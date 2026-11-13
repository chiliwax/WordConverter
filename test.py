import shutil
import sys
import os
import logging
import argparse
from pathlib import Path

# Ajout du dossier lib/ au path Python
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

try:
    # Projet appelant : la bibliotheque est vendue dans lib/wordconverter/.
    from wordconverter.pdf2word import PdfConverter, ProgressEvent, repair_html_outputs
except ModuleNotFoundError:
    # Ce depot : le module est a la racine.
    from pdf2word import PdfConverter, ProgressEvent, repair_html_outputs

try:
    from doc_meta.doc_meta import META_FILENAME
    from downloader.ps_downloader import long_path
except ModuleNotFoundError:
    # Script execute hors du projet appelant : pas de metadonnees a copier.
    META_FILENAME = None
    long_path = str


# ---------------------------------------------------------------------------
# Configuration du logging
# ---------------------------------------------------------------------------
# On branche le logger interne de la lib ("pdf2word") sur la console.
# Le logger racine de notre script est séparé pour garder les niveaux distincts.

logging.basicConfig(
    level=logging.WARNING,  # niveau par défaut pour les autres modules
    format="%(asctime)s [%(levelname)-8s] %(name)s - %(message)s",
    datefmt="%H:%M:%S",
    stream=sys.stdout,
)

# Logger de notre script
log = logging.getLogger("convert_pdfs")
log.setLevel(logging.DEBUG)

# Logger interne de pdf2word : on l'active en INFO pour voir ses traces internes
pdf2word_log = logging.getLogger("pdf2word")
pdf2word_log.setLevel(logging.INFO)

# ---------------------------------------------------------------------------
# METADATA HANDLER
# ---------------------------------------------------------------------------

def copy_metadata(source: Path, output: Path) -> bool:
    """Copie le sidecar de métadonnées du PDF source vers le dossier du fichier converti.

    Retourne True si la copie a réussi. Ne lève jamais d'exception : un échec
    de copie ne doit pas interrompre le reste du batch.
    """
    if META_FILENAME is None:
        return False

    src = Path(long_path(str(source.parent / META_FILENAME)))
    dst = Path(long_path(str(output.parent / META_FILENAME)))
    tmp = dst.with_suffix(".tmp")

    if not src.is_file():
        log.warning("Pas de %s à côté de %s : métadonnées non copiées.", META_FILENAME, source.name)
        return False
    try:
        shutil.copy2(src, tmp)
        os.replace(tmp, dst)
    except OSError as exc:
        log.warning("Copie des métadonnées impossible pour %s : %s", source.name, exc)
        return False
    return True

# ---------------------------------------------------------------------------
# Callback de progression
# ---------------------------------------------------------------------------

def on_event(event: ProgressEvent) -> None:
    """Reçoit chaque événement de progression émis par convert_many."""
    if event.phase == "started":
        log.info(
            "[%d/%d] Démarrage : %s",
            event.done + 1,  # +1 car done n'est pas encore incrémenté au départ
            event.total,
            event.source.name if event.source else "?",
        )

    elif event.phase == "completed":
        result = event.result
        log.info(
            "[%d/%d] ✓ Converti  : %s → %s (%.1f KB, %.1fs)",
            event.done,
            event.total,
            event.source.name if event.source else "?",
            result.output.name if result else "?",
            result.size_kb if result else 0.0,
            result.seconds if result else 0.0,
        )
        if event.source and result:
            copy_metadata(event.source, result.output)
        # Affichage de la progression en pourcentage
        eta_str = f", ETA {event.eta:.0f}s" if event.eta is not None else ""
        log.debug("   Progression : %.1f%%%s", event.percent, eta_str)

    elif event.phase == "failed":
        failure = event.failure
        log.error(
            "[%d/%d] ✗ Échec     : %s — %s",
            event.done,
            event.total,
            event.source.name if event.source else "?",
            failure.error if failure else "erreur inconnue",
        )

    elif event.phase == "finished":
        log.info(
            "Batch terminé : %d/%d traités en %.1fs",
            event.done,
            event.total,
            event.elapsed,
        )

# ---------------------------------------------------------------------------
# Parsing des arguments CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convertit des fichiers PDF en Word via PdfConverter.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,  # affiche les valeurs par défaut dans --help
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("input"),
        metavar="DIR",
        help="Dossier source contenant les fichiers PDF.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output"),
        metavar="DIR",
        help="Dossier de destination pour les fichiers convertis.",
    )
    parser.add_argument(
        "--log-file",
        type=Path,
        default=None,
        metavar="FILE",
        help="Fichier de log (optionnel, en plus de la console).",
    )
    parser.add_argument(
        "--repair",
        action="store_true",
        help="Ne convertit rien : répare les sorties HTML existantes dans --output "
             "(dossier d'images + liens), sans lancer Word.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Avec --repair : affiche ce qui serait fait sans rien modifier.",
    )
    parser.add_argument(
        "--images-folder",
        default="source_folder",
        metavar="NAME",
        help="Nom du dossier d'images attendu par la conversion.",
    )
    return parser.parse_args()        


def setup_logging(log_file: Path | None = None) -> None:
    """Configure la console (format [LEVEL], lu par l'UI) et, si demandé, un fichier."""
    # Console Windows en cp1252 : un message contenant un caractère hors page de
    # codes doit être remplacé, pas lever une exception dans le logger.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass

    handlers: list[logging.Handler] = []

    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    handlers.append(console)

    if log_file:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
        ))
        handlers.append(fh)

    # force=True : remplace d'éventuels handlers déjà installés (import multiple, tests)
    logging.basicConfig(level=logging.WARNING, handlers=handlers, force=True)

# ---------------------------------------------------------------------------
# Script principal
# ---------------------------------------------------------------------------



def run_repair(output_dir: Path, images_folder: str, dry_run: bool) -> int:
    """Répare les sorties HTML existantes, sans reconvertir et sans Word."""
    if not output_dir.is_dir():
        log.error("Le dossier de sortie '%s' n'existe pas ou n'est pas un dossier.", output_dir)
        return 1

    def on_result(result) -> None:
        if result.action == "repaired":
            details = []
            if result.renamed_from:
                details.append(f"dossier {result.renamed_from} -> {images_folder}")
            if result.links_replaced:
                details.append(f"{result.links_replaced} lien(s)")
            log.info("Réparé   : %s (%s)", result.path, ", ".join(details))
        elif result.action == "unchanged":
            log.debug("Déjà bon : %s", result.path)
        else:
            log.warning("Ignoré   : %s — %s", result.path, result.detail)

    summary = repair_html_outputs(
        output_dir,
        images_folder=images_folder,
        recursive=True,
        dry_run=dry_run,
        on_result=on_result,
    )

    suffix = " (dry-run, rien n'a été écrit)" if dry_run else ""
    log.info(
        "Réparation terminée : %d réparé(s), %d déjà bon(s), %d ignoré(s)%s.",
        len(summary.repaired),
        len(summary.unchanged),
        len(summary.skipped),
        suffix,
    )
    return 0


def main() -> None:
    args = parse_args()
    setup_logging(args.log_file)

    input_dir: Path = args.input
    output_dir: Path = args.output

    if args.dry_run and not args.repair:
        log.error("--dry-run s'utilise avec --repair.")
        sys.exit(2)

    if args.repair:
        sys.exit(run_repair(output_dir, args.images_folder, args.dry_run))

    if not input_dir.is_dir():
        log.error("Le dossier source '%s' n'existe pas ou n'est pas un dossier.", input_dir)
        sys.exit(1)

    pdfs = sorted(input_dir.rglob("*.pdf"))

    if not pdfs:
        log.warning("Aucun fichier PDF trouvé dans '%s'.", input_dir)
        return

    log.info("Démarrage de la conversion de %d fichier(s) PDF.", len(pdfs))

    with PdfConverter(format="html", workers=4) as conv:
        batch = conv.convert_many(
            str(input_dir),
            out_dir=str(output_dir),
            force=True,
            on_event=on_event,   # ← callback de progression
            recursive=True,
            timeout=600, # ← for bigger files (some don't have the time to be translate from pdf to word -> maybe we should adjust timeout according to file size)
            retries=3
        )

    # Résumé final
    log.info("Résultat : %d converti(s), %d échec(s).", len(batch.ok), len(batch.failed))

    if batch.failed:
        log.error("Fichiers en échec :")
        for failure in batch.failed:
            log.error("  - %s : %s", failure.source, failure.error)


if __name__ == "__main__":
    main()
