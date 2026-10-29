from __future__ import annotations

import unittest
from unittest import mock

import pdf2word
from tests.helpers import (
    FakeWorker,
    error_payload,
    make_job,
    success_payload,
)


class SlotRetryTests(unittest.TestCase):
    def make_job(self):
        return make_job(r"C:\Data\in.pdf", r"C:\Data\out.docx", retries=1)

    def test_retry_after_timeout(self):
        job = self.make_job()
        first = FakeWorker([pdf2word._WorkerTimeout("timed out")])
        second = FakeWorker([success_payload(job)])
        workers = [first, second]
        slot = pdf2word._Slot(0, lambda _index: workers.pop(0))

        with mock.patch("pdf2word.time.sleep") as sleep:
            result = slot.convert(job)

        self.assertEqual(result.output.name, "out.docx")
        self.assertTrue(first.killed)
        self.assertEqual(len(first.requests), 1)
        self.assertEqual(len(second.requests), 1)
        sleep.assert_any_call(10.0)

    def test_retry_after_worker_crash(self):
        job = self.make_job()
        first = FakeWorker([pdf2word._WorkerCrashed("worker died")])
        second = FakeWorker([success_payload(job)])
        workers = [first, second]
        slot = pdf2word._Slot(0, lambda _index: workers.pop(0))

        with mock.patch("pdf2word.time.sleep"):
            result = slot.convert(job)

        self.assertEqual(result.output.name, "out.docx")
        self.assertTrue(first.killed)
        self.assertEqual(len(second.requests), 1)

    def test_rpc_failure_restarts_worker(self):
        job = self.make_job()
        first = FakeWorker([
            error_payload("word", "The RPC server is unavailable (0x800706BE)")
        ])
        second = FakeWorker([success_payload(job)])
        workers = [first, second]
        slot = pdf2word._Slot(0, lambda _index: workers.pop(0))

        with mock.patch("pdf2word.time.sleep") as sleep:
            result = slot.convert(job)

        self.assertEqual(result.output.name, "out.docx")
        self.assertTrue(first.killed)
        self.assertEqual(len(second.requests), 1)
        sleep.assert_any_call(10.0)

    def test_validation_error_is_not_retried(self):
        job = make_job(r"C:\Data\in.pdf", r"C:\Data\out.docx", retries=3)
        worker = FakeWorker([error_payload("validation", "bad request")])
        slot = pdf2word._Slot(0, lambda _index: worker)

        with self.assertRaises(pdf2word.ValidationError):
            slot.convert(job)

        self.assertEqual(len(worker.requests), 1)

    def test_attempt_count_retries_plus_one(self):
        job = make_job(r"C:\Data\in.pdf", r"C:\Data\out.docx", retries=2)
        worker = FakeWorker([
            error_payload("word", "failure 1"),
            error_payload("word", "failure 2"),
            error_payload("word", "failure 3"),
        ])
        slot = pdf2word._Slot(0, lambda _index: worker)

        with mock.patch("pdf2word.time.sleep"):
            with self.assertRaises(pdf2word.WordError):
                slot.convert(job)

        self.assertEqual(len(worker.requests), 3)


if __name__ == "__main__":
    unittest.main()
