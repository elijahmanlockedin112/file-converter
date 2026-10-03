"""Tests for finding files, planning outputs and running batches.

These use a fake backend, so they run without Microsoft Word:
    python -m unittest
"""

import tempfile
import threading
import unittest
from pathlib import Path

from file_converter.backends import ConversionError
from file_converter.batch import Job, Status, find_sources, plan_jobs, run_jobs


class FakeBackend:
    """Writes a placeholder PDF, or fails for files named 'bad*'."""

    opened = 0

    def __enter__(self):
        FakeBackend.opened += 1
        return self

    def __exit__(self, *exc_info):
        pass

    def convert(self, src, dst):
        if src.name.startswith("bad"):
            raise ConversionError("bad file")
        dst.write_text(f"pdf of {src.name}")


class TempDirTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name).resolve()

    def tearDown(self):
        self._tmp.cleanup()

    def touch(self, *names):
        for name in names:
            path = self.root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("x")


class FindSourcesTest(TempDirTest):
    def test_finds_only_word_documents(self):
        self.touch("a.docx", "b.DOC", "~$a.docx", "notes.txt", "c.pdf")
        names = [s.path.name for s in find_sources([self.root])]
        self.assertEqual(names, ["a.docx", "b.DOC"])

    def test_subfolders_only_when_recursive(self):
        self.touch("top.docx", "sub/inner.docx")
        self.assertEqual(len(find_sources([self.root])), 1)
        self.assertEqual(len(find_sources([self.root], recursive=True)), 2)

    def test_same_file_listed_once(self):
        self.touch("a.docx")
        sources = find_sources([self.root / "a.docx", self.root, self.root / "A.DOCX"])
        self.assertEqual(len(sources), 1)


class PlanJobsTest(TempDirTest):
    def test_beside_originals_by_default(self):
        self.touch("sub/a.docx")
        (job,) = plan_jobs(find_sources([self.root], recursive=True))
        self.assertEqual(job.dst, self.root / "sub" / "a.pdf")

    def test_output_folder_mirrors_scanned_subfolders(self):
        self.touch("in/top.docx", "in/sub/inner.doc")
        out = self.root / "out"
        jobs = plan_jobs(find_sources([self.root / "in"], recursive=True), out)
        self.assertEqual(
            sorted(j.dst for j in jobs), [out / "sub" / "inner.pdf", out / "top.pdf"]
        )

    def test_name_clashes_get_numbered(self):
        self.touch("report.doc", "report.docx", "other/report.docx")
        sources = find_sources([self.root / "report.doc", self.root / "report.docx",
                                self.root / "other" / "report.docx"])
        out = self.root / "out"
        names = [j.dst.name for j in plan_jobs(sources, out)]
        self.assertEqual(names, ["report.pdf", "report (2).pdf", "report (3).pdf"])


class RunJobsTest(TempDirTest):
    def jobs(self, *names):
        self.touch(*names)
        return [Job(self.root / n, self.root / f"{Path(n).stem}.pdf", FakeBackend) for n in names]

    def test_failure_does_not_stop_the_batch(self):
        jobs = self.jobs("bad.docx", "good.docx")
        updates = []
        summary = run_jobs(jobs, on_update=lambda i, job, status, msg: updates.append(status))
        self.assertEqual((summary.done, summary.failed), (1, 1))
        self.assertEqual(
            updates, [Status.CONVERTING, Status.FAILED, Status.CONVERTING, Status.DONE]
        )

    def test_existing_outputs_skipped_unless_overwrite(self):
        jobs = self.jobs("a.docx")
        jobs[0].dst.write_text("old")
        self.assertEqual(run_jobs(jobs).skipped, 1)
        self.assertEqual(jobs[0].dst.read_text(), "old")
        self.assertEqual(run_jobs(jobs, overwrite=True).done, 1)
        self.assertEqual(jobs[0].dst.read_text(), "pdf of a.docx")

    def test_backend_not_started_when_everything_is_skipped(self):
        jobs = self.jobs("a.docx")
        jobs[0].dst.write_text("old")
        FakeBackend.opened = 0
        run_jobs(jobs)
        self.assertEqual(FakeBackend.opened, 0)

    def test_cancel_stops_before_next_file(self):
        jobs = self.jobs("a.docx", "b.docx")
        cancel = threading.Event()

        def on_update(i, job, status, msg):
            if status is Status.DONE:
                cancel.set()

        summary = run_jobs(jobs, on_update=on_update, cancel=cancel)
        self.assertEqual(summary.done, 1)
        self.assertTrue(summary.cancelled)
        self.assertFalse(jobs[1].dst.exists())


if __name__ == "__main__":
    unittest.main()
