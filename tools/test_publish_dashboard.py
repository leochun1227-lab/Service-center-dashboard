"""Exercise publishing against a disposable local Git remote; no GitHub access."""
from pathlib import Path
import subprocess
import tempfile
import unittest

from publish_dashboard import PUBLISH_FILES, ROOT, check_repository, find_git, git, publish


class PublishTests(unittest.TestCase):
    def test_publish_retry_and_preserve_unrelated_changes(self):
        executable = find_git()
        with tempfile.TemporaryDirectory(prefix="publish-test-", dir=ROOT / "outputs") as directory:
            base = Path(directory)
            remote = base / "remote.git"
            repo = base / "checkout"
            repo.mkdir()
            def run(*args):
                return git(repo, executable, *args)
            subprocess.run([executable, "init", "--bare", str(remote)], check=True, capture_output=True)
            run("init", "-b", "main")
            run("config", "user.name", "Publish Test")
            run("config", "user.email", "publish-test@example.invalid")
            for filename in (*PUBLISH_FILES, "unrelated.txt"):
                target = repo / filename
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_text("initial\n", encoding="utf-8")
            run("add", ".")
            run("commit", "-m", "Initial fixture")
            run("remote", "add", "origin", str(remote))
            run("push", "-u", "origin", "main")
            original = run("rev-parse", "HEAD")
            check_repository(repo, executable)
            self.assertEqual(run("rev-parse", "HEAD"), original)

            (repo / "unrelated.txt").write_text("user staged work\n", encoding="utf-8")
            run("add", "unrelated.txt")
            (repo / "dashboard-data.js").write_text("updated\n", encoding="utf-8")
            publish(repo, executable)
            published = run("rev-parse", "HEAD")
            self.assertNotEqual(published, original)
            self.assertEqual(run("diff-tree", "--no-commit-id", "--name-only", "-r", "HEAD"), "dashboard-data.js")
            self.assertEqual(run("diff", "--cached", "--name-only"), "unrelated.txt")
            self.assertEqual(run("rev-parse", "origin/main"), published)

            publish(repo, executable)
            self.assertEqual(run("rev-parse", "HEAD"), published)

            (repo / "dashboard-data.js").write_text("retry update\n", encoding="utf-8")
            run("remote", "set-url", "--push", "origin", str(base / "missing.git"))
            with self.assertRaisesRegex(RuntimeError, "git push failed"):
                publish(repo, executable)
            retry_commit = run("rev-parse", "HEAD")
            self.assertNotEqual(retry_commit, published)
            self.assertEqual(run("rev-parse", "origin/main"), published)
            run("remote", "set-url", "--push", "origin", str(remote))
            publish(repo, executable)
            self.assertEqual(run("rev-parse", "HEAD"), retry_commit)
            self.assertEqual(run("rev-parse", "origin/main"), retry_commit)

            other = base / "other"
            subprocess.run([executable, "clone", "-b", "main", str(remote), str(other)], check=True, capture_output=True)
            git(other, executable, "config", "user.name", "Remote Test")
            git(other, executable, "config", "user.email", "remote-test@example.invalid")
            (other / "overview.html").write_text("remote update\n", encoding="utf-8")
            git(other, executable, "commit", "-am", "Remote change")
            git(other, executable, "push", "origin", "main")
            (repo / "dashboard-data.js").write_text("local update kept\n", encoding="utf-8")
            with self.assertRaisesRegex(RuntimeError, "newer commits"):
                publish(repo, executable)
            self.assertEqual(run("rev-parse", "HEAD"), retry_commit)
            self.assertEqual((repo / "dashboard-data.js").read_text(), "local update kept\n")
            self.assertEqual(run("diff", "--cached", "--name-only"), "unrelated.txt")
            run("switch", "-c", "test-branch")
            with self.assertRaisesRegex(RuntimeError, "main branch"):
                check_repository(repo, executable)


if __name__ == "__main__":
    unittest.main()
