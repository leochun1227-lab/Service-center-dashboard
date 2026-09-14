"""Commit and push the files used by the Render static site."""
from __future__ import annotations

import argparse
from datetime import datetime
import os
from pathlib import Path
import shutil
import subprocess
import sys


ROOT = Path(__file__).resolve().parents[1]
PUBLISH_FILES = (
    "overview.html",
    "dashboard-data.js",
    "assets/service-order-icons.js",
    "assets/lucide-LICENSE",
    "render.yaml",
)


def find_git() -> str:
    executable = shutil.which("git")
    if executable:
        return executable
    # GitHub Desktop and Codex can supply Git even when it is absent from PATH.
    candidates = [
        Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/native/git/cmd/git.exe",
        Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "Git/cmd/git.exe",
    ]
    desktop = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local"))) / "GitHubDesktop"
    candidates.extend(sorted(desktop.glob("app-*/resources/app/git/cmd/git.exe"), reverse=True))
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    raise RuntimeError("Git was not found. Install Git or GitHub Desktop, then retry publishing.")


def git(repo: Path, executable: str, *args: str) -> str:
    result = subprocess.run(
        [executable, *args], cwd=repo, text=True, encoding="utf-8", errors="replace",
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        env={**os.environ, "GIT_TERMINAL_PROMPT": "0", "GCM_INTERACTIVE": "Never"},
    )
    if result.returncode:
        raise RuntimeError(f"git {args[0]} failed:\n{result.stderr.strip() or result.stdout.strip()}")
    return result.stdout.strip()


def check_repository(repo: Path, executable: str) -> None:
    if git(repo, executable, "branch", "--show-current") != "main":
        raise RuntimeError("Publishing requires the main branch. Switch to main and retry.")
    git(repo, executable, "remote", "get-url", "--push", "origin")
    for setting in ("user.name", "user.email"):
        git(repo, executable, "config", "--get", setting)
    if git(repo, executable, "ls-files", "--unmerged"):
        raise RuntimeError("Resolve the current Git conflicts before publishing.")
    for filename in PUBLISH_FILES:
        if not (repo / filename).is_file():
            raise RuntimeError(f"Required website file is missing: {filename}")


def publish(repo: Path, executable: str) -> None:
    check_repository(repo, executable)
    print("[RUN] Checking GitHub main...", flush=True)
    git(repo, executable, "fetch", "origin", "main")
    behind = int(git(repo, executable, "rev-list", "--count", "HEAD..FETCH_HEAD"))
    if behind:
        raise RuntimeError(
            "GitHub main has newer commits. Sync this repository in GitHub Desktop, "
            "then run RUN_PUBLISH_DATA.bat. The updated local data has been kept."
        )
    changes = git(repo, executable, "status", "--porcelain", "--untracked-files=normal", "--", *PUBLISH_FILES)
    if changes:
        print("[RUN] Committing dashboard website files...", flush=True)
        git(repo, executable, "add", "--", *PUBLISH_FILES)
        # --only excludes unrelated staged files from this commit.
        message = f"Update dashboard data {datetime.now():%Y-%m-%d %H:%M:%S}"
        git(repo, executable, "commit", "--only", "-m", message, "--", *PUBLISH_FILES)
    else:
        print("[OK] Website files unchanged; checking for an unpushed update.", flush=True)
    print("[RUN] Pushing dashboard to GitHub...", flush=True)
    # Always push, including after an earlier run committed but failed to push.
    git(repo, executable, "push", "origin", "HEAD:refs/heads/main")
    print("[SUCCESS] Dashboard pushed to GitHub.", flush=True)
    print("[INFO] Render deploys this commit when Auto-Deploy is set to On Commit.", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Check local publishing prerequisites without committing or pushing")
    args = parser.parse_args()
    try:
        executable = find_git()
        if args.check:
            check_repository(ROOT, executable)
            print("[OK] Git publishing prerequisites are ready.")
        else:
            publish(ROOT, executable)
        return 0
    except (RuntimeError, OSError) as error:
        print(f"[ERROR] {error}", file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
