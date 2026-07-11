"""
Phase 6 — Vercel Deploy
Pushes the freshly-injected dashboard HTML to Vercel as a static site,
giving it a public www link. Runs on the lab node after each successful
pipeline run — this file never runs anywhere else and needs the Vercel CLI
installed + authorized on whichever machine calls it (see README.md's
"Public hosting (Vercel)" section for the one-time setup).
"""

import logging
import shutil
import subprocess
import sys
from pathlib import Path

BASE_DIR     = Path(__file__).parent
DEPLOY_DIR   = BASE_DIR / "vercel_deploy"
DASHBOARD_FILE = BASE_DIR / "GL_Dashboard_v4_July2026.html"

# Stable production alias — Vercel re-points this to the latest prod deploy
# automatically, so it never changes even though each individual deploy also
# gets its own unique throwaway URL (that per-deploy URL is only useful for
# Vercel's own dashboard/audit trail, not for anyone to actually bookmark).
PRODUCTION_URL = "https://gl-campus-intelligence.vercel.app"

log = logging.getLogger(__name__)


def deploy(dashboard_path: Path = None, prod: bool = True) -> tuple:
    """
    Copy the current dashboard HTML into vercel_deploy/index.html and run
    `vercel deploy` from that directory. Returns (success, deployed_url) —
    deployed_url is the per-deploy URL Vercel just printed (or None on
    failure); callers that want the stable public link should use
    PRODUCTION_URL instead, since that's what recipients actually visit.
    Never raises — this must not take down the pipeline run that triggered
    it; a failed deploy just means the public link is stale until the next
    successful run, not that the whole pipeline failed.
    """
    src = dashboard_path or DASHBOARD_FILE
    if not src.exists():
        log.error("Vercel deploy skipped: dashboard file not found at %s", src)
        return False, None

    vercel_bin = shutil.which("vercel")
    if not vercel_bin:
        log.error(
            "Vercel deploy skipped: 'vercel' CLI not found on PATH. "
            "Install it with `npm install -g vercel` and run `vercel login` "
            "once on this machine (see README.md)."
        )
        return False, None

    DEPLOY_DIR.mkdir(exist_ok=True)
    dest = DEPLOY_DIR / "index.html"
    shutil.copy2(src, dest)
    log.info("Copied %s -> %s", src.name, dest)

    cmd = [vercel_bin, "deploy", "--yes"]
    if prod:
        cmd.append("--prod")

    try:
        result = subprocess.run(
            cmd, cwd=str(DEPLOY_DIR), capture_output=True, text=True, timeout=180,
        )
    except subprocess.TimeoutExpired:
        log.error("Vercel deploy timed out after 180s")
        return False, None

    if result.returncode != 0:
        log.error("Vercel deploy failed (exit %d):\n%s\n%s",
                   result.returncode, result.stdout, result.stderr)
        return False, None

    deployed_url = result.stdout.strip().splitlines()[-1] if result.stdout.strip() else None
    log.info("Vercel deploy succeeded: %s", deployed_url or "(no URL in output)")
    return True, deployed_url


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-8s %(message)s")
    ok, _url = deploy(prod="--dry-run" not in sys.argv)
    sys.exit(0 if ok else 1)
