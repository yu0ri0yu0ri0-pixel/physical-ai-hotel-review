"""
One-shot pipeline runner (no step-by-step confirmation).

  candidates → validate (PHYSICAL_AI / AUTOMATION_ONLY) → select 100 → match 100 conventional
  → crawl Trip.com reviews → download review photos (≤10/review) → raw dataset → crawl log → summary

Every stage is resume-safe; errors in one hotel never stop the run. Source Excel files are only read.

  python run_all.py                    # full live run
  python run_all.py --offline          # stages that need no network (selection + provisional matching)
  python run_all.py --limit 5 --headful   # smoke test on 5 hotels with a visible browser
"""
from __future__ import annotations

import argparse
import traceback

import common as U
import config as C

log = U.get_logger("run_all")


def stage(name, fn, *a, **k):
    log.info("========== %s ==========", name)
    try:
        return fn(*a, **k)
    except SystemExit:
        raise
    except Exception:  # noqa: BLE001
        log.error("Stage %s failed:\n%s", name, traceback.format_exc())
        return None


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--offline", action="store_true")
    ap.add_argument("--skip-discovery", action="store_true")
    ap.add_argument("--reselect", action="store_true")
    ap.add_argument("--limit", type=int, help="crawl only the first N hotels (smoke test)")
    ap.add_argument("--headful", action="store_true")
    ap.add_argument("--retry-partial", action="store_true")
    a = ap.parse_args()
    if a.headful:
        C.HEADLESS = False
    U.ensure_dirs()
    import validate_hotels, match_controls
    stage("1 validate + select unmanned/automated 100", validate_hotels.run, offline=a.offline, reselect=a.reselect)
    stage("2 match conventional controls", match_controls.run, offline=a.offline, skip_discovery=a.skip_discovery)
    if a.offline:
        log.info("Offline mode: crawling and image download skipped.")
    else:
        import crawl_tripcom_reviews, download_review_images
        stage("3 crawl Trip.com reviews", crawl_tripcom_reviews.run, limit=a.limit,
              retry_partial=a.retry_partial, headful=a.headful)
        stage("4 download review photos + build raw dataset", download_review_images.run)
    import make_summary
    stage("5 crawling log + sample summary", make_summary.run)


if __name__ == "__main__":
    main()
