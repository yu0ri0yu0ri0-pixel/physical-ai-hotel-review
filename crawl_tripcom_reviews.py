"""
Stage 3 — Crawl the newest (max 100) Trip.com reviews for every hotel in work/pairs_current.csv.

How it works (no API replay, no bypass)
  * Opens the hotel's public review page in Chromium (Playwright) and *listens* to the JSON
    responses the page itself requests (review endpoint observed in 2026: getHotelCommentInfo).
  * Tries to switch the list to "most recent" through the visible sort control, then follows the
    visible next-page / load-more control, one click at a time with random delays.
  * If the sort control cannot be found, it pages through more reviews (capped) and keeps the
    100 most recent by review date. The sort mode used is logged per hotel.
  * Review text is stored exactly as delivered (originalContent preferred over display text).
  * CAPTCHA, login walls and HTTP 403/429/430 are detected → hotel recorded as FAILED/PARTIAL,
    crawler moves on. After several consecutive blocks it pauses; after repeated blocks it stops
    cleanly so you can resume later.

Resume
  * Every new review is appended to work/reviews/{hotel_id}.jsonl immediately.
  * work/crawl_state.json is written after each hotel. Re-running skips SUCCESS hotels and never
    re-adds an existing review_id.

Usage
  python crawl_tripcom_reviews.py                 # all hotels, resume-aware
  python crawl_tripcom_reviews.py --probe H001    # diagnose one hotel (prints what was captured)
  python crawl_tripcom_reviews.py --only H001,H101 --headful
  python crawl_tripcom_reviews.py --retry-partial # also retry PARTIAL hotels
"""
from __future__ import annotations

import argparse
import json
import random
import sys
import time

import pandas as pd

import common as U
import config as C

log = U.get_logger("crawl")
TRANSIENT_REASONS = {"TIMEOUT", "ERROR", "NO_REVIEW_PAYLOAD", "HTTP_429"}


# ------------------------------------------------------------------ state

def load_state_raw() -> dict:
    """Crawl state keyed by Trip.com hotel code (stable across re-selection / re-matching)."""
    return U.read_json(C.CRAWL_STATE, {}) or {}


def load_state() -> dict:
    """hotel_id-keyed view of the crawl state for the current pair plan."""
    raw = load_state_raw()
    if not C.PAIRS_CURRENT_CSV.exists():
        return {}
    pairs = pd.read_csv(C.PAIRS_CURRENT_CSV, dtype={"hotel_id": str})
    out = {}
    for r in pairs.itertuples():
        code = U.trip_code_from_url(r.trip_review_url)
        if code in raw:
            out[r.hotel_id] = raw[code]
    return out


def save_state(state: dict) -> None:
    U.write_json_atomic(C.CRAWL_STATE, state)


def review_key(row) -> str:
    """Reviews are stored per Trip.com hotel code, so they stay attached to the right hotel even
    if hotel_id numbering changes between an offline and a live selection."""
    return U.trip_code_from_url(row.trip_review_url) or str(row.hotel_id)


def load_reviews(key: str) -> dict[str, dict]:
    path = C.REVIEW_DIR / f"{key}.jsonl"
    out: dict[str, dict] = {}
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            try:
                r = json.loads(line)
            except Exception:
                continue
            # a later line may upgrade a review first seen in default order to newest-order
            prev = out.get(r["review_id"])
            if prev is None or (r.get("newest_rank") is not None and prev.get("newest_rank") is None):
                out[r["review_id"]] = r
    return out


def append_reviews(key: str, reviews: list[dict]) -> None:
    if not reviews:
        return
    C.REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    with (C.REVIEW_DIR / f"{key}.jsonl").open("a", encoding="utf-8") as f:
        for r in reviews:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

# ------------------------------------------------------------------ page interaction

def _click_first(cap: U.PageCapture, selectors: list[str]) -> bool:
    page = cap.page
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count() == 0 or not loc.is_visible():
                continue
            loc.scroll_into_view_if_needed(timeout=5_000)
            loc.click(timeout=8_000)
            return True
        except Exception:
            continue
    return False


def _wait_new_payload(cap: U.PageCapture, timeout_s: float = 12.0) -> list:
    t0 = time.time()
    got = []
    while time.time() - t0 < timeout_s:
        cap.page.wait_for_timeout(500)
        got = cap.drain()
        if got:
            cap.page.wait_for_timeout(800)          # let sibling responses land
            got += cap.drain()
            break
    return got


def try_sort_newest(cap: U.PageCapture) -> tuple[bool, list]:
    if _click_first(cap, C.SORT_NEWEST_SELECTORS):
        return True, _wait_new_payload(cap)
    if _click_first(cap, C.SORT_OPEN_SELECTORS):
        U.polite_sleep((0.8, 1.5))
        if _click_first(cap, C.SORT_NEWEST_SELECTORS):
            return True, _wait_new_payload(cap)
    return False, []


_MARK_LOAD_MORE_JS = """(texts) => {
  document.querySelectorAll('[data-crawl-loadmore]').forEach(e => e.removeAttribute('data-crawl-loadmore'));
  const leaves = [...document.querySelectorAll('span,button,a,div')]
      .filter(e => e.children.length === 0 && texts.includes((e.textContent || '').trim()));
  for (const e of leaves.reverse()) {
    for (const el of [e, e.parentElement]) {
      if (!el) continue;
      const bg = getComputedStyle(el).backgroundColor;
      const r = el.getBoundingClientRect();
      const filled = bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent' && bg !== 'rgb(255, 255, 255)';
      if (filled && r.width >= 120 && r.height >= 30) { el.setAttribute('data-crawl-loadmore', '1'); return true; }
    }
  }
  return false;
}"""


def _click_load_more_button(cap: U.PageCapture) -> bool:
    """Click the review list's filled "Show More" button (never the per-review text expanders)."""
    page = cap.page
    try:
        if not page.evaluate(_MARK_LOAD_MORE_JS, C.LOAD_MORE_BUTTON_TEXTS):
            return False
        loc = page.locator("[data-crawl-loadmore='1']").first
        loc.scroll_into_view_if_needed(timeout=5_000)
        page.wait_for_timeout(400)
        loc.click(timeout=8_000)
        return True
    except Exception:
        return False


REVIEW_OPEN_SELECTORS = [
    "text=/^(See all reviews|All reviews|Read all reviews|查看全部点评|すべての口コミ|모든 리뷰)/i",
    "text=/^Reviews?\\s*\\(/i",
]

# ------------------------------------------------------------------ one hotel

def crawl_hotel(br: U.TripBrowser, row: pd.Series, entry: dict, probe: bool = False,
                snap: bool = False) -> dict:
    hid = row.hotel_id
    key = review_key(row)
    if entry.get("trip_code") and entry["trip_code"] != key:      # hotel_id now points to another hotel
        entry = {k: entry[k] for k in ("pair_id", "group", "hotel_name", "trip_review_url") if k in entry}
    entry["trip_code"] = key
    have = load_reviews(key)
    total_known = U.to_number(row.get("total_review_count"))
    entry.update(retry_count=entry.get("retry_count", -1) + 1, crawl_timestamp=U.now_iso())
    pages, stalls, sort_mode, reason = 0, 0, "DEFAULT_ORDER", ""
    newest_target = C.MAX_REVIEWS_PER_HOTEL
    cap = None

    newest_seen: list[str] = []          # review ids in Trip.com's "Most Recent" order (this session)

    def ingest(payloads, newest: bool = False) -> tuple[int, int]:
        """→ (new reviews stored, reviews seen in these payloads)"""
        nonlocal total_known
        U.save_raw(key, payloads)
        new, seen = [], 0
        for _url, data in payloads:
            t = U.parse_total_review_count(data)
            if t is not None and (total_known is None or t > 0):
                total_known = t
            for r in U.parse_reviews(data):
                seen += 1
                rid = r["review_id"]
                if newest and rid not in newest_seen:
                    newest_seen.append(rid)
                prev = have.get(rid)
                if prev is None or (newest and prev.get("newest_rank") is None):
                    r["crawled_at"] = prev.get("crawled_at") if prev else U.now_iso()
                    r["page_no"] = pages
                    r["newest_rank"] = newest_seen.index(rid) + 1 if newest else None
                    have[rid] = r
                    new.append(r)
        append_reviews(key, new)
        return len(new), seen

    try:
        cap = br.open(row.trip_review_url, C.REVIEW_RESPONSE_PATTERNS)
        blocked = cap.check_blocked(check_login_text=False)
        if blocked:
            reason = blocked
            raise U.Blocked(blocked)
        if not probe:
            try:                                      # hotel-level data rides on the same page
                U.meta_from_capture(cap, key, row.trip_hotel_url)
            except Exception:  # noqa: BLE001
                pass
        ingest(cap.drain())
        ingest([("embedded", b) for b in cap.embedded_json()])
        if not have:                                  # reviews may sit behind a tab/button
            if _click_first(cap, REVIEW_OPEN_SELECTORS):
                ingest(_wait_new_payload(cap))
        sorted_ok, payload = try_sort_newest(cap)
        if sorted_ok and payload:
            sort_mode = "NEWEST_UI"
            ingest(payload, newest=True)
        else:
            sorted_ok = False
        # how many to gather: with newest-first order 100 is enough; otherwise gather more and trim
        cap_n = newest_target if sorted_ok else min(int(total_known or 300), 300)
        before_session = len(have)

        def progress() -> int:
            return len(newest_seen) if sorted_ok else len(have)

        while progress() < cap_n and pages < C.MAX_REVIEW_PAGES:
            if total_known is not None and progress() >= total_known:
                break
            U.polite_sleep(C.ACTION_DELAY)
            if not (_click_load_more_button(cap) or _click_first(cap, C.NEXT_PAGE_SELECTORS)
                    or _click_first(cap, C.LOAD_MORE_SELECTORS)):
                # the public page stops offering "Show More" after a fixed number of pages
                reason = reason or ("VISIBLE_LIMIT_REACHED" if pages >= 1 else "NO_NEXT_PAGE_CONTROL")
                break
            pages += 1
            before = progress()
            _n_new, n_seen = ingest(_wait_new_payload(cap), newest=sorted_ok)
            if n_seen == 0 or progress() == before:
                stalls += 1
                b = cap.check_blocked(check_login_text=True)
                if b:
                    reason = b
                    break
                if stalls >= 2:
                    reason = reason or "PAGINATION_EXHAUSTED"
                    break
            else:
                stalls = 0
        if not have:
            b = cap.check_blocked(check_login_text=True)
            reason = b or "NO_REVIEW_PAYLOAD"
        if probe:
            print("\n=== PROBE", hid, row.hotel_name)
            print("matched JSON responses:", len(cap.parsed))
            for u, _ in cap.parsed[:15]:
                print("  ", u[:160])
            print("all XHR seen:", len(cap.all_xhr), "| block statuses:", cap.block_statuses)
            print("reviews parsed:", len(have), "| newest-order:", len(newest_seen), "| total_known:", total_known, "| sort:", sort_mode,
                  "| pages:", pages, "| reason:", reason)
            if have:
                sample = next(iter(have.values()))
                print(json.dumps({k: v for k, v in sample.items() if k != "image_urls_all"}, ensure_ascii=False,
                                 indent=1)[:1500])
                print("images on sample:", sample.get("image_urls_all", [])[:3])
            cap.snapshot(f"probe_{hid}")
        elif snap or len(have) == before_session == 0 or reason not in ("", "NO_NEXT_PAGE_CONTROL", "PAGINATION_EXHAUSTED", "VISIBLE_LIMIT_REACHED"):
            cap.snapshot(f"crawl_{hid}_{reason or 'empty'}")
    except U.Blocked:
        if cap:
            cap.snapshot(f"blocked_{hid}")
    except Exception as e:  # noqa: BLE001
        reason = "TIMEOUT" if "Timeout" in type(e).__name__ else "ERROR"
        entry["error_detail"] = f"{type(e).__name__}: {e}"[:300]
        if cap:
            cap.snapshot(f"error_{hid}")
    finally:
        if cap:
            try:
                cap.page.close()
            except Exception:
                pass

    target = min(C.MAX_REVIEWS_PER_HOTEL, int(total_known)) if total_known is not None else C.MAX_REVIEWS_PER_HOTEL
    n_newest = sum(1 for r in have.values() if r.get("newest_rank") is not None)
    in_sample = min(n_newest if sort_mode == "NEWEST_UI" else len(have), C.MAX_REVIEWS_PER_HOTEL)
    entry["collected_newest_order"] = n_newest
    if total_known == 0:
        status, reason = "SUCCESS", "NO_REVIEWS_ON_TRIPCOM"
    elif in_sample >= target and reason not in ("CAPTCHA", "LOGIN_REQUIRED", "HTTP_403", "HTTP_429", "HTTP_430"):
        status, reason = "SUCCESS", ""
    elif in_sample > 0:
        status = "PARTIAL"
        reason = reason or "FEWER_THAN_TARGET"
    else:
        status = "FAILED"
    entry.update(status=status, failure_reason=reason, target_review_count=target,
                 total_review_count_tripcom=total_known, collected_all=len(have), collected_in_sample=in_sample,
                 sort_mode=sort_mode, pages_clicked=pages)
    return entry

# ------------------------------------------------------------------ main loop

def should_crawl(entry: dict, retry_partial: bool) -> bool:
    st = entry.get("status")
    if st is None:
        return True
    if st == "SUCCESS":
        return False
    if entry.get("retry_count", 0) >= C.MAX_RETRIES:
        return False
    if st == "FAILED":
        return True
    return retry_partial or entry.get("failure_reason") in TRANSIENT_REASONS


def run(only: list[str] | None = None, limit: int | None = None, retry_partial: bool = False,
        headful: bool = False, probe: str | None = None, group: str | None = None) -> dict:
    U.ensure_dirs()
    pairs = pd.read_csv(C.PAIRS_CURRENT_CSV, dtype={"hotel_id": str})
    if probe:                       # diagnosis only: nothing is written to the real dataset/state
        only = [probe]
        C.REVIEW_DIR = C.WORK_DIR / "probe" / "reviews"
        C.RAW_RESPONSE_DIR = C.WORK_DIR / "probe" / "raw"
        C.REVIEW_DIR.mkdir(parents=True, exist_ok=True)
        C.RAW_RESPONSE_DIR.mkdir(parents=True, exist_ok=True)
    if only:
        pairs = pairs[pairs.hotel_id.isin(only)]
    if group:
        pairs = pairs[pairs.group == group]
    state = load_state_raw()
    todo = [r for _, r in pairs.iterrows()
            if probe or should_crawl(state.get(review_key(r), {}), retry_partial)]
    if limit:
        todo = todo[:limit]
    log.info("Hotels in plan: %d | to crawl now: %d", len(pairs), len(todo))
    consecutive_blocks, cooldowns = 0, 0
    with U.TripBrowser(headless=not (headful or probe)) as br:
        for n, row in enumerate(todo, 1):
            entry = dict(state.get(review_key(row), {})) if not probe else {}
            entry.update(hotel_id=row.hotel_id, pair_id=row.pair_id, group=row.group, hotel_name=row.hotel_name,
                         trip_review_url=row.trip_review_url)
            log.info("[%d/%d] %s %s", n, len(todo), row.hotel_id, row.hotel_name)
            entry = crawl_hotel(br, row, entry, probe=bool(probe), snap=n <= C.DEBUG_FIRST_N_HOTELS)
            if probe:
                continue
            state = load_state_raw()                  # merge-safe if another stage wrote meanwhile
            state[review_key(row)] = entry
            save_state(state)
            log.info("   → %s | %d reviews (target %s) | %s", entry["status"], entry["collected_in_sample"],
                     entry["target_review_count"], entry.get("failure_reason", ""))
            blocked = entry.get("failure_reason", "") in ("CAPTCHA", "LOGIN_REQUIRED", "HTTP_403", "HTTP_429", "HTTP_430")
            consecutive_blocks = consecutive_blocks + 1 if blocked else 0
            if consecutive_blocks >= C.CONSECUTIVE_BLOCK_LIMIT:
                cooldowns += 1
                if cooldowns > C.MAX_BLOCK_COOLDOWNS:
                    log.error("Repeated blocking. Stopping cleanly — re-run later to resume (no bypass attempted).")
                    break
                wait = random.uniform(*C.BLOCK_COOLDOWN)
                log.warning("%d consecutive blocked hotels → pausing %.0f s", consecutive_blocks, wait)
                time.sleep(wait)
                consecutive_blocks = 0
            if n < len(todo):
                U.polite_sleep(C.HOTEL_DELAY)
                if n % C.BATCH_SIZE == 0:
                    log.info("batch of %d done → cool-down", C.BATCH_SIZE)
                    U.polite_sleep(C.BATCH_COOLDOWN)
    return state


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--only", help="comma-separated hotel_ids")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--retry-partial", action="store_true")
    ap.add_argument("--headful", action="store_true", help="show the browser window")
    ap.add_argument("--probe", help="diagnose a single hotel_id and print captured data")
    ap.add_argument("--group", help="crawl only this group (UNMANNED_AUTOMATED or CONVENTIONAL)")
    a = ap.parse_args()
    if not C.PAIRS_CURRENT_CSV.exists():
        sys.exit("work/pairs_current.csv not found — run match_controls.py first")
    run(only=a.only.split(",") if a.only else None, limit=a.limit, retry_partial=a.retry_partial,
        headful=a.headful, probe=a.probe, group=a.group)
