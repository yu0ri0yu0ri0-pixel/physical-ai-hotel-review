"""
Stage 4 — Download user photos attached to each sampled review and build the raw dataset.

* Only images found INSIDE a review object are used (hotel gallery / room-sales photos are never
  touched, because they do not appear in review objects).
* Max 10 photos per review, in the order Trip.com lists them.
* Tries the original-resolution URL first (CDN resize suffix stripped), then the URL exactly as
  served. Files are saved unmodified (no resize/crop/re-encode); extension follows content type.
* Same URL → downloaded once; later occurrences point to the existing file.
* A failed photo never removes the review; failures are logged in work/image_manifest.csv.
* Resume: files already on disk / OK rows in the manifest are skipped.

Path layout: dataset/images/{pair_id}/{group}/{hotel_id}/{review_id}_{NN}.{ext}

Output: outputs/04_tripcom_reviews_raw.csv  (one row per review, newest ≤100 per hotel)
"""
from __future__ import annotations

import argparse
import csv
import random
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import requests

import common as U
import config as C
import crawl_tripcom_reviews as CR

log = U.get_logger("images")
_lock = threading.Lock()
MANIFEST_COLS = ["pair_id", "hotel_id", "review_id", "photo_index", "photo_url", "download_url_used", "variant",
                 "local_path", "bytes", "status", "error", "ts"]
EXT_BY_TYPE = {"image/jpeg": ".jpg", "image/jpg": ".jpg", "image/png": ".png", "image/webp": ".webp",
               "image/gif": ".gif", "image/heic": ".heic", "image/avif": ".avif"}


def sample_reviews(row) -> list[dict]:
    """Newest ≤100 reviews of a hotel (by review_date; crawl order as tie-breaker)."""
    revs = list(CR.load_reviews(CR.review_key(row)).values())
    if not revs:
        return []
    df = pd.DataFrame(revs)
    df["_order"] = range(len(df))
    df["_dt"] = pd.to_datetime(df["review_date"], errors="coerce", utc=True, format="mixed")
    # reviews delivered under Trip.com's "Most Recent" sort come first (newest by construction);
    # reviews only seen in the default "Most relevant" order fill up afterwards
    if "newest_rank" not in df.columns:
        df["newest_rank"] = None
    if df["newest_rank"].notna().any():
        # the study samples the *latest* reviews: when the newest sort worked, reviews that were only
        # delivered in the default "Most relevant" order are not part of the sample
        df = df[df["newest_rank"].notna()]
    df = df.sort_values(["_dt", "_order"], ascending=[False, True], na_position="last")
    return df.head(C.MAX_REVIEWS_PER_HOTEL).drop(columns=["_order", "_dt"]).to_dict("records")


def _copy_duplicate(src_rel: str, dest_stem: Path) -> str:
    """Same URL already downloaded: copy the local file (no network) so the filename still
    identifies its own review ({review_id}_{NN})."""
    import shutil
    src = C.ROOT / src_rel
    dest = Path(str(dest_stem) + src.suffix)
    if not dest.exists():
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
    return dest.relative_to(C.ROOT).as_posix()


def load_manifest() -> pd.DataFrame:
    if C.IMAGE_MANIFEST.exists():
        return pd.read_csv(C.IMAGE_MANIFEST, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    return pd.DataFrame(columns=MANIFEST_COLS)


def append_manifest(rows: list[dict]) -> None:
    new = not C.IMAGE_MANIFEST.exists()
    with _lock, C.IMAGE_MANIFEST.open("a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=MANIFEST_COLS)
        if new:
            w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in MANIFEST_COLS})


def _fetch(session: requests.Session, url: str) -> tuple[bytes, str]:
    resp = session.get(url, timeout=C.IMAGE_TIMEOUT, headers={"Referer": C.TRIP_BASE + "/"})
    resp.raise_for_status()
    ctype = resp.headers.get("Content-Type", "").split(";")[0].strip().lower()
    body = resp.content
    if not ctype.startswith("image/") or len(body) < 512:
        raise ValueError(f"not an image (type={ctype!r}, bytes={len(body)})")
    return body, ctype


def download_one(session: requests.Session, job: dict) -> dict:
    url = job["photo_url"]
    out = {**job, "ts": U.now_iso()}
    attempts = []
    up = U.upgrade_image_url(url)
    if up != url:
        attempts.append((up, "original_resolution"))
    attempts.append((url, "as_served"))
    last_err = ""
    for u, variant in attempts:
        for _try in range(2):
            try:
                body, ctype = _fetch(session, u)
                ext = EXT_BY_TYPE.get(ctype) or (Path(u.split("?")[0]).suffix or ".jpg")
                dest = Path(job["dest_stem"] + ext)
                dest.parent.mkdir(parents=True, exist_ok=True)
                tmp = dest.with_suffix(dest.suffix + ".part")
                tmp.write_bytes(body)
                tmp.replace(dest)
                out.update(download_url_used=u, variant=variant, local_path=dest.relative_to(C.ROOT).as_posix(),
                           bytes=len(body), status="OK", error="")
                time.sleep(random.uniform(*C.IMAGE_DELAY))
                return out
            except Exception as e:  # noqa: BLE001
                last_err = f"{type(e).__name__}: {e}"[:200]
                time.sleep(1.5)
    out.update(status="FAILED", error=last_err, local_path="", bytes=0, variant="", download_url_used="")
    return out


def run(skip_download: bool = False) -> pd.DataFrame:
    U.ensure_dirs()
    pairs = pd.read_csv(C.PAIRS_CURRENT_CSV, dtype={"hotel_id": str})
    man = load_manifest()
    ok = man[man.status == "OK"]
    done_keys = {(r.review_id, str(r.photo_index)) for r in ok.itertuples()}
    done_path = {(r.review_id, str(r.photo_index)): r.local_path for r in ok.itertuples()}   # later rows overwrite
    url_to_path = dict(zip(ok.photo_url, ok.local_path))
    jobs, dup_rows = [], []
    seen_urls_this_run: dict[str, tuple] = {}
    hotel_reviews: dict[str, list[dict]] = {}
    for _, h in pairs.iterrows():
        revs = sample_reviews(h)
        hotel_reviews[h.hotel_id] = revs
        for r in revs:
            for k, url in enumerate((r.get("image_urls_all") or [])[: C.MAX_IMAGES_PER_REVIEW], 1):
                key = (r["review_id"], str(k))
                base = {"pair_id": h.pair_id, "hotel_id": h.hotel_id, "review_id": r["review_id"],
                        "photo_index": k, "photo_url": url}
                stem = C.IMAGE_DIR / h.pair_id / h.group / h.hotel_id / f"{r['review_id']}_{k:02d}"
                if key in done_keys:
                    old = done_path.get(key, "")
                    expected = stem.parent.relative_to(C.ROOT).as_posix() + "/"
                    if old and not old.startswith(expected) and (C.ROOT / old).exists():
                        dup_rows.append({**base, "local_path": _copy_duplicate(old, stem), "status": "OK",
                                         "variant": "relocated_after_rematch", "download_url_used": "", "bytes": 0,
                                         "error": "", "ts": U.now_iso()})
                    continue
                if url in url_to_path and (C.ROOT / url_to_path[url]).exists():
                    dup_rows.append({**base, "local_path": _copy_duplicate(url_to_path[url], stem), "status": "OK",
                                     "variant": "copied_from_duplicate_url", "download_url_used": "", "bytes": 0,
                                     "error": "", "ts": U.now_iso()})
                    continue
                if url in seen_urls_this_run:
                    continue      # resolved after the first copy finishes (see below)
                seen_urls_this_run[url] = key
                existing = list(stem.parent.glob(stem.name + ".*")) if stem.parent.exists() else []
                existing = [p for p in existing if not p.name.endswith(".part") and p.stat().st_size > 0]
                if existing:
                    dup_rows.append({**base, "local_path": existing[0].relative_to(C.ROOT).as_posix(), "status": "OK",
                                     "variant": "already_on_disk", "download_url_used": "", "bytes": existing[0].stat().st_size,
                                     "error": "", "ts": U.now_iso()})
                    continue
                jobs.append({**base, "dest_stem": str(stem)})
    append_manifest(dup_rows)
    log.info("Images to download: %d (already done: %d, dedup/on-disk: %d)", len(jobs), len(done_keys), len(dup_rows))
    if jobs and not skip_download:
        session = requests.Session()
        session.headers.update({"User-Agent": C.USER_AGENT})
        done = 0
        with ThreadPoolExecutor(max_workers=C.IMAGE_WORKERS) as ex:
            futs = [ex.submit(download_one, session, j) for j in jobs]
            buf = []
            for f in as_completed(futs):
                res = f.result()
                res.pop("dest_stem", None)
                buf.append(res)
                done += 1
                if len(buf) >= 50:
                    append_manifest(buf)
                    buf = []
                    log.info("   images %d/%d", done, len(jobs))
            append_manifest(buf)
    # second-occurrence duplicates within this run → point at first copy
    man = load_manifest()
    ok = man[man.status == "OK"]
    url_to_path = dict(zip(ok.photo_url, ok.local_path))
    late = []
    have_keys = {(r.review_id, str(r.photo_index)) for r in man.itertuples()}
    for h_id, revs in hotel_reviews.items():
        prow = pairs[pairs.hotel_id == h_id].iloc[0]
        for r in revs:
            for k, url in enumerate((r.get("image_urls_all") or [])[: C.MAX_IMAGES_PER_REVIEW], 1):
                if (r["review_id"], str(k)) not in have_keys and url in url_to_path:
                    stem = C.IMAGE_DIR / prow.pair_id / prow.group / h_id / f"{r['review_id']}_{k:02d}"
                    late.append({"pair_id": prow.pair_id, "hotel_id": h_id, "review_id": r["review_id"],
                                 "photo_index": k, "photo_url": url, "local_path": _copy_duplicate(url_to_path[url], stem),
                                 "status": "OK", "variant": "copied_from_duplicate_url", "ts": U.now_iso()})
    append_manifest(late)
    return build_dataset(pairs, hotel_reviews)


def build_dataset(pairs: pd.DataFrame, hotel_reviews: dict[str, list[dict]]) -> pd.DataFrame:
    man = load_manifest()
    man = man.drop_duplicates(["review_id", "photo_index"], keep="last")   # manifest is append-only: last row wins
    pmap = {(r.review_id, str(r.photo_index)): r for r in man.itertuples()}
    state = CR.load_state()
    rows = []
    for _, h in pairs.iterrows():
        for r in hotel_reviews.get(h.hotel_id, []):
            urls = (r.get("image_urls_all") or [])
            row = {
                "pair_id": h.pair_id, "hotel_id": h.hotel_id, "group": h.group, "operational_type": h.operational_type,
                "hotel_name": h.hotel_name, "country": h.country, "city": h.city, "star_rating": h.star_rating,
                "hotel_overall_rating": h.overall_rating,
                "review_id": r["review_id"], "review_text": r.get("review_text", ""),
                "review_rating": r.get("review_rating"), "review_date": r.get("review_date"),
                "traveler_type": r.get("traveler_type"), "room_type": r.get("room_type"),
                "reviewer_name": r.get("reviewer_name"), "reviewer_country": r.get("reviewer_country"),
                "language": r.get("language"), "helpful_count": r.get("helpful_count"),
                "has_image": bool(urls), "image_count": len(urls),
            }
            n_ok = n_fail = 0
            for k in range(1, C.MAX_IMAGES_PER_REVIEW + 1):
                url = urls[k - 1] if k <= len(urls) else ""
                m = pmap.get((r["review_id"], str(k)))
                path = m.local_path if (m is not None and m.status == "OK") else ""
                n_ok += bool(path)
                n_fail += bool(m is not None and m.status == "FAILED")
                row[f"photo_{k}"] = path
                row[f"photo_url_{k}"] = url
            row.update({
                "photos_downloaded": n_ok, "photos_failed": n_fail,
                "text_source": r.get("text_source"), "review_title": r.get("review_title"),
                "rating_scale": r.get("rating_scale"), "stay_date": r.get("stay_date"),
                "hotel_reply_text": r.get("hotel_reply_text"), "reviewer_extra": r.get("reviewer_extra"),
                "crawl_sort_mode": state.get(h.hotel_id, {}).get("sort_mode"), "crawled_at": r.get("crawled_at"),
            })
            rows.append(row)
    cols = (["pair_id", "hotel_id", "group", "operational_type", "hotel_name", "country", "city", "star_rating",
             "hotel_overall_rating", "review_id", "review_text", "review_rating", "review_date", "traveler_type",
             "room_type", "reviewer_name", "reviewer_country", "language", "helpful_count", "has_image", "image_count"]
            + [f"photo_{k}" for k in range(1, 11)] + [f"photo_url_{k}" for k in range(1, 11)]
            + ["photos_downloaded", "photos_failed", "text_source", "review_title", "rating_scale", "stay_date",
               "hotel_reply_text", "reviewer_extra", "crawl_sort_mode", "crawled_at"])
    df = pd.DataFrame(rows, columns=cols)
    df.to_csv(C.REVIEWS_CSV, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_NONNUMERIC)
    log.info("Raw dataset: %d reviews → %s", len(df), C.REVIEWS_CSV.name)
    return df


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--build-only", action="store_true", help="rebuild 04 CSV without downloading")
    a = ap.parse_args()
    run(skip_download=a.build_only)
