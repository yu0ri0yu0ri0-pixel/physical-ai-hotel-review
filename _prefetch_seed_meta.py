"""Light Trip.com check (public review page, one visit each) of the seed control pool."""
import pandas as pd
import common as U, config as C, match_controls as M
log = U.get_logger("prefetch")
C.HOTEL_DELAY = (8.0, 14.0)
run_cfg = U.load_run_config()
unmanned = pd.read_excel(C.FINAL_UNMANNED, sheet_name="unmanned_automated_100")
pool = M.load_seed_pool()
pool = pool[~pool.trip_code.isin(M.unmanned_codes_all())]
cities = set(zip(unmanned.country, unmanned.city)) | {(k[0], c) for k, v in C.NEARBY_CITIES.items() for c in v}
rel = pool[[(a, b) in cities for a, b in zip(pool.country, pool.city)]]
log.info("seed pool %d, relevant %d", len(pool), len(rel))
with U.TripBrowser() as br:
    for n, (_, r) in enumerate(rel.iterrows(), 1):
        cached = (C.META_CACHE_DIR / f"{r.trip_code}.json").exists()
        m = U.fetch_hotel_meta(br, r.trip_hotel_url, run_cfg, r.source_id)
        log.info("[%d/%d] %s | %s | live=%s star=%s reviews=%s", n, len(rel), r.hotel_name[:50], m.get("fetch_status"),
                 m.get("live_type"), m.get("star_rating"), m.get("total_review_count"))
        if not cached:
            U.polite_sleep(C.HOTEL_DELAY)
