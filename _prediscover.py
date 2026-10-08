"""Run control discovery early for cities whose unmanned hotels are already crawled (meta cached).
Results are stored in work/discovered_controls.json and reused by match_controls.py."""
import sys
import pandas as pd
import common as U, config as C, match_controls as M
log = U.get_logger("prediscover")
countries = sys.argv[1].split(",") if len(sys.argv) > 1 else ["China"]
skip_cities = set(sys.argv[2].split(",")) if len(sys.argv) > 2 and sys.argv[2] else set()
if len(sys.argv) > 3:                      # separate result file so parallel runs never clobber each other
    M.DISCOVERED = C.WORK_DIR / f"discovered_controls_{sys.argv[3]}.json"
run_cfg = U.load_run_config()
un = pd.read_excel(C.FINAL_UNMANNED, sheet_name="unmanned_automated_100")
un["trip_code"] = un.trip_hotel_url.map(U.trip_code_from_url)
metas = {c: (U.read_json(C.META_CACHE_DIR / f"{c}.json", {}) or {}) for c in un.trip_code}
un["star_rating"] = [metas[c].get("star_rating") for c in un.trip_code]
ready = un.groupby(["country", "city"]).trip_code.apply(lambda s: all(metas[c].get("fetch_status") == "OK" for c in s))
ready = {k for k, v in ready.items() if v and k[0] in countries and k[1] not in skip_cities}
sub = un[[(a, b) in ready for a, b in zip(un.country, un.city)]]
log.info("cities ready for discovery: %s", sorted(ready))
excluded = M.unmanned_codes_all()
pool = M.load_seed_pool()
pool = pool[~pool.trip_code.isin(excluded)].reset_index(drop=True)
pm = {c: (U.read_json(C.META_CACHE_DIR / f"{c}.json", {}) or {}) for c in pool.trip_code}
pm = {c: {**m, **U.classify_live_from_meta(m)} for c, m in pm.items() if m.get("fetch_status") == "OK"}
pool = M.classify_controls(pool, pm)
with U.TripBrowser() as br:
    added = M.discover(sub, pool, excluded, run_cfg, br)
log.info("prediscovery added %d verified candidates", added)
