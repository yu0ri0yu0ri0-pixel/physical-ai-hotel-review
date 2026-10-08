"""Targeted extra discovery: cities where a 3-star unmanned hotel could only be paired with a 5-star
control. Verifies additional candidates whose star is within ±1 of the 3-star hotels."""
import sys
import pandas as pd
import common as U, config as C, match_controls as M
log = U.get_logger("starfix")
cities = sys.argv[1].split(",")
run_cfg = U.load_run_config()
un = pd.read_excel(C.FINAL_UNMANNED, sheet_name="unmanned_automated_100")
sub = un[un.city.isin(cities) & (un.star_rating <= 3)]
excluded = M.unmanned_codes_all()
pool = M.load_seed_pool()
pool = pool[~pool.trip_code.isin(excluded)].reset_index(drop=True)
known_codes = set(pool.trip_code)
pm = {c: (U.read_json(C.META_CACHE_DIR / f"{c}.json", {}) or {}) for c in pool.trip_code}
pool = M.classify_controls(pool, {c: m for c, m in pm.items() if m.get("fetch_status") == "OK"})
# count only controls that are star-compatible with the 3-star hotels as "have"
pool_view = pool[~(pool.city.isin(cities) & (pd.to_numeric(pool.star_rating, errors="coerce") > 4))]
have = pool_view[pool_view.city.isin(cities)].groupby("city").size()
log.info("3-star unmanned per city: %s | star-compatible controls: %s", sub.groupby("city").size().to_dict(), have.to_dict())
C.CONTROL_MARGIN = 0.34
pool_view = pool_view[~pool_view.city.isin(cities)]          # force a gap of n(+margin) for these cities
# keep every already known code out of the candidate list
ex = excluded | known_codes
with U.TripBrowser() as br:
    added = M.discover(sub, pool_view, ex, run_cfg, br)
log.info("star-fix discovery added %d", added)
