"""
Stage 2 — Build the conventional (staffed) control pool and match exactly one control to each of
the 100 unmanned/automated hotels.

Control pool sources (all read-only)
  * Control_Seed_Pool sheet of 호텔_표본_재구성_PhysicalAI100_매칭계획.xlsx (103 hotels)
  * 유인호텔_리스트.xlsx (merged by Trip.com code; adds nothing new if identical)
  * inputs/manual_control_candidates.xlsx (optional; columns: country, city, hotel_name, trip_hotel_url)
  * live Trip.com discovery for cities without enough controls (China cities in this study)

Control eligibility: final type CONVENTIONAL — no service robot, no strong front-desk automation
signal. Any hotel that is (or was) an unmanned candidate is never used as a control.

Matching: one-to-one optimal assignment (Hungarian algorithm, scipy) maximising a score that
encodes the spec's priority order (config.MATCH_WEIGHTS). Countries are never mixed.

Outputs
  outputs/02_final_conventional_100.xlsx + outputs/03_final_hotel_pairs_100.xlsx  (when 100/100 matched)
  outputs/02_conventional_matched_DRAFT.xlsx + outputs/03_hotel_pairs_DRAFT.xlsx (otherwise)
  outputs/stages/S1_control_pool_validated.xlsx, outputs/stages/S2_matching_log.xlsx
  work/pairs_current.csv  (input of the crawler)
"""
from __future__ import annotations

import argparse
import math
import re

import numpy as np
import pandas as pd
from scipy.optimize import linear_sum_assignment

import common as U
import config as C

log = U.get_logger("match")
W = C.MATCH_WEIGHTS
DISCOVERED = C.WORK_DIR / "discovered_controls.json"
CONTROL_REJECTS = C.WORK_DIR / "control_rejects.json"     # {trip_code: reason} from post-crawl QA
INFEASIBLE = -1e6

# ------------------------------------------------------------------ pool construction

def load_seed_pool() -> pd.DataFrame:
    src = U.find_input(C.UNMANNED_SOURCE_NAME)
    sp = pd.read_excel(src, sheet_name="Control_Seed_Pool")
    rows = []
    for _, r in sp.iterrows():
        rows.append({
            "control_source": "Control_Seed_Pool", "source_id": U.clean_str(r["Source_ID"]),
            "country": U.clean_str(r["Country"]), "city": U.clean_str(r["City"]),
            "hotel_name": U.clean_str(r["Hotel"]), "seed_star": U.to_number(r["Trip.com_Star"]),
            "grade": U.clean_str(r["Validation_Grade"]),
            "desk_evidence": (f"Grade {U.clean_str(r['Validation_Grade'])}: {U.clean_str(r['Human_Frontdesk'])}; "
                              f"robot tag: {U.clean_str(r['Service_Robot_Public'])}"),
            "evidence_url": U.clean_str(r["Evidence_URL"]),
            "trip_hotel_url": U.clean_str(r["Trip.com_Detail"]),
            "trip_review_url": U.clean_str(r["Trip.com_Review"]),
        })
    df = pd.DataFrame(rows)
    # merge the separate conventional list (same hotels expected); add anything new
    try:
        conv = pd.read_excel(U.find_input(C.CONVENTIONAL_SOURCE_NAME), sheet_name="Conventional_100")
        city_country = dict(zip(df.city, df.country))
        known = set(df.trip_hotel_url.map(U.trip_code_from_url))
        extra = []
        for _, r in conv.iterrows():
            code = U.trip_code_from_url(r["Trip.com 상세"])
            if code and code not in known:
                extra.append({
                    "control_source": "유인호텔_리스트", "source_id": f"CONV_{int(r['No.']):03d}",
                    "country": city_country.get(U.clean_str(r["City"]), ""), "city": U.clean_str(r["City"]),
                    "hotel_name": U.clean_str(r["Hotel"]), "seed_star": U.to_number(r["Trip.com Star"]),
                    "grade": U.clean_str(r["검증등급"]),
                    "desk_evidence": f"Grade {U.clean_str(r['검증등급'])}: {U.clean_str(r['직원 프런트/유인운영'])}",
                    "evidence_url": U.clean_str(r["Trip.com 상세"]),
                    "trip_hotel_url": U.clean_str(r["Trip.com 상세"]),
                    "trip_review_url": U.clean_str(r["Trip.com 리뷰(크롤링)"]),
                })
        if extra:
            df = pd.concat([df, pd.DataFrame(extra)], ignore_index=True)
        log.info("유인호텔_리스트 merged: %d additional hotels", len(extra))
    except FileNotFoundError:
        pass
    # optional manual candidates
    try:
        man = pd.read_excel(U.find_input(C.MANUAL_CONTROL_CANDIDATES_NAME))
        man = man.rename(columns=str.lower)
        man = man[~man["hotel_name"].astype(str).str.upper().str.startswith("EXAMPLE")]
        for _, r in man.iterrows():
            df.loc[len(df)] = {
                "control_source": "manual", "source_id": f"MAN_{len(df):03d}",
                "country": U.clean_str(r.get("country")), "city": U.clean_str(r.get("city")),
                "hotel_name": U.clean_str(r.get("hotel_name")), "seed_star": U.to_number(r.get("star_rating")),
                "grade": "M", "desk_evidence": "manual candidate: " + U.clean_str(r.get("notes")),
                "evidence_url": U.clean_str(r.get("evidence_url")) or U.clean_str(r.get("trip_hotel_url")),
                "trip_hotel_url": U.clean_str(r.get("trip_hotel_url")),
                "trip_review_url": U.review_url_from_detail(U.clean_str(r.get("trip_hotel_url"))),
            }
        log.info("manual control candidates: %d", len(man))
    except FileNotFoundError:
        pass
    for rec in U.read_json(DISCOVERED, []) or []:
        if rec.get("trip_hotel_url") and rec.get("discovery_ok", True):
            df.loc[len(df)] = {k: rec.get(k, "") for k in df.columns}
    df["trip_code"] = df["trip_hotel_url"].map(U.trip_code_from_url)
    df = df.drop_duplicates("trip_code", keep="first").reset_index(drop=True)
    bg = df["hotel_name"].map(U.brand_of)
    df["brand"] = bg.map(lambda t: t[0])
    df["brand_group"] = bg.map(lambda t: t[1])
    return df


def unmanned_codes_all() -> set[str]:
    """Every Trip.com code that appears as an unmanned candidate (selected, reserve or excluded)."""
    import validate_hotels as V
    cand = V.load_candidates()
    return set(cand["trip_code"].dropna())


def classify_controls(pool: pd.DataFrame, metas: dict[str, dict]) -> pd.DataFrame:
    out = []
    rejects = U.read_json(CONTROL_REJECTS, {}) or {}
    for _, r in pool.iterrows():
        meta = metas.get(r["trip_code"], {})
        desk_conf = {"A": "MEDIUM", "M": "LOW"}.get(r["grade"], "LOW")
        live = meta.get("live_type") if meta.get("fetch_status") == "OK" else None
        kw = f"robot_kw=[{meta.get('live_robot_kw','')}], automation_kw=[{meta.get('live_automation_kw','')}]"
        if live is None or live == "UNKNOWN":
            ftype, fconf = C.TYPE_CONV, desk_conf
            note = "live check unavailable — desk evidence only" if live is None else f"facility text not parsed ({kw})"
        elif r["trip_code"] in rejects:
            ftype, fconf, note = C.TYPE_UNCERTAIN, "LOW", f"REJECTED as control: {rejects[r['trip_code']]}"
        elif live == C.TYPE_CONV and ((meta.get("first_page_robot_mentions") or 0) > 0 or meta.get("page_robot_kw_anywhere")):
            ftype, fconf, note = C.TYPE_UNCERTAIN, "LOW", "REJECTED as control: robot mentioned in guest reviews / review tags"
        elif live == C.TYPE_CONV:
            ftype, fconf = C.TYPE_CONV, ("HIGH" if desk_conf == "MEDIUM" else "MEDIUM")
            note = (f"Trip.com facility ids/tags on the public review page show no 'Service robots' and no "
                    f"front-desk automation ({kw}); no robot mention on first review page")
        elif live == C.TYPE_PA:
            ftype, fconf, note = C.TYPE_PA, "MEDIUM", f"REJECTED as control: robot listed ({kw})"
        else:
            ftype, fconf, note = C.TYPE_AO, "MEDIUM", f"REJECTED as control: front-desk automation listed ({kw})"
        star = meta.get("star_rating") or r["seed_star"]
        out.append({
            "final_type": ftype, "confidence": fconf, "live_note": note,
            "star_rating": star, "overall_rating": meta.get("overall_rating"),
            "overall_rating_scale": meta.get("overall_rating_scale"),
            "total_review_count": meta.get("total_review_count"),
            "price_value": meta.get("price_value"), "price_currency": meta.get("price_currency"),
            "live_automation_weak_kw": meta.get("live_automation_weak_kw", ""),
            "meta_fetch_status": meta.get("fetch_status", "NOT_RUN"),
        })
    return pd.concat([pool.reset_index(drop=True), pd.DataFrame(out)], axis=1)

# ------------------------------------------------------------------ live discovery

def _parse_hotel_list(obj, city: str, country: str) -> list[dict]:
    found = {}
    for d in U.iter_dicts(obj):
        base = d.get("hotelBasicInfo") if isinstance(d.get("hotelBasicInfo"), dict) else d
        code = base.get("hotelId") or base.get("masterHotelId")
        name = base.get("hotelName") or base.get("hotelEnName")
        if not code or not isinstance(name, str):
            continue
        code = str(code)
        if code in found:            # outer record (with jump URL) is seen before its nested copy
            continue
        url = U.first_scalar(d, ["hotelJumpUrl", "jumpUrl", "detailUrl", "url"], 2)
        if not (isinstance(url, str) and "hotel-detail" in url):
            url = U.detail_url_from_code(code, city)
        if url.startswith("/"):
            url = C.TRIP_BASE + url
        found[code] = {
            "control_source": "trip_discovery", "source_id": f"DISC_{code}", "country": country, "city": city,
            "hotel_name": name, "seed_star": U.to_number(U._search_keys([d], U.STAR_KEYS, lambda v: 0 < (U.to_number(v) or 0) <= 5)),
            "grade": "D", "desk_evidence": "discovered via Trip.com city list",
            "evidence_url": url, "trip_hotel_url": url, "trip_review_url": U.review_url_from_detail(url),
        }
    return list(found.values())


_LIST_LINKS_JS = r"""() => {
  const hotels = [], subs = [];
  for (const a of document.querySelectorAll('a[href]')) {
    const h = a.href.split('?')[0].split('#')[0];
    if (/-hotel-detail-\d+\/[^/]+\/?$/.test(h)) hotels.push(h);
    else if (/-hotels-list-\d+\/zone\d+\/?$/.test(h)) subs.push(h);
  }
  return {hotels: [...new Set(hotels)], subs: [...new Set(subs)]};
}"""

# Names that are almost always 5-star; tried last when the city's unmanned hotels are 3-4 star
_LUXURY_RE = (r"hyatt|hilton|intercontinental|sofitel|bvlgari|regent|waldorf|ritz|marriott|shangri|kempinski|"
              r"four seasons|peninsula|mandarin|regis|westin|sheraton|conrad|rosewood|fairmont|wanda|"
              r"kerry|langham|banyan|aman|raffles|pullman|nuo|w hotel|park hyatt|grand|palace|resort|"
              r"hostel|apartment|homestay|inn\b|capsule|villa|guesthouse|b&b|youth")


def _city_slug_and_id(sub: pd.DataFrame, city: str) -> tuple[str | None, int | None]:
    """Trip.com city url slug + id, read from the unmanned hotels' own cached review-page data."""
    for code in sub.trip_code.dropna().astype(str):
        m = U.read_json(C.META_CACHE_DIR / f"{code}.json", {}) or {}
        if m.get("trip_city_url") and m.get("trip_city_id"):
            return m["trip_city_url"], int(m["trip_city_id"])
    for url in sub.trip_hotel_url.dropna():
        mm = re.search(r"/hotels/([a-z0-9-]+)-hotel-detail-", url)
        if mm and C.CITY_IDS.get(city):
            return mm.group(1), C.CITY_IDS[city]
    return None, None


def _collect_list_candidates(browser, city_url: str, city_id: int, want: int, city: str) -> list[str] | None:
    """Hotel detail URLs from the public city list page and its area ('zone') pages.
    None = blocked (caller stops discovery; nothing is bypassed)."""
    root = C.SEO_LIST_URL.format(city_url=city_url, city_id=city_id)
    queue, seen_pages, found = [root], set(), []
    prefix = f"/hotels/{city_url}-hotel-detail-"
    while queue and len(found) < want and len(seen_pages) < 14:
        url = queue.pop(0)
        if url in seen_pages:
            continue
        seen_pages.add(url)
        cap = None
        try:
            cap = browser.open(url, ["__no_capture__"])
            if cap.check_blocked(check_login_text=False):
                cap.snapshot(f"discovery_blocked_{city}")
                return None
            for _ in range(3):
                cap.page.mouse.wheel(0, 5000)
                U.polite_sleep((0.8, 1.6))
            res = cap.page.evaluate(_LIST_LINKS_JS)
            for h in res["hotels"]:
                if prefix in h and h not in found:
                    found.append(h)
            if url == root:
                queue.extend(s for s in res["subs"] if f"-hotels-list-{city_id}/" in s)
            log.info("   list page %s → %d candidate links so far", url.split("/hotels/")[1], len(found))
        except Exception as e:  # noqa: BLE001
            log.warning("discovery list error %s: %s", url, e)
        finally:
            if cap:
                try:
                    cap.page.close()
                except Exception:
                    pass
        U.polite_sleep(C.HOTEL_DELAY)
    return found


def control_is_acceptable(meta: dict, unmanned_stars: list[float]) -> tuple[bool, str]:
    """Conventional + comparable: no robot / front-desk automation on Trip.com, no robot talk on the
    first review page, enough reviews to sample, star within ±1 of an unmanned hotel in the city."""
    if meta.get("fetch_status") != "OK":
        return False, f"fetch {meta.get('fetch_status')}: {meta.get('fetch_error', '')}"
    if meta.get("live_type") != C.TYPE_CONV:
        return False, f"live_type={meta.get('live_type')} ({meta.get('live_robot_kw') or meta.get('live_automation_kw')})"
    names = f"{meta.get('trip_hotel_name', '')} | {meta.get('trip_hotel_name_cn', '')}"
    if re.search(C.NON_CONTROL_NAME_RE, names, re.I):
        return False, "property type / concept not a comparable conventional hotel (name)"
    if (meta.get("first_page_robot_mentions") or 0) > 0 or meta.get("page_robot_kw_anywhere"):
        return False, "robot mentioned in reviews / review tags"
    if (meta.get("total_review_count") or 0) < C.MIN_DISCOVERED_CONTROL_REVIEWS:
        return False, f"only {meta.get('total_review_count')} reviews"
    star = meta.get("star_rating")
    if unmanned_stars and star and min(abs(star - s) for s in unmanned_stars) > 1:
        return False, f"star {star} not within ±1 of {sorted(set(unmanned_stars))}"
    return True, ""


def discover(unmanned: pd.DataFrame, pool: pd.DataFrame, excluded_codes: set[str], run_cfg: dict,
             browser) -> int:
    """Find additional same-city conventional candidates for under-covered cities (live only).
    Source: Trip.com's public city list pages (+ area pages); every candidate is then verified on
    its own public review page (facility ids / tags)."""
    existing = U.read_json(DISCOVERED, []) or []
    known = set(pool.trip_code) | excluded_codes | {e.get("trip_code") for e in existing}
    unmanned = unmanned.copy()
    unmanned["trip_code"] = unmanned["trip_hotel_url"].map(U.trip_code_from_url)
    added = 0
    n_city = unmanned.groupby(["country", "city"]).size()
    need = (n_city + np.ceil(n_city * C.CONTROL_MARGIN).clip(lower=1)).rename("need")
    ok_pool = pool[(pool.final_type == C.TYPE_CONV) & (pool.meta_fetch_status == "OK")]
    have = ok_pool.groupby(["country", "city"]).size().rename("have")
    disc_ok = pd.Series([(e.get("country"), e.get("city")) for e in existing if e.get("discovery_ok")]).value_counts() \
        if existing else pd.Series(dtype=int)
    gaps = pd.concat([need, have], axis=1).fillna(0)
    gaps = gaps[gaps.have < gaps.need]
    for (country, city), g in gaps.iterrows():
        sub = unmanned[(unmanned.country == country) & (unmanned.city == city)]
        target = int(g.need - g.have)
        if target <= 0:
            continue
        stars = [float(s) for s in pd.to_numeric(sub.star_rating, errors="coerce").dropna()]
        city_url, city_id = _city_slug_and_id(sub, city)
        if not city_url:
            log.warning("discovery: no Trip.com city slug/id for %s — add manual candidates", city)
            continue
        log.info("discovery %s/%s: need %d more verified conventional hotels (unmanned stars %s)",
                 country, city, target, sorted(set(stars)))
        links = _collect_list_candidates(browser, city_url, city_id, want=max(40, target * 8), city=city)
        if links is None:
            log.warning("discovery blocked on %s — stopping discovery (no bypass)", city)
            return added
        cands = []
        for url in links:
            code = U.trip_code_from_url(url)
            if code and code not in known and code not in {c["trip_code"] for c in cands}:
                cands.append({"trip_code": code, "trip_hotel_url": url if url.endswith("/") else url + "/"})
        if stars and max(stars) <= 4:        # obvious luxury / non-hotel names last
            cands.sort(key=lambda c: bool(re.search(_LUXURY_RE, c["trip_hotel_url"].replace("-", " "))))
        verified, tried = 0, 0
        for h in cands:
            if verified >= target or tried >= target * 6 + 10:
                break
            tried += 1
            cached = (C.META_CACHE_DIR / f"{h['trip_code']}.json").exists()
            meta = U.fetch_hotel_meta(browser, h["trip_hotel_url"], run_cfg, f"disc_{city}")
            ok, why = control_is_acceptable(meta, stars)
            name = meta.get("trip_hotel_name") or h["trip_hotel_url"].rstrip("/").split("/")[-1]
            rec = {
                "control_source": "trip_discovery", "source_id": f"DISC_{h['trip_code']}", "country": country,
                "city": city, "hotel_name": name, "seed_star": meta.get("star_rating"), "grade": "D",
                "desk_evidence": "discovered via Trip.com public city list page; verified on the hotel's public "
                                 "review page (no 'Service robots' facility id/tag, no front-desk automation tag)",
                "evidence_url": U.review_url_from_detail(h["trip_hotel_url"]),
                "trip_hotel_url": h["trip_hotel_url"],
                "trip_review_url": U.review_url_from_detail(h["trip_hotel_url"]),
                "trip_code": h["trip_code"], "discovery_live_type": meta.get("live_type"),
                "discovery_ok": bool(ok), "discovery_reject_reason": why,
            }
            existing.append(rec)
            known.add(h["trip_code"])
            if ok:
                verified += 1
                added += 1
            log.info("   [%d] %s → %s %s", tried, name[:60], "OK" if ok else "rejected", why)
            U.write_json_atomic(DISCOVERED, existing)
            if not cached:
                U.polite_sleep(C.HOTEL_DELAY)
            if meta.get("fetch_status") == "BLOCKED":
                log.warning("discovery blocked while verifying — stopping discovery (no bypass)")
                return added
        log.info("discovery %s/%s: %d verified conventional candidates (%d tried, %d links)",
                 country, city, verified, tried, len(cands))
    return added

# ------------------------------------------------------------------ scoring & assignment

def _sim_log(a, b, span) -> float | None:
    if a is None or b is None or (isinstance(a, float) and math.isnan(a)) or (isinstance(b, float) and math.isnan(b)):
        return None
    if a <= 0 or b <= 0:
        return None
    return max(0.0, 1 - abs(math.log(a / b)) / math.log(span))


def score_pair(u: pd.Series, c: pd.Series) -> tuple[float, dict]:
    if u.country != c.country:
        return INFEASIBLE, {"reason": "different country"}
    comp = {}
    if u.city == c.city:
        comp["city"], level = W["same_city"], "SAME_CITY"
    elif c.city in C.NEARBY_CITIES.get((u.country, u.city), []):
        comp["city"], level = W["nearby_city"], "NEARBY_CITY"
    else:
        comp["city"], level = W["same_country"], "SAME_COUNTRY_ONLY"
    if u.brand != "Independent/Other" and u.brand == c.brand:
        comp["brand"], blevel = W["same_brand"], "SAME_BRAND"
    elif u.brand_group != "Independent/Other" and u.brand_group == c.brand_group:
        comp["brand"], blevel = W["same_group"], "SAME_GROUP"
    else:
        comp["brand"], blevel = 0.0, ""
    us, cs = U.to_number(u.get("star_rating")), U.to_number(c.get("star_rating"))
    if us is None or cs is None or (isinstance(us, float) and math.isnan(us)) or (isinstance(cs, float) and math.isnan(cs)):
        comp["star"] = W["star_unknown"]
    elif us == cs:
        comp["star"] = W["star_exact"]
    elif abs(us - cs) <= 1:
        comp["star"] = W["star_pm1"]
    else:
        comp["star"] = W["star_far"]
    same_cur = U.clean_str(u.get("price_currency")) and u.get("price_currency") == c.get("price_currency")
    ps = _sim_log(U.to_number(u.get("price_value")), U.to_number(c.get("price_value")), 2) if same_cur else None
    comp["price"] = W["price_max"] * ps if ps is not None else 0.0
    rs = _sim_log((U.to_number(u.get("total_review_count")) or 0) + 1 if u.get("total_review_count") is not None else None,
                  (U.to_number(c.get("total_review_count")) or 0) + 1 if c.get("total_review_count") is not None else None, 10)
    comp["reviews"] = W["reviews_max"] * rs if rs is not None else 0.0
    comp["area"] = W["area_token"] if (U.area_tokens(u.hotel_name) & U.area_tokens(c.hotel_name)) else 0.0
    comp["grade_a"] = W["grade_a"] if c.get("grade") == "A" else 0.0
    comp["level"] = level + (("+" + blevel) if blevel else "")
    return float(sum(v for k, v in comp.items() if k != "level")), comp


def assign(unmanned: pd.DataFrame, controls: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    n, m = len(unmanned), len(controls)
    S = np.full((n, max(m, 1)), INFEASIBLE)
    comps = {}
    for i, (_, u) in enumerate(unmanned.iterrows()):
        for j, (_, c) in enumerate(controls.iterrows()):
            s, comp = score_pair(u, c)
            S[i, j] = s
            comps[(i, j)] = comp
    rows, cols = linear_sum_assignment(-S) if m else (np.array([], int), np.array([], int))
    chosen = dict(zip(rows, cols))
    res, alts = [], []
    for i, (_, u) in enumerate(unmanned.iterrows()):
        j = chosen.get(i)
        ok = j is not None and S[i, j] > INFEASIBLE / 2
        comp = comps.get((i, j), {}) if ok else {}
        res.append({"pair_id": u.pair_id, "unmanned_hotel_id": u.hotel_id,
                    "control_idx": int(j) if ok else None, "match_score": round(S[i, j], 2) if ok else None,
                    "match_level": comp.get("level", "UNMATCHED"), **{f"score_{k}": v for k, v in comp.items() if k != "level"}})
        order = np.argsort(-S[i])[:4]
        for rank, jj in enumerate(order, 1):
            if S[i, jj] <= INFEASIBLE / 2:
                break
            alts.append({"pair_id": u.pair_id, "unmanned_hotel": u.hotel_name, "rank": rank,
                         "candidate": controls.iloc[jj].hotel_name, "candidate_city": controls.iloc[jj].city,
                         "score": round(S[i, jj], 2), "level": comps[(i, jj)]["level"],
                         "chosen": bool(ok and jj == j)})
    return pd.DataFrame(res), pd.DataFrame(alts)

# ------------------------------------------------------------------ reserve substitution

def substitute_from_reserve(unmanned: pd.DataFrame, eligible: pd.DataFrame, matches: pd.DataFrame,
                            metas: dict) -> tuple[pd.DataFrame, list[dict]]:
    """Replace unmanned hotels that have no feasible control with a RESERVE hotel (same country)
    that does. The reserve hotel inherits the hotel_id / pair_id slot; the swap is logged."""
    unmatched = matches[matches.match_level == "UNMATCHED"]
    if unmatched.empty:
        return unmanned, []
    try:
        res = pd.read_excel(C.STAGE1_CANDIDATES, sheet_name="reserve_and_excluded")
    except Exception:
        return unmanned, []
    res = res[res.selection_status == "RESERVE_BALANCE"].copy()
    used = set(matches.control_idx.dropna().astype(int))
    free = eligible.drop(index=list(used), errors="ignore")
    subs = []
    for _, m in unmatched.iterrows():
        i = unmanned.index[unmanned.pair_id == m.pair_id][0]
        u = unmanned.loc[i]
        for ri, r in res.iterrows():
            if r.country != u.country:
                continue
            meta = metas.get(U.clean_str(r.trip_code), {})
            cand = pd.Series({**r.to_dict(), "brand_group": U.brand_of(r.hotel_name)[1],
                              "star_rating": meta.get("star_rating", r.get("star_rating")),
                              "price_value": meta.get("price_value", r.get("price_value")),
                              "price_currency": meta.get("price_currency", r.get("price_currency")),
                              "total_review_count": meta.get("total_review_count", r.get("total_review_count"))})
            if any(score_pair(cand, c)[0] > INFEASIBLE / 2 and score_pair(cand, c)[1].get("level", "").startswith(("SAME_CITY", "NEARBY"))
                   for _, c in free.iterrows()):
                new = u.copy()
                for col in ["hotel_name", "country", "city", "brand", "brand_group", "star_rating", "overall_rating",
                            "total_review_count", "trip_hotel_url", "trip_review_url", "evidence_url",
                            "price_value", "price_currency"]:
                    new[col] = cand.get(col)
                new["operational_type"] = r.final_type
                new["confidence"] = r.final_confidence
                new["classification_evidence"] = f"{r.desk_evidence} | LIVE: {r.live_note} | SUBSTITUTED FROM RESERVE"
                unmanned.loc[i] = new
                subs.append({"pair_id": u.pair_id, "hotel_id": u.hotel_id, "removed": u.hotel_name,
                             "removed_reason": "no feasible same-city/nearby control", "substitute": r.hotel_name,
                             "substitute_candidate_id": r.candidate_id})
                res = res.drop(index=ri)
                break
    return unmanned, subs


# ------------------------------------------------------------------ main

SPEC_COLS = ["hotel_id", "pair_id", "group", "operational_type", "hotel_name", "country", "city", "brand",
             "star_rating", "overall_rating", "total_review_count", "trip_hotel_url", "trip_review_url",
             "classification_evidence", "evidence_url", "confidence", "price_value", "price_currency", "price_basis"]

MATCH_RULES = [
    ("Hard constraint", "Never match across countries."),
    ("Priority", "1 same city, 2 same brand (or same parent group), 3 same star, 4 star ±1, 5 similar price "
                 "(same stay conditions), 6 similar review volume. Implemented as additive weights whose "
                 "magnitudes make the ordering lexicographic (see config.MATCH_WEIGHTS)."),
    ("Fallbacks", "NEARBY_CITY = same metro area/prefecture (config.NEARBY_CITIES); SAME_COUNTRY_ONLY is a last "
                  "resort and is flagged."),
    ("Assignment", "Global one-to-one optimum (Hungarian algorithm) so no control is reused and no pair is "
                   "matched greedily at another pair's expense."),
    ("Control eligibility", "CONVENTIONAL only: hotels with a robot or strong self-check-in/kiosk/hologram "
                            "signal on Trip.com are rejected and logged."),
    ("Provisional", "In offline mode unmanned star/price/review counts are unknown, so matching relies on city, "
                    "neighbourhood tokens and source grade; the live run recomputes everything."),
]


def run(offline: bool, skip_discovery: bool = False) -> pd.DataFrame:
    U.ensure_dirs()
    run_cfg = U.load_run_config()
    unmanned = pd.read_excel(C.FINAL_UNMANNED, sheet_name="unmanned_automated_100")
    bg = unmanned["hotel_name"].map(U.brand_of)
    unmanned["brand_group"] = [g if b != "Independent/Other" else "Independent/Other" for b, g in bg]
    excluded_codes = unmanned_codes_all()
    pool = load_seed_pool()
    clash = pool.trip_code.isin(excluded_codes)
    if clash.any():
        log.warning("%d pool hotels are unmanned candidates → removed from control pool", clash.sum())
    pool = pool[~clash].reset_index(drop=True)

    metas: dict[str, dict] = {}
    browser_cm = None
    if not offline:
        browser_cm = U.TripBrowser()
        br = browser_cm.__enter__()
        cities = set(zip(unmanned.country, unmanned.city))
        nearby = {(k[0], c) for k, v in C.NEARBY_CITIES.items() for c in v}
        relevant = pool[[(a, b) in cities or (a, b) in nearby for a, b in zip(pool.country, pool.city)]]
        for n, (_, r) in enumerate(relevant.iterrows(), 1):
            cached = (C.META_CACHE_DIR / f"{r.trip_code}.json").exists()
            log.info("[%d/%d] control live check %s", n, len(relevant), r.hotel_name)
            metas[r.trip_code] = U.fetch_hotel_meta(br, r.trip_hotel_url, run_cfg, r.source_id)
            if not cached:
                U.polite_sleep(C.HOTEL_DELAY)
    pool = classify_controls(pool, metas)

    if not offline and not skip_discovery:
        added = discover(unmanned, pool, excluded_codes, run_cfg, br)
        log.info("discovery added %d verified candidates", added)
        if added:
            pool = load_seed_pool()
            pool = pool[~pool.trip_code.isin(excluded_codes)].reset_index(drop=True)
            for code in pool.trip_code:
                cache = U.read_json(C.META_CACHE_DIR / f"{code}.json")
                if cache:
                    metas[code] = cache
            pool = classify_controls(pool, metas)
    if browser_cm:
        browser_cm.__exit__(None, None, None)

    eligible = pool[pool.final_type == C.TYPE_CONV]
    if not offline:      # a control must be reachable and have reviews to sample
        usable = (eligible.meta_fetch_status == "OK") & (pd.to_numeric(eligible.total_review_count, errors="coerce").fillna(0) >= C.MIN_CONTROL_REVIEWS)
        log.info("eligible conventional controls: %d (dropped %d unreachable / too few reviews)", usable.sum(), (~usable).sum())
        eligible = eligible[usable]
    eligible = eligible.reset_index(drop=True)
    matches, alts = assign(unmanned, eligible)
    unmanned, subs = substitute_from_reserve(unmanned, eligible, matches, metas)
    if subs:
        log.info("substituted %d unmanned hotels from RESERVE → re-matching", len(subs))
        matches, alts = assign(unmanned, eligible)
        old_notes = pd.read_excel(C.FINAL_UNMANNED, sheet_name="Notes")
        notes01 = list(old_notes.itertuples(index=False, name=None)) + [
            ("substitution", f"{s['pair_id']}: {s['removed']} → {s['substitute']} ({s['removed_reason']})") for s in subs]
        cur = pd.read_excel(C.FINAL_UNMANNED, sheet_name="unmanned_automated_100")
        cur = cur.set_index("pair_id")
        upd = unmanned.set_index("pair_id")
        for col in cur.columns:
            if col in upd.columns:
                cur[col] = upd[col]
        U.write_excel(C.FINAL_UNMANNED, {"unmanned_automated_100": cur.reset_index()[
            ["hotel_id", "pair_id"] + [c for c in cur.columns if c != "hotel_id"]]}, notes01)

    # ---- assemble outputs
    ctrl_rows = []
    for _, m in matches.iterrows():
        if m.control_idx is None or (isinstance(m.control_idx, float) and math.isnan(m.control_idx)):
            continue
        c = eligible.iloc[int(m.control_idx)]
        k = int(m.pair_id.split("_")[1])
        ctrl_rows.append({
            "hotel_id": f"H{100 + k:03d}", "pair_id": m.pair_id, "group": C.GROUP_CONVENTIONAL,
            "operational_type": C.TYPE_CONV, "hotel_name": c.hotel_name, "country": c.country, "city": c.city,
            "brand": c.brand, "star_rating": c.star_rating, "overall_rating": c.overall_rating,
            "total_review_count": c.total_review_count, "trip_hotel_url": c.trip_hotel_url,
            "trip_review_url": c.trip_review_url,
            "classification_evidence": f"{c.desk_evidence} | LIVE: {c.live_note}",
            "evidence_url": c.evidence_url, "confidence": c.confidence,
            "price_value": c.price_value, "price_currency": c.price_currency,
            "price_basis": U.price_basis_text(run_cfg) if pd.notna(c.price_value) else "",
            "match_level": m.match_level, "match_score": m.match_score,
            "matched_unmanned_hotel_id": m.unmanned_hotel_id, "control_source": c.control_source,
            "control_source_id": c.source_id, "trip_code": c.trip_code,
            "match_status": "PROVISIONAL_OFFLINE" if offline else "MATCHED",
        })
    controls = pd.DataFrame(ctrl_rows, columns=SPEC_COLS + ["match_level", "match_score", "matched_unmanned_hotel_id",
                                                             "control_source", "control_source_id", "trip_code", "match_status"])
    um = unmanned[SPEC_COLS].merge(matches[["pair_id", "match_level", "match_score"]], on="pair_id", how="left")
    um["match_status"] = np.where(um.match_level == "UNMATCHED", "NO_CONTROL_YET",
                                  "PROVISIONAL_OFFLINE" if offline else "MATCHED")
    long = pd.concat([um, controls[SPEC_COLS + ["match_level", "match_score", "match_status"]]], ignore_index=True)
    long = long.sort_values(["pair_id", "group"], ascending=[True, False]).reset_index(drop=True)
    wide = unmanned[["pair_id", "hotel_id", "operational_type", "hotel_name", "country", "city", "brand", "star_rating"]] \
        .merge(controls[["pair_id", "hotel_id", "hotel_name", "city", "brand", "star_rating", "match_level", "match_score"]],
               on="pair_id", how="left", suffixes=("_unmanned", "_control"))
    n_matched = len(controls)
    complete = n_matched == C.N_PER_GROUP
    level_tab = matches.match_level.value_counts().rename_axis("match_level").reset_index(name="pairs")

    notes = [("generated_at", U.now_iso()), ("mode", "offline (provisional)" if offline else "live"),
             ("pairs matched", f"{n_matched} / {len(unmanned)}"),
             ("stay conditions (price)", U.price_basis_text(run_cfg))] + MATCH_RULES
    conv_path = C.FINAL_CONVENTIONAL if complete else C.DRAFT_CONVENTIONAL
    pairs_path = C.FINAL_PAIRS if complete else C.DRAFT_PAIRS
    for stale in (C.DRAFT_CONVENTIONAL, C.DRAFT_PAIRS) if complete else (C.FINAL_CONVENTIONAL, C.FINAL_PAIRS):
        stale.unlink(missing_ok=True)
    U.write_excel(conv_path, {"conventional": controls}, notes)
    U.write_excel(pairs_path, {"pairs_long_200": long, "pairs_wide": wide, "match_quality": level_tab}, notes)
    U.write_excel(C.STAGE1_CONTROLS, {
        "control_pool": pool,
        "rejected_controls": pool[pool.final_type != C.TYPE_CONV],
        "coverage_by_city": (unmanned.groupby(["country", "city"]).size().rename("unmanned_n").to_frame()
                             .join(eligible.groupby(["country", "city"]).size().rename("eligible_controls_same_city"))
                             .fillna(0).reset_index()),
    }, notes)
    U.write_excel(C.STAGE2_MATCHING_LOG, {"chosen_pairs": matches, "top_alternatives": alts,
                                          "unmatched": matches[matches.match_level == "UNMATCHED"],
                                          "substitutions": pd.DataFrame(subs, columns=["pair_id", "hotel_id", "removed",
                                              "removed_reason", "substitute", "substitute_candidate_id"])}, notes)
    crawl_cols = ["pair_id", "hotel_id", "group", "operational_type", "hotel_name", "country", "city", "star_rating",
                  "overall_rating", "total_review_count", "trip_hotel_url", "trip_review_url", "match_status"]
    long[crawl_cols].to_csv(C.PAIRS_CURRENT_CSV, index=False, encoding="utf-8-sig")
    log.info("Matched %d/%d pairs → %s", n_matched, len(unmanned), pairs_path.name)
    if not complete:
        log.warning("%d unmanned hotels have no eligible control yet (see S2_matching_log 'unmatched').",
                    len(unmanned) - n_matched)
    return long


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--offline", action="store_true", help="no network: match from existing pools only")
    ap.add_argument("--skip-discovery", action="store_true", help="live control checks but no city discovery")
    a = ap.parse_args()
    run(offline=a.offline, skip_discovery=a.skip_discovery)
