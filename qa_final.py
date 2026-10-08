"""
Final QA of the live run (read-only except for --write-rejects).

Checks
  * 01 = exactly 100 unmanned/automated, 02 = exactly 100 conventional, 03 = 100 pairs / 200 hotels
  * no duplicate hotel / Trip.com hotel code anywhere in the 200
  * same-city share, star difference distribution
  * every conventional hotel: no 'Service robots' facility id/tag on Trip.com, and how many of its
    collected reviews mention a robot (a control whose guests talk about a service robot is not a
    conventional hotel, whatever its facility list says)
  * 04 CSV: review counts per group, photo columns point to files that exist

  python qa_final.py                    # print report, write outputs/07_final_qa.xlsx
  python qa_final.py --write-rejects    # also add robot-mention controls to work/control_rejects.json
"""
from __future__ import annotations

import argparse
import sys

import pandas as pd

import common as U
import config as C
import crawl_tripcom_reviews as CR
import match_controls as M

QA_XLSX = C.OUTPUT_DIR / "07_final_qa.xlsx"
ROBOT_MENTION_REJECT = 2          # ≥ this many sampled reviews mentioning a robot → reject the control


def robot_mentions(code: str) -> tuple[int, int]:
    revs = [r for r in CR.load_reviews(code).values()]
    hits = sum(1 for r in revs if U.keyword_hits(r.get("review_text", ""), C.ROBOT_KEYWORDS))
    return hits, len(revs)


def run(write_rejects: bool = False) -> dict:
    sys.stdout.reconfigure(encoding="utf-8")
    res: dict = {}
    un = pd.read_excel(C.FINAL_UNMANNED, sheet_name="unmanned_automated_100")
    conv_path = C.FINAL_CONVENTIONAL if C.FINAL_CONVENTIONAL.exists() else C.DRAFT_CONVENTIONAL
    pairs_path = C.FINAL_PAIRS if C.FINAL_PAIRS.exists() else C.DRAFT_PAIRS
    cv = pd.read_excel(conv_path, sheet_name="conventional")
    long = pd.read_excel(pairs_path, sheet_name="pairs_long_200")
    wide = pd.read_excel(pairs_path, sheet_name="pairs_wide")
    long["trip_code"] = long.trip_hotel_url.map(U.trip_code_from_url)
    res["01_file"], res["01_rows"] = C.FINAL_UNMANNED.name, len(un)
    res["PHYSICAL_AI"] = int((un.operational_type == C.TYPE_PA).sum())
    res["AUTOMATION_ONLY"] = int((un.operational_type == C.TYPE_AO).sum())
    res["02_file"], res["02_rows"] = conv_path.name, len(cv)
    res["03_file"], res["03_hotels"] = pairs_path.name, len(long)
    res["03_pairs_complete"] = int((long.groupby("pair_id").group.nunique() == 2).sum())
    res["dup_trip_code"] = int(long.trip_code.duplicated().sum())
    res["dup_hotel_name_city"] = int(long.duplicated(["hotel_name", "city"]).sum())
    res["dup_hotel_id"] = int(long.hotel_id.duplicated().sum())
    res["controls_that_are_unmanned_candidates"] = int(
        long[long.group == C.GROUP_CONVENTIONAL].trip_code.isin(M.unmanned_codes_all()).sum())
    w = wide.copy()
    res["same_city_pairs"] = int((w.city_unmanned == w.city_control).sum())
    sd = (pd.to_numeric(w.star_rating_unmanned, errors="coerce") - pd.to_numeric(w.star_rating_control, errors="coerce")).abs()
    res["star_diff_0"] = int((sd == 0).sum())
    res["star_diff_1"] = int((sd == 1).sum())
    res["star_diff_gt1"] = int((sd > 1).sum())
    res["star_unknown"] = int(sd.isna().sum())

    rows, rejects = [], (U.read_json(M.CONTROL_REJECTS, {}) or {})
    state = CR.load_state_raw()
    for r in long.itertuples():
        meta = U.read_json(C.META_CACHE_DIR / f"{r.trip_code}.json", {}) or {}
        hits, n = robot_mentions(r.trip_code)
        robot_tag = bool(set(meta.get("facility_ids") or []) & set(C.ROBOT_FACILITY_IDS)) or \
            bool(U.keyword_hits("\n".join(meta.get("facility_hot_tags") or []), C.ROBOT_KEYWORDS))
        flag = ""
        if r.group == C.GROUP_CONVENTIONAL:
            if robot_tag:
                flag = "ROBOT_FACILITY_TAG"
            elif hits >= ROBOT_MENTION_REJECT:
                flag = f"ROBOT_IN_{hits}_REVIEWS"
            if flag and write_rejects:
                rejects[r.trip_code] = f"post-crawl QA: {flag} ({hits}/{n} collected reviews mention a robot)"
        st = state.get(r.trip_code, {})
        rows.append({"pair_id": r.pair_id, "hotel_id": r.hotel_id, "group": r.group,
                     "operational_type": r.operational_type, "hotel_name": r.hotel_name, "city": r.city,
                     "trip_code": r.trip_code, "star_rating": r.star_rating,
                     "trip_robot_facility_tag": robot_tag, "reviews_collected_all_orders": n,
                     "reviews_mentioning_robot": hits, "crawl_status": st.get("status", "NOT_CRAWLED"),
                     "failure_reason": st.get("failure_reason", ""), "qa_flag": flag})
    hq = pd.DataFrame(rows)
    res["controls_flagged"] = int((hq.qa_flag != "").sum())
    un_q = hq[hq.group == C.GROUP_UNMANNED]
    res["unmanned_with_trip_robot_tag"] = int(un_q.trip_robot_facility_tag.sum())
    res["unmanned_with_robot_in_reviews"] = int((un_q.reviews_mentioning_robot > 0).sum())
    if write_rejects:
        U.write_json_atomic(M.CONTROL_REJECTS, rejects)

    if C.REVIEWS_CSV.exists():
        d = pd.read_csv(C.REVIEWS_CSV, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        res["04_rows"] = len(d)
        res["04_dup_review_rows"] = int(d.duplicated(["hotel_id", "review_id"]).sum())
        res["04_reviews_unmanned"] = int((d.group == C.GROUP_UNMANNED).sum())
        res["04_reviews_conventional"] = int((d.group == C.GROUP_CONVENTIONAL).sum())
        res["04_reviews_with_photo"] = int((d.has_image.str.lower() == "true").sum())
        res["04_empty_text"] = int((d.review_text.str.strip() == "").sum())
        photo_cols = [f"photo_{k}" for k in range(1, 11)]
        paths = [p for c in photo_cols for p in d[c] if p]
        res["04_photo_cells"] = len(paths)
        res["04_photo_files_missing"] = sum(1 for p in paths if not (C.ROOT / p).exists())
        res["04_photo_url_cells"] = int(sum((d[f"photo_url_{k}"] != "").sum() for k in range(1, 11)))
        res["04_hotels_with_reviews"] = int(d.hotel_id.nunique())
        files = [p for p in C.IMAGE_DIR.rglob("*") if p.is_file()]
        res["image_files_on_disk"] = len(files)
    U.write_excel(QA_XLSX, {"checks": pd.DataFrame(list(res.items()), columns=["check", "value"]),
                            "hotel_qa": hq, "flagged_controls": hq[hq.qa_flag != ""]},
                  [("generated_at", U.now_iso()),
                   ("reviews_mentioning_robot", "keyword search (config.ROBOT_KEYWORDS) over every collected review of the hotel"),
                   ("qa_flag", f"conventional hotel with a Trip.com robot facility tag, or ≥{ROBOT_MENTION_REJECT} collected reviews mentioning a robot")])
    for k, v in res.items():
        print(f"{k}: {v}")
    if len(hq[hq.qa_flag != ""]):
        print(hq[hq.qa_flag != ""][["pair_id", "hotel_id", "hotel_name", "city", "reviews_mentioning_robot", "qa_flag"]].to_string())
    return res


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--write-rejects", action="store_true")
    run(write_rejects=ap.parse_args().write_rejects)
