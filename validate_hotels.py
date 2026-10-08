"""
Stage 1 — Validate unmanned/automated hotel candidates and select exactly 100.

Inputs (read-only)
  inputs/호텔_표본_재구성_PhysicalAI100_매칭계획.xlsx  (sheets Physical_AI_Final_100, Excluded_9)

What it does
  1. Builds a candidate table from BOTH sheets (the 9 Henn-na branches excluded from the old
     "Physical AI only" design are AUTOMATION_ONLY candidates under the new 3-type design).
  2. Desk classification from the evidence already recorded in the source workbook.
  3. (live mode) Opens each Trip.com detail page, reads facility text, star, rating, review count
     and a same-condition price, and combines live evidence with desk evidence.
  4. Selects exactly 100 PHYSICAL_AI + AUTOMATION_ONLY hotels. If more than 100 are eligible, it
     removes candidates from the most over-represented brand x city cell (only candidates with
     generic, non-HIGH evidence are removable) and keeps them as a RESERVE for substitution.

Outputs
  outputs/01_final_unmanned_automated_100.xlsx
  outputs/stages/S1_unmanned_candidates_validated.xlsx  (all candidates, decisions, QA flags)

Usage
  python validate_hotels.py --offline        # desk evidence only (no network)
  python validate_hotels.py                  # desk + live Trip.com verification
  python validate_hotels.py --reselect       # redo selection even if a live selection exists
"""
from __future__ import annotations

import argparse
import re
from urllib.parse import urlparse

import pandas as pd

import common as U
import config as C

log = U.get_logger("validate")
SELECTION_STATE = C.WORK_DIR / "selection_state.json"

CLASSIFICATION_RULES = [
    ("PHYSICAL_AI", "A physical robot performs a customer-facing service function on site (front-desk/"
                    "reception robot, delivery robot, concierge/guide robot, F&B delivery robot)."),
    ("AUTOMATION_ONLY", "No physical service robot confirmed, but the core touch point (check-in/check-out/"
                        "front desk) is automated: self check-in kiosk, automated or hologram reception, "
                        "unmanned reception, mobile self check-in. An in-room tablet alone does NOT qualify. "
                        "'Express check-in/check-out' as a Trip.com facility label is treated as a weak "
                        "signal only (it usually denotes a staffed fast-track service)."),
    ("CONVENTIONAL", "Main customer services are staff-delivered; no service robot and no core front-desk "
                     "automation."),
    ("UNCERTAIN", "Evidence conflicting or insufficient; excluded from both groups."),
    ("Branch-level rule", "Chain membership is never sufficient. Every branch is judged on its own evidence "
                          "(e.g. Henn-na branches split into PHYSICAL_AI and AUTOMATION_ONLY)."),
    ("Confidence HIGH", "Property-specific, current, primary evidence (official branch page / press release) "
                        "or a Trip.com facility listing verified in the current review cycle; or desk and "
                        "live Trip.com evidence agree."),
    ("Confidence MEDIUM", "OTA facility tag carried over from an earlier list and not yet re-verified, or "
                          "automation evidence that is indirect (e.g. chain-level page)."),
    ("Confidence LOW", "Chain-level only or conflicting evidence."),
    ("Evidence priority", "1 Trip.com detail page, 2 official hotel site, 3 official chain site, "
                          "4 credible press, 5 other OTAs."),
]

SELECTION_RULES = [
    ("Eligible", "final operational_type in {PHYSICAL_AI, AUTOMATION_ONLY}."),
    ("Exactly 100", "If >100 eligible: iteratively remove one hotel from the brand x city cell with the "
                    "most hotels; within the cell remove the hotel with the least specific evidence (generic "
                    "'Service robots' tag; non-HIGH confidence first, then any generic-tag hotel); ties -> "
                    "larger city total, then the "
                    "later source row. Removed hotels are kept as RESERVE (valid, substitutable)."),
    ("Rationale", "Reduces single-chain/single-city concentration (JI Hotel Hangzhou) while preserving "
                  "every AUTOMATION_ONLY hotel and every hotel with specific robot evidence."),
    ("If <100", "Shortfall is reported; in live mode match_controls.py may substitute from RESERVE, and new "
                "candidates can be added via Trip.com discovery (see README)."),
]

# ------------------------------------------------------------------ desk classification

def _branch_specific(url: str) -> bool:
    if not isinstance(url, str) or not url:
        return False
    path = [p for p in urlparse(url).path.split("/") if p and p not in ("en", "ja", "ko", "zh")]
    # a branch page (e.g. /en/sendai-kokubuncho/concept/); news posts are brand-level announcements
    return len(path) >= 1 and path[0] != "news" and "trip.com" not in url


def _brand(name: str, source_brand: str) -> tuple[str, str]:
    b, g = U.brand_of(name)
    if b != "Independent/Other":
        return b, g
    sb = U.clean_str(source_brand)
    if sb and sb.lower() != U.clean_str(name).lower() and sb != "Independent/Other":
        return sb, sb
    return "Independent/Other", "Independent/Other"


def load_candidates() -> pd.DataFrame:
    src = U.find_input(C.UNMANNED_SOURCE_NAME)
    final = pd.read_excel(src, sheet_name="Physical_AI_Final_100")
    excl = pd.read_excel(src, sheet_name="Excluded_9")
    rows = []
    for _, r in final.iterrows():
        status = U.clean_str(r["Validation_Status"])
        if status in ("KEEP_CONFIRMED_PHYSICAL_ROBOT", "KEEP_CONFIRMED_TRIPCOM_ROBOT_TAG"):
            conf = "HIGH"
        else:  # KEEP_EXISTING_CANDIDATE: Trip.com robot tag carried over, not independently re-audited
            conf = "MEDIUM"
        brand, group = _brand(r["Hotel"], r["Brand/Chain"])
        rows.append({
            "candidate_id": f"UC{int(r['Final_No']):03d}",
            "source_sheet": "Physical_AI_Final_100",
            "source_row_no": int(r["Final_No"]),
            "country": U.clean_str(r["Country"]), "city": U.clean_str(r["City"]),
            "hotel_name": U.clean_str(r["Hotel"]), "brand": brand, "brand_group": group,
            "technology_detail": U.clean_str(r["Physical_AI_Type"]),
            "prior_status": status,
            "desk_type": C.TYPE_PA, "desk_confidence": conf,
            "desk_evidence": f"[{status}] {U.clean_str(r['Physical_AI_Type'])} — {U.clean_str(r['Inclusion_Rationale'])}",
            "evidence_url": U.clean_str(r["Evidence_URL"]),
            "trip_hotel_url": U.clean_str(r["Trip.com_Detail"]),
            "trip_review_url": U.clean_str(r["Trip.com_Review"]),
        })
    for _, r in excl.iterrows():
        new_class = U.clean_str(r["New_Class"])
        ev_url = U.clean_str(r["Evidence_URL"])
        if new_class == "AUTOMATION_ONLY":
            dtype, conf = C.TYPE_AO, ("HIGH" if _branch_specific(ev_url) else "MEDIUM")
            note = "hologram / automated front desk; no physical robot at this branch"
        elif new_class == "CONVENTIONAL_OR_AUTOMATED":
            dtype, conf = C.TYPE_AO, "MEDIUM"
            note = "official notice: no robots/video tech, Express (machine) check-in → automated check-in"
        else:  # UNCERTAIN_EXCLUDE (robot presence unclear, Express Check-In confirmed)
            dtype, conf = C.TYPE_AO, "MEDIUM"
            note = "Express Check-In confirmed; robot presence unclear → conservatively AUTOMATION_ONLY"
        brand, group = _brand(r["Hotel"], "")
        rows.append({
            "candidate_id": f"UX{int(r['Original_No']):03d}",
            "source_sheet": "Excluded_9",
            "source_row_no": 100 + int(r["Original_No"]),
            "country": U.clean_str(r["Country"]), "city": U.clean_str(r["City"]),
            "hotel_name": U.clean_str(r["Hotel"]), "brand": brand, "brand_group": group,
            "technology_detail": U.clean_str(r["Previous_Type"]),
            "prior_status": f"{U.clean_str(r['Decision'])}/{new_class}",
            "desk_type": dtype, "desk_confidence": conf,
            "desk_evidence": f"[re-classified under 3-type scheme] {note}. Source note: {U.clean_str(r['Exclusion_Rationale'])}",
            "evidence_url": ev_url,
            "trip_hotel_url": U.clean_str(r["Trip.com_Detail"]),
            "trip_review_url": U.clean_str(r["Trip.com_Review"]),
        })
    df = pd.DataFrame(rows)
    df["trip_code"] = df["trip_hotel_url"].map(U.trip_code_from_url)
    df["evidence_specific"] = ~df["technology_detail"].str.strip().str.lower().isin(["service robots"])
    return df


def quality_flags(df: pd.DataFrame, control_codes: set[str]) -> pd.Series:
    flags = []
    dup = df["trip_code"].duplicated(keep=False)
    for i, r in df.iterrows():
        f = []
        if not r["trip_code"]:
            f.append("NO_TRIP_CODE")
        ev_code = U.trip_code_from_url(r["evidence_url"])
        if ev_code and r["trip_code"] and ev_code != r["trip_code"]:
            f.append(f"EVIDENCE_URL_ID_MISMATCH({ev_code}≠{r['trip_code']})")
        if "under-preparation" in (r["trip_hotel_url"] or ""):
            f.append("TRIP_SLUG_UNDER_PREPARATION(low review volume risk)")
        if dup[i]:
            f.append("DUPLICATE_TRIP_CODE")
        if r["trip_code"] in control_codes:
            f.append("ALSO_IN_CONTROL_POOL")
        flags.append("; ".join(f))
    return pd.Series(flags, index=df.index)

# ------------------------------------------------------------------ live combination

def combine(desk_type: str, desk_conf: str, meta: dict) -> tuple[str, str, str]:
    if not meta or meta.get("fetch_status") != "OK":
        why = (meta or {}).get("fetch_error") or "not run (offline)"
        return desk_type, desk_conf, f"live Trip.com check unavailable: {why}"
    live = meta.get("live_type")
    kw = f"robot_kw=[{meta.get('live_robot_kw','')}], automation_kw=[{meta.get('live_automation_kw','')}]"
    if live == "UNKNOWN":
        return desk_type, desk_conf, f"facility text not parsed; desk evidence kept ({kw})"
    if live == desk_type:
        return desk_type, "HIGH", f"desk evidence confirmed by current Trip.com facilities ({kw})"
    if desk_type == C.TYPE_PA and live == C.TYPE_AO:
        if desk_conf == "HIGH":
            return C.TYPE_PA, "MEDIUM", f"official robot evidence; Trip.com lists automation only ({kw})"
        return C.TYPE_AO, "MEDIUM", f"robot tag no longer listed; front-desk automation listed ({kw})"
    if desk_type == C.TYPE_PA and live == C.TYPE_CONV:
        if desk_conf == "HIGH":
            return C.TYPE_PA, "MEDIUM", f"official robot evidence; Trip.com facilities silent ({kw})"
        return C.TYPE_UNCERTAIN, "LOW", f"carried-over robot tag not found on current Trip.com page ({kw})"
    if desk_type == C.TYPE_AO and live == C.TYPE_PA:
        return C.TYPE_AO, "MEDIUM", f"Trip.com lists robot but official branch source says hologram/kiosk ({kw})"
    if desk_type == C.TYPE_AO and live == C.TYPE_CONV:
        # The public review page exposes facility *ids* plus only the top-8 facility names, and no
        # front-desk-automation label was observed there in this cycle, so Trip.com's silence is not
        # a contradiction of the official (hotel / chain) automation evidence: desk result is kept.
        return C.TYPE_AO, ("MEDIUM" if desk_conf == "HIGH" else desk_conf), (
            f"official automation evidence kept; Trip.com public facility tags carry no front-desk "
            f"automation label and no 'Service robots' tag ({kw})")
    return desk_type, desk_conf, f"unhandled combination desk={desk_type} live={live} ({kw})"

# ------------------------------------------------------------------ selection

def balance_select(elig: pd.DataFrame, n_target: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    sel = elig.copy()
    dropped = []
    step = 0
    while len(sel) > n_target:
        step += 1
        # independents are never pooled into one pseudo-brand cell
        brand_key = sel["brand"].where(sel["brand"] != "Independent/Other", "IND:" + sel["hotel_name"])
        cell_n = sel.assign(_bk=brand_key).groupby(["_bk", "city"])["candidate_id"].transform("count")
        city_n = sel.groupby("city")["candidate_id"].transform("count")
        pool = sel.assign(cell_n=cell_n, city_n=city_n)
        # Tier 1: generic "Service robots"-tag evidence that is not HIGH. Tier 2: any generic-tag
        # hotel (after the live check every generic tag is re-confirmed and therefore HIGH; without
        # this tier the balancing would fall through to hotels whose evidence is an official source
        # and drop AUTOMATION_ONLY branches instead of surplus same-brand same-city hotels).
        removable = pool[(~pool["evidence_specific"]) & (pool["final_confidence"] != "HIGH")]
        if removable.empty:
            removable = pool[~pool["evidence_specific"]]
        if removable.empty:
            removable = pool[pool["final_confidence"] != "HIGH"]
        if removable.empty:
            removable = pool
        removable = removable.sort_values(["cell_n", "city_n", "evidence_specific", "source_row_no"],
                                          ascending=[False, False, True, False])
        pick = removable.iloc[0]
        dropped.append({**pick.to_dict(), "selection_status": "RESERVE_BALANCE",
                        "selection_note": (f"step {step}: removed from {pick['brand']} × {pick['city']} "
                                           f"(cell n={int(pick['cell_n'])}, city n={int(pick['city_n'])}), "
                                           f"generic evidence='{pick['technology_detail']}'")})
        sel = sel.drop(index=pick.name)
    return sel, pd.DataFrame(dropped)


def run(offline: bool, reselect: bool, limit: int | None = None) -> pd.DataFrame:
    U.ensure_dirs()
    state = U.read_json(SELECTION_STATE, {})
    if C.FINAL_UNMANNED.exists() and state.get("mode") == "live" and not reselect:
        log.info("Live-validated selection already exists (frozen for resume safety) → reusing %s",
                 C.FINAL_UNMANNED.name)
        return pd.read_excel(C.FINAL_UNMANNED, sheet_name="unmanned_automated_100")

    df = load_candidates()
    log.info("Loaded %d unmanned/automated candidates (%d from Final_100, %d from Excluded_9)",
             len(df), (df.source_sheet == "Physical_AI_Final_100").sum(), (df.source_sheet == "Excluded_9").sum())
    try:
        cpool = pd.read_excel(U.find_input(C.UNMANNED_SOURCE_NAME), sheet_name="Control_Seed_Pool")
        control_codes = set(cpool["Trip.com_Detail"].map(U.trip_code_from_url).dropna())
    except Exception:
        control_codes = set()
    df["data_quality_flags"] = quality_flags(df, control_codes)

    run_cfg = U.load_run_config()
    metas: dict[str, dict] = {}
    if not offline:
        todo = df if limit is None else df.head(limit)
        with U.TripBrowser() as br:
            for n, (_, r) in enumerate(todo.iterrows(), 1):
                if not r["trip_code"]:
                    continue
                cached = (C.META_CACHE_DIR / f"{r['trip_code']}.json").exists()
                log.info("[%d/%d] live check %s %s", n, len(todo), r["candidate_id"], r["hotel_name"])
                metas[r["candidate_id"]] = U.fetch_hotel_meta(br, r["trip_hotel_url"], run_cfg, r["candidate_id"])
                if not cached:
                    U.polite_sleep(C.HOTEL_DELAY)
                    if n % C.BATCH_SIZE == 0:
                        U.polite_sleep(C.BATCH_COOLDOWN)

    finals = []
    for _, r in df.iterrows():
        meta = metas.get(r["candidate_id"], {})
        ftype, fconf, note = combine(r["desk_type"], r["desk_confidence"], meta)
        finals.append({
            "live_type": meta.get("live_type", ""),
            "live_robot_kw": meta.get("live_robot_kw", ""),
            "live_automation_kw": meta.get("live_automation_kw", ""),
            "live_automation_weak_kw": meta.get("live_automation_weak_kw", ""),
            "live_note": note,
            "final_type": ftype, "final_confidence": fconf,
            "star_rating": meta.get("star_rating"), "overall_rating": meta.get("overall_rating"),
            "overall_rating_scale": meta.get("overall_rating_scale"),
            "total_review_count": meta.get("total_review_count"),
            "price_value": meta.get("price_value"), "price_currency": meta.get("price_currency"),
            "price_basis": U.price_basis_text(run_cfg) if meta.get("price_value") else "",
            "trip_city_id": meta.get("trip_city_id"),
        })
    df = pd.concat([df, pd.DataFrame(finals, index=df.index)], axis=1)

    eligible = df[df["final_type"].isin([C.TYPE_PA, C.TYPE_AO])].copy()
    not_elig = df[~df.index.isin(eligible.index)].copy()
    not_elig["selection_status"] = "EXCLUDED_CLASSIFICATION"
    not_elig["selection_note"] = not_elig["live_note"]
    if len(eligible) > C.N_PER_GROUP:
        selected, reserve = balance_select(eligible, C.N_PER_GROUP)
    else:
        selected, reserve = eligible, pd.DataFrame(columns=list(eligible.columns) + ["selection_status", "selection_note"])
        if len(eligible) < C.N_PER_GROUP:
            log.warning("Only %d eligible unmanned/automated hotels (<%d). Add candidates (README §Shortfall).",
                        len(eligible), C.N_PER_GROUP)
    selected = selected.copy()
    selected["selection_status"] = "SELECTED"
    selected["selection_note"] = selected["live_note"]

    # Stable ordering → hotel_id / pair_id
    selected = selected.sort_values(["country", "city", "final_type", "brand", "hotel_name"]).reset_index(drop=True)
    selected["hotel_id"] = [f"H{i:03d}" for i in range(1, len(selected) + 1)]
    selected["pair_id"] = [f"PAIR_{i:03d}" for i in range(1, len(selected) + 1)]
    selected["group"] = C.GROUP_UNMANNED
    selected["operational_type"] = selected["final_type"]
    selected["confidence"] = selected["final_confidence"]
    selected["classification_evidence"] = selected["desk_evidence"] + " | LIVE: " + selected["live_note"]

    spec_cols = ["hotel_id", "pair_id", "group", "operational_type", "hotel_name", "country", "city",
                 "brand", "star_rating", "overall_rating", "total_review_count", "trip_hotel_url",
                 "trip_review_url", "classification_evidence", "evidence_url", "confidence",
                 "price_value", "price_currency", "price_basis"]
    extra_cols = ["brand_group", "technology_detail", "trip_code", "trip_city_id", "overall_rating_scale",
                  "desk_type", "desk_confidence", "live_type", "live_robot_kw", "live_automation_kw",
                  "data_quality_flags", "candidate_id", "source_sheet", "selection_note"]
    out = selected[spec_cols + extra_cols]

    mode = "offline" if offline else "live"
    notes = [
        ("validation_mode", mode + (" — desk evidence only; star/rating/price/review count are filled in live mode"
                                    if offline else " — desk + live Trip.com facility check")),
        ("generated_at", U.now_iso()),
        ("source", C.UNMANNED_SOURCE_NAME + " (read-only)"),
        ("candidates", str(len(df))),
        ("eligible", str(len(eligible))),
        ("selected", str(len(out))),
        ("PHYSICAL_AI", str((out.operational_type == C.TYPE_PA).sum())),
        ("AUTOMATION_ONLY", str((out.operational_type == C.TYPE_AO).sum())),
        ("stay conditions (price)", U.price_basis_text(run_cfg)),
    ] + CLASSIFICATION_RULES + SELECTION_RULES
    U.write_excel(C.FINAL_UNMANNED, {"unmanned_automated_100": out}, notes)

    all_cols = ["candidate_id", "source_sheet", "source_row_no", "country", "city", "hotel_name", "brand",
                "brand_group", "technology_detail", "evidence_specific", "prior_status", "desk_type",
                "desk_confidence", "desk_evidence", "live_type", "live_robot_kw", "live_automation_kw",
                "live_automation_weak_kw", "live_note", "final_type", "final_confidence", "star_rating",
                "overall_rating", "total_review_count", "price_value", "price_currency", "evidence_url",
                "trip_hotel_url", "trip_review_url", "trip_code", "data_quality_flags"]
    decisions = pd.concat([
        selected.assign()[all_cols + ["selection_status", "selection_note", "hotel_id"]],
        reserve.reindex(columns=all_cols + ["selection_status", "selection_note"]),
        not_elig.reindex(columns=all_cols + ["selection_status", "selection_note"]),
    ], ignore_index=True)
    brand_tab = (out.groupby("brand").size().rename("n").reset_index()
                 .assign(share=lambda t: t.n / t.n.sum()).sort_values("n", ascending=False))
    U.write_excel(C.STAGE1_CANDIDATES, {
        "all_candidate_decisions": decisions,
        "reserve_and_excluded": decisions[decisions.selection_status != "SELECTED"],
        "qa_flags": decisions[decisions.data_quality_flags.fillna("") != ""],
        "brand_concentration": brand_tab,
    }, notes)
    U.write_json_atomic(SELECTION_STATE, {"mode": mode, "created": U.now_iso(), "n_selected": len(out)})
    log.info("Selected %d (PA=%d, AO=%d); reserve=%d; excluded=%d → %s",
             len(out), (out.operational_type == C.TYPE_PA).sum(), (out.operational_type == C.TYPE_AO).sum(),
             len(reserve), len(not_elig), C.FINAL_UNMANNED.name)
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--offline", action="store_true", help="desk evidence only, no network")
    ap.add_argument("--reselect", action="store_true", help="redo selection even if a live one exists")
    ap.add_argument("--limit", type=int, help="live-check only the first N candidates (testing)")
    a = ap.parse_args()
    run(offline=a.offline, reselect=a.reselect, limit=a.limit)
