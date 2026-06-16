#!/usr/bin/env python3
"""
Maine RCV CVR Pipeline — v2
============================
Converts raw Maine SOS Cast Vote Record (CVR) xlsx files into the
rcv_{year}_{district}_{party}.json format consumed by the RCV map.

New in v2:
  - round_transitions replaces elimination_order: each transition is a LIST,
    supporting simultaneous multi-candidate elimination in a single round step.
    2018 CD2 reality: Fulford + Olson were eliminated simultaneously after
    Round 1 — Maine SOS certified exactly 2 rounds. The pipeline now reflects
    that accurately rather than manufacturing fictional intermediate rounds.
  - rounds_meta is derived automatically from round_transitions
  - UOCAVA_UNMAPPED synthetic town entry
  - Enhanced town round schema: leader + continuing fields
  - story metadata stubs (popups pre-populated from data)
  - --manifest flag: upserts election entry in data/manifest.json

Usage:
  python rcv_pipeline_v2.py \\
    --files congressd21.xlsx congressd22.xlsx congressd23.xlsx congressd24.xlsx \\
    --output data/rcv_2018_cd2_dem.json \\
    --geojson maine_map.html \\
    --election 2018_cd2_dem \\
    --manifest data/manifest.json

Official 2018 CD2 Dem Primary totals (Maine SOS RCV Report, certified 2018-06-22):
  Round 1: Golden 20,987 | St. Clair 17,742 | Olson 3,993 | Fulford 2,489
  Final:   Golden 23,611 | St. Clair 19,853
  Elimination: Fulford + Olson simultaneously after Round 1 → 2-round race
"""

import argparse
import io
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd


# ──────────────────────────────────────────────────────────────────────────────
# ELECTION CONFIGURATION
# ──────────────────────────────────────────────────────────────────────────────

CONFIG_2018_CD2_DEM = {
    "meta_id": "2018_cd2_dem",
    "race": "ME-CD2 Democratic Primary",
    "year": 2018,
    "total_ballots": 50845,
    "candidate_map": {
        # congressd21.xlsx format: "Last, First M. (NNNNN)"
        "Golden, Jared F.": "golden",
        "St. Clair, Lucas R.": "stclair",
        "Olson, Craig R.": "olson",
        "Fulford, Jonathan S.": "fulford",
        # congressd22–24.xlsx format: "Last, First M." (no numeric suffix)
        "Golden, Jared F": "golden",
        "St. Clair, Lucas R": "stclair",
        "Olson, Craig R": "olson",
        "Fulford, Jonathan S": "fulford",
    },
    "candidates_meta": [
        {"id": "golden",  "name": "Jared Golden",     "color": "#E07B39"},
        {"id": "stclair", "name": "Lucas St. Clair",  "color": "#3B7BC8"},
        {"id": "olson",   "name": "Craig Olson",      "color": "#5BA85A"},
        {"id": "fulford", "name": "Jonathan Fulford", "color": "#9B6BB5"},
    ],
    # round_transitions[i] = list of candidates eliminated simultaneously
    # AFTER round (i+1), producing round (i+2).
    # 2018: Fulford AND Olson were eliminated simultaneously after Round 1.
    # Maine SOS certified exactly 2 rounds — no intermediate rounds exist.
    "round_transitions": [
        ["fulford", "olson"],   # After Round 1 → Round 2 (Final)
    ],
    "winner": "golden",
    # official_district_totals: keyed by true round number (string).
    # Round 1 = four-candidate tally. Round 2 = the certified Final.
    "official_district_totals": {
        "1": {"golden": 20987, "stclair": 17742, "olson": 3993, "fulford": 2489},
        "2": {"golden": 23611, "stclair": 19853},
    },
    # Manifest display fields
    "manifest_label": "2018 CD2 Dem Primary",
    "manifest_meta": "4 candidates · 2 rounds · 50,845 ballots",
    "manifest_candidates": ["Golden", "St. Clair", "Olson", "Fulford"],
    "manifest_data_url": "data/rcv_2018_cd2_dem.json",
}


# ──────────────────────────────────────────────────────────────────────────────
# ROUNDS META DERIVATION
# ──────────────────────────────────────────────────────────────────────────────

def build_rounds_meta(config):
    """
    Derive rounds_meta from round_transitions.

    round_transitions[i] = candidates eliminated after round (i+1).
    Each entry can be a single candidate id string OR a list of ids
    for simultaneous multi-candidate elimination.

    Returns a list of dicts matching the visualization schema:
      {"round": N, "label": "...", "eliminated": [...] | None, "winner": id | None}
    """
    transitions = config["round_transitions"]
    num_rounds = len(transitions) + 1
    winner = config.get("winner")

    rounds = []
    for rn in range(1, num_rounds + 1):
        # eliminated = what gets eliminated AFTER this round
        trans_idx = rn - 1
        if trans_idx < len(transitions):
            elim = transitions[trans_idx]
            # Normalize to list always
            if isinstance(elim, str):
                elim = [elim]
        else:
            elim = None

        label = "Final" if rn == num_rounds else f"Round {rn}"

        entry = {
            "round": rn,
            "label": label,
            "eliminated": elim,  # list or None
        }
        if rn == num_rounds and winner:
            entry["winner"] = winner

        rounds.append(entry)

    return rounds


# ──────────────────────────────────────────────────────────────────────────────
# FILE LOADING
# ──────────────────────────────────────────────────────────────────────────────

def load_cvr_files(file_paths):
    """Load and concatenate CVR xlsx files. Handles two column-naming schemes."""
    dfs = []
    for path in file_paths:
        path = Path(path)
        try:
            df = pd.read_excel(path, engine="openpyxl")
            dfs.append(df)
            print(f"  {path.name}: {len(df):,} rows")
        except Exception:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                lines = content.split("\n")
                for j, line in enumerate(lines):
                    if line.startswith("Cast Vote Record"):
                        data_text = "\n".join(lines[j:])
                        break
                else:
                    print(f"WARNING: No data header in {path}", file=sys.stderr)
                    continue
                df = pd.read_csv(io.StringIO(data_text), sep="\t")
                dfs.append(df)
                print(f"  {path.name} (TSV fallback): {len(df):,} rows")
            except Exception as e2:
                print(f"WARNING: Failed to load {path}: {e2}", file=sys.stderr)

    if not dfs:
        raise ValueError("No CVR data loaded from any file")

    df_all = pd.concat(dfs, ignore_index=True)
    print(f"  Total: {len(df_all):,} rows from {len(file_paths)} file(s)")
    return df_all


# ──────────────────────────────────────────────────────────────────────────────
# NORMALIZATION
# ──────────────────────────────────────────────────────────────────────────────

def normalize_candidates(df, candidate_map):
    """
    Detect rank-choice columns (handles both observed naming schemes),
    consolidate into R1–R5, map to short candidate IDs as C1–C5.
    """
    rank_patterns = [
        ["1st Choice", "1st choice"],
        ["2nd Choice", "2nd choice"],
        ["3rd Choice", "3rd choice"],
        ["4th Choice", "4th choice"],
        ["5th Choice", "5th choice"],
    ]

    for rank_idx, patterns in enumerate(rank_patterns):
        merged = None
        for col in df.columns:
            if any(p in col for p in patterns):
                merged = df[col] if merged is None else merged.combine_first(df[col])
        df[f"R{rank_idx + 1}"] = merged if merged is not None else pd.Series([None] * len(df))

    def norm(val):
        if pd.isna(val) or str(val).strip().lower() in ("undervote", "overvote", ""):
            return None
        v = re.sub(r"\s*\(\d+\)\s*$", "", str(val)).strip()
        result = candidate_map.get(v)
        if result:
            return result
        # Try without trailing period on middle initial
        return candidate_map.get(re.sub(r"\.$", "", v))

    for r in range(1, 6):
        df[f"C{r}"] = df[f"R{r}"].apply(norm)

    return df


def normalize_towns(df):
    """
    Strip ward/district suffixes → uppercase.
    Tag UOCAVA rows and purely numeric precinct codes with '_UOCAVA_' sentinel.
    """
    def norm(p):
        if pd.isna(p):
            return None
        p = str(p).strip()
        if re.match(r"^\d+$", p) or "UOCAVA" in p.upper():
            return "_UOCAVA_"
        p = re.sub(
            r"\s+(W\d+[A-Z]*\d*|Ward\s+\d+|District\s+\d+|All|Dist\s+\d+)$",
            "", p, flags=re.IGNORECASE,
        )
        return p.strip().upper()

    df["TOWN"] = df["Precinct"].apply(norm)
    return df


# ──────────────────────────────────────────────────────────────────────────────
# RCV SIMULATION  (supports simultaneous multi-candidate elimination)
# ──────────────────────────────────────────────────────────────────────────────

def simulate_rcv(df, round_transitions, num_rounds):
    """
    Ballot-level RCV simulation supporting simultaneous elimination of
    multiple candidates in a single round transition.

    round_transitions[i] = list of candidate ids eliminated after round (i+1).
    All candidates in a transition group are eliminated AT THE SAME TIME:
    each ballot from any of them looks for the next choice not in the
    full eliminated set (including all newly eliminated candidates).

    Returns:
        round_votes:    {rn: {town: {candidate: count}}}
        round_received: {rn: {town: {candidate|'exhausted': count}}}
                        round_received[rn] = transfers that PRODUCED round rn's totals
    """
    ballots = df[["TOWN"] + [f"C{r}" for r in range(1, 6)]].to_dict("records")
    for b in ballots:
        b["active"] = b["C1"]

    cumulative_eliminated = set()
    round_votes = {}
    round_received = {}

    for rn in range(1, num_rounds + 1):
        # Tally active choices per town
        tv = defaultdict(lambda: defaultdict(int))
        for b in ballots:
            if b["active"]:
                tv[b["TOWN"]][b["active"]] += 1
        round_votes[rn] = {t: dict(v) for t, v in tv.items()}

        # Execute transition after this round if one exists
        trans_idx = rn - 1
        if trans_idx < len(round_transitions):
            newly_eliminated = set(round_transitions[trans_idx])
            # Build the full eliminated set BEFORE redistributing
            # so that transfers skip ALL eliminated candidates (including
            # those eliminated in earlier transitions AND these new ones).
            all_eliminated_now = cumulative_eliminated | newly_eliminated

            recv = defaultdict(lambda: defaultdict(int))
            for b in ballots:
                if b["active"] in newly_eliminated:
                    # Find next valid choice skipping everything eliminated so far
                    nxt = None
                    for r in range(2, 6):
                        c = b.get(f"C{r}")
                        if c and c not in all_eliminated_now:
                            nxt = c
                            break
                    b["active"] = nxt
                    t = b["TOWN"]
                    recv[t]["exhausted" if not nxt else nxt] += 1

            # round_received[rn+1] = the transfers that produce round rn+1's totals
            round_received[rn + 1] = {t: dict(v) for t, v in recv.items()}
            cumulative_eliminated = all_eliminated_now

    return round_votes, round_received


# ──────────────────────────────────────────────────────────────────────────────
# SCALING
# ──────────────────────────────────────────────────────────────────────────────

def compute_scale_factors(round_votes, official_r1):
    """Per-candidate scale from R1 sample (mapped towns only) → official totals."""
    samp_r1 = defaultdict(int)
    for town, tv in round_votes[1].items():
        if town == "_UOCAVA_":
            continue
        for c, v in tv.items():
            if c:
                samp_r1[c] += v

    scale = {c: official_r1[c] / samp_r1[c] for c in official_r1 if samp_r1.get(c, 0) > 0}
    global_scale = sum(official_r1.values()) / max(sum(samp_r1.values()), 1)

    print(f"  Sample R1 (mapped towns): {dict(samp_r1)}")
    print(f"  Official R1:              {official_r1}")
    print(f"  Scale factors: { {k: f'{v:.4f}x' for k,v in scale.items()} }")
    print(f"  Global scale:  {global_scale:.4f}x")
    return scale, global_scale


def scale_town_votes(votes_raw, active_candidates, r1_scale, global_scale):
    return {c: round(votes_raw.get(c, 0) * r1_scale.get(c, global_scale)) for c in active_candidates}


# ──────────────────────────────────────────────────────────────────────────────
# COUNTY LOOKUP
# ──────────────────────────────────────────────────────────────────────────────

def build_county_lookup(geojson_source):
    if not geojson_source:
        return {}
    path = Path(geojson_source)
    if not path.exists():
        return {}
    content = path.read_text(encoding="utf-8")
    names = re.findall(r'"NAME"\s*:\s*"([^"]+)"', content)
    counties = re.findall(r'"COUNTY"\s*:\s*"([^"]+)"', content)
    return {n.upper(): c for n, c in zip(names, counties)}


# ──────────────────────────────────────────────────────────────────────────────
# STORY METADATA GENERATOR
# ──────────────────────────────────────────────────────────────────────────────

def generate_story_metadata(config, rounds_meta, round_votes, round_received, r1_scale, global_scale):
    """
    Generate story stubs. headline/before/after = empty strings (app fills at runtime).
    popups are pre-populated with data-driven entries where informative.
    Handles simultaneous multi-candidate elimination correctly.
    """
    cand_names = {c["id"]: c["name"] for c in config["candidates_meta"]}
    official = config.get("official_district_totals", {})

    rounds_story = {}
    for rm in rounds_meta:
        rn = rm["round"]
        eliminated = rm.get("eliminated")  # list or None
        winner = rm.get("winner")
        popups = []

        if eliminated:
            # Aggregate district-level transfers that produced round rn+1
            recv_all = round_received.get(rn + 1, {})
            dist_recv = defaultdict(int)
            for tv in recv_all.values():
                for k, v in tv.items():
                    dist_recv[k] += v

            # Total ballots from all eliminated candidates (use official R1 where available)
            off_r1 = official.get("1", {})
            total_elim_ballots = sum(off_r1.get(e, 0) for e in eliminated)
            if not total_elim_ballots:
                # Fallback: scale sample
                total_elim_ballots = round(sum(
                    sum(round_votes[rn].get(t, {}).get(e, 0) for t in round_votes[rn] if t != "_UOCAVA_")
                    for e in eliminated
                ) * global_scale)

            top_recv = max(
                ((c, v) for c, v in dist_recv.items() if c != "exhausted"),
                key=lambda x: x[1], default=(None, 0)
            )
            exhausted_raw = dist_recv.get("exhausted", 0)
            exhausted_n = round(exhausted_raw * global_scale)
            exhausted_pct = (exhausted_n / total_elim_ballots * 100) if total_elim_ballots else 0

            elim_label = (
                cand_names.get(eliminated[0], eliminated[0])
                if len(eliminated) == 1
                else " & ".join(cand_names.get(e, e) for e in eliminated)
            )

            if top_recv[0]:
                popups.append({
                    "type": "TRANSFER WATCH",
                    "text": (
                        f"{cand_names.get(top_recv[0], top_recv[0])} absorbs the most "
                        f"of {elim_label}'s {total_elim_ballots:,} combined ballots district-wide."
                    )
                })

            if exhausted_pct > 5:
                popups.append({
                    "type": "BALLOTS EXHAUSTED",
                    "text": (
                        f"{exhausted_n:,} ballots ({exhausted_pct:.1f}%) exhaust "
                        f"after {elim_label}'s elimination — no next ranked continuing candidate."
                    )
                })

        if winner:
            popups.append({
                "type": "COUNTING NOTE",
                "text": (
                    f"{cand_names.get(winner, winner)} crosses the majority threshold "
                    f"of continuing ballots to win the district."
                )
            })

        rounds_story[str(rn)] = {
            "headline": "",
            "before": "",
            "after": "",
            "popups": popups,
        }

    return {"intro": "", "rounds": rounds_story}


# ──────────────────────────────────────────────────────────────────────────────
# UOCAVA SYNTHETIC TOWN
# ──────────────────────────────────────────────────────────────────────────────

def build_uocava_entry(round_votes, round_received, active_by_round, num_rounds):
    """Aggregate _UOCAVA_ ballots into a synthetic UOCAVA_UNMAPPED entry (unscaled)."""
    rounds_data = {}
    for rn in range(1, num_rounds + 1):
        votes_raw = round_votes[rn].get("_UOCAVA_", {})
        votes = {c: votes_raw.get(c, 0) for c in active_by_round[rn]}
        continuing = sum(votes.values())
        leader = max(votes, key=votes.get) if any(v > 0 for v in votes.values()) else None

        if rn == 1:
            received = None
            exhausted = 0
        else:
            recv_raw = round_received.get(rn, {}).get("_UOCAVA_", {})
            exhausted = recv_raw.get("exhausted", 0)
            recv_cands = {k: v for k, v in recv_raw.items() if k != "exhausted" and v > 0}
            received = recv_cands if recv_cands else None

        rounds_data[str(rn)] = {
            "votes": votes,
            "leader": leader,
            "continuing": continuing,
            "received": received,
            "exhausted": exhausted,
        }

    return {
        "type": "special",
        "mapped": False,
        "label": "UOCAVA / Unmapped Ballots",
        "county": None,
        "description": (
            "Ballots included in race totals but not assigned to a Maine municipality. "
            "Includes overseas and military ballots (UOCAVA) and any precincts "
            "with only numeric identifiers in the CVR."
        ),
        "rounds": rounds_data,
    }


# ──────────────────────────────────────────────────────────────────────────────
# OUTPUT BUILDER
# ──────────────────────────────────────────────────────────────────────────────

def build_json(df_valid, config, geojson_source=None):
    """Full pipeline: derive rounds → simulate → scale → build JSON."""

    rounds_meta = build_rounds_meta(config)
    num_rounds = len(rounds_meta)
    transitions = config["round_transitions"]
    official = config.get("official_district_totals", {})
    official_r1 = official.get("1", {})

    # Active candidates per round
    all_cands = [c["id"] for c in config["candidates_meta"]]
    active_by_round = {}
    elim_set = set()
    for rn in range(1, num_rounds + 1):
        active_by_round[rn] = [c for c in all_cands if c not in elim_set]
        if rn - 1 < len(transitions):
            for e in transitions[rn - 1]:
                elim_set.add(e)

    # Simulate (all candidates including _UOCAVA_)
    print("\n=== Simulating RCV ===")
    round_votes, round_received = simulate_rcv(df_valid, transitions, num_rounds)
    print(f"  {num_rounds} rounds, {len(transitions)} transition(s)")
    for i, t in enumerate(transitions):
        names = [c["name"] for c in config["candidates_meta"] if c["id"] in t]
        print(f"  After Round {i+1}: eliminate {' + '.join(names)} simultaneously")

    # Scale factors
    print("\n=== Scaling ===")
    if official_r1:
        r1_scale, g_scale = compute_scale_factors(round_votes, official_r1)
    else:
        r1_scale, g_scale = {}, 1.0
        print("  No official R1 totals — using raw sample counts")

    # County lookup
    county_map = build_county_lookup(geojson_source)

    # Per-town data
    all_towns = sorted({
        t for t in df_valid["TOWN"].unique()
        if t and isinstance(t, str) and t != "_UOCAVA_"
    })
    print(f"\n=== Building town data: {len(all_towns)} towns ===")

    towns_json = {}
    for town in all_towns:
        county = county_map.get(town, "Unknown")
        tr = {}
        for rn in range(1, num_rounds + 1):
            votes_raw = round_votes[rn].get(town, {})
            votes = scale_town_votes(votes_raw, active_by_round[rn], r1_scale, g_scale)
            continuing = sum(votes.values())
            leader = max(votes, key=votes.get) if any(v > 0 for v in votes.values()) else None

            if rn == 1:
                received = None
                exhausted = 0
            else:
                recv_raw = round_received.get(rn, {}).get(town, {})
                exhausted = round(recv_raw.get("exhausted", 0) * g_scale)
                recv_cands = {
                    k: round(v * g_scale)
                    for k, v in recv_raw.items()
                    if k != "exhausted" and isinstance(k, str) and v > 0
                }
                received = recv_cands if recv_cands else None

            tr[str(rn)] = {
                "votes": votes,
                "leader": leader,
                "continuing": continuing,
                "received": received,
                "exhausted": exhausted,
            }

        towns_json[town] = {"county": county, "rounds": tr}

    # UOCAVA synthetic entry
    uocava_raw = (df_valid["TOWN"] == "_UOCAVA_").sum()
    if uocava_raw > 0:
        print(f"  Adding UOCAVA_UNMAPPED ({uocava_raw:,} raw ballots, unscaled)")
        towns_json["UOCAVA_UNMAPPED"] = build_uocava_entry(
            round_votes, round_received, active_by_round, num_rounds
        )

    # District totals (official where available; scaled sample otherwise)
    def sum_round_scaled(rn):
        tot = defaultdict(int)
        for town, tv in round_votes[rn].items():
            if town == "_UOCAVA_":
                continue
            for c, v in tv.items():
                if c:
                    tot[c] += v
        return {c: round(tot.get(c, 0) * r1_scale.get(c, g_scale)) for c in active_by_round[rn]}

    district_totals = {}
    for rn in range(1, num_rounds + 1):
        if str(rn) in official:
            district_totals[str(rn)] = official[str(rn)]
        else:
            district_totals[str(rn)] = sum_round_scaled(rn)

    # Story metadata
    story = generate_story_metadata(
        config, rounds_meta, round_votes, round_received, r1_scale, g_scale
    )

    sample_size = len(df_valid[df_valid["TOWN"] != "_UOCAVA_"])
    total = config["total_ballots"]
    note = (
        f"Per-town data from {sample_size:,}-ballot CVR sample "
        f"(~{sample_size / total * 100:.1f}% of {total:,} total ballots); "
        "scaled to official Maine SOS district totals. "
        "Official R1 and Final round totals from SOS RCV Report (certified 2018-06-22); "
        "intermediate rounds estimated from sample ballot proportions."
        if official_r1 else
        "Per-town data from full CVR."
    )

    return {
        "meta": {
            "id": config["meta_id"],
            "race": config["race"],
            "year": config["year"],
            "total_ballots": total,
            "data_note": note,
            "candidates": config["candidates_meta"],
            "rounds": rounds_meta,
            "district_totals": district_totals,
        },
        "story": story,
        "towns": towns_json,
    }


# ──────────────────────────────────────────────────────────────────────────────
# VALIDATION
# ──────────────────────────────────────────────────────────────────────────────

def validate(output, official_r1, official_final):
    dt = output["meta"]["district_totals"]
    r1 = dt.get("1", {})
    final_key = max(dt.keys(), key=int)
    final = dt.get(final_key, {})

    print("\n=== VALIDATION ===")
    print("Round 1 vs. SOS certified:")
    for cand, off in official_r1.items():
        got = r1.get(cand, 0)
        diff = got - off
        status = "✓" if abs(diff) <= max(off * 0.02, 2) else "⚠"
        print(f"  {status} {cand:10s}  official={off:6,}  json={got:6,}  diff={diff:+,}")

    print(f"Final (Round {final_key}) vs. SOS certified:")
    for cand, off in official_final.items():
        got = final.get(cand, 0)
        diff = got - off
        status = "✓" if abs(diff) <= max(off * 0.02, 2) else "⚠"
        print(f"  {status} {cand:10s}  official={off:6,}  json={got:6,}  diff={diff:+,}")

    mapped = sum(1 for k, v in output["towns"].items() if v.get("type") != "special")
    print(f"\nMapped towns: {mapped}")
    uocava = output["towns"].get("UOCAVA_UNMAPPED")
    if uocava:
        r1v = uocava["rounds"].get("1", {}).get("votes", {})
        print(f"UOCAVA_UNMAPPED R1: { {c: v for c, v in r1v.items() if v > 0} }")

    # Print true round structure
    print("\nTrue round structure:")
    for rm in output["meta"]["rounds"]:
        elim = rm.get("eliminated")
        elim_str = " + ".join(elim) if elim else "—"
        win_str = f"  → winner: {rm.get('winner')}" if rm.get("winner") else ""
        print(f"  Round {rm['round']} ({rm['label']}): eliminated after = [{elim_str}]{win_str}")


# ──────────────────────────────────────────────────────────────────────────────
# MANIFEST
# ──────────────────────────────────────────────────────────────────────────────

def update_manifest(manifest_path, config):
    """Upsert election entry in manifest.json (creates file if absent)."""
    path = Path(manifest_path)
    manifest = []
    if path.exists():
        with open(path) as f:
            manifest = json.load(f)

    entry = {
        "id": config["meta_id"],
        "label": config["manifest_label"],
        "year": config["year"],
        "race": config["race"],
        "status": "available",
        "dataUrl": config["manifest_data_url"],
        "meta": config["manifest_meta"],
        "candidates": config["manifest_candidates"],
    }

    replaced = False
    for i, item in enumerate(manifest):
        if item.get("id") == config["meta_id"]:
            manifest[i] = entry
            replaced = True
            break
    if not replaced:
        manifest.append(entry)

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w") as f:
        json.dump(manifest, f, indent=2)
    action = "Updated" if replaced else "Added"
    print(f"  {action} '{config['meta_id']}' in {path} ({len(manifest)} election(s))")


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Maine RCV CVR Pipeline v2")
    parser.add_argument("--files", nargs="+", required=True, help="CVR xlsx/tsv files")
    parser.add_argument("--output", default="rcv_output.json", help="Output JSON path")
    parser.add_argument("--geojson", default=None, help="maine_map.html or GeoJSON for county lookup")
    parser.add_argument("--election", default="2018_cd2_dem", help="Election config key")
    parser.add_argument("--manifest", default=None, help="Path to manifest.json to upsert")
    args = parser.parse_args()

    configs = {
        "2018_cd2_dem": CONFIG_2018_CD2_DEM,
        # Add future elections here
    }

    if args.election not in configs:
        print(f"Unknown election: {args.election}. Available: {list(configs.keys())}", file=sys.stderr)
        sys.exit(1)

    config = configs[args.election]

    # Load
    print("=== Loading CVR files ===")
    df = load_cvr_files(args.files)

    # Normalize
    print("\n=== Normalizing ===")
    df = normalize_candidates(df, config["candidate_map"])
    df = normalize_towns(df)
    df_valid = df[df["TOWN"].notna()].copy()

    uocava_count = (df_valid["TOWN"] == "_UOCAVA_").sum()
    mapped_count = (df_valid["TOWN"] != "_UOCAVA_").sum()
    print(f"  Mapped ballots:          {mapped_count:,} across {df_valid[df_valid['TOWN'] != '_UOCAVA_']['TOWN'].nunique()} towns")
    print(f"  UOCAVA/unmapped ballots: {uocava_count:,}")

    # Build
    output = build_json(df_valid, config, geojson_source=args.geojson)

    # Validate
    official_r1 = config.get("official_district_totals", {}).get("1", {})
    final_key = str(len(build_rounds_meta(config)))
    official_final = config.get("official_district_totals", {}).get(final_key, {})
    if official_r1:
        validate(output, official_r1, official_final)

    # Write
    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w") as f:
        json.dump(output, f, indent=2)
    print(f"\n✓ {out_path} ({out_path.stat().st_size // 1024} KB)")

    # Manifest
    if args.manifest:
        print("\n=== Updating manifest ===")
        update_manifest(args.manifest, config)


if __name__ == "__main__":
    main()
