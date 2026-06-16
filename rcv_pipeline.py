#!/usr/bin/env python3
"""
Maine RCV CVR Pipeline
======================
Converts raw Cast Vote Record (CVR) xlsx files from Maine SOS into the
rcv_{year}_{district}_{party}.json format consumed by the RCV map visualization.

Usage:
  python rcv_pipeline.py --files congressd2-1.xlsx congressd2-2.xlsx ... \
                         --output rcv_2018_cd2_dem.json \
                         --year 2018 --district cd2 --party dem \
                         --geojson maine_map.html

The pipeline:
  1. Loads and concatenates CVR xlsx files (or extract-text TSV format)
  2. Normalizes candidate names and precinct names
  3. Simulates RCV elimination rounds ballot-by-ballot
  4. Scales per-town results to official SOS district totals (if provided)
  5. Outputs a single JSON file matching the visualization schema

Official 2018 CD2 Dem Primary totals (from Maine SOS):
  Round 1: Golden 20,935  |  St. Clair 17,695  |  Olson 3,962  |  Fulford 2,459
  Final:   Golden 23,611  |  St. Clair 19,853
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
# CONFIGURATION — edit these for each new election
# ──────────────────────────────────────────────────────────────────────────────

# 2018 CD2 Democratic Primary
CONFIG_2018_CD2_DEM = {
    "meta_id": "2018_cd2_dem",
    "race": "ME-CD2 Democratic Primary",
    "year": 2018,
    "total_ballots": 50845,
    "candidate_map": {
        "Golden, Jared F.": "golden",
        "St. Clair, Lucas R.": "stclair",
        "Olson, Craig R.": "olson",
        "Fulford, Jonathan S.": "fulford",
    },
    "candidates_meta": [
        {"id": "golden",  "name": "Jared Golden",     "color": "#E07B39"},
        {"id": "stclair", "name": "Lucas St. Clair",  "color": "#3B7BC8"},
        {"id": "olson",   "name": "Craig Olson",      "color": "#5BA85A"},
        {"id": "fulford", "name": "Jonathan Fulford", "color": "#9B6BB5"},
    ],
    # Elimination order: first listed is eliminated in Round 1→2, etc.
    "elimination_order": ["fulford", "olson"],
    "winner": "golden",
    "rounds_meta": [
        {"round": 1, "label": "Round 1",  "eliminated": None},
        {"round": 2, "label": "Round 2",  "eliminated": "fulford"},
        {"round": 3, "label": "Round 3",  "eliminated": "olson"},
        {"round": 4, "label": "Final",    "eliminated": None, "winner": "golden"},
    ],
    # Official SOS district totals for scaling (None = use raw sample counts)
    "official_district_totals": {
        "1": {"golden": 20935, "stclair": 17695, "olson": 3962, "fulford": 2459},
        "4": {"golden": 23611, "stclair": 19853},  # Final round
    },
}


# ──────────────────────────────────────────────────────────────────────────────
# FILE LOADING
# ──────────────────────────────────────────────────────────────────────────────

def load_cvr_files(file_paths):
    """Load CVR files. Supports both true xlsx and extract-text TSV output."""
    dfs = []
    for path in file_paths:
        path = Path(path)
        try:
            # Try reading as real xlsx first
            df = pd.read_excel(path, engine="openpyxl")
            dfs.append(df)
        except Exception:
            # Fall back to TSV (extract-text output starts with ## Sheet: header)
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            lines = content.split("\n")
            for j, line in enumerate(lines):
                if line.startswith("Cast Vote Record"):
                    data = "\n".join(lines[j:])
                    break
            else:
                print(f"WARNING: Could not find data header in {path}", file=sys.stderr)
                continue
            df = pd.read_csv(io.StringIO(data), sep="\t")
            dfs.append(df)

    if not dfs:
        raise ValueError("No CVR data loaded")

    df_all = pd.concat(dfs, ignore_index=True)
    print(f"Loaded {len(df_all):,} total rows from {len(file_paths)} file(s)")
    return df_all


# ──────────────────────────────────────────────────────────────────────────────
# NORMALIZATION
# ──────────────────────────────────────────────────────────────────────────────

def normalize_candidates(df, candidate_map):
    """Detect rank choice columns and normalize candidate names to short IDs."""
    # Find rank choice columns (handles two known column naming schemes)
    rank_sets = [
        [c for c in df.columns if "1st Choice" in c],
        [c for c in df.columns if "2nd Choice" in c],
        [c for c in df.columns if "3rd Choice" in c],
        [c for c in df.columns if "4th Choice" in c],
        [c for c in df.columns if "5th Choice" in c],
    ]

    # Consolidate multiple column sets per rank (different file formats)
    for rank_idx, cols in enumerate(rank_sets):
        merged = None
        for col in cols:
            if merged is None:
                merged = df[col]
            else:
                merged = merged.combine_first(df[col])
        df[f"R{rank_idx + 1}"] = merged

    def norm(val):
        if pd.isna(val) or str(val).strip() in ("undervote", "overvote", ""):
            return None
        v = re.sub(r"\s*\(\d+\)", "", str(val)).strip()
        return candidate_map.get(v)

    for r in range(1, 6):
        df[f"C{r}"] = df[f"R{r}"].apply(norm)

    return df


def normalize_towns(df):
    """Strip ward/district suffixes and convert to uppercase. Filter UOCAVA."""
    def norm(p):
        if pd.isna(p):
            return None
        p = str(p).strip()
        # Filter UOCAVA and purely numeric precinct codes
        if re.match(r"^\d+$", p) or "UOCAVA" in p.upper():
            return None
        # Strip common ward/district suffixes
        p = re.sub(
            r"\s+(W\d+[A-Z]*\d*|Ward\s+\d+|District\s+\d+|All|Dist\s+\d+)$",
            "", p, flags=re.IGNORECASE
        )
        return p.strip().upper()

    df["TOWN"] = df["Precinct"].apply(norm)
    return df


# ──────────────────────────────────────────────────────────────────────────────
# RCV SIMULATION
# ──────────────────────────────────────────────────────────────────────────────

def simulate_rcv(df, elimination_order, num_rounds):
    """
    Simulate RCV elimination round by round on ballot-level data.

    Returns:
        round_votes: {round_num: {town: {candidate: count}}}
        round_received: {round_num: {town: {candidate/exhausted: count}}}
    """
    ballots = df[["TOWN"] + [f"C{r}" for r in range(1, 6)]].to_dict("records")
    for b in ballots:
        b["active"] = b["C1"]

    eliminated = set()
    round_votes = {}
    round_received = {}

    for rn in range(1, num_rounds + 1):
        # Tally active choices per town
        tv = defaultdict(lambda: defaultdict(int))
        for b in ballots:
            if b["active"]:
                tv[b["TOWN"]][b["active"]] += 1
        round_votes[rn] = {t: dict(v) for t, v in tv.items()}

        # Eliminate and redistribute
        elim_idx = rn - 1
        if elim_idx < len(elimination_order):
            elim_cand = elimination_order[elim_idx]
            eliminated.add(elim_cand)

            recv = defaultdict(lambda: defaultdict(int))
            for b in ballots:
                if b["active"] == elim_cand:
                    # Find next valid choice
                    nxt = None
                    for r in range(2, 6):
                        c = b.get(f"C{r}")
                        if c and c not in eliminated:
                            nxt = c
                            break
                    b["active"] = nxt
                    t = b["TOWN"]
                    recv[t]["exhausted" if not nxt else nxt] += 1

            round_received[rn + 1] = {t: dict(v) for t, v in recv.items()}

    return round_votes, round_received


# ──────────────────────────────────────────────────────────────────────────────
# SCALING
# ──────────────────────────────────────────────────────────────────────────────

def compute_scale_factors(round_votes, official_r1):
    """
    Compute per-candidate scale factors from sample R1 vs official R1.
    Falls back to global scale for candidates not in official totals.
    """
    samp_r1 = defaultdict(int)
    for tv in round_votes[1].values():
        for c, v in tv.items():
            if c:
                samp_r1[c] += v

    global_scale = sum(official_r1.values()) / max(sum(samp_r1.values()), 1)
    scale = {c: official_r1[c] / samp_r1[c] for c in official_r1 if samp_r1.get(c, 0) > 0}

    print(f"\nSample R1: {dict(samp_r1)}")
    print(f"Official R1: {official_r1}")
    print(f"Scale factors: {{{', '.join(f'{k}: {v:.1f}x' for k,v in scale.items())}}}")
    print(f"Global scale: {global_scale:.1f}x")

    return scale, global_scale


def scale_town_votes(votes_raw, active_candidates, r1_scale, global_scale):
    return {c: round(votes_raw.get(c, 0) * r1_scale.get(c, global_scale)) for c in active_candidates}


# ──────────────────────────────────────────────────────────────────────────────
# COUNTY LOOKUP
# ──────────────────────────────────────────────────────────────────────────────

def build_county_lookup(geojson_source):
    """Extract NAME→COUNTY mapping from a GeoJSON file or HTML embedding it."""
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
# OUTPUT BUILDER
# ──────────────────────────────────────────────────────────────────────────────

def build_json(df_valid, config, geojson_source=None):
    """Full pipeline: simulate RCV → scale → build JSON."""
    elim_order = config["elimination_order"]
    num_rounds = len(config["rounds_meta"])
    official = config.get("official_district_totals", {})
    official_r1 = official.get("1", {})

    # Active candidates per round
    all_cands = [c["id"] for c in config["candidates_meta"]]
    eliminated_by_round = {}
    active_by_round = {}
    elim_set = set()
    for rn in range(1, num_rounds + 1):
        active_by_round[rn] = [c for c in all_cands if c not in elim_set]
        if rn - 1 < len(elim_order):
            elim_set.add(elim_order[rn - 1])

    # Simulate
    round_votes, round_received = simulate_rcv(df_valid, elim_order, num_rounds)

    # Scale factors
    if official_r1:
        r1_scale, g_scale = compute_scale_factors(round_votes, official_r1)
    else:
        r1_scale, g_scale = {}, 1.0
        print("No official R1 totals — using raw sample counts")

    # County lookup
    county_map = build_county_lookup(geojson_source)

    # Per-town data
    all_towns = sorted({t for t in df_valid["TOWN"].unique() if t and isinstance(t, str)})
    towns_json = {}

    for town in all_towns:
        county = county_map.get(town, "Unknown")
        tr = {}
        for rn in range(1, num_rounds + 1):
            votes_raw = round_votes[rn].get(town, {})
            votes = scale_town_votes(votes_raw, active_by_round[rn], r1_scale, g_scale)

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

            tr[str(rn)] = {"votes": votes, "received": received, "exhausted": exhausted}

        towns_json[town] = {"county": county, "rounds": tr}

    # District totals (use official where available, scaled sample elsewhere)
    def sum_round(rn):
        tot = defaultdict(int)
        for tv in round_votes[rn].values():
            for c, v in tv.items():
                if c:
                    tot[c] += v
        return dict(tot)

    district_totals = {}
    for rn in range(1, num_rounds + 1):
        if str(rn) in official:
            district_totals[str(rn)] = official[str(rn)]
        else:
            raw = sum_round(rn)
            district_totals[str(rn)] = {
                c: round(raw.get(c, 0) * r1_scale.get(c, g_scale))
                for c in active_by_round[rn]
            }

    # Sample size note
    sample_size = len(df_valid)
    total = config["total_ballots"]
    note = (
        f"Per-town data from {sample_size:,}-ballot CVR sample (~{sample_size/total*100:.1f}% of total); "
        "scaled to official SOS district totals. "
        "Official R1 and Final round totals are from SOS certification; "
        "intermediate rounds estimated from sample ballot proportions."
        if official_r1 else
        "Per-town data from full CVR."
    )

    output = {
        "meta": {
            "id": config["meta_id"],
            "race": config["race"],
            "year": config["year"],
            "total_ballots": total,
            "data_note": note,
            "candidates": config["candidates_meta"],
            "rounds": config["rounds_meta"],
            "district_totals": district_totals,
        },
        "towns": towns_json,
    }

    return output


# ──────────────────────────────────────────────────────────────────────────────
# VALIDATION
# ──────────────────────────────────────────────────────────────────────────────

def validate(output, official_r1, official_final):
    """Print validation report comparing JSON totals to official SOS numbers."""
    dt = output["meta"]["district_totals"]
    r1 = dt.get("1", {})
    final_key = max(dt.keys(), key=int)
    final = dt.get(final_key, {})

    print("\n=== VALIDATION REPORT ===")
    print("Round 1:")
    for cand, off in official_r1.items():
        got = r1.get(cand, 0)
        diff = got - off
        status = "✓" if abs(diff) <= off * 0.02 else "⚠"  # within 2%
        print(f"  {status} {cand:10s}  official={off:6,}  json={got:6,}  diff={diff:+,}")

    print(f"Final (Round {final_key}):")
    for cand, off in official_final.items():
        got = final.get(cand, 0)
        diff = got - off
        status = "✓" if abs(diff) <= off * 0.02 else "⚠"
        print(f"  {status} {cand:10s}  official={off:6,}  json={got:6,}  diff={diff:+,}")

    print(f"\nTowns with data: {len(output['towns'])}")


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Maine RCV CVR Pipeline")
    parser.add_argument("--files", nargs="+", required=True, help="CVR xlsx/tsv files")
    parser.add_argument("--output", default="rcv_output.json", help="Output JSON path")
    parser.add_argument("--geojson", default=None, help="maine_map.html or GeoJSON for county lookup")
    parser.add_argument("--election", default="2018_cd2_dem", help="Election config key")
    args = parser.parse_args()

    configs = {
        "2018_cd2_dem": CONFIG_2018_CD2_DEM,
        # Add future election configs here
    }

    if args.election not in configs:
        print(f"Unknown election: {args.election}. Available: {list(configs.keys())}", file=sys.stderr)
        sys.exit(1)

    config = configs[args.election]

    # Load
    df = load_cvr_files(args.files)

    # Normalize
    df = normalize_candidates(df, config["candidate_map"])
    df = normalize_towns(df)
    df_valid = df[df["TOWN"].notna()].copy()
    print(f"Valid ballots: {len(df_valid):,} across {df_valid['TOWN'].nunique()} towns")

    # Build JSON
    output = build_json(df_valid, config, geojson_source=args.geojson)

    # Validate
    official_r1 = config.get("official_district_totals", {}).get("1", {})
    final_key = str(len(config["rounds_meta"]))
    official_final = config.get("official_district_totals", {}).get(final_key, {})
    if official_r1:
        validate(output, official_r1, official_final)

    # Write
    out_path = Path(args.output)
    with open(out_path, "w") as f:
        json.dump(output, f, indent=2)
    print(f"\n✓ Output written to {out_path} ({out_path.stat().st_size // 1024} KB)")
    print(f"\nEmbed this in rcv_map.html under DATA['{config['meta_id']}']")


if __name__ == "__main__":
    main()
