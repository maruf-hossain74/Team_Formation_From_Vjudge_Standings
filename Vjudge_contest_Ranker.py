#!/usr/bin/env python3
"""
VJudge Team Maker — Local Excel Mode (improved & robust)
Requirements: pandas, openpyxl
You may put those in requirements.txt or install manually:
    pip install pandas openpyxl
"""

import os
import sys
import math
import subprocess
import pandas as pd
from collections import defaultdict

# ---------- Config ----------
REQUIREMENTS_FILE = "requirements.txt"   # keep lowercase, consistent
LEADERBOARDS_DIR = "Leaderboards"
OUT_FILE = "final_teams.xlsx"

# Points formula config — change to match your preferred formula
# Example: return math.ceil(1800 / (rank + 5))
POINTS_NUMERATOR = 1600
POINTS_OFFSET = 7

# Possible column names for username/handle and rank (case-insensitive matching)
USERNAME_COLUMNS = ["Username", "username", "Team", "team", "Handle", "handle"]
RANK_COLUMNS = ["Rank", "rank", "POSITION", "Position", "position"]


# ---------- Helpers ----------
def points_for_rank(rank: int) -> int:
    """Compute points based on rank using configurable numerator/offset."""
    if rank is None or rank < 1:
        return 0
    return math.ceil(POINTS_NUMERATOR / (rank + POINTS_OFFSET))


def safe_install_requirements(req_file: str):
    """Install dependencies listed in req_file (line by line). Non-fatal on errors."""
    if not os.path.exists(req_file):
        print(f"[info] requirements file '{req_file}' not found — skipping install.")
        return

    with open(req_file, "r", encoding="utf-8") as f:
        deps = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]

    if not deps:
        print(f"[info] No dependencies found in {req_file}.")
        return

    print(f"[info] Installing {len(deps)} dependencies from {req_file}...")
    for dep in deps:
        print(f"  → Installing: {dep}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
        except subprocess.CalledProcessError as e:
            print(f"  [warn] Failed to install {dep}: {e}. Continuing...")

    print("[info] Dependency installation attempt finished.")


def find_column_name(df: pd.DataFrame, candidates):
    """Return first matching column name in df columns from candidates, else None."""
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None


def read_excel_file(file_path: str):
    """Read an Excel file and return a list of (username, rank) tuples."""
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"❌ Error reading file {file_path}: {e}")
        return []

    user_col = find_column_name(df, USERNAME_COLUMNS)
    rank_col = find_column_name(df, RANK_COLUMNS)

    if user_col is None or rank_col is None:
        print(f"[warn] Couldn't find Username or Rank columns in {file_path}. "
              f"Available columns: {list(df.columns)}. Skipping file.")
        return []

    standings = []
    for _, row in df.iterrows():
        raw_user = row.get(user_col)
        raw_rank = row.get(rank_col)

        if pd.isna(raw_user):
            continue
        username = str(raw_user).strip()
        if not username:
            continue

        # Parse rank robustly
        rank = None
        if pd.isna(raw_rank):
            # leave rank as None (skip or treat as 0)
            pass
        else:
            try:
                # handle floats like 1.0
                rank = int(float(raw_rank))
                if rank < 1:
                    rank = None
            except Exception:
                rank = None

        if rank is None:
            # Skip rows without valid rank (or choose to set a default)
            continue

        standings.append((username, rank))

    return standings


def write_participants_and_teams_to_excel(all_scores: dict, team_size: int, out_file: str = OUT_FILE):
    """Write participants with their scores for each file and calculate the FinalPoints."""
    # Collect all usernames
    usernames = set()
    for file_scores in all_scores.values():
        usernames.update(file_scores.keys())

    # Build participant rows
    participants_data = []
    file_names_order = list(all_scores.keys())
    for username in usernames:
        row = {"Username": username}
        for file in file_names_order:
            row[file] = all_scores[file].get(username, 0)
        participants_data.append(row)

    if not participants_data:
        print("[warn] No participant data to write.")
        return

    participants_df = pd.DataFrame(participants_data)
    # Sum FinalPoints
    score_cols = [c for c in participants_df.columns if c != "Username"]
    participants_df["FinalPoints"] = participants_df[score_cols].sum(axis=1)
    # Sort descending by FinalPoints, tiebreak by Username for determinism
    participants_df = participants_df.sort_values(by=["FinalPoints", "Username"], ascending=[False, True]).reset_index(drop=True)

    # Write to Excel with teams
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        participants_df.to_excel(writer, sheet_name="Participants", index=False)
        print(f"[ok] Saved Participants sheet to '{out_file}'")

        # Build teams sequentially
        teams = [participants_df.iloc[i:i + team_size].reset_index(drop=True) for i in range(0, len(participants_df), team_size)]
        for i, team_df in enumerate(teams, start=1):
            sheet_name = f"Team_{i}"
            team_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"[ok] Saved {sheet_name} to '{out_file}'")


def main():
    # Optionally install requirements
    safe_install_requirements(REQUIREMENTS_FILE)

    # Check leaderboards directory
    if not os.path.isdir(LEADERBOARDS_DIR):
        print(f"❌ Leaderboards directory '{LEADERBOARDS_DIR}' not found. Please create it and put contest .xlsx files inside.")
        return

    # Accept both .xlsx and .xls
    files = [f for f in os.listdir(LEADERBOARDS_DIR) if f.lower().endswith((".xlsx", ".xls"))]
    if not files:
        print(f"❌ No Excel files (.xlsx/.xls) found in '{LEADERBOARDS_DIR}' directory. Exiting.")
        return

    # Ask for team size (robust)
    try:
        print()
        team_size_input = input("👥 Enter team size (default 3): ").strip()
        team_size = int(team_size_input) if team_size_input else 3
        if team_size <= 0:
            raise ValueError
    except Exception:
        print("[warn] Invalid team size provided. Defaulting to 3.")
        team_size = 3

    all_scores = defaultdict(dict)
    any_scores = False

    for fname in files:
        path = os.path.join(LEADERBOARDS_DIR, fname)
        print(f"\n📡 Processing: {path}")
        standings = read_excel_file(path)
        if not standings:
            print(f"[warn] No valid standings in {fname}. Skipping.")
            continue

        any_scores = True
        for username, rank in standings:
            pts = points_for_rank(rank)
            # If a user appears multiple times within a file (unlikely), keep best / or sum — here we keep the maximum for that file
            prev = all_scores[fname].get(username, 0)
            all_scores[fname][username] = max(prev, pts)

        print(f"[info] Processed {len(standings)} rows from {fname}.")

    if not any_scores:
        print("❌ No points computed from any files. Exiting.")
        return

    write_participants_and_teams_to_excel(all_scores, team_size, out_file=OUT_FILE)
    print("\n✅ Done. Output saved to:", OUT_FILE)


if __name__ == "__main__":
    print("🏆 ICPC/IUPC Team Formation — Local Excel Mode (improved)")
    print("Make sure the 'Leaderboards' folder contains the contest .xlsx/.xls files.")
    main()
