import os
import math
import pandas as pd
from collections import defaultdict

# ---------- Points formula ----------
def points_for_rank(rank: int) -> int:
    """Compute points based on rank."""
    return math.ceil(1800 / (rank + 5))

def read_excel_file(file_path: str):
    """Read an Excel file and return a list of (username, rank) tuples."""
    try:
        df = pd.read_excel(file_path)
        # Assuming columns are named "Username" and "Rank", adjust if necessary
        standings = []
        for _, row in df.iterrows():
            username = str(row['Username']).strip()
            try:
                rank = int(row['Rank'])
                standings.append((username, rank))
            except ValueError:
                continue  # If rank is not a valid integer, skip this row
        return standings
    except Exception as e:
        print(f"❌ Error reading file {file_path}: {e}")
        return []

def write_participants_and_teams_to_excel(all_scores: dict, team_size: int, out_file: str = "final_teams.xlsx"):
    """Write participants with their scores for each file and calculate the FinalPoints."""
    
    # Create a DataFrame with Username and points for each file
    participants_data = []
    usernames = set()
    for file, scores in all_scores.items():
        usernames.update(scores.keys())  # Collect all unique usernames
    
    for username in usernames:
        row = {'Username': username}
        for file in all_scores:
            row[file] = all_scores[file].get(username, 0)  # Default to 0 if user has no score in that file
        participants_data.append(row)
    
    # Create DataFrame
    participants_df = pd.DataFrame(participants_data)
    
    # Calculate FinalPoints as sum of all file columns
    file_columns = [col for col in participants_df.columns if col != 'Username']
    participants_df['FinalPoints'] = participants_df[file_columns].sum(axis=1)
    
    # Sort participants by FinalPoints (descending order)
    participants_df = participants_df.sort_values(by='FinalPoints', ascending=False)
    
    # Write the data to an Excel file
    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        # Write the participants sheet
        participants_df.to_excel(writer, sheet_name="Participants", index=False)
        print()
        print(f"💾 Saved Participants to 'Participants' sheet in {out_file}")
        
        # Form teams from sorted participants
        teams = [participants_df.iloc[i:i + team_size] for i in range(0, len(participants_df), team_size)]
        
        # Write each team to a separate sheet
        for idx, team in enumerate(teams):
            team_name = f"Team_{idx + 1}"
            team.to_excel(writer, sheet_name=team_name, index=False)
            print(f"💾 Saved {team_name} to {out_file}")

def main():
    print("🏆 VJudge Team Maker — Local Excel Mode")
    
    # Input directory where the Leaderboards XLSX files are stored
    Leaderboards_dir = "Leaderboards"
    
    # List all Excel files in the directory
    files = [f for f in os.listdir(Leaderboards_dir) if f.endswith('.xlsx')]
    if not files:
        print(f"❌ No Excel files found in '{Leaderboards_dir}' directory. Exiting.")
        return

    # Prompt for team size
    try:
        team_size = int(input("👥 Enter team size (default 3): ").strip() or "3")
    except ValueError:
        team_size = 3

    # Dictionary to accumulate individual file scores
    all_scores = defaultdict(dict)
    saved_any = False

    # Process each Excel file
    for file in files:
        file_path = os.path.join(Leaderboards_dir, file)
        print(f"\n📡 Processing: {file_path}")
        standings = read_excel_file(file_path)
        
        if standings:
            saved_any = True
            # Store points for each user per file
            for u, r in standings:
                points = points_for_rank(r)
                all_scores[file][u] = points  # Store per-file points for each user
        else:
            print(f"[WARN] No standings obtained from {file_path}")

    if saved_any:
        # Write all participants with their points to the Excel file
        write_participants_and_teams_to_excel(all_scores, team_size, out_file="final_teams.xlsx")
        print("\n✅ All done.")
    else:
        print("\n❌ No points computed. Nothing to save.")

if __name__ == "__main__":
    main()
