"""
🏆 VJudge Team Maker — Local Excel Mode

Working Process:
----------------
1. The script looks for Excel leaderboard files inside the folder named 'Leaderboards'. 
   Each file should represent a single contest and contain at least two columns: 'Username' and 'Rank'.

2. For each Excel file:
   - The script reads all rows and extracts participants' usernames and their ranks.
   - For each participant, it calculates the points using the formula:
        points = ceil(1800 / (rank + 5))

3. The points from all contests are aggregated:
   - Each participant's points from different contests are summed to calculate their total points.

4. The script prompts the user to enter a team size (default = 3).

5. Sorting and Team Formation:
   - Participants are sorted in descending order based on their total points (FinalPoints).
   - Teams are formed consecutively (top-ranked participants first) with the specified team size.

6. Output:
   - The final Excel file ('final_teams.xlsx') is generated.
   - Sheet "Participants": all participants with points per contest and FinalPoints.
   - Sheets "Team_1", "Team_2", ...: each sheet contains a single team with participants and their points.

7. Notes:
   - If a participant does not appear in a contest, they are assigned 0 points for that contest.
   - The script handles multiple Excel files in the 'Leaderboards' directory automatically.
   - Requires pandas and openpyxl libraries:
       pip install pandas openpyxl
       
       


# File name (in the same directory)
file_name = "requirements.txt"  # change this to your file name

# Get full path of the file
file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

# Check if file exists
if not os.path.exists(file_path):
    print(f"File '{file_name}' not found.")
    sys.exit(1)

# Read dependencies line by line
with open(file_path, "r") as f:
    dependencies = [line.strip() for line in f if line.strip() and not line.startswith("#")]

# Install each dependency
for dep in dependencies:
    subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
"""




import os
import sys
import math
import subprocess
import pandas as pd
from collections import defaultdict

# ---------- Points formula ----------
def points_for_rank(rank: int) -> int:
    """Compute points based on rank."""
    return math.ceil(1600 / (rank + 7))

def read_excel_file(file_path: str):
    """Read an Excel file and return a list of (username, rank) tuples."""
    try:
        df = pd.read_excel(file_path)
        # Assuming columns are named "Username" and "Rank", adjust if necessary
        standings = []
        for _, row in df.iterrows():
            username = str(row['Team']).strip()
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
    # File name (in the same directory)
    file_name = "Requirements.txt"  # change this to your file name

    # Get full path of the file
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

    # Check if file exists
    if not os.path.exists(file_path):
        print(f"File '{file_name}' not found.")
        sys.exit(1)

    # Read dependencies line by line
    with open(file_path, "r") as f:
        dependencies = [line.strip() for line in f if line.strip() and not line.startswith("#")]

    # Install each dependency
    for dep in dependencies:
        subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
        
    # Input directory where the Leaderboards XLSX files are stored
    Leaderboards_dir = "Leaderboards"
    
    # List all Excel files in the directory
    files = [f for f in os.listdir(Leaderboards_dir) if f.endswith('.xlsx')]
    if not files:
        print(f"❌ No Excel files found in '{Leaderboards_dir}' directory. Exiting.")
        return

    # Prompt for team size
    print("\n💡 Each team will be formed based on final ranking points.")
    print("   For example, if you enter 3, teams of 3 members will be formed sequentially from top scorers.\n")
    print()
    
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
    print()
    print("🏆 ICPC/IUPC Team Formation — Local Excel Mode")
    print("""
    📘 Instructions:
    1️⃣  Go to your VJudge contest page while logged in.
    2️⃣  Click the 'Setting' icon on the standings page.
    3️⃣  Then Click on 'Rank' to download the leaderboard file.
    4️⃣  Save all downloaded contest files inside the 'Leaderboards' folder.
        Example: Leaderboards/01.xlsx, Leaderboards/02.xlsx, Leaderboards/03.xlsx
    5️⃣  Then run this script to calculate total points and form balanced teams.
    """)

    main()
