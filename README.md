# next time formation will be calculated using the equation:
# Point = ceil(1600 / (rank + 7))


# Team Formation From VJudge Standings

This repository provides a **Python tool** to process VJudge contest results from Excel exports, calculate points for each participant, and form balanced teams automatically. It supports multiple contests and generates a final Excel file with both participant rankings and team allocations.

---

## Repository Structure


```

Team_Formation_From_Vjudge_Standings/

│

├─ Leaderboards/                  # Folder containing VJudge contest standings in Excel

│   ├─ 01.xlsx

│   ├─ 02.xlsx

│   ├─ 03.xlsx

│   └─ ...                        # Add as many contest Excel files as needed

│

├─ Final_Teams.xlsx               # Output Excel file containing participants & teams

├─ Vjudge_Contest_Ranker.py       # Main Python script

└─ README.md


````

---

## Requirements

- Python 3.10 or higher
- Python libraries:

```bash
pip install pandas openpyxl
````

---

## Excel File Format

Each contest Excel file in the `Leaderboards` folder must have the following columns:

| Username     | Rank |
| ------------ | ---- |
| participant1 | 1    |
| participant2 | 2    |
| participant3 | 3    |
| ...          | ...  |

* **Username**: The participant's exact name (case-sensitive).
* **Rank**: Integer representing the participant’s rank in that contest.
* Missing participants in a contest will automatically receive **0 points**.

---

## How to download contest standings as Excel File
<div align="center">
  <img src="https://github.com/maruf-hossain74/Sharing-Repo/blob/main/Images/Download_leaderboard.png" 
       alt="Vjudge standings" 
       width="50%">
</div>


---

## Points System

Points for each participant are calculated using the formula:

```
points = ceil(1800 / (rank + 5))
```

* Higher rank → more points.
* Points from multiple contests are **summed** to calculate `FinalPoints`.

---

## Usage Instructions

1. Place all contest Excel files in the `Leaderboards` folder.
2. Open a terminal in the project directory and run:

```bash
python Vjudge_Contest_Ranker.py
```

3. Enter the **team size** when prompted (default is 3).
4. The script will:

   * Read all contest Excel files.
   * Calculate points for each participant per contest.
   * Aggregate total points (`FinalPoints`) across contests.
   * Sort participants by total points.
   * Form teams consecutively based on rankings.
   * Output results in `Final_Teams.xlsx`:

     * **Participants sheet**: All participants with points per contest and total points.
     * **Team_1, Team_2, … sheets**: Each sheet contains one team.

---

## Workflow Diagram

```text
Leaderboards/             Python Script           Final_Teams.xlsx
--------------            --------------         -----------------
01.xlsx, 02.xlsx, 03.xlsx  --> Read & parse -->  Participants sheet
                            --> Calculate points --> Team_1, Team_2, ...
                            --> Aggregate points
                            --> Sort & form teams
```

---

## Example Workflow

1. Add contest files `01.xlsx`, `02.xlsx`, `03.xlsx` in the `Leaderboards` folder. (any file name can be added, it's not mendatory to add 01.xlsx or 02.xlsx or etc)
2. Run the script:

```
🏆 VJudge Team Maker — Local Excel Mode
👥 Enter team size (default 3): 3
```

## 3. Output:

<div align="center">
  <img src="https://github.com/maruf-hossain74/Sharing-Repo/blob/main/Images/output_Team_formation.png" 
       alt="Vjudge standings" 
       width="50%">
</div>

* `Final_Teams.xlsx` with sheets:

  * **Participants**: Shows usernames, points per contest, and `FinalPoints`.
  * **Team_1, Team_2, …**: Each sheet represents a team of participants.

---

## Notes

* Ensure Excel files are **not empty** and contain correct column names (`Username` and `Rank`).
* Participants missing from a contest are assigned **0 points** for that contest.
* The script currently works with **local Excel exports** only; it does not fetch data directly from VJudge online.

---

## Dependencies

* [pandas](https://pandas.pydata.org/)
* [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

Install via:

```bash
pip install pandas openpyxl
```

---

## License

MIT License © Maruf Hossain
