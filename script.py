import time
import os
import re
import shutil
import tempfile
import threading
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

BASE_DIR = r"C:\Users\miles\OneDrive\Desktop\onedrive"

CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")

WATCHES = [
    {
        "excel": os.path.join(BASE_DIR, "2026NJOCOPY.xlsx"),
    }]

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

STANDINGS_TAB = "Standings"
ORDINALS = {1: "1st", 2: "2nd", 3: "3rd"}


# ============================================================================
# Upload logic
# ============================================================================
def upload_to_drive(drive_service, excel_file, sheet_name, count):
    print(f"Change detected in {os.path.basename(excel_file)}! Uploading to '{sheet_name}'...")

    results = drive_service.files().list(
        q=f"name='{sheet_name}' and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)",
        includeItemsFromAllDrives=True,
        supportsAllDrives=True
    ).execute()

    files = results.get('files', [])
    if not files:
        print(f"Google Sheet '{sheet_name}' not found! Make sure the name matches exactly.")
        return None

    file_id = files[0]['id']

    temp_dir = tempfile.mkdtemp()
    temp_file = os.path.join(temp_dir, "temp_upload.xlsx")
    shutil.copy2(excel_file, temp_file)

    media = MediaFileUpload(
        temp_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )

    drive_service.files().update(
        fileId=file_id,
        media_body=media,
        body={},
        supportsAllDrives=True
    ).execute()

    media._fd.close()

    for _ in range(5):
        time.sleep(2)
        try:
            shutil.rmtree(temp_dir)
            break
        except Exception:
            continue

    print(f"Upload complete! '{sheet_name}' updated with formatting. | ##: {count}")
    print("-" * 40)

    return file_id


# ============================================================================
# Pool detection
# ============================================================================
def get_col_map_for_row(row_values):
    col_map = {}
    s_count = 0
    for c, label in enumerate(row_values):
        label = str(label).strip()
        if label == "S":
            s_count += 1
            col_map[f"S_{s_count}"] = c
        elif label:
            col_map[label] = c
    return col_map if "White" in col_map and "Dark" in col_map else None


def find_header_map_for_row(all_rows, target_idx):
    for r in range(target_idx, -1, -1):
        row = all_rows[r] if r < len(all_rows) else []
        col_map = get_col_map_for_row(row)
        if col_map:
            return col_map
    return None


def extract_pool_prefix(value):
    """'pt_M1' -> 'pt_M'. 'pt_O2-DIABLO ALLIANCE B' -> 'pt_O'
    (team name after a dash is ignored). '3rd pt_O-' -> None
    (no trailing digit, so it's a downstream placeholder, not a pool game)."""
    value = str(value).strip()
    code = value.split("-")[0].strip()  # drop team name if present
    m = re.match(r"^(.+?)(\d)$", code)
    return m.group(1) if m else None


def team_display_name(code):
    """'pt_O2-DIABLO ALLIANCE B' -> 'DIABLO ALLIANCE B'. Falls back to the
    raw code if there's no dash (shouldn't normally happen for a real team)."""
    code = str(code).strip()
    return code.split("-", 1)[1].strip() if "-" in code else code


def safe_get(row, idx, default=""):
    return row[idx] if idx is not None and idx < len(row) else default


def find_pools(all_rows):
    pools = {}
    for r_idx, row in enumerate(all_rows):
        col_map = get_col_map_for_row(row)
        if not col_map:
            col_map = find_header_map_for_row(all_rows, r_idx)
        if not col_map or "S_1" not in col_map or "S_2" not in col_map:
            continue

        white = str(safe_get(row, col_map.get("White"))).strip()
        dark = str(safe_get(row, col_map.get("Dark"))).strip()

        prefix = extract_pool_prefix(white) or extract_pool_prefix(dark)
        if not prefix:
            continue

        white_score = safe_get(row, col_map.get("S_1"))
        dark_score = safe_get(row, col_map.get("S_2"))

        pools.setdefault(prefix, []).append({
            "row": r_idx,
            "white": white,
            "dark": dark,
            "white_score": white_score,
            "dark_score": dark_score,
        })
    return pools


# ============================================================================
# Tiebreaker math
# ============================================================================
def compute_three_team_standings(games):
    teams = {}
    h2h = {}

    def ensure(name):
        if name not in teams:
            teams[name] = {"name": name, "pts": 0, "gs": 0, "ga": 0}

    for g in games:
        ensure(g["white"])
        ensure(g["dark"])

        w_score = float(g["white_score"])
        d_score = float(g["dark_score"])
        w_int, d_int = int(w_score // 1), int(d_score // 1)
        w_dec, d_dec = w_score - w_int, d_score - d_int

        if w_int > d_int:
            w_pts, d_pts = 4, 1
        elif d_int > w_int:
            w_pts, d_pts = 1, 4
        elif w_dec > d_dec:
            w_pts, d_pts = 3, 2
        elif d_dec > w_dec:
            w_pts, d_pts = 2, 3
        else:
            w_pts, d_pts = 2, 2

        teams[g["white"]]["pts"] += w_pts
        teams[g["dark"]]["pts"] += d_pts
        teams[g["white"]]["gs"] += w_int
        teams[g["white"]]["ga"] += d_int
        teams[g["dark"]]["gs"] += d_int
        teams[g["dark"]]["ga"] += w_int

        h2h[(g["white"], g["dark"])] = {"a": w_int, "b": d_int, "a_dec": w_dec, "b_dec": d_dec}
        h2h[(g["dark"], g["white"])] = {"a": d_int, "b": w_int, "a_dec": d_dec, "b_dec": w_dec}

    names = list(teams.keys())
    for n in names:
        teams[n]["gd"] = teams[n]["gs"] - teams[n]["ga"]

    narrative = ""
    locked = None
    pair = None

    def try_stage(stat_key, label, higher_is_better):
        nonlocal locked, pair, narrative
        if locked:
            return
        vals = [teams[n][stat_key] for n in names]
        if all(v == vals[0] for v in vals):
            narrative += f"All three tied on {label} ({vals[0]}). "
            return
        sorted_names = sorted(names, key=lambda n: teams[n][stat_key], reverse=higher_is_better)
        best = teams[sorted_names[0]][stat_key]
        worst = teams[sorted_names[2]][stat_key]
        best_count = sum(1 for n in names if teams[n][stat_key] == best)
        worst_count = sum(1 for n in names if teams[n][stat_key] == worst)

        if best_count == 1:
            locked = {"pos": 1, "team": sorted_names[0]}
            pair = [sorted_names[1], sorted_names[2]]
            narrative += f"{sorted_names[0]} has the best {label}, takes 1st. "
        elif worst_count == 1:
            locked = {"pos": 3, "team": sorted_names[2]}
            pair = [sorted_names[0], sorted_names[1]]
            narrative += f"{sorted_names[2]} is clearly behind on {label}, takes 3rd. "

    try_stage("pts", "Points", True)
    try_stage("gd", "Goal Differential", True)
    try_stage("gs", "Goals Scored", True)
    try_stage("ga", "Goals Allowed", False)

    true_tie = False
    rank = {}

    if not locked:
        true_tie = True
        narrative += "True three-way tie -- cannot be resolved by these stats."
    else:
        a, b = pair
        matchup = h2h[(a, b)]
        tie = False
        if matchup["a"] > matchup["b"]:
            winner, loser = a, b
        elif matchup["b"] > matchup["a"]:
            winner, loser = b, a
        elif matchup["a_dec"] > matchup["b_dec"]:
            winner, loser = a, b
            narrative += f"{a} won a shootout over {b}. "
        elif matchup["b_dec"] > matchup["a_dec"]:
            winner, loser = b, a
            narrative += f"{b} won a shootout over {a}. "
        else:
            tie = True
            winner = loser = None
            narrative += f"Head-to-head between {a} and {b} was also a tie."

        if locked["pos"] == 1:
            rank[1] = locked["team"]
            rank[2] = a if tie else winner
            rank[3] = b if tie else loser
        else:
            rank[3] = locked["team"]
            rank[1] = a if tie else winner
            rank[2] = b if tie else loser

    return {"narrative": narrative, "true_tie": true_tie, "rank": rank}


# ============================================================================
# Standings tab
# ============================================================================
def ensure_standings_sheet_exists(sheets_service, spreadsheet_id):
    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    tab_names = [s["properties"]["title"] for s in meta["sheets"]]

    if STANDINGS_TAB not in tab_names:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": STANDINGS_TAB}}}]}
        ).execute()
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{STANDINGS_TAB}!A1",
            valueInputOption="RAW",
            body={"values": [["Pool", "1st", "2nd", "3rd", "Narrative", "Last Updated"]]}
        ).execute()


def write_standings_row(sheets_service, spreadsheet_id, prefix, result):
    existing = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"{STANDINGS_TAB}!A:A"
    ).execute().get("values", [])

    target_row = None
    for i, row in enumerate(existing):
        if row and row[0] == prefix:
            target_row = i + 1
            break
    if target_row is None:
        target_row = len(existing) + 1

    if result["true_tie"]:
        row_values = [prefix, "TRUE TIE", "TRUE TIE", "TRUE TIE", result["narrative"], time.strftime("%Y-%m-%d %H:%M:%S")]
    else:
        row_values = [
            prefix,
            team_display_name(result["rank"][1]),
            team_display_name(result["rank"][2]),
            team_display_name(result["rank"][3]),
            result["narrative"],
            time.strftime("%Y-%m-%d %H:%M:%S")
        ]

    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{STANDINGS_TAB}!A{target_row}",
        valueInputOption="RAW",
        body={"values": [row_values]}
    ).execute()


# ============================================================================
# Downstream placeholder resolution -- e.g. "3rd pt_O-" -> "3rd pt_O-VIPER PIGEON"
# ============================================================================
def colnum_to_letter(n):
    """1 -> A, 2 -> B, ... 27 -> AA"""
    letters = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def resolve_downstream_placeholders(sheets_service, spreadsheet_id, tab, values, prefix, result):
    if result["true_tie"]:
        return  # nothing to resolve -- can't advance a true tie automatically

    for rank_num, team_code in result["rank"].items():
        ordinal = ORDINALS[rank_num]
        placeholder = f"{ordinal} {prefix}-"
        new_value = f"{ordinal} {prefix}-{team_display_name(team_code)}"

        for r_idx, row in enumerate(values):
            for c_idx, cell in enumerate(row):
                if str(cell).strip() == placeholder:
                    cell_range = f"{tab}!{colnum_to_letter(c_idx + 1)}{r_idx + 1}"
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=cell_range,
                        valueInputOption="RAW",
                        body={"values": [[new_value]]}
                    ).execute()
                    print(f"  Resolved '{placeholder}' -> '{new_value}' at {cell_range}")


# ============================================================================
# Main tiebreaker pass
# ============================================================================
def run_tiebreaker_for_spreadsheet(sheets_service, spreadsheet_id):
    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    tab_names = [s["properties"]["title"] for s in meta["sheets"] if s["properties"]["title"] != STANDINGS_TAB]

    ensure_standings_sheet_exists(sheets_service, spreadsheet_id)

    for tab in tab_names:
        values = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=tab
        ).execute().get("values", [])

        pools = find_pools(values)

        for prefix, games in pools.items():
            if len(games) != 3:
                continue
            all_scored = all(g["white_score"] not in ("", None) and g["dark_score"] not in ("", None) for g in games)
            if not all_scored:
                continue

            try:
                result = compute_three_team_standings(games)
                write_standings_row(sheets_service, spreadsheet_id, prefix, result)
                print(f"Tiebreaker computed for pool '{prefix}' in tab '{tab}'")
                resolve_downstream_placeholders(sheets_service, spreadsheet_id, tab, values, prefix, result)
            except Exception as e:
                print(f"Error computing tiebreaker for pool '{prefix}': {e}")


# ============================================================================
# Watch loop
# ============================================================================
def watch(config, drive_service, sheets_service):
    excel_file = config["excel"]
    sheet_name = config["sheet"]

    config["count"] += 1
    file_id = upload_to_drive(drive_service, excel_file, sheet_name, config["count"])
    if file_id:
        run_tiebreaker_for_spreadsheet(sheets_service, file_id)

    last_modified = os.path.getmtime(excel_file)

    while True:
        time.sleep(2)
        try:
            current_modified = os.path.getmtime(excel_file)
            if current_modified != last_modified:
                last_modified = current_modified
                config["count"] += 1
                file_id = upload_to_drive(drive_service, excel_file, sheet_name, config["count"])
                if file_id:
                    run_tiebreaker_for_spreadsheet(sheets_service, file_id)
        except Exception as e:
            print(f"Error watching {os.path.basename(excel_file)}: {e}")
            print("-" * 40)


if __name__ == "__main__":
    print("Starting watchers...")

    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    threads = []
    for config in WATCHES:
        t = threading.Thread(target=watch, args=(config, drive_service, sheets_service))
        t.daemon = True
        threads.append(t)
        t.start()

    while True:
        time.sleep(1)