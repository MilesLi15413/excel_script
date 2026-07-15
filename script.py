import time
import os
import re
import shutil
import tempfile
import threading
import datetime as dt
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False

BASE_DIR = r"C:\Users\miles\OneDrive\Desktop\onedrive"

CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")

WATCHES = [
    {
        "excel": os.path.join(BASE_DIR, "2026NJOCOPY.xlsx"),
        "live_copy": os.path.join(BASE_DIR, "2026NJOCOPY_LIVE.xlsx"),
        "sheet": "2026njo",
        "count": 0
    }
]

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
# Pool detection (shared by both local-Excel and Google Sheets paths)
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
    (no trailing digit, so it's a downstream placeholder, not a pool game).

    Only matches real pool-prefix conventions: one or more lowercase letters,
    an underscore, then a capital letter (e.g. 'pt_O', 'au_M', 'ag_A', 'bz_R'),
    followed by a single trailing seed digit. This deliberately does NOT match
    bracket-elimination codes like 'W29' or 'L28' (winner/loser of game #),
    which have no underscore and would otherwise get misread as a pool prefix
    with the last digit stripped off (e.g. 'L28' -> prefix 'L2', digit '8'),
    incorrectly bucketing unrelated bracket games together as a fake pool."""
    value = str(value).strip()
    code = value.split("-")[0].strip()  # drop team name if present
    m = re.match(r"^([a-z]+_[A-Z])(\d)$", code)
    return m.group(1) if m else None


def team_display_name(code):
    """'pt_O2-DIABLO ALLIANCE B' -> 'DIABLO ALLIANCE B'. Falls back to the
    raw code if there's no dash (shouldn't normally happen for a real team)."""
    code = str(code).strip()
    return code.split("-", 1)[1].strip() if "-" in code else code


def safe_get(row, idx, default=""):
    return row[idx] if idx is not None and idx < len(row) else default


EXCEL_EPOCH = dt.datetime(1899, 12, 31)


def normalize_score(value):
    """Some score cells (White/Dark 'S' columns) carry inherited date/time
    number formats even though a scorer just typed a plain number in them
    (e.g. '8'). openpyxl reads those back as datetime/time objects instead
    of the number, since it infers type from the cell's number_format. This
    reverses that back into the real numeric score."""
    if isinstance(value, dt.datetime):
        delta = value - EXCEL_EPOCH
        return delta.days + delta.seconds / 86400 + delta.microseconds / 86400000000
    if isinstance(value, dt.time):
        return (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6) / 86400
    return value


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

        white_score = normalize_score(safe_get(row, col_map.get("S_1")))
        dark_score = normalize_score(safe_get(row, col_map.get("S_2")))

        pools.setdefault(prefix, []).append({
            "row": r_idx,
            "white": white,
            "dark": dark,
            "white_score": white_score,
            "dark_score": dark_score,
        })
    return pools


# ============================================================================
# Tiebreaker math (shared)
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
# Downstream placeholder resolution -- shared regex helpers
# ============================================================================
def colnum_to_letter(n):
    """1 -> A, 2 -> B, ... 27 -> AA"""
    letters = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def placeholder_candidates(ordinal, prefix, team_code):
    """Two known conventions seen across tabs:
    - space+dash:  '1st pt_O-'  -> '1st pt_O-TEAM NAME'  (most tabs)
    - underscore:  '1st_pt_O'   -> '1st_pt_O-TEAM NAME'  (e.g. 12U_Coed_Champ-45)
    Returns {raw_placeholder_text: resolved_value}."""
    team_name = team_display_name(team_code)
    return {
        f"{ordinal} {prefix}-": f"{ordinal} {prefix}-{team_name}",
        f"{ordinal}_{prefix}": f"{ordinal}_{prefix}-{team_name}",
    }


# ============================================================================
# LIVE IN-PLACE: talk directly to your already-open Excel session via
# xlwings, so auto-advance edits land in the actual working file -- no
# closing required, no duplicate file needed.
# ============================================================================
# xw.Book(path) is the key piece: if that file is already open in a running
# Excel instance, xlwings attaches to that exact session (Excel still owns
# the file, so there's no lock conflict). If it's not open, xlwings opens it
# fresh. Either way we get a live handle to work with directly.
def get_live_workbook(excel_file):
    if not XLWINGS_AVAILABLE:
        print("[XLWINGS] xlwings not installed -- run: pip install xlwings")
        return None
    try:
        return xw.Book(excel_file)
    except Exception as e:
        print(f"[XLWINGS] Could not attach to workbook: {e}")
        return None


def read_sheet_values_xw(sht):
    """Normalizes xlwings' used_range.value (which can come back as a
    scalar, a flat list, or a list of lists depending on the range's shape)
    into a consistent list-of-rows shape, matching what find_pools() expects."""
    raw = sht.used_range.value
    if raw is None:
        return []
    if not isinstance(raw, list):
        return [[raw]]
    if not isinstance(raw[0], list):
        return [raw]
    return raw


def ensure_standings_sheet_xw(wb):
    names = [s.name for s in wb.sheets]
    if STANDINGS_TAB in names:
        return wb.sheets[STANDINGS_TAB]
    ws = wb.sheets.add(name=STANDINGS_TAB, after=wb.sheets[-1])
    ws.range((1, 1)).value = [["Pool", "1st", "2nd", "3rd", "Narrative", "Last Updated"]]
    return ws


def write_standings_row_xw(standings_ws, prefix, result):
    target_row = None
    r = 1
    while standings_ws.range((r, 1)).value not in (None, ""):
        if standings_ws.range((r, 1)).value == prefix:
            target_row = r
            break
        r += 1
    if target_row is None:
        target_row = r

    if result["true_tie"]:
        row_values = [prefix, "TRUE TIE", "TRUE TIE", "TRUE TIE", result["narrative"],
                      time.strftime("%Y-%m-%d %H:%M:%S")]
    else:
        row_values = [
            prefix,
            team_display_name(result["rank"][1]),
            team_display_name(result["rank"][2]),
            team_display_name(result["rank"][3]),
            result["narrative"],
            time.strftime("%Y-%m-%d %H:%M:%S")
        ]
    standings_ws.range((target_row, 1)).value = [row_values]


def resolve_downstream_placeholders_xw(ws, values, prefix, result):
    if result["true_tie"]:
        return

    for rank_num, team_code in result["rank"].items():
        ordinal = ORDINALS[rank_num]
        candidates = placeholder_candidates(ordinal, prefix, team_code)

        for r_idx, row in enumerate(values):
            for c_idx, cell in enumerate(row):
                cell_str = str(cell).strip() if cell is not None else ""
                if cell_str in candidates:
                    new_value = candidates[cell_str]
                    ws.range((r_idx + 1, c_idx + 1)).value = new_value
                    print(f"  [XLWINGS] Resolved '{cell_str}' -> '{new_value}' "
                          f"at row {r_idx + 1}, col {c_idx + 1} in '{ws.name}'")


def run_tiebreaker_live_xw(excel_file):
    """Attempts to run the full tiebreaker + auto-advance pass live, directly
    against your already-open Excel workbook via xlwings. Returns True if it
    ran (and saved), False if xlwings isn't available or something went
    wrong attaching (caller should fall back to the duplicate-file approach)."""
    wb = get_live_workbook(excel_file)
    if wb is None:
        return False

    try:
        standings_ws = ensure_standings_sheet_xw(wb)

        for ws in wb.sheets:
            if ws.name == STANDINGS_TAB:
                continue

            values = read_sheet_values_xw(ws)
            pools = find_pools(values)

            for prefix, games in pools.items():
                if len(games) != 3:
                    continue
                all_scored = all(
                    g["white_score"] not in ("", None) and g["dark_score"] not in ("", None)
                    for g in games
                )
                if not all_scored:
                    continue

                try:
                    result = compute_three_team_standings(games)
                    write_standings_row_xw(standings_ws, prefix, result)
                    print(f"[XLWINGS] Tiebreaker computed for pool '{prefix}' in tab '{ws.name}'")
                    resolve_downstream_placeholders_xw(ws, values, prefix, result)
                except Exception as e:
                    print(f"[XLWINGS] Error computing tiebreaker for pool '{prefix}' in '{ws.name}': {e}")

        wb.save()
        print("[XLWINGS] Auto-advance applied live to the open workbook.")
        return True
    except Exception as e:
        print(f"[XLWINGS] Unexpected error during live auto-advance: {e}")
        return False


# ============================================================================
# LIVE DUPLICATE: write auto-advance results into a separate file, never
# touching the file you actually have open in Excel
# ============================================================================
# Excel holds an exclusive lock on your working file for as long as it's
# open -- there's no reliable way for another process to write into that
# same file without you closing it first. Rather than fight that lock,
# auto-advance results get written into a SEPARATE file sitting next to your
# working copy (e.g. "2026NJOCOPY_LIVE.xlsx"). That duplicate is never
# opened by Excel, so nothing ever locks it -- writes to it always succeed.
#
# Flow: you edit scores in your working file -> script uploads a read-only
# copy to Drive (this already works fine while the file is open) -> Google
# Sheets computes the tiebreaker/auto-advance -> the finished result gets
# exported and saved as the live-copy file. Open the live-copy file anytime
# to see current standings and resolved bracket placeholders, side by side
# with your working file.
def download_sheet_as_excel(drive_service, spreadsheet_id, live_copy_path, attempts=5, delay=2):
    request = drive_service.files().export_media(
        fileId=spreadsheet_id,
        mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    data = request.execute()

    for attempt in range(attempts):
        try:
            with open(live_copy_path, "wb") as f:
                f.write(data)
            print(f"[LIVE COPY] Updated '{os.path.basename(live_copy_path)}' with auto-advance results.")
            return True
        except PermissionError:
            # only happens if you have the live-copy file itself open somewhere
            print(f"[LIVE COPY] '{os.path.basename(live_copy_path)}' is locked "
                  f"(attempt {attempt + 1}/{attempts}), retrying in {delay}s...")
            time.sleep(delay)

    print(f"[LIVE COPY] Could not write '{os.path.basename(live_copy_path)}' -- "
          "close it if you have it open, and it'll catch up on the next change.")
    return False



# ============================================================================
# GOOGLE SHEETS: same tiebreaker pass, applied to the uploaded copy
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


def resolve_downstream_placeholders(sheets_service, spreadsheet_id, tab, values, prefix, result):
    if result["true_tie"]:
        return  # nothing to resolve -- can't advance a true tie automatically

    for rank_num, team_code in result["rank"].items():
        ordinal = ORDINALS[rank_num]
        candidates = placeholder_candidates(ordinal, prefix, team_code)

        for r_idx, row in enumerate(values):
            for c_idx, cell in enumerate(row):
                cell_str = str(cell).strip()
                if cell_str in candidates:
                    new_value = candidates[cell_str]
                    cell_range = f"{tab}!{colnum_to_letter(c_idx + 1)}{r_idx + 1}"
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=cell_range,
                        valueInputOption="RAW",
                        body={"values": [[new_value]]}
                    ).execute()
                    print(f"  Resolved '{cell_str}' -> '{new_value}' at {cell_range}")


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
    live_copy_path = config["live_copy"]
    sheet_name = config["sheet"]

    def apply_auto_advance():
        # Try live in-place first (xlwings, talks directly to your open
        # Excel session). Falls back to the Sheets + duplicate-file path
        # if xlwings isn't installed or the file isn't currently open.
        live_ok = run_tiebreaker_live_xw(excel_file)

        # Upload either way, so Google Sheets stays in sync with whatever
        # just happened locally.
        file_id = upload_to_drive(drive_service, excel_file, sheet_name, config["count"])
        if file_id:
            run_tiebreaker_for_spreadsheet(sheets_service, file_id)
            if not live_ok:
                download_sheet_as_excel(drive_service, file_id, live_copy_path)
        return file_id

    config["count"] += 1
    apply_auto_advance()

    # xlwings' wb.save() just touched excel_file's mtime -- read it fresh
    # AFTER that save, so we don't mistake our own edit for a new change
    last_modified = os.path.getmtime(excel_file)

    while True:
        time.sleep(2)
        try:
            current_modified = os.path.getmtime(excel_file)
            if current_modified != last_modified:
                config["count"] += 1
                apply_auto_advance()

                # same reasoning: our own save just changed mtime again --
                # re-read it after the fact, not before, so the next poll
                # compares against the post-save state, not the pre-save one
                last_modified = os.path.getmtime(excel_file)
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