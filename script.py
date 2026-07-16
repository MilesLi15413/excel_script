import time
import os
import re
import shutil
import tempfile
import threading
import datetime as dt
import hashlib
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
        "sheet": "2026njo",
        "count": 0
    }
]

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

ORDINALS = {1: "1st", 2: "2nd", 3: "3rd"}


# ============================================================================
# Upload logic
# ============================================================================
def upload_to_drive(drive_service, excel_file, sheet_name, count):
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

    print(f"Sheets uploaded to '{sheet_name}'")

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

        # The decimal (e.g. .3 in "10.3") is a shootout marker, not a
        # fractional goal -- .3 means "won the shootout by 3". Below, the
        # RAW w_dec/d_dec still decide who won the shootout for points
        # purposes (3/2 split) and for head-to-head comparisons -- that
        # logic is unchanged. What's new: the shootout digit also gets
        # added to the regulation goal count to form a combined goal total,
        # and THAT total (not just the regulation score) is what feeds
        # Goals Scored / Goals Allowed / Goal Differential. Previously a
        # shootout win added nothing to GS/GA/GD, which meant an all-
        # shootout pool could never be separated by GD at all (everyone's
        # combined GD was still 0-0-0, exactly like a plain tie) -- now a
        # shootout win shows up in the goal totals the same way a
        # regulation goal would.
        w_shootout_goals = round(w_dec * 10)
        d_shootout_goals = round(d_dec * 10)
        w_total = w_int + w_shootout_goals
        d_total = d_int + d_shootout_goals

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
        teams[g["white"]]["gs"] += w_total
        teams[g["white"]]["ga"] += d_total
        teams[g["dark"]]["gs"] += d_total
        teams[g["dark"]]["ga"] += w_total

        # Head-to-head comparisons stay based on the original regulation
        # score + shootout decimal -- this is about who literally won that
        # specific game (same rule as points: regulation decides, shootout
        # decimal breaks a regulation tie), not about goal totals.
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


def placeholder_prefixes(ordinal, prefix):
    """Two known conventions seen across tabs:
    - space+dash:  '1st pt_O-'  (most tabs)
    - underscore:  '1st_pt_O'   (e.g. 12U_Coed_Champ-45)
    Returns the raw prefix strings to check a cell against via startswith --
    NOT exact match. This matters: once a placeholder gets resolved (e.g.
    '1st pt_O-' becomes '1st pt_O-DIABLO ALLIANCE B'), the cell no longer
    equals the raw unresolved text. If a later score correction changes who
    should actually be 1st, matching only the raw text would silently skip
    that cell forever, leaving the WRONG team advanced with no error. This
    version recognizes the cell whether it's still raw or already resolved
    with some (possibly now-stale) team name, so corrections can overwrite it."""
    return [f"{ordinal} {prefix}-", f"{ordinal}_{prefix}"]


def resolve_cell_if_match(get_cell_str, set_cell, ordinal, prefix, team_code, location_desc, verbose=True):
    """Checks one cell against both known placeholder conventions (raw or
    already-resolved) and overwrites it with the current correct team name
    if it matches and differs from what's already there. Returns True if it
    wrote a change. verbose=False suppresses the print (used on the Sheets
    side, since the live Excel side already prints the same result and we
    don't want every pool logged twice)."""
    team_name = team_display_name(team_code)
    cell_str = get_cell_str()

    for raw_prefix in placeholder_prefixes(ordinal, prefix):
        if cell_str == raw_prefix or cell_str.startswith(raw_prefix):
            separator = "" if raw_prefix.endswith("-") else "-"
            new_value = f"{raw_prefix}{separator}{team_name}"
            if cell_str != new_value:
                set_cell(new_value)
                if verbose:
                    verb = "Resolved" if cell_str == raw_prefix else "Updated (correction)"
                    print(f"  {verb} '{cell_str}' -> '{new_value}' at {location_desc}")
                return True
            return False
    return False


def reset_stale_placeholders(values, prefix, set_cell, location_desc_fn, verbose=True):
    """If a pool that was previously fully scored and resolved becomes
    incomplete again (a score got erased, a game row got deleted, etc.),
    any cell for this pool's placeholders that currently shows a resolved
    team name gets reset back to its raw, unresolved form (e.g.
    '1st pt_U-BULLDOG' -> '1st pt_U-') instead of silently keeping a stale
    answer that no longer reflects reality now that the pool is undecided
    again. Scans all three ranks against both known placeholder
    conventions. Returns True if it reset anything. verbose=False
    suppresses the print (Sheets side runs silently, same reasoning as
    resolve_cell_if_match)."""
    changed = False
    for rank_num in (1, 2, 3):
        ordinal = ORDINALS[rank_num]
        raw_forms = placeholder_prefixes(ordinal, prefix)

        for r_idx, row in enumerate(values):
            for c_idx, cell in enumerate(row):
                cell_str = str(cell).strip() if cell is not None else ""
                for raw in raw_forms:
                    if cell_str.startswith(raw) and cell_str != raw:
                        set_cell(r_idx, c_idx, raw)
                        if verbose:
                            print(f"  Reset stale '{cell_str}' -> '{raw}' at {location_desc_fn(r_idx, c_idx)} "
                                  f"(pool '{prefix}' no longer fully scored)")
                        changed = True
                        break
    return changed


# ============================================================================
# LIVE IN-PLACE: talk directly to your already-open Excel session via
# xlwings, so auto-advance edits land in the actual working file -- no
# closing required, no duplicate file needed.
# ============================================================================
# xw.Book(path) is the key piece: if that file is already open in a running
# Excel instance, xlwings attaches to that exact session (Excel still owns
# the file, so there's no lock conflict). If it's not open, xlwings opens it
# fresh. Either way we get a live handle to work with directly.
def get_live_workbook(excel_file, quiet=False):
    if not XLWINGS_AVAILABLE:
        if not quiet:
            print("[XLWINGS] xlwings not installed -- run: pip install xlwings")
        return None
    try:
        return xw.Book(excel_file)
    except Exception as e:
        if not quiet:
            print(f"[XLWINGS] Could not attach to workbook: {e}")
        return None


FIXED_READ_RANGE = "A1:Z475"


def read_sheet_values_xw(sht):
    """Normalizes xlwings' range values (which can come back as a scalar, a
    flat list, or a list of lists depending on the range's shape) into a
    consistent list-of-rows shape, matching what find_pools() expects.

    Deliberately reads a fixed, generously-sized range (A1:Z600) instead of
    sht.used_range. Excel's UsedRange is a tracked boundary that can lag
    behind very recent edits, especially edits near the current edge of
    that boundary -- a score typed seconds ago can be genuinely on the
    sheet but not yet reflected in used_range, until Excel recalculates it
    (e.g. on a full close/reopen). That caused exactly the symptom seen
    live: a newly-completed pool would resolve correctly moments later via
    the Sheets-side path (which re-reads the whole uploaded file fresh) but
    get silently missed by the live xlwings path THIS cycle. Reading a
    fixed range sidesteps the staleness entirely -- extra blank cells
    beyond your real data are harmless, find_pools() already ignores them."""
    raw = sht.range(FIXED_READ_RANGE).value
    if raw is None:
        return []
    if not isinstance(raw, list):
        return [[raw]]
    if not isinstance(raw[0], list):
        return [raw]
    return raw


def compute_data_signature(excel_file, busy_retries=2, busy_delay=1.0):
    """Returns a signature representing the actual game data (White/Dark/
    score cells across all sheets), NOT raw file bytes and NOT mtime. This
    matters because Excel's AutoSave (on by default for OneDrive files) or
    OneDrive's own sync process periodically re-saves the file in the
    background even when no score actually changed -- every save updates
    internal file metadata (an embedded 'last modified' timestamp inside
    the .xlsx itself), which changes the mtime AND the raw file bytes even
    with zero real data change. Comparing actual cell values instead means
    background autosave/sync churn no longer falsely triggers a new
    auto-advance cycle.

    A transient "Excel is busy" moment (same OLE error as elsewhere) is
    retried briefly rather than immediately falling back to a raw file-byte
    hash -- switching signature methods between calls would itself look
    like a false "change" even when nothing in the data actually moved,
    since a value-hash and a file-byte-hash of the same data never match.

    Falls back to hashing raw file bytes only if xlwings truly isn't
    available/attached after retries."""
    for attempt in range(busy_retries):
        wb = get_live_workbook(excel_file, quiet=True)
        if wb is not None:
            try:
                parts = []
                for ws in wb.sheets:
                    parts.append(f"{ws.name}:{read_sheet_values_xw(ws)}")
                return hashlib.md5("".join(parts).encode("utf-8", errors="ignore")).hexdigest()
            except Exception as e:
                if _is_excel_busy_error(e) and attempt < busy_retries - 1:
                    time.sleep(busy_delay)
                    continue
                break  # not a busy error, or out of retries -- fall through to file-hash
        else:
            break  # xlwings not available/attached at all -- no point retrying

    try:
        with open(excel_file, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()
    except Exception:
        return None



def resolve_downstream_placeholders_xw(ws, values, prefix, result):
    if result["true_tie"]:
        return

    for rank_num, team_code in result["rank"].items():
        ordinal = ORDINALS[rank_num]

        for r_idx, row in enumerate(values):
            for c_idx, cell in enumerate(row):
                cell_str = str(cell).strip() if cell is not None else ""

                def get_str(s=cell_str):
                    return s

                def set_val(v, r=r_idx, c=c_idx):
                    ws.range((r + 1, c + 1)).value = v

                resolve_cell_if_match(
                    get_cell_str=get_str,
                    set_cell=set_val,
                    ordinal=ordinal, prefix=prefix, team_code=team_code,
                    location_desc=f"[XLWINGS] row {r_idx + 1}, col {c_idx + 1} in '{ws.name}'"
                )


EXCEL_BUSY_HRESULT = -2146777998  # 0x800AC472 -- "Excel is busy right now, try again"


def _is_excel_busy_error(exc):
    return bool(exc.args) and isinstance(exc.args[0], int) and exc.args[0] == EXCEL_BUSY_HRESULT


_last_tab_signature = {}  # {(excel_file, tab_name): hash} -- persists across calls
                          # within this process, so unchanged tabs get skipped


def _hash_tab_values(values):
    return hashlib.md5(str(values).encode("utf-8", errors="ignore")).hexdigest()


def run_tiebreaker_live_xw(excel_file, busy_retries=3, busy_delay=1.5):
    """Attempts to run the full tiebreaker + auto-advance pass live, directly
    against your already-open Excel workbook via xlwings. Returns True if it
    ran (and saved), False if xlwings isn't available or something went
    wrong attaching (caller should fall back to the duplicate-file approach).

    Excel occasionally rejects an incoming automation call with OLE error
    0x800AC472 -- this specifically means "I'm busy right now" (mid-save,
    mid-recalculation, a cell still in edit mode, etc.), not a real failure.
    It usually clears within a second or two, so this retries a few times
    on that specific error before giving up and falling back.

    Also skips any tab whose actual data hasn't changed since the last run
    (tracked via _last_tab_signature) -- previously every tab got fully
    re-read and re-scanned on every single save, even ones untouched since
    the last change. Now only the tab(s) that actually changed get
    reprocessed, which matters more the more tabs/divisions a workbook has."""
    for attempt in range(busy_retries):
        wb = get_live_workbook(excel_file, quiet=(attempt > 0))
        if wb is None:
            return False

        try:
            any_processed = False

            for ws in wb.sheets:
                values = read_sheet_values_xw(ws)

                tab_key = (excel_file, ws.name)
                tab_hash = _hash_tab_values(values)
                if _last_tab_signature.get(tab_key) == tab_hash:
                    continue  # nothing changed on this tab since last run
                _last_tab_signature[tab_key] = tab_hash

                pools = find_pools(values)

                for prefix, games in pools.items():
                    is_complete = len(games) == 3 and all(
                        g["white_score"] not in ("", None) and g["dark_score"] not in ("", None)
                        for g in games
                    )

                    if not is_complete:
                        def set_cell_xw(r, c, v, ws=ws):
                            ws.range((r + 1, c + 1)).value = v

                        def loc_xw(r, c, ws=ws):
                            return f"row {r + 1}, col {c + 1} in '{ws.name}'"

                        if reset_stale_placeholders(values, prefix, set_cell_xw, loc_xw):
                            any_processed = True
                        continue

                    try:
                        result = compute_three_team_standings(games)
                        print(f"excel computed for pool '{prefix}' in tab '{ws.name}'")
                        resolve_downstream_placeholders_xw(ws, values, prefix, result)
                        any_processed = True
                    except Exception as e:
                        print(f"[XLWINGS] Error computing tiebreaker for pool '{prefix}' in '{ws.name}': {e}")

            if any_processed:
                wb.save()
                print("[XLWINGS] Auto-advance applied live to the open workbook.")
            return True
        except Exception as e:
            if _is_excel_busy_error(e) and attempt < busy_retries - 1:
                print(f"[XLWINGS] Excel is momentarily busy, retrying in {busy_delay}s... "
                      f"(attempt {attempt + 1}/{busy_retries})")
                time.sleep(busy_delay)
                continue
            print(f"[XLWINGS] Unexpected error during live auto-advance: {e}")
            return False

    return False



# ============================================================================
# GOOGLE SHEETS: same tiebreaker pass, applied to the uploaded copy
# ============================================================================

def queue_downstream_placeholder_writes(pending_writes, tab, values, prefix, result):
    """Same matching logic as resolve_cell_if_match, but queues the write
    into pending_writes instead of executing an API call immediately --
    lets run_tiebreaker_for_spreadsheet send every cell write from an
    entire run in a single batchUpdate instead of one call per cell."""
    if result["true_tie"]:
        return

    for rank_num, team_code in result["rank"].items():
        ordinal = ORDINALS[rank_num]

        for r_idx, row in enumerate(values):
            for c_idx, cell in enumerate(row):
                cell_str = str(cell).strip()
                cell_range = f"{tab}!{colnum_to_letter(c_idx + 1)}{r_idx + 1}"

                def get_str(s=cell_str):
                    return s

                def set_val(v, rng=cell_range):
                    pending_writes.append({"range": rng, "values": [[v]]})

                resolve_cell_if_match(
                    get_cell_str=get_str,
                    set_cell=set_val,
                    ordinal=ordinal, prefix=prefix, team_code=team_code,
                    location_desc=cell_range
                )


def run_tiebreaker_for_spreadsheet(sheets_service, spreadsheet_id):
    """Runs the tiebreaker pass across every tab, batching all reads and
    writes into as few Sheets API calls as possible -- each individual API
    call is a full network round trip, and the old per-cell approach could
    add up to dozens of sequential round trips per run. Now it's: one call
    to list tabs, one batched read of every tab's data, and one batched
    write for every resolved placeholder cell computed this run."""
    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    tab_names = [s["properties"]["title"] for s in meta["sheets"]]

    # One batched read for every tab's data, instead of one GET per tab.
    batch_get = sheets_service.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id, ranges=tab_names
    ).execute()
    tab_values = {
        tab_names[i]: vr.get("values", [])
        for i, vr in enumerate(batch_get.get("valueRanges", []))
    }

    pending_writes = []

    for tab in tab_names:
        values = tab_values.get(tab, [])
        pools = find_pools(values)

        for prefix, games in pools.items():
            is_complete = len(games) == 3 and all(
                g["white_score"] not in ("", None) and g["dark_score"] not in ("", None) for g in games
            )

            if not is_complete:
                def set_cell_sheets(r, c, v, tab=tab):
                    pending_writes.append({"range": f"{tab}!{colnum_to_letter(c + 1)}{r + 1}", "values": [[v]]})

                def loc_sheets(r, c, tab=tab):
                    return f"{tab}!{colnum_to_letter(c + 1)}{r + 1}"

                reset_stale_placeholders(values, prefix, set_cell_sheets, loc_sheets)
                continue

            try:
                result = compute_three_team_standings(games)
                print(f"Sheets computed for pool '{prefix}' in tab '{tab}'")
                queue_downstream_placeholder_writes(pending_writes, tab, values, prefix, result)
            except Exception as e:
                print(f"Error computing tiebreaker for pool '{prefix}': {e}")

    if pending_writes:
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"valueInputOption": "RAW", "data": pending_writes}
        ).execute()


# ============================================================================
# Watch loop
# ============================================================================
def watch(config, drive_service, sheets_service):
    excel_file = config["excel"]
    sheet_name = config["sheet"]

    def apply_auto_advance():
        print(f"Change detected in {os.path.basename(excel_file)}!")

        file_id = upload_to_drive(drive_service, excel_file, sheet_name, config["count"])
        if file_id:
            run_tiebreaker_for_spreadsheet(sheets_service, file_id)

        # Live in-place edit directly in your open Excel session. Runs
        # independently of the Sheets computation above -- if this fails
        # (xlwings missing, Excel not currently open with this file), the
        # Sheets side above already ran fine on its own; only your local
        # file misses the live edit until this succeeds on a later save.
        run_tiebreaker_live_xw(excel_file)

        print(f"FINISHED {config['count']}")
        print("-" * 40)
        return file_id

    config["count"] += 1
    apply_auto_advance()

    # Two-stage gate, so this only actually runs when you deliberately save
    # (Ctrl+S / Save button), not on every live in-memory keystroke:
    #
    # 1. mtime gate: os.path.getmtime() only changes when Excel actually
    #    WRITES the file to disk -- i.e. on a save. It does NOT change just
    #    because a cell was edited in memory. This is the "did a save just
    #    happen at all" check.
    #
    # 2. data-signature gate: even a real save can happen with zero actual
    #    score changes (e.g. Excel's own AutoSave firing on a timer, or you
    #    saving without having changed anything). compute_data_signature
    #    hashes the actual game data (not file bytes/metadata), so a save
    #    with no real content change doesn't trigger a wasted auto-advance
    #    cycle either.
    #
    # Both are read AFTER apply_auto_advance(), since our own live writes
    # (resolved placeholders, Standings tab) just changed both -- otherwise
    # we'd immediately mistake our own edit for a new user save.
    last_modified = os.path.getmtime(excel_file)
    last_signature = compute_data_signature(excel_file)

    while True:
        time.sleep(2)
        try:
            current_modified = os.path.getmtime(excel_file)
            if current_modified != last_modified:
                current_signature = compute_data_signature(excel_file)
                if current_signature != last_signature:
                    config["count"] += 1
                    apply_auto_advance()
                    last_signature = compute_data_signature(excel_file)
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