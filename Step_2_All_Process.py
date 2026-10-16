import os
from datetime import datetime, timedelta
import xlwings as xw


# ==========================================================
# CONFIG
# ==========================================================

YPR = r"C:\Users\srishti sinha\Documents\H5706\H5706\YPR"

STLY = r"C:\Users\srishti sinha\Documents\H5706\H5706\files (27)\stly_v2_pipeline\stly_v2\Final_Output"

TW = "Data_Table_TW"
LW = "Data_Table_LW"
SOTM = "Data_Table_SOTM"
YESTERDAY = "Data_Table_Yesterday"

OTB_2027 = "2027 OTB (Opera)"
OTB = "2026 OTB (Opera)"

STLY_SHEET = "STLY"


# ==========================================================
# DATES
# ==========================================================

today = datetime.today()
yesterday = today - timedelta(days=1)

lw_date = today - timedelta(days=7)
lw_fallback_date = today - timedelta(days=8)

month_start = today.replace(day=1)

prev_month = (
    month_start - timedelta(days=1)
).replace(day=1)


# ==========================================================
# YPR FILE PATHS
# ==========================================================

yesterday_ypr = os.path.join(
    YPR,
    f"H5706_YPR_{yesterday:%d_%B}.xlsx"
)

today_ypr = os.path.join(
    YPR,
    f"H5706_YPR_{today:%d_%B}.xlsx"
)


# ==========================================================
# TW STLY FILE
# ==========================================================

tw_stly = os.path.join(
    STLY,
    f"{today:%d_%b_%Y}",
    "STLY_Output.xlsx"
)


# ==========================================================
# LW STLY FILE - CURRENT DATE - 7 DAYS
# ==========================================================

lw_stly_7 = os.path.join(
    STLY,
    f"{lw_date:%d_%b_%Y}",
    "STLY_Output.xlsx"
)


# ==========================================================
# LW STLY FILE - CURRENT DATE - 8 DAYS
# ==========================================================

lw_stly_8 = os.path.join(
    STLY,
    f"{lw_fallback_date:%d_%b_%Y}",
    "STLY_Output.xlsx"
)


# ==========================================================
# SOTM STLY FILE
# ==========================================================

sotm_stly = os.path.join(
    STLY,
    f"{month_start:%d_%b_%Y}",
    "STLY_Output.xlsx"
)


# ==========================================================
# YESTERDAY STLY FILE
# ==========================================================

yesterday_stly_normal = os.path.join(
    STLY,
    f"{yesterday:%d_%b_%Y}",
    "STLY_Output.xlsx"
)


# ==========================================================
# LW FILE SELECTION
# ==========================================================

if os.path.exists(lw_stly_7):

    lw_stly = lw_stly_7
    lw_used_date = lw_date

    print(
        f"LW file found for Current Date - 7 Days: "
        f"{lw_date:%d-%b-%Y}"
    )

else:

    if os.path.exists(lw_stly_8):

        lw_stly = lw_stly_8
        lw_used_date = lw_fallback_date

        print(
            f"LW file NOT found for Current Date - 7 Days: "
            f"{lw_date:%d-%b-%Y}"
        )

        print(
            f"Using fallback Current Date - 8 Days: "
            f"{lw_fallback_date:%d-%b-%Y}"
        )

    else:

        raise FileNotFoundError(
            "LW STLY file not found for both:\n"
            f"7 Days: {lw_stly_7}\n"
            f"8 Days: {lw_stly_8}"
        )


# ==========================================================
# FIND LATEST PREVIOUS-DATE STLY FILE
# ==========================================================

def find_latest_previous_stly_file(root_folder):

    latest_file = None
    latest_folder_date = None

    print(
        "\nSearching for latest PREVIOUS-DATE "
        "STLY_Output.xlsx..."
    )

    for folder_name in os.listdir(root_folder):

        folder_path = os.path.join(
            root_folder,
            folder_name
        )

        if not os.path.isdir(folder_path):
            continue

        try:

            folder_date = datetime.strptime(
                folder_name,
                "%d_%b_%Y"
            )

        except ValueError:

            continue

        if folder_date.date() >= today.date():
            continue

        stly_file = os.path.join(
            folder_path,
            "STLY_Output.xlsx"
        )

        if not os.path.exists(stly_file):
            continue

        if (
            latest_folder_date is None
            or folder_date > latest_folder_date
        ):

            latest_folder_date = folder_date
            latest_file = stly_file

    if latest_file:

        print(
            "\nLatest PREVIOUS-DATE STLY file found:"
        )

        print(latest_file)

        print(
            f"File date: "
            f"{latest_folder_date:%d-%b-%Y}"
        )

        return latest_file, latest_folder_date

    return None, None


# ==========================================================
# YESTERDAY FILE SELECTION
# ==========================================================

if os.path.exists(yesterday_stly_normal):

    yesterday_stly = yesterday_stly_normal
    yesterday_source_date = yesterday

    print(
        "\nYesterday STLY file found:"
    )

    print(yesterday_stly)

else:

    print(
        "\nYesterday STLY file NOT found:"
    )

    print(yesterday_stly_normal)

    print(
        "\nSearching previous-date STLY files..."
    )

    (
        yesterday_stly,
        yesterday_source_date
    ) = find_latest_previous_stly_file(
        STLY
    )

    if yesterday_stly is None:

        raise FileNotFoundError(
            "Yesterday STLY file was not found "
            "and no previous-date STLY_Output.xlsx "
            "file was available."
        )

    print(
        "\nUsing latest previous-date file "
        "for Data_Table_Yesterday:"
    )

    print(yesterday_stly)


# ==========================================================
# GET SHEET SAFELY
# ==========================================================

def get_sheet(wb, sheet_name):

    target = sheet_name.strip().lower()

    for sh in wb.sheets:

        if sh.name.strip().lower() == target:
            return sh

    available = [
        sh.name
        for sh in wb.sheets
    ]

    raise ValueError(
        f"Sheet '{sheet_name}' not found.\n"
        f"Available sheets: {available}"
    )


# ==========================================================
# DATE CONVERSION
# ==========================================================

def dt(v):

    if isinstance(v, datetime):
        return v

    if (
        hasattr(v, "year")
        and hasattr(v, "month")
        and hasattr(v, "day")
    ):

        return datetime(
            v.year,
            v.month,
            v.day
        )

    if isinstance(v, str):

        for f in (
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%b-%Y",
            "%d-%B-%Y",
            "%d %b %Y",
            "%d %B %Y",
            "%m/%d/%Y",
            "%m-%d-%Y"
        ):

            try:

                return datetime.strptime(
                    v.strip(),
                    f
                )

            except:

                pass

    return None


# ==========================================================
# LAST ROW
# ==========================================================

def last_row(sh):

    return sh.range(
        "C" + str(
            sh.cells.last_cell.row
        )
    ).end("up").row


# ==========================================================
# STATUS / VISIBLE COLUMNS
# ==========================================================

def status_visible(sh):

    status = 12
    visible = 13

    for r in range(1, 11):

        for c in range(1, 50):

            v = sh.cells(r, c).value

            if v is None:
                continue

            v = str(v).strip().lower()

            if v == "status":
                status = c

            elif v == "visible":
                visible = c

    return status, visible


# ==========================================================
# PROCESS TW / LW / SOTM
# ==========================================================

def process_sheet(
    ysh,
    stly_path,
    sheet_name
):

    print(
        f"\nProcessing {sheet_name}"
    )

    if not os.path.exists(stly_path):

        raise FileNotFoundError(
            f"STLY file not found:\n"
            f"{stly_path}"
        )

    swb = app.books.open(
        stly_path
    )

    try:

        ssh = get_sheet(
            swb,
            STLY_SHEET
        )

        last_stly = last_row(
            ssh
        )

        data = ssh.range(
            f"A2:K{last_stly}"
        ).options(
            ndim=2
        ).value

        filtered = [
            row
            for row in data
            if row
            and len(row) >= 3
            and dt(row[2])
            and dt(row[2]) >= prev_month
        ]

        if not filtered:

            raise ValueError(
                f"No STLY data from "
                f"{prev_month:%d-%b-%Y} onward."
            )

        last_ypr = last_row(
            ysh
        )

        start = None

        for r in range(
            1,
            last_ypr + 1
        ):

            d = dt(
                ysh.cells(r, 3).value
            )

            if d and d >= prev_month:

                start = r
                break

        if start is None:

            raise ValueError(
                f"{prev_month:%d-%b-%Y} "
                f"not found in "
                f"{sheet_name} Column C."
            )

        end = (
            start
            + len(filtered)
            - 1
        )

        ysh.range(
            f"A{start}:M{last_ypr}"
        ).clear_contents()

        ysh.range(
            f"A{start}:K{end}"
        ).value = filtered

        status, visible = status_visible(
            ysh
        )

        ysh.range(
            (start, status),
            (end, status)
        ).value = 1

        ysh.range(
            (start, visible),
            (end, visible)
        ).value = 1

        ysh.visible = True

        print(
            f"{sheet_name}: "
            f"{len(filtered)} rows copied."
        )

        return len(filtered)

    finally:

        swb.close()


# ==========================================================
# PROCESS DATA_TABLE_YESTERDAY
# ==========================================================

def process_yesterday(
    ysh,
    stly_path
):

    print(
        "\nProcessing Data_Table_Yesterday"
    )

    print(
        f"Source STLY file:\n"
        f"{stly_path}"
    )

    if not os.path.exists(stly_path):

        raise FileNotFoundError(
            f"STLY file not found:\n"
            f"{stly_path}"
        )

    swb = app.books.open(
        stly_path
    )

    try:

        ssh = get_sheet(
            swb,
            STLY_SHEET
        )

        last_stly = last_row(
            ssh
        )

        data = ssh.range(
            f"A2:K{last_stly}"
        ).options(
            ndim=2
        ).value

        filtered = [
            row
            for row in data
            if row
            and len(row) >= 3
            and dt(row[2])
            and dt(row[2]) >= prev_month
        ]

        if not filtered:

            raise ValueError(
                f"No STLY data from "
                f"{prev_month:%d-%b-%Y} onward."
            )

        last_ypr = last_row(
            ysh
        )

        start = None

        for r in range(
            2,
            last_ypr + 1
        ):

            d = dt(
                ysh.cells(r, 3).value
            )

            if d and d >= prev_month:

                start = r
                break

        if start is None:

            raise ValueError(
                f"{prev_month:%d-%b-%Y} "
                f"not found in "
                f"Data_Table_Yesterday "
                f"Column C."
            )

        end = (
            start
            + len(filtered)
            - 1
        )

        ysh.range(
            f"A{start}:M{last_ypr}"
        ).clear_contents()

        ysh.range(
            f"A{start}:K{end}"
        ).value = filtered

        status, visible = status_visible(
            ysh
        )

        ysh.range(
            (start, status),
            (end, status)
        ).value = 1

        ysh.visible = True

        print(
            f"Data_Table_Yesterday "
            f"Status = 1 for rows "
            f"{start} to {end}"
        )

        print(
            f"Data_Table_Yesterday: "
            f"{len(filtered)} rows copied."
        )

        return len(filtered)

    finally:

        swb.close()


# ==========================================================
# START EXCEL
# ==========================================================

app = xw.App(
    visible=False,
    add_book=False
)

app.display_alerts = False
app.screen_updating = False

ywb = None

try:

    # ======================================================
    # STEP 1 - OPEN YESTERDAY YPR
    # ======================================================

    print("=" * 70)

    print(
        "STEP 1 : OPEN YESTERDAY YPR"
    )

    print("=" * 70)

    if not os.path.exists(
        yesterday_ypr
    ):

        raise FileNotFoundError(
            yesterday_ypr
        )

    ywb = app.books.open(
        yesterday_ypr
    )

    print(
        "Available YPR sheets:",
        [
            sh.name
            for sh in ywb.sheets
        ]
    )


    # ======================================================
    # STEP 2 - TW
    # ======================================================

    print(
        "\nSTEP 2 : DATA_TABLE_TW"
    )

    tw_sh = get_sheet(
        ywb,
        TW
    )

    tw_sh.visible = True

    tw_rows = process_sheet(
        tw_sh,
        tw_stly,
        TW
    )

    if os.path.exists(
        today_ypr
    ):

        os.remove(
            today_ypr
        )

    ywb.save(
        today_ypr
    )

    ywb.close()
    ywb = None


    # ======================================================
    # STEP 3 - LW
    # ======================================================

    print(
        "\nSTEP 3 : DATA_TABLE_LW"
    )

    print(
        f"LW source date selected: "
        f"{lw_used_date:%d-%b-%Y}"
    )

    print(
        f"LW source file:\n"
        f"{lw_stly}"
    )

    ywb = app.books.open(
        today_ypr
    )

    lw_sh = get_sheet(
        ywb,
        LW
    )

    lw_sh.visible = True

    lw_rows = process_sheet(
        lw_sh,
        lw_stly,
        LW
    )

    ywb.save(
        today_ypr
    )

    ywb.close()
    ywb = None


    # ======================================================
    # STEP 4 - SOTM
    # ======================================================

    print(
        "\nSTEP 4 : DATA_TABLE_SOTM"
    )

    ywb = app.books.open(
        today_ypr
    )

    sotm_sh = get_sheet(
        ywb,
        SOTM
    )

    sotm_sh.visible = True

    sotm_rows = process_sheet(
        sotm_sh,
        sotm_stly,
        SOTM
    )

    ywb.save(
        today_ypr
    )

    ywb.close()
    ywb = None


    # ======================================================
    # STEP 5 - DATA_TABLE_YESTERDAY
    # ======================================================

    print(
        "\nSTEP 5 : DATA_TABLE_YESTERDAY"
    )

    print(
        f"Selected source date: "
        f"{yesterday_source_date:%d-%b-%Y}"
    )

    print(
        f"Selected source file:\n"
        f"{yesterday_stly}"
    )

    ywb = app.books.open(
        today_ypr
    )

    yesterday_sh = get_sheet(
        ywb,
        YESTERDAY
    )

    yesterday_sh.visible = True

    yesterday_rows = process_yesterday(
        yesterday_sh,
        yesterday_stly
    )


    # ======================================================
    # STEP 6 - 2027 OTB (Opera)
    # ======================================================

    print(
        "\nSTEP 6 : 2027 OTB (Opera)"
    )

    otb_2027_sh = get_sheet(
        ywb,
        OTB_2027
    )

    otb_2027_sh.visible = True

    print(
        "Updating 2027 OTB (Opera) C2 "
        "with yesterday date..."
    )

    # ==================================================
    # C2 = YESTERDAY DATE
    #
    # Example:
    # 05 October 2026
    # ==================================================

    otb_2027_sh.range(
        "C2"
    ).value = yesterday

    otb_2027_sh.range(
        "C2"
    ).number_format = "dd mmmm yyyy"

    print(
        f"2027 OTB (Opera) C2 = "
        f"{yesterday:%d %B %Y}"
    )


    # ======================================================
    # STEP 7 - 2026 OTB (Opera)
    # ======================================================

    print(
        "\nSTEP 7 : 2026 OTB (Opera)"
    )

    otb_sh = get_sheet(
        ywb,
        OTB
    )

    otb_sh.visible = True

    print(
        "Unprotecting 2026 OTB (Opera)..."
    )

    otb_sh.api.Unprotect()

    print(
        "2026 OTB (Opera) sheet unprotected."
    )


    # ==================================================
    # A2 = YESTERDAY DAY ONLY
    # ==================================================

    otb_sh.range(
        "A2"
    ).value = yesterday.day

    otb_sh.range(
        "A2"
    ).number_format = "0"


    # ==================================================
    # C2 = YESTERDAY DATE
    #
    # Example:
    # 05 October 2026
    # ==================================================

    otb_sh.range(
        "C2"
    ).value = yesterday

    otb_sh.range(
        "C2"
    ).number_format = "dd mmmm yyyy"

    print(
        f"A2 = {yesterday.day}"
    )

    print(
        f"C2 = "
        f"{yesterday:%d %B %Y}"
    )


    # ==================================================
    # CURRENT MONTH REMAINING DAYS
    #
    # ONLY CHECK B4:B15
    #
    # Remaining days are written into
    # corresponding A cell.
    # ==================================================

    print(
        "\nUpdating current month remaining days..."
    )


    # --------------------------------------------------
    # FIRST DAY OF NEXT MONTH
    # --------------------------------------------------

    next_month = (
        today.replace(day=28)
        + timedelta(days=4)
    ).replace(day=1)


    # --------------------------------------------------
    # LAST DAY OF CURRENT MONTH
    # --------------------------------------------------

    last_day_current_month = (
        next_month - timedelta(days=1)
    ).day


    # --------------------------------------------------
    # REMAINING DAYS INCLUDING TODAY
    # --------------------------------------------------

    remaining_days = (
        last_day_current_month
        - today.day
        + 0
    )

    print(
        f"Current Month: "
        f"{today:%B %Y}"
    )

    print(
        f"Last Day of Month: "
        f"{last_day_current_month}"
    )

    print(
        f"Today's Day: "
        f"{today.day}"
    )

    print(
        f"Remaining Days: "
        f"{remaining_days}"
    )


    # ==================================================
    # FIND CURRENT MONTH ONLY IN B4:B15
    # ==================================================

    current_month_found = False
    current_month_row = None

    for r in range(4, 16):

        month_value = otb_sh.cells(
            r,
            2
        ).value

        if month_value is None:
            continue

        print(
            f"Checking B{r}: "
            f"{month_value}"
        )


        # ------------------------------------------------
        # CASE 1 - EXCEL DATE VALUE
        # ------------------------------------------------

        month_date = dt(month_value)

        if month_date:

            if (
                month_date.month == today.month
                and month_date.year == today.year
            ):

                otb_sh.cells(
                    r,
                    1
                ).value = remaining_days

                otb_sh.cells(
                    r,
                    1
                ).number_format = "0"

                current_month_found = True
                current_month_row = r

                print(
                    f"Current month found "
                    f"in B{r}"
                )

                print(
                    f"A{r} = "
                    f"{remaining_days}"
                )

                break

        else:

            continue


        # ------------------------------------------------
        # CASE 2 - TEXT MONTH
        #
        # Examples:
        #
        # Oct-26
        # Oct 2026
        # October-26
        # October 2026
        # 01-10-2026
        # 01/10/2026
        # ------------------------------------------------

        month_text = str(
            month_value
        ).strip()

        possible_formats = (
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%d.%m.%Y",
            "%d-%b-%Y",
            "%d-%B-%Y",
            "%d %b %Y",
            "%d %B %Y",
            "%b-%y",
            "%b %y",
            "%B-%y",
            "%B %y",
            "%b-%Y",
            "%b %Y",
            "%B-%Y",
            "%B %Y"
        )

        parsed_month = None

        for fmt in possible_formats:

            try:

                parsed_month = datetime.strptime(
                    month_text,
                    fmt
                )

                break

            except ValueError:

                pass

        if parsed_month:

            if (
                parsed_month.month == today.month
                and parsed_month.year == today.year
            ):

                otb_sh.cells(
                    r,
                    1
                ).value = remaining_days

                otb_sh.cells(
                    r,
                    1
                ).number_format = "0"

                current_month_found = True
                current_month_row = r

                print(
                    f"Current month "
                    f"'{month_text}' "
                    f"found in B{r}"
                )

                print(
                    f"A{r} = "
                    f"{remaining_days}"
                )

                break


    # ==================================================
    # CURRENT MONTH NOT FOUND
    # ==================================================

    if not current_month_found:

        print(
            "\nWARNING: Current month "
            f"'{today:%B %Y}' "
            "was not found in B4:B15."
        )

        print(
            "Checked only cells B4:B15."
        )


    # ======================================================
    # STEP 8 - HIDE FOUR DATA SHEETS
    # ======================================================

    print(
        "\nSTEP 8 : HIDING ALL FOUR SHEETS"
    )

    for sheet_name in [
        TW,
        LW,
        SOTM,
        YESTERDAY
    ]:

        sh = get_sheet(
            ywb,
            sheet_name
        )

        sh.visible = False

        print(
            f"{sheet_name} = HIDDEN"
        )


    # ======================================================
    # STEP 9 - MAKE 2026 OTB ACTIVE SHEET
    # ======================================================

    print(
        "\nSTEP 9 : "
        "SETTING 2026 OTB (Opera) "
        "AS ACTIVE SHEET"
    )

    otb_sh.visible = True

    otb_sh.activate()

    otb_sh.range(
        "A1"
    ).select()

    print(
        "2026 OTB (Opera) "
        "is now the active sheet."
    )


    # ======================================================
    # FINAL SAVE
    # ======================================================

    print(
        "\nFINAL SAVE"
    )

    ywb.save(
        today_ypr
    )

    print(
        "\n" + "=" * 70
    )

    print(
        "PROCESS COMPLETED SUCCESSFULLY"
    )

    print(
        "=" * 70
    )

    print(
        f"Today YPR       : "
        f"{today_ypr}"
    )

    print(
        f"TW rows         : "
        f"{tw_rows}"
    )

    print(
        f"LW rows         : "
        f"{lw_rows}"
    )

    print(
        f"SOTM rows       : "
        f"{sotm_rows}"
    )

    print(
        f"Yesterday rows  : "
        f"{yesterday_rows}"
    )

    print(
        "Yesterday Status: 1"
    )

    print(
        f"2027 OTB C2     : "
        f"{yesterday:%d %B %Y}"
    )

    print(
        f"OTB A2          : "
        f"{yesterday.day}"
    )

    print(
        f"OTB C2          : "
        f"{yesterday:%d %B %Y}"
    )

    print(
        f"Current Month   : "
        f"{today:%B %Y}"
    )

    print(
        f"Remaining Days  : "
        f"{remaining_days}"
    )

    if current_month_row:

        print(
            f"Month Row       : "
            f"{current_month_row}"
        )

        print(
            f"Month Cell      : "
            f"B{current_month_row}"
        )

        print(
            f"Remaining Cell  : "
            f"A{current_month_row}"
        )

    print(
        f"Data from       : "
        f"{prev_month:%d-%b-%Y}"
    )

    print(
        f"LW source used  : "
        f"{lw_used_date:%d-%b-%Y}"
    )

    print(
        f"Yesterday source: "
        f"{yesterday_source_date:%d-%b-%Y}"
    )

    print(
        f"Yesterday file  : "
        f"{yesterday_stly}"
    )

    print(
        "Columns copied  : A:K"
    )

    print(
        "TW/LW/SOTM Status: 1"
    )

    print(
        "Yesterday Status: 1"
    )

    print(
        "\n2027 OTB (Opera) "
        "C2 updated with yesterday date."
    )

    print(
        "\n2026 OTB (Opera) "
        "will open as the active sheet."
    )

    print(
        "All four data sheets "
        "are HIDDEN after final save."
    )

    print(
        "=" * 70
    )


except Exception as e:

    print(
        "\n" + "=" * 70
    )

    print(
        "ERROR OCCURRED"
    )

    print(
        "=" * 70
    )

    print(
        str(e)
    )

    print(
        "=" * 70
    )

    raise


finally:

    try:

        if ywb is not None:
            ywb.close()

    except:

        pass

    try:

        app.screen_updating = True
        app.display_alerts = True
        app.quit()

    except:

        pass