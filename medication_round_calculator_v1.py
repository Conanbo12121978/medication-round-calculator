"""
Medication Round Calculator v1
Reads:
  Master_Matrix
  Master_Long
  Round_Sets
  Rules

The Excel workbook is the source of truth.

Important:
- Dose #1 is always the actual STAT date/time.
- Master_Long supplies the selected Round_set, Next_round,
  Next_day, and Day_offset.
- For Q24H, the workbook currently has Day_offset = 1 for every row.
- A Round value of 24 means 00:00 on the date specified by Day_offset.
"""

from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd


# ============================================================
# Utility functions
# ============================================================

def _clock_to_minutes(value):
    """
    Convert clock value to minutes from midnight.

    Examples:
        06       -> 360
        06:00    -> 360
        24       -> 1440
        24:00    -> 1440

    Important:
        24:00 is represented internally as 1440.
        It will later be interpreted as 00:00 of the
        date specified by Day_offset.
    """
    if pd.isna(value):
        return None

    s = str(value).strip()

    if s.endswith(".0"):
        s = s[:-2]

    if ":" in s:
        h, m = map(int, s.split(":"))
    else:
        h, m = int(s), 0

    if h == 24 and m == 0:
        return 1440

    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ValueError(f"Invalid clock value: {value}")

    return h * 60 + m


def _parse_datetime(text):
    """
    Accept:
        DD/MM/YYYY HH:MM
    """
    return datetime.strptime(text, "%d/%m/%Y %H:%M")


def _round_times(round_set):
    """
    Convert a Round_set such as:
        06-12-18-24

    into minutes from midnight.
    """
    return [_clock_to_minutes(x) for x in str(round_set).split("-")]


# ============================================================
# Main calculator
# ============================================================

class MedicationRoundCalculator:

    def __init__(self, excel_file):

        self.excel_file = Path(excel_file)

        if not self.excel_file.exists():
            raise FileNotFoundError(
                f"Excel workbook not found: {self.excel_file}"
            )

        # ----------------------------------------------------
        # Read Excel workbook
        # ----------------------------------------------------

        self.master_long = pd.read_excel(
            self.excel_file,
            sheet_name="Master_Long"
        )

        self.round_sets = pd.read_excel(
            self.excel_file,
            sheet_name="Round_Sets"
        )

        self.rules = pd.read_excel(
            self.excel_file,
            sheet_name="Rules"
        )

        # ----------------------------------------------------
        # Validate Master_Long columns
        # ----------------------------------------------------

        required = {
            "STAT_time",
            "Frequency",
            "Round_set",
            "Next_round",
            "Next_day",
            "Day_offset"
        }

        missing = required - set(self.master_long.columns)

        if missing:
            raise ValueError(
                "Master_Long is missing columns: "
                + ", ".join(sorted(missing))
            )

    # ========================================================
    # Find Master_Long row
    # ========================================================

    def master_row(self, stat_dt, frequency):
        """
        Find the exact STAT-time rule for the requested frequency.
        """

        clock = stat_dt.strftime("%H:%M")

        df = self.master_long.copy()

        df["Frequency"] = (
            df["Frequency"]
            .astype(str)
            .str.strip()
        )

        df = df[df["Frequency"].eq(frequency)]

        if df.empty:
            raise ValueError(
                f"No Master_Long rows for {frequency}"
            )

        # ----------------------------------------------------
        # Exact STAT time match
        # ----------------------------------------------------

        exact = df[
            df["STAT_time"].map(
                self._same_clock(clock)
            )
        ]

        if not exact.empty:

            # ------------------------------------------------
            # Prefer Use_for_auto_recommendation = Yes
            # ------------------------------------------------

            if "Use_for_auto_recommendation" in exact.columns:

                yes = exact[
                    exact[
                        "Use_for_auto_recommendation"
                    ]
                    .astype(str)
                    .str.lower()
                    .isin(
                        ["yes", "y", "true", "1"]
                    )
                ]

                if not yes.empty:
                    exact = yes

            return exact.iloc[0]

        raise ValueError(
            f"No exact STAT_time={clock} rule "
            f"for {frequency}"
        )

    # ========================================================
    # Compare clock values
    # ========================================================

    @staticmethod
    def _same_clock(target):

        target_min = _clock_to_minutes(target)

        def compare(x):

            try:
                return (
                    _clock_to_minutes(x)
                    == target_min
                )

            except Exception:
                return False

        return compare

    # ========================================================
    # Calculate Dose #2
    # ========================================================

    @staticmethod
    def _next_datetime(stat_dt, row):
        """
        Use the workbook's:

            Next_round
            Next_day
            Day_offset

        Important rule:

            24 means 00:00 on the date specified
            by Day_offset.

        Therefore:

            Day_offset = 1
            Next_round = 24

        means:

            STAT date + 1 day at 00:00

        NOT:

            STAT date + 2 days at 00:00
        """

        # ----------------------------------------------------
        # Read Next_round
        # ----------------------------------------------------

        next_round = _clock_to_minutes(
            row["Next_round"]
        )

        # ----------------------------------------------------
        # Read Day_offset
        # ----------------------------------------------------

        if not pd.isna(row["Day_offset"]):

            day_offset = int(
                row["Day_offset"]
            )

        else:

            day_offset = 0

        # ----------------------------------------------------
        # Backward-compatible fallback
        #
        # If Next_day = Yes but Day_offset = 0,
        # automatically use +1 day.
        # ----------------------------------------------------

        next_day = (
            str(row["Next_day"])
            .strip()
            .lower()
            == "yes"
        )

        if next_day and day_offset == 0:
            day_offset = 1

        # ----------------------------------------------------
        # Next_round must exist
        # ----------------------------------------------------

        if next_round is None:

            raise ValueError(
                "Next_round is blank for "
                "this Master_Long row"
            )

        # ----------------------------------------------------
        # Base date
        # ----------------------------------------------------

        base_date = (
            stat_dt.date()
            + timedelta(days=day_offset)
        )

        # ----------------------------------------------------
        # IMPORTANT:
        #
        # 24:00 = 00:00 of base_date
        #
        # Do NOT add another day here.
        # ----------------------------------------------------

        if next_round == 1440:

            return datetime.combine(
                base_date,
                datetime.min.time()
            )

        # ----------------------------------------------------
        # Normal clock value
        # ----------------------------------------------------

        return (
            datetime.combine(
                base_date,
                datetime.min.time()
            )
            + timedelta(minutes=next_round)
        )

    # ========================================================
    # Find interval from Round_Sets
    # ========================================================

    def _get_interval(self, selected_round):

        rs = self.round_sets.copy()

        rs["Frequency"] = (
            rs["Frequency"]
            .astype(str)
            .str.strip()
        )

        rs["Round_set"] = (
            rs["Round_set"]
            .astype(str)
            .str.strip()
        )

        match = rs[
            rs["Round_set"].eq(
                selected_round
            )
        ]

        if match.empty:

            raise ValueError(
                f"Round_set "
                f"{selected_round!r} "
                f"not found in Round_Sets"
            )

        interval_hours = float(
            match.iloc[0]["Interval_hours"]
        )

        return timedelta(
            hours=interval_hours
        )

    # ========================================================
    # Main calculation
    # ========================================================

    def calculate(
        self,
        stat_text,
        frequency,
        dose_count=5,
        round_set=None
    ):
        """
        Return:

            DataFrame
            selected Round_set

        Dose #1:
            Always actual STAT datetime.

        Dose #2:
            Determined by Master_Long.

        Dose #3 onward:

            Q24H:
                every 24 hours from Dose #2.

            Other frequencies:
                use Interval_hours from Round_Sets.
        """

        # ----------------------------------------------------
        # Validate dose count
        # ----------------------------------------------------

        if dose_count < 1:

            raise ValueError(
                "dose_count must be >= 1"
            )

        # ----------------------------------------------------
        # Parse STAT datetime
        # ----------------------------------------------------

        stat_dt = _parse_datetime(
            stat_text
        )

        # ----------------------------------------------------
        # Find Master_Long row
        # ----------------------------------------------------

        row = self.master_row(
            stat_dt,
            frequency
        )

        # ----------------------------------------------------
        # Determine selected Round_set
        # ----------------------------------------------------

        if round_set:

            selected_round = (
                str(round_set)
                .strip()
            )

        else:

            selected_round = (
                str(row["Round_set"])
                .strip()
            )

        # ----------------------------------------------------
        # Dose #1
        # ----------------------------------------------------

        dates = [stat_dt]

        # ----------------------------------------------------
        # Dose #2
        # ----------------------------------------------------

        if dose_count >= 2:

            dose2 = self._next_datetime(
                stat_dt,
                row
            )

            dates.append(dose2)

        # ----------------------------------------------------
        # Determine interval for Dose #3 onward
        # ----------------------------------------------------

        if frequency == "Q24H":

            # Rules sheet:
            # Q24H Dose #3 onward continues
            # every 24 hours from Dose #2.

            interval = timedelta(
                hours=24
            )

        else:

            interval = self._get_interval(
                selected_round
            )

        # ----------------------------------------------------
        # Dose #3 onward
        # ----------------------------------------------------

        while len(dates) < dose_count:

            dates.append(
                dates[-1] + interval
            )

        # ----------------------------------------------------
        # Build result DataFrame
        # ----------------------------------------------------

        result = pd.DataFrame(
            {
                "Dose": range(
                    1,
                    dose_count + 1
                ),

                "Date": [
                    d.strftime("%d/%m/%Y")
                    for d in dates
                ],

                "Time": [
                    d.strftime("%H:%M")
                    for d in dates
                ],

                "DateTime": dates,
            }
        )

        return result, selected_round


# ============================================================
# Test / Validation
# ============================================================

if __name__ == "__main__":

    workbook = (
        "round_master_Q24H_fixed.xlsx"
    )

    calc = MedicationRoundCalculator(
        workbook
    )

    # --------------------------------------------------------
    # Test 1: Q24H
    #
    # STAT:
    #   23/09/2026 23:00
    #
    # Master_Long:
    #   Next_round = 24
    #   Day_offset = 1
    #
    # Expected:
    #
    #   Dose #1 = 23/09/2026 23:00
    #   Dose #2 = 24/09/2026 00:00
    #   Dose #3 = 25/09/2026 00:00
    #   Dose #4 = 26/09/2026 00:00
    #   Dose #5 = 27/09/2026 00:00
    # --------------------------------------------------------

    result, selected_round = calc.calculate(
        "23/09/2026 23:00",
        "Q24H",
        dose_count=5,
    )

    print(
        f"Round set: {selected_round}"
    )

    print(
        result.to_string(
            index=False
        )
    )
