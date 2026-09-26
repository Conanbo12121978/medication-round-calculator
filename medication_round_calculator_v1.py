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
- Master_Long supplies the selected Round_set, Next_round, Next_day, Day_offset.
- For Q24H, the workbook currently has Day_offset = 1 for every row.
- A Round value of 24 means 00:00 on the following calendar date.
"""

from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd


def _clock_to_minutes(value):
    """06, 06:00, 24, 24:00 -> minutes from midnight; 24:00 = 1440."""
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
    """Accept DD/MM/YYYY HH:MM."""
    return datetime.strptime(text, "%d/%m/%Y %H:%M")


def _round_times(round_set):
    return [_clock_to_minutes(x) for x in str(round_set).split("-")]


class MedicationRoundCalculator:
    def __init__(self, excel_file):
        self.excel_file = Path(excel_file)
        self.master_long = pd.read_excel(
            self.excel_file, sheet_name="Master_Long"
        )
        self.round_sets = pd.read_excel(
            self.excel_file, sheet_name="Round_Sets"
        )
        self.rules = pd.read_excel(
            self.excel_file, sheet_name="Rules"
        )

        required = {
            "STAT_time", "Frequency", "Round_set",
            "Next_round", "Next_day", "Day_offset"
        }
        missing = required - set(self.master_long.columns)
        if missing:
            raise ValueError(
                "Master_Long is missing columns: "
                + ", ".join(sorted(missing))
            )

    def master_row(self, stat_dt, frequency):
        """Find the exact STAT-time rule for the requested frequency."""
        clock = stat_dt.strftime("%H:%M")

        df = self.master_long.copy()
        df["Frequency"] = df["Frequency"].astype(str).str.strip()
        df = df[df["Frequency"].eq(frequency)]

        if df.empty:
            raise ValueError(f"No Master_Long rows for {frequency}")

        # Exact match first.
        exact = df[df["STAT_time"].map(self._same_clock(clock))]
        if not exact.empty:
            # Prefer rows marked for auto recommendation.
            if "Use_for_auto_recommendation" in exact.columns:
                yes = exact[
                    exact["Use_for_auto_recommendation"]
                    .astype(str).str.lower().isin(["yes", "y", "true", "1"])
                ]
                if not yes.empty:
                    exact = yes
            return exact.iloc[0]

        raise ValueError(
            f"No exact STAT_time={clock} rule for {frequency}"
        )

    @staticmethod
    def _same_clock(target):
        target_min = _clock_to_minutes(target)

        def compare(x):
            try:
                return _clock_to_minutes(x) == target_min
            except Exception:
                return False

        return compare

    @staticmethod
    def _next_datetime(stat_dt, row):
        """
        Use the workbook's Next_round + Day_offset.
        24:00 is normalized to 00:00 after the indicated date offset.
        """
        next_round = _clock_to_minutes(row["Next_round"])
        day_offset = int(row["Day_offset"]) if not pd.isna(row["Day_offset"]) else 0

        # If Next_day is Yes but Day_offset is blank/zero, use +1.
        next_day = str(row["Next_day"]).strip().lower() == "yes"
        if next_day and day_offset == 0:
            day_offset = 1

        if next_round is None:
            raise ValueError("Next_round is blank for this Master_Long row")

        base_date = stat_dt.date() + timedelta(days=day_offset)

        if next_round == 1440:
            # 24:00 of base_date = 00:00 of the following date.
            return datetime.combine(
                base_date + timedelta(days=1),
                datetime.min.time()
            )

        return datetime.combine(
            base_date,
            datetime.min.time()
        ) + timedelta(minutes=next_round)

    def calculate(self, stat_text, frequency, dose_count=5, round_set=None):
        """
        Return a DataFrame with Dose #1 onward.

        If round_set is omitted, the exact Master_Long row supplies it.
        """
        if dose_count < 1:
            raise ValueError("dose_count must be >= 1")

        stat_dt = _parse_datetime(stat_text)
        row = self.master_row(stat_dt, frequency)

        selected_round = (
            str(round_set).strip()
            if round_set
            else str(row["Round_set"]).strip()
        )

        # Dose #1 is always the actual STAT datetime.
        dates = [stat_dt]

        if dose_count >= 2:
            dose2 = self._next_datetime(stat_dt, row)
            dates.append(dose2)

        # Rules sheet says Q24H Dose #3 onward continues every 24 hours
        # from Dose #2. For other frequencies, use the Round_Sets interval.
        if frequency == "Q24H":
            interval = timedelta(hours=24)
        else:
            rs = self.round_sets.copy()
            rs["Frequency"] = rs["Frequency"].astype(str).str.strip()
            rs["Round_set"] = rs["Round_set"].astype(str).str.strip()

            match = rs[rs["Round_set"].eq(selected_round)]
            if match.empty:
                raise ValueError(
                    f"Round_set {selected_round!r} not found in Round_Sets"
                )

            interval_hours = float(match.iloc[0]["Interval_hours"])
            interval = timedelta(hours=interval_hours)

        while len(dates) < dose_count:
            dates.append(dates[-1] + interval)

        return pd.DataFrame(
            {
                "Dose": range(1, dose_count + 1),
                "Date": [d.strftime("%d/%m/%Y") for d in dates],
                "Time": [d.strftime("%H:%M") for d in dates],
                "DateTime": dates,
            }
        ), selected_round


if __name__ == "__main__":
    # First validation case:
    # STAT 23/09/2026 23:00, Q24H.
    # The current workbook says:
    #   Next_round = 24
    #   Next_day = Yes
    #   Day_offset = 1
    #
    # Therefore:
    #   Dose #1 = 23/09 23:00
    #   Dose #2 = 25/09 00:00
    #   Dose #3 = 26/09 00:00
    #   ...

    workbook = "round_master_Q24H_fixed.xlsx"
    calc = MedicationRoundCalculator(workbook)

    result, selected_round = calc.calculate(
        "23/09/2026 23:00",
        "Q24H",
        dose_count=5,
    )

    print(f"Round set: {selected_round}")
    print(result.to_string(index=False))
