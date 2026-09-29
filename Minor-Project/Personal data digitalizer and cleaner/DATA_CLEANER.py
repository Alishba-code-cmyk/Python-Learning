"""
Reads my Daily Activity Tracker (.xlsx) with openpyxl, validates it,
stores each day as an object, calculates indices + correlations,
saves cleaned data to JSON and offers a small menu.
"""
import json
from datetime import datetime, timedelta

import openpyxl

# 1. SETTINGS  

EXCEL_FILE = "12621146.xlsx"          
JSON_FILE = "12621146_cleaned.json"   

# The program finds each column by checking what the header text

REQUIRED_COLUMNS = ["date", "sleep", "fitness", "study", "coding", "class",
                    "attended", "other", "total", "free", "day",
                    "satisfact", "energy"]
OPTIONAL_COLUMNS = ["notes"]
# Words are turned into numbers  so correlations can be calculated.
SATISFACTION_SCALE = {"very satisfied": 5, "satisfied": 4, "neutral":3,
                      "unsatisfied": 2, "very unsatisfied": 1}
ENERGY_SCALE = {"low": 1, "medium": 2, "high": 3 }

# 2. CUSTOM EXCEPTION 
 
class InvalidDataError(Exception):
   pass

# 3. CLASS 

class DailyRecord:
    """One row of the Excel sheet = one DailyRecord object."""

    def __init__(self, date, sleep, fitness, study, coding, class_min,
                 attended, other, feeling, satisfaction_text, satisfaction,
                 energy_text, energy, notes):
        # A leading underscore means "private - don't touch from outside"
        self._date = date
        self._sleep = sleep
        self._fitness = fitness
        self._study = study
        self._coding = coding
        self._class = class_min
        self._attended = attended         
        self._other = other
        self._feeling = feeling
        self._satisfaction_text = satisfaction_text   # original word
        self._satisfaction = satisfaction             # number 1-5
        self._energy_text = energy_text
        self._energy = energy
        self._notes = notes


    def get_date(self):
        return self._date

    def get_sleep(self):
        return self._sleep

    def get_fitness(self):
        return self._fitness

    def get_study(self):
        return self._study

    def get_coding(self):
        return self._coding

    def get_class(self):
        return self._class

    def get_attended(self):
        return self._attended

    def get_other(self):
        return self._other

    def get_feeling(self):
        return self._feeling

    def get_satisfaction(self):
        return self._satisfaction

    def get_satisfaction_text(self):
        return self._satisfaction_text

    def get_energy(self):
        return self._energy

    def get_notes(self):
        return self._notes

    def total_minutes(self):
        return (self._sleep + self._fitness + self._study + self._coding
                + self._class + self._other)

    def free_minutes(self):
        return 1440 - self.total_minutes()

    def to_dict(self):
        """Convert the object to a dictionary so it can be saved as JSON."""
        return {
            "date": self._date.isoformat(),
            "sleep": self._sleep,
            "fitness": self._fitness,
            "study": self._study,
            "coding": self._coding,
            "class": self._class,
            "attended": self._attended,
            "other": self._other,
            "total_tracked": self.total_minutes(),
            "free_time": self.free_minutes(),
            "feeling": self._feeling,
            "satisfaction": self._satisfaction_text,
            "satisfaction_score": self._satisfaction,
            "energy": self._energy_text,
            "energy_score": self._energy,
            "notes": self._notes,
        }

    def __str__(self):
        return (f"{self._date} | Sleep {self._sleep:.0f} | Study {self._study:.0f} | "
                f"Coding {self._coding:.0f} | Class {self._class:.0f} | "
                f"{self._feeling} | {self._satisfaction_text} | Energy {self._energy_text}")

# 4. SMALL HELPER FUNCTIONS THAT CHECK ONE VALUE EACH
#    (each raises InvalidDataError if something is wrong)
def parse_date(value):
    """Return a date object from an Excel date cell ."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, str):
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass            # try the next format
    raise InvalidDataError(f"invalid or missing date: {value!r}")


def parse_minutes(value, name):
    """Minutes must be a number between 0 and 1440 (minutes in a day)."""
    if value is None or str(value).strip() == "":
        raise InvalidDataError(f"{name} is missing")
    try:
        minutes = float(value)
    except (TypeError, ValueError):
        raise InvalidDataError(f"{name} is not a number: {value!r}")
    if minutes < 0 or minutes > 1440:
        raise InvalidDataError(f"{name} out of range (0-1440): {minutes}")
    return minutes


def parse_scale(value, table, name):
    """Turn a rating into a number 1-5. Accepts numbers or words."""
    if value is None or str(value).strip() == "":
        raise InvalidDataError(f"{name} is missing")
    try:                                   # case 1: already a number
        number = float(value)
    except ValueError:                     # case 2: it is a word
        key = str(value).strip().lower()
        if key in table:
            return table[key]
        raise InvalidDataError(f"unknown {name} value: {value!r}")
    if 1 <= number <= 5:
        return number
    raise InvalidDataError(f"{name} out of range (1-5): {number}")

# 5. READING THE EXCEL FILE  (openpyxl)

def find_header_row(all_rows):
    """Return the index of the row whose cells include a header starting with 'date'."""
    for index, row in enumerate(all_rows):
        for cell in row:
            if cell is not None and str(cell).strip().lower().startswith("date"):
                return index
    raise InvalidDataError("Could not find a header row containing 'DATE'")


def find_columns(header_row):
    """Return a dictionary like {'date': 0, 'sleep': 1, ...} (column positions)."""
    headers = [str(h).strip().lower() if h is not None else "" for h in header_row]
    positions = {}
    for keyword in REQUIRED_COLUMNS + OPTIONAL_COLUMNS:
        for index, header in enumerate(headers):
            if header.startswith(keyword):
                positions[keyword] = index
                break
        else:   # for-else: runs only if 'break' never happened
            if keyword in REQUIRED_COLUMNS:
                raise InvalidDataError(f"No column starting with '{keyword}' found in the header row")
    return positions


def build_record(row, cols):
    """Check one Excel row and return a DailyRecord (or raise InvalidDataError)."""
    date = parse_date(row[cols["date"]])
    sleep = parse_minutes(row[cols["sleep"]], "Sleep")
    fitness = parse_minutes(row[cols["fitness"]], "Fitness")
    study = parse_minutes(row[cols["study"]], "Study")
    coding = parse_minutes(row[cols["coding"]], "Coding")
    class_min = parse_minutes(row[cols["class"]], "Class")
    other = parse_minutes(row[cols["other"]], "Other Activities")

    attended = row[cols["attended"]]
    try:
        attended = float(attended)
        if attended < 0:
            raise ValueError
    except (TypeError, ValueError):
        raise InvalidDataError(f"Attended must be a number 0 or more: {attended!r}")

    total = sleep + fitness + study + coding + class_min + other
    if total > 1440:
        raise InvalidDataError(f"total minutes ({total:.0f}) exceed 1440")

    # Cross-check with the TOTAL TRACKED and FREE TIME columns in the sheet
    sheet_total = parse_minutes(row[cols["total"]], "Total Tracked")
    sheet_free = parse_minutes(row[cols["free"]], "Free Time")
    if abs(sheet_total - total) > 0.5:
        raise InvalidDataError(f"Total Tracked says {sheet_total:.0f} but activities add up to {total:.0f}")
    if abs(sheet_free - (1440 - total)) > 0.5:
        raise InvalidDataError(f"Free Time says {sheet_free:.0f} but should be {1440 - total:.0f}")

    feeling = row[cols["day"]]
    if feeling is None or str(feeling).strip() == "":
        raise InvalidDataError("Day's Feeling is missing")
    feeling = str(feeling).strip()

    sat_raw = row[cols["satisfact"]]
    satisfaction = parse_scale(sat_raw, SATISFACTION_SCALE, "Satisfaction")
    en_raw = row[cols["energy"]]
    energy = parse_scale(en_raw, ENERGY_SCALE, "Energy")

    notes = ""
    if "notes" in cols and row[cols["notes"]] is not None:
        notes = str(row[cols["notes"]]).strip()

    return DailyRecord(date, sleep, fitness, study, coding, class_min,
                       attended, other, feeling, str(sat_raw).strip(),
                       satisfaction, str(en_raw).strip(), energy, notes)


def load_data(path):
    """Read the whole sheet. Returns (list_of_records, list_of_skipped_rows)."""
    try:
        workbook = openpyxl.load_workbook(path, data_only=True)
    except FileNotFoundError:
        print(f"ERROR: file '{path}' not found. Keep it in the same folder as this .py file.")
        return [], []
    except Exception as error:        # any other openpyxl problem (corrupt file etc.)
        print("ERROR: could not open the Excel file:", error)
        return [], []

    sheet = workbook.active
    all_rows = list(sheet.iter_rows(values_only=True))   # each row = tuple of cell values

    try:
        header_index = find_header_row(all_rows)
        cols = find_columns(all_rows[header_index])
    except InvalidDataError as error:
        print("ERROR:", error)
        return [], []

    records = []        # list -> good rows as DailyRecord objects
    skipped = []        # list -> (excel_row_number, reason) for bad rows
    seen_dates = set()  # set  -> to catch duplicate dates

    # index 0 = Excel row 1, so Excel row number = index + 1
    for index in range(header_index + 1, len(all_rows)):
        row = all_rows[index]
        row_number = index + 1
        if all(cell is None for cell in row):
            continue                                    # blank row, ignore
        try:
            record = build_record(row, cols)
            if record.get_date() in seen_dates:
                raise InvalidDataError("duplicate date")
            seen_dates.add(record.get_date())
            records.append(record)
        except InvalidDataError as error:
            skipped.append((row_number, str(error)))

    records.sort(key=lambda r: r.get_date())     # keep the date sequence correct
    return records, skipped


# ------------------------------------------------------------------
# 6. STATISTICS FUNCTIONS
# ------------------------------------------------------------------
def average(numbers):
    return sum(numbers) / len(numbers) if numbers else 0


def pearson(x_list, y_list):
    """Correlation coefficient (-1 to +1) written by hand, no numpy."""
    n = len(x_list)
    if n < 2:
        return 0
    mean_x, mean_y = average(x_list), average(y_list)
    top = sum((x - mean_x) * (y - mean_y) for x, y in zip(x_list, y_list))
    bottom_x = sum((x - mean_x) ** 2 for x in x_list) ** 0.5
    bottom_y = sum((y - mean_y) ** 2 for y in y_list) ** 0.5
    if bottom_x == 0 or bottom_y == 0:
        return 0            # one column never changes -> no correlation possible
    return top / (bottom_x * bottom_y)


def count_feelings(records):
    """Dictionary: feeling -> number of days."""
    counts = {}
    for r in records:
        counts[r.get_feeling()] = counts.get(r.get_feeling(), 0) + 1
    return counts


def longest_streak(records, condition):
    """Longest run of CONSECUTIVE calendar days where condition(record) is True."""
    best = current = 0
    previous_date = None
    for r in records:
        consecutive = previous_date is not None and r.get_date() - previous_date == timedelta(days=1)
        if condition(r):
            current = current + 1 if consecutive and current > 0 else 1
            best = max(best, current)
        else:
            current = 0
        previous_date = r.get_date()
    return best


# 7. THE INDICES 

def calculate_indices(records):
    if not records:
        return {}
    day = 1440   # minutes in a day

    avg_sleep = average([r.get_sleep() for r in records])
    avg_fit = average([r.get_fitness() for r in records])
    avg_code = average([r.get_coding() for r in records])
    avg_study = average([r.get_study() for r in records])
    avg_class = average([r.get_class() for r in records])
    avg_other = average([r.get_other() for r in records])
    avg_total = average([r.total_minutes() for r in records])
    avg_energy = average([r.get_energy() for r in records])
    avg_sat = average([r.get_satisfaction() for r in records])

    tpi = avg_code / day * 100                       # Tech Productivity
    aai = (avg_study + avg_class) / day * 100        # Academic Activity
    phai = avg_fit / day * 100                       # Physical Activity
    sri = min(avg_sleep / 480 * 100, 100)            # Sleep & Recovery (8 h = 100)
    groups = (avg_code, avg_study + avg_class, avg_fit, avg_other)   # tuple
    abi = (min(groups) / max(groups) * 100) if max(groups) > 0 else 0  # Activity Balance
    tui = avg_total / day * 100                      # Time Utilization
    ei = ((avg_energy + avg_sat) / 2) / 5 * 100      # Experience Index

    # DCI = Valid Recorded Days / Expected Days x 100  (from your slide)
    first, last = records[0].get_date(), records[-1].get_date()
    expected_days = (last - first).days + 1
    dci = len(records) / expected_days * 100

    pai = average([tpi, aai, phai, sri, abi, tui, ei])   # overall index

    return {"PAI": pai, "TPI": tpi, "AAI": aai, "PhAI": phai, "SRI": sri,
            "ABI": abi, "TUI": tui, "EI": ei, "DCI": dci}


def calculate_correlations(records):
    return {
        "Coding <-> Energy": pearson([r.get_coding() for r in records],
                                     [r.get_energy() for r in records]),
        "Sleep <-> Energy": pearson([r.get_sleep() for r in records],
                                    [r.get_energy() for r in records]),
        "Study <-> Satisfaction": pearson([r.get_study() for r in records],
                                          [r.get_satisfaction() for r in records]),
    }


def describe(value):
    """Words for a correlation number."""
    strength = abs(value)
    if strength >= 0.6:
        word = "strong"
    elif strength >= 0.3:
        word = "moderate"
    else:
        word = "weak / almost no"
    direction = "positive" if value > 0 else "negative"
    return f"{word} {direction} relationship"

# 8. MENU OPTIONS 
def show_summary(records):
    print(f"\nDays loaded: {len(records)}  ({records[0].get_date()} to {records[-1].get_date()})")
    print(f"Average sleep    : {average([r.get_sleep() for r in records]):.0f} min")
    print(f"Average study    : {average([r.get_study() for r in records]):.0f} min")
    print(f"Average coding   : {average([r.get_coding() for r in records]):.0f} min")
    print(f"Average class    : {average([r.get_class() for r in records]):.0f} min")
    print(f"Average fitness  : {average([r.get_fitness() for r in records]):.0f} min")
    print(f"Average other activity  : {average([r.get_other() for r in records]):.0f} min")
    print(f"Average free time: {average([r.free_minutes() for r in records]):.0f} min")
    print(f"Lectures attended: {sum(r.get_attended() for r in records):.0f} in total")
    best = max(records, key=lambda r: r.get_coding())      # max/min give a record
    worst = min(records, key=lambda r: r.get_coding())
    extremes = (best.get_date(), worst.get_date())         # tuple
    print(f"Most coding on {extremes[0]} ({best.get_coding():.0f} min), "
          f"least on {extremes[1]} ({worst.get_coding():.0f} min)")
    print("Days per feeling:", count_feelings(records))
    print("Unique feelings :", set(r.get_feeling() for r in records))
    print("Longest streak of 60+ min coding days:",
          longest_streak(records, lambda r: r.get_coding() >= 60))


def search_by_field(records, label, getter):
    text = input(f"Enter {label} to search (e.g. Good): ").strip().lower()
    found = [r for r in records if str(getter(r)).lower() == text]
    print(f"\n{len(found)} day(s) found")
    for r in found:
        print(" ", r)
        if r.get_notes():
            print("     Note:", r.get_notes())


def show_indices(records):
    print()
    for name, value in calculate_indices(records).items():
        print(f"{name:5}: {value:6.2f}")


def show_correlations(records):
    print()
    for name, value in calculate_correlations(records).items():
        print(f"{name:24}: r = {value:+.2f}  ->  {describe(value)}")


def show_findings(records):
    corr = calculate_correlations(records)
    print("\nMy findings (based on my own data):")
    for name, value in corr.items():
        print(f"- {name}: {describe(value)} (r = {value:+.2f})")
    idx = calculate_indices(records)
    print(f"- My data continuity is {idx['DCI']:.0f}% and my sleep index is {idx['SRI']:.0f}/100.")


def save_json(records, skipped):
    output = {
        "records": [r.to_dict() for r in records],
        "skipped_rows": [{"excel_row": n, "reason": why} for n, why in skipped],
        "indices": {k: round(v, 2) for k, v in calculate_indices(records).items()},
        "correlations": {k: round(v, 3) for k, v in calculate_correlations(records).items()},
    }
    try:
        with open(JSON_FILE, "w") as file:
            json.dump(output, file, indent=4)
        print(f"\nSaved {len(records)} records to {JSON_FILE}")
    except OSError as error:
        print("Could not write JSON file:", error)


def main():
    records, skipped = load_data(EXCEL_FILE)
    if not records:
        print("No valid data to work with. Program stopped.")
        return

    print(f"Loaded {len(records)} valid rows, skipped {len(skipped)} bad rows.")
    for row_number, reason in skipped:
        print(f"  Skipped Excel row {row_number}: {reason}")

    save_json(records, skipped)      # cleaned file is written automatically

    while True:
        print("\n===== MENU =====")
        print("1. Summary")
        print("2. Show days by Day's Feeling (e.g. Good)")
        print("3. Show days by Satisfaction (e.g. Very Satisfied)")
        print("4. Show indices")
        print("5. Show correlations")
        print("6. Show my findings")
        print("7. Save JSON again")
        print("0. Exit")
        choice = input("Choose: ").strip()

        if choice == "1":
            show_summary(records)
        elif choice == "2":
            search_by_field(records, "feeling", lambda r: r.get_feeling())
        elif choice == "3":
            search_by_field(records, "satisfaction", lambda r: r.get_satisfaction_text())
        elif choice == "4":
            show_indices(records)
        elif choice == "5":
            show_correlations(records)
        elif choice == "6":
            show_findings(records)
        elif choice == "7":
            save_json(records, skipped)
        elif choice == "0":
            print("Goodbye!")
            break
        else:
            print("Please type a number from the menu.")


# Runs main() only when this file is executed directly
if __name__ == "__main__":
    main()
