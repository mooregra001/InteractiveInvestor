# II_RollReport.py
import os
import logging
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from datetime import datetime, timedelta
import shutil

# Import centralized config and constants
from II_Config import load_config, update_config
from II_Constants import DEFAULT_BASE_PATH

# ----------------------------------------------------------------------
# LOGGING
# ----------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("roll_report.log", mode="a", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# ----------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------
def validate_date(date_str: str) -> bool:
    """Return True if date_str matches MM/DD/YYYY."""
    try:
        datetime.strptime(date_str, "%m/%d/%Y")
        return True
    except ValueError:
        logging.warning(f"Invalid date format: {date_str}")
        return False


def extract_date_from_filename(filepath: str) -> str:
    """
    Expected filename format: II_YYYYMMDD.xlsx
    Returns the YYYYMMDD string or raises ValueError.
    """
    filename = os.path.basename(filepath)
    if not (filename.startswith("II_") and filename.lower().endswith(".xlsx")):
        raise ValueError("Filename must be II_YYYYMMDD.xlsx")

    date_part = filename[3:11]  # II_YYYYMMDD.xlsx → positions 3–10
    if len(date_part) != 8 or not date_part.isdigit():
        raise ValueError(f"Date portion '{date_part}' is not YYYYMMDD")

    try:
        datetime.strptime(date_part, "%Y%m%d")
    except ValueError as exc:
        raise ValueError(f"Invalid date in filename: {date_part}") from exc

    return date_part


# ----------------------------------------------------------------------
# CORE ROLL FUNCTION
# ----------------------------------------------------------------------
def roll_to_next_business_day(workbook_path: str, business_date_str: str) -> str:
    """Copy the workbook to a new file named after the given business date."""
    logging.info(f"Rolling workbook: {workbook_path}")

    # 1. Validate source file
    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"Source workbook not found: {workbook_path}")

    # 2. Extract current date from filename
    try:
        current_date_str = extract_date_from_filename(workbook_path)
        current_date = datetime.strptime(current_date_str, "%Y%m%d")
        logging.info(f"Current file date: {current_date_str}")
    except Exception as e:
        logging.error(f"Cannot parse date from filename: {e}")
        raise

    # 3. Parse target business date
    if not validate_date(business_date_str):
        raise ValueError("Business date must be MM/DD/YYYY")
    business_date = datetime.strptime(business_date_str, "%m/%d/%Y")
    new_date_str = business_date.strftime("%Y%m%d")

    # 4. Validate target > source
    if business_date <= current_date:
        raise ValueError("Target business date must be after the source file date.")

    # 5. Resolve base folder from config
    config = load_config()
    base_path = os.path.normpath(config.get("base_path", DEFAULT_BASE_PATH))
    if not os.path.isdir(base_path):
        raise ValueError(f"Base path does not exist: {base_path}")

    # 6. Build destination path
    new_filename = f"II_{new_date_str}.xlsx"
    new_file_path = os.path.join(base_path, new_filename)
    logging.info(f"Target file: {new_file_path}")

    # 7. Prevent overwrite without confirmation
    if os.path.exists(new_file_path):
        root = tk.Tk()
        root.withdraw()
        confirm = messagebox.askyesno("Overwrite?", f"File already exists:\n{new_file_path}\n\nOverwrite?")
        root.destroy()
        if not confirm:
            raise RuntimeError("Operation cancelled by user.")

    # 8. Copy file
    try:
        shutil.copy2(workbook_path, new_file_path)
        logging.info(f"Copied to {new_file_path}")
    except Exception as e:
        logging.error(f"Copy failed: {e}")
        raise RuntimeError(f"Failed to copy file: {e}")

    # 9. Update config with new path
    update_config(new_file_path)

    return new_file_path


# ----------------------------------------------------------------------
# MAIN / GUI
# ----------------------------------------------------------------------
def main():
    logging.info("=== Roll-Report utility started ===")
    root = tk.Tk()
    root.withdraw()

    # Pre-fill with last used file if available
    config = load_config()
    initial_dir = None
    initial_file = None
    if config.get("excel_path") and os.path.exists(config["excel_path"]):
        initial_dir = os.path.dirname(config["excel_path"])
        initial_file = os.path.basename(config["excel_path"])

    # Select workbook
    workbook_path = filedialog.askopenfilename(
        title="Select the II_*.xlsx workbook",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=initial_dir,
        initialfile=initial_file
    )
    if not workbook_path:
        messagebox.showerror("Cancelled", "No file selected – exiting.")
        root.destroy()
        return

    # Default business date (previous business day)
    today = datetime.now()
    weekday = today.weekday()
    days_back = 3 if weekday == 0 else 1  # Monday → Friday, else yesterday
    default_date = (today - timedelta(days=days_back)).strftime("%m/%d/%Y")

    business_date = simpledialog.askstring(
        "Business Date",
        "Enter business date (MM/DD/YYYY):",
        initialvalue=default_date
    )
    root.destroy()

    if not business_date:
        messagebox.showerror("Cancelled", "No date entered – exiting.")
        return

    if not validate_date(business_date):
        messagebox.showerror("Invalid Date", "Use MM/DD/YYYY format.")
        return

    # Perform the roll
    try:
        new_path = roll_to_next_business_day(workbook_path, business_date)
        messagebox.showinfo(
            "Success",
            f"New report created:\n{new_path}"
        )
        logging.info("Roll completed successfully.")
    except Exception as exc:
        msg = str(exc)
        logging.error(f"Roll failed: {msg}")
        messagebox.showerror("Error", msg)


if __name__ == "__main__":
    main()