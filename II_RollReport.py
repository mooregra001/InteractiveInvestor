import os
import logging
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from datetime import datetime, timedelta
import shutil
import json
from II_Constants import DEFAULT_BASE_PATH, DEFAULT_EXCEL_PATH
from II_Config import load_config  # Import load_config from II_Config.py

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("roll_report.log", mode="a"),
        logging.StreamHandler()
    ]
)

def update_config(new_file_path, config_path="config.json"):
    """Update the configuration file with the new Excel file path."""
    config = load_config(config_path)
    config["excel_path"] = new_file_path
    try:
        with open(config_path, "w") as f:
            json.dump(config, f, indent=4)
        logging.info(f"Updated config with new file path: {new_file_path}")
    except Exception as e:
        logging.error(f"Failed to update config file {config_path}: {e}")
        raise RuntimeError(f"Failed to update config file: {e}")

def validate_date(date_str):
    """Validate date string in MM/DD/YYYY format."""
    try:
        datetime.strptime(date_str, "%m/%d/%Y")
        return True
    except ValueError:
        logging.warning(f"Invalid date format: {date_str}. Expected MM/DD/YYYY.")
        return False

def roll_to_next_business_day(workbook_path, business_date_str):
    """Roll the report to the specified business date and create folder structure."""
    logging.info(f"Starting roll_to_next_business_day for workbook: {workbook_path}")
    
    # Extract date from file name (e.g., II_20250917.xlsx -> 20250917)
    try:
        current_date_str = workbook_path[-12:-5]  # Adjusted to capture YYYYMMDD
        current_date = datetime.strptime(current_date_str, "%Y%m%d")
    except ValueError:
        logging.error(f"Invalid date format in file name: {current_date_str}")
        raise ValueError(f"Invalid date format in file name: {current_date_str}")
    
    # Validate business date
    try:
        business_date = datetime.strptime(business_date_str, "%m/%d/%Y")
    except ValueError:
        logging.error(f"Invalid business date: {business_date_str}")
        raise ValueError(f"Invalid business date: {business_date_str}")
    
    # Load config for base path
    config = load_config()
    base_path = config.get("base_path", DEFAULT_BASE_PATH)
    base_path = os.path.normpath(base_path.rstrip(os.sep))
    base_path = os.path.join(base_path, '')
    
    if not os.path.exists(base_path):
        logging.error(f"Base path does not exist: {base_path}")
        messagebox.showerror("Error", f"Base path does not exist: {base_path}")
        raise ValueError(f"Base path does not exist: {base_path}")
    
    # Create folder structure (YYYY/MM.MMM)
    year_folder = business_date.strftime("%Y")
    month_folder = business_date.strftime("%m.%b")
    folder_path = os.path.join(base_path, year_folder, month_folder)
    os.makedirs(folder_path, exist_ok=True)
    logging.info(f"Created folder structure: {folder_path}")
    
    # Generate new file name
    new_date_str = business_date.strftime("%Y%m%d")
    new_file_name = f"II_{new_date_str}.xlsx"
    new_file_path = os.path.join(folder_path, new_file_name)
    logging.info(f"New file path: {new_file_path}")
    
    # Copy the original file to the new file path
    try:
        shutil.copy(workbook_path, new_file_path)
        logging.info(f"Successfully copied {workbook_path} to {new_file_path}")
        # Update config.json with the new file path
        update_config(new_file_path)
    except Exception as e:
        logging.error(f"Failed to copy file: {e}")
        raise RuntimeError(f"Failed to copy file to {new_file_path}: {e}")
    
    return new_file_path

def main():
    """Main function to test rolling the report."""
    logging.info("Starting roll report test")
    root = tk.Tk()
    root.withdraw()

    # Prompt user to select the Excel workbook
    workbook_path = filedialog.askopenfilename(
        title="Select Excel Workbook",
        filetypes=[("Excel files", "*.xlsx *.xlsm")]
    )
    if not workbook_path:
        logging.error("No workbook selected")
        messagebox.showerror("Error", "No workbook selected.")
        root.destroy()
        return

    # Prompt user for business date
    today = datetime.now()
    today_weekday = today.weekday()
    days_back = 3 if today_weekday == 0 else 1  # Monday: go back 3 days
    default_date = (today - timedelta(days=days_back)).strftime("%m/%d/%Y")
    business_date = simpledialog.askstring(
        "Input",
        "Please enter business date (MM/DD/YYYY)",
        initialvalue=default_date
    )
    root.destroy()

    if not business_date:
        logging.error("No business date provided")
        messagebox.showerror("Error", "No business date provided.")
        return

    if not validate_date(business_date):
        logging.error("Invalid business date format")
        messagebox.showerror("Error", "Invalid date format. Use MM/DD/YYYY.")
        return

    try:
        # Roll to new file path
        new_excel_path = roll_to_next_business_day(workbook_path, business_date)
        logging.info(f"Successfully generated new file path: {new_excel_path}")
        messagebox.showinfo("Success", f"New file created: {new_excel_path}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {str(e)}")
        messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

if __name__ == "__main__":
    main()