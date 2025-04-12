import os
import time
import csv
import tkinter as tk
import traceback
from tkinter import ttk, messagebox, filedialog, Checkbutton, IntVar, StringVar
from threading import Thread, Event
from datetime import datetime
import subprocess
import sys
import threading

# Try multiple approaches for Edge WebDriver
try:
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (StaleElementReferenceException,
                                          TimeoutException,
                                          ElementClickInterceptedException,
                                          WebDriverException,
                                          NoSuchElementException)

    # Try to import WebDriver Manager (if installed)
    try:
        from webdriver_manager.microsoft import EdgeChromiumDriverManager
        webdriver_manager_available = True
    except ImportError:
        webdriver_manager_available = False

except ImportError as e:
    print(f"Error importing Selenium packages: {e}")
    print("Please install required packages: pip install selenium webdriver-manager")
    sys.exit(1)

import logging  # Import the logging module

class PatentProcessor:
    def __init__(self, config_dir="config", log_function=print):  # Added log_function
        self.config_dir = config_dir
        os.makedirs(self.config_dir, exist_ok=True)
        self.patent_headers_default_path = os.path.join(self.config_dir, 'patent_headers_default.csv')
        self.patent_column_settings_path = os.path.join(self.config_dir, 'patent_column_settings.csv')
        self.patent_filter_settings_path = os.path.join(self.config_dir, 'patent_filter_settings.csv')
        self.selected_columns = {}
        self.filter_settings = []
        self.default_headers = []
        self.log_message = log_function  # Use the provided log function
        self.load_default_headers()
        self.generate_default_column_settings()
        self.generate_default_filter_settings()
        self.retrieve_column_settings()
        self.retrieve_filter_settings()

    def load_default_headers(self):
        """Load default headers from patent_headers_default.csv"""
        self.default_headers = []
        try:
            with open(self.patent_headers_default_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                self.default_headers = [row[0] for row in reader]  # Assuming headers are in the first column
            self.log_message("Loaded default headers from CSV", "STATUS")
        except FileNotFoundError:
            self.log_message("patent_headers_default.csv not found.  Using hardcoded defaults.", "WARNING")
            # Hardcoded default headers if the file is missing
            self.default_headers = ["Application Number", "Title", "Applicant", "Filing Date", "Abstract"]
        except Exception as e:
            self.log_message(f"Error loading default headers: {str(e)}", "ERROR")
            self.default_headers = ["Application Number", "Title", "Applicant", "Filing Date", "Abstract"]

    def generate_default_column_settings(self):
        """Generate patent_column_settings.csv if it doesn't exist."""
        if not os.path.exists(self.patent_column_settings_path):
            try:
                with open(self.patent_column_settings_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["HeaderName", "Include"])  # Write header row
                    for header in self.default_headers:
                        writer.writerow([header, "TRUE"])  # Include all columns by default
                self.log_message("Generated default patent_column_settings.csv", "STATUS")
            except Exception as e:
                self.log_message(f"Error generating default patent_column_settings.csv: {str(e)}", "ERROR")

    def generate_default_filter_settings(self):
        """Generate patent_filter_settings.csv if it doesn't exist."""
        if not os.path.exists(self.patent_filter_settings_path):
            try:
                with open(self.patent_filter_settings_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["FilterField", "MatchOperator", "MatchValue"])  # Write header row
                    # Add some default filter settings (example)
                    writer.writerow(["Status", "equals", "New"])
                    writer.writerow(["Applicant", "contains", "IBM"])
                self.log_message("Generated default patent_filter_settings.csv", "STATUS")
            except Exception as e:
                self.log_message(f"Error generating default patent_filter_settings.csv: {str(e)}", "ERROR")

    def retrieve_filter_settings(self):
        """Retrieve filter settings from patent_filter_settings.csv"""
        try:
            with open(self.patent_filter_settings_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader, None)  # Skip header row
                self.filter_settings = []  # Clear existing settings
                for row in reader:
                    filter_field, match_operator, match_value = row
                    self.filter_settings.append((filter_field, match_operator.strip(), match_value)) # Strip whitespace
                self.log_message("Retrieved filter settings from CSV", "STATUS")
        except FileNotFoundError:
            self.log_message("patent_filter_settings.csv not found. Using defaults.", "WARNING")
            self.generate_default_filter_settings()
            self.retrieve_filter_settings() # Retry loading
        except Exception as e:
            self.log_message(f"Failed to retrieve filter settings from CSV: {str(e)}", "ERROR")
            self.filter_settings = []  # Use empty list in case of an error

    def retrieve_column_settings(self):
        """Retrieve column settings from patent_column_settings.csv"""
        try:
            with open(self.patent_column_settings_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader, None)  # Skip header row
                self.selected_columns = {}
                for row in reader:
                    column_name, include_str = row
                    self.selected_columns[column_name] = include_str.upper() == 'TRUE'
                self.log_message("Retrieved column settings from CSV", "STATUS")
        except FileNotFoundError:
            self.log_message("patent_column_settings.csv not found. Using defaults.", "WARNING")
            self.generate_default_column_settings()
            self.retrieve_column_settings() # Retry loading
        except Exception as e:
            self.log_message(f"Failed to retrieve column settings from CSV: {str(e)}", "ERROR")
            self.selected_columns = {}  # Default to including all columns

    def filter_data(self, data):
        """Filter the data based on the settings in self.filter_settings"""
        filtered_data = []
        for row in data:
            include_row = True
            for filter_field, match_operator, match_value in self.filter_settings:
                try:
                    value = row[filter_field]

                    if match_operator == "equals" or match_operator == "=":
                        if str(value) != str(match_value):  # Convert to strings for comparison
                            include_row = False
                            break
                    elif match_operator == "contains":
                        if str(match_value) not in str(value):
                            include_row = False
                            break
                    elif match_operator == "--":  # Does Not Contain
                        if str(match_value) in str(value):
                            include_row = False
                            break
                    elif match_operator == "greater_than" or match_operator == ">":
                        try:
                            if not (float(value) > float(match_value)):  # Convert to numbers for comparison
                                include_row = False
                                break
                        except ValueError:
                            self.log_message(f"Cannot compare '{value}' and '{match_value}' as numbers", "WARNING")
                            include_row = False
                            break
                    elif match_operator == "less_than" or match_operator == "<":
                        try:
                            if not (float(value) < float(match_value)):  # Convert to numbers for comparison
                                include_row = False
                                break
                        except ValueError:
                            self.log_message(f"Cannot compare '{value}' and '{match_value}' as numbers", "WARNING")
                            include_row = False
                            break
                    elif match_operator == "starts with" or match_operator == "STARTS WITH":
                        if not str(value).startswith(str(match_value)):
                            include_row = False
                            break
                    elif match_operator == "ends with" or match_operator == "ENDS WITH":
                        if not str(value).endswith(str(match_value)):
                            include_row = False
                            break
                    else:
                        self.log_message(f"Unknown match operator: {match_operator}", "WARNING")
                        include_row = False  # Default to exclude if operator is unknown

                except KeyError:
                    self.log_message(f"Filter field '{filter_field}' not found in data", "WARNING")
                    include_row = False
                    break  # Skip to the next row if a filter field is missing
                except TypeError as e:
                    self.log_message(f"Type Error during comparison: {str(e)}", "WARNING")
                    include_row = False
                    break
            if include_row:
                filtered_data.append(row)
        return filtered_data

    def write_filtered_data_to_csv(self, filtered_data, output_csv_path, fieldnames):
        """Write filtered data to a new CSV file"""
        try:
            with open(output_csv_path, 'w', newline='', encoding='utf-8') as outfile:
                writer = csv.DictWriter(outfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(filtered_data)
            self.log_message(f"Filtered data written to {output_csv_path}", "SUCCESS")
        except Exception as e:
            self.log_message(f"Error writing filtered data to CSV: {str(e)}", "ERROR")


class PatentAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Patent Application Automation")
        self.root.geometry("800x600")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.running = False
        self.stop_event = threading.Event()
        self.current_process = None
        self.driver = None  # Initialize driver to None
        self.settings_loaded = False

        # Initialize variables
        self.email_var = StringVar()
        self.folder_var = StringVar()
        self.application_numbers_text = None
        self.extra_emails_text = None  # Initialize extra_emails_text
        self.column_vars = {}
        self.table_headers = ["Application Number", "Email"]  # Basic headers
        self.ad_headers = []  # To store application detail headers
        self.filter_settings = {}

        # Configure grid layout
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Create UI elements
        self.create_widgets()

        # Initialize PatentProcessor - pass the log_message function
        self.patent_processor = PatentProcessor(log_function=self.log_message)

        # Load default headers AFTER log_text is initialized
        self.patent_processor.load_default_headers()

        # Load user settings
        self.load_user_settings()

        # Retrieve filter settings
        self.retrieve_filter_settings()

    def create_widgets(self):
        # Frame for input fields
        input_frame = ttk.Frame(self.root, padding=10)
        input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Email input
        email_label = ttk.Label(input_frame, text="Email:")
        email_label.grid(row=0, column=0, sticky=tk.W)
        email_entry = ttk.Entry(input_frame, textvariable=self.email_var, width=30)
        email_entry.grid(row=0, column=1, sticky=tk.W)

        # Folder selection
        folder_label = ttk.Label(input_frame, text="Save Folder:")
        folder_label.grid(row=1, column=0, sticky=tk.W)
        folder_entry = ttk.Entry(input_frame, textvariable=self.folder_var, width=30)
        folder_entry.grid(row=1, column=1, sticky=tk.W)
        folder_button = ttk.Button(input_frame, text="Browse", command=self.browse_folder)
        folder_button.grid(row=1, column=2, sticky=tk.W)

        # Application Numbers input
        app_numbers_label = ttk.Label(input_frame, text="Application Numbers (one per line):")
        app_numbers_label.grid(row=2, column=0, sticky=tk.W)
        self.application_numbers_text = Text(input_frame, height=5, width=40)
        self.application_numbers_text.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E))

        # Extra Emails input
        extra_emails_label = ttk.Label(input_frame, text="Extra Emails (one per line):")
        extra_emails_label.grid(row=3, column=0, sticky=tk.W)
        self.extra_emails_text = Text(input_frame, height=3, width=40)
        self.extra_emails_text.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E))

        # Buttons
        button_frame = ttk.Frame(self.root, padding=10)
        button_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))

        self.start_button = ttk.Button(button_frame, text="Start", command=self.start_process)
        self.start_button.grid(row=0, column=0, padx=5, sticky=tk.W)

        self.stop_button = ttk.Button(button_frame, text="Stop", command=self.stop_process, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=5, sticky=tk.W)

        self.column_settings_button = ttk.Button(button_frame, text="Column Settings", command=self.show_column_settings)
        self.column_settings_button.grid(row=0, column=2, padx=5, sticky=tk.W)

        self.filter_button = ttk.Button(button_frame, text="Filter Settings", command=self.show_filter_settings)
        self.filter_button.grid(row=0, column=3, padx=5, sticky=tk.W)

        self.generate_filtered_button = ttk.Button(button_frame, text="Generate Filtered Data", command=self.generate_filtered_data, state=tk.DISABLED)
        self.generate_filtered_button.grid(row=0, column=4, padx=5, sticky=tk.W)

        # Log Text Area
        self.log_text = Text(self.root, height=10, wrap=tk.WORD)
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.config(state=tk.DISABLED)

        # Scrollbar for Log Text Area
        log_scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.log_text['yscrollcommand'] = log_scrollbar.set

    def set_default_values(self):
        """Set default values for email and folder"""
        self.email_var.set("default@example.com")
        self.folder_var.set(os.path.join(os.path.expanduser("~"), "Documents"))

    def log_message(self, message, tag="STATUS"):
        """Log messages to the text widget with different tags"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"{timestamp} - {message}\n"
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, formatted_message, tag)
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)

    def start_processing(self):
        """Start the data processing thread"""
        if not self.running:
            self.running = True
            self.stop_event.clear()
            self.fetch_applications_btn.config(state=tk.DISABLED)
            email = self.email_var.get()
            folder = self.folder_var.get()
            app_numbers_text = self.app_numbers_text.get("1.0", tk.END).strip()
            self.application_numbers = [num.strip() for num in app_numbers_text.replace("\n", ",").split(",") if num.strip()]

            if not email or not folder or not self.application_numbers:
                messagebox.showerror("Error", "Please fill in all fields.")
                self.running = False
                self.fetch_applications_btn.config(state=tk.NORMAL)
                return

            # Create the output folder if it does not exist
            os.makedirs(folder, exist_ok=True)

            # Initialize the CSV file
            csv_file = os.path.join(os.path.dirname(__file__), 'config', 'patent_full_list.csv')
            try:
                self.initialize_csv_file(csv_file)
            except Exception as e:
                messagebox.showerror("Error", f"Could not initialize CSV file: {str(e)}")
                self.running = False
                self.fetch_applications_btn.config(state=tk.NORMAL)
                return

            self.log_message("Starting data processing...", "STATUS")
            self.current_process = Thread(target=self.full_workflow)
            self.current_process.start()
        else:
            self.log_message("Process already running.", "WARNING")

    def setup_ui(self):
        self.root.title("Patent Automation Suite v3.6 (Edge)")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)

        # Create styles for buttons
        style = ttk.Style()
        style.configure("Fetch.TButton", font=("Helvetica", 10, "bold"), foreground="green")  # Bold green font for Fetch button
        style.configure("Exit.TButton", font=("Helvetica", 10, "bold"), foreground="red")  # Red font for Exit button

        # Main container
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configuration Section
        config_frame = ttk.LabelFrame(main_frame, text="Configuration")
        config_frame.pack(fill=tk.X, pady=5)

        # Email Input
        ttk.Label(config_frame, text="Email Address:").grid(row=0, column=0, sticky=tk.W, padx=5)
        email_entry = ttk.Entry(config_frame, textvariable=self.email_var, width=50)
        email_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)

        # Extra Emails Input
        ttk.Label(config_frame, text="Extra Emails (comma-separated or one per row):").grid(row=1, column=0, sticky=tk.W, padx=5)
        extra_emails_text = tk.Text(config_frame, height=4, width=50)
        extra_emails_text.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        self.extra_emails_text = extra_emails_text

        # Folder Selection
        ttk.Label(config_frame, text="Output Folder:").grid(row=2, column=0, sticky=tk.W, padx=5)
        folder_entry = ttk.Entry(config_frame, textvariable=self.folder_var)
        folder_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Button(config_frame, text="Browse", command=self.select_folder).grid(row=2, column=2, padx=5)

        # Column Selection Button
        select_columns_btn = ttk.Button(config_frame, text="Select Columns", command=self.show_column_selection, style="Select.TButton")
        select_columns_btn.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)

        # Add application numbers input field
        ttk.Label(config_frame, text="Application Numbers (comma-separated or one per row):").grid(row=4, column=0, sticky=tk.W, padx=5)
        self.app_numbers_text = tk.Text(config_frame, height=4)
        self.app_numbers_text.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)

        # Add Filter Settings button
        select_filters_btn = ttk.Button(config_frame, text="Filter Settings", command=self.show_filter_settings, style="Select.TButton")
        select_filters_btn.grid(row=5, column=1, sticky=tk.W, padx=5, pady=5)

        # Control Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        # Fetch applications button
        self.fetch_applications_btn = ttk.Button(btn_frame, text="Fetch applications", command=self.start_processing, style="Fetch.TButton")
        self.fetch_applications_btn.pack(side=tk.LEFT, padx=5)

        # Generate Filtered Data button
        self.generate_filtered_data_btn = ttk.Button(btn_frame, text="Generate Filtered Data", command=self.generate_filtered_data, style="Generate.TButton")
        self.generate_filtered_data_btn.pack(side=tk.LEFT, padx=5)

        # Reset to Default Settings button
        self.reset_btn = ttk.Button(main_frame, text="Reset to Default Settings", command=self.reset_to_default, style="Reset.TButton")
        self.reset_btn.pack(pady=5)

        # Exit button
        self.exit_btn = ttk.Button(main_frame, text="Exit", command=self.stop_process, style="Exit.TButton")
        self.exit_btn.pack(pady=5)

        # Log Console
        log_frame = ttk.LabelFrame(main_frame, text="Execution Log")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, state=tk.DISABLED, font=('Consolas', 10))
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Configure log tags
        self.log_text.tag_config("STATUS", foreground="blue")
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("WARNING", foreground="orange")
        self.log_text.tag_config("ERROR", foreground="red")

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Configure log tags
        self.log_text.tag_config("STATUS", foreground="blue")
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("WARNING", foreground="orange")
        self.log_text.tag_config("ERROR", foreground="red")

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def initialize_csv_file(self, csv_file):
        """Initialize CSV file with headers including email column"""
        try:
            # Get selected columns
            selected_columns = self.get_selected_columns()
            selected_ad_columns = self.get_selected_ad_columns()

            with open(csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)

                # Write headers for all columns without filtering
                headers = ["Email", "Application Number"] + self.table_headers + self.ad_headers
                headers = list(set(headers))

                writer.writerow(headers)

            self.log_message(f"Created initial CSV file: {csv_file}", "SUCCESS")

        except Exception as e:
            self.log_message(f"Error initializing CSV: {str(e)}", "ERROR")
            raise

    def update_csv_with_data(self, csv_path, email, app_number, headers_col1, headers_col3, data_col1, data_col3):
        """Update CSV with both basic and application details data"""
        try:
            # Read existing CSV data
            existing_data = []
            headers = []

            with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader)  # Get headers
                existing_data = list(reader)

            # Get selected columns
            selected_basic_columns = self.get_selected_columns()
            selected_ad_columns = self.get_selected_ad_columns()

            # Prepare row data
            row_data = {
                "Email": email,
                "Application Number": app_number
            }

            # Add all basic column data without filtering
            for i, header in enumerate(self.table_headers):
                # Check if the index is within the range of basic_row
                # and if the header is a selected basic column
                row_data[header] = ''  # Default value
                #row_data[header] = basic_row[i] if i < len(basic_row) and header in selected_basic_columns else ''

            # Add application detail data
            # Add column 1 data
            for header, value in zip(headers_col1, data_col1):
                row_data[f"{header}_AD"] = value

            # Add column 3 data
            for header, value in zip(headers_col3, data_col3):
                row_data[f"{header}_AD"] = value

            # Find existing row or prepare new row
            row_index = -1
            for i, row in enumerate(existing_data):
                if row[0] == email and row[1] == str(app_number):
                    row_index = i
                    break

            # Create new row or update existing row
            if row_index == -1:
                new_row = []
                for header in headers:
                    new_row.append(row_data.get(header, ''))
                existing_data.append(new_row)
            else:
                # Update existing row
                for i, header in enumerate(headers):
                    if header in row_data:
                        existing_data[row_index][i] = row_data[header]

            # Write updated data back to CSV
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(existing_data)

            self.log_message(f"Successfully updated data for {app_number} (Email: {email})", "SUCCESS")

        except Exception as e:
            self.log_message(f"Error updating CSV: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

    def open_output_file(self, file_path):
        """Open the output CSV file with the default application"""
        try:
            if sys.platform.startswith('darwin'):  # macOS
                subprocess.Popen(['open', file_path])
            elif os.name == 'nt':  # Windows
                os.startfile(file_path)
            elif os.name == 'posix':  # Linux
                subprocess.Popen(['xdg-open', file_path])
            self.log_message(f"Opening file: {file_path}", "STATUS")
        except Exception as e:
            self.log_message(f"Error opening file: {str(e)}", "ERROR")

    def select_folder(self):
        """Open a folder selection dialog and update the folder variable"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_var.set(folder_selected)

    def show_column_selection(self):
        """Display a window for selecting columns to include in the output"""
        column_selection_window = tk.Toplevel(self.root)
        column_selection_window.title("Select Columns")

        # Basic Columns Frame
        basic_columns_frame = ttk.LabelFrame(column_selection_window, text="Basic Columns")
        basic_columns_frame.pack(padx=10, pady=5, fill=tk.X)

        # Application Detail Columns Frame
        ad_columns_frame = ttk.LabelFrame(column_selection_window, text="Application Detail Columns")
        ad_columns_frame.pack(padx=10, pady=5, fill=tk.X)

        # Create checkbuttons for basic columns
        for header in self.table_headers:
            var = self.column_vars.get(header)
            if var is None:
                var = IntVar(value=1)  # Default to selected if not found
                self.column_vars[header] = var
            chk = Checkbutton(basic_columns_frame, text=header, variable=var)
            chk.pack(anchor=tk.W)

        # Create checkbuttons for application detail columns
        for header in self.ad_headers:
            var = self.column_vars.get(header)
            if var is None:
                var = IntVar(value=1)  # Default to selected if not found
                self.column_vars[header] = var
            chk = Checkbutton(ad_columns_frame, text=header, variable=var)
            chk.pack(anchor=tk.W)

        # Buttons to apply and cancel
        button_frame = ttk.Frame(column_selection_window)
        button_frame.pack(pady=5)

        apply_btn = ttk.Button(button_frame, text="Apply", command=self.save_column_selections)
        apply_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = ttk.Button(button_frame, text="Cancel", command=column_selection_window.destroy)
        cancel_btn.pack(side=tk.LEFT, padx=5)

    def get_selected_columns(self):
        """Get the list of selected basic columns based on the checkbuttons"""
        selected_columns = [header for header in self.table_headers if self.column_vars[header].get() == 1]
        return selected_columns

    def get_selected_ad_columns(self):
        """Get the list of selected application detail columns based on the checkbuttons"""
        selected_columns = [header for header in self.ad_headers if self.column_vars[header].get() == 1]
        return selected_columns

    def save_column_selections(self):
        """Save column selections to patent_column_settings.csv"""
        try:
            csv_path = os.path.join(os.path.dirname(__file__), 'config', 'patent_column_settings.csv')
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["HeaderName", "Include"])  # Write header
                for header, var in self.column_vars.items():
                    writer.writerow([header, "TRUE" if var.get() == 1 else "FALSE"])
            self.log_message("Saved column selections to CSV", "STATUS")
            messagebox.showinfo("Success", "Column selections saved successfully.")
        except Exception as e:
            self.log_message(f"Error saving column selections: {str(e)}", "ERROR")
            messagebox.showerror("Error", f"Error saving column selections: {str(e)}")

    def load_user_settings(self):
        """Load user settings from user_settings.csv"""
        try:
            csv_path = os.path.join(os.path.dirname(__file__), 'config', 'user_settings.csv')
            with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                settings = dict(reader)
                self.email_var.set(settings.get('email', "default@example.com"))
                self.folder_var.set(settings.get('folder', os.path.join(os.path.expanduser("~"), "Documents")))
                extra_emails = settings.get('extra_emails', '')
                self.extra_emails_text.delete("1.0", tk.END)
                self.extra_emails_text.insert(tk.END, extra_emails)
            self.settings_loaded = True
            self.log_message("Loaded user settings from CSV", "STATUS")
        except FileNotFoundError:
            self.log_message("user_settings.csv not found. Using defaults.", "WARNING")
            self.set_default_values()
        except Exception as e:
            self.log_message(f"Failed to load user settings from CSV: {str(e)}", "ERROR")
            self.set_default_values()
        finally:
            if not self.settings_loaded:
                self.set_default_values()

    def save_user_settings(self):
        """Save user settings to user_settings.csv"""
        try:
            csv_path = os.path.join(os.path.dirname(__file__), 'config', 'user_settings.csv')
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                settings = {
                    'email': self.email_var.get(),
                    'folder': self.folder_var.get(),
                    'extra_emails': self.extra_emails_text.get("1.0", tk.END).strip()
                }
                for key, value in settings.items():
                    writer.writerow([key, value])
            self.log_message("Saved user settings to CSV", "STATUS")
        except Exception as e:
            self.log_message(f"Error saving user settings: {str(e)}", "ERROR")

    def reset_to_default(self):
        """Reset all settings to default values"""
        try:
            # Reset email and folder
            self.set_default_values()

            # Reset column selections to default (all selected)
            for var in self.column_vars.values():
                var.set(1)

            # Clear extra emails
            self.extra_emails_text.delete("1.0", tk.END)

            # Save the reset settings
            self.save_user_settings()

            self.log_message("Settings reset to default", "STATUS")
            messagebox.showinfo("Reset", "Settings have been reset to default.")

        except Exception as e:
            self.log_message(f"Error resetting to default: {str(e)}", "ERROR")
            messagebox.showerror("Error", f"Error resetting settings: {str(e)}")

    def stop_process(self):
        """Stop the running process"""
        if self.running:
            self.running = False
            self.stop_event.set()
            if self.current_process and self.current_process.is_alive():
                self.log_message("Stopping current process...", "WARNING")
                # Attempt to join the thread (give it a chance to stop gracefully)
                self.current_process.join(timeout=5)
                if self.current_process.is_alive():
                    self.log_message("Process did not stop gracefully.  May need to manually terminate.", "ERROR")

            if self.driver:
                try:
                    self.driver.quit()
                    self.log_message("WebDriver closed.", "STATUS")
                except Exception as e:
                    self.log_message(f"Error closing WebDriver: {str(e)}", "ERROR")

            self.fetch_applications_btn.config(state=tk.NORMAL)
            self.log_message("Process stopped.", "STATUS")
        else:
            self.root.destroy()

    def on_close(self):
        """Handle window close event"""
        self.stop_process()
        self.root.destroy()

    def full_workflow(self):
        """Complete automation workflow"""
        try:
            email = self.email_var.get()
            folder = self.folder_var.get()
            csv_file = os.path.join(os.path.dirname(__file__), 'config', 'patent_full_list.csv')

            # Get extra emails from the text widget
            extra_emails_str = self.extra_emails_text.get("1.0", tk.END).strip()
            extra_emails = [e.strip() for e in extra_emails_str.replace("\n", ",").split(",") if e.strip()]
            all_emails = [email] + extra_emails

            # Launch Edge browser
            self.launch_edge()

            if not self.driver:
                self.log_message("Failed to launch Edge.  Aborting.", "ERROR")
                self.cleanup_after_completion()
                return

            for app_number in self.application_numbers:
                if self.stop_event.is_set():
                    break

                try:
                    # Fetch data from the website
                    headers_col1, headers_col3, data_col1, data_col3 = self.fetch_data(app_number)

                    # Update CSV for each email
                    for current_email in all_emails:
                        self.update_csv_with_data(csv_file, current_email, app_number, headers_col1, headers_col3, data_col1, data_col3)

                except Exception as e:
                    self.log_message(f"Error processing application {app_number}: {str(e)}", "ERROR")
                    self.log_message(traceback.format_exc(), "ERROR")

            # Open the output file
            self.open_output_file(csv_file)

        except Exception as e:
            self.log_message(f"An error occurred during the workflow: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

        finally:
            self.cleanup_after_completion()

    def launch_edge(self):
        """Launch Edge browser with Selenium"""
        try:
            self.log_message("Launching Edge browser...", "STATUS")

            edge_options = Options()
            edge_options.add_argument("--headless=new")  # Run in headless mode
            edge_options.add_argument("--disable-gpu")  # Disable GPU acceleration
            edge_options.add_argument("--no-sandbox")  # For running in environments like Docker
            edge_options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
            edge_options.add_argument("--ignore-certificate-errors") # Ignore certificate errors
            edge_options.add_argument("--disable-extensions") # Disable extensions

            # Try WebDriver Manager first
            if webdriver_manager_available:
                try:
                    edge_service = Service(EdgeChromiumDriverManager().install())
                    self.driver = webdriver.Edge(service=edge_service, options=edge_options)
                    self.log_message("Edge launched successfully using WebDriver Manager", "SUCCESS")
                    return
                except Exception as e:
                    self.log_message(f"WebDriver Manager failed: {str(e)}", "WARNING")
                    self.log_message("Falling back to system PATH...", "WARNING")

            # Fallback to system PATH
            try:
                self.driver = webdriver.Edge(options=edge_options)
                self.log_message("Edge launched successfully using system PATH", "SUCCESS")
            except Exception as e:
                self.log_message(f"Failed to launch Edge: {str(e)}", "ERROR")
                self.driver = None  # Ensure driver is None if launch fails

        except Exception as e:
            self.log_message(f"Error launching Edge: {str(e)}", "ERROR")
            self.driver = None

    def fetch_data(self, app_number):
        """Fetch data from the website for a given application number"""
        try:
            self.log_message(f"Fetching data for application number: {app_number}", "STATUS")

            # Construct the URL
            url = f"https://www.ic.gc.ca/opic-cipo/cpd/eng/patent/{app_number}/summary"
            self.driver.get(url)

            # Wait for the page to load and elements to be present
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "wb-dtlst"))
            )

            # Extract data from column 1
            col1 = self.driver.find_element(By.XPATH, '//*[@id="wb-dtlst"]/dl[1]')
            data_pairs_col1 = [item.text for item in col1.find_elements(By.XPATH, './child::*')]
            headers_col1 = data_pairs_col1[::2]
            data_col1 = data_pairs_col1[1::2]

            # Extract data from column 3
            col3 = self.driver.find_element(By.XPATH, '//*[@id="wb-dtlst"]/dl[3]')
            data_pairs_col3 = [item.text for item in col3.find_elements(By.XPATH, './child::*')]
            headers_col3 = data_pairs_col3[::2]
            data_col3 = data_pairs_col3[1::2]

            self.log_message(f"Data fetched successfully for application number: {app_number}", "SUCCESS")
            return headers_col1, headers_col3, data_col1, data_col3

        except TimeoutException as e:
            self.log_message(f"Timeout while waiting for element: {str(e)}", "ERROR")
            return [], [], [], []  # Return empty lists in case of timeout

        except NoSuchElementException as e:
            self.log_message(f"Element not found: {str(e)}", "ERROR")
            return [], [], [], []  # Return empty lists if element is not found

        except Exception as e:
            self.log_message(f"Error fetching data: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")
            return [], [], [], []  # Return empty lists in case of error

    def cleanup_after_completion(self):
        """Clean up resources after the process completes"""
        try:
            if self.driver:
                self.driver.quit()
                self.log_message("WebDriver closed.", "STATUS")
            self.running = False
            self.fetch_applications_btn.config(state=tk.NORMAL)
            self.log_message("Process completed.", "STATUS")
        except Exception as e:
            self.log_message(f"Error during cleanup: {str(e)}", "ERROR")

    def retrieve_filter_settings(self):
        """Retrieve filter settings from the patent_filter_settings.csv file"""
        try:
            csv_path = os.path.join(os.path.dirname(__file__), 'config', 'patent_filter_settings.csv')
            with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                # Use PatentProcessor to retrieve filter settings
                self.patent_processor.retrieve_filter_settings()
            self.log_message("Retrieved filter settings from CSV", "STATUS")
        except Exception as e:
            self.log_message(f"Failed to retrieve filter settings from CSV: {str(e)}", "ERROR")
            # Set default filter settings
            self.filter_settings = {
                "Action": "Action Required",
                "Chapter": "4",
                "Status": "New"
            }

    def generate_filtered_data(self):
        """Generate a new CSV file with filtered data based on settings"""
        try:
            # Define paths
            input_csv_path = os.path.join(os.path.dirname(__file__), 'config', 'patent_full_list.csv')
            output_csv_path = os.path.join(self.folder_var.get(), 'patent_filtered_list.csv')

            # Read data from the input CSV
            with open(input_csv_path, 'r', newline='', encoding='utf-8') as infile:
                reader = csv.DictReader(infile)
                data = list(reader)

            # Filter data using PatentProcessor
            filtered_data = self.patent_processor.filter_data(data)

            # Write filtered data to the output CSV using PatentProcessor
            fieldnames = reader.fieldnames
            self.patent_processor.write_filtered_data_to_csv(filtered_data, output_csv_path, fieldnames)

            # Open the output file
            self.open_output_file(output_csv_path)

        except Exception as e:
            self.log_message(f"Error generating filtered data: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

    def show_filter_settings(self):
        """Display a window for viewing and editing filter settings."""
        filter_settings_window = tk.Toplevel(self.root)
        filter_settings_window.title("Filter Settings")

        # Frame to hold filter settings
        filter_frame = ttk.Frame(filter_settings_window)
        filter_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        # Display filter settings (basic example - needs improvement)
        for i, (filter_field, match_operator, match_value) in enumerate(self.patent_processor.filter_settings):
            ttk.Label(filter_frame, text=f"{filter_field}: {match_operator} {match_value}").grid(row=i, column=0, sticky=tk.W)

        # Add buttons to apply and cancel (and potentially add/edit filters)
        button_frame = ttk.Frame(filter_settings_window)
        button_frame.pack(pady=5)

        apply_btn = ttk.Button(button_frame, text="Apply", command=filter_settings_window.destroy)
        apply_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = ttk.Button(button_frame, text="Cancel", command=filter_settings_window.destroy)
        cancel_btn.pack(side=tk.LEFT, padx=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = PatentAutomationApp(root)
    root.mainloop()