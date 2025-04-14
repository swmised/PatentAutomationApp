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


class PatentAutomationApp:
    def __init__(self, root):
        self.root = root
        self.driver = None
        self.running = False
        self.stop_event = Event()
        self.application_numbers = []
        self.csv_data = []
        self.current_process = None

        self.column_vars = {}      # Add this
        self.ad_column_vars = {}   # Add this
        self.filter_settings = {}  # Add this
		
        self.email_var = tk.StringVar()
        self.folder_var = tk.StringVar()

        # Initialize UI components FIRST
        self.initialize_styles()
        self.setup_ui()

        # Retrieve headers from CSV file
        self.retrieve_headers_from_csv()

        # Retrieve filter settings from CSV file
        self.retrieve_filter_settings()

        # Initialize column selection variables for default headers
        for header in self.table_headers:
            self.column_vars[header] = IntVar(value=1)  # Default to selected
        for header in self.ad_headers:
            self.column_vars[header] = IntVar(value=1)  # Default to selected

        # Set default values if settings were not loaded
        if not hasattr(self, 'settings_loaded') or not self.settings_loaded:
            self.set_default_values()
    def initialize_styles(self):
        """Initialize custom UI styles"""
        self.style = ttk.Style()
        self.style.configure('LogToggle.TButton', 
                            font=('Segoe UI', 9, 'bold'),
                            foreground='#444444')
        self.style.configure('TButton', padding=6)

    def setup_ui(self):
        self.root.title("Patent Automation Suite v3.6 (Edge)")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create log panel FIRST
        self.log_panel = SlidingLogPanel(self.root)

        # Then build other components
        self.build_configuration_section(main_frame)

        # Control Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        # Fetch applications button
        self.fetch_applications_btn = ttk.Button(
            btn_frame, text="Fetch applications", command=self.start_processing, style="Fetch.TButton")
        self.fetch_applications_btn.pack(side=tk.LEFT, padx=5)

        # Generate Filtered Data button
        self.generate_filtered_data_btn = ttk.Button(
            btn_frame, text="Generate Filtered Data", command=self.generate_filtered_data, style="Generate.TButton")
        self.generate_filtered_data_btn.pack(side=tk.LEFT, padx=5)

        # Reset to Default Settings button
        self.reset_btn = ttk.Button(main_frame, text="Reset to Default Settings",
                                    command=self.reset_to_default, style="Reset.TButton")
        self.reset_btn.pack(pady=5)

        # Exit button
        self.exit_btn = ttk.Button(
            main_frame, text="Exit", command=self.stop_process, style="Exit.TButton")
        self.exit_btn.pack(pady=5)

        # Log Console
        log_frame = ttk.LabelFrame(main_frame, text="Execution Log")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD,
                                state=tk.DISABLED, font=('Consolas', 10))
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

    def log_message(self, message, level="info"):
        """Add message to log console with timestamp"""
        if hasattr(self, 'log_panel'):
            self.log_panel.log(message, level)
        else:
            # Fallback for early initialization messages
            print(f"[FALLBACK LOG] {message}")

    def retrieve_headers_from_csv(self):
        """Retrieve headers from the patent_headers_default.csv file"""
        try:
            csv_path = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_headers_default.csv')
            with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader)

                # Separate basic headers and application detail headers
                self.table_headers = ["Application Number"]
                self.ad_headers = []

                for header in headers:
                    if header.endswith("_AD"):
                        self.ad_headers.append(header)
                    else:
                        if header != "Application Number":  # Application Number is always included
                            self.table_headers.append(header)

            self.log_message(
                f"Retrieved {len(self.table_headers) - 1} basic headers and {len(self.ad_headers)} application detail headers from CSV", "STATUS")
        except Exception as e:
            self.log_message(
                f"Failed to retrieve headers from CSV: {str(e)}", "ERROR")
            self.log_message("Using default headers", "WARNING")

            # Default headers to show initially
            self.table_headers = [
                "Application Number", "Action", "Chapter", "Due Date", "Received Date",
                "Status", "Email", "Comments", "Checkbox"
            ]

            # Default application detail headers
            self.ad_headers = [
                "Application Number_AD", "Filing Date_AD", "Status_AD", "Title_AD",
                "Applicant_AD", "Agent_AD", "Exceptor_AD", "Classification_AD"
            ]

    def set_default_values(self):
        """Set default values for email and folder."""
        self.email_var.set("simon.chau@ised-isde.gc.ca")
        self.folder_var.set(os.path.join(os.path.expanduser('~'), 'Documents'))

    def start_processing(self):
        """Start complete automation workflow"""
        if not self.running:
            self.running = True
            self.stop_event.clear()
            self.fetch_applications_btn.config(state=tk.DISABLED)

            # Save user settings
            self.save_user_settings()

            # Capture application numbers from the text widget
            app_numbers_str = self.app_numbers_text.get("1.0", tk.END).strip()
            self.application_numbers = [num.strip() for num in app_numbers_str.replace(
                "\n", ",").split(",") if num.strip()]

            # Define the path for the patent_full_list.csv file
            csv_file = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_full_list.csv')

            # Initialize the CSV file
            self.initialize_csv_file(csv_file)

            self.current_process = Thread(
                target=self.full_workflow, daemon=True)
            self.current_process.start()

    def setup_ui(self):
        self.root.title("Patent Automation Suite v3.6 (Edge)")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)

        # Create styles for buttons
        style = ttk.Style()
        style.configure("Fetch.TButton", font=("Helvetica", 10, "bold"),
                        foreground="green")  # Bold green font for Fetch button
        style.configure("Exit.TButton", font=("Helvetica", 10, "bold"),
                        foreground="red")  # Red font for Exit button

        # Main container
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configuration Section
        config_frame = ttk.LabelFrame(main_frame, text="Configuration")
        config_frame.pack(fill=tk.X, pady=5)

        # Email Input
        ttk.Label(config_frame, text="Email Address:").grid(
            row=0, column=0, sticky=tk.W, padx=5)
        email_entry = ttk.Entry(
            config_frame, textvariable=self.email_var, width=50)
        email_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)

        # Extra Emails Input
        ttk.Label(config_frame, text="Extra Emails (comma-separated or one per row):").grid(
            row=1, column=0, sticky=tk.W, padx=5)
        extra_emails_text = tk.Text(config_frame, height=4, width=50)
        extra_emails_text.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        self.extra_emails_text = extra_emails_text

        # Folder Selection
        ttk.Label(config_frame, text="Output Folder:").grid(
            row=2, column=0, sticky=tk.W, padx=5)
        folder_entry = ttk.Entry(config_frame, textvariable=self.folder_var)
        folder_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Button(config_frame, text="Browse", command=self.select_folder).grid(
            row=2, column=2, padx=5)

        # Column Selection Button
        select_columns_btn = ttk.Button(
            config_frame, text="Select Columns", command=self.show_column_selection, style="Select.TButton")
        select_columns_btn.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)

        # Add application numbers input field
        ttk.Label(config_frame, text="Application Numbers (comma-separated or one per row):").grid(
            row=4, column=0, sticky=tk.W, padx=5)
        self.app_numbers_text = tk.Text(config_frame, height=4)
        self.app_numbers_text.grid(
            row=4, column=1, sticky=tk.EW, padx=5, pady=2)

        # Control Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        # Fetch applications button
        self.fetch_applications_btn = ttk.Button(
            btn_frame, text="Fetch applications", command=self.start_processing, style="Fetch.TButton")
        self.fetch_applications_btn.pack(side=tk.LEFT, padx=5)

        # Generate Filtered Data button
        self.generate_filtered_data_btn = ttk.Button(
            btn_frame, text="Generate Filtered Data", command=self.generate_filtered_data, style="Generate.TButton")
        self.generate_filtered_data_btn.pack(side=tk.LEFT, padx=5)

        # Reset to Default Settings button
        self.reset_btn = ttk.Button(main_frame, text="Reset to Default Settings",
                                    command=self.reset_to_default, style="Reset.TButton")
        self.reset_btn.pack(pady=5)

        # Exit button
        self.exit_btn = ttk.Button(
            main_frame, text="Exit", command=self.stop_process, style="Exit.TButton")
        self.exit_btn.pack(pady=5)

        # Log Console
        log_frame = ttk.LabelFrame(main_frame, text="Execution Log")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD,
                                state=tk.DISABLED, font=('Consolas', 10))
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
                headers = ["Email", "Application Number"] + \
                    self.table_headers + self.ad_headers
                headers = list(set(headers))

                writer.writerow(headers)

            self.log_message(
                f"Created initial CSV file: {csv_file}", "SUCCESS")

        except Exception as e:
            self.log_message(f"Error initializing CSV: {str(e)}", "ERROR")
            raise

    def update_csv_with_data(self, csv_path, email, app_number, basic_row, headers_col1, headers_col3, data_col1, data_col3):
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
                if i < len(basic_row):
                    row_data[header] = basic_row[i]

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

            self.log_message(
                f"Successfully updated data for {app_number} (Email: {email})", "SUCCESS")

        except Exception as e:
            self.log_message(f"Error updating CSV: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

    def open_output_file(self, file_path):
        """Open the output CSV file with the default application"""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(file_path)
            elif sys.platform == 'darwin':  # macOS
                subprocess.call(['open', file_path])
            else:  # Linux
                subprocess.call(['xdg-open', file_path])
            self.log_message(f"Opened output file: {file_path}", "SUCCESS")
        except Exception as e:
            self.log_message(f"Failed to open output file: {str(e)}", "ERROR")

    def process_email_applications(self, email, csv_file):
        """Process applications for a specific email"""
        try:
            # Get selected columns
            selected_columns = self.get_selected_columns()
            selected_ad_columns = self.get_selected_ad_columns()

            max_retries = 2
            failed_applications = set()
            processed_numbers = set()

            for attempt in range(max_retries + 1):
                current_failed_applications = set()

                # Make a copy of application numbers to avoid modification during iteration
                current_applications = self.application_numbers.copy()

                for app_number in current_applications:
                    # Skip if already processed successfully
                    if app_number in processed_numbers:
                        continue

                    # Check if stop requested
                    if self.stop_event.is_set():
                        self.log_message("Process aborted by user", "WARNING")
                        return

                    try:
                        self.log_message(
                            f"Processing application {app_number} for {email} (Attempt {attempt + 1})", "STATUS")

                        # Process application details
                        details_data = self.get_application_details(app_number)

                        # Update CSV with application details including email
                        self.update_csv_with_email_data(
                            csv_file,
                            email,
                            app_number,
                            details_data
                        )

                        processed_numbers.add(app_number)
                        self.log_message(
                            f"Successfully processed {app_number} for {email}", "SUCCESS")

                    except Exception as e:
                        self.log_message(
                            f"Attempt {attempt + 1}: Error processing application {app_number} for {email}: {str(e)}", "ERROR")
                        current_failed_applications.add(app_number)

                # Update failed applications set
                if attempt < max_retries:
                    failed_applications = current_failed_applications

                    if not failed_applications:
                        break

                    self.application_numbers = list(failed_applications)
                    self.log_message(
                        f"Retry attempt {attempt + 2} for {len(self.application_numbers)} applications", "STATUS")
                else:
                    # Final attempt - log permanent failures
                    if current_failed_applications:
                        self.log_message(
                            f"\n--- PERMANENT FAILURES FOR {email} ---", "ERROR")
                        for app in current_failed_applications:
                            self.log_message(
                                f"Failed to process application {app}", "ERROR")

            self.log_message(
                f"Processed {len(processed_numbers)} applications for {email}", "SUCCESS")
            if failed_applications:
                self.log_message(
                    f"Failed to process {len(failed_applications)} applications for {email}", "WARNING")

        except Exception as e:
            self.log_message(
                f"Error processing applications for {email}: {str(e)}", "ERROR")
            raise

    def load_user_settings(self):
        """Load user settings from patent_settings_default.csv"""
        try:
            settings_path = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_settings_default.csv')
            self.log_message(
                f"Loading user settings from: {settings_path}", "INFO")
            if os.path.exists(settings_path):
                with open(settings_path, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    settings = {}
                    for row in reader:
                        if len(row) != 2:
                            self.log_message(
                                f"Skipping invalid row: {row}", "WARNING")
                            continue
                        settings[row[0]] = row[1]

                # Load settings into variables
                self.email_var.set(settings.get("email", ""))
                self.folder_var.set(settings.get(
                    "folder", os.path.join(os.path.expanduser('~'), 'Documents')))
                self.app_numbers_text.insert(
                    "1.0", settings.get("application_numbers", ""))
                self.extra_emails_text.insert(
                    "1.0", settings.get("extra_emails", ""))

                for header in self.table_headers:
                    try:
                        self.column_vars[header].set(
                            int(settings.get(header, "1")))
                    except ValueError:
                        self.log_message(
                            f"Skipping invalid entry for {header}", "WARNING")

                for header in self.ad_headers:
                    try:
                        self.column_vars[header].set(
                            int(settings.get(header, "1")))
                    except ValueError:
                        self.log_message(
                            f"Skipping invalid entry for {header}", "WARNING")

                self.settings_loaded = True
                self.log_message("User settings loaded successfully", "STATUS")
            else:
                self.log_message(
                    f"No user settings file found at: {settings_path}, using default settings", "WARNING")
                self.set_default_values()
        except Exception as e:
            self.log_message(
                f"Failed to load user settings: {str(e)}", "ERROR")
            self.set_default_values()

    def save_user_settings(self):
        """Save user settings to patent_settings_default.csv"""
        try:
            settings_path = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_settings_default.csv')
            self.log_message(
                f"Saving user settings to: {settings_path}", "INFO")
            with open(settings_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["email", self.email_var.get()])
                writer.writerow(["folder", self.folder_var.get()])
                writer.writerow(
                    ["application_numbers", self.app_numbers_text.get("1.0", tk.END).strip()])
                writer.writerow(
                    ["extra_emails", self.extra_emails_text.get("1.0", tk.END).strip()])

                for header, var in self.column_vars.items():
                    writer.writerow([header, var.get()])

                for header, var in self.ad_column_vars.items():
                    writer.writerow([header, var.get()])

            self.log_message(
                f"User settings saved successfully to {settings_path}", "DEBUG")
        except Exception as e:
            self.log_message(
                f"Failed to save user settings: {str(e)}", "ERROR")

    def reset_to_default(self):
        """Reset settings to the original default settings"""
        self.email_var.set("simon.chau@ised-isde.gc.ca")
        self.folder_var.set(os.path.join(os.path.expanduser('~'), 'Documents'))

        for header in self.table_headers:
            self.column_vars[header].set(1)

        for header in self.ad_headers:
            self.ad_column_vars[header].set(1)

        self.app_numbers_text.delete("1.0", tk.END)
        self.extra_emails_text.delete("1.0", tk.END)

        self.log_message("Settings reset to default", "STATUS")

    def show_column_selection(self):
        """Show column selection dialog with both regular and application detail columns"""
        # Create a new top-level window
        column_window = tk.Toplevel(self.root)
        column_window.title("Select Columns to Include")
        column_window.geometry("800x600")
        # Set as transient to the main window
        column_window.transient(self.root)
        column_window.grab_set()  # Make modal

        # Create a notebook with tabs
        notebook = ttk.Notebook(column_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab 1: Basic Columns
        basic_frame = ttk.Frame(notebook)
        notebook.add(basic_frame, text="Basic Columns")

        # Add a note about Application Number
        ttk.Label(basic_frame, text="Note: 'Application Number' is always included").pack(
            anchor=tk.W, padx=5, pady=5)

        # Create a scrollable frame for basic columns
        basic_canvas = tk.Canvas(basic_frame)
        basic_scrollbar = ttk.Scrollbar(
            basic_frame, orient="vertical", command=basic_canvas.yview)
        basic_scrollable_frame = ttk.Frame(basic_canvas)

        basic_scrollable_frame.bind(
            "<Configure>",
            lambda e: basic_canvas.configure(
                scrollregion=basic_canvas.bbox("all"))
        )

        basic_canvas.create_window(
            (0, 0), window=basic_scrollable_frame, anchor="nw")
        basic_canvas.configure(yscrollcommand=basic_scrollbar.set)

        basic_canvas.pack(side="left", fill="both", expand=True)
        basic_scrollbar.pack(side="right", fill="y")

        # Add checkboxes and filter settings for each basic column (excluding Application Number)
        for header in sorted(self.table_headers):
            if header != "Application Number":  # Application Number is always included
                if header not in self.column_vars:
                    self.column_vars[header] = IntVar(
                        value=0)  # Default to deselected
                if header not in self.filter_settings:
                    self.filter_settings[header] = {
                        'operator': "--",
                        'value': "",
                        'operator_var': StringVar(value="--"),
                        'value_var': StringVar(value="")
                    }

                frame = ttk.Frame(basic_scrollable_frame)
                frame.pack(anchor=tk.W, padx=20, pady=2)

                Checkbutton(frame, text=header, variable=self.column_vars[header]).pack(
                    side=tk.LEFT)

                operator_var = self.filter_settings[header]['operator_var']
                ttk.Combobox(frame, textvariable=operator_var, values=[
                             "--", "IS", "BLANK", "CONTAINS", "STARTS WITH", "ENDS WITH"], width=10).pack(side=tk.LEFT, padx=5)

                value_var = self.filter_settings[header]['value_var']
                entry = ttk.Entry(frame, textvariable=value_var, width=20)
                if "Date" in header:
                    entry.insert(0, "yyyy/mm/dd")
                    entry.config(foreground="gray")
                entry.pack(side=tk.LEFT, padx=5)

        # Add buttons for Select All / Deselect All for basic columns
        btn_frame1 = ttk.Frame(basic_frame)
        btn_frame1.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame1, text="Select All Basic",
                   command=self.select_all_basic_columns).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame1, text="Deselect All Basic",
                   command=self.deselect_all_basic_columns).pack(side=tk.LEFT, padx=5)

        # Tab 2: Application Detail Columns
        ad_frame = ttk.Frame(notebook)
        notebook.add(ad_frame, text="Application Detail Columns")

        # Add a note about Application Detail columns
        ttk.Label(ad_frame, text="Select Application Detail columns to include:").pack(
            anchor=tk.W, padx=5, pady=5)

        # Create a scrollable frame for AD columns
        ad_canvas = tk.Canvas(ad_frame)
        ad_scrollbar = ttk.Scrollbar(
            ad_frame, orient="vertical", command=ad_canvas.yview)
        ad_scrollable_frame = ttk.Frame(ad_canvas)

        ad_scrollable_frame.bind(
            "<Configure>",
            lambda e: ad_canvas.configure(scrollregion=ad_canvas.bbox("all"))
        )

        ad_canvas.create_window(
            (0, 0), window=ad_scrollable_frame, anchor="nw")
        ad_canvas.configure(yscrollcommand=ad_scrollbar.set)

        ad_canvas.pack(side="left", fill="both", expand=True)
        ad_scrollbar.pack(side="right", fill="y")

        # Add checkboxes and filter settings for each application detail column
        for header in sorted(self.ad_headers):
            if header not in self.ad_column_vars:
                self.ad_column_vars[header] = IntVar(
                    value=0)  # Default to deselected
            if header not in self.filter_settings:
                self.filter_settings[header] = {
                    'operator': "--",
                    'value': "",
                    'operator_var': StringVar(value="--"),
                    'value_var': StringVar(value="")
                }

            frame = ttk.Frame(ad_scrollable_frame)
            frame.pack(anchor=tk.W, padx=20, pady=2)

            display_name = header[:-3] if header.endswith("_AD") else header
            Checkbutton(frame, text=display_name,
                        variable=self.ad_column_vars[header]).pack(side=tk.LEFT)

            operator_var = self.filter_settings[header]['operator_var']
            ttk.Combobox(frame, textvariable=operator_var, values=[
                         "--", "IS", "BLANK", "CONTAINS", "STARTS WITH", "ENDS WITH"], width=10).pack(side=tk.LEFT, padx=5)

            value_var = self.filter_settings[header]['value_var']
            entry = ttk.Entry(frame, textvariable=value_var, width=20)
            if "Date" in header:
                entry.insert(0, "yyyy/mm/dd")
                entry.config(foreground="gray")
            entry.pack(side=tk.LEFT, padx=5)

        # Add buttons for Select All / Deselect All for AD columns
        btn_frame2 = ttk.Frame(ad_frame)
        btn_frame2.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame2, text="Select All Details",
                   command=self.select_all_ad_columns).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame2, text="Deselect All Details",
                   command=self.deselect_all_ad_columns).pack(side=tk.LEFT, padx=5)

        # Add OK and Reset buttons at the bottom of the window
        btn_frame3 = ttk.Frame(column_window)
        btn_frame3.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame3, text="OK", command=lambda: self.save_and_close_column_window(
            column_window)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame3, text="Reset", command=self.reset_filter_settings).pack(
            side=tk.LEFT, padx=5)

    def save_filter_settings(self):
        """Save filter settings to the patent_filters_default.csv file"""
        try:
            csv_path = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_filters_default.csv')
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                for header, settings in self.filter_settings.items():
                    writer.writerow(
                        [header, settings['operator'], settings['value']])
            self.log_message("Filter settings saved successfully", "STATUS")
        except Exception as e:
            self.log_message(
                f"Failed to save filter settings: {str(e)}", "ERROR")

    def save_and_close_column_window(self, window):
        """Save the column selections and filter settings, then close the window"""
        for header in self.table_headers:
            if header != "Application Number":  # Application Number is always included
                self.column_vars[header].set(self.column_vars[header].get())
                self.filter_settings[header]['operator'] = self.filter_settings[header]['operator_var'].get(
                )
                self.filter_settings[header]['value'] = self.filter_settings[header]['value_var'].get(
                )

        for header in self.ad_headers:
            self.ad_column_vars[header].set(self.ad_column_vars[header].get())
            self.filter_settings[header]['operator'] = self.filter_settings[header]['operator_var'].get(
            )
            self.filter_settings[header]['value'] = self.filter_settings[header]['value_var'].get(
            )

        self.save_filter_settings()
        window.destroy()

    def reset_filter_settings(self):
        """Reset filter settings to default values"""
        for header in self.table_headers:
            if header != "Application Number":  # Application Number is always included
                self.filter_settings[header] = {
                    'operator': "--", 'value': "", 'operator_var': StringVar(value="--"), 'value_var': StringVar(value="")}

        for header in self.ad_headers:
            self.filter_settings[header] = {
                'operator': "--", 'value': "", 'operator_var': StringVar(value="--"), 'value_var': StringVar(value="")}

        self.log_message("Filter settings reset to default", "STATUS")

    def save_and_close_column_window(self, window):
        """Save the column selections and filter settings, then close the window"""
        for header in self.table_headers:
            if header != "Application Number":  # Application Number is always included
                self.column_vars[header].set(self.column_vars[header].get())
                self.filter_settings[header]['operator'] = self.filter_settings[header]['operator_var'].get(
                )
                self.filter_settings[header]['value'] = self.filter_settings[header]['value_var'].get(
                )

        for header in self.ad_headers:
            self.ad_column_vars[header].set(self.ad_column_vars[header].get())
            self.filter_settings[header]['operator'] = self.filter_settings[header]['operator_var'].get(
            )
            self.filter_settings[header]['value'] = self.filter_settings[header]['value_var'].get(
            )

        self.save_filter_settings()
        window.destroy()

    def select_all_basic_columns(self):
        """Select all basic columns"""
        for column, var in self.column_vars.items():
            var.set(1)

    def deselect_all_basic_columns(self):
        """Deselect all basic columns"""
        for column, var in self.column_vars.items():
            var.set(0)

    def select_all_ad_columns(self):
        """Select all application detail columns"""
        for column, var in self.ad_column_vars.items():
            var.set(1)

    def deselect_all_ad_columns(self):
        """Deselect all application detail columns"""
        for column, var in self.ad_column_vars.items():
            var.set(0)

    def get_selected_columns(self):
        """Retrieve the list of selected columns from the GUI."""
        selected_columns = []
        for header, var in self.column_vars.items():
            if var.get() == 1:
                selected_columns.append(header)
        return selected_columns

    def get_selected_ad_columns(self):
        """Retrieve the list of selected application detail columns from the GUI."""
        selected_ad_columns = []
        for column, var in self.ad_column_vars.items():
            if var.get() == 1:
                selected_ad_columns.append(column)
        return selected_ad_columns

    def update_headers(self, basic_headers, ad_headers):
        """Update the list of all headers and create selection variables for new headers"""
        # Update basic headers
        for header in basic_headers:
            if header not in self.table_headers:
                self.table_headers.append(header)
                self.column_vars[header] = IntVar(
                    value=1)  # Default to selected

        # Update application detail headers
        for header in ad_headers:
            if header not in self.ad_headers:
                self.ad_headers.append(header)
                self.ad_column_vars[header] = IntVar(
                    value=1)  # Default to selected

    def update_csv_with_email_data(self, csv_file, email, app_number, details_data):
        """Update CSV with application details including email"""
        try:
            # Read existing CSV data
            existing_data = []
            headers = []

            with open(csv_file, 'r', newline='', encoding='utf-8') as f:
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
                if i < len(details_data['basic_row']):
                    row_data[header] = details_data['basic_row'][i]

            # Add application detail data
            for header, value in zip(details_data['headers_col1'], details_data['data_col1']):
                row_data[f"{header}_AD"] = value

            for header, value in zip(details_data['headers_col3'], details_data['data_col3']):
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
            with open(csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(existing_data)

            self.log_message(
                f"Successfully updated data for {app_number} (Email: {email})", "SUCCESS")

        except Exception as e:
            self.log_message(f"Error updating CSV: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

    def generate_filtered_data(self):
        """Generate filtered data based on user settings"""
        try:
            # Define the path for the patent_full_list.csv file
            input_file = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_full_list.csv')

            # Check if the file exists
            if not os.path.exists(input_file):
                self.log_message(f"File not found: {input_file}", "ERROR")
                return

            # Read the existing CSV data
            with open(input_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                rows = list(reader)

            # Apply filters
            filtered_rows = []
            for row in rows:
                include_row = True
                for header, settings in self.filter_settings.items():
                    operator = settings['operator']
                    value = settings['value'].strip()
                    cell_value = row.get(header, "")

                    if operator == "IS" and cell_value != value:
                        include_row = False
                    elif operator == "BLANK" and cell_value != "":
                        include_row = False
                    elif operator == "CONTAINS" and value not in cell_value:
                        include_row = False
                    elif operator == "STARTS WITH" and not cell_value.startswith(value):
                        include_row = False
                    elif operator == "ENDS WITH" and not cell_value.endswith(value):
                        include_row = False

                    if not include_row:
                        break

                if include_row:
                    filtered_rows.append(row)

            # Define the path for the filtered data file with a timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            results_folder = os.path.join(os.path.dirname(__file__), 'Results')
            os.makedirs(results_folder, exist_ok=True)
            filtered_file = os.path.join(
                results_folder, f'filtered_patent_list_{timestamp}.csv')

            # Get headers from patent_headers_default.csv
            headers_path = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_headers_default.csv')
            with open(headers_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader)

            # Filter out unselected columns
            selected_headers = [header for header in headers if self.column_vars.get(
                header, IntVar(value=0)).get() == 1]

            # Write the filtered data to a new CSV file
            with open(filtered_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=selected_headers)
                writer.writeheader()
                for row in filtered_rows:
                    filtered_row = {header: row[header]
                                    for header in selected_headers if header in row}
                    writer.writerow(filtered_row)

            self.log_message(
                f"Filtered data saved to {filtered_file}", "SUCCESS")
            self.open_output_file(filtered_file)
        except Exception as e:
            self.log_message(
                f"Failed to generate filtered data: {str(e)}", "ERROR")

    def retrieve_filter_settings(self):
        """Retrieve filter settings from the patent_filters_default.csv file"""
        try:
            csv_path = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_filters_default.csv')
            if os.path.exists(csv_path):
                with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if len(row) == 3:
                            header, operator, value = row
                            self.filter_settings[header] = {
                                'operator': operator,
                                'value': value,
                                'operator_var': StringVar(value=operator),
                                'value_var': StringVar(value=value)
                            }
                self.log_message(
                    "Filter settings loaded successfully", "STATUS")
            else:
                self.log_message(
                    "No filter settings file found, using default settings", "WARNING")
        except Exception as e:
            self.log_message(
                f"Failed to load filter settings: {str(e)}", "ERROR")
            self.log_message("Using default filter settings", "WARNING")

    def select_folder(self):
        selected_dir = filedialog.askdirectory(
            initialdir=self.folder_var.get(),
            title="Select Output Folder"
        )
        if selected_dir:
            self.folder_var.set(selected_dir)

    def kill_edge_processes(self):
        """Kill any running Edge processes that might interfere with automation"""
        self.log_message("Checking for running Edge processes...", "STATUS")
        try:
            if os.name == 'nt':  # Windows
                os.system('taskkill /f /im msedge.exe >nul 2>&1')
                os.system('taskkill /f /im msedgedriver.exe >nul 2>&1')
            else:  # Linux/Mac
                os.system('pkill -f msedge > /dev/null 2>&1')
                os.system('pkill -f msedgedriver > /dev/null 2>&1')
            time.sleep(2)  # Give processes time to terminate
            self.log_message("Edge processes terminated", "SUCCESS")
        except Exception as e:
            self.log_message(
                f"Error terminating Edge processes: {str(e)}", "WARNING")

    def full_workflow(self):
        try:
            # Initialize Edge browser
            self.log_message("Initializing Microsoft Edge...", "STATUS")

            # Configure Edge options
            edge_options = Options()
            edge_options.add_experimental_option("detach", True)
            edge_options.add_argument("--disable-notifications")
            edge_options.add_argument("--start-maximized")
            edge_options.add_argument("--disable-extensions")
            edge_options.add_argument("--disable-popup-blocking")
            edge_options.add_argument("--disable-gpu")
            edge_options.add_argument("--no-sandbox")
            edge_options.add_argument("--disable-dev-shm-usage")
            edge_options.add_argument(
                f"--user-data-dir={os.path.join(os.path.expanduser('~'), 'edge_automation_profile')}")
            edge_options.add_argument(
                f"download.default_directory={self.folder_var.get()}")

            # Initialize driver with multiple approaches
            driver_initialized = False

            try:
                if webdriver_manager_available:
                    self.log_message(
                        "Trying WebDriver Manager approach...", "STATUS")
                    edge_driver_path = EdgeChromiumDriverManager().install()
                    service = Service(executable_path=edge_driver_path)
                    self.driver = webdriver.Edge(
                        service=service, options=edge_options)
                    driver_initialized = True
                    self.log_message(
                        "Edge initialized with WebDriver Manager", "SUCCESS")
            except Exception as e:
                self.log_message(
                    f"WebDriver Manager approach failed: {str(e)}", "WARNING")

            if not driver_initialized:
                try:
                    self.log_message(
                        "Trying default Service approach...", "STATUS")
                    service = Service()
                    self.driver = webdriver.Edge(
                        service=service, options=edge_options)
                    driver_initialized = True
                    self.log_message(
                        "Edge initialized with default Service", "SUCCESS")
                except Exception as e:
                    self.log_message(
                        f"Default Service approach failed: {str(e)}", "WARNING")

            if not driver_initialized:
                try:
                    self.log_message(
                        "Trying system PATH approach...", "STATUS")
                    self.driver = webdriver.Edge(options=edge_options)
                    driver_initialized = True
                    self.log_message(
                        "Edge initialized with system PATH", "SUCCESS")
                except Exception as e:
                    self.log_message(
                        f"System PATH approach failed: {str(e)}", "WARNING")

            if not driver_initialized:
                raise Exception("Failed to initialize Edge WebDriver")

            # Set page load timeout
            self.driver.set_page_load_timeout(60)

            # Get all emails to process
            primary_email = self.email_var.get()
            extra_emails_text = self.extra_emails_text.get(
                "1.0", tk.END).strip()
            all_emails = [primary_email]

            if extra_emails_text:
                extra_emails = [email.strip() for email in extra_emails_text.replace('\n', ',').split(',')
                                if email.strip() and email.strip() != primary_email]
                all_emails.extend(extra_emails)

            # Create timestamp for this run
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Define the path for the patent_full_list.csv file
            csv_file = os.path.join(os.path.dirname(
                __file__), 'config', 'patent_full_list.csv')

            # Process each email
            for email_index, current_email in enumerate(all_emails, 1):
                if self.stop_event.is_set():
                    break

                try:
                    self.log_message(
                        f"\nProcessing email {email_index}/{len(all_emails)}: {current_email}", "STATUS")

                    # Navigation sequence with retry mechanism
                    max_retries = 3
                    for retry in range(max_retries):
                        try:
                            if email_index == 1 or not self.driver.current_url.endswith('/patent'):
                                self.driver.get(
                                    "https://sp-wildfly-cipo-itm-patents-sp-prod.apps.ocp.prod.ised-isde.canada.ca/backoffice/patent")
                                # Wait for page to load
                                WebDriverWait(self.driver, 30).until(
                                    lambda driver: driver.execute_script(
                                        'return document.readyState') == 'complete'
                                )
                                # Initial button click with enhanced stale element handling
                                self.click_action_button_robust()
                            break
                        except Exception as e:
                            if retry == max_retries - 1:
                                raise
                            self.log_message(
                                f"Retry {retry + 1} for navigation: {str(e)}", "WARNING")
                            time.sleep(2)

                    # Set the current email
                    # Set the current email only if it's different from the current value
                    if self.email_var.get() != current_email:
                        self.email_var.set(current_email)

                    # User selection and processing
                    self.select_user()

                    # Proceed with data extraction
                    self.set_page_size()
                    self.extract_table_data()

                    # Process applications for this email
                    self.process_email_applications(current_email, csv_file)

                except Exception as e:
                    self.log_message(
                        f"Error processing email {current_email}: {str(e)}", "ERROR")
                    self.log_message(traceback.format_exc(), "ERROR")
                    continue

            self.log_message("\nAll email processing completed", "SUCCESS")

            # Open the output file
            self.open_output_file(csv_file)

        except WebDriverException as e:
            self.log_message(f"Edge initialization failed: {str(e)}", "ERROR")
            self.log_message(
                "Ensure Microsoft Edge is installed and updated", "WARNING")
            self.log_message(
                "Also verify that webdriver-manager is installed (pip install webdriver-manager)", "WARNING")
            self.log_message(
                "Try running the script with administrator privileges", "WARNING")
            self.log_message(
                "Check if Edge version matches msedgedriver version", "WARNING")
        except Exception as e:
            self.log_message(f"Process failed: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")
        finally:
            self.running = False
            self.root.after(
                0, lambda: self.fetch_applications_btn.config(state=tk.NORMAL))
            # Ensure driver is closed
            if hasattr(self, 'driver') and self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            self.driver = None  # Clear the driver reference

    def click_action_button_robust(self):
        """Enhanced robust handling for initial action button click with multiple fallback mechanisms"""
        self.log_message(
            "Clicking action button with robust handling...", "STATUS")

        # Multiple button selectors to try
        button_selectors = [
            "#root > div > div > div.button-panel > div:nth-child(1) > button",
            "div.button-panel button",
            "button.btn-primary",
            "button.MuiButton-root"
        ]

        # Wait for page to fully load
        time.sleep(5)

        # Try each selector with multiple approaches
        for selector in button_selectors:
            try:
                self.log_message(f"Trying selector: {selector}", "STATUS")

                # Approach 1: Standard WebDriverWait and click
                try:
                    button = WebDriverWait(self.driver, 20).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))

                    # Scroll into view
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", button)
                    time.sleep(1)

                    button.click()
                    self.log_message(
                        "Button clicked successfully with standard approach", "SUCCESS")

                    # Verify navigation success
                    WebDriverWait(self.driver, 30).until(
                        lambda d: "patent" in d.current_url.lower())

                    self.log_message("Navigation successful", "SUCCESS")
                    time.sleep(2)
                    return
                except StaleElementReferenceException:
                    self.log_message(
                        "Stale element detected, trying alternative approach", "WARNING")
                except Exception as e:
                    self.log_message(
                        f"Standard click failed: {str(e)}", "WARNING")

                # Approach 2: JavaScript click
                try:
                    self.log_message("Trying JavaScript click...", "STATUS")
                    self.driver.execute_script(f"""
                        var button = document.querySelector('{selector}');
                        if(button) {{
                            button.scrollIntoView({{behavior: 'smooth', block: 'center'}});
                            setTimeout(function() {{
                                button.click();
                            }}, 500);
                        }}
                    """)
                    time.sleep(3)

                    # Verify navigation success
                    if "patent" in self.driver.current_url.lower():
                        self.log_message(
                            "Navigation successful with JavaScript click", "SUCCESS")
                        time.sleep(2)
                        return
                except Exception as e:
                    self.log_message(
                        f"JavaScript click failed: {str(e)}", "WARNING")

                # Approach 3: Action chains
                try:
                    self.log_message("Trying action chains...", "STATUS")
                    button = self.driver.find_element(
                        By.CSS_SELECTOR, selector)
                    webdriver.ActionChains(self.driver)\
                        .move_to_element(button)\
                        .pause(1)\
                        .click()\
                        .perform()
                    time.sleep(3)

                    # Verify navigation success
                    if "patent" in self.driver.current_url.lower():
                        self.log_message(
                            "Navigation successful with action chains", "SUCCESS")
                        time.sleep(2)
                        return
                except Exception as e:
                    self.log_message(
                        f"Action chains failed: {str(e)}", "WARNING")

            except Exception as e:
                self.log_message(
                    f"Selector {selector} failed: {str(e)}", "WARNING")

        # Last resort: Try to find any button that might be the action button
        try:
            self.log_message("Trying to find any visible button...", "STATUS")
            buttons = self.driver.find_elements(By.TAG_NAME, "button")
            for button in buttons:
                try:
                    if button.is_displayed() and button.is_enabled():
                        button_text = button.text.lower()
                        self.log_message(
                            f"Found potential button: {button_text}", "STATUS")

                        # Scroll into view
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", button)
                        time.sleep(1)

                        # Click the button
                        button.click()
                        time.sleep(3)

                        # Verify navigation success
                        if "patent" in self.driver.current_url.lower():
                            self.log_message(
                                "Navigation successful with generic button", "SUCCESS")
                            time.sleep(2)
                            return
                except StaleElementReferenceException:
                    continue
                except Exception as e:
                    continue
        except Exception as e:
            self.log_message(
                f"Generic button approach failed: {str(e)}", "WARNING")

        # Final attempt: Direct navigation
        try:
            self.log_message("Trying direct navigation...", "STATUS")
            self.driver.get(
                "https://sp-wildfly-cipo-itm-patents-sp-prod.apps.ocp.prod.ised-isde.canada.ca/backoffice/patent/application")
            time.sleep(5)

            # Verify navigation success
            if "patent" in self.driver.current_url.lower() and "application" in self.driver.current_url.lower():
                self.log_message("Direct navigation successful", "SUCCESS")
                return
        except Exception as e:
            self.log_message(f"Direct navigation failed: {str(e)}", "ERROR")

        # If all approaches failed
        raise Exception("All action button click approaches failed")

    def select_user(self):
        """Handle user selection with email input"""
        try:
            self.log_message("Configuring user credentials...", "STATUS")

            # Open user selector with increased timeout
            user_selector = "#app-view-container > div.view-routes > div > div > div:nth-child(2) > div > div > div > div.d-flex.flex-row.flex-wrap.gap-3 > div.d-flex.flex-row.flex-wrap.gap-3.p-2.JlxaqMmJxkvC39GgqnSS > div:nth-child(1) > div.MuiAutocomplete-root.MuiAutocomplete-hasClearIcon.MuiAutocomplete-hasPopupIcon.css-dt8xbo > div > div > div > button.MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeMedium.MuiAutocomplete-popupIndicator.css-uge3vf"

            # Try multiple selector approaches
            try:
                selector_element = WebDriverWait(self.driver, 60).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, user_selector)))
                selector_element.click()
            except Exception as e:
                self.log_message(
                    f"Primary selector failed: {str(e)}", "WARNING")
                # Try alternative XPath approach
                try:
                    self.log_message(
                        "Trying alternative selector approach...", "STATUS")
                    # Try a more generic selector
                    selector_element = WebDriverWait(self.driver, 60).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiAutocomplete-popupIndicator')]")))
                    selector_element.click()
                except Exception as e2:
                    self.log_message(
                        f"Alternative selector failed: {str(e2)}", "ERROR")
                    # Try JavaScript click on any visible dropdown
                    self.driver.execute_script("""
                        var buttons = document.querySelectorAll('button');
                        for(var i=0; i<buttons.length; i++) {
                            if(buttons[i].offsetParent !== null) {
                                buttons[i].click();
                                break;
                            }
                        }
                    """)

            # Select user type with increased timeout and multiple approaches
            try:
                WebDriverWait(self.driver, 60).until(
                    EC.visibility_of_element_located((By.XPATH, "//li[contains(., 'User')]"))).click()
            except Exception as e:
                self.log_message(
                    f"User type selection failed: {str(e)}", "WARNING")
                # Try JavaScript approach
                self.driver.execute_script("""
                    var items = document.querySelectorAll('li');
                    for(var i=0; i<items.length; i++) {
                        if(items[i].textContent.includes('User')) {
                            items[i].click();
                            break;
                        }
                    }
                """)

            # Input email with increased timeout and multiple approaches
            try:
                email_field = WebDriverWait(self.driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#mui-6")))
                email_field.clear()
                email_field.send_keys(self.email_var.get())
            except Exception as e:
                self.log_message(
                    f"Email field selection failed: {str(e)}", "WARNING")
                # Try alternative approach
                try:
                    # Try a more generic input selector
                    email_fields = self.driver.find_elements(
                        By.TAG_NAME, "input")
                    for field in email_fields:
                        if field.is_displayed():
                            field.clear()
                            field.send_keys(self.email_var.get())
                            break
                except Exception as e2:
                    self.log_message(
                        f"Alternative email input failed: {str(e2)}", "ERROR")
                    # Try JavaScript approach
                    self.driver.execute_script(f"""
                        var inputs = document.querySelectorAll('input');
                        for(var i=0; i<inputs.length; i++) {{
                            if(inputs[i].offsetParent !== null) {{
                                inputs[i].value = '{self.email_var.get()}';
                                break;
                            }}
                        }}
                    """)

            # Enhanced confirmation handling with increased timeout
            confirm_selector = "#app-view-container > div.view-routes > div > div > div:nth-child(2) > div > div > div > div.d-flex.flex-row.flex-wrap.gap-3 > div.d-flex.flex-row.flex-wrap.gap-3.p-2.JlxaqMmJxkvC39GgqnSS > div:nth-child(1) > button"

            # Try multiple approaches to find and click the confirm button
            try:
                # Wait for button to be both present and clickable
                confirm_button = WebDriverWait(self.driver, 60).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, confirm_selector)))

                # Scroll into view with offset for fixed headers
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                    confirm_button
                )
                time.sleep(1)  # Allow scroll completion

                # Attempt multiple interaction methods
                try:
                    # Standard click first
                    confirm_button.click()
                except ElementClickInterceptedException:
                    # Fallback 1: JavaScript click
                    self.log_message(
                        "Using JavaScript click fallback...", "WARNING")
                    self.driver.execute_script(
                        "arguments[0].click();", confirm_button)
                except Exception as e:
                    # Fallback 2: Action chains with offset
                    self.log_message(
                        "Using action chains fallback...", "WARNING")
                    webdriver.ActionChains(self.driver)\
                        .move_to_element(confirm_button)\
                        .pause(0.5)\
                        .click(confirm_button)\
                        .perform()
            except Exception as e:
                self.log_message(
                    f"Confirm button selection failed: {str(e)}", "WARNING")
                # Try alternative approach - find any button that might be the confirm button
                try:
                    buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        if button.is_displayed() and button.is_enabled():
                            button_text = button.text.lower()
                            if "confirm" in button_text or "submit" in button_text or "ok" in button_text or button_text == "":
                                self.log_message(
                                    f"Found potential confirm button: {button_text}", "STATUS")
                                button.click()
                                break
                except Exception as e2:
                    self.log_message(
                        f"Alternative confirm button approach failed: {str(e2)}", "ERROR")
                    # Last resort: JavaScript click on any visible button
                    self.driver.execute_script("""
                        var buttons = document.querySelectorAll('button');
                        for(var i=0; i<buttons.length; i++) {
                            if(buttons[i].offsetParent !== null) {
                                buttons[i].click();
                                break;
                            }
                        }
                    """)

            # No longer waiting for confirmation - proceed directly
            self.log_message(
                "Email submitted, proceeding to next step", "SUCCESS")
            time.sleep(3)  # Allow interface update

        except Exception as e:
            self.log_message(f"User configuration failed: {str(e)}", "ERROR")
            raise

    def set_page_size(self):
        """Set table page size to maximum"""
        try:
            self.log_message("Setting page size...", "STATUS")

            # Wait for the page size selector to be present with increased timeout
            try:
                page_size_selector = WebDriverWait(self.driver, 60).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#pageSizeSelect"))
                )

                # Create a Select object
                select = Select(page_size_selector)

                # Select the option with value '100'
                select.select_by_value('100')

                # Wait for the table to reload with new page size
                WebDriverWait(self.driver, 60).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "tbody tr[data-cy='entityTable']"))
                )

                self.log_message(
                    "Page size successfully set to 100", "SUCCESS")
                # Additional buffer time to ensure page is fully loaded
                time.sleep(2)
            except Exception as e:
                self.log_message(
                    f"Standard page size selection failed: {str(e)}", "WARNING")

                # Try JavaScript approach
                try:
                    self.log_message(
                        "Trying JavaScript approach for page size...", "STATUS")
                    self.driver.execute_script("""
                        var selects = document.querySelectorAll('select');
                        for(var i=0; i<selects.length; i++) {
                            if(selects[i].offsetParent !== null) {
                                for(var j=0; j<selects[i].options.length; j++) {
                                    if(selects[i].options[j].value === '100') {
                                        selects[i].value = '100';
                                        var event = new Event('change', { bubbles: true });
                                        selects[i].dispatchEvent(event);
                                        break;
                                    }
                                }
                            }
                        }
                    """)
                    self.log_message(
                        "JavaScript page size selection attempted", "STATUS")
                    time.sleep(3)  # Wait for potential changes to take effect
                except Exception as e2:
                    self.log_message(
                        f"JavaScript page size selection failed: {str(e2)}", "WARNING")
                    # Continue anyway, as this is not critical

        except Exception as e:
            self.log_message(f"Error setting page size: {str(e)}", "WARNING")
            self.log_message("Continuing with default page size", "WARNING")

    def extract_table_data(self):
        """Extract data from the patent table using JavaScript to avoid stale element references"""
        try:
            # Wait for table to be present
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.TAG_NAME, "tbody"))
            )
            time.sleep(3)  # Additional buffer time

            # Get table headers first to ensure we have all column names
            table_headers = self.driver.execute_script("""
                var headers = [];
                var headerCells = document.querySelectorAll('thead th');
                for(var i=0; i<headerCells.length; i++) {
                    headers.push(headerCells[i].textContent.trim());
                }
                return headers;
            """)

            # Check if headers were found
            if table_headers and len(table_headers) > 0:
                # Update basic headers with actual table headers
                self.table_headers = table_headers
            else:
                self.log_message(
                    "No table headers found, using default headers", "WARNING")
                # Keep default headers if none found

            # Use JavaScript to extract table data to avoid stale element references
            self.log_message(
                "Using JavaScript extraction to avoid stale elements...", "STATUS")
            table_data = self.driver.execute_script("""
                var result = [];
                var appNumbers = [];
                var rows = document.querySelectorAll('tbody tr');

                for(var i=0; i<rows.length; i++) {
                    var rowData = [];
                    var cells = rows[i].querySelectorAll('td');

                    for(var j=0; j<cells.length; j++) {
                        var cellText = '';
                        var aElement = cells[j].querySelector('a');
                        var spanElement = cells[j].querySelector('span');

                        if(aElement) {
                            cellText = aElement.textContent.trim();
                        } else if(spanElement) {
                            cellText = spanElement.textContent.trim();
                        } else {
                            cellText = cells[j].textContent.trim();
                        }

                        rowData.push(cellText);
                    }

                    if(rowData.length > 0) {
                        result.push(rowData);
                        if(rowData[0]) {
                            appNumbers.push(rowData[0]);
                        }
                    }
                }

                return {data: result, appNumbers: appNumbers, rowCount: rows.length};
            """)

            if table_data and 'data' in table_data and 'appNumbers' in table_data:
                # Filter the application numbers based on the user-provided list
                if self.application_numbers:
                    filtered_data = []
                    filtered_app_numbers = []
                    for row, app_num in zip(table_data['data'], table_data['appNumbers']):
                        if app_num in self.application_numbers:
                            filtered_data.append(row)
                            filtered_app_numbers.append(app_num)

                    self.csv_data = filtered_data
                    self.application_numbers = filtered_app_numbers
                else:
                    self.csv_data = table_data['data']
                    self.application_numbers = table_data['appNumbers']

                row_count = len(self.application_numbers)

                self.log_message(
                    f"JavaScript extraction found {row_count} rows in the table", "STATUS")
                self.log_message(
                    f"Successfully extracted data for {len(self.application_numbers)} applications", "SUCCESS")
            else:
                self.log_message(
                    "JavaScript extraction failed to return expected data structure", "ERROR")
                raise Exception("Failed to extract table data")
        except Exception as e:
            self.log_message(
                f"Table data extraction failed: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")
            raise

    def get_application_details(self, app_number):
        """Get details for a single application number"""
        try:
            # Clear and input application number with explicit wait and JavaScript
            input_field = WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "#inputText1"))
            )

            # Use JavaScript to clear and set value
            self.driver.execute_script("arguments[0].value = '';", input_field)
            input_field.send_keys(app_number)

            # Wait and click search button using JavaScript
            search_button = WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "#root > div > div > div.container-fluid.d-flex.justify-content-end.align-bottom.p-2.pr-3.col > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > button"))
            )

            # Use JavaScript click to bypass potential overlay issues
            self.driver.execute_script("arguments[0].click();", search_button)

            # Wait for data to load
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#mui-1"))
            )
            time.sleep(3)  # Additional buffer time

            # Extract application details using JavaScript
            details_data = self.driver.execute_script("""
                var result = {
                    basic_row: [],
                    headers_col1: [],
                    headers_col3: [],
                    data_col1: [],
                    data_col3: []
                };

                // Extract basic row data
                var basicRow = document.querySelector("#basicRow");
                if (basicRow) {
                    result.basic_row = basicRow.querySelectorAll('td').map(cell => cell.textContent.trim());
                }

                // Extract column 1 data
                var rows = document.querySelectorAll("#app-view-container > div.view-routes > div > div.undefined.container-fluid.p-0.m-2 > div.tab-content > div.tab-pane.container-fluid.active > div.d-flex.flex-row.flex-wrap.IwoYpnkL8KrJER38CYqf > div.dFzbiZbd3Lke0_3uy4DA > table > tbody tr");
                for(var i=0; i<rows.length; i++) {
                    var cells = rows[i].querySelectorAll('td');
                    if(cells.length >= 4) {
                        result.headers_col1.push(cells[0].textContent.trim());
                        result.data_col1.push(cells[1].textContent.trim());
                        result.headers_col3.push(cells[2].textContent.trim());
                        result.data_col3.push(cells[3].textContent.trim());
                    }
                }

                return result;
            """)

            return details_data

        except Exception as e:
            self.log_message(
                f"Error getting application details for {app_number}: {str(e)}", "ERROR")
            raise

    def update_csv_with_specific_row(self, csv_path, app_number, app_details_headers_col1, app_details_headers_col3, app_details_data_col1, app_details_data_col3):
        """Update CSV with application details"""
        try:
            # Read existing data
            with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                existing = list(reader)

            # Get selected columns
            selected_basic_columns = self.get_selected_columns()
            selected_ad_columns = self.get_selected_ad_columns()

            # Base headers - filter based on selected columns
            # Always include Application Number
            base_headers = ["Application Number"]
            for col in self.table_headers:
                if col != "Application Number" and col in selected_basic_columns:
                    base_headers.append(col)

            # Prepare unique headers for App Details with _AD suffix
            def create_unique_header(header):
                # Remove any existing prefix and create a clean header with _AD suffix
                clean_header = header.replace('App_Details_Tab_', '').strip()
                return f"{clean_header}_AD"

            # Create unique headers for columns
            unique_headers_col1 = [create_unique_header(
                h) for h in app_details_headers_col1]
            unique_headers_col3 = [create_unique_header(
                h) for h in app_details_headers_col3]

            # Update the application detail headers list
            new_ad_headers = []
            for header in unique_headers_col1 + unique_headers_col3:
                if header not in self.ad_headers:
                    new_ad_headers.append(header)
                    self.ad_headers.append(header)
                    self.ad_column_vars[header] = IntVar(
                        value=1)  # Default to selected

            if new_ad_headers:
                self.log_message(
                    f"Found {len(new_ad_headers)} new application detail headers", "STATUS")

            # Filter application detail headers based on selection
            filtered_headers_col1 = [
                h for h in unique_headers_col1 if h in selected_ad_columns]
            filtered_headers_col3 = [
                h for h in unique_headers_col3 if h in selected_ad_columns]

            # Create filtered data based on selected headers
            filtered_data_col1 = [v for h, v in zip(
                unique_headers_col1, app_details_data_col1) if h in selected_ad_columns]
            filtered_data_col3 = [v for h, v in zip(
                unique_headers_col3, app_details_data_col3) if h in selected_ad_columns]

            # Combine base headers with filtered application detail headers
            all_headers = base_headers + filtered_headers_col1 + filtered_headers_col3

            # Find or create the target row
            target_row = next(
                (r for r in existing[1:] if r and r[0] == str(app_number)), None)
            if not target_row:
                target_row = [str(app_number)] + [''] * (len(all_headers) - 1)
                existing.append(target_row)

            # Ensure the row has enough columns
            while len(target_row) < len(all_headers):
                target_row.append('')

            # Update the headers
            existing[0] = all_headers

            # Map and update data
            header_map = {h: i for i, h in enumerate(all_headers)}

            # Update column 1 data
            for header, value in zip(filtered_headers_col1, filtered_data_col1):
                if header in header_map:
                    target_row[header_map[header]] = value

            # Update column 3 data
            for header, value in zip(filtered_headers_col3, filtered_data_col3):
                if header in header_map:
                    target_row[header_map[header]] = value

            # Write back to CSV
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerows(existing)

            self.log_message(
                f"Successfully updated {app_number} with {len(all_headers)} columns", "SUCCESS")

        except Exception as e:
            self.log_message(f"CSV Error: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

    def process_application_details(self, app_number, csv_file, email):
        """Process details for a single application number with email"""
        try:
            # Clear and input application number with explicit wait and JavaScript
            input_field = WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "#inputText1"))
            )

            # Use JavaScript to clear and set value
            self.driver.execute_script("arguments[0].value = '';", input_field)
            input_field.send_keys(app_number)

            # Wait and click search button using JavaScript
            search_button = WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "#root > div > div > div.container-fluid.d-flex.justify-content-end.align-bottom.p-2.pr-3.col > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > button"))
            )

            # Use JavaScript click to bypass potential overlay issues
            self.driver.execute_script("arguments[0].click();", search_button)

            # Wait for data to load
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#mui-1"))
            )
            time.sleep(3)

            # Extract application details using JavaScript
            details_data = self.driver.execute_script("""
                var col1_headers = [];
                var col1_data = [];
                var col3_headers = [];
                var col3_data = [];

                try {
                    var rows = document.querySelectorAll("#app-view-container > div.view-routes > div > div.undefined.container-fluid.p-0.m-2 > div.tab-content > div.tab-pane.container-fluid.active > div.d-flex.flex-row.flex-wrap.IwoYpnkL8KrJER38CYqf > div.dFzbiZbd3Lke0_3uy4DA > table > tbody tr");

                    for(var i=0; i<rows.length; i++) {
                        var cells = rows[i].querySelectorAll('td');
                        if(cells.length >= 4) {
                            col1_headers.push(cells[0].textContent.trim());
                            col1_data.push(cells[1].textContent.trim());
                            col3_headers.push(cells[2].textContent.trim());
                            col3_data.push(cells[3].textContent.trim());
                        }
                    }
                } catch(e) {
                    console.error("Error extracting details: " + e);
                }

                return {
                    col1_headers: col1_headers,
                    col1_data: col1_data,
                    col3_headers: col3_headers,
                    col3_data: col3_data
                };
            """)

            # Update CSV with application details including email
            if details_data and 'col1_headers' in details_data:
                self.update_csv_with_data(
                    csv_file,
                    email,
                    app_number,
                    details_data['col1_headers'],
                    details_data['col3_headers'],
                    details_data['col1_data'],
                    details_data['col3_data']
                )
            else:
                self.log_message(
                    f"Failed to extract details for {app_number}", "ERROR")
                raise Exception("Details extraction failed")

        except TimeoutException as te:
            self.log_message(
                f"Timeout processing {app_number}: {str(te)}", "ERROR")
            raise
        except Exception as e:
            self.log_message(
                f"Error processing application {app_number}: {str(e)}", "ERROR")
            raise

    def stop_process(self):
        """Stop the current process"""
        if self.running:
            self.stop_event.set()
            self.log_message("Aborting process...", "WARNING")
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            self.running = False
            self.fetch_applications_btn.config(state=tk.NORMAL)

    def _update_log(self, message, tag):
        """Update log text widget (must be called from main thread)"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def on_close(self):
        """Handle window close event"""
        if self.running:
            if messagebox.askyesno("Confirm Exit", "A process is still running. Are you sure you want to exit?"):
                self.stop_process()
                self.root.destroy()
        else:
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = PatentAutomationApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)  # Add proper close handler
    root.mainloop()
