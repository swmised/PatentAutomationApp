import os
import time
import csv
import tkinter as tk
import traceback
from tkinter import ttk, messagebox, filedialog
from threading import Thread, Event
from datetime import datetime
import subprocess
import sys

# ... (Selenium imports and other code as before)

class PatentAutomationApp:
    # --- Method Ordering is Key ---

    def _get_data_path(self, filename, use_config=False):
        # (Your existing _get_data_path code)
        pass

    def load_column_settings(self):
        # (Your existing load_column_settings code)
        pass

    def load_filter_settings(self):
        # (Your existing load_filter_settings code)
        pass

    def create_default_column_settings(self):
        # (Your existing create_default_column_settings code)
        pass

    def create_default_filter_settings(self):
        # (Your existing create_default_filter_settings code)
        pass

    def populate_filter_fields(self, headers=None):
        # (Your existing populate_filter_fields code)
        pass

    def populate_column_checkboxes(self):
        # (Your existing populate_column_checkboxes code)
        pass

    def save_column_settings(self):
        # (Your existing save_column_settings code)
        pass

    def save_filter_settings(self):
        # (Your existing save_filter_settings code)
        pass

    def __init__(self, root):
        # (Your existing __init__ code)
        pass

    def setup_ui(self):
        # (Your existing setup_ui code)
        pass

    def select_folder(self):
        # (Your existing select_folder code)
        pass

    def start_processing(self):
        # (Your existing start_processing code)
        pass

    def kill_edge_processes(self):
        # (Your existing kill_edge_processes code)
        pass

    def get_edge_version(self):
        # (Your existing get_edge_version code)
        pass

    def full_workflow(self):
        # (Your existing full_workflow code)
        pass

    def click_action_button(self):
        # (Your existing click_action_button code)
        pass

    def select_user(self):
        # (Your existing select_user code)
        pass

    def set_page_size(self):
        # (Your existing set_page_size code)
        pass

    def extract_table_data(self):
        # (Your existing extract_table_data code)
        pass

    def update_csv_with_specific_row(self, csv_path, app_number, app_details_headers_col1, app_details_headers_col3, app_details_data_col1, app_details_data_col3):
        # (Your existing update_csv_with_specific_row code)
        pass

    def process_application_details(self, app_number, csv_file):
        # (Your existing process_application_details code)
        pass

    def process_applications(self):
        # (Your existing process_applications code)
        pass

    def generate_csv(self):
        # (Your existing generate_csv code)
        pass

    def stop_process(self):
        # (Your existing stop_process code)
        pass

    def log_message(self, message, tag=""):
        # (Your existing log_message code)
        pass

    def _update_log(self, message, tag):
        # (Your existing _update_log code)
        pass

    def on_close(self):
        # (Your existing on_close code)
        pass    
    def __init__(self, root):
        self.root = root
        self.driver = None
        self.running = False
        self.stop_event = Event()
        self.application_numbers = []
        self.csv_data = []
        self.current_process = None

        # Determine the script's directory and config directory (MUST be done first)
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_dir = os.path.join(self.script_dir, "config")

        # Ensure the config directory exists
        os.makedirs(self.config_dir, exist_ok=True)

        # Initialize Tkinter variables FIRST
        self.email_var = tk.StringVar()
        self.folder_var = tk.StringVar()

        # Initialize column and filter settings
        self.column_vars = {}
        self.filter_conditions = {}
        self.filter_values = {}
        self.filter_entries = {}
        self.default_headers = [] # Add this line
        
        # THEN setup UI components
        self.setup_ui()

        # Load settings AFTER UI setup
        self.load_column_settings()
        self.load_filter_settings()

        # Set default values AFTER UI setup
        self.email_var.set("simon.chau@ised-isde.gc.ca")
        self.folder_var.set(os.path.join(os.path.expanduser('~'), 'Documents'))

    def setup_ui(self):
        self.root.title("Patent Automation Suite v3.1 (Edge)")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)

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

        # Folder Selection
        ttk.Label(config_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, padx=5)
        folder_entry = ttk.Entry(config_frame, textvariable=self.folder_var)
        folder_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Button(config_frame, text="Browse", command=self.select_folder).grid(row=1, column=2, padx=5)

        # Control Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        self.start_btn = ttk.Button(btn_frame, text="Start Processing", command=self.start_processing)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(btn_frame, text="Abort Process", command=self.stop_process)
        self.stop_btn.pack(side=tk.RIGHT, padx=5)

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

    def select_folder(self):
        selected_dir = filedialog.askdirectory(
            initialdir=self.folder_var.get(),
            title="Select Output Folder"
        )
        if selected_dir:
            self.folder_var.set(selected_dir)

    def start_processing(self):
        """Start complete automation workflow"""
        if not self.running:
            self.running = True
            self.stop_event.clear()
            self.start_btn.config(state=tk.DISABLED)
            self.current_process = Thread(target=self.full_workflow, daemon=True)
            self.current_process.start()

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
            self.log_message(f"Error terminating Edge processes: {str(e)}", "WARNING")

    def get_edge_version(self):
        """Get the installed Edge browser version"""
        try:
            if os.name == 'nt':  # Windows
                cmd = r'reg query "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon" /v version'
                result = subprocess.check_output(cmd, shell=True).decode('utf-8')
                version = result.strip().split()[-1]
            else:  # Linux
                cmd = "microsoft-edge --version"
                result = subprocess.check_output(cmd, shell=True).decode('utf-8')
                version = result.strip().split()[-1]

            self.log_message(f"Detected Edge version: {version}", "STATUS")
            return version
        except Exception as e:
            self.log_message(f"Could not detect Edge version: {str(e)}", "WARNING")
            return None

    def full_workflow(self):
        try:
            # Kill any running Edge processes
            self.kill_edge_processes()

            # Get Edge version
            edge_version = self.get_edge_version()

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
            edge_options.add_argument(f"--user-data-dir={os.path.join(os.path.expanduser('~'), 'edge_automation_profile')}")
            edge_options.add_argument(f"download.default_directory={self.folder_var.get()}")

            # Try to find Edge binary location
            edge_binary = None
            if os.name == 'nt':  # Windows
                possible_paths = [
                    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
                ]
                for path in possible_paths:
                    if os.path.exists(path):
                        edge_binary = path
                        break
            else:  # Linux
                possible_paths = [
                    "/usr/bin/microsoft-edge",
                    "/usr/bin/microsoft-edge-stable"
                ]
                for path in possible_paths:
                    if os.path.exists(path):
                        edge_binary = path
                        break

            if edge_binary:
                self.log_message(f"Using Edge binary: {edge_binary}", "STATUS")
                edge_options.binary_location = edge_binary

            # Try multiple approaches to initialize Edge WebDriver
            driver_initialized = False

            # Approach 1: Use WebDriver Manager if available
            if webdriver_manager_available:
                try:
                    self.log_message("Trying WebDriver Manager approach...", "STATUS")
                    edge_driver_path = EdgeChromiumDriverManager().install()
                    service = Service(executable_path=edge_driver_path)
                    self.driver = webdriver.Edge(service=service, options=edge_options)
                    driver_initialized = True
                    self.log_message("Edge initialized with WebDriver Manager", "SUCCESS")
                except Exception as e:
                    self.log_message(f"WebDriver Manager approach failed: {str(e)}", "WARNING")

            # Approach 2: Try with default Service (no path)
            if not driver_initialized:
                try:
                    self.log_message("Trying default Service approach...", "STATUS")
                    service = Service()
                    self.driver = webdriver.Edge(service=service, options=edge_options)
                    driver_initialized = True
                    self.log_message("Edge initialized with default Service", "SUCCESS")
                except Exception as e:
                    self.log_message(f"Default Service approach failed: {str(e)}", "WARNING")

            # Approach 3: Try with system PATH
            if not driver_initialized:
                try:
                    self.log_message("Trying system PATH approach...", "STATUS")
                    self.driver = webdriver.Edge(options=edge_options)
                    driver_initialized = True
                    self.log_message("Edge initialized with system PATH", "SUCCESS")
                except Exception as e:
                    self.log_message(f"System PATH approach failed: {str(e)}", "WARNING")

            # If all approaches failed, raise exception
            if not driver_initialized:
                raise Exception("All Edge initialization approaches failed")

            # Set page load timeout
            self.driver.set_page_load_timeout(60)

            # Navigation sequence
            self.log_message("Navigating to patent portal...", "STATUS")
            self.driver.get("https://sp-wildfly-cipo-itm-patents-sp-prod.apps.ocp.prod.ised-isde.canada.ca/backoffice/patent")

            # Initial button click
            self.click_action_button()

            # User selection and processing
            self.select_user()

            # Proceed directly to next steps without waiting for confirmation
            self.set_page_size()
            self.extract_table_data()
            self.process_applications()

            # Generate output
            self.generate_csv()

            self.log_message("Process completed successfully", "SUCCESS")

        except WebDriverException as e:
            self.log_message(f"Edge initialization failed: {str(e)}", "ERROR")
            self.log_message("Ensure Microsoft Edge is installed and updated", "WARNING")
            self.log_message("Also verify that webdriver-manager is installed (pip install webdriver-manager)", "WARNING")
            self.log_message("Try running the script with administrator privileges", "WARNING")
            self.log_message("Check if Edge version matches msedgedriver version", "WARNING")
        except Exception as e:
            self.log_message(f"Process failed: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")
        finally:
            self.running = False
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            # Ensure driver is closed
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass

    def click_action_button(self):
        """Handle initial action button click with Edge-specific adjustments"""
        try:
            self.log_message("Clicking action button...", "STATUS")
            button_selector = "#root > div > div > div.button-panel > div:nth-child(1) > button"

            # Wait with increased timeout
            button = WebDriverWait(self.driver, 60).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, button_selector)))

            # Edge-specific view adjustment
            if not button.is_displayed():
                self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", button)
                time.sleep(1)

            try:
                button.click()
            except ElementClickInterceptedException:
                self.log_message("Using JavaScript click fallback...", "WARNING")
                self.driver.execute_script("arguments[0].click();", button)

            # Verify navigation success
            WebDriverWait(self.driver, 60).until(
                lambda d: "patent" in d.current_url.lower())

            self.log_message("Navigation successful", "SUCCESS")
            time.sleep(2)  # Allow page stabilization

        except TimeoutException:
            self.log_message("Page transition timeout occurred", "ERROR")
            raise

    def select_user(self):
        """Handle user selection with email input"""
        try:
            self.log_message("Configuring user credentials...", "STATUS")

            # Open user selector with increased timeout
            user_selector = "#app-view-container > div.view-routes > div > div > div:nth-child(2) > div > div > div > div.d-flex.flex-row.flex-wrap.gap-3 > div.d-flex.flex-row.flex-wrap.gap-3.p-2.JlxaqMmJxkvC39GgqnSS > div:nth-child(1) > div.MuiAutocomplete-root.MuiAutocomplete-hasClearIcon.MuiAutocomplete-hasPopupIcon.css-dtxbo > div > div > div > button.MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeMedium.MuiAutocomplete-popupIndicator.css-uge3vf"

            # Try multiple selector approaches
            try:
                selector_element = WebDriverWait(self.driver, 60).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, user_selector)))
                selector_element.click()
            except Exception as e:
                self.log_message(f"Primary selector failed: {str(e)}", "WARNING")
                # Try alternative XPath approach
                try:
                    self.log_message("Trying alternative selector approach...", "STATUS")
                    # Try a more generic selector
                    selector_element = WebDriverWait(self.driver, 60).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiAutocomplete-popupIndicator')]")))
                    selector_element.click()
                except Exception as e2:
                    self.log_message(f"Alternative selector failed: {str(e2)}", "ERROR")
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
                self.log_message(f"User type selection failed: {str(e)}", "WARNING")
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
                self.log_message(f"Email field selection failed: {str(e)}", "WARNING")
                # Try alternative approach
                try:
                    # Try a more generic input selector
                    email_fields = self.driver.find_elements(By.TAG_NAME, "input")
                    for field in email_fields:
                        if field.is_displayed():
                            field.clear()
                            field.send_keys(self.email_var.get())
                            break
                except Exception as e2:
                    self.log_message(f"Alternative email input failed: {str(e2)}", "ERROR")
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
            confirm_selector = "#app-view-container > div.view-routes > div > div > div:nth-child(2) > div    > div > div > div.d-flex.flex-row.flex-wrap.gap-3 > div.d-flex.flex-row.flex-wrap.gap-3.p-2.JlxaqMmJxkvC39GgqnSS > div:nth-child(1) > button"

            # Try multiple approaches to find and click the confirm button
            try:
                # Wait for button to be both present and clickable
                confirm_button = WebDriverWait(self.driver, 60).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, confirm_selector)))

                # Scroll into view with offset for fixed headers
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({behavior: 'smooth',     block: 'center'});",
                    confirm_button
                )
                time.sleep(1)  # Allow scroll completion

                # Attempt multiple interaction methods
                try:
                    # Standard click first
                    confirm_button.click()
                except ElementClickInterceptedException:
                    # Fallback 1: JavaScript click
                    self.log_message("Using JavaScript click fallback...", "WARNING")
                    self.driver.execute_script("arguments[0].click();", confirm_button)
                except Exception as e:
                    # Fallback 2: Action chains with offset
                    self.log_message("Using action chains fallback...", "WARNING")
                    webdriver.ActionChains(self.driver)\
                        .move_to_element(confirm_button)\
                        .pause(0.5)\
                        .click(confirm_button)\
                        .perform()
            except Exception as e:
                self.log_message(f"Confirm button selection failed: {str(e)}", "WARNING")
                # Try alternative approach - find any button that might be the confirm button
                try:
                    buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        if button.is_displayed() and button.is_enabled():
                            button_text = button.text.lower()
                            if "confirm" in button_text or "submit" in button_text or "ok" in button_text or button_text == "":
                                self.log_message(f"Found potential confirm button: {button_text}", "STATUS")
                                button.click()
                                break
                except Exception as e2:
                    self.log_message(f"Alternative confirm button approach failed: {str(e2)}", "ERROR")
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
            self.log_message("Email submitted, proceeding to next step", "SUCCESS")
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
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#pageSizeSelect"))
                )

                # Create a Select object
                select = Select(page_size_selector)

                # Select the option with value '100'
                select.select_by_value('100')

                # Wait for the table to reload with new page size
                WebDriverWait(self.driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr[data-cy='entityTable']"))
                )

                self.log_message("Page size successfully set to 100", "SUCCESS")
                time.sleep(2)  # Additional buffer time to ensure page is fully loaded
            except Exception as e:
                self.log_message(f"Standard page size selection failed: {str(e)}", "WARNING")

                # Try JavaScript approach
                try:
                    self.log_message("Trying JavaScript approach for page size...", "STATUS")
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
                    self.log_message("JavaScript page size selection attempted", "STATUS")
                    time.sleep(3)  # Wait for potential changes to take effect
                except Exception as e2:
                    self.log_message(f"JavaScript page size selection failed: {str(e2)}", "WARNING")
                    # Continue anyway, as this is not critical

        except Exception as e:
            self.log_message(f"Error setting page size: {str(e)}", "WARNING")
            self.log_message("Continuing with default page size", "WARNING")

    def extract_table_data(self):
        """Extract data from the patent table using JavaScript to avoid stale element references"""
        try:
            self.log_message("Extracting table data...", "STATUS")

            # Wait for table to be present
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.TAG_NAME, "tbody"))
            )
            time.sleep(3)  # Additional buffer time

            # Use JavaScript to extract table data to avoid stale element references
            self.log_message("Using JavaScript extraction to avoid stale elements...", "STATUS")
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
                self.csv_data = table_data['data']
                self.application_numbers = table_data['appNumbers']
                row_count = table_data.get('rowCount', 0)

                self.log_message(f"JavaScript extraction found {row_count} rows in the table", "STATUS")
                self.log_message(f"Successfully extracted data for {len(self.application_numbers)} applications", "SUCCESS")

                # Verify we have application numbers
                if not self.application_numbers:
                    self.log_message("No application numbers found in the table", "WARNING")

                    # Try alternative approach with direct DOM access
                    self.log_message("Trying alternative extraction approach...", "STATUS")
                    alt_app_numbers = self.driver.execute_script("""
                        var numbers = [];
                        var cells = document.querySelectorAll('tbody tr td:first-child');
                        for(var i=0; i<cells.length; i++) {
                            var text = cells[i].textContent.trim();
                            if(text) numbers.push(text);
                        }
                        return numbers;
                    """)

                    if alt_app_numbers and len(alt_app_numbers) > 0:
                        self.application_numbers = alt_app_numbers
                        self.log_message(f"Alternative approach found {len(self.application_numbers)} application     numbers", "SUCCESS")
            else:
                self.log_message("JavaScript extraction failed to return expected data structure", "ERROR")
                raise Exception("Failed to extract table data")

        except Exception as e:
            self.log_message(f"Table data extraction failed: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")
            raise

    def update_csv_with_specific_row(self, csv_path, app_number, app_details_headers_col1, app_details_headers_col3, app_details_data_col1, app_details_data_col3):
        """Update CSV with application details"""
        try:
            # Read existing data
            with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                existing = list(reader)

            # Base headers
            base_headers = [
                "Application Number", "Action", "Chapter", "Due Date",
                "Received Date", "Status", "Email",     "Comments", "Checkbox"
            ]

            # Prepare unique headers for App Details with _AD suffix
            def create_unique_header(header):
                # Remove any existing prefix and create a clean header with _AD suffix
                clean_header = header.replace('App_Details_Tab_', '').strip()
                return f"{clean_header}_AD"

            # Create unique headers for columns
            unique_headers_col1 = [create_unique_header(h) for h in app_details_headers_col1]
            unique_headers_col3 = [create_unique_header(h) for h in app_details_headers_col3]

            # Combine base headers with new unique headers
            all_headers = base_headers + unique_headers_col1     + unique_headers_col3

            # Find or create the target row
            target_row = next((r for r in existing[1:] if r and r[0] == str(app_number)), None)
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
            for header, value in zip(unique_headers_col1, app_details_data_col1):
                if header in header_map:
                    target_row[header_map[header]] = value

            # Update column 3 data
            for header, value in zip(unique_headers_col3, app_details_data_col3):
                if header in header_map:
                    target_row[header_map[header]] = value

            # Write back to CSV
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerows(existing)

            self.log_message(f"Successfully updated {app_number} with {len(all_headers)} columns", "SUCCESS")

        except Exception as e:
            self.log_message(f"CSV Error: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

    def process_application_details(self, app_number, csv_file):
        """Process details for a single application number"""
        try:
            # Clear and input application number with explicit wait and JavaScript
            input_field = WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#inputText1"))
            )

            # Use JavaScript to clear and set value
            self.driver.execute_script("arguments[0].value = '';", input_field)
            input_field.send_keys(app_number)

            # Wait and click search button using JavaScript
            search_button = WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#root > div > div > div.container-fluid.d-flex.justify-content-end.align-bottom.p-2.pr-3.col > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > button"))
            )
            
            

            # Use JavaScript click to bypass potential overlay issues
            self.driver.execute_script("arguments[0].click();", search_button)

            # Wait for results to load
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#root > div > div > div.container-fluid.d-flex.justify-content-center.align-items-center.pr-3.col > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > div > div > div > div > div > div > div > table > tbody > tr:nth-child(1) > td:nth-child(1) > a"))
            )

            # Click the first result
            first_result_link = WebDriverWait(self.driver, 60).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#root > div > div > div.container-fluid.d-flex.justify-content-center.align-items-center.pr-3.col > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > div > div > div > div > div > div > div > table > tbody > tr:nth-child(1) > td:nth-child(1) > a"))
            )
            first_result_link.click()

            # Wait for details to load
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#root > div > div > div.container-fluid.d-flex.justify-content-center.align-items-center.pr-3.col > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > div > div > div > div > div > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > div > div:nth-child(1) > div > div > div:nth-child(1) > div > div > div:nth-child(2) > div > div > div > div > div > div > div > table > tbody > tr:nth-child(1) > td:nth-child(1)"))
            )

            # Extract details using JavaScript
            details = self.driver.execute_script("""
                var details = {
                    col1_headers: [],
                    col3_headers: [],
                    col1_data: [],
                    col3_data: []
                };

                // Function to extract headers and data from a table
                function extractTableDetails(tableSelector, headersArray, dataArray) {
                    var table = document.querySelector(tableSelector);
                    if (table) {
                        var headerCells = table.querySelectorAll('thead th');
                        headerCells.forEach(cell => headersArray.push(cell.textContent.trim()));

                        var dataRows = table.querySelectorAll('tbody tr');
                        dataRows.forEach(row => {
                            var firstCell = row.querySelector('td:nth-child(1)');
                            var secondCell = row.querySelector('td:nth-child(2)');
                            if (firstCell) dataArray.push(firstCell.textContent.trim());
                            if (secondCell) dataArray.push(secondCell.textContent.trim());
                        });
                    }
                }

                // Extract details from the first table
                extractTableDetails(
                    '#root > div > div > div.container-fluid.d-flex.justify-content-center.align-items-center.pr-3.col > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > div > div > div > div > div > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > div > div:nth-child(1) > div > div > div:nth-child(1) > div > div > div:nth-child(2) > div > div > div > div > div > div > div > table',
                    details.col1_headers,
                    details.col1_data
                );

                // Extract details from the second table
                extractTableDetails(
                    '#root > div > div > div.container-fluid.d-flex.justify-content-center.align-items-center.pr-3.col > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > div > div > div > div > div > div.container-fluid.ml-0.mt-1.yF5lUFS5qz0ybZFagftI > div > div:nth-child(1) > div > div > div:nth-child(2) > div > div > div:nth-child(2) > div > div > div > div > div > div > div > table',
                    details.col3_headers,
                    details.col3_data
                );

                return details;
            """)

            # Log the extracted details
            self.log_message(f"Extracted App Details for {app_number}", "SUCCESS")

            # Update CSV with extracted details
            self.update_csv_with_specific_row(
                csv_file,
                app_number,
                details['col1_headers'],
                details['col3_headers'],
                details['col1_data'],
                details['col3_data']
            )

        except Exception as e:
            self.log_message(f"Error processing details for application {app_number}: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")
            raise

    def process_applications(self):
        """Process each application in the table with retry mechanism."""
        try:
            self.log_message("Processing applications...", "STATUS")

            # Check if we have application numbers to process
            if not self.application_numbers:
                self.log_message("No application numbers found to process", "WARNING")
                return

            # Create CSV file for storing results
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_file = os.path.join(self.folder_var.get(), f"patent_data_{timestamp}.csv")

            # Get selected columns for CSV headers
            selected_columns = [
                header for header, var in self.column_vars.items() if var.get()
            ]

            # Initialize CSV with headers and initial data
            with open(csv_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(selected_columns)  # Use selected columns
                if self.csv_data:
                    # Write data, ensuring only selected columns are written
                    for row in self.csv_data:
                        writer.writerow(
                            [row[col] for col in selected_columns if col in row]
                        )  #
                else:
                    # If no CSV data, at least write application numbers
                    for app_num in self.application_numbers:
                        writer.writerow([app_num] + [""] * (len(selected_columns) - 1))

            self.log_message(f"Created initial CSV file: {csv_file}", "SUCCESS")
            self.log_message(f"Processing {len(self.application_numbers)} applications...", "STATUS")

            # Process applications with retry mechanism
            max_retries = 2
            failed_applications = set()
            processed_numbers = set()

            for attempt in range(max_retries + 1):
                current_failed_applications = set()

                # Make a copy of application numbers
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
                        self.log_message(f"Processing application {app_number} (Attempt {attempt + 1})", "STATUS")
                        self.process_application_details(app_number, csv_file)
                        processed_numbers.add(app_number)
                        self.log_message(f"Successfully processed {app_number}", "SUCCESS")
                    except Exception as e:
                        self.log_message(f"Attempt {attempt + 1}: Error processing application {app_number}: {str(e)}", "ERROR")
                        current_failed_applications.add(app_number)

                # Update failed applications set
                if attempt < max_retries:
                    failed_applications = current_failed_applications

                    # Break if no more failed applications
                    if not failed_applications:
                        break

                    # Prepare for next retry
                    self.application_numbers = list(failed_applications)
                    self.log_message(f"Retry attempt {attempt + 2} for {len(self.application_numbers)} applications", "STATUS")
                else:
                    # Final attempt - log permanent failures
                    if current_failed_applications:
                        self.log_message("\n--- PERMANENT FAILURE APPLICATIONS ---", "ERROR")
                        for app in current_failed_applications:
                            self.log_message(f"Permanent Error: Unable to process application {app}", "ERROR")

            self.log_message(f"Processed {len(processed_numbers)} applications successfully", "SUCCESS")
            if failed_applications:
                self.log_message(f"Failed to process {len(failed_applications)} applications", "WARNING")

            # Store the CSV file path for later use
            self.csv_file_path = csv_file

        except Exception as e:
            self.log_message(f"Error in application processing: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

    def generate_csv(self):
        """Generate final CSV file"""
        try:
            # Reuse the csv_file_path if available, otherwise construct a new one
            if hasattr(self, "csv_file_path"):
                csv_file = self.csv_file_path
            else:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                csv_file = os.path.join(self.folder_var.get(), f"patent_data_{timestamp}.csv")

            # Check if the file exists and is not empty
            if os.path.exists(csv_file) and os.path.getsize(csv_file) > 0:
                self.log_message(f"Final CSV file generated at: {csv_file}", "SUCCESS")
                # Open the directory containing the file
                if os.name == "nt":  # Windows
                    os.startfile(os.path.dirname(csv_file))
                elif os.name == "posix":  # Linux or macOS
                    subprocess.Popen(["xdg-open", os.path.dirname(csv_file)])
            else:
                self.log_message("No data was extracted, therefore no CSV file was generated.", "WARNING")
        except Exception as e:
            self.log_message(f"Error generating CSV: {str(e)}", "ERROR")
            self.log_message(traceback.format_exc(), "ERROR")

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
            self.start_btn.config(state=tk.NORMAL)

    def log_message(self, message, tag=""):
        """Add message to log console with timestamp (thread-safe)"""
        timestamp = datetime.now().strftime("[%H:%M:%S] ")
        full_message = timestamp + message
        print(f"log_message: {full_message}, tag={tag}")  # Debug print
        self.root.after(0, self._update_log, full_message, tag)

    def _update_log(self, full_message, tag):
        """Update log text widget (must be called from main thread)"""
        print(f"_update_log: {full_message}, tag={tag}")  # Debug print
        try:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, full_message + "\n", tag)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        except Exception as e:
            print(f"_update_log error: {e}")  # Print any errors here        
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
    root.mainloop()
            
            
            
            
            
