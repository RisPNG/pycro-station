import os
import time
import threading
from typing import List, Tuple, Dict, Any
from datetime import datetime

# PySide6 & qfluentwidgets UI imports
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QTextEdit,
    QWidget,
    QSizePolicy
)
from qfluentwidgets import PrimaryPushButton, PushButton, MessageBox

# Excel imports
import openpyxl

# Selenium imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains


class MainWidget(QWidget):
    """
    The main UI Widget for the Pycro.
    Handles layout, file selection, and spawning the background thread for processing.
    """
    log_message = Signal(str)
    processing_done = Signal(int, int, str)

    def __init__(self):
        super().__init__()
        self.setObjectName("ssi_widget")
        self.processor = None

        self._build_ui()
        self._connect_signals()

    # UI Construction
    def _build_ui(self):
        self.desc_label = QLabel("", self)
        self.desc_label.setWordWrap(True)
        self.desc_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.desc_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.desc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.desc_label.setStyleSheet(
            "color: #dcdcdc; background: transparent; padding: 6px; "
            "border: 1px solid #3a3a3a; border-radius: 6px;"
        )
        self.set_long_description("")

        self.select_btn = PrimaryPushButton("Select Excel Files", self)
        self.run_btn = PrimaryPushButton("Import to Smartsheet", self)
        self.stop_btn = PushButton("Stop", self)
        self.stop_btn.setEnabled(False)

        self.files_label = QLabel("Selected files", self)
        self.files_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.files_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.logs_label = QLabel("Process logs", self)
        self.logs_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.logs_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.files_box = QTextEdit(self)
        self.files_box.setReadOnly(True)
        self.files_box.setPlaceholderText("Selected files will appear here...")
        self.files_box.setStyleSheet(
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Live process log will appear here...")
        self.log_box.setStyleSheet(
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        main_layout.addWidget(self.desc_label, 1)

        row1_layout = QHBoxLayout()
        row1_layout.addStretch(1)
        row1_layout.addWidget(self.select_btn, 1)
        row1_layout.addStretch(1)
        main_layout.addLayout(row1_layout, 0)

        row2_layout = QHBoxLayout()
        row2_layout.addStretch(1)
        row2_layout.addWidget(self.run_btn, 1)
        row2_layout.addWidget(self.stop_btn, 1)
        row2_layout.addStretch(1)
        main_layout.addLayout(row2_layout, 0)

        row3_layout = QHBoxLayout()
        row3_layout.addWidget(self.files_label, 1)
        row3_layout.addWidget(self.logs_label, 1)
        main_layout.addLayout(row3_layout, 0)

        row4_layout = QHBoxLayout()
        row4_layout.addWidget(self.files_box, 1)
        row4_layout.addWidget(self.log_box, 1)
        main_layout.addLayout(row4_layout, 4)

    def set_long_description(self, text: str):
        clean = (text or "").strip()
        if clean:
            self.desc_label.setText(clean)
            self.desc_label.show()
        else:
            self.desc_label.clear()
            self.desc_label.hide()

    def _connect_signals(self):
        self.select_btn.clicked.connect(self.select_files)
        self.run_btn.clicked.connect(self.run_process)
        self.stop_btn.clicked.connect(self.stop_process)
        self.log_message.connect(self.append_log)
        self.processing_done.connect(self.on_processing_done)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Excel Files", "", "Excel Files (*.xlsx)")
        if files:
            self.files_box.setPlainText("\n".join(files))
        else:
            self.files_box.clear()

    def _selected_files(self) -> List[str]:
        text = self.files_box.toPlainText().strip()
        if not text:
            return[]
        return[line for line in text.split("\n") if line.strip()]

    def run_process(self):
        files = self._selected_files()
        if not files:
            MessageBox("Warning", "No files selected to process.", self).exec()
            return

        self.log_box.clear()
        self.log_message.emit("Process started...")

        self.run_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        self.processor = SmartsheetImporterProcessor(self.log_message.emit)

        def worker():
            ok_total, fail_total = 0, 0
            try:
                ok_total, fail_total = self.processor.process(files)
            except Exception as e:
                self.log_message.emit(f"ERROR: {e}")
            self.processing_done.emit(ok_total, fail_total, "")

        threading.Thread(target=worker, daemon=True).start()

    def stop_process(self):
        if self.processor:
            self.processor.stop()
            self.log_message.emit("🛑 Stop requested! Halting after current operation finishes...")
            self.stop_btn.setEnabled(False)

    def append_log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def on_processing_done(self, ok: int, fail: int, out_path: str):
        self.log_message.emit(f"Completed: {ok} success, {fail} failed.")

        self.run_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

        title = "Processing complete" if fail == 0 else "Processing finished with issues"
        lines =[f"Success (Rows): {ok}", f"Failed (Rows): {fail}"]

        msg = MessageBox(title, "\n".join(lines), self)
        msg.yesButton.setText("OK")
        msg.cancelButton.hide()
        msg.exec()


def get_widget():
    return MainWidget()


class SmartsheetImporterProcessor:
    def __init__(self, log_emit):
        self.log_emit = log_emit
        self.success_count = 0
        self.fail_count = 0
        self.is_stopped = False

    def log(self, msg: str):
        stamp = f"[{datetime.now().strftime('%H:%M:%S')}]"
        self.log_emit(f"{stamp} {msg}")

    def stop(self):
        self.is_stopped = True

    def process(self, file_paths: List[str]) -> Tuple[int, int]:
        for idx, path in enumerate(file_paths, start=1):
            if self.is_stopped:
                break

            file_name = os.path.basename(path)
            self.log(f"Loading data from: {file_name}")
            try:
                data = self.load_excel_data(path)
                self.log(f"Found {len(data)} rows of data. Starting automation...")
                self.automate_data_entry(data)
                self.log(f"Finished processing {file_name}.")
            except Exception as e:
                self.log(f"Error processing {file_name}: {e}")

        if self.is_stopped:
            self.log("Process was stopped by user.")

        return self.success_count, self.fail_count

    def load_excel_data(self, file_path: str) -> List[Dict[str, Any]]:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        headers =[cell.value for cell in sheet[1]]
        data =[
            {headers[col_idx]: row[col_idx] for col_idx in range(len(headers))}
            for row in sheet.iter_rows(min_row=2, values_only=True)
        ]
        return data

    def automate_data_entry(self, data: List[Dict[str, Any]]):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("detach", True)

        driver = webdriver.Chrome(options=chrome_options)
        try:
            driver.get("https://app.smartsheet.com/dynamicview/views/b3bbae0c-452a-489e-8c8a-8f58782e6784")

            try:
                WebDriverWait(driver, 300).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "view-description-container"))
                )
                self.log("Login detected. Starting the automation process.")
            except Exception as e:
                self.log(f"Login timeout or error: {e}")
                return

            row_index = 0
            for row in data:
                if self.is_stopped:
                    break

                retry_attempts = 10
                row_successful = False

                for attempt in range(retry_attempts):
                    if self.is_stopped:
                        break

                    try:
                        row_height = 40
                        scroll_increment = 600
                        container_xpath = ".//div[contains(@class, 'ReactVirtualized__Grid__innerScrollContainer')]"

                        # FIX 1: New row XPath - Smartsheet removed aria-label="row"
                        row_xpath = './/div[@role="row" and contains(@class, "data-row")]'

                        skip = False

                        while True:
                            if self.is_stopped:
                                break

                            try:
                                container = WebDriverWait(driver, 10).until(
                                    EC.presence_of_element_located((By.XPATH, container_xpath))
                                )
                                rows = container.find_elements(By.XPATH, row_xpath)

                                if not rows:
                                    self.log("No rows found, waiting for table to load...")
                                    time.sleep(1)
                                    continue

                                first_row = rows[0]
                                current_scroll_top = int(first_row.value_of_css_property("top").replace("px", ""))
                                target_row_top = row_index * row_height
                                self.log(f"Target Top: {target_row_top}px, Current Top: {current_scroll_top}px.")

                                if current_scroll_top <= target_row_top <= current_scroll_top + scroll_increment:
                                    rows = container.find_elements(By.XPATH, row_xpath)
                                    visible_row_index = (target_row_top - current_scroll_top) // row_height

                                    if visible_row_index < len(rows):
                                        current_row = rows[visible_row_index]
                                        try:
                                            driver.execute_script(
                                                "arguments[0].scrollIntoView({block: 'start', inline: 'nearest'});",
                                                current_row
                                            )
                                            rows = container.find_elements(By.XPATH, row_xpath)
                                            current_row = rows[visible_row_index]

                                            if (row.get("Skip") or "").upper() == "YES":
                                                self.log("Skipping row...")
                                                skip = True
                                                break

                                            try:
                                                current_row.click()
                                            except StaleElementReferenceException:
                                                self.log("Stale element on click, retrying...")
                                                continue
                                            self.log(f"Clicked row at index {row_index}.")
                                            break
                                        except Exception as e:
                                            self.log(f"Click attempt failed: {e}. Retrying...")
                                            time.sleep(1)
                                    else:
                                        self.log(f"Row index {row_index} is not visible. Loading new records.")
                                        try:
                                            new_button = WebDriverWait(driver, 10).until(
                                                EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "New")]'))
                                            )
                                            new_button.click()
                                        except StaleElementReferenceException:
                                            self.log("Stale element on new_button click, retrying...")
                                            continue
                                        break
                                else:
                                    container_height = container.size['height']
                                    scrollTop = driver.execute_script("return arguments[0].scrollTop", container)
                                    scrollHeight = driver.execute_script("return arguments[0].scrollHeight", container)

                                    if scrollTop + container_height >= scrollHeight:
                                        self.log(f"Reached bottom. Attempting to click row {row_index} directly.")
                                        rows = container.find_elements(By.XPATH, row_xpath)
                                        visible_row_index = (target_row_top - current_scroll_top) // row_height

                                        if visible_row_index < len(rows):
                                            current_row = rows[visible_row_index]
                                            driver.execute_script(
                                                "arguments[0].scrollIntoView({block: 'start', inline: 'nearest'});",
                                                current_row
                                            )
                                            rows = container.find_elements(By.XPATH, row_xpath)
                                            current_row = rows[visible_row_index]

                                            if (row.get("Skip") or "").upper() == "YES":
                                                self.log("Skipping row...")
                                                skip = True
                                                break

                                            try:
                                                current_row.click()
                                            except StaleElementReferenceException:
                                                self.log("Stale element (bottom) on click, retrying...")
                                                continue
                                            self.log(f"Clicked row at index {row_index} at bottom.")
                                            break
                                        else:
                                            self.log(f"Row index {row_index} is still not visible even at bottom. Loading new records.")
                                            try:
                                                new_button = WebDriverWait(driver, 10).until(
                                                    EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "New")]'))
                                                )
                                                new_button.click()
                                            except StaleElementReferenceException:
                                                self.log("Stale element on new_button click, retrying...")
                                                continue
                                            break
                                    else:
                                        self.log(f"Scrolling down to find row {row_index}...")
                                        driver.execute_script("arguments[0].scrollBy(0, 300);", container)
                                        time.sleep(0.3)
                                        continue

                            except TimeoutException:
                                self.log("Timeout waiting for elements. Check if page structure changed.")
                                return
                            except StaleElementReferenceException:
                                self.log("Stale element in row scroll loop, retrying...")
                                continue
                            except Exception as e:
                                self.log(f"An error occurred: {e}")
                                return

                        # Input fill and save validation code
                        if not skip and not self.is_stopped:
                            fill_save_attempts = 10
                            for fill_save_try in range(1, fill_save_attempts + 1):
                                if self.is_stopped:
                                    break

                                # FIX 2: Wait for side panel structural class, not randomized hash
                                WebDriverWait(driver, 10).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, "section.r-side-panel"))
                                )

                                for field_name, value in row.items():
                                    if (str(field_name).upper() in ["SKIP"] and str(value).upper() in ["YES"]) or value is None or str(field_name).upper() in["PRODUCT CATEGORY"]:
                                        self.log(f"Skipping field: {field_name}, Value: {value}")
                                        continue

                                    try:
                                        # FIX 3: Input Field Locator using stable 'form-field-wrapper' class
                                        robust_input_xpath = f'//b[contains(text(), "{field_name}")]/ancestor::div[contains(@class, "form-field-wrapper")]//input[not(@type="file")]'
                                        input_element = WebDriverWait(driver, 10).until(
                                            EC.presence_of_element_located((By.XPATH, robust_input_xpath))
                                        )

                                        self.log(f"Processing field: {field_name}, Value: {value}")
                                        driver.execute_script(
                                            "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", input_element
                                        )
                                        WebDriverWait(driver, 20).until(EC.visibility_of(input_element))
                                        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(input_element))

                                        if field_name in["MCO with Capability Limitation", "Country with Limitation", "MCO Limitation"]:
                                            values =[v.strip() for v in str(value).split(",")]
                                            try:
                                                input_element.clear()
                                                # FIX 4: Clear button container
                                                cont = input_element.find_element(By.XPATH, "./ancestor::div[contains(@class, 'form-field-wrapper')]")
                                                clear_button = WebDriverWait(cont, 3).until(
                                                    EC.element_to_be_clickable((By.XPATH, './/button[@aria-label="Clear selection"]'))
                                                )
                                                clear_button.click()
                                            except (TimeoutException, StaleElementReferenceException):
                                                pass

                                            for val in values:
                                                multi_retry = 3
                                                for i in range(multi_retry):
                                                    try:
                                                        self.log(f"Processing field: {field_name}, Value: {val}")
                                                        input_element.send_keys(val)
                                                        input_element.send_keys(Keys.ENTER)
                                                        input_element.send_keys(Keys.ESCAPE)
                                                        break
                                                    except StaleElementReferenceException:
                                                        self.log("Stale element: re-finding input for multivalue...")
                                                        input_element = WebDriverWait(driver, 10).until(
                                                            EC.presence_of_element_located((By.XPATH, robust_input_xpath))
                                                        )
                                                        continue

                                        elif field_name != "Fcty Group":
                                            if field_name == "Fcty Contact List":
                                                for i in range(3):
                                                    try:
                                                        try:
                                                            input_element.clear()
                                                            # FIX 5: Contact Clear button
                                                            clear_contact_button = WebDriverWait(driver, 3).until(
                                                                EC.element_to_be_clickable((
                                                                    By.CSS_SELECTOR,
                                                                    "div.contact-picker__clear-indicator",
                                                                ))
                                                            )
                                                            clear_contact_button.click()
                                                        except (TimeoutException, StaleElementReferenceException):
                                                            pass
                                                        time.sleep(1)
                                                        input_element.send_keys(str(value))
                                                        input_element.send_keys(Keys.ENTER)
                                                        break
                                                    except StaleElementReferenceException:
                                                        input_element = WebDriverWait(driver, 10).until(
                                                            EC.presence_of_element_located((By.XPATH, robust_input_xpath))
                                                        )
                                                        continue
                                            elif field_name in["Gender Dimension", "Dev Team", "Merchandising Classification"]:
                                                for i in range(3):
                                                    try:
                                                        try:
                                                            input_element.clear()
                                                        except (TimeoutException, StaleElementReferenceException):
                                                            pass
                                                        time.sleep(1)
                                                        input_element.send_keys(str(value))
                                                        dropdown_options = WebDriverWait(driver, 10).until(
                                                            EC.presence_of_all_elements_located((By.XPATH, '//div[@role="option"]'))
                                                        )
                                                        found = False
                                                        for option in dropdown_options:
                                                            if option.text.strip().upper() == value.strip().upper():
                                                                driver.execute_script("arguments[0].click();", option)
                                                                found = True
                                                                break
                                                        if not found:
                                                            raise Exception(f"Could not find option matching '{value}'")
                                                        break
                                                    except StaleElementReferenceException:
                                                        input_element = WebDriverWait(driver, 10).until(
                                                            EC.presence_of_element_located((By.XPATH, robust_input_xpath))
                                                        )
                                                        continue
                                            else:
                                                for i in range(3):
                                                    try:
                                                        input_element.clear()
                                                        input_element.send_keys(str(value))
                                                        break
                                                    except StaleElementReferenceException:
                                                        input_element = WebDriverWait(driver, 10).until(
                                                            EC.presence_of_element_located((By.XPATH, robust_input_xpath))
                                                        )
                                                        continue
                                    except Exception as e:
                                        self.log(f"Error processing field '{field_name}': {e}")
                                        continue

                                time.sleep(1)
                                self.log("Saving...")
                                try:
                                    save_button = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Save")]'))
                                    )
                                    save_button.click()
                                except StaleElementReferenceException:
                                    save_button = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Save")]'))
                                    )
                                    save_button.click()

                                time.sleep(1)
                                try:
                                    close_button = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.close-button'))
                                    )
                                    close_button.click()
                                except StaleElementReferenceException:
                                    close_button = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.close-button'))
                                    )
                                    close_button.click()

                                time.sleep(1)

                                detail_modal_still_open = False
                                try:
                                    WebDriverWait(driver, 3).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, "section.r-side-panel"))
                                    )
                                    detail_modal_still_open = True
                                except TimeoutException:
                                    detail_modal_still_open = False

                                if not detail_modal_still_open:
                                    row_successful = True
                                    break
                                else:
                                    try:
                                        error_msgs = driver.find_elements(By.XPATH, "//div[contains(@class,'error') or contains(@class,'invalid')]")
                                        if not error_msgs:
                                            error_msgs = driver.find_elements(By.XPATH, "//p[contains(@class,'error') or contains(@class,'warning')]")
                                        msgs =[e.text for e in error_msgs if e.text.strip()]
                                        if msgs:
                                            self.log(f"Save failed, validation error: {' | '.join(msgs)} (retry {fill_save_try})")
                                        else:
                                            self.log(f"Save failed, modal still open (retry {fill_save_try})")
                                    except Exception as err:
                                        self.log(f"Save failed (retry {fill_save_try}), error finding validation message: {err}")

                                    if fill_save_try == fill_save_attempts:
                                        self.log("Giving up after max retries due to validation error.")
                                        skip = True
                                    else:
                                        self.log("Retrying: Re-filling form and Save...")
                                    time.sleep(1)

                        if skip or row_successful:
                            self.success_count += 1
                        else:
                            self.fail_count += 1

                        row_index += 1
                        break

                    except StaleElementReferenceException:
                        self.log("Stale element in row processing, retrying row...")
                        time.sleep(1)
                        continue
                    except Exception as e:
                        self.log(f"Error processing row {row}, attempt {attempt + 1}: {e}")
                        if attempt == retry_attempts - 1:
                            self.log(f"Failed to process row after {retry_attempts} attempts: {row}")
                            self.fail_count += 1
                        time.sleep(5)

        finally:
            self.log("Process complete. The browser window has been left open for manual verification.")