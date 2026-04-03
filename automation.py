from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import sys
from typing import Iterable, Optional, Tuple

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    InvalidElementStateException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


CRM_URL = "https://crm.intl.hopeedu.com/#/home"
EXCEL_FILE = "data.xlsx"
DEFAULT_TIMEOUT = 20
SHORT_TIMEOUT = 8
MAX_RETRIES = 3
CONVERT_STATUSES = {"Completed-Successful", "Completed-Not Interested"}


@dataclass
class ProcessResult:
    phone: str
    success: bool
    message: str


def get_status(remark: str) -> str:
    text = str(remark).strip().lower()

    if any(token in text for token in ["wrong number", "not in service", "invalid", "cannot be reached"]):
        return "Completed-Invalid"
    if any(token in text for token in ["not interested", "nis"]):
        return "Completed-Not Interested"
    if any(token in text for token in ["no reply", "no answer", "hung up"]):
        return "Completed-No Reply"
    if any(token in text for token in ["picked up", "considering", "interested", "shared open day"]):
        return "Completed-Successful"
    return "Completed-No Reply"


def normalize_phone(value: object) -> str:
    if pd.isna(value):
        return ""
    raw = str(value).strip()
    if raw.endswith(".0"):
        raw = raw[:-2]
    return raw


def load_rows(excel_path: str) -> Iterable[Tuple[str, str]]:
    df = pd.read_excel(excel_path)
    required_columns = {"Phone No.", "Remarks"}
    missing = required_columns.difference(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in Excel: {', '.join(sorted(missing))}")

    for _, row in df.iterrows():
        phone = normalize_phone(row["Phone No."])
        remark = "" if pd.isna(row["Remarks"]) else str(row["Remarks"]).strip()
        if not phone:
            continue
        yield phone, remark


class CRMTaskLogger:
    def __init__(self) -> None:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-infobars")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options,
        )
        self.wait = WebDriverWait(self.driver, DEFAULT_TIMEOUT)
        self.short_wait = WebDriverWait(self.driver, SHORT_TIMEOUT)
        self.today = datetime.today().strftime("%Y-%m-%d")
        self.select_all_key = Keys.COMMAND if sys.platform == "darwin" else Keys.CONTROL

    def open(self) -> None:
        self.driver.get(CRM_URL)
        input("Login completed, press Enter to start... ")
        self.wait_for_page_ready()

    def close(self) -> None:
        self.driver.quit()

    def wait_for_page_ready(self) -> None:
        self.wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    @staticmethod
    def is_visible(element: WebElement) -> bool:
        try:
            return element.is_displayed() and element.size.get("height", 0) > 0 and element.size.get("width", 0) > 0
        except StaleElementReferenceException:
            return False

    def wait_for_visible(self, by: By, locator: str, timeout: Optional[int] = None) -> WebElement:
        wait = self.wait if timeout is None else WebDriverWait(self.driver, timeout)
        return wait.until(lambda d: next((el for el in d.find_elements(by, locator) if self.is_visible(el)), False))

    def scroll_into_view(self, element: WebElement) -> None:
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
            element,
        )

    def clear_input(self, element: WebElement) -> None:
        try:
            element.clear()
        except InvalidElementStateException:
            pass
        element.send_keys(self.select_all_key, "a")
        element.send_keys(Keys.DELETE)

    def safe_click(self, by: By, locator: str, description: str, timeout: Optional[int] = None) -> WebElement:
        last_error = None
        for _ in range(MAX_RETRIES):
            try:
                element = self.wait_for_visible(by, locator, timeout=timeout)
                self.scroll_into_view(element)
                self.wait.until(lambda d: element.is_enabled())
                try:
                    element.click()
                except (ElementClickInterceptedException, InvalidElementStateException, WebDriverException):
                    self.driver.execute_script("arguments[0].click();", element)
                return element
            except (TimeoutException, StaleElementReferenceException, WebDriverException) as exc:
                last_error = exc
        raise TimeoutException(f"Unable to click {description}: {last_error}")

    def safe_type(
        self,
        by: By,
        locator: str,
        value: str,
        description: str,
        clear_first: bool = True,
        timeout: Optional[int] = None,
    ) -> WebElement:
        last_error = None
        for _ in range(MAX_RETRIES):
            try:
                element = self.wait_for_visible(by, locator, timeout=timeout)
                self.scroll_into_view(element)
                self.wait.until(lambda d: element.is_enabled())
                if clear_first:
                    self.clear_input(element)
                element.send_keys(value)
                return element
            except (StaleElementReferenceException, InvalidElementStateException, WebDriverException) as exc:
                last_error = exc
        raise TimeoutException(f"Unable to type into {description}: {last_error}")

    def wait_for_loading_overlay(self) -> None:
        overlay_xpath = "//div[contains(@class,'el-loading-mask') and not(contains(@style,'display: none'))]"
        try:
            self.short_wait.until_not(
                lambda d: any(self.is_visible(el) for el in d.find_elements(By.XPATH, overlay_xpath))
            )
        except TimeoutException:
            pass

    def close_any_open_dropdown(self) -> None:
        try:
            self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        except Exception:
            pass

    def wait_for_global_search(self) -> None:
        self.wait_for_visible(
            By.XPATH,
            "//input[contains(@class,'el-input__inner') and (@placeholder='Search' or contains(@placeholder,'Search'))] | //i[contains(@class,'el-icon-search')]",
        )

    def reset_to_home(self) -> None:
        self.close_any_open_dropdown()
        self.driver.get(CRM_URL)
        self.wait_for_page_ready()
        self.wait_for_loading_overlay()
        self.wait_for_global_search()

    def search_student(self, phone: str) -> Optional[str]:
        self.wait_for_loading_overlay()
        self.close_any_open_dropdown()

        search_input_xpath = (
            "(//input[contains(@class,'el-input__inner') and not(@readonly) "
            "and (@placeholder='Search' or contains(@placeholder,'Search'))])[1]"
        )

        try:
            search_input = self.wait_for_visible(By.XPATH, search_input_xpath)
        except TimeoutException:
            self.safe_click(By.XPATH, "(//i[contains(@class,'el-icon-search')])[1]", "search icon", timeout=SHORT_TIMEOUT)
            search_input = self.wait_for_visible(By.XPATH, search_input_xpath)

        self.safe_type(By.XPATH, search_input_xpath, phone, "global search input")
        search_input.send_keys(Keys.ENTER)
        self.wait_for_loading_overlay()

        if self.has_student_rows():
            print("Found student under current results")
            return "lead"

        print("No lead result, trying Opportunity tab...")
        if self.try_open_result_tab("Opportunity") and self.has_student_rows():
            print("Found student under Opportunity")
            return "opportunity"

        print("No opportunity result, checking Leads tab...")
        if self.try_open_result_tab("Leads") and self.has_student_rows():
            print("Found student under Leads")
            return "lead"

        return None

    def try_open_result_tab(self, label: str) -> bool:
        tag_xpath = f"//span[contains(@class,'tagShow') and contains(normalize-space(),'{label}')]"
        try:
            tag = self.wait_for_visible(By.XPATH, tag_xpath, timeout=SHORT_TIMEOUT)
        except TimeoutException:
            return False

        try:
            tag.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", tag)

        self.wait_for_loading_overlay()

        row_xpath = (
            "//table[contains(@class,'el-table')]//tr[contains(@class,'el-table__row')]"
            "[.//button[contains(@class,'el-button--text') and not(.//span[normalize-space()='Edit'])]]"
        )

        try:
            self.wait.until(
                lambda d: any(self.is_visible(el) for el in d.find_elements(By.XPATH, row_xpath))
                or any(self.is_visible(el) for el in d.find_elements(By.XPATH, "//*[contains(.,'No data')]"))
                or any(self.is_visible(el) for el in d.find_elements(By.XPATH, "//*[contains(.,'No student found')]"))
            )
        except TimeoutException:
            pass

        return True

    def has_student_rows(self) -> bool:
        row_xpath = (
            "//table[contains(@class,'el-table')]//tr[contains(@class,'el-table__row')]"
            "[.//button[contains(@class,'el-button--text') and not(.//span[normalize-space()='Edit'])]]"
        )
        rows = [el for el in self.driver.find_elements(By.XPATH, row_xpath) if self.is_visible(el)]
        return len(rows) > 0

    def open_first_student_result(self, phone: str) -> None:
        row_xpath = (
            "(//table[contains(@class,'el-table')]//tr[contains(@class,'el-table__row')]"
            "[.//button[contains(@class,'el-button--text') and not(.//span[normalize-space()='Edit'])]])[1]"
        )
        row = self.wait_for_visible(By.XPATH, row_xpath)
        self.scroll_into_view(row)

        student_button_xpath = (
            ".//button[contains(@class,'el-button--text') and contains(@class,'el-button--medium') "
            "and not(.//span[normalize-space()='Edit'])]"
        )
        student_button = next(
            (el for el in row.find_elements(By.XPATH, student_button_xpath) if self.is_visible(el)),
            None
        )

        if student_button is None:
            raise TimeoutException(f"Unable to find student button for {phone}")

        self.scroll_into_view(student_button)

        try:
            student_button.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", student_button)

        self.wait_for_loading_overlay()
        self.wait.until(
            lambda d: any(
                self.is_visible(el)
                for el in d.find_elements(By.XPATH, "//div[@id='tab-third'] | //span[normalize-space()='Related']")
            )
        )

    def click_related_tab(self) -> None:
        related_xpath = "//div[@id='tab-third'] | //span[normalize-space()='Related']/ancestor::*[self::div or self::button][1]"
        self.safe_click(By.XPATH, related_xpath, "Related tab")
        self.wait_for_loading_overlay()

    def click_new_task(self) -> None:
        new_task_xpath = "//button[.//span[normalize-space()='New Task']]"
        self.safe_click(By.XPATH, new_task_xpath, "New Task button")
        self.wait_for_loading_overlay()
        self.wait_for_visible(By.XPATH, "//div[contains(@class,'leads_right') and .//h1[normalize-space()='Create New Task']]")

    def get_form_container(self) -> WebElement:
        container_xpath = "//div[contains(@class,'leads_right') and .//h1[normalize-space()='Create New Task']]"
        return self.wait_for_visible(By.XPATH, container_xpath)

    def find_field_container(self, form: WebElement, label_text: str) -> WebElement:
        label_xpath = (
            f".//label[normalize-space()='{label_text}' or contains(normalize-space(),'{label_text}')]"
            "/ancestor::*[contains(@class,'el-form-item')][1]"
        )
        container = next((el for el in form.find_elements(By.XPATH, label_xpath) if self.is_visible(el)), None)
        if container is None:
            raise NoSuchElementException(f"Field container not found for label: {label_text}")
        return container

    def type_in_field(self, form: WebElement, label_text: str, value: str) -> None:
        container = self.find_field_container(form, label_text)
        input_element = next(
            (el for el in container.find_elements(By.XPATH, ".//textarea | .//input[not(@readonly)]") if self.is_visible(el)),
            None,
        )
        if input_element is None:
            raise NoSuchElementException(f"Input not found for field: {label_text}")
        self.scroll_into_view(input_element)
        self.clear_input(input_element)
        input_element.send_keys(value)

    def open_dropdown(self, container: WebElement) -> None:
        trigger = next(
            (
                el for el in container.find_elements(
                    By.XPATH,
                    ".//*[contains(@class,'el-select') or contains(@class,'el-input')][1]",
                ) if self.is_visible(el)
            ),
            None,
        )
        if trigger is None:
            raise NoSuchElementException("Dropdown trigger not found")

        self.scroll_into_view(trigger)
        try:
            trigger.click()
        except (ElementClickInterceptedException, InvalidElementStateException, WebDriverException):
            self.driver.execute_script("arguments[0].click();", trigger)

        self.wait_for_loading_overlay()

    def select_visible_dropdown_option(self, option_text: str) -> None:
        option_xpath = (
            f"//div[contains(@class,'el-select-dropdown') and not(contains(@style,'display: none'))]"
            f"//li[not(contains(@class,'is-disabled'))]//span[normalize-space()='{option_text}']"
        )
        option = self.wait_for_visible(By.XPATH, option_xpath)
        self.scroll_into_view(option)
        try:
            option.click()
        except (ElementClickInterceptedException, InvalidElementStateException, WebDriverException):
            self.driver.execute_script("arguments[0].click();", option)

        self.wait_for_loading_overlay()
        self.close_any_open_dropdown()

    def select_dropdown_value(self, form: WebElement, label_text: str, option_text: str) -> None:
        container = self.find_field_container(form, label_text)
        self.open_dropdown(container)
        self.select_visible_dropdown_option(option_text)

    def set_due_date_today(self, form: WebElement) -> None:
        container = self.find_field_container(form, "Due Date")
        input_element = next((el for el in container.find_elements(By.XPATH, ".//input") if self.is_visible(el)), None)
        if input_element is None:
            raise NoSuchElementException("Due Date input not found")

        self.scroll_into_view(input_element)
        try:
            input_element.click()
        except WebDriverException:
            self.driver.execute_script("arguments[0].click();", input_element)

        today_cell_xpath = (
            "//div[contains(@class,'el-picker-panel') and not(contains(@style,'display: none'))]"
            "//td[contains(@class,'today') and not(contains(@class,'disabled'))]"
        )

        try:
            self.safe_click(By.XPATH, today_cell_xpath, "today date cell", timeout=SHORT_TIMEOUT)
        except TimeoutException:
            self.clear_input(input_element)
            input_element.send_keys(self.today)
            input_element.send_keys(Keys.TAB)

        self.wait_for_loading_overlay()

    def set_additional_information_status(self, form: WebElement, status_text: str) -> None:
        print(f"[STATUS] Looking for Additional Information -> {status_text}")

        status_order = [
            "Not started",
            "In Progress",
            "Completed-Successful",
            "Completed-Not Interested",
            "Completed-No Reply",
            "Completed-Invalid",
        ]

        if status_text not in status_order:
            raise NoSuchElementException(f"Unknown status text: {status_text}")

        target_index = status_order.index(status_text)

        additional_tab_xpath = (
            "//a[@title='Additional Information' or .//span[contains(normalize-space(),'Additional In')]]"
        )
        additional_tab = self.wait_for_visible(By.XPATH, additional_tab_xpath)
        try:
            additional_tab.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", additional_tab)

        self.wait_for_loading_overlay()

        section_xpath = (
            "//h2[normalize-space()='Additional Information']"
            "/ancestor::div[contains(@class,'section_list_div2')][1]"
        )
        section = self.wait_for_visible(By.XPATH, section_xpath)

        scroll_container = self.wait_for_visible(By.XPATH, "//div[contains(@class,'scrollList')]")
        self.driver.execute_script(
            "arguments[0].scrollTop = arguments[1].offsetTop - 120;",
            scroll_container,
            section
        )

        self.wait_for_loading_overlay()

        status_item_xpath = (
            ".//label[normalize-space()='Status']"
            "/ancestor::div[contains(@class,'el-form-item')][1]"
        )
        status_item = next(
            (el for el in section.find_elements(By.XPATH, status_item_xpath) if self.is_visible(el)),
            None
        )
        if status_item is None:
            raise NoSuchElementException("Status field inside Additional Information not found")

        trigger = next(
            (
                el for el in status_item.find_elements(
                    By.XPATH,
                    ".//div[contains(@class,'el-input') and contains(@class,'el-input--suffix')]"
                )
                if self.is_visible(el)
            ),
            None
        )
        if trigger is None:
            raise NoSuchElementException("Status trigger not found")

        value_input = next(
            (
                el for el in status_item.find_elements(
                    By.XPATH,
                    ".//input[contains(@class,'el-input__inner')]"
                )
                if self.is_visible(el)
            ),
            None
        )
        if value_input is None:
            raise NoSuchElementException("Status input not found")

        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", trigger)
        self.wait_for_loading_overlay()

        actions = ActionChains(self.driver)

        opened = False
        for attempt in range(4):
            try:
                actions.move_to_element(trigger).pause(0.2).click(trigger).perform()
                self.wait_for_loading_overlay()

                try:
                    value_input.send_keys(Keys.ARROW_DOWN)
                    self.wait_for_loading_overlay()
                except Exception:
                    pass

                dropdown = next(
                    (
                        el for el in status_item.find_elements(
                            By.XPATH,
                            ".//div[contains(@class,'el-select-dropdown') and not(contains(@style,'display: none'))]"
                        )
                        if self.is_visible(el)
                    ),
                    None
                )

                if dropdown is not None:
                    opened = True
                    print("[STATUS] Dropdown opened")
                    break
            except Exception as exc:
                print(f"[STATUS] Open attempt {attempt + 1} failed: {exc}")
                try:
                    self.driver.execute_script("arguments[0].click();", trigger)
                except Exception:
                    pass

        if opened:
            dropdown = next(
                (
                    el for el in status_item.find_elements(
                        By.XPATH,
                        ".//div[contains(@class,'el-select-dropdown') and not(contains(@style,'display: none'))]"
                    )
                    if self.is_visible(el)
                ),
                None
            )

            option = None
            if dropdown is not None:
                option = next(
                    (
                        el for el in dropdown.find_elements(
                            By.XPATH,
                            f".//li[contains(@class,'el-select-dropdown__item')]/span[normalize-space()='{status_text}']"
                        )
                        if self.is_visible(el)
                    ),
                    None
                )

            if option is not None:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", option)
                try:
                    actions.move_to_element(option).pause(0.2).click(option).perform()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", option)

                self.wait_for_loading_overlay()

                self.wait.until(lambda d: value_input.get_attribute("value") == status_text)
                actual_value = value_input.get_attribute("value")
                if actual_value != status_text:
                    raise TimeoutException(
                        f"Status was not applied correctly. Expected '{status_text}', got '{actual_value}'"
                    )

                print(f"[STATUS] Selected successfully via dropdown: {actual_value}")
                return

        print("[STATUS] Dropdown click path failed, using keyboard fallback")

        try:
            actions.move_to_element(trigger).pause(0.2).click(trigger).perform()
        except Exception:
            self.driver.execute_script("arguments[0].click();", trigger)

        self.wait_for_loading_overlay()

        try:
            value_input.send_keys(Keys.ENTER)
        except Exception:
            self.driver.execute_script("arguments[0].focus();", value_input)

        self.wait_for_loading_overlay()

        for _ in range(6):
            value_input.send_keys(Keys.ARROW_UP)

        for _ in range(target_index):
            value_input.send_keys(Keys.ARROW_DOWN)

        value_input.send_keys(Keys.ENTER)
        self.wait_for_loading_overlay()

        self.wait.until(lambda d: value_input.get_attribute("value") == status_text)

        actual_value = value_input.get_attribute("value")
        if actual_value != status_text:
            raise TimeoutException(
                f"Keyboard fallback failed. Expected '{status_text}', got '{actual_value}'"
            )

        print(f"[STATUS] Selected successfully via keyboard fallback: {actual_value}")

    def save_task(self) -> None:
        save_xpath = (
            "//div[contains(@class,'bottom_page')]//button[.//span[normalize-space()='Save']]"
            " | "
            "//button[.//span[normalize-space()='Save']]"
        )
        self.safe_click(By.XPATH, save_xpath, "Save button")

        def save_completed(driver):
            success_message = driver.find_elements(
                By.XPATH,
                "//div[contains(@class,'el-message') and contains(.,'Success')]"
            )
            visible_success = any(self.is_visible(el) for el in success_message)

            form_visible = any(
                self.is_visible(el)
                for el in driver.find_elements(
                    By.XPATH,
                    "//div[contains(@class,'leads_right') and .//h1[normalize-space()='Create New Task']]"
                )
            )

            return visible_success or not form_visible

        self.wait.until(save_completed)
        self.wait_for_loading_overlay()
        print("Save completed successfully")

    def convert_lead_to_opportunity(self) -> None:
        print("Checking whether lead conversion is needed...")

        try:
            lead_link_xpath = "//span[contains(@class,'p2')]//a[contains(.,'Leads')]"
            lead_link = self.wait_for_visible(By.XPATH, lead_link_xpath, timeout=SHORT_TIMEOUT)
            try:
                lead_link.click()
            except Exception:
                self.driver.execute_script("arguments[0].click();", lead_link)

            self.wait_for_loading_overlay()
        except TimeoutException:
            print("Lead link not found, skipping conversion")
            return

        convert_button_xpath = "//button[.//span[normalize-space()='Convert']]"

        try:
            convert_button = self.wait_for_visible(By.XPATH, convert_button_xpath, timeout=SHORT_TIMEOUT)
        except TimeoutException:
            print("No Convert button found, probably already converted or conversion not available. Skipping.")
            return

        try:
            convert_button.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", convert_button)

        self.wait_for_loading_overlay()

        confirm_convert_xpath = "(//button[.//span[normalize-space()='Convert']])[last()]"

        try:
            confirm_button = self.wait_for_visible(By.XPATH, confirm_convert_xpath, timeout=SHORT_TIMEOUT)
        except TimeoutException:
            print("Convert confirmation did not appear. Skipping conversion.")
            return

        try:
            confirm_button.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", confirm_button)

        self.wait_for_loading_overlay()

        try:
            self.wait.until(
                lambda d: any(
                    self.is_visible(el)
                    for el in d.find_elements(
                        By.XPATH,
                        "//div[contains(@class,'el-message') and (contains(.,'Success') or contains(.,'success'))]"
                    )
                )
                or not any(
                    self.is_visible(el)
                    for el in d.find_elements(By.XPATH, confirm_convert_xpath)
                )
            )
            print("Lead converted successfully")
        except TimeoutException:
            print("Conversion result not clearly shown, continuing anyway")

    def create_task(self, remark: str, status: str) -> None:
        form = self.get_form_container()

        self.type_in_field(form, "Comments", remark)
        self.select_dropdown_value(form, "Type", "Contact")
        self.select_dropdown_value(form, "Sub Type", "Outbound Call")
        self.type_in_field(form, "Subject", "Outbound Call")
        self.set_due_date_today(form)

        print(f"Trying to set Additional Information Status -> {status}")
        self.set_additional_information_status(form, status)

        print("Status confirmed, now saving...")
        self.save_task()

    def process_phone(self, phone: str, remark: str) -> ProcessResult:
        status = get_status(remark)
        last_error = ""

        for attempt in range(1, 3):
            try:
                self.reset_to_home()
                print(f"Processing {phone} | Attempt {attempt} | Evaluated Status -> {status}")

                record_type = self.search_student(phone)
                if record_type is None:
                    return ProcessResult(phone, False, "No student found in Lead or Opportunity after checking both tabs")

                self.open_first_student_result(phone)
                self.click_related_tab()
                self.click_new_task()
                self.create_task(remark, status)

                if record_type == "lead" and status in CONVERT_STATUSES:
                    self.convert_lead_to_opportunity()

                return ProcessResult(
                    phone,
                    True,
                    f"Task created with status: {status} | Source: {record_type}"
                )

            except TimeoutException as exc:
                last_error = f"Timeout on attempt {attempt}: {exc}"
                print(last_error)
            except Exception as exc:
                last_error = f"Unexpected error on attempt {attempt}: {exc}"
                print(last_error)

        return ProcessResult(phone, False, last_error)


def main() -> int:
    logger = CRMTaskLogger()
    results = []

    try:
        logger.open()

        for phone, remark in load_rows(EXCEL_FILE):
            result = logger.process_phone(phone, remark)
            results.append(result)

            if result.success:
                print(f"[SUCCESS] {result.phone}: {result.message}")
            else:
                print(f"[FAILED] {result.phone}: {result.message}")

    finally:
        input("Processing finished. Press Enter to close browser... ")
        logger.close()

    success_count = sum(1 for x in results if x.success)
    failed_count = len(results) - success_count
    print(f"Completed. Success: {success_count} | Failed: {failed_count}")

    return 0 if failed_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
