import concurrent
import logging
import os
import time
from threading import Lock
from typing import Any, Callable, Optional

import dateutil.parser as dparser
import my_libs.utils as Utils
import my_libs.web_driver as Driver
from my_libs.terapeak.terapeak_xlsx_writer import (
    DaysRange,
    MyTerapeakExcel,
    TerapeakData,
)
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

MAX_SCRAPE_ROW = 500


def process_keywords(keywords: list[str], output_dir: str) -> None:
    """
    Scrape product data for a list of keywords and save the results to the specified directory.

    This function manages the scraping process for each keyword by creating and running multiple threads to fetch and process data
    over different time ranges (30 days and 90 days). It handles image downloading, data extraction, and workbook management.

    Args:
        keywords (list[str]): A list of search keywords for which to scrape product data. Each keyword is processed to collect data
                              and images associated with it.
        output_dir (str): The directory path where the results, including scraped data and images, will be saved.

    Returns:
        None: This function does not return a value. It saves data and images to the specified directory and manages the lifecycle
              of the scraping tasks, including cleanup of temporary resources.

    Notes:
        - If no keywords are provided, the function logs a warning and skips the data fetch process.
        - The function ensures that an image folder is created or verified before starting the scraping process.
        - It uses a `ThreadPoolExecutor` to parallelize the scraping tasks, with a maximum of two concurrent workers.
        - After all tasks are completed, it saves the results into Excel workbooks, and then deletes the temporary
          image folder.
    """
    if not keywords:
        logging.warning("No keywords provided. Skipping data fetch.")
        return

    logging.info("Fetching and saving data for keywords: %s", ", ".join(keywords))

    product_images_folder_path = Utils.create_subfolder(
        output_dir, "Terapeak Product Images"
    )
    screenshots_folder_path = Utils.create_subfolder(output_dir, "Terapeak Screenshots")
    logging.info("Image folder created or ensured at: %s", product_images_folder_path)

    try:
        total_workbook = MyTerapeakExcel("All Terapeak Data", output_dir)
        total_workbook_lock = Lock()
        max_workers = 2

        # Create a queue to hold the WebDriver instances
        driver_pool = Driver.DriverPool(max_workers)

        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            keyword_tasks = {}

            for keyword in keywords:
                keyword = keyword.strip()
                if keyword:
                    # Create an instance of KeywordScraper for each keyword and day range
                    keyword_workbook = MyTerapeakExcel(keyword, output_dir)
                    scraper_30 = KeywordScraper(
                        keyword,
                        DaysRange.THIRTY,
                        keyword_workbook,
                        total_workbook,
                        output_dir,
                        product_images_folder_path,
                        screenshots_folder_path,
                        total_workbook_lock,
                    )
                    scraper_90 = KeywordScraper(
                        keyword,
                        DaysRange.NINETY,
                        keyword_workbook,
                        total_workbook,
                        output_dir,
                        product_images_folder_path,
                        screenshots_folder_path,
                        total_workbook_lock,
                    )

                    logging.info("Submitting tasks for search keyword: %s", keyword)

                    # Submit tasks and store futures
                    future_30 = executor.submit(
                        scraper_30.scrape_keyword_data, driver_pool
                    )
                    futures[future_30] = scraper_30

                    future_90 = executor.submit(
                        scraper_90.scrape_keyword_data, driver_pool
                    )
                    futures[future_90] = scraper_90

                    # Initialize tasks for keyword
                    keyword_tasks[keyword] = {"30": False, "90": False}
                else:
                    logging.warning("Empty search keyword encountered. Skipping.")

            # Collect results as futures complete
            for future in concurrent.futures.as_completed(futures):
                scraper = futures[future]
                keyword = scraper.keyword
                days_range = scraper.days_range.value

                try:
                    if days_range == 30:
                        keyword_tasks[keyword]["30"] = True
                    else:
                        keyword_tasks[keyword]["90"] = True
                except Exception as e:
                    logging.error(
                        f"Error processing future result for {keyword} ({days_range} days): {e}"
                    )
                    continue

                # Save workbook only after both tasks for a keyword are complete
                if keyword_tasks[keyword]["30"] and keyword_tasks[keyword]["90"]:
                    logging.info(
                        f"Both {keyword}'s days range task completed, saving workbook now..."
                    )
                    scraper.keyword_workbook.save_workbook()

            total_workbook.save_workbook()
        logging.info("All tasks completed successfully")
    except Exception as e:
        logging.error(f"Error while processing keywords: {e}")
    finally:
        driver_pool.cleanup()
        Utils.delete_folder(product_images_folder_path)


class KeywordScraper:
    def __init__(
        self,
        keyword: str,
        days_range: DaysRange,
        keyword_book: MyTerapeakExcel,
        total_book: MyTerapeakExcel,
        output_dir: str,
        product_images_folder_path: str,
        screenshots_folder_path: str,
        total_workbook_lock: Lock,
    ):
        """
        Initialize a KeywordScraper instance for scraping product data related to a specific keyword and day range.

        Args:
            keyword (str): The search keyword for which to scrape product data.
            days_range (DaysRange): The range of days to consider for scraping (e.g., 30 or 90 days).
            keyword_book (MyTerapeakExcel): An instance of MyTerapeakExcel for saving keyword-specific data.
            total_book (MyTerapeakExcel): An instance of MyTerapeakExcel for saving aggregated data across all keywords.
            output_dir (str): The directory path where data and images will be saved.

        Returns:
            None: Initializes the scraper instance without returning any value.
        """
        self.keyword: str = keyword
        self.days_range: DaysRange = days_range
        self.keyword_workbook: MyTerapeakExcel = keyword_book
        self.total_workbook: MyTerapeakExcel = total_book
        self.output_dir: str = output_dir
        self.driver: Optional[webdriver.Chrome] = None
        self.product_images_folder_path: str = product_images_folder_path
        self.screenshots_folder_path: str = screenshots_folder_path
        self.processed_rows: int = 0
        self.total_workbook_lock: Lock = total_workbook_lock
        self.logging_name: str = f"{self.keyword} Last {self.days_range.value} days"

    def scrape_keyword_data(self, driver_pool: Driver.DriverPool) -> None:
        """
        Perform the scraping process for the given keyword and day range.

        This method handles the complete scraping workflow including initializing the web driver, navigating to the URL,
        fetching and processing table rows, downloading images, and writing data to Excel workbooks. It also manages error handling
        and ensures proper closing of the web driver.

        Returns:
            None: This method does not return a value. It performs data scraping and processing, and manages the lifecycle of the
                  web driver.

        Notes:
            - Handles the cookie management and captcha login processes.
            - Collects data from each page and writes the processed data into workbooks.
            - Ensures the web driver is closed properly in the `finally` block.
        """
        try:
            self.driver = driver_pool.acquire()

            url = Utils.build_terapeak_url(self.keyword, self.days_range.value)
            logging.debug("Navigating to URL: %s", url)
            self.driver.get(url)

            all_extracted_data: list[dict[TerapeakData, Any]] = []

            page_num: int = 0

            while self.processed_rows < MAX_SCRAPE_ROW:
                try:
                    page_num += 1
                    Driver.check_ebay_captcha(self.driver, url)
                    time.sleep(2)
                    screenshot_path = os.path.join(
                        self.screenshots_folder_path,
                        f"{self.logging_name} Page {page_num}.png",
                    )
                    Utils.take_screenshot(screenshot_path, self.driver)
                    logging.info(
                        f"Fetching data for {self.logging_name} in page {page_num}"
                    )
                    rows = self.fetch_table_rows(self.driver, url)
                    if not rows:
                        logging.info(f"No more rows found for {self.logging_name}.")
                        break
                    try:
                        processed_data = self.process_rows_data(rows)
                        all_extracted_data.extend(processed_data)
                    except ValueError as ve:
                        logging.error(
                            f"Invalid data found while processing rows: {ve}. Stopping operation."
                        )
                        break

                    if self.next_page_available(rows):
                        url = Utils.build_terapeak_url(
                            self.keyword,
                            self.days_range.value,
                            offset=self.processed_rows,
                        )
                        self.driver.get(url)
                    else:
                        break

                except Exception as e:
                    logging.error(
                        f"Error while processing rows for {self.logging_name}: {e}"
                    )
                    break  # Exit loop on error

            sorted_data = sorted(
                all_extracted_data,
                key=lambda x: x.get(TerapeakData.AVG_SOLD_PRICE, 0),
                reverse=True,
            )
            self.write_sorted_data(sorted_data)

            total_sold_sum = self.calculate_total_sold(all_extracted_data)
            self.write_total_sold(total_sold_sum)
            logging.info(
                f"Wrote {self.processed_rows} rows of {self.logging_name} data."
            )

        except Exception as e:
            logging.error(
                "An error occurred during the scraping process for '%s': %s",
                self.logging_name,
                e,
            )

        finally:
            # Ensure the WebDriver is properly closed
            if self.driver:
                driver_pool.release(self.driver)

    def fetch_table_rows(self, driver: webdriver.Chrome, url: str) -> list[WebElement]:
        """
        Wait for and fetch the rows from the research table on a webpage.

        This method waits for the presence of the research table rows and handles scenarios where no sold results are found.

        Args:
            driver (webdriver.Chrome): The Selenium WebDriver instance used to interact with the webpage.
            url (str): The Terapeak search url in case the driver need to be refresh

        Returns:
            list[WebElement]: A list of WebElement objects representing the rows in the research table.

        Raises:
            NoSuchElementException: If the research table rows cannot be found.
            TimeoutException: If fetching the rows times out.

        Notes:
            - Uses a timeout for fetching rows.
            - Logs errors if the rows cannot be found.
        """
        attempts, max_retries = 0, 5

        while attempts < max_retries:
            try:
                # Wait for the presence of the research table rows
                WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, "tr.research-table-row")
                    )
                )

                current_url = driver.current_url
                if "tabName=ACTIVE" in current_url:
                    logging.info(
                        f"Detected ACTIVE tab instead of SOLD for {self.logging_name}. Skipping data extraction."
                    )
                    return []

                rows = driver.find_elements(By.CSS_SELECTOR, "tr.research-table-row")

                # If rows are found, return them early
                if rows:
                    logging.info(f"Fetched {len(rows)} rows from {self.logging_name}")
                    return rows

                # If no rows found, check for the error message
                try:
                    error_message = driver.find_element(
                        By.CSS_SELECTOR,
                        "div.research__generic-error .page-notice__title",
                    )
                    if "No sold results found" in error_message.text:
                        logging.info(
                            "No sold results found for keyword: %s", self.logging_name
                        )
                        return []

                except NoSuchElementException:
                    logging.error(
                        f"No rows and no error message found: {self.logging_name}"
                    )
                    raise Exception(
                        f"No results or error message found for {self.logging_name}"
                    )

            except (TimeoutException, NoSuchElementException) as e:
                logging.error(
                    f"Error fetching table rows for {self.logging_name}: {str(e)}",
                    exc_info=True,
                )
                attempts += 1
                if attempts <= max_retries:
                    logging.info(
                        f"Retrying to fetch table rows for {self.logging_name}... ({attempts} / {max_retries})"
                    )
                    try:
                        driver.get(url)
                        Driver.check_ebay_captcha(driver, url)
                    except Exception as e:
                        logging.error(
                            f"Error occurred when attempting to refresh driver: {e}"
                        )
            except Exception as e:
                logging.error(f"Unexpected error: {str(e)}")
                return []

        logging.error(
            "Failed to fetch table rows after %d attempts: %s",
            max_retries,
            self.logging_name,
        )
        return []

    def process_rows_data(
        self, rows: list[WebElement]
    ) -> list[dict[TerapeakData, Any]]:
        """
        Process table rows to extract data, download images, and return a list of dictionaries containing the extracted data.

        This method processes each row in the table to extract product data and download associated images. It limits the number of rows
        processed based on a predefined maximum limit.

        Args:
            rows (list[WebElement]): A list of WebElement objects representing the rows of the research table.

        Returns:
            (list[dict[TerapeakDataKey, Any]]): A list of dictionaries where each dictionary contains the data for a single row.

        Notes:
            - Handles image downloading and stores the path to each downloaded image.
            - Logs warnings for any issues encountered during data extraction.
        """

        extracted_data: list[dict[TerapeakData, Any]] = []
        logging.info(f"Processing {len(rows)} rows for {self.logging_name}")

        for index, row in enumerate(rows, start=self.processed_rows + 1):
            if self.processed_rows >= MAX_SCRAPE_ROW:
                logging.info(f"Reached maximum row limit of {MAX_SCRAPE_ROW}")
                break

            logging.debug(f"Processing row {index} for {self.logging_name}")
            data = self.parse_row_data(row)
            image_url = data.get(TerapeakData.IMAGE_URL)
            image_path = (
                Utils.download_image(
                    image_url,
                    self.product_images_folder_path,
                    f"Terapeak_{self.logging_name.replace(' ', '_')}_{index}",
                )
                if image_url
                else None
            )
            data[TerapeakData.IMAGE_PATH] = image_path
            self.processed_rows += 1
            extracted_data.append(data)
            logging.debug(f"Successfully processed row {index} for {self.logging_name}")

        logging.info(
            f"Finished processing {len(extracted_data)} rows for {self.logging_name}"
        )
        return extracted_data

    def parse_row_data(self, row: WebElement) -> dict[TerapeakData, Any]:
        """
        Extract and parse data from a row element in the research table for a given keyword.

        This method extracts various pieces of information from a row element, including title, average sold price, shipping cost,
        and other details. It handles cases where certain data may be missing or not present.

        Args:
            row (WebElement): The Selenium WebElement representing a row containing product data.

        Returns:
            (dict[TerapeakDataKey, Any]): A dictionary with `TerapeakDataKey` enum keys and their corresponding extracted values.

        Notes:
            - Uses helper functions to safely extract text and attributes from elements.
            - Handles cases where elements might be missing or extraction fails.
            - Logs debug information about the extracted data.
        """
        data: dict[TerapeakData, Any] = {}

        def safe_extract_text(
            selector, transform_func: Callable[[str], Any] = str
        ) -> Any:
            """
            Safely extract text from an element located by the CSS selector and transform it using a function.

            Args:
                selector: The CSS selector to locate the element.
                transform_func: Function to transform the extracted text.

            Returns:
                Any: Transformed text or None if extraction fails.
            """
            try:
                element = row.find_element(By.CSS_SELECTOR, selector)
                return transform_func(element.text)
            except (NoSuchElementException, ValueError):
                logging.debug("Failed to extract data for '%s'", selector)
                return None

        def safe_extract_attribute(selector, attribute_name) -> Any:
            """
            Safely extract an attribute value from an element located by the CSS selector.

            Args:
                selector: The CSS selector to locate the element.
                attribute_name: The attribute name to extract.

            Returns:
                Any: The attribute value or None if extraction fails.

            Raises:
                ValueError: If the row data doesn't match expected sold listing data format
            """
            try:
                element = row.find_element(By.CSS_SELECTOR, selector)
                return element.get_attribute(attribute_name)
            except NoSuchElementException:
                logging.debug(
                    "Failed to extract attribute '%s' for '%s'",
                    attribute_name,
                    selector,
                )
                return None

        data[TerapeakData.KEYWORD] = self.keyword
        title = safe_extract_text("div.research-table-row__product-info-name a span")
        if not title:
            title = safe_extract_text("div.research-table-row__product-info-name")
        else:
            url = safe_extract_attribute(
                "div.research-table-row__product-info-name a", "href"
            )
            data[TerapeakData.TITLE_HREF] = Utils.ebay_clean_product_url(url)
            data[TerapeakData.IMAGE_URL] = safe_extract_attribute(
                "div.__zoomable-thumbnail-inner img", "src"
            )
        data[TerapeakData.TITLE] = Utils.escape_quotes(title)
        data[TerapeakData.AVG_SOLD_PRICE] = safe_extract_text(
            "td.research-table-row__avgSoldPrice>div:first-child>div:first-child",
            lambda text: float(text.replace(",", "").replace("$", "")),
        )
        data[TerapeakData.AVG_SHIPPING_COST] = safe_extract_text(
            "td.research-table-row__avgShippingCost>div:first-child>div:first-child",
            lambda text: float(text.replace(",", "").replace("$", "")),
        )
        data[TerapeakData.TOTAL_SOLD] = safe_extract_text(
            "td.research-table-row__totalSoldCount>div:first-child>div:first-child",
            lambda text: int(text.replace(",", "")),
        )
        data[TerapeakData.ITEM_SALES] = safe_extract_text(
            "td.research-table-row__totalSalesValue>div:first-child>div:first-child",
            lambda text: float(text.replace(",", "").replace("$", "")),
        )
        date_last_sold = safe_extract_text(
            "td.research-table-row__dateLastSold>div:first-child>div:first-child",
            lambda text: dparser.parse(text, fuzzy=True),
        )
        data[TerapeakData.DATE_LAST_SOLD] = Utils.convert_to_excel_date(date_last_sold)

        # if data[DataKey.TITLE_HREF]:
        #     get_additional_product_data(data)

        logging.debug(f"Extracted data: {self.format_data_for_logging(data)}")

        if (
            data[TerapeakData.AVG_SOLD_PRICE] is None
            or data[TerapeakData.TOTAL_SOLD] is None
        ):
            raise ValueError(f"Invalid row data: Missing critical data.")

        return data

    def write_sorted_data(self, sorted_data: list[dict[TerapeakData, Any]]) -> None:
        """
        Write sorted data to the keyword-specific and total workbooks.

        This method writes the sorted product data to both the keyword-specific workbook and the total workbook for aggregation.

        Args:
            sorted_data (list[dict[TerapeakDataKey, Any]]): A list of dictionaries containing sorted product data.

        Notes:
            - The data is written to both the keyword-specific workbook and the total workbook.
        """
        for data in sorted_data:
            self.keyword_workbook.write_data_row(self.days_range, data)
            self.total_workbook.write_data_row(
                self.days_range, data, self.total_workbook_lock
            )

    def calculate_total_sold(self, datas: list[dict[TerapeakData, Any]]) -> int:
        """
        Calculate the total number of items sold from the given data.

        Args:
            datas (list[dict[TerapeakDataKey, Any]]): A list of dictionaries containing product data.

        Returns:
            int: The total number of items sold.

        Notes:
            - Sums up the `TOTAL_SOLD` values from each data dictionary.
        """
        return sum(data.get(TerapeakData.TOTAL_SOLD, 0) for data in datas)

    def write_total_sold(self, total_sold: int) -> None:
        """
        Write the total number of items sold to the keyword-specific workbook.

        Args:
            total_sold (int): The total number of items sold.

        Returns:
            None: This method does not return a value. It writes the total sold count to the workbook.

        Notes:
            - The total sold count is written to the keyword-specific workbook.
        """
        self.keyword_workbook.write_total_sold(self.days_range, total_sold)
        return

    def go_to_next_page(self, rows: list[WebElement]) -> bool:
        """
        Navigate to the next page if the "Next" button is enabled.

        Args:
            driver (webdriver.Chrome): The Selenium WebDriver instance.
            rows (list[WebElement]): List of row elements from the current page.

        Returns:
            bool: True if the navigation to the next page is successful, False otherwise.

        Notes:
            - Waits for the current page to become stale before returning.
            - Logs a message when no more pages are available.
        """
        try:
            if self.driver:
                next_button = self.driver.find_element(
                    By.CSS_SELECTOR, "button.pagination__next"
                )
                if next_button.is_enabled():
                    next_button.click()
                    WebDriverWait(self.driver, 60).until(EC.staleness_of(rows[0]))
                    return True
                else:
                    logging.info("No more pages to scrape.")
                    return False
            else:
                logging.error("Driver is not alive.")
                return False
        except Exception as e:
            logging.error(f"Error while navigating to the next page: {e}")
            return False

    def next_page_available(self, rows: list[WebElement]) -> bool:
        """
        Navigate to the next page if the "Next" button is enabled.

        Args:
            driver (webdriver.Chrome): The Selenium WebDriver instance.
            rows (list[WebElement]): List of row elements from the current page.

        Returns:
            bool: True if the navigation to the next page is successful, False otherwise.

        Notes:
            - Waits for the current page to become stale before returning.
            - Logs a message when no more pages are available.
        """
        try:
            if self.driver:
                next_button = self.driver.find_element(
                    By.CSS_SELECTOR, "button.pagination__next"
                )
                if next_button.is_enabled():
                    return True
                else:
                    logging.info("No more pages to scrape.")
                    return False
            else:
                logging.error("Driver is not alive.")
                return False
        except Exception as e:
            logging.error(f"Error while determining next page availability: {e}")
            return False

    def format_data_for_logging(self, data: dict[TerapeakData, Any]) -> str:
        """
        Format the extracted data into a readable string for logging purposes.

        Args:
            data (dict[TerapeakData, Any]): The data dictionary to format

        Returns:
            str: A formatted string representation of the data
        """
        return (
            f"\nData for '{self.keyword}' ({self.days_range.value} days):"
            f"\n  Title: {data.get(TerapeakData.TITLE, 'N/A')}"
            f"\n  URL: {data.get(TerapeakData.TITLE_HREF, 'N/A')}"
            f"\n  Avg Sold Price: ${data.get(TerapeakData.AVG_SOLD_PRICE, 0):.2f}"
            f"\n  Avg Shipping: ${data.get(TerapeakData.AVG_SHIPPING_COST, 0):.2f}"
            f"\n  Total Sold: {data.get(TerapeakData.TOTAL_SOLD, 0)}"
            f"\n  Total Sales: ${data.get(TerapeakData.ITEM_SALES, 0):.2f}"
            f"\n  Last Sold: {data.get(TerapeakData.DATE_LAST_SOLD, 'N/A')}"
        )
