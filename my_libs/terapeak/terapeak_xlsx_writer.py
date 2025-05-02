import my_libs.utils as Utils
from my_libs.dependencies import *


class TerapeakData(Enum):
    """
    Enum class representing the various keys used for extracting data from research table rows.

    Members:
        TITLE
        TITLE_HREF
        KEYWORD
        AVG_SOLD_PRICE
        AVG_SHIPPING_COST
        TOTAL_SOLD
        ITEM_SALES
        DATE_LAST_SOLD
        IMAGE_URL
        IMAGE_PATH

    Usage:
        The `TerapeakData` enum provides standardized keys for accessing specific data extracted from research table rows. Each key corresponds to a particular type of information related to products being analyzed.

    Notes:
        - Enum members are used as keys in dictionaries where data is organized and retrieved.
        - `IMAGE_PATH` represents the local path of the saved image file, in addition to `IMAGE_URL` for the image's URL.
    """

    TITLE = DataAttr(header="Title", column=2)
    TITLE_HREF = DataAttr()
    KEYWORD = DataAttr(header="Keyword", column=1)
    AVG_SOLD_PRICE = DataAttr(header="Avg Sold Price", column=3)
    AVG_SHIPPING_COST = DataAttr(header="Avg Shipping Cost", column=4)
    TOTAL_SOLD = DataAttr(header="Total Sold", column=5)
    ITEM_SALES = DataAttr(header="Total Sale", column=6)
    DATE_LAST_SOLD = DataAttr(header="Last Sold Date", column=7)
    IMAGE_URL = DataAttr()
    IMAGE_PATH = DataAttr(header="Image", column=0)


class DaysRange(Enum):
    """
    Enum class representing the range of days for data analysis.

    Members:
        THIRTY: Represents a 30-day range for data analysis.
        NINETY: Represents a 90-day range for data analysis.

    Usage:
        The `DaysRange` enum provides standardized day ranges for filtering or analyzing data over specific periods. Each member represents a distinct time range used in data queries or reports.

    Notes:
        - Enum members are used to specify the time frame for which data should be analyzed or reported.
    """

    THIRTY = 30
    NINETY = 90


class MyTerapeakExcel:
    """
    A class for managing an Excel workbook to store Terapeak data, including data for different date ranges.

    This class handles the creation of an Excel workbook, adding and formatting worksheets, and writing data to the sheets.

    Attributes:
        workbook (xlsxwriter.Workbook): The Excel workbook instance.
        last_30_days_sheet (xlsxwriter.worksheet.Worksheet): Worksheet for storing data from the last 30 days.
        last_90_days_sheet (xlsxwriter.worksheet.Worksheet): Worksheet for storing data from the last 90 days.
        formats (dict[FormatType, xlsxwriter.format.Format]): Dictionary of formats used in the workbook.
        row_counts (dict[DaysRange, int]): Dictionary tracking row counts for different date ranges.

    Methods:
        create_workbook(keyword: str, output_directory: str) -> xlsxwriter.Workbook:
            Creates a new Excel workbook with a filename based on the provided keyword.
        save_workbook() -> None:
            Adjusts column widths for all sheets then saves and closes the Excel workbook, handling potential errors.
        add_headers() -> None:
            Adds headers to each sheet in the workbook.
        write_data_row(days_range: DaysRange, data: dict[TerapeakDataKey, Any]) -> None:
            Writes a row of data to the appropriate worksheet based on the date range.
        write_total_sold(days_range: DaysRange, total_sold: int) -> None:
            Writes the total number of items sold to the appropriate worksheet.
    """

    def __init__(self, keyword: str, output_dir: str) -> None:
        """
        Initialize an Excel workbook for storing Terapeak data.

        Args:
            keyword (str): The keyword used to name the Excel file.
            output_dir (str): The directory where the workbook will be saved.

        Attributes:
            workbook (xlsxwriter.Workbook): The Excel workbook instance.
            last_30_days_sheet (xlsxwriter.worksheet.Worksheet): Worksheet for the last 30 days data.
            last_90_days_sheet (xlsxwriter.worksheet.Worksheet): Worksheet for the last 90 days data.
            formats (dict[FormatType, xlsxwriter.format.Format]): Dictionary of formats used in the workbook.
            row_counts (dict[DaysRange, int]): Dictionary tracking row counts for different day ranges.
        """
        self.row_counts: dict[DaysRange, int] = {
            DaysRange.THIRTY: 0,
            DaysRange.NINETY: 0,
        }
        self.workbook: xlsxwriter.Workbook = self.create_workbook(keyword, output_dir)
        self.formats: dict[FormatType, xlsxwriter.format.Format] = initialize_formats(
            self.workbook
        )
        self.last_30_days_sheet = self.workbook.add_worksheet("Last 30 days")
        self.last_90_days_sheet = self.workbook.add_worksheet("Last 90 days")
        self.add_headers()

    def create_workbook(
        self, keyword: str, output_directory: str
    ) -> xlsxwriter.Workbook:
        """
        Create a new Excel workbook with a filename based on the provided keyword.

        Args:
            keyword (str): The keyword used to name the Excel file.
            output_directory (str): The directory where the workbook will be saved.

        Returns:
            xlsxwriter.Workbook: The created Workbook instance.
        """
        output_file = os.path.join(output_directory, f"{keyword}.xlsx")
        workbook = xlsxwriter.Workbook(output_file)
        logging.info(f"Workbook created successfully at {output_file}")
        return workbook

    def save_workbook(self) -> None:
        """
        Adjust column widths for all sheets and save the workbook.

        This method will autofit the columns and set a specific width for the first column of each sheet,
        then save and close the workbook. This method handles saving the workbook and manages potential
        errors such as permission issues.

        Notes:
            - Autofits columns in all worksheets.
            - Sets the width for the first column to one-sixth of 100.
            - Saves and closes the workbook, handling any errors related to file permissions.
            - Logs a message indicating success or an error if the workbook cannot be saved.
            - Retries if the file is open elsewhere, prompting the user to close it and retry.
        """
        for sheet in self.workbook.worksheets():
            # Autofit columns
            sheet.autofit()
            # Set column width for the first column
            sheet.set_column(0, Utils.get_enum_col(TerapeakData.IMAGE_PATH), 100 / 6)
        while True:
            try:
                self.workbook.close()
                logging.info("Workbook successfully saved.")
                break  # Exit the loop if the workbook is saved successfully
            except OSError as e:
                if e.errno == errno.EACCES:  # Permission denied error
                    logging.error(
                        "PermissionError: Please close the Excel file if it is open and press Enter to retry."
                    )
                    input("Please close the Excel file and press Enter to retry...")
                else:
                    logging.error(
                        "An OSError occurred while saving the workbook: %s", e
                    )
                    input("An unexpected error occurred. Press Enter to retry...")
            except Exception as e:
                logging.error(
                    "An unexpected error occurred while saving the workbook: %s", e
                )
                input("An unexpected error occurred. Press Enter to retry...")

    def add_headers(self) -> None:
        """
        Add header rows to each worksheet in the workbook.

        This method writes a standard set of headers to all sheets in the workbook,
        formatted according to the predefined header format.
        """
        headers = Utils.get_enum_headers_row(TerapeakData)
        # Add headers to each sheet
        for sheet in self.workbook.worksheets():
            sheet.write_row(0, 0, headers, self.formats[FormatType.HEADER])
        for key in self.row_counts:
            self.row_counts[key] += 1

    def write_data_row(
        self,
        days_range: DaysRange,
        data: dict[TerapeakData, Any],
        lock: Optional[Lock] = None,
    ) -> None:
        """
        Add data, images, and hyperlinks to a specified row in the Excel sheet for the given date range.

        Args:
            days_range (DaysRange): The range of days (e.g., last 30 days or last 90 days) to determine the sheet.
            data (dict[TerapeakDataKey, Any]): A dictionary containing the extracted product data, with keys defined in `TerapeakDataKey`.

        Notes:
            - The row height is set to 100.
            - The method writes product data into specific columns, formats cells, and embeds an image if a path is provided.
            - Adds a hyperlink if a valid link is available; otherwise, writes the title as plain text.
            - The following columns are populated:
                - Column 0: Image (embedded if `IMAGE_PATH` is provided)
                - Column 1: Keyword
                - Column 2: Title (with hyperlink if available)
                - Column 3: Average Sold Price (formatted as currency)
                - Column 4: Average Shipping Cost (formatted as currency)
                - Column 5: Total Sold (formatted as number)
                - Column 6: Item Sales (formatted as currency)
                - Column 7: Date Last Sold (formatted as date)
        """
        sheet = (
            self.last_30_days_sheet
            if days_range == DaysRange.THIRTY
            else self.last_90_days_sheet
        )
        row_index = self.row_counts[days_range]
        sheet.set_row(row_index, 100)

        # Insert the image if an image path is provided
        image_path = data.get(TerapeakData.IMAGE_PATH)
        if image_path:
            sheet.embed_image(row_index, 0, image_path)

        fields_to_write = [
            (TerapeakData.TITLE, TerapeakData.TITLE_HREF),
            (TerapeakData.KEYWORD, None),
            (TerapeakData.AVG_SOLD_PRICE, None),
            (TerapeakData.AVG_SHIPPING_COST, None),
            (TerapeakData.TOTAL_SOLD, None),
            (TerapeakData.ITEM_SALES, None),
            (TerapeakData.DATE_LAST_SOLD, None),
        ]

        for data_key, url_key in fields_to_write:
            check_genuine = data_key == TerapeakData.TITLE
            is_date = data_key == TerapeakData.DATE_LAST_SOLD
            is_currency = data_key in (
                TerapeakData.AVG_SHIPPING_COST,
                TerapeakData.AVG_SOLD_PRICE,
                TerapeakData.ITEM_SALES,
            )

            Utils.write_data(
                sheet,
                self.formats,
                row_index,
                Utils.get_enum_col(data_key),
                data,
                data_key,
                url_key=url_key,
                check_genuine=check_genuine,
                is_date=is_date,
                is_currency=is_currency,
                lock=lock,
            )

        self.row_counts[days_range] += 1

    def write_total_sold(self, days_range: DaysRange, total_sold: int) -> None:
        sheet = (
            self.last_30_days_sheet
            if days_range == DaysRange.THIRTY
            else self.last_90_days_sheet
        )
        sheet.write_string(
            0,
            8,
            f"Total Sold = {total_sold}",
            cell_format=self.formats[FormatType.HEADER],
        )

    # def add_screenshot(self, days_range: DaysRange, file_path: str) -> None:
    #     """
    #     Add a screenshot image to the specified worksheet at the current row index.

    #     Args:
    #         days_range (DaysRange) The DaysRange to identify which worksheet to insert the image.
    #         file_path (str): The file path of the screenshot to insert.
    #     """
    #     sheet = (
    #         self.last_30_days_sheet
    #         if days_range == DaysRange.THIRTY
    #         else self.last_90_days_sheet
    #     )

    #     row_idx = self.row_counts[days_range]

    #     Utils.add_screenshot_to_sheet(sheet, row_idx, file_path)

    #     # Update the row count for the sheet
    #     self.row_counts[days_range] += 1
