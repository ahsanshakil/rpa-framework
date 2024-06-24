from robocorp.tasks import task
from robocorp import browser

from RPA.PDF import PDF
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Archive import Archive
import time

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=600,
    )

    open_robot_order_website()
    bypass_modal()
    download_csv_file()
    fill_form_with_orders_data()
    zip_archive()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order/")

def bypass_modal():
    """Bypasses the modal that opens when hitting the website"""
    page = browser.page()
    page.click("button:text('OK')")

def fill_form_with_orders_data() -> None:
    """Gets Data from orders.csv and fills the form on each request"""
    orders = get_orders()
    page = browser.page()

    for order in orders:
        print(f"Processing Order: {order}")

        page.select_option("#head", str(order['Head']))
        page.click(f"#id-body-{order['Body']}")
        page.get_by_placeholder("Enter the part number for the legs").fill(order["Legs"])
        page.fill("#address", order["Address"])
        page.click("#preview")

        while not page.locator("#receipt").is_visible():
            page.click("#order")
            time.sleep(1)

        pdf_path = store_receipt_as_pdf(order["Order number"])
        screenshot_path = take_robot_screenshot(order["Order number"])
        embed_screenshot_to_receipt(screenshot_path, pdf_path)
        page.click("#order-another")
        bypass_modal()


def get_orders():
    """Reads orders from CSV file"""
    tables = Tables()
    return tables.read_table_from_csv("orders.csv")

def download_csv_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def store_receipt_as_pdf(order_number):
    """Stores the order receipt as a PDF file"""
    page = browser.page()
    receipt_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf_path = f"output/receipts/receipt_{order_number}.pdf"
    pdf.html_to_pdf(receipt_results_html, pdf_path)
    print(f"Receipt saved as: {pdf_path}")
    return pdf_path

def take_robot_screenshot(order_number):
    """Takes a screenshot of the robot"""
    page = browser.page()
    page.screenshot(path=f"output/screenshots/robot_{order_number}.png", full_page=False)
    print(f"Robot screenshot saved as: output/screenshots/robot_{order_number}.png")
    return f"output/screenshots/robot_{order_number}.png"

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the screenshot of the robot into the PDF receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=pdf_file,
    )
    print(f"Screenshot embedded into receipt: {pdf_file}")

def zip_archive():
    """Creates a zip archive of the receipts folder"""

    lib = Archive()

    lib.archive_folder_with_zip("./output/receipts", 'receipts.zip', recursive=True)
