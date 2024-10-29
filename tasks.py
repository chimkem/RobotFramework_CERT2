from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    #browser.configure(slowmo=1000)
    page = browser.page()

    open_robot_order_website()

    """Loop the orders"""
    orders = get_orders()
    for order in orders:
        fill_the_form(order)
        
        screenshot = screenshot_robot(order["Order number"])
        receipt = store_receipt_as_pdf(order["Order number"])

        embed_screenshot_to_receipt(screenshot, receipt)
        page.click("button:text('Order another robot')")
        close_annoying_modal()

    archive_receipts()

# - - - - - - - - - - - - - - - - - - - - -

def open_robot_order_website():
    """Navigate to the URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()

def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    excel_results = Tables()
    return excel_results.read_table_from_csv("orders.csv", header=True)

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('Yep')")

# - - - FILL THE FORM - - -

def fill_the_form(order):
    page = browser.page()

    """Choose the parts"""
    page.select_option("#head", str(order["Head"]))
    page.click("#id-body-" + str(order["Body"]))
    page.fill("//label[contains(text(),'3.')]/following-sibling::input",  str(order["Legs"]))
    page.fill("#address", str(order["Address"]))

    """preview to get the image of the robot and finish the order"""
    page.click("button:text('Preview')")
    page.click("button:text('Order')")

    """In case of error"""
    max_attempts = 5
    for attempt in range(max_attempts):
        if page.locator(".alert.alert-danger").count() > 0: 
            page.click("button:text('Order')")
        else:
            break

# - - - - - - - - - - - - - - - - - - 

def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    pdf = PDF()
    
    pdf_result = ("output/receipt_from_order" + order_number + ".pdf")
    pdf.html_to_pdf(page.locator("#order-completion").inner_html(), pdf_result)
    return pdf_result

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    screenshot_result = ("output/screenshot_from_order" + order_number + ".png")
    page.locator("#robot-preview-image").screenshot(path=screenshot_result)
    return screenshot_result

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Append the Robot preview screenshot to the reciept PDF"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot, source_path= pdf_file, output_path=pdf_file)

def archive_receipts():
    """Make a ZIP for all of the receipts"""
    archive = Archive()
    archive.archive_folder_with_zip("output/receipts", "output/zip_receipts.zip")