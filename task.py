# +
"""Template robot with Python."""

import os

from Browser.utils.data_types import SelectAttribute
from RPA.Archive import Archive
from RPA.Browser.Playwright import Playwright
from RPA.Dialogs import Dialogs
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Robocorp.Vault import Vault
from RPA.Tables import Tables

# +
browser = Playwright()

output_dir = os.path.join(os.getcwd(), "output")
if not os.path.exists:
    os.mkdir(output_dir)


# -

def get_orders_file_from_user():
    dialog = Dialogs()
    dialog.add_text_input(
        name="orders_file",
        label="Orders file path"
    )

    input = dialog.run_dialog()
    return input.orders_file


def get_build_a_robot_url_from_vault():
    build_a_robot_url = Vault().get_secret("website_credentials")
    return build_a_robot_url["build_a_robot_url"]


def download_orders_file(file):
    http = HTTP()
    http.download(
        url=file,
        target_file=os.path.join(output_dir, "orders.csv"),
        overwrite=True
    )


def open_intranet_website(url):
    browser.open_browser(url)


def close_annoying_modal():
    browser.click(selector="css=button.btn.btn-dark")


def save_receipt_to_pdf(order_id):
    receipt_pdf_filename = order_id + "_order_receipt.pdf"

    browser.wait_for_elements_state(selector="id=receipt")
    receipt_html_file = browser.get_property(
        selector="id=receipt",
        property="outerHTML"
    )

    pdf = PDF()
    pdf.html_to_pdf(
        receipt_html_file,
        os.path.join(output_dir, receipt_pdf_filename)
    )


def save_robot_preview_to_file(order_id):
    robot_image_file_prefix = order_id + "_robot_preview"

    browser.wait_for_elements_state(
        selector="css=#robot-preview-image>img[alt='Head']"
    )
    browser.wait_for_elements_state(
        selector="css=#robot-preview-image>img[alt='Body']"
    )
    browser.wait_for_elements_state(
        selector="css=#robot-preview-image>img[alt='Legs']"
    )

    browser.take_screenshot(
        filename=os.path.join(output_dir, robot_image_file_prefix),
        selector="id=robot-preview-image"
    )


def generate_detailed_receipt_pdf(order_id):
    save_receipt_to_pdf(order_id)
    save_robot_preview_to_file(order_id)

    receipt_file = order_id + "_order_receipt.pdf"
    robot_preview_file = order_id + "_robot_preview.png"
    full_receipt_file = order_id + "_full_order_receipt.pdf"

    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[
            os.path.join(output_dir, receipt_file),
            os.path.join(output_dir, robot_preview_file),
        ],
        target_document=os.path.join(output_dir, full_receipt_file),
    )


def submit_form_for_one_order(order):
    browser.select_options_by(
        "css=select#head.custom-select",
        SelectAttribute["value"],
        order["Head"]
    )

    order_body_selector = 'id=id-body-' + order["Body"]
    browser.check_checkbox(order_body_selector)

    browser.type_text(
        "css=input[placeholder='Enter the part number for the legs']",
        order["Legs"]
    )

    browser.type_text(
        "css=input#address.form-control",
        order["Address"]
    )

    browser.click(selector="id=preview")
    click_submit_button_resiliently()


def click_submit_button_resiliently():
    browser.click(selector="id=order")

    while browser.get_element_state("css=.alert-danger"):
        browser.click(selector="id=order")


def process_one_order(order):
    submit_form_for_one_order(order)
    generate_detailed_receipt_pdf(order["Order number"])


def process_orders_using_data_from_orders_file():
    csv = Tables()
    orders = csv.read_table_from_csv(
        path=os.path.join(output_dir, "orders.csv"),
        header=True
    )

    for order in orders:
        process_one_order(order)
        browser.click(
            selector="id=order-another"
        )
        close_annoying_modal()


def display_path_of_receipts_zip_file(zip_file):
    dialog = Dialogs()
    dialog.add_text(f"Receipts zip file: {zip_file}")
    dialog.show_dialog()


def archive_all_receipts_in_zip_file():
    zip_file_path = os.path.join(output_dir, "order_receipts.zip")

    archiver = Archive()
    archiver.archive_folder_with_zip(
        output_dir,
        zip_file_path,
        include="*full*.pdf",
    )

    display_path_of_receipts_zip_file(zip_file_path)


# +
def main():
    try:
        orders_file_path = get_orders_file_from_user()
        build_a_robot_url = get_build_a_robot_url_from_vault()

        download_orders_file(orders_file_path)
        open_intranet_website(build_a_robot_url)
        close_annoying_modal()
        process_orders_using_data_from_orders_file()
        archive_all_receipts_in_zip_file()
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()

