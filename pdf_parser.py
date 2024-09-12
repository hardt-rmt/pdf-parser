import pymupdf
import pandas as pd


def split_pdf(input_pdf, output_directory):
    # Open the input PDF document
    print("Parsing the document...")
    pdf_document = pymupdf.open(input_pdf)
    purchase_order_details = get_page_purchase_order_details(pdf_document)
    purchase_orders = purchase_order_details[0]
    purchase_order_dates = purchase_order_details[1]
    vendor_numbers = purchase_order_details[2]
    count = 0

    print("Splitting the pdf file into the ", output_directory, " folder...")
    # Loop through each page and save it as a separate PDF
    for i in range(1, len(purchase_orders)):
        current_pdf = purchase_orders[i]
        previous_pdf = purchase_orders[i-1]

        if current_pdf == previous_pdf:
            count += 1
            continue
        else:
            if count > 0:
                start_page = i - count - 1
                end_page = i - 1
                count = 0
                create_sing_page_pdf(pdf_document, start_page, end_page, output_directory, purchase_orders[start_page])
                continue
            else:
                start_page = i - 1
                end_page = i - 1
                count = 0
                create_sing_page_pdf(pdf_document, start_page, end_page, output_directory, purchase_orders[start_page])
                continue

    if count > 1:
        start_page = len(purchase_orders) - 1 - count
        end_page = len(purchase_orders) - 1
        create_sing_page_pdf(pdf_document, start_page, end_page, output_directory, purchase_orders[start_page])
    else:
        start_page = len(purchase_orders) - 1
        end_page = len(purchase_orders) - 1
        create_sing_page_pdf(pdf_document, start_page, end_page, output_directory, purchase_orders[len(purchase_orders)-1])

    pdf_document.close()
    print("The document has been successfully split into individual files based on their purchase order numbers.")
    export_to_excel(purchase_order_dates, purchase_orders, vendor_numbers)


def create_sing_page_pdf(pdf_document, start_page, end_page, output_directory, purchase_order):
    # Create a new PDF with the single page
    single_page_pdf = pymupdf.open()
    single_page_pdf.insert_pdf(pdf_document, from_page=start_page, to_page=end_page)

    # Save the single page PDF to the specified directory
    output_path = f"{output_directory}/{purchase_order}.pdf"
    single_page_pdf.save(output_path)
    single_page_pdf.close()


def get_purchase_details(pdf_document, page_num, identifier, delimiter):
    page = pdf_document.load_page(page_num)
    potential_indices = []
    text = page.get_text("text")
    split_text = text.split("\n")

    for index, item in enumerate(split_text):
        if item.startswith(identifier):
            potential_indices.append(index)

    purchase_order_index = potential_indices[len(potential_indices) - 1]
    purchase_order = split_text[purchase_order_index]
    purchase_order_value = purchase_order[delimiter:]
    return purchase_order_index, purchase_order_value


def get_purchase_order_date(pdf_document, page_num, date_index):
    page = pdf_document.load_page(page_num)
    text = page.get_text("text")
    split_text = text.split("\n")
    parsed_date = split_text[date_index]
    return parsed_date[:10]


def get_page_purchase_order_details(pdf_document):
    purchase_orders = []
    purchase_order_dates = []
    vendor_numbers = []
    for page_num in range(len(pdf_document)):
        purchase_order_index, purchase_order = get_purchase_details(pdf_document, page_num, "BYU-", 4)
        purchase_orders.append(purchase_order)
        purchase_order_date = get_purchase_order_date(pdf_document, page_num, purchase_order_index + 1)
        purchase_order_dates.append(purchase_order_date)
        vendor_number = get_purchase_details(pdf_document, page_num, "Supplier:", 10)
        vendor_numbers.append(vendor_number[1])
    return [purchase_orders, purchase_order_dates, vendor_numbers]


def export_to_excel(purchase_order_dates, purchase_orders, vendor_numbers):
    data = {
        'Date': purchase_order_dates,
        'Purchase Order': purchase_orders,
        'Supplier Number': vendor_numbers
    }

    # Create DataFrame
    df = pd.DataFrame(data)

    # Export DataFrame to Excel file
    df.to_excel('purchase-order-details.xlsx', index=False)

    print("Data has been exported to the purchase-order-details.xlsx file")


pdf_path = input("Enter the file path: ")
output_dir = input("Enter the output directory: ")
split_pdf(pdf_path, output_dir)
