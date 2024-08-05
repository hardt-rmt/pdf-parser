import pymupdf
import pandas as pd


def split_pdf(input_pdf, output_dir):
    # Open the input PDF
    pdf_document = pymupdf.open(input_pdf)
    purchase_order_details = get_page_purchase_order_details(pdf_document)
    purchase_orders = purchase_order_details[0]
    purchase_order_dates = purchase_order_details[1]
    vendor_numbers = purchase_order_details[2]
    start_page = None

    # Loop through each page and save it as a separate PDF
    for i in range(1, len(purchase_orders)):
        current_pdf = purchase_orders[i]
        previous_pdf = purchase_orders[i-1]

        if current_pdf is None:
            if previous_pdf is not None:
                start_page = i - 1
                continue
            else:
                if i == len(purchase_orders) - 1:
                    end_page = i
                    create_sing_page_pdf(pdf_document, start_page, end_page, output_dir, purchase_orders[start_page])
                    break
                else:
                    continue
        else:
            if previous_pdf is not None:
                start_page = i - 1
                end_page = i - 1
                create_sing_page_pdf(pdf_document, start_page, end_page, output_dir, purchase_orders[start_page])
                continue
            else:
                end_page = i - 1
                create_sing_page_pdf(pdf_document, start_page, end_page, output_dir, purchase_orders[start_page])

    if purchase_orders[len(purchase_orders) - 1] is not None:
        start_page = len(purchase_orders) - 1
        end_page = len(purchase_orders) - 1
        create_sing_page_pdf(pdf_document, start_page, end_page, output_dir, purchase_orders[len(purchase_orders)-1])

    pdf_document.close()
    export_to_excel(purchase_order_dates, purchase_orders, vendor_numbers)


def create_sing_page_pdf(pdf_document, start_page, end_page, output_dir, purchase_order):
    # Create a new PDF with the single page
    single_page_pdf = pymupdf.open()
    single_page_pdf.insert_pdf(pdf_document, from_page=start_page, to_page=end_page)

    # Save the single page PDF to the specified directory
    output_path = f"{output_dir}/{purchase_order}.pdf"
    single_page_pdf.save(output_path)
    single_page_pdf.close()


def get_purchase_details(pdf_document, page_num, identifier, delimiter):
    page = pdf_document.load_page(page_num)
    purchase_order_index = None
    text = page.get_text("text")
    split_text = text.split("\n")
    for index, item in enumerate(split_text):
        if item.startswith(identifier):
            purchase_order_index = index
            break
        else:
            continue
    if purchase_order_index is not None:
        purchase_order = split_text[purchase_order_index]
        purchase_order_value = purchase_order[delimiter:]
        return purchase_order_value
    else:
        return None


def get_page_purchase_order_details(pdf_document):
    purchase_orders = []
    purchase_order_dates = []
    vendor_numbers = []
    for page_num in range(len(pdf_document)):
        purchase_order = get_purchase_details(pdf_document, page_num, "P.O. #", 7)
        purchase_orders.append(purchase_order)
        purchase_order_date = get_purchase_details(pdf_document, page_num, "P.O. Date", 10)
        purchase_order_dates.append(purchase_order_date)
        vendor_number = get_purchase_details(pdf_document, page_num, "Vendor #:", 10)
        vendor_numbers.append(vendor_number)
    return [purchase_orders, purchase_order_dates, vendor_numbers]


def export_to_excel(purchase_order_dates, purchase_orders, vendor_numbers):
    data = {
        'Date': purchase_order_dates,
        'Purchase Order': purchase_orders,
        'Vendor Number': vendor_numbers
    }

    # Create DataFrame
    df = pd.DataFrame(data)

    # Export DataFrame to Excel file
    df.to_excel('purchase-order-details.xlsx', index=False)

    print("Data has been exported to purchase-order-details.xlsx")


# Example usage
pdf = input("Enter PDF file path: ")
output_directory = input("Enter output directory: ")
split_pdf(pdf, output_directory)
