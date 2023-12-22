from datetime import datetime
from jproperties import Properties
import pandas as pd
import pypdf
import sys
import os


application_path = ''
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

print('application_path is ' + application_path)
properties_path = os.path.join(application_path, 'env/prod.properties')

configs = Properties()
with open(properties_path, 'rb') as read_prop:
    configs.load(read_prop)

resource_path = os.path.join(application_path, configs.get('path.resource').data)
excel_file_path = resource_path + "/" + configs.get('scan.keywords').data

# Path to the PDFs to scan
pdf_scan_path = configs.get("scan.folder").data
# Read the Excel file into a DataFrame
xlsDataFrame = pd.read_excel(excel_file_path)

# Create an empty list to store the matched keywords
matched_keywords = []
unmatched_keywords = []


def cache_pdf_pages(reader):
    pages = []
    for page_num in range(len(reader.pages)):
        pdf_page = reader.pages[page_num]
        pdf_text = pdf_page.extract_text()
        pages.append(pdf_text)
    return pages


def found_in_pdf(search_key, pages, verbose):
    found = False
    for page in pages:
        if page.find(search_key) > -1:
            found = True
        if verbose:
            print(page)
        if found:  # break for loop
            break
    return found


def rename_that_bitch(old_file_name, provider):
    formatted_date = datetime.now().date().strftime("%Y-%m-%d")
    new_file_name_to_rename = formatted_date + "-" + provider + ".pdf"

    new_pdf_file_path = os.path.join(pdf_scan_path, new_file_name_to_rename)
    os.rename(old_file_name, new_pdf_file_path)
    print("File " + old_file_name + " has been renamed to " + new_pdf_file_path + ".")


def _main():
    if os.path.exists(pdf_scan_path) and os.path.isdir(pdf_scan_path):
        # List all items (files and directories) in the specified folder
        pdf_file_list = os.listdir(pdf_scan_path)

        # Iterate through the items and filter for files
        for pdf_file in pdf_file_list:
            # print("iterating over " + pdf_file)
            pdf_file_path = os.path.join(pdf_scan_path, pdf_file)  # Get the full path of the item

            if os.path.isfile(pdf_file_path) and (os.path.splitext(pdf_file_path)[1]).lower() == ".pdf":
                pdf_reader = pypdf.PdfReader(pdf_file_path)
                pdf_pages = cache_pdf_pages(pdf_reader)

                key_value_matched = False
                new_file_name = ""
                for index, row in xlsDataFrame.iterrows():
                    if key_value_matched:
                        break
                    search_key_found = False
                    for (columnName, columnData) in xlsDataFrame.items():   # remember file name
                        xls_key = row[columnName]
                        if xlsDataFrame.columns.get_loc(columnName) == 0:   # skip 1st column w file names
                            new_file_name = xls_key
                        else:
                            if not pd.isnull(xls_key):
                                if xls_key.strip() and found_in_pdf(xls_key, pdf_pages, False):
                                    search_key_found = True
                                    matched_keywords.append(columnName + " : " + xls_key + " found in file : " + pdf_file)
                                else:
                                    search_key_found = False
                                if not search_key_found:    # means we didn't find what we were looking for => next row!
                                    break       # break from xls row to search
                            else:
                                key_value_matched = True    # make sure we break row iter
                                rename_that_bitch(pdf_file_path, new_file_name)
                                break                       # means we found what we were looking for; break column iter
            else:
                print("os.path.isfile " + pdf_file_path + " is no file, nor .pdf!")

    # Print the matched keywords
    if matched_keywords:
        print("==============================================================Matched keywords found in the following PDFs:")
        for keyword in matched_keywords:
            print(keyword)

    if unmatched_keywords:
        print("==============================================================Unmatched keywords in following PDFs:")
        for keyword in unmatched_keywords:
            print(keyword)
    else:
        print("No matched keywords found in all the PDF.")


if __name__ == '__main__':
    _main()