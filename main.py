# ==========================================================
# IMPORT LIBRARIES
# ==========================================================

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
import glob


# ==========================================================
# CONFIGURATION
# ==========================================================

DOWNLOAD_FOLDER = "rides"
EXCEL_FILE = "Invoices-facturas.xlsx"
OUTPUT_FILE = "facturas_resultado.xlsx"


# ==========================================================
# SETUP CHROME FOR AUTOMATIC DOWNLOADS
# ==========================================================

options = webdriver.ChromeOptions()

prefs = {
    "download.default_directory": os.path.abspath(DOWNLOAD_FOLDER),
    "download.prompt_for_download": False
}

options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)


# ==========================================================
# LOAD EXCEL DATA
# ==========================================================

excel = pd.read_excel(EXCEL_FILE)
total_invoices = len(excel)

print(f"Total invoices in Excel: {total_invoices}")


# ==========================================================
# STORE NOT FOUND INVOICES
# ==========================================================

not_found = []


# ==========================================================
# OPEN SRI PORTAL
# ==========================================================

driver.get("https://srienlinea.sri.gob.ec")

input("Login manually, go to 'Comprobantes Recibidos' and press ENTER...")


# ==========================================================
# FORMAT INVOICE NUMBER
# ==========================================================

def format_invoice(serie, numero):
    serie = str(serie).zfill(6)
    serie = serie[:3] + "-" + serie[3:]
    numero = str(numero).zfill(9)
    return f"Factura {serie}-{numero}"


# ==========================================================
# CHANGE PAGE IN SRI TABLE
# ==========================================================

def change_page():
    try:
        next_button = driver.find_element(By.CLASS_NAME, "ui-paginator-next")
        class_attr = next_button.get_attribute("class")

        if "ui-state-disabled" in class_attr:
            first_button = driver.find_element(By.CLASS_NAME, "ui-paginator-first")
            driver.execute_script("arguments[0].click();", first_button)
            print("Returning to first page")
        else:
            driver.execute_script("arguments[0].click();", next_button)
            print("Moving to next page")

        time.sleep(3)

    except Exception as e:
        print("Error changing page:", e)
        time.sleep(3)


# ==========================================================
# MAIN LOOP
# ==========================================================

for index, row in excel.iterrows():

    print("\n===================================")
    print(f"Processing invoice {index + 1} of {total_invoices}")
    print("===================================")

    ruc_excel = str(row["Ruc"])
    serie = row["Serie"]
    numero = row["Número"]

    invoice_excel = format_invoice(serie, numero)

    print("Searching:", invoice_excel)

    found = False
    pages_checked = 0


    while pages_checked < 30:

        rows = driver.find_elements(By.XPATH, "//tbody/tr")

        for table_row in rows:

            columns = table_row.find_elements(By.TAG_NAME, "td")

            if len(columns) < 3:
                continue

            ruc_sri = columns[1].text
            invoice_sri = columns[2].text

            if ruc_excel in ruc_sri and invoice_excel in invoice_sri:

                print("FOUND:", invoice_excel)

                invoice_number = invoice_excel.replace("Factura ", "")
                new_name = os.path.join(DOWNLOAD_FOLDER, invoice_number + ".pdf")

                if os.path.exists(new_name):
                    print("File already exists:", invoice_number)
                    excel.loc[index, "Estado"] = "ENCONTRADO"
                    found = True
                    break

                before = set(glob.glob(DOWNLOAD_FOLDER + "/*.pdf"))

                pdf_button = columns[10].find_element(By.TAG_NAME, "a")
                driver.execute_script("arguments[0].click();", pdf_button)

                print("Downloading RIDE...")

                new_file = None

                while new_file is None:
                    time.sleep(1)
                    after = set(glob.glob(DOWNLOAD_FOLDER + "/*.pdf"))
                    diff = after - before

                    if diff:
                        new_file = diff.pop()

                os.rename(new_file, new_name)

                print("Saved:", invoice_number)

                excel.loc[index, "Estado"] = "ENCONTRADO"

                found = True
                break

        if found:
            break

        pages_checked += 1
        change_page()


    if not found:
        print("NOT FOUND:", invoice_excel)
        excel.loc[index, "Estado"] = "NO ENCONTRADO"
        not_found.append(invoice_excel)


# ==========================================================
# SAVE RESULTS
# ==========================================================

excel.to_excel(OUTPUT_FILE, index=False)

print("\nResults saved in:", OUTPUT_FILE)


# ==========================================================
# FINAL SUMMARY
# ==========================================================

print("\n===================================")
print("NOT FOUND INVOICES")
print("===================================")

for invoice in not_found:
    print(invoice)

print("\nTotal not found:", len(not_found))

print("\n===================================")
print("PROCESS COMPLETED")
print("===================================")