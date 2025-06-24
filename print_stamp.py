from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

def save_last_serial(serial_number):
    """Save the last used serial number to a file"""
    try:
        with open("last_serial.txt", "w") as f:
            f.write(str(serial_number))
    except Exception as e:
        print(f"Error saving serial number: {str(e)}")

def load_last_serial():
    """Load the last used serial number from file"""
    try:
        with open("last_serial.txt", "r") as f:
            return f.read().strip()
    except FileNotFoundError:
        return None
    except Exception as e:
        print(f"Error loading serial number: {str(e)}")
        return None

def print_stamp(driver):
    print(">>> Starting e-Stamp printing process...")
    try:
        num_stamps = int(input("How many stamps do you want to print? "))
        print(f">>> You chose to print {num_stamps} stamps.")

        last_serial = load_last_serial()
        if last_serial:
            next_serial = str(int(last_serial) + 1).zfill(5)
            use_next = input(f"Next serial number is {next_serial}. Use it? (y/n): ").lower()
            if use_next == 'y':
                start_serial = next_serial
                print(f">>> Using serial number: {start_serial}")
            else:
                start_serial = input("Enter the 5-digit serial number to start with: ")
                print(f">>> User entered serial number: {start_serial}")
        else:
            start_serial = input("Enter the 5-digit serial number to start with: ")
            print(f">>> No previous serial found. Starting with: {start_serial}")

        if not start_serial.isdigit() or len(start_serial) != 5:
            print(">>> Invalid serial number. Please enter a valid 5-digit serial number.")
            return

        current_serial = int(start_serial)

        for i in range(num_stamps):
            print(f"\n>>> Processing stamp {i+1} of {num_stamps} with serial {str(current_serial).zfill(5)}")

            print(">>> Clicking on 'Generate Certificate' link...")
            generate_cert_link = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR,
                    'a[href="/eStampIndia/cert/certgen/CertGenServlet?sDoAction=DraftDetail&aAction=GenerateCertificate"].link2'))
            )
            generate_cert_link.click()

            print(">>> Clicking on certificate link...")
            try:
                cert_link = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "tr.bg-row1 a[href^='javascript:submissionDetails']"))
                )
                cert_link.click()
            except TimeoutException:
                print(">>> No stamp to print.")
                break

            print(">>> Clicking on 'Accept' button...")
            accept_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.NAME, "pAccept"))
            )
            accept_button.click()

            print(">>> Handling alert popup after accepting...")
            alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
            alert.accept()

            print(">>> Clicking on 'Print Stamp Certificate' button...")
            print_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.NAME, "PrintStampCertificateNew"))
            )
            print_button.click()

            print(">>> Handling print confirmation alert...")
            alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
            alert.accept()

            print(">>> Clicking on 'Final Print' button...")
            final_print_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "printBtn"))
            )
            final_print_button.click()

            print(">>> Entering serial number on form...")
            serial_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "serialNo"))
            )
            serial_input.clear()
            serial_input.send_keys(str(current_serial).zfill(5))

            print(">>> Waiting for system to populate certificate details...")

            # Wait for the <select> to have at least one valid option (value not empty, not disabled)
            try:
                WebDriverWait(driver, 15).until(
                    lambda d: any(
                        option.get_attribute("value") not in [None, "", "Select your option"]
                        and option.is_enabled()
                        for option in d.find_elements(By.CSS_SELECTOR, "#selectSl option")
                    )
                )
            except Exception:
                print("⚠️ Timeout waiting for valid option in selectSl dropdown.")

            # Get stationery serial number from valid option
            stationery_serial = ""
            try:
                select_elem = driver.find_element(By.ID, "selectSl")
                options = select_elem.find_elements(By.TAG_NAME, "option")
                # Find first valid option with a meaningful value
                valid_option = next(
                    (opt for opt in options if opt.get_attribute("value") not in [None, "", "Select your option"] and opt.is_enabled()),
                    None
                )
                if valid_option:
                    raw_value = valid_option.get_attribute("value")
                    if raw_value and "=" in raw_value:
                        stationery_serial = raw_value.split("=")[0].strip()
            except Exception as e:
                print(f"Could not read stationery serial: {e}")

            # Get certificate ID from page
            certificate_id = ""
            try:
                cert_elem = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//td/font[contains(@class, 'txt-body') and starts-with(text(), 'IN-GJ')]"))
                )
                certificate_id = cert_elem.text.strip()
            except Exception as e:
                print(f"⚠️ Could not read certificate ID: {e}")

            # Print results
            if certificate_id:
                print(f"Certificate ID: {certificate_id}")
            else:
                print("Certificate ID not found.")

            if stationery_serial:
                print(f"Stationery Serial Number : {stationery_serial}\n")
            else:
                print("Stationery Serial not found.\n")

            print(">>> Clicking 'Save' to store the serial number...")
            save_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.NAME, "Save"))
            )
            save_button.click()

            save_last_serial(current_serial)
            print(f">>> Serial number {str(current_serial).zfill(5)} saved successfully.")
            current_serial += 1

        print(f"\n>>> Total {num_stamps} Print completed. ")

    except ValueError:
        print(">>> Input error: Please enter valid numbers.")
    except Exception as e:
        print(f">>> Error during printing: {str(e)}")
