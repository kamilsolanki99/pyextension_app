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

        # last_serial = load_last_serial()
        # if last_serial:
        #     use_last = input(f"Last used serial number was {last_serial}. Use next number? (y/n): ").lower()
        #     if use_last == 'y':
        #         start_serial = str(int(last_serial) + 1).zfill(5)
        #         print(f">>> Using next serial number: {start_serial}")
        #     else:
        #         start_serial = input("Enter the last 5-digit serial number: ")
        #         print(f">>> User entered serial number: {start_serial}")
        # else:
        #     start_serial = input("Enter the last 5-digit serial number: ")
        #     print(f">>> No previous serial found. Starting with: {start_serial}")

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

            # print(">>> Clicking on 'OK' button after alert...")
            # ok_button = WebDriverWait(driver, 5).until(
            #     EC.element_to_be_clickable((By.NAME, "ackButton"))
            # )
            # ok_button.click()

            print(">>> Clicking on 'Final Print' button...")
            final_print_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "printBtn"))
            )
            final_print_button.click()

            print(">>> Waiting for system to populate dynamic dropdown...")
            WebDriverWait(driver, 10).until(
                lambda d: len(d.find_elements(By.CSS_SELECTOR, "#selectSl option")) > 1
            )

            select_element = driver.find_element(By.ID, "selectSl")
            options = select_element.find_elements(By.TAG_NAME, "option")

            for option in options:
                value = option.get_attribute("value")
                if value and "=" in value:
                    actual_serial = value.split("=")[0]
                    print(f"🖨️ Dynamic Serial Detected from Website: {actual_serial}")
                    break

            # ow click save
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
