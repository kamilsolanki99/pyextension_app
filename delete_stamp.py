import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def delete_stamp(driver):
    print(">>> Starting e-Stamp deletion process...")
    try:
        num_stamps = int(input("How many stamps do you want to delete? "))
        print(f">>> You chose to delete {num_stamps} stamps.")

        for i in range(num_stamps):
            print(f"\n>>> Deleting stamp {i+1} of {num_stamps}")

            print(">>> Finding 'Generate Certificate' link...")
            generate_cert_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 
                    'a[href="/eStampIndia/cert/certgen/CertGenServlet?sDoAction=DraftDetail&aAction=GenerateCertificate"].link2'))
            )
            print(">>> Clicking on 'Generate Certificate' link...")
            generate_cert_link.click()

            print(">>> Waiting for and clicking on certificate link...")
            try:
                cert_link = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "tr.bg-row1 a[href^='javascript:submissionDetails']"))
                )
                cert_link.click()
            except TimeoutException:
                print(">>> No stamp to be deleted.")
                break  # Stop further attempts

            print(">>> Waiting for and clicking 'Delete' button...")
            delete_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, "pDelete"))
            )
            delete_button.click()

            print(">>> Handling alert popup after clicking delete...")
            alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
            alert.accept()

            print(f">>> e-Stamp {i+1} deleted successfully!")
            time.sleep(0.01)

        print(f"\n>>> Deletion process completed.")

    except ValueError:
        print(">>> Input error: Please enter a valid number.")
    except Exception as e:
        print(f">>> Error during deletion: {str(e)}")
