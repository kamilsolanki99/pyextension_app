import os
import time
import warnings
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from generate_stamp import generate_stamp
from print_stamp import print_stamp
from delete_stamp import delete_stamp

warnings.filterwarnings('ignore')
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'

def show_welcome():
    """Display welcome message and tool information"""
    os.system("cls" if os.name == "nt" else "clear")
    print("=" * 50)
    print("           e-Stamp Automator v1.0")
    print("=" * 50)
    print("\nDeveloped by: KAMIL SOLANKI")
    print("\nWelcome to the e-Stamp Automation Tool!")
    print("This tool will help you manage your e-Stamp certificates efficiently.")
    print("\n" + "=" * 50)
    time.sleep(2)

def setup_driver():
    """Initialize and configure the Chrome WebDriver"""
    options = Options()
    # options.add_argument("--start-maximized")
    # options.add_argument("--guest")
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def wait_and_find_element(driver, by, value, timeout=10):
    """Utility function to wait for and find an element"""
    element = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )
    return element

def launch_browser():
    """Launch the browser and navigate to e-Stamp portal"""
    os.system("title e-stamp Automator" if os.name == "nt" else "")
    print("\nLaunching e-Stamp Portal...")

    url = "https://www.shcilestamp.com/eStampIndia/useradmin/UserAdminLoginServlet?rDoAction=LoadLoginPage"
    user_id = "gjdpipati"
    password = "Ddpkm@2016"

    try:
        driver = setup_driver()
        driver.get(url)

        user_input = driver.find_element(By.NAME, "userId")
        user_input.clear()
        user_input.send_keys(user_id)

        pass_input = driver.find_element(By.NAME, "userPwd")
        pass_input.clear()
        pass_input.send_keys(password)
        
        return driver
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

def wait_for_login(driver):
    """Wait for user to complete manual login process"""
    print("Waiting for login ...")
    try:
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Logout')]"))
        )
        print("Login successful!")
        return True
    except:
        print("Login not detected. Timeout.")
        return False

def post_login_menu(driver):
    print("\n=== Welcome to e-Stamp Automator ===")
    while True:
        print("\nChoose an action:")
        print("1. Generate")
        print("2. Print")
        print("3. Delete")
        print("4. Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            generate_stamp(driver)
        elif choice == '2':
            print_stamp(driver)
        elif choice == '3':
            delete_stamp(driver)
        elif choice == '4':
            print("\nExiting...")
            driver.quit()
            break
        else:
            print("Invalid choice. Please try again.")

def main():
    show_welcome()
    driver = launch_browser()
    if driver is None:
        print("Failed to launch browser. Exiting...")
        return

    if wait_for_login(driver):
        post_login_menu(driver)
    else:
        print("Exiting due to login failure.")
        driver.quit()

if __name__ == "__main__":
    main()
