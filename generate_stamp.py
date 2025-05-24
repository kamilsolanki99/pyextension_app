import time
import os
import sys
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def read_stamp_data():
    try:
        if getattr(sys, 'frozen', False):
            current_dir = os.path.dirname(sys.executable)
        else:
            current_dir = os.path.dirname(os.path.abspath(__file__))

        excel_path = os.path.join(current_dir, 'stamp_data.xlsx')

        if not os.path.exists(excel_path):
            print(f">>> Error: Excel file not found at {excel_path}")
            return None

        df = pd.read_excel(excel_path)
        if df is not None and not df.empty:
            print(">>> Excel file contents:")
            print(df.head())
            return df
        else:
            print(">>> Excel file is empty")
            return None
    except Exception as e:
        print(f">>> Error reading Excel file: {str(e)}")
        return None

def handle_all_alerts(driver, max_alerts=3):
    for _ in range(max_alerts):
        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            print(f">>> Alert handled: {alert.text}")
            alert.accept()
            time.sleep(0.01)
        except:
            break

def generate_stamp(driver):
    print(">>> Generating e-Stamp...")    
    
    # Ask for number of stamps to generate
    while True:
        try:
            num_stamps = int(input("How many stamps do you want to generate? "))
            if num_stamps > 0:
                break
            else:
                print("Please enter a positive number.")
        except ValueError:
            print("Please enter a valid number.")
    
    print(f">>> Will generate {num_stamps} stamps...")
    
    # Get stamp duty type and article choice once for all stamps
    try:
        for stamp_num in range(num_stamps):
            print(f"\n>>> Generating stamp {stamp_num + 1} of {num_stamps}")
            
            create_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((
                    By.CSS_SELECTOR, 
                    'a[href="/eStampIndia/submission/SubmissionServlet?rDoAction=LoadStampDuty"].link2'
                ))
            )
            create_link.click()
            time.sleep(0.01)
            print(">>> Clicked on Create Submission link")

            # Only ask for choices on first iteration
            if stamp_num == 0:
                while True:
                    print("\nPlease choose stamp duty type:")
                    print("1. Registerable Stamp Duty")
                    print("2. Non-Registerable Stamp Duty")
                    choice = input("Enter your choice (1 or 2): ")
                    global_choice = choice  # Store for future iterations
                    
                    if choice == "1":
                        radio_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((
                                By.CSS_SELECTOR, 
                                'input[name="StampDuty"][onclick="nonRegReset()"]'
                            ))
                        )
                        radio_button.click()
                        print(">>> Selected Registerable radio button")

                        print("\nAvailable Registerable Articles:")
                        print("1. Pledge, Loan or Hypothecation - Movable Property Loan/Debt not Exceed Rs. 1 Cr [6(2)(a)(i)]")
                        print("2. Power of Attorney (in any other case) [45 (h)]")
                        global_article_choice = input("Enter article choice ")
                        
                        dropdown = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, 'RegSD'))
                        )
                        dropdown.click()
                        
                        if global_article_choice == "1":
                            global_article_value = "GJ-RG-96"
                        elif global_article_choice == "2":
                            global_article_value = "GJ-RG-91"
                        else:
                            print("Invalid article choice ")
                            continue
                        
                        article_option = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, f'option[value="{global_article_value}"]'))
                        )
                        article_option.click()
                        print(f">>> Selected article option {global_article_choice}")
                        break

                    elif choice == "2":
                        radio_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((
                                By.CSS_SELECTOR, 
                                'input[name="StampDuty"][onclick="regReset()"]'
                            ))
                        )
                        radio_button.click()
                        print(">>> Selected Non-Registerable radio button")

                        print("\nAvailable Non-Registerable Articles:")
                        print("1. Agreement (not otherwise provided for) [5(h)]")
                        print("2. Affidavit [4]")
                        print("3. Letter of Guarantee [32]")
                        global_article_choice = input("Enter article choice ")
                        
                        dropdown = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, 'NonRegSD'))
                        )
                        dropdown.click()

                        if global_article_choice == "1":
                            global_article_value = "GJ-NRG-33H"
                        elif global_article_choice == "2":
                            global_article_value = "GJ-NRG-32"
                        elif global_article_choice == "3":
                            global_article_value = "GJ-NRG-66"
                        else:
                            print("Invalid article choice")
                            continue

                        article_option = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, f'option[value="{global_article_value}"]'))
                        )
                        article_option.click()
                        print(f">>> Selected article option {global_article_choice}")
                        break
                    else:
                        print("Invalid choice. Please enter 1 or 2.")
            else:
                # For subsequent iterations, use the stored choices
                if global_choice == "1":
                    radio_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((
                            By.CSS_SELECTOR, 
                            'input[name="StampDuty"][onclick="nonRegReset()"]'
                        ))
                    )
                    radio_button.click()
                    
                    dropdown = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, 'RegSD'))
                    )
                    dropdown.click()
                    
                    article_option = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, f'option[value="{global_article_value}"]'))
                    )
                    article_option.click()
                else:
                    radio_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((
                            By.CSS_SELECTOR, 
                            'input[name="StampDuty"][onclick="regReset()"]'
                        ))
                    )
                    radio_button.click()
                    
                    dropdown = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, 'NonRegSD'))
                    )
                    dropdown.click()
                    
                    article_option = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, f'option[value="{global_article_value}"]'))
                    )
                    article_option.click()
                print(">>> Selected previous choices automatically")

            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, "pNext"))
            )
            next_button.click()
            print(">>> Clicked Next button")

            print("\n>>> Reading data from Excel...")
            df = read_stamp_data()

            if df is not None:
                try:                    # Define all mandatory fields with their IDs
                    mandatory_fields = {
                        "Purchased_by": "TextField6Mand",
                        "Description": "TextArea8Mand",
                        "First_Party": "TextField11Mand",
                        "Second_Party": "TextField18Mand",
                        "Stamp_Duty_Paid_By": "TextField24Mand",
                        "First_Party_Mobile": "fpMobNo",
                        "Second_Party_Mobile": "spMobNo",
                        "Stamp_Duty_Amount": "TextField28Mand"
                    }                    # Get data from Excel if available
                    row_data = {}
                    if df is not None:
                        row = df.iloc[0]
                        for col_name in mandatory_fields.keys():
                            value = row.get(col_name, "")
                            # Handle NaN values
                            if pd.isna(value):
                                value = ""
                            row_data[col_name] = str(value).strip()

                    # Fill all mandatory fields (with data or blank)
                    for field_name, element_id in mandatory_fields.items():
                        try:
                            field = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, element_id))
                            )
                            driver.execute_script("arguments[0].scrollIntoView(true);", field)
                            time.sleep(0.01)
                            field.clear()
                            time.sleep(0.01)
                            
                            # Use Excel data if available, otherwise leave blank
                            value = row_data.get(field_name, "") if row_data else ""
                            field.send_keys(value)
                            
                            if value:
                                print(f">>> Successfully filled {field_name} with: {value}")
                            else:
                                print(f">>> Left {field_name} blank")
                        except Exception as e:
                            print(f">>> Failed to handle {field_name}: {str(e)}")

                    try:
                        save_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.NAME, "pSave"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView(true);", save_button)
                        time.sleep(0.01)
                        save_button.click()
                        print(">>> Clicked Save button")
                        handle_all_alerts(driver)  # Automatically handle up to 3 alerts
                        
                        if stamp_num < num_stamps - 1:  # If not the last stamp
                            print(">>> Preparing for next stamp...")
                        else:
                            print(">>> All stamps generated successfully!")
                            
                    except Exception as e:
                        print(f">>> Failed to click Save button: {str(e)}")

                except Exception as e:
                    print(f">>> Error in filling form: {str(e)}")
            else:
                print(">>> No data found in Excel file or error reading file")

    except Exception as e:
        print(f">>> Error during e-Stamp generation: {str(e)}")
