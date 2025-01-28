import time
import pandas as pd
from browser import Browser
from excel_handler import ExcelHandler

# Constants
WEBSITE_PATH = 'https://apps.powerapps.com/play/e/a3c669f6-ac2e-4e77-ad43-beab3e15bee7/a/d97f7b85-0658-4f12-a9ed-05fff961e2d7?tenantId=5b973f99-77df-4beb-b27d-aa0c70b8482c'
EXCEL_PATH = r'C:\Users\EV683XP\OneDrive - EY\Desktop\checklist\SN - Knowledge Article.xlsx'

# Load Excel data
df = pd.read_excel(EXCEL_PATH)

def run_automation(browser, excel_handler):
    """
    Run the automation process to fill out the decommission checklist.
    """
    browser.click_button("/html/body/div[1]/div/div/div/div[2]/div/div/div[15]/div/div/div/div/button/div/div")
    print("Asset Decommission button clicked")

    for index, row in df.iterrows():
        asset_name_value = row.iloc[4]
        status = row.iloc[5]

        asset_name = WebDriverWait(browser.browser, 15).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[3]/div/div/div[24]/div/div/div/div/input"))
        )
        asset_name.click()
        asset_name.clear()
        asset_name.send_keys(asset_name_value)
        print(f"Asset name '{asset_name_value}' populated")
        time.sleep(7)

        asset_name_text1, asset_name_text2, asset_name_text3 = None, None, None

        try:
            asset_name_xpath1 = WebDriverWait(browser.browser, 5).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[3]/div/div/div[47]/div/div/div/div/div[2]/div/div/div[1]/div[2]/div/div[1]/div/div/div/div/div"))
            )
            asset_name_text1 = asset_name_xpath1.text
        except (NoSuchElementException, TimeoutException):
            print("Element for asset_name_xpath1 not found")

        try:
            asset_name_xpath2 = WebDriverWait(browser.browser, 5).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[3]/div/div/div[47]/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div/div/div/div/div"))
            )
            asset_name_text2 = asset_name_xpath2.text
        except (NoSuchElementException, TimeoutException):
            print("Element for asset_name_xpath2 not found")

        try:
            asset_name_xpath3 = WebDriverWait(browser.browser, 5).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[3]/div/div/div[47]/div/div/div/div/div[2]/div/div/div[3]/div[2]/div/div[1]/div/div/div/div/div"))
            )
            asset_name_text3 = asset_name_xpath3.text
        except (NoSuchElementException, TimeoutException):
            print("Element for asset_name_xpath3 not found")

        if asset_name_text1 == asset_name_value:
            try:
                browser.click_button("/html/body/div[1]/div/div/div/div[3]/div/div/div[47]/div/div/div/div/div[2]/div/div/div[1]/div[2]/div/div[7]/div/div/div/div")
                print("Drop-down button for xpath1 clicked")
            except:
                print(f"Project {asset_name_value} not found for xpath1")
                continue
        elif asset_name_text2 == asset_name_value:
            try:
                browser.click_button("/html/body/div[1]/div/div/div/div[3]/div/div/div[47]/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[7]/div/div/div/div")
                print("Drop-down button for xpath2 clicked")
            except:
                print(f"Project {asset_name_value} not found for xpath2")
                continue
        elif asset_name_text3 == asset_name_value:
            try:
                browser.click_button("/html/body/div[1]/div/div/div/div[3]/div/div/div[47]/div/div/div/div/div[2]/div/div/div[3]/div[2]/div/div[7]/div/div/div/div")
                print("Drop-down button for xpath3 clicked")
            except:
                print(f"Project {asset_name_value} not found for xpath3")
                continue
        else:
            print(f"Asset name '{asset_name_value}' not found in any of the provided XPaths")
            continue

        try:
            decomm_button = WebDriverWait(browser.browser, 15).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[4]/div/div/div[18]/div/div/div/div/button/div"))
            )
            is_disabled = browser.browser.execute_script("return arguments[0].classList.contains('disabled');", decomm_button)

            if not is_disabled:
                decomm_button.click()
                time.sleep(7)
                print("Decomm checklist button clicked")
            else:
                raise Exception("Decomm checklist button is greyed out or disabled")
        except Exception as e:
            print(f'Decomm checklist button greyed out/unidentified - {e}')
            time.sleep(2)
            browser.click_button("/html/body/div[1]/div/div/div/div[4]/div/div/div[8]/div/div/div/div/button/div")
            print("<< button clicked")
            continue

        if browser.check_tick():
            browser.save_button()
            excel_handler.mark_green("Sheet1", index)
            browser.go_back()
            continue

        dropdown_xpath = "/html/body/div[1]/div/div/div/div[4]/div/div/div[7]/div/div/div/div/div[1]/div/div/div/div[66]/div/div/div/div[3]/div/div/div/div[1]/div[1]"
        dropdown_element = WebDriverWait(browser.browser, 15).until(
            EC.presence_of_element_located((By.XPATH, dropdown_xpath))
        )
        time.sleep(1)
        browser.browser.execute_script("arguments[0].scrollIntoView(true);", dropdown_element)
        time.sleep(2)
        browser.click_button(dropdown_xpath)
        print("Selected dropdown clicked")
        time.sleep(2)

        if status == "Done":
            browser.element_presence("/html/body/div[1]/div/div/div[2]/div/div/ul/li[1]")
            browser.click_button("/html/body/div[1]/div/div/div[2]/div/div/ul/li[1]")
            print("Done button clicked")
        elif status == "Not Required":
            browser.element_presence("/html/body/div[1]/div/div/div[2]/div/div/ul/li[3]")
            browser.click_button("/html/body/div[1]/div/div/div[2]/div/div/ul/li[3]")
            print("Not Required button clicked")
        elif status == "Pending":
            browser.element_presence("/html/body/div[1]/div/div/div[2]/div/div/ul/li[2]")
            browser.click_button("/html/body/div[1]/div/div/div[2]/div/div/ul/li[2]")
            print("Pending button clicked")
        else:
            print(f"Unknown status: {status}")
        time.sleep(2)
        browser.save_button()
        excel_handler.mark_green("Sheet1", index)
        browser.go_back()

if __name__ == '__main__':
    browser = Browser()
    excel_handler = ExcelHandler(EXCEL_PATH)
    time.sleep(2)
    browser.open_page(WEBSITE_PATH)
    time.sleep(15)
    browser.switch_to_iframe()
    run_automation(browser, excel_handler)
    browser.end_time()
    time.sleep(10)
