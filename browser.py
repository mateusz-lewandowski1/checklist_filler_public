import time
from datetime import datetime
from selenium import webdriver
from selenium.common import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.edge.options import Options

class Browser:
    def __init__(self):
        """
        Initialize the browser with Edge options and start time.
        """
        self.start_time = time.time()
        self.readable_start_time = datetime.fromtimestamp(self.start_time)
        self.formatted_start_time = self.readable_start_time.strftime('%Y-%m-%d %H:%M:%S')
        print(f'Starting up...\n{self.formatted_start_time}')

        # Set headless option for Edge browser
        self.edge_options = Options()
        self.edge_options.add_argument("--headless")
        self.edge_options.add_argument("--disable-gpu")
        self.browser = webdriver.Edge(options=self.edge_options)

    def open_page(self, url):
        """
        Open the specified URL and maximize the browser window.
        """
        print(f'Opening: {url}')
        self.browser.get(url)
        self.browser.maximize_window()

    def switch_to_iframe(self):
        """
        Switch to the iframe within the page.
        """
        try:
            WebDriverWait(self.browser, 15).until(
                EC.frame_to_be_available_and_switch_to_it((By.XPATH, "/html/body/div[2]/div[1]/iframe"))
            )
            print("Switched to iframe")
        except Exception as e:
            print(f"Failed to switch to iframe: {e}")

    def check_tick(self):
        """
        Check if the checklist is already ticked.
        """
        try:
            element = WebDriverWait(self.browser, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, "/html/body/div[1]/div/div/div/div[4]/div/div/div[7]/div/div/div/div/div[1]/div/div/div/div[66]/div/div/div/div[3]/div/div/div/div[1]/div[1]/ul/li")
                )
            )
            self.browser.execute_script("arguments[0].scrollIntoView(true);", element)
            time.sleep(2)
            print('The checklist is already ticked')
            return True
        except TimeoutException:
            print("Element not found within the time limit")
            return False

    def click_button(self, xpath):
        """
        Click a button identified by its XPath.
        """
        WebDriverWait(self.browser, 15).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        ).click()

    def element_presence(self, xpath):
        """
        Wait for the presence of an element identified by its XPath.
        """
        WebDriverWait(self.browser, 15).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )

    def end_time(self):
        """
        Calculate and print the total time elapsed for the automation process.
        """
        end_time = time.time()
        elapsed_seconds = end_time - self.start_time
        minutes = int(elapsed_seconds // 60)
        seconds = int(elapsed_seconds % 60)
        print(f"Total time elapsed: {minutes} minute(s) and {seconds} second(s)")
