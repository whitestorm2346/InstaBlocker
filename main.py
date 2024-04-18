from tkinter import *
from tkinter import ttk
from selenium import webdriver  # for operating the website
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as chromeOptions
from webdriver_manager.chrome import ChromeDriverManager

INSTA_LOGIN_URL = 'https://www.instagram.com/?hl=zh' 
AUTHENTICATE_URL = 'https://www.instagram.com/accounts/login/two_factor?hl=zh&next=%2F'

class InstaBlocker:
    def __init__(self) -> None:
        self.__init_driver__()
        self.__init_ui__()

    def __init_driver__(self) -> None:
        chrome_option = chromeOptions()
        chrome_option.add_argument('--log-level=3')
        self.driver = webdriver.Chrome(
            executable_path=ChromeDriverManager(
                version='114.0.5735.90').install(),
            options=chrome_option
        )

    def __init_ui__(self) -> None:
        pass

    def login(self, username, password) -> int:
        self.driver.get(INSTA_LOGIN_URL)

        username_input = self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[1]/div/label/input')
        username_input.text = username

        password_input = self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[2]/div/label/input')
        password_input.text = password

        login_btn = self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[3]/button')
        login_btn.click()

    def run(self, username, password, namelist) -> None:
        login_state = self.login(username, password)

        if login_state == 0:
            pass

        for username in namelist:
            self.driver.get(username)

            options_btn = self.driver.find_element(By.XPATH, '//*[@id="mount_0_0_Oh"]/div/div/div[2]/div/div/div[1]/div[1]/div[2]/div[2]/section/main/div/header/section/div[1]/div[3]/div/div/svg')
            options_btn.click()

            blockade_btn = self.driver.find_element(By.XPATH, '/html/body/div[7]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div/button[1]')
            
            if blockade_btn.text == '解除封鎖':
                cancel_blockade_btn = self.driver.find_element(By.XPATH, '/html/body/div[6]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div/button[6]')
                cancel_blockade_btn.click()
                continue
            else:
                blockade_btn.click()

            check_blockade_btn = self.driver.find_element(By.XPATH, '/html/body/div[7]/div[1]/div/div[2]/div/div/div/div/div/div/div[2]/button[1]')
            check_blockade_btn.click()

            close_blockade_btn = self.driver.find_element(By.XPATH, '/html/body/div[7]/div[1]/div/div[2]/div/div/div/div/div/div/div[2]/button')
            close_blockade_btn.click()


if __name__ == "__main__":
    insta_blocker = InstaBlocker()
    insta_blocker.run([])