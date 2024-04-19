import threading
from time import sleep
from tkinter import *
from tkinter import ttk
from selenium import webdriver  # for operating the website
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as chromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
import pandas as pd

INSTA_LOGIN_URL = 'https://www.instagram.com/?hl=zh' 
AUTHENTICATE_URL = 'https://www.instagram.com/accounts/login/two_factor?hl=zh&next=%2F'

ENGLISH = 'Times New Roman'
CHINESE = '微軟正黑體'

MY_USERNAME = 'andrew770426@gmail.com'
MY_PASSWORD = 'Whitestorm2346'

class MainUI:
    def __init__(self) -> None:
        self.init_main_frame()
        self.init_login_frame()
        self.init_buttons()

        self.threads = []  # init thread

    def __show_password_onclick__(self):
        if self.show_password_value.get():
            self.password_entry.config(show='')
        else:
            self.password_entry.config(show='\u25CF')

    def start_btn_onclick(self):
        pass

    def quit_btn_onclick(self):
        pass

    def init_main_frame(self) -> None:
        self.root = Tk()
        self.root.resizable(False, False)
        self.root.geometry("350x350")
        self.root.title('InstaBlocker')

    def init_login_frame(self) -> None:
        self.login_frame = LabelFrame(self.root)
        self.login_frame.config(text=' Instagram Login ', font=(ENGLISH, 12))
        self.login_frame.pack(side=TOP, fill='x', padx=10, pady=10)

        self.username_label = Label(self.login_frame)
        self.username_label.config(
            text='User Name', font=(ENGLISH, 14, 'bold'))
        self.username_label.pack(side=TOP, pady=5)

        self.username = StringVar(self.login_frame)
        self.username_entry = Entry(self.login_frame)
        self.username_entry.config(
            font=(ENGLISH, 12), textvariable=self.username)
        self.username_entry.pack(side=TOP, fill='x', padx=5, pady=10)

        self.password_label = Label(self.login_frame)
        self.password_label.config(
            text='Password', font=(ENGLISH, 14, 'bold'))
        self.password_label.pack(side=TOP, pady=5)

        self.password = StringVar(self.login_frame)
        self.password_entry = Entry(self.login_frame)
        self.password_entry.config(
            font=(ENGLISH, 12), textvariable=self.password, show='\u25CF')
        self.password_entry.pack(side=TOP, fill='x', padx=5, pady=10)

        # Set Login Inner Frame

        self.login_inner_frame = Frame(self.login_frame)
        self.login_inner_frame.pack(side=TOP, fill='x')

        self.show_password_value = BooleanVar()
        self.show_password_checkbox = Checkbutton(
            self.login_inner_frame, text='Show Password',
            variable=self.show_password_value, command=self.__show_password_onclick__)
        self.show_password_checkbox.pack(side=LEFT, padx=10)

    def init_buttons(self) -> None:
        self.start_btn = Button(self.root)
        self.start_btn.config(text='start', font=(ENGLISH, 14, 'bold'),
                              height=2, width=8, command=self.start_btn_onclick)
        self.start_btn.pack(side=LEFT, padx=10)

        self.quit_btn = Button(self.root)
        self.quit_btn.config(text='quit', font=(ENGLISH, 14, 'bold'),
                             height=2, width=8, command=self.quit_btn_onclick)
        self.quit_btn.pack(side=LEFT, padx=15)

    def run(self) -> None:
        self.root.mainloop()

class InstaBlocker:
    def __init__(self) -> None:
        self.__init_app__()

    def start_driver(self) -> None:
        chrome_option = chromeOptions()
        chrome_option.add_argument('--log-level=3')
        chrome_option.add_argument('--start-maximized')
        chromedriver_autoinstaller.install()
        self.driver = webdriver.Chrome(options=chrome_option)

    def __init_app__(self) -> None:
        self.app = MainUI()

    def login(self, username, password) -> int:
        self.driver.get(INSTA_LOGIN_URL)

        try:
            username_input = self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[1]/div/label/input')
        except Exception as e:
            return 1
        
        username_input.clear()
        username_input.send_keys(username)

        password_input = self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[2]/div/label/input')
        password_input.clear()
        password_input.send_keys(password)

        login_btn = self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[3]/button')
        login_btn.click()

        sleep(5)

    def authentication(self): 
        if self.driver.current_url != AUTHENTICATE_URL:
            return
        
        try:
            verify_code_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                '/html/body/div[2]/div/div/div[2]/div/div/div[1]/section/main/div/div/div[1]/div[2]/form/div[1]/div/label/input')))
            verify_code_input.clear()

            verify_code = input('輸入驗證碼: ')

            verify_code_input.send_keys(verify_code)
        except Exception as e:
            print(format(e))

        confirm_btn = self.driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/section/main/div/div/div[1]/div[2]/form/div[2]/button')
        confirm_btn.click()

        try:
            save_data_later_btn = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[1]/div[2]/section/main/div/div/div/div/div')))
            save_data_later_btn.click()
        except Exception as e:
            print(format(e))

        try:
            turn_on_notifications_btn = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                '/html/body/div[3]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div[3]/button[2]')))
            turn_on_notifications_btn.click()
        except Exception as e:
            print(format(e))

    def blockade(self, namelist) -> None:
        for username in namelist:
            self.driver.get(username)

            try:
                options_btn = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, 
                    '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[1]/div[2]/div[2]/section/main/div/header/section/div[1]/div[3]/div')))
                options_btn.click()
            except Exception as e:
                print(format(e))

            try:
                blockade_btn = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, 
                    '/html/body/div[7]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div/button[1]')))
            except Exception as e:
                print(format(e))
            
            if blockade_btn.text == '解除封鎖':
                cancel_blockade_btn = self.driver.find_element(By.XPATH, '/html/body/div[6]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div/button[6]')
                cancel_blockade_btn.click()
                continue
            else:
                blockade_btn.click()

            try:
                check_blockade_btn = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, 
                    '/html/body/div[7]/div[1]/div/div[2]/div/div/div/div/div/div/div[2]/button[1]')))
                check_blockade_btn.click()
            except Exception as e:
                print(format(e))

            try:
                close_blockade_btn = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, 
                    '/html/body/div[7]/div[1]/div/div[2]/div/div/div/div/div/div/div[2]/button')))
                close_blockade_btn.click()
            except Exception as e:
                print(format(e))

    def run(self) -> None:
        # self.app.run()
        self.start_driver()
        return


if __name__ == "__main__":
    data = pd.read_excel('qusun.ny-劣質粉絲名單.xlsx', header=7)
    filtered_urls = data[data['評分'] <= -55][data.columns[1]].tolist()

    insta_blocker = InstaBlocker()
    insta_blocker.run()
    insta_blocker.login(MY_USERNAME, MY_PASSWORD)
    insta_blocker.authentication()
    insta_blocker.blockade(filtered_urls)
