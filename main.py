import threading
from tkinter import *
from tkinter import ttk
from selenium import webdriver  # for operating the website
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as chromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
import pandas as pd

from colorama import init, Fore, Back, Style

DATA_FILE = 'test.xlsx'

INSTA_LOGIN_URL = 'https://www.instagram.com/?hl=zh' 
AUTHENTICATE_URL = 'https://www.instagram.com/accounts/login/two_factor?hl=zh&next=%2F'

ENGLISH = 'Times New Roman'
CHINESE = '微軟正黑體'

class CheckLogin:
    def __init__(self) -> None:
        self.__init_root__()
        self.__init_label__()
        self.__init_buttons__()
 
    def __start_button_onclick__(self):
        print('start button onclick')
        self.root.destroy()

    def __init_root__(self):
        self.root = Tk()
        self.root.resizable(False, False)
        self.root.geometry("300x300")
        self.root.title('InstaBlocker Check Login')

    def __init_label__(self):
        self.label = Label(self.root)
        self.label.config(
            text='Please finish the login process, then click the button below.', font=(ENGLISH, 14, 'bold'))
        self.label.pack(side=TOP, pady=5)

    def __init_buttons__(self):
        self.control_frame = Frame(self.root)
        self.control_frame.pack(side=TOP, fill='x', padx=15, pady=0)

        self.start_btn = Button(self.control_frame)
        self.start_btn.config(text='start', font=(ENGLISH, 14, 'bold'),
                              height=2, width=8, command=self.__start_button_onclick__)
        self.start_btn.pack(side=TOP, padx=10, pady=5)

    def run(self):
        self.root.mainloop()

class BlockadeLogView:
    def __init__(self, is_abort) -> None:
        self.__init_root__()
        self.__init_inner_frame__()
        self.__init_buttons__()

        self.is_aborted = is_abort

    def __init_root__(self):
        self.root = Tk()
        self.root.resizable(False, False)
        self.root.geometry("400x600")
        self.root.title('InstaBlocker Log View')

    def __scrollbar_resize__(self, event):
        size = (self.window_frame.winfo_reqwidth(),
                self.window_frame.winfo_reqheight())
        self.inner_canvas.config(scrollregion="0 0 %s %s" % size)

        if self.window_frame.winfo_reqwidth() != self.inner_canvas.winfo_width():
            self.inner_canvas.config(width=self.window_frame.winfo_reqwidth())

    def __abort_blockade__(self):
        self.is_aborted = True

    def __init_inner_frame__(self):
        self.inner_frame = Frame(self.root)
        self.inner_frame.pack(fill=BOTH, expand=True)

        self.inner_canvas = Canvas(self.inner_frame)
        self.inner_canvas.pack(side=LEFT, fill=BOTH, expand=True)

        self.scrollbar = ttk.Scrollbar(self.inner_frame, orient=VERTICAL,
                                       command=self.inner_canvas.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)

        self.inner_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.inner_canvas.bind('<Configure>', lambda e: self.inner_canvas.configure(
            scrollregion=self.inner_canvas.bbox('all')))

        self.window_frame = Frame(self.inner_canvas)

        self.inner_canvas.create_window(
            (0, 0), window=self.window_frame, anchor='nw')
        self.window_frame.bind('<Configure>', self.__scrollbar_resize__)

    def __init_buttons__(self):
        self.control_frame = Frame(self.root)
        self.control_frame.pack(side=BOTTOM, fill='x', padx=15, pady=0)

        self.start_btn = Button(self.control_frame)
        self.start_btn.config(text='abort', font=(ENGLISH, 14, 'bold'),
                              height=2, width=8, command=self.__abort_blockade__)
        self.start_btn.pack(side=TOP, padx=10, pady=5)

    def new_log(self, msg):
        log = Label(self.window_frame)
        log.config(text=msg, font=(ENGLISH, 14))
        log.pack(side=BOTTOM)

    def run(self):
        self.root.mainloop()

class InstaBlocker:
    def __init__(self) -> None:
        self.__init_data__()
        self.__init_path__()
        
    def __init_path__(self):
        self.path = {
            # BY XPATH
            'username_input': '//*[@id="loginForm"]/div/div[1]/div/label/input',
            'password_input': '//*[@id="loginForm"]/div/div[2]/div/label/input',
            'login_button': '//button[normalize-space()="登入"]',
            'verify_code_input': '/html/body/div[2]/div/div/div[2]/div/div/div[1]/section/main/div/div/div[1]/div[2]/form/div[1]/div/label/input',
            'confirm_button': '//button[normalize-space()="確認"]',
            'save_data_later_button': '//button[normalize-space()="稍後再說"]',
            'turn_on_notifications_button': '//button[normalize-space()="開啟通知"]',
            'blockade_button': '//button[normalize-space()="封鎖"]',
            'remove_blockade_button': '//button[normalize-space()="解除封鎖"]',
            'cancel_blockade_button': '//button[normalize-space()="取消"]',
            'inner_div': '/html/body/div[7]/div[1]/div/div[2]/div/div/div/div/div/div/div[2]',
            'close_blockade_button': '//button[normalize-space()="關閉"]',

            # BY CSS SELECTOR
            'options_button': '[aria-label="選項"]'
        }

    def __init_data__(self):
        try:
            self.data = pd.read_excel('qusun.ny-劣質粉絲名單.xlsx', header=7)
        except Exception as e:
            print('qusun.ny-劣質粉絲名單.xlsx not found')
            exit(1)

        self.namelist = self.data[self.data['評分'] <= -55][self.data.columns[1]].tolist()

    def __update_data__(self):
        header = pd.read_excel(DATA_FILE, nrows=7)

        with pd.ExcelWriter(DATA_FILE, engine='openpyxl') as writer:
            header.to_excel(writer, index=False)
            self.data.to_excel(writer, index=False, header=False, startrow=8)

    def start_driver(self) -> None:
        chrome_option = chromeOptions()
        chrome_option.add_argument('--log-level=3')
        chrome_option.add_argument('--start-maximized')
        chromedriver_autoinstaller.install()
        self.driver = webdriver.Chrome(options=chrome_option)
        
    def blockade(self, namelist) -> int:
        for username in namelist[:]:
            print(Fore.GREEN + f"{self.is_aborted}\n")

            if self.is_aborted:
                break

            self.driver.get(username)

            url_split = username.split('/')

            try:
                options_btn = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 
                    self.path['options_button'])))
                options_btn.click()
            except Exception as e:
                print('options_btn not found')
                print(format(e))
                self.blockade_log_view.new_log(f'{url_split[-1]} error: options_btn not found')
                continue

            try:
                remove_blockade_btn = WebDriverWait(self.driver, 1).until(
                    EC.presence_of_element_located((By.XPATH, 
                    self.path['remove_blockade_button'])))
                self.blockade_log_view.new_log(f'{url_split[-1]} blocked already')

                try:
                    cancel_blockade_btn = self.driver.find_element(By.XPATH, self.path['cancel_blockade_button'])
                    cancel_blockade_btn.click()
                except Exception as e:
                    print(format(e))

                continue
            except Exception as e:
                try:
                    blockade_btn = WebDriverWait(self.driver, 1).until(
                        EC.presence_of_element_located((By.XPATH, 
                        self.path['blockade_button'])))
                except Exception as e:
                    print('blockade_btn not found')
                    print(format(e))
                    self.blockade_log_view.new_log(f'{url_split[-1]} error: blockade_btn not found')
                    continue
            
            blockade_btn.click()

            try:
                div = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, 
                    self.path['inner_div'])))
                btns = div.find_elements(By.TAG_NAME, 'button')
                btns[0].click()

                self.blockade_log_view.new_log(f'{url_split[-1]} are blocked')
                namelist.remove(username)
            except Exception as e:
                print('check_blockade_btn not found')
                print(format(e))
                self.blockade_log_view.new_log(f'{url_split[-1]} error: check_blockade_btn not found')
                continue

            try:
                close_blockade_btn = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, 
                    self.path['close_blockade_button'])))
                close_blockade_btn.click()
            except Exception as e:
                print('close_blockade_btn not found')
                print(format(e))
                self.blockade_log_view.new_log(f'{url_split[-1]} error: close_btn not found')
                continue

        self.namelist = namelist
            
        return 0

    def __check_login__(self):
        self.check_login = CheckLogin()
        self.check_login.run()
    
    def __run_blockade__(self):
        self.blockade(self.namelist)

    def run(self) -> None:
        self.start_driver()

        check_login_threading = threading.Thread(target=self.__run_blockade__)
        check_login_threading.start()
        check_login_threading.join()

        self.is_aborted = False

        blockade_thread = threading.Thread(target=self.__run_blockade__)
        blockade_thread.start()

        self.blockade_log_view = BlockadeLogView(self.is_aborted)
        self.blockade_log_view.run()

        blockade_thread.join()
        self.blockade_log_view.root.destroy()
        self.__update_data__()


if __name__ == "__main__":
    init(autoreset=True)

    insta_blocker = InstaBlocker()
    insta_blocker.run()
