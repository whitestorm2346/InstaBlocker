import threading
from tkinter import *
from tkinter import ttk
from selenium import webdriver  # for operating the website
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as chromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import chromedriver_autoinstaller
import pandas as pd

from colorama import init, Fore, Style

DATA_FILE = 'qusun.ny-劣質粉絲名單.xlsx'

INSTA_LOGIN_URL = 'https://www.instagram.com/?hl=zh' 
AUTHENTICATE_URL = 'https://www.instagram.com/accounts/login/two_factor?hl=zh&next=%2F'

ENGLISH = 'Times New Roman'
CHINESE = '微軟正黑體'

is_aborted = False

class CheckLogin:
    def __init__(self) -> None:
        self.__init_root__()
        self.__init_label__()
        self.__init_buttons__()
 
    def __start_button_onclick__(self):
        self.root.destroy()

    def __init_root__(self):
        self.root = Tk()
        self.root.resizable(False, False)
        self.root.geometry("300x180")
        self.root.title('InstaBlocker Check Login')

    def __init_label__(self):
        self.label1 = Label(self.root)
        self.label1.config(
            text='Please finish the login process,', font=(ENGLISH, 14, 'bold'))
        self.label1.pack(side=TOP, pady=5)
        
        self.label2 = Label(self.root)
        self.label2.config(
            text='then click the button below.', font=(ENGLISH, 14, 'bold'))
        self.label2.pack(side=TOP, pady=5)

    def __init_buttons__(self):
        self.control_frame = Frame(self.root)
        self.control_frame.pack(side=BOTTOM, fill='x', padx=15, pady=10)

        self.start_btn = Button(self.control_frame)
        self.start_btn.config(text='start', font=(ENGLISH, 14, 'bold'),
                              height=2, width=8, command=self.__start_button_onclick__)
        self.start_btn.pack(side=TOP, padx=10, pady=5)

    def run(self):
        self.root.mainloop()

class BlockadeLogView:
    def __init__(self) -> None:
        self.__init_root__()
        self.__init_inner_frame__()
        self.__init_buttons__()

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
        global is_aborted

        print(Fore.YELLOW + 'abort button on click')
        is_aborted = True

        self.root.destroy()

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
            self.data = pd.read_excel(DATA_FILE, header=7)
        except Exception as e:
            print(f'{DATA_FILE} not found')
            exit(1)

        self.data = self.data[self.data['評分'] <= -55]
        self.namelist = self.data[self.data.columns[1]].tolist()

    def __update_data__(self):
        header = pd.read_excel(DATA_FILE, nrows=7)
        self.data = self.data[self.data[self.data.columns[1]].isin(self.namelist)]

        with pd.ExcelWriter(DATA_FILE, engine='openpyxl') as writer:
            header.to_excel(writer, index=False)
            self.data.to_excel(writer, index=False, header=False, startrow=8)

    def start_driver(self) -> None:
        chrome_option = chromeOptions()
        chrome_option.add_argument('--log-level=3')
        chrome_option.add_argument('--start-maximized')
        chromedriver_autoinstaller.install()
        self.driver = webdriver.Chrome(options=chrome_option)
        
    def blockade(self) -> int:
        global is_aborted

        blockced = []

        for username in self.namelist:
            url_split = username.split('/')

            if is_aborted:
                break

            print('\n' + Style.BRIGHT + '[account] ' + Fore.CYAN + f'{url_split[-1]}')

            self.driver.get(username)

            try:
                options_btn = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 
                    self.path['options_button'])))
                options_btn.click()
            except TimeoutException as e:
                try:
                    print(Style.BRIGHT + '[status] ' + Fore.LIGHTRED_EX + 'account not found')
                    self.blockade_log_view.new_log(f'{url_split[-1]} error: account not found')
                except Exception as e:
                    pass

                blockced.append(username)

                continue

            try:
                remove_blockade_btn = WebDriverWait(self.driver, 1).until(
                    EC.presence_of_element_located((By.XPATH, 
                    self.path['remove_blockade_button'])))
                
                try:
                    print(Style.BRIGHT + '[status] ' + Fore.YELLOW + 'blocked already')
                    self.blockade_log_view.new_log(f'{url_split[-1]} blocked already')
                except Exception as e:
                    pass

                try:
                    cancel_blockade_btn = self.driver.find_element(By.XPATH, self.path['cancel_blockade_button'])
                    cancel_blockade_btn.click()
                except NoSuchElementException as e:
                    print(Fore.RED + 'cancel_blockade_btn not found')

                blockced.append(username)

                continue
            except TimeoutException as e:
                try:
                    blockade_btn = WebDriverWait(self.driver, 1).until(
                        EC.presence_of_element_located((By.XPATH, 
                        self.path['blockade_button'])))
                except TimeoutException as e:
                    print(Fore.RED + 'blockade_btn not found')

                    try:
                        self.blockade_log_view.new_log(f'{url_split[-1]} error: blockade_btn not found')
                    except Exception as e:
                        pass

                    continue
            
            blockade_btn.click()

            try:
                div = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, 
                    self.path['inner_div'])))
                btns = div.find_elements(By.TAG_NAME, 'button')
                btns[0].click()

                print(Style.BRIGHT + '[status] ' + Fore.LIGHTGREEN_EX + 'blockade success')

                try:
                    self.blockade_log_view.new_log(f'{url_split[-1]} are blocked')
                except Exception as e:
                    pass

                blockced.append(username)
            except TimeoutException as e:
                print(Fore.RED + 'check_blockade_btn not found')

                try:
                    self.blockade_log_view.new_log(f'{url_split[-1]} error: check_blockade_btn not found')
                except Exception as e:
                    pass

                continue

            try:
                close_blockade_btn = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, 
                    self.path['close_blockade_button'])))
                close_blockade_btn.click()
            except TimeoutException as e:
                print(Fore.RED + 'close_blockade_btn not found')

                try:
                    self.blockade_log_view.new_log(f'{url_split[-1]} error: close_btn not found')
                except Exception as e:
                    pass

                continue

        self.namelist = list(set(self.namelist) - set(blockced))
            
        return 0

    def __check_login__(self):
        self.check_login = CheckLogin()
        self.check_login.run()

    def run(self) -> None:
        global is_aborted

        self.start_driver()

        self.driver.get(INSTA_LOGIN_URL)

        check_login_threading = threading.Thread(target=self.__check_login__)
        check_login_threading.start()
        check_login_threading.join()

        is_aborted = False

        blockade_thread = threading.Thread(target=self.blockade)
        blockade_thread.start()

        self.blockade_log_view = BlockadeLogView()
        self.blockade_log_view.run()

        blockade_thread.join()

        self.driver.close()

        print(Fore.YELLOW + 'Updating Data...')

        self.__update_data__()

        print(Fore.YELLOW + 'Finished Updating...')


if __name__ == "__main__":
    init(autoreset=True)

    insta_blocker = InstaBlocker()
    insta_blocker.run()
