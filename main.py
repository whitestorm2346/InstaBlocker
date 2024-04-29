from platform import platform
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

DATA_FILE = 'test.xlsx'

INSTA_LOGIN_URL = 'https://www.instagram.com/?hl=zh' 
AUTHENTICATE_URL = 'https://www.instagram.com/accounts/login/two_factor?hl=zh&next=%2F'

ENGLISH = 'Times New Roman'
CHINESE = '微軟正黑體'

class LoginForm:
    def __init__(self) -> None:
        self.__init_root__()
        self.__init_login_frame__()
        self.__init_buttons__()

    def __show_password_onclick__(self):
        if self.show_password_value.get():
            self.password_entry.config(show='')
        else:
            self.password_entry.config(show='\u25CF')

    def __start_button_onclick__(self):
        # TODO: implement the empty value checking

        print('start button onclick')
        self.root.destroy()

    def __init_root__(self):
        self.root = Tk()
        self.root.resizable(False, False)
        self.root.geometry("300x300")
        self.root.title('InstaBlocker Login')

    def __init_login_frame__(self):
        self.login_frame = LabelFrame(self.root)
        self.login_frame.config(text=' Login ', font=(ENGLISH, 12))
        self.login_frame.pack(side=TOP, fill='x', padx=10, pady=5)

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

    def __init_buttons__(self):
        self.control_frame = Frame(self.root)
        self.control_frame.pack(side=TOP, fill='x', padx=15, pady=0)

        self.start_btn = Button(self.root)
        self.start_btn.config(text='start', font=(ENGLISH, 14, 'bold'),
                              height=2, width=8, command=self.__start_button_onclick__)
        self.start_btn.pack(side=TOP, padx=10, pady=5)

    def run(self):
        self.root.mainloop()

class AuthenticateForm:
    def __init__(self) -> None:
        self.__init_root__()
        self.__init_authenticate_frame__()
        self.__init_buttons__()

    def __send_button_onclick__(self):
        # TODO: implement the empty value checking

        print('send button onclick')
        self.root.destroy()

    def __init_root__(self):
        self.root = Tk()
        self.root.resizable(False, False)
        self.root.geometry("200x200")
        self.root.title('Verify')

    def __init_authenticate_frame__(self):
        self.authenticate_frame = LabelFrame(self.root)
        self.authenticate_frame.config(text=' Authentication ', font=(ENGLISH, 12))
        self.authenticate_frame.pack(side=TOP, fill='x', padx=10, pady=5)

        self.verify_code_label = Label(self.authenticate_frame)
        self.verify_code_label.config(
            text='Verify Code', font=(ENGLISH, 14, 'bold'))
        self.verify_code_label.pack(side=TOP, pady=5)

        self.verify_code = StringVar(self.authenticate_frame)
        self.verify_code_entry = Entry(self.authenticate_frame)
        self.verify_code_entry.config(
            font=(ENGLISH, 12), textvariable=self.verify_code)
        self.verify_code_entry.pack(side=TOP, fill='x', padx=5, pady=10)

    def __init_buttons__(self):
        self.control_frame = Frame(self.root)
        self.control_frame.pack(side=TOP, fill='x', padx=15, pady=0)

        self.send_btn = Button(self.root)
        self.send_btn.config(text='send', font=(ENGLISH, 14, 'bold'),
                              height=2, width=8, command=self.__send_button_onclick__)
        self.send_btn.pack(side=TOP, padx=10, pady=5)

    def run(self):
        self.root.mainloop()

class BlockadeLogView:
    def __init__(self) -> None:
        self.__init_root__()
        self.__init_inner_frame__()

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

    def new_log(self, msg):
        log = Label(self.window_frame)
        log.config(text=msg, font=(ENGLISH, 14))
        log.pack(side=BOTTOM)

    def run(self):
        self.root.mainloop()

class InstaBlocker:
    def __init__(self) -> None:
        self.__init_data__()
        
        self.login_form = LoginForm()
        

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

    def login(self, username, password) -> int:
        self.driver.get(INSTA_LOGIN_URL)

        try:
            username_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                '//*[@id="loginForm"]/div/div[1]/div/label/input')))
        except Exception as e:
            print('User Already Logged In')
            print(format(e))
            return 1
        
        username_input.clear()
        username_input.send_keys(username)

        password_input = self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[2]/div/label/input')
        password_input.clear()
        password_input.send_keys(password)

        login_btn = self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[3]/button')
        login_btn.click()

        return 0

    def authentication(self) -> int:         
        try:
            verify_code_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                '/html/body/div[2]/div/div/div[2]/div/div/div[1]/section/main/div/div/div[1]/div[2]/form/div[1]/div/label/input')))
            verify_code_input.clear()

            self.authenticate_form = AuthenticateForm()
            self.authenticate_form.run()
            verify_code = self.authenticate_form.verify_code.get()

            verify_code_input.send_keys(verify_code)
        except Exception as e:
            if self.driver.current_url != AUTHENTICATE_URL:
                return 1
            else:
                print('verify_code_input not found')
                print(format(e))
            
                return -1

        confirm_btn = self.driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[1]/section/main/div/div/div[1]/div[2]/form/div[2]/button')
        confirm_btn.click()

        try:
            save_data_later_btn = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[1]/div[2]/section/main/div/div/div/div/div')))
            save_data_later_btn.click()
        except Exception as e:
            print('save_data_later_btn not found')
            print(format(e))

        try:
            turn_on_notifications_btn = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                '/html/body/div[3]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div[3]/button[2]')))
            turn_on_notifications_btn.click()
        except Exception as e:
            print('turn_on_notifications_btn not found')
            print(format(e))
        
        return 0

    def blockade(self, namelist) -> int:
        for username in namelist:
            self.driver.get(username)

            url_split = username.split('/')

            try:
                options_btn = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, 
                    '/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[1]/div[2]/div[2]/section/main/div/header/section/div[1]/div[3]/div')))
                options_btn.click()
            except Exception as e:
                print('options_btn not found')
                print(format(e))
                return -1

            try:
                blockade_btn = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, 
                    '/html/body/div[6]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div/button[1]')))
            except Exception as e:
                try:
                    blockade_btn = WebDriverWait(self.driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, 
                        '/html/body/div[6]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div/button[1]')))
                except Exception as e:
                    print('blockade_btn not found')
                    print(format(e))
                    self.blockade_log_view.new_log(f'{url_split[-1]} error: blockade_btn not found')
                    continue
            
            if blockade_btn.text == '解除封鎖':
                self.blockade_log_view.new_log(f'{url_split[-1]} blocked already')
                cancel_blockade_btn = self.driver.find_element(By.XPATH, '/html/body/div[6]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div/button[6]')
                cancel_blockade_btn.click()
                continue
            else:
                blockade_btn.click()
                self.blockade_log_view.new_log(f'{url_split[-1]} are blocked')

            try:
                check_blockade_btn = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, 
                    '/html/body/div[6]/div[1]/div/div[2]/div/div/div/div/div/div/div[2]/button[1]')))
                check_blockade_btn.click()
            except Exception as e:
                print('check_blockade_btn not found')
                print(format(e))
                self.blockade_log_view.new_log(f'{url_split[-1]} error: check_blockade_btn not found')
                continue

            try:
                close_blockade_btn = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, 
                    '/html/body/div[6]/div[1]/div/div[2]/div/div/div/div/div/div/div[2]/button')))
                close_blockade_btn.click()
            except Exception as e:
                print('close_blockade_btn not found')
                print(format(e))
                self.blockade_log_view.new_log(f'{url_split[-1]} error: close_btn not found')
                continue
            
        return 0

    def __run_blockade__(self):
        self.blockade(self.namelist)

    def run(self) -> None:
        self.login_form.run()
        self.start_driver()
        self.login(self.login_form.username.get(), self.login_form.password.get())
        self.authentication()

        blockade_thread = threading.Thread(target=self.__run_blockade__)
        blockade_thread.start()

        self.blockade_log_view = BlockadeLogView()
        self.blockade_log_view.run()

        self.__update_data__()


if __name__ == "__main__":
    try:
        data = pd.read_excel(DATA_FILE, header=7)
        header = pd.read_excel(DATA_FILE, nrows=7)
    except Exception as e:
        print(f'{DATA_FILE} not found')
        exit(1)

    filtered_data = data[data['評分'] <= -55]

    print(filtered_data)

    with pd.ExcelWriter(DATA_FILE, engine='openpyxl') as writer:
        header.to_excel(writer, index=False)
        filtered_data.to_excel(writer, index=False, header=False, startrow=8)

    # insta_blocker = InstaBlocker()
    # insta_blocker.run()
