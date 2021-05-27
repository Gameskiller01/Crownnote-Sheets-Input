import xpaths
from time import time
from re import findall
from html import unescape
from shutil import rmtree
from subprocess import call
from selenium import webdriver
from os import getenv, path, makedirs
from cryptography.fernet import Fernet
from shelve import open as shelve_open
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from undetected_chromedriver.v2 import Chrome as UndetectedChrome, ChromeOptions as UndetectedChromeOptions
from selenium.common.exceptions import TimeoutException, InvalidArgumentException, NoSuchElementException, StaleElementReferenceException, ElementNotInteractableException, ElementClickInterceptedException, WebDriverException

class EndCodeException(Exception):
    pass

def create_shelve_file(name, variable, value):
    with shelve_open(path.join(location, name)) as d:
        if variable not in d:
            d[variable] = value

def expand_root_element(element):
    shadow_root = driver_login.execute_script('return arguments[0].shadowRoot', element)
    return shadow_root

def raise_exception(message):
    call("cls", shell=True)
    print(message)
    raise EndCodeException

def driver_get(driver, website):
    timeout = time() + 60
    while True:
        try:
            driver.get(website)
            break
        except WebDriverException:
            call("cls", shell=True)
            print("Could not load website " + website + ". Retrying...")
        if time() > timeout:
            raise_exception("Took too long to load website " + website + ".")

def send_keys(key1, key2 = None, key3 = None):
    action = ActionChains(driver2)
    action.send_keys(key1)
    if key2 != None:
        action.send_keys(key2)
    if key3 != None:
        action.send_keys(key3)
    action.perform()

def clear_and_send_keys(path, text):
    driver2.find_element_by_xpath(path).clear()
    driver2.find_element_by_xpath(path).send_keys(text)

def hold_key(key):
    action = ActionChains(driver2)
    action.key_down(key)
    action.perform()

def release_key(key):
    action = ActionChains(driver2)
    action.key_up(key)
    action.perform()

def send_key_combo(key1, key2, key3 = None):
    action = ActionChains(driver2)
    action.key_down(key1)
    if key3 != None:
        action.key_down(key3)
    action.send_keys(key2)
    if key3 != None:
        action.key_up(key3)
    action.key_up(key1)
    action.perform()

def click_menu_item(main_menu, sub_menu, button):
    driver2.find_element_by_xpath(main_menu).click()
    action = ActionChains(driver2).move_to_element(driver2.find_element_by_xpath(sub_menu))
    action.perform()
    WebDriverWait(driver2, 5).until(EC.presence_of_element_located((By.XPATH, button))).click()

def double_click(path):
    action = ActionChains(driver2)
    action.double_click(driver2.find_element_by_xpath(path))
    action.perform()

def change_colour(colour):
    driver2.find_element_by_xpath(xpaths.sheets_text_colour_path).click()
    WebDriverWait(driver2, 5).until(EC.element_to_be_clickable((By.XPATH, colour))).click()

def get_rows_or_columns(row_or_column, message):
    timeout = time() + 5
    while True:
        try:
            if row_or_column == "column":
                try:
                    rows_or_columns = ord(driver2.find_element_by_xpath(xpaths.find_column_path_2).get_attribute("aria-label").split("Column ")[1].split(" ")[0]) - 64
                except NoSuchElementException:
                    rows_or_columns = ord(driver2.find_element_by_xpath(xpaths.find_column_path).get_attribute("aria-label").split("Delete column ")[1].split(" ")[0]) - 64
            elif row_or_column == "row":
                try:
                    rows_or_columns = int(driver2.find_element_by_xpath(xpaths.find_row_path_2).get_attribute("aria-label").split("Row ")[1].split(" ")[0])
                except NoSuchElementException:
                    rows_or_columns = int(driver2.find_element_by_xpath(xpaths.find_row_path).get_attribute("aria-label").split("Delete row ")[1].split(" ")[0])
            break
        except StaleElementReferenceException:
            pass
        if time() > timeout:
            raise_exception(message)
    return rows_or_columns

def go_to_cell(cell):
    toggle_keyboard_shortcuts("true")
    send_keys(Keys.F5)
    driver2.execute_script("arguments[0].setAttribute('autocomplete', 'off')", WebDriverWait(driver2, 7).until(EC.visibility_of_element_located((By.XPATH, xpaths.go_to_cell_input_path))));
    clear_and_send_keys(xpaths.go_to_cell_input_path, (cell, Keys.ENTER))

def columns_needed(length_of_charts):
    lengths = [1, 3, 5, 10, 20, 40, 75, 100]
    i = 0
    columns_needed = None
    if length_of_charts == 1:
        columns_needed = 5
    while columns_needed == None:
        if i == 7:
            columns_needed = i + 6
        elif length_of_charts > lengths[i] and length_of_charts <= lengths[i+1]:
            columns_needed = i + 6
        i = i + 1
    return columns_needed

def toggle_keyboard_shortcuts(true_or_false):
    send_key_combo(Keys.CONTROL, '/')
    WebDriverWait(driver2, 10).until(EC.element_to_be_clickable((By.XPATH, xpaths.toggle_shortcuts_path)))
    if driver2.find_element_by_xpath(xpaths.toggle_shortcuts_path).get_attribute("aria-pressed") != true_or_false:
        driver2.find_element_by_xpath(xpaths.toggle_shortcuts_path).click()
    send_keys(Keys.ESCAPE)

def colour_key(message, number, path):
    send_keys(Keys.ENTER)
    clear_and_send_keys(xpaths.edit_active_cell_path, (message, Keys.HOME, Keys.SHIFT + Keys.ARROW_RIGHT * number))
    change_colour(path)
    driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.ENTER)

location = path.join(getenv("APPDATA"), "Crownnote")
if not path.exists(location):
    makedirs(location)
create_shelve_file("variables", "username", None)
create_shelve_file("variables", "length_of_charts", None)
create_shelve_file("variables", "spreadsheet_link", None)

with shelve_open(path.join(location, "variables")) as d:
    if d["username"] == None:
        d["username"] = input("Enter your CrownNote username: ")
        call("cls", shell=True)
    if d["length_of_charts"] == None:
        while True:
            try:
                d["length_of_charts"] = int(input("Enter how many songs you have in each of your weekly CrownNote charts: "))
                call("cls", shell=True)
                if d["length_of_charts"] < 1:
                    print("The length of your CrownNote charts must be greater than 0. Please try again.")
                    continue
                break
            except ValueError:
                call("cls", shell=True)
                print("The length of your CrownNote charts must be an integer. Please try again.")
    if d["spreadsheet_link"] == None:
        d["spreadsheet_link"] = input("""Enter the share link to your spreadsheet (making sure that the mode is set to "Anyone on the Internet with this link can edit"): """).split("?")[0]
        call("cls", shell=True)
    username = d["username"]
    length_of_charts = d["length_of_charts"]
    spreadsheet_link = d["spreadsheet_link"]
personal_link_path = """//*[@id="block-views-32f0a13b6ed96df9665df3ef2668965a"]/div/div/div/div[@class='view-content']//a[contains(text(),'""" + username + """')]/parent::*/parent::*//a[@class='btn btn--plain']"""

with open('main.key', 'rb') as key:
    with open('.enc', 'rb') as enc_file:
        decrypted = str(Fernet(key.read()).decrypt(enc_file.read()))[2:][:-1].split("\\r\\n")

lengths = [3, 5, 10, 20, 40, 75, 100]

create_shelve_file("variables", "login", False)

try:
    
    headless=Options()
    headless.headless=True
    try:
        webdriver.Chrome(ChromeDriverManager().install(), options=headless).quit()
    except ValueError:
        raise_exception("It seems like Google Chrome is not installed. Please install it and try again.")
    call("cls", shell=True)
    
    with shelve_open(path.join(location, "variables")) as d:
        login = d["login"]
    if login == False:
        uc_options = UndetectedChromeOptions()
        uc_options.add_argument("--start-maximized")
        uc_options.add_argument("--user-data-dir=" + getenv("LOCALAPPDATA") + "\\Google\\Chrome\\Selenium User Data")
        
        driver_login = UndetectedChrome(options=uc_options)
        
        try:
            
            driver_get(driver_login, "chrome://settings/passwords")
            timeout = time() + 5
            while time() < timeout:
                try:
                    expand_root_element(expand_root_element(expand_root_element(expand_root_element(expand_root_element(expand_root_element(driver_login.find_element_by_tag_name("settings-ui")).find_element_by_tag_name("settings-main")).find_element_by_tag_name("settings-basic-page")).find_element_by_tag_name("settings-autofill-page")).find_element_by_tag_name("passwords-section")).find_element_by_tag_name("settings-toggle-button")).find_element_by_tag_name("cr-toggle").click()
                    break
                except ElementNotInteractableException:
                    pass
            driver_get(driver_login, "https://accounts.google.com/signin/v2/identifier")
            WebDriverWait(driver_login, 30).until(EC.presence_of_element_located((By.XPATH, xpaths.google_username_path))).send_keys(decrypted[2], Keys.ENTER)
            WebDriverWait(driver_login, 30).until(EC.visibility_of_element_located((By.XPATH, xpaths.google_password_path))).send_keys(decrypted[3], Keys.ENTER)
            WebDriverWait(driver_login, 30).until(lambda driver_login: driver_login.find_elements(By.XPATH, xpaths.google_phone_number_input_path) or driver_login.find_elements(By.XPATH, xpaths.google_confirm_phone_number_path) or driver_login.find_elements(By.XPATH, xpaths.google_text_path) or driver_login.find_elements(By.XPATH, xpaths.google_home_path))
            timeout = time() + 5
            while time() < timeout:
                try:
                    try:
                        driver_login.find_element_by_xpath(xpaths.google_phone_number_input_path).send_keys(decrypted[4], Keys.ENTER)
                    except NoSuchElementException:
                        try:
                            try:
                                driver_login.find_element_by_xpath(xpaths.google_confirm_phone_number_path).click()
                            except ElementClickInterceptedException:
                                WebDriverWait(driver_login, 10).until(EC.element_to_be_clickable((By.XPATH, xpaths.google_confirm_phone_number_path))).click()
                            WebDriverWait(driver_login, 30).until(EC.visibility_of_element_located((By.XPATH, xpaths.google_phone_number_input_path))).send_keys(decrypted[4], Keys.ENTER)
                        except NoSuchElementException:
                            try:
                                try:
                                    driver_login.find_element_by_xpath(xpaths.google_text_path).click()
                                except ElementClickInterceptedException:
                                    WebDriverWait(driver_login, 10).until(EC.element_to_be_clickable((By.XPATH, xpaths.google_text_path))).click()
                                WebDriverWait(driver_login, 30).until(lambda driver_login: driver_login.find_elements(By.XPATH, xpaths.google_too_many_attempts_path) or driver_login.find_elements(By.XPATH, xpaths.google_verification_input_path))
                                try:
                                    driver_login.find_element_by_xpath(xpaths.google_verification_input_path)
                                    driver_verification = webdriver.Chrome(ChromeDriverManager().install(), options=headless)
                                    call("cls", shell=True)
                                    try:
                                        driver_get(driver_verification, decrypted[5])
                                        WebDriverWait(driver_verification, 45).until(EC.visibility_of_element_located((By.XPATH, xpaths.recovery_phone_messages_path)))
                                        timeout2 = time() + 30
                                        while time() < timeout2:
                                            try:
                                                WebDriverWait(driver_verification, 30).until(EC.visibility_of_element_located((By.XPATH, xpaths.recovery_phone_messages_path)))
                                                driver_login.find_element_by_xpath(xpaths.google_verification_input_path).send_keys(driver_verification.find_element_by_xpath(xpaths.recovery_phone_google_message_path).get_attribute("innerHTML"), Keys.ENTER)
                                                break
                                            except NoSuchElementException:
                                                driver_verification.find_element_by_xpath(xpaths.recovery_phone_refresh_path).click()
                                    finally:
                                        driver_verification.close()
                                        driver_verification.quit()
                                except NoSuchElementException:
                                    pass
                            except NoSuchElementException:
                                pass
                    break
                except ElementNotInteractableException:
                    pass
            WebDriverWait(driver_login, 30).until(EC.presence_of_element_located((By.XPATH, xpaths.google_home_path)))
            with shelve_open(path.join(location, "variables")) as d:
                d["login"] = True
        
        except TimeoutException:
            pass
        finally:
            driver_login.close()
            driver_login.quit()
            with shelve_open(path.join(location, "variables")) as d:
                login = d["login"]
            if login == False:
                try:
                    rmtree(path.join(getenv("LOCALAPPDATA") + "\\Google\\Chrome", "Selenium User Data"))
                except FileNotFoundError:
                    pass
    
    options1 = Options()
    options1.add_argument("--start-maximized")
    options1.add_argument('--hide-scrollbars')
    options1.add_argument("--log-level=3")
    options1.headless=True
    caps = options1.to_capabilities()
    caps["pageLoadStrategy"] = "eager"
    
    options2 = Options()
    options2.add_argument("--window-size=1920,1080")
    options2.add_argument('--hide-scrollbars')
    options2.add_argument("--log-level=3")
    options2.add_argument("--user-data-dir=" + getenv("LOCALAPPDATA") + "\\Google\\Chrome\\Selenium User Data")
    options2.headless=True
    
    driver1 = webdriver.Chrome(ChromeDriverManager().install(), desired_capabilities=caps)
    driver2 = webdriver.Chrome(ChromeDriverManager().install(), options=options2)
    call("cls", shell=True)
    
    try:
        
        driver_get(driver1, "https://crownnote.com/user/login")
        WebDriverWait(driver1, 15).until(EC.presence_of_element_located((By.XPATH, xpaths.username_path)))
        driver1.find_element_by_xpath(xpaths.username_path).send_keys(decrypted[0])
        driver1.find_element_by_xpath(xpaths.password_path).send_keys(decrypted[1], Keys.ENTER)
        WebDriverWait(driver1, 30).until(EC.presence_of_element_located((By.XPATH, xpaths.user_menu_path)))
        driver_get(driver1, "https://crownnote.com/users/" + username.replace(" ", "-"))
        
        try:
            driver_get(driver2, spreadsheet_link)
        except InvalidArgumentException:
            with shelve_open(path.join(location, "variables")) as d:
                d["spreadsheet_link"] = None
            raise_exception("The spreadsheet link provided is invalid. Please verify the link and try again.")
        try:
            WebDriverWait(driver1, 45).until(EC.presence_of_element_located((By.XPATH, xpaths.button_path)))
        except TimeoutException:
            with shelve_open(path.join(location, "variables")) as d:
                d["username"] = None
            raise_exception("A CrownNote user with the username provided does not seem to exist, or they seem to have no charts. Please verify the username and try again.")
        try:
            WebDriverWait(driver2, 60).until(EC.presence_of_element_located((By.XPATH, xpaths.spreadsheet_loaded_path)))
        except TimeoutException:
            with shelve_open(path.join(location, "variables")) as d:
                d["spreadsheet_link"] = None
            raise_exception("The link provided does not seem to lead to a Google Spreadsheet. Please verify the link and try again.")
        try:
            send_keys(Keys.ENTER)
            driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.ENTER)
            send_keys(Keys.ARROW_UP)
        except NoSuchElementException:
            with shelve_open(path.join(location, "variables")) as d:
                d["spreadsheet_link"] = None
            raise_exception("""The spreadsheet link provided is not editable. Please set it to "Anyone on the Internet with this link can edit" and try again.""")
        
        driver2.find_element_by_xpath(xpaths.sheets_tools_button_path).click()
        if driver2.find_element_by_xpath(xpaths.sheets_disable_autocomplete_path).get_attribute("aria-checked") == "true":
            try:
                driver2.find_element_by_xpath(xpaths.sheets_disable_autocomplete_path).click()
            except ElementNotInteractableException:
                send_keys(Keys.ESCAPE)
                click_menu_item(xpaths.sheets_tools_button_path, xpaths.sheets_autocomplete_menu_path, xpaths.sheets_disable_autocomplete_path)
        else:
            send_keys(Keys.ESCAPE)
        
        send_key_combo(Keys.CONTROL, Keys.END)
        rows = get_rows_or_columns("row", "Could not get amount of rows in spreadsheet.")
        columns = get_rows_or_columns("column", "Could not get amount of columns in spreadsheet.")
        while rows < 3:
            send_key_combo(Keys.ALT, 'i')
            send_keys('b')
            rows = get_rows_or_columns("row", "Could not get amount of rows in spreadsheet.")
        try:
            driver2.find_element_by_xpath(xpaths.sheets_understood_path).click()
        except NoSuchElementException:
            pass
        while columns < columns_needed(length_of_charts):
            send_key_combo(Keys.ALT, 'i')
            send_keys('o')
            columns = get_rows_or_columns("column", "Could not get amount of columns in spreadsheet.")
        send_key_combo(Keys.CONTROL, Keys.HOME)

        create_shelve_file("variables", spreadsheet_link, False)
        with shelve_open(path.join(location, "variables")) as d:
            setup = d[spreadsheet_link]
        if setup == False:
            send_keys("Song Artist(s)", Keys.ARROW_RIGHT)
            send_keys("Song Name", Keys.ARROW_RIGHT)
            send_keys("Weeks at #1")
            i = 0
            while length_of_charts > lengths[i]:
                send_keys(Keys.ARROW_RIGHT, "Weeks in Top " + str(lengths[i]))
                i = i + 1
            if length_of_charts > 1:
                send_keys(Keys.ARROW_RIGHT, "Weeks in Top " + str(length_of_charts))
            send_keys(Keys.ARROW_DOWN, Keys.ARROW_UP)
            
            if get_rows_or_columns("column", "Could not get position of selection in spreadsheet.") != 1:
                hold_key(Keys.SHIFT)
                send_keys(Keys.ARROW_LEFT)
                while driver2.find_element_by_xpath(xpaths.sheets_selection_path).get_attribute("style").split("left: ")[1].split("px;")[0] != "0":
                    send_keys(Keys.ARROW_LEFT)
                release_key(Keys.SHIFT)
            if driver2.find_element_by_xpath(xpaths.sheets_bold_path).get_attribute("aria-pressed") == "false":
                send_key_combo(Keys.CONTROL, 'b')
            
            send_key_combo(Keys.CONTROL, Keys.SPACE)
            try:
                click_menu_item(xpaths.sheets_format_button_path, xpaths.sheets_text_wrapping_menu_path_2, xpaths.text_wrapping_clip_path)
            except NoSuchElementException:
                send_keys(Keys.ESCAPE)
                click_menu_item(xpaths.sheets_format_button_path, xpaths.sheets_text_wrapping_menu_path, xpaths.text_wrapping_clip_path)
            send_keys(Keys.ARROW_RIGHT)
            
            go_to_cell(chr(columns_needed(length_of_charts) + 64) + "2")
            colour_key("BLUE = Consecutive", 4, xpaths.sheets_blue_path)
            colour_key("RED = Non-Consecutive", 3, xpaths.sheets_red_path)
            
            send_key_combo(Keys.CONTROL, Keys.HOME)
            
            send_key_combo(Keys.SHIFT, Keys.ARROW_RIGHT)
            send_key_combo(Keys.CONTROL, Keys.SPACE)
            click_menu_item(xpaths.sheets_format_button_path, xpaths.sheets_number_menu_path, xpaths.format_plain_text_path)
            
            send_keys(Keys.ARROW_RIGHT)
            toggle_keyboard_shortcuts("false")
            while get_rows_or_columns("column", "Could not get position of selection in spreadsheet.") + 1 <= columns_needed(length_of_charts) - 2:
                send_key_combo(Keys.ALT, 'd')
                while driver2.find_element_by_xpath(xpaths.sheets_filter_views_menu_path).get_attribute("class") != "goog-menuitem apps-menuitem goog-submenu goog-menuitem-highlight":
                    send_keys(Keys.ARROW_DOWN)
                send_keys(Keys.ARROW_RIGHT, Keys.ENTER, Keys.ARROW_RIGHT)
                WebDriverWait(driver2, 10).until(EC.presence_of_element_located((By.XPATH, xpaths.filter_view_name_path))).send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
                driver2.find_element_by_xpath(xpaths.filter_view_name_path).send_keys(driver2.find_element_by_xpath(xpaths.formula_bar_path).get_attribute("innerHTML").split("<br>")[0], Keys.ENTER)
                try:
                    driver2.find_element_by_xpath(xpaths.filter_view_popup_dismiss_path).click()
                except NoSuchElementException:
                    pass
                driver2.find_element_by_xpath(xpaths.filter_view_input_path).send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
                driver2.find_element_by_xpath(xpaths.filter_view_input_path).send_keys("A1:" + chr(columns_needed(length_of_charts) + 62) + str(rows), Keys.ENTER)
                send_key_combo(Keys.CONTROL, 'r', Keys.ALT)
                WebDriverWait(driver2, 5).until(EC.presence_of_element_located((By.XPATH, xpaths.sort_column_path))).click()
                driver2.find_element_by_xpath(xpaths.filter_view_close_path).click()
            
            double_click(xpaths.sheets_name_path)
            WebDriverWait(driver2, 5).until(EC.presence_of_element_located((By.XPATH, xpaths.sheets_rename_input_path))).send_keys("Singles", Keys.ENTER)
            
            click_menu_item(xpaths.sheets_view_button_path, xpaths.sheets_freeze_menu_path, xpaths.freeze_one_row_path)
            click_menu_item(xpaths.sheets_view_button_path, xpaths.sheets_freeze_menu_path, xpaths.freeze_two_columns_path)
            
            with shelve_open(path.join(location, "variables")) as d:
                d[spreadsheet_link] = True
        
        buttons = driver1.find_elements_by_class_name("toggle-button")
        
        links = []
        for i, button in enumerate(buttons):
            list_path = """//*[@id="block-views-users_charts-user_chart_list"]/div/div/div/div[2]/div[""" + str((i * 2) + 2) + """]//div[@class='history-item chart-type--weekly-- one-chart-per-week']//a[@href]"""
            places = driver1.find_elements_by_xpath(list_path)
            for place in places:
                links.append([place.get_attribute("href"), place.get_attribute('innerHTML').split("</span>")[0].split(">")[-1]])

        create_shelve_file("variables", "first run", False)
        with shelve_open(path.join(location, "songs"))as d:
            for item in d:
                if d[item] == False:
                    del d[item]
            with shelve_open(path.join(location, "variables")) as b:
                if b["first run"] == False:
                    go_to_cell('A' + str(len(d.keys()) + 2))
        update_dict_created = False
        for link in links:
            song_links = []
            create_shelve_file("charts", link[1], False)
            with shelve_open(path.join(location, "charts")) as d:
                chart = d[link[1]]
            if chart == False:
                driver_get(driver1, link[0])
                WebDriverWait(driver1, 90).until(EC.presence_of_element_located((By.XPATH, xpaths.final_weeks_path)))
                call("cls", shell=True)
                print("Loaded chart dated " + link[1])
                
                with shelve_open(path.join(location, "variables")) as d:
                    first_run = d["first run"]
                if first_run == False:
                    song_places = driver1.find_elements_by_xpath(xpaths.song_debut_path)
                elif first_run == True:
                    song_places = driver1.find_elements_by_xpath(xpaths.song_link_path)
                    if update_dict_created == False:
                        song_complete_on_update = {}
                        update_dict_created = True
                for song_place in song_places:
                    song_links.append(song_place.get_attribute("href"))
                
                for song_link in song_links:
                    artist_links = []
                    driver_get(driver1, song_link)
                    WebDriverWait(driver1, 30).until(EC.presence_of_element_located((By.XPATH, xpaths.artists_path)))
                    artists = driver1.find_elements_by_xpath(xpaths.artists_path + "/a[@href]")
                    for artist in artists:
                        artist_links.append(artist.get_attribute("href"))
                    artist_array = [[None for i in range(2)] for j in range(len(artist_links))]
                    artists_full = unescape(" ".join(findall(r'>.+?<', driver1.find_element_by_xpath(xpaths.artists_path).get_attribute('innerHTML'))).replace("<", "").replace(">", "").replace(" , ", ",").replace("   ft.  ", "ft."))
                    song_name = unescape(driver1.find_element_by_xpath(xpaths.song_name_path).get_attribute('innerHTML'))
                    WebDriverWait(driver1, 60).until(EC.presence_of_element_located((By.XPATH, xpaths.user_list_path)))
                    personal_link = driver1.find_element_by_xpath(personal_link_path).get_attribute("href")
                    create_shelve_file("songs", personal_link, False)
                    if first_run == True and personal_link not in song_complete_on_update:
                        song_complete_on_update[personal_link] = False
                    
                    with shelve_open(path.join(location, "songs")) as d:
                        song_complete = d[personal_link]
                    if (first_run == False and song_complete == False) or (first_run == True and song_complete_on_update[personal_link] == False):
                        with shelve_open(path.join(location, "songs"))as d:
                            if first_run == True and song_complete == True:
                                go_to_cell('A' + str(list(d).index(personal_link) + 2))
                            elif first_run == True and song_complete == False:
                                go_to_cell('A' + str(len(d.keys()) + 1))
                        driver_get(driver1, personal_link)
                        i = 0
                        weeks_in_range = {"Weeks at #1": [0, "Consecutive"]}
                        while length_of_charts > lengths[i]:
                            weeks_in_range["Weeks in Top " + str(lengths[i])] = [0, "Consecutive"]
                            i = i + 1
                        if length_of_charts > 1:
                            weeks_in_range["Weeks in Top " + str(length_of_charts)] = [0, "Consecutive"]
                        WebDriverWait(driver1, 60).until(EC.presence_of_element_located((By.XPATH, xpaths.song_charts_list_path)))
                        current_row = get_rows_or_columns("row", "Could not get position of selection in spreadsheet.")
                        song_charts = driver1.find_elements_by_xpath(xpaths.song_charts_path)
                        song_chart_runs = [[None for i in range(2)] for j in range(len(song_charts))]
                        for i, song_chart in enumerate(song_charts):
                            j = 0
                            song_chart_runs[i][0] = int(song_chart.get_attribute("innerHTML").split("""<span class="history-position">""")[1].split("</span>")[0])
                            song_chart_runs[i][1] = datetime.strptime(song_chart.get_attribute("innerHTML").split("</span></a></span>")[0].split(">")[-1], '%B %d, %Y')
                            if song_chart_runs[i][0] == 1:
                                weeks_in_range["Weeks at #1"][0] = weeks_in_range["Weeks at #1"][0] + 1
                                if i > 0 and weeks_in_range["Weeks at #1"][0] > 1 and (song_chart_runs[i-1][0] != 1 or song_chart_runs[i-1][1] + timedelta(days=7) != song_chart_runs[i][1]):
                                    weeks_in_range["Weeks at #1"][1] = "Non-Consecutive"
                            while length_of_charts > lengths[j]:
                                if song_chart_runs[i][0] <= length_of_charts and song_chart_runs[i][0] <= lengths[j]:
                                    weeks_in_range["Weeks in Top " + str(lengths[j])][0] = weeks_in_range["Weeks in Top " + str(lengths[j])][0] + 1
                                    if i > 0 and weeks_in_range["Weeks in Top " + str(lengths[j])][0] > 1 and (song_chart_runs[i-1][0] > lengths[j] or song_chart_runs[i-1][1] + timedelta(days=7) != song_chart_runs[i][1]):
                                        weeks_in_range["Weeks in Top " + str(lengths[j])][1] = "Non-Consecutive"
                                j = j + 1
                            if song_chart_runs[i][0] <= length_of_charts and length_of_charts != 1:
                                weeks_in_range["Weeks in Top " + str(length_of_charts)][0] = weeks_in_range["Weeks in Top " + str(length_of_charts)][0] + 1
                                if i > 0 and weeks_in_range["Weeks in Top " + str(length_of_charts)][0] > 1 and (song_chart_runs[i-1][0] > length_of_charts or song_chart_runs[i-1][1] + timedelta(days=7) != song_chart_runs[i][1]):
                                    weeks_in_range["Weeks in Top " + str(length_of_charts)][1] = "Non-Consecutive"
                        if song_complete == False:
                            send_keys(Keys.ENTER)
                            clear_and_send_keys(xpaths.edit_active_cell_path, artists_full)
                            places_moved = 0
                            for i, artist_link in enumerate(artist_links):
                                driver_get(driver1, artist_link)
                                WebDriverWait(driver1, 60).until(EC.presence_of_element_located((By.XPATH, xpaths.edit_path)))
                                artist_array[i][0] = driver1.find_element_by_xpath(xpaths.edit_path).get_attribute("href")[:-5]
                                artist_array[i][1] = unescape(driver1.find_element_by_xpath(xpaths.artist_name_path).get_attribute("innerHTML"))
                                timeout = time() + 10
                                while driver2.execute_script("return window.getSelection().toString();") != artist_array[i][1]:
                                    driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.HOME, Keys.ARROW_RIGHT * places_moved)
                                    driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.SHIFT + Keys.ARROW_RIGHT * len(artist_array[i][1]))
                                    if time() > timeout:
                                        raise_exception("Couldn't link artist to artist page.")
                                places_moved = places_moved + len(artist_array[i][1])
                                driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.CONTROL + 'k')
                                WebDriverWait(driver2, 5).until(EC.presence_of_element_located((By.XPATH, xpaths.input_url_path)))
                                driver2.find_element_by_xpath(xpaths.input_url_path).send_keys(artist_array[i][0], Keys.ENTER)
                                if artists_full.split(artist_array[i][1])[1][:2] == ", ":
                                    driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.ARROW_RIGHT * 2)
                                    places_moved = places_moved + 2
                                elif artists_full.split(artist_array[i][1])[1][:5] == " ft. ":
                                    driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.ARROW_RIGHT * 5)
                                    places_moved = places_moved + 5
                            driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.ENTER)
                            
                            send_keys(Keys.ARROW_RIGHT, Keys.ARROW_UP, Keys.ENTER)
                            clear_and_send_keys(xpaths.edit_active_cell_path, (song_name, Keys.CONTROL + ('a' + 'k')))
                            WebDriverWait(driver2, 5).until(EC.presence_of_element_located((By.XPATH, xpaths.input_url_path)))
                            driver2.find_element_by_xpath(xpaths.input_url_path).send_keys(personal_link, Keys.ENTER * 2)
                            
                        gone_to_cell = False
                        for i, _ in enumerate(weeks_in_range):
                            if list(weeks_in_range.values())[i][0] > 0:
                                if gone_to_cell == False:
                                    go_to_cell(chr(i + 67) + str(current_row))
                                    gone_to_cell = True
                                else:
                                    send_keys(Keys.ARROW_RIGHT)
                                send_keys(list(weeks_in_range.values())[i][0])
                            if list(weeks_in_range.values())[i][0] > 1:
                                driver2.find_element_by_xpath(xpaths.edit_active_cell_path).send_keys(Keys.CONTROL + 'a')
                                if list(weeks_in_range.values())[i][1] == "Consecutive":
                                    change_colour(xpaths.sheets_blue_path)
                                elif list(weeks_in_range.values())[i][1] == "Non-Consecutive":
                                    change_colour(xpaths.sheets_red_path)
                        send_keys(Keys.ARROW_DOWN)
                        if first_run == False:
                            go_to_cell("A" + str(current_row + 1))
                        if current_row + 1 == rows:
                            driver2.find_element_by_xpath(xpaths.add_row_button_path).click()
                            rows = rows + int(driver2.find_element_by_xpath(xpaths.add_row_input_path).get_attribute("value"))
                            toggle_keyboard_shortcuts("false")
                            filter_views = driver2.find_elements_by_xpath(xpaths.filter_views_path)
                            filter_view_names = []
                            for filter_view in filter_views:
                                filter_view_names.append(filter_view.get_attribute("innerHTML").split("""<div class="goog-menuitem-checkbox"></div>""")[1])
                            for filter_view_name in filter_view_names:
                                send_key_combo(Keys.ALT, 'd')
                                while driver2.find_element_by_xpath(xpaths.sheets_filter_views_menu_path).get_attribute("class") != "goog-menuitem apps-menuitem goog-submenu goog-menuitem-highlight":
                                    send_keys(Keys.ARROW_DOWN)
                                send_keys(Keys.ARROW_RIGHT)
                                while driver2.find_element_by_xpath(xpaths.filter_views_path + """[contains(text(),'""" + filter_view_name + """')]/parent::*""").get_attribute("class") != "goog-menuitem apps-menuitem goog-option goog-menuitem-highlight":
                                    send_keys(Keys.ARROW_DOWN)
                                send_keys(Keys.ENTER)
                                WebDriverWait(driver2, 10).until(EC.presence_of_element_located((By.XPATH, xpaths.filter_view_input_path))).send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
                                driver2.find_element_by_xpath(xpaths.filter_view_input_path).send_keys("A1:" + chr(columns_needed(length_of_charts) + 62) + str(rows), Keys.ENTER)
                                try:
                                    driver2.find_element_by_xpath(xpaths.filter_view_popup_dismiss_path).click()
                                except NoSuchElementException:
                                    pass
                                try:
                                    driver2.find_element_by_xpath(xpaths.sheets_understood_path).click()
                                except NoSuchElementException:
                                    pass
                                driver2.find_element_by_xpath(xpaths.filter_view_close_path).click()
                        if song_complete == False:
                            with shelve_open(path.join(location, "songs")) as d:
                                d[personal_link] = True
                        if first_run == True:
                            song_complete_on_update[personal_link] = True
                with shelve_open(path.join(location, "charts")) as d:
                    d[link[1]] = True
        with shelve_open(path.join(location, "variables")) as d:
            if d["first run"] == False:
                d["first run"] = True
        print("Finished")
    
    except TimeoutException:
        call("cls", shell=True)
        print("Encountered a timeout error.")
    finally:
        input("Press Enter to exit.")
        driver1.close()
        driver1.quit()
        driver2.close()
        driver2.quit()

except EndCodeException:
    pass
