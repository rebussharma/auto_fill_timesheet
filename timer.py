import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import date, datetime
import os


def get_today_weekday_name():
    """
        Get name of the day: Monday, ... , Sunday
    :return: name of the day
    """
    return datetime.today().strftime('%A')[:3].upper()


def get_today_date():
    """
        Get today's full time
    :return: current date: in DD-MM-YYYY format
    """
    return date.today()


def get_today_day():
    """
        Get todays numeric date
    :return: date: 1, ..., 28/30/31
    """
    today = get_today_date()
    return today.day


def get_today_month():
    """
        Get current month in number
    :return: current month: 1, ..., 12
    """
    today = get_today_date()
    return today.month


def get_today_year():
    """
        Get current year in number
    :return: current year: 2022, ...
    """
    today = get_today_date()
    return today.year


def check_if_weekend():
    """
        Check if today is weekend. Weekend are 5th and 6th day of week.
        6th day being the last day, First fay is day 0
    :return: current year: 2022, ...
    """
    weekend_number = date.today().weekday()
    if weekend_number == 5 or weekend_number == 6:
        return True
    return False


def national_holiday_list():
    """
        A list of all national holidays
    :return: dict of holidays
    """
    holidays = {
        1: (1, 17),
        2: 21,
        5: 30,
        6: 19,
        7: 4,
        9: 5,
        10: 10,
        11: (11, 24),
        12: (25, 26)
    }
    return holidays


def is_leap_year(year):
    """
        Check if this year is leap year to make sure we count Feb 29
    :param year: current year in number
    :return: True if leap year else false
    """
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


def last_days_of_months():
    """
        Find how many days are in given month.
        Helpful to see when pay-period ends for given month
    :return: dict of total days in months
    """
    last_days = {
        1: 31,
        2: 29,
        3: 31,
        4: 30,
        5: 31,
        6: 30,
        7: 31,
        8: 31,
        9: 30,
        10: 31,
        11: 30,
        12: 31
    }
    return last_days


def get_last_day():
    """
        Get the last day for given month.
        Last day is needed to see if pay-period ends for this month this day
    :return: return last day number for current month
    """
    if is_leap_year(get_today_year()) and get_today_month() == 2:
        return 29
    return last_days_of_months().get(get_today_month())


def end_period():
    """
        Check if today is the end-period of timesheet. End periods are bi-weekly
            If yes, timesheet needs to be submitted!
    :return: True if today is end-period else false
    """
    if get_today_day() == 15 or get_today_day() == get_last_day():
        return True
    return False


def begin_period():
    """
        check if today is the day to start a new time sheet
    :return: True if today is 1st day or 16th day of month
    """
    if get_today_day() == 1 or get_today_day() == 16:
        return True
    return False


def check_if_national_holiday(today):
    """
        TODO: Feature to add a differnt chanrge for timesheet is cumbersome. So, its not implemented yet!
        check if today is holiday. If holiday fill in timesheet for Holiday charge code
    :param today: current day in number
    :return: True if today is holiday else false
    """
    holidays = national_holiday_list()
    holidays_this_month = holidays.get(get_today_month())

    if isinstance(holidays_this_month, tuple):
        if today in holidays_this_month:
            return True
        return False
    else:
        if today == holidays_this_month:
            return True
        return False


# Not implemented as this closes all instance
def close_all_chrome():
    """
        NOTE: This is not implemented as it closes all chrome instances
        Close ALL instances of chrome
    :return:
    """
    print("closing all instances of chrome")
    os.system("taskkill /im chrome.exe /f")


def open_chrome():
    """
        Opens a new instance of Google Chrome and navigates to timesheet url
    :return: chrome driver instance
    """
    print("opening google chrome")
    options = Options()
    options.add_argument("start-maximized")
    options.add_argument("--user-data-dir=C:\\Users\\RibashSharma\\AppData\\Local\\Google\\Chrome\\User Data")
    options.page_load_strategy = 'normal'
    driver = webdriver.Chrome(service=Service('C:\\Users\\RibashSharma\\.wdm\\drivers\\chromedriver\\win32\\99.0.4844'
                                              '.51\\chromedriver.exe'), options=options)
    print("navigating to timesheet")
    driver.get("https://aptive.unanet.biz/aptive/action/home")
    time.sleep(5)
    return driver


# if today is new period, create new timesheet
def click_create_timesheet_button(driver):
    print("Adding New Time Sheet")
    css_selector = "img[src='/aptive/images/add.png']"
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector))).click()
    time.sleep(3)


def save_new_timesheet(driver):
    button_save = driver.find_element(By.ID, "button_save")
    button_save.click()
    time.sleep(1)


# def add_timesheet_if_not_active():
# TODO needs to be done later

def click_edit_button(driver):
    """
        Clicks on the edit icon button to edit the timesheet
    :param driver: driver instance of opened chrome
    :return:
    """
    print("clicking edit button")
    css_selector = "img[src='/aptive/images/edit.png']"
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector))).click()


def find_add_time_field(driver):
    """
        Find the field that needs to be edited to add time
    :param driver: driver instance of opened chrome
    :return:
    """
    '''
        find all the spans with class type dow or dom
           and returns its parent element w/ class hours weekly
    '''
    e = (driver.find_elements(by=By.XPATH, value="//span[@class='dow' or @class='dom']/parent::th["
                                                 "@class='hours-weekday']"))

    '''
    Objective is to find the 'id' of the element that has PROJECT CODE
    Once that id is found, the id has row number in it. We get the row number and edit the hour field in the row
    '''
    #
    row_to_edit = driver.find_element(by=By.XPATH, value="//input[contains(@value, 'C_IBM 1034_S_IBM_')]")
    row_to_edit_id = row_to_edit.get_attribute('id')
    print(row_to_edit_id)
    row_number = row_to_edit_id[-1]

    for i in e:
        weekday = i.text[:3]
        day = int(i.text[3:])

        if get_today_weekday_name() == weekday and get_today_day() == day:
            parent_id = i.get_attribute("id")
            box_number = parent_id[parent_id.rindex('_') + 1:]
            input_field_id = 'd_r' + row_number + '_' + box_number
            return input_field_id

    return -1


def add_time(driver, input_field_id):
    """
        Add 8 hours of time to the field for today
    :param driver: chrome driver
    :param input_field_id: field that needs to be edited
    :return:
    """
    print("adding 8 hours")
    hours = driver.find_element(By.ID, input_field_id)
    hours.clear()
    hours.send_keys('8')
    time.sleep(1)


def save(driver):
    """
        Save the timesheet by clicking save button
    :param driver: driver instance of opened chrome
    :return:
    """
    print("saving the timesheet")
    button_save = driver.find_element(By.ID, "button_save")
    button_save.click()
    time.sleep(1)


def submit(driver):
    """
        submit timesheet by clicking submit button
    :param driver: chrome driver
    :return:
    """
    print("submitting timesheet")
    button_submit = driver.find_element(By.ID, "button_submit")
    button_submit.click()
    time.sleep(1)


def final_submit(driver):
    """
        submit timesheet again in new page by clicking submit button
    :param driver: chrome driver
    :return:
    """
    print("submitting timesheet second window")
    button_submit = driver.find_element(By.ID, "ts_submit")
    button_submit.click()
    time.sleep(1)


if __name__ == '__main__':

    input("All open chrome windows need to be manually closed to proceed.\n"
          "Press any key and(or) 'enter' to continue after closing open windows:")
    if check_if_weekend():
        print("TODAY IS WEEKEND, SO NO TIMESHEET :)\nExiting")
        time.sleep(2)
    else:
        if check_if_national_holiday(get_last_day()):
            print("TODAY IS A NATIONAL HOLIDAY, PLEASE MANUALLY UPDATE TIME-SHEET")
        else:
            chrome_driver = open_chrome()
            # if today is first day of the month OR 16th day of the month, start a new time sheet
            if begin_period():
                click_create_timesheet_button(chrome_driver)
                save_new_timesheet(chrome_driver)
            else:
                click_edit_button(chrome_driver)

            input_field_id = find_add_time_field(chrome_driver)
            if input_field_id != -1:
                add_time(chrome_driver, input_field_id)
                save(chrome_driver)

                if end_period():
                    submit(chrome_driver)
                    final_submit(chrome_driver)

            print("exiting")
            chrome_driver.close()
