
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import staleness_of
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.common.exceptions import NoSuchElementException        

import time
from contextlib import contextmanager
from datetime import datetime

driver = webdriver.Chrome()


# Thanks to http://www.obeythetestinggoat.com/how-to-get-selenium-to-wait-for-page-load-after-a-click.html
def wait_for(condition_function):
    start_time = time.time()
    while time.time() < start_time + 3:
        if condition_function():
            return True
        else:
            time.sleep(0.1)
    raise Exception('Timeout waiting for {}'.format(condition_function.__name__))

def parseSingleDate(date_str):
    return datetime.strptime(date_str, "%b %d, %Y %I:%M %p") if ':' in date_str else datetime.strptime(date_str, "%b %d, %Y")
def parseStartEnd(date_str):
    if '(All day)' in date_str:
        return datetime.strptime(date_str, "%b %d, %Y (All day)"), None
    
    try:
        s1, s2 = date_str.split(' - ')
    except ValueError:
        print(date_str)
        raise ValueError
    d1 = parseSingleDate(s1)

    d2 = datetime.strptime(s2, "%I:%M %p") if len(s2)<=8 else parseSingleDate(s2) 
    if d2.year == 1900:
        d2 = datetime(year=d1.year,month=d1.month,day=d1.day,hour=d2.hour,minute=d2.minute,second=d2.second)
    
    return d1, d2

class timer(object):
    def __init__(self, callback):
        self.callback = callback
    def __enter__(self):
        self.started = datetime.now()
    def __exit__(self, *_):
        if self.started:
            self.callback(datetime.now()-self.started)
class page_loaded(object):
    def __init__(self, browser):
        self.browser = browser
    def __enter__(self):
        self.old_page = self.browser.find_element_by_tag_name('html')
    def page_has_loaded(self):
        new_page = self.browser.find_element_by_tag_name('html')
        return new_page.id != self.old_page.id
    def __exit__(self, *_):
        wait_for(self.page_has_loaded)

@contextmanager
def wait_for_page_load(timeout=30):
    old_page = driver.find_element_by_tag_name('html')
    yield WebDriverWait(driver, timeout).until(staleness_of(old_page))
    #time.sleep(0.5)

def check_exists_by_xpath(xpath):
    try:
        return driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False

def submit():
    """Presses any 'submit' button"""
    safeclick(wait_for_element('//*[@type="submit"]'))
def safeclick(element, xtra_safe=False):
    if xtra_safe:
        loc = element.rect
        element = driver.execute_script(
            "return document.elementFromPoint(arguments[0], arguments[1]);",
            loc['x'],
            loc['y'])
    driver.execute_script('arguments[0].click()',element)

def wait_for_element(xpath, timeout = 30):
    return WebDriverWait(driver, timeout).until(presence_of_element_located((By.XPATH, xpath)))

def wait_for_element_pass(xpath='//*[@class="progress"]', timeout = 30):
    """Waits for an element to appear, then disappear. Default case: Progress bar"""
    while True: # Shitty do while loop
        element = wait_for_element(xpath, timeout=timeout/2)
        WebDriverWait(driver, timeout/2).until(staleness_of(element))
        time.sleep(0.1)
        if len(driver.find_elements_by_xpath(xpath)) == 0:
            break

def signin(username:str='',password:str=''):
    driver.get('https://go.microsoft.com/fwlink/p/?LinkID=873020')
    wait_for_page_load()

    # If you've reached the login page, prompt for user and pass
    #print(driver.current_url)
    if 'login' in driver.current_url:
        userbox = wait_for_element('//*[@name="loginfmt"]')
        userbox.send_keys(username)
        submit()

        wait_for_element_pass()

        passbox = driver.find_element_by_xpath('//*[@name="passwd"]')
        passbox.send_keys(password)
        submit()

        wait_for_page_load()
        submit()

    # Microsoft redirects here rather often, so we must sleep :(
    wait_for_element(xpath='//app-header-bar[@enable-navigation="true"]')
    driver.implicitly_wait(2.0)

    # Return if you end up on the teams page
    return 'teams.microsoft.com' in driver.current_url

def calendar():
    cal_xpath = '//button[@aria-label="Calendar Toolbar"]'
    if check_exists_by_xpath(cal_xpath): driver.find_element_by_xpath(cal_xpath).click()
    
    wait_for_page_load()
    if not 'calendarv2' in driver.current_url:
        driver.get('https://teams.microsoft.com/_#/calendarv2')
    
    # warning: calendar might still be loading
    return 'calendarv2' in driver.current_url

def cal_event():
    assert calendar()
    cal_base_xpath = '//div[@aria-label="Calendar grid view"]/div/div/div/div/div/div/div/div'

    cal_event_xpath = cal_base_xpath+'//div[contains(@class, "event") and contains(@role, "button")]'
    wait_for_element(xpath=cal_event_xpath)
    cal_events = driver.find_elements_by_xpath(cal_event_xpath)
    print('p e p e e')

    event_title = '//div[@class="default"]//div[contains(@class,"meeting-header-peek") and contains(@class,"__subject")]'
    event_date = '//div[@class="default"]//div[contains(@class,"meeting-header-peek") and contains(@class,"__date")]'
    event_loc = '//div[@class="default"]//div[contains(@class,"location-peek-location") and contains(@class,"__block")]'
    event_class = '//div[@class="default"]//div[contains(@class,"channel-peek-channel") and contains(@class,"__blockString")]'
    event_organs = '//div[@class="default"]//div[contains(@class,"participants-peek-participants") and contains(@class,"__peekParticipantsContainer")]'
    
    events = list()
    for cal_event in cal_events:
        #driver.execute_script('arguments[0].click()',cal_event)
        #cal_event.click()
        safeclick(cal_event)
        #driver.implicitly_wait(0.1)

        event_dict = dict()
        event_dict['Title'] = driver.find_element_by_xpath(event_title).text

        # need walrus to be gud
        # Currently slow if an element does not exist (~2s)
        if check_exists_by_xpath(event_date):
            date_string = driver.find_element_by_xpath(event_date).text.replace('\n',' ')
            start, end = parseStartEnd(date_string)
            event_dict['Start'] = start
            if end:
                event_dict['End'] = end
        
        # This location check takes forever. (~2s)
        if check_exists_by_xpath(event_loc):
            event_dict['Location'] = driver.find_element_by_xpath(event_loc).text

        
        if check_exists_by_xpath(event_class):
            event_dict['Team'] = driver.find_element_by_xpath(event_class).text
            
        
        if check_exists_by_xpath(event_organs):
            participants = driver.find_element_by_xpath(event_organs).text
            participants = participants.split('\n')
            event_dict['Participants'] = [tuple(participants[i:i+2]) for i in range(0,len(participants),2)]

        events.append(event_dict)
    return events


if __name__ == "__main__":
    print('Log into teams')
    username = input('Username: ')
    password = input('Password: ')
    print(signin(username=username,password=password))
    print(calendar())
    with timer(print):
        print('Calendar events:')
        print(len(cal_event()))
        print('Time to get events:')