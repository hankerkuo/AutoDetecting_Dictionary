# C:/Users/NR/AppData/Local/conda/conda/envs/py27/python
# -*- coding: utf-8 -*-
import pythoncom, pyHook
import win32clipboard
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

CTRL_DETECT = False
COPY_DETECT = False

def DownKey(event):
    global CTRL_DETECT
    global COPY_DETECT
    # print ('Key:', event.Key)
    if CTRL_DETECT is True:
        if event.Key == 'Lcontrol':
            print 'Directly search vocabulary:'
            hm.UnhookKeyboard()
            # in order to input unicode text(for chinese utf-8), add .decode('utf-8') after raw_input
            vocab = raw_input().decode('utf-8')

            background = driver.find_element_by_tag_name('body')
            background_2 = driver_2.find_element_by_tag_name('body')
            background.send_keys(Keys.CONTROL + Keys.HOME)
            background_2.send_keys(Keys.CONTROL + Keys.HOME)
            try:
                wait = WebDriverWait(driver, 5)
                wait_2 = WebDriverWait(driver_2, 5)
                elem = wait.until(EC.element_to_be_clickable((By.NAME, 'p')))
                elem_2 = wait_2.until(EC.element_to_be_clickable((By.NAME, 'q')))
            except:
                print "website connction error, please try again"

            elem.clear()
            elem_2.clear()
            elem.send_keys(vocab)
            elem_2.send_keys(vocab)
            elem.submit()
            elem_2.submit()
            hm.HookKeyboard()
            print 'Detecting copy action ...'
        CTRL_DETECT = False
    elif event.Key == 'Lcontrol':
        CTRL_DETECT = True
        COPY_DETECT = True
    return True


def OnKeyboardEvent(event):
    global COPY_DETECT
    # print ('MessageName:',event.MessageName)
    # print ('Message:',event.Message)
    # print ('Time:',event.Time)
    # print ('Window:',event.Window)
    # print ('WindowName:',event.WindowName)
    # print ('Ascii:', event.Ascii, chr(event.Ascii))
    # print ('Key:', event.Key)
    # print ('KeyID:', event.KeyID)
    # print ('ScanCode:', event.ScanCode)
    # print ('Extended:', event.Extended)
    # print ('Injected:', event.Injected)
    # print ('Alt', event.Alt)
    # print ('Transition', event.Transition)acquired
    # print ('---')
    if COPY_DETECT is True and event.Key == 'C':
        win32clipboard.OpenClipboard()
        # win32clipboard.CF_UNICODETEXT -> let the unicode be possible
        vocab = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()

        background = driver.find_element_by_tag_name('body')
        background_2 = driver_2.find_element_by_tag_name('body')
        background.send_keys(Keys.CONTROL + Keys.HOME)
        background_2.send_keys(Keys.CONTROL + Keys.HOME)
        try:
            wait = WebDriverWait(driver, 5)
            wait_2 = WebDriverWait(driver_2, 5)
            elem = wait.until(EC.element_to_be_clickable((By.NAME, 'p')))
            elem_2 = wait_2.until(EC.element_to_be_clickable((By.NAME, 'q')))
        except:
            print "website connction error, please try again"

        elem.clear()
        elem_2.clear()
        elem.send_keys(vocab)
        elem_2.send_keys(vocab)
        elem.submit()
        elem_2.submit()
    COPY_DETECT = False
    # if CTRL_DETECT is True and event.Key == 'Lcontrol':
    #     print 'Directly search vocabulary:'
    #     hm.UnhookKeyboard()
    #     # in order to input unicode text(for chinese utf-8), add .decode('utf-8') after raw_input
    #     vocab = raw_input().decode('utf-8')
    #
    #     background = driver.find_element_by_tag_name('body')
    #     background_2 = driver_2.find_element_by_tag_name('body')
    #     background.send_keys(Keys.CONTROL + Keys.HOME)
    #     background_2.send_keys(Keys.CONTROL + Keys.HOME)
    #     try:
    #         wait = WebDriverWait(driver, 5)
    #         wait_2 = WebDriverWait(driver_2, 5)
    #         elem = wait.until(EC.element_to_be_clickable((By.NAME, 'p')))
    #         elem_2 = wait_2.until(EC.element_to_be_clickable((By.NAME, 'q')))
    #     except:
    #         print "website connction error, please try again"
    #
    #     elem.clear()
    #     elem_2.clear()
    #     elem.send_keys(vocab)
    #     elem_2.send_keys(vocab)
    #     elem.submit()
    #     elem_2.submit()
    #     hm.HookKeyboard()
    #     print 'Detecting copy action ...'
# return True to pass the event to other handlers
#     return (event.Ascii in (ord('a'), ord('A')))
    return True

print 'Waiting for browser open ...'
driver = webdriver.Chrome()
driver_2 = webdriver.Chrome()
driver.set_window_position(-1866, 0)
driver_2.set_window_position(-950, 0)
# driver.maximize_window()
driver.get('https://tw.dictionary.search.yahoo.com')
driver_2.get('http://alldic.daum.net/index.do?dic=eng')

# create a hook manager
hm = pyHook.HookManager()
# watch for all keyboard events
hm.KeyDown = DownKey
hm.KeyUp = OnKeyboardEvent
# set the hook
hm.HookKeyboard()
# show able to detect
print 'Detecting copy action ...'
# wait forever
pythoncom.PumpMessages()

# close the browsers
driver.close()
driver_2.close()