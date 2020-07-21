import time
import pywinauto
from pywinauto import timings

while True:
    try:
        app = pywinauto.Application()
        app.connect(title='Cisco Webex Meetings')
        main_window = app.window(title='Cisco Webex Meetings')
        if main_window.is_visible():
            app.connect(title='Cisco Jabber')
            #app.window(title='Cisco Jabber').set_focus()
            app.window(title='Cisco Jabber').type_keys("%{F4}")
            time.sleep(180)
    except (pywinauto.findwindows.WindowNotFoundError, pywinauto.timings.TimeoutError, pywinauto.findwindows.ElementNotFoundError):
        print("Window not found.")
        time.sleep(5)
