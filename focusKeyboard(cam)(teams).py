import socket
import threading
import pyautogui
import time
import pywinauto
import warnings
import requests
warnings.simplefilter('ignore', category=UserWarning)
pyautogui.FAILSAFE = False
urlToCamera = "http://10.33.2.109:8000/"

cameraCommands = {"returnToHome":"""<?xml version="1.0"?><methodCall>
<methodName>returnToHome</methodName>
<params>
</params>
</methodCall>""", "startTracking":"""<?xml version="1.0"?><methodCall>
<methodName>startTracking</methodName>
<params>
<param>
<value>
<boolean>1</boolean>
</value>
</param>
</params>
</methodCall>""", "stopTracking":"""<?xml version="1.0"?><methodCall>
<methodName>stopTracking</methodName>
<params>
<param>
<value>
<boolean>1</boolean>
</value>
</param>
</params>
</methodCall>"""}
app = pywinauto.application.Application()
Teams = pywinauto.Application(backend="uia")
def find(list, item):
    try:
        s = list.index(item)
    except ValueError:
        s = -1
    return s
def presskeys(keys):
    keys = keys.split(",")
    for i in keys:
        pyautogui.keyDown(i)
    time.sleep(.1)
    for i in keys:
        pyautogui.keyUp(i)


def openAppList(backend="win32"):


    windows = pywinauto.Desktop(backend=backend).windows()
    names = []
    for w in windows:
        names.append(w.window_text())
    return names


def setFocus(name, ignore=[], backend="win32"):
    if backend == "uia":
        use = Teams
    elif backend == "win32":
        use = app
    openApps = openAppList(backend=backend)
    if name not in openApps:
        name = generateBestMacth(name, ignore, backend=backend)
        if name == None:
            return False

    try:
        #h = pywinauto.findwindows.find_window(title=name)
        #print(h)
        use.connect(title = name, found_index = 0)
        app_dialog = use.window(title=name, found_index = 0)
        app_dialog.exists()
        app_dialog.set_focus()
    except pywinauto.findwindows.WindowNotFoundError:
        return False

    return True
def generateBestMacth(name, ignore, backend="win32"):
    relevantApps = openAppList(backend=backend)
    containsName = []
    for i in relevantApps:
        if (name in i) and (i not in ignore):
            containsName.append(i)
    if len(containsName) <=0 :
        return
    else:
        s = containsName[0]

    for i in containsName:
        i = str(i)
        if len(i) < len(s):
            s = i
    return s
def startup():
    global conn

    global path
    global bank
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    bank = ""

    s.bind(("", 2001))

    s.listen(2)

    conn, ip = s.accept()
    print("Connected", ip)
    threading.Thread(target=checkConnectiomn).start()
def actions(command):
    if command != "t":
        print("Command:", command)



    # keys:ctrl,c*
    if command[:4] == "keys":
        presskeys(command[5:])
    # focus:zoom*
    if command[:5] == "focus":
        setFocus(command[6:])
    # zoomShortcut
    if command == "zoomShortcut":





        setFocus("VideoFrameWnd")


    if command == "teamsShortcut":


        setFocus("| Microsoft Teams", backend="uia")
    if command == "stopTracking":
        try:
            requests.post(urlToCamera, data=cameraCommands["stopTracking"])
        except:
            print("fail")
    if command == "startTracking":
       try:
           requests.post(urlToCamera, data=cameraCommands["startTracking"])
       except:
           print("fail")
    if command == "returnHome":
        try:
            requests.post(urlToCamera, data=cameraCommands["returnToHome"])
        except:
            print("fail")




def checkConnectiomn():
    while True:
        try:
            try:

                time.sleep(5)
                conn.send(b"test")


            except ConnectionResetError:
                break
        except ConnectionAbortedError:
            break
startup()
#bank = ""

while True:
    try:
        data = conn.recv(50)

        bank = bank + data.decode()

        while True:
            spot = bank.find("*")
            if spot != -1:
                command = bank[:spot]
                bank = bank[spot + 1:]
                actions(command)
            if spot == -1:
                break



    except Exception as e:
        print(e)
        print("restart")
        startup()

