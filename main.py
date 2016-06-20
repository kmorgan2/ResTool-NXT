#                     ____              ______               __   _   __ _  __ ______   _____    ____
#                    / __ \ ___   _____/_  __/____   ____   / /  / | / /| |/ //_  __/  |__  /   / __ \
#                   / /_/ // _ \ / ___/ / /  / __ \ / __ \ / /  /  |/ / |   /  / /      /_ <   / / / /
#                  / _, _//  __/(__  ) / /  / /_/ // /_/ // /  / /|  / /   |  / /     ___/ /_ / /_/ /
#                 /_/ |_| \___//____/ /_/   \____/ \____//_/  /_/ |_/ /_/|_| /_/     /____/(_)\____/
#        Developed by Kevin Morgan for ResTech, a division of Housing and Dining @ Northern Illinois University
#              Copyright(c) 2016. Board of Trustees of Northern Illinois University. All Rights Reserved.
#

# GUI Imports
from Tkinter import *
from ttk import Style, Progressbar, Notebook, Frame as TTKFrame, LabelFrame as TTKLabelFrame, Button as TTKButton, \
    Entry as TTKEntry, Label as TTKLabel

# Network/Web Imports
from twisted.internet import tksupport, reactor, threads, task
from twisted.internet.protocol import Protocol, Factory
# from twisted.internet.endpoints import TCP4ServerEndpoint
from twisted.logger import Logger, textFileLogObserver
import requests
import webbrowser

# General Imports
import os
import socket
import ctypes
import platform
import time
import logging

# Process/Window Control Imports
from pywinauto.application import Application
from pywinauto.controls.common_controls import ListViewWrapper
from win32gui import GetPixel, GetWindowDC, ReleaseDC
from win32com.client import Dispatch
import win32clipboard
import subprocess

# Registry and File Manipulation
import _winreg as reg
import io

# Disable warnings to console
sys.stderr = open(os.devnull, 'w')

# THE ALMIGHTY DEBUG FLAG. SET THIS TO TRUE WITH GREAT CAUTION
DEBUG = True
# THE LESS ALMIGHTY NO_POST FLAG. SET THIS TO TRUE UNTIL WE GET OUR API ENDPOINTS AND DEV ACCESS
NO_POST = False

# Logging
# noinspection PyTypeChecker
log = Logger(observer=textFileLogObserver(io.open(os.path.expanduser("~\Desktop\log.txt"), 'a')))
logging.getLogger("requests").setLevel(logging.WARNING)
log.info("Starting ResTool")

# Globals - API
API_HOST = 'intranet'  # Which subdomain to use (before restech.niu.edu)
if DEBUG:
    API_HOST = 'dev'  # Subdomain is dev if we're in development

API_BASE = "http://" + API_HOST + ".restech.niu.edu/tickets/api/restool"  # URL of the API directory

API_HEARTBEATS = API_BASE + "/heartbeats"  # API endpoint for heartbeats
API_EVENTS = API_BASE + "/events"  # API endpoint for events
API_TOKENS = API_BASE + "/tokens"  # API endpoint for tokens

EXPIRE_SECONDS = 60  # How long a default heartbeat should last

heartbeat = "Idle"  # The global for storing/modifying the heartbeat

api_count = 0  # How many times we've requested any API endpoint

# Check for Admin
if not ctypes.windll.shell32.IsUserAnAdmin():
    ctypes.windll.user32.MessageBoxW(0, u"Please make sure to run ResTool with Administrator rights.",
                                     u"Admin Required!", 16)
    log.critical("ResTool was started without admin privileges and was unable to run.")
    sys.exit()

# ------------------------------------------------------ GUI Init ------------------------------------------------------

root = Tk()  # create root window
tksupport.install(root)  # bind the GUI to a twisted reactor
root.protocol('WM_DELETE_WINDOW', None)  # unbind the close
root.iconbitmap('.\\icon.ico')  # icon is icon
root.title("ResTool NXT 2.0: Beta?")  # title is title
root.resizable(0, 0)  # prevent resize

# Create necessary styles
s = Style()
s.configure('White.TFrame', background='white')
s.configure('TLabelframe', background='white')
s.configure('TLabelframe.Label', background='white')
s.configure('Red.TFrame', background='red')
s.configure('Red.TLabel', foreground='white', background='red')


# And style modifiers
# ----------------------------------------------------------
# Function: red_status_area
# Purpose:  change the coloration of status area to red
# Preconds: GUI Initialized
# Postconds:GUI status area is red themed
# Arguments:
# ----------------------------------------------------------
def red_status_area():
    root_empty_space.configure(style="Red.TFrame")
    root_progressbar_label.configure(style="Red.TLabel")
    root.config(background='red')


# ----------------------------------------------------------
# Function: default_status_area
# Purpose:  change the coloration of the status area to def
# Preconds: GUI Initialized
# Postconds:GUI status area is default TTK style
# Arguments:
# ----------------------------------------------------------
def default_status_area():
    root_empty_space.configure(style="TFrame")
    root_progressbar_label.configure(style="TLabel")
    root.configure(background='SystemButtonFace')  # Discovered this is the default background color


# ---------------------------------------------------- GUI Classes -----------------------------------------------------


# A class encapsulation of the token dialog
class TokenDialog:
    # ----------------------------------------------------------
    # Function: TokenDialog::__init__
    # Purpose:  Initialize TokenDialog
    # Preconds: None
    # Postconds:Token Dialog is shown on screen or error thrown
    # Arguments:
    #           parent: the toplevel window we inherit from
    #           token_mgr: a Token class object for verification
    # ----------------------------------------------------------
    def __init__(self, parent, token_mgr):
        self.token_manager = token_mgr
        self.token = None
        top = self.top = Toplevel(parent)
        self.top.title("Enter ResTool Token")
        self.top.iconbitmap('icon.ico')  # icon is icon

        self.label = TTKLabel(top, text="Token:")
        self.label.grid(row=0, column=0)
        self.num = TTKEntry(top)
        self.num.grid(padx=5, row=0, column=1)

        Style().configure("Error.TLabel", foreground='red')

        b = TTKButton(top, text="OK", command=self.ok)
        b.grid(pady=5, columnspan=2)

    # ----------------------------------------------------------
    # Function: TokenDialog::ok
    # Purpose:  callback/command for the OK button. Verifies tok
    # Preconds: TokenDialog is initialized
    # Postconds:A valid token has been entered and the window
    #           dismissed, OR input field cleared, error colored
    # Arguments:
    # ----------------------------------------------------------
    def ok(self):
        if self.num.get():
            if self.token_manager.validate_token(self.num.get()):
                self.token = self.num.get()
                self.top.destroy()
            else:
                self.num.delete(0, END)
                self.label.configure(style="Error.TLabel")

    # ----------------------------------------------------------
    # Function: TokenDialog::get_token
    # Purpose:  return the token after the window is dismissed
    # Preconds: None
    # Postconds:None
    # Arguments:
    # ----------------------------------------------------------
    def get_token(self):
        return self.token


# A class encapsulation of the status bar
class ResToolStatusBar(Frame):
    # ----------------------------------------------------------
    # Function: ResToolStatusBar::__init__
    # Purpose:  initialize the status bar object
    # Preconds: have a token or are set offline
    # Postconds:status bar ready to be screened on GUI
    # Arguments:
    #           master: the Frame we inherit from
    #           **kw: kwargs as necessary, passed to sc init
    # ----------------------------------------------------------
    def __init__(self, master, **kw):
        Frame.__init__(self, master, **kw)
        # Display Windows version in leftmost box
        self.version_text = StringVar(value="Undefined")
        self.set_version()
        self.version_label = Label(self, textvariable=self.version_text, borderwidth=1, relief=SUNKEN)
        self.version_label.pack(side=LEFT, fill=X, expand=True)
        # Display IP in center box
        self.ip_text = StringVar(value="Not Connected")
        self.set_ip()
        self.ip_label = Label(self, textvariable=self.ip_text, borderwidth=1, relief=SUNKEN)
        self.ip_label.pack(side=LEFT, fill=X, expand=True)
        # Display token in rightmost box
        self.token_text = StringVar(value="Offline")
        self.token_label = Label(self, textvariable=self.token_text, anchor=W, borderwidth=1, relief=SUNKEN)
        self.token_label.pack(side=LEFT, fill=X, expand=True)
        self.token_label.bind("<Button-1>", lambda _: token_manager.record_token())
        self.set_ticket()

    # ----------------------------------------------------------
    # Function: ResToolStatusBar::set_version
    # Purpose:  place the windows version in the status bar
    # Preconds: version_text exists
    # Postconds:version_text displays windows version & bitness
    # Arguments:
    # ----------------------------------------------------------
    def set_version(self):
        try:
            _ = os.environ["PROGRAMFILES(X86)"]  # if there is a program files x86 dir
            bits = ' x64'
        except KeyError:  # it's like an if-else except all the logic relies on a KeyError (or lack thereof)
            bits = ' x32'
        value = 'Windows ' + str(platform.platform()).split('-')[1] + bits
        self.version_text.set(value)

    # ----------------------------------------------------------
    # Function: set_ip
    # Purpose:  place the IP address in the status bar
    # Preconds: ip_text exists
    # Postconds:ip_text displays the most useful IP
    # Arguments:
    # ----------------------------------------------------------
    def set_ip(self):
        # gets the valid hostname of this machine (borrowed this code from stack overflow)
        ip_address = [l for l in (
            [ip for ip in socket.gethostbyname_ex(socket.gethostname())[2] if not ip.startswith("127.")][:1], [
                [(soc.connect(('8.8.8.8', 80)), soc.getsockname()[0], soc.close()) for soc in
                 [socket.socket(socket.AF_INET, socket.SOCK_DGRAM)]][0][1]]) if l][0][0]
        self.ip_text.set(ip_address)

    # ----------------------------------------------------------
    # Function: set_ticket
    # Purpose:  place the ticket number in the status bar
    # Preconds: we have a token, token_text exists
    # Postconds:token_text displays ticket number or Offline
    # Arguments:
    # ----------------------------------------------------------
    def set_ticket(self):
        self.token_text.set(get_ticket())


# ---------------------------------------------------- Token Class -----------------------------------------------------
# ----------------------------------------------------------
# Function: open_registry
# Purpose:  initialize registry for use by other functions
# Preconds: Windows Registry is in a readable state
# Postconds:A valid handle to HKLM\Software\ResTech exists
# Arguments: 
# ----------------------------------------------------------
def open_registry(w=False):
    try:
        r = reg.OpenKey(reg.HKEY_LOCAL_MACHINE, "Software\\ResTech\\", 0, reg.KEY_ALL_ACCESS if w else reg.KEY_READ)
    except (
            WindowsError, KeyError):  # The ResTech Key in the registry does not exist. Create and return accordingly
        r = reg.CreateKeyEx(reg.HKEY_LOCAL_MACHINE, "Software\\ResTech\\", 0,
                            reg.KEY_ALL_ACCESS if w else reg.KEY_READ)
    return r


# A class to manage the token, saving and reading from the registry, and handling no-token occurrences
class Token:
    # ----------------------------------------------------------
    # Function: Token::__init__
    # Purpose:  initialize token management class
    # Preconds: None
    # Postconds:Token object initialized with token or offline
    # Arguments:
    #           token: a token to use before searching for one
    # ----------------------------------------------------------
    def __init__(self, token=None):
        self.token = token
        if not self.token:  # try to get it from the registry
            self.get_token(validate=True)
            if self.token:
                log.debug("Retrieved stored token from registry.")
        if NO_POST and not self.token:  # it's not in the registry, but we are offline
            self.set_token("Offline")
        elif not self.token:  # try to get it from the user
            self.record_token()

    # ----------------------------------------------------------
    # Function: Token::record_token
    # Purpose:  record a token from the user using TokenDialog
    # Preconds: Token pre-initialized
    # Postconds:Token recorded or set offline
    # Arguments:
    # ----------------------------------------------------------
    def record_token(self):
        global NO_POST
        NO_POST = False
        td = TokenDialog(root, self)
        while root.wait_window(td.top):
            time.sleep(5)
        token = td.get_token()
        if token():
            self.set_token(token)
        if not self.validate_token(accept_offline=False) or not self.get_token():
            self.set_token("Offline")
            NO_POST = True
            log.warn("No valid token supplied; user exited prompt. Operating offline.")
        else:
            log.debug("Token {token} recorded from user", token=self.token)

    # ----------------------------------------------------------
    # Function: Token::validate_token
    # Purpose:  validate the token with Intranet
    # Preconds: Token number initialized (or is supplied)
    # Postconds:None
    # Arguments:
    #           token: override stored token to verify
    # ----------------------------------------------------------
    def validate_token(self, token=None, accept_offline=True):
        if not token:
            token = self.token
        if token == "Offline" and accept_offline or NO_POST:
            return True
        if not token:
            return False

        global api_count
        api_count += 1
        log.debug("API called to validate Token. This is call {ac} since launch.", ac=api_count)

        return requests.get(API_TOKENS + "/" + token).status_code == 200

    # ----------------------------------------------------------
    # Function: Token::get_token
    # Purpose:  returns the token or retrieves it from the regi-
    #           stry. If validate is true, it checks validity
    # Preconds: Token exists or in registry
    # Postconds:Token may be validated
    # Arguments:
    #           validate: whether the token should be checked
    # ----------------------------------------------------------
    def get_token(self, validate=False, accept_offline=True):
        if not self.token:
            try:
                self.token = str(reg.QueryValueEx(open_registry(), "token")[0])
            except (WindowsError, KeyError):  # The Token Value in the ResTech does not exist. Return placeholder text.
                self.token = None
                log.debug("Attempt to get token when none was initialized or stored.")
        if validate and not (self.token and self.validate_token(accept_offline=accept_offline)):
            self.record_token()
        return self.token

    # ----------------------------------------------------------
    # Function: Token::set_token
    # Purpose:  sets the token in the registry
    # Preconds: None
    # Postconds:Token stored in registry
    # Arguments:
    #           value: the token to set in-object and -registry
    # ----------------------------------------------------------
    def set_token(self, value):
        self.token = value
        try:
            reg.SetValueEx(open_registry(True), "token", 0, reg.REG_SZ, value)
            log.info("Recorded new token {token}.", token=self.token)
        except (WindowsError, KeyError):  # This should be impossible. If it happens, something has gone wrong
            ctypes.windll.user32.MessageBoxW(0, u'Unrecoverable error in writing token to registry.', u'Registry Error',
                                             16)
            log.critical("Registry not in valid state when storing token.")
            sys.exit()

    # ----------------------------------------------------------
    # Function: Token::delete_token
    # Purpose:  deletes the token from the registry
    # Preconds: Token in registry
    # Postconds:Token not in registry
    # Arguments:
    # ----------------------------------------------------------
    def delete_token(self):
        self.token = None
        try:
            reg.DeleteKey(reg.HKEY_LOCAL_MACHINE, "Software\\ResTech\\")
            log.error("Removed recorded token.")
        except (WindowsError, KeyError):  # This should be impossible. If it happens, something has gone wrong
            ctypes.windll.user32.MessageBoxW(0, u'Unrecoverable error in removing ResTech registry data.',
                                             u'Registry Error', 16)
            log.critical("Registry not in valid state when removing token.")
            sys.exit()


token_manager = Token()  # Initialize the token in global scope


# -------------------------------------------------- Safe Mode Class ---------------------------------------------------
class SafeModeTracker:
    # ----------------------------------------------------------
    # Function: SafeModeTracker::__init__
    # Purpose:  initialize the safe mode tracker
    # Preconds: None
    # Postconds:Safe mode option identified
    # Arguments:
    # ----------------------------------------------------------
    def __init__(self):
        self.safe_mode_set = False
        self.get_safe_mode_setting()
        log.debug("On launch, safe mode was {value}.", value="enabled" if self.safe_mode_set else "disabled")

    # ----------------------------------------------------------
    # Function: SafeModeTracker::get_safe_mode_setting
    # Purpose:  check whether the system will boot to safe mode
    # Preconds: None
    # Postconds:Safe Mode button correctly reflects state
    # Arguments:
    # ----------------------------------------------------------
    def get_safe_mode_setting(self):
        output = subprocess.check_output(["bcdedit", "/enum", "{default}"], shell=True)
        if "safeboot" in output:
            self.safe_mode_set = True
            # Set GUI to reflect safe mode setting
            msc_button.configure(text="Disable")
        else:
            self.safe_mode_set = False
            msc_button.configure(text="Enable")

    # ----------------------------------------------------------
    # Function: safeModeTracker::enable_safe_mode
    # Purpose:  Enable safe mode on next boot
    # Preconds: Safe mode disabled for next boot
    # Postconds:Safe mode enabled for next boot
    # Arguments:
    # ----------------------------------------------------------
    def enable_safe_mode(self):
        # Modify the BCD Configuration to enable safe boot with networking
        subprocess.Popen(["bcdedit", "/set", "{default}", "safeboot", "network"], shell=True).wait()
        self.get_safe_mode_setting()
        if self.safe_mode_set:
            log.info("Safe mode has been enabled.")
        else:
            log.error("There was an error enabling safe mode.")

    # ----------------------------------------------------------
    # Function: SafeModeTracker::disable_safe_mode
    # Purpose:  Disable safe mode on next boot
    # Preconds: Safe mode enabled for next boot
    # Postconds:Safe mode disabled for next boot
    # Arguments:
    # ----------------------------------------------------------
    def disable_safe_mode(self):
        # Modify the BCD Configuration to disable safeboot
        subprocess.Popen(["bcdedit", "/deletevalue", "{default}", "safeboot"], shell=True).wait()
        self.get_safe_mode_setting()
        if not self.safe_mode_set:
            log.info("Safe mode has been disabled.")
        else:
            log.error("There was an error disabling safe mode.")

    # ----------------------------------------------------------
    # Function: toggle_safe_mode
    # Purpose:  Toggle safe mode option for next boot
    # Preconds: None
    # Postconds:Safe Mode setting for next boot inverted
    # Arguments:
    # ----------------------------------------------------------
    def toggle_safe_mode(self):
        if self.safe_mode_set:
            self.disable_safe_mode()
        else:
            self.enable_safe_mode()
        if ctypes.windll.user32.MessageBoxW(0,
                                            u"Would you like to restart now into " + (u'safe' if self.safe_mode_set else
                                                                                      u"normal") + u" mode?",
                                            u"Restart to switch modes?", 36) == 6:
            update_heartbeat("ResTool will restart the computer to switch to " + ('safe' if self.safe_mode_set else
                                                                                  "normal") + " mode.", 3600)
            create_startup_shortcut()
            subprocess.check_call(['shutdown', '/t', '0', '/r', '/f'])
        else:
            log.info("A restart was not performed after toggling safe mode.")


# --------------------------------------------------- API Functions ----------------------------------------------------


# ----------------------------------------------------------
# Function: update_heartbeat
# Purpose:  Sets the current heartbeat and posts to Intranet
# Preconds: heartbeat exists
# Postconds:heartbeat updated and posted to Intranet
# Arguments:
#           text: the text of the heartbeat
#           expire_seconds: how long the heartbeat is valid
# ----------------------------------------------------------
def update_heartbeat(text, expire_seconds=EXPIRE_SECONDS):
    global heartbeat
    heartbeat = text
    post_heartbeat(expire_seconds)


# ----------------------------------------------------------
# Function: post_heartbeat
# Purpose:  Posts the current heartbeat to the Intranet
# Preconds: heartbeat exists
# Postconds:heartbeat has been posted to Intranet
# Arguments:
#           expire_seconds: how long the heartbeat is valid
# ----------------------------------------------------------
def post_heartbeat(expire_seconds=EXPIRE_SECONDS):
    global heartbeat, API_HEARTBEATS
    if not NO_POST and not token_manager.get_token() == "Offline":
        global api_count
        api_count += 1
        log.debug("API called to post Heartbeat. This is call {ac} since launch.", ac=api_count)
        response = requests.post(API_HEARTBEATS,
                                 {'token': token_manager.get_token(), 'status': heartbeat,
                                  'expire_seconds': expire_seconds})
        try:
            assert response.status_code == 200
        except AssertionError:  # Error posting the heartbeat
            if response.status_code == 401:  # Bad token
                log.error('Invalid Token {token} when posting heartbeat.', token=token_manager.get_token())
                if token_manager.get_token(validate=True, accept_offline=False):
                    post_heartbeat(expire_seconds)
            elif response.status_code == 400:  # Bad status or expire_seconds
                log.error('Invalid data when posting heartbeat "{heartbeat}"', heartbea=heartbeat)
            else:  # Uncaught (500, 404, etc)
                log.error('Uncaught HTTP Error {rc}. token in POST heartbeats: with token: {t}, heartbeat:"{h}",' +
                          'and expire_seconds: {e}', t=token_manager.get_token(), h=heartbeat, e=expire_seconds,
                          rc=response.status_code)
                if token_manager.get_token(validate=True, accept_offline=False):
                    post_heartbeat(expire_seconds)
        return response.status_code
    return "Offline"


# ----------------------------------------------------------
# Function: post_event
# Purpose:  Post an event to the Intranet
# Preconds: None
# Postconds:Event has been posted to the Intranet
# Arguments:
#           text: the text of the event
# ----------------------------------------------------------
def post_event(text):
    if not NO_POST and not token_manager.get_token() == "Offline":

        global api_count
        api_count += 1
        log.debug("API called to post Event. This is call {ac} since launch.", ac=api_count)

        response = requests.post(API_EVENTS, {'token': token_manager.get_token(), 'text': text})
        try:
            assert response.status_code == 200
        except AssertionError:  # Error posting the event
            if response.status_code == 401:  # Bad token
                log.error('Invalid Token {token} when posting event.', token=token_manager.get_token())
                if token_manager.get_token(validate=True, accept_offline=False):
                    post_event(text)
            elif response.status_code == 400:  # Bad text
                log.error('Invalid data when posting event "{text}"', text=text)
            else:  # Uncaught (500, 404, etc)
                log.error('Uncaught HTTP Error {rc} in POST events: with token:{token} and event: "{text}"',
                          token=token_manager.get_token(),
                          text=text, rc=response.status_code)
                if token_manager.get_token(validate=True, accept_offline=False):
                    post_event(text)
        return response.status_code
    return "Offline"


# ----------------------------------------------------------
# Function: stop
# Purpose:  halt the program and inform Intranet
# Preconds: Program running
# Postconds:Program not running
# Arguments: 
# ----------------------------------------------------------
def stop():
    update_heartbeat("ResTool was closed at the request of the user.", 3600)
    log.info("ResTool was closed at the request of the user.")
    reactor.stop()
    return None


# ----------------------------------------------------------
# Function: invalidate_token
# Purpose:  Invalidate the current token
# Preconds: Valid token exists
# Postconds:Current token invalidated and unbound
# Arguments: 
# ----------------------------------------------------------
def invalidate_token():
    if not NO_POST:
        global api_count
        api_count += 1
        log.debug("API called to invalidate Token. This is call {ac} since launch.", ac=api_count)

        response = requests.post(API_TOKENS + "/" + token_manager.get_token(), {'action': 'invalidate'})
        try:
            assert response.status_code == 200
        except AssertionError:  # Error posting the event
            if response.status_code == 401:  # Bad token
                log.error('Invalid Token {token} when invalidating token.', token=token_manager.get_token())
                if token_manager.get_token(validate=True, accept_offline=False):
                    invalidate_token()
            elif response.status_code == 400:  # Bad action
                log.error('Invalid action "invalidate" when invalidating token.')
            else:  # Uncaught (500, 404, etc)
                log.error('Uncaught HTTP Error {rc} in POST token: {token} with action: invalidate',
                          token=token_manager.get_token(), rc=response.status_code)
                if token_manager.get_token(validate=True, accept_offline=False):
                    invalidate_token()
        return response.status_code
    token_manager.delete_token()
    return "Offline"


# ----------------------------------------------------------
# Function: get_ticket
# Purpose:  Get the ticket for the current token
# Preconds: Token exists and is valid
# Postconds:None
# Arguments: 
# ----------------------------------------------------------
def get_ticket():
    print token_manager.get_token()
    if not NO_POST and not token_manager.get_token() == "Offline":
        global api_count
        api_count += 1
        log.debug("API called to get Ticket. This is call {ac} since launch.", ac=api_count)

        response = requests.get(API_TOKENS + "/" + token_manager.get_token())
        try:
            assert response.status_code == 200
            ticket = response.json()['ticket']['id']
            return ticket
        except AssertionError:
            if response.status_code == 404:
                log.error("Invalid Token {token} when getting ticket.")
                if token_manager.get_token(validate=True, accept_offline=False):
                    get_ticket()
            else:
                log.error("Uncaught HTTP Error {rc} in GET token: {token}", token=token_manager.get_token(),
                          rc=response.status_code)
                if token_manager.get_token(validate=True, accept_offline=False):
                    get_ticket()
            return "ERROR"
    return "Offline"


# ----------------------------------------------------------
# Function: create_startup_shortcut
# Purpose:  Create a startup shortcut
# Preconds: None
# Postconds:A startup shortcut exists
# Arguments:
# ----------------------------------------------------------
def create_startup_shortcut():
    path = os.path.join('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp', "ResTool.lnk")
    workingdir = os.path.abspath('.')
    target = workingdir + '\ResTool NXT.exe'

    if DEBUG:
        workingdir += '\\dist'
        target = workingdir + '\\main.exe'

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = workingdir
    shortcut.IconLocation = target
    shortcut.save()


def delete_startup_shortcut():
    try:
        os.remove('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\ResTool.lnk')
        log.debug("On launch, restart link detected and removed")
    except OSError:
        log.debug("On launch, restart link was not detected.")


delete_startup_shortcut()  # Call this ASAP to make sure this is done


# -------------------------------------------------- Callback Methods --------------------------------------------------


# Called to thread-defer a GUI callback.
# ----------------------------------------------------------
# Function: app_defer
# Purpose:  Perform thread-safe execution of a function and:
#               -log its start
#               -update the GUI status area
#               -update the heartbeat
#               -bind callback and errback functions
# Preconds: None
# Postconds:The specified function deferred to open thread
# Arguments:
#           function: the function to defer
#           name: the name of the function for logging
# ----------------------------------------------------------
def app_defer(function, name):
    log.info("Running {name}", name=name)
    default_status_area()
    root_progressbar_label.configure(text="Running " + name)
    root_progressbar.config(mode='indeterminate')
    root_progressbar.start()
    d = threads.deferToThread(function)
    update_heartbeat("Running " + name)
    d.addCallbacks(app_callback, app_errback, errbackArgs=(name,))
    return None


# Called on app success
# ----------------------------------------------------------
# Function: app_callback
# Purpose:  Log successful completion of a deferred function
# Preconds: Deferred function completes successfully
# Postconds:Results logged and GUI updated to idle state
# Arguments:
#           return_text: data returned by deferred function
# ----------------------------------------------------------
def app_callback(return_text):
    if return_text is not None:
        log.info("Finished. " + return_text)
        # update Action
        post_event(return_text)
        update_heartbeat("Idle")
    # reset GUI
    root_progressbar_label.configure(text="Ready for adventure...")
    root_progressbar.config(mode='indeterminate')
    root_progressbar.start()


# Called on app error
# ----------------------------------------------------------
# Function: app_errback
# Purpose:  Log error termination of a deferred function
# Preconds: Deferred function throws an uncaught exception
# Postconds:Error logged and GUI updated to error state
# Arguments:
#           error: the uncaught exception
#           name: the name of the function for logging
# ----------------------------------------------------------
def app_errback(error, name):
    log.error("Error running {name}\n" + error.getErrorMessage(), name=name)
    if DEBUG:
        log.failure("", error)
    # update status
    post_event("While running " + str(name) + ", ResTool encountered an error:\n" + error.getErrorMessage())
    update_heartbeat("Error running " + str(name))
    # set GUI components
    red_status_area()
    root_progressbar_label.configure(text="Error in " + name)
    root_progressbar.config(mode='determinate', value=0)
    root_progressbar.stop()


# ----------------------------------------------------------
# Function: run_cf
# Purpose:  Run ComboFix scan. Win 8+ incompatible
# Preconds: ComboFix is compatible, CF.exe exists
# Postconds:ComboFix scan run and results gathered
# Arguments:
# ----------------------------------------------------------
def run_cf():
    return None


# Function to defer run_cf
def run_cf_defer():
    app_defer(run_cf, "ComboFix")


# ----------------------------------------------------------
# Function: run_mwb
# Purpose:  Run MalwareBytes Anti Malware Scan
# Preconds: MWB.exe exists
# Postconds:MWBAM scan run and results gathered
# Arguments: 
# ----------------------------------------------------------
def run_mwb():
    application = Application()
    if os.path.isfile("C:\Program Files (x86)\Malwarebytes Anti-Malware\mbam.exe"):
        try:
            application.start("C:\Program Files (x86)\Malwarebytes Anti-Malware\mbam.exe")
        except UserWarning:  # bitness error
            pass
        time.sleep(3)
    else:
        try:
            application.Start("programs/MWB.exe")
        except UserWarning:  # bitness error
            pass
        time.sleep(0.25)
        try:
            application.connect(title_re="Select Set.*")
        except UserWarning:  # bitness error
            pass
        application.SelectSetupLanguage.OK.ClickInput()
        application.SetupMalwarebytesAntiMalware.Wait("exists enabled visible ready", 30)
        application.SetupMalwarebytesAntiMalware.Next.ClickInput()
        application.SetupMalwarebytesAntiMalware.Iaccepttheagreement.Wait("exists enabled visible ready", 30)
        application.SetupMalwarebytesAntiMalware.Iaccepttheagreement.ClickInput()
        application.SetupMalwarebytesAntiMalware.Next.Wait("exists enabled visible ready", 30)
        application.SetupMalwarebytesAntiMalware.Next.ClickInput()
        application.SetupMalwarebytesAntiMalware.Next.Wait("exists enabled visible ready", 30)
        application.SetupMalwarebytesAntiMalware.Next.ClickInput()
        application.SetupMalwarebytesAntiMalware.Next.Wait("exists enabled visible ready", 30)
        application.SetupMalwarebytesAntiMalware.Next.ClickInput()
        application.SetupMalwarebytesAntiMalware.Next.Wait("exists enabled visible ready", 30)
        application.SetupMalwarebytesAntiMalware.Next.ClickInput()
        application.SetupMalwarebytesAntiMalware.Next.Wait("exists enabled visible ready", 30)
        application.SetupMalwarebytesAntiMalware.Next.ClickInput()
        application.SetupMalwarebytesAntiMalware.Install.Wait("exists enabled visible ready", 30)
        application.SetupMalwarebytesAntiMalware.Install.ClickInput()
        application.SetupMalwarebytesAntiMalware.Finish.Wait("exists enabled visible ready", 3600)
        application.SetupMalwarebytesAntiMalware.Finish.ClickInput()

        time.sleep(15)

    # Start scan window
    try:
        application.connect(title_re="Malwarebytes.*Home.*")
    except UserWarning:  # bitness error
        pass
    application.MalwarebytesAntiMalwareHome.ClickInput(coords=(493, 477))

    # ----------------------------------------------------------
    # Function: run_mwb::get_color
    # Purpose:  get the color of a pixel given XY coordinates
    # Preconds:window exists
    # Postconds:None
    # Arguments:
    #           window: a PyWinAuto window specification
    #           xpos: the pixel x-coordinate (left to right)
    #           ypos: the pixel y-coordinate (top to bottom)
    # ----------------------------------------------------------
    def get_color(window, xpos, ypos):
        handle = window.handle
        dc = GetWindowDC(handle)
        color = int(GetPixel(dc, xpos, ypos))
        ReleaseDC(handle, dc)
        return (color & 0xff), ((color >> 8) & 0xff), ((color >> 16) & 0xff)

    # Wait for "Finish" button to appear
    while get_color(application.MalwareBytesAntiMalwareHome, 350, 478) != (37, 120, 230):
        time.sleep(5)

    # Copy results to clipboard
    application.MalwarebytesAntiMalwareHome.ClickInput(coords=(800, 520))
    application.MalwarebytesAntiMalwareHome.TypeKeys("{DOWN}{ENTER}")
    time.sleep(0.1)
    win32clipboard.OpenClipboard()
    cb_results = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()

    txt_list = [(cb_results.split("Processes: ")[1].split('\r\n')[0]),
                (cb_results.split("Modules: ")[1].split('\r\n')[0]),
                (cb_results.split("Registry Keys: ")[1].split('\r\n')[0]),
                (cb_results.split("Registry Values: ")[1].split('\r\n')[0]),
                (cb_results.split("Registry Data: ")[1].split('\r\n')[0]),
                (cb_results.split("Folders: ")[1].split('\r\n')[0]),
                (cb_results.split("Files: ")[1].split('\r\n')[0]),
                (cb_results.split("Physical Sectors: ")[1].split('\r\n')[0])
                ]

    application.MalwarebytesAntiMalwareHome.Close()
    num_infect = sum([int(x) for x in txt_list])
    return str(num_infect) + " detections."


# Function to defer run_mwb
def run_mwb_defer():
    app_defer(run_mwb, "Malwarebytes")


# ----------------------------------------------------------
# Function: run_eset
# Purpose:  Run an ESET Online Scanner scan
# Preconds: ESET.exe exists, an active network connection
# Postconds:ESET scan run and results recorded
# Arguments:
# ----------------------------------------------------------
def run_eset():
    app = Application()
    try:
        app.Start("programs/ESET.exe")
    except UserWarning:  # bitness error
        pass
    time.sleep(0.5)
    try:
        app.connect(title_re="Terms.*")
    except UserWarning:  # bitness error
        pass
    app.Termsofuse.YesIaccepttheTermsOfUse.Wait("exists enabled visible ready", 30)
    app.Termsofuse.YesIaccepttheTermsOfUse.ClickInput()
    app.Termsofuse.Start.Wait("exists enabled visible ready", 30)
    app.Termsofuse.Start.ClickInput()
    time.sleep(7)
    try:
        app.connect(title_re="ESET.*")
    except UserWarning:  # bitness error
        pass
    time.sleep(0.5)
    app.ESETOnlineScanner.Enableprotectionofpotentiallyunwantedapplications.Wait("exists enabled visible ready", 30)
    app.ESETOnlineScanner.Enableprotectionofpotentiallyunwantedapplications.ClickInput()
    app.ESETOnlineScanner.Start.Wait("exists enabled visible ready", 30)
    app.ESETOnlineScanner.Start.ClickInput()

    # Wait for scan to finish
    app.ESETOnlineScanner.Uninstallapplicationonclose.Wait("exists enabled visible ready", 7200)

    # Grab number of infected from file
    eset_logfile = open('C:\Program Files (x86)\ESET\ESET Online Scanner\log.txt')
    eset_logfile_text = eset_logfile.read()
    num_cleaned = eset_logfile_text.split("# cleaned=")[1].split('\n')[0]
    num_found = eset_logfile_text.split("# found=")[1].split('\n')[0]
    eset_logfile.close()
    time.sleep(0.5)

    # Check if found = cleaned
    print(num_cleaned)
    print(num_found)
    if num_cleaned != num_found:
        raise ValueError(
            "Not all infected files were able to be cleaned!\n\tFound:" + num_found + "\n\tCleaned" + num_cleaned)

    app.ESETOnlineScanner.Uninstallapplicationonclose.ClickInput()
    app.ESETOnlineScanner.Finish.Wait("exists enabled visible ready", 30)
    app.ESETOnlineScanner.Finish.ClickInput()
    if app.ESETOnlineScanner.Close():
        return str(num_found) + " detections."


# Function to defer run_eset
def run_eset_defer():
    app_defer(run_eset, "ESET")


# ----------------------------------------------------------
# Function: run_sas
# Purpose:  Runs a SUEPERAntiSpyware Scan
# Preconds: SAS.exe exists
# Postconds:SAS scan ran and results recorded
# Arguments: 
# ----------------------------------------------------------
def run_sas():
    app = Application()
    try:
        app.Start("programs/SAS.exe")
    except UserWarning:  # bitness error
        pass
    time.sleep(1)
    try:
        app.connect(title_re="SUPERAntiSpyware.*")
    except UserWarning:  # bitness error
        pass
    app.SUPERAntiSpywareFreeEditionSetup.Next.Wait("exists enabled visible ready", 30)
    app.SUPERAntiSpywareFreeEditionSetup.Next.ClickInput()
    app.SUPERAntiSpywareFreeEditionSetup.IAgree.Wait("exists enabled visible ready", 30)
    app.SUPERAntiSpywareFreeEditionSetup.IAgree.ClickInput()
    app.SUPERAntiSpywareFreeEditionSetup.Onlyformecurrentuser.Wait("exists enabled visible ready", 30)
    app.SUPERAntiSpywareFreeEditionSetup.Onlyformecurrentuser.ClickInput()
    app.SUPERAntiSpywareFreeEditionSetup.Next.Wait("exists enabled visible ready", 30)
    app.SUPERAntiSpywareFreeEditionSetup.Next.ClickInput()
    app.SUPERAntiSpywareFreeEditionSetup.Next.Wait("exists enabled visible ready", 30)
    app.SUPERAntiSpywareFreeEditionSetup.Next.ClickInput()
    app.SUPERAntiSpywareFreeEditionSetup.Next.Wait("exists enabled visible ready", 30)
    app.SUPERAntiSpywareFreeEditionSetup.Next.ClickInput()
    app.SUPERAntiSpywareFreeEditionSetup.Finished.Wait("exists enabled visible ready", 30)
    app.SUPERAntiSpywareFreeEditionSetup.Finished.ClickInput()
    time.sleep(4)
    try:
        app.connect(title_re=".*Trial*").connect(title_re=".*Trial*")
    except UserWarning:  # bitness error
        pass
    finally:
        app.connect(title_re=".*Trial*")
    app.SUPERAntiSpywareProfessionalTrial.Decline.Wait("exists enabled visible ready", 30)
    app.SUPERAntiSpywareProfessionalTrial.Decline.ClickInput()
    time.sleep(2)
    try:
        app.connect(title_re=".*Edition")
    except UserWarning:  # bitness error
        pass
    # Click scan
    app.SUPERAntiSpywareFreeEdition[u'5'].ClickInput()
    app.SUPERAntiSpywareFreeEdition.ClickInput(coords=(500, 150))

    # ----------------------------------------------------------
    # Function: run_sas::get_color
    # Purpose:  get the color of a pixel given XY coordinates
    # Preconds:window exists
    # Postconds:None
    # Arguments:
    #           window: a PyWinAuto window specification
    #           xpos: the pixel x-coordinate (left to right)
    #           ypos: the pixel y-coordinate (top to bottom)
    # ----------------------------------------------------------
    def get_color(window, x, y):
        handle = window.handle
        dc = GetWindowDC(handle)
        color = int(GetPixel(dc, x, y))
        ReleaseDC(handle, dc)
        return (color & 0xff), ((color >> 8) & 0xff), ((color >> 16) & 0xff)

    # Wait for "Finish" button to appear
    while get_color(app.SUPERAntiSpywareFreeEdition, 350, 478) != (49, 58, 63):  # todo: Coords wrong
        time.sleep(5)

    return None


# Function to defer run_sas
def run_sas_defer():
    app_defer(run_sas, "SuperAntiSpyware")


# ----------------------------------------------------------
# Function: run_sb
# Purpose:  Run a SpyBot Search and Destroy Scan
# Preconds: SB.exe exists
# Postconds:SB scan run and results recorded
# Arguments: 
# ----------------------------------------------------------
def run_sb():
    application = Application()
    if not os.path.isfile("C:\Program Files (x86)\Spybot - Search & Destroy 2\SDScan.exe"):
        application.start("programs/SB.exe")
        time.sleep(2.5)
        application.connect(title_re=".*Select Setup Language.*")
        application.SelectSetupLanguage.OK.ClickInput()
        for _ in range(0, 3):
            application.SetupSpybotSearchDestroy.Next.ClickInput()
        application.SetupSpybotSearchDestroy.Iaccepttheagreement.ClickInput()
        application.SetupSpybotSearchDestroy.Next.ClickInput()
        application.SetupSpybotSearchDestroy.Install.ClickInput()
        while not application.SetupSpybotSearchDestroy.Finish.Exists():
            time.sleep(5)
            if application.Setup.OK.Exists():
                application.Setup.OK.ClickInput()
        application.SetupSpybotSearchDestroy.Finish.ClickInput()

    # Update
    application.start("C:\Program Files (x86)\Spybot - Search & Destroy 2\SDUpdate.exe")
    time.sleep(10)
    application.UpdateSpybotSearchDestroy.Update.Wait("exists enabled visible ready", 120)
    application.UpdateSpybotSearchDestroy.Update.ClickInput()
    application.UpdateSpybotSearchDestroy.Update.Wait("exists enabled visible ready", 36000)
    update_log = application.UpdateSpybotSearchDestroy['Edit'].WindowText()
    update_count = update_log.split("[+] Installed ")[-1].split(' updates')[0]
    try:
        int(update_count)
    except ValueError:
        update_count = '0'
    update_text = "Spybot Search and Destroy installed " + str(update_count) + " updates"
    if "Unable" in application.UpdateSpybotSearchDestroy["Edit"].WindowText():
        update_text += ", but errors were encountered. This is normal and the scan will continue"
    update_text += "."
    log.info(update_text)
    post_event(update_text)
    application.UpdateSpybotSearchDestroy.Close()

    # Scan
    application.start("C:\Program Files (x86)\Spybot - Search & Destroy 2\SDScan.exe")
    time.sleep(10)
    application.SystemScanSpybotSearchDestroy.StartaScan.ClickInput()
    while application.SystemScanSpybotSearchDestroy["Button2"].WindowText() == 'Stop scan':
        time.sleep(10)
    # Button changes control number 2->3
    application.SystemScanSpybotSearchDestroy["Button3"].Wait("exists enabled visible ready")
    if application.SystemScanSpybotSearchDestroy["Button3"].WindowText() == 'Show scan results':
        application.SystemScanSpybotSearchDestroy.Showscanresults.ClickInput()
    item_count = ListViewWrapper(application.SystemScanSpybotSearchDestroy.TListView.handle).ItemCount()
    if item_count > 0:
        application.SystemScanSpybotSearchDestroy.Fixselected.ClickInput()
        time.sleep(30)
    application.SystemScanSpybotSearchDestroy.Close()
    return "Spybot detected and removed " + str(item_count) + " threats."


# Function to defer run_sb
def run_sb_defer():
    app_defer(run_sb, "Spybot")


# ----------------------------------------------------------
# Function: run_hc
# Purpose:  Run a Trend Micro HouseCall Scan
# Preconds: HC.exe exists
# Postconds:HC scan run and results recorded
# Arguments: 
# ----------------------------------------------------------
def run_hc():
    application = Application().start("programs/HC.exe")
    time.sleep(2.0)
    application = application.connect(class_name_re="#32770", title=u'')  # connect to window
    application.Dialog[u'4'].ClickInput(coords=(35, 315))  # accept EULA
    application.Dialog[u'4'].ClickInput(coords=(550, 360))  # click OK

    application.Dialog[u'4'].ClickInput(coords=(460, 220))  # Click Settings

    application.Dialog[u'5'].ClickInput(coords=(25, 115))  # Click Full Scan
    application.Dialog[u'5'].ClickInput(coords=(300, 360))  # Click OK

    application.Dialog[u'4'].ClickInput(coords=(460, 150))  # Click Scan Now

    # Wait for scan completion

    return None


# Function to defer run_hc
def run_hc_defer():
    app_defer(run_hc, "TrendMicro HouseCall")


# ----------------------------------------------------------
# Function: run_cc
# Purpose:  Run a CCLeaner file and registry cleaning
# Preconds: CC.exe exists
# Postconds:CC file and at least 1 reg scan run and recorded
# Arguments: 
# ----------------------------------------------------------
def run_cc():
    # if installed
    results = []
    application = Application()
    if os.path.isfile("C:\Program Files\CCleaner\CCleaner.exe"):
        # Run Cleaner
        try:
            application.start("C:\Program Files\CCleaner\CCleaner.exe")
        except UserWarning:
            pass
        time.sleep(1)
        try:
            application.connect(title_re=".*CCleaner Professional.*")
        except UserWarning:
            pass
        application.CCleanerProfessional.Close()
        application.PiriformCCleanerProfessionalEdition.RunCleaner.Wait("exists enabled visible ready", 60)
        application.PiriformCCleanerProfessionalEdition.RunCleaner.ClickInput()
        application.Dialog.OK.Wait("exists enabled visible ready", 60)
        application.Dialog.OK.ClickInput()
        while not application.PiriformCCleanerProfessionalEdition.RunCleaner.IsEnabled():
            time.sleep(1)
            if application.top_window_() != application.PiriformCCleanerProfessionalEdition:
                if application.Dialog.Exists():
                    application.Dialog.Yes.ClickInput()
        results.append(application.PiriformCCleanerProfessionalEdition.Edit.WindowText().split('\r\n')[2])
        # Run Registry Cleaner
        application.PiriformCCleanerProfessionalEdition.Registry.ClickInput()
        registry_value = 1
        while not (registry_value == 0 or (len(results) > 2 and registry_value == results[-2])):
            application.PiriformCCleanerProfessionalEdition.ScanforIssues.Wait("exists enabled visible ready", 60)
            application.PiriformCCleanerProfessionalEdition.ScanforIssues.ClickInput()
            application.PiriformCCleanerProfessionalEdition.ScanforIssues.Wait("exists enabled visible ready", 3600)
            application.PiriformCCleanerProfessionalEdition.FixselectedIssues.ClickInput()
            application.CCleaner.Yes.Wait("exists enabled visible ready", 60)
            application.CCleaner.Yes.ClickInput()
            application.SaveAs.Save.Wait("exists enabled visible ready", 60)
            application.SaveAs.Save.ClickInput()
            application.Dialog.Wait("exists enabled visible ready", 3600)
            i = 0
            while application.Dialog.FixIssue.IsEnabled():
                i += 1
                application.Dialog.FixIssue.ClickInput()
            application.Dialog['Close'].ClickInput()
            results.append(i)
            registry_value = i
        # Close CCleaner
        application.PiriformCCleanerProfessionalEdition.Close()
        return "CCleaner generated the following results: \n\tCleaner: " + results[0] + "\n\tRegistry: " + ", ".join(
            ([str(e) for e in results[1:]])) + " (Stopped due to identical passes)" if results[-1] == results[
            -2] else ''

    else:
        # install
        try:
            application.start("programs/CC.exe")
        except UserWarning:
            pass
        application.CCleanerProfessionalSetup.Next.Wait("exists enabled visible ready", 60)
        application.CCleanerProfessionalSetup.Next.ClickInput()
        application.CCleanerProfessionalSetup.Install.Wait("exists enabled visible ready", 60)
        application.CCleanerProfessionalSetup.Install.ClickInput()
        application.CCleanerProfessionalSetup.Finish.Wait("exists enabled visible ready", 3600)
        application.CCleanerProfessionalSetup.ViewReleasenotes.ClickInput()
        application.CCleanerProfessionalSetup.Finish.ClickInput()
        time.sleep(2)
        try:
            application.connect(title_re=".*CCleaner Professional.*")
        except UserWarning:
            pass
        application.CCleanerProfessional.Close()
        application.PiriformCCleanerProfessionalEdition.Close()
        return run_cc()


# Function to defer run_cc
def run_cc_defer():
    app_defer(run_cc, "CCleaner")


# ----------------------------------------------------------
# Function: rem_scans
# Purpose:  Uninstall virus scanners, CCleaner, & All-in-one
# Preconds: None
# Postconds:MWB, ESET, SB, SAS, CC, AIO not installed
# Arguments: 
# ----------------------------------------------------------
def rem_scans():
    uninstalled_programs = []
    application = Application()

    # remove MWB
    if os.path.isfile("C:\Program Files (x86)\Malwarebytes Anti-Malware\unins000.exe"):
        try:
            application.start("C:\Program Files (x86)\Malwarebytes Anti-Malware\unins000.exe")
        except UserWarning:
            pass
        time.sleep(1)
        try:
            application.connect(title_re=".*Malwarebytes Anti-Malware Uninstall.*")
        except UserWarning:
            pass
        application.MalwarebytesAntiMalwareUninstall.Yes.Wait("exists visible enabled ready", 30)
        application.MalwarebytesAntiMalwareUninstall.Yes.ClickInput()
        application.MalwarebytesAntiMalwareUninstall.No.Wait("exists visible enabled ready", 3600)
        application.MalwarebytesAntiMalwareUninstall.No.ClickInput()
        uninstalled_programs.append("MWB")
    # remove ESET
    if os.path.isfile("C:\Program Files (x86)\ESET\ESET Online Scanner\OnlineScannerUninstaller.exe"):
        try:
            application.start("C:\Program Files (x86)\ESET\ESET Online Scanner\OnlineScannerUninstaller.exe")
        except UserWarning:
            pass
        application.ESETOnlineScanner.Yes.Wait("exists visible enabled ready", 30)
        application.ESETOnlineScanner.Yes.ClickInput()
        application.ESETOnlineScanner.OK.Wait("exists visible enabled ready", 3600)
        application.ESETOnlineScanner.OK.ClickInput()
        os.remove("C:\\Program Files (x86)\\ESET\\ESET Online Scanner\\OnlineScannerUninstaller.exe")
        os.rmdir("C:\\Program Files (x86)\\ESET\\ESET Online Scanner\\")
        uninstalled_programs.append("ESET")
    # remove SAS
    # todo: implement SAS removal
    # remove SB
    if os.path.isfile("C:\Program Files (x86)\Spybot - Search & Destroy 2\unins000.exe"):
        try:
            application.start("C:\Program Files (x86)\Spybot - Search & Destroy 2\unins000.exe")
        except UserWarning:
            pass
        time.sleep(1)
        try:
            application.connect(title_re=".*Spybot - Search & Destroy Uninstall.*")
        except UserWarning:
            pass
        application.SpybotSearchDestroyUninstall.Yes.Wait("exists visible enabled ready", 30)
        application.SpybotSearchDestroyUninstall.Yes.ClickInput()
        application.top_window_().Next.Wait("exists visible enabled ready", 60)
        application.top_window_().Next.ClickInput()
        application.top_window_().Uninstall.Wait("exists visible enabled ready", 60)
        application.top_window_().IwanttoinformyouwhyIuninstallmyreasonis.ClickInput()
        application.top_window_().Uninstall.ClickInput()
        application.SpybotSearchDestroyUninstall.No.Wait("exists visible enabled ready", 3600)
        application.SpybotSearchDestroyUninstall.No.ClickInput()
        uninstalled_programs.append("SB")
    # remove CCleaner - we may be uninstalling either pro or not, so we're vague about windows
    if os.path.isfile("C:\Program Files\CCleaner\uninst.exe"):
        try:
            application.start("C:\Program Files\CCleaner\uninst.exe")
        except UserWarning:
            pass
        time.sleep(1)
        try:
            application.connect(title_re=".*CCleaner.*Uninstall.*")
        except UserWarning:
            pass
        application.top_window_().Next.Wait("exists visible enabled ready", 60)
        application.top_window_().Next.ClickInput()
        application.top_window_().Uninstall.Wait("exists visible enabled ready", 60)
        application.top_window_().Uninstall.ClickInput()
        application.top_window_().Finish.Wait("exists visible enabled ready", 3600)
        application.top_window_().Finish.ClickInput()
        uninstalled_programs.append("CC")
    # remove AIO
    if os.path.isfile("C:\Program Files (x86)\Tweaking.com\Windows Repair (All in One)\uninstall.exe"):
        try:
            application.start('"C:\Program Files (x86)\Tweaking.com\Windows Repair (All in One)\uninstall.exe" /U: C:' +
                              '"\Program Files (x86)\Tweaking.com\Windows Repair (All in One)\Uninstall\uninstall.xml"')
        except UserWarning:
            pass
        application.TweakingcomWindowsRepairUninstaller.Next.Wait("exists visible enabled ready", 60)
        application.TweakingcomWindowsRepairUninstaller.Next.ClickInput()
        application.TweakingcomWindowsRepairUninstaller.Finish.Wait("exists visible enabled ready", 3600)
        application.TweakingcomWindowsRepairUninstaller.Finish.ClickInput()
        uninstalled_programs.append("AIO")
    return "Removed the following programs: " + ", ".join(uninstalled_programs)


# Function to defer rem_scans
def rem_scans_defer():
    app_defer(rem_scans, "Scanner Uninstall")


# ----------------------------------------------------------
# Function: msc_toggle
# Purpose:  Button callback to toggle safe mode
# Preconds: None
# Postconds:Safe mode toggled
# Arguments: 
# ----------------------------------------------------------
def msc_toggle():
    safe_mode_tracker.toggle_safe_mode()
    return None


# ----------------------------------------------------------
# Function: run_mwbar
# Purpose:  Run a Malwarebytes Anti Rootkit Scan
# Preconds: MWBAR.exe exists
# Postconds:MWBAR scan run and results recorded
# Arguments: 
# ----------------------------------------------------------
def run_mwbar():
    return None


# Function to defer run_mwbar
def run_mwbar_defer():
    app_defer(run_mwbar, "Malwarebytes AntiRootkit")


# ----------------------------------------------------------
# Function: run_tdss
# Purpose:  Run a Kaspersky TDSSKiller Scan
# Preconds: TDSS.exe exists
# Postconds:TDSSKiller scan run and results recorded
# Arguments: 
# ----------------------------------------------------------
def run_tdss():
    return None


# Function to defer run_tdss
def run_tdss_defer():
    app_defer(run_tdss, "Kaspersky TDSSKiller")


# ----------------------------------------------------------
# Function: get_mse
# Purpose:  Install Security Essentials/Windows Defender
# Preconds: MSE.exe exists
# Postconds:MSE installed or Windows Defender enabled
# Arguments: 
# ----------------------------------------------------------
def get_mse():
    # if windows 7 - This is untested b/c I didn't have a win7 computer to test on.
    if platform.win32_ver()[0] == 7:

        # create application
        application = Application()
        try:
            application.start("\programs\MSE.exe")
        except UserWarning:
            pass
        application.MicrosoftSecurityEssentials.Wait("exists visible enabled ready", 3600)

        # welcome
        application.MicrosoftSecurityEssentials.Next.Wait("exists visible enabled ready", 30)
        application.MicrosoftSecurityEssentials.Next.ClickInput()

        # EULA
        application.MicrosoftSecurityEssentials.Iaccept.Wait("exists visible enabled ready", 30)
        application.MicrosoftSecurityEssentials.Iaccept.ClickInput()

        # Customer Improvement
        application.MicrosoftSecurityEssentials.Next.Wait("exists visible enabled ready", 30)
        application.MicrosoftSecurityEssentials.JointheCustomerExperienceImprovementProgram.ClickInput()
        application.MicrosoftSecurityEssentials.Next.ClickInput()

        # Optimize Security
        application.MicrosoftSecurityEssentials.Next.Wait("exists visible enabled ready", 30)
        application.MicrosoftSecurityEssentials.Turnonautomaticsamplesubmission.ClickInput()
        application.MicrosoftSecurityEssentials.Next.ClickInput()

        # ready to install
        application.MicrosoftSecurityEssentials.Install.Wait("exists visible enabled ready", 30)
        application.MicrosoftSecurityEssentials.Install.ClickInput()

        application.MicrosoftSecurityEssentials.Finish.Wait("exists visible enabled ready", 3600)
        application.MicrosoftSecurityEssentials.Finish.ClickInput()
        return "Microsoft Security Essentials Successfully Installed"
    # if windows 8 or newer - enable
    else:
        subprocess.call("control.exe /name Microsoft.WindowsDefender")
        result = ctypes.windll.user32.MessageBoxW(0,
                                                  u"Windows Defender is open. If disabled, it must be manually " +
                                                  u"enabled.\nFollow the prompts to enable Windows Defender and " +
                                                  u"then click OK here.\nIf you are unable to enable Windows " +
                                                  u"Defender, please click Cancel instead.",
                                                  u"Manually Enable Windows Defender:",
                                                  4161)
        if result == 1:  # yes
            return "Windows Defender Manually Enabled."
        elif result == 2:  # no
            raise UserWarning("Windows Defender Not Manually Enabled.")


# Function to defer get_mse
def get_mse_defer():
    app_defer(get_mse, "Windows Defender Install")


# ----------------------------------------------------------
# Function: run_pnf
# Purpose:  Open Programs and Features
# Preconds: None
# Postconds:Programs and Features open
# Arguments: 
# ----------------------------------------------------------
def run_pnf():
    try:
        subprocess.check_call(['control.exe', 'appwiz.cpl'])
    except subprocess.CalledProcessError:
        log.error("Could not launch Programs and Features")


# ----------------------------------------------------------
# Function: run_temp
# Purpose:  Run Temporary File Cleanup
# Preconds: None
# Postconds:Temporary File Cleanup launched
# Arguments: 
# ----------------------------------------------------------
def run_temp():
    try:
        subprocess.check_call(['cleanmgr.exe'])
    except subprocess.CalledProcessError:
        log.error("Could not launch Temporary File Cleanup")


# Function to defer run_temp
def run_temp_defer():
    app_defer(run_temp, "Temporary File Cleanup")


# ----------------------------------------------------------
# Function: run_ticket
# Purpose:  Open the ticket in the system webbrowser
# Preconds: System has a web browser, Token is valid
# Postconds:Ticket (login page) is open in browser
# Arguments: 
# ----------------------------------------------------------
def run_ticket():
    ticket = get_ticket()
    if ticket not in ["Offline", "ERROR"]:
        webbrowser.open("http://" + API_HOST + ".restech.niu.edu/tickets/" + ticket)
    else:
        webbrowser.open("http://" + API_HOST + ".restech.niu.edu/tickets")
    return None


# ----------------------------------------------------------
# Function: run_ipconfig
# Purpose:  Run an IPConfig reset
# Preconds: System has a network connection
# Postconds:IPConfig Reset completed
# Arguments: 
# ----------------------------------------------------------
def run_ipconfig():
    post_event("IPConfig Reset started. Completion of this event may not be logged.")
    update_heartbeat("IPConfig Reset will cause this machine to become temporarily unreachable", 3600)
    try:
        subprocess.check_call(['ipconfig.exe', '/release'])
        subprocess.check_call(['ipconfig.exe', '/flushdns'])
        subprocess.check_call(['ipconfig.exe', '/renew'])
        return "IPConfig Reset Complete."
    except subprocess.CalledProcessError as e:
        raise UserWarning("Uncaught IPConfig Behavior: \n" + ''.join(ch for ch in e.output if ch not in ['x00', 'x08']))


# Function to defer run_ipconfig
def run_ipconfig_defer():
    app_defer(run_ipconfig, "IPConfig Reset")


# ----------------------------------------------------------
# Function: run_winsock
# Purpose:  Runs a Winsock Reset
# Preconds: System has a network connection
# Postconds:Winsock, firewall, IPV4 services/catalogs reset
# Arguments: 
# ----------------------------------------------------------
def run_winsock():
    post_event("Winsock Reset started. Completion of this event may not be logged.")
    update_heartbeat("Winsock Reset requires a restart; this machine may be unreachable until the restarted.", 3600)
    try:
        subprocess.call(['netsh.exe', 'int', 'ip', 'reset'])
        subprocess.call(['netsh.exe', 'winsock', 'reset'])
        subprocess.call(['netsh.exe', 'winsock', 'reset', 'catalog'])
        subprocess.call(['netsh.exe', 'advfirewall', 'reset'])

        if ctypes.windll.user32.MessageBoxW(0, u"Would you like to restart now in order to complete the Winsock Reset?",
                                            u"Restart Required!", 52) == 6:
            create_startup_shortcut()
            subprocess.check_call(['shutdown', '/t', '0', '/r', '/f'])
        else:
            log.warn("Restart not completed after Winsock Reset. Network behavior may be erratic.")
        return "Winsock Reset complete. Reboot Necessary."

    except subprocess.CalledProcessError as e:
        raise UserWarning("Uncaught Winsock Behavior: \n" + ''.join(ch for ch in e.output if ch not in ['x00', 'x08']))


# Function to defer run_winsock
def run_winsock_defer():
    app_defer(run_winsock, "Winsock Reset")


# ----------------------------------------------------------
# Function: run_hidapters
# Purpose:  Removes hidden adapters known to cause net probs
# Preconds: devcon.exe exists
# Postconds:hidden adapters removed from device manager
# Arguments: 
# ----------------------------------------------------------
def run_hidapters():
    root_progressbar.stop()
    root_progressbar.config(mode='determinate', value=0)
    post_event("Hidden Adapter Removal started. Completion of this event may not be logged.")
    classes = ["TEREDO", "ISATAP", "TUNMP", "6TO4"]
    progress_increment = 100.0 / (11 * len(classes))
    for adapter in classes:
        for instance in range(0, 10):  # remove things the old way
            time.sleep(0.1)
            subprocess.check_call(r'programs\devcon.exe -r remove "@ROOT\*' + adapter + r'\000' + str(instance) + '"')
            root_progressbar.step(progress_increment)
        # For Win10, everything is virtual and friendly and we can use wildcards
        subprocess.check_call('programs\devcon.exe -r remove *' + adapter + '*')

    root_progressbar.config(value=0)

    return "Hidden Adapter Removal complete."


# Function to defer run_hidapters
def run_hidapters_defer():
    app_defer(run_hidapters, "Hidden Adapter Removal")


# ----------------------------------------------------------
# Function: run_wifi
# Purpose:  Install the NIUWireless profile
# Preconds: Wi-Fi.xml exists, Computer has Wi-Fi, using WZC
# Postconds:Wi-Fi Profile installed without credentials
# Arguments: 
# ----------------------------------------------------------
def run_wifi():
    try:
        return subprocess.check_output(["netsh", "wlan", "add", "profile", 'filename="programs\Wi-Fi.xml"'])
    except subprocess.CalledProcessError as e:
        raise UserWarning("Uncaught WiFi Add Behavior: \n" + ''.join(ch for ch in e.output if ch not in ['x00', 'x08']))


# Function to defer run_wifi
def run_wifi_defer():
    app_defer(run_wifi, "WiFi Profile Install")


# ----------------------------------------------------------
# Function: run_speed
# Purpose:  Run a speedtest in the system browser
# Preconds: Computer has network connection, browser
# Postconds:Speedtest website open in browser
# Arguments: 
# ----------------------------------------------------------
def run_speed():
    webbrowser.open("http://speedtest.niu.edu")
    return None


# ----------------------------------------------------------
# Function: run_netcpl
# Purpose:  Open Network and Sharing Center
# Preconds: None
# Postconds:Network and Sharing Center open
# Arguments: 
# ----------------------------------------------------------
def run_netcpl():
    try:
        subprocess.check_call(['control.exe', '/name', 'Microsoft.NetworkAndSharingCenter'])
    except subprocess.CalledProcessError:
        log.error("Could not launch Network and Sharing Center")


# ----------------------------------------------------------
# Function: run_sfc
# Purpose:  Run a System File Checker Scan
# Preconds: None
# Postconds:SFC scan completed and results recorded
# Arguments: 
# ----------------------------------------------------------
def run_sfc():
    try:
        output = subprocess.check_output("sfc /scannow")
        output = ''.join(ch for ch in output if ch not in ['x00', 'x08'])
        if "did not find any" in output:
            return "SFC Passed with no corrupt files."
        elif "could not perform" in output:
            raise UserWarning("A system error requires SFC to be run in Safe Mode.")
        elif "successfully repaired them" in output:
            return "SFC Failed but repaired all corrupt files."
        elif "unable to fix" in output:
            return "SFC Failed and could not repair corrupt files."
        else:
            raise UserWarning("Uncaught SFC Behavior: \n" + ''.join(ch for ch in output if ch not in ['x00', 'x08']))
    except subprocess.CalledProcessError as e:
        raise UserWarning("Uncaught SFC Behavior: \n" + ''.join(ch for ch in e.output if ch not in ['x00', 'x08']))


# Function to defer run_sfc
def run_sfc_defer():
    app_defer(run_sfc, "System File Checker")


# ----------------------------------------------------------
# Function: run_dism
# Purpose:  Run a DISM health check and repair
# Preconds: System compatible with DISM Check-Image commands
# Postconds:DISM scan run and results recorded
# Arguments: 
# ----------------------------------------------------------
def run_dism():
    try:
        output = subprocess.check_output("dism /Online /Cleanup-Image /RestoreHealth")
        output = ''.join(ch for ch in output if ch not in ['x00', 'x08'])
        if "The restore operation completed successfully." in output:
            return "DISM Successfully verified and/or made repairs."
        else:
            raise UserWarning("DISM Failed:\n" + ''.join(ch for ch in output if ch not in ['x00', 'x08']))
    except subprocess.CalledProcessError as e:
        raise UserWarning("Uncaught DISM Behavior: \n" + ''.join(ch for ch in e.output if ch not in ['x00', 'x08']))


# Function to defer run_dism
def run_dism_defer():
    app_defer(run_dism, "DISM")


# ----------------------------------------------------------
# Function: run_aio
# Purpose:  Run an All-in-One repair
# Preconds: AIO.exe exists
# Postconds:AIO scan run and results log
# Arguments: 
# ----------------------------------------------------------
def run_aio():
    return None


# Function to defer run_aio
def run_aio_defer():
    app_defer(run_aio, "All-In-One Repair")


# ----------------------------------------------------------
# Function: run_defrag
# Purpose:  Optimize or defragment all volumes
# Preconds: 
# Postconds:Drives optimized
# Arguments: 
# ----------------------------------------------------------
def run_defrag():
    try:
        output = subprocess.check_output("defrag C: /H /O /U /V")
        output = ''.join(ch for ch in output if ch not in ['x00', 'x08'])
        if "The operation completed successfully" in output:
            return "Disk Optimization or Defragmentation Complete."
        elif "An operation is currently in progress" in output:
            raise UserWarning("Disk Optimization currently in progress.")
        else:
            raise UserWarning("Disk Optimization failed:\n" + ''.join(ch for ch in output if ch not in ['x00', 'x08']))
    except subprocess.CalledProcessError as e:
        raise UserWarning(
            "Uncaught Disk Optimization Behavior: \n" + ''.join(ch for ch in e.output if ch not in ['x00', 'x08']))


# Function to defer run_defrag
def run_defrag_defer():
    app_defer(run_defrag, "Disk Optimizer")


# ----------------------------------------------------------
# Function: run_chkdsk
# Purpose:  Run Chkdsk verification on the system volume
# Preconds: None
# Postconds:Chkdsk run and recorded, repair queued if needed
# Arguments: 
# ----------------------------------------------------------
def run_chkdsk():
    try:
        output = subprocess.check_output("chkdsk")
        output = ''.join(ch for ch in output if ch not in ['x00', 'x08'])
        if "found no problems." in output:
            return "CHKDSK found no errors"
        else:
            subprocess.check_call("fsutil set dirty")
            return "CHKDSK found errors. They will be repaired at next reboot"
    except subprocess.CalledProcessError as e:
        raise UserWarning("Uncaught CHKDSK Behavior: \n" + ''.join(ch for ch in e.output if ch not in ['x00', 'x08']))


# Function to defer run_chkdsk
def run_chkdsk_defer():
    app_defer(run_chkdsk, "Disk Check")


# ----------------------------------------------------------
# Function: run_dmc
# Purpose:  Run the Device Management Console
# Preconds: None
# Postconds:Device Manager is open
# Arguments: 
# ----------------------------------------------------------
def run_dmc():
    try:
        subprocess.check_call(['mmc.exe', 'devmgmt.msc'])
    except subprocess.CalledProcessError:
        log.error("Could not launch Device Manager")


# ----------------------------------------------------------
# Function: rem_mse
# Purpose:  Uninstall/Disable MSE or Windows Defender
# Preconds: MSE/WD installed
# Postconds:MSE/WD uninstalled or disabled
# Arguments: 
# ----------------------------------------------------------
def rem_mse():
    return None


# Function to defer rem_mse
def rem_mse_defer():
    app_defer(rem_mse, "Windows Defender Uninstall")


# ----------------------------------------------------------
# Function: run_cpl
# Purpose:  Open the Control Panel
# Preconds: None
# Postconds:Control Panel open
# Arguments: 
# ----------------------------------------------------------
def run_cpl():
    try:
        subprocess.check_call(['control.exe'])
    except subprocess.CalledProcessError:
        log.error("Could not launch Control Panel")


# ----------------------------------------------------------
# Function: run_reg
# Purpose:  Open Regedit
# Preconds: None
# Postconds:Regedit open
# Arguments: 
# ----------------------------------------------------------
def run_reg():
    try:
        subprocess.check_call(['regedit.exe'])
    except subprocess.CalledProcessError:
        log.error("Could not launch Registry Editor")
    return None


# ----------------------------------------------------------
# Function: run_awp
# Purpose:  Install anywhere printer driver
# Preconds: AWP.exe exists
# Postconds:Anywhere Printer driver installed
# Arguments: 
# ----------------------------------------------------------
def run_awp():
    application = Application()
    try:
        application.Start("programs/awp.exe")
    except UserWarning:  # bitness error
        pass
    application.PackageWindowsAnywherePrintClient.Accept.Wait("exists enabled visible ready", 30)
    application.PackageWindowsAnywherePrintClient.Accept.ClickInput()
    application.PackageWindowsAnywherePrintClient.Install.Wait("exists enabled visible ready", 30)
    application.PackageWindowsAnywherePrintClient.Install.ClickInput()
    application.PackageWindowsAnywherePrintClient.Finish.Wait("exists enabled visible ready", 3600)
    try:
        assert "success" in application.PackageWindowsAnywherePrintClient.Edit.TextBlock()
        application.PackageWindowsAnywherePrintClient.Finish.ClickInput()
        return "AnywherePrint Client Successfully Installed"
    except AssertionError:
        application.PackageWindowsAnywherePrintClient.Finish.ClickInput()
        raise UserWarning(
            "Uncaught AnywherePrint Install Behavior:/n" +
            application.PackageWindowsAnywherePrintClient.Edit.TextBlock())


# Function to defer run_awp
def run_awp_defer():
    app_defer(run_awp, "Anywhere Printer Install")


# ------------------------------------------------ End Callback Methods ------------------------------------------------

# -------------------------------------------------- Network Classes ---------------------------------------------------


# A Twisted protocol class to receive communication via ResTool Web.

# noinspection PyClassHasNoInit
class ResToolWebResponder(Protocol):
    funcMap = {'A': update_heartbeat, 'B': run_cf, 'C': run_mwb, 'D': run_eset, 'E': run_sas, 'F': run_sb, 'G': run_hc,
               'H': run_cc, 'I': run_sfc, 'J': rem_scans_defer, 'K': msc_toggle, 'L': run_mwbar, 'M': run_tdss,
               'N': get_mse,
               'O': run_pnf, 'P': run_temp, 'Q': run_ticket, 'R': run_ipconfig, 'S': run_winsock, 'T': run_hidapters,
               'U': run_wifi, 'V': run_speed, 'W': run_netcpl, 'X': run_dism, 'Y': run_aio, 'Z': run_defrag,
               'a': run_chkdsk, 'b': run_dmc, 'c': rem_mse, 'd': run_cpl, 'e': run_reg, 'f': run_awp, 'g': stop,
               'h': invalidate_token}
    nameMap = {'B': "ComboFix", 'C': "Malwarebytes", 'D': "ESET", 'E': "SuperAntiSpyware", 'F': "Spybot",
               'G': "TrendMicro HouseCall", 'H': "CCleaner", 'I': "System File Checker",
               'L': "Malwarebytes AntiRootkit", 'M': "Kaspersky TDSSKiller", 'N': "Windows Defender Install",
               'P': "Temporary File Cleanup", 'R': "IPConfig Reset", 'S': "Winsock Reset",
               'T': "Hidden Adapter Removal", 'U': "WiFi Profile Creation", 'V': "Speed Test", 'X': "DISM",
               'Y': "All-In-One Repair", 'Z': "Disk Defragmenter", 'a': "Disk Check", 'c': "Windows Defender Uninstall",
               'f': "AnywherePrint Client Install"}

    # ----------------------------------------------------------
    # Function: dataReceived
    # Purpose:  handle remote control command
    # Preconds: Data exists in TCP stack
    # Postconds:The requested program has been called
    # Arguments:
    # ----------------------------------------------------------
    def dataReceived(self, data):
        self.transport.write("0")
        if data[0] not in ["A", "J", "K", "O", "Q", "W", "b", "d", "e", "g", "h"]:  # These fxns are not action-worthy
            app_defer(self.funcMap[data[0]], self.nameMap[data[0]])
        else:
            self.funcMap[data[0]]()


# noinspection PyClassHasNoInit
class ResToolWebResponderFactory(Factory):
    # ----------------------------------------------------------
    # Function: buildProtocol
    # Purpose:  Construct a responder for TCP remote control cmd
    # Preconds: 
    # Postconds:a responder exists on addr
    # Arguments:
    #           addr: the address to respond on
    # ----------------------------------------------------------
    def buildProtocol(self, addr):
        return ResToolWebResponder()


# ---------------------------------------------------- Main Program ----------------------------------------------------

# GUI creation (main program initialization)

# --Main Notebook
root_notebook = Notebook(root)

# ----Tab: Virus Scans
tab__virus_scans = TTKFrame(style='White.TFrame')
# ------Left Third
virus_scans__left = TTKFrame(tab__virus_scans, style='White.TFrame')
av_scanners_group = TTKLabelFrame(virus_scans__left, text="AV Scans")
cf_button = TTKButton(av_scanners_group, text="ComboFix", command=run_cf_defer, state=DISABLED)
cf_button.pack(padx=5, pady=5, fill=X)
mwb_button = TTKButton(av_scanners_group, text="Malwarebytes", command=run_mwb_defer)
mwb_button.pack(padx=5, pady=5, fill=X)
eset_button = TTKButton(av_scanners_group, text="ESET", command=run_eset_defer)
eset_button.pack(padx=5, pady=5, fill=X)
sas_button = TTKButton(av_scanners_group, text="SAS", command=run_sas_defer)
sas_button.pack(padx=5, pady=5, fill=X)
sb_button = TTKButton(av_scanners_group, text="Spybot", command=run_sb_defer)
sb_button.pack(padx=5, pady=5, fill=X)
hc_button = TTKButton(av_scanners_group, text="HouseCall", command=run_hc_defer, state=DISABLED)
hc_button.pack(padx=5, pady=5, fill=X)
av_scanners_group.grid(sticky=N)
virus_scans__left.pack(side=LEFT, expand=True, pady=5, fill=X)
# ------Middle Third
virus_scans__center = TTKFrame(tab__virus_scans, style='White.TFrame')
system_cleanup_group = TTKLabelFrame(virus_scans__center, text="Cleanup")
cc_button = TTKButton(system_cleanup_group, text="CCleaner", command=run_cc_defer)
cc_button.pack(padx=5, pady=5, fill=X)
sfc_button = TTKButton(system_cleanup_group, text="SFC", command=run_sfc_defer)
sfc_button.pack(padx=5, pady=5, fill=X)
uns_button = TTKButton(system_cleanup_group, text="Rem. Scans", command=rem_scans_defer)
uns_button.pack(padx=5, pady=5, fill=X)
system_cleanup_group.grid(row=0, sticky=N)
msconfig_group = TTKLabelFrame(virus_scans__center, text="Safe Mode")
msc_button = TTKButton(msconfig_group, text="Enable", command=msc_toggle)
msc_button.pack(padx=5, pady=5, fill=X)
msconfig_group.grid(row=1, sticky=N)
virus_scans__center.pack(side=LEFT, expand=True, anchor=N, pady=5, fill=X)
# ------Right Third
virus_scans__right = TTKFrame(tab__virus_scans, style='White.TFrame')
rootkit_group = TTKLabelFrame(virus_scans__right, text="Rootkit")
mwbar_button = TTKButton(rootkit_group, text="MWBAR", command=run_mwbar_defer, state=DISABLED)
mwbar_button.pack(padx=5, pady=5, fill=X)
tdss_button = TTKButton(rootkit_group, text="TDSS", command=run_tdss_defer, state=DISABLED)
tdss_button.pack(padx=5, pady=5, fill=X)
rootkit_group.grid(row=0)
misc_group = TTKLabelFrame(virus_scans__right, text="Misc.")
mse_button = TTKButton(misc_group, text="Inst. MSE", command=get_mse_defer)
mse_button.pack(padx=5, pady=5, fill=X)
progs_button = TTKButton(misc_group, text="Prog&Feat", command=run_pnf)
progs_button.pack(padx=5, pady=5, fill=X)
tfr_button = TTKButton(misc_group, text="Temp Files", command=run_temp_defer)
tfr_button.pack(padx=5, pady=5, fill=X)
tic_button = TTKButton(misc_group, text="Ticket", command=run_ticket)
tic_button.pack(padx=5, pady=5, fill=X)
misc_group.grid(row=1)
virus_scans__right.pack(side=LEFT, expand=True, pady=5, fill=X)
root_notebook.add(tab__virus_scans, text="Virus Scans")

# ----Tab:OS Troubleshooting
tab__os_troubleshooting = TTKFrame(style='White.TFrame')
# ------Left Third
os_troubleshooting__left = TTKFrame(tab__os_troubleshooting, style='White.TFrame')
network_group = TTKLabelFrame(os_troubleshooting__left, text="Network")
ipc_button = TTKButton(network_group, text="IPConfig", command=run_ipconfig_defer)
ipc_button.pack(padx=5, pady=5, fill=X)
wsr_button = TTKButton(network_group, text="Winsock", command=run_winsock_defer)
wsr_button.pack(padx=5, pady=5, fill=X)
had_button = TTKButton(network_group, text="Hid. Adapters", command=run_hidapters_defer)
had_button.pack(padx=5, pady=5, fill=X)
wzc_button = TTKButton(network_group, text="NIUwireless", command=run_wifi_defer)
wzc_button.pack(padx=5, pady=5, fill=X)
sts_button = TTKButton(network_group, text="Speedtest", command=run_speed)
sts_button.pack(padx=5, pady=5, fill=X)
ncp_button = TTKButton(network_group, text="Network Cpl.", command=run_netcpl)
ncp_button.pack(padx=5, pady=5, fill=X)
network_group.pack()
os_troubleshooting__left.pack(side=LEFT, expand=True, fill=X, pady=5)
# ------Middle Third
os_troubleshooting__center = TTKFrame(tab__os_troubleshooting, style='White.TFrame')
system_group = TTKLabelFrame(os_troubleshooting__center, text="System")
sfc2_button = TTKButton(system_group, text="SFC", command=run_sfc_defer)
sfc2_button.pack(padx=5, pady=5, fill=X)
dism_button = TTKButton(system_group, text="DISM", command=run_dism_defer)
dism_button.pack(padx=5, pady=5, fill=X)
aio_button = TTKButton(system_group, text="All in One", command=run_aio_defer, state=DISABLED)
aio_button.pack(padx=5, pady=5, fill=X)
dfg_button = TTKButton(system_group, text="Defragment", command=run_defrag_defer)
dfg_button.pack(padx=5, pady=5, fill=X)
ckd_button = TTKButton(system_group, text="Disk Check", command=run_chkdsk_defer)
ckd_button.pack(padx=5, pady=5, fill=X)
dmc_button = TTKButton(system_group, text="Device Mgmt.", command=run_dmc)
dmc_button.pack(padx=5, pady=5, fill=X)
system_group.pack()
os_troubleshooting__center.pack(side=LEFT, expand=True, fill=X, pady=5)
# ------Right Third
os_troubleshooting__right = TTKFrame(tab__os_troubleshooting, style='White.TFrame')
miscellaneous_group = TTKLabelFrame(os_troubleshooting__right, text="Miscellaneous")
rmse_button = TTKButton(miscellaneous_group, text="Remove MSE", command=rem_mse_defer, state=DISABLED)
rmse_button.pack(padx=5, pady=5, fill=X)
cpl_button = TTKButton(miscellaneous_group, text="Control Panel", command=run_cpl)
cpl_button.pack(padx=5, pady=5, fill=X)
reg_button = TTKButton(miscellaneous_group, text="Registry", command=run_reg)
reg_button.pack(padx=5, pady=5, fill=X)
awp_button = TTKButton(miscellaneous_group, text="Pharos", command=run_awp_defer)
awp_button.pack(padx=5, pady=5, fill=X)
miscellaneous_group.pack()
os_troubleshooting__right.pack(side=LEFT, expand=True, fill=X, pady=5)
root_notebook.add(tab__os_troubleshooting, text="OS Troubleshooting")

# ----Tab:Removal Tools
tab__removal_tools = TTKFrame(style='White.TFrame')
# ------Left Third
removal_tools__left = TTKFrame(tab__removal_tools, style='White.TFrame')
removal_tools__left.pack(side=LEFT, expand=True)
# ------Middle Third
removal_tools__center = TTKFrame(tab__removal_tools, style='White.TFrame')
removal_tools__center.pack(side=LEFT, expand=True)
# ------Right Third
removal_tools__right = TTKFrame(tab__removal_tools, style='White.TFrame')
removal_tools__right.pack(side=LEFT, expand=True)
root_notebook.add(tab__removal_tools, text="Removal Tools")

# ----Tab:Auto
tab__auto = TTKFrame(style='White.TFrame')
root_notebook.add(tab__auto, text="Auto")
root_notebook.grid(row=0)

root_empty_space = TTKFrame()
root_empty_space.grid(row=1, pady=5)
root_progressbar = Progressbar(root, mode="indeterminate")
root_progressbar.grid(row=2, sticky=E + W, padx=10)
root_progressbar.start()
root_progressbar_label = TTKLabel(root, text="Ready for adventure...")
root_progressbar_label.grid(row=3)

status_bar = ResToolStatusBar(root)  # create a status bar class object
status_bar.grid(row=4, sticky=E + W)  # put it on da board

safe_mode_tracker = SafeModeTracker()

root.protocol('WM_DELETE_WINDOW', stop)  # bind the close

'''endpoint = TCP4ServerEndpoint(reactor, 8000)


    # ----------------------------------------------------------
    # Function: listen
    # Purpose:  listen for remote control data
    # Preconds: port for listening is open
    # Postconds:listening on port
    # Arguments:
    # ----------------------------------------------------------
    def listen():
    d = endpoint.listen(ResToolWebResponderFactory())

    # ----------------------------------------------------------
    # Function: listen_failed
    # Purpose:  Catch error listening to specified port
    # Preconds: Error listening on endpoint
    # Postconds:Error message displayed, program closed
    # Arguments:
    # ----------------------------------------------------------
    def listen_failed(_):
        ctypes.windll.user32.MessageBoxW(0,
                                         u"Unable to bind ResTool to the necessary network port.\n" +
                                         "Verify only one instance of ResTool is running\n" +
                                         "and no other services are using port 8000.\n" +
                                         "Restool will now terminate.",
                                         u"Network Conflict!", 16)
        log.critical("ResTool was unable to bind to the correct network port.")
        reactor.stop()

    d.addErrback(listen_failed)


reactor.callWhenRunning(listen)'''  # This is code for the remote functionality, which I will leave for future use

# Make sure heartbeats are posted regularly
heartbeatTask = task.LoopingCall(post_heartbeat)
heartbeatTask.start(59.9)

# noinspection PyUnresolvedReferences
reactor.run()
# -------------------------------------------------- End Main Program --------------------------------------------------
