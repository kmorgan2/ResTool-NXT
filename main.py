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
from ttk import *

# Network/Web Imports
from twisted.internet import tksupport, reactor, threads
from twisted.internet.protocol import Protocol, Factory
# from twisted.internet.endpoints import TCP4ServerEndpoint
from twisted.logger import Logger, textFileLogObserver
import requests
import webbrowser

# General Imports
import os
import socket
import ctypes
from platform import platform

# Registry and File Manipulation
import _winreg as reg
import io

# Process Control
import subprocess

# THE ALMIGHTY DEBUG FLAG. SET THIS TO TRUE WITH GREAT CAUTION
DEBUG = True
# THE LESS ALMIGHTY NO_POST FLAG. SET THIS TO TRUE UNTIL WE GET OUR API ENDPOINTS AND DEV ACCESS
NO_POST = True

# Logging
# noinspection PyTypeChecker
log = Logger(observer=textFileLogObserver(io.open(os.path.expanduser("~\Desktop\log.txt"), 'a')))
log.info("Starting ResTool")
# Globals
API_HOST = 'intranet'
if DEBUG:
    API_HOST = 'dev'
API_BASE = "http://" + API_HOST + ".restech.niu.edu/tickets/api/restool"
API_HEARTBEATS = API_BASE + "/heartbeats"
API_EVENTS = API_BASE + "/events"
API_TOKENS = API_BASE + "/token"

EXPIRE_SECONDS = 30

heartbeat = "Idle"

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
root.iconbitmap('icon.ico')  # icon is icon
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
def red_status_area():
    root_empty_space.configure(style="Red.TFrame")
    root_progressbar_label.configure(style="Red.TLabel")
    root.config(background='red')


def default_status_area():
    root_empty_space.configure(style="TFrame")
    root_progressbar_label.configure(style="TLabel")
    root.configure(background='SystemButtonFace')  # Discovered this is the default background color


# ---------------------------------------------------- GUI Classes -----------------------------------------------------

# A class encapsulation of the token dialog
class TokenDialog:
    def __init__(self, parent):
        self.token = None
        top = self.top = Toplevel(parent)
        self.top.title("New Ticket Dialog")
        self.top.iconbitmap('icon.ico')  # icon is icon

        Label(top, text="Token:").grid(row=0, column=0)
        self.num = Entry(top)
        self.num.grid(padx=5, row=0, column=1)

        b = Button(top, text="OK", command=self.ok)
        b.grid(pady=5, columnspan=2)

    def ok(self):
        if self.num.get():
            self.token = self.num.get()
            self.top.destroy()

    def get_token(self):
        return self.token


# A class encapsulation of the status bar


# noinspection PyCallByClass,PyTypeChecker
class ResToolStatusBar(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
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
        self.token_text = StringVar(value="No Ticket")
        self.token_label = Label(self, textvariable=self.token_text, anchor=W, borderwidth=1, relief=SUNKEN)
        self.token_label.pack(side=LEFT, fill=X, expand=True)
        self.set_ticket()

    def set_version(self):
        try:
            _ = os.environ["PROGRAMFILES(X86)"]  # if there is a program files x86 dir
            bits = ' x64'
        except KeyError:  # it's like an if-else except all the logic relies on a KeyError (or lack thereof)
            bits = ' x32'
        value = 'Windows ' + str(platform()).split('-')[1] + bits  # Concatenate (platform returns 'Windows-?-...')
        self.version_text.set(value)

    def set_ip(self):
        # gets the valid hostname of this machine (borrowed this code from stack overflow)
        ip_address = [l for l in (
            [ip for ip in socket.gethostbyname_ex(socket.gethostname())[2] if not ip.startswith("127.")][:1], [
                [(soc.connect(('8.8.8.8', 80)), soc.getsockname()[0], soc.close()) for soc in
                 [socket.socket(socket.AF_INET, socket.SOCK_DGRAM)]][0][1]]) if l][0][0]
        self.ip_text.set(ip_address)

    def set_ticket(self):
        self.token_text.set(get_ticket())


# ---------------------------------------------------- Token Class -----------------------------------------------------
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
    def __init__(self, token=None):
        self.token = token
        if not self.token:  # try to get it from the registry
            self.get_token()
            if self.token:
                log.info("Retrieved stored token from registry.")
        if NO_POST and not self.token:  # it's not in the registry, but we are offline
            self.set_token("Offline")
        elif not self.token:  # try to get it from the user
            td = TokenDialog(root)
            root.wait_window(td.top)
            self.set_token(td.get_token())
            log.info("Token {token} recorded from user", token=self.token)

    def get_token(self):
        if not self.token:
            try:
                self.token = str(reg.QueryValueEx(open_registry(), "token")[0])
            except (WindowsError, KeyError):  # The Token Value in the ResTech does not exist. Return placeholder text.
                self.token = None
                log.error("Attempt to get token when none was initialized or stored.")
        return self.token

    def set_token(self, value):
        self.token = value
        try:
            reg.SetValueEx(open_registry(True), "token", 0, reg.REG_SZ, value)
            log.info("Recorded new token {token}.", token=self.token)
        except (WindowsError, KeyError):  # This should be impossible. If it happens, something has gone wrong
            ctypes.windll.user32.MessageBoxW(0, u'Unrecoverable error in writing token to registry.', u'Registry Error',
                                             16)
            log.failure("Registry not in valid state when storing token.")

    def delete_token(self):
        self.token = None
        try:
            reg.DeleteKey(reg.HKEY_LOCAL_MACHINE, "Software\\ResTech\\")
            log.info("Removed recorded token.")
        except (WindowsError, KeyError):  # This should be impossible. If it happens, something has gone wrong
            ctypes.windll.user32.MessageBoxW(0, u'Unrecoverable error in removing ResTech registry data.',
                                             u'Registry Error', 16)
            log.failure("Registry not in valid state when removing token.")


token_manager = Token()


# -------------------------------------------------- Safe Mode Class ---------------------------------------------------
class SafeModeTracker:
    def __init__(self):
        self.safe_mode_set = False
        self.get_safe_mode_setting()
        log.info("On launch, safe mode was {value}.", value="enabled" if self.safe_mode_set else "disabled")

    def get_safe_mode_setting(self):
        output = subprocess.check_output(["bcdedit", "/enum", "{default}"], shell=True)
        if "safeboot" in output:
            self.safe_mode_set = True
            # Set GUI to reflect safe mode setting
            msc_button.configure(text="Disable")
        else:
            self.safe_mode_set = False
            msc_button.configure(text="Enable")

    def enable_safe_mode(self):
        # Modify the BCD Configuration to enable safe boot with networking
        subprocess.Popen(["bcdedit", "/set", "{default}", "safeboot", "network"], shell=True).wait()
        self.get_safe_mode_setting()
        if self.safe_mode_set:
            log.info("Safe mode has been enabled.")
        else:
            log.error("There was an error enabling safe mode.")

    def disable_safe_mode(self):
        # Modify the BCD Configuration to disable safeboot
        subprocess.Popen(["bcdedit", "/deletevalue", "{default}", "safeboot"], shell=True).wait()
        self.get_safe_mode_setting()
        if not self.safe_mode_set:
            log.info("Safe mode has been disabled.")
        else:
            log.error("There was an error disabling safe mode.")

    def toggle_safe_mode(self):
        if self.safe_mode_set:
            self.disable_safe_mode()
        else:
            self.enable_safe_mode()


# --------------------------------------------------- API Functions ----------------------------------------------------


def update_heartbeat(text):
    global heartbeat
    heartbeat = text
    post_heartbeat()


def post_heartbeat():
    global heartbeat, API_HEARTBEATS
    if not NO_POST and token_manager.get_token():
        response = requests.post(API_HEARTBEATS,
                                 {'token': token_manager.get_token(), 'status': heartbeat,
                                  'expire_seconds': EXPIRE_SECONDS})
        try:
            assert response.status_code == 200
        except AssertionError:  # Error posting the heartbeat
            if response.status_code == 401:  # Bad token
                log.error('Invalid Token {token} when posting heartbeat.', token=token_manager.get_token())
            elif response.status_code == 400:  # Bad status or expire_seconds
                log.error('Invalid data when posting heartbeat "{heartbeat}"', heartbea=heartbeat)
            else:  # Uncaught (500, 404, etc)
                log.error('Uncaught HTTP Error {rc}. token in POST heartbeats: with token: {t}, heartbeat:"{h}",' +
                          'and expire_seconds: {e}', t=token_manager.get_token(), h=heartbeat, e=EXPIRE_SECONDS)
        return response.status_code
    return "Offline"


def post_event(text):
    if not NO_POST and token_manager.get_token():
        response = requests.post(API_EVENTS, {'token': token_manager.get_token(), 'text': text})
        try:
            assert response.status_code == 200
        except AssertionError:  # Error posting the event
            if response.status_code == 401:  # Bad token
                log.error('Invalid Token {token} when posting event.', token=token_manager.get_token())
            elif response.status_code == 400:  # Bad text
                log.error('Invalid data when posting event "{text}"', text=text)
            else:  # Uncaught (500, 404, etc)
                log.error('Uncaught HTTP Error {rc} in POST events: with token:{token} and event: "{text}"',
                          token=token_manager.get_token(),
                          text=text)
        return response.status_code
    return "Offline"


def stop():
    update_heartbeat("ResTool was closed at the request of the user.")
    log.info("ResTool was closed at the request of the user.")
    reactor.stop()
    return 0


def invalidate_token():
    if not NO_POST:
        response = requests.post(API_EVENTS + "/" + token_manager.get_token(), {'action': 'invalidate'})
        try:
            assert response.status_code == 200
        except AssertionError:  # Error posting the event
            if response.status_code == 401:  # Bad token
                log.error('Invalid Token {token} when invalidating token.', token=token_manager.get_token())
            elif response.status_code == 400:  # Bad action
                log.error('Invalid action "invalidate" when invalidating token.')
            else:  # Uncaught (500, 404, etc)
                log.error('Uncaught HTTP Error {rc} in POST token: {token} with action: invalidate',
                          token=token_manager.get_token())
        return response.status_code
    token_manager.delete_token()
    return "Offline"


def get_ticket():
    if not NO_POST:
        response = requests.get(API_EVENTS + "/" + token_manager.get_token())
        try:
            assert response.status_code == 200
            ticket = response.json()['ticket']
            return ticket
        except AssertionError:
            if response.status_code == 404:
                log.error("Invalid Token {token} when getting ticket.")
            else:
                log.error("Uncaught HTTP Error {rc} in GET token: {token}")
            return "No Ticket"
    return "Offline"


# -------------------------------------------------- Callback Methods --------------------------------------------------


# Called to thread-defer a GUI callback.
def app_defer(function, name):
    log.info("Running {name}", name=name)
    default_status_area()
    root_progressbar_label.configure(text="Running " + name)
    root_progressbar.config(mode='indeterminate')
    root_progressbar.start()
    d = threads.deferToThread(function)
    update_heartbeat("Running " + name)
    d.addCallbacks(app_callback, app_errback, errbackArgs=(name,))
    return 0


# Called on app success
def app_callback(return_text):
    if return_text != 0:
        log.info("Finished. " + return_text)
        # update Action
        post_event(return_text)
        update_heartbeat("Idle")
    # reset GUI
    root_progressbar_label.configure(text="Ready for adventure...")
    root_progressbar.config(mode='indeterminate')
    root_progressbar.start()


# Called on app error
def app_errback(error, name):
    log.error("Error running {name}\n" + error.getErrorMessage(), name=name)
    # update status
    post_event("While running " + str(name) + ", ResTool encountered an error:\n" + error.getErrorMessage())
    update_heartbeat("Error running " + str(name))
    # set GUI components
    red_status_area()
    root_progressbar_label.configure(text="Error in " + name)
    root_progressbar.config(mode='determinate', value=0)
    root_progressbar.stop()


def run_cf():
    return 0


def run_cf_defer():
    app_defer(run_cf, "ComboFix")


def run_mwb():
    return 0


def run_mwb_defer():
    app_defer(run_mwb, "Malwarebytes")


def run_eset():
    return 0


def run_eset_defer():
    app_defer(run_eset, "ESET")


def run_sas():
    return 0


def run_sas_defer():
    app_defer(run_sas, "SuperAntiSpyware")


def run_sb():
    return 0


def run_sb_defer():
    app_defer(run_sb, "Spybot")


def run_hc():
    return 0


def run_hc_defer():
    app_defer(run_hc, "TrendMicro HouseCall")


def run_cc():
    return 0


def run_cc_defer():
    app_defer(run_cc, "CCleaner")


def rem_scans():
    return 0


def msc_toggle():
    safe_mode_tracker.toggle_safe_mode()
    return 0


def run_mwbar():
    return 0


def run_mwbar_defer():
    app_defer(run_mwbar, "Malwarebytes AntiRootkit")


def run_tdss():
    return 0


def run_tdss_defer():
    app_defer(run_tdss, "Kaspersky TDSSKiller")


def get_mse():
    return 0


def get_mse_defer():
    app_defer(get_mse, "Windows Defender Install")


def run_pnf():
    subprocess.call(['control.exe', 'appwiz.cpl'])
    return 0


def run_temp():
    subprocess.call(['cleanmgr.exe'])
    return 0


def run_temp_defer():
    app_defer(run_temp, "Temporary File Cleanup")


def run_ticket():
    ticket = get_ticket()
    if ticket not in ["Offline", "No Ticket"]:
        webbrowser.open("http://" + API_HOST + ".restech.niu.edu/tickets/" + ticket)
    else:
        webbrowser.open("http://" + API_HOST + ".restech.niu.edu/tickets")
    return 0


def run_ipconfig():
    update_heartbeat("IPConfig Reset will cause this machine to become temporarily unreachable")
    subprocess.call(['ipconfig.exe', '/release'])
    subprocess.call(['ipconfig.exe', '/flushdns'])
    subprocess.call(['ipconfig.exe', '/renew'])
    return "IPConfig Complete."


def run_ipconfig_defer():
    app_defer(run_ipconfig, "IPConfig Reset")


def run_winsock():
    update_heartbeat("Winsock Reset requires a restart; this machine may be unreachable until the restart is complete.")
    subprocess.call(['netsh.exe', 'int', 'ip', 'reset'])
    subprocess.call(['netsh.exe', 'winsock', 'reset'])
    subprocess.call(['netsh.exe', 'winsock', 'reset', 'catalog'])
    subprocess.call(['netsh.exe', 'advfirewall', 'reset'])
    print "Reboot Necessary."
    # Set ResTool to run on restart
    # Restart the computer
    return "Winsock Reset complete. Reboot Necessary."


def run_winsock_defer():
    app_defer(run_winsock, "Winsock Reset")


def run_hidapters():
    return 0


def run_hidapters_defer():
    app_defer(run_hidapters, "Hidden Adapter Removal")


def run_wifi():
    return 0


def run_wifi_defer():
    app_defer(run_wifi, "WiFi Profile Install")


def run_speed():
    webbrowser.open("http://speedtest.niu.edu")
    return 0


def run_netcpl():
    subprocess.call(['control.exe', '/name', 'Microsoft.NetworkAndSharingCenter'])
    return 0


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


def run_sfc_defer():
    app_defer(run_sfc, "System File Checker")


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


def run_dism_defer():
    app_defer(run_dism, "DISM")


def run_aio():
    return 0


def run_aio_defer():
    app_defer(run_aio, "All-In-One Repair")


def run_defrag():
    return 0


def run_defrag_defer():
    app_defer(run_defrag, "Disk Defragmenter")


def run_chkdsk():
    try:
        output = subprocess.check_output("chkdsk")
        output = ''.join(ch for ch in output if ch not in ['x00', 'x08'])
        if "found no problems." in output:
            return "CHKDSK found no errors"
        else:
            subprocess.call("fsutil set dirty")
            return "CHKDSK found errors. They will be repaired at next reboot"
    except subprocess.CalledProcessError as e:
        raise UserWarning("Uncaught CHKDSK Behavior: \n" + ''.join(ch for ch in e.output if ch not in ['x00', 'x08']))


def run_chkdsk_defer():
    app_defer(run_chkdsk, "Disk Check")


def run_dmc():
    subprocess.call(['mmc.exe', 'devmgmt.msc'])


def rem_mse():
    return 0


def rem_mse_defer():
    app_defer(rem_mse, "Windows Defender Uninstall")


def run_cpl():
    subprocess.call(['control.exe'])
    return 0


def run_reg():
    subprocess.call(['regedit.exe'])
    return 0


def run_awp():
    return 0


def run_awp_defer():
    app_defer(run_awp, "Anywhere Printer Install")


# ------------------------------------------------ End Callback Methods ------------------------------------------------

# -------------------------------------------------- Network Classes ---------------------------------------------------


# A Twisted protocol class to receive communication via ResTool Web.

# noinspection PyClassHasNoInit
class ResToolWebResponder(Protocol):
    funcMap = {'A': update_heartbeat, 'B': run_cf, 'C': run_mwb, 'D': run_eset, 'E': run_sas, 'F': run_sb, 'G': run_hc,
               'H': run_cc, 'I': run_sfc, 'J': rem_scans, 'K': msc_toggle, 'L': run_mwbar, 'M': run_tdss, 'N': get_mse,
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

    def dataReceived(self, data):
        self.transport.write("0")
        if data[0] not in ["A", "J", "K", "O", "Q", "W", "b", "d", "e", "g", "h"]:  # These fxns are not action-worthy
            app_defer(self.funcMap[data[0]], self.nameMap[data[0]])
        else:
            self.funcMap[data[0]]()


# noinspection PyClassHasNoInit
class ResToolWebResponderFactory(Factory):
    def buildProtocol(self, addr):
        return ResToolWebResponder()


# ---------------------------------------------------- Main Program ----------------------------------------------------

# GUI creation (main program initialization)

# --Main Notebook
root_notebook = Notebook(root)

# ----Tab: Virus Scans
tab__virus_scans = Frame(style='White.TFrame')
# ------Left Third
virus_scans__left = Frame(tab__virus_scans, style='White.TFrame')
av_scanners_group = LabelFrame(virus_scans__left, text="AV Scans")
cf_button = Button(av_scanners_group, text="ComboFix", command=run_cf_defer, state=DISABLED)
cf_button.pack(padx=5, pady=5, fill=X)
mwb_button = Button(av_scanners_group, text="Malwarebytes", command=run_mwb_defer, state=DISABLED)
mwb_button.pack(padx=5, pady=5, fill=X)
eset_button = Button(av_scanners_group, text="ESET", command=run_eset_defer, state=DISABLED)
eset_button.pack(padx=5, pady=5, fill=X)
sas_button = Button(av_scanners_group, text="SAS", command=run_sas_defer, state=DISABLED)
sas_button.pack(padx=5, pady=5, fill=X)
sb_button = Button(av_scanners_group, text="Spybot", command=run_sb_defer, state=DISABLED)
sb_button.pack(padx=5, pady=5, fill=X)
hc_button = Button(av_scanners_group, text="HouseCall", command=run_hc_defer, state=DISABLED)
hc_button.pack(padx=5, pady=5, fill=X)
av_scanners_group.grid(sticky=N)
virus_scans__left.pack(side=LEFT, expand=True, pady=5, fill=X)
# ------Middle Third
virus_scans__center = Frame(tab__virus_scans, style='White.TFrame')
system_cleanup_group = LabelFrame(virus_scans__center, text="Cleanup")
cc_button = Button(system_cleanup_group, text="CCleaner", command=run_cc_defer, state=DISABLED)
cc_button.pack(padx=5, pady=5, fill=X)
sfc_button = Button(system_cleanup_group, text="SFC", command=run_sfc_defer)
sfc_button.pack(padx=5, pady=5, fill=X)
uns_button = Button(system_cleanup_group, text="Rem. Scans", command=rem_scans, state=DISABLED)
uns_button.pack(padx=5, pady=5, fill=X)
system_cleanup_group.grid(row=0, sticky=N)
msconfig_group = LabelFrame(virus_scans__center, text="Safe Mode")
msc_button = Button(msconfig_group, text="Enable", command=msc_toggle)
msc_button.pack(padx=5, pady=5, fill=X)
msconfig_group.grid(row=1, sticky=N)
virus_scans__center.pack(side=LEFT, expand=True, anchor=N, pady=5, fill=X)
# ------Right Third
virus_scans__right = Frame(tab__virus_scans, style='White.TFrame')
rootkit_group = LabelFrame(virus_scans__right, text="Rootkit")
mwbar_button = Button(rootkit_group, text="MWBAR", command=run_mwbar_defer, state=DISABLED)
mwbar_button.pack(padx=5, pady=5, fill=X)
tdss_button = Button(rootkit_group, text="TDSS", command=run_tdss_defer, state=DISABLED)
tdss_button.pack(padx=5, pady=5, fill=X)
rootkit_group.grid(row=0)
misc_group = LabelFrame(virus_scans__right, text="Misc.")
mse_button = Button(misc_group, text="Inst. MSE", command=get_mse_defer, state=DISABLED)
mse_button.pack(padx=5, pady=5, fill=X)
progs_button = Button(misc_group, text="Prog&Feat", command=run_pnf)
progs_button.pack(padx=5, pady=5, fill=X)
tfr_button = Button(misc_group, text="Temp Files", command=run_temp_defer)
tfr_button.pack(padx=5, pady=5, fill=X)
tic_button = Button(misc_group, text="Ticket", command=run_ticket)
tic_button.pack(padx=5, pady=5, fill=X)
misc_group.grid(row=1)
virus_scans__right.pack(side=LEFT, expand=True, pady=5, fill=X)
root_notebook.add(tab__virus_scans, text="Virus Scans")

# ----Tab:OS Troubleshooting
tab__os_troubleshooting = Frame(style='White.TFrame')
# ------Left Third
os_troubleshooting__left = Frame(tab__os_troubleshooting, style='White.TFrame')
network_group = LabelFrame(os_troubleshooting__left, text="Network")
ipc_button = Button(network_group, text="IPConfig", command=run_ipconfig_defer)
ipc_button.pack(padx=5, pady=5, fill=X)
wsr_button = Button(network_group, text="Winsock", command=run_winsock_defer)
wsr_button.pack(padx=5, pady=5, fill=X)
had_button = Button(network_group, text="Hid. Adapters", command=run_hidapters_defer, state=DISABLED)
had_button.pack(padx=5, pady=5, fill=X)
wzc_button = Button(network_group, text="NIUwireless", command=run_wifi_defer, state=DISABLED)
wzc_button.pack(padx=5, pady=5, fill=X)
sts_button = Button(network_group, text="Speedtest", command=run_speed)
sts_button.pack(padx=5, pady=5, fill=X)
ncp_button = Button(network_group, text="Network Cpl.", command=run_netcpl)
ncp_button.pack(padx=5, pady=5, fill=X)
network_group.pack()
os_troubleshooting__left.pack(side=LEFT, expand=True, fill=X, pady=5)
# ------Middle Third
os_troubleshooting__center = Frame(tab__os_troubleshooting, style='White.TFrame')
system_group = LabelFrame(os_troubleshooting__center, text="System")
sfc2_button = Button(system_group, text="SFC", command=run_sfc_defer)
sfc2_button.pack(padx=5, pady=5, fill=X)
dism_button = Button(system_group, text="DISM", command=run_dism_defer)
dism_button.pack(padx=5, pady=5, fill=X)
aio_button = Button(system_group, text="All in One", command=run_aio_defer, state=DISABLED)
aio_button.pack(padx=5, pady=5, fill=X)
dfg_button = Button(system_group, text="Defragment", command=run_defrag_defer, state=DISABLED)
dfg_button.pack(padx=5, pady=5, fill=X)
ckd_button = Button(system_group, text="Disk Check", command=run_chkdsk_defer, state=DISABLED)
ckd_button.pack(padx=5, pady=5, fill=X)
dmc_button = Button(system_group, text="Device Mgmt.", command=run_dmc)
dmc_button.pack(padx=5, pady=5, fill=X)
system_group.pack()
os_troubleshooting__center.pack(side=LEFT, expand=True, fill=X, pady=5)
# ------Right Third
os_troubleshooting__right = Frame(tab__os_troubleshooting, style='White.TFrame')
miscellaneous_group = LabelFrame(os_troubleshooting__right, text="Miscellaneous")
rmse_button = Button(miscellaneous_group, text="Remove MSE", command=rem_mse_defer, state=DISABLED)
rmse_button.pack(padx=5, pady=5, fill=X)
cpl_button = Button(miscellaneous_group, text="Control Panel", command=run_cpl)
cpl_button.pack(padx=5, pady=5, fill=X)
reg_button = Button(miscellaneous_group, text="Registry", command=run_reg)
reg_button.pack(padx=5, pady=5, fill=X)
awp_button = Button(miscellaneous_group, text="Pharos", command=run_awp_defer, state=DISABLED)
awp_button.pack(padx=5, pady=5, fill=X)
miscellaneous_group.pack()
os_troubleshooting__right.pack(side=LEFT, expand=True, fill=X, pady=5)
root_notebook.add(tab__os_troubleshooting, text="OS Troubleshooting")

# ----Tab:Removal Tools
tab__removal_tools = Frame(style='White.TFrame')
# ------Left Third
removal_tools__left = Frame(tab__removal_tools, style='White.TFrame')
removal_tools__left.pack(side=LEFT, expand=True)
# ------Middle Third
removal_tools__center = Frame(tab__removal_tools, style='White.TFrame')
removal_tools__center.pack(side=LEFT, expand=True)
# ------Right Third
removal_tools__right = Frame(tab__removal_tools, style='White.TFrame')
removal_tools__right.pack(side=LEFT, expand=True)
root_notebook.add(tab__removal_tools, text="Removal Tools")

# ----Tab:Auto
tab__auto = Frame(style='White.TFrame')
root_notebook.add(tab__auto, text="Auto")
root_notebook.grid(row=0)

root_empty_space = Frame()
root_empty_space.grid(row=1, pady=5)
root_progressbar = Progressbar(root, mode="indeterminate")
root_progressbar.grid(row=2, sticky=E + W, padx=10)
root_progressbar.start()
root_progressbar_label = Label(root, text="Ready for adventure...")
root_progressbar_label.grid(row=3)

status_bar = ResToolStatusBar(root)  # create a status bar class object
status_bar.grid(row=4, sticky=E + W)  # put it on da board

safe_mode_tracker = SafeModeTracker()

root.protocol('WM_DELETE_WINDOW', stop)  # bind the close

'''endpoint = TCP4ServerEndpoint(reactor, 8000)


def listen():
    d = endpoint.listen(ResToolWebResponderFactory())

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
# noinspection PyUnresolvedReferences
reactor.run()
# -------------------------------------------------- End Main Program --------------------------------------------------