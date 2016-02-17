#                     ____              ______               __   _   __ _  __ ______    ____    ____
#                    / __ \ ___   _____/_  __/____   ____   / /  / | / /| |/ //_  __/   /_  /   / __ \
#                   / /_/ // _ \ / ___/ / /  / __ \ / __ \ / /  /  |/ / |   /  / /     __/ /   / / / /
#                  / _, _//  __/(__  ) / /  / /_/ // /_/ // /  / /|  / /   |  / /     / __/ _ / /_/ /
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
from twisted.internet.endpoints import TCP4ServerEndpoint
import requests

# General Imports
import os
import socket
import ctypes
from platform import platform

# Registry Manipulation
# import _winreg as winreg

# Process Control
import subprocess

# Globals
CURRENT_TOKEN = "2mD1O0ArdYRewnxCsjz7jR/NWOJnla98fmUCCFkxXU5bkBONiF0IekS6MY7bIZfp"
API_BASE_AUTODETECT = str("http://" + requests.get("http://restool-ip-kmorgan2.c9users.io/").text + "/api/")
API_BASE = API_BASE_AUTODETECT
globaldict = {}  # This is a horrible thing and I should not be doing it

# Check for Admin
if not ctypes.windll.shell32.IsUserAnAdmin():
    ctypes.windll.user32.MessageBoxW(0, u"Please make sure to run ResTool with Administrator rights.",
                                     u"Admin Required!", 16)
    sys.exit()


# ---------------------------------------------------- GUI Classes -----------------------------------------------------


# A class encapsulation of the ticket creation dialog
class TicketDialog:
    global globaldict

    def __init__(self, parent):
        top = self.top = Toplevel(parent)
        self.top.title("New Ticket Dialog")
        self.top.iconbitmap('icon.ico')  # icon is the iconfile

        Label(top, text="Ticket Number:").grid(row=0, column=0)
        self.num = Entry(top)
        self.num.grid(padx=5, row=0, column=1)

        b = Button(top, text="OK", command=self.ok)
        b.grid(pady=5, columnspan=2)

    def ok(self):
        globaldict['ticket_number'] = self.num.get()

        if self.num.get():
            self.top.destroy()


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
        # Display ticket in rightmost box
        self.ticket_text = StringVar(value="No Ticket")
        self.ticket_label = Label(self, textvariable=self.ticket_text, anchor=W, borderwidth=1, relief=SUNKEN)
        self.ticket_label.pack(side=LEFT, fill=X, expand=True)

    def set_version(self):
        try:
            _ = os.environ["PROGRAMFILES(X86)"]  # if there is a program files x86 dir
            bits = ' x64'
        except KeyError:  # it's like an if-else except all the logic relies on a KeyError (or lack thereof)
            bits = ' x32'
        value = 'Windows ' + str(platform()).split('-')[1] + bits  # Concatenate (platform returns 'Windows-?-...')
        self.version_text.set(value)

    def set_ip(self):
        # gets the valid hostname of this machine (stole this code)
        global globaldict
        globaldict['ip_addr'] = [l for l in (
            [ip for ip in socket.gethostbyname_ex(socket.gethostname())[2] if not ip.startswith("127.")][:1], [
                [(soc.connect(('8.8.8.8', 80)), soc.getsockname()[0], soc.close()) for soc in
                 [socket.socket(socket.AF_INET, socket.SOCK_DGRAM)]][0][1]]) if l][0][0]
        self.ip_text.set(globaldict['ip_addr'])

    def set_ticket(self):
        self.ticket_text.set("DEBUG")  # replace this


# -------------------------------------------------- END GUI Classes ---------------------------------------------------

# ---------------------------------------------------- API Classes -----------------------------------------------------
def status(text):
    return text


def action(text):
    return text


def stop():
    return 0


def close():
    return 0


# -------------------------------------------------- Callback Methods --------------------------------------------------

# Called to thread-defer a GUI callback.
def app_defer(function, name):
    default_status_area()
    root_progressbar_label.configure(text="Running " + name)
    root_progressbar.config(mode='indeterminate')
    root_progressbar.start()
    d = threads.deferToThread(function)
    status("Running " + name)
    d.addCallbacks(app_callback, app_errback, callbackArgs=(name,), errbackArgs=(name,))
    return 0


# Called on app success
def app_callback(return_text):
    # update Action
    action(return_text)
    # reset GUI
    root_progressbar_label.configure(text="Ready for adventure...")
    root_progressbar.config(mode='indeterminate')
    root_progressbar.start()


# Called on app error
def app_errback(error, name):
    # update status
    status("While running " + str(name) + ", ResTool encountered the following error:\n" + error.getErrorMessage())
    # set GUI components
    red_status_area()
    root_progressbar_label.configure(text="Error in " + name)
    root_progressbar.config(mode='determinate', value=0)
    root_progressbar.stop()


def run_cf():
    return "ComboFix performed about a million or so deletions."


def run_cf_defer():
    app_defer(run_cf, "ComboFix")


def run_mwb():
    raise UserWarning('Malwarebytes literally exploded... Good luck cleaning that up.')


def run_mwb_defer():
    app_defer(run_mwb, "Malwarebytes")


def run_eset():
    import time

    time.sleep(60)
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
    return 0


def run_ipconfig():
    status("IPConfig Reset will cause this machine to become temporarily unreachable")
    subprocess.call(['ipconfig.exe', '/release'])
    subprocess.call(['ipconfig.exe', '/flushdns'])
    subprocess.call(['ipconfig.exe', '/renew'])
    return 0


def run_ipconfig_defer():
    app_defer(run_ipconfig, "IPConfig Reset")


def run_winsock():
    status("Winsock Reset requires a restart; this machine will be unreachable until the restart has completed")
    subprocess.call(['netsh.exe', 'int', 'ip', 'reset'])
    subprocess.call(['netsh.exe', 'winsock', 'reset'])
    subprocess.call(['netsh.exe', 'winsock', 'reset', 'catalog'])
    subprocess.call(['netsh.exe', 'advfirewall', 'reset'])
    print "Reboot Necessary"
    # Set ResTool to run on restart
    # Restart the computer
    return 0


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
    return 0


def run_speed_defer():
    app_defer(run_speed, "Speed Test")


def run_netcpl():
    subprocess.call(['control.exe', '/name', 'Microsoft.NetworkAndSharingCenter'])
    return 0


def run_sfc():
    return 0


def run_sfc_defer():
    app_defer(run_sfc, "System File Checker")


def run_dism():
    return 0


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
    return 0


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


def run_reg():
    subprocess.call(['regedit.exe'])


def run_awp():
    return 0


def run_awp_defer():
    app_defer(run_awp, "Anywhere Printer Install")


# ------------------------------------------------ End Callback Methods ------------------------------------------------

# -------------------------------------------------- Network Classes ---------------------------------------------------


# A Twisted protocol class to receive communication via ResTool Web.

# noinspection PyClassHasNoInit
class ResToolWebResponder(Protocol):
    funcMap = {'A': status, 'B': run_cf, 'C': run_mwb, 'D': run_eset, 'E': run_sas, 'F': run_sb, 'G': run_hc,
               'H': run_cc, 'I': run_sfc, 'J': rem_scans, 'K': msc_toggle, 'L': run_mwbar, 'M': run_tdss, 'N': get_mse,
               'O': run_pnf, 'P': run_temp, 'Q': run_ticket, 'R': run_ipconfig, 'S': run_winsock, 'T': run_hidapters,
               'U': run_wifi, 'V': run_speed, 'W': run_netcpl, 'X': run_dism, 'Y': run_aio, 'Z': run_defrag,
               'a': run_chkdsk, 'b': run_dmc, 'c': rem_mse, 'd': run_cpl, 'e': run_reg, 'f': run_awp, 'g': stop,
               'h': close}
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

root = Tk()  # create root window
tksupport.install(root)  # bind the GUI to a twisted reactor
root.protocol('WM_DELETE_WINDOW', None)  # bind the close
root.iconbitmap('icon.ico')
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
sfc_button = Button(system_cleanup_group, text="SFC", command=run_sfc_defer, state=DISABLED)
sfc_button.pack(padx=5, pady=5, fill=X)
uns_button = Button(system_cleanup_group, text="Rem. Scans", command=rem_scans, state=DISABLED)
uns_button.pack(padx=5, pady=5, fill=X)
system_cleanup_group.grid(row=0, sticky=N)
msconfig_group = LabelFrame(virus_scans__center, text="MSConfig")
msc_button = Button(msconfig_group, text="Enable", command=msc_toggle, state=DISABLED)
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
tic_button = Button(misc_group, text="Ticket", command=run_ticket, state=DISABLED)
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
sts_button = Button(network_group, text="Speedtest", command=run_speed_defer, state=DISABLED)
sts_button.pack(padx=5, pady=5, fill=X)
ncp_button = Button(network_group, text="Network Cpl.", command=run_netcpl)
ncp_button.pack(padx=5, pady=5, fill=X)
network_group.pack()
os_troubleshooting__left.pack(side=LEFT, expand=True, fill=X, pady=5)
# ------Middle Third
os_troubleshooting__center = Frame(tab__os_troubleshooting, style='White.TFrame')
system_group = LabelFrame(os_troubleshooting__center, text="System")
sfc2_button = Button(system_group, text="SsFC", command=run_sfc_defer, state=DISABLED)
sfc2_button.pack(padx=5, pady=5, fill=X)
dism_button = Button(system_group, text="DISM", command=run_dism_defer, state=DISABLED)
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
cpl_button = Button(miscellaneous_group, text="Control Panel", command=run_cpl, state=DISABLED)
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

status_bar.set_ticket()

root.protocol('WM_DELETE_WINDOW', reactor.stop)  # bind the close

endpoint = TCP4ServerEndpoint(reactor, 8000)


def listen():
    d = endpoint.listen(ResToolWebResponderFactory())

    def listen_failed(_):
        ctypes.windll.user32.MessageBoxW(0,
                                         u"Unable to bind ResTool to the necessary network port.\n" +
                                         "Verify only one instance of ResTool is running\n" +
                                         "and no other services are using port 8000.\n" +
                                         "Restool will now terminate.",
                                         u"Network Conflict!", 16)
        reactor.stop()

    d.addErrback(listen_failed)


reactor.callWhenRunning(listen)
# noinspection PyUnresolvedReferences
reactor.run()
# -------------------------------------------------- End Main Program --------------------------------------------------
