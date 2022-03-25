from multiprocessing.connection import wait
import time
import openpyxl
from pywinauto import application, Desktop
import pywinauto.keyboard as keyboard
app1 = application.Application(backend='uia')
app = Desktop(backend='uia')
app1.start(r'C:\Program Files (x86)\VMware\VMware Horizon View Client\vmware-view.exe').connect(
    title='VMware Horizon Client', timeout=200)
time.sleep(5)
username = app.VMwareHorizonClient.child_window(
    title="User name:", auto_id="134", control_type="Edit").wrapper_object()
username.type_keys("Harsh.x.Gupta")
password = app.VMwareHorizonClient.child_window(
    title="Passcode:", auto_id="136", control_type="Edit").wrapper_object()
password.type_keys("Welcome26")
login = app.VMwareHorizonClient.child_window(
    title="Login", auto_id="1", control_type="Button").wrapper_object()
login.click_input()
time.sleep(5)
grid = app.VMwareHorizonClient.Static5.window_text()
wb = openpyxl.load_workbook('Grid.xlsx')
sheet = wb['Sheet1']
grid_value = str(sheet['{}'.format(grid[19:21])].value) + str(
    sheet['{}'.format(grid[24:26])].value) + str(sheet['{}'.format(grid[29:31])].value)
gridcode = app.VMwareHorizonClient.child_window(
    title="Next Code:", auto_id="134", control_type="Edit").wrapper_object()
gridcode.type_keys(grid_value)
login = app.VMwareHorizonClient.child_window(
    title="Login", auto_id="1", control_type="Button").wrapper_object()
login.click_input()
time.sleep(5)
keyboard.send_keys('{ENTER}')
#app.VMwareHorizonClient.print_control_identifiers()
