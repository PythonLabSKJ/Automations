import pyperclip
import openpyxl
import pywinauto.mouse as mouse
import pywinauto.keyboard as keyboard
import time
from pywinauto import application

app = application.Application(backend='uia')
# calling Cisco AnyConnect Application
app.start(r'C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe').connect(
    found_index=0, title='Cisco AnyConnect Secure Mobility Client', timeout=200)

# Enter the Quest VPN server
server = app.CiscoAnyConnectSecureMobilityClient.child_window(
    title="Ready to connect.", auto_id="1001", control_type="Edit").wrapper_object()
server.type_keys("Quest Employee - QDC", with_spaces=True)
conect = app.CiscoAnyConnectSecureMobilityClient.child_window(
    title="Connect", auto_id="1313", control_type="Button").wrapper_object()
conect.click_input()
time.sleep(10)

# Enter Quest Credentails
username = app.CiscoAnyConnectSecureMobilityClient.child_window(
    title="Username:", control_type="Edit").wrapper_object()
username.type_keys('^a{BACKSPACE}')
username.type_keys("Harsh.x.Gupta")
passwrd = app.CiscoAnyConnectSecureMobilityClient.child_window(
    title="Password:", control_type="Edit").wrapper_object()
passwrd.type_keys("Welcome26")
login = app.CiscoAnyConnectSecureMobilityClient.child_window(
    title="OK", control_type="Button").wrapper_object()
login.click_input()
time.sleep(5)

# Enter Grid Details
app.CiscoAnyConnectSecureMobilityClient.set_focus()

# app.CiscoAnyConnectSecureMobilityClient.print_control_identifiers()
# move the mouse and select the Grid details
mouse.right_click(coords=(555, 388))
keyboard.send_keys('+{a}')

# Copy the select Grid details
mouse.right_click(coords=(555, 388))
keyboard.send_keys('^{c}')

# store the grid details in variable
s = pyperclip.paste()
# print(s[19:21] + " " + s[24:26] + " "+s[29:31])
# wb = openpyxl.load_workbook(r'C:\Users\1035141\Documents\Python_Pratice\Grid.xlsx')

# fetching the Grid details from the Excel file
# open the excel sheet
wb = openpyxl.load_workbook('Grid.xlsx')
sheet = wb['Sheet1']

# take the grid code and find the value accordingly
grid_value = str(sheet['{}'.format(s[19:21])].value) + str(
    sheet['{}'.format(s[24:26])].value) + str(sheet['{}'.format(s[29:31])].value)
# print(grid_value)

#Enter the grid code value
grid = app.CiscoAnyConnectSecureMobilityClient.child_window(
    title="Answer:", control_type="Edit").wrapper_object()
grid.type_keys(grid_value)

# click on Contine for establish the connection
contine = app.CiscoAnyConnectSecureMobilityClient.child_window(
    title="Continue", control_type="Button", found_index=0).wrapper_object()
contine.click_input()
# time.sleep(5)
