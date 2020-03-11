# TODO
# make sure it ends when there are no more part numbers - done
# general error checking like if the part is not foud just go to the next
# and at the end give me the ones that were not found
# ^^^ that's kinda done
# now I need to test this on real data and jsut test this somemore
# because I'm not trying to look like I'm dumb

#TODO
#I have to do this because GSS  won't let me type
# map all the keys on the key board
# after this typing is not a problem 
# I'm pretty much gonna have to export this from another file
# maybe I can use locatecenter function from pyautogui
# the next part is figuring out how to input quantities 
# I either have to read them from the screen but that's hard 
# an easy way is to just type it in a textbox 
# We'll lets just get typing down first 



import xlrd
import pyautogui as gui
import time
# import chilimangoes
# gui.screenshot = chilimangoes.grab_screen
# gui.pyscreeze.screenshot = chilimangoes.grab_screen
# gui.size = lambda: chilimangoes.screen_size

# reading the entire file
loc = "./Inventory Transfer Sheet-2.0.xls"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# abbreviations
heading_row_number = 4
gss_part = 1
decription = 2
qty = 3
coming_from = 4
transfer_to = 5
going_to = 6
work_order = 7
picked_by = 8
received_by = 9

# Start the script
# for every row in the spreadsheet do this
where_to_start = 5


# Computer vision time
# this is the retrieve all button
def gotoRetrieveAll():
    get_button_found = gui.locateOnScreen(
        r'C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/gwenTest/retrieve all.PNG', confidence=0.9)
    # print(get_button_found)
    get_button_found_left_adjustment = 50
    get_button_found_top_adjustment = 40
    get_button_found_x = get_button_found.left + get_button_found_left_adjustment
    get_button_found_y = get_button_found.top + get_button_found_top_adjustment

# this is the part number input textbox


# the_whole_thing = gui.locateOnScreen(
#     r"C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/gwenTest/Capture.PNG",  confidence=0.9)

# gui.moveTo(the_whole_thing.left+15, the_whole_thing.top+15)
# gui.rightClick()


def enterPartNumber(string):
    part_number_textbox = gui.locateOnScreen(
        r"C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/gwenTest/Part Number text box.PNG")
    part_number_textbox_left_adjustment = 50
    part_number_textbox_top_adjustment = 25
    part_number_textbox_x = part_number_textbox.left + \
        part_number_textbox_left_adjustment
    part_number_textbox_y = part_number_textbox.top + \
        part_number_textbox_top_adjustment
    gui.moveTo(part_number_textbox_x, part_number_textbox_y)
    # gui.doubleClick()
    gui.click()
    gui.hotkey("backspace")
    # time.sleep(1)
    print("before")
    # gui.write(string)
    print("after")
    gui.hotkey("tab")
    gui.hotkey("tab")
    gui.hotkey("tab")
    gui.hotkey("tab")
    gui.hotkey("tab")
    gui.hotkey("tab")


# this ist he transfer fromm button
def enterFromBin(location):
    from_textbox = gui.locateOnScreen(
        r"C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/gwenTest/from bin textbox.PNG", confidence=0.9)
    from_textbox_left_adjustment = 75
    from_text_top_adjustment = 50
    from_textbox_x = from_textbox_left_adjustment + from_textbox.left
    from_textbox_y = from_text_top_adjustment + from_textbox.top
    gui.moveTo(from_textbox_x, from_textbox_y)
    gui.click()
    gui.hotkey("backspace")
    gui.write(location)
    gui.hotkey("tab")
# this is the transfer quantity


def enterDestinationBin(location):
    transfer_quantity_textboox = gui.locateOnScreen(
        r"C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/gwenTest/a to bin textbox.PNG", confidence=0.9)
    transfer_quantity_textboox_left_adjustment = 100
    transfer_quantity_textboox_top_adjustment = 70
    transfer_quantity_textboox_x = transfer_quantity_textboox_left_adjustment + \
        transfer_quantity_textboox.left
    transfer_quantity_textboox_y = transfer_quantity_textboox_top_adjustment + \
        transfer_quantity_textboox.top
    gui.moveTo(transfer_quantity_textboox_x, transfer_quantity_textboox_y)
    gui.click()
    gui.hotkey("backspace")
    gui.write(location)
    gui.hotkey("tab")


def enterTransferQuantity(number):
    transfer_bar = gui.locateOnScreen(
        r"C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/gwenTest/transfer quantity.PNG", confidence=0.9)
    transfer_bar_left_adjustment = 112
    transfer_bar_top_adjustment = 27
    transfer_bar_x = transfer_bar_left_adjustment + transfer_bar.left
    transfer_bar_y = transfer_bar_top_adjustment + transfer_bar.top
    gui.moveTo(transfer_bar_x, transfer_bar_y)
    gui.click()
    gui.hotkey('backspace')
    gui.write(number)
    gui.hotkey('tab')


# this is the transfer button
def clickOnTransferButton():
    button = gui.locateOnScreen(
        r"C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/gwenTest/transfer bar.PNG", confidence=0.9)
    button_left_adjustment = 60
    button_top_adjustment = 15
    button_x = button_left_adjustment + button.left
    button_y = button_top_adjustment + button.top
    gui.moveTo(button_x, button_y)
    gui.click()


def clickOnSelected():
    retrieve = gui.locateOnScreen(
        r"C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/gwenTest/all the retrieves.PNG", confidence=0.9)
    retrieve_left_adjustment = 150
    retrieve_top_adjustment = 15
    retrieve_x = retrieve_left_adjustment + retrieve.left
    retrieve_y = retrieve_top_adjustment + retrieve.top
    gui.moveTo(retrieve_x, retrieve_y)
    gui.rightClick()


gui.hotkey("control")
gui.hotkey()
for i in range(where_to_start, sheet.nrows):

    # get relevant information
    # I wish object destructuring was a thing python
    row = sheet.row_values(i)
    part = row[gss_part]
    # if(part == ""):
    #     continue
    part_current_location = row[coming_from]
    destination = row[going_to]
    quantity = row[qty]

    enterPartNumber(part)
    # enterFromBin(part_current_location)
    # enterDestinationBin(destination)
    # enterTransferQuantity(quantity)
    # clickOnTransferButton()
