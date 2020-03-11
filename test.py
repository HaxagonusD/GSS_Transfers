import xlrd
import pyautogui as gui

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
        './inventory_transfer_screen_images/retrieve all.PNG', confidence=0.9)
    # print(get_button_found)
    get_button_found_left_adjustment = 50
    get_button_found_top_adjustment = 40
    get_button_found_x = get_button_found.left + get_button_found_left_adjustment
    get_button_found_y = get_button_found.top + get_button_found_top_adjustment

# this is the part number input textbox


def enterPartNumber(string):
    part_number_textbox = gui.locateOnScreen(
        './inventory_transfer_screen_images/Part Number text box.PNG')
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
    gui.write(string)
    gui.hotkey("tab")







# this ist he transfer fromm button
def enterFromBin(location):
    from_textbox = gui.locateOnScreen(
        "./inventory_transfer_screen_images/from bin textbox.PNG")
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
        "./inventory_transfer_screen_images/a to bin textbox.PNG")
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
        "./inventory_transfer_screen_images/transfer quantity.PNG")
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
        "./inventory_transfer_screen_images/transfer bar.PNG")
    button_left_adjustment = 60
    button_top_adjustment = 15
    button_x = button_left_adjustment + button.left
    button_y = button_top_adjustment + button.top
    gui.moveTo(button_x,button_y)
    gui.click()

    


for i in range(where_to_start, sheet.nrows):
    
    # get relevant information
    # I wish object destructuring was a thing python
    row = sheet.row_values(i)
    part = row[gss_part]
    part_current_location = row[coming_from]
    destination = row[going_to]
    quantity = row[qty]

    enterPartNumber(part)
    enterFromBin(part_current_location)
    enterDestinationBin(destination)
    enterTransferQuantity(quantity)
    clickOnTransferButton()


