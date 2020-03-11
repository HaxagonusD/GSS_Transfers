import pyautogui as gui
import time

textbox = gui.locateOnScreen(
    r"C:/Users/DDGReceiving100/Desktop/Coding/python_scripts/test/GSS_Transfers/inventory_transfer_screen_images/textboox_chrome.PNG", confidence=0.9)
gui.moveTo(textbox.left+200, textbox.top+15)
gui.click()
gui.press("backspace")
gui.write("something")
gui.press("enter")
