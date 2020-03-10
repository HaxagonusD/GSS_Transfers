import xlrd
import pyautogui

image_location = pyautogui.locateOnScreen('./Capture.PNG', confidence=0.8)
the_button = pyautogui.locateOnScreen('./the button.PNG', confidence=0.8)
pyautogui.moveTo(the_button.left, the_button.top)
# pyautogui.moveTo(image_location.left+330, image_location.top+90)
print(the_button)

