"""
Inspired from https://github.com/AertonMarz/music-bot-new-world and developed further, 
using opencv's screenshot recognition.
Works with New World in Windowed Mode 1920x1080.
Use F10 to stop the script. Dependencies: See Import functions.
"""
import numpy as np
import cv2 as cv
from PIL import ImageGrab, Image
import time
import win32api as api
import win32con as con

def main():

    """
    main function for the program
    """

    #Image loading
    img_A = cv.imread('./img/a.png',0)
    img_D = cv.imread('./img/d.png',0)
    img_S = cv.imread('./img/s.png',0)
    img_lrclick = cv.imread('./img/lrclick.png',0)
    img_space = cv.imread('./img/space.png',0)

    REGION = (880, 975, 988, 1130)

    time.sleep(3)
    print("*Start*")

    # Keyboard Key Stroke handling
    def press_key(key):
        api.keybd_event(key, 0, 0, 0)
        time.sleep(0.05)
        api.keybd_event(key, 0, 0x0002, 0)  # 0x0002 corresponds to KEYEVENTF_KEYUP

    # Mouse Left+Right Click handling
    def click_leftright():
        api.mouse_event(con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.01)
        api.mouse_event(con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.01)        
        api.mouse_event(con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
        time.sleep(0.01)
        api.mouse_event(con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)

    #image detection handling
    def LocateImage(image, Region=None, Precision=0.56):
        img = ImageGrab.grab(bbox=REGION)
        img_cv = cv.cvtColor(np.array(img), cv.COLOR_RGB2BGR)

        img_rgb = np.array(img_cv)
        img_gray = cv.cvtColor(img_rgb, cv.COLOR_BGR2GRAY)
        #template = cv.imread(image, 0)

        res = cv.matchTemplate(img_gray, image, cv.TM_CCOEFF_NORMED)
        min_val, LocatedPrecision, min_loc, Position = cv.minMaxLoc(res)
        if LocatedPrecision > Precision:
            return Position[0], Position[1]

    last_e_press_time = time.time()

    while not api.GetAsyncKeyState(con.VK_F10) & 0x8000:
        current_time = time.time()

        #locate images and press the key accordingly - key's are expressed in ASCII
        if LocateImage(img_A) != None:
            press_key(65)
            print("A Found")

        if LocateImage(img_D) != None:
            press_key(68)
            print("D Found")

        if LocateImage(img_S) != None:
            press_key(83)
            print("S Found")

        if LocateImage(img_lrclick) != None:
            click_leftright()
            print("LR Found")
        
        if LocateImage(img_space) != None:
            press_key(32)
            print("Space Found")
        
        # Restart playing music
        if current_time - last_e_press_time >= 42:
            api.keybd_event(69, 0, 0, 0)
            time.sleep(3)
            api.keybd_event(69, 0, 0x0002, 0)  # 0x0002 corresponds to KEYEVENTF_KEYUP
            print("Restarting Music")
            api.keybd_event(69, 0, 0, 0)
            time.sleep(.5)
            api.keybd_event(69, 0, 0x0002, 0)  # 0x0002 corresponds to KEYEVENTF_KEYUP
            last_e_press_time = time.time()
    

# Runs the main function
if __name__ == '__main__':
    main()