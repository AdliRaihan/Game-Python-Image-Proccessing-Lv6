import imutils
import cv2
import pyautogui as keygg
from PIL import Image
import numpy
import keyboard
import win32gui
import win32ui
import win32con
from ctypes import windll
import win32com as comctl
import time

class startScanning:

    myArray = None
    currentImg = None
    selectedKey = "6"

    def __init__ (self, array, selectedKey):
        self.myArray = array
        hwnd = win32gui.FindWindow(None, 'Audition')

        left, top, right, bot = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bot - top
    
        wDC = win32gui.GetWindowDC(hwnd)
        dcObj=win32ui.CreateDCFromHandle(wDC)
        cDC=dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, w, h)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0,0),(w, h) , dcObj, (0,0), win32con.SRCCOPY)
        dataBitMap.SaveBitmapFile(cDC, 'test.jpg')
        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())
        self.selectedKey = selectedKey
        self.cropSelected()

    def cropSelected(self):
        ResDict = {
            "6" : (393,435),
            "7" : (282,322),
            "8" : (264,304),
            "9" : (246,386),
        }
        # Perkotak itu 35x35
        # 6 : 425:475 200:240
        # 6^(*2) : 425 + (res) | 200 + (res)
        # 7 : 35/2
        # Ambil Image
        imageTemplate = cv2.imread("test.jpg")
        (fX, fX2) = ResDict[self.selectedKey]

        for index in range(0, int(self.selectedKey)):
            if index == 0 :
                pass
            else:
                (x, x2) = ResDict[self.selectedKey]
                (fX, fX2) =  ((x + (40 * index + 1)), (x2 + (40 * index + 1))) 
                # Scan Image + trigger input keyboard
                # Lempar ke doScan def ( function )
            # 530:580
            cropImg = imageTemplate[535:575, fX:fX2].copy()
            # cv2.imshow("cropped", cropImg)
            # cv2.waitKey(0)
            self.doScan(cropImg)


    def doScan(self, imageCrop):
        imageArr = [
            'arrow_left.jpg', #kiri
            'arrow_left_up.jpg', #kiri-atas
            'arrow_top.jpg', #atas
            'arrow_top_right.jpg', # kanan atas
            'arrow_right.jpg', # kanan
            'arrow_bottom_right.jpg', # kanan bawah
            'arrow_bottom.jpg', # bawah
            'arrow_left_bottom.jpg', # kiri bawah
        ]
        if self.myArray is not None:
            for idx in imageArr:

                imageTemplate = cv2.imread(idx)
                image = imageCrop
                founded = None
                (h, w, d) = image.shape

                detection = cv2.matchTemplate(image,imageTemplate, cv2.TM_CCOEFF_NORMED)
                (min_val, max_val, minLoc, maxLoc) = cv2.minMaxLoc(detection)

                if founded is None or max_val > founded[0]:
                    founded = (max_val, maxLoc)

                (_, pos) = founded
                (x, y) = pos

                if max_val >= 0.85:
                    # print("IMAGE {}".format(idx))
                    # print("MAX VALUE : ", min_val, max_val, minLoc, maxLoc, idx)
                    # print("MEAN VALUE: ", max_val - min_val)
                    sendKey.whatKey(_, idx)
                # cv2.rectangle(image, pos, (x + 35, y + 35), (0,255,0), 1)
                # cv2.imshow('Detected', image)
                # cv2.waitKey(0)
        
        # print("================================\n")
# 8496607.0
# 5532369.0
# 5607378.0


class sendKey:


    def whatKey(self, keyLoad):

        listsKey = {
            'arrow_left.jpg': 'left',
            'arrow_left_up.jpg': 'home',
            'arrow_top.jpg': 'up',
            'arrow_top_right.jpg': 'page up',
            'arrow_bottom_right.jpg': 'page down',
            'arrow_right.jpg': 'right',
            'arrow_left_bottom.jpg' : 'end',
            'arrow_bottom.jpg':'down',
        }

        if keyLoad is not None:
            if keyLoad in listsKey:
                try:
                    # wsh = comctl.shell

                    # Google Chrome window title
                    # wsh.AppActivate("Audition")
                    # wsh.SendKeys(listsKey[keyLoad])
                    # print(listsKey[keyLoad])
                    keyboard.press(listsKey[keyLoad])
                    time.sleep(0.10)
                    keyboard.release(listsKey[keyLoad])
                    time.sleep(0.10)
                except:
                    print("KEY NOT FOUND")
                

class screenDetectionHax:
    
    def show_Menu():
        pImage = ["arrow_left.jpg", "arrow_left_up.jpg", "arrow_top.jpg"]
        always = 0
        while always == 0:
            print("Waiting for press!")
            test = keyboard.read_event()
            print(test.name)
            if test.name == '6':
                
                
                # idx = keyboard.read_event()
                # print(int(idx.name))
                # startScanning(pImage, idx.name)
                try:
                    idx = keyboard.read_event()
                    print(int(idx.name))
                    startScanning(pImage, idx.name)
                except:
                    print("what?")
            elif test.name == 'esc':
                pass
            # try:
            #     if keyboard.is_pressed('num5'):
            #         print('Pressed')
            #         break
            #     else:
            #         pass
            # except:
            #     break

        # doExit = 0
        # while doExit == 0:
        #     print("1. Start Scanning")
        #     print("0. Exit")

        #     x = input()
        #     if x != '':
        #         p = int(x)

        #     if p is not None:
        #         if p == 0:
        #             doExit = 1
        #         else :
        #             startScanning(pImage).doScan()
        #             doExit = 0

        #     else:
        #         # FORCE EXIT
        #         doExit = 1
                    

def main():
    screenDetectionHax.show_Menu()

if __name__ == "__main__":
    main()
