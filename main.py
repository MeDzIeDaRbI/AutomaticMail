# -*- coding: utf-8 -*-

import pyautogui as pg
import time,os
import pyperclip
import win32com.client

absPath = os.path.dirname(__file__)
def openMail():
    global absPath

    location = pg.locateOnScreen(absPath+'/pics/windows.png')
    x, y = pg.center(location)
    pg.click(x,y)
    #
    time.sleep(2)
    #
    location = pg.locateOnScreen(absPath + '/pics/openMail2.png')
    x, y = pg.center(location)
    pg.click(x, y)
    #
    time.sleep(5)
    #
    location = pg.locateOnScreen(absPath + '/pics/SupComMail.png')
    x, y = pg.center(location)
    pg.click(x, y)
    #
    time.sleep(5)
    #
    location = pg.locateOnScreen(absPath + '/pics/newMessage.png')
    x, y = pg.center(location)
    pg.click(x, y)
    #
    time.sleep(5)
    #
    location = pg.locateOnScreen(absPath + '/pics/a.png')
    x, y = pg.center(location)
    pg.click(x, y)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('mail')
    #
    time.sleep(5)
    #
    pg.click(x, y+35)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('Demande')
    #
    time.sleep(5)
    #
    pg.click(x, y + 80)
    shell = win32com.client.Dispatch("WScript.Shell")
    msg = generateMail('Mr', 'Kurby')
    for i in msg.split('\n'):
        shell.SendKeys(i)
        pg.press('enter')

    #
    time.sleep(2)
    #
    location = pg.locateOnScreen(absPath + r'/pics/send.png')
    x, y = pg.center(location)
    pg.click(x, y)
    #



def generateMail(abr, name):
    return '''Dear {0}. {1},
I am a sophomohjkhjkhjhkjhjkhjjhkjhjhjkhkjhjkhjhkjhkj
Med Zied Arbi
medzied.arbi@supcom.tn
(+216) 20 275 899'''.format(abr,name)


openMail()
