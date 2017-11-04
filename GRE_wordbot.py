from win32api import *
from win32gui import *
import win32con
import sys, os
import struct
import time
import sqlite3
import random

def WordList(ID):
    word = ""
    meaning = ""
    
    conn = sqlite3.connect('wordlist.db')
    print("Opener")

    cursor = conn.execute("SELECT WORDS,MEANING from WORDS where ID = %s ;" % (ID) )
    for row in cursor:
       print("WORD = ", row[0])
       word = row[0]
       print("MEANING = ", row[1])
       meaning = row[1]
    print("Operation done successfully")

    conn.close()
    
    return word, meaning

class WindowsBalloonTip:
    def __init__(self, title, msg):
        message_map = {
                win32con.WM_DESTROY: self.OnDestroy,
        }
        # Register the Window class.
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = "PythonTaskbar"
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        classAtom = RegisterClass(wc)
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( classAtom, "Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, hinst, None)
        UpdateWindow(self.hwnd)

        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, 0, "tooltip")
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, \
                         (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20,\
                          0, "Balloon  tooltip",msg,200,title))
        time.sleep(15)
        DestroyWindow(self.hwnd)
        classAtom = UnregisterClass(classAtom, hinst)
    
    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0) # Terminate the app.

def balloon_tip(title, msg):
    w=WindowsBalloonTip(title, msg)

if __name__ == '__main__':
    while 1 :
        number = random.randint(1,1500)
        word, meaning = WordList(number)
        balloon_tip(word,meaning)
        time.sleep(300)
