import pygame
import pyautogui
pyautogui.PAUSE = 0.01
pyautogui.FAILSAFE = False
from subprocess import Popen
import win32gui
from ctypes import HRESULT
from ctypes.wintypes import HWND
import comtypes
import comtypes.client


def Tabtip():
    class ITipInvocation(comtypes.IUnknown):
        _iid_ = comtypes.GUID("{37c994e7-432b-4834-a2f7-dce1f13b834b}")
        _methods_ = [
            comtypes.COMMETHOD([], HRESULT, "Toggle",
                    ( ['in'], HWND, "hwndDesktop" )
                    )
            ]

    dtwin = win32gui.GetDesktopWindow();
    ctsdk = comtypes.client.CreateObject("{4ce576fa-83dc-4F88-951c-9d0782b4e376}", interface=ITipInvocation)
    ctsdk.Toggle(dtwin);


tabtip = Popen('C:/Program Files/Common Files/microsoft shared/ink/tabtip.exe', shell= True)


pygame.init()
done = False
 
sensitivity = 25
mode = 0
plugin = False



if plugin:
    zoom_key = 'alt'

    def zoom_func(n):
        with pyautogui.hold(zoom_key):
            pyautogui.scroll(n)

else:
    zoom_key = 'ctrl'
    def zoom_func(n):
        if n > 0:
            sign = '+'
        else:
            sign = '-'

        with pyautogui.hold(zoom_key):
            pyautogui.press(sign)


joystick_count = pygame.joystick.get_count()


my_joystick = pygame.joystick.Joystick(0)
my_joystick.init()

while not done:
 
    for event in pygame.event.get():
        if event.type == pygame.JOYBUTTONDOWN:
            if event.button == 7:
                done = True
            elif event.button == 0:
                pyautogui.mouseDown(button='right')
            elif event.button == 2:
                pyautogui.mouseDown(button='left')
            elif event.button == 4:
                if mode == 1:
                    sensitivity = 25
                    mode = 0
                else:
                    mode = 1
                    sensitivity = 10
                
            elif event.button == 1:
                # with pyautogui.hold(zoom_key):
                #     pyautogui.scroll(60)
                zoom_func(60)
            elif event.button == 3:
                # with pyautogui.hold(zoom_key):
                #     pyautogui.scroll(-60)
                zoom_func(-60)


            elif event.button == 5:
                # try:
                    Tabtip()
                # except:
                #     tabtip.kill()
                #     tabtip = Popen('C:/Program Files/Common Files/microsoft shared/ink/tabtip.exe', shell= True)

            
        elif event.type == pygame.JOYBUTTONUP:
            if event.button == 0:
                pyautogui.mouseUp(button='right')
            elif event.button == 2:
                pyautogui.mouseUp(button='left')



    horiz_move_pos = my_joystick.get_axis(0)
    vert_move_pos = my_joystick.get_axis(1)
    
    horiz_scroll_pos = my_joystick.get_axis(2)
    vert_scroll_pos = my_joystick.get_axis(3)

    x_coord = int(horiz_move_pos * sensitivity)
    y_coord = int(vert_move_pos * sensitivity)


    pyautogui.move(x_coord, y_coord)
            
    x_scroll = int(horiz_scroll_pos * sensitivity * 1.5)
    y_scroll = int(vert_scroll_pos * sensitivity * -1.5)
    pyautogui.vscroll(y_scroll)
    pyautogui.hscroll(x_scroll)


pygame.quit()