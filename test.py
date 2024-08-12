import pandas as pd
import pyautogui
from time import sleep, time
import pytesseract
import cv2 #opencv
import numpy as np
import os
import keyboard

#click area preta para fazer a rolagem
x_wall= 1410
y_wall = 567
pyautogui.click(x_wall, y_wall)

pyautogui.scroll(+800)
sleep(1.5)