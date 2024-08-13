import pandas as pd
import pyautogui
from time import sleep, time
import pytesseract
import cv2 #opencv
import numpy as np
import os
import keyboard

pyautogui.moveTo(303,587)#mover o mouse parar area das luminaria para fazer a rolagem para cima
pyautogui.scroll(+100000)