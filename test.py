import pandas as pd
import pyautogui
from time import sleep, time
import pytesseract
import cv2 #opencv
import numpy as np
import os
import keyboard

#modificando nome do projeto
sleep(1.5)
nome_projet = pyautogui.locateCenterOnScreen('nome_projeto.png', confidence=0.6)
pyautogui.click(nome_projet.x, nome_projet.y)
sleep(1)
digitar_nome = pyautogui.locateCenterOnScreen('descri.png', confidence=0.6)
pyautogui.click(digitar_nome.x, digitar_nome.y)
sleep(0.4)
pyautogui.hotkey('shift', 'tab')

pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')