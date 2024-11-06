import pyautogui
import os
import pytesseract
import cv2
import numpy as np

leftalt = 1160
topalt = 205
widthalt = 50
heightalt = 26

# Exibir retângulo da área de captura na tela
screen = pyautogui.screenshot()
screen_np = np.array(screen)

# Desenhar o retângulo no local de captura
cv2.rectangle(screen_np, (leftalt, topalt), (leftalt + widthalt, topalt + heightalt), (0, 255, 0), 2)

# Mostrar a tela com o retângulo
cv2.imshow("Área de captura", cv2.cvtColor(screen_np, cv2.COLOR_BGR2RGB))
cv2.waitKey(0)
cv2.destroyAllWindows()


