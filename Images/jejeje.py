import ctypes
import datetime
import os
import sys
import time

import cv2
import keyboard
import numpy as np
import pandas as pd
import psutil
import pyautogui as pa
import pyperclip
from openpyxl import load_workbook

#cordenadasequis= x=1128, y=569
#cordenadasequis= x=1161, y=592

#im = pa.screenshot("sumasuma.png",region=(1128,569, 33, 23))


pa.hotkey("alt","tab")
#im = pa.screenshot("sumasuma.png",region=(1128,569, 33, 23))

time.sleep(1)
if pa.locateCenterOnScreen(r"C:\Python\Modificacion de Pedidos\Images\sumasuma.png", confidence=0.7,region=(1128,569, 33, 23)) == True:
    time.sleep(3)
    x1, y1 = pa.locateCenterOnScreen(r"C:\Python\Modificacion de Pedidos\Images\sumasuma.png", confidence=0.7,region=(1128,569, 33, 23))
    print(x1, y1)
    pa.moveTo(x=x1, y=y1)
else:
    pa.hotkey("alt","tab")
    print("NADA POR AQUI AMIGUITO")



# pa.hotkey("alt","tab")
# pa.moveTo(x=765, y=377)
# pa.dragTo(818, 616,0.5, button= 'left')
# pa.hotkey('ctrl','c')
# tped = (pyperclip.paste().replace('\n',',').replace(' ','').replace('\r','')).split(sep=",")
# print(tped)
# tped1=[]
# tped1= [print(tped1.append(tp)) for tp in tped if tp != "" ]
# # for tp in tped:
# #     if tp != "":
# #         tped1.append(tp)
# #     else:
# #         pass
# tped11 = np.array([tped1]).T
  
# print(tped11)  






# #DATAFRAME PANDAS
# file = (r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\Modificador de Pedidos 23.xlsx")
# df2= pd.read_excel(file, sheet_name="Pedido")
# pedidos = df2[df2["Estado"]=="Pendiente"]



# #EXCEL
# wb = load_workbook(r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\Modificador de Pedidos 23.xlsx")
# sheet = wb["Pedido"]


# while len(pedidos) > 0:
#     def PED_CONFIRM(df2):
#         xlsx_client, xlsx_estado = sheet["B1":"B" + str(len(df2["Estado"])+1)], sheet["E1":"E"+ str(len(df2["Estado"])+1)]
#         cli_est = zip(xlsx_client,xlsx_estado)
#         client= int(pedidos.iat[0,1])
#         for k,v  in cli_est:
#             if k[0].value == client:
#                 v[0].value = "OK"
#         wb.save(r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\Modificador de Pedidos 23.xlsx")
#         time.sleep(1)  
#         df2= pd.read_excel(file, sheet_name="Pedido")
#         pedidos = df2[df2["Estado"]=="Pendiente"]







    # xlsx_client, xlsx_estado = sheet["B1":"B" + str(len(df2["Estado"])+1)], sheet["E1":"E"+ str(len(df2["Estado"])+1)]
    # cli_est = zip(xlsx_client,xlsx_estado)
    # client= int(pedidos.iat[0,1])
    # print(client)
    # for k,v  in cli_est:
    #     if k[0].value == client:
    #         v[0].value = "OK"


    # wb.save(r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\Modificador de Pedidos 23.xlsx")
    # time.sleep(1)  
    # df2= pd.read_excel(file, sheet_name="Pedido")
    # pedidos = df2[df2["Estado"]=="Pendiente"]
    # print(pedidos)

#print(dict(cli_est))



#print(sheetrange)

      





#df3Eliminar= df3[df3["Tipo"].isin(["Eliminar"])]
#df3Modificar= df3[df3["Tipo"].isin(["Modificacion"])]
#df3Agregar= df3[df3["Tipo"].isin(["Agregar"])]





# df3Eliminar= df3[df3["Tipo"].isin(["Eliminar"])]
# cambiar= df3Eliminar[df3Eliminar["Cliente"].isin([client])]
# cambiar= cambiar["Estado"]= "ok"

# print(client)
# print(df3Eliminar[::-1])
# print(cambiar)
# print(df3Eliminar[::-1])