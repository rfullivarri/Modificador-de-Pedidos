import datetime
import os
import sys
import time

import keyboard
import numpy as np
import pandas as pd
import psutil
import pyautogui as pa
import pyperclip
from openpyxl import load_workbook

"""
Para reconocimiento de imagenes tenes OPENCV-PYTHON MATPLOTLIB Y NUMPY  
"""

#MAIN FILE
file = (r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\Modificador de Pedidos 23.xlsx")
#PASSWORD
df = pd.read_excel(file, sheet_name="Contraseñas")
user1 = df.iat[0,1]
password1 = df.iat[0,2]
#PEDIDOS SHEET
df2= pd.read_excel(file, sheet_name="Pedido")
pedidos = df2[df2["Estado"]=="Pendiente"]

try:
    client= int(pedidos.iat[0,1])
except IndexError:
    print("No hay Pedidos Pendientes de modificacion")
    sys.exit()

#MODIFICATION SHEET
df3= pd.read_excel(file, sheet_name="Modificacion") #, index="Cliente")

#EXCEL FILE OPENPYXL
wb = load_workbook(r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\Modificador de Pedidos 23.xlsx")
sheet = wb["Pedido"]


#FILTRO DE FECHAS
now = datetime.datetime.utcnow()
now = now.strftime('%d-%m-%y')
now2 = datetime.datetime.utcnow()
dayback = now2 - datetime.timedelta(days=10)
dayback = dayback.strftime('%d-%m-%y')


df3 = df3[df3["Estado"]=="Pendiente"]

df3Deleted = df3[df3["Tipo"].isin(["Eliminar"])]##
df3Deleted1 = df3Deleted[df3Deleted["Cliente"].isin([client])]##

df3Modify = df3[df3["Tipo"].isin(["Modificacion"])]##
df3Modify1 = df3Modify[df3Modify["Cliente"].isin([client])]##


df3Add = df3[df3["Tipo"].isin(["Agregar"])]##
df3Add1 = df3Add[df3Add["Cliente"].isin([client])]##

lendf3Deleted= len(df3Deleted1)
lendf3Modify= len(df3Modify1)
lendf3Add= len(df3Add1)


#INICIATE

is_open_truck ="pcsws.exe" in (i.name() for i in psutil.process_iter())

if is_open_truck == False:
    os.system("start C:\\truck.WS")
    time.sleep(2)
    pa.typewrite("AS400",0.2)
    keyboard.press("TAB")
    pa.typewrite("AS400",0.2)
    keyboard.press("ENTER")
    time.sleep(12)

    #LOGIN
    pa.typewrite(user1, 0.2)
    keyboard.press("TAB")
    pa.typewrite(password1, 0.2)
    keyboard.press("ENTER")
    time.sleep(3)
    pa.hotkey("win","up")
    for i in range(0,3):
        pa.typewrite("\n")
        time.sleep(2)
else:
    pass

while len(pedidos) > 0:
 #----------------------------------------------------SEARCH ORDER-----------------------------------------------------------------------------

    #pa.hotkey("alt","tab")
    for i in range(0,4):
            pa.press("f3")
            time.sleep(0.2) 

    FINDPED= [3,6,1]
    for i in FINDPED:
        pa.typewrite(f"{str(i)}\n")
        time.sleep(1.5)


    #CLIENT FILTER PENDING MODIFICATION
    for c in range(0,2):
        for i in range(0,(6-len(str(client)))):
            pa.typewrite("0")
        pa.typewrite(str(client),0.2)
    #pa.typewrite(str(client),0.2)
    pa.typewrite(dayback)
    pa.typewrite(now)
    for i in range(0,9):
            keyboard.press("TAB")
            time.sleep(0.15)
    for i in range(0,4):
            keyboard.press("SUPR")
            time.sleep(0.2)
    for i in range(0,4):
            pa.press("ENTER")
            time.sleep(0.15) 

    #CASH SALE FILTER
    time.sleep(1.5) 
    pa.press("f9")
    pa.typewrite("001\n", 0.2)
    pa.press("ENTER")

    #CUSTOMER SEARCH
    pa.moveTo(x=228, y=390)
    pa.dragTo(346, 388,0.5, button= 'left')
    pa.hotkey('ctrl','c')
    find_pedido = [pyperclip.paste()]

    #BREACKING POINT
    if find_pedido == ['(La lista']:
        xlsx_client, xlsx_comentario = sheet["B1":"B" + str(len(df2["Estado"])+1)], sheet["F1":"F"+ str(len(df2["Estado"])+1)]
        cli_com = zip(xlsx_client,xlsx_comentario)
        client= int(pedidos.iat[0,1])
        for k,v  in cli_com:
            if k[0].value == client:
                v[0].value = "NO SE ENCONTRO PEDIDO"
        time.sleep(1)        
        xlsx_estado = sheet["E1":"E"+ str(len(df2["Estado"])+1)]
        cli_est = zip(xlsx_client,xlsx_estado)
        client= int(pedidos.iat[0,1])
        for k,v  in cli_est:
            if k[0].value == client:
                v[0].value = "NO RESUELTO"            
        wb.save(r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\Modificador de Pedidos 23.xlsx")
        continue
    else:
        pass


    #TYPE OF ORDER
    pa.moveTo(x=765, y=377)
    pa.dragTo(818, 616,0.5, button= 'left')
    pa.hotkey('ctrl','c')
    tped = (pyperclip.paste().replace('\n',',').replace(' ','').replace('\r','')).split(sep=",")

    tped1 =[]
    for tp in tped:
        if tp != "":
            tped1.append(tp)
        else:
            pass
    tped11 = np.array([tped1]).T

    #ORDER CODE
    pa.moveTo(x=819, y=377)
    pa.dragTo(937, 616,0.5, button= 'left')
    pa.hotkey('ctrl','c')
    pedd = (pyperclip.paste().replace('\n',',').replace(' ','').replace('\r','')).split(sep=",")

    pedd1 =[]
    for pp in pedd:
        if pp != "":
            pedd1.append(pp)
        else:
            pass
    pedd11 = np.array([pedd1]).T

    #DATAFRAMES TO FIND TRANSFER ORDER
    dftped =  pd.DataFrame(tped11,columns=["T_Pedidos"])
    dfpedd =  pd.DataFrame(pedd11, columns=["Pedidos"])


    #TRANSFER ORDER
    xtruck361= 228
    ytruck361= 397

    pa.click(x=xtruck361,y=(ytruck361+(22*(len(dfpedd)-1))))
    pa.typewrite("13", 0.2)
    pa.press("ENTER")
    time.sleep(2)



 #-------------------------------------------------------------------------------------------------------------------------------------



 #--------------------------------------------------FIND TRANSFER ORDER--------------------------------------------------------------- 

    for i in range(0,5):
        pa.press("f3")
        time.sleep(0.5)

    MODPED= [3,1,1]
    for i in MODPED:
        pa.typewrite(f"{str(i)}\n")
        time.sleep(1.5)


    #FILTER IN TRANSFER ORDER PLACE (3 1 1) 
    tped311= int(dftped.loc[dftped.index[-1]])
    pedd311= int(dfpedd.loc[dfpedd.index[-1]])

    time.sleep(1)
    pa.click(x=685, y=612)
    pa.write(str(client), 0.2)
    pa.click(x=290, y=300)
    pa.write(str(tped311), 0.2)
    pa.click(x=532, y=300)
    pa.write(f"0{str(pedd311)}", 0.2)
    pa.press("ENTER")
    time.sleep(1.5)
    pa.click(x=228, y=324)
    time.sleep(0.5)
    pa.write("2")
    pa.press("ENTER")

    time.sleep(1.5)
    pa.click(x=953, y=394)
    time.sleep(1.5)
    for i in range(0,11):
            keyboard.press("SUPR")
            time.sleep(0.1)
    time.sleep(0.5)
    pa.click(x=436, y=420)
    pa.write("804")
    time.sleep(1)            
    pa.press("ENTER")
    time.sleep(3)



 #--------------------------------------------------------------------------------------------------------------------------------------------



 #--------------------------------------------------SCRAPPING ORDER DATA----------------------------------------------------------------------
    CODTruck1 =[]
    CANTTruck1 =[]
    plus=0

    while plus ==0:
        plus += 1
        #TRUCK ORDER CODE
        pa.moveTo(x=242, y=304)
        pa.dragTo(311, 615,0.5, button= 'left')
        pa.hotkey('ctrl','c')
        CODTruck= (pyperclip.paste().replace('\n',',').replace(' ','').replace('\r','')).split(sep=",")
        time.sleep(1)
        for COD in CODTruck:
            if COD != "":
                CODTruck1.append(COD)
            else:
                pass

        if len(CODTruck1) == 12:
            plus -= 1
            pass
        elif len(CODTruck1) == 24:
            plus -= 1
            pass
        else:
            pass    
        
        #SKU QUANTITIES IN THE ORDER  
        pa.moveTo(x=500, y=310)
        pa.dragTo(613, 615,0.5, button= 'left')
        pa.hotkey('ctrl','c')
        CANTTruck= (pyperclip.paste().replace('\n',',').replace(' ','').replace('\r','')).split(sep=",")
        time.sleep(1)
        for CANT in CANTTruck:
            if CANT != "":
                CANTTruck1.append(CANT)
            else:
                pass
        pa.press("pagedown")
        
    pa.press("pageup")    
    CANTTruck1 = np.array(CANTTruck1)
    CODTruck1 = np.array(CODTruck1)
    CODCANT= {CODTruck1[i]: CANTTruck1[i] for i in range(len(CODTruck1))}

 #------------------------------------------------------------------------------------------------------------------------------------------------



 #--------------------------------------------------MODIFICATIONS FUNCTIONS----------------------------------------------------------------------

    #MODIFICATION
    def  MODIFICATE():
        positionCX= 504
        positionCY= 324
        for skm in df3Modify["SKU"]: 
            pa.press("f11")
            time.sleep(1)
            dfsku = df3Modify[df3Modify["SKU"].isin([skm])]
            df3Cant = (dfsku["Cantidad"].astype("Int64")).to_string(index=False)
            for k , skt in enumerate(CODTruck1):
                pa.press("f11")
                if str(skt) == str(skm):
                    #SKU POSITIONS
                    if k >= 12: #NEXT PAGE
                        pa.press("pagedown")
                        if k-12 == 0:       
                            pa.moveTo(x=positionCX, y=(positionCY))
                            time.sleep(0.5)
                        else :
                            pa.moveTo(x=positionCX, y=(positionCY+(22*(k-12))))
                            time.sleep(0.5)   
                    elif k == 0: #FIRST POSITION
                        pa.press("f11")       
                        pa.moveTo(x=positionCX, y=(positionCY))
                        time.sleep(0.5) 
                    else: #EVERYONE ELSE
                        pa.press("f11")       
                        pa.moveTo(x=positionCX, y=(positionCY+(22*(k))))
                        time.sleep(1)

                    #WRITTING QUANTITIE    
                    if int(df3Cant) < 0: #NEGATIVES
                        newcant=(int(CODCANT.get(str(skm)))+ int(df3Cant))  #estilo de filtro 
                        pa.click()
                        time.sleep(1)
                        for i in range(0,(9-len(str(newcant)))):
                            pa.typewrite("0") 
                        pa.typewrite(str(newcant),0.1)
                        pa.press("ENTER")
                        pa.press("ENTER")
                        time.sleep(1)
                        pa.press("pageup")   
                    elif len(df3Cant) < 8: #POSITIVES
                        pa.click()
                        time.sleep(1)
                        for i in range(0,(9-len(df3Cant))):
                            pa.typewrite("0")
                        pa.typewrite(str(df3Cant),0.1)
                        pa.press("ENTER")
                        pa.press("ENTER")
                        time.sleep(1)
                        pa.press("pageup")
                    else:
                         pass   
                         
    #AGREGAR
    def ADD():    
        for ska in df3Add1: 
            pa.press("f6")
            dfsku = df3Add[df3Add["SKU"].isin([ska])]
            df3Cant = (dfsku["Cantidad"].astype("Int64")).to_string(index=False)
            for k , skt in enumerate(CODTruck1):
                if str(skt) == str(ska):
                    #print(MODIFICAR())        
                    print("ya esta ese sku amigo")
                else:
                    pass
            if len(str(ska)) < 5:
                pa.typewrite("0") 
                pa.typewrite(str(ska),0.1)
                for i in range(0,(9-len(df3Cant))):
                    pa.typewrite("0")
                pa.typewrite(str(df3Cant),0.1)
                pa.press("ENTER")
                pa.press("f11")
            else:
                pa.typewrite(str(ska),0.1)
                for i in range(0,(9-len(df3Cant))):
                    pa.typewrite("0")
                pa.typewrite(str(df3Cant),0.1)
                pa.press("ENTER")
                pa.press("f11")

    #DELETED
    def DELETED():
        positionDX= 219
        positionDY= 324

        for skd in df3Deleted1: 
            pa.hotkey("shift","f1")
            time.sleep(1)
            pa.moveTo(x=positionDX, y=positionDY)
            for k , skt in enumerate(CODTruck1):
                if str(skt) == str(skd):
                    if k == 0:
                        #pa.moveTo(x=positionDX, y=positionDY)
                        #
                        # time.sleep(1)
                        pa.click()
                        time.sleep(0.3)
                        pa.press("4")
                        time.sleep(0.5)
                        pa.press("ENTER")
                        pa.press("ENTER") 
                    else:           
                        pa.moveTo(x=positionDX, y=(positionDY+(22*(k))))
                        time.sleep(1)
                        pa.click()
                        pa.press("4")
                        pa.press("ENTER")
                        pa.press("ENTER")
                else:
                    pass    


 #-----------------------------------------------------------------------------------------------------------------------------------------------



 #---------------------------------------------------MODIFICATION FLOW---------------------------------------------------------------------------
    if lendf3Modify > 0:
        print(MODIFICATE())
    else:
        pass


    if lendf3Deleted > 0:
        print(DELETED())
    else:
        pass


    if lendf3Add > 0:
        print(ADD())
    else:
        pass


    time.sleep(2)
    pa.press("ENTER")
    pa.press("ENTER")
    pa.press("f12")
    time.sleep(2)
    pa.typewrite("7")
    time.sleep(3)

    pa.press("ENTER")
    pa.press("ENTER")
    pa.press("ENTER")
    time.sleep(2)
    
    for i in range(0,4):
        pa.press("f3")
        time.sleep(0.5)
        

    #ORDER CONFIRMATION IN EXCEL    
    xlsx_client, xlsx_estado = sheet["B1":"B" + str(len(df2["Estado"])+1)], sheet["E1":"E"+ str(len(df2["Estado"])+1)]
    cli_est = zip(xlsx_client,xlsx_estado)
    client= int(pedidos.iat[0,1])
    for k,v  in cli_est:
        if k[0].value == client:
            v[0].value = "OK"
    wb.save(r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\Modificador de Pedidos 23.xlsx")
    time.sleep(2)  
    df2= pd.read_excel(file, sheet_name="Pedido")
    pedidos = df2[df2["Estado"]=="Pendiente"]
    try:
        client= int(pedidos.iat[0,1])
    except  IndexError:
        client= 0 
        print("No hay mas pedidos por modificar")  
    time.sleep(3)

        
    



