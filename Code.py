import pandas as pd
import win32com.client
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import subprocess
from time import sleep
import sys  
import datetime
from email.mime.base import MIMEBase

hoje =  datetime.date.today()
hoje = pd.to_datetime(hoje, dayfirst=True)
hoje = hoje.strftime('%d/%m/%Y')
hoje = str(hoje.replace("/","."))


um_dia_atraz =  datetime.date.today() - datetime.timedelta(days=1)
um_dia_atraz = pd.to_datetime(um_dia_atraz, dayfirst=True)
um_dia_atraz = um_dia_atraz.strftime('%d/%m/%Y')
um_dia_atraz = str(um_dia_atraz.replace("/","."))


dois_dias_atras =  datetime.date.today() - datetime.timedelta(days=2)
dois_dias_atras = pd.to_datetime(dois_dias_atras, dayfirst=True)
dois_dias_atras = dois_dias_atras.strftime('%d/%m/%Y')
dois_dias_atras = str(dois_dias_atras.replace("/","."))


tres_dias_atras =  datetime.date.today() - datetime.timedelta(days=3)
tres_dias_atras = pd.to_datetime(tres_dias_atras, dayfirst=True)
tres_dias_atras = tres_dias_atras.strftime('%d/%m/%Y')
tres_dias_atras = str(tres_dias_atras.replace("/","."))

quatro_dias_atras =  datetime.date.today() - datetime.timedelta(days=4)
quatro_dias_atras = pd.to_datetime(quatro_dias_atras, dayfirst=True)
quatro_dias_atras = quatro_dias_atras.strftime('%d/%m/%Y')
quatro_dias_atras = str(quatro_dias_atras.replace("/","."))

cinco_dias_atras =  datetime.date.today() - datetime.timedelta(days=5)
cinco_dias_atras = pd.to_datetime(cinco_dias_atras, dayfirst=True)
cinco_dias_atras = cinco_dias_atras.strftime('%d/%m/%Y')
cinco_dias_atras = str(cinco_dias_atras.replace("/","."))


class SapGui():
    
    def __init__ (self):
        self.path = r"caminho onde sap está instalado"        
        subprocess.Popen(self.path)
        aberto = False
        sleep(1)
        while not aberto:
            sleep(1)
            try:
                sleep(1)
                self.SapGuiauto = win32com.client.GetObject('SAPGUI')
                application  = self.SapGuiauto.GetScriptingEngine
                self.connection = application.OpenConnection('010 - PRD - EP1 - SAP ECC Matriz')
                aberto = True
            except:
                pass
        connections = len(self.connection.Children)
        self.session = self.connection.Children(connections-1)
        self.session.findById("wnd[0]").maximize

    def SapLogin(self):
        try:
            self.session.findById("wnd[0]").maximize      
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "usuario_sap"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "senha_sap"
            self.session.findById("wnd[0]").sendVKey(0)
        except:
            print(sys.exc_info()[0])

SapGui().SapLogin()


def Zjob_todos(): 
 
        
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    conn = application.Children(0)  
    session = conn.Children(0)


    def zjob_entra():
        session.findById("wnd[0]/tbar[0]/okcd").text = "zjob"
        session.findById("wnd[0]").sendVKey(0)
    zjob_entra()
    
    try:
        def Pegar_txt_me2m_componente():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Compoente_Consultar_Pedido_De_Compra.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        Pegar_txt_me2m_componente() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def Pegar_txt_me2m_componente():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Compoente_Consultar_Pedido_De_Compra.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            Pegar_txt_me2m_componente()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def Pegar_txt_me2m_componente():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Compoente_Consultar_Pedido_De_Compra.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                Pegar_txt_me2m_componente()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def Pegar_txt_me2m_componente():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Compoente_Consultar_Pedido_De_Compra.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    Pegar_txt_me2m_componente() 
                except:
                    try: 
                        session.findById("wnd[0]").sendVKey(0)
                        def Pegar_txt_me2m_componente():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Compoente_Consultar_Pedido_De_Compra.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        Pegar_txt_me2m_componente()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def Pegar_txt_me2m_componente():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Compoente_Consultar_Pedido_De_Compra.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            Pegar_txt_me2m_componente()
                        except:
                            print('erro')


    
    try:
        def Pegar_txt_me2m_componente():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZCO144"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)

                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(-1, "TIME")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("TIME")
                session.findById("wnd[0]/tbar[1]/btn[40]").press()


                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZCO144.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        Pegar_txt_me2m_componente() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def Pegar_txt_me2m_componente():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZCO144"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)

                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(-1, "TIME")
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("TIME")
                    session.findById("wnd[0]/tbar[1]/btn[40]").press()



                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZCO144.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            Pegar_txt_me2m_componente()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def Pegar_txt_me2m_componente():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZCO144"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZCO144.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                Pegar_txt_me2m_componente()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def Pegar_txt_me2m_componente():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZCO144"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZCO144.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    Pegar_txt_me2m_componente() 
                except:
                    try: 
                        session.findById("wnd[0]").sendVKey(0)
                        def Pegar_txt_me2m_componente():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZCO144"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZCO144.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        Pegar_txt_me2m_componente()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def Pegar_txt_me2m_componente():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZCO144"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZCO144.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            Pegar_txt_me2m_componente()
                        except:
                            print('erro')
                            


    try:
        def Pegar_txt_me2m_fert():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Fert_Consultar_Pedido_De_Compra.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        Pegar_txt_me2m_fert() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def Pegar_txt_me2m_fert():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Fert_Consultar_Pedido_De_Compra.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            Pegar_txt_me2m_fert()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def Pegar_txt_me2m_fert():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Fert_Consultar_Pedido_De_Compra.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                Pegar_txt_me2m_fert()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def Pegar_txt_me2m_fert():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Fert_Consultar_Pedido_De_Compra.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    Pegar_txt_me2m_fert() 
                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0) 
                        def Pegar_txt_me2m_fert():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Fert_Consultar_Pedido_De_Compra.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        Pegar_txt_me2m_fert()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def Pegar_txt_me2m_fert():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ME2M"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Me2m_Fert_Consultar_Pedido_De_Compra.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            Pegar_txt_me2m_fert()
                        except:
                            print('erro')




    try:
        def SQVI_BLOQUEI():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "SQVI_BLOQUEI"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        SQVI_BLOQUEI() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def SQVI_BLOQUEI():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "SQVI_BLOQUEI"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            SQVI_BLOQUEI()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def SQVI_BLOQUEI():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "SQVI_BLOQUEI"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                SQVI_BLOQUEI()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def SQVI_BLOQUEI():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "SQVI_BLOQUEI"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    SQVI_BLOQUEI() 
                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0) 
                        def SQVI_BLOQUEI():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "SQVI_BLOQUEI"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        SQVI_BLOQUEI()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def SQVI_BLOQUEI():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "SQVI_BLOQUEI"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            SQVI_BLOQUEI()
                        except:
                            print('erro')




    try:
        def ZPP101():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZPP101"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZPP101_Fert_Ordem_Aberta.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        ZPP101() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def ZPP101():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZPP101"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZPP101_Fert_Ordem_Aberta.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            ZPP101()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def ZPP101():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZPP101"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZPP101_Fert_Ordem_Aberta.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                ZPP101()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def ZPP101():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZPP101"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZPP101_Fert_Ordem_Aberta.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    ZPP101() 
                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0) 
                        def ZPP101():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZPP101"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZPP101_Fert_Ordem_Aberta.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        ZPP101()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def ZPP101():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZPP101"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZPP101_Fert_Ordem_Aberta.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            ZPP101()
                        except:
                            print('erro')


    try:
        def ZSD138_CARTE():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        ZSD138_CARTE() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def ZSD138_CARTE():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            ZSD138_CARTE()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def ZSD138_CARTE():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                ZSD138_CARTE()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def ZSD138_CARTE():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    ZSD138_CARTE() 
                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0) 
                        def ZSD138_CARTE():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        ZSD138_CARTE()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def ZSD138_CARTE():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            ZSD138_CARTE()
                        except:
                            print('erro')



    try:
        def ZSD138_CARTE():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        ZSD138_CARTE() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def ZSD138_CARTE():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            ZSD138_CARTE()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def ZSD138_CARTE():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                ZSD138_CARTE()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def ZSD138_CARTE():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    ZSD138_CARTE() 
                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0) 
                        def ZSD138_CARTE():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        ZSD138_CARTE()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def ZSD138_CARTE():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZSD138_CARTE"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZSD138_carteira_fert.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            ZSD138_CARTE()
                        except:
                            print('erro')

    try:
        def ULTIMA_VENDA():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ULTIMA_VENDA"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ultima_venda.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        ULTIMA_VENDA() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def ULTIMA_VENDA():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ULTIMA_VENDA"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ultima_venda.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            ULTIMA_VENDA()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def ULTIMA_VENDA():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ULTIMA_VENDA"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ultima_venda.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                ULTIMA_VENDA()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def ULTIMA_VENDA():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ULTIMA_VENDA"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ultima_venda.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    ULTIMA_VENDA() 
                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0) 
                        def ULTIMA_VENDA():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ULTIMA_VENDA"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ultima_venda.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        ULTIMA_VENDA()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def ULTIMA_VENDA():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ULTIMA_VENDA"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ultima_venda.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            ULTIMA_VENDA()
                        except:
                            print('erro')
    try:
        def ZMM34_VENDA():
            def me2m ():
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZMM34_VENDA"
            me2m()          
            def roda_zjob():
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
            roda_zjob()
            def caminho():
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
            caminho()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZMM22_Fert_Total_Venda_8_Meses.txt"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            def voltar():
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3)  
            voltar()
        ZMM34_VENDA() 

    except:
        try:
            session.findById("wnd[0]").sendVKey(0)
            def ZMM34_VENDA():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZMM34_VENDA"
                me2m()          
                def roda_zjob():
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                roda_zjob()
                def caminho():
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                caminho()
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZMM22_Fert_Total_Venda_8_Meses.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            ZMM34_VENDA()     
        except:
            try:
                session.findById("wnd[0]").sendVKey(0)
                def ZMM34_VENDA():
                    def me2m ():
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZMM34_VENDA"
                    me2m()          
                    def roda_zjob():
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                    roda_zjob()
                    def caminho():
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    caminho()
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZMM22_Fert_Total_Venda_8_Meses.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    def voltar():
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)  
                    voltar()
                ZMM34_VENDA()   
            except:
                try:
                    session.findById("wnd[0]").sendVKey(0) 
                    def ZMM34_VENDA():
                        def me2m ():
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZMM34_VENDA"
                        me2m()          
                        def roda_zjob():
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                        roda_zjob()
                        def caminho():
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        caminho()
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZMM22_Fert_Total_Venda_8_Meses.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        def voltar():
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)  
                        voltar()
                    ZMM34_VENDA() 
                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0) 
                        def ZMM34_VENDA():
                            def me2m ():
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZMM34_VENDA"
                            me2m()          
                            def roda_zjob():
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                            roda_zjob()
                            def caminho():
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            caminho()
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZMM22_Fert_Total_Venda_8_Meses.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            def voltar():
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)  
                            voltar()
                        ZMM34_VENDA()
                    except: 
                        try:
                            session.findById("wnd[0]").sendVKey(0) 
                            def ZMM34_VENDA():
                                def me2m ():
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ZMM34_VENDA"
                                me2m()          
                                def roda_zjob():
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                roda_zjob()
                                def caminho():
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                caminho()
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZMM22_Fert_Total_Venda_8_Meses.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                def voltar():
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)  
                                voltar()
                            ZMM34_VENDA()
                        except:
                            print('erro')
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()  
Zjob_todos()


#0
import pandas as pd
import datetime
with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\ZCO144.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()



linhas = [x for x in linhas if 'Lista tecnica multi-nível' not in x]
linhas = [x for x in linhas if '-----------------------------------------' not in x]
linhas = [x for x in linhas if '-------------------------------------------' not in x]

linhas = [x for x in linhas if x[0][0]=='|']
linhas = [x.split('|') for x in linhas]
linhas = [x[1:-1] for x in linhas]

zco144 = []
for linha in linhas:     
    dados_temp = []
    for dado in linha:
        dado = dado.rstrip().lstrip()
        dados_temp.append(dado)
    zco144.append(dados_temp)

zco144 = pd.DataFrame(zco144)
zco144.drop_duplicates(inplace=True)
zco144 = zco144.rename(columns=zco144.iloc[0]).drop(zco144.index[0])
zco144 = zco144.rename(columns={'Comp. secundário': 'Componente','Tipo mat.sec.': 'Tipo'})
zco144 = zco144.loc[(zco144['Tipo'] != 'VERP')]
zco144 = zco144.sort_values(by=['Material','Nível','Sentido'])
zco144.drop(['Nível','Sentido'],axis=1,inplace=True)
zco144.drop_duplicates(inplace=True)
zco144.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco144.xlsx',index=None)


#1
import pandas as pd
import datetime
with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

linhas = [x for x in linhas if 'Sem dados' not in x]
linhas = [x for x in linhas if '--------------------------------------------------------------------------------' not in x]
linhas = [x for x in linhas if '-----------------------------------' not in x]
linhas = [x for x in linhas if '---------------------------------' not in x]
linhas = [x for x in linhas if '--------------------------' not in x]
linhas = [x for x in linhas if '-------------------------' not in x]
linhas = [x for x in linhas if 'Estatíst.dados' not in x]
linhas = [x for x in linhas if 'Registros transfs.' not in x]

linhas = [x for x in linhas if x[0][0]=='|']
linhas = [x.split('|') for x in linhas]
linhas = [x[1:-1] for x in linhas]

Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao = []
for linha in linhas:     
    dados_temp = []
    for dado in linha:
        dado = dado.rstrip().lstrip()
        dados_temp.append(dado)
    Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.append(dados_temp)

Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao = pd.DataFrame(Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao)

Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao = Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.rename(columns=Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.iloc[0]).drop(Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.index[0])
Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao = Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.rename(columns={
                                                                                                            'TxtBreveMaterial': 'Descr.',           
                                                                                                            'SM': 'Bloq.Prod',
                                                                                                            'Denominação': 'Descr.',
                                                                                                            'St': 'Bloq.Venda',
                                                                                                            'Denominação': 'Descr.',
                                                                                                            'EM': 'GE',                                                   
                                                                                                            })
#Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao['Criado'] = Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao['Criado'].str.replace('.', '/')
Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.xlsx',index=None)


#ESTOQUE FERT
import sqlalchemy
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
servidor_dns = 'Servidor'
servidor_database = 'Base ded dados'
url = f'mssql+pyodbc://@{servidor_dns}/{servidor_database}?trusted_connection=yes&driver=SQL+Server'
engine = sqlalchemy. create_engine (url)

Base_Mard = pd.read_sql("""
                    SELECT 
                    MARD.MATNR as Material,  
                    MARD.LABST as Livre , 
                    MARD.SPEME as Bloqueado , 
                    MARD.INSME as Qualidade , 
                    MARD.UMLME as TransfDeposito 
                    FROM MARD 
                    LEFT JOIN MARA 
                    ON MARA.MATNR = MARD.MATNR 
                    WHERE MARA.MTART in ( 'FERT')""",engine)

Base_Mard = Base_Mard.drop_duplicates()
Base_Mard['Material'] = Base_Mard['Material'].str.lstrip('00000000000000000')
Base_Mard['Livre'] = Base_Mard.groupby('Material')["Livre"].transform(np.sum)
Base_Mard['Bloqueado'] = Base_Mard.groupby('Material')["Bloqueado"].transform(np.sum)
Base_Mard['Qualidade'] = Base_Mard.groupby('Material')["Qualidade"].transform(np.sum)
Base_Mard['TransfDeposito'] = Base_Mard.groupby('Material')["TransfDeposito"].transform(np.sum)
Base_Mard = Base_Mard.drop_duplicates()


Marc = pd.read_sql("""
                    SELECT 
                    MARC.MATNR as Material,  
                    MARC.UMLMC as TransfCentro 

                    FROM MARC 
                    LEFT JOIN MARA 
                    ON MARA.MATNR = MARC.MATNR 
                    WHERE MARA.MTART in ( 'FERT')""",engine)

Marc = Marc.drop_duplicates()
Marc['Material'] = Marc['Material'].str.lstrip('00000000000000000')
Marc['TransfCentro'] = Marc.groupby('Material')["TransfCentro"].transform(np.sum)
Marc = Marc.drop_duplicates()


Base_Mard  = Base_Mard.merge(Marc, on='Material', how='left')

Base_Mard = Base_Mard.fillna({ 
                'Livre': 0, 
                'Bloqueado': 0,
                'Qualidade': 0,
                'TransfDeposito': 0, 
                'TransfCentro': 0,                    
                })
Base_Mard['Saldo'] = Base_Mard['Livre'] + Base_Mard['Bloqueado'] + Base_Mard['Qualidade'] + Base_Mard['TransfDeposito'] + Base_Mard['TransfCentro']
Estoque = Base_Mard.loc[:,['Material','Saldo']]
Estoque.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco32_estoque_fert.xlsx',index=None )



#ESTOQUE Componente
import sqlalchemy
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
servidor_dns = 'Servidor'
servidor_database = 'Base de dados'
url = f'mssql+pyodbc://@{servidor_dns}/{servidor_database}?trusted_connection=yes&driver=SQL+Server'
engine = sqlalchemy. create_engine (url)

Base_Mard_Comp = pd.read_sql("""
                    SELECT 
                    MARD.MATNR as Material,  
                    MARD.LABST as Livre , 
                    MARD.SPEME as Bloqueado , 
                    MARD.INSME as Qualidade , 
                    MARD.UMLME as TransfDeposito 
                    FROM MARD 
                    LEFT JOIN MARA 
                    ON MARA.MATNR = MARD.MATNR 
                    WHERE MARA.MTART in ('ZINT','FERT','HALB','ZPC2','ZPC3','ZPC4','ZPC5','ZPC6','ZINT','DBBS','ZAR1','ROH','VERP')""",engine)

Base_Mard_Comp = Base_Mard_Comp.drop_duplicates()
Base_Mard_Comp['Material'] = Base_Mard_Comp['Material'].str.lstrip('00000000000000000')
Base_Mard_Comp['Livre'] = Base_Mard_Comp.groupby('Material')["Livre"].transform(np.sum)
Base_Mard_Comp['Bloqueado'] = Base_Mard_Comp.groupby('Material')["Bloqueado"].transform(np.sum)
Base_Mard_Comp['Qualidade'] = Base_Mard_Comp.groupby('Material')["Qualidade"].transform(np.sum)
Base_Mard_Comp['TransfDeposito'] = Base_Mard_Comp.groupby('Material')["TransfDeposito"].transform(np.sum)
Base_Mard_Comp = Base_Mard_Comp.drop_duplicates()


Marc_componente = pd.read_sql("""
                    SELECT 
                    MARC.MATNR as Material,  
                    MARC.UMLMC as TransfCentro 

                    FROM MARC 
                    LEFT JOIN MARA 
                    ON MARA.MATNR = MARC.MATNR 
                    WHERE MARA.MTART in ( 'ZINT','FERT','HALB','ZPC2','ZPC3','ZPC4','ZPC5','ZPC6','ZINT','DBBS','ZAR1','ROH','VERP')""",engine)

Marc_componente = Marc_componente.drop_duplicates()
Marc_componente['Material'] = Marc_componente['Material'].str.lstrip('00000000000000000')
Marc_componente['TransfCentro'] = Marc_componente.groupby('Material')["TransfCentro"].transform(np.sum)
Marc_componente = Marc_componente.drop_duplicates()


Base_Mard_Comp  = Base_Mard_Comp.merge(Marc_componente, on='Material', how='left')
Base_Mard_Comp = Base_Mard_Comp.fillna({ 
                'Livre': 0, 
                'Bloqueado': 0,
                'Qualidade': 0,
                'TransfDeposito': 0, 
                'TransfCentro': 0,                  
                })
Base_Mard_Comp['Saldo'] = Base_Mard_Comp['Livre'] + Base_Mard_Comp['Bloqueado'] + Base_Mard_Comp['Qualidade'] + Base_Mard_Comp['TransfDeposito'] + Base_Mard_Comp['TransfCentro']
Base_Mard_Comp = Base_Mard_Comp.rename(columns={'Material': 'Componente'})
Estzco32_estoque_componenteoque = Base_Mard_Comp.loc[:,['Componente','Saldo']]
Estzco32_estoque_componenteoque.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco32_estoque_componente.xlsx',index=None)










#4

import pandas as pd
with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\ZMM22_Fert_Total_Venda_8_Meses.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

linhas = [x for x in linhas if '----------------------------------------------' not in x]
linhas = [x for x in linhas if '--------------------------------------------' not in x]
linhas = [x for x in linhas if '----------------------------' not in x]
linhas = [x for x in linhas if '-------------------------------' not in x]
linhas = [x for x in linhas if '---------------------------------' not in x]
linhas = [x for x in linhas if '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
linhas = [x for x in linhas if '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]

linhas = [x for x in linhas if 'Registros transfs.' not in x]
linhas = [x for x in linhas if 'Histórico de consumo dos últimos 12 meses ' not in x]

linhas = [x for x in linhas if 'Gerado em' not in x]
linhas = [x for x in linhas if 'Registros processados:' not in x]

linhas = [x for x in linhas if 'Lnhs.totais determinadas' not in x]
linhas = [x for x in linhas if 'Estatíst.dados' not in x]
linhas = [x for x in linhas if '|Material              |  X   |     |        |' not in x]
linhas = [x for x in linhas if '|Critérios de ordenação|Cresc.|Decr.|Subtotal|' not in x]
linhas = [x for x in linhas if '*' not in x]



linhas = [x for x in linhas if x[0][0]=='|']
linhas = [x.split('|') for x in linhas]
linhas = [x[1:-1] for x in linhas]
ZMM22_Fert_Total_Venda_8_Meses = []
for linha in linhas:     
    dados_temp = []
    
    for dado in linha:
        dado = dado.rstrip().lstrip()
        dados_temp.append(dado)
    ZMM22_Fert_Total_Venda_8_Meses.append(dados_temp)

ZMM22_Fert_Total_Venda_8_Meses = pd.DataFrame(ZMM22_Fert_Total_Venda_8_Meses)
ZMM22_Fert_Total_Venda_8_Meses = ZMM22_Fert_Total_Venda_8_Meses.drop_duplicates()
ZMM22_Fert_Total_Venda_8_Meses = ZMM22_Fert_Total_Venda_8_Meses.rename(columns=ZMM22_Fert_Total_Venda_8_Meses.iloc[0]).drop(ZMM22_Fert_Total_Venda_8_Meses.index[0])

#tirando ultimas linha em branco
len_ZMM22_Fert_Total_Venda_8_Meses = int(len(ZMM22_Fert_Total_Venda_8_Meses))
sete = int(7)
NumeroDeLinhas = int(len_ZMM22_Fert_Total_Venda_8_Meses - sete)
ZMM22_Fert_Total_Venda_8_Meses = ZMM22_Fert_Total_Venda_8_Meses.iloc[:NumeroDeLinhas, [0,2,3,4,5,6,7,8,9,10,11,12,13,14,15]]

ZMM22_Fert_Total_Venda_8_Meses['Total'] = ZMM22_Fert_Total_Venda_8_Meses['Total'].apply(lambda x: float(x.replace(".","").replace(",",".")))
ZMM22_Fert_Total_Venda_8_Meses['Média'] = ZMM22_Fert_Total_Venda_8_Meses['Média'].apply(lambda x: float(x.replace(".","").replace(",",".")))
ZMM22_Fert_Total_Venda_8_Meses.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ZMM22_Fert_Total_Venda_8_Meses.xlsx', index = None)


#5

import pandas as pd
with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\ZPP101_Fert_Ordem_Aberta.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

linhas = [x for x in linhas if 'Sem dados' not in x]
linhas = [x for x in linhas if 'Material' not in x]
linhas = [x for x in linhas if 'Material c/ Estoque e OP' not in x]
linhas = [x for x in linhas if 'Emitido' not in x]
linhas = [x for x in linhas if '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
linhas = [x for x in linhas if '---------------------------' not in x]
linhas = [x for x in linhas if '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
linhas = [x for x in linhas if x[0][0]=='|']
linhas = [x.split('|') for x in linhas]
linhas = [x[1:-1] for x in linhas]

ZPP101_Fert_Ordem_Aberta = []
for linha in linhas:     
    dados_temp = []
    for dado in linha:
        dado = dado.rstrip().lstrip()
        dados_temp.append(dado)
    ZPP101_Fert_Ordem_Aberta.append(dados_temp)



estrutura_df = {
    'Material': str,
    'Ordem': str,
    'Saldo': str,
}

ZPP101_Fert_Ordem_Aberta = pd.DataFrame(ZPP101_Fert_Ordem_Aberta)
ZPP101_Fert_Ordem_Aberta.columns = estrutura_df.keys()
ZPP101_Fert_Ordem_Aberta = ZPP101_Fert_Ordem_Aberta.astype(estrutura_df)
ZPP101_Fert_Ordem_Aberta['Saldo'] = ZPP101_Fert_Ordem_Aberta['Saldo'].apply(lambda x: str(x.replace("-","")))
ZPP101_Fert_Ordem_Aberta['Saldo'] = ZPP101_Fert_Ordem_Aberta['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
ZPP101_Fert_Ordem_Aberta[['Saldo']] = ZPP101_Fert_Ordem_Aberta[['Saldo']].apply(pd.to_numeric).round(4)
ZPP101_Fert_Ordem_Aberta.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ZPP101_Fert_Ordem_Aberta.xlsx' , index = None)


#6
import pandas as pd
with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\ZSD138_carteira_fert.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

linhas = [x for x in linhas if 'Sem dados' not in x]
linhas = [x for x in linhas if 'Material' not in x]
linhas = [x for x in linhas if '-----------------------------' not in x]
linhas = [x for x in linhas if '---------------------------' not in x]
linhas = [x for x in linhas if '----------------------------' not in x]
linhas = [x for x in linhas if 'Emissor  ' not in x]
linhas = [x for x in linhas if 'Data Emissão' not in x]
linhas = [x for x in linhas if x[0][0]=='|']
linhas = [x.split('|') for x in linhas]
linhas = [x[1:-1] for x in linhas]
ZSD138_carteira_fert = []
for linha in linhas:     
    dados_temp = []
    for dado in linha:
        dado = dado.rstrip().lstrip()
        dados_temp.append(dado)
    ZSD138_carteira_fert.append(dados_temp)

estrutura_df = {
    'Material': str,
    'Saldo': str,
}

ZSD138_carteira_fert = pd.DataFrame(ZSD138_carteira_fert)
ZSD138_carteira_fert.columns = estrutura_df.keys()
ZSD138_carteira_fert = ZSD138_carteira_fert.astype(estrutura_df)
ZSD138_carteira_fert['Saldo'] = ZSD138_carteira_fert['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
ZSD138_carteira_fert = ZSD138_carteira_fert.groupby("Material")["Saldo"].sum().round(3)
ZSD138_carteira_fert.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ZSD138_carteira_fert.xlsx')



from operator import index
import pandas as pd
import numpy as np
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\Me2m_Compoente_Consultar_Pedido_De_Compra.txt', encoding='latin1'  ) as f:
    linhas = f.readlines()
    f.close()   

linhas = [x for x in linhas if 'Sem dados' not in x]
linhas = [x for x in linhas if 'Material' not in x]
linhas = [x for x in linhas if '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
linhas = [x for x in linhas if '|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
linhas = [x for x in linhas if x[0][0]=='|']
linhas = [x.split('|') for x in linhas]
linhas = [x[1:-1] for x in linhas]

Me2m_Compoente_Consultar_Pedido_De_Compra = []
for linha in linhas:     
    dados_temp = []
    for dado in linha:
        dado = dado.rstrip().lstrip()
        dados_temp.append(dado)
    Me2m_Compoente_Consultar_Pedido_De_Compra.append(dados_temp)


Me2m_Compoente_Consultar_Pedido_De_Compra = pd.DataFrame(Me2m_Compoente_Consultar_Pedido_De_Compra )
Me2m_Compoente_Consultar_Pedido_De_Compra.head()

Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.iloc[:, [8,39,54]]



estrutura_df = {
    'Componente': str,
    'Saldo': str,
    'Un': str
}

Me2m_Compoente_Consultar_Pedido_De_Compra = pd.DataFrame(Me2m_Compoente_Consultar_Pedido_De_Compra, dtype=str)
Me2m_Compoente_Consultar_Pedido_De_Compra.columns = estrutura_df.keys()
Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.astype(estrutura_df)
Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.loc[(Me2m_Compoente_Consultar_Pedido_De_Compra['Componente'] != '')]
Me2m_Compoente_Consultar_Pedido_De_Compra['Saldo'] = Me2m_Compoente_Consultar_Pedido_De_Compra['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
Me2m_Compoente_Consultar_Pedido_De_Compra.head()
Me2m_Compoente_Consultar_Pedido_De_Compra['Saldo'] = Me2m_Compoente_Consultar_Pedido_De_Compra.groupby("Componente")["Saldo"].transform(np.sum)
Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.drop_duplicates()
Me2m_Compoente_Consultar_Pedido_De_Compra.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Me2m_Compoente_Consultar_Pedido_De_Compra.xlsx', index = None)



from operator import index
import pandas as pd
import numpy as np
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\Me2m_Compoente_Consultar_Pedido_De_Compra.txt', encoding='latin1'  ) as f:
    linhas = f.readlines()
    f.close()   

linhas = [x for x in linhas if 'Sem dados' not in x]
linhas = [x for x in linhas if 'Material' not in x]
linhas = [x for x in linhas if '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
linhas = [x for x in linhas if '|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
linhas = [x for x in linhas if x[0][0]=='|']
linhas = [x.split('|') for x in linhas]
linhas = [x[1:-1] for x in linhas]

Me2m_Compoente_Consultar_Pedido_De_Compra = []
for linha in linhas:     
    dados_temp = []
    for dado in linha:
        dado = dado.rstrip().lstrip()
        dados_temp.append(dado)
    Me2m_Compoente_Consultar_Pedido_De_Compra.append(dados_temp)


Me2m_Compoente_Consultar_Pedido_De_Compra = pd.DataFrame(Me2m_Compoente_Consultar_Pedido_De_Compra )
Me2m_Compoente_Consultar_Pedido_De_Compra.head()

Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.iloc[:, [8,39,54]]



estrutura_df = {
    'Material': str,
    'Saldo': str,
    'Un': str
}

Me2m_Compoente_Consultar_Pedido_De_Compra = pd.DataFrame(Me2m_Compoente_Consultar_Pedido_De_Compra, dtype=str)
Me2m_Compoente_Consultar_Pedido_De_Compra.columns = estrutura_df.keys()
Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.astype(estrutura_df)
Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.loc[(Me2m_Compoente_Consultar_Pedido_De_Compra['Material'] != '')]
Me2m_Compoente_Consultar_Pedido_De_Compra['Saldo'] = Me2m_Compoente_Consultar_Pedido_De_Compra['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
Me2m_Compoente_Consultar_Pedido_De_Compra.head()
Me2m_Compoente_Consultar_Pedido_De_Compra['Saldo'] = Me2m_Compoente_Consultar_Pedido_De_Compra.groupby("Material")["Saldo"].transform(np.sum)
Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.drop_duplicates()
Me2m_Compoente_Consultar_Pedido_De_Compra.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Me2m_Fert_Consultar_Pedido_De_Compra.xlsx', index = None)





from operator import index
import pandas as pd
with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\ultima_venda.txt', encoding='latin1'  ) as f:
    linhas = f.readlines()
    f.close()   

linhas = [x for x in linhas if 'Ùltima Data de um Tipo de' not in x]
linhas = [x for x in linhas if 'Data Base' not in x]
linhas = [x for x in linhas if 'Emitido em' not in x]
linhas = [x for x in linhas if 'Material' not in x]
linhas = [x for x in linhas if '--------------------------' not in x]
linhas = [x for x in linhas if '------------------------' not in x]

linhas = [x for x in linhas if x[0][0]=='|']
linhas = [x.split('|') for x in linhas]
linhas = [x[1:-1] for x in linhas]

ultima_venda = []
for linha in linhas:     
    dados_temp = []
    for dado in linha:
        dado = dado.rstrip().lstrip()
        dados_temp.append(dado)
    ultima_venda.append(dados_temp)

ultima_venda = pd.DataFrame(ultima_venda )


estrutura_df = {
    'Material': str,
    'Ult.Venda': str
}

ultima_venda = pd.DataFrame(ultima_venda, dtype=str)
ultima_venda.columns = estrutura_df.keys()
ultima_venda = ultima_venda.astype(estrutura_df)
ultima_venda.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ultima_venda.xlsx' , index = None)



import sqlalchemy
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
servidor_dns = 'Servidor'
servidor_database = 'Base de dados'
url = f'mssql+pyodbc://@{servidor_dns}/{servidor_database}?trusted_connection=yes&driver=SQL+Server'
engine = sqlalchemy. create_engine (url)

Mslb_Comp = pd.read_sql("""
                    SELECT 
                    SUBSTRING(MSLB.MATNR, PATINDEX('%[^0]%', MSLB.MATNR), LEN(MSLB.MATNR)) AS Componente,
                    ROUND(SUM (MSLB.LBLAB),2) as Saldo
                    FROM MSLB 
                    LEFT JOIN MARA 
                    ON MARA.MATNR = MSLB.MATNR 
                    WHERE MARA.MTART in ('ZINT','HALB','ZPC2','ZPC3','ZPC4','ZPC5','ZPC6','ZINT','DBBS','ZAR1','ROH','VERP')
                    GROUP BY MSLB.MATNR
                    HAVING SUM(MSLB.LBLAB) > 0
                    """,engine)

Mslb_Comp.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Porto_Comp.xlsx', index=None)




import sqlalchemy
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
servidor_dns = 'Servidor'
servidor_database = 'Base de dados'
url = f'mssql+pyodbc://@{servidor_dns}/{servidor_database}?trusted_connection=yes&driver=SQL+Server'
engine = sqlalchemy. create_engine (url)

Mslb_Fert = pd.read_sql("""
                    SELECT 
                    SUBSTRING(MSLB.MATNR, PATINDEX('%[^0]%', MSLB.MATNR), LEN(MSLB.MATNR)) AS Material,
                    ROUND(SUM (MSLB.LBLAB),2) as Saldo
                    FROM MSLB 
                    LEFT JOIN MARA 
                    ON MARA.MATNR = MSLB.MATNR 
                    WHERE MARA.MTART in ('FERT')
                    GROUP BY MSLB.MATNR
                    HAVING SUM(MSLB.LBLAB) > 0
                    """,engine)

Mslb_Fert.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Porto_Fert.xlsx', index=None)


import pandas as pd
import datetime 
import numpy as np

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
dados = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco144.xlsx',dtype = str)
dados = pd.DataFrame(dados, dtype=str)
material_zco144 = dados.loc[:,['Material']]
material_zco144 = material_zco144.drop_duplicates()
material_sqvi = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.xlsx' , dtype = str)
material_sqvi = material_sqvi.loc[(material_sqvi['GE'] != 'E') ]
material_sqvi = material_sqvi.loc[:,['Material']]
material_sqvi = material_sqvi.drop_duplicates()
material_sqvi = material_sqvi.merge(material_zco144, on='Material', how='outer', suffixes=['', '_'], indicator=True)
material_sqvi = material_sqvi.loc[(material_sqvi['_merge'] == 'left_only') ]
material_sqvi = material_sqvi.loc[:,['Material']]
dados = pd.concat([dados, material_sqvi])
dados = dados.fillna('')


zco32_estoque_componente = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco32_estoque_componente.xlsx', dtype = str)
dados = dados.merge(zco32_estoque_componente, on='Componente', how='left')
dados = dados.rename(columns={'Saldo': 'zco32_estoque_componente'})

Porto_Comp = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Porto_Comp.xlsx', dtype = str,)
dados = dados.merge(Porto_Comp, on='Componente', how='left')
dados = dados.rename(columns={'Saldo': 'Porto_Comp'})


Me2m_Compoente_Consultar_Pedido_De_Compra = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Me2m_Compoente_Consultar_Pedido_De_Compra.xlsx', dtype = str)
dados = dados.merge(Me2m_Compoente_Consultar_Pedido_De_Compra, on='Componente', how='left')
dados = dados.rename(columns={'Saldo': 'Me2m_Compoente_Consultar_Pedido_De_Compra'})

zco32_estoque_fert = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco32_estoque_fert.xlsx', dtype = str )
dados  = dados.merge(zco32_estoque_fert, on='Material', how='left')
dados = dados.rename(columns={'Saldo': 'zco32_estoque_fert'})

Porto_Fert = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Porto_Fert.xlsx', dtype = str,)
dados = dados.merge(Porto_Fert, on='Material', how='left')
dados = dados.rename(columns={'Saldo': 'Porto_Fert'})

Me2m_Fert_Consultar_Pedido_De_Compra = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Me2m_Fert_Consultar_Pedido_De_Compra.xlsx', dtype = str,)
dados = dados.merge(Me2m_Fert_Consultar_Pedido_De_Compra, on='Material', how='left')
dados = dados.rename(columns={'Saldo': 'Me2m_Fert_Consultar_Pedido_De_Compra'})


ZPP101_Fert_Ordem_Aberta = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ZPP101_Fert_Ordem_Aberta.xlsx', dtype = str)
dados  = dados.merge(ZPP101_Fert_Ordem_Aberta, on='Material', how='left')
dados = dados.rename(columns={'Saldo': 'ZPP101_Fert_Ordem_Aberta'})


ZSD138_carteira_fert = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ZSD138_carteira_fert.xlsx', dtype = str)
dados  = dados.merge(ZSD138_carteira_fert, on='Material', how='left')
dados = dados.rename(columns={'Saldo': 'ZSD138_carteira_fert'})

ZMM22_Fert_Total_Venda_8_Meses = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ZMM22_Fert_Total_Venda_8_Meses.xlsx', dtype = str,)
dados = dados.merge(ZMM22_Fert_Total_Venda_8_Meses, on='Material', how='left')
dados = dados.rename(columns={'Saldo': 'ZMM22_Fert_Total_Venda_8_Meses'})

ultima_venda = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ultima_venda.xlsx', dtype = str,)
dados = dados.merge(ultima_venda, on='Material', how='left')

Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao.xlsx' , dtype = str)
dados = dados.merge(Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao, on='Material', how='left')

dados = dados.fillna({'Criado': '14.04.2023'})

dados = dados.fillna({ 'Ult.Venda': '01.01.1999'})

data_day =  datetime.date.today()

dados["Criado"] = pd.to_datetime(dados["Criado"], dayfirst=True)

dados["Ult.Venda"] = pd.to_datetime(dados["Ult.Venda"], dayfirst=True)

dados['Meses_Cadastrado'] = (pd.to_datetime(data_day) - dados['Criado'])

dados['Meses_Ultima_Venda'] = (pd.to_datetime(data_day) - dados['Ult.Venda'])


dados['Criado'] = dados['Criado'].dt.strftime('%d/%m/%Y')

dados['Ult.Venda'] = dados['Ult.Venda'].dt.strftime('%d/%m/%Y')


dados['Meses_Cadastrado'] = ((dados['Meses_Cadastrado'] / np.timedelta64(1, 'D')).astype(int))
dados['Meses_Ultima_Venda'] = ((dados['Meses_Ultima_Venda'] / np.timedelta64(1, 'D')).astype(int))


dados['Meses_Cadastrado'] = dados['Meses_Cadastrado'] / 30
dados['Meses_Ultima_Venda'] = dados['Meses_Ultima_Venda'] / 30


dados = dados.rename(columns={  'ZMM22_Fert_Total_Venda_8_Meses' : 'Tot_Venda6Mês',
                        'zco32_estoque_fert' : 'Estoque_Fert', 
                        'ZPP101_Fert_Ordem_Aberta' : 'Orderm_Aberta',
                        'ZSD138_carteira_fert' : 'Cart_Aberta',
                        'Me2m_Fert_Consultar_Pedido_De_Compra' : 'Compra_Fert',
                        'Componente' : 'Comp.',
                        'Me2m_Compoente_Consultar_Pedido_De_Compra' : 'Compra_Comp',
                        'zco32_estoque_componente' : 'Estoque_Comp'
                        })


dados = dados.drop_duplicates()

dados = dados.fillna({ 
                'Estoque_Fert': 0, 
                'Compra_Fert': 0,
                'Orderm_Aberta': 0,
                'Cart_Aberta': 0, 
                'Porto_Fert': 0,
                'Tot_Venda6Mês': 0, 
                'Compra_Comp': 0,
                'Porto_Comp': 0, 
                'Estoque_Comp': 0,
                'Total': 0,
                'Média': 0,                    
                })



dados = dados.fillna('')


dados[[
'Estoque_Fert',
'Compra_Fert',
'Orderm_Aberta',
'Cart_Aberta',
'Porto_Fert',
'Compra_Comp',
'Porto_Comp',
'Estoque_Comp',
'Total',
'Média'
]] = dados[[
    'Estoque_Fert',
    'Compra_Fert',
    'Orderm_Aberta',
    'Cart_Aberta',
    'Porto_Fert',
    'Compra_Comp',
    'Porto_Comp',
    'Estoque_Comp',
    'Total',
    'Média'
    ]].apply(pd.to_numeric).round(4)


data_day =  datetime.date.today()

dia = data_day.day
mes = data_day.month
ano = data_day.year
name_excel = f'{ano}-{mes}-{dia}'



dados['Estoque_Comp_Soma'] = dados.groupby('Material')["Estoque_Comp"].transform(np.sum)
dados['Porto_Comp_Soma'] = dados.groupby('Material')["Porto_Comp"].transform(np.sum)
dados['Compra_Comp_Soma'] = dados.groupby('Material')["Compra_Comp"].transform(np.sum)



dados = dados.loc[(dados['GE'] != 'E') & (dados['Material'] != '9000SP') & (dados['Material'] != '9050SP')  & (dados['Material'] != '9100SP')  & (dados['Material'] != '9150SP') & (dados['Material'] != '9400SP') & (dados['Material'] != '9901SU') ]


dados['Fert_Estoq.Compra_Porto'] = dados['Estoque_Fert']  + dados['Compra_Fert'] + dados['Porto_Fert']
dados['Comp._Estoq_Compra_Porto'] = dados['Estoque_Comp_Soma']  + dados['Porto_Comp_Soma'] + dados['Compra_Comp_Soma']
dados = dados.drop_duplicates()

dados.drop(['Estoque_Comp_Soma'], axis=1, inplace=True)
dados.drop(['Porto_Comp_Soma'], axis=1, inplace=True)
dados.drop(['Compra_Comp_Soma'], axis=1, inplace=True)

def minha_funcao(row):
        #11 ordem aberta, #13 Cart_Aberta, #27 Total, #38Meses_Cadastrado, #39 Meses_Ultima_Venda, #40 estoque_fert, #41 estoque componente
    if (row[11]=='' and row[13]==0 and row[27]==0 and row[38]>12 and row[39]>12 and row[40]==0 and row[41]==0):
        return '(ELIMINAR)'
           
    #elif(row[38] <=12):
        #return '(CADASTRO NOVO) (<=12 MESES)'

    #elif(row[39] <=12 and row[27]!=0):
        #return '(ULTIMA VENDA RECENTE) (<=12 MESES)'


    #elif(row[28] == '01/01/1999' and row[27]==0 and row[40] >0 and row[41] == 0 and row[38] >12):
        #return '(NUNCA HOUVE VENDA) (FERT SALDO >0) (COMPONENTES SALDO ==0) (MATERIAL CADASTRADO >12 MESES) '


    #elif(row[28] == '01/01/1999' and row[27]==0 and row[40] ==0 and row[41] > 0 and row[38] >12):
        #return '(NUNCA HOUVE VENDA) (FERT SALDO ==0) (COMPONENTES SALDO >0) (MATERIAL CADASTRADO >12 MESES) '

  
    #elif(row[28] == '01/01/1999' and row[27]==0 and row[40] ==0 and row[41]==0 and row[38] >12):
        #return '(NUNCA HOUVE VENDA) (FERT SALDO ==0) (COMPONENTES SALDO ==0) (MATERIAL CADASTRADO >12 MESES) '


    #elif(row[28] == '01/01/1999' and row[27]==0 and row[40] >0 and row[41] > 0 and row[38] >12):
        #return '(NUNCA HOUVE VENDA) (FERT SALDO >0) (COMPONENTES SALDO >0) (MATERIAL CADASTRADO >12 MESES) '

    #elif(row[40]>0 and row[41]==0 and row[39]>12 and row[27] ==0 and row[38] >12):
        #return '(NÃO HOUVE VENDA ÚLTIMOS 12 MESES) (FERT SALDO >0) (COMPONENTES SALDO ==0) (CADASTRO >12 MESES)'

    #elif(row[40]==0 and row[41]>0 and row[39]>12 and row[27] ==0 and row[38] >12):
        #return '(NÃO HOUVE VENDA ÚLTIMOS 12 MESES) (FERT SALDO ==0) (COMPONENTES SALDO >0) (CADASTRO >12 MESES)'


    #elif(row[40]==0 and row[41] ==0 and row[39]>12 and row[27] ==0 and row[38] >12):
        #return '(NÃO HOUVE VENDA ÚLTIMOS 12 MESES) (FERT SALDO ==0) (COMPONENTES SALDO ==0) (CADASTRO >12 MESES)'

    #elif(row[40]>0 and row[41] >0 and row[39]>12 and row[27] ==0 and row[38] >12):
        #return '(NÃO HOUVE VENDA ÚLTIMOS 12 MESES) (FERT SALDO >0) (COMPONENTES SALDO >0) (CADASTRO >12 MESES)'



dados.insert(1,"Status",dados.apply(minha_funcao, axis=1) ,True)

dados['Ult.Venda'] = dados['Ult.Venda'].apply(lambda x: str(x.replace("01/01/1999","")))
dados['Meses_Ultima_Venda'] = dados['Meses_Ultima_Venda'].apply(str)
dados['Meses_Ultima_Venda'] = dados['Meses_Ultima_Venda'].apply(lambda x: str(x.replace("296.46666666666664","")))
dados = dados.rename(columns={'Porto_Comp': 'Fornecedor_Componente','Porto_Fert': 'Fornecedor_Fert'})
dados.to_excel(f'U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Jupyter\\{name_excel} - Analise Final DD.xlsx', index = None)





import pandas as pd
import win32com.client
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import subprocess
from time import sleep
import sys  
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


data_day =  datetime.date.today()

dia = data_day.day
mes = data_day.month
ano = data_day.year
name_excel = f'{ano}-{mes}-{dia}'
  
def enviar_email1():


    data_day =  datetime.date.today()

    dia = data_day.day
    mes = data_day.month
    ano = data_day.year
    name_excel = f'{ano}-{mes}-{dia}'
            
    remetente = 'email'
    senha_rede = '******' # Colocar aqui a senha do e-mail

    destinatario7 = 'destinario@.com.br'
    destinatario2 = 'destinario@.com.br'
    destinatario3 = 'destinario@.com.br'
    destinatario4 = 'destinario@.com.br'
    destinatario5 = 'destinario@.com.br'
    destinatario6 = 'destinario@.com.br'
    destinatario8 = 'destinario@.com.br'
    destinatario1 = 'destinario@.com.br'

    assunto = 'Relatório Dedo Duro'
    # Preenche abaixo o corpo da mensagem.
    texto = f"""



    "Bom dia, segue relatório" 
        

    OBS: MENSAGEM AUTOMÁTICA.

    """
    email_sender = remetente
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario1
    msg['To'] = destinatario2
    msg['To'] = destinatario3
    msg['To'] = destinatario4
    msg['To'] = destinatario5
    msg['To'] = destinatario6
    msg['To'] = destinatario7
    msg['To'] = destinatario8

    msg['Subject'] = assunto


    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(f"{name_excel} - Analise Final DD.xlsx", "rb").read())
    encoders.encode_base64(part)


    part.add_header('Content-Disposition', 'attachment', filename=f"{name_excel} - Analise Final DD.xlsx")
    msg.attach(part)

    msg.attach(MIMEText(_text=texto.encode('utf-8'), _charset='utf-8'))
    port = *** if '#empresa' in destinatario1 else 25
    server = smtplib.SMTP(host='smtp.office365.com', port=port)
    server.ehlo()
    server.starttls()
    server.login(remetente, senha_rede)
    text = msg.as_string()
    server.sendmail(email_sender, destinatario1, text)
    server.sendmail(email_sender, destinatario2, text)
    server.sendmail(email_sender, destinatario3, text)
    server.sendmail(email_sender, destinatario4, text)
    server.sendmail(email_sender, destinatario5, text)
    server.sendmail(email_sender, destinatario6, text)
    server.sendmail(email_sender, destinatario7, text)
    server.sendmail(email_sender, destinatario8, text) 
    print('Email enviado')
    server.quit()        
enviar_email1()
