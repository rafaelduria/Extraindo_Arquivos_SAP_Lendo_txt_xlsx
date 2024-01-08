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
from email import encoders
import sys



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
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"        
        subprocess.Popen(self.path)
        aberto = False
        sleep(1)
        while not aberto:
            sleep(1)
            try:
                sleep(1)
                self.SapGuiauto = win32com.client.GetObject('SAPGUI')
                application  = self.SapGuiauto.GetScriptingEngine
                self.connection = application.OpenConnection('sessaoSAP')
                aberto = True
            except:
                pass
        connections = len(self.connection.Children)
        self.session = self.connection.Children(connections-1)
        self.session.findById("wnd[0]").maximize

    def SapLogin(self):
        try:
            self.session.findById("wnd[0]").maximize      
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "login"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "senha"
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
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me2m.txt"
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
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me2m.txt"
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
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me2m.txt"
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
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me2m.txt"
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
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me2m.txt"
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
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me2m.txt"
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
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "sqvi.txt"
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
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "sqvi.txt"
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
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "sqvi.txt"
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
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "sqvi.txt"
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
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "sqvi.txt"
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
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "sqvi.txt"
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
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ordem_aberta.txt"
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
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ordem_aberta.txt"
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
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ordem_aberta.txt"
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
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ordem_aberta.txt"
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
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ordem_aberta.txt"
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
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ordem_aberta.txt"
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
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "carteira.txt"
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
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "carteira.txt"
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
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "carteira.txt"
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
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "carteira.txt"
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
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "carteira.txt"
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
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "carteira.txt"
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
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "venda_12_meses.txt"
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
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "venda_12_meses.txt"
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
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "venda_12_meses.txt"
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
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "venda_12_meses.txt"
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
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "venda_12_meses.txt"
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
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "venda_12_meses.txt"
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



import pandas as pd
import datetime
import numpy as np

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)



def limpar_txt(linhas):  
    linhas = [x for x in linhas if '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
    linhas = [x for x in linhas if '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
    linhas = [x for x in linhas if '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
    linhas = [x for x in linhas if '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]    
    linhas = [x for x in linhas if '-------------------------' not in x]
    linhas = [x for x in linhas if '--------------------------' not in x]
    linhas = [x for x in linhas if '---------------------------' not in x]
    linhas = [x for x in linhas if '----------------------------' not in x]
    linhas = [x for x in linhas if '-----------------------------' not in x]
    linhas = [x for x in linhas if '-------------------------------' not in x]
    linhas = [x for x in linhas if '---------------------------------' not in x]
    linhas = [x for x in linhas if '-----------------------------------' not in x]
    linhas = [x for x in linhas if '-----------------------------------------' not in x]
    linhas = [x for x in linhas if '-------------------------------------------' not in x]
    linhas = [x for x in linhas if '--------------------------------------------' not in x]
    linhas = [x for x in linhas if '----------------------------------------------' not in x]
    linhas = [x for x in linhas if '--------------------------------------------------------------------------------' not in x]
    linhas = [x for x in linhas if '*' not in x]
    linhas = [x for x in linhas if '|Critérios de ordenação|Cresc.|Decr.|Subtotal|' not in x]
    linhas = [x for x in linhas if '|Material              |  X   |     |        |' not in x]
    linhas = [x for x in linhas if 'Data Emissão' not in x]
    linhas = [x for x in linhas if 'Emissor  ' not in x]
    linhas = [x for x in linhas if 'Estatíst.dados' not in x]
    linhas = [x for x in linhas if 'Gerado em' not in x]
    linhas = [x for x in linhas if 'Histórico de consumo dos últimos 12 meses ' not in x]
    linhas = [x for x in linhas if 'Lista tecnica multi-nível' not in x]
    linhas = [x for x in linhas if 'Lnhs.totais determinadas' not in x]
    linhas = [x for x in linhas if 'Registros processados:' not in x]
    linhas = [x for x in linhas if 'Registros transfs.' not in x]    
    linhas = [x for x in linhas if 'Sem dados' not in x]
    linhas = [x for x in linhas if x[0][0]=='|']
    linhas = [x.split('|') for x in linhas]
    linhas = [x[1:-1] for x in linhas]

    dados_limpados = []
    for linha in linhas:     
        dados_temp = []
        for dado in linha:
            dado = dado.rstrip().lstrip()
            dados_temp.append(dado)
        dados_limpados.append(dados_temp)

    dados_limpados = pd.DataFrame(dados_limpados)
    dados_limpados = dados_limpados.rename(columns=dados_limpados.iloc[0]).drop(dados_limpados.index[0])
    dados_limpados = dados_limpados.query('Material != "Material"')
    dados_limpados = dados_limpados.query('Material != "" ')
    return dados_limpados





with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\ZCO144.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

zco144 = limpar_txt(linhas)

zco144 = zco144.rename(columns={'Comp. secundário': 'Comp','Tipo mat.sec.': 'Tipo','MEINS_4': 'UN' })
zco144 = zco144.loc[(zco144['Tipo'] != 'VERP') & (zco144['Tipo'] != 'ROH')]

zco144 = zco144.sort_values(by=['Material','Nível','Sentido'])
zco144.drop(['Nível','Sentido','Tipo'],axis=1,inplace=True)
zco144 = zco144.drop_duplicates()
zco144.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco144.xlsx',index=None)




with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\sqvi.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

sqvi = limpar_txt(linhas)

sqvi = sqvi.rename(columns={
'TxtBreveMaterial': 'Desc',      
'SM': 'Bloq Prod',
'Denominação': 'Desc',
'St': 'Bloq Venda',
'Denominação': 'Desc',
'EM': 'GE',                                                   
})

sqvi["Criado"] = pd.to_datetime(sqvi["Criado"], dayfirst=True)

sqvi['Meses Cadastrado'] = (pd.to_datetime(datetime.date.today()) - sqvi['Criado'])

sqvi['Criado'] = sqvi['Criado'].dt.strftime('%d/%m/%Y')

sqvi['Meses Cadastrado'] = ((sqvi['Meses Cadastrado'] / np.timedelta64(1, 'D')).astype(int))

sqvi['Meses Cadastrado'] = (sqvi['Meses Cadastrado'] / 30).round()

sqvi = sqvi.iloc[:,[0,6,7,4,5,8,9,10,3]]

sqvi.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\sqvi.xlsx',index=None)


with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\ultima_venda.txt', encoding='latin1'  ) as f:
    linhas = f.readlines()
    f.close()

ultima_venda = limpar_txt(linhas)

ultima_venda = ultima_venda.rename(columns={'Dt.lçto.': 'Ultima_Venda'})

ultima_venda["Ultima_Venda"] = pd.to_datetime(ultima_venda["Ultima_Venda"], dayfirst=True)

ultima_venda['Meses Ult Venda'] = (pd.to_datetime(datetime.date.today()) - ultima_venda['Ultima_Venda'])
ultima_venda['Ultima_Venda'] = ultima_venda['Ultima_Venda'].dt.strftime('%d/%m/%Y')

ultima_venda['Meses Ult Venda'] = ((ultima_venda['Meses Ult Venda'] / np.timedelta64(1, 'D')).astype(int))

ultima_venda['Meses Ult Venda'] = (ultima_venda['Meses Ult Venda'] / 30).round()

ultima_venda = ultima_venda.drop(['Ultima_Venda'], axis='columns')

ultima_venda.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ultima_venda.xlsx' , index = None)




with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\venda_12_meses.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

venda_12_meses = limpar_txt(linhas)

venda_12_meses = venda_12_meses.drop(['Descrição','Unidade'], axis='columns')
venda_12_meses = venda_12_meses.rename(columns={'Dt.lçto.': 'Ultima_Venda', 'Total':'Venda 12 Meses'})
venda_12_meses = venda_12_meses.iloc[:,[0,14,13,12,11,10,9,8,7,6,5,4,3,2]]
#Tirar zero 0
n = list(venda_12_meses.columns)
n1 = n[1]
n2 = n[2]
n3 = n[3]
n4 = n[4]
n5 = n[5]
n6 = n[6]
n7 = n[7]
n8 = n[8]
n9 = n[9]
n10 = n[10]
n11 = n[11]
n12 = n[12]
n13 = n[13]
venda_12_meses[n1] = venda_12_meses[n1].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n2] = venda_12_meses[n2].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n3] = venda_12_meses[n3].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n4] = venda_12_meses[n4].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n5] = venda_12_meses[n5].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n6] = venda_12_meses[n6].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n7] = venda_12_meses[n7].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n8] = venda_12_meses[n8].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n9] = venda_12_meses[n9].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n10] = venda_12_meses[n10].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n11] = venda_12_meses[n11].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n12] = venda_12_meses[n12].apply(lambda x: "" if x == "0" else x)
venda_12_meses[n13] = venda_12_meses[n13].apply(lambda x: "" if x == "0" else x)
venda_12_meses.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\venda_12_meses.xlsx', index = None)



with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\carteira.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

carteira = limpar_txt(linhas)

carteira = carteira.rename(columns={'Pend.Fornecer': 'Cart Venda'})
carteira['Cart Venda'] = carteira['Cart Venda'].apply(lambda x: float(x.replace(".","").replace(",",".")))
carteira = carteira.groupby(['Material','UMB'])[['Cart Venda']].sum().round(3).reset_index()[['Material','Cart Venda','UMB']]
carteira.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\carteira.xlsx', index=None)


with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\ordem_aberta.txt', encoding='latin1') as f:
    linhas = f.readlines()
    f.close()

ordem_aberta = limpar_txt(linhas)

def ordem_func(row):
    if (row['Ordem']!=''):
        return 'Sim'
ordem_aberta.insert(1,"Ordem Aberta",ordem_aberta.apply(ordem_func, axis=1) ,True)
ordem_aberta = ordem_aberta.loc[:,['Material','Ordem Aberta','Cob.+ OP']]
ordem_aberta.drop_duplicates(inplace=True)
ordem_aberta.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ordem_aberta.xlsx' , index = None)



with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\me2m.txt', encoding='latin1'  ) as f:
    linhas = f.readlines()
    f.close()   

me2m = limpar_txt(linhas)

me2m = me2m[['Material','a fornecer','UGE']]
me2m = me2m.rename(columns={'UGE': 'UPP'})

me2m = me2m.iloc[:,[0,1,3]]
me2m = me2m.rename(columns={'a fornecer':'Ped Compra Fert'})
me2m['Ped Compra Fert'] = me2m['Ped Compra Fert'].apply(lambda x: float(x.replace(".","").replace(",",".")))
me2m = me2m.groupby(['Material','UPP'])['Ped Compra Fert'].sum().round(3).apply(lambda x: "{:_.2f}".format(x).replace('.', ',').replace('_', '.')).reset_index()[['Material','Ped Compra Fert','UPP']]
me2m.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\me2m.xlsx', index = None)



import pandas as pd
import datetime 
import numpy as np
import sqlalchemy
from sqlalchemy import create_engine

servidor_dns = 'servidor'
servidor_database = 'inteligcom'
url = f'mssql+pyodbc://@{servidor_dns}/{servidor_database}?trusted_connection=yes&driver=SQL+Server'
engine = sqlalchemy. create_engine (url)


#Consulta Todos produtos "FERT" pois ZCO144 não traz itens revenda ou itens que não possui lista tecnica.
todos_fert_sql = pd.read_sql(
"""
SELECT 
DISTINCT
SUBSTRING(MARA.MATNR, PATINDEX('%[^0]%', MARA.MATNR), LEN(MARA.MATNR)) AS Material,
MARA.MEINS as Un
FROM MARA
WHERE MARA.MTART = 'FERT' AND MARA.LVORM <> 'X'
"""
,engine)


#Copia todos códifo fert com unidade de medida basica
todos_fert_sql_Un = todos_fert_sql.copy()


#Drop #eliminar coluna unidade "Un"
todos_fert_sql = todos_fert_sql.drop(['Un'], axis='columns')



#lista de materiais zco144
fert_zco144 = dados[['Material']]
#verificando juntação lista de materiais zco144 todos fert sql
todos_fert_sql = todos_fert_sql.merge(fert_zco144, on='Material', how='outer', suffixes=['', '_'], indicator=True)


#filtro apenas com materiais faltantes
todos_fert_sql = todos_fert_sql.loc[(todos_fert_sql['_merge'] == 'left_only')]
#deixando apenas coluna material
todos_fert_sql = todos_fert_sql[['Material']]


#colocando lista de materiais faltantes debaixo do dataframe dados
dados = pd.concat([dados, todos_fert_sql])


#estoque
Estoque_Sql = pd.read_sql(
"""
SELECT
DISTINCT
SUBSTRING(MARA.MATNR, PATINDEX('%[^0]%', MARA.MATNR), LEN(MARA.MATNR)) AS Material,
COALESCE((MARD.LABST),0) + COALESCE((MARD.SPEME),0) + COALESCE((MARD.UMLME),0) + COALESCE((MARC.UMLMC),0) + COALESCE((MSLB.LBLAB),0) + COALESCE((MARD.INSME),0) as Estoq
FROM MARA LEFT JOIN MARD ON MARD.MATNR = MARA.MATNR 
LEFT JOIN MARC  ON MARA.MATNR = MARC.MATNR
LEFT JOIN MSLB  ON MARA.MATNR = MSLB.MATNR
WHERE MARA.MTART in ('ZAR1','ZPC5','ROH','ZPC4','ZPC2','ZINT','FERT','ZPC3','HALB','ZAR3','ZAR2','DBBS','ZPC6') AND COALESCE((MARD.LABST),0) + COALESCE((MARD.SPEME),0) + COALESCE((MARD.INSME),0) + COALESCE((MARD.UMLME),0) + COALESCE((MARC.UMLMC),0) + COALESCE((MSLB.LBLAB),0) >0 AND (MARA.LVORM <> 'X') 
""",engine)


#Estoque_Sql['Estoq'] = Estoque_Sql.groupby('Material')['Estoq'].transform(sum)
Estoque_Sql = Estoque_Sql.groupby(['Material'])['Estoq'].sum().reset_index()


#Criar copia de Estoque
Estoque_Sql_Comp = Estoque_Sql.copy()
#Renomear coluna para fazer Merge
Estoque_Sql_Comp = Estoque_Sql_Comp.rename(columns={'Material': 'Comp', 'Estoq':'Estoq Comp'})
Estoque_Sql_Comp.head()


#Merge dados com Unidade de medida basica
dados = dados.merge(todos_fert_sql_Un, on='Material',how='left')
#Atualizando nomes de TH para MIL de PAK para PAC
dados['Un'] = dados['Un'].apply(lambda x: "MIL" if x == "TH" else x)
dados['Un'] = dados['Un'].apply(lambda x: "PAC" if x == "PAK" else x)
dados['Un'] = dados['Un'].apply(lambda x: "PEC" if x == "ST" else x)
dados['Un'] = dados['Un'].apply(lambda x: "PAR" if x == "PAA" else x)


#Merge dados com Estoque Material
dados = dados.merge(Estoque_Sql.drop_duplicates('Material'),how='left',on='Material')
dados.head()


#Reordenando colunas
dados = dados[['Material','Un','Estoq','Comp','UN']]
dados.head()


#Merge dados com Estoque Comp
dados = dados.merge(Estoque_Sql_Comp.drop_duplicates('Comp'),how='left',on='Comp')
dados.head()



ultima_venda = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ultima_venda.xlsx', dtype = str)
dados = dados.merge(ultima_venda, on='Material', how='left')
dados = dados.fillna({'Meses Ult Venda':9999})
dados[['Meses Ult Venda']] = dados[['Meses Ult Venda']].apply(pd.to_numeric)

venda_12_meses = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\venda_12_meses.xlsx', dtype = str,)
dados = dados.merge(venda_12_meses, on='Material', how='left')

sqvi = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\sqvi.xlsx' , dtype = str)
dados = dados.merge(sqvi, on='Material', how='left')
dados = dados.fillna({'Meses Cadastrado':0})
dados[['Meses Cadastrado']] = dados[['Meses Cadastrado']].apply(pd.to_numeric)

dados = dados.loc[(dados['GE'] != 'E') & (dados['Material'] != '9000SP') & (dados['Material'] != '9050SP')  & (dados['Material'] != '9100SP')  & (dados['Material'] != '9150SP') & (dados['Material'] != '9400SP') & (dados['Material'] != '9901SU') ]

carteira = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\carteira.xlsx', dtype = str)
dados  = dados.merge(carteira, on='Material', how='left')

ordem_aberta = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ordem_aberta.xlsx', dtype = str)
dados  = dados.merge(ordem_aberta, on='Material', how='left')
dados = dados.fillna({'Ordem Aberta':"Não"})

me2m = pd.read_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\me2m.xlsx', dtype = str)
dados = pd.merge(dados,me2m , right_on='Material', left_on='Material', how='left')

me2m_comp = me2m.copy()
me2m_comp = me2m_comp.rename(columns={'Material': 'Comp', 'Ped Compra Fert':'Ped Compra Comp', 'UPP':'Upp'})
dados  = dados.merge(me2m_comp, on='Comp', how='left')


#Ajustando descrições das colunas
dados = dados.rename(columns={'Desc': 'Desc GE', 'Desc.1':'Desc bloqueio prod', 'Desc.2':'Desc bloqueio venda', 'Desc.3':'Desc produto acabado'})



#Criando uma coluna  "Soma componentes"
Soma_componentes = dados.groupby(['Material'])[['Estoq Comp']].sum().reset_index().copy()



Soma_componentes = Soma_componentes.rename(columns={'Estoq Comp': 'Soma componentes'})

#dados  = dados.merge(Soma_componentes, on='Material', how='left')
dados = dados.merge(Soma_componentes.drop_duplicates('Material'),how='left',on='Material')

#Agrupar Estoque e tranformar valores númericos para formato brasileiro 000.000,00
dados['Estoq'] = dados['Estoq'].apply(lambda x: "{:_.3f}".format(x).replace('.', ',').replace('_', '.'))
dados['Estoq Comp'] = dados['Estoq Comp'].apply(lambda x: "{:_.3f}".format(x).replace('.', ',').replace('_', '.'))

#dados['EstoqQualidade'] = dados['EstoqQualidade'].apply(lambda x: "{:_.3f}".format(x).replace('.', ',').replace('_', '.'))
#dados['EstoqQualidade Comp'] = dados['EstoqQualidade Comp'].apply(lambda x: "{:_.3f}".format(x).replace('.', ',').replace('_', '.'))

#Deixando todos NULL para branco
dados = dados.fillna('')
dados['Estoq Comp'] = dados['Estoq Comp'].apply(lambda x: "" if x == "nan" else x)
dados['Estoq'] = dados['Estoq'].apply(lambda x: "" if x == "nan" else x)

#dados['EstoqQualidade'] = dados['EstoqQualidade'].apply(lambda x: "" if x == "nan" else x)
#dados['EstoqQualidade Comp'] = dados['EstoqQualidade Comp'].apply(lambda x: "" if x == "nan" else x)


def minha_funcao(row):
    if(row['Ordem Aberta']=='Não') and (row['Cart Venda']=='') and  (row['Venda 12 Meses']=='') and  (row['Meses Cadastrado']>12) and  (row['Estoq']=='') and  (row['Soma componentes']==0.0) and (row['Ped Compra Fert']=='') and (row['Ped Compra Comp']==''):
        return 'Eliminar'

    elif(row['Meses Cadastrado'] <=12):
        return 'Cadastro novo menor que 12 meses'

    elif(row['Meses Ult Venda'] <=12):
        return 'Última venda recente menor que 12 meses'


dados.insert(1,"Status",dados.apply(minha_funcao, axis=1) ,True)


dados = dados.drop(['Soma componentes'], axis='columns')

dados['Meses Ult Venda'] = dados['Meses Ult Venda'].apply(lambda x: "Nunca houve venda" if x == 9999 else x)    


name_excel = f'{(datetime.date.today()).year}-{(datetime.date.today()).month}-{(datetime.date.today()).day}'
dados.to_excel(f'U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Jupyter\\{name_excel} - Analise Final DD.xlsx',index=None, freeze_panes= (1,1))



def enviar_email1():


    data_day =  datetime.date.today()

    dia = data_day.day
    mes = data_day.month
    ano = data_day.year
    name_excel = f'{ano}-{mes}-{dia}'
            
    remetente = 'email@.com.br'
    senha_rede = 'senha**' # Colocar aqui a senha 

    destinatario1 = 'destinatario1@.com.br'
    assunto = 'aqui vai o assunto'
    # Preenche abaixo o corpo da mensagem.
    texto = f"""



    "Bom dia, segue relatório" 
        

    OBS: MENSAGEM AUTOMÁTICA.

    """
    email_sender = remetente
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario1
    msg['Subject'] = assunto


    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(f"{name_excel} - Analise Final DD.xlsx", "rb").read())
    encoders.encode_base64(part)


    part.add_header('Content-Disposition', 'attachment', filename=f"{name_excel} - Analise Final DD.xlsx")
    msg.attach(part)

    msg.attach(MIMEText(_text=texto.encode('utf-8'), _charset='utf-8'))
    port = 587 if 'empresa' in destinatario1 else 25
    server = smtplib.SMTP(host='smtp.office365.com', port=port)
    server.ehlo()
    server.starttls()
    server.login(remetente, senha_rede)
    text = msg.as_string()
    server.sendmail(email_sender, destinatario1, text)
    print('Email enviado')
    server.quit()        
enviar_email1()
