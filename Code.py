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
import numpy as np


def play():

    data_day =  datetime.date.today()
    dia = data_day.day
    mes = data_day.month
    ano = data_day.year
    hoje = (f'{dia}.{mes}.{ano}')
    um_dia_atraz = (f'{dia-1}.{mes}.{ano}')
    
    dois_dias_atras = (f'{dia-2}.{mes}.{ano}')
    tres_dias_atras = (f'{dia-3}.{mes}.{ano}')
    quatro_dias_atras = (f'{dia-4}.{mes}.{ano}')
    cinco_dias_atras = (f'{dia-5}.{mes}.{ano}')




    def Zjob():
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
                    self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "AQUI_VAI_LOGIN_SAP"
                    self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "AQUI_VAI_SENHA_SAP"
                    self.session.findById("wnd[0]").sendVKey(0)
                except:
                    print(sys.exc_info()[0])

        SapGui().SapLogin()  


        def Zjob_todos():      
            data_day =  datetime.date.today()
            dia = data_day.day
            mes = data_day.month
            ano = data_day.year
            hoje = (f'{dia}.{mes}.{ano}')
            um_dia_atraz = (f'{dia-1}.{mes}.{ano}')
            dois_dias_atras = (f'{dia-2}.{mes}.{ano}')
            tres_dias_atras = (f'{dia-3}.{mes}.{ano}')
            quatro_dias_atras = (f'{dia-4}.{mes}.{ano}')
            cinco_dias_atras = (f'{dia-5}.{mes}.{ano}')
                
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            conn = application.Children(0)  
            session = conn.Children(0)


            def zjob_entra():
                session.findById("wnd[0]/tbar[0]/okcd").text = "zjob"
                session.findById("wnd[0]").sendVKey(0)
            zjob_entra()
            

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



            def PORTO():
                def me2m ():
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "PORTO"
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

                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "porto.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()

                def voltar():
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)  
                voltar()
            PORTO()



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
        Zjob_todos() 




        def Pegar_Estoque_Fert():
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            conn = application.Children(0)  
            session = conn.Children(0)

            data_day =  datetime.date.today()
            dia = data_day.day
            mes = data_day.month
            ano = data_day.year
            hoje = (f'{dia}.{mes}.{ano}')
            um_dia_atraz = (f'{dia-1}.{mes}.{ano}')
            
            dois_dias_atras = (f'{dia-2}.{mes}.{ano}')
            tres_dias_atras = (f'{dia-3}.{mes}.{ano}')
            quatro_dias_atras = (f'{dia-4}.{mes}.{ano}')
            cinco_dias_atras = (f'{dia-5}.{mes}.{ano}')

            try:
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_FERT"
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_fert.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(3) 

            except:
                try:
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_FERT"
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_fert.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]").sendVKey(3)

                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0)                
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_FERT"
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_fert.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        session.findById("wnd[0]").sendVKey(3)
                        session.findById("wnd[0]").sendVKey(3)

                    except:
                        try:
                            session.findById("wnd[0]").sendVKey(0)
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_FERT"
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_fert.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            session.findById("wnd[0]").sendVKey(3)
                            session.findById("wnd[0]").sendVKey(3)
                     
                        except:
                            try:
                                session.findById("wnd[0]").sendVKey(0)
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_FERT"
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_fert.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                session.findById("wnd[0]").sendVKey(3)
                                session.findById("wnd[0]").sendVKey(3)
                            except:
                                try:
                                    session.findById("wnd[0]").sendVKey(0)
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_FERT"
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_fert.txt"
                                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                    session.findById("wnd[0]").sendVKey(3)
                                    session.findById("wnd[0]").sendVKey(3)
                  
                                except:
                                    print('erro')
                                    session.findById("wnd[1]").sendVKey(0)                            
        Pegar_Estoque_Fert()



        def Pegar_Estoque_ZINT():
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            conn = application.Children(0)  
            session = conn.Children(0)

            data_day =  datetime.date.today()
            dia = data_day.day
            mes = data_day.month
            ano = data_day.year
            hoje = (f'{dia}.{mes}.{ano}')
            um_dia_atraz = (f'{dia-1}.{mes}.{ano}')
            
            dois_dias_atras = (f'{dia-2}.{mes}.{ano}')
            tres_dias_atras = (f'{dia-3}.{mes}.{ano}')
            quatro_dias_atras = (f'{dia-4}.{mes}.{ano}')
            cinco_dias_atras = (f'{dia-5}.{mes}.{ano}')

            try:
                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_COMP"
                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{hoje}"
                session.findById("wnd[0]").sendVKey(8)
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_componente.txt"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]").close()
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()




            except:
                try:
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_COMP"
                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{um_dia_atraz}"
                    session.findById("wnd[0]").sendVKey(8)
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_componente.txt"
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]").close()
                    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
   
                except:
                    try:
                        session.findById("wnd[0]").sendVKey(0)                
                        session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_COMP"
                        session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{dois_dias_atras}"
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_componente.txt"
                        session.findById("wnd[1]/tbar[0]/btn[11]").press()
                        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                        session.findById("wnd[0]").sendVKey(0)
                        session.findById("wnd[0]").close()
                        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
       
                    except:
                        try:
                            session.findById("wnd[0]").sendVKey(0)
                            session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_COMP"
                            session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{tres_dias_atras}"
                            session.findById("wnd[0]").sendVKey(8)
                            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_componente.txt"
                            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                            session.findById("wnd[0]").sendVKey(0)
                            session.findById("wnd[0]").close()
                            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

   
                        except:
                            try:
                                session.findById("wnd[0]").sendVKey(0)
                                session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_COMP"
                                session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{quatro_dias_atras}"
                                session.findById("wnd[0]").sendVKey(8)
                                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_componente.txt"
                                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                                session.findById("wnd[0]").sendVKey(0)
                                session.findById("wnd[0]").close()
                                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
  
                            except:
                                try:
                                    session.findById("wnd[0]").sendVKey(0)
                                    session.findById("wnd[0]/usr/txtS_JOBN-LOW").text = "ESTOQUE_COMP"
                                    session.findById("wnd[0]/usr/ctxtS_DATA-LOW").text = f"{cinco_dias_atras}"
                                    session.findById("wnd[0]").sendVKey(8)
                                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
                                    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt"
                                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zco32_estoque_componente.txt"
                                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                                    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                                    session.findById("wnd[0]").sendVKey(0)
                                    session.findById("wnd[0]").close()
                                    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
   
                                except:
                                    print('erro')
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                                    session.findById("wnd[0]").sendVKey(0)
                                    session.findById("wnd[0]").close()
                                    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

        Pegar_Estoque_ZINT()
    Zjob()

    def Read_Txt_Txt():

        def Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao():

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
        Sqvi_Bloqueio_Produção_Venda_GE_Data_Criacao()


        def zco32_estoque_fert():
            with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\zco32_estoque_fert.txt', encoding='latin1') as f:
                linhas = f.readlines()
                f.close()

            linhas = [x for x in linhas if 'Sem dados' not in x]
            linhas = [x for x in linhas if 'Material' not in x]
            linhas = [x for x in linhas if '--------------------------------------------------------------------------------' not in x]
            linhas = [x for x in linhas if '--------------------------' not in x]
            linhas = [x for x in linhas if x[0][0]=='|']
            linhas = [x.split('|') for x in linhas]
            linhas = [x[1:-1] for x in linhas]

            zco32_estoque_fert = []
            for linha in linhas:     
                dados_temp = []
                for dado in linha:
                    dado = dado.rstrip().lstrip()
                    dados_temp.append(dado)
                zco32_estoque_fert.append(dados_temp)


            estrutura_df = {
                'Material': str,
                'Saldo': str
            }


            zco32_estoque_fert = pd.DataFrame(zco32_estoque_fert, dtype=str)
            zco32_estoque_fert.columns = estrutura_df.keys()
            zco32_estoque_fert = zco32_estoque_fert.astype(estrutura_df)

            zco32_estoque_fert['Saldo'] = zco32_estoque_fert['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
            zco32_estoque_fert = zco32_estoque_fert.groupby("Material")["Saldo"].sum().round(3)
            zco32_estoque_fert.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco32_estoque_fert.xlsx')
        zco32_estoque_fert()


        def zco32_estoque_componente():

            with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\zco32_estoque_componente.txt', encoding='latin1') as f:
                linhas = f.readlines()
                f.close()

            linhas = [x for x in linhas if 'Sem dados' not in x]
            linhas = [x for x in linhas if 'Material' not in x]
            linhas = [x for x in linhas if 'Análise de Estoques' not in x]
            linhas = [x for x in linhas if '------------------------------' not in x]
            linhas = [x for x in linhas if '----------------------------' not in x]
            linhas = [x for x in linhas if 'Emitido em' not in x]
            linhas = [x for x in linhas if 'Tipo Material:' not in x]
            linhas = [x for x in linhas if 'Gr.Estatístico:Todos' not in x]
            linhas = [x for x in linhas if x[0][0]=='|']
            linhas = [x.split('|') for x in linhas]
            linhas = [x[1:-1] for x in linhas]
            zco32_estoque_componente = []
            for linha in linhas:     
                dados_temp = []
                for dado in linha:
                    dado = dado.rstrip().lstrip()
                    dados_temp.append(dado)
                zco32_estoque_componente.append(dados_temp)


            estrutura_df = {
                'Componente': str,
                'Saldo': str,
            }

            zco32_estoque_componente = pd.DataFrame(zco32_estoque_componente, dtype=str)
            zco32_estoque_componente.columns = estrutura_df.keys()
            zco32_estoque_componente = zco32_estoque_componente.astype(estrutura_df)
            zco32_estoque_componente = zco32_estoque_componente.drop_duplicates()
            zco32_estoque_componente['Saldo'] = zco32_estoque_componente['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
            zco32_estoque_componente = zco32_estoque_componente.loc[(zco32_estoque_componente['Saldo'] > 0)]
            zco32_estoque_componente.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\zco32_estoque_componente.xlsx' , index = None) 
        zco32_estoque_componente()



        def ZMM22_Fert_Total_Venda_8_Meses():
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

            ZMM22_Fert_Total_Venda_8_Meses = ZMM22_Fert_Total_Venda_8_Meses.iloc[:NumeroDeLinhas, [0,2,3,4,5,6,7,8,10]]


            ZMM22_Fert_Total_Venda_8_Meses.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ZMM22_Fert_Total_Venda_8_Meses.xlsx', index = None)
        ZMM22_Fert_Total_Venda_8_Meses()




        def ZPP101_Fert_Ordem_Aberta():
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
            ZPP101_Fert_Ordem_Aberta[[    'Saldo']] = ZPP101_Fert_Ordem_Aberta[[     'Saldo'        ]].apply(pd.to_numeric).round(4)
            ZPP101_Fert_Ordem_Aberta.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\ZPP101_Fert_Ordem_Aberta.xlsx' , index = None)
        ZPP101_Fert_Ordem_Aberta()


        
        def ZSD138_carteira_fert():
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
        ZSD138_carteira_fert()

                
        def Me2m_Compoente_Consultar_Pedido_De_Compra():
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
            Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.iloc[:, [8,37]]
            estrutura_df = {
                'Componente': str,
                'Saldo': str
            }

            Me2m_Compoente_Consultar_Pedido_De_Compra = pd.DataFrame(Me2m_Compoente_Consultar_Pedido_De_Compra, dtype=str)
            Me2m_Compoente_Consultar_Pedido_De_Compra.columns = estrutura_df.keys()
            Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.astype(estrutura_df)
            Me2m_Compoente_Consultar_Pedido_De_Compra['Saldo'] = Me2m_Compoente_Consultar_Pedido_De_Compra['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
            Me2m_Compoente_Consultar_Pedido_De_Compra = Me2m_Compoente_Consultar_Pedido_De_Compra.groupby("Componente")["Saldo"].sum().round(3)
            Me2m_Compoente_Consultar_Pedido_De_Compra.fillna(value=False, inplace= True)
            Me2m_Compoente_Consultar_Pedido_De_Compra.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Me2m_Compoente_Consultar_Pedido_De_Compra.xlsx')
        Me2m_Compoente_Consultar_Pedido_De_Compra()
        


        def Me2m_Fert_Consultar_Pedido_De_Compra() :
            with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\Me2m_Fert_Consultar_Pedido_De_Compra.txt', encoding='latin1'  ) as f:
                linhas = f.readlines()
                f.close()   

            linhas = [x for x in linhas if 'Sem dados' not in x]
            linhas = [x for x in linhas if 'Material' not in x]
            linhas = [x for x in linhas if '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
            linhas = [x for x in linhas if '|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' not in x]
            linhas = [x for x in linhas if x[0][0]=='|']
            linhas = [x.split('|') for x in linhas]
            linhas = [x[1:-1] for x in linhas]

            Me2m_Fert_Consultar_Pedido_De_Compra = []
            for linha in linhas:     
                dados_temp = []
                for dado in linha:
                    dado = dado.rstrip().lstrip()
                    dados_temp.append(dado)
                Me2m_Fert_Consultar_Pedido_De_Compra.append(dados_temp)
            Me2m_Fert_Consultar_Pedido_De_Compra = pd.DataFrame(Me2m_Fert_Consultar_Pedido_De_Compra )
            Me2m_Fert_Consultar_Pedido_De_Compra = Me2m_Fert_Consultar_Pedido_De_Compra.iloc[:, [8,37]]
            estrutura_df = {
                'Material': str,
                'Saldo': str
            }

            Me2m_Fert_Consultar_Pedido_De_Compra = pd.DataFrame(Me2m_Fert_Consultar_Pedido_De_Compra, dtype=str)
            Me2m_Fert_Consultar_Pedido_De_Compra.columns = estrutura_df.keys()
            Me2m_Fert_Consultar_Pedido_De_Compra = Me2m_Fert_Consultar_Pedido_De_Compra.astype(estrutura_df)
            Me2m_Fert_Consultar_Pedido_De_Compra['Saldo'] = Me2m_Fert_Consultar_Pedido_De_Compra['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
            Me2m_Fert_Consultar_Pedido_De_Compra = Me2m_Fert_Consultar_Pedido_De_Compra.groupby("Material")["Saldo"].sum().round(3)
            Me2m_Fert_Consultar_Pedido_De_Compra.fillna(value=False, inplace= True)
            Me2m_Fert_Consultar_Pedido_De_Compra.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Me2m_Fert_Consultar_Pedido_De_Compra.xlsx')
        Me2m_Fert_Consultar_Pedido_De_Compra()


        def ultima_venda():
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
        ultima_venda()


        def Porto_Fert():
            with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\porto.txt', encoding='latin1'  ) as f:
                linhas = f.readlines()
                f.close()   


            linhas = [x for x in linhas if 'Campo de seleção    ' not in x]
            linhas = [x for x in linhas if 'Não Receb/Receb. com Saldo' not in x]
            linhas = [x for x in linhas if 'Layout' not in x]
            linhas = [x for x in linhas if 'Critérios' not in x]
            linhas = [x for x in linhas if 'Saldo do Pedido' not in x]
            linhas = [x for x in linhas if 'Estatíst.dados' not in x]
            linhas = [x for x in linhas if 'Registros transfs.' not in x]
            linhas = [x for x in linhas if 'Destes ocultados por filtro' not in x]
            linhas = [x for x in linhas if 'SISTEMA COMERCIO EXTERIOR' not in x]
            linhas = [x for x in linhas if 'Relatório de Follow Up (Modelo CISER)' not in x]
            linhas = [x for x in linhas if '-----------------------------------' not in x]
            linhas = [x for x in linhas if '---------------------------------' not in x]
            linhas = [x for x in linhas if '------------------------------' not in x]
            linhas = [x for x in linhas if '----------------------------' not in x]
            linhas = [x for x in linhas if 'Material' not in x]
            linhas = [x for x in linhas if x[0][0]=='|']
            linhas = [x.split('|') for x in linhas]
            linhas = [x[1:-1] for x in linhas]

            porto = []
            for linha in linhas:     
                dados_temp = []
                for dado in linha:
                    dado = dado.rstrip().lstrip()
                    dados_temp.append(dado)
                porto.append(dados_temp)

            porto = pd.DataFrame(porto )

            estrutura_df = {
                'Material': str,
                'Saldo': str
            }

            porto = pd.DataFrame(porto, dtype=str)
            porto.columns = estrutura_df.keys()
            porto = porto.astype(estrutura_df)
            porto['Saldo'] = porto['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
            porto = porto.groupby("Material")["Saldo"].sum().round(3)
            porto.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Porto_Fert.xlsx')
        Porto_Fert()



        def porto_componente():
            with open ('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Txt\\porto.txt', encoding='latin1'  ) as f:
                linhas = f.readlines()
                f.close()   


            linhas = [x for x in linhas if 'Campo de seleção    ' not in x]
            linhas = [x for x in linhas if 'Não Receb/Receb. com Saldo' not in x]
            linhas = [x for x in linhas if 'Layout' not in x]
            linhas = [x for x in linhas if 'Critérios' not in x]
            linhas = [x for x in linhas if 'Saldo do Pedido' not in x]
            linhas = [x for x in linhas if 'Estatíst.dados' not in x]
            linhas = [x for x in linhas if 'Registros transfs.' not in x]
            linhas = [x for x in linhas if 'Destes ocultados por filtro' not in x]
            linhas = [x for x in linhas if 'SISTEMA COMERCIO EXTERIOR' not in x]
            linhas = [x for x in linhas if 'Relatório de Follow Up (Modelo CISER)' not in x]
            linhas = [x for x in linhas if '-----------------------------------' not in x]
            linhas = [x for x in linhas if '---------------------------------' not in x]
            linhas = [x for x in linhas if '------------------------------' not in x]
            linhas = [x for x in linhas if '----------------------------' not in x]
            linhas = [x for x in linhas if 'Material' not in x]

            linhas = [x for x in linhas if x[0][0]=='|']
            linhas = [x.split('|') for x in linhas]
            linhas = [x[1:-1] for x in linhas]

            Porto_Comp = []
            for linha in linhas:     
                dados_temp = []
                for dado in linha:
                    dado = dado.rstrip().lstrip()
                    dados_temp.append(dado)
                Porto_Comp.append(dados_temp)


            estrutura_df = {
                'Componente': str,
                'Saldo': str
            }

            Porto_Comp = pd.DataFrame(Porto_Comp, dtype=str)
            Porto_Comp.columns = estrutura_df.keys()
            Porto_Comp = Porto_Comp.astype(estrutura_df)

            Porto_Comp['Saldo'] = Porto_Comp['Saldo'].apply(lambda x: float(x.replace(".","").replace(",",".")))
            Porto_Comp = Porto_Comp.groupby("Componente")["Saldo"].sum().round(3)

            Porto_Comp.to_excel('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Saldos\\Porto_Comp.xlsx')
        porto_componente()
    Read_Txt_Txt()


    def Read_xlsx():
        
        dados = pd.read_csv('U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Dados_Cs11\\dados_transacao.csv', sep=',', dtype = str, header=None)

        dados = pd.DataFrame(dados, dtype = str)
        estrutura_df = {
            'Fert': str,
            'Componente': str,
            'Tipo': str,
            'Status': str,
        }
        dados.columns = estrutura_df.keys()
        dados.drop(['Status'], axis=1, inplace=True)

        dados = dados.rename(columns={'Fert': 'Material'})
        dados['Componente'] = dados['Componente'].apply(lambda x: str(x.replace("nan","")))
        dados['Tipo'] = dados['Tipo'].apply(lambda x: str(x.replace("nan","")))


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


        dados = dados.fillna({ 'Ult.Venda': '01.01.1999'})

        data_day =  datetime.date.today()

        dados["Criado"] = pd.to_datetime(dados["Criado"], dayfirst=True)
        dados["Ult.Venda"] = pd.to_datetime(dados["Ult.Venda"], dayfirst=True)

        dados['DiasCadastrado'] = (pd.to_datetime(data_day) - dados['Criado'])
        dados['DiasVenda'] = (pd.to_datetime(data_day) - dados['Ult.Venda'])

        dados['Criado'] = dados['Criado'].dt.strftime('%d/%m/%Y')
        dados['Ult.Venda'] = dados['Ult.Venda'].dt.strftime('%d/%m/%Y')

        dados['DiasCadastrado'] = (dados['DiasCadastrado'] / np.timedelta64(1, 'D')).astype(int)
        dados['DiasVenda'] = (dados['DiasVenda'] / np.timedelta64(1, 'D')).astype(int)

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
        'Total'
        ]] = dados[[
            'Estoque_Fert',
            'Compra_Fert',
            'Orderm_Aberta',
            'Cart_Aberta',
            'Porto_Fert',
            'Compra_Comp',
            'Porto_Comp',
            'Estoque_Comp',
            'Total'
            ]].apply(pd.to_numeric).round(4)

        day = data_day.day
        month = data_day.month
        year = data_day.year

        dados['Estoque_Comp_Soma'] = dados.groupby('Material')["Estoque_Comp"].transform(np.sum)
        dados['Porto_Comp_Soma'] = dados.groupby('Material')["Porto_Comp"].transform(np.sum)
        dados['Compra_Comp_Soma'] = dados.groupby('Material')["Compra_Comp"].transform(np.sum)
        
        #REGRAS PARA ELIMINAR MATERIAL
        #FERT
        dados = dados.loc[(dados['GE'] != 'E')]
        dados = dados.loc[(dados['DiasVenda'] >= 365)]

        dados = dados.loc[(dados['DiasCadastrado'] >= 365)]
        dados = dados.loc[(dados['Estoque_Fert'] == 0)]
        dados = dados.loc[(dados['Porto_Fert'] == 0)]
        dados = dados.loc[(dados['Compra_Fert'] == 0)]
        dados = dados.loc[(dados['Ordem'] == '')]
        dados = dados.loc[(dados['Cart_Aberta'] == 0)]

        #Total Venda
        dados = dados.loc[(dados['Total'] == 0)]

        #Componente
        dados = dados.loc[(dados['Estoque_Comp_Soma'] == 0)]
        dados = dados.loc[(dados['Porto_Comp_Soma'] == 0)]
        dados = dados.loc[(dados['Compra_Comp_Soma'] == 0)]

        dados.to_excel(f'U:\\Controladoria\\Cadastro\\14 - AUTOMATIZAÇÃO PYTHON\\Projetos\\Rafa pilot\\Marcar item para eliminar\\Pro\\Jupyter\\{year}-{month}-{day} - Analise Final DD.xlsx', index = None)


    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders


    def Enviar_Email():       
        remetente = 'cadastro@ciser.com.br'
        senha_rede = 'ciscad01**' # Colocar aqui a senha do e-mail do cadastro@ciser.com.br
        destinatario = 'rafael.refundini@ciser.com.br'
        destinatario2 = 'cadastro@ciser.com.br'
        destinatario3 = 'rafael.ramalho@ciser.com.br'
        assunto = 'Relatório Marcar itens para eliminar'
        # Preenche abaixo o corpo da mensagem.
        texto = f"""
            Relatório Marcar itens para eliminar


            OBS: MENSAGEM AUTOMÁTICA.
        """
        email_sender = remetente
        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatario
        msg['To'] = destinatario2

        msg['Subject'] = assunto

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open("2023-1-24 - Analise Final DD.xlsx", "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="2023-1-24 - Analise Final DD.xlsx"')
        msg.attach(part)
        
        msg.attach(MIMEText(_text=texto.encode('utf-8'), _charset='utf-8'))
        port = 587 if '@ciser' in destinatario else 25
        server = smtplib.SMTP(host='smtp.office365.com', port=port)
        server.ehlo()
        server.starttls()
        server.login(remetente, senha_rede)
        text = msg.as_string()
        server.sendmail(email_sender, destinatario, text)
        server.sendmail(email_sender, destinatario2, text)
        server.sendmail(email_sender, destinatario3, text)
        print('Email enviado')
        server.quit()

    Enviar_Email()
 
    Read_xlsx()
if(__name__ == "__main__"):
    play()
