import sys, csv, xlsxwriter,os,openpyxl, win32com.client, pandas as pd
from xlsxwriter import workbook
from pandas import ExcelWriter
from openpyxl import load_workbook, Workbook, styles
from openpyxl.styles import Protection, NamedStyle


class csv_para_excel: #Classe para transformar o CSV em XLSX.

    def __init__ (self):#Funcao inicial da classe.
        return #Continuar sem necessidade de chamada.

    def tratar_csv(csv_sem_tratar): #Funcao que trata o CSV com ";" e troca para ",".
        with open(csv_sem_tratar, "rt") as fin: #Abre o arquivo CSV sem tratamento.
            with open(r"C:\Users\Administrador\Desktop\Projeto_Requisitos\CSV\CSV_Out\CSVTratado.csv", "wt") as fout: #Cria o arquivo CSV tratado.
                for line in fin: #Repetição para percorrer o texto.
                    fout.write(line.replace(";", ",")) #Altera o ";" por ",".
        global csv_file #Cria a variavel global csv_file.
        csv_file = (r"C:\Users\Administrador\Desktop\Projeto_Requisitos\CSV\CSV_Out\CSVTratado.csv") #Armazena na variavel global o PATH do CSV tratado.
        return csv_file #Retorna a variavel para o controlador.

    def getdf(csv_file):#Extrai os dados do CSV tratado e armazena em um objeto Pandas.
        global data_frame #Cria a variavel global data_frame
        data_frame = pd.read_csv(csv_file, encoding="ISO-8859-1", header=0, index_col=0 ) # Com a funcao read_csv do Pandas abre o arquivo , que ja extrai os dados CSV e transforma em um Objeto.
        return data_frame #Retorna o objeto para o controlador.

    def postxlsx(xlsx_file, xlsm_file,data_frame): #Função que pega o objeto (data_frame) com o Data Frame (DF) do Pandas.
        print(data_frame) #Printa o DF para checagem visual.
        with pd.ExcelWriter(xlsx_file, engine="openpyxl", read_only=False , mode="a") as writer: # Com a metodo ExcelWriter abre o arquivo xlsx aonde se deseja inserir os dados, com a biblioteca openpyxl, no modo (a) de "append", como writer. Notas:Provavelmente o "a" é disso, pesquisar.
            book  = load_workbook(xlsx_file) #Abre o xlsx com Workbook do openpyxl
            writer.book = book #Seleciona esse Workbook para ser a base do Writer.
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)#Copia as planilhas.
            data_frame.to_excel(writer, sheet_name='Pasta1', startrow=0, startcol=0) #Com o metodo to_excel ele passa o objeto (DF) do pandas, para o local especificado e com os parametros da engine e modo.
            book.close()
            print(data_frame.applymap(type))#Print dos tipos de dados
            return data_frame#Retorna a varaivel Data_Frame para o controlador

    def formatartabela(xlsx_file, xlsm_file, data_frame):#Usa o vbaProject.bin para jogar rodar a macro na planilha Excel xlsm. Com o XLSXWRITER.
            df = data_frame # Define o Data Frame
            filenamexlsx = xlsx_file # Passa a variavel do Path do XLSX
            writer = pd.ExcelWriter(filenamexlsx, engine='xlsxwriter') #Aciona o Writer, com os dados do XLSX
            df.to_excel(writer, sheet_name='Pasta1', index=False, startrow=0, startcol=0)#Com o metodo to_excel ele passa o objeto (DF) do pandas, para o local especificado e com os parametros da engine e modo.
            file_name_macro = ('reqxlsm')#Nome que o XLSM irá ter.
            workbook = writer.book #Definindo o Workbook do XLSXWRITER.
            workbook.filename = (file_name_macro+".xlsm")#Definindop o tipo do arquivo gerado.
            workbook.add_vba_project(r'C:\Users\Administrador\Desktop\Projeto_Requisitos\Back-End\Regras\MacroVBA\vbaProject.bin')#Le a Macro do arquivo vbaProject.bin
            writer.save()#Salva o Arquivo
            #Rodar a Macro com o win32com
            if os.path.exists(r'C:\Users\Administrador\Desktop\Projeto_Requisitos\reqxlsm.xlsm'):
                print("aqui")
                xl = win32com.client.Dispatch('Excel.Application')
                xl.Workbooks.Open(Filename =r'C:\Users\Administrador\Desktop\Projeto_Requisitos\reqxlsm.xlsm', ReadOnly = 0)
                xl.Application.Run("'reqxlsm.xlsm'!Padroniza")
                xl.Application.Quit()
                del xl
                print("Macro refresh completed!")
            writer.close()




    if __name__ == '__main__': #Controlador das funões da classe.
        tratar_csv(r"C:\Users\Administrador\Desktop\Projeto_Requisitos\CSV\CSV_In\CSVTesteSemTratamento.csv")#Chama a funcao que trata o CSV e como parametro a PAth do CSV sem tratamento.
        getdf(csv_file) #Chama a funcap que gera o data_frame, objeto com os dados CSV do pandas. Passando como parametro a PATH do CSV tratado.
        postxlsx(r"C:\Users\Administrador\Desktop\Projeto_Requisitos\XLSX\xlsxparacarga.xlsx",r"C:\Users\Administrador\Desktop\Projeto_Requisitos\XLSM\xlsmparacarga.xlsm", data_frame)#Chama a funcao
        formatartabela(r"C:\Users\Administrador\Desktop\Projeto_Requisitos\XLSX\xlsxparacarga.xlsx",r"C:\Users\Administrador\Desktop\Projeto_Requisitos\XLSM\xlsmparacarga.xlsm", data_frame)
        #formatartabela(r"C:\Users\Administrador\Desktop\Projeto_Requisitos\XLSX\Base\xlsxparacarga.xlsx")

csv_para_excel() #Chama a classe que da carga dos dados CSV no EXCEL
