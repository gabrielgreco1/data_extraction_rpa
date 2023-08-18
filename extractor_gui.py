import oracledb
import pandas as pd
import openpyxl
import datetime
from tkinter import *
from tkinter import messagebox
from infos.dados import Conexao

conexao = Conexao()
conection_string = f"{conexao.user}/{conexao.password}@{conexao.host}:{conexao.port}/{conexao.service_name}"
query_SRD = f"{conexao.query_srd}"
query_beneficios = f"{conexao.query_beneficios}"

janela = Tk()
janela.title("Extração de relatórios")
janela.geometry('500x500')

var = StringVar()
var.set("Selecione um relatório")

label_data = Label(janela, text="Digite a data do período (mm/aaaa):")
label_data.pack(padx=10, pady=10)

data_digitada = Entry(janela)
data_digitada.pack(padx=15, pady=15)

def selecionar_opcao():
    global data_formatada
    data_formatada = data_digitada.get()
    data_formatada = datetime.datetime.strptime(data_formatada, "%m/%Y").strftime("%Y%m")
    selected_option = var.get()
    if selected_option == 'Benefícios':
        benefícios()
    elif selected_option == 'Folha - SRD':
        folha()


def benefícios():

    # Função para conectar no banco de dados e extrair os dados
    def extraction():
            # Conexão no banco
            conection = oracledb.connect(conection_string)
            print("Conexão bem-sucedida com o banco de dados Oracle")
            cursor = conection.cursor()

            # Monta o parâmetro de data com a data atual
            # data_consulta = datetime.datetime.now().strftime('%Y%m') + str(dia_atual).zfill(2)

            #Executa a query e armazena na variável 
            cursor.execute(query_beneficios)
            
            SELECT = cursor.fetchall()
            
            # Recebe todos os dados, os transforma em string e armazena em uma lista
            resultado_f = []

            global column_names
            column_names = [desc[0] for desc in cursor.description] 

            for row in SELECT:
                campos = []

                for campo in row:
                    if isinstance(campo, bytes):
                        # Tratar campos do tipo bytes
                        campo_str = campo.decode('utf-8')
                    elif isinstance(campo, oracledb.LOB):
                        # Tratar campos do tipo LOB
                        campo_str = campo.read().decode('utf-8')
                    else:
                        # Converter campo para string
                        campo_str = str(campo)

                    campos.append(campo_str)

                resultado_f.append(campos)

            # Fecha a conexão com o banco e confirma a extração
            cursor.close()
            conection.close()
            print("Dados extraídos com sucesso! Dia: ", data_formatada)
            return resultado_f
    
            # Verifica a conexão ao banco
    try:
        resultado = extraction()
    except oracledb.DatabaseError as e:
        # Captura e trata a exceção de erro de conexão
        error, = e.args
        print("Erro ao conectar ao banco de dados Oracle:", error.message)

    # Adiciona os dados da lista a um Dataframe
    resultado_df = pd.DataFrame(resultado, columns=column_names)

    # Limpar a formatação dos dados
    for column in resultado_df:
        resultado_df[column] = resultado_df[column].str.strip("[]'")

    # Cola no excel limpando os dados antigos
    resultado_df.to_excel(f'beneficios{data_formatada}.xlsx', index=False, engine='openpyxl')
    print("Dados adicionados na planilha com sucesso!")
   
    print("-------------------------------------------------------------------------")  
    
def folha():

    # Função para conectar no banco de dados e extrair os dados
    def extraction():
            # Conexão no banco
            conection = oracledb.connect(conection_string)
            print("Conexão bem-sucedida com o banco de dados Oracle")
            cursor = conection.cursor()

            #Executa a query e armazena na variável 
            cursor.execute(query_SRD)
            
            SELECT = cursor.fetchall()
            
            # Recebe todos os dados, os transforma em string e armazena em uma lista
            resultado_f = []

            global column_names
            column_names = [desc[0] for desc in cursor.description] 

            for row in SELECT:
                campos = []

                for campo in row:
                    if isinstance(campo, bytes):
                        # Tratar campos do tipo bytes
                        campo_str = campo.decode('utf-8')
                    elif isinstance(campo, oracledb.LOB):
                        # Tratar campos do tipo LOB
                        campo_str = campo.read().decode('utf-8')
                    else:
                        # Converter campo para string
                        campo_str = str(campo)

                    campos.append(campo_str)

                resultado_f.append(campos)

            # Fecha a conexão com o banco e confirma a extração
            cursor.close()
            conection.close()
            print("Dados extraídos com sucesso! Dia: ", data_formatada)
            return resultado_f


    
        # Verifica a conexão ao banco
    try:
        resultado = extraction()
    except oracledb.DatabaseError as e:
        # Captura e trata a exceção de erro de conexão
        error, = e.args
        print("Erro ao conectar ao banco de dados Oracle:", error.message)

    # Adiciona os dados da lista a um Dataframe
    resultado_df = pd.DataFrame(resultado, columns=column_names)

    # Limpar a formatação dos dados
    for column in resultado_df:
        resultado_df[column] = resultado_df[column].str.strip("[]'")

    # Cola no excel limpando os dados antigos
    resultado_df.to_excel(f'folha{data_formatada}.xlsx', index=False, engine='openpyxl')
    print("Dados adicionados na planilha com sucesso!")
   
    print("-------------------------------------------------------------------------")        


     




texto1 = Label(janela, text="Selecione a rotina para realizar a extração de relatório")
texto1.pack(padx=20, pady=20)

opcoes_menu = OptionMenu(janela, var, "Benefícios" , "Folha")
opcoes_menu.pack(padx=21, pady=21)

button = Button(janela, text="Extrair", command = selecionar_opcao)
button.pack(padx=30, pady=30)

janela.mainloop()