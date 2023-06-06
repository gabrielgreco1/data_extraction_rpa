import oracledb
import pandas as pd
import openpyxl
import datetime
from tkinter import *
from tkinter import messagebox

janela = Tk()
janela.title("Extração de relatórios")
janela.geometry('500x500')

var = StringVar();
var.set("Selecione um relatório")

label_data = Label(janela, text="Digite a data de vencimento (dd/mm/aaaa):")
label_data.pack(padx=10, pady=10)

data_digitada = Entry(janela)
data_digitada.pack(padx=15, pady=15)

def selecionar_opcao():
    global data_formatada
    data_formatada = data_digitada.get()
    data_formatada = datetime.datetime.strptime(data_formatada, "%d/%m/%Y").strftime("%Y%m%d")
    selected_option = var.get()
    if selected_option == 'Contas a Pagar':
        contaspag()
    elif selected_option == 'Contas a Receber':
        contasrec()


def contaspag():
    # Define o dia de hoje
    dia_atual = datetime.datetime.now().day
    print("-------------------------------------------------------------------------")   

    # Função para conectar no banco de dados e extrair os dados
    def extraction():
            # Conexão no banco
            conection = oracledb.connect('')
            print("Conexão bem-sucedida com o banco de dados Oracle")
            cursor = conection.cursor()

            # Monta o parâmetro de data com a data atual
            # data_consulta = datetime.datetime.now().strftime('%Y%m') + str(dia_atual).zfill(2)

            #Executa a query e armazena na variável 
            cursor.execute("SELECT * FROM SE2010 WHERE E2_VENCTO = :data", data=data_formatada)
            SELECT = cursor.fetchall()
            
            # Recebe todos os dados, os transforma em string e armazena em uma lista
            resultado_f = []

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


    # Condição para caso seja dia 1 (planilha terá os dados excluídos)
    if dia_atual == 1:
        # Verifica a conexão ao banco
        try:
            resultado = extraction()
        except oracledb.DatabaseError as e:
            # Captura e trata a exceção de erro de conexão
            error, = e.args
            print("Erro ao conectar ao banco de dados Oracle:", error.message)

        # Adiciona os dados da lista a um Dataframe
        resultado_df = pd.DataFrame(resultado)

        # Limpar a formatação dos dados
        for column in resultado_df:
            resultado_df[column] = resultado_df[column].str.strip("[]'")

        # Cola no excel limpando os dados antigos
        resultado_df.to_excel('excel.xlsx', index=False, engine='openpyxl')
        print("Dados adicionados na planilha com sucesso!")

    # Condição para caso não seja dia 1 (manterá os dados da planilha e adicionará os novos)
    else:
        # Verifica a conexão ao banco
        try:
            resultado = extraction()
        except oracledb.DatabaseError as e:
            # Captura e trata a exceção de erro de conexão
            error, = e.args
            print("Erro ao conectar ao banco de dados Oracle:", error.message)


        # Adiciona os dados da lista a um Dataframe
        resultado_df = pd.DataFrame(resultado)

        # Limpar a formatação dos dados
        for column in resultado_df:
            resultado_df[column] = resultado_df[column].str.strip("[]'")
        
        # Verifica se já há dados no excel, e os armazena na variável 
        try:
            verificador = pd.read_excel('excel.xlsx')
        except:
            verificador = pd.DataFrame()

        # Caso não haja nada no excel, adiciona apenas os dados novos
        if verificador.empty:
            resultado_df.to_excel('excel.xlsx', index=False, engine='openpyxl')
            print("Dados adicionados na planilha com sucesso!")

        # Caso haja dados no excel, armazena em uma variável e transforma em DataFrame
        else:
            try:
                existing_data = pd.read_excel('excel.xlsx')
            except FileNotFoundError:
                existing_data = pd.DataFrame()
                
            # Concatenar os dados existentes com os novos dados
            merged_data = pd.concat([existing_data, resultado_df], ignore_index=True)
            # Salvar os dados de volta na planilha do Excel adicionando os dados novos
            merged_data.to_excel('excel.xlsx', index=False)
            print("Dados adicionados na planilha com sucesso!")
            
    print("-------------------------------------------------------------------------")        


def contasrec():
    # Define o dia de hoje
    dia_atual = datetime.datetime.now().day
    print("-------------------------------------------------------------------------")   

    # Função para conectar no banco de dados e extrair os dados
    def extraction():
            # Conexão no banco
            conection = oracledb.connect('')
            print("Conexão bem-sucedida com o banco de dados Oracle")
            cursor = conection.cursor()

            # Monta o parâmetro de data com a data atual
            # data_consulta = '20230501'

            #Executa a query e armazena na variável 
            cursor.execute("SELECT * FROM SE1010 WHERE E1_VENCTO = :data", data=data_formatada)
            SELECT = cursor.fetchall()
            
            # Recebe todos os dados, os transforma em string e armazena em uma lista
            resultado_f = []

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


    # Condição para caso seja dia 1 (planilha terá os dados excluídos)
    if dia_atual == 1:
        # Verifica a conexão ao banco
        try:
            resultado = extraction()
        except oracledb.DatabaseError as e:
            # Captura e trata a exceção de erro de conexão
            error, = e.args
            print("Erro ao conectar ao banco de dados Oracle:", error.message)

        # Adiciona os dados da lista a um Dataframe
        resultado_df = pd.DataFrame(resultado)

        # Limpar a formatação dos dados
        for column in resultado_df:
            resultado_df[column] = resultado_df[column].str.strip("[]'")

        # Cola no excel limpando os dados antigos
        resultado_df.to_excel('excel.xlsx', index=False, engine='openpyxl')
        print("Dados adicionados na planilha com sucesso!")

    # Condição para caso não seja dia 1 (manterá os dados da planilha e adicionará os novos)
    else:
        # Verifica a conexão ao banco
        try:
            resultado = extraction()
        except oracledb.DatabaseError as e:
            # Captura e trata a exceção de erro de conexão
            error, = e.args
            print("Erro ao conectar ao banco de dados Oracle:", error.message)


        # Adiciona os dados da lista a um Dataframe
        resultado_df = pd.DataFrame(resultado)

        # Limpar a formatação dos dados
        for column in resultado_df:
            resultado_df[column] = resultado_df[column].str.strip("[]'")
        
        # Verifica se já há dados no excel, e os armazena na variável 
        try:
            verificador = pd.read_excel('excel.xlsx')
        except:
            verificador = pd.DataFrame()

        # Caso não haja nada no excel, adiciona apenas os dados novos
        if verificador.empty:
            resultado_df.to_excel('excel.xlsx', index=False, engine='openpyxl')
            print("Dados adicionados na planilha com sucesso!")

        # Caso haja dados no excel, armazena em uma variável e transforma em DataFrame
        else:
            try:
                existing_data = pd.read_excel('excel.xlsx')
            except FileNotFoundError:
                existing_data = pd.DataFrame()
                
            # Concatenar os dados existentes com os novos dados
            merged_data = pd.concat([existing_data, resultado_df], ignore_index=True)
            # Salvar os dados de volta na planilha do Excel adicionando os dados novos
            merged_data.to_excel('excel.xlsx', index=False)
            print("Dados adicionados na planilha com sucesso!")
            
    print("-------------------------------------------------------------------------")        




texto1 = Label(janela, text="Selecione a rotina para realizar a extração de relatório")
texto1.pack(padx=20, pady=20)

opcoes_menu = OptionMenu(janela, var, "Contas a Pagar" , "Contas a Receber")
opcoes_menu.pack(padx=21, pady=21)

button = Button(janela, text="Extrair", command = selecionar_opcao)
button.pack(padx=30, pady=30)

janela.mainloop()
