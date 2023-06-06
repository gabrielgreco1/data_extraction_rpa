import oracledb
import pandas as pd
import openpyxl
import datetime

# Define o dia de hoje
dia_inicial = '01'
dia_atual = '10'
print("-------------------------------------------------------------------------")   

# Função para conectar no banco de dados e extrair os dados
def extraction():
        # Conexão no banco
        conection = oracledb.connect('')
        print("Conexão bem-sucedida com o banco de dados Oracle")
        cursor = conection.cursor()
        
        # Monta o parâmetro de data com a data atual
        data_inicial = datetime.datetime.now().strftime('%Y%m') + str(dia_inicial)
        data_final = datetime.datetime.now().strftime('%Y%m') + str(dia_atual).zfill(2)
        #Executa a query e armazena na variável 
        cursor.execute("SELECT * FROM SE2010 WHERE E2_VENCTO BETWEEN :data_inicial AND :data_final", data_inicial=data_inicial, data_final=data_final)
        SELECT = cursor.fetchall()
        
        # cursor.execute("SELECT * FROM SE2130 WHERE E2_VENCTO = BETWEEN :data_inicial AND :data_final", data_inicial=data_inicial, data_final=data_final)
        # result2 = cursor.fetchall()

        # SELECT = result1 + result2

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
        print("Dados extraídos com sucesso! Período: ", data_inicial, " até ", data_final)
        return resultado_f
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
    # Concatenar os dados existentes com os novos dados
    merged_data = resultado_df.drop_duplicates(keep='first')

    # Salvar os dados de volta na planilha do Excel adicionando os dados novos
    merged_data.to_excel('excel.xlsx', index=False)
    print("Processo finalizado!")

print("-------------------------------------------------------------------------")  

