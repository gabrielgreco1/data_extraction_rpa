import oracledb
import pandas as pd
import openpyxl
import datetime
from Infos.dados import Conexao

conexao = Conexao()
conection_string = f"{conexao.user}/{conexao.password}@{conexao.host}:{conexao.port}/{conexao.service_name}"
query = f"{conexao.query}"

def extraction():
    # Conexão no banco
    conection = oracledb.connect(conection_string)
    print("Conexão bem-sucedida com o banco de dados Oracle")
    cursor = conection.cursor()
    cursor.execute(query)
    SELECT = cursor.fetchall()

    resultado_f = []

    # Obter os nomes das colunas
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
    print("Dados extraídos com sucesso! Período: ", conexao.data_inicial, " até ", conexao.data_final)
    return resultado_f


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

# Verifica se já há dados no excel, e os armazena na variável 
try:
    verificador = pd.read_excel('C:\\Base.xlsx')
except:
    verificador = pd.DataFrame()

# Caso não haja nada no excel, adiciona apenas os dados novos
if verificador.empty:
    resultado_df.to_excel('C:\\Base.xlsx', index=False, engine='openpyxl', header = True)
    print("Dados adicionados na planilha com sucesso!")

# Caso haja dados no excel, armazena em uma variável e transforma em DataFrame
else:    
    # Concatenar os dados existentes com os novos dados
    merged_data = resultado_df.drop_duplicates(keep='first')

    # Salvar os dados de volta na planilha do Excel adicionando os dados novos
    merged_data.to_excel('C:\\Base.xlsx', index=False)
    print("Processo finalizado!")