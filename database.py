import pandas as pd
import mysql.connector
from tkinter import messagebox, filedialog
from utils import formatar_data_db, formatar_valor_protheus

def conectar_db():
    try:
        conn = mysql.connector.connect(
            host='10.1.1.24',
            user='root',
            password='!POTapczuk123kj',
            database='db_sagg_consult'
        )
        return conn
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao conectar ao banco de dados: {e}")
        return None

def verificar_existencia_bb(cursor, data, historico, valor, agencia, conta):
    sql = """
    SELECT COUNT(*) FROM extratobb 
    WHERE data = %s AND historico = %s AND valor = %s AND agencia = %s AND conta = %s
    """
    cursor.execute(sql, (data, historico, valor, agencia, conta))
    return cursor.fetchone()[0] > 0

def inserir_dados_banco_brasil(df, agencia, conta):
    conn = conectar_db()
    if conn is None:
        return

    cursor = conn.cursor()
    total_inseridos = 0

    for _, row in df.iterrows():
        data = formatar_data_db(row['Data'])
        historico = row['Historico']
        valor = row['Valor R$']

        if not verificar_existencia_bb(cursor, data, historico, valor, agencia, conta):
            sql = "INSERT INTO extratobb (data, historico, valor, agencia, conta) VALUES (%s, %s, %s, %s, %s)"
            cursor.execute(sql, (data, historico, valor, agencia, conta))
            total_inseridos += 1

    conn.commit()
    cursor.close()
    conn.close()

    return total_inseridos

def verificar_existencia_itau(cursor, data, razao_social, valor, agencia, conta):
    sql = """
    SELECT COUNT(*) FROM extratoitau 
    WHERE data = %s AND razao = %s AND valor = %s AND agencia = %s AND conta = %s
    """
    cursor.execute(sql, (data, razao_social, valor, agencia, conta))
    return cursor.fetchone()[0] > 0

def inserir_dados_itau(df, agencia, conta):
    conn = conectar_db()
    if conn is None:
        return 0

    cursor = conn.cursor()
    total_inseridos = 0

    for _, row in df.iterrows():
        data = formatar_data_db(row['Data'])
        razao_social = row['Razão Social'] if pd.notna(row['Razão Social']) else ''  # Se NaN, usa string vazia
        valor = row['Valor (R$)']

        if not verificar_existencia_itau(cursor, data, razao_social, valor, agencia, conta):
            if pd.notna(data) and pd.notna(valor):
                sql = "INSERT INTO extratoitau (data, razao, valor, agencia, conta) VALUES (%s, %s, %s, %s, %s)"
                cursor.execute(sql, (data, razao_social, valor, agencia, conta))
                total_inseridos += 1

    conn.commit()
    cursor.close()
    conn.close()

    return total_inseridos

def verificar_existencia_protheus(cursor, data, titulo, saidas):
    sql = """
    SELECT COUNT(*) FROM extratoprotheus 
    WHERE data = %s AND titulo = %s AND saidas = %s
    """
    cursor.execute(sql, (data, titulo, saidas))
    return cursor.fetchone()[0] > 0

def inserir_dados_protheus(df):
    conn = conectar_db()
    if conn is None:
        return

    cursor = conn.cursor()
    total_inseridos = 0

    # Filtrar valores não válidos na coluna SAIDAS
    df = df.dropna(subset=['DATA', 'SAIDAS'])
    df['SAIDAS'] = df['SAIDAS'].apply(formatar_valor_protheus)  # Formata corretamente

    # Remover linhas onde SAIDAS é zero ou menor
    df = df[df['SAIDAS'] > 0]

    for _, row in df.iterrows():
        data = formatar_data_db(row['DATA'])
        titulo = row['PREFIXO/TITULO'] if 'PREFIXO/TITULO' in row and pd.notna(row['PREFIXO/TITULO']) else None
        saidas = row['SAIDAS']

        try:
            sql = "INSERT INTO extratoprotheus (data, titulo, saidas, natureza) VALUES (%s, %s, %s, %s)"
            cursor.execute(sql, (data, titulo, saidas, None))
            total_inseridos += 1
        except Exception as e:
            print(f'Erro ao inserir: {e}')

    conn.commit()
    cursor.close()
    conn.close()

    return total_inseridos

def carregar_dataframe(cursor, query):
    cursor.execute(query)
    columns = [col[0] for col in cursor.description]
    data = cursor.fetchall()
    return pd.DataFrame(data, columns=columns)

def processar_dados(df_filtrado, extratobb_df, extratoitau_df, analise_data):
    if df_filtrado.empty:
        return  # Evita processamento se o DataFrame filtrado estiver vazio

    for _, row in df_filtrado.iterrows():
        valor = row['saidas']
        data = row['data']

        valor_bb = extratobb_df.loc[extratobb_df['valor'] == valor]
        valor_itau = extratoitau_df.loc[extratoitau_df['valor'] == valor]

        if not valor_bb.empty:
            analise_data.append({'DATA': data, 'VALOR': valor, 'BANCO': 'BANCO DO BRASIL'})
        elif not valor_itau.empty:
            analise_data.append({'DATA': data, 'VALOR': valor, 'BANCO': 'ITAÚ'})
        else:
            analise_data.append({'DATA': data, 'VALOR': valor, 'BANCO': 'NÃO LOCALIZADO'})

def analise_financeira(mes_selecionado, ano_selecionado, meses):
    try:
        conn = conectar_db()
        if conn is None:
            return
        
        cursor = conn.cursor()
        
        # Carregar todos os DataFrames necessários
        extratoprotheus_df = carregar_dataframe(cursor, "SELECT * FROM extratoprotheus")
        extratoprotheus_df['data'] = pd.to_datetime(extratoprotheus_df['data'], errors='coerce')

        extratobb_df = carregar_dataframe(cursor, "SELECT * FROM extratobb")
        extratobb_df['data'] = pd.to_datetime(extratobb_df['data'], errors='coerce')

        extratoitau_df = carregar_dataframe(cursor, "SELECT * FROM extratoitau")
        extratoitau_df['data'] = pd.to_datetime(extratoitau_df['data'], errors='coerce')

        analise_data = []
        nao_localizados_data = []  # Lista para armazenar não localizados
        extratoprotheus_filtrado = pd.DataFrame()

        # Filtragem de dados baseada no mês e ano selecionados
        if mes_selecionado != "Geral":
            mes_numero = meses.index(mes_selecionado) + 1  # Ajuste para o mês (1 a 12)
            extratoprotheus_filtrado = extratoprotheus_df[
                (extratoprotheus_df['data'].dt.month == mes_numero) & 
                (extratoprotheus_df['data'].dt.year == int(ano_selecionado))
            ]
        else:
            extratoprotheus_filtrado = extratoprotheus_df

        # Processar os dados filtrados
        processar_dados(extratoprotheus_filtrado, extratobb_df, extratoitau_df, analise_data)

        # Adicionar dados não localizados
        for _, row in extratobb_df.iterrows():
            valor = row['valor']
            data = row['data']
            if extratoprotheus_df[extratoprotheus_df['saidas'] == valor].empty:
                nao_localizados_data.append({'DATA': data, 'VALOR': valor, 'BANCO': 'BANCO DO BRASIL', 'INFO': 'NÃO LOCALIZADO'})

        for _, row in extratoitau_df.iterrows():
            valor = row['valor']
            data = row['data']
            if extratoprotheus_df[extratoprotheus_df['saidas'] == valor].empty:
                nao_localizados_data.append({'DATA': data, 'VALOR': valor, 'BANCO': 'ITAÚ', 'INFO': 'NÃO LOCALIZADO'})

        # Cria o DataFrame da análise e o DataFrame dos não localizados
        analise_df = pd.DataFrame(analise_data)
        nao_localizados_df = pd.DataFrame(nao_localizados_data)

        # Exibir mensagens de aviso se não houver dados
        if analise_df.empty:
            messagebox.showwarning("Aviso", "Nenhum dado encontrado para a seleção.")
            return
        
        # Formatar as datas
        analise_df['DATA'] = pd.to_datetime(analise_df['DATA'], errors='coerce').dt.strftime('%d/%m/%Y')
        nao_localizados_df['DATA'] = pd.to_datetime(nao_localizados_df['DATA'], errors='coerce').dt.strftime('%d/%m/%Y')

        # Nome padrão do arquivo
        nome_arquivo = f"analise_financeira_sagg_{mes_selecionado}{ano_selecionado}.xlsx" if mes_selecionado != "Geral" else "analise_financeira_sagg_diversos.xlsx"

        # Solicitar ao usuário para salvar o arquivo
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   initialfile=nome_arquivo,
                                                   title="Salvar Análise Financeira",
                                                   filetypes=[("Excel files", "*.xlsx;*.xls")])
        if save_path:
            # Salvar o DataFrame em múltiplas folhas
            with pd.ExcelWriter(save_path) as writer:
                analise_df.to_excel(writer, sheet_name='Protheus', index=False)
                nao_localizados_df.to_excel(writer, sheet_name='Não Localizados', index=False)

            messagebox.showinfo("Sucesso", "Análise financeira gerada com sucesso!")
        
        cursor.close()
        conn.close()

    except Exception as e:
        messagebox.showerror("Erro", str(e))