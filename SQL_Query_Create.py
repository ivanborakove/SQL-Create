import pandas as pd
import easygui as gui
import sys

def update_table(excel_file):
    table_name = gui.enterbox("Digite o nome da tabela:")
    column_names = gui.enterbox("Digite os nomes das colunas (separados por vírgula):")
    if not table_name or not column_names:
        gui.msgbox("Tabela ou colunas não informadas!", "Aviso")
        return
    column_names = [col.strip() for col in column_names.split(',')]
    df = pd.read_excel(excel_file)
    sql_queries = []
    for col_name in column_names:
        msg = f"Selecione as colunas do Excel para buscar os valores para '{col_name}':"
        selected_columns = gui.multchoicebox(msg, "Seleção de colunas", df.columns)
        while not selected_columns:
            response = gui.msgbox("Selecione pelo menos uma coluna!", "Aviso")
            selected_columns = gui.multchoicebox(msg, "Seleção de colunas", df.columns)

        # armazenar os valores de cada coluna em uma lista
        values = []
        for sel_col in selected_columns:
            values.append(df[sel_col].tolist())

        # loop sobre as listas para construir a string de atualização
        for vals in zip(*values):
            update_str = ""
            for i, sel_col in enumerate(selected_columns):
                value = vals[i]
                if pd.isna(value):
                    update_str += f"{col_name} = NULL"
                elif isinstance(value, str):
                    update_str += f"{col_name} = '{value}'"
                else:
                    update_str += f"{col_name} = {value}"
                if i < len(selected_columns) - 1:
                    update_str += ", "
            sql_query = f"UPDATE {table_name} SET {update_str}"
            sql_queries.append(sql_query)

    gui.textbox("Resultado", "SQL Queries", "\n".join(sql_queries))
    gui.msgbox("UPDATE concluído.", "Mensagem")

# Menu principal
options = {
    "UPDATE": update_table,
    "SAIR": sys.exit
}

while True:
    user_choice = gui.buttonbox("Escolha uma opção:", "Menu principal", ["UPDATE", "SAIR"])
    if not user_choice or user_choice == "SAIR":
        sys.exit()
    excel_file = gui.fileopenbox("Selecione o arquivo Excel:")
    if not excel_file:
        gui.msgbox("Arquivo não selecionado!", "Aviso")
        continue
    options[user_choice](excel_file)
