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
    updates_dict = {col: [] for col in column_names}
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
            for i, sel_col in enumerate(selected_columns):
                value = vals[i]
                if pd.isna(value):
                    updates_dict[col_name].append("NULL")
                elif isinstance(value, str):
                    updates_dict[col_name].append(f"'{value}'")
                else:
                    updates_dict[col_name].append(str(value))

    # construir as consultas SQL
    for i in range(len(updates_dict[column_names[0]])):
        update_str = ""
        for col_name in column_names:
            update_str += f"{col_name} = {updates_dict[col_name][i]}, "
        sql_query = f"UPDATE {table_name} SET {update_str[:-2]}"
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
