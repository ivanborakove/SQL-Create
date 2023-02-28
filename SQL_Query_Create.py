import pandas as pd
import easygui as gui
import sys

def get_user_input():
    table_name = gui.enterbox("Digite o nome da tabela:")
    column_names_str = gui.enterbox("Digite os nomes das colunas (separados por vírgula):")
    column_names = [col.strip() for col in column_names_str.split(',') if col.strip()]
    if not table_name or not column_names:
        gui.msgbox("Tabela ou colunas não informadas!", "Aviso")
        sys.exit()
    return table_name, column_names

def get_selected_columns(df, column_names):
    sql_queries = []
    for index, row in df.iterrows():
        set_columns = []
        for col_name in column_names:
            while True:
                msg = f"Selecione as colunas do Excel para buscar o valor para '{col_name}':"
                selected_columns = gui.multchoicebox(msg, "Seleção de colunas", df.columns)
                if not selected_columns:
                    response = gui.msgbox("Selecione pelo menos uma coluna!", "Aviso")
                else:
                    sel_col = selected_columns[0]
                    set_columns.append(f"{col_name} = {row[sel_col]}")
                    break
        set_str = ', '.join(set_columns)
        sql_query = f"UPDATE {table_name} SET {set_str}"
        sql_queries.append(sql_query)
    return sql_queries

def show_results(sql_queries):
    if not sql_queries:
        gui.msgbox("Nenhum registro selecionado para atualização!", "Aviso")
        sys.exit()
    gui.textbox("Resultado", "SQL Queries", "\n".join(sql_queries))
    gui.msgbox("UPDATE concluído.", "Mensagem")

# Menu principal
options = {
    "UPDATE": lambda f: show_results(get_selected_columns(pd.read_excel(f), column_names)),
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
    table_name, column_names = get_user_input()
    options[user_choice](excel_file)
