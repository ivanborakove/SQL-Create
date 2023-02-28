import pandas as pd
import easygui as gui
import sys

def get_table_name():
    table_name = gui.enterbox("Digite o nome da tabela:")
    if not table_name:
        gui.msgbox("Tabela não informada!", "Aviso")
        return get_table_name()
    return table_name

def get_column_names():
    column_names = gui.enterbox("Digite os nomes das colunas (separados por vírgula):")
    if not column_names:
        gui.msgbox("Colunas não informadas!", "Aviso")
        return get_column_names()
    column_names = [col.strip() for col in column_names.split(',')]
    return column_names

def get_selected_columns(df, col_name):
    selected_columns = gui.multchoicebox(f"Selecione as colunas do Excel para buscar os valores para '{col_name}':", 
                                          "Seleção de colunas", df.columns)
    if not selected_columns:
        gui.msgbox("Selecione pelo menos uma coluna!", "Aviso")
        return get_selected_columns(df, col_name)
    return selected_columns

def get_values(df, selected_columns):
    values = []
    for sel_col in selected_columns:
        values.append(df[sel_col].tolist())
    return values

def build_update_string(column_names, updates_dict, i):
    update_str = ""
    for col_name in column_names:
        update_str += f"{col_name} = {updates_dict[col_name][i]}, "
    return update_str[:-2]

def update_table(excel_file):
    df = pd.read_excel(excel_file)
    table_name = get_table_name()
    column_names = get_column_names()
    sql_queries = []
    updates_dict = {col: [] for col in column_names}
    
    for col_name in column_names:
        selected_columns = get_selected_columns(df, col_name)
        values = get_values(df, selected_columns)

        for vals in zip(*values):
            for i, sel_col in enumerate(selected_columns):
                value = vals[i]
                if pd.isna(value):
                    updates_dict[col_name].append("NULL")
                elif isinstance(value, str):
                    updates_dict[col_name].append(f"'{value}'")
                else:
                    updates_dict[col_name].append(str(value))

    for i in range(len(updates_dict[column_names[0]])):
        update_str = build_update_string(column_names, updates_dict, i)
        sql_query = f"UPDATE {table_name} SET {update_str}"
        sql_queries.append(sql_query)

    gui.textbox("Resultado", "SQL Queries", "\n".join(sql_queries))
    gui.msgbox("UPDATE concluído.", "Mensagem")

def main_menu():
    user_choice = gui.buttonbox("Escolha uma opção:", "Menu principal", ["UPDATE", "SAIR"])
    if not user_choice or user_choice == "SAIR":
        sys.exit()
    excel_file = gui.fileopenbox("Selecione o arquivo Excel:")
    if not excel_file:
        gui.msgbox("Arquivo não selecionado!", "Aviso")
        return main_menu
    update_table(excel_file)
    return main_menu()

if __name__ == "__main__":
    main_menu()
