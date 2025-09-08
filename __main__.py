#!/bin/env python
''' Module to  '''
import os
import sys
import json
from tkinter import messagebox
from tkinter import filedialog
import xlwt
import pandas
import openpyxl

if getattr(sys, 'frozen', False):
    BASE_FOLDER = os.path.dirname(sys.executable)
else:
    BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))

class InvalidFileException(Exception):
    ''' Custom exception class for specific error handling '''

class RelationshipException(Exception):
    ''' Custom exception class for relationship issues '''

def print_center_presentation(text: str, lenght: int) -> None:
    ''' Function to centralize text on defined lenght '''
    spaces = ((lenght - len(text)) / 2) - 4
    result = '# '
    result += str(' ' * int(spaces))
    result += text
    result += str(' ' * int(spaces))
    result += ' #'
    print(result)

def print_header_presentation() -> None:
    ''' Function to print startup banner '''
    presentation_lenght = 100
    print()
    print('#' * presentation_lenght)
    print_center_presentation('Bem vindos ao programa de separação de planilhas do MestreRuan', presentation_lenght)
    print_center_presentation('Repositório: https://github.com/decimo3/AndreWorkSheetSpliter', presentation_lenght)
    print('#' * presentation_lenght)
    print()


def check_if_folder_or_file_exist(
        path_find: str
        ) -> None:
    ''' Function to check if a folder or file exists and raise exception if not '''
    if not os.path.exists(path_find):
        show_popup_error(f'O arquivo {path_find} não foi encontrado!')
        raise FileNotFoundError()

def create_folder_if_not_exist(
        folder_path: str
        ) -> None:
    ''' Function to creates a folder if it does not already exist. '''
    if not os.path.exists(folder_path):
        print(f'Criada pasta: {folder_path}')
        os.mkdir(folder_path)

def get_dataframe_from_excel(
        file_path: str,
        sheetname: str
        ) -> pandas.DataFrame:
    ''' Function to get DataFrame from Excel file '''
    print(f'Obtendo informações do arquivo {file_path}...')
    workbook = openpyxl.open(file_path)
    worksheet = workbook[sheetname]
    rows = worksheet.iter_rows(values_only=True)
    head = next(rows)
    body = list(rows)
    print('Planilha carregada em memória com sucesso!')
    return pandas.DataFrame(body, columns=head)

def export_dataframe_to_xls(
        dataframe: pandas.DataFrame,
        file_path: str,
        total_value: float
        ) -> None:
    ''' Export a DataFrame to a legacy .xls file using xlwt '''
    print(f'Exportando planilha {file_path}...')
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Planilha1')
    # Write total
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
    style.pattern = pattern
    sheet.write(0, 6, total_value, style)
    # Write head
    for col, head in enumerate(dataframe.columns):
        sheet.write(1, col, head)
    # Write body
    for row_idx, row in enumerate(dataframe.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            sheet.write(row_idx + 1, col_idx, value)
    workbook.save(file_path)
    print(f'Planilha {file_path} exportada com sucesso!')

def get_distinct_list_from_dataframe_column(
        dataframe: pandas.DataFrame,
        column_name: str
        ) -> list:
    ''' Function to get distinct values of a column and create folders '''
    distinct_values = set(dataframe[column_name].to_list())
    print(f'Valores distintos obtido da coluna {column_name}: {distinct_values}.')
    return distinct_values

def create_folder_and_place_filtred_dataframe(
        dataframe: pandas.DataFrame,
        column_name: str,
        distinct_value: str,
        directory_path: str,
        sumarize_column: str,
        create_directory: bool = True
        ) -> pandas.DataFrame:
    ''' Function to create folder from distinct values and place a filtred dataframe '''
    if create_directory:
        folder_path = os.path.join(directory_path, str(distinct_value))
    else:
        folder_path = directory_path
    file_path = os.path.join(folder_path, str(distinct_value) + '.xls')
    create_folder_if_not_exist(folder_path)
    dataframe = dataframe[dataframe[column_name] == distinct_value]
    total = dataframe[sumarize_column].sum(numeric_only=True)
    export_dataframe_to_xls(dataframe, file_path, total)
    return dataframe

def show_popup_error(message: str) -> None:
    ''' Function to show a popup message about erros '''
    print('ERRO: ' + message)
    messagebox.showerror('Erro!', message=message)

def show_popup_info(message: str) -> None:
    ''' Function to show a popup message about erros '''
    print('INFO: ' + message)
    messagebox.showinfo('Info!', message=message)

def recursive_split_and_export(
        dataframe: pandas.DataFrame,
        column_names: list,
        base_directory: str,
        sumarize_column: str
    ) -> None:
    ''' Function to recursively split dataframe and export based on column_names '''
    if not column_names:
        return
    current_column = column_names[0]
    is_last = len(column_names) == 1
    distinct_values = get_distinct_list_from_dataframe_column(dataframe, current_column)
    for value in distinct_values:
        filtered_df = create_folder_and_place_filtred_dataframe(dataframe, current_column, value, base_directory, sumarize_column, create_directory = not is_last)
        next_directory = os.path.join(base_directory, str(value)) if not is_last else base_directory
        recursive_split_and_export(filtered_df, column_names[1:], next_directory, sumarize_column)

if __name__ == '__main__':
    print_header_presentation()
    configuration_filepath = os.path.join(BASE_FOLDER, 'AndreWorkSheetSpliter.json')
    check_if_folder_or_file_exist(configuration_filepath)
    with open(configuration_filepath, mode='r', encoding='utf8') as file:
        configs = json.load(file)
    filepath = sys.argv[1] if len(sys.argv) > 1 else filedialog.askopenfilename()
    check_if_folder_or_file_exist(filepath)
    if not filepath.lower().split('.')[-1] in {'xlsx', 'xlsm'}:
        show_popup_error(f'O arquivo {filepath} não é válido!')
        raise InvalidFileException()
    tax_table_filepath = os.path.join(BASE_FOLDER, 'ISS.xlsx')
    check_if_folder_or_file_exist(tax_table_filepath)
    tax_table_dataframe = get_dataframe_from_excel(tax_table_filepath, 'Planilha1')
    base_directory = os.path.dirname(filepath)
    dataframe = get_dataframe_from_excel(filepath, 'Lista_Pedidos')
    dataframe = pandas.merge(left=dataframe, right=tax_table_dataframe, how='left', left_on='Municipio', right_on='MUNICÍPIO')
    dataframe_with_null_values = dataframe[dataframe['ALÍQUOTA'].isnull()]
    if len(dataframe_with_null_values) != 0:
        show_popup_error(f'Há valores nulos no relacionamento!\n{dataframe_with_null_values['Municipio']}')
        raise RelationshipException()
    dataframe = dataframe[['Contrato', 'Codigo', 'NumeroConformidade',
        'DescricaoServico', 'Quantidade', 'ValorUnitario', 'ValorTotal',
        'CodigoLeiComp', 'ALÍQUOTA', 'Municipio', 'DomicilioFiscal', 'GrupoComprador',
        'DescricaoGrupoComprador', 'ProvedorDescricao', 'Posicao', 'CodigoBaremo',
        'DescricaoBaremo', 'CodigoOrdem', 'Estado', 'MotivoRecusa']]
    dataframe = dataframe.rename(columns={'ALÍQUOTA': 'ISS'})
    base_directory = os.path.join(base_directory, 'AMPLA')
    create_folder_if_not_exist(base_directory)
    base_directory = os.path.join(base_directory, 'COMERCIAL')
    create_folder_if_not_exist(base_directory)
    criteria = ['CodigoLeiComp', 'Municipio', 'NumeroConformidade']
    recursive_split_and_export(dataframe, criteria, base_directory)
    show_popup_info(f'Relatórios exportados em {base_directory}')
