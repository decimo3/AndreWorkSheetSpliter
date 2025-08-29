#!/bin/env python
''' Module to  '''
import os
import sys
import xlwt
import pandas
import openpyxl

if getattr(sys, 'frozen', False):
    BASE_FOLDER = os.path.dirname(sys.executable)
else:
    BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))

class InvalidFileException(Exception):
    ''' Custom exception class for specific error handling '''

def print_center_presentation(text: str, lenght: int) -> None:
    ''' Function to centralize text on defined lenght '''
    result = '# '
    result += text
    result += str(' ' * (len(result) - (lenght - 4)))
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
        raise FileNotFoundError(f'O arquivo {path_find} não foi encontrado!')

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
        file_path: str
        ) -> None:
    ''' Export a DataFrame to a legacy .xls file using xlwt '''
    print(f'Exportando planilha {file_path}...')
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Planilha1')
    # Write head
    for col, head in enumerate(dataframe.columns):
        sheet.write(0, col, head)
    # Write body
    for row_idx, row in enumerate(dataframe.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            sheet.write(row_idx, col_idx, value)
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
        directory_path: str
        ) -> pandas.DataFrame:
    ''' Function to create folder from distinct values and place a filtred dataframe '''
    folder_path = os.path.join(directory_path, str(distinct_value))
    file_path = os.path.join(folder_path, str(distinct_value) + '.xls')
    create_folder_if_not_exist(folder_path)
    dataframe = dataframe[dataframe[column_name] == distinct_value]
    export_dataframe_to_xls(dataframe, file_path)
    return dataframe

if __name__ == '__main__':
    print_header_presentation()
    check_if_folder_or_file_exist(sys.argv[1])
    if not sys.argv[1].lower().endswith('.xlsx'):
        raise InvalidFileException(f'O arquivo {sys.argv[1]} não é válido!')
    tax_table_filepath = os.path.join(BASE_FOLDER, 'ISS.xlsx')
    check_if_folder_or_file_exist(tax_table_filepath)
    tax_table_dataframe = get_dataframe_from_excel(tax_table_filepath, 'Planilha1')
    base_directory = os.path.dirname(sys.argv[1])
    dataframe = get_dataframe_from_excel(sys.argv[1], 'Lista_Pedidos')
    dataframe = dataframe[['Contrato', 'Codigo', 'NumeroConformidade',
        'DescricaoServico', 'Quantidade', 'ValorUnitario', 'ValorTotal',
        'CodigoLeiComp', 'Municipio', 'DomicilioFiscal', 'GrupoComprador',
        'DescricaoGrupoComprador', 'ProvedorDescricao', 'Posicao', 'CodigoBaremo',
        'DescricaoBaremo', 'CodigoOrdem', 'Estado', 'MotivoRecusa']]
    base_directory = os.path.join(base_directory, 'AMPLA')
    create_folder_if_not_exist(base_directory)
    base_directory = os.path.join(base_directory, 'COMERCIAL')
    create_folder_if_not_exist(base_directory)
    for x in get_distinct_list_from_dataframe_column(dataframe, 'CodigoLeiComp'):
        dataframe_filterby_codigoleicomp = create_folder_and_place_filtred_dataframe(dataframe, 'CodigoLeiComp', x, base_directory)
        base_directory1 = os.path.join(base_directory, str(x))
        for y in get_distinct_list_from_dataframe_column(dataframe_filterby_codigoleicomp, 'Municipio'):
            dataframe_filterby_municipio = create_folder_and_place_filtred_dataframe(dataframe_filterby_codigoleicomp, 'Municipio', y, base_directory1)
            base_directory2 = os.path.join(base_directory1, str(y))
            for z in get_distinct_list_from_dataframe_column(dataframe_filterby_municipio, 'NumeroConformidade'):
                dataframe_filterby_conformidade = create_folder_and_place_filtred_dataframe(dataframe_filterby_municipio, 'NumeroConformidade', z, base_directory2)
    print(f'Relatórios exportados em {base_directory}')
    input('Pressione qualquer tecla para finalizar.')
