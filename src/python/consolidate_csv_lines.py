"""
Script para limpeza e padronização de arquivos de extração de dados do AutoCAD via comando DATAEXTRACTION.

Este script foi desenvolvido para tratar arquivos exportados do AutoCAD no formato `.txt` ou `.csv`,
com o objetivo de corrigir problemas comuns causados por quebras de linha inesperadas, tags RTF, e 
vírgulas presentes em textos que conflitam com o delimitador CSV.

Funcionalidades principais:
---------------------------
1. **correct_line_break**: Corrige quebras de linha indesejadas em arquivos de texto onde conteúdos de células
   foram divididos em múltiplas linhas, unindo-as com base na identificação de linhas que não começam com dígito.

2. **remove_rtf_tags**: Remove tags RTF e códigos de formatação de texto (como `{\*\...}` ou `\n`) que podem estar
   presentes em textos extraídos do AutoCAD, tornando o conteúdo legível e limpo.

3. **decode_unicode_escapes**: Decodifica sequências Unicode escapadas (ex: `\\u00e7` → `ç`) para texto UTF-8 legível.

4. **replace_commas**: Substitui vírgulas dentro de textos que deveriam ser números decimais ou expressões como
   "PISO EL + 23,070", evitando confusão com o delimitador CSV. A substituição é feita apenas em linhas que excedem
   o número de campos esperados, indicando que há vírgulas internas indesejadas.

5. **read_csv**: Lê arquivos `.csv` ou `.txt` com delimitador personalizado (padrão: tabulação).

6. **write_csv**: Escreve a saída limpa em um novo arquivo `.csv` com um delimitador seguro (padrão: pipe `|`), 
   evitando conflitos com vírgulas no conteúdo.

Recomendações de uso:
---------------------
- Utilize o formato **TXT (delimitado por tabulação)** ao exportar do AutoCAD com o comando `DATAEXTRACTION`, 
  para evitar conflitos causados por vírgulas em textos.

- Caso seja necessário usar CSV, o script `replace_commas` ajuda a mitigar os efeitos de vírgulas internas, 
  substituindo-as por pontos para preservar o significado decimal e manter a integridade dos campos.

Notas:
------
- O script sobrescreve o arquivo de entrada durante as etapas intermediárias.
- A versão final é salva com delimitador `|`, ideal para importação futura no Excel ou bancos de dados.
"""

import csv
import re
import codecs
import io

def correct_line_break(input_file_path):
    new_lines = []
    header = True
    
    with open(input_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    for line in lines:
        if header:
            new_lines.append(line)
            header = False
        elif line.strip():  # Check if line is not empty
            if not line[0].isdigit():
                new_lines[-1] = new_lines[-1].strip() + line
            else:
                new_lines.append(line)
    
    # Write the merged content to a new file
    with open(input_file_path, 'w', encoding='utf-8') as output_file:
        output_file.writelines(new_lines)

def remove_rtf_tags(input_file_path):
    new_lines = []
    pattern = r"\{\*?\\[^{}]+}|[{}]|\\\n?[A-Za-z]+\n?(?:-?\d+)?[ ]?"
    
    with open(input_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    for line in lines:
        if line.strip():  # Check if line is not empty
            cleaned_text = re.sub(pattern, "", line)
            cleaned_text = re.sub(r"\\'[0-9a-fA-F]{2}", "", cleaned_text)
            cleaned_text = re.sub(r"\\\n", "", cleaned_text)
            new_lines.append(cleaned_text)
    
    # Write the merged content to a new file
    with open(input_file_path, 'w', encoding='utf-8') as output_file:
        output_file.writelines(new_lines)

def decode_unicode_escapes(text):
    return codecs.decode(text, 'unicode_escape')

def replace_commas(input_file_path, delimiter = ",", debug = False):
    #1,Text,BdB201460-0000-V-FE0047.dwg,ROMANS,0,6,,,,,,,,,,,,,,,,,,,,,
    #1,Text,BdB201460-0000-V-FE0041.dwg,R80,PISO EL + 23,070,,,,,,,,,,,,,,,,,,,,,
    #1,MText,BdB201460-0000-V-FE0046.dwg,,,,,,,,,,,,,,,,,,,;,A PROPRIEDADE INTELECTUAL DESTE DESENHO ? TRATADA NO CONTRATO, SENDO QUE O FORNECEDOR AUTORIZA, DESDE J?, O USO DESTE DESENHO PELA VALE E/OU EMPRESAS DO GRUPO, INCLUSIVE O COMPARTILHAMENTO COM TERCEIROS PARA FINS DE MANUTEN??O, OPERA??O E AQUISI??O DE PE?AS DE REPOSI??O.,,,
    pattern  = r'(\b\d{1,}(?:,\d{1,})+(?:\.\d+)?\b)' # Match "0,6"
    pattern += r'|(\b[A-Za-z ]+ \+ \d{1,}(?:,\d{1,})+(?:\.\d+)?\b)' # Match "PISO EL + 23,070"
    pattern = re.compile(pattern)
    
    def comma_replacer(match):
        #print(f'DEBUG: {match.group().replace(',', '.')}')
        # Replace commas with a chosen character or format, e.g., remove or change to dot
        return match.group().replace(',', '.')  # Example: remove commas
    
    number_of_fields = 0
    new_lines = []
    line_counter = 1
    
    with open(input_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
        #new_lines.append(header) # Append the header.
        number_of_fields = len(lines[0].split(delimiter)) # get the number of available fields
    
    for line in lines:
        fields = line.split(delimiter)
        if line.strip(): # Check if line is not empty
            if len(fields) > number_of_fields: # Check if the field index is within the range of available fields
                match = pattern.match(line)
                cleaned_text = pattern.sub(comma_replacer, line)
                new_lines.append(cleaned_text)
                #print(f'DEBUG: line_counter {line_counter} number_of_fields {number_of_fields} len(fields) {len(fields)} fields {fields}')
                if debug:
                    print(f'DEBUG: {",".join(fields)}')
            else:
                new_lines.append(line)
            
            line_counter += 1
    
    # Write the merged content to a new file
    with open(input_file_path, 'w', encoding='utf-8') as output_file:
        output_file.writelines(new_lines)

def read_csv(input_file_path, delimiter='\t'):
    with open(input_file_path, 'r', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile, delimiter=delimiter)
        return list(reader)  # Read and return all rows at once

def write_csv(processed_lines, output_file_path, delimiter='|'):
    with open(output_file_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile, quoting=csv.QUOTE_ALL, delimiter=delimiter)
        writer.writerows(processed_lines)

# Example usage
input_path = r'C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\01_projetos\1_em_desenvolvimento\BdB201460\Projeto\DataExtraction\BdB201460 - DataExtraction - 2025-07-08\BdB201460 - DataExtraction - 2025-07-08.csv'

correct_line_break(input_path)

remove_rtf_tags(input_path)

replace_commas(input_file_path = input_path, delimiter = ",", debug = False)

replace_commas(input_file_path = input_path, delimiter = ",", debug = True)

write_csv(read_csv(input_path, delimiter='\t'), input_path, delimiter='|')

input("Press enter to continue...")



# Extrai os dados do autocad (DataExtraction)
# Processa os dados no Excel.
# Copia todos os tags (linhas, válvulas, tie-ins, etc.) para a planilha Sheet1. A formatação deve ser do tupo texto, ou o PROCV pode ter problemas para encontrar os valores.
# Filtra a coluna "Verificação DataExtraction->GerenciamentoTag" da tabela "DataExtraction" com os valores "#N/A". Nessa coluna temos os tags nosso.
# Rode o macro ProcurarEColorirLinhas para marcar as linhas na planilha GerenciamentoTag que deverão ser deletadas
# Verifique as linhas marcadas para remoção em relação a planilha de ID, pois podem haver tags nessa planilha que não devem ser removidos.
