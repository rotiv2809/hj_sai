import win32com.client as win32
from pathlib import Path
from Codigos_auxiliares.limit_margins import limit_margins_to_half_horizontal
from Codigos_auxiliares.set_gab import format_word_in_two_columns
from Codigos_auxiliares.rename import rename_only_file_in_folder
from Google_Drive.puxar_arq import main
from clear import clear_folder
from clear import folders_to_clear

import os

# Limpa todas as pastas antes do download
for folder in folders_to_clear:
    clear_folder(folder)

print("Pastas limpas com sucesso!")

#importando os arquivos do google drive

# IDs das pastas no Google Drive e seus nomes personalizados
folder_data = {
    "1chJtKUNdd0ip2twLqf42Z6K98EHUuXEY": "capa",
    "1nxvMceD_5QewLjNMp8YNN2B_XwD8hJk9": "fundo",
    "1CnpMtGRsLEeX58deJmM_dM8-uINNR5-Y": "enunciado",
    "1kQbWNlo1hpqXixNgxTkf0RpMZcuc6QAC": "gabarito"
}

# Baixar arquivos de cada pasta
for folder_id, folder_name in folder_data.items():
    print(f"Baixando arquivos da pasta {folder_name}...")
    main(folder_id, custom_folder_name=folder_name)

#renomear os arquivos para que fique tudo certo

script_dir = os.path.dirname(os.path.abspath(__file__))
folder_path_1 = os.path.join(script_dir,'downloads/capa')
folder_path_2 = os.path.join(script_dir,'downloads/fundo')
folder_path_3 = os.path.join(script_dir,'downloads/enunciado')
folder_path_4 = os.path.join(script_dir,'downloads/gabarito')

rename_only_file_in_folder(folder_path_1,'capa.pdf')
rename_only_file_in_folder(folder_path_2,'fundo_1.png')
rename_only_file_in_folder(folder_path_3,'teste.docx')
rename_only_file_in_folder(folder_path_4,'gab.docx')

script_dir = os.path.dirname(os.path.abspath(__file__))
word_file = os.path.join(folder_path_3, 'teste.docx')
gab_file = os.path.join(folder_path_4, 'gab.docx')
output_file = os.path.join(script_dir, 'Docx_Gerados/novo.docx')
output_file_gab = os.path.join(script_dir, 'Docx_Gerados/gab_com_formatacao.docx')


# Ajusta as margens no novo arquivo
limit_margins_to_half_horizontal(str(word_file),str(output_file))
format_word_in_two_columns(str(gab_file),str(output_file_gab))

# Esse código aqui gera os documentos Word formatados em Docx_Gerados, é necessário ajustar o tamanho das imagens para que não estrapolem a margem no documento final!

