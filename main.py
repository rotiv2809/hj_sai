import win32com.client as win32
from pathlib import Path
from Codigos_auxiliares.limit_margins import limit_margins_to_half_horizontal
from Codigos_auxiliares.set_gab import format_word_in_two_columns

# Caminhos dos arquivos
word_file = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Docx_Base/teste.docx")
output_file = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Docx_Gerados/novo.docx")
gab_file = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Docx_Base/gab.docx")
output_file_gab = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Docx_Gerados/gab_com_formatacao.docx")


# Ajusta as margens no novo arquivo
limit_margins_to_half_horizontal(str(word_file),str(output_file))
format_word_in_two_columns(str(gab_file),str(output_file_gab))

# Esse código aqui gera os documentos Word formatados em Docx_Gerados, é necessário ajustar o tamanho das imagens para que não estrapolem a margem no documento final!

