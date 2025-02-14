import win32com.client as win32
from pathlib import Path
from Codigos_auxiliares.pdf_enun import add_background_image_and_convert_to_pdf
from Codigos_auxiliares.pdf_gab import add_background_image_and_convert_to_pdf_gab
from Codigos_auxiliares.merge_pdf import merge_pdfs
import os

#gab com form -> pdf, novo -> pdf, juntar, juntar com a capa

script_dir = os.path.dirname(os.path.abspath(__file__))
capa_file = os.path.join(script_dir, 'downloads/capa/capa.pdf')
word_file = os.path.join(script_dir, 'Docx_Gerados/novo.docx')
image_path = os.path.join(script_dir, 'downloads/fundo/fundo_1.png')
pdf_output = os.path.join(script_dir, 'Documentos_Intermediarios/novo.pdf')
word_file_gab = os.path.join(script_dir, 'Docx_Gerados/gab_com_formatacao.docx')
word_file_gab_enum = os.path.join(script_dir, 'Docx_Gerados/gab_enum_com_formatacao.docx')
pdf_output_gab = os.path.join(script_dir, 'Documentos_Intermediarios/gab.pdf')
pdf_output_gab_enum = os.path.join(script_dir, 'Documentos_Intermediarios/gab_enum.pdf')
output = os.path.join(script_dir, 'Documentos_Intermediarios/documento_junto.pdf')
output_int = os.path.join(script_dir, 'Documentos_Intermediarios/documento_int.pdf')
output_final = os.path.join(script_dir, 'Output_pdf/documento_final.pdf')


# Executa o script
add_background_image_and_convert_to_pdf(str(word_file), str(image_path), str(pdf_output))
add_background_image_and_convert_to_pdf_gab(str(word_file_gab), str(image_path), str(pdf_output_gab))
add_background_image_and_convert_to_pdf_gab(str(word_file_gab_enum), str(image_path), str(pdf_output_gab_enum))
merge_pdfs(str(pdf_output),str(pdf_output_gab),str(output))
merge_pdfs(str(output),str(pdf_output_gab_enum), str(output_int))
merge_pdfs(str(capa_file),str(output_int), str(output_final))

# Esse código aqui pega os dois documentos formatados, do gabarito e do enunciado e transforma em PDF, dps junta com a capa e sai tudo isso junto num resultado só em
# Output_pdf, onde estará o documento pronto.
