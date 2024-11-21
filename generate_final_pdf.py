import win32com.client as win32
from pathlib import Path
from Codigos_auxiliares.pdf_enun import add_background_image_and_convert_to_pdf
from Codigos_auxiliares.pdf_gab import add_background_image_and_convert_to_pdf_gab
from Codigos_auxiliares.merge_pdf import merge_pdfs

#gab com form -> pdf, novo -> pdf, juntar, juntar com a capa

capa_file = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Documentos_Fixos/capa.pdf")
word_file = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Docx_Gerados/novo.docx")
image_path = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Documentos_Fixos/fundo_1.png")
pdf_output = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Documentos_Intermediarios/novo.pdf")



word_file_gab = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Docx_Gerados/gab_com_formatacao.docx")
pdf_output_gab = Path("C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Documentos_Intermediarios/gab.pdf")

output = "C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Documentos_Intermediarios/documento_junto.pdf"
output_final = "C:/Users/vitor/OneDrive/Documentos/IME Aulas/Projetos pessoais/hj_sai/Output_pdf/documento_final.pdf"

# Executa o script
add_background_image_and_convert_to_pdf(str(word_file), str(image_path), str(pdf_output))
add_background_image_and_convert_to_pdf_gab(str(word_file_gab), str(image_path), str(pdf_output_gab))
merge_pdfs(str(pdf_output),str(pdf_output_gab),str(output))
merge_pdfs(str(capa_file),str(output), str(output_final))

# Esse código aqui pega os dois documentos formatados, do gabarito e do enunciado e transforma em PDF, dps junta com a capa e sai tudo isso junto num resultado só em
# Output_pdf, onde estará o documento pronto.
