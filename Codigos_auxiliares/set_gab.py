import os
from pathlib import Path
import win32com.client as win32

def format_word_in_two_columns(input_file, output_file):
    # Verifica se o arquivo existe
    if not os.path.exists(input_file):
        print(f"Erro: O arquivo Word não foi encontrado em: {input_file}")
        return

    try:
        # Inicializa o Word
        word_app = win32.gencache.EnsureDispatch("Word.Application")
        word_app.Visible = False  # Oculta o Word
        doc = word_app.Documents.Open(input_file)

        # Aplica formatação de duas colunas na primeira seção
        section = doc.Sections(1)
        section.PageSetup.TextColumns.SetCount(2)  # Define 2 colunas

        # Ajusta o espaço entre as colunas, se necessário (em pontos)
        section.PageSetup.TextColumns.Spacing = 36  # 36 pontos = 0,5 polegada

        # Salva o documento formatado
        doc.SaveAs(output_file)
        print(f"Documento formatado salvo em: {output_file}")
    except Exception as e:
        print(f"Ocorreu um erro ao formatar o documento: {e}")
    finally:
        # Fecha o documento e o Word
        doc.Close(SaveChanges=False)
        word_app.Quit()

