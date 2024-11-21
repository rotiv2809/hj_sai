import win32com.client as win32
from pathlib import Path

def limit_margins_to_half_horizontal(word_file, output_file):
    word_app = None
    doc = None  # Inicializa o documento como None
    try:
        # Inicializa o Word
        word_app = win32.gencache.EnsureDispatch("Word.Application")
        word_app.Visible = False  # Oculta o Word
        doc = word_app.Documents.Open(word_file)

        # Obtém a largura total da página
        page_width = doc.PageSetup.PageWidth

        # Define margens para limitar o texto até a metade horizontal
        doc.PageSetup.TopMargin = 72  # Margem superior padrão (1 polegada)
        doc.PageSetup.BottomMargin = 72  # Margem inferior padrão
        doc.PageSetup.LeftMargin = 72  # Margem esquerda padrão
        doc.PageSetup.RightMargin = int(page_width / 2)  # Margem direita até o meio da página

        # Salva o novo documento
        doc.SaveAs(output_file)
        print(f"Documento salvo com margens ajustadas em: {output_file}")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        # Fecha o documento e o Word, se abertos
        if doc:
            doc.Close(SaveChanges=True)
        if word_app:
            word_app.Quit()