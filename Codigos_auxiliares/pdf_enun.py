import os
from pathlib import Path
import win32com.client as win32

def add_background_image_and_convert_to_pdf(word_file, image_path, pdf_output):
    # Verifica se os arquivos existem
    if not os.path.exists(word_file):
        print(f"Erro: O arquivo Word não foi encontrado em: {word_file}")
        return
    if not os.path.exists(image_path):
        print(f"Erro: A imagem não foi encontrada em: {image_path}")
        return

    try:
        # Inicializa o Word
        word_app = win32.gencache.EnsureDispatch("Word.Application")
        word_app.Visible = False  # Oculta o Word
        doc = word_app.Documents.Open(word_file)

        # Calcula o tamanho total da página
        page_width = doc.PageSetup.PageWidth
        page_height = doc.PageSetup.PageHeight

        # Define margens para limitar o texto até metade da página
        section = doc.Sections(1)
        section.PageSetup.TopMargin = 0
        section.PageSetup.BottomMargin = 72  # 1 polegada
        section.PageSetup.LeftMargin = 30
        section.PageSetup.RightMargin = page_width / 2 + 5 # Limita à metade da página

        # Adiciona a imagem de fundo a todas as páginas (usando cabeçalhos)
        for section in doc.Sections:
            header = section.Headers(1)  # Cabeçalho primário
            shape = header.Shapes.AddPicture(
                FileName=str(image_path),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=-30,  # Posicionado na borda esquerda
                Top=-40,  # Ajuste fino para deslocar a imagem levemente para cima
                Width=page_width,  # Largura total da página
                Height=page_height + 10  # Altura ligeiramente ajustada para evitar cortes
            )
            shape.ZOrder(5)  # Envia a imagem para o fundo

        # Salva o documento como PDF
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf_output),
            ExportFormat=17,  # 17 é o valor para PDF
            OpenAfterExport=False
        )
        print(f"Documento salvo como PDF em: {pdf_output}")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        # Fecha o documento e o Word
        doc.Close(SaveChanges=False)
        word_app.Quit()

