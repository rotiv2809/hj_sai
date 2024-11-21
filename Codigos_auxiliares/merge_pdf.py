from PyPDF2 import PdfMerger

def merge_pdfs(pdf1_path, pdf2_path, output_path):
    try:
        # Cria o objeto PdfMerger
        merger = PdfMerger()

        # Adiciona os dois PDFs
        merger.append(pdf1_path)
        merger.append(pdf2_path)

        # Salva o PDF combinado no caminho de saída
        merger.write(output_path)
        merger.close()

        print(f"PDFs combinados com sucesso! Salvo em: {output_path}")
    except Exception as e:
        print(f"Ocorreu um erro ao combinar os PDFs: {e}")

