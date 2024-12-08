from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials
import io
import os

# Função para autenticar no Google Drive
def authenticate_google_drive():
    ## Obtenha o diretório do script
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Caminho para o arquivo credentials.json
    credentials_path = os.path.join(script_dir, 'credentials.json')
    
    # Carregar as credenciais
    creds = Credentials.from_service_account_file(credentials_path)
    
    return build('drive', 'v3', credentials=creds)

# Função para listar arquivos em uma pasta específica
def list_files_in_folder(service, folder_id):
    results = service.files().list(
        q=f"'{folder_id}' in parents",  # Filtra arquivos pela pasta
        spaces='drive',
        fields="files(id, name)"
    ).execute()
    return results.get('files', [])

# Função para fazer download de um arquivo
def download_file(service, file_id, destination):
    request = service.files().get_media(fileId=file_id)
    with io.FileIO(destination, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}% concluído.")
    print(f"Arquivo baixado: {destination}")

# Função principal
def main():
    # Autenticação
    service = authenticate_google_drive()

    # ID da pasta no Google Drive (substitua pelo ID da sua pasta)
    folder_id = "1XVgzu0hQIH-3fJkB7mbWQ3tcj3JWsIkg"

    # Listar arquivos na pasta
    print("Listando arquivos na pasta...")
    files = list_files_in_folder(service, folder_id)

    if not files:
        print("Nenhum arquivo encontrado na pasta.")
        return

    # Criar uma pasta local para armazenar os downloads
    os.makedirs("downloads", exist_ok=True)

    # Baixar todos os arquivos da pasta
    for file in files:
        file_name = file['name']
        file_id = file['id']
        print(f"Baixando arquivo: {file_name}...")
        download_file(service, file_id, f"./downloads/{file_name}")

    print("Todos os arquivos foram baixados com sucesso!")

# Executar o programa
if __name__ == "__main__":
    main()
