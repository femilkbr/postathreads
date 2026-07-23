import os
import sys
import time
import requests
from docx import Document
#from dotenv import load_dotenv

# Carrega o arquivo .env se existir localmente (no GitHub Actions lê direto dos Secrets)
#load_dotenv()

# Obtém as credenciais das variáveis de ambiente e remove espaços em branco extras
THREADS_USER_ID = os.getenv("THREADS_USER_ID", "").strip()
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "").strip()

def publicar_no_threads(texto: str) -> bool:
    """
    Publica um texto no Threads via API Graph da Meta com tratamento de codificação UTF-8.
    """
    if not THREADS_USER_ID or not ACCESS_TOKEN:
        print("Erro: As variáveis THREADS_USER_ID ou ACCESS_TOKEN não foram encontradas!")
        return False

    base_url = f"https://graph.threads.net/v1.0/{THREADS_USER_ID}"

    # 1. ETAPA DE CRIAÇÃO DO CONTÊINER (TEXTO)
    payload_creation = {
        'media_type': 'TEXT',
        'text': texto,
        'access_token': ACCESS_TOKEN
    }

    # Header explícito UTF-8 essencial para aceitar quebras de linha (\n) e emojis sem quebrar
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded; charset=utf-8'
    }

    res_creation = requests.post(
        f"{base_url}/threads",
        data=payload_creation,
        headers=headers
    )

    if res_creation.status_code != 200:
        print(f"Erro na chamada à API do Threads: {res_creation.text}")
        return False

    dados_resposta = res_creation.json()
    creation_id = dados_resposta.get("id")

    if not creation_id:
        print("Erro: ID de criação não retornado pela API do Threads.")
        return False

    print(f"Contêiner criado com sucesso! ID: {creation_id}")

    # Aguarda 5 segundos para que os servidores do Meta processem o contêiner
    time.sleep(5)

    # 2. ETAPA DE PUBLICAÇÃO DO CONTÊINER
    payload_publish = {
        'creation_id': creation_id,
        'access_token': ACCESS_TOKEN
    }

    res_publish = requests.post(
        f"{base_url}/threads_publish",
        data=payload_publish,
        headers=headers
    )

    if res_publish.status_code == 200:
        print("✅ Post publicado com sucesso no Threads!")
        return True
    else:
        print(f"Erro ao publicar no Threads: {res_publish.text}")
        return False

def processar_e_postar_do_docx(caminho_docx: str):
    """
    Lê o arquivo .docx, envia o primeiro post separado por '---' para o Threads,
    e regrava o arquivo .docx sem a mensagem enviada.
    """
    if not os.path.exists(caminho_docx):
        print(f"Erro: O arquivo '{caminho_docx}' não foi encontrado.")
        sys.exit(1)

    doc = Document(caminho_docx)

    # Junta o texto mantendo as quebras de linha dos parágrafos
    texto_completo = "\n".join([p.text for p in doc.paragraphs])

    # Separa os posts usando o delimitador '---'
    posts = [p.strip() for p in texto_completo.split("---") if p.strip()]

    if not posts:
        print("Nenhum post pendente encontrado no arquivo .docx!")
        return

    # Pega o primeiro post da fila
    post_atual = posts[0]

    print("\n--------------------------------------------------")
    print("📢 Publicando o post do dia:")
    print("--------------------------------------------------")
    print(f"{post_atual}")
    print("--------------------------------------------------\n")

    sucesso = publicar_no_threads(post_atual)

    # Se a publicação teve êxito, remove a mensagem publicada e regrava o arquivo
    if sucesso:
        posts_restantes = posts[1:]

        novo_doc = Document()
        for idx, post in enumerate(posts_restantes):
            for linha in post.split("\n"):
                novo_doc.add_paragraph(linha)
            # Adiciona o separador se ainda houverem posts na fila
            if idx < len(posts_restantes) - 1:
                novo_doc.add_paragraph("---")

        novo_doc.save(caminho_docx)
        print(f"📝 Arquivo '{caminho_docx}' atualizado! (Restam {len(posts_restantes)} mensagens no arquivo).")
    else:
        print("⚠️ Falha ao publicar. O arquivo .docx não foi modificado.")
        sys.exit(1)

if __name__ == "__main__":
    # Permite passar o nome do arquivo via linha de comando ou assume 'arquivo.docx'
    nome_arquivo = sys.argv[1] if len(sys.argv) > 1 else "arquivo.docx"
    processar_e_postar_do_docx(nome_arquivo)
