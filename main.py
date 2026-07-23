# Insira seus dados do Threads
THREADS_USER_ID = "37071976939115477"
ACCESS_TOKEN = "THAAPEI9zx65xBYlpreEI2QjNwWC0xZAXdpakZASTzd4cFYxQTZAuWU0tUElHREMtc1JMVXQ0NHlsUTYzRG04YndWd3BvWnNiOXp6X0RCRkxMeHhQRFV0NnpHdjhlQWNHeElMX2l5c2lfVjNtc2FubW5vcWlad3lrX185UHZAFeUc3TG5NTERUMUtZAczQtZAWZA1blQ3Sk0xcVUyNXk5bEc1VTkyOGFLaXdPdwZDZD"

import os
import sys
import time
import requests
from docx import Document

GRAPH_API_BASE = "https://graph.threads.net/v1.0"
MAX_CHARS = 500  # limite atual de caracteres por post no Threads


# --------------------------------------------------------------------------
# 1. Extração do texto do .docx
# --------------------------------------------------------------------------
def extrair_paragrafos(caminho_docx: str) -> list[str]:
    """Lê o .docx e retorna a lista de parágrafos não vazios, na ordem original."""
    documento = Document(caminho_docx)
    paragrafos = [p.text.strip() for p in documento.paragraphs]
    paragrafos = [p for p in paragrafos if p]  # remove parágrafos em branco
    if not paragrafos:
        raise ValueError("Nenhum texto encontrado no documento.")
    return paragrafos


def montar_texto_completo(paragrafos: list[str]) -> str:
    """Junta os parágrafos preservando a separação entre eles."""
    return "\n\n".join(paragrafos)


# --------------------------------------------------------------------------
# 2. Divisão em posts (respeitando o limite de caracteres do Threads)
# --------------------------------------------------------------------------
def dividir_em_posts(paragrafos: list[str], limite: int = MAX_CHARS) -> list[str]:
    """
    Agrupa parágrafos em blocos que caibam no limite de caracteres.
    Se um único parágrafo for maior que o limite, ele é quebrado por
    palavras como último recurso.
    """
    posts = []
    bloco_atual = ""

    for paragrafo in paragrafos:
        candidato = f"{bloco_atual}\n\n{paragrafo}" if bloco_atual else paragrafo

        if len(candidato) <= limite:
            bloco_atual = candidato
            continue

        # o parágrafo não cabe junto com o bloco atual: fecha o bloco atual
        if bloco_atual:
            posts.append(bloco_atual)
            bloco_atual = ""

        if len(paragrafo) <= limite:
            bloco_atual = paragrafo
        else:
            # parágrafo sozinho já excede o limite: quebra por palavras
            posts.extend(_quebrar_por_palavras(paragrafo, limite))

    if bloco_atual:
        posts.append(bloco_atual)

    return posts


def _quebrar_por_palavras(texto: str, limite: int) -> list[str]:
    palavras = texto.split()
    partes = []
    parte_atual = ""
    for palavra in palavras:
        candidato = f"{parte_atual} {palavra}".strip()
        if len(candidato) <= limite:
            parte_atual = candidato
        else:
            if parte_atual:
                partes.append(parte_atual)
            parte_atual = palavra
    if parte_atual:
        partes.append(parte_atual)
    return partes


# --------------------------------------------------------------------------
# 3. Chamadas à API do Threads
# --------------------------------------------------------------------------
def criar_container(user_id: str, access_token: str, texto: str, reply_to_id: str | None = None) -> str:
    """Cria um container de mídia do tipo TEXT e devolve o creation_id."""
    url = f"{GRAPH_API_BASE}/{user_id}/threads"
    payload = {
        "media_type": "TEXT",
        "text": texto,
        "access_token": access_token,
    }
    if reply_to_id:
        payload["reply_to_id"] = reply_to_id

    resposta = requests.post(url, data=payload, timeout=30)
    resposta.raise_for_status()
    return resposta.json()["id"]


def publicar_container(user_id: str, access_token: str, creation_id: str) -> str:
    """Publica o container criado e devolve o ID do post publicado."""
    url = f"{GRAPH_API_BASE}/{user_id}/threads_publish"
    payload = {
        "creation_id": creation_id,
        "access_token": access_token,
    }
    resposta = requests.post(url, data=payload, timeout=30)
    resposta.raise_for_status()
    return resposta.json()["id"]


def postar_texto(user_id: str, access_token: str, texto: str, reply_to_id: str | None = None) -> str:
    """Cria e publica um único post de texto. Devolve o ID do post publicado."""
    creation_id = criar_container(user_id, access_token, texto, reply_to_id=reply_to_id)
    # pequena espera recomendada pela documentação antes de publicar
    time.sleep(2)
    post_id = publicar_container(user_id, access_token, creation_id)
    return post_id


# --------------------------------------------------------------------------
# 4. Orquestração
# --------------------------------------------------------------------------
def publicar_docx_no_threads(caminho_docx: str, user_id: str, access_token: str) -> list[str]:
    """
    Lê o .docx e publica no Threads. Se o texto couber em um post, publica
    um único post. Caso contrário, publica uma sequência encadeada de posts
    (thread), respeitando os parágrafos.
    """
    paragrafos = extrair_paragrafos(caminho_docx)
    texto_completo = montar_texto_completo(paragrafos)

    ids_publicados = []

    if len(texto_completo) <= MAX_CHARS:
        post_id = postar_texto(user_id, access_token, texto_completo)
        ids_publicados.append(post_id)
        print(f"Post único publicado com sucesso. ID: {post_id}")
        return ids_publicados

    partes = dividir_em_posts(paragrafos)
    print(f"Texto excede {MAX_CHARS} caracteres. Será publicado como {len(partes)} posts encadeados.")

    reply_to_id = None
    for i, parte in enumerate(partes, start=1):
        post_id = postar_texto(user_id, access_token, parte, reply_to_id=reply_to_id)
        ids_publicados.append(post_id)
        print(f"Post {i}/{len(partes)} publicado. ID: {post_id}")
        reply_to_id = post_id
        time.sleep(3)  # respiro entre posts consecutivos

    return ids_publicados


# --------------------------------------------------------------------------
# 5. Execução via linha de comando
# --------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Uso: python postar_threads.py caminho/para/arquivo.docx")
        sys.exit(1)

    caminho_docx = sys.argv[1]

    user_id = os.environ.get("THREADS_USER_ID")
    access_token = os.environ.get("THREADS_ACCESS_TOKEN")

    if not user_id or not access_token:
        print(
            "Defina as variáveis de ambiente THREADS_USER_ID e "
            "THREADS_ACCESS_TOKEN antes de executar o script."
        )
        sys.exit(1)

    try:
        publicar_docx_no_threads(caminho_docx, user_id, access_token)
    except requests.HTTPError as erro:
        print(f"Erro na chamada à API do Threads: {erro.response.text}")
        sys.exit(1)
    except Exception as erro:
        print(f"Erro: {erro}")
        sys.exit(1)


if __name__ == "__main__":
    main()