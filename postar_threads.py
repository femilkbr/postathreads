#!/usr/bin/env python3
"""
Posta o texto de um arquivo .docx no Threads (Meta), preservando parágrafos.

Como funciona:
- O .docx funciona como uma fila: você escreve um ou mais textos, e separa
  cada um deles com um parágrafo contendo apenas o separador (por padrão
  "---", configurável pela variável de ambiente THREADS_SEPARADOR).
- A cada execução, o script pega SOMENTE o primeiro bloco de texto (tudo
  antes do primeiro separador encontrado — ou até o fim do arquivo, se não
  houver separador) e publica no Threads.
- Depois de publicar com sucesso, o script apaga esse bloco (incluindo o
  separador) do arquivo .docx, para que ele não seja publicado de novo na
  próxima execução. O restante do arquivo permanece intacto.
- A API do Threads limita cada post a 500 caracteres. Se o bloco couber
  nesse limite, publica um único post. Se for maior, o texto é dividido em
  pedaços (sem quebrar um parágrafo no meio, sempre que possível) e
  publicado como uma sequência de posts encadeados (cada um respondendo ao
  anterior via reply_to_id), formando uma "thread" de verdade.

Exemplo de conteúdo do .docx:
    Primeiro parágrafo do primeiro post.
    Segundo parágrafo do primeiro post.
    ---
    Texto do segundo post, que só será publicado na próxima execução.
    ---
    Texto do terceiro post.

Pré-requisitos:
- pip install python-docx requests python-dotenv
- Uma conta de desenvolvedor Meta com um app que tenha as permissões
  threads_basic e threads_content_publish, e um token de acesso de usuário
  válido para a conta do Threads (THREADS_ACCESS_TOKEN) e o ID do usuário do
  Threads (THREADS_USER_ID).
- Um arquivo chamado ".env" na mesma pasta do script, com o conteúdo:
    THREADS_USER_ID=seu_id
    THREADS_ACCESS_TOKEN=seu_token

Uso:
    python postar_threads.py caminho/para/texto.docx
"""

import os
import sys
import time
import requests
from docx import Document
from dotenv import load_dotenv

# Carrega do .env se existir localmente; no GitHub Actions, lê das env vars do runner
load_dotenv()

THREADS_USER_ID = os.getenv("THREADS_USER_ID")
THREADS_ACCESS_TOKEN = os.getenv("THREADS_ACCESS_TOKEN")

GRAPH_API_BASE = "https://graph.threads.net/v1.0"
MAX_CHARS = 500  # limite atual de caracteres por post no Threads
SEPARADOR_PADRAO = "---"  # linha que marca o fim de um bloco de texto no .docx


# --------------------------------------------------------------------------
# 1. Extração do próximo bloco de texto do .docx
# --------------------------------------------------------------------------
def extrair_proximo_bloco(documento: Document, separador: str = SEPARADOR_PADRAO):
    """
    Percorre os parágrafos do documento, do início, e devolve:
    - lista de textos (não vazios) do primeiro bloco, até encontrar um
      parágrafo igual ao separador (ou até o fim do documento, se não houver
      separador);
    - lista dos objetos Paragraph correspondentes (incluindo o separador,
      se encontrado), para serem removidos do arquivo após a publicação.

    Se o documento estiver vazio (nada a publicar), devolve (None, None).
    """
    bloco_textos = []
    paragrafos_para_remover = []

    for paragrafo in documento.paragraphs:
        texto = paragrafo.text.strip()

        if texto == separador:
            paragrafos_para_remover.append(paragrafo)
            break

        paragrafos_para_remover.append(paragrafo)
        if texto:
            bloco_textos.append(texto)

    if not bloco_textos:
        return None, None

    return bloco_textos, paragrafos_para_remover


def montar_texto_completo(paragrafos: list[str]) -> str:
    """Junta os parágrafos preservando a separação entre eles."""
    return "\n\n".join(paragrafos)


def remover_paragrafos(paragrafos) -> None:
    """Remove os parágrafos informados da árvore XML do documento."""
    for paragrafo in paragrafos:
        elemento = paragrafo._element
        elemento.getparent().remove(elemento)


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
def publicar_docx_no_threads(
    caminho_docx: str,
    user_id: str,
    access_token: str,
    separador: str = SEPARADOR_PADRAO,
) -> list[str]:
    """
    Abre o .docx, pega apenas o PRÓXIMO bloco de texto (delimitado pelo
    separador) e publica no Threads. Se o bloco couber em um post, publica
    um único post; caso contrário, publica como posts encadeados (thread).

    Após a publicação ter sucesso, remove o bloco publicado (e o separador)
    do arquivo .docx e salva, para que não seja publicado de novo.
    """
    documento = Document(caminho_docx)
    bloco_textos, paragrafos_para_remover = extrair_proximo_bloco(documento, separador)

    if bloco_textos is None:
        print("Nenhum texto pendente para publicar no arquivo.")
        return []

    texto_completo = montar_texto_completo(bloco_textos)
    ids_publicados = []

    if len(texto_completo) <= MAX_CHARS:
        post_id = postar_texto(user_id, access_token, texto_completo)
        ids_publicados.append(post_id)
        print(f"Post único publicado com sucesso. ID: {post_id}")
    else:
        partes = dividir_em_posts(bloco_textos)
        print(f"Bloco excede {MAX_CHARS} caracteres. Será publicado como {len(partes)} posts encadeados.")

        reply_to_id = None
        for i, parte in enumerate(partes, start=1):
            post_id = postar_texto(user_id, access_token, parte, reply_to_id=reply_to_id)
            ids_publicados.append(post_id)
            print(f"Post {i}/{len(partes)} publicado. ID: {post_id}")
            reply_to_id = post_id
            time.sleep(3)  # respiro entre posts consecutivos

    # Só remove o bloco do arquivo depois que TODOS os posts foram
    # publicados com sucesso (se algo falhar antes, uma exceção interrompe
    # a execução e o texto permanece intacto no .docx para nova tentativa).
    remover_paragrafos(paragrafos_para_remover)
    documento.save(caminho_docx)
    print(f"Bloco publicado removido de '{caminho_docx}'.")

    return ids_publicados


# --------------------------------------------------------------------------
# 5. Execução via linha de comando
# --------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Uso: python postar_threads.py caminho/para/arquivo.docx")
        sys.exit(1)

    caminho_docx = sys.argv[1]

    # Carrega as variáveis definidas no arquivo .env (se existir) para o
    # ambiente do processo, sem sobrescrever variáveis já exportadas no shell.
    load_dotenv()

    user_id = os.environ.get("THREADS_USER_ID")
    access_token = os.environ.get("THREADS_ACCESS_TOKEN")

    if not user_id or not access_token:
        print(
            "THREADS_USER_ID e THREADS_ACCESS_TOKEN não encontrados.\n"
            "Crie um arquivo .env na mesma pasta do script com:\n"
            "  THREADS_USER_ID=seu_id\n"
            "  THREADS_ACCESS_TOKEN=seu_token\n"
            "(ou exporte essas variáveis no terminal antes de rodar o script)."
        )
        sys.exit(1)

    separador = os.environ.get("THREADS_SEPARADOR", SEPARADOR_PADRAO)

    try:
        publicar_docx_no_threads(caminho_docx, user_id, access_token, separador=separador)
    except requests.HTTPError as erro:
        print(f"Erro na chamada à API do Threads: {erro.response.text}")
        sys.exit(1)
    except Exception as erro:
        print(f"Erro: {erro}")
        sys.exit(1)


if __name__ == "__main__":
    main()
