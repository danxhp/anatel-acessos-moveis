# -*- coding: utf-8 -*-
"""
monitor_anatel_smp.py
=====================
Verifica, uma vez por dia, se a Anatel já publicou a competência mensal NOVA de
acessos móveis (Serviço Móvel Pessoal - SMP) no Portal de Dados Abertos.

Fonte confirmada (Inventario_de_Bases_de_Dados.csv da Anatel, item 2):
    - Dataset : "Acessos - Telefonia Móvel" (acessos do SMP)
    - slug/id : acessos-autorizadas-smp
    - Página  : https://dados.gov.br/dataset/acessos-autorizadas-smp
    - Painel  : https://informacoes.anatel.gov.br/paineis/acessos/telefonia-movel
    - Encoding: latin-1 (ISO-8859-1) com separador ";"

Estratégia EM CAMADAS para descobrir os recursos (a API do dados.gov.br passou a
exigir chave; por isso há fallback):
    A) CKAN legado   : /api/3/action/package_show?id=<slug>
    B) API nova      : /dados/api/publico/conjuntos-dados/<slug>
                       (envia header chave-api-dados-abertos se CHAVE_DADOS_GOV existir)
    C) Raspagem HTML : extrai links .csv/.zip da página do dataset (sem chave)

Não usa Selenium nem renderiza o painel JS. Apenas requests + pandas.

Uso:
    python monitor_anatel_smp.py            # execução normal (silenciosa se sem novidade)
    python monitor_anatel_smp.py --verbose  # mostra detalhes (recursos, colunas, etc.)
"""

import argparse
import io
import json
import os
import re
import smtplib
import struct
import sys
import time
import zipfile
import zlib
from datetime import datetime, date
from email.message import EmailMessage
import unicodedata

import requests
import pandas as pd

# Console do Windows costuma ser cp1252 e quebra com emojis/acentos. Força UTF-8.
for _fluxo in (sys.stdout, sys.stderr):
    try:
        _fluxo.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass  # versões antigas: segue sem reconfigurar

# ===========================================================================
# CONFIGURAÇÃO (ajuste aqui no topo)
# ===========================================================================

# Competência que estou esperando (formato "AAAA-MM"). Hoje: abril/2026.
COMPETENCIA_ALVO = "2026-04"

# Slug/id do dataset SMP no dados.gov.br.
DATASET_SLUG = "acessos-autorizadas-smp"

# Pasta onde os CSVs baixados ficam guardados (criada se não existir).
PASTA_DOWNLOAD = r"c:\dev\Acessos_Moveis\dados_anatel"

# Arquivos de apoio (ficam ao lado deste script).
_DIR = os.path.dirname(os.path.abspath(__file__))
ARQ_LOG = os.path.join(_DIR, "anatel_smp_log.txt")
ARQ_ESTADO = os.path.join(_DIR, "anatel_smp_estado.json")
ARQ_EMAIL_CONFIG = os.path.join(_DIR, "email_config.json")

# User-Agent de navegador comum (a Anatel/dados.gov.br às vezes bloqueia clientes "crus").
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Timeouts (conexão, leitura) e política de retry com backoff exponencial.
TIMEOUT = (15, 120)          # segundos: (conectar, ler)
MAX_TENTATIVAS = 3           # nº de tentativas por requisição
BACKOFF_BASE = 2             # 1ª espera 2s, depois 4s, depois 8s...

# Chave opcional do dados.gov.br (registro gratuito). Definida via variável de
# ambiente para não ficar hardcoded:  setx CHAVE_DADOS_GOV "sua-chave-aqui"
CHAVE_DADOS_GOV = os.environ.get("CHAVE_DADOS_GOV", "").strip()

# URL do painel, citada nos avisos quando a API parece defasada.
URL_PAINEL = "https://informacoes.anatel.gov.br/paineis/acessos/telefonia-movel"

# Tamanho do chunk ao ler o CSV em pedaços (evita estourar memória no consolidado).
CHUNKSIZE = 200_000

# Tipos de produto excluídos do total "manchete" da Anatel. O arquivo oficial
# Acessos_Telefonia_Movel_Total.csv NÃO conta M2M nem pontos de serviço; excluindo
# esses dois, o detalhe bate exatamente com o total oficial (acessos "pessoais").
EXCLUIR_TIPO_PRODUTO = {"M2M", "PONTO_DE_SERVICO"}

# Variável global ligada por --verbose.
VERBOSE = False


# ===========================================================================
# UTILIDADES GERAIS
# ===========================================================================

def log_verbose(msg):
    """Imprime apenas em modo --verbose (para diagnóstico)."""
    if VERBOSE:
        print(f"   · {msg}")


def normaliza(texto):
    """Minúsculas + sem acento + sem espaços nas pontas. Usado p/ casar nomes de coluna."""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return t.strip().lower()


def gravar_log(resultado):
    """Acrescenta uma linha no log local com data/hora da checagem + resultado."""
    carimbo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(ARQ_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{carimbo}] {resultado}\n")
    except OSError as e:
        print(f"⚠️  Não consegui escrever no log ({ARQ_LOG}): {e}")


def carregar_estado():
    """Lê o estado anterior (última competência vista) para decidir se houve novidade."""
    if os.path.exists(ARQ_ESTADO):
        try:
            with open(ARQ_ESTADO, "r", encoding="utf-8") as f:
                return json.load(f)
        except (OSError, json.JSONDecodeError):
            return {}
    return {}


def salvar_estado(estado):
    """Persiste o estado atual (competência vista + carimbo da última checagem)."""
    try:
        with open(ARQ_ESTADO, "w", encoding="utf-8") as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)
    except OSError as e:
        print(f"⚠️  Não consegui salvar o estado ({ARQ_ESTADO}): {e}")


def http_get(url, headers=None, stream=False):
    """
    GET com retry + backoff exponencial. Trata timeout/erros de rede com mensagem
    clara. Retorna o objeto Response (status 2xx) ou levanta a última exceção.
    """
    cabecalhos = {"User-Agent": USER_AGENT}
    if headers:
        cabecalhos.update(headers)

    ultima_excecao = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            resp = requests.get(url, headers=cabecalhos, timeout=TIMEOUT, stream=stream)
            resp.raise_for_status()
            return resp
        except requests.exceptions.HTTPError as e:
            # Erro HTTP (4xx/5xx): normalmente não adianta repetir 401/403/404.
            codigo = e.response.status_code if e.response is not None else "?"
            ultima_excecao = e
            log_verbose(f"HTTP {codigo} em {url}")
            if codigo in (401, 403, 404):
                break  # não insiste: precisa de chave / não existe
        except (requests.exceptions.Timeout,
                requests.exceptions.ConnectionError) as e:
            ultima_excecao = e
            log_verbose(f"Falha de rede ({type(e).__name__}) em {url}")
        except requests.exceptions.RequestException as e:
            ultima_excecao = e
            log_verbose(f"Erro de requisição em {url}: {e}")

        if tentativa < MAX_TENTATIVAS:
            espera = BACKOFF_BASE ** tentativa
            log_verbose(f"Tentativa {tentativa} falhou; aguardando {espera}s...")
            time.sleep(espera)

    raise ultima_excecao


# ===========================================================================
# CAMADA DE DESCOBERTA DOS RECURSOS
# ===========================================================================

def _camada_ckan_legado():
    """Camada A: CKAN clássico package_show."""
    url = f"https://dados.gov.br/api/3/action/package_show?id={DATASET_SLUG}"
    resp = http_get(url, headers={"Accept": "application/json"})
    dados = resp.json()
    recursos = []
    for r in dados.get("result", {}).get("resources", []):
        recursos.append({
            "titulo": r.get("name") or r.get("url", ""),
            "url": r.get("url", ""),
            "formato": (r.get("format") or "").upper(),
            "atualizado_em": r.get("last_modified") or r.get("created"),
        })
    return recursos


def _camada_api_nova():
    """Camada B: API nova do dados.gov.br (usa chave se disponível)."""
    url = f"https://dados.gov.br/dados/api/publico/conjuntos-dados/{DATASET_SLUG}"
    headers = {"Accept": "application/json"}
    if CHAVE_DADOS_GOV:
        headers["chave-api-dados-abertos"] = CHAVE_DADOS_GOV
        log_verbose("Usando CHAVE_DADOS_GOV no header da API nova.")
    else:
        log_verbose("CHAVE_DADOS_GOV não definida; tentando API nova sem chave.")
    resp = http_get(url, headers=headers)
    dados = resp.json()
    recursos = []
    # A API nova usa a chave "recursos" (lista). Cada recurso traz:
    #   link, titulo, formato e dataUltimaAtualizacaoArquivo (dd/mm/aaaa).
    lista = dados.get("recursos") or dados.get("resources") or []
    for r in lista:
        url = (r.get("link") or r.get("url") or "").replace("\\", "/").strip()
        # Só interessam recursos de dados (CSV/ZIP); ignora glossário PDF etc.
        if not url.lower().endswith((".csv", ".zip")):
            continue
        recursos.append({
            "titulo": r.get("titulo") or r.get("name") or url.rsplit("/", 1)[-1],
            "url": url,
            "formato": (r.get("formato") or r.get("format") or "").upper(),
            "atualizado_em": (r.get("dataUltimaAtualizacaoArquivo")
                              or r.get("dataAtualizacao") or r.get("last_modified")),
        })
    return recursos


def _camada_html():
    """Camada C: raspa a página HTML do dataset e extrai links de CSV/ZIP."""
    paginas = [
        f"https://dados.gov.br/dados/conjuntos-dados/{DATASET_SLUG}",
        f"https://dados.gov.br/dataset/{DATASET_SLUG}",
    ]
    # Captura URLs absolutas terminando em .csv/.zip OU apontando para os PDA da Anatel.
    padrao = re.compile(
        r'https?://[^\s"\'<>]+?(?:\.csv|\.zip)'
        r'|https?://[^\s"\'<>]*anatel\.gov\.br/dadosabertos[^\s"\'<>]+',
        re.IGNORECASE,
    )
    vistos = set()
    recursos = []
    for pagina in paginas:
        try:
            resp = http_get(pagina, headers={"Accept": "text/html"})
        except requests.exceptions.RequestException:
            continue
        for url in padrao.findall(resp.text):
            url = url.rstrip('.,);')
            if url in vistos:
                continue
            vistos.add(url)
            fmt = "ZIP" if url.lower().endswith(".zip") else "CSV"
            recursos.append({
                "titulo": url.rsplit("/", 1)[-1],
                "url": url,
                "formato": fmt,
                "atualizado_em": None,  # HTML não traz a data de modificação confiável
            })
        if recursos:
            break  # já achou na primeira página que funcionou
    return recursos


def descobrir_recursos():
    """
    Tenta cada camada em ordem; retorna a primeira lista de recursos não vazia.
    Se todas falharem, retorna lista vazia (o chamador decide o que fazer).
    """
    camadas = [
        ("CKAN legado", _camada_ckan_legado),
        ("API nova", _camada_api_nova),
        ("Raspagem HTML", _camada_html),
    ]
    for nome, funcao in camadas:
        try:
            recursos = funcao()
            if recursos:
                print(f"🔎 Recursos obtidos via: {nome} ({len(recursos)} recurso(s)).")
                return recursos
            log_verbose(f"Camada '{nome}' retornou vazio.")
        except requests.exceptions.RequestException as e:
            log_verbose(f"Camada '{nome}' falhou: {e}")
        except (ValueError, KeyError) as e:
            log_verbose(f"Camada '{nome}' resposta inesperada: {e}")
    return []


def escolher_recurso(recursos):
    """
    Prioriza o CSV do ANO CORRENTE (evita baixar o histórico gigante). Se não houver
    arquivo por ano, cai no consolidado CSV. ZIP é aceito como último caso.
    """
    ano_corrente = str(date.today().year)
    csvs = [r for r in recursos if r["formato"] == "CSV" or r["url"].lower().endswith(".csv")]
    zips = [r for r in recursos if r["formato"] == "ZIP" or r["url"].lower().endswith(".zip")]

    # 1) CSV cujo título/URL contenha o ano corrente.
    for r in csvs:
        alvo = f"{r['titulo']} {r['url']}".lower()
        if ano_corrente in alvo and "movel" in normaliza(alvo):
            log_verbose(f"Escolhido CSV do ano {ano_corrente}: {r['titulo']}")
            return r
    for r in csvs:
        if ano_corrente in f"{r['titulo']} {r['url']}":
            log_verbose(f"Escolhido CSV contendo {ano_corrente}: {r['titulo']}")
            return r

    # 2) Qualquer CSV (provável consolidado).
    if csvs:
        log_verbose(f"Sem CSV por ano; usando CSV: {csvs[0]['titulo']}")
        return csvs[0]

    # 3) ZIP como fallback.
    if zips:
        log_verbose(f"Sem CSV; usando ZIP: {zips[0]['titulo']}")
        return zips[0]

    return None


# ===========================================================================
# DOWNLOAD E LEITURA DO CSV
# ===========================================================================

def _req_headers(extra=None):
    """Headers de navegador para os requests diretos (Range)."""
    h = {"User-Agent": USER_AGENT}
    if extra:
        h.update(extra)
    return h


def _zip_listar_membros(url):
    """
    Lê SÓ o diretório central do ZIP remoto (últimos ~3 MB) via HTTP Range, sem baixar
    os ~3 GB. Retorna {nome: {"method", "comp", "lho"}} de cada arquivo de dentro.
    """
    h = requests.head(url, headers=_req_headers(), timeout=TIMEOUT)
    h.raise_for_status()
    size = int(h.headers["Content-Length"])
    if h.headers.get("Accept-Ranges") != "bytes":
        raise RuntimeError("Servidor não suporta HTTP Range; não dá pra extrair membro.")
    inicio = max(0, size - 3_000_000)
    r = requests.get(url, headers=_req_headers({"Range": f"bytes={inicio}-"}),
                     timeout=TIMEOUT)
    r.raise_for_status()
    dados = r.content
    membros = {}
    sig = b"PK\x01\x02"           # assinatura do cabeçalho do diretório central
    i = 0
    while True:
        j = dados.find(sig, i)
        if j < 0:
            break
        method = struct.unpack("<H", dados[j + 10:j + 12])[0]
        comp = struct.unpack("<I", dados[j + 20:j + 24])[0]
        fnlen = struct.unpack("<H", dados[j + 28:j + 30])[0]
        extlen = struct.unpack("<H", dados[j + 30:j + 32])[0]
        cmtlen = struct.unpack("<H", dados[j + 32:j + 34])[0]
        lho = struct.unpack("<I", dados[j + 42:j + 46])[0]
        nome = dados[j + 46:j + 46 + fnlen].decode("latin-1", "replace")
        membros[nome] = {"method": method, "comp": comp, "lho": lho}
        i = j + 46 + fnlen + extlen + cmtlen
    if not membros:
        raise RuntimeError("Não consegui ler o diretório central do ZIP.")
    return membros


def _zip_extrair_membro(url, info, destino):
    """
    Baixa e descomprime UM membro do ZIP via Range, em streaming e com RETOMADA em caso
    de queda de conexão (a Anatel derruba o download). Grava o CSV cru em 'destino'.
    """
    lho = info["lho"]
    cab = requests.get(url, headers=_req_headers({"Range": f"bytes={lho}-{lho + 29}"}),
                       timeout=TIMEOUT)
    cab.raise_for_status()
    b = cab.content
    fnlen = struct.unpack("<H", b[26:28])[0]
    extlen = struct.unpack("<H", b[28:30])[0]
    off = lho + 30 + fnlen + extlen        # início dos dados comprimidos do membro
    comp = info["comp"]
    deflate = info["method"] == 8
    dec = zlib.decompressobj(-15) if deflate else None

    pos = 0
    falhas = 0
    with open(destino, "wb") as out:
        while pos < comp:
            try:
                r = requests.get(
                    url,
                    headers=_req_headers({"Range": f"bytes={off + pos}-{off + comp - 1}"}),
                    stream=True, timeout=TIMEOUT,
                )
                r.raise_for_status()
                for chunk in r.iter_content(chunk_size=1 << 20):   # 1 MB
                    if not chunk:
                        continue
                    pos += len(chunk)
                    out.write(dec.decompress(chunk) if deflate else chunk)
                falhas = 0
            except (requests.exceptions.ConnectionError,
                    requests.exceptions.ChunkedEncodingError,
                    requests.exceptions.Timeout) as e:
                falhas += 1
                if falhas > 10:
                    raise RuntimeError(f"Muitas quedas baixando o membro do ZIP: {e}")
                log_verbose(f"Conexão caiu em {pos:,}/{comp:,} bytes; retomando "
                            f"(tentativa {falhas}).")
                time.sleep(BACKOFF_BASE)
        if deflate:
            out.write(dec.flush())
    return destino


def _escolher_membros(membros):
    """
    Escolhe os CSVs de DADOS do ano corrente (ex.: '..._2026_1S.csv', '..._2026_2S.csv'),
    ignorando as variantes '_Colunas', totais e PDFs. Se não houver arquivo do ano,
    cai no pequeno '..._Total.csv' (só total mensal, sem quebra). Retorna lista de nomes.
    """
    ano = str(date.today().year)
    detalhe = [n for n in membros
               if re.search(rf"_{ano}_\dS\.csv$", n, re.IGNORECASE)
               and "colunas" not in n.lower()]
    if detalhe:
        return sorted(detalhe)
    total = [n for n in membros if n.lower().endswith("total.csv")
             and "pre_pos" not in n.lower()]
    if total:
        log_verbose(f"Sem arquivo de {ano}; usando fallback {total[0]} (só total).")
        return total[:1]
    raise RuntimeError(f"Nenhum CSV de dados do ano {ano} nem total encontrado no ZIP.")


def baixar(recurso):
    """
    Recurso ZIP: lê o índice remoto, escolhe o(s) CSV(s) do ano corrente e extrai SÓ
    esse(s) membro(s) via HTTP Range (ex.: ~168 MB do ano em vez dos ~3 GB do ZIP).
    Recurso CSV: baixa direto. Retorna LISTA de caminhos locais.
    """
    os.makedirs(PASTA_DOWNLOAD, exist_ok=True)
    url = recurso["url"]

    if not url.lower().endswith(".zip"):
        destino = os.path.join(PASTA_DOWNLOAD, "acessos_telefonia_movel.csv")
        print(f"⬇️  Baixando {url}")
        resp = http_get(url, stream=True)
        with open(destino, "wb") as f:
            for chunk in resp.iter_content(chunk_size=1 << 16):
                if chunk:
                    f.write(chunk)
        return [destino]

    print("🔎 Lendo índice do ZIP remoto (sem baixar os ~3 GB)...")
    membros = _zip_listar_membros(url)
    escolhidos = _escolher_membros(membros)
    caminhos = []
    for nome in escolhidos:
        info = membros[nome]
        destino = os.path.join(PASTA_DOWNLOAD, os.path.basename(nome))
        print(f"⬇️  Extraindo {nome} (~{info['comp'] / 1e6:.0f} MB compr.) via Range...")
        _zip_extrair_membro(url, info, destino)
        caminhos.append(destino)
    return caminhos


def detectar_formato(caminho):
    """
    Detecta encoding + separador pelo cabeçalho. ATENÇÃO: os arquivos da Anatel variam —
    o detalhado é UTF-8 com BOM; outros podem ser latin-1. Como latin-1 NUNCA falha ao
    decodificar, não basta "abriu": escolhemos a combinação cujo cabeçalho realmente
    expõe as colunas-chave (Ano e Mês). Retorna (enc, sep, mapa).
    """
    tentativas = [("utf-8-sig", ";"), ("latin-1", ";"),
                  ("utf-8-sig", ","), ("latin-1", ",")]
    reserva = None
    for enc, sep in tentativas:
        try:
            head = pd.read_csv(caminho, sep=sep, encoding=enc, nrows=0)
        except (UnicodeDecodeError, pd.errors.ParserError, ValueError):
            continue
        if len(head.columns) <= 1:
            continue
        mapa = _mapear_colunas(list(head.columns))
        if mapa["ano"] and mapa["mes"]:
            log_verbose(f"Encoding={enc}, sep='{sep}'. Colunas: {list(head.columns)}")
            return enc, sep, mapa
        reserva = reserva or (enc, sep, mapa)
    if reserva:
        return reserva
    raise RuntimeError("Não consegui interpretar o cabeçalho do CSV (encoding/sep).")


def _mapear_colunas(colunas):
    """Casa as colunas reais com os papéis que precisamos (acento/caixa normalizados).
    Para tecnologia, PREFERE 'Tecnologia Geração' (já vem como 2G/3G/4G/5G)."""
    mapa = {"ano": None, "mes": None, "qtd": None, "operadora": None,
            "tecnologia": None, "produto": None}
    norms = {col: normaliza(col) for col in colunas}
    for col, n in norms.items():
        if mapa["ano"] is None and n == "ano":
            mapa["ano"] = col
        elif mapa["mes"] is None and n == "mes":
            mapa["mes"] = col
        elif mapa["qtd"] is None and n in ("acessos", "quantidade", "qtde", "qtd_acessos"):
            mapa["qtd"] = col
        elif mapa["operadora"] is None and n in ("grupo economico", "empresa",
                                                 "prestadora", "nome prestadora"):
            mapa["operadora"] = col
        elif mapa["produto"] is None and n in ("tipo de produto", "tipo produto"):
            mapa["produto"] = col
    # Tecnologia: prioriza a coluna de Geração; senão, a "Tecnologia" crua.
    for alvo in ("tecnologia geracao", "tecnologia"):
        for col, n in norms.items():
            if n == alvo:
                mapa["tecnologia"] = col
                break
        if mapa["tecnologia"]:
            break
    return mapa


def competencia_mais_recente(caminhos, enc, sep, mapa):
    """
    Lê o(s) CSV(s) em UM passe (chunks), agregando por competência dentro de cada chunk
    p/ memória baixa. Identifica a competência (Ano/Mês) mais recente entre todos os
    arquivos e devolve só as linhas agregadas dela.
    Retorna (comp_str "AAAA-MM", DataFrame_agregado_da_competência).
    """
    col_ano, col_mes = mapa["ano"], mapa["mes"]
    if not col_ano or not col_mes:
        raise RuntimeError("Não encontrei as colunas 'Ano' e 'Mês'. "
                           f"Colunas mapeadas: {mapa}")
    dims = [col_ano, col_mes]
    for papel in ("operadora", "tecnologia"):
        if mapa[papel]:
            dims.append(mapa[papel])
    col_qtd = mapa["qtd"]
    col_prod = mapa.get("produto")
    # 'produto' entra só para filtrar M2M/ponto (não vira dimensão de agrupamento).
    usecols = dims + ([col_qtd] if col_qtd else []) + ([col_prod] if col_prod else [])

    partes = []
    for caminho in caminhos:
        leitor = pd.read_csv(caminho, sep=sep, encoding=enc, usecols=usecols,
                             chunksize=CHUNKSIZE)
        for chunk in leitor:
            chunk[col_ano] = pd.to_numeric(chunk[col_ano], errors="coerce")
            chunk[col_mes] = pd.to_numeric(chunk[col_mes], errors="coerce")
            chunk = chunk.dropna(subset=[col_ano, col_mes])
            # Exclui M2M / pontos de serviço para bater com o total oficial da Anatel.
            if col_prod and EXCLUIR_TIPO_PRODUTO:
                excl = chunk[col_prod].astype(str).str.strip().str.upper()
                chunk = chunk[~excl.isin(EXCLUIR_TIPO_PRODUTO)]
            if chunk.empty:
                continue
            if col_qtd:
                chunk[col_qtd] = pd.to_numeric(chunk[col_qtd], errors="coerce").fillna(0)
                g = chunk.groupby(dims, dropna=False)[col_qtd].sum().reset_index()
            else:
                g = chunk[dims].drop_duplicates()
            partes.append(g)

    if not partes:
        raise RuntimeError("Não consegui ler nenhuma competência válida no(s) arquivo(s).")
    acc = pd.concat(partes, ignore_index=True)
    if col_qtd:
        acc = acc.groupby(dims, dropna=False)[col_qtd].sum().reset_index()
    max_ano = int(acc[col_ano].max())
    max_mes = int(acc[acc[col_ano] == max_ano][col_mes].max())
    comp_str = f"{max_ano:04d}-{max_mes:02d}"
    log_verbose(f"Competência mais recente: {comp_str}")
    df_comp = acc[(acc[col_ano] == max_ano) & (acc[col_mes] == max_mes)].copy()
    return comp_str, df_comp


# ===========================================================================
# AGREGADOS (total / operadora / tecnologia)
# ===========================================================================

# Mapeamento dos grupos econômicos para as marcas que interessam. ATENÇÃO: na base da
# Anatel a coluna "Grupo Econômico" usa o nome do CONTROLADOR: Telefônica=Vivo,
# Telecom Americas=Claro (América Móvil), Telecom Italia=TIM.
def _classifica_operadora(valor):
    n = normaliza(valor)
    if "vivo" in n or "telefonica" in n:
        return "Vivo"
    if "claro" in n or "america movil" in n or "telecom americas" in n or "embratel" in n:
        return "Claro"
    if "telecom italia" in n or n == "tim" or n.startswith("tim "):
        return "TIM"
    return "Outras"


# Mapeamento de tecnologia -> geração.
def _classifica_tecnologia(valor):
    n = normaliza(valor)
    if any(t in n for t in ("nr", "5g")):
        return "5G"
    if any(t in n for t in ("lte", "4g")):
        return "4G"
    if any(t in n for t in ("wcdma", "hspa", "umts", "3g")):
        return "3G"
    if any(t in n for t in ("gsm", "2g")):
        return "2G"
    return "Outras"


def montar_agregados(df_comp, mapa):
    """Calcula total Brasil + quebra por operadora e por tecnologia (se houver colunas)."""
    col_qtd = mapa["qtd"]
    resultado = {"total": None, "operadora": None, "tecnologia": None}

    if not col_qtd:
        log_verbose("Coluna de quantidade não encontrada; agregados indisponíveis.")
        return resultado

    qtd = pd.to_numeric(df_comp[col_qtd], errors="coerce").fillna(0)
    df = df_comp.assign(_qtd=qtd)
    resultado["total"] = int(df["_qtd"].sum())

    if mapa["operadora"]:
        op = df[mapa["operadora"]].map(_classifica_operadora)
        resultado["operadora"] = df.groupby(op)["_qtd"].sum().astype("int64").to_dict()

    if mapa["tecnologia"]:
        tec = df[mapa["tecnologia"]].map(_classifica_tecnologia)
        resultado["tecnologia"] = df.groupby(tec)["_qtd"].sum().astype("int64").to_dict()

    return resultado


def _fmt(n):
    """Formata inteiro com separador de milhar no padrão brasileiro."""
    return f"{n:,}".replace(",", ".")


def imprimir_agregados(agg):
    """Imprime total Brasil + quebras de forma legível."""
    if agg["total"] is not None:
        print(f"   Total Brasil: {_fmt(agg['total'])} acessos")
    if agg["operadora"]:
        print("   Por operadora:")
        for nome in ("Vivo", "Claro", "TIM", "Outras"):
            if nome in agg["operadora"]:
                print(f"      - {nome}: {_fmt(agg['operadora'][nome])}")
    if agg["tecnologia"]:
        print("   Por tecnologia:")
        for nome in ("2G", "3G", "4G", "5G", "Outras"):
            if nome in agg["tecnologia"]:
                print(f"      - {nome}: {_fmt(agg['tecnologia'][nome])}")
    if agg["total"] is None:
        print("   (Granularidade de quantidade/operadora/tecnologia não disponível "
              "neste recurso.)")


# ===========================================================================
# ALERTA POR E-MAIL
# ===========================================================================

# Template de configuração. O envio só ocorre se "enabled" = true e houver
# credenciais SMTP. No GitHub Actions este arquivo é escrito a partir do secret
# EMAIL_CONFIG_JSON; localmente você pode editar o arquivo gerado.
DEFAULT_EMAIL_CONFIG = {
    "enabled": False,
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_username": "",          # seu e-mail (login SMTP)
    "smtp_password": "",          # senha de APP do Gmail (não a senha normal)
    "from_addr": "",             # opcional; se vazio usa smtp_username
    "to_addrs": ["danxhp@gmail.com"],
    "subject_prefix": "[Anatel SMP]",
}


def load_email_config():
    """
    Lê email_config.json. Se não existir, cria um template e retorna None.
    Retorna o dict só se estiver 'enabled' e com credenciais/destinatários válidos.
    """
    if not os.path.exists(ARQ_EMAIL_CONFIG):
        with open(ARQ_EMAIL_CONFIG, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_EMAIL_CONFIG, f, indent=2, ensure_ascii=False)
        print(f"ℹ️  Criado template {os.path.basename(ARQ_EMAIL_CONFIG)}. "
              "Edite (enabled=true + credenciais SMTP) para receber e-mails.")
        return None
    try:
        with open(ARQ_EMAIL_CONFIG, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"⚠️  Falha ao ler email_config.json: {e}")
        return None
    if not cfg.get("enabled"):
        log_verbose("E-mail desabilitado (enabled=false).")
        return None
    if not cfg.get("smtp_username") or not cfg.get("smtp_password"):
        print("⚠️  email_config.json sem smtp_username/smtp_password.")
        return None
    if not cfg.get("to_addrs"):
        print("⚠️  email_config.json com to_addrs vazio.")
        return None
    return cfg


def _render_email_html(comp_extenso, atualizado, agg):
    """Monta o corpo HTML do alerta com total Brasil + quebras."""
    def linhas(dic, ordem):
        if not dic:
            return ""
        itens = "".join(
            f'<tr><td style="padding:4px 0;color:#2d3748;">{nome}</td>'
            f'<td style="padding:4px 0;text-align:right;font-weight:600;color:#1a202c;">'
            f'{_fmt(dic[nome])}</td></tr>'
            for nome in ordem if nome in dic
        )
        return itens

    total = _fmt(agg["total"]) if agg["total"] is not None else "—"
    op = linhas(agg["operadora"], ("Vivo", "Claro", "TIM", "Outras"))
    tec = linhas(agg["tecnologia"], ("2G", "3G", "4G", "5G", "Outras"))
    bloco_op = (f'<div style="font-size:11px;font-weight:600;color:#718096;'
                f'text-transform:uppercase;letter-spacing:1px;margin:18px 0 6px;">'
                f'Por operadora</div><table width="100%">{op}</table>') if op else ""
    bloco_tec = (f'<div style="font-size:11px;font-weight:600;color:#718096;'
                 f'text-transform:uppercase;letter-spacing:1px;margin:18px 0 6px;">'
                 f'Por tecnologia</div><table width="100%">{tec}</table>') if tec else ""

    return f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f1f3f5;font-family:-apple-system,
BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1a202c;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
style="background:#f1f3f5;padding:24px 12px;"><tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0"
style="background:#fff;border-radius:10px;overflow:hidden;max-width:600px;width:100%;
box-shadow:0 1px 3px rgba(0,0,0,.06),0 4px 16px rgba(0,0,0,.04);">
<tr><td bgcolor="#1e3a8a" style="padding:28px 32px;background-color:#1e3a8a;">
<div style="font-size:11px;font-weight:700;color:#bfdbfe;letter-spacing:1.5px;
text-transform:uppercase;">Anatel · Acessos móveis (SMP)</div>
<div style="margin-top:10px;font-size:22px;font-weight:700;color:#fff;">
Dado novo: {comp_extenso}</div>
<div style="margin-top:8px;font-size:13px;color:#dbeafe;">Atualizado em {atualizado}</div>
</td></tr>
<tr><td style="padding:24px 32px 4px;">
<div style="font-size:11px;font-weight:600;color:#718096;text-transform:uppercase;
letter-spacing:1px;margin-bottom:6px;">Total Brasil</div>
<div style="font-size:26px;font-weight:700;color:#1a202c;">{total} <span
style="font-size:14px;font-weight:500;color:#718096;">acessos</span></div>
{bloco_op}{bloco_tec}
</td></tr>
<tr><td style="padding:18px 32px 24px;">
<a href="{URL_PAINEL}" style="display:inline-block;background:#3182ce;color:#fff;
padding:12px 22px;border-radius:6px;text-decoration:none;font-weight:600;font-size:14px;">
Abrir painel da Anatel</a></td></tr>
<tr><td style="padding:14px 32px;background:#f7fafc;border-top:1px solid #e2e8f0;
font-size:12px;color:#718096;">Enviado pelo monitor Anatel SMP ·
{datetime.now():%d/%m/%Y %H:%M}</td></tr>
</table></td></tr></table></body></html>"""


def _render_email_text(comp_extenso, atualizado, agg):
    """Versão texto puro do alerta (fallback)."""
    linhas = [f"ANATEL — Acessos móveis (SMP)",
              f"Dado novo: {comp_extenso} (atualizado em {atualizado})", ""]
    if agg["total"] is not None:
        linhas.append(f"Total Brasil: {_fmt(agg['total'])} acessos")
    if agg["operadora"]:
        linhas.append("Por operadora:")
        linhas += [f"  - {n}: {_fmt(agg['operadora'][n])}"
                   for n in ("Vivo", "Claro", "TIM", "Outras") if n in agg["operadora"]]
    if agg["tecnologia"]:
        linhas.append("Por tecnologia:")
        linhas += [f"  - {n}: {_fmt(agg['tecnologia'][n])}"
                   for n in ("2G", "3G", "4G", "5G", "Outras") if n in agg["tecnologia"]]
    linhas += ["", f"Painel: {URL_PAINEL}"]
    return "\n".join(linhas)


def enviar_email(cfg, comp_extenso, atualizado, agg):
    """Envia o alerta SMTP (porta 465=SSL, demais=STARTTLS). Retorna True/False."""
    msg = EmailMessage()
    prefixo = cfg.get("subject_prefix", "[Anatel SMP]")
    msg["Subject"] = f"{prefixo} Dado novo: {comp_extenso}"
    msg["From"] = cfg.get("from_addr") or cfg["smtp_username"]
    destinos = cfg["to_addrs"]
    if isinstance(destinos, str):
        destinos = [destinos]
    msg["To"] = ", ".join(destinos)
    msg.set_content(_render_email_text(comp_extenso, atualizado, agg))
    msg.add_alternative(_render_email_html(comp_extenso, atualizado, agg), subtype="html")

    try:
        host = cfg["smtp_host"]
        port = int(cfg.get("smtp_port", 587))
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=30) as s:
                s.login(cfg["smtp_username"], cfg["smtp_password"])
                s.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=30) as s:
                s.starttls()
                s.login(cfg["smtp_username"], cfg["smtp_password"])
                s.send_message(msg)
        print(f"📧 E-mail enviado para {msg['To']}: {msg['Subject']}")
        return True
    except (smtplib.SMTPException, OSError) as e:
        print(f"⚠️  Falha ao enviar e-mail: {e}")
        return False


# ===========================================================================
# FLUXO PRINCIPAL
# ===========================================================================

def main():
    global VERBOSE
    parser = argparse.ArgumentParser(
        description="Monitora a publicação da competência SMP (acessos móveis) da Anatel."
    )
    parser.add_argument("--verbose", action="store_true",
                        help="Mostra detalhes e força a saída completa mesmo sem novidade.")
    parser.add_argument("--force", action="store_true",
                        help="Ignora o gate de data e baixa/relê o arquivo mesmo sem mudança.")
    args = parser.parse_args()
    VERBOSE = args.verbose

    print(f"📡 Monitor Anatel SMP — alvo: {COMPETENCIA_ALVO} "
          f"(checagem em {datetime.now():%Y-%m-%d %H:%M})")

    estado = carregar_estado()
    comp_anterior = estado.get("ultima_competencia")
    data_arquivo_anterior = estado.get("ultima_atualizacao_arquivo")

    # 1) Descobrir recursos do dataset.
    recursos = descobrir_recursos()
    if not recursos:
        msg = ("Não consegui acessar os recursos do dataset por nenhuma camada "
               "(CKAN/API/HTML). A API do dados.gov.br pode exigir chave "
               "(defina CHAVE_DADOS_GOV) ou estar fora do ar. "
               f"Confira manualmente o painel: {URL_PAINEL}")
        print(f"❌ {msg}")
        gravar_log(f"ERRO: {msg}")
        sys.exit(2)

    # 2) Escolher e baixar o recurso relevante.
    recurso = escolher_recurso(recursos)
    if not recurso:
        msg = "Nenhum recurso CSV/ZIP utilizável foi encontrado no dataset."
        print(f"❌ {msg}")
        gravar_log(f"ERRO: {msg}")
        sys.exit(2)

    data_arquivo = recurso.get("atualizado_em")
    if data_arquivo:
        print(f"📦 Recurso: {recurso['titulo']} (arquivo atualizado em {data_arquivo})")
    else:
        print(f"📦 Recurso: {recurso['titulo']}")

    # 2b) GATE DE DATA: o arquivo é o consolidado de ~3 GB. Em vez de baixá-lo a cada
    #     execução, comparamos a data de atualização do arquivo (vinda da API, barata)
    #     com a da última checagem. Se não mudou, não há competência nova — saímos sem
    #     baixar nada. O download pesado só acontece quando o arquivo realmente muda
    #     (~1x/mês). Use --force para ignorar o gate.
    if (not args.force and data_arquivo and comp_anterior is not None
            and data_arquivo == data_arquivo_anterior):
        print(f"   (sem novidade — arquivo inalterado desde {data_arquivo}; "
              f"competência conhecida: {comp_anterior}. Pulando download de ~3 GB.)")
        estado["ultima_checagem"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        salvar_estado(estado)
        gravar_log(f"SEM_MUDANCA arquivo={data_arquivo} competencia={comp_anterior}")
        return

    try:
        caminhos = baixar(recurso)
    except (requests.exceptions.RequestException, OSError, RuntimeError,
            zipfile.BadZipFile) as e:
        msg = f"Falha ao baixar/extrair o recurso: {e}. Confira o painel: {URL_PAINEL}"
        print(f"❌ {msg}")
        gravar_log(f"ERRO: {msg}")
        sys.exit(2)

    # 3) Ler o(s) CSV(s) e achar a competência mais recente.
    try:
        enc, sep, mapa = detectar_formato(caminhos[0])
        comp_atual, df_comp = competencia_mais_recente(caminhos, enc, sep, mapa)
    except (RuntimeError, pd.errors.ParserError, ValueError) as e:
        msg = f"Falha ao interpretar o CSV: {e}"
        print(f"❌ {msg}")
        gravar_log(f"ERRO: {msg}")
        sys.exit(2)

    # 4) Comparar com o alvo e montar a saída.
    houve_novidade = (comp_atual != comp_anterior)
    primeira_execucao = (comp_anterior is None)
    atingiu_alvo = (comp_atual >= COMPETENCIA_ALVO)
    agg = montar_agregados(df_comp, mapa)  # calculado sempre (usado na tela e no e-mail)

    # Mês por extenso para mensagem amigável.
    meses_pt = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    ano_a, mes_a = COMPETENCIA_ALVO.split("-")
    ano_c, mes_c = comp_atual.split("-")
    alvo_extenso = f"{meses_pt[int(mes_a)]}/{ano_a}"
    atual_extenso = f"{meses_pt[int(mes_c)]}/{ano_c}"
    atualizado = recurso.get("atualizado_em") or "data não informada pela API"

    if atingiu_alvo:
        print(f"\n✅ DADO NOVO DISPONÍVEL: {alvo_extenso} já está na base "
              f"(atualizado em {atualizado})")
        imprimir_agregados(agg)
        resultado_log = f"DISPONIVEL competencia={comp_atual} (alvo={COMPETENCIA_ALVO})"
    else:
        print(f"\n⏳ AINDA NÃO: a competência mais recente é {atual_extenso}. "
              f"{alvo_extenso.capitalize()} ainda não publicado.")
        # Aviso de possível defasagem API x painel (não afirmar que "não saiu").
        print(f"   ℹ️  A API/dados abertos pode estar defasada vs. o painel. "
              f"Se quiser confirmar, veja: {URL_PAINEL}")
        resultado_log = f"PENDENTE competencia={comp_atual} (alvo={COMPETENCIA_ALVO})"

    # 5) Alerta por e-mail: dispara só quando surge uma competência NOVA (mudou desde
    #    a última checagem). Na primeira execução apenas registra o estado, sem e-mail
    #    (evita disparar no "marco zero"). Idempotente: não reenvia se nada mudou.
    if houve_novidade and not primeira_execucao:
        cfg = load_email_config()
        if cfg:
            enviar_email(cfg, atual_extenso, atualizado, agg)
    elif primeira_execucao:
        load_email_config()  # garante criação do template já na 1ª rodada
        log_verbose("Primeira execução: estado inicial registrado, sem e-mail.")

    # 6) Log + estado. Silencioso/idempotente quando NÃO houve mudança.
    gravar_log(resultado_log)
    salvar_estado({
        "ultima_competencia": comp_atual,
        "ultima_atualizacao_arquivo": data_arquivo,
        "ultima_checagem": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "alvo": COMPETENCIA_ALVO,
    })

    if not houve_novidade and not VERBOSE:
        # Sem novidade desde a última execução: mantém discreto.
        print("   (sem novidade desde a última checagem)")

    log_verbose(f"Estado salvo. Competência anterior: {comp_anterior} -> atual: {comp_atual}")


if __name__ == "__main__":
    main()
