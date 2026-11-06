"""
Verificador de links quebrados

Rodar no computador:     streamlit run app.py
Publicar para a equipe:  Streamlit Community Cloud (app.py + requirements.txt em um repositório do GitHub)
"""
import io
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.parse import urldefrag, urljoin, urlparse

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; VerificadorDeLinks/1.0)"}
IGNORAR = ("#", "mailto:", "tel:", "javascript:", "data:", "sms:", "whatsapp:")
NAO_HTML = (".pdf", ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".zip", ".rar", ".doc",
            ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".mp3", ".mp4", ".css", ".js")
BLOQUEIO = {401, 403, 429, 999}  # destinos que costumam bloquear robôs: conferir no navegador
SECOES = {"header": "Cabeçalho", "nav": "Menu", "main": "Conteúdo principal",
          "aside": "Barra lateral", "footer": "Rodapé"}
ERROS = {"ConnectTimeout": "Timeout", "ReadTimeout": "Timeout",
         "ConnectionError": "Falha de conexão ou DNS", "SSLError": "Erro de certificado SSL",
         "TooManyRedirects": "Redirecionamentos demais", "InvalidURL": "URL inválida"}
COLUNAS = ["Página", "Título da página", "Link", "Status", "Situação", "URL final", "Texto âncora",
           "Seção da página", "Título mais próximo", "Linha no HTML", "Trecho HTML"]
CARACTERES_ILEGAIS = r"[\x00-\x08\x0b\x0c\x0e-\x1f]"  # o Excel recusa estes caracteres


# ----------------------------------------------------------------- motor de análise

def dominio_de(url):
    return urlparse(url).netloc.lower().removeprefix("www.")


def normalizar(url):
    return re.sub(r"^https?://(www\.)?", "", url.strip().lower()).rstrip("/")


def casa(link, procurado, exato):
    a, b = normalizar(link), normalizar(procurado)
    return a == b if exato else b in a


def situacao(status):
    if isinstance(status, int) and status < 400:
        return "OK"
    return "VERIFICAR" if status in BLOQUEIO else "QUEBRADO"


def baixar_pagina(url, timeout):
    """Baixa uma página. Devolve (status, html ou None, url final)."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout, stream=True)
        html = None
        if r.status_code < 400 and "html" in r.headers.get("Content-Type", "").lower():
            if r.encoding in (None, "ISO-8859-1"):
                r.encoding = r.apparent_encoding
            html = r.text
        r.close()
        return r.status_code, html, r.url
    except Exception as e:
        return ERROS.get(type(e).__name__, type(e).__name__), None, url


def checar_link(url, timeout):
    """Confere um link com HEAD e, se falhar, com GET. Devolve (status, url final)."""
    erro = "ERRO"
    for metodo in ("HEAD", "GET"):
        try:
            r = requests.request(metodo, url, headers=HEADERS, timeout=timeout,
                                 allow_redirects=True, stream=True)
            r.close()
            if r.status_code < 400 or metodo == "GET":
                return r.status_code, r.url
        except Exception as e:
            erro = ERROS.get(type(e).__name__, type(e).__name__)
    return erro, url


def extrair_links(html, base):
    """Lista os links <a> da página e onde cada um está (seção, título mais próximo, linha do HTML)."""
    soup = BeautifulSoup(html, "html.parser")
    titulo = soup.title.get_text(strip=True) if soup.title else ""
    achados = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.lower().startswith(IGNORAR):
            continue
        try:
            destino = urldefrag(urljoin(base, href))[0]
        except ValueError:
            continue
        if not destino.lower().startswith(("http://", "https://")):
            continue
        secao = a.find_parent(list(SECOES))
        cabecalho = a.find_previous(["h1", "h2", "h3", "h4", "h5", "h6"])
        texto = a.get_text(" ", strip=True)
        if not texto and a.img:
            texto = f"[imagem] {a.img.get('alt', '')}".strip()
        achados.append({
            "Link": destino,
            "Texto âncora": texto,
            "Seção da página": SECOES[secao.name] if secao else "Corpo da página",
            "Título mais próximo": cabecalho.get_text(" ", strip=True)[:100] if cabecalho else "",
            "Linha no HTML": a.sourceline,
            "Trecho HTML": str(a)[:250],
        })
    return titulo, achados


def analisar(inicio, rastrear, limite, caminho, timeout, workers, procurado, exato, progresso):
    """Lê as páginas (rastreando o site, se pedido), confere cada link e devolve os resultados."""
    paginas, linhas, conhecidos = [], [], {}
    fila, vistas = list(inicio), {u.rstrip("/") for u in inicio}
    dominio = dominio_de(inicio[0])

    with ThreadPoolExecutor(workers) as ex:
        # 1) lê as páginas, em levas
        while fila and len(paginas) < limite:
            n = limite - len(paginas)
            leva, fila = fila[:n], fila[n:]
            resultados = ex.map(lambda u: baixar_pagina(u, timeout), leva)
            for url, (status, html, final) in zip(leva, resultados):
                titulo, achados = extrair_links(html, final) if html else ("", [])
                conhecidos[url] = (status, final)
                paginas.append({"Página": url, "Título": titulo, "Status": status,
                                "Situação": situacao(status), "Links encontrados": len(achados)})
                for a in achados:
                    linhas.append({"Página": url, "Título da página": titulo, **a})
                    link = a["Link"]
                    if (rastrear and link.rstrip("/") not in vistas and dominio_de(link) == dominio
                            and (not caminho or caminho in link)
                            and not link.lower().split("?")[0].endswith(NAO_HTML)):
                        vistas.add(link.rstrip("/"))
                        fila.append(link)
            progresso(f"Páginas lidas: {len(paginas)} (na fila: {len(fila)})",
                      0.5 * len(paginas) / (len(paginas) + len(fila)))

        # 2) confere os links únicos que ainda não têm status
        checados = dict(conhecidos)
        unicos = sorted({l["Link"] for l in linhas} - set(conhecidos))
        futuros = {ex.submit(checar_link, u, timeout): u for u in unicos}
        for i, f in enumerate(as_completed(futuros), 1):
            checados[futuros[f]] = f.result()
            if i % 5 == 0 or i == len(unicos):
                progresso(f"Conferindo links: {i} de {len(unicos)}", 0.5 + 0.5 * i / len(unicos))

    # 3) monta as tabelas
    for l in linhas:
        status, final = checados[l["Link"]]
        l["Status"], l["Situação"] = status, situacao(status)
        l["URL final"] = "" if final.rstrip("/") == l["Link"].rstrip("/") else final

    links = pd.DataFrame(linhas, columns=COLUNAS)
    paginas_df = pd.DataFrame(paginas, columns=["Página", "Título", "Status", "Situação", "Links encontrados"])
    quebrados = links[links["Situação"] == "QUEBRADO"]
    verificar = links[links["Situação"] == "VERIFICAR"]
    buscou = bool(procurado.strip())
    if buscou:
        achados = links[links["Link"].map(lambda u: casa(u, procurado, exato)).astype(bool)]
    else:
        achados = links.iloc[0:0]
    resumo = pd.DataFrame([
        ("Data da análise", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Páginas analisadas", len(paginas_df)),
        ("Páginas que não abriram", int((paginas_df["Situação"] != "OK").sum())),
        ("Links encontrados (com repetição)", len(links)),
        ("Links únicos conferidos", links["Link"].nunique()),
        ("Links quebrados (ocorrências)", len(quebrados)),
        ("Links para conferir manualmente (ocorrências)", len(verificar)),
        ("Link procurado", procurado.strip() if buscou else "—"),
        ("Ocorrências do link procurado", len(achados) if buscou else "—"),
    ], columns=["Indicador", "Valor"])
    return {"resumo": resumo, "paginas": paginas_df, "links": links, "quebrados": quebrados,
            "verificar": verificar, "achados": achados, "buscou": buscou}


def gerar_excel(res):
    abas = {"Resumo": res["resumo"], "Links quebrados": res["quebrados"],
            "Conferir manualmente": res["verificar"]}
    if res["buscou"]:
        abas["Link procurado"] = res["achados"]
    abas["Todos os links"] = res["links"]
    abas["Páginas"] = res["paginas"]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, df in abas.items():
            df.replace(CARACTERES_ILEGAIS, "", regex=True).to_excel(w, sheet_name=nome, index=False)
            ws = w.sheets[nome]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="1F4E78")
            for col in ws.columns:
                maior = max(len(str(c.value or "")) for c in col[:300])
                ws.column_dimensions[col[0].column_letter].width = min(max(maior + 2, 10), 60)
    return buf.getvalue()


def ler_urls(arquivo, texto):
    """Junta as URLs coladas e as do arquivo (.txt, .csv ou .xlsx, em qualquer coluna)."""
    candidatas = texto.split()
    if arquivo is not None:
        nome = arquivo.name.lower()
        if nome.endswith(".txt"):
            candidatas += arquivo.getvalue().decode("utf-8-sig", "ignore").split()
        else:
            if nome.endswith(".xlsx"):
                tabelas = pd.read_excel(arquivo, header=None, dtype=str, sheet_name=None).values()
            else:
                tabelas = [pd.read_csv(arquivo, header=None, dtype=str, sep=None, engine="python",
                                       encoding_errors="ignore")]
            for t in tabelas:
                candidatas += [str(v) for v in t.to_numpy().ravel()]
    urls = [u.strip() for u in candidatas if u.strip().lower().startswith(("http://", "https://"))]
    return list(dict.fromkeys(urls))


# ----------------------------------------------------------------- interface

def tabela(df):
    if "Status" in df.columns:
        df = df.astype({"Status": str})
    st.dataframe(df, hide_index=True)


def mostrar(res):
    st.divider()
    m = st.columns(4)
    m[0].metric("Páginas analisadas", len(res["paginas"]))
    m[1].metric("Links encontrados", len(res["links"]))
    m[2].metric("Links quebrados", len(res["quebrados"]))
    m[3].metric("Ocorrências do link procurado", len(res["achados"]) if res["buscou"] else "—")
    st.download_button("Baixar relatório em Excel", res["excel"], file_name=res["arquivo"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    falhas = int((res["paginas"]["Situação"] != "OK").sum())
    if falhas:
        st.warning(f"{falhas} página(s) não puderam ser lidas. Veja o motivo na aba Páginas.")

    t1, t2, t3, t4, t5 = st.tabs(["Quebrados", "Conferir manualmente", "Link procurado",
                                  "Todos os links", "Páginas"])
    with t1:
        if res["quebrados"].empty:
            st.success("Nenhum link quebrado nas páginas analisadas.")
        else:
            tabela(res["quebrados"])
    with t2:
        st.caption("Estes destinos responderam 401, 403 ou 429, o que costuma ser bloqueio a robôs. "
                   "Abra no navegador para confirmar.")
        tabela(res["verificar"])
    with t3:
        if not res["buscou"]:
            st.info("Digite um link no campo “Procurar um link específico” e analise de novo.")
        elif res["achados"].empty:
            st.warning("O link procurado não aparece nas páginas analisadas.")
        else:
            tabela(res["achados"])
    with t4:
        tabela(res["links"])
    with t5:
        tabela(res["paginas"])


def interface():
    st.set_page_config(page_title="Verificador de links", page_icon="🔗", layout="wide")
    st.title("Verificador de links quebrados")
    st.caption("Analise um site ou uma lista de páginas e baixe o relatório em Excel com cada link "
               "quebrado e o lugar da página onde ele está.")

    with st.sidebar:
        st.header("Configurações")
        limite = st.slider("Máximo de páginas ao rastrear um site", 10, 2000, 200, step=10)
        caminho = st.text_input("Rastrear só URLs que contenham (opcional)", placeholder="/es/",
                                help="Ex.: /es/ mantém o rastreamento dentro dessa seção do site.")
        timeout = st.slider("Tempo máximo por requisição (segundos)", 5, 60, 15)
        workers = st.slider("Requisições simultâneas", 1, 20, 8,
                            help="Valores altos aceleram a análise, mas o site pode bloquear o acesso.")

    c1, c2 = st.columns([4, 1], vertical_alignment="bottom")
    procurado = c1.text_input("Procurar um link específico (opcional)", key="procurado",
                              placeholder="https://es.loja.sebrae.com.br/agropecuaria")
    exato = c2.checkbox("Correspondência exata", key="exato",
                        help="Desmarcado, encontra qualquer link que contenha o texto digitado.")

    def rodar(urls, rastrear, max_paginas):
        barra = st.progress(0.0, text="Iniciando…")
        try:
            res = analisar(urls, rastrear, max_paginas, caminho.strip(), timeout, workers,
                           procurado, exato, lambda texto, p: barra.progress(min(p, 1.0), text=texto))
        finally:
            barra.empty()
        res["arquivo"] = f"relatorio_links_{datetime.now():%Y%m%d_%H%M}.xlsx"
        res["excel"] = gerar_excel(res)
        st.session_state["res"] = res

    aba_url, aba_lote = st.tabs(["Analisar uma URL", "Lote de páginas"])

    with aba_url:
        url = st.text_input("Endereço do site ou da página", key="url_unica",
                            placeholder="https://sebrae.com.br/es/empreendedores")
        rastrear = st.checkbox("Rastrear o site inteiro (segue os links internos)", value=True,
                               key="rastrear", help="Respeita o máximo de páginas da barra lateral.")
        if st.button("Analisar", type="primary", key="btn_url"):
            u = url.strip()
            if not u:
                st.warning("Digite o endereço que você quer analisar.")
            else:
                if not u.lower().startswith(("http://", "https://")):
                    u = "https://" + u
                rodar([urldefrag(u)[0]], rastrear, limite)

    with aba_lote:
        arquivo = st.file_uploader("Arquivo com as páginas (.xlsx, .csv ou .txt)",
                                   type=["xlsx", "csv", "txt"], key="arquivo",
                                   help="Lê todos os endereços que começam com http:// ou https://, em qualquer coluna.")
        colado = st.text_area("Ou cole os endereços aqui, um por linha", height=150, key="colado")
        if st.button("Analisar todas as páginas", type="primary", key="btn_lote"):
            urls = ler_urls(arquivo, colado)
            if not urls:
                st.warning("Nenhum endereço válido encontrado. Use endereços completos, "
                           "começando com http:// ou https://.")
            else:
                rodar(urls, False, len(urls))

    if st.session_state.get("res"):
        mostrar(st.session_state["res"])


if __name__ == "__main__":
    interface()
