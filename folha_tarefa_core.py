# -*- coding: utf-8 -*-
"""
================================================================================
 NUCLEO COMPARTILHADO - GERADOR DE FOLHA TAREFA
 Consorcio Possebon / CPL - UTAA/TR - CPC RNEST
================================================================================
Toda a logica de leitura das planilhas, formatacao e montagem do HTML/PDF fica
aqui, para ser usada tanto pelo script de linha de comando (gerar_folha_tarefa.py)
quanto pelo aplicativo web (app_web.py). Assim os dois ficam sempre identicos.
"""

import os
import glob
import base64

import pandas as pd
from weasyprint import HTML

# ==============================================================================
# CAMINHOS DO PROJETO
# ==============================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
LOGO_PATH = os.path.join(BASE_DIR, "logo", "logo1.png")
CONTADOR_PATH = os.path.join(BASE_DIR, "controle_numero_folha.txt")

# Prefixos usados para localizar automaticamente o arquivo mais recente na pasta input/
PREFIXO_SPOOLS = "U-54-ControlTUB_-_Situacao_Geral_de_Spools_-_CPC_RNEST"
PREFIXO_JUNTAS = "U-54-ControlTUB_-_Situacao_Geral_de_Juntas_-_CPC_RNEST"

# Nomes das abas dentro dos arquivos originais
ABA_SPOOLS = "Mapa de Spools"
ABA_JUNTAS = "MAPA DE JUNTAS"

# Linhas de cabecalho reais dentro das planilhas (0-indexado, layout ControlTUB)
HEADER_ROW_SPOOLS = 5   # cabecalho esta na linha 6 do Excel
HEADER_ROW_JUNTAS = 6   # cabecalho esta na linha 7 do Excel

os.makedirs(OUTPUT_DIR, exist_ok=True)


# ==============================================================================
# LOCALIZACAO E LEITURA DAS PLANILHAS
# ==============================================================================

def encontrar_arquivo_mais_recente(pasta, prefixo):
    """Procura na pasta 'input' o arquivo .xlsx mais recente que comece com 'prefixo'."""
    candidatos = [
        f for f in glob.glob(os.path.join(pasta, "*.xlsx"))
        if os.path.basename(f).startswith(prefixo)
    ]
    if not candidatos:
        raise FileNotFoundError(
            f"\nNenhum arquivo encontrado em '{pasta}' comecando com:\n"
            f"  {prefixo}\n"
            f"Coloque o arquivo exportado do ControlTUB (mesmo com data/hora no nome) "
            f"dentro da pasta 'input'."
        )
    candidatos.sort(key=os.path.getmtime, reverse=True)
    return candidatos[0]


def carregar_planilhas(caminho_spools, caminho_juntas):
    """Le os dois arquivos xlsx (caminho ou objeto tipo arquivo) e devolve os DataFrames prontos."""
    df_spools = pd.read_excel(
        caminho_spools, sheet_name=ABA_SPOOLS, header=HEADER_ROW_SPOOLS
    ).fillna("-")
    df_juntas = pd.read_excel(
        caminho_juntas, sheet_name=ABA_JUNTAS, header=HEADER_ROW_JUNTAS
    ).fillna("-")
    return df_spools, df_juntas


def filtrar_spools(df_spools, isometrico_alvo="", linha_alvo="", spool_alvo="", limite=None):
    """Aplica os filtros de Isometrico / Tag da Linha / numero do Spool."""
    filtrado = df_spools.copy()

    if isometrico_alvo and isometrico_alvo.strip():
        filtrado = filtrado[
            filtrado["Isométrico"].astype(str).str.contains(isometrico_alvo.strip(), case=False, na=False)
        ]

    if linha_alvo and linha_alvo.strip():
        filtrado = filtrado[
            filtrado["Linha"].astype(str).str.contains(linha_alvo.strip(), case=False, na=False)
        ]

    if spool_alvo and spool_alvo.strip():
        filtrado = filtrado[
            filtrado["Spool"].astype(str).str.lstrip("0") == str(spool_alvo).strip().lstrip("0")
        ]

    if limite:
        filtrado = filtrado.head(int(limite))

    return filtrado


# ==============================================================================
# NUMERACAO DA FOLHA TAREFA (contador persistente)
# ==============================================================================

def obter_numero_folha(avancar=True):
    """
    Le o contador persistente e devolve o proximo numero (F0001, F0002...).
    Se avancar=False, so consulta o proximo numero sem gravar (util para
    pre-visualizacao no app web antes de confirmar a geracao).
    """
    numero_atual = 0
    if os.path.exists(CONTADOR_PATH):
        try:
            with open(CONTADOR_PATH, "r", encoding="utf-8") as f:
                numero_atual = int(f.read().strip())
        except (ValueError, OSError):
            numero_atual = 0

    proximo = numero_atual + 1

    if avancar:
        with open(CONTADOR_PATH, "w", encoding="utf-8") as f:
            f.write(str(proximo))

    return proximo


# ==============================================================================
# STATUS MANUAL (marcar "ja fiz" direto na tela, sem esperar o ControlTUB)
# ==============================================================================
# Guarda, por spool, as datas informadas manualmente na tela (ex: "ja fiz o
# corte hoje"). So e usado como PREENCHIMENTO PROVISORIO: se a planilha oficial
# do ControlTUB ja tiver uma data naquele campo, ela sempre prevalece.

STATUS_MANUAL_PATH = os.path.join(BASE_DIR, "status_manual.json")

# (rotulo exibido, nome da coluna na planilha de Spools)
CAMPOS_CHECKLIST_SPOOL = [
    ("Corte", "Data_Corte"),
    ("VA (Ajuste)", "Data_VA_Fab"),
    ("Solda", "Data_Soldagem_Fab"),
    ("Exame Visual", "Data_VS_Fab"),
    ("E.N.D", "Data_End_Fab"),
    ("Dimensional de Fábrica", "Data_Dim_Fab"),
    ("Pintura de Fundo", "Data_Pint_Fundo"),
    ("Pintura Interm.", "Data_Pint_Intermediária"),
]


def chave_spool(unidade, isometrico, spool):
    return f"{unidade}|{isometrico}|{spool}"


def carregar_status_manual():
    if not os.path.exists(STATUS_MANUAL_PATH):
        return {}
    try:
        import json
        with open(STATUS_MANUAL_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (ValueError, OSError):
        return {}


def _salvar_status_manual_completo(dados):
    import json
    with open(STATUS_MANUAL_PATH, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


def definir_status_manual(chave, campo, valor):
    """valor = string de data (ex '30/07/2026') para marcar feito, ou None para desmarcar."""
    dados = carregar_status_manual()
    dados.setdefault(chave, {})
    if valor:
        dados[chave][campo] = valor
    else:
        dados[chave].pop(campo, None)
        if not dados[chave]:
            dados.pop(chave, None)
    _salvar_status_manual_completo(dados)


def obter_spool_com_overrides(spool_row, overrides=None):
    """Devolve um dict do spool onde campos em branco na planilha sao preenchidos
    com o status manual salvo na tela (a planilha oficial sempre tem prioridade)."""
    if overrides is None:
        overrides = carregar_status_manual()
    spool_dict = spool_row.to_dict() if hasattr(spool_row, "to_dict") else dict(spool_row)
    chave = chave_spool(
        spool_dict.get("Unidade", "-"), spool_dict.get("Isométrico", "-"), spool_dict.get("Spool", "-")
    )
    status_spool = overrides.get(chave, {})
    for _label, campo in CAMPOS_CHECKLIST_SPOOL:
        valor_atual = str(spool_dict.get(campo, "-")).strip()
        if valor_atual in ("-", "nan", "None", "") and status_spool.get(campo):
            spool_dict[campo] = status_spool[campo]
    return spool_dict


def campos_manuais_do_spool(spool_row, overrides=None):
    """Devolve o conjunto de colunas do checklist do spool cujo valor exibido
    veio do preenchimento manual na tela (e nao da planilha oficial do
    ControlTUB) - usado para marcar essas datas em vermelho com '(IM)'."""
    if overrides is None:
        overrides = carregar_status_manual()
    spool_dict = spool_row.to_dict() if hasattr(spool_row, "to_dict") else dict(spool_row)
    chave = chave_spool(
        spool_dict.get("Unidade", "-"), spool_dict.get("Isométrico", "-"), spool_dict.get("Spool", "-")
    )
    status_spool = overrides.get(chave, {})
    manuais = set()
    for _label, campo in CAMPOS_CHECKLIST_SPOOL:
        valor_atual = str(spool_dict.get(campo, "-")).strip()
        if valor_atual in ("-", "nan", "None", "") and status_spool.get(campo):
            manuais.add(campo)
    return manuais


def aplicar_overrides_dataframe(df, overrides=None):
    """Aplica o status manual em todas as linhas de um DataFrame de spools
    (usado antes de gerar o PDF, para o manual entrar tambem no documento)."""
    if overrides is None:
        overrides = carregar_status_manual()
    if not overrides:
        return df
    df = df.copy()
    for idx, row in df.iterrows():
        merged = obter_spool_com_overrides(row, overrides)
        for _label, campo in CAMPOS_CHECKLIST_SPOOL:
            if campo in df.columns:
                df.at[idx, campo] = merged.get(campo, df.at[idx, campo])
    return df


def montar_resumo_programacao(df_spools):
    """Resumo por Programa\u00e7\u00e3o de Fabrica\u00e7\u00e3o (coluna Prog_Fab), no estilo do
    painel do ControlTUB: quantidade de spools, peso total e % concluido de
    cada etapa (considerando tambem o status marcado manualmente na tela)."""
    df = aplicar_overrides_dataframe(df_spools)

    def _percentual_feito(serie):
        total = len(serie)
        if total == 0:
            return 0.0
        feitos = serie.astype(str).str.strip().apply(lambda v: v not in ("-", "nan", "None", "")).sum()
        return round(100 * feitos / total, 1)

    def _peso_numerico(serie):
        return pd.to_numeric(serie, errors="coerce").fillna(0).sum()

    linhas = []
    programacoes = df["Prog_Fab"].astype(str).str.strip()
    programacoes = programacoes.where(~programacoes.isin(["", "-", "nan", "None"]), "Sem Programa\u00e7\u00e3o")
    for prog in sorted(programacoes.unique()):
        grupo = df[programacoes == prog]
        linhas.append({
            "Programa\u00e7\u00e3o": prog,
            "Qtd. Spools": len(grupo),
            "Peso Total (kg)": round(_peso_numerico(grupo.get("Peso", pd.Series(dtype=float))), 1),
            "% Solda": _percentual_feito(grupo.get("Data_Soldagem_Fab", pd.Series(dtype=str))),
            "% END": _percentual_feito(grupo.get("Data_End_Fab", pd.Series(dtype=str))),
            "% Dimensional": _percentual_feito(grupo.get("Data_Dim_Fab", pd.Series(dtype=str))),
        })

    return pd.DataFrame(linhas).sort_values("Programa\u00e7\u00e3o").reset_index(drop=True)


def peso_total_programacao(df_spools_completo, prog_fab_valor):
    """Soma o peso (coluna Peso) de todos os spools que pertencem a mesma
    Programacao de Fabricacao (Prog_Fab) do spool informado."""
    if df_spools_completo is None or "Prog_Fab" not in df_spools_completo.columns:
        return None
    prog_col = df_spools_completo["Prog_Fab"].astype(str).str.strip()
    valor_str = str(prog_fab_valor).strip()
    if valor_str in ("", "-", "nan", "None"):
        grupo = df_spools_completo[prog_col.isin(["", "-", "nan", "None"])]
    else:
        grupo = df_spools_completo[prog_col == valor_str]
    peso = pd.to_numeric(grupo.get("Peso", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    return round(peso, 1)


def formatar_peso_br(valor):
    """Formata um numero no padrao brasileiro (ex: 216166.5 -> '216.166,5')."""
    if valor is None:
        return "-"
    texto = f"{valor:,.1f}"
    return texto.replace(",", "@").replace(".", ",").replace("@", ".")


# ==============================================================================
# OCORRENCIAS / PENDENCIAS (banco SQLite - motivo, setor, usuario, fotos)
# ==============================================================================
# Cada vez que um item do checklist nao pode ser marcado como feito, o usuario
# pode registrar o motivo. Fica guardado em banco SQLite (mais robusto que
# arquivo texto/JSON para varios usuarios e para anexar fotos) e aparece tanto
# na tela quanto na Folha Tarefa em PDF.

import sqlite3
import datetime as _datetime

DB_PATH = os.path.join(BASE_DIR, "folha_tarefa.db")
EVIDENCIAS_DIR = os.path.join(BASE_DIR, "evidencias")
os.makedirs(EVIDENCIAS_DIR, exist_ok=True)

SETORES = ["Planejamento", "Constru\u00e7\u00e3o e Montagem", "Qualidade", "Pintura"]

MOTIVOS_PADRAO = [
    "Falta de acesso",
    "\u00c1rea bloqueada",
    "Sem suporta\u00e7\u00e3o",
    "Spool sem tag definido",
    "Spool n\u00e3o est\u00e1 soldado",
    "Falta documenta\u00e7\u00e3o de soldagem",
    "Spool desalinhado",
    "Falta acoplamento",
    "Outro (descrever)",
]


def _conectar_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ocorrencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chave TEXT NOT NULL,
            atividade TEXT NOT NULL,
            motivo TEXT NOT NULL,
            setor TEXT,
            usuario TEXT,
            data_hora TEXT NOT NULL,
            fotos TEXT
        )
    """)
    return conn


def adicionar_ocorrencia(chave, atividade, motivo, setor, usuario, fotos=None):
    """Grava uma ocorrencia (motivo de nao realizacao) para um item do checklist."""
    conn = _conectar_db()
    try:
        conn.execute(
            "INSERT INTO ocorrencias (chave, atividade, motivo, setor, usuario, data_hora, fotos) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (
                chave, atividade, motivo, setor or "-", usuario or "-",
                _datetime.datetime.now().strftime("%d/%m/%y %H:%M"),
                ",".join(fotos) if fotos else "",
            ),
        )
        conn.commit()
    finally:
        conn.close()


def ocorrencias_do_spool(chave):
    """Devolve a lista de ocorrencias registradas para um spool, mais recente primeiro."""
    conn = _conectar_db()
    try:
        cur = conn.execute(
            "SELECT atividade, motivo, setor, usuario, data_hora, fotos "
            "FROM ocorrencias WHERE chave = ? ORDER BY id DESC",
            (chave,),
        )
        colunas = ["atividade", "motivo", "setor", "usuario", "data_hora", "fotos"]
        return [dict(zip(colunas, linha)) for linha in cur.fetchall()]
    finally:
        conn.close()


def substituir_ocorrencia_aberta(chave, atividade, motivo, setor, usuario, fotos=None):
    """Registra a pendencia de uma atividade, substituindo qualquer pendencia
    anterior da MESMA atividade nesse spool (evita ficar um registro antigo
    sem foto exibido enquanto um mais novo com foto existe)."""
    limpar_ocorrencia_atividade(chave, atividade)
    adicionar_ocorrencia(chave, atividade, motivo, setor, usuario, fotos=fotos)


def limpar_ocorrencia_atividade(chave, atividade):
    """Remove a pendencia registrada para essa atividade (ex: quando ela e marcada como feita)."""
    conn = _conectar_db()
    try:
        conn.execute(
            "DELETE FROM ocorrencias WHERE chave = ? AND atividade = ?", (chave, atividade)
        )
        conn.commit()
    finally:
        conn.close()



def salvar_fotos_evidencia(chave, arquivos_upload):
    """Salva as fotos enviadas (bytes) em disco e devolve a lista de nomes de arquivo salvos."""
    if not arquivos_upload:
        return []
    pasta_chave = os.path.join(EVIDENCIAS_DIR, chave.replace("|", "_").replace("/", "-"))
    os.makedirs(pasta_chave, exist_ok=True)
    nomes_salvos = []
    carimbo = _datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    for i, arquivo in enumerate(arquivos_upload):
        nome_original = getattr(arquivo, "name", f"foto_{i}.jpg")
        nome_final = f"{carimbo}_{i}_{nome_original}"
        caminho_final = os.path.join(pasta_chave, nome_final)
        with open(caminho_final, "wb") as f:
            f.write(arquivo.getbuffer() if hasattr(arquivo, "getbuffer") else arquivo.read())
        nomes_salvos.append(os.path.join(os.path.basename(pasta_chave), nome_final))
    return nomes_salvos


def montar_bloco_ocorrencias_html(chave, fotos_labels=None):
    """HTML (mesmo visual do documento) com as ocorrencias/pendencias registradas
    para o spool - aparece tanto na tela quanto no PDF, logo apos as juntas.

    fotos_labels: dict opcional {caminho_relativo: "Foto-01"} - quando informado,
    mostra so a referencia ("Ver Foto-01") em vez da miniatura embutida (usado na
    geracao do PDF em lote, onde as fotos ficam reunidas na ultima pagina)."""
    ocorrencias = ocorrencias_do_spool(chave)
    if not ocorrencias:
        return '<div class="sem-ocorrencias">Nenhuma pend&ecirc;ncia registrada para este spool.</div>'

    itens_html = ""
    for oc in ocorrencias:
        fotos_html = ""
        caminhos = [c for c in oc["fotos"].split(",") if c] if oc["fotos"] else []
        if caminhos and fotos_labels is not None:
            rotulos = [fotos_labels.get(c, "?") for c in caminhos]
            fotos_html = f'<div class="ocorrencia-fotos-ref">📷 Ver {", ".join(rotulos)} (anexo)</div>'
        elif caminhos:
            for caminho_relativo in caminhos:
                caminho_absoluto = os.path.join(EVIDENCIAS_DIR, caminho_relativo)
                if os.path.exists(caminho_absoluto):
                    with open(caminho_absoluto, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode("utf-8")
                    ext = caminho_absoluto.rsplit(".", 1)[-1].lower()
                    mime = "jpeg" if ext in ("jpg", "jpeg") else ext
                    fotos_html += f'<img src="data:image/{mime};base64,{b64}" />'
            if fotos_html:
                fotos_html = f'<div class="ocorrencia-fotos">{fotos_html}</div>'
        itens_html += f"""
        <div class="ocorrencia-item">
            <div class="ocorrencia-linha1">
                <span class="ocorrencia-setor">{oc['setor']}</span>
                <span class="ocorrencia-datahora">{oc['data_hora']}</span>
            </div>
            <div><span class="ocorrencia-atividade">{oc['atividade']}:</span>
                 <span class="ocorrencia-motivo">{oc['motivo']}</span></div>
            <div class="ocorrencia-usuario">Usu&aacute;rio: {oc['usuario']}</div>
            {fotos_html}
        </div>
        """

    return f'<div class="ocorrencias-box">{itens_html}</div>'


def montar_anexo_fotos(spools_do_lote, logo_html=None, codigo_base_folha=None):
    """Percorre todos os spools do lote, reune as fotos de todas as ocorrencias
    (na ordem em que aparecem) e monta:
    - fotos_labels: dict {chave: {caminho_relativo: "Foto-01"}} para cada spool
    - html_anexo: uma pagina extra (HTML) com todas as fotos numeradas, com o
      MESMO cabecalho e rodape de assinaturas das demais paginas, para ficar
      sempre na ULTIMA pagina do documento gerado."""
    fotos_labels = {}
    cartoes_html = ""
    contador = 0

    for _, spool in spools_do_lote.iterrows():
        chave = chave_spool(
            spool.get("Unidade", "-"), spool.get("Isométrico", "-"), spool.get("Spool", "-")
        )
        fotos_labels.setdefault(chave, {})
        for oc in ocorrencias_do_spool(chave):
            if not oc["fotos"]:
                continue
            for caminho_relativo in oc["fotos"].split(","):
                if not caminho_relativo or caminho_relativo in fotos_labels[chave]:
                    continue
                caminho_absoluto = os.path.join(EVIDENCIAS_DIR, caminho_relativo)
                if not os.path.exists(caminho_absoluto):
                    continue
                contador += 1
                rotulo = f"Foto-{contador:02d}"
                fotos_labels[chave][caminho_relativo] = rotulo

                with open(caminho_absoluto, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode("utf-8")
                ext = caminho_absoluto.rsplit(".", 1)[-1].lower()
                mime = "jpeg" if ext in ("jpg", "jpeg") else ext

                cartoes_html += f"""
                <div class="anexo-foto-card">
                    <img src="data:image/{mime};base64,{b64}" />
                    <div class="anexo-foto-titulo">{rotulo}</div>
                    <div class="anexo-foto-legenda">
                        Isom&eacute;trico {spool.get('Isométrico', '-')} &middot; Spool {spool.get('Spool', '-')}<br>
                        {oc['atividade']}: {oc['motivo']}<br>
                        Anexado por: {oc['usuario']}
                    </div>
                </div>
                """

    if contador == 0:
        return fotos_labels, ""

    cabecalho_html = ""
    if logo_html is not None:
        cabecalho_html = f"""
        <div class="doc-header">
            <div class="logo-box">{logo_html}</div>
            <div class="header-center">
                <div class="header-title">Folha Tarefa &ndash; Controle de Fabrica&ccedil;&atilde;o e Montagem</div>
                <div class="header-sub">Cons&oacute;rcio Possebon / CPL &middot; UTAA/TR &middot; CPC RNEST</div>
            </div>
            <div class="doc-number-box">
                <div class="doc-number-label">N&ordm; Folha Tarefa</div>
                <div class="doc-number-value">{codigo_base_folha or '-'}</div>
            </div>
        </div>
        """

    html_anexo = f"""
    <div class="page">
        {cabecalho_html}
        <div class="section-title" style="font-size:10pt; margin-bottom:8px;">Anexo &mdash; Fotos das Ocorr&ecirc;ncias</div>
        <div class="anexo-fotos-grid">{cartoes_html}</div>
        <div class="footer-signatures">
            <table class="signatures-table">
                <tr><th>Fabricante</th><th>Constru&ccedil;&atilde;o e Montagem</th><th>Qualidade</th><th>Fiscaliza&ccedil;&atilde;o</th></tr>
                <tr>
                    <td><div class="sign-line">Assinatura / Carimbo</div><div class="sign-date">Data: ____/____/________</div></td>
                    <td><div class="sign-line">Assinatura / Carimbo</div><div class="sign-date">Data: ____/____/________</div></td>
                    <td><div class="sign-line">Assinatura / Carimbo</div><div class="sign-date">Data: ____/____/________</div></td>
                    <td><div class="sign-line">Assinatura / Carimbo</div><div class="sign-date">Data: ____/____/________</div></td>
                </tr>
            </table>
            __PAGEINFO__
            <div class="developer-credit">Desenvolvido por F&aacute;bio Lima</div>
        </div>
    </div>
    """
    return fotos_labels, html_anexo


# ==============================================================================
# FORMATACAO DE VALORES
# ==============================================================================

def carregar_logo_base64(caminho_logo=None):
    caminho = caminho_logo or LOGO_PATH
    if not os.path.exists(caminho):
        return None
    with open(caminho, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def montar_logo_html(logo_b64):
    if logo_b64:
        return f'<img class="logo-img" src="data:image/png;base64,{logo_b64}" />'
    return '<div class="logo-placeholder">LOGO DA OBRA</div>'


def metros_para_mm(valor_m):
    try:
        if str(valor_m).strip() in ["-", "nan", "NaT", "None", ""]:
            return "-"
        v = str(valor_m).replace(",", ".")
        return f"{float(v) * 1000:.0f}"
    except Exception:
        return str(valor_m)


def formatar_data(val):
    """Devolve a data no padrao brasileiro com ano de 2 digitos (ex: 23/07/26)."""
    val_str = str(val).strip()
    if val_str in ["-", "nan", "NaT", "None", "", "NaN"]:
        return "-"

    # Numero serial do Excel (ex: 46221.0)
    try:
        f_val = float(val_str.replace(",", "."))
        if f_val > 30000:
            dt = pd.to_datetime(f_val, unit="D", origin="1899-12-30")
            return dt.strftime("%d/%m/%y")
    except Exception:
        pass

    # Qualquer outro formato de data (com ou sem hora): "2026-07-18 00:00:00",
    # "18/07/2026", "18/07/26"...
    val_sem_hora = val_str.split(" ")[0]
    # formato ISO (AAAA-MM-DD) nao e ambiguo, entao nao precisa de dayfirst
    eh_iso = len(val_sem_hora) >= 8 and val_sem_hora[4:5] == "-"
    dt = pd.to_datetime(val_sem_hora, errors="coerce", dayfirst=not eh_iso)
    if pd.notna(dt):
        return dt.strftime("%d/%m/%y")

    return val_sem_hora[:10]


def check_realizado(data_val, manual=False):
    fmt_date = formatar_data(data_val)
    if fmt_date == "-":
        return '<span class="pendente">[ &nbsp;&nbsp;&nbsp; ]</span>'
    if manual:
        return f'<span class="feito-manual">{fmt_date} <span class="tag-im">(IM)</span></span>'
    return f'<span class="feito">{fmt_date}</span>'


def check_texto(text_val):
    val = str(text_val).strip()
    if val in ["-", "nan", "NaT", "None", "", "NaN", "0"]:
        return '<span class="pendente">[ &nbsp;&nbsp;&nbsp; ]</span>'
    return f'<span class="prog">{val}</span>'


# ==============================================================================
# CSS DO DOCUMENTO (paleta baseada na logo Possebon / CPL)
# ==============================================================================

CSS_DOCUMENTO = """
    @page {
        size: A4;
        margin: 30mm 8mm 28mm 8mm;
        background-color: #ffffff;
        @top-center { content: element(cabecalho-corrente); width: 194mm; margin-bottom: 3mm; }
        @bottom-center { content: element(rodape-corrente); width: 194mm; margin-top: 3mm; }
    }
    *, *::before, *::after { box-sizing: border-box; }
    body {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: #1e293b; margin: 0; padding: 0; line-height: 1.2;
    }
    .page { position: relative; page-break-after: always; }

    :root {
        --azul-principal: #0093d3;
        --azul-claro: #60b1df;
        --azul-fundo: #eaf6fc;
    }

    .doc-header {
        position: running(cabecalho-corrente);
        display: flex; align-items: center; justify-content: space-between;
        border-bottom: 3px solid #0093d3; padding-bottom: 5px; width: 100%;
    }
    .logo-box { width: 30%; display: flex; align-items: center; }
    .logo-img { max-height: 68px; max-width: 100%; object-fit: contain; }
    .logo-placeholder {
        border: 1.5px dashed #94a3b8; color: #94a3b8; font-size: 7pt;
        text-align: center; padding: 10px 4px; width: 100%; border-radius: 3px;
    }
    .header-center { width: 44%; text-align: center; }
    .header-title {
        font-size: 11pt; font-weight: bold; color: #0093d3; text-transform: uppercase;
        margin: 0;
    }
    .header-sub {
        font-size: 7pt; font-weight: bold; color: #475569; letter-spacing: 0.4px;
        margin-top: 1px; text-transform: uppercase;
    }
    .doc-number-box { width: 26%; text-align: right; }
    .doc-number-label { font-size: 6.5pt; color: #64748b; font-weight: bold; text-transform: uppercase; }
    .doc-number-value { font-size: 13pt; font-weight: 900; color: #0093d3; }

    .highlight-bar {
        display: flex; justify-content: space-between; background-color: var(--azul-fundo);
        border: 1.5px solid var(--azul-claro); border-radius: 4px; margin-bottom: 6px;
    }
    .hl-item { width: 33.33%; padding: 6px 10px; text-align: center; border-right: 1.5px solid var(--azul-claro); }
    .hl-item:last-child { border-right: none; }
    .hl-title { font-size: 7.5pt; color: #475569; font-weight: bold; margin-bottom: 2px; text-transform: uppercase; }
    .hl-value { font-size: 13pt; color: #0f172a; font-weight: 900; }

    .status-banner {
        background-color: #0093d3; color: #ffffff; text-align: center; padding: 6px;
        font-size: 11pt; font-weight: bold; border-radius: 4px; margin-bottom: 8px;
        text-transform: uppercase; letter-spacing: 0.5px;
    }

    .tech-grid {
        display: flex; flex-wrap: wrap; border: 1px solid #cbd5e1; border-radius: 4px;
        margin-bottom: 8px; background-color: #fafafa;
    }
    .tech-item { width: 20%; padding: 4px 6px; border-right: 1px solid #cbd5e1; border-bottom: 1px solid #cbd5e1; }
    .tech-item:nth-child(5n) { border-right: none; }
    .tech-item:nth-last-child(-n+5) { border-bottom: none; }
    .t-label { font-size: 6pt; color: #64748b; font-weight: bold; text-transform: uppercase; }
    .t-value { font-size: 8pt; color: #1e293b; font-weight: bold; margin-top: 1px; }

    .section-title { font-size: 8.5pt; font-weight: bold; color: #0093d3; margin: 8px 0 3px 0; text-transform: uppercase; }
    .checklist-table { width: 100%; border-collapse: collapse; margin-bottom: 6px; font-size: 7.5pt; }
    .checklist-table th {
        background-color: var(--azul-fundo); color: #1e293b; font-weight: bold; text-align: center;
        padding: 5px 2px; border: 1px solid var(--azul-claro); font-size: 7pt;
    }
    .checklist-table td { border: 1px solid #cbd5e1; text-align: center; padding: 5px 2px; vertical-align: middle; }

    .checklist-table td.label-col {
        text-align: left; font-weight: bold; background-color: #f8fafc; padding-left: 8px;
        width: 35%; text-transform: uppercase; font-size: 6.8pt;
    }
    .checklist-table td.status-col { width: 15%; font-size: 7pt; padding: 4px 2px; }

    .checklist-table td.joint-col { text-align: left; font-weight: bold; background-color: #f8fafc; padding-left: 4px; font-size: 7pt; }
    .checkbox-cell { font-size: 10pt; color: #94a3b8; font-weight: normal; }

    .feito { color: #15803d; font-weight: bold; font-size: 6.5pt; }
    .feito-manual { color: #dc2626; font-weight: bold; font-size: 6.5pt; }
    .tag-im { color: #dc2626; font-weight: 900; font-size: 6pt; }
    .pendente { font-size: 10pt; color: #94a3b8; }
    .prog { color: #0093d3; font-weight: bold; font-size: 6.5pt; }
    .naplicavel { color: #94a3b8; font-style: italic; font-size: 6.8pt; }

    .legenda-tabela { font-size: 6.5pt; color: #64748b; margin: -3px 0 8px 2px; font-style: italic; }

    .ocorrencias-box { border: 1px solid #cbd5e1; border-radius: 4px; margin-bottom: 8px; overflow: hidden; }
    .ocorrencia-item { padding: 6px 8px; border-bottom: 1px solid #e2e8f0; font-size: 7pt; }
    .ocorrencia-item:last-child { border-bottom: none; }
    .ocorrencia-linha1 { display: flex; justify-content: space-between; margin-bottom: 2px; }
    .ocorrencia-setor { font-weight: bold; color: #0093d3; text-transform: uppercase; font-size: 6.5pt; }
    .ocorrencia-datahora { color: #64748b; font-size: 6.3pt; }
    .ocorrencia-atividade { font-weight: bold; color: #1e293b; }
    .ocorrencia-motivo { color: #b45309; font-style: italic; }
    .ocorrencia-usuario { color: #475569; font-size: 6.5pt; }
    .ocorrencia-fotos { margin-top: 4px; }
    .ocorrencia-fotos img {
        max-height: 60px; max-width: 90px; object-fit: cover; margin-right: 4px;
        border: 1px solid #cbd5e1; border-radius: 3px;
    }
    .ocorrencia-fotos-ref { margin-top: 3px; font-size: 6.5pt; color: #0093d3; font-weight: bold; }
    .sem-ocorrencias { font-size: 7pt; color: #94a3b8; font-style: italic; padding: 6px; }

    .anexo-fotos-grid { display: flex; flex-wrap: wrap; gap: 8px; }
    .anexo-foto-card {
        width: 30%; border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px;
        text-align: center; background: #fafafa;
    }
    .anexo-foto-card img { max-width: 100%; max-height: 140px; object-fit: contain; border-radius: 3px; }
    .anexo-foto-titulo { font-weight: bold; color: #0093d3; font-size: 8pt; margin-top: 4px; }
    .anexo-foto-legenda { font-size: 6.5pt; color: #64748b; margin-top: 2px; }

    .footer-pageinfo { text-align: center; font-size: 6.5pt; color: #64748b; margin-top: 4px; }
    .developer-credit { text-align: right; font-size: 6pt; color: #94a3b8; font-style: italic; margin-top: 2px; }

    .footer-signatures { position: running(rodape-corrente); width: 100%; }
    .signatures-table { width: 100%; border-collapse: collapse; font-size: 7.5pt; }
    .signatures-table th {
        background-color: var(--azul-fundo); border: 1px solid #cbd5e1; color: #0093d3;
        padding: 4px 2px; text-align: center; width: 25%; font-size: 6.8pt;
    }
    .signatures-table td { border: 1px solid #cbd5e1; vertical-align: bottom; height: 45px; text-align: center; padding: 5px; }
    .sign-line { border-top: 1px solid #64748b; margin-top: 20px; padding-top: 2px; font-weight: bold; color: #334155; font-size: 6.5pt; }
    .sign-date { text-align: left; margin-top: 2px; font-size: 6pt; color: #64748b; }
"""


def montar_bloco_info_spool(spool, peso_total_prog=None):
    """Bloco de identificacao do spool (Isometrico/Linha/Spool/Peso Total da
    Programacao + situacao SGER + dados tecnicos), sem o cabecalho de pagina
    nem os checklists. Usado na tela e no PDF. Usa o STATUS DE FABRICACAO
    (SGER), ja que esta Folha Tarefa e de Fabricacao (nao de Montagem)."""
    iso_atual = str(spool.get("Isométrico", "-"))
    num_spool_atual = str(spool.get("Spool", "-"))
    comp_mm = metros_para_mm(spool.get("Comp (m)", "-"))
    peso_prog_texto = formatar_peso_br(peso_total_prog) if peso_total_prog is not None else "-"

    return f"""
        <div class="highlight-bar">
            <div class="hl-item" style="width: 20%;">
                <div class="hl-title">Isom&eacute;trico</div>
                <div class="hl-value">{iso_atual}</div>
            </div>
            <div class="hl-item" style="width: 37%;">
                <div class="hl-title">Tag da Linha</div>
                <div class="hl-value">{spool.get('Linha', '-')}</div>
            </div>
            <div class="hl-item" style="width: 12%;">
                <div class="hl-title">Spool</div>
                <div class="hl-value">{num_spool_atual}</div>
            </div>
            <div class="hl-item" style="width: 31%;">
                <div class="hl-title">Peso Total da Programa&ccedil;&atilde;o (kg)</div>
                <div class="hl-value">{peso_prog_texto}</div>
            </div>
        </div>

        <div class="status-banner">
            SITUA&Ccedil;&Atilde;O (SGER): {spool.get('STATUS DE FABRICAÇÃO (SGER)', '-')}
        </div>

        <div class="tech-grid">
            <div class="tech-item"><div class="t-label">Fabricante</div><div class="t-value">{spool.get('Site Fabricante', '-')}</div></div>
            <div class="tech-item"><div class="t-label">Material</div><div class="t-value">{spool.get('Mat', '-')}</div></div>
            <div class="tech-item"><div class="t-label">Di&acirc;metro Nominal</div><div class="t-value">{spool.get('Ø', '-')}</div></div>
            <div class="tech-item"><div class="t-label">Espessura</div><div class="t-value">{spool.get('ESP', '-')}</div></div>
            <div class="tech-item"><div class="t-label">Peso (kg)</div><div class="t-value">{spool.get('Peso', '-')}</div></div>

            <div class="tech-item"><div class="t-label">Compr. (mm)</div><div class="t-value">{comp_mm}</div></div>
            <div class="tech-item"><div class="t-label">&Aacute;rea Pint. (m²)</div><div class="t-value">{spool.get('Area_m²', '-')}</div></div>
            <div class="tech-item"><div class="t-label">Cond. Pintura</div><div class="t-value">{spool.get('Cond_Pint', '-')}</div></div>
            <div class="tech-item"><div class="t-label">Prog. Fab.</div><div class="t-value" style="color: #0093d3;">{spool.get('Prog_Fab', '-')}</div></div>
            <div class="tech-item"><div class="t-label">Romaneio</div><div class="t-value" style="color: #0093d3;">{spool.get('Romaneio', '-')}</div></div>
        </div>
        """


def montar_tabela_juntas_html(spool, df_juntas):
    """Monta so a tabela de checklist por junta (reutilizada no PDF e na tela).
    So mostra juntas de PIPE (fabricacao) - juntas de CAMPO ficam de fora,
    pois essa Folha Tarefa e de Fabricacao."""
    iso_atual = str(spool.get("Isométrico", "-"))
    num_spool_atual = str(spool.get("Spool", "-"))

    juntas_do_spool = df_juntas[
        (df_juntas["Isometrico"].astype(str) == iso_atual)
        & (df_juntas["Spool"].astype(str).str.lstrip("0") == num_spool_atual.lstrip("0"))
    ].copy()

    coluna_local = "C/P" if "C/P" in juntas_do_spool.columns else "Local"
    if coluna_local in juntas_do_spool.columns:
        juntas_do_spool = juntas_do_spool[
            juntas_do_spool[coluna_local].astype(str).str.strip().str.upper() == "PIPE"
        ]

    juntas_do_spool["Junta_Num"] = pd.to_numeric(juntas_do_spool["Junta"], errors="coerce")
    juntas_do_spool = juntas_do_spool.sort_values(by="Junta_Num")

    juntas_html = ""
    for _, junta in juntas_do_spool.iterrows():
        tipo_junta = str(junta.get("Tp_Junta", "-")).strip().upper()
        if tipo_junta == "AN":
            va_html = '<span class="naplicavel">N/A</span>'
        else:
            va_html = check_realizado(junta.get("Data_Va", "-"))

        juntas_html += f"""
        <tr>
            <td class="joint-col">{junta.get('Junta', '-')}<br>
                <span style="font-size: 6pt; color: #64748b; font-weight: normal;">Sit: {junta.get('Status_Junta', '-')}</span>
            </td>
            <td>{junta.get('N_Insp.', '-')}</td>
            <td>{junta.get('Ø', '-')} / {junta.get('Espess.', '-')}</td>
            <td>{junta.get('Tp_Junta', '-')}</td>
            <td>{junta.get(coluna_local, '-')}</td>
            <td>{check_texto(junta.get('Prog.Fabr.', '-'))}</td>
            <td>{va_html}</td>
            <td>{check_realizado(junta.get('Data_Solda', '-'))}</td>
            <td>{check_realizado(junta.get('Data_Vs', '-'))}</td>
            <td>{check_realizado(junta.get('Data_Rx/Us', '-'))}</td>
        </tr>
        """
    if juntas_html == "":
        juntas_html = "<tr><td colspan='10'>Nenhuma junta de PIPE encontrada para este spool.</td></tr>"

    tabela = f"""
    <table class="checklist-table">
        <thead>
            <tr>
                <th style="width: 15%;">N&ordm; da Junta</th><th style="width: 8%;">N&iacute;vel Insp.</th><th style="width: 11%;">Dim / Esp</th>
                <th style="width: 10%;">Tipo</th><th style="width: 9%;">Local</th><th style="width: 9%;">Prog.</th>
                <th style="width: 11%;">VA (Ajuste)</th><th style="width: 9%;">Solda</th>
                <th style="width: 11%;">Exame Visual</th><th style="width: 7%;">END</th>
            </tr>
        </thead>
        <tbody>
            {juntas_html}
        </tbody>
    </table>
    <div class="legenda-tabela">N/A = N&atilde;o Aplic&aacute;vel &middot; <span class="tag-im">(IM)</span> = Inserido Manualmente</div>
    """
    return tabela, len(juntas_do_spool)


def montar_checklist_spool_html(spool, campos_manuais=None):
    """Tabela do checklist do spool, com o MESMO visual do PDF (bordas/grade).
    campos_manuais: conjunto de nomes de coluna cujo valor foi preenchido
    manualmente na tela (aparecem em vermelho com a marca (IM))."""
    campos_manuais = campos_manuais or set()

    def _cr(campo):
        return check_realizado(spool.get(campo), manual=(campo in campos_manuais))

    return f"""
    <table class="checklist-table">
        <tbody>
            <tr>
                <td class="label-col">Corte</td>
                <td class="status-col">{_cr('Data_Corte')}</td>
                <td class="label-col">E.N.D</td>
                <td class="status-col">{_cr('Data_End_Fab')}</td>
            </tr>
            <tr>
                <td class="label-col">VA (Ajuste)</td>
                <td class="status-col">{_cr('Data_VA_Fab')}</td>
                <td class="label-col">Dimensional de F&aacute;brica</td>
                <td class="status-col">{_cr('Data_Dim_Fab')}</td>
            </tr>
            <tr>
                <td class="label-col">Solda</td>
                <td class="status-col">{_cr('Data_Soldagem_Fab')}</td>
                <td class="label-col">Pintura de Fundo</td>
                <td class="status-col">{_cr('Data_Pint_Fundo')}</td>
            </tr>
            <tr>
                <td class="label-col">Exame Visual</td>
                <td class="status-col">{_cr('Data_VS_Fab')}</td>
                <td class="label-col">Pintura Interm.</td>
                <td class="status-col">{_cr('Data_Pint_Intermediária')}</td>
            </tr>
        </tbody>
    </table>
    """


# ==============================================================================
# MONTAGEM DE UMA PAGINA (1 spool)
# ==============================================================================

def montar_pagina_spool(spool_original, df_juntas, numero_documento, logo_html,
                         pagina_atual=None, total_paginas=None, fotos_labels=None,
                         peso_total_prog=None):
    overrides = carregar_status_manual()
    spool = obter_spool_com_overrides(spool_original, overrides)
    campos_manuais = campos_manuais_do_spool(spool_original, overrides)

    iso_atual = str(spool.get("Isométrico", "-"))
    num_spool_atual = str(spool.get("Spool", "-"))

    tabela_juntas_html, _qtd_juntas = montar_tabela_juntas_html(spool, df_juntas)
    chave_atual = chave_spool(spool.get("Unidade", "-"), iso_atual, num_spool_atual)
    bloco_ocorrencias_html = montar_bloco_ocorrencias_html(chave_atual, fotos_labels=fotos_labels)

    pageinfo_html = ""
    if pagina_atual and total_paginas:
        pageinfo_html = f'<div class="footer-pageinfo">Folha {pagina_atual} de {total_paginas}</div>'

    return f"""
    <div class="page">
        <div class="doc-header">
            <div class="logo-box">{logo_html}</div>
            <div class="header-center">
                <div class="header-title">Folha Tarefa &ndash; Controle de Fabrica&ccedil;&atilde;o e Montagem</div>
                <div class="header-sub">Cons&oacute;rcio Possebon / CPL &middot; UTAA/TR &middot; CPC RNEST</div>
            </div>
            <div class="doc-number-box">
                <div class="doc-number-label">N&ordm; Folha Tarefa</div>
                <div class="doc-number-value">{numero_documento}</div>
            </div>
        </div>

        {montar_bloco_info_spool(spool, peso_total_prog=peso_total_prog)}

        <div class="section-title">1. Checklist do Spool (Acompanhamento)</div>
        {montar_checklist_spool_html(spool, campos_manuais=campos_manuais)}

        <div class="section-title">2. Checklist por Junta Soldada</div>
        {tabela_juntas_html}

        <div class="section-title">3. Ocorr&ecirc;ncias / Pend&ecirc;ncias</div>
        {bloco_ocorrencias_html}

        <div class="footer-signatures">
            <table class="signatures-table">
                <tr><th>Fabricante</th><th>Constru&ccedil;&atilde;o e Montagem</th><th>Qualidade</th><th>Fiscaliza&ccedil;&atilde;o</th></tr>
                <tr>
                    <td><div class="sign-line">Assinatura / Carimbo</div><div class="sign-date">Data: ____/____/________</div></td>
                    <td><div class="sign-line">Assinatura / Carimbo</div><div class="sign-date">Data: ____/____/________</div></td>
                    <td><div class="sign-line">Assinatura / Carimbo</div><div class="sign-date">Data: ____/____/________</div></td>
                    <td><div class="sign-line">Assinatura / Carimbo</div><div class="sign-date">Data: ____/____/________</div></td>
                </tr>
            </table>
            {pageinfo_html}
            <div class="developer-credit">Desenvolvido por F&aacute;bio Lima</div>
        </div>
    </div>
    """


def montar_documento_completo(spools_filtrados, df_juntas, codigo_base_folha, logo_html, df_spools_completo=None):
    """Monta o HTML completo (todas as paginas) para o conjunto de spools selecionado.
    As fotos das ocorrencias ficam reunidas numa pagina de anexo ao final
    (Foto-01, Foto-02...), referenciadas em cada Folha Tarefa.

    df_spools_completo: DataFrame com TODOS os spools da planilha (nao so os
    filtrados/selecionados), usado para calcular o peso total de cada
    Programacao de Fabricacao corretamente mesmo quando so parte dos spools
    foi selecionada para gerar a Folha Tarefa. Se nao informado, usa apenas
    os spools deste lote."""
    df_base_peso = df_spools_completo if df_spools_completo is not None else spools_filtrados
    cache_peso_prog = {}

    def _peso_prog_cache(valor):
        if valor not in cache_peso_prog:
            cache_peso_prog[valor] = peso_total_programacao(df_base_peso, valor)
        return cache_peso_prog[valor]

    fotos_labels, html_anexo = montar_anexo_fotos(spools_filtrados, logo_html=logo_html, codigo_base_folha=codigo_base_folha)
    total_paginas = len(spools_filtrados) + (1 if html_anexo else 0)

    html_completo = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>{CSS_DOCUMENTO}</style>
</head>
<body>
"""
    pagina_num = 0
    for _, spool in spools_filtrados.iterrows():
        pagina_num += 1
        numero_documento = f"{codigo_base_folha}/{pagina_num}"
        chave_spool_atual = chave_spool(
            spool.get("Unidade", "-"), spool.get("Isométrico", "-"), spool.get("Spool", "-")
        )
        html_completo += montar_pagina_spool(
            spool, df_juntas, numero_documento, logo_html,
            pagina_atual=pagina_num, total_paginas=total_paginas,
            fotos_labels=fotos_labels.get(chave_spool_atual, {}),
            peso_total_prog=_peso_prog_cache(spool.get("Prog_Fab", "-")),
        )

    if html_anexo:
        html_anexo = html_anexo.replace(
            "__PAGEINFO__",
            f'<div class="footer-pageinfo">Folha {total_paginas} de {total_paginas}</div>',
        )
        html_completo += html_anexo

    html_completo += "\n</body>\n</html>"
    return html_completo


def gerar_pdf_bytes(html_completo):
    """Renderiza o HTML em PDF e devolve os bytes (sem gravar em disco)."""
    return HTML(string=html_completo).write_pdf()


def gerar_pdf_arquivo(html_completo, caminho_pdf):
    """Renderiza o HTML em PDF e grava direto no caminho informado."""
    HTML(string=html_completo).write_pdf(caminho_pdf)
