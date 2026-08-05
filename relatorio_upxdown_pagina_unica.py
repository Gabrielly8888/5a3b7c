# -*- coding: utf-8 -*-
r"""
Relatório Executivo Up x Down — HTML v4
Entrada : updown.xlsx
Saída   : relatorio upxdown.html
"""

import json
import re
import unicodedata
from datetime import datetime, timedelta, timezone
from pathlib import Path

import numpy as np
import pandas as pd

# Conexão com banco de dados (opcional — se falhar, faixa de vida fica como "Não calculado")
try:
    import psycopg2
    PSYCOPG2_OK = True
except ImportError:
    PSYCOPG2_OK = False

PASTA            = Path(__file__).resolve().parent
ARQUIVO_EXCEL    = PASTA / "updown.xlsx"
ARQUIVO_SAIDA    = PASTA / "relatorio upxdown.html"
ARQUIVO_BASE_CSV = PASTA / "base_tratada_upxdown.csv"
ARQUIVO_OUTROS_CSV = PASTA / "analitico_movimentos_outros.csv"
ARQUIVO_ANALITICO_CONSIDERADOS = PASTA / "analitico_casos_considerados.csv"
ARQUIVO_ANALITICO_NAO_CONSIDERADOS = PASTA / "analitico_casos_nao_considerados.csv"
MESES_PT = {
    1:"Jan.",2:"Fev.",3:"Mar.",4:"Abr.",5:"Mai.",6:"Jun.",
    7:"Jul.",8:"Ago.",9:"Set.",10:"Out.",11:"Nov.",12:"Dez."
}
MESES_LONGO = {
    1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",
    7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro"
}

def limpar_texto(x):
    if pd.isna(x): return ""
    return str(x).replace("\t"," ").replace("\n"," ").strip()

def normalizar(x, padrao="Não informado"):
    x = limpar_texto(x)
    if x in ["","nan","None","—","-"]: return padrao
    return x

def remover_acentos(txt):
    txt = str(txt or "")
    return "".join(c for c in unicodedata.normalize("NFD",txt) if unicodedata.category(c)!="Mn")

def texto_chave(txt): return remover_acentos(txt).lower().strip()

def eh_nome_valido(x):
    x = limpar_texto(x)
    if not x: return False
    # Remove sufixo ".0" que o pandas gera ao converter número para string
    x_sem_ponto = re.sub(r"\.0+$", "", x).strip()
    xl = x_sem_ponto.lower()
    # Variações de "sem proprietário" e não informado
    invalidos = [
        "sem proprietário","sem proprietario","não informado","nao informado",
        "sem dono","sem responsável","sem responsavel","não identificado",
        "nao identificado","n/a","na","none","null","—","-","s/n","s/proprietario",
        "s/proprietário","usuário inexistente","usuario inexistente","não informado"
    ]
    if xl in invalidos: return False
    # Número puro (inteiro, decimal, CPF 000.000.000-00, CNPJ, telefone, etc.)
    # Remove pontos, traços, barras, parênteses, espaços e verifica se só sobram dígitos
    so_digitos = re.sub(r"[\s\.\-\/\(\)]", "", x_sem_ponto)
    if so_digitos.isdigit(): return False
    # Verifica se começa com dígito (código, matrícula, etc.)
    if re.match(r"^\d", x_sem_ponto): return False
    letras = len(re.findall(r"[A-Za-zÀ-ÿ]", x_sem_ponto))
    nums   = len(re.findall(r"\d", x_sem_ponto))
    # Precisa ter pelo menos 3 letras e nenhum dígito
    return letras >= 3 and nums == 0

def encurtar_produto(x, limite=46):
    x = normalizar(x)
    if len(x) <= limite: return x
    partes = [p.strip() for p in re.split(r";|,|\+",x) if p.strip()]
    if len(partes) >= 2:
        curto = partes[0]+" + outros"
        return curto if len(curto)<=limite else curto[:limite-1]+"…"
    return x[:limite-1]+"…"

def classificar_movimento(row):
    """
    Classificação oficial dos movimentos.

    Regra atual:
    - Upgrade: tipo/tema = Alteração no contrato (Up) ou Integração de prateleira
    - Downgrade: tipo/tema = Alteração no contrato (Down)
    - Criação de URL: tipo/tema = Criação de URL
    - Migração: tipo/tema = Migração para número oficial
    - Robô: tipo/tema = Robô personalizado
    - Teste: tipo/tema = Teste
    - Alteração de Plano: tipo/tema = Alteração de Plano

    Observação: no arquivo atual esses nomes estão principalmente na coluna "tipo".
    Mesmo assim, a função olha também "tema" para evitar erro caso a base venha com
    essa informação em outra coluna.
    """
    tipo = texto_chave(row.get("tipo", ""))
    tema = texto_chave(row.get("tema", ""))
    tipo_servico = texto_chave(row.get("tipo_servico", ""))

    campos = [tipo, tema, tipo_servico]

    def tem_exato(valor):
        valor = texto_chave(valor)
        return any(c == valor for c in campos)

    def contem(valor):
        valor = texto_chave(valor)
        return any(valor in c for c in campos if c)

    if tem_exato("Alteração no contrato (Down)"):
        return "Downgrade"

    if tem_exato("Alteração no contrato (Up)") or tem_exato("Integração de prateleira"):
        return "Upgrade"

    if tem_exato("Criação de URL"):
        return "Criação de URL"

    if tem_exato("Migração para número oficial"):
        return "Migração"

    if tem_exato("Robô personalizado") or contem("robo personalizado"):
        return "Robô"

    if tem_exato("Teste"):
        return "Teste"

    if tem_exato("Alteração de Plano"):
        return "Alteração de Plano"

    # Compatibilidade com bases antigas em que a coluna "tema" vinha resumida
    # como Upgrade/Downgrade/Migração/Solicitação de Teste/Solicitação de robô.
    # Mantém o código funcionando sem puxar Bonificação/Desconto para o dashboard.
    if tema == "downgrade":
        return "Downgrade"
    if tema == "upgrade" and (tipo in ["", "nao informado", "—", "-"] or tipo_servico == "up"):
        return "Upgrade"
    if "migracao para numero oficial" in tema:
        return "Migração"
    if "solicitacao de teste" in tema or tema == "teste":
        return "Teste"
    if "solicitacao de robo" in tema or "robo personalizado" in tema:
        return "Robô"
    if "criacao de url" in tema:
        return "Criação de URL"

    return "__EXCLUIR__"

def classificar_motivo_down(row):
    txt = texto_chave(str(row.get("motivo_downgrade",""))+" "+str(row.get("descricao","")))
    if not txt or txt in ["nao informado","sem informacao","—","-"]: return "Sem detalhe informado"
    if any(k in txt for k in ["custo","redu","finance","caro","valor","preco","preço","orçamento","orcamento","cortar gasto","econom"]): return "Redução de custo"
    if any(k in txt for k in ["nao usa","nao utiliza","nao usou","sem uso","baixa utilizacao","pouco uso","nao adapt","nao achou util","nao viu valor","sem valor agregado"]): return "Baixa utilização / sem valor percebido"
    if any(k in txt for k in ["cancel","encerr","churn","parou","deixou de","nao quer mais"]): return "Cancelamento / encerramento"
    if any(k in txt for k in ["instab","erro","bug","problema","falha","lento","nao funciona","trav","limitacao","limitação","tecnico","técnico"]): return "Problema técnico / limitação"
    if any(k in txt for k in ["manual","automat","processo","operacao","operação","layout","organiz","visualizacao","identificacao"]): return "Processo/uso operacional"
    if any(k in txt for k in ["outra ferramenta","outras ferramentas","teams","zoho","liderhub","cpi","solucao propria","solução própria","preferem outra","migrando"]): return "Troca de solução / concorrente"
    if any(k in txt for k in ["desnecess","nao precisa","não precisa","sem necessidade"]): return "Produto desnecessário"
    return "Outros motivos descritos"

if not ARQUIVO_EXCEL.exists():
    # Mantém o caminho oficial configurado acima.
    # Se o script for executado em outra pasta para teste/validação, usa um updown.xlsx
    # salvo ao lado do próprio .py. No seu computador, se o caminho oficial existir,
    # nada muda.
    _arquivo_excel_local = Path(__file__).with_name("updown.xlsx")
    if _arquivo_excel_local.exists():
        PASTA = _arquivo_excel_local.parent
        ARQUIVO_EXCEL = _arquivo_excel_local
        ARQUIVO_SAIDA = PASTA / "relatorio upxdown.html"
        ARQUIVO_BASE_CSV = PASTA / "base_tratada_upxdown.csv"
        ARQUIVO_OUTROS_CSV = PASTA / "analitico_movimentos_outros.csv"
        ARQUIVO_ANALITICO_CONSIDERADOS = PASTA / "analitico_casos_considerados.csv"
        ARQUIVO_ANALITICO_NAO_CONSIDERADOS = PASTA / "analitico_casos_nao_considerados.csv"
    else:
        raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO_EXCEL}")

df = pd.read_excel(ARQUIVO_EXCEL)
df.columns = [str(c).strip() for c in df.columns]

colunas_necessarias = ["dia","proprietario","tema","produtos_digisac","tipo_servico","motivo_downgrade","descricao","qtd_servicos","total_up_RS","total_down_RS"]
faltantes = [c for c in colunas_necessarias if c not in df.columns]
if faltantes: raise ValueError(f"Colunas não encontradas: {faltantes}")

if "subdominio" not in df.columns:
    df["subdominio"] = "Não informado"

if "card_financeiro" not in df.columns:
    df["card_financeiro"] = ""

if "tipo" not in df.columns:
    df["tipo"] = ""

df["dia"] = pd.to_datetime(df["dia"], errors="coerce", dayfirst=True)
df = df[~df["dia"].isna()].copy()

for col in ["proprietario","tema","tipo","produtos_digisac","tipo_servico","motivo_downgrade","descricao","subdominio"]:
    df[col] = df[col].apply(normalizar)

# Remove linhas sem subdomínio — sem subdomínio não é possível identificar a conta
SEM_SUBDOMINIO = {"não informado", "nao informado", "—", "-", "", "nan", "none", "null"}
antes = len(df)
df = df[~df["subdominio"].str.lower().str.strip().isin(SEM_SUBDOMINIO)].copy()
print(f"ℹ️  Linhas removidas por ausência de subdomínio: {antes - len(df)} (restaram {len(df)})")

# Separa os casos com e sem link no Card Financeiro
# - com link: entram no relatório
# - sem link: ficam em um analítico separado para auditoria
antes = len(df)
df["card_financeiro"] = (
    df["card_financeiro"]
    .fillna("")
    .astype(str)
    .str.strip()
)
_mask_card_financeiro = df["card_financeiro"].str.startswith("http", na=False)
df_sem_card_financeiro = df[~_mask_card_financeiro].copy()
df = df[_mask_card_financeiro].copy()
print(f"ℹ️  Linhas separadas sem Card Financeiro: {len(df_sem_card_financeiro)} | consideradas: {len(df)}")


# ==========================================
# Padronização dos produtos
# Regra importante:
# - Quando vier mais de um produto separado por ";", cada item é tratado separado.
# - Exemplo:
#   "Funil de vendas;IA - Resumo Inteligente - Plano 1000;IA - Texto Mágico - Plano 1000"
#   vira:
#   "Funil de vendas; Resumo Inteligente; Texto Mágico"
# ==========================================
def padronizar_produto_item(produto):
    produto = normalizar(produto)
    produto = re.sub(r"\s+", " ", produto).strip()

    chave = texto_chave(produto)

    # IA CSAT
    if re.search(r"\bia csat\b", chave):
        return "IA CSAT"

    # IA Copiloto
    if re.search(r"\bia copiloto\b", chave):
        return "IA Copiloto"

    # Transcrição de Áudio
    if "transcricao de audio" in chave:
        return "Transcrição de Áudio"

    # IA Agent / IA Agente
    if re.search(r"\bia agente\b|\bia agent\b", chave):
        return "IA Agent"

    # Texto Mágico
    if "texto magico" in chave:
        return "Texto Mágico"

    # Resumo Inteligente
    if "resumo inteligente" in chave:
        return "Resumo Inteligente"

    # Combo de IA
    if "combo de ia" in chave:
        return "Combo de IA's"

    return produto


def padronizar_produtos_digisac(valor):
    valor = normalizar(valor)

    # Se tiver mais de um produto, separa pelo ponto e vírgula
    partes = [p.strip() for p in str(valor).split(";") if p.strip()]

    if not partes:
        return valor

    partes_padronizadas = []
    for parte in partes:
        novo = padronizar_produto_item(parte)
        if novo and novo not in partes_padronizadas:
            partes_padronizadas.append(novo)

    return "; ".join(partes_padronizadas)


df["produtos_digisac"] = df["produtos_digisac"].apply(padronizar_produtos_digisac)


EXCLUIR = {"gabrielly oliveira","nicholas prudente","processos interno","processos internos"}
df = df[~df["proprietario"].apply(lambda x: texto_chave(x) in EXCLUIR)].copy()

df["qtd_servicos"] = df["produtos_digisac"].apply(
    lambda x: len([p for p in str(x).split(";") if p.strip() and p.strip().lower() not in ["não informado","nao informado","nan",""]])
).clip(lower=1)  # mínimo 1 para não zerar linhas sem produto identificado
df["total_up_RS"]   = pd.to_numeric(df["total_up_RS"],  errors="coerce").fillna(0)
df["total_down_RS"] = pd.to_numeric(df["total_down_RS"],errors="coerce").fillna(0)

df["Ano"]     = df["dia"].dt.year.astype(int)
df["MesNum"]  = df["dia"].dt.month.astype(int)
df["Mês"]     = df["MesNum"].map(MESES_PT)
df["MesLongo"]= df["MesNum"].map(MESES_LONGO)
df["AnoMes"]  = df["Ano"].astype(str)+"-"+df["MesNum"].astype(str).str.zfill(2)
df["Periodo"] = df["Mês"]+"/"+df["Ano"].astype(str).str[-2:]
df["Data"]    = df["dia"].dt.strftime("%d/%m/%Y")
df["Movimento"]     = df.apply(classificar_movimento, axis=1)
df["Situacao_Down"] = df.apply(classificar_motivo_down, axis=1)

# ── EXCLUSÃO: tudo que não é um movimento válido é descartado ─────────────
# (antes virava "Outros"; agora não é considerado em nada).
# Guarda o que foi removido num CSV só para você auditar.
df_excluidos = df[df["Movimento"] == "__EXCLUIR__"].copy()
if len(df_excluidos):
    cols_aud = [c for c in ["Data", "Ano", "Mês", "subdominio", "tema", "tipo", "produtos_digisac", "card_financeiro",
                            "tipo_servico", "proprietario", "motivo_downgrade", "descricao",
                            "qtd_servicos", "total_up_RS", "total_down_RS"] if c in df_excluidos.columns]
    df_excluidos[cols_aud].to_csv(PASTA / "linhas_excluidas_sem_movimento.csv",
                                  sep=";", index=False, encoding="utf-8-sig")
print(f"🗑️  Linhas excluídas (tema fora dos movimentos válidos): {len(df_excluidos)}")
if len(df_excluidos):
    print("    Temas descartados:")
    for t, q in df_excluidos["tema"].value_counts().items():
        print(f"      · {t!r}: {q}")
df = df[df["Movimento"] != "__EXCLUIR__"].copy()
print(f"✅ Linhas consideradas no relatório: {len(df)}")
# ── fim da exclusão ───────────────────────────────────────────────────────

df["Proprietario_Nome_Valido"] = df["proprietario"].apply(eh_nome_valido)
df["Produto_Curto"] = df["produtos_digisac"].apply(encurtar_produto)
df["proprietario_filtro"] = np.where(df["Proprietario_Nome_Valido"], df["proprietario"], "Usuário inexistente")

# ── DIAGNÓSTICO: remova este bloco após validar ──────────────────────────
invalidos_unicos = sorted(df[~df["Proprietario_Nome_Valido"]]["proprietario"].dropna().unique().tolist())
validos_unicos   = sorted(df[ df["Proprietario_Nome_Valido"]]["proprietario"].dropna().unique().tolist())
print(f"\n{'='*60}")
print(f"PROPRIETÁRIOS INVÁLIDOS → 'Usuário inexistente' ({len(invalidos_unicos)} únicos):")
for v in invalidos_unicos[:40]:
    print(f"  · {repr(v)}")
if len(invalidos_unicos) > 40:
    print(f"  ... e mais {len(invalidos_unicos)-40}")
print(f"\nPROPRIETÁRIOS VÁLIDOS ({len(validos_unicos)} únicos):")
for v in validos_unicos[:20]:
    print(f"  · {repr(v)}")
if len(validos_unicos) > 20:
    print(f"  ... e mais {len(validos_unicos)-20}")
print('='*60+'\n')
# ── fim do diagnóstico ───────────────────────────────────────────────────

base = df[[
    "Data","Ano","MesNum","Mês","MesLongo","AnoMes","Periodo",
    "proprietario","proprietario_filtro","Proprietario_Nome_Valido","subdominio",
    "tema","tipo","produtos_digisac","Produto_Curto","card_financeiro","tipo_servico",
    "motivo_downgrade","Situacao_Down","descricao",
    "qtd_servicos","total_up_RS","total_down_RS","Movimento"
]].copy()

# ══════════════════════════════════════════════════════════════════════
# Faixa de vida do cliente
# Consulta agnus.vw_assinaturas no banco para buscar criacao_url
# por subdominio, calcula dias entre criacao_url e dia do movimento,
# e classifica na faixa de vida.
# ══════════════════════════════════════════════════════════════════════
DB_CONFIG = dict(
    host     = "10.10.5.216",
    port     = 5433,
    dbname   = "dw",
    user     = "gabrielly.oliveira",
    password = "qNNy&QT;$1Kz$a6LQI/5",
)

def classificar_faixa_vida(dias):
    if pd.isna(dias) or dias is None:
        return "Não calculado"
    d = int(dias)
    if   d <=  30: return "1 - 0 a 30"
    elif d <=  60: return "2 - 31 a 60"
    elif d <=  90: return "3 - 61 a 90"
    elif d <= 120: return "4 - 91 a 120"
    elif d <= 180: return "5 - 121 a 180"
    elif d <= 240: return "6 - 181 a 240"
    elif d <= 360: return "7 - 241 a 360"
    elif d <= 540: return "8 - 361 a 540"
    elif d <= 720: return "9 - 541 a 720"
    else:          return "10 - Acima de 721"

def buscar_criacao_url(subdominios_unicos):
    """Consulta o banco e retorna dict {subdominio: criacao_url}."""
    if not PSYCOPG2_OK:
        print("⚠️  psycopg2 não instalado. Execute: pip install psycopg2-binary")
        return {}
    if not subdominios_unicos:
        return {}
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        # Monta a lista de subdomínios para o IN (...)
        lista = tuple(subdominios_unicos) if len(subdominios_unicos) > 1 else (list(subdominios_unicos)[0],)
        placeholders = ",".join(["%s"] * len(lista))
        query = f"""
            SELECT DISTINCT ON (subdominio)
                subdominio,
                criacao_url
            FROM agnus.vw_assinaturas
            WHERE subdominio IN ({placeholders})
            ORDER BY subdominio, criacao_url ASC
        """
        cur = conn.cursor()
        cur.execute(query, lista)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        resultado = {row[0]: row[1] for row in rows if row[1] is not None}
        print(f"✅ Faixa de vida: {len(resultado)} subdomínios encontrados no banco (de {len(subdominios_unicos)} únicos).")
        return resultado
    except Exception as e:
        print(f"⚠️  Erro ao consultar banco para faixa de vida: {e}")
        return {}

# Coleta subdomínios únicos válidos da base
subdominios_validos = set(
    base["subdominio"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
) - {"", "Não informado", "nan", "None"}

mapa_criacao = buscar_criacao_url(subdominios_validos)

# Aplica na base
base["criacao_url"] = (
    base["subdominio"]
    .astype(str)
    .str.strip()
    .map(mapa_criacao)
)
base["criacao_url"] = pd.to_datetime(base["criacao_url"], errors="coerce")
base["dia_dt"]      = pd.to_datetime(base["Data"], format="%d/%m/%Y", errors="coerce")

base["Tempo_vida_cli"] = (base["dia_dt"] - base["criacao_url"]).dt.days
base["Faixa_Vida"]    = base["Tempo_vida_cli"].apply(classificar_faixa_vida)

# Remove colunas auxiliares
base.drop(columns=["dia_dt"], inplace=True)

print(f"📊 Distribuição de faixas de vida:\n{base['Faixa_Vida'].value_counts().to_string()}\n")

# ── DIAGNÓSTICO DOWNGRADE 2025: remova após validar ──────────────────────
down2025 = base[(base["Movimento"] == "Downgrade") & (base["Ano"] == 2025)].copy()
down2025 = down2025.sort_values(["MesNum","Data"])
print(f"\n{'='*70}")
print(f"DOWNGRADES 2025 — total de linhas: {len(down2025)}  |  soma qtd_servicos: {down2025['qtd_servicos'].sum()}")
print(f"{'='*70}")
print(f"{'Data':<12} {'Mês':<6} {'Proprietário':<30} {'Tema':<35} {'Produto':<30} {'Qtd':>4} {'Down R$':>10}")
print("-"*130)
for _, r in down2025.iterrows():
    print(f"{r['Data']:<12} {r['Mês']:<6} {str(r['proprietario'])[:29]:<30} {str(r['tema'])[:34]:<35} {str(r['produtos_digisac'])[:29]:<30} {r['qtd_servicos']:>4} {r['total_down_RS']:>10.2f}")
print('='*70+'\n')

# Salva também em CSV para fácil comparação
down2025.to_csv(PASTA / "diagnostico_down2025.csv", sep=";", index=False, encoding="utf-8-sig")
print(f"📋 CSV de diagnóstico salvo em: {PASTA / 'diagnostico_down2025.csv'}")
# ── fim do diagnóstico ───────────────────────────────────────────────────

# ── DIAGNÓSTICO UPGRADE 2026: remova após validar ────────────────────────
up2026 = base[(base["Movimento"] == "Upgrade") & (base["Ano"] == 2026)].copy()
up2026 = up2026.sort_values(["MesNum","Data"])
print(f"\n{'='*70}")
print(f"UPGRADES 2026 — total de linhas: {len(up2026)}  |  soma qtd_servicos: {up2026['qtd_servicos'].sum()}")
print(f"{'='*70}")
print(f"{'Data':<12} {'Mês':<6} {'Proprietário':<30} {'Tema':<35} {'Produto':<30} {'Qtd':>4} {'Up R$':>10}")
print("-"*130)
for _, r in up2026.iterrows():
    print(f"{r['Data']:<12} {r['Mês']:<6} {str(r['proprietario'])[:29]:<30} {str(r['tema'])[:34]:<35} {str(r['produtos_digisac'])[:29]:<30} {r['qtd_servicos']:>4} {r['total_up_RS']:>10.2f}")
print('='*70+'\n')

# Salva também em CSV para fácil comparação
up2026.to_csv(PASTA / "diagnostico_up2026.csv", sep=";", index=False, encoding="utf-8-sig")
print(f"📋 CSV de diagnóstico salvo em: {PASTA / 'diagnostico_up2026.csv'}")
# ── fim do diagnóstico ───────────────────────────────────────────────────

# ==========================================================
# Base auxiliar para visão de Produtos
# Aqui o produto é separado pelo ";".
# Exemplo:
# "Servidores; Web chat" vira duas linhas:
# "Servidores" e "Web chat".
# Assim a aba Produtos soma cada produto individualmente,
# mesmo quando ele vem junto com outros produtos na mesma linha.
# ==========================================================
base_produtos = base.copy()
base_produtos["Produto_Individual"] = (
    base_produtos["produtos_digisac"]
    .astype(str)
    .str.split(";")
)

base_produtos = base_produtos.explode("Produto_Individual").copy()
base_produtos["Produto_Individual"] = base_produtos["Produto_Individual"].astype(str).str.strip()
base_produtos = base_produtos[
    (base_produtos["Produto_Individual"] != "") &
    (base_produtos["Produto_Individual"].str.lower() != "nan") &
    (base_produtos["Produto_Individual"].str.lower() != "não informado")
].copy()

base_produtos["Produto_Curto"] = base_produtos["Produto_Individual"].apply(encurtar_produto)
base_produtos["produtos_digisac_original"] = base_produtos["produtos_digisac"]
base_produtos["produtos_digisac"] = base_produtos["Produto_Individual"]

# ==========================================================
# Regra de contagem x valor quando o movimento tem mais de 1
# produto na mesma linha original (ex.: "Servidores; Web chat"):
#   - CONTAGEM: continua 1 linha por produto — são produtos
#     diferentes, então é correto contar cada um.
#   - VALOR (R$): pertence ao movimento inteiro, não a cada
#     produto individualmente. Por isso o valor fica só na
#     1ª linha de cada movimento explodido; nas demais linhas
#     do mesmo movimento o valor é zerado. Assim, somar a coluna
#     de valor aqui dá o mesmo total do Resumo Executivo (que
#     soma 1 vez por movimento) — mesmo a contagem de linhas
#     sendo maior (1 por produto).
# `explode` preserva o índice original em todas as linhas
# geradas a partir da mesma linha-mãe, então agrupar por esse
# índice identifica com exatidão quais linhas vieram do mesmo
# movimento.
# ==========================================================
_ordem_no_movimento = base_produtos.groupby(level=0).cumcount()
base_produtos.loc[_ordem_no_movimento > 0, "total_up_RS"] = 0
base_produtos.loc[_ordem_no_movimento > 0, "total_down_RS"] = 0

# qtd_servicos também passa a ser 1 por linha: antes repetia a
# contagem do pacote inteiro (ex.: "2") em cada uma das linhas
# explodidas, o que inflava a soma. Agora cada linha representa
# exatamente 1 produto/serviço.
base_produtos["qtd_servicos"] = 1

# ==========================================================
# De-para de Categorias — embutido no código.
# Para adicionar ou mover um produto, edite este dicionário.
# Produtos não listados ficam como "Outros".
# ==========================================================
MAPA_CATEGORIAS = {
    # ── Conexão ──────────────────────────────────────────
    "Conexão Instagram Direct":               "Conexão",
    "Cancelamento da conexão WABA":           "Conexão",
    "Conexão adicional de WhatsApp":          "Conexão",
    "Conexão Facebook Messenger":             "Conexão",
    "Web chat":                               "Conexão",
    "Conexão adicional de WhatsApp - WABA":   "Conexão",
    "Primeira ativação de número oficial":    "Conexão",
    "Migração para número oficial":           "Conexão",
    "E-mail adicional":                       "Conexão",
    "Reclame Aqui":                           "Conexão",
    # ── IA ───────────────────────────────────────────────
    "IA Copiloto":                            "IA",
    "Texto Mágico":                           "IA",
    "Resumo Inteligente":                     "IA",
    "Transcrição de Áudio":                   "IA",
    "IA CSAT":                                "IA",
    "IA Agent":                               "IA",
    "Combo Start":                            "IA",
    "Combo de IA's":                          "IA",
    "Combo Essencial":                        "IA",
    "Combo Performance":                      "IA",
    "IA Resumo Inteligente":                  "IA",
    "IA Texto Mágico":                        "IA",
    "IA Transcrição de Áudio":                "IA",
    # ── Integração ───────────────────────────────────────
    "Integração":                             "Integração",
    "Integração de prateleira":               "Integração",
    # ── Módulos ──────────────────────────────────────────
    "Chat interno":                           "Módulos",
    "Funil de vendas":                        "Módulos",
    "Distribuição automática":                "Módulos",
    # ── Plano ────────────────────────────────────────────
    "Plano Acelera para Plano Base":          "Plano",
    "Plano Base para Plano Acelera":          "Plano",
    # ── Robô ─────────────────────────────────────────────
    "Robô personalizado":                     "Robô",
    # ── Servidor ─────────────────────────────────────────
    "Servidores":                             "Servidor",
    # ── SMS ──────────────────────────────────────────────
    "Pacotes SMS Disparo em Massa":           "SMS",
    # ── URL ──────────────────────────────────────────────
    "URL":                                    "URL",
    # ── Usuários ─────────────────────────────────────────
    "Usuário adicional":                      "Usuários",
}
mapa_cat = MAPA_CATEGORIAS
print(f"✅ {len(mapa_cat)} produtos mapeados para categorias (embutido no código).")

# Aplica o mapa: produto_individual → categoria
base_produtos["Categoria"] = (
    base_produtos["produtos_digisac"]
    .astype(str)
    .str.strip()
    .map(mapa_cat)
    .fillna("Outros")
)

def gerar_analitico_auditoria(df_origem, status_consideracao, motivo_consideracao):
    """
    Gera um analítico no mesmo formato da base_analitica,
    inclusive explodindo produtos e mantendo a regra de valor apenas na 1ª linha
    de cada movimento.
    """
    if df_origem is None or len(df_origem) == 0:
        return pd.DataFrame(columns=list(base_analitica.columns) if "base_analitica" in globals() else [])

    tmp = df_origem.copy()

    # Garante colunas mínimas
    for c in ["proprietario","tema","tipo","produtos_digisac","tipo_servico","motivo_downgrade","descricao","subdominio","card_financeiro"]:
        if c not in tmp.columns:
            tmp[c] = "Não informado" if c != "card_financeiro" else ""
        tmp[c] = tmp[c].apply(normalizar) if c != "card_financeiro" else tmp[c].fillna("").astype(str).str.strip()

    tmp["dia"] = pd.to_datetime(tmp["dia"], errors="coerce", dayfirst=True)
    tmp = tmp[~tmp["dia"].isna()].copy()

    tmp["produtos_digisac"] = tmp["produtos_digisac"].apply(padronizar_produtos_digisac)
    tmp = tmp[~tmp["proprietario"].apply(lambda x: texto_chave(x) in EXCLUIR)].copy()

    tmp["qtd_servicos"] = tmp["produtos_digisac"].apply(
        lambda x: len([p for p in str(x).split(";") if p.strip() and p.strip().lower() not in ["não informado","nao informado","nan",""]])
    ).clip(lower=1)
    tmp["total_up_RS"]   = pd.to_numeric(tmp.get("total_up_RS", 0), errors="coerce").fillna(0)
    tmp["total_down_RS"] = pd.to_numeric(tmp.get("total_down_RS", 0), errors="coerce").fillna(0)

    tmp["Ano"]      = tmp["dia"].dt.year.astype(int)
    tmp["MesNum"]   = tmp["dia"].dt.month.astype(int)
    tmp["Mês"]      = tmp["MesNum"].map(MESES_PT)
    tmp["MesLongo"] = tmp["MesNum"].map(MESES_LONGO)
    tmp["AnoMes"]   = tmp["Ano"].astype(str)+"-"+tmp["MesNum"].astype(str).str.zfill(2)
    tmp["Periodo"]  = tmp["Mês"]+"/"+tmp["Ano"].astype(str).str[-2:]
    tmp["Data"]     = tmp["dia"].dt.strftime("%d/%m/%Y")

    if "Movimento" not in tmp.columns:
        tmp["Movimento"] = tmp.apply(classificar_movimento, axis=1)
    if "Situacao_Down" not in tmp.columns:
        tmp["Situacao_Down"] = tmp.apply(classificar_motivo_down, axis=1)

    tmp["Proprietario_Nome_Valido"] = tmp["proprietario"].apply(eh_nome_valido)
    tmp["Produto_Curto"] = tmp["produtos_digisac"].apply(encurtar_produto)
    tmp["proprietario_filtro"] = np.where(tmp["Proprietario_Nome_Valido"], tmp["proprietario"], "Usuário inexistente")

    base_tmp = tmp[[
        "Data","Ano","MesNum","Mês","MesLongo","AnoMes","Periodo",
        "proprietario","proprietario_filtro","Proprietario_Nome_Valido","subdominio",
        "tema","tipo","produtos_digisac","Produto_Curto","card_financeiro","tipo_servico",
        "motivo_downgrade","Situacao_Down","descricao",
        "qtd_servicos","total_up_RS","total_down_RS","Movimento"
    ]].copy()

    # Faixa de vida também no analítico de auditoria
    subs = set(base_tmp["subdominio"].dropna().astype(str).str.strip().unique()) - {"", "Não informado", "nan", "None"}
    mapa_tmp = buscar_criacao_url(subs) if len(subs) else {}
    base_tmp["criacao_url"] = base_tmp["subdominio"].astype(str).str.strip().map(mapa_tmp)
    base_tmp["criacao_url"] = pd.to_datetime(base_tmp["criacao_url"], errors="coerce")
    base_tmp["dia_dt"] = pd.to_datetime(base_tmp["Data"], format="%d/%m/%Y", errors="coerce")
    base_tmp["Tempo_vida_cli"] = (base_tmp["dia_dt"] - base_tmp["criacao_url"]).dt.days
    base_tmp["Faixa_Vida"] = base_tmp["Tempo_vida_cli"].apply(classificar_faixa_vida)
    base_tmp.drop(columns=["dia_dt"], inplace=True)

    ana = base_tmp.copy()
    ana["Produto_Individual"] = ana["produtos_digisac"].astype(str).str.split(";")
    ana = ana.explode("Produto_Individual").copy()
    ana["Produto_Individual"] = ana["Produto_Individual"].astype(str).str.strip()
    ana = ana[(ana["Produto_Individual"] != "") & (ana["Produto_Individual"].str.lower() != "nan") & (ana["Produto_Individual"].str.lower() != "não informado")].copy()
    ana["Produto_Curto"] = ana["Produto_Individual"].apply(encurtar_produto)
    ana["produtos_digisac_original"] = ana["produtos_digisac"]
    ana["produtos_digisac"] = ana["Produto_Individual"]

    ordem = ana.groupby(level=0).cumcount()
    ana.loc[ordem > 0, "total_up_RS"] = 0
    ana.loc[ordem > 0, "total_down_RS"] = 0
    ana["qtd_servicos"] = 1

    ana["Categoria"] = ana["produtos_digisac"].astype(str).str.strip().map(mapa_cat).fillna("Outros")
    ana["Status_Consideracao"] = status_consideracao
    ana["Motivo_Consideracao"] = motivo_consideracao

    def _exibir_prop_local(x):
        s = str(x).strip()
        if re.fullmatch(r"\d[\d.\s]*", s):
            return "Usuário não identificado"
        return x
    ana["proprietario"] = ana["proprietario"].apply(_exibir_prop_local)

    return ana

# ==========================================================
# Base auxiliar para as tabelas analíticas
# Aqui também separa os produtos pelo ";".
# O restante das colunas permanece igual.
# ==========================================================
# IMPORTANTE: a Base Analítica precisa bater com o Resumo Executivo.
# Antes ela usava `base_produtos`, que é uma base "explodida": quando um
# movimento tem mais de um produto (ex.: "Servidores; Web chat"), essa
# base transforma 1 linha em 2 (uma por produto) — e isso inflava tanto
# a contagem de registros quanto a soma de R$ na aba Analítica, sem que
# o Resumo (que conta 1 vez por movimento) refletisse o mesmo aumento.
# Por isso agora a Base Analítica usa a `base` original (1 linha por
# movimento, igual ao Resumo). A visão por produto individual continua
# existindo — e correta — só na aba Produtos (base_produtos), que é o
# lugar certo para granularidade por produto.
# A Base Analítica mostra 1 linha por produto (correto contar assim,
# já que são produtos diferentes), mas o valor R$ de cada linha já
# vem corrigido lá de cima em `base_produtos`: só a 1ª linha de cada
# movimento carrega o valor, as demais ficam com R$ 0. Resultado:
# contagem por produto, soma de R$ batendo com o Resumo Executivo.
base_analitica = base_produtos.copy()
base_analitica["Status_Consideracao"] = "Considerado"
base_analitica["Motivo_Consideracao"] = "Possui link no Card Financeiro e movimento válido"

# Analíticos de auditoria
# 1) Casos considerados: mesmo conteúdo da Base Analítica do relatório
analitico_considerados = base_analitica.copy()

# 2) Casos não considerados: sem Card Financeiro + temas fora dos movimentos válidos
analitico_sem_card = gerar_analitico_auditoria(
    df_sem_card_financeiro,
    "Não considerado",
    "Sem link no Card Financeiro"
)
analitico_movimento_invalido = gerar_analitico_auditoria(
    df_excluidos,
    "Não considerado",
    "Tema fora dos movimentos válidos"
)
analitico_nao_considerados = pd.concat(
    [analitico_sem_card, analitico_movimento_invalido],
    ignore_index=True
)

base.to_csv(ARQUIVO_BASE_CSV, sep=";", index=False, encoding="utf-8-sig")
analitico_considerados.to_csv(ARQUIVO_ANALITICO_CONSIDERADOS, sep=";", index=False, encoding="utf-8-sig")
analitico_nao_considerados.to_csv(ARQUIVO_ANALITICO_NAO_CONSIDERADOS, sep=";", index=False, encoding="utf-8-sig")
print(f"📋 Analítico de casos considerados salvo em: {ARQUIVO_ANALITICO_CONSIDERADOS}")
print(f"📋 Analítico de casos não considerados salvo em: {ARQUIVO_ANALITICO_NAO_CONSIDERADOS}")

# (O antigo "analítico de movimentos Outros" foi removido: o movimento
# "Outros" não existe mais — essas linhas agora são excluídas e ficam
# registradas em "linhas_excluidas_sem_movimento.csv".)

# Lista ordenada de categorias (para o filtro)
categorias_lista = sorted(base_produtos["Categoria"].dropna().unique().tolist())

# Faixa_Vida, Tempo_vida_cli e criacao_url já vêm da `base`, pois `base_produtos`
# foi criada a partir dela. Não fazemos novo merge aqui para evitar colunas duplicadas
# como Faixa_Vida_x/Faixa_Vida_y, que quebram os filtros do dashboard.

# Ordem natural das faixas para filtro
ORDEM_FAIXAS = [
    "1 - 0 a 30","2 - 31 a 60","3 - 61 a 90","4 - 91 a 120",
    "5 - 121 a 180","6 - 181 a 240","7 - 241 a 360",
    "8 - 361 a 540","9 - 541 a 720","10 - Acima de 721","Não calculado",
]
faixas_presentes = base["Faixa_Vida"].dropna().unique().tolist()
faixas_lista = [f for f in ORDEM_FAIXAS if f in faixas_presentes]

# Proprietário que é só número (ex.: "82573565") vira "Usuário não identificado"
# na exibição e na exportação (tabela e CSV usam estas colunas).
def exibir_proprietario(x):
    s = str(x).strip()
    if re.fullmatch(r"\d[\d.\s]*", s):
        return "Usuário não identificado"
    return x

for _df in (base, base_produtos, base_analitica, analitico_considerados, analitico_nao_considerados):
    _df["proprietario"] = _df["proprietario"].apply(exibir_proprietario)

payload = {
    "dados": base.replace({np.nan:None}).to_dict(orient="records"),
    "dados_produtos": base_produtos.replace({np.nan:None}).to_dict(orient="records"),
    "dados_analitica": base_analitica.replace({np.nan:None}).to_dict(orient="records"),
    "anos":  sorted(base["Ano"].dropna().unique().astype(int).tolist()),
    "meses": [{"num":k,"nome":v,"longo":MESES_LONGO[k]} for k,v in MESES_PT.items()],
    "produtos":       sorted(base_produtos["produtos_digisac"].dropna().unique().tolist()),
    "categorias":     categorias_lista,
    "faixas_vida":    faixas_lista,
    "proprietarios":  sorted([x for x in base["proprietario_filtro"].dropna().unique().tolist() if x!="Usuário inexistente"]),
    "movimentos":     sorted(base["Movimento"].dropna().unique().tolist()),
    "situacoes_down": sorted(base["Situacao_Down"].dropna().unique().tolist()),
}

def limpar_json_obj(obj):
    if isinstance(obj,dict): return {k:limpar_json_obj(v) for k,v in obj.items()}
    if isinstance(obj,list): return [limpar_json_obj(v) for v in obj]
    if isinstance(obj,pd.Timestamp): return obj.strftime("%d/%m/%Y")
    try:
        if pd.isna(obj): return None
    except: pass
    return obj

payload = limpar_json_obj(payload)
payload_json = json.dumps(payload, ensure_ascii=False)

# ══════════════════════════════════════════════════════════════════════
# Data/hora de geração do relatório
# ══════════════════════════════════════════════════════════════════════
FUSO_BRASILIA = timezone(timedelta(hours=-3))
dt_atualizacao = datetime.now(FUSO_BRASILIA).strftime("%d/%m/%Y às %H:%M")

# ══════════════════════════════════════════════════════════════════════
# HTML
# ══════════════════════════════════════════════════════════════════════
html = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Relatório Executivo Up x Down</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
:root{
  --bg:#F0F4F8;--card:#fff;--ink:#061736;--muted:#667085;--line:#dce8f3;
  --cyan:#16b8cf;--cyan2:#ddf7fb;--blue:#04357c;--blue2:#0050bf;
  --red:#c43444;--green:#059669;--amber:#D97706;
  --shadow:0 2px 8px rgba(10,38,75,.07);--radius:16px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);font-family:'Segoe UI',Arial,sans-serif;color:var(--ink)}
.app{min-height:100vh}
.side{display:none}
.nav{width:54px;height:54px;border-radius:18px;display:flex;align-items:center;justify-content:center;color:#8aa1b7;cursor:pointer;transition:.15s}
.nav:hover{background:#f0f9fb;color:var(--cyan)}
.nav.active{background:var(--cyan2);color:var(--cyan)}
.nav svg{width:22px;height:22px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
.main{width:100%;padding:0 28px 28px}
.page{display:block;scroll-margin-top:18px}
#page-home{margin-left:-28px;margin-right:-28px}
#page-exec{padding-top:28px}

/* COVER */
.cover{min-height:260px;background:linear-gradient(135deg,#0F2444 0%,#1E3A5F 55%,#0EA5E9 130%);padding:28px 48px;color:#fff;display:flex;flex-direction:column;position:relative;overflow:hidden}
.cover::before{content:"";position:absolute;width:700px;height:700px;border-radius:50%;border:1px solid rgba(255,255,255,.06);right:-180px;top:-200px}
.cover-brand{font-size:15px;font-weight:700;opacity:.7;letter-spacing:.5px;display:flex;align-items:center;gap:8px}
.cover-brand-dot{width:8px;height:8px;background:var(--cyan);border-radius:50%}
.cover h1{font-size:40px;line-height:1.05;font-weight:200;letter-spacing:-1px;margin-top:20px}
.cover h1 b{font-weight:800;color:var(--cyan)}
.cover-sub{margin-top:14px;max-width:560px;font-size:14px;line-height:1.5;color:rgba(255,255,255,.65)}
.cover-update{font-size:15px;color:#ffffff;letter-spacing:.3px;display:flex;align-items:center;gap:8px}
.cover-update::before{content:"";display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--cyan)}
.home-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:16px;margin-top:52px;max-width:920px}
.home-card{background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);border-radius:18px;padding:26px 20px;cursor:pointer;transition:.2s}
.home-card:hover{background:rgba(255,255,255,.13);transform:translateY(-4px)}
.home-card .ico{width:46px;height:46px;background:rgba(14,165,233,.22);border-radius:12px;display:flex;align-items:center;justify-content:center;margin-bottom:16px}
.home-card .ico svg{width:22px;height:22px;stroke:#7DD3FC;fill:none;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.home-card h2{font-size:14px;font-weight:700;color:#fff;line-height:1.3}
.home-card p{font-size:12px;color:rgba(255,255,255,.5);margin-top:4px}
.home-foot{font-size:24px;font-weight:800;letter-spacing:-1px;opacity:.18;margin-top:auto;padding-top:60px}

/* HERO */
.hero{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:18px 22px;margin-bottom:20px;box-shadow:var(--shadow)}
.hero-top{display:flex;align-items:center;justify-content:space-between;gap:16px}
.hero-title{display:flex;align-items:center;gap:12px}
.hero-title .ico{width:44px;height:44px;border-radius:50%;background:var(--cyan2);color:var(--cyan);display:flex;align-items:center;justify-content:center}
.hero-title .ico svg{width:22px;height:22px;stroke:currentColor;fill:none;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.hero h1{font-size:24px;font-weight:800;margin:0}
.hero-sub{font-size:13px;color:var(--muted);margin-top:2px}
.reset{border:0;background:#eef9fb;color:var(--cyan);border-radius:10px;padding:10px 14px;font-weight:800;cursor:pointer;display:flex;align-items:center;gap:7px;font-size:13px;transition:.15s}
.reset:hover{background:var(--cyan2)}
.reset svg{width:14px;height:14px;stroke:currentColor;fill:none;stroke-width:2.2;stroke-linecap:round;stroke-linejoin:round}

/* FILTROS MULTI-SELECT */
.filters{display:flex;flex-wrap:wrap;gap:10px;margin-top:12px}
.filter{position:relative;width:180px}
.filter-label{font-size:10px;font-weight:800;color:#344054;margin-bottom:4px;text-transform:uppercase;letter-spacing:.5px}
.multi-btn{width:100%;min-height:32px;border:1.5px solid #cbd8e6;background:#fff;border-radius:8px;padding:5px 28px 5px 10px;text-align:left;font-size:12px;color:#101828;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;position:relative;transition:.15s}
.multi-btn:hover{border-color:#94A3B8}
.multi-btn::after{content:"▾";position:absolute;right:11px;top:50%;transform:translateY(-50%);color:#667085;font-size:11px}
.multi.open .multi-btn{border-color:var(--cyan);box-shadow:0 0 0 3px rgba(22,184,207,.12)}
.multi-panel{display:none;position:absolute;z-index:200;top:calc(100% + 5px);left:0;right:0;min-width:220px;background:#fff;border:1.5px solid #cbd8e6;border-radius:14px;box-shadow:0 12px 32px rgba(7,31,64,.15);padding:10px;max-height:280px;overflow:auto}
.multi.open .multi-panel{display:block}
.multi-actions{display:flex;gap:6px;margin-bottom:8px}
.mini-btn{border:1px solid var(--line);background:#f8fbfd;border-radius:7px;padding:5px 9px;font-size:11px;font-weight:800;cursor:pointer;color:#0f365e;transition:.12s}
.mini-btn:hover{background:var(--line)}
.search-inp{width:100%;height:32px;border:1px solid var(--line);border-radius:8px;margin-bottom:8px;padding:0 9px;font-size:12px;outline:none}
.search-inp:focus{border-color:var(--cyan)}
.check{display:flex;gap:8px;align-items:center;padding:7px 4px;font-size:13px;color:#101828;cursor:pointer;border-radius:7px}
.check:hover{background:#f8fbfd}
.check input{accent-color:var(--cyan);width:14px;height:14px;cursor:pointer;flex-shrink:0}

/* KPI TOOLTIP */
.kpi-wrap{position:relative}
.kpi-info{position:absolute;top:14px;right:14px;width:16px;height:16px;border-radius:50%;background:#E0F2FE;color:#0EA5E9;font-size:10px;font-weight:800;display:flex;align-items:center;justify-content:center;cursor:default;flex-shrink:0;line-height:1}
.kpi-info:hover .kpi-tip{display:block}
.kpi-tip{display:none;position:absolute;top:22px;right:0;width:220px;background:#1E3A5F;color:#fff;font-size:11px;line-height:1.6;padding:10px 12px;border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.2);z-index:300;font-weight:400}
.kpi-tip::before{content:"";position:absolute;top:-5px;right:8px;width:10px;height:10px;background:#1E3A5F;transform:rotate(45deg);border-radius:2px}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px}
.kpi{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:18px 20px;box-shadow:var(--shadow);position:relative}
.kpi .label{font-size:11px;color:#667085;font-weight:700;text-transform:uppercase;letter-spacing:.5px}
.kpi .value{font-size:28px;font-weight:800;margin:8px 0 4px;color:var(--blue);letter-spacing:-1px}
.kpi .sub{font-size:12px;color:#667085}
.kpi .kpi-valor{font-size:14px;font-weight:700;color:#334155;margin-bottom:4px}
.kpi.kpi-up{border-top:3px solid #16b8cf}
.kpi.kpi-dn{border-top:3px solid #04357c}
.kpi.kpi-saldo{border-top:3px solid #059669}
.badge{display:inline-flex;align-items:center;font-size:11px;font-weight:800;padding:3px 9px;border-radius:999px}
.badge-up{background:#DCFCE7;color:#15803D}
.badge-dn{background:#FEE2E2;color:#B91C1C}

/* GRIDS E CARDS */
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.grid3{display:grid;grid-template-columns:1.45fr 1fr;gap:16px;margin-bottom:16px}
.card{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:22px 24px;box-shadow:var(--shadow);margin-bottom:16px}
.card h3{font-size:16px;font-weight:800;margin-bottom:16px}
.card-hd{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px}
.card-hd h3{margin:0}

/* INSIGHT CARDS */
.ic-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.exec-insights-card{margin-bottom:16px;padding:16px 18px}
.exec-insights-card .ic{min-height:92px}
.exec-chart-card{margin-bottom:16px}
.ic{border-radius:10px;padding:15px 16px;border-left:3px solid transparent}
.ic.pos{background:#ECFDF5;border-left-color:#059669}
.ic.warn{background:#FEF2F2;border-left-color:#DC2626}
.ic.info{background:#E0F2FE;border-left-color:#0EA5E9}
.ic.amb{background:#FFFBEB;border-left-color:#D97706}
.ic .il{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px}
.ic .iv{font-size:20px;font-weight:800;line-height:1.1;margin-bottom:4px}
.ic .id{font-size:12px;line-height:1.5}
.ic.pos .il,.ic.pos .iv,.ic.pos .id{color:#065F46}
.ic.warn .il,.ic.warn .iv,.ic.warn .id{color:#991B1B}
.ic.info .il,.ic.info .iv,.ic.info .id{color:#0C4A6E}
.ic.amb .il,.ic.amb .iv,.ic.amb .id{color:#78350F}
.ic-note{margin-top:12px;padding:13px 15px;background:#F8FAFC;border-radius:10px;font-size:13px;color:#475569;line-height:1.65}
.ic-note strong{color:var(--ink);font-weight:700}

/* BARRAS HORIZONTAIS CSS */
.bar-row{display:flex;align-items:center;gap:10px;margin-bottom:9px;padding:2px 0}
.bar-lbl{font-size:12px;color:#475569;width:200px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bar-track{flex:1;height:10px;background:#F0F4F8;border-radius:999px;overflow:hidden}
.bar-fill{height:100%;border-radius:999px;transition:width .4s ease}
.bar-cnt{font-size:12px;font-weight:700;color:var(--ink);min-width:80px;text-align:right;flex-shrink:0}

/* NOTE BAR */
.note-bar{background:#E0F2FE;border-radius:10px;padding:10px 14px;font-size:12px;color:#0C4A6E;margin-top:14px}

/* FAIXA DE VIDA */
.fv-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px}
.fv-kpi{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:16px 18px;position:relative}
.fv-kpi.kpi-up{border-top:3px solid #16b8cf}
.fv-kpi.kpi-dn{border-top:3px solid #04357c}
.fv-kpi.kpi-sal{border-top:3px solid #059669}
.fv-kpi.kpi-mig{border-top:3px solid #60A5FA}
.fv-kpi.kpi-tot{border-top:3px solid #94A3B8}
.fv-kpi-lbl{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;color:#667085;margin-bottom:6px}
.fv-kpi-val{font-size:26px;font-weight:800;letter-spacing:-1px;color:#061736;margin-bottom:3px}
.fv-kpi-brl{font-size:13px;font-weight:700;color:#334155;margin-bottom:3px}
.fv-kpi-sub{font-size:11px;color:#667085}
.fv-insight{border-radius:10px;padding:14px 16px}
.fv-insight.up{background:#ECFDF5;border-left:3px solid #059669}
.fv-insight.dn{background:#FEF2F2;border-left:3px solid #DC2626}
.fv-insight.alert{background:#FFFBEB;border-left:3px solid #D97706}
.fv-lbl{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px}
.fv-val{font-size:18px;font-weight:800;line-height:1.2;margin-bottom:3px}
.fv-sub{font-size:12px;line-height:1.5}
.fv-insight.up .fv-lbl,.fv-insight.up .fv-val,.fv-insight.up .fv-sub{color:#065F46}
.fv-insight.dn .fv-lbl,.fv-insight.dn .fv-val,.fv-insight.dn .fv-sub{color:#991B1B}
.fv-insight.alert .fv-lbl,.fv-insight.alert .fv-val,.fv-insight.alert .fv-sub{color:#78350F}
.fv-insight-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:20px}
.fv-bars{display:grid;grid-template-columns:1fr 1fr;gap:0 40px}
.fv-bar-row{display:flex;align-items:center;gap:8px;margin-bottom:6px;cursor:pointer;border-radius:6px;padding:3px 4px;transition:.12s}
.fv-bar-row:hover{background:#F0F9FB}
.fv-bar-row:hover .fv-bar-lbl{color:#0EA5E9;font-weight:700}
.fv-bar-lbl{font-size:11px;color:#475569;width:130px;flex-shrink:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;transition:.12s}
.fv-bar-track{flex:1;height:8px;background:#F0F4F8;border-radius:999px;overflow:hidden}
.fv-bar-fill-up{height:100%;border-radius:999px;background:#16b8cf;transition:width .3s}
.fv-bar-fill-dn{height:100%;border-radius:999px;background:#04357c;transition:width .3s}
.fv-bar-cnt{font-size:11px;font-weight:700;color:var(--ink);min-width:32px;text-align:right;flex-shrink:0}
.fv-bar-hint{font-size:10px;color:#94A3B8;margin-left:2px}
.fv-section-title{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;color:#94A3B8;margin-bottom:10px}
.fv-side-insights{background:#F8FAFC;border:1px solid var(--line);border-radius:14px;padding:16px;box-shadow:var(--shadow)}
.fv-side-title{font-size:12px;font-weight:900;text-transform:uppercase;letter-spacing:.5px;color:#0C4A6E;margin-bottom:12px}
.fv-insight-line{background:#fff;border:1px solid #E2E8F0;border-radius:12px;padding:11px 12px;margin-bottom:10px}
.fv-insights-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:10px}
.fv-insights-grid .fv-insight-line{margin-bottom:0}
.fv-insight-line span{display:block;font-size:10px;font-weight:900;text-transform:uppercase;letter-spacing:.45px;color:#64748B;margin-bottom:4px}
.fv-insight-line strong{display:block;font-size:18px;line-height:1.15;color:#061736;margin-bottom:3px}
.fv-insight-line small{display:block;font-size:11px;color:#667085;line-height:1.35}
.fv-mini-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-top:12px}
.fv-mini-grid div{background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:10px 8px;text-align:center}
.fv-mini-grid b{display:block;font-size:18px;color:#061736;line-height:1}
.fv-mini-grid span{display:block;font-size:10px;color:#64748B;margin-top:4px;text-transform:uppercase;font-weight:800}

/* Drill-down faixa */
.fv-drill-header{display:flex;align-items:center;gap:10px;margin-bottom:14px;padding:9px 14px;background:#E0F2FE;border-radius:10px;font-size:13px}
.fv-drill-back{border:none;background:none;color:#0C4A6E;font-weight:700;cursor:pointer;font-size:13px;padding:0;text-decoration:underline}
.fv-drill-tag{font-size:11px;background:#BFDBFE;color:#1E40AF;padding:2px 8px;border-radius:999px}
.fv-drill-table{width:100%;border-collapse:collapse;font-size:12px;margin-top:4px}
.fv-drill-table th{background:#F8FAFC;color:#667085;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.4px;padding:7px 10px;border-bottom:1px solid var(--line);text-align:left}
.fv-drill-table th.r{text-align:right}
.fv-drill-table td{padding:6px 10px;border-bottom:1px solid #F1F5F9;color:#475569;vertical-align:middle}
.fv-drill-table td.r{text-align:right}
.fv-drill-table tbody tr:hover td{background:#FAFBFD}

/* TABELA MENSAL */
.mt{width:100%;border-collapse:collapse;font-size:13px;min-width:520px}
.mt th{padding:10px 14px;background:#F8FAFC;color:#667085;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.4px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap}
.mt th:first-child{text-align:left}
.mt td{padding:10px 14px;text-align:right;border-bottom:1px solid #F1F5F9;color:#475569}
.mt td:first-child{text-align:left;font-weight:700;color:var(--ink)}
.mt .tr-tot td{background:#F8FAFC;font-weight:800;color:var(--ink);border-top:2px solid var(--line)}
.pill{display:inline-flex;align-items:center;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:800}
.p-up{background:#DCFCE7;color:#15803D}.p-dn{background:#FEE2E2;color:#B91C1C}
.p-mig{background:#DBEAFE;color:#1D4ED8}.p-tst{background:#E2E8F0;color:#334155}.p-oth{background:#F3F4F6;color:#374151}
.p-url{background:#FEF3C7;color:#92400E}.p-dsc{background:#EDE9FE;color:#4C1D95}
.p-rob{background:#D1FAE5;color:#065F46}.p-tst{background:#E0F2FE;color:#0C4A6E}

/* DATA TABLES */
.tbar{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.tbar h3{font-size:15px;font-weight:800;margin:0}
.btn{border:1px solid var(--line);background:#fff;border-radius:10px;padding:9px 14px;font-size:12px;font-weight:700;cursor:pointer;color:#475569;display:inline-flex;align-items:center;gap:6px;transition:.15s}
.btn:hover{background:#F8FAFC}
.btn svg{width:14px;height:14px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
.btn.pri{background:var(--cyan);border-color:var(--cyan);color:#fff}
.btn.pri:hover{background:#0ea0b5}
.table-wrap{overflow:auto;border-radius:10px;border:1px solid var(--line);max-height:500px}
table.dt{width:100%;border-collapse:collapse;font-size:12px;min-width:900px}
table.dt th{position:sticky;top:0;background:#F8FAFC;color:#667085;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.4px;padding:7px 10px;border-bottom:2px solid var(--line);text-align:left;z-index:2;white-space:nowrap}
table.dt td{padding:6px 10px;border-bottom:1px solid #F1F5F9;color:#475569;vertical-align:top}
table.dt tbody tr:hover td{background:#FAFBFD}
table.dt .r{text-align:right}
.pager{display:flex;align-items:center;justify-content:flex-end;gap:10px;margin-top:12px;font-size:13px;color:#667085}
.csv-box{display:none;width:100%;height:160px;margin-top:10px;border:1px solid var(--line);border-radius:10px;padding:10px;font-family:Consolas,monospace;font-size:12px;color:var(--ink)}

@media (max-width:900px){
  .main{padding:0 14px 20px}
  #page-home{margin-left:-14px;margin-right:-14px}
  .cover{min-height:220px;padding:24px 22px}
  .cover h1{font-size:32px;margin-top:16px}
  .cover-sub{font-size:13px}
  .kpis{grid-template-columns:repeat(2,1fr)}
  .grid2,.grid3,.ic-grid{grid-template-columns:1fr}
}
@media (max-width:560px){
  .cover h1{font-size:28px}
  .kpis{grid-template-columns:1fr}
  .hero-top{align-items:flex-start;flex-direction:column}
  .filter{width:100%}
}
</style>
</head>
<body>
<div class="app" id="app">



<main class="main">

<!-- CAPA -->
<section id="page-home" class="page">
<div class="cover">
  <div class="cover-brand"><div class="cover-brand-dot"></div>digisac</div>
  <h1>Relatório<br><b>Up × Down</b></h1>
  <p class="cover-sub">Painel executivo para acompanhar expansão, redução, saldo operacional, impacto financeiro, produtos críticos e motivos de downgrade.</p>
  <div style="display:flex;align-items:flex-end;justify-content:space-between;margin-top:auto;padding-top:24px">
    <div class="home-foot">ikatec</div>
    <div class="cover-update">Atualizado: __DT_ATUALIZACAO__</div>
  </div>
</div>
</section>

<!-- EXECUTIVO -->
<section id="page-exec" class="page">
  <div class="hero">
    <div class="hero-top">
      <div class="hero-title">
        <div class="ico"><svg viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg></div>
        <div><h1>Resumo Executivo</h1><div class="hero-sub">Visão estratégica: volume, valor, saldo e tendência do período</div></div>
      </div>
      <button class="reset" onclick="resetFiltros()">
        <svg viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3"/></svg>Limpar filtros
      </button>
    </div>
    <div id="filters-exec" class="filters"></div>
  </div>
  <div class="kpis">
    <div class="kpi kpi-up">
      <div class="kpi-info">i<div class="kpi-tip">Conta todos os movimentos classificados como <strong>Upgrade</strong> no período — inclui upgrades diretos e Criação de URL. O valor é a soma financeira dos upgrades (coluna total_up_RS).</div></div>
      <div class="label">Upgrade</div>
      <div class="value" id="kUp">0</div>
      <div class="sub kpi-valor" id="kValorUp">R$ 0</div>
      <div class="sub"><span class="badge badge-up" id="kUpPct">0%</span> do volume</div>
    </div>
    <div class="kpi kpi-dn">
      <div class="kpi-info">i<div class="kpi-tip">Conta todos os movimentos classificados como <strong>Downgrade</strong> no período. O valor é a soma financeira dos downgrades (coluna total_down_RS).</div></div>
      <div class="label">Downgrade</div>
      <div class="value" id="kDown">0</div>
      <div class="sub kpi-valor" id="kValorDown">R$ 0</div>
      <div class="sub"><span class="badge badge-dn" id="kDownPct">0%</span> do volume</div>
    </div>
    <div class="kpi kpi-saldo">
      <div class="kpi-info">i<div class="kpi-tip"><strong>Saldo = Upgrade − Downgrade</strong><br>Qtd: diferença entre a quantidade de upgrades e downgrades.<br>Valor: diferença financeira (Valor Up − Valor Down).<br>Saldo positivo indica que o período teve mais expansão do que redução.</div></div>
      <div class="label">Saldo Up − Down</div>
      <div class="value" id="kSaldo">0</div>
      <div class="sub kpi-valor" id="kSaldoValor">R$ 0</div>
      <div class="sub" id="kSaldoTxt">—</div>
    </div>
    <div class="kpi">
      <div class="kpi-info">i<div class="kpi-tip">Agrupa os movimentos que <strong>não são Upgrade nem Downgrade</strong>: Criação de URL, Migração, Teste, Robô e Alteração de Plano. Não entram no cálculo de saldo.</div></div>
      <div class="label">Outros Produtos</div>
      <div class="value" id="kOutros">0</div>
      <div class="sub" id="kOutrosTxt">—</div>
    </div>
    <div class="kpi">
      <div class="kpi-info">i<div class="kpi-tip"><strong>Total de movimentos</strong> no período filtrado (uma linha por produto: Upgrade + Downgrade + Criação de URL + Migração + Teste + Robô + Alteração de Plano).<br>O valor é a soma financeira geral (Up + Down), sem duplicar quando há mais de um produto.</div></div>
      <div class="label">Movimentos</div>
      <div class="value" id="kTotal">0</div>
      <div class="sub kpi-valor" id="kTotalValor">R$ 0</div>
      <div class="sub" id="kPeriodo">—</div>
    </div>
  </div>
  <div class="card exec-insights-card">
    <h3>Insights do período</h3>
    <div id="execInsights"></div>
  </div>
  <div class="card exec-chart-card">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
      <h3 style="margin:0" id="chartExecTitle">Volume mensal por tipo</h3>
      <div style="display:flex;background:#F0F4F8;border-radius:10px;padding:3px;gap:2px">
        <button id="btnQtd" onclick="setChartMode('qtd')" style="border:none;border-radius:8px;padding:7px 14px;font-size:12px;font-weight:800;cursor:pointer;background:var(--cyan);color:#fff;transition:.15s">Quantidade</button>
        <button id="btnVal" onclick="setChartMode('val')" style="border:none;border-radius:8px;padding:7px 14px;font-size:12px;font-weight:800;cursor:pointer;background:transparent;color:#667085;transition:.15s">Valor R$</button>
      </div>
    </div>
    <div style="position:relative;height:360px"><canvas id="chartExec"></canvas></div>
  </div>
  <!-- Card de Faixa de Vida — linha inteira -->
  <div class="card" style="margin-bottom:16px">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
      <h3 style="margin:0">Movimentações por ciclo de vida do cliente</h3>
      <span style="font-size:12px;color:#94A3B8">faixas, movimentações e produtos</span>
    </div>
    <div id="execFaixaVida"></div>
  </div>
  <div class="grid2" style="align-items:stretch;grid-template-columns:1fr">
    <div class="card" style="margin-bottom:0">
      <h3>Tabela mensal — Quantidade</h3>
      <div class="table-wrap"><div id="tblMensalQtd"></div></div>
    </div>
    <div class="card" style="margin-bottom:0">
      <h3>Tabela mensal — Valor (R$)</h3>
      <div class="table-wrap"><div id="tblMensalVal"></div></div>
    </div>
  </div>
</section>

<!-- PRODUTOS -->
<section id="page-prod" class="page">
  <div class="hero">
    <div class="hero-top">
      <div class="hero-title">
        <div class="ico"><svg viewBox="0 0 24 24"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/></svg></div>
        <div><h1 id="prodPageTitle">Produtos</h1><div class="hero-sub" id="prodPageSub">Expansão, redução e impacto financeiro por categoria</div></div>
      </div>
      <button class="reset" onclick="resetFiltros()"><svg viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3"/></svg>Limpar filtros</button>
    </div>
    <div id="filters-prod" class="filters"></div>
  </div>
  <!-- Breadcrumb de drill-down -->
  <div id="prod-breadcrumb" style="display:none;align-items:center;gap:8px;margin-bottom:14px;font-size:13px;color:#667085;background:#E0F2FE;border-radius:10px;padding:10px 16px;">
    <svg viewBox="0 0 24 24" style="width:14px;height:14px;stroke:#0EA5E9;fill:none;stroke-width:2.5;stroke-linecap:round;flex-shrink:0"><polyline points="15 18 9 12 15 6"/></svg>
    <button onclick="prodDrillBack()" style="border:none;background:none;color:#0C4A6E;font-weight:700;cursor:pointer;font-size:13px;padding:0;text-decoration:underline">Voltar para categorias</button>
    <span style="color:#94A3B8">›</span>
    <span style="font-weight:700;color:#0C4A6E" id="prod-breadcrumb-cat"></span>
    <span style="font-size:11px;background:#BFDBFE;color:#1E40AF;padding:2px 8px;border-radius:999px;margin-left:4px">detalhe</span>
  </div>
  <!-- Dica de interatividade (só aparece na visão de categorias) -->
  <div id="prod-drill-hint" style="display:flex;align-items:center;gap:7px;margin-bottom:14px;font-size:12px;color:#667085;background:#F8FAFC;border-radius:10px;padding:9px 14px;border:1px solid #E2E8F0">
    <svg viewBox="0 0 24 24" style="width:13px;height:13px;stroke:#94A3B8;fill:none;stroke-width:2;flex-shrink:0"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
    Clique em uma barra para ver os produtos daquela categoria
  </div>
  <div class="grid2">
    <div class="card"><div class="card-hd"><h3 id="hPupQ">Upgrade — Quantidade</h3></div><div id="bPupQ"></div></div>
    <div class="card"><div class="card-hd"><h3 id="hPdnQ">Downgrade — Quantidade</h3></div><div id="bPdnQ"></div></div>
  </div>
  <div class="grid2">
    <div class="card"><div class="card-hd"><h3 id="hPupV">Upgrade — Valor</h3></div><div id="bPupV"></div></div>
    <div class="card"><div class="card-hd"><h3 id="hPdnV">Downgrade — Valor</h3></div><div id="bPdnV"></div></div>
  </div>
  <div class="card">
    <div class="tbar"><h3 id="tblProdTitle">Resumo por Categoria</h3><button class="btn pri" onclick="expProd()"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>Exportar</button></div>
    <div class="table-wrap" id="tblProd"></div><div class="pager" id="pgProd"></div>
  </div>
</section>

<!-- MOTIVOS DOWN -->
<section id="page-down" class="page">
  <div class="hero">
    <div class="hero-top">
      <div class="hero-title">
        <div class="ico"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="8 12 12 16 16 12"/><line x1="12" y1="8" x2="12" y2="16"/></svg></div>
        <div><h1>Motivos de Downgrade</h1><div class="hero-sub">Situações, causas e padrões nos processos de redução</div></div>
      </div>
      <button class="reset" onclick="resetFiltros()"><svg viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3"/></svg>Limpar filtros</button>
    </div>
    <div id="filters-down" class="filters"></div>
  </div>
  <div class="grid3">
    <div class="card"><div class="card-hd"><h3>Distribuição de situações</h3></div><div id="barsSit"></div></div>
    <div class="card"><div class="card-hd"><h3>Análise de motivos</h3></div><div id="downIns"></div></div>
  </div>
  <div class="card">
    <div class="tbar"><h3>Base de Downgrade — motivos e observações</h3><button class="btn pri" onclick="expDown()"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>Exportar</button></div>
    <div class="table-wrap" id="tblDown"></div><div class="pager" id="pgDown"></div>
  </div>
</section>

<!-- PROPRIETÁRIOS -->
<section id="page-prop" class="page">
  <div class="hero">
    <div class="hero-top">
      <div class="hero-title">
        <div class="ico"><svg viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg></div>
        <div><h1>Proprietários</h1><div class="hero-sub">Somente proprietários com nome identificado</div></div>
      </div>
      <button class="reset" onclick="resetFiltros()"><svg viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3"/></svg>Limpar filtros</button>
    </div>
    <div id="filters-prop" class="filters"></div>
  </div>
  <div class="grid2">
    <div class="card"><div class="card-hd"><h3>Upgrade — Quantidade</h3></div><div id="bOupQ"></div></div>
    <div class="card"><div class="card-hd"><h3>Downgrade — Quantidade</h3></div><div id="bOdnQ"></div></div>
  </div>
  <div class="grid2">
    <div class="card"><div class="card-hd"><h3>Upgrade — Valor</h3></div><div id="bOupV"></div></div>
    <div class="card"><div class="card-hd"><h3>Downgrade — Valor</h3></div><div id="bOdnV"></div></div>
  </div>
  <div class="card">
    <div class="tbar"><h3>Resumo por Proprietário</h3><button class="btn pri" onclick="expProp()"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>Exportar</button></div>
    <div class="table-wrap" id="tblProp"></div><div class="pager" id="pgProp"></div>
  </div>
</section>

<!-- ANALÍTICO -->
<section id="page-ana" class="page">
  <div class="hero">
    <div class="hero-top">
      <div class="hero-title">
        <div class="ico"><svg viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg></div>
        <div><h1>Base Analítica</h1><div class="hero-sub">Todos os registros com filtros e exportação</div></div>
      </div>
      <button class="reset" onclick="resetFiltros()"><svg viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-3"/></svg>Limpar filtros</button>
    </div>
    <div id="filters-ana" class="filters"></div>
  </div>
  <div class="card">
    <div class="tbar">
      <h3 id="anaCnt">0 registros</h3>
      <div style="display:flex;gap:8px">
        <button class="btn pri" onclick="expAna()"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>Exportar CSV</button>
        <button class="btn" onclick="copyCSV()"><svg viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>Copiar</button>
        <button class="btn" onclick="toggleCSV()"><svg viewBox="0 0 24 24"><polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/></svg>Ver CSV</button>
      </div>
    </div>
    <textarea id="csvBox" class="csv-box"></textarea>
    <div class="table-wrap" id="tblAna"></div><div class="pager" id="pgAna"></div>
  </div>
</section>

</main>
</div>

<script>
const DATA = __PAYLOAD_JSON__;
const DADOS = DATA.dados;
const DADOS_PRODUTOS = DATA.dados_produtos || DATA.dados;
const DADOS_ANALITICA = DATA.dados_analitica || DATA.dados_produtos || DATA.dados;
const PS = 20;
let CP = {prod:1,prop:1,down:1,ana:1};
let current = 'report';

let SEL = {
  exec:{ano:[],mes:[],movimento:[],proprietario:[],faixa_vida:[]},
  prod:{ano:[],mes:[],movimento:[],categoria:[],faixa_vida:[]},
  down:{ano:[],mes:[],produto:[],situacao:[],faixa_vida:[]},
  prop:{ano:[],mes:[],movimento:[],proprietario:[],faixa_vida:[]},
  ana: {ano:[],mes:[],movimento:[],produto:[],proprietario:[],faixa_vida:[]}
};

// Estado do drill-down de Produtos
let PROD_DRILL = null; // null = visão de categorias; string = categoria selecionada

function id(x){return document.getElementById(x)}
function esc(s){return String(s??'').replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]))}
// Proprietário só com números (ex.: "82573565") vira "Usuário não identificado".
function propNome(p){
  let s=String(p??'').trim();
  return /^\d[\d.\s]*$/.test(s) ? 'Usuário não identificado' : s;
}
function fmt(v){return new Intl.NumberFormat('pt-BR').format(Math.round(Number(v||0)))}
function brl(v){return new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL',maximumFractionDigits:0}).format(Number(v||0))}
function pct(v){return Number(v||0).toFixed(1).replace('.',',')+' %'}
function sum(rs,c){return rs.reduce((a,r)=>a+(Number(r[c])||0),0)}
function sumM(rs,m){return rs.filter(r=>r.Movimento===m).length}
function grp(rs,f,vf,mov){
  let m={};
  rs.forEach(r=>{
    if(mov && r.Movimento!==mov) return;
    let k=r[f]||'Não informado';
    m[k]=(m[k]||0)+(Number(r[vf])||0);
  });
  return Object.entries(m).map(([key,value])=>({key,value})).sort((a,b)=>b.value-a.value);
}
function grpCnt(rs,f,mov){
  let m={};
  rs.forEach(r=>{
    if(mov && r.Movimento!==mov) return;
    let k=r[f]||'Não informado';
    m[k]=(m[k]||0)+1;
  });
  return Object.entries(m).map(([key,value])=>({key,value})).sort((a,b)=>b.value-a.value);
}
function top1(a){return a&&a.length?a[0]:{key:'—',value:0}}
function periodTxt(rs){
  let ms=[...new Set(rs.map(r=>r.AnoMes))].sort();
  if(!ms.length) return '—';
  let f=rs.find(r=>r.AnoMes===ms[0]), l=rs.find(r=>r.AnoMes===ms[ms.length-1]);
  return ms.length===1 ? f.Periodo : f.Periodo+' a '+l.Periodo;
}
function months(rs){
  return [...new Set(rs.map(r=>r.AnoMes))].sort().map(am=>{
    let r=rs.find(x=>x.AnoMes===am);
    return {am, label:r.Periodo, mes:r.MesNum};
  });
}

const CR={};
function newChart(cid,cfg){
  if(CR[cid]) CR[cid].destroy();
  let c=id(cid); if(!c) return;
  CR[cid]=new Chart(c,cfg);
}

function fRows(page){
  let f=SEL[page];
  // Regra oficial do relatório:
  // - Painel Executivo, Motivos Down e Base Analítica usam a base explodida por produto.
  //   Assim, se um mesmo movimento tiver 2, 3, 4 ou mais produtos, a contagem considera
  //   cada produto separadamente.
  // - O valor financeiro NÃO duplica, porque no Python o valor fica apenas na 1ª linha
  //   do movimento explodido e fica zerado nas demais linhas do mesmo movimento.
  // Resultado: contagem do Painel bate com a Base Analítica, e o R$ continua batendo
  // com o movimento original.
  let fonte = page === 'prod' ? DADOS_PRODUTOS : (page === 'exec' || page === 'ana' || page === 'down' ? DADOS_ANALITICA : DADOS);
  return fonte.filter(r=>
    (!f.ano||!f.ano.length||f.ano.includes(String(r.Ano)))&&
    (!f.mes||!f.mes.length||f.mes.includes(String(r.MesNum)))&&
    (!f.movimento||!f.movimento.length||f.movimento.includes(r.Movimento))&&
    (!f.produto||!f.produto.length||f.produto.includes(r.produtos_digisac))&&
    (!f.categoria||!f.categoria.length||f.categoria.includes(r.Categoria))&&
    (!f.faixa_vida||!f.faixa_vida.length||f.faixa_vida.includes(r.Faixa_Vida))&&
    (!f.proprietario||!f.proprietario.length||f.proprietario.includes(r.proprietario_filtro))&&
    (!f.situacao||!f.situacao.length||f.situacao.includes(r.Situacao_Down))
  );
}

// Base original para os gráficos de VALOR da aba Produtos.
// Motivo: quando uma linha tem vários produtos separados por ";",
// o valor da linha é a soma do conjunto e não pode ser dividido por produto individual.
function fRowsProdValor(){
  let f = SEL.prod;
  return DADOS_PRODUTOS.filter(r=>{
    return (
      (!f.ano||!f.ano.length||f.ano.includes(String(r.Ano)))&&
      (!f.mes||!f.mes.length||f.mes.includes(String(r.MesNum)))&&
      (!f.movimento||!f.movimento.length||f.movimento.includes(r.Movimento))&&
      (!f.categoria||!f.categoria.length||f.categoria.includes(r.Categoria))
    );
  });
}

function makeFilter(containerId, page, key, label, values, fmtFn, searchable){
  let wrap = document.createElement('div');
  wrap.className = 'filter';
  let lbl = document.createElement('div');
  lbl.className = 'filter-label';
  lbl.textContent = label;
  let outer = document.createElement('div');
  outer.className = 'multi';
  outer.dataset.page = page;
  outer.dataset.key  = key;

  let opts = [{v:'_all_',t:'Todos'}, ...values.map(v=>({v:String(v),t:fmtFn?fmtFn(String(v)):String(v)}))];

  let btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'multi-btn';
  btn.textContent = 'Todos';
  btn.onclick = function(e){
    e.stopPropagation();
    document.querySelectorAll('.multi.open').forEach(m=>{if(m!==outer)m.classList.remove('open')});
    outer.classList.toggle('open');
  };

  let panel = document.createElement('div');
  panel.className = 'multi-panel';

  if(searchable !== false){
    let si = document.createElement('input');
    si.type='text'; si.className='search-inp'; si.placeholder='Pesquisar...';
    si.oninput = function(e){
      let q = e.target.value.toLowerCase();
      panel.querySelectorAll('.check').forEach(l=>{
        l.style.display = l.dataset.txt.includes(q) ? '' : 'none';
      });
    };
    panel.appendChild(si);
  }

  let acts = document.createElement('div');
  acts.className = 'multi-actions';
  let bAll = document.createElement('button');
  bAll.type='button'; bAll.className='mini-btn'; bAll.textContent='Todos';
  bAll.onclick = function(){
    panel.querySelectorAll('input[type=checkbox]').forEach(i=>i.checked=i.value==='_all_');
    syncMulti(outer);
  };
  let bClr = document.createElement('button');
  bClr.type='button'; bClr.className='mini-btn'; bClr.textContent='Limpar';
  bClr.onclick = function(){
    panel.querySelectorAll('input[type=checkbox]').forEach(i=>i.checked=false);
    panel.querySelector('input[value="_all_"]').checked=true;
    syncMulti(outer);
  };
  acts.appendChild(bAll); acts.appendChild(bClr); panel.appendChild(acts);

  opts.forEach(function(o){
    let lbl2 = document.createElement('label');
    lbl2.className = 'check';
    lbl2.dataset.txt = o.t.toLowerCase();
    let cb = document.createElement('input');
    cb.type='checkbox'; cb.value=o.v;
    if(o.v==='_all_') cb.checked=true;
    cb.onchange = function(){
      if(cb.value==='_all_' && cb.checked){
        panel.querySelectorAll('input[type=checkbox]').forEach(i=>{ if(i!==cb) i.checked=false; });
      } else if(cb.checked){
        let ac=panel.querySelector('input[value="_all_"]'); if(ac) ac.checked=false;
      } else {
        let any=[...panel.querySelectorAll('input[type=checkbox]')].some(i=>i.value!=='_all_'&&i.checked);
        if(!any){ let ac=panel.querySelector('input[value="_all_"]'); if(ac) ac.checked=true; }
      }
      syncMulti(outer);
    };
    let sp = document.createElement('span'); sp.textContent=o.t;
    lbl2.appendChild(cb); lbl2.appendChild(sp);
    panel.appendChild(lbl2);
  });

  outer.appendChild(btn); outer.appendChild(panel);
  wrap.appendChild(lbl); wrap.appendChild(outer);
  id(containerId).appendChild(wrap);
}

function syncMulti(outer){
  let page=outer.dataset.page, key=outer.dataset.key;
  let checked=[...outer.querySelectorAll('input[type=checkbox]:checked')].map(i=>i.value);
  let isAll = !checked.length || checked.includes('_all_');
  SEL[page][key] = isAll ? [] : checked;
  let btn=outer.querySelector('.multi-btn');
  if(isAll){ btn.textContent='Todos'; }
  else if(checked.length<=2){
    btn.textContent=checked.map(v=>{
      let el=outer.querySelector('input[value="'+CSS.escape(v)+'"]');
      return el ? el.closest('.check').querySelector('span').textContent : v;
    }).join(', ');
  } else { btn.textContent=checked.length+' selecionados'; }

  // Propagar filtros compartilhados (ano, mes, movimento, proprietario) para todas as páginas
  const SHARED_KEYS = ['ano','mes','movimento','proprietario'];
  if(SHARED_KEYS.includes(key)){
    Object.keys(SEL).forEach(function(pg){
      if(pg === page) return;
      if(SEL[pg][key] === undefined) return;
      SEL[pg][key] = SEL[page][key].slice();
      // Atualizar visual do filtro correspondente na outra página
      let otherOuter = document.querySelector('#filters-'+pg+' .multi[data-key="'+key+'"]');
      if(!otherOuter) return;
      let otherPanel = otherOuter.querySelector('.multi-panel');
      otherPanel.querySelectorAll('input[type=checkbox]').forEach(function(cb){
        if(isAll){
          cb.checked = cb.value === '_all_';
        } else {
          cb.checked = checked.includes(cb.value);
        }
      });
      let otherBtn = otherOuter.querySelector('.multi-btn');
      if(isAll){ otherBtn.textContent='Todos'; }
      else if(checked.length<=2){
        otherBtn.textContent=checked.map(function(v){
          let el=otherPanel.querySelector('input[value="'+CSS.escape(v)+'"]');
          return el ? el.closest('.check').querySelector('span').textContent : v;
        }).join(', ');
      } else { otherBtn.textContent=checked.length+' selecionados'; }
    });
  }

  // Quando qualquer filtro da aba Produtos muda, volta para visão de categorias
  if(page === 'prod') PROD_DRILL = null;
  // Quando qualquer filtro do executivo muda, volta para visão geral de faixas
  if(page === 'exec') FV_DRILL = null;

  CP={prod:1,prop:1,down:1,ana:1};
  renderAll();
}

function initFilters(){
  let anos  = DATA.anos.map(String);
  let meses = DATA.meses.map(m=>String(m.num));
  let mFmt  = function(v){ let m=DATA.meses.find(x=>String(x.num)===v); return m?m.nome:v; };

  makeFilter('filters-exec','exec','ano','Ano',anos,null,false);
  makeFilter('filters-exec','exec','mes','Mês',meses,mFmt,false);
  makeFilter('filters-exec','exec','movimento','Movimento',DATA.movimentos,null,false);
  makeFilter('filters-exec','exec','proprietario','Proprietário do Movimento',DATA.proprietarios);
  makeFilter('filters-exec','exec','faixa_vida','Faixa de Vida',(DATA.faixas_vida||[]),null,false);

  makeFilter('filters-prod','prod','ano','Ano',anos,null,false);
  makeFilter('filters-prod','prod','mes','Mês',meses,mFmt,false);
  makeFilter('filters-prod','prod','movimento','Movimento',DATA.movimentos,null,false);
  makeFilter('filters-prod','prod','categoria','Categoria',(DATA.categorias||[]));
  makeFilter('filters-prod','prod','faixa_vida','Faixa de Vida',(DATA.faixas_vida||[]),null,false);

  makeFilter('filters-down','down','ano','Ano',anos,null,false);
  makeFilter('filters-down','down','mes','Mês',meses,mFmt,false);
  makeFilter('filters-down','down','produto','Produto',DATA.produtos);
  makeFilter('filters-down','down','situacao','Situação',DATA.situacoes_down);
  makeFilter('filters-down','down','faixa_vida','Faixa de Vida',(DATA.faixas_vida||[]),null,false);

  makeFilter('filters-prop','prop','ano','Ano',anos,null,false);
  makeFilter('filters-prop','prop','mes','Mês',meses,mFmt,false);
  makeFilter('filters-prop','prop','movimento','Movimento',DATA.movimentos,null,false);
  makeFilter('filters-prop','prop','proprietario','Proprietário do Movimento',DATA.proprietarios);
  makeFilter('filters-prop','prop','faixa_vida','Faixa de Vida',(DATA.faixas_vida||[]),null,false);

  makeFilter('filters-ana','ana','ano','Ano',anos,null,false);
  makeFilter('filters-ana','ana','mes','Mês',meses,mFmt,false);
  makeFilter('filters-ana','ana','movimento','Movimento',DATA.movimentos,null,false);
  makeFilter('filters-ana','ana','produto','Produto',DATA.produtos);
  makeFilter('filters-ana','ana','proprietario','Proprietário do Movimento',DATA.proprietarios);
  makeFilter('filters-ana','ana','faixa_vida','Faixa de Vida',(DATA.faixas_vida||[]),null,false);
}

function resetFiltros(){
  if(current==='home') return;
  // Reset current page filters
  Object.keys(SEL[current]).forEach(k=>SEL[current][k]=[]);
  document.querySelectorAll('#filters-'+current+' .multi').forEach(function(outer){
    outer.querySelectorAll('input[type=checkbox]').forEach(i=>i.checked=i.value==='_all_');
    outer.querySelector('.multi-btn').textContent='Todos';
  });
  // Reset drill-down de produtos quando estiver na aba prod
  if(current==='prod') PROD_DRILL = null;
  // Reset drill-down de faixa de vida quando estiver no executivo
  if(current==='exec') FV_DRILL = null;
  // Reset shared keys in all other pages
  const SHARED_KEYS = ['ano','mes','movimento','proprietario'];
  Object.keys(SEL).forEach(function(pg){
    if(pg===current) return;
    SHARED_KEYS.forEach(function(k){
      if(SEL[pg][k]!==undefined) SEL[pg][k]=[];
    });
    let cont = document.querySelector('#filters-'+pg);
    if(!cont) return;
    SHARED_KEYS.forEach(function(k){
      let outer = cont.querySelector('.multi[data-key="'+k+'"]');
      if(!outer) return;
      outer.querySelectorAll('input[type=checkbox]').forEach(i=>i.checked=i.value==='_all_');
      outer.querySelector('.multi-btn').textContent='Todos';
    });
  });
  CP={prod:1,prop:1,down:1,ana:1};
  renderAll();
}

document.addEventListener('click',function(e){
  if(!e.target.closest('.multi')) document.querySelectorAll('.multi.open').forEach(m=>m.classList.remove('open'));
});

Chart.defaults.font.family="'Segoe UI',Arial,sans-serif";
Chart.defaults.font.size=12;
Chart.defaults.color='#475569';
const CUP='#00AFC8', CDN='#002F6C', CURL='#D97706', CMIG='#1D4ED8', CTST='#475569', CPLAN='#6D28D9', CROB='#047857', COUT='#1E293B';
const SCOL=['#0EA5E9','#1E3A5F','#0369A1','#38BDF8','#075985','#7DD3FC','#0C4A6E','#BAE6FD'];

function renderExec(){
  let rs=fRows('exec');
  let tot=rs.length, up=sumM(rs,'Upgrade'), dn=sumM(rs,'Downgrade');
  let url=sumM(rs,'Criação de URL'), mig=sumM(rs,'Migração'), tst=sumM(rs,'Teste'), plan=sumM(rs,'Alteração de Plano'), rob=sumM(rs,'Robô'), out=sumM(rs,'Outros');
  let sal=up-dn, vup=sum(rs.filter(r=>r.Movimento==='Upgrade'),'total_up_RS'), vdn=sum(rs.filter(r=>r.Movimento==='Downgrade'),'total_down_RS');

  let outros = url + mig + tst + plan + rob + out;
  id('kTotal').textContent=fmt(tot);
  id('kTotalValor').textContent=brl(vup+vdn);
  id('kPeriodo').textContent=periodTxt(rs);
  id('kUp').textContent=fmt(up);
  id('kUpPct').textContent=pct(tot?up/tot*100:0);
  id('kValorUp').textContent=brl(vup);
  id('kDown').textContent=fmt(dn);
  id('kDownPct').textContent=pct(tot?dn/tot*100:0);
  id('kValorDown').textContent=brl(vdn);
  id('kSaldo').textContent=(sal>=0?'+':'')+fmt(sal);
  id('kSaldo').style.color=sal>=0?'#059669':'#DC2626';
  id('kSaldoValor').textContent=brl(vup-vdn);
  id('kSaldoValor').style.color=vup-vdn>=0?'#059669':'#DC2626';
  id('kSaldoTxt').textContent=sal>=0?'saldo positivo ▲':'saldo negativo ▼';
  id('kOutros').textContent=fmt(outros);
  id('kOutrosTxt').textContent='Criação de URL: '+fmt(url)+' · Migração: '+fmt(mig)+' · Teste: '+fmt(tst)+' · Robô: '+fmt(rob)+' · Alt. Plano: '+fmt(plan)+(out?' · Outros: '+fmt(out):'');

  let ms=months(rs);
  let lbs=ms.map(m=>m.label);
  let dUp =ms.map(m=>sumM(rs.filter(r=>r.AnoMes===m.am),'Upgrade'));
  let dDn =ms.map(m=>sumM(rs.filter(r=>r.AnoMes===m.am),'Downgrade'));
  let dUrl=ms.map(m=>sumM(rs.filter(r=>r.AnoMes===m.am),'Criação de URL'));
  let dMig=ms.map(m=>sumM(rs.filter(r=>r.AnoMes===m.am),'Migração'));
  let dTst=ms.map(m=>sumM(rs.filter(r=>r.AnoMes===m.am),'Teste'));
  let dPlan=ms.map(m=>sumM(rs.filter(r=>r.AnoMes===m.am),'Alteração de Plano'));
  let dRob=ms.map(m=>sumM(rs.filter(r=>r.AnoMes===m.am),'Robô'));
  let dOut=ms.map(m=>sumM(rs.filter(r=>r.AnoMes===m.am),'Outros'));
  let dSal=dUp.map((v,i)=>v-dDn[i]);
  let dUpV =ms.map(m=>sum(rs.filter(r=>r.AnoMes===m.am&&r.Movimento==='Upgrade'),'total_up_RS'));
  let dDnV =ms.map(m=>sum(rs.filter(r=>r.AnoMes===m.am&&r.Movimento==='Downgrade'),'total_down_RS'));
  let dUrlV=ms.map(m=>sum(rs.filter(r=>r.AnoMes===m.am&&r.Movimento==='Criação de URL'),'total_up_RS'));
  let dMigV=ms.map(m=>sum(rs.filter(r=>r.AnoMes===m.am&&r.Movimento==='Migração'),'total_up_RS'));
  let dTstV=ms.map(m=>sum(rs.filter(r=>r.AnoMes===m.am&&r.Movimento==='Teste'),'total_up_RS'));
  let dPlanV=ms.map(m=>sum(rs.filter(r=>r.AnoMes===m.am&&r.Movimento==='Alteração de Plano'),'total_up_RS'));
  let dRobV=ms.map(m=>sum(rs.filter(r=>r.AnoMes===m.am&&r.Movimento==='Robô'),'total_up_RS'));
  let dOutV=ms.map(m=>sum(rs.filter(r=>r.AnoMes===m.am&&r.Movimento==='Outros'),'total_up_RS'));
  let dSalV=dUpV.map((v,i)=>v-dDnV[i]);

  window._execChartData = {lbs, dUp, dDn, dUrl, dMig, dTst, dPlan, dRob, dOut, dSal, dUpV, dDnV, dUrlV, dMigV, dTstV, dPlanV, dRobV, dOutV, dSalV};
  window._execChartMode = window._execChartMode || 'qtd';
  buildExecChart();

  renderExecInsights(rs,up,dn,sal,vup,vdn,tot);
  renderExecFaixaVida(rs);
  renderMensal(rs);
}

function setChartMode(mode){
  window._execChartMode = mode;
  id('btnQtd').style.background = mode==='qtd'?'var(--cyan)':'transparent';
  id('btnQtd').style.color       = mode==='qtd'?'#fff':'#667085';
  id('btnVal').style.background  = mode==='val'?'var(--cyan)':'transparent';
  id('btnVal').style.color       = mode==='val'?'#fff':'#667085';
  buildExecChart();
}

function buildExecChart(){
  const d=window._execChartData; if(!d) return;
  const isVal = window._execChartMode==='val';
  const upD  = isVal?d.dUpV :d.dUp;
  const dnD  = isVal?d.dDnV :d.dDn;
  const urlD = isVal?d.dUrlV:d.dUrl;
  const migD = isVal?d.dMigV:d.dMig;
  const tstD = isVal?d.dTstV:d.dTst;
  const planD = isVal?d.dPlanV:d.dPlan;
  const robD = isVal?d.dRobV:d.dRob;
  const outD = isVal?d.dOutV:d.dOut;

  id('chartExecTitle').textContent = isVal?'Valor mensal por tipo (R$)':'Volume mensal por tipo';

  const LABEL_COLORS = {
    'Upgrade':  {border:'#00AFC8', text:'#007C8E'},
    'Downgrade':{border:'#002F6C', text:'#002F6C'},
    'Criação de URL': {border:'#D97706', text:'#92400E'},
    'Migração': {border:'#2563EB', text:'#1D4ED8'},
    'Teste':    {border:'#64748B', text:'#334155'},
    'Alteração de Plano': {border:'#7C3AED', text:'#4C1D95'},
    'Robô':     {border:'#059669', text:'#065F46'},
    'Outros':   {border:'#334155', text:'#1E293B'},
  };

  // Balões de quantidade/valor em cima das barras.
  // Ajuste feito para NÃO esconder Desconto/Robô e NÃO deixar balões sobrepostos.
  // Regra: tenta primeiro em cima da barra; se bater em outro balão, desloca para os lados
  // e depois sobe uma linha, mantendo o balão visível dentro da área do gráfico.
  const datalabelPlugin = {
    id:'datalabelUpDown',
    afterDatasetsDraw(chart){
      const ctx = chart.ctx;
      const area = chart.chartArea;
      const boxes = [];
      const margem = 4;

      function overlap(a,b){
        return !(a.x+a.w < b.x || b.x+b.w < a.x || a.y+a.h < b.y || b.y+b.h < a.y);
      }
      function clamp(v,min,max){ return Math.max(min, Math.min(max, v)); }

      chart.data.datasets.forEach((ds,i)=>{
        const colors = LABEL_COLORS[ds.label];
        if(!colors) return;
        const meta = chart.getDatasetMeta(i);

        meta.data.forEach((bar,j)=>{
          const val = Number(ds.data[j] || 0);
          if(!val) return; // não mostra zero no gráfico para não poluir

          const allValues = chart.data.datasets.flatMap(x => x.data || []).filter(x => Number(x)>0).map(Number);
          const maxVal = Math.max(...allValues, 0);
          // No gráfico de valor, valores muito pequenos continuam apenas no tooltip.
          // No gráfico de quantidade, aparece tudo: 1, 2, 3 etc.
          if(isVal && maxVal && val < maxVal * 0.035) return;

          const txt = isVal
            ? new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL',maximumFractionDigits:0}).format(val)
            : String(val);

          const px = isVal ? 9 : 10;
          ctx.font = 'bold '+px+'px Segoe UI,Arial,sans-serif';
          const tw = ctx.measureText(txt).width;
          const bw = tw + 12;
          const bh = 17;

          const xTentativas = [0,-14,14,-28,28,-42,42,-56,56];
          const yTentativas = [0,20,40,60,80,100,120];

          let escolhido = null;
          for(const yExtra of yTentativas){
            for(const xExtra of xTentativas){
              let bx = bar.x - bw/2 + xExtra;
              let by = bar.y - bh - 7 - yExtra;

              bx = clamp(bx, area.left + margem, area.right - bw - margem);
              by = Math.max(area.top + margem, by);

              const box = {x:bx-3, y:by-3, w:bw+6, h:bh+6};
              if(!boxes.some(b=>overlap(box,b))){
                escolhido = {bx,by,box};
                break;
              }
            }
            if(escolhido) break;
          }

          // Último recurso: ainda mostra o balão, mas joga para uma posição alternada.
          if(!escolhido){
            const bx = clamp(bar.x - bw/2 + ((i%2===0)?-36:36), area.left + margem, area.right - bw - margem);
            const by = area.top + margem + ((j+i)%5)*19;
            escolhido = {bx,by,box:{x:bx-3,y:by-3,w:bw+6,h:bh+6}};
          }

          boxes.push(escolhido.box);

          ctx.save();
          ctx.fillStyle = '#ffffff';
          ctx.strokeStyle = colors.border;
          ctx.lineWidth = 1.5;
          ctx.beginPath();
          ctx.roundRect(escolhido.bx, escolhido.by, bw, bh, 5);
          ctx.fill();
          ctx.stroke();
          ctx.fillStyle = colors.text;
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';
          ctx.fillText(txt, escolhido.bx + bw/2, escolhido.by + bh/2);
          ctx.restore();
        });
      });
    }
  };

  newChart('chartExec',{type:'bar',data:{labels:d.lbs,datasets:[
    {label:'Upgrade',   data:upD, backgroundColor:CUP,  borderRadius:4,order:2},
    {label:'Downgrade', data:dnD, backgroundColor:CDN,  borderRadius:4,order:2},
    {label:'Criação de URL', data:urlD,backgroundColor:CURL, borderRadius:4,order:2},
    {label:'Migração',  data:migD,backgroundColor:CMIG, borderRadius:4,order:2},
    {label:'Teste',     data:tstD,backgroundColor:CTST, borderRadius:4,order:2},
    {label:'Alteração de Plano', data:planD,backgroundColor:CPLAN, borderRadius:4,order:2},
    {label:'Robô',      data:robD,backgroundColor:CROB, borderRadius:4,order:2}
  ]},options:{
    responsive:true,maintainAspectRatio:false,
    plugins:{
      legend:{position:'top',labels:{boxWidth:12,padding:16,usePointStyle:true}},
      tooltip:{mode:'index',intersect:false,callbacks:{label:function(ctx){
        if(isVal) return ' '+ctx.dataset.label+': '+new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL',maximumFractionDigits:0}).format(ctx.raw);
        return ' '+ctx.dataset.label+': '+ctx.raw;
      }}}
    },
    scales:{
      x:{grid:{display:false},ticks:{color:'#94A3B8'}},
      y:{display:false,afterDataLimits:ax=>{ax.max=ax.max*1.55}}
    }
  },plugins:[datalabelPlugin]});
}

function renderExecInsights(rs,up,dn,sal,vup,vdn,tot){
  let ms=months(rs);
  let byM=ms.map(m=>{let r=rs.filter(x=>x.AnoMes===m.am);return{label:m.label,up:sumM(r,'Upgrade'),dn:sumM(r,'Downgrade'),tot:r.length}});
  let forte=[...byM].sort((a,b)=>b.tot-a.tot)[0]||{label:'—',tot:0};
  let pior =[...byM].sort((a,b)=>b.dn -a.dn )[0]||{label:'—',dn:0};
  let pUp=top1(grpCnt(rs,'Produto_Curto','Upgrade'));
  let pDn=top1(grpCnt(rs,'Produto_Curto','Downgrade'));
  let sit=top1(grpCnt(rs.filter(r=>r.Movimento==='Downgrade'),'Situacao_Down'));
  let pd=tot?dn/tot*100:0;
  let risco=pd>=30?'Alto':pd>=20?'Moderado':'Controlado';
  let rc=pd>=30?'warn':pd>=20?'amb':'pos';
  id('execInsights').innerHTML=`
  <div class="fv-side-insights" style="box-shadow:none;margin:0">
    <div class="fv-side-title">Insights do gráfico</div>
    <div class="fv-insights-grid">
      <div class="fv-insight-line">
        <span>Mês de maior volume</span>
        <strong>${forte.label}</strong>
        <small>${fmt(forte.tot)} movimentos registrados</small>
      </div>
      <div class="fv-insight-line">
        <span>Pressão de downgrade</span>
        <strong>${risco}</strong>
        <small>${pct(pd)} do volume são downgrades</small>
      </div>
      <div class="fv-insight-line">
        <span>Principal motivo de down</span>
        <strong>${esc(sit.key)}</strong>
        <small>${fmt(sit.value)} ocorrência(s)</small>
      </div>
      <div class="fv-insight-line">
        <span>Top upgrade</span>
        <strong>${esc(pUp.key)}</strong>
        <small>${fmt(pUp.value)} serviços</small>
      </div>
      <div class="fv-insight-line">
        <span>Top downgrade</span>
        <strong>${esc(pDn.key)}</strong>
        <small>${fmt(pDn.value)} serviços</small>
      </div>
      <div class="fv-insight-line">
        <span>Mês de maior pressão</span>
        <strong>${pior.label}</strong>
        <small>${fmt(pior.dn)} downgrades</small>
      </div>
    </div>
  </div>`;
}

// Estado do drill-down de faixa de vida
let FV_DRILL = null; // null = visão geral; string = faixa selecionada

function renderExecFaixaVida(rs){
  const ORDEM = ["1 - 0 a 30","2 - 31 a 60","3 - 61 a 90","4 - 91 a 120",
    "5 - 121 a 180","6 - 181 a 240","7 - 241 a 360",
    "8 - 361 a 540","9 - 541 a 720","10 - Acima de 721","Não calculado"];

  let com = rs.filter(r=>r.Faixa_Vida);
  if(!com.length){
    id('execFaixaVida').innerHTML=`<div style="color:#94A3B8;font-size:13px;padding:12px 0">Nenhum dado de faixa de vida disponível — verifique a conexão com o banco.</div>`;
    return;
  }

  function produtosDaLinha(r){
    let txt = String(r.produtos_digisac||r.Produto_Curto||'Não informado');
    return txt.split(';').map(x=>x.trim()).filter(x=>x && x.toLowerCase()!=='nan' && x.toLowerCase()!=='não informado' && x.toLowerCase()!=='nao informado');
  }
  function topProduto(rows, movimento=null){
    let map = {};
    rows.filter(r=>!movimento || r.Movimento===movimento).forEach(r=>{
      produtosDaLinha(r).forEach(p=>{ map[p]=(map[p]||0)+1; });
    });
    let arr = Object.entries(map).map(([key,value])=>({key,value})).sort((a,b)=>b.value-a.value || a.key.localeCompare(b.key));
    return arr[0] || {key:'—',value:0};
  }
  function produtoSet(rows, movimento=null){
    let st = new Set();
    rows.filter(r=>!movimento || r.Movimento===movimento).forEach(r=>produtosDaLinha(r).forEach(p=>st.add(p)));
    return st;
  }
  function movimentoPredominante(rows){
    let arr = ['Upgrade','Downgrade','Migração','Teste','Desconto','Robô'].map(m=>({key:m,value:rows.filter(r=>r.Movimento===m).length})).sort((a,b)=>b.value-a.value);
    return arr[0] || {key:'—',value:0};
  }

  // DRILL: ao clicar em uma faixa, mostra detalhe da faixa e clientes
  if(FV_DRILL !== null){
    let drillRs = rs.filter(r=>r.Faixa_Vida===FV_DRILL);
    let up = drillRs.filter(r=>r.Movimento==='Upgrade').length;
    let dn = drillRs.filter(r=>r.Movimento==='Downgrade').length;
    let mig = drillRs.filter(r=>r.Movimento==='Migração').length;
    let out = drillRs.filter(r=>r.Movimento==='Outros').length;
    let pred = movimentoPredominante(drillRs);
    let prodG = topProduto(drillRs);
    let prodUp = topProduto(drillRs,'Upgrade');
    let prodDn = topProduto(drillRs,'Downgrade');

    let subMap = {};
    drillRs.forEach(r=>{
      let k = r.subdominio||'—';
      if(!subMap[k]) subMap[k]={sub:k,up:0,dn:0,mig:0,out:0,tot:0,vup:0,vdn:0,dias:r.Tempo_vida_cli,prod:{}};
      subMap[k].tot++;
      if(r.Movimento==='Upgrade'){subMap[k].up++;subMap[k].vup+=(Number(r.total_up_RS)||0);} 
      else if(r.Movimento==='Downgrade'){subMap[k].dn++;subMap[k].vdn+=(Number(r.total_down_RS)||0);} 
      else if(r.Movimento==='Migração') subMap[k].mig++;
      else subMap[k].out++;
      produtosDaLinha(r).forEach(p=>{subMap[k].prod[p]=(subMap[k].prod[p]||0)+1;});
    });
    let subs = Object.values(subMap).sort((a,b)=>b.tot-a.tot);
    let topProdSub = obj => {
      let a=Object.entries(obj.prod||{}).sort((x,y)=>y[1]-x[1]);
      return a.length?a[0][0]:'—';
    };

    id('execFaixaVida').innerHTML = `
      <div class="fv-drill-header">
        <button class="fv-drill-back" onclick="fvDrillBack()">← Voltar para resumo das faixas</button>
        <span style="color:#94A3B8">›</span>
        <span style="font-weight:800;color:#0C4A6E">${esc(FV_DRILL)}</span>
        <span class="fv-drill-tag">${fmt(drillRs.length)} movimentos</span>
      </div>
      <div style="display:grid;grid-template-columns:1fr;gap:16px;margin-bottom:16px">
        <div class="fv-side-insights">
          <div class="fv-side-title">Insights da faixa</div>
          <div class="fv-insights-grid">
            <div class="fv-insight-line"><span>Movimento predominante</span><strong>${esc(pred.key)}</strong><small>${fmt(pred.value)} registros</small></div>
            <div class="fv-insight-line"><span>Produto mais movimentado</span><strong>${esc(prodG.key)}</strong><small>${fmt(prodG.value)} movimentações</small></div>
            <div class="fv-insight-line"><span>Produto líder em Upgrade</span><strong>${esc(prodUp.key)}</strong><small>${fmt(prodUp.value)} upgrades</small></div>
            <div class="fv-insight-line"><span>Produto líder em Downgrade</span><strong>${esc(prodDn.key)}</strong><small>${fmt(prodDn.value)} downgrades</small></div>
          </div>
          <div class="fv-mini-grid">
            <div><b>${fmt(up)}</b><span>Up</span></div><div><b>${fmt(dn)}</b><span>Down</span></div><div><b>${fmt(mig)}</b><span>Mig.</span></div><div><b>${fmt(out)}</b><span>Outros</span></div>
          </div>
        </div>
        <div>
          <div class="fv-section-title" style="margin-bottom:10px">Subdomínios da faixa</div>
          <div class="table-wrap" style="max-height:430px">
            <table class="fv-drill-table">
              <thead><tr>
                <th>Subdomínio</th><th class="r">Dias</th><th class="r">Up</th><th class="r">Down</th><th class="r">Mig.</th><th class="r">Outros</th><th class="r">Total</th><th>Produto destaque</th>
              </tr></thead>
              <tbody>${subs.map(r=>`<tr>
                <td><strong>${esc(r.sub)}</strong></td>
                <td class="r" style="color:#64748B">${r.dias!=null?fmt(r.dias):'—'}</td>
                <td class="r" style="color:#059669">${r.up||'—'}</td>
                <td class="r" style="color:#DC2626">${r.dn||'—'}</td>
                <td class="r">${r.mig||'—'}</td>
                <td class="r">${r.out||'—'}</td>
                <td class="r"><strong>${fmt(r.tot)}</strong></td>
                <td>${esc(topProdSub(r))}</td>
              </tr>`).join('')}</tbody>
            </table>
          </div>
        </div>
      </div>`;
    return;
  }

  // RESUMO por faixa: tabela + insights ao lado
  let faixas = ORDEM.filter(f=>com.some(r=>r.Faixa_Vida===f));
  let resumo = faixas.map(f=>{
    let rows = com.filter(r=>r.Faixa_Vida===f);
    let up = rows.filter(r=>r.Movimento==='Upgrade').length;
    let dn = rows.filter(r=>r.Movimento==='Downgrade').length;
    let mig = rows.filter(r=>r.Movimento==='Migração').length;
    let out = rows.filter(r=>r.Movimento==='Outros').length;
    return {
      faixa:f, mov:rows.length, up, dn, mig, out,
      pred:movimentoPredominante(rows),
      prod:topProduto(rows),
      prodUp:topProduto(rows,'Upgrade'),
      prodDn:topProduto(rows,'Downgrade')
    };
  }).sort((a,b)=>b.mov-a.mov);

  let faixaMais = resumo[0] || {faixa:'—',mov:0};
  let predGeral = movimentoPredominante(com);
  let prodGeral = topProduto(com);
  let prodUpGeral = topProduto(com,'Upgrade');
  let prodDnGeral = topProduto(com,'Downgrade');
  let novosDown = com.filter(r=>r.Faixa_Vida==='1 - 0 a 30' && r.Movimento==='Downgrade').length;

  // Produto em comum entre a faixa que mais faz upgrade e a faixa que mais faz downgrade
  let topFaixaUp = [...resumo].sort((a,b)=>b.up-a.up)[0];
  let topFaixaDn = [...resumo].sort((a,b)=>b.dn-a.dn)[0];
  let comumTxt = 'Sem produto em comum';
  if(topFaixaUp && topFaixaDn){
    let setUp = produtoSet(com.filter(r=>r.Faixa_Vida===topFaixaUp.faixa),'Upgrade');
    let setDn = produtoSet(com.filter(r=>r.Faixa_Vida===topFaixaDn.faixa),'Downgrade');
    let comuns = [...setUp].filter(p=>setDn.has(p));
    if(comuns.length){
      let score = comuns.map(p=>({key:p,value:com.filter(r=>produtosDaLinha(r).includes(p)).length})).sort((a,b)=>b.value-a.value);
      comumTxt = score.slice(0,3).map(x=>x.key).join(', ');
    }
  }
  let semFaixa = rs.filter(r=>!r.Faixa_Vida).length;

  id('execFaixaVida').innerHTML = `
    <div style="display:grid;grid-template-columns:1fr;gap:16px;align-items:start">
      <div class="fv-side-insights">
        <div class="fv-side-title">Insights das faixas</div>
        <div class="fv-insights-grid">
          <div class="fv-insight-line"><span>Faixa mais movimentada</span><strong>${esc(faixaMais.faixa)}</strong><small>${fmt(faixaMais.mov)} movimentações</small></div>
          <div class="fv-insight-line"><span>Movimento predominante</span><strong>${esc(predGeral.key)}</strong><small>${fmt(predGeral.value)} registros</small></div>
          <div class="fv-insight-line"><span>Produto mais movimentado</span><strong>${esc(prodGeral.key)}</strong><small>${fmt(prodGeral.value)} movimentações</small></div>
          <div class="fv-insight-line"><span>Produto líder em Upgrade</span><strong>${esc(prodUpGeral.key)}</strong><small>${fmt(prodUpGeral.value)} upgrades</small></div>
          <div class="fv-insight-line"><span>Produto líder em Downgrade</span><strong>${esc(prodDnGeral.key)}</strong><small>${fmt(prodDnGeral.value)} downgrades</small></div>
          <div class="fv-insight-line"><span>Clientes novos em downgrade</span><strong>${fmt(novosDown)} clientes</strong><small>faixa 1 - 0 a 30 dias</small></div>
        </div>
      </div>
      <div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
          <div>
            <div class="fv-section-title" style="margin-bottom:2px">Resumo por faixa: movimentações e produtos</div>
            <div style="font-size:11px;color:#94A3B8">Clique em uma linha para abrir os clientes da faixa.</div>
          </div>
        </div>
        <div class="table-wrap" style="max-height:430px">
          <table class="fv-drill-table">
            <thead><tr>
              <th>Faixa</th>
              <th class="r">Mov.</th>
              <th class="r">Up</th>
              <th class="r">Down</th>
              <th class="r">Mig.</th>
              <th class="r">Outros</th>
              <th>Predominante</th>
              <th>Produto destaque</th>
              <th>Mais Up</th>
              <th>Mais Down</th>
            </tr></thead>
            <tbody>${resumo.map(r=>`<tr onclick="fvDrillInto('${r.faixa.replace(/'/g,"\\'")}')" style="cursor:pointer" title="Clique para ver detalhes da ${esc(r.faixa)}">
              <td><strong>${esc(r.faixa)}</strong></td>
              <td class="r"><strong>${fmt(r.mov)}</strong></td>
              <td class="r" style="color:#059669">${fmt(r.up)}</td>
              <td class="r" style="color:#DC2626">${fmt(r.dn)}</td>
              <td class="r">${fmt(r.mig)}</td>
              <td class="r">${fmt(r.out)}</td>
              <td>${esc(r.pred.key)}</td>
              <td>${esc(r.prod.key)}</td>
              <td>${esc(r.prodUp.key)}</td>
              <td>${esc(r.prodDn.key)}</td>
            </tr>`).join('')}</tbody>
          </table>
        </div>
        ${semFaixa>0?`<div style="margin-top:10px;font-size:11px;color:#94A3B8">${fmt(semFaixa)} registro${semFaixa>1?'s':''} sem faixa calculada (subdomínio não encontrado no banco).</div>`:''}
      </div>
    </div>`;
}

function fvDrillInto(faixa){
  FV_DRILL = faixa;
  let rs = fRows('exec');
  renderExecFaixaVida(rs);
}

function fvDrillBack(){
  FV_DRILL = null;
  let rs = fRows('exec');
  renderExecFaixaVida(rs);
}

function renderMensal(rs){
  let anos=[...new Set(rs.map(r=>r.Ano))].sort();
  if(!anos.length){id('tblMensalQtd').innerHTML='';id('tblMensalVal').innerHTML='';return}
  let ano=anos[anos.length-1];
  let sub=rs.filter(r=>r.Ano===ano);
  let meses=[...new Set(sub.map(r=>r.MesNum))].sort((a,b)=>a-b);
  let mN=m=>DATA.meses.find(x=>x.num===m).nome;

  const ALL_MOVS = ['Upgrade','Downgrade','Criação de URL','Migração','Teste','Robô','Alteração de Plano'];
  function pillClass(mv){
    if(mv==='Upgrade') return 'p-up';
    if(mv==='Downgrade') return 'p-dn';
    if(mv==='Criação de URL') return 'p-url';
    if(mv==='Migração') return 'p-mig';
    if(mv==='Teste') return 'p-tst';
    if(mv==='Robô') return 'p-rob';
    if(mv==='Alteração de Plano') return 'p-dsc';
    return 'p-oth';
  }

  // --- Tabela de Quantidade ---
  let hQ=`<table class="mt"><thead><tr><th>Tipo</th>${meses.map(m=>`<th>${mN(m)}</th>`).join('')}<th>Total</th></tr></thead><tbody>`;
  ALL_MOVS.forEach(mv=>{
    let vs=meses.map(m=>sumM(sub.filter(r=>r.MesNum===m),mv));
    let t=vs.reduce((a,b)=>a+b,0); if(!t) return;
    hQ+=`<tr><td><span class="pill ${pillClass(mv)}">${mv}</span></td>${vs.map(v=>`<td>${v?fmt(v):'—'}</td>`).join('')}<td><strong>${fmt(t)}</strong></td></tr>`;
  });
  let totsQ=meses.map(m=>sub.filter(r=>r.MesNum===m).length);
  hQ+=`<tr class="tr-tot"><td>Total ${ano}</td>${totsQ.map(v=>`<td>${fmt(v)}</td>`).join('')}<td>${fmt(totsQ.reduce((a,b)=>a+b,0))}</td></tr></tbody></table>`;
  id('tblMensalQtd').innerHTML=hQ;

  // --- Tabela de Valor ---
  let hV=`<table class="mt"><thead><tr><th>Tipo</th>${meses.map(m=>`<th>${mN(m)}</th>`).join('')}<th>Total</th></tr></thead><tbody>`;
  const valColMap={'Upgrade':'total_up_RS','Downgrade':'total_down_RS'};
  ALL_MOVS.forEach(mv=>{
    let col=valColMap[mv]||'total_up_RS';
    let vs=meses.map(m=>sum(sub.filter(r=>r.MesNum===m&&r.Movimento===mv),col));
    let t=vs.reduce((a,b)=>a+b,0); if(!t) return;
    hV+=`<tr><td><span class="pill ${pillClass(mv)}">${mv}</span></td>${vs.map(v=>`<td>${v?brl(v):'—'}</td>`).join('')}<td><strong>${brl(t)}</strong></td></tr>`;
  });
  let totsVUp=meses.map(m=>sum(sub.filter(r=>r.MesNum===m),'total_up_RS'));
  let totsVDn=meses.map(m=>sum(sub.filter(r=>r.MesNum===m),'total_down_RS'));
  let totsV=totsVUp.map((v,i)=>v+totsVDn[i]);
  hV+=`<tr class="tr-tot"><td>Total ${ano}</td>${totsV.map(v=>`<td>${brl(v)}</td>`).join('')}<td>${brl(totsV.reduce((a,b)=>a+b,0))}</td></tr></tbody></table>`;
  id('tblMensalVal').innerHTML=hV;
}

function barRows(divId,data,money,color){
  let top=data.slice(0,12), mx=top.reduce((a,d)=>Math.max(a,d.value),1);
  id(divId).innerHTML=top.map(d=>`
    <div class="bar-row">
      <div class="bar-lbl" title="${esc(d.key)}">${esc(d.key)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.round(d.value/mx*100)}%;background:${color}"></div></div>
      <div class="bar-cnt">${money?brl(d.value):fmt(d.value)}</div>
    </div>`).join('');
}

// Barra clicável: ao clicar entra no drill-down da categoria
function barRowsDrill(divId, data, money, color, onClickFn){
  let top=data.slice(0,15), mx=top.reduce((a,d)=>Math.max(a,d.value),1);
  id(divId).innerHTML=top.map(d=>{
    const pctW=Math.round(d.value/mx*100);
    const clickAttr = onClickFn ? `onclick="${onClickFn}('${d.key.replace(/'/g,"\\'")}')"`:'';
    const cursor = onClickFn ? 'cursor:pointer;' : '';
    const hoverStyle = onClickFn ? 'onmouseover="this.style.background=\'#F0F9FF\'" onmouseout="this.style.background=\'\'"' : '';
    return `<div class="bar-row" style="${cursor}border-radius:8px;padding:4px 6px;transition:.15s;" ${clickAttr} ${hoverStyle} title="${onClickFn?'Clique para ver produtos de '+esc(d.key):esc(d.key)}">
      <div class="bar-lbl" style="${onClickFn?'color:#0EA5E9;font-weight:700;':''}" title="${esc(d.key)}">${esc(d.key)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${pctW}%;background:${color}"></div></div>
      <div class="bar-cnt">${money?brl(d.value):fmt(d.value)}${onClickFn?'<span style="margin-left:5px;font-size:10px;color:#94A3B8">›</span>':''}</div>
    </div>`;
  }).join('');
}

function renderProd(){
  let rsQtd = fRows('prod');
  let rsVal = fRowsProdValor();

  if(PROD_DRILL === null){
    // ── Visão agrupada: barras por CATEGORIA ──
    id('prod-breadcrumb').style.display = 'none';
    id('prod-drill-hint').style.display = 'flex';
    id('prodPageTitle').textContent = 'Produtos';
    id('prodPageSub').textContent = 'Expansão, redução e impacto financeiro por categoria';
    id('hPupQ').textContent = 'Upgrade — Quantidade por Categoria';
    id('hPdnQ').textContent = 'Downgrade — Quantidade por Categoria';
    id('hPupV').textContent = 'Upgrade — Valor por Categoria';
    id('hPdnV').textContent = 'Downgrade — Valor por Categoria';
    id('tblProdTitle').textContent = 'Resumo por Categoria';

    barRowsDrill('bPupQ', grpCnt(rsQtd,'Categoria','Upgrade'),   false, CUP, 'prodDrillInto');
    barRowsDrill('bPdnQ', grpCnt(rsQtd,'Categoria','Downgrade'), false, CDN, 'prodDrillInto');
    barRowsDrill('bPupV', grp(rsVal,'Categoria','total_up_RS','Upgrade'),     true,  CUP, 'prodDrillInto');
    barRowsDrill('bPdnV', grp(rsVal,'Categoria','total_down_RS','Downgrade'), true,  CDN, 'prodDrillInto');

    renderResumo(rsQtd,'Categoria','tblProd','pgProd','prod');
  } else {
    // ── Visão detalhada: produtos de UMA categoria ──
    let rsQtdDrill = rsQtd.filter(r=>r.Categoria===PROD_DRILL);
    let rsValDrill = rsVal.filter(r=>r.Categoria===PROD_DRILL);

    id('prod-breadcrumb').style.display = 'flex';
    id('prod-drill-hint').style.display = 'none';
    id('prod-breadcrumb-cat').textContent = PROD_DRILL;
    id('prodPageTitle').textContent = PROD_DRILL;
    id('prodPageSub').textContent = 'Produtos da categoria ' + PROD_DRILL;
    id('hPupQ').textContent = 'Upgrade — Quantidade';
    id('hPdnQ').textContent = 'Downgrade — Quantidade';
    id('hPupV').textContent = 'Upgrade — Valor';
    id('hPdnV').textContent = 'Downgrade — Valor';
    id('tblProdTitle').textContent = 'Resumo por Produto — ' + PROD_DRILL;

    barRowsDrill('bPupQ', grpCnt(rsQtdDrill,'Produto_Curto','Upgrade'),   false, CUP, null);
    barRowsDrill('bPdnQ', grpCnt(rsQtdDrill,'Produto_Curto','Downgrade'), false, CDN, null);
    barRowsDrill('bPupV', grp(rsValDrill,'Produto_Curto','total_up_RS','Upgrade'),     true, CUP, null);
    barRowsDrill('bPdnV', grp(rsValDrill,'Produto_Curto','total_down_RS','Downgrade'), true, CDN, null);

    renderResumo(rsQtdDrill,'produtos_digisac','tblProd','pgProd','prod');
  }
}

function prodDrillInto(cat){
  PROD_DRILL = cat;
  CP.prod = 1;
  renderProd();
}

function prodDrillBack(){
  PROD_DRILL = null;
  CP.prod = 1;
  renderProd();
}

function renderDown(){
  let rs=fRows('down').filter(r=>r.Movimento==='Downgrade');
  let g=grpCnt(rs,'Situacao_Down');
  let tot=rs.length||1;
  id('barsSit').innerHTML=g.map((d,i)=>`
    <div class="bar-row">
      <div class="bar-lbl" title="${esc(d.key)}">${esc(d.key)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.round(d.value/tot*100)}%;background:${SCOL[i%SCOL.length]}"></div></div>
      <div class="bar-cnt">${fmt(d.value)} <span style="color:#94A3B8;font-size:11px">${pct(d.value/tot*100)}</span></div>
    </div>`).join('');
  let sit=top1(g), prod=top1(grpCnt(rs,'Produto_Curto')), val=top1(grp(rs,'produtos_digisac','total_down_RS'));
  id('downIns').innerHTML=`
    <div class="ic warn" style="margin-bottom:12px">
      <div class="il">Situação mais recorrente</div>
      <div class="iv" style="font-size:16px">${esc(sit.key)}</div>
      <div class="id">${fmt(sit.value)} casos — ${pct(sit.value/tot*100)} dos downgrades</div>
    </div>
    <div class="ic-note">
      <strong>${fmt(tot)}</strong> downgrades no recorte.<br>
      Produto mais afetado: <strong>${esc(prod.key)}</strong> (${fmt(prod.value)} serviços).<br>
      Maior impacto financeiro: <strong>${esc(val.key)}</strong> — ${brl(val.value)}.
    </div>`;
  renderDownTbl(rs);
}

function renderDownTbl(rs){
  let data=rs.slice().sort((a,b)=>String(b.AnoMes).localeCompare(String(a.AnoMes)));
  let tp=Math.max(1,Math.ceil(data.length/PS));
  CP.down=Math.min(CP.down,tp);
  let sl=data.slice((CP.down-1)*PS,CP.down*PS);
  id('tblDown').innerHTML=`<table class="dt"><thead><tr>
    <th style="min-width:88px">Data</th><th style="min-width:180px">Produto</th>
    <th style="min-width:120px">Proprietário</th><th style="min-width:140px">Situação</th>
    <th style="min-width:160px">Motivo</th><th style="min-width:240px">Observação</th>
    <th style="min-width:55px" class="r">Qtd.</th><th style="min-width:95px" class="r">Valor Down</th>
  </tr></thead><tbody>${sl.map(r=>`<tr>
    <td>${esc(r.Data)}</td><td><strong>${esc(r.produtos_digisac)}</strong></td>
    <td>${esc(propNome(r.proprietario))}</td>
    <td><span class="pill p-dn">${esc(r.Situacao_Down)}</span></td>
    <td>${esc(r.motivo_downgrade)}</td><td style="color:#64748B">${esc(r.descricao)}</td>
    <td class="r">${fmt(r.qtd_servicos)}</td><td class="r">${brl(r.total_down_RS)}</td>
  </tr>`).join('')}</tbody></table>`;
  renderPager('pgDown','down',tp,CP.down);
  window._dl=data;
}

function renderProp(){
  let rs=fRows('prop').filter(r=>r.Proprietario_Nome_Valido);
  barRows('bOupQ',grpCnt(rs,'proprietario_filtro','Upgrade'),false,CUP);
  barRows('bOdnQ',grpCnt(rs,'proprietario_filtro','Downgrade'),false,CDN);
  barRows('bOupV',grp(rs,'proprietario_filtro','total_up_RS','Upgrade'),true,CUP);
  barRows('bOdnV',grp(rs,'proprietario_filtro','total_down_RS','Downgrade'),true,CDN);
  renderResumo(rs,'proprietario_filtro','tblProp','pgProp','prop');
}

function renderResumo(rs,field,divId,pagerId,pageKey){
  let keys=[...new Set(rs.map(r=>r[field]))].sort();
  let data=keys.map(k=>{
    let r=rs.filter(x=>x[field]===k);
    return{nome:k,up:sumM(r,'Upgrade'),dn:sumM(r,'Downgrade'),url:sumM(r,'Criação de URL'),mig:sumM(r,'Migração'),tst:sumM(r,'Teste'),plan:sumM(r,'Alteração de Plano'),rob:sumM(r,'Robô'),tot:r.length,vup:sum(r.filter(x=>x.Movimento==='Upgrade'),'total_up_RS'),vdn:sum(r.filter(x=>x.Movimento==='Downgrade'),'total_down_RS')};
  }).filter(r=>r.tot>0).sort((a,b)=>b.tot-a.tot);
  let tp=Math.max(1,Math.ceil(data.length/PS));
  CP[pageKey]=Math.min(CP[pageKey],tp);
  let sl=data.slice((CP[pageKey]-1)*PS,CP[pageKey]*PS);
  let lbl=field==='produtos_digisac'?'Produto':'Proprietário';
  id(divId).innerHTML=`<table class="dt"><thead><tr>
    <th style="min-width:200px">${lbl}</th>
    <th style="min-width:65px" class="r">Up</th><th style="min-width:65px" class="r">Down</th>
    <th style="min-width:75px" class="r">Migração</th><th style="min-width:65px" class="r">Teste</th>
    <th style="min-width:75px" class="r">Desconto</th><th style="min-width:65px" class="r">Robô</th>
    <th style="min-width:65px" class="r">Total</th>
    <th style="min-width:100px" class="r">Valor Up</th><th style="min-width:100px" class="r">Valor Down</th>
  </tr></thead><tbody>${sl.map(r=>`<tr>
    <td><strong>${esc(r.nome)}</strong></td>
    <td class="r" style="color:#059669">${r.up?fmt(r.up):'—'}</td>
    <td class="r" style="color:#DC2626">${r.dn?fmt(r.dn):'—'}</td>
    <td class="r">${r.mig?fmt(r.mig):'—'}</td><td class="r">${r.tst?fmt(r.tst):'—'}</td>
    <td class="r">${r.dsc?fmt(r.dsc):'—'}</td><td class="r">${r.rob?fmt(r.rob):'—'}</td>
    <td class="r"><strong>${fmt(r.tot)}</strong></td>
    <td class="r">${brl(r.vup)}</td><td class="r">${brl(r.vdn)}</td>
  </tr>`).join('')}</tbody></table>`;
  renderPager(pagerId,pageKey,tp,CP[pageKey]);
}

function renderAna(){
  let rs=fRows('ana').sort((a,b)=>String(b.AnoMes).localeCompare(String(a.AnoMes)));
  id('anaCnt').textContent=fmt(rs.length)+' registro'+(rs.length!==1?'s':'');
  let tp=Math.max(1,Math.ceil(rs.length/PS));
  CP.ana=Math.min(CP.ana,tp);
  let sl=rs.slice((CP.ana-1)*PS,CP.ana*PS);
  id('tblAna').innerHTML=`<table class="dt"><thead><tr>
    <th style="min-width:88px">Data</th><th style="min-width:100px">Movimento</th>
    <th style="min-width:150px">Faixa de Vida</th>
    <th style="min-width:70px" class="r">Dias</th>
    <th style="min-width:150px">Subdomínio</th>
    <th style="min-width:180px">Produto</th><th style="min-width:150px">Proprietário</th>
    <th style="min-width:160px">Motivo Down</th><th style="min-width:260px">Observação</th>
    <th style="min-width:55px" class="r">Qtd.</th>
    <th style="min-width:100px" class="r">Valor Up</th><th style="min-width:110px" class="r">Valor Down</th>
  </tr></thead><tbody>${sl.map(r=>`<tr>
    <td>${esc(r.Data)}</td>
    <td><span class="pill ${r.Movimento==='Upgrade'?'p-up':r.Movimento==='Downgrade'?'p-dn':r.Movimento==='Migração'?'p-mig':r.Movimento==='Teste'?'p-tst':r.Movimento==='Criação de URL'?'p-url':r.Movimento==='Robô'?'p-rob':r.Movimento==='Alteração de Plano'?'p-dsc':'p-oth'}">${esc(r.Movimento)}</span></td>
    <td><span style="font-size:11px;font-weight:700;background:#EFF6FF;color:#1D4ED8;padding:3px 8px;border-radius:999px">${esc(r.Faixa_Vida||'—')}</span></td>
    <td class="r" style="color:#64748B">${r.Tempo_vida_cli!=null?fmt(r.Tempo_vida_cli):'—'}</td>
    <td>${esc(r.subdominio)}</td>
    <td><strong>${esc(r.produtos_digisac)}</strong></td>
    <td>${esc(propNome(r.proprietario))}</td>
    <td>${esc(r.motivo_downgrade)}</td>
    <td style="color:#64748B">${esc(r.descricao)}</td>
    <td class="r">${fmt(r.qtd_servicos)}</td>
    <td class="r" style="color:#059669">${r.total_up_RS?brl(r.total_up_RS):'—'}</td>
    <td class="r" style="color:#DC2626">${r.total_down_RS?brl(r.total_down_RS):'—'}</td>
  </tr>`).join('')}</tbody></table>`;
  renderPager('pgAna','ana',tp,CP.ana);
  window._ar=rs; prepCSV(rs);
}

function renderPager(divId,key,tp,cur){
  id(divId).innerHTML=`
    <button class="btn" onclick="movePg('${key}',-1)"${cur<=1?' disabled':''}>← Anterior</button>
    <span>Página ${cur} de ${tp||1}</span>
    <button class="btn" onclick="movePg('${key}',1)"${cur>=tp?' disabled':''}>Próxima →</button>`;
}
function movePg(k,d){CP[k]+=d; if(CP[k]<1)CP[k]=1; renderAll();}

const ACOLS=['Data','Ano','Mês','Movimento','Faixa_Vida','Tempo_vida_cli','subdominio','proprietario','produtos_digisac','produtos_digisac_original','tipo_servico','Situacao_Down','motivo_downgrade','descricao','qtd_servicos','total_up_RS','total_down_RS'];
function toCSV(rs,cols){return cols.join(';')+'\n'+rs.map(r=>cols.map(c=>'"'+String(r[c]??'').replace(/"/g,'""').replace(/\n/g,' ')+'"').join(';')).join('\n')}
function dlCSV(rs,cols,name){let b=new Blob(['\ufeff'+toCSV(rs,cols)],{type:'text/csv;charset=utf-8;'});let u=URL.createObjectURL(b);let a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(u)}
function expProd(){dlCSV(fRows('prod'),['Data','Movimento','Faixa_Vida','Tempo_vida_cli','subdominio','produtos_digisac','produtos_digisac_original','proprietario','qtd_servicos','total_up_RS','total_down_RS','motivo_downgrade','descricao'],'produtos_upxdown.csv')}
function expDown(){dlCSV(fRows('down').filter(r=>r.Movimento==='Downgrade'),['Data','Ano','Mês','Faixa_Vida','Tempo_vida_cli','subdominio','proprietario','produtos_digisac','produtos_digisac_original','Situacao_Down','motivo_downgrade','descricao','qtd_servicos','total_down_RS'],'downgrades.csv')}
function expProp(){dlCSV(fRows('prop').filter(r=>r.Proprietario_Nome_Valido),['Data','Movimento','Faixa_Vida','Tempo_vida_cli','subdominio','proprietario','produtos_digisac','qtd_servicos','total_up_RS','total_down_RS'],'proprietarios.csv')}
function expAna(){dlCSV(window._ar||[],ACOLS,'base_analitica.csv')}
function prepCSV(rs){id('csvBox').value=toCSV(rs,ACOLS)}
function copyCSV(){navigator.clipboard.writeText(id('csvBox').value).then(()=>alert('CSV copiado!'))}
function toggleCSV(){let b=id('csvBox');b.style.display=b.style.display==='block'?'none':'block'}

function showPage(p){
  const destino = id('page-'+p);
  if(destino) destino.scrollIntoView({behavior:'smooth', block:'start'});
}

function renderAll(){
  renderExec();
  renderProd();
  renderDown();
  renderProp();
  renderAna();
}

initFilters();
renderAll();
</script>
</body>
</html>'''

html = html.replace('__PAYLOAD_JSON__', payload_json)
html = html.replace('__DT_ATUALIZACAO__', dt_atualizacao)

ARQUIVO_SAIDA.write_text(html, encoding='utf-8')
print("✅ Gerado com sucesso!")
print(f"📄 {ARQUIVO_SAIDA}")
print(f"📊 {ARQUIVO_BASE_CSV}")
print(f"🗑️  {PASTA / 'linhas_excluidas_sem_movimento.csv'} (linhas descartadas, se houver)")
