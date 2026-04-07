import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import os
import pickle

hoje = datetime.today()
limite = hoje - timedelta(days=30)

output_file = "leads_filtrados.csv"
cnae_file = "cnaes.csv"
empresas_file = "K3241.K03200Y0.D60314.EMPRECSV"
municipios_file = "Municípios.xlsx"
checkpoint_file = "checkpoint.pkl"

print("="*60)
print("🚀 INICIANDO PROCESSAMENTO DE LEADS (VERSÃO OTIMIZADA)")
print("="*60)

# ============================================
# MAPEAMENTO DE COLUNAS (NOVO)
# ============================================
COLUNAS = {
    0: "CNPJ_BASICO",
    1: "CNPJ_ORDEM",
    2: "CNPJ_DV",
    3: "IDENTIFICADOR_MATRIZ_FILIAL",
    4: "NOME_FANTASIA",
    5: "SITUACAO_CADASTRAL",
    6: "DATA_SITUACAO_CADASTRAL",
    7: "MOTIVO_SITUACAO_CADASTRAL",
    8: "NOME_CIDADE_EXTERIOR",
    9: "PAIS",
    10: "DATA_INICIO_ATIVIDADE",
    11: "CNAE_PRINCIPAL",
    12: "CNAE_SECUNDARIO",
    13: "TIPO_LOGRADOURO",
    14: "LOGRADOURO",
    15: "NUMERO",
    16: "COMPLEMENTO",
    17: "BAIRRO",
    18: "CEP",
    19: "UF",
    20: "COD_MUNICIPIO",
    21: "DDD1",
    22: "TELEFONE1",
    23: "DDD2",
    24: "TELEFONE2",
    25: "DDD_FAX",
    26: "FAX",
    27: "EMAIL",
    28: "SITUACAO_ESPECIAL",
    29: "DATA_SITUACAO_ESPECIAL"
}

# CORREÇÃO: Adicionadas FAX e DDD_FAX
COLS_USADAS = [
    "CNPJ_BASICO", "CNPJ_ORDEM", "CNPJ_DV",
    "IDENTIFICADOR_MATRIZ_FILIAL",
    "NOME_FANTASIA",
    "SITUACAO_CADASTRAL",
    "DATA_INICIO_ATIVIDADE",
    "CNAE_PRINCIPAL", "CNAE_SECUNDARIO",
    "TIPO_LOGRADOURO", "LOGRADOURO", "NUMERO",
    "COMPLEMENTO", "BAIRRO", "CEP",
    "UF", "COD_MUNICIPIO",
    "DDD1", "TELEFONE1",
    "DDD2", "TELEFONE2",
    "DDD_FAX", "FAX",  # CORREÇÃO: Adicionadas estas colunas
    "EMAIL",
    "SITUACAO_ESPECIAL"
]

IDX_USADAS = [i for i, nome in COLUNAS.items() if nome in COLS_USADAS]

# ============================================
# PASSO 1: CARREGAR DESCRIÇÕES DE CNAE
# ============================================
print("\n📚 [1/4] Carregando descrições de CNAE...")
descricoes_cnae = {}
try:
    df_cnae = pd.read_csv(cnae_file, sep=";", header=None, names=["codigo", "descricao"])
    codigo_limpo = df_cnae["codigo"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(7)
    descricao_limpa = df_cnae["descricao"].str.replace(';', ',', regex=False).str.replace('"', "'", regex=False)
    descricoes_cnae = dict(zip(codigo_limpo, descricao_limpa))
    print(f"  ✅ {len(descricoes_cnae)} descrições carregadas")
except Exception as e:
    print(f"  ⚠️ Erro: {e}")

# ============================================
# PASSO 1.5: CARREGAR MUNICÍPIOS DO EXCEL
# ============================================
print("\n🏙️ [1.5/4] Carregando dados de municípios do Excel...")
municipios_dict = {}

try:
    df_municipios = pd.read_excel(
        municipios_file,
        engine='openpyxl',
        dtype=str,
        header=None
    )

    print(f"  ✅ Arquivo carregado com {len(df_municipios)} linhas e {df_municipios.shape[1]} colunas")
    print(f"  Primeiras 5 linhas:\n{df_municipios.head()}")

    df_municipios.columns = ['COD_MUNICIPIO', 'MUNICIPIO', 'UF']
    df_municipios = df_municipios.dropna(subset=['COD_MUNICIPIO', 'MUNICIPIO', 'UF'])

    df_municipios['COD_MUNICIPIO'] = df_municipios['COD_MUNICIPIO'].astype(str).str.strip().str.zfill(5)
    df_municipios['UF'] = df_municipios['UF'].astype(str).str.strip().str.upper()
    df_municipios['MUNICIPIO'] = df_municipios['MUNICIPIO'].astype(str).str.strip()

    df_municipios['CHAVE_MUNICIPIO'] = df_municipios['UF'] + '_' + df_municipios['COD_MUNICIPIO']
    municipios_dict = dict(zip(df_municipios['CHAVE_MUNICIPIO'], df_municipios['MUNICIPIO']))

    print(f"  ✅ {len(municipios_dict):,} municípios carregados com sucesso")
    print(f"  Exemplo: RO_{'0000001'.zfill(7)} -> {municipios_dict.get('RO_' + '0000001'.zfill(7), 'N/A')}")

except Exception as e:
    print(f"  ⚠️ Erro ao carregar municípios: {e}")
    print(f"  ⚠️ Continuando sem a coluna de município...")
    municipios_dict = {}

# ============================================
# FUNÇÕES AUXILIARES
# ============================================
def limpar_numero(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\D", "", str(x).replace(".0", ""))

def buscar_municipio(uf, cod_municipio):
    if pd.isna(uf) or pd.isna(cod_municipio):
        return ""
    try:
        uf_limpo = str(uf).strip().upper()
        cod_limpo = str(cod_municipio).strip().zfill(5)
        chave = uf_limpo + '_' + cod_limpo
        return municipios_dict.get(chave, "")
    except:
        return ""

def formatar_capital(capital):
    if pd.isna(capital) or capital == "":
        return ""
    try:
        valor = float(str(capital).replace(',', '.'))
        return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return str(capital)

def montar_tel_vec(ddd_s, num_s):
    """Versão vetorizada de montar_tel."""
    ddd = ddd_s.astype(str).str.replace(r"\D", "", regex=True).str.replace(r"\.0$", "", regex=True)
    num = num_s.astype(str).str.replace(r"\D", "", regex=True).str.replace(r"\.0$", "", regex=True)
    tel = ddd + num
    all_zero = tel.str.fullmatch(r"0*") | (tel == "")
    fmt11 = "(" + tel.str[:2] + ") " + tel.str[2:7] + "-" + tel.str[7:]
    fmt10 = "(" + tel.str[:2] + ") " + tel.str[2:6] + "-" + tel.str[6:]
    result = pd.Series("", index=tel.index)
    result = result.where(tel.str.len() != 11, fmt11)
    result = result.where(tel.str.len() != 10, fmt10)
    result[all_zero] = ""
    return result

def limpar_texto(t):
    if pd.isna(t):
        return ""
    return str(t).replace(';', ',').replace('"', "'").replace('\n', ' ')

def limpar_cnae_sec(x):
    if pd.isna(x):
        return ""
    cnaes = [re.sub(r"\D", "", cnae).zfill(7) for cnae in str(x).split(",")]
    return ",".join(filter(None, cnaes))

situacao_map = {"01": "NULA", "2": "ATIVA", "3": "SUSPENSA", "4": "INAPTA", "08": "BAIXADA"}

# ============================================
# PASSO 2: COLETAR CNPJs + BUFFERAR DADOS FILTRADOS
# ============================================
print("\n🔍 [2/4] Coletando CNPJs dos estabelecimentos...")

if os.path.exists(checkpoint_file):
    print("  🔄 Checkpoint encontrado! Continuando...")
    with open(checkpoint_file, 'rb') as f:
        checkpoint = pickle.load(f)
        cnpjs_necessarios = checkpoint['cnpjs']
        total_estabelecimentos = checkpoint['total_leads']
        chunks_processados = checkpoint['chunks']
        filtered_chunks = checkpoint.get('filtered_chunks', [])
    print(f"  ✅ Retomando do chunk {chunks_processados}")
else:
    cnpjs_necessarios = set()
    total_estabelecimentos = 0
    chunks_processados = 0
    filtered_chunks = []

chunk_num = 0

for chunk in pd.read_csv(
    "K3241.K03200Y0.D60314.ESTABELE",
    sep=";",
    encoding="latin1",
    header=None,
    engine="python",
    on_bad_lines="skip",
    chunksize=1_000_000,
    usecols=IDX_USADAS,
):
    chunk_num += 1

    if chunk_num <= chunks_processados:
        continue

    chunk.columns = [COLUNAS[i] for i in IDX_USADAS]

    chunk["DATA_INICIO_ATIVIDADE"] = pd.to_datetime(chunk["DATA_INICIO_ATIVIDADE"], format="%Y%m%d", errors="coerce")

    mask = (
        (chunk["DATA_INICIO_ATIVIDADE"] >= limite) &
        (chunk["NOME_FANTASIA"].notna()) &
        (chunk["UF"].notna()) &
        (chunk["SITUACAO_CADASTRAL"] == 2) &
        (
            chunk["EMAIL"].notna() |
            chunk["TELEFONE1"].notna() |
            chunk["FAX"].notna()  # CORREÇÃO: Agora FAX existe
        ) &
        (chunk["SITUACAO_ESPECIAL"].isna())
    )

    filtrado = chunk[mask]
    total_estabelecimentos += len(filtrado)

    if len(filtrado) > 0:
        filtered_chunks.append(filtrado.copy())
        cnpjs_limpos = filtrado["CNPJ_BASICO"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
        cnpjs_validos = [c for c in cnpjs_limpos.unique() if c.isdigit() and len(c) == 8]
        cnpjs_necessarios.update(cnpjs_validos)

    if chunk_num % 20 == 0:
        with open(checkpoint_file, 'wb') as f:
            pickle.dump({
                'cnpjs': cnpjs_necessarios,
                'total_leads': total_estabelecimentos,
                'chunks': chunk_num,
                'filtered_chunks': filtered_chunks
            }, f)
        print(f"  💾 Chunk {chunk_num}: {len(cnpjs_necessarios):,} CNPJs | {total_estabelecimentos:,} leads")

if chunk_num > 0 and chunk_num % 20 != 0:
    with open(checkpoint_file, 'wb') as f:
        pickle.dump({
            'cnpjs': cnpjs_necessarios,
            'total_leads': total_estabelecimentos,
            'chunks': chunk_num,
            'filtered_chunks': filtered_chunks
        }, f)
    print(f"  💾 Checkpoint final: Chunk {chunk_num} | {len(cnpjs_necessarios):,} CNPJs | {total_estabelecimentos:,} leads")

print(f"\n  ✅ {len(cnpjs_necessarios):,} CNPJs únicos coletados")
print(f"  ✅ {total_estabelecimentos:,} leads potenciais")

if os.path.exists(checkpoint_file):
    os.remove(checkpoint_file)

# ============================================
# PASSO 3: BUSCAR DADOS DAS EMPRESAS
# ============================================

print("\n📂 [3/4] Buscando dados das empresas... (versão otimizada)")

cnpjs_para_buscar = frozenset(cnpjs_necessarios)
dados_empresas = {}

def mapear_porte_seguro(porte_codigo):
    if pd.isna(porte_codigo):
        return "NÃO INFORMADO"
    try:
        # int(float(...)) cobre tanto int (1) quanto float (1.0) lidos pelo pandas
        porte_str = str(int(float(porte_codigo))).zfill(2)
    except:
        return "NÃO INFORMADO"
    return {
        "01": "MICRO EMPRESA",
        "03": "EMPRESA DE PEQUENO PORTE",
        "05": "DEMAIS"
    }.get(porte_str, "NÃO INFORMADO")

chunk_emp_num = 0

for chunk_emp in pd.read_csv(
    empresas_file,
    sep=";",
    encoding="latin1",
    header=None,
    engine="c",
    on_bad_lines="skip",
    chunksize=1_000_000,
    usecols=[0, 4, 5]  # CNPJ_BASICO, CAPITAL, PORTE
):
    if len(chunk_emp) == 0:
        continue

    chunk_emp.columns = ["CNPJ_BASICO", "CAPITAL_SOCIAL", "PORTE_BRUTO"]
    chunk_emp["CNPJ_BASICO"] = chunk_emp["CNPJ_BASICO"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
    chunk_emp = chunk_emp[chunk_emp["CNPJ_BASICO"].isin(cnpjs_para_buscar)]

    if not chunk_emp.empty:
        chunk_emp["PORTE"] = chunk_emp["PORTE_BRUTO"].apply(mapear_porte_seguro)
        dados_empresas.update({
            cnpj: {"CAPITAL_SOCIAL": capital, "PORTE": porte}
            for cnpj, capital, porte in zip(
                chunk_emp["CNPJ_BASICO"],
                chunk_emp["CAPITAL_SOCIAL"],
                chunk_emp["PORTE"]
            )
        })

    chunk_emp_num += 1
    if chunk_emp_num % 10 == 0:
        print(f"  Chunk {chunk_emp_num}: {len(dados_empresas):,}/{len(cnpjs_necessarios):,} encontrados")

print(f"\n  ✅ Dados encontrados para {len(dados_empresas):,} de {len(cnpjs_necessarios):,} empresas")

# Dicts planos para lookup vetorizado
porte_map   = {cnpj: v["PORTE"]          for cnpj, v in dados_empresas.items()}
capital_map = {cnpj: v["CAPITAL_SOCIAL"] for cnpj, v in dados_empresas.items()}

# ============================================
# PASSO 4: GERAR CSV A PARTIR DOS DADOS BUFFERIZADOS
# ============================================
print("\n💾 [4/4] Gerando arquivo final...")

primeiro = True
total_geral = 0

if not filtered_chunks:
    print("  ⚠️ Nenhum dado filtrado encontrado.")
else:
    for chunk_idx, dados_raw in enumerate(filtered_chunks, start=1):
        print(f"  Processando bloco {chunk_idx}/{len(filtered_chunks)}...")
        dados = dados_raw.copy()

        if dados.empty:
            continue

        try:
            # CNPJ vetorizado
            cnpj_base  = dados["CNPJ_BASICO"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
            cnpj_ordem = dados["CNPJ_ORDEM"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(4)
            cnpj_dv    = dados["CNPJ_DV"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(2)
            cnpj_14    = (cnpj_base + cnpj_ordem + cnpj_dv).str.zfill(14)
            cnpj_formatado = (
                cnpj_14.str[:2] + "." + cnpj_14.str[2:5] + "." +
                cnpj_14.str[5:8] + "/" + cnpj_14.str[8:12] + "-" + cnpj_14.str[12:]
            )

            # CNAE vetorizado (reutilizado para descrição)
            cnae_s = dados["CNAE_PRINCIPAL"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(7)
            cnae_formatado = (
                cnae_s.str[:2] + "." + cnae_s.str[2:4] + "-" +
                cnae_s.str[4] + "/" + cnae_s.str[5:]
            )

            # CEP vetorizado
            cep_s = dados["CEP"].astype(str).str.replace(r"\.0$", "", regex=True).str.replace(r"\D", "", regex=True)
            cep_formatado = pd.Series(
                np.where(cep_s.str.len() == 8, cep_s.str[:5] + "-" + cep_s.str[5:], ""),
                index=dados.index
            )

            # Porte e capital via .map()
            porte   = cnpj_base.map(porte_map).fillna("NÃO INFORMADO")
            capital = cnpj_base.map(capital_map).fillna("")

            # Município vetorizado
            if municipios_dict:
                chave_mun = (
                    dados["UF"].fillna("").astype(str).str.strip().str.upper()
                    + "_" +
                    dados["COD_MUNICIPIO"].fillna("").astype(str).str.strip().str.zfill(5)
                )
                municipio = chave_mun.map(municipios_dict).fillna("")
            else:
                municipio = pd.Series("", index=dados.index)

            # Situação cadastral vetorizada
            sit_s = dados["SITUACAO_CADASTRAL"].fillna("").astype(str)
            situacao_final = sit_s.map(situacao_map).fillna(sit_s.where(sit_s != "", "NÃO INFORMADO"))

            # Identificador Matriz/Filial vetorizado
            imf_s = dados["IDENTIFICADOR_MATRIZ_FILIAL"].astype(str).str.strip()
            imf_fallback = (
                dados["IDENTIFICADOR_MATRIZ_FILIAL"].fillna("").astype(str)
                .str.replace(";", ",", regex=False)
                .str.replace('"', "'", regex=False)
                .str.replace("\n", " ", regex=False)
            )
            identificador = imf_s.map({"1": "MATRIZ", "2": "FILIAL"}).fillna(imf_fallback)

            final = pd.DataFrame({
                "CNPJ": cnpj_formatado,
                "IDENTIFICADOR_MATRIZ_FILIAL": identificador,
                "NOME": dados["NOME_FANTASIA"].apply(limpar_texto),
                "DATA_ABERTURA": dados["DATA_INICIO_ATIVIDADE"],
                "CNAE": cnae_formatado,
                "CNAE_FISCAL_SECUNDARIA": dados["CNAE_SECUNDARIO"].apply(limpar_cnae_sec),
                "DESCRICAO_CNAE": cnae_s.map(descricoes_cnae).fillna("Descrição não encontrada"),
                "PORTE": porte,
                "CAPITAL_SOCIAL": capital.apply(formatar_capital),
                "SITUACAO_CADASTRAL": situacao_final,
                "EMAIL": dados["EMAIL"].apply(limpar_texto),
                "TELEFONE1": montar_tel_vec(dados["DDD1"], dados["TELEFONE1"]),
                "TELEFONE2": montar_tel_vec(dados["DDD2"], dados["TELEFONE2"]),
                "TIPO_LOGRADOURO": dados["TIPO_LOGRADOURO"].apply(limpar_texto),
                "LOGRADOURO": dados["LOGRADOURO"].apply(limpar_texto),
                "NUMERO": dados["NUMERO"].apply(limpar_texto),
                "COMPLEMENTO": dados["COMPLEMENTO"].apply(limpar_texto),
                "BAIRRO": dados["BAIRRO"].apply(limpar_texto),
                "CEP": cep_formatado,
                "UF": dados["UF"].apply(limpar_texto),
                "MUNICIPIO": municipio
            })

            final.to_csv(output_file, mode="a", header=primeiro, index=False, sep=";", quoting=1)
            primeiro = False
            total_geral += len(final)

        except Exception as e:
            print(f"  ⚠️ Erro no bloco {chunk_idx}: {e}")

print("\n" + "="*60)
print(f"✅ CONCLUÍDO!")
print(f"📊 Total de leads gerados: {total_geral}")
print(f"📁 Arquivo salvo: {output_file}")
print("="*60)
