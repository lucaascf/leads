import pandas as pd
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
    for _, row in df_cnae.iterrows():
        codigo_limpo = re.sub(r"\D", "", str(row["codigo"])).zfill(7)
        descricoes_cnae[codigo_limpo] = row["descricao"].replace(';', ',').replace('"', "'")
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
# FUNÇÕES OTIMIZADAS
# ============================================
def limpar_numero(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\D", "", str(x).replace(".0", ""))

def limpar_numero_series(s):
    return s.astype(str).str.replace(r"\D", "", regex=True).str.replace(".0", "", regex=False)

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

# ============================================
# PASSO 2: COLETAR CNPJs (OTIMIZADO)
# ============================================
print("\n🔍 [2/4] Coletando CNPJs dos estabelecimentos...")

if os.path.exists(checkpoint_file):
    print("  🔄 Checkpoint encontrado! Continuando...")
    with open(checkpoint_file, 'rb') as f:
        checkpoint = pickle.load(f)
        cnpjs_necessarios = checkpoint['cnpjs']
        total_estabelecimentos = checkpoint['total_leads']
        chunks_processados = checkpoint['chunks']
    print(f"  ✅ Retomando do chunk {chunks_processados}")
else:
    cnpjs_necessarios = set()
    total_estabelecimentos = 0
    chunks_processados = 0

chunk_num = 0

for chunk in pd.read_csv(
    "K3241.K03200Y0.D60314.ESTABELE",
    sep=";",
    encoding="latin1",
    header=None,
    engine="python",
    on_bad_lines="skip",
    chunksize=1000000,
    usecols=IDX_USADAS
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
            chunk["TELEFONE1"].notna()
        ) &
        (chunk["SITUACAO_ESPECIAL"].isna())
    )
    
    filtrado = chunk[mask]
    total_estabelecimentos += len(filtrado)
    
    if len(filtrado) > 0:
        cnpjs_limpos = filtrado["CNPJ_BASICO"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
        cnpjs_necessarios.update(cnpjs_limpos.unique())
    
    if chunk_num % 10 == 0:
        with open(checkpoint_file, 'wb') as f:
            pickle.dump({
                'cnpjs': cnpjs_necessarios,
                'total_leads': total_estabelecimentos,
                'chunks': chunk_num
            }, f)
        print(f"  💾 Chunk {chunk_num}: {len(cnpjs_necessarios):,} CNPJs | {total_estabelecimentos:,} leads")

if chunk_num > 0 and chunk_num % 10 != 0:
    with open(checkpoint_file, 'wb') as f:
        pickle.dump({
            'cnpjs': cnpjs_necessarios,
            'total_leads': total_estabelecimentos,
            'chunks': chunk_num
        }, f)
    print(f"  💾 Checkpoint final: Chunk {chunk_num} | {len(cnpjs_necessarios):,} CNPJs | {total_estabelecimentos:,} leads")

print(f"\n  ✅ {len(cnpjs_necessarios):,} CNPJs únicos coletados")
print(f"  ✅ {total_estabelecimentos:,} leads potenciais")

if os.path.exists(checkpoint_file):
    os.remove(checkpoint_file)

# ============================================
# PASSO 3: BUSCAR DADOS DAS EMPRESAS
# ============================================
print("\n📂 [3/4] Buscando dados das empresas...")

cnpjs_para_buscar = frozenset(cnpjs_necessarios)
dados_empresas = {}
chunk_emp_num = 0

try:
    for chunk_emp in pd.read_csv(
        empresas_file,
        sep=";",
        encoding="latin1",
        header=None,
        engine="python",
        on_bad_lines="skip",
        chunksize=1000000,
        usecols=[0, 4, 5]
    ):
        if len(chunk_emp) == 0:
            continue
        
        for idx in range(len(chunk_emp)):
            try:
                cnpj_base = str(chunk_emp.iloc[idx, 0]).strip().zfill(8)
                
                if cnpj_base in cnpjs_para_buscar:
                    capital = chunk_emp.iloc[idx, 1] if pd.notna(chunk_emp.iloc[idx, 1]) else ""
                    porte_codigo = str(chunk_emp.iloc[idx, 2]) if pd.notna(chunk_emp.iloc[idx, 2]) else "0"
                    
                    if porte_codigo == "1":
                        porte_desc = "MEI"
                    elif porte_codigo == "3":
                        porte_desc = "ME/EPP"
                    elif porte_codigo == "5":
                        porte_desc = "DEMAIS"
                    else:
                        porte_desc = "NÃO INFORMADO"
                    
                    dados_empresas[cnpj_base] = {
                        "CAPITAL_SOCIAL": capital,
                        "PORTE": porte_desc
                    }
            except Exception as e:
                continue
        
        chunk_emp_num += 1
        if chunk_emp_num % 10 == 0:
            print(f"  Chunk {chunk_emp_num}: {len(dados_empresas):,}/{len(cnpjs_necessarios):,} encontrados")
            
except Exception as e:
    print(f"  ⚠️ Erro ao processar empresas: {e}")

print(f"\n  ✅ Dados encontrados para {len(dados_empresas):,} de {len(cnpjs_necessarios):,} empresas")

# ============================================
# FUNÇÕES AUXILIARES
# ============================================
def formatar_cnpj(cnpj):
    cnpj = limpar_numero(cnpj).zfill(14)
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

def formatar_cnae(cnae):
    cnae = limpar_numero(cnae).zfill(7)
    return f"{cnae[:2]}.{cnae[2:4]}-{cnae[4]}/{cnae[5:]}"

def formatar_capital(capital):
    if pd.isna(capital) or capital == "":
        return ""
    try:
        valor = float(str(capital).replace(',', '.'))
        return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return str(capital)

def montar_tel(ddd, num):
    ddd = limpar_numero(ddd)
    num = limpar_numero(num)

    tel = ddd + num

    # 🔥 BLOQUEIO REAL (resolve seu problema)
    if not tel or set(tel) == {"0"}:
        return ""

    if len(tel) == 11:
        return f"({tel[:2]}) {tel[2:7]}-{tel[7:]}"
    elif len(tel) == 10:
        return f"({tel[:2]}) {tel[2:6]}-{tel[6:]}"

    return ""

def limpar_texto(t):
    if pd.isna(t):
        return ""
    return str(t).replace(';', ',').replace('"', "'").replace('\n', ' ')

def formatar_cep(valor):
    try:
        cep_str = str(valor).replace('.0', '').replace(',', '')
        cep_limpo = re.sub(r"\D", "", cep_str)
        if len(cep_limpo) == 8:
            return f"{cep_limpo[:5]}-{cep_limpo[5:]}"
        return ""
    except:
        return ""
    
def limpar_cnae_sec(x):
    if pd.isna(x):
        return ""
    cnaes = [re.sub(r"\D", "", cnae).zfill(7) for cnae in str(x).split(",")]
    return ",".join(filter(None, cnaes))

situacao_map = {"01": "NULA", "2": "ATIVA", "3": "SUSPENSA", "4": "INAPTA", "08": "BAIXADA"}

# ============================================
# PASSO 4: GERAR CSV (COM CHECKPOINT SEGURO)
# ============================================
print("\n💾 [4/4] Gerando arquivo final...")

checkpoint_final = "checkpoint_final.pkl"
primeiro = True
total_geral = 0
chunk_num = 0

if os.path.exists(checkpoint_final):
    print("  🔄 Checkpoint encontrado! Continuando...")
    try:
        with open(checkpoint_final, 'rb') as f:
            checkpoint = pickle.load(f)
            chunk_num = checkpoint['chunk']
            total_geral = checkpoint['total']
            primeiro = False
        print(f"  ✅ Retomando do chunk {chunk_num} com {total_geral:,} leads")
    except:
        print("  ⚠️ Checkpoint corrompido, começando do zero")
        chunk_num = 0
        total_geral = 0
        primeiro = True

chunk_atual = 0

for chunk in pd.read_csv(
    "K3241.K03200Y0.D60314.ESTABELE",
    sep=";",
    encoding="latin1",
    header=None,
    engine="python",
    on_bad_lines="skip",
    chunksize=1000000,
    usecols=IDX_USADAS
):
    chunk_atual += 1
    
    if chunk_atual <= chunk_num:
        continue
    
    print(f"  Processando chunk {chunk_atual}...")
    
    chunk.columns = [COLUNAS[i] for i in IDX_USADAS]
    
    chunk["DATA_INICIO_ATIVIDADE"] = pd.to_datetime(chunk["DATA_INICIO_ATIVIDADE"], format="%Y%m%d", errors="coerce")
    
    mask = (
        (chunk["DATA_INICIO_ATIVIDADE"] >= limite) &
        (chunk["NOME_FANTASIA"].notna()) &
        (chunk["UF"].notna()) &
        (chunk["SITUACAO_CADASTRAL"] == 2) &
        (
            chunk["EMAIL"].notna() |
            chunk["TELEFONE1"].notna()
        ) &
        (chunk["SITUACAO_ESPECIAL"].isna())
    )
    
    dados = chunk[mask].copy()
    
    if not dados.empty:
        try:
            cnpj_base = dados["CNPJ_BASICO"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
            cnpj_completo = (cnpj_base + 
                            dados["CNPJ_ORDEM"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(4) + 
                            dados["CNPJ_DV"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(2))
            
            porte = cnpj_base.apply(lambda x: dados_empresas.get(x, {}).get("PORTE", "NÃO INFORMADO"))
            capital = cnpj_base.apply(lambda x: dados_empresas.get(x, {}).get("CAPITAL_SOCIAL", ""))
            
            cep_formatado = dados["CEP"].apply(formatar_cep)
            
            if municipios_dict:
                municipio = dados.apply(
                    lambda row: buscar_municipio(row["UF"], row["COD_MUNICIPIO"]), 
                    axis=1
                )
            else:
                municipio = ""
            
            final = pd.DataFrame({
                "CNPJ": cnpj_completo.apply(formatar_cnpj),
                "IDENTIFICADOR_MATRIZ_FILIAL": dados["IDENTIFICADOR_MATRIZ_FILIAL"].apply(
                    lambda x: "MATRIZ" if str(x).strip() == "1" else ("FILIAL" if str(x).strip() == "2" else limpar_texto(x))
                ),
                "NOME": dados["NOME_FANTASIA"].apply(limpar_texto),
                "DATA_ABERTURA": dados["DATA_INICIO_ATIVIDADE"],
                "CNAE": dados["CNAE_PRINCIPAL"].apply(formatar_cnae),
                "CNAE_FISCAL_SECUNDARIA": dados["CNAE_SECUNDARIO"].apply(limpar_cnae_sec),
                "DESCRICAO_CNAE": dados["CNAE_PRINCIPAL"].apply(lambda x: descricoes_cnae.get(limpar_numero(x).zfill(7), "Descrição não encontrada")),
                "PORTE": porte,
                "CAPITAL_SOCIAL": capital.apply(formatar_capital),
                "SITUACAO_CADASTRAL": dados["SITUACAO_CADASTRAL"].fillna("").astype(str).apply(lambda x: situacao_map.get(x, x if x else "NÃO INFORMADO")),
                "EMAIL": dados["EMAIL"].apply(limpar_texto),
                "TELEFONE1": dados.apply(lambda x: montar_tel(x["DDD1"], x["TELEFONE1"]), axis=1),
                "TELEFONE2": dados.apply(lambda x: montar_tel(x["DDD2"], x["TELEFONE2"]), axis=1),
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
            print(f"  ⚠️ Erro no chunk {chunk_atual}: {e}")
    
    if chunk_atual % 5 == 0:
        with open(checkpoint_final, 'wb') as f:
            pickle.dump({'chunk': chunk_atual, 'total': total_geral}, f)
        print(f"  💾 Checkpoint salvo: chunk {chunk_atual} | {total_geral:,} leads")

if chunk_atual > 0:
    with open(checkpoint_final, 'wb') as f:
        pickle.dump({'chunk': chunk_atual, 'total': total_geral}, f)
    print(f"  💾 Checkpoint final salvo: {total_geral:,} leads")

print("\n" + "="*60)
print(f"✅ CONCLUÍDO!")
print(f"📊 Total de leads gerados: {total_geral}")
print(f"📁 Arquivo salvo: {output_file}")
print("="*60)