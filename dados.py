"""
MÓDULO DE DADOS — Carregamento, validação e processamento
Fonte centralizada para todas as abas da dashboard.

Regras de fonte:
- RELATORIO: cadastro de patrimônios com NF (foto atual, 1 linha por patrimônio)
- OS: histórico de todas as ordens de serviço (fonte temporal por período)
- CONTRATOS + config: parque total da rede (inclui equipamentos sem NF)
- BASE_CRUZADA: cruzamento consolidado (CONTRATOS + OS + classificação)
"""

import streamlit as st
import pandas as pd
import os

# Caminho do arquivo de dados (na raiz do projeto)
DATA_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(DATA_DIR, 'NFxPRODUTO__1_.xlsx')

# ============================================================
# PADRONIZAÇÃO CANÔNICA DE ASSUNTOS
# Mapeia valores antigos/variantes → nome padrão usado nas análises
# ============================================================

PADRONIZACAO_ASSUNTO = {
    'INSTALACAO INTERNET': 'INSTALACAO',
    'Instalação Internet (descontinuado)': 'INSTALACAO',
    'Instalação Repetidor Wirelles': 'MESH',
    'Instalação Serviço Cortesia': 'INSTALACAO CORTESIA',
    'Instalação de Telefone': 'INSTALACAO TELEFONE',
    'MANUTENCAO DE REDE': 'MANUTENCAO',
    'MANUTENÇÃO TÉCNICA': 'MANUTENCAO',
    'MUDANÇA DE ENDEREÇO': 'MUDANCA ENDEREÇO',
    'SERVIÇOS TÉCNICOS DIVERSOS': 'MESH',
    'UPGRADE - EQUIPAMENTO': 'UPGRADE',
    '0.1.4 RETIRADA DE REPETIDOR WIRELESS': 'RETIRADA REPETIDOR',
    '0.1.5 RETIRADA ORDEM DE COLETA': 'RETIRADA COLETA',
    '0.1.6 RETIRADA PONTO DE INTERNET': 'RETIRADA DE PONTO',
    'RETIRADA ORDEM DE COLETA': 'RETIRADA COLETA',
    'RETIRADA PONTO DE INTERNET': 'RETIRADA DE PONTO',
}

# Mapeamento DE→PARA completo para a aba config (inclui itens ignorados)
MAPEAMENTO_DE_PARA = {
    **PADRONIZACAO_ASSUNTO,
    'Emitir Taxa': '****',
    'POS VENDA (BRASILIA)': '****',
}


def fmt(numero):
    """Formata número inteiro com ponto como separador de milhares"""
    return f"{int(numero):,}".replace(",", ".")


@st.cache_data
def load_data():
    """Carrega e processa todos os dados da planilha Excel.

    Retorna dict com DataFrames prontos para uso.
    Inclui validações e mensagens de erro claras.
    """
    if not os.path.exists(DATA_FILE):
        st.error(f"Arquivo de dados não encontrado: {DATA_FILE}")
        st.stop()

    try:
        # --- Carregar abas ---
        notas = pd.read_excel(DATA_FILE, sheet_name='NOTAS')
        os_df = pd.read_excel(DATA_FILE, sheet_name='OS')
        relatorio = pd.read_excel(DATA_FILE, sheet_name='RELATORIO')
        contratos = pd.read_excel(DATA_FILE, sheet_name='CONTRATOS')
        config = pd.read_excel(DATA_FILE, sheet_name='config')

        # Abas opcionais (podem não existir ou estar vazias)
        try:
            obsoletos = pd.read_excel(DATA_FILE, sheet_name='OBSOLETOS')
        except Exception:
            obsoletos = pd.DataFrame()

        try:
            reaproveitados = pd.read_excel(DATA_FILE, sheet_name='REAPROVEITADOS')
        except Exception:
            reaproveitados = pd.DataFrame()

        try:
            negativado = pd.read_excel(DATA_FILE, sheet_name='NEGATIVADO')
        except Exception:
            negativado = pd.DataFrame()

        try:
            base_cruzada_raw = pd.read_excel(DATA_FILE, sheet_name='BASE_CRUZADA', header=None)
            base_cruzada = _processar_base_cruzada(base_cruzada_raw)
        except Exception:
            base_cruzada = pd.DataFrame()

        try:
            retirada = pd.read_excel(DATA_FILE, sheet_name='RETIRADA')
        except Exception:
            retirada = pd.DataFrame()

    except Exception as e:
        st.error(f"Erro ao carregar planilha: {e}")
        st.stop()

    # --- Validações de colunas obrigatórias ---
    _validar_colunas(os_df, 'OS', [
        'ID _Ordem de Serviço', 'data_fechamento_OS', 'id_patrimonio',
        'ASSUNTO PADRONIZADO'
    ])
    _validar_colunas(relatorio, 'RELATORIO', [
        'NF', 'DATA NF', 'PATRIMONIO', 'DESCRICAO', 'ASSUNTO OS',
        'DATA ÚLTIMA OS', 'LOCAL_EQUIPAMENTO'
    ])

    # --- Conversões de data ---
    relatorio['DATA NF'] = pd.to_datetime(relatorio['DATA NF'], errors='coerce')
    relatorio['DATA ÚLTIMA OS'] = pd.to_datetime(relatorio['DATA ÚLTIMA OS'], errors='coerce')
    os_df['data_abertura_OS'] = pd.to_datetime(os_df['data_abertura_OS'], errors='coerce')
    os_df['data_fechamento_OS'] = pd.to_datetime(os_df['data_fechamento_OS'], errors='coerce')

    # --- Normalizar IDs para string ---
    for c in ['NF', 'PATRIMONIO', 'PRODUTO ID']:
        if c in relatorio.columns:
            relatorio[c] = (
                relatorio[c].astype(str).fillna('')
                .str.replace(r'\\,', '', regex=True)
                .str.replace(r'\.0$', '', regex=True)
            )

    # Limpar OS sem patrimônio (guardar contagem antes de remover)
    _os_total_bruto = len(os_df)
    _os_sem_patrimonio = 0
    if 'id_patrimonio' in os_df.columns:
        _os_sem_patrimonio = int(os_df['id_patrimonio'].isna().sum())
        os_df = os_df.dropna(subset=['id_patrimonio'])
        os_df['id_patrimonio'] = (
            os_df['id_patrimonio'].astype(str)
            .str.replace(r'\\,', '', regex=True)
            .str.replace(r'\.0$', '', regex=True)
        )

    # Remover OS duplicadas (mesmo ID de OS, manter primeira ocorrência)
    _os_duplicadas = int(os_df.duplicated(subset=['ID _Ordem de Serviço'], keep='first').sum())
    os_df = os_df.drop_duplicates(subset=['ID _Ordem de Serviço'], keep='first')

    # --- Padronizar ASSUNTO PADRONIZADO ---
    # Re-mapeia valores antigos/variantes para nomenclatura canônica
    if 'ASSUNTO PADRONIZADO' in os_df.columns:
        os_df['ASSUNTO PADRONIZADO'] = os_df['ASSUNTO PADRONIZADO'].replace(PADRONIZACAO_ASSUNTO)

    # --- Calcular CICLO na aba OS ---
    # CICLO = contagem cumulativa de OS por patrimônio (ordem cronológica)
    os_df = os_df.sort_values(['id_patrimonio', 'data_fechamento_OS'])
    os_df['CICLO'] = os_df.groupby('id_patrimonio').cumcount() + 1

    # --- STATUS_EQUIPAMENTO no RELATORIO (baseado no último CICLO) ---
    if 'PATRIMONIO' in relatorio.columns and 'id_patrimonio' in os_df.columns:
        ultimo_ciclo = os_df.groupby('id_patrimonio')['CICLO'].last().reset_index()
        ultimo_ciclo.columns = ['PATRIMONIO', 'ULTIMO_CICLO']
        relatorio = relatorio.merge(ultimo_ciclo, on='PATRIMONIO', how='left')
        relatorio['STATUS_EQUIPAMENTO'] = relatorio['ULTIMO_CICLO'].apply(
            lambda x: 'REUTILIZADO' if pd.notna(x) and x > 1 else 'NOVO'
        )
    else:
        relatorio['STATUS_EQUIPAMENTO'] = 'NOVO'

    # --- Mapeamento de obsolescência ---
    obs_map = dict(zip(config['MODELO'], config['OBSOLETO?']))

    # Modelos confirmados como NÃO obsoletos (não constam na aba config)
    _modelos_nao_obsoletos = ['ONT ZTE F6600P', 'ONU ZTE F6600P', 'ROTEADOR ZTE H3601 MESH']
    for modelo in _modelos_nao_obsoletos:
        if modelo not in obs_map:
            obs_map[modelo] = 'Não'
            # Adicionar ao DataFrame config para consistência nas validações
            nova_linha = pd.DataFrame({'MODELO': [modelo], 'OBSOLETO?': ['Não']})
            config = pd.concat([config, nova_linha], ignore_index=True)

    # Aplicar em CONTRATOS
    if 'Descrição eqpto' in contratos.columns:
        contratos['OBSOLETO'] = contratos['Descrição eqpto'].map(obs_map).fillna('Não')
        contratos['id_patrimonio_str'] = (
            contratos['id_patrimonio'].astype(str)
            .str.replace(r'\.0$', '', regex=True)
        )

    # Aplicar em BASE_CRUZADA
    if not base_cruzada.empty and 'modelo' in base_cruzada.columns:
        base_cruzada['OBSOLETO'] = base_cruzada['modelo'].map(obs_map).fillna('Não')

    # Mapeamentos DE→PARA canônicos (adicionar ao config se ausentes)
    de_existentes = set(config['DE'].dropna())
    for de_val, para_val in MAPEAMENTO_DE_PARA.items():
        if de_val not in de_existentes:
            nova_linha = pd.DataFrame({'DE': [de_val], 'PARA': [para_val]})
            config = pd.concat([config, nova_linha], ignore_index=True)

    # Mapeamento DE→PARA para padronização de assuntos
    assunto_map = dict(zip(config['DE'].dropna(), config['PARA'].dropna()))

    return {
        'notas': notas,
        'os': os_df,
        'relatorio': relatorio,
        'contratos': contratos,
        'config': config,
        'obsoletos': obsoletos,
        'reaproveitados': reaproveitados,
        'negativado': negativado,
        'base_cruzada': base_cruzada,
        'retirada': retirada,
        'obs_map': obs_map,
        'assunto_map': assunto_map,
        # Metadados de limpeza (para Auditoria)
        '_limpeza': {
            'os_total_bruto': _os_total_bruto,
            'os_sem_patrimonio': _os_sem_patrimonio,
            'os_duplicadas': _os_duplicadas,
        },
    }


def _processar_base_cruzada(raw_df):
    """Processa BASE_CRUZADA que não tem header na planilha."""
    if len(raw_df) < 2:
        return pd.DataFrame()

    df = raw_df.copy()
    df.columns = [
        'id_patrimonio', 'serie', 'modelo', 'cliente',
        'status_contrato', 'status_internet', 'data_mov',
        'tem_os', 'ciclo', 'classificacao'
    ]
    # Remover primeira linha (NaN)
    df = df.iloc[1:].reset_index(drop=True)
    df['data_mov'] = pd.to_datetime(df['data_mov'], errors='coerce')
    df['id_patrimonio_str'] = (
        df['id_patrimonio'].astype(str)
        .str.replace(r'\.0$', '', regex=True)
    )
    return df


def _validar_colunas(df, nome_aba, colunas_obrigatorias):
    """Valida se as colunas obrigatórias existem no DataFrame."""
    faltando = [c for c in colunas_obrigatorias if c not in df.columns]
    if faltando:
        st.error(
            f"Aba '{nome_aba}' com colunas faltando: {', '.join(faltando)}. "
            f"Colunas encontradas: {', '.join(df.columns)}"
        )
        st.stop()


# --- Funções auxiliares para análise por período ---

def get_os_periodo(os_df, data_inicio, data_fim):
    """Filtra OS pelo período de fechamento."""
    return os_df[
        (os_df['data_fechamento_OS'] >= pd.to_datetime(data_inicio)) &
        (os_df['data_fechamento_OS'] <= pd.to_datetime(data_fim))
    ]


def enriquecer_com_relatorio(patrimonios_series, relatorio):
    """Cruza lista de patrimônios com RELATORIO para obter NF, modelo, data NF.

    Args:
        patrimonios_series: Series ou set de id_patrimonio (str)
        relatorio: DataFrame do RELATORIO

    Returns:
        DataFrame com colunas do RELATORIO filtrado pelos patrimônios
    """
    pat_set = set(str(p) for p in patrimonios_series if pd.notna(p))
    rel = relatorio.copy()
    rel['PAT_STR'] = rel['PATRIMONIO'].astype(str).str.replace('.0', '', regex=False)
    return rel[rel['PAT_STR'].isin(pat_set)]


def get_meses_disponiveis(os_df):
    """Retorna lista de meses com dados de OS, ordenados do mais recente ao mais antigo."""
    datas = os_df['data_fechamento_OS'].dropna()
    if datas.empty:
        return []
    meses = sorted(datas.dt.to_period('M').unique())
    return list(reversed(meses))


def periodo_do_mes(mes_period):
    """Retorna (data_inicio, data_fim) de um período mensal."""
    data_inicio = mes_period.start_time
    data_fim = mes_period.end_time
    return data_inicio, data_fim
