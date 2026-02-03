"""
ANÁLISE MENSAL — Detalhamento operacional por mês
Fonte híbrida: OS (dados temporais) + RELATORIO (cadastro NF/modelo)
Comparativo obrigatório: mês selecionado vs mês anterior
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from dados import (
    load_data, fmt, get_os_periodo, enriquecer_com_relatorio,
    get_meses_disponiveis, periodo_do_mes,
)

st.set_page_config(
    page_title="Analise Mensal - Equipamentos",
    page_icon="📅",
    layout="wide",
)


# ============================================================
# CÁLCULOS POR PERÍODO (fonte: OS + RELATORIO para enriquecer)
# ============================================================

@st.cache_data
def calcular_ativacoes(os_df, relatorio, data_inicio, data_fim):
    """Conta instalações no período usando OS, enriquece com RELATORIO.

    Evita o 'efeito foto' do RELATORIO: OS preserva cada evento.
    """
    os_per = get_os_periodo(os_df, data_inicio, data_fim)
    os_inst = os_per[
        os_per['ASSUNTO PADRONIZADO'].str.contains('INSTALAC', case=False, na=False)
    ]

    patrimonios = os_inst['id_patrimonio'].dropna().unique()
    rel_enriq = enriquecer_com_relatorio(patrimonios, relatorio)

    tabela = []
    if len(rel_enriq) > 0:
        rel_total = relatorio.copy()
        for (nf, modelo), grupo in rel_enriq.groupby(['NF', 'DESCRICAO']):
            total_nf = len(rel_total[(rel_total['NF'] == nf) & (rel_total['DESCRICAO'] == modelo)])
            ativados = len(grupo)
            tabela.append({
                'NF': nf,
                'Modelo': modelo,
                'Data NF': grupo['DATA NF'].min(),
                'Comprados (Total NF)': total_nf,
                'Ativados': ativados,
                'Taxa': ativados / total_nf if total_nf > 0 else 0,
            })

    df = pd.DataFrame(tabela)
    if len(df) > 0:
        df = df.sort_values('Ativados', ascending=False)

    # Contagem de ativados COM NF (consistente com Taxa de Ativação)
    total_ativados_nf = int(df['Ativados'].sum()) if len(df) > 0 else 0
    return df, total_ativados_nf


@st.cache_data
def calcular_ativacoes_acumuladas(os_df, relatorio, data_fim):
    """Ativações acumuladas: todas as instalações desde sempre até data_fim."""
    os_acum = os_df[os_df['data_fechamento_OS'] <= pd.to_datetime(data_fim)]
    os_inst = os_acum[
        os_acum['ASSUNTO PADRONIZADO'].str.contains('INSTALAC', case=False, na=False)
    ]
    patrimonios = os_inst['id_patrimonio'].dropna().unique()
    rel_enriq = enriquecer_com_relatorio(patrimonios, relatorio)

    tabela = []
    if len(rel_enriq) > 0:
        rel_total = relatorio.copy()
        for (nf, modelo), grupo in rel_enriq.groupby(['NF', 'DESCRICAO']):
            total_nf = len(rel_total[(rel_total['NF'] == nf) & (rel_total['DESCRICAO'] == modelo)])
            ativados = len(grupo)
            tabela.append({
                'NF': nf,
                'Modelo': modelo,
                'Data NF': grupo['DATA NF'].min(),
                'Comprados (Total NF)': total_nf,
                'Ativados Acumulado': ativados,
                'Taxa Acumulada': ativados / total_nf if total_nf > 0 else 0,
            })

    df = pd.DataFrame(tabela)
    if len(df) > 0:
        df = df.sort_values('Ativados Acumulado', ascending=False)
    return df


@st.cache_data
def calcular_manutencao(os_df, data_inicio, data_fim):
    """Manutenção: Novos vs Reutilizados por tipo.

    Usa CICLO da OS (não do RELATORIO) para evitar viés temporal.
    CICLO=1 na linha da OS = novo, CICLO>1 = reutilizado.
    """
    os_per = get_os_periodo(os_df, data_inicio, data_fim)

    tipos = ['MANUTENCAO', 'MESH', 'UPGRADE']
    resultado = {}
    for tipo in tipos:
        os_tipo = os_per[os_per['ASSUNTO PADRONIZADO'] == tipo]
        total = len(os_tipo)
        novos = len(os_tipo[os_tipo['CICLO'] == 1])
        reutilizados = len(os_tipo[os_tipo['CICLO'] > 1])
        resultado[tipo] = {
            'total': total,
            'novos': novos,
            'reutilizados': reutilizados,
            'pct_novos': novos / total if total > 0 else 0,
            'pct_reutilizados': reutilizados / total if total > 0 else 0,
        }
    return resultado


@st.cache_data
def calcular_parque_rede(contratos, config_obs_map, os_df, data_fim):
    """Equipamentos na rede usando CONTRATOS + config.

    CONTRATOS tem TODOS os equipamentos (com e sem NF).
    """
    # Filtrar contratos ativos
    ativos = contratos[contratos['status_contrato'] == 'Ativo'].copy()
    total_ativos = len(ativos)

    # Obsolescência
    obs_ativos = len(ativos[ativos['OBSOLETO'] == 'Sim'])
    nao_obs_ativos = len(ativos[ativos['OBSOLETO'] == 'Não'])

    # Média de chamados dos clientes com equipamento obsoleto
    # Cruzar patrimônios obsoletos ativos com OS
    pat_obs = set(
        ativos[ativos['OBSOLETO'] == 'Sim']['id_patrimonio_str'].dropna()
    )
    os_obs = os_df[os_df['id_patrimonio'].isin(pat_obs)]
    # Contar OS por cliente (usando id do patrimônio como proxy)
    if len(os_obs) > 0 and 'ID_cliente' in os_obs.columns:
        media_chamados = os_obs.groupby('ID_cliente').size().mean()
    else:
        media_chamados = 0

    # Top modelos obsoletos ativos
    top_obsoletos = pd.DataFrame()
    if 'Descrição eqpto' in ativos.columns:
        obs_df = ativos[ativos['OBSOLETO'] == 'Sim']
        if len(obs_df) > 0:
            top_obsoletos = (
                obs_df.groupby('Descrição eqpto')
                .size()
                .reset_index(name='Qtd Ativa')
                .sort_values('Qtd Ativa', ascending=False)
                .head(10)
            )
            top_obsoletos.columns = ['Modelo', 'Qtd Ativa']

    return {
        'total_ativos': total_ativos,
        'obs_ativos': obs_ativos,
        'nao_obs_ativos': nao_obs_ativos,
        'media_chamados_obs': media_chamados,
        'top_obsoletos': top_obsoletos,
    }


# ============================================================
# FORMATAÇÃO DE TABELAS
# ============================================================

def formatar_tabela_ativacoes(df, col_ativados, col_taxa):
    """Formata tabela de ativações para exibição."""
    if len(df) == 0:
        return df
    exib = df.copy()
    if 'Data NF' in exib.columns:
        exib['Data NF'] = pd.to_datetime(exib['Data NF'], errors='coerce').dt.strftime('%d/%m/%Y')
    exib[col_taxa] = (exib[col_taxa] * 100).round(1).astype(str) + '%'
    return exib


def render_delta(valor_atual, valor_anterior, is_pct=False):
    """Retorna string de delta para st.metric."""
    if valor_anterior is None or valor_anterior == 0:
        return None
    diff = valor_atual - valor_anterior
    if is_pct:
        return f"{diff:+.1f}pp"
    return f"{diff:+,}".replace(",", ".")


# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

def main():
    st.title("📅 Analise Mensal")
    st.markdown("**Detalhamento operacional por mês com comparativo**")
    st.markdown("---")

    with st.spinner('Carregando dados...'):
        data = load_data()

    os_df = data['os']
    relatorio = data['relatorio']
    contratos = data['contratos']

    # ========================================
    # FILTROS
    # ========================================

    st.subheader("Configuracoes")
    col1, col2, col3 = st.columns(3)

    meses = get_meses_disponiveis(os_df)
    if not meses:
        st.warning("Nenhum dado de OS encontrado.")
        return

    with col1:
        mes_labels = [m.strftime('%B/%Y').capitalize() for m in meses]
        mes_idx = st.selectbox("Mes de Analise", range(len(meses)), format_func=lambda i: mes_labels[i])
        mes_selecionado = meses[mes_idx]

    # Mês anterior automático
    mes_anterior = mes_selecionado - 1

    with col2:
        modelos_lista = ['Todos'] + sorted([
            m for m in relatorio['DESCRICAO'].unique() if pd.notna(m)
        ])
        modelo_filtro = st.selectbox("Modelo", modelos_lista)

    with col3:
        nf_lista = ['Todos'] + sorted(
            relatorio['NF'].dropna().astype(str).unique().tolist(), reverse=True
        )
        nf_filtro = st.selectbox("Nota Fiscal", nf_lista)

    st.caption(
        f"Periodo: **{mes_selecionado.strftime('%B/%Y')}** | "
        f"Comparativo: **{mes_anterior.strftime('%B/%Y')}**"
    )

    # Aviso de confiabilidade para períodos anteriores a mai/2024
    _limite_confiavel = pd.Period('2024-05', 'M')
    if mes_selecionado < _limite_confiavel:
        st.warning(
            f"**Dados parciais para {mes_selecionado.strftime('%B/%Y')}.**  \n"
            f"Antes de maio/2024, grande parte das OS nao possuia patrimonio vinculado "
            f"(2020-2022: 100%, 2023: ~69%, jan-abr/2024: ~8%). "
            f"Os KPIs de ativacoes e manutencao estao subcontados neste periodo. "
            f"Consulte a aba Auditoria > Ingestao para detalhes."
        )

    st.markdown("---")

    # Calcular períodos
    ini_atual, fim_atual = periodo_do_mes(mes_selecionado)
    ini_anterior, fim_anterior = periodo_do_mes(mes_anterior)

    # Aplicar filtros de modelo/NF no OS e RELATORIO
    os_filtrado = os_df.copy()
    rel_filtrado = relatorio.copy()

    if modelo_filtro != 'Todos':
        os_filtrado = os_filtrado[os_filtrado['descricao_produto'] == modelo_filtro]
        rel_filtrado = rel_filtrado[rel_filtrado['DESCRICAO'] == modelo_filtro]

    if nf_filtro != 'Todos':
        # Filtrar RELATORIO pela NF, depois filtrar OS pelos patrimônios dessa NF
        rel_filtrado = rel_filtrado[rel_filtrado['NF'].astype(str) == str(nf_filtro)]
        pat_nf = set(rel_filtrado['PATRIMONIO'].astype(str))
        os_filtrado = os_filtrado[os_filtrado['id_patrimonio'].isin(pat_nf)]

    # ========================================
    # SEÇÃO 1: ATIVAÇÕES POR NF
    # ========================================

    st.subheader("1. Ativacoes por Nota Fiscal")

    # Período atual
    df_ativ_atual, total_inst_atual = calcular_ativacoes(
        os_filtrado, rel_filtrado, ini_atual, fim_atual
    )
    # Período anterior
    df_ativ_ant, total_inst_ant = calcular_ativacoes(
        os_filtrado, rel_filtrado, ini_anterior, fim_anterior
    )

    # Calcular acumulado (necessário para KPIs)
    df_acum = calcular_ativacoes_acumuladas(os_filtrado, rel_filtrado, fim_atual)
    df_acum_ant = calcular_ativacoes_acumuladas(os_filtrado, rel_filtrado, fim_anterior)

    total_com_nf = len(rel_filtrado)
    total_acum = int(df_acum['Ativados Acumulado'].sum()) if len(df_acum) > 0 else 0
    total_acum_ant = int(df_acum_ant['Ativados Acumulado'].sum()) if len(df_acum_ant) > 0 else 0

    taxa_ativ_mes = total_inst_atual / total_com_nf * 100 if total_com_nf > 0 else 0
    taxa_ativ_ant = total_inst_ant / total_com_nf * 100 if total_com_nf > 0 else 0
    pendentes = total_com_nf - total_acum
    pendentes_ant = total_com_nf - total_acum_ant
    taxa_acum = total_acum / total_com_nf * 100 if total_com_nf > 0 else 0
    taxa_acum_ant = total_acum_ant / total_com_nf * 100 if total_com_nf > 0 else 0

    # KPIs com delta
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric(
            "Ativacoes no Mes",
            fmt(total_inst_atual),
            delta=render_delta(total_inst_atual, total_inst_ant),
        )
    with c2:
        st.metric(
            "Taxa de Ativacao (Mes)",
            f"{taxa_ativ_mes:.1f}%",
            delta=render_delta(taxa_ativ_mes, taxa_ativ_ant, is_pct=True),
        )
    with c3:
        st.metric(
            "Pendentes de Ativacao",
            fmt(pendentes),
            delta=render_delta(pendentes, pendentes_ant),
            delta_color="inverse",
        )
    with c4:
        st.metric(
            "Taxa Ativacao Acumulada",
            f"{taxa_acum:.1f}%",
            delta=render_delta(taxa_acum, taxa_acum_ant, is_pct=True),
        )

    # Tabela do período
    st.markdown(f"**Ativacoes em {mes_selecionado.strftime('%B/%Y')}**")
    if len(df_ativ_atual) > 0:
        st.dataframe(
            formatar_tabela_ativacoes(df_ativ_atual, 'Ativados', 'Taxa'),
            use_container_width=True, height=280,
        )
    else:
        st.info("Nenhuma instalacao no periodo.")

    # Acumulado (tabela)
    st.markdown(f"**Ativacoes Acumuladas ate {mes_selecionado.strftime('%B/%Y')}**")
    if len(df_acum) > 0:
        st.dataframe(
            formatar_tabela_ativacoes(df_acum, 'Ativados Acumulado', 'Taxa Acumulada'),
            use_container_width=True, height=280,
        )
    else:
        st.info("Nenhuma ativacao acumulada encontrada.")

    st.markdown("---")

    # ========================================
    # SEÇÃO 2: MANUTENÇÃO — NOVOS vs REUTILIZADOS
    # ========================================

    st.subheader("2. Manutencao: Novos vs Reutilizados")
    st.caption("Fonte: OS do periodo. CICLO=1 na OS = Novo | CICLO>1 = Reutilizado")

    manut_atual = calcular_manutencao(os_filtrado, ini_atual, fim_atual)
    manut_ant = calcular_manutencao(os_filtrado, ini_anterior, fim_anterior)

    # Cards por tipo
    col_m1, col_m2, col_m3 = st.columns(3)
    for col, tipo in zip([col_m1, col_m2, col_m3], ['MANUTENCAO', 'MESH', 'UPGRADE']):
        with col:
            d = manut_atual[tipo]
            d_ant = manut_ant[tipo]
            titulo = 'Manutencao' if tipo == 'MANUTENCAO' else tipo.capitalize()
            st.markdown(f"**{titulo}**")
            st.metric(
                "Total OS", fmt(d['total']),
                delta=render_delta(d['total'], d_ant['total']),
            )
            if d['total'] > 0:
                st.metric(
                    "Novos", f"{fmt(d['novos'])} ({d['pct_novos']*100:.1f}%)",
                )
                st.metric(
                    "Reutilizados", f"{fmt(d['reutilizados'])} ({d['pct_reutilizados']*100:.1f}%)",
                )

    # Tabela consolidada
    rows = []
    for tipo in ['MANUTENCAO', 'MESH', 'UPGRADE']:
        d = manut_atual[tipo]
        d_ant = manut_ant[tipo]
        titulo = 'Manutencao' if tipo == 'MANUTENCAO' else tipo.capitalize()
        rows.append({
            'Tipo': titulo,
            'Total Mes': d['total'],
            'Novos Mes': d['novos'],
            '% Novos': f"{d['pct_novos']*100:.1f}%",
            'Reutil. Mes': d['reutilizados'],
            '% Reutil.': f"{d['pct_reutilizados']*100:.1f}%",
            'Total Mes Ant.': d_ant['total'],
            'Delta': d['total'] - d_ant['total'],
        })
    t_at = sum(manut_atual[t]['total'] for t in ['MANUTENCAO', 'MESH', 'UPGRADE'])
    t_n = sum(manut_atual[t]['novos'] for t in ['MANUTENCAO', 'MESH', 'UPGRADE'])
    t_r = sum(manut_atual[t]['reutilizados'] for t in ['MANUTENCAO', 'MESH', 'UPGRADE'])
    t_ant = sum(manut_ant[t]['total'] for t in ['MANUTENCAO', 'MESH', 'UPGRADE'])
    rows.append({
        'Tipo': 'TOTAL',
        'Total Mes': t_at,
        'Novos Mes': t_n,
        '% Novos': f"{t_n/t_at*100:.1f}%" if t_at > 0 else "0.0%",
        'Reutil. Mes': t_r,
        '% Reutil.': f"{t_r/t_at*100:.1f}%" if t_at > 0 else "0.0%",
        'Total Mes Ant.': t_ant,
        'Delta': t_at - t_ant,
    })
    st.dataframe(pd.DataFrame(rows), use_container_width=True)

    st.markdown("---")

    # ========================================
    # SEÇÃO 3: PARQUE NA REDE
    # ========================================

    st.subheader("3. Equipamentos na Rede")
    st.caption("Fonte: CONTRATOS + config (parque total, incluindo equipamentos sem NF)")

    parque = calcular_parque_rede(contratos, data['obs_map'], os_df, fim_atual)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Total Ativos na Rede", fmt(parque['total_ativos']))
    with c2:
        st.metric("Obsoletos Ativos", fmt(parque['obs_ativos']))
    with c3:
        st.metric("Nao Obsoletos Ativos", fmt(parque['nao_obs_ativos']))
    with c4:
        st.metric("Media Chamados (Obsoletos)", f"{parque['media_chamados_obs']:.1f}")

    if len(parque['top_obsoletos']) > 0:
        st.markdown("**Top 10 Modelos Obsoletos Ativos**")
        st.dataframe(parque['top_obsoletos'], use_container_width=True)

    st.markdown("---")

    # ========================================
    # SEÇÃO 4: RETORNO DE EQUIPAMENTOS (FUTURO)
    # ========================================

    st.subheader("4. Retorno de Equipamentos")
    st.info(
        "Secao prevista para dados futuros. "
        "Aguardando preenchimento da aba RETIRADA com dados de OS de retirada."
    )
    st.markdown("""
    **Indicadores previstos:**
    - Total de retornos no mes
    - Quantos eram obsoletos
    - Quantos para descarte tecnico
    - Quantos para garantia
    - Quantos para reutilizacao
    """)

    st.markdown("---")

    # ========================================
    # SEÇÃO 5: CONTROLE DE RETIRADAS (FUTURO)
    # ========================================

    st.subheader("5. Controle de Retiradas")

    # Dados parciais disponíveis: OS de retirada + NEGATIVADO
    os_retirada_atual = get_os_periodo(os_df, ini_atual, fim_atual)
    os_retirada_atual = os_retirada_atual[
        os_retirada_atual['ASSUNTO PADRONIZADO'].str.contains('RETIRADA', case=False, na=False)
    ]
    os_retirada_ant = get_os_periodo(os_df, ini_anterior, fim_anterior)
    os_retirada_ant = os_retirada_ant[
        os_retirada_ant['ASSUNTO PADRONIZADO'].str.contains('RETIRADA', case=False, na=False)
    ]

    negativado = data['negativado']
    total_neg = len(negativado) if not negativado.empty else 0

    c1, c2 = st.columns(2)
    with c1:
        st.metric(
            "OS de Retirada no Mes",
            fmt(len(os_retirada_atual)),
            delta=render_delta(len(os_retirada_atual), len(os_retirada_ant)),
        )
    with c2:
        st.metric("Total Negativados (acumulado)", fmt(total_neg))

    st.info(
        "Detalhamento completo aguardando preenchimento da aba RETIRADA. "
        "Indicadores previstos: nao retornaram (multa), retornaram → obsoletos / reutilizacao / garantia."
    )

    # ========================================
    # RODAPÉ
    # ========================================

    st.markdown("---")
    st.markdown(f"""
    <div style='text-align: center; color: #666; font-size: 0.9em;'>
        Analise: {mes_selecionado.strftime('%B/%Y')} |
        Comparativo: {mes_anterior.strftime('%B/%Y')} |
        Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
