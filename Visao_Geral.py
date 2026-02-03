"""
VISÃO GERAL — Painel Executivo
Foto atual do parque + delta vs mês anterior
Fontes: CONTRATOS+config (parque total) | RELATORIO (equipamentos com NF)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np

from dados import (
    load_data, fmt, get_os_periodo, enriquecer_com_relatorio,
    get_meses_disponiveis, periodo_do_mes,
)


# ============================================
# PROCESSAMENTO
# ============================================

@st.cache_data
def calcular_kpis_parque(contratos, relatorio, negativado):
    """KPIs do parque total usando CONTRATOS + RELATORIO."""
    # CONTRATOS = parque total (com e sem NF)
    ativos = contratos[contratos['status_contrato'] == 'Ativo']
    total_rede = len(ativos)
    obs_ativos = len(ativos[ativos['OBSOLETO'] == 'Sim'])
    nao_obs_ativos = len(ativos[ativos['OBSOLETO'] == 'Não'])

    # RELATORIO = equipamentos com NF (foto atual)
    instalados = len(relatorio[relatorio['LOCAL_EQUIPAMENTO'] == 'INSTALADO'])
    em_estoque = len(relatorio[relatorio['LOCAL_EQUIPAMENTO'] == 'EM ESTOQUE'])
    em_rma = len(relatorio[relatorio['LOCAL_EQUIPAMENTO'] == 'RMA'])
    com_tecnico = len(relatorio[relatorio['LOCAL_EQUIPAMENTO'] == 'COM TÉCNICO'])

    # Taxa de utilização (instalados / total com NF)
    total_com_nf = len(relatorio)
    taxa_utilizacao = instalados / total_com_nf * 100 if total_com_nf > 0 else 0

    # Negativados
    total_neg = len(negativado) if not negativado.empty else 0
    neg_contratos = len(contratos[contratos['status_contrato'] == 'Negativado'])

    return {
        'total_rede': total_rede,
        'obs_ativos': obs_ativos,
        'nao_obs_ativos': nao_obs_ativos,
        'total_com_nf': total_com_nf,
        'instalados': instalados,
        'em_estoque': em_estoque,
        'em_rma': em_rma,
        'com_tecnico': com_tecnico,
        'taxa_utilizacao': taxa_utilizacao,
        'negativados': neg_contratos,
    }


@st.cache_data
def gerar_resumo_nf(relatorio, config):
    """Tabela resumo por NF+Modelo (baseada no RELATORIO)."""
    obs_map = dict(zip(config['MODELO'], config['OBSOLETO?']))
    combinacoes = relatorio[['NF', 'DESCRICAO']].drop_duplicates()

    rows = []
    for _, row in combinacoes.iterrows():
        nf, modelo = row['NF'], row['DESCRICAO']
        df_nf = relatorio[(relatorio['NF'] == nf) & (relatorio['DESCRICAO'] == modelo)]

        comprados = len(df_nf)
        ativados_novos = len(df_nf[
            (df_nf['LOCAL_EQUIPAMENTO'] == 'INSTALADO') &
            (df_nf['STATUS_EQUIPAMENTO'] == 'NOVO')
        ])
        ativados_reutil = len(df_nf[
            (df_nf['LOCAL_EQUIPAMENTO'] == 'INSTALADO') &
            (df_nf['STATUS_EQUIPAMENTO'] == 'REUTILIZADO')
        ])
        em_estoque = len(df_nf[df_nf['LOCAL_EQUIPAMENTO'] == 'EM ESTOQUE'])
        em_rma = len(df_nf[df_nf['LOCAL_EQUIPAMENTO'] == 'RMA'])
        com_tecnico = len(df_nf[df_nf['LOCAL_EQUIPAMENTO'] == 'COM TÉCNICO'])

        rows.append({
            'NF': nf,
            'MODELO': modelo,
            'DATA': df_nf['DATA NF'].min(),
            'COMPRADOS': comprados,
            'ATIVADOS_NOVOS': ativados_novos,
            'TAXA_ATIVACAO': ativados_novos / comprados if comprados > 0 else 0,
            'ATIVADOS_REUTIL': ativados_reutil,
            'EM_ESTOQUE': em_estoque,
            'PERC_ESTOQUE': em_estoque / comprados if comprados > 0 else 0,
            'EM_RMA': em_rma,
            'PERC_RMA': em_rma / comprados if comprados > 0 else 0,
            'COM_TECNICO': com_tecnico,
            'OBSOLETO': obs_map.get(modelo, 'Nao'),
        })

    df = pd.DataFrame(rows)
    if len(df) > 0:
        df = df.sort_values('DATA', ascending=False)
    return df


@st.cache_data
def gerar_evolucao_mensal(os_df, relatorio):
    """Dados de evolução mensal (instalações por mês via OS)."""
    os_inst = os_df[
        os_df['ASSUNTO PADRONIZADO'].str.contains('INSTALAC', case=False, na=False)
    ].copy()
    os_inst['MES'] = os_inst['data_fechamento_OS'].dt.to_period('M')

    evolucao = os_inst.groupby('MES').size().reset_index(name='Instalacoes')
    evolucao['MES_STR'] = evolucao['MES'].astype(str)
    return evolucao


# ============================================
# GRÁFICOS
# ============================================

def chart_distribuicao(relatorio):
    """Pizza de distribuição por LOCAL_EQUIPAMENTO."""
    dist = relatorio['LOCAL_EQUIPAMENTO'].value_counts().reset_index()
    dist.columns = ['Status', 'Quantidade']

    cores = {
        'INSTALADO': '#28a745', 'EM ESTOQUE': '#007bff',
        'RMA': '#ffc107', 'COM TÉCNICO': '#dc3545',
        'DESCONTINUADO': '#6c757d',
    }
    fig = px.pie(
        dist, values='Quantidade', names='Status',
        title='Distribuicao por Status (Equipamentos com NF)',
        color='Status', color_discrete_map=cores, hole=0.4,
    )
    fig.update_traces(textposition='inside', textinfo='percent+label')
    fig.update_layout(height=380, margin=dict(t=50, b=20, l=20, r=20))
    return fig


def chart_evolucao(evolucao):
    """Linha de evolução de instalações por mês."""
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=evolucao['MES_STR'], y=evolucao['Instalacoes'],
        mode='lines+markers', name='Instalacoes',
        line=dict(color='#28a745', width=2), marker=dict(size=8),
    ))
    fig.update_layout(
        title='Evolucao Mensal de Instalacoes (via OS)',
        xaxis_title='Mes', yaxis_title='Instalacoes',
        height=380, margin=dict(t=50, b=50, l=50, r=20),
    )
    return fig


def chart_sankey(df_resumo, nf):
    """Sankey de fluxo para uma NF específica."""
    df_nf = df_resumo[df_resumo['NF'] == nf]
    if df_nf.empty:
        return None
    row = df_nf.iloc[0]

    nodes, flows = ['Comprados'], []

    def add(label, value):
        if value > 0:
            if label not in nodes:
                nodes.append(label)
            flows.append(('Comprados', label, int(value)))

    add('Ativados (Novos)', row.get('ATIVADOS_NOVOS', 0))
    add('Ativados (Reutilizados)', row.get('ATIVADOS_REUTIL', 0))
    add('Em Estoque', row.get('EM_ESTOQUE', 0))
    add('Em RMA', row.get('EM_RMA', 0))
    add('Com Tecnico', row.get('COM_TECNICO', 0))

    if not flows:
        return None

    node_colors = {
        'Comprados': '#0056b3', 'Ativados (Novos)': '#1e7e34',
        'Ativados (Reutilizados)': '#d39e00', 'Em Estoque': '#117a8b',
        'Em RMA': '#e66100', 'Com Tecnico': '#bd2130',
    }
    colors = [node_colors.get(n, '#545b62') for n in nodes]

    source = [nodes.index(f[0]) for f in flows]
    target = [nodes.index(f[1]) for f in flows]
    value = [f[2] for f in flows]

    link_colors = []
    for _, t, _ in flows:
        c = node_colors.get(t, '#545b62')
        r, g, b = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
        link_colors.append(f'rgba({r},{g},{b},0.5)')

    fig = go.Figure(go.Sankey(
        node=dict(label=nodes, pad=25, thickness=30, color=colors),
        link=dict(source=source, target=target, value=value, color=link_colors),
    ))
    fig.update_layout(
        title=f'Fluxo NF: {nf}', height=500,
        margin=dict(t=80, b=40, l=40, r=40),
    )
    return fig


# ============================================
# INTERFACE
# ============================================

def main():
    st.set_page_config(
        page_title="Visao Geral", page_icon="📊",
        layout="wide", initial_sidebar_state="expanded",
    )

    st.title("📊 Dashboard - Ciclo de Vida de Equipamentos")
    st.markdown("**Painel Executivo**")
    st.markdown("---")

    with st.spinner('Carregando dados...'):
        data = load_data()

    relatorio = data['relatorio']
    contratos = data['contratos']
    os_df = data['os']
    config = data['config']
    negativado = data['negativado']

    df_resumo = gerar_resumo_nf(relatorio, config)
    kpis = calcular_kpis_parque(contratos, relatorio, negativado)

    # ========================================
    # SIDEBAR — FILTROS
    # ========================================

    with st.sidebar:
        st.header("Filtros")

        st.subheader("Modelo")
        modelos = ['Todos'] + sorted([m for m in df_resumo['MODELO'].unique() if pd.notna(m)])
        modelo_sel = st.selectbox("Modelo", modelos)

        st.subheader("Nota Fiscal")
        nfs = ['Todas'] + sorted([n for n in df_resumo['NF'].unique() if pd.notna(n)], reverse=True)
        nf_sel = st.selectbox("NF", nfs)

    # ========================================
    # APLICAR FILTROS na tabela resumo
    # ========================================

    df_filt = df_resumo.copy()
    if modelo_sel != 'Todos':
        df_filt = df_filt[df_filt['MODELO'] == modelo_sel]
    if nf_sel != 'Todas':
        df_filt = df_filt[df_filt['NF'] == nf_sel]

    # ========================================
    # KPIs PARQUE TOTAL
    # ========================================

    st.subheader("Parque de Equipamentos")
    st.caption("Fonte: CONTRATOS (parque total) + RELATORIO (equipamentos com NF)")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("Total na Rede", fmt(kpis['total_rede']))
    with c2:
        st.metric("Instalados (c/ NF)", fmt(kpis['instalados']))
    with c3:
        st.metric("Em Estoque", fmt(kpis['em_estoque']))
    with c4:
        st.metric("Em RMA", fmt(kpis['em_rma']))
    with c5:
        st.metric("Com Tecnico", fmt(kpis['com_tecnico']))

    c6, c7, c8 = st.columns(3)
    with c6:
        st.metric("Obsoletos Ativos", fmt(kpis['obs_ativos']))
    with c7:
        st.metric("Taxa Utilizacao (c/ NF)", f"{kpis['taxa_utilizacao']:.1f}%")
    with c8:
        st.metric("Negativados", fmt(kpis['negativados']))

    st.markdown("---")

    # ========================================
    # TABELA RESUMO POR NF (filtrada)
    # ========================================

    st.subheader("Resumo por Nota Fiscal")

    if len(df_filt) > 0:
        df_exib = df_filt.copy()
        df_exib['DATA'] = df_exib['DATA'].dt.strftime('%d/%m/%Y')
        df_exib['TAXA_ATIVACAO'] = (df_exib['TAXA_ATIVACAO'] * 100).round(1).astype(str) + '%'
        df_exib['PERC_ESTOQUE'] = (df_exib['PERC_ESTOQUE'] * 100).round(1).astype(str) + '%'
        df_exib['PERC_RMA'] = (df_exib['PERC_RMA'] * 100).round(1).astype(str) + '%'

        colunas = [
            'NF', 'MODELO', 'DATA', 'COMPRADOS', 'ATIVADOS_NOVOS', 'TAXA_ATIVACAO',
            'ATIVADOS_REUTIL', 'EM_ESTOQUE', 'PERC_ESTOQUE', 'EM_RMA', 'PERC_RMA',
            'COM_TECNICO', 'OBSOLETO',
        ]
        st.dataframe(df_exib[colunas], use_container_width=True, height=400)

        csv = df_exib[colunas].to_csv(index=False).encode('utf-8-sig')
        st.download_button("Baixar CSV", data=csv, file_name=f"resumo_{datetime.now():%Y%m%d}.csv", mime="text/csv")
    else:
        st.info("Nenhum dado para os filtros selecionados.")

    st.markdown("---")

    # ========================================
    # SANKEY
    # ========================================

    st.subheader("Fluxo por Nota Fiscal")
    if nf_sel != 'Todas':
        fig_sankey = chart_sankey(df_resumo, nf_sel)
        if fig_sankey:
            st.plotly_chart(fig_sankey, use_container_width=True)
        else:
            st.info("Dados insuficientes para o Sankey desta NF.")
    else:
        st.info("Selecione uma NF na barra lateral para visualizar o fluxo.")

    st.markdown("---")

    # ========================================
    # GRÁFICOS
    # ========================================

    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(chart_distribuicao(relatorio), use_container_width=True)
    with col2:
        evolucao = gerar_evolucao_mensal(os_df, relatorio)
        st.plotly_chart(chart_evolucao(evolucao), use_container_width=True)

    # ========================================
    # RODAPÉ
    # ========================================

    st.markdown("---")
    st.markdown(f"""
    <div style='text-align: center; color: #666; font-size: 0.9em;'>
        Atualizado em: {datetime.now():%d/%m/%Y %H:%M:%S} |
        Parque total: {fmt(kpis['total_rede'])} equipamentos |
        Com NF: {fmt(kpis['total_com_nf'])}
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
