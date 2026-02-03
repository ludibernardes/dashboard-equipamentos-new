"""
AUDITORIA — Validação de dados + Consulta de dados brutos
Seções:
  A. Integridade dos dados
  B. Qualidade da ingestão
  C. Cruzamentos e divergências
  D. Contadores gerais
  E. Dados brutos (busca por patrimônio, NF, cliente)
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from dados import load_data, fmt

st.set_page_config(
    page_title="Auditoria - Equipamentos",
    page_icon="🔍",
    layout="wide",
)


# ============================================================
# CÁLCULOS DE AUDITORIA
# ============================================================

@st.cache_data
def calcular_integridade(os_df, relatorio, contratos, config, limpeza):
    """Verifica integridade dos dados entre as abas."""
    # 1. OS sem patrimônio (valor real, contado ANTES da remoção no load_data)
    os_sem_pat = limpeza.get('os_sem_patrimonio', 0)

    # 2. Patrimônios no RELATORIO sem nenhuma OS
    pat_rel = set(relatorio['PATRIMONIO'].dropna().astype(str))
    pat_os = set(os_df['id_patrimonio'].dropna().astype(str))
    pat_sem_os = pat_rel - pat_os
    pat_os_sem_rel = pat_os - pat_rel

    # 3. Patrimônios no CONTRATOS sem NF (não estão no RELATORIO)
    if 'id_patrimonio_str' in contratos.columns:
        pat_contratos = set(contratos['id_patrimonio_str'].dropna())
        pat_sem_nf = pat_contratos - pat_rel
    else:
        pat_contratos = set()
        pat_sem_nf = set()

    # 4. Modelos sem mapeamento de obsolescência
    modelos_rel = set(relatorio['DESCRICAO'].dropna().unique())
    modelos_config = set(config['MODELO'].dropna().unique())
    modelos_sem_map = modelos_rel - modelos_config

    # 5. Assuntos não mapeados (DE→PARA)
    assuntos_os = set(os_df['Descrição Assunto'].dropna().unique()) if 'Descrição Assunto' in os_df.columns else set()
    assuntos_de = set(config['DE'].dropna().unique())
    assuntos_sem_map = assuntos_os - assuntos_de

    # 6. ASSUNTO PADRONIZADO com valor '****' (não mapeado)
    if 'ASSUNTO PADRONIZADO' in os_df.columns:
        nao_mapeados = os_df[os_df['ASSUNTO PADRONIZADO'] == '****']
        qtd_nao_mapeados = len(nao_mapeados)
    else:
        qtd_nao_mapeados = 0

    return {
        'os_sem_patrimonio': os_sem_pat,
        'pat_relatorio_sem_os': pat_sem_os,
        'qtd_pat_sem_os': len(pat_sem_os),
        'pat_os_sem_relatorio': pat_os_sem_rel,
        'qtd_pat_os_sem_rel': len(pat_os_sem_rel),
        'pat_contratos_sem_nf': pat_sem_nf,
        'qtd_sem_nf': len(pat_sem_nf),
        'modelos_sem_mapeamento': modelos_sem_map,
        'qtd_modelos_sem_map': len(modelos_sem_map),
        'assuntos_sem_mapeamento': assuntos_sem_map,
        'qtd_assuntos_sem_map': len(assuntos_sem_map),
        'qtd_assunto_nao_mapeado': qtd_nao_mapeados,
        'total_pat_relatorio': len(pat_rel),
        'total_pat_os': len(pat_os),
        'total_pat_contratos': len(pat_contratos),
    }


@st.cache_data
def calcular_qualidade_ingestao(os_df, limpeza):
    """Analisa qualidade da ingestão de OS."""
    datas = os_df['data_fechamento_OS'].dropna()

    if datas.empty:
        return {
            'data_min': None, 'data_max': None,
            'meses_cobertos': 0, 'total_os': 0,
            'os_bruto': limpeza.get('os_total_bruto', 0),
            'os_sem_patrimonio': limpeza.get('os_sem_patrimonio', 0),
            'duplicados_removidos': limpeza.get('os_duplicadas', 0),
            'cobertura': pd.DataFrame(),
        }

    # Cobertura mensal
    meses = datas.dt.to_period('M')
    cobertura = meses.value_counts().sort_index().reset_index()
    cobertura.columns = ['Mes', 'Qtd OS']
    cobertura['Mes'] = cobertura['Mes'].astype(str)

    return {
        'data_min': datas.min(),
        'data_max': datas.max(),
        'meses_cobertos': meses.nunique(),
        'total_os': len(os_df),
        'os_bruto': limpeza.get('os_total_bruto', 0),
        'os_sem_patrimonio': limpeza.get('os_sem_patrimonio', 0),
        'duplicados_removidos': limpeza.get('os_duplicadas', 0),
        'cobertura': cobertura,
    }


@st.cache_data
def calcular_cruzamentos(relatorio, contratos, os_df):
    """Cruza RELATORIO vs CONTRATOS para encontrar divergências."""
    resultados = {}

    # RELATORIO vs CONTRATOS — patrimônios ativos sem NF
    if 'id_patrimonio_str' in contratos.columns:
        ativos = contratos[contratos['status_contrato'] == 'Ativo']
        pat_ativos = set(ativos['id_patrimonio_str'].dropna())
        pat_rel = set(relatorio['PATRIMONIO'].dropna().astype(str))

        # Ativos no CONTRATOS sem registro no RELATORIO
        ativos_sem_rel = pat_ativos - pat_rel
        resultados['ativos_sem_relatorio'] = len(ativos_sem_rel)
        resultados['total_ativos'] = len(pat_ativos)
        resultados['cobertura_relatorio'] = len(pat_ativos & pat_rel) / len(pat_ativos) * 100 if len(pat_ativos) > 0 else 0

        # Patrimônios no RELATORIO que não estão ativos no CONTRATOS
        rel_sem_ativo = pat_rel - pat_ativos
        # Podem estar negativados ou cancelados
        resultados['rel_sem_contrato_ativo'] = len(rel_sem_ativo)
    else:
        resultados['ativos_sem_relatorio'] = 0
        resultados['total_ativos'] = 0
        resultados['cobertura_relatorio'] = 0
        resultados['rel_sem_contrato_ativo'] = 0

    # LOCAL_EQUIPAMENTO vs status_contrato
    if 'LOCAL_EQUIPAMENTO' in relatorio.columns:
        dist_local = relatorio['LOCAL_EQUIPAMENTO'].value_counts().reset_index()
        dist_local.columns = ['Local', 'Quantidade']
        resultados['distribuicao_local'] = dist_local
    else:
        resultados['distribuicao_local'] = pd.DataFrame()

    return resultados


@st.cache_data
def calcular_contadores(data):
    """Contadores gerais de todas as abas."""
    contadores = []

    abas = {
        'NOTAS': data['notas'],
        'OS': data['os'],
        'RELATORIO': data['relatorio'],
        'CONTRATOS': data['contratos'],
        'config': data['config'],
        'OBSOLETOS': data['obsoletos'],
        'REAPROVEITADOS': data['reaproveitados'],
        'NEGATIVADO': data['negativado'],
        'BASE_CRUZADA': data['base_cruzada'],
        'RETIRADA': data['retirada'],
    }

    for nome, df in abas.items():
        if df is not None and not df.empty:
            contadores.append({
                'Aba': nome,
                'Linhas': len(df),
                'Colunas': len(df.columns),
            })
        else:
            contadores.append({
                'Aba': nome,
                'Linhas': 0,
                'Colunas': 0,
            })

    return pd.DataFrame(contadores)


# ============================================================
# DADOS BRUTOS — BUSCA
# ============================================================

def buscar_patrimonio(patrimonio, data):
    """Busca informações de um patrimônio em todas as abas."""
    pat_str = str(patrimonio).strip()
    resultados = {}

    # RELATORIO
    rel = data['relatorio']
    rel_match = rel[rel['PATRIMONIO'].astype(str) == pat_str]
    if len(rel_match) > 0:
        resultados['RELATORIO'] = rel_match

    # OS
    os_df = data['os']
    os_match = os_df[os_df['id_patrimonio'].astype(str) == pat_str]
    if len(os_match) > 0:
        resultados['OS'] = os_match.sort_values('data_fechamento_OS', ascending=False)

    # CONTRATOS
    contratos = data['contratos']
    if 'id_patrimonio_str' in contratos.columns:
        c_match = contratos[contratos['id_patrimonio_str'] == pat_str]
        if len(c_match) > 0:
            resultados['CONTRATOS'] = c_match

    # BASE_CRUZADA
    bc = data['base_cruzada']
    if not bc.empty and 'id_patrimonio_str' in bc.columns:
        bc_match = bc[bc['id_patrimonio_str'] == pat_str]
        if len(bc_match) > 0:
            resultados['BASE_CRUZADA'] = bc_match

    return resultados


def buscar_nf(nf, data):
    """Busca informações de uma NF em todas as abas."""
    nf_str = str(nf).strip()
    resultados = {}

    # RELATORIO
    rel = data['relatorio']
    rel_match = rel[rel['NF'].astype(str) == nf_str]
    if len(rel_match) > 0:
        resultados['RELATORIO'] = rel_match

    # NOTAS
    notas = data['notas']
    if 'Número NF' in notas.columns:
        n_match = notas[notas['Número NF'].astype(str) == nf_str]
        if len(n_match) > 0:
            resultados['NOTAS'] = n_match

    return resultados


def buscar_cliente(cliente_id, data):
    """Busca informações de um cliente em todas as abas."""
    cli_str = str(cliente_id).strip()
    resultados = {}

    # OS
    os_df = data['os']
    if 'ID_cliente' in os_df.columns:
        os_match = os_df[os_df['ID_cliente'].astype(str) == cli_str]
        if len(os_match) > 0:
            resultados['OS'] = os_match.sort_values('data_fechamento_OS', ascending=False)

    # CONTRATOS
    contratos = data['contratos']
    for col in ['ID_cliente', 'id_cliente', 'ID Cliente']:
        if col in contratos.columns:
            c_match = contratos[contratos[col].astype(str) == cli_str]
            if len(c_match) > 0:
                resultados['CONTRATOS'] = c_match
            break

    # NEGATIVADO
    neg = data['negativado']
    if not neg.empty:
        for col in neg.columns:
            if 'cliente' in col.lower() or 'id' in col.lower():
                n_match = neg[neg[col].astype(str) == cli_str]
                if len(n_match) > 0:
                    resultados['NEGATIVADO'] = n_match
                    break

    return resultados


# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

def main():
    st.title("🔍 Auditoria de Dados")
    st.markdown("**Validacao, cruzamentos e consulta de dados brutos**")
    st.markdown("---")

    with st.spinner('Carregando dados...'):
        data = load_data()

    os_df = data['os']
    relatorio = data['relatorio']
    contratos = data['contratos']
    config = data['config']
    limpeza = data.get('_limpeza', {})

    # Tabs para organizar as seções
    tab_integ, tab_ingest, tab_cruz, tab_cont, tab_brutos = st.tabs([
        "Integridade", "Ingestao", "Cruzamentos", "Contadores", "Dados Brutos"
    ])

    # ========================================
    # A. INTEGRIDADE DOS DADOS
    # ========================================

    with tab_integ:
        st.subheader("A. Integridade dos Dados")
        st.caption("Verifica consistencia entre as abas da planilha e impacto nas analises")

        integ = calcular_integridade(os_df, relatorio, contratos, config, limpeza)

        # Resumo visual
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Patrimonios RELATORIO", fmt(integ['total_pat_relatorio']))
        with c2:
            st.metric("Patrimonios OS (unicos)", fmt(integ['total_pat_os']))
        with c3:
            st.metric("Patrimonios CONTRATOS", fmt(integ['total_pat_contratos']))

        st.markdown("---")

        # Alertas com notas de impacto
        st.markdown("**Alertas de Integridade**")

        col1, col2 = st.columns(2)
        with col1:
            # OS sem patrimônio
            status_0 = "🟢" if integ['os_sem_patrimonio'] == 0 else "🟡"
            st.markdown(f"{status_0} **OS sem patrimonio:** {fmt(integ['os_sem_patrimonio'])}")
            if integ['os_sem_patrimonio'] > 0:
                st.caption(
                    "Tratamento: removidas antes de qualquer analise. Nao afetam KPIs.\n\n"
                    "**Padrao identificado:** a ausencia de patrimonio esta concentrada nos periodos mais antigos. "
                    "Entre 2020-2022, 100% das OS nao possuiam patrimonio vinculado. "
                    "Em 2023, cerca de 69% estavam sem patrimonio. "
                    "A partir de abril/2024, o indice caiu para ~8%, e de maio/2024 em diante e proximo de 0%. "
                    "74% dessas OS sao do tipo INSTALACAO e 13% MANUTENCAO. "
                    "Consequencia: KPIs de ativacoes e manutencao ficam subcontados para periodos anteriores a maio/2024."
                )

            # Patrimônios RELATORIO sem OS
            status_1 = "🟢" if integ['qtd_pat_sem_os'] == 0 else "🟡"
            st.markdown(f"{status_1} **Patrimonios no RELATORIO sem OS:** {fmt(integ['qtd_pat_sem_os'])}")
            if integ['qtd_pat_sem_os'] > 0:
                st.caption(
                    "Impacto: equipamentos comprados (com NF) mas sem nenhuma OS. "
                    "Contam como 'Pendentes de Ativacao' na Analise Mensal."
                )
                with st.expander(f"Ver {min(integ['qtd_pat_sem_os'], 50)} primeiros"):
                    st.write(sorted(list(integ['pat_relatorio_sem_os']))[:50])

            # Patrimônios no OS sem RELATORIO
            status_2 = "🟢" if integ['qtd_pat_os_sem_rel'] == 0 else "🔴"
            st.markdown(f"{status_2} **Patrimonios com OS sem RELATORIO:** {fmt(integ['qtd_pat_os_sem_rel'])}")
            if integ['qtd_pat_os_sem_rel'] > 0:
                st.caption(
                    "Impacto: equipamentos sem NF que tiveram OS. "
                    "NAO aparecem nas ativacoes/manutencoes por NF. "
                    "Visiveis apenas na secao 'Equipamentos na Rede' (via CONTRATOS)."
                )
                with st.expander(f"Ver {min(integ['qtd_pat_os_sem_rel'], 50)} primeiros"):
                    st.write(sorted(list(integ['pat_os_sem_relatorio']))[:50])

            # Contratos sem NF
            status_3 = "🟢" if integ['qtd_sem_nf'] == 0 else "🟡"
            st.markdown(f"{status_3} **Patrimonios em CONTRATOS sem NF:** {fmt(integ['qtd_sem_nf'])}")
            if integ['qtd_sem_nf'] > 0:
                st.caption(
                    "Impacto: equipamentos na rede sem nota fiscal. "
                    "Aparecem apenas nos totais de CONTRATOS (secao Equipamentos na Rede)."
                )

        with col2:
            # Modelos sem mapeamento
            status_4 = "🟢" if integ['qtd_modelos_sem_map'] == 0 else "🟡"
            st.markdown(f"{status_4} **Modelos sem mapeamento de obsolescencia:** {integ['qtd_modelos_sem_map']}")
            if integ['qtd_modelos_sem_map'] > 0:
                st.caption(
                    "Impacto: tratados como NAO OBSOLETOS por padrao (.fillna). "
                    "Considere adicionar ao mapeamento config para precisao."
                )
                with st.expander("Ver modelos"):
                    st.write(sorted(list(integ['modelos_sem_mapeamento'])))

            # Assuntos não mapeados
            status_5 = "🟢" if integ['qtd_assuntos_sem_map'] == 0 else "🟡"
            st.markdown(f"{status_5} **Assuntos de OS sem mapeamento (DE->PARA):** {integ['qtd_assuntos_sem_map']}")
            if integ['qtd_assuntos_sem_map'] > 0:
                st.caption(
                    "Impacto: descricoes de assunto que nao tem correspondencia no config. "
                    "Se o ASSUNTO PADRONIZADO ja existir na OS, nao ha problema."
                )
                with st.expander("Ver assuntos"):
                    st.write(sorted(list(integ['assuntos_sem_mapeamento'])))

            # OS com ASSUNTO '****'
            status_6 = "🟢" if integ['qtd_assunto_nao_mapeado'] == 0 else "🔴"
            st.markdown(f"{status_6} **OS com ASSUNTO PADRONIZADO = '****':** {fmt(integ['qtd_assunto_nao_mapeado'])}")
            if integ['qtd_assunto_nao_mapeado'] > 0:
                st.caption(
                    "Impacto: EXCLUIDAS de toda analise (nao casam com INSTALAC, "
                    "MANUTENCAO, MESH, UPGRADE). Dados perdidos silenciosamente."
                )

    # ========================================
    # B. QUALIDADE DA INGESTÃO
    # ========================================

    with tab_ingest:
        st.subheader("B. Qualidade da Ingestao")
        st.caption("Analisa cobertura temporal, limpeza aplicada e duplicados na aba OS")

        qual = calcular_qualidade_ingestao(os_df, limpeza)

        # Pipeline de limpeza
        st.markdown("**Pipeline de Limpeza Aplicado**")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("OS Bruto (arquivo)", fmt(qual['os_bruto']))
        with c2:
            st.metric("Sem Patrimonio (removidas)", fmt(qual['os_sem_patrimonio']))
        with c3:
            st.metric("Duplicadas (removidas)", fmt(qual['duplicados_removidos']))
        with c4:
            st.metric("OS Final (usadas)", fmt(qual['total_os']))

        if qual['duplicados_removidos'] > 0:
            st.info(
                f"{fmt(qual['duplicados_removidos'])} OS duplicadas foram removidas automaticamente. "
                f"Sem duplicatas, os KPIs de ativacao e manutencao refletem contagens corretas."
            )
        elif qual['os_bruto'] > 0:
            st.success("Nenhuma duplicata encontrada na base de OS.")

        st.markdown("---")

        # Alerta de confiabilidade por período
        if qual['os_sem_patrimonio'] > 0:
            st.markdown("**Confiabilidade por Periodo**")
            st.error(
                f"**{fmt(qual['os_sem_patrimonio'])} OS sem patrimonio** foram removidas antes das analises. "
                f"Essas OS existem na planilha mas nao podem ser vinculadas a equipamentos especificos."
            )
            st.markdown(
                "| Periodo | % OS sem patrimonio | Impacto nas analises |\n"
                "|---|---|---|\n"
                "| **2020 a 2022** | **100%** | Nenhum dado disponivel para ativacoes/manutencao |\n"
                "| **2023** | **~69%** | Dados severamente subcontados |\n"
                "| **Jan-Abr 2024** | **~8%** | Dados parcialmente subcontados |\n"
                "| **Mai/2024 em diante** | **~0%** | Dados confiaveis e completos |"
            )
            st.caption(
                "Causa: antes de mai/2024, o sistema nao vinculava patrimonio as OS. "
                "98% das OS sem patrimonio tem status_comodato='Devolvido' "
                "(vinculo desfeito na devolucao do equipamento). "
                "Na Analise Mensal, meses anteriores a mai/2024 exibem um aviso automatico."
            )
            st.markdown("---")

        # Cobertura temporal
        st.markdown("**Cobertura Temporal**")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Meses Cobertos", fmt(qual['meses_cobertos']))
        with c2:
            if qual['data_min'] is not None:
                st.metric("Primeira OS", qual['data_min'].strftime('%d/%m/%Y'))
            else:
                st.metric("Primeira OS", "N/A")
        with c3:
            if qual['data_max'] is not None:
                st.metric("Ultima OS", qual['data_max'].strftime('%d/%m/%Y'))
            else:
                st.metric("Ultima OS", "N/A")

        # Cobertura mensal
        st.markdown("**Cobertura Mensal (OS por mes)**")
        if len(qual['cobertura']) > 0:
            st.dataframe(qual['cobertura'], use_container_width=True, height=400)
        else:
            st.info("Sem dados de cobertura.")

    # ========================================
    # C. CRUZAMENTOS E DIVERGÊNCIAS
    # ========================================

    with tab_cruz:
        st.subheader("C. Cruzamentos e Divergencias")
        st.caption("Compara RELATORIO vs CONTRATOS e identifica inconsistencias")

        cruz = calcular_cruzamentos(relatorio, contratos, os_df)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Ativos no CONTRATOS", fmt(cruz['total_ativos']))
        with c2:
            st.metric("Ativos sem RELATORIO", fmt(cruz['ativos_sem_relatorio']))
        with c3:
            st.metric("Cobertura do RELATORIO", f"{cruz['cobertura_relatorio']:.1f}%")

        if cruz['rel_sem_contrato_ativo'] > 0:
            st.info(
                f"{fmt(cruz['rel_sem_contrato_ativo'])} patrimonios no RELATORIO "
                f"sem contrato ativo (podem ser negativados ou cancelados)."
            )

        # Distribuição LOCAL_EQUIPAMENTO
        st.markdown("**Distribuicao por LOCAL_EQUIPAMENTO (RELATORIO)**")
        if len(cruz.get('distribuicao_local', pd.DataFrame())) > 0:
            st.dataframe(cruz['distribuicao_local'], use_container_width=True)

    # ========================================
    # D. CONTADORES GERAIS
    # ========================================

    with tab_cont:
        st.subheader("D. Contadores Gerais")
        st.caption("Totais por aba da planilha")

        contadores = calcular_contadores(data)
        st.dataframe(contadores, use_container_width=True)

        st.markdown("---")

        # Detalhes extras
        st.markdown("**Detalhes Adicionais**")
        c1, c2 = st.columns(2)

        with c1:
            # NFs únicas
            nfs_unicas = relatorio['NF'].dropna().nunique()
            st.metric("NFs Unicas (RELATORIO)", fmt(nfs_unicas))

            # Modelos únicos
            modelos_unicos = relatorio['DESCRICAO'].dropna().nunique()
            st.metric("Modelos Unicos (RELATORIO)", fmt(modelos_unicos))

        with c2:
            # Range de datas NF
            datas_nf = relatorio['DATA NF'].dropna()
            if len(datas_nf) > 0:
                st.metric("Primeira NF", datas_nf.min().strftime('%d/%m/%Y'))
                st.metric("Ultima NF", datas_nf.max().strftime('%d/%m/%Y'))

    # ========================================
    # E. DADOS BRUTOS — BUSCA
    # ========================================

    with tab_brutos:
        st.subheader("E. Dados Brutos")
        st.caption("Consulta dados em todas as abas por patrimonio, NF ou cliente")

        tipo_busca = st.radio(
            "Buscar por:",
            ["Patrimonio", "Nota Fiscal", "Cliente"],
            horizontal=True,
        )

        if tipo_busca == "Patrimonio":
            valor = st.text_input("ID do Patrimonio", placeholder="Ex: 123456")
            if valor:
                resultados = buscar_patrimonio(valor, data)
                if resultados:
                    for aba, df in resultados.items():
                        st.markdown(f"**{aba}** ({len(df)} registros)")
                        st.dataframe(df, use_container_width=True)
                        st.markdown("---")
                else:
                    st.warning(f"Patrimonio '{valor}' nao encontrado em nenhuma aba.")

        elif tipo_busca == "Nota Fiscal":
            nf_lista = sorted(relatorio['NF'].dropna().astype(str).unique().tolist(), reverse=True)
            valor = st.selectbox("Selecione a NF", [''] + nf_lista)
            if valor:
                resultados = buscar_nf(valor, data)
                if resultados:
                    for aba, df in resultados.items():
                        st.markdown(f"**{aba}** ({len(df)} registros)")
                        st.dataframe(df, use_container_width=True)
                        st.markdown("---")

                    # Resumo da NF
                    if 'RELATORIO' in resultados:
                        rel_nf = resultados['RELATORIO']
                        st.markdown("**Resumo da NF**")
                        c1, c2, c3, c4 = st.columns(4)
                        with c1:
                            st.metric("Total Equipamentos", fmt(len(rel_nf)))
                        with c2:
                            inst = len(rel_nf[rel_nf['LOCAL_EQUIPAMENTO'] == 'INSTALADO'])
                            st.metric("Instalados", fmt(inst))
                        with c3:
                            est = len(rel_nf[rel_nf['LOCAL_EQUIPAMENTO'] == 'EM ESTOQUE'])
                            st.metric("Em Estoque", fmt(est))
                        with c4:
                            rma = len(rel_nf[rel_nf['LOCAL_EQUIPAMENTO'] == 'RMA'])
                            st.metric("Em RMA", fmt(rma))
                else:
                    st.warning(f"NF '{valor}' nao encontrada.")

        elif tipo_busca == "Cliente":
            valor = st.text_input("ID do Cliente", placeholder="Ex: 12345")
            if valor:
                resultados = buscar_cliente(valor, data)
                if resultados:
                    for aba, df in resultados.items():
                        st.markdown(f"**{aba}** ({len(df)} registros)")
                        st.dataframe(df, use_container_width=True)
                        st.markdown("---")
                else:
                    st.warning(f"Cliente '{valor}' nao encontrado em nenhuma aba.")

        st.markdown("---")

        # Opção de exportar dados brutos por aba
        st.subheader("Exportar Dados Brutos")
        aba_export = st.selectbox("Selecione a aba para exportar", [
            'RELATORIO', 'OS', 'CONTRATOS', 'NOTAS', 'config',
            'OBSOLETOS', 'REAPROVEITADOS', 'NEGATIVADO', 'BASE_CRUZADA', 'RETIRADA',
        ])

        aba_map = {
            'RELATORIO': data['relatorio'],
            'OS': data['os'],
            'CONTRATOS': data['contratos'],
            'NOTAS': data['notas'],
            'config': data['config'],
            'OBSOLETOS': data['obsoletos'],
            'REAPROVEITADOS': data['reaproveitados'],
            'NEGATIVADO': data['negativado'],
            'BASE_CRUZADA': data['base_cruzada'],
            'RETIRADA': data['retirada'],
        }

        df_export = aba_map.get(aba_export, pd.DataFrame())
        if df_export is not None and not df_export.empty:
            st.markdown(f"**{aba_export}**: {fmt(len(df_export))} linhas, {len(df_export.columns)} colunas")
            st.dataframe(df_export, use_container_width=True, height=400)

            csv = df_export.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                f"Baixar {aba_export} (CSV)",
                data=csv,
                file_name=f"{aba_export}_{datetime.now():%Y%m%d}.csv",
                mime="text/csv",
            )
        else:
            st.info(f"Aba '{aba_export}' esta vazia.")

    # ========================================
    # RODAPÉ
    # ========================================

    st.markdown("---")
    st.markdown(f"""
    <div style='text-align: center; color: #666; font-size: 0.9em;'>
        Auditoria gerada em: {datetime.now():%d/%m/%Y %H:%M:%S}
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
