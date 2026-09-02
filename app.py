import io

import pandas as pd

import streamlit as st

from datetime import datetime


from database import (
    importar_pessoas,
    obter_estatisticas,
    obter_historico,
    registrar_voto,
    obter_pessoas,
    obter_configuracao,
    finalizar_votacao,
    reabrir_votacao,
    votacao_ativa,
)


from utils.cracha import leitor_cracha


# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(

    page_title="Reclame Aqui - Votação",

    page_icon="🗳️",

    layout="wide"

)


# ============================================================
# ESTATÍSTICAS
# ============================================================

def mostrar_estatisticas():

    estatisticas = (
        obter_estatisticas()
    )

    col1, col2, col3, col4 = (
        st.columns(4)
    )

    with col1:

        st.metric(
            "👥 CADASTRADOS",
            estatisticas["total"]
        )

    with col2:

        st.metric(
            "🗳️ VOTOS",
            estatisticas["votaram"]
        )

    with col3:

        st.metric(
            "⏳ NÃO VOTARAM",
            estatisticas["nao_votaram"]
        )

    with col4:

        st.metric(
            "📊 PARTICIPAÇÃO",
            f"{estatisticas['percentual']:.1f}%"
        )


# ============================================================
# PROCESSAR VOTO
# ============================================================

def processar_voto(
    identificador
):

    identificador = str(
        identificador
    ).strip()

    if not identificador:

        st.warning(
            "⚠️ Nenhum crachá foi informado."
        )

        return

    resultado = registrar_voto(
        identificador
    )


    # ========================================================
    # ENCERRADA
    # ========================================================

    if resultado["resultado"] == "encerrada":

        st.error(
            "🔴 VOTAÇÃO ENCERRADA"
        )

        return


    # ========================================================
    # CONTABILIZADO
    # ========================================================

    if resultado["resultado"] == "contabilizado":

        st.success(
            "✅ VOTO CONTABILIZADO"
        )

        st.subheader(
            f"👤 {resultado['nome']}"
        )

        col1, col2 = st.columns(2)

        with col1:

            st.write(
                f"**ID Mat:** "
                f"{resultado['id_mat']}"
            )

        with col2:

            st.write(
                f"**Mat:** "
                f"{resultado['mat']}"
            )

        st.write(
            "🕐 **Voto registrado com sucesso.**"
        )


    # ========================================================
    # DUPLICADO
    # ========================================================

    elif resultado["resultado"] == "duplicado":

        st.error(
            "🚫 VOTO DUPLICADO"
        )

        st.subheader(
            f"👤 {resultado['nome']}"
        )

        st.warning(
            "Este colaborador já possui voto registrado."
        )


    # ========================================================
    # NÃO ENCONTRADO
    # ========================================================

    elif resultado["resultado"] == "nao_encontrado":

        st.warning(
            "⚠️ COLABORADOR NÃO ENCONTRADO"
        )

        st.write(
            f"Identificador lido: "
            f"**{identificador}**"
        )


# ============================================================
# EXPORTAR EXCEL FINAL
# ============================================================

def gerar_excel_final():

    pessoas = obter_pessoas()

    historico, colunas_historico = (
        obter_historico()
    )

    estatisticas = (
        obter_estatisticas()
    )

    configuracao = (
        obter_configuracao()
    )

    df_resultado = pd.DataFrame(

        [

            {

                "ID Mat":
                    pessoa.get("id_mat", ""),

                "Mat":
                    pessoa.get("mat", ""),

                "Nome":
                    pessoa.get("nome", ""),

                "Voto":
                    pessoa.get("voto", ""),

                "Data/Hora Voto":
                    pessoa.get("data_voto", ""),

            }

            for pessoa in pessoas

        ]

    )


    df_historico = pd.DataFrame(

        historico,

        columns=colunas_historico

    )


    data_finalizacao = (
        configuracao.get(
            "data_finalizacao"
        )
    )


    df_resumo = pd.DataFrame({

        "Indicador": [

            "Total de colaboradores",

            "Votos registrados",

            "Não votaram",

            "Participação",

            "Total de tentativas",

            "Votos duplicados",

            "Não encontrados",

            "Data de finalização",

        ],

        "Resultado": [

            estatisticas["total"],

            estatisticas["votaram"],

            estatisticas["nao_votaram"],

            f"{estatisticas['percentual']:.2f}%",

            estatisticas["total_tentativas"],

            estatisticas["duplicados"],

            estatisticas["nao_encontrados"],

            data_finalizacao or "",

        ]

    })


    buffer = io.BytesIO()


    with pd.ExcelWriter(

        buffer,

        engine="openpyxl"

    ) as writer:

        df_resumo.to_excel(

            writer,

            index=False,

            sheet_name="RESUMO"

        )

        df_resultado.to_excel(

            writer,

            index=False,

            sheet_name="RESULTADO"

        )

        df_historico.to_excel(

            writer,

            index=False,

            sheet_name="HISTÓRICO"

        )


    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# CABEÇALHO
# ============================================================

st.title(
    "🗳️ Reclame Aqui - Contabilização de Votos"
)

st.caption(
    "Sistema de votação online"
)


# ============================================================
# STATUS
# ============================================================

if votacao_ativa():

    st.success(
        "🟢 VOTAÇÃO ATIVA"
    )

else:

    st.error(
        "🔴 VOTAÇÃO ENCERRADA"
    )


# ============================================================
# ABAS
# ============================================================

aba_votacao, aba_dashboard, aba_historico, aba_importacao = st.tabs(

    [

        "🗳️ VOTAÇÃO",

        "📊 DASHBOARD",

        "📋 HISTÓRICO",

        "📥 IMPORTAR CADASTRO"

    ]

)


# ============================================================
# VOTAÇÃO
# ============================================================

with aba_votacao:

    mostrar_estatisticas()

    st.divider()


    if not obter_pessoas():

        st.warning(
            "⚠️ Nenhum cadastro foi importado."
        )

        st.info(
            "Importe o cadastro na aba "
            "'📥 IMPORTAR CADASTRO'."
        )


    elif not votacao_ativa():

        st.error(
            "🔴 A VOTAÇÃO FOI ENCERRADA."
        )

        st.info(
            "Acesse o Dashboard para "
            "baixar o resultado final."
        )


    else:

        st.subheader(
            "🎫 Leitura do Crachá"
        )

        st.success(
            "🟢 SISTEMA PRONTO — "
            "APROXIME O CRACHÁ"
        )

        st.caption(
            "O leitor envia o código seguido de ENTER."
        )


        identificador = st.text_input(

            "ID Mat / Mat",

            key="campo_cracha",

            placeholder="Aproxime o crachá..."

        )


        leitor_cracha()


        if st.button(

            "🗳️ REGISTRAR VOTO",

            use_container_width=True,

            type="primary"

        ):

            processar_voto(
                identificador
            )


# ============================================================
# DASHBOARD
# ============================================================

with aba_dashboard:

    st.subheader(
        "📊 Dashboard da Votação"
    )

    estatisticas = (
        obter_estatisticas()
    )


    col1, col2, col3, col4 = (
        st.columns(4)
    )


    with col1:

        st.metric(
            "👥 TOTAL CADASTRADOS",
            estatisticas["total"]
        )


    with col2:

        st.metric(
            "🗳️ VOTOS",
            estatisticas["votaram"]
        )


    with col3:

        st.metric(
            "⏳ NÃO VOTARAM",
            estatisticas["nao_votaram"]
        )


    with col4:

        st.metric(
            "📊 PARTICIPAÇÃO",
            f"{estatisticas['percentual']:.1f}%"
        )


    st.divider()


    st.subheader(
        "📈 Progresso da votação"
    )


    percentual = (
        estatisticas["percentual"]
        / 100
    )


    st.progress(
        min(
            max(
                percentual,
                0.0
            ),
            1.0
        )
    )


    st.write(

        f"**{estatisticas['votaram']}** "
        f"de **{estatisticas['total']}** "
        f"colaboradores já votaram."

    )


    st.divider()


    col1, col2, col3 = (
        st.columns(3)
    )


    with col1:

        st.metric(
            "🔄 TENTATIVAS",
            estatisticas[
                "total_tentativas"
            ]
        )


    with col2:

        st.metric(
            "🚫 DUPLICADOS",
            estatisticas[
                "duplicados"
            ]
        )


    with col3:

        st.metric(
            "⚠️ NÃO ENCONTRADOS",
            estatisticas[
                "nao_encontrados"
            ]
        )


    st.divider()


    if estatisticas["total"] == 0:

        st.info(
            "Nenhum colaborador cadastrado."
        )

    elif (
        estatisticas["votaram"]
        == estatisticas["total"]
    ):

        st.success(
            "🎉 TODOS OS COLABORADORES "
            "JÁ VOTARAM!"
        )

    else:

        st.info(

            f"⏳ Ainda faltam "
            f"**{estatisticas['nao_votaram']}** "
            f"colaborador(es) votar."

        )


    # ========================================================
    # FINALIZAÇÃO
    # ========================================================

    st.divider()

    st.subheader(
        "🏁 Finalizar contabilização"
    )


    if votacao_ativa():

        st.warning(

            "⚠️ Ao finalizar a votação, "
            "novos votos serão bloqueados."

        )


        confirmar = st.checkbox(

            "Estou ciente e desejo finalizar "
            "a contabilização."

        )


        if st.button(

            "🏁 FINALIZAR CONTABILIZAÇÃO",

            use_container_width=True,

            type="primary",

            disabled=not confirmar

        ):

            finalizar_votacao()

            st.success(
                "🏁 VOTAÇÃO FINALIZADA COM SUCESSO!"
            )

            st.rerun()


    else:

        st.success(
            "🔴 Esta votação já está encerrada."
        )


        arquivo_final = (
            gerar_excel_final()
        )


        data_nome = datetime.now().strftime(
            "%Y-%m-%d_%H-%M-%S"
        )


        st.download_button(

            label="📥 BAIXAR PLANILHA FINAL",

            data=arquivo_final,

            file_name=(
                f"resultado_votacao_"
                f"{data_nome}.xlsx"
            ),

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            use_container_width=True,

            type="primary"

        )


# ============================================================
# HISTÓRICO
# ============================================================

with aba_historico:

    st.subheader(
        "📋 Histórico de Leituras"
    )


    pesquisa = st.text_input(

        "🔎 Pesquisar",

        placeholder=(
            "Nome, matrícula, ID Mat ou resultado..."
        )

    )


    dados, colunas = obter_historico(
        pesquisa
    )


    if dados:

        df = pd.DataFrame(

            dados,

            columns=colunas

        )


        resultados = [

            "TODOS",

            "VOTO CONTABILIZADO",

            "VOTO DUPLICADO",

            "COLABORADOR NÃO ENCONTRADO"

        ]


        filtro = st.selectbox(

            "Filtrar resultado",

            resultados

        )


        if filtro != "TODOS":

            df = df[
                df["Resultado"] == filtro
            ]


        st.write(

            f"**{len(df)}** registro(s) encontrado(s)."

        )


        st.dataframe(

            df,

            use_container_width=True,

            hide_index=True

        )


        buffer = io.BytesIO()


        with pd.ExcelWriter(

            buffer,

            engine="openpyxl"

        ) as writer:

            df.to_excel(

                writer,

                index=False,

                sheet_name="Histórico"

            )


        st.download_button(

            "📥 BAIXAR HISTÓRICO EM EXCEL",

            data=buffer.getvalue(),

            file_name="historico_votacao.xlsx",

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            use_container_width=True

        )

    else:

        st.info(
            "Nenhum registro encontrado."
        )


# ============================================================
# IMPORTAÇÃO
# ============================================================

with aba_importacao:

    st.subheader(
        "📥 Importar cadastro"
    )


    st.write(
        "A planilha precisa possuir exatamente:"
    )


    st.code(
        "ID Mat | Mat | Nome | Voto"
    )


    arquivo = st.file_uploader(

        "Selecione a planilha Excel",

        type=["xlsx"]

    )


    if arquivo:

        try:

            df = pd.read_excel(

                arquivo,

                dtype=str

            )


            df.columns = [

                str(col).strip()

                for col in df.columns

            ]


            esperadas = [

                "ID Mat",

                "Mat",

                "Nome",

                "Voto"

            ]


            if list(df.columns) != esperadas:

                st.error(
                    "❌ Estrutura da planilha inválida."
                )

            else:

                for coluna in esperadas:

                    df[coluna] = (

                        df[coluna]

                        .fillna("")

                        .astype(str)

                        .str.strip()

                    )


                df["Voto"] = (

                    df["Voto"]

                    .str.upper()

                )


                invalidos = df[
                    ~df["Voto"].isin(
                        [
                            "",
                            "SIM"
                        ]
                    )
                ]


                if not invalidos.empty:

                    st.error(
                        "❌ Existem valores inválidos "
                        "na coluna Voto."
                    )

                    st.dataframe(
                        invalidos
                    )

                else:

                    st.success(
                        "✅ Estrutura válida."
                    )


                    st.dataframe(

                        df.head(20),

                        use_container_width=True,

                        hide_index=True

                    )


                    st.divider()


                    if st.button(

                        "📥 IMPORTAR / ATUALIZAR CADASTRO",

                        use_container_width=True,

                        type="primary"

                    ):

                        lista = []

                        for _, linha in df.iterrows():

                            lista.append({

                                "id_mat":
                                    linha["ID Mat"],

                                "mat":
                                    linha["Mat"],

                                "nome":
                                    linha["Nome"],

                                "voto":
                                    linha["Voto"],

                            })


                        try:

                            resultado = (
                                importar_pessoas(
                                    lista
                                )
                            )


                            st.success(
                                "✅ CADASTRO IMPORTADO!"
                            )


                            col1, col2, col3 = (
                                st.columns(3)
                            )


                            with col1:

                                st.metric(
                                    "NOVOS",
                                    resultado[0]
                                )


                            with col2:

                                st.metric(
                                    "ATUALIZADOS",
                                    resultado[1]
                                )


                            with col3:

                                st.metric(
                                    "IGNORADOS",
                                    resultado[2]
                                )


                        except Exception as erro:

                            st.error(
                                f"❌ Erro: {erro}"
                            )

        except Exception as erro:

            st.error(
                f"❌ Erro ao ler Excel: {erro}"
            )