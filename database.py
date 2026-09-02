import os
from datetime import datetime, timezone

import pandas as pd
import streamlit as st
from supabase import create_client


# ============================================================
# CONFIGURAÇÃO DO SUPABASE
# ============================================================

def obter_secret(nome):
    """
    Obtém uma configuração primeiro do Streamlit Secrets
    e depois das variáveis de ambiente.
    """

    try:
        valor = st.secrets.get(
            nome,
            None
        )

        if valor:
            return valor

    except Exception:
        pass

    return os.getenv(
        nome,
        ""
    )


SUPABASE_URL = obter_secret(
    "SUPABASE_URL"
)

SUPABASE_KEY = obter_secret(
    "SUPABASE_KEY"
)


if not SUPABASE_URL:
    raise RuntimeError(
        "SUPABASE_URL não configurada."
    )


if not SUPABASE_KEY:
    raise RuntimeError(
        "SUPABASE_KEY não configurada."
    )


supabase = create_client(
    SUPABASE_URL,
    SUPABASE_KEY
)


# ============================================================
# CONFIGURAÇÃO DE PAGINAÇÃO
# ============================================================

# O Supabase/PostgREST pode retornar no máximo 1000 registros
# por consulta dependendo da configuração do projeto.
#
# Todas as consultas grandes deste arquivo utilizam paginação.
# Assim, o sistema funciona com:
#
# 1.433 registros
# 2.000 registros
# 5.000 registros
# 10.000 registros
# etc.
#
# sem depender do limite de uma única consulta.

TAMANHO_PAGINA = 1000


# ============================================================
# NORMALIZAÇÃO
# ============================================================

def normalizar_identificador(valor):
    """
    Normaliza um identificador.

    Regras:
    - remove espaços;
    - trata valores vazios;
    - remove .0 de valores numéricos vindos do Excel;
    - remove zeros à esquerda.

    Exemplos:

    000123 -> 123
    00123  -> 123
    123    -> 123
    123.0  -> 123
    """

    if valor is None:
        return ""

    try:

        if pd.isna(valor):
            return ""

    except Exception:

        pass

    texto = str(
        valor
    ).strip()

    if not texto:
        return ""

    texto = texto.replace(
        "\u00a0",
        ""
    )

    texto = texto.replace(
        " ",
        ""
    )

    # --------------------------------------------------------
    # Corrigir números vindos do Excel
    # --------------------------------------------------------

    if texto.endswith(".0"):

        parte = texto[:-2]

        if parte.isdigit():
            texto = parte

    # --------------------------------------------------------
    # Remover zeros à esquerda
    # --------------------------------------------------------

    texto = texto.lstrip("0")

    # Mantemos "0" como resultado para o valor numérico 0.
    if texto == "":
        return "0"

    return texto


def limpar_texto(valor):
    """
    Limpa texto geral.
    """

    if valor is None:
        return ""

    try:

        if pd.isna(valor):
            return ""

    except Exception:

        pass

    return str(
        valor
    ).strip()


# ============================================================
# CONFIGURAÇÃO DA VOTAÇÃO
# ============================================================

def obter_configuracao():

    resposta = (
        supabase
        .table("configuracao_votacao")
        .select("*")
        .eq(
            "id",
            1
        )
        .limit(1)
        .execute()
    )

    if not resposta.data:
        return None

    return resposta.data[0]


def votacao_ativa():

    configuracao = obter_configuracao()

    if not configuracao:
        return False

    return bool(
        configuracao.get(
            "votacao_ativa",
            False
        )
    )


def finalizar_votacao():

    resposta = supabase.rpc(
        "finalizar_votacao"
    ).execute()

    if not resposta.data:

        return {
            "sucesso": False,
            "mensagem":
                "Não foi possível finalizar a votação.",
        }

    if isinstance(
        resposta.data,
        list
    ):

        return resposta.data[0]

    return resposta.data


def reabrir_votacao():

    resposta = (
        supabase
        .table("configuracao_votacao")
        .update(
            {
                "votacao_ativa": True,
                "data_finalizacao": None,
            }
        )
        .eq(
            "id",
            1
        )
        .execute()
    )

    return bool(
        resposta.data
    )


# ============================================================
# BUSCAR TODAS AS PESSOAS
# ============================================================

def _buscar_todas_pessoas():
    """
    Busca TODOS os colaboradores cadastrados.

    IMPORTANTE:
    Não utiliza uma única consulta, pois o Supabase pode
    limitar o retorno a 1000 registros.

    A função pagina automaticamente:

        0    - 999
        1000 - 1999
        2000 - 2999
        ...

    até encontrar o final dos dados.
    """

    dados = []

    inicio = 0

    while True:

        fim = (
            inicio
            + TAMANHO_PAGINA
            - 1
        )

        resposta = (
            supabase
            .table("pessoas")
            .select(
                """
                id,
                id_mat,
                mat,
                nome,
                voto,
                data_voto
                """
            )
            .order(
                "id"
            )
            .range(
                inicio,
                fim
            )
            .execute()
        )

        lote = (
            resposta.data
            or []
        )

        dados.extend(
            lote
        )

        # ----------------------------------------------------
        # Se vier menos que 1000, chegamos ao final.
        # ----------------------------------------------------

        if len(lote) < TAMANHO_PAGINA:
            break

        inicio += TAMANHO_PAGINA

    return dados


# ============================================================
# IMPORTAÇÃO DE PESSOAS
# ============================================================

def importar_pessoas(pessoas):
    """
    Importa um DataFrame ou lista de dicionários.

    A validação ocorre COMPLETAMENTE antes da gravação.

    Isso evita:
    - importação parcial;
    - IDs duplicados;
    - dados inválidos;
    - inconsistência no Supabase.

    Os votos SIM já existentes no banco são preservados.
    """

    # ========================================================
    # VALIDAR TIPO
    # ========================================================

    if pessoas is None:

        raise ValueError(
            "Nenhum cadastro foi fornecido."
        )

    # ========================================================
    # CONVERTER PARA DATAFRAME
    # ========================================================

    if isinstance(
        pessoas,
        pd.DataFrame
    ):

        df = pessoas.copy()

    elif isinstance(
        pessoas,
        list
    ):

        df = pd.DataFrame(
            pessoas
        )

    else:

        raise TypeError(
            "importar_pessoas() aceita somente "
            "pandas.DataFrame ou lista de dicionários."
        )

    # ========================================================
    # COLUNAS OBRIGATÓRIAS
    # ========================================================

    colunas_esperadas = [
        "ID Mat",
        "Mat",
        "Nome",
        "Voto",
    ]

    faltantes = [
        coluna
        for coluna in colunas_esperadas
        if coluna not in df.columns
    ]

    if faltantes:

        raise ValueError(
            "Colunas ausentes: "
            +
            ", ".join(
                faltantes
            )
        )

    df = df[
        colunas_esperadas
    ].copy()

    # ========================================================
    # LIMPAR DADOS
    # ========================================================

    for coluna in colunas_esperadas:

        df[coluna] = df[coluna].apply(
            limpar_texto
        )

    # ========================================================
    # REMOVER LINHAS TOTALMENTE VAZIAS
    # ========================================================

    df = df[
        ~(
            (df["ID Mat"] == "")
            &
            (df["Mat"] == "")
            &
            (df["Nome"] == "")
            &
            (df["Voto"] == "")
        )
    ].copy()

    # ========================================================
    # VALIDAR CAMPOS OBRIGATÓRIOS
    # ========================================================

    problemas = []

    for indice, linha in df.iterrows():

        linha_excel = (
            int(indice) + 2
        )

        id_original = limpar_texto(
            linha["ID Mat"]
        )

        mat_original = limpar_texto(
            linha["Mat"]
        )

        nome_original = limpar_texto(
            linha["Nome"]
        )

        voto_original = limpar_texto(
            linha["Voto"]
        ).upper()

        if not id_original:

            problemas.append(
                {
                    "linha": linha_excel,
                    "tipo": "ID Mat vazio",
                    "valor": "",
                }
            )

        if not mat_original:

            problemas.append(
                {
                    "linha": linha_excel,
                    "tipo": "Mat vazio",
                    "valor": "",
                }
            )

        if not nome_original:

            problemas.append(
                {
                    "linha": linha_excel,
                    "tipo": "Nome vazio",
                    "valor": "",
                }
            )

        if voto_original not in {
            "",
            "SIM",
        }:

            problemas.append(
                {
                    "linha": linha_excel,
                    "tipo": "Voto inválido",
                    "valor": voto_original,
                }
            )

    # ========================================================
    # PARAR SE EXISTIREM ERROS BÁSICOS
    # ========================================================

    if problemas:

        detalhes = []

        for problema in problemas:

            detalhes.append(
                (
                    f"Linha {problema['linha']}: "
                    f"{problema['tipo']}"
                    +
                    (
                        f" -> {problema['valor']}"
                        if problema["valor"]
                        else ""
                    )
                )
            )

        raise ValueError(
            "O cadastro possui erros:\n\n"
            +
            "\n".join(
                detalhes
            )
        )

    # ========================================================
    # NORMALIZAR IDs
    # ========================================================

    df["_id_mat_normalizado"] = (
        df["ID Mat"]
        .apply(
            normalizar_identificador
        )
    )

    # ========================================================
    # VALIDAR IDS VAZIOS APÓS NORMALIZAÇÃO
    # ========================================================

    ids_vazios = df[
        df["_id_mat_normalizado"] == ""
    ]

    if not ids_vazios.empty:

        linhas = []

        for indice in ids_vazios.index:

            linhas.append(
                str(
                    int(indice) + 2
                )
            )

        raise ValueError(
            "Existem IDs Mat inválidos nas "
            "linhas: "
            +
            ", ".join(
                linhas
            )
        )

    # ========================================================
    # IDENTIFICAR DUPLICIDADES
    # ========================================================

    contagem_ids = (
        df["_id_mat_normalizado"]
        .value_counts()
    )

    ids_duplicados = (
        contagem_ids[
            contagem_ids > 1
        ]
    )

    if not ids_duplicados.empty:

        detalhes_duplicados = []

        for id_normalizado in (
            ids_duplicados.index
        ):

            ocorrencias = df[
                df["_id_mat_normalizado"]
                ==
                id_normalizado
            ]

            linhas = []

            valores_originais = []

            for indice, linha in (
                ocorrencias.iterrows()
            ):

                linhas.append(
                    str(
                        int(indice) + 2
                    )
                )

                valores_originais.append(
                    limpar_texto(
                        linha["ID Mat"]
                    )
                )

            detalhes_duplicados.append(
                (
                    f"ID normalizado '{id_normalizado}' "
                    f"nas linhas "
                    f"{', '.join(linhas)} "
                    f"(valores originais: "
                    f"{', '.join(valores_originais)})"
                )
            )

        raise ValueError(
            "Foram encontrados IDs Mat duplicados "
            "após a normalização:\n\n"
            +
            "\n".join(
                detalhes_duplicados
            )
            +
            "\n\n"
            "Corrija os IDs duplicados no Excel "
            "antes de importar."
        )

    # ========================================================
    # BUSCAR DADOS EXISTENTES
    #
    # IMPORTANTE:
    # _buscar_todas_pessoas() já utiliza paginação.
    # Portanto, aqui recuperamos TODOS os registros,
    # inclusive quando houver mais de 1000.
    # ========================================================

    pessoas_existentes = (
        _buscar_todas_pessoas()
    )

    # ========================================================
    # MAPA DOS VOTOS EXISTENTES
    # ========================================================

    votos_existentes = {}

    for pessoa_existente in (
        pessoas_existentes
    ):

        id_existente = (
            normalizar_identificador(
                pessoa_existente.get(
                    "id_mat",
                    ""
                )
            )
        )

        if not id_existente:
            continue

        votos_existentes[
            id_existente
        ] = {
            "voto":
                limpar_texto(
                    pessoa_existente.get(
                        "voto",
                        ""
                    )
                ).upper(),

            "data_voto":
                pessoa_existente.get(
                    "data_voto"
                ),
        }

    # ========================================================
    # PREPARAR REGISTROS
    # ========================================================

    agora = datetime.now(
        timezone.utc
    ).isoformat()

    registros_importacao = []

    for indice, linha in (
        df.iterrows()
    ):

        id_original = (
            limpar_texto(
                linha["ID Mat"]
            )
        )

        mat_original = (
            limpar_texto(
                linha["Mat"]
            )
        )

        nome = (
            limpar_texto(
                linha["Nome"]
            )
        )

        voto_excel = (
            limpar_texto(
                linha["Voto"]
            ).upper()
        )

        id_mat = (
            linha["_id_mat_normalizado"]
        )

        mat = normalizar_identificador(
            mat_original
        )

        # ----------------------------------------------------
        # VERIFICAR REGISTRO EXISTENTE
        # ----------------------------------------------------

        existente = (
            votos_existentes.get(
                id_mat
            )
        )

        if existente:

            voto_banco = (
                existente.get(
                    "voto",
                    ""
                )
            )

            data_voto_banco = (
                existente.get(
                    "data_voto"
                )
            )

            # ----------------------------------------------
            # SE JÁ VOTOU, PRESERVAR
            # ----------------------------------------------

            if voto_banco == "SIM":

                voto_final = "SIM"

                data_voto_final = (
                    data_voto_banco
                )

            # ----------------------------------------------
            # SE EXCEL TEM SIM
            # ----------------------------------------------

            elif voto_excel == "SIM":

                voto_final = "SIM"

                data_voto_final = (
                    data_voto_banco
                    or agora
                )

            # ----------------------------------------------
            # SEM VOTO
            # ----------------------------------------------

            else:

                voto_final = ""

                data_voto_final = None

        else:

            if voto_excel == "SIM":

                voto_final = "SIM"

                data_voto_final = agora

            else:

                voto_final = ""

                data_voto_final = None

        registros_importacao.append(
            {
                "id_mat": id_mat,
                "mat": mat,
                "nome": nome,
                "voto": voto_final,
                "data_voto": data_voto_final,
            }
        )

    # ========================================================
    # GRAVAÇÃO EM LOTES
    # ========================================================

    tamanho_lote = 500

    total = len(
        registros_importacao
    )

    for inicio in range(
        0,
        total,
        tamanho_lote
    ):

        lote = registros_importacao[
            inicio:
            inicio + tamanho_lote
        ]

        (
            supabase
            .table("pessoas")
            .upsert(
                lote,
                on_conflict="id_mat"
            )
            .execute()
        )

    return {
        "sucesso": True,
        "quantidade": total,
    }


# ============================================================
# REGISTRAR VOTO
# ============================================================

def registrar_voto(
    identificador
):

    resposta = supabase.rpc(
        "registrar_voto",
        {
            "p_identificador":
                str(
                    identificador
                )
        }
    ).execute()

    if not resposta.data:
        return None

    if isinstance(
        resposta.data,
        list
    ):

        return resposta.data[0]

    return resposta.data


# ============================================================
# REGISTRAR HISTÓRICO
# ============================================================

def registrar_historico(
    identificador,
    id_mat=None,
    mat=None,
    nome=None,
    resultado=None
):

    registro = {
        "identificador":
            limpar_texto(
                identificador
            ),

        "id_mat":
            limpar_texto(
                id_mat
            ),

        "mat":
            limpar_texto(
                mat
            ),

        "nome":
            limpar_texto(
                nome
            ),

        "resultado":
            limpar_texto(
                resultado
            ),

        "data_hora":
            datetime.now(
                timezone.utc
            ).isoformat(),
    }

    resposta = (
        supabase
        .table("historico")
        .insert(
            registro
        )
        .execute()
    )

    return resposta.data


# ============================================================
# HISTÓRICO
# ============================================================

def obter_historico():
    """
    Busca TODOS os registros do histórico.

    A consulta é paginada para evitar o limite de 1000
    registros do Supabase/PostgREST.
    """

    dados = []

    inicio = 0

    while True:

        fim = (
            inicio
            + TAMANHO_PAGINA
            - 1
        )

        resposta = (
            supabase
            .table("historico")
            .select("*")
            .order(
                "data_hora",
                desc=True
            )
            .range(
                inicio,
                fim
            )
            .execute()
        )

        lote = (
            resposta.data
            or []
        )

        dados.extend(
            lote
        )

        if len(lote) < TAMANHO_PAGINA:
            break

        inicio += TAMANHO_PAGINA

    # --------------------------------------------------------
    # Garantir DataFrame consistente mesmo sem registros
    # --------------------------------------------------------

    if not dados:

        return pd.DataFrame(
            columns=[
                "id",
                "identificador",
                "id_mat",
                "mat",
                "nome",
                "resultado",
                "data_hora",
            ]
        )

    return pd.DataFrame(
        dados
    )


# ============================================================
# PESSOAS
# ============================================================

def obter_pessoas():
    """
    Busca TODOS os colaboradores.

    Nunca fica limitado aos primeiros 1000 registros.
    """

    dados = (
        _buscar_todas_pessoas()
    )

    # --------------------------------------------------------
    # Garantir DataFrame consistente mesmo sem registros
    # --------------------------------------------------------

    if not dados:

        return pd.DataFrame(
            columns=[
                "id",
                "id_mat",
                "mat",
                "nome",
                "voto",
                "data_voto",
            ]
        )

    return pd.DataFrame(
        dados
    )


# ============================================================
# ESTATÍSTICAS
# ============================================================

def obter_estatisticas():
    """
    Calcula as estatísticas utilizando TODOS os registros.

    IMPORTANTE:

    A versão anterior fazia:

        .table("pessoas")
        .select("id,voto")
        .execute()

    Essa consulta podia retornar somente 1000 registros.

    Agora utilizamos as funções paginadas, garantindo que
    1433, 2000, 5000 ou mais colaboradores sejam contabilizados.
    """

    # ========================================================
    # BUSCAR TODAS AS PESSOAS
    # ========================================================

    pessoas = (
        _buscar_todas_pessoas()
    )

    # ========================================================
    # TOTAL DE CADASTRADOS
    # ========================================================

    total = len(
        pessoas
    )

    # ========================================================
    # CONTAR VOTOS
    # ========================================================

    votos = 0

    for pessoa in pessoas:

        voto = (
            limpar_texto(
                pessoa.get(
                    "voto",
                    ""
                )
            ).upper()
        )

        if voto == "SIM":

            votos += 1

    # ========================================================
    # NÃO VOTARAM
    # ========================================================

    nao_votaram = (
        total - votos
    )

    # ========================================================
    # PARTICIPAÇÃO
    # ========================================================

    if total > 0:

        participacao = (
            votos / total
        ) * 100

    else:

        participacao = 0

    # ========================================================
    # BUSCAR TODO O HISTÓRICO
    # ========================================================

    historico = []

    inicio = 0

    while True:

        fim = (
            inicio
            + TAMANHO_PAGINA
            - 1
        )

        resposta_historico = (
            supabase
            .table("historico")
            .select(
                "resultado"
            )
            .order(
                "id"
            )
            .range(
                inicio,
                fim
            )
            .execute()
        )

        lote = (
            resposta_historico.data
            or []
        )

        historico.extend(
            lote
        )

        if len(lote) < TAMANHO_PAGINA:
            break

        inicio += TAMANHO_PAGINA

    # ========================================================
    # TOTAL DE TENTATIVAS
    # ========================================================

    tentativas = len(
        historico
    )

    # ========================================================
    # CONTADORES DO HISTÓRICO
    # ========================================================

    duplicados = 0

    nao_encontrados = 0

    for registro in historico:

        resultado = (
            limpar_texto(
                registro.get(
                    "resultado",
                    ""
                )
            ).upper()
        )

        if resultado == "VOTO DUPLICADO":

            duplicados += 1

        elif (
            resultado
            ==
            "COLABORADOR NÃO ENCONTRADO"
        ):

            nao_encontrados += 1

    # ========================================================
    # RETORNO
    # ========================================================

    return {
        "total":
            total,

        "votos":
            votos,

        "nao_votaram":
            nao_votaram,

        "participacao":
            participacao,

        "tentativas":
            tentativas,

        "duplicados":
            duplicados,

        "nao_encontrados":
            nao_encontrados,
    }