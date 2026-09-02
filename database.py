import os

from datetime import datetime, timezone

from supabase import create_client


# ============================================================
# CONEXÃO
# ============================================================

SUPABASE_URL = os.environ.get(
    "SUPABASE_URL"
)

SUPABASE_KEY = os.environ.get(
    "SUPABASE_KEY"
)


if not SUPABASE_URL or not SUPABASE_KEY:

    raise RuntimeError(
        "SUPABASE_URL e SUPABASE_KEY não foram configuradas."
    )


supabase = create_client(
    SUPABASE_URL,
    SUPABASE_KEY
)


# ============================================================
# NORMALIZAÇÃO
# ============================================================

def normalizar_identificador(valor):

    if valor is None:

        return ""

    valor = str(valor).strip()

    if not valor:

        return ""

    if valor.endswith(".0"):

        valor = valor[:-2]

    valor = valor.replace(
        " ",
        ""
    )

    valor = valor.lstrip("0")

    if valor == "":

        return "0"

    return valor


# ============================================================
# STATUS DA VOTAÇÃO
# ============================================================

def votacao_ativa():

    resposta = (
        supabase
        .table("configuracao_votacao")
        .select("votacao_ativa")
        .eq("id", 1)
        .single()
        .execute()
    )

    if not resposta.data:

        return False

    return bool(
        resposta.data["votacao_ativa"]
    )


def finalizar_votacao():

    agora = datetime.now(
        timezone.utc
    ).isoformat()

    resposta = (
        supabase
        .table("configuracao_votacao")
        .update({
            "votacao_ativa": False,
            "data_finalizacao": agora
        })
        .eq("id", 1)
        .execute()
    )

    return resposta.data


def reabrir_votacao():

    resposta = (
        supabase
        .table("configuracao_votacao")
        .update({
            "votacao_ativa": True,
            "data_finalizacao": None
        })
        .eq("id", 1)
        .execute()
    )

    return resposta.data


# ============================================================
# IMPORTAR PESSOAS
# ============================================================

def importar_pessoas(lista_pessoas):

    inseridos = 0
    atualizados = 0
    ignorados = 0

    for pessoa in lista_pessoas:

        id_mat_original = pessoa.get(
            "id_mat",
            ""
        )

        mat_original = pessoa.get(
            "mat",
            ""
        )

        nome = str(
            pessoa.get(
                "nome",
                ""
            )
        ).strip()

        id_mat = normalizar_identificador(
            id_mat_original
        )

        mat = str(
            mat_original
        ).strip()

        if (
            not id_mat
            or not mat
            or not nome
        ):

            ignorados += 1

            continue

        existente = (
            supabase
            .table("pessoas")
            .select("id")
            .eq("id_mat", id_mat)
            .execute()
        )

        if existente.data:

            supabase.table(
                "pessoas"
            ).update({

                "mat": mat,

                "nome": nome

            }).eq(
                "id_mat",
                id_mat
            ).execute()

            atualizados += 1

        else:

            voto = str(
                pessoa.get(
                    "voto",
                    ""
                )
            ).strip().upper()

            if voto not in (
                "",
                "SIM"
            ):

                voto = ""

            supabase.table(
                "pessoas"
            ).insert({

                "id_mat": id_mat,

                "mat": mat,

                "nome": nome,

                "voto": voto,

                "data_voto": None

            }).execute()

            inseridos += 1

    return (
        inseridos,
        atualizados,
        ignorados
    )


# ============================================================
# LOCALIZAR COLABORADOR
# ============================================================

def localizar_colaborador(
    identificador
):

    identificador_normalizado = (
        normalizar_identificador(
            identificador
        )
    )

    if not identificador_normalizado:

        return None

    pessoas = (
        supabase
        .table("pessoas")
        .select("*")
        .execute()
    )

    for pessoa in pessoas.data:

        id_mat = normalizar_identificador(
            pessoa.get("id_mat")
        )

        mat = normalizar_identificador(
            pessoa.get("mat")
        )

        if (
            identificador_normalizado == id_mat
            or identificador_normalizado == mat
        ):

            return pessoa

    return None


# ============================================================
# REGISTRAR VOTO
# ============================================================

def registrar_voto(
    identificador
):

    if not votacao_ativa():

        return {
            "resultado": "encerrada"
        }

    colaborador = localizar_colaborador(
        identificador
    )

    if colaborador is None:

        registrar_historico(
            identificador=identificador,
            id_mat="",
            mat="",
            nome="",
            resultado="COLABORADOR NÃO ENCONTRADO"
        )

        return {
            "resultado": "nao_encontrado",
            "identificador": identificador
        }

    voto = str(
        colaborador.get(
            "voto",
            ""
        )
    ).strip().upper()

    if voto == "SIM":

        registrar_historico(
            identificador=identificador,
            id_mat=str(
                colaborador["id_mat"]
            ),
            mat=str(
                colaborador["mat"]
            ),
            nome=str(
                colaborador["nome"]
            ),
            resultado="VOTO DUPLICADO"
        )

        return {
            "resultado": "duplicado",
            **colaborador
        }

    agora = datetime.now(
        timezone.utc
    ).isoformat()

    (
        supabase
        .table("pessoas")
        .update({

            "voto": "SIM",

            "data_voto": agora

        })
        .eq(
            "id",
            colaborador["id"]
        )
        .execute()
    )

    registrar_historico(
        identificador=identificador,
        id_mat=str(
            colaborador["id_mat"]
        ),
        mat=str(
            colaborador["mat"]
        ),
        nome=str(
            colaborador["nome"]
        ),
        resultado="VOTO CONTABILIZADO"
    )

    return {
        "resultado": "contabilizado",
        **colaborador
    }


# ============================================================
# HISTÓRICO
# ============================================================

def registrar_historico(
    identificador,
    id_mat,
    mat,
    nome,
    resultado
):

    agora = datetime.now(
        timezone.utc
    ).isoformat()

    resposta = (
        supabase
        .table("historico")
        .insert({

            "identificador": identificador,

            "id_mat": id_mat,

            "mat": mat,

            "nome": nome,

            "resultado": resultado,

            "data_hora": agora

        })
        .execute()
    )

    return resposta.data


def obter_historico(
    pesquisa=""
):

    consulta = (
        supabase
        .table("historico")
        .select("*")
        .order(
            "id",
            desc=True
        )
    )

    if pesquisa:

        # Busca ampla feita localmente após recuperar os registros.
        resposta = consulta.execute()

        termo = str(
            pesquisa
        ).lower().strip()

        dados = []

        for item in resposta.data:

            texto = " ".join([

                str(
                    item.get(
                        "identificador",
                        ""
                    )
                ),

                str(
                    item.get(
                        "id_mat",
                        ""
                    )
                ),

                str(
                    item.get(
                        "mat",
                        ""
                    )
                ),

                str(
                    item.get(
                        "nome",
                        ""
                    )
                ),

                str(
                    item.get(
                        "resultado",
                        ""
                    )
                ),

            ]).lower()

            if termo in texto:

                dados.append(item)

    else:

        resposta = consulta.execute()

        dados = resposta.data

    resultado = []

    for item in dados:

        resultado.append([

            item.get("id"),

            item.get("identificador"),

            item.get("id_mat"),

            item.get("mat"),

            item.get("nome"),

            item.get("resultado"),

            item.get("data_hora"),

        ])

    colunas = [

        "ID",

        "Identificador",

        "ID Mat",

        "Mat",

        "Nome",

        "Resultado",

        "Data/Hora"

    ]

    return resultado, colunas


# ============================================================
# ESTATÍSTICAS
# ============================================================

def obter_estatisticas():

    resposta = (
        supabase
        .table("pessoas")
        .select(
            "voto",
            count="exact"
        )
        .execute()
    )

    pessoas = resposta.data

    total = len(pessoas)

    votaram = sum(

        1

        for pessoa in pessoas

        if str(
            pessoa.get(
                "voto",
                ""
            )
        ).upper() == "SIM"

    )

    nao_votaram = (
        total - votaram
    )

    historico = (
        supabase
        .table("historico")
        .select(
            "resultado"
        )
        .execute()
    )

    registros = historico.data

    duplicados = sum(

        1

        for item in registros

        if item["resultado"]
        == "VOTO DUPLICADO"

    )

    nao_encontrados = sum(

        1

        for item in registros

        if item["resultado"]
        == "COLABORADOR NÃO ENCONTRADO"

    )

    total_tentativas = len(
        registros
    )

    percentual = (

        votaram
        / total
        * 100

        if total > 0

        else 0

    )

    return {

        "total": total,

        "votaram": votaram,

        "nao_votaram": nao_votaram,

        "duplicados": duplicados,

        "nao_encontrados": nao_encontrados,

        "total_tentativas": total_tentativas,

        "percentual": percentual,

    }


# ============================================================
# DADOS PARA EXPORTAÇÃO
# ============================================================

def obter_pessoas():

    resposta = (
        supabase
        .table("pessoas")
        .select("*")
        .order(
            "id_mat"
        )
        .execute()
    )

    return resposta.data


def obter_configuracao():

    resposta = (
        supabase
        .table("configuracao_votacao")
        .select("*")
        .eq(
            "id",
            1
        )
        .single()
        .execute()
    )

    return resposta.data