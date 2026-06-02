from pathlib import Path
import pandas as pd
import subprocess

# =====================================================
# CONFIGURAÇÃO
# =====================================================

REPO = Path(__file__).parent

PASTA_UPLOAD = REPO / "upload"

ARQ_DEPARA = REPO / "de_para.xlsx"

ARQ_OBRASCONTRATOS = PASTA_UPLOAD / "OBRASCONTRATOS.xlsx"

de_para = pd.read_excel(
    ARQ_DEPARA,
    dtype={
        "Contrato": str,
        "Contrato_SIAD": str,
        "Portal de Compras": str
    }
)

# =====================================================
# GERA CONTRATOS_SIAD E ITENS
# =====================================================

print(f"Procurando arquivo: {ARQ_OBRASCONTRATOS}")
print(f"Existe? {ARQ_OBRASCONTRATOS.exists()}")

if ARQ_OBRASCONTRATOS.exists():

    print("Processando OBRASCONTRATOS.xlsx...")

    abas = {
        "contratos_siad": "contratos_siad.xlsx",
        "itens": "itens.xlsx"
    }

    for aba, arquivo_saida in abas.items():

        print(f"Gerando {arquivo_saida}...")

        df = pd.read_excel(
            ARQ_OBRASCONTRATOS,
            sheet_name=aba,
            header=None
        )

        # remove primeira linha
        df = df.iloc[1:]

        # remove primeira coluna
        df = df.iloc[:, 1:]

        # transforma primeira linha em cabeçalho
        df.columns = df.iloc[0]

        # remove linha utilizada como cabeçalho
        df = df.iloc[1:]

        # -----------------------------------------
        # ADICIONA CONTRATO A PARTIR DO SIAD
        # -----------------------------------------

        if "Número Contrato" in df.columns:

            mapa_contrato = (
                de_para[
                    ["Contrato_SIAD", "Contrato"]
                ]
                .drop_duplicates()
            )

            mapa_contrato["Contrato_SIAD"] = (
                mapa_contrato["Contrato_SIAD"]
                .astype(str)
                .str.strip()
            )

            df["Número Contrato"] = (
                df["Número Contrato"]
                .astype(str)
                .str.strip()
            )

            df = df.merge(
                mapa_contrato,
                left_on="Número Contrato",
                right_on="Contrato_SIAD",
                how="left"
            )

            if "Contrato_SIAD" in df.columns:
                df.drop(
                    columns=["Contrato_SIAD"],
                    inplace=True
                )

            if "Contrato" in df.columns:

                cols = list(df.columns)

                cols.remove("Contrato")

                idx = cols.index("Número Contrato") + 1

                cols.insert(
                    idx,
                    "Contrato"
                )

                df = df[cols]

        destino = PASTA_UPLOAD / arquivo_saida

        df.to_excel(
            destino,
            index=False
        )

        print(f"Criado: {arquivo_saida}")

    try:
        ARQ_OBRASCONTRATOS.unlink()
        print("OBRASCONTRATOS.xlsx removido.")
    except Exception as e:
        print(
            f"Não foi possível remover "
            f"OBRASCONTRATOS.xlsx: {e}"
        )

else:
    print("OBRASCONTRATOS.xlsx não encontrado.")


# =====================================================
# DE PARA
# =====================================================

print("\nAtualizando arquivos...")

de_para.columns = de_para.columns.str.strip()

de_para["Contrato"] = (
    de_para["Contrato"]
    .astype(str)
    .str.strip()
)

de_para["Contrato_SIAD"] = (
    de_para["Contrato_SIAD"]
    .astype(str)
    .str.strip()
)


ARQUIVOS = [
    "Contratos.xlsx",
    "Coordenadas.xlsx",
    "Fiscais.xlsx",
    "Municipios.xlsx",
    "Obra.xlsx",
    "Situacao.xlsx",
    "Trechos.xlsx"
]

contratos_sem_mapeamento = set()

# =====================================================
# PROCESSA ARQUIVOS
# =====================================================

for nome_arquivo in ARQUIVOS:

    caminho = PASTA_UPLOAD / nome_arquivo

    if not caminho.exists():
        print(f"Arquivo não encontrado: {nome_arquivo}")
        continue

    print(f"\nProcessando {nome_arquivo}")

    df = pd.read_excel(
        caminho,
        dtype=str
    )

    df.columns = df.columns.str.strip()

    if "Contrato" not in df.columns:
        print(f"Coluna Contrato não encontrada em {nome_arquivo}")
        continue

    # -------------------------------------------------
    # MERGE COM DE-PARA
    # -------------------------------------------------

    df = df.merge(
        de_para,
        on="Contrato",
        how="left",
        suffixes=("", "_novo")
    )

    # -------------------------------------------------
    # CONTRATOS SEM MAPEAMENTO
    # -------------------------------------------------

    if "Contrato_SIAD_novo" in df.columns:

        nao_mapeados = df[
            df["Contrato_SIAD_novo"]
            .fillna("")
            .str.strip()
            .eq("")
        ]

        if not nao_mapeados.empty:

            contratos_sem_mapeamento.update(
                nao_mapeados["Contrato"]
                .dropna()
                .astype(str)
                .tolist()
            )

    # -------------------------------------------------
    # ATUALIZA CONTRATO_SIAD
    # TODOS OS 7 ARQUIVOS
    # -------------------------------------------------

    if "Contrato_SIAD_novo" in df.columns:

        if "Contrato_SIAD" in df.columns:

            df["Contrato_SIAD"] = (
                df["Contrato_SIAD_novo"]
                .combine_first(df["Contrato_SIAD"])
            )

        else:

            df["Contrato_SIAD"] = df["Contrato_SIAD_novo"]

        df.drop(
            columns=["Contrato_SIAD_novo"],
            inplace=True
        )

    # -------------------------------------------------
    # PORTAL DE COMPRAS
    # SOMENTE CONTRATOS E FISCAIS
    # -------------------------------------------------

    if nome_arquivo in ["Contratos.xlsx", "Fiscais.xlsx"]:

        if "Portal de Compras_novo" in df.columns:

            if "Portal de Compras" in df.columns:

                df["Portal de Compras"] = (
                    df["Portal de Compras_novo"]
                    .combine_first(df["Portal de Compras"])
                )

            else:

                df["Portal de Compras"] = (
                    df["Portal de Compras_novo"]
                )

            df.drop(
                columns=["Portal de Compras_novo"],
                inplace=True
            )

    else:

        if "Portal de Compras_novo" in df.columns:

            df.drop(
                columns=["Portal de Compras_novo"],
                inplace=True
            )

    # -------------------------------------------------
    # REMOVE COLUNAS AUXILIARES
    # -------------------------------------------------

    for coluna in [
        "Contrato_SIAD_novo",
        "Portal de Compras_novo"
    ]:

        if coluna in df.columns:
            df.drop(
                columns=[coluna],
                inplace=True
            )

    # -------------------------------------------------
    # POSICIONA CONTRATO_SIAD
    # LOGO APÓS CONTRATO
    # -------------------------------------------------

    if "Contrato_SIAD" in df.columns:

        colunas = list(df.columns)

        colunas.remove("Contrato_SIAD")

        indice = colunas.index("Contrato") + 1

        colunas.insert(
            indice,
            "Contrato_SIAD"
        )

        df = df[colunas]

    # -------------------------------------------------
    # SALVA
    # -------------------------------------------------

    with pd.ExcelWriter(
        caminho,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False
        )

    print(f"{nome_arquivo} atualizado")

# =====================================================
# GIT
# =====================================================

print("\nVerificando alterações...")

resultado = subprocess.run(
    ["git", "status", "--porcelain"],
    cwd=REPO,
    capture_output=True,
    text=True
)

if not resultado.stdout.strip():
    print("Nenhuma alteração encontrada.")
    raise SystemExit(0)

print("Enviando alterações para o GitHub...")

subprocess.run(
    ["git", "add", "upload"],
    cwd=REPO,
    check=True
)

subprocess.run(
    [
        "git",
        "commit",
        "-m",
        "Atualização automática portal obras"
    ],
    cwd=REPO,
    check=True
)

subprocess.run(
    ["git", "push"],
    cwd=REPO,
    check=True
)

print("GitHub atualizado com sucesso.")