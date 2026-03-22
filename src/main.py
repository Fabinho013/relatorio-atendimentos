import pandas as pd
import logging
import sys

ARQUIVO = "data/atendimentos.csv"

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s"
)


def carregar_dados(caminho):
    """Carrega o arquivo CSV e limpa os nomes das colunas."""
    try:
        df = pd.read_csv(caminho, encoding="latin-1", sep=";")
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    df.columns = df.columns.str.strip().str.lower()

    colunas_esperadas = {"inicio", "fim", "atendente", "data"}
    if not colunas_esperadas.issubset(df.columns):
        raise ValueError("❌ O arquivo não contém todas as colunas obrigatórias")

    return df


def limpar_dados(df):
    """Remove dados inválidos e converte horários."""
    df = df.copy()

    df = df.dropna(subset=["inicio", "fim", "atendente", "data"])

    df["inicio"] = pd.to_datetime(df["inicio"], format="%H:%M:%S", errors="coerce")
    df["fim"] = pd.to_datetime(df["fim"], format="%H:%M:%S", errors="coerce")
    df["data"] = pd.to_datetime(df["data"], errors="coerce")

    df = df.dropna(subset=["inicio", "fim", "data"])

    df = df[df["fim"] >= df["inicio"]]

    return df


def calcular_duracao(df):
    """Calcula duração dos atendimentos."""
    df = df.copy()
    df["duracao"] = df["fim"] - df["inicio"]
    return df


def formatar_tempo_medio(df):
    """Retorna tempo médio formatado."""
    media = df["duracao"].mean()

    if pd.isna(media):
        return 0, 0

    total_seg = int(media.total_seconds())
    return total_seg // 60, total_seg % 60


def obter_metricas(df):
    """Centraliza cálculos principais."""
    total = len(df)
    por_atendente = df["atendente"].value_counts()
    por_dia = df["data"].dt.date.value_counts().sort_index()
    minutos, segundos = formatar_tempo_medio(df)

    return {
        "total": total,
        "por_atendente": por_atendente,
        "por_dia": por_dia,
        "tempo_medio": (minutos, segundos)
    }


def mostrar_relatorio_terminal(metricas):
    print("\n📊 Total de atendimentos:", metricas["total"])

    print("\n" + "-" * 35)
    print("📋 Atendimentos por atendente")
    print("-" * 35)
    for atendente, qtd in metricas["por_atendente"].items():
        print(f"{atendente}: {qtd} atendimentos")

    print("\n⏱ Tempo médio de atendimento:")
    m, s = metricas["tempo_medio"]
    print(f"{m}m {s}s")

    print("\n" + "-" * 35)
    print("📅 Atendimentos por dia")
    print("-" * 35)
    for dia, qtd in metricas["por_dia"].items():
        print(f"{dia}: {qtd} atendimentos")


def gerar_excel(df, metricas):
    with pd.ExcelWriter("relatorio.xlsx", engine="openpyxl") as writer:

        # Atendentes
        df_atendente = metricas["por_atendente"].reset_index()
        df_atendente.columns = ["Atendente", "Quantidade"]
        df_atendente.to_excel(writer, sheet_name="Atendentes", index=False)

        # Por dia
        df_dia = metricas["por_dia"].reset_index()
        df_dia.columns = ["Data", "Quantidade"]
        df_dia.to_excel(writer, sheet_name="Por Dia", index=False)

        # Resumo
        m, s = metricas["tempo_medio"]
        resumo = pd.DataFrame({
            "Métrica": ["Total de Atendimentos", "Tempo Médio"],
            "Valor": [metricas["total"], f"{m}m {s}s"]
        })

        resumo.to_excel(writer, sheet_name="Resumo", index=False)

    logging.info("Relatório Excel gerado: relatorio.xlsx")


def gerar_txt(metricas):
    with open("relatorio.txt", "w", encoding="utf-8") as f:
        f.write("RELATÓRIO DE ATENDIMENTOS\n")
        f.write("=" * 40 + "\n\n")

        f.write("Atendimentos por atendente:\n")
        for atendente, qtd in metricas["por_atendente"].items():
            f.write(f"{atendente}: {qtd} atendimentos\n")

        f.write("\n")
        f.write(f"Total de atendimentos: {metricas['total']}\n\n")

        m, s = metricas["tempo_medio"]
        f.write(f"Tempo médio: {m}m {s}s\n\n")

        f.write("Atendimentos por dia:\n")
        for dia, qtd in metricas["por_dia"].items():
            f.write(f"{dia}: {qtd} atendimentos\n")

    logging.info("Relatório TXT gerado: relatorio.txt")


def main():
    logging.info("Sistema iniciado")

    try:
        df = carregar_dados(ARQUIVO)
        df = limpar_dados(df)
        df = calcular_duracao(df)

        metricas = obter_metricas(df)

        mostrar_relatorio_terminal(metricas)
        gerar_excel(df, metricas)
        gerar_txt(metricas)

    except Exception as e:
        logging.error(f"Erro: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()