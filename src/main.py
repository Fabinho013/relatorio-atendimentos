import pandas as pd

ARQUIVO = "data/atendimentos.csv"


def carregar_dados(caminho):
    """Carrega o arquivo CSV e limpa os nomes das colunas."""
    try:
        df = pd.read_csv(caminho, encoding="latin-1", sep=";")
    except FileNotFoundError:
        print("❌ Arquivo não encontrado:", caminho)
        exit()

    # limpar nomes das colunas
    df.columns = df.columns.str.strip().str.lower()

    return df


def limpar_dados(df):
    """Remove dados inválidos e converte horários."""

    # remover linhas com valores vazios importantes
    df = df.dropna(subset=["inicio", "fim", "atendente"])

    # converter horários
    df["inicio"] = pd.to_datetime(df["inicio"], format="%H:%M:%S", errors="coerce")
    df["fim"] = pd.to_datetime(df["fim"], format="%H:%M:%S", errors="coerce")

    # remover horários inválidos
    df = df.dropna(subset=["inicio", "fim"])

    # remover atendimentos com horário incorreto
    df = df[df["fim"] >= df["inicio"]]

    return df


def calcular_duracao(df):
    """Calcula duração dos atendimentos."""
    df["duracao"] = df["fim"] - df["inicio"]
    return df


def mostrar_atendentes(df):
    print("\n" + "-" * 35)
    print("📋 Atendimentos por atendente")
    print("-" * 35)

    por_atendente = df["atendente"].value_counts()

    for atendente, qtd in por_atendente.items():
        print(f"{atendente}: {qtd} atendimentos")


def mostrar_tempo_medio(df):
    media = df["duracao"].mean()

    total_seg = int(media.total_seconds())
    minutos = total_seg // 60
    segundos = total_seg % 60

    print("\n⏱ Tempo médio de atendimento:")
    print(f"{minutos}m {segundos}s")


def mostrar_por_dia(df):
    print("\n" + "-" * 35)
    print("📅 Atendimentos por dia")
    print("-" * 35)

    por_dia = df["data"].value_counts().sort_index()

    for dia, qtd in por_dia.items():
        print(f"{dia}: {qtd} atendimentos")
        
def gerar_excel(df, total):
    with pd.ExcelWriter("relatorio.xlsx", engine="openpyxl") as writer:

        # Atendimentos por atendente
        por_atendente = df["atendente"].value_counts().reset_index()
        por_atendente.columns = ["Atendente", "Quantidade"]
        por_atendente.to_excel(writer, sheet_name="Atendentes", index=False)

        # Atendimentos por dia
        por_dia = df["data"].value_counts().sort_index().reset_index()
        por_dia.columns = ["Data", "Quantidade"]
        por_dia.to_excel(writer, sheet_name="Por Dia", index=False)

        # Resumo
        media = df["duracao"].mean()
        total_seg = int(media.total_seconds())
        minutos = total_seg // 60
        segundos = total_seg % 60

        resumo = pd.DataFrame({
            "Métrica": ["Total de Atendimentos", "Tempo Médio"],
            "Valor": [total, f"{minutos}m {segundos}s"]
        })

        resumo.to_excel(writer, sheet_name="Resumo", index=False)

    print("✅ Relatório Excel gerado: relatorio.xlsx")
    
def gerar_relatorio(df, total):
    with open("relatorio.txt", "w", encoding="utf-8") as f:
        f.write("RELATÓRIO DE ATENDIMENTOS\n")
        f.write("=" * 40 + "\n\n")

        f.write("Atendimentos por atendente:\n")
        por_atendente = df["atendente"].value_counts()

        for atendente, qtd in por_atendente.items():
            f.write(f"{atendente}: {qtd} atendimentos\n")

        f.write("\n")

        f.write(f"Total de atendimentos: {total}\n\n")

        media = df["duracao"].mean()
        total_seg = int(media.total_seconds())
        minutos = total_seg // 60
        segundos = total_seg % 60

        f.write(f"Tempo médio: {minutos}m {segundos}s\n\n")

        f.write("Atendimentos por dia:\n")
        por_dia = df["data"].value_counts().sort_index()

        for dia, qtd in por_dia.items():
            f.write(f"{dia}: {qtd} atendimentos\n")

    print("\n✅ Relatório gerado: relatorio.txt")


def main():
    print("📊 Sistema de Relatório de Atendimentos iniciado")

    df = carregar_dados(ARQUIVO)
    df = limpar_dados(df)
    df = calcular_duracao(df)

    total = len(df)

    print(f"\n📊 Total de atendimentos: {total}")

    gerar_excel(df, total)

    mostrar_atendentes(df)

    mostrar_tempo_medio(df)

    mostrar_por_dia(df)

    gerar_relatorio(df, total)

if __name__ == "__main__":
    main()