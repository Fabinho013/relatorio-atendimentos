import pandas as pd
import chardet
import unicodedata
from django.shortcuts import render
from django.http import HttpResponse
from io import BytesIO


def gerar_excel_response(metricas):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df_atendente = pd.DataFrame(
            list(metricas["por_atendente"].items()),
            columns=["Atendente", "Quantidade"]
        )
        df_atendente.to_excel(writer, sheet_name="Atendentes", index=False)

        df_dia = pd.DataFrame(
            list(metricas["por_dia"].items()),
            columns=["Data", "Quantidade"]
        )
        df_dia.to_excel(writer, sheet_name="Por Dia", index=False)

        df_media = pd.DataFrame(
            list(metricas["media_por_atendente"].items()),
            columns=["Atendente", "Tempo Médio"]
        )
        df_media.to_excel(writer, sheet_name="Média Atendente", index=False)

        resumo = pd.DataFrame({
            "Métrica": ["Total", "Tempo Médio Geral"],
            "Valor": [metricas["total"], metricas["tempo_medio"]]
        })
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

        for sheet in writer.sheets.values():
            for col in sheet.columns:
                max_length = 0
                col_letter = col[0].column_letter

                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass

                sheet.column_dimensions[col_letter].width = max_length + 2

    output.seek(0)

    response = HttpResponse(
        output,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = 'attachment; filename="relatorio.xlsx"'

    return response


def normalizar_nome(nome):
    if pd.isna(nome):
        return nome

    nome = str(nome).strip()

    # Normaliza unicode
    nome = unicodedata.normalize("NFKD", nome)

    # Remove acentos completamente
    nome = "".join(c for c in nome if not unicodedata.combining(c))

    # Remove espaços duplicados
    nome = " ".join(nome.split())

    return nome.title()


def home(request):
    context = {}

    if request.method == "POST" and "download" in request.POST:
        metricas = request.session.get("metricas")

        if not metricas:
            return render(request, "relatorios/home.html", {
                "erro": "Nenhum relatório disponível para download."
            })

        return gerar_excel_response(metricas)

    elif request.method == "POST":
        arquivos = request.FILES.getlist("arquivo")

        if not arquivos:
            context["erro"] = "Nenhum arquivo enviado."
            return render(request, "relatorios/home.html", context)

        try:
            dfs = []
            erros = []

            for arquivo in arquivos:
                nome = arquivo.name.lower()
                arquivo_bytes = arquivo.read()
                buffer = BytesIO(arquivo_bytes)

                if nome.endswith(".csv"):
                    resultado = chardet.detect(arquivo_bytes)
                    encoding = resultado["encoding"] or "utf-8"

                    try:
                        df = pd.read_csv(
                            buffer,
                            encoding=encoding,
                            sep=";",
                            encoding_errors="replace"
                        )
                    except Exception:
                        buffer.seek(0)
                        df = pd.read_csv(
                            buffer,
                            encoding="latin-1",
                            sep=";",
                            encoding_errors="replace"
                        )

                elif nome.endswith(".xlsx"):
                    buffer.seek(0)
                    df = pd.read_excel(buffer)

                else:
                    erros.append(f"{arquivo.name} ignorado (formato não suportado)")
                    continue

                df.columns = df.columns.str.strip().str.lower()

                colunas_esperadas = {"inicio", "fim", "atendente", "data"}
                colunas_recebidas = set(df.columns)

                faltando = colunas_esperadas - colunas_recebidas

                if faltando:
                    erros.append(
                        f"❌ <strong>{arquivo.name}</strong><br>"
                        f"👉 Faltando: {', '.join(faltando)}<br>"
                        f"👉 Recebidas: {', '.join(colunas_recebidas)}"
                    )
                    continue

                df = df.dropna(subset=["inicio", "fim", "atendente", "data"])

                df["atendente"] = df["atendente"].apply(normalizar_nome)

                df["inicio"] = pd.to_datetime(df["inicio"], format="%H:%M:%S", errors="coerce")
                df["fim"] = pd.to_datetime(df["fim"], format="%H:%M:%S", errors="coerce")
                df["data"] = pd.to_datetime(df["data"], format="%d/%m/%Y", errors="coerce")

                df = df.dropna(subset=["inicio", "fim", "data"])
                df = df[df["fim"] >= df["inicio"]]

                df["duracao"] = df["fim"] - df["inicio"]

                dfs.append(df)

            if not dfs:
                raise ValueError("Nenhum arquivo válido enviado.")

            df_final = pd.concat(dfs, ignore_index=True)

            total = len(df_final)

            por_atendente = df_final["atendente"].value_counts()

            por_dia = (
                df_final["data"]
                .dt.strftime("%d/%m/%Y")
                .value_counts()
                .sort_index()
            )

            media = df_final["duracao"].mean()
            total_seg = int(media.total_seconds()) if pd.notna(media) else 0
            minutos = total_seg // 60
            segundos = total_seg % 60

            media_por_atendente = (
                df_final.groupby("atendente")["duracao"]
                .mean()
                .apply(lambda x: f"{int(x.total_seconds() // 60)}m {int(x.total_seconds() % 60)}s")
            )

            request.session["metricas"] = {
                "total": total,
                "por_atendente": por_atendente.to_dict(),
                "por_dia": por_dia.to_dict(),
                "tempo_medio": f"{minutos}m {segundos}s",
                "media_por_atendente": media_por_atendente.to_dict()
            }

            context = {
                "total": total,
                "por_atendente": por_atendente.items(),
                "por_dia": por_dia.items(),
                "tempo_medio": f"{minutos}m {segundos}s",
                "media_por_atendente": media_por_atendente.items(),
                "erro": " | ".join(erros) if erros else None
            }

        except Exception as e:
            context["erro"] = str(e)

    return render(request, "relatorios/home.html", context)