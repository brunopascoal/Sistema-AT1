import streamlit as st
import pandas as pd
import plotly.express as px
import io


# Funções de Análise
def load_data(file):
    data = pd.read_excel(file)
    data.columns = data.iloc[1]
    data = data.drop([0, 1])

    data.rename(columns={"Servi¿os Bruto": "Serviços Bruto"}, inplace=True)

    data["Vlr. Itens Bruto"] = pd.to_numeric(
        data["Vlr. Itens Bruto"], errors="coerce"
    ).fillna(0)
    data["Serviços Bruto"] = pd.to_numeric(
        data["Serviços Bruto"], errors="coerce"
    ).fillna(0)
    data["Valor Total"] = data["Vlr. Itens Bruto"] + data["Serviços Bruto"]

    # Adicionar coluna 'Aging'
    data["Dias na Oficina"] = pd.to_numeric(
        data["Dias na Oficina"], errors="coerce"
    ).fillna(0)
    data["Aging"] = data["Dias na Oficina"].apply(
        lambda x: (
            "Até 30 dias"
            if x <= 30
            else (
                "De 31 a 60 dias"
                if x <= 60
                else "De 61 a 90 dias" if x <= 90 else "Acima de 121 dias"
            )
        )
    )
    return data


def analyze_summary(data):
    pivot_table = data.pivot_table(
        index=["Tipo", "Descrição"],
        values=["Vlr. Itens Bruto", "Serviços Bruto", "Valor Total", "Numero"],
        aggfunc={
            "Vlr. Itens Bruto": "sum",
            "Serviços Bruto": "sum",
            "Valor Total": "sum",
            "Numero": "count",
        },
        margins=True,
        margins_name="Total",
    ).rename(columns={"Numero": "Quantidade de OS"})

    # Adicionar coluna de variação
    total_value = pivot_table.loc["Total", "Valor Total"]
    pivot_table["Variação (%)"] = (pivot_table["Valor Total"] / total_value) * 100
    pivot_table["Variação (%)"] = pivot_table["Variação (%)"].round(2)

    return pivot_table


def analyze_aging_by_service(data):
    pivot_table = data.pivot_table(
        index=["Descrição"],
        columns="Aging",
        values="Valor Total",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Total",
    )
    os_count = data.pivot_table(
        index=["Descrição"],
        values="Numero",
        aggfunc="count",
        margins=True,
        margins_name="Total",
    ).rename(columns={"Numero": "Quantidade de OS"})
    pivot_table = pivot_table.join(os_count)

    # Adicionar coluna de variação
    total_value = pivot_table.loc["Total", "Total"]
    pivot_table["Variação (%)"] = (pivot_table["Total"] / total_value) * 100
    pivot_table["Variação (%)"] = pivot_table["Variação (%)"].round(2)

    return pivot_table


def analyze_aging(data):
    # Adicionar campo "Valor Total" (Vlr. Itens Bruto + Serviços Bruto)
    data["Valor Total"] = data["Vlr. Itens Bruto"] + data["Serviços Bruto"]

    # Criar a tabela dinâmica
    pivot_table = data.pivot_table(
        index="Aging",
        values=["Vlr. Itens Bruto", "Serviços Bruto", "Valor Total", "Numero"],
        aggfunc={
            "Vlr. Itens Bruto": "sum",
            "Serviços Bruto": "sum",
            "Valor Total": "sum",
            "Numero": "count",
        },
        margins=True,
        margins_name="Total",
    ).rename(columns={"Numero": "Quantidade de OS"})

    # Calcular a variação
    total_valor_total = pivot_table.at["Total", "Valor Total"]
    pivot_table["Variação"] = pivot_table["Valor Total"] / total_valor_total

    return pivot_table


def at1_os(data):
    return data


def analyze_synthetic_comparative_type(data_old, data_new):
    def summarize(data):
        return data.pivot_table(
            index=["Tipo", "Descrição"],
            values=["Valor Total", "Numero"],
            aggfunc={"Valor Total": "sum", "Numero": "count"},
        ).rename(columns={"Numero": "Quantidade de OS"})

    old_summary = summarize(data_old).rename(
        columns={
            "Quantidade de OS": "Quantidade de OS_Antigo",
            "Valor Total": "Valor Total_Antigo",
        }
    )
    new_summary = summarize(data_new).rename(
        columns={
            "Quantidade de OS": "Quantidade de OS_Atual",
            "Valor Total": "Valor Total_Atual",
        }
    )

    comparative = old_summary.join(new_summary, how="outer").fillna(0)
    comparative[["Quantidade de OS_Antigo", "Quantidade de OS_Atual"]] = comparative[
        ["Quantidade de OS_Antigo", "Quantidade de OS_Atual"]
    ].astype(int)

    # Adicionando os totais no final
    total_row = pd.DataFrame(
        {
            "Quantidade de OS_Antigo": [comparative["Quantidade de OS_Antigo"].sum()],
            "Valor Total_Antigo": [comparative["Valor Total_Antigo"].sum()],
            "Quantidade de OS_Atual": [comparative["Quantidade de OS_Atual"].sum()],
            "Valor Total_Atual": [comparative["Valor Total_Atual"].sum()],
        },
        index=[("Total", "")],
    )

    comparative = pd.concat([comparative, total_row])

    # Restabelecendo os índices como multi-índice
    comparative.index = pd.MultiIndex.from_tuples(
        comparative.index, names=["Tipo", "Descrição"]
    )

    return comparative


def analyze_comparative_type_time(data_old, data_new):
    def summarize_value(data):
        return data.pivot_table(
            index=["Tipo"], values="Valor Total", aggfunc="sum"
        ).rename(columns={"Valor Total": "Valor Total"})

    def summarize_count(data):
        return data.pivot_table(
            index=["Tipo"], values="Numero", aggfunc="count"
        ).rename(columns={"Numero": "Quantidade de OS"})

    old_values = summarize_value(data_old).rename(
        columns={"Valor Total": "Valor Total_Antigo"}
    )
    new_values = summarize_value(data_new).rename(
        columns={"Valor Total": "Valor Total_Atual"}
    )
    comparative_values = old_values.join(new_values, how="outer").fillna(0)

    old_counts = summarize_count(data_old).rename(
        columns={"Quantidade de OS": "Quantidade de OS_Antigo"}
    )
    new_counts = summarize_count(data_new).rename(
        columns={"Quantidade de OS": "Quantidade de OS_Atual"}
    )
    comparative_counts = old_counts.join(new_counts, how="outer").fillna(0)
    comparative_counts[["Quantidade de OS_Antigo", "Quantidade de OS_Atual"]] = (
        comparative_counts[
            ["Quantidade de OS_Antigo", "Quantidade de OS_Atual"]
        ].astype(int)
    )

    # Adicionando os totais nas tabelas
    total_values_row = pd.DataFrame(
        {
            "Valor Total_Antigo": [comparative_values["Valor Total_Antigo"].sum()],
            "Valor Total_Atual": [comparative_values["Valor Total_Atual"].sum()],
        },
        index=["Total"],
    )
    comparative_values = pd.concat([comparative_values, total_values_row])

    total_counts_row = pd.DataFrame(
        {
            "Quantidade de OS_Antigo": [
                comparative_counts["Quantidade de OS_Antigo"].sum()
            ],
            "Quantidade de OS_Atual": [
                comparative_counts["Quantidade de OS_Atual"].sum()
            ],
        },
        index=["Total"],
    )
    comparative_counts = pd.concat([comparative_counts, total_counts_row])

    return comparative_values, comparative_counts


def analyze_synthetic_comparative_age(data_old, data_new):
    def summarize(data):
        return data.pivot_table(
            index="Aging",
            values=["Valor Total", "Numero"],
            aggfunc={"Valor Total": "sum", "Numero": "count"},
        ).rename(columns={"Numero": "Quantidade de OS"})

    old_summary = summarize(data_old).rename(
        columns={
            "Quantidade de OS": "Quantidade de OS_Antigo",
            "Valor Total": "Valor Total_Antigo",
        }
    )
    new_summary = summarize(data_new).rename(
        columns={
            "Quantidade de OS": "Quantidade de OS_Atual",
            "Valor Total": "Valor Total_Atual",
        }
    )

    comparative = old_summary.join(new_summary, how="outer").fillna(0)
    comparative[["Quantidade de OS_Antigo", "Quantidade de OS_Atual"]] = comparative[
        ["Quantidade de OS_Antigo", "Quantidade de OS_Atual"]
    ].astype(int)

    # Adicionando a linha total
    total_row = pd.DataFrame(
        {
            "Quantidade de OS_Antigo": [comparative["Quantidade de OS_Antigo"].sum()],
            "Valor Total_Antigo": [comparative["Valor Total_Antigo"].sum()],
            "Quantidade de OS_Atual": [comparative["Quantidade de OS_Atual"].sum()],
            "Valor Total_Atual": [comparative["Valor Total_Atual"].sum()],
        },
        index=["Total"],
    )

    comparative = pd.concat([comparative, total_row])

    # Restabelecendo os índices
    comparative.index.name = "Aging"

    return comparative


def analyze_comparative_age(data_old, data_new):
    def summarize_value(data):
        return data.pivot_table(
            index="Aging", values="Valor Total", aggfunc="sum"
        ).rename(columns={"Valor Total": "Valor Total"})

    def summarize_count(data):
        return data.pivot_table(index="Aging", values="Numero", aggfunc="count").rename(
            columns={"Numero": "Quantidade de OS"}
        )

    old_values = summarize_value(data_old).rename(
        columns={"Valor Total": "Valor Total_Antigo"}
    )
    new_values = summarize_value(data_new).rename(
        columns={"Valor Total": "Valor Total_Atual"}
    )
    comparative_values = old_values.join(new_values, how="outer").fillna(0)

    old_counts = summarize_count(data_old).rename(
        columns={"Quantidade de OS": "Quantidade de OS_Antigo"}
    )
    new_counts = summarize_count(data_new).rename(
        columns={"Quantidade de OS": "Quantidade de OS_Atual"}
    )
    comparative_counts = old_counts.join(new_counts, how="outer").fillna(0)
    comparative_counts[["Quantidade de OS_Antigo", "Quantidade de OS_Atual"]] = (
        comparative_counts[
            ["Quantidade de OS_Antigo", "Quantidade de OS_Atual"]
        ].astype(int)
    )

    # Adicionando a linha total em comparative_values
    total_values_row = pd.DataFrame(
        {
            "Valor Total_Antigo": [comparative_values["Valor Total_Antigo"].sum()],
            "Valor Total_Atual": [comparative_values["Valor Total_Atual"].sum()],
        },
        index=["Total"],
    )
    comparative_values = pd.concat([comparative_values, total_values_row])
    comparative_values.index.name = "Aging"

    # Adicionando a linha total em comparative_counts
    total_counts_row = pd.DataFrame(
        {
            "Quantidade de OS_Antigo": [
                comparative_counts["Quantidade de OS_Antigo"].sum()
            ],
            "Quantidade de OS_Atual": [
                comparative_counts["Quantidade de OS_Atual"].sum()
            ],
        },
        index=["Total"],
    )
    comparative_counts = pd.concat([comparative_counts, total_counts_row])
    comparative_counts.index.name = "Aging"

    return comparative_values, comparative_counts


# Função para compilar todas as análises
def analyze_all(data_current, data_previous):
    at1_os_result = at1_os(
        data_current
    )  # Certifique-se de que at1_os é uma função definida
    summary_data = analyze_summary(data_current)
    aging_by_service = analyze_aging_by_service(data_current)
    aging_data = analyze_aging(data_current)
    synthetic_comparative_type = analyze_synthetic_comparative_type(
        data_previous, data_current
    )
    comparative_type_time_values, comparative_type_time_counts = (
        analyze_comparative_type_time(data_previous, data_current)
    )
    synthetic_comparative_age = analyze_synthetic_comparative_age(
        data_previous, data_current
    )
    comparative_age_values, comparative_age_counts = analyze_comparative_age(
        data_previous, data_current
    )

    return {
        "AT1 - OSs em aberto": at1_os_result,
        "Análise por Tipo": summary_data,
        "Aging por Atendimento": aging_by_service,
        "Aging": aging_data,
        "Análise Sint Comp Tipo": synthetic_comparative_type,
        "Análises Comp Tipo Tempo - V": comparative_type_time_values,
        "Análises Comp Tipo Tempo - Q": comparative_type_time_counts,
        "Análises Sint Comp AGE": synthetic_comparative_age,
        "Análises Comp AGE - V": comparative_age_values,
        "Análises Comp AGE - Q": comparative_age_counts,
    }


# Função para gerar um Excel com todas as análises e incluir o gráfico dinâmico
def generate_excel(data_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, data in data_dict.items():
            data.to_excel(writer, sheet_name=sheet_name)

        # Criar gráfico dinâmico na planilha "Aging por Atendimento"
        workbook = writer.book
        worksheet = writer.sheets["Aging por Atendimento"]

        # Configurar o gráfico de pizza
        chart = workbook.add_chart({"type": "pie"})

        # Adicionar séries ao gráfico
        chart.add_series(
            {
                "name": "Distribuição por Descrição (Aging por Atendimento)",
                "categories": [
                    "Aging por Atendimento",
                    1,
                    0,
                    len(data_dict["Aging por Atendimento"]) - 1,
                    0,
                ],
                "values": [
                    "Aging por Atendimento",
                    1,
                    len(data_dict["Aging por Atendimento"].columns) - 2,
                    len(data_dict["Aging por Atendimento"]) - 1,
                    len(data_dict["Aging por Atendimento"].columns) - 2,
                ],
            }
        )

        # Adicionar título e estilo ao gráfico
        chart.set_title({"name": "Distribuição por Descrição (Aging por Atendimento)"})
        chart.set_style(10)

        # Inserir o gráfico na planilha
        worksheet.insert_chart("G2", chart)

    output.seek(0)
    return output


# Função principal com interface Streamlit
def main():
    st.title("Análise de OS Abertas")
    file_current = st.file_uploader(
        "Carregar arquivo atual", type=["xlsx"], key="current"
    )
    file_previous = st.file_uploader(
        "Carregar arquivo anterior", type=["xlsx"], key="previous"
    )

    analysis_options = {
        "AT1 - OSs em aberto": at1_os,
        "Análise por Tipo": analyze_summary,
        "Aging por Atendimento": analyze_aging_by_service,
        "Aging": analyze_aging,
        "Análise Sintéticas Comparativas Tipo": analyze_synthetic_comparative_type,
        "Análises Comparativas Tipo Tempo": analyze_comparative_type_time,
        "Análises Sintéticas Comparativas AGE": analyze_synthetic_comparative_age,
        "Análises Comparativas AGE": analyze_comparative_age,
    }

    if file_current and file_previous:
        data_current = load_data(file_current)
        data_previous = load_data(file_previous)

        choice = st.selectbox("Escolha a análise:", list(analysis_options.keys()))

        if (
            choice == "Análises Comparativas Tipo Tempo"
            or choice == "Análises Comparativas AGE"
        ):
            comparative_values, comparative_counts = analysis_options[choice](
                data_previous, data_current
            )
            st.write(f"Resultado Comparativo de Valores ({choice}):")
            st.dataframe(comparative_values)

            st.write(f"Resultado Comparativo de Quantidade de OS ({choice}):")
            st.dataframe(comparative_counts)
        elif (
            choice == "Análise Sintéticas Comparativas Tipo"
            or choice == "Análises Sintéticas Comparativas AGE"
        ):
            result = analysis_options[choice](data_previous, data_current)
            st.write(f"Resultado da {choice}:")
            st.dataframe(result)
        elif choice == "Aging por Atendimento":
            pivot_table = analysis_options[choice](data_current)
            st.write(f"Resultado da {choice}:")
            st.dataframe(pivot_table)
        else:
            result = analysis_options[choice](data_current)
            st.write(f"Resultado da {choice}:")
            st.dataframe(result)

        # Adicionar botão de download para todas as análises
        all_analyses = analyze_all(data_current, data_previous)

        # Verificar a saída
        print(type(all_analyses))  # Deve ser <class 'dict'>

        # Verifique se all_analyses é um dicionário
        if isinstance(all_analyses, dict):
            excel_data = generate_excel(all_analyses)
            st.download_button(
                label="Download de Todas as Análises em Excel",
                data=excel_data,
                file_name="Todas_as_Análises.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("Erro: a saída das análises não é um dicionário.")


if __name__ == "__main__":
    main()
