import streamlit as st
import pandas as pd
import itertools
import io

def find_subset_sum(numbers, target, tol=0.01):
    n = len(numbers)
    for r in range(1, n + 1):
        for subset in itertools.combinations(range(n), r):
            s = sum(numbers[i] for i in subset)
            if abs(s - target) <= tol:
                return subset
    return None

def распределить_работы(raboty, sostavy):
    raboty_sorted = raboty.sort_values(by="Стоимость", ascending=False).reset_index(drop=True)
    sostavy_sorted = sostavy.sort_values(by="Стоимость", ascending=False).reset_index(drop=True).copy()
    sostavy_sorted["Остаток"] = sostavy_sorted["Стоимость"]

    result_rows = []

    def add_result(rabota, sostav_name, amount, note):
        nds_ratio = rabota["НДС"] / rabota["Стоимость"] if rabota["Стоимость"] != 0 else 0
        nds_val = amount * nds_ratio
        result_rows.append({
            "Состав": sostav_name,
            "Наименование работ": rabota["Наименование работ"],
            "Сумма": amount,
            "НДС": nds_val,
            "ОригинальныйID": rabota["Номер"],
            "Примечание": note
        })

    works = raboty_sorted.copy()

    while len(works) > 0:
        current = works.iloc[0]
        remaining = current["Стоимость"]
        nds_total = current["НДС"]

        # 2.1 Проверяем целиком
        placed = False
        for i, row in sostavy_sorted.iterrows():
            if row["Остаток"] >= remaining:
                add_result(current, row["Состав"], remaining, "Полностью")
                sostavy_sorted.at[i, "Остаток"] -= remaining
                placed = True
                break

        if placed:
            works = works.iloc[1:].reset_index(drop=True)
            continue

        # 2.2 Ищем точное подмножество остатков
        остатки = sostavy_sorted["Остаток"].tolist()
        subset_idx = find_subset_sum(остатки, remaining)
        if subset_idx is not None:
            for idx in subset_idx:
                portion = sostavy_sorted.at[idx, "Остаток"]
                add_result(current, sostavy_sorted.at[idx, "Состав"], portion, "Частично")
                sostavy_sorted.at[idx, "Остаток"] = 0
            works = works.iloc[1:].reset_index(drop=True)
            continue

        # 2.3 Частичное распределение — самый большой остаток
        max_idx = sostavy_sorted["Остаток"].idxmax()
        max_ost = sostavy_sorted.at[max_idx, "Остаток"]
        add_result(current, sostavy_sorted.at[max_idx, "Состав"], max_ost, "Частично (по максимуму)")
        sostavy_sorted.at[max_idx, "Остаток"] = 0

        # Обновляем работы — заменяем текущую работу остатком
        remaining -= max_ost
        nds_ratio = nds_total / current["Стоимость"] if current["Стоимость"] != 0 else 0
        nds_new = nds_ratio * remaining

        works.at[0, "Стоимость"] = remaining
        works.at[0, "НДС"] = nds_new
        works = works.sort_values(by="Стоимость", ascending=False).reset_index(drop=True)

    return pd.DataFrame(result_rows)

st.title("Распределение работ по составам")

uploaded_file = st.file_uploader("Загрузите Excel-файл с листами 'Работы' и 'Составы'", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        if "Работы" not in xls.sheet_names or "Составы" not in xls.sheet_names:
            st.error("В файле должны быть листы с именами 'Работы' и 'Составы'")
        else:
            raboty = pd.read_excel(xls, sheet_name="Работы")
            sostavy = pd.read_excel(xls, sheet_name="Составы")

            result = распределить_работы(raboty, sostavy)

            st.write("Результат распределения:")
            st.dataframe(result)

            # Итоговые суммы для удобной проверки
            total_sum = result["Сумма"].sum()
            total_nds = result["НДС"].sum()
            st.markdown(f"**Итоговая сумма:** {total_sum:,.2f}")
            st.markdown(f"**Итоговый НДС:** {total_nds:,.2f}")

            # Кнопка для скачивания результата
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                raboty.to_excel(writer, sheet_name="Работы", index=False)
                sostavy.to_excel(writer, sheet_name="Составы", index=False)
                result.to_excel(writer, sheet_name="Распределение", index=False)
            output.seek(0)

            st.download_button(
                label="Скачать результат в Excel",
                data=output,
                file_name="Распределение_работ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
