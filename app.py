import streamlit as st
import pandas as pd

def распределить_работы(raboty, sostavy):
    raboty_sorted = raboty.sort_values(by="Стоимость", ascending=False).reset_index(drop=True)
    sostavy_sorted = sostavy.sort_values(by="Стоимость", ascending=False).reset_index(drop=True).copy()
    sostavy_sorted["Остаток"] = sostavy_sorted["Стоимость"]

    rows = []
    for _, rabota in raboty_sorted.iterrows():
        remaining = rabota["Стоимость"]
        nds_ratio = rabota["НДС"] / rabota["Стоимость"] if rabota["Стоимость"] != 0 else 0

        for i, sostav in sostavy_sorted.iterrows():
            if remaining <= 0:
                break
            available = sostav["Остаток"]
            if available <= 0:
                continue
            take = min(available, remaining)
            nds_part = take * nds_ratio
            rows.append({
                "Состав": sostav["Состав"],
                "Наименование работ": rabota["Наименование работ"],
                "Сумма": take,
                "НДС": nds_part,
                "ОригинальныйID": rabota["Номер"],
                "Примечание": "Частично" if take < remaining else "Полностью"
            })
            sostavy_sorted.at[i, "Остаток"] -= take
            remaining -= take

    return pd.DataFrame(rows)

st.title("Распределение работ по составам")

uploaded_file = st.file_uploader("Загрузите Excel-файл с листами 'Работы' и 'Составы'", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    if "Работы" not in xls.sheet_names or "Составы" not in xls.sheet_names:
        st.error("В файле должны быть листы с именами 'Работы' и 'Составы'")
    else:
        raboty = pd.read_excel(xls, sheet_name="Работы")
        sostavy = pd.read_excel(xls, sheet_name="Составы")

        # Запускаем распределение
        result = распределить_работы(raboty, sostavy)

        st.write("Результат распределения:")
        st.dataframe(result)

        # Генерируем Excel для скачивания
        import io
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
