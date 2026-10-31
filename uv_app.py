from io import BytesIO

import matplotlib.pyplot as plt
import streamlit as st
import xlsxwriter


st.title("UVスペクトル重ね描き")
st.write("UVデータのTXTファイルを選んでください。")

files = st.file_uploader(
    "データファイルを選択",
    type=["txt"],
    accept_multiple_files=True,
)

if files:
    fig, ax = plt.subplots()
    spectra = []

    for file in files:
        raw_data = file.getvalue()

        try:
            text = raw_data.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw_data.decode("cp932", errors="replace")

        lines = text.splitlines()
        reading_data = False
        wavelength = []
        absorbance = []

        for line in lines:
            if line.strip().upper() == "XYDATA":
                reading_data = True
                continue

            if not reading_data:
                continue

            columns = line.split()
            if len(columns) < 2:
                continue

            try:
                wavelength.append(float(columns[0]))
                absorbance.append(float(columns[1]))
            except ValueError:
                continue

        if wavelength:
            ax.plot(wavelength, absorbance, label=file.name)
            spectra.append(
                {
                    "name": file.name,
                    "wavelength": wavelength,
                    "absorbance": absorbance,
                }
            )
        else:
            st.warning(f"{file.name}: XYDATAの数値を読み取れませんでした。")

    if spectra:
        ax.set_xlabel("Wavelength (nm)")
        ax.set_ylabel("Absorbance")
        ax.set_title("UV-Vis spectra")
        ax.grid(True, alpha=0.3)
        ax.legend()
        st.pyplot(fig)

        # Excelファイルをメモリ上に作る
        excel_buffer = BytesIO()
        workbook = xlsxwriter.Workbook(
            excel_buffer,
            {"in_memory": True},
        )

        data_sheet = workbook.add_worksheet("Spectra")
        chart_sheet = workbook.add_worksheet("Chart")
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})

        chart = workbook.add_chart(
            {
                "type": "scatter",
                "subtype": "straight",
            }
        )

        for spectrum_number, spectrum in enumerate(spectra):
            column = spectrum_number * 2

            data_sheet.write(
                0,
                column,
                f"{spectrum['name']} - Wavelength (nm)",
                header_format,
            )
            data_sheet.write(
                0,
                column + 1,
                f"{spectrum['name']} - Absorbance",
                header_format,
            )

            for row, (x_value, y_value) in enumerate(
                zip(spectrum["wavelength"], spectrum["absorbance"]),
                start=1,
            ):
                data_sheet.write_number(row, column, x_value)
                data_sheet.write_number(row, column + 1, y_value)

            last_row = len(spectrum["wavelength"])

            chart.add_series(
                {
                    "name": spectrum["name"],
                    "categories": [
                        "Spectra",
                        1,
                        column,
                        last_row,
                        column,
                    ],
                    "values": [
                        "Spectra",
                        1,
                        column + 1,
                        last_row,
                        column + 1,
                    ],
                }
            )

            data_sheet.set_column(column, column + 1, 24)

        chart.set_title({"name": "UV-Vis spectra"})
        chart.set_x_axis({"name": "Wavelength (nm)"})
        chart.set_y_axis({"name": "Absorbance"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 900, "height": 540})

        chart_sheet.insert_chart("A1", chart)
        data_sheet.freeze_panes(1, 0)

        workbook.close()
        excel_file = excel_buffer.getvalue()

        st.download_button(
            label="Excelファイルをダウンロード",
            data=excel_file,
            file_name="uv_spectra.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("グラフにできるデータがありませんでした。")