import streamlit as st
import pandas as pd
import io
import xlwt

####################
# Cabeçalho
header_style = xlwt.XFStyle()
header_font = xlwt.Font()
header_font.name = 'Calibri'
header_font.bold = True
header_font.height = 220  # 11 pt
header_style.font = header_font

# Dados
data_style = xlwt.XFStyle()
data_font = xlwt.Font()
data_font.name = 'Calibri'
data_font.height = 220  # 10 pt
data_style.font = data_font

####################



st.markdown(
    "<h1 style='text-align: center;'>Conversor de xlsx para xls formato MV</h1>",
    unsafe_allow_html=True
)


uploaded_file = st.file_uploader("Upload an image", type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df = df.rename(columns={'PROCEDIMENTO':'CD', 'VR_TOTAL': 'VALOR'})
    df['FIM_ARQUIVO'] = '9'
    df['VL_UCO'] = ''
    df['TP_FILME'] = ''
    df['NR_AUXILIAR'] = ''
    df['CD_PORTE_ANESTESICO'] = ''
    df['CD_PORTE_MEDICO'] = ''
    df['VL_PERC_BANDA_UCO'] = ''
    df['VL_PERC_PESO_PORTE'] = ''
    df['NR_INCIDENCIAS'] = ''
    df['DT_VIGENCIA'] = ''
    df = df[['CD', 'VALOR', 'VL_UCO', 'TP_FILME', 'NR_AUXILIAR', 'CD_PORTE_ANESTESICO', 'CD_PORTE_MEDICO', 'VL_PERC_BANDA_UCO', 'VL_PERC_PESO_PORTE',
             'NR_INCIDENCIAS', 'DT_VIGENCIA', 'FIM_ARQUIVO']]

    st.dataframe(df)

    output = io.BytesIO()
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Dados")

     # Cabeçalho
    for col, name in enumerate(df.columns):
        sheet.write(0, col, str(name), header_style)

    # Dados
    for row in range(len(df)):
        for col in range(len(df.columns)):
            value = df.iat[row, col]

            if pd.isna(value):
                sheet.write(row + 1, col, "", data_style)
            elif hasattr(value, "item"):
                sheet.write(row + 1, col, value.item(), data_style)
            else:
                sheet.write(row + 1, col, value, data_style)

    workbook.save(output)
    output.seek(0)

    st.download_button(
        label="Baixar arquivo XLS",
        data=output,
        file_name="arquivo_mv.xls",
        mime="application/vnd.ms-excel"
    )