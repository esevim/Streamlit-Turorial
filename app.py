import pandas as pd
import streamlit as st
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import plotly.express as px
from PIL import Image

st.set_page_config(page_title='Survey Results')
st.header('Survey results 2021')
st.subheader('Was the tutorial helpful?')


### --- Load DataFrame
excel_file = 'Survey_Results.xlsx'
sheet_name = 'DATA'

df = pd.read_excel(excel_file, 
                   sheet_name=sheet_name,
                   usecols= 'B:D',
                   header=3)

df_participants = pd.read_excel(excel_file,
                                sheet_name=sheet_name,
                                usecols='F:G',
                                header=3)


### -- Upload a file within app.
# uploaded_file = st.file_uploader('Choose a file')
# if uploaded_file is not None:
#     #read excel
#     df = pd.read_excel(uploaded_file, 
#                    sheet_name=sheet_name,
#                    usecols= 'B:D',
#                    header=3)

#     df_participants = pd.read_excel(uploaded_file,
#                                 sheet_name=sheet_name,
#                                 usecols='F:G',
#                                 header=3)
# else:
#     st.warning('you need to upload a csv or excel file.')


df_participants.dropna(inplace=True)

# --- Streamlit Selection
department = df['Department'].unique().tolist()
ages = df['Age'].unique().tolist()

age_selection = st.slider('Age',
                        min_value= min(ages),
                        max_value= max(ages),
                        value=(min(ages),max(ages)))

department_selection = st.multiselect('Department:',
                                        department,
                                        default=department)

# --- Filter Dataframe based on Selection

mask = (df['Age'].between(*age_selection)) & (df['Department'].isin(department_selection))
number_of_results = df[mask].shape[0]
st.markdown(f'*Available Results: {number_of_results}*')


# --- GROUP DATAFRAME AFTER SELECTION
df_grouped = df[mask].groupby(by=['Rating']).count()[['Age']]
df_grouped = df_grouped.rename(columns={'Age': 'Votes'})
df_grouped = df_grouped.reset_index()

# --- PLOT BAR CHART
bar_chart = px.bar(df_grouped,
                   x='Rating',
                   y='Votes',
                   text='Votes',
                   color_discrete_sequence = ['#F63366']*len(df_grouped),
                   template= 'plotly_white')
st.plotly_chart(bar_chart)


st.dataframe(df)
# st.dataframe(df_participants)

pie_chart = px.pie(df_participants,
                    title = 'Total No. of Participants',
                    values = 'Participants',
                    names = 'Departments')

st.plotly_chart(pie_chart)

# image= Image.open('Image/example image.jpg')
# st.image(image,
#         caption = 'test image',
#         use_column_width = True)


### --- Download Button ---
# df.to_excel('file.xlsx')

# st.download_button(
#    label = 'Press to Download',
#    data = df_xlsx,
#    file_name = "file.xlsx",
#     mime = "text/excel",
#    key='download-csv'
# )

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    # format1 = workbook.add_format({'num_format': '0.00'}) 
    # worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

df_xlsx = to_excel(df_participants)

st.download_button(label='📥 Download Current Result',
                                data=df_xlsx,
                                file_name= 'df_test.xlsx')