import pandas as pd
import numpy as np
import streamlit as st
from PIL import Image
import os
import warnings
warnings.filterwarnings("ignore")
import html
import xlsxwriter
# from mimetypes import guess_extension


st.set_option('deprecation.showPyplotGlobalUse', False)

st.set_page_config(layout="wide")

# front end elements of the web page

html_temp = """
    <header style="font-size:10px;width=60">
    <div style ="background-color:lightgreen;border-style:solid">
    <img src="https://renom.in/wp-content/uploads/2022/02/cropped-renom-logo.86b197ce-e1644472068111-1.png" style="float:Right;width:200px;height:70px;border:orange; border-width:10px; border-style:solid;">
    <h2 style ="color:black;text-align:center;">RESCA Temperature Analysis</h1>
    <h4 style ="color:black;text-align:center;"> Renom Energy Services Pvt Ltd </h3>
    </header>
    <body>
    <p></p>
    <h6 style ="color:black;text-align:center;"> Application to check Abnormal Temperature Observations in various Temp. KPI's of Turbine </h6>
    """

st.markdown(html_temp, unsafe_allow_html=True)



st.markdown(
         f"""
         <style>
         .stApp {{
             background-image: url("http://cdn.walkthroughindia.com/wp-content/uploads/2016/02/Muppandal-Windfarm-Kanyakumari-647x400.jpg");
             background-attachment: fixed;
             background-size: cover
         }}
         </style>
         """,
         unsafe_allow_html=True
     )



#https://static.stacker.com/s3fs-public/styles/sar_screen_maximum_large/s3/croppedshutterstock1934730587IKL8jpg_0.webp

uploaded_file = st.sidebar.file_uploader("Upload file", type=['xlsx'])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    df0=pd.read_excel(uploaded_file,0)
    df1=pd.read_excel(uploaded_file,1)
    df2=pd.read_excel(uploaded_file,2)
    df3=pd.read_excel(uploaded_file,3)
    df4=pd.read_excel(uploaded_file,4)
    df5=pd.read_excel(uploaded_file,5)
    df6=pd.read_excel(uploaded_file,6)
    df7=pd.read_excel(uploaded_file,7)
    df8=pd.read_excel(uploaded_file,8)
    df9=pd.read_excel(uploaded_file,9)
    dfa=pd.read_excel(uploaded_file,10)

    df0 = df0.replace('-',np.nan)
    df1 = df1.replace('-',np.nan)
    df2 = df2.replace('-',np.nan)
    df3 = df3.replace('-',np.nan)
    df4 = df4.replace('-',np.nan)
    df5 = df5.replace('-',np.nan)
    df6 = df6.replace('-',np.nan)
    df7 = df7.replace('-',np.nan)
    df8 = df8.replace('-',np.nan)
    df9 = df9.replace('-',np.nan)
    dfa = dfa.replace('-',np.nan)

    E30_pivot=df0.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=[ 'Excitation Heat SinkTemperature','Front Bearing Temperature', 'Rear Bearing Temperature',
        'Motor Blade A Temperature', 'Motor Blade B Temperature','Motor Blade C Temperature', 'Rotor Temperature', 'Stator Temperature','Controller Cabinet Temperature',
        'Nacelle Control cabinet Temperature', 'Converter Cabinet Temperature','Converter Heat Sink 1 Temperature','Converter Heat Sink 2 Temperature',
        'Converter Heat Sink 3 Temperature'],aggfunc='mean')
    E30_pivot = E30_pivot.round(2)

    E33_pivot=df1.pivot_table(index=['Farm Name','Turbine Name','Model'],values=['Nacelle Temperature','Front Bearing Temperature', 'Rear Bearing Temperature',
        'Motor Blade A Temperature', 'Motor Blade B Temperature',
        'Motor Blade C Temperature', 'Rotor Temperature', 'Stator Temperature',
        'Transformer Temperature', 'Converter Cabinet Temperature', 'Converter Heat Sink 1 Temperature',
        'Converter Heat Sink 2 Temperature',
        'Converter Heat Sink 3 Temperature',
        'Converter Heat Sink 4 Temperature'],aggfunc='mean')
    E33_pivot = E33_pivot.round(2)

    E40_pivot=df2.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=['Excitation Heat SinkTemperature','Heat Sink Rectifier 1 Temperature', 'Heat Sink Rectifier 2 Temperature', 'Motor A Temperature',
        'Motor B Temperature', 'Motor C Temperature', 'Rotor Temperature','Stator Temperature', 'Controller Cabinet Temperature', 'Nacelle Control cabinet Temperature',
        'Converter 1 Heat Sink ChopTemperature','Converter 2 Heat Sink ChopTemperature', 'Yaw Break Temperature'],aggfunc='mean')
    E40_pivot = E40_pivot.round(2)

    E48_pivot=df3.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=['Converter 1 Step Up Chop Temperature',
        'Converter 2 Step Up Chop Temperature',
        'Converter 3 Step Up Chop Temperature',
        'Excitation Heat SinkTemperature', 'Front Bearing Temperature',
        'Rear Bearing Temperature', 'Heat Sink Rectifier 1 Temperature',
        'Heat Sink Rectifier 2 Temperature', 'Motor A Temperature',
        'Motor B Temperature', 'Motor C Temperature', 'Rotor 1 Temperature',
        'Rotor 2 Temperature', 'Stator 1 Temperature', 'Stator 2 Temperature', 'Yaw Inverter Heat Sink 1 Temperature',
        'Yaw Inverter Heat Sink 2 Temperature'],aggfunc='mean')
    E48_pivot = E48_pivot.round(2)

    E53_pivot=df4.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=['Converter 1 Step Up Chop Temperature',
        'Converter 2 Step Up Chop Temperature',
        'Converter 3 Step Up Chop Temperature',
        'Excitation Heat SinkTemperature', 'Front Bearing Temperature',
        'Rear Bearing Temperature', 'Heat Sink Rectifier 1 Temperature',
        'Heat Sink Rectifier 2 Temperature', 'Motor A Temperature',
        'Motor B Temperature', 'Motor C Temperature', 'Rotor 1 Temperature',
        'Rotor 2 Temperature', 'Stator 1 Temperature', 'Stator 2 Temperature','Yaw Inverter Heat Sink 1 Temperature',
        'Yaw Inverter Heat Sink 2 Temperature'],aggfunc='mean')
    E53_pivot = E53_pivot.round(2)

    NM48_pivot=df5.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=['Thyristor Temperature','Nacelle Temperature', 'Gearbearing Temperature','Generator 1 Temperature',
        'Generator 2 Temperature'],aggfunc='mean')
    NM48_pivot = NM48_pivot.round(2)

    V39_pivot=df6.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=['Controller Temperature', 'GearBox Temperature',
        'Generator Temperature', 'Nacelle Temperature', 'Hydraulic Temperature','Main Bearing Temperature'],aggfunc='mean')
    V39_pivot = V39_pivot.round(2)
    

    V47_pivot=df7.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=[ 'Controller Temperature', 'GearBox Temperature',
        'Generator Temperature', 'Nacelle Temperature', 'Hydraulic Temperature',
        'Main Bearing Temperature'],aggfunc='mean')
    V47_pivot = V47_pivot.round(2)

    V77_pivot=df8.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=['AC Inductor Temperature',
        'DC Inductor Temperature', 'DC Link Capacitor Temperature',
        'Generator Capacitor Temperature', 'Chopper IGBT Temperature',
        'Grid Up IGBT L1A Temperature', 'Grid Up IGBT L1B Temperature',
        'Grid Up IGBT L2A Temperature', 'Grid Up IGBT L2B Temperature',
        'Grid Up IGBT L3A Temperature', 'Grid Up IGBT L3B Temperature',
        'Pitch Cabinate sys1 Temperature', 'Pitch Cabinate sys2 Temperature',
        'Pitch Cabinate sys3 Temperature', 'Pitch Converter sys1 Temperature',
        'Pitch Converter sys2 Temperature', 'Pitch Converter sys3 Temperature',
        'Pitch Motor sys1 Temperature', 'Pitch Motor sys2 Temperature',
        'Pitch Motor sys3 Temperature', 'Pitch Capacitor sys1 Temperature',
        'Pitch Capacitor sys2 Temperature', 'Pitch Capacitor sys3 Temperature',
        'Step Up IGBT 1 Temperature', 'Step Up IGBT 2 Temperature',
        'Step Up IGBT 3 Temperature', 'Rectifier Temperature'],aggfunc='mean')
    V77_pivot = V77_pivot.round(2)

    V82_pivot=df9.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=['Gear Oil Temperature', 'Generator Temperature',
        'Nacelle Temperature', 'Gear Bearing Front Temperature','Gear Bearing Rear Temperature',
        'Thyristor Temperature'],aggfunc='mean')
    V82_pivot = V82_pivot.round(2)
    
    V87_pivot=dfa.pivot_table(index=['Farm Name', 'Turbine Name', 'Model'],values=['AC Inductor Temperature',
        'DC Inductor Temperature', 'DC Link Capacitor Temperature',
        'Generator Capacitor Temperature', 'Chopper IGBT Temperature',
        'Grid Up IGBT L1A Temperature', 'Grid Up IGBT L1B Temperature',
        'Grid Up IGBT L2A Temperature', 'Grid Up IGBT L2B Temperature',
        'Grid Up IGBT L3A Temperature', 'Grid Up IGBT L3B Temperature',
        'Pitch Cabinate sys1 Temperature', 'Pitch Cabinate sys2 Temperature',
        'Pitch Cabinate sys3 Temperature', 'Pitch Converter sys1 Temperature',
        'Pitch Converter sys2 Temperature', 'Pitch Converter sys3 Temperature',
        'Pitch Motor sys1 Temperature', 'Pitch Motor sys2 Temperature',
        'Pitch Motor sys3 Temperature', 'Pitch Capacitor sys1 Temperature',
        'Pitch Capacitor sys2 Temperature', 'Pitch Capacitor sys3 Temperature',
        'Step Up IGBT 1 Temperature', 'Step Up IGBT 2 Temperature',
        'Step Up IGBT 3 Temperature', 'Rectifier Temperature'],aggfunc='mean')
    V87_pivot = V87_pivot.round(2)
    

    def highlight_cell(cell):
        if type(cell) != str and cell < 0 :
            return 'background-color: skyblue'
        elif type(cell) != str and cell > 80:
            return 'background-color : red'
        elif cell > 70 and cell < 80 :
            return 'background-color : orange'
        elif cell > 60 and cell < 70 :
            return 'background-color : yellow'
        elif cell == 0 :
            return 'background-color : pink'
        elif cell == None :
            return 'background-color: green'
        
        
    E30_pivot = E30_pivot.style.applymap(highlight_cell)
    E33_pivot = E33_pivot.style.applymap(highlight_cell)
    E40_pivot = E40_pivot.style.applymap(highlight_cell)
    E48_pivot = E48_pivot.style.applymap(highlight_cell)
    E53_pivot = E53_pivot.style.applymap(highlight_cell)
    NM48_pivot = NM48_pivot.style.applymap(highlight_cell)

    V39_pivot = V39_pivot.style.applymap(highlight_cell)
    V47_pivot = V47_pivot.style.applymap(highlight_cell)
    V77_pivot = V77_pivot.style.applymap(highlight_cell)
    V82_pivot = V82_pivot.style.applymap(highlight_cell)
    V87_pivot = V87_pivot.style.applymap(highlight_cell)



    dflist=[E30_pivot,E33_pivot,E40_pivot,E48_pivot,E53_pivot,NM48_pivot,V39_pivot,V47_pivot,V77_pivot,V82_pivot,V87_pivot]



    Excelwriter = pd.ExcelWriter("weekly temp report.xlsx",engine="xlsxwriter")

    for i, df in enumerate(dflist) :
        df.to_excel(Excelwriter, sheet_name="Sheet" + str(i+1),index=True)
    #And finally save the file
    Excelwriter.save()
    
     
    df111 = pd.read_excel("weekly temp report.xlsx")
    st.write(df111)

df = df111
if df is not None:
    file_name = st.text_input('Weekly Temp Report')
    download = st.download_button(label='Download Excel', data=df.to_excel(index=False, header=True), key='download')

if file_name and download:
    with open(file_name, "wb") as f:
        f.write(download)
    st.success(f"File '{file_name}' has been downloaded successfully.")

