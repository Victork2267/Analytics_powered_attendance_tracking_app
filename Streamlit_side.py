import streamlit as st
import pandas as pd
import os
import openpyxl
from math import pi
import plotly.express as px
#Open attendance xl file

def createbar(df,X, y, title):
    fig = px.bar(df, x=X, y=y)
    fig.update_layout(title_text=title)
    st.plotly_chart(fig, use_container_width=True)

def createdistplot(df, X, title):
    fig = px.histogram(df, x=X, marginal="box", hover_data=df.columns)
    fig.update_layout(title_text=title)
    st.plotly_chart(fig, use_container_width=True)

def createpie(df,X, y, title):
    fig =px.pie(df,values=y,names=X)
    fig.update_layout(title_text=title)
    st.plotly_chart(fig, use_container_width=True)

ospath = os.path.dirname(os.path.abspath(__file__))
print(ospath)

streamlit_as = pd.read_excel(open(r"{}\Revised_attendance_sheet.xlsm".format(ospath),'rb'),sheet_name="Attendance_Sheet")
streamlit_ms = pd.read_excel(open(r"{}\Revised_attendance_sheet.xlsm".format(ospath),'rb'),sheet_name="streamlit_reference_sheet")

book = openpyxl.load_workbook(r"{}\Revised_attendance_sheet.xlsm".format(ospath))
Input_label = book["Master_Control"]["B6"].value
Input_value = book["Master_Control"]["B8"].value
#print(Input_label)

st.header("Attendance tracking App")
st.subheader("Attendance Breakdown")
st.write(streamlit_as)

streamlit_as_YN_y = streamlit_as[Input_label].value_counts().tolist()
streamlit_as_YN_x = streamlit_as[Input_label].value_counts().keys().tolist()

streamlit_as_YN_df = pd.DataFrame({"Input_Label":streamlit_as_YN_x,"Proportion":streamlit_as_YN_y})

createpie(df=streamlit_as_YN_df,X="Input_Label", y="Proportion", title="")

field_1 = streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 1", "Cols"].iloc[0]
field_2 = streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 2", "Cols"].iloc[0]
field_3 = streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 3", "Cols"].iloc[0]
field_4 = streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 4", "Cols"].iloc[0]
field_5 = streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 5", "Cols"].iloc[0]

streamlit_as_Y = streamlit_as[streamlit_as[Input_label] == Input_value]

Special_field_col = book["streamlit_reference_sheet"]["B9"].value
Special_field_value = book["streamlit_reference_sheet"]["B10"].value

Special_cond = streamlit_as[streamlit_as[Special_field_col] == Special_field_value]

n = 0
for i in Special_cond[Input_label]:
    if i == Input_value:
        n =+ 1

st.subheader("Special Guest Tracking")
if n == 0:
    st.info("No Special Guests have arrived yet")
else:
    st.warning("Take note, certain Special Guests have arrived")
    Special_cond_have = Special_cond[Special_cond[Input_label] == Input_value]
    st.write(Special_cond_have)

st.subheader("Graphical representation for Field 1 - {}".format(field_1))
if not field_1 == "None":
    if streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 1", "Choice of graph"].iloc[0] == "displot":
        createdistplot(df=streamlit_as_Y, X=field_1, title="breakdown")
    elif streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 1", "Choice of graph"].iloc[0] == "V_bar":
        print("bar code - graph 1")
        y = streamlit_as_Y[field_1].value_counts().tolist()
        X = streamlit_as_Y[field_1].value_counts().keys().tolist()

        data = pd.DataFrame({"Categories":X,"Count":y})
        createbar(df=data,X="Categories", y="Count", title="breakdown")

else:
    st.write("No Data field selected")

st.subheader("Graphical representation for Field 2 - {}".format(field_2))
if not field_2 == "None":
    if streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 2", "Choice of graph"].iloc[0] == "displot":
        createdistplot(df=streamlit_as_Y, X=field_2, title="breakdown")
    elif streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 2", "Choice of graph"].iloc[0] == "V_bar":
        print("bar code - graph 2")

        y = streamlit_as_Y[field_2].value_counts().tolist()
        X = streamlit_as_Y[field_2].value_counts().keys().tolist()

        data = pd.DataFrame({"Categories":X,"Count":y})
        createbar(df=data,X="Categories", y="Count", title="breakdown")
else:
    st.write("No Data field selected")

st.subheader("Graphical representation for Field 3 - {}".format(field_3))
if not field_3 == "None":
    print(streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 3", "Choice of graph"].iloc[0])
    if streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 3", "Choice of graph"].iloc[0] == "displot":
        print("Pie code - graph 3")
        createdistplot(df=streamlit_as_Y, X=field_3, title="breakdown")

    elif streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 3", "Choice of graph"].iloc[0] == "V_bar":
        print("bar code - graph 3")

        y = streamlit_as_Y[field_3].value_counts().tolist()
        X = streamlit_as_Y[field_3].value_counts().keys().tolist()

        data = pd.DataFrame({"Categories":X,"Count":y})
        createbar(df=data,X="Categories", y="Count", title="breakdown")
else:
    st.write("No Data field selected")

st.subheader("Graphical representation for Field 4 - {}".format(field_4))
if not field_4 == "None":
    if streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 4", "Choice of graph"].iloc[0] == "displot":
        print("Pie code - graph 4")
        createdistplot(df=streamlit_as_Y, X=field_4, title="breakdown")

    elif streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 4", "Choice of graph"].iloc[0] == "V_bar":
        print("bar code - graph 4")

        y = streamlit_as_Y[field_4].value_counts().tolist()
        X = streamlit_as_Y[field_4].value_counts().keys().tolist()

        data = pd.DataFrame({"Categories":X,"Count":y})
        createbar(df=data,X="Categories", y="Count", title="breakdown")
else:
    st.write("No Data field selected")

st.subheader("Graphical representation for Field 5 - {}".format(field_5))
if not field_5 == "None":
    if streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 5", "Choice of graph"].iloc[0] == "displot":
        createdistplot(df=streamlit_as_Y, X=field_5, title="breakdown")
    elif streamlit_ms.loc[streamlit_ms["Data_col"] == "Data 5", "Choice of graph"].iloc[0] == "V_bar":
        print("bar code - graph 5")

        y = streamlit_as_Y[field_5].value_counts().tolist()
        X = streamlit_as_Y[field_5].value_counts().keys().tolist()

        data = pd.DataFrame({"Categories":X,"Count":y})
        createbar(df=data,X="Categories", y="Count", title="breakdown")
else:
    st.write("No Data field selected")
