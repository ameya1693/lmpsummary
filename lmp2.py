import streamlit as st
import pandas as pd
import os
from openpyxl import Workbook

def main():
    st.title("LMP Summary")

    os.chdir("/Users/ameya/Desktop/LMP Platform data")
    path = "/Users/ameya/Desktop/LMP Platform data"
    files = os.listdir(path)
    AllNames = pd.DataFrame()

    if f[-4:] == ".csv":

        for f in files:
            info = pd.read_csv(f, 'Sheet1')
            AllNames = AllNames.append(info)
    st.dataframe(AllNames.head())

    all_columns = AllNames.columns.to_list()
    if st.sidebar("Select Node/Zone to show"):
        selected_columns = st.multiselect("Select Columns", all_columns)
    #if st.sidebar.checkbox("Select Run Years"):
        run_years = [2020, 2025, 2030, 2035, 2040, 2045, 2050]
        #selected_runyears = st.multiselect("Run Years", run_years)
            #if st.sidebar.checkbox("Select Data Item"):
                #data_items = st.multiselect("Select Items", "Locational Marginal Price ($/MWH), LMP Total Congestion Component ($/MWH),LMP Loss Component ($/MWH)")

        Avg = AllNames.loc[(AllNames['Data Item']== 'Locational Marginal Price ($/MWH)') |(AllNames['Year']== 2020),selected_columns].mean()
        st.write(pd.DataFrame((Avg),columns=["2020"]))

    #writer = pd.ExcelWriter("LMP Summary.xlsx")
    #AllNames.to_excel(writer,"Sheet1")
    #writer.save()


if __name__ == '__main__':
    main()

