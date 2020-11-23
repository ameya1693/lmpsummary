import streamlit as st
import pandas as pd
import os
from openpyxl import Workbook
from io import StringIO

def main():
    st.title("LMP Summary")

    os.chdir("/Users/ameya/Desktop/CD")
    path = "/Users/ameya/Desktop/CD"
    files = os.listdir(path)
    AllNames = pd.DataFrame()



    for f in files:
        info = pd.read_excel(f)
        AllNames = AllNames.append(info)
    st.dataframe(AllNames.head())


    all_columns = AllNames.columns.to_list()
    #if st.sidebar("Select Node/Zone to show"):
    selected_columns = st.multiselect("Select Columns", all_columns)
    #if st.sidebar.checkbox("Select Run Years"):
    run_years = [2020, 2025, 2030, 2035, 2040, 2045, 2050]
        #selected_runyears = st.multiselect("Run Years", run_years)
            #if st.sidebar.checkbox("Select Data Item"):
                #data_items = st.multiselect("Select Items", "Locational Marginal Price ($/MWH), LMP Total Congestion Component ($/MWH),LMP Loss Component ($/MWH)")

    Avg1 = AllNames.loc[(AllNames['Data Item']== 'Locational Marginal Price ($/MWH)')&(AllNames['Year']== 2020),selected_columns].mean()
    Avg2 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
                AllNames['Year'] == 2025), selected_columns].mean()
    Avg3 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
                AllNames['Year'] == 2030), selected_columns].mean()
    Avg4 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
                AllNames['Year'] == 2035), selected_columns].mean()
    Avg5 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
                AllNames['Year'] == 2040), selected_columns].mean()
    Avg6 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
                AllNames['Year'] == 2045), selected_columns].mean()
    Avg7 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
                AllNames['Year'] == 2050), selected_columns].mean()


    lmp_summary= pd.DataFrame(zip(Avg1, Avg2, Avg3, Avg4, Avg5, Avg6, Avg7), columns=["2020", "2025", "2030", "2035", "2040", "2045", "2050"], index=[selected_columns])
    st.dataframe(lmp_summary)

    writer = pd.ExcelWriter("LMP Summary.xlsx")
    lmp_summary.to_excel(writer,"Sheet1")
    writer.save()


if __name__ == '__main__':
  main()

