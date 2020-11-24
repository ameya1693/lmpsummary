import streamlit as st
import pandas as pd
import os
from openpyxl import Workbook
from io import StringIO

def main():
    st.title("LMP Summary")

    os.chdir("/Users/ameya/Desktop/ABCD")
    path = "/Users/ameya/Desktop/ABCD"
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

    AllNames.loc[(AllNames['Month'] > 5) & (AllNames['Month'] < 10), 'Season'] = 'Summer'
    AllNames.loc[(AllNames['Month'] == 1), 'Season'] = 'Winter'
    AllNames.loc[(AllNames['Month'] == 2), 'Season'] = 'Winter'
    AllNames.loc[(AllNames['Month'] == 12), 'Season'] = 'Winter'
    AllNames.loc[((AllNames['Month'] <= 5) & (AllNames['Month'] > 2)) | (AllNames['Month'] == 11), 'Season'] = 'Summer'

    # print(AllNames['Season'])

    AllNames.loc[(AllNames['Season'] == 'Winter') & (((AllNames['HourOfDay'] > 4) & (AllNames['HourOfDay'] < 12)) | (
                (AllNames['HourOfDay'] > 18) & (AllNames['HourOfDay'] < 23))), 'Peak/Off-Peak'] = 'Peak'
    AllNames.loc[(AllNames['Season'] == 'Winter') & (
                ((AllNames['HourOfDay'] <= 18) & (AllNames['HourOfDay'] >= 12)) | (AllNames['HourOfDay'] <= 4) | (
                    AllNames['HourOfDay'] >= 23)), 'Peak/Off-Peak'] = 'OffPeak'

    AllNames.loc[(AllNames['Season'] == 'Shoulder') & (((AllNames['HourOfDay'] > 5) & (AllNames['HourOfDay'] < 11)) | (
                (AllNames['HourOfDay'] > 17) & (AllNames['HourOfDay'] < 24))), 'Peak/Off-Peak'] = 'Peak'
    AllNames.loc[(AllNames['Season'] == 'Shoulder') & (
                ((AllNames['HourOfDay'] <= 17) & (AllNames['HourOfDay'] >= 11)) | (AllNames['HourOfDay'] <= 5) | (
                    AllNames['HourOfDay'] == 24)), 'Peak/Off-Peak'] = 'OffPeak'

    AllNames.loc[(AllNames['Season'] == 'Summer') & (
    ((AllNames['HourOfDay'] > 13) & (AllNames['HourOfDay'] < 22))), 'Peak/Off-Peak'] = 'Peak'
    AllNames.loc[(AllNames['Season'] == 'Summer') & (
    ((AllNames['HourOfDay'] <= 13) | (AllNames['HourOfDay'] >= 22))), 'Peak/Off-Peak'] = 'OffPeak'


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


    lmp_ATC= pd.DataFrame(zip(Avg1, Avg2, Avg3, Avg4, Avg5, Avg6, Avg7), columns=["2020", "2025", "2030", "2035", "2040", "2045", "2050"], index=[selected_columns])
    st.dataframe(lmp_ATC)

    Avg8 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
                AllNames['Year'] == 2020)& (AllNames['Peak/Off-Peak']== 'OffPeak'), selected_columns].mean()
    Avg9 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
            AllNames['Year'] == 2025)& (AllNames['Peak/Off-Peak']== 'OffPeak'), selected_columns].mean()
    Avg10 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
            AllNames['Year'] == 2030)& (AllNames['Peak/Off-Peak']== 'OffPeak'), selected_columns].mean()
    Avg11 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
            AllNames['Year'] == 2035)& (AllNames['Peak/Off-Peak']== 'OffPeak'), selected_columns].mean()
    Avg12 = AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
            AllNames['Year'] == 2040)& (AllNames['Peak/Off-Peak']== 'OffPeak'), selected_columns].mean()
    Avg13= AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
            AllNames['Year'] == 2045)& (AllNames['Peak/Off-Peak']== 'OffPeak'), selected_columns].mean()
    Avg14= AllNames.loc[(AllNames['Data Item'] == 'Locational Marginal Price ($/MWH)') & (
            AllNames['Year'] == 2050)& (AllNames['Peak/Off-Peak']== 'OffPeak'), selected_columns].mean()

    lmp_offpeak = pd.DataFrame(zip(Avg8, Avg9, Avg10, Avg11, Avg12, Avg13, Avg14),
                           columns=["2020", "2025", "2030", "2035", "2040", "2045", "2050"], index=[selected_columns])
    st.dataframe(lmp_offpeak)



    dataframes = [lmp_offpeak, lmp_ATC]
    lmp_summary = pd.concat(dataframes)

    writer = pd.ExcelWriter("LMP Summary.xlsx")
    lmp_summary.to_excel(writer,"Sheet1")
    writer.save()


if __name__ == '__main__':
  main()

