import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import io

# Set Streamlit page config
st.set_page_config(page_title="Service Desk Analytics", layout="wide")

st.title("📊 Service Desk Analytics Tool")

# File uploader
uploaded_file = st.file_uploader("Upload your ticket data (Excel file)", type=["xlsx"])

if uploaded_file:
    # Load Excel file
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    sheet_names = xls.sheet_names
    selected_sheet = st.selectbox("Select a sheet to analyze", sheet_names)

    df = pd.read_excel(xls, sheet_name=selected_sheet, engine='openpyxl')

    st.subheader("📈 Forecast: Ticket Volume for Next 14 Days")
    if 'Processed Tickets' in df.columns and 'Work days' in df.columns:
        df_forecast = df[['Processed Tickets', 'Work days']].dropna()
        daily_avg = df_forecast['Processed Tickets'] / df_forecast['Work days']
        ts = pd.Series(daily_avg.values, index=pd.date_range(end=pd.Timestamp.today(), periods=len(daily_avg)))

        model = ExponentialSmoothing(ts, trend='add', seasonal=None)
        fit = model.fit()
        forecast = fit.forecast(14)

        fig, ax = plt.subplots()
        ts.plot(ax=ax, label='Historical')
        forecast.plot(ax=ax, label='Forecast', color='orange')
        ax.set_title("Ticket Volume Forecast")
        ax.legend()
        st.pyplot(fig)

        st.subheader("🧮 Shift Planning")
        agent_capacity = st.number_input("Tickets an agent can handle per day", min_value=1, value=40)
        forecast_df = pd.DataFrame({
            'Date': forecast.index,
            'Forecasted Tickets': forecast.values,
            'Agents Needed': (forecast.values / agent_capacity).astype(int)
        })
        st.dataframe(forecast_df)

        # Download shift plan
        csv = forecast_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Shift Plan CSV", csv, "shift_plan.csv", "text/csv")

    st.subheader("🔥 Heatmap: Average Tickets per Day by Region and Employee")
    if 'Location' in df.columns and 'Employee Name' in df.columns and 'Avg. Tickets / day' in df.columns:
        df_heatmap = df[['Location', 'Employee Name', 'Avg. Tickets / day']].dropna()
        # Map locations to regions
        region_map = {
            'Bangalore': 'APAC',
            'Mumbai': 'APAC',
            'London': 'EMEA',
            'Warsaw': 'EMEA',
            'Houston': 'AMER',
            'Denver': 'AMER'
        }
        df_heatmap['Region'] = df_heatmap['Location'].map(region_map).fillna('Other')

        pivot_table = df_heatmap.pivot_table(index='Employee Name', columns='Region', values='Avg. Tickets / day', aggfunc='mean')

        fig, ax = plt.subplots(figsize=(12, 8))
        sns.heatmap(pivot_table, annot=True, fmt=".1f", cmap="YlGnBu", ax=ax)
        ax.set_title("Average Tickets per Day by Region and Employee")
        st.pyplot(fig)
    else:
        st.warning("Required columns for heatmap not found in the selected sheet.")
else:
    st.info("Please upload an Excel file to begin analysis.")
