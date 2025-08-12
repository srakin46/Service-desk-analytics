import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from io import BytesIO

# Set Streamlit page config
st.set_page_config(page_title="Service Desk Analytics", layout="wide")

# Title
st.title("📊 Service Desk Analytics Tool")

# File uploader
uploaded_file = st.file_uploader("Upload your ticket data file (Excel)", type=["xlsx"])

if uploaded_file:
    # Load Excel file
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    sheet_names = xls.sheet_names
    sheet = st.selectbox("Select sheet for forecasting", sheet_names)

    # Load selected sheet
    df = pd.read_excel(xls, sheet_name=sheet, engine='openpyxl')

    # Forecasting logic for 'indira' sheet
    if sheet.lower() == 'indira':
        st.subheader("📈 Forecast: Ticket Volume for Next 14 Days")

        # Extract date column
        date_col = df.columns[0]
        df_forecast = df[[date_col]].dropna()
        df_forecast[date_col] = pd.to_datetime(df_forecast[date_col], errors='coerce')
        df_forecast = df_forecast.dropna()

        # Count tickets per day
        ticket_counts = df_forecast.groupby(date_col).size().reset_index(name='ticket_count')
        ticket_counts = ticket_counts.set_index(date_col).asfreq('D', fill_value=0)

        # Fit Holt-Winters model
        model = ExponentialSmoothing(ticket_counts['ticket_count'], seasonal='add', seasonal_periods=7)
        fit = model.fit()
        forecast = fit.forecast(14)

        # Plot forecast
        fig, ax = plt.subplots(figsize=(10, 4))
        ticket_counts['ticket_count'].plot(ax=ax, label='Historical')
        forecast.plot(ax=ax, label='Forecast', color='orange')
        ax.set_title("Forecasted Ticket Volume for Next 14 Days")
        ax.set_ylabel("Tickets")
        ax.legend()
        st.pyplot(fig)

        # Show forecast table
        forecast_df = forecast.reset_index()
        forecast_df.columns = ['Date', 'Forecasted Tickets']
        st.dataframe(forecast_df)

        # Download forecast CSV
        csv = forecast_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Forecast CSV", data=csv, file_name="forecast.csv", mime="text/csv")
    else:
        st.info("Please select the 'indira' sheet to view the forecast.")
