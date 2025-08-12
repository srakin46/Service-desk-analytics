import streamlit as st
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from io import BytesIO

st.set_page_config(page_title="Service Desk Analytics", layout="wide")

st.title("📊 Service Desk Analytics Tool")

uploaded_file = st.file_uploader("Upload your Excel ticket data", type=["xlsx", "xls"])

if uploaded_file:
    sheet_names = pd.ExcelFile(uploaded_file, engine='openpyxl').sheet_names
    selected_sheet = st.selectbox("Select a sheet to analyze", sheet_names)

    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, engine='openpyxl')

    st.subheader("📈 Preview of Uploaded Data")
    st.dataframe(df.head())

    # Heatmap Section
    st.subheader("🔥 Ticket Volume Heatmap by Hour and Weekday")
    if 'Date' in df.columns and 'Hour' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df['Weekday'] = df['Date'].dt.day_name()
        heatmap_data = df.groupby(['Weekday', 'Hour']).size().unstack(fill_value=0)
        fig, ax = plt.subplots(figsize=(12, 6))
        sns.heatmap(heatmap_data, cmap="YlGnBu", ax=ax)
        st.pyplot(fig)
    else:
        st.warning("Columns 'Date' and 'Hour' are required for heatmap generation.")

    # Forecast Section
    st.subheader("📅 Forecast Ticket Volume for Next 14 Days")
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        daily_counts = df['Date'].value_counts().sort_index()
        model = ExponentialSmoothing(daily_counts, trend='add', seasonal=None)
        fit = model.fit()
        forecast = fit.forecast(14)
        forecast_df = pd.DataFrame({'Date': forecast.index, 'Forecasted Tickets': forecast.values})
        st.line_chart(forecast_df.set_index('Date'))
    else:
        st.warning("Column 'Date' is required for forecasting.")

    # Shift Planning Section
    st.subheader("👥 Shift Planning")
    agent_capacity = st.number_input("Tickets per agent per hour", min_value=1, value=5)
    shift_hours = st.number_input("Shift duration (hours)", min_value=1, value=8)

    if 'Date' in df.columns:
        total_daily_forecast = forecast_df['Forecasted Tickets'].sum() / 14
        tickets_per_agent_per_day = agent_capacity * shift_hours
        agents_needed = int(np.ceil(total_daily_forecast / tickets_per_agent_per_day))

        st.markdown(f"**Estimated Daily Ticket Volume:** {int(total_daily_forecast)}")
        st.markdown(f"**Agents Needed per Day:** {agents_needed}")

        shift_plan = pd.DataFrame({
            'Date': forecast_df['Date'],
            'Forecasted Tickets': forecast_df['Forecasted Tickets'].astype(int),
            'Agents Needed': (forecast_df['Forecasted Tickets'] / tickets_per_agent_per_day).apply(np.ceil).astype(int)
        })

        st.dataframe(shift_plan)

        # Export option
        csv = shift_plan.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Shift Plan CSV", data=csv, file_name="shift_plan.csv", mime='text/csv')
    else:
        st.warning("Column 'Date' is required for shift planning.")
else:
    st.info("Please upload an Excel file to begin analysis.")
