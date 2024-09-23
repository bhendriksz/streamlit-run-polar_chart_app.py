#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Sep 23 17:21:18 2024

@author: bramhendriksz
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import numpy as np
import io

# Integration of Segoe UI web fonts
st.markdown("""
    <style>
    @font-face {
        font-family: 'SegoeUI';
        src: url(//c.s-microsoft.com/static/fonts/segoe-ui/west-european/light/latest.woff2) format('woff2'),
             url(//c.s-microsoft.com/static/fonts/segoe-ui/west-european/light/latest.woff) format('woff'),
             url(//c.s-microsoft.com/static/fonts/segoe-ui/west-european/light/latest.ttf) format('truetype');
        font-weight: 100;
    }
    @font-face {
        font-family: 'SegoeUI';
        src: url(//c.s-microsoft.com/static/fonts/segoe-ui/west-european/bold/latest.woff2) format('woff2'),
             url(//c.s-microsoft.com/static/fonts/segoe-ui/west-european/bold/latest.woff) format('woff'),
             url(//c.s-microsoft.com/static/fonts/segoe-ui/west-european/bold/latest.ttf) format('truetype');
        font-weight: 700;
    }
    body {
        font-family: 'SegoeUI';
    }
    </style>
""", unsafe_allow_html=True)

# Define a function to extract data from the uploaded Excel file
def extract_data(df, department_name, sheet_name, start_row, end_row, columns):
    statements = df.iloc[start_row:end_row, columns['statement']].values
    responses = df.iloc[start_row:end_row, columns['responses']].values
    weighted_average = df.iloc[start_row:end_row, columns['weighted_average']].values

    data = {
        "Statement": statements,
        "Strongly Disagree": responses[:, 0],
        "Disagree": responses[:, 1],
        "Neutral": responses[:, 2],
        "Agree": responses[:, 3],
        "Strongly Agree": responses[:, 4],
        "Weighted Average": weighted_average
    }

    department_df = pd.DataFrame(data)
    department_df['Department'] = department_name
    department_df['Question'] = sheet_name
    
    return department_df

# Streamlit app starts here
st.title("Bar Chart Generator")

uploaded_files = st.file_uploader("Upload Excel files for each department", type=["xlsx"], accept_multiple_files=True)
if uploaded_files:
    sheets_info = {
        "Question 4": {"start_row": 2, "end_row": 14, "columns": {"statement": 0, "responses": slice(1, 6), "weighted_average": 6}},
        "Question 5": {"start_row": 2, "end_row": 15, "columns": {"statement": 0, "responses": slice(1, 6), "weighted_average": 6}},
        "Question 6": {"start_row": 2, "end_row": 14, "columns": {"statement": 0, "responses": slice(1, 6), "weighted_average": 6}},
        "Question 7": {"start_row": 2, "end_row": 14, "columns": {"statement": 0, "responses": slice(1, 6), "weighted_average": 6}},
    }

    combined_df = pd.DataFrame()

    for uploaded_file in uploaded_files:
        department_name = st.text_input(f"Enter the department name for file: {uploaded_file.name}")
        if department_name:
            for sheet, info in sheets_info.items():
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                sheet_df = extract_data(df, department_name, sheet, info['start_row'], info['end_row'], info['columns'])
                combined_df = pd.concat([combined_df, sheet_df], ignore_index=True)

    statement_options = combined_df["Statement"].unique()
    selected_statement = st.selectbox("Select a statement for visualization", statement_options)

    if selected_statement:
        st.write(f"Generating chart for: {selected_statement}")
        df_statement = combined_df[combined_df['Statement'] == selected_statement]
        df_statement = df_statement.pivot(index='Department', columns='Question', values=["Strongly Disagree", "Disagree", "Neutral", "Agree", "Strongly Agree"])
        df_statement.columns = [col[0] for col in df_statement.columns]

        df_percentage = df_statement.div(df_statement.sum(axis=1), axis=0) * 100

        fig, ax = plt.subplots(figsize=(12, 8), dpi=300)  # Increase DPI for high quality
        bars = df_percentage.plot(kind='barh', stacked=True, color=["#C00000", "#DE7E35", "#FFFBB9", "#A7C23D", "#4F7A27"], ax=ax, zorder=3)

        ax.xaxis.grid(True, color='gray', linestyle='--', linewidth=0.5, zorder=1)
        for spine in ax.spines.values():
            spine.set_visible(False)

        ax.set_title(selected_statement, fontsize=16, weight='bold', pad=20)
        ax.set_xlabel('Percentage of Responses', fontsize=14)
        ax.set_ylabel('Department', fontsize=14)
        ax.set_xticks(np.arange(0, 101, 10))
        ax.set_xticklabels([f'{i}%' for i in range(0, 101, 10)])
        ax.set_yticklabels(df_percentage.index)
        ax.yaxis.set_ticks_position('none')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=5, title='Response', frameon=False)

        plt.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=600, bbox_inches='tight')
        buf.seek(0)
        st.pyplot(fig)

        st.download_button(
            label="Download chart as PNG",
            data=buf,
            file_name=f"{selected_statement}.png",
            mime="image/png"
        )

    st.markdown("""
    ### Instructions:
    1. Upload the Excel files for each department.
    2. Enter the department name for each file.
    3. Select a statement from the dropdown menu.
    4. View the bar chart and download it in high quality.
    """)
