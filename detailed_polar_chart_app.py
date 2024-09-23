#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 29 14:53:39 2024

@author: bramhendriksz
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import io

# Load font properties for Segoe UI
font_properties = FontProperties(family='sans-serif', size=18)

# Function to create radar chart for a given question's data
def create_radar_chart(data, categories, title, sheet_name):
    colors = ['#7CAEAD', '#917670', '#CDB486']
    light_colors = ['#C0E4E0', '#D6CCC8', '#E4D9D3']

    figsize = (8, 8)
    dpi = 120
    segment_angle = 2 * np.pi / len(categories)
    angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False)

    for k, category in enumerate(categories):
        fig, ax = plt.subplots(figsize=figsize, subplot_kw={'projection': 'polar'}, dpi=dpi)
        ax.set_theta_direction(-1)
        ax.set_theta_offset(np.pi / 2)
        ax.set_ylim(0, 5)
        ax.set_yticks(np.arange(1, 6))
        ax.set_yticklabels([])  # Hide radial grid labels
        ax.set_xticklabels([])  # Remove angle labels
        ax.yaxis.grid(True, linewidth=0.75)
        ax.xaxis.grid(False)
        ax.spines['polar'].set_visible(False)

        bars = []

        for i, (cat, values) in enumerate(data.items()):
            base_angle = angles[i]
            sub_segment_width = segment_angle / len(values)
            offset = 0.21 if len(values) == 5 else 0.265
            
            for j, value in enumerate(values):
                sub_angle = base_angle + offset + j * sub_segment_width
                color = colors[i] if i == k else light_colors[i]
                bar = ax.bar(sub_angle, value, width=sub_segment_width, color=color, alpha=0.75, edgecolor='white', label=cat if (j == 0 and i == k) else "")
                
                if i == k:
                    ax.text(sub_angle, value - 0.75, f'{value:.2f}', ha='center', va='center', color='black', fontweight='bold', fontsize=17, fontproperties=font_properties)
                    
                    if j == 0:
                        bars.append(bar)

        for angle in angles:
            ax.axvline(x=angle, color='gray', linestyle='--', linewidth=1)

        ax.legend(handles=[b[0] for b in bars], labels=[category], loc='upper right', bbox_to_anchor=(1.1, 1.1), prop=font_properties)

        plt.tight_layout()
        fig.patch.set_alpha(0.0)
        ax.patch.set_alpha(1.0)
        
        st.pyplot(fig)

        # Save the figure as a downloadable high-res PNG
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=600)  # Save as high-res PNG
        buf.seek(0)  # Move the buffer's position to the beginning
        st.download_button(
            label=f"Download Chart for {category}",
            data=buf,
            file_name=f"radar_chart_{sheet_name}_{category}.png",
            mime="image/png"
        )

# Streamlit App
st.title("Radar Chart Visualization App")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    sheet_name = st.selectbox('Select the sheet name', ['Question 4', 'Question 5', 'Question 6', 'Question 7', 'Question 8'])
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    
    if sheet_name == 'Question 4':
        data_q4 = {
            'Strategic Alignment': df.iloc[3:7, 12].values,
            'Balance': df.iloc[7:11, 12].values,
            'Maximal Value': df.iloc[11:15, 12].values
        }
        # Apply the specific average calculation for the "balance" category
        balance_values = data_q4['Balance']
        transformed_balance_values = np.array([balance_values[0], balance_values[3], 5 - balance_values[1], 5 - balance_values[2]])
        
        averages_q4 = {key: (np.mean(transformed_balance_values) if key == 'Balance' else np.mean(values)) for key, values in data_q4.items()}
        
        categories_q4 = list(data_q4.keys())
        create_radar_chart(data_q4, categories_q4, "Portfolio Success Visualization", sheet_name)
        
    elif sheet_name == 'Question 5':
        data_q5 = {
            'Portfolio Mindset': df.iloc[3:8, 12].values,
            'Focus': df.iloc[8:12, 12].values,
            'Agility': df.iloc[12:16, 12].values
        }
        averages_q5 = {key: np.mean(values) for key, values in data_q5.items()}
        categories_q5 = list(data_q5.keys())
        create_radar_chart(data_q5, categories_q5, "Effectiveness Visualization", sheet_name)
        
    elif sheet_name == 'Question 6':
        data_q6 = {
            'Evidence': df.iloc[3:7, 12].values,
            'Informal Power': df.iloc[7:11, 12].values,
            'Opinion': df.iloc[11:15, 12].values
        }
        averages_q6 = {key: np.mean(values) for key, values in data_q6.items()}
        categories_q6 = list(data_q6.keys())
        create_radar_chart(data_q6, categories_q6, "Decision Making Visualization", sheet_name)
        
    elif sheet_name == 'Question 7':
        data_q7 = {
            'Cross-functional collaboration': df.iloc[3:7, 12].values,
            'Critical thinking': df.iloc[7:11, 12].values,
            'Market immersion': df.iloc[11:15, 12].values
        }
        averages_q7 = {key: np.mean(values) for key, values in data_q7.items()}
        categories_q7 = list(data_q7.keys())
        create_radar_chart(data_q7, categories_q7, "Input Processes Visualization", sheet_name)
        
    elif sheet_name == 'Question 8':
        data_q8 = {
            'Culture': df.iloc[3:7, 12].values
        }
        averages_q8 = {key: np.mean(values) for key, values in data_q8.items()}
        categories_q8 = list(data_q8.keys())
        create_radar_chart(data_q8, categories_q8, "Collective Ambition Visualization", sheet_name)

    # Show instructions
    st.markdown("""
    ### Instructions:
    1. Upload an Excel file containing the data.
    2. Select the appropriate sheet name.
    3. View and download the resulting radar chart for each category in high quality.
    """)
