#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 20 15:22:43 2024

@author: bramhendriksz
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

# Define font properties for Segoe UI
segoe_ui = FontProperties(fname='/path/to/your/Segoe UI Bold.ttf', size=25)  # Update the path

# Function to create polar chart
def create_polar_chart(data, averages, categories, colors, title):
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw={'projection': 'polar'}, dpi=120)
    ax.set_theta_direction(-1)
    ax.set_theta_offset(np.pi / 2)
    ax.set_ylim(0, 5)
    ax.set_yticks(np.arange(1, 6))
    ax.set_yticklabels([], fontproperties=segoe_ui)  # Hide radial grid labels
    ax.set_xticklabels([], fontproperties=segoe_ui)  # Remove angle labels
    ax.yaxis.grid(True, linewidth=0.75)  # Adjust the width of the circular grid lines
    ax.xaxis.grid(False)
    ax.spines['polar'].set_visible(False)

    # The angle for each segment
    segment_angle = 2 * np.pi / len(categories)
    angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False)
    angles = [(angle + segment_angle / 2) % (2 * np.pi) for angle in angles]  # Center bars on angles

    # Draw bars and add standard deviation lines
    legend_handles = []
    cap_length = 0.02  # Length of the cap line, adjust as needed

    for i, (category, avg) in enumerate(averages.items()):
        # Draw bars
        bar = ax.bar(angles[i], avg, color=colors[i], alpha=0.75, width=segment_angle, label=category)
        legend_handles.append(bar)
        
        # Annotate bars with the average value
        ax.text(angles[i], avg - 1, f'{avg:.2f}', ha='center', va='bottom', color='black', fontweight='bold', fontsize=20, fontproperties=segoe_ui)

    # Add separation lines at the edges of each bar
    separation_angles = [(angle - segment_angle / 2) % (2 * np.pi) for angle in angles] + [2 * np.pi]
    for angle in separation_angles:
        ax.axvline(x=angle, color='gray', linestyle='--', linewidth=1)

    # Add a legend
    ax.legend(handles=[h[0] for h in legend_handles], labels=categories, loc='upper right', bbox_to_anchor=(1.1, 1.1), fontsize=20, prop=segoe_ui)

    plt.tight_layout()
    return fig

# Streamlit app starts here
st.title("Polar Chart App")

# Let the user upload an Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    sheet_name = st.selectbox('Select the sheet name', ['Question 4', 'Question 5', 'Question 6', 'Question 7', 'Question 8'])

    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    
    if sheet_name == 'Question 4':
        data = {
            'Strategic Alignment': df.iloc[3:7, 12].values,
            'Balance': df.iloc[7:11, 12].values,
            'Maximal Value': df.iloc[11:15, 12].values
        }
        
        # Transform balance values
        balance_values = data['Balance']
        transformed_balance_values = np.array([balance_values[0], balance_values[3], 5 - balance_values[1], 5 - balance_values[2]])

        # Calculate averages
        averages = {key: (np.mean(transformed_balance_values) if key == 'Balance' else np.mean(values)) for key, values in data.items()}

        # Define categories and colors
        categories = list(data.keys())
        colors = ['#7CAEAD', '#917670', '#CDB486']  # Custom colors
        
        # Generate the chart
        fig = create_polar_chart(data, averages, categories, colors, "Question 4: Polar Chart")
        st.pyplot(fig)
        
    elif sheet_name == 'Question 5':
        data = {
            'Portfolio Mindset': df.iloc[3:8, 12].values,
            'Focus': df.iloc[8:12, 12].values,
            'Agility': df.iloc[12:16, 12].values
        }

        # Calculate averages
        averages = {key: np.mean(values) for key, values in data.items()}
        
        # Define categories and colors
        categories = list(data.keys())
        colors = ['#7CAEAD', '#917670', '#CDB486']  # Custom colors
        
        # Generate the chart
        fig = create_polar_chart(data, averages, categories, colors, "Question 5: Polar Chart")
        st.pyplot(fig)

    elif sheet_name == 'Question 6':
        data = {
            'Evidence': df.iloc[3:7, 12].values,
            'Informal Power': df.iloc[7:11, 12].values,
            'Opinion': df.iloc[11:15, 12].values
        }

        # Calculate averages
        averages = {key: np.mean(values) for key, values in data.items()}
        
        # Define categories and colors
        categories = list(data.keys())
        colors = ['#7CAEAD', '#917670', '#CDB486']  # Custom colors
        
        # Generate the chart
        fig = create_polar_chart(data, averages, categories, colors, "Question 6: Polar Chart")
        st.pyplot(fig)
    
    elif sheet_name == 'Question 7':
        data = {
            'Cross-functional collaboration': df.iloc[3:7, 12].values,
            'Critical thinking': df.iloc[7:11, 12].values,
            'Market immersion': df.iloc[11:15, 12].values
        }

        # Calculate averages
        averages = {key: np.mean(values) for key, values in data.items()}
        
        # Define categories and colors
        categories = list(data.keys())
        colors = ['#7CAEAD', '#917670', '#CDB486']  # Custom colors
        
        # Generate the chart
        fig = create_polar_chart(data, averages, categories, colors, "Question 7: Polar Chart")
        st.pyplot(fig)
    
    elif sheet_name == 'Question 8':
        data = {
            'Culture': df.iloc[3:7, 12].values
        }

        # Calculate averages
        averages = {key: np.mean(values) for key, values in data.items()}
        
        # Define categories and colors
        categories = list(data.keys())
        colors = ['#7CAEAD', '#917670', '#CDB486']  # Custom colors
        
        # Generate the chart
        fig = create_polar_chart(data, averages, categories, colors, "Question 8: Polar Chart")
        st.pyplot(fig)

# Add instructions
st.markdown("""
### Instructions:
1. Upload an Excel file containing the data.
2. Select the appropriate sheet name.
3. View the resulting polar chart.
""")
