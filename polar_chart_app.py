#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 20 15:22:43 2024

@author: bramhendriksz
"""
pip install matplotlib

import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import io  # Import io for in-memory file handling

# Define font properties for a general sans-serif font (alternative to Segoe UI)
font_properties = FontProperties(family='sans-serif', size=18)

# Function to create polar chart
def create_polar_chart(data, averages, categories, colors, title):
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw={'projection': 'polar'}, dpi=120)
    ax.set_theta_direction(-1)
    ax.set_theta_offset(np.pi / 2)
    ax.set_ylim(0, 5)
    ax.set_yticks(np.arange(1, 6))
    ax.set_yticklabels([], fontproperties=font_properties)  # Hide radial grid labels
    ax.set_xticklabels([], fontproperties=font_properties)  # Remove angle labels
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
        ax.text(angles[i], avg - 1, f'{avg:.2f}', ha='center', va='bottom', color='black', fontweight='bold', fontsize=16, fontproperties=font_properties)

    # Add separation lines at the edges of each bar
    separation_angles = [(angle - segment_angle / 2) % (2 * np.pi) for angle in angles] + [2 * np.pi]
    for angle in separation_angles:
        ax.axvline(x=angle, color='gray', linestyle='--', linewidth=1)

    # Add a legend
    ax.legend(handles=[h[0] for h in legend_handles], labels=categories, loc='upper right', bbox_to_anchor=(1.1, 1.1), fontsize=20, prop=font_properties)

    plt.tight_layout()

    # Set transparent background for the figure
    fig.patch.set_alpha(0.0)  # Make the figure background transparent
    ax.patch.set_alpha(1.0)   # Keep the polar chart itself non-transparent

    return fig

# Streamlit app starts here
st.title("Polar Chart App")

# Let the user upload an Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    sheet_name = st.selectbox('Select the sheet name', ['Question 4', 'Question 5', 'Question 6', 'Question 7', 'Question 8'])

    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    
    # Define the structure for each question
    questions = {
        'Question 4': {
            'categories': ['Strategic Alignment', 'Balance', 'Maximal Value'],
            'colors': ['#7CAEAD', '#917670', '#CDB486'],
            'data_ranges': [(3, 7), (7, 11), (11, 15)],
            'transform': lambda values: np.array([values[0], values[3], 5 - values[1], 5 - values[2]])  # Specific transformation for balance
        },
        'Question 5': {
            'categories': ['Portfolio Mindset', 'Focus', 'Agility'],
            'colors': ['#7CAEAD', '#917670', '#CDB486'],
            'data_ranges': [(3, 8), (8, 12), (12, 16)]
        },
        'Question 6': {
            'categories': ['Evidence', 'Informal Power', 'Opinion'],
            'colors': ['#7CAEAD', '#917670', '#CDB486'],
            'data_ranges': [(3, 7), (7, 11), (11, 15)]
        },
        'Question 7': {
            'categories': ['Cross-functional collaboration', 'Critical thinking', 'Market immersion'],
            'colors': ['#7CAEAD', '#917670', '#CDB486'],
            'data_ranges': [(3, 7), (7, 11), (11, 15)]
        },
        'Question 8': {
            'categories': ['Culture'],
            'colors': ['#7CAEAD'],
            'data_ranges': [(3, 7)]
        }
    }
    
    # Extract the relevant information for the selected question
    question = questions[sheet_name]
    categories = question['categories']
    colors = question['colors']
    data_ranges = question['data_ranges']
    
    # Extract data from the selected sheet
    data = {}
    for idx, category in enumerate(categories):
        start_row, end_row = data_ranges[idx]
        data[category] = df.iloc[start_row:end_row, 12].values
    
    # Calculate the average values
    if sheet_name == 'Question 4':
        # Apply the transformation to Balance
        balance_values = data['Balance']
        transformed_balance_values = question['transform'](balance_values)
        averages = {key: (np.mean(transformed_balance_values) if key == 'Balance' else np.mean(values)) for key, values in data.items()}
    else:
        averages = {key: np.mean(values) for key, values in data.items()}
    
    # Generate the chart
    fig = create_polar_chart(data, averages, categories, colors, f"{sheet_name}: Polar Chart")
    st.pyplot(fig)

    # Save figure to a BytesIO object
    buf = io.BytesIO()
    fig.savefig(buf, format="png")  # Save as PNG
    buf.seek(0)  # Move the buffer's position to the beginning

    # Create a download button
    st.download_button(
        label=f"Download Chart as PNG for {sheet_name}",
        data=buf,
        file_name=f"polar_chart_{sheet_name}.png",
        mime="image/png"
    )

    # Show instructions
    st.markdown("""
    ### Instructions:
    1. Upload an Excel file containing the data.
    2. Select the appropriate sheet name.
    3. View and download the resulting polar chart.
    """)
