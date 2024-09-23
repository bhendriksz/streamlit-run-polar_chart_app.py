import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import io  # Import io for in-memory file handling

# Define font properties for a general sans-serif font
font_properties = FontProperties(family='sans-serif', size=18)

# Function to create detailed polar charts
def create_detailed_polar_chart(data, categories, colors, light_colors, title, question_number):
    segment_angle = 2 * np.pi / len(categories)
    angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False)
    
    # Calculate sub-segment width
    sub_segment_width = segment_angle / 4  # Adjust sub-segment width for clarity
    
    for k, category in enumerate(categories):
        fig, ax = plt.subplots(figsize=(8, 8), subplot_kw={'projection': 'polar'}, dpi=300)
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
            offset = 0.265  # Adjust the offset value

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
        
        # Save figure to a BytesIO object with high DPI for sharpness
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=600)  # Save as PNG with high DPI (600)
        buf.seek(0)  # Move the buffer's position to the beginning

        st.download_button(
            label=f"Download Chart as High-Quality PNG for {category}",
            data=buf,
            file_name=f"polar_chart_{question_number}_{category}.png",
            mime="image/png"
        )

# Streamlit app starts here
st.title("Detailed Polar Chart Visualization App")

# Let the user upload an Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    sheet_name = st.selectbox('Select the sheet name', ['Question 4', 'Question 5', 'Question 6', 'Question 7', 'Question 8'])

    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    
    questions = {
        'Question 4': {
            'categories': ['Strategic Alignment', 'Balance', 'Maximal Value'],
            'colors': ['#7CAEAD', '#917670', '#CDB486'],
            'light_colors': ['#C0E4E0', '#D6CCC8', '#E4D9D3'],
            'data_ranges': [(3, 7), (7, 11), (11, 15)],
            'transform': lambda values: np.array([values[0], values[3], 5 - values[1], 5 - values[2]])  # Specific transformation for balance
        },
        'Question 5': {
            'categories': ['Portfolio Mindset', 'Focus', 'Agility'],
            'colors': ['#7CAEAD', '#917670', '#CDB486'],
            'light_colors': ['#C0E4E0', '#D6CCC8', '#E4D9D3'],
            'data_ranges': [(3, 8), (8, 12), (12, 16)]
        },
        'Question 6': {
            'categories': ['Evidence', 'Informal Power', 'Opinion'],
            'colors': ['#7CAEAD', '#917670', '#CDB486'],
            'light_colors': ['#C0E4E0', '#D6CCC8', '#E4D9D3'],
            'data_ranges': [(3, 7), (7, 11), (11, 15)]
        },
        'Question 7': {
            'categories': ['Cross-functional collaboration', 'Critical thinking', 'Market immersion'],
            'colors': ['#7CAEAD', '#917670', '#CDB486'],
            'light_colors': ['#C0E4E0', '#D6CCC8', '#E4D9D3'],
            'data_ranges': [(3, 7), (7, 11), (11, 15)]
        },
        'Question 8': {
            'categories': ['Culture'],
            'colors': ['#7CAEAD'],
            'light_colors': ['#C0E4E0'],
            'data_ranges': [(3, 7)]
        }
    }
    
    # Extract the relevant information for the selected question
    question = questions[sheet_name]
    categories = question['categories']
    colors = question['colors']
    light_colors = question['light_colors']
    data_ranges = question['data_ranges']
    
    # Extract data from the selected sheet
    data = {}
    for idx, category in enumerate(categories):
        start_row, end_row = data_ranges[idx]
        data[category] = df.iloc[start_row:end_row, 12].values
    
    # Calculate averages for balance if Question 4
    if sheet_name == 'Question 4':
        balance_values = data['Balance']
        transformed_balance_values = question['transform'](balance_values)
        averages = {key: (np.mean(transformed_balance_values) if key == 'Balance' else np.mean(values)) for key, values in data.items()}
    else:
        averages = {key: np.mean(values) for key, values in data.items()}
    
    # Generate and display the detailed polar chart
    create_detailed_polar_chart(data, categories, colors, light_colors, sheet_name, sheet_name)
    
    st.write("Averages:", averages)
