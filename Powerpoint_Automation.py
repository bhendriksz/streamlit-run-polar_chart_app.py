import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from collections import defaultdict
import os
import colorsys

# Helper function to convert RGB values
def RGB(r, g, b):
    return RGBColor(r, g, b)

# Function to convert HSV to RGB and return as RGBColor
def hsv_to_rgb(h, s, v):
    r, g, b = colorsys.hsv_to_rgb(h, s, v)
    return RGB(int(r * 255), int(g * 255), int(b * 255))

# Define a mapping from department (two-letter code) to colors
department_colors = {}

# Function to assign a unique color to each department using HSV color space
def assign_department_color(department):
    if department not in department_colors:
        hue = len(department_colors) / 12.0  # Adjust denominator to control the spread of hues
        department_colors[department] = hsv_to_rgb(hue, 0.8, 0.9)  # Set saturation and value to fixed levels
    return department_colors[department]

# Function to add bullets to a slide with grouping by department
def add_bullets_to_slide_with_grouping(slide, projects, rows, cols, cell_width_cm, cell_height_cm, bol_diameter_pt):
    bullet_count = defaultdict(int)
    department_shapes = defaultdict(list)

    # Convert cell width and height from cm to inches
    cell_width_in = cell_width_cm / 2.54  # 1 inch = 2.54 cm
    cell_height_in = cell_height_cm / 2.54

    # Convert bullet (boll) diameter from pt to inches
    bol_diameter_in = bol_diameter_pt / 72.0  # 72 points = 1 inch

    # Loop over each project
    for afkorting, spf, project_title, project_type, source_slide_idx in projects:
        col_letter = spf[0]  # First letter for column
        row_number = int(spf[1:])  # Number for the row
        col_number = ord(col_letter.upper()) - ord('A') + 1  # Convert column letter to index

        # Only place the bullet if it fits within the grid
        if 1 <= row_number <= rows and 1 <= col_number <= cols:
            bullet_count[(row_number, col_number)] += 1
            bullet_idx = bullet_count[(row_number, col_number)] - 1

            # Maximum of 5 bullets per row, and 2 rows per cell
            visible_bullet_idx = bullet_idx % 10  # 0-9 visible bullets
            row_offset = (visible_bullet_idx // 5) * bol_diameter_in  # Offset for second row of bullets
            col_offset = (visible_bullet_idx % 5) * bol_diameter_in  # Offset for bullet position

            # Calculate the position on the slide in inches
            xPos = Inches(1 + (col_number - 1) * cell_width_in + col_offset)
            yPos = Inches(1 + (row_number - 1) * cell_height_in + row_offset)

            # Add bullet for the project
            shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, xPos, yPos, Inches(bol_diameter_in), Inches(bol_diameter_in))
            shape.text = afkorting

            # Assign a unique color to the department based on the two-letter code
            department_color = assign_department_color(afkorting[:2])
            fill = shape.fill
            fill.solid()
            fill.fore_color.rgb = department_color

            # Set text font style and size
            text_frame = shape.text_frame
            text_frame.text = afkorting
            text_frame.paragraphs[0].font.name = "Arial"
            text_frame.paragraphs[0].font.size = Pt(12)
            text_frame.paragraphs[0].font.bold = True
            text_frame.paragraphs[0].font.italic = True

            # Collect shapes for grouping by department
            department_shapes[afkorting[:2]].append(shape)

# Function to process and modify the PowerPoint presentation
def process_presentation(ppt_path, rows, cols, cell_width_cm, cell_height_cm, bol_diameter_pt):
    try:
        # Open the PowerPoint file
        prs = Presentation(ppt_path)

        # Initialize a dictionary to store department-wise slides
        department_slides = defaultdict(list)
        department_initiatives_ideas = defaultdict(list)
        department_tasks = defaultdict(list)

        # Loop through all slides and shapes to find tables
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_table:
                    tbl = shape.table
                    afkorting_col = None
                    spf_col = None
                    title_col = None
                    type_col = None

                    # Identify the columns "AFKORTING", "SPF", "TITEL", and "PROJECT / IDEA / TASK"
                    for col_idx, cell in enumerate(tbl.rows[0].cells):
                        header_text = cell.text.upper()
                        if header_text == "AFKORTING":
                            afkorting_col = col_idx
                        elif header_text == "SPF":
                            spf_col = col_idx
                        elif header_text == "TITEL":
                            title_col = col_idx
                        elif header_text == "PROJECT / IDEA / TASK":
                            type_col = col_idx

                    if afkorting_col is not None and spf_col is not None and type_col is not None:
                        for row in tbl.rows[1:]:
                            afkorting = row.cells[afkorting_col].text
                            spf = row.cells[spf_col].text
                            project_title = row.cells[title_col].text if title_col is not None else ""
                            project_type = row.cells[type_col].text

                            if len(afkorting) >= 2:
                                department = afkorting[:2]  # First two letters define the department
                                department_slides[department].append((afkorting, spf, project_title, project_type, slide))

                                # Categorize into initiatives/ideas or tasks
                                if project_type in ['Initiative', 'Idea']:
                                    department_initiatives_ideas[department].append((afkorting, spf, project_title, project_type, slide))
                                elif project_type == 'Task':
                                    department_tasks[department].append((afkorting, spf, project_title, project_type, slide))

        # Create slide with all projects, using different colors for each department
        new_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Add a blank slide layout
        title_shape = new_slide.shapes.title
        title_shape.text = "All Projects by Department"

        all_projects = [project for projects in department_slides.values() for project in projects]
        add_bullets_to_slide_with_grouping(new_slide, all_projects, rows, cols, cell_width_cm, cell_height_cm, bol_diameter_pt)

        # Save the modified presentation
        output_path = os.path.join(os.path.dirname(ppt_path), "updated_presentation_departments.pptx")
        prs.save(output_path)

        return output_path

    except Exception as e:
        st.error(f"An error occurred: {e}")
        raise e


# Streamlit UI setup
st.title("PowerPoint Processing Tool")

uploaded_ppt = st.file_uploader("Upload your PowerPoint presentation", type=["pptx"])

if uploaded_ppt is not None:
    with open("uploaded_presentation.pptx", "wb") as f:
        f.write(uploaded_ppt.read())

    st.success("PowerPoint successfully uploaded!")

    rows = st.number_input("Number of rows", min_value=1, step=1, value=23)
    cols = st.number_input("Number of columns", min_value=1, step=1, value=15)
    cell_width_cm = st.number_input("Width of each cell (cm)", min_value=0.1, step=0.1, value=4.1)  # Set input in cm
    cell_height_cm = st.number_input("Height of each cell (cm)", min_value=0.1, step=0.1, value=1.78)  # Set input in cm
    bol_diameter_pt = st.number_input("Diameter of the bullet (pt)", min_value=1.0, step=0.5, value=9.0)  # Set input in pt

    if st.button("Process Presentation"):
        output_ppt_path = process_presentation("uploaded_presentation.pptx", rows, cols, cell_width_cm, cell_height_cm, bol_diameter_pt)
        st.success("Presentation processed successfully!")
        with open(output_ppt_path, "rb") as f:
            st.download_button("Download Updated PowerPoint", f, "updated_presentation_departments.pptx", "application/vnd.ms-powerpoint")

