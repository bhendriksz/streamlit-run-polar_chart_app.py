import streamlit as st
import pythoncom
import win32com.client as win32
import os

# Function to add and run the macro in the PowerPoint presentation
def add_and_run_macro(ppt_path, rows, cols, cell_width, cell_height):
    pythoncom.CoInitialize()
    
    try:
        ppt_app = win32.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True  # Make sure PowerPoint is visible for debugging purposes

        # Open the presentation
        ppt_path = os.path.abspath(ppt_path)  # Ensure the path is absolute
        if not os.path.exists(ppt_path):
            raise FileNotFoundError(f"The PowerPoint file was not found at: {ppt_path}")

        presentation = ppt_app.Presentations.Open(ppt_path)

        # VBA code to be added to the presentation
        vba_code = f"""
        Sub AddBulletsAndNumberingBasedOnTableLocationsWithHyperlinks()
            Dim sld As Slide
            Dim tblShape As Shape
            Dim rowsCount As Integer
            Dim colsCount As Integer
            Dim bolDiameter As Single
            Dim bolShape As Shape
            Dim location As String
            Dim colLetter As String
            Dim rowNumber As Integer
            Dim colNumber As Integer
            Dim xPos As Single, yPos As Single
            Dim abbreviationText As String
            Dim projectTitle As String
            Dim titleColumnIndex As Integer
            Dim bulletCount() As Integer
            
            ' Set the number of rows and columns based on your specific image
            rowsCount = {rows}
            colsCount = {cols}

            ' Initialize the bullet count array
            ReDim bulletCount(1 To rowsCount, 1 To colsCount)

            ' Set the diameter of the bullets
            bolDiameter = 9

            ' Refer to the active presentation and add a new slide
            Set sld = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)

            ' Add a table with the specified dimensions
            Set tblShape = sld.Shapes.AddTable(rowsCount, colsCount, 10, 10, {cell_width * 28.35 * cols}, {cell_height * 28.35 * rows})

            ' Set the table background to white
            tblShape.Fill.ForeColor.RGB = RGB(255, 255, 255)

            ' Adding numbering to the table based on chessboard notation
            For j = 1 To colsCount
                For i = 1 To rowsCount
                    tblShape.Table.Cell(i, j).Shape.TextFrame.TextRange.Text = Chr(j + 64) & i
                    tblShape.Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Size = 7
                    tblShape.Table.Columns(j).Width = Len(tblShape.Table.Cell(i, j).Shape.TextFrame.TextRange.Text) * 16.4
                Next i
            Next j

            ' Set the text color of the cells in the top row to black
            For j = 1 To colsCount
                tblShape.Table.Cell(1, j).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
            Next j
        End Sub
        """

        # Add the VBA code to the presentation
        module = presentation.VBProject.VBComponents.Add(win32.constants.vbext_ct_StdModule)
        module.CodeModule.AddFromString(vba_code)

        # Run the macro
        ppt_app.Run("AddBulletsAndNumberingBasedOnTableLocationsWithHyperlinks")

        # Save the modified presentation
        output_path = os.path.join(os.path.dirname(ppt_path), "updated_presentation.pptm")
        presentation.SaveAs(output_path)
        presentation.Close()
        ppt_app.Quit()

        return output_path

    except Exception as e:
        st.error(f"An error occurred: {e}")
        if ppt_app:
            ppt_app.Quit()
        raise e

# Set up the Streamlit UI
st.title("PowerPoint Macro Tool")

# Upload PowerPoint file
uploaded_ppt = st.file_uploader("Upload your PowerPoint presentation", type=["pptx", "pptm"])

if uploaded_ppt is not None:
    # Save the uploaded file temporarily
    with open("uploaded_presentation.pptx", "wb") as f:
        f.write(uploaded_ppt.read())

    st.success("PowerPoint successfully uploaded!")

    # Input fields for variables
    rows = st.number_input("Number of rows", min_value=1, step=1, value=23)
    cols = st.number_input("Number of columns", min_value=1, step=1, value=15)
    cell_width = st.number_input("Width of the cell (cm)", min_value=0.1, step=0.1, value=1.62)
    cell_height = st.number_input("Height of the cell (cm)", min_value=0.1, step=0.1, value=0.7)
    
    if st.button("Run Macro"):
        # Add and run the macro with the provided parameters
        output_ppt_path = add_and_run_macro("uploaded_presentation.pptx", rows, cols, cell_width, cell_height)
        
        st.success("Macro executed successfully!")
        
        # Provide the updated presentation for download
        with open(output_ppt_path, "rb") as f:
            st.download_button(
                label="Download Updated PowerPoint",
                data=f,
                file_name="updated_presentation.pptm",
                mime="application/vnd.ms-powerpoint"
            )

