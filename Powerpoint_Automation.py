Created on Thu Sep 26 10:48:08 2024

@author: bramhendriksz
"""

import streamlit as st
import pythoncom
import win32com.client as win32
import os

# Functie om de macro aan de PowerPoint toe te voegen en uit te voeren
def add_and_run_macro(ppt_path, rows, cols, cell_width, cell_height):
    pythoncom.CoInitialize()
    ppt_app = win32.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True  # Zorg ervoor dat PowerPoint zichtbaar is
    
    # Open de presentatie
    presentation = ppt_app.Presentations.Open(ppt_path)

    # VBA-code toevoegen aan de presentatie
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
        bolDiameter = {9}

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

        ' Find the slide with the table that contains the 'Afkorting' column and store the text underneath
        Dim sldWithAfkorting As Slide
        Dim shpWithAfkortingTable As Shape
        Dim afkortingTbl As Table

        ' Go through each slide to find the Abbreviation table
        For Each sldWithAfkorting In ActivePresentation.Slides
            For Each shpWithAfkortingTable In sldWithAfkorting.Shapes
                If shpWithAfkortingTable.HasTable Then
                    Set afkortingTbl = shpWithAfkortingTable.Table
                    ' Determine which column is 'AFKORTING' and 'TITEL'
                    titleColumnIndex = 0  ' Reset the index for safety
                    For i = 1 To afkortingTbl.Columns.Count
                        If UCase(afkortingTbl.Cell(1, i).Shape.TextFrame.TextRange.Text) = "AFKORTING" Then
                            For j = 2 To afkortingTbl.Rows.Count
                                abbreviationText = afkortingTbl.Cell(j, i).Shape.TextFrame.TextRange.Text
                                For k = 1 To afkortingTbl.Columns.Count
                                    If UCase(afkortingTbl.Cell(1, k).Shape.TextFrame.TextRange.Text) = "SPF" Then
                                        location = afkortingTbl.Cell(j, k).Shape.TextFrame.TextRange.Text
                                        If location <> "" Then
                                            colLetter = Left(location, 1)
                                            rowNumber = Val(Mid(location, 2))
                                            colNumber = Asc(UCase(colLetter)) - Asc("A") + 1

                                            ' Increment the bullet count for this cell
                                            bulletCount(rowNumber, colNumber) = bulletCount(rowNumber, colNumber) + 1

                                            ' Calculate the position of the bullet based on its index
                                            Dim bulletIndex As Integer
                                            bulletIndex = (bulletCount(rowNumber, colNumber) - 1) Mod 5 + 1

                                            Dim colOffset As Single, rowOffset As Single
                                            colOffset = (bulletIndex - 1) * bolDiameter
                                            rowOffset = 0 ' All bullets are in the same row, so rowOffset is always 0

                                            xPos = tblShape.Left + (colNumber - 1) * (tblShape.Width / colsCount) + colOffset
                                            yPos = tblShape.Top + (rowNumber - 1) * (tblShape.Height / rowsCount) + rowOffset

                                            ' Add the bullet shape
                                            Set bolShape = sld.Shapes.AddShape(msoShapeOval, xPos, yPos, bolDiameter, bolDiameter)
                                            bolShape.TextFrame.TextRange.Text = abbreviationText
                                            bolShape.Fill.ForeColor.RGB = RGB(0, 0, 0) ' Black bullet
                                            bolShape.ActionSettings(ppMouseClick).Action = ppActionHyperlink
                                            bolShape.ActionSettings(ppMouseClick).Hyperlink.Address = "#" & sldWithAfkorting.SlideIndex
                                        End If
                                    ElseIf UCase(afkortingTbl.Cell(1, k).Shape.TextFrame.TextRange.Text) = "TITEL" Then
                                        titleColumnIndex = k
                                    End If
                                Next k
                                If titleColumnIndex > 0 Then
                                    projectTitle = afkortingTbl.Cell(j, titleColumnIndex).Shape.TextFrame.TextRange.Text
                                    If Not bolShape Is Nothing Then
                                        bolShape.ActionSettings(ppMouseClick).Hyperlink.ScreenTip = abbreviationText & " : " & projectTitle
                                    End If
                                End If
                            Next j
                            Exit For
                        End If
                    Next i
                    Exit For
                End If
            Next shpWithAfkortingTable
        Next sldWithAfkorting

        ' Change the border color of every cell in the table to black
        Dim newTbl As Table
        Set newTbl = tblShape.Table
        For i = 1 To newTbl.Rows.Count
            For j = 1 To newTbl.Columns.Count
                With newTbl.Cell(i, j).Borders
                    .Item(ppBorderTop).ForeColor.RGB = RGB(0, 0, 0)
                    .Item(ppBorderLeft).ForeColor.RGB = RGB(0, 0, 0)
                    .Item(ppBorderBottom).ForeColor.RGB = RGB(0, 0, 0)
                    .Item(ppBorderRight).ForeColor.RGB = RGB(0, 0, 0)
                End With
            Next j
        Next i
    End Sub
    """

    # Voeg de VBA-code toe aan de presentatie
    ppt_app.VBE.ActiveVBProject.VBComponents.Add(win32.constants.vbext_ct_StdModule).CodeModule.AddFromString(vba_code)
    
    # Voer de macro uit
    ppt_app.Run("AddBulletsAndNumberingBasedOnTableLocationsWithHyperlinks")

    # Sla de gewijzigde presentatie op
    output_path = os.path.join(os.path.dirname(ppt_path), "updated_presentation.pptm")
    presentation.SaveAs(output_path)
    presentation.Close()
    ppt_app.Quit()

    return output_path

# Streamlit UI opzetten
st.title("PowerPoint Macro Tool")

# Upload PowerPoint bestand
uploaded_ppt = st.file_uploader("Upload je PowerPoint presentatie", type=["pptx"])

if uploaded_ppt is not None:
    # Sla het geüploade bestand tijdelijk op
    with open("uploaded_presentation.pptx", "wb") as f:
        f.write(uploaded_ppt.read())
    
    st.success("PowerPoint succesvol geüpload!")

    # Vraag om invoer voor de variabelen
    rows = st.number_input("Aantal rijen", min_value=1, step=1, value=23)
    cols = st.number_input("Aantal kolommen", min_value=1, step=1, value=15)
    cell_width = st.number_input("Breedte van de cel (cm)", min_value=0.1, step=0.1, value=1.62)
    cell_height = st.number_input("Hoogte van de cel (cm)", min_value=0.1, step=0.1, value=0.7)
    
    if st.button("Voer macro uit"):
        # Voeg de macro toe en voer deze uit met de opgegeven parameters
        output_ppt_path = add_and_run_macro("uploaded_presentation.pptx", rows, cols, cell_width, cell_height)
        
        st.success("Macro succesvol uitgevoerd!")
        
        # Bied de geüpdatete presentatie aan voor download
        with open(output_ppt_path, "rb") as f:
            st.download_button(
                label="Download aangepaste PowerPoint",
                data=f,
                file_name="updated_presentation.pptm",
                mime="application/vnd.ms-powerpoint"
            )
