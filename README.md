# PowerPointSlideCutter
Cuts Power Point many Slides into Seperate Peresntations. Given there is a power point with multiple slides in it. We want to have a powerpoint for each slide containing only that slide. 

Acording to Chatti (https://chatgpt.com/share/67cf3047-3c4c-8002-a416-62495ed32df3)
Best Solution: Using VBA to Duplicate Slides


Steps:

- Open PowerPoint and press ALT + F11 to open the VBA editor.
- Paste this VBA code in a module:

"""
Sub SplitSlidesIntoSeparatePresentations()
    Dim oPres As Presentation
    Dim oNewPres As Presentation
    Dim oSlide As Slide
    Dim i As Integer
    Dim outputFolder As String
    
    ' Set active presentation
    Set oPres = ActivePresentation
    
    ' Define output folder
    outputFolder = oPres.Path & "\Split_Slides\"
    MkDir outputFolder
    
    ' Loop through slides
    For i = 1 To oPres.Slides.Count
        ' Create new presentation
        Set oNewPres = Presentations.Add
        oNewPres.PageSetup.SlideWidth = oPres.PageSetup.SlideWidth
        oNewPres.PageSetup.SlideHeight = oPres.PageSetup.SlideHeight
        
        ' Duplicate slide
        oPres.Slides(i).Copy
        oNewPres.Slides.Paste
        
        ' Save as new PowerPoint file
        oNewPres.SaveAs outputFolder & "Slide_" & i & ".pptx"
        oNewPres.Close
    Next i
    
    MsgBox "Slides saved in " & outputFolder, vbInformation, "Done!"
End Sub
"""
- Run the script (Press F5 in VBA).
- PowerPoint will create a folder with each slide as a separate PowerPoint file.







📌 Steps to Paste and Run the VBA Code in PowerPoint

1️⃣ Open PowerPoint and load your presentation.
2️⃣ Open the VBA Editor:

    Press ALT + F11 to open the Microsoft Visual Basic for Applications (VBA) Editor.
    If that doesn't work, go to Developer > Visual Basic (if the Developer tab is enabled).

3️⃣ Insert a New Module:

    In the VBA editor, click on "Insert" in the top menu.
    Select "Module" from the dropdown.
    A new blank window will appear on the right side.

4️⃣ Paste the Code:

    Copy the VBA script I provided.
    Paste it into the blank module window.

5️⃣ Run the Script:

    Click inside the code window.
    Press F5 or click Run > Run Sub/UserForm in the top menu.
    The script will now run, creating separate PowerPoint files.

6️⃣ Check the Output Folder:

    The script will create a folder called "Split_Slides" in the same location as your PowerPoint file.
    Inside that folder, you'll find each slide saved as a separate .pptx file.



