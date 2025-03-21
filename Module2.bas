Attribute VB_Name = "Module2"
Sub SplitSlidesIntoSeparatePresentations()
    Dim oPres As Presentation
    Dim oNewPres As Presentation
    Dim oSlide As Slide
    Dim i As Integer
    Dim outputFolder As String
    Dim slideTitle As String
    Dim firstShape As Shape
    Dim existingTitles As Collection
    Dim titleCount As Integer
    Dim uniqueTitle As String
    
    ' Set active presentation
    Set oPres = ActivePresentation
    
    ' Define output folder
    outputFolder = oPres.Path & "\Split_Slides\"
    MkDir outputFolder
    
    ' Initialize the collection for existing titles
    Set existingTitles = New Collection
    
    ' Loop through slides
    For i = 1 To oPres.Slides.Count
        ' Create new presentation
        Set oNewPres = Presentations.Add
        oNewPres.PageSetup.SlideWidth = oPres.PageSetup.SlideWidth
        oNewPres.PageSetup.SlideHeight = oPres.PageSetup.SlideHeight
        
        ' Duplicate slide
        oPres.Slides(i).Copy
        oNewPres.Slides.Paste
        
        ' Initialize slideTitle
        slideTitle = ""
        
        ' Check if there are shapes on the slide
        If oPres.Slides(i).Shapes.Count > 0 Then
            Dim shapeFound As Boolean
            shapeFound = False
            
            ' Loop through all shapes on the slide
            Dim j As Integer
            For j = 1 To oPres.Slides(i).Shapes.Count
                Set firstShape = oPres.Slides(i).Shapes(j)
                Debug.Print firstShape.Name
                
                ' Check if the shape has a text frame and contains text
                If firstShape.HasTextFrame Then
                    If firstShape.TextFrame.HasText Then
                        slideTitle = firstShape.TextFrame.TextRange.Text
                        shapeFound = True
                        Debug.Print "Slide Title: " & slideTitle
                        Exit For ' Exit the loop once the first text-containing shape is found
                    End If
                End If
            Next j
            
            If Not shapeFound Then
                Debug.Print "No text found in any shape on slide " & i
            End If
        End If
        
        ' Check if slideTitle is empty and assign a default name if necessary
        If slideTitle = "" Then
            slideTitle = "Slide_" & i
        End If
        
        ' Clean the title to create a valid filename
        slideTitle = CleanFileName(slideTitle)
        
        ' Check for existing titles and create a unique title
        uniqueTitle = slideTitle
        titleCount = 1
        
        ' Loop to find a unique title
        On Error Resume Next ' Ignore errors for duplicate keys
        Do While True
            existingTitles.Add uniqueTitle, uniqueTitle ' Try to add the title
            If Err.Number = 0 Then
                ' No error means the title is unique
                Exit Do
            Else
                ' Title already exists, increment the count and create a new title
                titleCount = titleCount + 1
                uniqueTitle = slideTitle & "_" & titleCount
                Err.Clear ' Clear the error for the next iteration
            End If
        Loop
        On Error GoTo 0 ' Resume normal error handling
        
        ' Print the unique title to the Immediate Window
        Debug.Print "Saving slide " & i & " as: " & uniqueTitle
        
        ' Save as new PowerPoint file
        oNewPres.SaveAs outputFolder & uniqueTitle & ".pptx"
        oNewPres.Close
    Next i
    
    MsgBox "Slides saved in " & outputFolder, vbInformation, "Done!"
End Sub

Function CleanFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    
    ' Define invalid characters for file names
    invalidChars = "/\:*?""<>|"
    
    ' Replace invalid characters with an underscore
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i
    
    ' Return cleaned file name
    CleanFileName = Trim(fileName)
End Function
