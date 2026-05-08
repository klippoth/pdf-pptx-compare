Attribute VB_Name = "SlideExport"
Option Explicit

Private Const POSIX_PATH_SEPARATOR As String = "/"

Public Sub InsertReferenceSlidesFromManifest(ByVal manifestPath As String)
    Dim presentation As Presentation
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim fileHandle As Integer
    Dim lineText As String
    Dim columns() As String
    Dim actionType As String
    Dim imagePath As String
    Dim insertAfterIndex As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    If Len(Trim$(manifestPath)) = 0 Then
        Err.Raise vbObjectError + 1000, "SlideExport", "Manifest path is required."
    End If

    If Dir$(manifestPath) = vbNullString Then
        Err.Raise vbObjectError + 1001, "SlideExport", "Manifest file was not found."
    End If

    Set presentation = ActivePresentation
    If presentation Is Nothing Then
        Err.Raise vbObjectError + 1002, "SlideExport", "No active presentation is open."
    End If

    slideWidth = presentation.PageSetup.SlideWidth
    slideHeight = presentation.PageSetup.SlideHeight

    fileHandle = FreeFile
    Open manifestPath For Input As #fileHandle
    On Error GoTo CleanFail

    Do While Not EOF(fileHandle)
        Line Input #fileHandle, lineText

        If Len(Trim$(lineText)) = 0 Then
            GoTo ContinueLoop
        End If

        columns = Split(lineText, vbTab)
        If UBound(columns) < 2 Then
            Err.Raise vbObjectError + 1003, "SlideExport", "Invalid manifest line: " & lineText
        End If

        actionType = UCase$(Trim$(columns(0)))
        imagePath = Trim$(columns(2))

        If actionType = "INSERT" Then
            insertAfterIndex = CLng(Trim$(columns(1)))
            InsertReferenceSlide presentation, insertAfterIndex + 2, imagePath, slideWidth, slideHeight
        ElseIf actionType = "APPEND" Then
            InsertReferenceSlide presentation, presentation.Slides.Count + 1, imagePath, slideWidth, slideHeight
        Else
            Err.Raise vbObjectError + 1004, "SlideExport", "Unknown manifest action: " & actionType
        End If

ContinueLoop:
    Loop

CleanExit:
    On Error Resume Next
    If fileHandle <> 0 Then
        Close #fileHandle
    End If
    Exit Sub

CleanFail:
    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next
    If fileHandle <> 0 Then
        Close #fileHandle
    End If
    Err.Raise errorNumber, "SlideExport", errorDescription
End Sub

Public Sub TestInsertReferenceSlidesFromManifest()
    Dim manifestPath As String

    manifestPath = InputBox( _
        Prompt:="Enter the full POSIX path to a manifest TSV file.", _
        Title:="Test InsertReferenceSlidesFromManifest", _
        Default:=DefaultFolderPath("reference-slide-manifest.tsv") _
    )

    If Len(Trim$(manifestPath)) = 0 Then
        Exit Sub
    End If

    InsertReferenceSlidesFromManifest manifestPath
    MsgBox "Reference slide insertion finished.", vbInformation, "SlideExport"
End Sub

Private Sub InsertReferenceSlide( _
    ByVal presentation As Presentation, _
    ByVal slideIndex As Long, _
    ByVal imagePath As String, _
    ByVal slideWidth As Single, _
    ByVal slideHeight As Single)

    Dim slideItem As Slide
    Dim picture As Shape

    Set slideItem = presentation.Slides.Add(slideIndex, ppLayoutBlank)
    slideItem.Name = NextReferenceSlideName(presentation)

    Set picture = slideItem.Shapes.AddPicture(imagePath, msoFalse, msoTrue, 0, 0, slideWidth, slideHeight)
    picture.Name = "PDF_ORIGINAL"
End Sub

Private Function NextReferenceSlideName(ByVal presentation As Presentation) As String
    Dim candidate As String
    Dim suffix As Long

    suffix = 1
    Do
        candidate = "PDF_ORIGINAL_" & Format$(suffix, "000")
        If Not SlideNameExists(presentation, candidate) Then
            NextReferenceSlideName = candidate
            Exit Function
        End If
        suffix = suffix + 1
    Loop
End Function

Private Function SlideNameExists(ByVal presentation As Presentation, ByVal slideName As String) As Boolean
    Dim slideItem As Slide

    For Each slideItem In presentation.Slides
        If slideItem.Name = slideName Then
            SlideNameExists = True
            Exit Function
        End If
    Next slideItem
End Function

Private Function DefaultFolderPath(ByVal leafName As String) As String
    Dim homePath As String

    homePath = Environ$("HOME")
    If Len(homePath) = 0 Then
        DefaultFolderPath = POSIX_PATH_SEPARATOR & "tmp" & POSIX_PATH_SEPARATOR & leafName
        Exit Function
    End If

    If Right$(homePath, 1) = POSIX_PATH_SEPARATOR Then
        DefaultFolderPath = homePath & "Downloads" & POSIX_PATH_SEPARATOR & leafName
    Else
        DefaultFolderPath = homePath & POSIX_PATH_SEPARATOR & "Downloads" & POSIX_PATH_SEPARATOR & leafName
    End If
End Function
