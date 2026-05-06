Attribute VB_Name = "SlideExport"
Option Explicit

Public Function ExportSlidesAsPng( _
    ByVal presentationPath As String, _
    ByVal outputFolder As String, _
    Optional ByVal scaleWidth As Long = 1400, _
    Optional ByVal scaleHeight As Long = 788) As String

    Dim presentation As Presentation
    Dim targetFolder As String

    targetFolder = outputFolder
    If Len(targetFolder) = 0 Then
        Err.Raise vbObjectError + 1000, "SlideExport", "Output folder is required."
    End If

    If Right$(targetFolder, 1) = Application.PathSeparator Then
        targetFolder = Left$(targetFolder, Len(targetFolder) - 1)
    End If

    EnsureFolderExists targetFolder

    Set presentation = Presentations.Open(presentationPath, msoFalse, msoFalse, msoFalse)
    On Error GoTo CleanFail

    presentation.Export Path:=targetFolder, FilterName:="PNG", ScaleWidth:=scaleWidth, ScaleHeight:=scaleHeight
    ExportSlidesAsPng = targetFolder

CleanExit:
    On Error Resume Next
    If Not presentation Is Nothing Then
        presentation.Close
    End If
    Exit Function

CleanFail:
    ExportSlidesAsPng = "ERROR: " & Err.Number & " - " & Err.Description
    Resume CleanExit
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then
        Exit Sub
    End If

    MkDirRecursive folderPath
End Sub

Private Sub MkDirRecursive(ByVal folderPath As String)
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Long

    pathParts = Split(folderPath, Application.PathSeparator)
    If InStr(folderPath, ":") > 0 And Left$(folderPath, 1) <> Application.PathSeparator Then
        currentPath = pathParts(0)
        i = 1
    Else
        currentPath = Application.PathSeparator
        i = 1
    End If

    For i = i To UBound(pathParts)
        If Len(pathParts(i)) = 0 Then
            GoTo ContinueLoop
        End If

        If Right$(currentPath, 1) = Application.PathSeparator Then
            currentPath = currentPath & pathParts(i)
        Else
            currentPath = currentPath & Application.PathSeparator & pathParts(i)
        End If

        If Len(Dir$(currentPath, vbDirectory)) = 0 Then
            MkDir currentPath
        End If

ContinueLoop:
    Next i
End Sub
