Option Explicit
'Author Ejaskhan
'http://msdn.microsoft.com/en-us/library/office/bb251061(v=office.12).aspx
Const ppSaveAsPDF = 32

Dim oFSO ' Public reference to FileSystemObject
Dim oPPT ' Public reference to PowerPoint App

Main

Sub Main()
    Dim sInput

    If wscript.Arguments.Count <> 1 Then
        Wscript.Echo "You need to specify input and output files."
        wscript.Quit
    End If

    Set oPPT = CreateObject("PowerPoint.Application")
    If CDbl(oPPT.Version) < 12 Then
        Wscript.Echo "PowerPoint version must be 2007 or later!"
        oPPT.Visible = True
        oPPT.Quit
        Set oPPT = Nothing
        wscript.Quit
    End If
    ' Store Input Argument and detect execute mode (single file / Folder batch mode)
    sInput = wscript.Arguments(0)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If IsPptFile(sInput) Then
        PPT2PDF sInput
    ElseIf oFSO.FolderExists(sInput) Then
		Wscript.Echo "Hey Blum, file(s) creation started at: " & Now
        BatchPPT2PDF sInput
    Else
        Wscript.Echo """" & sInput & """ is not a PPT file or Folder!"
    End If
    ' Close PowerPoint app if no other presentations are opened
    If oPPT.Presentations.Count = 0 Then oPPT.Quit
    Set oPPT = Nothing
    Set oFSO = Nothing
End Sub

Private Sub BatchPPT2PDF(sFDR)
    Dim oFDR, oFile
    'Wscript.Echo String(50, Chr(151))
    For Each oFile In oFSO.GetFolder(sFDR).Files
        If IsPptFile(oFile.Name) Then
            PPT2PDF(oFile)
        End If
    Next
    For Each oFDR In oFSO.GetFolder(sFDR).SubFolders
        BatchPPT2PDF oFDR
    Next
	Wscript.Echo "File(s) got generated!!! Please check the output folder, " & sFDR
	Wscript.Echo "Thank you BLUM, @EK"
End Sub

Private Function IsPptFile(sFile)
    IsPptFile = (InStr(1, Right(sFile, InStrRev(sFile, ".")), "ppt") > 0)
End Function

Private Sub PPT2PDF(sFile)
    On Error Resume Next
    Dim sPDF, oPres
    sPDF = Left(sFile,InstrRev(sFile,".")) & "pdf"
    Set oPres = oPPT.Presentations.Open(sFile, True, False, False) ' Read Only, No Title, No Window
    Err.Clear
    oPres.SaveAs sPDF, ppSaveAsPDF
    oPres.Close
    Set oPres = Nothing
    If Err.Number = 0 Then
        
    Else
        Wscript.Echo "X" & vbTab & sPDF & " [ERR " & Err.Number & ": " & Err.Description & "]"
        Err.Clear
    End If
End Sub