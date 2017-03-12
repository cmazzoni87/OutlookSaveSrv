''Allows user to save read-only excel files (in this case a xlam) using XCOPY and archives previous version

Option Explicit
Sub SaveCopy():
    Dim aWkb As Workbook, bkpWkb As Workbook
    Dim strDt As String, bkpStr As String, plyStrng As String
    
    Set aWkb = ThisWorkbook
    strDt = Format(Now(), "YYYY")
    bkpStr = aWkb.Path & "\Backup_" & strDt
    If Dir(bkpStr, vbDirectory) = "" Then
        MkDir bkpStr
    End If
    If Dir(bkpStr & "\" & Format(Now(), "YYYYMMDD"), vbDirectory) = "" Then
        MkDir bkpStr & "\" & Format(Now(), "YYYYMMDD")
        plyStrng = bkpStr & "\" & Format(Now(), "YYYYMMDD")
        aWkb.SaveCopyAs plyStrng & "\" & Left(aWkb.Name, 8) & "_" & Format(Now(), "hhmmss") & ".xlam"
    Else
        plyStrng = bkpStr & "\" & Format(Now(), "YYYYMMDD")
        aWkb.SaveCopyAs plyStrng & "\" & Left(aWkb.Name, 8) & "_" & Format(Now(), "hhmmss") & ".xlam"
    End If
    Call SetAttr(plyStrng & "\" & Left(aWkb.Name, 8) & "_" & Format(Now(), "hhmmss") & ".xlam", vbReadOnly)
End Sub

Sub moveToProd()
    Dim aWkb As Workbook, bkpWkb As Workbook
    Dim strDt As String, bkpStr As String, plyStrng As String
    Dim fso As Object
    
    Set aWkb = ThisWorkbook
    If aWkb.ReadOnly = False Then
        MsgBox "This file is not Read Only, please save the file normally"
        Exit Sub
    End If
    strDt = ThisWorkbook.Path & "\Older_Versions"
    If Dir(strDt, vbDirectory) = "" Then
        MkDir strDt
    End If
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(aWkb.FullName, strDt & "\" & Left(aWkb.Name, 8) & "_" & Format(Now(), "YYYYMMDDhhmmss") & ".xlam")
    plyStrng = Format(Now(), "hhmmss") & ".xlam"
    aWkb.SaveCopyAs aWkb.Path & "\xCOPY" & plyStrng
    strDt = aWkb.Path & "\xCOPY" & plyStrng
    Call SetAttr(strDt, vbReadOnly)
    Shell "xcopy " & strDt & " " & ThisWorkbook.FullName & " /r /y"
    Application.Wait (Now + TimeValue("0:00:02"))
    Call SetAttr(strDt, vbNormal)
    Kill strDt
End Sub



