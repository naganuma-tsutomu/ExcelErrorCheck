Attribute VB_Name = "Reload"

Sub Reload_module()
    Reload_module_shortcut_delete
    ThisWorkbook.load_from_conf ".\..\..\src\libdef.txt"
End Sub

Private Sub Reload_module_shortcut_delete()
    For Each component In ThisWorkbook.VBProject.VBComponents
        If InStr(component.Name, "Reload") <> 0 Then
            Application.MacroOptions Macro:=component.Name & ".Reload_module", ShortcutKey:=""
        End If
    Next component
End Sub

Sub AllReplace()
    
    Dim f       As String
    Dim t       As String
    Dim i       As Integer
    Dim sht     As Excel.Worksheet
    
    f = InputBox("’uŠ·‘ÎÛ•¶š—ñ‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢")
    
    If Trim$(f) <> "" Then
        t = InputBox("’uŠ·Œã•¶š—ñ‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢")
    End If

    If Trim$(t) = "" Then
        Exit Sub
    End If
    
    For i = 1 To ActiveWorkbook.Sheets.Count
        Set sht = ActiveWorkbook.Sheets(i)
        
        If sht.Name = f Then
            sht.Name = t
        End If

        sht.Cells.Replace What:=f, Replacement:=t, LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, MatchByte:=False
    Next

    Call MsgBox("Š®—¹")
End Sub 