Attribute VB_Name = "Reload"

Sub Reload_module()
    Reload_module_shortcut_delete
    ThisWorkbook.load_from_conf ".\..\..\ÉVÉXÉeÉÄê›íË\libdef.txt"
End Sub

Private Sub Reload_module_shortcut_delete()
    For Each component In ThisWorkbook.VBProject.VBComponents
        If InStr(component.Name, "Reload") <> 0 Then
            Application.MacroOptions Macro:=component.Name & ".Reload_module", ShortcutKey:=""
        End If
    Next component
End Sub