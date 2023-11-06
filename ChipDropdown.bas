Attribute VB_Name = "ChipDropdown"
Sub ColFootMixChip_Click()
    HideDropdown "ColFoot Mix Dropdown"
End Sub

Sub HideDropdown(dropdownName As String)
    If ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse Then
        ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoTrue
    Else
        ActiveSheet.Shapes.Range(Array(dropdownName)).Visible = msoFalse
    End If
End Sub
