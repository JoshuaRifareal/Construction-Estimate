Attribute VB_Name = "ChipDropdown"

Sub ColFootMixChip_Click()
    HideDropdown "ColFoot Mix Chip", "ColFoot Mix Options"
    PositionDropdown "ColFoot Mix Chip", "ColFoot Mix Options", "OptionsLeave"
    PositionOptions "ColFoot Mix Chip","Option 1", "Option1Hover"
    PositionOptions "ColFoot Mix Chip","Option 2", "Option2Hover", "Option 1"
    PositionOptions "ColFoot Mix Chip","Option 3", "Option3Hover", "Option 2"
    PositionOptions "ColFoot Mix Chip","Option 4", "Option4Hover", "Option 3"
End Sub


