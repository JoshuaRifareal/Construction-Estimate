Attribute VB_Name = "ChipDropdown"

Sub ColFootMixChip_Click()
    HideDropdown "ColFoot Mix Chip", "ColFoot Mix Options"
    PositionDropdown "ColFoot Mix Chip", "ColFoot Mix Options", "ColFoot Mix Option 1", "ColFoot Mix Panel", "ColFoot_Mix_OptionsLeave"
    PositionOptions "ColFoot Mix Chip", "ColFoot Mix Option 1", "ColFoot_Mix_Option1Hover"
    PositionOptions "ColFoot Mix Chip", "ColFoot Mix Option 2", "ColFoot_Mix_Option2Hover", "ColFoot Mix Option 1"
    PositionOptions "ColFoot Mix Chip", "ColFoot Mix Option 3", "ColFoot_Mix_Option3Hover", "ColFoot Mix Option 2"
    PositionOptions "ColFoot Mix Chip", "ColFoot Mix Option 4", "ColFoot_Mix_Option4Hover", "ColFoot Mix Option 3"
End Sub


