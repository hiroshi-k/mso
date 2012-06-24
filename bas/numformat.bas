Attribute VB_Name = "NumberFormat"

Sub マイナス▲黒()
    Selection.NumberFormatLocal = "#,##0;""▲""#,##0"
End Sub

Sub マイナス▲赤()
    Selection.NumberFormatLocal = "#,##0;[赤]""▲""#,##0"
End Sub
