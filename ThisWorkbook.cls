VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Worksheets("Sheet1").Range("B3").value = "'" + Format(Now(), "yyyy/mm/dd")
    ' 起動の時に、PI Serverに値を書き込みたい場合は、
    ' 下記の行のコメントのマークを削除してください。
    'WriteDataToPI_Click
End Sub

