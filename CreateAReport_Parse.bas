Attribute VB_Name = "CreateAReport_Parse"
Sub Enroll_Date()
Attribute Enroll_Date.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    Dim LRow As Long

    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    Lastrow = Range("A" & Rows.Count).End(xlUp).Row 'finds last row index
    
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Enrollment Date"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=DATEVALUE(LEFT(RC[1],10))"
    Range("M2").AutoFill Destination:=Range("M2:M" & Lastrow)
    
    Columns("M:M").Select
    Selection.NumberFormat = "m/d/yyyy"
    
    Range("A1").Select
    Selection.AutoFilter
    
End Sub
