Attribute VB_Name = "syrup_mod_1"
Public prevParcel As Long
Public prevShape As Shape

Sub HighlightRowDynamic()
    Dim shp As Shape
    Dim rowNum As Long
    Dim lastCol As Long
    Dim c As Long
    Dim r As Range
    Dim selCol As Long
    
    
    Application.EnableEvents = False
    
    ' Store the currently selected column
    selCol = ActiveCell.Column
        
    ' Clear previous highlights
    Call ClearAllHighlights
    
    ' Current shape
    Set shp = ActiveSheet.Shapes(Application.Caller)
    shp.Fill.Transparency = 0.6
    shp.Fill.ForeColor.RGB = RGB(253, 191, 86)
    Set prevShape = shp
    
    ' Get parcel number from shape
    Dim parcelNum As Long
    parcelNum = CLng(Replace(shp.Name, "Val_", ""))
    
    ' Find the row for that parcel
    For Each r In ActiveSheet.Range("A2:A" & ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row)
        If r.Value = parcelNum Then
            rowNum = r.Row
            Exit For
        End If
    Next r
    
    ' Highlight populated columns in the row
    lastCol = Cells(rowNum, Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Not IsEmpty(Cells(rowNum, c)) Then
            Cells(rowNum, c).Interior.Color = RGB(253, 191, 86)
        End If
    Next c
    
    ' Store parcel for clearing next time
    prevParcel = parcelNum
    
    ' After highlighting the row via shape
    Cells(rowNum, selCol).Select
    

    Application.EnableEvents = True
End Sub

Sub ClearAllHighlights()
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    lastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    
    For r = 2 To lastRow  ' skip header
        Range(Cells(r, 1), Cells(r, lastCol)).Interior.ColorIndex = xlNone
    Next r
    
    ' Reset previous shape
    If Not prevShape Is Nothing Then
        On Error Resume Next
        prevShape.Fill.Transparency = 1
        prevShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
        On Error GoTo 0
        Set prevShape = Nothing
    End If
End Sub

