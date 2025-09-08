VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim lastCol As Long, c As Long
    Dim shpName As String, shp As Shape

    Application.EnableEvents = False

    ' Only reset prevShape if it exists
    If Not prevShape Is Nothing Then
        On Error Resume Next
        prevShape.Fill.Transparency = 1
        prevShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
        On Error GoTo 0
        Set prevShape = Nothing
    End If

    ' Clear all row highlights
    Call ClearAllHighlights

    ' If clicked outside populated rows, exit
    If Target.CountLarge > 1 Or IsEmpty(Target.Value) Or Target.Row = 1 _
        Or LCase(Cells(Target.Row, 1).Value) = "totals" Then
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Highlight row (only populated columns)
    lastCol = Cells(Target.Row, Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Not IsEmpty(Cells(Target.Row, c)) Then
            Cells(Target.Row, c).Interior.Color = RGB(253, 191, 86)
        End If
    Next c

    ' Highlight corresponding shape if it exists
    shpName = "Val_" & Cells(Target.Row, 1).Value
    On Error Resume Next
    Set shp = Me.Shapes(shpName)
    If Not shp Is Nothing Then
        shp.Fill.Transparency = 0.6
        shp.Fill.ForeColor.RGB = RGB(253, 191, 86)
        Set prevShape = shp
    End If
    On Error GoTo 0

    ' Store parcel
    prevParcel = Cells(Target.Row, 1).Value

    
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Calculate()
    Dim r As Range
    Dim lastCol As Long, c As Long
    Dim shp As Shape

    If prevParcel = 0 Then Exit Sub

    ' Find row of stored parcel after sort
    For Each r In ActiveSheet.Range("A2:A" & ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row)
        If r.Value = prevParcel Then
            ' Highlight row
            lastCol = Cells(r.Row, Columns.Count).End(xlToLeft).Column
            For c = 1 To lastCol
                If Not IsEmpty(Cells(r.Row, c)) Then
                    Cells(r.Row, c).Interior.Color = RGB(253, 191, 86)
                End If
            Next c

            ' Highlight shape
            On Error Resume Next
            Set shp = Me.Shapes("Val_" & r.Value)
            If Not shp Is Nothing Then
                shp.Fill.Transparency = 0.6
                shp.Fill.ForeColor.RGB = RGB(253, 191, 86)
                Set prevShape = shp
            Else
                Set prevShape = Nothing
            End If
            On Error GoTo 0
                 
            Exit For
        End If
    Next r
End Sub




