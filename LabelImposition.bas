Attribute VB_Name = "LabelImposition"
Option Explicit

' Import this source into a CorelDRAW VBA project. No binary GMS is supplied.
' DEMO-220: fictional machine, 220 mm web, 300 mm rapport, k=0.97.
' Python companion verifies the arithmetic, not execution of this VBA module.

Public Function MaxColumns(ByVal labelWidth As Double, ByVal webWidth As Double, _
    ByVal clearance As Double, ByVal railWidth As Double) As Long
    If labelWidth <= 0 Or webWidth <= 0 Or clearance < 0 Or railWidth <= 0 Then _
        Err.Raise 5, , "Invalid width or rail parameters"
    MaxColumns = Int((webWidth - 2 * (clearance + railWidth) + 0.00000001) / labelWidth)
End Function

Public Function MaxRows(ByVal labelHeight As Double, ByVal rapport As Double) As Long
    If labelHeight <= 0 Or rapport <= 0 Then Err.Raise 5, , "Invalid height or rapport"
    MaxRows = Int((rapport + 0.00000001) / labelHeight)
End Function

Private Function AskNumber(ByVal prompt As String, ByVal defaultValue As Double) As Double
    Dim answer As String
    answer = InputBox(prompt & vbCrLf & "Cancel aborts without generating a layout.", _
                      "Label imposition DEMO", CStr(defaultValue))
    If Len(Trim$(answer)) = 0 Then Err.Raise 18, , "Cancelled"
    If Not IsNumeric(answer) Then Err.Raise 5, , "A number is required"
    AskNumber = CDbl(answer)
End Function

Public Sub BuildDemoRapport()
    Dim doc As Document, source As Shape, outputPage As Page, outputLayer As Layer
    Dim dup As Shape, railShape As Shape, composition As Shape
    Dim allShapes As New ShapeRange
    Dim oldUnit As cdrUnit, oldReference As cdrReferencePoint, oldOptimization As Boolean
    Dim groupStarted As Boolean, stateCaptured As Boolean, errText As String, errNumber As Long
    Dim w As Double, h As Double, web As Double, rapport As Double, factor As Double
    Dim cols As Long, rows As Long, c As Long, r As Long
    Dim n As Double, leftX As Double, pitch As Double, railWidth As Double, clearance As Double
    Dim reverse As Boolean, mode As Double
    Dim originX As Double, topY As Double
    On Error GoTo Failed
    If Documents.Count = 0 Then Err.Raise 5, , "Open a COPY of a label document first."
    If ActiveSelectionRange.Count <> 1 Then Err.Raise 5, , "Select exactly one grouped label."
    Set source = ActiveSelectionRange(1)
    If source.Type <> cdrGroupShape Then Err.Raise 5, , "The selected label must be grouped."
    Set doc = ActiveDocument
    oldUnit = doc.Unit
    oldReference = doc.ReferencePoint
    oldOptimization = Application.Optimization
    stateCaptured = True
    doc.Unit = cdrMillimeter
    doc.ReferencePoint = cdrTopLeft
    w = source.SizeWidth: h = source.SizeHeight
    web = AskNumber("Web width (mm). DEMO machine = 220", 220)
    rapport = AskNumber("Rapport BEFORE distortion (mm). DEMO = 300", 300)
    factor = AskNumber("Length compression factor. DEMO ONLY = 0.97", 0.97)
    clearance = 1.5: railWidth = 2
    If factor <= 0 Or factor > 1 Then Err.Raise 5, , "Compression factor must be >0 and <=1."
    If MaxColumns(w, web, clearance, railWidth) < 1 Or MaxRows(h, rapport) < 1 Then _
        Err.Raise 5, , "The selected label does not fit the web or rapport."
    n = AskNumber("Columns (streams). Maximum = " & MaxColumns(w, web, clearance, railWidth), _
                  MaxColumns(w, web, clearance, railWidth))
    If n <> Fix(n) Or n < 1 Or n > MaxColumns(w, web, clearance, railWidth) Then _
        Err.Raise 5, , "Invalid number of columns"
    cols = CLng(n)
    n = AskNumber("Labels per stream. Maximum = " & MaxRows(h, rapport), MaxRows(h, rapport))
    If n <> Fix(n) Or n < 1 Or n > MaxRows(h, rapport) Then Err.Raise 5, , "Invalid number of rows"
    rows = CLng(n)
    If CDbl(cols) * rows > 10000 Then Err.Raise 5, , "Demo limit: 10000 labels"
    mode = AskNumber("Print mode: 0 = direct, 1 = mirror horizontally (confirm with printer)", 0)
    If mode <> 0 And mode <> 1 Then Err.Raise 5, , "Print mode must be 0 or 1"
    reverse = (mode = 1)
    leftX = (web - cols * w) / 2
    pitch = rapport / rows
    doc.BeginCommandGroup "DEMO label imposition"
    groupStarted = True
    Application.Optimization = True
    Set outputPage = doc.AddPages(1)
    outputPage.Activate
    outputPage.SetSize web, rapport * factor
    originX = outputPage.LeftX
    topY = outputPage.TopY
    Set outputLayer = outputPage.CreateLayer("DEMO layout - verify before printing")
    For r = 0 To rows - 1
        For c = 0 To cols - 1
            Set dup = source.Duplicate(0, 0)
            dup.MoveToLayer outputLayer
            dup.SetPositionEx cdrTopLeft, originX + leftX + c * w, topY - r * pitch
            allShapes.Add dup
        Next c
    Next r
    ' Simplified demonstration rails; real rail geometry is machine-specific.
    Set railShape = outputLayer.CreateRectangle2(originX + leftX - clearance - railWidth, topY, railWidth, rapport)
    railShape.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    railShape.Outline.SetNoOutline
    allShapes.Add railShape
    Set railShape = outputLayer.CreateRectangle2(originX + leftX + cols * w + clearance, topY, railWidth, rapport)
    railShape.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
    railShape.Outline.SetNoOutline
    allShapes.Add railShape
    Set composition = allShapes.Group
    ' Compress the WHOLE composition along length after imposition, as requested.
    composition.SetSize composition.SizeWidth, composition.SizeHeight * factor
    If reverse Then composition.Flip cdrFlipHorizontal
    composition.SetPositionEx cdrTopLeft, originX + leftX - clearance - railWidth, topY
    composition.CreateSelection
    doc.EndCommandGroup
    groupStarted = False
    doc.Unit = oldUnit
    doc.ReferencePoint = oldReference
    Application.Optimization = oldOptimization
    Application.Refresh
    MsgBox "DEMO: " & cols & " x " & rows & " labels on a NEW page." & vbCrLf & _
           "Original group unchanged. Review geometry and print settings before export.", vbInformation
    Exit Sub
Failed:
    errText = Err.Description: errNumber = Err.Number
    On Error Resume Next
    If groupStarted Then doc.EndCommandGroup
    If stateCaptured Then
        doc.Unit = oldUnit
        doc.ReferencePoint = oldReference
        Application.Optimization = oldOptimization
    End If
    Application.Refresh
    On Error GoTo 0
    If errNumber <> 18 Then MsgBox errText & vbCrLf & _
        "If a partial output page was created, Undo the DEMO label imposition group.", vbExclamation
End Sub

Public Sub TestLayoutMath()
    If MaxColumns(50, 220, 1.5, 2) <> 4 Then Err.Raise 5, , "Column test failed"
    If MaxRows(40, 300) <> 7 Then Err.Raise 5, , "Row test failed"
    If Abs(300 * 0.97 - 291) > 0.000001 Then Err.Raise 5, , "Distortion test failed"
    MsgBox "Math assertions passed inside VBA. This does not validate printing.", vbInformation
End Sub
