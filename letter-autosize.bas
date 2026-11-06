Attribute VB_Name = "Auto_Size_Letter"

Sub FitLetterAndScale()
    Const MARGIN As Double = 0.25        ' paper inches on each side
    Dim pg As Visio.Page, ps As Visio.Shape
    Dim L As Double, B As Double, R As Double, T As Double
    Dim dw As Double, dh As Double, k As Double, k2 As Double
    Dim paperW As Double, paperH As Double
    Dim undoID As Long

    Set pg = ActivePage
    Set ps = pg.PageSheet
    If pg.Shapes.Count = 0 Then MsgBox "Page is empty.": Exit Sub

    undoID = Application.BeginUndoScope("Fit to Letter and Scale")
    On Error GoTo Fail

    ' 1. Measure the drawing (drawing units)
    pg.BoundingBox visBBoxUprightWH + visBBoxExtents + visBBoxDrawingCoords, L, B, R, T
    dw = R - L
    dh = T - B

    ' 2. Letter, orientation matched to the drawing
    If dw >= dh Then
        paperW = 11: paperH = 8.5
        ps.CellsU("PrintPageOrientation").FormulaU = "2"   ' landscape
    Else
        paperW = 8.5: paperH = 11
        ps.CellsU("PrintPageOrientation").FormulaU = "1"   ' portrait
    End If
    ps.CellsU("PaperKind").FormulaU = "1"                  ' Letter

    ' 3. Scale factor = drawing units per paper inch
    k = dw / (paperW - 2 * MARGIN)
    k2 = dh / (paperH - 2 * MARGIN)
    If k2 > k Then k = k2

    ps.CellsU("DrawingResizeType").FormulaU = "2"          ' Auto Size off
    ps.CellsU("DrawingScale").ResultIU = ps.CellsU("PageScale").ResultIU * k
    ps.CellsU("PageWidth").ResultIU = paperW * k
    ps.CellsU("PageHeight").ResultIU = paperH * k

    ' 4. Center everything on the page
    pg.CenterDrawing

    Application.EndUndoScope undoID, True
    Exit Sub
Fail:
    Application.EndUndoScope undoID, False
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub