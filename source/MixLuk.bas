Attribute VB_Name = "MixLuk"
'===============================================================================
'   ??????          : MixLuk
'   ??????          : 2022.10.29
'   ?????           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   ?????           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

'===============================================================================

Private Const a As Double = 5 '??

'===============================================================================

Sub MakeDashes()

    If RELEASE Then On Error GoTo Catch
    
    Dim Lines As ShapeRange
    With InitData.GetShapes(LayerMustBeEnabled:=True)
        If .IsError Then Exit Sub
        Set Lines = .Shapes
    End With
    Dim Source As ShapeRange
    Set Source = ActiveSelectionRange
    ActiveDocument.Unit = cdrMillimeter
    
    lib_elvin.BoostStart "MakeDashes", RELEASE
    
    Dim Line As Shape
    Dim InvalidCount As Long
    For Each Line In Lines
        If IsPerpendicularLine(Line) Then
            MakeDashesOnLine Line, a
            ContinueLine Line, a
        Else
            InvalidCount = InvalidCount + 1
        End If
    Next Line
    
    Source.CreateSelection
    
    If InvalidCount > 0 Then _
        VBA.MsgBox "????????? " & InvalidCount & " ????? ?????????"
    
Finally:
    lib_elvin.BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================

Private Function IsPerpendicularLine(ByVal Line As Shape) As Boolean
    Const Tolerance As Double = 0.001
    IsPerpendicularLine = _
        DiffWithinTolerance(Line.SizeWidth, 0, Tolerance) _
     Or DiffWithinTolerance(Line.SizeHeight, 0, Tolerance)
End Function

Private Function MakeDashesOnLine( _
                     ByVal Line As Shape, _
                     ByVal Offset As Double _
                 ) As ShapeRange
    Dim Dashes As ShapeRange
    Set Dashes = CreateShapeRange
    
    With Line.Curve.Nodes
        Dashes.Add _
            FlipIfNeeded( _
                MakeDashAtNode(.First, Offset), Line _
            )
        Dashes.Add _
            FlipIfNeeded( _
                MakeDashAtNode(.Last, Offset), Line _
            )
    End With
    Dashes.SetOutlineProperties Line.Outline.Width * 2
    Set MakeDashesOnLine = Dashes
End Function

Private Function MakeDashAtNode( _
                     ByVal Node As Node, _
                     ByVal Offset As Double _
                 ) As Shape
    Set MakeDashAtNode = _
        ActiveLayer.CreateLineSegment( _
            Node.PositionX - Offset, Node.PositionY - Offset, _
            Node.PositionX + Offset, Node.PositionY + Offset _
        )
End Function

Private Function FlipIfNeeded( _
                     ByVal Dash As Shape, _
                     ByVal Line As Shape _
                 ) As Shape
    If Not lib_elvin.IsLandscape(Line) Then Dash.Flip cdrFlipHorizontal
    Set FlipIfNeeded = Dash
End Function

Private Function ContinueLine( _
                     ByVal Line As Shape, _
                     ByVal Offset As Double _
                 ) As Shape
    If lib_elvin.IsLandscape(Line) Then
        Line.SizeWidth = Line.SizeWidth + Offset * 2
    Else
        Line.SizeHeight = Line.SizeHeight + Offset * 2
    End If
    Set ContinueLine = Line
End Function


'===============================================================================
' # ?????

Private Sub testSomething()
'
End Sub
