Attribute VB_Name = "LibCoreTest"
Option Explicit

'===============================================================================

Public Sub TestGetRotatedRect()
    Dim Rect As Rect
    Set Rect = ActiveLayer.CreateRectangle(0, 0, 3, 6).BoundingBox
    ActiveLayer.CreateRectangleRect GetRotatedRect(Rect)
End Sub

Public Sub TestShow()
    Show Empty
    Show Null
    Show 3
    Show "3"
    Show New Collection
    Show Pack(1, 2, Pack("x", "y", "z"))
    Show CreateColor
End Sub

Public Sub UnitContains()
    Debug.Assert Contains(Array(1, 2, 3), 1) = True
    Debug.Assert Contains(Array(1, 2, 3), "x") = False
    Debug.Assert Contains(Array(1, 2, 3), 3) = True
    Debug.Print "Contains is OK"
End Sub

Public Sub UnitContainsAll()
    Debug.Assert ContainsAll(Array(1, 2, 3), Array(3, 1, 2)) = True
    Debug.Assert ContainsAll(Array(1, 2, 3), Array(3, 1, 4)) = False
    Debug.Assert ContainsAll(Array(1, 2, 3), Array(1)) = True
    Debug.Print "ContainsAll is OK"
End Sub

Public Sub UnitDeduplicate()
    Debug.Assert Deduplicate(Array(1, 1, 2, 1, 2)).Count = 2
    Debug.Assert Deduplicate(Array(1, 1, 2, 1, 2))(2) = 2
    Debug.Print "Deduplicate is OK"
End Sub

Public Sub UnitHasPosition()
    Debug.Assert HasPosition(CreateRect) = True
    Debug.Assert HasPosition(New Collection) = False
    Debug.Assert HasPosition(123) = False
    Debug.Print "HasPosition is OK"
End Sub

Public Sub UnitHasSize()
    Debug.Assert HasSize(CreateRect) = True
    Debug.Assert HasPosition(New Collection) = False
    Debug.Assert HasPosition(123) = False
    Debug.Print "HasSize is OK"
End Sub

Public Sub UnitIsJust()
    Debug.Assert IsJust(0) = True
    Debug.Assert IsJust(1) = True
    Debug.Assert IsJust(New Collection) = True
    Debug.Assert IsJust(Empty) = False
    Debug.Assert IsJust(Null) = False
    Debug.Assert IsJust(Nothing) = False
    Debug.Assert IsJust(VBA.CVErr(ErrorCodes.ErrorInvalidArgument)) = False
    Debug.Print "IsJust is OK"
End Sub

Public Sub UnitNumberToFitArea()
    Debug.Assert _
        NumberToFitArea( _
            CreateRect(0, 0, 10, 10), _
            CreateRect(0, 0, 100, 100) _
        ) = 100
    Debug.Assert _
        NumberToFitArea( _
            CreateRect(0, 0, 10, 20), _
            CreateRect(0, 0, 10, 20) _
        ) = 1
    Debug.Assert _
        NumberToFitArea( _
            CreateRect(0, 0, 10, 20), _
            CreateRect(0, 0, 5, 5) _
        ) = 0
    Debug.Assert _
        NumberToFitArea( _
            CreateRect(0, 0, 10, 20), _
            CreateRect(0, 0, 21, 21) _
        ) = 2
End Sub

Public Sub UnitSpaceBox()
    With SpaceBox(CreateRect(0, 0, 100, 100), 20)
        Debug.Assert .Width = 140
        Debug.Assert .Height = 140
    End With
End Sub

Public Sub UnitSwap()
    Dim x As Long, y As Long
    x = 1
    y = 2
    Swap x, y
    Debug.Assert x = 2
    Debug.Assert y = 1
End Sub
