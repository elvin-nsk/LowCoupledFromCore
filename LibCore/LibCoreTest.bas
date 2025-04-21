Attribute VB_Name = "LibCoreTest"
Option Explicit

'===============================================================================

Public Sub TestAllSpecs()
    SpecContains
    SpecContainsAll
    SpecDeduplicate
    SpecExtractSubstrings
    SpecFString
    SpecHasPosition
    SpecHasSize
    SpecIsJust
    SpecNumberToFitArea
    SpecSpaceBox
    SpecSwap
    Debug.Print "All PASSED"
End Sub

'===============================================================================
' # Manual tests

Public Sub TestGetRotatedRect()
    Dim Rect As Rect
    Set Rect = ActiveLayer.CreateRectangle(0, 0, 3, 6).BoundingBox
    ActiveLayer.CreateRectangleRect GetRotatedRect(Rect)
End Sub

Public Sub TestShapeName()
    ShapeName(ActiveShape) = "Gradio WebUI for creators and developers, featuring key TTS (Edge-TTS, kokoro) and zero-shot Voice Cloning (E2 & F5-TTS, CosyVoice), with Whisper audio processing, YouTube download, Demucs vocal isolation, and multilingual translation."
    Show ShapeName(ActiveShape)
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

'===============================================================================
' # Auto tests

Public Sub SpecColorToShowable()
    'Show ColorToShowable(ActiveShape.Fill.UniformColor)
    Debug.Assert ColorToShowable(Cyan) = "C:100 M:0 Y:0 K:0"
End Sub

Public Sub SpecContains()
    Debug.Assert Contains(Array(1, 2, 3), 1) = True
    Debug.Assert Contains(Array(1, 2, 3), "x") = False
    Debug.Assert Contains(Array(1, 2, 3), 3) = True
End Sub

Public Sub SpecContainsAll()
    Debug.Assert ContainsAll(Array(1, 2, 3), Array(3, 1, 2)) = True
    Debug.Assert ContainsAll(Array(1, 2, 3), Array(3, 1, 4)) = False
    Debug.Assert ContainsAll(Array(1, 2, 3), Array(1)) = True
End Sub

Public Sub SpecDeduplicate()
    Debug.Assert Deduplicate(Array(1, 1, 2, 1, 2)).Count = 2
    Debug.Assert Deduplicate(Array(1, 1, 2, 1, 2))(2) = 2
End Sub

Public Sub SpecExtractSubstrings()
    Debug.Assert ExtractSubstrings("Это {важный} фрагмент", "{}")(1) = "важный"
    Debug.Assert ExtractSubstrings("Это _важный_ фрагмент", "_")(1) = "важный"
    Debug.Assert ExtractSubstrings("Это _важные_ _фрагменты_", "_")(2) = "фрагменты"
End Sub

Public Sub SpecFString()
    Dim Text As String: Text = "I have {0} coins {1} cents each."
    Debug.Assert FString(Text, 10, 5) = "I have 10 coins 5 cents each."
    Debug.Assert Text = "I have {0} coins {1} cents each."
    Text = "I have fake {q}coins{q}."
    Debug.Assert FString(Text) = "I have fake " & Chr(34) & "coins" & Chr(34) & "."
End Sub

Public Sub SpecHasPosition()
    Debug.Assert HasPosition(CreateRect) = True
    Debug.Assert HasPosition(New Collection) = False
    Debug.Assert HasPosition(123) = False
End Sub

Public Sub SpecHasSize()
    Debug.Assert HasSize(CreateRect) = True
    Debug.Assert HasPosition(New Collection) = False
    Debug.Assert HasPosition(123) = False
End Sub

Public Sub SpecIndexOfChar()
    Debug.Assert IndexOfChar("Text", "e") = 2
    Debug.Assert IndexOfChar("Text", "t") = 1
    Debug.Assert IndexOfChar("Text", "t", 2) = 4
    Debug.Assert IndexOfChar("Text", "t", , True) = 4
End Sub

Public Sub SpecIsJust()
    Debug.Assert IsJust(0) = True
    Debug.Assert IsJust(1) = True
    Debug.Assert IsJust(New Collection) = True
    Debug.Assert IsJust(Empty) = False
    Debug.Assert IsJust(Null) = False
    Debug.Assert IsJust(Nothing) = False
    Debug.Assert IsJust(VBA.CVErr(ErrorCodes.ErrorInvalidArgument)) = False
End Sub

Public Sub SpecNumberToFitArea()
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

Public Sub SpecSpaceBox()
    With SpaceBox(CreateRect(0, 0, 100, 100), 20)
        Debug.Assert .Width = 140
        Debug.Assert .Height = 140
    End With
End Sub

Public Sub SpecSwap()
    Dim x As Long, y As Long
    x = 1
    y = 2
    Swap x, y
    Debug.Assert x = 2
    Debug.Assert y = 1
End Sub
