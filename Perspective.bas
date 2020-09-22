Attribute VB_Name = "modPerspBlt"
Option Explicit
Public Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const HTCAPTION As Integer = 2
Public Const WS_EX_LAYERED As Long = &H80000
Public Const GWL_EXSTYLE As Integer = (-20)
Public Const LWA_COLORKEY As Integer = &H1
Public Const LWA_ALPHA As Integer = &H2

Public Const WM_NCLBUTTONDOWN As Integer = &HA1
Public Const COLORONCOLOR As Long = 3
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" ( _
                                                ByVal hdc As Long, _
                                                ByVal X As Long, _
                                                ByVal Y As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal nHeight As Long, _
                                                ByVal hSrcDC As Long, _
                                                ByVal xSrc As Long, _
                                                ByVal ySrc As Long, _
                                                ByVal nSrcWidth As Long, _
                                                ByVal nSrcHeight As Long, _
                                                ByVal dwRop As Long) As Long

Public Sub PerspBltX(ByVal outDC As Long, ByVal outX As Long, ByVal outY As Long, _
                     ByVal OutWidth As Long, ByVal outStartHeight As Long, ByVal outEndHeight As Long, _
                     ByVal outYOff As Long, ByVal inDC As Long, ByVal inWidth As Long, ByVal inHeight As Long)

  Dim loopx As Long
  Dim InterpPos As Single
  Dim InterpH As Long
  Dim StartLoop As Long
  Dim EndLoop As Long

    If OutWidth = 0 Then
        Exit Sub
    End If

    StartLoop = 0
    EndLoop = OutWidth

    If OutWidth < 0 Then
        StartLoop = OutWidth
        EndLoop = 0
    End If

    SetStretchBltMode outDC, COLORONCOLOR

    For loopx = StartLoop To EndLoop
        InterpPos = loopx / OutWidth
        InterpH = InterpPos * (outEndHeight - outStartHeight)
        StretchBlt outDC, loopx + outX, outY + (InterpPos * outYOff), 1, outStartHeight + InterpH, inDC, (InterpPos * inWidth), 0, 1, inHeight, vbSrcCopy
    Next loopx

End Sub

Public Sub PerspBltY(ByVal outDC As Long, ByVal outX As Long, ByVal outY As Long, _
                     ByVal outStartWidth As Long, ByVal outEndWidth As Long, ByVal OutHeight As Long, _
                     ByVal outXOff As Long, ByVal inDC As Long, ByVal inWidth As Long, ByVal inHeight As Long)

  Dim LoopY As Long
  Dim InterpPos As Single
  Dim InterpW As Long
  Dim StartLoop As Long
  Dim EndLoop As Long

    If OutHeight = 0 Then
        Exit Sub
    End If

    StartLoop = 0
    EndLoop = OutHeight

    If OutHeight < 0 Then
        StartLoop = OutHeight
        EndLoop = 0
    End If

    SetStretchBltMode outDC, COLORONCOLOR
    For LoopY = StartLoop To EndLoop
        InterpPos = LoopY / OutHeight
        InterpW = InterpPos * (outEndWidth - outStartWidth)
        StretchBlt outDC, outX + (InterpPos * outXOff), LoopY + outY, outStartWidth + InterpW, 1, inDC, 0, (InterpPos * inHeight), inWidth, 1, vbSrcCopy
    Next LoopY

End Sub

