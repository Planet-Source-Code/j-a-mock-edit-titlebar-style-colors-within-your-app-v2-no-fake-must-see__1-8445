VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlColor1    As Long
Private mlColor2    As Long
Private mfAngle     As Single

Private Const mlDefColor1   As Long = &HFFFFD0
Private Const mlDefColor2   As Long = &H400000
Private Const mfDefAngle    As Single = 315

Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180

Private Type PointSng
    X   As Single
    Y   As Single
End Type

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private Type RectAPI
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Const PS_SOLID As Long = 0
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Function Draw(picObj As Object) As Boolean

Dim lRet    As Long
Dim lIdx    As Long
Dim lTime   As Long
Dim uRect   As RectAPI

    On Error GoTo LocalError
    
    lRet = LockWindowUpdate(picObj.hWnd)
    
    lRet = GetClientRect(picObj.hWnd, uRect)
    
    If lRet <> 0 Then
        If uRect.Right > 1 And uRect.Bottom > 1 Then
            lIdx = DrawGradient(picObj.hDC, uRect.Right, uRect.Bottom)
            Draw = (lIdx > 0)
        End If
    End If
        
NormalExit:
    lRet = LockWindowUpdate(0)
    Exit Function
    
LocalError:
    MsgBox Err.Description, vbExclamation
    Resume NormalExit

End Function
Public Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long

Dim lIdx    As Long
Dim lRed    As Long
Dim lGrn    As Long
Dim lBlu    As Long
Dim fRedStp As Single
Dim fGrnStp As Single
Dim fBluStp As Single

    If lSteps < 2 Then lSteps = 2
    
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000
    
    fRedStp = Div(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
    fGrnStp = Div(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
    fBluStp = Div(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))
    
    ReDim laRetColors(lSteps - 1)
    laRetColors(0) = lColor1
    laRetColors(lSteps - 1) = lColor2
    For lIdx = 1 To lSteps - 2
        laRetColors(lIdx) = CLng(lRed + (fRedStp * CSng(lIdx))) + _
            (CLng(lGrn + (fGrnStp * CSng(lIdx))) * &H100&) + _
            (CLng(lBlu + (fBluStp * CSng(lIdx))) * &H10000)
    Next lIdx
    
    BlendColors = lSteps

End Function
Private Function DrawGradient(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long

Dim bDone       As Boolean
Dim iIncX       As Integer
Dim iIncY       As Integer
Dim lIdx        As Long
Dim lRet        As Long
Dim hPen        As Long
Dim hOldPen     As Long
Dim lPointCnt   As Long
Dim laColors()  As Long
Dim fMovX       As Single
Dim fMovY       As Single
Dim fDist       As Single
Dim fAngle      As Single
Dim fLongSide   As Single
Dim uTmpPt      As POINTAPI
Dim uaPts()     As POINTAPI
Dim uaTmpPts()  As PointSng
    
    On Error GoTo LocalError
    
    ReDim uaTmpPts(2)
    uaTmpPts(2).X = Int(lWidth / 2)
    uaTmpPts(2).Y = Int(lHeight / 2)
    
    fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
    fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2
    
    uaTmpPts(0).X = uaTmpPts(2).X - fDist
    uaTmpPts(0).Y = uaTmpPts(2).Y
    uaTmpPts(1).X = uaTmpPts(2).X + fDist
    uaTmpPts(1).Y = uaTmpPts(2).Y
    
    fAngle = CDbl(mfAngle + 90) - Int(Int(CDbl(mfAngle + 90) / 360#) * 360#)
    
    Call RotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
    Call RotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)
    
    Select Case mfAngle
        Case 0 To 90
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                iIncX = 0
                iIncY = -1
            End If
        Case 90 To 180
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                iIncX = 0
                iIncY = -1
            End If
        Case 180 To 270
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                iIncX = 0
                iIncY = 1
            End If
        Case Else
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                iIncX = 0
                iIncY = 1
            End If
    End Select
    
    ReDim uaPts(999)
    
    uaPts(0).X = uaTmpPts(0).X + fMovX
    uaPts(0).Y = uaTmpPts(0).Y + fMovY
    uaPts(1).X = uaTmpPts(1).X + fMovX
    uaPts(1).Y = uaTmpPts(1).Y + fMovY
    
    lIdx = 2
    
    While Not bDone
        uaPts(lIdx).X = uaPts(lIdx - 2).X + iIncX
        uaPts(lIdx).Y = uaPts(lIdx - 2).Y + iIncY
        lIdx = lIdx + 1
        Select Case True
            Case iIncX > 0
                bDone = uaPts(lIdx - 1).X > lWidth And uaPts(lIdx - 2).X > lWidth
            Case iIncX < 0
                bDone = uaPts(lIdx - 1).X < 0 And uaPts(lIdx - 2).X < 0
            Case iIncY > 0
                bDone = uaPts(lIdx - 1).Y > lHeight And uaPts(lIdx - 2).Y > lHeight
            Case iIncY < 0
                bDone = uaPts(lIdx - 1).Y < 0 And uaPts(lIdx - 2).Y < 0
        End Select
        If (lIdx Mod 1000) = 0 Then
            ReDim Preserve uaPts(UBound(uaPts) + 1000)
        End If
    Wend
    
    ReDim Preserve uaPts(lIdx - 1)
    
    lRet = BlendColors(mlColor1, mlColor2, lIdx / 2, laColors)
    
    For lIdx = 0 To UBound(uaPts) - 1 Step 2
        lRet = MoveToEx(hDC, uaPts(lIdx).X, uaPts(lIdx).Y, uTmpPt)
        hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx / 2)))
        hOldPen = SelectObject(hDC, hPen)
        lRet = LineTo(hDC, uaPts(lIdx + 1).X, uaPts(lIdx + 1).Y)
        lRet = SelectObject(hDC, hOldPen)
        lRet = DeleteObject(hPen)
    Next lIdx
    
    DrawGradient = lIdx
    
NormalExit:
    
    Erase laColors
    Erase uaPts
    Erase uaTmpPts
    Exit Function
    
LocalError:
    MsgBox Err.Description, vbExclamation, "GradientRect.cls"
    DrawGradient = 0
    Resume
    
End Function

Private Sub RotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Single)

Dim fDX         As Single
Dim fDY         As Single
Dim fRadians    As Single

    fRadians = fDegrees * RADS
    fDX = uRotatePt.X - uAxisPt.X
    fDY = uRotatePt.Y - uAxisPt.Y
    uRotatePt.X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
    uRotatePt.Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    
End Sub

Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double

    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If

End Function

Public Property Let Color1(ByVal lData As Long)
    
Dim lIdx As Long

    mlColor1 = lData
    If mlColor1 < 0 Then
        lIdx = (mlColor1 And Not &H80000000)
        If lIdx >= 0 And lIdx <= 24 Then
            mlColor1 = GetSysColor(lIdx)
        End If
    End If
    
End Property


Public Property Get Color1() As Long
    Color1 = mlColor1
End Property

Public Property Let Color2(ByVal lData As Long)

Dim lIdx As Long

    mlColor2 = lData
    If mlColor2 < 0 Then
        lIdx = (mlColor2 And Not &H80000000)
        If lIdx >= 0 And lIdx <= 24 Then
            mlColor2 = GetSysColor(lIdx)
        End If
    End If

End Property

Public Property Get Color2() As Long
    Color2 = mlColor2
End Property

Public Property Let Angle(ByVal fData As Single)
    
    mfAngle = CDbl(fData) - Int(Int(CDbl(fData) / 360#) * 360#)

End Property

Public Property Get Angle() As Single
    Angle = mfAngle
End Property

Private Sub Class_Initialize()

    mlColor1 = mlDefColor1
    mlColor2 = mlDefColor2
    mfAngle = mfDefAngle
    
End Sub
