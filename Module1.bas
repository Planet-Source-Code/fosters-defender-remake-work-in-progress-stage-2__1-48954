Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount& Lib "kernel32" ()
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Enum PodState
    Free = 1
    Captured = 2
    Destroyed = 3
    Falling = 4
End Enum
Public Type udtPod
    X As Long
    Y As Long
    TerrainPoint As Long
    TerrainPointDirection As Long
    State As PodState
    LastTickCount As Long
    TimeTilNextPodMove As Long
End Type
Public Type udtExplosion
    Free As Boolean
    X As Long
    Y As Long
    CurrentFrame As Long
    TimeTillNextFrame As Long
    LastFrameTime As Long
End Type
Public Type udtSmallExplosion
    Free As Boolean
    X As Long
    Y As Long
    CurrentFrame As Long
    TimeTillNextFrame As Long
    LastFrameTime As Long
End Type
Public Enum ShipDirection
    DefenderLeft = 1
    DefenderRight = 2
End Enum
Public Enum ShipType
    xwing = 1
    falcon = 2
End Enum
Public Enum LaserColor
    greenlaser = 14483420
    redlaser = 14474495
    BlueLaser = 16768220
End Enum
Public Enum LaserSegment
    BuildLaser = 1
    HazeEffect = 2
    DestroyLaser = 3
End Enum
Public Type udtLaser
    Free As Boolean
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    Segment As LaserSegment
    SegmentCount As Long
    MainColor As LaserColor
    CurrentColor As Long
    LastTickCount As Long
    TimeTillNextAction As Long
End Type

Public Type udtStarFieldStars
    X As Long
    Y As Long
    Speed As Long
    StartColor As Long
    CurrentColor As Long
    ColorDirection As Long
End Type
Public Type udtDefender
    X As Long
    Y As Long
    Direction As ShipDirection
    Ship As ShipType
    CurrentFrame As Long
    ChangingDirection As Boolean
    yMovement As Long
End Type
Public Type udtTerrainPoint
    X As Long
    Y As Long
End Type
Public Type udtTerrain
    Points() As udtTerrainPoint
    NumPoints As Long
    StartColor As Long
    CurrentColor As Long
    ColorDirection As Long
    X As Long
End Type
Public Const TopOfScreen As Long = 40
Public Const BottomOfScreen As Long = 20
Public lNumPods As Long
Public Pods() As udtPod
Public Terrain As udtTerrain
Public Defender As udtDefender
Public Explosion() As udtExplosion
Public SmallExplosion() As udtExplosion
Public Lasers() As udtLaser
Public Const LaserLength As Long = 370
Public Const Pi = 3.1415926535898
Public StarFieldStars() As udtStarFieldStars
Public NumStars As Long

Function GimmeX(ByRef aIn As Single, ByRef lIn As Long) As Long
    GimmeX = sIn(aIn * (Pi / 180)) * lIn
End Function
Function GimmeY(ByRef aIn As Single, ByRef lIn As Long) As Long
    GimmeY = Cos(aIn * (Pi / 180)) * lIn
End Function
Public Function AdjustBrightness(ByRef RGB_In As Long, ByRef ShiftPercentage As Integer, Optional GoExtreme As Boolean = False) As Long
Dim lColor As Long
Dim r As Single, G As Single, B As Single

    lColor = RGB_In
    r = lColor Mod &H100
    lColor = lColor \ &H100
    G = lColor Mod &H100
    lColor = lColor \ &H100
    B = lColor Mod &H100

    r = r + ((r / 100) * ShiftPercentage)
    G = G + ((G / 100) * ShiftPercentage)
    B = B + ((B / 100) * ShiftPercentage)
    
    If r > 255 Or G > 255 Or B > 255 Then
        If GoExtreme Then
            If r > 255 Then r = 255
            If G > 255 Then G = 255
            If B > 255 Then B = 255
            AdjustBrightness = RGB(r, G, B)
        Else
            AdjustBrightness = RGB_In
        End If
    ElseIf r < 20 Or G < 20 Or B < 20 Then
        If GoExtreme Then
            If r < 20 Then r = 0
            If G < 20 Then G = 0
            If B < 20 Then B = 0
        Else
            AdjustBrightness = RGB_In
        End If
    Else
        AdjustBrightness = RGB(r, G, B)
    End If
End Function
