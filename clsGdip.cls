VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGdip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary
'@@require
'@@reference
'@@license
'@@author
'@@create
'@@modify
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       接口继承
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有常量
'------------------------------------------------------------------------------
Public Enum ErrorTypes
    GDIPlusFailed
    SaveImageToFileFailed
End Enum

Public Enum SmoothingModeEnum
    SmoothingModeDefault = 0&
    SmoothingModeHighSpeed = 1&
    SmoothingModeHighQuality = 2&
    SmoothingModeNone = 3&
    SmoothingModeAntiAlias8x4 = 4&
    SmoothingModeAntiAlias = 4&
    SmoothingModeAntiAlias8x8 = 5&
End Enum

'------------------------------------------------------------------------------
'       公有数据类型
'------------------------------------------------------------------------------
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    Type As Long
    value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

'------------------------------------------------------------------------------
'       公有变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有API
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       事件声明
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       私有常量
'------------------------------------------------------------------------------
Private Const JPEGEncoder As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Private Const BMPEncoder As String = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
Private Const GIFEncoder As String = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
Private Const TIFFEncoder As String = "{557CF405-1A04-11D3-9A73-0000F81EF32E}"
Private Const PNGEncoder As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
Private Const QualityParameter As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有变量
'------------------------------------------------------------------------------
' Token
Private pvlToken As Long

' Handles
Private pvlGraph As Long
Private pvlHDC As Long
Private pvlPen As Long
Private pvoStdPicture As StdPicture
' Tools
Private pvoPen As clsGdipPen
' Other
Private pvtGdiInput As GdiplusStartupInput
Private pvlStartUpRet As Long
Private pvtSmoothMode As SmoothingModeEnum
Private errHandler As clsError

'------------------------------------------------------------------------------
'       属性变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有API
'------------------------------------------------------------------------------
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, _
    inputbuf As GdiplusStartupInput, _
    ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, _
    ByVal hPal As Long, _
    Bitmap As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal imageHandle As Long, _
    ByVal fileName As Long, _
    clsidEncoder As GUID, _
    encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, _
    id As GUID) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, _
    hGraphics As Long) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal Graphics As Long, _
    ByVal Pen As Long, _
    ByVal X1 As Single, _
    ByVal Y1 As Single, _
    ByVal X2 As Single, _
    ByVal Y2 As Single) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, _
    ByVal Width As Single, _
    ByVal unit As Long, _
    Pen As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, _
    ByVal SmoothingMd As Long) As Long

'//////////////////////////////////////////////////////////////////////////////
'//
'//      类
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       初始化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    pvtGdiInput.GdiplusVersion = 1
    pvlStartUpRet = GdiplusStartup(pvlToken, pvtGdiInput, 0)
    If pvlStartUpRet <> 0 Then
        errHandler.RaiseError ErrorTypes.GDIPlusFailed, "clsGDI", "Failed to initialize GDI+"
    End If
    Call pmInit
End Sub


'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
    pvoPen.ClosePen
    GdiplusShutdown pvlToken
End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'//      事件处理
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Private Sub pmInit()
    pvtSmoothMode = SmoothingModeDefault
    
    Set errHandler = New clsError
    errHandler.BaseErrorNumber = &H1
    
    Set pvoPen = New clsGdipPen
End Sub

'//////////////////////////////////////////////////////////////////////////////
'//
'//      继承实现
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Public Sub SetStdPicture(newStdPicture As StdPicture)
    Set pvoStdPicture = newStdPicture
End Sub

Public Sub SetHDC(ByVal newHDC As Long)
    pvlHDC = newHDC
End Sub

Public Sub SaveStdPictureToFile(ByVal fileName As String, Optional ByVal quality As Byte = 80)
    Dim lBitmapRes As Long, lBitmap As Long
    Dim lSaveImageRes As Long

    lBitmapRes = GdipCreateBitmapFromHBITMAP(pvoStdPicture.Handle, 0, lBitmap)

    If lBitmapRes = 0 Then
        Dim tJpgEncoder As GUID
        Dim tParams As EncoderParameters
    
        '初始化解码器的GUID标识
        CLSIDFromString StrPtr(JPEGEncoder), tJpgEncoder
        
        '设置解码器参数
        tParams.Count = 1
        With tParams.Parameter ' Quality
            '得到Quality参数的GUID标识
            CLSIDFromString StrPtr(QualityParameter), .GUID
            .NumberOfValues = 1
            .Type = 4
            .value = VarPtr(quality)
        End With
    
        '保存图像
        lSaveImageRes = GdipSaveImageToFile(lBitmap, StrPtr(fileName), tJpgEncoder, tParams)
        If lSaveImageRes <> 0 Then
            errHandler.RaiseError ErrorTypes.SaveImageToFileFailed, "clsGDI", "Failed to save to file."
        End If

        '销毁GDI+图像
        GdipDisposeImage lBitmap
    End If
End Sub

Public Sub SetSmoothMode(ByVal SmoothingMode As SmoothingModeEnum)
    pvtSmoothMode = SmoothingMode
End Sub

Public Sub DrawLine(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
    Dim lRet As Long
    lRet = GdipCreateFromHDC(pvlHDC, pvlGraph)
    lRet = GdipSetSmoothingMode(pvlGraph, pvtSmoothMode)
    GdipDrawLine pvlGraph, pvoPen.Handle, X1, Y1, X2, Y2
End Sub
