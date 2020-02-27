VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGdipPen"
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
Public Enum PenColor
    Black = vbButtonStandard
End Enum

Public Enum GdiPenError
    CreatePenFailed
    DeletePenFailed
    SetPenColorFailed
End Enum

Public Enum GpUnitEnum  ' aka Unit
   UnitWorld      ' 0 -- World coordinate (non-physical unit)
   UnitDisplay    ' 1 -- Variable -- for PageTransform only
   UnitPixel      ' 2 -- Each unit is one device pixel.
   UnitPoint      ' 3 -- Each unit is a printer's point, or 1/72 inch.
   UnitInch       ' 4 -- Each unit is 1 inch.
   UnitDocument   ' 5 -- Each unit is 1/300 inch.
   UnitMillimeter ' 6 -- Each unit is 1 millimeter.
End Enum

'------------------------------------------------------------------------------
'       公有数据类型
'------------------------------------------------------------------------------


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


'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有变量
'------------------------------------------------------------------------------
' Handles
Private pvlPen As Long
Private pvbPenInitialized As Boolean
' Others
Private lRet As Long
Private errHandler As clsError

'------------------------------------------------------------------------------
'       属性变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有API
'------------------------------------------------------------------------------
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, _
    ByVal Width As Single, _
    ByVal unit As Long, _
    Pen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal Pen As Long) As Long
Private Declare Function GdipSetPenColor Lib "gdiplus" (ByVal Pen As Long, _
    ByVal ARGB As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, _
    lpSource As Any, ByVal cbCopy As Long)
'//////////////////////////////////////////////////////////////////////////////
'//
'//      类
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       初始化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    pvbPenInitialized = False
    Set errHandler = New clsError
    errHandler.BaseErrorNumber = &H2
    lRet = GdipCreatePen1(&HFF562B00, 1, GpUnitEnum.UnitPixel, pvlPen)
    If lRet <> 0 Then
        errHandler.RaiseError GdiPenError.CreatePenFailed, "clsGdipPen", "Failed to create Pen."
    End If
    pvbPenInitialized = True
End Sub

'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Call ClosePen
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
Private Function ARGB(ByVal Alpha As Byte, ByVal Red As Byte, _
    ByVal Green As Byte, ByVal Blue As Byte) As Long
    ARGB = ARGBFromRGB(Alpha, RGB(Red, Green, Blue))
End Function

Private Function ARGBFromRGB(ByVal Alpha As Byte, ByVal RGB As Long)
    CopyMemory ByVal VarPtr(ARGBFromRGB) + 3, Alpha, 1
    CopyMemory ByVal VarPtr(ARGBFromRGB) + 2, VarPtr(RGB) + 2, 3
End Function

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
Public Property Get Handle() As Long
    Handle = pvlPen
End Property

'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Public Sub SetPenColor(ByVal RGB As Long, ByVal Alpha As Byte)
    lRet = GdipSetPenColor(pvlPen, ARGBFromRGB(Alpha, RGB))
    If lRet <> 0 Then
        errHandler.RaiseError GdiPenError.SetPenColorFailed, "clsGdipPen", "Failed to set Pen color."
    End If
End Sub

Public Sub ClosePen()
    If Not pvbPenInitialized Then
        Exit Sub
    End If
    lRet = GdipDeletePen(pvlPen)
    If lRet <> 0 Then
        errHandler.RaiseError GdiPenError.DeletePenFailed, "clsGdipPen", "Failed to delete Pen handle."
    End If
    pvbPenInitialized = False
End Sub


