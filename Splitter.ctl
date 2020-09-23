VERSION 5.00
Begin VB.UserControl Splitter 
   Alignable       =   -1  'True
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   3150
   ToolboxBitmap   =   "Splitter.ctx":0000
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   1560
      ScaleHeight     =   2895
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "Splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------
' Splitter Control by Tim Humphrey
' 3/14/2001
' zzhumphreyt@techie.com
'
' ----------
'
' Usage: Add controls to the Splitter control the same way you do for a frame;
' set Child1 and/or Child2 to the names of the controls you want to be resized.
'
' - All sizes are in twips
'
' - The list box can only be resized in certain increments, 195 twips,
'   if you add one resize the Splitter control to match it.
'
' - While resizing, the escape key may be pressed to cancel and undo
'   the dragging
'
' ----------
'
' Edit History
'   10/15/2001
'       o Added AllowResize property
'       o The minimum SplitterSize can now be as low as 0
'   4/17/2001
'       o Added BorderStyle, MaxSize, MaxSizeAppliesTo,
'         SplitterPos, CurrSplitterPos, CurrRatioFromTop
'         and Maintain properties
'       o Rewrote code, where necessary, to support new properties and to
'         reduce complexity
'       o CurrSplitterPos and CurrRatioFromTop always report accurate readings;
'         previously RatioFromTop assumed CurrRatioFromTop's functionality
'         and could sometimes report inaccurate readings
'       o Out of necessity, gave invalid property values proper values
'       o Started all enumerations at 0, breaks compatibility with previous
'         version on OrientationConstants
'       o Removed design-time splitter appearance to make control easier to grab
'       o Reduced default splitter size
'   3/14/2001
'       o Initial creation
'--------------------------------------------------

Option Explicit
Option Compare Text

'-------------------- Enumerations --------------------
Public Enum AppearanceConstants
    vbFlat = 0
    vb3D
End Enum

Public Enum BorderConstants
    vbBSNone = 0
    vbFixedSingle
End Enum

Public Enum MaintainConstants
    MN_POS = 0
    MN_RATIO
End Enum

Public Enum MaxAppliesToConstants
    MX_CHILD1 = 0
    MX_CHILD2
End Enum

Public Enum OrientationConstants
    OC_HORIZONTAL = 0
    OC_VERTICAL
End Enum

'-------------------- Constants --------------------
'----- Property strings
Const kStrBorderStyle As String = "BorderStyle"
Const kStrSplitterAppearance As String = "SplitterAppearance"
Const kStrSplitterBorder As String = "SplitterBorder"
Const kstrSplitterColor As String = "SplitterColor"

Const kStrOrientation As String = "Orientation"
Const kStrSplitterSize As String = "SplitterSize"

Const kstrMaintain As String = "Maintain"
Const kStrSplitterPos As String = "SplitterPos"
Const kStrRatioFromTop As String = "RatioFromTop"

Const kStrChild1 As String = "Child1"
Const kStrChild2 As String = "Child2"

Const kStrMaxSize As String = "MaxSize"
Const kstrMaxSizeAppliesTo As String = "MaxSizeAppliesTo"
Const kStrMinSize1 As String = "MinSize1"
Const kStrMinSize2 As String = "MinSize2"
Const kStrMinSizeAux As String = "MinSizeAux"

Const kStrAllowResize As String = "AllowResize"
Const kStrLiveUpdate As String = "LiveUpdate"

'----- Defaults
Const kDefBorderStyle As Integer = vbBSNone
Const kDefSplitterAppearance As Integer = vb3D
Const kDefSplitterBorder As Integer = vbFixedSingle
Const kDefSplitterColor As Long = &H404040

Const kDefOrientation As Integer = OC_HORIZONTAL
Const kDefSplitterSize As Integer = 75

Const kDefMaintain As Integer = MN_RATIO
Const kDefSplitterPos As Integer = 0
Const kDefRatioFromTop As Single = 0.5

Const kDefChild1 As String = ""
Const kDefChild2 As String = ""

Const kDefMaxSize As Long = 0
Const kDefMaxSizeAppliesTo As Integer = MX_CHILD1
Const kDefMinSize1 As Long = 255
Const kDefMinSize2 As Long = 255
Const kDefMinSizeAux As Long = 255

Const kDefAllowResize As Boolean = True
Const kDefLiveUpdate As Boolean = True

'----- Busy bit-masks
Const kBusySplitterPos As Integer = &H1
Const kBusyRatioFromTop As Integer = &H2
Const kBusyCurrSplitterPos As Integer = &H4
Const kBusyCurrRatioFromTop As Integer = &H8

'-------------------- Variables --------------------
'----- Public properties
Private mSplitterAppearance As AppearanceConstants
Private mSplitterBorder As BorderConstants
Private mSplitterColor As Long

Private mOrientation As OrientationConstants
Private mSplitterSize As Integer

Private mMaintain As MaintainConstants
Private mSplitterPos As Integer
Private mRatioFromTop As Single

Private mChild1 As String
Private mChild2 As String

Private mMaxSize As Integer
Private mMaxSizeAppliesTo As MaxAppliesToConstants
Private mMinSize1 As Integer
Private mMinSize2 As Integer
Private mMinSizeAux As Integer

Private mAllowResize As Boolean
Private mLiveUpdate As Boolean

'----- Private properties
Private mAvailableAuxSpace As Integer
Private mMinRequiredSpace As Integer
Private mCurrRatioFromTop As Single

'----- Control use
Private gBusy As Integer
Private gResizeChildren As Boolean
Private gMoving As Boolean
Private gOrigPos As Integer
Private gOrigPoint As Integer

'-------------------- Events --------------------
Public Event Resize()
Attribute Resize.VB_Description = "Occurs when the child controls are resized."

'-------------------- API Types --------------------
Private Type Point
    X As Long
    Y As Long
End Type

'-------------------- API Functions --------------------
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As Boolean
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Point) As Boolean

Private Function CalcAvailableAuxSpace() As Integer
    Dim result As Integer
    
    Select Case Orientation
    Case OC_HORIZONTAL
        If UserControl.ScaleHeight > MinSizeAux Then
            result = UserControl.ScaleHeight
        Else
            result = MinSizeAux
        End If
    Case OC_VERTICAL
        If UserControl.ScaleWidth > MinSizeAux Then
            result = UserControl.ScaleWidth
        Else
            result = MinSizeAux
        End If
    End Select
    
    CalcAvailableAuxSpace = result
End Function

Private Function CalcMinRequiredSpace() As Integer
    CalcMinRequiredSpace = MinSize1 + SplitterSize + MinSize2
End Function

Private Function GetAvailableSpace() As Integer
    Select Case Orientation
    Case OC_HORIZONTAL
        GetAvailableSpace = UserControl.ScaleWidth
    Case OC_VERTICAL
        GetAvailableSpace = UserControl.ScaleHeight
    End Select
End Function

Private Function PosToRatio(availableSpace As Integer, pos As Integer) As Single
    If availableSpace > 0 Then
        PosToRatio = (pos + (SplitterSize \ 2)) / availableSpace
    Else
        PosToRatio = 0
    End If
End Function

Private Function RatioToPos(availableSpace As Integer, ratio As Single) As Integer
    RatioToPos = (availableSpace * ratio) - (SplitterSize \ 2)
End Function

Private Sub ResizeChildren()
    '-------------------- Variables --------------------
    Dim vObjChild1 As Object
    Dim vObjChild2 As Object
    
    Dim newLeft1 As Integer
    Dim newTop1 As Integer
    Dim newWidth1 As Integer
    Dim newHeight1 As Integer
    
    Dim newLeft2 As Integer
    Dim newTop2 As Integer
    Dim newWidth2 As Integer
    Dim newHeight2 As Integer
    
    '-------------------- Code --------------------
    If gResizeChildren Then
        UserControl.AutoRedraw = False
        
        Set vObjChild1 = objChild1
        Set vObjChild2 = objChild2
        
        'Hack around evil ListView control
        If Not (vObjChild1 Is Nothing) And (TypeName(vObjChild1) = "ListView") Then
            newLeft1 = -15
            newTop1 = -15
            newWidth1 = 30
            newHeight1 = 30
        End If
        
        If Not (vObjChild2 Is Nothing) And (TypeName(vObjChild2) = "ListView") Then
            newLeft2 = -15
            newTop2 = -15
            newWidth2 = 30
            newHeight2 = 30
        End If
        
        Select Case Orientation
        Case OC_HORIZONTAL
            If Not (vObjChild1 Is Nothing) Then
                newLeft1 = newLeft1 + 0
                newTop1 = newTop1 + 0
                newWidth1 = newWidth1 + CurrSplitterPos
                newHeight1 = newHeight1 + AvailableAuxSpace
                
                vObjChild1.Move newLeft1, newTop1, newWidth1, newHeight1
            End If
            
            If Not (vObjChild2 Is Nothing) Then
                newLeft2 = newLeft2 + CurrSplitterPos + SplitterSize
                newTop2 = newTop2 + 0
                newHeight2 = newHeight2 + AvailableAuxSpace
                
                If UserControl.ScaleWidth - (CurrSplitterPos + SplitterSize) >= MinSize2 Then
                    newWidth2 = newWidth2 + UserControl.ScaleWidth - (CurrSplitterPos + SplitterSize)
                Else
                    newWidth2 = newWidth2 + MinSize2
                End If
                
                vObjChild2.Move newLeft2, newTop2, newWidth2, newHeight2
            End If
        Case OC_VERTICAL
            If Not (vObjChild1 Is Nothing) Then
                newLeft1 = newLeft1 + 0
                newTop1 = newTop1 + 0
                newWidth1 = newWidth1 + AvailableAuxSpace
                newHeight1 = newHeight1 + CurrSplitterPos
                
                vObjChild1.Move newLeft1, newTop1, newWidth1, newHeight1
            End If
            
            If Not (vObjChild2 Is Nothing) Then
                newLeft2 = newLeft2 + 0
                newTop2 = newTop2 + CurrSplitterPos + SplitterSize
                newWidth2 = newWidth2 + AvailableAuxSpace
                
                If UserControl.ScaleHeight - (CurrSplitterPos + SplitterSize) >= MinSize2 Then
                    newHeight2 = newHeight2 + UserControl.ScaleHeight - (CurrSplitterPos + SplitterSize)
                Else
                    newHeight2 = newHeight2 + MinSize2
                End If
                
                vObjChild2.Move newLeft2, newTop2, newWidth2, newHeight2
            End If
        End Select
        
        RaiseEvent Resize
        
        UserControl.AutoRedraw = True
    End If
End Sub

Private Sub ResizeSplitter()
    Dim newPos As Integer
    
    Select Case Orientation
    Case OC_HORIZONTAL
        Select Case Maintain
        Case MN_POS
            newPos = SplitterPos
        Case MN_RATIO
            newPos = RatioToPos(UserControl.ScaleWidth, RatioFromTop)
        End Select
        
        newPos = VerifyNewPos(UserControl.ScaleWidth, newPos)
        picSplitter.Move newPos, 0, SplitterSize, AvailableAuxSpace
        CurrSplitterPos = newPos
    Case OC_VERTICAL
        Select Case Maintain
        Case MN_POS
            newPos = SplitterPos
        Case MN_RATIO
            newPos = RatioToPos(UserControl.ScaleHeight, RatioFromTop)
        End Select
        
        newPos = VerifyNewPos(UserControl.ScaleHeight, newPos)
        picSplitter.Move 0, newPos, AvailableAuxSpace, SplitterSize
        CurrSplitterPos = newPos
    End Select
End Sub

Private Sub UpdateSplitter()
    Select Case Maintain
    Case MN_POS
        CurrSplitterPos = SplitterPos
    Case MN_RATIO
        CurrRatioFromTop = RatioFromTop
    End Select
End Sub

Private Function VerifyNewPos(availableSpace As Integer, pos As Integer) As Integer
    '-------------------- Variables --------------------
    Dim newPos As Integer
    Dim lowerBound As Integer
    Dim size1Violated As Boolean
    
    '-------------------- Code --------------------
    If availableSpace > MinRequiredSpace Then
        newPos = pos
        
        'Correct bounds if needed
        If newPos < 0 Then
            newPos = 0
        End If
        If (newPos + SplitterSize) > availableSpace Then
            newPos = availableSpace - SplitterSize
        End If
        
        'Check MaxSize
        If MaxSize > 0 Then
            Select Case MaxSizeAppliesTo
            Case MX_CHILD1
                If newPos > MaxSize Then
                    newPos = MaxSize
                End If
            Case MX_CHILD2
                lowerBound = availableSpace - MaxSize - SplitterSize
                If newPos < lowerBound Then
                    newPos = lowerBound
                End If
            End Select
        End If
        
        'See if Child1 bounds violated
        size1Violated = False
        If newPos <= MinSize1 Then
            newPos = MinSize1
            size1Violated = True
        End If
        
        'See if Child2 bounds violated
        If Not size1Violated Then
            If (newPos + SplitterSize) > (availableSpace - MinSize2) Then
                newPos = availableSpace - MinSize2 - SplitterSize
            End If
        End If
    Else
        newPos = MinSize1
    End If
    
    VerifyNewPos = newPos
End Function

Public Property Get BorderStyle() As BorderConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for the control."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(value As BorderConstants)
    UserControl.BorderStyle = value
    PropertyChanged kStrBorderStyle
    UserControl_Resize
End Property

Public Property Get SplitterAppearance() As AppearanceConstants
Attribute SplitterAppearance.VB_Description = "The appearance of the splitter bar, only used while the splitter bar is moving and LiveUpdate is false."
Attribute SplitterAppearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SplitterAppearance = mSplitterAppearance
End Property

Public Property Let SplitterAppearance(value As AppearanceConstants)
    mSplitterAppearance = value
    PropertyChanged kStrSplitterAppearance
End Property

Public Property Get SplitterBorder() As BorderConstants
Attribute SplitterBorder.VB_Description = "The border style for the spiltter bar, only used while the splitter bar is moving and LiveUpdate is false."
Attribute SplitterBorder.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SplitterBorder = mSplitterBorder
End Property

Public Property Let SplitterBorder(value As BorderConstants)
    mSplitterBorder = value
    PropertyChanged kStrSplitterBorder
End Property

Public Property Get SplitterColor() As OLE_COLOR
Attribute SplitterColor.VB_Description = "The color of the splitter bar, only used while the splitter bar is moving and LiveUpdate is false."
Attribute SplitterColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SplitterColor = mSplitterColor
End Property

Public Property Let SplitterColor(value As OLE_COLOR)
    mSplitterColor = value
    PropertyChanged kstrSplitterColor
End Property

Public Property Get Orientation() As OrientationConstants
Attribute Orientation.VB_Description = "The flow of the child controls."
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Position"
    Orientation = mOrientation
End Property

Public Property Let Orientation(value As OrientationConstants)
    Dim oldPos As Integer
    
    oldPos = CurrSplitterPos
    mOrientation = value
    
    Select Case mOrientation
    Case OC_HORIZONTAL
        picSplitter.MousePointer = vbSizeWE
    Case OC_VERTICAL
        picSplitter.MousePointer = vbSizeNS
    End Select
    
    If Maintain = MN_POS Then
        CurrSplitterPos = oldPos
    End If
    
    PropertyChanged kStrOrientation
    UserControl_Resize
End Property

Public Property Get SplitterSize() As Integer
Attribute SplitterSize.VB_Description = "Returns/sets the size of the splitter bar."
Attribute SplitterSize.VB_ProcData.VB_Invoke_Property = ";Position"
    SplitterSize = mSplitterSize
End Property

Public Property Let SplitterSize(value As Integer)
    If value >= 0 Then
        mSplitterSize = value
    Else
        mSplitterSize = 0
    End If
    
    PropertyChanged kStrSplitterSize
    MinRequiredSpace = CalcMinRequiredSpace
    UserControl_Resize
End Property

Public Property Get Maintain() As MaintainConstants
Attribute Maintain.VB_Description = "Determines how the splitter bar changes when the control is resized."
Attribute Maintain.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Maintain = mMaintain
End Property

Public Property Let Maintain(value As MaintainConstants)
    mMaintain = value
    PropertyChanged kstrMaintain
    UpdateSplitter
End Property

Public Property Get SplitterPos() As Integer
Attribute SplitterPos.VB_Description = "Returns/sets the desired position of the splitter bar."
Attribute SplitterPos.VB_ProcData.VB_Invoke_Property = ";Position"
    SplitterPos = mSplitterPos
End Property

Public Property Let SplitterPos(value As Integer)
    If (gBusy And kBusySplitterPos) = 0 Then
        gBusy = gBusy + kBusySplitterPos
        
        If value >= 0 Then
            mSplitterPos = value
        Else
            mSplitterPos = 0
        End If
        PropertyChanged kStrSplitterPos
        
        'SplitterPos and RatioFromTop update each other, must prevent an infinite loop
        If (gBusy And kBusyRatioFromTop) = 0 Then
            RatioFromTop = PosToRatio(GetAvailableSpace, mSplitterPos)
            CurrSplitterPos = mSplitterPos
        End If
        
        gBusy = gBusy - kBusySplitterPos
    End If
End Property

Public Property Get CurrSplitterPos() As Integer
Attribute CurrSplitterPos.VB_Description = "Returns the current position of the splitter bar."
    Select Case Orientation
    Case OC_HORIZONTAL
        CurrSplitterPos = picSplitter.Left
    Case OC_VERTICAL
        CurrSplitterPos = picSplitter.Top
    End Select
End Property

Private Property Let CurrSplitterPos(value As Integer)
    Dim newPos As Integer
    
    If (gBusy And kBusyCurrSplitterPos) = 0 Then
        gBusy = gBusy + kBusyCurrSplitterPos
        
        If value >= 0 Then
            newPos = value
        Else
            newPos = 0
        End If
        
        Select Case Orientation
        Case OC_HORIZONTAL
            picSplitter.Left = VerifyNewPos(UserControl.ScaleWidth, newPos)
            CurrRatioFromTop = PosToRatio(UserControl.ScaleWidth, picSplitter.Left)
        Case OC_VERTICAL
            picSplitter.Top = VerifyNewPos(UserControl.ScaleHeight, newPos)
            CurrRatioFromTop = PosToRatio(UserControl.ScaleHeight, picSplitter.Top)
        End Select
        
        If (gBusy And kBusyCurrRatioFromTop) = 0 Then
            ResizeChildren
        End If
        
        gBusy = gBusy - kBusyCurrSplitterPos
    End If
End Property

Public Property Get RatioFromTop() As Single
Attribute RatioFromTop.VB_Description = "Returns/sets the desired percentage from the top/left to place the splitter bar."
Attribute RatioFromTop.VB_ProcData.VB_Invoke_Property = ";Position"
    RatioFromTop = mRatioFromTop
End Property

Public Property Let RatioFromTop(value As Single)
    If (gBusy And kBusyRatioFromTop) = 0 Then
        gBusy = gBusy + kBusyRatioFromTop
        
        Select Case True
        Case (value >= 0) And (value <= 1)
            mRatioFromTop = value
        Case value < 0
            mRatioFromTop = 0
        Case Else
            mRatioFromTop = 1
        End Select
        PropertyChanged kStrRatioFromTop
        
        'SplitterPos and RatioFromTop update each other, must prevent an infinite loop
        If (gBusy And kBusySplitterPos) = 0 Then
            SplitterPos = RatioToPos(GetAvailableSpace, mRatioFromTop)
            CurrRatioFromTop = mRatioFromTop
        End If
        
        gBusy = gBusy - kBusyRatioFromTop
    End If
End Property

Public Property Get CurrRatioFromTop() As Single
Attribute CurrRatioFromTop.VB_Description = "Returns the current percentage from the top/left of the splitter bar."
    CurrRatioFromTop = mCurrRatioFromTop
End Property

Private Property Let CurrRatioFromTop(value As Single)
    Dim newRatio As Single
    Dim availableSpace As Integer
    
    If (gBusy And kBusyCurrRatioFromTop) = 0 Then
        gBusy = gBusy + kBusyCurrRatioFromTop
        
        Select Case True
        Case (value >= 0) And (value <= 1)
            newRatio = value
        Case value < 0
            newRatio = 0
        Case Else
            newRatio = 1
        End Select
        
        availableSpace = GetAvailableSpace
        CurrSplitterPos = VerifyNewPos(availableSpace, RatioToPos(availableSpace, newRatio))
        mCurrRatioFromTop = PosToRatio(availableSpace, CurrSplitterPos)
        
        If (gBusy And kBusyCurrSplitterPos) = 0 Then
            ResizeChildren
        End If
        
        gBusy = gBusy - kBusyCurrRatioFromTop
    End If
End Property

Public Property Get Child1() As String
Attribute Child1.VB_Description = "Returns/sets the name of the control to appear at the left/top."
Attribute Child1.VB_ProcData.VB_Invoke_Property = ";Misc"
    Child1 = mChild1
End Property

Public Property Let Child1(value As String)
    mChild1 = value
    PropertyChanged kStrChild1
    ResizeChildren
End Property

Private Property Get objChild1() As Object
    '-------------------- Variables --------------------
    Dim found As Boolean
    Dim i As Integer
    
    '-------------------- Code --------------------
    If Child1 <> "" Then
        found = False
        
        For i = 0 To UserControl.ContainedControls.Count - 1
            If UserControl.ContainedControls(i).Name = Child1 Then
                Set objChild1 = UserControl.ContainedControls(i)
                found = True
                Exit For
            End If
        Next
        
        If Not found Then
            Set objChild1 = Nothing
        End If
    Else
        Set objChild1 = Nothing
    End If
End Property

Public Property Get Child2() As String
Attribute Child2.VB_Description = "Returns/sets the name of the control to appear at the right/bottom."
Attribute Child2.VB_ProcData.VB_Invoke_Property = ";Misc"
    Child2 = mChild2
End Property

Public Property Let Child2(value As String)
    mChild2 = value
    PropertyChanged kStrChild2
    ResizeChildren
End Property

Private Property Get objChild2() As Object
    '-------------------- Variables --------------------
    Dim found As Boolean
    Dim i As Integer
    
    '-------------------- Code --------------------
    If Child2 <> "" Then
        found = False
        
        For i = 0 To UserControl.ContainedControls.Count - 1
            If UserControl.ContainedControls(i).Name = Child2 Then
                Set objChild2 = UserControl.ContainedControls(i)
                found = True
                Exit For
            End If
        Next
        
        If Not found Then
            Set objChild2 = Nothing
        End If
    Else
        Set objChild2 = Nothing
    End If
End Property

Public Property Get MaxSize() As Integer
Attribute MaxSize.VB_Description = "Returns/sets the maximum size of a child; 0 is unlimited."
Attribute MaxSize.VB_ProcData.VB_Invoke_Property = ";Position"
    MaxSize = mMaxSize
End Property

Public Property Let MaxSize(value As Integer)
    Dim newMax As Integer
    
    If value >= 0 Then
        newMax = value
    Else
        newMax = 0
    End If
    
    If newMax > 0 Then
        Select Case MaxSizeAppliesTo
        Case MX_CHILD1
            If newMax < MinSize1 Then
                newMax = MinSize1
            End If
        Case MX_CHILD2
            If newMax < MinSize2 Then
                newMax = MinSize2
            End If
        End Select
    End If
    
    mMaxSize = newMax
    PropertyChanged kStrMaxSize
    UpdateSplitter
End Property

Public Property Get MaxSizeAppliesTo() As MaxAppliesToConstants
Attribute MaxSizeAppliesTo.VB_Description = "Determines which child the MaxSize property applies to."
Attribute MaxSizeAppliesTo.VB_ProcData.VB_Invoke_Property = ";Position"
    MaxSizeAppliesTo = mMaxSizeAppliesTo
End Property

Public Property Let MaxSizeAppliesTo(value As MaxAppliesToConstants)
    mMaxSizeAppliesTo = value
    PropertyChanged kstrMaxSizeAppliesTo
    UpdateSplitter
End Property

Public Property Get MinSize1() As Long
Attribute MinSize1.VB_Description = "Returns/sets the minimum size allowed for Child1."
Attribute MinSize1.VB_ProcData.VB_Invoke_Property = ";Position"
    MinSize1 = mMinSize1
End Property

Public Property Let MinSize1(value As Long)
    If value >= 0 Then
        mMinSize1 = value
    Else
        mMinSize1 = 0
    End If
    
    If (MaxSize > 0) And (MaxSizeAppliesTo = MX_CHILD1) Then
        If mMinSize1 > MaxSize Then
            mMinSize1 = MaxSize
        End If
    End If
    
    PropertyChanged kStrMinSize1
    MinRequiredSpace = CalcMinRequiredSpace
    UpdateSplitter
End Property

Public Property Get MinSize2() As Long
Attribute MinSize2.VB_Description = "Returns/sets the minimum size allowed for Child2."
Attribute MinSize2.VB_ProcData.VB_Invoke_Property = ";Position"
    MinSize2 = mMinSize2
End Property

Public Property Let MinSize2(value As Long)
    If value >= 0 Then
        mMinSize2 = value
    Else
        mMinSize2 = 0
    End If
    
    If (MaxSize > 0) And (MaxSizeAppliesTo = MX_CHILD2) Then
        If mMinSize2 > MaxSize Then
            mMinSize2 = MaxSize
        End If
    End If
    
    PropertyChanged kStrMinSize2
    MinRequiredSpace = CalcMinRequiredSpace
    UpdateSplitter
End Property

Public Property Get MinSizeAux() As Long
Attribute MinSizeAux.VB_Description = "Returns/sets the minimum size allowed for the control for the opposite orientation."
Attribute MinSizeAux.VB_ProcData.VB_Invoke_Property = ";Position"
    MinSizeAux = mMinSizeAux
End Property

Public Property Let MinSizeAux(value As Long)
    If value >= 0 Then
        mMinSizeAux = value
    Else
        mMinSizeAux = 0
    End If
    
    PropertyChanged kStrMinSizeAux
    UserControl_Resize
End Property

Public Property Get AllowResize() As Boolean
Attribute AllowResize.VB_Description = "True if the user can move the splitter bar."
Attribute AllowResize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowResize = mAllowResize
End Property

Public Property Let AllowResize(value As Boolean)
    mAllowResize = value
    picSplitter.Visible = value
    PropertyChanged kStrAllowResize
End Property

Public Property Get LiveUpdate() As Boolean
Attribute LiveUpdate.VB_Description = "True if the child controls should be resized as the splitter bar is moved."
Attribute LiveUpdate.VB_ProcData.VB_Invoke_Property = ";Behavior"
    LiveUpdate = mLiveUpdate
End Property

Public Property Let LiveUpdate(value As Boolean)
    mLiveUpdate = value
    PropertyChanged kStrLiveUpdate
End Property

Private Property Get AvailableAuxSpace() As Integer
    AvailableAuxSpace = mAvailableAuxSpace
End Property

Private Property Let AvailableAuxSpace(value As Integer)
    If value >= 0 Then
        mAvailableAuxSpace = value
    Else
        mAvailableAuxSpace = 0
    End If
End Property

Private Property Get MinRequiredSpace() As Integer
    MinRequiredSpace = mMinRequiredSpace
End Property

Private Property Let MinRequiredSpace(value As Integer)
    If value >= 0 Then
        mMinRequiredSpace = value
    Else
        mMinRequiredSpace = 0
    End If
End Property

Private Sub picSplitter_KeyPress(KeyAscii As Integer)
    If gMoving And (KeyAscii = vbKeyEscape) Then
        If LiveUpdate Then
            CurrSplitterPos = gOrigPos
            ResizeChildren
        Else
            picSplitter.BackColor = vbButtonFace
            picSplitter.BorderStyle = vbBSNone
            CurrSplitterPos = gOrigPos
        End If
        
        gMoving = False
    End If
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '-------------------- Variables --------------------
    Dim originalPoint As Point
    Dim modifiedPoint As Point
    Dim offsetX As Long
    Dim offsetY As Long
    
    '-------------------- Code --------------------
    If Button = vbLeftButton Then
        gOrigPos = CurrSplitterPos
        Select Case Orientation
        Case OC_HORIZONTAL
            gOrigPoint = X
        Case OC_VERTICAL
            gOrigPoint = Y
        End Select
        
        picSplitter.ZOrder 0
        If Not LiveUpdate Then
            'Changing the picture box to include a border will alter the
            'scalewidth/scaleheight, thereby immediately triggering a
            'mouse moved event; we must compensate for this
            GetCursorPos originalPoint
            ScreenToClient picSplitter.hWnd, originalPoint
            
            picSplitter.Appearance = SplitterAppearance
            picSplitter.BackColor = SplitterColor
            picSplitter.BorderStyle = SplitterBorder
            
            GetCursorPos modifiedPoint
            ScreenToClient picSplitter.hWnd, modifiedPoint
            
            Select Case Orientation
            Case OC_HORIZONTAL
                offsetX = (originalPoint.X - modifiedPoint.X) * Screen.TwipsPerPixelX
                gOrigPoint = gOrigPoint - offsetX
            Case OC_VERTICAL
                offsetY = (originalPoint.Y - modifiedPoint.Y) * Screen.TwipsPerPixelY
                gOrigPoint = gOrigPoint - offsetY
            End Select
        End If
        gMoving = True
    End If
End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '-------------------- Variables --------------------
    Dim currPos As Integer
    Dim newPos As Integer
    Dim availableSpace As Integer
    
    '-------------------- Code --------------------
    If gMoving And (Button = vbLeftButton) Then
        '----- Calculate bounds
        Select Case Orientation
        Case OC_HORIZONTAL
            currPos = picSplitter.Left
            newPos = picSplitter.Left + (X - gOrigPoint)
            availableSpace = UserControl.ScaleWidth
        Case OC_VERTICAL
            currPos = picSplitter.Top
            newPos = picSplitter.Top + (Y - gOrigPoint)
            availableSpace = UserControl.ScaleHeight
        End Select
        
        newPos = VerifyNewPos(availableSpace, newPos)
        If currPos <> newPos Then
            If LiveUpdate Then
                CurrSplitterPos = newPos
            Else
                gResizeChildren = False
                CurrSplitterPos = newPos
                gResizeChildren = True
            End If
        End If
    End If
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If Not LiveUpdate Then
            picSplitter.BackColor = vbButtonFace
            picSplitter.BorderStyle = vbBSNone
            ResizeChildren
        End If
        
        SplitterPos = CurrSplitterPos
        gMoving = False
    End If
End Sub

Private Sub UserControl_InitProperties()
    gResizeChildren = False
    
    BorderStyle = kDefBorderStyle
    SplitterAppearance = kDefSplitterAppearance
    SplitterBorder = kDefSplitterBorder
    SplitterColor = kDefSplitterColor
    
    Orientation = kDefOrientation
    SplitterSize = kDefSplitterSize
    
    Maintain = kDefMaintain
    SplitterPos = RatioToPos(GetAvailableSpace, kDefRatioFromTop)
    RatioFromTop = kDefRatioFromTop
    
    Child1 = kDefChild1
    Child2 = kDefChild2
    
    MaxSize = kDefMaxSize
    MaxSizeAppliesTo = kDefMaxSizeAppliesTo
    MinSize1 = kDefMinSize1
    MinSize2 = kDefMinSize2
    MinSizeAux = kDefMinSizeAux
    
    AllowResize = kDefAllowResize
    LiveUpdate = kDefLiveUpdate
    
    gResizeChildren = True
    ResizeChildren
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    gResizeChildren = False
    
    With PropBag
        BorderStyle = .ReadProperty(kStrBorderStyle, kDefBorderStyle)
        SplitterAppearance = .ReadProperty(kStrSplitterAppearance, kDefSplitterAppearance)
        SplitterBorder = .ReadProperty(kStrSplitterBorder, kDefSplitterBorder)
        SplitterColor = .ReadProperty(kstrSplitterColor, kDefSplitterColor)
        
        Orientation = .ReadProperty(kStrOrientation, kDefOrientation)
        SplitterSize = .ReadProperty(kStrSplitterSize, kDefSplitterSize)
        
        Maintain = .ReadProperty(kstrMaintain, kDefMaintain)
        SplitterPos = .ReadProperty(kStrSplitterPos, kDefSplitterPos)
        RatioFromTop = .ReadProperty(kStrRatioFromTop, kDefRatioFromTop)
        
        Child1 = .ReadProperty(kStrChild1, kDefChild1)
        Child2 = .ReadProperty(kStrChild2, kDefChild2)
        
        MaxSize = .ReadProperty(kStrMaxSize, kDefMaxSize)
        MaxSizeAppliesTo = .ReadProperty(kstrMaxSizeAppliesTo, kDefMaxSizeAppliesTo)
        MinSize1 = .ReadProperty(kStrMinSize1, kDefMinSize1)
        MinSize2 = .ReadProperty(kStrMinSize2, kDefMinSize2)
        MinSizeAux = .ReadProperty(kStrMinSizeAux, kDefMinSizeAux)
        
        AllowResize = .ReadProperty(kStrAllowResize, kDefAllowResize)
        LiveUpdate = .ReadProperty(kStrLiveUpdate, kDefLiveUpdate)
    End With
    
    gResizeChildren = True
    ResizeChildren
End Sub

Private Sub UserControl_Resize()
    AvailableAuxSpace = CalcAvailableAuxSpace
    ResizeSplitter
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty kStrBorderStyle, BorderStyle, kDefBorderStyle
        .WriteProperty kStrSplitterAppearance, SplitterAppearance, kDefSplitterAppearance
        .WriteProperty kStrSplitterBorder, SplitterBorder, kDefSplitterBorder
        .WriteProperty kstrSplitterColor, SplitterColor, kDefSplitterColor
        
        .WriteProperty kStrOrientation, Orientation, kDefOrientation
        .WriteProperty kStrSplitterSize, SplitterSize, kDefSplitterSize
        
        .WriteProperty kstrMaintain, Maintain, kDefMaintain
        .WriteProperty kStrSplitterPos, SplitterPos, kDefSplitterPos
        .WriteProperty kStrRatioFromTop, RatioFromTop, kDefRatioFromTop
        
        .WriteProperty kStrChild1, Child1, kDefChild1
        .WriteProperty kStrChild2, Child2, kDefChild2
        
        .WriteProperty kStrMaxSize, MaxSize, kDefMaxSize
        .WriteProperty kstrMaxSizeAppliesTo, MaxSizeAppliesTo, kDefMaxSizeAppliesTo
        .WriteProperty kStrMinSize1, MinSize1, kDefMinSize1
        .WriteProperty kStrMinSize2, MinSize2, kDefMinSize2
        .WriteProperty kStrMinSizeAux, MinSizeAux, kDefMinSizeAux
        
        .WriteProperty kStrAllowResize, AllowResize, kDefAllowResize
        .WriteProperty kStrLiveUpdate, LiveUpdate, kDefLiveUpdate
    End With
End Sub
