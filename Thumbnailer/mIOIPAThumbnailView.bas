Attribute VB_Name = "mIOIPAThumbnailView"
'========================================================================================
' FileName:    mIOIPAThumbnailView.bas (mIOleInPlaceActivate.bas)
' Author:      Mike Gainer, Matt Curland and Bill Storage
' Date:        09 January 1999
'
' Requires:    OleGuids.tlb (in IDE only)
'
' Description:
' Allows you to replace the standard IOLEInPlaceActiveObject interface for a
' UserControl with a customisable one.  This allows you to take control
' of focus in VB controls.
'
' The code could be adapted to replace other UserControl OLE interfaces.
'
' ---------------------------------------------------------------------------------------
' Visit vbAccelerator, advanced, free source for VB programmers
' http://vbaccelerator.com
'========================================================================================

Option Explicit

'========================================================================================
' Lightweight object definition
'========================================================================================

Public Type IPAOHookStructThumbnailView
    lpVTable    As Long                    'VTable pointer
    IPAOReal    As IOleInPlaceActiveObject 'Un-AddRefed pointer for forwarding calls
    Ctl         As ucThumbnailView         'Un-AddRefed native class pointer for making Friend calls
    ThisPointer As Long
End Type

'========================================================================================
' API
'========================================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsEqualGUID Lib "ole32" (iid1 As GUID2, iid2 As GUID2) As Long

Private Type GUID2
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

'========================================================================================
' Constants and member variables
'========================================================================================

Private Const S_OK                  As Long = 0
Private IID_IOleInPlaceActiveObject As GUID2
Private m_IPAOVTable(9)             As Long

'========================================================================================
' Functions
'========================================================================================

Public Sub InitIPAO(IPAOHookStruct As IPAOHookStructThumbnailView, Ctl As ucThumbnailView)
    
  Dim IPAO As IOleInPlaceActiveObject
    
    With IPAOHookStruct
        Set IPAO = Ctl
        Call CopyMemory(.IPAOReal, IPAO, 4)
        Call CopyMemory(.Ctl, Ctl, 4)
        .lpVTable = GetVTable
        .ThisPointer = VarPtr(IPAOHookStruct)
    End With
End Sub

Public Sub TerminateIPAO(IPAOHookStruct As IPAOHookStructThumbnailView)
    With IPAOHookStruct
        Call CopyMemory(.IPAOReal, 0&, 4)
        Call CopyMemory(.Ctl, 0&, 4)
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function GetVTable() As Long

    ' Set up the vTable for the interface and return a pointer to it
    If (m_IPAOVTable(0) = 0) Then
        m_IPAOVTable(0) = AddressOfFunction(AddressOf QueryInterface)
        m_IPAOVTable(1) = AddressOfFunction(AddressOf AddRef)
        m_IPAOVTable(2) = AddressOfFunction(AddressOf Release)
        m_IPAOVTable(3) = AddressOfFunction(AddressOf GetWindow)
        m_IPAOVTable(4) = AddressOfFunction(AddressOf ContextSensitiveHelp)
        m_IPAOVTable(5) = AddressOfFunction(AddressOf TranslateAccelerator)
        m_IPAOVTable(6) = AddressOfFunction(AddressOf OnFrameWindowActivate)
        m_IPAOVTable(7) = AddressOfFunction(AddressOf OnDocWindowActivate)
        m_IPAOVTable(8) = AddressOfFunction(AddressOf ResizeBorder)
        m_IPAOVTable(9) = AddressOfFunction(AddressOf EnableModeless)
        '--- init GUID2
        With IID_IOleInPlaceActiveObject
            .Data1 = &H117&
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
    End If
    GetVTable = VarPtr(m_IPAOVTable(0))
End Function

Private Function AddressOfFunction(LPFN As Long) As Long
    ' Work around, VB thinks lPtr = AddressOf Method is an error
    AddressOfFunction = LPFN
End Function

'========================================================================================
' Interface implemenattion
'========================================================================================

Private Function AddRef(This As IPAOHookStructThumbnailView) As Long
    AddRef = This.IPAOReal.AddRef
End Function

Private Function Release(This As IPAOHookStructThumbnailView) As Long
    Release = This.IPAOReal.Release
End Function

Private Function QueryInterface(This As IPAOHookStructThumbnailView, riid As GUID2, pvObj As Long) As Long
    ' Install the interface if required
    If (IsEqualGUID(riid, IID_IOleInPlaceActiveObject)) Then
        ' Install alternative IOleInPlaceActiveObject interface implemented here
        pvObj = This.ThisPointer
        Call AddRef(This)
        QueryInterface = 0
      Else
        ' Use the default support for the interface:
        QueryInterface = This.IPAOReal.QueryInterface(ByVal VarPtr(riid), pvObj)
    End If
End Function

Private Function GetWindow(This As IPAOHookStructThumbnailView, phwnd As Long) As Long
    GetWindow = This.IPAOReal.GetWindow(phwnd)
End Function

Private Function ContextSensitiveHelp(This As IPAOHookStructThumbnailView, ByVal fEnterMode As Long) As Long
    ContextSensitiveHelp = This.IPAOReal.ContextSensitiveHelp(fEnterMode)
End Function

Private Function TranslateAccelerator(This As IPAOHookStructThumbnailView, lpMsg As Msg) As Long
    ' Check if we want to override the handling of this key code:
    If (This.Ctl.frTranslateAccel(lpMsg)) Then
        TranslateAccelerator = S_OK
      Else
        ' If not pass it on to the standard UserControl TranslateAccelerator method:
        TranslateAccelerator = This.IPAOReal.TranslateAccelerator(ByVal VarPtr(lpMsg))
    End If
End Function

Private Function OnFrameWindowActivate(This As IPAOHookStructThumbnailView, ByVal fActivate As Long) As Long
    OnFrameWindowActivate = This.IPAOReal.OnFrameWindowActivate(fActivate)
End Function

Private Function OnDocWindowActivate(This As IPAOHookStructThumbnailView, ByVal fActivate As Long) As Long
    OnDocWindowActivate = This.IPAOReal.OnDocWindowActivate(fActivate)
End Function

Private Function ResizeBorder(This As IPAOHookStructThumbnailView, prcBorder As RECT, ByVal puiWindow As IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
    ResizeBorder = This.IPAOReal.ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
End Function

Private Function EnableModeless(This As IPAOHookStructThumbnailView, ByVal fEnable As Long) As Long
    EnableModeless = This.IPAOReal.EnableModeless(fEnable)
End Function
