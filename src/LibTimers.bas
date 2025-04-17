Attribute VB_Name = "LibTimers"
Option Explicit

#Const Windows = (Mac = 0)
#Const x64 = Win64

#If Windows Then
    #If VBA7 Then
        Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByVal pvargSrc As LongPtr)
    #Else
        Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByVal pvargSrc As Long)
    #End If
#End If

#If VBA7 = 0 Then
    Public Enum LongPtr: [_]: End Enum
#End If

#If x64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const ptrSize = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const ptrSize = 4
#End If

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Type FakeVariant
    vt As Integer
    wReserved(0 To 2) As Integer
    ptrs(0 To 1) As LongPtr
End Type

Private Sub EntryPoint(): End Sub

Public Function GetTimerProc(ByVal st As SafeTimer) As LongPtr
    If st Is Nothing Then Exit Function
    Static pa As PointerAccessor
    Const FADF_AUTO As Long = &H1
    Const FADF_FIXEDSIZE As Long = &H10
    '
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = ptrSize
        pa.sa.cLocks = 1
        #If Windows Then
            MemLongPtrRef(VarPtr(pa)) = VarPtr(pa.sa)
        #End If
    End If
    '
    pa.sa.pvData = ObjPtr(st)
    pa.sa.rgsabound0.cElements = 1
    pa.sa.pvData = pa.arr(0) + ptrSize * 9
    Dim aPtr As LongPtr: aPtr = pa.arr(0) 'SafeTimer.DummyASM
    '
#If x64 Then
    pa.sa.pvData = aPtr
    If (pa.arr(0) And &HFFFFFF) <> &HC1894C Then
                                  pa.arr(0) = &HC1894C      '4C89C1   MOV RCX,R8              ;nIDEvent (instance)
        pa.sa.pvData = aPtr + 3:  pa.arr(0) = &HCA894C      '4C89CA   MOV RDX,R9              ;wTime
        pa.sa.pvData = aPtr + 6:  pa.arr(0) = &H18B48       '488B01   MOV RAX,QWORD PTR [RCX] ;vtbl
        pa.sa.pvData = aPtr + 9:  pa.arr(0) = &H55          '55       PUSH RBP
        pa.sa.pvData = aPtr + 10: pa.arr(0) = &HEC8B48      '488BEC   MOV RBP,RSP
        pa.sa.pvData = aPtr + 13: pa.arr(0) = &H20EC8348    '4883EC20 SUB RSP,0x20
        pa.sa.pvData = aPtr + 17: pa.arr(0) = &H4050FF      'FF5040   CALL QWORD PTR [RAX+40] ;SafeTimer.TimerProc
        pa.sa.pvData = aPtr + 20: pa.arr(0) = &H20C48348    '4883C420 ADD RSP,0x20
        pa.sa.pvData = aPtr + 24: pa.arr(0) = &H5D          '5D       POP RBP
        pa.sa.pvData = aPtr + 25: pa.arr(0) = &HC3          'C3       RET
    End If
    '
    'Note that EBMode is found at EntryPoint+37 and VBA does the work for us
    GetTimerProc = VBA.Int(AddressOf EntryPoint)
    pa.sa.pvData = GetTimerProc + 55
    pa.arr(0) = aPtr
#Else
    '
    '
    '
    '
    '
    '
    '
    '
    '
    '
#End If
    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = NullPtr
End Function

#If Windows Then
Private Property Let MemLongPtrRef(ByVal memAddress As LongPtr _
                                 , ByVal newValue As LongPtr)
    Const VT_BYREF As Long = &H4000
    Dim memValue As Variant
    Dim remoteVT As Variant
    Dim fv As FakeVariant
    '
    fv.ptrs(0) = VarPtr(memValue)
    fv.vt = vbInteger + VT_BYREF
    VariantCopy remoteVT, VarPtr(fv) 'Init VarType ByRef
    '
#If x64 Then 'Cannot assign LongLong ByRef
    Dim c As Currency
    RemoteAssign memValue, VarPtr(newValue), remoteVT, vbCurrency + VT_BYREF, c, memValue
    RemoteAssign memValue, memAddress, remoteVT, vbCurrency + VT_BYREF, memValue, c
#Else 'Can assign Long ByRef
    RemoteAssign memValue, memAddress, remoteVT, vbLong + VT_BYREF, memValue, newValue
#End If
End Property
Private Sub RemoteAssign(ByRef memValue As Variant _
                       , ByRef memAddress As LongPtr _
                       , ByRef remoteVT As Variant _
                       , ByVal newVT As VbVarType _
                       , ByRef targetVariable As Variant _
                       , ByRef newValue As Variant)
    memValue = memAddress
    remoteVT = newVT
    targetVariable = newValue
    remoteVT = vbEmpty
End Sub
#End If
