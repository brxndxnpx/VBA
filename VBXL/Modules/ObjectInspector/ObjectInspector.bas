Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TTYPEDESC
    #If Win64 Then
    pTypeDesc As LongLong
    #Else
    pTypeDesc As Long
    #End If
    vt As Integer
End Type

Private Type TPARAMDESC
    #If Win64 Then
    pPARAMDESCEX As LongLong
    #Else
    pPARAMDESCEX As Long
    #End If
    wParamFlags As Integer
End Type

Private Type TELEMDESC
    tdesc  As TTYPEDESC
    pdesc  As TPARAMDESC
End Type

Type TYPEATTR
    aGUID As GUID
    LCID As Long
    dwReserved As Long
    memidConstructor As Long
    memidDestructor As Long
    #If Win64 Then
    lpstrSchema As LongLong
    #Else
    lpstrSchema As Long
    #End If
    cbSizeInstance As Integer
    typekind As Long
    cFuncs As Integer
    cVars As Integer
    cImplTypes As Integer
    cbSizeVft As Integer
    cbAlignment As Integer
    wTypeFlags As Integer
    wMajorVerNum As Integer
    wMinorVerNum As Integer
    tdescAlias As Long
    idldescType As Long
End Type

Type FUNCDESC
    memid As Long
    #If Win64 Then
    lReserved1 As LongLong
    lprgelemdescParam As LongLong
    #Else
    lReserved1 As Long
    lprgelemdescParam As Long
    #End If
    funckind As Long
    INVOKEKIND As Long
    CallConv As Long
    cParams As Integer
    cParamsOpt As Integer
    oVft As Integer
    cReserved2 As Integer
    elemdescFunc As TELEMDESC
    wFuncFlags As Integer
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
#End If


'Private Const TableOffsetFac = 1

Sub TESTX()
    Dim ct As Variant
'    Dim d As New CheckRunOnBaseDocument
'
'    d.DocumentID = "123456"
'
'    ct = GetPropertySetterType(d, "DocumentID")
'
'    Stop
    
    
    Dim FS    As New FileSystemObject
    Dim F1    As New Folder
    Dim F2    As Scripting.Folder
    
    
'    VbCallType
    
    
    Set F2 = FS.GetFolder("C:\Users\bphommathep.TOTALMM\OneDrive - TMM, Inc\Desktop\1.20")
    
    ct = GetPropertySetterType(F2, "Files")
    
'    F1.Injest F2

    Stop

    Dim Properties As List
    '      Set Properties = GetObjectMembers(F2, FuncType:=0)
    Set Properties = GetObjectMembers(F2, VbCallType.VbGet)
      
      
    Dim PropSource As List
    Dim PropTarget As List
      
    Set PropSource = GetObjectMembers(F2, VbCallType.VbGet)
    Set PropTarget = GetObjectMembers(F1, VbCallType.VbGet)
      
      
      
End Sub



Public Function GetPropertySetterType(ByVal Target As Object, ByVal PropertyName) As VbCallType
    Dim TypeAttribute As TYPEATTR, FunctionDescription As FUNCDESC
    Dim aGUID(0 To 11) As Long, lFuncsCount As Long
    Dim ITypeInfo As IUnknown, IDispatch As IUnknown
    Dim PropName As String
      
    #If Win64 Then
        Const TableOffsetFac = 2
        Dim aTYPEATTR() As LongLong, aFUNCDESC() As LongLong, farPtr As LongLong
    #Else
        Const TableOffsetFac = 1
        Dim aTYPEATTR() As Long, aFUNCDESC() As Long, farPtr As Long
    #End If
    
    Const CC_STDCALL As Long = 4
    Const IUNK_QueryInterface As Long = 0
    Const IDSP_GetTypeInfo As Long = 16 * TableOffsetFac
    Const ITYP_GetTypeAttr As Long = 12 * TableOffsetFac
    Const ITYP_GetFuncDesc As Long = 20 * TableOffsetFac
    Const ITYP_GetDocument As Long = 48 * TableOffsetFac
    Const ITYP_ReleaseTypeAttr As Long = 76 * TableOffsetFac
    Const ITYP_ReleaseFuncDesc As Long = 80 * TableOffsetFac

    aGUID(0) = &H20400
    aGUID(2) = &HC0&
    aGUID(3) = &H46000000
      
    CallFunction_COM ObjPtr(Target), IUNK_QueryInterface, vbLong, CC_STDCALL, VarPtr(aGUID(0)), VarPtr(IDispatch)
      
    If IDispatch Is Nothing Then Exit Function

    CallFunction_COM ObjPtr(IDispatch), IDSP_GetTypeInfo, vbLong, CC_STDCALL, 0&, 0&, VarPtr(ITypeInfo)
    If ITypeInfo Is Nothing Then Exit Function
      
    CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetTypeAttr, vbLong, CC_STDCALL, VarPtr(farPtr)
    If farPtr = 0& Then Exit Function

    CopyMemory ByVal VarPtr(TypeAttribute), ByVal farPtr, LenB(TypeAttribute)
      
    ReDim aTYPEATTR(LenB(TypeAttribute))
      
    CopyMemory ByVal VarPtr(aTYPEATTR(0)), TypeAttribute, UBound(aTYPEATTR)
      
    CallFunction_COM ObjPtr(ITypeInfo), ITYP_ReleaseTypeAttr, vbEmpty, CC_STDCALL, farPtr
      
    For lFuncsCount = 0 To TypeAttribute.cFuncs - 1
        CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetFuncDesc, vbLong, CC_STDCALL, lFuncsCount, VarPtr(farPtr)
            
        If farPtr = 0 Then Exit For
            
        CopyMemory ByVal VarPtr(FunctionDescription), ByVal farPtr, LenB(FunctionDescription)
            
        ReDim aFUNCDESC(LenB(FunctionDescription))
            
        CopyMemory ByVal VarPtr(aFUNCDESC(0)), FunctionDescription, UBound(aFUNCDESC)
            
        CallFunction_COM ObjPtr(ITypeInfo), ITYP_ReleaseFuncDesc, vbEmpty, CC_STDCALL, farPtr
        CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(PropName), 0, 0, 0
        CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(PropName), 0, 0, 0

        With FunctionDescription
            If .INVOKEKIND <> VbGet And PropName = PropertyName Then GetPropertySetterType = .INVOKEKIND: Exit Function
        End With
    Next

End Function

Function GetObjectMembers(ByVal Target As Object, Optional ByVal FuncType As VbCallType) As List
    Dim TypeAttribute As TYPEATTR, FunctionDescription As FUNCDESC
    Dim aGUID(0 To 11) As Long, lFuncsCount As Long
    Dim ITypeInfo As IUnknown, IDispatch As IUnknown
    Dim PropName As String, PropInfo As PropertyInfo, PropList As New List
      
    #If Win64 Then
        Const TableOffsetFac = 2
        Dim aTYPEATTR() As LongLong, aFUNCDESC() As LongLong, farPtr As LongLong
    #Else
        Const TableOffsetFac = 1
        Dim aTYPEATTR() As Long, aFUNCDESC() As Long, farPtr As Long
    #End If
    
    Const CC_STDCALL As Long = 4
    Const IUNK_QueryInterface As Long = 0
    Const IDSP_GetTypeInfo As Long = 16 * TableOffsetFac
    Const ITYP_GetTypeAttr As Long = 12 * TableOffsetFac
    Const ITYP_GetFuncDesc As Long = 20 * TableOffsetFac
    Const ITYP_GetDocument As Long = 48 * TableOffsetFac
    Const ITYP_ReleaseTypeAttr As Long = 76 * TableOffsetFac
    Const ITYP_ReleaseFuncDesc As Long = 80 * TableOffsetFac

    aGUID(0) = &H20400
    aGUID(2) = &HC0&
    aGUID(3) = &H46000000
      
      
    '      Dim TargetPtr as longlong
      
      
    CallFunction_COM ObjPtr(Target), IUNK_QueryInterface, vbLong, CC_STDCALL, VarPtr(aGUID(0)), VarPtr(IDispatch)
      
    If IDispatch Is Nothing Then Exit Function

    CallFunction_COM ObjPtr(IDispatch), IDSP_GetTypeInfo, vbLong, CC_STDCALL, 0&, 0&, VarPtr(ITypeInfo)
    If ITypeInfo Is Nothing Then Exit Function
      
    CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetTypeAttr, vbLong, CC_STDCALL, VarPtr(farPtr)
    If farPtr = 0& Then Exit Function

    CopyMemory ByVal VarPtr(TypeAttribute), ByVal farPtr, LenB(TypeAttribute)
      
    ReDim aTYPEATTR(LenB(TypeAttribute))
      
    CopyMemory ByVal VarPtr(aTYPEATTR(0)), TypeAttribute, UBound(aTYPEATTR)
      
    CallFunction_COM ObjPtr(ITypeInfo), ITYP_ReleaseTypeAttr, vbEmpty, CC_STDCALL, farPtr
      
      
    For lFuncsCount = 0 To TypeAttribute.cFuncs - 1
        '            Call CallFunction_COM(ObjPtr(ITypeInfo), ITYP_GetFuncDesc, vbLong, CC_STDCALL, lFuncsCount, VarPtr(farPtr))
        CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetFuncDesc, vbLong, CC_STDCALL, lFuncsCount, VarPtr(farPtr)
            
        '            If farPtr = 0 Then MsgBox "error": Exit For
        If farPtr = 0 Then Exit For
            
        CopyMemory ByVal VarPtr(FunctionDescription), ByVal farPtr, LenB(FunctionDescription)
            
        ReDim aFUNCDESC(LenB(FunctionDescription))
            
        CopyMemory ByVal VarPtr(aFUNCDESC(0)), FunctionDescription, UBound(aFUNCDESC)
            
        CallFunction_COM ObjPtr(ITypeInfo), ITYP_ReleaseFuncDesc, vbEmpty, CC_STDCALL, farPtr
        CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(PropName), 0, 0, 0
        CallFunction_COM ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(PropName), 0, 0, 0
            
        '            Call CallFunction_COM(ObjPtr(ITypeInfo), ITYP_ReleaseFuncDesc, vbEmpty, CC_STDCALL, farPtr)
        '            Call CallFunction_COM(ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(sName), 0, 0, 0)
        '            Call CallFunction_COM(ObjPtr(ITypeInfo), ITYP_GetDocument, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(sName), 0, 0, 0)

        With FunctionDescription
            If FuncType Then
                If .INVOKEKIND = FuncType Then
                    Set PropInfo = New PropertyInfo
                    Set PropInfo.Owner = Target
                    PropInfo.PropertyName = PropName
                    PropInfo.CallType = .INVOKEKIND
                    PropList.Add PropInfo
                End If
            Else
                Set PropInfo = New PropertyInfo
                Set PropInfo.Owner = Target
                PropInfo.PropertyName = PropName
                PropInfo.CallType = .INVOKEKIND
                PropList.Add PropInfo
            End If
        End With
            
        PropName = vbNullString
    Next
    
    Set GetObjectMembers = PropList
End Function




#If Win64 Then
Private Function CallFunction_COM(ByVal InterfacePointer As LongLong, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    '     Private Function CallFunction_COM(ByVal InterfacePointer As LongLong, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    Dim vParamPtr() As LongLong
#Else
Private Function CallFunction_COM(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    '     Private Function CallFunction_COM(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    Dim vParamPtr() As Long
#End If
      
If InterfacePointer = 0& Or VTableOffset < 0& Then Exit Function
If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function
      
Dim pIndex    As Long, pCount As Long
Dim vParamType() As Integer
Dim vRtn      As Variant, vParams() As Variant
      
vParams() = FunctionParameters()
pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
If pCount = 0& Then
    ReDim vParamPtr(0 To 0)
    ReDim vParamType(0 To 0)
Else
    ReDim vParamPtr(0 To pCount - 1&)
    ReDim vParamType(0 To pCount - 1&)
    For pIndex = 0& To pCount - 1&
        vParamPtr(pIndex) = VarPtr(vParams(pIndex))
        vParamType(pIndex) = VarType(vParams(pIndex))
    Next
End If
      
pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, pCount, _
                      vParamType(0), vParamPtr(0), vRtn)
If pIndex = 0& Then
    CallFunction_COM = vRtn
Else
    SetLastError pIndex
End If

End Function


