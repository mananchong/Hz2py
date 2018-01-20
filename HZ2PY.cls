VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HZ2PY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'*
'* Module:          HzToPy
'* Update:          2011-09-23
'* Author:          tt.t
'*
'* Description:     将中文字符串转换为拼音，就这些。原先这里写了太多废话，删了。
'*
'* Theory：         原理依然是通过IFELanguage接口实现。
'*                  唯一需要解释的是如何解决多音字正确注音的问题。
'*                  IFELanguage接口是能够正确返回很多多音字拼音的，但多音字的读音只有特定词汇中
'*                  才能确认，因此在解析拼音时候不能把词拆成单字，否则多音字返回的拼音就很可能不对。
'*                  之前版本中就是因为把词拆开获取拼音导致多音字拼音错误。
'*                  这次的更新利用接口返回数据中标识每个拼音长度的数组实现了对返回拼音
'*                  的按字拆分，无需再把词拆成字获取单个字的拼音，从而解决了多音字问题。
'*                  需要说明的是，VB_MORRSLT结构就是MS文档中的MORRSLT结构，但是VBA自定义结构
'*                  无法实现不按4字节对齐，使得不得不修改MORRSLT的定义方式，能这样修改只能说运气不错，
'*                  因为被修改的部分刚好获取拼音用不到。
'*
'*
'* Histroy:
'*                  2011-09-23
'*                  ● 重写主要代码，支持多音字，提高了运行效率。
'*                  ● 取拼音首字时，ao, ai, ei, ou, er作为首字而不是原来的第一个字母。
'*                  ● 为函数增加了注音方式选择，hàn可以显示为han或han4。
'*                  ● 函数的使用与之前版本兼容，将模块中函数代码和HZ2PY类代码覆盖之前版本即可实现升级，无需修改文档中的公式。
'*                  2011-04-07
'*                  ● 更正CoTaskMemFree传递参数错误，消除了Win7等环境下崩溃。
'*                  2007-04-03
'*                  ● 更正redim时vba数组默认起始值错误。
'*                  2007-04-02
'*                  ● 最初版本，实现了由汉字获取拼音。
'*
'***************************************************************************

Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type VB_MORRSLT
    dwSize As Long          '4
    pwchOutput As Long      '4
    cchOutput As Integer    '2+(2)，VBA内存对齐闹得，折腾了好一阵才确认问题所在，唉
    Block1 As Long          '4
    pchInputPos As Long     '4
    pchOutputIdxWDD As Long '4
    pchReadIdxWDD As Long   '4
    paMonoRubyPos As Long   '4
    pWDD As Long            '4
    cWDD As Integer         '2
    pPrivate As Long        '4
    BLKBuff As Long         '4
End Type

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
        
Private Declare Function CLSIDFromString Lib "ole32.dll" _
        (ByVal lpszProgID As Long, pCLSID As GUID) As Long
        
Private Declare Function CoCreateInstance Lib "ole32" ( _
        rclsid As GUID, ByVal pUnkOuter As Long, _
        ByVal dwClsContext As Long, riid As GUID, _
        ByRef ppv As Long) As Long

Private Declare Function DispCallFunc Lib "oleaut32" _
        (ByVal pvInstance As Long, ByVal oVft As Long, _
        ByVal cc As Long, ByVal vtReturn As Integer, _
        ByVal cActuals As Long, prgvt As Integer, _
        prgpvarg As Long, pvargResult As Variant) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" (pv As Long)

Dim MSIME_GUID As GUID          'MSIME's GUID
Dim IFELanguage_GUID As GUID    'IFELanguage's GUID
Dim IFELanguage As Long         'Pointer to IFELanguage interface
Dim PinYinArray() As String
Dim HzLen As Integer

Dim pvSeperator As String
Dim pvUseSeperator As Boolean
Dim pvInitialOnly As Boolean
Dim pvOnlyOneChar As Boolean
Dim pvNonChnUseSep As Boolean

Public Function GetPinYin(HzStr As String) As String
    Dim i As Integer
    Dim Py As String
    Dim IsPy As Boolean
    
    GetPinYin = ""
    If IFELanguage = 0 Then
        GetPinYin = "未发现运行环境，请安装微软拼音2.0以上版本！"
        Exit Function
    End If
    If HzStr = "" Then Exit Function
    HzLen = Len(HzStr)
    Call IFELanguage_GetMorphResult(HzStr)
    For i = 1 To HzLen
        Py = PinYinArray(i)
        IsPy = Py <> ""
        If Not IsPy Then Py = Mid(HzStr, i, 1)
        If pvInitialOnly Then Py = GetInitial(Py)
        If pvOnlyOneChar Then Py = VBA.Left(Py, 1)
        GetPinYin = GetPinYin & Py & IIf(IsPy, pvSeperator, "")
    Next i
    If IsPy And pvSeperator <> "" Then GetPinYin = Left(GetPinYin, Len(GetPinYin) - 1)
End Function

Property Get Seperator() As String
    Seperator = pvSeperator
End Property

Property Let Seperator(Value As String)
    pvSeperator = Value
End Property

Property Get InitialOnly() As Boolean
    pvUseSeperator = pvInitialOnly
End Property

Property Let InitialOnly(Value As Boolean)
    pvInitialOnly = Value
End Property

Property Get OnlyOneChar() As Boolean
    pvUseSeperator = pvOnlyOneChar
End Property

Property Let OnlyOneChar(Value As Boolean)
    pvOnlyOneChar = Value
End Property

Public Function AdjustPhoneticNotation(Py As String, ByVal pn As Integer) As String
    Dim i As Integer
    Dim c As String
    
    If pn = -1 Then
        AdjustPhoneticNotation = Py
        Exit Function
    Else
        For i = 1 To Len(Py)
            c = VBA.Mid(Py, i, 1)
            Select Case Asc(c)
            Case VBA.Asc("ā") To VBA.Asc("à")
                c = "a" & IIf(pn = 0, "", (VBA.Asc(c) - VBA.Asc("ā") + 1))
            Case VBA.Asc("ē") To VBA.Asc("è")
                c = "e" & IIf(pn = 0, "", (VBA.Asc(c) - VBA.Asc("ē") + 1))
            Case VBA.Asc("ī") To VBA.Asc("ì")
                c = "i" & IIf(pn = 0, "", (VBA.Asc(c) - VBA.Asc("ī") + 1))
            Case VBA.Asc("ō") To VBA.Asc("ò")
                c = "o" & IIf(pn = 0, "", (VBA.Asc(c) - VBA.Asc("ō") + 1))
            Case VBA.Asc("ū") To VBA.Asc("ù")
                c = "u" & IIf(pn = 0, "", (VBA.Asc(c) - VBA.Asc("ū") + 1))
            Case VBA.Asc("ǖ") To VBA.Asc("ǜ")
                c = "u" & IIf(pn = 0, "", (VBA.Asc(c) - VBA.Asc("ǖ") + 1))
            Case VBA.Asc("ü")
                c = "u"
            Case VBA.Asc("ɡ")
                c = "g"
            End Select
            AdjustPhoneticNotation = AdjustPhoneticNotation & c
        Next i
    End If
End Function

Private Function GetInitial(Py As String) As String
    GetInitial = VBA.Mid(Py, 1, 2)
    Select Case AdjustPhoneticNotation(GetInitial, 0)
    Case "ch", "sh", "zh", "ao", "ai", "ei", "ou", "er"
    Case Else
        GetInitial = VBA.Left(GetInitial, 1)
    End Select
End Function

Private Function IFELanguage_GetMorphResult(HzStr As String) As String
    Dim ret As Variant
    Dim pArgs(0 To 5) As Long
    Dim vt(0 To 5) As Integer
    Dim Args(0 To 5) As Long
    Dim ResultPtr As Long
    Dim TinyM As VB_MORRSLT
    Dim Py() As Byte
    Dim i As Integer
    Dim j As Integer
    Dim PinyinIndexArray() As Integer
        
    IFELanguage_GetMorphResult = ""
    If IFELanguage = 0 Then Exit Function
    
    Args(0) = &H30000
    Args(1) = &H40000100
    Args(2) = Len(HzStr)
    Args(3) = StrPtr(HzStr)
    Args(4) = 0
    Args(5) = VarPtr(ResultPtr)
        
    For i = 0 To 5
        vt(i) = vbLong
        pArgs(i) = VarPtr(Args(i)) - 8
    Next
    
    Call DispCallFunc(IFELanguage, 20, 4, vbLong, 6, vt(0), pArgs(0), ret)
    Call MoveMemory(TinyM, ByVal ResultPtr, Len(TinyM))

    ReDim PinyinIndexArray(0 To HzLen - 1)
    ReDim PinYinArray(1 To HzLen)
    If TinyM.cchOutput > 0 Then
        ReDim Py(0 To TinyM.cchOutput * 2 - 1)
        Call MoveMemory(Py(0), ByVal TinyM.pwchOutput, TinyM.cchOutput * 2)
        IFELanguage_GetMorphResult = Py
        Call MoveMemory(PinyinIndexArray(0), ByVal TinyM.paMonoRubyPos + 2, HzLen * 2)
        j = 0
        For i = 0 To HzLen - 1
            PinYinArray(i + 1) = VBA.Mid(IFELanguage_GetMorphResult, j + 1, PinyinIndexArray(i) - j)
            j = PinyinIndexArray(i)
        Next i
    End If
    
    Call CoTaskMemFree(ByVal ResultPtr)
End Function

Private Sub IFELanguage_Open()
    Dim ret As Variant
    
    Call DispCallFunc(IFELanguage, 4, 4, vbLong, 0, 0, 0, ret)
    Call DispCallFunc(IFELanguage, 12, 4, vbLong, 0, 0, 0, ret)
End Sub

Private Sub IFELanguage_Close()
    Dim ret As Variant
    
    If IFELanguage = 0 Then Exit Sub
    Call DispCallFunc(IFELanguage, 8, 4, vbLong, 0, 0, 0, ret)
    Call DispCallFunc(IFELanguage, 16, 4, vbLong, 0, 0, 0, ret)
End Sub

Private Function GenerateGUID()
    Dim Rlt As Long
    
    'MSIME.China GUID = "{E4288337-873B-11D1-BAA0-00AA00BBB8C0}"
    Rlt = CLSIDFromString(StrPtr("MSIME.China"), MSIME_GUID)
    'IFELanguage GUID = "{019F7152-E6DB-11d0-83C3-00C04FDDB82E}"
    With IFELanguage_GUID
        .Data1 = &H19F7152
        .Data2 = &HE6DB
        .Data3 = &H11D0
        .Data4(0) = &H83
        .Data4(1) = &HC3
        .Data4(2) = &H0
        .Data4(3) = &HC0
        .Data4(4) = &H4F
        .Data4(5) = &HDD
        .Data4(6) = &HB8
        .Data4(7) = &H2E
    End With
    GenerateGUID = Rlt = 0
End Function

Private Sub Class_Initialize()
    IFELanguage = 0
    pvSeperator = ""
    GenerateGUID
    If CoCreateInstance(MSIME_GUID, 0, 1, IFELanguage_GUID, IFELanguage) = 0 Then Call IFELanguage_Open
End Sub

Private Sub Class_Terminate()
    If IFELanguage <> 0 Then Call IFELanguage_Close
End Sub
