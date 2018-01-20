Attribute VB_Name = "NewMacros"
Sub AddPinYin()


    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With Selection
    Dim s As String
    s = .Text
    .Text = "(  )"
    .Range.PhoneticGuide Text:=HzToPy(s), _
    Alignment:=wdPhoneticGuideAlignmentCenter, _
    FontSize:=10   '<-------这个值是修改拼音字体大小的'
    
    
    End With

End Sub


Public Function HzToPy(Hz As String, _
        Optional Sep As String = "", _
        Optional NotationType As Integer = -1, _
        Optional ShowInitialOnly As Boolean = False, _
        Optional ShowOnlyOneChar As Boolean = False) As String
        
    Dim hp As HZ2PY
    
    Set hp = New HZ2PY          '创建类
    hp.Seperator = Sep
    hp.InitialOnly = ShowInitialOnly
    hp.OnlyOneChar = ShowOnlyOneChar
    HzToPy = hp.GetPinYin(Hz)
    HzToPy = hp.AdjustPhoneticNotation(HzToPy, NotationType)
    Set hp = Nothing            '释放类
End Function

