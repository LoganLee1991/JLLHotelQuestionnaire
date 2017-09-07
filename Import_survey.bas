Attribute VB_Name = "Import_survey"
Public Sub getSummary()
    Dim Title As String
    Dim Y As Integer
    Dim filt As String
    Dim filterIndex As Integer
    Dim fileName As Variant
    Title = "Import survey"
    'Select the input files
    msg = "Please select the survey files you want to import"
    Y = MsgBox(msg, vbOKCancel, Title)
    If Y = 2 Then Exit Sub
            
    filt = "All Excel Files (*.xls), *.xls,"
    filterIndex = 1
    
    fileName = Application.GetOpenFilename _
        (FileFilter:=filt, _
        filterIndex:=filterIndex, _
        Title:=Title, _
        MultiSelect:=True)
        
    If Not IsArray(fileName) Then
        MsgBox "No file were selected."
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    Dim i As Variant
    For i = LBound(fileName) To UBound(fileName)
        Msg1 = Msg1 & fileName(i) & vbCrLf
    Next i
    'MsgBox "You selected:" & vbCrLf & Msg1 & " " & i
    
   ' On Error Resume Next
    
   ' Application.DisplayAlerts = False
    
    'Copy survey to the summary file
    Dim surveybookName As String
    Dim summarybookName As String
    Dim surveyBook As Workbook
    Dim summaryBook As Workbook
    Dim surveySheet As Worksheet
    Dim summarySheet As Worksheet
    
    Set summaryBook = ActiveWorkbook
    Set summarySheet = ActiveWorkbook.Worksheets("Summary")
    'MsgBox "summarybook name:" & summarybookName
    
    Dim filePath As String
    Dim x As Variant
    
    For x = LBound(fileName) To UBound(fileName)
            
        Workbooks.Open fileName:=fileName(x), UpdateLinks:=0
        filePath = fileName(x) 'get file's path
        surveybookName = ActiveWorkbook.Name
        Set surveyBook = Workbooks(surveybookName)
        Set surveySheet = Workbooks(surveybookName).Worksheets("Survey")
        
        summarySheet.Activate
        Range("a65536").End(xlUp).Select
        Dim benchCell As Range
        Set benchCell = ActiveCell.Offset(1, 0) '每个样本的编号所在的单元格，以此作为定位的基准单元格
        
        If ActiveCell.Value = "编号" Then
        benchCell.Value = 1
        Else
        benchCell.Value = ActiveCell.Value + 1
        End If
        '-------------------------------------------------------------
        'Part1:饭店和业主信息
        '-------------------------------------------------------------
        surveySheet.Range("c8:c17").Copy
        summarySheet.Activate
        benchCell.Offset(0, 1).Select
        Selection.PasteSpecial Paste:=xlValues, Transpose:=True
        '-------------------------------------------------------------
        benchCell.Offset(0, 11) = Workbooks(surveybookName).Worksheets("Survey").OptionButton1.Value
        benchCell.Offset(0, 12) = Workbooks(surveybookName).Worksheets("Survey").k18.Value
        benchCell.Offset(0, 13) = Workbooks(surveybookName).Worksheets("Survey").optionbutton2.Value
        
        Dim explain1 As String
        explain1 = surveySheet.Range("d18").Value
        benchCell.Offset(0, 14) = Right(explain1, Len(explain1) - Len("请说明:"))
        '-------------------------------------------------------------
        surveySheet.Range("c19:c25").Copy
        summarySheet.Activate
        benchCell.Offset(0, 15).Select
        Selection.PasteSpecial Paste:=xlValues, Transpose:=True
        '-------------------------------------------------------------
        benchCell.Offset(0, 22) = Workbooks(surveybookName).Worksheets("Survey").OptionButton3.Value
        benchCell.Offset(0, 23) = Workbooks(surveybookName).Worksheets("Survey").optionbutton4.Value
        '-------------------------------------------------------------
        benchCell.Offset(0, 24) = surveySheet.Range("c27").Value
        '-------------------------------------------------------------
        benchCell.Offset(0, 25) = Workbooks(surveybookName).Worksheets("Survey").optionbutton87.Value
        benchCell.Offset(0, 26) = Workbooks(surveybookName).Worksheets("Survey").optionbutton88.Value
        benchCell.Offset(0, 27) = Workbooks(surveybookName).Worksheets("Survey").optionbutton89.Value
        benchCell.Offset(0, 28) = Workbooks(surveybookName).Worksheets("Survey").optionbutton90.Value
        benchCell.Offset(0, 29) = Workbooks(surveybookName).Worksheets("Survey").OptionButton91.Value
        '-------------------------------------------------------------
        benchCell.Offset(0, 30) = surveySheet.Range("c30").Value
        '-------------------------------------------------------------
        benchCell.Offset(0, 31) = Workbooks(surveybookName).Worksheets("Survey").OptionButton10.Value
        benchCell.Offset(0, 32) = Workbooks(surveybookName).Worksheets("Survey").optionbutton11.Value
        benchCell.Offset(0, 33) = Workbooks(surveybookName).Worksheets("Survey").OptionButton12.Value
        benchCell.Offset(0, 34) = Workbooks(surveybookName).Worksheets("Survey").optionbutton13.Value
        benchCell.Offset(0, 35) = Workbooks(surveybookName).Worksheets("Survey").OptionButton14.Value
        
        Dim explain2 As String
        explain2 = surveySheet.Range("d32").Value
        benchCell.Offset(0, 36) = Right(explain2, Len(explain2) - Len("请说明:"))
        '-------------------------------------------------------------
        surveySheet.Range("c33:c35").Copy
        summarySheet.Activate
        benchCell.Offset(0, 37).Select
        Selection.PasteSpecial Paste:=xlValues, Transpose:=True
        '-------------------------------------------------------------
        benchCell.Offset(0, 40) = Workbooks(surveybookName).Worksheets("Survey").OptionButton5.Value
        benchCell.Offset(0, 41) = Workbooks(surveybookName).Worksheets("Survey").optionbutton6.Value
        benchCell.Offset(0, 42) = Workbooks(surveybookName).Worksheets("Survey").OptionButton7.Value
        benchCell.Offset(0, 43) = Workbooks(surveybookName).Worksheets("Survey").optionbutton8.Value
        benchCell.Offset(0, 44) = Workbooks(surveybookName).Worksheets("Survey").OptionButton9.Value
        '-------------------------------------------------------------
        
        '-------------------------------------------------------------
        'Part2：饭店业绩统计
        '-------------------------------------------------------------
        surveySheet.Range("c43:c55").Copy
        summarySheet.Activate
        benchCell.Offset(0, 45).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("d43:d55").Copy
        summarySheet.Activate
        benchCell.Offset(0, 58).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("e43:e55").Copy
        summarySheet.Activate
        benchCell.Offset(0, 71).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        '-------------------------------------------------------------
        
        '-------------------------------------------------------------
        'Part3:饭店客源构成
        '-------------------------------------------------------------
        surveySheet.Range("c62:c68").Copy
        summarySheet.Activate
        benchCell.Offset(0, 84).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("d62:d68").Copy
        summarySheet.Activate
        benchCell.Offset(0, 91).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("e62:e67").Copy
        summarySheet.Activate
        benchCell.Offset(0, 98).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        Dim explain3 As String
        explain3 = surveySheet.Range("b67").Value
        benchCell.Offset(0, 104) = Right(explain3, Len(explain3) - Len("请说明其他客源："))
        '-------------------------------------------------------------
        
        '-------------------------------------------------------------
        'Part4:饭店收入构成
        '-------------------------------------------------------------
        surveySheet.Range("c75:c87").Copy
        summarySheet.Activate
        benchCell.Offset(0, 105).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("d75:d87").Copy
        summarySheet.Activate
        benchCell.Offset(0, 118).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("e75:e87").Copy
        summarySheet.Activate
        benchCell.Offset(0, 131).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
                
        benchCell.Offset(0, 144) = Workbooks(surveybookName).Worksheets("Survey").OptionButton15.Value
        benchCell.Offset(0, 145) = Workbooks(surveybookName).Worksheets("Survey").optionbutton16.Value
        benchCell.Offset(0, 146) = Workbooks(surveybookName).Worksheets("Survey").optionbutton17.Value
        benchCell.Offset(0, 147) = Workbooks(surveybookName).Worksheets("Survey").optionbutton69.Value
        benchCell.Offset(0, 148) = Workbooks(surveybookName).Worksheets("Survey").optionbutton70.Value
        benchCell.Offset(0, 149) = Workbooks(surveybookName).Worksheets("Survey").optionbutton71.Value
        
        benchCell.Offset(0, 150) = Workbooks(surveybookName).Worksheets("Survey").optionbutton18.Value
        benchCell.Offset(0, 151) = Workbooks(surveybookName).Worksheets("Survey").optionbutton19.Value
        benchCell.Offset(0, 152) = Workbooks(surveybookName).Worksheets("Survey").optionbutton20.Value
        benchCell.Offset(0, 153) = Workbooks(surveybookName).Worksheets("Survey").optionbutton92.Value
        benchCell.Offset(0, 154) = Workbooks(surveybookName).Worksheets("Survey").optionbutton93.Value
        benchCell.Offset(0, 155) = Workbooks(surveybookName).Worksheets("Survey").optionbutton94.Value
        
        benchCell.Offset(0, 156) = Workbooks(surveybookName).Worksheets("Survey").optionbutton24.Value
        benchCell.Offset(0, 157) = Workbooks(surveybookName).Worksheets("Survey").optionbutton25.Value
        benchCell.Offset(0, 158) = Workbooks(surveybookName).Worksheets("Survey").optionbutton26.Value
        benchCell.Offset(0, 159) = Workbooks(surveybookName).Worksheets("Survey").optionbutton98.Value
        benchCell.Offset(0, 160) = Workbooks(surveybookName).Worksheets("Survey").optionbutton99.Value
        benchCell.Offset(0, 161) = Workbooks(surveybookName).Worksheets("Survey").optionbutton100.Value
        
        benchCell.Offset(0, 162) = Workbooks(surveybookName).Worksheets("Survey").optionbutton27.Value
        benchCell.Offset(0, 163) = Workbooks(surveybookName).Worksheets("Survey").optionbutton28.Value
        benchCell.Offset(0, 164) = Workbooks(surveybookName).Worksheets("Survey").optionbutton29.Value
        benchCell.Offset(0, 165) = Workbooks(surveybookName).Worksheets("Survey").optionbutton101.Value
        benchCell.Offset(0, 166) = Workbooks(surveybookName).Worksheets("Survey").optionbutton102.Value
        benchCell.Offset(0, 167) = Workbooks(surveybookName).Worksheets("Survey").optionbutton103.Value
        
        benchCell.Offset(0, 168) = Workbooks(surveybookName).Worksheets("Survey").optionbutton30.Value
        benchCell.Offset(0, 169) = Workbooks(surveybookName).Worksheets("Survey").optionbutton31.Value
        benchCell.Offset(0, 170) = Workbooks(surveybookName).Worksheets("Survey").optionbutton32.Value
        benchCell.Offset(0, 171) = Workbooks(surveybookName).Worksheets("Survey").optionbutton104.Value
        benchCell.Offset(0, 172) = Workbooks(surveybookName).Worksheets("Survey").optionbutton105.Value
        benchCell.Offset(0, 173) = Workbooks(surveybookName).Worksheets("Survey").optionbutton106.Value
        
        benchCell.Offset(0, 174) = Workbooks(surveybookName).Worksheets("Survey").optionbutton33.Value
        benchCell.Offset(0, 175) = Workbooks(surveybookName).Worksheets("Survey").optionbutton34.Value
        benchCell.Offset(0, 176) = Workbooks(surveybookName).Worksheets("Survey").optionbutton35.Value
        benchCell.Offset(0, 177) = Workbooks(surveybookName).Worksheets("Survey").optionbutton107.Value
        benchCell.Offset(0, 178) = Workbooks(surveybookName).Worksheets("Survey").optionbutton108.Value
        benchCell.Offset(0, 179) = Workbooks(surveybookName).Worksheets("Survey").optionbutton109.Value
        
        benchCell.Offset(0, 180) = Workbooks(surveybookName).Worksheets("Survey").optionbutton36.Value
        benchCell.Offset(0, 181) = Workbooks(surveybookName).Worksheets("Survey").optionbutton37.Value
        benchCell.Offset(0, 182) = Workbooks(surveybookName).Worksheets("Survey").optionbutton38.Value
        benchCell.Offset(0, 183) = Workbooks(surveybookName).Worksheets("Survey").optionbutton110.Value
        benchCell.Offset(0, 184) = Workbooks(surveybookName).Worksheets("Survey").optionbutton111.Value
        benchCell.Offset(0, 185) = Workbooks(surveybookName).Worksheets("Survey").optionbutton112.Value
        
        benchCell.Offset(0, 186) = Workbooks(surveybookName).Worksheets("Survey").optionbutton39.Value
        benchCell.Offset(0, 187) = Workbooks(surveybookName).Worksheets("Survey").optionbutton40.Value
        benchCell.Offset(0, 188) = Workbooks(surveybookName).Worksheets("Survey").optionbutton41.Value
        benchCell.Offset(0, 189) = Workbooks(surveybookName).Worksheets("Survey").optionbutton113.Value
        benchCell.Offset(0, 190) = Workbooks(surveybookName).Worksheets("Survey").optionbutton114.Value
        benchCell.Offset(0, 191) = Workbooks(surveybookName).Worksheets("Survey").optionbutton115.Value
        
        benchCell.Offset(0, 192) = Workbooks(surveybookName).Worksheets("Survey").optionbutton42.Value
        benchCell.Offset(0, 193) = Workbooks(surveybookName).Worksheets("Survey").optionbutton43.Value
        benchCell.Offset(0, 194) = Workbooks(surveybookName).Worksheets("Survey").optionbutton44.Value
        benchCell.Offset(0, 195) = Workbooks(surveybookName).Worksheets("Survey").optionbutton116.Value
        benchCell.Offset(0, 196) = Workbooks(surveybookName).Worksheets("Survey").optionbutton117.Value
        benchCell.Offset(0, 197) = Workbooks(surveybookName).Worksheets("Survey").optionbutton118.Value
        
        benchCell.Offset(0, 198) = Workbooks(surveybookName).Worksheets("Survey").optionbutton45.Value
        benchCell.Offset(0, 199) = Workbooks(surveybookName).Worksheets("Survey").optionbutton46.Value
        benchCell.Offset(0, 200) = Workbooks(surveybookName).Worksheets("Survey").optionbutton47.Value
        benchCell.Offset(0, 201) = Workbooks(surveybookName).Worksheets("Survey").optionbutton119.Value
        benchCell.Offset(0, 202) = Workbooks(surveybookName).Worksheets("Survey").optionbutton120.Value
        benchCell.Offset(0, 203) = Workbooks(surveybookName).Worksheets("Survey").optionbutton121.Value
        
        benchCell.Offset(0, 204) = Workbooks(surveybookName).Worksheets("Survey").optionbutton60.Value
        benchCell.Offset(0, 205) = Workbooks(surveybookName).Worksheets("Survey").optionbutton61.Value
        benchCell.Offset(0, 206) = Workbooks(surveybookName).Worksheets("Survey").optionbutton62.Value
        benchCell.Offset(0, 207) = Workbooks(surveybookName).Worksheets("Survey").optionbutton122.Value
        benchCell.Offset(0, 208) = Workbooks(surveybookName).Worksheets("Survey").optionbutton123.Value
        benchCell.Offset(0, 209) = Workbooks(surveybookName).Worksheets("Survey").optionbutton124.Value
        
        benchCell.Offset(0, 210) = Workbooks(surveybookName).Worksheets("Survey").optionbutton63.Value
        benchCell.Offset(0, 211) = Workbooks(surveybookName).Worksheets("Survey").optionbutton64.Value
        benchCell.Offset(0, 212) = Workbooks(surveybookName).Worksheets("Survey").optionbutton65.Value
        benchCell.Offset(0, 213) = Workbooks(surveybookName).Worksheets("Survey").optionbutton125.Value
        benchCell.Offset(0, 214) = Workbooks(surveybookName).Worksheets("Survey").optionbutton126.Value
        benchCell.Offset(0, 215) = Workbooks(surveybookName).Worksheets("Survey").optionbutton127.Value
        
        benchCell.Offset(0, 216) = Workbooks(surveybookName).Worksheets("Survey").optionbutton66.Value
        benchCell.Offset(0, 217) = Workbooks(surveybookName).Worksheets("Survey").optionbutton67.Value
        benchCell.Offset(0, 218) = Workbooks(surveybookName).Worksheets("Survey").optionbutton68.Value
        benchCell.Offset(0, 219) = Workbooks(surveybookName).Worksheets("Survey").optionbutton128.Value
        benchCell.Offset(0, 220) = Workbooks(surveybookName).Worksheets("Survey").optionbutton129.Value
        benchCell.Offset(0, 221) = Workbooks(surveybookName).Worksheets("Survey").optionbutton130.Value

        Dim explain4 As String
        explain4 = surveySheet.Range("b87").Value
        benchCell.Offset(0, 222) = Right(explain4, Len(explain4) - Len("请说明其他收入："))
        '-------------------------------------------------------------
        
        '-------------------------------------------------------------
        'Part5:饭店会议设施使用情况
        '-------------------------------------------------------------
        surveySheet.Range("c94:c99").Copy
        summarySheet.Activate
        benchCell.Offset(0, 223).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("d94:d99").Copy
        summarySheet.Activate
        benchCell.Offset(0, 229).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        benchCell.Offset(0, 235) = Workbooks(surveybookName).Worksheets("Survey").optionbutton131.Value
        benchCell.Offset(0, 236) = Workbooks(surveybookName).Worksheets("Survey").optionbutton132.Value
        benchCell.Offset(0, 237) = Workbooks(surveybookName).Worksheets("Survey").optionbutton133.Value
        benchCell.Offset(0, 238) = Workbooks(surveybookName).Worksheets("Survey").optionbutton134.Value
        benchCell.Offset(0, 239) = Workbooks(surveybookName).Worksheets("Survey").optionbutton135.Value
        benchCell.Offset(0, 240) = Workbooks(surveybookName).Worksheets("Survey").optionbutton136.Value
        
        benchCell.Offset(0, 241) = Workbooks(surveybookName).Worksheets("Survey").optionbutton137.Value
        benchCell.Offset(0, 242) = Workbooks(surveybookName).Worksheets("Survey").optionbutton138.Value
        benchCell.Offset(0, 243) = Workbooks(surveybookName).Worksheets("Survey").optionbutton139.Value
        benchCell.Offset(0, 244) = Workbooks(surveybookName).Worksheets("Survey").optionbutton140.Value
        benchCell.Offset(0, 245) = Workbooks(surveybookName).Worksheets("Survey").optionbutton141.Value
        benchCell.Offset(0, 246) = Workbooks(surveybookName).Worksheets("Survey").optionbutton142.Value
        
        benchCell.Offset(0, 247) = Workbooks(surveybookName).Worksheets("Survey").optionbutton143.Value
        benchCell.Offset(0, 248) = Workbooks(surveybookName).Worksheets("Survey").optionbutton144.Value
        benchCell.Offset(0, 249) = Workbooks(surveybookName).Worksheets("Survey").optionbutton145.Value
        benchCell.Offset(0, 250) = Workbooks(surveybookName).Worksheets("Survey").optionbutton146.Value
        benchCell.Offset(0, 251) = Workbooks(surveybookName).Worksheets("Survey").optionbutton147.Value
        benchCell.Offset(0, 252) = Workbooks(surveybookName).Worksheets("Survey").optionbutton148.Value
        
        benchCell.Offset(0, 253) = Workbooks(surveybookName).Worksheets("Survey").optionbutton149.Value
        benchCell.Offset(0, 254) = Workbooks(surveybookName).Worksheets("Survey").optionbutton150.Value
        benchCell.Offset(0, 255) = Workbooks(surveybookName).Worksheets("Survey").optionbutton151.Value
        benchCell.Offset(0, 256) = Workbooks(surveybookName).Worksheets("Survey").optionbutton152.Value
        benchCell.Offset(0, 257) = Workbooks(surveybookName).Worksheets("Survey").optionbutton153.Value
        benchCell.Offset(0, 258) = Workbooks(surveybookName).Worksheets("Survey").optionbutton154.Value
        
        benchCell.Offset(0, 259) = Workbooks(surveybookName).Worksheets("Survey").optionbutton155.Value
        benchCell.Offset(0, 260) = Workbooks(surveybookName).Worksheets("Survey").optionbutton156.Value
        benchCell.Offset(0, 261) = Workbooks(surveybookName).Worksheets("Survey").optionbutton157.Value
        benchCell.Offset(0, 262) = Workbooks(surveybookName).Worksheets("Survey").optionbutton158.Value
        benchCell.Offset(0, 263) = Workbooks(surveybookName).Worksheets("Survey").optionbutton159.Value
        benchCell.Offset(0, 264) = Workbooks(surveybookName).Worksheets("Survey").optionbutton160.Value
        
        benchCell.Offset(0, 265) = Workbooks(surveybookName).Worksheets("Survey").optionbutton161.Value
        benchCell.Offset(0, 266) = Workbooks(surveybookName).Worksheets("Survey").optionbutton162.Value
        benchCell.Offset(0, 267) = Workbooks(surveybookName).Worksheets("Survey").optionbutton163.Value
        benchCell.Offset(0, 268) = Workbooks(surveybookName).Worksheets("Survey").optionbutton164.Value
        benchCell.Offset(0, 269) = Workbooks(surveybookName).Worksheets("Survey").optionbutton165.Value
        benchCell.Offset(0, 270) = Workbooks(surveybookName).Worksheets("Survey").optionbutton166.Value
       
        Dim explain5 As String
        explain5 = surveySheet.Range("b99").Value
        benchCell.Offset(0, 271) = Right(explain5, Len(explain5) - Len("请说明其他会议设施："))
        '-------------------------------------------------------------
        'Part6:饭店销售渠道构成
        '-------------------------------------------------------------
        surveySheet.Range("c105:c112").Copy
        summarySheet.Activate
        benchCell.Offset(0, 272).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("d105:d112").Copy
        summarySheet.Activate
        benchCell.Offset(0, 280).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        benchCell.Offset(0, 288) = Workbooks(surveybookName).Worksheets("Survey").optionbutton167.Value
        benchCell.Offset(0, 289) = Workbooks(surveybookName).Worksheets("Survey").optionbutton168.Value
        benchCell.Offset(0, 290) = Workbooks(surveybookName).Worksheets("Survey").optionbutton169.Value
        benchCell.Offset(0, 291) = Workbooks(surveybookName).Worksheets("Survey").optionbutton170.Value
        benchCell.Offset(0, 292) = Workbooks(surveybookName).Worksheets("Survey").optionbutton171.Value
        benchCell.Offset(0, 293) = Workbooks(surveybookName).Worksheets("Survey").optionbutton172.Value
        
        benchCell.Offset(0, 294) = Workbooks(surveybookName).Worksheets("Survey").optionbutton173.Value
        benchCell.Offset(0, 295) = Workbooks(surveybookName).Worksheets("Survey").optionbutton174.Value
        benchCell.Offset(0, 296) = Workbooks(surveybookName).Worksheets("Survey").optionbutton175.Value
        benchCell.Offset(0, 297) = Workbooks(surveybookName).Worksheets("Survey").optionbutton176.Value
        benchCell.Offset(0, 298) = Workbooks(surveybookName).Worksheets("Survey").optionbutton177.Value
        benchCell.Offset(0, 299) = Workbooks(surveybookName).Worksheets("Survey").optionbutton178.Value
        
        benchCell.Offset(0, 300) = Workbooks(surveybookName).Worksheets("Survey").optionbutton179.Value
        benchCell.Offset(0, 301) = Workbooks(surveybookName).Worksheets("Survey").optionbutton180.Value
        benchCell.Offset(0, 302) = Workbooks(surveybookName).Worksheets("Survey").optionbutton181.Value
        benchCell.Offset(0, 303) = Workbooks(surveybookName).Worksheets("Survey").optionbutton182.Value
        benchCell.Offset(0, 304) = Workbooks(surveybookName).Worksheets("Survey").optionbutton183.Value
        benchCell.Offset(0, 305) = Workbooks(surveybookName).Worksheets("Survey").optionbutton184.Value
        
        benchCell.Offset(0, 306) = Workbooks(surveybookName).Worksheets("Survey").optionbutton185.Value
        benchCell.Offset(0, 307) = Workbooks(surveybookName).Worksheets("Survey").optionbutton186.Value
        benchCell.Offset(0, 308) = Workbooks(surveybookName).Worksheets("Survey").optionbutton187.Value
        benchCell.Offset(0, 309) = Workbooks(surveybookName).Worksheets("Survey").optionbutton188.Value
        benchCell.Offset(0, 310) = Workbooks(surveybookName).Worksheets("Survey").optionbutton189.Value
        benchCell.Offset(0, 311) = Workbooks(surveybookName).Worksheets("Survey").optionbutton190.Value
        
        benchCell.Offset(0, 312) = Workbooks(surveybookName).Worksheets("Survey").optionbutton191.Value
        benchCell.Offset(0, 313) = Workbooks(surveybookName).Worksheets("Survey").optionbutton192.Value
        benchCell.Offset(0, 314) = Workbooks(surveybookName).Worksheets("Survey").optionbutton193.Value
        benchCell.Offset(0, 315) = Workbooks(surveybookName).Worksheets("Survey").optionbutton194.Value
        benchCell.Offset(0, 316) = Workbooks(surveybookName).Worksheets("Survey").optionbutton195.Value
        benchCell.Offset(0, 317) = Workbooks(surveybookName).Worksheets("Survey").optionbutton196.Value
        
        benchCell.Offset(0, 318) = Workbooks(surveybookName).Worksheets("Survey").optionbutton197.Value
        benchCell.Offset(0, 319) = Workbooks(surveybookName).Worksheets("Survey").optionbutton198.Value
        benchCell.Offset(0, 320) = Workbooks(surveybookName).Worksheets("Survey").optionbutton199.Value
        benchCell.Offset(0, 321) = Workbooks(surveybookName).Worksheets("Survey").optionbutton200.Value
        benchCell.Offset(0, 322) = Workbooks(surveybookName).Worksheets("Survey").optionbutton201.Value
        benchCell.Offset(0, 323) = Workbooks(surveybookName).Worksheets("Survey").optionbutton202.Value
        
        benchCell.Offset(0, 324) = Workbooks(surveybookName).Worksheets("Survey").optionbutton203.Value
        benchCell.Offset(0, 325) = Workbooks(surveybookName).Worksheets("Survey").optionbutton204.Value
        benchCell.Offset(0, 326) = Workbooks(surveybookName).Worksheets("Survey").optionbutton205.Value
        benchCell.Offset(0, 327) = Workbooks(surveybookName).Worksheets("Survey").optionbutton206.Value
        benchCell.Offset(0, 328) = Workbooks(surveybookName).Worksheets("Survey").optionbutton207.Value
        benchCell.Offset(0, 329) = Workbooks(surveybookName).Worksheets("Survey").optionbutton208.Value
        
        benchCell.Offset(0, 330) = Workbooks(surveybookName).Worksheets("Survey").optionbutton209.Value
        benchCell.Offset(0, 331) = Workbooks(surveybookName).Worksheets("Survey").optionbutton210.Value
        benchCell.Offset(0, 332) = Workbooks(surveybookName).Worksheets("Survey").optionbutton211.Value
        benchCell.Offset(0, 333) = Workbooks(surveybookName).Worksheets("Survey").optionbutton212.Value
        benchCell.Offset(0, 334) = Workbooks(surveybookName).Worksheets("Survey").optionbutton213.Value
        benchCell.Offset(0, 335) = Workbooks(surveybookName).Worksheets("Survey").optionbutton214.Value

        Dim explain6 As String
        explain6 = surveySheet.Range("b111").Value
        benchCell.Offset(0, 336) = Right(explain6, Len(explain6) - Len("请说明其他："))
        '-------------------------------------------------------------
        
        '-------------------------------------------------------------
        'Part7：年度预算与资产管理
        '-------------------------------------------------------------
        benchCell.Offset(0, 337) = surveySheet.Range("c117")
        benchCell.Offset(0, 340) = surveySheet.Range("d117")
        
        benchCell.Offset(0, 338) = Workbooks(surveybookName).Worksheets("Survey").optionbutton215.Value
        benchCell.Offset(0, 339) = Workbooks(surveybookName).Worksheets("Survey").optionbutton216.Value
        
        benchCell.Offset(0, 341) = Workbooks(surveybookName).Worksheets("Survey").optionbutton219.Value
        benchCell.Offset(0, 342) = Workbooks(surveybookName).Worksheets("Survey").optionbutton225.Value
        
        benchCell.Offset(0, 343) = Workbooks(surveybookName).Worksheets("Survey").optionbutton217.Value
        benchCell.Offset(0, 344) = Workbooks(surveybookName).Worksheets("Survey").optionbutton218.Value
        
        benchCell.Offset(0, 345) = Workbooks(surveybookName).Worksheets("Survey").optionbutton235.Value
        benchCell.Offset(0, 346) = Workbooks(surveybookName).Worksheets("Survey").optionbutton236.Value
        benchCell.Offset(0, 347) = Workbooks(surveybookName).Worksheets("Survey").optionbutton238.Value
        benchCell.Offset(0, 348) = Workbooks(surveybookName).Worksheets("Survey").optionbutton237.Value
        benchCell.Offset(0, 349) = Workbooks(surveybookName).Worksheets("Survey").optionbutton239.Value
        
        Dim explain7 As String
        explain7 = surveySheet.Range("f123").Value
        benchCell.Offset(0, 350) = Right(explain7, Len(explain7) - Len("请说明："))
        
        Dim explain8 As String
        explain8 = surveySheet.Range("f124").Value
        benchCell.Offset(0, 351) = Right(explain8, Len(explain8) - Len("请说明："))

        '-------------------------------------------------------------
        
        '-------------------------------------------------------------
        'Part8:饭店成本构成
        '-------------------------------------------------------------
        surveySheet.Range("c128:c136").Copy
        summarySheet.Activate
        benchCell.Offset(0, 352).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("d128:d136").Copy
        summarySheet.Activate
        benchCell.Offset(0, 361).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        surveySheet.Range("e128:e136").Copy
        summarySheet.Activate
        benchCell.Offset(0, 370).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Transpose:=True
        
        Dim explain9 As String
        explain9 = surveySheet.Range("B135").Value
        benchCell.Offset(0, 379) = Right(explain9, Len(explain9) - Len("请说明其他成本："))
        
        benchCell.Offset(0, 380) = Workbooks(surveybookName).Worksheets("Survey").optionbutton230.Value
        benchCell.Offset(0, 381) = Workbooks(surveybookName).Worksheets("Survey").optionbutton231.Value
        benchCell.Offset(0, 382) = Workbooks(surveybookName).Worksheets("Survey").optionbutton232.Value
        benchCell.Offset(0, 383) = Workbooks(surveybookName).Worksheets("Survey").optionbutton233.Value
        
        Dim explain10 As String
        explain10 = surveySheet.Range("E139").Value
        benchCell.Offset(0, 384) = Right(explain10, Len(explain10) - Len("请说明:"))
        
        Dim explain11 As String '????此处问卷存在格式问题！！！！！！！！！！！！！！！
        explain11 = surveySheet.Range("C140").Value
        benchCell.Offset(0, 385) = Right(explain11, Len(explain11) - Len("每次改造/预算："))
        '-------------------------------------------------------------
        
        '-------------------------------------------------------------
        'Part9:未来预测
        '-------------------------------------------------------------
        benchCell.Offset(0, 386) = Workbooks(surveybookName).Worksheets("Survey").optionbutton54.Value
        benchCell.Offset(0, 387) = Workbooks(surveybookName).Worksheets("Survey").optionbutton76.Value
        benchCell.Offset(0, 388) = Workbooks(surveybookName).Worksheets("Survey").optionbutton220.Value
        benchCell.Offset(0, 389) = Workbooks(surveybookName).Worksheets("Survey").optionbutton221.Value
        benchCell.Offset(0, 390) = Workbooks(surveybookName).Worksheets("Survey").optionbutton222.Value
        benchCell.Offset(0, 391) = Workbooks(surveybookName).Worksheets("Survey").optionbutton223.Value
        benchCell.Offset(0, 392) = Workbooks(surveybookName).Worksheets("Survey").optionbutton21.Value
        
        Dim explain12 As String
        explain12 = surveySheet.Range("e144").Value
        benchCell.Offset(0, 393) = Right(explain12, Len(explain12) - Len("请说明:"))
        
        benchCell.Offset(0, 394) = Workbooks(surveybookName).Worksheets("Survey").optionbutton48.Value
        benchCell.Offset(0, 395) = Workbooks(surveybookName).Worksheets("Survey").optionbutton49.Value
        benchCell.Offset(0, 396) = Workbooks(surveybookName).Worksheets("Survey").optionbutton50.Value

        benchCell.Offset(0, 397) = Workbooks(surveybookName).Worksheets("Survey").optionbutton51.Value
        benchCell.Offset(0, 398) = Workbooks(surveybookName).Worksheets("Survey").optionbutton77.Value
        benchCell.Offset(0, 399) = Workbooks(surveybookName).Worksheets("Survey").optionbutton78.Value
        benchCell.Offset(0, 400) = Workbooks(surveybookName).Worksheets("Survey").optionbutton52.Value
        benchCell.Offset(0, 401) = Workbooks(surveybookName).Worksheets("Survey").optionbutton72.Value
        benchCell.Offset(0, 402) = Workbooks(surveybookName).Worksheets("Survey").optionbutton73.Value
        benchCell.Offset(0, 403) = Workbooks(surveybookName).Worksheets("Survey").optionbutton74.Value
        benchCell.Offset(0, 404) = Workbooks(surveybookName).Worksheets("Survey").optionbutton53.Value
        benchCell.Offset(0, 405) = Workbooks(surveybookName).Worksheets("Survey").optionbutton75.Value
        
        Dim explain13 As String
        explain13 = surveySheet.Range("e147").Value
        benchCell.Offset(0, 406) = Right(explain13, Len(explain13) - Len("请说明:"))
        
        benchCell.Offset(0, 407) = Workbooks(surveybookName).Worksheets("Survey").optionbutton55.Value
        benchCell.Offset(0, 408) = Workbooks(surveybookName).Worksheets("Survey").optionbutton56.Value
        benchCell.Offset(0, 409) = Workbooks(surveybookName).Worksheets("Survey").optionbutton57.Value
        benchCell.Offset(0, 410) = Workbooks(surveybookName).Worksheets("Survey").optionbutton58.Value
        benchCell.Offset(0, 411) = Workbooks(surveybookName).Worksheets("Survey").optionbutton59.Value
        
        benchCell.Offset(0, 412) = Workbooks(surveybookName).Worksheets("Survey").optionbutton79.Value
        benchCell.Offset(0, 413) = Workbooks(surveybookName).Worksheets("Survey").optionbutton80.Value
        benchCell.Offset(0, 414) = Workbooks(surveybookName).Worksheets("Survey").optionbutton81.Value
        benchCell.Offset(0, 415) = Workbooks(surveybookName).Worksheets("Survey").optionbutton82.Value
        benchCell.Offset(0, 416) = Workbooks(surveybookName).Worksheets("Survey").optionbutton83.Value
        benchCell.Offset(0, 417) = Workbooks(surveybookName).Worksheets("Survey").optionbutton84.Value
        benchCell.Offset(0, 418) = Workbooks(surveybookName).Worksheets("Survey").optionbutton85.Value
        benchCell.Offset(0, 419) = Workbooks(surveybookName).Worksheets("Survey").optionbutton86.Value
        benchCell.Offset(0, 420) = Workbooks(surveybookName).Worksheets("Survey").optionbutton224.Value
        
        Dim explain14 As String
        explain14 = surveySheet.Range("f150").Value
        benchCell.Offset(0, 421) = Right(explain14, Len(explain14) - Len("请说明:"))
        
        benchCell.Offset(0, 422) = Workbooks(surveybookName).Worksheets("Survey").CheckBox63.Value
        benchCell.Offset(0, 423) = Workbooks(surveybookName).Worksheets("Survey").CheckBox64.Value
        benchCell.Offset(0, 424) = Workbooks(surveybookName).Worksheets("Survey").CheckBox65.Value
        benchCell.Offset(0, 425) = Workbooks(surveybookName).Worksheets("Survey").CheckBox1.Value
        benchCell.Offset(0, 426) = Workbooks(surveybookName).Worksheets("Survey").CheckBox2.Value
        benchCell.Offset(0, 427) = Workbooks(surveybookName).Worksheets("Survey").CheckBox66.Value
        benchCell.Offset(0, 428) = Workbooks(surveybookName).Worksheets("Survey").CheckBox67.Value
        benchCell.Offset(0, 429) = Workbooks(surveybookName).Worksheets("Survey").CheckBox68.Value
        benchCell.Offset(0, 430) = Workbooks(surveybookName).Worksheets("Survey").CheckBox4.Value
        benchCell.Offset(0, 431) = Workbooks(surveybookName).Worksheets("Survey").CheckBox5.Value
        benchCell.Offset(0, 432) = Workbooks(surveybookName).Worksheets("Survey").CheckBox3.Value
        
        Dim explain15 As String
        explain15 = surveySheet.Range("f153").Value
        benchCell.Offset(0, 433) = Right(explain15, Len(explain15) - Len("请说明:"))

        benchCell.Offset(0, 434) = Workbooks(surveybookName).Worksheets("Survey").CheckBox6.Value
        benchCell.Offset(0, 435) = Workbooks(surveybookName).Worksheets("Survey").CheckBox69.Value
        benchCell.Offset(0, 436) = Workbooks(surveybookName).Worksheets("Survey").CheckBox70.Value
        benchCell.Offset(0, 437) = Workbooks(surveybookName).Worksheets("Survey").CheckBox71.Value
        benchCell.Offset(0, 438) = Workbooks(surveybookName).Worksheets("Survey").CheckBox72.Value
        benchCell.Offset(0, 439) = Workbooks(surveybookName).Worksheets("Survey").CheckBox73.Value
        benchCell.Offset(0, 440) = Workbooks(surveybookName).Worksheets("Survey").CheckBox74.Value
        benchCell.Offset(0, 441) = Workbooks(surveybookName).Worksheets("Survey").CheckBox75.Value
        benchCell.Offset(0, 442) = Workbooks(surveybookName).Worksheets("Survey").CheckBox76.Value
        benchCell.Offset(0, 443) = Workbooks(surveybookName).Worksheets("Survey").CheckBox77.Value
        benchCell.Offset(0, 444) = Workbooks(surveybookName).Worksheets("Survey").CheckBox78.Value
        benchCell.Offset(0, 445) = Workbooks(surveybookName).Worksheets("Survey").CheckBox79.Value
        benchCell.Offset(0, 446) = Workbooks(surveybookName).Worksheets("Survey").CheckBox80.Value
        benchCell.Offset(0, 447) = Workbooks(surveybookName).Worksheets("Survey").CheckBox81.Value
        benchCell.Offset(0, 448) = Workbooks(surveybookName).Worksheets("Survey").CheckBox82.Value
        benchCell.Offset(0, 449) = Workbooks(surveybookName).Worksheets("Survey").CheckBox83.Value
        benchCell.Offset(0, 450) = Workbooks(surveybookName).Worksheets("Survey").CheckBox84.Value
        benchCell.Offset(0, 451) = Workbooks(surveybookName).Worksheets("Survey").CheckBox85.Value
        benchCell.Offset(0, 452) = Workbooks(surveybookName).Worksheets("Survey").CheckBox86.Value
        benchCell.Offset(0, 453) = Workbooks(surveybookName).Worksheets("Survey").CheckBox87.Value
        benchCell.Offset(0, 454) = Workbooks(surveybookName).Worksheets("Survey").CheckBox88.Value
        benchCell.Offset(0, 455) = Workbooks(surveybookName).Worksheets("Survey").CheckBox89.Value
        benchCell.Offset(0, 456) = Workbooks(surveybookName).Worksheets("Survey").CheckBox90.Value
        benchCell.Offset(0, 457) = Workbooks(surveybookName).Worksheets("Survey").CheckBox91.Value
        benchCell.Offset(0, 458) = Workbooks(surveybookName).Worksheets("Survey").CheckBox92.Value
        benchCell.Offset(0, 459) = Workbooks(surveybookName).Worksheets("Survey").CheckBox93.Value
        benchCell.Offset(0, 460) = Workbooks(surveybookName).Worksheets("Survey").CheckBox94.Value
        benchCell.Offset(0, 461) = Workbooks(surveybookName).Worksheets("Survey").CheckBox95.Value

        Dim explain16 As String
        explain16 = surveySheet.Range("f156").Value
        benchCell.Offset(0, 462) = Right(explain16, Len(explain16) - Len("请说明:"))
        
        benchCell.Offset(0, 463) = Workbooks(surveybookName).Worksheets("Survey").CheckBox96.Value
        benchCell.Offset(0, 464) = Workbooks(surveybookName).Worksheets("Survey").CheckBox97.Value
        benchCell.Offset(0, 465) = Workbooks(surveybookName).Worksheets("Survey").CheckBox98.Value
        benchCell.Offset(0, 466) = Workbooks(surveybookName).Worksheets("Survey").CheckBox99.Value
        benchCell.Offset(0, 467) = Workbooks(surveybookName).Worksheets("Survey").CheckBox100.Value
        benchCell.Offset(0, 468) = Workbooks(surveybookName).Worksheets("Survey").CheckBox101.Value
        benchCell.Offset(0, 469) = Workbooks(surveybookName).Worksheets("Survey").CheckBox102.Value
        benchCell.Offset(0, 470) = Workbooks(surveybookName).Worksheets("Survey").CheckBox103.Value
        benchCell.Offset(0, 471) = Workbooks(surveybookName).Worksheets("Survey").CheckBox104.Value
        benchCell.Offset(0, 472) = Workbooks(surveybookName).Worksheets("Survey").CheckBox105.Value
        benchCell.Offset(0, 473) = Workbooks(surveybookName).Worksheets("Survey").CheckBox106.Value
        benchCell.Offset(0, 474) = Workbooks(surveybookName).Worksheets("Survey").CheckBox107.Value
        benchCell.Offset(0, 475) = Workbooks(surveybookName).Worksheets("Survey").CheckBox108.Value
        benchCell.Offset(0, 476) = Workbooks(surveybookName).Worksheets("Survey").CheckBox109.Value
        benchCell.Offset(0, 477) = Workbooks(surveybookName).Worksheets("Survey").CheckBox110.Value
        benchCell.Offset(0, 478) = Workbooks(surveybookName).Worksheets("Survey").CheckBox111.Value
        benchCell.Offset(0, 479) = Workbooks(surveybookName).Worksheets("Survey").CheckBox112.Value
        benchCell.Offset(0, 480) = Workbooks(surveybookName).Worksheets("Survey").CheckBox113.Value
        benchCell.Offset(0, 481) = Workbooks(surveybookName).Worksheets("Survey").CheckBox114.Value
        benchCell.Offset(0, 482) = Workbooks(surveybookName).Worksheets("Survey").CheckBox115.Value
        benchCell.Offset(0, 483) = Workbooks(surveybookName).Worksheets("Survey").CheckBox116.Value
        benchCell.Offset(0, 484) = Workbooks(surveybookName).Worksheets("Survey").CheckBox117.Value
        benchCell.Offset(0, 485) = Workbooks(surveybookName).Worksheets("Survey").CheckBox118.Value
        benchCell.Offset(0, 486) = Workbooks(surveybookName).Worksheets("Survey").CheckBox119.Value
        benchCell.Offset(0, 487) = Workbooks(surveybookName).Worksheets("Survey").CheckBox120.Value
        benchCell.Offset(0, 488) = Workbooks(surveybookName).Worksheets("Survey").CheckBox121.Value
        benchCell.Offset(0, 489) = Workbooks(surveybookName).Worksheets("Survey").CheckBox122.Value
        benchCell.Offset(0, 490) = Workbooks(surveybookName).Worksheets("Survey").CheckBox123.Value

        Dim explain17 As String
        explain17 = surveySheet.Range("f158").Value
        benchCell.Offset(0, 491) = Right(explain17, Len(explain17) - Len("请说明:"))
        '-------------------------------------------------------------
        'MsgBox "active workbook is" & surveybookName
        surveyBook.Close
        
        
        
    Next x
    summarySheet.Activate
End Sub
