VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BEION_S3_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BEION S3-3 精子质量分析仪
'厂商:北昂医疗
'作者:hj
'用户:威远县人民医院

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'标准函数：从文件获取检验结果
'   strFile：包含路径的文件名
'   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
'   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
'   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
'   返回的检验结果数组(各解析程序必须按以下标准组织结果)
'   元素之间以|分隔
'   第0个元素：检验时间
'   第1个元素：样本序号
'   第2个元素：检验人
'   第3个元素：标本
'   第4个元素：是否质控品
'   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
'       如：第5i个元素为检验项目，第5i+1个元素为检验结果

    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str操作员 As String, str日期 As String
    Dim strFilePath As String, strPath As String
    Dim strFileName As String, n As Integer
    Dim objStream As TextStream
    Dim strLIne As String, varItem As Variant, varTmp As Variant
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim varStr As Variant, strImg As String, str_Line As String
    Dim lngLoop As Long, i As Integer, j As Integer
    Dim strTmp As String
    Dim strData As String           '解析数据
    Dim intCount As Integer         '计数
    Dim blnBegin As Boolean         '是否开始数据
    Dim blnItem As Boolean          '是否开始一个项目
    Dim intDay As Integer
    Dim aItem() As String           '项目
    Dim strDate As String           '时间
    Dim fs, f, f1, s, sf
    Dim strImgFolder As String
    
    On Error GoTo errHandle

    If strFile = "" Then MsgBox "未设置读取文件路径，请检查！", vbInformation, "中联接口": Exit Function
    strFilePath = objFile.GetParentFolderName(strFile)
    
    intDay = Abs(DateDiff("d", dtEnd, dtStart))
    ReDim aResults(0)
    aResults(0) = ""
    For i = 0 To intDay
        strFileName = Dir(strFilePath & "\" & Format(dtStart + i, "yyyymmdd") & "*.TXT")
        strImgFolder = strFilePath & "\" & Replace(strFileName, ".txt", "\")
        Do While strFileName <> ""
            Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFileName)
            strImgFolder = strFilePath & "\" & Replace(strFileName, ".txt", "\")
            n = 0
            Do Until objStream.AtEndOfStream
                strLIne = objStream.ReadLine
                If InStr(strLIne, "标本号") Then
                    str标本号 = Replace(Mid(strLIne, InStr(strLIne, "value=") + 6), ">", "")
                ElseIf InStr(strLIne, "报告日期") Then
                    str日期 = Format(Mid(strLIne, InStr(strLIne, "value=") + 6, 10), "yyyy-mm-dd")
                ElseIf InStr(strLIne, "image file") Then
                    strLIne = Replace(strLIne, ">", "")
                    strTmp = strTmp & "^" & Mid(strLIne, InStr(strLIne, ".") - 2, 1) & ";113;0;" & Mid(strLIne, InStr(strLIne, "=") + 1)
                ElseIf InStr(strLIne, "pcode") Then
                    i = InStr(strLIne, "name=") + 5
                    j = InStr(strLIne, "value=") + 6
                    If InStr(strLIne, "化学染色") Or InStr(strLIne, "细胞外铁") Or InStr(strLIne, "细胞内铁") Then j = InStr(strLIne, "value=") + 6
                    If InStr(strLIne, "粒细胞:红细胞") Or InStr(strLIne, "血片共数有核细胞") Or InStr(strLIne, "骨髓共数有核细胞") Then j = InStr(strLIne, "value=") + 6
                    str检验结果 = str检验结果 & "|" & Mid(strLIne, i, InStr(i, strLIne, ";") - i) & "|" & IIf(Mid(strLIne, j, InStr(j, strLIne, ";") - j) = "", "0", Mid(strLIne, j, InStr(j, strLIne, ";") - j))
                ElseIf InStr(strLIne, "statis file=") Then
                    strLIne = Replace(strLIne, ">", "")
                    strImg = Mid(strLIne, InStr(strLIne, "title=") + 6)
                    str_Line = Split(strLIne, ";")(0)
                    strPath = Mid(str_Line, InStr(str_Line, "file=") + 5)
                    'strChenel = objFile.GetParentFolderName(strPath)
                    
                    strTmp = strTmp & "^" & (n + 1) & ";103;0;" & strImgFolder & strPath
                    n = n + 1
                'ElseIf InStr(strLIne, "<trace file=") And strChenel <> "" Then
                ElseIf InStr(strLIne, "<trace file=") Then
                    strLIne = Replace(strLIne, ">", "")
                    strImg = Mid(strLIne, InStr(strLIne, "file=") + 5)
                    'strTmp = strTmp & "^traces;103;0;" & strChenel & "\" & strImg
                    
                    strTmp = strTmp & "^" & (n + 1) & ";103;0;" & strImgFolder & strImg
                    n = n + 1
                End If
            Loop
            If str标本号 <> "" And str检验结果 <> "" Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |精液|0" & str检验结果
                If strTmp <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strTmp, 2)
                str检验结果 = "": strTmp = ""
            End If
            strTmp = ""
            n = 0
            strFileName = Dir
        Loop
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    'Resume
    WriteLog "BEION", strFile, Err.Number & " " & Err.Description
    Resume
End Function


