VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BC5180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器：五分类血液细胞分析仪
'型号：BC5180 CRP
'厂家：迈瑞
'用户：吴川市第四人民医院
'时间：2022.03.12
'

'公共变量
Dim mstrCode As String
'Base64对照表
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="


'以下是必须的接口函数

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, _
                   strResult As String, _
                   strReserved As String, _
                   strCmd As String)
    Dim lngS, lngE As Long
    Dim strCode As String
    Dim segment, element As Variant
    Dim i As Integer
    Dim str组件分隔符, str重复分隔符, str转义分隔符, str子组件分隔符 As String
    Dim strIsQuality, strSampleID, strIterm, strDateTime As String
    Dim strBase64, strWBC, strRBC, strPLT, strDIFF, strImg As String
    
    i = 0
    strResult = ""
    
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lngS = InStr(mstrCode, "")
    
    If lngS > 0 Then
        lngE = InStr(mstrCode, "") '这里面已经有结束字符了
    Else
        lngE = 0
    End If
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid$(mstrCode, lngS + 1, lngE - lngS)
        mstrCode = Mid$(mstrCode, lngE + 1)
    
        segment = Split(strCode, Chr$(&HD))

        For i = LBound(segment) To UBound(segment)

            'If IsEmpty(segment(i)) = False And segment(i) <> vbCrLf And segment(i) <> Chr$(&HD) And segment(i) <> Chr$(&HA) And segment(i) <> "" Then
            If segment(i) Like "*" & "|" & "*" Then
                element = Split(segment(i), "|")

                Select Case element(0)

                    Case "MSH"
                        str组件分隔符 = Mid$(element(1), 1, 1)
                        str重复分隔符 = Mid$(element(1), 2, 1)
                        str转义分隔符 = Mid$(element(1), 3, 1)
                        str子组件分隔符 = Mid$(element(1), 4, 1)
                        strIsQuality = IIf(element(10) = "P", 0, 1)

                    Case "OBR"
                        strSampleID = element(3)
                        strDateTime = Format(Mid$(element(7), 1, 2) & "-" & Mid$(element(7), 3, 2) & "-" & Mid$(element(7), 5, 2) & " " & Mid$(element(7), 7, 2) & ":" & Mid$(element(7), 9, 2) & ":" & Mid$(element(7), 11, 2), "yyyy-mm-dd hh:mm:ss")

                    Case "OBX"

                        If element(2) = "NM" Then
                            strIterm = strIterm & "|" & Split(element(3), "^")(1) & "|" & element(5)
                        End If
                
                        If element(2) = "ED" And Split(element(3), "^")(1) = "WBC Histogram. Binary" Then
                            strBase64 = Split(element(5), "^")(4)
                            strWBC = SaveBase64Bmp("WBC", strSampleID, strBase64)
                        End If
                
                        If element(2) = "ED" And Split(element(3), "^")(1) = "RBC Histogram. Binary" Then
                            strBase64 = Split(element(5), "^")(4)
                            strRBC = SaveBase64Bmp("RBC", strSampleID, strBase64)
                        End If
                
                        If element(2) = "ED" And Split(element(3), "^")(1) = "PLT Histogram. Binary" Then
                            strBase64 = Split(element(5), "^")(4)
                            strPLT = SaveBase64Bmp("PLT", strSampleID, strBase64)
                        End If
                
                        If element(2) = "ED" And Split(element(3), "^")(1) = "WBC DIFF Scattergram. BIN" Then
                            strBase64 = Split(element(5), "^")(4)
                            strDIFF = SaveBase64Bmp("DIFF", strSampleID, strBase64)
                        End If

                End Select

            End If

        Next

        If Val(strSampleID) > 0 And strIterm <> "" Then
            strResult = strResult & "||" & strDateTime & "|" & strSampleID & "||血液|" & strIsQuality & strIterm
            strImg = ""

            If strWBC <> "" Then
                strImg = strImg & "^WBC;101;1;" & strWBC
            End If

            If strRBC <> "" Then
                strImg = strImg & "^RBC;101;1;" & strRBC
            End If

            If strPLT <> "" Then
                strImg = strImg & "^PLT;101;1;" & strPLT
            End If

            If strDIFF <> "" Then
                strImg = strImg & "^DIFF;101;1;" & strDIFF
            End If

            If strImg <> "" Then
                strResult = strResult & vbCrLf & Mid$(strImg, 2)
            End If
            
        End If

        lngS = InStr(mstrCode, "") '开始字符VT
        lngE = InStr(mstrCode, "") '里面已经有了结束字符FS
    Loop
    
    If strResult <> "" Then
        strResult = Mid$(strResult, 3)
    End If

End Sub

'---------------------------------------------------------------------------------------------------------------
'以下是内部函数
Private Function SaveBase64Bmp(ByVal strType As String, ByVal str标本号 As String, ByVal str2Decode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
    
    Dim ByteData() As Byte, lngCount As Long, strFilename As String, lngFileNum
    
    lngCount = Len(str2Decode)
    ReDim ByteData(lngCount / 4 * 3)
    lngCount = 0
    ' for each 4 character group....
    For lPtr = 1 To Len(str2Decode) Step 4
        iLen = 4
        For iCtr = 0 To 3
            ' retrive the base 64 value, 4 at a time
            iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                ' not found
                Case 0: Exit Function
            End Select
        Next

        ' convert the 4, 6 bit values into 3, 8 bit values
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)

        ' add the three new characters to the output string
        For iCtr = 1 To iLen - 1
            ByteData(lngCount) = bits(iCtr)
            lngCount = lngCount + 1
        Next
    Next
    
    strFilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFilename For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBase64Bmp = strFilename

End Function









