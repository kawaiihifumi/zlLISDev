VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU680_V2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:AU680
'厂商:
'作者:吴品江
'时间：2022/4/1
'用户:遂溪县中医院
'难点1：AU680传来的结果中，空格的数量不是固定的。需要将连续的多个空格替换为单个空格。
'难点2：同样是结果中，可能会在数值后面带一个“r”，需要去掉。
'难点3：每一次传输的结果都会被“DB”“DE”包裹，去除之后才可正常解析。

Private mstrCode As String, mstrWaitSend As String, strRackNO As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    
    Dim lngWasteS, lngWasteE, lngS, lngE As Long
    Dim strCode, strSample, strDateTime, strQuality, strIterm, strTemp As String
    Dim varArray As Variant
    Dim i, flag As Integer
    
    strResult = "": strReserved = "": flag = 0
    
    lngWasteS = InStr(strSource, "DB")
    lngWasteE = InStr(strSource, "DE")
    
'    '每次接收到数据，首先把最外层的壳去掉
'    If lngWasteS > 0 And lngWasteE > 0 Then
'        strSource = Mid$(strSource, lngWasteS + 4, lngWasteE - (lngWasteS + 4))
'    Else
'        strReserved = strSource
'        '如果走了这条路，希望有提示......
'        MsgBox "退出Analyse！", vbInformation, "中联接口"
'        Exit Sub
'    End If

    'If lngWasteS > 0 And lngWasteE > 0 Then strSource = Mid$(strSource, lngWasteS + 4, lngWasteE - (lngWasteS + 4))

    
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    '初始化局部变量
    strQuality = "0"
    
    '错误处理
    On Error GoTo errHandle
    
    Call WriteLog("AU680.Analyse 开始解析数据：", strSource, strResult)
    
    
    
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    Else
        Exit Sub
    End If
    
    lngS = InStr(mstrCode, "D")
    If lngS > 0 Then
        lngE = InStr(lngS, mstrCode, "")
    Else
        lngE = InStr(mstrCode, "")
    End If
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid$(mstrCode, lngS + 2, lngE - (lngS + 2))
        mstrCode = Mid$(mstrCode, lngE + 1)
        
        Do While InStr(strCode, "  ") > 0
            strCode = Replace$(strCode, "  ", " ")
        Loop
        
        varArray = Split(strCode, " ")
        
        strDateTime = Format(Now, "YYYY-MM-DD HH:MM:SS")
        strSample = varArray(2)
        
        '循环解析当前标本，把指标名、结果值都看作是相同的项。
        For i = LBound(varArray) + 3 To UBound(varArray) - 1
            If i = 3 Then
                varArray(i) = Mid$(varArray(i), 2)
            End If
            
            '通道码
            '结果的负号会黏在通道码的后面，需要给个flag判断
            If InStr(varArray(i), "-") Then
                strTemp = Mid$(varArray(i), 1, InStr(varArray(i), "-") - 1)
                strIterm = strIterm & "|" & strTemp
                flag = 1
            '指标结果
            ElseIf InStr(varArray(i), "War") > 0 Then
                strTemp = Mid$(varArray(i), 1, InStr(varArray(i), "War") - 1)
                If flag = 1 Then
                    strIterm = strIterm & "|-" & strTemp
                    flag = 0
                ElseIf flag = 0 Then
                    strIterm = strIterm & "|" & strTemp
                End If
            ElseIf InStr(varArray(i), "r") > 0 Then
                strTemp = Mid$(varArray(i), 1, InStr(varArray(i), "r") - 1)
                If flag = 1 Then
                    strIterm = strIterm & "|-" & strTemp
                    flag = 0
                ElseIf flag = 0 Then
                    strIterm = strIterm & "|" & strTemp
                End If
            Else
                strIterm = strIterm & "|" & varArray(i)
            End If
        Next
        
        If strDateTime <> "" And strSample <> "" And strIterm <> "" Then
            strResult = strResult & "||" & strDateTime & "|" & strSample & "| |血液|" & strQuality & strIterm
        End If
        
        lngS = InStr(mstrCode, "D")
        If lngS > 0 Then
            lngE = InStr(lngS, mstrCode, "")
        Else
            lngE = InStr(mstrCode, "")
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid$(strResult, 3)
    
errHandle:
    Call WriteLog("AU680.Analyse 错误：", Err.Number, Err.Description)
End Sub

