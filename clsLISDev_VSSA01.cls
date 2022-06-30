VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_VSSA01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪 器：维生素分析仪
'型 号：VSS-A-01
'厂 商：中芯维康
'作 者：吴品江
'用 户：吴川市第四人民医院
'运行库：mdPublicByWuPinjiang.bas
'ASCII <STX> = <0x02> 正文开始
'ASCII <ETX> = <0x03> 正文结束
'ASCII <VT> = <0x0B> 垂直制表符
'ASCII <FS> = <0x1C> 文件分割符
'ASCII <CR> = <0x0D> 回车
'ASCII <LF> = <0x0A> 换行
'ASCII <SPACE> = <0x20> 空格


'待解析数据
Private strAllDataFromDevice As String
Private mstrRack As String
Private mstrWaitSend As String


'获取设备开始发送命令
Public Function GetStartCmd() As String
    GetStartCmd = ""
End Function

'获取设备停止发送命令
Public Function GetEndCmd() As String
    GetEndCmd = ""
End Function

'自动应答指令（定时应答）
Public Function GetAnswerCmd() As String
    GetAnswerCmd = Chr(&H6)
End Function

'解析LIS查询到的医嘱项目，重组为仪器可识别的格式
Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    Dim varSample As Variant, iCount As Integer, iLoop As Integer
    Dim strRecord As String, strSampleNO As String
    Dim iEmergency As Integer, strSampleBarCode As String, strItem As String
    Dim strSampleType As String
    Dim strSendItem As String, strLIne As String
    
    '2022-06-28 16:41:44|1|测试病人|血清|routine|220628000173| , |514^男性^1999-05-07^22岁^ ^CeShiBingRen^0^ ^内二科^202100001^202100001| | |VE|VB1|VB9
    
    varSample = Split(strSampleInfo, "||")
    
    SendSample = ""
    
    blnSuccess = True
    
    If iSendStep = 0 Then
        For iCount = LBound(varSample) To UBound(varSample)
            strRecord = varSample(iCount)
            If Val(Split(strRecord, "|")(1)) = 0 Then
                strSampleNO = Trim(Split(strRecord, "|")(5))
            Else
                strSampleNO = Val(Split(strRecord, "|")(1))
            End If
            
            strSampleType = Split(strRecord, "|")(3)
            iEmergency = Val(Split(strRecord, "|")(4))
            strSampleBarCode = Trim(Split(strRecord, "|")(5))
            strItem = ""
            
            For iLoop = 10 To UBound(Split(strRecord, "|"))
                    strItem = strItem & Split(strRecord, "|")(iLoop) & "^^^\"
            Next
            
            If strSampleBarCode <> "" And strItem <> "" Then
                strItem = Mid(strItem, 1, Len(strItem) - 1)
                
                'First phase
                strLIne = Chr(&H2) & "1H|\^&|||VSS-A-01^20200401^V1.1.3|||||||QA|1394-97|" & Replace(Split(Now, " ")(0), "/", "")
                strSendItem = strSendItem & strLIne & Chr(&HD)
                
                'Second phase
                strLIne = "P|1||||^^||^^|||||||||||||||||||||||||||"
                strSendItem = strSendItem & strLIne & Chr(&HD)
                
                'Third phase
                strLIne = "O|1|^^|" & strSampleNO & "^^" & strSampleBarCode & "|" & strItem & "|" & IIf(iEmergency = 1, "S", "R")
                strLIne = strLIne & "||||||||||||||||||||O|||||"
                strSendItem = strSendItem & strLIne & Chr(&HD)
                
                'Forth phase
                strLIne = "L|1|N"
                strSendItem = strSendItem & strLIne & Chr(&HD) & Chr(&H3) & GetCheckCode(strLIne) & vbLf
                
                If SendSample = "" Then
                    SendSample = SendSample & Chr(&H5)
                    iSendStep = 1
                    mstrWaitSend = mstrWaitSend & "," & strSendItem & Chr(&H4)
                Else
                    mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & strSendItem & Chr(&H4)
                End If

            End If
        Next
    Else
        If InStr(strResponse, Chr(&H6)) > 0 Then
            If mstrWaitSend <> "" Then
                If Left(mstrWaitSend, 1) = "," Then
                    mstrWaitSend = Mid(mstrWaitSend, 2)
                End If
                
                If InStr(mstrWaitSend, ",") > 0 Then
                    SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
                    mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                Else
                     SendSample = mstrWaitSend
                     mstrWaitSend = ""
                End If
            End If
        End If
        If mstrWaitSend = "" Then iSendStep = 0
        Exit Function
    End If
    
End Function

'双向函数
Private Function Duplexing(ByRef varLine As Variant, _
                           ByRef strResult As String, _
                           ByRef strCmd As String) As String
    On Error GoTo ErrorHandle
    
    Dim strSampleBarCode As String
    strSampleBarCode = ""
    
    strSampleBarCode = Val(Replace(Trim(Split(varLine(1), "|")(2)), "^", ""))
    'mstrRack = Split(strSampleBarCode, "^")(2) & "^" & Split(strSampleBarCode, "^")(3) & "^" & Split(strSampleBarCode, "^")(4)
    
    strResult = strResult & "||" & Format(Now, "YYYY-MM-DD") & "|^^" & strSampleBarCode
    
    '通知程序，开始双向
    strCmd = "1|" & strCmd
    
    
    Duplexing = "1"
    
ErrorHandle:

    Exit Function

End Function

'单向函数
Private Function Singlex(ByRef varLine As Variant, ByRef strResult As String) As String

    Dim strDate, strSampleNO, strQuality, strEmergency, strItem As String

    Dim varElement As Variant

    Dim n          As Integer
    
    On Error GoTo ErrorHandle
    
    strDate = "": strSampleNO = "": strQuality = "P": strEmergency = "": strItem = ""
    '这里是原本的单向
    '右边界-1是因为按照<CR>拆分，会有一个""作为最后一项，这个空字符串会导致溢出
    For n = LBound(varLine) To UBound(varLine) - 1
       varElement = Split(varLine(n), "|")
        
        If varElement(0) = "1H" Then
            strDate = Convert2Date(varElement(13))
        ElseIf varElement(0) = "O" Then
            strSampleNO = varElement(3)
        ElseIf varElement(0) = "R" Then
            strItem = strItem & "|" & Split(varElement(2), "^")(1) & "|" & Split(varElement(3), "^")(0)
        End If
    Next
    
    If strDate <> "" And strSampleNO <> "" And strItem <> "" Then
        strResult = strResult & "||" & strDate & "|" & strSampleNO & "| |血清|" & IIf(strQuality = "Q", 1, 0) & strItem
    End If
    
    Singlex = "1"
ErrorHandle:

    Exit Function

End Function

'解析仪器结果
Public Sub Analyse(ByVal strSource As String, _
                   strResult As String, _
                   strReserved As String, _
                   strCmd As String)

    On Error GoTo ErrorHandle

    Dim strNextSpecimen As String

    Dim varLine As Variant

    Dim lngS, lngE As Long

    Dim DuplexingResult, SinglexResult As String

    '一般应答06
    If InStr(strSource, Chr(&H5)) > 0 Or InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)

    If strSource <> "" Then
        strAllDataFromDevice = strAllDataFromDevice & strSource
    Else

        Exit Sub

    End If

    lngS = InStr(strAllDataFromDevice, Chr(&H2))

    If lngS > 0 Then
        lngE = InStr(lngS, strAllDataFromDevice, Chr(&H3))
    Else
        lngE = InStr(strAllDataFromDevice, Chr(&H3))
    End If
    
    Do While lngS > 0 And lngE > 0
    
        strNextSpecimen = Mid$(strAllDataFromDevice, lngS + 1, lngE - lngS - 1)
        strAllDataFromDevice = Mid$(strAllDataFromDevice, lngE + 1)
        
        varLine = Split(strNextSpecimen, Chr(&HD))
        
        '单双向分流
        If varLine(1) Like "Q" & "*" Then
            DuplexingResult = Duplexing(varLine, strResult, strCmd)
        ElseIf varLine(1) Like "O" & "*" Then
            SinglexResult = Singlex(varLine, strResult)
        End If
        
        lngS = InStr(strAllDataFromDevice, Chr(&H2))

        If lngS > 0 Then
            lngE = InStr(lngS, strAllDataFromDevice, Chr(&H3))
        Else
            lngE = InStr(strAllDataFromDevice, Chr(&H3))
        End If
        
    Loop
    
    If strResult <> "" Then
        strResult = Mid$(strResult, 3)
    End If
    
ErrorHandle:

    Exit Sub

End Sub



