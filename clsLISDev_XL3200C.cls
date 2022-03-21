VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XL3200C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器：凝血分析仪
'型号：XL3200C
'厂家：众驰伟业
'用户：遂溪中医院
'时间：2022.03.04
'作者：WuPinjiang 13420716981（mobile）
'


'不知为何，在此处声明没用，还是要在Analyse函数中声明
'Private mstrCode As String


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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim lngS, lngE As Long
    
    Dim mstrCode, strCode, strDateTime, strNo, strItem, strQuality As String
    
    Dim intCount As Integer
    
    
    strResult = ""
    strQuality = "P"
    
    'UTF-8转ANSI
    'strSource = UTF8_Decode(strSource)
    
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lngS = InStr(mstrCode, "")
    If lngS > 0 Then
        lngE = InStr(mstrCode, "")
    Else
        lngE = 0
    End If
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid(mstrCode, lngS, lngE - lngS + 1)
        mstrCode = Mid(mstrCode, lngE + 1)
        
        strCode = Split(strCode, Chr$(&HD) & Chr$(&HA))
        
        strDateTime = Split(strCode(4), ",")(1) & " " & Split(strCode(5), ",")(1)
        strNo = Split(strCode(1), ",")(1)
        For intCount = 6 To UBound(strCode)
            strItem = "|" & Split(strCode(intCount), ",")(0) & "|" & Split(strCode(intCount), ",")(1)
        Next
        
        If strNo <> "" And strItem <> "" Then
            strResult = strResult & "||" & strDateTime & "|" & strNo & "||血液|" & IIf(strQuality = "Q", 1, 0) & strItem
        End If
    Loop
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
    End If

End Sub

