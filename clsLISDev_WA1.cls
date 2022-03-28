VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_WA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'仪器：电解质分析仪
'型号：WA-1
'用户：遂溪县中医院
'厂家：科方
'
'


Private mstrCode As String

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
    Dim lngS, lngE As Long
    Dim strCode, strDateTime, strSampleID, strIterm, strQuality As String
    Dim varArray As Variant
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lngS = InStr(mstrCode, "@@||")
    lngE = InStr(mstrCode, "||##")
    
    strCode = Mid$(mstrCode, lngS + 4, lngE - (lngS + 4))
    mstrCode = Mid$(mstrCode, lngE + 4)
    
    While lngS > 0 And lngE > 0
        varArray = Split(strCode, "||")
        strDateTime = varArray(3)
        strSampleID = varArray(0)
        strIterm = strIterm & "|" & varArray(1) & varArray(2)
        strResult = strResult & "||" & strDateTime & "|" & strSampleID & "| |血液|" & strQuality & strIterm
        
        lngS = InStr(mstrCode, "@@||")
        lngE = InStr(mstrCode, "||##")
        strCode = Mid$(mstrCode, lngS + 4, lngE - (lngS + 4))
        mstrCode = Mid$(mstrCode, lngE + 4)
    Loop
    
    If strResult <> "" Then
        strResult = Mid$(strResult, 3)
    End If
End Sub
