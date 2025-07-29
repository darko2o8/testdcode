VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsZgXsToPzTGZS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public ILoginable As New UnkObject


Public Property Get Business() '1102EC50
  loc_1102EC8F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102ECA5: Set var_18 =
  loc_1102ECB0: GoTo loc_1102ECBC
  loc_1102ECBB: Exit Sub
  loc_1102ECBC: ' Referenced from: 1102ECB0
End Sub

Public Property Set Business(bBus) '1102ECF0
  loc_1102ED2E: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102ED41: Set bBus =
End Sub

Public Sub ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine) '1102ED70
  Dim var_8010 As 
  Dim var_124 As ADODB.Connection
  loc_1102EE00: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102EE14: var_20 = cMenuId
  loc_1102EE1C: var_18 = cMenuName
  loc_1102EE24: var_1C = cAuthId
  loc_1102EE2C: var_24 = cCmdLine
  loc_1102EE36: var_8004 = 0()
  loc_1102EE67:   Set var_8008 =
  loc_1102EEA1:   Set var_800C =
  loc_1102EEE6:   var_5C = 00000000h."UfSystemDb".00000000h
  loc_1102EF0E:   var_6C = "cAcc_Id".00000000h
  loc_1102EF36:   var_7C = "cIYear".00000000h
  loc_1102EF61:   var_8C = "cUserId".00000000h
  loc_1102EFDF:   ebx+110F6020h.InitPubs("CurDate".00000000h(12), "CurDate".00000000h(8), "CurDate".00000000h(4), "SysPassword".00000000h(12), "SysPassword".00000000h(8), "SysPassword".00000000h(4), "SysPassword".00000000h)
  loc_1102F01D:   var_8010 = CreateObject(global_1100C90C)
  loc_1102F02B:   Set var_8010 = 
  loc_1102F050:   ADODB.Connection.ConnectionTimeout = 0
  loc_1102F057:   If  < 0 Then
  loc_1102F084:   var_134 = CheckObj(var_124, global_1100BD8C, 52)
  loc_1102F094:   var_38 = CStr("UfDbName".00000000h)
  loc_1102F0ED:   clsZgXsToPzTGZS.Business = 0
  loc_1102F121:   var_8014 = clsZgXsToPzTGZS.Proc_9_10_1102F7D0(global_1100AE28, global_1100AE28, global_1100AE28, -1)
  loc_1102F1E5:   Set ebx+110F620Ch = " œûíº@J?þÆŠµf" & Chr(12)()
  loc_1102F1F3:   Set var_34 = ebx+110F6014h.CreateFormEnv("U8V9_CY_FTZJG.frmZGXSToPzTGZS", ebx+110F620Ch)
  loc_1102F29E:   ebx+110F6014h.ShowForm(ebx+110F620Ch, "AA", "asdasdasd")
  loc_1102F2B8:   Set ebx+110F620Ch = " œûíº@J?þÆŠµf" & Chr(12)()
  loc_1102F2C7: End If
  loc_1102F2CC: GoTo loc_1102F311
  loc_1102F310: Exit Sub
  loc_1102F311: ' Referenced from: 1102F2CC
End Sub

Public Sub CloseForm() '1102F360

End Sub

Public Function getFormObj(objForm) '1102F3C0
  loc_1102F404: GoTo loc_1102F410
  loc_1102F40F: Exit Function
  loc_1102F410: ' Referenced from: 1102F404
End Function

Public Function getFormHandle() '1102F450

End Function

Public Function getFormHandleEx() '1102F4C0

End Function

Public Function CallFunction(cMenuId, cMenuName, cAuthId, cCmdLine) '1102F660
  loc_1102F6E4: var_8004 = clsZgXsToPzTGZS.ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine)
  loc_1102F6FF: GoTo loc_1102F70B
  loc_1102F70A: Exit Function
  loc_1102F70B: ' Referenced from: 1102F6FF
End Function

Private  Proc_9_8_1102F530(arg_C, arg_10, arg_14, arg_18) '1102F530
  Dim var_28 As var_2C
  loc_1102F5E7: Set var_28 = clsZgXsToPzTGZS.CallFunction(arg_C, arg_10, arg_14, arg_18)
  loc_1102F5F5: GoTo loc_1102F610
  loc_1102F5FB: If var_4 Then
  loc_1102F606: End If
  loc_1102F60F: Exit Sub
  loc_1102F610: ' Referenced from: 1102F5F5
End Sub

Private Sub Proc_9_9_1102F750

End Sub

Private Sub Proc_9_10_1102F7D0

End Sub
