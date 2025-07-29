VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsZgCgToPz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public ILoginable As New UnkObject


Public Property Get Business() '1102D330
  loc_1102D36F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102D385: Set var_18 =
  loc_1102D390: GoTo loc_1102D39C
  loc_1102D39B: Exit Sub
  loc_1102D39C: ' Referenced from: 1102D390
End Sub

Public Property Set Business(bBus) '1102D3D0
  loc_1102D40E: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102D421: Set bBus =
End Sub

Public Sub ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine) '1102D450
  Dim var_8010 As 
  Dim var_124 As ADODB.Connection
  loc_1102D4E0: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102D4F4: var_20 = cMenuId
  loc_1102D4FC: var_18 = cMenuName
  loc_1102D504: var_1C = cAuthId
  loc_1102D50C: var_24 = cCmdLine
  loc_1102D516: var_8004 = 0()
  loc_1102D547:   Set var_8008 =
  loc_1102D581:   Set var_800C =
  loc_1102D5C6:   var_5C = 00000000h."UfSystemDb".00000000h
  loc_1102D5EE:   var_6C = "cAcc_Id".00000000h
  loc_1102D616:   var_7C = "cIYear".00000000h
  loc_1102D641:   var_8C = "cUserId".00000000h
  loc_1102D6BF:   ebx+110F6020h.InitPubs("CurDate".00000000h(12), "CurDate".00000000h(8), "CurDate".00000000h(4), "SysPassword".00000000h(12), "SysPassword".00000000h(8), "SysPassword".00000000h(4), "SysPassword".00000000h)
  loc_1102D6FD:   var_8010 = CreateObject(global_1100C90C)
  loc_1102D70B:   Set var_8010 = 
  loc_1102D726:   var_B4 = "ufmeta_"
  loc_1102D759:   ecx =  & "cAcc_Id".00000000h
  loc_1102D791:   ADODB.Connection.ConnectionTimeout = 0
  loc_1102D798:   If  < 0 Then
  loc_1102D7C5:   var_134 = CheckObj(var_124, global_1100BD8C, 52)
  loc_1102D7D5:   var_38 = CStr("UfDbName".00000000h)
  loc_1102D82E:   clsZgCgToPz.Business = 0
  loc_1102D862:   var_8018 = clsZgCgToPz.Proc_7_10_1102DF40(global_1100AE28, global_1100AE28, global_1100AE28, -1)
  loc_1102D93C:   Set ebx+110F61BCh = {387D67A4-1DCD-4A2C-8E32AF6DC1C5DBD8}()
  loc_1102D94A:   Set var_34 = ebx+110F6014h.CreateFormEnv("U8V9_CY_FTZJG.frmZGCGToPz", ebx+110F61BCh)
  loc_1102DA01:   ebx+110F6014h.ShowForm(ebx+110F61BCh, "AA", "asdasdasd")
  loc_1102DA1B:   Set ebx+110F61BCh = {387D67A4-1DCD-4A2C-8E32AF6DC1C5DBD8}()
  loc_1102DA2A: End If
  loc_1102DA2F: GoTo loc_1102DA74
  loc_1102DA73: Exit Sub
  loc_1102DA74: ' Referenced from: 1102DA2F
End Sub

Public Sub CloseForm() '1102DAD0

End Sub

Public Function getFormObj(objForm) '1102DB30
  loc_1102DB74: GoTo loc_1102DB80
  loc_1102DB7F: Exit Function
  loc_1102DB80: ' Referenced from: 1102DB74
End Function

Public Function getFormHandle() '1102DBC0

End Function

Public Function getFormHandleEx() '1102DC30

End Function

Public Function CallFunction(cMenuId, cMenuName, cAuthId, cCmdLine) '1102DDD0
  loc_1102DE54: var_8004 = clsZgCgToPz.ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine)
  loc_1102DE6F: GoTo loc_1102DE7B
  loc_1102DE7A: Exit Function
  loc_1102DE7B: ' Referenced from: 1102DE6F
End Function

Private  Proc_7_8_1102DCA0(arg_C, arg_10, arg_14, arg_18) '1102DCA0
  Dim var_28 As var_2C
  loc_1102DD57: Set var_28 = clsZgCgToPz.CallFunction(arg_C, arg_10, arg_14, arg_18)
  loc_1102DD65: GoTo loc_1102DD80
  loc_1102DD6B: If var_4 Then
  loc_1102DD76: End If
  loc_1102DD7F: Exit Sub
  loc_1102DD80: ' Referenced from: 1102DD65
End Sub

Private Sub Proc_7_9_1102DEC0

End Sub

Private Sub Proc_7_10_1102DF40

End Sub
