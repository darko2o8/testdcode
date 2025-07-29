VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsZgXsToPzTGZP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public ILoginable As New UnkObject


Public Property Get Business() '1102DFC0
  loc_1102DFFF: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102E015: Set var_18 =
  loc_1102E020: GoTo loc_1102E02C
  loc_1102E02B: Exit Sub
  loc_1102E02C: ' Referenced from: 1102E020
End Sub

Public Property Set Business(bBus) '1102E060
  loc_1102E09E: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102E0B1: Set bBus =
End Sub

Public Sub ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine) '1102E0E0
  Dim var_8010 As 
  Dim var_124 As ADODB.Connection
  loc_1102E170: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102E184: var_20 = cMenuId
  loc_1102E18C: var_18 = cMenuName
  loc_1102E194: var_1C = cAuthId
  loc_1102E19C: var_24 = cCmdLine
  loc_1102E1A6: var_8004 = 0()
  loc_1102E1D7:   Set var_8008 =
  loc_1102E211:   Set var_800C =
  loc_1102E256:   var_5C = 00000000h."UfSystemDb".00000000h
  loc_1102E27E:   var_6C = "cAcc_Id".00000000h
  loc_1102E2A6:   var_7C = "cIYear".00000000h
  loc_1102E2D1:   var_8C = "cUserId".00000000h
  loc_1102E34F:   ebx+110F6020h.InitPubs("CurDate".00000000h(12), "CurDate".00000000h(8), "CurDate".00000000h(4), "SysPassword".00000000h(12), "SysPassword".00000000h(8), "SysPassword".00000000h(4), "SysPassword".00000000h)
  loc_1102E38D:   var_8010 = CreateObject(global_1100C90C)
  loc_1102E39B:   Set var_8010 = 
  loc_1102E3B6:   var_B4 = "ufmeta_"
  loc_1102E3E9:   ecx =  & "cAcc_Id".00000000h
  loc_1102E421:   ADODB.Connection.ConnectionTimeout = 0
  loc_1102E428:   If  < 0 Then
  loc_1102E455:   var_134 = CheckObj(var_124, global_1100BD8C, 52)
  loc_1102E465:   var_38 = CStr("UfDbName".00000000h)
  loc_1102E4BE:   clsZgXsToPzTGZP.Business = 0
  loc_1102E4F2:   var_8018 = clsZgXsToPzTGZP.Proc_8_10_1102EBD0(global_1100AE28, global_1100AE28, global_1100AE28, -1)
  loc_1102E5CC:   Set ebx+110F61E4h = {E48D88C5-AD04-4EC5-AF86E066140021E9}()
  loc_1102E5DA:   Set var_34 = ebx+110F6014h.CreateFormEnv("U8V9_CY_FTZJG.frmZGXSToPzTGZP", ebx+110F61E4h)
  loc_1102E691:   ebx+110F6014h.ShowForm(ebx+110F61E4h, "AA", "asdasdasd")
  loc_1102E6AB:   Set ebx+110F61E4h = {E48D88C5-AD04-4EC5-AF86E066140021E9}()
  loc_1102E6BA: End If
  loc_1102E6BF: GoTo loc_1102E704
  loc_1102E703: Exit Sub
  loc_1102E704: ' Referenced from: 1102E6BF
End Sub

Public Sub CloseForm() '1102E760

End Sub

Public Function getFormObj(objForm) '1102E7C0
  loc_1102E804: GoTo loc_1102E810
  loc_1102E80F: Exit Function
  loc_1102E810: ' Referenced from: 1102E804
End Function

Public Function getFormHandle() '1102E850

End Function

Public Function getFormHandleEx() '1102E8C0

End Function

Public Function CallFunction(cMenuId, cMenuName, cAuthId, cCmdLine) '1102EA60
  loc_1102EAE4: var_8004 = clsZgXsToPzTGZP.ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine)
  loc_1102EAFF: GoTo loc_1102EB0B
  loc_1102EB0A: Exit Function
  loc_1102EB0B: ' Referenced from: 1102EAFF
End Function

Private  Proc_8_8_1102E930(arg_C, arg_10, arg_14, arg_18) '1102E930
  Dim var_28 As var_2C
  loc_1102E9E7: Set var_28 = clsZgXsToPzTGZP.CallFunction(arg_C, arg_10, arg_14, arg_18)
  loc_1102E9F5: GoTo loc_1102EA10
  loc_1102E9FB: If var_4 Then
  loc_1102EA06: End If
  loc_1102EA0F: Exit Sub
  loc_1102EA10: ' Referenced from: 1102E9F5
End Sub

Private Sub Proc_8_9_1102EB50

End Sub

Private Sub Proc_8_10_1102EBD0

End Sub
