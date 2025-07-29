VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsSaleBillDr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public ILoginable As New UnkObject


Public Property Get Business() '1102BAA0
  loc_1102BADF: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102BAF5: Set var_18 =
  loc_1102BB00: GoTo loc_1102BB0C
  loc_1102BB0B: Exit Sub
  loc_1102BB0C: ' Referenced from: 1102BB00
End Sub

Public Property Set Business(bBus) '1102BB40
  loc_1102BB7E: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102BB91: Set bBus =
End Sub

Public Sub ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine) '1102BBC0
  Dim var_8010 As 
  Dim var_124 As ADODB.Connection
  loc_1102BC50: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102BC64: var_20 = cMenuId
  loc_1102BC6C: var_18 = cMenuName
  loc_1102BC74: var_1C = cAuthId
  loc_1102BC7C: var_24 = cCmdLine
  loc_1102BC86: var_8004 = 0()
  loc_1102BCB7:   Set var_8008 =
  loc_1102BCF1:   Set var_800C =
  loc_1102BD36:   var_5C = 00000000h."UfSystemDb".00000000h
  loc_1102BD5E:   var_6C = "cAcc_Id".00000000h
  loc_1102BD86:   var_7C = "cIYear".00000000h
  loc_1102BDB1:   var_8C = "cUserId".00000000h
  loc_1102BE2F:   ebx+110F6020h.InitPubs("CurDate".00000000h(12), "CurDate".00000000h(8), "CurDate".00000000h(4), "SysPassword".00000000h(12), "SysPassword".00000000h(8), "SysPassword".00000000h(4), "SysPassword".00000000h)
  loc_1102BE6D:   var_8010 = CreateObject(global_1100C90C)
  loc_1102BE7B:   Set var_8010 = 
  loc_1102BE96:   var_B4 = "ufmeta_"
  loc_1102BEC9:   ecx =  & "cAcc_Id".00000000h
  loc_1102BF01:   ADODB.Connection.ConnectionTimeout = 0
  loc_1102BF08:   If  < 0 Then
  loc_1102BF35:   var_134 = CheckObj(var_124, global_1100BD8C, 52)
  loc_1102BF45:   var_38 = CStr("UfDbName".00000000h)
  loc_1102BF9E:   clsSaleBillDr.Business = 0
  loc_1102BFD2:   var_8018 = clsSaleBillDr.Proc_5_10_1102C6B0(global_1100AE28, global_1100AE28, global_1100AE28, -1)
  loc_1102C0AC:   Set ebx+110F616Ch = {1AEDA0C9-8F32-4B90-A98D7DE0E72E1BA5}()
  loc_1102C0BA:   Set var_34 = ebx+110F6014h.CreateFormEnv("U8V9_CY_FTZJG.frmSaleBillDr", ebx+110F616Ch)
  loc_1102C171:   ebx+110F6014h.ShowForm(ebx+110F616Ch, "AA", "asdasdasd")
  loc_1102C18B:   Set ebx+110F616Ch = {1AEDA0C9-8F32-4B90-A98D7DE0E72E1BA5}()
  loc_1102C19A: End If
  loc_1102C19F: GoTo loc_1102C1E4
  loc_1102C1E3: Exit Sub
  loc_1102C1E4: ' Referenced from: 1102C19F
End Sub

Public Sub CloseForm() '1102C240

End Sub

Public Function getFormObj(objForm) '1102C2A0
  loc_1102C2E4: GoTo loc_1102C2F0
  loc_1102C2EF: Exit Function
  loc_1102C2F0: ' Referenced from: 1102C2E4
End Function

Public Function getFormHandle() '1102C330

End Function

Public Function getFormHandleEx() '1102C3A0

End Function

Public Function CallFunction(cMenuId, cMenuName, cAuthId, cCmdLine) '1102C540
  loc_1102C5C4: var_8004 = clsSaleBillDr.ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine)
  loc_1102C5DF: GoTo loc_1102C5EB
  loc_1102C5EA: Exit Function
  loc_1102C5EB: ' Referenced from: 1102C5DF
End Function

Private  Proc_5_8_1102C410(arg_C, arg_10, arg_14, arg_18) '1102C410
  Dim var_28 As var_2C
  loc_1102C4C7: Set var_28 = clsSaleBillDr.CallFunction(arg_C, arg_10, arg_14, arg_18)
  loc_1102C4D5: GoTo loc_1102C4F0
  loc_1102C4DB: If var_4 Then
  loc_1102C4E6: End If
  loc_1102C4EF: Exit Sub
  loc_1102C4F0: ' Referenced from: 1102C4D5
End Sub

Private Sub Proc_5_9_1102C630

End Sub

Private Sub Proc_5_10_1102C6B0

End Sub
