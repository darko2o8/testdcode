VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsGzToPzTGZP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public ILoginable As New UnkObject


Public Property Get Business() '1102A180
  loc_1102A1BF: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102A1D5: Set var_18 =
  loc_1102A1E0: GoTo loc_1102A1EC
  loc_1102A1EB: Exit Sub
  loc_1102A1EC: ' Referenced from: 1102A1E0
End Sub

Public Property Set Business(bBus) '1102A220
  loc_1102A25E: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102A271: Set bBus =
End Sub

Public Sub ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine) '1102A2A0
  Dim var_8010 As 
  Dim var_124 As ADODB.Connection
  loc_1102A330: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102A344: var_20 = cMenuId
  loc_1102A34C: var_18 = cMenuName
  loc_1102A354: var_1C = cAuthId
  loc_1102A35C: var_24 = cCmdLine
  loc_1102A366: var_8004 = 0()
  loc_1102A397:   Set var_8008 =
  loc_1102A3D1:   Set var_800C =
  loc_1102A416:   var_5C = 00000000h."UfSystemDb".00000000h
  loc_1102A43E:   var_6C = "cAcc_Id".00000000h
  loc_1102A466:   var_7C = "cIYear".00000000h
  loc_1102A491:   var_8C = "cUserId".00000000h
  loc_1102A50F:   ebx+110F6020h.InitPubs("CurDate".00000000h(12), "CurDate".00000000h(8), "CurDate".00000000h(4), "SysPassword".00000000h(12), "SysPassword".00000000h(8), "SysPassword".00000000h(4), "SysPassword".00000000h)
  loc_1102A54D:   var_8010 = CreateObject(global_1100C90C)
  loc_1102A55B:   Set var_8010 = 
  loc_1102A576:   var_B4 = "ufmeta_"
  loc_1102A5A9:   ecx =  & "cAcc_Id".00000000h
  loc_1102A5E1:   ADODB.Connection.ConnectionTimeout = 0
  loc_1102A5E8:   If  < 0 Then
  loc_1102A615:   var_134 = CheckObj(var_124, global_1100BD8C, 52)
  loc_1102A625:   var_38 = CStr("UfDbName".00000000h)
  loc_1102A67E:   clsGzToPzTGZP.Business = 0
  loc_1102A6B2:   var_8018 = clsGzToPzTGZP.Proc_3_10_1102AD90(global_1100AE28, global_1100AE28, global_1100AE28, -1)
  loc_1102A78C:   Set ebx+110F611Ch = {C9E70198-A9AF-4EE6-9C0CE544CAB12930}()
  loc_1102A79A:   Set var_34 = ebx+110F6014h.CreateFormEnv("U8V9_CY_FTZJG.frmGzToPzTGZP", ebx+110F611Ch)
  loc_1102A851:   ebx+110F6014h.ShowForm(ebx+110F611Ch, "AA", "frmGzToPzTGZP")
  loc_1102A86B:   Set ebx+110F611Ch = {C9E70198-A9AF-4EE6-9C0CE544CAB12930}()
  loc_1102A87A: End If
  loc_1102A87F: GoTo loc_1102A8C4
  loc_1102A8C3: Exit Sub
  loc_1102A8C4: ' Referenced from: 1102A87F
End Sub

Public Sub CloseForm() '1102A920

End Sub

Public Function getFormObj(objForm) '1102A980
  loc_1102A9C4: GoTo loc_1102A9D0
  loc_1102A9CF: Exit Function
  loc_1102A9D0: ' Referenced from: 1102A9C4
End Function

Public Function getFormHandle() '1102AA10

End Function

Public Function getFormHandleEx() '1102AA80

End Function

Public Function CallFunction(cMenuId, cMenuName, cAuthId, cCmdLine) '1102AC20
  loc_1102ACA4: var_8004 = clsGzToPzTGZP.ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine)
  loc_1102ACBF: GoTo loc_1102ACCB
  loc_1102ACCA: Exit Function
  loc_1102ACCB: ' Referenced from: 1102ACBF
End Function

Private  Proc_3_8_1102AAF0(arg_C, arg_10, arg_14, arg_18) '1102AAF0
  Dim var_28 As var_2C
  loc_1102ABA7: Set var_28 = clsGzToPzTGZP.CallFunction(arg_C, arg_10, arg_14, arg_18)
  loc_1102ABB5: GoTo loc_1102ABD0
  loc_1102ABBB: If var_4 Then
  loc_1102ABC6: End If
  loc_1102ABCF: Exit Sub
  loc_1102ABD0: ' Referenced from: 1102ABB5
End Sub

Private Sub Proc_3_9_1102AD10

End Sub

Private Sub Proc_3_10_1102AD90

End Sub
