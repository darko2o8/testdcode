VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsGzDepMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public ILoginable As New UnkObject


Public Property Get Business() '110294F0
  loc_1102952F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11029545: Set var_18 =
  loc_11029550: GoTo loc_1102955C
  loc_1102955B: Exit Sub
  loc_1102955C: ' Referenced from: 11029550
End Sub

Public Property Set Business(bBus) '11029590
  loc_110295CE: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110295E1: Set bBus =
End Sub

Public Sub ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine) '11029610
  Dim var_8010 As 
  Dim var_124 As ADODB.Connection
  loc_110296A0: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110296B4: var_20 = cMenuId
  loc_110296BC: var_18 = cMenuName
  loc_110296C4: var_1C = cAuthId
  loc_110296CC: var_24 = cCmdLine
  loc_110296D6: var_8004 = 0()
  loc_11029707:   Set var_8008 =
  loc_11029741:   Set var_800C =
  loc_1102975E:   var_4C = "UfSystemDb".00000000h
  loc_11029786:   var_5C = "UfDbName".00000000h
  loc_110297AE:   var_6C = "cAcc_Id".00000000h
  loc_110297D6:   var_7C = "cIYear".00000000h
  loc_11029801:   var_8C = "cUserId".00000000h
  loc_1102982C:   var_9C = "CurDate".00000000h
  loc_11029857:   var_AC = "SysPassword".00000000h
  loc_1102987F:   "InitPubs".00000007h
  loc_110298BD:   var_8010 = CreateObject(global_1100C90C)
  loc_110298CB:   Set var_8010 = 
  loc_110298E6:   var_B4 = "ufmeta_"
  loc_11029919:   ecx =  & "cAcc_Id".00000000h
  loc_11029951:   ADODB.Connection.ConnectionTimeout = 0
  loc_11029958:   If  < 0 Then
  loc_11029985:   var_134 = CheckObj(var_124, global_1100BD8C, 52)
  loc_11029995:   var_38 = CStr("UfDbName".00000000h)
  loc_110299EE:   clsGzDepMap.Business = 0
  loc_11029A22:   var_8018 = clsGzDepMap.Proc_2_10_1102A100(global_1100AE28, global_1100AE28, global_1100AE28, -1)
  loc_11029AFC:   Set ebx+110F60F4h = {DB60EA12-FDE6-4B9F-AFBDBD71C6AE0C0A}()
  loc_11029B0A:   Set var_34 = ebx+110F6014h.CreateFormEnv("U8V9_CY_FTZJG.frmGzDepMap", ebx+110F60F4h)
  loc_11029BC1:   ebx+110F6014h.ShowForm(ebx+110F60F4h, "AA", "frmGzDepMap")
  loc_11029BDB:   Set ebx+110F60F4h = {DB60EA12-FDE6-4B9F-AFBDBD71C6AE0C0A}()
  loc_11029BEA: End If
  loc_11029BEF: GoTo loc_11029C34
  loc_11029C33: Exit Sub
  loc_11029C34: ' Referenced from: 11029BEF
End Sub

Public Sub CloseForm() '11029C90

End Sub

Public Function getFormObj(objForm) '11029CF0
  loc_11029D34: GoTo loc_11029D40
  loc_11029D3F: Exit Function
  loc_11029D40: ' Referenced from: 11029D34
End Function

Public Function getFormHandle() '11029D80

End Function

Public Function getFormHandleEx() '11029DF0

End Function

Public Function CallFunction(cMenuId, cMenuName, cAuthId, cCmdLine) '11029F90
  loc_1102A014: var_8004 = clsGzDepMap.ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine)
  loc_1102A02F: GoTo loc_1102A03B
  loc_1102A03A: Exit Function
  loc_1102A03B: ' Referenced from: 1102A02F
End Function

Private  Proc_2_8_11029E60(arg_C, arg_10, arg_14, arg_18) '11029E60
  Dim var_28 As var_2C
  loc_11029F17: Set var_28 = clsGzDepMap.CallFunction(arg_C, arg_10, arg_14, arg_18)
  loc_11029F25: GoTo loc_11029F40
  loc_11029F2B: If var_4 Then
  loc_11029F36: End If
  loc_11029F3F: Exit Sub
  loc_11029F40: ' Referenced from: 11029F25
End Sub

Private Sub Proc_2_9_1102A080

End Sub

Private Sub Proc_2_10_1102A100

End Sub
