VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsXsCbToPz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public ILoginable As New UnkObject


Public Property Get Business() '1102C730
  loc_1102C76F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102C785: Set var_18 =
  loc_1102C790: GoTo loc_1102C79C
  loc_1102C79B: Exit Sub
  loc_1102C79C: ' Referenced from: 1102C790
End Sub

Public Property Set Business(bBus) '1102C7D0
  loc_1102C80E: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102C821: Set bBus =
End Sub

Public Sub ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine) '1102C850
  Dim var_8010 As 
  Dim var_124 As ADODB.Connection
  loc_1102C8E0: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102C8F4: var_20 = cMenuId
  loc_1102C8FC: var_18 = cMenuName
  loc_1102C904: var_1C = cAuthId
  loc_1102C90C: var_24 = cCmdLine
  loc_1102C916: var_8004 = 0()
  loc_1102C947:   Set var_8008 =
  loc_1102C981:   Set var_800C =
  loc_1102C9C6:   var_5C = 00000000h."UfSystemDb".00000000h
  loc_1102C9EE:   var_6C = "cAcc_Id".00000000h
  loc_1102CA16:   var_7C = "cIYear".00000000h
  loc_1102CA41:   var_8C = "cUserId".00000000h
  loc_1102CABF:   ebx+110F6020h.InitPubs("CurDate".00000000h(12), "CurDate".00000000h(8), "CurDate".00000000h(4), "SysPassword".00000000h(12), "SysPassword".00000000h(8), "SysPassword".00000000h(4), "SysPassword".00000000h)
  loc_1102CAFD:   var_8010 = CreateObject(global_1100C90C)
  loc_1102CB0B:   Set var_8010 = 
  loc_1102CB30:   ADODB.Connection.ConnectionTimeout = 0
  loc_1102CB37:   If  < 0 Then
  loc_1102CB64:   var_134 = CheckObj(var_124, global_1100BD8C, 52)
  loc_1102CB74:   var_38 = CStr("UfDbName".00000000h)
  loc_1102CBCD:   clsXsCbToPz.Business = 0
  loc_1102CC01:   var_8014 = clsXsCbToPz.Proc_6_10_1102D2B0(global_1100AE28, global_1100AE28, global_1100AE28, -1)
  loc_1102CCC5:   Set ebx+110F6194h = {B1D986BA-A388-4DE3-94416804DF20E4E4}()
  loc_1102CCD3:   Set var_34 = ebx+110F6014h.CreateFormEnv("U8V9_CY_FTZJG.frmXSCBToPz", ebx+110F6194h)
  loc_1102CD7E:   ebx+110F6014h.ShowForm(ebx+110F6194h, "AA", "asdasdasd")
  loc_1102CD98:   Set ebx+110F6194h = {B1D986BA-A388-4DE3-94416804DF20E4E4}()
  loc_1102CDA7: End If
  loc_1102CDAC: GoTo loc_1102CDF1
  loc_1102CDF0: Exit Sub
  loc_1102CDF1: ' Referenced from: 1102CDAC
End Sub

Public Sub CloseForm() '1102CE40

End Sub

Public Function getFormObj(objForm) '1102CEA0
  loc_1102CEE4: GoTo loc_1102CEF0
  loc_1102CEEF: Exit Function
  loc_1102CEF0: ' Referenced from: 1102CEE4
End Function

Public Function getFormHandle() '1102CF30

End Function

Public Function getFormHandleEx() '1102CFA0

End Function

Public Function CallFunction(cMenuId, cMenuName, cAuthId, cCmdLine) '1102D140
  loc_1102D1C4: var_8004 = clsXsCbToPz.ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine)
  loc_1102D1DF: GoTo loc_1102D1EB
  loc_1102D1EA: Exit Function
  loc_1102D1EB: ' Referenced from: 1102D1DF
End Function

Private  Proc_6_8_1102D010(arg_C, arg_10, arg_14, arg_18) '1102D010
  Dim var_28 As var_2C
  loc_1102D0C7: Set var_28 = clsXsCbToPz.CallFunction(arg_C, arg_10, arg_14, arg_18)
  loc_1102D0D5: GoTo loc_1102D0F0
  loc_1102D0DB: If var_4 Then
  loc_1102D0E6: End If
  loc_1102D0EF: Exit Sub
  loc_1102D0F0: ' Referenced from: 1102D0D5
End Sub

Private Sub Proc_6_9_1102D230

End Sub

Private Sub Proc_6_10_1102D2B0

End Sub
