VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsGzToPzTGZS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public ILoginable As New UnkObject


Public Property Get Business() '1102AE10
  loc_1102AE4F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102AE65: Set var_18 =
  loc_1102AE70: GoTo loc_1102AE7C
  loc_1102AE7B: Exit Sub
  loc_1102AE7C: ' Referenced from: 1102AE70
End Sub

Public Property Set Business(bBus) '1102AEB0
  loc_1102AEEE: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102AF01: Set bBus =
End Sub

Public Sub ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine) '1102AF30
  Dim var_8010 As 
  Dim var_124 As ADODB.Connection
  loc_1102AFC0: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102AFD4: var_20 = cMenuId
  loc_1102AFDC: var_18 = cMenuName
  loc_1102AFE4: var_1C = cAuthId
  loc_1102AFEC: var_24 = cCmdLine
  loc_1102AFF6: var_8004 = 0()
  loc_1102B027:   Set var_8008 =
  loc_1102B061:   Set var_800C =
  loc_1102B0A6:   var_5C = 00000000h."UfSystemDb".00000000h
  loc_1102B0CE:   var_6C = "cAcc_Id".00000000h
  loc_1102B0F6:   var_7C = "cIYear".00000000h
  loc_1102B121:   var_8C = "cUserId".00000000h
  loc_1102B19F:   ebx+110F6020h.InitPubs("CurDate".00000000h(12), "CurDate".00000000h(8), "CurDate".00000000h(4), "SysPassword".00000000h(12), "SysPassword".00000000h(8), "SysPassword".00000000h(4), "SysPassword".00000000h)
  loc_1102B1DD:   var_8010 = CreateObject(global_1100C90C)
  loc_1102B1EB:   Set var_8010 = 
  loc_1102B206:   var_B4 = "ufmeta_"
  loc_1102B239:   ecx =  & "cAcc_Id".00000000h
  loc_1102B271:   ADODB.Connection.ConnectionTimeout = 0
  loc_1102B278:   If  < 0 Then
  loc_1102B2A5:   var_134 = CheckObj(var_124, global_1100BD8C, 52)
  loc_1102B2B5:   var_38 = CStr("UfDbName".00000000h)
  loc_1102B30E:   clsGzToPzTGZS.Business = 0
  loc_1102B342:   var_8018 = clsGzToPzTGZS.Proc_4_10_1102BA20(global_1100AE28, global_1100AE28, global_1100AE28, -1)
  loc_1102B41C:   Set ebx+110F6144h = {BFB65897-A5B9-4E6D-990636F07A7B082D}()
  loc_1102B42A:   Set var_34 = ebx+110F6014h.CreateFormEnv("U8V9_CY_FTZJG.frmGzToPzTGZS", ebx+110F6144h)
  loc_1102B4E1:   ebx+110F6014h.ShowForm(ebx+110F6144h, "AA", "frmGzToPzTGZS")
  loc_1102B4FB:   Set ebx+110F6144h = {BFB65897-A5B9-4E6D-990636F07A7B082D}()
  loc_1102B50A: End If
  loc_1102B50F: GoTo loc_1102B554
  loc_1102B553: Exit Sub
  loc_1102B554: ' Referenced from: 1102B50F
End Sub

Public Sub CloseForm() '1102B5B0

End Sub

Public Function getFormObj(objForm) '1102B610
  loc_1102B654: GoTo loc_1102B660
  loc_1102B65F: Exit Function
  loc_1102B660: ' Referenced from: 1102B654
End Function

Public Function getFormHandle() '1102B6A0

End Function

Public Function getFormHandleEx() '1102B710

End Function

Public Function CallFunction(cMenuId, cMenuName, cAuthId, cCmdLine) '1102B8B0
  loc_1102B934: var_8004 = clsGzToPzTGZS.ShowModeless(cMenuId, cMenuName, cAuthId, cCmdLine)
  loc_1102B94F: GoTo loc_1102B95B
  loc_1102B95A: Exit Function
  loc_1102B95B: ' Referenced from: 1102B94F
End Function

Private  Proc_4_8_1102B780(arg_C, arg_10, arg_14, arg_18) '1102B780
  Dim var_28 As var_2C
  loc_1102B837: Set var_28 = clsGzToPzTGZS.CallFunction(arg_C, arg_10, arg_14, arg_18)
  loc_1102B845: GoTo loc_1102B860
  loc_1102B84B: If var_4 Then
  loc_1102B856: End If
  loc_1102B85F: Exit Sub
  loc_1102B860: ' Referenced from: 1102B845
End Sub

Private Sub Proc_4_9_1102B9A0

End Sub

Private Sub Proc_4_10_1102BA20

End Sub
