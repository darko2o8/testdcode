VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C908002B2F49FB}#1.2#0"; "C:\Windows\SysWow64\Comdlg32.ocx"
Begin VB.Form frmZGCGToPz
  Caption = "暂估采购导转凭证"
  BackColor = &H80000005&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmZGCGToPz.frx":0000
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 345
  ClientWidth = 9255
  ClientHeight = 5700
  Appearance = 0 'Flat
  Begin C1SizerLibCtl.C1Elastic Pic1
    Left = 3360
    Top = 3240
    Width = 5025
    Height = 675
    Visible = 0   'False
    TabStop = 0   'False
    TabIndex = 3
    OleObjectBlob = "frmZGCGToPz.frx":014A
    Begin VB.Label Label3
      Caption = "正在分析数据，请稍候。。。"
      Left = 120
      Top = 240
      Width = 4800
      Height = 180
      TabIndex = 4
      AutoSize = -1  'True
      BackStyle = 0 'Transparent
    End
  End
  Begin C1SizerLibCtl.C1Elastic C1Elastic1
    Left = 0
    Top = 0
    Width = 12045
    Height = 5700
    TabStop = 0   'False
    TabIndex = 0
    OleObjectBlob = "frmZGCGToPz.frx":0293
    Begin AIFCmp1.asxStatusBar sBar
      Left = 0
      Top = 5355
      Width = 12045
      Height = 345
      OleObjectBlob = "frmZGCGToPz.frx":04BC
    End
    Begin C1SizerLibCtl.C1Elastic C1Elastic2
      Left = 0
      Top = 0
      Width = 12045
      Height = 435
      TabStop = 0   'False
      TabIndex = 1
      OleObjectBlob = "frmZGCGToPz.frx":05EC
    End
    Begin VSFlex8LCtl.VSFlexGrid VFG
      Left = 0
      Top = 1800
      Width = 12045
      Height = 3540
      TabIndex = 2
      OleObjectBlob = "frmZGCGToPz.frx":0747
      Begin MSComDlg.CommonDialog dlg
        OleObjectBlob = "frmZGCGToPz.frx":0BB0
        Left = 0
        Top = 0
      End
    End
    Begin AIFCmp1.asxPanel asxPanel1
      Left = 0
      Top = 450
      Width = 12045
      Height = 1335
      OleObjectBlob = "frmZGCGToPz.frx":0C14
      Begin VB.CheckBox Chk1
        Caption = "是否使用数据源中的成本中心作为部门核算用"
        BackColor = &HFFCABB&
        Left = 120
        Top = 750
        Width = 3975
        Height = 255
        TabIndex = 16
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 3
        Left = 9045
        Top = 945
        Width = 690
        Height = 330
        TabIndex = 9
        OleObjectBlob = "frmZGCGToPz.frx":0CF4
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 4
        Left = 9825
        Top = 945
        Width = 810
        Height = 330
        TabIndex = 13
        OleObjectBlob = "frmZGCGToPz.frx":0E94
      End
      Begin VB.CheckBox Chk
        Caption = "Check1"
        Index = 0
        Left = 8550
        Top = 480
        Width = 1695
        Height = 300
        Visible = 0   'False
        TabIndex = 6
        Value = 1
      End
      Begin VB.CheckBox Chk
        Caption = "Check1"
        Index = 1
        Left = 8550
        Top = 120
        Width = 1575
        Height = 300
        Visible = 0   'False
        TabIndex = 5
        Value = 1
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 0
        Left = 5925
        Top = 960
        Width = 960
        Height = 330
        TabIndex = 7
        OleObjectBlob = "frmZGCGToPz.frx":1034
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 1
        Left = 6960
        Top = 945
        Width = 960
        Height = 330
        TabIndex = 8
        OleObjectBlob = "frmZGCGToPz.frx":122C
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 0
        Left = 120
        Top = 105
        Width = 6525
        Height = 270
        TabIndex = 10
        OleObjectBlob = "frmZGCGToPz.frx":13FC
        ToolTipText = "项目大类"
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 2
        Left = 7995
        Top = 960
        Width = 960
        Height = 330
        Visible = 0   'False
        TabIndex = 11
        OleObjectBlob = "frmZGCGToPz.frx":1558
      End
      Begin TDBDate6Ctl.TDBDate TDBDate
        Left = 3180
        Top = 1020
        Width = 2505
        Height = 285
        TabIndex = 12
        OleObjectBlob = "frmZGCGToPz.frx":16FC
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 1
        Left = 30
        Top = 420
        Width = 3255
        Height = 270
        TabIndex = 14
        OleObjectBlob = "frmZGCGToPz.frx":19EB
        ToolTipText = "借方科目编码"
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 2
        Left = 3360
        Top = 420
        Width = 3255
        Height = 270
        TabIndex = 15
        OleObjectBlob = "frmZGCGToPz.frx":1B4F
        ToolTipText = "贷方科目编码"
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 3
        Left = 60
        Top = 1020
        Width = 3015
        Height = 270
        TabIndex = 17
        OleObjectBlob = "frmZGCGToPz.frx":1CB3
        ToolTipText = "部门编码"
      End
    End
  End
End

Attribute VB_Name = "frmZGCGToPz"


Private  TDBText_UnknownEvent_B(arg_C) '110BE330
  Dim var_6C As frmZGCGToPz.dlg
  loc_110BE38D: If arg_C = 0 Then
  loc_110BE3A9:   Set var_6C = frmZGCGToPz.dlg
  loc_110BE3DB:   var_6C.FileName = var_4C
  loc_110BE3FD:   var_6C.DialogTitle = var_4C
  loc_110BE41F:   var_6C.Filter = var_4C
  loc_110BE43E:   var_6C.CancelError = var_4C
  loc_110BE448:   var_6C.ShowOpen
  loc_110BE460:   var_6C.FileName = var_6C
  loc_110BE4A2:   If (var_30 = global_1100AE28) Then
  loc_110BE4B4:     var_6C.FileName = Me
  loc_110BE4F1:     arg_C = frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110BE55C:   End If
  loc_110BE568: Else
  loc_110BE56E:   GoTo loc_110BE55E
  loc_110BE59C:   Exit Sub
  loc_110BE59D: End If
End Sub

Private Sub Form_Load() '110AB700
  Dim var_1C As Variant
  Dim var_24 As var_20.DispID_03E8
  Dim var_20 As var_1C.DispID_03E8
  loc_110AB766: If var_18 <= 3 Then
  loc_110AB789:   var_18 = frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110AB7B5:   var_34 = var_20.DispID_03E8
  loc_110AB7CA:   Set var_24 = var_20.DispID_03E8
  loc_110AB81D:   var_18 = 1+var_18
  loc_110AB822:   GoTo loc_110AB75D
  loc_110AB827: End If
  loc_110AB840: Set var_1C = frmZGCGToPz.TDBDate
  loc_110AB847: var_34 = var_1C.DispID_03E8
  loc_110AB85C: Set var_20 = var_1C.DispID_03E8
  loc_110AB868: var_20.UnkVCall_00000030h
  loc_110AB8D7: frmZGCGToPz.TDBDate.DispID_0000 = Date
  loc_110AB8F9: Set var_1C = frmZGCGToPz.APB
  loc_110AB906: var_1C.UnkVCall_00000040h
  loc_110AB944: var_20.DispID_80010007 = var_1C.DispID_03E8
  loc_110AB96B: Set var_1C = frmZGCGToPz.APB
  loc_110AB978: var_1C.UnkVCall_00000040h
  loc_110AB9B3: var_20.DispID_80010007 = var_1C.DispID_03E8
  loc_110AB9CF: var_8004 = frmZGCGToPz.Proc_15_10_110A4830(var_1C)
  loc_110AB9DC: var_58 = frmZGCGToPz.getBTData
  loc_110ABA04: GoTo loc_110ABA27
  loc_110ABA26: Exit Sub
  loc_110ABA27: ' Referenced from: 110ABA04
End Sub

Private Sub Form_Resize() '110ABA50
  loc_110ABADD: var_38 = frmZGCGToPz.Pic1.DispID_80010005
  loc_110ABB01: var_48 = frmZGCGToPz.Pic1.DispID_80010006
  loc_110ABB14: var_EC = var_48.ScaleWidth
  loc_110ABB4B: If global_110F6000 = 0 Then
  loc_110ABB55: Else
  loc_110ABB60: End If
  loc_110ABB60: var_70 = ((var_EC - CSgn(var_38)) / 2)
  loc_110ABB75: var_F0 = var_48.ScaleHeight
  loc_110ABBB3: If global_110F6000 = 0 Then
  loc_110ABBBD: Else
  loc_110ABBC8: End If
  loc_110ABCD3: frmZGCGToPz.Pic1.DispID_80011002(var_70, ((var_F0 - CSgn(var_48)) / 2), CSgn(frmZGCGToPz.Pic1.DispID_80010005), CSgn(frmZGCGToPz.Pic1.DispID_80010006))
  loc_110ABD1C: GoTo loc_110ABD56
End Sub

Private  APB_UnknownEvent_9(arg_C) '110BC420
  Dim var_24 As Variant
  loc_110BC49D: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110BC4A6: var_E8 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110BC4CD: arg_C = frmZGCGToPz.APB.UnkVCall_00000040h
  loc_110BC511: var_D4 = var_28.DispID_FFFFFDFA
  loc_110BC53B: var_8008 = (var_D4 = "加载数据")
  loc_110BC543: If var_8008 = 0 Then
  loc_110BC574:   Set var_24 = frmZGCGToPz.TDBText
  loc_110BC582:   var_C0 = var_24
  loc_110BC588:   var_24.UnkVCall_00000040h
  loc_110BC5BF:   var_EC = var_1C
  loc_110BC5DA:   var_20 = var_28.DispID_0000
  loc_110BC5EA:   var_8014 = Scripting.FileSystemObject.FileExists
  loc_110BC632:   If Not (var_BC) Then
  loc_110BC67D:     var_80 = "文件不存在或非法路径！ "
  loc_110BC69E:     MsgBox(var_80, 64, "提示", 10, 10)
  loc_110BC6C4:   Else
  loc_110BC6D4:     If var_18 > 2 Then GoTo loc_110BC8B0
  loc_110BC6FB:     var_18 = frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110BC732:     var_40 = var_28.DispID_0000
  loc_110BC78D:     If (Proc_0_11_11029000(8, var_28, var_20) = global_1100AE28) + 1 = 0 Then
  loc_110BC79E:       var_18 = 1+var_18
  loc_110BC7A1:       GoTo loc_110BC6CB
  loc_110BC7A6:     End If
  loc_110BC7C7:     var_18 = frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110BC81B:     var_80 = "提示"
  loc_110BC863:     MsgBox(var_28.DispID_8001004A & "不能为空，请输入。 ", 64, var_80, 10, 10)
  loc_110BC8A1:   End If
  loc_110BC8AB:   GoTo loc_110BCCB3
  loc_110BC90A:   If (frmZGCGToPz.Chk1.Value = 0) Then
  loc_110BC91E:     Set var_24 = frmZGCGToPz.TDBText
  loc_110BC92F:     var_24.UnkVCall_00000040h
  loc_110BC963:     var_40 = var_28.DispID_0000
  loc_110BC9C1:     If (Proc_0_11_11029000(8, var_24, 3) = global_1100AE28) + 1 Then
  loc_110BC9E6:       frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110BCA3A:       var_80 = "提示"
  loc_110BCA82:       MsgBox(var_28.DispID_8001004A & "不能为空，请输入。 ", 64, var_80, 10, 10)
  loc_110BCAB5:       GoTo loc_110BC896
  loc_110BCABA:     End If
  loc_110BCABA:   End If
  loc_110BCACC:   If frmZGCGToPz.FillData >= 0 Then GoTo loc_110BC8A1
  loc_110BCADE:   var_BC = CheckObj(8, global_1100D22C, 1788)
  loc_110BCAE9: End If
  loc_110BCAF5: var_8040 = (var_D4 = "取消加载")
  loc_110BCAFD: If var_8040 = 0 Then
  loc_110BCB22:   var_80 = "提示信息"
  loc_110BCB33:   var_48 = var_80
  loc_110BCB5C:   var_30 = "是否取消数据载入？" & vbCrLf & "取消数据载入，数据将全部清空。"
  loc_110BCB7B:   MsgBox(var_30, 292, var_48, var_58, var_68)
  loc_110BCBB2:   If (MsgBox(var_30, 292, var_48, var_58, var_68) = 6) = 8 Then GoTo loc_110BC8A3
  loc_110BCBBE:   GoTo loc_110BC8A3
  loc_110BCBC3: End If
  loc_110BCBD5: var_804C = (var_D4 = "凭证导入")
  loc_110BCBD9: If var_804C = 0 Then
  loc_110BCBDE:   var_8050 = frmZGCGToPz.Proc_15_13_110B3300(var_30)
  loc_110BCBE4:   GoTo loc_110BC8A3
  loc_110BCBE9: End If
  loc_110BCBF5: var_8054 = (var_D4 = "导出")
  loc_110BCBF9: If var_8054 = 0 Then GoTo loc_110BC8A3
  loc_110BCC0F: If (var_D4 = global_1100EBD4) Then GoTo loc_110BC8A3
  loc_110BCC46: Set var_24 = CInt(8)
  loc_110BCC54: var_8060 = Global.Unload var_58
  loc_110BCC75: GoTo loc_110BC8A3
  loc_110BCCB2: Exit Sub
  loc_110BCCB3: ' Referenced from: 110BC8AB
End Sub

Private Sub Chk1_Click() '110AB530
  loc_110AB5D1: If (frmZGCGToPz.Chk1.Value = 0) Then
  loc_110AB5F1:   frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110AB62D:   var_1C.DispID_8001000D = True
  loc_110AB63D: Else
  loc_110AB65A:   frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110AB6A4: End If
  loc_110AB6BB: GoTo loc_110AB6D1
  loc_110AB6D0: Exit Sub
  loc_110AB6D1: ' Referenced from: 110AB6BB
End Sub

Public Sub ExitForm(Cancel, UnloadMode) '110A4750
  Dim var_18 As Global
  loc_110A478F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110A47BA: Set var_18 = Me
  loc_110A47C2: var_8008 = Global.Unload
  loc_110A47FC: GoTo loc_110A4808
  loc_110A4807: Exit Sub
  loc_110A4808: ' Referenced from: 110A47FC
End Sub

Public Function FillData() '110A61D0
  Dim var_88 As Variant
  Dim var_60 As Variant
  Dim var_54 As Variant
  Dim var_38 As Variant
  Dim var_30 As Me
  Dim var_2C As ADODB.Recordset
  loc_110A62C9: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110A62DF: var_218 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110A6323: frmZGCGToPz.VFG.DispID_0007 = 1
  loc_110A6346: Set var_88 = frmZGCGToPz.Label3
  loc_110A6350: var_1F0 = var_88
  loc_110A6356: var_88.Caption = "正在打开Excel数据表，请稍候。。。"
  loc_110A63C9: frmZGCGToPz.Pic1.DispID_80010007 = True
  loc_110A63F5: frmZGCGToPz.Pic1.DispID_FFFFFDDA
  loc_110A640F: var_8004 = CreateObject(global_1100D5A4)
  loc_110A641A: Set var_60 = CreateObject(global_1100D5A4)
  loc_110A64D4: frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110A674D: var_80 = var_8C.DispID_0000
  loc_110A6763: var_80 = var_60.UnkVCall_000000D0h.UnkVCall_0000004Ch
  loc_110A67DA: var_88 = 0.Tag
  loc_110A6882: var_88.Activate
  loc_110A68E9: Set var_78 = var_88.UsedRange
  loc_110A6918: Set var_88 = frmZGCGToPz.Label3
  loc_110A6922: var_1F0 = var_88
  loc_110A6928: var_88.Caption = "正在填充数据，请稍候。。。"
  loc_110A699B: frmZGCGToPz.Pic1.DispID_80010007 = True
  loc_110A69C8: frmZGCGToPz.Pic1.DispID_FFFFFDDA
  loc_110A6A02: Set var_88 = frmZGCGToPz.APB
  loc_110A6A10: var_1F0 = var_88
  loc_110A6A16: var_88.UnkVCall_00000040h
  loc_110A6AAC: Set var_88 = frmZGCGToPz.APB
  loc_110A6ABA: var_1F0 = var_88
  loc_110A6AC0: var_88.UnkVCall_00000040h
  loc_110A6B66: frmZGCGToPz.APB.UnkVCall_00000040h
  loc_110A6C54: var_C4 = var_78.Rows.Count - 2
  loc_110A6CF6: frmZGCGToPz.sBar.DispID_6803001E(1100D68Ch & var_C4 & "条记录")
  loc_110A6D43: var_30 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[T_CY_ZGCG_Temp]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1) "
  loc_110A6D52: var_8014 = var_30 & "CREATE TABLE [T_CY_ZGCG_Temp](cVenCode VARCHAR(50) NULL,cVenName VARCHAR(50) NULL,cInvCode VARCHAR(50) NULL,cDepCode VARCHAR(50) NULL,iQuantity float NULL,iMoney Money NULL)"
  loc_110A6D5D: var_30 = var_8014
  loc_110A6DA2: var_A4 = UnkObj.UnkVCall_00000040h
  loc_110A6DE6: var_30 = "DELETE FROM [T_CY_ZGCG_Temp]"
  loc_110A6E7F: Set var_88 = frmZGCGToPz.TDBText
  loc_110A6E91: var_1F0 = var_88
  loc_110A6E97: var_88.UnkVCall_00000040h
  loc_110A6EE6: var_44 = Proc_0_11_11029000(9, var_88, 1)
  loc_110A6F14: Set var_88 = frmZGCGToPz.TDBText
  loc_110A6F26: var_1F0 = var_88
  loc_110A6F2C: var_88.UnkVCall_00000040h
  loc_110A6F7B: var_40 = Proc_0_11_11029000(9, var_88, 2)
  loc_110A6FA9: Set var_88 = frmZGCGToPz.Chk1
  loc_110A6FB9: var_1F0 = var_88
  loc_110A700C: If (var_88.Value = 0) Then
  loc_110A7026:   Set var_88 = frmZGCGToPz.TDBText
  loc_110A7038:   var_1F0 = var_88
  loc_110A703E:   var_88.UnkVCall_00000040h
  loc_110A7069:   var_8C = 0
  loc_110A7073:   var_9C = var_8C
  loc_110A708D:   var_24 = Proc_0_11_11029000(9, var_88, 3)
  loc_110A70B0: Else
  loc_110A70B7: End If
  loc_110A70CF: Set var_88 = frmZGCGToPz.TDBDate
  loc_110A70FB: var_AC = var_88.DispID_004E
  loc_110A710B: var_C4)
  loc_110A7169: var_74 = CByte("DateToPeriod".00000001h)
  loc_110A71D3: var_B4 = var_78.Rows.Count
  loc_110A722A: If var_18 <= CLng(var_B4 + 1) Then
  loc_110A7238:   If global_56 = 0 Then
  loc_110A72A8:     var_38.UnkVCall_00000064h
  loc_110A7404:     var_264 = (Proc_0_11_11029000(var_8C.Cells(var_18, 1).value, var_8C, var_38) = "汇总") + 1
  loc_110A7486:     var_80 = Proc_0_11_11029000(var_88.Cells(var_18, 3).value, 2, var_110)
  loc_110A74B0:     var_1F8 = (var_80 = global_1100AE28) + 1
  loc_110A74FE:     If var_1F8 = 0 Then
  loc_110A7531:       var_80 = CStr(var_18(-2))
  loc_110A75C5:       Set var_88 = frmZGCGToPz.sBar
  loc_110A75CC:       var_88.DispID_6803001E("正在填充数据：" & var_80 & "条记录")
  loc_110A7675:       var_38.UnkVCall_00000064h
  loc_110A7711:       var_B4 = var_88.Cells(var_18, 1).value
  loc_110A771B:       var_8050 = Proc_0_10_11028DD0(var_B4, "INSERT INTO [T_CY_ZGCG_Temp](cVenCode,cVenName,cInvCode,cDepCode,iQuantity,iMoney) VALUES (", var_38)
  loc_110A7725:       var_80 = var_8050
  loc_110A7884:       var_B4 = var_88.Cells(var_18, 2).value
  loc_110A7898:       var_80 = Proc_0_10_11028DD0(var_B4, 2 & var_80 & global_1100AC40, var_88)
  loc_110A79FB:       var_B4 = var_88.Cells(var_18, 3).value
  loc_110A7A0F:       var_80 = Proc_0_10_11028DD0(var_B4, var_110 & var_80 & global_1100AC40, var_88)
  loc_110A7A90:       var_80 = Proc_0_10_11028DD0(&H4008, var_108, var_88)
  loc_110A7B95:       var_B4 = var_88.Cells(var_18, 8).value
  loc_110A7BDB:       var_15C = (var_5C = 0)
  loc_110A7C46:       var_807C = vbNull & var_80 & global_1100AC40 & IIf((var_5C = 0), var_80, Proc_0_10_11028DD0(var_B4, var_88, var_24)) & 1100AC40h
  loc_110A7DD7:       var_B4 = var_88.Cells(var_18, 5).value
  loc_110A7FE0:       var_30 = var_807C & var_88.Cells(var_18, 7).value & 1100AC40h & Format(var_88.Cells(var_18, 7).value, "0.00") & 1100BD88h
  loc_110A80AF:       var_28 = var_28(1)
  loc_110A80BA:       If var_18 Mod 00000064h = 0 Then
  loc_110A80BC:         DoEvents
  loc_110A80C2:       End If
  loc_110A80D2:       var_18 = 1+var_18
  loc_110A80D5:       GoTo loc_110A7224
  loc_110A80DA:     End If
  loc_110A8129:     frmZGCGToPz.VFG.DispID_0007 = 1
  loc_110A8155:     global_56 = 0
  loc_110A8169:     Set var_88 = frmZGCGToPz.APB
  loc_110A817B:     var_1F0 = var_88
  loc_110A8181:     var_88.UnkVCall_00000040h
  loc_110A821A:     Set var_88 = frmZGCGToPz.APB
  loc_110A822C:     var_1F0 = var_88
  loc_110A8232:     var_88.UnkVCall_00000040h
  loc_110A82CB:     Set var_88 = frmZGCGToPz.APB
  loc_110A82DD:     var_1F0 = var_88
  loc_110A82E3:     var_88.UnkVCall_00000040h
  loc_110A82EA:     If var_88.UnkVCall_00000040h < 0 Then
  loc_110A82F0:       GoTo loc_110A84A0
  loc_110A82F5:     End If
  loc_110A831D:     Set var_88 = frmZGCGToPz.APB
  loc_110A832F:     var_1F0 = var_88
  loc_110A8335:     var_88.UnkVCall_00000040h
  loc_110A83CE:     Set var_88 = frmZGCGToPz.APB
  loc_110A83E0:     var_1F0 = var_88
  loc_110A83E6:     var_88.UnkVCall_00000040h
  loc_110A847F:     Set var_88 = frmZGCGToPz.APB
  loc_110A8491:     var_1F0 = var_88
  loc_110A8497:     var_88.UnkVCall_00000040h
  loc_110A849E:     If var_88.UnkVCall_00000040h < 0 Then
  loc_110A84A0:       ' Referenced from: 110A82F0
  loc_110A84AF:       var_88.UnkVCall_00000040h = CheckObj(var_1F0, global_1100D678, 64)
  loc_110A84B5:     End If
  loc_110A84B5:   End If
  loc_110A84E9:   var_8C.DispID_80010007 = var_10C
  loc_110A8508: End If
  loc_110A851C: var_80 = CStr(var_74)
  loc_110A8540: var_64 = "暂估采购" & var_80 & global_1100D708
  loc_110A857B: var_1F0 = var_2C
  loc_110A8581: var_1EC = ADODB.Recordset.State
  loc_110A85AC: If var_1EC = 1 Then
  loc_110A85CA:   var_1F0 = var_2C
  loc_110A85D0:   var_809C = ADODB.Recordset.Close
  loc_110A85F4: End If
  loc_110A8608: var_80 = "SELECT '" & var_44
  loc_110A8645: var_80AC = var_80 & "' AS cCode," & "cVenCode,cVenName,cInvCode,cDepCode,SUM(iQuantity) AS iQuantity,SUM(iMoney) AS iMoney " & "FROM [T_CY_ZGCG_Temp] GROUP BY cVenCode,cVenName,cInvCode,cDepCode ORDER BY cVenCode,cInvCode "
  loc_110A86BA: var_1F0 = var_2C
  loc_110A86F7: var_80B4 = ADODB.Recordset.Open(var_80AC, var_110, var_80AC, var_108, 9)
  loc_110A873E: var_1F0 = var_2C
  loc_110A8744: var_1E8 = ADODB.Recordset.EOF
  loc_110A876D: If var_1E8 = 0 Then
  loc_110A8793:   var_1F0 = var_2C
  loc_110A87D4:   var_1F8 = ADODB.Recordset.Fields
  loc_110A880A:   ADODB.Recordset.8 = Forms
  loc_110A8859:   var_80 = Proc_0_11_11029000(9, var_110, "cVenName")
  loc_110A889F:   If (var_80 = global_1100AE28) Then
  loc_110A88E7:     var_1F0 = var_2C
  loc_110A8916:     var_10C = "cVenName"
  loc_110A8928:     var_1F8 = ADODB.Recordset.Fields
  loc_110A895E:     ADODB.Recordset.8 = Forms
  loc_110A898C:     var_200 = var_8C
  loc_110A89D3:     var_4C = var_64 & "/" & var_8C
  loc_110A8A07:   End If
  loc_110A8CDC:   var_C4 = "1" & Chr(9) & 1100AE28h & Chr(9) & frmZGCGToPz.TDBDate.DispID_004E & Chr(9) & 1100D6D4h & Chr(9) & 1100C008h & Chr(9) & var_4C
  loc_110A8D4D:   var_1F0 = var_2C
  loc_110A8D8E:   var_1F8 = ADODB.Recordset.Fields
  loc_110A8DC4:   ADODB.Recordset.8 = Forms
  loc_110A8ECF:   var_1F0 = var_2C
  loc_110A8F1A:   var_1F8 = ADODB.Recordset.Fields
  loc_110A8F46:   ADODB.Recordset.8 = Forms
  loc_110A9068:   var_C4 = 9 & Chr(9) & Proc_0_11_11029000(9, var_120, "cCode") & Chr(9) & Proc_0_11_11029000(9, var_120, "iMoney") & Chr(9) & 1100C008h
  loc_110A90D9:   var_1F0 = var_2C
  loc_110A9124:   var_1F8 = ADODB.Recordset.Fields
  loc_110A9150:   ADODB.Recordset.8 = Forms
  loc_110A93F9:   var_B4 = 9 & Chr(9) & Proc_0_11_11029000(9, var_120, "iQuantity") & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9)
  loc_110A951D:   var_8128 = var_B4 & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110A958B:   var_1F0 = var_2C
  loc_110A95CC:   var_1F8 = ADODB.Recordset.Fields
  loc_110A9602:   ADODB.Recordset.8 = Forms
  loc_110A9834:   var_C4 = var_8128 & Chr(9) & Proc_0_11_11029000(9, var_120, "cDepCode") & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110A992D:   var_1F0 = var_2C
  loc_110A996E:   var_1F8 = ADODB.Recordset.Fields
  loc_110A99A4:   ADODB.Recordset.8 = Forms
  loc_110A9A32:   var_50 = 9 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_120, "cInvCode")
  loc_110A9A86:   var_1F0 = var_2C
  loc_110A9A8C:   var_8158 = ADODB.Recordset.MoveNext
  loc_110A9B02:   frmZGCGToPz.VFG.DispID_0080(var_50)
  loc_110A9B17:   GoTo loc_110A871B
  loc_110A9B1C: End If
  loc_110A9B3C: var_1F0 = var_2C
  loc_110A9B42: var_1EC = ADODB.Recordset.State
  loc_110A9B6D: If var_1EC = 1 Then
  loc_110A9B8B:   var_1F0 = var_2C
  loc_110A9B91:   var_8164 = ADODB.Recordset.Close
  loc_110A9BB5: End If
  loc_110A9C06: var_8174 = "SELECT '" & var_40 & "' AS cCode," & "cVenCode,cVenName,SUM(iMoney) AS iMoney " & "FROM [T_CY_ZGCG_Temp] GROUP BY cVenCode,cVenName ORDER BY cVenCode  "
  loc_110A9C7B: var_1F0 = var_2C
  loc_110A9CB8: var_817C = ADODB.Recordset.Open(var_8174, var_110, var_8174, var_108, 9)
  loc_110A9CFF: var_1F0 = var_2C
  loc_110A9D05: var_1E8 = ADODB.Recordset.EOF
  loc_110A9D2E: If var_1E8 = 0 Then
  loc_110A9D54:   var_1F0 = var_2C
  loc_110A9D95:   var_1F8 = ADODB.Recordset.Fields
  loc_110A9DCB:   ADODB.Recordset.8 = Forms
  loc_110A9E60:   If (Proc_0_11_11029000(9, var_110, "cVenName") = global_1100AE28) Then
  loc_110A9EA8:     var_1F0 = var_2C
  loc_110A9ED7:     var_10C = "cVenName"
  loc_110A9EE9:     var_1F8 = ADODB.Recordset.Fields
  loc_110A9F1F:     ADODB.Recordset.8 = Forms
  loc_110A9F4D:     var_200 = var_8C
  loc_110A9F94:     var_4C = var_64 & "/" & var_8C
  loc_110A9FC8:   End If
  loc_110AA29D:   var_C4 = "1" & Chr(9) & 1100AE28h & Chr(9) & frmZGCGToPz.TDBDate.DispID_004E & Chr(9) & 1100D6D4h & Chr(9) & 1100C008h & Chr(9) & var_4C
  loc_110AA30E:   var_1F0 = var_2C
  loc_110AA34F:   var_1F8 = ADODB.Recordset.Fields
  loc_110AA385:   ADODB.Recordset.8 = Forms
  loc_110AA518:   var_1F0 = var_2C
  loc_110AA559:   var_1F8 = ADODB.Recordset.Fields
  loc_110AA58F:   ADODB.Recordset.8 = Forms
  loc_110AA60F:   var_E4 = 9 & Chr(9) & Proc_0_11_11029000(9, var_120, "cCode") & Chr(9) & 1100C008h & Chr(9) & Proc_0_11_11029000(9, var_120, "iMoney")
  loc_110AA959:   var_C4 = var_E4 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110AAA9E:   var_10C = var_C4 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110AABEA:   var_1F0 = var_2C
  loc_110AAC2B:   var_1F8 = ADODB.Recordset.Fields
  loc_110AAC61:   ADODB.Recordset.8 = Forms
  loc_110AAC85:   var_8C = 0
  loc_110AAC8F:   var_BC = var_8C
  loc_110AACE1:   var_E4 = var_10C & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_120, "cVenCode")
  loc_110AAE19:   var_50 = var_E4 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110AAE53:   var_1F0 = var_2C
  loc_110AAE59:   var_8210 = ADODB.Recordset.MoveNext
  loc_110AAECF:   frmZGCGToPz.VFG.DispID_0080(var_50)
  loc_110AAEE4:   GoTo loc_110A9CDC
  loc_110AAEE9: End If
  loc_110AAF0B: var_1EC = ADODB.Recordset.State
  loc_110AAF30: If var_1EC = 1 Then
  loc_110AAF50:   var_821C = ADODB.Recordset.Close
  loc_110AAF6E: End If
  loc_110AB029: frmZGCGToPz.sBar.DispID_6803001E("有效数据共" & CStr(var_28) & global_1100FE7C)
  loc_110AB091: frmZGCGToPz.APB.UnkVCall_00000040h
  loc_110AB123: Set var_88 = frmZGCGToPz.APB
  loc_110AB131: var_1F0 = var_88
  loc_110AB137: var_88.UnkVCall_00000040h
  loc_110AB1C9: Set var_88 = frmZGCGToPz.APB
  loc_110AB1D7: var_1F0 = var_88
  loc_110AB1DD: var_88.UnkVCall_00000040h
  loc_110AB28D: frmZGCGToPz.Pic1.DispID_80010007 = var_10C
  loc_110AB2AC: Set var_88 = frmZGCGToPz.TDBText
  loc_110AB2BC: var_88.UnkVCall_00000040h
  loc_110AB314: var_9C = var_8C
  loc_110AB38C: var_88.ForeColor = False
  loc_110AB3C5: var_110 = var_60.UnkVCall_00000398h
  loc_110AB3FA: Set var_38 = {000208D7-0000-0000-C000000000000046}()
  loc_110AB40A: Set var_54 = {000208DA-0000-0000-C000000000000046}()
  loc_110AB41A: Set var_60 = {000208D5-0000-0000-C000000000000046}()
  loc_110AB42E: GoTo loc_110AB4A4
  loc_110AB4A3: Exit Function
  loc_110AB4A4: ' Referenced from: 110AB42E
End Function

Public Function getWBHL(sWhere) '110BCCF0
  Dim var_1C As ADODB.Recordset
  Dim var_2C As Me
  loc_110BCD50: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110BCD5C: var_98 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110BCD84: var_40 = Trim(sWhere)
  loc_110BCDB5: If (var_40 <> 1100AE28h) Then
  loc_110BCDE3:   var_20 = "SELECT * FROM exch WHERE 1=1 " & " AND " & sWhere
  loc_110BCDF0: Else
  loc_110BCDFC: End If
  loc_110BCE0C: var_20 = var_20 & " order by cexch_name, itype, iperiod, cdate"
  loc_110BCE76: var_78 = var_1C
  loc_110BCE85: var_8018 = ADODB.Recordset.Open(var_20, var_5C, var_20, var_54, 9)
  loc_110BCEEB: If ADODB.Recordset.EOF Then
  loc_110BCEFA:   var_24 = CStr(0)
  loc_110BCF05: Else
  loc_110BCF27:   var_2C = ADODB.Recordset.Fields
  loc_110BCF54:   var_58 = "NFLAT"
  loc_110BCF6D:   ADODB.Recordset.8 = Forms
  loc_110BCFBE:   var_24 = var_40
  loc_110BCFE0: End If
  loc_110BCFFE: var_8030 = ADODB.Recordset.Close
  loc_110BD01D: GoTo loc_110BD05B
  loc_110BD023: If var_4 Then
  loc_110BD02E: End If
  loc_110BD05A: Exit Function
  loc_110BD05B: ' Referenced from: 110BD01D
End Function

Public Function getBTData() '110BE5D0
  Dim var_24 As ADODB.Recordset
  Dim var_38 As Variant
  loc_110BE660: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110BE66A: On Error GoTo loc_110BEEF3
  loc_110BE6A5: var_28 = 1 & "IF NOT EXISTS (SELECT * FROM [" & "]..Sysobjects "
  loc_110BE710: var_8018 =  & var_28 & "WHERE Name = 'T_CY_ZGCG_Setting') " & "CREATE TABLE [" & "]..[T_CY_ZGCG_Setting](cJFKmCode VARCHAR(50) NULL," & "cDFKmCode VARCHAR(50) NULL,cDepCode VARCHAR(50) NULL,bDep Bit NOT NULL)"
  loc_110BE717: var_28 = var_8018
  loc_110BE747: var_54 = UnkObj.UnkVCall_00000040h
  loc_110BE799: var_28 = var_38 & "SELECT * FROM [" & "]..[T_CY_ZGCG_Setting]"
  loc_110BE7D3: var_EC = ADODB.Recordset.State
  loc_110BE7F8: If var_EC = 1 Then
  loc_110BE814:   var_802C = ADODB.Recordset.Close
  loc_110BE832: End If
  loc_110BE8BE: var_8034 = ADODB.Recordset.Open(var_28, var_B0, var_28, var_A8, 9)
  loc_110BE911: var_E8 = ADODB.Recordset.EOF
  loc_110BE92D: If var_E8 = 0 Then
  loc_110BE955:   var_38 = ADODB.Recordset.Fields
  loc_110BE973:   var_AC = "cjfkmcode"
  loc_110BE9A7:   ADODB.Recordset.8 = Forms
  loc_110BEA12:   frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110BEA48:   var_44.DispID_0000 = Proc_0_11_11029000(9, var_B0, "cjfkmcode")
  loc_110BEAAE:   var_100 = ADODB.Recordset.Fields
  loc_110BEAB9:   var_AC = "cDFKmCode"
  loc_110BEAED:   ADODB.Recordset.8 = Forms
  loc_110BEB5B:   frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110BEB91:   var_44.DispID_0000 = Proc_0_11_11029000(9, var_B0, "cDFKmCode")
  loc_110BEBF7:   var_100 = ADODB.Recordset.Fields
  loc_110BEC02:   var_AC = "cDepCode"
  loc_110BEC36:   ADODB.Recordset.8 = Forms
  loc_110BECA4:   frmZGCGToPz.TDBText.UnkVCall_00000040h
  loc_110BECDA:   var_44.DispID_0000 = Proc_0_11_11029000(9, var_B0, "cDepCode")
  loc_110BED29:   var_38 = ADODB.Recordset.Fields
  loc_110BED47:   var_AC = "bDep"
  loc_110BED7B:   ADODB.Recordset.8 = Forms
  loc_110BEE5C:   frmZGCGToPz.Chk1.Value = CInt(IIf((0 = True), 1, 0))
  loc_110BEEB2: End If
  loc_110BEEDA: If ADODB.Recordset.Close < 0 Then
  loc_110BEEEC:   var_8064 = CheckObj(var_24, global_1100ADFC, 128)
  loc_110BEEF3:   ' Referenced from: 110BE66A
  loc_110BEEF8:   var_8068 = Err
  loc_110BEF03:   Set var_38 = Err
  loc_110BEF88:   MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_110BEFB5: End If
  loc_110BEFB5: Exit Sub
  loc_110BEFC0: GoTo loc_110BF017
  loc_110BF016: Exit Function
  loc_110BF017: ' Referenced from: 110BEFC0
End Function

Public Function UpdateBTData() '110BF060
  Dim var_48 As Variant
  Dim var_50 As frmZGCGToPz.TDBText
  Dim var_58 As frmZGCGToPz.TDBText
  Dim var_20 As Me
  loc_110BF0F0: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110BF0F8: var_FC = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110BF100: On Error GoTo loc_110BF4C6
  loc_110BF13B: var_20 = 1 & "DELETE FROM [" & "]..[T_CY_ZGCG_Setting]"
  loc_110BF17A: var_6C = UnkObj.UnkVCall_00000040h
  loc_110BF203: Set var_48 = frmZGCGToPz.TDBText
  loc_110BF209: var_D4 = var_48
  loc_110BF218: var_48.UnkVCall_00000040h
  loc_110BF258: Set var_50 = frmZGCGToPz.TDBText
  loc_110BF25E: var_DC = var_50
  loc_110BF26D: var_50.UnkVCall_00000040h
  loc_110BF28E: var_54 = 0
  loc_110BF295: var_74 = var_54
  loc_110BF2AD: Set var_58 = frmZGCGToPz.TDBText
  loc_110BF2B3: var_E4 = var_58
  loc_110BF2C2: var_58.UnkVCall_00000040h
  loc_110BF2E3: var_5C = 0
  loc_110BF2EA: var_84 = var_5C
  loc_110BF302: var_8018 = Proc_0_10_11028DD0(9, var_48 & "INSERT INTO [" & "]..[T_CY_ZGCG_Setting]" & "(cJFKmCode,cDFKmCode,cDepCode,bDep) VALUES (", var_58)
  loc_110BF366: var_8034 = var_54 & Proc_0_10_11028DD0(9, var_50 & Proc_0_10_11028DD0(9, 3 & var_8018 & global_1100AC40, var_5C) & global_1100AC40, 2)
  loc_110BF441: var_20 = var_8034 & global_1100AC40 & CStr(frmZGCGToPz.Chk1.Value) & global_1100BD88
  loc_110BF4C1: GoTo loc_110BF597
  loc_110BF4C6: ' Referenced from: 110BF100
  loc_110BF4CB: var_8048 = Err
  loc_110BF4D6: Set var_48 = Err
  loc_110BF567: MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_110BF597: ' Referenced from: 110BF4C1
  loc_110BF597: Exit Sub
  loc_110BF5A2: GoTo loc_110BF611
  loc_110BF610: Exit Function
  loc_110BF611: ' Referenced from: 110BF5A2
End Function

Private Sub Proc_15_10_110A4830
  Dim var_58 As frmZGCGToPz.VFG
  loc_110A4871: Set var_58 = frmZGCGToPz.VFG
  loc_110A48C2: var_58.DispID_005D = frmZGCGToPz.VFG
  loc_110A4903: var_58.DispID_0067 = frmZGCGToPz.VFG
  loc_110A4922: var_58.DispID_0041 = frmZGCGToPz.VFG
  loc_110A49CC: var_58.DispID_00A5("...")
  loc_110A4AF4: var_58.DispID_008A(4)
  loc_110A4B37: var_58.DispID_0079(450)
  loc_110A4B5B: var_58.DispID_0019 = True
  loc_110A4B9B: var_58.DispID_007B(True)
  loc_110A4BE4: var_58.DispID_009D(5)
  loc_110A4C29: var_58.DispID_0090("业务号")
  loc_110A4C6C: var_58.DispID_0077(4)
  loc_110A4CAF: var_58.DispID_0078(700)
  loc_110A4CF7: var_58.DispID_0090("状态")
  loc_110A4D3D: var_58.DispID_0077(4)
  loc_110A4D83: var_58.DispID_0078(700)
  loc_110A4DCB: var_58.DispID_0090("制单日期")
  loc_110A4E11: var_58.DispID_0077(1)
  loc_110A4E57: var_58.DispID_0078(1000)
  loc_110A4E9C: var_58.DispID_0090("凭证类别字")
  loc_110A4EDE: var_58.DispID_0077(4)
  loc_110A4F20: var_58.DispID_0078(700)
  loc_110A4F68: var_58.DispID_0090("附单据数")
  loc_110A4FAC: var_58.DispID_0077(var_3C)
  loc_110A4FF2: var_58.DispID_0078(var_3C)
  loc_110A503A: var_58.DispID_0090(var_3C)
  loc_110A5080: var_58.DispID_0077(var_3C)
  loc_110A50C6: var_58.DispID_0078(var_3C)
  loc_110A510E: var_58.DispID_0090(var_3C)
  loc_110A5154: var_58.DispID_0077(var_3C)
  loc_110A519A: var_58.DispID_0078(var_3C)
  loc_110A51E2: var_58.DispID_0090(var_3C)
  loc_110A5226: var_58.DispID_0077(var_3C)
  loc_110A526C: var_58.DispID_0078(var_3C)
  loc_110A52B4: var_58.DispID_009C(var_3C)
  loc_110A52FC: var_58.DispID_0090(var_3C)
  loc_110A5342: var_58.DispID_0077(var_3C)
  loc_110A5388: var_58.DispID_0078(var_3C)
  loc_110A53D0: var_58.DispID_009C(var_3C)
  loc_110A5418: var_58.DispID_0090(var_3C)
  loc_110A545E: var_58.DispID_0077(var_3C)
  loc_110A54A4: var_58.DispID_0078(var_3C)
  loc_110A54EC: var_58.DispID_009C(var_3C)
  loc_110A5534: var_58.DispID_0090(var_3C)
  loc_110A557A: var_58.DispID_0077(var_3C)
  loc_110A55C0: var_58.DispID_0078(var_3C)
  loc_110A5608: var_58.DispID_009C(var_3C)
  loc_110A5650: var_58.DispID_0090(var_3C)
  loc_110A5696: var_58.DispID_0077(var_3C)
  loc_110A56DC: var_58.DispID_0078(var_3C)
  loc_110A5724: var_58.DispID_009C(var_3C)
  loc_110A576C: var_58.DispID_0090(var_3C)
  loc_110A57B2: var_58.DispID_0077(var_3C)
  loc_110A57F8: var_58.DispID_0078(var_3C)
  loc_110A5840: var_58.DispID_0090(var_3C)
  loc_110A5886: var_58.DispID_0077(var_3C)
  loc_110A58CC: var_58.DispID_0078(var_3C)
  loc_110A5914: var_58.DispID_0090(var_3C)
  loc_110A595A: var_58.DispID_0077(var_3C)
  loc_110A59A0: var_58.DispID_0078(var_3C)
  loc_110A59E8: var_58.DispID_0090(var_3C)
  loc_110A5A2E: var_58.DispID_0077(var_3C)
  loc_110A5A74: var_58.DispID_0078(var_3C)
  loc_110A5ABC: var_58.DispID_0090(var_3C)
  loc_110A5B02: var_58.DispID_0077(var_3C)
  loc_110A5B48: var_58.DispID_0078(var_3C)
  loc_110A5B90: var_58.DispID_0090(var_3C)
  loc_110A5BD6: var_58.DispID_0077(var_3C)
  loc_110A5C1C: var_58.DispID_0078(var_3C)
  loc_110A5C64: var_58.DispID_0090(var_3C)
  loc_110A5CAA: var_58.DispID_0077(var_3C)
  loc_110A5CF0: var_58.DispID_0078(var_3C)
  loc_110A5D38: var_58.DispID_0090(var_3C)
  loc_110A5D7E: var_58.DispID_0077(var_3C)
  loc_110A5DC4: var_58.DispID_0078(var_3C)
  loc_110A5E0C: var_58.DispID_0090(var_3C)
  loc_110A5E52: var_58.DispID_0077(var_3C)
  loc_110A5E98: var_58.DispID_0078(var_3C)
  loc_110A5EE0: var_58.DispID_0090(var_3C)
  loc_110A5F26: var_58.DispID_0077(var_3C)
  loc_110A5F6C: var_58.DispID_0078(var_3C)
  loc_110A5FB4: var_58.DispID_0090(var_3C)
  loc_110A5FFA: var_58.DispID_0077(var_3C)
  loc_110A6040: var_58.DispID_0078(var_3C)
  loc_110A605C: If 10 <= &H14 Then
  loc_110A609C:   var_58.DispID_00AC(var_3C)
  loc_110A60B4:   var_14 = 1+var_14
  loc_110A60B7:   GoTo loc_110A6058
  loc_110A60B9: End If
  loc_110A60F9: var_58.DispID_00AC(var_3C)
  loc_110A613E: var_58.DispID_00AC(var_3C)
  loc_110A6183: var_58.DispID_00AC(var_3C)
End Sub

Private Sub Proc_15_11_110ABD80
  Dim var_7C As Variant
  Dim var_1F8 As Label
  Dim var_80 As Variant
  Dim var_88 As frmZGCGToPz.Label3
  loc_110ABE6A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110ABE72: var_228 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110ABE78: var_8004 = ecx
  loc_110ABEEE: If var_14 <= CLng(frmZGCGToPz.VFG.DispID_0007)(-1) Then
  loc_110ABEFF:   var_800C = frmZGCGToPz.Proc_15_12_110ADC20(vbNull)
  loc_110ABF9D:   frmZGCGToPz.VFG.DispID_0082(22, var_58)
  loc_110AC081:   If (frmZGCGToPz.VFG.DispID_0082(var_14, 22) = global_1100AE28) + 1 Then
  loc_110AC101:     frmZGCGToPz.VFG.DispID_0082(1, 285267764)
  loc_110AC235:     frmZGCGToPz.VFG.DispID_009E(var_14, 1, var_14, 1, 16711680)
  loc_110AC255:     Set var_7C = frmZGCGToPz.Label3
  loc_110AC262:     var_1F8 = var_7C
  loc_110AC2AC:     var_7C.Caption = "分析: 第(" & CStr(vbNull) & ")行信息----有效"
  loc_110AC2FE:     frmZGCGToPz.Pic1.DispID_FFFFFDDA
  loc_110AC311:   Else
  loc_110AC38B:     frmZGCGToPz.VFG.DispID_0082(1, 285267820)
  loc_110AC4BF:     frmZGCGToPz.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110AC4DF:     Set var_80 = frmZGCGToPz.Label3
  loc_110AC4EC:     var_1F8 = var_80
  loc_110AC5CD:     var_80.Caption = "分析:   第(" & CStr(vbNull) & ")行信息----" & frmZGCGToPz.VFG.DispID_0082(var_14, 22)
  loc_110AC638:     frmZGCGToPz.Pic1.DispID_FFFFFDDA
  loc_110AC64A:   End If
  loc_110AC65A:   var_14 = 1+var_14
  loc_110AC65D:   GoTo loc_110ABEE0
  loc_110AC662: End If
  loc_110AC6C9: If var_14 <= CLng(frmZGCGToPz.VFG.DispID_0007)(-1) Then
  loc_110AC741:   var_A0 = frmZGCGToPz.VFG.DispID_0082(var_14, 2)
  loc_110AC75F:   var_B8)
  loc_110AC8EF:   var_8048 = frmZGCGToPz.VFG.DispID_0082(var_14, frmZGCGToPz.VFG)
  loc_110AC926:   var_4C = CCur(0)
  loc_110AC929:   var_48 = var_8048
  loc_110AC935:   var_40 = CCur(0)
  loc_110AC938:   var_3C = var_8048
  loc_110AC944:   var_34 = var_14
  loc_110AC94D:   var_30 = var_14
  loc_110AC956:   var_160 = CByte("DateToPeriod".00000001h)
  loc_110AC9F3:   var_B8)
  loc_110ACA72:   Set var_80 = frmZGCGToPz.VFG
  loc_110ACA98:   var_8064 = (frmZGCGToPz.VFG.DispID_0082(var_14, 3) = var_80.DispID_0082(var_14, 3))
  loc_110ACAC5:   var_1A0 = var_8064 + 1
  loc_110ACB3F:   var_806C = (var_8048 = frmZGCGToPz.VFG.DispID_0082(var_14, ""))
  loc_110ACB66:   var_1E0 = var_806C + 1
  loc_110ACC68:   If CBool((frmZGCGToPz.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_8064 + 1 And var_806C + 1) Then
  loc_110ACD2D:     If (frmZGCGToPz.VFG.DispID_0082(var_14, 22) = global_1100AE28) Then
  loc_110ACD36:     End If
  loc_110ACD3B:     If var_24 = 0 Then
  loc_110ACDE4:       var_16C = var_48
  loc_110ACE28:       var_9C = var_1F0
  loc_110ACE74:       var_4C = CCur(var_4C + Format(Val(frmZGCGToPz.VFG.DispID_0082(var_14, 7)), "#.00"))
  loc_110ACE77:       var_48 = var_D8
  loc_110ACF57:       var_16C = var_3C
  loc_110ACF9B:       var_9C = var_1F0
  loc_110ACFE7:       var_40 = CCur(var_40 + Format(Val(frmZGCGToPz.VFG.DispID_0082(var_14, 8)), "#.00"))
  loc_110ACFEA:       var_3C = var_D8
  loc_110AD02A:     End If
  loc_110AD04B:     var_14 = var_14(1)
  loc_110AD04E:     var_30 = var_30(1)
  loc_110AD070:     var_80A0 = CLng(frmZGCGToPz.VFG.DispID_0007)
  loc_110AD08B:     var_1F8 = (var_14 > 0)
  loc_110AD0AF:     If var_1F8 = 0 Then GoTo loc_110AC950
  loc_110AD0B5:   End If
  loc_110AD0BA:   If var_24 = 0 Then
  loc_110AD0CE:     Set var_7C = frmZGCGToPz.Chk
  loc_110AD0D9:     var_1F8 = var_7C
  loc_110AD0DF:     Set var_80 = var_7C(1)
  loc_110AD10A:     var_200 = var_80
  loc_110AD110:     var_1EC = var_80.Value
  loc_110AD164:     If (var_1EC = 1) Then
  loc_110AD194:       If (Abs(var_4C - var_40) <> 0.01) >= 0 Then
  loc_110AD19D:       End If
  loc_110AD19D:     End If
  loc_110AD1A2:     If var_24 Then
  loc_110AD1A8:     End If
  loc_110AD1C8:     var_1C = var_34
  loc_110AD1CD:     If var_34 <= (var_30 - 1) Then
  loc_110AD291:       If (frmZGCGToPz.VFG.DispID_0082(var_1C, 22) = global_1100AE28) + 1 Then
  loc_110AD319:         frmZGCGToPz.VFG.DispID_0082(1, 285267820)
  loc_110AD3AD:         frmZGCGToPz.VFG.DispID_0082(22, "凭证借贷不平衡或某分录有错误")
  loc_110AD4E1:         frmZGCGToPz.VFG.DispID_009E(var_1C, 1, var_1C, 1, 255)
  loc_110AD4F3:       End If
  loc_110AD503:       GoTo loc_110AD1C2
  loc_110AD508:     End If
  loc_110AD519:     var_44 = var_44(1)
  loc_110AD52A:     Set var_88 = frmZGCGToPz.Label3
  loc_110AD55D:     var_1F8 = var_88
  loc_110AD66A:     Set var_80 = frmZGCGToPz.VFG
  loc_110AD742:     var_80D4 = "分析: 第[" & frmZGCGToPz.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmZGCGToPz.VFG.DispID_0082(var_34, var_14)
  loc_110AD764:     var_78 = var_80D4 & "]号凭证借贷不平衡"
  loc_110AD778:     var_88.Caption = var_78
  loc_110AD77F:     If var_78 < 0 Then
  loc_110AD785:       GoTo loc_110ADA03
  loc_110AD78A:     End If
  loc_110AD79B:     var_20 = var_20(1)
  loc_110AD7AC:     Set var_88 = frmZGCGToPz.Label3
  loc_110AD7DF:     var_1F8 = var_88
  loc_110AD8EC:     Set var_80 = frmZGCGToPz.VFG
  loc_110AD9C4:     var_80F8 = "分析: 第[" & frmZGCGToPz.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmZGCGToPz.VFG.DispID_0082(var_34, frmZGCGToPz.VFG.DispID_0082(var_34, var_14))
  loc_110AD9E6:     var_78 = var_80F8 & "]号凭证有效"
  loc_110AD9FA:     var_88.Caption = var_78
  loc_110ADA01:     If var_78 >= 0 Then GoTo loc_110ADA12
  loc_110ADA03:     ' Referenced from: 110AD785
  loc_110ADA0C:     var_78 = CheckObj(var_1F8, global_1100D574, 84)
  loc_110ADA12:   End If
  loc_110ADA94:   frmZGCGToPz.Pic1.DispID_FFFFFDDA
  loc_110ADAC5:   var_14 = 1+var_14(-1)
  loc_110ADAC8:   GoTo loc_110AC6C3
  loc_110ADACD: End If
  loc_110ADAD2: If var_44 > 0 Then
  loc_110ADAD9:   If var_20 > 0 Then
  loc_110ADAF4:   Else
  loc_110ADB0D:   Else
  loc_110ADB17:     var_8108 = frmZGCGToPz.Proc_15_14_110BD0A0(var_1EC)
  loc_110ADB25:     If var_1EC Then
  loc_110ADB40:     Else
  loc_110ADB48:       var_18 = ecx
  loc_110ADB51:       GoTo loc_110ADBEB
  loc_110ADBEA:       Exit Sub
  loc_110ADBEB:     End If
  loc_110ADBEB:   End If
  loc_110ADBEB: End If
  loc_110ADBEB: ' Referenced from: 110ADB51
End Sub

Private  Proc_15_12_110ADC20(arg_C) '110ADC20
  Dim var_58 As frmZGCGToPz.VFG
  Dim var_20 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA19-EB96-11D2-AF06000021009B21}()
  Dim var_18 As {3302AA41-EB96-11D2-AF06000021009B21}()
  Dim var_1C As {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_110ADD1C: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110ADD2C: var_210 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110ADE0B: If (frmZGCGToPz.VFG.DispID_0082(arg_C, 2) = global_1100AE28) + 1 Then
  loc_110ADE15:   var_24 = "制单日期为空"
  loc_110ADE26: Else
  loc_110ADEC1:   var_78 = frmZGCGToPz.VFG.DispID_0082(arg_C, 2)
  loc_110ADEFB:   If Proc_0_9_11028500(var_80, global_110B32D2, ) Then
  loc_110ADFA4:     var_78 = frmZGCGToPz.VFG.DispID_0082(arg_C, 2)
  loc_110ADFAE:     var_90)
  loc_110ADFC0:     var_48 = var_90
  loc_110ADFF2:     var_118 = var_48
  loc_110AE000:     var_114 = var_44
  loc_110AE034:     var_80 = "AccountOpen".0.0
  loc_110AE065:     If (var_80 < var_80) Then
  loc_110AE06F:       var_24 = "日期超前总账系统启用日期"
  loc_110AE080:     Else
  loc_110AE086:       var_154 = var_44
  loc_110AE08C:       var_1A4 = var_44
  loc_110AE098:       var_158 = var_48
  loc_110AE09F:       var_1A8 = var_48
  loc_110AE14C:       var_80 = "AccountYMD".0.00000002h("AccountYMD".0, var_13C)
  loc_110AE246:       If CBool( Or ((global_110B32D2 < var_80) > "AccountYMD".0.00000002h(var_180, var_18C))) Then
  loc_110AE250:         var_24 = "日期必须在当前会计年度内"
  loc_110AE261:       Else
  loc_110AE27E:         var_118 = var_48
  loc_110AE2D2:         var_80 = "DateToPeriod".00000001h - 1
  loc_110AE360:         If CBool("AccountYMD".0.00000001h) Then
  loc_110AE36A:           var_24 = "已结账月份不能制单"
  loc_110AE37B:         Else
  loc_110AE457:           If (frmZGCGToPz.VFG.DispID_0082(arg_C, 3) = global_1100AE28) + 1 Then
  loc_110AE461:             var_24 = "凭证类别字为空"
  loc_110AE472:           Else
  loc_110AE501:             var_8034 = frmZGCGToPz.VFG.DispID_0082(arg_C, 3)
  loc_110AE511:             var_80 = 8
  loc_110AE514:             var_78 = var_8034
  loc_110AE55B:             var_8038 = CBool(Not("pzlbCheck".00000001h(, fs:[00000000h], , global_110B32D2, global_110B32D2, var_74, var_8034, var_7C)))
  loc_110AE592:             If var_8038 Then
  loc_110AE59C:               var_24 = "凭证类别字非法"
  loc_110AE5AD:             Else
  loc_110AE684:               If (frmZGCGToPz.VFG.DispID_0082(arg_C, var_128) = global_1100AE28) + 1 Then
  loc_110AE68E:                 var_24 = "业务号为空"
  loc_110AE69F:               Else
  loc_110AE729:                 var_8044 = frmZGCGToPz.VFG.DispID_0082(arg_C, var_128)
  loc_110AE739:                 var_80 = 8
  loc_110AE73C:                 var_78 = var_8044
  loc_110AE77F:                 var_90 = "GenLen".00000001h(fs:[00000000h], , global_110B32D2, global_110B32D2, global_110B32D2, var_74, var_8044, var_7C)
  loc_110AE7C7:                 If (var_90 > 30) Then
  loc_110AE7D1:                   var_24 = "业务号超长"
  loc_110AE7E2:                 Else
  loc_110AE8C1:                   If (frmZGCGToPz.VFG.DispID_0082(arg_C, 5) = global_1100AE28) + 1 Then
  loc_110AE8CB:                     var_24 = "摘要为空"
  loc_110AE8DC:                   Else
  loc_110AE997:                     var_8058 = InStr(1, frmZGCGToPz.VFG.DispID_0082(arg_C, 5), "|", 0)
  loc_110AE9BD:                     var_220 = (var_8058 > 0)
  loc_110AEA13:                     var_80 = frmZGCGToPz.VFG.DispID_0082(arg_C, 5)
  loc_110AEB34:                     If (((var_8058 > 0) Or (InStr(1, var_80, """", 0) > 0)) Or (InStr(1, frmZGCGToPz.VFG.DispID_0082(arg_C, 5), "'", 0) > 0)) Then
  loc_110AEB3E:                       var_24 = "摘要含有非法字符"
  loc_110AEB4F:                     Else
  loc_110AEBE1:                       var_806C = frmZGCGToPz.VFG.DispID_0082(arg_C, 5)
  loc_110AEBF1:                       var_80 = 8
  loc_110AEBF4:                       var_78 = var_806C
  loc_110AEC37:                       var_90 = "GenLen".00000001h(global_110B32D2, global_110B32D2, global_110B32D2, global_110B32D2, global_110B32D2, var_74, var_806C, var_7C)
  loc_110AEC80:                       If (var_90 > 120) Then
  loc_110AEC8A:                         var_24 = "摘要超长"
  loc_110AEC9B:                       Else
  loc_110AED78:                         If (frmZGCGToPz.VFG.DispID_0082(arg_C, 6) = global_1100AE28) + 1 Then
  loc_110AED82:                           var_24 = "科目为空"
  loc_110AED93:                         Else
  loc_110AEE22:                           var_807C = frmZGCGToPz.VFG.DispID_0082(arg_C, 6)
  loc_110AEE32:                           var_80 = 8
  loc_110AEE35:                           var_78 = var_807C
  loc_110AEEB5:                           var_40 = "kmCheck".00000002h(var_807C, var_150, var_15C)
  loc_110AEEE7:                           var_8084 = (var_40 = global_1100AE28)
  loc_110AEEEF:                           If var_8084 = 0 Then
  loc_110AEEF9:                             var_24 = "科目非法"
  loc_110AEF0A:                           Else
  loc_110AEF48:                             var_118 = arg_C
  loc_110AEFAF:                             frmZGCGToPz.VFG.DispID_0082(6, var_40)
  loc_110AEFC9:                             var_118 = var_40
  loc_110AF01B:                             var_128 = var_20
  loc_110AF069:                             "kmCodeToProperties".00000002h
  loc_110AF086:                             Set var_20 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110AF0B4:                             var_1F0 = var_20
  loc_110AF0BA:                             var_1D4 = var_20.UnkVCall_00000114h
  loc_110AF0E6:                             If var_1D4 = 0 Then
  loc_110AF0F0:                               var_24 = "科目非末级"
  loc_110AF101:                             Else
  loc_110AF1DF:                               If (frmZGCGToPz.VFG.DispID_0082(arg_C, 7) = global_1100AE28) Then
  loc_110AF2BB:                                 If Not (IsNumeric(frmZGCGToPz.VFG.DispID_0082(arg_C, 7))) Then
  loc_110AF2C5:                                   var_24 = "借方金额非法"
  loc_110AF2D6:                                 Else
  loc_110AF37F:                                   var_80A4 = CDbl(Val(frmZGCGToPz.VFG.DispID_0082(arg_C, 7)))
  loc_110AF41A:                                   var_80 = frmZGCGToPz.VFG.DispID_0082(arg_C, 7)
  loc_110AF442:                                   var_22C = CDbl(Val(var_80))
  loc_110AF458:                                   var_80B0 = CDbl(-9999999999999.99)
  loc_110AF470:                                   GoTo loc_110AF474
  loc_110AF4C2:                                   If (eax Or 0) Then
  loc_110AF4CC:                                     var_24 = "借方金额超范围"
  loc_110AF4DD:                                   Else
  loc_110AF4DD:                                   End If
  loc_110AF5BB:                                   If (frmZGCGToPz.VFG.DispID_0082(arg_C, 8) = global_1100AE28) Then
  loc_110AF697:                                     If Not (IsNumeric(frmZGCGToPz.VFG.DispID_0082(arg_C, 8))) Then
  loc_110AF6A1:                                       var_24 = "贷方金额非法"
  loc_110AF6B2:                                     Else
  loc_110AF75B:                                       var_80C8 = CDbl(Val(frmZGCGToPz.VFG.DispID_0082(arg_C, 8)))
  loc_110AF7F6:                                       var_80 = frmZGCGToPz.VFG.DispID_0082(arg_C, 8)
  loc_110AF81E:                                       var_238 = CDbl(Val(var_80))
  loc_110AF834:                                       var_80D4 = CDbl(-9999999999999.99)
  loc_110AF84C:                                       GoTo loc_110AF850
  loc_110AF89E:                                       If (eax Or 0) Then
  loc_110AF8A8:                                         var_24 = "贷方金额超范围"
  loc_110AF8B9:                                       Else
  loc_110AF8B9:                                       End If
  loc_110AFA31:                                       var_74 = var_1E0
  loc_110AFAA3:                                       var_C4 = var_1E8
  loc_110AFB1D:                                       var_80E8 = (Format(Val(frmZGCGToPz.VFG.DispID_0082(arg_C, 7)), "#.00") <> 0) And (Format(Val(frmZGCGToPz.VFG.DispID_0082(arg_C, 8)), "#.00") <> 0)
  loc_110AFB96:                                       If CBool(var_80E8) Then
  loc_110AFBA0:                                         var_24 = "借方金额和贷方金额不能同时不为0"
  loc_110AFBB1:                                       Else
  loc_110AFD29:                                         var_74 = var_1E0
  loc_110AFD9B:                                         var_C4 = var_1E8
  loc_110AFE15:                                         var_8100 = (Format(Val(frmZGCGToPz.VFG.DispID_0082(arg_C, 7)), "#.00") = 0) And (Format(Val(frmZGCGToPz.VFG.DispID_0082(arg_C, 8)), "#.00") = 0)
  loc_110AFE8E:                                         If CBool(var_8100) Then
  loc_110AFE98:                                           var_24 = "借方金额和贷方金额不能同时为0"
  loc_110AFEA9:                                         Else
  loc_110AFEC9:                                           var_1F0 = var_20
  loc_110AFF1B:                                           If (var_20.UnkVCall_0000007Ch = global_1100AE28) Then
  loc_110AFFFF:                                             If (frmZGCGToPz.VFG.DispID_0082(arg_C, 9) = global_1100AE28) Then
  loc_110B00DB:                                               If Not (IsNumeric(frmZGCGToPz.VFG.DispID_0082(arg_C, 9))) Then
  loc_110B00E5:                                                 var_24 = "数量数值非法"
  loc_110B00F6:                                               Else
  loc_110B00F6:                                               End If
  loc_110B00F6:                                             End If
  loc_110B0116:                                             var_1F0 = var_20
  loc_110B0168:                                             If (var_20.UnkVCall_0000006Ch = global_1100AE28) Then
  loc_110B024C:                                               If (frmZGCGToPz.VFG.DispID_0082(arg_C, 10) = global_1100AE28) Then
  loc_110B0328:                                                 If Not (IsNumeric(frmZGCGToPz.VFG.DispID_0082(arg_C, 10))) Then
  loc_110B0332:                                                   var_24 = "外币金额非法"
  loc_110B0343:                                                 Else
  loc_110B03EC:                                                   var_813C = CDbl(Val(frmZGCGToPz.VFG.DispID_0082(arg_C, 10)))
  loc_110B04AF:                                                   var_244 = CDbl(Val(frmZGCGToPz.VFG.DispID_0082(arg_C, 10)))
  loc_110B04C5:                                                   var_8148 = CDbl(-9999999999999.99)
  loc_110B04DD:                                                   GoTo loc_110B04E1
  loc_110B052F:                                                   If (eax Or 0) Then
  loc_110B0539:                                                     var_24 = "外币超范围"
  loc_110B054A:                                                   Else
  loc_110B054A:                                                   End If
  loc_110B0628:                                                   If (frmZGCGToPz.VFG.DispID_0082(arg_C, 11) = global_1100AE28) Then
  loc_110B0704:                                                     If Not (IsNumeric(frmZGCGToPz.VFG.DispID_0082(arg_C, 11))) Then
  loc_110B070E:                                                       var_24 = "汇率数值非法"
  loc_110B071F:                                                     Else
  loc_110B071F:                                                     End If
  loc_110B071F:                                                   End If
  loc_110B07FD:                                                   If (frmZGCGToPz.VFG.DispID_0082(arg_C, 12) = global_1100AE28) Then
  loc_110B0894:                                                     var_8164 = frmZGCGToPz.VFG.DispID_0082(arg_C, 12)
  loc_110B08A7:                                                     var_78 = var_8164
  loc_110B08EA:                                                     var_90 = "GenLen".00000001h(global_110B32D2, global_110B32D2, global_110B32D2, global_110B32D2, global_110B32D2, var_74, var_8164, var_7C)
  loc_110B0904:                                                     var_1F0 = (var_90 > 20)
  loc_110B0933:                                                     If var_1F0 = 0 Then GoTo loc_110B0A79
  loc_110B0941:                                                     var_24 = "制单人姓名超长"
  loc_110B0952:                                                   Else
  loc_110B0971:                                                     var_118 = arg_C
  loc_110B0A4D:                                                     frmZGCGToPz.VFG.DispID_0082(12, "UserCurrent".00000000h.00000000h)
  loc_110B0A9C:                                                     var_1F0 = var_20
  loc_110B0ACE:                                                     If var_20.UnkVCall_0000010Ch Then
  loc_110B0BB2:                                                       If (frmZGCGToPz.VFG.DispID_0082(arg_C, 13) = global_1100AE28) Then
  loc_110B0C49:                                                         var_817C = frmZGCGToPz.VFG.DispID_0082(arg_C, 13)
  loc_110B0C5C:                                                         var_78 = var_817C
  loc_110B0C8B:                                                         var_90 = "JsfsCheck".00000001h(1, global_110B32D2, global_110B32D2, global_110B32D2, global_110B32D2, var_74, var_817C, var_7C)
  loc_110B0CDB:                                                         If CBool(Not(var_90)) Then
  loc_110B0CE5:                                                           var_24 = "结算方式非法"
  loc_110B0CF6:                                                         Else
  loc_110B0CF6:                                                         End If
  loc_110B0CF6:                                                       End If
  loc_110B0D19:                                                       var_1F0 = var_20
  loc_110B0D1F:                                                       var_1D4 = var_20.UnkVCall_0000010Ch
  loc_110B0D66:                                                       var_1F8 = var_20
  loc_110B0D6C:                                                       var_1D8 = var_20.UnkVCall_00000094h
  loc_110B0DB3:                                                       var_200 = var_20
  loc_110B0E0B:                                                       If (var_20.UnkVCall_0000009Ch = 0) = 0 Then
  loc_110B0EEF:                                                         If (frmZGCGToPz.VFG.DispID_0082(arg_C, 14) = global_1100AE28) Then
  loc_110B0F86:                                                           var_8198 = frmZGCGToPz.VFG.DispID_0082(arg_C, 14)
  loc_110B0F99:                                                           var_78 = var_8198
  loc_110B0FDC:                                                           var_90 = "GenLen".00000001h(1, global_110B32D2, global_110B32D2, global_110B32D2, global_110B32D2, var_74, var_8198, var_7C)
  loc_110B1025:                                                           If (var_90 > 10) Then
  loc_110B102F:                                                             var_24 = "票号超长"
  loc_110B1040:                                                           Else
  loc_110B1040:                                                           End If
  loc_110B111E:                                                           If (frmZGCGToPz.VFG.DispID_0082(arg_C, 15) = global_1100AE28) Then
  loc_110B11B5:                                                             var_81A8 = frmZGCGToPz.VFG.DispID_0082(arg_C, 15)
  loc_110B11C8:                                                             var_78 = var_81A8
  loc_110B11F7:                                                             var_90 = "DateCheck".00000001h(1, global_110B32D2, global_110B32D2, global_110B32D2, global_110B32D2, var_74, var_81A8, var_7C)
  loc_110B1247:                                                             If CBool(Not(var_90)) Then
  loc_110B1251:                                                               var_24 = "票号发生日期非法"
  loc_110B1262:                                                             Else
  loc_110B1262:                                                             End If
  loc_110B1262:                                                           End If
  loc_110B1285:                                                           var_1F0 = var_20
  loc_110B12D2:                                                           var_1F8 = var_20
  loc_110B12D8:                                                           var_1D8 = var_20.UnkVCall_0000008Ch
  loc_110B133B:                                                           If (var_20.UnkVCall_000000A4h = 0) = 0 Then
  loc_110B13FA:                                                             If (frmZGCGToPz.VFG.DispID_0082(arg_C, 16) = global_1100AE28) Then
  loc_110B14A4:                                                               var_78 = frmZGCGToPz.VFG.DispID_0082(arg_C, 16)
  loc_110B1524:                                                               var_38 = "BmCheck".00000002h(var_154, 0, var_15C)
  loc_110B1556:                                                               var_81C8 = (var_38 = global_1100AE28)
  loc_110B155E:                                                               If var_81C8 = 0 Then
  loc_110B1568:                                                                 var_24 = "部门非法"
  loc_110B1579:                                                               Else
  loc_110B1596:                                                                 var_118 = arg_C
  loc_110B1620:                                                                 frmZGCGToPz.VFG.DispID_0082(16, var_38)
  loc_110B1655:                                                                 var_1F0 = var_20
  loc_110B1687:                                                                 If var_20.UnkVCall_000000A4h Then
  loc_110B1695:                                                                   var_118 = var_38
  loc_110B16E7:                                                                   var_128 = var_28
  loc_110B1735:                                                                   "BmToProperties".00000002h
  loc_110B1752:                                                                   Set var_28 = {3302AA19-EB96-11D2-AF06000021009B21}()
  loc_110B1780:                                                                   var_1F0 = var_28
  loc_110B1786:                                                                   var_1D4 = var_28.UnkVCall_00000034h
  loc_110B17AC:                                                                   If var_1D4 = 0 Then
  loc_110B17BA:                                                                     var_24 = "部门非末级"
  loc_110B17CB:                                                                   Else
  loc_110B17D3:                                                                     var_24 = "部门为空"
  loc_110B17E4:                                                                   Else
  loc_110B1866:                                                                     frmZGCGToPz.VFG.DispID_0082(var_128, 285257256)
  loc_110B1878:                                                                   End If
  loc_110B1878:                                                                 End If
  loc_110B189B:                                                                 var_1F0 = var_20
  loc_110B18CD:                                                                 If var_20.UnkVCall_0000008Ch Then
  loc_110B1979:                                                                   var_81E0 = (frmZGCGToPz.VFG.DispID_0082(arg_C, &H11) = global_1100AE28)
  loc_110B19B1:                                                                   If var_81E0 Then
  loc_110B1A5D:                                                                     var_81E8 = (frmZGCGToPz.VFG.DispID_0082(arg_C, 16) = global_1100AE28)
  loc_110B1AB9:                                                                     If var_81E8 + 1 Then
  loc_110B1B3C:                                                                       var_78 = frmZGCGToPz.VFG.DispID_0082(arg_C, &H11)
  loc_110B1BCA:                                                                       var_90 = "ZyCheck".00000003h(var_174, "BmCheck".00000002h(var_154, 80020004h, var_15C), var_17C)
  loc_110B1BDF:                                                                       var_34 = var_90
  loc_110B1C11:                                                                       var_81F4 = (var_34 = global_1100AE28)
  loc_110B1C19:                                                                       If var_81F4 = 0 Then
  loc_110B1C23:                                                                         var_24 = "职员非法"
  loc_110B1C34:                                                                       Else
  loc_110B1C51:                                                                         var_118 = arg_C
  loc_110B1CDB:                                                                         frmZGCGToPz.VFG.DispID_0082(&H11, var_34)
  loc_110B1CFA:                                                                         var_118 = var_34
  loc_110B1D47:                                                                         var_128 = var_18
  loc_110B1D95:                                                                         "ZyToProperties".00000002h
  loc_110B1DB2:                                                                         Set var_18 = {3302AA41-EB96-11D2-AF06000021009B21}()
  loc_110B1DC0:                                                                         var_118 = arg_C
  loc_110B1E01:                                                                         var_1F0 = var_18
  loc_110B1EBA:                                                                         frmZGCGToPz.VFG.DispID_0082(var_128, var_18.UnkVCall_0000002Ch)
  loc_110B1EDA:                                                                       Else
  loc_110B1F50:                                                                         var_158 = var_38
  loc_110B1F5D:                                                                         var_78 = frmZGCGToPz.VFG.DispID_0082(8, var_128)
  loc_110B2012:                                                                         var_34 = "ZyCheck".00000003h(var_164, 0, var_16C)
  loc_110B2044:                                                                         var_8208 = (var_34 = global_1100AE28)
  loc_110B204C:                                                                         If var_8208 = 0 Then
  loc_110B2056:                                                                           var_24 = "职员不在指定部门内"
  loc_110B2067:                                                                         Else
  loc_110B20A5:                                                                           var_118 = arg_C
  loc_110B210C:                                                                           frmZGCGToPz.VFG.DispID_0082(&H11, var_34)
  loc_110B211E:                                                                         End If
  loc_110B211E:                                                                       End If
  loc_110B211E:                                                                     End If
  loc_110B2141:                                                                     var_1F0 = var_20
  loc_110B2173:                                                                     If var_20.UnkVCall_00000094h Then
  loc_110B221F:                                                                       var_8214 = (frmZGCGToPz.VFG.DispID_0082(arg_C, &H12) = global_1100AE28)
  loc_110B2230:                                                                       var_1F0 = var_8214
  loc_110B2257:                                                                       If var_1F0 = 0 Then GoTo loc_110B2745
  loc_110B2301:                                                                       var_78 = frmZGCGToPz.VFG.DispID_0082(arg_C, &H12)
  loc_110B2381:                                                                       var_3C = "KhCheck".00000002h(var_154, 0, var_15C)
  loc_110B23B3:                                                                       var_8220 = (var_3C = global_1100AE28)
  loc_110B23BB:                                                                       If var_8220 = 0 Then
  loc_110B23C5:                                                                         var_24 = "客户非法"
  loc_110B23D6:                                                                       Else
  loc_110B2414:                                                                         var_118 = arg_C
  loc_110B247B:                                                                         frmZGCGToPz.VFG.DispID_0082(&H12, var_3C)
  loc_110B248D:                                                                       End If
  loc_110B24B0:                                                                       var_1F0 = var_20
  loc_110B24E2:                                                                       If var_20.UnkVCall_0000009Ch Then
  loc_110B258E:                                                                         var_822C = (frmZGCGToPz.VFG.DispID_0082(arg_C, &H13) = global_1100AE28)
  loc_110B25C6:                                                                         If var_822C Then
  loc_110B2670:                                                                           var_78 = frmZGCGToPz.VFG.DispID_0082(arg_C, &H13)
  loc_110B2722:                                                                           var_8238 = ("GysCheck".00000002h(var_154, 0, var_15C) = global_1100AE28)
  loc_110B272A:                                                                           If var_8238 = 0 Then
  loc_110B2734:                                                                             var_24 = "供应商非法"
  loc_110B2740:                                                                             GoTo loc_110B3293
  loc_110B274D:                                                                             var_24 = "客户为空"
  loc_110B275E:                                                                           Else
  loc_110B277B:                                                                             var_118 = arg_C
  loc_110B27E6:                                                                           Else
  loc_110B27EE:                                                                             var_24 = "供应商为空"
  loc_110B27FF:                                                                           Else
  loc_110B281C:                                                                             var_118 = arg_C
  loc_110B2884:                                                                           End If
  loc_110B28A8:                                                                           frmZGCGToPz.VFG.DispID_0082(var_128, 285257256)
  loc_110B28DD:                                                                           var_1F0 = var_20
  loc_110B292A:                                                                           var_1F8 = var_20
  loc_110B2930:                                                                           var_1D8 = var_20.UnkVCall_0000009Ch
  loc_110B296E:                                                                           If (var_20.UnkVCall_00000094h = 0) = 0 Then
  loc_110B2A1A:                                                                             var_8248 = (frmZGCGToPz.VFG.DispID_0082(arg_C, &H14) = global_1100AE28)
  loc_110B2A52:                                                                             If var_8248 Then
  loc_110B2A75:                                                                               var_118 = arg_C
  loc_110B2AE9:                                                                               var_824C = frmZGCGToPz.VFG.DispID_0082(var_118, var_128)
  loc_110B2AFC:                                                                               var_78 = var_824C
  loc_110B2B3F:                                                                               var_90 = "GenLen".00000001h(global_110B32D2, var_118, var_128, , global_110B32D2, var_74, var_824C, var_7C)
  loc_110B2B88:                                                                               If (var_90 > 20) Then
  loc_110B2B92:                                                                                 var_24 = "业务员超长"
  loc_110B2BA3:                                                                               Else
  loc_110B2BA3:                                                                               End If
  loc_110B2BA3:                                                                             End If
  loc_110B2BC3:                                                                             var_1F0 = var_20
  loc_110B2C1B:                                                                             If (var_20.UnkVCall_000000ACh = global_1100AE28) Then
  loc_110B2CC7:                                                                               var_8260 = (frmZGCGToPz.VFG.DispID_0082(arg_C, &H15) = global_1100AE28)
  loc_110B2CFF:                                                                               If var_8260 Then
  loc_110B2D25:                                                                                 var_1F0 = var_20
  loc_110B2D58:                                                                                 var_8268 = (var_20.UnkVCall_000000ACh = global_1100AE28)
  loc_110B2D7D:                                                                                 If var_8268 Then
  loc_110B2DA3:                                                                                   var_1F0 = var_20
  loc_110B2DD3:                                                                                   var_78 = var_20.UnkVCall_000000ACh
  loc_110B2E81:                                                                                   var_88 = frmZGCGToPz.VFG.DispID_0082(arg_C, &H15)
  loc_110B2F09:                                                                                   var_A0 = "XmCheck".00000003h(var_164, Not(8), var_16C)
  loc_110B2F1E:                                                                                   var_2C = var_A0
  loc_110B2F57:                                                                                   var_8278 = (var_2C = global_1100AE28)
  loc_110B2F5F:                                                                                   If var_8278 = 0 Then
  loc_110B2F69:                                                                                     var_24 = "项目非法"
  loc_110B2F7A:                                                                                   Else
  loc_110B2FA6:                                                                                     var_4C = var_20.UnkVCall_000000ACh
  loc_110B2FD4:                                                                                     var_128 = var_2C
  loc_110B3005:                                                                                     Set var_58 = var_1C
  loc_110B3087:                                                                                     "XmToProperties".00000003h
  loc_110B30A4:                                                                                     Set var_1C = {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_110B30F9:                                                                                     If var_1C.UnkVCall_00000034h Then
  loc_110B3107:                                                                                       var_24 = "项目已结算"
  loc_110B3118:                                                                                     Else
  loc_110B3142:                                                                                       var_118 = %cobj
  loc_110B31BA:                                                                                       frmZGCGToPz.VFG.DispID_0082(&H15, 285257256)
  loc_110B31D7:                                                                                     Else
  loc_110B31DF:                                                                                       var_24 = "项目为空"
  loc_110B31F0:                                                                                     Else
  loc_110B31F8:                                                                                       var_24 = "制单日期非法"
  loc_110B31FE:                                                                                     End If
  loc_110B31FE:                                                                                   End If
  loc_110B3204:                                                                                   GoTo loc_110B3293
  loc_110B320D:                                                                                   If var_4 Then
  loc_110B3218:                                                                                   End If
  loc_110B3292:                                                                                   Exit Sub
  loc_110B3293:                                                                                 End If
  loc_110B3293:                                                                               End If
  loc_110B3293:                                                                             End If
  loc_110B3293:                                                                           End If
  loc_110B3293:                                                                         End If
  loc_110B3293:                                                                       End If
  loc_110B3293:                                                                     End If
  loc_110B3293:                                                                   End If
  loc_110B3293:                                                                 End If
  loc_110B3293:                                                               End If
  loc_110B3293:                                                             End If
  loc_110B3293:                                                           End If
  loc_110B3293:                                                         End If
  loc_110B3293:                                                       End If
  loc_110B3293:                                                     End If
  loc_110B3293:                                                   End If
  loc_110B3293:                                                 End If
  loc_110B3293:                                               End If
  loc_110B3293:                                             End If
  loc_110B3293:                                           End If
  loc_110B3293:                                         End If
  loc_110B3293:                                       End If
  loc_110B3293:                                     End If
  loc_110B3293:                                   End If
  loc_110B3293:                                 End If
  loc_110B3293:                               End If
  loc_110B3293:                             End If
  loc_110B3293:                           End If
  loc_110B3293:                         End If
  loc_110B3293:                       End If
  loc_110B3293:                     End If
  loc_110B3293:                   End If
  loc_110B3293:                 End If
  loc_110B3293:               End If
  loc_110B3293:             End If
  loc_110B3293:           End If
  loc_110B3293:         End If
  loc_110B3293:       End If
  loc_110B3293:     End If
  loc_110B3293:   End If
  loc_110B3293: End If
  loc_110B3293: ' Referenced from: 110B3204
End Sub

Private Sub Proc_15_13_110B3300
  Dim var_9C As Variant
  Dim var_8034 As Label
  Dim var_8074 As Label
  Dim var_A0 As Variant
  Dim var_38 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA47-EB96-11D2-AF06000021009B21}()
  loc_110B345A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110B3460: var_294 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110B3486: Set var_9C = frmZGCGToPz.VFG
  loc_110B34D0: If (CLng(var_9C.DispID_0007) < 2) Then
  loc_110B34FE:   var_800C = = Global.Screen
  loc_110B3520:   var_8010 = ecx
  loc_110B3528:   var_8010 = var_9C.UnkVCall_0000007Ch
  loc_110B3595:   var_C8 = "提示信息"
  loc_110B3597:   var_150 = "没有可生成用友凭证的数据。"
  loc_110B35A6: Else
  loc_110B3656:   var_264 = ("GetAccInfo".00000002h(, , fs:[00000000h], , "GL", var_16C, "dGLStartDate", var_174) = 1100AE28h)
  loc_110B3670:   If var_264 = 0 Then GoTo loc_110B37B1
  loc_110B369E:   var_801C = = Global.Screen
  loc_110B36C0:   var_8020 = ecx
  loc_110B36C8:   var_8020 = var_9C.UnkVCall_0000007Ch
  loc_110B3735:   var_C8 = "提示信息"
  loc_110B3737:   var_150 = "总账系统尚未启用，不能进行凭证引入！"
  loc_110B3741: End If
  loc_110B3773: MsgBox(var_150, 64, var_C8, var_D8, var_E8)
  loc_110B37A0: Exit Sub
  loc_110B37AC: GoTo loc_110BC3B7
  loc_110B37BB: var_8024 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[VouchNum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " & " CREATE TABLE VouchNum(iperiod tinyint NULL ,csign varchar(8) NULL ,ino_id int NULL,constraint index1 unique(iperiod,csign,ino_id))"
  loc_110B37C1: var_B0 = var_8024
  loc_110B3820: var_D8.00000001h(0, , , , "3缷M鋎?", var_AC, var_8024, var_B4)
  loc_110B3840: On Error GoTo 0
  loc_110B3846: var_B0 = %ecx = %S_edx_S
  loc_110B3868: var_78 = "AS13"
  loc_110B3880: var_78)
  loc_110B38AA: If Not (var_78)) Then
  loc_110B38DB:   If Global.Screen < 0 Then
  loc_110B38EC:   End If
  loc_110B38F6:   var_8030 = ecx
  loc_110B3905:   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110B3918:   Else
  loc_110B3929:     call var_8034 = var_9C(var_9C, frmZGCGToPz.Label3, var_9C, global_1100C47C, 0000007Ch)
  loc_110B392B:     var_264 = var_8034
  loc_110B3939:     Label3.Caption = "正在进行数据分析，请稍等..."
  loc_110B3966:     var_150 = True
  loc_110B39A9:     call var_8038 = var_9C(var_9C, frmZGCGToPz.Pic1, global_80010007, 0000000Bh, var_154, True, var_14C)
  loc_110B39AC:     var_8038.DispID_0000 =
  loc_110B39D5:     call var_803C = var_9C(var_9C, frmZGCGToPz.Pic1, global_FFFFFDDA, var_9C = var_9C)
  loc_110B39D8:     var_803C.DispID_0000
  loc_110B39F7:     var_8040 = .Proc_15_11_110ABD80(var_24C)
  loc_110B3A05:     If var_24C = 2 Then
  loc_110B3A0B:       var_150 = %ecx = %S_edx_S
  loc_110B3A4E:       call var_8044 = var_9C(var_9C, frmZGCGToPz.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C)
  loc_110B3A51:       var_8044.DispID_0000 =
  loc_110B3AED:       MsgBox("数据源中没有合法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110B3B2A:       var_24C = %ecx = %S_edx_S
  loc_110B3B50:       "AS13")
  loc_110B3B92:       var_B8 = Global.Screen
  loc_110B3BB4:       var_804C = ecx
  loc_110B3BC3:       If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110B3BD6:       Else
  loc_110B3BD8:         If var_804C = 1 Then
  loc_110B3BDE:           var_150 = %ecx = %S_edx_S
  loc_110B3C21:           call var_8050 = var_9C(var_9C, frmZGCGToPz.Pic1, global_80010007, 0000000Bh, var_154, var_14C = var_9C, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_110B3C24:           var_8050.DispID_0000 =
  loc_110B3CC0:           MsgBox("数据源中含有非法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110B3CFD:           var_24C = %ecx = %S_edx_S
  loc_110B3D23:           "AS13")
  loc_110B3D65:           var_B8 = Global.Screen
  loc_110B3D87:           var_8058 = ecx
  loc_110B3D96:           If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110B3DA9:           Else
  loc_110B3DAB:             If var_8058 = 3 Then
  loc_110B3DB1:               var_150 = %ecx = %S_edx_S
  loc_110B3DF4:               call var_805C = var_9C(var_9C, frmZGCGToPz.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_110B3DF7:               var_805C.DispID_0000 =
  loc_110B3E93:               MsgBox("数据源中指定的凭证号无效或重号，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110B3ED0:               var_24C = %ecx = %S_edx_S
  loc_110B3EF6:               "AS13")
  loc_110B3F38:               var_B8 = Global.Screen
  loc_110B3F5A:               var_8064 = ecx
  loc_110B3F69:               If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110B3F7C:               Else
  loc_110B3FBE:                 var_C8 = "提示信息"
  loc_110B3FE4:                 var_B8 = "数据源中的数据已全部通过检查，是否开始引入？"
  loc_110B4008:                 MsgBox(var_B8, 36, var_C8, var_D8, var_E8)
  loc_110B404D:                 If (MsgBox(var_B8, 36, var_C8, var_D8, var_E8) = 7) Then
  loc_110B4098:                   call var_8068 = var_9C(var_9C, frmZGCGToPz.Pic1, global_80010007, 0000000Bh, var_154, frmZGCGToPz.Pic1, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_110B409B:                   var_8068.DispID_0000 =
  loc_110B40C1:                   var_24C = %ecx = %S_edx_S
  loc_110B40E7:                   "AS13")
  loc_110B4129:                   var_B8 = Global.Screen
  loc_110B414B:                   var_8070 = ecx
  loc_110B415A:                   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110B416D:                   Else
  loc_110B416E:                     On Error GoTo 0
  loc_110B4185:                     call var_8074 = var_9C(var_9C, frmZGCGToPz.Label3, var_9C = var_9C, var_9C, global_1100C47C, 0000007Ch)
  loc_110B4187:                     var_264 = var_8074
  loc_110B4195:                     Label3.Caption = "正在写数据，请稍等..."
  loc_110B41D9:                     call var_8078 = var_9C(var_9C, frmZGCGToPz.Pic1, global_FFFFFDDA, 00000000h)
  loc_110B41DC:                     var_8078.DispID_0000
  loc_110B4213:                     Set var_74 = CreateObject("UfDbKit.UfRecordset", 0)
  loc_110B422A:                     var_150 = "SELECT TOP 1 * FROM GL_accvouch"
  loc_110B429F:                     Set var_74 = "DataMdb".00000000h.00000001h(var_14C, "SELECT TOP 1 * FROM GL_accvouch", var_154)
  loc_110B42D3:                     call var_8084 = var_9C(var_9C, frmZGCGToPz.VFG, 00000007h, 00000000h)
  loc_110B4337:                     If var_24 <= CLng(var_8084.DispID_0000)(-1) Then
  loc_110B4341:                       var_2A8 = var_24
  loc_110B4347:                       var_150 = var_24
  loc_110B43C4:                       call var_8090 = var_9C(var_9C, frmZGCGToPz.VFG, 00000082h, 00000002h, 3, var_174, 2, var_16C, 00000003h, var_154, var_24, var_14C)
  loc_110B43DE:                       var_C0 = var_8090.DispID_0000
  loc_110B43FC:                       var_D8)
  loc_110B4454:                       var_70 = CByte("DateToPeriod".00000001h(8, var_D4))
  loc_110B448D:                       var_150 = var_2A8
  loc_110B4506:                       call var_809C = var_9C(var_9C, frmZGCGToPz.VFG, 00000082h, 00000002h, 3, var_174, 3, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110B4525:                       var_58 = var_809C.DispID_0000
  loc_110B4549:                       var_150 = var_2A8
  loc_110B45C6:                       call var_80A4 = var_9C(var_9C, frmZGCGToPz.VFG, 00000082h, 00000002h, 3, var_174, 0, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110B45E5:                       var_64 = var_80A4.DispID_0000
  loc_110B4609:                       var_150 = var_2A8
  loc_110B4686:                       call var_80AC = var_9C(var_9C, frmZGCGToPz.VFG, 00000082h, 00000002h, 3, var_174, 1, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110B46EF:                       If (var_80AC.DispID_0000 = global_1100D76C) Then
  loc_110B4706:                         call var_80B8 = var_9C(var_A8, frmZGCGToPz.Label3)
  loc_110B4708:                         var_264 = var_80B8
  loc_110B4818:                         var_80 = "正在处理：第[" & frmZGCGToPz.VFG.DispID_0082(var_2A8, 2) & " - "
  loc_110B4959:                         var_D8 = frmZGCGToPz.VFG.DispID_0082(var_2A8, 0)
  loc_110B49A0:                         var_98 = var_80 & frmZGCGToPz.VFG.DispID_0082(var_2A8, 3) & " - " & var_D8 & "]号凭证"
  loc_110B49B0:                         var_98 = var_80B8.UnkVCall_00000054h
  loc_110B4A6B:                         frmZGCGToPz.Pic1.DispID_FFFFFDDA
  loc_110B4A9F:                         var_3C = var_24
  loc_110B4AB3:                         Set var_9C = frmZGCGToPz.Chk
  loc_110B4AB5:                         var_264 = var_9C
  loc_110B4AC7:                         Set var_A0 = var_9C(0)
  loc_110B4AEB:                         var_26C = var_A0
  loc_110B4B55:                         If (var_A0.Value = 1) Then
  loc_110B4B88:                           var_24C = CInt("cIYear".00000000h)
  loc_110B4B9D:                           var_24C, var_70)
  loc_110B4BAA:                           var_54 = var_24C, var_70)
  loc_110B4BBB:                         Else
  loc_110B4BD1:                           var_80E8 = .Proc_15_15_110BDEB0(var_70)
  loc_110B4BE3:                           var_54 = var_258
  loc_110B4BE6:                         End If
  loc_110B4BEB:                         If var_54 > 0 Then
  loc_110B4BF3:                           On Error GoTo loc_110BA55D
  loc_110B4C2C:                           "wksAlias".00000000h.00000000h(var_58)
  loc_110B4C4B:                           var_1A0 = var_70
  loc_110B4D14:                           var_D8)
  loc_110B4DC0:                           var_80FC = (var_58 = frmZGCGToPz.VFG.DispID_0082(var_24, 3))
  loc_110B4DCD:                           var_1F0 = var_80FC + 1
  loc_110B4E89:                           var_8104 = (var_64 = frmZGCGToPz.VFG.DispID_0082(var_24, 0))
  loc_110B4E96:                           var_240 = var_8104 + 1
  loc_110B4F2C:                           var_8110 = (frmZGCGToPz.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_80FC + 1 And var_8104 + 1
  loc_110B4FB8:                           If CBool(var_8110) Then
  loc_110B5059:                             var_C0 = frmZGCGToPz.VFG.DispID_0082(var_24, 6)
  loc_110B5096:                             var_1A0 = var_38
  loc_110B5104:                             "kmCodeToProperties".00000002h
  loc_110B5124:                             Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110B515D:                             var_74.AddNew
  loc_110B5168:                             var_150 = "ibook"
  loc_110B51D9:                             var_74.DispID_0000(0)
  loc_110B51DB:                             var_1A0 = "iPeriod"
  loc_110B528A:                             var_C0 = frmZGCGToPz.VFG.DispID_0082(var_24, 2)
  loc_110B52A8:                             var_D8)
  loc_110B5341:                             var_74.DispID_0000("DateToPeriod".00000001h)
  loc_110B5376:                             var_190 = "csign"
  loc_110B5483:                             var_74.DispID_0000(frmZGCGToPz.VFG.DispID_0082(var_24, 3))
  loc_110B54AA:                             var_190 = "isignseq"
  loc_110B55CA:                             var_74.DispID_0000(Proc_0_4_11026BD0(frmZGCGToPz.VFG.DispID_0082(var_24, 3), var_64, var_258))
  loc_110B55F5:                             var_150 = "ino_id"
  loc_110B5667:                             var_74.DispID_0000(var_54)
  loc_110B5669:                             var_190 = "dbill_date"
  loc_110B5718:                             var_C0 = frmZGCGToPz.VFG.DispID_0082(var_24, 2)
  loc_110B5736:                             var_D8)
  loc_110B5793:                             var_74.DispID_0000(var_D8)
  loc_110B57C1:                             var_190 = "idoc"
  loc_110B57D9:                             var_150 = var_24
  loc_110B58E2:                             var_74.DispID_0000(Val(frmZGCGToPz.VFG.DispID_0082(var_150, 4)))
  loc_110B590D:                             var_160 = "ctext1"
  loc_110B5974:                             var_74.DispID_0000(var_150)
  loc_110B597B:                             var_160 = "ctext2"
  loc_110B59E2:                             var_74.DispID_0000(var_150)
  loc_110B59E9:                             var_150 = "cbill"
  loc_110B5A57:                             var_74.DispID_0000("cUserName".00000000h(, var_14C, "cbill", var_154))
  loc_110B5A6D:                             var_160 = "cbook"
  loc_110B5AD4:                             var_74.DispID_0000(var_150)
  loc_110B5ADB:                             var_160 = "ccheck"
  loc_110B5B42:                             var_74.DispID_0000(var_150)
  loc_110B5B49:                             var_160 = "ccashier"
  loc_110B5BB0:                             var_74.DispID_0000(var_150)
  loc_110B5BB7:                             var_160 = "iflag"
  loc_110B5C1E:                             var_74.DispID_0000(var_150)
  loc_110B5C25:                             var_160 = "coutaccset"
  loc_110B5C8C:                             var_74.DispID_0000(var_150)
  loc_110B5C93:                             var_160 = "ioutyear"
  loc_110B5CFA:                             var_74.DispID_0000(var_150)
  loc_110B5D01:                             var_160 = "coutsysver"
  loc_110B5D68:                             var_74.DispID_0000(var_150)
  loc_110B5D6F:                             var_160 = "coutsysname"
  loc_110B5DD6:                             var_74.DispID_0000(var_150)
  loc_110B5DDD:                             var_170 = "ioutperiod"
  loc_110B5E7A:                             var_74.DispID_0000(var_74.DispID_0000("iPeriod"))
  loc_110B5E8B:                             var_170 = "doutbilldate"
  loc_110B5F4E:                             var_74.DispID_0000(CStr(var_74.DispID_0000("dbill_date")))
  loc_110B5F6B:                             var_150 = "iYear"
  loc_110B5FD9:                             var_74.DispID_0000("cIYear".00000000h(var_58, var_14C, "iYear", var_154))
  loc_110B60D7:                             var_74.DispID_0000("cIYear".00000000h(, var_16C, "iYPeriod", var_174) & Format(var_70, "00"))
  loc_110B6105:                             var_160 = "coutsign"
  loc_110B616C:                             var_74.DispID_0000(var_70)
  loc_110B616E:                             var_190 = "coutno_id"
  loc_110B627B:                             var_74.DispID_0000(frmZGCGToPz.VFG.DispID_0082(var_24, 3))
  loc_110B62A7:                             var_150 = "bvouchedit"
  loc_110B6316:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110B631D:                             var_150 = "bvouchaddordele"
  loc_110B638E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110B6395:                             var_150 = "bvouchmoneyhold"
  loc_110B6406:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110B640D:                             var_150 = "bvalueedit"
  loc_110B647E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110B6485:                             var_150 = "bcodeedit"
  loc_110B64F6:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110B64FD:                             var_150 = "bPCSedit"
  loc_110B656E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110B6575:                             var_150 = "bDeptedit"
  loc_110B65E6:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110B65ED:                             var_150 = "bItemedit"
  loc_110B665E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110B6665:                             var_150 = "inid"
  loc_110B66D7:                             var_74.DispID_0000(1)
  loc_110B66D9:                             var_190 = "cdigest"
  loc_110B67EA:                             var_74.DispID_0000(frmZGCGToPz.VFG.DispID_0082(var_24, 5))
  loc_110B6811:                             var_190 = "cCode"
  loc_110B6920:                             var_74.DispID_0000(frmZGCGToPz.VFG.DispID_0082(var_24, 6))
  loc_110B69C8:                             var_7C = var_38.UnkVCall_0000006Ch
  loc_110B6A13:                             var_8150 = (var_38.UnkVCall_0000006Ch = global_1100AE28)
  loc_110B6A20:                             var_160 = var_8150 + 1
  loc_110B6AAB:                             var_74.DispID_0000(IIf(var_8150 + 1, vbNull, 0))
  loc_110B6B90:                             var_1B0 = "md"
  loc_110B6BD9:                             var_BC = var_25C
  loc_110B6C60:                             var_74.DispID_0000(Format(Val(frmZGCGToPz.VFG.DispID_0082(var_24, 7)), "#.00"))
  loc_110B6D51:                             var_1B0 = "mc"
  loc_110B6D9A:                             var_BC = var_25C
  loc_110B6E21:                             var_74.DispID_0000(Format(Val(frmZGCGToPz.VFG.DispID_0082(var_24, 8)), "#.00"))
  loc_110B6EE9:                             If (var_74.DispID_0000("md") <> 0) Then
  loc_110B6F5E:                               If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110B6F69:                                 var_150 = "md_f"
  loc_110B6FDA:                                 var_74.DispID_0000(0)
  loc_110B6FE4:                               Else
  loc_110B7097:                                 var_1B0 = "md_f"
  loc_110B70E0:                                 var_BC = var_25C
  loc_110B7167:                                 var_74.DispID_0000(Format(Val(frmZGCGToPz.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_110B71A8:                               End If
  loc_110B721A:                               If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_110B7225:                                 var_150 = "nd_s"
  loc_110B7296:                                 var_74.DispID_0000(0)
  loc_110B72A0:                               Else
  loc_110B72AF:                               Else
  loc_110B731E:                                 If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110B7329:                                   var_150 = "mc_f"
  loc_110B739A:                                   var_74.DispID_0000(0)
  loc_110B73A4:                                 Else
  loc_110B7457:                                   var_1B0 = "mc_f"
  loc_110B74A0:                                   var_BC = var_25C
  loc_110B7527:                                   var_74.DispID_0000(Format(Val(frmZGCGToPz.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_110B7568:                                 End If
  loc_110B75DA:                                 If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_110B75E1:                                   GoTo loc_110B7225
  loc_110B75E6:                                 End If
  loc_110B75F0:                               End If
  loc_110B770A:                               var_74.DispID_0000(Val(frmZGCGToPz.VFG.DispID_0082(var_24, 9)))
  loc_110B7730:                             End If
  loc_110B77A2:                             If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110B77AD:                               var_150 = "nfrat"
  loc_110B781E:                               var_74.DispID_0000(0)
  loc_110B7828:                             Else
  loc_110B794C:                               var_74.DispID_0000(Val(frmZGCGToPz.VFG.DispID_0082(var_24, 11)))
  loc_110B7972:                             End If
  loc_110B79C7:                             If var_38.UnkVCall_0000010Ch Then
  loc_110B7A5E:                               var_1F0 = "csettle"
  loc_110B7B45:                               var_81A4 = (frmZGCGToPz.VFG.DispID_0082(var_24, 13) = global_1100AE28)
  loc_110B7B52:                               var_1E0 = var_81A4 + 1
  loc_110B7BDD:                               var_74.DispID_0000(IIf(var_81A4 + 1, vbNull, frmZGCGToPz.VFG.DispID_0082(var_24, 13)))
  loc_110B7C36:                             End If
  loc_110B7C5F:                             var_24C = var_38.UnkVCall_0000010Ch
  loc_110B7CAC:                             var_250 = var_38.UnkVCall_00000094h
  loc_110B7D4B:                             If (var_38.UnkVCall_0000009Ch = 0) = 0 Then
  loc_110B7DE2:                               var_1F0 = "cn_id"
  loc_110B7E91:                               var_E0 = frmZGCGToPz.VFG.DispID_0082(var_24, 14)
  loc_110B7EC9:                               var_81BC = (frmZGCGToPz.VFG.DispID_0082(var_24, 14) = global_1100AE28)
  loc_110B7ED6:                               var_1E0 = var_81BC + 1
  loc_110B7F61:                               var_74.DispID_0000(IIf(var_81BC + 1, vbNull, var_E0))
  loc_110B8048:                               var_1F0 = "dt_date"
  loc_110B80F7:                               var_D0 = frmZGCGToPz.VFG.DispID_0082(var_24, 15)
  loc_110B8115:                               var_E0)
  loc_110B8142:                               var_81C8 = (frmZGCGToPz.VFG.DispID_0082(var_24, 15) = global_1100AE28)
  loc_110B814F:                               var_1E0 = var_81C8 + 1
  loc_110B81DA:                               var_74.DispID_0000(IIf(var_81C8 + 1, vbNull, var_E0))
  loc_110B82C8:                               var_1F0 = "cname"
  loc_110B83AF:                               var_81D4 = (frmZGCGToPz.VFG.DispID_0082(var_24, &H14) = global_1100AE28)
  loc_110B83BC:                               var_1E0 = var_81D4 + 1
  loc_110B8447:                               var_74.DispID_0000(IIf(var_81D4 + 1, vbNull, frmZGCGToPz.VFG.DispID_0082(var_24, &H14)))
  loc_110B84A0:                             End If
  loc_110B8516:                             var_250 = var_38.UnkVCall_0000008Ch
  loc_110B8554:                             If (var_38.UnkVCall_000000A4h = 0) = 0 Then
  loc_110B855E:                               var_150 = var_24
  loc_110B85EB:                               var_1F0 = "cdept_id"
  loc_110B86D2:                               var_81E8 = (frmZGCGToPz.VFG.DispID_0082(var_150, 16) = global_1100AE28)
  loc_110B86DF:                               var_1E0 = var_81E8 + 1
  loc_110B876A:                               var_74.DispID_0000(IIf(var_81E8 + 1, vbNull, frmZGCGToPz.VFG.DispID_0082(var_24, 16)))
  loc_110B87C5:                             Else
  loc_110B87CA:                               var_160 = "cdept_id"
  loc_110B8831:                               var_74.DispID_0000(var_150)
  loc_110B8836:                             End If
  loc_110B888B:                             If var_38.UnkVCall_0000008Ch Then
  loc_110B8895:                               var_150 = var_24
  loc_110B8922:                               var_1F0 = "cperson_id"
  loc_110B8A09:                               var_81F8 = (frmZGCGToPz.VFG.DispID_0082(var_150, &H11) = global_1100AE28)
  loc_110B8A16:                               var_1E0 = var_81F8 + 1
  loc_110B8AA1:                               var_74.DispID_0000(IIf(var_81F8 + 1, vbNull, frmZGCGToPz.VFG.DispID_0082(var_24, &H11)))
  loc_110B8AFC:                             Else
  loc_110B8B01:                               var_160 = "cperson_id"
  loc_110B8B68:                               var_74.DispID_0000(var_150)
  loc_110B8B6D:                             End If
  loc_110B8BC2:                             If var_38.UnkVCall_00000094h Then
  loc_110B8BCC:                               var_150 = var_24
  loc_110B8C59:                               var_1F0 = "ccus_id"
  loc_110B8D40:                               var_8208 = (frmZGCGToPz.VFG.DispID_0082(var_150, &H12) = global_1100AE28)
  loc_110B8D4D:                               var_1E0 = var_8208 + 1
  loc_110B8DD8:                               var_74.DispID_0000(IIf(var_8208 + 1, vbNull, frmZGCGToPz.VFG.DispID_0082(var_24, &H12)))
  loc_110B8E33:                             Else
  loc_110B8E38:                               var_160 = "ccus_id"
  loc_110B8E9F:                               var_74.DispID_0000(var_150)
  loc_110B8EA4:                             End If
  loc_110B8EF9:                             If var_38.UnkVCall_0000009Ch Then
  loc_110B8F03:                               var_150 = var_24
  loc_110B8F90:                               var_1F0 = "csup_id"
  loc_110B9077:                               var_8218 = (frmZGCGToPz.VFG.DispID_0082(var_150, &H13) = global_1100AE28)
  loc_110B9084:                               var_1E0 = var_8218 + 1
  loc_110B910F:                               var_74.DispID_0000(IIf(var_8218 + 1, vbNull, frmZGCGToPz.VFG.DispID_0082(var_24, &H13)))
  loc_110B916A:                             Else
  loc_110B916F:                               var_160 = "csup_id"
  loc_110B91D6:                               var_74.DispID_0000(var_150)
  loc_110B91DB:                             End If
  loc_110B9254:                             If (var_38.UnkVCall_000000ACh = global_1100AE28) Then
  loc_110B925E:                               var_150 = var_24
  loc_110B92EB:                               var_1F0 = "citem_id"
  loc_110B93D2:                               var_822C = (frmZGCGToPz.VFG.DispID_0082(var_150, &H15) = global_1100AE28)
  loc_110B93DF:                               var_1E0 = var_822C + 1
  loc_110B946A:                               var_74.DispID_0000(IIf(var_822C + 1, vbNull, frmZGCGToPz.VFG.DispID_0082(var_24, &H15)))
  loc_110B9547:                               var_7C = var_38.UnkVCall_000000ACh
  loc_110B9598:                               var_8238 = (var_38.UnkVCall_000000ACh = global_1100AE28)
  loc_110B95A5:                               var_160 = var_8238 + 1
  loc_110B9630:                               var_74.DispID_0000(IIf(var_8238 + 1, vbNull, 0))
  loc_110B966A:                             Else
  loc_110B966F:                               var_160 = "citem_id"
  loc_110B96D6:                               var_74.DispID_0000(var_150)
  loc_110B96DD:                               var_160 = "citem_class"
  loc_110B9744:                               var_74.DispID_0000(var_150)
  loc_110B9749:                             End If
  loc_110B974E:                             var_160 = "ccode_equal"
  loc_110B97B5:                             var_74.DispID_0000(var_150)
  loc_110B97BC:                             var_160 = "iflagbank"
  loc_110B9823:                             var_74.DispID_0000(var_150)
  loc_110B982A:                             var_160 = "iflagperson"
  loc_110B9891:                             var_74.DispID_0000(var_150)
  loc_110B989E:                             var_74.Update
  loc_110B98B5:                             var_24 = var_24(1)
  loc_110B98C6:                             var_68 = var_68(1)
  loc_110B98FB:                             var_823C = CLng(frmZGCGToPz.VFG.DispID_0007)
  loc_110B9917:                             var_264 = (var_24(1) > 0)
  loc_110B993E:                             If var_264 = 0 Then GoTo loc_110B4C48
  loc_110B9944:                           End If
  loc_110B9977:                           "wksAlias".00000000h.00000000h
  loc_110B99A4:                           Set var_9C = frmZGCGToPz.Chk
  loc_110B99A6:                           var_264 = var_9C
  loc_110B99B8:                           Set var_A0 = var_9C(0)
  loc_110B99DC:                           var_26C = var_A0
  loc_110B9A46:                           If (var_A0.Value = 1) Then
  loc_110B9A54:                             var_70, var_58)
  loc_110B9A59:                           End If
  loc_110B9A5B:                           On Error GoTo 0
  loc_110B9A92:                           var_250 = CInt("cIYear".00000000h)
  loc_110B9ABC:                           var_24C, var_250, var_70, var_58)
  loc_110B9AC6:                           var_5C = var_24C, var_250, var_70, var_58)
  loc_110B9B09:                           var_250 = CInt("cIYear".00000000h)
  loc_110B9B3D:                           var_48 = r_250, var_70, var_58) var_250, var_70, var_58)
  loc_110B9B4F:                           var_150 = "select * from GL_accvouch where ibook=0 and iYear="
  loc_110B9B77:                           var_170 = var_70
  loc_110B9B9B:                           var_824C = Proc_0_4_11026BD0(var_58, var_54, var_54)
  loc_110B9BA0:                           var_190 = var_824C
  loc_110B9BC8:                           var_1B0 = var_54
  loc_110B9C21:                           var_D8 = 1 & "cIYear".00000000h(, 1, 1) & " and iperiod="
  loc_110B9C8A:                           var_128 = var_D8 & var_70 & " and isignseq=" & var_824C & " and ino_id=" & var_54
  loc_110B9CF3:                           Set var_74 = "DataMdb".00000000h.00000001h
  loc_110B9D92:                           If CBool(Not(var_74.EOF)) Then
  loc_110B9DEA:                             If CBool(Not(var_74.EOF)) Then
  loc_110B9DF3:                               var_170 = var_70
  loc_110B9E08:                               var_150 = "iPeriod"
  loc_110B9E2C:                               var_180 = "csign"
  loc_110B9E40:                               var_1D0 = var_54
  loc_110B9E51:                               var_1B0 = "ino_id"
  loc_110B9FA8:                               If CBool((var_70 = var_14C) And (var_58 = var_D8) And (var_54 = var_1AC)) Then
  loc_110B9FB3:                                 var_150 = "mc"
  loc_110BA035:                                 var_180 = "ccode_equal"
  loc_110BA049:                                 If (var_14C <> 0) Then
  loc_110BA075:                                   var_8278 = (var_5C = global_1100AE28)
  loc_110BA082:                                   var_160 = var_8278 + 1
  loc_110BA0AF:                                   var_C8 = IIf(var_8278 + 1, vbNull, var_5C)
  loc_110BA129:                                 Else
  loc_110BA14F:                                   var_827C = (var_48 = global_1100AE28)
  loc_110BA15C:                                   var_160 = var_827C + 1
  loc_110BA189:                                   var_C8 = IIf(var_827C + 1, vbNull, var_48)
  loc_110BA1FE:                                 End If
  loc_110BA214:                                 var_74.Update
  loc_110BA25E:                                 var_180 = var_38
  loc_110BA2A5:                                 var_B8 = var_74.DispID_0000("cCode")
  loc_110BA302:                                 "kmCodeToProperties".00000002h
  loc_110BA322:                                 Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110BA341:                                 var_150 = "citem_class"
  loc_110BA3A8:                                 If IsNull(var_74.DispID_0000(var_150)) Then
  loc_110BA3BD:                                 Else
  loc_110BA3FE:                                   var_180 = var_28
  loc_110BA445:                                   var_B8 = var_74.DispID_0000(var_150)
  loc_110BA4A2:                                   "XmClassIDToProperties".00000002h
  loc_110BA502:                                   var_78 = {3302AA47-EB96-11D2-AF06000021009B21}().UnkVCall_0000002Ch
  loc_110BA533:                                 End If
  loc_110BA541:                                 var_68 = var_68(1)
  loc_110BA54F:                                 var_74.MoveNext
  loc_110BA558:                                 GoTo loc_110B9D9F
  loc_110BA590:                                 "wksAlias".00000000h.00000000h
  loc_110BA5A8:                                 var_30 = var_3C
  loc_110BA5BD:                                 var_1A0 = var_70
  loc_110BA686:                                 var_D8)
  loc_110BA732:                                 var_829C = (var_58 = frmZGCGToPz.VFG.DispID_0082(var_30, 3))
  loc_110BA73F:                                 var_1F0 = var_829C + 1
  loc_110BA7FB:                                 var_82A4 = (var_64 = frmZGCGToPz.VFG.DispID_0082(var_30, 0))
  loc_110BA808:                                 var_240 = var_82A4 + 1
  loc_110BA89E:                                 var_82B0 = (frmZGCGToPz.VFG.DispID_0082(var_30, 2) = "DateToPeriod".00000001h) And var_829C + 1 And var_82A4 + 1
  loc_110BA92A:                                 If CBool(var_82B0) Then
  loc_110BA934:                                   var_150 = var_30
  loc_110BA9F0:                                   frmZGCGToPz.VFG.DispID_0082(1, "-")
  loc_110BAB70:                                   frmZGCGToPz.VFG.DispID_009E(var_30, 1, var_30, 1, &HFF)
  loc_110BAB85:                                   var_150 = var_30
  loc_110BAC41:                                   frmZGCGToPz.VFG.DispID_0082(&H16, "数据提交错或该数据已经被导入----未引入")
  loc_110BAC60:                                   var_30 = var_30(1)
  loc_110BAC8C:                                   var_82B8 = CLng(frmZGCGToPz.VFG.DispID_0007)
  loc_110BACA8:                                   var_264 = (var_30 > 0)
  loc_110BACCF:                                   If var_264 = 0 Then GoTo loc_110BA5BA
  loc_110BACD5:                                 End If
  loc_110BACD8:                                 var_24 = var_30
  loc_110BACEC:                                 Set var_9C = frmZGCGToPz.Chk
  loc_110BACEE:                                 var_264 = var_9C
  loc_110BAD00:                                 Set var_A0 = var_9C(0)
  loc_110BAD24:                                 var_26C = var_A0
  loc_110BAD8E:                                 If (var_A0.Value = 1) Then
  loc_110BAE8A:                                   "unLockVouch".00000004h(var_180, var_BC, var_C4, 0, var_74, var_70, var_58, var_16C, var_54, &H4002, var_184)
  loc_110BAE93:                                 End If
  loc_110BAE98:                                 var_150 = "VouchNum"
  loc_110BAF0D:                                 Set var_34 = "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C4, 0, var_14C, "VouchNum", var_154)
  loc_110BAF2E:                                 var_150 = "delete from vouchnum"
  loc_110BAF8C:                                 "DataMdb".00000000h.00000001h(1, 1, var_180, var_BC, var_C4, var_14C, "delete from vouchnum", var_154)
  loc_110BAFE9:                                 frmZGCGToPz.Pic1.DispID_80010007 = var_150
  loc_110BAFFD:                                 var_82C4 = Resume(0)
  loc_110BB003:                               End If
  loc_110BB003:                             End If
  loc_110BB003:                           End If
  loc_110BB021:                           var_24 = var_27C+(var_24 - 1)
  loc_110BB024:                           GoTo loc_110B432C
  loc_110BB029:                         End If
  loc_110BB02C:                         var_1A0 = var_70
  loc_110BB0F5:                         var_D8)
  loc_110BB1A1:                         var_82D0 = (var_58 = frmZGCGToPz.VFG.DispID_0082(var_24, 3))
  loc_110BB1AE:                         var_1F0 = var_82D0 + 1
  loc_110BB26A:                         var_82D8 = (var_64 = frmZGCGToPz.VFG.DispID_0082(var_24, 0))
  loc_110BB277:                         var_240 = var_82D8 + 1
  loc_110BB30D:                         var_82E4 = (frmZGCGToPz.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_82D0 + 1 And var_82D8 + 1
  loc_110BB31A:                         var_264 = CBool(var_82E4)
  loc_110BB399:                         If var_264 = 0 Then GoTo loc_110BB003
  loc_110BB3B0:                         Set var_9C = frmZGCGToPz.Chk
  loc_110BB3B2:                         var_264 = var_9C
  loc_110BB3C4:                         Set var_A0 = var_9C(0)
  loc_110BB3E8:                         var_26C = var_A0
  loc_110BB42B:                         var_274 = (var_A0.Value = 1)
  loc_110BB456:                         var_150 = var_24
  loc_110BB477:                         var_190 = "网络共享冲突----未引入"
  loc_110BB481:                         If var_274 = 0 Then
  loc_110BB483:                           var_190 = "指定的凭证号无效或重号----未引入"
  loc_110BB48D:                         End If
  loc_110BB51E:                         frmZGCGToPz.VFG.DispID_0082(var_170, var_190)
  loc_110BB53D:                         var_24 = var_24(1)
  loc_110BB543:                         var_2A8 = var_24(1)
  loc_110BB572:                         var_82EC = CLng(frmZGCGToPz.VFG.DispID_0007)
  loc_110BB58E:                         var_264 = (var_2A8 > 0)
  loc_110BB5B5:                         If var_264 = 0 Then GoTo loc_110BB029
  loc_110BB5BB:                         GoTo loc_110BB003
  loc_110BB5C0:                       End If
  loc_110BB5C3:                       var_1A0 = var_70
  loc_110BB68E:                       var_D8)
  loc_110BB73C:                       var_82F8 = (var_58 = frmZGCGToPz.VFG.DispID_0082(var_2A8, 3))
  loc_110BB749:                       var_1F0 = var_82F8 + 1
  loc_110BB807:                       var_8300 = (var_64 = frmZGCGToPz.VFG.DispID_0082(var_2A8, 0))
  loc_110BB814:                       var_240 = var_8300 + 1
  loc_110BB8AA:                       var_830C = (frmZGCGToPz.VFG.DispID_0082(var_2A8, 2) = "DateToPeriod".00000001h) And var_82F8 + 1 And var_8300 + 1
  loc_110BB8B7:                       var_264 = CBool(var_830C)
  loc_110BB936:                       If var_264 = 0 Then GoTo loc_110BB003
  loc_110BBA2D:                       If (frmZGCGToPz.VFG.DispID_0082(var_2A8, &H16) = global_1100AE28) + 1 Then
  loc_110BBA33:                         var_150 = var_2A8
  loc_110BBAEC:                         Set var_9C = frmZGCGToPz.VFG
  loc_110BBAEF:                         var_9C.DispID_0082(&H16, "凭证借贷不平衡或某分录有错误----未引入")
  loc_110BBB00:                         GoTo loc_110BB5C0
  loc_110BBB05:                       End If
  loc_110BBBCF:                       var_C0 = frmZGCGToPz.VFG.DispID_0082(frmZGCGToPz.VFG, &H16) & "----未引入"
  loc_110BBC6C:                       frmZGCGToPz.VFG.DispID_0082(&H16, var_C0)
  loc_110BBCA9:                       GoTo loc_110BB5C0
  loc_110BBCAE:                     End If
  loc_110BBCF6:                     frmZGCGToPz.Pic1.DispID_80010007 = var_150
  loc_110BBD0D:                     If var_2C Then
  loc_110BBD1D:                       var_24C = frmZGCGToPz.UpdateBTData
  loc_110BBDC5:                       MsgBox("数据引入已完成，数据已生成用友凭证。", 64, "提示信息", 10, 10)
  loc_110BBE37:                       frmZGCGToPz.VFG.DispID_0007 = 1
  loc_110BBE91:                       frmZGCGToPz.VFG.DispID_0007 = 1
  loc_110BBF2C:                       frmZGCGToPz.sBar.DispID_6803001E("数量合计：0.0000")
  loc_110BBFC3:                       frmZGCGToPz.sBar.DispID_6803001E("金额合计：0.00")
  loc_110BC05A:                       Set var_9C = frmZGCGToPz.sBar
  loc_110BC05D:                       var_9C.DispID_6803001E("税额合计：0.00")
  loc_110BC073:                     Else
  loc_110BC0FA:                       MsgBox("数据没有被引入，原因请查看最后一列中的说明。", 64, "提示信息", 10, 10)
  loc_110BC127:                     End If
  loc_110BC12C:                     var_150 = "VouchNum"
  loc_110BC1A5:                     Set var_34 = "DataMdb".00000000h.00000001h(var_180, var_BC, var_C0, var_C4, var_C8, var_14C, "VouchNum", var_154)
  loc_110BC1C6:                     var_150 = "delete  from vouchnum"
  loc_110BC216:                     "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C0, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_110BC26B:                     "AS13")
  loc_110BC2A4:                     var_B8 = Global.Screen
  loc_110BC2C6:                     var_8330 = ecx
  loc_110BC2D5:                     If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110BC2DF:                     End If
  loc_110BC2DF:                   End If
  loc_110BC2DF:                 End If
  loc_110BC2DF:               End If
  loc_110BC2DF:             End If
  loc_110BC2E0:             var_8330 = CheckObj(var_9C, global_1100C47C, 124)
  loc_110BC2E6:           End If
  loc_110BC2E6:         End If
  loc_110BC2E6:       End If
  loc_110BC2E6:     End If
  loc_110BC2E6:   End If
  loc_110BC2E6: End If
  loc_110BC2F2: Exit Sub
  loc_110BC2FE: GoTo loc_110BC3B7
  loc_110BC3B6: Exit Sub
  loc_110BC3B7: ' Referenced from: 110B37AC
  loc_110BC3B7: ' Referenced from: 110BC2FE
End Sub

Private Sub Proc_15_14_110BD0A0
  Dim var_58 As Variant
  Dim var_5C As Variant
  Dim var_64 As frmZGCGToPz.Label3
  Dim var_1D0 As Label
  loc_110BD18D: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110BD196: var_1F0 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110BD1B3: Set var_58 = frmZGCGToPz.Chk
  loc_110BD1BD: var_1D0 = var_58
  loc_110BD1C3: Set var_5C = var_58(0)
  loc_110BD1EE: var_1D8 = var_5C
  loc_110BD231: var_1E0 = (var_5C.Value = 1)
  loc_110BD247: If var_1E0 = 0 Then
  loc_110BD2AC:   If var_14 <= CLng(frmZGCGToPz.VFG.DispID_0007)(-1) Then
  loc_110BD321:     var_7C = frmZGCGToPz.VFG.DispID_0082(var_14, 2)
  loc_110BD33C:     var_94)
  loc_110BD397:     var_30 = CByte("DateToPeriod".00000001h)
  loc_110BD4F1:     Set var_64 = frmZGCGToPz.Label3
  loc_110BD51B:     var_1D0 = var_64
  loc_110BD6D1:     var_94 = frmZGCGToPz.VFG.DispID_0082(var_14, frmZGCGToPz.VFG)
  loc_110BD704:     var_8038 = "正在处理：第[" & frmZGCGToPz.VFG.DispID_0082(var_14, 2) & " - " & frmZGCGToPz.VFG.DispID_0082(var_14, 3) & " - " & var_94 & "]号凭证是否重号"
  loc_110BD723:     var_64.Caption = var_8038
  loc_110BD7B2:     var_803C = frmZGCGToPz.Proc_15_15_110BDEB0(var_30)
  loc_110BD7C7:     If var_1CC <= 0 Then
  loc_110BD7D9:       var_13C = var_30
  loc_110BD870:       var_94)
  loc_110BD909:       var_804C = (frmZGCGToPz.VFG.DispID_0082(var_14, 3) = frmZGCGToPz.VFG.DispID_0082(var_14, 3))
  loc_110BD936:       var_17C = var_804C + 1
  loc_110BD9AD:       var_8054 = (frmZGCGToPz.VFG.DispID_0082(var_14, frmZGCGToPz.VFG) = frmZGCGToPz.VFG.DispID_0082(var_14, ""))
  loc_110BD9D4:       var_1BC = var_8054 + 1
  loc_110BDACF:       If CBool((frmZGCGToPz.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_804C + 1 And var_8054 + 1) Then
  loc_110BDB62:         frmZGCGToPz.VFG.DispID_0082(var_10C, 285267820)
  loc_110BDC96:         frmZGCGToPz.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110BDD2A:         frmZGCGToPz.VFG.DispID_0082(var_10C, "指定的凭证号无效或重号")
  loc_110BDD75:         var_8068 = CLng(frmZGCGToPz.VFG.DispID_0007)
  loc_110BDD93:         var_1D0 = (var_14(1) > 0)
  loc_110BDDB0:         If var_1D0 = 0 Then GoTo loc_110BD7D3
  loc_110BDDB6:       End If
  loc_110BDDC4:     Else
  loc_110BDDCD:     End If
  loc_110BDDDA:     var_14 = 1+var_14
  loc_110BDDDD:     GoTo loc_110BD2A6
  loc_110BDDE2:   End If
  loc_110BDDE2: End If
  loc_110BDDE7: GoTo loc_110BDE78
  loc_110BDE77: Exit Sub
  loc_110BDE78: ' Referenced from: 110BDDE7
End Sub

Private  Proc_15_15_110BDEB0(arg_C, arg_10, arg_14) '110BDEB0
  loc_110BDF49: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110BDF52: var_168 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110BDF7B: If IsNumeric(arg_14) Then
  loc_110BDF8A:   var_8008 = CLng(Val(arg_14))
  loc_110BDF94:   If var_8008 > 0 Then
  loc_110BDFA0:     If var_8008 <= 9999 Then
  loc_110BE01C:       var_8028 = "select * from GL_accvouch where iperiod >=" & CStr(arg_C) & " and isignseq>=" & CStr(0) & " and ino_id>=" & CStr(var_8008)
  loc_110BE031:       var_44 = var_8028
  loc_110BE083:       Set var_1C = "DataMdb".00000000h.00000001h(fs:[00000000h], , , , , var_40, var_8028, var_48)
  loc_110BE0C8:       var_8030 = Proc_0_4_11026BD0(arg_10, , )
  loc_110BE0E9:       var_8034 = CBool(var_1C.EOF)
  loc_110BE0FD:       If var_8034 = 0 Then
  loc_110BE128:         var_F4 = arg_C
  loc_110BE1E6:         var_8040 = (var_1C.DispID_0000("iPeriod") = arg_C) And (var_1C.DispID_0000("isignseq") = (Proc_0_4_11026BD0(arg_10, , ) And 255))
  loc_110BE256:         var_804C = CBool(Not(var_8040 And (var_1C.DispID_0000("ino_id") = var_8008)))
  loc_110BE27B:         If var_804C = 0 Then GoTo loc_110BE280
  loc_110BE27D:       End If
  loc_110BE28B:       var_1C.oClose
  loc_110BE294:     End If
  loc_110BE294:   End If
  loc_110BE294: End If
  loc_110BE29A: GoTo loc_110BE2FF
  loc_110BE2FE: Exit Sub
  loc_110BE2FF: ' Referenced from: 110BE29A
End Sub
