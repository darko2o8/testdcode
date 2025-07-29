VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C908002B2F49FB}#1.2#0"; "C:\Windows\SysWow64\Comdlg32.ocx"
Begin VB.Form frmZGXSToPzTGZP
  Caption = "销售暂估导转凭证（TGZP）"
  BackColor = &H80000005&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmZGXSToPzTGZP.frx":0000
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 345
  ClientWidth = 9255
  ClientHeight = 5700
  Appearance = 0 'Flat
  Begin C1SizerLibCtl.C1Elastic Pic1
    Left = 3300
    Top = 3480
    Width = 5025
    Height = 675
    Visible = 0   'False
    TabStop = 0   'False
    TabIndex = 3
    OleObjectBlob = "frmZGXSToPzTGZP.frx":014A
    Begin VB.Label Label3
      Caption = "正在分析数据，请稍候。。。"
      Left = 210
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
    OleObjectBlob = "frmZGXSToPzTGZP.frx":0293
    Begin AIFCmp1.asxStatusBar sBar
      Left = 0
      Top = 5355
      Width = 12045
      Height = 345
      OleObjectBlob = "frmZGXSToPzTGZP.frx":04BA
    End
    Begin C1SizerLibCtl.C1Elastic C1Elastic2
      Left = 0
      Top = 0
      Width = 12045
      Height = 435
      TabStop = 0   'False
      TabIndex = 1
      OleObjectBlob = "frmZGXSToPzTGZP.frx":05EA
    End
    Begin VSFlex8LCtl.VSFlexGrid VFG
      Left = 0
      Top = 1500
      Width = 12045
      Height = 3840
      TabIndex = 2
      OleObjectBlob = "frmZGXSToPzTGZP.frx":0751
      Begin MSComDlg.CommonDialog dlg
        OleObjectBlob = "frmZGXSToPzTGZP.frx":0BBA
        Left = 0
        Top = 0
      End
    End
    Begin AIFCmp1.asxPanel asxPanel1
      Left = 0
      Top = 450
      Width = 12045
      Height = 1035
      OleObjectBlob = "frmZGXSToPzTGZP.frx":0C1E
      Begin AIFCmp1.asxPowerButton APB
        Index = 3
        Left = 9750
        Top = 615
        Width = 690
        Height = 360
        TabIndex = 9
        OleObjectBlob = "frmZGXSToPzTGZP.frx":0CFE
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 4
        Left = 10485
        Top = 615
        Width = 810
        Height = 360
        TabIndex = 13
        OleObjectBlob = "frmZGXSToPzTGZP.frx":0E9E
      End
      Begin VB.CheckBox Chk
        Caption = "Check1"
        Index = 0
        Left = 8670
        Top = 150
        Width = 615
        Height = 180
        Visible = 0   'False
        TabIndex = 6
        Value = 1
      End
      Begin VB.CheckBox Chk
        Caption = "Check1"
        Index = 1
        Left = 8670
        Top = 0
        Width = 615
        Height = 180
        Visible = 0   'False
        TabIndex = 5
        Value = 1
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 0
        Left = 6705
        Top = 615
        Width = 960
        Height = 360
        TabIndex = 7
        OleObjectBlob = "frmZGXSToPzTGZP.frx":103E
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 1
        Left = 7725
        Top = 615
        Width = 960
        Height = 360
        TabIndex = 8
        OleObjectBlob = "frmZGXSToPzTGZP.frx":1236
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 0
        Left = 120
        Top = 75
        Width = 6525
        Height = 270
        TabIndex = 10
        OleObjectBlob = "frmZGXSToPzTGZP.frx":1406
        ToolTipText = "项目大类"
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 2
        Left = 8730
        Top = 615
        Width = 960
        Height = 360
        Visible = 0   'False
        TabIndex = 11
        OleObjectBlob = "frmZGXSToPzTGZP.frx":1562
      End
      Begin TDBDate6Ctl.TDBDate TDBDate
        Left = 3840
        Top = 690
        Width = 2805
        Height = 285
        TabIndex = 12
        OleObjectBlob = "frmZGXSToPzTGZP.frx":1706
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 1
        Left = 60
        Top = 390
        Width = 3255
        Height = 270
        TabIndex = 14
        OleObjectBlob = "frmZGXSToPzTGZP.frx":19F5
        ToolTipText = "借方科目编码"
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 2
        Left = 3390
        Top = 390
        Width = 3255
        Height = 270
        TabIndex = 15
        OleObjectBlob = "frmZGXSToPzTGZP.frx":1B59
        ToolTipText = "贷方科目编码"
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 3
        Left = 420
        Top = 720
        Width = 2895
        Height = 270
        TabIndex = 16
        OleObjectBlob = "frmZGXSToPzTGZP.frx":1CBD
        ToolTipText = "部门编码"
      End
    End
  End
End

Attribute VB_Name = "frmZGXSToPzTGZP"


Private Sub Form_Load() '110C7600
  Dim var_1C As Variant
  Dim var_24 As var_20.DispID_03E8
  Dim var_20 As var_1C.DispID_03E8
  loc_110C7666: If var_18 <= 3 Then
  loc_110C7689:   var_18 = frmZGXSToPzTGZP.TDBText.UnkVCall_00000040h
  loc_110C76B5:   var_34 = var_20.DispID_03E8
  loc_110C76CA:   Set var_24 = var_20.DispID_03E8
  loc_110C771D:   var_18 = 1+var_18
  loc_110C7722:   GoTo loc_110C765D
  loc_110C7727: End If
  loc_110C7740: Set var_1C = frmZGXSToPzTGZP.TDBDate
  loc_110C7747: var_34 = var_1C.DispID_03E8
  loc_110C775C: Set var_20 = var_1C.DispID_03E8
  loc_110C7768: var_20.UnkVCall_00000030h
  loc_110C77D7: frmZGXSToPzTGZP.TDBDate.DispID_0000 = Date
  loc_110C77F9: Set var_1C = frmZGXSToPzTGZP.APB
  loc_110C7806: var_1C.UnkVCall_00000040h
  loc_110C7844: var_20.DispID_80010007 = var_1C.DispID_03E8
  loc_110C786B: Set var_1C = frmZGXSToPzTGZP.APB
  loc_110C7878: var_1C.UnkVCall_00000040h
  loc_110C78B3: var_20.DispID_80010007 = var_1C.DispID_03E8
  loc_110C78CF: var_8004 = frmZGXSToPzTGZP.Proc_16_9_110BF730(var_1C)
  loc_110C78DC: var_58 = frmZGXSToPzTGZP.getBTData
  loc_110C7904: GoTo loc_110C7927
  loc_110C7926: Exit Sub
  loc_110C7927: ' Referenced from: 110C7904
End Sub

Private Sub Form_Resize() '110C7950
  loc_110C79DD: var_38 = frmZGXSToPzTGZP.Pic1.DispID_80010005
  loc_110C7A01: var_48 = frmZGXSToPzTGZP.Pic1.DispID_80010006
  loc_110C7A14: var_EC = var_48.ScaleWidth
  loc_110C7A4B: If global_110F6000 = 0 Then
  loc_110C7A55: Else
  loc_110C7A60: End If
  loc_110C7A60: var_70 = ((var_EC - CSgn(var_38)) / 2)
  loc_110C7A75: var_F0 = var_48.ScaleHeight
  loc_110C7AB3: If global_110F6000 = 0 Then
  loc_110C7ABD: Else
  loc_110C7AC8: End If
  loc_110C7BD3: frmZGXSToPzTGZP.Pic1.DispID_80011002(var_70, ((var_F0 - CSgn(var_48)) / 2), CSgn(frmZGXSToPzTGZP.Pic1.DispID_80010005), CSgn(frmZGXSToPzTGZP.Pic1.DispID_80010006))
  loc_110C7C1C: GoTo loc_110C7C56
End Sub

Private  TDBText_UnknownEvent_B(arg_C) '110DADC0
  Dim var_6C As frmZGXSToPzTGZP.dlg
  loc_110DAE1D: If arg_C = 0 Then
  loc_110DAE39:   Set var_6C = frmZGXSToPzTGZP.dlg
  loc_110DAE6B:   var_6C.FileName = var_4C
  loc_110DAE8D:   var_6C.DialogTitle = var_4C
  loc_110DAEAF:   var_6C.Filter = var_4C
  loc_110DAECE:   var_6C.CancelError = var_4C
  loc_110DAED8:   var_6C.ShowOpen
  loc_110DAEF0:   var_6C.FileName = var_6C
  loc_110DAF32:   If (var_30 = global_1100AE28) Then
  loc_110DAF44:     var_6C.FileName = Me
  loc_110DAF81:     arg_C = frmZGXSToPzTGZP.TDBText.UnkVCall_00000040h
  loc_110DAFEC:   End If
  loc_110DAFF8: Else
  loc_110DAFFE:   GoTo loc_110DAFEE
  loc_110DB02C:   Exit Sub
  loc_110DB02D: End If
End Sub

Private  APB_UnknownEvent_9(arg_C) '110D82C0
  Dim var_24 As Variant
  loc_110D833D: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110D8346: var_E8 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110D836D: arg_C = frmZGXSToPzTGZP.APB.UnkVCall_00000040h
  loc_110D83B1: var_D4 = var_28.DispID_FFFFFDFA
  loc_110D83E1: var_8008 = (var_D4 = "加载数据")
  loc_110D83E5: If var_8008 = 0 Then
  loc_110D8429:   frmZGXSToPzTGZP.TDBText.UnkVCall_00000040h
  loc_110D8471:   var_20 = var_28.DispID_0000
  loc_110D8473:   var_EC = var_1C
  loc_110D8487:   var_8014 = Scripting.FileSystemObject.FileExists
  loc_110D84CF:   If Not (var_BC) Then
  loc_110D851A:     var_80 = "文件不存在或非法路径！ "
  loc_110D853B:     MsgBox(var_80, 64, "提示", 10, 10)
  loc_110D8561:   Else
  loc_110D8571:     If var_18 <= 3 Then
  loc_110D8598:       var_18 = frmZGXSToPzTGZP.TDBText.UnkVCall_00000040h
  loc_110D85CC:       var_40 = var_28.DispID_0000
  loc_110D862A:       If (Proc_0_11_11029000(8, var_28, var_20) = global_1100AE28) + 1 = 0 Then
  loc_110D863B:         var_18 = 1+var_18
  loc_110D863E:         GoTo loc_110D8568
  loc_110D8643:       End If
  loc_110D8664:       var_18 = frmZGXSToPzTGZP.TDBText.UnkVCall_00000040h
  loc_110D86B8:       var_80 = "提示"
  loc_110D8700:       MsgBox(var_28.DispID_8001004A & "不能为空，请输入。 ", 64, var_80, 10, 10)
  loc_110D8743:     Else
  loc_110D8755:       If frmZGXSToPzTGZP.FillData < 0 Then
  loc_110D8767:         var_BC = CheckObj(8, global_1100D2F0, 1788)
  loc_110D8772:       End If
  loc_110D877E:       var_BC = var_D4
  loc_110D8782:       If var_BC = 0 Then
  loc_110D87A7:         var_80 = "提示信息"
  loc_110D87B8:         var_48 = var_80
  loc_110D87DF:         var_30 = "是否取消数据载入？" & vbCrLf & "取消数据载入，数据将全部清空。"
  loc_110D87FE:         MsgBox(var_30, 292, var_48, var_58, var_68)
  loc_110D8838:         If (MsgBox(var_30, 292, var_48, var_58, var_68) = 6) = 0 Then GoTo loc_110D88DE
  loc_110D8849:       Else
  loc_110D8855:         var_8034 = var_D4 & "凭证导入"
  loc_110D8859:         If var_8034 = 0 Then
  loc_110D885E:           var_8038 = frmZGXSToPzTGZP.Proc_16_12_110CF170("取消加载")
  loc_110D8866:         Else
  loc_110D8876:           If var_D4 & "导出" Then
  loc_110D8884:             var_8040 = var_D4 & global_1100EBD4
  loc_110D8888:             If var_8040 = 0 Then
  loc_110D88B5:               Set var_24 = CInt(8)
  loc_110D88BD:               var_8048 = Global.Unload var_28
  loc_110D88DE:             End If
  loc_110D88DE:           End If
  loc_110D88DE:         End If
  loc_110D88DE:       End If
  loc_110D88DE:     End If
  loc_110D88DE:   End If
  loc_110D88DE: End If
  loc_110D88EA: GoTo loc_110D8925
  loc_110D8924: Exit Sub
  loc_110D8925: ' Referenced from: 110D88EA
End Sub

Public Sub ExitForm(Cancel, UnloadMode) '110BF650
  Dim var_18 As Global
  loc_110BF68F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110BF6BA: Set var_18 = Me
  loc_110BF6C2: var_8008 = Global.Unload
  loc_110BF6FC: GoTo loc_110BF708
  loc_110BF707: Exit Sub
  loc_110BF708: ' Referenced from: 110BF6FC
End Sub

Public Function FillData() '110C1080
  Dim var_CC As Variant
  Dim var_64 As Variant
  Dim var_58 As Variant
  Dim var_3C As Variant
  Dim var_34 As Me
  Dim var_2C As ADODB.Recordset
  Dim var_D4 As Me
  loc_110C1215: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110C122B: var_2DC = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110C126F: frmZGXSToPzTGZP.VFG.DispID_0007 = 1
  loc_110C1292: Set var_CC = frmZGXSToPzTGZP.Label3
  loc_110C129C: var_2A4 = var_CC
  loc_110C12A2: var_CC.Caption = "正在打开Excel数据表，请稍候。。。"
  loc_110C1315: frmZGXSToPzTGZP.Pic1.DispID_80010007 = True
  loc_110C1341: frmZGXSToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_110C1367: Set var_CC = frmZGXSToPzTGZP.TDBText
  loc_110C1377: var_CC.UnkVCall_00000040h
  loc_110C139B: var_E0 = var_D0
  loc_110C13C2: var_48 = Proc_0_11_11029000(9, var_CC, 1)
  loc_110C13ED: Set var_CC = frmZGXSToPzTGZP.TDBText
  loc_110C13FF: var_2A4 = var_CC
  loc_110C1405: var_CC.UnkVCall_00000040h
  loc_110C1436: var_E0 = var_D0
  loc_110C1450: var_44 = Proc_0_11_11029000(9, var_CC, 2)
  loc_110C147B: Set var_CC = frmZGXSToPzTGZP.TDBText
  loc_110C148D: var_2A4 = var_CC
  loc_110C1493: var_CC.UnkVCall_00000040h
  loc_110C14BD: var_E0 = var_D0
  loc_110C14DE: var_24 = Proc_0_11_11029000(9, var_CC, 3)
  loc_110C14FD: var_8010 = CreateObject(global_1100D5A4)
  loc_110C1508: Set var_64 = CreateObject(global_1100D5A4)
  loc_110C151B: var_D4 = var_64.UnkVCall_000000D0h
  loc_110C15CF: Set var_CC = frmZGXSToPzTGZP.TDBText
  loc_110C15E0: var_2A4 = var_CC
  loc_110C15E6: var_CC.UnkVCall_00000040h
  loc_110C186C: var_94 = var_D0.DispID_0000
  loc_110C187C: var_94 = var_D4.UnkVCall_0000004Ch
  loc_110C18F6: var_CC = 0.Tag
  loc_110C199E: var_CC.Activate
  loc_110C1A08: Set var_8C = var_CC.UsedRange
  loc_110C1A37: Set var_CC = frmZGXSToPzTGZP.Label3
  loc_110C1A41: var_2A4 = var_CC
  loc_110C1A47: var_CC.Caption = "正在填充数据，请稍候。。。"
  loc_110C1ABA: frmZGXSToPzTGZP.Pic1.DispID_80010007 = True
  loc_110C1AE7: frmZGXSToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_110C1B21: Set var_CC = frmZGXSToPzTGZP.APB
  loc_110C1B2F: var_2A4 = var_CC
  loc_110C1B35: var_CC.UnkVCall_00000040h
  loc_110C1BCB: Set var_CC = frmZGXSToPzTGZP.APB
  loc_110C1BD9: var_2A4 = var_CC
  loc_110C1BDF: var_CC.UnkVCall_00000040h
  loc_110C1C85: frmZGXSToPzTGZP.APB.UnkVCall_00000040h
  loc_110C1D74: var_108 = 1100D68Ch & var_8C.Rows.Count
  loc_110C1E07: frmZGXSToPzTGZP.sBar.DispID_6803001E(var_108 & "条记录")
  loc_110C1E52: var_34 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[T_CY_ZGXSTGZP_Temp]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1) "
  loc_110C1E61: var_8020 = var_34 & "CREATE TABLE [T_CY_ZGXSTGZP_Temp](cCode VARCHAR(50) NULL,cCusCode VARCHAR(50) NULL,cCusName VARCHAR(50) NULL,cInvCode VARCHAR(50) NULL,cDepCode VARCHAR(50) NULL,cXL VARCHAR(50) NULL,iQuantity float NULL,iMoney Money NULL,iMoney1 Money NULL)"
  loc_110C1E6C: var_34 = var_8020
  loc_110C1EB1: var_E8 = UnkObj.UnkVCall_00000040h
  loc_110C1EF5: var_34 = "DELETE FROM [T_CY_ZGXSTGZP_Temp]"
  loc_110C1F8F: Set var_CC = frmZGXSToPzTGZP.TDBDate
  loc_110C1FAD: var_F0 = var_CC.DispID_004E
  loc_110C1FCB: var_108)
  loc_110C2029: var_84 = CByte("DateToPeriod".00000001h)
  loc_110C20BC: var_3C.UnkVCall_00000064h
  loc_110C2154: var_F8 = var_CC.Cells(1, 1).value
  loc_110C2168: var_70 = Proc_0_11_11029000(var_F8, var_3C, 2)
  loc_110C229B: var_30 = Proc_0_11_11029000(var_CC.Cells(1, 2).value, var_CC, var_1B4)
  loc_110C2303: var_F8 = var_8C.Rows.Count
  loc_110C235E: If var_18 <= CLng(var_F8 + 1) Then
  loc_110C2369:   If global_56 = 0 Then
  loc_110C23DA:     var_F8.BackColor = var_1B0
  loc_110C253A:     var_330 = (Proc_0_11_11029000(var_D0.Cells(var_18, 1).value, var_D0, var_CC) = "汇总") + 1
  loc_110C25EC:     var_2AC = (Proc_0_11_11029000(var_CC.Cells(var_18, 1).value, var_1B0, var_1AC) = global_1100AE28) + 1
  loc_110C263E:     If var_2AC = 0 Then
  loc_110C26FD:       Set var_CC = frmZGXSToPzTGZP.sBar
  loc_110C2704:       var_CC.DispID_6803001E("正在填充数据：" & CStr(vbNull) & "条记录")
  loc_110C275E:       var_2A4 = var_2C
  loc_110C2764:       var_2A0 = ADODB.Recordset.State
  loc_110C278F:       If var_2A0 = 1 Then
  loc_110C27AD:         var_2A4 = var_2C
  loc_110C27B3:         var_8060 = ADODB.Recordset.Close
  loc_110C27D7:       End If
  loc_110C2841:       ADODB.Recordset.BackColor = 1
  loc_110C28D9:       var_F8 = var_CC.Cells(var_18, 4).value
  loc_110C292D:       var_1C0 = var_70
  loc_110C2933:       var_1B0 = var_44
  loc_110C2960:       var_1E0 = Proc_0_11_11029000(var_F8, var_CC, var_CC)
  loc_110C2966:       var_1D0 = var_30
  loc_110C2977:       var_1F0 = var_24
  loc_110C297D:       var_8068 = Proc_0_10_11028DD0(&H4008, "INSERT INTO [T_CY_ZGXSTGZP_Temp](cCode,cCusCode,cCusName,cInvCode,cDepCode,cXL,iQuantity,iMoney,iMoney1) VALUES (", var_CC)
  loc_110C29F9:       var_8080 = Proc_0_10_11028DD0(&H4008, var_D0 & Proc_0_10_11028DD0(&H4008, var_CC & var_8068 & global_1100AC40, 2) & global_1100AC40, var_CC)
  loc_110C2A75:       var_8098 = Proc_0_10_11028DD0(&H4008, var_CC & Proc_0_10_11028DD0(&H4008, 1 & var_8080 & global_1100AC40, var_D0) & global_1100AC40, 0)
  loc_110C2B4B:       var_F8.BackColor = CInt(1)
  loc_110C2BDD:       var_168 = var_D4.Cells(var_18, 10).value
  loc_110C2CEC:       var_F8 = var_CC.Cells(var_18, 3).value
  loc_110C2EA5:       var_128 = var_D0.Cells(var_18, 6).value
  loc_110C2EC7:       var_148 = var_D4 & Proc_0_10_11028DD0(var_F8, var_D0 & var_8098 & global_1100AC40, var_CC) & global_1100AC40 & var_128 & 1100AC40h
  loc_110C3094:       var_F8 = var_CC.Cells(var_18, 12).value
  loc_110C313A:       var_34 = var_148 & Format(var_168, "0.00") & 1100AC40h & Format(var_F8, "0.00") & 1100BD88h
  loc_110C320D:       var_28 = var_28(1)
  loc_110C3218:       If var_18 Mod 00000064h = 0 Then
  loc_110C321A:         DoEvents
  loc_110C3220:       End If
  loc_110C3230:       var_18 = 1+var_18
  loc_110C3233:       GoTo loc_110C2358
  loc_110C3238:     End If
  loc_110C3284:     frmZGXSToPzTGZP.VFG.DispID_0007 = 1
  loc_110C3299:     global_56 = 0
  loc_110C32C1:     Set var_CC = frmZGXSToPzTGZP.APB
  loc_110C32D3:     var_2A4 = var_CC
  loc_110C32D9:     var_CC.UnkVCall_00000040h
  loc_110C336F:     Set var_CC = frmZGXSToPzTGZP.APB
  loc_110C3381:     var_2A4 = var_CC
  loc_110C3387:     var_CC.UnkVCall_00000040h
  loc_110C3431:     frmZGXSToPzTGZP.APB.UnkVCall_00000040h
  loc_110C3496:   Else
  loc_110C34BB:     Set var_CC = frmZGXSToPzTGZP.APB
  loc_110C34CD:     var_2A4 = var_CC
  loc_110C34D3:     var_CC.UnkVCall_00000040h
  loc_110C3569:     Set var_CC = frmZGXSToPzTGZP.APB
  loc_110C357B:     var_2A4 = var_CC
  loc_110C3581:     var_CC.UnkVCall_00000040h
  loc_110C362B:     frmZGXSToPzTGZP.APB.UnkVCall_00000040h
  loc_110C368B:   End If
  loc_110C3696: End If
  loc_110C36D2: var_68 = global_11012A94 & CStr(var_84) & "月销售"
  loc_110C3710: var_2A4 = var_2C
  loc_110C3716: var_2A0 = ADODB.Recordset.State
  loc_110C3741: If var_2A0 = 1 Then
  loc_110C375F:   var_2A4 = var_2C
  loc_110C3765:   var_80CC = ADODB.Recordset.Close
  loc_110C3789: End If
  loc_110C37C2: var_80D8 = "SELECT '" & var_48 & "' AS cCode,cCusCode,cCusName,cXL,SUM(iMoney1) AS iMoney1 " & "FROM [T_CY_ZGXSTGZP_Temp] GROUP BY cCusCode,cCusName,cXL"
  loc_110C383E: var_2A4 = var_2C
  loc_110C386D: var_80E0 = ADODB.Recordset.Open(var_80D8, var_1B4, var_80D8, var_1AC, 9)
  loc_110C3900: var_2AC = ADODB.Recordset.Fields
  loc_110C391B: ADODB.Recordset.8 = Forms
  loc_110C393F: var_D0 = 0
  loc_110C3949: var_E0 = var_D0
  loc_110C39AC: If (Proc_0_11_11029000(9, var_1B4, "cCusName") = global_1100AE28) Then
  loc_110C3A21:   var_2AC = ADODB.Recordset.Fields
  loc_110C3A3C:   ADODB.Recordset.8 = Forms
  loc_110C3A66:   var_E0 = var_D0
  loc_110C3B19:   var_1C0 = "cCusName"
  loc_110C3B2E:   var_2BC = ADODB.Recordset.Fields
  loc_110C3B49:   ADODB.Recordset.8 = Forms
  loc_110C3BB4:   var_50 = "cXL" & Proc_0_11_11029000(9, var_68, var_1B4) & "/" & var_F8
  loc_110C3C10: End If
  loc_110C3C39: var_2A4 = var_2C
  loc_110C3C3F: var_29C = ADODB.Recordset.EOF
  loc_110C3C65: If var_29C = 0 Then
  loc_110C3F3C:   var_108 = "1" & Chr(9) & 1100AE28h & Chr(9) & frmZGXSToPzTGZP.TDBDate.DispID_004E & Chr(9) & 1100D6D4h & Chr(9) & 1100C008h & Chr(9) & var_50
  loc_110C3FAD:   var_2A4 = var_2C
  loc_110C3FEE:   var_2AC = ADODB.Recordset.Fields
  loc_110C4024:   ADODB.Recordset.8 = Forms
  loc_110C412F:   var_2A4 = var_2C
  loc_110C417A:   var_2AC = ADODB.Recordset.Fields
  loc_110C41A6:   ADODB.Recordset.8 = Forms
  loc_110C42C8:   var_108 = 9 & Chr(9) & Proc_0_11_11029000(9, var_1C4, "cCode") & Chr(9) & Proc_0_11_11029000(9, var_1C4, "iMoney1") & Chr(9) & 1100C008h
  loc_110C4411:   var_1B0 = var_108 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C4570:   var_108 = var_1B0 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C46B9:   var_1B0 = var_108 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C4801:   var_2A4 = var_2C
  loc_110C484C:   var_2AC = ADODB.Recordset.Fields
  loc_110C4878:   ADODB.Recordset.8 = Forms
  loc_110C48F8:   var_128 = var_1B0 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1C4, "cCusCode")
  loc_110C4AB8:   var_54 = var_128 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C4AF6:   var_2A4 = var_2C
  loc_110C4AFC:   var_8180 = ADODB.Recordset.MoveNext
  loc_110C4B72:   frmZGXSToPzTGZP.VFG.DispID_0080(var_54)
  loc_110C4B87:   GoTo loc_110C3C16
  loc_110C4B8C: End If
  loc_110C4BAF: var_2A4 = var_2C
  loc_110C4BB5: var_2A0 = ADODB.Recordset.State
  loc_110C4BE0: If var_2A0 = 1 Then
  loc_110C4BFE:   var_2A4 = var_2C
  loc_110C4C04:   var_818C = ADODB.Recordset.Close
  loc_110C4C28: End If
  loc_110C4C3F: var_8190 = "SELECT cCode,cCusCode,cCusName,cDepCode,cInvCode,cXL,SUM(iQuantity) AS iQuantity,SUM(iMoney) AS iMoney " & "FROM [T_CY_ZGXSTGZP_Temp] GROUP BY cCode,cCusCode,cCusName,cDepCode,cInvCode,cXL"
  loc_110C4CB4: var_2A4 = var_2C
  loc_110C4CF1: var_8198 = ADODB.Recordset.Open(var_8190, var_1B4, var_8190, var_1AC, 9)
  loc_110C4D38: var_2A4 = var_2C
  loc_110C4D3E: var_29C = ADODB.Recordset.EOF
  loc_110C4D64: If var_29C = 0 Then
  loc_110C503B:   var_108 = "1" & Chr(9) & 1100AE28h & Chr(9) & frmZGXSToPzTGZP.TDBDate.DispID_004E & Chr(9) & 1100D6D4h & Chr(9) & 1100C008h & Chr(9) & var_50
  loc_110C50AC:   var_2A4 = var_2C
  loc_110C50ED:   var_2AC = ADODB.Recordset.Fields
  loc_110C5123:   ADODB.Recordset.8 = Forms
  loc_110C52B6:   var_2A4 = var_2C
  loc_110C52F7:   var_2AC = ADODB.Recordset.Fields
  loc_110C532D:   ADODB.Recordset.8 = Forms
  loc_110C53AD:   var_128 = 9 & Chr(9) & Proc_0_11_11029000(9, var_1C4, "cCode") & Chr(9) & 1100C008h & Chr(9) & Proc_0_11_11029000(9, var_1C4, "iMoney")
  loc_110C5438:   var_2A4 = var_2C
  loc_110C5483:   var_2AC = ADODB.Recordset.Fields
  loc_110C54AF:   ADODB.Recordset.8 = Forms
  loc_110C55C0:   var_F8 = var_128 & Chr(9) & Proc_0_11_11029000(9, var_1C4, "iQuantity") & Chr(9) & Proc_0_11_11029000(9, var_1C4, "iQuantity") & Chr(9) & Proc_0_11_11029000(9, var_1C4, "iQuantity") & Chr(9)
  loc_110C56E4:   var_81E8 = var_F8 & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C58EA:   var_2A4 = var_2C
  loc_110C592B:   var_2AC = ADODB.Recordset.Fields
  loc_110C5961:   ADODB.Recordset.8 = Forms
  loc_110C59E1:   var_128 = var_81E8 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1C4, "cDepCode")
  loc_110C5AF4:   var_2A4 = var_2C
  loc_110C5B35:   var_2AC = ADODB.Recordset.Fields
  loc_110C5B6B:   ADODB.Recordset.8 = Forms
  loc_110C5C3E:   var_1B0 = var_128 & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1C4, "cCusCode") & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1C4, "cCusCode")
  loc_110C5D86:   var_2A4 = var_2C
  loc_110C5DC7:   var_2AC = ADODB.Recordset.Fields
  loc_110C5DFD:   ADODB.Recordset.8 = Forms
  loc_110C5E7D:   var_128 = var_1B0 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1C4, "cInvCode")
  loc_110C5E8B:   var_54 = var_128
  loc_110C5EE3:   var_2A4 = var_2C
  loc_110C5EE9:   var_822C = ADODB.Recordset.MoveNext
  loc_110C5F5F:   frmZGXSToPzTGZP.VFG.DispID_0080(var_54)
  loc_110C5F74:   GoTo loc_110C4D15
  loc_110C5F79: End If
  loc_110C5F9C: var_2A4 = var_2C
  loc_110C5FA2: var_2A0 = ADODB.Recordset.State
  loc_110C5FCD: If var_2A0 = 1 Then
  loc_110C5FEB:   var_2A4 = var_2C
  loc_110C5FF1:   var_8238 = ADODB.Recordset.Close
  loc_110C6015: End If
  loc_110C60A1: var_2A4 = var_2C
  loc_110C60DE: var_8244 = ADODB.Recordset.Open("SELECT '21710112' AS cCode,SUM(ISNULL(iMoney1,0)-ISNULL(iMoney,0)) AS iMoney " & "FROM [T_CY_ZGXSTGZP_Temp] ", var_1B4, "SELECT '21710112' AS cCode,SUM(ISNULL(iMoney1,0)-ISNULL(iMoney,0)) AS iMoney " & "FROM [T_CY_ZGXSTGZP_Temp] ", var_1AC, 9)
  loc_110C6125: var_2A4 = var_2C
  loc_110C612B: var_29C = ADODB.Recordset.EOF
  loc_110C6151: If var_29C = 0 Then
  loc_110C6428:   var_108 = "1" & Chr(9) & 1100AE28h & Chr(9) & frmZGXSToPzTGZP.TDBDate.DispID_004E & Chr(9) & 1100D6D4h & Chr(9) & 1100C008h & Chr(9) & var_50
  loc_110C6499:   var_2A4 = var_2C
  loc_110C64DA:   var_2AC = ADODB.Recordset.Fields
  loc_110C6510:   ADODB.Recordset.8 = Forms
  loc_110C66A3:   var_2A4 = var_2C
  loc_110C66E4:   var_2AC = ADODB.Recordset.Fields
  loc_110C671A:   ADODB.Recordset.8 = Forms
  loc_110C673E:   var_D0 = 0
  loc_110C6748:   var_100 = var_D0
  loc_110C679A:   var_128 = 9 & Chr(9) & Proc_0_11_11029000(9, var_1C4, "cCode") & Chr(9) & 1100C008h & Chr(9) & Proc_0_11_11029000(9, var_1C4, "iMoney")
  loc_110C6AE4:   var_108 = var_128 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C6C2D:   var_1B0 = var_108 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C6D8C:   var_108 = var_1B0 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C6EAA:   var_54 = var_108 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110C6EE8:   var_2A4 = var_2C
  loc_110C6EEE:   var_82B8 = ADODB.Recordset.MoveNext
  loc_110C6F64:   frmZGXSToPzTGZP.VFG.DispID_0080(var_54)
  loc_110C6F79:   GoTo loc_110C6102
  loc_110C6F7E: End If
  loc_110C703C: frmZGXSToPzTGZP.sBar.DispID_6803001E("有效数据共" & CStr(var_28) & global_1100FE7C)
  loc_110C70A7: frmZGXSToPzTGZP.APB.UnkVCall_00000040h
  loc_110C7139: Set var_CC = frmZGXSToPzTGZP.APB
  loc_110C7147: var_2A4 = var_CC
  loc_110C714D: var_CC.UnkVCall_00000040h
  loc_110C71DF: Set var_CC = frmZGXSToPzTGZP.APB
  loc_110C71ED: var_2A4 = var_CC
  loc_110C71F3: var_CC.UnkVCall_00000040h
  loc_110C72A3: frmZGXSToPzTGZP.Pic1.DispID_80010007 = var_1B0
  loc_110C72C2: Set var_CC = frmZGXSToPzTGZP.TDBText
  loc_110C72D2: var_CC.UnkVCall_00000040h
  loc_110C732A: var_E0 = var_D0
  loc_110C73A2: var_CC.ForeColor = False
  loc_110C73DB: var_1B4 = var_64.UnkVCall_00000398h
  loc_110C7410: Set var_3C = {000208D7-0000-0000-C000000000000046}()
  loc_110C7420: Set var_58 = {000208DA-0000-0000-C000000000000046}()
  loc_110C7430: Set var_64 = {000208D5-0000-0000-C000000000000046}()
  loc_110C7447: GoTo loc_110C7544
  loc_110C7543: Exit Function
  loc_110C7544: ' Referenced from: 110C7447
End Function

Public Function getWBHL(sWhere) '110D8960
  Dim var_1C As ADODB.Recordset
  Dim var_2C As Me
  loc_110D89C0: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110D89CC: var_98 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110D89F4: var_40 = Trim(sWhere)
  loc_110D8A25: If (var_40 <> 1100AE28h) Then
  loc_110D8A53:   var_20 = "SELECT * FROM exch WHERE 1=1 " & " AND " & sWhere
  loc_110D8A60: Else
  loc_110D8A6C: End If
  loc_110D8A7C: var_20 = var_20 & " order by cexch_name, itype, iperiod, cdate"
  loc_110D8AE6: var_78 = var_1C
  loc_110D8AF5: var_8018 = ADODB.Recordset.Open(var_20, var_5C, var_20, var_54, 9)
  loc_110D8B5B: If ADODB.Recordset.EOF Then
  loc_110D8B6A:   var_24 = CStr(0)
  loc_110D8B75: Else
  loc_110D8B97:   var_2C = ADODB.Recordset.Fields
  loc_110D8BC4:   var_58 = "NFLAT"
  loc_110D8BDD:   ADODB.Recordset.8 = Forms
  loc_110D8C2E:   var_24 = var_40
  loc_110D8C50: End If
  loc_110D8C6E: var_8030 = ADODB.Recordset.Close
  loc_110D8C8D: GoTo loc_110D8CCB
  loc_110D8C93: If var_4 Then
  loc_110D8C9E: End If
  loc_110D8CCA: Exit Function
  loc_110D8CCB: ' Referenced from: 110D8C8D
End Function

Public Function getBTData() '110D9FA0
  Dim var_24 As ADODB.Recordset
  Dim var_38 As Variant
  loc_110DA024: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110DA02E: On Error GoTo loc_110DA700
  loc_110DA069: var_28 = 1 & "IF NOT EXISTS (SELECT * FROM [" & "]..Sysobjects "
  loc_110DA0D4: var_8018 =  & var_28 & "WHERE Name = 'T_CY_ZGXS_Setting') " & "CREATE TABLE [" & "]..[T_CY_ZGXS_Setting](cJFKmCode VARCHAR(50) NULL," & "cDFKmCode VARCHAR(50) NULL,cDepCode VARCHAR(50) NULL,bDep Bit NOT NULL)"
  loc_110DA0DB: var_28 = var_8018
  loc_110DA10B: var_54 = UnkObj.UnkVCall_00000040h
  loc_110DA15D: var_28 = var_38 & "SELECT * FROM [" & "]..[T_CY_ZGXS_Setting]"
  loc_110DA197: var_BC = ADODB.Recordset.State
  loc_110DA1BC: If var_BC = 1 Then
  loc_110DA1D8:   var_802C = ADODB.Recordset.Close
  loc_110DA1F6: End If
  loc_110DA276: var_8034 = ADODB.Recordset.Open(var_28, var_90, var_28, var_88, 9)
  loc_110DA2C9: var_B8 = ADODB.Recordset.EOF
  loc_110DA2E5: If var_B8 = 0 Then
  loc_110DA30D:   var_38 = ADODB.Recordset.Fields
  loc_110DA32B:   var_8C = "cjfkmcode"
  loc_110DA35F:   ADODB.Recordset.8 = Forms
  loc_110DA3CA:   frmZGXSToPzTGZP.TDBText.UnkVCall_00000040h
  loc_110DA400:   var_44.DispID_0000 = Proc_0_11_11029000(9, var_90, "cjfkmcode")
  loc_110DA466:   var_D0 = ADODB.Recordset.Fields
  loc_110DA471:   var_8C = "cDFKmCode"
  loc_110DA4A5:   ADODB.Recordset.8 = Forms
  loc_110DA513:   frmZGXSToPzTGZP.TDBText.UnkVCall_00000040h
  loc_110DA549:   var_44.DispID_0000 = Proc_0_11_11029000(9, var_90, "cDFKmCode")
  loc_110DA5AF:   var_D0 = ADODB.Recordset.Fields
  loc_110DA5BA:   var_8C = "cDepCode"
  loc_110DA5EE:   ADODB.Recordset.8 = Forms
  loc_110DA65C:   frmZGXSToPzTGZP.TDBText.UnkVCall_00000040h
  loc_110DA692:   var_44.DispID_0000 = Proc_0_11_11029000(9, var_90, "cDepCode")
  loc_110DA6BF: End If
  loc_110DA6E7: If ADODB.Recordset.Close < 0 Then
  loc_110DA6F9:   var_8058 = CheckObj(var_24, global_1100ADFC, 128)
  loc_110DA700:   ' Referenced from: 110DA02E
  loc_110DA705:   var_805C = Err
  loc_110DA710:   Set var_38 = Err
  loc_110DA795:   MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_110DA7C2: End If
  loc_110DA7C2: Exit Sub
  loc_110DA7CD: GoTo loc_110DA816
  loc_110DA815: Exit Function
  loc_110DA816: ' Referenced from: 110DA7CD
End Function

Public Function UpdateBTData() '110DA860
  Dim var_48 As Variant
  Dim var_50 As frmZGXSToPzTGZP.TDBText
  Dim var_58 As frmZGXSToPzTGZP.TDBText
  Dim var_20 As Me
  loc_110DA8EA: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110DA8F2: var_F8 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110DA8FA: On Error GoTo loc_110DAC3F
  loc_110DA935: var_20 = 1 & "DELETE FROM [" & "]..[T_CY_ZGXS_Setting]"
  loc_110DA974: var_6C = UnkObj.UnkVCall_00000040h
  loc_110DA9FD: Set var_48 = frmZGXSToPzTGZP.TDBText
  loc_110DAA03: var_D0 = var_48
  loc_110DAA12: var_48.UnkVCall_00000040h
  loc_110DAA52: Set var_50 = frmZGXSToPzTGZP.TDBText
  loc_110DAA58: var_D8 = var_50
  loc_110DAA67: var_50.UnkVCall_00000040h
  loc_110DAA88: var_54 = 0
  loc_110DAA8F: var_74 = var_54
  loc_110DAAA7: Set var_58 = frmZGXSToPzTGZP.TDBText
  loc_110DAAB8: var_58.UnkVCall_00000040h
  loc_110DAAD3: var_5C = 0
  loc_110DAADA: var_84 = var_5C
  loc_110DAAF2: var_8018 = Proc_0_10_11028DD0(9, var_48 & "INSERT INTO [" & "]..[T_CY_ZGXS_Setting]" & "(cJFKmCode,cDFKmCode,cDepCode,bDep) VALUES (", var_58)
  loc_110DAB56: var_8034 = var_54 & Proc_0_10_11028DD0(9, var_50 & Proc_0_10_11028DD0(9, 3 & var_8018 & global_1100AC40, var_5C) & global_1100AC40, 2)
  loc_110DABD4: var_20 = var_8034 & global_1100AC40 & "0)"
  loc_110DAC3A: GoTo loc_110DAD10
  loc_110DAC3F: ' Referenced from: 110DA8FA
  loc_110DAC44: var_8040 = Err
  loc_110DAC4F: Set var_48 = Err
  loc_110DACE0: MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_110DAD10: ' Referenced from: 110DAC3A
  loc_110DAD10: Exit Sub
  loc_110DAD1B: GoTo loc_110DAD8A
  loc_110DAD89: Exit Function
  loc_110DAD8A: ' Referenced from: 110DAD1B
End Function

Private Sub Proc_16_9_110BF730
  Dim var_58 As frmZGXSToPzTGZP.VFG
  loc_110BF771: Set var_58 = frmZGXSToPzTGZP.VFG
  loc_110BF7C2: var_58.DispID_005D = frmZGXSToPzTGZP.VFG
  loc_110BF803: var_58.DispID_0067 = frmZGXSToPzTGZP.VFG
  loc_110BF822: var_58.DispID_0041 = frmZGXSToPzTGZP.VFG
  loc_110BF8CC: var_58.DispID_00A5("...")
  loc_110BF9F4: var_58.DispID_008A(4)
  loc_110BFA37: var_58.DispID_0079(450)
  loc_110BFA77: var_58.DispID_007B(True)
  loc_110BFA9B: var_58.DispID_0019 = True
  loc_110BFAE0: var_58.DispID_0090("业务号")
  loc_110BFB23: var_58.DispID_0077(4)
  loc_110BFB66: var_58.DispID_0078(700)
  loc_110BFBAE: var_58.DispID_0090("状态")
  loc_110BFBF4: var_58.DispID_0077(4)
  loc_110BFC3A: var_58.DispID_0078(700)
  loc_110BFC82: var_58.DispID_0090("制单日期")
  loc_110BFCC8: var_58.DispID_0077(1)
  loc_110BFD0E: var_58.DispID_0078(1000)
  loc_110BFD53: var_58.DispID_0090("凭证类别字")
  loc_110BFD95: var_58.DispID_0077(4)
  loc_110BFDD7: var_58.DispID_0078(700)
  loc_110BFE1F: var_58.DispID_0090("附单据数")
  loc_110BFE63: var_58.DispID_0077(var_3C)
  loc_110BFEA9: var_58.DispID_0078(var_3C)
  loc_110BFEF1: var_58.DispID_0090(var_3C)
  loc_110BFF37: var_58.DispID_0077(var_3C)
  loc_110BFF7D: var_58.DispID_0078(var_3C)
  loc_110BFFC5: var_58.DispID_0090(var_3C)
  loc_110C000B: var_58.DispID_0077(var_3C)
  loc_110C0051: var_58.DispID_0078(var_3C)
  loc_110C0099: var_58.DispID_0090(var_3C)
  loc_110C00DD: var_58.DispID_0077(var_3C)
  loc_110C0123: var_58.DispID_0078(var_3C)
  loc_110C016B: var_58.DispID_009C(var_3C)
  loc_110C01B3: var_58.DispID_0090(var_3C)
  loc_110C01F9: var_58.DispID_0077(var_3C)
  loc_110C023F: var_58.DispID_0078(var_3C)
  loc_110C0287: var_58.DispID_009C(var_3C)
  loc_110C02CF: var_58.DispID_0090(var_3C)
  loc_110C0315: var_58.DispID_0077(var_3C)
  loc_110C035B: var_58.DispID_0078(var_3C)
  loc_110C03A3: var_58.DispID_009C(var_3C)
  loc_110C03EB: var_58.DispID_0090(var_3C)
  loc_110C0431: var_58.DispID_0077(var_3C)
  loc_110C0477: var_58.DispID_0078(var_3C)
  loc_110C04BF: var_58.DispID_009C(var_3C)
  loc_110C0507: var_58.DispID_0090(var_3C)
  loc_110C054D: var_58.DispID_0077(var_3C)
  loc_110C0593: var_58.DispID_0078(var_3C)
  loc_110C05DB: var_58.DispID_009C(var_3C)
  loc_110C0623: var_58.DispID_0090(var_3C)
  loc_110C0669: var_58.DispID_0077(var_3C)
  loc_110C06AF: var_58.DispID_0078(var_3C)
  loc_110C06F7: var_58.DispID_0090(var_3C)
  loc_110C073D: var_58.DispID_0077(var_3C)
  loc_110C0783: var_58.DispID_0078(var_3C)
  loc_110C07CB: var_58.DispID_0090(var_3C)
  loc_110C0811: var_58.DispID_0077(var_3C)
  loc_110C0857: var_58.DispID_0078(var_3C)
  loc_110C089F: var_58.DispID_0090(var_3C)
  loc_110C08E5: var_58.DispID_0077(var_3C)
  loc_110C092B: var_58.DispID_0078(var_3C)
  loc_110C0973: var_58.DispID_0090(var_3C)
  loc_110C09B9: var_58.DispID_0077(var_3C)
  loc_110C09FF: var_58.DispID_0078(var_3C)
  loc_110C0A47: var_58.DispID_0090(var_3C)
  loc_110C0A8D: var_58.DispID_0077(var_3C)
  loc_110C0AD3: var_58.DispID_0078(var_3C)
  loc_110C0B1B: var_58.DispID_0090(var_3C)
  loc_110C0B61: var_58.DispID_0077(var_3C)
  loc_110C0BA7: var_58.DispID_0078(var_3C)
  loc_110C0BEF: var_58.DispID_0090(var_3C)
  loc_110C0C35: var_58.DispID_0077(var_3C)
  loc_110C0C7B: var_58.DispID_0078(var_3C)
  loc_110C0CC3: var_58.DispID_0090(var_3C)
  loc_110C0D09: var_58.DispID_0077(var_3C)
  loc_110C0D4F: var_58.DispID_0078(var_3C)
  loc_110C0D97: var_58.DispID_0090(var_3C)
  loc_110C0DDD: var_58.DispID_0077(var_3C)
  loc_110C0E23: var_58.DispID_0078(var_3C)
  loc_110C0E6B: var_58.DispID_0090(var_3C)
  loc_110C0EB1: var_58.DispID_0077(var_3C)
  loc_110C0EF7: var_58.DispID_0078(var_3C)
  loc_110C0F13: If 10 <= &H14 Then
  loc_110C0F53:   var_58.DispID_00AC(var_3C)
  loc_110C0F6B:   var_14 = 1+var_14
  loc_110C0F6E:   GoTo loc_110C0F0F
  loc_110C0F70: End If
  loc_110C0FB0: var_58.DispID_00AC(var_3C)
  loc_110C0FF5: var_58.DispID_00AC(var_3C)
  loc_110C103A: var_58.DispID_00AC(var_3C)
End Sub

Private Sub Proc_16_10_110C7C80
  Dim var_7C As Variant
  Dim var_1F8 As Label
  Dim var_80 As Variant
  Dim var_88 As frmZGXSToPzTGZP.Label3
  loc_110C7D6A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110C7D72: var_228 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110C7D78: var_8004 = ecx
  loc_110C7DEE: If var_14 <= CLng(frmZGXSToPzTGZP.VFG.DispID_0007)(-1) Then
  loc_110C7DFF:   var_800C = frmZGXSToPzTGZP.Proc_16_11_110C9B20(vbNull)
  loc_110C7E9D:   frmZGXSToPzTGZP.VFG.DispID_0082(22, var_58)
  loc_110C7F81:   If (frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 22) = global_1100AE28) + 1 Then
  loc_110C8001:     frmZGXSToPzTGZP.VFG.DispID_0082(1, 285267764)
  loc_110C8135:     frmZGXSToPzTGZP.VFG.DispID_009E(var_14, 1, var_14, 1, 16711680)
  loc_110C8155:     Set var_7C = frmZGXSToPzTGZP.Label3
  loc_110C8162:     var_1F8 = var_7C
  loc_110C81AC:     var_7C.Caption = "分析: 第(" & CStr(vbNull) & ")行信息----有效"
  loc_110C81FE:     frmZGXSToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_110C8211:   Else
  loc_110C828B:     frmZGXSToPzTGZP.VFG.DispID_0082(1, 285267820)
  loc_110C83BF:     frmZGXSToPzTGZP.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110C83DF:     Set var_80 = frmZGXSToPzTGZP.Label3
  loc_110C83EC:     var_1F8 = var_80
  loc_110C84CD:     var_80.Caption = "分析:   第(" & CStr(vbNull) & ")行信息----" & frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 22)
  loc_110C8538:     frmZGXSToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_110C854A:   End If
  loc_110C855A:   var_14 = 1+var_14
  loc_110C855D:   GoTo loc_110C7DE0
  loc_110C8562: End If
  loc_110C85C9: If var_14 <= CLng(frmZGXSToPzTGZP.VFG.DispID_0007)(-1) Then
  loc_110C8641:   var_A0 = frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 2)
  loc_110C865F:   var_B8)
  loc_110C87EF:   var_8048 = frmZGXSToPzTGZP.VFG.DispID_0082(var_14, frmZGXSToPzTGZP.VFG)
  loc_110C8826:   var_4C = CCur(0)
  loc_110C8829:   var_48 = var_8048
  loc_110C8835:   var_40 = CCur(0)
  loc_110C8838:   var_3C = var_8048
  loc_110C8844:   var_34 = var_14
  loc_110C884D:   var_30 = var_14
  loc_110C8856:   var_160 = CByte("DateToPeriod".00000001h)
  loc_110C88F3:   var_B8)
  loc_110C8972:   Set var_80 = frmZGXSToPzTGZP.VFG
  loc_110C8998:   var_8064 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 3) = var_80.DispID_0082(var_14, 3))
  loc_110C89C5:   var_1A0 = var_8064 + 1
  loc_110C8A3F:   var_806C = (var_8048 = frmZGXSToPzTGZP.VFG.DispID_0082(var_14, ""))
  loc_110C8A66:   var_1E0 = var_806C + 1
  loc_110C8B68:   If CBool((frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_8064 + 1 And var_806C + 1) Then
  loc_110C8C2D:     If (frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 22) = global_1100AE28) Then
  loc_110C8C36:     End If
  loc_110C8C3B:     If var_24 = 0 Then
  loc_110C8CE4:       var_16C = var_48
  loc_110C8D28:       var_9C = var_1F0
  loc_110C8D74:       var_4C = CCur(var_4C + Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 7)), "#.00"))
  loc_110C8D77:       var_48 = var_D8
  loc_110C8E57:       var_16C = var_3C
  loc_110C8E9B:       var_9C = var_1F0
  loc_110C8EE7:       var_40 = CCur(var_40 + Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 8)), "#.00"))
  loc_110C8EEA:       var_3C = var_D8
  loc_110C8F2A:     End If
  loc_110C8F4B:     var_14 = var_14(1)
  loc_110C8F4E:     var_30 = var_30(1)
  loc_110C8F70:     var_80A0 = CLng(frmZGXSToPzTGZP.VFG.DispID_0007)
  loc_110C8F8B:     var_1F8 = (var_14 > 0)
  loc_110C8FAF:     If var_1F8 = 0 Then GoTo loc_110C8850
  loc_110C8FB5:   End If
  loc_110C8FBA:   If var_24 = 0 Then
  loc_110C8FCE:     Set var_7C = frmZGXSToPzTGZP.Chk
  loc_110C8FD9:     var_1F8 = var_7C
  loc_110C8FDF:     Set var_80 = var_7C(1)
  loc_110C900A:     var_200 = var_80
  loc_110C9010:     var_1EC = var_80.Value
  loc_110C9064:     If (var_1EC = 1) Then
  loc_110C9094:       If (Abs(var_4C - var_40) <> 0.01) >= 0 Then
  loc_110C909D:       End If
  loc_110C909D:     End If
  loc_110C90A2:     If var_24 Then
  loc_110C90A8:     End If
  loc_110C90C8:     var_1C = var_34
  loc_110C90CD:     If var_34 <= (var_30 - 1) Then
  loc_110C9191:       If (frmZGXSToPzTGZP.VFG.DispID_0082(var_1C, 22) = global_1100AE28) + 1 Then
  loc_110C9219:         frmZGXSToPzTGZP.VFG.DispID_0082(1, 285267820)
  loc_110C92AD:         frmZGXSToPzTGZP.VFG.DispID_0082(22, "凭证借贷不平衡或某分录有错误")
  loc_110C93E1:         frmZGXSToPzTGZP.VFG.DispID_009E(var_1C, 1, var_1C, 1, 255)
  loc_110C93F3:       End If
  loc_110C9403:       GoTo loc_110C90C2
  loc_110C9408:     End If
  loc_110C9419:     var_44 = var_44(1)
  loc_110C942A:     Set var_88 = frmZGXSToPzTGZP.Label3
  loc_110C945D:     var_1F8 = var_88
  loc_110C956A:     Set var_80 = frmZGXSToPzTGZP.VFG
  loc_110C9642:     var_80D4 = "分析: 第[" & frmZGXSToPzTGZP.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmZGXSToPzTGZP.VFG.DispID_0082(var_34, var_14)
  loc_110C9664:     var_78 = var_80D4 & "]号凭证借贷不平衡"
  loc_110C9678:     var_88.Caption = var_78
  loc_110C967F:     If var_78 < 0 Then
  loc_110C9685:       GoTo loc_110C9903
  loc_110C968A:     End If
  loc_110C969B:     var_20 = var_20(1)
  loc_110C96AC:     Set var_88 = frmZGXSToPzTGZP.Label3
  loc_110C96DF:     var_1F8 = var_88
  loc_110C97EC:     Set var_80 = frmZGXSToPzTGZP.VFG
  loc_110C98C4:     var_80F8 = "分析: 第[" & frmZGXSToPzTGZP.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmZGXSToPzTGZP.VFG.DispID_0082(var_34, frmZGXSToPzTGZP.VFG.DispID_0082(var_34, var_14))
  loc_110C98E6:     var_78 = var_80F8 & "]号凭证有效"
  loc_110C98FA:     var_88.Caption = var_78
  loc_110C9901:     If var_78 >= 0 Then GoTo loc_110C9912
  loc_110C9903:     ' Referenced from: 110C9685
  loc_110C990C:     var_78 = CheckObj(var_1F8, global_1100D574, 84)
  loc_110C9912:   End If
  loc_110C9994:   frmZGXSToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_110C99C5:   var_14 = 1+var_14(-1)
  loc_110C99C8:   GoTo loc_110C85C3
  loc_110C99CD: End If
  loc_110C99D2: If var_44 > 0 Then
  loc_110C99D9:   If var_20 > 0 Then
  loc_110C99F4:   Else
  loc_110C9A0D:   Else
  loc_110C9A17:     var_8108 = frmZGXSToPzTGZP.Proc_16_13_110D8D10(var_1EC)
  loc_110C9A25:     If var_1EC Then
  loc_110C9A40:     Else
  loc_110C9A48:       var_18 = ecx
  loc_110C9A51:       GoTo loc_110C9AEB
  loc_110C9AEA:       Exit Sub
  loc_110C9AEB:     End If
  loc_110C9AEB:   End If
  loc_110C9AEB: End If
  loc_110C9AEB: ' Referenced from: 110C9A51
End Sub

Private  Proc_16_11_110C9B20(arg_C) '110C9B20
  Dim var_58 As frmZGXSToPzTGZP.VFG
  Dim var_20 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA19-EB96-11D2-AF06000021009B21}()
  Dim var_18 As {3302AA41-EB96-11D2-AF06000021009B21}()
  Dim var_1C As {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_110C9C1C: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110C9C2C: var_210 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110C9D0B: If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 2) = global_1100AE28) + 1 Then
  loc_110C9D15:   var_24 = "制单日期为空"
  loc_110C9D26: Else
  loc_110C9DC1:   var_78 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 2)
  loc_110C9DFB:   If Proc_0_9_11028500(var_80, global_110CF14E, ) Then
  loc_110C9EA4:     var_78 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 2)
  loc_110C9EAE:     var_90)
  loc_110C9EC0:     var_48 = var_90
  loc_110C9EF2:     var_118 = var_48
  loc_110C9F00:     var_114 = var_44
  loc_110C9F34:     var_80 = "AccountOpen".0.0
  loc_110C9F65:     If (var_80 < var_80) Then
  loc_110C9F6F:       var_24 = "日期超前总账系统启用日期"
  loc_110C9F80:     Else
  loc_110C9F86:       var_154 = var_44
  loc_110C9F8C:       var_1A4 = var_44
  loc_110C9F98:       var_158 = var_48
  loc_110C9F9F:       var_1A8 = var_48
  loc_110CA04C:       var_80 = "AccountYMD".0.00000002h("AccountYMD".0, var_13C)
  loc_110CA146:       If CBool( Or ((global_110CF14E < var_80) > "AccountYMD".0.00000002h(var_180, var_18C))) Then
  loc_110CA150:         var_24 = "日期必须在当前会计年度内"
  loc_110CA161:       Else
  loc_110CA17E:         var_118 = var_48
  loc_110CA1D2:         var_80 = "DateToPeriod".00000001h - 1
  loc_110CA260:         If CBool("AccountYMD".0.00000001h) Then
  loc_110CA26A:           var_24 = "已结账月份不能制单"
  loc_110CA27B:         Else
  loc_110CA357:           If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 3) = global_1100AE28) + 1 Then
  loc_110CA361:             var_24 = "凭证类别字为空"
  loc_110CA372:           Else
  loc_110CA401:             var_8034 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 3)
  loc_110CA411:             var_80 = 8
  loc_110CA414:             var_78 = var_8034
  loc_110CA45B:             var_8038 = CBool(Not("pzlbCheck".00000001h(, fs:[00000000h], , global_110CF14E, global_110CF14E, var_74, var_8034, var_7C)))
  loc_110CA492:             If var_8038 Then
  loc_110CA49C:               var_24 = "凭证类别字非法"
  loc_110CA4AD:             Else
  loc_110CA584:               If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, var_128) = global_1100AE28) + 1 Then
  loc_110CA58E:                 var_24 = "业务号为空"
  loc_110CA59F:               Else
  loc_110CA629:                 var_8044 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, var_128)
  loc_110CA639:                 var_80 = 8
  loc_110CA63C:                 var_78 = var_8044
  loc_110CA67F:                 var_90 = "GenLen".00000001h(fs:[00000000h], , global_110CF14E, global_110CF14E, global_110CF14E, var_74, var_8044, var_7C)
  loc_110CA6C7:                 If (var_90 > 30) Then
  loc_110CA6D1:                   var_24 = "业务号超长"
  loc_110CA6E2:                 Else
  loc_110CA7C1:                   If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 5) = global_1100AE28) + 1 Then
  loc_110CA7CB:                     var_24 = "摘要为空"
  loc_110CA7DC:                   Else
  loc_110CA9AC:                     var_80 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 5)
  loc_110CAA25:                     If (((InStr(1, frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 5), "'", 0) > 0) Or (InStr(1, frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 5), "|", 0) > 0)) Or (InStr(1, var_80, """", 0) > 0)) Then
  loc_110CAA2F:                       var_24 = "摘要含有非法字符"
  loc_110CAA40:                     Else
  loc_110CAAD2:                       var_806C = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 5)
  loc_110CAAE2:                       var_80 = 8
  loc_110CAAE5:                       var_78 = var_806C
  loc_110CAB28:                       var_90 = "GenLen".00000001h(global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, var_74, var_806C, var_7C)
  loc_110CAB71:                       If (var_90 > 120) Then
  loc_110CAB7B:                         var_24 = "摘要超长"
  loc_110CAB8C:                       Else
  loc_110CAC69:                         If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 6) = global_1100AE28) + 1 Then
  loc_110CAC73:                           var_24 = "科目为空"
  loc_110CAC84:                         Else
  loc_110CAD13:                           var_807C = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 6)
  loc_110CAD23:                           var_80 = 8
  loc_110CAD26:                           var_78 = var_807C
  loc_110CADA6:                           var_40 = "kmCheck".00000002h(var_807C, var_150, var_15C)
  loc_110CADD8:                           var_8084 = (var_40 = global_1100AE28)
  loc_110CADE0:                           If var_8084 = 0 Then
  loc_110CADEA:                             var_24 = "科目非法"
  loc_110CADFB:                           Else
  loc_110CAE39:                             var_118 = arg_C
  loc_110CAEA0:                             frmZGXSToPzTGZP.VFG.DispID_0082(6, var_40)
  loc_110CAEBA:                             var_118 = var_40
  loc_110CAF0C:                             var_128 = var_20
  loc_110CAF5A:                             "kmCodeToProperties".00000002h
  loc_110CAF77:                             Set var_20 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110CAFA5:                             var_1F0 = var_20
  loc_110CAFAB:                             var_1D4 = var_20.UnkVCall_00000114h
  loc_110CAFD7:                             If var_1D4 = 0 Then
  loc_110CAFE1:                               var_24 = "科目非末级"
  loc_110CAFF2:                             Else
  loc_110CB0D0:                               If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 7) = global_1100AE28) Then
  loc_110CB1AC:                                 If Not (IsNumeric(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 7))) Then
  loc_110CB1B6:                                   var_24 = "借方金额非法"
  loc_110CB1C7:                                 Else
  loc_110CB270:                                   var_80A4 = CDbl(Val(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 7)))
  loc_110CB30B:                                   var_80 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 7)
  loc_110CB333:                                   var_228 = CDbl(Val(var_80))
  loc_110CB349:                                   var_80B0 = CDbl(-9999999999999.99)
  loc_110CB361:                                   GoTo loc_110CB365
  loc_110CB3B3:                                   If (eax Or 0) Then
  loc_110CB3BD:                                     var_24 = "借方金额超范围"
  loc_110CB3CE:                                   Else
  loc_110CB3CE:                                   End If
  loc_110CB4AC:                                   If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 8) = global_1100AE28) Then
  loc_110CB588:                                     If Not (IsNumeric(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 8))) Then
  loc_110CB592:                                       var_24 = "贷方金额非法"
  loc_110CB5A3:                                     Else
  loc_110CB64C:                                       var_80C8 = CDbl(Val(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 8)))
  loc_110CB6E7:                                       var_80 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 8)
  loc_110CB70F:                                       var_234 = CDbl(Val(var_80))
  loc_110CB725:                                       var_80D4 = CDbl(-9999999999999.99)
  loc_110CB73D:                                       GoTo loc_110CB741
  loc_110CB78F:                                       If (eax Or 0) Then
  loc_110CB799:                                         var_24 = "贷方金额超范围"
  loc_110CB7AA:                                       Else
  loc_110CB7AA:                                       End If
  loc_110CB922:                                       var_74 = var_1E0
  loc_110CB994:                                       var_C4 = var_1E8
  loc_110CBA0E:                                       var_80E8 = (Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 7)), "#.00") <> 0) And (Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 8)), "#.00") <> 0)
  loc_110CBA87:                                       If CBool(var_80E8) Then
  loc_110CBA91:                                         var_24 = "借方金额和贷方金额不能同时不为0"
  loc_110CBAA2:                                       Else
  loc_110CBC1A:                                         var_74 = var_1E0
  loc_110CBC8C:                                         var_C4 = var_1E8
  loc_110CBD06:                                         var_8100 = (Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 7)), "#.00") = 0) And (Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 8)), "#.00") = 0)
  loc_110CBD7F:                                         If CBool(var_8100) Then
  loc_110CBD89:                                           var_24 = "借方金额和贷方金额不能同时为0"
  loc_110CBD9A:                                         Else
  loc_110CBDBA:                                           var_1F0 = var_20
  loc_110CBE0C:                                           If (var_20.UnkVCall_0000007Ch = global_1100AE28) Then
  loc_110CBEC9:                                             var_1F0 = (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 9) = global_1100AE28)
  loc_110CBEF0:                                             If var_1F0 = 0 Then GoTo loc_110CC0A6
  loc_110CBFA4:                                             var_1F0 = Not (IsNumeric(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 9)))
  loc_110CBFCC:                                             If var_1F0 = 0 Then GoTo loc_110CC0A6
  loc_110CBFDA:                                             var_24 = "数量数值非法"
  loc_110CBFEB:                                           Else
  loc_110CC008:                                             var_118 = arg_C
  loc_110CC094:                                             frmZGXSToPzTGZP.VFG.DispID_0082(9, 285257256)
  loc_110CC0C6:                                             var_1F0 = var_20
  loc_110CC118:                                             If (var_20.UnkVCall_0000006Ch = global_1100AE28) Then
  loc_110CC1FC:                                               If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 10) = global_1100AE28) Then
  loc_110CC2D8:                                                 If Not (IsNumeric(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 10))) Then
  loc_110CC2E2:                                                   var_24 = "外币金额非法"
  loc_110CC2F3:                                                 Else
  loc_110CC39C:                                                   var_813C = CDbl(Val(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 10)))
  loc_110CC45F:                                                   var_240 = CDbl(Val(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 10)))
  loc_110CC475:                                                   var_8148 = CDbl(-9999999999999.99)
  loc_110CC48D:                                                   GoTo loc_110CC491
  loc_110CC4DF:                                                   If (eax Or 0) Then
  loc_110CC4E9:                                                     var_24 = "外币超范围"
  loc_110CC4FA:                                                   Else
  loc_110CC4FA:                                                   End If
  loc_110CC5D8:                                                   If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 11) = global_1100AE28) Then
  loc_110CC6B4:                                                     If Not (IsNumeric(frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 11))) Then
  loc_110CC6BE:                                                       var_24 = "汇率数值非法"
  loc_110CC6CF:                                                     Else
  loc_110CC6CF:                                                     End If
  loc_110CC6CF:                                                   End If
  loc_110CC7AD:                                                   If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 12) = global_1100AE28) Then
  loc_110CC844:                                                     var_8164 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 12)
  loc_110CC857:                                                     var_78 = var_8164
  loc_110CC89A:                                                     var_90 = "GenLen".00000001h(global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, var_74, var_8164, var_7C)
  loc_110CC8B4:                                                     var_1F0 = (var_90 > 20)
  loc_110CC8E3:                                                     If var_1F0 = 0 Then GoTo loc_110CCA29
  loc_110CC8F1:                                                     var_24 = "制单人姓名超长"
  loc_110CC902:                                                   Else
  loc_110CC921:                                                     var_118 = arg_C
  loc_110CC9FD:                                                     frmZGXSToPzTGZP.VFG.DispID_0082(12, "UserCurrent".00000000h.00000000h)
  loc_110CCA4C:                                                     var_1F0 = var_20
  loc_110CCA7E:                                                     If var_20.UnkVCall_0000010Ch Then
  loc_110CCB62:                                                       If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 13) = global_1100AE28) Then
  loc_110CCBF9:                                                         var_817C = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 13)
  loc_110CCC0C:                                                         var_78 = var_817C
  loc_110CCC3B:                                                         var_90 = "JsfsCheck".00000001h(1, global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, var_74, var_817C, var_7C)
  loc_110CCC8B:                                                         If CBool(Not(var_90)) Then
  loc_110CCC95:                                                           var_24 = "结算方式非法"
  loc_110CCCA6:                                                         Else
  loc_110CCCA6:                                                         End If
  loc_110CCCA6:                                                       End If
  loc_110CCCC9:                                                       var_1F0 = var_20
  loc_110CCCCF:                                                       var_1D4 = var_20.UnkVCall_0000010Ch
  loc_110CCD16:                                                       var_1F8 = var_20
  loc_110CCD1C:                                                       var_1D8 = var_20.UnkVCall_00000094h
  loc_110CCD63:                                                       var_200 = var_20
  loc_110CCDBB:                                                       If (var_20.UnkVCall_0000009Ch = 0) = 0 Then
  loc_110CCE9F:                                                         If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 14) = global_1100AE28) Then
  loc_110CCF36:                                                           var_8198 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 14)
  loc_110CCF49:                                                           var_78 = var_8198
  loc_110CCF8C:                                                           var_90 = "GenLen".00000001h(1, global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, var_74, var_8198, var_7C)
  loc_110CCFD5:                                                           If (var_90 > 10) Then
  loc_110CCFDF:                                                             var_24 = "票号超长"
  loc_110CCFF0:                                                           Else
  loc_110CCFF0:                                                           End If
  loc_110CD0CE:                                                           If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 15) = global_1100AE28) Then
  loc_110CD165:                                                             var_81A8 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 15)
  loc_110CD178:                                                             var_78 = var_81A8
  loc_110CD1A7:                                                             var_90 = "DateCheck".00000001h(1, global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, var_74, var_81A8, var_7C)
  loc_110CD1F7:                                                             If CBool(Not(var_90)) Then
  loc_110CD201:                                                               var_24 = "票号发生日期非法"
  loc_110CD212:                                                             Else
  loc_110CD212:                                                             End If
  loc_110CD212:                                                           End If
  loc_110CD235:                                                           var_1F0 = var_20
  loc_110CD282:                                                           var_1F8 = var_20
  loc_110CD288:                                                           var_1D8 = var_20.UnkVCall_0000008Ch
  loc_110CD2EB:                                                           If (var_20.UnkVCall_000000A4h = 0) = 0 Then
  loc_110CD3AA:                                                             If (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 16) = global_1100AE28) Then
  loc_110CD454:                                                               var_78 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 16)
  loc_110CD4D4:                                                               var_38 = "BmCheck".00000002h(var_154, 0, var_15C)
  loc_110CD506:                                                               var_81C8 = (var_38 = global_1100AE28)
  loc_110CD50E:                                                               If var_81C8 = 0 Then
  loc_110CD518:                                                                 var_24 = "部门非法"
  loc_110CD529:                                                               Else
  loc_110CD546:                                                                 var_118 = arg_C
  loc_110CD5D0:                                                                 frmZGXSToPzTGZP.VFG.DispID_0082(16, var_38)
  loc_110CD605:                                                                 var_1F0 = var_20
  loc_110CD637:                                                                 If var_20.UnkVCall_000000A4h Then
  loc_110CD645:                                                                   var_118 = var_38
  loc_110CD697:                                                                   var_128 = var_28
  loc_110CD6E5:                                                                   "BmToProperties".00000002h
  loc_110CD702:                                                                   Set var_28 = {3302AA19-EB96-11D2-AF06000021009B21}()
  loc_110CD730:                                                                   var_1F0 = var_28
  loc_110CD736:                                                                   var_1D4 = var_28.UnkVCall_00000034h
  loc_110CD75C:                                                                   If var_1D4 = 0 Then
  loc_110CD76A:                                                                     var_24 = "部门非末级"
  loc_110CD77B:                                                                   Else
  loc_110CD783:                                                                     var_24 = "部门为空"
  loc_110CD794:                                                                   Else
  loc_110CD816:                                                                     frmZGXSToPzTGZP.VFG.DispID_0082(var_128, 285257256)
  loc_110CD828:                                                                   End If
  loc_110CD828:                                                                 End If
  loc_110CD84B:                                                                 var_1F0 = var_20
  loc_110CD87D:                                                                 If var_20.UnkVCall_0000008Ch Then
  loc_110CD929:                                                                   var_81E0 = (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H11) = global_1100AE28)
  loc_110CD961:                                                                   If var_81E0 Then
  loc_110CDA0D:                                                                     var_81E8 = (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, 16) = global_1100AE28)
  loc_110CDA69:                                                                     If var_81E8 + 1 Then
  loc_110CDAEC:                                                                       var_78 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H11)
  loc_110CDB7A:                                                                       var_90 = "ZyCheck".00000003h(var_174, "BmCheck".00000002h(var_154, 80020004h, var_15C), var_17C)
  loc_110CDB8F:                                                                       var_34 = var_90
  loc_110CDBC1:                                                                       var_81F4 = (var_34 = global_1100AE28)
  loc_110CDBC9:                                                                       If var_81F4 = 0 Then
  loc_110CDBD3:                                                                         var_24 = "职员非法"
  loc_110CDBE4:                                                                       Else
  loc_110CDC01:                                                                         var_118 = arg_C
  loc_110CDC8B:                                                                         frmZGXSToPzTGZP.VFG.DispID_0082(&H11, var_34)
  loc_110CDCAA:                                                                         var_118 = var_34
  loc_110CDCF7:                                                                         var_128 = var_18
  loc_110CDD45:                                                                         "ZyToProperties".00000002h
  loc_110CDD62:                                                                         Set var_18 = {3302AA41-EB96-11D2-AF06000021009B21}()
  loc_110CDD70:                                                                         var_118 = arg_C
  loc_110CDDB1:                                                                         var_1F0 = var_18
  loc_110CDE6A:                                                                         frmZGXSToPzTGZP.VFG.DispID_0082(var_128, var_18.UnkVCall_0000002Ch)
  loc_110CDE8A:                                                                       Else
  loc_110CDF00:                                                                         var_158 = var_38
  loc_110CDF0D:                                                                         var_78 = frmZGXSToPzTGZP.VFG.DispID_0082(8, var_128)
  loc_110CDFC2:                                                                         var_34 = "ZyCheck".00000003h(var_164, 0, var_16C)
  loc_110CDFF4:                                                                         var_8208 = (var_34 = global_1100AE28)
  loc_110CDFFC:                                                                         If var_8208 = 0 Then
  loc_110CE006:                                                                           var_24 = "职员不在指定部门内"
  loc_110CE017:                                                                         Else
  loc_110CE055:                                                                           var_118 = arg_C
  loc_110CE0BC:                                                                           frmZGXSToPzTGZP.VFG.DispID_0082(&H11, var_34)
  loc_110CE0CE:                                                                         End If
  loc_110CE0CE:                                                                       End If
  loc_110CE0CE:                                                                     End If
  loc_110CE0F1:                                                                     var_1F0 = var_20
  loc_110CE123:                                                                     If var_20.UnkVCall_00000094h Then
  loc_110CE1CF:                                                                       var_8214 = (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H12) = global_1100AE28)
  loc_110CE1E0:                                                                       var_1F0 = var_8214
  loc_110CE207:                                                                       If var_1F0 = 0 Then GoTo loc_110CE6F5
  loc_110CE2B1:                                                                       var_78 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H12)
  loc_110CE331:                                                                       var_3C = "KhCheck".00000002h(var_154, 0, var_15C)
  loc_110CE363:                                                                       var_8220 = (var_3C = global_1100AE28)
  loc_110CE36B:                                                                       If var_8220 = 0 Then
  loc_110CE375:                                                                         var_24 = "客户非法"
  loc_110CE386:                                                                       Else
  loc_110CE3C4:                                                                         var_118 = arg_C
  loc_110CE42B:                                                                         frmZGXSToPzTGZP.VFG.DispID_0082(&H12, var_3C)
  loc_110CE43D:                                                                       End If
  loc_110CE460:                                                                       var_1F0 = var_20
  loc_110CE492:                                                                       If var_20.UnkVCall_0000009Ch Then
  loc_110CE53E:                                                                         var_822C = (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H13) = global_1100AE28)
  loc_110CE54F:                                                                         var_1F0 = var_822C
  loc_110CE576:                                                                         If var_1F0 = 0 Then GoTo loc_110CEAAE
  loc_110CE620:                                                                         var_78 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H13)
  loc_110CE6A0:                                                                         var_30 = "GysCheck".00000002h(var_154, 0, var_15C)
  loc_110CE6D2:                                                                         var_8238 = (var_30 = global_1100AE28)
  loc_110CE6DA:                                                                         If var_8238 = 0 Then
  loc_110CE6E4:                                                                           var_24 = "供应商非法"
  loc_110CE6F0:                                                                           GoTo loc_110CF10F
  loc_110CE6FD:                                                                           var_24 = "客户为空"
  loc_110CE70E:                                                                         Else
  loc_110CE74C:                                                                           var_118 = arg_C
  loc_110CE7B3:                                                                           frmZGXSToPzTGZP.VFG.DispID_0082(&H13, var_30)
  loc_110CE7C5:                                                                         End If
  loc_110CE7E8:                                                                         var_1F0 = var_20
  loc_110CE835:                                                                         var_1F8 = var_20
  loc_110CE83B:                                                                         var_1D8 = var_20.UnkVCall_0000009Ch
  loc_110CE879:                                                                         If (var_20.UnkVCall_00000094h = 0) = 0 Then
  loc_110CE925:                                                                           var_8248 = (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H14) = global_1100AE28)
  loc_110CE95D:                                                                           If var_8248 Then
  loc_110CE9F4:                                                                             var_824C = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H14)
  loc_110CEA07:                                                                             var_78 = var_824C
  loc_110CEA4A:                                                                             var_90 = "GenLen".00000001h(global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, global_110CF14E, var_74, var_824C, var_7C)
  loc_110CEA93:                                                                             If (var_90 > 20) Then
  loc_110CEA9D:                                                                               var_24 = "业务员超长"
  loc_110CEAA9:                                                                               GoTo loc_110CF10F
  loc_110CEAB6:                                                                               var_24 = "供应商为空"
  loc_110CEAC7:                                                                             Else
  loc_110CEAC7:                                                                             End If
  loc_110CEAC7:                                                                           End If
  loc_110CEAE7:                                                                           var_1F0 = var_20
  loc_110CEB50:                                                                           If (var_20.UnkVCall_000000ACh = global_1100AE28) Then
  loc_110CEBEB:                                                                             var_8260 = (frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H15) = global_1100AE28)
  loc_110CEBFC:                                                                             var_1F0 = var_8260
  loc_110CEC23:                                                                             If var_1F0 = 0 Then GoTo loc_110CF07A
  loc_110CEC49:                                                                             var_1F0 = var_20
  loc_110CEC79:                                                                             var_78 = var_20.UnkVCall_000000ACh
  loc_110CED27:                                                                             var_88 = frmZGXSToPzTGZP.VFG.DispID_0082(arg_C, &H15)
  loc_110CEDC4:                                                                             var_2C = "XmCheck".00000003h(var_164, Not(8), var_16C)
  loc_110CEDFD:                                                                             var_8270 = (var_2C = global_1100AE28)
  loc_110CEE05:                                                                             If var_8270 = 0 Then
  loc_110CEE0F:                                                                               var_24 = "项目为空或非法"
  loc_110CEE20:                                                                             Else
  loc_110CEE4C:                                                                               var_4C = var_20.UnkVCall_000000ACh
  loc_110CEE7A:                                                                               var_128 = var_2C
  loc_110CEEAB:                                                                               Set var_58 = var_1C
  loc_110CEF2D:                                                                               "XmToProperties".00000003h
  loc_110CEF7F:                                                                               var_1D4 = {3302AA4B-EB96-11D2-AF06000021009B21}().UnkVCall_00000034h
  loc_110CEF9F:                                                                               If var_1D4 Then
  loc_110CEFAD:                                                                                 var_24 = "项目已结算"
  loc_110CEFBE:                                                                               Else
  loc_110CF04F:                                                                                 frmZGXSToPzTGZP.VFG.DispID_0082(&H15, 285257256)
  loc_110CF06C:                                                                               Else
  loc_110CF074:                                                                                 var_24 = "制单日期非法"
  loc_110CF07A:                                                                               End If
  loc_110CF080:                                                                               GoTo loc_110CF10F
  loc_110CF089:                                                                               If var_4 Then
  loc_110CF094:                                                                               End If
  loc_110CF10E:                                                                               Exit Sub
  loc_110CF10F:                                                                             End If
  loc_110CF10F:                                                                           End If
  loc_110CF10F:                                                                         End If
  loc_110CF10F:                                                                       End If
  loc_110CF10F:                                                                     End If
  loc_110CF10F:                                                                   End If
  loc_110CF10F:                                                                 End If
  loc_110CF10F:                                                               End If
  loc_110CF10F:                                                             End If
  loc_110CF10F:                                                           End If
  loc_110CF10F:                                                         End If
  loc_110CF10F:                                                       End If
  loc_110CF10F:                                                     End If
  loc_110CF10F:                                                   End If
  loc_110CF10F:                                                 End If
  loc_110CF10F:                                               End If
  loc_110CF10F:                                             End If
  loc_110CF10F:                                           End If
  loc_110CF10F:                                         End If
  loc_110CF10F:                                       End If
  loc_110CF10F:                                     End If
  loc_110CF10F:                                   End If
  loc_110CF10F:                                 End If
  loc_110CF10F:                               End If
  loc_110CF10F:                             End If
  loc_110CF10F:                           End If
  loc_110CF10F:                         End If
  loc_110CF10F:                       End If
  loc_110CF10F:                     End If
  loc_110CF10F:                   End If
  loc_110CF10F:                 End If
  loc_110CF10F:               End If
  loc_110CF10F:             End If
  loc_110CF10F:           End If
  loc_110CF10F:         End If
  loc_110CF10F:       End If
  loc_110CF10F:     End If
  loc_110CF10F:   End If
  loc_110CF10F: End If
  loc_110CF10F: ' Referenced from: 110CF080
End Sub

Private Sub Proc_16_12_110CF170
  Dim var_9C As Variant
  Dim var_8034 As Label
  Dim var_8074 As Label
  Dim var_A0 As Variant
  Dim var_38 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA47-EB96-11D2-AF06000021009B21}()
  loc_110CF2CA: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110CF2D0: var_294 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110CF2F6: Set var_9C = frmZGXSToPzTGZP.VFG
  loc_110CF340: If (CLng(var_9C.DispID_0007) < 2) Then
  loc_110CF36E:   var_800C = = Global.Screen
  loc_110CF390:   var_8010 = ecx
  loc_110CF398:   var_8010 = var_9C.UnkVCall_0000007Ch
  loc_110CF405:   var_C8 = "提示信息"
  loc_110CF407:   var_150 = "没有可生成用友凭证的数据。"
  loc_110CF416: Else
  loc_110CF4C6:   var_264 = ("GetAccInfo".00000002h(, , fs:[00000000h], , "GL", var_16C, "dGLStartDate", var_174) = 1100AE28h)
  loc_110CF4E0:   If var_264 = 0 Then GoTo loc_110CF621
  loc_110CF50E:   var_801C = = Global.Screen
  loc_110CF530:   var_8020 = ecx
  loc_110CF538:   var_8020 = var_9C.UnkVCall_0000007Ch
  loc_110CF5A5:   var_C8 = "提示信息"
  loc_110CF5A7:   var_150 = "总账系统尚未启用，不能进行凭证引入！"
  loc_110CF5B1: End If
  loc_110CF5E3: MsgBox(var_150, 64, var_C8, var_D8, var_E8)
  loc_110CF610: Exit Sub
  loc_110CF61C: GoTo loc_110D824E
  loc_110CF62B: var_8024 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[VouchNum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " & " CREATE TABLE VouchNum(iperiod tinyint NULL ,csign varchar(8) NULL ,ino_id int NULL,constraint index1 unique(iperiod,csign,ino_id))"
  loc_110CF631: var_B0 = var_8024
  loc_110CF690: var_D8.00000001h(0, , , , "3缷M鋎?", var_AC, var_8024, var_B4)
  loc_110CF6B0: On Error GoTo 0
  loc_110CF6B6: var_B0 = %ecx = %S_edx_S
  loc_110CF6D8: var_78 = "AS13"
  loc_110CF6F0: var_78)
  loc_110CF71A: If Not (var_78)) Then
  loc_110CF74B:   If Global.Screen < 0 Then
  loc_110CF75C:   End If
  loc_110CF766:   var_8030 = ecx
  loc_110CF775:   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110CF784:     var_8030 = CheckObj(var_9C, global_1100C47C, 124)
  loc_110CF78F:   End If
  loc_110CF7A0:   call var_8034 = var_9C(var_9C, frmZGXSToPzTGZP.Label3)
  loc_110CF7A2:   var_264 = var_8034
  loc_110CF7B0:   Label3.Caption = "正在进行数据分析，请稍等..."
  loc_110CF7DD:   var_150 = True
  loc_110CF820:   call var_8038 = var_9C(var_9C, frmZGXSToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, True, var_14C)
  loc_110CF823:   var_8038.DispID_0000 =
  loc_110CF84C:   call var_803C = var_9C(var_9C, frmZGXSToPzTGZP.Pic1, global_FFFFFDDA, var_9C = var_9C)
  loc_110CF84F:   var_803C.DispID_0000
  loc_110CF86E:   var_8040 = .Proc_16_10_110C7C80(var_24C)
  loc_110CF87C:   If var_24C = 2 Then
  loc_110CF882:     var_150 = %ecx = %S_edx_S
  loc_110CF8C5:     call var_8044 = var_9C(var_9C, frmZGXSToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C)
  loc_110CF8C8:     var_8044.DispID_0000 =
  loc_110CF964:     MsgBox("数据源中没有合法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110CF9A1:     var_24C = %ecx = %S_edx_S
  loc_110CF9C7:     "AS13")
  loc_110CFA09:     var_B8 = Global.Screen
  loc_110CFA2B:     var_804C = ecx
  loc_110CFA3A:     If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110CFA49:       var_804C = CheckObj(var_9C, global_1100C47C, 124)
  loc_110CFA54:     End If
  loc_110CFA56:     If var_804C = 1 Then
  loc_110CFA5C:       var_150 = %ecx = %S_edx_S
  loc_110CFA9F:       call var_8050 = var_9C(var_9C, frmZGXSToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, var_14C = var_9C, var_14C)
  loc_110CFAA2:       var_8050.DispID_0000 =
  loc_110CFB3E:       MsgBox("数据源中含有非法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110CFB7B:       var_24C = %ecx = %S_edx_S
  loc_110CFBA1:       "AS13")
  loc_110CFBE3:       var_B8 = Global.Screen
  loc_110CFC05:       var_8058 = ecx
  loc_110CFC14:       If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110CFC23:         var_8058 = CheckObj(var_9C, global_1100C47C, 124)
  loc_110CFC2E:       End If
  loc_110CFC30:       If var_8058 = 3 Then
  loc_110CFC36:         var_150 = %ecx = %S_edx_S
  loc_110CFC79:         call var_805C = var_9C(var_9C, frmZGXSToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C)
  loc_110CFC7C:         var_805C.DispID_0000 =
  loc_110CFD18:         MsgBox("数据源中指定的凭证号无效或重号，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110CFD55:         var_24C = %ecx = %S_edx_S
  loc_110CFD7B:         "AS13")
  loc_110CFDBD:         var_B8 = Global.Screen
  loc_110CFDDF:         var_8064 = ecx
  loc_110CFDEE:         If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110CFDFD:           var_8064 = CheckObj(var_9C, global_1100C47C, 124)
  loc_110CFE08:         End If
  loc_110CFE4A:         var_C8 = "提示信息"
  loc_110CFE70:         var_B8 = "数据源中的数据已全部通过检查，是否开始引入？"
  loc_110CFE94:         MsgBox(var_B8, 36, var_C8, var_D8, var_E8)
  loc_110CFED9:         If (MsgBox(var_B8, 36, var_C8, var_D8, var_E8) = 7) Then
  loc_110CFF24:           call var_8068 = var_9C(var_9C, frmZGXSToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, frmZGXSToPzTGZP.Pic1, var_14C)
  loc_110CFF27:           var_8068.DispID_0000 =
  loc_110CFF4D:           var_24C = %ecx = %S_edx_S
  loc_110CFF73:           "AS13")
  loc_110CFFB5:           var_B8 = Global.Screen
  loc_110CFFD7:           var_8070 = ecx
  loc_110CFFE6:           If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110CFFF5:             var_8070 = CheckObj(var_9C, global_1100C47C, 124)
  loc_110D0000:           End If
  loc_110D0001:           On Error GoTo 0
  loc_110D0018:           call var_8074 = var_9C(var_9C, frmZGXSToPzTGZP.Label3, var_9C = var_9C)
  loc_110D001A:           var_264 = var_8074
  loc_110D0028:           Label3.Caption = "正在写数据，请稍等..."
  loc_110D006C:           call var_8078 = var_9C(var_9C, frmZGXSToPzTGZP.Pic1, global_FFFFFDDA, 00000000h)
  loc_110D006F:           var_8078.DispID_0000
  loc_110D00A6:           Set var_74 = CreateObject("UfDbKit.UfRecordset", 0)
  loc_110D00BD:           var_150 = "SELECT TOP 1 * FROM GL_accvouch"
  loc_110D0132:           Set var_74 = "DataMdb".00000000h.00000001h(var_14C, "SELECT TOP 1 * FROM GL_accvouch", var_154)
  loc_110D0166:           call var_8084 = var_9C(var_9C, frmZGXSToPzTGZP.VFG, 00000007h, 00000000h)
  loc_110D01CA:           If var_24 <= CLng(var_8084.DispID_0000)(-1) Then
  loc_110D01D4:             var_2A8 = var_24
  loc_110D01DA:             var_150 = var_24
  loc_110D0257:             call var_8090 = var_9C(var_9C, frmZGXSToPzTGZP.VFG, 00000082h, 00000002h, 3, var_174, 2, var_16C, 00000003h, var_154, var_24, var_14C)
  loc_110D0271:             var_C0 = var_8090.DispID_0000
  loc_110D028F:             var_D8)
  loc_110D02E7:             var_70 = CByte("DateToPeriod".00000001h(8, var_D4))
  loc_110D0320:             var_150 = var_2A8
  loc_110D0399:             call var_809C = var_9C(var_9C, frmZGXSToPzTGZP.VFG, 00000082h, 00000002h, 3, var_174, 3, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110D03B8:             var_58 = var_809C.DispID_0000
  loc_110D03DC:             var_150 = var_2A8
  loc_110D0459:             call var_80A4 = var_9C(var_9C, frmZGXSToPzTGZP.VFG, 00000082h, 00000002h, 3, var_174, 0, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110D0478:             var_64 = var_80A4.DispID_0000
  loc_110D049C:             var_150 = var_2A8
  loc_110D0519:             call var_80AC = var_9C(var_9C, frmZGXSToPzTGZP.VFG, 00000082h, 00000002h, 3, var_174, 1, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110D0582:             If (var_80AC.DispID_0000 = global_1100D76C) Then
  loc_110D0599:               call var_80B8 = var_9C(var_A8, frmZGXSToPzTGZP.Label3)
  loc_110D059B:               var_264 = var_80B8
  loc_110D06AB:               var_80 = "正在处理：第[" & frmZGXSToPzTGZP.VFG.DispID_0082(var_2A8, 2) & " - "
  loc_110D07EC:               var_D8 = frmZGXSToPzTGZP.VFG.DispID_0082(var_2A8, 0)
  loc_110D0833:               var_98 = var_80 & frmZGXSToPzTGZP.VFG.DispID_0082(var_2A8, 3) & " - " & var_D8 & "]号凭证"
  loc_110D0843:               var_98 = var_80B8.UnkVCall_00000054h
  loc_110D08FE:               frmZGXSToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_110D0932:               var_3C = var_24
  loc_110D0946:               Set var_9C = frmZGXSToPzTGZP.Chk
  loc_110D0948:               var_264 = var_9C
  loc_110D095A:               Set var_A0 = var_9C(0)
  loc_110D097E:               var_26C = var_A0
  loc_110D09E8:               If (var_A0.Value = 1) Then
  loc_110D0A1B:                 var_24C = CInt("cIYear".00000000h)
  loc_110D0A30:                 var_24C, var_70)
  loc_110D0A3D:                 var_54 = var_24C, var_70)
  loc_110D0A4E:               Else
  loc_110D0A64:                 var_80E8 = .Proc_16_14_110D9B20(var_70)
  loc_110D0A76:                 var_54 = var_258
  loc_110D0A79:               End If
  loc_110D0A7E:               If var_54 > 0 Then
  loc_110D0A86:                 On Error GoTo loc_110D63F0
  loc_110D0ABF:                 "wksAlias".00000000h.00000000h(var_58)
  loc_110D0ADE:                 var_1A0 = var_70
  loc_110D0BA7:                 var_D8)
  loc_110D0C53:                 var_80FC = (var_58 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 3))
  loc_110D0C60:                 var_1F0 = var_80FC + 1
  loc_110D0D1C:                 var_8104 = (var_64 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 0))
  loc_110D0D29:                 var_240 = var_8104 + 1
  loc_110D0DC6:                 var_8114 = CBool((frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_80FC + 1 And var_8104 + 1)
  loc_110D0E4B:                 If var_8114 Then
  loc_110D0EEC:                   var_C0 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 6)
  loc_110D0F29:                   var_1A0 = var_38
  loc_110D0F97:                   "kmCodeToProperties".00000002h
  loc_110D0FB7:                   Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110D0FF0:                   var_74.AddNew
  loc_110D0FFB:                   var_150 = "ibook"
  loc_110D106C:                   var_74.DispID_0000(0)
  loc_110D106E:                   var_1A0 = "iPeriod"
  loc_110D111D:                   var_C0 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 2)
  loc_110D113B:                   var_D8)
  loc_110D11D4:                   var_74.DispID_0000("DateToPeriod".00000001h)
  loc_110D1209:                   var_190 = "csign"
  loc_110D1316:                   var_74.DispID_0000(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 3))
  loc_110D133D:                   var_190 = "isignseq"
  loc_110D145D:                   var_74.DispID_0000(Proc_0_4_11026BD0(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 3), var_64, var_258))
  loc_110D1488:                   var_150 = "ino_id"
  loc_110D14FA:                   var_74.DispID_0000(var_54)
  loc_110D14FC:                   var_190 = "dbill_date"
  loc_110D15AB:                   var_C0 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 2)
  loc_110D15C9:                   var_D8)
  loc_110D1626:                   var_74.DispID_0000(var_D8)
  loc_110D1654:                   var_190 = "idoc"
  loc_110D166C:                   var_150 = var_24
  loc_110D1775:                   var_74.DispID_0000(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_150, 4)))
  loc_110D17A0:                   var_160 = "ctext1"
  loc_110D1807:                   var_74.DispID_0000(var_150)
  loc_110D180E:                   var_160 = "ctext2"
  loc_110D1875:                   var_74.DispID_0000(var_150)
  loc_110D187C:                   var_150 = "cbill"
  loc_110D18EA:                   var_74.DispID_0000("cUserName".00000000h(, var_14C, "cbill", var_154))
  loc_110D1900:                   var_160 = "cbook"
  loc_110D1967:                   var_74.DispID_0000(var_150)
  loc_110D196E:                   var_160 = "ccheck"
  loc_110D19D5:                   var_74.DispID_0000(var_150)
  loc_110D19DC:                   var_160 = "ccashier"
  loc_110D1A43:                   var_74.DispID_0000(var_150)
  loc_110D1A4A:                   var_160 = "iflag"
  loc_110D1AB1:                   var_74.DispID_0000(var_150)
  loc_110D1AB8:                   var_160 = "coutaccset"
  loc_110D1B1F:                   var_74.DispID_0000(var_150)
  loc_110D1B26:                   var_160 = "ioutyear"
  loc_110D1B8D:                   var_74.DispID_0000(var_150)
  loc_110D1B94:                   var_160 = "coutsysver"
  loc_110D1BFB:                   var_74.DispID_0000(var_150)
  loc_110D1C02:                   var_160 = "coutsysname"
  loc_110D1C69:                   var_74.DispID_0000(var_150)
  loc_110D1C70:                   var_170 = "ioutperiod"
  loc_110D1D0D:                   var_74.DispID_0000(var_74.DispID_0000("iPeriod"))
  loc_110D1D1E:                   var_170 = "doutbilldate"
  loc_110D1DE1:                   var_74.DispID_0000(CStr(var_74.DispID_0000("dbill_date")))
  loc_110D1DFE:                   var_150 = "iYear"
  loc_110D1E6C:                   var_74.DispID_0000("cIYear".00000000h(var_58, var_14C, "iYear", var_154))
  loc_110D1F6A:                   var_74.DispID_0000("cIYear".00000000h(, var_16C, "iYPeriod", var_174) & Format(var_70, "00"))
  loc_110D1F98:                   var_160 = "coutsign"
  loc_110D1FFF:                   var_74.DispID_0000(var_70)
  loc_110D2001:                   var_190 = "coutno_id"
  loc_110D210E:                   var_74.DispID_0000(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 3))
  loc_110D213A:                   var_150 = "bvouchedit"
  loc_110D21A9:                   var_74.DispID_0000(FFFFFFFFh)
  loc_110D21B0:                   var_150 = "bvouchaddordele"
  loc_110D2221:                   var_74.DispID_0000(FFFFFFFFh)
  loc_110D2228:                   var_150 = "bvouchmoneyhold"
  loc_110D2299:                   var_74.DispID_0000(FFFFFFFFh)
  loc_110D22A0:                   var_150 = "bvalueedit"
  loc_110D2311:                   var_74.DispID_0000(FFFFFFFFh)
  loc_110D2318:                   var_150 = "bcodeedit"
  loc_110D2389:                   var_74.DispID_0000(FFFFFFFFh)
  loc_110D2390:                   var_150 = "bPCSedit"
  loc_110D2401:                   var_74.DispID_0000(FFFFFFFFh)
  loc_110D2408:                   var_150 = "bDeptedit"
  loc_110D2479:                   var_74.DispID_0000(FFFFFFFFh)
  loc_110D2480:                   var_150 = "bItemedit"
  loc_110D24F1:                   var_74.DispID_0000(FFFFFFFFh)
  loc_110D24F8:                   var_150 = "inid"
  loc_110D256A:                   var_74.DispID_0000(1)
  loc_110D256C:                   var_190 = "cdigest"
  loc_110D267D:                   var_74.DispID_0000(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 5))
  loc_110D26A4:                   var_190 = "cCode"
  loc_110D27B3:                   var_74.DispID_0000(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 6))
  loc_110D285B:                   var_7C = var_38.UnkVCall_0000006Ch
  loc_110D28A6:                   var_8150 = (var_38.UnkVCall_0000006Ch = global_1100AE28)
  loc_110D28B3:                   var_160 = var_8150 + 1
  loc_110D293E:                   var_74.DispID_0000(IIf(var_8150 + 1, vbNull, 0))
  loc_110D2A23:                   var_1B0 = "md"
  loc_110D2A6C:                   var_BC = var_25C
  loc_110D2AF3:                   var_74.DispID_0000(Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 7)), "#.00"))
  loc_110D2BE4:                   var_1B0 = "mc"
  loc_110D2C2D:                   var_BC = var_25C
  loc_110D2CB4:                   var_74.DispID_0000(Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 8)), "#.00"))
  loc_110D2D7C:                   If (var_74.DispID_0000("md") <> 0) Then
  loc_110D2DF1:                     If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110D2DFC:                       var_150 = "md_f"
  loc_110D2E6D:                       var_74.DispID_0000(0)
  loc_110D2E77:                     Else
  loc_110D2F2A:                       var_1B0 = "md_f"
  loc_110D2F73:                       var_BC = var_25C
  loc_110D2FFA:                       var_74.DispID_0000(Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_110D303B:                     End If
  loc_110D30AD:                     If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_110D30B8:                       var_150 = "nd_s"
  loc_110D3129:                       var_74.DispID_0000(0)
  loc_110D3133:                     Else
  loc_110D3142:                     Else
  loc_110D31B1:                       If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110D31BC:                         var_150 = "mc_f"
  loc_110D322D:                         var_74.DispID_0000(0)
  loc_110D3237:                       Else
  loc_110D32EA:                         var_1B0 = "mc_f"
  loc_110D3333:                         var_BC = var_25C
  loc_110D33BA:                         var_74.DispID_0000(Format(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_110D33FB:                       End If
  loc_110D346D:                       If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_110D3474:                         GoTo loc_110D30B8
  loc_110D3479:                       End If
  loc_110D3483:                     End If
  loc_110D359D:                     var_74.DispID_0000(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 9)))
  loc_110D35C3:                   End If
  loc_110D3635:                   If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110D3640:                     var_150 = "nfrat"
  loc_110D36B1:                     var_74.DispID_0000(0)
  loc_110D36BB:                   Else
  loc_110D37DF:                     var_74.DispID_0000(Val(frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 11)))
  loc_110D3805:                   End If
  loc_110D385A:                   If var_38.UnkVCall_0000010Ch Then
  loc_110D38F1:                     var_1F0 = "csettle"
  loc_110D39D8:                     var_81A4 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 13) = global_1100AE28)
  loc_110D39E5:                     var_1E0 = var_81A4 + 1
  loc_110D3A70:                     var_74.DispID_0000(IIf(var_81A4 + 1, vbNull, frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 13)))
  loc_110D3AC9:                   End If
  loc_110D3AF2:                   var_24C = var_38.UnkVCall_0000010Ch
  loc_110D3B3F:                   var_250 = var_38.UnkVCall_00000094h
  loc_110D3BDE:                   If (var_38.UnkVCall_0000009Ch = 0) = 0 Then
  loc_110D3C75:                     var_1F0 = "cn_id"
  loc_110D3D24:                     var_E0 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 14)
  loc_110D3D5C:                     var_81BC = (frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 14) = global_1100AE28)
  loc_110D3D69:                     var_1E0 = var_81BC + 1
  loc_110D3DF4:                     var_74.DispID_0000(IIf(var_81BC + 1, vbNull, var_E0))
  loc_110D3EDB:                     var_1F0 = "dt_date"
  loc_110D3F8A:                     var_D0 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 15)
  loc_110D3FA8:                     var_E0)
  loc_110D3FD5:                     var_81C8 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 15) = global_1100AE28)
  loc_110D3FE2:                     var_1E0 = var_81C8 + 1
  loc_110D406D:                     var_74.DispID_0000(IIf(var_81C8 + 1, vbNull, var_E0))
  loc_110D415B:                     var_1F0 = "cname"
  loc_110D4242:                     var_81D4 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_24, &H14) = global_1100AE28)
  loc_110D424F:                     var_1E0 = var_81D4 + 1
  loc_110D42DA:                     var_74.DispID_0000(IIf(var_81D4 + 1, vbNull, frmZGXSToPzTGZP.VFG.DispID_0082(var_24, &H14)))
  loc_110D4333:                   End If
  loc_110D43A9:                   var_250 = var_38.UnkVCall_0000008Ch
  loc_110D43E7:                   If (var_38.UnkVCall_000000A4h = 0) = 0 Then
  loc_110D43F1:                     var_150 = var_24
  loc_110D447E:                     var_1F0 = "cdept_id"
  loc_110D4565:                     var_81E8 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_150, 16) = global_1100AE28)
  loc_110D4572:                     var_1E0 = var_81E8 + 1
  loc_110D45FD:                     var_74.DispID_0000(IIf(var_81E8 + 1, vbNull, frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 16)))
  loc_110D4658:                   Else
  loc_110D465D:                     var_160 = "cdept_id"
  loc_110D46C4:                     var_74.DispID_0000(var_150)
  loc_110D46C9:                   End If
  loc_110D471E:                   If var_38.UnkVCall_0000008Ch Then
  loc_110D4728:                     var_150 = var_24
  loc_110D47B5:                     var_1F0 = "cperson_id"
  loc_110D489C:                     var_81F8 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_150, &H11) = global_1100AE28)
  loc_110D48A9:                     var_1E0 = var_81F8 + 1
  loc_110D4934:                     var_74.DispID_0000(IIf(var_81F8 + 1, vbNull, frmZGXSToPzTGZP.VFG.DispID_0082(var_24, &H11)))
  loc_110D498F:                   Else
  loc_110D4994:                     var_160 = "cperson_id"
  loc_110D49FB:                     var_74.DispID_0000(var_150)
  loc_110D4A00:                   End If
  loc_110D4A55:                   If var_38.UnkVCall_00000094h Then
  loc_110D4A5F:                     var_150 = var_24
  loc_110D4AEC:                     var_1F0 = "ccus_id"
  loc_110D4BD3:                     var_8208 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_150, &H12) = global_1100AE28)
  loc_110D4BE0:                     var_1E0 = var_8208 + 1
  loc_110D4C6B:                     var_74.DispID_0000(IIf(var_8208 + 1, vbNull, frmZGXSToPzTGZP.VFG.DispID_0082(var_24, &H12)))
  loc_110D4CC6:                   Else
  loc_110D4CCB:                     var_160 = "ccus_id"
  loc_110D4D32:                     var_74.DispID_0000(var_150)
  loc_110D4D37:                   End If
  loc_110D4D8C:                   If var_38.UnkVCall_0000009Ch Then
  loc_110D4D96:                     var_150 = var_24
  loc_110D4E23:                     var_1F0 = "csup_id"
  loc_110D4F0A:                     var_8218 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_150, &H13) = global_1100AE28)
  loc_110D4F17:                     var_1E0 = var_8218 + 1
  loc_110D4FA2:                     var_74.DispID_0000(IIf(var_8218 + 1, vbNull, frmZGXSToPzTGZP.VFG.DispID_0082(var_24, &H13)))
  loc_110D4FFD:                   Else
  loc_110D5002:                     var_160 = "csup_id"
  loc_110D5069:                     var_74.DispID_0000(var_150)
  loc_110D506E:                   End If
  loc_110D50E7:                   If (var_38.UnkVCall_000000ACh = global_1100AE28) Then
  loc_110D50F1:                     var_150 = var_24
  loc_110D517E:                     var_1F0 = "citem_id"
  loc_110D5265:                     var_822C = (frmZGXSToPzTGZP.VFG.DispID_0082(var_150, &H15) = global_1100AE28)
  loc_110D5272:                     var_1E0 = var_822C + 1
  loc_110D52FD:                     var_74.DispID_0000(IIf(var_822C + 1, vbNull, frmZGXSToPzTGZP.VFG.DispID_0082(var_24, &H15)))
  loc_110D53DA:                     var_7C = var_38.UnkVCall_000000ACh
  loc_110D542B:                     var_8238 = (var_38.UnkVCall_000000ACh = global_1100AE28)
  loc_110D5438:                     var_160 = var_8238 + 1
  loc_110D54C3:                     var_74.DispID_0000(IIf(var_8238 + 1, vbNull, 0))
  loc_110D54FD:                   Else
  loc_110D5502:                     var_160 = "citem_id"
  loc_110D5569:                     var_74.DispID_0000(var_150)
  loc_110D5570:                     var_160 = "citem_class"
  loc_110D55D7:                     var_74.DispID_0000(var_150)
  loc_110D55DC:                   End If
  loc_110D55E1:                   var_160 = "ccode_equal"
  loc_110D5648:                   var_74.DispID_0000(var_150)
  loc_110D564F:                   var_160 = "iflagbank"
  loc_110D56B6:                   var_74.DispID_0000(var_150)
  loc_110D56BD:                   var_160 = "iflagperson"
  loc_110D5724:                   var_74.DispID_0000(var_150)
  loc_110D5731:                   var_74.Update
  loc_110D5748:                   var_24 = var_24(1)
  loc_110D5759:                   var_68 = var_68(1)
  loc_110D578E:                   var_823C = CLng(frmZGXSToPzTGZP.VFG.DispID_0007)
  loc_110D57AA:                   var_264 = (var_24(1) > 0)
  loc_110D57D1:                   If var_264 = 0 Then GoTo loc_110D0ADB
  loc_110D57D7:                 End If
  loc_110D580A:                 "wksAlias".00000000h.00000000h
  loc_110D5837:                 Set var_9C = frmZGXSToPzTGZP.Chk
  loc_110D5839:                 var_264 = var_9C
  loc_110D584B:                 Set var_A0 = var_9C(0)
  loc_110D586F:                 var_26C = var_A0
  loc_110D58D9:                 If (var_A0.Value = 1) Then
  loc_110D58E7:                   var_70, var_58)
  loc_110D58EC:                 End If
  loc_110D58EE:                 On Error GoTo 0
  loc_110D5925:                 var_250 = CInt("cIYear".00000000h)
  loc_110D594F:                 var_24C, var_250, var_70, var_58)
  loc_110D5959:                 var_5C = var_24C, var_250, var_70, var_58)
  loc_110D599C:                 var_250 = CInt("cIYear".00000000h)
  loc_110D59D0:                 var_48 = r_250, var_70, var_58) var_250, var_70, var_58)
  loc_110D59E2:                 var_150 = "select * from GL_accvouch where ibook=0 and iYear="
  loc_110D5A0A:                 var_170 = var_70
  loc_110D5A2E:                 var_824C = Proc_0_4_11026BD0(var_58, var_54, var_54)
  loc_110D5A33:                 var_190 = var_824C
  loc_110D5A5B:                 var_1B0 = var_54
  loc_110D5AB4:                 var_D8 = 1 & "cIYear".00000000h(, 1, 1) & " and iperiod="
  loc_110D5B1D:                 var_128 = var_D8 & var_70 & " and isignseq=" & var_824C & " and ino_id=" & var_54
  loc_110D5B86:                 Set var_74 = "DataMdb".00000000h.00000001h
  loc_110D5C25:                 If CBool(Not(var_74.EOF)) Then
  loc_110D5C7D:                   If CBool(Not(var_74.EOF)) Then
  loc_110D5C86:                     var_170 = var_70
  loc_110D5C9B:                     var_150 = "iPeriod"
  loc_110D5CBF:                     var_180 = "csign"
  loc_110D5CD3:                     var_1D0 = var_54
  loc_110D5CE4:                     var_1B0 = "ino_id"
  loc_110D5E3B:                     If CBool((var_70 = var_14C) And (var_58 = var_D8) And (var_54 = var_1AC)) Then
  loc_110D5E46:                       var_150 = "mc"
  loc_110D5EC8:                       var_180 = "ccode_equal"
  loc_110D5EDC:                       If (var_14C <> 0) Then
  loc_110D5F08:                         var_8278 = (var_5C = global_1100AE28)
  loc_110D5F15:                         var_160 = var_8278 + 1
  loc_110D5F42:                         var_C8 = IIf(var_8278 + 1, vbNull, var_5C)
  loc_110D5FBC:                       Else
  loc_110D5FE2:                         var_827C = (var_48 = global_1100AE28)
  loc_110D5FEF:                         var_160 = var_827C + 1
  loc_110D601C:                         var_C8 = IIf(var_827C + 1, vbNull, var_48)
  loc_110D6091:                       End If
  loc_110D60A7:                       var_74.Update
  loc_110D60F1:                       var_180 = var_38
  loc_110D6138:                       var_B8 = var_74.DispID_0000("cCode")
  loc_110D6195:                       "kmCodeToProperties".00000002h
  loc_110D61B5:                       Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110D61D4:                       var_150 = "citem_class"
  loc_110D623B:                       If IsNull(var_74.DispID_0000(var_150)) Then
  loc_110D6250:                       Else
  loc_110D6291:                         var_180 = var_28
  loc_110D62D8:                         var_B8 = var_74.DispID_0000(var_150)
  loc_110D6335:                         "XmClassIDToProperties".00000002h
  loc_110D6395:                         var_78 = {3302AA47-EB96-11D2-AF06000021009B21}().UnkVCall_0000002Ch
  loc_110D63C6:                       End If
  loc_110D63D4:                       var_68 = var_68(1)
  loc_110D63E2:                       var_74.MoveNext
  loc_110D63EB:                       GoTo loc_110D5C32
  loc_110D63F0:                       ' Referenced from: 110D0A86
  loc_110D6423:                       "wksAlias".00000000h.00000000h
  loc_110D643B:                       var_30 = var_3C
  loc_110D6450:                       var_1A0 = var_70
  loc_110D6519:                       var_D8)
  loc_110D65C5:                       var_829C = (var_58 = frmZGXSToPzTGZP.VFG.DispID_0082(var_30, 3))
  loc_110D65D2:                       var_1F0 = var_829C + 1
  loc_110D668E:                       var_82A4 = (var_64 = frmZGXSToPzTGZP.VFG.DispID_0082(var_30, 0))
  loc_110D669B:                       var_240 = var_82A4 + 1
  loc_110D6731:                       var_82B0 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_30, 2) = "DateToPeriod".00000001h) And var_829C + 1 And var_82A4 + 1
  loc_110D67BD:                       If CBool(var_82B0) Then
  loc_110D67C7:                         var_150 = var_30
  loc_110D6883:                         frmZGXSToPzTGZP.VFG.DispID_0082(1, "-")
  loc_110D6A03:                         frmZGXSToPzTGZP.VFG.DispID_009E(var_30, 1, var_30, 1, &HFF)
  loc_110D6A18:                         var_150 = var_30
  loc_110D6AD4:                         frmZGXSToPzTGZP.VFG.DispID_0082(&H16, "数据提交错或该数据已经被导入----未引入")
  loc_110D6AF3:                         var_30 = var_30(1)
  loc_110D6B1F:                         var_82B8 = CLng(frmZGXSToPzTGZP.VFG.DispID_0007)
  loc_110D6B3B:                         var_264 = (var_30 > 0)
  loc_110D6B62:                         If var_264 = 0 Then GoTo loc_110D644D
  loc_110D6B68:                       End If
  loc_110D6B6B:                       var_24 = var_30
  loc_110D6B7F:                       Set var_9C = frmZGXSToPzTGZP.Chk
  loc_110D6B81:                       var_264 = var_9C
  loc_110D6B93:                       Set var_A0 = var_9C(0)
  loc_110D6BB7:                       var_26C = var_A0
  loc_110D6C21:                       If (var_A0.Value = 1) Then
  loc_110D6D1D:                         "unLockVouch".00000004h(var_180, var_BC, var_C4, 0, var_74, var_70, var_58, var_16C, var_54, &H4002, var_184)
  loc_110D6D26:                       End If
  loc_110D6D2B:                       var_150 = "VouchNum"
  loc_110D6DA0:                       Set var_34 = "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C4, 0, var_14C, "VouchNum", var_154)
  loc_110D6DC1:                       var_150 = "delete  from vouchnum"
  loc_110D6E1F:                       "DataMdb".00000000h.00000001h(1, 1, var_180, var_BC, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_110D6E7C:                       frmZGXSToPzTGZP.Pic1.DispID_80010007 = var_150
  loc_110D6E90:                       var_82C4 = Resume(0)
  loc_110D6E96:                     End If
  loc_110D6E96:                   End If
  loc_110D6E96:                 End If
  loc_110D6EB4:                 var_24 = var_27C+(var_24 - 1)
  loc_110D6EB7:                 GoTo loc_110D01BF
  loc_110D6EBC:               End If
  loc_110D6EBF:               var_1A0 = var_70
  loc_110D6F88:               var_D8)
  loc_110D7034:               var_82D0 = (var_58 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 3))
  loc_110D7041:               var_1F0 = var_82D0 + 1
  loc_110D70FD:               var_82D8 = (var_64 = frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 0))
  loc_110D710A:               var_240 = var_82D8 + 1
  loc_110D71A7:               var_82E8 = CBool((frmZGXSToPzTGZP.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_82D0 + 1 And var_82D8 + 1)
  loc_110D71AD:               var_264 = var_82E8
  loc_110D722C:               If var_264 = 0 Then GoTo loc_110D6E96
  loc_110D7243:               Set var_9C = frmZGXSToPzTGZP.Chk
  loc_110D7245:               var_264 = var_9C
  loc_110D7257:               Set var_A0 = var_9C(0)
  loc_110D727B:               var_26C = var_A0
  loc_110D72BE:               var_274 = (var_A0.Value = 1)
  loc_110D72E9:               var_150 = var_24
  loc_110D730A:               var_190 = "网络共享冲突----未引入"
  loc_110D7314:               If var_274 = 0 Then
  loc_110D7316:                 var_190 = "指定的凭证号无效或重号----未引入"
  loc_110D7320:               End If
  loc_110D73B1:               frmZGXSToPzTGZP.VFG.DispID_0082(var_170, var_190)
  loc_110D73D0:               var_24 = var_24(1)
  loc_110D73D6:               var_2A8 = var_24(1)
  loc_110D7405:               var_82EC = CLng(frmZGXSToPzTGZP.VFG.DispID_0007)
  loc_110D7421:               var_264 = (var_2A8 > 0)
  loc_110D7448:               If var_264 = 0 Then GoTo loc_110D6EBC
  loc_110D744E:               GoTo loc_110D6E96
  loc_110D7453:             End If
  loc_110D7456:             var_1A0 = var_70
  loc_110D7521:             var_D8)
  loc_110D75CF:             var_82F8 = (var_58 = frmZGXSToPzTGZP.VFG.DispID_0082(var_2A8, 3))
  loc_110D75DC:             var_1F0 = var_82F8 + 1
  loc_110D769A:             var_8300 = (var_64 = frmZGXSToPzTGZP.VFG.DispID_0082(var_2A8, 0))
  loc_110D76A7:             var_240 = var_8300 + 1
  loc_110D7744:             var_8310 = CBool((frmZGXSToPzTGZP.VFG.DispID_0082(var_2A8, 2) = "DateToPeriod".00000001h) And var_82F8 + 1 And var_8300 + 1)
  loc_110D774A:             var_264 = var_8310
  loc_110D77C9:             If var_264 = 0 Then GoTo loc_110D6E96
  loc_110D78C0:             If (frmZGXSToPzTGZP.VFG.DispID_0082(var_2A8, &H16) = global_1100AE28) + 1 Then
  loc_110D78C6:               var_150 = var_2A8
  loc_110D797F:               Set var_9C = frmZGXSToPzTGZP.VFG
  loc_110D7982:               var_9C.DispID_0082(&H16, "凭证借贷不平衡或某分录有错误----未引入")
  loc_110D7993:               GoTo loc_110D7453
  loc_110D7998:             End If
  loc_110D7A62:             var_C0 = frmZGXSToPzTGZP.VFG.DispID_0082(frmZGXSToPzTGZP.VFG, &H16) & "----未引入"
  loc_110D7AFF:             frmZGXSToPzTGZP.VFG.DispID_0082(&H16, var_C0)
  loc_110D7B3C:             GoTo loc_110D7453
  loc_110D7B41:           End If
  loc_110D7B89:           frmZGXSToPzTGZP.Pic1.DispID_80010007 = var_150
  loc_110D7BA0:           If var_2C Then
  loc_110D7C32:             MsgBox("数据引入已完成，数据已生成用友凭证。", 64, "提示信息", 10, 10)
  loc_110D7CA4:             frmZGXSToPzTGZP.VFG.DispID_0007 = 1
  loc_110D7CFE:             frmZGXSToPzTGZP.VFG.DispID_0007 = 1
  loc_110D7D99:             frmZGXSToPzTGZP.sBar.DispID_6803001E(1100AE28h)
  loc_110D7E30:             frmZGXSToPzTGZP.sBar.DispID_6803001E(1100AE28h)
  loc_110D7EC7:             Set var_9C = frmZGXSToPzTGZP.sBar
  loc_110D7ECA:             var_9C.DispID_6803001E(1100AE28h)
  loc_110D7EE0:           Else
  loc_110D7F67:             MsgBox("数据没有被引入，原因请查看最后一列中的说明。", 64, "提示信息", 10, 10)
  loc_110D7F94:           End If
  loc_110D7F99:           var_150 = "VouchNum"
  loc_110D8010:           Set var_34 = "DataMdb".00000000h.00000001h(var_180, var_BC, var_C0, var_C4, var_C8, var_14C, "VouchNum", var_154)
  loc_110D8031:           var_150 = "delete  from vouchnum"
  loc_110D8085:           "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C0, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_110D80DA:           "AS13")
  loc_110D80FA:           var_24C = frmZGXSToPzTGZP.UpdateBTData
  loc_110D8143:           var_B8 = Global.Screen
  loc_110D8161:           var_8330 = ecx
  loc_110D8169:           var_8330 = var_9C.UnkVCall_0000007Ch
  loc_110D817D:         End If
  loc_110D817D:       End If
  loc_110D817D:     End If
  loc_110D817D:   End If
  loc_110D817D: End If
  loc_110D8189: Exit Sub
  loc_110D8195: GoTo loc_110D824E
  loc_110D824D: Exit Sub
  loc_110D824E: ' Referenced from: 110CF61C
  loc_110D824E: ' Referenced from: 110D8195
End Sub

Private Sub Proc_16_13_110D8D10
  Dim var_58 As Variant
  Dim var_5C As Variant
  Dim var_64 As frmZGXSToPzTGZP.Label3
  Dim var_1D0 As Label
  loc_110D8DFD: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110D8E06: var_1F0 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110D8E23: Set var_58 = frmZGXSToPzTGZP.Chk
  loc_110D8E2D: var_1D0 = var_58
  loc_110D8E33: Set var_5C = var_58(0)
  loc_110D8E5E: var_1D8 = var_5C
  loc_110D8EA1: var_1E0 = (var_5C.Value = 1)
  loc_110D8EB7: If var_1E0 = 0 Then
  loc_110D8F1C:   If var_14 <= CLng(frmZGXSToPzTGZP.VFG.DispID_0007)(-1) Then
  loc_110D8F91:     var_7C = frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 2)
  loc_110D8FAC:     var_94)
  loc_110D9007:     var_30 = CByte("DateToPeriod".00000001h)
  loc_110D9161:     Set var_64 = frmZGXSToPzTGZP.Label3
  loc_110D918B:     var_1D0 = var_64
  loc_110D9341:     var_94 = frmZGXSToPzTGZP.VFG.DispID_0082(var_14, frmZGXSToPzTGZP.VFG)
  loc_110D935D:     var_8034 = "正在处理：第[" & frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 2) & " - " & frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 3) & " - " & var_94
  loc_110D9393:     var_64.Caption = var_8034 & "]号凭证是否重号"
  loc_110D9422:     var_803C = frmZGXSToPzTGZP.Proc_16_14_110D9B20(var_30)
  loc_110D9437:     If var_1CC <= 0 Then
  loc_110D9449:       var_13C = var_30
  loc_110D94E0:       var_94)
  loc_110D9579:       var_804C = (frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 3) = frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 3))
  loc_110D95A6:       var_17C = var_804C + 1
  loc_110D961D:       var_8054 = (frmZGXSToPzTGZP.VFG.DispID_0082(var_14, frmZGXSToPzTGZP.VFG) = frmZGXSToPzTGZP.VFG.DispID_0082(var_14, ""))
  loc_110D9644:       var_1BC = var_8054 + 1
  loc_110D973F:       If CBool((frmZGXSToPzTGZP.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_804C + 1 And var_8054 + 1) Then
  loc_110D97D2:         frmZGXSToPzTGZP.VFG.DispID_0082(var_10C, 285267820)
  loc_110D9906:         frmZGXSToPzTGZP.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110D999A:         frmZGXSToPzTGZP.VFG.DispID_0082(var_10C, "指定的凭证号无效或重号")
  loc_110D99E5:         var_8068 = CLng(frmZGXSToPzTGZP.VFG.DispID_0007)
  loc_110D9A03:         var_1D0 = (var_14(1) > 0)
  loc_110D9A20:         If var_1D0 = 0 Then GoTo loc_110D9443
  loc_110D9A26:       End If
  loc_110D9A34:     Else
  loc_110D9A3D:     End If
  loc_110D9A4A:     var_14 = 1+var_14
  loc_110D9A4D:     GoTo loc_110D8F16
  loc_110D9A52:   End If
  loc_110D9A52: End If
  loc_110D9A57: GoTo loc_110D9AE8
  loc_110D9AE7: Exit Sub
  loc_110D9AE8: ' Referenced from: 110D9A57
End Sub

Private  Proc_16_14_110D9B20(arg_C, arg_10, arg_14) '110D9B20
  loc_110D9BB9: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110D9BC2: var_168 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110D9BEB: If IsNumeric(arg_14) Then
  loc_110D9BFA:   var_8008 = CLng(Val(arg_14))
  loc_110D9C04:   If var_8008 > 0 Then
  loc_110D9C10:     If var_8008 <= 9999 Then
  loc_110D9C8C:       var_8028 = "select * from GL_accvouch where iperiod >=" & CStr(arg_C) & " and isignseq>=" & CStr(0) & " and ino_id>=" & CStr(var_8008)
  loc_110D9CA1:       var_44 = var_8028
  loc_110D9CF3:       Set var_1C = "DataMdb".00000000h.00000001h(fs:[00000000h], , , , , var_40, var_8028, var_48)
  loc_110D9D38:       var_8030 = Proc_0_4_11026BD0(arg_10, , )
  loc_110D9D59:       var_8034 = CBool(var_1C.EOF)
  loc_110D9D6D:       If var_8034 = 0 Then
  loc_110D9D98:         var_F4 = arg_C
  loc_110D9E56:         var_8040 = (var_1C.DispID_0000("iPeriod") = arg_C) And (var_1C.DispID_0000("isignseq") = (Proc_0_4_11026BD0(arg_10, , ) And 255))
  loc_110D9EC6:         var_804C = CBool(Not(var_8040 And (var_1C.DispID_0000("ino_id") = var_8008)))
  loc_110D9EEB:         If var_804C = 0 Then GoTo loc_110D9EF0
  loc_110D9EED:       End If
  loc_110D9EFB:       var_1C.oClose
  loc_110D9F04:     End If
  loc_110D9F04:   End If
  loc_110D9F04: End If
  loc_110D9F0A: GoTo loc_110D9F6F
  loc_110D9F6E: Exit Sub
  loc_110D9F6F: ' Referenced from: 110D9F0A
End Sub
