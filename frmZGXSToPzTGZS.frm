VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C908002B2F49FB}#1.2#0"; "C:\Windows\SysWow64\Comdlg32.ocx"
Begin VB.Form frmZGXSToPzTGZS
  Caption = "销售暂估导转凭证（TGZS）"
  BackColor = &H80000005&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmZGXSToPzTGZS.frx":0000
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 345
  ClientWidth = 9255
  ClientHeight = 5700
  LockControls = -1  'True
  Appearance = 0 'Flat
  Begin C1SizerLibCtl.C1Elastic Pic1
    Left = 3300
    Top = 3480
    Width = 5025
    Height = 675
    Visible = 0   'False
    TabStop = 0   'False
    TabIndex = 3
    OleObjectBlob = "frmZGXSToPzTGZS.frx":014A
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
    OleObjectBlob = "frmZGXSToPzTGZS.frx":0293
    Begin AIFCmp1.asxStatusBar sBar
      Left = 0
      Top = 5355
      Width = 12045
      Height = 345
      OleObjectBlob = "frmZGXSToPzTGZS.frx":04BC
    End
    Begin C1SizerLibCtl.C1Elastic C1Elastic2
      Left = 0
      Top = 0
      Width = 12045
      Height = 435
      TabStop = 0   'False
      TabIndex = 1
      OleObjectBlob = "frmZGXSToPzTGZS.frx":05EC
    End
    Begin VSFlex8LCtl.VSFlexGrid VFG
      Left = 0
      Top = 1260
      Width = 12045
      Height = 4080
      TabIndex = 2
      OleObjectBlob = "frmZGXSToPzTGZS.frx":0753
      Begin MSComDlg.CommonDialog dlg
        OleObjectBlob = "frmZGXSToPzTGZS.frx":0BBC
        Left = 0
        Top = 0
      End
    End
    Begin AIFCmp1.asxPanel asxPanel1
      Left = 0
      Top = 450
      Width = 12045
      Height = 795
      OleObjectBlob = "frmZGXSToPzTGZS.frx":0C20
      Begin VB.ComboBox Cbo
        Style = 2
        Left = 630
        Top = 60
        Width = 4545
        Height = 300
        Visible = 0   'False
        TabIndex = 14
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 3
        Left = 7935
        Top = 435
        Width = 600
        Height = 270
        TabIndex = 9
        OleObjectBlob = "frmZGXSToPzTGZS.frx":0D00
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 4
        Left = 8595
        Top = 435
        Width = 720
        Height = 270
        TabIndex = 13
        OleObjectBlob = "frmZGXSToPzTGZS.frx":0EA0
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
        Left = 5175
        Top = 435
        Width = 870
        Height = 270
        TabIndex = 7
        OleObjectBlob = "frmZGXSToPzTGZS.frx":1040
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 1
        Left = 6090
        Top = 435
        Width = 870
        Height = 270
        TabIndex = 8
        OleObjectBlob = "frmZGXSToPzTGZS.frx":1238
      End
      Begin TDBText6Ctl.TDBText TDBText
        Left = 30
        Top = 435
        Width = 5115
        Height = 270
        TabIndex = 10
        OleObjectBlob = "frmZGXSToPzTGZS.frx":1408
        ToolTipText = "项目大类"
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 2
        Left = 7005
        Top = 435
        Width = 870
        Height = 270
        Visible = 0   'False
        TabIndex = 11
        OleObjectBlob = "frmZGXSToPzTGZS.frx":1564
      End
      Begin TDBDate6Ctl.TDBDate TDBDate
        Left = 9390
        Top = 420
        Width = 2385
        Height = 285
        TabIndex = 12
        OleObjectBlob = "frmZGXSToPzTGZS.frx":1708
      End
      Begin VB.Label Label1
        Caption = "选择："
        Left = 90
        Top = 120
        Width = 555
        Height = 345
        Visible = 0   'False
        TabIndex = 15
        BackStyle = 0 'Transparent
      End
    End
  End
End

Attribute VB_Name = "frmZGXSToPzTGZS"


Private  APB_UnknownEvent_9(arg_C) '110F3630
  Dim var_20 As Variant
  Dim var_AC As Scripting.FileSystemObject
  loc_110F36A7: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110F36B0: var_C4 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110F36D7: arg_C = frmZGXSToPzTGZS.APB.UnkVCall_00000040h
  loc_110F3715: var_B8 = var_24.DispID_FFFFFDFA
  loc_110F3749: var_8008 = (var_B8 = "加载数据")
  loc_110F374D: If var_8008 = 0 Then
  loc_110F3770:   var_AC = var_18
  loc_110F37AB:   var_1C = frmZGXSToPzTGZS.TDBText.DispID_0000
  loc_110F37BB:   var_8014 = Scripting.FileSystemObject.FileExists
  loc_110F37F9:   If Not (var_A8) Then
  loc_110F385C:     MsgBox("文件不存在或非法路径！ ", 64, "提示", 10, 10)
  loc_110F3882:   Else
  loc_110F3894:     If frmZGXSToPzTGZS.FillData < 0 Then
  loc_110F38A6:       var_A8 = CheckObj(%ecx = %S_edx_S = %S_edx_S, " 滬砗丂J?嶠茒礷" & Chr(12), 1788)
  loc_110F38B1:     End If
  loc_110F38BD:     call ebx("取消加载", var_B8, var_1C, var_A8, var_24)
  loc_110F38C1:     If ebx("取消加载", var_B8, var_1C, var_A8, var_24) = 0 Then
  loc_110F38F1:       var_44 = "提示信息"
  loc_110F391F:       var_2C = "是否取消数据载入？" & vbCrLf & "取消数据载入，数据将全部清空。"
  loc_110F393B:       MsgBox(var_2C, 292, var_44, var_54, var_64)
  loc_110F3972:       If (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6) = 0 Then GoTo loc_110F3A27
  loc_110F3983:     Else
  loc_110F398F:       (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_110F3993:       If var_B8 = 0 Then
  loc_110F3998:         var_8020 = frmZGXSToPzTGZS.Proc_17_10_110EA530("凭证导入")
  loc_110F39A3:       Else
  loc_110F39AF:         (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_110F39B3:         If var_8020 Then
  loc_110F39C1:           (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_110F39C5:           If var_B8 = 0 Then
  loc_110F39F8:             Set var_20 = var_C4 = %S_edx_S
  loc_110F3A06:             var_8028 = Global.Unload var_B8
  loc_110F3A27:           End If
  loc_110F3A27:         End If
  loc_110F3A27:       End If
  loc_110F3A27:     End If
  loc_110F3A27:   End If
  loc_110F3A27: End If
  loc_110F3A2F: GoTo loc_110F3A66
  loc_110F3A65: Exit Sub
  loc_110F3A66: ' Referenced from: 110F3A2F
End Sub

Private Sub Form_Load() '110E29B0
  Dim var_18 As Variant
  Dim var_1C As var_18.DispID_03E8
  loc_110E2A1A: Set var_18 = frmZGXSToPzTGZS.TDBText
  loc_110E2A21: var_2C = var_18.DispID_03E8
  loc_110E2A42: var_18.DispID_03E8.UnkVCall_00000030h
  loc_110E2A90: Set var_18 = frmZGXSToPzTGZS.TDBDate
  loc_110E2A97: var_2C = var_18.DispID_03E8
  loc_110E2AAC: Set var_1C = var_18.DispID_03E8
  loc_110E2AB8: var_1C.UnkVCall_00000030h
  loc_110E2B27: frmZGXSToPzTGZS.TDBDate.DispID_0000 = Date
  loc_110E2B56: frmZGXSToPzTGZS.APB.UnkVCall_00000040h
  loc_110E2B94: var_1C.DispID_80010007 = var_30
  loc_110E2BBB: Set var_18 = frmZGXSToPzTGZS.APB
  loc_110E2BC8: var_18.UnkVCall_00000040h
  loc_110E2C03: var_1C.DispID_80010007 = var_30
  loc_110E2C1F: var_8004 = frmZGXSToPzTGZS.Proc_17_7_110DB140(var_18)
  loc_110E2C31: GoTo loc_110E2C50
  loc_110E2C4F: Exit Sub
  loc_110E2C50: ' Referenced from: 110E2C31
End Sub

Private Sub Form_Resize() '110E2C70
  loc_110E2CFD: var_38 = frmZGXSToPzTGZS.Pic1.DispID_80010005
  loc_110E2D21: var_48 = frmZGXSToPzTGZS.Pic1.DispID_80010006
  loc_110E2D34: var_EC = var_48.ScaleWidth
  loc_110E2D6B: If global_110F6000 = 0 Then
  loc_110E2D75: Else
  loc_110E2D80: End If
  loc_110E2D80: var_70 = ((var_EC - CSgn(var_38)) / 2)
  loc_110E2D95: var_F0 = var_48.ScaleHeight
  loc_110E2DD3: If global_110F6000 = 0 Then
  loc_110E2DDD: Else
  loc_110E2DE8: End If
  loc_110E2EF3: frmZGXSToPzTGZS.Pic1.DispID_80011002(var_70, ((var_F0 - CSgn(var_48)) / 2), CSgn(frmZGXSToPzTGZS.Pic1.DispID_80010005), CSgn(frmZGXSToPzTGZS.Pic1.DispID_80010006))
  loc_110E2F3C: GoTo loc_110E2F76
End Sub

Private Sub TDBText_UnknownEvent_B '110F3AA0
  Dim var_64 As frmZGXSToPzTGZS.dlg
  loc_110F3B07: Set var_64 = frmZGXSToPzTGZS.dlg
  loc_110F3B39: var_64.FileName = var_48
  loc_110F3B5E: var_64.DialogTitle = var_48
  loc_110F3B83: var_64.Filter = var_48
  loc_110F3BA5: var_64.CancelError = var_48
  loc_110F3BAF: var_64.ShowOpen
  loc_110F3BC1: var_64.FileName = var_64
  loc_110F3C07: If (var_64 = global_1100AE28) Then
  loc_110F3C15:   var_64.FileName = Me
  loc_110F3C5D:   frmZGXSToPzTGZS.TDBText.DispID_0000 = var_2C
  loc_110F3C87: End If
  loc_110F3C93: GoTo loc_110F3CBB
  loc_110F3CBA: Exit Sub
  loc_110F3CBB: ' Referenced from: 110F3C93
End Sub

Public Sub ExitForm(Cancel, UnloadMode) '110DB060
  Dim var_18 As Global
  loc_110DB09F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110DB0CA: Set var_18 = Me
  loc_110DB0D2: var_8008 = Global.Unload
  loc_110DB10C: GoTo loc_110DB118
  loc_110DB117: Exit Sub
  loc_110DB118: ' Referenced from: 110DB10C
End Sub

Public Function FillData() '110DCA90
  Dim var_B0 As Variant
  Dim var_58 As Variant
  Dim var_B4 As frmZGXSToPzTGZS.TDBText
  Dim var_4C As Variant
  Dim var_38 As Variant
  Dim var_30 As Me
  Dim var_2C As ADODB.Recordset
  loc_110DCC16: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110DCC1C: var_310 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110DCC70: frmZGXSToPzTGZS.VFG.DispID_0007 = 1
  loc_110DCC93: Set var_B0 = frmZGXSToPzTGZS.Label3
  loc_110DCC9D: var_2E4 = var_B0
  loc_110DCCA3: var_B0.Caption = "正在打开Excel数据表，请稍候。。。"
  loc_110DCD16: frmZGXSToPzTGZS.Pic1.DispID_80010007 = True
  loc_110DCD42: frmZGXSToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_110DCD5C: var_8004 = CreateObject(global_1100D5A4)
  loc_110DCD67: Set var_58 = CreateObject(global_1100D5A4)
  loc_110DCD76: var_B0 = var_58.UnkVCall_000000D0h
  loc_110DCE09: var_2E8 = var_B0
  loc_110DD053: Set var_B4 = frmZGXSToPzTGZS.TDBText
  loc_110DD07B: var_84 = var_B4.DispID_0000
  loc_110DD08B: var_84 = var_B0.UnkVCall_0000004Ch
  loc_110DD0FE: var_B0 = var_4C.Tag
  loc_110DD1A6: var_B0.Activate
  loc_110DD20D: Set var_7C = var_B0.UsedRange
  loc_110DD23C: Set var_B0 = frmZGXSToPzTGZS.Label3
  loc_110DD246: var_2E4 = var_B0
  loc_110DD24C: var_B0.Caption = "正在填充数据，请稍候。。。"
  loc_110DD2BF: frmZGXSToPzTGZS.Pic1.DispID_80010007 = True
  loc_110DD2EC: frmZGXSToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_110DD326: Set var_B0 = frmZGXSToPzTGZS.APB
  loc_110DD334: var_2E4 = var_B0
  loc_110DD33A: var_B0.UnkVCall_00000040h
  loc_110DD3D0: Set var_B0 = frmZGXSToPzTGZS.APB
  loc_110DD3DE: var_2E4 = var_B0
  loc_110DD3E4: var_B0.UnkVCall_00000040h
  loc_110DD48A: frmZGXSToPzTGZS.APB.UnkVCall_00000040h
  loc_110DD576: var_E8 = 1100D68Ch & var_7C.Rows.Count
  loc_110DD609: frmZGXSToPzTGZS.sBar.DispID_6803001E(var_E8 & "条记录")
  loc_110DD654: var_30 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[T_CY_ZGXS_Temp]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1) "
  loc_110DD663: var_8014 = var_30 & "CREATE TABLE [T_CY_ZGXS_Temp](cCode VARCHAR(50) NULL,cCusCode VARCHAR(50) NULL,cInvCode VARCHAR(50) NULL,cDepCode VARCHAR(50) NULL,iQuantity float NULL,iMoney Money NULL,iMoney1 Money NULL)"
  loc_110DD66E: var_30 = var_8014
  loc_110DD6B3: var_C8 = UnkObj.UnkVCall_00000040h
  loc_110DD6F7: var_30 = "DELETE FROM [T_CY_ZGXS_Temp]"
  loc_110DD791: Set var_B0 = frmZGXSToPzTGZS.TDBDate
  loc_110DD7AF: var_D0 = var_B0.DispID_004E
  loc_110DD7CD: var_E8)
  loc_110DD82B: var_74 = CByte("DateToPeriod".00000001h)
  loc_110DD8BB: var_38.UnkVCall_00000064h
  loc_110DD967: var_60 = Proc_0_11_11029000(var_B0.Cells(1, 1).value, var_38, 2)
  loc_110DD996: var_24 = "1411"
  loc_110DD9DA: var_D8 = var_7C.Rows.Count
  loc_110DDA35: If var_18 <= CLng(var_D8 + 1) Then
  loc_110DDA40:   If global_56 = 0 Then
  loc_110DDAB1:     var_D8.BackColor = var_1C0
  loc_110DDBFA:     var_350 = (Proc_0_11_11029000(var_B4.Cells(var_18, 1).value, var_B4, var_B0) = "汇总") + 1
  loc_110DDCB0:     var_2EC = (Proc_0_11_11029000(var_B0.Cells(var_18, 1).value, var_1C4, var_1C0) = global_1100AE28) + 1
  loc_110DDD02:     If var_2EC = 0 Then
  loc_110DDDC1:       Set var_B0 = frmZGXSToPzTGZS.sBar
  loc_110DDDC8:       var_B0.DispID_6803001E("正在填充数据：" & CStr(vbNull) & "条记录")
  loc_110DDE22:       var_2E4 = var_2C
  loc_110DDE28:       var_2E0 = ADODB.Recordset.State
  loc_110DDE53:       If var_2E0 = 1 Then
  loc_110DDE71:         var_2E4 = var_2C
  loc_110DDE77:         var_8050 = ADODB.Recordset.Close
  loc_110DDE9B:       End If
  loc_110DDF05:       ADODB.Recordset.BackColor = 1
  loc_110DDFB1:       var_68 = Proc_0_11_11029000(var_B0.Cells(var_18, 4).value, var_B0, var_1BC)
  loc_110DDFF1:       var_1C0 = var_68
  loc_110DE088:       var_2E4 = var_2C
  loc_110DE0C5:       var_8064 = ADODB.Recordset.Open(var_B0 & Proc_0_10_11028DD0(&H4008, "SELECT * FROM Inventory WHERE cInvCode=", var_B0), var_1C4, var_B0 & Proc_0_10_11028DD0(&H4008, "SELECT * FROM Inventory WHERE cInvCode=", var_B0), var_1BC, 9)
  loc_110DE10C:       var_2E4 = var_2C
  loc_110DE138:       If ADODB.Recordset.EOF Then
  loc_110DE144:       Else
  loc_110DE167:         var_2E4 = var_2C
  loc_110DE16D:         var_B0 = ADODB.Recordset.Fields
  loc_110DE196:         var_1C0 = "cInvCCode"
  loc_110DE1A8:         var_2EC = var_B0
  loc_110DE1DE:         ADODB.Recordset.8 = Forms
  loc_110DE20C:         var_2F4 = var_B4
  loc_110DE2B3:         var_2FC = Mid$(var_B4.UnkVCall_00000034h, 1, 2)
  loc_110DE2C5:         var_8078 = (var_2FC = "31")
  loc_110DE2CD:         If var_8078 = 0 Then
  loc_110DE2D6:         Else
  loc_110DE2EF:           If (var_2FC = "32") Then
  loc_110DE2F6:           End If
  loc_110DE2F6:         End If
  loc_110DE2F6:       End If
  loc_110DE33C:       var_1E0 = var_68
  loc_110DE353:       var_1F0 = var_24
  loc_110DE359:       var_8080 = Proc_0_10_11028DD0(&H4008, "INSERT INTO [T_CY_ZGXS_Temp](cCode,cCusCode,cInvCode,cDepCode,iQuantity,iMoney,iMoney1) VALUES (", var_1C4)
  loc_110DE3D5:       var_8098 = Proc_0_10_11028DD0(&H4008, var_B4 & Proc_0_10_11028DD0(&H4008, "51019002" & var_8080 & global_1100AC40, var_1BC) & global_1100AC40, var_1D4)
  loc_110DE4D4:       var_B4.BackColor = CInt(1)
  loc_110DE566:       var_118 = var_B4.Cells(var_18, 10).value
  loc_110DE5A5:       var_118.BackColor = CInt(1)
  loc_110DE637:       var_178 = var_B8.Cells(var_18, 12).value
  loc_110DE82D:       var_E8 = 0 & Proc_0_10_11028DD0(&H4008, var_60 & var_8098 & global_1100AC40, var_1CC) & global_1100AC40 & var_B0.Cells(var_18, 6).value
  loc_110DE890:       var_30 = var_E8 & 1100AC40h & Format(var_118, "0.00") & 1100AC40h & Format(var_178, "0.00") & 1100BD88h
  loc_110DE9BD:       var_28 = var_28(1)
  loc_110DE9C5:       If var_18 Mod 00000064h = 0 Then
  loc_110DE9C7:         DoEvents
  loc_110DE9CD:       End If
  loc_110DE9DD:       var_18 = 1+var_18
  loc_110DE9E0:       GoTo loc_110DDA2F
  loc_110DE9E5:     End If
  loc_110DEA31:     frmZGXSToPzTGZS.VFG.DispID_0007 = 1
  loc_110DEA46:     global_56 = 0
  loc_110DEA6E:     Set var_B0 = frmZGXSToPzTGZS.APB
  loc_110DEA80:     var_2E4 = var_B0
  loc_110DEA86:     var_B0.UnkVCall_00000040h
  loc_110DEB1C:     Set var_B0 = frmZGXSToPzTGZS.APB
  loc_110DEB2E:     var_2E4 = var_B0
  loc_110DEB34:     var_B0.UnkVCall_00000040h
  loc_110DEBDE:     frmZGXSToPzTGZS.APB.UnkVCall_00000040h
  loc_110DEC43:   Else
  loc_110DEC68:     Set var_B0 = frmZGXSToPzTGZS.APB
  loc_110DEC7A:     var_2E4 = var_B0
  loc_110DEC80:     var_B0.UnkVCall_00000040h
  loc_110DED16:     Set var_B0 = frmZGXSToPzTGZS.APB
  loc_110DED28:     var_2E4 = var_B0
  loc_110DED2E:     var_B0.UnkVCall_00000040h
  loc_110DEDD8:     frmZGXSToPzTGZS.APB.UnkVCall_00000040h
  loc_110DEE38:   End If
  loc_110DEE43: End If
  loc_110DEE7C: var_44 = global_11013374 & CStr(var_74) & global_1100D708
  loc_110DEEBC: var_2E0 = ADODB.Recordset.State
  loc_110DEEE1: If var_2E0 = 1 Then
  loc_110DEF01:   var_80C8 = ADODB.Recordset.Close
  loc_110DEF1F: End If
  loc_110DEFB6: var_2E4 = var_2C
  loc_110DEFE5: var_80D4 = ADODB.Recordset.Open("SELECT '113109' AS cCode,cCusCode,SUM(iMoney1) AS iMoney1 " & "FROM [T_CY_ZGXS_Temp] GROUP BY cCusCode", var_1C4, "SELECT '113109' AS cCode,cCusCode,SUM(iMoney1) AS iMoney1 " & "FROM [T_CY_ZGXS_Temp] GROUP BY cCusCode", var_1BC, 9)
  loc_110DF032: var_2E4 = var_2C
  loc_110DF038: var_2DC = ADODB.Recordset.EOF
  loc_110DF05E: If var_2DC = 0 Then
  loc_110DF335:   var_E8 = "1" & Chr(9) & 1100AE28h & Chr(9) & frmZGXSToPzTGZS.TDBDate.DispID_004E & Chr(9) & 1100D6D4h & Chr(9) & 1100C008h & Chr(9) & var_44
  loc_110DF3A6:   var_2E4 = var_2C
  loc_110DF3E7:   var_2EC = ADODB.Recordset.Fields
  loc_110DF41D:   ADODB.Recordset.8 = Forms
  loc_110DF528:   var_2E4 = var_2C
  loc_110DF573:   var_2EC = ADODB.Recordset.Fields
  loc_110DF59F:   ADODB.Recordset.8 = Forms
  loc_110DF6C1:   var_E8 = 9 & Chr(9) & Proc_0_11_11029000(9, var_1D4, "cCode") & Chr(9) & Proc_0_11_11029000(9, var_1D4, "iMoney1") & Chr(9) & 1100C008h
  loc_110DF80A:   var_1C0 = var_E8 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110DF969:   var_E8 = var_1C0 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110DFAB2:   var_1C0 = var_E8 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110DFBFA:   var_2E4 = var_2C
  loc_110DFC45:   var_2EC = ADODB.Recordset.Fields
  loc_110DFC71:   ADODB.Recordset.8 = Forms
  loc_110DFCF1:   var_108 = var_1C0 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1D4, "cCusCode")
  loc_110DFEB1:   var_48 = var_108 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110DFEEF:   var_2E4 = var_2C
  loc_110DFEF5:   var_8150 = ADODB.Recordset.MoveNext
  loc_110DFF6B:   frmZGXSToPzTGZS.VFG.DispID_0080(var_48)
  loc_110DFF80:   GoTo loc_110DF00F
  loc_110DFF85: End If
  loc_110DFFA8: var_2E4 = var_2C
  loc_110DFFAE: var_2E0 = ADODB.Recordset.State
  loc_110DFFD9: If var_2E0 = 1 Then
  loc_110DFFF7:   var_2E4 = var_2C
  loc_110DFFFD:   var_815C = ADODB.Recordset.Close
  loc_110E0021: End If
  loc_110E0038: var_8160 = "SELECT cCode,cCusCode,cDepCode,cInvCode,SUM(iQuantity) AS iQuantity,SUM(iMoney) AS iMoney " & "FROM [T_CY_ZGXS_Temp] GROUP BY cCode,cCusCode,cDepCode,cInvCode"
  loc_110E00AD: var_2E4 = var_2C
  loc_110E00EA: var_8168 = ADODB.Recordset.Open(var_8160, var_1C4, var_8160, var_1BC, 9)
  loc_110E0131: var_2E4 = var_2C
  loc_110E0137: var_2DC = ADODB.Recordset.EOF
  loc_110E015D: If var_2DC = 0 Then
  loc_110E0434:   var_E8 = "1" & Chr(9) & 1100AE28h & Chr(9) & frmZGXSToPzTGZS.TDBDate.DispID_004E & Chr(9) & 1100D6D4h & Chr(9) & 1100C008h & Chr(9) & var_44
  loc_110E04A5:   var_2E4 = var_2C
  loc_110E04E6:   var_2EC = ADODB.Recordset.Fields
  loc_110E051C:   ADODB.Recordset.8 = Forms
  loc_110E06AF:   var_2E4 = var_2C
  loc_110E06F0:   var_2EC = ADODB.Recordset.Fields
  loc_110E0726:   ADODB.Recordset.8 = Forms
  loc_110E07A6:   var_108 = 9 & Chr(9) & Proc_0_11_11029000(9, var_1D4, "cCode") & Chr(9) & 1100C008h & Chr(9) & Proc_0_11_11029000(9, var_1D4, "iMoney")
  loc_110E0831:   var_2E4 = var_2C
  loc_110E087C:   var_2EC = ADODB.Recordset.Fields
  loc_110E08A8:   ADODB.Recordset.8 = Forms
  loc_110E09B9:   var_D8 = var_108 & Chr(9) & Proc_0_11_11029000(9, var_1D4, "iQuantity") & Chr(9) & Proc_0_11_11029000(9, var_1D4, "iQuantity") & Chr(9) & Proc_0_11_11029000(9, var_1D4, "iQuantity") & Chr(9)
  loc_110E0ADD:   var_81B8 = var_D8 & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110E0CE3:   var_2E4 = var_2C
  loc_110E0D24:   var_2EC = ADODB.Recordset.Fields
  loc_110E0D5A:   ADODB.Recordset.8 = Forms
  loc_110E0DDA:   var_108 = var_81B8 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1D4, "cDepCode")
  loc_110E0EED:   var_2E4 = var_2C
  loc_110E0F2E:   var_2EC = ADODB.Recordset.Fields
  loc_110E0F64:   ADODB.Recordset.8 = Forms
  loc_110E1037:   var_1C0 = var_108 & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1D4, "cCusCode") & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1D4, "cCusCode")
  loc_110E117F:   var_2E4 = var_2C
  loc_110E11C0:   var_2EC = ADODB.Recordset.Fields
  loc_110E11F6:   ADODB.Recordset.8 = Forms
  loc_110E1276:   var_108 = var_1C0 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & Proc_0_11_11029000(9, var_1D4, "cInvCode")
  loc_110E1284:   var_48 = var_108
  loc_110E12DC:   var_2E4 = var_2C
  loc_110E12E2:   var_81FC = ADODB.Recordset.MoveNext
  loc_110E1358:   frmZGXSToPzTGZS.VFG.DispID_0080(var_48)
  loc_110E136D:   GoTo loc_110E010E
  loc_110E1372: End If
  loc_110E1395: var_2E4 = var_2C
  loc_110E139B: var_2E0 = ADODB.Recordset.State
  loc_110E13C6: If var_2E0 = 1 Then
  loc_110E13E4:   var_2E4 = var_2C
  loc_110E13EA:   var_8208 = ADODB.Recordset.Close
  loc_110E140E: End If
  loc_110E149A: var_2E4 = var_2C
  loc_110E14D7: var_8214 = ADODB.Recordset.Open("SELECT '21710112' AS cCode,SUM(ISNULL(iMoney1,0)-ISNULL(iMoney,0)) AS iMoney " & "FROM [T_CY_ZGXS_Temp] ", var_1C4, "SELECT '21710112' AS cCode,SUM(ISNULL(iMoney1,0)-ISNULL(iMoney,0)) AS iMoney " & "FROM [T_CY_ZGXS_Temp] ", var_1BC, 9)
  loc_110E151E: var_2E4 = var_2C
  loc_110E1524: var_2DC = ADODB.Recordset.EOF
  loc_110E154A: If var_2DC = 0 Then
  loc_110E1821:   var_E8 = "1" & Chr(9) & 1100AE28h & Chr(9) & frmZGXSToPzTGZS.TDBDate.DispID_004E & Chr(9) & 1100D6D4h & Chr(9) & 1100C008h & Chr(9) & var_44
  loc_110E1892:   var_2E4 = var_2C
  loc_110E18D3:   var_2EC = ADODB.Recordset.Fields
  loc_110E1909:   ADODB.Recordset.8 = Forms
  loc_110E1A9C:   var_2E4 = var_2C
  loc_110E1ADD:   var_2EC = ADODB.Recordset.Fields
  loc_110E1B13:   ADODB.Recordset.8 = Forms
  loc_110E1B37:   var_B4 = 0
  loc_110E1B41:   var_E0 = var_B4
  loc_110E1B93:   var_108 = 9 & Chr(9) & Proc_0_11_11029000(9, var_1D4, "cCode") & Chr(9) & 1100C008h & Chr(9) & Proc_0_11_11029000(9, var_1D4, "iMoney")
  loc_110E1EDD:   var_E8 = var_108 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110E2026:   var_1C0 = var_E8 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110E2185:   var_E8 = var_1C0 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110E22A3:   var_48 = var_E8 & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h & Chr(9) & 1100AE28h
  loc_110E22E1:   var_2E4 = var_2C
  loc_110E22E7:   var_8288 = ADODB.Recordset.MoveNext
  loc_110E235D:   frmZGXSToPzTGZS.VFG.DispID_0080(var_48)
  loc_110E2372:   GoTo loc_110E14FB
  loc_110E2377: End If
  loc_110E2435: frmZGXSToPzTGZS.sBar.DispID_6803001E("有效数据共" & CStr(var_28) & global_1100FE7C)
  loc_110E24A0: frmZGXSToPzTGZS.APB.UnkVCall_00000040h
  loc_110E2532: Set var_B0 = frmZGXSToPzTGZS.APB
  loc_110E2540: var_2E4 = var_B0
  loc_110E2546: var_B0.UnkVCall_00000040h
  loc_110E25D8: Set var_B0 = frmZGXSToPzTGZS.APB
  loc_110E25E6: var_2E4 = var_B0
  loc_110E25EC: var_B0.UnkVCall_00000040h
  loc_110E269C: frmZGXSToPzTGZS.Pic1.DispID_80010007 = var_1C0
  loc_110E2728: var_C0 = frmZGXSToPzTGZS.TDBText
  loc_110E2775: var_1C8 = var_4C.UnkVCall_0000006Ch
  loc_110E27AE: var_1C4 = var_58.UnkVCall_00000398h
  loc_110E27E3: Set var_38 = {000208D7-0000-0000-C000000000000046}()
  loc_110E27F3: Set var_4C = {000208DA-0000-0000-C000000000000046}()
  loc_110E2803: Set var_58 = {000208D5-0000-0000-C000000000000046}()
  loc_110E2817: GoTo loc_110E290D
  loc_110E290C: Exit Function
  loc_110E290D: ' Referenced from: 110E2817
End Function

Public Function getWBHL(sWhere) '110F3CF0
  Dim var_1C As ADODB.Recordset
  Dim var_2C As Me
  loc_110F3D50: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110F3D5C: var_98 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110F3D84: var_40 = Trim(sWhere)
  loc_110F3DB5: If (var_40 <> 1100AE28h) Then
  loc_110F3DE3:   var_20 = "SELECT * FROM exch WHERE 1=1 " & " AND " & sWhere
  loc_110F3DF0: Else
  loc_110F3DFC: End If
  loc_110F3E0C: var_20 = var_20 & " order by cexch_name, itype, iperiod, cdate"
  loc_110F3E76: var_78 = var_1C
  loc_110F3E85: var_8018 = ADODB.Recordset.Open(var_20, var_5C, var_20, var_54, 9)
  loc_110F3EEB: If ADODB.Recordset.EOF Then
  loc_110F3EFA:   var_24 = CStr(0)
  loc_110F3F05: Else
  loc_110F3F27:   var_2C = ADODB.Recordset.Fields
  loc_110F3F54:   var_58 = "NFLAT"
  loc_110F3F6D:   ADODB.Recordset.8 = Forms
  loc_110F3FBE:   var_24 = var_40
  loc_110F3FE0: End If
  loc_110F3FFE: var_8030 = ADODB.Recordset.Close
  loc_110F401D: GoTo loc_110F405B
  loc_110F4023: If var_4 Then
  loc_110F402E: End If
  loc_110F405A: Exit Function
  loc_110F405B: ' Referenced from: 110F401D
End Function

Private Sub Proc_17_7_110DB140
  Dim var_58 As frmZGXSToPzTGZS.VFG
  loc_110DB181: Set var_58 = frmZGXSToPzTGZS.VFG
  loc_110DB1D2: var_58.DispID_005D = frmZGXSToPzTGZS.VFG
  loc_110DB213: var_58.DispID_0067 = frmZGXSToPzTGZS.VFG
  loc_110DB232: var_58.DispID_0041 = frmZGXSToPzTGZS.VFG
  loc_110DB2DC: var_58.DispID_00A5("...")
  loc_110DB404: var_58.DispID_008A(4)
  loc_110DB447: var_58.DispID_0079(450)
  loc_110DB487: var_58.DispID_007B(True)
  loc_110DB4AB: var_58.DispID_0019 = True
  loc_110DB4F0: var_58.DispID_0090("业务号")
  loc_110DB533: var_58.DispID_0077(4)
  loc_110DB576: var_58.DispID_0078(700)
  loc_110DB5BE: var_58.DispID_0090("状态")
  loc_110DB604: var_58.DispID_0077(4)
  loc_110DB64A: var_58.DispID_0078(700)
  loc_110DB692: var_58.DispID_0090("制单日期")
  loc_110DB6D8: var_58.DispID_0077(1)
  loc_110DB71E: var_58.DispID_0078(1000)
  loc_110DB763: var_58.DispID_0090("凭证类别字")
  loc_110DB7A5: var_58.DispID_0077(4)
  loc_110DB7E7: var_58.DispID_0078(700)
  loc_110DB82F: var_58.DispID_0090("附单据数")
  loc_110DB873: var_58.DispID_0077(var_3C)
  loc_110DB8B9: var_58.DispID_0078(var_3C)
  loc_110DB901: var_58.DispID_0090(var_3C)
  loc_110DB947: var_58.DispID_0077(var_3C)
  loc_110DB98D: var_58.DispID_0078(var_3C)
  loc_110DB9D5: var_58.DispID_0090(var_3C)
  loc_110DBA1B: var_58.DispID_0077(var_3C)
  loc_110DBA61: var_58.DispID_0078(var_3C)
  loc_110DBAA9: var_58.DispID_0090(var_3C)
  loc_110DBAED: var_58.DispID_0077(var_3C)
  loc_110DBB33: var_58.DispID_0078(var_3C)
  loc_110DBB7B: var_58.DispID_009C(var_3C)
  loc_110DBBC3: var_58.DispID_0090(var_3C)
  loc_110DBC09: var_58.DispID_0077(var_3C)
  loc_110DBC4F: var_58.DispID_0078(var_3C)
  loc_110DBC97: var_58.DispID_009C(var_3C)
  loc_110DBCDF: var_58.DispID_0090(var_3C)
  loc_110DBD25: var_58.DispID_0077(var_3C)
  loc_110DBD6B: var_58.DispID_0078(var_3C)
  loc_110DBDB3: var_58.DispID_009C(var_3C)
  loc_110DBDFB: var_58.DispID_0090(var_3C)
  loc_110DBE41: var_58.DispID_0077(var_3C)
  loc_110DBE87: var_58.DispID_0078(var_3C)
  loc_110DBECF: var_58.DispID_009C(var_3C)
  loc_110DBF17: var_58.DispID_0090(var_3C)
  loc_110DBF5D: var_58.DispID_0077(var_3C)
  loc_110DBFA3: var_58.DispID_0078(var_3C)
  loc_110DBFEB: var_58.DispID_009C(var_3C)
  loc_110DC033: var_58.DispID_0090(var_3C)
  loc_110DC079: var_58.DispID_0077(var_3C)
  loc_110DC0BF: var_58.DispID_0078(var_3C)
  loc_110DC107: var_58.DispID_0090(var_3C)
  loc_110DC14D: var_58.DispID_0077(var_3C)
  loc_110DC193: var_58.DispID_0078(var_3C)
  loc_110DC1DB: var_58.DispID_0090(var_3C)
  loc_110DC221: var_58.DispID_0077(var_3C)
  loc_110DC267: var_58.DispID_0078(var_3C)
  loc_110DC2AF: var_58.DispID_0090(var_3C)
  loc_110DC2F5: var_58.DispID_0077(var_3C)
  loc_110DC33B: var_58.DispID_0078(var_3C)
  loc_110DC383: var_58.DispID_0090(var_3C)
  loc_110DC3C9: var_58.DispID_0077(var_3C)
  loc_110DC40F: var_58.DispID_0078(var_3C)
  loc_110DC457: var_58.DispID_0090(var_3C)
  loc_110DC49D: var_58.DispID_0077(var_3C)
  loc_110DC4E3: var_58.DispID_0078(var_3C)
  loc_110DC52B: var_58.DispID_0090(var_3C)
  loc_110DC571: var_58.DispID_0077(var_3C)
  loc_110DC5B7: var_58.DispID_0078(var_3C)
  loc_110DC5FF: var_58.DispID_0090(var_3C)
  loc_110DC645: var_58.DispID_0077(var_3C)
  loc_110DC68B: var_58.DispID_0078(var_3C)
  loc_110DC6D3: var_58.DispID_0090(var_3C)
  loc_110DC719: var_58.DispID_0077(var_3C)
  loc_110DC75F: var_58.DispID_0078(var_3C)
  loc_110DC7A7: var_58.DispID_0090(var_3C)
  loc_110DC7ED: var_58.DispID_0077(var_3C)
  loc_110DC833: var_58.DispID_0078(var_3C)
  loc_110DC87B: var_58.DispID_0090(var_3C)
  loc_110DC8C1: var_58.DispID_0077(var_3C)
  loc_110DC907: var_58.DispID_0078(var_3C)
  loc_110DC923: If 10 <= &H14 Then
  loc_110DC963:   var_58.DispID_00AC(var_3C)
  loc_110DC97B:   var_14 = 1+var_14
  loc_110DC97E:   GoTo loc_110DC91F
  loc_110DC980: End If
  loc_110DC9C0: var_58.DispID_00AC(var_3C)
  loc_110DCA05: var_58.DispID_00AC(var_3C)
  loc_110DCA4A: var_58.DispID_00AC(var_3C)
End Sub

Private Sub Proc_17_8_110E2FA0
  Dim var_7C As Variant
  Dim var_1F8 As Label
  Dim var_80 As Variant
  Dim var_88 As frmZGXSToPzTGZS.Label3
  loc_110E308A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110E3092: var_228 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110E3098: var_8004 = ecx
  loc_110E310E: If var_14 <= CLng(frmZGXSToPzTGZS.VFG.DispID_0007)(-1) Then
  loc_110E311F:   var_800C = frmZGXSToPzTGZS.Proc_17_9_110E4E40(vbNull)
  loc_110E31BD:   frmZGXSToPzTGZS.VFG.DispID_0082(22, var_58)
  loc_110E32A1:   If (frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 22) = global_1100AE28) + 1 Then
  loc_110E3321:     frmZGXSToPzTGZS.VFG.DispID_0082(1, 285267764)
  loc_110E3455:     frmZGXSToPzTGZS.VFG.DispID_009E(var_14, 1, var_14, 1, 16711680)
  loc_110E3475:     Set var_7C = frmZGXSToPzTGZS.Label3
  loc_110E3482:     var_1F8 = var_7C
  loc_110E34CC:     var_7C.Caption = "分析: 第(" & CStr(vbNull) & ")行信息----有效"
  loc_110E351E:     frmZGXSToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_110E3531:   Else
  loc_110E35AB:     frmZGXSToPzTGZS.VFG.DispID_0082(1, 285267820)
  loc_110E36DF:     frmZGXSToPzTGZS.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110E36FF:     Set var_80 = frmZGXSToPzTGZS.Label3
  loc_110E370C:     var_1F8 = var_80
  loc_110E37ED:     var_80.Caption = "分析:   第(" & CStr(vbNull) & ")行信息----" & frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 22)
  loc_110E3858:     frmZGXSToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_110E386A:   End If
  loc_110E387A:   var_14 = 1+var_14
  loc_110E387D:   GoTo loc_110E3100
  loc_110E3882: End If
  loc_110E38E9: If var_14 <= CLng(frmZGXSToPzTGZS.VFG.DispID_0007)(-1) Then
  loc_110E3961:   var_A0 = frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 2)
  loc_110E397F:   var_B8)
  loc_110E3B0F:   var_8048 = frmZGXSToPzTGZS.VFG.DispID_0082(var_14, frmZGXSToPzTGZS.VFG)
  loc_110E3B46:   var_4C = CCur(0)
  loc_110E3B49:   var_48 = var_8048
  loc_110E3B55:   var_40 = CCur(0)
  loc_110E3B58:   var_3C = var_8048
  loc_110E3B64:   var_34 = var_14
  loc_110E3B6D:   var_30 = var_14
  loc_110E3B76:   var_160 = CByte("DateToPeriod".00000001h)
  loc_110E3C13:   var_B8)
  loc_110E3C92:   Set var_80 = frmZGXSToPzTGZS.VFG
  loc_110E3CB8:   var_8064 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 3) = var_80.DispID_0082(var_14, 3))
  loc_110E3CE5:   var_1A0 = var_8064 + 1
  loc_110E3D5F:   var_806C = (var_8048 = frmZGXSToPzTGZS.VFG.DispID_0082(var_14, ""))
  loc_110E3D86:   var_1E0 = var_806C + 1
  loc_110E3E88:   If CBool((frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_8064 + 1 And var_806C + 1) Then
  loc_110E3F4D:     If (frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 22) = global_1100AE28) Then
  loc_110E3F56:     End If
  loc_110E3F5B:     If var_24 = 0 Then
  loc_110E4004:       var_16C = var_48
  loc_110E4048:       var_9C = var_1F0
  loc_110E4094:       var_4C = CCur(var_4C + Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 7)), "#.00"))
  loc_110E4097:       var_48 = var_D8
  loc_110E4177:       var_16C = var_3C
  loc_110E41BB:       var_9C = var_1F0
  loc_110E4207:       var_40 = CCur(var_40 + Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 8)), "#.00"))
  loc_110E420A:       var_3C = var_D8
  loc_110E424A:     End If
  loc_110E426B:     var_14 = var_14(1)
  loc_110E426E:     var_30 = var_30(1)
  loc_110E4290:     var_80A0 = CLng(frmZGXSToPzTGZS.VFG.DispID_0007)
  loc_110E42AB:     var_1F8 = (var_14 > 0)
  loc_110E42CF:     If var_1F8 = 0 Then GoTo loc_110E3B70
  loc_110E42D5:   End If
  loc_110E42DA:   If var_24 = 0 Then
  loc_110E42EE:     Set var_7C = frmZGXSToPzTGZS.Chk
  loc_110E42F9:     var_1F8 = var_7C
  loc_110E42FF:     Set var_80 = var_7C(1)
  loc_110E432A:     var_200 = var_80
  loc_110E4330:     var_1EC = var_80.Value
  loc_110E4384:     If (var_1EC = 1) Then
  loc_110E43B4:       If (Abs(var_4C - var_40) <> 0.01) >= 0 Then
  loc_110E43BD:       End If
  loc_110E43BD:     End If
  loc_110E43C2:     If var_24 Then
  loc_110E43C8:     End If
  loc_110E43E8:     var_1C = var_34
  loc_110E43ED:     If var_34 <= (var_30 - 1) Then
  loc_110E44B1:       If (frmZGXSToPzTGZS.VFG.DispID_0082(var_1C, 22) = global_1100AE28) + 1 Then
  loc_110E4539:         frmZGXSToPzTGZS.VFG.DispID_0082(1, 285267820)
  loc_110E45CD:         frmZGXSToPzTGZS.VFG.DispID_0082(22, "凭证借贷不平衡或某分录有错误")
  loc_110E4701:         frmZGXSToPzTGZS.VFG.DispID_009E(var_1C, 1, var_1C, 1, 255)
  loc_110E4713:       End If
  loc_110E4723:       GoTo loc_110E43E2
  loc_110E4728:     End If
  loc_110E4739:     var_44 = var_44(1)
  loc_110E474A:     Set var_88 = frmZGXSToPzTGZS.Label3
  loc_110E477D:     var_1F8 = var_88
  loc_110E488A:     Set var_80 = frmZGXSToPzTGZS.VFG
  loc_110E4962:     var_80D4 = "分析: 第[" & frmZGXSToPzTGZS.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmZGXSToPzTGZS.VFG.DispID_0082(var_34, var_14)
  loc_110E4984:     var_78 = var_80D4 & "]号凭证借贷不平衡"
  loc_110E4998:     var_88.Caption = var_78
  loc_110E499F:     If var_78 < 0 Then
  loc_110E49A5:       GoTo loc_110E4C23
  loc_110E49AA:     End If
  loc_110E49BB:     var_20 = var_20(1)
  loc_110E49CC:     Set var_88 = frmZGXSToPzTGZS.Label3
  loc_110E49FF:     var_1F8 = var_88
  loc_110E4B0C:     Set var_80 = frmZGXSToPzTGZS.VFG
  loc_110E4BE4:     var_80F8 = "分析: 第[" & frmZGXSToPzTGZS.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmZGXSToPzTGZS.VFG.DispID_0082(var_34, frmZGXSToPzTGZS.VFG.DispID_0082(var_34, var_14))
  loc_110E4C06:     var_78 = var_80F8 & "]号凭证有效"
  loc_110E4C1A:     var_88.Caption = var_78
  loc_110E4C21:     If var_78 >= 0 Then GoTo loc_110E4C32
  loc_110E4C23:     ' Referenced from: 110E49A5
  loc_110E4C2C:     var_78 = CheckObj(var_1F8, global_1100D574, 84)
  loc_110E4C32:   End If
  loc_110E4CB4:   frmZGXSToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_110E4CE5:   var_14 = 1+var_14(-1)
  loc_110E4CE8:   GoTo loc_110E38E3
  loc_110E4CED: End If
  loc_110E4CF2: If var_44 > 0 Then
  loc_110E4CF9:   If var_20 > 0 Then
  loc_110E4D14:   Else
  loc_110E4D2D:   Else
  loc_110E4D37:     var_8108 = frmZGXSToPzTGZS.Proc_17_11_110F40A0(var_1EC)
  loc_110E4D45:     If var_1EC Then
  loc_110E4D60:     Else
  loc_110E4D68:       var_18 = ecx
  loc_110E4D71:       GoTo loc_110E4E0B
  loc_110E4E0A:       Exit Sub
  loc_110E4E0B:     End If
  loc_110E4E0B:   End If
  loc_110E4E0B: End If
  loc_110E4E0B: ' Referenced from: 110E4D71
End Sub

Private  Proc_17_9_110E4E40(arg_C) '110E4E40
  Dim var_58 As frmZGXSToPzTGZS.VFG
  Dim var_20 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA19-EB96-11D2-AF06000021009B21}()
  Dim var_18 As {3302AA41-EB96-11D2-AF06000021009B21}()
  Dim var_1C As {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_110E4F3C: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110E4F4C: var_210 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110E502B: If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 2) = global_1100AE28) + 1 Then
  loc_110E5035:   var_24 = "制单日期为空"
  loc_110E5046: Else
  loc_110E50E1:   var_78 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 2)
  loc_110E511B:   If Proc_0_9_11028500(var_80, global_110EA50C, ) Then
  loc_110E51C4:     var_78 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 2)
  loc_110E51CE:     var_90)
  loc_110E51E0:     var_48 = var_90
  loc_110E5212:     var_118 = var_48
  loc_110E5220:     var_114 = var_44
  loc_110E5254:     var_80 = "AccountOpen".0.0
  loc_110E5285:     If (var_80 < var_80) Then
  loc_110E528F:       var_24 = "日期超前总账系统启用日期"
  loc_110E52A0:     Else
  loc_110E52A6:       var_154 = var_44
  loc_110E52AC:       var_1A4 = var_44
  loc_110E52B8:       var_158 = var_48
  loc_110E52BF:       var_1A8 = var_48
  loc_110E536C:       var_80 = "AccountYMD".0.00000002h("AccountYMD".0, var_13C)
  loc_110E5466:       If CBool( Or ((global_110EA50C < var_80) > "AccountYMD".0.00000002h(var_180, var_18C))) Then
  loc_110E5470:         var_24 = "日期必须在当前会计年度内"
  loc_110E5481:       Else
  loc_110E549E:         var_118 = var_48
  loc_110E54F2:         var_80 = "DateToPeriod".00000001h - 1
  loc_110E5580:         If CBool("AccountYMD".0.00000001h) Then
  loc_110E558A:           var_24 = "已结账月份不能制单"
  loc_110E559B:         Else
  loc_110E5677:           If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 3) = global_1100AE28) + 1 Then
  loc_110E5681:             var_24 = "凭证类别字为空"
  loc_110E5692:           Else
  loc_110E5721:             var_8034 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 3)
  loc_110E5731:             var_80 = 8
  loc_110E5734:             var_78 = var_8034
  loc_110E577B:             var_8038 = CBool(Not("pzlbCheck".00000001h(, fs:[00000000h], , global_110EA50C, global_110EA50C, var_74, var_8034, var_7C)))
  loc_110E57B2:             If var_8038 Then
  loc_110E57BC:               var_24 = "凭证类别字非法"
  loc_110E57CD:             Else
  loc_110E58A4:               If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, var_128) = global_1100AE28) + 1 Then
  loc_110E58AE:                 var_24 = "业务号为空"
  loc_110E58BF:               Else
  loc_110E5949:                 var_8044 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, var_128)
  loc_110E5959:                 var_80 = 8
  loc_110E595C:                 var_78 = var_8044
  loc_110E599F:                 var_90 = "GenLen".00000001h(fs:[00000000h], , global_110EA50C, global_110EA50C, global_110EA50C, var_74, var_8044, var_7C)
  loc_110E59E7:                 If (var_90 > 30) Then
  loc_110E59F1:                   var_24 = "业务号超长"
  loc_110E5A02:                 Else
  loc_110E5AE1:                   If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 5) = global_1100AE28) + 1 Then
  loc_110E5AEB:                     var_24 = "摘要为空"
  loc_110E5AFC:                   Else
  loc_110E5BB7:                     var_8058 = InStr(1, frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 5), "|", 0)
  loc_110E5BDD:                     var_220 = (var_8058 > 0)
  loc_110E5C33:                     var_80 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 5)
  loc_110E5D54:                     If (((var_8058 > 0) Or (InStr(1, var_80, """", 0) > 0)) Or (InStr(1, frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 5), "'", 0) > 0)) Then
  loc_110E5D5E:                       var_24 = "摘要含有非法字符"
  loc_110E5D6F:                     Else
  loc_110E5E01:                       var_806C = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 5)
  loc_110E5E11:                       var_80 = 8
  loc_110E5E14:                       var_78 = var_806C
  loc_110E5E57:                       var_90 = "GenLen".00000001h(global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, var_74, var_806C, var_7C)
  loc_110E5EA0:                       If (var_90 > 120) Then
  loc_110E5EAA:                         var_24 = "摘要超长"
  loc_110E5EBB:                       Else
  loc_110E5F98:                         If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 6) = global_1100AE28) + 1 Then
  loc_110E5FA2:                           var_24 = "科目为空"
  loc_110E5FB3:                         Else
  loc_110E6042:                           var_807C = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 6)
  loc_110E6052:                           var_80 = 8
  loc_110E6055:                           var_78 = var_807C
  loc_110E60D5:                           var_40 = "kmCheck".00000002h(var_807C, var_150, var_15C)
  loc_110E6107:                           var_8084 = (var_40 = global_1100AE28)
  loc_110E610F:                           If var_8084 = 0 Then
  loc_110E6119:                             var_24 = "科目非法"
  loc_110E612A:                           Else
  loc_110E6168:                             var_118 = arg_C
  loc_110E61CF:                             frmZGXSToPzTGZS.VFG.DispID_0082(6, var_40)
  loc_110E61E9:                             var_118 = var_40
  loc_110E623B:                             var_128 = var_20
  loc_110E6289:                             "kmCodeToProperties".00000002h
  loc_110E62A6:                             Set var_20 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110E62D4:                             var_1F0 = var_20
  loc_110E62DA:                             var_1D4 = var_20.UnkVCall_00000114h
  loc_110E6306:                             If var_1D4 = 0 Then
  loc_110E6310:                               var_24 = "科目非末级"
  loc_110E6321:                             Else
  loc_110E63FF:                               If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 7) = global_1100AE28) Then
  loc_110E64DB:                                 If Not (IsNumeric(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 7))) Then
  loc_110E64E5:                                   var_24 = "借方金额非法"
  loc_110E64F6:                                 Else
  loc_110E659F:                                   var_80A4 = CDbl(Val(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 7)))
  loc_110E663A:                                   var_80 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 7)
  loc_110E6662:                                   var_22C = CDbl(Val(var_80))
  loc_110E6678:                                   var_80B0 = CDbl(-9999999999999.99)
  loc_110E6690:                                   GoTo loc_110E6694
  loc_110E66E2:                                   If (eax Or 0) Then
  loc_110E66EC:                                     var_24 = "借方金额超范围"
  loc_110E66FD:                                   Else
  loc_110E66FD:                                   End If
  loc_110E67DB:                                   If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 8) = global_1100AE28) Then
  loc_110E68B7:                                     If Not (IsNumeric(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 8))) Then
  loc_110E68C1:                                       var_24 = "贷方金额非法"
  loc_110E68D2:                                     Else
  loc_110E697B:                                       var_80C8 = CDbl(Val(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 8)))
  loc_110E6A16:                                       var_80 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 8)
  loc_110E6A3E:                                       var_238 = CDbl(Val(var_80))
  loc_110E6A54:                                       var_80D4 = CDbl(-9999999999999.99)
  loc_110E6A6C:                                       GoTo loc_110E6A70
  loc_110E6ABE:                                       If (eax Or 0) Then
  loc_110E6AC8:                                         var_24 = "贷方金额超范围"
  loc_110E6AD9:                                       Else
  loc_110E6AD9:                                       End If
  loc_110E6C51:                                       var_74 = var_1E0
  loc_110E6CC3:                                       var_C4 = var_1E8
  loc_110E6D3D:                                       var_80E8 = (Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 7)), "#.00") <> 0) And (Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 8)), "#.00") <> 0)
  loc_110E6DB6:                                       If CBool(var_80E8) Then
  loc_110E6DC0:                                         var_24 = "借方金额和贷方金额不能同时不为0"
  loc_110E6DD1:                                       Else
  loc_110E6F49:                                         var_74 = var_1E0
  loc_110E6FBB:                                         var_C4 = var_1E8
  loc_110E7035:                                         var_8100 = (Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 7)), "#.00") = 0) And (Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 8)), "#.00") = 0)
  loc_110E70AE:                                         If CBool(var_8100) Then
  loc_110E70B8:                                           var_24 = "借方金额和贷方金额不能同时为0"
  loc_110E70C9:                                         Else
  loc_110E70E9:                                           var_1F0 = var_20
  loc_110E713B:                                           If (var_20.UnkVCall_0000007Ch = global_1100AE28) Then
  loc_110E71F8:                                             var_1F0 = (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 9) = global_1100AE28)
  loc_110E721F:                                             If var_1F0 = 0 Then GoTo loc_110E73D5
  loc_110E72D3:                                             var_1F0 = Not (IsNumeric(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 9)))
  loc_110E72FB:                                             If var_1F0 = 0 Then GoTo loc_110E73D5
  loc_110E7309:                                             var_24 = "数量数值非法"
  loc_110E731A:                                           Else
  loc_110E7337:                                             var_118 = arg_C
  loc_110E73C3:                                             frmZGXSToPzTGZS.VFG.DispID_0082(9, 285257256)
  loc_110E73F5:                                             var_1F0 = var_20
  loc_110E7447:                                             If (var_20.UnkVCall_0000006Ch = global_1100AE28) Then
  loc_110E752B:                                               If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 10) = global_1100AE28) Then
  loc_110E7607:                                                 If Not (IsNumeric(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 10))) Then
  loc_110E7611:                                                   var_24 = "外币金额非法"
  loc_110E7622:                                                 Else
  loc_110E76CB:                                                   var_813C = CDbl(Val(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 10)))
  loc_110E778E:                                                   var_244 = CDbl(Val(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 10)))
  loc_110E77A4:                                                   var_8148 = CDbl(-9999999999999.99)
  loc_110E77BC:                                                   GoTo loc_110E77C0
  loc_110E780E:                                                   If (eax Or 0) Then
  loc_110E7818:                                                     var_24 = "外币超范围"
  loc_110E7829:                                                   Else
  loc_110E7829:                                                   End If
  loc_110E7907:                                                   If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 11) = global_1100AE28) Then
  loc_110E79E3:                                                     If Not (IsNumeric(frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 11))) Then
  loc_110E79ED:                                                       var_24 = "汇率数值非法"
  loc_110E79FE:                                                     Else
  loc_110E79FE:                                                     End If
  loc_110E79FE:                                                   End If
  loc_110E7ADC:                                                   If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 12) = global_1100AE28) Then
  loc_110E7B73:                                                     var_8164 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 12)
  loc_110E7B86:                                                     var_78 = var_8164
  loc_110E7BC9:                                                     var_90 = "GenLen".00000001h(global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, var_74, var_8164, var_7C)
  loc_110E7BE3:                                                     var_1F0 = (var_90 > 20)
  loc_110E7C12:                                                     If var_1F0 = 0 Then GoTo loc_110E7D58
  loc_110E7C20:                                                     var_24 = "制单人姓名超长"
  loc_110E7C31:                                                   Else
  loc_110E7C50:                                                     var_118 = arg_C
  loc_110E7D2C:                                                     frmZGXSToPzTGZS.VFG.DispID_0082(12, "UserCurrent".00000000h.00000000h)
  loc_110E7D7B:                                                     var_1F0 = var_20
  loc_110E7DAD:                                                     If var_20.UnkVCall_0000010Ch Then
  loc_110E7E91:                                                       If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 13) = global_1100AE28) Then
  loc_110E7F28:                                                         var_817C = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 13)
  loc_110E7F3B:                                                         var_78 = var_817C
  loc_110E7F6A:                                                         var_90 = "JsfsCheck".00000001h(1, global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, var_74, var_817C, var_7C)
  loc_110E7FBA:                                                         If CBool(Not(var_90)) Then
  loc_110E7FC4:                                                           var_24 = "结算方式非法"
  loc_110E7FD5:                                                         Else
  loc_110E7FD5:                                                         End If
  loc_110E7FD5:                                                       End If
  loc_110E7FF8:                                                       var_1F0 = var_20
  loc_110E7FFE:                                                       var_1D4 = var_20.UnkVCall_0000010Ch
  loc_110E8045:                                                       var_1F8 = var_20
  loc_110E804B:                                                       var_1D8 = var_20.UnkVCall_00000094h
  loc_110E8092:                                                       var_200 = var_20
  loc_110E80EA:                                                       If (var_20.UnkVCall_0000009Ch = 0) = 0 Then
  loc_110E81CE:                                                         If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 14) = global_1100AE28) Then
  loc_110E8265:                                                           var_8198 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 14)
  loc_110E8278:                                                           var_78 = var_8198
  loc_110E82BB:                                                           var_90 = "GenLen".00000001h(1, global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, var_74, var_8198, var_7C)
  loc_110E8304:                                                           If (var_90 > 10) Then
  loc_110E830E:                                                             var_24 = "票号超长"
  loc_110E831F:                                                           Else
  loc_110E831F:                                                           End If
  loc_110E83FD:                                                           If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 15) = global_1100AE28) Then
  loc_110E8494:                                                             var_81A8 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 15)
  loc_110E84A7:                                                             var_78 = var_81A8
  loc_110E84D6:                                                             var_90 = "DateCheck".00000001h(1, global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, var_74, var_81A8, var_7C)
  loc_110E8526:                                                             If CBool(Not(var_90)) Then
  loc_110E8530:                                                               var_24 = "票号发生日期非法"
  loc_110E8541:                                                             Else
  loc_110E8541:                                                             End If
  loc_110E8541:                                                           End If
  loc_110E8564:                                                           var_1F0 = var_20
  loc_110E85B1:                                                           var_1F8 = var_20
  loc_110E85B7:                                                           var_1D8 = var_20.UnkVCall_0000008Ch
  loc_110E861A:                                                           If (var_20.UnkVCall_000000A4h = 0) = 0 Then
  loc_110E86D9:                                                             If (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 16) = global_1100AE28) Then
  loc_110E8783:                                                               var_78 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 16)
  loc_110E8803:                                                               var_38 = "BmCheck".00000002h(var_154, 0, var_15C)
  loc_110E8835:                                                               var_81C8 = (var_38 = global_1100AE28)
  loc_110E883D:                                                               If var_81C8 = 0 Then
  loc_110E8847:                                                                 var_24 = "部门非法"
  loc_110E8858:                                                               Else
  loc_110E8875:                                                                 var_118 = arg_C
  loc_110E88FF:                                                                 frmZGXSToPzTGZS.VFG.DispID_0082(16, var_38)
  loc_110E8934:                                                                 var_1F0 = var_20
  loc_110E8966:                                                                 If var_20.UnkVCall_000000A4h Then
  loc_110E8974:                                                                   var_118 = var_38
  loc_110E89C6:                                                                   var_128 = var_28
  loc_110E8A14:                                                                   "BmToProperties".00000002h
  loc_110E8A31:                                                                   Set var_28 = {3302AA19-EB96-11D2-AF06000021009B21}()
  loc_110E8A5F:                                                                   var_1F0 = var_28
  loc_110E8A65:                                                                   var_1D4 = var_28.UnkVCall_00000034h
  loc_110E8A8B:                                                                   If var_1D4 = 0 Then
  loc_110E8A99:                                                                     var_24 = "部门非末级"
  loc_110E8AAA:                                                                   Else
  loc_110E8AB2:                                                                     var_24 = "部门为空"
  loc_110E8AC3:                                                                   Else
  loc_110E8B45:                                                                     frmZGXSToPzTGZS.VFG.DispID_0082(var_128, 285257256)
  loc_110E8B57:                                                                   End If
  loc_110E8B57:                                                                 End If
  loc_110E8B7A:                                                                 var_1F0 = var_20
  loc_110E8BAC:                                                                 If var_20.UnkVCall_0000008Ch Then
  loc_110E8C58:                                                                   var_81E0 = (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H11) = global_1100AE28)
  loc_110E8C90:                                                                   If var_81E0 Then
  loc_110E8D3C:                                                                     var_81E8 = (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, 16) = global_1100AE28)
  loc_110E8D98:                                                                     If var_81E8 + 1 Then
  loc_110E8E1B:                                                                       var_78 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H11)
  loc_110E8EA9:                                                                       var_90 = "ZyCheck".00000003h(var_174, "BmCheck".00000002h(var_154, 80020004h, var_15C), var_17C)
  loc_110E8EBE:                                                                       var_34 = var_90
  loc_110E8EF0:                                                                       var_81F4 = (var_34 = global_1100AE28)
  loc_110E8EF8:                                                                       If var_81F4 = 0 Then
  loc_110E8F02:                                                                         var_24 = "职员非法"
  loc_110E8F13:                                                                       Else
  loc_110E8F30:                                                                         var_118 = arg_C
  loc_110E8FBA:                                                                         frmZGXSToPzTGZS.VFG.DispID_0082(&H11, var_34)
  loc_110E8FD9:                                                                         var_118 = var_34
  loc_110E9026:                                                                         var_128 = var_18
  loc_110E9074:                                                                         "ZyToProperties".00000002h
  loc_110E9091:                                                                         Set var_18 = {3302AA41-EB96-11D2-AF06000021009B21}()
  loc_110E909F:                                                                         var_118 = arg_C
  loc_110E90E0:                                                                         var_1F0 = var_18
  loc_110E9199:                                                                         frmZGXSToPzTGZS.VFG.DispID_0082(var_128, var_18.UnkVCall_0000002Ch)
  loc_110E91B9:                                                                       Else
  loc_110E922F:                                                                         var_158 = var_38
  loc_110E923C:                                                                         var_78 = frmZGXSToPzTGZS.VFG.DispID_0082(8, var_128)
  loc_110E92F1:                                                                         var_34 = "ZyCheck".00000003h(var_164, 0, var_16C)
  loc_110E9323:                                                                         var_8208 = (var_34 = global_1100AE28)
  loc_110E932B:                                                                         If var_8208 = 0 Then
  loc_110E9335:                                                                           var_24 = "职员不在指定部门内"
  loc_110E9346:                                                                         Else
  loc_110E9384:                                                                           var_118 = arg_C
  loc_110E93EB:                                                                           frmZGXSToPzTGZS.VFG.DispID_0082(&H11, var_34)
  loc_110E93FD:                                                                         End If
  loc_110E93FD:                                                                       End If
  loc_110E93FD:                                                                     End If
  loc_110E9420:                                                                     var_1F0 = var_20
  loc_110E9452:                                                                     If var_20.UnkVCall_00000094h Then
  loc_110E94FE:                                                                       var_8214 = (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H12) = global_1100AE28)
  loc_110E950F:                                                                       var_1F0 = var_8214
  loc_110E9536:                                                                       If var_1F0 = 0 Then GoTo loc_110E9A24
  loc_110E95E0:                                                                       var_78 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H12)
  loc_110E9660:                                                                       var_3C = "KhCheck".00000002h(var_154, 0, var_15C)
  loc_110E9692:                                                                       var_8220 = (var_3C = global_1100AE28)
  loc_110E969A:                                                                       If var_8220 = 0 Then
  loc_110E96A4:                                                                         var_24 = "客户非法"
  loc_110E96B5:                                                                       Else
  loc_110E96F3:                                                                         var_118 = arg_C
  loc_110E975A:                                                                         frmZGXSToPzTGZS.VFG.DispID_0082(&H12, var_3C)
  loc_110E976C:                                                                       End If
  loc_110E978F:                                                                       var_1F0 = var_20
  loc_110E97C1:                                                                       If var_20.UnkVCall_0000009Ch Then
  loc_110E986D:                                                                         var_822C = (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H13) = global_1100AE28)
  loc_110E987E:                                                                         var_1F0 = var_822C
  loc_110E98A5:                                                                         If var_1F0 = 0 Then GoTo loc_110E9DDD
  loc_110E994F:                                                                         var_78 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H13)
  loc_110E99CF:                                                                         var_30 = "GysCheck".00000002h(var_154, 0, var_15C)
  loc_110E9A01:                                                                         var_8238 = (var_30 = global_1100AE28)
  loc_110E9A09:                                                                         If var_8238 = 0 Then
  loc_110E9A13:                                                                           var_24 = "供应商非法"
  loc_110E9A1F:                                                                           GoTo loc_110EA4CD
  loc_110E9A2C:                                                                           var_24 = "客户为空"
  loc_110E9A3D:                                                                         Else
  loc_110E9A7B:                                                                           var_118 = arg_C
  loc_110E9AE2:                                                                           frmZGXSToPzTGZS.VFG.DispID_0082(&H13, var_30)
  loc_110E9AF4:                                                                         End If
  loc_110E9B17:                                                                         var_1F0 = var_20
  loc_110E9B64:                                                                         var_1F8 = var_20
  loc_110E9B6A:                                                                         var_1D8 = var_20.UnkVCall_0000009Ch
  loc_110E9BA8:                                                                         If (var_20.UnkVCall_00000094h = 0) = 0 Then
  loc_110E9C54:                                                                           var_8248 = (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H14) = global_1100AE28)
  loc_110E9C8C:                                                                           If var_8248 Then
  loc_110E9D23:                                                                             var_824C = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H14)
  loc_110E9D36:                                                                             var_78 = var_824C
  loc_110E9D79:                                                                             var_90 = "GenLen".00000001h(global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, global_110EA50C, var_74, var_824C, var_7C)
  loc_110E9DC2:                                                                             If (var_90 > 20) Then
  loc_110E9DCC:                                                                               var_24 = "业务员超长"
  loc_110E9DD8:                                                                               GoTo loc_110EA4CD
  loc_110E9DE5:                                                                               var_24 = "供应商为空"
  loc_110E9DF6:                                                                             Else
  loc_110E9DF6:                                                                             End If
  loc_110E9DF6:                                                                           End If
  loc_110E9E16:                                                                           var_1F0 = var_20
  loc_110E9E6E:                                                                           If (var_20.UnkVCall_000000ACh = global_1100AE28) Then
  loc_110E9F1A:                                                                             var_8260 = (frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H15) = global_1100AE28)
  loc_110E9F52:                                                                             If var_8260 Then
  loc_110E9F78:                                                                               var_1F0 = var_20
  loc_110E9FAB:                                                                               var_8268 = (var_20.UnkVCall_000000ACh = global_1100AE28)
  loc_110E9FD0:                                                                               If var_8268 Then
  loc_110E9FF6:                                                                                 var_1F0 = var_20
  loc_110EA026:                                                                                 var_78 = var_20.UnkVCall_000000ACh
  loc_110EA0BE:                                                                                 var_8270 = frmZGXSToPzTGZS.VFG.DispID_0082(arg_C, &H15)
  loc_110EA0D4:                                                                                 var_88 = var_8270
  loc_110EA15C:                                                                                 var_A0 = "XmCheck".00000003h(var_164, Not(8), var_16C)
  loc_110EA171:                                                                                 var_2C = var_A0
  loc_110EA1AA:                                                                                 var_8278 = (var_2C = global_1100AE28)
  loc_110EA1B2:                                                                                 If var_8278 = 0 Then
  loc_110EA1BC:                                                                                   var_24 = "项目非法"
  loc_110EA1CD:                                                                                 Else
  loc_110EA1F9:                                                                                   var_4C = var_20.UnkVCall_000000ACh
  loc_110EA227:                                                                                   var_128 = var_2C
  loc_110EA258:                                                                                   Set var_58 = var_1C
  loc_110EA2DA:                                                                                   "XmToProperties".00000003h
  loc_110EA2F7:                                                                                   Set var_1C = {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_110EA34C:                                                                                   If var_1C.UnkVCall_00000034h Then
  loc_110EA35A:                                                                                     var_24 = "项目已结算"
  loc_110EA36B:                                                                                   Else
  loc_110EA395:                                                                                     var_118 = %cobj
  loc_110EA40D:                                                                                     frmZGXSToPzTGZS.VFG.DispID_0082(&H15, 285257256)
  loc_110EA42A:                                                                                   Else
  loc_110EA432:                                                                                     var_24 = "制单日期非法"
  loc_110EA438:                                                                                   End If
  loc_110EA438:                                                                                 End If
  loc_110EA438:                                                                               End If
  loc_110EA43E:                                                                               GoTo loc_110EA4CD
  loc_110EA447:                                                                               If var_4 Then
  loc_110EA452:                                                                               End If
  loc_110EA4CC:                                                                               Exit Sub
  loc_110EA4CD:                                                                             End If
  loc_110EA4CD:                                                                           End If
  loc_110EA4CD:                                                                         End If
  loc_110EA4CD:                                                                       End If
  loc_110EA4CD:                                                                     End If
  loc_110EA4CD:                                                                   End If
  loc_110EA4CD:                                                                 End If
  loc_110EA4CD:                                                               End If
  loc_110EA4CD:                                                             End If
  loc_110EA4CD:                                                           End If
  loc_110EA4CD:                                                         End If
  loc_110EA4CD:                                                       End If
  loc_110EA4CD:                                                     End If
  loc_110EA4CD:                                                   End If
  loc_110EA4CD:                                                 End If
  loc_110EA4CD:                                               End If
  loc_110EA4CD:                                             End If
  loc_110EA4CD:                                           End If
  loc_110EA4CD:                                         End If
  loc_110EA4CD:                                       End If
  loc_110EA4CD:                                     End If
  loc_110EA4CD:                                   End If
  loc_110EA4CD:                                 End If
  loc_110EA4CD:                               End If
  loc_110EA4CD:                             End If
  loc_110EA4CD:                           End If
  loc_110EA4CD:                         End If
  loc_110EA4CD:                       End If
  loc_110EA4CD:                     End If
  loc_110EA4CD:                   End If
  loc_110EA4CD:                 End If
  loc_110EA4CD:               End If
  loc_110EA4CD:             End If
  loc_110EA4CD:           End If
  loc_110EA4CD:         End If
  loc_110EA4CD:       End If
  loc_110EA4CD:     End If
  loc_110EA4CD:   End If
  loc_110EA4CD: End If
  loc_110EA4CD: ' Referenced from: 110EA43E
End Sub

Private Sub Proc_17_10_110EA530
  Dim var_9C As Variant
  Dim var_8034 As Label
  Dim var_8074 As Label
  Dim var_A0 As Variant
  Dim var_38 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA47-EB96-11D2-AF06000021009B21}()
  loc_110EA68A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110EA690: var_294 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110EA6B6: Set var_9C = frmZGXSToPzTGZS.VFG
  loc_110EA700: If (CLng(var_9C.DispID_0007) < 2) Then
  loc_110EA72E:   var_800C = = Global.Screen
  loc_110EA750:   var_8010 = ecx
  loc_110EA758:   var_8010 = var_9C.UnkVCall_0000007Ch
  loc_110EA7C5:   var_C8 = "提示信息"
  loc_110EA7C7:   var_150 = "没有可生成用友凭证的数据。"
  loc_110EA7D6: Else
  loc_110EA886:   var_264 = ("GetAccInfo".00000002h(, , fs:[00000000h], , "GL", var_16C, "dGLStartDate", var_174) = 1100AE28h)
  loc_110EA8A0:   If var_264 = 0 Then GoTo loc_110EA9E1
  loc_110EA8CE:   var_801C = = Global.Screen
  loc_110EA8F0:   var_8020 = ecx
  loc_110EA8F8:   var_8020 = var_9C.UnkVCall_0000007Ch
  loc_110EA965:   var_C8 = "提示信息"
  loc_110EA967:   var_150 = "总账系统尚未启用，不能进行凭证引入！"
  loc_110EA971: End If
  loc_110EA9A3: MsgBox(var_150, 64, var_C8, var_D8, var_E8)
  loc_110EA9D0: Exit Sub
  loc_110EA9DC: GoTo loc_110F35C1
  loc_110EA9EB: var_8024 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[VouchNum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " & " CREATE TABLE VouchNum(iperiod tinyint NULL ,csign varchar(8) NULL ,ino_id int NULL,constraint index1 unique(iperiod,csign,ino_id))"
  loc_110EA9F1: var_B0 = var_8024
  loc_110EAA50: var_D8.00000001h(0, , , , "3缷M鋎?", var_AC, var_8024, var_B4)
  loc_110EAA70: On Error GoTo 0
  loc_110EAA76: var_B0 = %ecx = %S_edx_S
  loc_110EAA98: var_78 = "AS13"
  loc_110EAAB0: var_78)
  loc_110EAADA: If Not (var_78)) Then
  loc_110EAB0B:   If Global.Screen < 0 Then
  loc_110EAB1C:   End If
  loc_110EAB26:   var_8030 = ecx
  loc_110EAB35:   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110EAB48:   Else
  loc_110EAB59:     call var_8034 = var_9C(var_9C, frmZGXSToPzTGZS.Label3, var_9C, global_1100C47C, 0000007Ch)
  loc_110EAB5B:     var_264 = var_8034
  loc_110EAB69:     Label3.Caption = "正在进行数据分析，请稍等..."
  loc_110EAB96:     var_150 = True
  loc_110EABD9:     call var_8038 = var_9C(var_9C, frmZGXSToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, True, var_14C)
  loc_110EABDC:     var_8038.DispID_0000 =
  loc_110EAC05:     call var_803C = var_9C(var_9C, frmZGXSToPzTGZS.Pic1, global_FFFFFDDA, var_9C = var_9C)
  loc_110EAC08:     var_803C.DispID_0000
  loc_110EAC27:     var_8040 = .Proc_17_8_110E2FA0(var_24C)
  loc_110EAC35:     If var_24C = 2 Then
  loc_110EAC3B:       var_150 = %ecx = %S_edx_S
  loc_110EAC7E:       call var_8044 = var_9C(var_9C, frmZGXSToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C)
  loc_110EAC81:       var_8044.DispID_0000 =
  loc_110EAD1D:       MsgBox("数据源中没有合法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110EAD5A:       var_24C = %ecx = %S_edx_S
  loc_110EAD80:       "AS13")
  loc_110EADC2:       var_B8 = Global.Screen
  loc_110EADE4:       var_804C = ecx
  loc_110EADF3:       If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110EAE06:       Else
  loc_110EAE08:         If var_804C = 1 Then
  loc_110EAE0E:           var_150 = %ecx = %S_edx_S
  loc_110EAE51:           call var_8050 = var_9C(var_9C, frmZGXSToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, var_14C = var_9C, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_110EAE54:           var_8050.DispID_0000 =
  loc_110EAEF0:           MsgBox("数据源中含有非法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110EAF2D:           var_24C = %ecx = %S_edx_S
  loc_110EAF53:           "AS13")
  loc_110EAF95:           var_B8 = Global.Screen
  loc_110EAFB7:           var_8058 = ecx
  loc_110EAFC6:           If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110EAFD9:           Else
  loc_110EAFDB:             If var_8058 = 3 Then
  loc_110EAFE1:               var_150 = %ecx = %S_edx_S
  loc_110EB024:               call var_805C = var_9C(var_9C, frmZGXSToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_110EB027:               var_805C.DispID_0000 =
  loc_110EB0C3:               MsgBox("数据源中指定的凭证号无效或重号，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110EB100:               var_24C = %ecx = %S_edx_S
  loc_110EB126:               "AS13")
  loc_110EB168:               var_B8 = Global.Screen
  loc_110EB18A:               var_8064 = ecx
  loc_110EB199:               If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110EB1AC:               Else
  loc_110EB1EE:                 var_C8 = "提示信息"
  loc_110EB214:                 var_B8 = "数据源中的数据已全部通过检查，是否开始引入？"
  loc_110EB238:                 MsgBox(var_B8, 36, var_C8, var_D8, var_E8)
  loc_110EB27D:                 If (MsgBox(var_B8, 36, var_C8, var_D8, var_E8) = 7) Then
  loc_110EB2C8:                   call var_8068 = var_9C(var_9C, frmZGXSToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, frmZGXSToPzTGZS.Pic1, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_110EB2CB:                   var_8068.DispID_0000 =
  loc_110EB2F1:                   var_24C = %ecx = %S_edx_S
  loc_110EB317:                   "AS13")
  loc_110EB359:                   var_B8 = Global.Screen
  loc_110EB37B:                   var_8070 = ecx
  loc_110EB38A:                   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110EB39D:                   Else
  loc_110EB39E:                     On Error GoTo 0
  loc_110EB3B5:                     call var_8074 = var_9C(var_9C, frmZGXSToPzTGZS.Label3, var_9C = var_9C, var_9C, global_1100C47C, 0000007Ch)
  loc_110EB3B7:                     var_264 = var_8074
  loc_110EB3C5:                     Label3.Caption = "正在写数据，请稍等..."
  loc_110EB409:                     call var_8078 = var_9C(var_9C, frmZGXSToPzTGZS.Pic1, global_FFFFFDDA, 00000000h)
  loc_110EB40C:                     var_8078.DispID_0000
  loc_110EB443:                     Set var_74 = CreateObject("UfDbKit.UfRecordset", 0)
  loc_110EB45A:                     var_150 = "SELECT TOP 1 * FROM GL_accvouch"
  loc_110EB4CF:                     Set var_74 = "DataMdb".00000000h.00000001h(var_14C, "SELECT TOP 1 * FROM GL_accvouch", var_154)
  loc_110EB503:                     call var_8084 = var_9C(var_9C, frmZGXSToPzTGZS.VFG, 00000007h, 00000000h)
  loc_110EB567:                     If var_24 <= CLng(var_8084.DispID_0000)(-1) Then
  loc_110EB571:                       var_2A8 = var_24
  loc_110EB577:                       var_150 = var_24
  loc_110EB5F4:                       call var_8090 = var_9C(var_9C, frmZGXSToPzTGZS.VFG, 00000082h, 00000002h, 3, var_174, 2, var_16C, 00000003h, var_154, var_24, var_14C)
  loc_110EB60E:                       var_C0 = var_8090.DispID_0000
  loc_110EB62C:                       var_D8)
  loc_110EB684:                       var_70 = CByte("DateToPeriod".00000001h(8, var_D4))
  loc_110EB6BD:                       var_150 = var_2A8
  loc_110EB736:                       call var_809C = var_9C(var_9C, frmZGXSToPzTGZS.VFG, 00000082h, 00000002h, 3, var_174, 3, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110EB755:                       var_58 = var_809C.DispID_0000
  loc_110EB779:                       var_150 = var_2A8
  loc_110EB7F6:                       call var_80A4 = var_9C(var_9C, frmZGXSToPzTGZS.VFG, 00000082h, 00000002h, 3, var_174, 0, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110EB815:                       var_64 = var_80A4.DispID_0000
  loc_110EB839:                       var_150 = var_2A8
  loc_110EB8B6:                       call var_80AC = var_9C(var_9C, frmZGXSToPzTGZS.VFG, 00000082h, 00000002h, 3, var_174, 1, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_110EB91F:                       If (var_80AC.DispID_0000 = global_1100D76C) Then
  loc_110EB936:                         call var_80B8 = var_9C(var_A8, frmZGXSToPzTGZS.Label3)
  loc_110EB938:                         var_264 = var_80B8
  loc_110EBA48:                         var_80 = "正在处理：第[" & frmZGXSToPzTGZS.VFG.DispID_0082(var_2A8, 2) & " - "
  loc_110EBB89:                         var_D8 = frmZGXSToPzTGZS.VFG.DispID_0082(var_2A8, 0)
  loc_110EBBD0:                         var_98 = var_80 & frmZGXSToPzTGZS.VFG.DispID_0082(var_2A8, 3) & " - " & var_D8 & "]号凭证"
  loc_110EBBE0:                         var_98 = var_80B8.UnkVCall_00000054h
  loc_110EBC9B:                         frmZGXSToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_110EBCCF:                         var_3C = var_24
  loc_110EBCE3:                         Set var_9C = frmZGXSToPzTGZS.Chk
  loc_110EBCE5:                         var_264 = var_9C
  loc_110EBCF7:                         Set var_A0 = var_9C(0)
  loc_110EBD1B:                         var_26C = var_A0
  loc_110EBD85:                         If (var_A0.Value = 1) Then
  loc_110EBDB8:                           var_24C = CInt("cIYear".00000000h)
  loc_110EBDCD:                           var_24C, var_70)
  loc_110EBDDA:                           var_54 = var_24C, var_70)
  loc_110EBDEB:                         Else
  loc_110EBE01:                           var_80E8 = .Proc_17_12_110F4EB0(var_70)
  loc_110EBE13:                           var_54 = var_258
  loc_110EBE16:                         End If
  loc_110EBE1B:                         If var_54 > 0 Then
  loc_110EBE23:                           On Error GoTo loc_110F178D
  loc_110EBE5C:                           "wksAlias".00000000h.00000000h(var_58)
  loc_110EBE7B:                           var_1A0 = var_70
  loc_110EBF44:                           var_D8)
  loc_110EBFF0:                           var_80FC = (var_58 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 3))
  loc_110EBFFD:                           var_1F0 = var_80FC + 1
  loc_110EC0B9:                           var_8104 = (var_64 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 0))
  loc_110EC0C6:                           var_240 = var_8104 + 1
  loc_110EC15C:                           var_8110 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_80FC + 1 And var_8104 + 1
  loc_110EC1E8:                           If CBool(var_8110) Then
  loc_110EC289:                             var_C0 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 6)
  loc_110EC2C6:                             var_1A0 = var_38
  loc_110EC334:                             "kmCodeToProperties".00000002h
  loc_110EC354:                             Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110EC38D:                             var_74.AddNew
  loc_110EC398:                             var_150 = "ibook"
  loc_110EC409:                             var_74.DispID_0000(0)
  loc_110EC40B:                             var_1A0 = "iPeriod"
  loc_110EC4BA:                             var_C0 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 2)
  loc_110EC4D8:                             var_D8)
  loc_110EC571:                             var_74.DispID_0000("DateToPeriod".00000001h)
  loc_110EC5A6:                             var_190 = "csign"
  loc_110EC6B3:                             var_74.DispID_0000(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 3))
  loc_110EC6DA:                             var_190 = "isignseq"
  loc_110EC7FA:                             var_74.DispID_0000(Proc_0_4_11026BD0(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 3), var_64, var_258))
  loc_110EC825:                             var_150 = "ino_id"
  loc_110EC897:                             var_74.DispID_0000(var_54)
  loc_110EC899:                             var_190 = "dbill_date"
  loc_110EC948:                             var_C0 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 2)
  loc_110EC966:                             var_D8)
  loc_110EC9C3:                             var_74.DispID_0000(var_D8)
  loc_110EC9F1:                             var_190 = "idoc"
  loc_110ECA09:                             var_150 = var_24
  loc_110ECB12:                             var_74.DispID_0000(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_150, 4)))
  loc_110ECB3D:                             var_160 = "ctext1"
  loc_110ECBA4:                             var_74.DispID_0000(var_150)
  loc_110ECBAB:                             var_160 = "ctext2"
  loc_110ECC12:                             var_74.DispID_0000(var_150)
  loc_110ECC19:                             var_150 = "cbill"
  loc_110ECC87:                             var_74.DispID_0000("cUserName".00000000h(, var_14C, "cbill", var_154))
  loc_110ECC9D:                             var_160 = "cbook"
  loc_110ECD04:                             var_74.DispID_0000(var_150)
  loc_110ECD0B:                             var_160 = "ccheck"
  loc_110ECD72:                             var_74.DispID_0000(var_150)
  loc_110ECD79:                             var_160 = "ccashier"
  loc_110ECDE0:                             var_74.DispID_0000(var_150)
  loc_110ECDE7:                             var_160 = "iflag"
  loc_110ECE4E:                             var_74.DispID_0000(var_150)
  loc_110ECE55:                             var_160 = "coutaccset"
  loc_110ECEBC:                             var_74.DispID_0000(var_150)
  loc_110ECEC3:                             var_160 = "ioutyear"
  loc_110ECF2A:                             var_74.DispID_0000(var_150)
  loc_110ECF31:                             var_160 = "coutsysver"
  loc_110ECF98:                             var_74.DispID_0000(var_150)
  loc_110ECF9F:                             var_160 = "coutsysname"
  loc_110ED006:                             var_74.DispID_0000(var_150)
  loc_110ED00D:                             var_170 = "ioutperiod"
  loc_110ED0AA:                             var_74.DispID_0000(var_74.DispID_0000("iPeriod"))
  loc_110ED0BB:                             var_170 = "doutbilldate"
  loc_110ED17E:                             var_74.DispID_0000(CStr(var_74.DispID_0000("dbill_date")))
  loc_110ED19B:                             var_150 = "iYear"
  loc_110ED209:                             var_74.DispID_0000("cIYear".00000000h(var_58, var_14C, "iYear", var_154))
  loc_110ED307:                             var_74.DispID_0000("cIYear".00000000h(, var_16C, "iYPeriod", var_174) & Format(var_70, "00"))
  loc_110ED335:                             var_160 = "coutsign"
  loc_110ED39C:                             var_74.DispID_0000(var_70)
  loc_110ED39E:                             var_190 = "coutno_id"
  loc_110ED4AB:                             var_74.DispID_0000(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 3))
  loc_110ED4D7:                             var_150 = "bvouchedit"
  loc_110ED546:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110ED54D:                             var_150 = "bvouchaddordele"
  loc_110ED5BE:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110ED5C5:                             var_150 = "bvouchmoneyhold"
  loc_110ED636:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110ED63D:                             var_150 = "bvalueedit"
  loc_110ED6AE:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110ED6B5:                             var_150 = "bcodeedit"
  loc_110ED726:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110ED72D:                             var_150 = "bPCSedit"
  loc_110ED79E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110ED7A5:                             var_150 = "bDeptedit"
  loc_110ED816:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110ED81D:                             var_150 = "bItemedit"
  loc_110ED88E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110ED895:                             var_150 = "inid"
  loc_110ED907:                             var_74.DispID_0000(1)
  loc_110ED909:                             var_190 = "cdigest"
  loc_110EDA1A:                             var_74.DispID_0000(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 5))
  loc_110EDA41:                             var_190 = "cCode"
  loc_110EDB50:                             var_74.DispID_0000(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 6))
  loc_110EDBF8:                             var_7C = var_38.UnkVCall_0000006Ch
  loc_110EDC43:                             var_8150 = (var_38.UnkVCall_0000006Ch = global_1100AE28)
  loc_110EDC50:                             var_160 = var_8150 + 1
  loc_110EDCDB:                             var_74.DispID_0000(IIf(var_8150 + 1, vbNull, 0))
  loc_110EDDC0:                             var_1B0 = "md"
  loc_110EDE09:                             var_BC = var_25C
  loc_110EDE90:                             var_74.DispID_0000(Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 7)), "#.00"))
  loc_110EDF81:                             var_1B0 = "mc"
  loc_110EDFCA:                             var_BC = var_25C
  loc_110EE051:                             var_74.DispID_0000(Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 8)), "#.00"))
  loc_110EE119:                             If (var_74.DispID_0000("md") <> 0) Then
  loc_110EE18E:                               If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110EE199:                                 var_150 = "md_f"
  loc_110EE20A:                                 var_74.DispID_0000(0)
  loc_110EE214:                               Else
  loc_110EE2C7:                                 var_1B0 = "md_f"
  loc_110EE310:                                 var_BC = var_25C
  loc_110EE397:                                 var_74.DispID_0000(Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_110EE3D8:                               End If
  loc_110EE44A:                               If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_110EE455:                                 var_150 = "nd_s"
  loc_110EE4C6:                                 var_74.DispID_0000(0)
  loc_110EE4D0:                               Else
  loc_110EE4DF:                               Else
  loc_110EE54E:                                 If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110EE559:                                   var_150 = "mc_f"
  loc_110EE5CA:                                   var_74.DispID_0000(0)
  loc_110EE5D4:                                 Else
  loc_110EE687:                                   var_1B0 = "mc_f"
  loc_110EE6D0:                                   var_BC = var_25C
  loc_110EE757:                                   var_74.DispID_0000(Format(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_110EE798:                                 End If
  loc_110EE80A:                                 If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_110EE811:                                   GoTo loc_110EE455
  loc_110EE816:                                 End If
  loc_110EE820:                               End If
  loc_110EE93A:                               var_74.DispID_0000(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 9)))
  loc_110EE960:                             End If
  loc_110EE9D2:                             If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_110EE9DD:                               var_150 = "nfrat"
  loc_110EEA4E:                               var_74.DispID_0000(0)
  loc_110EEA58:                             Else
  loc_110EEB7C:                               var_74.DispID_0000(Val(frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 11)))
  loc_110EEBA2:                             End If
  loc_110EEBF7:                             If var_38.UnkVCall_0000010Ch Then
  loc_110EEC8E:                               var_1F0 = "csettle"
  loc_110EED75:                               var_81A4 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 13) = global_1100AE28)
  loc_110EED82:                               var_1E0 = var_81A4 + 1
  loc_110EEE0D:                               var_74.DispID_0000(IIf(var_81A4 + 1, vbNull, frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 13)))
  loc_110EEE66:                             End If
  loc_110EEE8F:                             var_24C = var_38.UnkVCall_0000010Ch
  loc_110EEEDC:                             var_250 = var_38.UnkVCall_00000094h
  loc_110EEF7B:                             If (var_38.UnkVCall_0000009Ch = 0) = 0 Then
  loc_110EF012:                               var_1F0 = "cn_id"
  loc_110EF0C1:                               var_E0 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 14)
  loc_110EF0F9:                               var_81BC = (frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 14) = global_1100AE28)
  loc_110EF106:                               var_1E0 = var_81BC + 1
  loc_110EF191:                               var_74.DispID_0000(IIf(var_81BC + 1, vbNull, var_E0))
  loc_110EF278:                               var_1F0 = "dt_date"
  loc_110EF327:                               var_D0 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 15)
  loc_110EF345:                               var_E0)
  loc_110EF372:                               var_81C8 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 15) = global_1100AE28)
  loc_110EF37F:                               var_1E0 = var_81C8 + 1
  loc_110EF40A:                               var_74.DispID_0000(IIf(var_81C8 + 1, vbNull, var_E0))
  loc_110EF4F8:                               var_1F0 = "cname"
  loc_110EF5DF:                               var_81D4 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_24, &H14) = global_1100AE28)
  loc_110EF5EC:                               var_1E0 = var_81D4 + 1
  loc_110EF677:                               var_74.DispID_0000(IIf(var_81D4 + 1, vbNull, frmZGXSToPzTGZS.VFG.DispID_0082(var_24, &H14)))
  loc_110EF6D0:                             End If
  loc_110EF746:                             var_250 = var_38.UnkVCall_0000008Ch
  loc_110EF784:                             If (var_38.UnkVCall_000000A4h = 0) = 0 Then
  loc_110EF78E:                               var_150 = var_24
  loc_110EF81B:                               var_1F0 = "cdept_id"
  loc_110EF902:                               var_81E8 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_150, 16) = global_1100AE28)
  loc_110EF90F:                               var_1E0 = var_81E8 + 1
  loc_110EF99A:                               var_74.DispID_0000(IIf(var_81E8 + 1, vbNull, frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 16)))
  loc_110EF9F5:                             Else
  loc_110EF9FA:                               var_160 = "cdept_id"
  loc_110EFA61:                               var_74.DispID_0000(var_150)
  loc_110EFA66:                             End If
  loc_110EFABB:                             If var_38.UnkVCall_0000008Ch Then
  loc_110EFAC5:                               var_150 = var_24
  loc_110EFB52:                               var_1F0 = "cperson_id"
  loc_110EFC39:                               var_81F8 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_150, &H11) = global_1100AE28)
  loc_110EFC46:                               var_1E0 = var_81F8 + 1
  loc_110EFCD1:                               var_74.DispID_0000(IIf(var_81F8 + 1, vbNull, frmZGXSToPzTGZS.VFG.DispID_0082(var_24, &H11)))
  loc_110EFD2C:                             Else
  loc_110EFD31:                               var_160 = "cperson_id"
  loc_110EFD98:                               var_74.DispID_0000(var_150)
  loc_110EFD9D:                             End If
  loc_110EFDF2:                             If var_38.UnkVCall_00000094h Then
  loc_110EFDFC:                               var_150 = var_24
  loc_110EFE89:                               var_1F0 = "ccus_id"
  loc_110EFF70:                               var_8208 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_150, &H12) = global_1100AE28)
  loc_110EFF7D:                               var_1E0 = var_8208 + 1
  loc_110F0008:                               var_74.DispID_0000(IIf(var_8208 + 1, vbNull, frmZGXSToPzTGZS.VFG.DispID_0082(var_24, &H12)))
  loc_110F0063:                             Else
  loc_110F0068:                               var_160 = "ccus_id"
  loc_110F00CF:                               var_74.DispID_0000(var_150)
  loc_110F00D4:                             End If
  loc_110F0129:                             If var_38.UnkVCall_0000009Ch Then
  loc_110F0133:                               var_150 = var_24
  loc_110F01C0:                               var_1F0 = "csup_id"
  loc_110F02A7:                               var_8218 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_150, &H13) = global_1100AE28)
  loc_110F02B4:                               var_1E0 = var_8218 + 1
  loc_110F033F:                               var_74.DispID_0000(IIf(var_8218 + 1, vbNull, frmZGXSToPzTGZS.VFG.DispID_0082(var_24, &H13)))
  loc_110F039A:                             Else
  loc_110F039F:                               var_160 = "csup_id"
  loc_110F0406:                               var_74.DispID_0000(var_150)
  loc_110F040B:                             End If
  loc_110F0484:                             If (var_38.UnkVCall_000000ACh = global_1100AE28) Then
  loc_110F048E:                               var_150 = var_24
  loc_110F051B:                               var_1F0 = "citem_id"
  loc_110F0602:                               var_822C = (frmZGXSToPzTGZS.VFG.DispID_0082(var_150, &H15) = global_1100AE28)
  loc_110F060F:                               var_1E0 = var_822C + 1
  loc_110F069A:                               var_74.DispID_0000(IIf(var_822C + 1, vbNull, frmZGXSToPzTGZS.VFG.DispID_0082(var_24, &H15)))
  loc_110F0777:                               var_7C = var_38.UnkVCall_000000ACh
  loc_110F07C8:                               var_8238 = (var_38.UnkVCall_000000ACh = global_1100AE28)
  loc_110F07D5:                               var_160 = var_8238 + 1
  loc_110F0860:                               var_74.DispID_0000(IIf(var_8238 + 1, vbNull, 0))
  loc_110F089A:                             Else
  loc_110F089F:                               var_160 = "citem_id"
  loc_110F0906:                               var_74.DispID_0000(var_150)
  loc_110F090D:                               var_160 = "citem_class"
  loc_110F0974:                               var_74.DispID_0000(var_150)
  loc_110F0979:                             End If
  loc_110F097E:                             var_160 = "ccode_equal"
  loc_110F09E5:                             var_74.DispID_0000(var_150)
  loc_110F09EC:                             var_160 = "iflagbank"
  loc_110F0A53:                             var_74.DispID_0000(var_150)
  loc_110F0A5A:                             var_160 = "iflagperson"
  loc_110F0AC1:                             var_74.DispID_0000(var_150)
  loc_110F0ACE:                             var_74.Update
  loc_110F0AE5:                             var_24 = var_24(1)
  loc_110F0AF6:                             var_68 = var_68(1)
  loc_110F0B2B:                             var_823C = CLng(frmZGXSToPzTGZS.VFG.DispID_0007)
  loc_110F0B47:                             var_264 = (var_24(1) > 0)
  loc_110F0B6E:                             If var_264 = 0 Then GoTo loc_110EBE78
  loc_110F0B74:                           End If
  loc_110F0BA7:                           "wksAlias".00000000h.00000000h
  loc_110F0BD4:                           Set var_9C = frmZGXSToPzTGZS.Chk
  loc_110F0BD6:                           var_264 = var_9C
  loc_110F0BE8:                           Set var_A0 = var_9C(0)
  loc_110F0C0C:                           var_26C = var_A0
  loc_110F0C76:                           If (var_A0.Value = 1) Then
  loc_110F0C84:                             var_70, var_58)
  loc_110F0C89:                           End If
  loc_110F0C8B:                           On Error GoTo 0
  loc_110F0CC2:                           var_250 = CInt("cIYear".00000000h)
  loc_110F0CEC:                           var_24C, var_250, var_70, var_58)
  loc_110F0CF6:                           var_5C = var_24C, var_250, var_70, var_58)
  loc_110F0D39:                           var_250 = CInt("cIYear".00000000h)
  loc_110F0D6D:                           var_48 = r_250, var_70, var_58) var_250, var_70, var_58)
  loc_110F0D7F:                           var_150 = "select * from GL_accvouch where ibook=0 and iYear="
  loc_110F0DA7:                           var_170 = var_70
  loc_110F0DCB:                           var_824C = Proc_0_4_11026BD0(var_58, var_54, var_54)
  loc_110F0DD0:                           var_190 = var_824C
  loc_110F0DF8:                           var_1B0 = var_54
  loc_110F0E51:                           var_D8 = 1 & "cIYear".00000000h(, 1, 1) & " and iperiod="
  loc_110F0EBA:                           var_128 = var_D8 & var_70 & " and isignseq=" & var_824C & " and ino_id=" & var_54
  loc_110F0F23:                           Set var_74 = "DataMdb".00000000h.00000001h
  loc_110F0FC2:                           If CBool(Not(var_74.EOF)) Then
  loc_110F101A:                             If CBool(Not(var_74.EOF)) Then
  loc_110F1023:                               var_170 = var_70
  loc_110F1038:                               var_150 = "iPeriod"
  loc_110F105C:                               var_180 = "csign"
  loc_110F1070:                               var_1D0 = var_54
  loc_110F1081:                               var_1B0 = "ino_id"
  loc_110F11D8:                               If CBool((var_70 = var_14C) And (var_58 = var_D8) And (var_54 = var_1AC)) Then
  loc_110F11E3:                                 var_150 = "mc"
  loc_110F1265:                                 var_180 = "ccode_equal"
  loc_110F1279:                                 If (var_14C <> 0) Then
  loc_110F12A5:                                   var_8278 = (var_5C = global_1100AE28)
  loc_110F12B2:                                   var_160 = var_8278 + 1
  loc_110F12DF:                                   var_C8 = IIf(var_8278 + 1, vbNull, var_5C)
  loc_110F1359:                                 Else
  loc_110F137F:                                   var_827C = (var_48 = global_1100AE28)
  loc_110F138C:                                   var_160 = var_827C + 1
  loc_110F13B9:                                   var_C8 = IIf(var_827C + 1, vbNull, var_48)
  loc_110F142E:                                 End If
  loc_110F1444:                                 var_74.Update
  loc_110F148E:                                 var_180 = var_38
  loc_110F14D5:                                 var_B8 = var_74.DispID_0000("cCode")
  loc_110F1532:                                 "kmCodeToProperties".00000002h
  loc_110F1552:                                 Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110F1571:                                 var_150 = "citem_class"
  loc_110F15D8:                                 If IsNull(var_74.DispID_0000(var_150)) Then
  loc_110F15ED:                                 Else
  loc_110F162E:                                   var_180 = var_28
  loc_110F1675:                                   var_B8 = var_74.DispID_0000(var_150)
  loc_110F16D2:                                   "XmClassIDToProperties".00000002h
  loc_110F1732:                                   var_78 = {3302AA47-EB96-11D2-AF06000021009B21}().UnkVCall_0000002Ch
  loc_110F1763:                                 End If
  loc_110F1771:                                 var_68 = var_68(1)
  loc_110F177F:                                 var_74.MoveNext
  loc_110F1788:                                 GoTo loc_110F0FCF
  loc_110F17C0:                                 "wksAlias".00000000h.00000000h
  loc_110F17D8:                                 var_30 = var_3C
  loc_110F17ED:                                 var_1A0 = var_70
  loc_110F18B6:                                 var_D8)
  loc_110F1962:                                 var_829C = (var_58 = frmZGXSToPzTGZS.VFG.DispID_0082(var_30, 3))
  loc_110F196F:                                 var_1F0 = var_829C + 1
  loc_110F1A2B:                                 var_82A4 = (var_64 = frmZGXSToPzTGZS.VFG.DispID_0082(var_30, 0))
  loc_110F1A38:                                 var_240 = var_82A4 + 1
  loc_110F1ACE:                                 var_82B0 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_30, 2) = "DateToPeriod".00000001h) And var_829C + 1 And var_82A4 + 1
  loc_110F1B5A:                                 If CBool(var_82B0) Then
  loc_110F1B64:                                   var_150 = var_30
  loc_110F1C20:                                   frmZGXSToPzTGZS.VFG.DispID_0082(1, "-")
  loc_110F1DA0:                                   frmZGXSToPzTGZS.VFG.DispID_009E(var_30, 1, var_30, 1, &HFF)
  loc_110F1DB5:                                   var_150 = var_30
  loc_110F1E71:                                   frmZGXSToPzTGZS.VFG.DispID_0082(&H16, "数据提交错或该数据已经被导入----未引入")
  loc_110F1E90:                                   var_30 = var_30(1)
  loc_110F1EBC:                                   var_82B8 = CLng(frmZGXSToPzTGZS.VFG.DispID_0007)
  loc_110F1ED8:                                   var_264 = (var_30 > 0)
  loc_110F1EFF:                                   If var_264 = 0 Then GoTo loc_110F17EA
  loc_110F1F05:                                 End If
  loc_110F1F08:                                 var_24 = var_30
  loc_110F1F1C:                                 Set var_9C = frmZGXSToPzTGZS.Chk
  loc_110F1F1E:                                 var_264 = var_9C
  loc_110F1F30:                                 Set var_A0 = var_9C(0)
  loc_110F1F54:                                 var_26C = var_A0
  loc_110F1FBE:                                 If (var_A0.Value = 1) Then
  loc_110F20BA:                                   "unLockVouch".00000004h(var_180, var_BC, var_C4, 0, var_74, var_70, var_58, var_16C, var_54, &H4002, var_184)
  loc_110F20C3:                                 End If
  loc_110F20C8:                                 var_150 = "VouchNum"
  loc_110F213D:                                 Set var_34 = "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C4, 0, var_14C, "VouchNum", var_154)
  loc_110F215E:                                 var_150 = "delete  from vouchnum"
  loc_110F21BC:                                 "DataMdb".00000000h.00000001h(1, 1, var_180, var_BC, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_110F2219:                                 frmZGXSToPzTGZS.Pic1.DispID_80010007 = var_150
  loc_110F222D:                                 var_82C4 = Resume(0)
  loc_110F2233:                               End If
  loc_110F2233:                             End If
  loc_110F2233:                           End If
  loc_110F2251:                           var_24 = var_27C+(var_24 - 1)
  loc_110F2254:                           GoTo loc_110EB55C
  loc_110F2259:                         End If
  loc_110F225C:                         var_1A0 = var_70
  loc_110F2325:                         var_D8)
  loc_110F23D1:                         var_82D0 = (var_58 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 3))
  loc_110F23DE:                         var_1F0 = var_82D0 + 1
  loc_110F249A:                         var_82D8 = (var_64 = frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 0))
  loc_110F24A7:                         var_240 = var_82D8 + 1
  loc_110F253D:                         var_82E4 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_82D0 + 1 And var_82D8 + 1
  loc_110F254A:                         var_264 = CBool(var_82E4)
  loc_110F25C9:                         If var_264 = 0 Then GoTo loc_110F2233
  loc_110F25E0:                         Set var_9C = frmZGXSToPzTGZS.Chk
  loc_110F25E2:                         var_264 = var_9C
  loc_110F25F4:                         Set var_A0 = var_9C(0)
  loc_110F2618:                         var_26C = var_A0
  loc_110F265B:                         var_274 = (var_A0.Value = 1)
  loc_110F2686:                         var_150 = var_24
  loc_110F26A7:                         var_190 = "网络共享冲突----未引入"
  loc_110F26B1:                         If var_274 = 0 Then
  loc_110F26B3:                           var_190 = "指定的凭证号无效或重号----未引入"
  loc_110F26BD:                         End If
  loc_110F274E:                         frmZGXSToPzTGZS.VFG.DispID_0082(var_170, var_190)
  loc_110F276D:                         var_24 = var_24(1)
  loc_110F2773:                         var_2A8 = var_24(1)
  loc_110F27A2:                         var_82EC = CLng(frmZGXSToPzTGZS.VFG.DispID_0007)
  loc_110F27BE:                         var_264 = (var_2A8 > 0)
  loc_110F27E5:                         If var_264 = 0 Then GoTo loc_110F2259
  loc_110F27EB:                         GoTo loc_110F2233
  loc_110F27F0:                       End If
  loc_110F27F3:                       var_1A0 = var_70
  loc_110F28BE:                       var_D8)
  loc_110F296C:                       var_82F8 = (var_58 = frmZGXSToPzTGZS.VFG.DispID_0082(var_2A8, 3))
  loc_110F2979:                       var_1F0 = var_82F8 + 1
  loc_110F2A37:                       var_8300 = (var_64 = frmZGXSToPzTGZS.VFG.DispID_0082(var_2A8, 0))
  loc_110F2A44:                       var_240 = var_8300 + 1
  loc_110F2ADA:                       var_830C = (frmZGXSToPzTGZS.VFG.DispID_0082(var_2A8, 2) = "DateToPeriod".00000001h) And var_82F8 + 1 And var_8300 + 1
  loc_110F2AE7:                       var_264 = CBool(var_830C)
  loc_110F2B66:                       If var_264 = 0 Then GoTo loc_110F2233
  loc_110F2C5D:                       If (frmZGXSToPzTGZS.VFG.DispID_0082(var_2A8, &H16) = global_1100AE28) + 1 Then
  loc_110F2C63:                         var_150 = var_2A8
  loc_110F2D1C:                         Set var_9C = frmZGXSToPzTGZS.VFG
  loc_110F2D1F:                         var_9C.DispID_0082(&H16, "凭证借贷不平衡或某分录有错误----未引入")
  loc_110F2D30:                         GoTo loc_110F27F0
  loc_110F2D35:                       End If
  loc_110F2DFF:                       var_C0 = frmZGXSToPzTGZS.VFG.DispID_0082(frmZGXSToPzTGZS.VFG, &H16) & "----未引入"
  loc_110F2E9C:                       frmZGXSToPzTGZS.VFG.DispID_0082(&H16, var_C0)
  loc_110F2ED9:                       GoTo loc_110F27F0
  loc_110F2EDE:                     End If
  loc_110F2F26:                     frmZGXSToPzTGZS.Pic1.DispID_80010007 = var_150
  loc_110F2F3D:                     If var_2C Then
  loc_110F2FCF:                       MsgBox("数据引入已完成，数据已生成用友凭证。", 64, "提示信息", 10, 10)
  loc_110F3041:                       frmZGXSToPzTGZS.VFG.DispID_0007 = 1
  loc_110F309B:                       frmZGXSToPzTGZS.VFG.DispID_0007 = 1
  loc_110F3136:                       frmZGXSToPzTGZS.sBar.DispID_6803001E(1100AE28h)
  loc_110F31CD:                       frmZGXSToPzTGZS.sBar.DispID_6803001E(1100AE28h)
  loc_110F3264:                       Set var_9C = frmZGXSToPzTGZS.sBar
  loc_110F3267:                       var_9C.DispID_6803001E(1100AE28h)
  loc_110F327D:                     Else
  loc_110F3304:                       MsgBox("数据没有被引入，原因请查看最后一列中的说明。", 64, "提示信息", 10, 10)
  loc_110F3331:                     End If
  loc_110F3336:                     var_150 = "VouchNum"
  loc_110F33AF:                     Set var_34 = "DataMdb".00000000h.00000001h(var_180, var_BC, var_C0, var_C4, var_C8, var_14C, "VouchNum", var_154)
  loc_110F33D0:                     var_150 = "delete  from vouchnum"
  loc_110F3420:                     "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C0, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_110F3475:                     "AS13")
  loc_110F34AE:                     var_B8 = Global.Screen
  loc_110F34D0:                     var_8330 = ecx
  loc_110F34DF:                     If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110F34E9:                     End If
  loc_110F34E9:                   End If
  loc_110F34E9:                 End If
  loc_110F34E9:               End If
  loc_110F34E9:             End If
  loc_110F34EA:             var_8330 = CheckObj(var_9C, global_1100C47C, 124)
  loc_110F34F0:           End If
  loc_110F34F0:         End If
  loc_110F34F0:       End If
  loc_110F34F0:     End If
  loc_110F34F0:   End If
  loc_110F34F0: End If
  loc_110F34FC: Exit Sub
  loc_110F3508: GoTo loc_110F35C1
  loc_110F35C0: Exit Sub
  loc_110F35C1: ' Referenced from: 110EA9DC
  loc_110F35C1: ' Referenced from: 110F3508
End Sub

Private Sub Proc_17_11_110F40A0
  Dim var_58 As Variant
  Dim var_5C As Variant
  Dim var_64 As frmZGXSToPzTGZS.Label3
  Dim var_1D0 As Label
  loc_110F418D: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110F4196: var_1F0 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110F41B3: Set var_58 = frmZGXSToPzTGZS.Chk
  loc_110F41BD: var_1D0 = var_58
  loc_110F41C3: Set var_5C = var_58(0)
  loc_110F41EE: var_1D8 = var_5C
  loc_110F4231: var_1E0 = (var_5C.Value = 1)
  loc_110F4247: If var_1E0 = 0 Then
  loc_110F42AC:   If var_14 <= CLng(frmZGXSToPzTGZS.VFG.DispID_0007)(-1) Then
  loc_110F4321:     var_7C = frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 2)
  loc_110F433C:     var_94)
  loc_110F4397:     var_30 = CByte("DateToPeriod".00000001h)
  loc_110F44F1:     Set var_64 = frmZGXSToPzTGZS.Label3
  loc_110F451B:     var_1D0 = var_64
  loc_110F46D1:     var_94 = frmZGXSToPzTGZS.VFG.DispID_0082(var_14, frmZGXSToPzTGZS.VFG)
  loc_110F46ED:     var_8034 = "正在处理：第[" & frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 2) & " - " & frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 3) & " - " & var_94
  loc_110F4723:     var_64.Caption = var_8034 & "]号凭证是否重号"
  loc_110F47B2:     var_803C = frmZGXSToPzTGZS.Proc_17_12_110F4EB0(var_30)
  loc_110F47C7:     If var_1CC <= 0 Then
  loc_110F47D9:       var_13C = var_30
  loc_110F4870:       var_94)
  loc_110F4909:       var_804C = (frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 3) = frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 3))
  loc_110F4936:       var_17C = var_804C + 1
  loc_110F49AD:       var_8054 = (frmZGXSToPzTGZS.VFG.DispID_0082(var_14, frmZGXSToPzTGZS.VFG) = frmZGXSToPzTGZS.VFG.DispID_0082(var_14, ""))
  loc_110F49D4:       var_1BC = var_8054 + 1
  loc_110F4ACF:       If CBool((frmZGXSToPzTGZS.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_804C + 1 And var_8054 + 1) Then
  loc_110F4B62:         frmZGXSToPzTGZS.VFG.DispID_0082(var_10C, 285267820)
  loc_110F4C96:         frmZGXSToPzTGZS.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110F4D2A:         frmZGXSToPzTGZS.VFG.DispID_0082(var_10C, "指定的凭证号无效或重号")
  loc_110F4D75:         var_8068 = CLng(frmZGXSToPzTGZS.VFG.DispID_0007)
  loc_110F4D93:         var_1D0 = (var_14(1) > 0)
  loc_110F4DB0:         If var_1D0 = 0 Then GoTo loc_110F47D3
  loc_110F4DB6:       End If
  loc_110F4DC4:     Else
  loc_110F4DCD:     End If
  loc_110F4DDA:     var_14 = 1+var_14
  loc_110F4DDD:     GoTo loc_110F42A6
  loc_110F4DE2:   End If
  loc_110F4DE2: End If
  loc_110F4DE7: GoTo loc_110F4E78
  loc_110F4E77: Exit Sub
  loc_110F4E78: ' Referenced from: 110F4DE7
End Sub

Private  Proc_17_12_110F4EB0(arg_C, arg_10, arg_14) '110F4EB0
  loc_110F4F49: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110F4F52: var_168 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110F4F7B: If IsNumeric(arg_14) Then
  loc_110F4F8A:   var_8008 = CLng(Val(arg_14))
  loc_110F4F94:   If var_8008 > 0 Then
  loc_110F4FA0:     If var_8008 <= 9999 Then
  loc_110F501C:       var_8028 = "select * from GL_accvouch where iperiod >=" & CStr(arg_C) & " and isignseq>=" & CStr(0) & " and ino_id>=" & CStr(var_8008)
  loc_110F5031:       var_44 = var_8028
  loc_110F5083:       Set var_1C = "DataMdb".00000000h.00000001h(fs:[00000000h], , , , , var_40, var_8028, var_48)
  loc_110F50C8:       var_8030 = Proc_0_4_11026BD0(arg_10, , )
  loc_110F50E9:       var_8034 = CBool(var_1C.EOF)
  loc_110F50FD:       If var_8034 = 0 Then
  loc_110F5128:         var_F4 = arg_C
  loc_110F51E6:         var_8040 = (var_1C.DispID_0000("iPeriod") = arg_C) And (var_1C.DispID_0000("isignseq") = (Proc_0_4_11026BD0(arg_10, , ) And 255))
  loc_110F5256:         var_804C = CBool(Not(var_8040 And (var_1C.DispID_0000("ino_id") = var_8008)))
  loc_110F527B:         If var_804C = 0 Then GoTo loc_110F5280
  loc_110F527D:       End If
  loc_110F528B:       var_1C.oClose
  loc_110F5294:     End If
  loc_110F5294:   End If
  loc_110F5294: End If
  loc_110F529A: GoTo loc_110F52FF
  loc_110F52FE: Exit Sub
  loc_110F52FF: ' Referenced from: 110F529A
End Sub
