VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C908002B2F49FB}#1.2#0"; "C:\Windows\SysWow64\Comdlg32.ocx"
Begin VB.Form frmGzDepMap
  Caption = "工资部门对照表"
  BackColor = &H80000005&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmGzDepMap.frx":0000
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
    OleObjectBlob = "frmGzDepMap.frx":014A
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
    OleObjectBlob = "frmGzDepMap.frx":0293
    Begin AIFCmp1.asxStatusBar sBar
      Left = 0
      Top = 5355
      Width = 12045
      Height = 345
      OleObjectBlob = "frmGzDepMap.frx":04BC
    End
    Begin C1SizerLibCtl.C1Elastic C1Elastic2
      Left = 0
      Top = 0
      Width = 12045
      Height = 435
      TabStop = 0   'False
      TabIndex = 1
      OleObjectBlob = "frmGzDepMap.frx":05EC
    End
    Begin VSFlex8LCtl.VSFlexGrid VFG
      Left = 0
      Top = 1260
      Width = 12045
      Height = 4080
      TabIndex = 2
      OleObjectBlob = "frmGzDepMap.frx":0745
      Begin MSComDlg.CommonDialog dlg
        OleObjectBlob = "frmGzDepMap.frx":0BAE
        Left = 0
        Top = 0
      End
    End
    Begin AIFCmp1.asxPanel asxPanel1
      Left = 0
      Top = 450
      Width = 12045
      Height = 795
      OleObjectBlob = "frmGzDepMap.frx":0C12
      Begin AIFCmp1.asxToolButton APB
        Index = 0
        Left = 90
        Top = 60
        Width = 705
        Height = 285
        OleObjectBlob = "frmGzDepMap.frx":0CF2
      End
      Begin VB.ComboBox Cbo
        Style = 2
        Left = 12030
        Top = 60
        Width = 4545
        Height = 300
        Visible = 0   'False
        TabIndex = 6
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 0
        Left = 150
        Top = 420
        Width = 3495
        Height = 270
        TabIndex = 5
        OleObjectBlob = "frmGzDepMap.frx":0DF6
        ToolTipText = "朗新部门名称"
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 1
        Left = 3660
        Top = 420
        Width = 1935
        Height = 270
        TabIndex = 7
        OleObjectBlob = "frmGzDepMap.frx":0F5A
        ToolTipText = "类型"
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 2
        Left = 5760
        Top = 420
        Width = 3495
        Height = 270
        TabIndex = 8
        OleObjectBlob = "frmGzDepMap.frx":10B6
        ToolTipText = "用友部门编码"
      End
      Begin AIFCmp1.asxToolButton APB
        Index = 1
        Left = 828
        Top = 60
        Width = 705
        Height = 285
        OleObjectBlob = "frmGzDepMap.frx":121A
      End
      Begin AIFCmp1.asxToolButton APB
        Index = 2
        Left = 1566
        Top = 60
        Width = 705
        Height = 285
        OleObjectBlob = "frmGzDepMap.frx":131E
      End
      Begin AIFCmp1.asxToolButton APB
        Index = 3
        Left = 2304
        Top = 60
        Width = 705
        Height = 285
        OleObjectBlob = "frmGzDepMap.frx":1422
      End
      Begin AIFCmp1.asxToolButton APB
        Index = 4
        Left = 3042
        Top = 60
        Width = 705
        Height = 285
        OleObjectBlob = "frmGzDepMap.frx":1526
      End
      Begin AIFCmp1.asxToolButton APB
        Index = 5
        Left = 3780
        Top = 60
        Width = 705
        Height = 285
        OleObjectBlob = "frmGzDepMap.frx":162A
      End
      Begin AIFCmp1.asxToolButton APB
        Index = 6
        Left = 4518
        Top = 60
        Width = 705
        Height = 285
        OleObjectBlob = "frmGzDepMap.frx":172E
      End
      Begin AIFCmp1.asxToolButton APB
        Index = 7
        Left = 5256
        Top = 60
        Width = 705
        Height = 285
        OleObjectBlob = "frmGzDepMap.frx":1832
      End
      Begin AIFCmp1.asxToolButton APB
        Index = 8
        Left = 6000
        Top = 60
        Width = 705
        Height = 285
        Visible = 0   'False
        OleObjectBlob = "frmGzDepMap.frx":1936
      End
    End
  End
End

Attribute VB_Name = "frmGzDepMap"


Private  APB_UnknownEvent_9(arg_C) '1104C0D0
  Dim var_24 As Variant
  loc_1104C15C: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1104C165: var_120 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1104C17C: Set var_24 = frmGzDepMap.APB
  loc_1104C187: var_E8 = var_24
  loc_1104C192: arg_C = var_24.UnkVCall_00000040h
  loc_1104C1D8: var_F4 = var_28.DispID_FFFFFDFA
  loc_1104C206: var_8008 = (var_F4 = "刷新")
  loc_1104C20E: If var_8008 = 0 Then
  loc_1104C222:   If frmGzDepMap.FillData >= 0 Then GoTo loc_1104D2E2
  loc_1104C234:   var_E4 = CheckObj(Me, global_1100C54C, 1788)
  loc_1104C23F: End If
  loc_1104C253: If (var_F4 = "取消加载") Then
  loc_1104C265:   var_8010 = (var_F4 = "全选")
  loc_1104C26D:   If var_8010 = 0 Then
  loc_1104C2C2:     If (CLng(frmGzDepMap.VFG.DispID_0007) <= 1) Then GoTo loc_1104D2E2
  loc_1104C2E1:     var_40 = frmGzDepMap.VFG.DispID_0007
  loc_1104C322:     If 1 > CLng(var_40)(-1) Then GoTo loc_1104D2E2
  loc_1104C3C0:     frmGzDepMap.VFG.DispID_0082(frmGzDepMap.VFG, 285274092)
  loc_1104C3E2:     GoTo loc_1104C31B
  loc_1104C3E7:   End If
  loc_1104C3F3:   var_8020 = (var_F4 = "全消")
  loc_1104C3FB:   If var_8020 = 0 Then
  loc_1104C412:     Set var_24 = frmGzDepMap.VFG
  loc_1104C419:     call 1+1(var_40, var_24, 00000007h, var_8020, var_28)
  loc_1104C450:     If (CLng(1+1(var_40, var_24, 00000007h, var_8020, var_28)) <= 1) Then GoTo loc_1104D2E2
  loc_1104C468:     Set var_24 = frmGzDepMap.VFG
  loc_1104C46F:     call 1+1(var_40, var_24, 00000007h, 00000000h)
  loc_1104C4B0:     If 1 > CLng(1+1(var_40, var_24, 00000007h, 00000000h))(-1) Then GoTo loc_1104D2E2
  loc_1104C54E:     frmGzDepMap.VFG.DispID_0082(frmGzDepMap.VFG, 285257256)
  loc_1104C570:     GoTo loc_1104C4A9
  loc_1104C575:   End If
  loc_1104C581:   var_8030 = (var_F4 = "删除")
  loc_1104C589:   If var_8030 = 0 Then
  loc_1104C595:     var_1C = var_8030
  loc_1104C5A3:     Set var_24 = frmGzDepMap.VFG
  loc_1104C5AA:     call 1+1(var_40, var_24, 00000007h, var_8030)
  loc_1104C5F0:     If var_18 <= CLng(1+1(var_40, var_24, 00000007h, var_8030))(-1) Then
  loc_1104C61D:       var_88 = var_18
  loc_1104C665:       Set var_24 = frmGzDepMap.VFG
  loc_1104C66C:       call 1+1(var_40, var_24, 00000082h, 00000002h, var_B0, var_AC, var_24, var_A4, 00000003h, var_8C, var_18, var_84)
  loc_1104C689:       var_8040 = (1+1(var_40, var_24, 00000082h, 00000002h, var_B0, var_AC, var_24, var_A4, 00000003h, var_8C, var_18, var_84) = global_1100EFEC)
  loc_1104C6B7:       If var_8040 = 0 Then
  loc_1104C6C7:         var_1C = var_1C(1)
  loc_1104C6CA:       End If
  loc_1104C6DF:       var_18 = 1+var_18
  loc_1104C6E2:       GoTo loc_1104C5E6
  loc_1104C6E7:     End If
  loc_1104C6EC:     If var_1C = 0 Then
  loc_1104C72C:       var_50 = "提示信息"
  loc_1104C747:       var_40 = "请选择相应的数据！"
  loc_1104C75B:       MsgBox(var_40, 64, var_50, 10, 10)
  loc_1104C781:     Else
  loc_1104C78B:       var_8044 = .Proc_11_10_1104E0E0(var_E4)
  loc_1104C799:       If var_E4 Then
  loc_1104C7A9:         var_E4 = .FillData
  loc_1104C7B1:         If var_E4 < 0 Then
  loc_1104C7C3:           var_E4 = CheckObj(var_E4 = %S_edx_S, global_1100C54C, 1788)
  loc_1104C7CE:         End If
  loc_1104C7DA:         var_8048 = (var_F4 = "新增")
  loc_1104C7E2:         If var_8048 = 0 Then
  loc_1104C810:           call edi(var_24, frmGzDepMap.TDBText)
  loc_1104C81D:           var_804C = UnkObj.UnkVCall_00000040h
  loc_1104C89E:           call edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28)
  loc_1104C8AB:           var_8050 = UnkObj.UnkVCall_00000040h
  loc_1104C92E:           call edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104C93B:           var_8054 = UnkObj.UnkVCall_00000040h
  loc_1104C99D:           var_8058 = .Proc_11_8_1104B050(esi+0000003Ah)
  loc_1104C9A8:         Else
  loc_1104C9B4:           var_805C = (var_F4 = "修改")
  loc_1104C9BC:           If var_805C = 0 Then
  loc_1104C9D3:             var_8060 = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104C9DA:             var_40.DispID_0000 = var_24
  loc_1104CA11:             If (CLng(var_40) < 2) Then GoTo loc_1104D2E2
  loc_1104CA2F:             var_8068 = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CA36:             var_40.DispID_0000 = var_24
  loc_1104CA58:             var_88 = CLng(var_8068)
  loc_1104CAB1:             var_8070 = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CAB8:             var_50.DispID_0000 = var_28
  loc_1104CADC:             var_8078 = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CAE7:             var_E8 = var_8078
  loc_1104CAED:             var_8078.UnkVCall_00000040h
  loc_1104CB2D:             var_30.DispID_0000 = var_8070
  loc_1104CB74:             var_807C = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CB7B:             var_40.DispID_0000 = var_24
  loc_1104CB9D:             var_88 = CLng(var_807C)
  loc_1104CBF5:             var_8084 = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CBFC:             var_50.DispID_0000 = var_28
  loc_1104CC20:             var_808C = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CC2B:             var_E8 = var_808C
  loc_1104CC31:             var_808C.UnkVCall_00000040h
  loc_1104CC71:             var_30.DispID_0000 = var_8084
  loc_1104CCB8:             var_8090 = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CCBF:             var_40.DispID_0000 = var_24
  loc_1104CCE1:             var_88 = CLng(var_8090)
  loc_1104CD3A:             var_8098 = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CD41:             var_50.DispID_0000 = var_28
  loc_1104CD65:             var_80A0 = edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText, edi(var_24, frmGzDepMap.TDBText), 00000000h, var_28), 00000001h, var_28)
  loc_1104CD72:             var_80A0.UnkVCall_00000040h
  loc_1104CDAC:             var_30.DispID_0000 = var_8098
  loc_1104CDE8:             var_80A4 = .Proc_11_8_1104B050(esi+0000003Ah)
  loc_1104CDF3:           Else
  loc_1104CDFF:             var_80A8 = (var_F4 = "放弃")
  loc_1104CE07:             If var_80A8 = 0 Then
  loc_1104CE33:               call var_80AC = var_80A0(var_24, frmGzDepMap.TDBText, var_80A0, 00000002h, var_30, var_2C, frmGzDepMap.TDBText)
  loc_1104CE40:               var_80AC.UnkVCall_00000040h
  loc_1104CEC1:               call var_80B0 = var_80A0(var_24, frmGzDepMap.TDBText, var_80AC, 00000000h, var_28)
  loc_1104CECE:               var_80B0.UnkVCall_00000040h
  loc_1104CF51:               call var_80B4 = var_80A0(var_24, frmGzDepMap.TDBText, var_80B0, 00000001h, var_28)
  loc_1104CF5E:               var_80B4.UnkVCall_00000040h
  loc_1104CFC0:               var_80B8 = .Proc_11_8_1104B050(esi+0000003Ah)
  loc_1104CFCB:             Else
  loc_1104CFD7:               var_80BC = (var_F4 = "保存")
  loc_1104CFDF:               If var_80BC = 0 Then
  loc_1104CFFB:                 If var_80BC <= 2 Then
  loc_1104D00F:                   call var_80C0 = var_80B4(var_24, frmGzDepMap.TDBText, var_80B4, 00000002h, var_28)
  loc_1104D01C:                   var_E8 = var_80C0
  loc_1104D043:                   var_28 = 0
  loc_1104D04A:                   var_38 = var_28
  loc_1104D07E:                   var_F0 = (Proc_0_11_11029000(var_40, var_28) = global_1100AE28) + 1
  loc_1104D0A4:                   If var_F0 = 0 Then
  loc_1104D0B7:                     var_18 = var_110+var_18
  loc_1104D0BA:                     GoTo loc_1104CFF2
  loc_1104D0BF:                   End If
  loc_1104D0CD:                   call var_80CC = var_80B4(var_24, frmGzDepMap.TDBText)
  loc_1104D0DA:                   var_E8 = var_80CC
  loc_1104D0E0:                   var_18 = var_80CC.UnkVCall_00000040h
  loc_1104D10D:                   var_38.DispID_0000 = global_8001004A
  loc_1104D185:                   MsgBox(var_40 & "不能为空，请输入！", 64, "提示信息", 10, var_80)
  loc_1104D1D3:                   call var_80D8 = var_80B4(var_24, frmGzDepMap.TDBText)
  loc_1104D1E2:                   var_18 = var_80D8.UnkVCall_00000040h
  loc_1104D205:                   var_28.DispID_80011000
  loc_1104D21F:                 Else
  loc_1104D22D:                   var_80DC = .Proc_11_9_1104D370(var_80D8(58))
  loc_1104D23B:                   If var_E4 Then
  loc_1104D24A:                     var_80E0 = .Proc_11_8_1104B050(var_80D8(58))
  loc_1104D262:                     If .FillData < 0 Then
  loc_1104D270:                       var_E4 = CheckObj(var_80D8, global_1100C54C, 1788)
  loc_1104D278:                     End If
  loc_1104D284:                     var_80E4 = (var_F4 = global_1100EBD4)
  loc_1104D28C:                     If var_80E4 = 0 Then
  loc_1104D2B9:                       Set var_24 = var_80D8
  loc_1104D2C1:                       var_80EC = Global.Unload var_28
  loc_1104D2E2:                     End If
  loc_1104D2E2:                   End If
  loc_1104D2E2:                 End If
  loc_1104D2E2:               End If
  loc_1104D2E2:             End If
  loc_1104D2E2:           End If
  loc_1104D2E2:         End If
  loc_1104D2E2:       End If
  loc_1104D2E2:     End If
  loc_1104D2E2:   End If
  loc_1104D2E2: End If
  loc_1104D2E2: ' Referenced from: 1104D276
  loc_1104D2EE: GoTo loc_1104D331
  loc_1104D330: Exit Sub
  loc_1104D331: ' Referenced from: 1104D2EE
End Sub

Private Sub VFG_UnknownEvent_A '1104EA50
  loc_1104EB01: If CBool(frmGzDepMap.VFG.DispID_8001000D) + 1 = 0 Then
  loc_1104EB3F:   var_A4 = (CLng(frmGzDepMap.VFG.DispID_0011) < 1)
  loc_1104EB5C:   If var_A4 = 0 Then
  loc_1104EC3E:     If (frmGzDepMap.VFG.DispID_0082(CLng(frmGzDepMap.VFG.DispID_0011), "") = global_1100EFEC) + 1 Then
  loc_1104EC5F:       var_8018 = CLng(frmGzDepMap.VFG.DispID_0011)
  loc_1104ECB9:     Else
  loc_1104ECD8:       var_801C = CLng(frmGzDepMap.VFG.DispID_0008)
  loc_1104ED30:     End If
  loc_1104ED54:     frmGzDepMap.VFG.DispID_0082(frmGzDepMap.VFG, 285274092)
  loc_1104ED76:   End If
  loc_1104ED76: End If
  loc_1104ED82: GoTo loc_1104EDB1
  loc_1104EDB0: Exit Sub
  loc_1104EDB1: ' Referenced from: 1104ED82
End Sub

Private Sub VFG_UnknownEvent_17 '1104E9E0

End Sub

Private Sub Form_Load() '1104AEA0
  Dim var_1C As frmGzDepMap.TDBText
  Dim var_24 As var_20.DispID_03E8
  loc_1104AF09: If ebx <= 2 Then
  loc_1104AF1A:   Set var_1C = frmGzDepMap.TDBText
  loc_1104AF2A:   var_1C.UnkVCall_00000040h
  loc_1104AF56:   var_34 = var_20.DispID_03E8
  loc_1104AF6B:   Set var_24 = var_20.DispID_03E8
  loc_1104AFB7:   var_4C = var_4C + ebx
  loc_1104AFC4:   GoTo loc_1104AEFE
  loc_1104AFC9: End If
  loc_1104AFC9: var_8004 = frmGzDepMap.Proc_11_7_11049E70(16763579)
  loc_1104AFD6: var_38 = frmGzDepMap.FillData
  loc_1104AFF9: var_8008 = frmGzDepMap.Proc_11_8_1104B050(global_58, global_1100D720, var_1C)
  loc_1104B007: GoTo loc_1104B02A
  loc_1104B029: Exit Sub
  loc_1104B02A: ' Referenced from: 1104B007
End Sub

Private Sub Form_Resize() '1104BDA0
  loc_1104BE2D: var_38 = frmGzDepMap.Pic1.DispID_80010005
  loc_1104BE51: var_48 = frmGzDepMap.Pic1.DispID_80010006
  loc_1104BE64: var_EC = var_48.ScaleWidth
  loc_1104BE9B: If global_110F6000 = 0 Then
  loc_1104BEA5: Else
  loc_1104BEB0: End If
  loc_1104BEB0: var_70 = ((var_EC - CSgn(var_38)) / 2)
  loc_1104BEC5: var_F0 = var_48.ScaleHeight
  loc_1104BF03: If global_110F6000 = 0 Then
  loc_1104BF0D: Else
  loc_1104BF18: End If
  loc_1104C023: frmGzDepMap.Pic1.DispID_80011002(var_70, ((var_F0 - CSgn(var_48)) / 2), CSgn(frmGzDepMap.Pic1.DispID_80010005), CSgn(frmGzDepMap.Pic1.DispID_80010006))
  loc_1104C06C: GoTo loc_1104C0A6
End Sub

Public Sub ExitForm(Cancel, UnloadMode) '11049D90
  Dim var_18 As Global
  loc_11049DCF: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11049DFA: Set var_18 = Me
  loc_11049E02: var_8008 = Global.Unload
  loc_11049E3C: GoTo loc_11049E48
  loc_11049E47: Exit Sub
  loc_11049E48: ' Referenced from: 11049E3C
End Sub

Public Function FillData() '1104A490
  Dim var_1C As ADODB.Recordset
  loc_1104A508: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1104A51E: var_C8 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1104A55F: frmGzDepMap.VFG.DispID_0007 = 1
  loc_1104A58A: frmGzDepMap.Label3.Caption = "正在读取数据表，请稍候。。。"
  loc_1104A5F3: frmGzDepMap.Pic1.DispID_80010007 = True
  loc_1104A616: frmGzDepMap.Pic1.DispID_FFFFFDDA
  loc_1104A649: var_AC = ADODB.Recordset.State
  loc_1104A66E: If var_AC = 1 Then
  loc_1104A68E:   var_800C = ADODB.Recordset.Close
  loc_1104A6AC: End If
  loc_1104A74D: var_B0 = var_1C
  loc_1104A785: var_801C = ADODB.Recordset.Open( & "SELECT * FROM " & "..T_CY_GZ_SL_DepSetting ", var_90,  & "SELECT * FROM " & "..T_CY_GZ_SL_DepSetting ", var_88, 9)
  loc_1104A7CE: var_A8 = ADODB.Recordset.EOF
  loc_1104A7EE: If var_A8 = 0 Then
  loc_1104A812:   frmGzDepMap.Label3.Caption = "正在填充数据，请稍候。。。"
  loc_1104A87B:   frmGzDepMap.Pic1.DispID_80010007 = True
  loc_1104A89E:   frmGzDepMap.Pic1.DispID_FFFFFDDA
  loc_1104A8D1:   var_A8 = ADODB.Recordset.EOF
  loc_1104A8F1:   If var_A8 = 0 Then
  loc_1104A977:     var_B8 = ADODB.Recordset.Fields
  loc_1104A994:     ADODB.Recordset.8 = Forms
  loc_1104AAA7:     var_B8 = ADODB.Recordset.Fields
  loc_1104AAC4:     ADODB.Recordset.8 = Forms
  loc_1104ABD7:     var_B8 = ADODB.Recordset.Fields
  loc_1104ABF4:     ADODB.Recordset.8 = Forms
  loc_1104AC5C:     var_84 = var_24 & Chr(9) & Proc_0_11_11029000(9, var_A0, "GZ_Dep") & Chr(9) & Proc_0_11_11029000(9, var_A0, "GZ_Type") & Chr(9) & Proc_0_11_11029000(9, var_A0, "UF_DepCode")
  loc_1104AC6A:     var_24 = var_84
  loc_1104ACBB:     var_8050 = ADODB.Recordset.MoveNext
  loc_1104AD25:     frmGzDepMap.VFG.DispID_0080(var_24)
  loc_1104AD2E:     GoTo loc_1104A8A7
  loc_1104AD33:   End If
  loc_1104AD58:   var_AC = ADODB.Recordset.State
  loc_1104AD7D:   If var_AC = 1 Then
  loc_1104AD9D:     var_805C = ADODB.Recordset.Close
  loc_1104ADBB:   End If
  loc_1104ADBB: End If
  loc_1104AE04: frmGzDepMap.Pic1.DispID_80010007 = var_8C
  loc_1104AE14: GoTo loc_1104AE52
  loc_1104AE51: Exit Function
  loc_1104AE52: ' Referenced from: 1104AE14
End Function

Private Sub Proc_11_7_11049E70
  Dim var_58 As frmGzDepMap.VFG
  loc_11049EB1: Set var_58 = frmGzDepMap.VFG
  loc_11049F02: var_58.DispID_005D = frmGzDepMap.VFG
  loc_11049F43: var_58.DispID_0067 = frmGzDepMap.VFG
  loc_11049F62: var_58.DispID_0041 = frmGzDepMap.VFG
  loc_11049FC2: var_58.DispID_0047 = frmGzDepMap.VFG
  loc_1104A0D0: var_58.DispID_008A(4)
  loc_1104A113: var_58.DispID_0079(300)
  loc_1104A156: var_58.DispID_0077(4)
  loc_1104A199: var_58.DispID_0078(700)
  loc_1104A1DA: var_58.DispID_0077(1)
  loc_1104A220: var_58.DispID_0078(1700)
  loc_1104A266: var_58.DispID_0077(1)
  loc_1104A2AC: var_58.DispID_0078(700)
  loc_1104A2EE: var_58.DispID_0077(1)
  loc_1104A330: var_58.DispID_0078(1700)
  loc_1104A375: var_58.DispID_0090("选择标志")
  loc_1104A3BD: var_58.DispID_0090("朗新部门名称")
  loc_1104A405: var_58.DispID_0090("类型")
  loc_1104A44A: var_58.DispID_0090("用友部门编码")
End Sub

Private  Proc_11_8_1104B050(arg_C) '1104B050
  Dim var_14 As Variant
  loc_1104B08C: If arg_C Then
  loc_1104B092:   If Not Asm.le_flag Then GoTo loc_1104BD6D
  loc_1104B09B:   If arg_C > 2 Then GoTo loc_1104BD6D
  loc_1104B0C2:   frmGzDepMap.APB.UnkVCall_00000040h
  loc_1104B0FE:   var_18.DispID_6803001B = var_1C
  loc_1104B125:   Set var_14 = frmGzDepMap.APB
  loc_1104B134:   var_3C = var_14
  loc_1104B137:   var_14.UnkVCall_00000040h
  loc_1104B199:   Set var_14 = frmGzDepMap.APB
  loc_1104B1A8:   var_3C = var_14
  loc_1104B1AB:   var_14.UnkVCall_00000040h
  loc_1104B20D:   Set var_14 = frmGzDepMap.APB
  loc_1104B21C:   var_3C = var_14
  loc_1104B21F:   var_14.UnkVCall_00000040h
  loc_1104B281:   Set var_14 = frmGzDepMap.APB
  loc_1104B290:   var_3C = var_14
  loc_1104B293:   var_14.UnkVCall_00000040h
  loc_1104B2F5:   Set var_14 = frmGzDepMap.APB
  loc_1104B304:   var_3C = var_14
  loc_1104B307:   var_14.UnkVCall_00000040h
  loc_1104B369:   Set var_14 = frmGzDepMap.APB
  loc_1104B378:   var_3C = var_14
  loc_1104B37B:   var_14.UnkVCall_00000040h
  loc_1104B3B7:   var_18.DispID_6803001B = True
  loc_1104B3DE:   Set var_14 = frmGzDepMap.APB
  loc_1104B3ED:   var_3C = var_14
  loc_1104B3F0:   var_14.UnkVCall_00000040h
  loc_1104B42C:   var_18.DispID_6803001B = True
  loc_1104B45A:   Set var_14 = frmGzDepMap.APB
  loc_1104B469:   var_3C = var_14
  loc_1104B46C:   var_14.UnkVCall_00000040h
  loc_1104B4A8:   var_18.DispID_6803001B = var_20
  loc_1104B4CF:   Set var_14 = frmGzDepMap.TDBText
  loc_1104B4DE:   var_3C = var_14
  loc_1104B4E1:   var_14.UnkVCall_00000040h
  loc_1104B519:   var_18.DispID_000F = var_20
  loc_1104B540:   Set var_14 = frmGzDepMap.TDBText
  loc_1104B54F:   var_3C = var_14
  loc_1104B552:   var_14.UnkVCall_00000040h
  loc_1104B58A:   var_18.DispID_000F = var_20
  loc_1104B5B1:   Set var_14 = frmGzDepMap.TDBText
  loc_1104B5C0:   var_3C = var_14
  loc_1104B5C3:   var_14.UnkVCall_00000040h
  loc_1104B5FB:   var_18.DispID_000F = var_20
  loc_1104B640:   frmGzDepMap.VFG.DispID_8001000D = var_20
  loc_1104B659: Else
  loc_1104B67E:   frmGzDepMap.APB.UnkVCall_00000040h
  loc_1104B6BA:   var_18.DispID_6803001B = var_20
  loc_1104B6E1:   Set var_14 = frmGzDepMap.APB
  loc_1104B6F0:   var_3C = var_14
  loc_1104B6F3:   var_14.UnkVCall_00000040h
  loc_1104B72F:   var_18.DispID_6803001B = var_20
  loc_1104B756:   Set var_14 = frmGzDepMap.APB
  loc_1104B765:   var_3C = var_14
  loc_1104B768:   var_14.UnkVCall_00000040h
  loc_1104B7A4:   var_18.DispID_6803001B = var_20
  loc_1104B7CB:   Set var_14 = frmGzDepMap.APB
  loc_1104B7DA:   var_3C = var_14
  loc_1104B7DD:   var_14.UnkVCall_00000040h
  loc_1104B819:   var_18.DispID_6803001B = var_20
  loc_1104B840:   Set var_14 = frmGzDepMap.APB
  loc_1104B84F:   var_3C = var_14
  loc_1104B852:   var_14.UnkVCall_00000040h
  loc_1104B88E:   var_18.DispID_6803001B = var_20
  loc_1104B8B5:   Set var_14 = frmGzDepMap.APB
  loc_1104B8C4:   var_3C = var_14
  loc_1104B8C7:   var_14.UnkVCall_00000040h
  loc_1104B903:   var_18.DispID_6803001B = var_20
  loc_1104B92A:   Set var_14 = frmGzDepMap.APB
  loc_1104B939:   var_3C = var_14
  loc_1104B93C:   var_14.UnkVCall_00000040h
  loc_1104B977:   var_18.DispID_6803001B = var_20
  loc_1104B99E:   Set var_14 = frmGzDepMap.APB
  loc_1104B9AD:   var_3C = var_14
  loc_1104B9B0:   var_14.UnkVCall_00000040h
  loc_1104B9EB:   var_18.DispID_6803001B = var_20
  loc_1104BA12:   Set var_14 = frmGzDepMap.APB
  loc_1104BA21:   var_3C = var_14
  loc_1104BA24:   var_14.UnkVCall_00000040h
  loc_1104BA60:   var_18.DispID_6803001B = var_20
  loc_1104BA87:   Set var_14 = frmGzDepMap.TDBText
  loc_1104BA96:   var_3C = var_14
  loc_1104BA99:   var_14.UnkVCall_00000040h
  loc_1104BAD2:   var_18.DispID_000F = var_20
  loc_1104BAF9:   Set var_14 = frmGzDepMap.TDBText
  loc_1104BB08:   var_3C = var_14
  loc_1104BB0B:   var_14.UnkVCall_00000040h
  loc_1104BB44:   var_18.DispID_000F = var_20
  loc_1104BB6B:   Set var_14 = frmGzDepMap.TDBText
  loc_1104BB7A:   var_3C = var_14
  loc_1104BB7D:   var_14.UnkVCall_00000040h
  loc_1104BBB6:   var_18.DispID_000F = var_20
  loc_1104BBFC:   frmGzDepMap.VFG.DispID_8001000D = var_20
  loc_1104BC19:   Set var_14 = frmGzDepMap.TDBText
  loc_1104BC28:   var_3C = var_14
  loc_1104BC2B:   var_14.UnkVCall_00000040h
  loc_1104BC66:   var_18.DispID_0000 = var_20
  loc_1104BC8D:   Set var_14 = frmGzDepMap.TDBText
  loc_1104BC9C:   var_3C = var_14
  loc_1104BC9F:   var_14.UnkVCall_00000040h
  loc_1104BCDA:   var_18.DispID_0000 = var_20
  loc_1104BD20:   frmGzDepMap.TDBText.UnkVCall_00000040h
  loc_1104BD72:   GoTo loc_1104BD88
  loc_1104BD87:   Exit Sub
  loc_1104BD88: End If
  loc_1104BD88: ' Referenced from: 1104BD72
End Sub

Private  Proc_11_9_1104D370(arg_C) '1104D370
  Dim var_54 As Variant
  Dim var_5C As Variant
  Dim var_64 As frmGzDepMap.TDBText
  Dim var_58 As frmGzDepMap.VFG
  Dim var_60 As frmGzDepMap.VFG
  loc_1104D424: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1104D42A: var_178 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1104D432: On Error GoTo loc_1104DF3D
  loc_1104D44C: Set var_54 = frmGzDepMap.TDBText
  loc_1104D45C: var_54.UnkVCall_00000040h
  loc_1104D47A: var_70 = var_58
  loc_1104D492: Set var_5C = frmGzDepMap.TDBText
  loc_1104D4A3: var_5C.UnkVCall_00000040h
  loc_1104D4C1: var_80 = var_60
  loc_1104D536: var_8018 = Proc_0_10_11028DD0(9, var_60 & Proc_0_10_11028DD0(9, var_5C & "SELECT * FROM " & "..T_CY_GZ_SL_DepSetting  WHERE GZ_Dep=", 1) & " AND GZ_Type=", var_54)
  loc_1104D637: var_8024 = ADODB.Recordset.Open(0 & var_8018, var_D4, 0 & var_8018, var_CC, 9)
  loc_1104D684: var_14C = ADODB.Recordset.EOF
  loc_1104D6AA: If var_14C = 0 Then
  loc_1104D6D2:   var_E0 = "提示信息"
  loc_1104D731:   MsgBox("已存在相应的朗新部门名称和类型，无法保存。", 64, var_E0, 10, 10)
  loc_1104D73C: Else
  loc_1104D743:   If arg_C = 1 Then
  loc_1104D75B:     call edi("INSERT INTO ", var_E4, var_E0, var_DC, 00000003h, 00000003h, FFFFFFFFh, var_58, 00000001h)
  loc_1104D762:     var_30 = edi("INSERT INTO ", var_E4, var_E0, var_DC, 00000003h, 00000003h, FFFFFFFFh, var_58, 00000001h)
  loc_1104D76A:     call edi("..T_CY_GZ_SL_DepSetting(GZ_Dep,GZ_Type,UF_DepCode) ", var_30)
  loc_1104D771:     var_28 = edi("..T_CY_GZ_SL_DepSetting(GZ_Dep,GZ_Type,UF_DepCode) ", var_30)
  loc_1104D78A:     Set var_54 = frmGzDepMap.TDBText
  loc_1104D790:     var_150 = var_54
  loc_1104D79F:     var_54.UnkVCall_00000040h
  loc_1104D7C0:     var_58 = 0
  loc_1104D7C7:     var_70 = var_58
  loc_1104D7DF:     Set var_5C = frmGzDepMap.TDBText
  loc_1104D7E5:     var_158 = var_5C
  loc_1104D7F4:     var_5C.UnkVCall_00000040h
  loc_1104D815:     var_60 = 0
  loc_1104D81C:     var_80 = var_60
  loc_1104D837:     Set var_64 = frmGzDepMap.TDBText
  loc_1104D83D:     var_160 = var_64
  loc_1104D84C:     var_64.UnkVCall_00000040h
  loc_1104D86D:     var_68 = 0
  loc_1104D874:     var_90 = var_68
  loc_1104D88D:     call edi("VALUES (", var_28, var_64, 00000002h, var_68, var_5C, 00000001h, var_60, var_54, 00000000h, var_58)
  loc_1104D894:     var_30 = edi("VALUES (", var_28, var_64, 00000002h, var_68, var_5C, 00000001h, var_60, var_54, 00000000h, var_58)
  loc_1104D8A5:     var_34 = Proc_0_10_11028DD0(var_78, var_30)
  loc_1104D8A8:     call edi(var_34)
  loc_1104D8AF:     var_38 = edi(var_34)
  loc_1104D8B7:     call edi(global_1100AC40, var_38)
  loc_1104D8D2:     var_40 = Proc_0_10_11028DD0(var_88, edi(global_1100AC40, var_38))
  loc_1104D8D5:     call edi(var_40)
  loc_1104D8DC:     var_44 = edi(var_40)
  loc_1104D8E4:     call edi(global_1100AC40, var_44)
  loc_1104D8FF:     var_4C = Proc_0_10_11028DD0(var_98, edi(global_1100AC40, var_44))
  loc_1104D902:     call edi(var_4C)
  loc_1104D909:     var_50 = edi(var_4C)
  loc_1104D911:     call edi(global_1100BD88, var_50)
  loc_1104D918:     var_28 = edi(global_1100BD88, var_50)
  loc_1104D973:   Else
  loc_1104D985:     call edi("UPDATE ", 00000003h, var_78, var_88, var_98, 00000003h, var_54, var_5C, var_64, 00000009h, var_30, var_34, var_38, var_3C, var_40, var_44)
  loc_1104D98C:     var_30 = edi("UPDATE ", 00000003h, var_78, var_88, var_98, 00000003h, var_54, var_5C, var_64, 00000009h, var_30, var_34, var_38, var_3C, var_40, var_44)
  loc_1104D994:     call edi("..T_CY_GZ_SL_DepSetting ", var_30, var_48, var_4C, var_50)
  loc_1104D99B:     var_28 = edi("..T_CY_GZ_SL_DepSetting ", var_30, var_48, var_4C, var_50)
  loc_1104D9B4:     Set var_54 = frmGzDepMap.TDBText
  loc_1104D9BA:     var_150 = var_54
  loc_1104D9C9:     var_54.UnkVCall_00000040h
  loc_1104D9EA:     var_58 = 0
  loc_1104D9F1:     var_70 = var_58
  loc_1104DA09:     Set var_5C = frmGzDepMap.TDBText
  loc_1104DA0F:     var_158 = var_5C
  loc_1104DA1E:     var_5C.UnkVCall_00000040h
  loc_1104DA3F:     var_60 = 0
  loc_1104DA46:     var_80 = var_60
  loc_1104DA61:     Set var_64 = frmGzDepMap.TDBText
  loc_1104DA67:     var_160 = var_64
  loc_1104DA76:     var_64.UnkVCall_00000040h
  loc_1104DA97:     var_68 = 0
  loc_1104DA9E:     var_90 = var_68
  loc_1104DAB7:     call edi("SET GZ_Dep=", var_28, var_64, 00000002h, var_68, var_5C, 00000001h, var_60, var_54, 00000000h, var_58)
  loc_1104DAC5:     var_8038 = Proc_0_10_11028DD0(9, edi("SET GZ_Dep=", var_28, var_64, 00000002h, var_68, var_5C, 00000001h, var_60, var_54, 00000000h, var_58))
  loc_1104DACF:     var_34 = var_8038
  loc_1104DAD2:     call edi(var_34)
  loc_1104DAD9:     var_38 = edi(var_34)
  loc_1104DAE1:     call edi(",GZ_Type=", var_38)
  loc_1104DAFC:     var_40 = Proc_0_10_11028DD0(9, edi(",GZ_Type=", var_38))
  loc_1104DAFF:     call edi(var_40)
  loc_1104DB06:     var_44 = edi(var_40)
  loc_1104DB0E:     call edi(",UF_DepCode=", var_44)
  loc_1104DB29:     var_4C = Proc_0_10_11028DD0(9, edi(",UF_DepCode=", var_44))
  loc_1104DB2C:     call edi(var_4C)
  loc_1104DB33:     var_28 = edi(var_4C)
  loc_1104DC48:     var_90 = frmGzDepMap.VFG.DispID_0082(CLng(frmGzDepMap.VFG.DispID_0011), 1)
  loc_1104DD08:     var_C0 = frmGzDepMap.VFG.DispID_0082(CLng(frmGzDepMap.VFG.DispID_0011), 2)
  loc_1104DD21:     call edi(" WHERE GZ_Dep=", var_28)
  loc_1104DD3C:     var_34 = Proc_0_10_11028DD0(8, edi(" WHERE GZ_Dep=", var_28))
  loc_1104DD3F:     call edi(var_34)
  loc_1104DD46:     var_38 = edi(var_34)
  loc_1104DD4E:     call edi(" AND GZ_Type=", var_38)
  loc_1104DD69:     var_40 = Proc_0_10_11028DD0(8, edi(" AND GZ_Type=", var_38))
  loc_1104DD6C:     call edi(var_40)
  loc_1104DD73:     var_28 = edi(var_40)
  loc_1104DDD2:   End If
  loc_1104DE10:   var_78 = UnkObj.UnkVCall_00000040h
  loc_1104DE92:   frmGzDepMap.Pic1.DispID_80010007 = var_D0
  loc_1104DF19:   MsgBox("成功保存。", 64, "提示信息", 10, 10)
  loc_1104DF38:   GoTo loc_1104E00B
  loc_1104DF3D:   var_805C = Err
  loc_1104DF48:   Set var_54 = Err
  loc_1104DF57:   var_30 = Err.Description
  loc_1104DFE3:   MsgBox(0, 16, "提示信息", 10, 10)
  loc_1104DFF2: End If
  loc_1104E00B: ' Referenced from: 1104DF38
  loc_1104E016: Exit Sub
  loc_1104E021: GoTo loc_1104E0A8
  loc_1104E0A7: Exit Sub
  loc_1104E0A8: ' Referenced from: 1104E021
End Sub

Private Sub Proc_11_10_1104E0E0
  Dim var_44 As Variant
  Dim var_48 As Variant
  Dim var_110 As Label
  loc_1104E16A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1104E172: var_130 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1104E17A: On Error GoTo loc_1104E803
  loc_1104E193: var_8004 = UnkObj.UnkVCall_00000044h
  loc_1104E21E: If 1 <= CLng(frmGzDepMap.VFG.DispID_0007)(-1) Then
  loc_1104E2BA:   var_8014 = (frmGzDepMap.VFG.DispID_0082(1, frmGzDepMap.VFG.DispID_0007) = global_1100EFEC)
  loc_1104E2E8:   var_8018 = Unknown_2DA80(1)
  loc_1104E2FC:   Set var_48 = frmGzDepMap.Label3
  loc_1104E302:   var_110 = var_48
  loc_1104E3D1:   var_48.Caption = "正在处理：第[" & frmGzDepMap.VFG.DispID_0082(1, 2) & "]的数据。"
  loc_1104E463:   frmGzDepMap.Pic1.DispID_80010007 = True
  loc_1104E490:   frmGzDepMap.Pic1.DispID_FFFFFDDA
  loc_1104E52C:   var_60 = frmGzDepMap.VFG.DispID_0082(var_20, 1)
  loc_1104E59D:   var_78 = frmGzDepMap.VFG.DispID_0082(var_20, 2)
  loc_1104E5AD:   var_80 = var_78
  loc_1104E61C:   var_8044 = Proc_0_10_11028DD0(8,  & Proc_0_10_11028DD0(8,  & "DELETE FROM " & "..T_CY_GZ_SL_DepSetting  WHERE GZ_Dep=", ) & " AND GZ_Type=", )
  loc_1104E630:   var_24 = fs:[00000000h] & var_8044
  loc_1104E6B4:   var_58 = UnkObj.UnkVCall_00000040h
  loc_1104E6F3:   var_20 = 1+var_20
  loc_1104E6F8:   GoTo loc_1104E217
  loc_1104E6FD: End If
  loc_1104E70C: var_804C = = var_78.Name
  loc_1104E77A: frmGzDepMap.Pic1.DispID_80010007 = var_90
  loc_1104E7F8: MsgBox("成功删除。", 64, "提示信息", 10, 10)
  loc_1104E7FE: GoTo loc_1104E92A
  loc_1104E803: ' Referenced from: 1104E17A
  loc_1104E851: frmGzDepMap.Pic1.DispID_80010007 = var_90
  loc_1104E871: var_8050 = UnkObj.UnkVCall_0000004Ch
  loc_1104E88F: var_8054 = Err
  loc_1104E89A: Set var_44 = Err
  loc_1104E8A5: var_2C = Err.Description
  loc_1104E91F: MsgBox(0, 16, "提示信息", 10, 10)
  loc_1104E92A: ' Referenced from: 1104E7FE
  loc_1104E948: Exit Sub
  loc_1104E953: GoTo loc_1104E9A4
  loc_1104E9A3: Exit Sub
  loc_1104E9A4: ' Referenced from: 1104E953
End Sub
