VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C908002B2F49FB}#1.2#0"; "C:\Windows\SysWow64\Comdlg32.ocx"
Begin VB.Form frmXSCBToPz
  Caption = "凭证导入（SS销售成本结转凭证）"
  BackColor = &H80000005&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmXSCBToPz.frx":0000
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 345
  ClientWidth = 9255
  ClientHeight = 6000
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
    OleObjectBlob = "frmXSCBToPz.frx":014A
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
    Height = 6000
    TabStop = 0   'False
    TabIndex = 0
    OleObjectBlob = "frmXSCBToPz.frx":0293
    Begin AIFCmp1.asxStatusBar sBar
      Left = 0
      Top = 5655
      Width = 12045
      Height = 345
      OleObjectBlob = "frmXSCBToPz.frx":04BC
    End
    Begin C1SizerLibCtl.C1Elastic C1Elastic2
      Left = 0
      Top = 0
      Width = 12045
      Height = 435
      TabStop = 0   'False
      TabIndex = 1
      OleObjectBlob = "frmXSCBToPz.frx":05EC
    End
    Begin VSFlex8LCtl.VSFlexGrid VFG
      Left = 0
      Top = 1260
      Width = 12045
      Height = 4380
      TabIndex = 2
      OleObjectBlob = "frmXSCBToPz.frx":0757
      Begin MSComDlg.CommonDialog dlg
        OleObjectBlob = "frmXSCBToPz.frx":0BC0
        Left = 0
        Top = 0
      End
    End
    Begin AIFCmp1.asxPanel asxPanel1
      Left = 0
      Top = 450
      Width = 12045
      Height = 795
      OleObjectBlob = "frmXSCBToPz.frx":0C24
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
        Visible = 0   'False
        TabIndex = 9
        OleObjectBlob = "frmXSCBToPz.frx":0D04
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 4
        Left = 8595
        Top = 435
        Width = 720
        Height = 270
        Visible = 0   'False
        TabIndex = 13
        OleObjectBlob = "frmXSCBToPz.frx":0EA4
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
        OleObjectBlob = "frmXSCBToPz.frx":1044
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 1
        Left = 6090
        Top = 435
        Width = 870
        Height = 270
        TabIndex = 8
        OleObjectBlob = "frmXSCBToPz.frx":123C
      End
      Begin TDBText6Ctl.TDBText TDBText
        Left = 30
        Top = 435
        Width = 5115
        Height = 270
        TabIndex = 10
        OleObjectBlob = "frmXSCBToPz.frx":140C
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
        OleObjectBlob = "frmXSCBToPz.frx":1568
      End
      Begin TDBDate6Ctl.TDBDate TDBDate
        Left = 9390
        Top = 420
        Width = 2385
        Height = 285
        TabIndex = 12
        OleObjectBlob = "frmXSCBToPz.frx":170C
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
  Begin VB.Menu mnuSet
    Visible = 0   'False
    Caption = "设置"
    Begin VB.Menu mnuKM
      Caption = "科目映射"
    End
  End
End

Attribute VB_Name = "frmXSCBToPz"


Private Sub TDBText_UnknownEvent_B '11048780
  Dim var_64 As frmXSCBToPz.dlg
  loc_110487E7: Set var_64 = frmXSCBToPz.dlg
  loc_11048819: var_64.FileName = var_48
  loc_1104883E: var_64.DialogTitle = var_48
  loc_11048863: var_64.Filter = var_48
  loc_11048885: var_64.CancelError = var_48
  loc_1104888F: var_64.ShowOpen
  loc_110488A1: var_64.FileName = var_64
  loc_110488E7: If (var_64 = global_1100AE28) Then
  loc_110488F5:   var_64.FileName = Me
  loc_1104893D:   frmXSCBToPz.TDBText.DispID_0000 = var_2C
  loc_11048967: End If
  loc_11048973: GoTo loc_1104899B
  loc_1104899A: Exit Sub
  loc_1104899B: ' Referenced from: 11048973
End Sub

Private Sub Form_Load() '11035600
  Dim var_18 As Variant
  Dim var_1C As var_18.DispID_03E8
  loc_1103566A: Set var_18 = frmXSCBToPz.TDBText
  loc_11035671: var_2C = var_18.DispID_03E8
  loc_11035692: var_18.DispID_03E8.UnkVCall_00000030h
  loc_110356E0: Set var_18 = frmXSCBToPz.TDBDate
  loc_110356E7: var_2C = var_18.DispID_03E8
  loc_110356FC: Set var_1C = var_18.DispID_03E8
  loc_11035708: var_1C.UnkVCall_00000030h
  loc_11035777: frmXSCBToPz.TDBDate.DispID_0000 = Date
  loc_1103579B: Set var_18 = frmXSCBToPz.APB
  loc_110357A8: var_18.UnkVCall_00000040h
  loc_11035800: var_8004 = frmXSCBToPz.Proc_10_9_1102F930(var_18)
  loc_11035812: GoTo loc_11035831
  loc_11035830: Exit Sub
  loc_11035831: ' Referenced from: 11035812
End Sub

Private Sub Form_Resize() '11035860
  loc_110358ED: var_38 = frmXSCBToPz.Pic1.DispID_80010005
  loc_11035911: var_48 = frmXSCBToPz.Pic1.DispID_80010006
  loc_11035924: var_EC = var_48.ScaleWidth
  loc_1103595B: If global_110F6000 = 0 Then
  loc_11035965: Else
  loc_11035970: End If
  loc_11035970: var_70 = ((var_EC - CSgn(var_38)) / 2)
  loc_11035985: var_F0 = var_48.ScaleHeight
  loc_110359C3: If global_110F6000 = 0 Then
  loc_110359CD: Else
  loc_110359D8: End If
  loc_11035AE3: frmXSCBToPz.Pic1.DispID_80011002(var_70, ((var_F0 - CSgn(var_48)) / 2), CSgn(frmXSCBToPz.Pic1.DispID_80010005), CSgn(frmXSCBToPz.Pic1.DispID_80010006))
  loc_11035B2C: GoTo loc_11035B66
End Sub

Private  APB_UnknownEvent_9(arg_C) '11048300
  Dim var_20 As Variant
  Dim var_AC As Scripting.FileSystemObject
  loc_11048377: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11048380: var_C4 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110483A7: arg_C = frmXSCBToPz.APB.UnkVCall_00000040h
  loc_110483E5: var_B8 = var_24.DispID_FFFFFDFA
  loc_11048419: var_8008 = (var_B8 = "加载数据")
  loc_1104841D: If var_8008 = 0 Then
  loc_11048440:   var_AC = var_18
  loc_1104845B:   Set var_20 = frmXSCBToPz.TDBText
  loc_1104847B:   var_1C = frmXSCBToPz.TDBText.DispID_0000
  loc_1104848B:   var_8014 = Scripting.FileSystemObject.FileExists
  loc_110484C9:   If Not (var_A8) Then
  loc_1104852C:     MsgBox("文件不存在或非法路径！ ", 64, "提示", 10, 10)
  loc_11048552:   Else
  loc_11048564:     If frmXSCBToPz.FillData < 0 Then
  loc_11048576:       var_A8 = CheckObj(%ecx = %S_edx_S = %S_edx_S, global_1100D168, 1788)
  loc_11048581:     End If
  loc_1104858D:     call edi("取消加载", var_B8, var_1C, var_A8, var_24)
  loc_11048591:     If edi("取消加载", var_B8, var_1C, var_A8, var_24) = 0 Then
  loc_110485C1:       var_44 = "提示信息"
  loc_110485EF:       var_2C = "是否取消数据载入？" & vbCrLf & "取消数据载入，数据将全部清空。"
  loc_1104860B:       MsgBox(var_2C, 292, var_44, var_54, var_64)
  loc_11048645:       If (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6) = 0 Then GoTo loc_11048705
  loc_11048656:     Else
  loc_11048662:       (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_11048666:       If eax = 0 Then
  loc_1104866B:         var_8020 = frmXSCBToPz.Proc_10_12_1103D110("凭证导入")
  loc_11048676:       Else
  loc_11048682:         (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_11048686:         If var_B8 = 0 Then
  loc_1104868B:           var_8024 = .Proc_10_14_110489D0("导出")
  loc_11048693:         Else
  loc_1104869F:           (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_110486A3:           If var_8024 = 0 Then
  loc_110486D6:             Set var_20 = var_20 = %S_edx_S
  loc_110486E4:             var_802C = Global.Unload var_B8
  loc_11048705:           End If
  loc_11048705:         End If
  loc_11048705:       End If
  loc_11048705:     End If
  loc_11048705:   End If
  loc_11048705: End If
  loc_1104870D: GoTo loc_11048744
  loc_11048743: Exit Sub
  loc_11048744: ' Referenced from: 1104870D
End Sub

Public Sub ExitForm(Cancel, UnloadMode) '1102F850
  Dim var_18 As Global
  loc_1102F88F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1102F8BA: Set var_18 = Me
  loc_1102F8C2: var_8008 = Global.Unload
  loc_1102F8FC: GoTo loc_1102F908
  loc_1102F907: Exit Sub
  loc_1102F908: ' Referenced from: 1102F8FC
End Sub

Public Function FillData() '11031400
  Dim var_A0 As Variant
  Dim var_80 As Variant
  Dim var_A4 As Variant
  Dim var_64 As Variant
  Dim var_44 As Variant
  loc_1103152F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11031535: var_230 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11031589: frmXSCBToPz.VFG.DispID_0007 = 1
  loc_110315AC: Set var_A0 = frmXSCBToPz.Label3
  loc_110315AE: var_20C = var_A0
  loc_110315BC: var_A0.Caption = "正在打开Excel数据表，请稍候。。。"
  loc_1103162F: frmXSCBToPz.Pic1.DispID_80010007 = True
  loc_1103165B: frmXSCBToPz.Pic1.DispID_FFFFFDDA
  loc_11031672: On Error GoTo loc_11035119
  loc_1103167D: var_8004 = CreateObject(global_1100D5A4)
  loc_11031688: Set var_80 = CreateObject(global_1100D5A4)
  loc_11031697: var_A0 = var_80.UnkVCall_000000D0h
  loc_110316BE: var_210 = var_A0
  loc_1103194A: Set var_A4 = frmXSCBToPz.TDBText
  loc_11031974: var_94 = var_A4.DispID_0000
  loc_11031984: var_94 = var_A0.UnkVCall_0000004Ch
  loc_110319F7: var_A0 = var_64.Tag
  loc_11031A71: var_44.BackColor = CInt(1)
  loc_11031A9A: var_A0.Activate
  loc_11031B04: Set var_8C = var_A0.UsedRange
  loc_11031B6A: Set var_A0 = frmXSCBToPz.Pic1
  loc_11031B71: var_A0.DispID_80010007 = var_130
  loc_11031BDB: var_44.UnkVCall_00000064h
  loc_11031C71: var_C8 = var_A0.Cells(1, 5).value
  loc_11031C90: var_8014 = (Proc_0_11_11029000(var_C8, var_44, 2) = "借方")
  loc_11031C9E: var_1E0 = var_8014
  loc_11031CE5: var_C8.BackColor = CInt(1)
  loc_11031E63: If CBool(var_8014 Or (LCase(Proc_0_11_11029000(var_A4.Cells(1, 7).value, var_A4, var_134)) <> "贷方")) Then
  loc_11031EB7:   frmXSCBToPz.Pic1.DispID_80010007 = var_130
  loc_11031EDA:   Set var_A0 = frmXSCBToPz.TDBText
  loc_11031F9A:   var_138 = var_64.UnkVCall_0000006Ch
  loc_11031FD3:   var_134 = var_80.UnkVCall_00000398h
  loc_11032008:   Set var_44 = {000208D7-0000-0000-C000000000000046}()
  loc_11032018:   Set var_64 = {000208DA-0000-0000-C000000000000046}()
  loc_11032028:   Set var_80 = {000208D5-0000-0000-C000000000000046}()
  loc_110320C0:   MsgBox("与所要求的格式不符！ ", 64, "提示", 10, 10)
  loc_110320F2: Else
  loc_11032103:   Set var_A0 = frmXSCBToPz.Label3
  loc_11032109:   var_20C = var_A0
  loc_11032117:   var_A0.Caption = "正在分析数据，请稍候。。。"
  loc_1103218E:   frmXSCBToPz.Pic1.DispID_80010007 = True
  loc_110321BF:   frmXSCBToPz.Pic1.DispID_FFFFFDDA
  loc_110321F9:   Set var_A0 = frmXSCBToPz.APB
  loc_110321FF:   var_20C = var_A0
  loc_11032211:   var_A0.UnkVCall_00000040h
  loc_110322A7:   Set var_A0 = frmXSCBToPz.APB
  loc_110322AD:   var_20C = var_A0
  loc_110322BF:   var_A0.UnkVCall_00000040h
  loc_11032355:   Set var_A0 = frmXSCBToPz.APB
  loc_11032369:   var_A0.UnkVCall_00000040h
  loc_110323FD:   var_150 = "条记录"
  loc_11032430:   var_B8 = var_8C.Rows
  loc_11032459:   var_D8 = var_B8.Count - 1
  loc_1103247E:   var_F8 = var_140 & var_D8 & var_150
  loc_11032481:   call edi(var_F8, var_A0, 00000002h, var_A4, var_A0, 00000001h, var_A4, var_A0, 00000000h, var_A4, var_130, var_12C, var_B8, var_B4, var_A0, var_AC)
  loc_11032483:   var_100 = edi(var_F8, var_A0, 00000002h, var_A4, var_A0, 00000001h, var_A4, var_A0, 00000000h, var_A4, var_130, var_12C, var_B8, var_B4, var_A0, var_AC)
  loc_110324FD:   frmXSCBToPz.sBar.DispID_6803001E(var_100)
  loc_1103258C:   frmXSCBToPz.VFG.DispID_0007 = var_130
  loc_110325B6:   Set var_A0 = frmXSCBToPz.TDBDate
  loc_110325C4:   var_B8 = var_A0.DispID_004E
  loc_110325CE:   call edi(var_B8, 0000000Ah, var_144, 80020004h, var_13C, 00000409h, var_130, var_12C, var_A0, var_138, var_134, var_130, var_12C, var_148, var_144, var_140)
  loc_110325D0:   var_C0 = edi(var_B8, 0000000Ah, var_144, 80020004h, var_13C, 00000409h, var_130, var_12C, var_A0, var_138, var_134, var_130, var_12C, var_148, var_144, var_140)
  loc_110325EE:   var_D8)
  loc_11032640:   var_8024 = CByte("DateToPeriod".00000001h)
  loc_11032691:   Set var_A0 = frmXSCBToPz.TDBDate
  loc_1103269F:   var_B8 = var_A0.DispID_004E
  loc_110326A9:   call edi(var_B8, var_13C, var_158, var_154, var_150, var_14C, 0000000Ah, var_164, 80020004h, var_15C, 0000000Ah, var_174, 80020004h, var_16C, 0000000Ah, var_184)
  loc_110326AB:   var_C0 = edi(var_B8, var_13C, var_158, var_154, var_150, var_14C, 0000000Ah, var_164, 80020004h, var_15C, 0000000Ah, var_174, 80020004h, var_16C, 0000000Ah, var_184)
  loc_110326D6:   call edi(Year(var_C0), 80020004h, var_17C, 0000000Ah, var_194, 80020004h, var_18C, 0000000Ah, var_1A4, 80020004h, var_19C, 0000000Ah, var_1B4)
  loc_11032713:   var_50 = "转"
  loc_11032746:   var_C8 = var_8C.Rows.Count
  loc_11032785:   If var_20 <= CLng(var_C8) Then
  loc_11032793:     If global_56 = 0 Then
  loc_11032800:       var_44.UnkVCall_00000064h
  loc_11032851:       var_C8.BackColor = CInt(1)
  loc_110328EC:       var_802C = Proc_0_11_11029000(var_A4.Cells(var_20, 2).value, var_A4, var_44)
  loc_1103290E:       var_248 = (stk@FEC4(00000002h, var_134, 1, var_12C, var_A0) = global_1100AE28) + 1
  loc_1103299C:       var_8034 = Proc_0_11_11029000(var_A0.Cells(var_20, 1).value)
  loc_110329C6:       var_214 = (stk@FEC4 = global_1100AE28) + 1
  loc_11032A27:       If var_214 = 0 Then
  loc_11032A41:         Set var_A0 = frmXSCBToPz.Cbo
  loc_11032A47:         var_20C = var_A0
  loc_11032A87:         var_803C = var_A0.Text & global_1100D6EC
  loc_11032A9C:         var_8040 = stk@FEC4 & var_58
  loc_11032AE0:         Set var_A4 = frmXSCBToPz.TDBDate
  loc_11032AEE:         var_B8 = var_A4.DispID_004E
  loc_11032AF8:         call edi(var_B8)
  loc_11032B5B:         call edi(stk@FEC4 & global_1100D700 & Month(edi(var_B8)) & 1100D708h)
  loc_11032BD4:         var_60 = "1"
  loc_11032C39:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11032C9C:         Set var_A0 = frmXSCBToPz.TDBDate
  loc_11032CAA:         var_D8 = var_A0.DispID_004E
  loc_11032CB4:         call edi(var_D8)
  loc_11032CEF:         call edi(var_60 & Chr(9) & edi(var_D8))
  loc_11032D90:         call edi(var_60 & Chr(9) & var_50)
  loc_11032E18:         call edi(var_60 & Chr(9) & 1100C008h)
  loc_11032E9F:         call edi(var_60 & Chr(9) & var_5C)
  loc_11032F44:         var_44.UnkVCall_00000064h
  loc_11032FDA:         var_E8 = var_A0.Cells(var_20, 6).value
  loc_11033022:         call edi(var_60 & Chr(9) & Proc_0_11_11029000(var_E8, var_44, 2), var_144, 1, var_13C, var_A0)
  loc_1103317E:         var_E8 = var_A0.Cells(var_20, 4).value
  loc_110331C6:         call edi(var_60 & Chr(9) & Proc_0_12_110291B0(var_E8, var_A0))
  loc_1103326F:         call edi(var_60 & Chr(9) & 1100C008h)
  loc_110333F2:         call edi(var_60 & Chr(9) & Proc_0_12_110291B0(var_A0.Cells(var_20, 3).value, var_A0))
  loc_1103349B:         call edi(var_60 & Chr(9) & 1100C008h)
  loc_11033523:         call edi(var_60 & Chr(9) & 1100C008h)
  loc_11033583:         var_C8 = var_60 & Chr(9)
  loc_110335B5:         call edi( & "cUserName".00000000h)
  loc_11033644:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_110336CC:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11033751:         var_D8 = var_60 & Chr(9) & 1100AE28h
  loc_11033754:         call edi(var_D8)
  loc_1103388F:         var_E8 = var_A0.Cells(var_20, 5).value
  loc_110338D7:         call edi(var_60 & Chr(9) & Proc_0_11_11029000(var_E8, var_A0))
  loc_11033980:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11033A08:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11033A90:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11033B18:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11033C96:         call edi(var_60 & Chr(9) & Proc_0_11_11029000(var_A0.Cells(var_20, 2).value, var_A0))
  loc_11033D2F:         frmXSCBToPz.VFG.DispID_0080(var_60)
  loc_11033D4C:         var_60 = "1"
  loc_11033DB1:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11033E14:         Set var_A0 = frmXSCBToPz.TDBDate
  loc_11033E22:         var_D8 = var_A0.DispID_004E
  loc_11033E2C:         call edi(var_D8)
  loc_11033E67:         call edi(var_60 & Chr(9) & edi(var_D8))
  loc_11033F08:         call edi(var_60 & Chr(9) & var_50)
  loc_11033F90:         call edi(var_60 & Chr(9) & 1100C008h)
  loc_11034017:         call edi(var_60 & Chr(9) & var_5C)
  loc_110340BC:         var_44.UnkVCall_00000064h
  loc_11034152:         var_E8 = var_A0.Cells(var_20, 8).value
  loc_1103419A:         call edi(var_60 & Chr(9) & Proc_0_11_11029000(var_E8, var_44, 2), var_144, 1, var_13C, var_A0)
  loc_11034243:         call edi(var_60 & Chr(9) & 1100C008h)
  loc_1103437E:         var_E8 = var_A0.Cells(var_20, 4).value
  loc_110343C6:         call edi(var_60 & Chr(9) & Proc_0_12_110291B0(var_E8, var_A0))
  loc_1103456A:         call edi(var_60 & Chr(9) & Proc_0_12_110291B0(var_A0.Cells(var_20, 3).value, var_A0))
  loc_11034613:         call edi(var_60 & Chr(9) & 1100C008h)
  loc_1103469B:         call edi(var_60 & Chr(9) & 1100C008h)
  loc_110346FB:         var_C8 = var_60 & Chr(9)
  loc_1103472D:         call edi( & "cUserName".00000000h)
  loc_110347BC:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11034844:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_110348C9:         var_D8 = var_60 & Chr(9) & 1100AE28h
  loc_110348CC:         call edi(var_D8)
  loc_11034A07:         var_E8 = var_A0.Cells(var_20, 7).value
  loc_11034A4F:         call edi(var_60 & Chr(9) & Proc_0_11_11029000(var_E8, var_A0))
  loc_11034AF8:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11034B80:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11034C08:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11034C90:         call edi(var_60 & Chr(9) & 1100AE28h)
  loc_11034E0E:         call edi(var_60 & Chr(9) & Proc_0_11_11029000(var_A0.Cells(var_20, 2).value, var_A0))
  loc_11034EA7:         frmXSCBToPz.VFG.DispID_0080(var_60)
  loc_11034ECA:         var_20 = 1+var_20
  loc_11034ECD:         GoTo loc_1103277F
  loc_11034ED2:       End If
  loc_11034F23:       frmXSCBToPz.VFG.DispID_0007 = 1
  loc_11034F35:       global_56 = 0
  loc_11034F70:       frmXSCBToPz.APB.UnkVCall_00000040h
  loc_11035010:       frmXSCBToPz.APB.UnkVCall_00000040h
  loc_110350B2:       frmXSCBToPz.APB.UnkVCall_00000040h
  loc_110350FE:       var_A4.DispID_80010007 = var_130
  loc_11035119:     End If
  loc_11035119:   End If
  loc_11035157:   frmXSCBToPz.APB.UnkVCall_00000040h
  loc_110351A9:   var_A4.DispID_80010007 = var_130
  loc_110351E9:   Set var_A0 = frmXSCBToPz.APB
  loc_110351EB:   var_20C = var_A0
  loc_110351FD:   var_A0.UnkVCall_00000040h
  loc_1103528F:   Set var_A0 = frmXSCBToPz.APB
  loc_11035291:   var_20C = var_A0
  loc_110352A3:   var_A0.UnkVCall_00000040h
  loc_11035353:   frmXSCBToPz.Pic1.DispID_80010007 = var_130
  loc_11035398:   var_B0 = frmXSCBToPz.TDBText
  loc_1103542C:   var_138 = var_64.UnkVCall_0000006Ch
  loc_11035465:   var_134 = var_80.UnkVCall_00000398h
  loc_1103549A:   Set var_44 = {000208D7-0000-0000-C000000000000046}()
  loc_110354AA:   Set var_64 = {000208DA-0000-0000-C000000000000046}()
  loc_110354BA:   Set var_80 = {000208D5-0000-0000-C000000000000046}()
  loc_110354CB: End If
  loc_110354CB: Exit Sub
  loc_110354D6: GoTo loc_11035556
  loc_11035555: Exit Function
  loc_11035556: ' Referenced from: 110354D6
End Function

Public Function GetKmCode(pDepCode, pKmCode) '11049150
  Dim var_30 As ADODB.Recordset
  Dim var_4C As Me
  loc_110491C8: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110491D0: var_B0 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110491F3: On Error GoTo loc_110494E4
  loc_110491FC: var_68 = pKmCode
  loc_1104920A: var_78 = pDepCode
  loc_1104926F: var_8018 = Proc_0_10_11028DD0(var_80,  & Proc_0_10_11028DD0(&H4008, 1 & "SELECT * FROM " & "..T_CY_KmSetting WHERE cKMCode=", ) & " AND cDepCode=", )
  loc_11049283: var_20 =  & var_8018
  loc_11049323: var_8024 = ADODB.Recordset.Open(var_20, var_6C, var_20, var_64, 9)
  loc_11049380: var_84 = ADODB.Recordset.EOF
  loc_1104939F: If var_84 = 0 Then
  loc_110493C0:   var_4C = ADODB.Recordset.Fields
  loc_110493DE:   var_68 = "cKmcodeUF"
  loc_11049406:   ADODB.Recordset.8 = Forms
  loc_11049457:   var_24 = var_60
  loc_110494A1:   If ADODB.Recordset.Close < 0 Then
  loc_110494AF:     var_8038 = CheckObj(var_30, global_1100ADFC, 128)
  loc_110494B3:   End If
  loc_110494D4:   If ADODB.Recordset.Close < 0 Then
  loc_110494E4:   End If
  loc_110494F6:   Set var_30 = ADODB.Recordset()
  loc_1104950A: End If
  loc_1104950A: Exit Sub
  loc_11049515: GoTo loc_11049563
  loc_1104951B: If var_C Then
  loc_11049526: End If
  loc_11049562: Exit Function
  loc_11049563: ' Referenced from: 11049515
End Function

Public Function GetUFDepCode(pDepCode) '110495B0
  Dim var_2C As ADODB.Recordset
  Dim var_3C As Me
  loc_11049619: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11049627: var_28 = pDepCode
  loc_11049634: On Error GoTo loc_110498E2
  loc_1104963D: var_58 = var_28
  loc_11049696: var_20 =  & Proc_0_10_11028DD0(&H4008, 1 & "SELECT * FROM " & "..T_CY_DeptSetting WHERE cDepCode=", )
  loc_1104971B: var_8018 = ADODB.Recordset.Open(var_20, var_5C, var_20, var_54, 9)
  loc_11049772: var_74 = ADODB.Recordset.EOF
  loc_1104978B: If var_74 = 0 Then
  loc_110497AF:   var_3C = ADODB.Recordset.Fields
  loc_110497CD:   var_58 = "cDepcodeUF"
  loc_110497F5:   ADODB.Recordset.8 = Forms
  loc_11049846:   var_24 = var_50
  loc_11049890:   If ADODB.Recordset.Close < 0 Then
  loc_1104989E:     var_802C = CheckObj(var_2C, global_1100ADFC, 128)
  loc_110498A2:   End If
  loc_110498A8:   var_24 = var_28
  loc_110498D2:   If ADODB.Recordset.Close < 0 Then
  loc_110498E2:   End If
  loc_110498F4:   Set var_2C = ADODB.Recordset()
  loc_11049900:   var_24 = var_28
  loc_11049906: End If
  loc_11049906: Exit Sub
  loc_11049911: GoTo loc_11049953
  loc_11049917: If var_C Then
  loc_11049922: End If
  loc_11049952: Exit Function
  loc_11049953: ' Referenced from: 11049911
End Function

Public Function GetUFXMCode(pDepCode) '110499A0
  Dim var_2C As ADODB.Recordset
  Dim var_3C As Me
  loc_11049A09: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11049A17: var_28 = pDepCode
  loc_11049A24: On Error GoTo loc_11049CD2
  loc_11049A2D: var_58 = var_28
  loc_11049A86: var_20 =  & Proc_0_10_11028DD0(&H4008, 1 & "SELECT * FROM " & "..T_CY_XMSetting WHERE cDepCode=", )
  loc_11049B0B: var_8018 = ADODB.Recordset.Open(var_20, var_5C, var_20, var_54, 9)
  loc_11049B62: var_74 = ADODB.Recordset.EOF
  loc_11049B7B: If var_74 = 0 Then
  loc_11049B9F:   var_3C = ADODB.Recordset.Fields
  loc_11049BBD:   var_58 = "cXMcodeUF"
  loc_11049BE5:   ADODB.Recordset.8 = Forms
  loc_11049C36:   var_24 = var_50
  loc_11049C80:   If ADODB.Recordset.Close < 0 Then
  loc_11049C8E:     var_802C = CheckObj(var_2C, global_1100ADFC, 128)
  loc_11049C92:   End If
  loc_11049C98:   var_24 = var_28
  loc_11049CC2:   If ADODB.Recordset.Close < 0 Then
  loc_11049CD2:   End If
  loc_11049CE4:   Set var_2C = ADODB.Recordset()
  loc_11049CF0:   var_24 = var_28
  loc_11049CF6: End If
  loc_11049CF6: Exit Sub
  loc_11049D01: GoTo loc_11049D43
  loc_11049D07: If var_C Then
  loc_11049D12: End If
  loc_11049D42: Exit Function
  loc_11049D43: ' Referenced from: 11049D01
End Function

Private Sub Proc_10_9_1102F930
  Dim var_58 As frmXSCBToPz.VFG
  loc_1102F971: Set var_58 = frmXSCBToPz.VFG
  loc_1102F9C2: var_58.DispID_005D = frmXSCBToPz.VFG
  loc_1102FA03: var_58.DispID_0067 = frmXSCBToPz.VFG
  loc_1102FA22: var_58.DispID_0041 = frmXSCBToPz.VFG
  loc_1102FB93: var_58.DispID_008A(4)
  loc_1102FBD6: var_58.DispID_0079(400)
  loc_1102FC3C: var_58.DispID_007B(True)
  loc_1102FC81: var_58.DispID_0090("业务号")
  loc_1102FCC4: var_58.DispID_0078(700)
  loc_1102FD07: var_58.DispID_0077(4)
  loc_1102FD4F: var_58.DispID_0090("状态")
  loc_1102FD95: var_58.DispID_0078(700)
  loc_1102FDDB: var_58.DispID_0077(4)
  loc_1102FE23: var_58.DispID_0090("制单日期")
  loc_1102FE69: var_58.DispID_0078(1000)
  loc_1102FEAF: var_58.DispID_0077(1)
  loc_1102FEF4: var_58.DispID_0090("凭证类别字")
  loc_1102FF36: var_58.DispID_0078(1000)
  loc_1102FF78: var_58.DispID_0077(4)
  loc_1102FFC0: var_58.DispID_0090("附单据数")
  loc_11030006: var_58.DispID_0078(800)
  loc_1103004A: var_58.DispID_0077(var_3C)
  loc_11030092: var_58.DispID_0090(var_3C)
  loc_110300D8: var_58.DispID_0078(var_3C)
  loc_1103011E: var_58.DispID_0077(var_3C)
  loc_11030166: var_58.DispID_0090(var_3C)
  loc_110301AC: var_58.DispID_0078(var_3C)
  loc_110301F2: var_58.DispID_0077(var_3C)
  loc_1103023A: var_58.DispID_0090(var_3C)
  loc_11030280: var_58.DispID_0078(var_3C)
  loc_110302C8: var_58.DispID_009C(var_3C)
  loc_1103030C: var_58.DispID_0077(var_3C)
  loc_11030354: var_58.DispID_0090(var_3C)
  loc_1103039A: var_58.DispID_0078(var_3C)
  loc_110303E2: var_58.DispID_009C(var_3C)
  loc_11030428: var_58.DispID_0077(var_3C)
  loc_11030470: var_58.DispID_0090(var_3C)
  loc_110304B6: var_58.DispID_0078(var_3C)
  loc_110304FE: var_58.DispID_009C(var_3C)
  loc_11030544: var_58.DispID_0077(var_3C)
  loc_1103058C: var_58.DispID_0090(var_3C)
  loc_110305D2: var_58.DispID_0078(var_3C)
  loc_1103061A: var_58.DispID_009C(var_3C)
  loc_11030660: var_58.DispID_0077(var_3C)
  loc_110306A8: var_58.DispID_0090(var_3C)
  loc_110306EE: var_58.DispID_0078(var_3C)
  loc_11030736: var_58.DispID_009C(var_3C)
  loc_1103077C: var_58.DispID_0077(var_3C)
  loc_110307C4: var_58.DispID_0090(var_3C)
  loc_1103080A: var_58.DispID_0078(var_3C)
  loc_11030852: var_58.DispID_0090(var_3C)
  loc_11030898: var_58.DispID_0078(var_3C)
  loc_110308E0: var_58.DispID_0090(var_3C)
  loc_11030926: var_58.DispID_0078(var_3C)
  loc_1103096E: var_58.DispID_0090(var_3C)
  loc_110309B4: var_58.DispID_0078(var_3C)
  loc_110309FC: var_58.DispID_0090(var_3C)
  loc_11030A42: var_58.DispID_0078(var_3C)
  loc_11030A8A: var_58.DispID_0090(var_3C)
  loc_11030AD0: var_58.DispID_0078(var_3C)
  loc_11030B18: var_58.DispID_0090(var_3C)
  loc_11030B5E: var_58.DispID_0078(var_3C)
  loc_11030BA6: var_58.DispID_0090(var_3C)
  loc_11030BEC: var_58.DispID_0078(var_3C)
  loc_11030C34: var_58.DispID_0090(var_3C)
  loc_11030C7A: var_58.DispID_0078(var_3C)
  loc_11030CC2: var_58.DispID_0090(var_3C)
  loc_11030D08: var_58.DispID_0078(var_3C)
  loc_11030D50: var_58.DispID_0090(var_3C)
  loc_11030D96: var_58.DispID_0078(var_3C)
  loc_11030DDE: var_58.DispID_0090(var_3C)
  loc_11030E26: var_58.DispID_0090(var_3C)
  loc_11030E6E: var_58.DispID_0090(var_3C)
  loc_11030EB6: var_58.DispID_0090(var_3C)
  loc_11030EFE: var_58.DispID_0090(var_3C)
  loc_11030F46: var_58.DispID_0090(var_3C)
  loc_11030F8E: var_58.DispID_0090(var_3C)
  loc_11030FD4: var_58.DispID_0078(var_3C)
  loc_1103101A: var_58.DispID_00AC(var_3C)
  loc_11031060: var_58.DispID_00AC(var_3C)
  loc_110310A6: var_58.DispID_00AC(var_3C)
  loc_110310EC: var_58.DispID_00AC(var_3C)
  loc_11031132: var_58.DispID_00AC(var_3C)
  loc_11031178: var_58.DispID_00AC(var_3C)
  loc_110311BE: var_58.DispID_00AC(var_3C)
  loc_11031204: var_58.DispID_00AC(var_3C)
  loc_1103124A: var_58.DispID_00AC(var_3C)
  loc_11031290: var_58.DispID_00AC(var_3C)
  loc_110312D6: var_58.DispID_00AC(var_3C)
  loc_110312F2: If 22 <= &H1C Then
  loc_11031332:   var_58.DispID_00AC(var_3C)
  loc_1103134A:   var_14 = 1+var_14
  loc_1103134D:   GoTo loc_110312EE
  loc_1103134F: End If
  loc_11031362: If 12 <= &H1D Then
  loc_110313A2:   var_58.DispID_0077(var_3C)
  loc_110313B6:   var_14 = 1+var_14
  loc_110313B9:   GoTo loc_1103135E
  loc_110313BB: End If
End Sub

Private Sub Proc_10_10_11035B90
  Dim var_7C As Variant
  Dim var_1F8 As Label
  Dim var_80 As Variant
  Dim var_88 As frmXSCBToPz.Label3
  loc_11035C7A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11035C82: var_228 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11035C88: var_8004 = ecx
  loc_11035CFE: If var_14 <= CLng(frmXSCBToPz.VFG.DispID_0007)(-1) Then
  loc_11035D0F:   var_800C = frmXSCBToPz.Proc_10_11_11037A30(vbNull)
  loc_11035DAD:   frmXSCBToPz.VFG.DispID_0082(29, var_58)
  loc_11035E91:   If (frmXSCBToPz.VFG.DispID_0082(var_14, 29) = global_1100AE28) + 1 Then
  loc_11035F11:     frmXSCBToPz.VFG.DispID_0082(1, 285267764)
  loc_11036045:     frmXSCBToPz.VFG.DispID_009E(var_14, 1, var_14, 1, 16711680)
  loc_11036065:     Set var_7C = frmXSCBToPz.Label3
  loc_11036072:     var_1F8 = var_7C
  loc_110360BC:     var_7C.Caption = "分析: 第(" & CStr(vbNull) & ")行信息----有效"
  loc_1103610E:     frmXSCBToPz.Pic1.DispID_FFFFFDDA
  loc_11036121:   Else
  loc_1103619B:     frmXSCBToPz.VFG.DispID_0082(1, 285267820)
  loc_110362CF:     frmXSCBToPz.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110362EF:     Set var_80 = frmXSCBToPz.Label3
  loc_110362FC:     var_1F8 = var_80
  loc_110363DD:     var_80.Caption = "分析:   第(" & CStr(vbNull) & ")行信息----" & frmXSCBToPz.VFG.DispID_0082(var_14, 29)
  loc_11036448:     frmXSCBToPz.Pic1.DispID_FFFFFDDA
  loc_1103645A:   End If
  loc_1103646A:   var_14 = 1+var_14
  loc_1103646D:   GoTo loc_11035CF0
  loc_11036472: End If
  loc_110364D9: If var_14 <= CLng(frmXSCBToPz.VFG.DispID_0007)(-1) Then
  loc_11036551:   var_A0 = frmXSCBToPz.VFG.DispID_0082(var_14, 2)
  loc_1103656F:   var_B8)
  loc_110366FF:   var_8048 = frmXSCBToPz.VFG.DispID_0082(var_14, frmXSCBToPz.VFG)
  loc_11036736:   var_4C = CCur(0)
  loc_11036739:   var_48 = var_8048
  loc_11036745:   var_40 = CCur(0)
  loc_11036748:   var_3C = var_8048
  loc_11036754:   var_34 = var_14
  loc_1103675D:   var_30 = var_14
  loc_11036766:   var_160 = CByte("DateToPeriod".00000001h)
  loc_11036803:   var_B8)
  loc_11036882:   Set var_80 = frmXSCBToPz.VFG
  loc_110368A8:   var_8064 = (frmXSCBToPz.VFG.DispID_0082(var_14, 3) = var_80.DispID_0082(var_14, 3))
  loc_110368D5:   var_1A0 = var_8064 + 1
  loc_1103694F:   var_806C = (var_8048 = frmXSCBToPz.VFG.DispID_0082(var_14, ""))
  loc_11036976:   var_1E0 = var_806C + 1
  loc_11036A78:   If CBool((frmXSCBToPz.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_8064 + 1 And var_806C + 1) Then
  loc_11036B3D:     If (frmXSCBToPz.VFG.DispID_0082(var_14, 29) = global_1100AE28) Then
  loc_11036B46:     End If
  loc_11036B4B:     If var_24 = 0 Then
  loc_11036BF4:       var_16C = var_48
  loc_11036C38:       var_9C = var_1F0
  loc_11036C84:       var_4C = CCur(var_4C + Format(Val(frmXSCBToPz.VFG.DispID_0082(var_14, 7)), "#.00"))
  loc_11036C87:       var_48 = var_D8
  loc_11036D67:       var_16C = var_3C
  loc_11036DAB:       var_9C = var_1F0
  loc_11036DF7:       var_40 = CCur(var_40 + Format(Val(frmXSCBToPz.VFG.DispID_0082(var_14, 8)), "#.00"))
  loc_11036DFA:       var_3C = var_D8
  loc_11036E3A:     End If
  loc_11036E5B:     var_14 = var_14(1)
  loc_11036E5E:     var_30 = var_30(1)
  loc_11036E80:     var_80A0 = CLng(frmXSCBToPz.VFG.DispID_0007)
  loc_11036E9B:     var_1F8 = (var_14 > 0)
  loc_11036EBF:     If var_1F8 = 0 Then GoTo loc_11036760
  loc_11036EC5:   End If
  loc_11036ECA:   If var_24 = 0 Then
  loc_11036EDE:     Set var_7C = frmXSCBToPz.Chk
  loc_11036EE9:     var_1F8 = var_7C
  loc_11036EEF:     Set var_80 = var_7C(1)
  loc_11036F1A:     var_200 = var_80
  loc_11036F20:     var_1EC = var_80.Value
  loc_11036F74:     If (var_1EC = 1) Then
  loc_11036FA4:       If (Abs(var_4C - var_40) <> 0.01) >= 0 Then
  loc_11036FAD:       End If
  loc_11036FAD:     End If
  loc_11036FB2:     If var_24 Then
  loc_11036FB8:     End If
  loc_11036FD8:     var_1C = var_34
  loc_11036FDD:     If var_34 <= (var_30 - 1) Then
  loc_110370A1:       If (frmXSCBToPz.VFG.DispID_0082(var_1C, 29) = global_1100AE28) + 1 Then
  loc_11037129:         frmXSCBToPz.VFG.DispID_0082(1, 285267820)
  loc_110371BD:         frmXSCBToPz.VFG.DispID_0082(29, "凭证借贷不平衡或某分录有错误")
  loc_110372F1:         frmXSCBToPz.VFG.DispID_009E(var_1C, 1, var_1C, 1, 255)
  loc_11037303:       End If
  loc_11037313:       GoTo loc_11036FD2
  loc_11037318:     End If
  loc_11037329:     var_44 = var_44(1)
  loc_1103733A:     Set var_88 = frmXSCBToPz.Label3
  loc_1103736D:     var_1F8 = var_88
  loc_1103747A:     Set var_80 = frmXSCBToPz.VFG
  loc_11037552:     var_80D4 = "分析: 第[" & frmXSCBToPz.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmXSCBToPz.VFG.DispID_0082(var_34, var_14)
  loc_11037574:     var_78 = var_80D4 & "]号凭证借贷不平衡"
  loc_11037588:     var_88.Caption = var_78
  loc_1103758F:     If var_78 < 0 Then
  loc_11037595:       GoTo loc_11037813
  loc_1103759A:     End If
  loc_110375AB:     var_20 = var_20(1)
  loc_110375BC:     Set var_88 = frmXSCBToPz.Label3
  loc_110375EF:     var_1F8 = var_88
  loc_110376FC:     Set var_80 = frmXSCBToPz.VFG
  loc_110377D4:     var_80F8 = "分析: 第[" & frmXSCBToPz.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmXSCBToPz.VFG.DispID_0082(var_34, frmXSCBToPz.VFG.DispID_0082(var_34, var_14))
  loc_110377F6:     var_78 = var_80F8 & "]号凭证有效"
  loc_1103780A:     var_88.Caption = var_78
  loc_11037811:     If var_78 >= 0 Then GoTo loc_11037822
  loc_11037813:     ' Referenced from: 11037595
  loc_1103781C:     var_78 = CheckObj(var_1F8, global_1100D574, 84)
  loc_11037822:   End If
  loc_110378A4:   frmXSCBToPz.Pic1.DispID_FFFFFDDA
  loc_110378D5:   var_14 = 1+var_14(-1)
  loc_110378D8:   GoTo loc_110364D3
  loc_110378DD: End If
  loc_110378E2: If var_44 > 0 Then
  loc_110378E9:   If var_20 > 0 Then
  loc_11037904:   Else
  loc_1103791D:   Else
  loc_11037927:     var_8108 = frmXSCBToPz.Proc_10_13_11047510(var_1EC)
  loc_11037935:     If var_1EC Then
  loc_11037950:     Else
  loc_11037958:       var_18 = ecx
  loc_11037961:       GoTo loc_110379FB
  loc_110379FA:       Exit Sub
  loc_110379FB:     End If
  loc_110379FB:   End If
  loc_110379FB: End If
  loc_110379FB: ' Referenced from: 11037961
End Sub

Private  Proc_10_11_11037A30(arg_C) '11037A30
  Dim var_58 As frmXSCBToPz.VFG
  Dim var_20 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA19-EB96-11D2-AF06000021009B21}()
  Dim var_18 As {3302AA41-EB96-11D2-AF06000021009B21}()
  Dim var_1C As {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_11037B2C: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11037B3C: var_210 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11037C1B: If (frmXSCBToPz.VFG.DispID_0082(arg_C, 2) = global_1100AE28) + 1 Then
  loc_11037C25:   var_24 = "制单日期为空"
  loc_11037C36: Else
  loc_11037CD1:   var_78 = frmXSCBToPz.VFG.DispID_0082(arg_C, 2)
  loc_11037D0B:   If Proc_0_9_11028500(var_80, global_1103D0E7, ) Then
  loc_11037DB4:     var_78 = frmXSCBToPz.VFG.DispID_0082(arg_C, 2)
  loc_11037DBE:     var_90)
  loc_11037DD0:     var_48 = var_90
  loc_11037E02:     var_118 = var_48
  loc_11037E10:     var_114 = var_44
  loc_11037E44:     var_80 = "AccountOpen".0.0
  loc_11037E75:     If (var_80 < var_80) Then
  loc_11037E7F:       var_24 = "日期超前总账系统启用日期"
  loc_11037E90:     Else
  loc_11037E96:       var_154 = var_44
  loc_11037E9C:       var_1A4 = var_44
  loc_11037EAD:       var_158 = var_48
  loc_11037EB3:       var_1A8 = var_48
  loc_11037F6E:       var_80 = "AccountYMD".0.00000002h("AccountYMD".0, var_13C)
  loc_1103806A:       If CBool( Or ((global_1103D0E7 < var_80) > "AccountYMD".0.00000002h(var_180, var_18C))) Then
  loc_11038074:         var_24 = "日期必须在当前会计年度内"
  loc_11038085:       Else
  loc_110380A2:         var_118 = var_48
  loc_110380F6:         var_80 = "DateToPeriod".00000001h - 1
  loc_11038184:         If CBool("AccountYMD".0.00000001h) Then
  loc_1103818E:           var_24 = "已结账月份不能制单"
  loc_1103819F:         Else
  loc_1103827B:           If (frmXSCBToPz.VFG.DispID_0082(arg_C, 3) = global_1100AE28) + 1 Then
  loc_11038285:             var_24 = "凭证类别字为空"
  loc_11038296:           Else
  loc_11038325:             var_8034 = frmXSCBToPz.VFG.DispID_0082(arg_C, 3)
  loc_11038335:             var_80 = 8
  loc_11038338:             var_78 = var_8034
  loc_1103837F:             var_8038 = CBool(Not("pzlbCheck".00000001h(, fs:[00000000h], , global_1103D0E7, global_1103D0E7, var_74, var_8034, var_7C)))
  loc_110383B6:             If var_8038 Then
  loc_110383C0:               var_24 = "凭证类别字非法"
  loc_110383D1:             Else
  loc_110384A8:               If (frmXSCBToPz.VFG.DispID_0082(arg_C, var_128) = global_1100AE28) + 1 Then
  loc_110384B2:                 var_24 = "业务号为空"
  loc_110384C3:               Else
  loc_1103854D:                 var_8044 = frmXSCBToPz.VFG.DispID_0082(arg_C, var_128)
  loc_1103855D:                 var_80 = 8
  loc_11038560:                 var_78 = var_8044
  loc_110385A3:                 var_90 = "GenLen".00000001h(fs:[00000000h], , global_1103D0E7, global_1103D0E7, global_1103D0E7, var_74, var_8044, var_7C)
  loc_110385EB:                 If (var_90 > 30) Then
  loc_110385F5:                   var_24 = "业务号超长"
  loc_11038606:                 Else
  loc_110386E5:                   If (frmXSCBToPz.VFG.DispID_0082(arg_C, 5) = global_1100AE28) + 1 Then
  loc_110386EF:                     var_24 = "摘要为空"
  loc_11038700:                   Else
  loc_1103884D:                     var_80 = frmXSCBToPz.VFG.DispID_0082(arg_C, 5)
  loc_11038968:                     If (((InStr(1, frmXSCBToPz.VFG.DispID_0082(arg_C, 5), "|", 0) > 0) Or (InStr(1, var_80, """", 0) > 0)) Or (InStr(1, frmXSCBToPz.VFG.DispID_0082(arg_C, 5), "'", 0) > 0)) Then
  loc_11038972:                       var_24 = "摘要含有非法字符"
  loc_11038983:                     Else
  loc_11038A15:                       var_806C = frmXSCBToPz.VFG.DispID_0082(arg_C, 5)
  loc_11038A25:                       var_80 = 8
  loc_11038A28:                       var_78 = var_806C
  loc_11038A6B:                       var_90 = "GenLen".00000001h(global_1103D0E7, global_1103D0E7, global_1103D0E7, global_1103D0E7, global_1103D0E7, var_74, var_806C, var_7C)
  loc_11038AB4:                       If (var_90 > 120) Then
  loc_11038ABE:                         var_24 = "摘要超长"
  loc_11038ACF:                       Else
  loc_11038BAC:                         If (frmXSCBToPz.VFG.DispID_0082(arg_C, 6) = global_1100AE28) + 1 Then
  loc_11038BB6:                           var_24 = "科目为空"
  loc_11038BC7:                         Else
  loc_11038C56:                           var_807C = frmXSCBToPz.VFG.DispID_0082(arg_C, 6)
  loc_11038C66:                           var_80 = 8
  loc_11038C69:                           var_78 = var_807C
  loc_11038CE9:                           var_40 = "kmCheck".00000002h(var_807C, var_150, var_15C)
  loc_11038D1B:                           var_8084 = (var_40 = global_1100AE28)
  loc_11038D23:                           If var_8084 = 0 Then
  loc_11038D2D:                             var_24 = "科目非法"
  loc_11038D3E:                           Else
  loc_11038D92:                             var_118 = arg_C
  loc_11038DFA:                             frmXSCBToPz.VFG.DispID_0082(6, var_40)
  loc_11038E19:                             var_118 = var_40
  loc_11038E66:                             var_128 = var_20
  loc_11038EB4:                             "kmCodeToProperties".00000002h
  loc_11038ED1:                             Set var_20 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_11038EFF:                             var_1F0 = var_20
  loc_11038F05:                             var_1D4 = var_20.UnkVCall_00000114h
  loc_11038F31:                             If var_1D4 = 0 Then
  loc_11038F3B:                               var_24 = "科目非末级"
  loc_11038F4C:                             Else
  loc_11038F6F:                               var_1F0 = var_20
  loc_11038FA1:                               If var_20.UnkVCall_00000174h Then
  loc_11038FAB:                                 var_24 = "科目已封存"
  loc_11038FBC:                               Else
  loc_1103909A:                                 If (frmXSCBToPz.VFG.DispID_0082(arg_C, 7) = global_1100AE28) Then
  loc_11039176:                                   If Not (IsNumeric(frmXSCBToPz.VFG.DispID_0082(arg_C, 7))) Then
  loc_11039180:                                     var_24 = "借方金额非法"
  loc_11039191:                                   Else
  loc_1103923A:                                     var_80A8 = CDbl(Val(frmXSCBToPz.VFG.DispID_0082(arg_C, 7)))
  loc_110392E8:                                     var_80 = frmXSCBToPz.VFG.DispID_0082(arg_C, 7)
  loc_11039310:                                     var_228 = CDbl(Val(var_80))
  loc_11039326:                                     var_80B4 = CDbl(-9999999999999.99)
  loc_1103933E:                                     GoTo loc_11039342
  loc_11039390:                                     If (eax Or 0) Then
  loc_1103939A:                                       var_24 = "借方金额超范围"
  loc_110393AB:                                     Else
  loc_110393AB:                                     End If
  loc_11039489:                                     If (frmXSCBToPz.VFG.DispID_0082(arg_C, 8) = global_1100AE28) Then
  loc_11039565:                                       If Not (IsNumeric(frmXSCBToPz.VFG.DispID_0082(arg_C, 8))) Then
  loc_1103956F:                                         var_24 = "贷方金额非法"
  loc_11039580:                                       Else
  loc_11039629:                                         var_80CC = CDbl(Val(frmXSCBToPz.VFG.DispID_0082(arg_C, 8)))
  loc_110396D7:                                         var_80 = frmXSCBToPz.VFG.DispID_0082(arg_C, 8)
  loc_110396FF:                                         var_234 = CDbl(Val(var_80))
  loc_11039715:                                         var_80D8 = CDbl(-9999999999999.99)
  loc_1103972D:                                         GoTo loc_11039731
  loc_1103977F:                                         If (eax Or 0) Then
  loc_11039789:                                           var_24 = "贷方金额超范围"
  loc_1103979A:                                         Else
  loc_1103979A:                                         End If
  loc_11039923:                                         var_74 = var_1E0
  loc_11039995:                                         var_C4 = var_1E8
  loc_11039A0F:                                         var_80EC = (Format(Val(frmXSCBToPz.VFG.DispID_0082(arg_C, 7)), "#.00") <> 0) And (Format(Val(frmXSCBToPz.VFG.DispID_0082(arg_C, 8)), "#.00") <> 0)
  loc_11039A88:                                         If CBool(var_80EC) Then
  loc_11039A92:                                           var_24 = "借方金额和贷方金额不能同时不为0"
  loc_11039AA3:                                         Else
  loc_11039C2C:                                           var_74 = var_1E0
  loc_11039C9E:                                           var_C4 = var_1E8
  loc_11039D18:                                           var_8104 = (Format(Val(frmXSCBToPz.VFG.DispID_0082(arg_C, 7)), "#.00") = 0) And (Format(Val(frmXSCBToPz.VFG.DispID_0082(arg_C, 8)), "#.00") = 0)
  loc_11039D91:                                           If CBool(var_8104) Then
  loc_11039D9B:                                             var_24 = "借方金额和贷方金额不能同时为0"
  loc_11039DAC:                                           Else
  loc_11039DCC:                                             var_1F0 = var_20
  loc_11039E1E:                                             If (var_20.UnkVCall_0000007Ch = global_1100AE28) Then
  loc_11039F02:                                               If (frmXSCBToPz.VFG.DispID_0082(arg_C, 9) = global_1100AE28) Then
  loc_11039FDE:                                                 If Not (IsNumeric(frmXSCBToPz.VFG.DispID_0082(arg_C, 9))) Then
  loc_11039FE8:                                                   var_24 = "数量数值非法"
  loc_11039FF9:                                                 Else
  loc_11039FF9:                                                 End If
  loc_11039FF9:                                               End If
  loc_1103A019:                                               var_1F0 = var_20
  loc_1103A06B:                                               If (var_20.UnkVCall_0000006Ch = global_1100AE28) Then
  loc_1103A14F:                                                 If (frmXSCBToPz.VFG.DispID_0082(arg_C, 10) = global_1100AE28) Then
  loc_1103A22B:                                                   If Not (IsNumeric(frmXSCBToPz.VFG.DispID_0082(arg_C, 10))) Then
  loc_1103A235:                                                     var_24 = "外币金额非法"
  loc_1103A246:                                                   Else
  loc_1103A2EF:                                                     var_8140 = CDbl(Val(frmXSCBToPz.VFG.DispID_0082(arg_C, 10)))
  loc_1103A3C5:                                                     var_240 = CDbl(Val(frmXSCBToPz.VFG.DispID_0082(arg_C, 10)))
  loc_1103A3DB:                                                     var_814C = CDbl(-9999999999999.99)
  loc_1103A3F3:                                                     GoTo loc_1103A3F7
  loc_1103A445:                                                     If (eax Or 0) Then
  loc_1103A44F:                                                       var_24 = "外币超范围"
  loc_1103A460:                                                     Else
  loc_1103A460:                                                     End If
  loc_1103A53E:                                                     If (frmXSCBToPz.VFG.DispID_0082(arg_C, 11) = global_1100AE28) Then
  loc_1103A61A:                                                       If Not (IsNumeric(frmXSCBToPz.VFG.DispID_0082(arg_C, 11))) Then
  loc_1103A624:                                                         var_24 = "汇率数值非法"
  loc_1103A635:                                                       Else
  loc_1103A635:                                                       End If
  loc_1103A635:                                                     End If
  loc_1103A658:                                                     var_1F0 = var_20
  loc_1103A68A:                                                     If var_20.UnkVCall_0000010Ch Then
  loc_1103A76E:                                                       If (frmXSCBToPz.VFG.DispID_0082(arg_C, 13) = global_1100AE28) Then
  loc_1103A805:                                                         var_816C = frmXSCBToPz.VFG.DispID_0082(arg_C, 13)
  loc_1103A818:                                                         var_78 = var_816C
  loc_1103A847:                                                         var_90 = "JsfsCheck".00000001h(global_1103D0E7, global_1103D0E7, global_1103D0E7, global_1103D0E7, global_1103D0E7, var_74, var_816C, var_7C)
  loc_1103A897:                                                         If CBool(Not(var_90)) Then
  loc_1103A8A1:                                                           var_24 = "结算方式非法"
  loc_1103A8B2:                                                         Else
  loc_1103A8B2:                                                         End If
  loc_1103A8B2:                                                       End If
  loc_1103A8D5:                                                       var_1F0 = var_20
  loc_1103A8DB:                                                       var_1D4 = var_20.UnkVCall_0000010Ch
  loc_1103A922:                                                       var_1F8 = var_20
  loc_1103A928:                                                       var_1D8 = var_20.UnkVCall_00000094h
  loc_1103A96F:                                                       var_200 = var_20
  loc_1103A9C7:                                                       If (var_20.UnkVCall_0000009Ch = 0) = 0 Then
  loc_1103AAAB:                                                         If (frmXSCBToPz.VFG.DispID_0082(arg_C, 14) = global_1100AE28) Then
  loc_1103AB42:                                                           var_8188 = frmXSCBToPz.VFG.DispID_0082(arg_C, 14)
  loc_1103AB55:                                                           var_78 = var_8188
  loc_1103AB98:                                                           var_90 = "GenLen".00000001h(global_1103D0E7, global_1103D0E7, global_1103D0E7, global_1103D0E7, global_1103D0E7, var_74, var_8188, var_7C)
  loc_1103ABE1:                                                           If (var_90 > 10) Then
  loc_1103ABEB:                                                             var_24 = "票号超长"
  loc_1103ABFC:                                                           Else
  loc_1103ABFC:                                                           End If
  loc_1103ACDA:                                                           If (frmXSCBToPz.VFG.DispID_0082(arg_C, 15) = global_1100AE28) Then
  loc_1103AD71:                                                             var_8198 = frmXSCBToPz.VFG.DispID_0082(arg_C, 15)
  loc_1103AD84:                                                             var_78 = var_8198
  loc_1103ADB3:                                                             var_90 = "DateCheck".00000001h(global_1103D0E7, global_1103D0E7, global_1103D0E7, global_1103D0E7, global_1103D0E7, var_74, var_8198, var_7C)
  loc_1103AE03:                                                             If CBool(Not(var_90)) Then
  loc_1103AE0D:                                                               var_24 = "票号发生日期非法"
  loc_1103AE1E:                                                             Else
  loc_1103B000:                                                               var_81A8 = (Format(frmXSCBToPz.VFG.DispID_0082(arg_C, 15), "yyyy-mm-dd") > Format(frmXSCBToPz.VFG.DispID_0082(arg_C, 2), "yyyy-mm-dd"))
  loc_1103B062:                                                               If var_81A8 Then
  loc_1103B06C:                                                                 var_24 = "票号发生日期大于制单日期"
  loc_1103B07D:                                                               Else
  loc_1103B07D:                                                               End If
  loc_1103B07D:                                                             End If
  loc_1103B0A0:                                                             var_1F0 = var_20
  loc_1103B0ED:                                                             var_1F8 = var_20
  loc_1103B0F3:                                                             var_1D8 = var_20.UnkVCall_0000008Ch
  loc_1103B156:                                                             If (var_20.UnkVCall_000000A4h = 0) = 0 Then
  loc_1103B215:                                                               If (frmXSCBToPz.VFG.DispID_0082(arg_C, 16) = global_1100AE28) Then
  loc_1103B2BF:                                                                 var_78 = frmXSCBToPz.VFG.DispID_0082(arg_C, 16)
  loc_1103B33F:                                                                 var_38 = "BmCheck".00000002h(var_154, 0, var_15C)
  loc_1103B371:                                                                 var_81C4 = (var_38 = global_1100AE28)
  loc_1103B379:                                                                 If var_81C4 = 0 Then
  loc_1103B383:                                                                   var_24 = "部门非法"
  loc_1103B394:                                                                 Else
  loc_1103B3B7:                                                                   var_118 = arg_C
  loc_1103B452:                                                                   frmXSCBToPz.VFG.DispID_0082(16, var_38)
  loc_1103B487:                                                                   var_1F0 = var_20
  loc_1103B4B9:                                                                   If var_20.UnkVCall_000000A4h Then
  loc_1103B4C7:                                                                     var_118 = var_38
  loc_1103B519:                                                                     var_128 = var_28
  loc_1103B567:                                                                     "BmToProperties".00000002h
  loc_1103B584:                                                                     Set var_28 = {3302AA19-EB96-11D2-AF06000021009B21}()
  loc_1103B5B2:                                                                     var_1F0 = var_28
  loc_1103B5B8:                                                                     var_1D4 = var_28.UnkVCall_00000034h
  loc_1103B5DE:                                                                     If var_1D4 = 0 Then
  loc_1103B5EC:                                                                       var_24 = "部门非末级"
  loc_1103B5FD:                                                                     Else
  loc_1103B605:                                                                       var_24 = "部门为空"
  loc_1103B616:                                                                     Else
  loc_1103B6AE:                                                                       frmXSCBToPz.VFG.DispID_0082(var_128, 1100AE28h)
  loc_1103B6C0:                                                                     End If
  loc_1103B6C0:                                                                   End If
  loc_1103B6E3:                                                                   var_1F0 = var_20
  loc_1103B715:                                                                   If var_20.UnkVCall_0000008Ch Then
  loc_1103B7C1:                                                                     var_81DC = (frmXSCBToPz.VFG.DispID_0082(arg_C, &H11) = global_1100AE28)
  loc_1103B7F9:                                                                     If var_81DC Then
  loc_1103B8A5:                                                                       var_81E4 = (frmXSCBToPz.VFG.DispID_0082(arg_C, 16) = global_1100AE28)
  loc_1103B901:                                                                       If var_81E4 + 1 Then
  loc_1103B984:                                                                         var_78 = frmXSCBToPz.VFG.DispID_0082(arg_C, &H11)
  loc_1103BA12:                                                                         var_90 = "ZyCheck".00000003h(var_174, "BmCheck".00000002h(var_154, 80020004h, var_15C), var_17C)
  loc_1103BA27:                                                                         var_34 = var_90
  loc_1103BA59:                                                                         var_81F0 = (var_34 = global_1100AE28)
  loc_1103BA61:                                                                         If var_81F0 = 0 Then
  loc_1103BA6B:                                                                           var_24 = "职员非法"
  loc_1103BA7C:                                                                         Else
  loc_1103BA9F:                                                                           var_118 = arg_C
  loc_1103BB3A:                                                                           frmXSCBToPz.VFG.DispID_0082(&H11, var_34)
  loc_1103BB59:                                                                           var_118 = var_34
  loc_1103BBA6:                                                                           var_128 = var_18
  loc_1103BBF4:                                                                           "ZyToProperties".00000002h
  loc_1103BC11:                                                                           Set var_18 = {3302AA41-EB96-11D2-AF06000021009B21}()
  loc_1103BC1F:                                                                           var_118 = arg_C
  loc_1103BC60:                                                                           var_1F0 = var_18
  loc_1103BD19:                                                                           frmXSCBToPz.VFG.DispID_0082(var_128, var_18.UnkVCall_0000002Ch)
  loc_1103BD39:                                                                         Else
  loc_1103BDAF:                                                                           var_158 = var_38
  loc_1103BDBC:                                                                           var_78 = frmXSCBToPz.VFG.DispID_0082(8, var_128)
  loc_1103BE71:                                                                           var_34 = "ZyCheck".00000003h(var_164, 0, var_16C)
  loc_1103BEA3:                                                                           var_8204 = (var_34 = global_1100AE28)
  loc_1103BEAB:                                                                           If var_8204 = 0 Then
  loc_1103BEB5:                                                                             var_24 = "职员不在指定部门内"
  loc_1103BEC6:                                                                           Else
  loc_1103BF1A:                                                                             var_118 = arg_C
  loc_1103BF82:                                                                             frmXSCBToPz.VFG.DispID_0082(&H11, var_34)
  loc_1103BF94:                                                                           End If
  loc_1103BF94:                                                                         End If
  loc_1103BF94:                                                                       End If
  loc_1103BFB7:                                                                       var_1F0 = var_20
  loc_1103BFE9:                                                                       If var_20.UnkVCall_00000094h Then
  loc_1103C095:                                                                         var_8210 = (frmXSCBToPz.VFG.DispID_0082(arg_C, &H12) = global_1100AE28)
  loc_1103C0A6:                                                                         var_1F0 = var_8210
  loc_1103C0CD:                                                                         If var_1F0 = 0 Then GoTo loc_1103C5D2
  loc_1103C177:                                                                         var_78 = frmXSCBToPz.VFG.DispID_0082(arg_C, &H12)
  loc_1103C1F7:                                                                         var_3C = "KhCheck".00000002h(var_154, 0, var_15C)
  loc_1103C229:                                                                         var_821C = (var_3C = global_1100AE28)
  loc_1103C231:                                                                         If var_821C = 0 Then
  loc_1103C23B:                                                                           var_24 = "客户非法"
  loc_1103C24C:                                                                         Else
  loc_1103C2A0:                                                                           var_118 = arg_C
  loc_1103C308:                                                                           frmXSCBToPz.VFG.DispID_0082(&H12, var_3C)
  loc_1103C31A:                                                                         End If
  loc_1103C33D:                                                                         var_1F0 = var_20
  loc_1103C36F:                                                                         If var_20.UnkVCall_0000009Ch Then
  loc_1103C41B:                                                                           var_8228 = (frmXSCBToPz.VFG.DispID_0082(arg_C, &H13) = global_1100AE28)
  loc_1103C42C:                                                                           var_1F0 = var_8228
  loc_1103C453:                                                                           If var_1F0 = 0 Then GoTo loc_1103C9A2
  loc_1103C4FD:                                                                           var_78 = frmXSCBToPz.VFG.DispID_0082(arg_C, &H13)
  loc_1103C57D:                                                                           var_30 = "GysCheck".00000002h(var_154, 0, var_15C)
  loc_1103C5AF:                                                                           var_8234 = (var_30 = global_1100AE28)
  loc_1103C5B7:                                                                           If var_8234 = 0 Then
  loc_1103C5C1:                                                                             var_24 = "供应商非法"
  loc_1103C5CD:                                                                             GoTo loc_1103D0A8
  loc_1103C5DA:                                                                             var_24 = "客户为空"
  loc_1103C5EB:                                                                           Else
  loc_1103C63F:                                                                             var_118 = arg_C
  loc_1103C6A7:                                                                             frmXSCBToPz.VFG.DispID_0082(&H13, var_30)
  loc_1103C6B9:                                                                           End If
  loc_1103C6DC:                                                                           var_1F0 = var_20
  loc_1103C729:                                                                           var_1F8 = var_20
  loc_1103C72F:                                                                           var_1D8 = var_20.UnkVCall_0000009Ch
  loc_1103C76D:                                                                           If (var_20.UnkVCall_00000094h = 0) = 0 Then
  loc_1103C819:                                                                             var_8244 = (frmXSCBToPz.VFG.DispID_0082(arg_C, &H14) = global_1100AE28)
  loc_1103C851:                                                                             If var_8244 Then
  loc_1103C8E8:                                                                               var_8248 = frmXSCBToPz.VFG.DispID_0082(arg_C, &H14)
  loc_1103C8FB:                                                                               var_78 = var_8248
  loc_1103C93E:                                                                               var_90 = "GenLen".00000001h(global_1103D0E7, global_1103D0E7, 1, global_1103D0E7, global_1103D0E7, var_74, var_8248, var_7C)
  loc_1103C987:                                                                               If (var_90 > 20) Then
  loc_1103C991:                                                                                 var_24 = "业务员超长"
  loc_1103C99D:                                                                                 GoTo loc_1103D0A8
  loc_1103C9AA:                                                                                 var_24 = "供应商为空"
  loc_1103C9BB:                                                                               Else
  loc_1103C9BB:                                                                               End If
  loc_1103C9BB:                                                                             End If
  loc_1103C9DB:                                                                             var_1F0 = var_20
  loc_1103CA33:                                                                             If (var_20.UnkVCall_000000ACh = global_1100AE28) Then
  loc_1103CADF:                                                                               var_825C = (frmXSCBToPz.VFG.DispID_0082(arg_C, &H15) = global_1100AE28)
  loc_1103CB17:                                                                               If var_825C Then
  loc_1103CB3D:                                                                                 var_1F0 = var_20
  loc_1103CB70:                                                                                 var_8264 = (var_20.UnkVCall_000000ACh = global_1100AE28)
  loc_1103CB95:                                                                                 If var_8264 Then
  loc_1103CBBB:                                                                                   var_1F0 = var_20
  loc_1103CC99:                                                                                   var_88 = frmXSCBToPz.VFG.DispID_0082(arg_C, &H15)
  loc_1103CD21:                                                                                   var_A0 = "XmCheck".00000003h(var_164, Format(var_20.UnkVCall_000000ACh, 8), var_16C)
  loc_1103CD36:                                                                                   var_2C = var_A0
  loc_1103CD6F:                                                                                   var_8274 = (var_2C = global_1100AE28)
  loc_1103CD77:                                                                                   If var_8274 = 0 Then
  loc_1103CD81:                                                                                     var_24 = "项目非法"
  loc_1103CD92:                                                                                   Else
  loc_1103CDBE:                                                                                     var_4C = var_20.UnkVCall_000000ACh
  loc_1103CDEC:                                                                                     var_128 = var_2C
  loc_1103CE1D:                                                                                     Set var_58 = var_1C
  loc_1103CE9F:                                                                                     "XmToProperties".00000003h
  loc_1103CEBC:                                                                                     Set var_1C = {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_1103CF11:                                                                                     If var_1C.UnkVCall_00000034h Then
  loc_1103CF1F:                                                                                       var_24 = "项目已结算"
  loc_1103CF30:                                                                                     Else
  loc_1103CF80:                                                                                       var_118 = %cobj
  loc_1103CFE8:                                                                                       frmXSCBToPz.VFG.DispID_0082(&H15, 1100AE28h)
  loc_1103D005:                                                                                     Else
  loc_1103D00D:                                                                                       var_24 = "制单日期非法"
  loc_1103D013:                                                                                     End If
  loc_1103D013:                                                                                   End If
  loc_1103D013:                                                                                 End If
  loc_1103D019:                                                                                 GoTo loc_1103D0A8
  loc_1103D022:                                                                                 If var_4 Then
  loc_1103D02D:                                                                                 End If
  loc_1103D0A7:                                                                                 Exit Sub
  loc_1103D0A8:                                                                               End If
  loc_1103D0A8:                                                                             End If
  loc_1103D0A8:                                                                           End If
  loc_1103D0A8:                                                                         End If
  loc_1103D0A8:                                                                       End If
  loc_1103D0A8:                                                                     End If
  loc_1103D0A8:                                                                   End If
  loc_1103D0A8:                                                                 End If
  loc_1103D0A8:                                                               End If
  loc_1103D0A8:                                                             End If
  loc_1103D0A8:                                                           End If
  loc_1103D0A8:                                                         End If
  loc_1103D0A8:                                                       End If
  loc_1103D0A8:                                                     End If
  loc_1103D0A8:                                                   End If
  loc_1103D0A8:                                                 End If
  loc_1103D0A8:                                               End If
  loc_1103D0A8:                                             End If
  loc_1103D0A8:                                           End If
  loc_1103D0A8:                                         End If
  loc_1103D0A8:                                       End If
  loc_1103D0A8:                                     End If
  loc_1103D0A8:                                   End If
  loc_1103D0A8:                                 End If
  loc_1103D0A8:                               End If
  loc_1103D0A8:                             End If
  loc_1103D0A8:                           End If
  loc_1103D0A8:                         End If
  loc_1103D0A8:                       End If
  loc_1103D0A8:                     End If
  loc_1103D0A8:                   End If
  loc_1103D0A8:                 End If
  loc_1103D0A8:               End If
  loc_1103D0A8:             End If
  loc_1103D0A8:           End If
  loc_1103D0A8:         End If
  loc_1103D0A8:       End If
  loc_1103D0A8:     End If
  loc_1103D0A8:   End If
  loc_1103D0A8: End If
  loc_1103D0A8: ' Referenced from: 1103D019
End Sub

Private Sub Proc_10_12_1103D110
  Dim var_9C As Variant
  Dim var_A8 As frmXSCBToPz.Label3
  Dim var_A0 As Variant
  Dim var_260 As Label
  Dim var_38 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA47-EB96-11D2-AF06000021009B21}()
  loc_1103D264: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1103D26A: var_290 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1103D28A: Set var_9C = frmXSCBToPz.VFG
  loc_1103D2DA: If (CLng(var_9C.DispID_0007) < 2) Then
  loc_1103D308:   var_800C = = Global.Screen
  loc_1103D32A:   var_8010 = ecx
  loc_1103D332:   var_8010 = var_9C.UnkVCall_0000007Ch
  loc_1103D39F:   var_C8 = "提示信息"
  loc_1103D3A1:   var_150 = "没有可生成用友凭证的数据。"
  loc_1103D3B0: Else
  loc_1103D460:   var_260 = ("GetAccInfo".00000002h(, , fs:[00000000h], , "GL", var_16C, "dGLStartDate", var_174) = 1100AE28h)
  loc_1103D47A:   If var_260 = 0 Then GoTo loc_1103D5BB
  loc_1103D4A8:   var_801C = = Global.Screen
  loc_1103D4CA:   var_8020 = ecx
  loc_1103D4D2:   var_8020 = var_9C.UnkVCall_0000007Ch
  loc_1103D53F:   var_C8 = "提示信息"
  loc_1103D541:   var_150 = "总账系统尚未启用，不能进行凭证引入！"
  loc_1103D54B: End If
  loc_1103D57D: MsgBox(var_150, 64, var_C8, var_D8, var_E8)
  loc_1103D5AA: Exit Sub
  loc_1103D5B6: GoTo loc_110474A5
  loc_1103D5C5: var_8024 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[VouchNum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " & " CREATE TABLE VouchNum(iperiod tinyint NULL ,csign varchar(8) NULL ,ino_id int NULL,constraint index1 unique(iperiod,csign,ino_id))"
  loc_1103D5CB: var_B0 = var_8024
  loc_1103D62A: var_D8.00000001h(0, , , , "3缷M鋎?", var_AC, var_8024, var_B4)
  loc_1103D64A: On Error GoTo 0
  loc_1103D650: var_B0 = %ecx = %S_edx_S
  loc_1103D672: var_78 = "AS13"
  loc_1103D68A: var_78)
  loc_1103D6B4: If Not (var_78)) Then
  loc_1103D6E5:   If Global.Screen < 0 Then
  loc_1103D6F6:   End If
  loc_1103D700:   var_8030 = ecx
  loc_1103D70F:   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1103D722:   Else
  loc_1103D733:     Set var_9C = frmXSCBToPz.Label3
  loc_1103D739:     var_260 = var_9C
  loc_1103D747:     var_9C.Caption = "正在进行数据分析，请稍等..."
  loc_1103D7BE:     frmXSCBToPz.Pic1.DispID_80010007 = True
  loc_1103D7EE:     frmXSCBToPz.Pic1.DispID_FFFFFDDA
  loc_1103D80D:     var_8034 = .Proc_10_10_11035B90(var_24C)
  loc_1103D81B:     If var_24C = 2 Then
  loc_1103D864:       Set var_9C = frmXSCBToPz.Pic1
  loc_1103D86B:       var_9C.DispID_80010007 = %ecx = %S_edx_S
  loc_1103D90A:       MsgBox("数据源中没有合法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1103D947:       var_24C = %ecx = %S_edx_S
  loc_1103D96D:       "AS13")
  loc_1103D9AF:       var_B8 = Global.Screen
  loc_1103D9D1:       var_803C = ecx
  loc_1103D9E0:       If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1103D9F3:       Else
  loc_1103D9F5:         If var_803C = 1 Then
  loc_1103DA3E:           Set var_9C = frmXSCBToPz.Pic1
  loc_1103DA45:           var_9C.DispID_80010007 = %ecx = %S_edx_S
  loc_1103DAE4:           MsgBox("数据源中含有非法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1103DB21:           var_24C = %ecx = %S_edx_S
  loc_1103DB47:           "AS13")
  loc_1103DB89:           var_B8 = Global.Screen
  loc_1103DBAB:           var_8044 = ecx
  loc_1103DBBA:           If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1103DBCD:           Else
  loc_1103DBCF:             If var_8044 = 3 Then
  loc_1103DC18:               Set var_9C = frmXSCBToPz.Pic1
  loc_1103DC1F:               var_9C.DispID_80010007 = %ecx = %S_edx_S
  loc_1103DCBE:               MsgBox("数据源中指定的凭证号无效或重号，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1103DCFB:               var_24C = %ecx = %S_edx_S
  loc_1103DD21:               "AS13")
  loc_1103DD63:               var_B8 = Global.Screen
  loc_1103DD85:               var_804C = ecx
  loc_1103DD94:               If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1103DDA7:               Else
  loc_1103DDE9:                 var_C8 = "提示信息"
  loc_1103DE0F:                 var_B8 = "数据源中的数据已全部通过检查，是否开始引入？"
  loc_1103DE33:                 MsgBox(var_B8, 36, var_C8, var_D8, var_E8)
  loc_1103DE78:                 If (MsgBox(var_B8, 36, var_C8, var_D8, var_E8) = 7) Then
  loc_1103DEC1:                   Set var_9C = frmXSCBToPz.Pic1
  loc_1103DEC8:                   var_9C.DispID_80010007 = %ecx = %S_edx_S
  loc_1103DEEE:                   var_24C = %ecx = %S_edx_S
  loc_1103DF14:                   "AS13")
  loc_1103DF56:                   var_B8 = Global.Screen
  loc_1103DF78:                   var_8054 = ecx
  loc_1103DF87:                   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1103DF9A:                   Else
  loc_1103DF9B:                     On Error GoTo 0
  loc_1103DFB2:                     Set var_9C = frmXSCBToPz.Label3
  loc_1103DFB8:                     var_260 = var_9C
  loc_1103DFC6:                     var_9C.Caption = "正在写数据，请稍等..."
  loc_1103E011:                     frmXSCBToPz.Pic1.DispID_FFFFFDDA
  loc_1103E048:                     Set var_74 = CreateObject("UfDbKit.UfRecordset", 0)
  loc_1103E05F:                     var_150 = "SELECT TOP 1 * FROM GL_accvouch"
  loc_1103E0AE:                     var_B8 = "DataMdb".00000000h
  loc_1103E0D4:                     Set var_74 = var_C8
  loc_1103E116:                     call var_8060 = frmXSCBToPz.VFG(var_B8, frmXSCBToPz.VFG, 00000007h, 00000000h, frmXSCBToPz.VFG, global_1100C47C, 0000007Ch, global_1100C47C, 0000007Ch, global_1100C47C, 0000007Ch, global_1100C47C, 0000007Ch, global_1100C47C, 0000007Ch)
  loc_1103E16C:                     If var_24 <= CLng(var_8060)(-1) Then
  loc_1103E176:                       var_2A4 = var_24
  loc_1103E17C:                       var_150 = var_24
  loc_1103E207:                       call var_806C = frmXSCBToPz.VFG(var_B8, frmXSCBToPz.VFG, 00000082h, 00000002h, 3, var_174, 2, var_16C, 00000003h, var_154, var_24, var_14C)
  loc_1103E213:                       var_C0 = var_806C
  loc_1103E231:                       var_D8)
  loc_1103E289:                       var_70 = CByte("DateToPeriod".00000001h(8, var_D4))
  loc_1103E2C2:                       var_150 = var_2A4
  loc_1103E349:                       call var_8078 = frmXSCBToPz.VFG(var_B8, frmXSCBToPz.VFG, 00000082h, 00000002h, 3, var_174, 3, var_16C, 00000003h, var_154, var_2A4, var_14C)
  loc_1103E35A:                       var_58 = var_8078
  loc_1103E37E:                       var_150 = var_2A4
  loc_1103E409:                       call var_8080 = frmXSCBToPz.VFG(var_B8, frmXSCBToPz.VFG, 00000082h, 00000002h, 3, var_174, 0, var_16C, 00000003h, var_154, var_2A4, var_14C)
  loc_1103E41A:                       var_64 = var_8080
  loc_1103E43E:                       var_150 = var_2A4
  loc_1103E4C9:                       call var_8088 = frmXSCBToPz.VFG(var_B8, frmXSCBToPz.VFG, 00000082h, 00000002h, 3, var_174, 1, var_16C, 00000003h, var_154, var_2A4, var_14C)
  loc_1103E524:                       If (var_8088 = global_1100D76C) Then
  loc_1103E53B:                         Set var_A8 = frmXSCBToPz.Label3
  loc_1103E541:                         var_260 = var_A8
  loc_1103E651:                         var_80 = "正在处理：第[" & frmXSCBToPz.VFG.DispID_0082(var_2A4, 2) & " - "
  loc_1103E792:                         var_D8 = frmXSCBToPz.VFG.DispID_0082(var_2A4, 0)
  loc_1103E7E9:                         var_A8.Caption = var_80 & frmXSCBToPz.VFG.DispID_0082(var_2A4, 3) & " - " & var_D8 & "]号凭证"
  loc_1103E8A2:                         frmXSCBToPz.Pic1.DispID_FFFFFDDA
  loc_1103E8D4:                         var_3C = var_24
  loc_1103E8E8:                         Set var_9C = frmXSCBToPz.Chk
  loc_1103E8F0:                         var_260 = var_9C
  loc_1103E92C:                         var_24C = var_9C(0).Value
  loc_1103E959:                         var_270 = (var_24C = 1)
  loc_1103E97B:                         If (var_24C = 1) Then
  loc_1103E9AE:                           var_24C = CInt("cIYear".00000000h)
  loc_1103E9C3:                           var_24C, var_70)
  loc_1103E9D0:                           var_54 = var_24C, var_70)
  loc_1103E9E1:                         Else
  loc_1103E9ED:                           var_70, var_58)
  loc_1103E9FA:                           var_54 = var_70, var_58)
  loc_1103E9FD:                         End If
  loc_1103EA02:                         If var_54 > 0 Then
  loc_1103EA0A:                           On Error GoTo loc_11045892
  loc_1103EA43:                           "wksAlias".00000000h.00000000h(var_64)
  loc_1103EA68:                           var_1A0 = var_70
  loc_1103EB31:                           var_D8)
  loc_1103EBDD:                           var_80D0 = (var_58 = frmXSCBToPz.VFG.DispID_0082(var_24, 3))
  loc_1103EBEA:                           var_1F0 = var_80D0 + 1
  loc_1103ECA6:                           var_80D8 = (var_64 = frmXSCBToPz.VFG.DispID_0082(var_24, 0))
  loc_1103ECB3:                           var_240 = var_80D8 + 1
  loc_1103ED49:                           var_80E4 = (frmXSCBToPz.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_80D0 + 1 And var_80D8 + 1
  loc_1103EDD5:                           If CBool(var_80E4) Then
  loc_1103EE76:                             var_C0 = frmXSCBToPz.VFG.DispID_0082(var_24, 6)
  loc_1103EEB3:                             var_1A0 = var_38
  loc_1103EF21:                             "kmCodeToProperties".00000002h
  loc_1103EF41:                             Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_1103EF7E:                             var_74.AddNew
  loc_1103EF89:                             var_150 = "iPeriod"
  loc_1103EFF9:                             var_74.DispID_0000(var_70)
  loc_1103F000:                             var_150 = "ibook"
  loc_1103F071:                             var_74.DispID_0000(0)
  loc_1103F073:                             var_190 = "csign"
  loc_1103F180:                             var_74.DispID_0000(frmXSCBToPz.VFG.DispID_0082(var_24, 3))
  loc_1103F1A7:                             var_190 = "isignseq"
  loc_1103F2C7:                             var_74.DispID_0000(Proc_0_4_11026BD0(frmXSCBToPz.VFG.DispID_0082(var_24, 3), var_58))
  loc_1103F2F2:                             var_150 = "ino_id"
  loc_1103F364:                             var_74.DispID_0000(var_54)
  loc_1103F366:                             var_190 = "dbill_date"
  loc_1103F415:                             var_C0 = frmXSCBToPz.VFG.DispID_0082(var_24, 2)
  loc_1103F433:                             var_D8)
  loc_1103F490:                             var_74.DispID_0000(var_D8)
  loc_1103F4BE:                             var_190 = "idoc"
  loc_1103F4D6:                             var_150 = var_24
  loc_1103F5DF:                             var_74.DispID_0000(Val(frmXSCBToPz.VFG.DispID_0082(var_150, 4)))
  loc_1103F60A:                             var_160 = "ctext1"
  loc_1103F671:                             var_74.DispID_0000(var_150)
  loc_1103F678:                             var_160 = "ctext2"
  loc_1103F6DF:                             var_74.DispID_0000(var_150)
  loc_1103F6E1:                             var_190 = "cbill"
  loc_1103F6F9:                             var_150 = var_24
  loc_1103F7F2:                             var_74.DispID_0000(frmXSCBToPz.VFG.DispID_0082(var_150, 12))
  loc_1103F81E:                             var_160 = "cbook"
  loc_1103F883:                             var_74.DispID_0000(var_150)
  loc_1103F88A:                             var_160 = "ccheck"
  loc_1103F8F1:                             var_74.DispID_0000(var_150)
  loc_1103F8F8:                             var_160 = "ccashier"
  loc_1103F95F:                             var_74.DispID_0000(var_150)
  loc_1103F966:                             var_160 = "iflag"
  loc_1103F9CD:                             var_74.DispID_0000(var_150)
  loc_1103F9D4:                             var_160 = "coutaccset"
  loc_1103FA3B:                             var_74.DispID_0000(var_150)
  loc_1103FA42:                             var_160 = "ioutyear"
  loc_1103FAA9:                             var_74.DispID_0000(var_150)
  loc_1103FAB0:                             var_160 = "coutsysver"
  loc_1103FB17:                             var_74.DispID_0000(var_150)
  loc_1103FB1E:                             var_160 = "coutsysname"
  loc_1103FB85:                             var_74.DispID_0000(var_150)
  loc_1103FB8C:                             var_170 = "ioutperiod"
  loc_1103FC25:                             var_74.DispID_0000(var_74.DispID_0000("iPeriod"))
  loc_1103FC36:                             var_170 = "doutbilldate"
  loc_1103FCF5:                             var_74.DispID_0000(CStr(var_74.DispID_0000("dbill_date")))
  loc_1103FD12:                             var_150 = "iYear"
  loc_1103FD80:                             var_74.DispID_0000("cIYear".00000000h(, var_14C, "iYear", var_154))
  loc_1103FDCB:                             var_150 = var_70
  loc_1103FE7E:                             var_74.DispID_0000("cIYear".00000000h(, var_16C, "iYPeriod", var_174) & Format(var_70, "00"))
  loc_1103FEAC:                             var_160 = "coutsign"
  loc_1103FF13:                             var_74.DispID_0000(var_150)
  loc_1103FF1A:                             var_160 = "coutno_id"
  loc_1103FF81:                             var_74.DispID_0000(var_150)
  loc_1103FF88:                             var_150 = "bvouchedit"
  loc_1103FFF9:                             var_74.DispID_0000(FFFFFFFFh)
  loc_11040000:                             var_150 = "bvouchaddordele"
  loc_11040071:                             var_74.DispID_0000(FFFFFFFFh)
  loc_11040078:                             var_150 = "bvouchmoneyhold"
  loc_110400E9:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110400F0:                             var_150 = "bvalueedit"
  loc_11040161:                             var_74.DispID_0000(FFFFFFFFh)
  loc_11040168:                             var_150 = "bcodeedit"
  loc_110401D9:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110401E0:                             var_150 = "bPCSedit"
  loc_11040251:                             var_74.DispID_0000(FFFFFFFFh)
  loc_11040258:                             var_150 = "bDeptedit"
  loc_110402C9:                             var_74.DispID_0000(FFFFFFFFh)
  loc_110402D0:                             var_150 = "bItemedit"
  loc_11040341:                             var_74.DispID_0000(FFFFFFFFh)
  loc_11040348:                             var_150 = "inid"
  loc_110403BA:                             var_74.DispID_0000(1)
  loc_110403BC:                             var_190 = "cdigest"
  loc_110404CD:                             var_74.DispID_0000(frmXSCBToPz.VFG.DispID_0082(var_24, 5))
  loc_110404F4:                             var_190 = "cCode"
  loc_11040603:                             var_74.DispID_0000(frmXSCBToPz.VFG.DispID_0082(var_24, 6))
  loc_110406AB:                             var_7C = var_38.UnkVCall_0000006Ch
  loc_110406F6:                             var_8120 = (var_38.UnkVCall_0000006Ch = global_1100AE28)
  loc_11040703:                             var_160 = var_8120 + 1
  loc_1104078E:                             var_74.DispID_0000(IIf(var_8120 + 1, vbNull, 0))
  loc_11040873:                             var_1B0 = "md"
  loc_110408BC:                             var_BC = var_258
  loc_11040943:                             var_74.DispID_0000(Format(Val(frmXSCBToPz.VFG.DispID_0082(var_24, 7)), "#.00"))
  loc_11040A34:                             var_1B0 = "mc"
  loc_11040A7D:                             var_BC = var_258
  loc_11040B04:                             var_74.DispID_0000(Format(Val(frmXSCBToPz.VFG.DispID_0082(var_24, 8)), "#.00"))
  loc_11040BC8:                             If (var_74.DispID_0000("md") <> 0) Then
  loc_11040C3D:                               If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_11040C48:                                 var_150 = "md_f"
  loc_11040CB9:                                 var_74.DispID_0000(0)
  loc_11040CC3:                               Else
  loc_11040D76:                                 var_1B0 = "md_f"
  loc_11040DBF:                                 var_BC = var_258
  loc_11040E46:                                 var_74.DispID_0000(Format(Val(frmXSCBToPz.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_11040E87:                               End If
  loc_11040EF9:                               If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_11040F04:                                 var_150 = "nd_s"
  loc_11040F75:                                 var_74.DispID_0000(0)
  loc_11040F7F:                               Else
  loc_11040F8E:                               Else
  loc_11040FFD:                                 If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_11041008:                                   var_150 = "mc_f"
  loc_11041079:                                   var_74.DispID_0000(0)
  loc_11041083:                                 Else
  loc_11041136:                                   var_1B0 = "mc_f"
  loc_1104117F:                                   var_BC = var_258
  loc_11041206:                                   var_74.DispID_0000(Format(Val(frmXSCBToPz.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_11041247:                                 End If
  loc_110412B9:                                 If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_110412C0:                                   GoTo loc_11040F04
  loc_110412C5:                                 End If
  loc_110412CF:                               End If
  loc_110413E9:                               var_74.DispID_0000(Val(frmXSCBToPz.VFG.DispID_0082(var_24, 9)))
  loc_1104140F:                             End If
  loc_11041481:                             If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_1104148C:                               var_150 = "nfrat"
  loc_110414FD:                               var_74.DispID_0000(0)
  loc_11041507:                             Else
  loc_1104162B:                               var_74.DispID_0000(Val(frmXSCBToPz.VFG.DispID_0082(var_24, 11)))
  loc_11041651:                             End If
  loc_110416A6:                             If var_38.UnkVCall_0000010Ch Then
  loc_1104173D:                               var_1F0 = "csettle"
  loc_11041824:                               var_8174 = (frmXSCBToPz.VFG.DispID_0082(var_24, 13) = global_1100AE28)
  loc_11041831:                               var_1E0 = var_8174 + 1
  loc_110418BC:                               var_74.DispID_0000(IIf(var_8174 + 1, vbNull, frmXSCBToPz.VFG.DispID_0082(var_24, 13)))
  loc_11041915:                             End If
  loc_1104193E:                             var_24C = var_38.UnkVCall_0000010Ch
  loc_1104198B:                             var_250 = var_38.UnkVCall_00000094h
  loc_11041A2A:                             If (var_38.UnkVCall_0000009Ch = 0) = 0 Then
  loc_11041AC1:                               var_1F0 = "cn_id"
  loc_11041B70:                               var_E0 = frmXSCBToPz.VFG.DispID_0082(var_24, 14)
  loc_11041BA8:                               var_818C = (frmXSCBToPz.VFG.DispID_0082(var_24, 14) = global_1100AE28)
  loc_11041BB5:                               var_1E0 = var_818C + 1
  loc_11041C40:                               var_74.DispID_0000(IIf(var_818C + 1, vbNull, var_E0))
  loc_11041D27:                               var_1F0 = "dt_date"
  loc_11041DD6:                               var_D0 = frmXSCBToPz.VFG.DispID_0082(var_24, 15)
  loc_11041DF4:                               var_E0)
  loc_11041E21:                               var_8198 = (frmXSCBToPz.VFG.DispID_0082(var_24, 15) = global_1100AE28)
  loc_11041E2E:                               var_1E0 = var_8198 + 1
  loc_11041EB9:                               var_74.DispID_0000(IIf(var_8198 + 1, vbNull, var_E0))
  loc_11041FA7:                               var_1F0 = "cname"
  loc_110420A7:                               var_81A4 = (frmXSCBToPz.VFG.DispID_0082(var_24, &H14) = global_1100AE28)
  loc_110420B4:                               var_1D0 = var_81A4 + 1
  loc_1104213F:                               var_74.DispID_0000(IIf(var_81A4 + 1, "-", frmXSCBToPz.VFG.DispID_0082(var_24, &H14)))
  loc_11042198:                             End If
  loc_1104220E:                             var_250 = var_38.UnkVCall_0000008Ch
  loc_1104224C:                             If (var_38.UnkVCall_000000A4h = 0) = 0 Then
  loc_11042256:                               var_150 = var_24
  loc_110422E3:                               var_1F0 = "cdept_id"
  loc_110423CA:                               var_81B8 = (frmXSCBToPz.VFG.DispID_0082(var_150, 16) = global_1100AE28)
  loc_110423D7:                               var_1E0 = var_81B8 + 1
  loc_11042462:                               var_74.DispID_0000(IIf(var_81B8 + 1, vbNull, frmXSCBToPz.VFG.DispID_0082(var_24, 16)))
  loc_110424BD:                             Else
  loc_110424C2:                               var_160 = "cdept_id"
  loc_11042529:                               var_74.DispID_0000(var_150)
  loc_1104252E:                             End If
  loc_11042583:                             If var_38.UnkVCall_0000008Ch Then
  loc_1104258D:                               var_150 = var_24
  loc_1104261A:                               var_1F0 = "cperson_id"
  loc_11042701:                               var_81C8 = (frmXSCBToPz.VFG.DispID_0082(var_150, &H11) = global_1100AE28)
  loc_1104270E:                               var_1E0 = var_81C8 + 1
  loc_11042799:                               var_74.DispID_0000(IIf(var_81C8 + 1, vbNull, frmXSCBToPz.VFG.DispID_0082(var_24, &H11)))
  loc_110427F4:                             Else
  loc_110427F9:                               var_160 = "cperson_id"
  loc_11042860:                               var_74.DispID_0000(var_150)
  loc_11042865:                             End If
  loc_110428BA:                             If var_38.UnkVCall_00000094h Then
  loc_110428C4:                               var_150 = var_24
  loc_11042951:                               var_1F0 = "ccus_id"
  loc_11042A38:                               var_81D8 = (frmXSCBToPz.VFG.DispID_0082(var_150, &H12) = global_1100AE28)
  loc_11042A45:                               var_1E0 = var_81D8 + 1
  loc_11042AD0:                               var_74.DispID_0000(IIf(var_81D8 + 1, vbNull, frmXSCBToPz.VFG.DispID_0082(var_24, &H12)))
  loc_11042B2B:                             Else
  loc_11042B30:                               var_160 = "ccus_id"
  loc_11042B97:                               var_74.DispID_0000(var_150)
  loc_11042B9C:                             End If
  loc_11042BF1:                             If var_38.UnkVCall_0000009Ch Then
  loc_11042BFB:                               var_150 = var_24
  loc_11042C88:                               var_1F0 = "csup_id"
  loc_11042D6F:                               var_81E8 = (frmXSCBToPz.VFG.DispID_0082(var_150, &H13) = global_1100AE28)
  loc_11042D7C:                               var_1E0 = var_81E8 + 1
  loc_11042E07:                               var_74.DispID_0000(IIf(var_81E8 + 1, vbNull, frmXSCBToPz.VFG.DispID_0082(var_24, &H13)))
  loc_11042E62:                             Else
  loc_11042E67:                               var_160 = "csup_id"
  loc_11042ECE:                               var_74.DispID_0000(var_150)
  loc_11042ED3:                             End If
  loc_11042F4C:                             If (var_38.UnkVCall_000000ACh = global_1100AE28) Then
  loc_11042F56:                               var_150 = var_24
  loc_11042FE3:                               var_1F0 = "citem_id"
  loc_110430CA:                               var_81FC = (frmXSCBToPz.VFG.DispID_0082(var_150, &H15) = global_1100AE28)
  loc_110430D7:                               var_1E0 = var_81FC + 1
  loc_11043162:                               var_74.DispID_0000(IIf(var_81FC + 1, vbNull, frmXSCBToPz.VFG.DispID_0082(var_24, &H15)))
  loc_1104323F:                               var_7C = var_38.UnkVCall_000000ACh
  loc_11043290:                               var_8208 = (var_38.UnkVCall_000000ACh = global_1100AE28)
  loc_1104329D:                               var_160 = var_8208 + 1
  loc_11043328:                               var_74.DispID_0000(IIf(var_8208 + 1, vbNull, 0))
  loc_11043362:                             Else
  loc_11043367:                               var_160 = "citem_id"
  loc_110433CE:                               var_74.DispID_0000(var_150)
  loc_110433D5:                               var_160 = "citem_class"
  loc_1104343C:                               var_74.DispID_0000(var_150)
  loc_11043441:                             End If
  loc_11043446:                             var_160 = "ccode_equal"
  loc_110434AD:                             var_74.DispID_0000(var_150)
  loc_110434B4:                             var_160 = "iflagbank"
  loc_1104351B:                             var_74.DispID_0000(var_150)
  loc_11043522:                             var_160 = "iflagperson"
  loc_11043589:                             var_74.DispID_0000(var_150)
  loc_110435E3:                             If var_38.UnkVCall_0000017Ch Then
  loc_110435ED:                               var_150 = var_24
  loc_1104367A:                               var_1F0 = "cDefine1"
  loc_1104377A:                               var_8218 = (frmXSCBToPz.VFG.DispID_0082(var_150, &H16) = global_1100AE28)
  loc_11043787:                               var_1D0 = var_8218 + 1
  loc_11043812:                               var_74.DispID_0000(IIf(var_8218 + 1, "-", frmXSCBToPz.VFG.DispID_0082(var_24, &H16)))
  loc_1104386D:                             Else
  loc_11043872:                               var_160 = "cDefine1"
  loc_110438D9:                               var_74.DispID_0000(var_150)
  loc_110438DE:                             End If
  loc_11043933:                             If var_38.UnkVCall_00000184h Then
  loc_1104393D:                               var_150 = var_24
  loc_110439CA:                               var_1F0 = "cDefine2"
  loc_11043ACA:                               var_8228 = (frmXSCBToPz.VFG.DispID_0082(var_150, &H1B) = global_1100AE28)
  loc_11043AD7:                               var_1D0 = var_8228 + 1
  loc_11043B62:                               var_74.DispID_0000(IIf(var_8228 + 1, "-", frmXSCBToPz.VFG.DispID_0082(var_24, &H1B)))
  loc_11043BBD:                             Else
  loc_11043BC2:                               var_160 = "cDefine2"
  loc_11043C29:                               var_74.DispID_0000(var_150)
  loc_11043C2E:                             End If
  loc_11043C83:                             If var_38.UnkVCall_0000018Ch Then
  loc_11043C8D:                               var_150 = var_24
  loc_11043D1A:                               var_1F0 = "cDefine3"
  loc_11043E1A:                               var_8238 = (frmXSCBToPz.VFG.DispID_0082(var_150, &H1A) = global_1100AE28)
  loc_11043E27:                               var_1D0 = var_8238 + 1
  loc_11043EB2:                               var_74.DispID_0000(IIf(var_8238 + 1, "-", frmXSCBToPz.VFG.DispID_0082(var_24, &H1A)))
  loc_11043F0D:                             Else
  loc_11043F12:                               var_160 = "cDefine3"
  loc_11043F79:                               var_74.DispID_0000(var_150)
  loc_11043F7E:                             End If
  loc_11043FD3:                             If var_38.UnkVCall_000001B4h Then
  loc_1104406A:                               var_1F0 = "cDefine8"
  loc_1104416A:                               var_8248 = (frmXSCBToPz.VFG.DispID_0082(var_24, &H1C) = global_1100AE28)
  loc_11044177:                               var_1D0 = var_8248 + 1
  loc_11044202:                               var_74.DispID_0000(IIf(var_8248 + 1, "-", frmXSCBToPz.VFG.DispID_0082(var_24, &H1C)))
  loc_1104425B:                             End If
  loc_110442B0:                             If var_38.UnkVCall_000001C4h Then
  loc_11044347:                               var_1F0 = "cDefine10"
  loc_11044447:                               var_8258 = (frmXSCBToPz.VFG.DispID_0082(var_24, &H18) = global_1100AE28)
  loc_11044454:                               var_1D0 = var_8258 + 1
  loc_110444DF:                               var_74.DispID_0000(IIf(var_8258 + 1, "-", frmXSCBToPz.VFG.DispID_0082(var_24, &H18)))
  loc_11044538:                             End If
  loc_1104458D:                             If var_38.UnkVCall_000001DCh Then
  loc_11044597:                               var_150 = var_24
  loc_11044624:                               var_1F0 = "cDefine12"
  loc_11044724:                               var_8268 = (frmXSCBToPz.VFG.DispID_0082(var_150, &H17) = global_1100AE28)
  loc_11044731:                               var_1D0 = var_8268 + 1
  loc_110447BC:                               var_74.DispID_0000(IIf(var_8268 + 1, "-", frmXSCBToPz.VFG.DispID_0082(var_24, &H17)))
  loc_11044817:                             Else
  loc_1104481C:                               var_160 = "cDefine12"
  loc_11044883:                               var_74.DispID_0000(var_150)
  loc_11044888:                             End If
  loc_110448DD:                             If var_38.UnkVCall_000001E4h Then
  loc_110448E7:                               var_150 = var_24
  loc_11044974:                               var_1F0 = "cDefine13"
  loc_11044A74:                               var_8278 = (frmXSCBToPz.VFG.DispID_0082(var_150, &H19) = global_1100AE28)
  loc_11044A81:                               var_1D0 = var_8278 + 1
  loc_11044B0C:                               var_74.DispID_0000(IIf(var_8278 + 1, "-", frmXSCBToPz.VFG.DispID_0082(var_24, &H19)))
  loc_11044B67:                             Else
  loc_11044B6C:                               var_160 = "cDefine13"
  loc_11044BD3:                               var_74.DispID_0000(var_150)
  loc_11044BD8:                             End If
  loc_11044BE3:                             var_74.Update
  loc_11044BFA:                             var_24 = var_24(1)
  loc_11044C0B:                             var_68 = var_68(1)
  loc_11044C40:                             var_827C = CLng(frmXSCBToPz.VFG.DispID_0007)
  loc_11044C5C:                             var_260 = (var_24(1) > 0)
  loc_11044C83:                             If var_260 = 0 Then GoTo loc_1103EA65
  loc_11044C89:                           End If
  loc_11044CBC:                           "wksAlias".00000000h.00000000h
  loc_11044CE9:                           Set var_9C = frmXSCBToPz.Chk
  loc_11044CEF:                           var_260 = var_9C
  loc_11044D01:                           Set var_A0 = var_9C(0)
  loc_11044D25:                           var_268 = var_A0
  loc_11044D8F:                           If (var_A0.Value = 1) Then
  loc_11044D9D:                             var_70, var_58)
  loc_11044DA2:                           End If
  loc_11044DA4:                           On Error GoTo 0
  loc_11044DDB:                           var_250 = CInt("cIYear".00000000h)
  loc_11044E05:                           var_24C, var_250, var_70, var_58)
  loc_11044E0F:                           var_5C = var_24C, var_250, var_70, var_58)
  loc_11044E52:                           var_250 = CInt("cIYear".00000000h)
  loc_11044E86:                           var_48 = r_250, var_70, var_58) var_250, var_70, var_58)
  loc_11044E98:                           var_150 = "select * from GL_accvouch where ibook=0 and iYear="
  loc_11044EC0:                           var_170 = var_70
  loc_11044EE4:                           var_828C = Proc_0_4_11026BD0(var_58, var_54, var_54)
  loc_11044EE9:                           var_190 = var_828C
  loc_11044F11:                           var_1B0 = var_54
  loc_11044F6A:                           var_D8 = 1 & "cIYear".00000000h(, 1, 1) & " and iperiod="
  loc_11044FD3:                           var_128 = var_D8 & var_70 & " and isignseq=" & var_828C & " and ino_id=" & var_54
  loc_1104503C:                           Set var_74 = "DataMdb".00000000h.00000001h
  loc_110450DB:                           If CBool(Not(var_74.EOF)) Then
  loc_11045133:                             If CBool(Not(var_74.EOF)) Then
  loc_1104513C:                               var_170 = var_70
  loc_11045151:                               var_150 = "iPeriod"
  loc_11045175:                               var_180 = "csign"
  loc_11045189:                               var_1D0 = var_54
  loc_1104519A:                               var_1B0 = "ino_id"
  loc_110452E5:                               If CBool((var_70 = var_14C) And (var_58 = var_D8) And (var_54 = var_1AC)) Then
  loc_110452F0:                                 var_150 = "mc"
  loc_1104536E:                                 var_180 = "ccode_equal"
  loc_11045382:                                 If (var_14C <> 0) Then
  loc_110453AE:                                   var_82B8 = (var_5C = global_1100AE28)
  loc_110453BB:                                   var_160 = var_82B8 + 1
  loc_110453E8:                                   var_C8 = IIf(var_82B8 + 1, vbNull, var_5C)
  loc_11045462:                                 Else
  loc_11045488:                                   var_82BC = (var_48 = global_1100AE28)
  loc_11045495:                                   var_160 = var_82BC + 1
  loc_110454C2:                                   var_C8 = IIf(var_82BC + 1, vbNull, var_48)
  loc_11045537:                                 End If
  loc_1104554D:                                 var_74.Update
  loc_11045597:                                 var_180 = var_38
  loc_110455DE:                                 var_B8 = var_74.DispID_0000("cCode")
  loc_11045637:                                 "kmCodeToProperties".00000002h
  loc_11045657:                                 Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_1104567A:                                 var_150 = "citem_class"
  loc_110456DD:                                 If IsNull(var_74.DispID_0000(var_150)) Then
  loc_110456F2:                                 Else
  loc_11045733:                                   var_180 = var_28
  loc_1104577A:                                   var_B8 = var_74.DispID_0000(var_150)
  loc_110457D3:                                   "XmClassIDToProperties".00000002h
  loc_11045837:                                   var_78 = {3302AA47-EB96-11D2-AF06000021009B21}().UnkVCall_0000002Ch
  loc_11045868:                                 End If
  loc_11045876:                                 var_68 = var_68(1)
  loc_11045884:                                 var_74.MoveNext
  loc_1104588D:                                 GoTo loc_110450E8
  loc_110458C5:                                 "wksAlias".00000000h.00000000h
  loc_110458DD:                                 var_30 = var_3C
  loc_110458F2:                                 var_1A0 = var_70
  loc_110459BB:                                 var_D8)
  loc_11045A67:                                 var_82DC = (var_58 = frmXSCBToPz.VFG.DispID_0082(var_30, 3))
  loc_11045A74:                                 var_1F0 = var_82DC + 1
  loc_11045B30:                                 var_82E4 = (var_64 = frmXSCBToPz.VFG.DispID_0082(var_30, 0))
  loc_11045B3D:                                 var_240 = var_82E4 + 1
  loc_11045BD3:                                 var_82F0 = (frmXSCBToPz.VFG.DispID_0082(var_30, 2) = "DateToPeriod".00000001h) And var_82DC + 1 And var_82E4 + 1
  loc_11045C5F:                                 If CBool(var_82F0) Then
  loc_11045C69:                                   var_150 = var_30
  loc_11045D29:                                   frmXSCBToPz.VFG.DispID_0082(1, "-")
  loc_11045EAD:                                   frmXSCBToPz.VFG.DispID_009E(var_30, 1, var_30, 1, &HFF)
  loc_11045EC2:                                   var_150 = var_30
  loc_11045F82:                                   frmXSCBToPz.VFG.DispID_0082(&H16, "数据提交错或该数据已经被导入----未引入")
  loc_11045FA1:                                   var_30 = var_30(1)
  loc_11045FCD:                                   var_82F8 = CLng(frmXSCBToPz.VFG.DispID_0007)
  loc_11045FE9:                                   var_260 = (var_30 > 0)
  loc_11046010:                                   If var_260 = 0 Then GoTo loc_110458EF
  loc_11046016:                                 End If
  loc_11046019:                                 var_24 = var_30
  loc_1104602D:                                 Set var_9C = frmXSCBToPz.Chk
  loc_11046033:                                 var_260 = var_9C
  loc_11046045:                                 Set var_A0 = var_9C(0)
  loc_11046069:                                 var_268 = var_A0
  loc_110460D3:                                 If (var_A0.Value = 1) Then
  loc_110461CF:                                   "unLockVouch".00000004h(var_180, var_BC, var_C4, 0, var_74, var_70, var_58, var_16C, var_54, &H4002, var_184)
  loc_110461D8:                                 End If
  loc_110461DD:                                 var_150 = "VouchNum"
  loc_11046252:                                 Set var_34 = "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C4, 0, var_14C, "VouchNum", var_154)
  loc_11046273:                                 var_150 = "delete  from vouchnum"
  loc_110462D1:                                 "DataMdb".00000000h.00000001h(1, 1, var_180, var_BC, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_11046332:                                 frmXSCBToPz.Pic1.DispID_80010007 = var_150
  loc_11046346:                                 var_8304 = Resume(0)
  loc_1104634C:                               End If
  loc_1104634C:                             End If
  loc_1104634C:                           End If
  loc_1104636A:                           var_24 = var_278+(var_24 - 1)
  loc_1104636D:                           GoTo loc_1103E161
  loc_11046372:                         End If
  loc_1104637B:                         var_1A0 = var_70
  loc_11046444:                         var_D8)
  loc_110464F0:                         var_8310 = (var_58 = frmXSCBToPz.VFG.DispID_0082(var_24, 3))
  loc_110464FD:                         var_1F0 = var_8310 + 1
  loc_110465B9:                         var_8318 = (var_64 = frmXSCBToPz.VFG.DispID_0082(var_24, 0))
  loc_110465C6:                         var_240 = var_8318 + 1
  loc_1104665C:                         var_8324 = (frmXSCBToPz.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_8310 + 1 And var_8318 + 1
  loc_11046669:                         var_260 = CBool(var_8324)
  loc_110466E8:                         If var_260 = 0 Then GoTo loc_1104634C
  loc_110466FF:                         Set var_9C = frmXSCBToPz.Chk
  loc_11046705:                         var_260 = var_9C
  loc_11046717:                         Set var_A0 = var_9C(0)
  loc_1104673B:                         var_268 = var_A0
  loc_1104677E:                         var_270 = (var_A0.Value = 1)
  loc_110467A9:                         var_150 = var_24
  loc_110467CA:                         var_190 = "网络共享冲突----未引入"
  loc_110467D4:                         If var_270 = 0 Then
  loc_110467D6:                           var_190 = "指定的凭证号无效或重号----未引入"
  loc_110467E0:                         End If
  loc_11046875:                         frmXSCBToPz.VFG.DispID_0082(var_170, var_190)
  loc_11046894:                         var_24 = var_24(1)
  loc_1104689A:                         var_2A4 = var_24(1)
  loc_110468C9:                         var_832C = CLng(frmXSCBToPz.VFG.DispID_0007)
  loc_110468E5:                         var_260 = (var_2A4 > 0)
  loc_1104690C:                         If var_260 = 0 Then GoTo loc_11046378
  loc_11046912:                         GoTo loc_1104634C
  loc_11046917:                       End If
  loc_1104691A:                       var_1A0 = var_70
  loc_110469E5:                       var_D8)
  loc_11046A93:                       var_8338 = (var_58 = frmXSCBToPz.VFG.DispID_0082(var_2A4, 3))
  loc_11046AA0:                       var_1F0 = var_8338 + 1
  loc_11046B5E:                       var_8340 = (var_64 = frmXSCBToPz.VFG.DispID_0082(var_2A4, 0))
  loc_11046B6B:                       var_240 = var_8340 + 1
  loc_11046C01:                       var_834C = (frmXSCBToPz.VFG.DispID_0082(var_2A4, 2) = "DateToPeriod".00000001h) And var_8338 + 1 And var_8340 + 1
  loc_11046C0E:                       var_260 = CBool(var_834C)
  loc_11046C8D:                       If var_260 = 0 Then GoTo loc_1104634C
  loc_11046D84:                       If (frmXSCBToPz.VFG.DispID_0082(var_2A4, &H16) = global_1100AE28) + 1 Then
  loc_11046D8A:                         var_150 = var_2A4
  loc_11046E43:                         Set var_9C = frmXSCBToPz.VFG
  loc_11046E4A:                         var_9C.DispID_0082(&H16, "凭证借贷不平衡或某分录有错误----未引入")
  loc_11046E5B:                         GoTo loc_11046917
  loc_11046E60:                       End If
  loc_11046F2A:                       var_C0 = frmXSCBToPz.VFG.DispID_0082(frmXSCBToPz.VFG, &H16) & "----未引入"
  loc_11046FCB:                       frmXSCBToPz.VFG.DispID_0082(&H16, var_C0)
  loc_11047008:                       GoTo loc_11046917
  loc_1104700D:                     End If
  loc_1104705F:                     frmXSCBToPz.Pic1.DispID_80010007 = var_150
  loc_11047094:                     var_160 = "提示信息"
  loc_1104709E:                     If var_2C Then
  loc_11047102:                       MsgBox("数据引入已完成，数据已生成用友凭证。", 64, var_160, var_D8, var_E8)
  loc_11047173:                       Set var_9C = frmXSCBToPz.VFG
  loc_1104718D:                     Else
  loc_110471E8:                       MsgBox("数据没有被引入，原因请查看最后一列中的说明。", 64, var_160, var_D8, var_E8)
  loc_11047217:                     End If
  loc_1104721C:                     var_150 = "VouchNum"
  loc_11047293:                     Set var_34 = "DataMdb".00000000h.00000001h(var_180, var_BC, var_C0, var_C4, var_C8, var_14C, "VouchNum", var_154)
  loc_110472B0:                     var_150 = "delete  from vouchnum"
  loc_11047304:                     "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C0, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_11047359:                     "AS13")
  loc_11047392:                     var_B8 = Global.Screen
  loc_110473B4:                     var_8370 = ecx
  loc_110473C3:                     If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110473CD:                     End If
  loc_110473CD:                   End If
  loc_110473CD:                 End If
  loc_110473CD:               End If
  loc_110473CD:             End If
  loc_110473CE:             var_8370 = CheckObj(var_9C, global_1100C47C, 124)
  loc_110473D4:           End If
  loc_110473D4:         End If
  loc_110473D4:       End If
  loc_110473D4:     End If
  loc_110473D4:   End If
  loc_110473D4: End If
  loc_110473E0: Exit Sub
  loc_110473EC: GoTo loc_110474A5
  loc_110474A4: Exit Sub
  loc_110474A5: ' Referenced from: 1103D5B6
  loc_110474A5: ' Referenced from: 110473EC
End Sub

Private Sub Proc_10_13_11047510
  Dim var_58 As Variant
  Dim var_5C As Variant
  Dim var_64 As frmXSCBToPz.Label3
  Dim var_1CC As Label
  loc_110475F7: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11047600: var_1EC = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1104761D: Set var_58 = frmXSCBToPz.Chk
  loc_11047627: var_1CC = var_58
  loc_1104762D: Set var_5C = var_58(0)
  loc_11047658: var_1D4 = var_5C
  loc_1104769B: var_1DC = (var_5C.Value = 1)
  loc_110476B1: If var_1DC = 0 Then
  loc_11047716:   If var_14 <= CLng(frmXSCBToPz.VFG.DispID_0007)(-1) Then
  loc_1104778B:     var_7C = frmXSCBToPz.VFG.DispID_0082(var_14, 2)
  loc_110477A6:     var_94)
  loc_11047801:     var_30 = CByte("DateToPeriod".00000001h)
  loc_110478AB:     var_24 = frmXSCBToPz.VFG.DispID_0082(var_14, 3)
  loc_1104795B:     Set var_64 = frmXSCBToPz.Label3
  loc_11047985:     var_1CC = var_64
  loc_11047B3B:     var_94 = frmXSCBToPz.VFG.DispID_0082(var_14, frmXSCBToPz.VFG)
  loc_11047B6E:     var_8038 = "正在处理：第[" & frmXSCBToPz.VFG.DispID_0082(var_14, 2) & " - " & frmXSCBToPz.VFG.DispID_0082(var_14, 3) & " - " & var_94 & "]号凭证是否重号"
  loc_11047B8D:     var_64.Caption = var_8038
  loc_11047C22:     If r_24) var_24) <= 0 Then
  loc_11047C34:       var_13C = var_30
  loc_11047CCB:       var_94)
  loc_11047D64:       var_8048 = (var_24 = frmXSCBToPz.VFG.DispID_0082(var_14, 3))
  loc_11047D91:       var_17C = var_8048 + 1
  loc_11047E08:       var_8050 = (frmXSCBToPz.VFG.DispID_0082(var_14, frmXSCBToPz.VFG) = frmXSCBToPz.VFG.DispID_0082(var_14, ""))
  loc_11047E2F:       var_1BC = var_8050 + 1
  loc_11047F2A:       If CBool((frmXSCBToPz.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_8048 + 1 And var_8050 + 1) Then
  loc_11047FB2:         frmXSCBToPz.VFG.DispID_0082(1, 285267820)
  loc_110480E6:         frmXSCBToPz.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_1104817A:         frmXSCBToPz.VFG.DispID_0082(29, "指定的凭证号无效或重号")
  loc_110481C5:         var_8064 = CLng(frmXSCBToPz.VFG.DispID_0007)
  loc_110481E3:         var_1CC = (var_14(1) > 0)
  loc_11048200:         If var_1CC = 0 Then GoTo loc_11047C2E
  loc_11048206:       End If
  loc_11048214:     Else
  loc_1104821D:     End If
  loc_1104822A:     var_14 = 1+var_14
  loc_1104822D:     GoTo loc_11047710
  loc_11048232:   End If
  loc_11048232: End If
  loc_11048237: GoTo loc_110482C8
  loc_110482C7: Exit Sub
  loc_110482C8: ' Referenced from: 11048237
End Sub

Private Sub Proc_10_14_110489D0
  Dim var_30 As Variant
  Dim var_34 As Variant
  loc_11048A4B: var_8C = "Excel 97/2000 (*.xls)|*.xls"
  loc_11048A94: frmXSCBToPz.dlg.Filter = var_90
  loc_11048AE8: frmXSCBToPz.dlg.FileName = var_90
  loc_11048B0B: frmXSCBToPz.dlg.ShowSave
  loc_11048B2E: Set var_30 = frmXSCBToPz.dlg
  loc_11048B7E: If Len(var_30) + 1 = 0 Then
  loc_11048B86:   On Error GoTo loc_11048F9F
  loc_11048BD0:   frmXSCBToPz.dlg.FileName = var_D8
  loc_11048BE5:   var_24 = var_44
  loc_11048BF5:   var_8014 = Scripting.FileSystemObject.FileExists
  loc_11048C32:   If var_D8 Then
  loc_11048C71:     var_54 = "询问"
  loc_11048C90:     var_44 = "相同名称的文件已经存在，是否替换原来的文件？"
  loc_11048CA4:     MsgBox(var_44, 33, var_54, var_64, var_74)
  loc_11048CD2:     If MsgBox(var_44, 33, var_54, var_64, var_74) - 1 + 1 = 0 Then GoTo loc_110490D3
  loc_11048D0B:     Set var_30 = frmXSCBToPz.dlg
  loc_11048D2B:     var_24 = var_30
  loc_11048D3B:     var_8020 = Scripting.FileSystemObject.DeleteFile(var_24, False)
  loc_11048D6E:   End If
  loc_11048D7D:   Scripting.FileSystemObject.MousePointer = CInt(11)
  loc_11048DAC:   Set var_30 = frmXSCBToPz.dlg
  loc_11048DB7:   var_30.FileName = var_24
  loc_11048E5A:   Set var_34 = frmXSCBToPz.VFG
  loc_11048E61:   var_34.DispID_0097(var_30, 6, 7)
  loc_11048E99:   var_34.MousePointer = ecx
  loc_11048ED3:   frmXSCBToPz.dlg.FileName = var_24
  loc_11048F55:   MsgBox("文件已输出到" & var_44, 65, "提示", 10, 10)
  loc_11048F8F:   Exit Sub
  loc_11048F9A:   GoTo loc_11049127
  loc_11048F9F:   ' Referenced from: 11048B86
  loc_11048FA5:   var_8038 = Err
  loc_11048FAC:   Set var_30 = Err
  loc_11048FD8:   var_803C = Err
  loc_11048FDF:   Set var_34 = Err
  loc_1104908E:   MsgBox(CStr(Err.Number) & vbCrLf & Err.Description, 64, "提示", 10, 10)
  loc_110490D3: End If
  loc_110490D3: Exit Sub
  loc_110490DE: GoTo loc_11049127
  loc_11049126: Exit Sub
  loc_11049127: ' Referenced from: 11048F9A
  loc_11049127: ' Referenced from: 110490DE
End Sub
