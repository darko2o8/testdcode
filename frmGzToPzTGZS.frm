VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C908002B2F49FB}#1.2#0"; "C:\Windows\SysWow64\Comdlg32.ocx"
Begin VB.Form frmGzToPzTGZS
  Caption = "工资导转凭证（TGZS）"
  BackColor = &H80000005&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmGzToPzTGZS.frx":0000
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
    OleObjectBlob = "frmGzToPzTGZS.frx":014A
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
    OleObjectBlob = "frmGzToPzTGZS.frx":0293
    Begin AIFCmp1.asxStatusBar sBar
      Left = 0
      Top = 5355
      Width = 12045
      Height = 345
      OleObjectBlob = "frmGzToPzTGZS.frx":04BC
    End
    Begin C1SizerLibCtl.C1Elastic C1Elastic2
      Left = 0
      Top = 0
      Width = 12045
      Height = 435
      TabStop = 0   'False
      TabIndex = 1
      OleObjectBlob = "frmGzToPzTGZS.frx":05EC
    End
    Begin VSFlex8LCtl.VSFlexGrid VFG
      Left = 0
      Top = 1260
      Width = 12045
      Height = 4080
      TabIndex = 2
      OleObjectBlob = "frmGzToPzTGZS.frx":074F
      Begin MSComDlg.CommonDialog dlg
        OleObjectBlob = "frmGzToPzTGZS.frx":0BB8
        Left = 0
        Top = 0
      End
    End
    Begin AIFCmp1.asxPanel asxPanel1
      Left = 0
      Top = 450
      Width = 12045
      Height = 795
      OleObjectBlob = "frmGzToPzTGZS.frx":0C1C
      Begin TDBNumLite6Ctl.TDBNumLite TDBNum
        Index = 0
        Left = 120
        Top = 405
        Width = 1695
        Height = 270
        TabIndex = 16
        OleObjectBlob = "frmGzToPzTGZS.frx":0CFC
      End
      Begin VB.ComboBox Cbo
        Style = 2
        Left = 12030
        Top = 60
        Width = 4545
        Height = 300
        Visible = 0   'False
        TabIndex = 14
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 3
        Left = 7935
        Top = 75
        Width = 600
        Height = 270
        TabIndex = 9
        OleObjectBlob = "frmGzToPzTGZS.frx":0E68
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 4
        Left = 8595
        Top = 75
        Width = 720
        Height = 270
        TabIndex = 13
        OleObjectBlob = "frmGzToPzTGZS.frx":1008
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
        Top = 75
        Width = 870
        Height = 270
        TabIndex = 7
        OleObjectBlob = "frmGzToPzTGZS.frx":11A8
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 1
        Left = 6090
        Top = 75
        Width = 870
        Height = 270
        TabIndex = 8
        OleObjectBlob = "frmGzToPzTGZS.frx":13A0
      End
      Begin TDBText6Ctl.TDBText TDBText
        Left = 30
        Top = 75
        Width = 5115
        Height = 270
        TabIndex = 10
        OleObjectBlob = "frmGzToPzTGZS.frx":1570
        ToolTipText = "项目大类"
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 2
        Left = 7005
        Top = 75
        Width = 870
        Height = 270
        Visible = 0   'False
        TabIndex = 11
        OleObjectBlob = "frmGzToPzTGZS.frx":16CC
      End
      Begin TDBDate6Ctl.TDBDate TDBDate
        Left = 9390
        Top = 420
        Width = 2385
        Height = 285
        TabIndex = 12
        OleObjectBlob = "frmGzToPzTGZS.frx":1870
      End
      Begin TDBNumLite6Ctl.TDBNumLite TDBNum
        Index = 1
        Left = 2160
        Top = 405
        Width = 2295
        Height = 270
        Visible = 0   'False
        TabIndex = 17
        OleObjectBlob = "frmGzToPzTGZS.frx":1B5F
      End
      Begin VB.Label Label1
        Caption = "选择："
        Left = 11490
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

Attribute VB_Name = "frmGzToPzTGZS"


Private Sub TDBText_UnknownEvent_B '11093560
  Dim var_64 As frmGzToPzTGZS.dlg
  loc_110935C7: Set var_64 = frmGzToPzTGZS.dlg
  loc_110935F9: var_64.FileName = var_48
  loc_1109361E: var_64.DialogTitle = var_48
  loc_11093643: var_64.Filter = var_48
  loc_11093665: var_64.CancelError = var_48
  loc_1109366F: var_64.ShowOpen
  loc_11093681: var_64.FileName = var_64
  loc_110936C7: If (var_64 = global_1100AE28) Then
  loc_110936D5:   var_64.FileName = Me
  loc_1109371D:   frmGzToPzTGZS.TDBText.DispID_0000 = var_2C
  loc_11093747: End If
  loc_11093753: GoTo loc_1109377B
  loc_1109377A: Exit Sub
  loc_1109377B: ' Referenced from: 11093753
End Sub

Private  APB_UnknownEvent_9(arg_C) '110930F0
  Dim var_20 As Variant
  Dim var_AC As Scripting.FileSystemObject
  loc_11093167: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11093170: var_C4 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11093197: arg_C = frmGzToPzTGZS.APB.UnkVCall_00000040h
  loc_110931D5: var_B8 = var_24.DispID_FFFFFDFA
  loc_11093209: var_8008 = (var_B8 = "加载数据")
  loc_1109320D: If var_8008 = 0 Then
  loc_11093230:   var_AC = var_18
  loc_1109326B:   var_1C = frmGzToPzTGZS.TDBText.DispID_0000
  loc_1109327B:   var_8014 = Scripting.FileSystemObject.FileExists
  loc_110932B9:   If Not (var_A8) Then
  loc_1109331C:     MsgBox("文件不存在或非法路径！ ", 64, "提示", 10, 10)
  loc_11093342:   Else
  loc_11093354:     If frmGzToPzTGZS.FillDataNew < 0 Then
  loc_11093366:       var_A8 = CheckObj(%ecx = %S_edx_S = %S_edx_S, global_1100CFFC, 1788)
  loc_11093371:     End If
  loc_1109337D:     call ebx("取消加载", var_B8, var_1C, var_A8, var_24)
  loc_11093381:     If ebx("取消加载", var_B8, var_1C, var_A8, var_24) = 0 Then
  loc_110933B1:       var_44 = "提示信息"
  loc_110933DF:       var_2C = "是否取消数据载入？" & vbCrLf & "取消数据载入，数据将全部清空。"
  loc_110933FB:       MsgBox(var_2C, 292, var_44, var_54, var_64)
  loc_11093432:       If (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6) = 0 Then GoTo loc_110934E7
  loc_11093443:     Else
  loc_1109344F:       (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_11093453:       If var_B8 = 0 Then
  loc_11093458:         var_8020 = frmGzToPzTGZS.Proc_13_14_1108A020("凭证导入")
  loc_11093463:       Else
  loc_1109346F:         (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_11093473:         If var_8020 Then
  loc_11093481:           (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_11093485:           If var_B8 = 0 Then
  loc_110934B8:             Set var_20 = var_C4 = %S_edx_S
  loc_110934C6:             var_8028 = Global.Unload var_B8
  loc_110934E7:           End If
  loc_110934E7:         End If
  loc_110934E7:       End If
  loc_110934E7:     End If
  loc_110934E7:   End If
  loc_110934E7: End If
  loc_110934EF: GoTo loc_11093526
  loc_11093525: Exit Sub
  loc_11093526: ' Referenced from: 110934EF
End Sub

Private Sub Form_Load() '110823F0
  Dim var_18 As Variant
  Dim var_1C As var_18.DispID_03E8
  Dim var_20 As var_1C.DispID_03E8
  loc_11082460: Set var_18 = frmGzToPzTGZS.TDBText
  loc_11082467: var_30 = var_18.DispID_03E8
  loc_11082488: var_18.DispID_03E8.UnkVCall_00000030h
  loc_110824D6: Set var_18 = frmGzToPzTGZS.TDBDate
  loc_110824DD: var_30 = var_18.DispID_03E8
  loc_110824F2: Set var_1C = var_18.DispID_03E8
  loc_110824FE: var_1C.UnkVCall_00000030h
  loc_1108254D: frmGzToPzTGZS.TDBNum.UnkVCall_00000040h
  loc_11082579: var_30 = var_1C.DispID_03E8
  loc_1108258E: Set var_20 = var_1C.DispID_03E8
  loc_110825ED: frmGzToPzTGZS.TDBNum.UnkVCall_00000040h
  loc_11082619: var_30 = var_1C.DispID_03E8
  loc_1108262E: Set var_20 = var_1C.DispID_03E8
  loc_110826AD: frmGzToPzTGZS.TDBDate.DispID_0000 = Date
  loc_110826CF: Set var_18 = frmGzToPzTGZS.APB
  loc_110826DC: var_18.UnkVCall_00000040h
  loc_1108271A: var_1C.DispID_80010007 = var_18.DispID_03E8
  loc_11082741: Set var_18 = frmGzToPzTGZS.APB
  loc_1108274E: var_18.UnkVCall_00000040h
  loc_11082789: var_1C.DispID_80010007 = var_18.DispID_03E8
  loc_110827A5: var_8004 = frmGzToPzTGZS.Proc_13_11_11074C60(var_18)
  loc_110827B2: var_54 = frmGzToPzTGZS.getBTData
  loc_110827DA: GoTo loc_110827FD
  loc_110827FC: Exit Sub
  loc_110827FD: ' Referenced from: 110827DA
End Sub

Private Sub Form_Resize() '11082820
  loc_110828AD: var_38 = frmGzToPzTGZS.Pic1.DispID_80010005
  loc_110828D1: var_48 = frmGzToPzTGZS.Pic1.DispID_80010006
  loc_110828E4: var_EC = var_48.ScaleWidth
  loc_1108291B: If global_110F6000 = 0 Then
  loc_11082925: Else
  loc_11082930: End If
  loc_11082930: var_70 = ((var_EC - CSgn(var_38)) / 2)
  loc_11082945: var_F0 = var_48.ScaleHeight
  loc_11082983: If global_110F6000 = 0 Then
  loc_1108298D: Else
  loc_11082998: End If
  loc_11082AA3: frmGzToPzTGZS.Pic1.DispID_80011002(var_70, ((var_F0 - CSgn(var_48)) / 2), CSgn(frmGzToPzTGZS.Pic1.DispID_80010005), CSgn(frmGzToPzTGZS.Pic1.DispID_80010006))
  loc_11082AEC: GoTo loc_11082B26
End Sub

Public Sub ExitForm(Cancel, UnloadMode) '11074B80
  Dim var_18 As Global
  loc_11074BBF: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11074BEA: Set var_18 = Me
  loc_11074BF2: var_8008 = Global.Unload
  loc_11074C2C: GoTo loc_11074C38
  loc_11074C37: Exit Sub
  loc_11074C38: ' Referenced from: 11074C2C
End Sub

Public Function FillDataNew() '11076570
  Dim var_B4 As Variant
  Dim var_68 As Variant
  Dim var_B8 As frmGzToPzTGZS.TDBText
  Dim var_5C As Variant
  Dim var_3C As Variant
  Dim var_34 As Me
  Dim var_2C As ADODB.Recordset
  loc_110766C9: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110766DF: var_2E8 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11076723: frmGzToPzTGZS.VFG.DispID_0007 = 1
  loc_11076746: Set var_B4 = frmGzToPzTGZS.Label3
  loc_11076750: var_2C8 = var_B4
  loc_11076756: var_B4.Caption = "正在打开Excel数据表，请稍候。。。"
  loc_110767C9: frmGzToPzTGZS.Pic1.DispID_80010007 = True
  loc_110767F5: frmGzToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_1107680F: var_8004 = CreateObject(global_1100D5A4)
  loc_1107681A: Set var_68 = CreateObject(global_1100D5A4)
  loc_11076829: var_B4 = var_68.UnkVCall_000000D0h
  loc_110768B0: var_2CC = var_B4
  loc_11076AF8: Set var_B8 = frmGzToPzTGZS.TDBText
  loc_11076B20: var_94 = var_B8.DispID_0000
  loc_11076B30: var_94 = var_B4.UnkVCall_0000004Ch
  loc_11076BA3: var_B4 = var_5C.Tag
  loc_11076C1D: var_3C.BackColor = CInt(1)
  loc_11076C46: var_B4.Activate
  loc_11076CB2: Set var_90 = var_B4.UsedRange
  loc_11076D15: Set var_B4 = frmGzToPzTGZS.Pic1
  loc_11076D1C: var_B4.DispID_80010007 = var_1A4
  loc_11076D92: var_3C.UnkVCall_00000064h
  loc_11076E26: var_DC = var_B4.Cells(4, 1).value
  loc_11076E3D: var_94 = Proc_0_11_11029000(var_DC, var_3C, 2)
  loc_11076E45: var_8014 = (var_94 = "人员编号")
  loc_11076E60: var_254 = var_8014
  loc_11076E9A: var_DC.BackColor = CInt(1)
  loc_11076F2A: var_FC = var_B8.Cells(4, 4).value
  loc_110770E6: var_19C = var_8014 Or (LCase(Proc_0_11_11029000(var_FC, var_B8, var_1A8)) <> "部门名称") Or (LCase(Proc_0_11_11029000(var_BC.Cells(4, 7).value, var_BC, 1)) <> "岗位性质")
  loc_11077185: If CBool(var_19C) Then
  loc_110771D9:   frmGzToPzTGZS.Pic1.DispID_80010007 = var_1A4
  loc_1107724C:   var_C4 = frmGzToPzTGZS.TDBText
  loc_110772BC:   var_1AC = var_5C.UnkVCall_0000006Ch
  loc_110772F5:   var_1A8 = var_68.UnkVCall_00000398h
  loc_1107732A:   Set var_3C = {000208D7-0000-0000-C000000000000046}()
  loc_1107733A:   Set var_5C = {000208DA-0000-0000-C000000000000046}()
  loc_1107734A:   Set var_68 = {000208D5-0000-0000-C000000000000046}()
  loc_110773E2:   MsgBox("与所要求的格式不符！ ", 64, "提示", 10, 10)
  loc_11077410: Else
  loc_11077421:   Set var_B4 = frmGzToPzTGZS.Label3
  loc_1107742F:   var_2C8 = var_B4
  loc_11077435:   var_B4.Caption = "正在填充数据，请稍候。。。"
  loc_110774AC:   frmGzToPzTGZS.Pic1.DispID_80010007 = True
  loc_110774DD:   frmGzToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_11077517:   Set var_B4 = frmGzToPzTGZS.APB
  loc_11077529:   var_2C8 = var_B4
  loc_1107752F:   var_B4.UnkVCall_00000040h
  loc_110775C5:   Set var_B4 = frmGzToPzTGZS.APB
  loc_110775D7:   var_2C8 = var_B4
  loc_110775DD:   var_B4.UnkVCall_00000040h
  loc_11077687:   frmGzToPzTGZS.APB.UnkVCall_00000040h
  loc_1107775A:   var_CC = var_90.Rows
  loc_11077783:   var_EC = var_CC.Count - 2
  loc_11077826:   Set var_B4 = frmGzToPzTGZS.sBar
  loc_1107782D:   var_B4.DispID_6803001E(1100D68Ch & var_EC & "条记录")
  loc_11077874:   var_34 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[T_CY_GzZGZS_Temp]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1) "
  loc_11077883:   var_8030 = var_34 & "CREATE TABLE [T_CY_GzZGZS_Temp](iNo INT NULL,cCode VARCHAR(50) NULL,cDepCode VARCHAR(50) NULL,cGzItem VARCHAR(50),iMoney Money NULL)"
  loc_1107788E:   call edi(var_B4, 00000002h, var_B8, var_B4, 00000001h, var_B8, var_B4, 00000000h, var_B8, 2, var_1A0, var_CC, var_C8, var_C4, var_C0, 0000000Ah)
  loc_110778D3:   var_CC = UnkObj.UnkVCall_00000040h
  loc_11077917:   var_34 = "DELETE FROM [T_CY_GzZGZS_Temp]"
  loc_110779BB:   Set var_B4 = frmGzToPzTGZS.TDBDate
  loc_110779E7:   var_D4 = var_B4.DispID_004E
  loc_110779F7:   var_EC)
  loc_11077A55:   var_78 = CByte("DateToPeriod".00000001h)
  loc_11077AAE:   var_DC = var_90.Rows.Count
  loc_11077AE9:   If var_18 <= CLng(var_DC) Then
  loc_11077AF7:     If global_56 = 0 Then
  loc_11077B03:       var_1B4 = var_18
  loc_11077B64:       var_3C.UnkVCall_00000064h
  loc_11077BC6:       var_DC.BackColor = var_1E4
  loc_11077C62:       var_8040 = Proc_0_11_11029000(var_B8.Cells(var_18, 6).value, var_B8, var_3C)
  loc_11077C6F:       call edi(00000002h, var_1A8, var_1A4, var_1A0, var_B4, var_B4, var_1B8, 80020004h, var_1B0, 00000409h, var_1A0, var_B4, var_1AC, var_1A8, var_1A4, var_1A0)
  loc_11077C77:       var_8044 = (edi(00000002h, var_1A8, var_1A4, var_1A0, var_B4, var_B4, var_1B8, 80020004h, var_1B0, 00000409h, var_1A0, var_B4, var_1AC, var_1A8, var_1A4, var_1A0) = global_1100AE28)
  loc_11077C94:       var_390 = var_8044 + 1
  loc_11077D12:       var_8048 = Proc_0_11_11029000(var_B4.Cells(var_1B4, var_1C4).value, &H4003, var_1B8)
  loc_11077D1F:       call edi(var_1B4, var_1B0, var_1CC, var_1C8, var_1C4, var_1C0, var_1DC, var_1D8, var_1D4, var_1D0, var_1EC, var_1E8, var_1E4, var_1E0, 0000000Ah, var_1F8)
  loc_11077D27:       var_804C = (edi(var_1B4, var_1B0, var_1CC, var_1C8, var_1C4, var_1C0, var_1DC, var_1D8, var_1D4, var_1D0, var_1EC, var_1E8, var_1E4, var_1E0, 0000000Ah, var_1F8) = global_1100AE28)
  loc_11077D4D:       var_2D0 = (var_8044 + 1 And var_804C)
  loc_11077D9B:       If var_2D0 = 0 Then
  loc_11077DC3:         var_8050 = CStr(var_18(-2))
  loc_11077DD1:         call edi("正在填充数据：", 80020004h, var_1F0, 0000000Ah, var_208, 80020004h, var_200, 0000000Ah, var_218)
  loc_11077DD4:         var_8054 =  & edi("正在填充数据：", 80020004h, var_1F0, 0000000Ah, var_208, 80020004h, var_200, 0000000Ah, var_218)
  loc_11077DE2:         call edi
  loc_11077E65:         Set var_B4 = frmGzToPzTGZS.sBar
  loc_11077E6C:         var_B4.DispID_6803001E(8 & "条记录")
  loc_11077EC0:         var_54 = var_54(1)
  loc_11077F1C:         var_3C.UnkVCall_00000064h
  loc_11077FB4:         var_DC = var_B4.Cells(var_18, &H55).value
  loc_11077FBE:         var_805C = Proc_0_11_11029000(var_DC, var_3C, 2)
  loc_11077FC8:         call edi(var_1A8, 1, var_1A0, var_B4)
  loc_110780ED:         var_DC = var_B4.Cells(var_18, &H4C).value
  loc_110780F7:         var_8060 = Proc_0_12_110291B0(var_DC, var_B4)
  loc_11078104:         call edi
  loc_11078178:         var_80 = Format(0, "0.00")
  loc_11078236:         var_DC.BackColor = CInt(1)
  loc_110782CE:         var_DC = var_B4.Cells(var_18, 7).value
  loc_110782D8:         var_8064 = Proc_0_11_11029000(var_DC, var_B4, 1)
  loc_110782E5:         call edi(00000001h)
  loc_110783D5:         var_8068 = Proc_0_11_11029000(var_B8.Cells(var_18, 6).value, var_B8)
  loc_110783E2:         call edi
  loc_110783F2:         var_300 = var_A0
  loc_11078416:         call edi(var_9C)
  loc_11078425:         call edi(edi(var_9C))
  loc_11078433:         var_806C = = frmGzToPzTGZS.GetKmCode("工资", edi(edi(var_9C)), )
  loc_11078462:         call edi
  loc_110784CF:         var_34 = "INSERT INTO [T_CY_GzZGZS_Temp](iNo,cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_110784EC:         var_1A4 = var_64
  loc_110784FA:         var_1B4 = var_24
  loc_11078500:         var_8070 = CStr(var_54)
  loc_1107850E:         call edi(var_34)
  loc_11078511:         var_8074 =  & edi(var_34)
  loc_1107851F:         call edi
  loc_11078527:         var_8078 = 8 & global_1100AC40
  loc_11078535:         call edi
  loc_1107853F:         var_807C = Proc_0_10_11028DD0(&H4008, 8)
  loc_1107854C:         call edi
  loc_1107854F:         var_8080 =  & 8
  loc_1107855D:         call edi
  loc_11078565:         var_8084 = 8 & global_1100AC40
  loc_11078573:         call edi
  loc_1107857D:         var_8088 = Proc_0_10_11028DD0(&H4008, 8)
  loc_1107858A:         call edi
  loc_1107858D:         var_808C =  & 8
  loc_1107859B:         call edi
  loc_110785A3:         var_8090 = 8 & ",'工资',"
  loc_110785AE:         call edi
  loc_110785FF:         var_8094 = CStr(var_80)
  loc_1107860D:         call edi(var_7C, var_34)
  loc_11078610:         var_8098 =  & edi(var_7C, var_34)
  loc_1107861B:         call edi
  loc_11078632:         var_809C = var_34 & global_1100BD88
  loc_1107863D:         call edi
  loc_110787C0:         var_DC = var_B4.Cells(var_18, &H4D).value
  loc_110787CA:         var_80A0 = Proc_0_12_110291B0(var_DC, var_B4)
  loc_110787D7:         call edi
  loc_1107884B:         var_80 = Format(0, "0.00")
  loc_11078909:         var_DC.BackColor = CInt(1)
  loc_110789A1:         var_DC = var_B4.Cells(var_18, 7).value
  loc_110789AB:         var_80A4 = Proc_0_11_11029000(var_DC, var_B4, 1)
  loc_110789B8:         call edi(00000001h)
  loc_11078AA8:         var_80A8 = Proc_0_11_11029000(var_B8.Cells(var_18, 6).value, var_B8)
  loc_11078AB5:         call edi
  loc_11078AC5:         var_310 = var_A0
  loc_11078AE9:         call edi(var_9C)
  loc_11078AF8:         call edi(edi(var_9C))
  loc_11078B06:         var_80AC = = frmGzToPzTGZS.GetKmCode("社保", edi(edi(var_9C)), )
  loc_11078B35:         call edi
  loc_11078BA2:         var_34 = "INSERT INTO [T_CY_GzZGZS_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11078BBF:         var_1A4 = var_64
  loc_11078BD0:         var_1B4 = var_24
  loc_11078BD6:         var_80B0 = Proc_0_10_11028DD0(&H4008, var_34)
  loc_11078BE3:         call edi
  loc_11078BE6:         var_80B4 =  & 8
  loc_11078BF4:         call edi
  loc_11078BFC:         var_80B8 = 8 & global_1100AC40
  loc_11078C0A:         call edi
  loc_11078C14:         var_80BC = Proc_0_10_11028DD0(&H4008, 8)
  loc_11078C21:         call edi
  loc_11078C24:         var_80C0 =  & 8
  loc_11078C32:         call edi
  loc_11078C3A:         var_80C4 = 8 & ",'社保',"
  loc_11078C45:         call edi
  loc_11078C81:         var_80C8 = CStr(var_80)
  loc_11078C8F:         call edi(var_7C, var_34)
  loc_11078C92:         var_80CC =  & edi(var_7C, var_34)
  loc_11078C9D:         call edi
  loc_11078CB4:         var_80D0 = var_34 & global_1100BD88
  loc_11078CBF:         call edi
  loc_11078E42:         var_DC = var_B4.Cells(var_18, &H4F).value
  loc_11078E4C:         var_80D4 = Proc_0_12_110291B0(var_DC, var_B4)
  loc_11078E59:         call edi
  loc_11078ECD:         var_80 = Format(0, "0.00")
  loc_11078F8B:         var_DC.BackColor = CInt(1)
  loc_11079023:         var_DC = var_B4.Cells(var_18, 7).value
  loc_1107902D:         var_80D8 = Proc_0_11_11029000(var_DC, var_B4, 1)
  loc_1107903A:         call edi(00000001h)
  loc_1107912A:         var_80DC = Proc_0_11_11029000(var_B8.Cells(var_18, 6).value, var_B8)
  loc_11079137:         call edi
  loc_11079147:         var_320 = var_A0
  loc_1107916B:         call edi(var_9C)
  loc_1107917A:         call edi(edi(var_9C))
  loc_11079188:         var_80E0 = = frmGzToPzTGZS.GetKmCode("公积金", edi(edi(var_9C)), )
  loc_110791B7:         call edi
  loc_11079224:         var_34 = "INSERT INTO [T_CY_GzZGZS_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11079241:         var_1A4 = var_64
  loc_11079252:         var_1B4 = var_24
  loc_11079258:         var_80E4 = Proc_0_10_11028DD0(&H4008, var_34)
  loc_11079265:         call edi
  loc_11079268:         var_80E8 =  & 8
  loc_11079276:         call edi
  loc_1107927E:         var_80EC = 8 & global_1100AC40
  loc_1107928C:         call edi
  loc_11079296:         var_80F0 = Proc_0_10_11028DD0(&H4008, 8)
  loc_110792A3:         call edi
  loc_110792A6:         var_80F4 =  & 8
  loc_110792B4:         call edi
  loc_110792BC:         var_80F8 = 8 & ",'公积金',"
  loc_110792C7:         call edi
  loc_11079303:         var_80FC = CStr(var_80)
  loc_11079311:         call edi(var_7C, var_34)
  loc_11079314:         var_8100 =  & edi(var_7C, var_34)
  loc_1107931F:         call edi
  loc_11079336:         var_8104 = var_34 & global_1100BD88
  loc_11079341:         call edi
  loc_110793D6:         Set var_B4 = frmGzToPzTGZS.TDBNum
  loc_110793E8:         var_2C8 = var_B4
  loc_110793EE:         var_B4.UnkVCall_00000040h
  loc_11079428:         var_8108 = var_B8.DispID_0043
  loc_11079436:         call edi(var_B4, 00000000h, var_B8)
  loc_1107948B:         If (edi(var_B4, 00000000h, var_B8) = global_1100AE28) + 1 Then
  loc_110794A0:         Else
  loc_110794B4:           Set var_B4 = frmGzToPzTGZS.TDBNum
  loc_110794C6:           var_2C8 = var_B4
  loc_110794CC:           var_B4.UnkVCall_00000040h
  loc_110794FC:           var_CC = var_B8.DispID_0043
  loc_11079506:           var_8110 = var_CC
  loc_11079514:           call edi(var_B4, 00000000h, var_B8)
  loc_1107951D:           var_44 = edi(var_B4, 00000000h, var_B8)
  loc_11079551:         End If
  loc_1107962A:         var_FC = var_B8.Cells(var_18, 37).value
  loc_11079634:         var_8114 = Proc_0_12_110291B0(var_FC, var_B8)
  loc_11079641:         call edi
  loc_110796DE:         var_328 = var_9C
  loc_11079765:         var_8118 = Proc_0_12_110291B0(var_B4.Cells(var_18, &H24).value, var_B4)
  loc_11079772:         call edi
  loc_1107978D:         call edi
  loc_110797A8:         var_80 = (8 + 8)
  loc_1107981F:         var_DC = "0.00"
  loc_11079842:         If global_110F6000 = 0 Then
  loc_1107984C:         Else
  loc_1107985D:         End If
  loc_1107992B:         var_DC.BackColor = CInt(1)
  loc_110799C3:         var_DC = var_B4.Cells(var_18, 7).value
  loc_110799CD:         var_811C = Proc_0_11_11029000(var_DC, var_B4, 1)
  loc_110799DA:         call edi(00000001h)
  loc_11079ACA:         var_8120 = Proc_0_11_11029000(var_B8.Cells(var_18, 6).value, var_B8)
  loc_11079AD7:         call edi
  loc_11079AE7:         var_330 = var_A0
  loc_11079B0B:         call edi(var_9C)
  loc_11079B1A:         call edi(edi(var_9C))
  loc_11079B28:         var_8124 = = frmGzToPzTGZS.GetKmCode("奖金", edi(edi(var_9C)), )
  loc_11079B57:         call edi
  loc_11079BC4:         var_34 = "INSERT INTO [T_CY_GzZGZS_Temp](iNo,cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11079BE1:         var_1A4 = var_64
  loc_11079BEF:         var_1B4 = var_24
  loc_11079BF5:         var_8128 = CStr(var_54)
  loc_11079C03:         call edi(var_34)
  loc_11079C06:         var_812C =  & edi(var_34)
  loc_11079C14:         call edi
  loc_11079C1C:         var_8130 = 8 & global_1100AC40
  loc_11079C2A:         call edi
  loc_11079C34:         var_8134 = Proc_0_10_11028DD0(&H4008, 8)
  loc_11079C41:         call edi
  loc_11079C44:         var_8138 =  & 8
  loc_11079C52:         call edi
  loc_11079C5A:         var_813C = 8 & global_1100AC40
  loc_11079C68:         call edi
  loc_11079C72:         var_8140 = Proc_0_10_11028DD0(&H4008, 8)
  loc_11079C7F:         call edi
  loc_11079C82:         var_8144 =  & 8
  loc_11079C90:         call edi
  loc_11079C98:         var_8148 = 8 & ",'奖金',"
  loc_11079CA3:         call edi
  loc_11079CF4:         var_814C = CStr(Format(((var_80 * var_44) / 12), var_DC))
  loc_11079D02:         call edi(var_7C, var_34)
  loc_11079D05:         var_8150 =  & edi(var_7C, var_34)
  loc_11079D10:         call edi
  loc_11079D27:         var_8154 = var_34 & global_1100BD88
  loc_11079D32:         call edi
  loc_11079DC7:         var_28 = var_28(1)
  loc_11079DCF:         If var_18 Mod 00000064h = 0 Then
  loc_11079DD1:           DoEvents
  loc_11079DD7:         End If
  loc_11079DD7:       End If
  loc_11079DE7:       var_18 = 1+var_18
  loc_11079DEA:       GoTo loc_11077AE3
  loc_11079DEF:     End If
  loc_11079E3E:     frmGzToPzTGZS.VFG.DispID_0007 = 1
  loc_11079E6A:     global_56 = 0
  loc_11079E7E:     Set var_B4 = frmGzToPzTGZS.APB
  loc_11079E90:     var_2C8 = var_B4
  loc_11079E96:     var_B4.UnkVCall_00000040h
  loc_11079F2F:     Set var_B4 = frmGzToPzTGZS.APB
  loc_11079F41:     var_2C8 = var_B4
  loc_11079F47:     var_B4.UnkVCall_00000040h
  loc_11079FE0:     Set var_B4 = frmGzToPzTGZS.APB
  loc_11079FF2:     var_2C8 = var_B4
  loc_11079FF8:     var_B4.UnkVCall_00000040h
  loc_1107A069:   End If
  loc_1107A083:   var_1B4 = var_78
  loc_1107A0F9:   var_8158 = "cIYear".00000000h & 1100D700h & var_78 & "月应付工资"
  loc_1107A104:   call edi(00000002h, var_B8, var_B4, 00000001h, var_B8, var_B4, 00000000h, var_B8)
  loc_1107A14C:   var_2C8 = var_2C
  loc_1107A152:   var_2C4 = ADODB.Recordset.State
  loc_1107A17D:   If var_2C4 = 1 Then
  loc_1107A19B:     var_2C8 = var_2C
  loc_1107A1A1:     var_8164 = ADODB.Recordset.Close
  loc_1107A1C5:   End If
  loc_1107A23B:   var_2C8 = var_2C
  loc_1107A278:   var_816C = ADODB.Recordset.Open("SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZS_Temp] WHERE cGzItem='工资' GROUP BY ccode,cdepcode", var_1A8, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZS_Temp] WHERE cGzItem='工资' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_1107A2C7:   var_2C8 = var_2C
  loc_1107A2CD:   var_2C0 = ADODB.Recordset.EOF
  loc_1107A2F3:   If var_2C0 = 0 Then
  loc_1107A301:     var_58 = "1"
  loc_1107A366:     var_8174 = var_58 & Chr(9) & var_1B4
  loc_1107A371:     call edi(var_1B8, var_1B4, var_1B0, 00000003h, 00000001h, FFFFFFFFh)
  loc_1107A420:     var_817C = var_58 & Chr(9) & frmGzToPzTGZS.TDBDate.DispID_004E
  loc_1107A42B:     call edi
  loc_1107A4C2:     var_8180 = var_58 & Chr(9) & 1100D6D4h
  loc_1107A4CD:     call edi
  loc_1107A54A:     var_8184 = var_58 & Chr(9) & 1100C008h
  loc_1107A555:     call edi
  loc_1107A5D1:     var_8188 = var_58 & Chr(9) & var_50
  loc_1107A5DC:     call edi
  loc_1107A63F:     var_2C8 = var_2C
  loc_1107A680:     var_2D0 = ADODB.Recordset.Fields
  loc_1107A6B6:     ADODB.Recordset.8 = Forms
  loc_1107A6DA:     var_B8 = 0
  loc_1107A6E4:     var_E4 = var_B8
  loc_1107A739:     var_8194 = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_1107A744:     call edi(var_1B0, var_B8)
  loc_1107A7C1:     var_2C8 = var_2C
  loc_1107A80C:     var_2D0 = ADODB.Recordset.Fields
  loc_1107A838:     ADODB.Recordset.8 = Forms
  loc_1107A863:     var_B8 = 0
  loc_1107A86D:     var_E4 = var_B8
  loc_1107A8BB:     var_81A0 = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_1107A8C6:     call edi(var_1B0, var_B8)
  loc_1107A95D:     var_81A4 = var_58 & Chr(9) & 1100C008h
  loc_1107A968:     call edi
  loc_1107A9E5:     var_81A8 = var_58 & Chr(9) & 1100AE28h
  loc_1107A9F0:     call edi
  loc_1107AA6D:     var_81AC = var_58 & Chr(9) & 1100AE28h
  loc_1107AA78:     call edi
  loc_1107AAF5:     var_81B0 = var_58 & Chr(9) & 1100AE28h
  loc_1107AB00:     call edi
  loc_1107AB7D:     var_81B4 = var_58 & Chr(9) & 1100AE28h
  loc_1107AB88:     call edi
  loc_1107AC05:     var_81B8 = var_58 & Chr(9) & 1100AE28h
  loc_1107AC10:     call edi
  loc_1107AC8D:     var_81BC = var_58 & Chr(9) & 1100AE28h
  loc_1107AC98:     call edi
  loc_1107AD15:     var_81C0 = var_58 & Chr(9) & 1100AE28h
  loc_1107AD20:     call edi
  loc_1107AD83:     var_2C8 = var_2C
  loc_1107ADC4:     var_2D0 = ADODB.Recordset.Fields
  loc_1107ADFA:     ADODB.Recordset.8 = Forms
  loc_1107AE1E:     var_B8 = 0
  loc_1107AE28:     var_E4 = var_B8
  loc_1107AE7D:     var_81CC = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_1107AE88:     call edi(var_1B0, var_B8)
  loc_1107AF1F:     var_81D0 = var_58 & Chr(9) & 1100AE28h
  loc_1107AF2A:     call edi
  loc_1107AFA7:     var_81D4 = var_58 & Chr(9) & 1100AE28h
  loc_1107AFB2:     call edi
  loc_1107B02F:     var_81D8 = var_58 & Chr(9) & 1100AE28h
  loc_1107B03A:     call edi
  loc_1107B0B7:     var_81DC = var_58 & Chr(9) & 1100AE28h
  loc_1107B0C2:     call edi
  loc_1107B13F:     var_81E0 = var_58 & Chr(9) & 1100AE28h
  loc_1107B14A:     call edi
  loc_1107B18B:     var_2C8 = var_2C
  loc_1107B1D6:     var_2D0 = ADODB.Recordset.Fields
  loc_1107B202:     ADODB.Recordset.8 = Forms
  loc_1107B22D:     var_B8 = 0
  loc_1107B237:     var_C4 = var_B8
  loc_1107B247:     var_81E8 = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_1107B254:     call edi(var_1A0, var_B8)
  loc_1107B26C:     call edi
  loc_1107B27E:     var_20 = (8 + var_20)
  loc_1107B2D2:     var_2C8 = var_2C
  loc_1107B2D8:     var_81F0 = ADODB.Recordset.MoveNext
  loc_1107B34E:     frmGzToPzTGZS.VFG.DispID_0080(var_58)
  loc_1107B363:     GoTo loc_1107A2A4
  loc_1107B368:   End If
  loc_1107B38B:   var_2C8 = var_2C
  loc_1107B3B6:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_1107B3C4:     var_58 = "1"
  loc_1107B429:     var_81F8 = var_58 & Chr(9) & 1100AE28h
  loc_1107B434:     call edi
  loc_1107B4E3:     var_8200 = var_58 & Chr(9) & frmGzToPzTGZS.TDBDate.DispID_004E
  loc_1107B4EE:     call edi
  loc_1107B585:     var_8204 = var_58 & Chr(9) & 1100D6D4h
  loc_1107B590:     call edi
  loc_1107B60D:     var_8208 = var_58 & Chr(9) & 1100C008h
  loc_1107B618:     call edi
  loc_1107B694:     var_820C = var_58 & Chr(9) & var_50
  loc_1107B69F:     call edi
  loc_1107B71C:     var_8210 = var_58 & Chr(9) & "215101"
  loc_1107B727:     call edi
  loc_1107B7A4:     var_8214 = var_58 & Chr(9) & 1100C008h
  loc_1107B7AF:     call edi
  loc_1107B801:     var_1B0 = var_1C
  loc_1107B834:     var_8218 = var_58 & Chr(9) & var_20
  loc_1107B83F:     call edi
  loc_1107B8BC:     var_821C = var_58 & Chr(9) & 1100AE28h
  loc_1107B8C7:     call edi
  loc_1107B944:     var_8220 = var_58 & Chr(9) & 1100AE28h
  loc_1107B94F:     call edi
  loc_1107B9CC:     var_8224 = var_58 & Chr(9) & 1100AE28h
  loc_1107B9D7:     call edi
  loc_1107BA54:     var_8228 = var_58 & Chr(9) & 1100AE28h
  loc_1107BA5F:     call edi
  loc_1107BADC:     var_822C = var_58 & Chr(9) & 1100AE28h
  loc_1107BAE7:     call edi
  loc_1107BB64:     var_8230 = var_58 & Chr(9) & 1100AE28h
  loc_1107BB6F:     call edi
  loc_1107BBEC:     var_8234 = var_58 & Chr(9) & 1100AE28h
  loc_1107BBF7:     call edi
  loc_1107BC74:     var_8238 = var_58 & Chr(9) & 1100AE28h
  loc_1107BC7F:     call edi
  loc_1107BCFC:     var_823C = var_58 & Chr(9) & 1100AE28h
  loc_1107BD07:     call edi
  loc_1107BD84:     var_8240 = var_58 & Chr(9) & 1100AE28h
  loc_1107BD8F:     call edi
  loc_1107BE0C:     var_8244 = var_58 & Chr(9) & 1100AE28h
  loc_1107BE17:     call edi
  loc_1107BE94:     var_8248 = var_58 & Chr(9) & 1100AE28h
  loc_1107BE9F:     call edi
  loc_1107BF1C:     var_824C = var_58 & Chr(9) & 1100AE28h
  loc_1107BF27:     call edi
  loc_1107BF91:     frmGzToPzTGZS.VFG.DispID_0080(var_58)
  loc_1107BFA6:   End If
  loc_1107BFD6:   var_1C4 = var_78
  loc_1107BFEF:   var_1A4 = "计提"
  loc_1107C05E:   var_8250 =  & "cIYear".0 & 1100D700h & var_78 & "月社保"
  loc_1107C069:   call edi
  loc_1107C0B8:   var_2C8 = var_2C
  loc_1107C0BE:   var_2C4 = ADODB.Recordset.State
  loc_1107C0E9:   If var_2C4 = 1 Then
  loc_1107C107:     var_2C8 = var_2C
  loc_1107C10D:     var_825C = ADODB.Recordset.Close
  loc_1107C131:   End If
  loc_1107C1A7:   var_2C8 = var_2C
  loc_1107C1E4:   var_8264 = ADODB.Recordset.Open("SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZS_Temp] WHERE cGzItem='社保' GROUP BY ccode,cdepcode", var_1A8, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZS_Temp] WHERE cGzItem='社保' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_1107C22B:   var_2C8 = var_2C
  loc_1107C231:   var_2C0 = ADODB.Recordset.EOF
  loc_1107C257:   If var_2C0 = 0 Then
  loc_1107C265:     var_58 = "2"
  loc_1107C2CA:     var_826C = var_58 & Chr(9) & var_1B4
  loc_1107C2D5:     call edi(var_1B8, var_1B4, var_1B0, 00000003h, 00000001h, FFFFFFFFh)
  loc_1107C384:     var_8274 = var_58 & Chr(9) & frmGzToPzTGZS.TDBDate.DispID_004E
  loc_1107C38F:     call edi
  loc_1107C426:     var_8278 = var_58 & Chr(9) & 1100D6D4h
  loc_1107C431:     call edi
  loc_1107C4AE:     var_827C = var_58 & Chr(9) & 1100C008h
  loc_1107C4B9:     call edi
  loc_1107C535:     var_8280 = var_58 & Chr(9) & var_50
  loc_1107C540:     call edi
  loc_1107C5A3:     var_2C8 = var_2C
  loc_1107C5E4:     var_2D0 = ADODB.Recordset.Fields
  loc_1107C61A:     ADODB.Recordset.8 = Forms
  loc_1107C63E:     var_B8 = 0
  loc_1107C648:     var_E4 = var_B8
  loc_1107C69D:     var_828C = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_1107C6A8:     call edi(var_1B0, var_B8)
  loc_1107C725:     var_2C8 = var_2C
  loc_1107C770:     var_2D0 = ADODB.Recordset.Fields
  loc_1107C79C:     ADODB.Recordset.8 = Forms
  loc_1107C7C7:     var_B8 = 0
  loc_1107C7D1:     var_E4 = var_B8
  loc_1107C81F:     var_8298 = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_1107C82A:     call edi(var_1B0, var_B8)
  loc_1107C8C1:     var_829C = var_58 & Chr(9) & 1100C008h
  loc_1107C8CC:     call edi
  loc_1107C949:     var_82A0 = var_58 & Chr(9) & 1100AE28h
  loc_1107C954:     call edi
  loc_1107C9D1:     var_82A4 = var_58 & Chr(9) & 1100AE28h
  loc_1107C9DC:     call edi
  loc_1107CA59:     var_82A8 = var_58 & Chr(9) & 1100AE28h
  loc_1107CA64:     call edi
  loc_1107CAE1:     var_82AC = var_58 & Chr(9) & 1100AE28h
  loc_1107CAEC:     call edi
  loc_1107CB69:     var_82B0 = var_58 & Chr(9) & 1100AE28h
  loc_1107CB74:     call edi
  loc_1107CBF1:     var_82B4 = var_58 & Chr(9) & 1100AE28h
  loc_1107CBFC:     call edi
  loc_1107CC79:     var_82B8 = var_58 & Chr(9) & 1100AE28h
  loc_1107CC84:     call edi
  loc_1107CCE7:     var_2C8 = var_2C
  loc_1107CD28:     var_2D0 = ADODB.Recordset.Fields
  loc_1107CD5E:     ADODB.Recordset.8 = Forms
  loc_1107CD82:     var_B8 = 0
  loc_1107CD8C:     var_E4 = var_B8
  loc_1107CDE1:     var_82C4 = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_1107CDEC:     call edi(var_1B0, var_B8)
  loc_1107CE83:     var_82C8 = var_58 & Chr(9) & 1100AE28h
  loc_1107CE8E:     call edi
  loc_1107CF0B:     var_82CC = var_58 & Chr(9) & 1100AE28h
  loc_1107CF16:     call edi
  loc_1107CF93:     var_82D0 = var_58 & Chr(9) & 1100AE28h
  loc_1107CF9E:     call edi
  loc_1107D01B:     var_82D4 = var_58 & Chr(9) & 1100AE28h
  loc_1107D026:     call edi
  loc_1107D0A3:     var_82D8 = var_58 & Chr(9) & 1100AE28h
  loc_1107D0AE:     call edi
  loc_1107D0EF:     var_2C8 = var_2C
  loc_1107D13A:     var_2D0 = ADODB.Recordset.Fields
  loc_1107D166:     ADODB.Recordset.8 = Forms
  loc_1107D191:     var_B8 = 0
  loc_1107D19B:     var_C4 = var_B8
  loc_1107D1AB:     var_82E0 = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_1107D1B8:     call edi(var_1A0, var_B8)
  loc_1107D1D0:     call edi
  loc_1107D1E2:     var_20 = (8 + var_20)
  loc_1107D236:     var_2C8 = var_2C
  loc_1107D23C:     var_82E8 = ADODB.Recordset.MoveNext
  loc_1107D2B2:     frmGzToPzTGZS.VFG.DispID_0080(var_58)
  loc_1107D2C7:     GoTo loc_1107C208
  loc_1107D2CC:   End If
  loc_1107D2EF:   var_2C8 = var_2C
  loc_1107D31A:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_1107D328:     var_58 = "2"
  loc_1107D38D:     var_82F0 = var_58 & Chr(9) & 1100AE28h
  loc_1107D398:     call edi
  loc_1107D447:     var_82F8 = var_58 & Chr(9) & frmGzToPzTGZS.TDBDate.DispID_004E
  loc_1107D452:     call edi
  loc_1107D4E9:     var_82FC = var_58 & Chr(9) & 1100D6D4h
  loc_1107D4F4:     call edi
  loc_1107D571:     var_8300 = var_58 & Chr(9) & 1100C008h
  loc_1107D57C:     call edi
  loc_1107D5F8:     var_8304 = var_58 & Chr(9) & var_50
  loc_1107D603:     call edi
  loc_1107D680:     var_8308 = var_58 & Chr(9) & "215303"
  loc_1107D68B:     call edi
  loc_1107D708:     var_830C = var_58 & Chr(9) & 1100C008h
  loc_1107D713:     call edi
  loc_1107D765:     var_1B0 = var_1C
  loc_1107D798:     var_8310 = var_58 & Chr(9) & var_20
  loc_1107D7A3:     call edi
  loc_1107D820:     var_8314 = var_58 & Chr(9) & 1100AE28h
  loc_1107D82B:     call edi
  loc_1107D8A8:     var_8318 = var_58 & Chr(9) & 1100AE28h
  loc_1107D8B3:     call edi
  loc_1107D930:     var_831C = var_58 & Chr(9) & 1100AE28h
  loc_1107D93B:     call edi
  loc_1107D9B8:     var_8320 = var_58 & Chr(9) & 1100AE28h
  loc_1107D9C3:     call edi
  loc_1107DA40:     var_8324 = var_58 & Chr(9) & 1100AE28h
  loc_1107DA4B:     call edi
  loc_1107DAC8:     var_8328 = var_58 & Chr(9) & 1100AE28h
  loc_1107DAD3:     call edi
  loc_1107DB50:     var_832C = var_58 & Chr(9) & 1100AE28h
  loc_1107DB5B:     call edi
  loc_1107DBD8:     var_8330 = var_58 & Chr(9) & 1100AE28h
  loc_1107DBE3:     call edi
  loc_1107DC60:     var_8334 = var_58 & Chr(9) & 1100AE28h
  loc_1107DC6B:     call edi
  loc_1107DCE8:     var_8338 = var_58 & Chr(9) & 1100AE28h
  loc_1107DCF3:     call edi
  loc_1107DD70:     var_833C = var_58 & Chr(9) & 1100AE28h
  loc_1107DD7B:     call edi
  loc_1107DDF8:     var_8340 = var_58 & Chr(9) & 1100AE28h
  loc_1107DE03:     call edi
  loc_1107DE80:     var_8344 = var_58 & Chr(9) & 1100AE28h
  loc_1107DE8B:     call edi
  loc_1107DEF5:     frmGzToPzTGZS.VFG.DispID_0080(var_58)
  loc_1107DF0A:   End If
  loc_1107DF3A:   var_1C4 = var_78
  loc_1107DF53:   var_1A4 = "计提"
  loc_1107DFC2:   var_8348 =  & "cIYear".0 & 1100D700h & var_78 & "月住房公积金"
  loc_1107DFCD:   call edi
  loc_1107E01C:   var_2C8 = var_2C
  loc_1107E022:   var_2C4 = ADODB.Recordset.State
  loc_1107E04D:   If var_2C4 = 1 Then
  loc_1107E06B:     var_2C8 = var_2C
  loc_1107E071:     var_8354 = ADODB.Recordset.Close
  loc_1107E095:   End If
  loc_1107E10B:   var_2C8 = var_2C
  loc_1107E148:   var_835C = ADODB.Recordset.Open("SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZS_Temp] WHERE cGzItem='公积金' GROUP BY ccode,cdepcode", var_1A8, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZS_Temp] WHERE cGzItem='公积金' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_1107E18F:   var_2C8 = var_2C
  loc_1107E195:   var_2C0 = ADODB.Recordset.EOF
  loc_1107E1BB:   If var_2C0 = 0 Then
  loc_1107E1C9:     var_58 = "3"
  loc_1107E22E:     var_8364 = var_58 & Chr(9) & var_1B4
  loc_1107E239:     call edi(var_1B8, var_1B4, var_1B0, 00000003h, 00000001h, FFFFFFFFh)
  loc_1107E2E8:     var_836C = var_58 & Chr(9) & frmGzToPzTGZS.TDBDate.DispID_004E
  loc_1107E2F3:     call edi
  loc_1107E38A:     var_8370 = var_58 & Chr(9) & 1100D6D4h
  loc_1107E395:     call edi
  loc_1107E412:     var_8374 = var_58 & Chr(9) & 1100C008h
  loc_1107E41D:     call edi
  loc_1107E499:     var_8378 = var_58 & Chr(9) & var_50
  loc_1107E4A4:     call edi
  loc_1107E507:     var_2C8 = var_2C
  loc_1107E548:     var_2D0 = ADODB.Recordset.Fields
  loc_1107E57E:     ADODB.Recordset.8 = Forms
  loc_1107E5A2:     var_B8 = 0
  loc_1107E5AC:     var_E4 = var_B8
  loc_1107E601:     var_8384 = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_1107E60C:     call edi(var_1B0, var_B8)
  loc_1107E689:     var_2C8 = var_2C
  loc_1107E6D4:     var_2D0 = ADODB.Recordset.Fields
  loc_1107E700:     ADODB.Recordset.8 = Forms
  loc_1107E72B:     var_B8 = 0
  loc_1107E735:     var_E4 = var_B8
  loc_1107E783:     var_8390 = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_1107E78E:     call edi(var_1B0, var_B8)
  loc_1107E825:     var_8394 = var_58 & Chr(9) & 1100C008h
  loc_1107E830:     call edi
  loc_1107E8AD:     var_8398 = var_58 & Chr(9) & 1100AE28h
  loc_1107E8B8:     call edi
  loc_1107E935:     var_839C = var_58 & Chr(9) & 1100AE28h
  loc_1107E940:     call edi
  loc_1107E9BD:     var_83A0 = var_58 & Chr(9) & 1100AE28h
  loc_1107E9C8:     call edi
  loc_1107EA45:     var_83A4 = var_58 & Chr(9) & 1100AE28h
  loc_1107EA50:     call edi
  loc_1107EACD:     var_83A8 = var_58 & Chr(9) & 1100AE28h
  loc_1107EAD8:     call edi
  loc_1107EB55:     var_83AC = var_58 & Chr(9) & 1100AE28h
  loc_1107EB60:     call edi
  loc_1107EBDD:     var_83B0 = var_58 & Chr(9) & 1100AE28h
  loc_1107EBE8:     call edi
  loc_1107EC4B:     var_2C8 = var_2C
  loc_1107EC8C:     var_2D0 = ADODB.Recordset.Fields
  loc_1107ECC2:     ADODB.Recordset.8 = Forms
  loc_1107ECE6:     var_B8 = 0
  loc_1107ECF0:     var_E4 = var_B8
  loc_1107ED45:     var_83BC = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_1107ED50:     call edi(var_1B0, var_B8)
  loc_1107EDE7:     var_83C0 = var_58 & Chr(9) & 1100AE28h
  loc_1107EDF2:     call edi
  loc_1107EE6F:     var_83C4 = var_58 & Chr(9) & 1100AE28h
  loc_1107EE7A:     call edi
  loc_1107EEF7:     var_83C8 = var_58 & Chr(9) & 1100AE28h
  loc_1107EF02:     call edi
  loc_1107EF7F:     var_83CC = var_58 & Chr(9) & 1100AE28h
  loc_1107EF8A:     call edi
  loc_1107F007:     var_83D0 = var_58 & Chr(9) & 1100AE28h
  loc_1107F012:     call edi
  loc_1107F053:     var_2C8 = var_2C
  loc_1107F09E:     var_2D0 = ADODB.Recordset.Fields
  loc_1107F0CA:     ADODB.Recordset.8 = Forms
  loc_1107F0F5:     var_B8 = 0
  loc_1107F0FF:     var_C4 = var_B8
  loc_1107F10F:     var_83D8 = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_1107F11C:     call edi(var_1A0, var_B8)
  loc_1107F134:     call edi
  loc_1107F146:     var_20 = (8 + var_20)
  loc_1107F19A:     var_2C8 = var_2C
  loc_1107F1A0:     var_83E0 = ADODB.Recordset.MoveNext
  loc_1107F216:     frmGzToPzTGZS.VFG.DispID_0080(var_58)
  loc_1107F22B:     GoTo loc_1107E16C
  loc_1107F230:   End If
  loc_1107F253:   var_2C8 = var_2C
  loc_1107F27E:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_1107F28C:     var_58 = "3"
  loc_1107F2F1:     var_83E8 = var_58 & Chr(9) & 1100AE28h
  loc_1107F2FC:     call edi
  loc_1107F3AB:     var_83F0 = var_58 & Chr(9) & frmGzToPzTGZS.TDBDate.DispID_004E
  loc_1107F3B6:     call edi
  loc_1107F44D:     var_83F4 = var_58 & Chr(9) & 1100D6D4h
  loc_1107F458:     call edi
  loc_1107F4D5:     var_83F8 = var_58 & Chr(9) & 1100C008h
  loc_1107F4E0:     call edi
  loc_1107F55C:     var_83FC = var_58 & Chr(9) & var_50
  loc_1107F567:     call edi
  loc_1107F5E4:     var_8400 = var_58 & Chr(9) & "217601"
  loc_1107F5EF:     call edi
  loc_1107F66C:     var_8404 = var_58 & Chr(9) & 1100C008h
  loc_1107F677:     call edi
  loc_1107F6C9:     var_1B0 = var_1C
  loc_1107F6FC:     var_8408 = var_58 & Chr(9) & var_20
  loc_1107F707:     call edi
  loc_1107F784:     var_840C = var_58 & Chr(9) & 1100AE28h
  loc_1107F78F:     call edi
  loc_1107F80C:     var_8410 = var_58 & Chr(9) & 1100AE28h
  loc_1107F817:     call edi
  loc_1107F894:     var_8414 = var_58 & Chr(9) & 1100AE28h
  loc_1107F89F:     call edi
  loc_1107F91C:     var_8418 = var_58 & Chr(9) & 1100AE28h
  loc_1107F927:     call edi
  loc_1107F9A4:     var_841C = var_58 & Chr(9) & 1100AE28h
  loc_1107F9AF:     call edi
  loc_1107FA2C:     var_8420 = var_58 & Chr(9) & 1100AE28h
  loc_1107FA37:     call edi
  loc_1107FAB4:     var_8424 = var_58 & Chr(9) & 1100AE28h
  loc_1107FABF:     call edi
  loc_1107FB3C:     var_8428 = var_58 & Chr(9) & 1100AE28h
  loc_1107FB47:     call edi
  loc_1107FBC4:     var_842C = var_58 & Chr(9) & 1100AE28h
  loc_1107FBCF:     call edi
  loc_1107FC4C:     var_8430 = var_58 & Chr(9) & 1100AE28h
  loc_1107FC57:     call edi
  loc_1107FCD4:     var_8434 = var_58 & Chr(9) & 1100AE28h
  loc_1107FCDF:     call edi
  loc_1107FD5C:     var_8438 = var_58 & Chr(9) & 1100AE28h
  loc_1107FD67:     call edi
  loc_1107FDE4:     var_843C = var_58 & Chr(9) & 1100AE28h
  loc_1107FDEF:     call edi
  loc_1107FE59:     frmGzToPzTGZS.VFG.DispID_0080(var_58)
  loc_1107FE6E:   End If
  loc_1107FE9E:   var_1C4 = var_78
  loc_1107FEB7:   var_1A4 = "计提"
  loc_1107FF26:   var_8440 =  & "cIYear".0 & 1100D700h & var_78 & "月奖金"
  loc_1107FF31:   call edi
  loc_1107FF80:   var_2C8 = var_2C
  loc_1107FF86:   var_2C4 = ADODB.Recordset.State
  loc_1107FFB1:   If var_2C4 = 1 Then
  loc_1107FFCF:     var_2C8 = var_2C
  loc_1107FFD5:     var_844C = ADODB.Recordset.Close
  loc_1107FFF9:   End If
  loc_1108006F:   var_2C8 = var_2C
  loc_110800AC:   var_8454 = ADODB.Recordset.Open("SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZS_Temp] WHERE cGzItem='奖金' GROUP BY ccode,cdepcode", var_1A8, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZS_Temp] WHERE cGzItem='奖金' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_110800F3:   var_2C8 = var_2C
  loc_110800F9:   var_2C0 = ADODB.Recordset.EOF
  loc_1108011F:   If var_2C0 = 0 Then
  loc_1108012D:     var_58 = "4"
  loc_11080192:     var_845C = var_58 & Chr(9) & var_1B4
  loc_1108019D:     call edi(var_1B8, var_1B4, var_1B0, 00000003h, 00000001h, FFFFFFFFh)
  loc_1108024C:     var_8464 = var_58 & Chr(9) & frmGzToPzTGZS.TDBDate.DispID_004E
  loc_11080257:     call edi
  loc_110802EE:     var_8468 = var_58 & Chr(9) & 1100D6D4h
  loc_110802F9:     call edi
  loc_11080376:     var_846C = var_58 & Chr(9) & 1100C008h
  loc_11080381:     call edi
  loc_110803FD:     var_8470 = var_58 & Chr(9) & var_50
  loc_11080408:     call edi
  loc_1108046B:     var_2C8 = var_2C
  loc_110804AC:     var_2D0 = ADODB.Recordset.Fields
  loc_110804E2:     ADODB.Recordset.8 = Forms
  loc_11080506:     var_B8 = 0
  loc_11080510:     var_E4 = var_B8
  loc_11080565:     var_847C = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_11080570:     call edi(var_1B0, var_B8)
  loc_110805ED:     var_2C8 = var_2C
  loc_11080638:     var_2D0 = ADODB.Recordset.Fields
  loc_11080664:     ADODB.Recordset.8 = Forms
  loc_1108068F:     var_B8 = 0
  loc_11080699:     var_E4 = var_B8
  loc_110806E7:     var_8488 = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_110806F2:     call edi(var_1B0, var_B8)
  loc_11080789:     var_848C = var_58 & Chr(9) & 1100C008h
  loc_11080794:     call edi
  loc_11080811:     var_8490 = var_58 & Chr(9) & 1100AE28h
  loc_1108081C:     call edi
  loc_11080899:     var_8494 = var_58 & Chr(9) & 1100AE28h
  loc_110808A4:     call edi
  loc_11080921:     var_8498 = var_58 & Chr(9) & 1100AE28h
  loc_1108092C:     call edi
  loc_110809A9:     var_849C = var_58 & Chr(9) & 1100AE28h
  loc_110809B4:     call edi
  loc_11080A31:     var_84A0 = var_58 & Chr(9) & 1100AE28h
  loc_11080A3C:     call edi
  loc_11080AB9:     var_84A4 = var_58 & Chr(9) & 1100AE28h
  loc_11080AC4:     call edi
  loc_11080B41:     var_84A8 = var_58 & Chr(9) & 1100AE28h
  loc_11080B4C:     call edi
  loc_11080BAF:     var_2C8 = var_2C
  loc_11080BF0:     var_2D0 = ADODB.Recordset.Fields
  loc_11080C26:     ADODB.Recordset.8 = Forms
  loc_11080C4A:     var_B8 = 0
  loc_11080C54:     var_E4 = var_B8
  loc_11080CA9:     var_84B4 = var_58 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_11080CB4:     call edi(var_1B0, var_B8)
  loc_11080D4B:     var_84B8 = var_58 & Chr(9) & 1100AE28h
  loc_11080D56:     call edi
  loc_11080DD3:     var_84BC = var_58 & Chr(9) & 1100AE28h
  loc_11080DDE:     call edi
  loc_11080E5B:     var_84C0 = var_58 & Chr(9) & 1100AE28h
  loc_11080E66:     call edi
  loc_11080EE3:     var_84C4 = var_58 & Chr(9) & 1100AE28h
  loc_11080EEE:     call edi
  loc_11080F6B:     var_84C8 = var_58 & Chr(9) & 1100AE28h
  loc_11080F76:     call edi
  loc_11080FB7:     var_2C8 = var_2C
  loc_11081002:     var_2D0 = ADODB.Recordset.Fields
  loc_1108102E:     ADODB.Recordset.8 = Forms
  loc_11081059:     var_B8 = 0
  loc_11081063:     var_C4 = var_B8
  loc_11081073:     var_84D0 = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_11081080:     call edi(var_1A0, var_B8)
  loc_11081098:     call edi
  loc_110810AA:     var_20 = (8 + var_20)
  loc_110810FE:     var_2C8 = var_2C
  loc_11081104:     var_84D8 = ADODB.Recordset.MoveNext
  loc_1108117A:     frmGzToPzTGZS.VFG.DispID_0080(var_58)
  loc_1108118F:     GoTo loc_110800D0
  loc_11081194:   End If
  loc_110811B7:   var_2C8 = var_2C
  loc_110811E2:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_110811F0:     var_58 = "4"
  loc_11081255:     var_84E0 = var_58 & Chr(9) & 1100AE28h
  loc_11081260:     call edi
  loc_1108130F:     var_84E8 = var_58 & Chr(9) & frmGzToPzTGZS.TDBDate.DispID_004E
  loc_1108131A:     call edi
  loc_110813B1:     var_84EC = var_58 & Chr(9) & 1100D6D4h
  loc_110813BC:     call edi
  loc_11081439:     var_84F0 = var_58 & Chr(9) & 1100C008h
  loc_11081444:     call edi
  loc_110814C0:     var_84F4 = var_58 & Chr(9) & var_50
  loc_110814CB:     call edi
  loc_11081548:     var_84F8 = var_58 & Chr(9) & "215107"
  loc_11081553:     call edi
  loc_110815D0:     var_84FC = var_58 & Chr(9) & 1100C008h
  loc_110815DB:     call edi
  loc_1108162D:     var_1B0 = var_1C
  loc_11081660:     var_8500 = var_58 & Chr(9) & var_20
  loc_1108166B:     call edi
  loc_110816E8:     var_8504 = var_58 & Chr(9) & 1100AE28h
  loc_110816F3:     call edi
  loc_11081770:     var_8508 = var_58 & Chr(9) & 1100AE28h
  loc_1108177B:     call edi
  loc_110817F8:     var_850C = var_58 & Chr(9) & 1100AE28h
  loc_11081803:     call edi
  loc_11081880:     var_8510 = var_58 & Chr(9) & 1100AE28h
  loc_1108188B:     call edi
  loc_11081908:     var_8514 = var_58 & Chr(9) & 1100AE28h
  loc_11081913:     call edi
  loc_11081990:     var_8518 = var_58 & Chr(9) & 1100AE28h
  loc_1108199B:     call edi
  loc_11081A18:     var_851C = var_58 & Chr(9) & 1100AE28h
  loc_11081A23:     call edi
  loc_11081AA0:     var_8520 = var_58 & Chr(9) & 1100AE28h
  loc_11081AAB:     call edi
  loc_11081B28:     var_8524 = var_58 & Chr(9) & 1100AE28h
  loc_11081B33:     call edi
  loc_11081BB0:     var_8528 = var_58 & Chr(9) & 1100AE28h
  loc_11081BBB:     call edi
  loc_11081C38:     var_852C = var_58 & Chr(9) & 1100AE28h
  loc_11081C43:     call edi
  loc_11081CC0:     var_8530 = var_58 & Chr(9) & 1100AE28h
  loc_11081CCB:     call edi
  loc_11081D48:     var_8534 = var_58 & Chr(9) & 1100AE28h
  loc_11081D53:     call edi
  loc_11081DBD:     frmGzToPzTGZS.VFG.DispID_0080(var_58)
  loc_11081DD2:   End If
  loc_11081DEC:   var_8538 = CStr(var_28)
  loc_11081DFA:   call edi("有效数据共")
  loc_11081E03:   var_853C =  & edi("有效数据共")
  loc_11081E0D:   call edi
  loc_11081E90:   frmGzToPzTGZS.sBar.DispID_6803001E(8 & global_1100FE7C)
  loc_11081EFB:   frmGzToPzTGZS.APB.UnkVCall_00000040h
  loc_11081F8D:   Set var_B4 = frmGzToPzTGZS.APB
  loc_11081F9B:   var_2C8 = var_B4
  loc_11081FA1:   var_B4.UnkVCall_00000040h
  loc_11082033:   Set var_B4 = frmGzToPzTGZS.APB
  loc_11082041:   var_2C8 = var_B4
  loc_11082047:   var_B4.UnkVCall_00000040h
  loc_110820F7:   frmGzToPzTGZS.Pic1.DispID_80010007 = var_1A4
  loc_11082183:   var_C4 = frmGzToPzTGZS.TDBText
  loc_110821D0:   var_1AC = var_5C.UnkVCall_0000006Ch
  loc_11082209:   var_1A8 = var_68.UnkVCall_00000398h
  loc_1108223E:   Set var_3C = {000208D7-0000-0000-C000000000000046}()
  loc_1108224E:   Set var_5C = {000208DA-0000-0000-C000000000000046}()
  loc_1108225E:   Set var_68 = {000208D5-0000-0000-C000000000000046}()
  loc_1108226F: End If
  loc_11082275: GoTo loc_11082345
  loc_11082344: Exit Function
  loc_11082345: ' Referenced from: 11082275
End Function

Public Function getWBHL(sWhere) '110937B0
  Dim var_1C As ADODB.Recordset
  Dim var_2C As Me
  loc_11093810: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1109381C: var_98 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11093844: var_40 = Trim(sWhere)
  loc_11093875: If (var_40 <> 1100AE28h) Then
  loc_110938A3:   var_20 = "SELECT * FROM exch WHERE 1=1 " & " AND " & sWhere
  loc_110938B0: Else
  loc_110938BC: End If
  loc_110938CC: var_20 = var_20 & " order by cexch_name, itype, iperiod, cdate"
  loc_11093936: var_78 = var_1C
  loc_11093945: var_8018 = ADODB.Recordset.Open(var_20, var_5C, var_20, var_54, 9)
  loc_110939AB: If ADODB.Recordset.EOF Then
  loc_110939BA:   var_24 = CStr(0)
  loc_110939C5: Else
  loc_110939E7:   var_2C = ADODB.Recordset.Fields
  loc_11093A14:   var_58 = "NFLAT"
  loc_11093A2D:   ADODB.Recordset.8 = Forms
  loc_11093A7E:   var_24 = var_40
  loc_11093AA0: End If
  loc_11093ABE: var_8030 = ADODB.Recordset.Close
  loc_11093ADD: GoTo loc_11093B1B
  loc_11093AE3: If var_4 Then
  loc_11093AEE: End If
  loc_11093B1A: Exit Function
  loc_11093B1B: ' Referenced from: 11093ADD
End Function

Public Function GetKmCode(pGZ_Item, pGZ_Type, pGZ_Type1) '11094DF0
  Dim var_34 As ADODB.Recordset
  loc_11094E7D: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11094E85: var_D0 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11094EB0: On Error GoTo loc_1109520C
  loc_11094EB9: var_78 = pGZ_Item
  loc_11094EC7: var_88 = pGZ_Type
  loc_11094ED6: var_98 = pGZ_Type1
  loc_11094F44: var_8018 = Proc_0_10_11028DD0(&H4008,  & Proc_0_10_11028DD0(var_80, 1 & "SELECT * FROM " & "..T_CY_GZ_TGZS_KmSetting WHERE GZ_Item=", ) & " AND GZ_Type=", )
  loc_11095046: var_8030 = ADODB.Recordset.Open(fs:[00000000h] & Proc_0_10_11028DD0(&H4008,  & var_8018 & " AND GZ_Type1=", ), var_7C, fs:[00000000h] & Proc_0_10_11028DD0(&H4008,  & var_8018 & " AND GZ_Type1=", ), var_74, 9)
  loc_1109508F: var_A4 = ADODB.Recordset.EOF
  loc_110950B1: If var_A4 = 0 Then
  loc_110950F1:   var_B0 = ADODB.Recordset.Fields
  loc_110950F7:   var_78 = "UF_KMCode"
  loc_11095126:   ADODB.Recordset.8 = Forms
  loc_11095177:   var_2C = var_70
  loc_110951BD:   If ADODB.Recordset.Close < 0 Then
  loc_110951CB:     var_8044 = CheckObj(var_34, global_1100ADFC, 128)
  loc_110951D3:   End If
  loc_110951EE:   var_804C = ADODB.Recordset.Close
  loc_1109521E:   Set var_34 = ADODB.Recordset()
  loc_11095232: End If
  loc_11095232: Exit Sub
  loc_1109523D: GoTo loc_11095297
  loc_11095243: If var_C Then
  loc_1109524E: End If
  loc_11095296: Exit Function
  loc_11095297: ' Referenced from: 1109523D
End Function

Public Function GetUFDepCode(pDepCode, pType) '110952F0
  Dim var_30 As ADODB.Recordset
  Dim var_4C As Me
  loc_11095368: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11095370: var_B0 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11095382: var_2C = pDepCode
  loc_11095393: On Error GoTo loc_11095693
  loc_1109539C: var_68 = var_2C
  loc_110953AA: var_78 = pType
  loc_1109540F: var_8018 = Proc_0_10_11028DD0(var_80,  & Proc_0_10_11028DD0(&H4008, 1 & "SELECT * FROM " & "..T_CY_GZ_SL_DepSetting WHERE GZ_Dep=", ) & " AND GZ_Type=", )
  loc_11095423: var_24 =  & var_8018
  loc_110954C3: var_8024 = ADODB.Recordset.Open(var_24, var_6C, var_24, var_64, 9)
  loc_11095520: var_84 = ADODB.Recordset.EOF
  loc_1109553C: If var_84 = 0 Then
  loc_11095560:   var_4C = ADODB.Recordset.Fields
  loc_1109557E:   var_68 = "UF_DepCode"
  loc_110955A6:   ADODB.Recordset.8 = Forms
  loc_110955F7:   var_28 = var_60
  loc_11095641:   If ADODB.Recordset.Close < 0 Then
  loc_1109564F:     var_8038 = CheckObj(var_30, global_1100ADFC, 128)
  loc_11095653:   End If
  loc_11095659:   var_28 = var_2C
  loc_11095683:   If ADODB.Recordset.Close < 0 Then
  loc_11095693:   End If
  loc_110956A5:   Set var_30 = ADODB.Recordset()
  loc_110956B1:   var_28 = var_2C
  loc_110956B7: End If
  loc_110956B7: Exit Sub
  loc_110956C2: GoTo loc_11095710
  loc_110956C8: If var_C Then
  loc_110956D3: End If
  loc_1109570F: Exit Function
  loc_11095710: ' Referenced from: 110956C2
End Function

Public Function getBTData() '11095760
  Dim var_24 As ADODB.Recordset
  Dim var_38 As Variant
  loc_110957E4: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110957EE: On Error GoTo loc_11095D77
  loc_11095829: var_28 = 1 & "IF NOT EXISTS (SELECT * FROM [" & "]..Sysobjects "
  loc_11095894: var_8018 =  & var_28 & "WHERE Name = 'T_CY_GZ_ZGZS_Setting') " & "CREATE TABLE [" & "]..[T_CY_GZ_ZGZS_Setting](fXS1 FLOAT NULL," & "fXS2 FLOAT NULL)"
  loc_1109589B: var_28 = var_8018
  loc_110958CB: var_54 = UnkObj.UnkVCall_00000040h
  loc_1109591D: var_28 = var_38 & "SELECT * FROM [" & "]..[T_CY_GZ_ZGZS_Setting]"
  loc_11095957: var_BC = ADODB.Recordset.State
  loc_1109597C: If var_BC = 1 Then
  loc_11095998:   var_802C = ADODB.Recordset.Close
  loc_110959B6: End If
  loc_11095A36: var_8034 = ADODB.Recordset.Open(var_28, var_90, var_28, var_88, 9)
  loc_11095A89: var_B8 = ADODB.Recordset.EOF
  loc_11095AA5: If var_B8 = 0 Then
  loc_11095ACD:   var_38 = ADODB.Recordset.Fields
  loc_11095AEB:   var_8C = "fXS1"
  loc_11095B1F:   ADODB.Recordset.8 = Forms
  loc_11095B8A:   frmGzToPzTGZS.TDBNum.UnkVCall_00000040h
  loc_11095BC0:   var_44.DispID_0000 = Proc_0_12_110291B0(9, var_90, "fXS1")
  loc_11095C26:   var_D0 = ADODB.Recordset.Fields
  loc_11095C31:   var_8C = "fXS1"
  loc_11095C65:   ADODB.Recordset.8 = Forms
  loc_11095CD3:   frmGzToPzTGZS.TDBNum.UnkVCall_00000040h
  loc_11095D09:   var_44.DispID_0000 = Proc_0_12_110291B0(9, var_90, "fXS1")
  loc_11095D36: End If
  loc_11095D5E: If ADODB.Recordset.Close < 0 Then
  loc_11095D70:   var_8050 = CheckObj(var_24, global_1100ADFC, 128)
  loc_11095D77:   ' Referenced from: 110957EE
  loc_11095D7C:   var_8054 = Err
  loc_11095D87:   Set var_38 = Err
  loc_11095E0C:   MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_11095E39: End If
  loc_11095E39: Exit Sub
  loc_11095E44: GoTo loc_11095E8D
  loc_11095E8C: Exit Function
  loc_11095E8D: ' Referenced from: 11095E44
End Function

Public Function UpdateBTData() '11095ED0
  Dim var_3C As Variant
  Dim var_44 As frmGzToPzTGZS.TDBNum
  Dim var_20 As Me
  loc_11095F48: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11095F52: On Error GoTo loc_1109621B
  loc_11095F8D: var_20 = 1 & "DELETE FROM [" & "]..[T_CY_GZ_ZGZS_Setting]"
  loc_11095FC6: var_58 = UnkObj.UnkVCall_00000040h
  loc_11096046: Set var_3C = frmGzToPzTGZS.TDBNum
  loc_1109604C: var_BC = var_3C
  loc_1109605B: var_3C.UnkVCall_00000040h
  loc_11096095: var_60 = var_40.DispID_0043
  loc_110960B0: Set var_44 = frmGzToPzTGZS.TDBNum
  loc_110960B6: var_C4 = var_44
  loc_110960C5: var_44.UnkVCall_00000040h
  loc_110960FF: var_80 = var_48.DispID_0043
  loc_11096130: var_8028 = 1 & Proc_0_12_110291B0(8, var_3C & "INSERT INTO [" & "]..[T_CY_GZ_ZGZS_Setting]" & "(fXS1,fXS2) VALUES (", var_44) & global_1100AC40
  loc_11096164: var_20 = var_3C & Proc_0_12_110291B0(8, var_8028, var_48) & global_1100BD88
  loc_11096216: GoTo loc_110962DD
  loc_1109621B: ' Referenced from: 11095F52
  loc_11096220: var_8038 = Err
  loc_1109622B: Set var_3C = Err
  loc_110962B0: MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_110962DD: ' Referenced from: 11096216
  loc_110962DD: Exit Sub
  loc_110962E8: GoTo loc_1109633D
  loc_1109633C: Exit Function
  loc_1109633D: ' Referenced from: 110962E8
End Function

Private Sub Proc_13_11_11074C60
  Dim var_58 As frmGzToPzTGZS.VFG
  loc_11074CA1: Set var_58 = frmGzToPzTGZS.VFG
  loc_11074CF2: var_58.DispID_005D = frmGzToPzTGZS.VFG
  loc_11074D33: var_58.DispID_0067 = frmGzToPzTGZS.VFG
  loc_11074D52: var_58.DispID_0041 = frmGzToPzTGZS.VFG
  loc_11074DFC: var_58.DispID_00A5("...")
  loc_11074F24: var_58.DispID_008A(4)
  loc_11074F67: var_58.DispID_0079(450)
  loc_11074F8B: var_58.DispID_0019 = True
  loc_11074FCB: var_58.DispID_007B(True)
  loc_11075010: var_58.DispID_0090("业务号")
  loc_11075053: var_58.DispID_0077(4)
  loc_11075096: var_58.DispID_0078(700)
  loc_110750DE: var_58.DispID_0090("状态")
  loc_11075124: var_58.DispID_0077(4)
  loc_1107516A: var_58.DispID_0078(700)
  loc_110751B2: var_58.DispID_0090("制单日期")
  loc_110751F8: var_58.DispID_0077(1)
  loc_1107523E: var_58.DispID_0078(1000)
  loc_11075283: var_58.DispID_0090("凭证类别字")
  loc_110752C5: var_58.DispID_0077(4)
  loc_11075307: var_58.DispID_0078(700)
  loc_1107534F: var_58.DispID_0090("附单据数")
  loc_11075393: var_58.DispID_0077(var_3C)
  loc_110753D9: var_58.DispID_0078(var_3C)
  loc_11075421: var_58.DispID_0090(var_3C)
  loc_11075467: var_58.DispID_0077(var_3C)
  loc_110754AD: var_58.DispID_0078(var_3C)
  loc_110754F5: var_58.DispID_0090(var_3C)
  loc_1107553B: var_58.DispID_0077(var_3C)
  loc_11075581: var_58.DispID_0078(var_3C)
  loc_110755C9: var_58.DispID_0090(var_3C)
  loc_1107560D: var_58.DispID_0077(var_3C)
  loc_11075653: var_58.DispID_0078(var_3C)
  loc_1107569B: var_58.DispID_009C(var_3C)
  loc_110756E3: var_58.DispID_0090(var_3C)
  loc_11075729: var_58.DispID_0077(var_3C)
  loc_1107576F: var_58.DispID_0078(var_3C)
  loc_110757B7: var_58.DispID_009C(var_3C)
  loc_110757FF: var_58.DispID_0090(var_3C)
  loc_11075845: var_58.DispID_0077(var_3C)
  loc_1107588B: var_58.DispID_0078(var_3C)
  loc_110758D3: var_58.DispID_009C(var_3C)
  loc_1107591B: var_58.DispID_0090(var_3C)
  loc_11075961: var_58.DispID_0077(var_3C)
  loc_110759A7: var_58.DispID_0078(var_3C)
  loc_110759EF: var_58.DispID_009C(var_3C)
  loc_11075A37: var_58.DispID_0090(var_3C)
  loc_11075A7D: var_58.DispID_0077(var_3C)
  loc_11075AC3: var_58.DispID_0078(var_3C)
  loc_11075B0B: var_58.DispID_009C(var_3C)
  loc_11075B53: var_58.DispID_0090(var_3C)
  loc_11075B99: var_58.DispID_0077(var_3C)
  loc_11075BDF: var_58.DispID_0078(var_3C)
  loc_11075C27: var_58.DispID_0090(var_3C)
  loc_11075C6D: var_58.DispID_0077(var_3C)
  loc_11075CB3: var_58.DispID_0078(var_3C)
  loc_11075CFB: var_58.DispID_0090(var_3C)
  loc_11075D41: var_58.DispID_0077(var_3C)
  loc_11075D87: var_58.DispID_0078(var_3C)
  loc_11075DCF: var_58.DispID_0090(var_3C)
  loc_11075E15: var_58.DispID_0077(var_3C)
  loc_11075E5B: var_58.DispID_0078(var_3C)
  loc_11075EA3: var_58.DispID_0090(var_3C)
  loc_11075EE9: var_58.DispID_0077(var_3C)
  loc_11075F2F: var_58.DispID_0078(var_3C)
  loc_11075F77: var_58.DispID_0090(var_3C)
  loc_11075FBD: var_58.DispID_0077(var_3C)
  loc_11076003: var_58.DispID_0078(var_3C)
  loc_1107604B: var_58.DispID_0090(var_3C)
  loc_11076091: var_58.DispID_0077(var_3C)
  loc_110760D7: var_58.DispID_0078(var_3C)
  loc_1107611F: var_58.DispID_0090(var_3C)
  loc_11076165: var_58.DispID_0077(var_3C)
  loc_110761AB: var_58.DispID_0078(var_3C)
  loc_110761F3: var_58.DispID_0090(var_3C)
  loc_11076239: var_58.DispID_0077(var_3C)
  loc_1107627F: var_58.DispID_0078(var_3C)
  loc_110762C7: var_58.DispID_0090(var_3C)
  loc_1107630D: var_58.DispID_0077(var_3C)
  loc_11076353: var_58.DispID_0078(var_3C)
  loc_1107639B: var_58.DispID_0090(var_3C)
  loc_110763E1: var_58.DispID_0077(var_3C)
  loc_11076427: var_58.DispID_0078(var_3C)
  loc_11076443: If 9 <= &H15 Then
  loc_11076483:   var_58.DispID_00AC(var_3C)
  loc_1107649B:   var_14 = 1+var_14
  loc_1107649E:   GoTo loc_1107643F
  loc_110764A0: End If
  loc_110764DF: var_58.DispID_00AC(var_3C)
  loc_11076525: var_58.DispID_00AC(var_3C)
End Sub

Private Sub Proc_13_12_11082B50
  Dim var_7C As Variant
  Dim var_1F8 As Label
  Dim var_80 As Variant
  Dim var_88 As frmGzToPzTGZS.Label3
  loc_11082C3A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11082C42: var_228 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11082C48: var_8004 = ecx
  loc_11082CBE: If var_14 <= CLng(frmGzToPzTGZS.VFG.DispID_0007)(-1) Then
  loc_11082CCF:   var_800C = frmGzToPzTGZS.Proc_13_13_110849F0(vbNull)
  loc_11082D6D:   frmGzToPzTGZS.VFG.DispID_0082(22, var_58)
  loc_11082E51:   If (frmGzToPzTGZS.VFG.DispID_0082(var_14, 22) = global_1100AE28) + 1 Then
  loc_11082ED1:     frmGzToPzTGZS.VFG.DispID_0082(1, 285267764)
  loc_11083005:     frmGzToPzTGZS.VFG.DispID_009E(var_14, 1, var_14, 1, 16711680)
  loc_11083025:     Set var_7C = frmGzToPzTGZS.Label3
  loc_11083032:     var_1F8 = var_7C
  loc_1108307C:     var_7C.Caption = "分析: 第(" & CStr(vbNull) & ")行信息----有效"
  loc_110830CE:     frmGzToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_110830E1:   Else
  loc_1108315B:     frmGzToPzTGZS.VFG.DispID_0082(1, 285267820)
  loc_1108328F:     frmGzToPzTGZS.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110832AF:     Set var_80 = frmGzToPzTGZS.Label3
  loc_110832BC:     var_1F8 = var_80
  loc_1108339D:     var_80.Caption = "分析:   第(" & CStr(vbNull) & ")行信息----" & frmGzToPzTGZS.VFG.DispID_0082(var_14, 22)
  loc_11083408:     frmGzToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_1108341A:   End If
  loc_1108342A:   var_14 = 1+var_14
  loc_1108342D:   GoTo loc_11082CB0
  loc_11083432: End If
  loc_11083499: If var_14 <= CLng(frmGzToPzTGZS.VFG.DispID_0007)(-1) Then
  loc_11083511:   var_A0 = frmGzToPzTGZS.VFG.DispID_0082(var_14, 2)
  loc_1108352F:   var_B8)
  loc_110836BF:   var_8048 = frmGzToPzTGZS.VFG.DispID_0082(var_14, frmGzToPzTGZS.VFG)
  loc_110836F6:   var_4C = CCur(0)
  loc_110836F9:   var_48 = var_8048
  loc_11083705:   var_40 = CCur(0)
  loc_11083708:   var_3C = var_8048
  loc_11083714:   var_34 = var_14
  loc_1108371D:   var_30 = var_14
  loc_11083726:   var_160 = CByte("DateToPeriod".00000001h)
  loc_110837C3:   var_B8)
  loc_11083842:   Set var_80 = frmGzToPzTGZS.VFG
  loc_11083868:   var_8064 = (frmGzToPzTGZS.VFG.DispID_0082(var_14, 3) = var_80.DispID_0082(var_14, 3))
  loc_11083895:   var_1A0 = var_8064 + 1
  loc_1108390F:   var_806C = (var_8048 = frmGzToPzTGZS.VFG.DispID_0082(var_14, ""))
  loc_11083936:   var_1E0 = var_806C + 1
  loc_11083A38:   If CBool((frmGzToPzTGZS.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_8064 + 1 And var_806C + 1) Then
  loc_11083AFD:     If (frmGzToPzTGZS.VFG.DispID_0082(var_14, 22) = global_1100AE28) Then
  loc_11083B06:     End If
  loc_11083B0B:     If var_24 = 0 Then
  loc_11083BB4:       var_16C = var_48
  loc_11083BF8:       var_9C = var_1F0
  loc_11083C44:       var_4C = CCur(var_4C + Format(Val(frmGzToPzTGZS.VFG.DispID_0082(var_14, 7)), "#.00"))
  loc_11083C47:       var_48 = var_D8
  loc_11083D27:       var_16C = var_3C
  loc_11083D6B:       var_9C = var_1F0
  loc_11083DB7:       var_40 = CCur(var_40 + Format(Val(frmGzToPzTGZS.VFG.DispID_0082(var_14, 8)), "#.00"))
  loc_11083DBA:       var_3C = var_D8
  loc_11083DFA:     End If
  loc_11083E1B:     var_14 = var_14(1)
  loc_11083E1E:     var_30 = var_30(1)
  loc_11083E40:     var_80A0 = CLng(frmGzToPzTGZS.VFG.DispID_0007)
  loc_11083E5B:     var_1F8 = (var_14 > 0)
  loc_11083E7F:     If var_1F8 = 0 Then GoTo loc_11083720
  loc_11083E85:   End If
  loc_11083E8A:   If var_24 = 0 Then
  loc_11083E9E:     Set var_7C = frmGzToPzTGZS.Chk
  loc_11083EA9:     var_1F8 = var_7C
  loc_11083EAF:     Set var_80 = var_7C(1)
  loc_11083EDA:     var_200 = var_80
  loc_11083EE0:     var_1EC = var_80.Value
  loc_11083F34:     If (var_1EC = 1) Then
  loc_11083F64:       If (Abs(var_4C - var_40) <> 0.01) >= 0 Then
  loc_11083F6D:       End If
  loc_11083F6D:     End If
  loc_11083F72:     If var_24 Then
  loc_11083F78:     End If
  loc_11083F98:     var_1C = var_34
  loc_11083F9D:     If var_34 <= (var_30 - 1) Then
  loc_11084061:       If (frmGzToPzTGZS.VFG.DispID_0082(var_1C, 22) = global_1100AE28) + 1 Then
  loc_110840E9:         frmGzToPzTGZS.VFG.DispID_0082(1, 285267820)
  loc_1108417D:         frmGzToPzTGZS.VFG.DispID_0082(22, "凭证借贷不平衡或某分录有错误")
  loc_110842B1:         frmGzToPzTGZS.VFG.DispID_009E(var_1C, 1, var_1C, 1, 255)
  loc_110842C3:       End If
  loc_110842D3:       GoTo loc_11083F92
  loc_110842D8:     End If
  loc_110842E9:     var_44 = var_44(1)
  loc_110842FA:     Set var_88 = frmGzToPzTGZS.Label3
  loc_1108432D:     var_1F8 = var_88
  loc_1108443A:     Set var_80 = frmGzToPzTGZS.VFG
  loc_11084512:     var_80D4 = "分析: 第[" & frmGzToPzTGZS.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmGzToPzTGZS.VFG.DispID_0082(var_34, var_14)
  loc_11084534:     var_78 = var_80D4 & "]号凭证借贷不平衡"
  loc_11084548:     var_88.Caption = var_78
  loc_1108454F:     If var_78 < 0 Then
  loc_11084555:       GoTo loc_110847D3
  loc_1108455A:     End If
  loc_1108456B:     var_20 = var_20(1)
  loc_1108457C:     Set var_88 = frmGzToPzTGZS.Label3
  loc_110845AF:     var_1F8 = var_88
  loc_110846BC:     Set var_80 = frmGzToPzTGZS.VFG
  loc_11084794:     var_80F8 = "分析: 第[" & frmGzToPzTGZS.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmGzToPzTGZS.VFG.DispID_0082(var_34, frmGzToPzTGZS.VFG.DispID_0082(var_34, var_14))
  loc_110847B6:     var_78 = var_80F8 & "]号凭证有效"
  loc_110847CA:     var_88.Caption = var_78
  loc_110847D1:     If var_78 >= 0 Then GoTo loc_110847E2
  loc_110847D3:     ' Referenced from: 11084555
  loc_110847DC:     var_78 = CheckObj(var_1F8, global_1100D574, 84)
  loc_110847E2:   End If
  loc_11084864:   frmGzToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_11084895:   var_14 = 1+var_14(-1)
  loc_11084898:   GoTo loc_11083493
  loc_1108489D: End If
  loc_110848A2: If var_44 > 0 Then
  loc_110848A9:   If var_20 > 0 Then
  loc_110848C4:   Else
  loc_110848DD:   Else
  loc_110848E7:     var_8108 = frmGzToPzTGZS.Proc_13_15_11093B60(var_1EC)
  loc_110848F5:     If var_1EC Then
  loc_11084910:     Else
  loc_11084918:       var_18 = ecx
  loc_11084921:       GoTo loc_110849BB
  loc_110849BA:       Exit Sub
  loc_110849BB:     End If
  loc_110849BB:   End If
  loc_110849BB: End If
  loc_110849BB: ' Referenced from: 11084921
End Sub

Private  Proc_13_13_110849F0(arg_C) '110849F0
  Dim var_58 As frmGzToPzTGZS.VFG
  Dim var_20 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA19-EB96-11D2-AF06000021009B21}()
  Dim var_18 As {3302AA41-EB96-11D2-AF06000021009B21}()
  Dim var_1C As {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_11084AEC: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11084AFC: var_210 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11084BDB: If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 2) = global_1100AE28) + 1 Then
  loc_11084BE5:   var_24 = "制单日期为空"
  loc_11084BF6: Else
  loc_11084C91:   var_78 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 2)
  loc_11084CCB:   If Proc_0_9_11028500(var_80, global_11089FFD, ) Then
  loc_11084D74:     var_78 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 2)
  loc_11084D7E:     var_90)
  loc_11084D90:     var_48 = var_90
  loc_11084DC2:     var_118 = var_48
  loc_11084DD0:     var_114 = var_44
  loc_11084E04:     var_80 = "AccountOpen".0.0
  loc_11084E35:     If (var_80 < var_80) Then
  loc_11084E3F:       var_24 = "日期超前总账系统启用日期"
  loc_11084E50:     Else
  loc_11084E56:       var_154 = var_44
  loc_11084E5C:       var_1A4 = var_44
  loc_11084E68:       var_158 = var_48
  loc_11084E6F:       var_1A8 = var_48
  loc_11084F1C:       var_80 = "AccountYMD".0.00000002h("AccountYMD".0, var_13C)
  loc_11085016:       If CBool( Or ((global_11089FFD < var_80) > "AccountYMD".0.00000002h(var_180, var_18C))) Then
  loc_11085020:         var_24 = "日期必须在当前会计年度内"
  loc_11085031:       Else
  loc_1108504E:         var_118 = var_48
  loc_110850A2:         var_80 = "DateToPeriod".00000001h - 1
  loc_11085130:         If CBool("AccountYMD".0.00000001h) Then
  loc_1108513A:           var_24 = "已结账月份不能制单"
  loc_1108514B:         Else
  loc_11085227:           If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 3) = global_1100AE28) + 1 Then
  loc_11085231:             var_24 = "凭证类别字为空"
  loc_11085242:           Else
  loc_110852D1:             var_8034 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 3)
  loc_110852E1:             var_80 = 8
  loc_110852E4:             var_78 = var_8034
  loc_1108532B:             var_8038 = CBool(Not("pzlbCheck".00000001h(, fs:[00000000h], , global_11089FFD, global_11089FFD, var_74, var_8034, var_7C)))
  loc_11085362:             If var_8038 Then
  loc_1108536C:               var_24 = "凭证类别字非法"
  loc_1108537D:             Else
  loc_11085454:               If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, var_128) = global_1100AE28) + 1 Then
  loc_1108545E:                 var_24 = "业务号为空"
  loc_1108546F:               Else
  loc_110854F9:                 var_8044 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, var_128)
  loc_11085509:                 var_80 = 8
  loc_1108550C:                 var_78 = var_8044
  loc_1108554F:                 var_90 = "GenLen".00000001h(fs:[00000000h], , global_11089FFD, global_11089FFD, global_11089FFD, var_74, var_8044, var_7C)
  loc_11085597:                 If (var_90 > 30) Then
  loc_110855A1:                   var_24 = "业务号超长"
  loc_110855B2:                 Else
  loc_11085691:                   If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 5) = global_1100AE28) + 1 Then
  loc_1108569B:                     var_24 = "摘要为空"
  loc_110856AC:                   Else
  loc_11085767:                     var_8058 = InStr(1, frmGzToPzTGZS.VFG.DispID_0082(arg_C, 5), "|", 0)
  loc_1108578D:                     var_220 = (var_8058 > 0)
  loc_110857E3:                     var_80 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 5)
  loc_11085904:                     If (((var_8058 > 0) Or (InStr(1, var_80, """", 0) > 0)) Or (InStr(1, frmGzToPzTGZS.VFG.DispID_0082(arg_C, 5), "'", 0) > 0)) Then
  loc_1108590E:                       var_24 = "摘要含有非法字符"
  loc_1108591F:                     Else
  loc_110859B1:                       var_806C = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 5)
  loc_110859C1:                       var_80 = 8
  loc_110859C4:                       var_78 = var_806C
  loc_11085A07:                       var_90 = "GenLen".00000001h(global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, var_74, var_806C, var_7C)
  loc_11085A50:                       If (var_90 > 120) Then
  loc_11085A5A:                         var_24 = "摘要超长"
  loc_11085A6B:                       Else
  loc_11085B48:                         If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 6) = global_1100AE28) + 1 Then
  loc_11085B52:                           var_24 = "科目为空"
  loc_11085B63:                         Else
  loc_11085BF2:                           var_807C = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 6)
  loc_11085C02:                           var_80 = 8
  loc_11085C05:                           var_78 = var_807C
  loc_11085C85:                           var_40 = "kmCheck".00000002h(var_807C, var_150, var_15C)
  loc_11085CB7:                           var_8084 = (var_40 = global_1100AE28)
  loc_11085CBF:                           If var_8084 = 0 Then
  loc_11085CC9:                             var_24 = "科目非法"
  loc_11085CDA:                           Else
  loc_11085D18:                             var_118 = arg_C
  loc_11085D7F:                             frmGzToPzTGZS.VFG.DispID_0082(6, var_40)
  loc_11085D99:                             var_118 = var_40
  loc_11085DEB:                             var_128 = var_20
  loc_11085E39:                             "kmCodeToProperties".00000002h
  loc_11085E56:                             Set var_20 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_11085E84:                             var_1F0 = var_20
  loc_11085E8A:                             var_1D4 = var_20.UnkVCall_00000114h
  loc_11085EB6:                             If var_1D4 = 0 Then
  loc_11085EC0:                               var_24 = "科目非末级"
  loc_11085ED1:                             Else
  loc_11085FAF:                               If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 7) = global_1100AE28) Then
  loc_1108608B:                                 If Not (IsNumeric(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 7))) Then
  loc_11086095:                                   var_24 = "借方金额非法"
  loc_110860A6:                                 Else
  loc_1108614F:                                   var_80A4 = CDbl(Val(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 7)))
  loc_110861EA:                                   var_80 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 7)
  loc_11086212:                                   var_22C = CDbl(Val(var_80))
  loc_11086228:                                   var_80B0 = CDbl(-9999999999999.99)
  loc_11086240:                                   GoTo loc_11086244
  loc_11086292:                                   If (eax Or 0) Then
  loc_1108629C:                                     var_24 = "借方金额超范围"
  loc_110862AD:                                   Else
  loc_110862AD:                                   End If
  loc_1108638B:                                   If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 8) = global_1100AE28) Then
  loc_11086467:                                     If Not (IsNumeric(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 8))) Then
  loc_11086471:                                       var_24 = "贷方金额非法"
  loc_11086482:                                     Else
  loc_1108652B:                                       var_80C8 = CDbl(Val(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 8)))
  loc_110865C6:                                       var_80 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 8)
  loc_110865EE:                                       var_238 = CDbl(Val(var_80))
  loc_11086604:                                       var_80D4 = CDbl(-9999999999999.99)
  loc_1108661C:                                       GoTo loc_11086620
  loc_1108666E:                                       If (eax Or 0) Then
  loc_11086678:                                         var_24 = "贷方金额超范围"
  loc_11086689:                                       Else
  loc_11086689:                                       End If
  loc_11086801:                                       var_74 = var_1E0
  loc_11086873:                                       var_C4 = var_1E8
  loc_110868ED:                                       var_80E8 = (Format(Val(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 7)), "#.00") <> 0) And (Format(Val(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 8)), "#.00") <> 0)
  loc_11086966:                                       If CBool(var_80E8) Then
  loc_11086970:                                         var_24 = "借方金额和贷方金额不能同时不为0"
  loc_11086981:                                       Else
  loc_11086AF9:                                         var_74 = var_1E0
  loc_11086B6B:                                         var_C4 = var_1E8
  loc_11086BE5:                                         var_8100 = (Format(Val(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 7)), "#.00") = 0) And (Format(Val(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 8)), "#.00") = 0)
  loc_11086C5E:                                         If CBool(var_8100) Then
  loc_11086C68:                                           var_24 = "借方金额和贷方金额不能同时为0"
  loc_11086C79:                                         Else
  loc_11086C99:                                           var_1F0 = var_20
  loc_11086CEB:                                           If (var_20.UnkVCall_0000007Ch = global_1100AE28) Then
  loc_11086DCF:                                             If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 9) = global_1100AE28) Then
  loc_11086EAB:                                               If Not (IsNumeric(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 9))) Then
  loc_11086EB5:                                                 var_24 = "数量数值非法"
  loc_11086EC6:                                               Else
  loc_11086EC6:                                               End If
  loc_11086EC6:                                             End If
  loc_11086EE6:                                             var_1F0 = var_20
  loc_11086F38:                                             If (var_20.UnkVCall_0000006Ch = global_1100AE28) Then
  loc_1108701C:                                               If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 10) = global_1100AE28) Then
  loc_110870F8:                                                 If Not (IsNumeric(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 10))) Then
  loc_11087102:                                                   var_24 = "外币金额非法"
  loc_11087113:                                                 Else
  loc_110871BC:                                                   var_813C = CDbl(Val(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 10)))
  loc_1108727F:                                                   var_244 = CDbl(Val(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 10)))
  loc_11087295:                                                   var_8148 = CDbl(-9999999999999.99)
  loc_110872AD:                                                   GoTo loc_110872B1
  loc_110872FF:                                                   If (eax Or 0) Then
  loc_11087309:                                                     var_24 = "外币超范围"
  loc_1108731A:                                                   Else
  loc_1108731A:                                                   End If
  loc_110873F8:                                                   If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 11) = global_1100AE28) Then
  loc_110874D4:                                                     If Not (IsNumeric(frmGzToPzTGZS.VFG.DispID_0082(arg_C, 11))) Then
  loc_110874DE:                                                       var_24 = "汇率数值非法"
  loc_110874EF:                                                     Else
  loc_110874EF:                                                     End If
  loc_110874EF:                                                   End If
  loc_110875CD:                                                   If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 12) = global_1100AE28) Then
  loc_11087664:                                                     var_8164 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 12)
  loc_11087677:                                                     var_78 = var_8164
  loc_110876BA:                                                     var_90 = "GenLen".00000001h(global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, var_74, var_8164, var_7C)
  loc_110876D4:                                                     var_1F0 = (var_90 > 20)
  loc_11087703:                                                     If var_1F0 = 0 Then GoTo loc_11087849
  loc_11087711:                                                     var_24 = "制单人姓名超长"
  loc_11087722:                                                   Else
  loc_11087741:                                                     var_118 = arg_C
  loc_1108781D:                                                     frmGzToPzTGZS.VFG.DispID_0082(12, "UserCurrent".00000000h.00000000h)
  loc_1108786C:                                                     var_1F0 = var_20
  loc_1108789E:                                                     If var_20.UnkVCall_0000010Ch Then
  loc_11087982:                                                       If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 13) = global_1100AE28) Then
  loc_11087A19:                                                         var_817C = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 13)
  loc_11087A2C:                                                         var_78 = var_817C
  loc_11087A5B:                                                         var_90 = "JsfsCheck".00000001h(1, global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, var_74, var_817C, var_7C)
  loc_11087AAB:                                                         If CBool(Not(var_90)) Then
  loc_11087AB5:                                                           var_24 = "结算方式非法"
  loc_11087AC6:                                                         Else
  loc_11087AC6:                                                         End If
  loc_11087AC6:                                                       End If
  loc_11087AE9:                                                       var_1F0 = var_20
  loc_11087AEF:                                                       var_1D4 = var_20.UnkVCall_0000010Ch
  loc_11087B36:                                                       var_1F8 = var_20
  loc_11087B3C:                                                       var_1D8 = var_20.UnkVCall_00000094h
  loc_11087B83:                                                       var_200 = var_20
  loc_11087BDB:                                                       If (var_20.UnkVCall_0000009Ch = 0) = 0 Then
  loc_11087CBF:                                                         If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 14) = global_1100AE28) Then
  loc_11087D56:                                                           var_8198 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 14)
  loc_11087D69:                                                           var_78 = var_8198
  loc_11087DAC:                                                           var_90 = "GenLen".00000001h(1, global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, var_74, var_8198, var_7C)
  loc_11087DF5:                                                           If (var_90 > 10) Then
  loc_11087DFF:                                                             var_24 = "票号超长"
  loc_11087E10:                                                           Else
  loc_11087E10:                                                           End If
  loc_11087EEE:                                                           If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 15) = global_1100AE28) Then
  loc_11087F85:                                                             var_81A8 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 15)
  loc_11087F98:                                                             var_78 = var_81A8
  loc_11087FC7:                                                             var_90 = "DateCheck".00000001h(1, global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, var_74, var_81A8, var_7C)
  loc_11088017:                                                             If CBool(Not(var_90)) Then
  loc_11088021:                                                               var_24 = "票号发生日期非法"
  loc_11088032:                                                             Else
  loc_11088032:                                                             End If
  loc_11088032:                                                           End If
  loc_11088055:                                                           var_1F0 = var_20
  loc_110880A2:                                                           var_1F8 = var_20
  loc_110880A8:                                                           var_1D8 = var_20.UnkVCall_0000008Ch
  loc_1108810B:                                                           If (var_20.UnkVCall_000000A4h = 0) = 0 Then
  loc_110881CA:                                                             If (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 16) = global_1100AE28) Then
  loc_11088274:                                                               var_78 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, 16)
  loc_110882F4:                                                               var_38 = "BmCheck".00000002h(var_154, 0, var_15C)
  loc_11088326:                                                               var_81C8 = (var_38 = global_1100AE28)
  loc_1108832E:                                                               If var_81C8 = 0 Then
  loc_11088338:                                                                 var_24 = "部门非法"
  loc_11088349:                                                               Else
  loc_11088366:                                                                 var_118 = arg_C
  loc_110883F0:                                                                 frmGzToPzTGZS.VFG.DispID_0082(16, var_38)
  loc_11088425:                                                                 var_1F0 = var_20
  loc_11088457:                                                                 If var_20.UnkVCall_000000A4h Then
  loc_11088465:                                                                   var_118 = var_38
  loc_110884B7:                                                                   var_128 = var_28
  loc_11088505:                                                                   "BmToProperties".00000002h
  loc_11088522:                                                                   Set var_28 = {3302AA19-EB96-11D2-AF06000021009B21}()
  loc_11088550:                                                                   var_1F0 = var_28
  loc_11088556:                                                                   var_1D4 = var_28.UnkVCall_00000034h
  loc_1108857C:                                                                   If var_1D4 = 0 Then
  loc_1108858A:                                                                     var_24 = "部门非末级"
  loc_1108859B:                                                                   Else
  loc_110885A3:                                                                     var_24 = "部门为空"
  loc_110885B4:                                                                   Else
  loc_11088636:                                                                     frmGzToPzTGZS.VFG.DispID_0082(var_128, 285257256)
  loc_11088648:                                                                   End If
  loc_11088648:                                                                 End If
  loc_1108866B:                                                                 var_1F0 = var_20
  loc_1108869D:                                                                 If var_20.UnkVCall_0000008Ch Then
  loc_11088749:                                                                   var_81E0 = (frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H11) = global_1100AE28)
  loc_11088781:                                                                   If var_81E0 Then
  loc_1108882D:                                                                     var_81E8 = (frmGzToPzTGZS.VFG.DispID_0082(arg_C, 16) = global_1100AE28)
  loc_11088889:                                                                     If var_81E8 + 1 Then
  loc_1108890C:                                                                       var_78 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H11)
  loc_1108899A:                                                                       var_90 = "ZyCheck".00000003h(var_174, "BmCheck".00000002h(var_154, 80020004h, var_15C), var_17C)
  loc_110889AF:                                                                       var_34 = var_90
  loc_110889E1:                                                                       var_81F4 = (var_34 = global_1100AE28)
  loc_110889E9:                                                                       If var_81F4 = 0 Then
  loc_110889F3:                                                                         var_24 = "职员非法"
  loc_11088A04:                                                                       Else
  loc_11088A21:                                                                         var_118 = arg_C
  loc_11088AAB:                                                                         frmGzToPzTGZS.VFG.DispID_0082(&H11, var_34)
  loc_11088ACA:                                                                         var_118 = var_34
  loc_11088B17:                                                                         var_128 = var_18
  loc_11088B65:                                                                         "ZyToProperties".00000002h
  loc_11088B82:                                                                         Set var_18 = {3302AA41-EB96-11D2-AF06000021009B21}()
  loc_11088B90:                                                                         var_118 = arg_C
  loc_11088BD1:                                                                         var_1F0 = var_18
  loc_11088C8A:                                                                         frmGzToPzTGZS.VFG.DispID_0082(var_128, var_18.UnkVCall_0000002Ch)
  loc_11088CAA:                                                                       Else
  loc_11088D20:                                                                         var_158 = var_38
  loc_11088D2D:                                                                         var_78 = frmGzToPzTGZS.VFG.DispID_0082(8, var_128)
  loc_11088DE2:                                                                         var_34 = "ZyCheck".00000003h(var_164, 0, var_16C)
  loc_11088E14:                                                                         var_8208 = (var_34 = global_1100AE28)
  loc_11088E1C:                                                                         If var_8208 = 0 Then
  loc_11088E26:                                                                           var_24 = "职员不在指定部门内"
  loc_11088E37:                                                                         Else
  loc_11088E75:                                                                           var_118 = arg_C
  loc_11088EDC:                                                                           frmGzToPzTGZS.VFG.DispID_0082(&H11, var_34)
  loc_11088EEE:                                                                         End If
  loc_11088EEE:                                                                       End If
  loc_11088EEE:                                                                     End If
  loc_11088F11:                                                                     var_1F0 = var_20
  loc_11088F43:                                                                     If var_20.UnkVCall_00000094h Then
  loc_11088FEF:                                                                       var_8214 = (frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H12) = global_1100AE28)
  loc_11089000:                                                                       var_1F0 = var_8214
  loc_11089027:                                                                       If var_1F0 = 0 Then GoTo loc_11089515
  loc_110890D1:                                                                       var_78 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H12)
  loc_11089151:                                                                       var_3C = "KhCheck".00000002h(var_154, 0, var_15C)
  loc_11089183:                                                                       var_8220 = (var_3C = global_1100AE28)
  loc_1108918B:                                                                       If var_8220 = 0 Then
  loc_11089195:                                                                         var_24 = "客户非法"
  loc_110891A6:                                                                       Else
  loc_110891E4:                                                                         var_118 = arg_C
  loc_1108924B:                                                                         frmGzToPzTGZS.VFG.DispID_0082(&H12, var_3C)
  loc_1108925D:                                                                       End If
  loc_11089280:                                                                       var_1F0 = var_20
  loc_110892B2:                                                                       If var_20.UnkVCall_0000009Ch Then
  loc_1108935E:                                                                         var_822C = (frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H13) = global_1100AE28)
  loc_1108936F:                                                                         var_1F0 = var_822C
  loc_11089396:                                                                         If var_1F0 = 0 Then GoTo loc_110898CE
  loc_11089440:                                                                         var_78 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H13)
  loc_110894C0:                                                                         var_30 = "GysCheck".00000002h(var_154, 0, var_15C)
  loc_110894F2:                                                                         var_8238 = (var_30 = global_1100AE28)
  loc_110894FA:                                                                         If var_8238 = 0 Then
  loc_11089504:                                                                           var_24 = "供应商非法"
  loc_11089510:                                                                           GoTo loc_11089FBE
  loc_1108951D:                                                                           var_24 = "客户为空"
  loc_1108952E:                                                                         Else
  loc_1108956C:                                                                           var_118 = arg_C
  loc_110895D3:                                                                           frmGzToPzTGZS.VFG.DispID_0082(&H13, var_30)
  loc_110895E5:                                                                         End If
  loc_11089608:                                                                         var_1F0 = var_20
  loc_11089655:                                                                         var_1F8 = var_20
  loc_1108965B:                                                                         var_1D8 = var_20.UnkVCall_0000009Ch
  loc_11089699:                                                                         If (var_20.UnkVCall_00000094h = 0) = 0 Then
  loc_11089745:                                                                           var_8248 = (frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H14) = global_1100AE28)
  loc_1108977D:                                                                           If var_8248 Then
  loc_11089814:                                                                             var_824C = frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H14)
  loc_11089827:                                                                             var_78 = var_824C
  loc_1108986A:                                                                             var_90 = "GenLen".00000001h(global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, global_11089FFD, var_74, var_824C, var_7C)
  loc_110898B3:                                                                             If (var_90 > 20) Then
  loc_110898BD:                                                                               var_24 = "业务员超长"
  loc_110898C9:                                                                               GoTo loc_11089FBE
  loc_110898D6:                                                                               var_24 = "供应商为空"
  loc_110898E7:                                                                             Else
  loc_110898E7:                                                                             End If
  loc_110898E7:                                                                           End If
  loc_11089907:                                                                           var_1F0 = var_20
  loc_1108995F:                                                                           If (var_20.UnkVCall_000000ACh = global_1100AE28) Then
  loc_11089A0B:                                                                             var_8260 = (frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H15) = global_1100AE28)
  loc_11089A43:                                                                             If var_8260 Then
  loc_11089A69:                                                                               var_1F0 = var_20
  loc_11089A9C:                                                                               var_8268 = (var_20.UnkVCall_000000ACh = global_1100AE28)
  loc_11089AC1:                                                                               If var_8268 Then
  loc_11089AE7:                                                                                 var_1F0 = var_20
  loc_11089B17:                                                                                 var_78 = var_20.UnkVCall_000000ACh
  loc_11089BC5:                                                                                 var_88 = frmGzToPzTGZS.VFG.DispID_0082(arg_C, &H15)
  loc_11089C4D:                                                                                 var_A0 = "XmCheck".00000003h(var_164, Not(8), var_16C)
  loc_11089C62:                                                                                 var_2C = var_A0
  loc_11089C9B:                                                                                 var_8278 = (var_2C = global_1100AE28)
  loc_11089CA3:                                                                                 If var_8278 = 0 Then
  loc_11089CAD:                                                                                   var_24 = "项目非法"
  loc_11089CBE:                                                                                 Else
  loc_11089CEA:                                                                                   var_4C = var_20.UnkVCall_000000ACh
  loc_11089D18:                                                                                   var_128 = var_2C
  loc_11089D49:                                                                                   Set var_58 = var_1C
  loc_11089DCB:                                                                                   "XmToProperties".00000003h
  loc_11089DE8:                                                                                   Set var_1C = {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_11089E3D:                                                                                   If var_1C.UnkVCall_00000034h Then
  loc_11089E4B:                                                                                     var_24 = "项目已结算"
  loc_11089E5C:                                                                                   Else
  loc_11089E86:                                                                                     var_118 = %cobj
  loc_11089EFE:                                                                                     frmGzToPzTGZS.VFG.DispID_0082(&H15, 285257256)
  loc_11089F1B:                                                                                   Else
  loc_11089F23:                                                                                     var_24 = "制单日期非法"
  loc_11089F29:                                                                                   End If
  loc_11089F29:                                                                                 End If
  loc_11089F29:                                                                               End If
  loc_11089F2F:                                                                               GoTo loc_11089FBE
  loc_11089F38:                                                                               If var_4 Then
  loc_11089F43:                                                                               End If
  loc_11089FBD:                                                                               Exit Sub
  loc_11089FBE:                                                                             End If
  loc_11089FBE:                                                                           End If
  loc_11089FBE:                                                                         End If
  loc_11089FBE:                                                                       End If
  loc_11089FBE:                                                                     End If
  loc_11089FBE:                                                                   End If
  loc_11089FBE:                                                                 End If
  loc_11089FBE:                                                               End If
  loc_11089FBE:                                                             End If
  loc_11089FBE:                                                           End If
  loc_11089FBE:                                                         End If
  loc_11089FBE:                                                       End If
  loc_11089FBE:                                                     End If
  loc_11089FBE:                                                   End If
  loc_11089FBE:                                                 End If
  loc_11089FBE:                                               End If
  loc_11089FBE:                                             End If
  loc_11089FBE:                                           End If
  loc_11089FBE:                                         End If
  loc_11089FBE:                                       End If
  loc_11089FBE:                                     End If
  loc_11089FBE:                                   End If
  loc_11089FBE:                                 End If
  loc_11089FBE:                               End If
  loc_11089FBE:                             End If
  loc_11089FBE:                           End If
  loc_11089FBE:                         End If
  loc_11089FBE:                       End If
  loc_11089FBE:                     End If
  loc_11089FBE:                   End If
  loc_11089FBE:                 End If
  loc_11089FBE:               End If
  loc_11089FBE:             End If
  loc_11089FBE:           End If
  loc_11089FBE:         End If
  loc_11089FBE:       End If
  loc_11089FBE:     End If
  loc_11089FBE:   End If
  loc_11089FBE: End If
  loc_11089FBE: ' Referenced from: 11089F2F
End Sub

Private Sub Proc_13_14_1108A020
  Dim var_9C As Variant
  Dim var_8034 As Label
  Dim var_8074 As Label
  Dim var_A0 As Variant
  Dim var_38 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA47-EB96-11D2-AF06000021009B21}()
  loc_1108A17A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1108A180: var_294 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1108A1A6: Set var_9C = frmGzToPzTGZS.VFG
  loc_1108A1F0: If (CLng(var_9C.DispID_0007) < 2) Then
  loc_1108A21E:   var_800C = = Global.Screen
  loc_1108A240:   var_8010 = ecx
  loc_1108A248:   var_8010 = var_9C.UnkVCall_0000007Ch
  loc_1108A2B5:   var_C8 = "提示信息"
  loc_1108A2B7:   var_150 = "没有可生成用友凭证的数据。"
  loc_1108A2C6: Else
  loc_1108A376:   var_264 = ("GetAccInfo".00000002h(, , fs:[00000000h], , "GL", var_16C, "dGLStartDate", var_174) = 1100AE28h)
  loc_1108A390:   If var_264 = 0 Then GoTo loc_1108A4D1
  loc_1108A3BE:   var_801C = = Global.Screen
  loc_1108A3E0:   var_8020 = ecx
  loc_1108A3E8:   var_8020 = var_9C.UnkVCall_0000007Ch
  loc_1108A455:   var_C8 = "提示信息"
  loc_1108A457:   var_150 = "总账系统尚未启用，不能进行凭证引入！"
  loc_1108A461: End If
  loc_1108A493: MsgBox(var_150, 64, var_C8, var_D8, var_E8)
  loc_1108A4C0: Exit Sub
  loc_1108A4CC: GoTo loc_1109307D
  loc_1108A4DB: var_8024 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[VouchNum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " & " CREATE TABLE VouchNum(iperiod tinyint NULL ,csign varchar(8) NULL ,ino_id int NULL,constraint index1 unique(iperiod,csign,ino_id))"
  loc_1108A4E1: var_B0 = var_8024
  loc_1108A540: var_D8.00000001h(0, , , , "3缷M鋎?", var_AC, var_8024, var_B4)
  loc_1108A560: On Error GoTo 0
  loc_1108A566: var_B0 = %ecx = %S_edx_S
  loc_1108A588: var_78 = "AS13"
  loc_1108A5A0: var_78)
  loc_1108A5CA: If Not (var_78)) Then
  loc_1108A5FB:   If Global.Screen < 0 Then
  loc_1108A60C:   End If
  loc_1108A616:   var_8030 = ecx
  loc_1108A625:   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1108A638:   Else
  loc_1108A649:     call var_8034 = var_9C(var_9C, frmGzToPzTGZS.Label3, var_9C, global_1100C47C, 0000007Ch)
  loc_1108A64B:     var_264 = var_8034
  loc_1108A659:     Label3.Caption = "正在进行数据分析，请稍等..."
  loc_1108A686:     var_150 = True
  loc_1108A6C9:     call var_8038 = var_9C(var_9C, frmGzToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, True, var_14C)
  loc_1108A6CC:     var_8038.DispID_0000 =
  loc_1108A6F5:     call var_803C = var_9C(var_9C, frmGzToPzTGZS.Pic1, global_FFFFFDDA, var_9C = var_9C)
  loc_1108A6F8:     var_803C.DispID_0000
  loc_1108A717:     var_8040 = .Proc_13_12_11082B50(var_24C)
  loc_1108A725:     If var_24C = 2 Then
  loc_1108A72B:       var_150 = %ecx = %S_edx_S
  loc_1108A76E:       call var_8044 = var_9C(var_9C, frmGzToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C)
  loc_1108A771:       var_8044.DispID_0000 =
  loc_1108A80D:       MsgBox("数据源中没有合法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1108A84A:       var_24C = %ecx = %S_edx_S
  loc_1108A870:       "AS13")
  loc_1108A8B2:       var_B8 = Global.Screen
  loc_1108A8D4:       var_804C = ecx
  loc_1108A8E3:       If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1108A8F6:       Else
  loc_1108A8F8:         If var_804C = 1 Then
  loc_1108A8FE:           var_150 = %ecx = %S_edx_S
  loc_1108A941:           call var_8050 = var_9C(var_9C, frmGzToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, var_14C = var_9C, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_1108A944:           var_8050.DispID_0000 =
  loc_1108A9E0:           MsgBox("数据源中含有非法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1108AA1D:           var_24C = %ecx = %S_edx_S
  loc_1108AA43:           "AS13")
  loc_1108AA85:           var_B8 = Global.Screen
  loc_1108AAA7:           var_8058 = ecx
  loc_1108AAB6:           If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1108AAC9:           Else
  loc_1108AACB:             If var_8058 = 3 Then
  loc_1108AAD1:               var_150 = %ecx = %S_edx_S
  loc_1108AB14:               call var_805C = var_9C(var_9C, frmGzToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_1108AB17:               var_805C.DispID_0000 =
  loc_1108ABB3:               MsgBox("数据源中指定的凭证号无效或重号，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1108ABF0:               var_24C = %ecx = %S_edx_S
  loc_1108AC16:               "AS13")
  loc_1108AC58:               var_B8 = Global.Screen
  loc_1108AC7A:               var_8064 = ecx
  loc_1108AC89:               If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1108AC9C:               Else
  loc_1108ACDE:                 var_C8 = "提示信息"
  loc_1108AD04:                 var_B8 = "数据源中的数据已全部通过检查，是否开始引入？"
  loc_1108AD28:                 MsgBox(var_B8, 36, var_C8, var_D8, var_E8)
  loc_1108AD6D:                 If (MsgBox(var_B8, 36, var_C8, var_D8, var_E8) = 7) Then
  loc_1108ADB8:                   call var_8068 = var_9C(var_9C, frmGzToPzTGZS.Pic1, global_80010007, 0000000Bh, var_154, frmGzToPzTGZS.Pic1, var_14C, var_9C, global_1100C47C, 0000007Ch)
  loc_1108ADBB:                   var_8068.DispID_0000 =
  loc_1108ADE1:                   var_24C = %ecx = %S_edx_S
  loc_1108AE07:                   "AS13")
  loc_1108AE49:                   var_B8 = Global.Screen
  loc_1108AE6B:                   var_8070 = ecx
  loc_1108AE7A:                   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_1108AE8D:                   Else
  loc_1108AE8E:                     On Error GoTo 0
  loc_1108AEA5:                     call var_8074 = var_9C(var_9C, frmGzToPzTGZS.Label3, var_9C = var_9C, var_9C, global_1100C47C, 0000007Ch)
  loc_1108AEA7:                     var_264 = var_8074
  loc_1108AEB5:                     Label3.Caption = "正在写数据，请稍等..."
  loc_1108AEF9:                     call var_8078 = var_9C(var_9C, frmGzToPzTGZS.Pic1, global_FFFFFDDA, 00000000h)
  loc_1108AEFC:                     var_8078.DispID_0000
  loc_1108AF33:                     Set var_74 = CreateObject("UfDbKit.UfRecordset", 0)
  loc_1108AF4A:                     var_150 = "SELECT TOP 1 * FROM GL_accvouch"
  loc_1108AFBF:                     Set var_74 = "DataMdb".00000000h.00000001h(var_14C, "SELECT TOP 1 * FROM GL_accvouch", var_154)
  loc_1108AFF3:                     call var_8084 = var_9C(var_9C, frmGzToPzTGZS.VFG, 00000007h, 00000000h)
  loc_1108B057:                     If var_24 <= CLng(var_8084.DispID_0000)(-1) Then
  loc_1108B061:                       var_2A8 = var_24
  loc_1108B067:                       var_150 = var_24
  loc_1108B0E4:                       call var_8090 = var_9C(var_9C, frmGzToPzTGZS.VFG, 00000082h, 00000002h, 3, var_174, 2, var_16C, 00000003h, var_154, var_24, var_14C)
  loc_1108B0FE:                       var_C0 = var_8090.DispID_0000
  loc_1108B11C:                       var_D8)
  loc_1108B174:                       var_70 = CByte("DateToPeriod".00000001h(8, var_D4))
  loc_1108B1AD:                       var_150 = var_2A8
  loc_1108B226:                       call var_809C = var_9C(var_9C, frmGzToPzTGZS.VFG, 00000082h, 00000002h, 3, var_174, 3, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_1108B245:                       var_58 = var_809C.DispID_0000
  loc_1108B269:                       var_150 = var_2A8
  loc_1108B2E6:                       call var_80A4 = var_9C(var_9C, frmGzToPzTGZS.VFG, 00000082h, 00000002h, 3, var_174, 0, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_1108B305:                       var_64 = var_80A4.DispID_0000
  loc_1108B329:                       var_150 = var_2A8
  loc_1108B3A6:                       call var_80AC = var_9C(var_9C, frmGzToPzTGZS.VFG, 00000082h, 00000002h, 3, var_174, 1, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_1108B40F:                       If (var_80AC.DispID_0000 = global_1100D76C) Then
  loc_1108B426:                         call var_80B8 = var_9C(var_A8, frmGzToPzTGZS.Label3)
  loc_1108B428:                         var_264 = var_80B8
  loc_1108B538:                         var_80 = "正在处理：第[" & frmGzToPzTGZS.VFG.DispID_0082(var_2A8, 2) & " - "
  loc_1108B679:                         var_D8 = frmGzToPzTGZS.VFG.DispID_0082(var_2A8, 0)
  loc_1108B6C0:                         var_98 = var_80 & frmGzToPzTGZS.VFG.DispID_0082(var_2A8, 3) & " - " & var_D8 & "]号凭证"
  loc_1108B6D0:                         var_98 = var_80B8.UnkVCall_00000054h
  loc_1108B78B:                         frmGzToPzTGZS.Pic1.DispID_FFFFFDDA
  loc_1108B7BF:                         var_3C = var_24
  loc_1108B7D3:                         Set var_9C = frmGzToPzTGZS.Chk
  loc_1108B7D5:                         var_264 = var_9C
  loc_1108B7E7:                         Set var_A0 = var_9C(0)
  loc_1108B80B:                         var_26C = var_A0
  loc_1108B875:                         If (var_A0.Value = 1) Then
  loc_1108B8A8:                           var_24C = CInt("cIYear".00000000h)
  loc_1108B8BD:                           var_24C, var_70)
  loc_1108B8CA:                           var_54 = var_24C, var_70)
  loc_1108B8DB:                         Else
  loc_1108B8F1:                           var_80E8 = .Proc_13_16_11094970(var_70)
  loc_1108B903:                           var_54 = var_258
  loc_1108B906:                         End If
  loc_1108B90B:                         If var_54 > 0 Then
  loc_1108B913:                           On Error GoTo loc_1109127D
  loc_1108B94C:                           "wksAlias".00000000h.00000000h(var_58)
  loc_1108B96B:                           var_1A0 = var_70
  loc_1108BA34:                           var_D8)
  loc_1108BAE0:                           var_80FC = (var_58 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 3))
  loc_1108BAED:                           var_1F0 = var_80FC + 1
  loc_1108BBA9:                           var_8104 = (var_64 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 0))
  loc_1108BBB6:                           var_240 = var_8104 + 1
  loc_1108BC4C:                           var_8110 = (frmGzToPzTGZS.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_80FC + 1 And var_8104 + 1
  loc_1108BCD8:                           If CBool(var_8110) Then
  loc_1108BD79:                             var_C0 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 6)
  loc_1108BDB6:                             var_1A0 = var_38
  loc_1108BE24:                             "kmCodeToProperties".00000002h
  loc_1108BE44:                             Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_1108BE7D:                             var_74.AddNew
  loc_1108BE88:                             var_150 = "ibook"
  loc_1108BEF9:                             var_74.DispID_0000(0)
  loc_1108BEFB:                             var_1A0 = "iPeriod"
  loc_1108BFAA:                             var_C0 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 2)
  loc_1108BFC8:                             var_D8)
  loc_1108C061:                             var_74.DispID_0000("DateToPeriod".00000001h)
  loc_1108C096:                             var_190 = "csign"
  loc_1108C1A3:                             var_74.DispID_0000(frmGzToPzTGZS.VFG.DispID_0082(var_24, 3))
  loc_1108C1CA:                             var_190 = "isignseq"
  loc_1108C2EA:                             var_74.DispID_0000(Proc_0_4_11026BD0(frmGzToPzTGZS.VFG.DispID_0082(var_24, 3), var_64, var_258))
  loc_1108C315:                             var_150 = "ino_id"
  loc_1108C387:                             var_74.DispID_0000(var_54)
  loc_1108C389:                             var_190 = "dbill_date"
  loc_1108C438:                             var_C0 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 2)
  loc_1108C456:                             var_D8)
  loc_1108C4B3:                             var_74.DispID_0000(var_D8)
  loc_1108C4E1:                             var_190 = "idoc"
  loc_1108C4F9:                             var_150 = var_24
  loc_1108C602:                             var_74.DispID_0000(Val(frmGzToPzTGZS.VFG.DispID_0082(var_150, 4)))
  loc_1108C62D:                             var_160 = "ctext1"
  loc_1108C694:                             var_74.DispID_0000(var_150)
  loc_1108C69B:                             var_160 = "ctext2"
  loc_1108C702:                             var_74.DispID_0000(var_150)
  loc_1108C709:                             var_150 = "cbill"
  loc_1108C777:                             var_74.DispID_0000("cUserName".00000000h(, var_14C, "cbill", var_154))
  loc_1108C78D:                             var_160 = "cbook"
  loc_1108C7F4:                             var_74.DispID_0000(var_150)
  loc_1108C7FB:                             var_160 = "ccheck"
  loc_1108C862:                             var_74.DispID_0000(var_150)
  loc_1108C869:                             var_160 = "ccashier"
  loc_1108C8D0:                             var_74.DispID_0000(var_150)
  loc_1108C8D7:                             var_160 = "iflag"
  loc_1108C93E:                             var_74.DispID_0000(var_150)
  loc_1108C945:                             var_160 = "coutaccset"
  loc_1108C9AC:                             var_74.DispID_0000(var_150)
  loc_1108C9B3:                             var_160 = "ioutyear"
  loc_1108CA1A:                             var_74.DispID_0000(var_150)
  loc_1108CA21:                             var_160 = "coutsysver"
  loc_1108CA88:                             var_74.DispID_0000(var_150)
  loc_1108CA8F:                             var_160 = "coutsysname"
  loc_1108CAF6:                             var_74.DispID_0000(var_150)
  loc_1108CAFD:                             var_170 = "ioutperiod"
  loc_1108CB9A:                             var_74.DispID_0000(var_74.DispID_0000("iPeriod"))
  loc_1108CBAB:                             var_170 = "doutbilldate"
  loc_1108CC6E:                             var_74.DispID_0000(CStr(var_74.DispID_0000("dbill_date")))
  loc_1108CC8B:                             var_150 = "iYear"
  loc_1108CCF9:                             var_74.DispID_0000("cIYear".00000000h(var_58, var_14C, "iYear", var_154))
  loc_1108CDF7:                             var_74.DispID_0000("cIYear".00000000h(, var_16C, "iYPeriod", var_174) & Format(var_70, "00"))
  loc_1108CE25:                             var_160 = "coutsign"
  loc_1108CE8C:                             var_74.DispID_0000(var_70)
  loc_1108CE8E:                             var_190 = "coutno_id"
  loc_1108CF9B:                             var_74.DispID_0000(frmGzToPzTGZS.VFG.DispID_0082(var_24, 3))
  loc_1108CFC7:                             var_150 = "bvouchedit"
  loc_1108D036:                             var_74.DispID_0000(FFFFFFFFh)
  loc_1108D03D:                             var_150 = "bvouchaddordele"
  loc_1108D0AE:                             var_74.DispID_0000(FFFFFFFFh)
  loc_1108D0B5:                             var_150 = "bvouchmoneyhold"
  loc_1108D126:                             var_74.DispID_0000(FFFFFFFFh)
  loc_1108D12D:                             var_150 = "bvalueedit"
  loc_1108D19E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_1108D1A5:                             var_150 = "bcodeedit"
  loc_1108D216:                             var_74.DispID_0000(FFFFFFFFh)
  loc_1108D21D:                             var_150 = "bPCSedit"
  loc_1108D28E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_1108D295:                             var_150 = "bDeptedit"
  loc_1108D306:                             var_74.DispID_0000(FFFFFFFFh)
  loc_1108D30D:                             var_150 = "bItemedit"
  loc_1108D37E:                             var_74.DispID_0000(FFFFFFFFh)
  loc_1108D385:                             var_150 = "inid"
  loc_1108D3F7:                             var_74.DispID_0000(1)
  loc_1108D3F9:                             var_190 = "cdigest"
  loc_1108D50A:                             var_74.DispID_0000(frmGzToPzTGZS.VFG.DispID_0082(var_24, 5))
  loc_1108D531:                             var_190 = "cCode"
  loc_1108D640:                             var_74.DispID_0000(frmGzToPzTGZS.VFG.DispID_0082(var_24, 6))
  loc_1108D6E8:                             var_7C = var_38.UnkVCall_0000006Ch
  loc_1108D733:                             var_8150 = (var_38.UnkVCall_0000006Ch = global_1100AE28)
  loc_1108D740:                             var_160 = var_8150 + 1
  loc_1108D7CB:                             var_74.DispID_0000(IIf(var_8150 + 1, vbNull, 0))
  loc_1108D8B0:                             var_1B0 = "md"
  loc_1108D8F9:                             var_BC = var_25C
  loc_1108D980:                             var_74.DispID_0000(Format(Val(frmGzToPzTGZS.VFG.DispID_0082(var_24, 7)), "#.00"))
  loc_1108DA71:                             var_1B0 = "mc"
  loc_1108DABA:                             var_BC = var_25C
  loc_1108DB41:                             var_74.DispID_0000(Format(Val(frmGzToPzTGZS.VFG.DispID_0082(var_24, 8)), "#.00"))
  loc_1108DC09:                             If (var_74.DispID_0000("md") <> 0) Then
  loc_1108DC7E:                               If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_1108DC89:                                 var_150 = "md_f"
  loc_1108DCFA:                                 var_74.DispID_0000(0)
  loc_1108DD04:                               Else
  loc_1108DDB7:                                 var_1B0 = "md_f"
  loc_1108DE00:                                 var_BC = var_25C
  loc_1108DE87:                                 var_74.DispID_0000(Format(Val(frmGzToPzTGZS.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_1108DEC8:                               End If
  loc_1108DF3A:                               If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_1108DF45:                                 var_150 = "nd_s"
  loc_1108DFB6:                                 var_74.DispID_0000(0)
  loc_1108DFC0:                               Else
  loc_1108DFCF:                               Else
  loc_1108E03E:                                 If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_1108E049:                                   var_150 = "mc_f"
  loc_1108E0BA:                                   var_74.DispID_0000(0)
  loc_1108E0C4:                                 Else
  loc_1108E177:                                   var_1B0 = "mc_f"
  loc_1108E1C0:                                   var_BC = var_25C
  loc_1108E247:                                   var_74.DispID_0000(Format(Val(frmGzToPzTGZS.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_1108E288:                                 End If
  loc_1108E2FA:                                 If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_1108E301:                                   GoTo loc_1108DF45
  loc_1108E306:                                 End If
  loc_1108E310:                               End If
  loc_1108E42A:                               var_74.DispID_0000(Val(frmGzToPzTGZS.VFG.DispID_0082(var_24, 9)))
  loc_1108E450:                             End If
  loc_1108E4C2:                             If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_1108E4CD:                               var_150 = "nfrat"
  loc_1108E53E:                               var_74.DispID_0000(0)
  loc_1108E548:                             Else
  loc_1108E66C:                               var_74.DispID_0000(Val(frmGzToPzTGZS.VFG.DispID_0082(var_24, 11)))
  loc_1108E692:                             End If
  loc_1108E6E7:                             If var_38.UnkVCall_0000010Ch Then
  loc_1108E77E:                               var_1F0 = "csettle"
  loc_1108E865:                               var_81A4 = (frmGzToPzTGZS.VFG.DispID_0082(var_24, 13) = global_1100AE28)
  loc_1108E872:                               var_1E0 = var_81A4 + 1
  loc_1108E8FD:                               var_74.DispID_0000(IIf(var_81A4 + 1, vbNull, frmGzToPzTGZS.VFG.DispID_0082(var_24, 13)))
  loc_1108E956:                             End If
  loc_1108E97F:                             var_24C = var_38.UnkVCall_0000010Ch
  loc_1108E9CC:                             var_250 = var_38.UnkVCall_00000094h
  loc_1108EA6B:                             If (var_38.UnkVCall_0000009Ch = 0) = 0 Then
  loc_1108EB02:                               var_1F0 = "cn_id"
  loc_1108EBB1:                               var_E0 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 14)
  loc_1108EBE9:                               var_81BC = (frmGzToPzTGZS.VFG.DispID_0082(var_24, 14) = global_1100AE28)
  loc_1108EBF6:                               var_1E0 = var_81BC + 1
  loc_1108EC81:                               var_74.DispID_0000(IIf(var_81BC + 1, vbNull, var_E0))
  loc_1108ED68:                               var_1F0 = "dt_date"
  loc_1108EE17:                               var_D0 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 15)
  loc_1108EE35:                               var_E0)
  loc_1108EE62:                               var_81C8 = (frmGzToPzTGZS.VFG.DispID_0082(var_24, 15) = global_1100AE28)
  loc_1108EE6F:                               var_1E0 = var_81C8 + 1
  loc_1108EEFA:                               var_74.DispID_0000(IIf(var_81C8 + 1, vbNull, var_E0))
  loc_1108EFE8:                               var_1F0 = "cname"
  loc_1108F0CF:                               var_81D4 = (frmGzToPzTGZS.VFG.DispID_0082(var_24, &H14) = global_1100AE28)
  loc_1108F0DC:                               var_1E0 = var_81D4 + 1
  loc_1108F167:                               var_74.DispID_0000(IIf(var_81D4 + 1, vbNull, frmGzToPzTGZS.VFG.DispID_0082(var_24, &H14)))
  loc_1108F1C0:                             End If
  loc_1108F236:                             var_250 = var_38.UnkVCall_0000008Ch
  loc_1108F274:                             If (var_38.UnkVCall_000000A4h = 0) = 0 Then
  loc_1108F27E:                               var_150 = var_24
  loc_1108F30B:                               var_1F0 = "cdept_id"
  loc_1108F3F2:                               var_81E8 = (frmGzToPzTGZS.VFG.DispID_0082(var_150, 16) = global_1100AE28)
  loc_1108F3FF:                               var_1E0 = var_81E8 + 1
  loc_1108F48A:                               var_74.DispID_0000(IIf(var_81E8 + 1, vbNull, frmGzToPzTGZS.VFG.DispID_0082(var_24, 16)))
  loc_1108F4E5:                             Else
  loc_1108F4EA:                               var_160 = "cdept_id"
  loc_1108F551:                               var_74.DispID_0000(var_150)
  loc_1108F556:                             End If
  loc_1108F5AB:                             If var_38.UnkVCall_0000008Ch Then
  loc_1108F5B5:                               var_150 = var_24
  loc_1108F642:                               var_1F0 = "cperson_id"
  loc_1108F729:                               var_81F8 = (frmGzToPzTGZS.VFG.DispID_0082(var_150, &H11) = global_1100AE28)
  loc_1108F736:                               var_1E0 = var_81F8 + 1
  loc_1108F7C1:                               var_74.DispID_0000(IIf(var_81F8 + 1, vbNull, frmGzToPzTGZS.VFG.DispID_0082(var_24, &H11)))
  loc_1108F81C:                             Else
  loc_1108F821:                               var_160 = "cperson_id"
  loc_1108F888:                               var_74.DispID_0000(var_150)
  loc_1108F88D:                             End If
  loc_1108F8E2:                             If var_38.UnkVCall_00000094h Then
  loc_1108F8EC:                               var_150 = var_24
  loc_1108F979:                               var_1F0 = "ccus_id"
  loc_1108FA60:                               var_8208 = (frmGzToPzTGZS.VFG.DispID_0082(var_150, &H12) = global_1100AE28)
  loc_1108FA6D:                               var_1E0 = var_8208 + 1
  loc_1108FAF8:                               var_74.DispID_0000(IIf(var_8208 + 1, vbNull, frmGzToPzTGZS.VFG.DispID_0082(var_24, &H12)))
  loc_1108FB53:                             Else
  loc_1108FB58:                               var_160 = "ccus_id"
  loc_1108FBBF:                               var_74.DispID_0000(var_150)
  loc_1108FBC4:                             End If
  loc_1108FC19:                             If var_38.UnkVCall_0000009Ch Then
  loc_1108FC23:                               var_150 = var_24
  loc_1108FCB0:                               var_1F0 = "csup_id"
  loc_1108FD97:                               var_8218 = (frmGzToPzTGZS.VFG.DispID_0082(var_150, &H13) = global_1100AE28)
  loc_1108FDA4:                               var_1E0 = var_8218 + 1
  loc_1108FE2F:                               var_74.DispID_0000(IIf(var_8218 + 1, vbNull, frmGzToPzTGZS.VFG.DispID_0082(var_24, &H13)))
  loc_1108FE8A:                             Else
  loc_1108FE8F:                               var_160 = "csup_id"
  loc_1108FEF6:                               var_74.DispID_0000(var_150)
  loc_1108FEFB:                             End If
  loc_1108FF74:                             If (var_38.UnkVCall_000000ACh = global_1100AE28) Then
  loc_1108FF7E:                               var_150 = var_24
  loc_1109000B:                               var_1F0 = "citem_id"
  loc_110900F2:                               var_822C = (frmGzToPzTGZS.VFG.DispID_0082(var_150, &H15) = global_1100AE28)
  loc_110900FF:                               var_1E0 = var_822C + 1
  loc_1109018A:                               var_74.DispID_0000(IIf(var_822C + 1, vbNull, frmGzToPzTGZS.VFG.DispID_0082(var_24, &H15)))
  loc_11090267:                               var_7C = var_38.UnkVCall_000000ACh
  loc_110902B8:                               var_8238 = (var_38.UnkVCall_000000ACh = global_1100AE28)
  loc_110902C5:                               var_160 = var_8238 + 1
  loc_11090350:                               var_74.DispID_0000(IIf(var_8238 + 1, vbNull, 0))
  loc_1109038A:                             Else
  loc_1109038F:                               var_160 = "citem_id"
  loc_110903F6:                               var_74.DispID_0000(var_150)
  loc_110903FD:                               var_160 = "citem_class"
  loc_11090464:                               var_74.DispID_0000(var_150)
  loc_11090469:                             End If
  loc_1109046E:                             var_160 = "ccode_equal"
  loc_110904D5:                             var_74.DispID_0000(var_150)
  loc_110904DC:                             var_160 = "iflagbank"
  loc_11090543:                             var_74.DispID_0000(var_150)
  loc_1109054A:                             var_160 = "iflagperson"
  loc_110905B1:                             var_74.DispID_0000(var_150)
  loc_110905BE:                             var_74.Update
  loc_110905D5:                             var_24 = var_24(1)
  loc_110905E6:                             var_68 = var_68(1)
  loc_1109061B:                             var_823C = CLng(frmGzToPzTGZS.VFG.DispID_0007)
  loc_11090637:                             var_264 = (var_24(1) > 0)
  loc_1109065E:                             If var_264 = 0 Then GoTo loc_1108B968
  loc_11090664:                           End If
  loc_11090697:                           "wksAlias".00000000h.00000000h
  loc_110906C4:                           Set var_9C = frmGzToPzTGZS.Chk
  loc_110906C6:                           var_264 = var_9C
  loc_110906D8:                           Set var_A0 = var_9C(0)
  loc_110906FC:                           var_26C = var_A0
  loc_11090766:                           If (var_A0.Value = 1) Then
  loc_11090774:                             var_70, var_58)
  loc_11090779:                           End If
  loc_1109077B:                           On Error GoTo 0
  loc_110907B2:                           var_250 = CInt("cIYear".00000000h)
  loc_110907DC:                           var_24C, var_250, var_70, var_58)
  loc_110907E6:                           var_5C = var_24C, var_250, var_70, var_58)
  loc_11090829:                           var_250 = CInt("cIYear".00000000h)
  loc_1109085D:                           var_48 = r_250, var_70, var_58) var_250, var_70, var_58)
  loc_1109086F:                           var_150 = "select * from GL_accvouch where ibook=0 and iYear="
  loc_11090897:                           var_170 = var_70
  loc_110908BB:                           var_824C = Proc_0_4_11026BD0(var_58, var_54, var_54)
  loc_110908C0:                           var_190 = var_824C
  loc_110908E8:                           var_1B0 = var_54
  loc_11090941:                           var_D8 = 1 & "cIYear".00000000h(, 1, 1) & " and iperiod="
  loc_110909AA:                           var_128 = var_D8 & var_70 & " and isignseq=" & var_824C & " and ino_id=" & var_54
  loc_11090A13:                           Set var_74 = "DataMdb".00000000h.00000001h
  loc_11090AB2:                           If CBool(Not(var_74.EOF)) Then
  loc_11090B0A:                             If CBool(Not(var_74.EOF)) Then
  loc_11090B13:                               var_170 = var_70
  loc_11090B28:                               var_150 = "iPeriod"
  loc_11090B4C:                               var_180 = "csign"
  loc_11090B60:                               var_1D0 = var_54
  loc_11090B71:                               var_1B0 = "ino_id"
  loc_11090CC8:                               If CBool((var_70 = var_14C) And (var_58 = var_D8) And (var_54 = var_1AC)) Then
  loc_11090CD3:                                 var_150 = "mc"
  loc_11090D55:                                 var_180 = "ccode_equal"
  loc_11090D69:                                 If (var_14C <> 0) Then
  loc_11090D95:                                   var_8278 = (var_5C = global_1100AE28)
  loc_11090DA2:                                   var_160 = var_8278 + 1
  loc_11090DCF:                                   var_C8 = IIf(var_8278 + 1, vbNull, var_5C)
  loc_11090E49:                                 Else
  loc_11090E6F:                                   var_827C = (var_48 = global_1100AE28)
  loc_11090E7C:                                   var_160 = var_827C + 1
  loc_11090EA9:                                   var_C8 = IIf(var_827C + 1, vbNull, var_48)
  loc_11090F1E:                                 End If
  loc_11090F34:                                 var_74.Update
  loc_11090F7E:                                 var_180 = var_38
  loc_11090FC5:                                 var_B8 = var_74.DispID_0000("cCode")
  loc_11091022:                                 "kmCodeToProperties".00000002h
  loc_11091042:                                 Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_11091061:                                 var_150 = "citem_class"
  loc_110910C8:                                 If IsNull(var_74.DispID_0000(var_150)) Then
  loc_110910DD:                                 Else
  loc_1109111E:                                   var_180 = var_28
  loc_11091165:                                   var_B8 = var_74.DispID_0000(var_150)
  loc_110911C2:                                   "XmClassIDToProperties".00000002h
  loc_11091222:                                   var_78 = {3302AA47-EB96-11D2-AF06000021009B21}().UnkVCall_0000002Ch
  loc_11091253:                                 End If
  loc_11091261:                                 var_68 = var_68(1)
  loc_1109126F:                                 var_74.MoveNext
  loc_11091278:                                 GoTo loc_11090ABF
  loc_110912B0:                                 "wksAlias".00000000h.00000000h
  loc_110912C8:                                 var_30 = var_3C
  loc_110912DD:                                 var_1A0 = var_70
  loc_110913A6:                                 var_D8)
  loc_11091452:                                 var_829C = (var_58 = frmGzToPzTGZS.VFG.DispID_0082(var_30, 3))
  loc_1109145F:                                 var_1F0 = var_829C + 1
  loc_1109151B:                                 var_82A4 = (var_64 = frmGzToPzTGZS.VFG.DispID_0082(var_30, 0))
  loc_11091528:                                 var_240 = var_82A4 + 1
  loc_110915BE:                                 var_82B0 = (frmGzToPzTGZS.VFG.DispID_0082(var_30, 2) = "DateToPeriod".00000001h) And var_829C + 1 And var_82A4 + 1
  loc_1109164A:                                 If CBool(var_82B0) Then
  loc_11091654:                                   var_150 = var_30
  loc_11091710:                                   frmGzToPzTGZS.VFG.DispID_0082(1, "-")
  loc_11091890:                                   frmGzToPzTGZS.VFG.DispID_009E(var_30, 1, var_30, 1, &HFF)
  loc_110918A5:                                   var_150 = var_30
  loc_11091961:                                   frmGzToPzTGZS.VFG.DispID_0082(&H16, "数据提交错或该数据已经被导入----未引入")
  loc_11091980:                                   var_30 = var_30(1)
  loc_110919AC:                                   var_82B8 = CLng(frmGzToPzTGZS.VFG.DispID_0007)
  loc_110919C8:                                   var_264 = (var_30 > 0)
  loc_110919EF:                                   If var_264 = 0 Then GoTo loc_110912DA
  loc_110919F5:                                 End If
  loc_110919F8:                                 var_24 = var_30
  loc_11091A0C:                                 Set var_9C = frmGzToPzTGZS.Chk
  loc_11091A0E:                                 var_264 = var_9C
  loc_11091A20:                                 Set var_A0 = var_9C(0)
  loc_11091A44:                                 var_26C = var_A0
  loc_11091AAE:                                 If (var_A0.Value = 1) Then
  loc_11091BAA:                                   "unLockVouch".00000004h(var_180, var_BC, var_C4, 0, var_74, var_70, var_58, var_16C, var_54, &H4002, var_184)
  loc_11091BB3:                                 End If
  loc_11091BB8:                                 var_150 = "VouchNum"
  loc_11091C2D:                                 Set var_34 = "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C4, 0, var_14C, "VouchNum", var_154)
  loc_11091C4E:                                 var_150 = "delete  from vouchnum"
  loc_11091CAC:                                 "DataMdb".00000000h.00000001h(1, 1, var_180, var_BC, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_11091D09:                                 frmGzToPzTGZS.Pic1.DispID_80010007 = var_150
  loc_11091D1D:                                 var_82C4 = Resume(0)
  loc_11091D23:                               End If
  loc_11091D23:                             End If
  loc_11091D23:                           End If
  loc_11091D41:                           var_24 = var_27C+(var_24 - 1)
  loc_11091D44:                           GoTo loc_1108B04C
  loc_11091D49:                         End If
  loc_11091D4C:                         var_1A0 = var_70
  loc_11091E15:                         var_D8)
  loc_11091EC1:                         var_82D0 = (var_58 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 3))
  loc_11091ECE:                         var_1F0 = var_82D0 + 1
  loc_11091F8A:                         var_82D8 = (var_64 = frmGzToPzTGZS.VFG.DispID_0082(var_24, 0))
  loc_11091F97:                         var_240 = var_82D8 + 1
  loc_1109202D:                         var_82E4 = (frmGzToPzTGZS.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_82D0 + 1 And var_82D8 + 1
  loc_1109203A:                         var_264 = CBool(var_82E4)
  loc_110920B9:                         If var_264 = 0 Then GoTo loc_11091D23
  loc_110920D0:                         Set var_9C = frmGzToPzTGZS.Chk
  loc_110920D2:                         var_264 = var_9C
  loc_110920E4:                         Set var_A0 = var_9C(0)
  loc_11092108:                         var_26C = var_A0
  loc_1109214B:                         var_274 = (var_A0.Value = 1)
  loc_11092176:                         var_150 = var_24
  loc_11092197:                         var_190 = "网络共享冲突----未引入"
  loc_110921A1:                         If var_274 = 0 Then
  loc_110921A3:                           var_190 = "指定的凭证号无效或重号----未引入"
  loc_110921AD:                         End If
  loc_1109223E:                         frmGzToPzTGZS.VFG.DispID_0082(var_170, var_190)
  loc_1109225D:                         var_24 = var_24(1)
  loc_11092263:                         var_2A8 = var_24(1)
  loc_11092292:                         var_82EC = CLng(frmGzToPzTGZS.VFG.DispID_0007)
  loc_110922AE:                         var_264 = (var_2A8 > 0)
  loc_110922D5:                         If var_264 = 0 Then GoTo loc_11091D49
  loc_110922DB:                         GoTo loc_11091D23
  loc_110922E0:                       End If
  loc_110922E3:                       var_1A0 = var_70
  loc_110923AE:                       var_D8)
  loc_1109245C:                       var_82F8 = (var_58 = frmGzToPzTGZS.VFG.DispID_0082(var_2A8, 3))
  loc_11092469:                       var_1F0 = var_82F8 + 1
  loc_11092527:                       var_8300 = (var_64 = frmGzToPzTGZS.VFG.DispID_0082(var_2A8, 0))
  loc_11092534:                       var_240 = var_8300 + 1
  loc_110925CA:                       var_830C = (frmGzToPzTGZS.VFG.DispID_0082(var_2A8, 2) = "DateToPeriod".00000001h) And var_82F8 + 1 And var_8300 + 1
  loc_110925D7:                       var_264 = CBool(var_830C)
  loc_11092656:                       If var_264 = 0 Then GoTo loc_11091D23
  loc_1109274D:                       If (frmGzToPzTGZS.VFG.DispID_0082(var_2A8, &H16) = global_1100AE28) + 1 Then
  loc_11092753:                         var_150 = var_2A8
  loc_1109280C:                         Set var_9C = frmGzToPzTGZS.VFG
  loc_1109280F:                         var_9C.DispID_0082(&H16, "凭证借贷不平衡或某分录有错误----未引入")
  loc_11092820:                         GoTo loc_110922E0
  loc_11092825:                       End If
  loc_110928EF:                       var_C0 = frmGzToPzTGZS.VFG.DispID_0082(frmGzToPzTGZS.VFG, &H16) & "----未引入"
  loc_1109298C:                       frmGzToPzTGZS.VFG.DispID_0082(&H16, var_C0)
  loc_110929C9:                       GoTo loc_110922E0
  loc_110929CE:                     End If
  loc_11092A16:                     frmGzToPzTGZS.Pic1.DispID_80010007 = var_150
  loc_11092A2D:                     If var_2C Then
  loc_11092A3D:                       var_24C = frmGzToPzTGZS.UpdateBTData
  loc_11092AE5:                       MsgBox("数据引入已完成，数据已生成用友凭证。", 64, "提示信息", 10, 10)
  loc_11092B57:                       frmGzToPzTGZS.VFG.DispID_0007 = 1
  loc_11092BF2:                       frmGzToPzTGZS.sBar.DispID_6803001E(1100AE28h)
  loc_11092C89:                       frmGzToPzTGZS.sBar.DispID_6803001E(1100AE28h)
  loc_11092D20:                       Set var_9C = frmGzToPzTGZS.sBar
  loc_11092D23:                       var_9C.DispID_6803001E(1100AE28h)
  loc_11092D39:                     Else
  loc_11092DC0:                       MsgBox("数据没有被引入，原因请查看最后一列中的说明。", 64, "提示信息", 10, 10)
  loc_11092DED:                     End If
  loc_11092DF2:                     var_150 = "VouchNum"
  loc_11092E6B:                     Set var_34 = "DataMdb".00000000h.00000001h(var_180, var_BC, var_C0, var_C4, var_C8, var_14C, "VouchNum", var_154)
  loc_11092E8C:                     var_150 = "delete  from vouchnum"
  loc_11092EDC:                     "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C0, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_11092F31:                     "AS13")
  loc_11092F6A:                     var_B8 = Global.Screen
  loc_11092F8C:                     var_8330 = ecx
  loc_11092F9B:                     If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_11092FA5:                     End If
  loc_11092FA5:                   End If
  loc_11092FA5:                 End If
  loc_11092FA5:               End If
  loc_11092FA5:             End If
  loc_11092FA6:             var_8330 = CheckObj(var_9C, global_1100C47C, 124)
  loc_11092FAC:           End If
  loc_11092FAC:         End If
  loc_11092FAC:       End If
  loc_11092FAC:     End If
  loc_11092FAC:   End If
  loc_11092FAC: End If
  loc_11092FB8: Exit Sub
  loc_11092FC4: GoTo loc_1109307D
  loc_1109307C: Exit Sub
  loc_1109307D: ' Referenced from: 1108A4CC
  loc_1109307D: ' Referenced from: 11092FC4
End Sub

Private Sub Proc_13_15_11093B60
  Dim var_58 As Variant
  Dim var_5C As Variant
  Dim var_64 As frmGzToPzTGZS.Label3
  Dim var_1D0 As Label
  loc_11093C4D: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11093C56: var_1F0 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11093C73: Set var_58 = frmGzToPzTGZS.Chk
  loc_11093C7D: var_1D0 = var_58
  loc_11093C83: Set var_5C = var_58(0)
  loc_11093CAE: var_1D8 = var_5C
  loc_11093CF1: var_1E0 = (var_5C.Value = 1)
  loc_11093D07: If var_1E0 = 0 Then
  loc_11093D6C:   If var_14 <= CLng(frmGzToPzTGZS.VFG.DispID_0007)(-1) Then
  loc_11093DE1:     var_7C = frmGzToPzTGZS.VFG.DispID_0082(var_14, 2)
  loc_11093DFC:     var_94)
  loc_11093E57:     var_30 = CByte("DateToPeriod".00000001h)
  loc_11093FB1:     Set var_64 = frmGzToPzTGZS.Label3
  loc_11093FDB:     var_1D0 = var_64
  loc_11094191:     var_94 = frmGzToPzTGZS.VFG.DispID_0082(var_14, frmGzToPzTGZS.VFG)
  loc_110941AD:     var_8034 = "正在处理：第[" & frmGzToPzTGZS.VFG.DispID_0082(var_14, 2) & " - " & frmGzToPzTGZS.VFG.DispID_0082(var_14, 3) & " - " & var_94
  loc_110941E3:     var_64.Caption = var_8034 & "]号凭证是否重号"
  loc_11094272:     var_803C = frmGzToPzTGZS.Proc_13_16_11094970(var_30)
  loc_11094287:     If var_1CC <= 0 Then
  loc_11094299:       var_13C = var_30
  loc_11094330:       var_94)
  loc_110943C9:       var_804C = (frmGzToPzTGZS.VFG.DispID_0082(var_14, 3) = frmGzToPzTGZS.VFG.DispID_0082(var_14, 3))
  loc_110943F6:       var_17C = var_804C + 1
  loc_1109446D:       var_8054 = (frmGzToPzTGZS.VFG.DispID_0082(var_14, frmGzToPzTGZS.VFG) = frmGzToPzTGZS.VFG.DispID_0082(var_14, ""))
  loc_11094494:       var_1BC = var_8054 + 1
  loc_1109458F:       If CBool((frmGzToPzTGZS.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_804C + 1 And var_8054 + 1) Then
  loc_11094622:         frmGzToPzTGZS.VFG.DispID_0082(var_10C, 285267820)
  loc_11094756:         frmGzToPzTGZS.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_110947EA:         frmGzToPzTGZS.VFG.DispID_0082(var_10C, "指定的凭证号无效或重号")
  loc_11094835:         var_8068 = CLng(frmGzToPzTGZS.VFG.DispID_0007)
  loc_11094853:         var_1D0 = (var_14(1) > 0)
  loc_11094870:         If var_1D0 = 0 Then GoTo loc_11094293
  loc_11094876:       End If
  loc_11094884:     Else
  loc_1109488D:     End If
  loc_1109489A:     var_14 = 1+var_14
  loc_1109489D:     GoTo loc_11093D66
  loc_110948A2:   End If
  loc_110948A2: End If
  loc_110948A7: GoTo loc_11094938
  loc_11094937: Exit Sub
  loc_11094938: ' Referenced from: 110948A7
End Sub

Private  Proc_13_16_11094970(arg_C, arg_10, arg_14) '11094970
  loc_11094A09: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11094A12: var_168 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11094A3B: If IsNumeric(arg_14) Then
  loc_11094A4A:   var_8008 = CLng(Val(arg_14))
  loc_11094A54:   If var_8008 > 0 Then
  loc_11094A60:     If var_8008 <= 9999 Then
  loc_11094ADC:       var_8028 = "select * from GL_accvouch where iperiod >=" & CStr(arg_C) & " and isignseq>=" & CStr(0) & " and ino_id>=" & CStr(var_8008)
  loc_11094AF1:       var_44 = var_8028
  loc_11094B43:       Set var_1C = "DataMdb".00000000h.00000001h(fs:[00000000h], , , , , var_40, var_8028, var_48)
  loc_11094B88:       var_8030 = Proc_0_4_11026BD0(arg_10, , )
  loc_11094BA9:       var_8034 = CBool(var_1C.EOF)
  loc_11094BBD:       If var_8034 = 0 Then
  loc_11094BE8:         var_F4 = arg_C
  loc_11094CA6:         var_8040 = (var_1C.DispID_0000("iPeriod") = arg_C) And (var_1C.DispID_0000("isignseq") = (Proc_0_4_11026BD0(arg_10, , ) And 255))
  loc_11094D16:         var_804C = CBool(Not(var_8040 And (var_1C.DispID_0000("ino_id") = var_8008)))
  loc_11094D3B:         If var_804C = 0 Then GoTo loc_11094D40
  loc_11094D3D:       End If
  loc_11094D4B:       var_1C.oClose
  loc_11094D54:     End If
  loc_11094D54:   End If
  loc_11094D54: End If
  loc_11094D5A: GoTo loc_11094DBF
  loc_11094DBE: Exit Sub
  loc_11094DBF: ' Referenced from: 11094D5A
End Sub
