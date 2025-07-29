VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C908002B2F49FB}#1.2#0"; "C:\Windows\SysWow64\Comdlg32.ocx"
Begin VB.Form frmGzToPzTGZP
  Caption = "工资导转凭证（TGZP）"
  BackColor = &H80000005&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmGzToPzTGZP.frx":0000
  LinkTopic = "Form1"
  ClientLeft = 60
  ClientTop = 345
  ClientWidth = 9255
  ClientHeight = 6645
  Appearance = 0 'Flat
  Begin C1SizerLibCtl.C1Elastic Pic1
    Left = 3300
    Top = 3480
    Width = 5025
    Height = 675
    Visible = 0   'False
    TabStop = 0   'False
    TabIndex = 3
    OleObjectBlob = "frmGzToPzTGZP.frx":014A
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
    Width = 16245
    Height = 6645
    TabStop = 0   'False
    TabIndex = 0
    OleObjectBlob = "frmGzToPzTGZP.frx":0293
    Begin AIFCmp1.asxStatusBar sBar
      Left = 0
      Top = 6300
      Width = 16245
      Height = 345
      OleObjectBlob = "frmGzToPzTGZP.frx":04C0
    End
    Begin C1SizerLibCtl.C1Elastic C1Elastic2
      Left = 0
      Top = 0
      Width = 16245
      Height = 435
      TabStop = 0   'False
      TabIndex = 1
      OleObjectBlob = "frmGzToPzTGZP.frx":05F0
    End
    Begin VSFlex8LCtl.VSFlexGrid VFG
      Left = 0
      Top = 1260
      Width = 16245
      Height = 5025
      TabIndex = 2
      OleObjectBlob = "frmGzToPzTGZP.frx":0753
      Begin MSComDlg.CommonDialog dlg
        OleObjectBlob = "frmGzToPzTGZP.frx":0BBC
        Left = 0
        Top = 0
      End
    End
    Begin AIFCmp1.asxPanel asxPanel1
      Left = 0
      Top = 450
      Width = 16245
      Height = 795
      OleObjectBlob = "frmGzToPzTGZP.frx":0C20
      Begin TDBNumLite6Ctl.TDBNumLite TDBNum
        Index = 0
        Left = 120
        Top = 405
        Width = 1695
        Height = 270
        TabIndex = 16
        OleObjectBlob = "frmGzToPzTGZP.frx":0D00
      End
      Begin VB.ComboBox Cbo
        Style = 2
        Left = 12000
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
        OleObjectBlob = "frmGzToPzTGZP.frx":0E6C
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 4
        Left = 8595
        Top = 75
        Width = 720
        Height = 270
        TabIndex = 13
        OleObjectBlob = "frmGzToPzTGZP.frx":100C
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
        OleObjectBlob = "frmGzToPzTGZP.frx":11AC
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 1
        Left = 6090
        Top = 75
        Width = 870
        Height = 270
        TabIndex = 8
        OleObjectBlob = "frmGzToPzTGZP.frx":13A4
      End
      Begin TDBText6Ctl.TDBText TDBText
        Left = 30
        Top = 75
        Width = 5115
        Height = 270
        TabIndex = 10
        OleObjectBlob = "frmGzToPzTGZP.frx":1574
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
        OleObjectBlob = "frmGzToPzTGZP.frx":16D0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate
        Left = 9390
        Top = 420
        Width = 2385
        Height = 285
        TabIndex = 12
        OleObjectBlob = "frmGzToPzTGZP.frx":1874
      End
      Begin TDBNumLite6Ctl.TDBNumLite TDBNum
        Index = 1
        Left = 2160
        Top = 405
        Width = 2295
        Height = 270
        Visible = 0   'False
        TabIndex = 17
        OleObjectBlob = "frmGzToPzTGZP.frx":1B63
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

Attribute VB_Name = "frmGzToPzTGZP"


Private Sub TDBText_UnknownEvent_B '11071DD0
  Dim var_64 As frmGzToPzTGZP.dlg
  loc_11071E37: Set var_64 = frmGzToPzTGZP.dlg
  loc_11071E69: var_64.FileName = var_48
  loc_11071E8E: var_64.DialogTitle = var_48
  loc_11071EB3: var_64.Filter = var_48
  loc_11071ED5: var_64.CancelError = var_48
  loc_11071EDF: var_64.ShowOpen
  loc_11071EF1: var_64.FileName = var_64
  loc_11071F37: If (var_64 = global_1100AE28) Then
  loc_11071F45:   var_64.FileName = Me
  loc_11071F8D:   frmGzToPzTGZP.TDBText.DispID_0000 = var_2C
  loc_11071FB7: End If
  loc_11071FC3: GoTo loc_11071FEB
  loc_11071FEA: Exit Sub
  loc_11071FEB: ' Referenced from: 11071FC3
End Sub

Private  APB_UnknownEvent_9(arg_C) '11071960
  Dim var_20 As Variant
  Dim var_AC As Scripting.FileSystemObject
  loc_110719D7: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110719E0: var_C4 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11071A07: arg_C = frmGzToPzTGZP.APB.UnkVCall_00000040h
  loc_11071A45: var_B8 = var_24.DispID_FFFFFDFA
  loc_11071A79: var_8008 = (var_B8 = "加载数据")
  loc_11071A7D: If var_8008 = 0 Then
  loc_11071AA0:   var_AC = var_18
  loc_11071ADB:   var_1C = frmGzToPzTGZP.TDBText.DispID_0000
  loc_11071AEB:   var_8014 = Scripting.FileSystemObject.FileExists
  loc_11071B29:   If Not (var_A8) Then
  loc_11071B8C:     MsgBox("文件不存在或非法路径！ ", 64, "提示", 10, 10)
  loc_11071BB2:   Else
  loc_11071BC4:     If frmGzToPzTGZP.FillDataNew < 0 Then
  loc_11071BD6:       var_A8 = CheckObj(%ecx = %S_edx_S = %S_edx_S, global_1100CDAC, 1792)
  loc_11071BE1:     End If
  loc_11071BED:     call ebx("取消加载", var_B8, var_1C, var_A8, var_24)
  loc_11071BF1:     If ebx("取消加载", var_B8, var_1C, var_A8, var_24) = 0 Then
  loc_11071C21:       var_44 = "提示信息"
  loc_11071C4F:       var_2C = "是否取消数据载入？" & vbCrLf & "取消数据载入，数据将全部清空。"
  loc_11071C6B:       MsgBox(var_2C, 292, var_44, var_54, var_64)
  loc_11071CA2:       If (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6) = 0 Then GoTo loc_11071D57
  loc_11071CB3:     Else
  loc_11071CBF:       (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_11071CC3:       If var_B8 = 0 Then
  loc_11071CC8:         var_8020 = frmGzToPzTGZP.Proc_12_15_11068870("凭证导入")
  loc_11071CD3:       Else
  loc_11071CDF:         (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_11071CE3:         If var_8020 Then
  loc_11071CF1:           (MsgBox(var_2C, 292, var_44, var_54, var_64) = 6)
  loc_11071CF5:           If var_B8 = 0 Then
  loc_11071D28:             Set var_20 = var_C4 = %S_edx_S
  loc_11071D36:             var_8028 = Global.Unload var_B8
  loc_11071D57:           End If
  loc_11071D57:         End If
  loc_11071D57:       End If
  loc_11071D57:     End If
  loc_11071D57:   End If
  loc_11071D57: End If
  loc_11071D5F: GoTo loc_11071D96
  loc_11071D95: Exit Sub
  loc_11071D96: ' Referenced from: 11071D5F
End Sub

Private Sub Form_Load() '11060C40
  Dim var_18 As Variant
  Dim var_1C As var_18.DispID_03E8
  Dim var_20 As var_1C.DispID_03E8
  loc_11060CB0: Set var_18 = frmGzToPzTGZP.TDBText
  loc_11060CB7: var_30 = var_18.DispID_03E8
  loc_11060CD8: var_18.DispID_03E8.UnkVCall_00000030h
  loc_11060D26: Set var_18 = frmGzToPzTGZP.TDBDate
  loc_11060D2D: var_30 = var_18.DispID_03E8
  loc_11060D42: Set var_1C = var_18.DispID_03E8
  loc_11060D4E: var_1C.UnkVCall_00000030h
  loc_11060D9D: frmGzToPzTGZP.TDBNum.UnkVCall_00000040h
  loc_11060DC9: var_30 = var_1C.DispID_03E8
  loc_11060DDE: Set var_20 = var_1C.DispID_03E8
  loc_11060E3D: frmGzToPzTGZP.TDBNum.UnkVCall_00000040h
  loc_11060E69: var_30 = var_1C.DispID_03E8
  loc_11060E7E: Set var_20 = var_1C.DispID_03E8
  loc_11060EFD: frmGzToPzTGZP.TDBDate.DispID_0000 = Date
  loc_11060F1F: Set var_18 = frmGzToPzTGZP.APB
  loc_11060F2C: var_18.UnkVCall_00000040h
  loc_11060F6A: var_1C.DispID_80010007 = var_18.DispID_03E8
  loc_11060F91: Set var_18 = frmGzToPzTGZP.APB
  loc_11060F9E: var_18.UnkVCall_00000040h
  loc_11060FD9: var_1C.DispID_80010007 = var_18.DispID_03E8
  loc_11060FF5: var_8004 = frmGzToPzTGZP.Proc_12_12_1104EEC0(var_18)
  loc_11061002: var_54 = frmGzToPzTGZP.getBTData
  loc_1106102A: GoTo loc_1106104D
  loc_1106104C: Exit Sub
  loc_1106104D: ' Referenced from: 1106102A
End Sub

Private Sub Form_Resize() '11061070
  loc_110610FD: var_38 = frmGzToPzTGZP.Pic1.DispID_80010005
  loc_11061121: var_48 = frmGzToPzTGZP.Pic1.DispID_80010006
  loc_11061134: var_EC = var_48.ScaleWidth
  loc_1106116B: If global_110F6000 = 0 Then
  loc_11061175: Else
  loc_11061180: End If
  loc_11061180: var_70 = ((var_EC - CSgn(var_38)) / 2)
  loc_11061195: var_F0 = var_48.ScaleHeight
  loc_110611D3: If global_110F6000 = 0 Then
  loc_110611DD: Else
  loc_110611E8: End If
  loc_110612F3: frmGzToPzTGZP.Pic1.DispID_80011002(var_70, ((var_F0 - CSgn(var_48)) / 2), CSgn(frmGzToPzTGZP.Pic1.DispID_80010005), CSgn(frmGzToPzTGZP.Pic1.DispID_80010006))
  loc_1106133C: GoTo loc_11061376
End Sub

Public Sub ExitForm(Cancel, UnloadMode) '1104EDE0
  Dim var_18 As Global
  loc_1104EE1F: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1104EE4A: Set var_18 = Me
  loc_1104EE52: var_8008 = Global.Unload
  loc_1104EE8C: GoTo loc_1104EE98
  loc_1104EE97: Exit Sub
  loc_1104EE98: ' Referenced from: 1104EE8C
End Sub

Public Function FillData() '11050810

End Function

Public Function FillDataNew() '11050880
  Dim var_B0 As Variant
  Dim var_64 As Variant
  Dim var_B4 As Variant
  Dim var_58 As Variant
  Dim var_3C As Variant
  Dim var_34 As Me
  Dim var_1A8 As 1
  Dim var_CC As var_C8
  Dim var_C4 As var_C0
  Dim var_98 As Variant
  Dim var_1B0 As Variant
  Dim var_1C4 As var_1C0
  Dim var_2C As ADODB.Recordset
  Dim var_1B8 As 285257256
  Dim var_1A0 As 0
  loc_110509E2: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110509F8: var_2E8 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11050A3C: frmGzToPzTGZP.VFG.DispID_0007 = 1
  loc_11050A5F: Set var_B0 = frmGzToPzTGZP.Label3
  loc_11050A69: var_2C8 = var_B0
  loc_11050A6F: var_B0.Caption = "正在打开Excel数据表，请稍候。。。"
  loc_11050AE2: frmGzToPzTGZP.Pic1.DispID_80010007 = True
  loc_11050B0E: frmGzToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_11050B28: var_8004 = CreateObject(global_1100D5A4)
  loc_11050B33: Set var_64 = CreateObject(global_1100D5A4)
  loc_11050B42: var_B0 = var_64.UnkVCall_000000D0h
  loc_11050BC9: var_2CC = var_B0
  loc_11050E11: Set var_B4 = frmGzToPzTGZP.TDBText
  loc_11050E39: var_90 = var_B4.DispID_0000
  loc_11050E49: var_90 = var_B0.UnkVCall_0000004Ch
  loc_11050EBC: var_B0 = var_58.Tag
  loc_11050F36: var_3C.BackColor = CInt(1)
  loc_11050F5F: var_B0.Activate
  loc_11050FCB: Set var_8C = var_B0.UsedRange
  loc_1105102A: Set var_B0 = frmGzToPzTGZP.Pic1
  loc_11051031: var_B0.DispID_80010007 = var_1A4
  loc_110510A1: var_3C.UnkVCall_00000064h
  loc_11051139: var_DC = var_B0.Cells(5, &H48).value
  loc_11051150: var_90 = Proc_0_11_11029000(var_DC, var_3C, 2)
  loc_11051158: var_8014 = (var_90 = "工会会费")
  loc_1105117D: var_254 = var_8014
  loc_110511BD: var_DC.BackColor = var_1F4
  loc_1105124C: var_FC = var_B4.Cells(5, 73).value
  loc_110512D4: var_FC.BackColor = CInt(1)
  loc_11051400: var_19C = var_8014 Or (LCase(Proc_0_11_11029000(var_FC, var_B4, var_1A8)) <> "餐费扣款") Or (LCase(Proc_0_11_11029000(var_B8.Cells(var_274, 74).value, var_B8, 1)) <> "其他扣款")
  loc_1105149D: If CBool(var_19C) Then
  loc_110514F1:   frmGzToPzTGZP.Pic1.DispID_80010007 = var_1A4
  loc_11051556:   var_C4 = frmGzToPzTGZP.TDBText
  loc_110515DA:   var_1AC = var_58.UnkVCall_0000006Ch
  loc_11051613:   var_1A8 = var_64.UnkVCall_00000398h
  loc_11051648:   Set var_3C = {000208D7-0000-0000-C000000000000046}()
  loc_11051658:   Set var_58 = {000208DA-0000-0000-C000000000000046}()
  loc_11051668:   Set var_64 = {000208D5-0000-0000-C000000000000046}()
  loc_110516FB:   MsgBox("与所要求的格式不符！ ", 64, "提示", 10, 10)
  loc_1105172D: Else
  loc_1105173E:   Set var_B0 = frmGzToPzTGZP.Label3
  loc_1105174C:   var_2C8 = var_B0
  loc_11051752:   var_B0.Caption = "正在填充数据，请稍候。。。"
  loc_110517C9:   frmGzToPzTGZP.Pic1.DispID_80010007 = True
  loc_110517FA:   frmGzToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_11051834:   Set var_B0 = frmGzToPzTGZP.APB
  loc_11051846:   var_2C8 = var_B0
  loc_1105184C:   var_B0.UnkVCall_00000040h
  loc_110518E2:   Set var_B0 = frmGzToPzTGZP.APB
  loc_110518F4:   var_2C8 = var_B0
  loc_110518FA:   var_B0.UnkVCall_00000040h
  loc_110519A4:   frmGzToPzTGZP.APB.UnkVCall_00000040h
  loc_11051A9B:   var_EC = var_8C.Rows.Count - 2
  loc_11051B3D:   frmGzToPzTGZP.sBar.DispID_6803001E(1100D68Ch & var_EC & "条记录")
  loc_11051B88:   var_34 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[T_CY_GzZGZP_Temp]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1) "
  loc_11051B97:   var_8030 = var_34 & "CREATE TABLE [T_CY_GzZGZP_Temp](cCode VARCHAR(50) NULL,cDepCode VARCHAR(50) NULL,cGzItem VARCHAR(50),iMoney Money NULL)"
  loc_11051BE7:   var_CC = UnkObj.UnkVCall_00000040h
  loc_11051C2B:   var_34 = "DELETE FROM [T_CY_GzZGZP_Temp]"
  loc_11051CC8:   Set var_B0 = frmGzToPzTGZP.TDBDate
  loc_11051CF4:   var_D4 = var_B0.DispID_004E
  loc_11051D00:   var_EC)
  loc_11051D5E:   var_74 = CByte("DateToPeriod".00000001h)
  loc_11051DFA:   If var_18 <= CLng(var_8C.Rows.Count) Then
  loc_11051E08:     If global_56 = 0 Then
  loc_11051E78:       var_3C.UnkVCall_00000064h
  loc_11051F1A:       var_8040 = Proc_0_11_11029000(var_B0.Cells(var_18, 14).value, var_3C, 2)
  loc_11051F42:       var_2CC = (vbNull = global_1100AE28) + 1
  loc_11051F7C:       If var_2CC = 0 Then
  loc_11051FA4:         var_8048 = CStr(var_18(-2))
  loc_11051FB2:         Set "正在填充数据：" = CheckObj(var_3C, global_1100D5D4, 100)
  loc_11051FB5:         var_804C = var_B0 & "正在填充数据："
  loc_1105203E:         Set var_B0 = frmGzToPzTGZP.sBar
  loc_11052045:         var_B0.DispID_6803001E(frmGzToPzTGZP.TDBText & "条记录")
  loc_110520E6:         var_3C.UnkVCall_00000064h
  loc_11052188:         var_8054 = Proc_0_11_11029000(var_B0.Cells(var_18, &H61).value, var_3C, 2)
  loc_11052223:         1.BackColor = 1
  loc_110522C5:         var_8058 = Proc_0_12_110291B0(var_B0.Cells(var_18, &H48).value, var_B0, var_1A0)
  loc_110522D2:         Set var_B0 = frmGzToPzTGZP.sBar
  loc_110523B7:         var_805C = Proc_0_12_110291B0(var_B4.Cells(var_1F4, 73).value, var_B4, 1)
  loc_1105240C:         frmGzToPzTGZP.sBar.BackColor = CInt(1)
  loc_110524AA:         var_8060 = Proc_0_12_110291B0(var_B8.Cells(var_18, &H4A).value, var_B8, 0)
  loc_110524EF:         var_B4.BackColor = CInt(1)
  loc_1105258B:         var_8064 = Proc_0_12_110291B0(var_BC.Cells(var_274, 76).value, var_BC, var_1A0)
  loc_11052598:         Set var_CC = var_C8
  loc_110525C8:         var_2F8 = var_A0
  loc_110525D4:         var_300 = var_A8
  loc_110525E2:         var_304 = var_AC
  loc_11052606:         Set var_C4 = var_C0
  loc_11052621:         Set 0000000Ah = var_1B8
  loc_1105264C:         Set 80020004h = var_1B0
  loc_11052677:         Set 00000409h = CheckObj(var_3C, global_1100D5D4, 100)
  loc_11052803:         var_C4.BackColor = CInt(1)
  loc_1105289B:         var_DC = var_B0.Cells(var_18, 14).value
  loc_110528A5:         var_8068 = Proc_0_11_11029000(var_DC, var_B0, 1)
  loc_110528B2:         Set 00000001h = frmGzToPzTGZP.sBar
  loc_110529A2:         var_806C = Proc_0_11_11029000(vbNull.Cells(var_1F4, 15).value, var_B4, 2)
  loc_110529BF:         var_30C = var_9C
  loc_110529E3:         Set var_98 = CheckObj(var_3C, global_1100D5D4, 100)
  loc_11052A00:         var_1B4 = frmGzToPzTGZP.GetKmCode("工资", var_90, var_1B8)
  loc_11052A9A:         var_34 = "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11052AB7:         var_1A4 = var_60
  loc_11052AC8:         var_1B4 = var_24
  loc_11052ACE:         var_8070 = Proc_0_10_11028DD0(&H4008, var_34, var_1C8)
  loc_11052AE4:         var_8074 = 1 & var_1C0
  loc_11052AF6:         var_8078 = vbNull & global_1100AC40
  loc_11052B0A:         var_807C = Proc_0_10_11028DD0(&H4008, 2, var_1E8)
  loc_11052B1A:         var_8080 = 10 & CheckObj(var_3C, global_1100D5D4, 100)
  loc_11052B2C:         var_8084 = -2147352572 & ",'工资',"
  loc_11052B33:         Set  = 
  loc_11052B6F:         var_8088 = CStr(Format((1033 + (global_80020004 + (10 + var_C4))), "0.00"))
  loc_11052B80:         var_808C =  & "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11052B87:         Set  = 
  loc_11052B9E:         var_8090 = var_34 & global_1100BD88
  loc_11052BA5:         Set  = 
  loc_11052D27:         var_DC = var_B0.Cells(var_18, &H4D).value
  loc_11052D31:         var_8094 = Proc_0_12_110291B0(var_DC, var_B0)
  loc_11052D3E:         Set  = 
  loc_11052E6C:         var_DC.BackColor = CInt(1)
  loc_11052F04:         var_DC = var_B0.Cells(var_18, 14).value
  loc_11052F0E:         var_8098 = Proc_0_11_11029000(var_DC, var_B0, 1)
  loc_11052F1B:         Set 00000001h = 
  loc_1105300B:         var_809C = Proc_0_11_11029000(vbNull.Cells(var_1F4, 15).value, var_B4)
  loc_11053018:         Set  = 
  loc_11053028:         var_31C = var_9C
  loc_1105304C:         Set var_98 = 
  loc_11053069:         var_80A0 = = frmGzToPzTGZP.GetKmCode("社保", , )
  loc_11053098:         Set  = 
  loc_11053103:         var_34 = "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11053120:         var_1A4 = var_60
  loc_11053131:         var_1B4 = var_24
  loc_11053137:         var_80A4 = Proc_0_10_11028DD0(&H4008, var_34)
  loc_11053144:         Set  = 
  loc_1105314D:         var_80A8 =  & 0
  loc_11053157:         Set  = 
  loc_1105315F:         var_80AC = 0 & global_1100AC40
  loc_11053169:         Set  = 
  loc_11053173:         var_80B0 = Proc_0_10_11028DD0(&H4008, 0)
  loc_11053180:         Set  = 
  loc_11053183:         var_80B4 =  & 0
  loc_1105318D:         Set  = 
  loc_11053195:         var_80B8 = 0 & ",'社保',"
  loc_1105319C:         Set  = 
  loc_110531D8:         var_80BC = CStr(Format(0, "0.00"))
  loc_110531E9:         var_80C0 =  & "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_110531F0:         Set  = 
  loc_11053207:         var_80C4 = var_34 & global_1100BD88
  loc_1105320E:         Set  = 
  loc_11053390:         var_DC = var_B0.Cells(var_18, &H51).value
  loc_1105339A:         var_80C8 = Proc_0_12_110291B0(var_DC, var_B0)
  loc_110533A7:         Set  = 
  loc_110534D5:         var_DC.BackColor = CInt(1)
  loc_1105356D:         var_DC = var_B0.Cells(var_18, 14).value
  loc_11053577:         var_80CC = Proc_0_11_11029000(var_DC, var_B0, 1)
  loc_11053584:         Set 00000001h = 
  loc_11053674:         var_80D0 = Proc_0_11_11029000(vbNull.Cells(var_1F4, 15).value, var_B4)
  loc_11053681:         Set  = 
  loc_11053691:         var_32C = var_9C
  loc_110536B5:         Set var_98 = 
  loc_110536D2:         var_80D4 = = frmGzToPzTGZP.GetKmCode("公积金", , )
  loc_11053701:         Set  = 
  loc_1105376C:         var_34 = "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11053789:         var_1A4 = var_60
  loc_1105379A:         var_1B4 = var_24
  loc_110537A0:         var_80D8 = Proc_0_10_11028DD0(&H4008, var_34)
  loc_110537AD:         Set  = 
  loc_110537B6:         var_80DC =  & 0
  loc_110537C0:         Set  = 
  loc_110537C8:         var_80E0 = 0 & global_1100AC40
  loc_110537D2:         Set  = 
  loc_110537DC:         var_80E4 = Proc_0_10_11028DD0(&H4008, 0)
  loc_110537E9:         Set  = 
  loc_110537EC:         var_80E8 =  & 0
  loc_110537F6:         Set  = 
  loc_110537FE:         var_80EC = 0 & ",'公积金',"
  loc_11053805:         Set  = 
  loc_11053841:         var_80F0 = CStr(Format(0, "0.00"))
  loc_11053852:         var_80F4 =  & "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11053859:         Set  = 
  loc_11053870:         var_80F8 = var_34 & global_1100BD88
  loc_11053877:         Set  = 
  loc_1105390B:         Set var_B0 = frmGzToPzTGZP.TDBNum
  loc_1105391D:         var_2C8 = var_B0
  loc_11053923:         var_B0.UnkVCall_00000040h
  loc_1105395D:         var_80FC = vbNull.DispID_0043
  loc_110539C0:         If (0 = global_1100AE28) + 1 Then
  loc_110539D5:         Else
  loc_110539E9:           Set var_B0 = frmGzToPzTGZP.TDBNum
  loc_110539FB:           var_2C8 = var_B0
  loc_11053A01:           var_B0.UnkVCall_00000040h
  loc_11053A3B:           var_8104 = vbNull.DispID_0043
  loc_11053A52:           var_44 = var_B0
  loc_11053A86:         End If
  loc_11053B69:         var_8108 = Proc_0_12_110291B0(vbNull.Cells(var_1F4, 37).value, var_B4, var_B4)
  loc_11053B76:         Set var_B4 = 
  loc_11053C13:         var_334 = var_98
  loc_11053C9A:         var_810C = Proc_0_12_110291B0(var_B0.Cells(var_18, &H24).value, var_B0)
  loc_11053CA7:         Set  = 
  loc_11053CC2:         Set  = 
  loc_11053CDD:         var_7C = (0 + 0)
  loc_11053D54:         var_DC = "0.00"
  loc_11053D77:         If global_110F6000 = 0 Then
  loc_11053D81:         Else
  loc_11053D92:         End If
  loc_11053E64:         var_DC.BackColor = CInt(1)
  loc_11053EFC:         var_DC = var_B0.Cells(var_18, 14).value
  loc_11053F06:         var_8110 = Proc_0_11_11029000(var_DC, var_B0, 1)
  loc_11053F13:         Set 00000001h = 
  loc_11054003:         var_8114 = Proc_0_11_11029000(var_B4.Cells(var_1F4, 15).value, var_B4)
  loc_11054010:         Set  = 
  loc_11054020:         var_33C = var_9C
  loc_11054044:         Set var_98 = 
  loc_11054061:         var_8118 = = frmGzToPzTGZP.GetKmCode("奖金", , )
  loc_11054090:         Set  = 
  loc_110540FB:         var_34 = "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11054118:         var_1A4 = var_60
  loc_11054129:         var_1B4 = var_24
  loc_1105412F:         var_811C = Proc_0_10_11028DD0(&H4008, var_34)
  loc_1105413C:         Set  = 
  loc_11054145:         var_8120 =  & 0
  loc_1105414F:         Set  = 
  loc_11054157:         var_8124 = 0 & global_1100AC40
  loc_11054161:         Set  = 
  loc_1105416B:         var_8128 = Proc_0_10_11028DD0(&H4008, 0)
  loc_11054178:         Set  = 
  loc_1105417B:         var_812C =  & 0
  loc_11054185:         Set  = 
  loc_1105418D:         var_8130 = 0 & ",'奖金',"
  loc_11054194:         Set  = 
  loc_110541D0:         var_8134 = CStr(Format(((var_7C * var_44) / 12), var_DC))
  loc_110541E1:         var_8138 =  & "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_110541E8:         Set  = 
  loc_110541FF:         var_813C = var_34 & global_1100BD88
  loc_11054206:         Set  = 
  loc_11054388:         var_DC = var_B0.Cells(var_18, &H48).value
  loc_11054392:         var_8140 = Proc_0_12_110291B0(var_DC, var_B0)
  loc_1105439F:         Set  = 
  loc_1105440B:         var_7C = Format(var_90, "0.00")
  loc_1105456A:         var_DC = var_B0.Cells(var_18, &H48).value
  loc_11054574:         var_8144 = Proc_0_12_110291B0(var_DC, var_B0, 1)
  loc_11054581:         Set 00000001h = 
  loc_1105465C:         var_FC = var_B4.Cells(var_1F4, 73).value
  loc_11054666:         var_8148 = Proc_0_12_110291B0(var_FC, var_B4)
  loc_11054673:         Set  = 
  loc_110546BB:         var_FC.BackColor = CInt(1)
  loc_1105474F:         var_11C = var_B8.Cells(var_18, &H4A).value
  loc_11054759:         var_814C = Proc_0_12_110291B0(var_11C, var_B8)
  loc_11054766:         Set  = 
  loc_1105479E:         var_11C.BackColor = CInt(1)
  loc_1105483A:         var_8150 = Proc_0_12_110291B0(var_BC.Cells(var_274, 76).value, var_BC)
  loc_11054847:         Set  = 
  loc_11054877:         var_348 = var_A0
  loc_11054883:         var_350 = var_A8
  loc_11054891:         var_354 = var_AC
  loc_110548B5:         Set  = 
  loc_110548D0:         Set  = 
  loc_110548FB:         Set  = 
  loc_11054926:         Set  = 
  loc_11054AD5:         var_34 = "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11054AE4:         var_8154 = var_34 & "'550213','1100','工会经费',"
  loc_11054AEF:         Set 00000001h = 1
  loc_11054AFD:         var_8158 = CStr(Format((Format((0 + (0 + (0 + 0))), "0.00") * 0.02), "0.00"))
  loc_11054B0E:         var_815C = 1 & "INSERT INTO [T_CY_GzZGZP_Temp](cCode,cDepCode,cGzItem,iMoney) VALUES ("
  loc_11054B19:         Set 00000001h = 
  loc_11054B30:         var_8160 = var_34 & global_1100BD88
  loc_11054B3B:         Set  = 
  loc_11054BCD:         var_28 = var_28(1)
  loc_11054BD8:         If var_18 Mod 00000064h = 0 Then
  loc_11054BDA:           DoEvents
  loc_11054BE0:         End If
  loc_11054BE0:       End If
  loc_11054BF0:       var_18 = 1+var_18
  loc_11054BF3:       GoTo loc_11051DF4
  loc_11054BF8:     End If
  loc_11054C47:     frmGzToPzTGZP.VFG.DispID_0007 = 1
  loc_11054C5D:     global_56 = 0
  loc_11054C87:     Set var_B0 = frmGzToPzTGZP.APB
  loc_11054C99:     var_2C8 = var_B0
  loc_11054C9F:     var_B0.UnkVCall_00000040h
  loc_11054D38:     Set var_B0 = frmGzToPzTGZP.APB
  loc_11054D4A:     var_2C8 = var_B0
  loc_11054D50:     var_B0.UnkVCall_00000040h
  loc_11054DE9:     Set var_B0 = frmGzToPzTGZP.APB
  loc_11054DFB:     var_2C8 = var_B0
  loc_11054E01:     var_B0.UnkVCall_00000040h
  loc_11054E72:   End If
  loc_11054E7B:   var_1B4 = var_74
  loc_11054EFD:   var_8164 = "cIYear".00000000h & 1100D700h & var_74 & "月应付工资"
  loc_11054F08:   Set 00000002h = var_B4
  loc_11054F54:   var_2C8 = var_2C
  loc_11054F5A:   var_2C4 = ADODB.Recordset.State
  loc_11054F85:   If var_2C4 = 1 Then
  loc_11054FA3:     var_2C8 = var_2C
  loc_11054FA9:     var_8170 = ADODB.Recordset.Close
  loc_11054FCD:   End If
  loc_11055030:   var_2C8 = var_2C
  loc_11055078:   var_8178 = ADODB.Recordset.Open(8, vbNull, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZP_Temp] WHERE cGzItem='工资' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_110550C7:   var_2C8 = var_2C
  loc_110550CD:   var_2C0 = ADODB.Recordset.EOF
  loc_110550F3:   If var_2C0 = 0 Then
  loc_11055101:     var_54 = "1"
  loc_1105515E:     var_8180 = var_54 & Chr(9) & 1100AE28h
  loc_11055214:     var_8188 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_110552B2:     var_818C = var_54 & Chr(9) & 1100D6D4h
  loc_110552BD:     Set 00000001h = -1
  loc_11055336:     var_8190 = var_54 & Chr(9) & 1100C008h
  loc_110553B9:     var_8194 = var_54 & Chr(9) & var_50
  loc_11055427:     var_2C8 = var_2C
  loc_11055461:     var_2D0 = ADODB.Recordset.Fields
  loc_11055499:     ADODB.Recordset.8 = Forms
  loc_11055518:     var_81A0 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_110555A0:     var_2C8 = var_2C
  loc_110555E6:     var_2D0 = ADODB.Recordset.Fields
  loc_11055612:     ADODB.Recordset.8 = Forms
  loc_11055691:     var_81AC = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_1105572F:     var_81B0 = var_54 & Chr(9) & 1100C008h
  loc_1105573A:     Set 00000000h = 0
  loc_110557B3:     var_81B4 = var_54 & Chr(9) & 1100AE28h
  loc_110557BE:     Set  = 
  loc_11055837:     var_81B8 = var_54 & Chr(9) & 1100AE28h
  loc_11055842:     Set  = 
  loc_110558BB:     var_81BC = var_54 & Chr(9) & 1100AE28h
  loc_110558C6:     Set  = 
  loc_1105593F:     var_81C0 = var_54 & Chr(9) & 1100AE28h
  loc_1105594A:     Set  = 
  loc_110559C3:     var_81C4 = var_54 & Chr(9) & 1100AE28h
  loc_110559CE:     Set  = 
  loc_11055A47:     var_81C8 = var_54 & Chr(9) & 1100AE28h
  loc_11055A52:     Set  = 
  loc_11055ACB:     var_81CC = var_54 & Chr(9) & 1100AE28h
  loc_11055AD6:     Set  = 
  loc_11055B39:     var_2C8 = var_2C
  loc_11055B73:     var_2D0 = ADODB.Recordset.Fields
  loc_11055BAB:     ADODB.Recordset.8 = Forms
  loc_11055C2A:     var_81D8 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_11055CC8:     var_81DC = var_54 & Chr(9) & 1100AE28h
  loc_11055CD3:     Set  = 
  loc_11055D4C:     var_81E0 = var_54 & Chr(9) & 1100AE28h
  loc_11055D57:     Set  = 
  loc_11055DD0:     var_81E4 = var_54 & Chr(9) & 1100AE28h
  loc_11055DDB:     Set  = 
  loc_11055E54:     var_81E8 = var_54 & Chr(9) & 1100AE28h
  loc_11055E5F:     Set  = 
  loc_11055ED8:     var_81EC = var_54 & Chr(9) & 1100AE28h
  loc_11055EE3:     Set  = 
  loc_11055F28:     var_2C8 = var_2C
  loc_11055F6E:     var_2D0 = ADODB.Recordset.Fields
  loc_11055F9A:     ADODB.Recordset.8 = Forms
  loc_11055FDF:     var_81F4 = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_11056004:     Set  = 
  loc_11056016:     var_20 = (0 + var_20)
  loc_1105606A:     var_2C8 = var_2C
  loc_11056070:     var_81FC = ADODB.Recordset.MoveNext
  loc_110560E6:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_110560FB:     GoTo loc_110550A4
  loc_11056100:   End If
  loc_11056123:   var_2C8 = var_2C
  loc_1105614E:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_1105615C:     var_54 = "1"
  loc_110561B9:     var_8204 = var_54 & Chr(9) & 1100AE28h
  loc_110561C4:     Set  = 
  loc_1105626F:     var_820C = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_1105627A:     Set  = 
  loc_1105630D:     var_8210 = var_54 & Chr(9) & 1100D6D4h
  loc_11056318:     Set  = 
  loc_11056391:     var_8214 = var_54 & Chr(9) & 1100C008h
  loc_1105639C:     Set  = 
  loc_11056414:     var_8218 = var_54 & Chr(9) & var_50
  loc_1105641F:     Set  = 
  loc_11056498:     var_821C = var_54 & Chr(9) & "215101"
  loc_110564A3:     Set  = 
  loc_1105651C:     var_8220 = var_54 & Chr(9) & 1100C008h
  loc_11056527:     Set  = 
  loc_11056579:     var_1B0 = var_1C
  loc_110565AC:     var_8224 = var_54 & Chr(9) & var_20
  loc_110565B7:     Set  = 
  loc_11056630:     var_8228 = var_54 & Chr(9) & 1100AE28h
  loc_1105663B:     Set  = 
  loc_110566B4:     var_822C = var_54 & Chr(9) & 1100AE28h
  loc_110566BF:     Set  = 
  loc_11056738:     var_8230 = var_54 & Chr(9) & 1100AE28h
  loc_11056743:     Set  = 
  loc_110567BC:     var_8234 = var_54 & Chr(9) & 1100AE28h
  loc_110567C7:     Set  = 
  loc_11056840:     var_8238 = var_54 & Chr(9) & 1100AE28h
  loc_1105684B:     Set  = 
  loc_110568C4:     var_823C = var_54 & Chr(9) & 1100AE28h
  loc_110568CF:     Set  = 
  loc_11056948:     var_8240 = var_54 & Chr(9) & 1100AE28h
  loc_11056953:     Set  = 
  loc_110569CC:     var_8244 = var_54 & Chr(9) & 1100AE28h
  loc_110569D7:     Set  = 
  loc_11056A50:     var_8248 = var_54 & Chr(9) & 1100AE28h
  loc_11056A5B:     Set  = 
  loc_11056AD4:     var_824C = var_54 & Chr(9) & 1100AE28h
  loc_11056ADF:     Set  = 
  loc_11056B58:     var_8250 = var_54 & Chr(9) & 1100AE28h
  loc_11056B63:     Set  = 
  loc_11056BDC:     var_8254 = var_54 & Chr(9) & 1100AE28h
  loc_11056BE7:     Set  = 
  loc_11056C60:     var_8258 = var_54 & Chr(9) & 1100AE28h
  loc_11056C6B:     Set  = 
  loc_11056CD9:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_11056CEE:   End If
  loc_11056D07:   var_1C4 = var_74
  loc_11056D20:   var_1A4 = "计提"
  loc_11056DA1:   var_825C =  & "cIYear".0 & 1100D700h & var_74 & "月社保"
  loc_11056DAC:   Set  = 
  loc_11056DFF:   var_2C8 = var_2C
  loc_11056E05:   var_2C4 = ADODB.Recordset.State
  loc_11056E30:   If var_2C4 = 1 Then
  loc_11056E4E:     var_2C8 = var_2C
  loc_11056E54:     var_8268 = ADODB.Recordset.Close
  loc_11056E78:   End If
  loc_11056EDB:   var_2C8 = var_2C
  loc_11056F23:   var_8270 = ADODB.Recordset.Open(8, vbNull, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZP_Temp] WHERE cGzItem='社保' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_11056F6A:   var_2C8 = var_2C
  loc_11056F70:   var_2C0 = ADODB.Recordset.EOF
  loc_11056F96:   If var_2C0 = 0 Then
  loc_11056FA4:     var_54 = "2"
  loc_11057001:     var_8278 = var_54 & Chr(9) & 1100AE28h
  loc_110570B7:     var_8280 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_11057155:     var_8284 = var_54 & Chr(9) & 1100D6D4h
  loc_11057160:     Set 00000001h = -1
  loc_110571D9:     var_8288 = var_54 & Chr(9) & 1100C008h
  loc_110571E4:     Set  = 
  loc_1105725C:     var_828C = var_54 & Chr(9) & var_50
  loc_11057267:     Set  = 
  loc_110572CA:     var_2C8 = var_2C
  loc_11057304:     var_2D0 = ADODB.Recordset.Fields
  loc_1105733C:     ADODB.Recordset.8 = Forms
  loc_110573BB:     var_8298 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_11057443:     var_2C8 = var_2C
  loc_11057489:     var_2D0 = ADODB.Recordset.Fields
  loc_110574B5:     ADODB.Recordset.8 = Forms
  loc_11057534:     var_82A4 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_110575D2:     var_82A8 = var_54 & Chr(9) & 1100C008h
  loc_110575DD:     Set  = 
  loc_11057656:     var_82AC = var_54 & Chr(9) & 1100AE28h
  loc_11057661:     Set  = 
  loc_110576DA:     var_82B0 = var_54 & Chr(9) & 1100AE28h
  loc_110576E5:     Set  = 
  loc_1105775E:     var_82B4 = var_54 & Chr(9) & 1100AE28h
  loc_11057769:     Set  = 
  loc_110577E2:     var_82B8 = var_54 & Chr(9) & 1100AE28h
  loc_110577ED:     Set  = 
  loc_11057866:     var_82BC = var_54 & Chr(9) & 1100AE28h
  loc_11057871:     Set  = 
  loc_110578EA:     var_82C0 = var_54 & Chr(9) & 1100AE28h
  loc_110578F5:     Set  = 
  loc_1105796E:     var_82C4 = var_54 & Chr(9) & 1100AE28h
  loc_11057979:     Set  = 
  loc_110579DC:     var_2C8 = var_2C
  loc_11057A16:     var_2D0 = ADODB.Recordset.Fields
  loc_11057A4E:     ADODB.Recordset.8 = Forms
  loc_11057ACD:     var_82D0 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_11057B6B:     var_82D4 = var_54 & Chr(9) & 1100AE28h
  loc_11057B76:     Set  = 
  loc_11057BEF:     var_82D8 = var_54 & Chr(9) & 1100AE28h
  loc_11057BFA:     Set  = 
  loc_11057C73:     var_82DC = var_54 & Chr(9) & 1100AE28h
  loc_11057C7E:     Set  = 
  loc_11057CF7:     var_82E0 = var_54 & Chr(9) & 1100AE28h
  loc_11057D02:     Set  = 
  loc_11057D7B:     var_82E4 = var_54 & Chr(9) & 1100AE28h
  loc_11057D86:     Set  = 
  loc_11057DCB:     var_2C8 = var_2C
  loc_11057E11:     var_2D0 = ADODB.Recordset.Fields
  loc_11057E3D:     ADODB.Recordset.8 = Forms
  loc_11057E82:     var_82EC = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_11057EA7:     Set  = 
  loc_11057EB9:     var_20 = (0 + var_20)
  loc_11057F0D:     var_2C8 = var_2C
  loc_11057F13:     var_82F4 = ADODB.Recordset.MoveNext
  loc_11057F89:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_11057F9E:     GoTo loc_11056F47
  loc_11057FA3:   End If
  loc_11057FC6:   var_2C8 = var_2C
  loc_11057FF1:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_11057FFF:     var_54 = "2"
  loc_1105805C:     var_82FC = var_54 & Chr(9) & 1100AE28h
  loc_11058067:     Set  = 
  loc_11058112:     var_8304 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_1105811D:     Set  = 
  loc_110581B0:     var_8308 = var_54 & Chr(9) & 1100D6D4h
  loc_110581BB:     Set  = 
  loc_11058234:     var_830C = var_54 & Chr(9) & 1100C008h
  loc_1105823F:     Set  = 
  loc_110582B7:     var_8310 = var_54 & Chr(9) & var_50
  loc_110582C2:     Set  = 
  loc_1105833B:     var_8314 = var_54 & Chr(9) & "215303"
  loc_11058346:     Set  = 
  loc_110583BF:     var_8318 = var_54 & Chr(9) & 1100C008h
  loc_110583CA:     Set  = 
  loc_1105841C:     var_1B0 = var_1C
  loc_1105844F:     var_831C = var_54 & Chr(9) & var_20
  loc_1105845A:     Set  = 
  loc_110584D3:     var_8320 = var_54 & Chr(9) & 1100AE28h
  loc_110584DE:     Set  = 
  loc_11058557:     var_8324 = var_54 & Chr(9) & 1100AE28h
  loc_11058562:     Set  = 
  loc_110585DB:     var_8328 = var_54 & Chr(9) & 1100AE28h
  loc_110585E6:     Set  = 
  loc_1105865F:     var_832C = var_54 & Chr(9) & 1100AE28h
  loc_1105866A:     Set  = 
  loc_110586E3:     var_8330 = var_54 & Chr(9) & 1100AE28h
  loc_110586EE:     Set  = 
  loc_11058767:     var_8334 = var_54 & Chr(9) & 1100AE28h
  loc_11058772:     Set  = 
  loc_110587EB:     var_8338 = var_54 & Chr(9) & 1100AE28h
  loc_110587F6:     Set  = 
  loc_1105886F:     var_833C = var_54 & Chr(9) & 1100AE28h
  loc_1105887A:     Set  = 
  loc_110588F3:     var_8340 = var_54 & Chr(9) & 1100AE28h
  loc_110588FE:     Set  = 
  loc_11058977:     var_8344 = var_54 & Chr(9) & 1100AE28h
  loc_11058982:     Set  = 
  loc_110589FB:     var_8348 = var_54 & Chr(9) & 1100AE28h
  loc_11058A06:     Set  = 
  loc_11058A7F:     var_834C = var_54 & Chr(9) & 1100AE28h
  loc_11058A8A:     Set  = 
  loc_11058B03:     var_8350 = var_54 & Chr(9) & 1100AE28h
  loc_11058B0E:     Set  = 
  loc_11058B7C:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_11058B91:   End If
  loc_11058BAA:   var_1C4 = var_74
  loc_11058BC3:   var_1A4 = "计提"
  loc_11058C44:   var_8354 =  & "cIYear".0 & 1100D700h & var_74 & "月住房公积金"
  loc_11058C4F:   Set  = 
  loc_11058CA2:   var_2C8 = var_2C
  loc_11058CA8:   var_2C4 = ADODB.Recordset.State
  loc_11058CD3:   If var_2C4 = 1 Then
  loc_11058CF1:     var_2C8 = var_2C
  loc_11058CF7:     var_8360 = ADODB.Recordset.Close
  loc_11058D1B:   End If
  loc_11058D7E:   var_2C8 = var_2C
  loc_11058DC6:   var_8368 = ADODB.Recordset.Open(8, vbNull, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZP_Temp] WHERE cGzItem='公积金' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_11058E0D:   var_2C8 = var_2C
  loc_11058E13:   var_2C0 = ADODB.Recordset.EOF
  loc_11058E39:   If var_2C0 = 0 Then
  loc_11058E47:     var_54 = "3"
  loc_11058EA4:     var_8370 = var_54 & Chr(9) & 1100AE28h
  loc_11058F5A:     var_8378 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_11058FF8:     var_837C = var_54 & Chr(9) & 1100D6D4h
  loc_11059003:     Set 00000001h = -1
  loc_1105907C:     var_8380 = var_54 & Chr(9) & 1100C008h
  loc_11059087:     Set  = 
  loc_110590FF:     var_8384 = var_54 & Chr(9) & var_50
  loc_1105910A:     Set  = 
  loc_1105916D:     var_2C8 = var_2C
  loc_110591A7:     var_2D0 = ADODB.Recordset.Fields
  loc_110591DF:     ADODB.Recordset.8 = Forms
  loc_1105925E:     var_8390 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_110592E6:     var_2C8 = var_2C
  loc_1105932C:     var_2D0 = ADODB.Recordset.Fields
  loc_11059358:     ADODB.Recordset.8 = Forms
  loc_110593D7:     var_839C = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_11059475:     var_83A0 = var_54 & Chr(9) & 1100C008h
  loc_11059480:     Set  = 
  loc_110594F9:     var_83A4 = var_54 & Chr(9) & 1100AE28h
  loc_11059504:     Set  = 
  loc_1105957D:     var_83A8 = var_54 & Chr(9) & 1100AE28h
  loc_11059588:     Set  = 
  loc_11059601:     var_83AC = var_54 & Chr(9) & 1100AE28h
  loc_1105960C:     Set  = 
  loc_11059685:     var_83B0 = var_54 & Chr(9) & 1100AE28h
  loc_11059690:     Set  = 
  loc_11059709:     var_83B4 = var_54 & Chr(9) & 1100AE28h
  loc_11059714:     Set  = 
  loc_1105978D:     var_83B8 = var_54 & Chr(9) & 1100AE28h
  loc_11059798:     Set  = 
  loc_11059811:     var_83BC = var_54 & Chr(9) & 1100AE28h
  loc_1105981C:     Set  = 
  loc_1105987F:     var_2C8 = var_2C
  loc_110598B9:     var_2D0 = ADODB.Recordset.Fields
  loc_110598F1:     ADODB.Recordset.8 = Forms
  loc_11059970:     var_83C8 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_11059A0E:     var_83CC = var_54 & Chr(9) & 1100AE28h
  loc_11059A19:     Set  = 
  loc_11059A92:     var_83D0 = var_54 & Chr(9) & 1100AE28h
  loc_11059A9D:     Set  = 
  loc_11059B16:     var_83D4 = var_54 & Chr(9) & 1100AE28h
  loc_11059B21:     Set  = 
  loc_11059B9A:     var_83D8 = var_54 & Chr(9) & 1100AE28h
  loc_11059BA5:     Set  = 
  loc_11059C1E:     var_83DC = var_54 & Chr(9) & 1100AE28h
  loc_11059C29:     Set  = 
  loc_11059C6E:     var_2C8 = var_2C
  loc_11059CB4:     var_2D0 = ADODB.Recordset.Fields
  loc_11059CE0:     ADODB.Recordset.8 = Forms
  loc_11059D25:     var_83E4 = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_11059D4A:     Set  = 
  loc_11059D5C:     var_20 = (0 + var_20)
  loc_11059DB0:     var_2C8 = var_2C
  loc_11059DB6:     var_83EC = ADODB.Recordset.MoveNext
  loc_11059E2C:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_11059E41:     GoTo loc_11058DEA
  loc_11059E46:   End If
  loc_11059E69:   var_2C8 = var_2C
  loc_11059E94:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_11059EA2:     var_54 = "3"
  loc_11059EFF:     var_83F4 = var_54 & Chr(9) & 1100AE28h
  loc_11059F0A:     Set  = 
  loc_11059FB5:     var_83FC = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_11059FC0:     Set  = 
  loc_1105A053:     var_8400 = var_54 & Chr(9) & 1100D6D4h
  loc_1105A05E:     Set  = 
  loc_1105A0D7:     var_8404 = var_54 & Chr(9) & 1100C008h
  loc_1105A0E2:     Set  = 
  loc_1105A15A:     var_8408 = var_54 & Chr(9) & var_50
  loc_1105A165:     Set  = 
  loc_1105A1DE:     var_840C = var_54 & Chr(9) & "217601"
  loc_1105A1E9:     Set  = 
  loc_1105A262:     var_8410 = var_54 & Chr(9) & 1100C008h
  loc_1105A26D:     Set  = 
  loc_1105A2BF:     var_1B0 = var_1C
  loc_1105A2F2:     var_8414 = var_54 & Chr(9) & var_20
  loc_1105A2FD:     Set  = 
  loc_1105A376:     var_8418 = var_54 & Chr(9) & 1100AE28h
  loc_1105A381:     Set  = 
  loc_1105A3FA:     var_841C = var_54 & Chr(9) & 1100AE28h
  loc_1105A405:     Set  = 
  loc_1105A47E:     var_8420 = var_54 & Chr(9) & 1100AE28h
  loc_1105A489:     Set  = 
  loc_1105A502:     var_8424 = var_54 & Chr(9) & 1100AE28h
  loc_1105A50D:     Set  = 
  loc_1105A586:     var_8428 = var_54 & Chr(9) & 1100AE28h
  loc_1105A591:     Set  = 
  loc_1105A60A:     var_842C = var_54 & Chr(9) & 1100AE28h
  loc_1105A615:     Set  = 
  loc_1105A68E:     var_8430 = var_54 & Chr(9) & 1100AE28h
  loc_1105A699:     Set  = 
  loc_1105A712:     var_8434 = var_54 & Chr(9) & 1100AE28h
  loc_1105A71D:     Set  = 
  loc_1105A796:     var_8438 = var_54 & Chr(9) & 1100AE28h
  loc_1105A7A1:     Set  = 
  loc_1105A81A:     var_843C = var_54 & Chr(9) & 1100AE28h
  loc_1105A825:     Set  = 
  loc_1105A89E:     var_8440 = var_54 & Chr(9) & 1100AE28h
  loc_1105A8A9:     Set  = 
  loc_1105A922:     var_8444 = var_54 & Chr(9) & 1100AE28h
  loc_1105A92D:     Set  = 
  loc_1105A9A6:     var_8448 = var_54 & Chr(9) & 1100AE28h
  loc_1105A9B1:     Set  = 
  loc_1105AA1F:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_1105AA34:   End If
  loc_1105AA4D:   var_1C4 = var_74
  loc_1105AA66:   var_1A4 = "计提"
  loc_1105AAE7:   var_844C =  & "cIYear".0 & 1100D700h & var_74 & "月中方人员奖金"
  loc_1105AAF2:   Set  = 
  loc_1105AB45:   var_2C8 = var_2C
  loc_1105AB4B:   var_2C4 = ADODB.Recordset.State
  loc_1105AB76:   If var_2C4 = 1 Then
  loc_1105AB94:     var_2C8 = var_2C
  loc_1105AB9A:     var_8458 = ADODB.Recordset.Close
  loc_1105ABBE:   End If
  loc_1105AC21:   var_2C8 = var_2C
  loc_1105AC69:   var_8460 = ADODB.Recordset.Open(8, vbNull, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZP_Temp] WHERE cGzItem='奖金' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_1105ACB0:   var_2C8 = var_2C
  loc_1105ACB6:   var_2C0 = ADODB.Recordset.EOF
  loc_1105ACDC:   If var_2C0 = 0 Then
  loc_1105ACEA:     var_54 = "4"
  loc_1105AD47:     var_8468 = var_54 & Chr(9) & 1100AE28h
  loc_1105ADFD:     var_8470 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_1105AE9B:     var_8474 = var_54 & Chr(9) & 1100D6D4h
  loc_1105AEA6:     Set 00000001h = -1
  loc_1105AF1F:     var_8478 = var_54 & Chr(9) & 1100C008h
  loc_1105AF2A:     Set  = 
  loc_1105AFA2:     var_847C = var_54 & Chr(9) & var_50
  loc_1105AFAD:     Set  = 
  loc_1105B010:     var_2C8 = var_2C
  loc_1105B04A:     var_2D0 = ADODB.Recordset.Fields
  loc_1105B082:     ADODB.Recordset.8 = Forms
  loc_1105B101:     var_8488 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_1105B189:     var_2C8 = var_2C
  loc_1105B1CF:     var_2D0 = ADODB.Recordset.Fields
  loc_1105B1FB:     ADODB.Recordset.8 = Forms
  loc_1105B27A:     var_8494 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_1105B318:     var_8498 = var_54 & Chr(9) & 1100C008h
  loc_1105B323:     Set  = 
  loc_1105B39C:     var_849C = var_54 & Chr(9) & 1100AE28h
  loc_1105B3A7:     Set  = 
  loc_1105B420:     var_84A0 = var_54 & Chr(9) & 1100AE28h
  loc_1105B42B:     Set  = 
  loc_1105B4A4:     var_84A4 = var_54 & Chr(9) & 1100AE28h
  loc_1105B4AF:     Set  = 
  loc_1105B528:     var_84A8 = var_54 & Chr(9) & 1100AE28h
  loc_1105B533:     Set  = 
  loc_1105B5AC:     var_84AC = var_54 & Chr(9) & 1100AE28h
  loc_1105B5B7:     Set  = 
  loc_1105B630:     var_84B0 = var_54 & Chr(9) & 1100AE28h
  loc_1105B63B:     Set  = 
  loc_1105B6B4:     var_84B4 = var_54 & Chr(9) & 1100AE28h
  loc_1105B6BF:     Set  = 
  loc_1105B722:     var_2C8 = var_2C
  loc_1105B75C:     var_2D0 = ADODB.Recordset.Fields
  loc_1105B794:     ADODB.Recordset.8 = Forms
  loc_1105B813:     var_84C0 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_1105B8B1:     var_84C4 = var_54 & Chr(9) & 1100AE28h
  loc_1105B8BC:     Set  = 
  loc_1105B935:     var_84C8 = var_54 & Chr(9) & 1100AE28h
  loc_1105B940:     Set  = 
  loc_1105B9B9:     var_84CC = var_54 & Chr(9) & 1100AE28h
  loc_1105B9C4:     Set  = 
  loc_1105BA3D:     var_84D0 = var_54 & Chr(9) & 1100AE28h
  loc_1105BA48:     Set  = 
  loc_1105BAC1:     var_84D4 = var_54 & Chr(9) & 1100AE28h
  loc_1105BACC:     Set  = 
  loc_1105BB11:     var_2C8 = var_2C
  loc_1105BB57:     var_2D0 = ADODB.Recordset.Fields
  loc_1105BB83:     ADODB.Recordset.8 = Forms
  loc_1105BBC8:     var_84DC = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_1105BBED:     Set  = 
  loc_1105BBFF:     var_20 = (0 + var_20)
  loc_1105BC53:     var_2C8 = var_2C
  loc_1105BC59:     var_84E4 = ADODB.Recordset.MoveNext
  loc_1105BCCF:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_1105BCE4:     GoTo loc_1105AC8D
  loc_1105BCE9:   End If
  loc_1105BD0C:   var_2C8 = var_2C
  loc_1105BD37:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_1105BD45:     var_54 = "4"
  loc_1105BDA2:     var_84EC = var_54 & Chr(9) & 1100AE28h
  loc_1105BDAD:     Set  = 
  loc_1105BE58:     var_84F4 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_1105BE63:     Set  = 
  loc_1105BEF6:     var_84F8 = var_54 & Chr(9) & 1100D6D4h
  loc_1105BF01:     Set  = 
  loc_1105BF7A:     var_84FC = var_54 & Chr(9) & 1100C008h
  loc_1105BF85:     Set  = 
  loc_1105BFFD:     var_8500 = var_54 & Chr(9) & var_50
  loc_1105C008:     Set  = 
  loc_1105C081:     var_8504 = var_54 & Chr(9) & "215107"
  loc_1105C08C:     Set  = 
  loc_1105C105:     var_8508 = var_54 & Chr(9) & 1100C008h
  loc_1105C110:     Set  = 
  loc_1105C162:     var_1B0 = var_1C
  loc_1105C195:     var_850C = var_54 & Chr(9) & var_20
  loc_1105C1A0:     Set  = 
  loc_1105C219:     var_8510 = var_54 & Chr(9) & 1100AE28h
  loc_1105C224:     Set  = 
  loc_1105C29D:     var_8514 = var_54 & Chr(9) & 1100AE28h
  loc_1105C2A8:     Set  = 
  loc_1105C321:     var_8518 = var_54 & Chr(9) & 1100AE28h
  loc_1105C32C:     Set  = 
  loc_1105C3A5:     var_851C = var_54 & Chr(9) & 1100AE28h
  loc_1105C3B0:     Set  = 
  loc_1105C429:     var_8520 = var_54 & Chr(9) & 1100AE28h
  loc_1105C434:     Set  = 
  loc_1105C4AD:     var_8524 = var_54 & Chr(9) & 1100AE28h
  loc_1105C4B8:     Set  = 
  loc_1105C531:     var_8528 = var_54 & Chr(9) & 1100AE28h
  loc_1105C53C:     Set  = 
  loc_1105C5B5:     var_852C = var_54 & Chr(9) & 1100AE28h
  loc_1105C5C0:     Set  = 
  loc_1105C639:     var_8530 = var_54 & Chr(9) & 1100AE28h
  loc_1105C644:     Set  = 
  loc_1105C6BD:     var_8534 = var_54 & Chr(9) & 1100AE28h
  loc_1105C6C8:     Set  = 
  loc_1105C741:     var_8538 = var_54 & Chr(9) & 1100AE28h
  loc_1105C74C:     Set  = 
  loc_1105C7C5:     var_853C = var_54 & Chr(9) & 1100AE28h
  loc_1105C7D0:     Set  = 
  loc_1105C849:     var_8540 = var_54 & Chr(9) & 1100AE28h
  loc_1105C854:     Set  = 
  loc_1105C8C2:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_1105C8D7:   End If
  loc_1105C8F0:   var_1C4 = var_74
  loc_1105C909:   var_1A4 = "计提"
  loc_1105C98A:   var_8544 =  & "cIYear".0 & 1100D700h & var_74 & "月中方人员有薪假工资"
  loc_1105C995:   Set  = 
  loc_1105C9E8:   var_2C8 = var_2C
  loc_1105C9EE:   var_2C4 = ADODB.Recordset.State
  loc_1105CA19:   If var_2C4 = 1 Then
  loc_1105CA37:     var_2C8 = var_2C
  loc_1105CA3D:     var_8550 = ADODB.Recordset.Close
  loc_1105CA61:   End If
  loc_1105CAC4:   var_2C8 = var_2C
  loc_1105CB0C:   var_8558 = ADODB.Recordset.Open(8, vbNull, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZP_Temp] WHERE cGzItem='有薪假工资' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_1105CB53:   var_2C8 = var_2C
  loc_1105CB59:   var_2C0 = ADODB.Recordset.EOF
  loc_1105CB7F:   If var_2C0 = 0 Then
  loc_1105CB8D:     var_54 = "5"
  loc_1105CBEA:     var_8560 = var_54 & Chr(9) & 1100AE28h
  loc_1105CCA0:     var_8568 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_1105CD3E:     var_856C = var_54 & Chr(9) & 1100D6D4h
  loc_1105CD49:     Set 00000001h = -1
  loc_1105CDC2:     var_8570 = var_54 & Chr(9) & 1100C008h
  loc_1105CDCD:     Set  = 
  loc_1105CE45:     var_8574 = var_54 & Chr(9) & var_50
  loc_1105CE50:     Set  = 
  loc_1105CEB3:     var_2C8 = var_2C
  loc_1105CEED:     var_2D0 = ADODB.Recordset.Fields
  loc_1105CF25:     ADODB.Recordset.8 = Forms
  loc_1105CFA4:     var_8580 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_1105D02C:     var_2C8 = var_2C
  loc_1105D072:     var_2D0 = ADODB.Recordset.Fields
  loc_1105D09E:     ADODB.Recordset.8 = Forms
  loc_1105D11D:     var_858C = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_1105D1BB:     var_8590 = var_54 & Chr(9) & 1100C008h
  loc_1105D1C6:     Set  = 
  loc_1105D23F:     var_8594 = var_54 & Chr(9) & 1100AE28h
  loc_1105D24A:     Set  = 
  loc_1105D2C3:     var_8598 = var_54 & Chr(9) & 1100AE28h
  loc_1105D2CE:     Set  = 
  loc_1105D347:     var_859C = var_54 & Chr(9) & 1100AE28h
  loc_1105D352:     Set  = 
  loc_1105D3CB:     var_85A0 = var_54 & Chr(9) & 1100AE28h
  loc_1105D3D6:     Set  = 
  loc_1105D44F:     var_85A4 = var_54 & Chr(9) & 1100AE28h
  loc_1105D45A:     Set  = 
  loc_1105D4D3:     var_85A8 = var_54 & Chr(9) & 1100AE28h
  loc_1105D4DE:     Set  = 
  loc_1105D557:     var_85AC = var_54 & Chr(9) & 1100AE28h
  loc_1105D562:     Set  = 
  loc_1105D5C5:     var_2C8 = var_2C
  loc_1105D5FF:     var_2D0 = ADODB.Recordset.Fields
  loc_1105D637:     ADODB.Recordset.8 = Forms
  loc_1105D6B6:     var_85B8 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_1105D754:     var_85BC = var_54 & Chr(9) & 1100AE28h
  loc_1105D75F:     Set  = 
  loc_1105D7D8:     var_85C0 = var_54 & Chr(9) & 1100AE28h
  loc_1105D7E3:     Set  = 
  loc_1105D85C:     var_85C4 = var_54 & Chr(9) & 1100AE28h
  loc_1105D867:     Set  = 
  loc_1105D8E0:     var_85C8 = var_54 & Chr(9) & 1100AE28h
  loc_1105D8EB:     Set  = 
  loc_1105D964:     var_85CC = var_54 & Chr(9) & 1100AE28h
  loc_1105D96F:     Set  = 
  loc_1105D9B4:     var_2C8 = var_2C
  loc_1105D9FA:     var_2D0 = ADODB.Recordset.Fields
  loc_1105DA26:     ADODB.Recordset.8 = Forms
  loc_1105DA6B:     var_85D4 = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_1105DA90:     Set  = 
  loc_1105DAA2:     var_20 = (0 + var_20)
  loc_1105DAF6:     var_2C8 = var_2C
  loc_1105DAFC:     var_85DC = ADODB.Recordset.MoveNext
  loc_1105DB72:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_1105DB87:     GoTo loc_1105CB30
  loc_1105DB8C:   End If
  loc_1105DBAF:   var_2C8 = var_2C
  loc_1105DBDA:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_1105DBE8:     var_54 = "5"
  loc_1105DC45:     var_85E4 = var_54 & Chr(9) & 1100AE28h
  loc_1105DC50:     Set  = 
  loc_1105DCFB:     var_85EC = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_1105DD06:     Set  = 
  loc_1105DD99:     var_85F0 = var_54 & Chr(9) & 1100D6D4h
  loc_1105DDA4:     Set  = 
  loc_1105DE1D:     var_85F4 = var_54 & Chr(9) & 1100C008h
  loc_1105DE28:     Set  = 
  loc_1105DEA0:     var_85F8 = var_54 & Chr(9) & var_50
  loc_1105DEAB:     Set  = 
  loc_1105DF24:     var_85FC = var_54 & Chr(9) & "215109"
  loc_1105DF2F:     Set  = 
  loc_1105DFA8:     var_8600 = var_54 & Chr(9) & 1100C008h
  loc_1105DFB3:     Set  = 
  loc_1105E005:     var_1B0 = var_1C
  loc_1105E038:     var_8604 = var_54 & Chr(9) & var_20
  loc_1105E043:     Set  = 
  loc_1105E0BC:     var_8608 = var_54 & Chr(9) & 1100AE28h
  loc_1105E0C7:     Set  = 
  loc_1105E140:     var_860C = var_54 & Chr(9) & 1100AE28h
  loc_1105E14B:     Set  = 
  loc_1105E1C4:     var_8610 = var_54 & Chr(9) & 1100AE28h
  loc_1105E1CF:     Set  = 
  loc_1105E248:     var_8614 = var_54 & Chr(9) & 1100AE28h
  loc_1105E253:     Set  = 
  loc_1105E2CC:     var_8618 = var_54 & Chr(9) & 1100AE28h
  loc_1105E2D7:     Set  = 
  loc_1105E350:     var_861C = var_54 & Chr(9) & 1100AE28h
  loc_1105E35B:     Set  = 
  loc_1105E3D4:     var_8620 = var_54 & Chr(9) & 1100AE28h
  loc_1105E3DF:     Set  = 
  loc_1105E458:     var_8624 = var_54 & Chr(9) & 1100AE28h
  loc_1105E463:     Set  = 
  loc_1105E4DC:     var_8628 = var_54 & Chr(9) & 1100AE28h
  loc_1105E4E7:     Set  = 
  loc_1105E560:     var_862C = var_54 & Chr(9) & 1100AE28h
  loc_1105E56B:     Set  = 
  loc_1105E5E4:     var_8630 = var_54 & Chr(9) & 1100AE28h
  loc_1105E5EF:     Set  = 
  loc_1105E668:     var_8634 = var_54 & Chr(9) & 1100AE28h
  loc_1105E673:     Set  = 
  loc_1105E6EC:     var_8638 = var_54 & Chr(9) & 1100AE28h
  loc_1105E6F7:     Set  = 
  loc_1105E765:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_1105E77A:   End If
  loc_1105E793:   var_1C4 = var_74
  loc_1105E7AC:   var_1A4 = "计提"
  loc_1105E82D:   var_863C =  & "cIYear".0 & 1100D700h & var_74 & "月工会经费"
  loc_1105E838:   Set  = 
  loc_1105E88B:   var_2C8 = var_2C
  loc_1105E891:   var_2C4 = ADODB.Recordset.State
  loc_1105E8BC:   If var_2C4 = 1 Then
  loc_1105E8DA:     var_2C8 = var_2C
  loc_1105E8E0:     var_8648 = ADODB.Recordset.Close
  loc_1105E904:   End If
  loc_1105E967:   var_2C8 = var_2C
  loc_1105E9AF:   var_8650 = ADODB.Recordset.Open(8, vbNull, "SELECT ccode,cdepcode,sum(iMoney) AS iMoney FROM [T_CY_GzZGZP_Temp] WHERE cGzItem='工会经费' GROUP BY ccode,cdepcode", var_1A0, 9)
  loc_1105E9F6:   var_2C8 = var_2C
  loc_1105E9FC:   var_2C0 = ADODB.Recordset.EOF
  loc_1105EA22:   If var_2C0 = 0 Then
  loc_1105EA30:     var_54 = "6"
  loc_1105EA8D:     var_8658 = var_54 & Chr(9) & 1100AE28h
  loc_1105EB43:     var_8660 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_1105EBE1:     var_8664 = var_54 & Chr(9) & 1100D6D4h
  loc_1105EBEC:     Set 00000001h = -1
  loc_1105EC65:     var_8668 = var_54 & Chr(9) & 1100C008h
  loc_1105EC70:     Set  = 
  loc_1105ECE8:     var_866C = var_54 & Chr(9) & var_50
  loc_1105ECF3:     Set  = 
  loc_1105ED56:     var_2C8 = var_2C
  loc_1105ED90:     var_2D0 = ADODB.Recordset.Fields
  loc_1105EDC8:     ADODB.Recordset.8 = Forms
  loc_1105EE47:     var_8678 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cCode")
  loc_1105EECF:     var_2C8 = var_2C
  loc_1105EF15:     var_2D0 = ADODB.Recordset.Fields
  loc_1105EF41:     ADODB.Recordset.8 = Forms
  loc_1105EFC0:     var_8684 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "iMoney")
  loc_1105F05E:     var_8688 = var_54 & Chr(9) & 1100C008h
  loc_1105F069:     Set  = 
  loc_1105F0E2:     var_868C = var_54 & Chr(9) & 1100AE28h
  loc_1105F0ED:     Set  = 
  loc_1105F166:     var_8690 = var_54 & Chr(9) & 1100AE28h
  loc_1105F171:     Set  = 
  loc_1105F1EA:     var_8694 = var_54 & Chr(9) & 1100AE28h
  loc_1105F1F5:     Set  = 
  loc_1105F26E:     var_8698 = var_54 & Chr(9) & 1100AE28h
  loc_1105F279:     Set  = 
  loc_1105F2F2:     var_869C = var_54 & Chr(9) & 1100AE28h
  loc_1105F2FD:     Set  = 
  loc_1105F376:     var_86A0 = var_54 & Chr(9) & 1100AE28h
  loc_1105F381:     Set  = 
  loc_1105F3FA:     var_86A4 = var_54 & Chr(9) & 1100AE28h
  loc_1105F405:     Set  = 
  loc_1105F468:     var_2C8 = var_2C
  loc_1105F4A2:     var_2D0 = ADODB.Recordset.Fields
  loc_1105F4DA:     ADODB.Recordset.8 = Forms
  loc_1105F559:     var_86B0 = var_54 & Chr(9) & Proc_0_11_11029000(9, var_1B8, "cDepCode")
  loc_1105F5F7:     var_86B4 = var_54 & Chr(9) & 1100AE28h
  loc_1105F602:     Set  = 
  loc_1105F67B:     var_86B8 = var_54 & Chr(9) & 1100AE28h
  loc_1105F686:     Set  = 
  loc_1105F6FF:     var_86BC = var_54 & Chr(9) & 1100AE28h
  loc_1105F70A:     Set  = 
  loc_1105F783:     var_86C0 = var_54 & Chr(9) & 1100AE28h
  loc_1105F78E:     Set  = 
  loc_1105F807:     var_86C4 = var_54 & Chr(9) & 1100AE28h
  loc_1105F812:     Set  = 
  loc_1105F857:     var_2C8 = var_2C
  loc_1105F89D:     var_2D0 = ADODB.Recordset.Fields
  loc_1105F8C9:     ADODB.Recordset.8 = Forms
  loc_1105F8F4:     var_B4 = 0
  loc_1105F8FE:     var_C4 = var_B4
  loc_1105F90E:     var_86CC = Proc_0_12_110291B0(9, var_1A8, "iMoney")
  loc_1105F933:     Set  = 
  loc_1105F945:     var_20 = (0 + var_20)
  loc_1105F999:     var_2C8 = var_2C
  loc_1105F99F:     var_86D4 = ADODB.Recordset.MoveNext
  loc_1105FA15:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_1105FA2A:     GoTo loc_1105E9D3
  loc_1105FA2F:   End If
  loc_1105FA52:   var_2C8 = var_2C
  loc_1105FA7D:   If ADODB.Recordset.RecordCount >= 1 Then
  loc_1105FA8B:     var_54 = "6"
  loc_1105FAE8:     var_86DC = var_54 & Chr(9) & 1100AE28h
  loc_1105FAF3:     Set  = 
  loc_1105FB9E:     var_86E4 = var_54 & Chr(9) & frmGzToPzTGZP.TDBDate.DispID_004E
  loc_1105FBA9:     Set  = 
  loc_1105FC3C:     var_86E8 = var_54 & Chr(9) & 1100D6D4h
  loc_1105FC47:     Set  = 
  loc_1105FCC0:     var_86EC = var_54 & Chr(9) & 1100C008h
  loc_1105FCCB:     Set  = 
  loc_1105FD43:     var_86F0 = var_54 & Chr(9) & var_50
  loc_1105FD4E:     Set  = 
  loc_1105FDC7:     var_86F4 = var_54 & Chr(9) & "219101"
  loc_1105FDD2:     Set  = 
  loc_1105FE4B:     var_86F8 = var_54 & Chr(9) & 1100C008h
  loc_1105FE56:     Set  = 
  loc_1105FEA8:     var_1B0 = var_1C
  loc_1105FEDB:     var_86FC = var_54 & Chr(9) & var_20
  loc_1105FEE6:     Set  = 
  loc_1105FF5F:     var_8700 = var_54 & Chr(9) & 1100AE28h
  loc_1105FF6A:     Set  = 
  loc_1105FFE3:     var_8704 = var_54 & Chr(9) & 1100AE28h
  loc_1105FFEE:     Set  = 
  loc_11060067:     var_8708 = var_54 & Chr(9) & 1100AE28h
  loc_11060072:     Set  = 
  loc_110600EB:     var_870C = var_54 & Chr(9) & 1100AE28h
  loc_110600F6:     Set  = 
  loc_1106016F:     var_8710 = var_54 & Chr(9) & 1100AE28h
  loc_1106017A:     Set  = 
  loc_110601F3:     var_8714 = var_54 & Chr(9) & 1100AE28h
  loc_110601FE:     Set  = 
  loc_11060277:     var_8718 = var_54 & Chr(9) & 1100AE28h
  loc_11060282:     Set  = 
  loc_110602FB:     var_871C = var_54 & Chr(9) & 1100AE28h
  loc_11060306:     Set  = 
  loc_1106037F:     var_8720 = var_54 & Chr(9) & 1100AE28h
  loc_1106038A:     Set  = 
  loc_11060403:     var_8724 = var_54 & Chr(9) & 1100AE28h
  loc_1106040E:     Set  = 
  loc_11060487:     var_8728 = var_54 & Chr(9) & 1100AE28h
  loc_11060492:     Set  = 
  loc_1106050B:     var_872C = var_54 & Chr(9) & 1100AE28h
  loc_11060516:     Set  = 
  loc_1106058F:     var_8730 = var_54 & Chr(9) & 1100AE28h
  loc_1106059A:     Set  = 
  loc_11060608:     frmGzToPzTGZP.VFG.DispID_0080(var_54)
  loc_1106061D:   End If
  loc_11060637:   var_8734 = CStr(var_28)
  loc_11060645:   Set "有效数据共" = 
  loc_1106064E:   var_8738 =  & "有效数据共"
  loc_11060658:   Set  = 
  loc_110606D6:   frmGzToPzTGZP.sBar.DispID_6803001E(0 & global_1100FE7C)
  loc_11060742:   frmGzToPzTGZP.APB.UnkVCall_00000040h
  loc_110607D4:   Set var_B0 = frmGzToPzTGZP.APB
  loc_110607E2:   var_2C8 = var_B0
  loc_110607E8:   var_B0.UnkVCall_00000040h
  loc_1106087A:   Set var_B0 = frmGzToPzTGZP.APB
  loc_11060888:   var_2C8 = var_B0
  loc_1106088E:   var_B0.UnkVCall_00000040h
  loc_1106093E:   frmGzToPzTGZP.Pic1.DispID_80010007 = var_1A4
  loc_110609A2:   var_C4 = frmGzToPzTGZP.TDBText
  loc_11060A1D:   var_1AC = var_58.UnkVCall_0000006Ch
  loc_11060A56:   var_1A8 = var_64.UnkVCall_00000398h
  loc_11060A8B:   Set var_3C = {000208D7-0000-0000-C000000000000046}()
  loc_11060A9B:   Set var_58 = {000208DA-0000-0000-C000000000000046}()
  loc_11060AAB:   Set var_64 = {000208D5-0000-0000-C000000000000046}()
  loc_11060ABC: End If
  loc_11060AC2: GoTo loc_11060B99
  loc_11060B98: Exit Function
  loc_11060B99: ' Referenced from: 11060AC2
End Function

Public Function getWBHL(sWhere) '11072020
  Dim var_1C As ADODB.Recordset
  Dim var_2C As Me
  loc_11072080: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_1107208C: var_98 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110720B4: var_40 = Trim(sWhere)
  loc_110720E5: If (var_40 <> 1100AE28h) Then
  loc_11072113:   var_20 = "SELECT * FROM exch WHERE 1=1 " & " AND " & sWhere
  loc_11072120: Else
  loc_1107212C: End If
  loc_1107213C: var_20 = var_20 & " order by cexch_name, itype, iperiod, cdate"
  loc_110721A6: var_78 = var_1C
  loc_110721B5: var_8018 = ADODB.Recordset.Open(var_20, var_5C, var_20, var_54, 9)
  loc_1107221B: If ADODB.Recordset.EOF Then
  loc_1107222A:   var_24 = CStr(0)
  loc_11072235: Else
  loc_11072257:   var_2C = ADODB.Recordset.Fields
  loc_11072284:   var_58 = "NFLAT"
  loc_1107229D:   ADODB.Recordset.8 = Forms
  loc_110722EE:   var_24 = var_40
  loc_11072310: End If
  loc_1107232E: var_8030 = ADODB.Recordset.Close
  loc_1107234D: GoTo loc_1107238B
  loc_11072353: If var_4 Then
  loc_1107235E: End If
  loc_1107238A: Exit Function
  loc_1107238B: ' Referenced from: 1107234D
End Function

Public Function GetKmCode(pGZ_Item, pGZ_Type, pGZ_Type1) '11073660
  Dim var_34 As ADODB.Recordset
  Dim var_50 As Me
  loc_110736DE: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110736E6: var_B4 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11073708: var_30 = pGZ_Type1
  loc_11073711: On Error GoTo loc_11073A0B
  loc_1107371A: var_6C = pGZ_Item
  loc_11073728: var_7C = pGZ_Type
  loc_11073793: var_8018 = Proc_0_10_11028DD0(&H4008,  & Proc_0_10_11028DD0(&H4008, 1 & "SELECT * FROM " & "..T_CY_GZ_TGZP_KmSetting WHERE GZ_Item=", ) & " AND GZ_Type=", )
  loc_110737A7: var_20 =  & var_8018
  loc_1107384A: var_8024 = ADODB.Recordset.Open(var_20, var_70, var_20, var_68, 9)
  loc_110738A7: var_88 = ADODB.Recordset.EOF
  loc_110738C6: If var_88 = 0 Then
  loc_110738E7:   var_50 = ADODB.Recordset.Fields
  loc_11073905:   var_6C = "UF_KMCode"
  loc_1107392D:   ADODB.Recordset.8 = Forms
  loc_1107397E:   var_2C = var_64
  loc_110739C8:   If ADODB.Recordset.Close < 0 Then
  loc_110739D6:     var_8038 = CheckObj(var_34, global_1100ADFC, 128)
  loc_110739DA:   End If
  loc_110739FB:   If ADODB.Recordset.Close < 0 Then
  loc_11073A0B:   End If
  loc_11073A1D:   Set var_34 = ADODB.Recordset()
  loc_11073A31: End If
  loc_11073A31: Exit Sub
  loc_11073A3C: GoTo loc_11073A8A
  loc_11073A42: If var_C Then
  loc_11073A4D: End If
  loc_11073A89: Exit Function
  loc_11073A8A: ' Referenced from: 11073A3C
End Function

Public Function GetUFDepCode(pDepCode, pType) '11073AE0
  Dim var_30 As ADODB.Recordset
  Dim var_4C As Me
  loc_11073B58: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11073B60: var_B0 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11073B72: var_2C = pDepCode
  loc_11073B7A: var_20 = pType
  loc_11073B83: On Error GoTo loc_11073EA5
  loc_11073B92: var_8004 = (var_20 = "高科技")
  loc_11073B9A: If var_8004 = 0 Then
  loc_11073BA4:   var_28 = "2000"
  loc_11073BAB: Else
  loc_11073BAE:   var_68 = var_2C
  loc_11073BBC:   var_78 = var_20
  loc_11073C21:   var_801C = Proc_0_10_11028DD0(var_80,  & Proc_0_10_11028DD0(&H4008, 1 & "SELECT * FROM " & "..T_CY_GZ_SL_DepSetting WHERE GZ_Dep=", ) & " AND GZ_Type=", )
  loc_11073C35:   var_24 =  & var_801C
  loc_11073CD5:   var_8028 = ADODB.Recordset.Open(var_24, var_6C, var_24, var_64, 9)
  loc_11073D32:   var_84 = ADODB.Recordset.EOF
  loc_11073D4E:   If var_84 = 0 Then
  loc_11073D72:     var_4C = ADODB.Recordset.Fields
  loc_11073D90:     var_68 = "UF_DepCode"
  loc_11073DB8:     ADODB.Recordset.8 = Forms
  loc_11073E09:     var_28 = var_60
  loc_11073E53:     If ADODB.Recordset.Close < 0 Then
  loc_11073E61:       var_803C = CheckObj(var_30, global_1100ADFC, 128)
  loc_11073E65:     End If
  loc_11073E6B:     var_28 = var_2C
  loc_11073E95:     If ADODB.Recordset.Close < 0 Then
  loc_11073EA5:     End If
  loc_11073EB7:     Set var_30 = ADODB.Recordset()
  loc_11073EC3:     var_28 = var_2C
  loc_11073EC9:   End If
  loc_11073EC9: End If
  loc_11073EC9: ' Referenced from: 11073E63
  loc_11073EC9: Exit Sub
  loc_11073ED4: GoTo loc_11073F22
  loc_11073EDA: If var_C Then
  loc_11073EE5: End If
  loc_11073F21: Exit Function
  loc_11073F22: ' Referenced from: 11073ED4
End Function

Public Function getBTData() '11073F70
  Dim var_24 As ADODB.Recordset
  Dim var_38 As Variant
  loc_11073FF4: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11073FFE: On Error GoTo loc_11074587
  loc_11074039: var_28 = 1 & "IF NOT EXISTS (SELECT * FROM [" & "]..Sysobjects "
  loc_110740A4: var_8018 =  & var_28 & "WHERE Name = 'T_CY_GZ_ZGZP_Setting') " & "CREATE TABLE [" & "]..[T_CY_GZ_ZGZP_Setting](fXS1 FLOAT NULL," & "fXS2 FLOAT NULL)"
  loc_110740AB: var_28 = var_8018
  loc_110740DB: var_54 = UnkObj.UnkVCall_00000040h
  loc_1107412D: var_28 = var_38 & "SELECT * FROM [" & "]..[T_CY_GZ_ZGZP_Setting]"
  loc_11074167: var_BC = ADODB.Recordset.State
  loc_1107418C: If var_BC = 1 Then
  loc_110741A8:   var_802C = ADODB.Recordset.Close
  loc_110741C6: End If
  loc_11074246: var_8034 = ADODB.Recordset.Open(var_28, var_90, var_28, var_88, 9)
  loc_11074299: var_B8 = ADODB.Recordset.EOF
  loc_110742B5: If var_B8 = 0 Then
  loc_110742DD:   var_38 = ADODB.Recordset.Fields
  loc_110742FB:   var_8C = "fXS1"
  loc_1107432F:   ADODB.Recordset.8 = Forms
  loc_1107439A:   frmGzToPzTGZP.TDBNum.UnkVCall_00000040h
  loc_110743D0:   var_44.DispID_0000 = Proc_0_12_110291B0(9, var_90, "fXS1")
  loc_11074436:   var_D0 = ADODB.Recordset.Fields
  loc_11074441:   var_8C = "fXS1"
  loc_11074475:   ADODB.Recordset.8 = Forms
  loc_110744E3:   frmGzToPzTGZP.TDBNum.UnkVCall_00000040h
  loc_11074519:   var_44.DispID_0000 = Proc_0_12_110291B0(9, var_90, "fXS1")
  loc_11074546: End If
  loc_1107456E: If ADODB.Recordset.Close < 0 Then
  loc_11074580:   var_8050 = CheckObj(var_24, global_1100ADFC, 128)
  loc_11074587:   ' Referenced from: 11073FFE
  loc_1107458C:   var_8054 = Err
  loc_11074597:   Set var_38 = Err
  loc_1107461C:   MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_11074649: End If
  loc_11074649: Exit Sub
  loc_11074654: GoTo loc_1107469D
  loc_1107469C: Exit Function
  loc_1107469D: ' Referenced from: 11074654
End Function

Public Function UpdateBTData() '110746E0
  Dim var_3C As Variant
  Dim var_44 As frmGzToPzTGZP.TDBNum
  Dim var_20 As Me
  loc_11074758: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_11074762: On Error GoTo loc_11074A2B
  loc_1107479D: var_20 = 1 & "DELETE FROM [" & "]..[T_CY_GZ_ZGZP_Setting]"
  loc_110747D6: var_58 = UnkObj.UnkVCall_00000040h
  loc_11074856: Set var_3C = frmGzToPzTGZP.TDBNum
  loc_1107485C: var_BC = var_3C
  loc_1107486B: var_3C.UnkVCall_00000040h
  loc_110748A5: var_60 = var_40.DispID_0043
  loc_110748C0: Set var_44 = frmGzToPzTGZP.TDBNum
  loc_110748C6: var_C4 = var_44
  loc_110748D5: var_44.UnkVCall_00000040h
  loc_1107490F: var_80 = var_48.DispID_0043
  loc_11074940: var_8028 = 1 & Proc_0_12_110291B0(8, var_3C & "INSERT INTO [" & "]..[T_CY_GZ_ZGZP_Setting]" & "(fXS1,fXS2) VALUES (", var_44) & global_1100AC40
  loc_11074974: var_20 = var_3C & Proc_0_12_110291B0(8, var_8028, var_48) & global_1100BD88
  loc_11074A26: GoTo loc_11074AED
  loc_11074A2B: ' Referenced from: 11074762
  loc_11074A30: var_8038 = Err
  loc_11074A3B: Set var_3C = Err
  loc_11074AC0: MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_11074AED: ' Referenced from: 11074A26
  loc_11074AED: Exit Sub
  loc_11074AF8: GoTo loc_11074B4D
  loc_11074B4C: Exit Function
  loc_11074B4D: ' Referenced from: 11074AF8
End Function

Private Sub Proc_12_12_1104EEC0
  Dim var_58 As frmGzToPzTGZP.VFG
  loc_1104EF01: Set var_58 = frmGzToPzTGZP.VFG
  loc_1104EF52: var_58.DispID_005D = frmGzToPzTGZP.VFG
  loc_1104EF93: var_58.DispID_0067 = frmGzToPzTGZP.VFG
  loc_1104EFB2: var_58.DispID_0041 = frmGzToPzTGZP.VFG
  loc_1104F05C: var_58.DispID_00A5("...")
  loc_1104F184: var_58.DispID_008A(4)
  loc_1104F1C7: var_58.DispID_0079(450)
  loc_1104F1EB: var_58.DispID_0019 = True
  loc_1104F22B: var_58.DispID_007B(True)
  loc_1104F274: var_58.DispID_009D(5)
  loc_1104F2B9: var_58.DispID_0090("业务号")
  loc_1104F2FC: var_58.DispID_0077(4)
  loc_1104F33F: var_58.DispID_0078(700)
  loc_1104F387: var_58.DispID_0090("状态")
  loc_1104F3CD: var_58.DispID_0077(4)
  loc_1104F413: var_58.DispID_0078(700)
  loc_1104F45B: var_58.DispID_0090("制单日期")
  loc_1104F4A1: var_58.DispID_0077(1)
  loc_1104F4E7: var_58.DispID_0078(1000)
  loc_1104F52C: var_58.DispID_0090("凭证类别字")
  loc_1104F56E: var_58.DispID_0077(4)
  loc_1104F5B0: var_58.DispID_0078(700)
  loc_1104F5F8: var_58.DispID_0090("附单据数")
  loc_1104F63C: var_58.DispID_0077(var_3C)
  loc_1104F682: var_58.DispID_0078(var_3C)
  loc_1104F6CA: var_58.DispID_0090(var_3C)
  loc_1104F710: var_58.DispID_0077(var_3C)
  loc_1104F756: var_58.DispID_0078(var_3C)
  loc_1104F79E: var_58.DispID_0090(var_3C)
  loc_1104F7E4: var_58.DispID_0077(var_3C)
  loc_1104F82A: var_58.DispID_0078(var_3C)
  loc_1104F872: var_58.DispID_0090(var_3C)
  loc_1104F8B6: var_58.DispID_0077(var_3C)
  loc_1104F8FC: var_58.DispID_0078(var_3C)
  loc_1104F944: var_58.DispID_009C(var_3C)
  loc_1104F98C: var_58.DispID_0090(var_3C)
  loc_1104F9D2: var_58.DispID_0077(var_3C)
  loc_1104FA18: var_58.DispID_0078(var_3C)
  loc_1104FA60: var_58.DispID_009C(var_3C)
  loc_1104FAA8: var_58.DispID_0090(var_3C)
  loc_1104FAEE: var_58.DispID_0077(var_3C)
  loc_1104FB34: var_58.DispID_0078(var_3C)
  loc_1104FB7C: var_58.DispID_009C(var_3C)
  loc_1104FBC4: var_58.DispID_0090(var_3C)
  loc_1104FC0A: var_58.DispID_0077(var_3C)
  loc_1104FC50: var_58.DispID_0078(var_3C)
  loc_1104FC98: var_58.DispID_009C(var_3C)
  loc_1104FCE0: var_58.DispID_0090(var_3C)
  loc_1104FD26: var_58.DispID_0077(var_3C)
  loc_1104FD6C: var_58.DispID_0078(var_3C)
  loc_1104FDB4: var_58.DispID_009C(var_3C)
  loc_1104FDFC: var_58.DispID_0090(var_3C)
  loc_1104FE42: var_58.DispID_0077(var_3C)
  loc_1104FE88: var_58.DispID_0078(var_3C)
  loc_1104FED0: var_58.DispID_0090(var_3C)
  loc_1104FF16: var_58.DispID_0077(var_3C)
  loc_1104FF5C: var_58.DispID_0078(var_3C)
  loc_1104FFA4: var_58.DispID_0090(var_3C)
  loc_1104FFEA: var_58.DispID_0077(var_3C)
  loc_11050030: var_58.DispID_0078(var_3C)
  loc_11050078: var_58.DispID_0090(var_3C)
  loc_110500BE: var_58.DispID_0077(var_3C)
  loc_11050104: var_58.DispID_0078(var_3C)
  loc_1105014C: var_58.DispID_0090(var_3C)
  loc_11050192: var_58.DispID_0077(var_3C)
  loc_110501D8: var_58.DispID_0078(var_3C)
  loc_11050220: var_58.DispID_0090(var_3C)
  loc_11050266: var_58.DispID_0077(var_3C)
  loc_110502AC: var_58.DispID_0078(var_3C)
  loc_110502F4: var_58.DispID_0090(var_3C)
  loc_1105033A: var_58.DispID_0077(var_3C)
  loc_11050380: var_58.DispID_0078(var_3C)
  loc_110503C8: var_58.DispID_0090(var_3C)
  loc_1105040E: var_58.DispID_0077(var_3C)
  loc_11050454: var_58.DispID_0078(var_3C)
  loc_1105049C: var_58.DispID_0090(var_3C)
  loc_110504E2: var_58.DispID_0077(var_3C)
  loc_11050528: var_58.DispID_0078(var_3C)
  loc_11050570: var_58.DispID_0090(var_3C)
  loc_110505B6: var_58.DispID_0077(var_3C)
  loc_110505FC: var_58.DispID_0078(var_3C)
  loc_11050644: var_58.DispID_0090(var_3C)
  loc_1105068A: var_58.DispID_0077(var_3C)
  loc_110506D0: var_58.DispID_0078(var_3C)
  loc_110506EC: If 9 <= &H15 Then
  loc_1105072C:   var_58.DispID_00AC(var_3C)
  loc_11050744:   var_14 = 1+var_14
  loc_11050747:   GoTo loc_110506E8
  loc_11050749: End If
  loc_11050788: var_58.DispID_00AC(var_3C)
  loc_110507CE: var_58.DispID_00AC(var_3C)
End Sub

Private Sub Proc_12_13_110613A0
  Dim var_7C As Variant
  Dim var_1F8 As Label
  Dim var_80 As Variant
  Dim var_88 As frmGzToPzTGZP.Label3
  loc_1106148A: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11061492: var_228 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11061498: var_8004 = ecx
  loc_1106150E: If var_14 <= CLng(frmGzToPzTGZP.VFG.DispID_0007)(-1) Then
  loc_1106151F:   var_800C = frmGzToPzTGZP.Proc_12_14_11063240(vbNull)
  loc_110615BD:   frmGzToPzTGZP.VFG.DispID_0082(22, var_58)
  loc_110616A1:   If (frmGzToPzTGZP.VFG.DispID_0082(var_14, 22) = global_1100AE28) + 1 Then
  loc_11061721:     frmGzToPzTGZP.VFG.DispID_0082(1, 285267764)
  loc_11061855:     frmGzToPzTGZP.VFG.DispID_009E(var_14, 1, var_14, 1, 16711680)
  loc_11061875:     Set var_7C = frmGzToPzTGZP.Label3
  loc_11061882:     var_1F8 = var_7C
  loc_110618CC:     var_7C.Caption = "分析: 第(" & CStr(vbNull) & ")行信息----有效"
  loc_1106191E:     frmGzToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_11061931:   Else
  loc_110619AB:     frmGzToPzTGZP.VFG.DispID_0082(1, 285267820)
  loc_11061ADF:     frmGzToPzTGZP.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_11061AFF:     Set var_80 = frmGzToPzTGZP.Label3
  loc_11061B0C:     var_1F8 = var_80
  loc_11061BED:     var_80.Caption = "分析:   第(" & CStr(vbNull) & ")行信息----" & frmGzToPzTGZP.VFG.DispID_0082(var_14, 22)
  loc_11061C58:     frmGzToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_11061C6A:   End If
  loc_11061C7A:   var_14 = 1+var_14
  loc_11061C7D:   GoTo loc_11061500
  loc_11061C82: End If
  loc_11061CE9: If var_14 <= CLng(frmGzToPzTGZP.VFG.DispID_0007)(-1) Then
  loc_11061D61:   var_A0 = frmGzToPzTGZP.VFG.DispID_0082(var_14, 2)
  loc_11061D7F:   var_B8)
  loc_11061F0F:   var_8048 = frmGzToPzTGZP.VFG.DispID_0082(var_14, frmGzToPzTGZP.VFG)
  loc_11061F46:   var_4C = CCur(0)
  loc_11061F49:   var_48 = var_8048
  loc_11061F55:   var_40 = CCur(0)
  loc_11061F58:   var_3C = var_8048
  loc_11061F64:   var_34 = var_14
  loc_11061F6D:   var_30 = var_14
  loc_11061F76:   var_160 = CByte("DateToPeriod".00000001h)
  loc_11062013:   var_B8)
  loc_11062092:   Set var_80 = frmGzToPzTGZP.VFG
  loc_110620B8:   var_8064 = (frmGzToPzTGZP.VFG.DispID_0082(var_14, 3) = var_80.DispID_0082(var_14, 3))
  loc_110620E5:   var_1A0 = var_8064 + 1
  loc_1106215F:   var_806C = (var_8048 = frmGzToPzTGZP.VFG.DispID_0082(var_14, ""))
  loc_11062186:   var_1E0 = var_806C + 1
  loc_11062288:   If CBool((frmGzToPzTGZP.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_8064 + 1 And var_806C + 1) Then
  loc_1106234D:     If (frmGzToPzTGZP.VFG.DispID_0082(var_14, 22) = global_1100AE28) Then
  loc_11062356:     End If
  loc_1106235B:     If var_24 = 0 Then
  loc_11062404:       var_16C = var_48
  loc_11062448:       var_9C = var_1F0
  loc_11062494:       var_4C = CCur(var_4C + Format(Val(frmGzToPzTGZP.VFG.DispID_0082(var_14, 7)), "#.00"))
  loc_11062497:       var_48 = var_D8
  loc_11062577:       var_16C = var_3C
  loc_110625BB:       var_9C = var_1F0
  loc_11062607:       var_40 = CCur(var_40 + Format(Val(frmGzToPzTGZP.VFG.DispID_0082(var_14, 8)), "#.00"))
  loc_1106260A:       var_3C = var_D8
  loc_1106264A:     End If
  loc_1106266B:     var_14 = var_14(1)
  loc_1106266E:     var_30 = var_30(1)
  loc_11062690:     var_80A0 = CLng(frmGzToPzTGZP.VFG.DispID_0007)
  loc_110626AB:     var_1F8 = (var_14 > 0)
  loc_110626CF:     If var_1F8 = 0 Then GoTo loc_11061F70
  loc_110626D5:   End If
  loc_110626DA:   If var_24 = 0 Then
  loc_110626EE:     Set var_7C = frmGzToPzTGZP.Chk
  loc_110626F9:     var_1F8 = var_7C
  loc_110626FF:     Set var_80 = var_7C(1)
  loc_1106272A:     var_200 = var_80
  loc_11062730:     var_1EC = var_80.Value
  loc_11062784:     If (var_1EC = 1) Then
  loc_110627B4:       If (Abs(var_4C - var_40) <> 0.01) >= 0 Then
  loc_110627BD:       End If
  loc_110627BD:     End If
  loc_110627C2:     If var_24 Then
  loc_110627C8:     End If
  loc_110627E8:     var_1C = var_34
  loc_110627ED:     If var_34 <= (var_30 - 1) Then
  loc_110628B1:       If (frmGzToPzTGZP.VFG.DispID_0082(var_1C, 22) = global_1100AE28) + 1 Then
  loc_11062939:         frmGzToPzTGZP.VFG.DispID_0082(1, 285267820)
  loc_110629CD:         frmGzToPzTGZP.VFG.DispID_0082(22, "凭证借贷不平衡或某分录有错误")
  loc_11062B01:         frmGzToPzTGZP.VFG.DispID_009E(var_1C, 1, var_1C, 1, 255)
  loc_11062B13:       End If
  loc_11062B23:       GoTo loc_110627E2
  loc_11062B28:     End If
  loc_11062B39:     var_44 = var_44(1)
  loc_11062B4A:     Set var_88 = frmGzToPzTGZP.Label3
  loc_11062B7D:     var_1F8 = var_88
  loc_11062C8A:     Set var_80 = frmGzToPzTGZP.VFG
  loc_11062D62:     var_80D4 = "分析: 第[" & frmGzToPzTGZP.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmGzToPzTGZP.VFG.DispID_0082(var_34, var_14)
  loc_11062D84:     var_78 = var_80D4 & "]号凭证借贷不平衡"
  loc_11062D98:     var_88.Caption = var_78
  loc_11062D9F:     If var_78 < 0 Then
  loc_11062DA5:       GoTo loc_11063023
  loc_11062DAA:     End If
  loc_11062DBB:     var_20 = var_20(1)
  loc_11062DCC:     Set var_88 = frmGzToPzTGZP.Label3
  loc_11062DFF:     var_1F8 = var_88
  loc_11062F0C:     Set var_80 = frmGzToPzTGZP.VFG
  loc_11062FE4:     var_80F8 = "分析: 第[" & frmGzToPzTGZP.VFG.DispID_0082(var_34, 2) & " - " & var_80.DispID_0082(var_34, 3) & " - " & frmGzToPzTGZP.VFG.DispID_0082(var_34, frmGzToPzTGZP.VFG.DispID_0082(var_34, var_14))
  loc_11063006:     var_78 = var_80F8 & "]号凭证有效"
  loc_1106301A:     var_88.Caption = var_78
  loc_11063021:     If var_78 >= 0 Then GoTo loc_11063032
  loc_11063023:     ' Referenced from: 11062DA5
  loc_1106302C:     var_78 = CheckObj(var_1F8, global_1100D574, 84)
  loc_11063032:   End If
  loc_110630B4:   frmGzToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_110630E5:   var_14 = 1+var_14(-1)
  loc_110630E8:   GoTo loc_11061CE3
  loc_110630ED: End If
  loc_110630F2: If var_44 > 0 Then
  loc_110630F9:   If var_20 > 0 Then
  loc_11063114:   Else
  loc_1106312D:   Else
  loc_11063137:     var_8108 = frmGzToPzTGZP.Proc_12_16_110723D0(var_1EC)
  loc_11063145:     If var_1EC Then
  loc_11063160:     Else
  loc_11063168:       var_18 = ecx
  loc_11063171:       GoTo loc_1106320B
  loc_1106320A:       Exit Sub
  loc_1106320B:     End If
  loc_1106320B:   End If
  loc_1106320B: End If
  loc_1106320B: ' Referenced from: 11063171
End Sub

Private  Proc_12_14_11063240(arg_C) '11063240
  Dim var_58 As frmGzToPzTGZP.VFG
  Dim var_20 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA19-EB96-11D2-AF06000021009B21}()
  Dim var_18 As {3302AA41-EB96-11D2-AF06000021009B21}()
  Dim var_1C As {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_1106333C: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1106334C: var_210 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1106342B: If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 2) = global_1100AE28) + 1 Then
  loc_11063435:   var_24 = "制单日期为空"
  loc_11063446: Else
  loc_110634E1:   var_78 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 2)
  loc_1106351B:   If Proc_0_9_11028500(var_80, global_1106884D, ) Then
  loc_110635C4:     var_78 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 2)
  loc_110635CE:     var_90)
  loc_110635E0:     var_48 = var_90
  loc_11063612:     var_118 = var_48
  loc_11063620:     var_114 = var_44
  loc_11063654:     var_80 = "AccountOpen".0.0
  loc_11063685:     If (var_80 < var_80) Then
  loc_1106368F:       var_24 = "日期超前总账系统启用日期"
  loc_110636A0:     Else
  loc_110636A6:       var_154 = var_44
  loc_110636AC:       var_1A4 = var_44
  loc_110636B8:       var_158 = var_48
  loc_110636BF:       var_1A8 = var_48
  loc_1106376C:       var_80 = "AccountYMD".0.00000002h("AccountYMD".0, var_13C)
  loc_11063866:       If CBool( Or ((global_1106884D < var_80) > "AccountYMD".0.00000002h(var_180, var_18C))) Then
  loc_11063870:         var_24 = "日期必须在当前会计年度内"
  loc_11063881:       Else
  loc_1106389E:         var_118 = var_48
  loc_110638F2:         var_80 = "DateToPeriod".00000001h - 1
  loc_11063980:         If CBool("AccountYMD".0.00000001h) Then
  loc_1106398A:           var_24 = "已结账月份不能制单"
  loc_1106399B:         Else
  loc_11063A77:           If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 3) = global_1100AE28) + 1 Then
  loc_11063A81:             var_24 = "凭证类别字为空"
  loc_11063A92:           Else
  loc_11063B21:             var_8034 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 3)
  loc_11063B31:             var_80 = 8
  loc_11063B34:             var_78 = var_8034
  loc_11063B7B:             var_8038 = CBool(Not("pzlbCheck".00000001h(, fs:[00000000h], , global_1106884D, global_1106884D, var_74, var_8034, var_7C)))
  loc_11063BB2:             If var_8038 Then
  loc_11063BBC:               var_24 = "凭证类别字非法"
  loc_11063BCD:             Else
  loc_11063CA4:               If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, var_128) = global_1100AE28) + 1 Then
  loc_11063CAE:                 var_24 = "业务号为空"
  loc_11063CBF:               Else
  loc_11063D49:                 var_8044 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, var_128)
  loc_11063D59:                 var_80 = 8
  loc_11063D5C:                 var_78 = var_8044
  loc_11063D9F:                 var_90 = "GenLen".00000001h(fs:[00000000h], , global_1106884D, global_1106884D, global_1106884D, var_74, var_8044, var_7C)
  loc_11063DE7:                 If (var_90 > 30) Then
  loc_11063DF1:                   var_24 = "业务号超长"
  loc_11063E02:                 Else
  loc_11063EE1:                   If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 5) = global_1100AE28) + 1 Then
  loc_11063EEB:                     var_24 = "摘要为空"
  loc_11063EFC:                   Else
  loc_11063FB7:                     var_8058 = InStr(1, frmGzToPzTGZP.VFG.DispID_0082(arg_C, 5), "|", 0)
  loc_11063FDD:                     var_220 = (var_8058 > 0)
  loc_11064033:                     var_80 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 5)
  loc_11064154:                     If (((var_8058 > 0) Or (InStr(1, var_80, """", 0) > 0)) Or (InStr(1, frmGzToPzTGZP.VFG.DispID_0082(arg_C, 5), "'", 0) > 0)) Then
  loc_1106415E:                       var_24 = "摘要含有非法字符"
  loc_1106416F:                     Else
  loc_11064201:                       var_806C = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 5)
  loc_11064211:                       var_80 = 8
  loc_11064214:                       var_78 = var_806C
  loc_11064257:                       var_90 = "GenLen".00000001h(global_1106884D, global_1106884D, global_1106884D, global_1106884D, global_1106884D, var_74, var_806C, var_7C)
  loc_110642A0:                       If (var_90 > 120) Then
  loc_110642AA:                         var_24 = "摘要超长"
  loc_110642BB:                       Else
  loc_11064398:                         If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 6) = global_1100AE28) + 1 Then
  loc_110643A2:                           var_24 = "科目为空"
  loc_110643B3:                         Else
  loc_11064442:                           var_807C = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 6)
  loc_11064452:                           var_80 = 8
  loc_11064455:                           var_78 = var_807C
  loc_110644D5:                           var_40 = "kmCheck".00000002h(var_807C, var_150, var_15C)
  loc_11064507:                           var_8084 = (var_40 = global_1100AE28)
  loc_1106450F:                           If var_8084 = 0 Then
  loc_11064519:                             var_24 = "科目非法"
  loc_1106452A:                           Else
  loc_11064568:                             var_118 = arg_C
  loc_110645CF:                             frmGzToPzTGZP.VFG.DispID_0082(6, var_40)
  loc_110645E9:                             var_118 = var_40
  loc_1106463B:                             var_128 = var_20
  loc_11064689:                             "kmCodeToProperties".00000002h
  loc_110646A6:                             Set var_20 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_110646D4:                             var_1F0 = var_20
  loc_110646DA:                             var_1D4 = var_20.UnkVCall_00000114h
  loc_11064706:                             If var_1D4 = 0 Then
  loc_11064710:                               var_24 = "科目非末级"
  loc_11064721:                             Else
  loc_110647FF:                               If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 7) = global_1100AE28) Then
  loc_110648DB:                                 If Not (IsNumeric(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 7))) Then
  loc_110648E5:                                   var_24 = "借方金额非法"
  loc_110648F6:                                 Else
  loc_1106499F:                                   var_80A4 = CDbl(Val(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 7)))
  loc_11064A3A:                                   var_80 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 7)
  loc_11064A62:                                   var_22C = CDbl(Val(var_80))
  loc_11064A78:                                   var_80B0 = CDbl(-9999999999999.99)
  loc_11064A90:                                   GoTo loc_11064A94
  loc_11064AE2:                                   If (eax Or 0) Then
  loc_11064AEC:                                     var_24 = "借方金额超范围"
  loc_11064AFD:                                   Else
  loc_11064AFD:                                   End If
  loc_11064BDB:                                   If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 8) = global_1100AE28) Then
  loc_11064CB7:                                     If Not (IsNumeric(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 8))) Then
  loc_11064CC1:                                       var_24 = "贷方金额非法"
  loc_11064CD2:                                     Else
  loc_11064D7B:                                       var_80C8 = CDbl(Val(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 8)))
  loc_11064E16:                                       var_80 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 8)
  loc_11064E3E:                                       var_238 = CDbl(Val(var_80))
  loc_11064E54:                                       var_80D4 = CDbl(-9999999999999.99)
  loc_11064E6C:                                       GoTo loc_11064E70
  loc_11064EBE:                                       If (eax Or 0) Then
  loc_11064EC8:                                         var_24 = "贷方金额超范围"
  loc_11064ED9:                                       Else
  loc_11064ED9:                                       End If
  loc_11065051:                                       var_74 = var_1E0
  loc_110650C3:                                       var_C4 = var_1E8
  loc_1106513D:                                       var_80E8 = (Format(Val(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 7)), "#.00") <> 0) And (Format(Val(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 8)), "#.00") <> 0)
  loc_110651B6:                                       If CBool(var_80E8) Then
  loc_110651C0:                                         var_24 = "借方金额和贷方金额不能同时不为0"
  loc_110651D1:                                       Else
  loc_11065349:                                         var_74 = var_1E0
  loc_110653BB:                                         var_C4 = var_1E8
  loc_11065435:                                         var_8100 = (Format(Val(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 7)), "#.00") = 0) And (Format(Val(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 8)), "#.00") = 0)
  loc_110654AE:                                         If CBool(var_8100) Then
  loc_110654B8:                                           var_24 = "借方金额和贷方金额不能同时为0"
  loc_110654C9:                                         Else
  loc_110654E9:                                           var_1F0 = var_20
  loc_1106553B:                                           If (var_20.UnkVCall_0000007Ch = global_1100AE28) Then
  loc_1106561F:                                             If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 9) = global_1100AE28) Then
  loc_110656FB:                                               If Not (IsNumeric(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 9))) Then
  loc_11065705:                                                 var_24 = "数量数值非法"
  loc_11065716:                                               Else
  loc_11065716:                                               End If
  loc_11065716:                                             End If
  loc_11065736:                                             var_1F0 = var_20
  loc_11065788:                                             If (var_20.UnkVCall_0000006Ch = global_1100AE28) Then
  loc_1106586C:                                               If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 10) = global_1100AE28) Then
  loc_11065948:                                                 If Not (IsNumeric(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 10))) Then
  loc_11065952:                                                   var_24 = "外币金额非法"
  loc_11065963:                                                 Else
  loc_11065A0C:                                                   var_813C = CDbl(Val(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 10)))
  loc_11065ACF:                                                   var_244 = CDbl(Val(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 10)))
  loc_11065AE5:                                                   var_8148 = CDbl(-9999999999999.99)
  loc_11065AFD:                                                   GoTo loc_11065B01
  loc_11065B4F:                                                   If (eax Or 0) Then
  loc_11065B59:                                                     var_24 = "外币超范围"
  loc_11065B6A:                                                   Else
  loc_11065B6A:                                                   End If
  loc_11065C48:                                                   If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 11) = global_1100AE28) Then
  loc_11065D24:                                                     If Not (IsNumeric(frmGzToPzTGZP.VFG.DispID_0082(arg_C, 11))) Then
  loc_11065D2E:                                                       var_24 = "汇率数值非法"
  loc_11065D3F:                                                     Else
  loc_11065D3F:                                                     End If
  loc_11065D3F:                                                   End If
  loc_11065E1D:                                                   If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 12) = global_1100AE28) Then
  loc_11065EB4:                                                     var_8164 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 12)
  loc_11065EC7:                                                     var_78 = var_8164
  loc_11065F0A:                                                     var_90 = "GenLen".00000001h(global_1106884D, global_1106884D, global_1106884D, global_1106884D, global_1106884D, var_74, var_8164, var_7C)
  loc_11065F24:                                                     var_1F0 = (var_90 > 20)
  loc_11065F53:                                                     If var_1F0 = 0 Then GoTo loc_11066099
  loc_11065F61:                                                     var_24 = "制单人姓名超长"
  loc_11065F72:                                                   Else
  loc_11065F91:                                                     var_118 = arg_C
  loc_1106606D:                                                     frmGzToPzTGZP.VFG.DispID_0082(12, "UserCurrent".00000000h.00000000h)
  loc_110660BC:                                                     var_1F0 = var_20
  loc_110660EE:                                                     If var_20.UnkVCall_0000010Ch Then
  loc_110661D2:                                                       If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 13) = global_1100AE28) Then
  loc_11066269:                                                         var_817C = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 13)
  loc_1106627C:                                                         var_78 = var_817C
  loc_110662AB:                                                         var_90 = "JsfsCheck".00000001h(1, global_1106884D, global_1106884D, global_1106884D, global_1106884D, var_74, var_817C, var_7C)
  loc_110662FB:                                                         If CBool(Not(var_90)) Then
  loc_11066305:                                                           var_24 = "结算方式非法"
  loc_11066316:                                                         Else
  loc_11066316:                                                         End If
  loc_11066316:                                                       End If
  loc_11066339:                                                       var_1F0 = var_20
  loc_1106633F:                                                       var_1D4 = var_20.UnkVCall_0000010Ch
  loc_11066386:                                                       var_1F8 = var_20
  loc_1106638C:                                                       var_1D8 = var_20.UnkVCall_00000094h
  loc_110663D3:                                                       var_200 = var_20
  loc_1106642B:                                                       If (var_20.UnkVCall_0000009Ch = 0) = 0 Then
  loc_1106650F:                                                         If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 14) = global_1100AE28) Then
  loc_110665A6:                                                           var_8198 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 14)
  loc_110665B9:                                                           var_78 = var_8198
  loc_110665FC:                                                           var_90 = "GenLen".00000001h(1, global_1106884D, global_1106884D, global_1106884D, global_1106884D, var_74, var_8198, var_7C)
  loc_11066645:                                                           If (var_90 > 10) Then
  loc_1106664F:                                                             var_24 = "票号超长"
  loc_11066660:                                                           Else
  loc_11066660:                                                           End If
  loc_1106673E:                                                           If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 15) = global_1100AE28) Then
  loc_110667D5:                                                             var_81A8 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 15)
  loc_110667E8:                                                             var_78 = var_81A8
  loc_11066817:                                                             var_90 = "DateCheck".00000001h(1, global_1106884D, global_1106884D, global_1106884D, global_1106884D, var_74, var_81A8, var_7C)
  loc_11066867:                                                             If CBool(Not(var_90)) Then
  loc_11066871:                                                               var_24 = "票号发生日期非法"
  loc_11066882:                                                             Else
  loc_11066882:                                                             End If
  loc_11066882:                                                           End If
  loc_110668A5:                                                           var_1F0 = var_20
  loc_110668F2:                                                           var_1F8 = var_20
  loc_110668F8:                                                           var_1D8 = var_20.UnkVCall_0000008Ch
  loc_1106695B:                                                           If (var_20.UnkVCall_000000A4h = 0) = 0 Then
  loc_11066A1A:                                                             If (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 16) = global_1100AE28) Then
  loc_11066AC4:                                                               var_78 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, 16)
  loc_11066B44:                                                               var_38 = "BmCheck".00000002h(var_154, 0, var_15C)
  loc_11066B76:                                                               var_81C8 = (var_38 = global_1100AE28)
  loc_11066B7E:                                                               If var_81C8 = 0 Then
  loc_11066B88:                                                                 var_24 = "部门非法"
  loc_11066B99:                                                               Else
  loc_11066BB6:                                                                 var_118 = arg_C
  loc_11066C40:                                                                 frmGzToPzTGZP.VFG.DispID_0082(16, var_38)
  loc_11066C75:                                                                 var_1F0 = var_20
  loc_11066CA7:                                                                 If var_20.UnkVCall_000000A4h Then
  loc_11066CB5:                                                                   var_118 = var_38
  loc_11066D07:                                                                   var_128 = var_28
  loc_11066D55:                                                                   "BmToProperties".00000002h
  loc_11066D72:                                                                   Set var_28 = {3302AA19-EB96-11D2-AF06000021009B21}()
  loc_11066DA0:                                                                   var_1F0 = var_28
  loc_11066DA6:                                                                   var_1D4 = var_28.UnkVCall_00000034h
  loc_11066DCC:                                                                   If var_1D4 = 0 Then
  loc_11066DDA:                                                                     var_24 = "部门非末级"
  loc_11066DEB:                                                                   Else
  loc_11066DF3:                                                                     var_24 = "部门为空"
  loc_11066E04:                                                                   Else
  loc_11066E86:                                                                     frmGzToPzTGZP.VFG.DispID_0082(var_128, 285257256)
  loc_11066E98:                                                                   End If
  loc_11066E98:                                                                 End If
  loc_11066EBB:                                                                 var_1F0 = var_20
  loc_11066EED:                                                                 If var_20.UnkVCall_0000008Ch Then
  loc_11066F99:                                                                   var_81E0 = (frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H11) = global_1100AE28)
  loc_11066FD1:                                                                   If var_81E0 Then
  loc_1106707D:                                                                     var_81E8 = (frmGzToPzTGZP.VFG.DispID_0082(arg_C, 16) = global_1100AE28)
  loc_110670D9:                                                                     If var_81E8 + 1 Then
  loc_1106715C:                                                                       var_78 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H11)
  loc_110671EA:                                                                       var_90 = "ZyCheck".00000003h(var_174, "BmCheck".00000002h(var_154, 80020004h, var_15C), var_17C)
  loc_110671FF:                                                                       var_34 = var_90
  loc_11067231:                                                                       var_81F4 = (var_34 = global_1100AE28)
  loc_11067239:                                                                       If var_81F4 = 0 Then
  loc_11067243:                                                                         var_24 = "职员非法"
  loc_11067254:                                                                       Else
  loc_11067271:                                                                         var_118 = arg_C
  loc_110672FB:                                                                         frmGzToPzTGZP.VFG.DispID_0082(&H11, var_34)
  loc_1106731A:                                                                         var_118 = var_34
  loc_11067367:                                                                         var_128 = var_18
  loc_110673B5:                                                                         "ZyToProperties".00000002h
  loc_110673D2:                                                                         Set var_18 = {3302AA41-EB96-11D2-AF06000021009B21}()
  loc_110673E0:                                                                         var_118 = arg_C
  loc_11067421:                                                                         var_1F0 = var_18
  loc_110674DA:                                                                         frmGzToPzTGZP.VFG.DispID_0082(var_128, var_18.UnkVCall_0000002Ch)
  loc_110674FA:                                                                       Else
  loc_11067570:                                                                         var_158 = var_38
  loc_1106757D:                                                                         var_78 = frmGzToPzTGZP.VFG.DispID_0082(8, var_128)
  loc_11067632:                                                                         var_34 = "ZyCheck".00000003h(var_164, 0, var_16C)
  loc_11067664:                                                                         var_8208 = (var_34 = global_1100AE28)
  loc_1106766C:                                                                         If var_8208 = 0 Then
  loc_11067676:                                                                           var_24 = "职员不在指定部门内"
  loc_11067687:                                                                         Else
  loc_110676C5:                                                                           var_118 = arg_C
  loc_1106772C:                                                                           frmGzToPzTGZP.VFG.DispID_0082(&H11, var_34)
  loc_1106773E:                                                                         End If
  loc_1106773E:                                                                       End If
  loc_1106773E:                                                                     End If
  loc_11067761:                                                                     var_1F0 = var_20
  loc_11067793:                                                                     If var_20.UnkVCall_00000094h Then
  loc_1106783F:                                                                       var_8214 = (frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H12) = global_1100AE28)
  loc_11067850:                                                                       var_1F0 = var_8214
  loc_11067877:                                                                       If var_1F0 = 0 Then GoTo loc_11067D65
  loc_11067921:                                                                       var_78 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H12)
  loc_110679A1:                                                                       var_3C = "KhCheck".00000002h(var_154, 0, var_15C)
  loc_110679D3:                                                                       var_8220 = (var_3C = global_1100AE28)
  loc_110679DB:                                                                       If var_8220 = 0 Then
  loc_110679E5:                                                                         var_24 = "客户非法"
  loc_110679F6:                                                                       Else
  loc_11067A34:                                                                         var_118 = arg_C
  loc_11067A9B:                                                                         frmGzToPzTGZP.VFG.DispID_0082(&H12, var_3C)
  loc_11067AAD:                                                                       End If
  loc_11067AD0:                                                                       var_1F0 = var_20
  loc_11067B02:                                                                       If var_20.UnkVCall_0000009Ch Then
  loc_11067BAE:                                                                         var_822C = (frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H13) = global_1100AE28)
  loc_11067BBF:                                                                         var_1F0 = var_822C
  loc_11067BE6:                                                                         If var_1F0 = 0 Then GoTo loc_1106811E
  loc_11067C90:                                                                         var_78 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H13)
  loc_11067D10:                                                                         var_30 = "GysCheck".00000002h(var_154, 0, var_15C)
  loc_11067D42:                                                                         var_8238 = (var_30 = global_1100AE28)
  loc_11067D4A:                                                                         If var_8238 = 0 Then
  loc_11067D54:                                                                           var_24 = "供应商非法"
  loc_11067D60:                                                                           GoTo loc_1106880E
  loc_11067D6D:                                                                           var_24 = "客户为空"
  loc_11067D7E:                                                                         Else
  loc_11067DBC:                                                                           var_118 = arg_C
  loc_11067E23:                                                                           frmGzToPzTGZP.VFG.DispID_0082(&H13, var_30)
  loc_11067E35:                                                                         End If
  loc_11067E58:                                                                         var_1F0 = var_20
  loc_11067EA5:                                                                         var_1F8 = var_20
  loc_11067EAB:                                                                         var_1D8 = var_20.UnkVCall_0000009Ch
  loc_11067EE9:                                                                         If (var_20.UnkVCall_00000094h = 0) = 0 Then
  loc_11067F95:                                                                           var_8248 = (frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H14) = global_1100AE28)
  loc_11067FCD:                                                                           If var_8248 Then
  loc_11068064:                                                                             var_824C = frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H14)
  loc_11068077:                                                                             var_78 = var_824C
  loc_110680BA:                                                                             var_90 = "GenLen".00000001h(global_1106884D, global_1106884D, global_1106884D, global_1106884D, global_1106884D, var_74, var_824C, var_7C)
  loc_11068103:                                                                             If (var_90 > 20) Then
  loc_1106810D:                                                                               var_24 = "业务员超长"
  loc_11068119:                                                                               GoTo loc_1106880E
  loc_11068126:                                                                               var_24 = "供应商为空"
  loc_11068137:                                                                             Else
  loc_11068137:                                                                             End If
  loc_11068137:                                                                           End If
  loc_11068157:                                                                           var_1F0 = var_20
  loc_110681AF:                                                                           If (var_20.UnkVCall_000000ACh = global_1100AE28) Then
  loc_1106825B:                                                                             var_8260 = (frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H15) = global_1100AE28)
  loc_11068293:                                                                             If var_8260 Then
  loc_110682B9:                                                                               var_1F0 = var_20
  loc_110682EC:                                                                               var_8268 = (var_20.UnkVCall_000000ACh = global_1100AE28)
  loc_11068311:                                                                               If var_8268 Then
  loc_11068337:                                                                                 var_1F0 = var_20
  loc_11068367:                                                                                 var_78 = var_20.UnkVCall_000000ACh
  loc_11068415:                                                                                 var_88 = frmGzToPzTGZP.VFG.DispID_0082(arg_C, &H15)
  loc_1106849D:                                                                                 var_A0 = "XmCheck".00000003h(var_164, Not(8), var_16C)
  loc_110684B2:                                                                                 var_2C = var_A0
  loc_110684EB:                                                                                 var_8278 = (var_2C = global_1100AE28)
  loc_110684F3:                                                                                 If var_8278 = 0 Then
  loc_110684FD:                                                                                   var_24 = "项目非法"
  loc_1106850E:                                                                                 Else
  loc_1106853A:                                                                                   var_4C = var_20.UnkVCall_000000ACh
  loc_11068568:                                                                                   var_128 = var_2C
  loc_11068599:                                                                                   Set var_58 = var_1C
  loc_1106861B:                                                                                   "XmToProperties".00000003h
  loc_11068638:                                                                                   Set var_1C = {3302AA4B-EB96-11D2-AF06000021009B21}()
  loc_1106868D:                                                                                   If var_1C.UnkVCall_00000034h Then
  loc_1106869B:                                                                                     var_24 = "项目已结算"
  loc_110686AC:                                                                                   Else
  loc_110686D6:                                                                                     var_118 = %cobj
  loc_1106874E:                                                                                     frmGzToPzTGZP.VFG.DispID_0082(&H15, 285257256)
  loc_1106876B:                                                                                   Else
  loc_11068773:                                                                                     var_24 = "制单日期非法"
  loc_11068779:                                                                                   End If
  loc_11068779:                                                                                 End If
  loc_11068779:                                                                               End If
  loc_1106877F:                                                                               GoTo loc_1106880E
  loc_11068788:                                                                               If var_4 Then
  loc_11068793:                                                                               End If
  loc_1106880D:                                                                               Exit Sub
  loc_1106880E:                                                                             End If
  loc_1106880E:                                                                           End If
  loc_1106880E:                                                                         End If
  loc_1106880E:                                                                       End If
  loc_1106880E:                                                                     End If
  loc_1106880E:                                                                   End If
  loc_1106880E:                                                                 End If
  loc_1106880E:                                                               End If
  loc_1106880E:                                                             End If
  loc_1106880E:                                                           End If
  loc_1106880E:                                                         End If
  loc_1106880E:                                                       End If
  loc_1106880E:                                                     End If
  loc_1106880E:                                                   End If
  loc_1106880E:                                                 End If
  loc_1106880E:                                               End If
  loc_1106880E:                                             End If
  loc_1106880E:                                           End If
  loc_1106880E:                                         End If
  loc_1106880E:                                       End If
  loc_1106880E:                                     End If
  loc_1106880E:                                   End If
  loc_1106880E:                                 End If
  loc_1106880E:                               End If
  loc_1106880E:                             End If
  loc_1106880E:                           End If
  loc_1106880E:                         End If
  loc_1106880E:                       End If
  loc_1106880E:                     End If
  loc_1106880E:                   End If
  loc_1106880E:                 End If
  loc_1106880E:               End If
  loc_1106880E:             End If
  loc_1106880E:           End If
  loc_1106880E:         End If
  loc_1106880E:       End If
  loc_1106880E:     End If
  loc_1106880E:   End If
  loc_1106880E: End If
  loc_1106880E: ' Referenced from: 1106877F
End Sub

Private Sub Proc_12_15_11068870
  Dim var_9C As Variant
  Dim var_8034 As Label
  Dim var_8074 As Label
  Dim var_A0 As Variant
  Dim var_38 As {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  Dim var_28 As {3302AA47-EB96-11D2-AF06000021009B21}()
  loc_110689CA: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110689D0: var_294 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110689F6: Set var_9C = frmGzToPzTGZP.VFG
  loc_11068A40: If (CLng(var_9C.DispID_0007) < 2) Then
  loc_11068A6E:   var_800C = = Global.Screen
  loc_11068A90:   var_8010 = ecx
  loc_11068A98:   var_8010 = var_9C.UnkVCall_0000007Ch
  loc_11068B05:   var_C8 = "提示信息"
  loc_11068B07:   var_150 = "没有可生成用友凭证的数据。"
  loc_11068B16: Else
  loc_11068BC6:   var_264 = ("GetAccInfo".00000002h(, , fs:[00000000h], , "GL", var_16C, "dGLStartDate", var_174) = 1100AE28h)
  loc_11068BE0:   If var_264 = 0 Then GoTo loc_11068D21
  loc_11068C0E:   var_801C = = Global.Screen
  loc_11068C30:   var_8020 = ecx
  loc_11068C38:   var_8020 = var_9C.UnkVCall_0000007Ch
  loc_11068CA5:   var_C8 = "提示信息"
  loc_11068CA7:   var_150 = "总账系统尚未启用，不能进行凭证引入！"
  loc_11068CB1: End If
  loc_11068CE3: MsgBox(var_150, 64, var_C8, var_D8, var_E8)
  loc_11068D10: Exit Sub
  loc_11068D1C: GoTo loc_110718F4
  loc_11068D2B: var_8024 = "IF NOT EXISTS (SELECT * FROM Sysobjects WHERE ID = object_id(N'[dbo].[VouchNum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " & " CREATE TABLE VouchNum(iperiod tinyint NULL ,csign varchar(8) NULL ,ino_id int NULL,constraint index1 unique(iperiod,csign,ino_id))"
  loc_11068D31: var_B0 = var_8024
  loc_11068D90: var_D8.00000001h(0, , , , "3缷M鋎?", var_AC, var_8024, var_B4)
  loc_11068DB0: On Error GoTo 0
  loc_11068DB6: var_B0 = %ecx = %S_edx_S
  loc_11068DD8: var_78 = "AS13"
  loc_11068DF0: var_78)
  loc_11068E1A: If Not (var_78)) Then
  loc_11068E4B:   If Global.Screen < 0 Then
  loc_11068E5C:   End If
  loc_11068E66:   var_8030 = ecx
  loc_11068E75:   If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_11068E84:     var_8030 = CheckObj(var_9C, global_1100C47C, 124)
  loc_11068E8F:   End If
  loc_11068EA0:   call var_8034 = var_9C(var_9C, frmGzToPzTGZP.Label3)
  loc_11068EA2:   var_264 = var_8034
  loc_11068EB0:   Label3.Caption = "正在进行数据分析，请稍等..."
  loc_11068EDD:   var_150 = True
  loc_11068F20:   call var_8038 = var_9C(var_9C, frmGzToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, True, var_14C)
  loc_11068F23:   var_8038.DispID_0000 =
  loc_11068F4C:   call var_803C = var_9C(var_9C, frmGzToPzTGZP.Pic1, global_FFFFFDDA, var_9C = var_9C)
  loc_11068F4F:   var_803C.DispID_0000
  loc_11068F6E:   var_8040 = .Proc_12_13_110613A0(var_24C)
  loc_11068F7C:   If var_24C = 2 Then
  loc_11068F82:     var_150 = %ecx = %S_edx_S
  loc_11068FC5:     call var_8044 = var_9C(var_9C, frmGzToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C)
  loc_11068FC8:     var_8044.DispID_0000 =
  loc_11069064:     MsgBox("数据源中没有合法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_110690A1:     var_24C = %ecx = %S_edx_S
  loc_110690C7:     "AS13")
  loc_11069109:     var_B8 = Global.Screen
  loc_1106912B:     var_804C = ecx
  loc_1106913A:     If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_11069149:       var_804C = CheckObj(var_9C, global_1100C47C, 124)
  loc_11069154:     End If
  loc_11069156:     If var_804C = 1 Then
  loc_1106915C:       var_150 = %ecx = %S_edx_S
  loc_1106919F:       call var_8050 = var_9C(var_9C, frmGzToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, var_14C = var_9C, var_14C)
  loc_110691A2:       var_8050.DispID_0000 =
  loc_1106923E:       MsgBox("数据源中含有非法的数据，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1106927B:       var_24C = %ecx = %S_edx_S
  loc_110692A1:       "AS13")
  loc_110692E3:       var_B8 = Global.Screen
  loc_11069305:       var_8058 = ecx
  loc_11069314:       If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_11069323:         var_8058 = CheckObj(var_9C, global_1100C47C, 124)
  loc_1106932E:       End If
  loc_11069330:       If var_8058 = 3 Then
  loc_11069336:         var_150 = %ecx = %S_edx_S
  loc_11069379:         call var_805C = var_9C(var_9C, frmGzToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, var_9C = var_9C, var_14C)
  loc_1106937C:         var_805C.DispID_0000 =
  loc_11069418:         MsgBox("数据源中指定的凭证号无效或重号，无法生成用友凭证，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_11069455:         var_24C = %ecx = %S_edx_S
  loc_1106947B:         "AS13")
  loc_110694BD:         var_B8 = Global.Screen
  loc_110694DF:         var_8064 = ecx
  loc_110694EE:         If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110694FD:           var_8064 = CheckObj(var_9C, global_1100C47C, 124)
  loc_11069508:         End If
  loc_1106954A:         var_C8 = "提示信息"
  loc_11069570:         var_B8 = "数据源中的数据已全部通过检查，是否开始引入？"
  loc_11069594:         MsgBox(var_B8, 36, var_C8, var_D8, var_E8)
  loc_110695D9:         If (MsgBox(var_B8, 36, var_C8, var_D8, var_E8) = 7) Then
  loc_11069624:           call var_8068 = var_9C(var_9C, frmGzToPzTGZP.Pic1, global_80010007, 0000000Bh, var_154, frmGzToPzTGZP.Pic1, var_14C)
  loc_11069627:           var_8068.DispID_0000 =
  loc_1106964D:           var_24C = %ecx = %S_edx_S
  loc_11069673:           "AS13")
  loc_110696B5:           var_B8 = Global.Screen
  loc_110696D7:           var_8070 = ecx
  loc_110696E6:           If var_9C.UnkVCall_0000007Ch < 0 Then
  loc_110696F5:             var_8070 = CheckObj(var_9C, global_1100C47C, 124)
  loc_11069700:           End If
  loc_11069701:           On Error GoTo 0
  loc_11069718:           call var_8074 = var_9C(var_9C, frmGzToPzTGZP.Label3, var_9C = var_9C)
  loc_1106971A:           var_264 = var_8074
  loc_11069728:           Label3.Caption = "正在写数据，请稍等..."
  loc_1106976C:           call var_8078 = var_9C(var_9C, frmGzToPzTGZP.Pic1, global_FFFFFDDA, 00000000h)
  loc_1106976F:           var_8078.DispID_0000
  loc_110697A6:           Set var_74 = CreateObject("UfDbKit.UfRecordset", 0)
  loc_110697BD:           var_150 = "SELECT TOP 1 * FROM GL_accvouch"
  loc_11069832:           Set var_74 = "DataMdb".00000000h.00000001h(var_14C, "SELECT TOP 1 * FROM GL_accvouch", var_154)
  loc_11069866:           call var_8084 = var_9C(var_9C, frmGzToPzTGZP.VFG, 00000007h, 00000000h)
  loc_110698CA:           If var_24 <= CLng(var_8084.DispID_0000)(-1) Then
  loc_110698D4:             var_2A8 = var_24
  loc_110698DA:             var_150 = var_24
  loc_11069957:             call var_8090 = var_9C(var_9C, frmGzToPzTGZP.VFG, 00000082h, 00000002h, 3, var_174, 2, var_16C, 00000003h, var_154, var_24, var_14C)
  loc_11069971:             var_C0 = var_8090.DispID_0000
  loc_1106998F:             var_D8)
  loc_110699E7:             var_70 = CByte("DateToPeriod".00000001h(8, var_D4))
  loc_11069A20:             var_150 = var_2A8
  loc_11069A99:             call var_809C = var_9C(var_9C, frmGzToPzTGZP.VFG, 00000082h, 00000002h, 3, var_174, 3, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_11069AB8:             var_58 = var_809C.DispID_0000
  loc_11069ADC:             var_150 = var_2A8
  loc_11069B59:             call var_80A4 = var_9C(var_9C, frmGzToPzTGZP.VFG, 00000082h, 00000002h, 3, var_174, 0, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_11069B78:             var_64 = var_80A4.DispID_0000
  loc_11069B9C:             var_150 = var_2A8
  loc_11069C19:             call var_80AC = var_9C(var_9C, frmGzToPzTGZP.VFG, 00000082h, 00000002h, 3, var_174, 1, var_16C, 00000003h, var_154, var_2A8, var_14C)
  loc_11069C82:             If (var_80AC.DispID_0000 = global_1100D76C) Then
  loc_11069C99:               call var_80B8 = var_9C(var_A8, frmGzToPzTGZP.Label3)
  loc_11069C9B:               var_264 = var_80B8
  loc_11069DAB:               var_80 = "正在处理：第[" & frmGzToPzTGZP.VFG.DispID_0082(var_2A8, 2) & " - "
  loc_11069EEC:               var_D8 = frmGzToPzTGZP.VFG.DispID_0082(var_2A8, 0)
  loc_11069F33:               var_98 = var_80 & frmGzToPzTGZP.VFG.DispID_0082(var_2A8, 3) & " - " & var_D8 & "]号凭证"
  loc_11069F43:               var_98 = var_80B8.UnkVCall_00000054h
  loc_11069FFE:               frmGzToPzTGZP.Pic1.DispID_FFFFFDDA
  loc_1106A032:               var_3C = var_24
  loc_1106A046:               Set var_9C = frmGzToPzTGZP.Chk
  loc_1106A048:               var_264 = var_9C
  loc_1106A05A:               Set var_A0 = var_9C(0)
  loc_1106A07E:               var_26C = var_A0
  loc_1106A0E8:               If (var_A0.Value = 1) Then
  loc_1106A11B:                 var_24C = CInt("cIYear".00000000h)
  loc_1106A130:                 var_24C, var_70)
  loc_1106A13D:                 var_54 = var_24C, var_70)
  loc_1106A14E:               Else
  loc_1106A164:                 var_80E8 = .Proc_12_17_110731E0(var_70)
  loc_1106A176:                 var_54 = var_258
  loc_1106A179:               End If
  loc_1106A17E:               If var_54 > 0 Then
  loc_1106A186:                 On Error GoTo loc_1106FAF0
  loc_1106A1BF:                 "wksAlias".00000000h.00000000h(var_58)
  loc_1106A1DE:                 var_1A0 = var_70
  loc_1106A2A7:                 var_D8)
  loc_1106A353:                 var_80FC = (var_58 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 3))
  loc_1106A360:                 var_1F0 = var_80FC + 1
  loc_1106A41C:                 var_8104 = (var_64 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 0))
  loc_1106A429:                 var_240 = var_8104 + 1
  loc_1106A4C6:                 var_8114 = CBool((frmGzToPzTGZP.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_80FC + 1 And var_8104 + 1)
  loc_1106A54B:                 If var_8114 Then
  loc_1106A5EC:                   var_C0 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 6)
  loc_1106A629:                   var_1A0 = var_38
  loc_1106A697:                   "kmCodeToProperties".00000002h
  loc_1106A6B7:                   Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_1106A6F0:                   var_74.AddNew
  loc_1106A6FB:                   var_150 = "ibook"
  loc_1106A76C:                   var_74.DispID_0000(0)
  loc_1106A76E:                   var_1A0 = "iPeriod"
  loc_1106A81D:                   var_C0 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 2)
  loc_1106A83B:                   var_D8)
  loc_1106A8D4:                   var_74.DispID_0000("DateToPeriod".00000001h)
  loc_1106A909:                   var_190 = "csign"
  loc_1106AA16:                   var_74.DispID_0000(frmGzToPzTGZP.VFG.DispID_0082(var_24, 3))
  loc_1106AA3D:                   var_190 = "isignseq"
  loc_1106AB5D:                   var_74.DispID_0000(Proc_0_4_11026BD0(frmGzToPzTGZP.VFG.DispID_0082(var_24, 3), var_64, var_258))
  loc_1106AB88:                   var_150 = "ino_id"
  loc_1106ABFA:                   var_74.DispID_0000(var_54)
  loc_1106ABFC:                   var_190 = "dbill_date"
  loc_1106ACAB:                   var_C0 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 2)
  loc_1106ACC9:                   var_D8)
  loc_1106AD26:                   var_74.DispID_0000(var_D8)
  loc_1106AD54:                   var_190 = "idoc"
  loc_1106AD6C:                   var_150 = var_24
  loc_1106AE75:                   var_74.DispID_0000(Val(frmGzToPzTGZP.VFG.DispID_0082(var_150, 4)))
  loc_1106AEA0:                   var_160 = "ctext1"
  loc_1106AF07:                   var_74.DispID_0000(var_150)
  loc_1106AF0E:                   var_160 = "ctext2"
  loc_1106AF75:                   var_74.DispID_0000(var_150)
  loc_1106AF7C:                   var_150 = "cbill"
  loc_1106AFEA:                   var_74.DispID_0000("cUserName".00000000h(, var_14C, "cbill", var_154))
  loc_1106B000:                   var_160 = "cbook"
  loc_1106B067:                   var_74.DispID_0000(var_150)
  loc_1106B06E:                   var_160 = "ccheck"
  loc_1106B0D5:                   var_74.DispID_0000(var_150)
  loc_1106B0DC:                   var_160 = "ccashier"
  loc_1106B143:                   var_74.DispID_0000(var_150)
  loc_1106B14A:                   var_160 = "iflag"
  loc_1106B1B1:                   var_74.DispID_0000(var_150)
  loc_1106B1B8:                   var_160 = "coutaccset"
  loc_1106B21F:                   var_74.DispID_0000(var_150)
  loc_1106B226:                   var_160 = "ioutyear"
  loc_1106B28D:                   var_74.DispID_0000(var_150)
  loc_1106B294:                   var_160 = "coutsysver"
  loc_1106B2FB:                   var_74.DispID_0000(var_150)
  loc_1106B302:                   var_160 = "coutsysname"
  loc_1106B369:                   var_74.DispID_0000(var_150)
  loc_1106B370:                   var_170 = "ioutperiod"
  loc_1106B40D:                   var_74.DispID_0000(var_74.DispID_0000("iPeriod"))
  loc_1106B41E:                   var_170 = "doutbilldate"
  loc_1106B4E1:                   var_74.DispID_0000(CStr(var_74.DispID_0000("dbill_date")))
  loc_1106B4FE:                   var_150 = "iYear"
  loc_1106B56C:                   var_74.DispID_0000("cIYear".00000000h(var_58, var_14C, "iYear", var_154))
  loc_1106B66A:                   var_74.DispID_0000("cIYear".00000000h(, var_16C, "iYPeriod", var_174) & Format(var_70, "00"))
  loc_1106B698:                   var_160 = "coutsign"
  loc_1106B6FF:                   var_74.DispID_0000(var_70)
  loc_1106B701:                   var_190 = "coutno_id"
  loc_1106B80E:                   var_74.DispID_0000(frmGzToPzTGZP.VFG.DispID_0082(var_24, 3))
  loc_1106B83A:                   var_150 = "bvouchedit"
  loc_1106B8A9:                   var_74.DispID_0000(FFFFFFFFh)
  loc_1106B8B0:                   var_150 = "bvouchaddordele"
  loc_1106B921:                   var_74.DispID_0000(FFFFFFFFh)
  loc_1106B928:                   var_150 = "bvouchmoneyhold"
  loc_1106B999:                   var_74.DispID_0000(FFFFFFFFh)
  loc_1106B9A0:                   var_150 = "bvalueedit"
  loc_1106BA11:                   var_74.DispID_0000(FFFFFFFFh)
  loc_1106BA18:                   var_150 = "bcodeedit"
  loc_1106BA89:                   var_74.DispID_0000(FFFFFFFFh)
  loc_1106BA90:                   var_150 = "bPCSedit"
  loc_1106BB01:                   var_74.DispID_0000(FFFFFFFFh)
  loc_1106BB08:                   var_150 = "bDeptedit"
  loc_1106BB79:                   var_74.DispID_0000(FFFFFFFFh)
  loc_1106BB80:                   var_150 = "bItemedit"
  loc_1106BBF1:                   var_74.DispID_0000(FFFFFFFFh)
  loc_1106BBF8:                   var_150 = "inid"
  loc_1106BC6A:                   var_74.DispID_0000(1)
  loc_1106BC6C:                   var_190 = "cdigest"
  loc_1106BD7D:                   var_74.DispID_0000(frmGzToPzTGZP.VFG.DispID_0082(var_24, 5))
  loc_1106BDA4:                   var_190 = "cCode"
  loc_1106BEB3:                   var_74.DispID_0000(frmGzToPzTGZP.VFG.DispID_0082(var_24, 6))
  loc_1106BF5B:                   var_7C = var_38.UnkVCall_0000006Ch
  loc_1106BFA6:                   var_8150 = (var_38.UnkVCall_0000006Ch = global_1100AE28)
  loc_1106BFB3:                   var_160 = var_8150 + 1
  loc_1106C03E:                   var_74.DispID_0000(IIf(var_8150 + 1, vbNull, 0))
  loc_1106C123:                   var_1B0 = "md"
  loc_1106C16C:                   var_BC = var_25C
  loc_1106C1F3:                   var_74.DispID_0000(Format(Val(frmGzToPzTGZP.VFG.DispID_0082(var_24, 7)), "#.00"))
  loc_1106C2E4:                   var_1B0 = "mc"
  loc_1106C32D:                   var_BC = var_25C
  loc_1106C3B4:                   var_74.DispID_0000(Format(Val(frmGzToPzTGZP.VFG.DispID_0082(var_24, 8)), "#.00"))
  loc_1106C47C:                   If (var_74.DispID_0000("md") <> 0) Then
  loc_1106C4F1:                     If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_1106C4FC:                       var_150 = "md_f"
  loc_1106C56D:                       var_74.DispID_0000(0)
  loc_1106C577:                     Else
  loc_1106C62A:                       var_1B0 = "md_f"
  loc_1106C673:                       var_BC = var_25C
  loc_1106C6FA:                       var_74.DispID_0000(Format(Val(frmGzToPzTGZP.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_1106C73B:                     End If
  loc_1106C7AD:                     If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_1106C7B8:                       var_150 = "nd_s"
  loc_1106C829:                       var_74.DispID_0000(0)
  loc_1106C833:                     Else
  loc_1106C842:                     Else
  loc_1106C8B1:                       If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_1106C8BC:                         var_150 = "mc_f"
  loc_1106C92D:                         var_74.DispID_0000(0)
  loc_1106C937:                       Else
  loc_1106C9EA:                         var_1B0 = "mc_f"
  loc_1106CA33:                         var_BC = var_25C
  loc_1106CABA:                         var_74.DispID_0000(Format(Val(frmGzToPzTGZP.VFG.DispID_0082(var_24, 10)), "#.00"))
  loc_1106CAFB:                       End If
  loc_1106CB6D:                       If (var_38.UnkVCall_0000007Ch = global_1100AE28) + 1 Then
  loc_1106CB74:                         GoTo loc_1106C7B8
  loc_1106CB79:                       End If
  loc_1106CB83:                     End If
  loc_1106CC9D:                     var_74.DispID_0000(Val(frmGzToPzTGZP.VFG.DispID_0082(var_24, 9)))
  loc_1106CCC3:                   End If
  loc_1106CD35:                   If (var_38.UnkVCall_0000006Ch = global_1100AE28) + 1 Then
  loc_1106CD40:                     var_150 = "nfrat"
  loc_1106CDB1:                     var_74.DispID_0000(0)
  loc_1106CDBB:                   Else
  loc_1106CEDF:                     var_74.DispID_0000(Val(frmGzToPzTGZP.VFG.DispID_0082(var_24, 11)))
  loc_1106CF05:                   End If
  loc_1106CF5A:                   If var_38.UnkVCall_0000010Ch Then
  loc_1106CFF1:                     var_1F0 = "csettle"
  loc_1106D0D8:                     var_81A4 = (frmGzToPzTGZP.VFG.DispID_0082(var_24, 13) = global_1100AE28)
  loc_1106D0E5:                     var_1E0 = var_81A4 + 1
  loc_1106D170:                     var_74.DispID_0000(IIf(var_81A4 + 1, vbNull, frmGzToPzTGZP.VFG.DispID_0082(var_24, 13)))
  loc_1106D1C9:                   End If
  loc_1106D1F2:                   var_24C = var_38.UnkVCall_0000010Ch
  loc_1106D23F:                   var_250 = var_38.UnkVCall_00000094h
  loc_1106D2DE:                   If (var_38.UnkVCall_0000009Ch = 0) = 0 Then
  loc_1106D375:                     var_1F0 = "cn_id"
  loc_1106D424:                     var_E0 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 14)
  loc_1106D45C:                     var_81BC = (frmGzToPzTGZP.VFG.DispID_0082(var_24, 14) = global_1100AE28)
  loc_1106D469:                     var_1E0 = var_81BC + 1
  loc_1106D4F4:                     var_74.DispID_0000(IIf(var_81BC + 1, vbNull, var_E0))
  loc_1106D5DB:                     var_1F0 = "dt_date"
  loc_1106D68A:                     var_D0 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 15)
  loc_1106D6A8:                     var_E0)
  loc_1106D6D5:                     var_81C8 = (frmGzToPzTGZP.VFG.DispID_0082(var_24, 15) = global_1100AE28)
  loc_1106D6E2:                     var_1E0 = var_81C8 + 1
  loc_1106D76D:                     var_74.DispID_0000(IIf(var_81C8 + 1, vbNull, var_E0))
  loc_1106D85B:                     var_1F0 = "cname"
  loc_1106D942:                     var_81D4 = (frmGzToPzTGZP.VFG.DispID_0082(var_24, &H14) = global_1100AE28)
  loc_1106D94F:                     var_1E0 = var_81D4 + 1
  loc_1106D9DA:                     var_74.DispID_0000(IIf(var_81D4 + 1, vbNull, frmGzToPzTGZP.VFG.DispID_0082(var_24, &H14)))
  loc_1106DA33:                   End If
  loc_1106DAA9:                   var_250 = var_38.UnkVCall_0000008Ch
  loc_1106DAE7:                   If (var_38.UnkVCall_000000A4h = 0) = 0 Then
  loc_1106DAF1:                     var_150 = var_24
  loc_1106DB7E:                     var_1F0 = "cdept_id"
  loc_1106DC65:                     var_81E8 = (frmGzToPzTGZP.VFG.DispID_0082(var_150, 16) = global_1100AE28)
  loc_1106DC72:                     var_1E0 = var_81E8 + 1
  loc_1106DCFD:                     var_74.DispID_0000(IIf(var_81E8 + 1, vbNull, frmGzToPzTGZP.VFG.DispID_0082(var_24, 16)))
  loc_1106DD58:                   Else
  loc_1106DD5D:                     var_160 = "cdept_id"
  loc_1106DDC4:                     var_74.DispID_0000(var_150)
  loc_1106DDC9:                   End If
  loc_1106DE1E:                   If var_38.UnkVCall_0000008Ch Then
  loc_1106DE28:                     var_150 = var_24
  loc_1106DEB5:                     var_1F0 = "cperson_id"
  loc_1106DF9C:                     var_81F8 = (frmGzToPzTGZP.VFG.DispID_0082(var_150, &H11) = global_1100AE28)
  loc_1106DFA9:                     var_1E0 = var_81F8 + 1
  loc_1106E034:                     var_74.DispID_0000(IIf(var_81F8 + 1, vbNull, frmGzToPzTGZP.VFG.DispID_0082(var_24, &H11)))
  loc_1106E08F:                   Else
  loc_1106E094:                     var_160 = "cperson_id"
  loc_1106E0FB:                     var_74.DispID_0000(var_150)
  loc_1106E100:                   End If
  loc_1106E155:                   If var_38.UnkVCall_00000094h Then
  loc_1106E15F:                     var_150 = var_24
  loc_1106E1EC:                     var_1F0 = "ccus_id"
  loc_1106E2D3:                     var_8208 = (frmGzToPzTGZP.VFG.DispID_0082(var_150, &H12) = global_1100AE28)
  loc_1106E2E0:                     var_1E0 = var_8208 + 1
  loc_1106E36B:                     var_74.DispID_0000(IIf(var_8208 + 1, vbNull, frmGzToPzTGZP.VFG.DispID_0082(var_24, &H12)))
  loc_1106E3C6:                   Else
  loc_1106E3CB:                     var_160 = "ccus_id"
  loc_1106E432:                     var_74.DispID_0000(var_150)
  loc_1106E437:                   End If
  loc_1106E48C:                   If var_38.UnkVCall_0000009Ch Then
  loc_1106E496:                     var_150 = var_24
  loc_1106E523:                     var_1F0 = "csup_id"
  loc_1106E60A:                     var_8218 = (frmGzToPzTGZP.VFG.DispID_0082(var_150, &H13) = global_1100AE28)
  loc_1106E617:                     var_1E0 = var_8218 + 1
  loc_1106E6A2:                     var_74.DispID_0000(IIf(var_8218 + 1, vbNull, frmGzToPzTGZP.VFG.DispID_0082(var_24, &H13)))
  loc_1106E6FD:                   Else
  loc_1106E702:                     var_160 = "csup_id"
  loc_1106E769:                     var_74.DispID_0000(var_150)
  loc_1106E76E:                   End If
  loc_1106E7E7:                   If (var_38.UnkVCall_000000ACh = global_1100AE28) Then
  loc_1106E7F1:                     var_150 = var_24
  loc_1106E87E:                     var_1F0 = "citem_id"
  loc_1106E965:                     var_822C = (frmGzToPzTGZP.VFG.DispID_0082(var_150, &H15) = global_1100AE28)
  loc_1106E972:                     var_1E0 = var_822C + 1
  loc_1106E9FD:                     var_74.DispID_0000(IIf(var_822C + 1, vbNull, frmGzToPzTGZP.VFG.DispID_0082(var_24, &H15)))
  loc_1106EADA:                     var_7C = var_38.UnkVCall_000000ACh
  loc_1106EB2B:                     var_8238 = (var_38.UnkVCall_000000ACh = global_1100AE28)
  loc_1106EB38:                     var_160 = var_8238 + 1
  loc_1106EBC3:                     var_74.DispID_0000(IIf(var_8238 + 1, vbNull, 0))
  loc_1106EBFD:                   Else
  loc_1106EC02:                     var_160 = "citem_id"
  loc_1106EC69:                     var_74.DispID_0000(var_150)
  loc_1106EC70:                     var_160 = "citem_class"
  loc_1106ECD7:                     var_74.DispID_0000(var_150)
  loc_1106ECDC:                   End If
  loc_1106ECE1:                   var_160 = "ccode_equal"
  loc_1106ED48:                   var_74.DispID_0000(var_150)
  loc_1106ED4F:                   var_160 = "iflagbank"
  loc_1106EDB6:                   var_74.DispID_0000(var_150)
  loc_1106EDBD:                   var_160 = "iflagperson"
  loc_1106EE24:                   var_74.DispID_0000(var_150)
  loc_1106EE31:                   var_74.Update
  loc_1106EE48:                   var_24 = var_24(1)
  loc_1106EE59:                   var_68 = var_68(1)
  loc_1106EE8E:                   var_823C = CLng(frmGzToPzTGZP.VFG.DispID_0007)
  loc_1106EEAA:                   var_264 = (var_24(1) > 0)
  loc_1106EED1:                   If var_264 = 0 Then GoTo loc_1106A1DB
  loc_1106EED7:                 End If
  loc_1106EF0A:                 "wksAlias".00000000h.00000000h
  loc_1106EF37:                 Set var_9C = frmGzToPzTGZP.Chk
  loc_1106EF39:                 var_264 = var_9C
  loc_1106EF4B:                 Set var_A0 = var_9C(0)
  loc_1106EF6F:                 var_26C = var_A0
  loc_1106EFD9:                 If (var_A0.Value = 1) Then
  loc_1106EFE7:                   var_70, var_58)
  loc_1106EFEC:                 End If
  loc_1106EFEE:                 On Error GoTo 0
  loc_1106F025:                 var_250 = CInt("cIYear".00000000h)
  loc_1106F04F:                 var_24C, var_250, var_70, var_58)
  loc_1106F059:                 var_5C = var_24C, var_250, var_70, var_58)
  loc_1106F09C:                 var_250 = CInt("cIYear".00000000h)
  loc_1106F0D0:                 var_48 = r_250, var_70, var_58) var_250, var_70, var_58)
  loc_1106F0E2:                 var_150 = "select * from GL_accvouch where ibook=0 and iYear="
  loc_1106F10A:                 var_170 = var_70
  loc_1106F12E:                 var_824C = Proc_0_4_11026BD0(var_58, var_54, var_54)
  loc_1106F133:                 var_190 = var_824C
  loc_1106F15B:                 var_1B0 = var_54
  loc_1106F1B4:                 var_D8 = 1 & "cIYear".00000000h(, 1, 1) & " and iperiod="
  loc_1106F21D:                 var_128 = var_D8 & var_70 & " and isignseq=" & var_824C & " and ino_id=" & var_54
  loc_1106F286:                 Set var_74 = "DataMdb".00000000h.00000001h
  loc_1106F325:                 If CBool(Not(var_74.EOF)) Then
  loc_1106F37D:                   If CBool(Not(var_74.EOF)) Then
  loc_1106F386:                     var_170 = var_70
  loc_1106F39B:                     var_150 = "iPeriod"
  loc_1106F3BF:                     var_180 = "csign"
  loc_1106F3D3:                     var_1D0 = var_54
  loc_1106F3E4:                     var_1B0 = "ino_id"
  loc_1106F53B:                     If CBool((var_70 = var_14C) And (var_58 = var_D8) And (var_54 = var_1AC)) Then
  loc_1106F546:                       var_150 = "mc"
  loc_1106F5C8:                       var_180 = "ccode_equal"
  loc_1106F5DC:                       If (var_14C <> 0) Then
  loc_1106F608:                         var_8278 = (var_5C = global_1100AE28)
  loc_1106F615:                         var_160 = var_8278 + 1
  loc_1106F642:                         var_C8 = IIf(var_8278 + 1, vbNull, var_5C)
  loc_1106F6BC:                       Else
  loc_1106F6E2:                         var_827C = (var_48 = global_1100AE28)
  loc_1106F6EF:                         var_160 = var_827C + 1
  loc_1106F71C:                         var_C8 = IIf(var_827C + 1, vbNull, var_48)
  loc_1106F791:                       End If
  loc_1106F7A7:                       var_74.Update
  loc_1106F7F1:                       var_180 = var_38
  loc_1106F838:                       var_B8 = var_74.DispID_0000("cCode")
  loc_1106F895:                       "kmCodeToProperties".00000002h
  loc_1106F8B5:                       Set var_38 = {A02D7144-EE7B-4456-BE5F9B89A6496607}()
  loc_1106F8D4:                       var_150 = "citem_class"
  loc_1106F93B:                       If IsNull(var_74.DispID_0000(var_150)) Then
  loc_1106F950:                       Else
  loc_1106F991:                         var_180 = var_28
  loc_1106F9D8:                         var_B8 = var_74.DispID_0000(var_150)
  loc_1106FA35:                         "XmClassIDToProperties".00000002h
  loc_1106FA95:                         var_78 = {3302AA47-EB96-11D2-AF06000021009B21}().UnkVCall_0000002Ch
  loc_1106FAC6:                       End If
  loc_1106FAD4:                       var_68 = var_68(1)
  loc_1106FAE2:                       var_74.MoveNext
  loc_1106FAEB:                       GoTo loc_1106F332
  loc_1106FAF0:                       ' Referenced from: 1106A186
  loc_1106FB23:                       "wksAlias".00000000h.00000000h
  loc_1106FB3B:                       var_30 = var_3C
  loc_1106FB50:                       var_1A0 = var_70
  loc_1106FC19:                       var_D8)
  loc_1106FCC5:                       var_829C = (var_58 = frmGzToPzTGZP.VFG.DispID_0082(var_30, 3))
  loc_1106FCD2:                       var_1F0 = var_829C + 1
  loc_1106FD8E:                       var_82A4 = (var_64 = frmGzToPzTGZP.VFG.DispID_0082(var_30, 0))
  loc_1106FD9B:                       var_240 = var_82A4 + 1
  loc_1106FE31:                       var_82B0 = (frmGzToPzTGZP.VFG.DispID_0082(var_30, 2) = "DateToPeriod".00000001h) And var_829C + 1 And var_82A4 + 1
  loc_1106FEBD:                       If CBool(var_82B0) Then
  loc_1106FEC7:                         var_150 = var_30
  loc_1106FF83:                         frmGzToPzTGZP.VFG.DispID_0082(1, "-")
  loc_11070103:                         frmGzToPzTGZP.VFG.DispID_009E(var_30, 1, var_30, 1, &HFF)
  loc_11070118:                         var_150 = var_30
  loc_110701D4:                         frmGzToPzTGZP.VFG.DispID_0082(&H16, "数据提交错或该数据已经被导入----未引入")
  loc_110701F3:                         var_30 = var_30(1)
  loc_1107021F:                         var_82B8 = CLng(frmGzToPzTGZP.VFG.DispID_0007)
  loc_1107023B:                         var_264 = (var_30 > 0)
  loc_11070262:                         If var_264 = 0 Then GoTo loc_1106FB4D
  loc_11070268:                       End If
  loc_1107026B:                       var_24 = var_30
  loc_1107027F:                       Set var_9C = frmGzToPzTGZP.Chk
  loc_11070281:                       var_264 = var_9C
  loc_11070293:                       Set var_A0 = var_9C(0)
  loc_110702B7:                       var_26C = var_A0
  loc_11070321:                       If (var_A0.Value = 1) Then
  loc_1107041D:                         "unLockVouch".00000004h(var_180, var_BC, var_C4, 0, var_74, var_70, var_58, var_16C, var_54, &H4002, var_184)
  loc_11070426:                       End If
  loc_1107042B:                       var_150 = "VouchNum"
  loc_110704A0:                       Set var_34 = "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C4, 0, var_14C, "VouchNum", var_154)
  loc_110704C1:                       var_150 = "delete  from vouchnum"
  loc_1107051F:                       "DataMdb".00000000h.00000001h(1, 1, var_180, var_BC, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_1107057C:                       frmGzToPzTGZP.Pic1.DispID_80010007 = var_150
  loc_11070590:                       var_82C4 = Resume(0)
  loc_11070596:                     End If
  loc_11070596:                   End If
  loc_11070596:                 End If
  loc_110705B4:                 var_24 = var_27C+(var_24 - 1)
  loc_110705B7:                 GoTo loc_110698BF
  loc_110705BC:               End If
  loc_110705BF:               var_1A0 = var_70
  loc_11070688:               var_D8)
  loc_11070734:               var_82D0 = (var_58 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 3))
  loc_11070741:               var_1F0 = var_82D0 + 1
  loc_110707FD:               var_82D8 = (var_64 = frmGzToPzTGZP.VFG.DispID_0082(var_24, 0))
  loc_1107080A:               var_240 = var_82D8 + 1
  loc_110708A7:               var_82E8 = CBool((frmGzToPzTGZP.VFG.DispID_0082(var_24, 2) = "DateToPeriod".00000001h) And var_82D0 + 1 And var_82D8 + 1)
  loc_110708AD:               var_264 = var_82E8
  loc_1107092C:               If var_264 = 0 Then GoTo loc_11070596
  loc_11070943:               Set var_9C = frmGzToPzTGZP.Chk
  loc_11070945:               var_264 = var_9C
  loc_11070957:               Set var_A0 = var_9C(0)
  loc_1107097B:               var_26C = var_A0
  loc_110709BE:               var_274 = (var_A0.Value = 1)
  loc_110709E9:               var_150 = var_24
  loc_11070A0A:               var_190 = "网络共享冲突----未引入"
  loc_11070A14:               If var_274 = 0 Then
  loc_11070A16:                 var_190 = "指定的凭证号无效或重号----未引入"
  loc_11070A20:               End If
  loc_11070AB1:               frmGzToPzTGZP.VFG.DispID_0082(var_170, var_190)
  loc_11070AD0:               var_24 = var_24(1)
  loc_11070AD6:               var_2A8 = var_24(1)
  loc_11070B05:               var_82EC = CLng(frmGzToPzTGZP.VFG.DispID_0007)
  loc_11070B21:               var_264 = (var_2A8 > 0)
  loc_11070B48:               If var_264 = 0 Then GoTo loc_110705BC
  loc_11070B4E:               GoTo loc_11070596
  loc_11070B53:             End If
  loc_11070B56:             var_1A0 = var_70
  loc_11070C21:             var_D8)
  loc_11070CCF:             var_82F8 = (var_58 = frmGzToPzTGZP.VFG.DispID_0082(var_2A8, 3))
  loc_11070CDC:             var_1F0 = var_82F8 + 1
  loc_11070D9A:             var_8300 = (var_64 = frmGzToPzTGZP.VFG.DispID_0082(var_2A8, 0))
  loc_11070DA7:             var_240 = var_8300 + 1
  loc_11070E44:             var_8310 = CBool((frmGzToPzTGZP.VFG.DispID_0082(var_2A8, 2) = "DateToPeriod".00000001h) And var_82F8 + 1 And var_8300 + 1)
  loc_11070E4A:             var_264 = var_8310
  loc_11070EC9:             If var_264 = 0 Then GoTo loc_11070596
  loc_11070FC0:             If (frmGzToPzTGZP.VFG.DispID_0082(var_2A8, &H16) = global_1100AE28) + 1 Then
  loc_11070FC6:               var_150 = var_2A8
  loc_1107107F:               Set var_9C = frmGzToPzTGZP.VFG
  loc_11071082:               var_9C.DispID_0082(&H16, "凭证借贷不平衡或某分录有错误----未引入")
  loc_11071093:               GoTo loc_11070B53
  loc_11071098:             End If
  loc_11071162:             var_C0 = frmGzToPzTGZP.VFG.DispID_0082(frmGzToPzTGZP.VFG, &H16) & "----未引入"
  loc_110711FF:             frmGzToPzTGZP.VFG.DispID_0082(&H16, var_C0)
  loc_1107123C:             GoTo loc_11070B53
  loc_11071241:           End If
  loc_11071289:           frmGzToPzTGZP.Pic1.DispID_80010007 = var_150
  loc_110712A0:           If var_2C Then
  loc_11071332:             MsgBox("数据引入已完成，数据已生成用友凭证。", 64, "提示信息", 10, 10)
  loc_110713A4:             frmGzToPzTGZP.VFG.DispID_0007 = 1
  loc_1107143F:             frmGzToPzTGZP.sBar.DispID_6803001E(1100AE28h)
  loc_110714D6:             frmGzToPzTGZP.sBar.DispID_6803001E(1100AE28h)
  loc_1107156D:             Set var_9C = frmGzToPzTGZP.sBar
  loc_11071570:             var_9C.DispID_6803001E(1100AE28h)
  loc_11071586:           Else
  loc_1107160D:             MsgBox("数据没有被引入，原因请查看最后一列中的说明。", 64, "提示信息", 10, 10)
  loc_1107163A:           End If
  loc_1107163F:           var_150 = "VouchNum"
  loc_110716B6:           Set var_34 = "DataMdb".00000000h.00000001h(var_180, var_BC, var_C0, var_C4, var_C8, var_14C, "VouchNum", var_154)
  loc_110716D7:           var_150 = "delete  from vouchnum"
  loc_1107172B:           "DataMdb".00000000h.00000001h(1, var_180, var_BC, var_C0, var_C4, var_14C, "delete  from vouchnum", var_154)
  loc_11071780:           "AS13")
  loc_110717A0:           var_24C = frmGzToPzTGZP.UpdateBTData
  loc_110717E9:           var_B8 = Global.Screen
  loc_11071807:           var_8330 = ecx
  loc_1107180F:           var_8330 = var_9C.UnkVCall_0000007Ch
  loc_11071823:         End If
  loc_11071823:       End If
  loc_11071823:     End If
  loc_11071823:   End If
  loc_11071823: End If
  loc_1107182F: Exit Sub
  loc_1107183B: GoTo loc_110718F4
  loc_110718F3: Exit Sub
  loc_110718F4: ' Referenced from: 11068D1C
  loc_110718F4: ' Referenced from: 1107183B
End Sub

Private Sub Proc_12_16_110723D0
  Dim var_58 As Variant
  Dim var_5C As Variant
  Dim var_64 As frmGzToPzTGZP.Label3
  Dim var_1D0 As Label
  loc_110724BD: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110724C6: var_1F0 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110724E3: Set var_58 = frmGzToPzTGZP.Chk
  loc_110724ED: var_1D0 = var_58
  loc_110724F3: Set var_5C = var_58(0)
  loc_1107251E: var_1D8 = var_5C
  loc_11072561: var_1E0 = (var_5C.Value = 1)
  loc_11072577: If var_1E0 = 0 Then
  loc_110725DC:   If var_14 <= CLng(frmGzToPzTGZP.VFG.DispID_0007)(-1) Then
  loc_11072651:     var_7C = frmGzToPzTGZP.VFG.DispID_0082(var_14, 2)
  loc_1107266C:     var_94)
  loc_110726C7:     var_30 = CByte("DateToPeriod".00000001h)
  loc_11072821:     Set var_64 = frmGzToPzTGZP.Label3
  loc_1107284B:     var_1D0 = var_64
  loc_11072A01:     var_94 = frmGzToPzTGZP.VFG.DispID_0082(var_14, frmGzToPzTGZP.VFG)
  loc_11072A1D:     var_8034 = "正在处理：第[" & frmGzToPzTGZP.VFG.DispID_0082(var_14, 2) & " - " & frmGzToPzTGZP.VFG.DispID_0082(var_14, 3) & " - " & var_94
  loc_11072A53:     var_64.Caption = var_8034 & "]号凭证是否重号"
  loc_11072AE2:     var_803C = frmGzToPzTGZP.Proc_12_17_110731E0(var_30)
  loc_11072AF7:     If var_1CC <= 0 Then
  loc_11072B09:       var_13C = var_30
  loc_11072BA0:       var_94)
  loc_11072C39:       var_804C = (frmGzToPzTGZP.VFG.DispID_0082(var_14, 3) = frmGzToPzTGZP.VFG.DispID_0082(var_14, 3))
  loc_11072C66:       var_17C = var_804C + 1
  loc_11072CDD:       var_8054 = (frmGzToPzTGZP.VFG.DispID_0082(var_14, frmGzToPzTGZP.VFG) = frmGzToPzTGZP.VFG.DispID_0082(var_14, ""))
  loc_11072D04:       var_1BC = var_8054 + 1
  loc_11072DFF:       If CBool((frmGzToPzTGZP.VFG.DispID_0082(var_14, 2) = "DateToPeriod".00000001h) And var_804C + 1 And var_8054 + 1) Then
  loc_11072E92:         frmGzToPzTGZP.VFG.DispID_0082(var_10C, 285267820)
  loc_11072FC6:         frmGzToPzTGZP.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_1107305A:         frmGzToPzTGZP.VFG.DispID_0082(var_10C, "指定的凭证号无效或重号")
  loc_110730A5:         var_8068 = CLng(frmGzToPzTGZP.VFG.DispID_0007)
  loc_110730C3:         var_1D0 = (var_14(1) > 0)
  loc_110730E0:         If var_1D0 = 0 Then GoTo loc_11072B03
  loc_110730E6:       End If
  loc_110730F4:     Else
  loc_110730FD:     End If
  loc_1107310A:     var_14 = 1+var_14
  loc_1107310D:     GoTo loc_110725D6
  loc_11073112:   End If
  loc_11073112: End If
  loc_11073117: GoTo loc_110731A8
  loc_110731A7: Exit Sub
  loc_110731A8: ' Referenced from: 11073117
End Sub

Private  Proc_12_17_110731E0(arg_C, arg_10, arg_14) '110731E0
  loc_11073279: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11073282: var_168 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110732AB: If IsNumeric(arg_14) Then
  loc_110732BA:   var_8008 = CLng(Val(arg_14))
  loc_110732C4:   If var_8008 > 0 Then
  loc_110732D0:     If var_8008 <= 9999 Then
  loc_1107334C:       var_8028 = "select * from GL_accvouch where iperiod >=" & CStr(arg_C) & " and isignseq>=" & CStr(0) & " and ino_id>=" & CStr(var_8008)
  loc_11073361:       var_44 = var_8028
  loc_110733B3:       Set var_1C = "DataMdb".00000000h.00000001h(fs:[00000000h], , , , , var_40, var_8028, var_48)
  loc_110733F8:       var_8030 = Proc_0_4_11026BD0(arg_10, , )
  loc_11073419:       var_8034 = CBool(var_1C.EOF)
  loc_1107342D:       If var_8034 = 0 Then
  loc_11073458:         var_F4 = arg_C
  loc_11073516:         var_8040 = (var_1C.DispID_0000("iPeriod") = arg_C) And (var_1C.DispID_0000("isignseq") = (Proc_0_4_11026BD0(arg_10, , ) And 255))
  loc_11073586:         var_804C = CBool(Not(var_8040 And (var_1C.DispID_0000("ino_id") = var_8008)))
  loc_110735AB:         If var_804C = 0 Then GoTo loc_110735B0
  loc_110735AD:       End If
  loc_110735BB:       var_1C.oClose
  loc_110735C4:     End If
  loc_110735C4:   End If
  loc_110735C4: End If
  loc_110735CA: GoTo loc_1107362F
  loc_1107362E: Exit Sub
  loc_1107362F: ' Referenced from: 110735CA
End Sub
