VERSION 5.00
Object = "{00000000-0000-0000-0000-000000000000}##0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C908002B2F49FB}#1.2#0"; "C:\Windows\SysWow64\Comdlg32.ocx"
Begin VB.Form frmSaleBillDr
  Caption = "销售发票导入"
  BackColor = &H80000005&
  WindowState = 2
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  Icon = "frmSaleBillDr.frx":0000
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
    OleObjectBlob = "frmSaleBillDr.frx":014A
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
    Width = 10515
    Height = 5700
    TabStop = 0   'False
    TabIndex = 0
    OleObjectBlob = "frmSaleBillDr.frx":0293
    Begin AIFCmp1.asxStatusBar sBar
      Left = 0
      Top = 5355
      Width = 10515
      Height = 345
      OleObjectBlob = "frmSaleBillDr.frx":04BC
    End
    Begin C1SizerLibCtl.C1Elastic C1Elastic2
      Left = 0
      Top = 0
      Width = 10515
      Height = 435
      TabStop = 0   'False
      TabIndex = 1
      OleObjectBlob = "frmSaleBillDr.frx":05EC
    End
    Begin VSFlex8LCtl.VSFlexGrid VFG
      Left = 0
      Top = 1140
      Width = 10515
      Height = 4200
      TabIndex = 2
      OleObjectBlob = "frmSaleBillDr.frx":0743
      Begin MSComDlg.CommonDialog dlg
        OleObjectBlob = "frmSaleBillDr.frx":0BAC
        Left = 0
        Top = 0
      End
    End
    Begin AIFCmp1.asxPanel asxPanel1
      Left = 0
      Top = 450
      Width = 10515
      Height = 675
      OleObjectBlob = "frmSaleBillDr.frx":0C10
      Begin TDBDate6Ctl.TDBDate TDBDate
        Left = 4680
        Top = 360
        Width = 2385
        Height = 285
        TabIndex = 9
        OleObjectBlob = "frmSaleBillDr.frx":0CF0
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 0
        Left = 7155
        Top = 315
        Width = 930
        Height = 330
        TabIndex = 5
        OleObjectBlob = "frmSaleBillDr.frx":0FDF
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 1
        Left = 8160
        Top = 315
        Width = 930
        Height = 330
        TabIndex = 6
        OleObjectBlob = "frmSaleBillDr.frx":11D7
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 0
        Left = 120
        Top = 45
        Width = 6945
        Height = 270
        TabIndex = 7
        OleObjectBlob = "frmSaleBillDr.frx":13A7
        ToolTipText = "项目大类"
      End
      Begin AIFCmp1.asxPowerButton APB
        Index = 2
        Left = 9165
        Top = 315
        Width = 1080
        Height = 330
        Visible = 0   'False
        TabIndex = 8
        OleObjectBlob = "frmSaleBillDr.frx":1503
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 1
        Left = 90
        Top = 360
        Width = 2295
        Height = 270
        TabIndex = 10
        OleObjectBlob = "frmSaleBillDr.frx":16A7
        ToolTipText = "部门编码"
      End
      Begin TDBText6Ctl.TDBText TDBText
        Index = 2
        Left = 2400
        Top = 360
        Width = 2235
        Height = 270
        TabIndex = 11
        OleObjectBlob = "frmSaleBillDr.frx":1807
        ToolTipText = "本单位开户银行编码"
      End
    End
  End
End

Attribute VB_Name = "frmSaleBillDr"


Private  APB_UnknownEvent_9(arg_C) '110A3000
  Dim var_2C As Variant
  Dim var_1C As ADODB.Recordset
  loc_110A3089: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110A3092: var_F4 = __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110A30B9: arg_C = frmSaleBillDr.APB.UnkVCall_00000040h
  loc_110A30FD: var_E0 = var_30.DispID_FFFFFDFA
  loc_110A3127: var_8008 = (var_E0 = "加载数据")
  loc_110A312F: If var_8008 = 0 Then
  loc_110A3160:   Set var_2C = frmSaleBillDr.TDBText
  loc_110A316E:   var_CC = var_2C
  loc_110A3174:   var_2C.UnkVCall_00000040h
  loc_110A31AB:   var_F8 = var_24
  loc_110A31C6:   var_28 = var_30.DispID_0000
  loc_110A31D6:   var_8014 = Scripting.FileSystemObject.FileExists
  loc_110A321E:   If Not (var_C4) Then
  loc_110A328D:     MsgBox("文件不存在或非法路径！ ", 64, "提示", 10, 10)
  loc_110A32BD:   Else
  loc_110A32CD:     If var_18 <= 2 Then
  loc_110A32F4:       var_18 = frmSaleBillDr.TDBText.UnkVCall_00000040h
  loc_110A332B:       var_48 = var_30.DispID_0000
  loc_110A3386:       If (Proc_0_11_11029000(8, var_30, global_110A38B8) = global_1100AE28) + 1 = 0 Then
  loc_110A3397:         var_18 = 1+var_18
  loc_110A339A:         GoTo loc_110A32C4
  loc_110A339F:       End If
  loc_110A33C0:       var_18 = frmSaleBillDr.TDBText.UnkVCall_00000040h
  loc_110A3439:       var_28 = var_30.DispID_8001004A
  loc_110A345F:       MsgBox(var_28 & "不能为空，请输入。 ", 64, "提示", 10, var_80)
  loc_110A34C0:       var_18 = frmSaleBillDr.TDBText.UnkVCall_00000040h
  loc_110A34E3:       var_30.DispID_80011000
  loc_110A34F8:       GoTo loc_110A32AE
  loc_110A34FD:     End If
  loc_110A3563:     If (Proc_0_11_11029000(frmSaleBillDr.TDBDate.DispID_0000, var_30, var_30) = global_1100AE28) + 1 Then
  loc_110A35D2:       MsgBox("请输入制单日期！ ", 64, "提示", 10, 10)
  loc_110A3618:       var_C8 = ADODB.Recordset.State
  loc_110A363D:       If var_C8 = 1 Then
  loc_110A365D:         var_803C = ADODB.Recordset.Close
  loc_110A367B:       End If
  loc_110A3690:       Set var_2C = frmSaleBillDr.TDBDate
  loc_110A3697:       var_2C.DispID_80011000
  loc_110A36A9:       GoTo loc_110A32AE
  loc_110A36AE:     End If
  loc_110A36C0:     If global_0008.FillData >= 0 Then GoTo loc_110A32AE
  loc_110A36D2:     var_C4 = CheckObj(8, global_1100D0B8, 1788)
  loc_110A36DD:   End If
  loc_110A36E9:   var_8040 = (var_E0 = "取消加载")
  loc_110A36F1:   If var_8040 = 0 Then
  loc_110A372A:     var_50 = "提示信息"
  loc_110A373A:     var_8044 = "是否取消数据载入？" & vbCrLf
  loc_110A3745:     call ebx(var_28, var_C4, var_2C, 0, var_30, var_30)
  loc_110A3756:     var_38 = ebx(var_28, var_C4, var_2C, 0, var_30, var_30) & "取消数据载入，数据将全部清空。"
  loc_110A3772:     MsgBox(var_38, 292, var_50, var_60, var_70)
  loc_110A37A9:     If (MsgBox(var_38, 292, var_50, var_60, var_70) = 6) = var_1C Then GoTo loc_110A32B0
  loc_110A37B5:     GoTo loc_110A32B0
  loc_110A37BA:   End If
  loc_110A37CC:   var_804C = (var_E0 = "发票导入")
  loc_110A37D0:   If var_804C = 0 Then
  loc_110A37D5:     var_8050 = frmSaleBillDr.Proc_14_11_1109D450
  loc_110A37DB:     GoTo loc_110A32B0
  loc_110A37E0:   End If
  loc_110A37F0:   If (var_E0 = global_1100EBD4) Then GoTo loc_110A32B0
  loc_110A3827:   Set var_2C = CInt(8)
  loc_110A3835:   var_805C = Global.Unload
  loc_110A3856:   GoTo loc_110A32B0
  loc_110A3893:   Exit Sub
  loc_110A3894: End If
End Sub

Private  TDBText_UnknownEvent_B(arg_C) '110A38E0
  Dim var_6C As frmSaleBillDr.dlg
  loc_110A393D: If arg_C = 0 Then
  loc_110A3959:   Set var_6C = frmSaleBillDr.dlg
  loc_110A398B:   var_6C.FileName = var_4C
  loc_110A39AD:   var_6C.DialogTitle = var_4C
  loc_110A39CF:   var_6C.Filter = var_4C
  loc_110A39EE:   var_6C.CancelError = var_4C
  loc_110A39F8:   var_6C.ShowOpen
  loc_110A3A10:   var_6C.FileName = var_6C
  loc_110A3A52:   If (var_30 = global_1100AE28) Then
  loc_110A3A64:     var_6C.FileName = Me
  loc_110A3AA1:     arg_C = frmSaleBillDr.TDBText.UnkVCall_00000040h
  loc_110A3B0C:   End If
  loc_110A3B18: Else
  loc_110A3B1E:   GoTo loc_110A3B0E
  loc_110A3B4C:   Exit Sub
  loc_110A3B4D: End If
End Sub

Private Sub Form_Load() '1109A780
  Dim var_1C As Variant
  Dim var_24 As var_20.DispID_03E8
  Dim var_20 As var_1C.DispID_03E8
  loc_1109A7E6: If var_18 <= 2 Then
  loc_1109A809:   var_18 = frmSaleBillDr.TDBText.UnkVCall_00000040h
  loc_1109A835:   var_34 = var_20.DispID_03E8
  loc_1109A84A:   Set var_24 = var_20.DispID_03E8
  loc_1109A89D:   var_18 = 1+var_18
  loc_1109A8A2:   GoTo loc_1109A7DD
  loc_1109A8A7: End If
  loc_1109A8C0: Set var_1C = frmSaleBillDr.TDBDate
  loc_1109A8C7: var_34 = var_1C.DispID_03E8
  loc_1109A8DC: Set var_20 = var_1C.DispID_03E8
  loc_1109A8E8: var_20.UnkVCall_00000030h
  loc_1109A957: frmSaleBillDr.TDBDate.DispID_0000 = Date
  loc_1109A97B: Set var_1C = frmSaleBillDr.APB
  loc_1109A988: var_1C.UnkVCall_00000040h
  loc_1109A9E0: var_8004 = frmSaleBillDr.Proc_14_8_11096450(var_1C)
  loc_1109A9ED: var_58 = frmSaleBillDr.getBTData
  loc_1109AA15: GoTo loc_1109AA38
  loc_1109AA37: Exit Sub
  loc_1109AA38: ' Referenced from: 1109AA15
End Sub

Private Sub Form_Resize() '1109AA60
  loc_1109AAED: var_38 = frmSaleBillDr.Pic1.DispID_80010005
  loc_1109AB11: var_48 = frmSaleBillDr.Pic1.DispID_80010006
  loc_1109AB24: var_EC = var_48.ScaleWidth
  loc_1109AB5B: If global_110F6000 = 0 Then
  loc_1109AB65: Else
  loc_1109AB70: End If
  loc_1109AB70: var_70 = ((var_EC - CSgn(var_38)) / 2)
  loc_1109AB85: var_F0 = var_48.ScaleHeight
  loc_1109ABC3: If global_110F6000 = 0 Then
  loc_1109ABCD: Else
  loc_1109ABD8: End If
  loc_1109ACE3: frmSaleBillDr.Pic1.DispID_80011002(var_70, ((var_F0 - CSgn(var_48)) / 2), CSgn(frmSaleBillDr.Pic1.DispID_80010005), CSgn(frmSaleBillDr.Pic1.DispID_80010006))
  loc_1109AD2C: GoTo loc_1109AD66
End Sub

Public Sub ExitForm(Cancel, UnloadMode) '11096370
  Dim var_18 As Global
  loc_110963AF: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110963DA: Set var_18 = Me
  loc_110963E2: var_8008 = Global.Unload
  loc_1109641C: GoTo loc_11096428
  loc_11096427: Exit Sub
  loc_11096428: ' Referenced from: 1109641C
End Sub

Public Function FillData() '110975A0
  Dim var_64 As Variant
  Dim var_44 As Variant
  Dim var_40 As Variant
  Dim var_38 As Variant
  loc_110976C1: frmSaleBillDr.VFG.DispID_0007 = 1
  loc_110976DE: Set var_64 = frmSaleBillDr.Label3
  loc_110976E0: var_1F8 = var_64
  loc_110976EE: var_64.Caption = "正在打开Excel数据表，请稍候。。。"
  loc_1109775B: frmSaleBillDr.Pic1.DispID_80010007 = True
  loc_11097781: frmSaleBillDr.Pic1.DispID_FFFFFDDA
  loc_110977A1: Set var_64 = frmSaleBillDr.TDBText
  loc_110977AE: var_64.UnkVCall_00000040h
  loc_110977CC: var_78 = var_68
  loc_110977EA: var_24 = Proc_0_11_11029000(var_80, var_64, 1)
  loc_11097803: var_8008 = CreateObject(global_1100D5A4)
  loc_1109780E: Set var_44 = CreateObject(global_1100D5A4)
  loc_1109789C: Set var_64 = frmSaleBillDr.TDBText
  loc_110978A2: var_1F8 = var_64
  loc_110978B0: var_64.UnkVCall_00000040h
  loc_11097B0D: var_80 = var_68.DispID_0000
  loc_11097B22: var_54 = var_80
  loc_11097B32: var_54 = var_44.UnkVCall_000000D0h.UnkVCall_0000004Ch
  loc_11097B94: var_64 = 0.Tag
  loc_11097C30: var_64.Activate
  loc_11097C76: var_80 = var_64.UsedRange
  loc_11097C8B: Set var_4C = var_80
  loc_11097CB1: Set var_64 = frmSaleBillDr.Label3
  loc_11097CB3: var_1F8 = var_64
  loc_11097CC1: var_64.Caption = "正在填充数据，请稍候。。。"
  loc_11097D2E: frmSaleBillDr.Pic1.DispID_80010007 = True
  loc_11097D55: frmSaleBillDr.Pic1.DispID_FFFFFDDA
  loc_11097D89: Set var_64 = frmSaleBillDr.APB
  loc_11097D8B: var_1F8 = var_64
  loc_11097D9A: var_64.UnkVCall_00000040h
  loc_11097E24: Set var_64 = frmSaleBillDr.APB
  loc_11097E26: var_1F8 = var_64
  loc_11097E35: var_64.UnkVCall_00000040h
  loc_11097ECC: frmSaleBillDr.APB.UnkVCall_00000040h
  loc_11097F85: var_80 = var_4C.Rows
  loc_11098046: Set var_64 = frmSaleBillDr.sBar
  loc_1109804D: var_64.DispID_6803001E(1100D68Ch & var_80.Count - 2 & "条记录")
  loc_1109808E: On Error GoTo loc_1109A1C9
  loc_110980E9: var_38.UnkVCall_00000064h
  loc_11098168: var_80 = var_64.Cells(1, 1)
  loc_11098179: var_90 = var_80.value
  loc_11098279: var_80 = var_64.Cells(1, 2)
  loc_1109829E: var_2C = Proc_0_11_11029000(var_80.value, var_64, var_11C)
  loc_110982EC: var_80 = var_4C.Rows
  loc_110982FD: var_90 = var_80.Count
  loc_11098351: If var_20 <= CLng(var_90 + 1) Then
  loc_1109835F:   If global_56 = 0 Then
  loc_110983C9:     var_90.BackColor = var_118
  loc_1109844A:     var_80 = var_64.Cells(var_20, 1)
  loc_11098484:     var_1FC = (Proc_0_11_11029000(var_80.value, var_64, var_118) = global_1100AE28) + 1
  loc_110984B7:     If var_1FC = 0 Then
  loc_11098568:       frmSaleBillDr.sBar.DispID_6803001E("正在导入数据：" & CStr(vbNull) & "条记录")
  loc_110985BA:       var_80 = Chr(9)
  loc_11098639:       var_80 = Chr(9)
  loc_11098654:       Set var_64 = frmSaleBillDr.TDBDate
  loc_110986FE:       var_80 = Chr(9)
  loc_1109873A:       var_A0 = "1" & var_80 & 1100AE28h & var_80 & var_64.DispID_004E & var_80 & Proc_0_11_11029000("1" & var_80 & 1100AE28h & var_80, var_38, 2)
  loc_1109877C:       var_80 = Chr(9)
  loc_110987FA:       var_80 = Chr(9)
  loc_11098857:       var_38.UnkVCall_00000064h
  loc_110988EA:       var_B0 = var_64.Cells(var_20, 4).value
  loc_1109898C:       var_80 = Chr(9)
  loc_11098A7C:       var_B0 = var_64.Cells(var_20, 5).value
  loc_11098ABE:       var_D0 = var_64.Cells(var_20, 4) & var_80 & var_24 & var_80 & var_24 & var_80 & Proc_0_11_11029000(var_B0, var_38, 2) & var_80 & Proc_0_11_11029000(var_B0, var_64, var_12C)
  loc_11098B1E:       var_80 = Chr(9)
  loc_11098B7B:       var_B0.BackColor = var_128
  loc_11098C0E:       var_B0 = var_64.Cells(var_20, 6).value
  loc_11098D81:       var_B0 = var_64.Cells(var_20, 10).value
  loc_11098E6A:       var_D0 = var_68.Cells(var_20, 6).value
  loc_11098E83:       var_118 = var_D0 & var_80 & Proc_0_11_11029000(var_B0, var_64, 1)
  loc_11098E99:       var_80 = Chr(9)
  loc_11098EBF:       var_F0 = "0.000000"
  loc_11098EE5:       var_260 = Proc_0_12_110291B0(var_B0, var_64, var_124)
  loc_11098EF4:       var_58 = Proc_0_12_110291B0(var_D0, var_68, var_64)
  loc_11098F04:       If global_110F6000 = 0 Then
  loc_11098F0E:       Else
  loc_11098F1F:       End If
  loc_110990E7:       var_B0 = var_64.Cells(var_20, 10).value
  loc_110990FB:       var_54 = Proc_0_12_110291B0(var_B0, var_64, 1)
  loc_11099116:       var_80 = Chr(9)
  loc_1109919E:       var_F0 = var_118 & var_80 & Format((var_260 / var_58), var_F0) & var_80 & Format((var_260 / var_58), var_F0) & var_80 & Format(0, "0.00")
  loc_110992E6:       var_B0 = var_64.Cells(var_20, 12).value
  loc_110993CF:       var_D0 = var_68.Cells(var_20, 6).value
  loc_110993E8:       var_118 = var_F0
  loc_110993FE:       var_80 = Chr(9)
  loc_11099424:       var_F0 = "0.000000"
  loc_1109944A:       var_268 = Proc_0_12_110291B0(var_B0, var_64, 1)
  loc_11099459:       var_58 = Proc_0_12_110291B0(var_D0, var_68, 1)
  loc_11099469:       If global_110F6000 = 0 Then
  loc_11099473:       Else
  loc_11099484:       End If
  loc_1109964C:       var_B0 = var_64.Cells(var_20, 12).value
  loc_11099660:       var_54 = Proc_0_12_110291B0(var_B0, var_64, 1)
  loc_1109967B:       var_80 = Chr(9)
  loc_11099703:       var_F0 = var_118 & var_80 & Format((var_268 / var_58), var_F0) & var_80 & Format((var_268 / var_58), var_F0) & var_80 & Format(0, "0.00")
  loc_1109984B:       var_B0 = var_64.Cells(var_20, 12).value
  loc_11099934:       var_D0 = var_68.Cells(var_20, 10).value
  loc_11099963:       var_80 = Chr(9)
  loc_11099AC4:       var_80 = Chr(9)
  loc_11099BE5:       var_90 = "0.00" & var_80 & Format((Proc_0_12_110291B0(var_64.Cells(var_20, 13).value, var_64, 1) - Proc_0_12_110291B0(var_D0, var_68, 1)), "0.00") & var_80
  loc_11099C89:       frmSaleBillDr.VFG.DispID_0080(var_90 & Proc_0_11_11029000(frmSaleBillDr.VFG.Cells(var_20, 13).value, frmSaleBillDr.VFG, 1))
  loc_11099CA7:       var_28 = var_28(1)
  loc_11099CB7:       If var_20 Mod 00000064h = 0 Then
  loc_11099CB9:         DoEvents
  loc_11099CBF:       End If
  loc_11099CCD:       var_20 = 1+var_20
  loc_11099CD0:       GoTo loc_1109834B
  loc_11099CD5:     End If
  loc_11099D23:     frmSaleBillDr.VFG.DispID_0007 = 1
  loc_11099D32:     global_56 = 0
  loc_11099D67:     frmSaleBillDr.APB.UnkVCall_00000040h
  loc_11099DFC:     frmSaleBillDr.APB.UnkVCall_00000040h
  loc_11099E82:     Set var_64 = frmSaleBillDr.APB
  loc_11099E93:     var_64.UnkVCall_00000040h
  loc_11099E9A:     If var_64.UnkVCall_00000040h < 0 Then
  loc_11099EA0:       GoTo loc_1109A010
  loc_11099EA5:     End If
  loc_11099EDB:     frmSaleBillDr.APB.UnkVCall_00000040h
  loc_11099F24:     var_68.DispID_80010007 = var_118
  loc_11099F70:     frmSaleBillDr.APB.UnkVCall_00000040h
  loc_11099FF6:     Set var_64 = frmSaleBillDr.APB
  loc_1109A007:     var_64.UnkVCall_00000040h
  loc_1109A00E:     If var_64.UnkVCall_00000040h < 0 Then
  loc_1109A010:       ' Referenced from: 11099EA0
  loc_1109A019:       var_64.UnkVCall_00000040h = CheckObj(var_64, global_1100D678, 64)
  loc_1109A01F:     End If
  loc_1109A01F:   End If
  loc_1109A050:   var_68.DispID_80010007 = var_118
  loc_1109A065: End If
  loc_1109A10F: frmSaleBillDr.sBar.DispID_6803001E("有效数据共" & CStr(var_28) & global_1100FE7C)
  loc_1109A1B7: frmSaleBillDr.sBar.DispID_6803001E(1100AE28h)
  loc_1109A1C9: ' Referenced from: 1109808E
  loc_1109A201: frmSaleBillDr.APB.UnkVCall_00000040h
  loc_1109A250: var_68.DispID_80010007 = var_118
  loc_1109A287: Set var_64 = frmSaleBillDr.APB
  loc_1109A289: var_1F8 = var_64
  loc_1109A298: var_64.UnkVCall_00000040h
  loc_1109A31E: Set var_64 = frmSaleBillDr.APB
  loc_1109A320: var_1F8 = var_64
  loc_1109A32F: var_64.UnkVCall_00000040h
  loc_1109A3D3: frmSaleBillDr.Pic1.DispID_80010007 = var_118
  loc_1109A3EC: Set var_64 = frmSaleBillDr.TDBText
  loc_1109A3F9: var_64.UnkVCall_00000040h
  loc_1109A42F: var_78 = var_68
  loc_1109A4B4: var_64.ForeColor = False
  loc_1109A4E7: var_11C = var_44.UnkVCall_00000398h
  loc_1109A51C: Set var_38 = {000208D7-0000-0000-C000000000000046}()
  loc_1109A52C: Set var_40 = {000208DA-0000-0000-C000000000000046}()
  loc_1109A53C: Set var_44 = {000208D5-0000-0000-C000000000000046}()
  loc_1109A550: var_80B0 = Err
  loc_1109A557: Set var_64 = Err
  loc_1109A59A: If (Err.Number) Then
  loc_1109A5A0:   var_80B4 = Err
  loc_1109A5A7:   Set var_64 = Err
  loc_1109A5B2:   var_54 = Err.Description
  loc_1109A63E:   MsgBox(0, 16, "提示", 10, 10)
  loc_1109A671: End If
  loc_1109A671: Exit Sub
  loc_1109A67D: GoTo loc_1109A6FE
  loc_1109A6FD: Exit Function
  loc_1109A6FE: ' Referenced from: 1109A67D
End Function

Public Function getBTData() '110A3B80
  Dim var_24 As ADODB.Recordset
  Dim var_38 As Variant
  loc_110A3C04: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110A3C0E: On Error GoTo loc_110A4185
  loc_110A3C49: var_28 = 1 & "IF NOT EXISTS (SELECT * FROM [" & "]..Sysobjects "
  loc_110A3C8F: var_8014 =  & var_28 & "WHERE Name = 'T_CY_XSFPDr_Setting') " & "CREATE TABLE [" & "]..[T_CY_XSFPDr_Setting](cDepCode VARCHAR(50) NULL,cBankCode VARCHAR(50) NULL)"
  loc_110A3C96: var_28 = var_8014
  loc_110A3CD9: var_54 = UnkObj.UnkVCall_00000040h
  loc_110A3D2B: var_28 = var_38 & "SELECT * FROM [" & "]..[T_CY_XSFPDr_Setting]"
  loc_110A3D65: var_BC = ADODB.Recordset.State
  loc_110A3D8A: If var_BC = 1 Then
  loc_110A3DA6:   var_8028 = ADODB.Recordset.Close
  loc_110A3DC4: End If
  loc_110A3E44: var_8030 = ADODB.Recordset.Open(var_28, var_90, var_28, var_88, 9)
  loc_110A3E97: var_B8 = ADODB.Recordset.EOF
  loc_110A3EB3: If var_B8 = 0 Then
  loc_110A3EDB:   var_38 = ADODB.Recordset.Fields
  loc_110A3EF9:   var_8C = "cDepCode"
  loc_110A3F2D:   ADODB.Recordset.8 = Forms
  loc_110A3F98:   frmSaleBillDr.TDBText.UnkVCall_00000040h
  loc_110A3FCE:   var_44.DispID_0000 = Proc_0_11_11029000(9, var_90, "cDepCode")
  loc_110A4034:   var_D0 = ADODB.Recordset.Fields
  loc_110A403F:   var_8C = "cBankCode"
  loc_110A4073:   ADODB.Recordset.8 = Forms
  loc_110A40E1:   frmSaleBillDr.TDBText.UnkVCall_00000040h
  loc_110A4117:   var_44.DispID_0000 = Proc_0_11_11029000(9, var_90, "cBankCode")
  loc_110A4144: End If
  loc_110A416C: If ADODB.Recordset.Close < 0 Then
  loc_110A417E:   var_804C = CheckObj(var_24, global_1100ADFC, 128)
  loc_110A4185:   ' Referenced from: 110A3C0E
  loc_110A418A:   var_8050 = Err
  loc_110A4195:   Set var_38 = Err
  loc_110A421A:   MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_110A4247: End If
  loc_110A4247: Exit Sub
  loc_110A4252: GoTo loc_110A429B
  loc_110A429A: Exit Function
  loc_110A429B: ' Referenced from: 110A4252
End Function

Public Function UpdateBTData() '110A42E0
  Dim var_3C As Variant
  Dim var_44 As frmSaleBillDr.TDBText
  Dim var_20 As Me
  loc_110A4358: call __vbaAptOffset(global_1100A430, Me, 0, 0, 0)
  loc_110A4362: On Error GoTo loc_110A45EE
  loc_110A439D: var_20 = 1 & "DELETE FROM [" & "]..[T_CY_XSFPDr_Setting]"
  loc_110A43D6: var_58 = UnkObj.UnkVCall_00000040h
  loc_110A4456: Set var_3C = frmSaleBillDr.TDBText
  loc_110A445C: var_BC = var_3C
  loc_110A446B: var_3C.UnkVCall_00000040h
  loc_110A44AE: Set var_44 = frmSaleBillDr.TDBText
  loc_110A44B4: var_C4 = var_44
  loc_110A44C3: var_44.UnkVCall_00000040h
  loc_110A44E4: var_48 = 0
  loc_110A44EB: var_60 = var_48
  loc_110A4519: var_8020 = 2 & Proc_0_10_11028DD0(9, var_3C & "INSERT INTO [" & "]..[T_CY_XSFPDr_Setting]" & "(cDepCode,cBankCode) VALUES (", var_44) & global_1100AC40
  loc_110A454A: var_20 = var_3C & Proc_0_10_11028DD0(9, var_8020, var_48) & global_1100BD88
  loc_110A45E9: GoTo loc_110A46B0
  loc_110A45EE: ' Referenced from: 110A4362
  loc_110A45F3: var_8030 = Err
  loc_110A45FE: Set var_3C = Err
  loc_110A4683: MsgBox(Err.Description, 16, "提示信息", 10, 10)
  loc_110A46B0: ' Referenced from: 110A45E9
  loc_110A46B0: Exit Sub
  loc_110A46BB: GoTo loc_110A4710
  loc_110A470F: Exit Function
  loc_110A4710: ' Referenced from: 110A46BB
End Function

Private Sub Proc_14_8_11096450
  Dim var_68 As frmSaleBillDr.VFG
  loc_11096492: Set var_68 = frmSaleBillDr.VFG
  loc_110964E3: var_68.DispID_005D = frmSaleBillDr.VFG
  loc_11096524: var_68.DispID_0067 = frmSaleBillDr.VFG
  loc_11096543: var_68.DispID_0041 = frmSaleBillDr.VFG
  loc_110965A3: var_68.DispID_0047 = frmSaleBillDr.VFG
  loc_110966B1: var_68.DispID_008A(4)
  loc_110966F4: var_68.DispID_0079(450)
  loc_11096734: var_68.DispID_007B(True)
  loc_11096779: var_68.DispID_0090("单号")
  loc_110967BC: var_68.DispID_0077(4)
  loc_110967FF: var_68.DispID_0078(600)
  loc_11096847: var_68.DispID_0090("状态")
  loc_1109688D: var_68.DispID_0077(4)
  loc_110968D3: var_68.DispID_0078(500)
  loc_1109691B: var_68.DispID_0090("制单日期")
  loc_11096961: var_68.DispID_0077(1)
  loc_110969A7: var_68.DispID_0078(1000)
  loc_110969EC: var_68.DispID_0090("客户号")
  loc_11096A2E: var_68.DispID_0077(1)
  loc_11096A70: var_68.DispID_0078(700)
  loc_11096AB8: var_68.DispID_0090("部门")
  loc_11096AFE: var_68.DispID_0077(1)
  loc_11096B44: var_68.DispID_0078(700)
  loc_11096B8C: var_68.DispID_0090("品番")
  loc_11096BD2: var_68.DispID_0077(1)
  loc_11096C18: var_68.DispID_0078(1300)
  loc_11096C60: var_68.DispID_0090("品名")
  loc_11096CA6: var_68.DispID_0077(1)
  loc_11096CEC: var_68.DispID_0078(1500)
  loc_11096D34: var_68.DispID_0090("数量")
  loc_11096D78: var_68.DispID_0077(var_4C)
  loc_11096DBE: var_68.DispID_0078(var_4C)
  loc_11096E06: var_68.DispID_009C(var_4C)
  loc_11096E4E: var_68.DispID_0090(var_4C)
  loc_11096E94: var_68.DispID_0077(var_4C)
  loc_11096EDA: var_68.DispID_0078(var_4C)
  loc_11096F22: var_68.DispID_009C(var_4C)
  loc_11096F6A: var_68.DispID_0090(var_4C)
  loc_11096FB0: var_68.DispID_0077(var_4C)
  loc_11096FF6: var_68.DispID_0078(var_4C)
  loc_1109703E: var_68.DispID_009C(var_4C)
  loc_11097086: var_68.DispID_0090(var_4C)
  loc_110970CC: var_68.DispID_0077(var_4C)
  loc_11097112: var_68.DispID_0078(var_4C)
  loc_1109715A: var_68.DispID_009C(var_4C)
  loc_110971A2: var_68.DispID_0090(var_4C)
  loc_110971E8: var_68.DispID_0077(var_4C)
  loc_1109722E: var_68.DispID_0078(var_4C)
  loc_11097276: var_68.DispID_009C(var_4C)
  loc_110972BE: var_68.DispID_0090(var_4C)
  loc_11097304: var_68.DispID_0077(var_4C)
  loc_1109734A: var_68.DispID_0078(var_4C)
  loc_11097392: var_68.DispID_009C(var_4C)
  loc_110973DA: var_68.DispID_0090(var_4C)
  loc_11097420: var_68.DispID_0077(var_4C)
  loc_11097466: var_68.DispID_0078(var_4C)
  loc_110974B3: If 14 <= CLng(28)(-1) Then
  loc_110974F3:   var_68.DispID_00AC(var_4C)
  loc_1109750B:   var_14 = 1+var_14
  loc_1109750E:   GoTo loc_110974AC
  loc_11097510: End If
  loc_1109754F: var_68.DispID_00AC(var_4C)
  loc_11097569: GoTo loc_11097575
  loc_11097574: Exit Sub
  loc_11097575: ' Referenced from: 11097569
End Sub

Private Sub Proc_14_9_1109AD90
  Dim var_58 As Variant
  Dim var_130 As Label
  Dim var_5C As frmSaleBillDr.Label3
  loc_1109AE11: var_8004 = ecx
  loc_1109AE7E: If var_14 <= CLng(frmSaleBillDr.VFG.DispID_0007)(-1) Then
  loc_1109AE8F:   var_800C = frmSaleBillDr.Proc_14_10_1109BF60(vbNull)
  loc_1109AF18:   frmSaleBillDr.VFG.DispID_0082(22, var_44)
  loc_1109AFED:   If (frmSaleBillDr.VFG.DispID_0082(var_14, 22) = global_1100AE28) + 1 Then
  loc_1109B067:     frmSaleBillDr.VFG.DispID_0082(1, 285267764)
  loc_1109B182:     frmSaleBillDr.VFG.DispID_009E(var_14, 1, var_14, 1, 16711680)
  loc_1109B1A2:     Set var_58 = frmSaleBillDr.Label3
  loc_1109B1AF:     var_130 = var_58
  loc_1109B1F9:     var_58.Caption = "分析: 第(" & CStr(vbNull) & ")行信息----有效"
  loc_1109B24B:     frmSaleBillDr.Pic1.DispID_FFFFFDDA
  loc_1109B25E:   Else
  loc_1109B2D2:     frmSaleBillDr.VFG.DispID_0082(1, 285267820)
  loc_1109B3ED:     frmSaleBillDr.VFG.DispID_009E(var_14, 1, var_14, 1, 255)
  loc_1109B40D:     Set var_5C = frmSaleBillDr.Label3
  loc_1109B41A:     var_130 = var_5C
  loc_1109B4EF:     var_5C.Caption = "分析:   第(" & CStr(vbNull) & ")行信息----" & frmSaleBillDr.VFG.DispID_0082(var_14, 22)
  loc_1109B557:     frmSaleBillDr.Pic1.DispID_FFFFFDDA
  loc_1109B569:   End If
  loc_1109B579:   var_14 = 1+var_14
  loc_1109B57C:   GoTo loc_1109AE73
  loc_1109B581: End If
  loc_1109B5E2: If var_14 <= CLng(frmSaleBillDr.VFG.DispID_0007)(-1) Then
  loc_1109B67F:   var_34 = var_14
  loc_1109B688:   var_30 = var_14
  loc_1109B739:   If (frmSaleBillDr.VFG.DispID_0082(var_14, frmSaleBillDr.VFG) = frmSaleBillDr.VFG.DispID_0082(var_14, frmSaleBillDr.VFG)) + 1 Then
  loc_1109B7F2:     If (frmSaleBillDr.VFG.DispID_0082(var_14, 22) = global_1100AE28) Then
  loc_1109B7FB:     End If
  loc_1109B81C:     var_14 = var_14(1)
  loc_1109B81F:     var_30 = var_30(1)
  loc_1109B83E:     var_8058 = CLng(frmSaleBillDr.VFG.DispID_0007)
  loc_1109B859:     var_130 = (var_14 > 0)
  loc_1109B87A:     If var_130 = 0 Then GoTo loc_1109B68B
  loc_1109B880:   End If
  loc_1109B885:   If var_28 Then
  loc_1109B8AB:     var_1C = var_34
  loc_1109B8B0:     If var_34 <= (var_30 - 1) Then
  loc_1109B968:       If (frmSaleBillDr.VFG.DispID_0082(var_1C, 22) = global_1100AE28) + 1 Then
  loc_1109B9EA:         frmSaleBillDr.VFG.DispID_0082(1, 285267820)
  loc_1109BA78:         frmSaleBillDr.VFG.DispID_0082(22, "某分录有错误")
  loc_1109BB93:         frmSaleBillDr.VFG.DispID_009E(var_1C, 1, var_1C, 1, 255)
  loc_1109BBA5:       End If
  loc_1109BBB5:       GoTo loc_1109B8A5
  loc_1109BBBA:     End If
  loc_1109BBCB:     var_3C = var_3C(1)
  loc_1109BBD9:     Set var_5C = frmSaleBillDr.Label3
  loc_1109BBEC:     var_130 = var_5C
  loc_1109BCAB:     var_5C.Caption = "分析: 第[" & frmSaleBillDr.VFG.DispID_0082(var_34, 18) & "]销售发票错误"
  loc_1109BD0D:   Else
  loc_1109BD1E:     var_20 = var_20(1)
  loc_1109BD2C:     Set var_5C = frmSaleBillDr.Label3
  loc_1109BD3C:     var_130 = var_5C
  loc_1109BDB0:     Set var_58 = frmSaleBillDr.VFG
  loc_1109BE09:     var_5C.Caption = "分析:   第[" & var_58.DispID_0082(var_34, 0) & "]销售发票有效"
  loc_1109BE66:   End If
  loc_1109BE69:   var_58.DispID_FFFFFDDA
  loc_1109BE97:   var_14 = 1+var_14(-1)
  loc_1109BE9A:   GoTo loc_1109B5DC
  loc_1109BE9F: End If
  loc_1109BEA4: If var_3C > 0 Then
  loc_1109BEAB:   If var_20 > 0 Then
  loc_1109BEC2:   Else
  loc_1109BED7:   Else
  loc_1109BEDF:     var_18 = ecx
  loc_1109BEE7:     GoTo loc_1109BF22
  loc_1109BF21:     Exit Sub
  loc_1109BF22:   End If
  loc_1109BF22: End If
  loc_1109BF22: ' Referenced from: 1109BEE7
End Sub

Private  Proc_14_10_1109BF60(arg_C) '1109BF60
  Dim var_18 As ADODB.Recordset
  Dim var_38 As Variant
  Dim var_3C As frmSaleBillDr.VFG
  loc_1109BFDB: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1109BFE6: var_11C = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1109C013: var_F4 = ADODB.Recordset.State
  loc_1109C038: If var_F4 = 1 Then
  loc_1109C058:   var_800C = ADODB.Recordset.Close
  loc_1109C076: End If
  loc_1109C0B2: var_124 = var_18
  loc_1109C0CB: var_F8 = var_18
  loc_1109C102: var_8014 = ADODB.Recordset.Open(8, var_98, "select IsNULL(Max(iPeriod),0) +1 As iMonth From GL_Mend where bFlag_SA=1", var_90, 9)
  loc_1109C1D1: var_38 = ADODB.Recordset.Fields
  loc_1109C1FC: var_100 = var_38
  loc_1109C21F: ADODB.Recordset.8 = Forms
  loc_1109C2A3: If (Month(frmSaleBillDr.VFG.DispID_0082(arg_C, 2)) < var_7C) Then
  loc_1109C2AF: Else
  loc_1109C2D4:   var_F4 = ADODB.Recordset.State
  loc_1109C2F9:   If var_F4 = 1 Then
  loc_1109C319:     var_802C = ADODB.Recordset.Close
  loc_1109C337:   End If
  loc_1109C398:   var_4C = frmSaleBillDr.VFG.DispID_0082(arg_C, 3)
  loc_1109C3B1:   var_54 = var_4C
  loc_1109C448:   var_F8 = var_18
  loc_1109C47F:   var_8040 = ADODB.Recordset.Open(8, var_98, var_7C & Proc_0_10_11028DD0(8, "select * From customer where cCuscode=", var_3C), var_90, 9)
  loc_1109C4E8:   If ADODB.Recordset.EOF Then
  loc_1109C4F4:   Else
  loc_1109C541:     var_100 = ADODB.Recordset.Fields
  loc_1109C564:     ADODB.Recordset.8 = Forms
  loc_1109C631:     frmSaleBillDr.VFG.DispID_0082(18, var_4C)
  loc_1109C686:     var_F4 = ADODB.Recordset.State
  loc_1109C6AB:     If var_F4 = 1 Then
  loc_1109C6CB:       var_8058 = ADODB.Recordset.Close
  loc_1109C6E9:     End If
  loc_1109C74D:     var_4C = frmSaleBillDr.VFG.DispID_0082(arg_C, var_A4)
  loc_1109C766:     var_54 = var_4C
  loc_1109C7E3:     var_F8 = var_18
  loc_1109C828:     var_806C = ADODB.Recordset.Open(8, var_98, var_4C & Proc_0_10_11028DD0(8, "select * From department where bdepend=1 and cdepcode=", var_38), var_90, 9)
  loc_1109C891:     If ADODB.Recordset.EOF Then
  loc_1109C89D:     Else
  loc_1109C8C2:       var_F4 = ADODB.Recordset.State
  loc_1109C8E7:       If var_F4 = 1 Then
  loc_1109C907:         var_807C = ADODB.Recordset.Close
  loc_1109C925:       End If
  loc_1109C989:       var_4C = frmSaleBillDr.VFG.DispID_0082(arg_C, var_A4)
  loc_1109C9A2:       var_54 = var_4C
  loc_1109CA39:       var_F8 = var_18
  loc_1109CA70:       var_8090 = ADODB.Recordset.Open(8, var_98, .VTable_110F601C, var_90, 9)
  loc_1109CAD9:       If ADODB.Recordset.EOF Then
  loc_1109CAE5:       Else
  loc_1109CB32:         var_100 = ADODB.Recordset.Fields
  loc_1109CB55:         ADODB.Recordset.8 = Forms
  loc_1109CC22:         frmSaleBillDr.VFG.DispID_0082(19, var_4C)
  loc_1109CC9F:         var_100 = ADODB.Recordset.Fields
  loc_1109CCC2:         ADODB.Recordset.8 = Forms
  loc_1109CCF8:         var_80A4 = CBool(0)
  loc_1109CD1A:         If var_80A4 = 0 Then
  loc_1109CD21:         End If
  loc_1109CDDC:         If (frmSaleBillDr.VFG.DispID_0082(arg_C, var_A4) = "50") + 1 Then
  loc_1109CDEB:           var_80B0 = ("0" = global_1100B518)
  loc_1109CDF3:           If var_80B0 Then GoTo loc_1109CF58
  loc_1109CE81:           frmSaleBillDr.VFG.DispID_0082(var_A4, CStr(var_80B0))
  loc_1109CE9C:         End If
  loc_1109CF48:         If Not (IsNumeric(frmSaleBillDr.VFG.DispID_0082(arg_C, var_A4))) Then
  loc_1109CF53:           GoTo loc_1109D29F
  loc_1109CFDE:           var_80C4 = CDbl(Val(frmSaleBillDr.VFG.DispID_0082(arg_C, var_A4)))
  loc_1109D091:           var_80CC = CDbl(Val(frmSaleBillDr.VFG.DispID_0082(arg_C, 7)))
  loc_1109D0A9:           GoTo loc_1109D0AD
  loc_1109D0EF:           If (edi Or 0) = 0 Then GoTo loc_1109CE9C
  loc_1109D0FF:         Else
  loc_1109D187:           var_80D4 = CDbl(Val(frmSaleBillDr.VFG.DispID_0082(arg_C, var_A4)))
  loc_1109D236:           var_80DC = CDbl(Val(frmSaleBillDr.VFG.DispID_0082(arg_C, 11)))
  loc_1109D24E:           GoTo loc_1109D252
  loc_1109D294:           If (esi Or 0) Then
  loc_1109D29F:           End If
  loc_1109D29F:         End If
  loc_1109D29F:       End If
  loc_1109D29F:     End If
  loc_1109D29F:   End If
  loc_1109D2A2:   var_20 = "金额超范围"
  loc_1109D2CD:   var_F4 = ADODB.Recordset.State
  loc_1109D2F2:   If var_F4 = 1 Then
  loc_1109D320:     If ADODB.Recordset.Close < 0 Then
  loc_1109D326:       GoTo loc_1109D39E
  loc_1109D328:     End If
  loc_1109D34D:     var_F4 = ADODB.Recordset.State
  loc_1109D372:     If var_F4 = 1 Then
  loc_1109D39C:       If ADODB.Recordset.Close < 0 Then
  loc_1109D39E:         ' Referenced from: 1109D326
  loc_1109D3AA:         var_80F4 = CheckObj(var_18, global_1100ADFC, 128)
  loc_1109D3B0:       End If
  loc_1109D3B0:     End If
  loc_1109D3B0:   End If
  loc_1109D3B0: End If
  loc_1109D3B6: GoTo loc_1109D40E
  loc_1109D3BC: If var_4 Then
  loc_1109D3C7: End If
  loc_1109D40D: Exit Sub
  loc_1109D40E: ' Referenced from: 1109D3B6
End Sub

Private Sub Proc_14_11_1109D450
  Dim var_74 As Variant
  Dim var_8054 As Label
  Dim var_78 As Variant
  Dim var_184 As Label
  Dim var_3C As Me
  loc_1109D510: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1109D516: var_1B4 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_1109D539: Set var_74 = frmSaleBillDr.VFG
  loc_1109D580: If (CLng(var_74.DispID_0007) < 2) Then
  loc_1109D5AB:   var_800C = = Global.Screen
  loc_1109D5CA:   var_8010 = ecx
  loc_1109D5D2:   var_8010 = var_74.UnkVCall_0000007Ch
  loc_1109D67A:   MsgBox("没有可生成用友销售发票的数据。", 64, "提示信息", 10, 10)
  loc_1109D6A7:   Exit Sub
  loc_1109D6B7: Else
  loc_1109D6B8:   On Error GoTo 0
  loc_1109D6CC:   call edi(var_74, frmSaleBillDr.Label3, %ecx = %S_edx_S, "3缷M鋎?")
  loc_1109D6CE:   var_184 = edi(var_74, frmSaleBillDr.Label3, %ecx = %S_edx_S, "3缷M鋎?")
  loc_1109D6DC:   Label3.Caption = "正在进行数据分析，请稍等..."
  loc_1109D706:   var_C0 = True
  loc_1109D746:   call edi(var_74, frmSaleBillDr.Pic1, global_80010007, 0000000Bh, var_C4, True, var_BC)
  loc_1109D749:   edi(var_74, frmSaleBillDr.Pic1, global_80010007, 0000000Bh, var_C4, True, var_BC).DispID_0000 =
  loc_1109D76C:   call edi(var_74, frmSaleBillDr.Pic1, global_FFFFFDDA, %ecx = %S_edx_S)
  loc_1109D76F:   edi(var_74, frmSaleBillDr.Pic1, global_FFFFFDDA, var_74 = var_74).DispID_0000
  loc_1109D78B:   var_8014 = .Proc_14_9_1109AD90(var_17C)
  loc_1109D799:   If var_17C = 2 Then
  loc_1109D7E1:     call edi(var_74, frmSaleBillDr.Pic1, global_80010007, 0000000Bh, var_C4, frmSaleBillDr.Pic1, var_BC)
  loc_1109D7E4:     edi(var_74, frmSaleBillDr.Pic1, global_80010007, 0000000Bh, var_C4, frmSaleBillDr.Pic1, var_BC).DispID_0000 =
  loc_1109D87A:     MsgBox("数据源中没有合法的数据，无法生成用友销售发票，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1109D8CC:     var_801C = = Global.Screen
  loc_1109D8EB:     var_8020 = ecx
  loc_1109D8FA:     If var_74.UnkVCall_0000007Ch < 0 Then
  loc_1109D900:       GoTo loc_1109DD61
  loc_1109D905:     End If
  loc_1109D907:     If var_8020 = 1 Then
  loc_1109D94F:       call var_8024 = var_74(var_74, frmSaleBillDr.Pic1, global_80010007, 0000000Bh, var_C4, frmSaleBillDr.Pic1, var_BC)
  loc_1109D952:       var_8024.DispID_0000 =
  loc_1109D9E8:       MsgBox("数据源中含有非法的数据，无法生成用友销售发票，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1109DA3A:       var_802C = = Global.Screen
  loc_1109DA59:       var_8030 = ecx
  loc_1109DA68:       If var_74.UnkVCall_0000007Ch < 0 Then
  loc_1109DA6E:         GoTo loc_1109DD61
  loc_1109DA73:       End If
  loc_1109DA75:       If var_8030 = 3 Then
  loc_1109DABD:         call var_8034 = var_74(var_74, frmSaleBillDr.Pic1, global_80010007, 0000000Bh, var_C4, frmSaleBillDr.Pic1, var_BC)
  loc_1109DAC0:         var_8034.DispID_0000 =
  loc_1109DB56:         MsgBox("数据源中指定的凭证号无效或重号，无法生成用友销售发票，请认真查看出错信息。", 64, "提示信息", 10, 10)
  loc_1109DBA8:         var_803C = = Global.Screen
  loc_1109DBC7:         var_8040 = ecx
  loc_1109DBD6:         If var_74.UnkVCall_0000007Ch < 0 Then
  loc_1109DBDC:           GoTo loc_1109DD61
  loc_1109DBE1:         End If
  loc_1109DC23:         var_98 = "提示信息"
  loc_1109DC49:         var_88 = "数据源中的数据已全部通过检查，是否开始引入？"
  loc_1109DC6D:         MsgBox(var_88, 36, var_98, var_A8, var_B8)
  loc_1109DC7D:         var_184 = (MsgBox(var_88, 36, var_98, var_A8, var_B8) = 7)
  loc_1109DCB2:         If var_184 = 0 Then GoTo loc_1109DD89
  loc_1109DCFA:         call var_8044 = var_74(var_74, frmSaleBillDr.Pic1, global_80010007, 0000000Bh, var_C4, frmSaleBillDr.Pic1, var_BC)
  loc_1109DCFD:         var_8044.DispID_0000 =
  loc_1109DD31:         var_804C = = Global.Screen
  loc_1109DD50:         var_8050 = ecx
  loc_1109DD5F:         If var_74.UnkVCall_0000007Ch < 0 Then
  loc_1109DD6A:           var_8050 = CheckObj(var_74, global_1100C47C, 124)
  loc_1109DD70:         End If
  loc_1109DD70:       End If
  loc_1109DD70:     End If
  loc_1109DD70:   End If
  loc_1109DD79:   Exit Sub
  loc_1109DD84:   GoTo loc_110A2F9E
  loc_1109DD8A:   On Error GoTo 0
  loc_1109DD9E:   call var_8054 = var_74(var_74, frmSaleBillDr.Label3, var_74 = var_74, "3缷M鋎?")
  loc_1109DDA0:   var_184 = var_8054
  loc_1109DDAE:   Label3.Caption = "正在写数据，请稍等..."
  loc_1109DDE9:   call var_8058 = var_74(var_74, frmSaleBillDr.Pic1, global_FFFFFDDA, var_74 = var_74)
  loc_1109DDEC:   var_8058.DispID_0000
  loc_1109DE0C:   call var_805C = var_74(var_74, frmSaleBillDr.TDBText)
  loc_1109DE19:   var_805C.UnkVCall_00000040h
  loc_1109DE5F:   var_58 = Proc_0_11_11029000(9, var_805C, 2)
  loc_1109DEEE:   If var_20 <= CLng(frmSaleBillDr.VFG.DispID_0007)(-1) Then
  loc_1109DEF8:     var_1C0 = var_20
  loc_1109DF9B:     var_44 = frmSaleBillDr.VFG.DispID_0082(var_20, 0)
  loc_1109E098:     If (frmSaleBillDr.VFG.DispID_0082(var_1C0, 1) = global_1100D76C) Then
  loc_1109E0AC:       Set var_78 = frmSaleBillDr.Label3
  loc_1109E0B2:       var_184 = var_78
  loc_1109E197:       var_78.Caption = "正在处理：第[" & frmSaleBillDr.VFG.DispID_0082(var_1C0, 0) & "]的销售发票。"
  loc_1109E201:       frmSaleBillDr.Pic1.DispID_FFFFFDDA
  loc_1109E216:       var_38 = var_20
  loc_1109E21B:       On Error GoTo loc_110A22EB
  loc_1109E237:       var_8084 = UnkObj.UnkVCall_00000044h
  loc_1109E381:       If (var_44 = Trim(frmSaleBillDr.VFG.DispID_0082(var_20, var_E0))) Then
  loc_1109E38C:         If var_5C = 1 Then
  loc_1109E39A:           var_3C = "SELECT cNumber AS Maxnumber FROM VoucherHistory WITH (UPDLOCK) WHERE CardNumber='07' AND cContent IS NULL"
  loc_1109E3C9:           var_180 = ADODB.Recordset.State
  loc_1109E3F4:           If var_180 = 1 Then
  loc_1109E418:             var_8098 = ADODB.Recordset.Close
  loc_1109E43C:           End If
  loc_1109E4E1:           var_80A0 = ADODB.Recordset.Open(var_3C, var_C4, var_3C, var_BC, 9)
  loc_1109E554:           If ADODB.Recordset.EOF Then
  loc_1109E55E:             var_54 = "0000000001"
  loc_1109E56C:             var_3C = "INSERT INTO VoucherHistory(CardNumber,cNumber) values('07','1')"
  loc_1109E577:           Else
  loc_1109E5BE:             var_18C = ADODB.Recordset.Fields
  loc_1109E5C4:             var_C0 = "MaxNumber"
  loc_1109E605:             ADODB.Recordset.8 = Forms
  loc_1109E626:             var_194 = var_78
  loc_1109E6D4:             var_54 = Format(var_78.UnkVCall_00000034h + 1, var_E0)
  loc_1109E768:             var_18C = ADODB.Recordset.Fields
  loc_1109E76E:             var_C0 = "MaxNumber"
  loc_1109E7AF:             ADODB.Recordset.8 = Forms
  loc_1109E7D0:             var_194 = var_78
  loc_1109E865:             var_B8 = "UPDATE VoucherHistory SET cNumber='" & var_78.UnkVCall_00000034h + 1 & "' WHERE  CardNumber='07' AND cContent is NULL"
  loc_1109E877:             var_3C = var_B8
  loc_1109E8B0:           End If
  loc_1109EB14:           var_48, var_40, "00")
  loc_1109EBAB:           var_C0 = frmSaleBillDr.VFG.DispID_0082(var_20, 4)
  loc_1109EC2E:           var_80EC = Proc_0_10_11028DD0(&H4008, "BILLVOUCH" & Proc_0_10_11028DD0(&H4008, "53," & "'应收',", CStr("cAcc_Id".00000000h)) & global_1100AC40, CLng(frmSaleBillDr.VFG.DispID_0007)(-1))
  loc_1109ECE6:           var_8110 = "INSERT INTO SaleBillVouch(" & "ivtid," & "csource," & "cdepcode," & "ccuscode," & "bcashsale," & "sbvid," & "csbvcode,"
  loc_1109ECF2:           var_C0 = var_54
  loc_1109ED0D:           var_8114 = Proc_0_10_11028DD0(&H4008, frmSaleBillDr.VFG.DispID_0082(var_20, 3) & var_80EC & global_1100AC40 & "0," & CStr(var_48) & global_1100AC40, var_BC)
  loc_1109EE18:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_C0, 2)
  loc_1109EE56:           var_64 = 1 & Proc_0_10_11028DD0(8, var_78 & var_8114 & global_1100AC40 & "26,", 1) & global_1100AC40
  loc_1109F003:           var_90 = frmSaleBillDr.VFG.DispID_0082(1, 13)
  loc_1109F01E:           var_8164 = Proc_0_10_11028DD0(8, var_BC & Proc_0_10_11028DD0("cUserName".00000000h, var_C4, 1) & global_1100AC40 & "1," & "13,", var_78)
  loc_1109F0A5:           var_8178 = var_8110 & "cvouchtype," & "ddate," & "cmaker," & "iexchrate," & "itaxrate," & "cmemo," & "breturnflag," & "bfirst,"
  loc_1109F0D5:           var_C0 = var_58
  loc_1109F143:           var_8194 = var_CC & Proc_0_10_11028DD0(&H4008, var_D4 & var_8164 & global_1100AC40 & "0," & "0,", var_D0) & global_1100AC40 & "'人民币',"
  loc_1109F1FB:           var_90 = frmSaleBillDr.VFG.DispID_0082(1, &H12)
  loc_1109F292:           var_64 = 1 & Proc_0_10_11028DD0(8, var_8194, 0) & global_1100AC40 & "0,"
  loc_1109F364:           var_81D4 = var_8178 & "cbcode," & "cexch_name," & "ccusname," & "idisp," & "cchecker," & "bcredit," & "ioutgolden," & "iverifystate,"
  loc_1109F3BE:           var_81E8 = var_78 & Proc_0_10_11028DD0("cUserName".00000000h, -1, 1) & global_1100AC40 & "0," & "NULL," & "0," & "NULL," & "0,"
  loc_1109F40C:           var_3C = var_81D4 & "ireturncount," & "iswfcontrolled," & "cverifier" & ") VALUES (" & var_81E8 & "NULL)"
  loc_1109F44B:           var_88 = UnkObj.UnkVCall_00000040h
  loc_1109F50D:           var_820C = CLng(frmSaleBillDr.VFG.DispID_0007)
  loc_1109F522:           (var_40 - var_820C) = (var_40 - var_820C) + var_20
  loc_1109F52A:           (var_40 - var_820C)+var_20 = (var_40 - var_820C)+var_20 + 1
  loc_1109F612:           Set var_74 = frmSaleBillDr.VFG
  loc_1109F630:           var_90 = var_74.DispID_0082(var_20, 5)
  loc_1109F64B:           var_8224 = Proc_0_10_11028DD0(8, CStr(var_48) & global_1100AC40 & CStr((var_40 - var_820C)+var_20+1) & global_1100AC40, var_74)
  loc_1109F753:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 7)
  loc_1109F819:           var_8254 = "INSERT INTO SaleBillVouchs(" & "sbvid," & "autoid," & "cinvcode," & "iquantity," & "inum," & "iquotedprice," & "iunitprice,"
  loc_1109F8BE:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 8)
  loc_1109F8D9:           var_825C = Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8,  & var_8224 & global_1100AC40) & global_1100AC40 & "NULL," & "0,")
  loc_1109F9E1:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 10)
  loc_1109FB04:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 9)
  loc_1109FB3B:           var_828C =  & Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8,  & var_825C & global_1100AC40) & global_1100AC40) & global_1100AC40
  loc_1109FC27:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 12)
  loc_1109FD4A:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 11)
  loc_1109FE91:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 8)
  loc_1109FEAC:           var_82C8 = Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8, var_828C) & global_1100AC40) & global_1100AC40 & "0,")
  loc_1109FFB4:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 9)
  loc_110A0032:           var_82E8 = var_8254 & "itaxunitprice," & "imoney," & "itax," & "isum," & "idiscount," & "inatunitprice," & "inatmoney," & "inattax,"
  loc_110A00D7:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 12)
  loc_110A010E:           var_82F8 =  & Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8,  & var_82C8 & global_1100AC40) & global_1100AC40) & global_1100AC40
  loc_110A01FA:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 11)
  loc_110A0350:           var_8340 = var_82E8 & "inatsum," & "inatdiscount," & "isbvid," & "imoneysum," & "iexchsum," & "cclue," & "cincomesub," & "ctaxsub,"
  loc_110A0386:           var_834C =  & Proc_0_12_110291B0(8, var_82F8) & global_1100AC40 & "0," & "0," & "0," & "0," & "NULL," & "NULL," & "NULL," & "NULL,"
  loc_110A0515:           var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, &H13)
  loc_110A0593:           var_8394 = var_8340 & "ibatch," & "bsettleall," & "rdsid," & "itb," & "isosid," & "idlsid," & "kl,kl2," & "cinvname," & "itaxrate,"
  loc_110A05A5:           var_8398 =  & Proc_0_10_11028DD0(8, var_834C & "0," & "NULL," & "0," & "NULL," & "NULL," & "100,100,") & global_1100AC40 & "13,"
  loc_110A06B3:           var_83D4 = var_8394 & "foutquantity," & "foutnum," & "fsalecost," & "fsaleprice," & "iinvexchrate," & "ipbvsid," & "ccode," & "csocode,"
  loc_110A0755:           var_83F8 = var_8398 & "0," & "0," & "0," & "0," & "NULL," & "NULL," & "NULL," & "NULL," & "0," & "NULL," & "NULL," & "NULL,"
  loc_110A07AF:           var_840C = var_83D4 & "bgsp," & "ccontractid," & "ccontracttagcode," & "ccontractrowguid," & "cmassunit," & "bqaneedcheck," & "bqaurgency,"
  loc_110A08D8:           var_844C = var_840C & "cbaccounter," & "bcosting," & "cordercode," & "iorderrowno," & "irowno," & "idemandtype," & "cdemandcode,"
  loc_110A08EA:           var_8450 = var_83F8 & "0," & "0," & "0," & "NULL," & "0," & "NULL," & "NULL," & CStr(1) & global_1100AC40 & "NULL," & "NULL,"
  loc_110A09D4:           var_8484 = var_844C & "cdemandmemo," & "cdemandid," & "idemandseq," & "cbdlcode," & "iinvsncount," & "bneedsign," & "cgathingcode,"
  loc_110A0A9A:           var_84B0 = var_8450 & "NULL," & "NULL," & "NULL," & "NULL," & "0," & "0," & "NULL," & "0," & "NULL," & "0," & "0," & "NULL)"
  loc_110A0AC4:           var_3C = var_8484 & "ftaxpasum," & "fpasum," & "fnattaxpasum," & "fnatpasum," & "body_outid" & ") VALUES (" & var_84B0
  loc_110A0B0A:           If UnkObj.UnkVCall_00000040h < 0 Then
  loc_110A0B29:           Else
  loc_110A0BB2:             var_84CC = CLng(frmSaleBillDr.VFG.DispID_0007)
  loc_110A0BC7:             (var_40 - var_84CC) = (var_40 - var_84CC) + var_20
  loc_110A0BCF:             (var_40 - var_84CC)+var_20 = (var_40 - var_84CC)+var_20 + 1
  loc_110A0CD5:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 5)
  loc_110A0CF0:             var_84E4 = Proc_0_10_11028DD0(8, CStr(var_48) & global_1100AC40 & CStr((var_40 - var_84CC)+var_20+1) & global_1100AC40, global_1100BD8C)
  loc_110A0DDA:             Set var_74 = frmSaleBillDr.VFG
  loc_110A0DF8:             var_90 = var_74.DispID_0082(var_20, 7)
  loc_110A0EBE:             var_8514 = "INSERT INTO SaleBillVouchs(" & "sbvid," & "autoid," & "cinvcode," & "iquantity," & "inum," & "iquotedprice," & "iunitprice,"
  loc_110A0F63:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 8)
  loc_110A0F7E:             var_851C = Proc_0_12_110291B0(8, frmSaleBillDr.VFG & Proc_0_12_110291B0(8, 64 & var_84E4 & global_1100AC40, -1) & global_1100AC40 & "NULL," & "0,")
  loc_110A1086:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 10)
  loc_110A11A9:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 9)
  loc_110A11E0:             var_854C =  & Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8,  & var_851C & global_1100AC40) & global_1100AC40) & global_1100AC40
  loc_110A12CC:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 12)
  loc_110A13EF:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 11)
  loc_110A1536:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 8)
  loc_110A1551:             var_8588 = Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8, var_854C) & global_1100AC40) & global_1100AC40 & "0,")
  loc_110A1659:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 9)
  loc_110A16D7:             var_85A8 = var_8514 & "itaxunitprice," & "imoney," & "itax," & "isum," & "idiscount," & "inatunitprice," & "inatmoney," & "inattax,"
  loc_110A177C:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 12)
  loc_110A17B3:             var_85B8 =  & Proc_0_12_110291B0(8,  & Proc_0_12_110291B0(8,  & var_8588 & global_1100AC40) & global_1100AC40) & global_1100AC40
  loc_110A189F:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, 11)
  loc_110A19F5:             var_8600 = var_85A8 & "inatsum," & "inatdiscount," & "isbvid," & "imoneysum," & "iexchsum," & "cclue," & "cincomesub," & "ctaxsub,"
  loc_110A1A2B:             var_860C =  & Proc_0_12_110291B0(8, var_85B8) & global_1100AC40 & "0," & "0," & "0," & "0," & "NULL," & "NULL," & "NULL," & "NULL,"
  loc_110A1BBA:             var_90 = frmSaleBillDr.VFG.DispID_0082(var_20, &H13)
  loc_110A1C38:             var_8654 = var_8600 & "ibatch," & "bsettleall," & "rdsid," & "itb," & "isosid," & "idlsid," & "kl,kl2," & "cinvname," & "itaxrate,"
  loc_110A1C4A:             var_8658 =  & Proc_0_10_11028DD0(8, var_860C & "0," & "NULL," & "0," & "NULL," & "NULL," & "100,100,") & global_1100AC40 & "13,"
  loc_110A1D34:             var_868C = var_8654 & "foutquantity," & "foutnum," & "fsalecost," & "fsaleprice," & "iinvexchrate," & "ipbvsid," & "ccode,"
  loc_110A1DFA:             var_86B8 = var_8658 & "0," & "0," & "0," & "0," & "NULL," & "NULL," & "NULL," & "NULL," & "0," & "NULL," & "NULL," & "NULL,"
  loc_110A1E30:             var_86C4 = var_868C & "csocode," & "bgsp," & "ccontractid," & "ccontracttagcode," & "ccontractrowguid," & "cmassunit," & "bqaneedcheck,"
  loc_110A1F5B:             var_8704 = var_86C4 & "bqaurgency," & "cbaccounter," & "bcosting," & "cordercode," & "iorderrowno," & "irowno," & "idemandtype,"
  loc_110A1F91:             var_8710 = var_86B8 & "0," & "0," & "0," & "NULL," & "0," & "NULL," & "NULL," & CStr(vbNull) & global_1100AC40 & "NULL," & "NULL,"
  loc_110A2057:             var_873C = var_8704 & "cdemandcode," & "cdemandmemo," & "cdemandid," & "idemandseq," & "cbdlcode," & "iinvsncount," & "bneedsign,"
  loc_110A2141:             var_8770 = var_8710 & "NULL," & "NULL," & "NULL," & "NULL," & "0," & "0," & "NULL," & "0," & "NULL," & "0," & "0," & "NULL)"
  loc_110A2153:             var_8774 = var_873C & "cgathingcode," & "ftaxpasum," & "fpasum," & "fnattaxpasum," & "fnatpasum," & "body_outid" & ") VALUES ("
  loc_110A216B:             var_3C = var_8774 & var_8770
  loc_110A21B1:             If UnkObj.UnkVCall_00000040h >= 0 Then GoTo loc_110A21CE
  loc_110A21C7:           End If
  loc_110A21C8:           var_88 = CheckObj(global_1100BD8C, 64, -1)
  loc_110A21CE:         End If
  loc_110A2202:         var_5C = var_5C(1)
  loc_110A2208:         var_1C0 = var_20(1)
  loc_110A2238:         var_877C = CLng(frmSaleBillDr.VFG.DispID_0007)
  loc_110A2254:         var_184 = (var_1C0 > 0)
  loc_110A2278:         If var_184 = 0 Then GoTo loc_1109E262
  loc_110A227E:       End If
  loc_110A228D:       var_8780 = UnkObj.UnkVCall_00000048h
  loc_110A22BB:       var_17C = frmSaleBillDr.UpdateBTData
  loc_110A22E0:       On Error GoTo 0
  loc_110A22E6:       GoTo loc_110A2B00
  loc_110A22FA:       var_8784 = UnkObj.UnkVCall_0000004Ch
  loc_110A2442:       If (var_44 = Trim(frmSaleBillDr.VFG.DispID_0082(var_38, 0))) Then
  loc_110A244B:         var_C0 = var_38
  loc_110A2508:         frmSaleBillDr.VFG.DispID_0082(1, "-")
  loc_110A268A:         frmSaleBillDr.VFG.DispID_009E(var_38, 1, var_38, 1, &HFF)
  loc_110A269F:         var_C0 = var_38
  loc_110A26C0:         var_8790 = Err
  loc_110A26D1:         var_184 = Err
  loc_110A279F:         frmSaleBillDr.VFG.DispID_0082(&H16, "未引入:" & Err.Description)
  loc_110A2804:         var_8798 = CLng(frmSaleBillDr.VFG.DispID_0007)
  loc_110A281F:         var_184 = (var_38(1) > 0)
  loc_110A2843:         If var_184 = 0 Then GoTo loc_110A2324
  loc_110A2849:       End If
  loc_110A2895:       frmSaleBillDr.Pic1.DispID_80010007 = var_C0
  loc_110A28A6:       var_879C = Resume(0)
  loc_110A28B7:     Else
  loc_110A299B:       If (var_44 = frmSaleBillDr.VFG.DispID_0082(var_1C0, 0)) + 1 Then
  loc_110A29A7:         var_C0 = var_1C0
  loc_110A2A64:         frmSaleBillDr.VFG.DispID_0082(&H16, "网络共享冲突----未引入")
  loc_110A2A84:         var_20 = var_20(1)
  loc_110A2ABA:         var_87A8 = CLng(frmSaleBillDr.VFG.DispID_0007)
  loc_110A2AD6:         var_184 = (var_20(1) > 0)
  loc_110A2AFA:         If var_184 = 0 Then GoTo loc_110A28B7
  loc_110A2B00:       End If
  loc_110A2B00:     End If
  loc_110A2B20:     var_20 = var_19C+(var_20 - 1)
  loc_110A2B23:     GoTo loc_1109DEE7
  loc_110A2B28:   End If
  loc_110A2B73:   frmSaleBillDr.Pic1.DispID_80010007 = var_C0
  loc_110A2BAB:   var_D0 = "提示信息"
  loc_110A2BB5:   If var_2C Then
  loc_110A2C1B:     MsgBox("数据引入已完成，数据已生成用友销售发票。", 64, var_D0, var_A8, var_B8)
  loc_110A2C8A:     frmSaleBillDr.VFG.DispID_0007 = 1
  loc_110A2D1B:     frmSaleBillDr.sBar.DispID_6803001E(1100AE28h)
  loc_110A2DAC:     frmSaleBillDr.sBar.DispID_6803001E(1100AE28h)
  loc_110A2E3D:     Set var_74 = frmSaleBillDr.sBar
  loc_110A2E40:     var_74.DispID_6803001E(1100AE28h)
  loc_110A2E53:   Else
  loc_110A2EAE:     MsgBox("数据没有被引入，原因请查看最后一列中的说明。", 64, var_D0, var_A8, var_B8)
  loc_110A2EDB:   End If
  loc_110A2F01:   var_87B0 = = Global.Screen
  loc_110A2F20:   var_87B4 = ecx
  loc_110A2F28:   var_87B4 = var_74.UnkVCall_0000007Ch
  loc_110A2F45:   Exit Sub
  loc_110A2F50:   GoTo loc_110A2F9E
  loc_110A2F9D:   Exit Sub
  loc_110A2F9E: End If
  loc_110A2F9E: ' Referenced from: 1109DD84
  loc_110A2F9E: ' Referenced from: 110A2F50
End Sub
