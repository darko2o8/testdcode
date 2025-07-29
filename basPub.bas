'VA: 1100ADA4
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private  Proc_0_0_110254E0(arg_10, arg_14, arg_18) '110254E0
  Dim var_14 As Variant
  Dim var_30 As 0
  Dim var_84 As Me
  loc_11025531: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11025545: var_1C = arg_10
  loc_1102554D: var_20 = arg_14
  loc_11025555: var_18 = arg_18
  loc_11025577: ADODB.Command.CommandType = CInt(4)
  loc_110255CE: ADODB.Command.ActiveConnection = CInt(9)
  loc_1102561D: ADODB.Command.CommandText = "sp_GetID"
  loc_11025666: var_84 = ADODB.Command.Parameters
  loc_110256B8: var_8018 = ADODB.Command.CreateParameter("@RemoteID", 200, 1, 2, 8)
  loc_110256F9: var_30.var_30 = Forms
  loc_11025759: var_84 = ADODB.Command.Parameters
  loc_110257AD: var_8024 = ADODB.Command.CreateParameter("@cAcc_ID", 200, 1, 3, 8)
  loc_110257EE: var_30.var_30 = Forms
  loc_1102584E: var_84 = ADODB.Command.Parameters
  loc_110258A2: var_8030 = ADODB.Command.CreateParameter("@cVouchType", 200, 1, 50, 8)
  loc_110258E3: var_30.var_30 = Forms
  loc_11025943: var_84 = ADODB.Command.Parameters
  loc_11025994: var_803C = ADODB.Command.CreateParameter("@iAmount", 3, 1, 10, 3)
  loc_110259D5: var_30.var_30 = Forms
  loc_11025A35: var_84 = ADODB.Command.Parameters
  loc_11025A88: var_8048 = ADODB.Command.CreateParameter("@iFatherId", 3, 2, 10, 10)
  loc_11025AC9: var_30.var_30 = Forms
  loc_11025B29: var_84 = ADODB.Command.Parameters
  loc_11025B7C: var_8054 = ADODB.Command.CreateParameter("@iChildId", 3, 2, 10, 10)
  loc_11025BBD: var_30.var_30 = Forms
  loc_11025C2D: var_805C = ADODB.Command.Execute(10, 10, -1)
  loc_11025C7F: var_28 = ADODB.Command.Parameters
  loc_11025CED: ADODB.Command.var_40 = Forms
  loc_11025D05: var_8064 = CLng(10)
  loc_11025D4E: var_28 = ADODB.Command.Parameters
  loc_11025DB2: ADODB.Command.var_40 = Forms
  loc_11025DCE: var_806C = CLng(10)
  loc_11025E07: Set var_14 = "懨顧物?斪粭椃浺"()
  loc_11025E12: GoTo loc_11025E3C
  loc_11025E3B: Exit Sub
  loc_11025E3C: ' Referenced from: 11025E12
End Sub

Private  Proc_0_1_11025E80(arg_C) '11025E80
  loc_11025EE6: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11025F24: var_8C = arg_C
  loc_11025F84: If CBool("TaskExec".00000002h) Then
  loc_11025F97: Else
  loc_11025FB5:   var_E8 = "LogState".0
  loc_11025FDA:   var_8008 = (var_E8 = 20)
  loc_11025FE3:   If var_8008 = 0 Then
  loc_11026011:     If (var_E8 = 22) Then
  loc_1102601F:     Else
  loc_11026025:       If arg_C Then
  loc_11026091:         If (Trim("ShareString".0) = 1100AE28h) Then
  loc_11026163:           MsgBox(1100AFB8h & "FuncName".00000001h & "]功能暂时不能执行！  ", 64, "提示信息", 10, 10)
  loc_11026185:         Else
  loc_110261FE:           MsgBox(0 & "ShareString" & Space(3), 64, "提示信息", 10, 10)
  loc_1102621B:         End If
  loc_11026227:       End If
  loc_11026227:     End If
  loc_11026227:   End If
  loc_11026234:   "ClearError".0(0, , , fs:[00000000h], )
  loc_11026242:   GoTo loc_11026268
  loc_11026267:   Exit Sub
  loc_11026268: End If
  loc_11026268: ' Referenced from: 11026242
End Sub

Private Sub Proc_0_2_11026290
  loc_110262DA: If IsNull(Me) Then
  loc_110262E4:   var_18 = " is null "
  loc_110262F4: Else
  loc_1102631B:   If (Me = 1100AE28h) Then
  loc_11026325:     var_18 = "=''"
  loc_11026335:   Else
  loc_11026384:     var_18 = "= '" & Trim(Me) & "'"
  loc_110263A6:     GoTo loc_110263CF
  loc_110263AC:     If var_4 Then
  loc_110263B7:     End If
  loc_110263CE:     Exit Sub
  loc_110263CF:   End If
  loc_110263CF: End If
  loc_110263CF: ' Referenced from: 110263A6
End Sub

Private  Proc_0_3_110263F0(arg_C, arg_10, arg_14, arg_18) '110263F0
  Dim var_24 As ADODB.Recordset
  Dim var_60 As Me
  loc_1102648C: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11026492: var_F8 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110264BC: If Me Then
  loc_110264CB:   var_2C = "SELECT min(inid), ccode" & " From GL_accvouch"
  loc_110264D5: Else
  loc_110264E2:   var_2C = var_2C & " Where md<>0 and iyear="
  loc_110264EA: End If
  loc_11026564: var_802C = var_2C & " Where mc<>0 and iyear=" & CStr(arg_C) & " and iperiod=" & CStr(arg_10) & " and csign='" & arg_14 & "' and iNo_id="
  loc_11026685: var_8044 = ADODB.Recordset.Open(var_802C & CStr(arg_18) & " Group By ccode" & " ORDER BY min(inid)", var_A0, var_802C & CStr(arg_18) & " Group By ccode" & " ORDER BY min(inid)", var_98, 9)
  loc_1102673F: If ((ADODB.Recordset.BOF = 0) Or (ADODB.Recordset.EOF = 0)) Then
  loc_1102676A:   var_C8 = ADODB.Recordset.EOF
  loc_11026786:   If var_C8 = 0 Then
  loc_110267C1:     var_60 = ADODB.Recordset.Fields
  loc_110267DF:     var_9C = "cCode"
  loc_11026813:     ADODB.Recordset.8 = Forms
  loc_110268C9:     If (Proc_0_13_110293B0(var_28 & var_74, var_64, var_74) >= 50) = 0 Then
  loc_11026904:       var_60 = ADODB.Recordset.Fields
  loc_11026922:       var_9C = "cCode"
  loc_11026956:       ADODB.Recordset.8 = Forms
  loc_110269E9:       var_28 = var_28 & var_74 & var_BC
  loc_11026A40:       If ADODB.Recordset.MoveNext >= 0 Then GoTo loc_11026745
  loc_11026A52:       var_806C = CheckObj(var_24, global_1100ADFC, 144)
  loc_11026A59:     End If
  loc_11026A59:   End If
  loc_11026A65:   If Len(var_28) > 0 Then
  loc_11026AB1:     var_28 = Left(var_28, (Proc_0_13_110293B0(var_28, var_64, var_74) - 1))
  loc_11026ABC:   End If
  loc_11026ABC: End If
  loc_11026ADA: var_8080 = ADODB.Recordset.Close
  loc_11026B06: Set var_24 = ADODB.Recordset()
  loc_11026B0E: On Error GoTo 0
  loc_11026B19: GoTo loc_11026B97
  loc_11026B1F: If var_8 Then
  loc_11026B2A: End If
  loc_11026B96: Exit Sub
  loc_11026B97: ' Referenced from: 11026B19
End Sub

Private Sub Proc_0_4_11026BD0
  loc_11026C12: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11026C41: var_8008 =  & Proc_0_2_11026290(&H4008, "select * from dsign where csign", )
  loc_11026C54: var_24 = var_8008
  loc_11026CA5: Set var_18 = "DataMdb".0.00000001h(, , fs:[00000000h], , , var_20, var_8008, var_28)
  loc_11026CFB: If CBool(Not(var_18.EOF)) Then
  loc_11026D41:   var_14 = CByte(var_18.Fields("isignseq"))
  loc_11026D51: Else
  loc_11026D64:   GoTo loc_11026D87
  loc_11026D86:   Exit Sub
  loc_11026D87: End If
  loc_11026D87: ' Referenced from: 11026D64
End Sub

Private  Proc_0_5_11026DB0(arg_C, arg_10) '11026DB0
  Dim var_190 As ADODB.Recordset
  loc_11026E70: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11026E79: var_198 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11026EA2: If IsNumeric(arg_10) Then
  loc_11026EB1:   var_8008 = CLng(Val(arg_10))
  loc_11026EB9:   var_24 = var_8008
  loc_11026EBC:   If var_8008 > 0 Then
  loc_11026EC7:     If var_8008 <= 9999 Then
  loc_11026EEC:       Set var_190 = var_20
  loc_11026F68:       var_802C = "select * from GL_accvouch where iperiod >=" & CStr(Me) & " and isignseq>=" & CStr(0) & " and ino_id>=" & CStr(var_24)
  loc_11026F6F:       var_18 = var_802C
  loc_11026FB8:       var_F0 = var_18
  loc_11027016:       var_8030 = ADODB.Recordset.Open(8, var_F4, var_18, var_EC, 9)
  loc_11027048:       var_8034 = Proc_0_4_11026BD0(arg_C, var_104, .VTable_110F601C 'Ignore this)
  loc_11027087:       If ADODB.Recordset.EOF Then
  loc_1102708C:         var_14 = var_24
  loc_11027094:       Else
  loc_110270D7:         var_F0 = "iPeriod"
  loc_110270E8:         var_154 = ADODB.Recordset.Fields
  loc_110270FD:         ADODB.Recordset.8 = Forms
  loc_1102714C:         var_100 = Me
  loc_1102718C:         var_168 = ADODB.Recordset.Fields
  loc_110271A9:         ADODB.Recordset.8 = Forms
  loc_11027239:         var_17C = ADODB.Recordset.Fields
  loc_11027256:         ADODB.Recordset.8 = Forms
  loc_1102735E:         If CBool(Not((var_68 = Me) And (var_88 = var_1C) And (var_B8 = var_24))) Then
  loc_11027363:           var_14 = var_24
  loc_11027366:         End If
  loc_11027366:       End If
  loc_1102736F:       var_8050 = ADODB.Recordset.Close
  loc_1102739E:     End If
  loc_1102739E:   End If
  loc_1102739E: End If
  loc_110273A4: GoTo loc_11027432
  loc_11027431: Exit Sub
  loc_11027432: ' Referenced from: 110273A4
End Sub

Private  Proc_0_6_11027470(arg_C, arg_10) '11027470
  Dim var_14 As ADODB.Recordset
  Dim var_1C As Me
  loc_110274D3: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_110274D9: var_98 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11027501: var_7C = ADODB.Recordset.State
  loc_11027523: If var_7C = 1 Then
  loc_11027543:   var_800C = ADODB.Recordset.Close
  loc_11027561: End If
  loc_1102756C: var_5C = arg_10
  loc_110275DE: var_8028 = Proc_0_10_11028DD0(&H4008, "SELECT MAX(ino_id) AS MaxID FROM GL_AccVouch WHERE iYear=" & CStr(Me) & " AND iperiod=" & CStr(arg_C) & " AND csign=", )
  loc_1102765A: var_80 = var_14
  loc_1102768F: var_8034 = ADODB.Recordset.Open( & var_8028, var_60,  & var_8028, var_58, 9)
  loc_110276EF: If ADODB.Recordset.EOF Then
  loc_110276FD: Else
  loc_1102775A:   var_88 = ADODB.Recordset.Fields
  loc_1102776F:   ADODB.Recordset.8 = Forms
  loc_11027790:   var_44 = 0
  loc_11027797:   var_4C = var_44
  loc_110277D1:   var_18 = CLng((Proc_0_12_110291B0(9, var_60, "MaxID") + 1))
  loc_110277EC: End If
  loc_1102780E: var_7C = ADODB.Recordset.State
  loc_11027830: If var_7C = 1 Then
  loc_11027850:   var_8050 = ADODB.Recordset.Close
  loc_1102786E: End If
  loc_11027879: var_5C = arg_10
  loc_110278B1: var_8060 = Proc_0_10_11028DD0(&H4008, "SELECT MAX(ino_id) AS MaxID FROM Vouchnum WHERE iperiod=" & CStr(arg_C) & " AND csign=", var_58)
  loc_11027922: var_80 = var_14
  loc_11027957: var_806C = ADODB.Recordset.Open(var_44 & var_8060, var_60, var_44 & var_8060, var_58, 9)
  loc_1102799A: var_78 = ADODB.Recordset.EOF
  loc_110279B7: If var_78 = 0 Then
  loc_11027A1A:   var_88 = ADODB.Recordset.Fields
  loc_11027A2F:   ADODB.Recordset.8 = Forms
  loc_11027A55:   var_4C = var_44
  loc_11027A7B:   var_B0 = var_18
  loc_11027A9A:   var_807C = CDbl((Proc_0_12_110291B0(9, var_60, "MaxID") + 1))
  loc_11027AB2:   GoTo loc_11027AB6
  loc_11027AE2:   If eax Then
  loc_11027B45:     var_88 = ADODB.Recordset.Fields
  loc_11027B5A:     ADODB.Recordset.8 = Forms
  loc_11027B7B:     var_44 = 0
  loc_11027B82:     var_4C = var_44
  loc_11027BBC:     var_18 = CLng((Proc_0_12_110291B0(9, var_60, "MaxID") + 1))
  loc_11027BD7:   End If
  loc_11027BD7: End If
  loc_11027BEE: var_5C = arg_10
  loc_11027C34: var_809C = var_44 & Proc_0_10_11028DD0(&H4008, "INSERT INTO VouchNum (iperiod,csign,ino_id) VALUES (" & CStr(arg_C) & global_1100AC40, var_58)
  loc_11027C75: var_1C = var_809C & global_1100AC40 & CStr(var_18) & global_1100BD88
  loc_11027D0C: GoTo loc_11027D53
  loc_11027D52: Exit Sub
  loc_11027D53: ' Referenced from: 11027D0C
End Sub

Private  Proc_0_7_11027D90(arg_C, arg_10) '11027D90
  loc_11027DE1: call __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11027DEA: var_64 = __vbaAptOffset(global_1100A430, 0, 0, 0)
  loc_11027DFB: var_50 = arg_C
  loc_11027E77: var_8020 =  & Proc_0_10_11028DD0(&H4008, "DELETE FROM Vouchnum WHERE iperiod=" & CStr(Me) & " AND csign=", ) & " AND ino_id=" & CStr(arg_10)
  loc_11027E7E: var_14 = var_8020
  loc_11027ED8: var_48 = UnkObj.UnkVCall_00000040h
  loc_11027F0D: GoTo loc_11027F49
  loc_11027F48: Exit Sub
  loc_11027F49: ' Referenced from: 11027F0D
End Sub

Private  Proc_0_8_11027F70(arg_C) '11027F70
  loc_11027FF6: var_8004 = IsNull(arg_C)
  loc_11027FFC: var_A4 = var_8004
  loc_1102805E: var_B0 = CBool(var_8004 Or (arg_C = 1100AE28h))
  loc_11028071: If var_B0 = 0 Then
  loc_110280EC:   If (InStr(1, arg_C, "-", 0) > "") Then
  loc_110280F8:   Else
  loc_1102815F:     If (InStr(1, arg_C, ".", 0) > "") Then
  loc_11028168:     Else
  loc_110281BC:       var_B0 = (InStr(1, arg_C, "/", 0) > "")
  loc_110281CF:       If var_B0 = 0 Then GoTo loc_110281DF
  loc_110281D6:     End If
  loc_110281D6:   End If
  loc_110281D9:   var_18 = "/"
  loc_110281F0:   If (var_18 = global_1100AE28) Then
  loc_11028229:     var_8020 = CInt(InStr(1, arg_C, var_18, 0))
  loc_110282C8:     var_28 = Mid(arg_C, var_8020(1), Len(arg_C))
  loc_110282F1:     var_8030 = InStr(1, var_28, var_18, 0)
  loc_110282FC:     If var_8030 > 0 Then
  loc_1102842B:       var_8050 = IsDate(Left(arg_C, (var_8020 - 1)) & "-" & Left(Left(arg_C, (var_8020 - 1)) & "-" & Left(var_28, (var_8030 - 1)) & "-" & Mid(var_28, var_8030(1), Len(var_28)), (var_8030 - 1)) & "-" & Mid(Left(arg_C, (var_8020 - 1)) & "-" & Left(var_28, (var_8030 - 1)) & "-" & Mid(var_28, var_8030(1), Len(var_28)), var_8030(1), Len(Left(arg_C, (var_8020 - 1)) & "-" & Left(var_28, (var_8030 - 1)) & "-" & Mid(var_28, var_8030(1), Len(var_28)))))
  loc_11028434:       If var_8050 Then
  loc_11028461:         var_3C = DateSerial(CInt(CInt(CInt(0))), 0, 0)
  loc_11028467:       End If
  loc_11028467:     End If
  loc_11028467:   End If
  loc_11028467: End If
  loc_1102846C: GoTo loc_110284A9
  loc_11028472: If var_4 Then
  loc_1102847D: End If
  loc_110284A8: Exit Sub
  loc_110284A9: ' Referenced from: 1102846C
End Sub

Private Sub Proc_0_9_11028500
  loc_11028625: var_8004 = IsNull(Me)
  loc_1102862B: var_198 = var_8004
  loc_1102867E: var_80 = var_8004 Or (Me = 1100AE28h)
  loc_1102868B: var_294 = CBool(var_80)
  loc_110286A4: If var_294 = 0 Then
  loc_11028853:   var_110 = (InStr(1, Me, ".", 0) > "") And (InStr(1, Me, "-", 0) > "") Or (InStr(1, Me, ".", 0) > "") And (InStr(1, Me, "/", 0) > "")
  loc_110288EC:   var_294 = CBool(var_110 Or (InStr(1, Me, "/", 0) > "") And (InStr(1, Me, "-", 0) > ""))
  loc_1102892B:   If var_294 = 0 Then
  loc_11028931:     var_178 = "-"
  loc_11028981:     If CBool(InStr(1, Me, var_178, 0)) Then
  loc_1102898D:     Else
  loc_110289DD:       If CBool(InStr(1, Me, var_178, 0)) Then
  loc_110289E6:       Else
  loc_11028A20:         var_294 = CBool(InStr(1, Me, var_178, 0))
  loc_11028A36:         If var_294 = 0 Then GoTo loc_11028CE2
  loc_11028A41:       End If
  loc_11028A41:     End If
  loc_11028A44:     var_38 = "/"
  loc_11028A7D:     var_8044 = CInt(InStr(1, Me, var_38, 0))
  loc_11028A94:     If var_8044 > 0 Then
  loc_11028AA4:       var_68 = (var_8044 - 1)
  loc_11028ABA:       var_80 = Mid(Me, 1, (var_8044 - 1))
  loc_11028B12:       var_80 = Mid(Me, var_8044(1), 10)
  loc_11028B27:       var_50 = var_80
  loc_11028B4F:       var_8054 = InStr(1, var_50, var_38, 0)
  loc_11028B5A:       If var_8054 > 0 Then
  loc_11028B6D:         var_68 = (var_8054 - 1)
  loc_11028BA2:         var_80 = Mid(var_50, 1, (var_8054 - 1))
  loc_11028C06:         var_80 = Mid(var_50, var_8054(1), 10)
  loc_11028C88:         On Error GoTo loc_11028CDB
  loc_11028CA2:         var_34 = DateValue(var_80 & "-" & var_80 & "-" & var_80)
  loc_11028CB5:         If IsDate(var_34) Then
  loc_11028CC8:           var_8074 = var_34
  loc_11028CCA:           Exit Sub
  loc_11028CD6:           GoTo loc_11028D84
  loc_11028CDB:           ' Referenced from: 11028C88
  loc_11028CE2:         End If
  loc_11028CE2:       End If
  loc_11028CE2:     End If
  loc_11028CE2:   End If
  loc_11028CE2: End If
  loc_11028CE2: Exit Sub
  loc_11028CEE: GoTo loc_11028D84
  loc_11028D83: Exit Sub
  loc_11028D84: ' Referenced from: 11028CD6
  loc_11028D84: ' Referenced from: 11028CEE
End Sub

Private Sub Proc_0_10_11028DD0
  loc_11028E54: var_8004 = IsNull(Me)
  loc_11028ED1: var_C0 = var_8004
  loc_11028F5B: var_18 = IIf((Trim(Me) = 1100AE28h) Or var_8004, "NULL", "'" & Trim(Me) & "'")
  loc_11028F9C: GoTo loc_11028FE6
  loc_11028FA2: If var_4 Then
  loc_11028FAD: End If
  loc_11028FE5: Exit Sub
  loc_11028FE6: ' Referenced from: 11028F9C
End Sub

Private Sub Proc_0_11_11029000
  loc_1102906C: var_8004 = IsNull(Me)
  loc_110290BE: var_A0 = var_8004
  loc_1102911D: var_18 = IIf((Trim(Me) = 1100AE28h) Or var_8004, 1100AE28h, Trim(Me))
  loc_11029151: GoTo loc_1102918D
  loc_11029157: If var_4 Then
  loc_11029162: End If
  loc_1102918C: Exit Sub
  loc_1102918D: ' Referenced from: 11029151
End Sub

Private Sub Proc_0_12_110291B0
  loc_11029228: var_8004 = IsNull(Me)
  loc_1102923C: var_E0 = IsEmpty(Me)
  loc_11029285: var_C0 = var_E0
  loc_110292AC: var_B0 = var_8004
  loc_11029314: var_18 = IIf((Trim(Me) = 1100AE28h) Or var_8004 Or var_E0, 1100C008h, Trim(Me))
  loc_11029352: GoTo loc_11029395
  loc_11029358: If var_4 Then
  loc_11029363: End If
  loc_11029394: Exit Sub
  loc_11029395: ' Referenced from: 11029352
End Sub

Private Sub Proc_0_13_110293B0
  loc_1102940A: If 1 <= Len(Me) Then
  loc_1102944D:   var_800C = Asc(CStr(Mid(var_48, 1, 1)))
  loc_11029474:   If var_800C >= 0 Then
  loc_1102947B:     If var_800C <= 255 Then
  loc_11029487:       var_1C = var_1C(1)
  loc_1102948C:     Else
  loc_1102948C:     End If
  loc_11029496:     var_1C = var_1C(1)
  loc_11029499:   End If
  loc_110294A5:   GoTo loc_11029406
  loc_110294AA: End If
  loc_110294AF: GoTo loc_110294CE
  loc_110294CD: Exit Sub
  loc_110294CE: ' Referenced from: 110294AF
End Sub
