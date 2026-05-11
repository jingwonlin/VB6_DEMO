Attribute VB_Name = "Variable"
'程式版本: VB6-2007.07.11.01
Option Explicit
Option Base 0
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public rcn As New ADODB.Connection, rcn1 As New ADODB.Connection, rcn2 As New ADODB.Connection, rcn_rpt As New ADODB.Connection, rcn_bar As New ADODB.Connection, rcn_ora As New ADODB.Connection
Public rcn_str As String, rcn_str_local As String, rcn_str_rpt As String, mis_str As String
Public brs As New ADODB.Recordset, lrs As New ADODB.Recordset
Public first_prog As Form, prog_s210 As Form, prog As String * 15, mode As Integer, err1_old As Integer
Public tty As String * 60, stzone As String, lib_buf As String * 200
Public msg_prompt As String, msg_title As String, Response As Integer, err2_cnt As Integer
Public msg_button As Integer, msg_response As Integer, passwd_response As Integer, order_response As Integer
Public first_prog_large As Boolean, second_prog_large As Boolean, auto1_old As String
Public computername As String, loginuser As String * 10, loginname As String * 8, righ As String * 96, super As String * 1, dept As String * 10, loginpass As String * 5
Public lono_buf As String * 8, stno_buf As String * 5, sitm_buf As String * 19, sitm1_buf As String * 19, srid_buf As String * 1, prog_buf As String * 4
Public awno As String * 1, bank(6) As Integer, bay(6) As Integer, LEVEL(6) As Integer, DB_TYPE As String
Public AW_NUM As Integer, STNO_NUM As Integer, PLC1_NUM As Integer, CRAN_NUM As Integer, PALLET_NUM(6) As Integer
Public cran1_index(10) As Integer, port_id(10) As Integer, cran_move_flag As Boolean, cran_move_flag1 As String, start_port As Integer
Public err1_place As String, refresh1_screen As Integer, project_name As String * 4, plant_name As String * 7, get_s360_ok As Boolean
Public old_a40_load As String * 1, backup_dir As String, cran_buf As String * 20
Public palt_plc_write_asrs As Integer, palt_plc_write_asrs_old As Integer, cmd_plc_write_asrs_old As String * 8
Public strStno As String   '' PLC Slave Station No.
Public straddr As String   '' Read Wxxxx Address
Public strPoint As String  '' Read Lenght
Public strData As String, plc_buf As String * 2000, title As String
Public ora_flag As Boolean, mveq_s210 As String * 10, lamp(10) As Integer, change_car As Integer
Public dbuf(9) As String, cran2_cnt(5) As Integer
Type lo_display
  status As Integer
End Type

Global Const EQ_NULL  As String * 10 = "FINISH    "
Global Const EQ_CRAN1 As String * 10 = "#1天車    "  'S210-->Label20(1)
Global Const EQ_CRAN2 As String * 10 = "#2天車    "
Global Const EQ_CRAN3 As String * 10 = "#3天車    "
Global Const EQ_CRAN4 As String * 10 = "#4天車    "
Global Const EQ_CRAN5 As String * 10 = "#5天車    "
Global Const EQ_CRAN6 As String * 10 = "#6天車    "
Global Const EQ_CRAN7 As String * 10 = "#7天車    "
Global Const EQ_CRAN8 As String * 10 = "#8天車    "
Global Const EQ_CRAN9 As String * 10 = "#9天車    "
Global Const EQ_CRAN10 As String * 10 = "#0天車    " 'S210-->Label20(0)
Global Const EQ_STV1  As String * 10 = "STV1      "  'S210-->Label20(10)
Global Const EQ_STV2  As String * 10 = "STV2      "
Global Const EQ_STV3  As String * 10 = "STV3      "
Global Const EQ_STV4  As String * 10 = "STV4      "
Global Const EQ_STV5  As String * 10 = "STV5      "
Global Const EQ_STV6  As String * 10 = "STV6      "
Global Const EQ_STV7  As String * 10 = "STV7      "
Global Const EQ_STV8  As String * 10 = "STV8      "
Global Const EQ_STV9  As String * 10 = "STV9      "
Global Const EQ_STV10 As String * 10 = "STV10     "
Global Const EQ_STV11 As String * 10 = "STV11     "
Global Const EQ_STV12 As String * 10 = "STV12     "
Global Const EQ_STV13 As String * 10 = "STV13     "
Global Const EQ_STV14 As String * 10 = "STV14     "
Global Const EQ_STV15 As String * 10 = "STV15     "
Global Const EQ_STV16 As String * 10 = "STV16     "
Global Const EQ_STV17 As String * 10 = "STV17     "
Global Const EQ_STV18 As String * 10 = "STV18     "
Global Const EQ_STV19 As String * 10 = "STV19     "
Global Const EQ_STV20 As String * 10 = "STV20     "
Global Const EQ_STV21 As String * 10 = "STV21     "
Global Const EQ_STV22 As String * 10 = "STV22     "
Global Const EQ_STV23 As String * 10 = "STV23     "
Global Const EQ_STV24 As String * 10 = "STV24     "
Global Const EQ_STV25 As String * 10 = "STV25     "
Global Const EQ_STV26 As String * 10 = "STV26     "  'S210-->Label20(35)
Global Const EQ_STV27 As String * 10 = "STV27     "
Global Const EQ_STV28 As String * 10 = "STV28     "
Global Const EQ_STV29 As String * 10 = "STV29     "
Global Const EQ_STV30 As String * 10 = "STV30     "
Global Const EQ_STV31 As String * 10 = "STV31     "
Global Const EQ_STV32 As String * 10 = "STV32     "
Global Const EQ_STV33 As String * 10 = "STV33     "
Global Const EQ_STV34 As String * 10 = "STV34     "
Global Const EQ_STV35 As String * 10 = "STV35     "
Global Const EQ_STV36 As String * 10 = "STV36     "
Global Const EQ_MONO1 As String * 10 = "MONO1     "  'S210-->Label20(51)
Global Const EQ_MONO2 As String * 10 = "MONO2     "
Global Const EQ_MONO3 As String * 10 = "MONO3     "
Global Const EQ_MONO4 As String * 10 = "MONO4     "  'S210-->Label20(54)
Global Const EQ_LIFTER1 As String * 10 = "LIFTER1   " 'S210-->Label20(61)
Global Const EQ_LIFTER2 As String * 10 = "LIFTER2   "
Global Const EQ_LIFTER3 As String * 10 = "LIFTER3   "
Global Const EQ_LIFTER4 As String * 10 = "LIFTER4   "
Global Const EQ_LIFTER5 As String * 10 = "LIFTER5   "
Global Const EQ_LIFTER6 As String * 10 = "LIFTER6   "
Global Const EQ_ROTATE As String * 10 = "ROTATE    "



Global Const MANUALCOLOR As Integer = 1          'BLUE
Global Const LOADCOLOR As Integer = 10           'THIN_GREEN
Global Const ForbidColor As Integer = 11         'THIN_BLUE
Global Const ErrColor As Integer = 12            'THIN_RED
Global Const AUTOCOLOR As Integer = 14           'YELLOW
Global Const NOLDCOLOR As Integer = 15           'WHITE

Global Const SG_EMPTY As String * 1 = "0"
Global Const SG_LKIN As String * 1 = "S"
Global Const SG_LKOUT As String * 1 = "R"
Global Const SG_STOR As String * 1 = "F"
Global Const SG_ERR As String * 1 = "E"          '問題庫格
Global Const SG_NODATA As String * 1 = "N"       '在庫但無庫存資料庫格

Global Const MV_NONE As String * 1 = "0"
Global Const MV_NOW As String * 1 = "1"
Global Const MV_ERR1 As String * 1 = "2"
Global Const MV_ERR As String * 1 = "E"

'Global Const LOAD As String * 1 = "1"
Global Const NOLD As String * 1 = "0"

Global Const FB_NONE As String * 1 = "0"
Global Const FB_ALL As String * 1 = "X"

Global Const ST_RQST As String * 1 = "0"
Global Const ST_WAIT As String * 1 = "1"
Global Const ST_WORK As String * 1 = "2"

Global Const TR_IN As String * 1 = "S"          '入庫作業
Global Const TR_OUT As String * 1 = "R"         '出庫作業
Global Const TR_CW As String * 1 = "1"          '工作檔維護作業
Global Const TR_LO As String * 1 = "2"          '庫格維護作業
Global Const TR_TEST As String * 1 = "T"        '測試作業

Global Const TR_NORMAL_OUT As String * 1 = "1"  '一般出庫或指定庫格出庫
Global Const TR_PALLETS_OUT As String * 1 = "2" '整疊空棧板出庫
Global Const TR_QOUT As String * 1 = "3"        '盤點出庫
Global Const TR_MOVE_OUT As String * 1 = "4"    '搬倉出庫
Global Const TR_NO_OUT As String * 1 = "5"      '轉帳出庫
Global Const TR_000000_OUT As String * 1 = "6"  '卸載出庫到站號
Global Const TR_INVENTORY_OUT As String * 1 = "7" '平面倉庫出庫

Global Const TR_INVENTORY_IN As String * 1 = "9" '平面倉庫入庫
Global Const TR_NORMAL_IN As String * 1 = "A"   '一般入庫
Global Const TR_PALLETS_IN As String * 1 = "B"  '整疊空棧板入庫
Global Const TR_QIN As String * 1 = "C"         '盤點入庫
Global Const TR_MOVE_IN As String * 1 = "D"     '搬倉入庫
Global Const TR_BAR_IN As String * 1 = "E"      '條碼讀取自動入庫
Global Const TR_OUT_REIN As String * 1 = "F"    '撿取後再入庫
Global Const TR_MIX_IN As String * 1 = "G"      '併板入庫
Global Const TR_000000_IN As String * 1 = "H"   '卸載入庫到庫格
Global Const TR_MOVE_STNO As String * 1 = "I"   '站間搬運

Global Const TR_MIN As String * 1 = "K"         '維護新增
Global Const TR_MOUT As String * 1 = "L"        '維護刪除
Global Const TR_MOD As String * 1 = "M"         '維護更正前資料
Global Const TR_MOD1 As String * 1 = "N"        '維護更正後資料

Global Const TR_3000 As String * 1 = "O"        '空出荷
Global Const TR_3002 As String * 1 = "P"        '先入品
Global Const TR_CMP As String * 1 = "Q"         '手動完成
Global Const TR_DLT As String * 1 = "R"         '工作取消
Global Const TR_RESEND As String * 1 = "S"      '命令重送
Global Const TR_REJECT As String * 1 = "U"      '
Global Const TR_FULL As String * 1 = "V"        '強制在庫
Global Const TR_EMPTY As String * 1 = "W"       '強制空庫

'Global Const TR_CONVEYOR As String * 1 = "U"    '要移入之輸送機有殘留板
Public Sub init_parameter1()
  Dim jj As Long
  On Error GoTo err1_rtn
  Call Get_ComputerName
  If UCase(Left(computername, 8)) = "TPNPCENG" Then
    'computername = "SKGF3"
    computername = "GF2"
  End If
  If Left(computername, 3) = "GF2" Or Left(computername, 12) = "SKNPCEMD195" Then
    rcn_str = "Provider=SQLOLEDB.1;Data Source=10.115.80.195;Initial Catalog=sqldb;Persist Security Info=False;User ID=sa;password=nanya"
    project_name = "SK01": plant_name = "玻纖二"
  ElseIf Left(computername, 5) = "SKGF3" Then
    rcn_str = "Provider=SQLOLEDB.1;Data Source=192.115.64.77;Initial Catalog=sqldb;Persist Security Info=False;User ID=sa;password=nanya"
    project_name = "SK02": plant_name = "玻纖三"
  End If
  If Left(computername, 8) = "TPNPCENG" Then
     project_name = "SK01": plant_name = "玻纖二"
     rcn_str_local = "Provider=SQLOLEDB.1;Data Source=.;Initial Catalog=sqldb;Persist Security Info=False;User ID=sa;"
     rcn.ConnectionTimeout = 3: rcn.CursorLocation = adUseClient:  rcn.Open rcn_str_local
     rcn1.ConnectionTimeout = 3: rcn1.CursorLocation = adUseClient:  rcn1.Open rcn_str_local
  Else
     rcn.ConnectionTimeout = 3: rcn.CursorLocation = adUseClient:  rcn.Open rcn_str
     rcn1.ConnectionTimeout = 3: rcn1.CursorLocation = adUseClient:  rcn1.Open rcn_str
  End If
  cran_move_flag = True: cran_move_flag1 = "right_home_position": backup_dir = "c:\"
  refresh1_screen = 0: prog_buf = "    ": lono_buf = "        ": loginuser = "          ": loginpass = "     ": loginname = "        ": super = "N"
  err1_place = "": err2_cnt = 0: DB_TYPE = "SQL": PLC1_NUM = 7: STNO_NUM = 5
  Exit Sub
err1_rtn:
  Call MsgBox("test=" & Err.Number & "," & Err.Description)
  Call err2_rtn("init_parameter1")
End Sub
Public Sub backup_rtn()
  On Error GoTo err1_rtn
  Call Backup("lo"): Call Backup("sg")
  If Left(project_name, 2) <> "SO" Then Call Backup("pd")
  If Left(project_name, 2) = "SK" Then
    Call Backup("p1"): Call Backup("p2"): Call Backup("p3"): Call Backup("p4"): Call Backup("p5"): Call Backup("p6"): Call Backup("p7")
    If Left(project_name, 4) <= "SK03" Then
      Call Backup("p8"): Call Backup("p9"): Call Backup("p0")
      If Left(project_name, 4) = "SK02" Then Call Backup("pb")
    End If
  Else
    If Left(project_name, 2) <> "SO" Then Call Backup("pt")
  End If
  Exit Sub
err1_rtn:
  Call err2_rtn("backup_rtn")
End Sub

Public Function menu_enable(mode As Boolean) As Boolean
  If menu_enable_rtn(mode) = False Then Exit Function
  If super = "Y" Then
    first_prog.A002.Enabled = True: first_prog.A007.Enabled = True: first_prog.A150.Enabled = True: first_prog.A151.Enabled = True: first_prog.A152.Enabled = True: first_prog.A153.Enabled = True: first_prog.A155.Enabled = True: first_prog.A156.Enabled = True: first_prog.A157.Enabled = True: first_prog.A159.Enabled = True: first_prog.A15A.Enabled = True: first_prog.A15B.Enabled = True: first_prog.A244.Enabled = True
  Else
    first_prog.A002.Enabled = False: first_prog.A007.Enabled = False: first_prog.A150.Enabled = False: first_prog.A151.Enabled = False: first_prog.A152.Enabled = False: first_prog.A153.Enabled = False: first_prog.A155.Enabled = False: first_prog.A156.Enabled = False: first_prog.A157.Enabled = False: first_prog.A159.Enabled = False: first_prog.A15A.Enabled = False: first_prog.A15B.Enabled = False ': first_prog.A244.Enabled = False
  End If
  If (awno = "1" And Mid(computername, 4, 4) < "1A70") Or super = "Y" Then
    first_prog.A243.Enabled = True: first_prog.A244.Enabled = True
  Else
    first_prog.A243.Enabled = False: first_prog.A244.Enabled = False
  End If
  If ((awno = "1" And Mid(computername, 4, 1) = "1") Or super = "Y") And Left(project_name, 4) = "SK04" Then first_prog.A245.Visible = True Else first_prog.A245.Visible = False
  If (awno = "1" And Mid(computername, 4, 4) >= "1A70") Or super = "Y" Then first_prog.A242.Enabled = True Else first_prog.A242.Enabled = False:
  If check_right(awno, 7) Then '盤點作業
    first_prog.A313.Enabled = True: first_prog.A340.Enabled = True: first_prog.A590.Enabled = True
  Else
    first_prog.A313.Enabled = False: first_prog.A340.Enabled = False: first_prog.A590.Enabled = False
  End If
  first_prog.A313.Visible = False: first_prog.A340.Visible = False: first_prog.A590.Visible = False
  If check_right(awno, 11) Then first_prog.A350.Enabled = True Else first_prog.A350.Enabled = False
  
  first_prog.A1B0.Enabled = True '使用者更改密碼作業
  If check_right(awno, 16) Then '列印作業
    first_prog.A511.Enabled = True '入出庫日報表列印作業
'    first_prog.A610.Enabled = True '庫存明細表列印作業
    first_prog.A620.Enabled = True '庫存彙總表列印作業
    first_prog.A590.Enabled = True '盤點差異明細表列印作業
  Else
    first_prog.A511.Enabled = False '入出庫日報表列印作業
'    first_prog.A610.Enabled = False '庫存明細表列印作業
    first_prog.A620.Enabled = False '庫存彙總表列印作業
    first_prog.A590.Enabled = False '盤點差異明細表列印作業
  End If
  If Left(project_name, 2) = "SA" Then
    first_prog.A111.Visible = False
  ElseIf Left(project_name, 4) = "SK01" Then
    first_prog.A251.Visible = True
  ElseIf Left(project_name, 4) = "SK02" Then
    first_prog.A251.Visible = False
  ElseIf Left(project_name, 4) = "SK03" Then
    If super = "Y" Or awno = "5" Then
      first_prog.A248.Visible = True: first_prog.A359.Visible = True
    Else
      first_prog.A248.Visible = False: first_prog.A359.Visible = False
    End If
  End If
  first_prog.A1E0.Visible = False
  first_prog.A311.Visible = True '撿取後再入庫
  If awno = "1" Then
    first_prog.A312.Visible = False '併板後再入庫
    first_prog.A314.Visible = False '入庫作業
    first_prog.A318.Visible = False '入庫作業
    first_prog.A330.Visible = False '出庫作業
    first_prog.A330.Caption = "成品出庫作業-S330"
    first_prog.A350.Caption = "庫存查詢及指定庫格出庫作業-S350"
    first_prog.A351.Visible = True
    first_prog.A360.Caption = "庫格資料查詢及維護作-S360"
    first_prog.A361.Visible = False: first_prog.A362.Visible = False
    first_prog.A5A0.Visible = False '原料及成品庫格佔用率統計查詢作業
    If UCase(Left(computername, 5)) = "NANYA" Or UCase(Mid(computername, 4, 2)) = "1G" Then
    '  first_prog.A515.Enabled = True: first_prog.A332.Enabled = True: first_prog.A5B0.Enabled = True: first_prog.A333.Enabled = True
    Else
    '  first_prog.A515.Enabled = False: first_prog.A332.Enabled = False: first_prog.A5B0.Enabled = False: first_prog.A333.Enabled = False
    End If
  ElseIf awno = "5" Or super = "Y" Or Mid(computername, 4, 4) = "1A90" Or (Left(computername, 9) = "SERVER-SK" Or Mid(computername, 4, 4) = "0000") Then
    first_prog.A352.Enabled = True
  End If
  If awno = "2" Or awno = "3" Or super = "Y" Then
    first_prog.A354.Enabled = True
  Else
    first_prog.A354.Enabled = False
  End If
  If awno = "4" Or super = "Y" Then
    first_prog.A35A.Visible = True: first_prog.A35B.Visible = True
  Else
    first_prog.A35A.Visible = False: first_prog.A35B.Visible = False
  End If
  mode = True
  If mode = False Then 'Disable main menu
    first_prog.WindowState = 1
  Else
    first_prog.WindowState = 2: prog = "S001": Call Sys_MsgShow(first_prog, "請選擇工作別"): Screen.MousePointer = 0: first_prog_large = True
  End If
  first_prog.manage.Enabled = mode: first_prog.basic.Enabled = mode: first_prog.store.Enabled = mode: first_prog.inquiry.Enabled = mode: first_prog.report.Enabled = mode: first_prog.Mono.Enabled = mode: first_prog.login.Enabled = mode: first_prog.quit.Enabled = mode
  menu_enable = True
End Function
Public Function get_st_text(opno As String) As String
'改下面1行
  get_st_text = UCase(Mid(computername, 5, STNO_NUM)): Exit Function
  If UCase(Mid(computername, 4, 4)) = "1A00" Then
    If opno = "R" Then get_st_text = "A03 "
  ElseIf UCase(Mid(computername, 4, 4)) = "1A10" Then
    If opno = "R" Then get_st_text = "A12 "
  ElseIf UCase(Mid(computername, 4, 4)) = "1A30" Then
    If opno = "R" Then get_st_text = "A32 "
  End If
End Function
Public Function get_cran_no(awno1 As String, lono As String) As Integer
  Dim ii As Integer, cnt As Integer, brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  If is_lono(awno1, lono) Then
    cnt = 0
    For ii = 1 To Val(awno1) - 1
      brs2.Open "select count(*) from cr where cr_awno='" & CStr(ii) & "' and cr_mveq like '#%'", rcn1, adOpenKeyset, adLockOptimistic
      cnt = cnt + brs2(0): Set brs2 = Nothing
    Next ii
    If (Val(Left(lono, 2)) Mod 2) = 1 Then ii = Val(Left(lono, 2)) + 1 Else ii = Val(Left(lono, 2))
'   get_cran_no=((Val(Left(lono, 2)) + 1) \ 2)+cnt
    get_cran_no = (ii \ 2) + cnt
    If get_cran_no >= 10 Then get_cran_no = 0
  Else
    If awno1 = "1" Or Left(lono, 1) = "A" Then get_cran_no = 1
    If awno1 = "2" Or Left(lono, 1) = "B" Then get_cran_no = 2
    If awno1 = "3" Or Left(lono, 1) = "C" Then get_cran_no = 3
    If awno1 = "4" Or Left(lono, 1) = "D" Then get_cran_no = 4
    If awno1 = "5" Or Left(lono, 1) = "E" Then get_cran_no = 5
    If awno1 = "6" Or Left(lono, 1) = "E" Then get_cran_no = 5
  End If
  Exit Function
err1_rtn:
  Call err2_rtn("get_cran_no")
End Function

Public Sub put_sg_for_pallet(awno1 As String, lono As String, palletname As String, palletnum As Integer)
  rcn1.Execute "delete from sg where sg_awno='" & awno1 & "' and sg_lono='" & lono & "'"
  '                                   awno            lono           ptno                       quty              pqty sitm             spec odno        ltno            mode0 mark quty1          lot1  quty2 lot2     quty3 lot3            weight
  rcn1.Execute "insert into sg values('" & awno1 & "','" & lono & "','" & Trim(palletname) & "'," & palletnum & ",0,'" & get_now() & "','0','          ','             ','P','0'," & palletnum & ",'          ',0,'          ',0,'          ',0)"
  rcn1.Execute "update lo set lo_prog='    ' where lo_awno='" & awno1 & "' and lo_lono='" & lono & "'"
End Sub

Public Function get_zone(awno1 As String, zone As Integer) As String
  get_zone = "END"
  If Left(project_name, 4) = "SK01" Or Left(project_name, 4) = "SK02" Then
    If awno1 = "1" Then
      If zone = 10 Then get_zone = "原紗"
    ElseIf awno1 = "2" Then
      If zone = 10 Then get_zone = "經軸"
    ElseIf awno1 = "3" Then
      If zone = 10 Then get_zone = "織軸"
      If zone = 11 Then get_zone = "虛擬庫格"
    ElseIf awno1 = "4" Then
      If zone = 10 Then get_zone = "胚布小"
      If Left(project_name, 4) = "SK02" Then
        If zone = 11 Then get_zone = "胚布大"
      End If
      'If zone = 11 Then get_zone = "胚檢"
    ElseIf awno1 = "5" Then
      If zone = 10 Then get_zone = "成品布" '1-6層
      If zone = 11 Then get_zone = "整疊空棧板" '7-25層
      If Left(project_name, 4) = "SK02" Then
        If zone = 12 Then get_zone = "空棧板外售"
      End If
    End If
  ElseIf Left(project_name, 4) >= "SK04" And Left(project_name, 4) <= "SK06" Then
    If awno1 = "1" Then
      If zone = 10 Then get_zone = "保稅原紗"
      If zone = 11 Then get_zone = "完稅原紗"
    ElseIf awno1 = "2" Then
      If zone = 10 Then get_zone = "保稅經軸"
      If zone = 11 Then get_zone = "完稅經軸"
    ElseIf awno1 = "3" Then
      If zone = 10 Then get_zone = "保稅織軸"
      If zone = 11 Then get_zone = "完稅織軸"
      If zone = 12 Then get_zone = "穿綜車"
      If zone = 13 Then get_zone = "虛擬庫格"
    ElseIf awno1 = "4" Then
      If zone = 10 Then get_zone = "保稅胚布"
      If zone = 11 Then get_zone = "完稅胚布"
      If zone = 12 Then get_zone = "保稅胚檢"
      If zone = 13 Then get_zone = "完稅胚檢"
    ElseIf awno1 = "5" Then
      If zone = 10 Then get_zone = "保稅成品布"
      If zone = 11 Then get_zone = "完稅成品布"
      If zone = 12 Then get_zone = "保稅成檢或退布"
      If zone = 13 Then get_zone = "完稅成檢或退布"
      If zone = 14 Then get_zone = "整疊空棧板"
    ElseIf awno1 = "6" Then
      If zone = 10 Then get_zone = "平面倉庫"
    End If
  ElseIf Left(project_name, 4) = "SK03" Then
    If awno1 = "1" Then
      If zone = 10 Then get_zone = "原紗"
    ElseIf awno1 = "2" Then
      If zone = 10 Then get_zone = "經軸"
    ElseIf awno1 = "3" Then
      If zone = 10 Then get_zone = "織軸"
      If zone = 11 Then get_zone = "穿綜車"
      If zone = 12 Then get_zone = "虛擬庫格"
    ElseIf awno1 = "4" Then
      If zone = 10 Then get_zone = "胚布"
      'If zone = 11 Then get_zone = "胚檢"
    ElseIf awno1 = "5" Then
      If zone = 10 Then get_zone = "小成品布" '1-6層
      If zone = 11 Then get_zone = "大成品布" '7-25層
      If zone = 12 Then get_zone = "小整疊空棧板" '1-6層
      If zone = 13 Then get_zone = "大整疊空棧板" '7-25層
    ElseIf awno1 = "6" Then
      If zone = 10 Then get_zone = "經軸"
    End If
  End If
End Function
Public Function get_awnm(awnm As String) As String
  Dim arr1, i As Integer
  arr1 = Array("", "原紗自動倉庫系統", "經軸自動倉庫系統", "織軸自動倉庫系統", "胚布自動倉庫系統", "成品自動倉庫系統", "                ")
  If IsNumeric(awnm) Then
    If Val(awnm) > AW_NUM Then Exit Function
    get_awnm = arr1(Val(awnm))
  Else
    For i = 1 To AW_NUM
      If awnm = arr1(i) Then
        get_awnm = CStr(i): Exit Function
      End If
    Next i
  End If
End Function

Public Function get_srnm(srid As String) As String
  Dim buf, buf1, i As Integer
  buf = Array("", "1", "2", "3", "4", "5", "6", "7", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "END")
  buf1 = Array("", "一般出庫      ", "空棧板出庫    ", "盤點出庫      ", "搬倉出庫      ", "轉帳出庫      ", "卸載出庫到站號", "平面倉庫出庫  ", "平面倉庫入庫  ", "一般入庫      ", "空棧板入庫    ", "盤點後再入庫  ", "搬倉入庫      ", "條碼讀取自動入", "撿取後再入庫  ", "併板後再入庫  ", "卸載入庫到庫格", "站間搬運      ", "維護新增      ", "維護刪除      ", "維護更正前資料", "維護更正後資料", "空出荷        ", "先入品        ", "手動完成      ", "工作取消      ", "命令重送      ", "測試作業      ", "異常排出      ", "強制在庫      ", "強制空庫      ")
  get_srnm = "              ": i = 1
  Do
    If Left(buf(i), 3) = "END" Then Exit Do
    If srid = buf(i) Then
      get_srnm = buf1(i): Exit Function
    End If
    i = i + 1
  Loop
  If srid = TR_NORMAL_IN Then
    If awno = "1" Then get_srnm = "裂片玻璃入庫  "
    If awno = "2" Then get_srnm = "玻璃或SMD入庫 "
    If awno = "3" Then get_srnm = "原料或成品入庫"
  ElseIf srid = TR_NORMAL_OUT Then
    If awno = "1" Then get_srnm = "裂片玻璃出庫  "
    If awno = "2" Then get_srnm = "玻璃或SMD出庫 "
    If awno = "3" Then get_srnm = "原料或成品出庫"
  End If
End Function
Public Function get_awno(port_id As Integer) As String
  If Left(prog, 4) = "CRAN" Then
    get_awno = Mid(prog, 5, 1)
  ElseIf Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Then
    If Left(prog, 4) = "PLC6" Then get_awno = "5" Else get_awno = Mid(prog, 4, 1)
  ElseIf Left(prog, 4) = "MONO" Then
    get_awno = "1"
  Else
    MsgBox ("無法決定get_awno")
  End If
  Exit Function
End Function
    '針對工作檔維護作業而言,主要只是作投入tr檔
Public Sub put_tr(awno1 As String, srid As String, palt1 As Integer, lono1 As String)
  Dim brs2 As New ADODB.Recordset
  Dim lib_buf As String, opno As String, lono As String, stno As String, name1 As String, sitm_buf1 As String, sitm11 As String, odno As String, ptno As String, sitm As String, ltno As String
  Dim ii As Integer, jj As Integer, palt As Integer
  Dim sg_exist As Boolean
  Dim quty_buf, quty, pqty As Integer
  palt = palt1: lono = lono1: sitm11 = get_now()
  opno = get_opno(srid): stno = "      "
  If opno = TR_CW Or opno = TR_LO Then
      brs2.Open "select pd_name from pd where pd_user='" & loginuser & "'", rcn1, adOpenKeyset, adLockOptimistic
      name1 = brs2.Fields(0): Set brs2 = Nothing
  End If
  If palt <> 0 Then 'opno=TR_CW or TR_IN or TR_OUT
      brs2.Open "select cw_lono from cw where cw_awno='" & awno1 & "' and cw_palt=" & palt, rcn1, adOpenKeyset, adLockOptimistic
      If Not brs2.EOF Then
        '考慮同一庫格有多筆資料時必須使用sg
        lono = brs2(0): Set brs2 = Nothing
        If is_lono(awno1, lono) Then GoTo rtn
      End If
      Set brs2 = Nothing
      If opno = TR_CW Or opno = TR_LO Then
      'If opno = TR_CW Then
          rcn1.Execute "update cw set cw_name='" & name1 & "' where cw_awno='" & awno1 & "' and cw_palt=" & palt
      End If
      rcn1.Execute "insert into tr select * from cw where cw_awno='" & awno1 & "' and cw_palt=" & palt
      brs2.Open "select cw_sitm from cw where cw_awno='" & awno1 & "' and cw_palt=" & palt, rcn1, adOpenKeyset, adLockOptimistic
      rcn1.Execute "update tr set tr_srid='" & srid & "',tr_srnm='" & get_srnm(srid) & "',tr_opno='" & opno & "',tr_sitm='" & sitm11 & "',tr_date='" & Left(sitm11, 10) & "' where tr_awno='" & awno1 & "' and tr_palt=" & palt & " and tr_sitm='" & brs2!cw_sitm & "'"
      Set brs2 = Nothing
      brs2.Open "select * from cw where cw_awno='" & awno1 & "' and cw_palt=" & palt, rcn1, adOpenKeyset, adLockOptimistic
      Call backup1("tr" & awno1, brs2)
      Set brs2 = Nothing
  Else 'opno=TR_LO
rtn:
      If opno = TR_CW Or opno = TR_LO Then
        rcn1.Execute "update sg set sg_name='" & name1 & "' where sg_awno='" & awno1 & "' and sg_lono='" & lono & "'"
      End If
      rcn1.Execute "insert into tr select * from sg where sg_awno='" & awno1 & "' and sg_lono='" & lono & "'"
      brs2.Open "select sg_sitm from sg where sg_awno='" & awno1 & "' and sg_lono='" & lono & "'", rcn1, adOpenKeyset, adLockOptimistic
      Do
        If brs2.EOF Then Exit Do
        rcn1.Execute "update tr set tr_srid='" & srid & "',tr_srnm='" & get_srnm(srid) & "',tr_opno='" & opno & "',tr_sitm='" & sitm11 & "',tr_date='" & Left(sitm11, 10) & "' where tr_awno='" & awno1 & "' and tr_lono='" & lono & "' and tr_sitm='" & brs2!sg_sitm & "'"
        sitm11 = Left(sitm11, 17) & Right("00" & (Val(Mid(sitm11, 18, 2)) + 1), 2)
        brs2.MoveNext
      Loop
      Set brs2 = Nothing
      brs2.Open "select * from sg where sg_awno='" & awno1 & "' and sg_lono='" & lono & "'", rcn1, adOpenKeyset, adLockOptimistic
      Call backup1("tr" & awno1, brs2)
      Set brs2 = Nothing
  End If
  'If opno = TR_IN Or opno = TR_CW Then
  'rcn1.Execute("insert into tr values('" & awno1 & "'," & brs_cw!cw_palt & ",'" & Trim(lono) & "','" & brs2!sg_ptno & "'," & brs2!sg_quty & "," & brs2!sg_pqty & ",'" & brs2!sg_sitm & "','" & brs2!sg_sitm & "','" & brs2!sg_spec & "','" & brs2!sg_odno & "','" & brs2!sg_ltno & "','" & brs2!sg_mark & "','" & Left(brs2!sg_sitm, 10) & "','" & opno & "','" & srid & "','" & Trim(stno) & "','" & Trim(name1) & "','" & pick & "','" & brs_cw!cw_now & "->" & brs_cw!cw_nxt & "','" & get_srnm(srid) & "','" & brs_lo!lo_prog & "'," & brs2!sg_quty1 & ",'" & brs2!sg_cusno & "','" & brs2!sg_cusname & "','" & brs2!sg_glue & "','" & brs2!sg_mchno & "','" & brs2!sg_bar & "','" & brs2!sg_seqno & "'," & brs2!sg_quty2 & ",'" & brs2!sg_tax & "','" & brs2!sg_err1 & "'," & brs2!sg_mtk & ",'" & brs2!sg_bill & "','" & brs2!sg_lot & "','" & brs_lo!lo_mhno & "','" & brs_lo!lo_class & "','" & brs_lo!lo_stat & "','" & brs_lo!lo_whno & "','" & brs_lo!lo_type0 & "','" & Left(brs2!sg_sitm, 7) & "')")
  'Else
  '    rcn1.Execute("insert into tr values('" & awno1 & "'," & brs_cw!cw_palt & ",'" & Trim(lono) & "','" & brs2!sg_ptno & "'," & brs2!sg_quty & "," & brs2!sg_pqty & ",'" & sitm_buf1 & "','" & brs2!sg_sitm & "','" & brs2!sg_spec & "','" & brs2!sg_odno & "','" & brs2!sg_ltno & "','" & brs2!sg_mark & "','" & Left(sitm_buf1, 10) & "','" & opno & "','" & srid & "','" & Trim(stno) & "','" & Trim(name1) & "','" & pick & "','" & brs_cw!cw_now & "->" & brs_cw!cw_nxt & "','" & get_srnm(srid) & "','" & brs_lo!lo_prog & "'," & brs2!sg_quty1 & ",'" & brs2!sg_cusno & "','" & brs2!sg_cusname & "','" & brs2!sg_glue & "','" & brs2!sg_mchno & "','" & brs2!sg_bar & "','" & brs2!sg_seqno & "'," & brs2!sg_quty2 & ",'" & brs2!sg_tax & "','" & brs2!sg_err1 & "'," & brs2!sg_mtk & ",'" & brs2!sg_bill & "','" & brs2!sg_lot & "','" & brs_lo!lo_mhno & "','" & brs_lo!lo_class & "','" & brs_lo!lo_stat & "','" & brs_lo!lo_whno & "','" & brs_lo!lo_type0 & "','" & Left(sitm_buf1, 7) & "')")
  '    sitm_buf1 = Left(sitm_buf1, 17) & Right("00" & (Val(Mid(sitm_buf1, 18, 2)) + 1), 2)
  'End If
  If opno = TR_IN Or opno = TR_OUT Or opno = TR_TEST Then
      lib_buf = get_field_value(awno1, "cw_palt", "I", CStr(palt), "cw_from", "S")
      If is_lono(awno1, lib_buf) Then
          rcn1.Execute "update lo set lo_sgst='0' where lo_awno='" & awno1 & "' and lo_lono='" & lib_buf & "'"
          rcn1.Execute "delete sg where sg_awno='" & awno1 & "' and sg_lono='" & lib_buf & "'"
      End If
      lib_buf = get_field_value(awno1, "cw_palt", "I", CStr(palt), "cw_to", "S")
      If is_lono(awno1, lib_buf) Then
          brs2.Open "select * from sg where sg_awno='" & awno1 & "' and sg_lono='" & lib_buf & "'", rcn1, adOpenKeyset, adLockOptimistic
          If brs2.EOF Then
              rcn1.Execute "update lo set lo_sgst='E' where lo_awno='" & awno1 & "' and lo_lono='" & lib_buf & "'"
          Else
              rcn1.Execute "update lo set lo_sgst='F' where lo_awno='" & awno1 & "' and lo_lono='" & lib_buf & "'"
          End If
          Set brs2 = Nothing
      End If
      rcn1.Execute "delete cw where cw_awno='" & awno1 & "' and cw_palt='" & palt & "'"
  End If
rtn1:
  Call syslog("", 0, 0, "put_tr ok,srid=" & get_srnm(srid) & ",palt=" & palt & ",lono=" & lono)
  'err1_place = backup_dir
  'ii = FreeFile() : FileOpen(ii, backup_dir & "tr.log", OpenMode.Append) : err1_place = ""
  'PrintLine(ii, get_srnm(srid) & "|倉庫" & awno1 & "|庫格" & lono & "|站號" & stno & "|料號" & ptno & "|庫存量" & quty & "|出庫量" & pqty & "|訂號" & odno & "|時間" & sitm & "|序號" & palt) : FileClose(ii)
  Exit Sub
err1_rtn:
  Call err2_rtn("put_tr")
End Sub

Public Function handle_3000_3001(awno As String, palt As Integer, err1 As Integer) As Boolean
  Dim rcn2 As New ADODB.Connection, ii As Integer, brs2 As New ADODB.Recordset, brs3 As New ADODB.Recordset, lono1 As String, zone As String
  On Error GoTo err1_rtn
  handle_3000_3001 = False
  brs2.Open "select * from cw where cw_awno='" & awno & "' and cw_palt=" & palt, rcn1, adOpenKeyset, adLockOptimistic
  If Not brs2.EOF Then
    If err1 = 3000 Then '空出荷
    'If brs2!cw_err = 153 Or (Left(project_name, 2) <> "SU" And brs2!cw_err = 605) Or brs2!cw_err = 606 Or brs2!cw_err = 3000 Then '空出荷
      If Left(prog, 1) = "S" Then
        Response = MsgBox(brs2!cw_mveq & "空出荷異常!" & "序號" & palt, , "訊息")
      End If
      rcn1.Execute "update lo set lo_sgst='0' where lo_awno='" & awno & "' and lo_lono='" & brs2!cw_from & "'"
      rcn1.Execute "delete from sg where sg_awno='" & awno & "' and sg_lono='" & brs2!cw_from & "' "
      If Left(project_name, 4) = "SK02" And awno = "5" Then
        rcn1.Execute "update cr set cr_err=0,cr_rest='5',cr_mvst='0' where cr_mveq='" & brs2!cw_mveq & "' "
      Else
        rcn1.Execute "update cr set cr_err=0,cr_rest='2',cr_mvst='0' where cr_mveq='" & brs2!cw_mveq & "' "
      End If
      Call put_tr(awno, TR_3000, palt, "      ")
      rcn1.Execute "delete from cw where cw_awno='" & awno & "' and cw_palt=" & palt
      If Left(prog, 1) = "S" Then
        Response = MsgBox(brs2!cw_mveq & "原出庫庫格'" & brs2!cw_from & "'狀態改為空庫,出庫作業取消,請自行選擇其它庫格出庫!", , "訊息")
      Else
        Call syslog1("handle_3000_3001", palt, 0, brs2!cw_mveq & "原出庫庫格'" & brs2!cw_from & "'狀態改為空庫,出庫作業取消")
      End If
    ElseIf err1 = 3001 Then '先入品
    'ElseIf brs2!cw_err = 154 Or brs2!cw_err = 614 Or brs2!cw_err = 3001 Or brs2!cw_err = 3002 Then '先入品
      If is_lono(awno, brs2!cw_to) Then
        If Left(prog, 1) = "S" Then
          Response = MsgBox(brs2!cw_mveq & "先入品異常!" & "序號" & palt, , "訊息")
        End If
        Call put_tr(awno, TR_3002, palt, "      ")
        brs3.Open "select * from lo where lo_awno='" & awno & "' and lo_lono='" & brs2!cw_to & "'", rcn1, adOpenKeyset, adLockOptimistic
        zone = brs3!lo_zone: Set brs3 = Nothing
'找替代庫格
        rcn2.ConnectionTimeout = 3: rcn2.CursorLocation = adUseClient: rcn2.Open rcn_str
        brs3.Open "select * from lo where lo_awno='" & awno & "' and lo_zone='" & zone & "' and lo_sgst='0' and lo_fbst='0' order by lo_sisq", rcn2, adOpenKeyset, adLockOptimistic
        Do While Not brs3.EOF
            If get_cran_no(awno, brs2!cw_to) = get_cran_no(awno, brs3!lo_lono) Then Exit Do
            brs3.MoveNext
        Loop
        If brs3.EOF Then
          If Left(prog, 1) = "S" Then
            Set brs3 = Nothing: rcn2.Close: Response = MsgBox(brs2!cw_mveq & "倉庫已滿倉,找不到空的替代庫格,無法自動更換庫格!,請按其它功能鍵", , "異常訊息")
            Set brs2 = Nothing: Exit Function
          Else
            Set brs3 = Nothing: rcn2.Close: Call syslog("handle_3000_3001", palt, 0, brs2!cw_mveq & "倉庫已滿倉,找不到空的替代庫格,無法自動更換庫格!")
            Set brs2 = Nothing:  Exit Function
          End If
        End If
        lono1 = brs3!lo_lono: Set brs3 = Nothing
        'rcn.Close: rcn.ConnectionTimeout = 3: rcn.CursorLocation = adUseClient: rcn.Open rcn_str
'若不作上一行動作則下一行指令會咬住,rcn會有問題
        rcn2.Execute "update lo set lo_sgst='S' where lo_awno='" & awno & "' and lo_lono='" & lono1 & "'"
        rcn2.Execute "update sg set sg_lono='" & lono1 & "',sg_to='" & lono1 & "' where sg_awno='" & awno & "' and sg_lono='" & brs2!cw_to & "'"
        rcn2.Execute "update lo set lo_sgst='E' where lo_awno='" & awno & "' and lo_lono='" & brs2!cw_to & "'"
'cran異常時,切手動後moti="9"(搬運完),電腦命令自動會完成(cw_stat="0"),再切回自動時原異常碼仍會送出,此時必須下reset D748=1 命令,則moti="9",異常碼變為0
'下reset D748=1 命令時,異常碼變為0,moti="9"(搬運完),在mapath中會造成命令完成(cw_stat="0"),因之針對先入品狀況,不可以讓命令完成,cw_stat必須維持"2",下二行再去改為"1"
        If Left(project_name, 4) = "SK02" And awno = "5" Then
          rcn2.Execute "update cr set cr_rest='5',cr_mvst='0' where cr_awno='" & awno & "' and cr_mveq='" & brs2!cw_mveq & "' "
        Else
          rcn2.Execute "update cr set cr_rest='2',cr_mvst='0' where cr_awno='" & awno & "' and cr_mveq='" & brs2!cw_mveq & "' "
        End If
        If Left(prog, 1) = "S" Then
          Response = MsgBox(brs2!cw_mveq & "原入庫庫格'" & brs2!cw_to & "'狀態改為問題庫格,找到另一個空的替代庫格'" & lono1 & "'" & "準備再入庫!", , "訊息")
        Else
          Call syslog1("handle_3000_3001", palt, 0, brs2!cw_mveq & "原入庫庫格'" & brs2!cw_to & "'狀態改為問題庫格,找到另一個空的替代庫格'" & lono1 & "'" & "準備再入庫!")
        End If
        rcn2.Execute "update cw set cw_err=0,cw_stat='1',cw_now='(車上)',cw_first='Y',cw_lono='" & lono1 & "',cw_to='" & lono1 & "',cw_nxt='" & lono1 & "' where cw_awno='" & awno & "' and cw_palt=" & palt
        Call put_tr(awno, TR_3002, palt, "      ")
        rcn2.Close
      Else
        rcn.Execute "update cr set cr_err=0,cr_rest='2',cr_mvst='0' where cr_mveq='" & brs2!cw_mveq & "' "
        If Left(prog, 1) = "S" Then
          Response = MsgBox(brs2!cw_mveq & "要卸載到輸送機" & brs2!cw_to & "時,卻偵測到有載異常,請先手動執行輸送機卸載然後按手動完成功能鍵", , "訊息")
        Else
          Call syslog1("handle_3000_3001", palt, 0, brs2!cw_mveq & "要卸載到輸送機" & brs2!cw_to & "時,卻偵測到有載異常,請先手動執行輸送機卸載然後按手動完成功能鍵")
        End If
        rcn.Execute "update cw set cw_err=0,cw_stat='1',cw_now='(車上)',cw_first='Y' where cw_awno='" & awno & "' and cw_palt=" & palt
        Call put_tr(awno, TR_3002, palt, "      ")
      End If
    End If
  End If
  Set brs2 = Nothing
  handle_3000_3001 = True
  Exit Function
err1_rtn:
  Call err2_rtn("handle_3000_3001")
End Function
Public Sub update_sg_cran(awno As String)
  Dim brs2 As New ADODB.Recordset, buf As String, ii As Integer
  On Error GoTo err1_rtn
  ii = 0: brs2.Open "select * from sg where sg_awno='" & awno & "'", rcn, adOpenKeyset, adLockOptimistic
  Do
    If brs2.EOF Then Exit Do
    buf = ((Val(Left(brs2!sg_lono, 2)) + 1) \ 2)
    If brs2!sg_cran <> ((Val(Left(brs2!sg_lono, 2)) + 1) \ 2) Then
      ii = ii + 1: Call syslog1(CStr(ii), Val(brs2!sg_cran), Val(((Val(Left(brs2!sg_lono, 2)) + 1) \ 2)), brs2!sg_ptno)
      rcn1.Execute "update sg set sg_cran='" & ((Val(Left(brs2!sg_lono, 2)) + 1) \ 2) & "' where sg_awno='" & awno & "' and sg_lono='" & brs2!sg_lono & "'"
    End If
    brs2.MoveNext
  Loop
  Set brs2 = Nothing
  Exit Sub
err1_rtn:
  Call err2_rtn("update_sg_cran")
End Sub


Public Sub Get_ComputerName()
  Dim jj As Long, computername1 As String, ii As Integer, buf As String
  On Error GoTo err1_rtn
  jj = 20: computername1 = String(20, 0) '不可小於20
  GetComputerName computername1, jj: computername1 = UCase(computername1)
  For ii = 1 To 12
    buf = Mid(computername1, ii, 1)
    If Not ((Mid(computername1, ii, 1) >= "A" And Mid(computername1, ii, 1) <= "Z") Or (Mid(computername1, ii, 1) >= "0" And Mid(computername1, ii, 1) <= "9") Or Mid(computername1, ii, 1) = "-" Or Mid(computername1, ii, 1) = "_") Then
      GoTo rtn1
    End If
  Next ii
rtn1:
  computername = Mid(computername1, 1, ii - 1)
  ii = Len(Trim(computername))
  Exit Sub
err1_rtn:
  Call err2_rtn("Get_ComputerName")
End Sub
Public Function update_palt1_stno1(mode As String, stno As String, palt1 As Integer, Optional stno1 As String) As Boolean
  Dim ii As Integer, cnt As Integer, brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  update_palt1_stno1 = True
  rcn1.Execute "update cm set cm_reply1=0 where cm_cvar='" & mode & "'"
  rcn1.Execute "update cm set cm_reply=1,cm_from='" & Left(stno, STNO_NUM) & "',cm_palt=" & palt1 & ",cm_to='" & Left(stno1, STNO_NUM) & "' where cm_cvar='" & mode & "'"
  If mode = "3" Or mode = "4" Then
    For ii = 1 To 20
      Call Wait(0.5)
      brs2.Open "select * from cm where cm_cvar='" & mode & "'", rcn1, adOpenKeyset, adLockOptimistic
      If Not brs2.EOF Then
        If brs2!cm_reply1 = 10 Then
          Set brs2 = Nothing: Exit For
        End If
      End If
      Set brs2 = Nothing
    Next ii
    If ii >= 21 Then
      rcn1.Execute "update cm set cm_reply=2 where cm_cvar='" & mode & "'"
      If mode = "3" Or mode = "4" Then
        MsgBox ("PLC timeout"): update_palt1_stno1 = False
      End If
      Exit Function
    End If
  Else
    Call Wait(4)
  End If
  rcn1.Execute "update cm set cm_reply=0 where cm_cvar='" & mode & "'"
  Exit Function
err1_rtn:
  Call err2_rtn("update_palt1_stno1")
End Function

Public Function check_ascii(port_id As Integer, buf As String, cnt As Integer, show1 As String) As Boolean
  Dim i As Integer
  For i = 1 To cnt
    If Not ((Mid(buf, i, 1) >= "0" And Mid(buf, i, 1) <= "9") Or Mid(buf, i, 1) = " " Or (Mid(buf, i, 1) >= "A" And Mid(buf, i, 1) <= "Z")) Then
      If show1 = "show" Then Call syslog("check_ascii", port_id, 0, "buf=" & Mid(buf, 1, cnt) & ",第" & i & "byte err", "show")
      check_ascii = False: Exit Function
    End If
  Next i
  check_ascii = True
  Exit Function
err1_rtn:
  Call err2_rtn("check_ascii")
End Function

Public Function put_bcc() As Byte
  Dim i() As Byte, j As Byte, ii As Integer
  i = StrConv("21K01", vbFromUnicode)
  j = 0
  For ii = 0 To 4
    j = i(ii) Xor j
  Next
  ReDim i(6)
  i(5) = 16
  j = j Xor i(5)
  i(6) = 3
  j = j Xor i(6)
  Exit Function
err1_rtn:
  Call err2_rtn("put_bcc")
End Function

Public Sub watch_dog_read(model As String, palt As Integer, cnt As Integer, prog2 As String)
  Dim brs2 As New ADODB.Recordset, cnt1 As Integer, palt1  As Integer
  On Error GoTo err1_rtn
  brs2.Open "select pa_" & model & " from pa", rcn, adOpenKeyset, adLockOptimistic
  If palt = brs2(0) Then
    cnt = cnt + 1
    If cnt >= 12 Then
      cnt1 = cnt: palt1 = palt: Call syslog1("", cnt1, palt1, "initial:" & prog2)
      Call Shell(App.Path & "\" & prog2 & ".exe", 1): palt = 0: cnt = 0
    End If
  Else
    palt = brs2(0): cnt = 0
  End If
  Set brs2 = Nothing
  Exit Sub
err1_rtn:
  Call err2_rtn("watch_dog_read")
End Sub
Public Sub watch_dog_write(model As String, palt As Integer)
  Dim brs2 As New ADODB.Recordset, sitm11 As String
  On Error GoTo err1_rtn
  If Left(model, 2) <> "pa" Then GoTo rtn
  Call syslog("", palt, 0, model)
  sitm11 = get_now()
  If (Val(Mid(sitm11, 18, 2)) Mod 3) = 0 Then
rtn:
    brs2.Open "select " & model & " from " & Left(model, 2), rcn, adOpenKeyset, adLockOptimistic
    palt = brs2(0) + 1: Set brs2 = Nothing
    If palt >= 32767 Then palt = 1
    rcn.Execute "update " & Left(model, 2) & " set " & model & " = " & palt
  End If
  Exit Sub
err1_rtn:
  Call err2_rtn("watch_dog_write")
End Sub

Public Sub init_parameter()
Dim jj As Long, brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  'rcn.BeginTrans: rcn.CommitTrans: rcn.RollbackTrans
  brs2.Open "select count(distinct(lo_awno)) from lo", rcn, adOpenKeyset, adLockOptimistic
  AW_NUM = brs2.Fields(0): Set brs2 = Nothing
  For jj = 1 To AW_NUM
    brs2.Open "select count(distinct(substring(lo_lono,1,2))) from lo where lo_awno='" & jj & "'", rcn, adOpenKeyset, adLockOptimistic
    bank(jj) = brs2.Fields(0): Set brs2 = Nothing
    brs2.Open "select count(distinct(substring(lo_lono,3,3))) from lo where lo_awno='" & jj & "'", rcn, adOpenKeyset, adLockOptimistic
    bay(jj) = brs2.Fields(0): Set brs2 = Nothing
    brs2.Open "select count(distinct(substring(lo_lono,6,3))) from lo where lo_awno='" & jj & "'", rcn, adOpenKeyset, adLockOptimistic
    LEVEL(jj) = brs2.Fields(0): Set brs2 = Nothing
  Next jj
  awno = "1"
  If Left(project_name, 4) = "SK01" Then
    If Mid(computername, 1, 9) = "GF2_PC200" Or Mid(computername, 1, 9) = "GF2_PC201" Then
      awno = "1"
    ElseIf Mid(computername, 1, 9) = "GF2_PC202" Then
      awno = "2"
    ElseIf Mid(computername, 1, 9) = "GF2_PC207" Or Mid(computername, 1, 9) = "GF2_PC211" Then
      awno = "3"
    ElseIf Mid(computername, 1, 9) = "GF2_PC208" Or Mid(computername, 1, 9) = "GF2_PC210" Or Mid(computername, 1, 9) = "GF2_PC214" Or Mid(computername, 1, 9) = "GF2_PC216" Then
      awno = "4"
    ElseIf Mid(computername, 1, 9) = "GF2_PC221" Then
      awno = "5"
    End If
  ElseIf Left(project_name, 4) = "SK02" Then
    awno = Mid(computername, 7, 1) 'SKGF3_1_01
    If awno < "1" Or awno > "5" Then awno = "5"
    '虛擬庫格作業
    bay(3) = 14
  End If
  brs2.Open "select count(*) from cr where cr_mveq like '#%'", rcn, adOpenKeyset, adLockOptimistic
  CRAN_NUM = brs2.Fields(0): Set brs2 = Nothing
  Exit Sub
err1_rtn:
  Call err2_rtn("init_parameter")
End Sub
                                     
'執行搬倉作業,sg檔及lo檔資料移轉
Public Function lono_data_move(awno1 As String, from1 As String, to1 As String) As Boolean
  Dim brs2 As New ADODB.Recordset, brs3 As New ADODB.Recordset
  brs2.Open "select * from sg where sg_awno='" & awno1 & "' and sg_lono='" & from1 & "'", rcn1, adOpenKeyset, adLockOptimistic
  If brs2.EOF Then
    Response = MsgBox("執行搬倉作業,但庫格編號=" & to1 & "將會沒有庫存資料,是否仍要執行入庫作業,倘若確定執行入庫作業時,庫格狀態將被改為問題庫格,請立即在{庫格資料查詢及維護作業-S360}中輸入庫存資料", vbOKCancel)
    If Response = vbCancel Then
      Set brs2 = Nothing: lono_data_move = True: Exit Function
    End If
    rcn1.Execute "update lo set lo_sgst='E' where lo_awno='" & awno1 & "' and lo_lono='" & to1 & "'"
  End If
  brs3.Open "select lo_prog from lo where lo_awno='" & awno1 & "' and lo_lono='" & from1 & "'", rcn1, adOpenKeyset, adLockOptimistic
  If Not brs3.EOF Then rcn1.Execute "update lo set lo_prog='" & brs3(0) & "' where lo_lono='" & to1 & "'"
  Set brs2 = Nothing: Set brs3 = Nothing
  rcn1.Execute "delete from sg where sg_awno='" & awno1 & "' and sg_lono='" & to1 & "'"
  rcn1.Execute "update sg set sg_lono='" & to1 & "' where sg_awno='" & awno1 & "' and sg_lono='" & from1 & "'"
  lono_data_move = False
End Function

Public Sub mapath_rtn()
  Dim err1 As Integer, ii As Integer, jj As Integer, lib_buf As String * 100, buf As String * 100, month1 As Integer
  On Error GoTo err1_rtn
  lib_buf = get_now() '2001/06/08 12:30:12'
  If Left(project_name, 2) = "SK" And Left(prog, 6) = "MAPATH" Then '上午12:01 , get_now()讀出值為2008/07/31 00:01
    If Left(lib_buf, 16) = (Left(get_string_value1("", "", "pa_sitm1"), 13) & ":01") Then
      Call syslog1("", 0, 0, "盤點轉檔ck start1")
      rcn.Execute "delete ck": rcn.Execute "insert into ck select * from sg where sg_mark='正常紗' or sg_mark='正常經軸' or sg_mark='異常經軸' or sg_mark='織軸' or sg_mark='未穿綜織軸' or sg_mark='已穿綜織軸' or sg_mark='打結織軸' or sg_mark='異常織軸' or sg_mark='虛擬庫格' or sg_mark='送C/D可外售' or sg_mark='送CD不可外售' or sg_mark='配布' or sg_mark='退漿' or sg_mark='胚檢' or sg_mark='成品布' or sg_mark='成檢' or sg_mark='待印刷' or sg_mark = '待改裁' or sg_mark = '待判' or sg_mark='客訴布' or sg_mark='牽料布' or sg_mark='暫存布' or sg_mark='追回布'"
      Call syslog1("", 0, 0, "盤點轉檔ck end1")
      Call Wait(61)
    End If
  End If
  '每日凌晨00:30:00-03或12:30:00-03自動作清檔動作
  If (Left(prog, 6) = "MAPATH" And Val(Mid(lib_buf, 15, 2)) = 0 And Val(Mid(lib_buf, 18, 2)) >= 0 And Val(Mid(lib_buf, 18, 2)) <= 3) Or _
     (Left(prog, 6) = "BACKUP" And (Val(Mid(lib_buf, 15, 2)) Mod 10) = 0 And Val(Mid(lib_buf, 18, 2)) >= 0 And Val(Mid(lib_buf, 18, 2)) <= 3) Or _
     (Left(prog, 6) <> "MAPATH" And Val(Mid(lib_buf, 15, 2)) = 30 And Val(Mid(lib_buf, 18, 2)) >= 0 And Val(Mid(lib_buf, 18, 2)) <= 3) Then
    'If ((Val(Mid(lib_buf, 12, 2)) Mod 2) = 0) And Val(Mid(lib_buf, 15, 2)) = 30 And Val(Mid(lib_buf, 18, 2)) >= 0 And Val(Mid(lib_buf, 18, 2)) <= 3 Then
    'If Left(prog, 6) <> "MAPATH" And Left(project_name, 2) <> "SO" Then GoTo rtn10
    If Val(Mid(lib_buf, 6, 2)) = 1 And Val(Mid(lib_buf, 9, 2)) >= 1 And Val(Mid(lib_buf, 9, 2)) <= 3 Then '1月1日清盤點選取日期
      If Left(project_name, 2) <> "SO" Then rcn.Execute "update ck set ck_date='          '"
    End If
    If Val(Mid(lib_buf, 9, 2)) = 1 Or Val(Mid(lib_buf, 9, 2)) = 15 Then '每月1號作清檔作業
      Call syslog1("", 0, 0, "mapath: delete tr,er,ck start")
      'put_cw()中
      'Else '針對站對站case,sg不存在的情況
      '  Global Const TR_LO As String * 1 = "2"          '庫格維護作業
      '  brs2.Open "select * from tr where tr_opno='2'", rcn1, adOpenKeyset, adLockOptimistic
      '  玻纖case,tr存6個月總計219203 records,select * from tr需時1 mins,所以tr不可太大
      '  brs2.Open "select * from tr", rcn1, adOpenKeyset, adLockOptimistic
      lib_buf = get_previous_month(12)
      If Left(project_name, 2) <> "SO" Then
        rcn.Execute "delete er where er_sitm <= '" & lib_buf & "'"
        rcn.Execute "delete ck where ck_sitm <= '" & lib_buf & "'"
      End If
      If Left(project_name, 2) = "SP" Then
        rcn.Execute "delete lj where lj_sitm <= '" & lib_buf & "'"
      End If
      For ii = 1 To 31
        Mid(lib_buf, 9, 2) = Format(ii, "00") 'date1 = Format(Now - ii, "YYYYMMDDhhmmss"): DoEvents
        rcn.Execute "delete tr where tr_date = '" & Left(lib_buf, 10) & "'"
        DoEvents
        'For jj = 0 To 24  even if no tr data,this for loop still spend much time
        '  rcn.Execute "delete tr where tr_sitm like '" & Left(lib_buf, 11) & Right("00" & jj, 2) & "%'"
        '  If (Left(project_name, 2) = "SK" ) Then
        '    rcn.Execute "delete mo where mo_sitm like '" & Left(lib_buf, 11) & Right("00" & jj, 2) & "%'"
        '  End If
        '  DoEvents
        'Next jj
      Next ii
      rcn.Execute "delete tr where tr_sitm <= '" & Left(lib_buf, 10) & "  '"
      If Left(project_name, 2) = "SK" Then
        rcn.Execute "delete mo where mo_sitm <= '" & Left(lib_buf, 10) & "  '"
      End If
      Call syslog1("", 0, 0, "mapath: delete tr,er,ck ok")
      ii = FreeFile
      'lib_buf = App.Path & "\cp.bat"
      Open Trim(App.Path & "\cp.bat") For Output As ii
      Print #ii, "copy c:\log\*.* c:\log1\" & Chr(13) & Chr(10) & "del /Q c:\log\*.*"
      Close ii
      'Shell指令只能在NT上執行
      Call Shell(Trim(App.Path & "\cp.bat"), 1)
      If Left(project_name, 2) = "SK" Then
        ii = FreeFile
        Open Trim(App.Path & "\del_shiftr.bat") For Output As ii
        Print #ii, "del /Q c:\system\shiftr_old\S" & Mid(lib_buf, 1, 4) & Mid(lib_buf, 6, 2) & "*.*"
        Close ii
        Call Shell(Trim(App.Path & "\del_shiftr.bat"), 1)
        
        lib_buf = get_previous_month(2) '1天兩班共1000筆
        For ii = 1 To 31
          Mid(lib_buf, 9, 2) = Format(ii, "00") 'date1 = Format(Now - ii, "YYYYMMDDhhmmss"): DoEvents
          rcn.Execute "delete mj where mj_mode='1' and substring(mj_sitm,1,10) = '" & Left(lib_buf, 10) & "'"
          rcn.Execute "delete mj where mj_mode='2' and substring(mj_sitm,1,10) = '" & Left(lib_buf, 10) & "'"
          rcn.Execute "delete mi where substring(mi_sitm,1,10) = '" & Left(lib_buf, 10) & "'"
          DoEvents
        Next ii
        rcn.Execute "delete mj where mj_mode='1' and mj_sitm <= '" & Left(lib_buf, 10) & "'"
        rcn.Execute "delete mj where mj_mode='2' and mj_sitm <= '" & Left(lib_buf, 10) & "'"
        rcn.Execute "delete mi where mi_sitm <= '" & Left(lib_buf, 10) & "'"
      End If
      GoTo rtn10
    Else
rtn10:
      If Left(prog, 6) = "BACKUP" And (Val(Mid(lib_buf, 15, 2)) Mod 30) = 0 Then
        Call syslog1("", 0, 0, prog & ": backup")
      Else
        Call syslog1("", 0, 0, prog & ": backup")
      End If
      Call backup_rtn
      For ii = 1 To AW_NUM
        If Dir(backup_dir & "tr" & ii & ".txt") <> "" Then Kill backup_dir & "tr" & ii & ".txt"
      Next ii
      Call syslog1("", 0, 0, prog & ": backup ok")
    End If
    Call Wait(3)
  End If
  DoEvents
  Exit Sub
err1_rtn:
  Call err2_rtn("mapath_rtn")
End Sub
Public Function get_previous_month(month1 As Integer) As String
  Dim lib_buf As String, ii As Integer
  lib_buf = get_now(): ii = Val(Mid(lib_buf, 6, 2))
  If ii > month1 Then
    ii = ii - month1
  Else
    ii = ii + 12 - month1: Mid(lib_buf, 1, 4) = Right("0000" & (Val(Mid(lib_buf, 1, 4)) - 1), 4)
  End If
  Mid(lib_buf, 6, 2) = Format(ii, "00"): get_previous_month = lib_buf
End Function

Public Sub Backup(ByRef mode As String)
  Dim brs2 As New ADODB.Recordset, fname As String, lib_buf As String, dir_buf As String
  On Error GoTo err1_rtn
  lib_buf = get_now() '2001/06/08 12:30:12'
  If (Val(Mid(lib_buf, 9, 2)) Mod 2) = 0 Then
    dir_buf = backup_dir & "backup2"
  Else
    dir_buf = backup_dir & "backup1"
  End If
  If Dir(dir_buf, vbDirectory) = "" Then MkDir (dir_buf)
  brs2.Open "select * from " & mode, rcn1, adOpenKeyset, adLockOptimistic
  If Not brs2.EOF Then
    fname = dir_buf & "\" & mode & ".txt"
    If Dir(fname) <> "" Then
      lib_buf = backup_dir & "system\" & mode & ".txt": FileCopy fname, lib_buf
      Kill (fname)
    End If
    Call backup1(mode, brs2)
  End If
  Set brs2 = Nothing
  Exit Sub
err1_rtn:
  If Left(Err.Description, 19) = "Invalid object name" Then
    MsgBox ("backup:" & mode & "," & Err.Description): Exit Sub
  End If
  Call err2_rtn("backup:" & mode)
End Sub
Public Sub backup1(ByRef mode As String, brs2 As ADODB.Recordset)
  Dim ii As Integer, jj As Integer, fname As String, lib_buf As String, dir_buf As String
  On Error GoTo err1_rtn
  lib_buf = get_now() '2001/06/08 12:30:12'
  If (Val(Mid(lib_buf, 9, 2)) Mod 2) = 0 Then
    dir_buf = backup_dir & "backup2"
  Else
    dir_buf = backup_dir & "backup1"
  End If
  If Dir(dir_buf, vbDirectory) = "" Then MkDir (dir_buf)
  fname = dir_buf & "\" & mode & ".txt"
  If Dir(fname) <> "" Then Kill (fname)
  ii = FreeFile()
  Open fname For Output As ii
' Open dir_buf & "\" & mode & ".txt" For Append As ii '不可用append,會一直長大
' If Dir(backup_dir & mode & ".txt") <> "" Then Kill fname
  lib_buf = ""
  For jj = 1 To brs2.Fields.Count 'rdo無Fields之屬性
    If jj < brs2.Fields.Count Then
      lib_buf = Trim(lib_buf) & brs2.Fields(jj - 1).name & ","
    Else
      lib_buf = Trim(lib_buf) & brs2.Fields(jj - 1).name
    End If
  Next jj
  Print #ii, lib_buf
  Do While Not brs2.EOF
    lib_buf = ""
    For jj = 1 To brs2.Fields.Count 'rdo無Fields之屬性
      lib_buf = Trim(lib_buf) & brs2.Fields(jj - 1) & ","
    Next jj
    '2,0,          ,0,01001002,01001002,        ,        ,        ,0,   ,N,2004/02/19 15:25:14,2004/02/19 15:25:14,2004/02/19,2,K,維護新增      ,系統管理,    ,0,0,空棧 ,106 ,          ,                    ,   ,   ,A1,QW11,A,          ,0,0,0,0,0,空棧板      ,0,      , ,0,  ,GFA123,          ,0,0,保稅,0,0,0,  ,0,0,       ,0,
    Print #ii, lib_buf
    brs2.MoveNext
  Loop
  Close ii
  Exit Sub
err1_rtn:
  Call err2_rtn("backup1:" & mode)
End Sub
Public Sub restore1(mode As String)
  Dim brs2 As New ADODB.Recordset, fname As String, buf As String, buf1 As String, ii As Integer, jj As Integer, kk As Integer
  Dim type1(100) As Integer, para
  On Error GoTo err1_rtn
  err1_place = backup_dir
  fname = mode & ".txt"
  If Dir(fname) = "" Then Exit Sub
  err1_place = ""
  brs2.Open "select * from " & Mid(mode, 12, 2), rcn1, adOpenKeyset, adLockOptimistic
  type1(0) = brs2.Fields.Count 'rdo無Fields之屬性
  For kk = 1 To brs2.Fields.Count
    type1(kk) = brs2.Fields(kk - 1).Type 'CHAR:129,INT:3,FLOAT:5
    buf = brs2.Fields(kk - 1).name 'sg_awno
    ii = brs2.Fields(kk - 1).DefinedSize 'INT:4,FLOAT:8
  Next kk
  Set brs2 = Nothing
  rcn1.Execute "delete " & Mid(mode, 12, 2)
  Open fname For Input As #1
  Line Input #1, buf
  Do While Not EOF(1)
    Line Input #1, buf
    buf1 = "insert into " & Mid(mode, 12, 2) & " values("
    For kk = 1 To type1(0)
      If kk = 1 Then
        jj = InStr(1, buf, ",", 0)
        para = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      Else
        para = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      End If
      If kk < type1(0) Then
        If type1(kk) = 129 Then buf1 = buf1 & "'" & Trim(para) & "',"
        If type1(kk) = 3 Or type1(kk) = 5 Then buf1 = buf1 & Val(para) & ","
      Else
        If type1(kk) = 129 Then buf1 = buf1 & "'" & Trim(para) & "')"
        If type1(kk) = 3 Or type1(kk) = 5 Then buf1 = buf1 & Val(para) & ")"
      End If
    Next kk
    rcn1.Execute buf1
  Loop
  Close #1
  Exit Sub
err1_rtn:
  Call err2_rtn("restore1:" & mode)
End Sub


Public Sub restore(mode As String)
  Dim brs2 As New ADODB.Recordset, fname As String, buf As String, ii As Integer, jj As Integer
  Dim awno0 As String, lono As String, zone As String, sgst As String, fbst As String, sisq As String, cran As String, sitm As String, date0 As String, prog As String, palt2 As Integer, quty2 As Integer, hilo As String
  Dim user As String, name As String, pass As String, dept As String, righ, super As String
  Dim ptno As String, spec As String, unit As String, quty As String, weight As String, zonename As String, awno1 As String, awno2 As String, awno3 As String, awno4 As String, awno5 As String, awno6 As String, date1 As String, date2 As String, date3 As String, date4 As String, date5 As String, date6 As String, ptno3 As String
  Dim ccl As String, bas1 As String, bas2 As String, wid1 As String, wid2 As String, leng As String
  Dim vendor As String, ltno2 As String, ltno As String, grade As String, process As String
  On Error GoTo err1_rtn
'注意
  err1_place = backup_dir
  fname = backup_dir & mode & ".txt"
  If Dir(fname) = "" Then Exit Sub
  err1_place = ""
  rcn1.Execute "delete " & mode
  Open fname For Input As #1
  Do While Not EOF(1)
    Line Input #1, buf
    If mode = "lo" Then
      jj = InStr(1, buf, ",", 0)
      awno0 = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      lono = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      zone = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      sgst = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      fbst = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      sisq = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      cran = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      sitm = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      If Left(project_name, 2) <> "SO" Then
        date0 = Mid(buf, ii + 1, jj - ii - 1): ii = jj
        prog = Mid(buf, ii + 1, 4)
        rcn1.Execute "insert into " & mode & " values('" & awno & "','" & lono & "','" & zone & "','" & sgst & "','" & fbst & "'," & Val(sisq) & ",'" & cran & "','" & sitm & "','" & date0 & "','" & prog & "')"
      Else
        date0 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        prog = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        palt2 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        quty2 = Mid(buf, ii + 1, jj - ii - 1): ii = jj
        hilo = Mid(buf, ii + 1, 1)
        rcn1.Execute "insert into " & mode & " values('" & awno & "','" & lono & "','" & zone & "','" & sgst & "','" & fbst & "'," & Val(sisq) & ",'" & cran & "','" & sitm & "','" & date0 & "','" & prog & "'," & Val(palt2) & "," & Val(quty2) & ",'" & hilo & "')"
      End If
    ElseIf mode = "pd" Then
      jj = InStr(1, buf, ",", 0)
      user = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      name = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      pass = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      dept = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      righ = Mid(buf, ii + 1, jj - ii - 1): ii = jj
      super = Mid(buf, ii + 1, 1)
      rcn1.Execute "insert into " & mode & " values('" & user & "','" & name & "','" & pass & "','" & dept & "','" & righ & "','" & super & "')"
    ElseIf mode = "pt" Then
      jj = InStr(1, buf, ",", 0)
      ptno = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      spec = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      If Left(project_name, 2) = "SF" Then
        bas1 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        bas2 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        wid1 = Mid(buf, ii + 1, jj - ii - 1): ii = jj
        wid2 = Mid(buf, ii + 1, 4)
        rcn1.Execute "insert into " & mode & " values('" & ptno & "','" & spec & "'," & Val(bas1) & "," & Val(bas2) & "," & Val(wid1) & "," & Val(wid2) & ")"
      Else
        unit = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        quty = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        weight = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        zonename = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        sitm = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        awno1 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        awno2 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        awno3 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        awno4 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        awno5 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        awno6 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        date1 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        date2 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        date3 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        date4 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
        If Left(project_name, 2) = "SJ" Then
          date5 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
          date6 = Mid(buf, ii + 1, jj - ii - 1): ii = jj
          ptno3 = Mid(buf, ii + 1, 25)
          rcn1.Execute "insert into " & mode & " values('" & ptno & "','" & spec & "','" & unit & "'," & Val(quty) & "," & Val(weight) & ",'" & zonename & "','" & sitm & "','" & awno1 & "','" & awno2 & "','" & awno3 & "','" & awno4 & "','" & awno5 & "','" & awno6 & "','" & date1 & "','" & date2 & "','" & date3 & "','" & date4 & "','" & date5 & "','" & date6 & "','" & ptno3 & "')"
        Else
          date5 = Mid(buf, ii + 1, jj - ii - 1): ii = jj
          date6 = Mid(buf, ii + 1, 10)
          rcn1.Execute "insert into " & mode & " values('" & ptno & "','" & spec & "','" & unit & "'," & Val(quty) & "," & Val(weight) & ",'" & zonename & "','" & sitm & "','" & awno1 & "','" & awno2 & "','" & awno3 & "','" & awno4 & "','" & awno5 & "','" & awno6 & "','" & date1 & "','" & date2 & "','" & date3 & "','" & date4 & "','" & date5 & "','" & date6 & "')"
        End If
      End If
    ElseIf mode = "p1" Then
      jj = InStr(1, buf, ",", 0)
      ptno = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      spec = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      ccl = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      bas1 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      bas2 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      wid1 = Mid(buf, ii + 1, jj - ii - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      wid2 = Mid(buf, ii + 1, jj - ii - 1): ii = jj
      leng = Mid(buf, ii + 1, 4)
      rcn1.Execute "insert into " & mode & " values('" & ptno & "','" & spec & "','" & ccl & "'," & Val(bas1) & "," & Val(bas2) & "," & Val(wid1) & "," & Val(wid2) & "," & Val(leng) & ")"
    ElseIf mode = "p2" Then
      jj = InStr(1, buf, ",", 0)
      ptno = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      ltno2 = Mid(buf, ii + 1, jj - ii - 1): ii = jj
      ltno = Mid(buf, ii + 1, 20)
      rcn1.Execute "insert into " & mode & " values('" & ptno & "','" & ltno2 & "','" & ltno & "')"
    ElseIf mode = "p3" Then
      jj = InStr(1, buf, ",", 0)
      awno0 = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      grade = Mid(buf, ii + 1, jj - ii - 1): ii = jj
      rcn1.Execute "insert into " & mode & " values('" & awno & "','" & grade & "')"
    ElseIf mode = "p4" Then
      jj = InStr(1, buf, ",", 0)
      ltno = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      rcn1.Execute "insert into " & mode & " values('" & ltno & "')"
    ElseIf mode = "p5" Then
      jj = InStr(1, buf, ",", 0)
      process = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      rcn1.Execute "insert into " & mode & " values('" & process & "')"
    ElseIf mode = "p6" Then
      jj = InStr(1, buf, ",", 0)
      vendor = Mid(buf, 1, jj - 1): ii = jj: jj = InStr(jj + 1, buf, ",", 0)
      rcn1.Execute "insert into " & mode & " values('" & vendor & "')"
    End If
  Loop
  Close #1
  Exit Sub
err1_rtn:
  If err1_place = backup_dir Then
    MsgBox (backup_dir & "不存在"): err1_place = ""
  End If
  Call err2_rtn("restore:" & mode)
End Sub


Public Sub s001rtn(mode As String)
  Dim brs2 As New ADODB.Recordset, myvalue As String
  On Error GoTo err1_rtn
  Exit Sub
err1_rtn:
  Call err2_rtn(mode)
End Sub

Public Function menu_enable_rtn(mode As Boolean) As Boolean
  Dim brs2 As New ADODB.Recordset, ii As Integer
  On Error GoTo err1_rtn
  'If UCase(Left(computername, 3)) = "COM" Then
  '  first_prog.manage.Enabled = mode: first_prog.store.Enabled = mode: first_prog.inquiry.Enabled = mode: first_prog.report.Enabled = mode: first_prog.login.Enabled = mode
  '  menu_enable_rtn = False: Exit Function
  'End If
  Call s001rtn("Timer1_Timer")
  brs2.Open "select pd_super from pd where pd_user='" & loginuser & "'", rcn1, adOpenKeyset, adLockOptimistic
  If brs2.EOF Then
    If mode = False Then 'Disable main menu
      Set brs2 = Nothing: MsgBox ("請先執行系統登入的動作"): menu_enable_rtn = False: Exit Function
    End If
  End If
  Set brs2 = Nothing
  If super = "Y" Then
    first_prog.A160.Enabled = True: first_prog.A1A0.Enabled = True: first_prog.A1C0.Enabled = True: first_prog.A1D0.Visible = False: first_prog.A1E0.Enabled = True
  Else
    first_prog.A160.Enabled = False: first_prog.A1A0.Enabled = False: first_prog.A1C0.Enabled = False: first_prog.A1D0.Visible = False: first_prog.A1E0.Enabled = False
  End If
  If super = "Y" Or Mid(righ, (Val(awno) - 1) * 16 + 12, 1) = "1" Then first_prog.A170.Enabled = True Else first_prog.A170.Enabled = False
  If Left(project_name, 4) <> "SI04" Then
    If PALLET_NUM(Val(awno)) = 0 Then first_prog.A370.Visible = False Else first_prog.A370.Visible = True
    'If AW_NUM <= 1 Then first_prog.A580.Visible = False Else first_prog.A580.Visible = True '設備使用率統計查詢作業
    If check_right(awno, 4) Then '物流搬運作業
      first_prog.A250.Enabled = True
      If Left(project_name, 2) = "SI" Then
        first_prog.A251.Enabled = True ': first_prog.A252.Enabled = True
      End If
    Else
      first_prog.A250.Enabled = False
      If Left(project_name, 2) = "SI" Then
        first_prog.A251.Enabled = False ': first_prog.A252.Enabled = False
      End If
    End If
  End If
  'If check_right(awno, 10) Then '存區規劃
  '  first_prog.A1C0.Enabled = True
  '  If Left(project_name, 2) = "SG" Then
  '    first_prog.A1C1.Enabled = True: first_prog.A1C1.Visible = True
  '  End If
  'Else
  '  first_prog.A1C0.Enabled = False
  '  If Left(project_name, 2) = "SG" Then
  '    first_prog.A1C1.Enabled = False: first_prog.A1C1.Visible = False
  '  End If
  'End If
  menu_enable_rtn = True
  Exit Function
err1_rtn:
  Call err2_rtn("menu_enable_rtn")
End Function
'bit_no:0-15
Public Function bit_test(buf As Long, bit_no As Integer) As Boolean
  Dim ii As Integer, jj As Long
  jj = 1: ii = 1
  Do While ii <= bit_no
    jj = jj * 2: ii = ii + 1
  Loop
  bit_test = buf And jj 'bit_no=4,jj=16
End Function

Public Function pallet_auto_out(awno1 As String, stno As String, mark As String, Optional cran As String) As Boolean
  Dim cnt(8) As String, brs2 As New ADODB.Recordset, brs3 As New ADODB.Recordset, lono As String * 8, ii As Integer
  On Error GoTo err1_rtn
  pallet_auto_out = False
  If Left(project_name, 2) = "SP" And Left(stno, 1) = "A" Then
    If cran = "3" Then brs3.Open "select * from cw where cw_nxt='C014'", rcn, adOpenKeyset, adLockOptimistic
    If cran = "6" Then brs3.Open "select * from cw where cw_nxt='C022'", rcn, adOpenKeyset, adLockOptimistic
    If cran = "7" Then brs3.Open "select * from cw where cw_nxt='C026'", rcn, adOpenKeyset, adLockOptimistic
  Else
    brs3.Open "select * from cw where cw_to='" & stno & "'", rcn, adOpenKeyset, adLockOptimistic
  End If
  If Not brs3.EOF Then
    Call syslog("pallet_auto_out", brs3!cw_palt, Val(cran), "stno=" & stno & ",存在往迄站的工作,因之無法再投入工作檔")
    'If brs3!cw_err <> 4 Then rcn.Execute "update cw set cw_err=4 where cw_awno='" & awno1 & "' and cw_palt=" & brs3!cw_palt
  Else
    If Left(project_name, 2) = "SP" Then GoTo rtn1
    If Left(project_name, 4) = "SK02" Then GoTo rtn2
    If Not check_io_wrong(awno1, " ", Left(stno, STNO_NUM)) Then
rtn2:
      If Left(project_name, 4) = "SK02" Then
        brs2.Open "select * from sg,lo where sg_awno='" & awno1 & "' and sg_mark like '" & Trim(mark) & "%'  and sg_ptno like '" & Trim(cran) & "%' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
      ElseIf Left(project_name, 2) = "SK" Then
        brs2.Open "select * from sg,lo where sg_awno='" & awno1 & "' and sg_mark like '" & Trim(mark) & "%' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
      ElseIf Left(project_name, 2) = "SU" Then
        brs2.Open "select * from sg,lo where sg_awno='" & awno1 & "' and sg_ptno like '" & Trim(mark) & "%' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
      ElseIf Left(project_name, 2) = "SP" Then
rtn1:
        If cran >= "1" And cran <= "9" Then
          brs2.Open "select count(*) from sg,lo where sg_awno='" & awno1 & "' and sg_ptno like '" & Trim(mark) & "%' and sg_cran='" & cran & "' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F'", rcn1, adOpenKeyset, adLockOptimistic
          If brs2(0) <= 30 Then rcn1.Execute "update st set st_err=5 where st_awno='" & awno1 & "' and st_stno = '" & Left(stno, STNO_NUM) & "'"
          Set brs2 = Nothing
          brs2.Open "select * from sg,lo where sg_awno='" & awno1 & "' and sg_ptno like '" & Trim(mark) & "%' and sg_cran='" & cran & "' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
        Else
          cnt(0) = 0: cnt(3) = 0: cnt(6) = 0: cnt(7) = 0
          brs2.Open "select * from sg4 where sg_ptno like '" & Trim(mark) & "%' and (sg_cran='3' or sg_cran='6' or sg_cran='7')", rcn1, adOpenKeyset, adLockOptimistic
          Do
            If brs2.EOF Then Exit Do
            cnt(Val(brs2!sg_cran)) = brs2!sg_cnt
            brs2.MoveNext
          Loop
          Set brs2 = Nothing
          If cnt(3) >= cnt(0) Then
            ii = 3: cnt(0) = cnt(3)
          End If
          If cnt(6) >= cnt(0) Then
            ii = 6: cnt(0) = cnt(6)
          End If
          If cnt(7) >= cnt(0) Then
            ii = 7: cnt(0) = cnt(7)
          End If
          If cnt(0) = 0 Then GoTo rtn
          brs2.Open "select count(*) from sg,lo where sg_awno='" & awno1 & "' and sg_ptno like '" & Trim(mark) & "%' and sg_cran='" & CStr(ii) & "' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F'", rcn1, adOpenKeyset, adLockOptimistic
          If brs2(0) <= 30 Then rcn1.Execute "update st set st_err=5 where st_awno='" & awno1 & "' and st_stno = '" & Left(stno, STNO_NUM) & "'"
          Set brs2 = Nothing
          brs2.Open "select * from sg,lo where sg_awno='" & awno1 & "' and sg_ptno like '" & Trim(mark) & "%' and sg_cran='" & CStr(ii) & "' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
        End If
      Else
        If cran >= "1" And cran <= "9" Then
          brs2.Open "select * from sg,lo where sg_awno='" & awno1 & "' and sg_ptno like '" & Trim(mark) & "%' and sg_cran='" & cran & "' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
        Else
          brs2.Open "select * from sg,lo where sg_awno='" & awno1 & "' and sg_ptno like '" & Trim(mark) & "%' and sg_lono=lo_lono and lo_fbst='0' and lo_sgst='F' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
        End If
      End If
      lono = "        "
      Do
        If brs2.EOF Then Exit Do
        If get_string_value(awno1, brs2!sg_lono, "lo_sgst") = "F" And get_string_value(awno1, brs2!sg_lono, "lo_fbst") = "0" Then
          lono = brs2!sg_lono: Exit Do
        End If
        brs2.MoveNext
      Loop
      Set brs2 = Nothing
      If lono <> "        " Then
        rcn1.Execute "update lo set lo_sgst='R' where lo_awno='" & awno1 & "' and lo_lono='" & lono & "'"
        If Left(project_name, 2) <> "SK" And Left(mark, 6) <> "PALLET" Then
          Call put_cw(awno1, TR_NORMAL_OUT, lono, stno, loginname, "N")
        Else
          Call put_cw(awno1, TR_PALLETS_OUT, lono, stno, loginname, "N")
        End If
        rcn1.Execute "update st set st_err=0 where st_awno='" & awno1 & "' and st_stno = '" & Left(stno, STNO_NUM) & "'"
        pallet_auto_out = True
      Else
rtn:
        If get_integer_value1(awno1, Left(stno, STNO_NUM), "st_err") <> 6 Then rcn1.Execute "update st set st_err=6 where st_awno='" & awno1 & "' and st_stno = '" & Left(stno, STNO_NUM) & "'"
        Call syslog("pallet_auto_out", 0, 0, "沒棧板," & stno & "," & mark & "," & cran)
      End If
    End If
  End If
  Set brs3 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("pallet_auto_out")
End Function

Public Function pallet_auto_in(awno1 As String, stno As String, palletname As String, palletnum As Integer, zone As String, Optional cran As String) As Boolean
  Dim brs2 As New ADODB.Recordset, brs3 As New ADODB.Recordset, brs4 As New ADODB.Recordset, lono As String * 8
  Dim sitm_buf As String, flag As Boolean
  On Error GoTo err1_rtn
  pallet_auto_in = False
  brs3.Open "select * from cw where cw_from='" & stno & "'", rcn, adOpenKeyset, adLockOptimistic
  If Left(project_name, 2) = "SP" Or brs3.EOF = True Then
    If Not check_io_wrong(awno1, Left(stno, STNO_NUM), " ") Then
      lono = "        "
      If cran >= "1" And cran <= "9" Then
        flag = get_empty_lono(awno1, zone, lono, "pallet_auto_in", "no_show", cran)
      Else
        flag = get_empty_lono(awno1, zone, lono, "pallet_auto_in", "no_show")
      End If
      If flag = False Then
        rcn1.Execute "delete from sg where sg_awno='" & awno1 & "' and sg_lono='" & lono & "'"
        brs2.Open "select * from tr where tr_srid='K'", rcn, adOpenKeyset, adLockOptimistic
        'brs2.Open "select * from tr where tr_opno='2'", rcn, adOpenKeyset, adLockOptimistic
        Do
          If brs2.EOF Then
            rcn1.Execute "update lo set lo_sgst='0' where lo_awno='" & awno1 & "' and lo_lono='" & lono & "'"
            Set brs2 = Nothing: MsgBox ("pallet_auto_in:tr檔沒有tr_srid=K(新增資料)異常,無法轉出sg,請執行庫格資料查詢及維護作業-S360,新增一筆庫存資料後即可開始執行本作業"): Exit Function
          End If
          brs4.Open "select count(*) from tr where tr_awno='" & brs2!tr_awno & "' and tr_lono='" & brs2!tr_lono & "' and tr_ptno='" & brs2!tr_ptno & "' and tr_sitm='" & brs2!tr_sitm & "'", rcn1, adOpenKeyset, adLockOptimistic
          If brs4(0) = 1 Then
            Set brs4 = Nothing: Exit Do
          End If
          Set brs4 = Nothing
          brs2.MoveNext
        Loop
        rcn1.Execute "insert into sg select * from tr where tr_awno='" & brs2!tr_awno & "' and tr_lono='" & brs2!tr_lono & "' and tr_ptno='" & brs2!tr_ptno & "' and tr_sitm='" & brs2!tr_sitm & "'"
        sitm_buf = get_now()
        If Left(project_name, 4) = "SK02" Or Left(project_name, 4) = "SK04" Or Left(project_name, 4) = "SK05" Or Left(project_name, 2) = "SU" Then
          '注意對SK專案言,sg_ptno='" & Trim(palletname) &"'最多只能5位數
          rcn1.Execute "update sg set sg_mark='" & Trim(palletname) & "',sg_prod=' ',sg_ptno='" & Left(palletname, 2) & "',sg_spec=' ',sg_mchno=' ',sg_class1=' ',sg_leng=0,sg_grade=' ',sg_widt=0,sg_quty=" & palletnum & ",sg_awno='" & awno1 & "',sg_palt=0,sg_mveq=' ',sg_stat='0',sg_lono='" & lono & "',sg_from='" & stno & "',sg_to='" & lono & "',sg_now=' ',sg_nxt=' ',sg_err=0,sg_first='N',sg_sitm='" & sitm_buf & "',sg_sitm1='" & sitm_buf & "',sg_date='" & Left(sitm_buf, 10) & "',sg_opno='" & get_opno(TR_PALLETS_IN) & "',sg_srid='" & TR_PALLETS_IN & "',sg_srnm='" & get_srnm(TR_PALLETS_IN) & "',sg_name='" & Trim(loginname) & "',sg_pqty=0,sg_odno=' ',sg_ltno=' ' where sg_awno='" & brs2!tr_awno & "' and sg_lono='" & brs2!tr_lono & "' and sg_ptno='" & brs2!tr_ptno & "' and sg_sitm='" & brs2!tr_sitm & "'"
          'rcn1.Execute "update sg set sg_ptno='" & Trim(palletname) & "',sg_quty=" & palletnum & ",sg_awno='" & awno1 & "',sg_palt=0,sg_mveq=' ',sg_stat='0',sg_lono='" & lono & "',sg_from='" & stno & "',sg_to='" & lono & "',sg_now=' ',sg_nxt=' ',sg_err=0,sg_first='N',sg_sitm='" & sitm_buf & "',sg_sitm1='" & sitm_buf & "',sg_date='" & Left(sitm_buf, 10) & "',sg_opno='" & get_opno(TR_PALLETS_IN) & "',sg_srid='" & TR_PALLETS_IN & "',sg_srnm='" & get_srnm(TR_PALLETS_IN) & "',sg_name='" & Trim(loginname) & "',sg_pqty=0,sg_odno=' ',sg_ltno=' ' where sg_awno='" & brs2!tr_awno & "' and sg_lono='" & brs2!tr_lono & "' and sg_ptno='" & brs2!tr_ptno & "' and sg_sitm='" & brs2!tr_sitm & "'"
        Else
          rcn1.Execute "update sg set sg_cran='" & ((Val(Left(lono, 2)) + 1) \ 2) & "',sg_ptno='" & Trim(palletname) & "',sg_quty=" & palletnum & ",sg_awno='" & awno1 & "',sg_palt=0,sg_mveq=' ',sg_stat='0',sg_lono='" & lono & "',sg_from='" & stno & "',sg_to='" & lono & "',sg_now=' ',sg_nxt=' ',sg_err=0,sg_first='N',sg_sitm='" & sitm_buf & "',sg_sitm1='" & sitm_buf & "',sg_date='" & Left(sitm_buf, 10) & "',sg_opno='" & get_opno(TR_PALLETS_IN) & "',sg_srid='" & TR_PALLETS_IN & "',sg_srnm='" & get_srnm(TR_PALLETS_IN) & "',sg_name='" & Trim(loginname) & "',sg_pqty=0,sg_odno=' ',sg_ltno=' ' where sg_awno='" & brs2!tr_awno & "' and sg_lono='" & brs2!tr_lono & "' and sg_ptno='" & brs2!tr_ptno & "' and sg_sitm='" & brs2!tr_sitm & "'"
        End If
        Set brs2 = Nothing
        'Call put_sg_for_pallet(awno1, lono, palletname, palletnum)
        Call put_cw(awno1, TR_PALLETS_IN, stno, lono, loginname, "N")
        rcn1.Execute "update st set st_err=0 where st_awno='" & awno1 & "' and st_stno = '" & Left(stno, STNO_NUM) & "'"
        pallet_auto_in = True
      Else
        If get_integer_value1(awno1, Left(stno, STNO_NUM), "st_err") <> 7 Then rcn1.Execute "update st set st_err=7 where st_awno='" & awno1 & "' and st_stno = '" & Left(stno, STNO_NUM) & "'"
      End If
    End If
  End If
  Set brs3 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("pallet_auto_in")
End Function
Public Function pallet_auto_move(awno1 As String, from1 As String, to1 As String) As Boolean
  Dim brs2 As New ADODB.Recordset, brs3 As New ADODB.Recordset
  Dim sitm_buf As String, flag As Boolean
  On Error GoTo err1_rtn
  pallet_auto_move = False
  If get_string_value1(awno1, Left(to1, STNO_NUM), "st_load") = "1" Then Exit Function
  brs3.Open "select * from cw where cw_from like '" & Trim(from1) & "%'", rcn, adOpenKeyset, adLockOptimistic
  If brs3.EOF Then
    If Not check_io_wrong(awno1, Left(from1, STNO_NUM), " ") Then
      Call put_cw(awno1, TR_MOVE_STNO, from1, to1, loginname, "N")
      pallet_auto_move = True
    End If
  End If
  Set brs3 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("pallet_auto_move")
End Function


Public Sub version1_rtn()
  Dim brs2 As New ADODB.Recordset, ver As String * 13, ii As Integer, buf As String, awno1 As String * 1
  On Error GoTo err1_rtn
  Exit Sub
err1_rtn:
  Call err2_rtn("version1_rtn")
End Sub

Public Sub shell_rtn()
  Dim ii As Integer
  On Error GoTo err1_rtn
  Call Shell(App.Path & "\MAPATH連線程式.exe", 1)
  For ii = 1 To PLC1_NUM
    Call Shell(App.Path & "\PLC" & ii & "連線程式.exe", 1)
  Next ii
  For ii = 1 To CRAN_NUM
    Call Shell(App.Path & "\CRAN" & ii & "連線程式.exe", 1)
  Next ii
  Exit Sub
err1_rtn:
  Call err2_rtn("shell_rtn")
End Sub

Public Function increase_cnt(fieldname As String, mveq As String, now1 As String) As Integer
  Dim brs2 As New ADODB.Recordset, jj As Long
  On Error GoTo err1_rtn
  If Left(fieldname, 2) = "cr" Then
    brs2.Open "select cr_cnt,cr_erno from cr where cr_mveq='" & mveq & "'", rcn1, adOpenKeyset, adLockOptimistic
    If brs2.EOF = True Then
      Set brs2 = Nothing: Response = MsgBox("搬運設備" & mveq & "不存在異常"): End
    End If
    If Left(fieldname, 6) = "cr_cnt" Then
      jj = brs2!cr_cnt
      If jj >= 99999 Then jj = 1 Else jj = jj + 1
      rcn1.Execute "update cr set cr_cnt=" & jj & " where cr_mveq='" & mveq & "'"
    Else
      jj = brs2!cr_erno
      If jj >= 99999 Then jj = 1 Else jj = jj + 1
      rcn1.Execute "update cr set cr_erno=" & jj & " where cr_mveq='" & mveq & "'"
    End If
  Else
    brs2.Open "select pa_cnt from pa", rcn1, adOpenKeyset, adLockOptimistic
    If brs2.EOF = True Then
      Set brs2 = Nothing: Response = MsgBox("pa_cnt不存在異常"): End
    End If
    jj = brs2!pa_cnt
    If jj >= 99999 Then jj = 1 Else jj = jj + 1
    rcn1.Execute "update pa set pa_cnt=" & jj
  End If
  Set brs2 = Nothing: Exit Function
err1_rtn:
  Call err2_rtn("increase_cnt")
End Function

Public Function check_status(awno1 As String, mveq As String, now1 As String, nxt1 As String, Optional layer As String) As Integer
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  check_status = 0
  If Left(mveq, 1) = "#" Or Left(mveq, 3) = "STV" Then
    brs2.Open "select * from cr where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'", rcn1, adOpenKeyset, adLockOptimistic
    If brs2.EOF = True Then
      Set brs2 = Nothing: Response = MsgBox("搬運設備" & mveq & "不存在異常"): End
    End If
    If Left(project_name, 2) <> "SO" Then
      If Not (is_stno(awno1, Left(now1, STNO_NUM)) And get_string_value1(awno1, Left(now1, STNO_NUM), "st_stat") = "V") Then
      'If Left(project_name, 2) <> "SA" Then '棧板在車上
        If Left(mveq, 3) = "STV" And (layer = "1" Or layer = "2") Then '"1" mean 要走下層,"2" mean 要走上層
          If brs2!cr_load = layer Or brs2!cr_load = "3" Then check_status = 69
        Else
          If brs2!cr_load <> "0" Then check_status = 69
        End If
        If Left(project_name, 4) = "SU01" And (mveq = EQ_STV16 Or mveq = EQ_STV17) Then
          check_status = 0
        End If
      End If
      If brs2!cr_auto <> "1" Then check_status = 63
    Else
      If brs2!cr_auto <> "0" Then check_status = 63
    End If
    If brs2!cr_fbst <> FB_NONE Then check_status = 61
    If brs2!cr_mvst = MV_NOW Then check_status = 1
    '須Mark掉,否則若PLC沒異常卻送出異常碼,如此命令將無法下達,例Epoxy2系統STV8,PLC一直送err=3(意思為G24起站無載,但G24明明為虛擬站)
    'If brs2!cr_err <> 0 Then check_status = brs2!cr_err '希望命令能夠重覆下,所以此處不作此項判斷
    Set brs2 = Nothing
    If check_status > 0 Then Exit Function
  End If
  If is_stno(awno1, now1) Then
    brs2.Open "select * from st where st_awno='" & awno1 & "' and st_stno='" & Left(now1, STNO_NUM) & "'", rcn1, adOpenKeyset, adLockOptimistic
    If brs2.EOF = True Then
      Set brs2 = Nothing: Response = MsgBox("站號" & now1 & "不存在異常"): End
    End If
    If brs2!st_fbst <> FB_NONE Then
      check_status = 62: Set brs2 = Nothing: Exit Function
    End If
    Set brs2 = Nothing
  End If
  If is_stno(awno1, nxt1) Then
    brs2.Open "select * from st where st_awno='" & awno1 & "' and st_stno='" & Left(nxt1, STNO_NUM) & "'", rcn1, adOpenKeyset, adLockOptimistic
    If brs2.EOF = True Then
      Set brs2 = Nothing: Response = MsgBox("站號" & nxt1 & "不存在異常1"): End
    End If
    If brs2!st_fbst <> FB_NONE Then check_status = 62
    Set brs2 = Nothing
  End If
  Exit Function
err1_rtn:
  Call err2_rtn("check_status")
End Function

Public Function get_empty_lono(awno1 As String, zone As String, lono As String, srid As String, Optional show1 As String, Optional cran As String, Optional bay As String) As Boolean
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  get_empty_lono = True
  If Not is_lono(awno1, lono) Then
    If (Left(project_name, 4) >= "SK01" And Left(project_name, 4) <= "SK03") And awno1 = "5" And (Left(cran, 2) = "01" Or Left(cran, 2) = "02") Then
      brs2.Open "select lo_lono from lo where lo_awno='" & awno1 & "' and lo_zone='" & zone & "' and lo_sgst='0' and lo_fbst='0'  and substring(lo_lono,1,2) = '" & Left(cran, 2) & "' order by lo_sisq", rcn1, adOpenKeyset, adLockOptimistic
    ElseIf cran = "PAPER" Then
      brs2.Open "select lo_lono from lo where lo_awno='" & awno1 & "' and lo_zone='" & zone & "' and lo_sgst='0' and lo_fbst='0' and substring(lo_lono,6,3) = '003' order by lo_sisq", rcn1, adOpenKeyset, adLockOptimistic
    ElseIf cran >= "1" And cran <= "9" Then
      If Len(Trim(bay)) = 6 Then
        brs2.Open "select lo_lono from lo where lo_awno='" & awno1 & "' and lo_cran='" & cran & "' and lo_zone='" & zone & "' and lo_sgst='0' and lo_fbst='0' and substring(lo_lono,3,3) >= '" & Mid(bay, 1, 3) & "' and substring(lo_lono,3,3) <= '" & Mid(bay, 4, 3) & "' order by lo_sisq", rcn1, adOpenKeyset, adLockOptimistic
      Else
        brs2.Open "select lo_lono from lo where lo_awno='" & awno1 & "' and lo_cran='" & cran & "' and lo_zone='" & zone & "' and lo_sgst='0' and lo_fbst='0' order by lo_sisq", rcn1, adOpenKeyset, adLockOptimistic
      End If
    Else
      If Len(Trim(bay)) = 6 Then
        brs2.Open "select lo_lono from lo where lo_awno='" & awno1 & "' and lo_zone='" & zone & "' and lo_sgst='0' and lo_fbst='0' and substring(lo_lono,3,3) >= '" & Mid(bay, 1, 3) & "' and substring(lo_lono,3,3) <= '" & Mid(bay, 4, 3) & "'  order by lo_sisq", rcn1, adOpenKeyset, adLockOptimistic
      Else
        brs2.Open "select lo_lono from lo where lo_awno='" & awno1 & "' and lo_zone='" & zone & "' and lo_sgst='0' and lo_fbst='0' order by lo_sisq", rcn1, adOpenKeyset, adLockOptimistic
      End If
    End If
    If brs2.EOF Then
      Set brs2 = Nothing
      If Left(prog, 3) <> "PLC" And Left(prog, 3) <> "STV" And Left(prog, 4) <> "CRAN" And Left(prog, 6) <> "MAPATH" And Left(prog, 6) <> "RS6000" And Left(prog, 4) <> "TTCS" And Left(prog, 4) <> "MONO" Then
      'If Left(project_name, 2) <> "SJ" And srid <> "pallet_auto_in" Then
        If show1 <> "no_show" Then
          Response = MsgBox(get_awnm(awno1) & get_zone(awno1, Val(zone) + 9) & "存區找不到空庫格!")
        Else
          GoTo rtn1
        End If
      Else
rtn1:
        Call syslog("", 0, 0, get_awnm(awno1) & get_zone(awno1, Val(zone) + 9) & "存區找不到空庫格!")
      End If
      Exit Function
    End If
    lono = brs2(0): Set brs2 = Nothing
  End If
  rcn1.Execute "update lo set lo_sgst='S' where lo_awno='" & awno1 & "' and lo_lono='" & lono & "'"
  If srid <> "pallet_auto_in" And srid <> "bar" Then
    If Left(project_name, 2) = "SG" And awno1 = "3" Then
      If zone = "1" Then Response = MsgBox("選取到單層庫格,庫格編號=" & lono & "確定要執行" & srid, vbOKCancel) Else Response = MsgBox("選取到雙層庫格,庫格編號=" & lono & "確定要執行" & srid, vbOKCancel)
    ElseIf Left(project_name, 2) = "SF" And awno1 = "2" Then
      If zone = "1" Then Response = MsgBox("選取到小庫格,庫格編號=" & lono & "確定要執行" & srid, vbOKCancel) Else Response = MsgBox("選取到大庫格,庫格編號=" & lono & "確定要執行" & srid, vbOKCancel)
    Else
      If Left(prog, 3) <> "PLC" And Left(prog, 3) <> "STV" And Left(prog, 4) <> "CRAN" And Left(prog, 6) <> "MAPATH" And Left(prog, 6) <> "RS6000" And Left(prog, 4) <> "TTCS" And Left(prog, 4) <> "MONO" Then
        If show1 <> "no_show" Then Response = MsgBox("庫格編號=" & lono & "確定要執行" & srid, vbOKCancel)
      End If
    End If
    If Response = vbCancel Then
      rcn1.Execute "update lo set lo_sgst='0' where lo_awno='" & awno1 & "' and lo_lono='" & lono & "'"
      lono = "Cancel": Exit Function
    End If
  End If
  get_empty_lono = False
  Exit Function
err1_rtn:
  Call err2_rtn("get_empty_lono")
End Function

'mode="S" means string,mode="I" means integer
'eg: zonename=get_field_value(awno,"pt_ptno","S","PALLET","pt_zonename","S")
Public Function get_field_value(awno1 As String, keyfield As String, keymode As String, keyvalue As String, getfield As String, getmode As String) As String
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  If (keymode <> "S" And keymode <> "I") Or (getmode <> "S" And getmode <> "I") Then
    MsgBox ("mode不等於S或I異常 for get_field_value function"): get_field_value = "Err1": Exit Function
  End If
  If keymode = "S" Then
    If LCase(Left(getfield, 2)) = "st" And Len(Trim(keyvalue)) = 2 Then 'Fa,FA,fa case,SQL讀取時會視為相同,所以需再判斷
      brs2.Open "select " & LCase(getfield) & ",st_stno from " & LCase(Left(getfield, 2)) & " where " & LCase(Left(getfield, 2)) & "_awno='" & awno1 & "' and " & keyfield & "='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
      Do
        If brs2.EOF Then Exit Do
        If brs2!st_stno = Left(keyvalue, STNO_NUM) Then Exit Do
        brs2.MoveNext
      Loop
    ElseIf LCase(Left(getfield, 2)) <> "pt" Then
      brs2.Open "select " & LCase(getfield) & " from " & LCase(Left(getfield, 2)) & " where " & LCase(Left(getfield, 2)) & "_awno='" & awno1 & "' and " & keyfield & "='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
    Else
      brs2.Open "select " & LCase(getfield) & " from " & LCase(Left(getfield, 2)) & " where " & keyfield & "='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
    End If
  Else
    If LCase(Left(getfield, 2)) <> "pt" Then
      brs2.Open "select " & LCase(getfield) & " from " & LCase(Left(getfield, 2)) & " where " & LCase(Left(getfield, 2)) & "_awno='" & awno1 & "' and " & keyfield & "=" & Val(keyvalue), rcn1, adOpenKeyset, adLockOptimistic
    Else
      brs2.Open "select " & LCase(getfield) & " from " & LCase(Left(getfield, 2)) & " where " & keyfield & "=" & Val(keyvalue), rcn1, adOpenKeyset, adLockOptimistic
    End If
  End If
  If brs2.EOF Then
    If Left(project_name, 2) = "SJ" And LCase(Left(getfield, 6)) = "sg_bar" Then
      get_field_value = ""
    Else
      If Left(project_name, 2) = "SK" And LCase(Left(getfield, 2)) = "cw" Then
         Call syslog1("get_field_value", 0, 567, "awno=" & awno1 & ",讀取" & LCase(Left(getfield, 2)) & "檔,讀取欄位" & LCase(getfield) & "資料不存在錯誤!(使用鍵值欄位" & LCase(keyfield) & ",鍵值為" & keyvalue & ")")
      Else
         MsgBox ("awno=" & awno1 & ",讀取" & LCase(Left(getfield, 2)) & "檔,讀取欄位" & LCase(getfield) & "資料不存在錯誤!(使用鍵值欄位" & LCase(keyfield) & ",鍵值為" & keyvalue & ")")
      End If
      get_field_value = "Err2"
    End If
  Else
    If getmode = "S" Then get_field_value = brs2(0) Else get_field_value = CStr(brs2(0))
  End If
  Set brs2 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("get_field_value" & ",awno1=" & awno1 & "," & "keyfield=" & keyfield & "," & "getfield=" & getfield, "noMsgBox")
  'Call err2_rtn("get_field_value")
End Function
'mode="S" means string,mode="I" means integer
'eg: err1=get_field_value_integer(awno,"cw_palt","I",123,"cw_err","I")
Public Function get_field_value_integer(awno1 As String, keyfield As String, keymode As String, keyvalue As String, getfield As String, getmode As String) As Integer
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  If (keymode <> "S" And keymode <> "I") Or (getmode <> "S" And getmode <> "I") Then
    MsgBox ("mode不等於S或I異常 for get_field_value_integer function"): get_field_value_integer = 32760: Exit Function
  End If
  If keymode = "S" Then
    If LCase(Left(getfield, 2)) = "st" And Len(Trim(keyvalue)) = 2 Then 'Fa,FA,fa case,SQL讀取時會視為相同,所以需再判斷
      brs2.Open "select " & LCase(getfield) & ",st_stno from " & LCase(Left(getfield, 2)) & " where " & LCase(Left(getfield, 2)) & "_awno='" & awno1 & "' and " & keyfield & "='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
      Do
        If brs2.EOF Then Exit Do
        If brs2!st_stno = Left(keyvalue, STNO_NUM) Then Exit Do
        brs2.MoveNext
      Loop
    ElseIf LCase(Left(getfield, 2)) <> "pt" Then
      brs2.Open "select " & LCase(getfield) & " from " & LCase(Left(getfield, 2)) & " where " & LCase(Left(getfield, 2)) & "_awno='" & awno1 & "' and " & keyfield & "='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
    Else
      brs2.Open "select " & LCase(getfield) & " from " & LCase(Left(getfield, 2)) & " where " & keyfield & "='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
    End If
  Else
    If LCase(Left(getfield, 2)) <> "pt" Then
      brs2.Open "select " & LCase(getfield) & " from " & LCase(Left(getfield, 2)) & " where " & LCase(Left(getfield, 2)) & "_awno='" & awno1 & "' and " & keyfield & "=" & Val(keyvalue), rcn1, adOpenKeyset, adLockOptimistic
    Else
      brs2.Open "select " & LCase(getfield) & " from " & LCase(Left(getfield, 2)) & " where " & keyfield & "=" & Val(keyvalue), rcn1, adOpenKeyset, adLockOptimistic
    End If
  End If
  If brs2.EOF Then
    If Left(project_name, 2) = "SK" And LCase(Left(getfield, 2)) = "cw" Then
       Call syslog1("get_field_value", 0, 567, "awno=" & awno1 & ",讀取" & LCase(Left(getfield, 2)) & "檔,讀取欄位" & LCase(getfield) & "資料不存在錯誤!(使用鍵值欄位" & LCase(keyfield) & ",鍵值為" & keyvalue & ")")
    Else
       MsgBox ("awno=" & awno1 & ",讀取" & LCase(Left(getfield, 2)) & "檔,讀取欄位" & LCase(getfield) & "資料不存在錯誤!(使用鍵值欄位" & LCase(keyfield) & ",鍵值為" & keyvalue & ")")
    End If
    get_field_value_integer = 32761: Exit Function
  Else
    If getmode = "S" Then get_field_value_integer = Val(brs2(0)) Else get_field_value_integer = brs2(0)
  End If
  Set brs2 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("get_field_value_integer" & ",awno1=" & awno1 & "," & "keyfield=" & keyfield & "," & "getfield=" & getfield, "noMsgBox")
  'Call err2_rtn("get_field_value_integer")
End Function

Public Function get_integer_value(awno1 As String, keyvalue As String, fieldname As String, Optional mode As Integer) As Integer
  Dim bcn2 As New ADODB.Connection, brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  bcn2.ConnectionTimeout = 3: bcn2.CursorLocation = adUseClient:  bcn2.Open rcn_str
  If LCase(Left(fieldname, 2)) = "st" Then
    If Len(Trim(keyvalue)) = 2 Then 'Fa,FA,fa case,SQL讀取時會視為相同,所以需再判斷
      brs2.Open "select " & LCase(fieldname) & ",st_stno from st where st_awno='" & awno1 & "' and st_stno='" & keyvalue & "'", bcn2, adOpenKeyset, adLockOptimistic
      Do
        If brs2.EOF Then Exit Do
        If brs2!st_stno = Left(keyvalue, STNO_NUM) Then Exit Do
        brs2.MoveNext
      Loop
    Else
      brs2.Open "select " & LCase(fieldname) & " from st where st_awno='" & awno1 & "' and st_stno='" & keyvalue & "'", bcn2, adOpenKeyset, adLockOptimistic
    End If
  ElseIf LCase(Left(fieldname, 2)) = "cr" Then
    brs2.Open "select " & LCase(fieldname) & " from cr where cr_awno='" & awno1 & "' and cr_mveq='" & keyvalue & "'", bcn2, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "cw" Then
    brs2.Open "select " & LCase(fieldname) & " from cw where cw_awno='" & awno1 & "' and cw_palt=" & Val(keyvalue), bcn2, adOpenKeyset, adLockOptimistic
  End If
  If brs2.EOF Then
    If LCase(Left(fieldname, 2)) = "st" Then
      brs2.Close: bcn2.Close: MsgBox ("站號不存在錯誤get_integer_value!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "cr" Then
      brs2.Close: bcn2.Close: MsgBox ("設備代號不存在錯誤get_integer_value!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    Else
      brs2.Close: bcn2.Close: MsgBox ("工作序號不存在錯誤get_integer_value!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    End If
  Else
    get_integer_value = brs2(0)
  End If
  brs2.Close: bcn2.Close
  Exit Function
err1_rtn:
  Call err2_rtn("get_integer_value" & ",awno1=" & awno1 & "," & "fieldname=" & fieldname, "noMsgBox")
  'Call err2_rtn("get_integer_value")
End Function
Public Function get_integer_value1(awno1 As String, keyvalue As String, fieldname As String, Optional mode As Integer) As Integer
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  If LCase(Left(fieldname, 2)) = "st" Then
    If Len(Trim(keyvalue)) = 2 Then 'Fa,FA,fa case,SQL讀取時會視為相同,所以需再判斷
      brs2.Open "select " & LCase(fieldname) & ",st_stno from st where st_awno='" & awno1 & "' and st_stno='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
      Do
        If brs2.EOF Then Exit Do
        If brs2!st_stno = Left(keyvalue, STNO_NUM) Then Exit Do
        brs2.MoveNext
      Loop
    Else
      brs2.Open "select " & LCase(fieldname) & " from st where st_awno='" & awno1 & "' and st_stno='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
    End If
  ElseIf LCase(Left(fieldname, 2)) = "cr" Then
    brs2.Open "select " & LCase(fieldname) & " from cr where cr_awno='" & awno1 & "' and cr_mveq='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "cw" Then
    brs2.Open "select " & LCase(fieldname) & " from cw where cw_awno='" & awno1 & "' and cw_palt=" & Val(keyvalue), rcn1, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "pa" Then
    brs2.Open "select " & LCase(fieldname) & " from pa", rcn1, adOpenKeyset, adLockOptimistic
  End If
  If brs2.EOF Then
    If LCase(Left(fieldname, 2)) = "st" Then
      Set brs2 = Nothing: MsgBox ("站號不存在錯誤get_integer_value1!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "cr" Then
      Set brs2 = Nothing: MsgBox ("設備代號不存在錯誤get_integer_value1!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "pa" Then
      Set brs2 = Nothing: MsgBox ("pa檔不存在錯誤get_integer_value1!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    Else
      Set brs2 = Nothing: MsgBox ("工作序號不存在錯誤get_integer_value1!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    End If
  Else
    get_integer_value1 = brs2(0)
  End If
  Set brs2 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("get_integer_value1" & ",awno1=" & awno1 & "," & "fieldname=" & fieldname, "noMsgBox")
  'Call err2_rtn("get_integer_value1")
End Function
'針對sg檔,以Do Loop搜尋某項要出庫ptno的sg_lono時,必須呼叫此,其它情況一律呼叫get_string_value1
Public Function get_string_value(awno1 As String, keyvalue As String, fieldname As String, Optional mode As Integer) As String
  Dim bcn2 As New ADODB.Connection, brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  bcn2.ConnectionTimeout = 3: bcn2.CursorLocation = adUseClient: bcn2.Open rcn_str
  If LCase(Left(fieldname, 2)) = "lo" Then
    brs2.Open "select " & LCase(fieldname) & " from lo where lo_awno='" & awno1 & "' and lo_lono='" & keyvalue & "'", bcn2, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "cr" Then
    brs2.Open "select " & LCase(fieldname) & " from cr where cr_awno='" & awno1 & "' and cr_mveq='" & keyvalue & "'", bcn2, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "st" Then
    If Len(Trim(keyvalue)) = 2 Then 'Fa,FA,fa case,SQL讀取時會視為相同,所以需再判斷
      brs2.Open "select " & LCase(fieldname) & ",st_stno from st where st_awno='" & awno1 & "' and st_stno='" & keyvalue & "'", bcn2, adOpenKeyset, adLockOptimistic
      Do
        If brs2.EOF Then Exit Do
        If Left(Trim(brs2!st_stno), Len(Trim(keyvalue))) = Left(Trim(keyvalue), Len(Trim(keyvalue))) Then Exit Do
        brs2.MoveNext
      Loop
    Else
      brs2.Open "select " & LCase(fieldname) & " from st where st_awno='" & awno1 & "' and st_stno='" & keyvalue & "'", bcn2, adOpenKeyset, adLockOptimistic
    End If
  ElseIf LCase(Left(fieldname, 2)) = "pa" Then
    brs2.Open "select " & LCase(fieldname) & " from pa", bcn2, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "pd" Then
    brs2.Open "select " & LCase(fieldname) & " from pd where pd_user='" & loginuser & "'", bcn2, adOpenKeyset, adLockOptimistic
  End If
  If brs2.EOF Then
    If LCase(Left(fieldname, 2)) = "lo" Then
      brs2.Close: bcn2.Close: MsgBox ("庫格編號不存在錯誤get_string_value!!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "cr" Then
      brs2.Close: bcn2.Close: MsgBox ("設備代號不存在錯誤get_string_valu!!!e,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "st" Then
      brs2.Close: bcn2.Close
      If LCase(Left(fieldname, 7)) = "st_stat" Then '判斷是否虛擬站號
        get_string_value = "R"
      Else
        MsgBox ("站號不存在錯誤get_string_value!"):
      End If
      Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "pa" Then
      brs2.Close: bcn2.Close: MsgBox ("pa檔不存在錯誤get_string_value!!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "pd" Then
'      brs2.Close: bcn2.Close: MsgBox ("人員編號不存在錯誤!!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    End If
  Else
    get_string_value = brs2(0)
  End If
  brs2.Close: bcn2.Close
  Exit Function
err1_rtn:
  'If Left(project_name, 4) = "SU01" Then
  Call err2_rtn("get_string_value" & ",awno1=" & awno1 & "," & "keyvalue=" & keyvalue & "," & "fieldname=" & fieldname, "noMsgBox")
  'Else
  '  Call err2_rtn("get_string_value")
  'End If
End Function

Public Function get_string_value1(awno1 As String, keyvalue As String, fieldname As String, Optional mode As Integer) As String
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  If LCase(Left(fieldname, 2)) = "lo" Then
    brs2.Open "select " & LCase(fieldname) & " from lo where lo_awno='" & awno1 & "' and lo_lono='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "cr" Then
    brs2.Open "select " & LCase(fieldname) & " from cr where cr_awno='" & awno1 & "' and cr_mveq='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "st" Then
    If Len(Trim(keyvalue)) = 2 Then 'Fa,FA,fa case,SQL讀取時會視為相同,所以需再判斷
      brs2.Open "select " & LCase(fieldname) & ",st_stno from st where st_awno='" & awno1 & "' and st_stno='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
      Do
        If brs2.EOF Then Exit Do
        If Left(Trim(brs2!st_stno), Len(Trim(keyvalue))) = Left(Trim(keyvalue), Len(Trim(keyvalue))) Then Exit Do
        brs2.MoveNext
      Loop
    Else
      brs2.Open "select " & LCase(fieldname) & " from st where st_awno='" & awno1 & "' and st_stno='" & keyvalue & "'", rcn1, adOpenKeyset, adLockOptimistic
    End If
  ElseIf LCase(Left(fieldname, 2)) = "pa" Then
    brs2.Open "select " & LCase(fieldname) & " from pa", rcn1, adOpenKeyset, adLockOptimistic
  ElseIf LCase(Left(fieldname, 2)) = "pd" Then
    brs2.Open "select " & LCase(fieldname) & " from pd where pd_user='" & loginuser & "'", rcn1, adOpenKeyset, adLockOptimistic
  End If
  If brs2.EOF Then
    If LCase(Left(fieldname, 2)) = "lo" Then
      Set brs2 = Nothing: MsgBox ("庫格編號不存在錯誤get_string_value1!!!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "cr" Then
      Set brs2 = Nothing: MsgBox ("設備代號不存在錯誤get_string_value1!!!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "st" Then
      Set brs2 = Nothing
      If LCase(Left(fieldname, 7)) = "st_stat" Then '判斷是否虛擬站號
        get_string_value1 = "R"
      Else
        MsgBox ("站號不存在錯誤get_string_value1!"):
      End If
      Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "pa" Then
      Set brs2 = Nothing: MsgBox ("pa檔不存在錯誤get_string_value1!!!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    ElseIf LCase(Left(fieldname, 2)) = "pd" Then
'      Set brs2 = Nothing: MsgBox ("人員編號不存在錯誤!!!!,awno1=" & awno1 & ",keyvalue=" & keyvalue & ",fieldname=" & fieldname & ",mode=" & mode): Exit Function
    End If
  Else
    get_string_value1 = brs2(0)
  End If
  Set brs2 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("get_string_value1" & ",awno1=" & awno1 & "," & "keyvalue=" & keyvalue & "," & "fieldname=" & fieldname, "noMsgBox")
  'Call err2_rtn("get_string_value1")
End Function

Public Function get_dt_valu(awno1 As String, clss As Integer) As Integer
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  brs2.Open "select dt_valu from dt where dt_awno='" & awno1 & "' and dt_clss=" & clss, rcn1, adOpenKeyset, adLockOptimistic
  get_dt_valu = brs2(0)
  Set brs2 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("get_dt_valu")
End Function


Public Function check_io_wrong(awno1 As String, from1 As String, to1 As String, Optional prog1 As String) As Boolean
  Dim brs2 As New ADODB.Recordset, buf As String
  On Error GoTo err1_rtn
  If Left(from1, 1) <> " " Then
    If is_lono(awno1, from1) Then
      brs2.Open "select * from lo where lo_awno='" & awno1 & "' and lo_lono='" & from1 & "'", rcn1, adOpenKeyset, adLockOptimistic
      If brs2!lo_sgst <> SG_STOR Then
        If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
          Response = MsgBox("此庫格編號=" & from1 & "目前非在庫中,是否仍要執行出庫作業", vbOKCancel)
          If Response = vbCancel Then
            Set brs2 = Nothing: check_io_wrong = True: Exit Function
          End If
        Else
          Set brs2 = Nothing: check_io_wrong = True: Exit Function
        End If
      ElseIf brs2!lo_fbst = FB_ALL Then
        Set brs2 = Nothing
        If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
          Response = MsgBox("此庫格編號=" & from1 & "目前禁用中!", , "警告訊息")
        End If
        check_io_wrong = True: Exit Function
      End If
      Set brs2 = Nothing
    ElseIf is_stno(awno1, from1) Then
      brs2.Open "select * from st where st_awno='" & awno1 & "' and st_stno='" & from1 & "'", rcn1, adOpenKeyset, adLockOptimistic
      If brs2!st_fbst = FB_ALL Then
        Set brs2 = Nothing
        If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
          Response = MsgBox("此站號=" & from1 & "目前禁用中!", , "警告訊息")
        Else
          'pallet_auto_in副程式使用
          If get_integer_value1(awno1, Left(from1, STNO_NUM), "st_err") <> 62 Then rcn1.Execute "update st set st_err=62 where st_awno='" & awno1 & "' and st_stno = '" & Left(from1, STNO_NUM) & "'"
        End If
        check_io_wrong = True: Exit Function
      ElseIf brs2!st_load = NOLD Then
        If Left(prog1, 3) = "S25" Then
          Response = MsgBox("此站號=" & from1 & "目前無載,是否仍要執行入庫作業!", vbOKCancel, "警告訊息")
          If Response = vbCancel Then
            Set brs2 = Nothing: check_io_wrong = True: Exit Function
          End If
        ElseIf brs2!st_stat <> "V" Then
          'buf = get_pfile("bypass_st")
          'If Left(buf, 1) <> "Y" Then
            Set brs2 = Nothing
            If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
              Response = MsgBox("此站號=" & from1 & "目前無載!", , "警告訊息")
            End If
            check_io_wrong = True: Exit Function
          'End If
        End If
      End If
      Set brs2 = Nothing
      If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
        brs2.Open "select * from cw where cw_nxt='" & from1 & "'", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs2.EOF Then
          Response = MsgBox("工作檔中存在一筆下一站為" & from1 & "的出庫作業(工作序號為'" & brs2!cw_palt & "'),是否仍要執行入庫作業", vbOKCancel)
          If Response = vbCancel Then
            Set brs2 = Nothing: check_io_wrong = True: Exit Function
          End If
        End If
        Set brs2 = Nothing
        brs2.Open "select * from cw where cw_now='" & from1 & "'", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs2.EOF Then
          If Left(prog1, 3) = "S25" Then
            Response = MsgBox("工作檔中存在一筆目前站為" & from1 & "的入庫作業(工作序號為'" & brs2!cw_palt & "'),是否仍要執行入庫作業", vbOKCancel, "警告訊息")
            If Response = vbCancel Then
              Set brs2 = Nothing: check_io_wrong = True: Exit Function
            End If
          Else
            'Response = MsgBox("工作檔中存在一筆目前站為" & from1 & "的入庫作業(工作序號為'" & brs2!cw_palt & "')")
            'Set brs2 = Nothing: check_io_wrong = True: Exit Function
          End If
        End If
        Set brs2 = Nothing
      End If
    End If
  End If
  If Left(to1, 1) <> " " Then
    If is_lono(awno1, to1) Then
      brs2.Open "select * from lo where lo_awno='" & awno1 & "' and lo_lono='" & to1 & "'", rcn1, adOpenKeyset, adLockOptimistic
      If brs2!lo_sgst <> SG_EMPTY Then
        If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
          Response = MsgBox("此庫格編號=" & to1 & "目前非空庫,是否仍要執行入庫作業", vbOKCancel)
          If Response = vbCancel Then
            Set brs2 = Nothing: check_io_wrong = True: Exit Function
          End If
        Else
          Set brs2 = Nothing: check_io_wrong = True: Exit Function
        End If
      ElseIf brs2!lo_fbst = FB_ALL Then
        Set brs2 = Nothing
        If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
          Response = MsgBox("此庫格編號=" & to1 & "目前禁用中!", , "警告訊息")
        End If
        check_io_wrong = True: Exit Function
      End If
      Set brs2 = Nothing
    ElseIf is_stno(awno1, to1) Then
      brs2.Open "select * from st where st_awno='" & awno1 & "' and st_stno='" & to1 & "'", rcn1, adOpenKeyset, adLockOptimistic
      If brs2!st_fbst = FB_ALL Then
        Set brs2 = Nothing
        If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
          Response = MsgBox("此站號=" & to1 & "目前禁用中!", , "警告訊息")
        Else
          'pallet_auto_out副程式使用
          If get_integer_value1(awno1, Left(to1, STNO_NUM), "st_err") <> 62 Then rcn1.Execute "update st set st_err=62 where st_awno='" & awno1 & "' and st_stno = '" & Left(to1, STNO_NUM) & "'"
        End If
        check_io_wrong = True: Exit Function
      End If
      Set brs2 = Nothing
      If Not (Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS") Then
        brs2.Open "select * from cw where cw_from='" & to1 & "'", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs2.EOF Then
          Response = MsgBox("工作檔中存在一筆起始站為" & to1 & "的入庫作業(工作序號為'" & brs2!cw_palt & "'),是否仍要執行出庫作業", vbOKCancel)
          If Response = vbCancel Then
            Set brs2 = Nothing: check_io_wrong = True: Exit Function
          End If
        End If
      End If
      Set brs2 = Nothing
    End If
  End If
  check_io_wrong = False
  Exit Function
err1_rtn:
  Call err2_rtn("check_io_wrong")
End Function

Public Function get_lo_lono(lono As ComboBox, awno1 As String) As ComboBox
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  lono.Clear
  brs2.Open "select lo_lono from lo where lo_awno='" & awno1 & "' and (lo_lono like '0%' or lo_lono like '1%') order by lo_lono", rcn1, adOpenKeyset, adLockOptimistic
  Do
    If brs2.EOF Then Exit Do
    lono.AddItem brs2(0): brs2.MoveNext
  Loop
  Set brs2 = Nothing: Set get_lo_lono = lono
  Exit Function
err1_rtn:
  Call err2_rtn("get_lo_lono")
End Function

Public Function get_st_stno(stno As ComboBox, awno1 As String, opno As String) As ComboBox
  Dim brs2 As New ADODB.Recordset, ii As Integer, jj As Integer
  On Error GoTo err1_rtn
  stno.Clear
  brs2.Open "select * from st where st_awno='" & awno1 & "' order by st_stno", rcn1, adOpenKeyset, adLockOptimistic
  Do
    If brs2.EOF Then Exit Do
    If brs2!st_opno = opno Or brs2!st_opno = "A" Or opno = "A" Then
      If Left(project_name, 2) = "SP" Then
        '             10.110.71.57-60: 後段電腦
        '             10.110.71.61-65: 前段電腦
        If Mid(computername, 1, 12) >= "MLNPCPL4Q057" And Mid(computername, 1, 12) <= "MLNPCPL4Q060" Then
          If Left(brs2!st_stno, 1) < "D" Then GoTo rtn
        End If
        If Mid(computername, 1, 12) >= "MLNPCPL4Q061" And Mid(computername, 1, 12) <= "MLNPCPL4Q065" Then
          If Left(brs2!st_stno, 1) >= "D" Then GoTo rtn
        End If
      End If
      stno.AddItem brs2!st_stno
rtn:
'      If Left(brs2!st_stno, STNO_NUM) = Mid(computername, 5, STNO_NUM) Then stno = brs2!st_stno
    End If
    brs2.MoveNext
  Loop
  Set brs2 = Nothing
  If Left(project_name, 4) = "SK01" And awno1 = "4" And opno = "S" Then
    For ii = 21 To 37
      For jj = 1 To 24
        stno.AddItem "A" & (Format(ii, "00")) & (Format(jj, "00"))
      Next jj
    Next ii
  End If
  If Left(project_name, 2) >= "SK04" And awno1 = "5" Then
    stno.AddItem "CCL"
    'If opno = "S" Then
      stno.AddItem "F01": stno.AddItem "F11": stno.AddItem "F19": stno.AddItem "F20"
      stno.AddItem "F12": stno.AddItem "F10": stno.AddItem "F02"
    'End If
  End If
  Set get_st_stno = stno
  If Left(project_name, 2) = "SK" Then get_st_stno = get_st_text(opno) Else get_st_stno = ""
  Exit Function
err1_rtn:
  Call err2_rtn("get_st_stno")
End Function

Public Function combo_choose_wrong(combo As ComboBox, caption1 As String, Optional prog1 As String) As Boolean
  Dim i As Integer
  If combo.ListCount = 0 Then
    Response = MsgBox(caption1 & "-欄位沒有東西可以選擇錯誤!"): Screen.MousePointer = 0: combo_choose_wrong = True
    Exit Function
  End If
  For i = 0 To combo.ListCount - 1
    If Trim(combo) = Trim(combo.List(i)) Then Exit For
  Next i
  If i = combo.ListCount Then
    If Left(prog1, 3) = "S25" Then
      Response = MsgBox(caption1 & "-欄位選擇錯誤,是否仍要執行物流搬運作業!", vbOKCancel, "警告訊息")
      If Response = vbCancel Then
        Screen.MousePointer = 0: combo_choose_wrong = True: Exit Function
      End If
    Else
      Response = MsgBox(caption1 & "-欄位選擇錯誤,請重新選擇!")
      Screen.MousePointer = 0: combo_choose_wrong = True: Exit Function
    End If
  End If
  combo_choose_wrong = False
End Function

Public Function get_cw_stno(awno1 As String, now1 As String, nxt1 As String) As String
  '由cw_now,cw_nxt決定cw_stno[STNO_NUM]
  get_cw_stno = "000"
  If is_lono(awno1, now1) And is_lono(awno1, nxt1) Then '針對搬倉作業
    get_cw_stno = Left("00R" & "     ", STNO_NUM)
  ElseIf (Not is_lono(awno1, now1)) And Left(now1, 1) <> "(" Then
    get_cw_stno = Left(now1, STNO_NUM)
  Else
    get_cw_stno = Left(nxt1, STNO_NUM)
  End If
End Function

Public Function convert_sitm(sitm As String) As String
  convert_sitm = (Val(Mid(sitm, 1, 4)) - 1911) & Mid(sitm, 6, 2) & Mid(sitm, 9, 2) & Mid(sitm, 12, 2) & Mid(sitm, 15, 2) & Mid(sitm, 18, 2)
End Function
Public Function get_local_now() As String
  Dim buf As String
  On Error GoTo err1_rtn
  buf = Date '2007/5/25
  buf = time '下午 01:02:03
  If Left(buf, 2) = "下午" Then
    get_local_now = Left(get_now(), 11) & (Val(Mid(buf, 4, 2)) + 12) & Mid(buf, 6, 6)
  Else
    get_local_now = Left(get_now(), 11) & Mid(buf, 4, 8)
  End If
  Exit Function
err1_rtn:
  Call err2_rtn("get_local_now")
End Function
Public Function get_now(Optional mode1 As Integer) As String '取得系統時間(西元)
  Dim brs2 As New ADODB.Recordset, qsql As String, lib_buf As String * 100, ii As Integer
  On Error GoTo err1_rtn
  'get_now = "2000/09/25 12:20:30": Exit Function
  If Left(DB_TYPE, 3) = "SQL" Then
    brs2.Open "select getdate()", rcn1, adOpenKeyset, adLockOptimistic
    get_now = Format(brs2.Fields(0), "YYYY/MM/DD hh:mm:ss") '會讀到2004-03-02 12:20:33
    Set brs2 = Nothing
    If mode1 = 12 Then
      get_now = (Val(Mid(get_now, 1, 4)) - 1911) & Mid(get_now, 6, 2) & Mid(get_now, 9, 2) & Mid(get_now, 12, 2) & Mid(get_now, 15, 2) & Mid(get_now, 18, 2)
    Else
      Mid(get_now, 5, 1) = "/": Mid(get_now, 8, 1) = "/"
    End If
    Exit Function
  ElseIf Left(DB_TYPE, 6) = "ORACLE" Then
    brs2.Open "SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD HH24:MI:SS') FROM dual", rcn1, adOpenKeyset, adLockOptimistic
    get_now = brs2(0)
    Set brs2 = Nothing
  Else
    get_now = Format(Now, "YYYY/MM/DD hh:mm:ss")
  End If
  Exit Function
err1_rtn:
  If Left(computername, 8) = "QSNPCEPO" Then
    Call err2_rtn("get_now")
  Else
    Response = MsgBox("call get_now:Please check DB_TYPE=" & DB_TYPE & ",確定要離開本系統", vbOKCancel)
    If Response = vbOK Then End
  End If
End Function
Public Function find_lon(ByVal zone As String, ByRef lono As String) As Integer
'  Dim qsql As String
'  Dim rps As rdoQuery
'  qsql = "{?=CALL FIND_LON(?,?)}"
'  Set rps = rcn1.CreateQuery("find_lon", qsql)
'  rps.rdoParameters(0).Direction = rdParamReturnValue
'  'CALL FIND_LON(?,?)中2個?代表2個輸入或輸出的引數,此處第1個為輸入,第2個輸出
'  rps.rdoParameters(1) = zone
'  rps.rdoParameters(2).Direction = rdParamOutput
'  rps.Execute
'  lono = rps.rdoParameters(2)
'  find_lon = rps.rdoParameters(0)
'  Set rps = Nothing
err1_rtn:
  Call err2_rtn("find_lon")
End Function

Public Function check_right(awno1 As String, index As Integer) As Boolean
'  Dim bcn2 As rdoConnection, brs2 As New ADODB.Recordset
'  On Error GoTo err1_rtn
  If super = "Y" Or Mid(righ, (Val(awno1) - 1) * 16 + index, 1) = "1" Then check_right = True Else check_right = False
'  Set bcn2 = rdoEnvironments(0).OpenConnection("sqldb", rdDriverNoPrompt, False)
'  brs2.open "select * from pd where pd_user='" & loginuser & "'", rdOpenKeyset, rdConcurRowver)
'  If brs2!pd_super = "Y" Or Mid(brs2!pd_righ, (Val(awno1) - 1) * 16 + index, 1) = "1" Then check_right = True Else check_right = False
'  brs2.Close: bcn2.Close
  Exit Function
err1_rtn:
  Call err2_rtn("check_right")
End Function

Public Function put_cw(awno1 As String, srid As String, from1 As String, to1 As String, name1 As String, first1 As String, Optional show1 As String) As Integer
  Dim brs2 As New ADODB.Recordset, cvar0 As String, ii As Integer
  Dim model As String * 3, new_srid As String * 1, palt As Integer, cnt As Integer
  On Error GoTo err1_rtn
  If name1 = "" Then name1 = String(" ", 8)
  If srid = TR_TEST Then
      new_srid = srid
  ElseIf Left(from1, 1) = "(" Then
      If is_lono(awno1, to1) Then new_srid = TR_000000_IN Else new_srid = TR_000000_OUT
  ElseIf is_lono(awno1, from1) = False And is_lono(awno1, to1) = False Then
      new_srid = TR_MOVE_STNO
  Else
      new_srid = srid
  End If
  cnt = 0: sitm_buf = get_now()
  Do While cnt < 3
      brs2.Open "select * from dt where dt_awno='" & awno1 & "' and dt_clss=1", rcn1, adOpenKeyset, adLockOptimistic
      If Left(project_name, 2) = "SK" And awno1 = "1" Then
        palt = brs2!dt_valu + 1: If palt >= 8000 Then palt = 1
      ElseIf Left(project_name, 2) = "SP" Then
        If (Left(from1, 1) >= "A" And Left(from1, 1) <= "C") Or (Left(to1, 1) >= "A" And Left(to1, 1) <= "C") Then
          palt = brs2!dt_valu1 + 1: If palt >= 16001 Then palt = 1 '製程端用1-16000
        Else
          palt = brs2!dt_valu + 1: If palt >= 32001 Or palt < 16001 Then palt = 16001 '出貨端用16001-32000
        End If
      Else
        palt = brs2!dt_valu + 1: If palt >= 32767 Then palt = 1
      End If
      Set brs2 = Nothing
      For ii = 1 To 1000
        brs2.Open "select cw_palt from cw where cw_awno='" & awno1 & "' and cw_palt=" & palt, rcn1, adOpenKeyset, adLockOptimistic
        If brs2.EOF Then Exit For
        If Left(project_name, 4) = "SK03" And awno1 = "5" Then
          palt = palt + 1: If palt >= 9999 Then palt = 1 'CCL3地下道只能寫到9999
        ElseIf Left(project_name, 4) >= "SK04" And awno1 = "1" Then
          palt = palt + 1: If palt >= 8000 Then palt = 1
        ElseIf Left(project_name, 2) = "SP" Then
          If (Left(from1, 1) >= "A" And Left(from1, 1) <= "C") Or (Left(to1, 1) >= "A" And Left(to1, 1) <= "C") Then
            palt = palt + 1: If palt >= 16001 Then palt = 1 '製程端用1-16000
          Else
            palt = palt + 1: If palt >= 32001 Or palt < 16001 Then palt = 16001 '出貨端用16001-32000
          End If
        Else
          palt = palt + 1: If palt >= 32767 Then palt = 1
        End If
        Set brs2 = Nothing: DoEvents
      Next ii
      Set brs2 = Nothing
      If Left(project_name, 2) = "SP" Then
        If (Left(from1, 1) >= "A" And Left(from1, 1) <= "C") Or (Left(to1, 1) >= "A" And Left(to1, 1) <= "C") Then
          rcn1.Execute "update dt set dt_valu1=" & palt & " where dt_awno='" & awno1 & "' and dt_clss=1"
        Else
          rcn1.Execute "update dt set dt_valu=" & palt & " where dt_awno='" & awno1 & "' and dt_clss=1"
        End If
      Else
        rcn1.Execute "update dt set dt_valu=" & palt & " where dt_awno='" & awno1 & "' and dt_clss=1"
      End If
      cvar0 = "3"
      If is_lono(awno1, from1) Then
          rcn1.Execute "update lo set lo_sgst='R',lo_sitm='" & sitm_buf & "' where lo_awno='" & awno1 & "' and lo_lono='" & from1 & "'"
          brs2.Open "select sg_sitm from sg where sg_awno='" & awno1 & "' and sg_lono='" & from1 & "' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
          If Not brs2.EOF Then
            sitm_buf = brs2(0): Set brs2 = Nothing
            '同一庫格有多筆資料時,sitm務必不同
            rcn1.Execute "update sg set sg_first='" & first1 & "',sg_stat='0',sg_palt=" & palt & ",sg_from='" & Left(Trim(from1) & "     ", 8) & "',sg_to='" & Left(Trim(to1) & "     ", 8) & "',sg_name='" & name1 & "',sg_srid='" & new_srid & "',sg_srnm='" & get_srnm(new_srid) & "',sg_opno='" & get_opno(new_srid) & "',sg_mveq='          ',sg_now='        ',sg_nxt='        ' where sg_awno='" & awno1 & "' and sg_lono='" & from1 & "' and sg_sitm='" & sitm_buf & "'"
            If Left(project_name, 4) = "SP01" And is_stno(awno1, to1) Then
              If Left(to1, 1) <= "C" Then cvar0 = "1" Else cvar0 = "2"
              rcn1.Execute "update sg set sg_cvar='" & cvar0 & "' where sg_awno='" & awno1 & "' and sg_lono='" & from1 & "' and sg_sitm='" & sitm_buf & "'"
            End If
            rcn1.Execute "insert into cw select * from sg where sg_awno='" & awno1 & "' and sg_lono='" & from1 & "' and sg_sitm='" & sitm_buf & "'"
            If Left(project_name, 4) <> "SP01" Then
              rcn1.Execute "update cw set cw_sitm='" & get_now() & "' where cw_awno='" & awno1 & "' and cw_lono='" & from1 & "' and cw_sitm='" & sitm_buf & "'"
            End If
          Else
            Set brs2 = Nothing
            GoTo rtn
          End If
      ElseIf is_lono(awno1, to1) Then
          brs2.Open "select sg_sitm from sg where sg_awno='" & awno1 & "' and sg_lono='" & to1 & "' order by sg_sitm", rcn1, adOpenKeyset, adLockOptimistic
          If Not brs2.EOF Then
            sitm_buf = brs2(0): Set brs2 = Nothing
            rcn1.Execute "update lo set lo_sgst='S',lo_sitm='" & sitm_buf & "' where lo_awno='" & awno1 & "' and lo_lono='" & to1 & "'"
            rcn1.Execute "update sg set sg_first='" & first1 & "',sg_stat='0',sg_palt=" & palt & ",sg_from='" & Left(Trim(from1) & "     ", 8) & "',sg_to='" & Left(Trim(to1) & "     ", 8) & "',sg_name='" & name1 & "',sg_srid='" & new_srid & "',sg_srnm='" & get_srnm(new_srid) & "',sg_opno='" & get_opno(new_srid) & "',sg_mveq='          ',sg_now='        ',sg_nxt='        ' where sg_awno='" & awno1 & "' and sg_lono='" & to1 & "' and sg_sitm='" & sitm_buf & "'"
            If Left(project_name, 4) = "SP01" And is_stno(awno1, from1) Then
              If Left(from1, 1) <= "C" Then cvar0 = "1" Else cvar0 = "2"
              rcn1.Execute "update sg set sg_cvar='" & cvar0 & "' where sg_awno='" & awno1 & "' and sg_lono='" & to1 & "' and sg_sitm='" & sitm_buf & "'"
            End If
            rcn1.Execute "insert into cw select * from sg where sg_awno='" & awno1 & "' and sg_lono='" & to1 & "' and sg_sitm='" & sitm_buf & "'"
            rcn1.Execute "update cw set cw_sitm='" & get_now() & "' where cw_awno='" & awno1 & "' and cw_lono='" & to1 & "' and cw_sitm='" & sitm_buf & "'"
          Else
            rcn1.Execute "update lo set lo_sgst='E',lo_sitm='" & sitm_buf & "' where lo_awno='" & awno1 & "' and lo_lono='" & to1 & "'"
            Set brs2 = Nothing: GoTo rtn
          End If
      Else '針對站對站case,sg不存在的情況
rtn:
          brs2.Open "select * from sg", rcn, adOpenKeyset, adLockOptimistic
          'Global Const TR_LO As String * 1 = "2"          '庫格維護作業
          'brs2.Open "select * from tr where tr_opno='2'", rcn1, adOpenKeyset, adLockOptimistic
          '玻纖case,tr存6個月總計219203 records,select * from tr需時1 mins,所以tr不可太大
          'brs2.Open "select * from tr", rcn1, adOpenKeyset, adLockOptimistic
          If brs2.EOF Then
            Set brs2 = Nothing: MsgBox ("put_cw:sg檔沒東西異常,無法轉出cw,請執行庫格資料查詢及維護作業-S360,新增一筆庫存資料後即可開始執行本作業"): Exit Function
          End If
          'sg_stat='2',防insert cw後立即執行,insert cw後cw_stat再改為'0'
          rcn1.Execute "update sg set sg_stat='2' where sg_awno='" & brs2!sg_awno & "' and sg_lono='" & brs2!sg_lono & "' and sg_ptno='" & brs2!sg_ptno & "' and sg_sitm='" & brs2!sg_sitm & "'"
          If Left(project_name, 4) = "SP01" Then
            If is_stno(awno1, from1) Then
              If Left(from1, 1) <= "C" Then cvar0 = "1" Else cvar0 = "2"
              rcn1.Execute "update sg set sg_cvar='" & cvar0 & "',sg_mode='" & show1 & "' where sg_awno='" & brs2!sg_awno & "' and sg_lono='" & brs2!sg_lono & "' and sg_ptno='" & brs2!sg_ptno & "' and sg_sitm='" & brs2!sg_sitm & "'"
            ElseIf is_stno(awno1, to1) Then
              If Left(to1, 1) <= "C" Then cvar0 = "1" Else cvar0 = "2"
              rcn1.Execute "update sg set sg_cvar='" & cvar0 & "' where sg_awno='" & brs2!sg_awno & "' and sg_lono='" & brs2!sg_lono & "' and sg_ptno='" & brs2!sg_ptno & "' and sg_sitm='" & brs2!sg_sitm & "'"
            End If
          End If
          rcn1.Execute "insert into cw select * from sg where sg_awno='" & brs2!sg_awno & "' and sg_lono='" & brs2!sg_lono & "' and sg_ptno='" & brs2!sg_ptno & "' and sg_sitm='" & brs2!sg_sitm & "'"
          If Left(from1, 4) = "(車上)" Then
            rcn1.Execute "update cw set cw_awno='" & awno1 & "',cw_palt=" & palt & ",cw_mveq='" & mveq_s210 & "',cw_stat='1',cw_lono=' ',cw_from='" & from1 & "',cw_to='" & to1 & "',cw_now='" & from1 & "',cw_nxt='" & to1 & "',cw_err=0,cw_first='" & first1 & "',cw_sitm='" & sitm_buf & "',cw_sitm1='" & sitm_buf & "',cw_date='" & Left(sitm_buf, 10) & "',cw_opno='" & get_opno(new_srid) & "',cw_srid='" & new_srid & "',cw_srnm='" & get_srnm(new_srid) & "',cw_name='" & name1 & "',cw_quty=0,cw_pqty=0,cw_ptno=' ',cw_odno=' ',cw_ltno=' ' where cw_awno='" & brs2!sg_awno & "' and cw_lono='" & brs2!sg_lono & "' and cw_ptno='" & brs2!sg_ptno & "' and cw_sitm='" & brs2!sg_sitm & "'"
          Else
            rcn1.Execute "update cw set cw_awno='" & awno1 & "',cw_palt=" & palt & ",cw_mveq=' ',cw_stat='0',cw_lono=' ',cw_from='" & from1 & "',cw_to='" & to1 & "',cw_now=' ',cw_nxt=' ',cw_err=0,cw_first='" & first1 & "',cw_sitm='" & sitm_buf & "',cw_sitm1='" & sitm_buf & "',cw_date='" & Left(sitm_buf, 10) & "',cw_opno='" & get_opno(new_srid) & "',cw_srid='" & new_srid & "',cw_srnm='" & get_srnm(new_srid) & "',cw_name='" & name1 & "',cw_quty=0,cw_pqty=0,cw_ptno=' ',cw_odno=' ',cw_ltno=' ' where cw_awno='" & brs2!sg_awno & "' and cw_lono='" & brs2!sg_lono & "' and cw_ptno='" & brs2!sg_ptno & "' and cw_sitm='" & brs2!sg_sitm & "'"
          End If
          Set brs2 = Nothing
          '                                    awno            palt         mveq stat lono from            to            now nxt Err stno first            sitm               sitm1              date                         opno                         srid               srnm                         name            prog quty pqty ptno spec odno ltno vendor berry grade mchno class prod bas widt leng net gros mark loi process head leng1 check ccl bar  gros1 gros2 tax nohead glue seq mark1 quty1 quty2 tty seq1  match equal a651barcode a652barcode lono1 mark2
          'rcn1.Execute "insert into cw values('" & awno1 & "'," & palt & ",' ' ,'0' ,' ' ,'" & from1 & "','" & to1 & "',' ',' ',0  ,' ' ,'" & first1 & "','" & sitm_buf & "','" & sitm_buf & "','" & Left(sitm_buf, 10) & "','" & get_opno(new_srid) & "','" & new_srid & "','" & get_srnm(new_srid) & "','" & name1 & "',' ' ,0   ,0   ,' ' ,' ' ,' ' ,' ' ,' '   ,' '  ,' '  ,' '  ,' '  ,' ' ,0  ,0   ,0   ,0  ,0   ,' ' ,0  ,' ',' ',0,' ',' ',' ',0,0,' ',0,0,0,' ',0,0,' ',0,' ',' ',' ',' ',' ',' ')"
      End If
      brs2.Open "select cw_palt from cw where cw_awno='" & awno1 & "' and cw_palt=" & palt, rcn1, adOpenKeyset, adLockOptimistic
      If Not brs2.EOF Then
         Set brs2 = Nothing: Exit Do
      End If
       Set brs2 = Nothing: cnt = cnt + 1: Wait (1)
  Loop
  If cnt = 2 Then
      Call syslog("put_cw", palt, cnt, "投入工作檔時序號重覆" & cnt & "次異常,new_srid=" & new_srid & ",srid=" & srid & ",from=" & from1 & ",to=" & to1)
      Exit Function
  End If
  If Left(prog, 3) <> "PLC" And Left(prog, 3) <> "STV" And Left(prog, 4) <> "CRAN" And Left(prog, 6) <> "MAPATH" And Left(prog, 6) <> "RS6000" And Left(prog, 4) <> "TTCS" And Left(prog, 4) <> "MONO" Then
      If Left(prog, 4) <> "S330" And Left(prog, 4) <> "S001" And Left(prog, 4) <> "S340" Then
          If show1 <> "no_show" Then Response = MsgBox("投入工作檔完成,開始執行" & Trim(get_srnm(new_srid)) & "作業")
      End If
  End If
rtn1:
  If Left(project_name, 2) = "SK" Then
    Call syslog1("put_cw", 0, 0, "new_srid=" & get_srnm(new_srid) & ",srid=" & srid & ",palt=" & palt & ",from=" & from1 & ",to=" & to1)
  End If
  put_cw = palt: Exit Function
err1_rtn:
  Call err2_rtn("put_cw")
End Function

Public Function get_opno(srid As String) As String
  If srid <= "7" Then
    get_opno = TR_OUT
  ElseIf srid <= "I" Then
    get_opno = TR_IN
  ElseIf srid <= "N" Then
    get_opno = TR_LO
  ElseIf srid = TR_TEST Then
    get_opno = TR_TEST
  Else
    get_opno = TR_CW
  End If
End Function
Public Function is_lono(awno1 As String, lono As String) As Boolean
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  'If Left(lono, 1) <> "0" And Left(lono, 1) <> "1" Then
  '  is_lono = False
  'Else
    brs2.Open "select lo_lono from lo where lo_awno='" & awno1 & "' and lo_lono='" & Trim(lono) & "'", rcn1, adOpenKeyset, adLockOptimistic
    If brs2.EOF = False Then is_lono = True Else is_lono = False
    Set brs2 = Nothing
  'End If
  Exit Function
err1_rtn:
  Call err2_rtn("is_lono")
End Function
Public Function is_lono_sg(awno1 As String, lono As String) As Boolean
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  If Left(lono, 1) <> "0" And Left(lono, 1) <> "1" Then
    is_lono_sg = False
  Else
    brs2.Open "select sg_lono from sg where sg_awno='" & awno1 & "' and sg_lono='" & lono & "'", rcn1, adOpenKeyset, adLockOptimistic
    If brs2.EOF = False Then is_lono_sg = True Else is_lono_sg = False
    Set brs2 = Nothing
  End If
  Exit Function
err1_rtn:
  Call err2_rtn("is_lono_sg")
End Function

Public Function is_stno(awno1 As String, stno As String) As Boolean
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  brs2.Open "select st_stno from st where st_awno='" & awno1 & "' and st_stno like '" & Trim(stno) & "%'", rcn1, adOpenKeyset, adLockOptimistic
  'brs2.Open "select st_stno from st where st_awno='" & awno1 & "' and st_stno='" & Left(stno, STNO_NUM) & "'", rcn1, adOpenKeyset, adLockOptimistic
  If brs2.EOF = False Then is_stno = True Else is_stno = False
  Set brs2 = Nothing: Exit Function
err1_rtn:
  Call err2_rtn("is_stno")
End Function

Public Function check_block(nxt As String, mveq As String, stat As String) As Boolean
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  check_block = False
  If stat >= ST_RQST And stat <= ST_WORK Then
    If Left(mveq, 1) = " " Then
'異常碼: -2147467259,Your transaction (ProcessID#30及ProcessID#39) was deadlocked with another process and has been chosen as the deadlock victim.Rerun your transaction.
      brs2.Open "select * from cw where cw_nxt='" & nxt & "' and cw_stat='" & stat & "'", rcn1, adOpenKeyset, adLockOptimistic
    Else
      brs2.Open "select * from cw where cw_nxt='" & nxt & "' and cw_mveq='" & mveq & "' and cw_stat='" & stat & "'", rcn1, adOpenKeyset, adLockOptimistic
    End If
    If brs2.EOF = False Then check_block = True
  Else
    brs2.Open "select * from cw where cw_nxt='" & nxt & "' and cw_mveq='" & mveq & "'", rcn1, adOpenKeyset, adLockOptimistic
    If brs2.EOF = False Then check_block = True
  End If
  Set brs2 = Nothing: Exit Function
err1_rtn:
  Call err2_rtn("check_block")
End Function

Public Sub WCenter_Place(XFrm As Form, Cfrm As Form, Pmode As Integer)
  Dim Stbar As Integer
  On Error GoTo err1_rtn
  If (Pmode = 0) Then                                               ' For Screen
    Cfrm.Left = (Screen.Width - Cfrm.Width) * 9 / 20                ' Center form horizontally.
    Cfrm.Top = (Screen.Height - Cfrm.Height) * 9 / 20               ' Center form vertically.
  ElseIf (Pmode = 1) Then                                           ' For first_prog
    If (Mid(XFrm.name, 2, 3) = "000" Or Mid(XFrm.name, 2, 3) = "001") Then
      Stbar = XFrm.StatusBar.Height * 2 / 3
    Else
      Stbar = 0
    End If
    Cfrm.Left = (XFrm.Width - Cfrm.Width) / 2 + XFrm.Left            ' Center form horizontally.
    Cfrm.Top = (XFrm.Height + Stbar - Cfrm.Height) / 2 + XFrm.Top    ' Center form vertically.
  End If
  Exit Sub
err1_rtn:
  Call err2_rtn("WCenter_Place")
End Sub
Public Sub Wait(WaitSeconds As Single)
  Dim StartTime As Single, CurrTime As Single
  StartTime = Timer
  Do
    DoEvents
    CurrTime = Timer
    If (CurrTime < StartTime) Then
      WaitSeconds = WaitSeconds - (86400 - StartTime)
      StartTime = 0
    End If
  Loop While CurrTime < StartTime + WaitSeconds
End Sub
Public Sub Sys_MsgShow(Cfrm As Form, ErMsg As String)
  Dim brs2 As New ADODB.Recordset
  On Error GoTo err1_rtn
  brs2.Open "select * from pd where pd_user='" & loginuser & "'", rcn1, adOpenKeyset, adLockOptimistic
  If Not brs2.EOF Then Cfrm.Label999 = "登錄者: " & brs2!pd_name & "  ,  " & ErMsg & "     最新版本: " & get_string_value1("", "", "pa_ver") & "  ,  現在時間: " & get_now()
  'If Not brs2.EOF Then Cfrm.Label999 = "登錄者: " & brs2!pd_name & "  ,  " & ErMsg & "  ,  系統版本: " & PassWd.Label1(1) & "  ,  現在時間: " & get_now()
'  If Not brs2.EOF Then Cfrm.StatusBar1.Panels(1).Text = "登錄者: " & brs2!pd_name & " , " & ErMsg & " , 系統版本: " & PassWd.Label1(1)
  Set brs2 = Nothing
  DoEvents
  Exit Sub
err1_rtn:
  Call err2_rtn("Sys_MsgShow")
End Sub
Public Sub err2_rtn(Msg As String, Optional mode As String)
  Dim er As rdoError, lib_buf1 As String
  If mode = "show" Or Left(prog, 4) = "S001" Or Left(prog, 4) = "CRAN" Or Left(prog, 3) = "PLC" Or Left(prog, 6) = "MAPATH" Then
    first_prog.Print Err, Error
  End If
  For Each er In rdoErrors
    If mode = "show" Or Left(prog, 4) = "S001" Or Left(prog, 4) = "CRAN" Or Left(prog, 3) = "PLC" Or Left(prog, 6) = "MAPATH" Then
      first_prog.Print er.Description, er.Number
    End If
  Next er
  lib_buf1 = prog & "-" & Msg & ":系統異常,行號為:" & Erl & ",異常碼:" & Str(Err.Number) & Chr(13) & Err.Description
  Call syslog1("err2_rtn:" & Msg, 999, 0, lib_buf1)
  If Err.Number = -2147467259 Then   'SQL deadlock case,或[DBNETLIB]ConnectionOpen: SQL Server不存在或拒絕存取,或連線失敗(for SO專案Save程式)
    Call Wait(2)
    If False And Mid(Trim(Err.Description), 1, 10) = "[DBNMPNTW]" Then '[DBNMPNTW] ConnectionWrite,rcn被close
      If Left(computername, 8) = "QSNPCEPO" Then
        rcn.Close: rcn1.Close
        rcn.ConnectionTimeout = 3: rcn.CursorLocation = adUseClient:  rcn.Open rcn_str
        rcn1.ConnectionTimeout = 3: rcn1.CursorLocation = adUseClient:  rcn1.Open rcn_str
      End If
    ElseIf Mid(Trim(Err.Description), 1, 16) = "Your transaction" Or Mid(Trim(Err.Description), 1, 11) = "Transaction" Then
'異常碼: -2147467259,Your transaction (ProcessID#30及ProcessID#39) was deadlocked with another process and has been chosen as the deadlock victim.Rerun your transaction.
'[CRAN7-err2_rtn:st_comd_cran] (999,0) CRAN7-st_comd_cran:系統異常,行號為:3432,異常碼:-2147467259Transaction (Process ID 65) was deadlocked on lock resources with another process and has been chosen as the deadlock victim. Rerun the transaction.,行號為:3432
'[CRAN1-err2_rtn:st_comd_cran] (999,0) CRAN1-st_comd_cran:系統異常,行號為:3432,異常碼:-2147467259Transaction (Process ID 61) was deadlocked on lock resources with another process and has been chosen as the deadlock victim. Rerun the transaction.,行號為:3432
'[Check_block]
'[rd_data_cran]
    End If
  'ElseIf Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS" or left(prog,4)= "SAVE" Then
  '  End
  Else
    If mode = "noMsgBox" Then Call Wait(5) Else Response = MsgBox(lib_buf1, , Err.Source) ': End
  End If
  If Left(project_name, 2) = "SO" Then End
  'If True Or err2_cnt >= 5 Then
  '   Response = MsgBox(lib_buf1, , Err.Source): End
  'Else
  '  rcn.Close: rcn1.Close: Call Wait(2): err2_cnt = err2_cnt + 1
  '  rcn.ConnectionTimeout = 3: rcn.CursorLocation = adUseClient:  rcn.Open rcn_str
  '  rcn1.ConnectionTimeout = 3: rcn1.CursorLocation = adUseClient:  rcn1.Open rcn_str
  'End If
End Sub
'Open "c:\ys2\tty.txt" For Input As #1
'Do While Not EOF(1)
'  Line Input #1, tty
'Loop
'Close #1   ' 關閉檔案。
'注意log檔稍大時(200行左右),當連線程式要寫一個訊息時,會讓連線程式變很慢(會idle近5秒),所以試車完時務必取消寫入log檔指令
Public Sub syslog(prog1 As String, err1 As Integer, err2 As Integer, Msg As String, Optional mode As String)
Dim ct As String, ii As Integer
Dim fname As String
Dim lib_buf As String * 300, buf As String
    ct = Format(Now, "YYYY/MM/DD hh:mm:ss")
    lib_buf = Mid(ct, 1, 19) & " [" & Trim(prog) & "-" & Trim(prog1) & "] " & "(" & err1 & "," & err2 & ") " & Msg
    If False Then 'Left(prog, 4) = "S001" Then
       If Dir("c:\log", vbDirectory) = "" Then
          MkDir ("c:\log"): MkDir ("c:\log1")
       End If
       fname = "c:\log\s" & Mid(ct, 9, 2) & ".log"
       If Dir(fname) <> "" Then
          ii = FreeFile
          Open fname For Input As ii
          Line Input #ii, buf: Close #ii
          If Mid(buf, 1, 10) <> Mid(ct, 1, 10) Then Kill fname
       End If
       ii = FreeFile
       Open fname For Append As ii
       Print #ii, lib_buf
       Close #ii
       Debug.Print lib_buf
    End If
    If refresh1_screen = 44 Then
       refresh1_screen = 0: first_prog.Cls
    Else
       refresh1_screen = refresh1_screen + 1
    End If
    If mode = "show" Or Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS" Or Left(prog, 4) = "SAVE" Then first_prog.Print lib_buf
  If Err.Number = -2147467259 Then 'SQL deadlock case,或[DBNETLIB]ConnectionOpen: SQL Server不存在或拒絕存取,或連線失敗(for SO專案Save程式)
    Call Wait(2)
    If Mid(Err.Description, 1, 10) = "[DBNMPNTW]" Then '[DBNMPNTW] ConnectionWrite,rcn被close
      If Left(computername, 8) = "QSNPCEPO" Then
        Call syslog("[DBNMPNTW]", 999, 888, lib_buf)
        rcn.Close: rcn1.Close
        rcn.ConnectionTimeout = 3: rcn.CursorLocation = adUseClient:  rcn.Open rcn_str
        rcn1.ConnectionTimeout = 3: rcn1.CursorLocation = adUseClient:  rcn1.Open rcn_str
      End If
    End If
  End If
End Sub
Public Sub syslog1(prog1 As String, err1 As Integer, err2 As Integer, Msg As String, Optional mode As String)
Dim ct As String, ii As Integer
Dim fname As String
Dim lib_buf As String * 300, buf As String
    If Left(project_name, 2) = "SF" And err1 <> 999 Then Exit Sub
    If Dir("c:\log", vbDirectory) = "" Then
       MkDir ("c:\log"): MkDir ("c:\log1")
    End If
    ct = Format(Now, "YYYY/MM/DD hh:mm:ss")
    fname = "c:\log\s" & Mid(ct, 9, 2) & ".log"
    If Dir(fname) <> "" Then
       ii = FreeFile
       Open fname For Input As ii
       Line Input #ii, buf: Close #ii
       If Mid(buf, 1, 10) <> Mid(ct, 1, 10) Then Kill fname
    End If
    ii = FreeFile
    Open fname For Append As ii
    lib_buf = Mid(ct, 1, 19) & " [" & Trim(prog) & "-" & Trim(prog1) & "] " & "(" & err1 & "," & err2 & ") " & Msg & ",行號為:" & Erl
    Print #ii, lib_buf
    Close #ii
    If refresh1_screen = 44 Then
       refresh1_screen = 0: first_prog.Cls
    Else
       refresh1_screen = refresh1_screen + 1
    End If
    If mode = "show" Or Left(prog, 3) = "PLC" Or Left(prog, 3) = "STV" Or Left(prog, 4) = "CRAN" Or Left(prog, 6) = "MAPATH" Or Left(prog, 4) = "MONO" Or Left(prog, 4) = "TTCS" Or Left(prog, 4) = "SAVE" Then first_prog.Print lib_buf
End Sub


'"FFFFFFFF","80888888"會產生溢位異常 ,"79888888"不會產生溢位異常 ,"0FFFFFFF"==>268435455 , "00FFFFFF"==>16777215
Public Function hextodec(buff As String, len1 As Integer) As Long
  Dim i As Integer, k As Long, m As Integer
'  If Left(buff, 1) = "F" Then
'    hextodec = 99999999: Exit Function
'  End If
  If Len(Trim(buff)) >= 8 And buff >= "79888888" Then
    hextodec = 99999999: Exit Function
  End If
  hextodec = 0: k = 1
  For i = 1 To len1
    If Mid(buff, len1 - i + 1, 1) >= "A" And Mid(buff, len1 - i + 1, 1) <= "F" Then
      m = Asc(Mid(buff, len1 - i + 1, 1)) - 55
    Else
      m = Val(Mid(buff, len1 - i + 1, 1))
    End If
    If i = 1 Then
      hextodec = m
    Else
      k = k * 16: hextodec = hextodec + m * k
    End If
  Next i
End Function

'由PLC讀出之HEX字串作轉換,例model=4表將len1=8 bytes的字串"00310032"轉成2 bytes的字串"12",例model=2表將len1=4 bytes的字串"3132"轉成2 bytes的字串"12"
Public Function hextoasc(model As Integer, buff As String, len1 As Integer) As String
  Dim i As Integer, j As Long, m As Integer
  hextoasc = ""
  m = len1 / model
  For i = 1 To m
    j = hextodec(Mid(buff, (i - 1) * model + 1, model), model)
    hextoasc = hextoasc & Chr(j)
  Next i
End Function
Public Function cran_read_fx(port_id As Integer, addr As Integer, leng As Integer, lib_buf1 As String, mveq As String) As Boolean
  Dim LL As Long, wbuf As String, ii As Integer, jj As Integer, lib_buf As String * 300
  Dim awno1 As String * 1, tXDMID$, r$
  On Error GoTo err1_rtn
  awno1 = get_awno(port_id)
  '要由addr=D100開始讀leng組D暫存器,讀出時每組會佔4bytes
  'addr = 60: leng = 13  '要由addr=D100開始讀leng組D暫存器,讀出時每組會佔4bytes
  'addr轉換公式 : Right("00" + Hex(addr * 2 + Val("&H1000")), 4) , addr=D100轉換成10C8
  'leng轉換公式 : Right("00" + Trim(Hex(leng * 2)), 2) , leng=13轉換成1A
  '直接連線火狐狸的programming port,可以少買一片RS232卡(走標準協定)
  tXDMID$ = "0" + Right("00" + Hex(addr * 2 + Val("&H1000")), 4) + Right("00" + Trim(Hex(leng * 2)), 2) + Chr(3)
  Call sio_flush(port_id, 2): lib_buf = ""
  wbuf = Chr(2) + tXDMID$ + Sumcheck(tXDMID$)
  Call sio_write(port_id, wbuf, Len(wbuf)) 'Len(wbuf)=11
  If (leng * 4 + 4) > 100 Then Call Wait(0.4) Else Call Wait(0.3)
  LL = sio_read(port_id, lib_buf, leng * 4 + 4)  '讀到Stx + data + Etx + 2 Bytes Sumcheck
  If LL <> (leng * 4 + 4) Then
    ii = get_integer_value1(awno1, mveq, "cr_err")
    lib_buf = "cran read err,port=" & port_id & ",length=" & LL & ",buf=" & Left(lib_buf, leng * 4 + 4): ii = LL: Call syslog("Cran_rd_data_cran1", leng * 4 + 4, ii, lib_buf)
    If LL = 0 Then
      DoEvents 'Call wait1_100ms(port_id,1)
      If ii <> 67 Then rcn1.Execute "update cr set cr_err=67 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
    Else
      DoEvents 'Call wait1_100ms(port_id,1)
      If ii <> 65 Then rcn1.Execute "update cr set cr_err=65 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
    End If
    cran_read_fx = False: Exit Function
  End If
  r$ = Sumcheck(Mid(lib_buf, 2, leng * 4 + 1))
  If r$ <> Mid(lib_buf, leng * 4 + 3, 2) Then
    lib_buf = "cran read check sum err,length=" & LL & ",buf=" & Left(lib_buf, leng * 4 + 4): ii = LL: Call syslog("Cran_rd_data_cran1", leng * 4 + 4, ii, lib_buf)
    DoEvents 'Call wait1_100ms(port_id, 1)
    If ii <> 66 Then rcn1.Execute "update cr set cr_err=66 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
    cran_read_fx = False: Exit Function
  End If
'  If get_integer_value1(awno1, mveq, "cr_err") >= 65 And get_integer_value1(awno1, mveq, "cr_err") <= 67 Then rcn1.Execute "update cr set cr_err=0 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  lib_buf1 = lib_buf
  cran_read_fx = True
  Exit Function
err1_rtn:
  Call err2_rtn("cran_read_fx")
End Function
Public Function Sumcheck(BUFFER_sum As String) As String
  Dim res As Integer, ii As Integer
  res = 0
  For ii = 1 To Len(BUFFER_sum)
    res = res + Asc(Mid(BUFFER_sum, ii, 1))
    'Debug.Print RESULT%
  Next ii
  Sumcheck = Right("00" + Hex(res), 2)
End Function
Public Function cran_write_fx(port_id As Integer, addr As Integer, leng As Integer, data1 As String, mveq As String, Optional palt As Integer, Optional now1 As String, Optional nxt1 As String) As Boolean
  Dim LL As Long, wbuf As String, ii As Integer, jj As Integer, lib_buf As String * 300
  Dim awno1 As String * 1
  On Error GoTo err1_rtn
  awno1 = get_awno(port_id)
  '要由addr=D100開始寫leng組資料至D暫存器,寫入時每組會佔4bytes
  'addr轉換公式 : Right("00" + Hex(addr * 2 + Val("&H1000")), 4) , addr=D100轉換成10C8
  'leng轉換公式 : Right("00" + Trim(Hex(leng * 2)), 2) , leng=13轉換成1A
  Dim tXDMID$
  ' 1 10C8 0C 0000000001000100080002000F
  tXDMID$ = "1" + Right("00" + Hex(addr * 2 + Val("&H1000")), 4) + Right("00" + Trim(Hex(leng * 2)), 2)
  '寫leng組資料至D暫存器
  'Dim buf1 as string*4
  'For ii = 1 To leng
  '  jj = ii + 10
  '  buf1 = Mid(Right("0000" + Hex(jj), 4), 3, 2) + Mid(Right("0000" + Hex(jj), 4), 1, 2)
  '  tXDMID$ = tXDMID$ + buf1
  'Next ii
  tXDMID$ = tXDMID$ + Mid(data1, 1, 4 * leng) + Chr(3)
  Call sio_flush(port_id, 2)
  wbuf = Chr(2) + tXDMID$ + Sumcheck(tXDMID$)
  jj = Len(wbuf)
  Call sio_write(port_id, wbuf, Len(wbuf)) 'Len(wbuf)=11+leng*4
  Call Wait(0.3): lib_buf = "": LL = sio_read(port_id, lib_buf, leng * 4 + 4)
  If LL <> 1 Or Mid(lib_buf, 1, 1) <> Chr(6) Then
    lib_buf = "cran_write_fx err,wbuf=" & wbuf & ",rbuf=" & Mid(lib_buf, 1, 1): ii = LL: Call syslog("cran_write_fx", addr, leng, lib_buf)
    ii = get_integer_value1(awno1, mveq, "cr_err")
    If LL = 0 Then
      DoEvents 'Call wait1_100ms(port_id,1)
      If ii <> 67 Then
'        rcn1.Execute "update cr set cr_err=67 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
        rcn1.Execute "update cw set cw_err=67 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
      End If
    Else
      DoEvents 'Call wait1_100ms(port_id,1)
      If ii <> 65 Then
'        rcn1.Execute "update cr set cr_err=65 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
        rcn1.Execute "update cw set cw_err=65 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
      End If
    End If
    cran_write_fx = False: Exit Function
  End If
  If (addr < 751 Or addr > 765) And data1 <> "000000000000000000000000" Then
    If Left(project_name, 4) = "SU01" Then
      Call syslog1("cran_fx", addr, palt, now1 & "," & nxt1 & "," & wbuf)
    Else
      Call syslog1("cran_write_fx_rtn", addr, leng, wbuf)
    End If
  End If
  cran_write_fx = True
  Exit Function
err1_rtn:
  Call err2_rtn("cran_write_fx")
End Function
'response_num可能為4,5,6(for SF project)
Public Function plc_write_asrs(port_id As Integer, leng As Integer, data1 As String, mveq As String, response_num As Integer, Optional awno2 As String) As Boolean
  Dim LL As Long, wbuf As String, ii As Integer, jj As Integer, lib_buf As String * 300
  Dim awno1 As String * 1, buf2 As String * 2
  On Error GoTo err1_rtn
  plc_write_asrs = False
  If Left(project_name, 4) = "SP01" Then
    awno1 = "1"
  Else
    If awno2 >= "1" And awno2 <= "9" Then awno1 = awno2 Else awno1 = get_awno(port_id)
  End If
  Call sio_flush(port_id, 2)
  wbuf = data1: lib_buf = ""
  Call sio_write(port_id, wbuf, leng)
  If Left(project_name, 2) = "SI" Then Call Wait(1.5) Else Call Wait(0.5)
  lib_buf = "": LL = sio_read(port_id, lib_buf, response_num)
  If Left(project_name, 2) <> "SF" Then buf2 = Mid(lib_buf, response_num - 3, 2) Else buf2 = Mid(lib_buf, response_num - 4, 2)
  If LL <> response_num Or buf2 <> "11" Then
    ii = LL
    If Left(project_name, 2) = "SI" Then
      If palt_plc_write_asrs_old <> palt_plc_write_asrs And cmd_plc_write_asrs_old <> Left(wbuf, 8) Then
        lib_buf = "err,mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",plc回=" & buf2 & ",command = " & Left(wbuf, 8)
        Call syslog1("plc_write_asrs", response_num, ii, lib_buf)
        palt_plc_write_asrs_old = palt_plc_write_asrs: cmd_plc_write_asrs_old = Left(wbuf, 8)
      End If
    ElseIf Left(project_name, 2) = "SK" Then
      If Left(project_name, 4) = "SK02" Then
        If Left(get_string_value1("", "", "pa_ccl3"), 3) = "log" Then
          jj = LL: Call syslog1("st_comd_cran", leng, jj, "awno=" & awno1 & ",wbuf=" & Left(wbuf, leng) & ",rbuf=" & Left(lib_buf, jj))
        End If
      ElseIf Mid(wbuf, 3, 2) = "F20" Then
        Call syslog1("plc_write_asrs", ii, response_num, Left(wbuf, 8) & ",response=" & Left(buf2, 2))
        Call Wait(2)
      End If
    End If
    lib_buf = "send command err: mveq=" & mveq & ",leng=" & leng & ",err=" & buf2 & ",command=" & wbuf: ii = LL: Call syslog("plc_write_asrs", response_num, ii, lib_buf)
    If mveq <> " " Then ii = get_integer_value1(awno1, mveq, "cr_err")
    DoEvents 'Call wait1_100ms(port_id, 1)
    If LL = 0 Then
      If mveq = " " And Left(project_name, 2) = "SG" Then '控制自動門B13,B63開關
        rcn1.Execute "update st set st_err=67 where st_stno='" & Mid(wbuf, 4, STNO_NUM) & " '"
      Else
        If ii <> 67 Then
'          rcn1.Execute "update cr set cr_err=67 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
          rcn1.Execute "update cw set cw_err=67 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
        End If
      End If
    ElseIf buf2 = "00" Then
      Call syslog("plc_write_asrs", 53, 53, "天車狀態為非自動或搬運中或異常,因之暫不接受命令")
      If mveq = " " And Left(project_name, 2) = "SG" Then
        rcn1.Execute "update st set st_err=" & 53 & " where st_stno='" & Mid(wbuf, 4, STNO_NUM) & " '"
      Else
        If ii <> 53 Then
'          rcn1.Execute "update cr set cr_err=53 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
          rcn1.Execute "update cw set cw_err=53 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
        End If
      End If
    Else
      If mveq = " " And Left(project_name, 2) = "SG" Then
        rcn1.Execute "update st set st_err=" & Val(buf2) & " where st_stno='" & Mid(wbuf, 4, STNO_NUM) & " '"
      Else
        If ii <> Val(buf2) Then
'          rcn1.Execute "update cr set cr_err=" & Val(buf2) & " where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
          rcn1.Execute "update cw set cw_err=" & Val(buf2) & " where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
        End If
      End If
    End If
    plc_write_asrs = False: Exit Function
  End If
  palt_plc_write_asrs_old = 0: cmd_plc_write_asrs_old = "00000000"
  If Mid(data1, response_num - 3, 2) <> "04" Then
    lib_buf = "ok,mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",command = " & wbuf
    If Left(project_name, 2) = "SK" Then
      If awno1 = "1" Or awno1 = "5" Then
        'ii = LL: Call syslog1("plc_write_asrs", response_num, ii, Left(Trim(lib_buf), Len(Trim(lib_buf)) - 2))
      End If
    End If
  End If
  plc_write_asrs = True
  Exit Function
err1_rtn:
  Call err2_rtn("plc_write_asrs")
End Function
'response_num可能為4,5,6(for SF project)
Public Function plc_write_asrs_sk1(port_id As Integer, leng As Integer, data1 As String, mveq As String, response_num As Integer, Optional awno2 As String) As Boolean
  Dim LL As Long, wbuf As String, ii As Integer, jj As Integer, lib_buf As String * 300
  Dim awno1 As String * 1, buf2 As String * 2
  On Error GoTo err1_rtn
  plc_write_asrs_sk1 = False
  If Left(project_name, 4) = "SP01" Then
    awno1 = "1"
  Else
    If awno2 >= "1" And awno2 <= "9" Then awno1 = awno2 Else awno1 = get_awno(port_id)
  End If
  Call sio_flush(port_id, 2)
  wbuf = data1
  Call sio_write(port_id, wbuf, leng)
  If Left(project_name, 2) = "SI" Then Call Wait(1.5) Else Call Wait(0.5)
  lib_buf = "": LL = sio_read(port_id, lib_buf, response_num)
  If Left(project_name, 2) <> "SF" Then buf2 = Mid(lib_buf, response_num - 3, 2) Else buf2 = Mid(lib_buf, response_num - 4, 2)
  If LL <> response_num Or buf2 <> "11" Then
    ii = LL
    If Left(project_name, 2) = "SI" Then
      If palt_plc_write_asrs_old <> palt_plc_write_asrs And cmd_plc_write_asrs_old <> Left(wbuf, 8) Then
        lib_buf = "err,mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",plc回=" & buf2 & ",command = " & Left(wbuf, 8)
        Call syslog1("plc_write_asrs_sk1", response_num, ii, lib_buf)
        palt_plc_write_asrs_old = palt_plc_write_asrs: cmd_plc_write_asrs_old = Left(wbuf, 8)
      End If
    ElseIf Left(project_name, 2) = "SK" Then
      If Mid(wbuf, 3, 2) = "F20" Then
        Call syslog1("plc_write_asrs_sk1", ii, response_num, Left(wbuf, 8) & ",response=" & Left(buf2, 2))
        Call Wait(2)
      End If
    End If
    lib_buf = "send command err: mveq=" & mveq & ",leng=" & leng & ",err=" & buf2 & ",command=Chr(2)+" & wbuf: ii = LL: Call syslog("plc_write_asrs_sk1", response_num, ii, lib_buf)
    If mveq <> " " Then ii = get_integer_value1(awno1, mveq, "cr_err")
    DoEvents 'Call wait1_100ms(port_id, 1)
    If LL = 0 Then
      If mveq = " " And Left(project_name, 2) = "SG" Then '控制自動門B13,B63開關
        rcn1.Execute "update st set st_err=67 where st_stno='" & Mid(wbuf, 4, STNO_NUM) & " '"
      Else
        If ii <> 67 Then
'          rcn1.Execute "update cr set cr_err=67 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
          rcn1.Execute "update cw set cw_err=67 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
        End If
      End If
    ElseIf buf2 = "00" Then
      Call syslog("plc_write_asrs_sk1", 53, 53, "天車狀態為非自動或搬運中或異常,因之暫不接受命令")
      If mveq = " " And Left(project_name, 2) = "SG" Then
        rcn1.Execute "update st set st_err=" & 53 & " where st_stno='" & Mid(wbuf, 4, STNO_NUM) & " '"
      Else
        If ii <> 53 Then
'          rcn1.Execute "update cr set cr_err=53 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
          rcn1.Execute "update cw set cw_err=53 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
        End If
      End If
    Else
      If mveq = " " And Left(project_name, 2) = "SG" Then
        rcn1.Execute "update st set st_err=" & Val(buf2) & " where st_stno='" & Mid(wbuf, 4, STNO_NUM) & " '"
      Else
        If ii <> Val(buf2) Then
'          rcn1.Execute "update cr set cr_err=" & Val(buf2) & " where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
          rcn1.Execute "update cw set cw_err=" & Val(buf2) & " where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
        End If
      End If
    End If
    If Left(project_name, 4) = "SK01" Then
      If Mid(wbuf, 2, 2) = "00" Or Mid(wbuf, 2, 2) = "08" Then
        lib_buf = "send command err: mveq=" & mveq & ",leng=" & leng & ",err=" & buf2 & ",command=Chr(2)+" & wbuf: ii = LL: Call syslog1("plc_write_asrs_sk1", response_num, ii, lib_buf)
        plc_write_asrs_sk1 = True: Exit Function
      End If
    End If
    plc_write_asrs_sk1 = False: Exit Function
  End If
  palt_plc_write_asrs_old = 0: cmd_plc_write_asrs_old = "00000000"
  If Mid(data1, response_num - 3, 2) <> "04" Then
    lib_buf = "ok,mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",command = " & wbuf
    If Left(project_name, 2) = "SK" Then
      If awno1 = "1" Or awno1 = "5" Then
        ii = LL: Call syslog1("plc_write_asrs_sk1", response_num, ii, Left(Trim(lib_buf), Len(Trim(lib_buf)) - 2))
      End If
    End If
  End If
  plc_write_asrs_sk1 = True
  Exit Function
err1_rtn:
  Call err2_rtn("plc_write_asrs_sk1")
End Function

'葉樹堅吊車
Public Function rd_data_cran(port_id As Integer, mveq As String, Optional awno2 As String) As Boolean
Dim brs5 As New ADODB.Recordset, brs3 As New ADODB.Recordset, brs4 As New ADODB.Recordset
Dim err1 As Integer, ii As Integer, jj As Integer, kk As Integer, lib_buf As String * 300, lib_buf1 As String * 300
Dim LL As Long, wbuf As String
Dim auto1 As String * 1, mvst As String * 1, old_mvst As String * 1, load1 As String * 1
Dim array_buf, com1 As Integer, buf1 As String
Dim buf As String, loct As String * 6, now2 As String * 3, now3 As String * 1, moti As String * 1
Dim awno1 As String * 1, stv As String * 6, cnt As Integer, err2 As Integer, palt2 As Integer, lamp1 As Integer
  On Error GoTo err1_rtn
  rd_data_cran = False
  If Left(project_name, 2) = "SP" Then
    awno1 = "1"
  Else
    If awno2 >= "1" And awno2 <= "9" Then awno1 = awno2 Else awno1 = get_awno(port_id)
  End If
  Call sio_flush(port_id, 2): lib_buf = ""
  wbuf = "02" & Chr(13) & Chr(10)
  Call sio_write(port_id, wbuf, 4)
  Call Wait(0.5): LL = sio_read(port_id, lib_buf, 22)
  If LL <> 22 Then
     If Left(project_name, 4) = "SK02" Then
       If Left(get_string_value1("", "", "pa_ccl3"), 3) = "log" Then
         jj = LL: Call syslog1("rd_data_cran", 22, jj, "awno=" & awno1 & ",buf=" & Left(lib_buf, 22))
       End If
       err1 = cran2_cnt(Val(awno1))
       If cran2_cnt(Val(awno1)) <= 2 Then
         cran2_cnt(Val(awno1)) = cran2_cnt(Val(awno1)) + 1: Exit Function
       End If
     End If
     ii = get_integer_value1(awno1, mveq, "cr_err")
     lib_buf = "cran read err,length=" & LL & ",buf=" & Left(lib_buf, 22)
     jj = LL: Call syslog("Cran_rd_data_cran", 22, jj, lib_buf)
     If Left(project_name, 4) = "SK04" And Left(mveq, 2) = "#3" Then
       If LL <> 0 Then Call syslog1("Cran_rd_data_cran3", 22, ii, lib_buf)
     Else
       If LL = 0 Then
          DoEvents 'Call wait1_100ms(port_id,1)
          If ii <> 67 Then rcn.Execute "update cr set cr_err=67 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
       Else
          DoEvents 'Call wait1_100ms(port_id, 1)
          If Left(project_name, 4) <> "SP01" Then
            If ii <> 65 Then rcn.Execute "update cr set cr_err=65 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
          End If
       End If
     End If
     Exit Function
  End If
  cran2_cnt(Val(awno1)) = 0
  If check_ascii(port_id, lib_buf, 20, "show") = False Then Exit Function
  brs5.Open "select * from cr where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'", rcn, adOpenKeyset, adLockOptimistic
  If brs5.EOF Then
     Set brs5 = Nothing
     Call syslog("rd_data_cran", 0, 0, "搬運設備" & mveq & "不存在設備檔中")
     Call Wait(0.1): Exit Function
  End If
  If Left(lib_buf, 20) <> cran_buf Then
    'cran_buf = Left(lib_buf, 20): If awno1 <> "1" And awno1 <> "5" Then Call syslog1("", 0, 0, cran_buf)
  End If
  buf = lib_buf
  auto1 = Mid(buf, 1, 1): load1 = Mid(buf, 10, 1): mvst = Mid(buf, 2, 1)
  If auto1 = "4" Then
    auto1 = "1": lamp1 = 1 'SP case使用lamp1
  Else
    auto1 = "0": lamp1 = 2 'SP case使用lamp1
  End If
  If brs5!cr_auto <> auto1 Then rcn.Execute "update cr set cr_auto='" & auto1 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  If brs5!cr_load <> load1 Then rcn.Execute "update cr set cr_load='" & load1 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  If Left(project_name, 2) = "SP" Then
    If load1 = "1" Then '記錄天車接受命令後曾有載過,控制cw完成用
      If brs5!cr_load1 <> load1 Then
        rcn.Execute "update cr set cr_load1='" & load1 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
        If Left(project_name, 2) = "SP" Then
          'rcn.Execute "insert into cs select * from cr where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
          'rcn.Execute "update cs set cs_enc1='" & get_now() & "' where cs_awno='" & awno1 & "' and cs_mveq='" & mveq & "' and cs_palt=" & brs5!cr_palt
          If brs5!cr_palt = 0 Then
            Call syslog1("rd_data_cran_err", brs5!cr_palt, 0, brs5!cr_mveq & "," & brs5!cr_from & "-->" & brs5!cr_to)
          Else
            Call syslog1("rd_data_cran", brs5!cr_palt, 0, brs5!cr_mveq & "," & brs5!cr_from & "-->" & brs5!cr_to)
          End If
        End If
      End If
    End If
  End If
  'loct(6):目前走行位置,now2(3):目前昇降位置,now3(1):目前叉牙位置,moti(1):目前天車動作
  'Mid(buf, 3, 3)='000***'表停在站的位置
  If Left(project_name, 4) = "SK04" And (Mid(buf, 3, 6) = "000001" Or Mid(buf, 3, 6) = "000002") Then 'A66及A69入出庫站位置
    loct = "027   "
  Else
    If Mid(buf, 3, 3) = "000" Then 'Mid(buf, 3, 3)='000***'表停在站的位置
      loct = Mid(buf, 3, 6)
    Else
      loct = Mid(buf, 3, 3) & "   "
    End If
  End If
  If brs5!cr_loct <> loct Then rcn.Execute "update cr set cr_loct='" & loct & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  now2 = Mid(buf, 6, 3): now3 = Mid(buf, 13, 1)
  If brs5!cr_now2 <> now2 Then rcn.Execute "update cr set cr_now2='" & now2 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  If brs5!cr_now3 <> now3 Then rcn.Execute "update cr set cr_now3='" & now3 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  err1 = Val(Mid(buf, 17, 4))
  If mvst = "0" Or mvst = "2" Then mvst = MV_NONE
  If auto1_old <> auto1 And auto1 = "0" Then
    'Call syslog1("auto11", brs5!cr_palt, err1, "auto1=" & auto1_old & "," & auto1 & ",mvst=" & brs5!cr_mvst & "," & mvst): auto1_old = auto1
'2004/08/02 21:51:37 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=14006,command = 010102401300000004,行號為:0
'2004/08/02 22:52:42 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=14006,command = 03,行號為:0
'2004/08/02 22:52:43 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=14006,command = 010000000000000004,行號為:0
'2004/08/02 22:55:56 [CRAN5-auto11] (14006,3097) auto1=1,0,mvst=3,3,行號為:0'
'3097:叉牙運動中載物超出左右側
'2004/08/02 22:59:10 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=14006,command = 03,行號為:0
'2004/08/02 22:59:11 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=14006,command = 010000000000000004,行號為:0
'2004/08/02 22:59:57 [CRAN5-auto12] (14006,0) auto1=0,1,mvst=1,0,行號為:0
  
'2004/07/31 13:41:21 [CRAN5-auto12] (13654,0) auto1=1,1,mvst=1,0,行號為:0
'2004/07/31 13:44:02 [CRAN5-auto11] (0,0) auto1=1,0,mvst=0,0,行號為:0
'2004/07/31 13:59:49 [CRAN5-auto11] (0,3097) auto1=,0,mvst=3,3,行號為:0
'2004/07/31 14:06:50 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=13655,command = 010205100100000001,行號為:0
'2004/07/31 14:08:36 [CRAN5-auto12] (13655,0) auto1=0,1,mvst=1,0,行號為:0
  
'2004/07/31 21:03:46 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=13708,command = 010000000702006013,行號為:0
'2004/07/31 21:05:41 [CRAN5-auto11] (13708,0) auto1=1,0,mvst=3,0,行號為:0
'2004/07/31 23:39:49 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=13708,command = 03,行號為:0
'2004/07/31 23:39:50 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=13709,command = 010200101200000006,行號為:0
'2004/07/31 23:42:46 [CRAN5-auto12] (13709,0) auto1=0,1,mvst=1,0,行號為:0
'2004/07/31 23:45:33 [CRAN5-plc_write_asrs] (4,4) ok,mveq=#5天車      ,palt=13710,command = 010201201200000006,行號為:0
'2004/07/31 23:48:07 [CRAN5-auto12] (13710,0) auto1=1,1,mvst=1,0,行號為:0
  End If
  '          舊搬運狀態=搬運中 And 新搬運狀態=無搬運
  If Left(project_name, 2) = "SK" Or Left(project_name, 2) = "SU" Then
  'If Left(project_name, 4) = "SK04" Or Left(project_name, 4) = "SK05" Or Left(project_name, 2) = "SU" Then
    If brs5!cr_mvst = MV_NOW And mvst = MV_NONE Then
      '搬運完成,轉移搬運狀態,cw_stat由"2"變為"0"
      'Call syslog1("auto12", brs5!cr_palt, err1, "auto1=" & auto1_old & "," & auto1 & ",mvst=" & brs5!cr_mvst & "," & mvst): auto1_old = auto1
      If auto1 = "1" Then
        brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='2'", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs3.EOF Then
          If err1 < 102 Then
             lib_buf = "cran creq_mvst: mveq=" & mveq & ",palt=" & brs5!cr_palt: Call syslog("rd_data_cran", 0, 0, lib_buf)
             If brs5!cr_palt = 0 Then
                Call syslog("rd_data_cran:%s", 3, 0, "mveq=" & mveq & ",palt=0 err")
             Else
                Call increase_cnt("cr_cnt", brs5!cr_mveq, brs5!cr_from)
                GoTo rtn10
                If False And Left(brs3!cw_nxt, 3) >= "B11" And Left(brs3!cw_nxt, 3) <= "B61" Then
                  '針對STV
                   cnt = 0
                   Do
                      If get_string_value1(awno1, Left("STV" & (Val(Mid(brs3!cw_nxt, 2, 1)) + 4) & "       ", 10), "cr_load") <> "0" Then Exit Do
                      If cnt = 20 Then
                         Call syslog("", 0, 0, "STV=" & (Val(Mid(brs3!cw_nxt, 2, 1)) + 4) & "卸載 timeout")
                         If get_integer_value1(awno1, Left(brs3!cw_nxt, 3), "st_err") <> 68 Then rcn.Execute "update st set st_err=68 where st_awno='" & awno1 & "' and st_stno='" & Left(brs3!cw_nxt, STNO_NUM) & "'"
                         Set brs3 = Nothing: Set brs5 = Nothing
                         Exit Function
                      End If
                      cnt = cnt + 1: Call Wait(0.5) 'Call wait1_100ms(port_id, 5)
                   Loop
                ElseIf is_stno(awno1, Left(brs3!cw_nxt, STNO_NUM)) And get_string_value1(awno1, Left(brs3!cw_nxt, STNO_NUM), "st_stat") <> "V" Then
                   cnt = 0
                   Do
                     'If get_integer_value1(awno1, Left(brs3!cw_nxt, STNO_NUM), "st_err") <> 11 Then rCn.Execute "update st set st_err=11 where st_stno='" & Left(brs3!cw_nxt, STNO_NUM) & "'"
                      If get_string_value1(awno1, Left(brs3!cw_nxt, STNO_NUM), "st_load") <> "0" Then Exit Do
                      If cnt = 20 Then
                         Call syslog("", 0, 0, "stno=" & Left(brs3!cw_nxt, STNO_NUM) & "卸載 timeout")
                         If get_integer_value1(awno1, Left(brs3!cw_nxt, STNO_NUM), "st_err") <> 68 Then rcn.Execute "update st set st_err=68 where st_awno='" & awno1 & "' and st_stno='" & Left(brs3!cw_nxt, STNO_NUM) & "'"
                         If Left(project_name, 2) = "SJ" Then
                            If Left(brs3!cw_nxt, 3) = "A28" Then GoTo rtn10
                         End If
                         Set brs3 = Nothing: Set brs5 = Nothing
                         Exit Function
                      End If
                      cnt = cnt + 1: Call Wait(0.5)
                   Loop
                End If
rtn10:
                'rCn.Execute "update st set st_err=0 where st_stno='" & Left(brs3!cw_nxt, STNO_NUM) & "'"
                If Left(project_name, 4) = "SK04" Then
                   If Left(mveq, 2) = "#1" Then
                      If Mid(brs3!cw_nxt, 4, 2) = "25" Or Mid(brs3!cw_nxt, 4, 2) = "26" Or Mid(brs3!cw_nxt, 1, 3) = "A69" Then '024位置安全
                         Call put_cw(awno1, TR_MOVE_STNO, "(車上)", "00024" & get_string_value1(awno1, EQ_CRAN1, "cr_now2"), loginname, "Y")
                      End If
                   ElseIf Left(mveq, 2) = "#4" Then
                      If Left(brs3!cw_nxt, 3) = "D20" And Left(brs3!cw_nxt, 3) <> Left(brs3!cw_to, 3) Then
                         'Call put_cw(awno1, TR_MOVE_STNO, "D20     ", brs3!cw_to, loginname, "N")
                      End If
                   End If
                End If
                If is_stno(brs3!cw_awno, Left(brs3!cw_nxt, STNO_NUM)) And get_string_value1(brs3!cw_awno, Left(brs3!cw_nxt, STNO_NUM), "st_stat") <> "V" Then
                  'plc更新太慢,卸載一完成mapath立刻更新成有載狀態
                  rcn1.Execute "update st set st_load='1' where st_stno = '" & Left(brs3!cw_nxt, STNO_NUM) & "'"
                End If
                rcn.Execute "update cw set cw_stat='0' where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cr_palt
                rcn.Execute "update cr set cr_palt=0 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
             End If
          End If
        End If
        Set brs3 = Nothing
      End If
    End If
  Else
    'If Left(project_name, 4) = "SP01" And (Left(mveq, 2) >= "#1" And Left(mveq, 2) <= "#3") Then
    If mvst = MV_NONE Then
      '搬運完成,轉移搬運狀態,cw_stat由"2"變為"0"
      If auto1 = "1" And load1 = "0" Then
        brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_palt=" & brs5!cr_palt & " and (cw_stat='1' or cw_stat='2')", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs3.EOF Then
          If err1 < 102 Then
             lib_buf = "cran creq_mvst: mveq=" & mveq & ",palt=" & brs5!cr_palt: Call syslog("rd_data_cran", 0, 0, lib_buf)
             If brs5!cr_palt = 0 Then
                Call syslog("rd_data_cran:%s", 3, 0, "mveq=" & mveq & ",palt=0 err")
             Else '只要brs5!cr_palt <> 0就會完成
                'GoTo rtn12
                If Left(project_name, 2) = "SP" Then
                  If Not is_lono(awno1, brs3!cw_now) Then '入庫
                    '命令已傳天車,天車異常時會造成誤動作,資料已入庫但實際仍未入庫
                    brs4.Open "select * from st where st_stno='" & Left(brs3!cw_now, STNO_NUM) & "'", rcn1, adOpenKeyset, adLockOptimistic
                    If Not brs4.EOF Then '命令已傳天車,天車異常時會造成誤動作,資料已入庫但實際仍未入庫
                      If brs4!st_palt1 <> brs5!cr_palt Then '棧板移走後,應該brs4!st_palt1=0
                        Set brs4 = Nothing: GoTo rtn12 '可完成
                      Else
                        Call syslog("rd_data_cran:", brs4!st_palt1, brs5!cr_palt, "mveq=" & mveq & ",入庫序號一致所以不能完成")
                        Set brs4 = Nothing
                      End If
                    Else
                      Set brs4 = Nothing: GoTo rtn12
                    End If
                  Else
                    GoTo rtn12
                  End If
                Else
rtn12:
                  If brs5!cr_load1 = "1" Then '記錄天車接受命令後曾有載過,控制cw完成用
                    Call increase_cnt("cr_cnt", brs5!cr_mveq, brs5!cr_from)
                    rcn.Execute "update cw set cw_stat='0' where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cr_palt
                    rcn.Execute "update cr set cr_palt=0,cr_load1='0' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
                  End If
                End If
             End If
          End If
        End If
        Set brs3 = Nothing
      End If
    End If
  End If
  If brs5!cr_mvst <> mvst Then rcn.Execute "update cr set cr_mvst='" & mvst & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
     '新的設備異常碼 <> 0 And 新的設備異常碼 <> 舊的設備異常碼,在er檔中寫入一筆
  If err1 <> brs5!cr_err Then
     rcn.Execute "update cr set cr_err=" & err1 & " where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
     rcn.Execute "update cw set cw_err=" & err1 & " where cw_mveq='" & mveq & "' and cw_stat='2'"
     If Left(project_name, 2) = "SJ" Then
       'LED詳細說明請參考SJ900.bas
        If err1 <> 0 Then
           ii = sio_write(3, Chr(2) & Chr(127) & Chr(33) & Chr(42) & Chr(32) & "天車異常,異常碼:" & err1 & Chr(3) & Chr(0), 26)
           lib_buf1 = Mid(Err_Translate(err1), 4, 20) '左方的4若改為6會有問題
           ii = sio_write(3, Chr(2) & Chr(127) & Chr(34) & Chr(42) & Chr(32) & lib_buf1 & Chr(3) & Chr(0), 26)
           rcn.Execute "update pa set pa_err=" & err1
        Else
           ii = sio_write(3, Chr(2) & Chr(127) & Chr(33) & Chr(42) & Chr(32) & "歡迎使用自動倉庫系統" & Chr(3) & Chr(0), 26)
           ii = sio_write(3, Chr(2) & Chr(127) & Chr(34) & Chr(42) & Chr(32) & " 全廠自動化控制系統 " & Chr(3) & Chr(0), 26)
           rcn.Execute "update pa set pa_err=0"
        End If
     ElseIf Left(project_name, 2) = "SK" Then
       If Left(project_name, 4) = "SK04" Or Left(project_name, 4) = "SK05" Then
          brs3.Open "select pa_led from pa", rcn1, adOpenKeyset, adLockOptimistic
          Call InitialComPort(brs3(0), 9600, False, 8, 1)
          If err1 <> 0 Then
             lib_buf = Err_Translate(err1): jj = 0
             For ii = 1 To 40
                buf1 = Mid(lib_buf, ii, 1): kk = Asc(buf1)  '中文字"完"讀到kk=-22599,"一"讀到kk=-23488
                If Not (kk > 0 And kk < 255) Then jj = jj + 2 Else jj = jj + 1
                If jj >= 39 Then Exit For
             Next ii
             Call SendMessage(Mid(lib_buf, 1, ii), brs3(0), 1, 40)
          Else
             Call SendMessage("歡迎使用玻纖布廠" & get_awnm(awno1) & Mid(get_now(), 3, 8), brs3(0), 1, 40)
          End If
          Call Wait(1) '務必作
          '''frmMain.mscLED(brs3(0) - 1).PortOpen = False
          Set brs3 = Nothing
       End If
     ElseIf Left(project_name, 2) = "SP" Then
        If Mid(brs5!cr_mveq, 2, 1) >= "1" And Mid(brs5!cr_mveq, 2, 1) < "4" Then
          com1 = Val(Mid(brs5!cr_mveq, 2, 1)) + 2
        ElseIf Mid(brs5!cr_mveq, 2, 1) = "4" Or Mid(brs5!cr_mveq, 2, 1) = "5" Then
          com1 = 6
        Else
          com1 = Val(Mid(brs5!cr_mveq, 2, 1)) + 1
        End If
        If err1 = 0 Then
          lib_buf = Mid(brs5!cr_mveq, 1, 2) & "天車異常已排除": rcn.Execute "update le set le_message='" & lib_buf & "' where le_com=" & com1 & " and le_mode=2"
        Else
          lamp1 = 3
          lib_buf = Err_Translate(err1): jj = 0
          For ii = 1 To 40
            buf1 = Mid(lib_buf, ii, 1): kk = Asc(buf1)  '中文字"完"讀到kk=-22599,"一"讀到kk=-23488
            If Not (kk > 0 And kk < 255) Then jj = jj + 2 Else jj = jj + 1
            If jj >= 39 Then Exit For
          Next ii
          Mid(lib_buf, 1, 2) = Mid(brs5!cr_mveq, 1, 2) & "  "
          rcn1.Execute "update le set le_message='" & Mid(lib_buf, 1, ii) & "' where le_com=" & com1 & " and le_mode=2"
          If err1 = 3000 Then
            If Left(brs5!cr_to, STNO_NUM) = "C014" Or Left(brs5!cr_to, STNO_NUM) = "C022" Or Left(brs5!cr_to, STNO_NUM) = "C026" Then
              Call handle_3000_3001(awno1, brs5!cr_palt, 3000)
            End If
          End If
          If err1 = 3001 Or err1 = 3002 Then Call handle_3000_3001(awno1, brs5!cr_palt, 3001)
        End If
     Else
       If err1 <> 0 Then
         'If err1 = 3000 Then Call handle_3000_3001(awno1, brs5!cr_palt, 3000)
         If err1 = 3001 Or err1 = 3002 Then Call handle_3000_3001(awno1, brs5!cr_palt, 3001)
       End If
     End If
     If err1 <> 0 And err1 >= 3000 Then
        brs3.Open "select * from er where er_awno='" & awno1 & "' and er_palt=" & brs5!cr_palt & " and er_err=" & err1, rcn1, adOpenKeyset, adLockOptimistic
        If brs3.EOF Then
          Call increase_cnt("cr_erno", brs5!cr_mveq, brs5!cr_from)
          '2002/10/03發生The statement has been terminated異常
          err1_place = "insert_er"
          rcn.Execute "insert into er values('" & awno1 & "','" & mveq & "'," & brs5!cr_palt & ",'" & brs5!cr_from & "','" & brs5!cr_to & "','" & mvst & "','" & brs5!cr_fbst & "','" & auto1 & "','" & loct & "','" & load1 & "'," & err1 & ",'" & get_now() & "')"
          err1_place = ""
        End If
        Set brs3 = Nothing
     End If
  End If
  If brs5!cr_rest >= "2" And brs5!cr_rest <= "8" Then
     If plc_write_asrs(port_id, 4, "04" & Chr(13) & Chr(10), mveq, 4, awno1) = False Then
        Set brs5 = Nothing: rd_data_cran = True: Exit Function
     End If
     If brs5!cr_rest = "2" Then '處理異常復歸
        If Left(project_name, 2) = "SJ" Then
           ii = sio_write(3, Chr(2) & Chr(127) & Chr(33) & Chr(42) & Chr(32) & "歡迎使用自動倉庫系統" & Chr(3) & Chr(0), 26)
           ii = sio_write(3, Chr(2) & Chr(127) & Chr(34) & Chr(42) & Chr(32) & " 全廠自動化控制系統 " & Chr(3) & Chr(0), 26)
           rcn.Execute "update pa set pa_err=0"
        End If
        lib_buf = "03" & Chr(13) & Chr(10): ii = 4
     ElseIf brs5!cr_rest = "3" Then '點亮CCD監視器燈光
        lib_buf = "0701" & Chr(13) & Chr(10): ii = 6
     ElseIf brs5!cr_rest = "4" Then '熄滅CCD監視器燈光
        lib_buf = "0702" & Chr(13) & Chr(10): ii = 6
     ElseIf brs5!cr_rest >= "5" And brs5!cr_rest <= "8" Then '天車三軸歸位
        lib_buf = "060" & (Val(brs5!cr_rest) - 4) & Chr(13) & Chr(10): ii = 6
     End If
     If plc_write_asrs(port_id, ii, lib_buf, mveq, 4, awno1) = True Then
'      Set brs5 = Nothing: Exit Function
'    End If
        rcn.Execute "update cr set cr_rest='" & "0" & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
     End If
  End If
  If Left(project_name, 2) = "SP" Then
    If lamp(Val(Mid(mveq, 2, 1))) <> lamp1 Then
      If update_palt1_stno1("1", "800" & Mid(mveq, 2, 1), lamp1) Then Call syslog("前lamp", lamp(Val(Mid(mveq, 2, 1))), lamp1, "")
      If update_palt1_stno1("2", "900" & Mid(mveq, 2, 1), lamp1) Then Call syslog("後lamp", lamp(Val(Mid(mveq, 2, 1))), lamp1, "")
      lamp(Val(Mid(mveq, 2, 1))) = lamp1
    End If
  ElseIf Left(project_name, 4) = "SK04" And Left(mveq, 2) = "#1" Then
     If Mid(brs5!cr_from, 4, 2) = "25" Or Mid(brs5!cr_from, 4, 2) = "26" Or Mid(brs5!cr_to, 4, 2) = "25" Or Mid(brs5!cr_to, 4, 2) = "26" Or Mid(brs5!cr_to, 1, 3) = "A69" Then
       brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='2' and cw_err <> 3014 and cw_palt = " & brs5!cr_palt, rcn1, adOpenKeyset, adLockOptimistic
       'brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='2' and cw_palt <> " & brs5!cr_palt, rcn, adOpenKeyset, adLockOptimistic
       If Not brs3.EOF Then
         'If Mid(brs3!cw_now, 4, 2) = "25" Or Mid(brs3!cw_now, 4, 2) = "26" Or Mid(brs3!cw_nxt, 4, 2) = "25" Or Mid(brs3!cw_nxt, 4, 2) = "26" Or Mid(brs3!cw_nxt, 1, 3) = "A69" Then
         'If Mid(brs3!cw_now, 4, 2) = "25" Or Mid(brs3!cw_now, 4, 2) = "26" Or Mid(brs3!cw_now, 1, 3) = "A66" Or Mid(brs3!cw_nxt, 4, 2) = "25" Or Mid(brs3!cw_nxt, 4, 2) = "26" Or Mid(brs3!cw_nxt, 1, 3) = "A69" Then
'2004/07/14 04:58:05 [CRAN1-rd_data_cran] (5312,0) now=A71     ,A90=0,A91=0,行號為:0
'2004/07/14 04:58:08 [CRAN1-rd_data_cran] (5312,0) now=A71     ,A90=0,A91=0,行號為:0
'2004/07/14 04:58:10 [CRAN1-rd_data_cran] (5312,0) now=A71     ,A90=0,A91=0,行號為:0

'2004/07/14 04:58:13 [CRAN1-rd_data_cran] (5312,0) now=A71     ,A90=3,A91=1,行號為:0
'2004/07/14 04:58:14 [CRAN1-rd_data_cran] (5312,1) now=A71     ,A90=3,A91=1,行號為:0
'2004/07/14 04:58:15 [CRAN1-rd_data_cran] (5312,2) now=A71     ,A90=1,A91=0,行號為:0
'2004/07/14 04:58:16 [CRAN1-rd_data_cran] (5312,3) now=A71     ,A90=1,A91=0,行號為:0
'2004/07/14 04:58:17 [CRAN1-rd_data_cran] (5312,4) now=A71     ,A90=1,A91=0,行號為:0
'2004/07/14 04:58:17 [PLC1-A90] (0,0) 0010,行號為:0
'2004/07/14 04:58:17 [PLC1-put_cw] (0,0) new_srid=站間搬運      ,srid=I,palt=5313,from=A92     ,to=A92     ,行號為:0
'2004/07/14 04:58:18 [CRAN1-rd_data_cran] (5312,5) now=A71     ,A90=0,A91=0,行號為:0
'2004/07/14 04:58:18 [CRAN1-rd_data_cran1] (5312,3014) 天車緊急停止,ii=5,now=A71     ,nxt=A69     ,行號為:0
'2004/07/14 04:58:18 [CRAN1-plc_write_asrs] (4,4) ok,mveq=#1天車      ,palt=5312,command = 05,行號為:0

           
           ii = 0
           Do
             If ii >= 5 Then Exit Do
             lib_buf = get_string_value1("", "", "pa_a90")
             If Mid(lib_buf, 1, 1) = "0" Or Mid(lib_buf, 2, 1) = "0" Then
             'If get_string_value1(awno1, "A90", "st_load") = "0" And get_string_value1(awno1, "A91", "st_load") = "0" Then
               Exit Do
             End If
             'Call syslog1("rd_data_cran", brs3!cw_palt, ii, "now=" & brs3!cw_now & ",A90-A93=" & Mid(lib_buf, 1, 4))
             ii = ii + 1: Call Wait(1)
           Loop
           'Call syslog1("rd_data_cran1", brs3!cw_palt, ii, "now=" & brs3!cw_now & ",A90-A93=" & Mid(lib_buf, 1, 4))
           If ii >= 5 Then
             Call syslog1("rd_data_cran2", brs3!cw_palt, 3014, "天車緊急停止,ii=" & ii & ",now=" & brs3!cw_now & ",nxt=" & brs3!cw_nxt)
             lib_buf = "05" & Chr(13) & Chr(10): ii = 4 '天車緊急停止
             Call plc_write_asrs(port_id, ii, lib_buf, mveq, 4, awno1)
             GoTo rtn20
           End If
           'If get_string_value1(awno1, "A90", "st_load") <> NOLD Or get_string_value1(awno1, "A91", "st_load") <> NOLD Then
           '  Call syslog1("rd_data_cran", brs3!cw_palt, 0, "天車緊急停止,now=" & brs3!cw_now & ",nxt=" & brs3!cw_nxt)
           '  lib_buf = "05" & Chr(13) & Chr(10): ii = 4 '天車緊急停止
           '  Call plc_write_asrs(port_id, ii, lib_buf, mveq, 4, awno1)
           '  GoTo rtn20
           'End If
           brs4.Open "select * from cw where cw_mveq='ROTATE    ' and cw_stat='2'", rcn1, adOpenKeyset, adLockOptimistic
           If Not brs4.EOF Then
             Call syslog1("rd_data_cran", brs3!cw_palt, 3014, "天車緊急停止ROTATE,now=" & brs3!cw_now & ",nxt=" & brs3!cw_nxt)
             lib_buf = "05" & Chr(13) & Chr(10): ii = 4 '天車緊急停止
             Call plc_write_asrs(port_id, ii, lib_buf, mveq, 4, awno1)
           End If
           Set brs4 = Nothing
         'End If
       End If
rtn20:
       Set brs3 = Nothing
     End If
  End If
  Set brs5 = Nothing
  If Left(project_name, 2) = "SJ" Then
     If get_string_value1(awno1, "A40 ", "st_load") <> old_a40_load Then
        old_a40_load = get_string_value1(awno1, "A40 ", "st_load")
        If old_a40_load = "1" Then '防突A40
           Call plc_write_asrs(port_id, 4, "05" & Chr(13) & Chr(10), mveq, 4, awno1)
        End If
     End If
  End If
  rd_data_cran = True
  Exit Function
err1_rtn:
  Set brs3 = Nothing: Set brs4 = Nothing: Set brs5 = Nothing
  If Err.Number = -2147467259 Then   'SQL deadlock case,或[DBNETLIB]ConnectionOpen: SQL Server不存在或拒絕存取,或連線失敗(for SO專案Save程式)
    Call syslog1("rd_data_cran", 999, 0, prog & ":系統異常,行號為:" & Erl & ",異常碼:" & Str(Err.Number) & Chr(13) & Err.Description)
    Call Wait(2)
  ElseIf Left(project_name, 4) = "SK04" And awno1 = "3" Then
    If Err.Number = -2147217833 Then 'String or binary data would be truncated
      ii = LL: Call syslog1("", 22, ii, "err=" & Err.Number & ",buf=" & Left(lib_buf, 20) & ",String or binary data would be truncated", "show")
    End If
  Else
    If err1_place = "insert_er" Then
       Call syslog1(err1_place, 0, 0, "test")
       Call syslog1(err1_place, 0, 0, awno1 & "," & mveq & "," & brs5!cr_palt & "," & brs5!cr_from & "," & brs5!cr_to & "," & mvst & "," & brs5!cr_fbst & "," & auto1 & "," & loct & "," & load1 & "," & err1 & "," & get_now())
       err1_place = ""
    Else
       Call err2_rtn("rd_data_cran")
    End If
  End If
End Function
Public Function plc_read(port_id As Integer, addr As String, data1 As String, leng As Integer, mveq As String) As Boolean
  Dim awno1 As String, buf As String, LL As Long, wbuf As String, ii As Integer, jj As Integer, lib_buf As String * 300
  On Error GoTo err1_rtn
  awno1 = "1"
  wbuf = "00FFWR0" & addr & Right("0" & CStr(Hex(leng)), 2) 'addr=R0704
  buf = checksum(wbuf, 14)
  wbuf = Chr(5) & wbuf & Left(buf, 2) & Chr(13) & Chr(10)
  Call sio_flush(port_id, 2)
  Call sio_write(port_id, wbuf, 19)
  Call Wait(0.5): lib_buf = "": LL = sio_read(port_id, lib_buf, (leng * 4 + 10))
  If Left(lib_buf, 1) = Chr(21) Then 'NAK
    Response = MsgBox("plc_read:PLC回NAK異常,port_id=" & port_id & ",addr=" & addr & ",leng=" & (leng * 4 + 10) & ",leng1=" & LL & ",err code=" & Mid(lib_buf, 6, 2) & ",mveq=" & Left(mveq, 10))
    plc_read = False:  Exit Function
  End If
  buf = checksum(Mid(lib_buf, 2, (leng * 4 + 5)), (leng * 4 + 5))
  If LL <> (leng * 4 + 10) Or Mid(lib_buf, (leng * 4 + 7), 2) <> Left(buf, 2) Then
    If LL <> (leng * 4 + 10) Then
      ii = get_integer_value1(awno1, mveq, "cr_err")
      If LL = 0 Then
        If ii <> 67 Then rcn1.Execute "update cr set cr_err=67 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
      Else
        If ii <> 65 Then rcn1.Execute "update cr set cr_err=65 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
      End If
      Response = MsgBox("plc_read:與PLC交訊異常,port_id=" & port_id & ",addr=" & addr & ",leng=" & (leng * 4 + 10) & ",leng1=" & LL & ",mveq=" & Left(mveq, 10))
    Else
      Response = MsgBox("plc_read:checksum異常,port_id=" & port_id & ",addr=" & addr & ",leng=" & (leng * 4 + 10) & ",leng1=" & LL & ",mveq=" & Left(mveq, 10))
    End If
' chr(13) :CR
    Call sio_write(port_id, Chr(21) & "00FF" & Chr(13) & Chr(10), 7)
    plc_read = False:  Exit Function
  End If
  Call sio_write(port_id, Chr(6) & "00FF" & Chr(13) & Chr(10), 7)
  data1 = Mid(lib_buf, 6, leng * 4)
  plc_read = True
  Exit Function
err1_rtn:
  plc_read = False
  Call err2_rtn("plc_read")
End Function
Public Function plc_write(port_id As Integer, addr As String, data1 As String) As Boolean
  Dim buf As String, LL As Long, wbuf As String, ii As Integer, jj As Integer, lib_buf As String * 300, leng As Integer
  On Error GoTo err1_rtn
  'buf = checksum("00FFBR3ABCD", 11)
  ii = Len(data1) - 1: leng = ii / 4 'addr=R0704
  wbuf = "00FFWW0" & addr & Right("0" & CStr(Hex(leng)), 2) & Left(data1, ii)
  buf = checksum(wbuf, ii + 14)
  wbuf = Chr(5) & wbuf & Left(buf, 2) & Chr(13) & Chr(10)
  Call sio_flush(port_id, 2)
  Call sio_write(port_id, wbuf, ii + 19)
  Call Wait(0.3): lib_buf = "": LL = sio_read(port_id, lib_buf, 9)
  If Left(lib_buf, 1) = Chr(21) Then 'NAK,LL=9
    Response = MsgBox("plc_write:PLC回NAK異常,port_id=" & port_id & ",addr=" & addr & ",leng=" & leng & ",leng1=" & LL & ",err code=" & Mid(lib_buf, 6, 2))
    plc_write = False: Exit Function
  End If
  If LL <> 7 Then
    Response = MsgBox("plc_write:與PLC交訊異常,port_id=" & port_id & ",addr=" & addr & ",leng=" & leng & ",leng=" & LL)
    plc_write = False: Exit Function
  ElseIf Left(lib_buf, 1) <> Chr(6) Then '非ACK
    Response = MsgBox("plc_write:PLC不回ACK異常,port_id=" & port_id & ",addr=" & addr & ",leng=" & leng & ",leng=" & LL)
    plc_write = False: Exit Function
  End If
  plc_write = True
  Exit Function
err1_rtn:
  plc_write = False
  Call err2_rtn("plc_write")
End Function
Public Function checksum(buf As String, leng As Integer) As String
  Dim ii As Integer, sum1
  sum1 = 0
  For ii = 1 To leng
    sum1 = sum1 + Asc(Mid(buf, ii, 1))
  Next ii
  checksum = CStr(Hex(sum1))
  checksum = Right(CStr(Hex(sum1)), 2)
End Function
'葉樹堅吊車
Public Function rd_data_cran_new(port_id As Integer, mveq As String) As Boolean
Dim brs5 As New ADODB.Recordset, brs3 As New ADODB.Recordset, brs4 As New ADODB.Recordset
Dim err1 As Integer, ii As Integer, jj As Integer, kk As Integer, lib_buf As String * 300, lib_buf1 As String * 300
Dim LL As Long, wbuf As String
Dim auto1 As String * 1, mvst As String * 1, old_mvst As String * 1, load1 As String * 1
Dim array_buf
Dim buf As String, loct As String * 6, now2 As String * 3, now3 As String * 1, moti As String * 1
Dim awno1 As String * 1, stv As String * 6, cnt As Integer, err2 As Integer, palt2 As Integer
  On Error GoTo err1_rtn
  rd_data_cran_new = False
  awno1 = "1"
  If plc_read(port_id, "D0170", lib_buf, 11, mveq) = False Then Exit Function
  brs5.Open "select * from cr where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'", rcn, adOpenKeyset, adLockOptimistic
  If brs5.EOF Then
     Set brs5 = Nothing
     Call syslog("rd_data_cran_new", 0, 0, "搬運設備" & mveq & "不存在設備檔中")
     Call Wait(0.1): Exit Function
  End If
  buf = lib_buf
  auto1 = Mid(buf, 4 * (1 - 1) + 4, 1): load1 = Mid(buf, 4 * (5 - 1) + 4, 1): mvst = Mid(buf, 4 * (2 - 1) + 4, 1)
  If auto1 = "4" Then auto1 = "1" Else auto1 = "0"
  If brs5!cr_auto <> auto1 Then rcn.Execute "update cr set cr_auto='" & auto1 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  If brs5!cr_load <> load1 Then rcn.Execute "update cr set cr_load='" & load1 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  'loct(6):目前走行位置,now2(3):目前昇降位置,now3(1):目前叉牙位置,moti(1):目前天車動作
  loct = Mid(buf, 4 * (3 - 1) + 2, 3) & "   "
  If brs5!cr_loct <> loct Then rcn.Execute "update cr set cr_loct='" & loct & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  now2 = Mid(buf, 4 * (4 - 1) + 2, 3): now3 = Mid(buf, 4 * (8 - 1) + 4, 1)
  If brs5!cr_now2 <> now2 Then rcn.Execute "update cr set cr_now2='" & now2 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  If brs5!cr_now3 <> now3 Then rcn.Execute "update cr set cr_now3='" & now3 & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
  err1 = Val(Mid(buf, 4 * (9 - 1) + 1, 4))
  If mvst = "0" Or mvst = "2" Then mvst = MV_NONE
  If auto1_old <> auto1 And auto1 = "0" Then
  End If
'          舊搬運狀態=搬運中 And 新搬運狀態=無搬運
  If brs5!cr_mvst = MV_NOW And mvst = MV_NONE Then
    '搬運完成,轉移搬運狀態,cw_stat由"2"變為"0"
     If auto1 = "1" Then
       brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='2'", rcn1, adOpenKeyset, adLockOptimistic
       If Not brs3.EOF Then
          If err1 < 102 Then
             lib_buf = "cran creq_mvst: mveq=" & mveq & ",palt=" & brs5!cr_palt: Call syslog("rd_data_cran_new", 0, 0, lib_buf)
             If brs5!cr_palt = 0 Then
                Call syslog("rd_data_cran_new:%s", 3, 0, "mveq=" & mveq & ",palt=0 err")
             Else
                Call increase_cnt("cr_cnt", brs5!cr_mveq, brs5!cr_from)
                'rCn.Execute "update st set st_err=0 where st_stno='" & Left(brs3!cw_nxt, STNO_NUM) & "'"
                rcn.Execute "update cw set cw_stat='0' where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cr_palt
                rcn.Execute "update cr set cr_palt=0 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
             End If
          End If
       End If
       Set brs3 = Nothing
    End If
  End If
  If brs5!cr_mvst <> mvst Then rcn.Execute "update cr set cr_mvst='" & mvst & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
     '新的設備異常碼 <> 0 And 新的設備異常碼 <> 舊的設備異常碼,在er檔中寫入一筆
  If err1 <> brs5!cr_err Then
     rcn.Execute "update cr set cr_err=" & err1 & " where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
     rcn.Execute "update cw set cw_err=" & err1 & " where cw_mveq='" & mveq & "' and cw_stat='2'"
     If err1 <> 0 And err1 >= 3000 Then
        brs3.Open "select * from er where er_awno='" & awno1 & "' and er_palt=" & brs5!cr_palt & " and er_err=" & err1, rcn1, adOpenKeyset, adLockOptimistic
        If brs3.EOF Then
          Call increase_cnt("cr_erno", brs5!cr_mveq, brs5!cr_from)
          '2002/10/03發生The statement has been terminated異常
          err1_place = "insert_er"
          rcn.Execute "insert into er values('" & awno1 & "','" & mveq & "'," & brs5!cr_palt & ",'" & brs5!cr_from & "','" & brs5!cr_to & "','" & mvst & "','" & brs5!cr_fbst & "','" & auto1 & "','" & loct & "','" & load1 & "'," & err1 & ",'" & get_now() & "')"
          err1_place = ""
        End If
        Set brs3 = Nothing
     End If
  End If
  If brs5!cr_rest >= "2" And brs5!cr_rest <= "8" Then
    If plc_write(port_id, "D0160", "000A") = False Then
      Set brs5 = Nothing: rd_data_cran_new = True: Exit Function
    End If
    If brs5!cr_rest = "2" Then '處理異常復歸
      If plc_write(port_id, "D0161", "000B") = True Then
         rcn.Execute "update cr set cr_rest='" & "0" & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
      End If
    ElseIf brs5!cr_rest >= "5" And brs5!cr_rest <= "8" Then '天車三軸歸位
      If plc_write(port_id, "D0163", "000B") = True Then
         rcn.Execute "update cr set cr_rest='" & "0" & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
      End If
    End If
  End If
  Set brs5 = Nothing
  rd_data_cran_new = True: Exit Function
err1_rtn:
  If err1_place = "insert_er" Then
     Call syslog1(err1_place, 0, 0, "test")
     Call syslog1(err1_place, 0, 0, awno1 & "," & mveq & "," & brs5!cr_palt & "," & brs5!cr_from & "," & brs5!cr_to & "," & mvst & "," & brs5!cr_fbst & "," & auto1 & "," & loct & "," & load1 & "," & err1 & "," & get_now())
     err1_place = ""
  Else
     Call err2_rtn("rd_data_cran_new")
  End If
End Function

'葉樹堅吊車
Public Function st_comd_cran(port_id As Integer, mveq As String, mode As String, Optional awno2 As String) As Boolean
Dim err1 As Integer, ii As Integer, jj As Integer, lib_buf As String * 300, palt_first As Integer, palt_first_flag As Boolean
Dim brs5 As New ADODB.Recordset, brs3 As New ADODB.Recordset, from1, to1
Dim LL As Long, wbuf As String, palt1 As Integer, nxt1 As String
Dim awno1 As String * 1, array_buf
Dim d007 As Integer, d008 As Integer, d009 As Integer
  On Error GoTo err1_rtn
  palt_first_flag = False: palt_first = 0: st_comd_cran = True
  If Left(project_name, 2) = "SP" Then
    awno1 = "1"
  Else
    If awno2 >= "1" And awno2 <= "9" Then awno1 = awno2 Else awno1 = get_awno(port_id)
  End If
  brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='2' order by cw_mveq,cw_stat", rcn, adOpenKeyset, adLockOptimistic
  If Not brs5.EOF Then
     Set brs5 = Nothing: Exit Function
  End If
  Set brs5 = Nothing
  brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_now='(車上)' and cw_stat='1' order by cw_now", rcn, adOpenKeyset, adLockOptimistic
  If Not brs5.EOF Then GoTo rtn0
  Set brs5 = Nothing
'發生次數最多
'在下一行CRAN1及CRAN3同時發生,常發生
'異常碼: -2147467259,Your transaction (ProcessID#30及ProcessID#39) was deadlocked with another process and has been chosen as the deadlock victim.Rerun your transaction.
  brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_first='Y' and cw_stat='1' order by cw_sitm", rcn, adOpenKeyset, adLockOptimistic
  If Not brs5.EOF Then
    If Left(project_name, 2) = "SK" Then
      If Left(brs5!cw_now, 3) = "E12" Or Left(brs5!cw_now, 3) = "E16" Or Left(brs5!cw_now, 3) = "F02" Or Left(brs5!cw_nxt, 3) = "F00" Then GoTo rtn1
    End If
    palt_first = brs5!cw_palt: GoTo rtn0
  End If
rtn1:
  Set brs5 = Nothing
  If Left(mode, 3) = "B01" Then
     ii = ii
  End If
rtn2:
  If Left(mode, 3) = "000" Then '搬棧板至天車上並停住case
     brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_nxt='00000000'", rcn, adOpenKeyset, adLockOptimistic
     If Not brs5.EOF = True Then GoTo rtn20
     'brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_stno='0000' order by cw_mveq,cw_stat,cw_stno,cw_sitm", rcn, adOpenKeyset, adLockOptimistic
  Else
     brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_stno='" & Left(mode, STNO_NUM) & "' order by cw_mveq,cw_stat,cw_stno,cw_sitm", rcn, adOpenKeyset, adLockOptimistic
  End If
  Do
     If brs5.EOF = True Then 'cw_first='Y的工作若無法執行時務必要將cw_first改為'N',否則所有cw_first='N'的工作都不會執行
'      If palt_first <> 0 Then rCn.Execute "update cw set cw_first='N' where cw_palt=" & palt_first
        Set brs5 = Nothing: st_comd_cran = False: Exit Do
     End If
rtn0:
     If brs5!cw_opno = TR_TEST Then GoTo rtn
     brs3.Open "select * from cw where cw_mveq = '" & brs5!cw_mveq & "' and cw_stat='2'", rcn1, adOpenKeyset, adLockOptimistic
     If Not brs3.EOF Then
       If brs5!cw_err <> 49 Then rcn.Execute "update cw set cw_err=49 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
       Set brs3 = Nothing: GoTo rtn
     End If
     Set brs3 = Nothing
     If (brs5!cw_err >= 102 And brs5!cw_err <= 499) Or (brs5!cw_err >= 600 And brs5!cw_err <= 699) Or (brs5!cw_err >= 3000 And brs5!cw_err <= 3999) Then GoTo rtn
     err1 = check_status(awno1, brs5!cw_mveq, brs5!cw_now, brs5!cw_nxt)
     If err1 > 0 Then
        If Not (err1 = 69 And Left(brs5!cw_now, 1) = "(") Then
           If brs5!cw_err <> err1 Then rcn.Execute "update cw set cw_err=" & err1 & " where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           GoTo rtn 'Exit Do
        End If
     End If
     'check是否有其它作業cw_now="(車上)"(表示東西在車上情況)
     If Left(brs5!cw_now, 1) <> "(" Then
        brs3.Open "select * from cw where cw_now='(車上)' and cw_awno='" & awno1 & "' and cw_palt <> " & brs5!cw_palt & " and cw_mveq='" & mveq & "'", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs3.EOF Then
           If brs5!cw_err <> 2 Then rcn.Execute "update cw set cw_err=2 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           Set brs3 = Nothing: GoTo rtn 'Exit Do
        End If
        Set brs3 = Nothing
        If brs5!cw_first <> "Y" Then
          brs3.Open "select * from cw where cw_first='Y' and cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'", rcn1, adOpenKeyset, adLockOptimistic
          If Not brs3.EOF Then
            If brs5!cw_err <> 2 Then rcn.Execute "update cw set cw_err=2 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
            Set brs3 = Nothing: GoTo rtn 'Exit Do
          End If
          Set brs3 = Nothing
        End If
        Set brs3 = Nothing
     End If
     If is_stno(awno1, brs5!cw_now) And Left(brs5!cw_now, 1) <> "(" And brs5!cw_first <> "Y" Then
        brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_now like '" & Trim(brs5!cw_now) & "%' order by cw_sitm", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs3.EOF Then
          If brs3!cw_palt <> brs5!cw_palt And get_string_value1(awno1, brs3!cw_now, "st_fbst") <> "X" Then
            If brs5!cw_err <> 48 Then rcn.Execute "update cw set cw_err=48 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
            Set brs3 = Nothing: GoTo rtn
          End If
        End If
        Set brs3 = Nothing
     End If
     If Left(project_name, 4) = "SK01" And awno1 = "4" And Left(brs5!cw_to, 4) >= "D007" And Left(brs5!cw_to, 4) <= "D009" Then GoTo rtn67
     If is_stno(awno1, brs5!cw_nxt) And brs5!cw_first <> "Y" Then
        brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_nxt like '" & Trim(brs5!cw_nxt) & "%' order by cw_sitm", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs3.EOF Then
          If brs3!cw_palt <> brs5!cw_palt And get_string_value1(awno1, brs3!cw_nxt, "st_fbst") <> "X" Then
            If brs5!cw_err <> 48 Then rcn.Execute "update cw set cw_err=48 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
            Set brs3 = Nothing: GoTo rtn
          End If
        End If
        Set brs3 = Nothing
     End If
rtn67:
     If Left(project_name, 4) = "SK01" And awno1 = "4" Then
       If Left(brs5!cw_to, 4) >= "D007" And Left(brs5!cw_to, 4) <= "D009" Then
         '強制兩站會輪流
         brs3.Open "select * from cw where substring(cw_to,1,4)='" & Left(brs5!cw_to, 4) & "' and ((substring(cw_nxt,1,4)='D002' and cw_stat='0') or substring(cw_now,1,4)='D002')", rcn1, adOpenKeyset, adLockOptimistic
         'brs3.Open "select * from cw where substring(cw_to,1,4)='" & Left(brs5!cw_to, 4) & "' and (cw_stat='2' or cw_to='" & brs5!cw_nxt & "')", rcn1, adOpenKeyset, adLockOptimistic
         If Not brs3.EOF Then
           If brs5!cw_err <> 28 Then rcn.Execute "update cw set cw_err=28 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           Set brs3 = Nothing: GoTo rtn
         End If
         Set brs3 = Nothing
         '強制3站會輪流
         brs3.Open "select count(*) from cw where substring(cw_now,1,1)='0' and cw_stat='1' and substring(cw_to,1,4)='D007'", rcn1, adOpenKeyset, adLockOptimistic
         d007 = brs3(0): Set brs3 = Nothing
         brs3.Open "select count(*) from cw where substring(cw_now,1,1)='0' and cw_stat='1' and substring(cw_to,1,4)='D008'", rcn1, adOpenKeyset, adLockOptimistic
         d008 = brs3(0): Set brs3 = Nothing
         brs3.Open "select count(*) from cw where substring(cw_now,1,1)='0' and cw_stat='1' and substring(cw_to,1,4)='D009'", rcn1, adOpenKeyset, adLockOptimistic
         d009 = brs3(0): Set brs3 = Nothing
         If d007 > 0 And d008 > 0 And d009 > 0 Then
           If Left(brs5!cw_to, 4) = dbuf(7) Or Left(brs5!cw_to, 4) = dbuf(8) Then
             If brs5!cw_err <> 28 Then rcn.Execute "update cw set cw_err=28 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
             GoTo rtn
           End If
         End If
       End If
       If Left(brs5!cw_now, 4) = "D001" And Left(brs5!cw_nxt, 4) = "D004" Then
         If brs5!cw_ivar < 25 Then GoTo rtn
       End If
       If Left(brs5!cw_now, 4) = "D006" And Left(brs5!cw_nxt, 1) = "0" Then
         If brs5!cw_fvar < 25 Then GoTo rtn
       End If
     ElseIf Left(project_name, 4) = "SK03" And awno1 = "3" Then
       If Left(brs5!cw_now, 4) = "C026" And Left(brs5!cw_ltno, 10) = "change car" Then '棧板換車
         If change_car < 220 Then '220/60分
           If brs5!cw_err <> 100 Then rcn.Execute "update cw set cw_err=100 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           GoTo rtn
         End If
         brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and substring(cw_now,1,4)='" & Left(brs5!cw_now, 4) & "' order by cw_sitm", rcn1, adOpenKeyset, adLockOptimistic
         If Not brs3.EOF Then
           If (Left(brs5!cw_now, 1) <> "(" And brs5!cw_first <> "Y") And brs3!cw_palt <> brs5!cw_palt Then
             If brs5!cw_err <> 48 Then rcn.Execute "update cw set cw_err=48 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
             Set brs3 = Nothing: GoTo rtn
           End If
         End If
         Set brs3 = Nothing
       End If
       If Left(brs5!cw_nxt, 4) = "C026" Then
         brs3.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and substring(cw_nxt,1,4)='" & Left(brs5!cw_nxt, 4) & "' order by cw_sitm", rcn1, adOpenKeyset, adLockOptimistic
         If Not brs3.EOF Then
           If (Left(brs5!cw_now, 1) <> "(" And brs5!cw_first <> "Y") And brs3!cw_palt <> brs5!cw_palt Then
             If brs5!cw_err <> 48 Then rcn.Execute "update cw set cw_err=48 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
             Set brs3 = Nothing: GoTo rtn
           End If
         End If
         Set brs3 = Nothing
       End If
     ElseIf Left(project_name, 2) = "SP" Then
       If Left(brs5!cw_nxt, 4) = "C014" Or Left(brs5!cw_nxt, 4) = "C022" Or Left(brs5!cw_nxt, 4) = "C026" Then
         If get_string_value1("", "", "pa_flag") = "Y" Then
           If brs5!cw_err <> 82 Then rcn.Execute "update cw set cw_err=82 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           GoTo rtn
         End If
       End If
       'If Left(brs5!cw_now, STNO_NUM) = "B007" Or Left(brs5!cw_now, STNO_NUM) = "B009" Or Left(brs5!cw_now, STNO_NUM) = "B011" Or Left(brs5!cw_now, STNO_NUM) = "B013" Or Left(brs5!cw_now, STNO_NUM) = "B015" Or Left(brs5!cw_now, STNO_NUM) = "B017" Or Left(brs5!cw_now, STNO_NUM) = "B019" Or Left(brs5!cw_now, STNO_NUM) = "B021" Or Left(brs5!cw_now, STNO_NUM) = "B023" Or Left(brs5!cw_now, STNO_NUM) = "C008" Or Left(brs5!cw_now, STNO_NUM) = "C010" Or Left(brs5!cw_now, STNO_NUM) = "C012" Or Left(brs5!cw_now, STNO_NUM) = "C016" Or Left(brs5!cw_now, STNO_NUM) = "C018" Or Left(brs5!cw_now, STNO_NUM) = "C020" Or Left(brs5!cw_now, STNO_NUM) = "C024" Or Left(brs5!cw_now, STNO_NUM) = "C028" Or Left(brs5!cw_now, STNO_NUM) = "C030" Or Left(brs5!cw_now, STNO_NUM) = "D007" Or Left(brs5!cw_now, STNO_NUM) = "D015" Or Left(brs5!cw_now, STNO_NUM) = "D017" Then
         brs3.Open "select * from st where st_stno='" & Left(brs5!cw_now, 4) & "'", rcn1, adOpenKeyset, adLockOptimistic
         If Not brs3.EOF Then
           palt1 = brs3!st_palt1: Set brs3 = Nothing
           If brs5!cw_from <> brs5!cw_now And palt1 <> brs5!cw_palt Then '非起站
             If brs5!cw_err <> 77 Then rcn.Execute "update cw set cw_err=77 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
             If Left(brs5!cw_now, STNO_NUM) = "B007" Or Left(brs5!cw_now, STNO_NUM) = "B009" Or Left(brs5!cw_now, STNO_NUM) = "B011" Or Left(brs5!cw_now, STNO_NUM) = "B013" Or Left(brs5!cw_now, STNO_NUM) = "B015" Or Left(brs5!cw_now, STNO_NUM) = "B017" Or Left(brs5!cw_now, STNO_NUM) = "B019" Or Left(brs5!cw_now, STNO_NUM) = "B021" Or Left(brs5!cw_now, STNO_NUM) = "B023" Or Left(brs5!cw_now, STNO_NUM) = "C008" Or Left(brs5!cw_now, STNO_NUM) = "C010" Or Left(brs5!cw_now, STNO_NUM) = "C012" Or Left(brs5!cw_now, STNO_NUM) = "C016" Or Left(brs5!cw_now, STNO_NUM) = "C018" Or Left(brs5!cw_now, STNO_NUM) = "C020" Or Left(brs5!cw_now, STNO_NUM) = "C024" Or Left(brs5!cw_now, STNO_NUM) = "C028" Or Left(brs5!cw_now, STNO_NUM) = "C030" Or Left(brs5!cw_now, STNO_NUM) = "D007" Or Left(brs5!cw_now, STNO_NUM) = "D015" Or Left(brs5!cw_now, STNO_NUM) = "D017" Then
               If palt1 <> 0 Then
                 brs3.Open "select * from cw where cw_awno='" & awno1 & "' and cw_now like '" & Left(brs5!cw_now, 4) & "%' and cw_palt=" & palt1, rcn1, adOpenKeyset, adLockOptimistic
                 If Not brs3.EOF Then
                   rcn.Execute "update cw set cw_stat='0' where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                   Call syslog1("st_comd_cran_err", brs5!cw_palt, palt1, Left(brs5!cw_now, 4) & "-->" & Left(brs5!cw_nxt, 8) & " end")
                   Set brs3 = Nothing: GoTo rtn3
                 Else
                   Set brs3 = Nothing:  GoTo rtn
                 End If
               End If
             End If
             GoTo rtn
           End If
         End If
         Set brs3 = Nothing
       'End If
     ElseIf Left(project_name, 2) = "SK" Then
       If Left(mveq, 2) = "#5" Then
         If Left(project_name, 4) = "SK05" Then
           If Left(brs5!cw_now, 3) = "E02" Or Left(brs5!cw_now, 3) = "E20" Or Left(brs5!cw_now, 3) = "E21" Then 'E02入庫
             If conveyor_full(awno1, Left(brs5!cw_now, 3), 3, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
           End If
           If Left(brs5!cw_nxt, 3) = "E02" Then
             brs3.Open "select * from cw where cw_from='E01     '", rcn, adOpenKeyset, adLockOptimistic
             If Not brs3.EOF Then
               If brs5!cw_err <> 46 Then rcn.Execute "update cw set cw_err=46 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               Set brs3 = Nothing: GoTo rtn
             End If
             Set brs3 = Nothing
           ElseIf Left(brs5!cw_nxt, 3) = "E20" Then
             brs3.Open "select * from cw where cw_from='E25     '", rcn, adOpenKeyset, adLockOptimistic
             If Not brs3.EOF Then
               If brs5!cw_err <> 46 Then rcn.Execute "update cw set cw_err=46 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               Set brs3 = Nothing: GoTo rtn
             End If
             Set brs3 = Nothing
           ElseIf Left(brs5!cw_nxt, 3) = "E21" Then
             brs3.Open "select * from cw where cw_from='E26     '", rcn, adOpenKeyset, adLockOptimistic
             If Not brs3.EOF Then
               If brs5!cw_err <> 46 Then rcn.Execute "update cw set cw_err=46 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               Set brs3 = Nothing: GoTo rtn
             End If
             Set brs3 = Nothing
           End If
         End If
         If Left(brs5!cw_now, 3) = "E12" Then
           If conveyor_full(awno1, Left(brs5!cw_now, 3), 3, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
         ElseIf Left(brs5!cw_now, 3) = "E16" Then 'E16入庫
           If conveyor_full(awno1, Left(brs5!cw_now, 3), 3, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
         ElseIf Left(brs5!cw_nxt, 3) = "E16" Then
           brs3.Open "select * from cw where cw_from='E15     '", rcn, adOpenKeyset, adLockOptimistic
           If Not brs3.EOF Then
             If brs5!cw_err <> 46 Then rcn.Execute "update cw set cw_err=46 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
             Set brs3 = Nothing: GoTo rtn
           End If
           Set brs3 = Nothing
         ElseIf Mid(brs5!cw_nxt, 1, 3) = "F00" Then
           If get_string_value1(awno1, "F01", "st_load") = "1" Then
             If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
             GoTo rtn
           End If
           If conveyor_full(awno1, "F01", 2, brs5!cw_palt, brs5!cw_err, 10) Then GoTo rtn
           If get_string_value1(awno1, "F21", "st_load") <> "1" Then 'F21代表是否允許出庫至F00之命令,當允許時放'1'
             If brs5!cw_err <> 39 Then rcn.Execute "update cw set cw_err=39 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
             GoTo rtn 'Exit Do
           End If
           'If Left(brs5!cw_now, 1) <> "(" Then
           '  brs3.Open "select * from cw where cw_from like '0%' and cw_nxt='F00     ' and cw_stno='F00' order by cw_sitm", rcn1, adOpenKeyset, adLockOptimistic
           '  If Not brs3.EOF Then
           '    If brs5!cw_palt <> brs3!cw_palt Then
           '      If brs5!cw_err <> 48 Then rcn.Execute "update cw set cw_err=48 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           '      Set brs3 = Nothing: GoTo rtn 'Exit Do
           '    End If
           '  End If
           '  Set brs3 = Nothing
           'End If
           ii = 0
           Do
             If ii >= 5 Then Exit Do
             If get_string_value1(awno1, "F00", "st_load") = "1" Then
               If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
             ii = ii + 1: Call Wait(1)
           Loop
         ElseIf Mid(brs5!cw_now, 1, 3) = "F02" Then
           If conveyor_full(awno1, Left(brs5!cw_now, 3), 3, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
           ii = 0
           Do
             If ii >= 5 Then Exit Do
             If get_string_value1(awno1, "F02", "st_load") = "0" Then
               If brs5!cw_err <> 3 Then rcn.Execute "update cw set cw_err=3 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
             ii = ii + 1: Call Wait(1)
           Loop
         End If
       ElseIf Left(mveq, 2) = "#1" Then
         If Left(project_name, 4) = "SK04" Then
           If Mid(brs5!cw_now, 4, 2) = "25" Or Mid(brs5!cw_now, 4, 2) = "26" Or Mid(brs5!cw_now, 1, 3) = "A66" Or Mid(brs5!cw_nxt, 4, 2) = "25" Or Mid(brs5!cw_nxt, 4, 2) = "26" Or Mid(brs5!cw_nxt, 1, 3) = "A69" Then
             ii = 0
             Do
               If ii >= 5 Then Exit Do
               If get_string_value1(awno1, "A90", "st_load") <> NOLD Or get_string_value1(awno1, "A91", "st_load") <> NOLD Or get_string_value1(awno1, "A92", "st_load") <> NOLD Or get_string_value1(awno1, "A93", "st_load") <> NOLD Then
                 If brs5!cw_err <> 15 Then rcn.Execute "update cw set cw_err=15 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                 If Mid(brs5!cw_nxt, 1, 3) = "A69" Then
                   Call syslog1("st_data_cran", brs5!cw_palt, ii, "now=" & brs5!cw_now & ",A90=" & get_string_value1(awno1, "A90", "st_load") & ",A92=" & get_string_value1(awno1, "A92", "st_load"))
                 End If
                 GoTo rtn
               End If
               brs3.Open "select * from cw where cw_mveq='ROTATE    '", rcn1, adOpenKeyset, adLockOptimistic
               If Not brs3.EOF Then
                 If brs5!cw_err <> 23 Then rcn.Execute "update cw set cw_err=23 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                 Set brs3 = Nothing: GoTo rtn 'Exit Do
               End If
               Set brs3 = Nothing
               ii = ii + 1: Call Wait(1)
             Loop
             If Mid(brs5!cw_nxt, 1, 3) = "A69" Then
               'Call syslog1("st_data_cran1", brs5!cw_palt, ii, "now=" & brs5!cw_now & ",A90=" & get_string_value1(awno1, "A90", "st_load") & ",A92=" & get_string_value1(awno1, "A92", "st_load"))
             End If
           End If
           If brs5!cw_now <> "00024001" Then
             brs3.Open "select * from cw where cw_now='00024001' and cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'", rcn, adOpenKeyset, adLockOptimistic
             If Not brs3.EOF Then
               If brs5!cw_err <> 2 Then rcn.Execute "update cw set cw_err=2 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               Set brs3 = Nothing: GoTo rtn 'Exit Do
             End If
             Set brs3 = Nothing
           End If
           If Left(brs5!cw_nxt, 3) = "A69" Then
             If get_string_value1(awno1, "A95", "st_load") <> "1" Then 'A95-A69不允許天車下出庫命令 或 A97-A69不允許STV4下命令
               If brs5!cw_err <> 26 Then rcn.Execute "update cw set cw_err=26 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
           End If
         End If
         If Left(brs5!cw_now, 3) = "A66" Then
           If Left(project_name, 4) = "SK04" Then
             If conveyor_full(awno1, Left(brs5!cw_now, 3), 2, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
             If get_string_value1(awno1, "A96", "st_load") <> "1" Then 'A94-A66不允許STV3下命令 或 A96-A66不允許天車下入庫命令
               If brs5!cw_err <> 25 Then rcn.Execute "update cw set cw_err=25 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
           Else
             If conveyor_full(awno1, Left(brs5!cw_now, 3), 5, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
           End If
         End If
         If Left(brs5!cw_nxt, 3) = "A69" Then
           If Left(project_name, 4) = "SK04" Then
             brs3.Open "select * from cw where cw_now='A68     ' and cw_mveq='MONO      '", rcn1, adOpenKeyset, adLockOptimistic
             If Not brs3.EOF Then
               If brs5!cw_err <> 37 Then rcn.Execute "update cw set cw_err=37 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               Set brs3 = Nothing: GoTo rtn
             End If
             Set brs3 = Nothing
           Else
             If conveyor_full(awno1, "A68", 5, brs5!cw_palt, brs5!cw_err, 10) Then GoTo rtn
           End If
         End If
       End If
     End If
     'wbuf = get_pfile("bypass_st")
     'If Left(wbuf, 1) = "Y" Then GoTo rtn20
     If Left(project_name, 4) >= "SK01" And Left(project_name, 4) <= "SK03" Then
       If Left(project_name, 4) = "SK03" Then
         If awno1 = "1" Then GoTo rtn20
       End If
       'If get_string_value1("", "", "pa_loadcheck") = "N" Then GoTo rtn20
     End If
     If is_stno(awno1, brs5!cw_now) And get_string_value1(awno1, Trim(brs5!cw_now), "st_stat") <> "V" Then
'起站須有載
        If get_string_value1(awno1, Trim(brs5!cw_now), "st_load") = NOLD Then
           If brs5!cw_err <> 3 Then rcn.Execute "update cw set cw_err=3 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           GoTo rtn
        End If
     End If
'迄站須無鎖定
     If is_stno(awno1, brs5!cw_nxt) And check_block(brs5!cw_nxt, " ", ST_WORK) = True Then
       If Not (Left(project_name, 4) >= "SK01" And Left(project_name, 4) <= "SK03" And awno1 = "4") Then
         If brs5!cw_err <> 4 Then rcn.Execute "update cw set cw_err=4 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
         GoTo rtn 'Exit Do
       End If
     End If
     If is_stno(awno1, brs5!cw_nxt) And get_string_value1(awno1, Trim(brs5!cw_nxt), "st_stat") <> "V" Then
'迄站須無載
        If Left(project_name, 2) = "SJ" Then
'A05-->A04-->A03檢查A04無載就可以下命令,注意不必檢查A05迄站須無載狀態
           from1 = Array("A05", "A13", "A21", "A28", "A33", "END")
           to1 = Array("A04", "A12", "A20", "A27", "A32")
           For ii = 0 To 1000
              If Left(brs5!cw_nxt, 3) = from1(ii) Or from1(ii) = "END" Then Exit For
           Next ii
           If from1(ii) <> "END" Then
              lib_buf = to1(ii)
              If get_string_value1(awno1, Left(lib_buf, 3) & " ", "st_load") <> "0" Then
'輸送機組全部為有載狀態,因之命令無法下達
                 If brs5!cw_err <> 10 Then rcn.Execute "update cw set cw_err=10 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                 GoTo rtn
              End If
          End If
        Else
           If get_string_value1(awno1, Trim(brs5!cw_nxt), "st_load") <> "0" Then
rtn10:
              If Left(project_name, 2) = "SK" Then
                 If Left(brs5!cw_nxt, 3) = "D02" And get_string_value1(awno1, Trim(brs5!cw_nxt), "st_load") = "1" Then
                    GoTo rtn20
                 End If
              ElseIf Left(project_name, 2) = "SP" And Left(brs5!cw_nxt, 1) = "D" Then
                nxt1 = Mid(brs5!cw_nxt, 1, 2) & Right("00" & (Val(Mid(brs5!cw_nxt, 3, 2)) + 1), 2)
                If get_string_value1(awno1, nxt1, "st_load") = "0" Then
                  GoTo rtn20
                End If
              End If
              If palt_first <> 0 Then
                 'Err_Translate = "此為優先執行之命令,但迄站為有載狀態,因之出庫命令無法下達,請注意這將會造成所有其它的入出庫命令都無法執行,因之請速排除有載狀態"
                 If brs5!cw_err <> 9 Then rcn.Execute "update cw set cw_err=9 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
'palt_first_flag = True
              Else
                 If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
              End If
              GoTo rtn
           Else
              If Left(project_name, 2) = "SI" Then
                 If Left(brs5!cw_nxt, 3) = "H13" And get_string_value1(awno1, "H12", "st_load") <> "0" Then
                    GoTo rtn10
                 End If
              End If
           End If
        End If
     End If
rtn20:
     lib_buf = ""
     If Left(project_name, 2) = "SG" Then
        array_buf = Array("C05 ", "C01 ", "C25 ", "C21 ", "C41 ", "C91 ", "C15 ", "C11 ", "C31 ", "C81 ", "END")
     ElseIf Left(project_name, 2) = "SI" Then
        If Left(brs5!cw_now, 3) = "G04" Then
           If get_string_value1(awno1, EQ_STV4, "cr_load") <> "1" Then GoTo rtn
        ElseIf Left(brs5!cw_now, 3) = "G26" Then
           If get_string_value1(awno1, EQ_STV5, "cr_load") <> "1" Then GoTo rtn
        End If
        array_buf = Array("H13 ", "H23 ", "G04 ", "G26 ", "H33 ", "H43 ", "H53 ", "H72 ", "H71 ", "END")
     ElseIf Left(project_name, 2) = "SJ" Then
       If Left(project_name, 4) = "SJ02" Then array_buf = Array("A02 ", "A05 ", "A11 ", "A13 ", "A21 ", "A28 ", "A31 ", "A33 ", "END")
       If Left(project_name, 4) = "SJ03" Then array_buf = Array("G07 ", "G12 ", "G13 ", "H02 ", "H03 ", "H08 ", "G09 ", "E01 ", "F01 ", "C03 ", "D15 ", "D08 ", "C04 ", "C08 ", "END")
     ElseIf Left(project_name, 2) = "SK" Then
       If Left(project_name, 4) = "SK01" Then
         If Mid(mveq, 2, 1) = "1" Then array_buf = Array("A001 ", "A002 ", "A003 ", "A004 ", "A005 ", "A006 ", "A007 ", "A008 ", "END")
         If Mid(mveq, 2, 1) = "2" Then array_buf = Array("B001 ", "B002 ", "B003 ", "B004 ", "B005 ", "B006 ", "B007 ", "B008 ", "END")
         If Mid(mveq, 2, 1) = "3" Then array_buf = Array("C001 ", "C002 ", "C003 ", "C004 ", "C005 ", "END")
         If Mid(mveq, 2, 1) = "4" Then array_buf = Array("D001 ", "D002 ", "D003 ", "D004 ", "D005 ", "D006 ", "END")
         If Mid(mveq, 2, 1) = "5" Then array_buf = Array("E001 ", "E002 ", "END")
       ElseIf Left(project_name, 4) = "SK02" Then
         If Mid(mveq, 2, 1) = "1" Then array_buf = Array("A001 ", "A002 ", "A003 ", "A004 ", "A005 ", "A006 ", "A007 ", "A008 ", "A009 ", "END")
         If Mid(mveq, 2, 1) = "2" Then array_buf = Array("B001 ", "B002 ", "B003 ", "B004 ", "B005 ", "B006 ", "B007 ", "END")
         If Mid(mveq, 2, 1) = "3" Then array_buf = Array("C001 ", "C002 ", "C003 ", "C004 ", "C005 ", "C006 ", "END")
         If Mid(mveq, 2, 1) = "4" Then array_buf = Array("D001 ", "D002 ", "D003 ", "D004 ", "D005 ", "END")
         If Mid(mveq, 2, 1) = "5" Then array_buf = Array("E001 ", "E002 ", "E003 ", "E004 ", "E005 ", "E006 ", "E007 ", "E008 ", "E009 ", "END")
       ElseIf Left(project_name, 4) = "SK03" Then
         If Mid(mveq, 2, 1) = "1" Then array_buf = Array("A011 ", "A012 ", "A013 ", "A014 ", "A015 ", "END")
         If Mid(mveq, 2, 1) = "7" Then array_buf = Array("A021 ", "A022 ", "A023 ", "A024 ", "A025 ", "A026 ", "END")
         If Mid(mveq, 2, 1) = "2" Then array_buf = Array("B011 ", "B012 ", "B013 ", "B014 ", "END")
         If Mid(mveq, 2, 1) = "3" Then array_buf = Array("C021 ", "C022 ", "C023 ", "C024 ", "C025 ", "C026 ", "C027 ", "C028 ", "END")
         If Mid(mveq, 2, 1) = "4" Then array_buf = Array("D051 ", "D052 ", "D053 ", "D054 ", "D055 ", "D056 ", "END")
         If Mid(mveq, 2, 1) = "5" Then array_buf = Array("E071 ", "E072 ", "E073 ", "E074 ", "E075 ", "E076 ", "E077 ", "E078 ", "E079 ", "E080 ", "E081 ", "END")
         If Mid(mveq, 2, 1) = "6" Then array_buf = Array("F041 ", "F042 ", "F043 ", "F044 ", "F045 ", "END")
       ElseIf Left(project_name, 4) = "SK04" Then
         If Mid(mveq, 2, 1) = "1" Then
           If Left(brs5!cw_now, 3) = "A66" And is_lono(awno1, brs5!cw_nxt) Then
              '讀取磅秤
              'jj = sio_open(9): If jj < 0 Then GoTo err1_rtn
              'jj = sio_ioctl(9, B9600, BIT_7 Or STOP_1 Or P_EVEN): If jj < 0 Then GoTo err1_rtn
              'jj = sio_flush(9, 2): If jj < 0 Then GoTo err1_rtn
              'Call Wait(1): wbuf = "": jj = sio_read(9, wbuf, 36)
              'sio_close (9)
              If jj = 36 Then
                'jj = InStr(1, wbuf, "+", 0)
                'rcn.Execute "update pa set pa_gros1=" & Val(Mid(wbuf, jj + 1, 7))
                'rcn.Execute "update sg set sg_gros1=" & Val(Mid(wbuf, jj + 1, 7)) & " where sg_to='" & brs5!cw_to & "'"
              End If
            End If
            array_buf = Array("A66", "A69", "A70", "A71", "A72", "A80", "A81", "A82", "END")
         End If
         If Mid(mveq, 2, 1) = "2" Then array_buf = Array("B10", "B11", "B12", "B13", "B14", "B15", "B30", "B31", "B32", "B33", "B34", "END")
         If Mid(mveq, 2, 1) = "3" Then array_buf = Array("C10", "C11", "C12", "C13", "C14", "C30", "C31", "END")
         If Mid(mveq, 2, 1) = "4" Then array_buf = Array("D10", "D12", "D14", "D16", "D00", "D01", "D02", "D20", "END")
         If Mid(mveq, 2, 1) = "5" Then array_buf = Array("F00", "F02", "E12", "E14", "E16", "E20", "E21", "END")
       ElseIf Left(project_name, 4) = "SK05" Then
         If Mid(mveq, 2, 1) = "1" Then array_buf = Array("A72", "A70", "A71", "A80", "A81", "A66", "A69", "END")
         If Mid(mveq, 2, 1) = "2" Then array_buf = Array("B19", "B18", "B17", "B16", "B15", "B14", "B13", "B12", "B11", "B10", "END")
         If Mid(mveq, 2, 1) = "3" Then array_buf = Array("C16", "C15", "C14", "C13", "C12", "C11", "C10", "END")
         If Mid(mveq, 2, 1) = "4" Then array_buf = Array("D02", "D01", "D00", "D14", "D12", "D10", "D20", "END")
         If Mid(mveq, 2, 1) = "5" Then array_buf = Array("E02", "E16", "E14", "E12", "E21", "E20", "F02", "F00", "END")
       End If
     ElseIf Left(project_name, 4) = "SP01" Then
       If Mid(mveq, 2, 1) = 1 Then array_buf = Array("D001", "B007", "C008", "END") ' "00R"表示搬倉作業
       If Mid(mveq, 2, 1) = 2 Then array_buf = Array("D003", "B009", "C010", "END")
       If Mid(mveq, 2, 1) = 3 Then array_buf = Array("D005", "D007", "B011", "C012", "C014", "END")
       If Mid(mveq, 2, 1) = 4 Then array_buf = Array("D009", "B013", "C016", "END")
       If Mid(mveq, 2, 1) = 5 Then array_buf = Array("D011", "B015", "C018", "END")
       If Mid(mveq, 2, 1) = 6 Then array_buf = Array("D013", "D015", "B017", "C020", "C022", "END")
       If Mid(mveq, 2, 1) = 7 Then array_buf = Array("D017", "D019", "B019", "C024", "C026", "END")
       If Mid(mveq, 2, 1) = 8 Then array_buf = Array("D021", "B021", "C028", "END")
       If Mid(mveq, 2, 1) = 9 Then array_buf = Array("D023", "B023", "C030", "END")
     End If
     If Left(brs5!cw_now, 1) = "(" Then
        lib_buf = "00000000"
     'ElseIf Left(mode, 3) = "000" Then
     ElseIf Left(brs5!cw_now, 5) = "99000" Then '天車移位至某站但不搬棧板的情況,99000001-->000000,test用
        Mid(lib_buf, 1, 16) = brs5!cw_now & "00000000": GoTo rtn13
     ElseIf Left(brs5!cw_now, 3) = "000" Then '天車移位至某站並搬棧板至天車上的情況,00000001-->000000,test用
        Mid(lib_buf, 1, 16) = brs5!cw_now & "00000000": GoTo rtn13
     ElseIf Left(brs5!cw_nxt, 2) = "00" Then '天車移位至某bank某level的情況,不搬棧板,00099099-->000000,SK04 case
        Mid(lib_buf, 1, 16) = brs5!cw_nxt & "00000000": GoTo rtn13
     ElseIf is_lono(awno1, brs5!cw_now) Then
        If (Val(Mid(brs5!cw_now, 1, 2)) Mod 2) = 0 Then Mid(lib_buf, 1, 2) = "02" Else Mid(lib_buf, 1, 2) = "01"
        Mid(lib_buf, 3, 6) = Mid(brs5!cw_now, 3, 6)
     Else
        For ii = 0 To 1000
           If Left(array_buf(ii), 3) = "END" Then Exit For
           If Left(array_buf(ii), STNO_NUM) = Mid(brs5!cw_now, 1, STNO_NUM) Then
              lib_buf = Right("00000000" & (ii + 1), 8)
              Exit For
           End If
        Next ii
     End If
     If is_lono(awno1, brs5!cw_nxt) Then
        If (Val(Mid(brs5!cw_nxt, 1, 2)) Mod 2) = 0 Then Mid(lib_buf, 1 + 8, 2) = "02" Else Mid(lib_buf, 1 + 8, 2) = "01"
        Mid(lib_buf, 3 + 8, 6) = Mid(brs5!cw_nxt, 3, 6)
     Else
        For ii = 0 To 1000
           If Left(array_buf(ii), 3) = "END" Then Exit For
           If Left(array_buf(ii), STNO_NUM) = Mid(brs5!cw_nxt, 1, STNO_NUM) Then
              Mid(lib_buf, 1 + 8, 8) = Right("00000000" & (ii + 1), 8)
              Exit For
           End If
        Next ii
     End If
rtn13:
     palt_plc_write_asrs = brs5!cw_palt
     If Left(project_name, 2) = "SP" Then
       If plc_write_asrs_sp(port_id, 4, "04" & Chr(13) & Chr(10), mveq, 4, "", brs5!cw_now, brs5!cw_nxt) = False Then GoTo rtn 'Exit Do
       If plc_write_asrs_sp(port_id, 20, "01" & Left(lib_buf, 16) & Chr(13) & Chr(10), mveq, 4, "", brs5!cw_now, brs5!cw_nxt) = False Then GoTo rtn 'Exit Do
     Else
       If plc_write_asrs(port_id, 4, "04" & Chr(13) & Chr(10), mveq, 4, awno1) = False Then GoTo rtn 'Exit Do
       If plc_write_asrs(port_id, 20, "01" & Left(lib_buf, 16) & Chr(13) & Chr(10), mveq, 4, awno1) = False Then GoTo rtn 'Exit Do
     End If
     rcn.Execute "update cw set cw_stat='2',cw_err=0 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt & " and cw_from='" & brs5!cw_from & "'"
     rcn.Execute "update cr set cr_err=0,cr_palt=" & brs5!cw_palt & ",cr_from='" & brs5!cw_now & "',cr_to='" & brs5!cw_nxt & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
     If Left(project_name, 2) = "SK" Then
       If Left(project_name, 4) = "SK01" And awno1 = "4" Then
         If Left(brs5!cw_to, 4) >= "D007" And Left(brs5!cw_to, 4) <= "D009" Then
           For ii = 1 To 8
             dbuf(ii - 1) = dbuf(ii)
           Next ii
           dbuf(8) = Left(brs5!cw_to, 4)
         End If
       ElseIf Left(project_name, 4) = "SK03" And awno1 = "3" Then
         If Left(brs5!cw_now, 4) = "C026" Then change_car = 0
         If Left(brs5!cw_nxt, 4) = "C026" Then change_car = 1
       ElseIf Left(project_name, 4) = "SK04" Or Left(project_name, 4) = "SK05" Then
         '在rd_data_cran中天車搬運結束時才產生虛擬站D20命令
         If Left(mveq, 2) = "#4" And Left(brs5!cw_nxt, 3) = "D20" Then GoTo rtn3
         brs3.Open "select pa_led from pa", rcn, adOpenKeyset, adLockOptimistic
         Call InitialComPort(brs3(0), 9600, False, 8, 1)
         Call SendMessage("歡迎使用玻纖布廠" & get_awnm(awno1) & Mid(get_now(), 3, 8), brs3(0), 1, 40)
         Call Wait(1) '務必作
         '''frmMain.mscLED(brs3(0) - 1).PortOpen = False
         Set brs3 = Nothing
       End If
     End If
'虛擬站命令分裂
     If Trim(brs5!cw_nxt) <> Trim(brs5!cw_to) And get_string_value1(awno1, brs5!cw_nxt, "st_stat") = "V" Then
        rcn.Execute "update cw set cw_to='" & brs5!cw_nxt & "' where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        ii = put_cw(awno1, brs5!cw_srid, brs5!cw_nxt, brs5!cw_to, brs5!cw_name, "N")
        If Left(project_name, 2) = "SI" Then
           rcn.Execute "update cw set cw_ptno='" & brs5!cw_ptno & "',cw_odno='" & brs5!cw_odno & "',cw_ltno='" & brs5!cw_ltno & "',cw_quty=" & brs5!cw_quty & ",cw_mark='" & brs5!cw_mark & "' where cw_palt=" & ii
        End If
     End If
rtn3:
     Call syslog("st_comd_cran", 0, 0, "st_comd:palt=" & brs5!cw_palt & ",now=" & brs5!cw_now & ",nxt=" & brs5!cw_nxt & ",palt_first=" & palt_first)
     st_comd_cran = False: palt_first_flag = True
     Exit Do
rtn:
     'If Left(project_name, 2) = "SP" Then Exit Do
     If Left(brs5!cw_now, 1) = "(" Then Exit Do
     brs5.MoveNext
  Loop
  Set brs5 = Nothing
'cw_first='Y的工作若無法執行時務必要將cw_first改為'N',否則所有cw_first='N'的工作都不會執行
  If palt_first <> 0 And palt_first_flag = False Then
     palt_first = 0: GoTo rtn2
'rCn.Execute "update cw set cw_first='N' where cw_palt=" & palt_first & " and cw_stat = '1'"
  End If
  Exit Function
err1_rtn:
  Set brs3 = Nothing:  Set brs5 = Nothing
  If Err.Number = -2147467259 Then   'SQL deadlock case,或[DBNETLIB]ConnectionOpen: SQL Server不存在或拒絕存取,或連線失敗(for SO專案Save程式)
    Call syslog1("err2_rtn:st_comd_cran", 999, 0, prog & "-st_comd_cran:系統異常,行號為:" & Erl & ",異常碼:" & Str(Err.Number) & Chr(13) & Err.Description)
    Call Wait(2)
  Else
    Call err2_rtn("st_comd_cran")
  End If
End Function
'response_num可能為4,5,6(for SF project)
Public Function plc_write_asrs_sp(port_id As Integer, leng As Integer, data1 As String, mveq As String, response_num As Integer, Optional awno2 As String, Optional now1 As String, Optional nxt1 As String) As Boolean
  Dim LL As Long, wbuf As String, ii As Integer, jj As Integer, lib_buf As String * 300
  Dim awno1 As String * 1, buf2 As String * 2
  On Error GoTo err1_rtn
  plc_write_asrs_sp = False
  If Left(project_name, 2) = "SP" Then
    awno1 = "1"
  Else
    If awno2 >= "1" And awno2 <= "9" Then awno1 = awno2 Else awno1 = get_awno(port_id)
  End If
  Call sio_flush(port_id, 2)
  wbuf = data1
  Call sio_write(port_id, wbuf, leng)
  Call Wait(0.5)
  lib_buf = "": LL = sio_read(port_id, lib_buf, response_num)
  buf2 = Mid(lib_buf, response_num - 3, 2)
  If LL <> response_num Or buf2 <> "11" Then
    ii = LL
    lib_buf = "send command err: mveq=" & mveq & ",leng=" & leng & ",err=" & buf2 & ",command=" & wbuf: ii = LL: Call syslog("plc_write_asrs", response_num, ii, lib_buf)
    If mveq <> " " Then ii = get_integer_value1(awno1, mveq, "cr_err")
    DoEvents 'Call wait1_100ms(port_id, 1)
    If LL = 0 Then
      If ii <> 67 Then
'          rcn1.Execute "update cr set cr_err=67 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
        rcn1.Execute "update cw set cw_err=67 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
      End If
    ElseIf buf2 = "00" Then
      Call syslog("plc_write_asrs", 53, 53, "天車狀態為非自動或搬運中或異常,因之暫不接受命令")
      If ii <> 53 Then
'          rcn1.Execute "update cr set cr_err=53 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
        rcn1.Execute "update cw set cw_err=53 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
      End If
    ElseIf Left(project_name, 2) = "SP" And (Mid(buf2, 1, 1) = "0" Or Mid(buf2, 2, 1) = "0") Then
      Call syslog("plc_write_asrs", 53, 53, "天車狀態為非自動或搬運中或異常,因之暫不接受命令")
      If ii <> 53 Then
'          rcn1.Execute "update cr set cr_err=53 where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
        rcn1.Execute "update cw set cw_err=53 where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
      End If
      If (Mid(buf2, 1, 2) = "01" Or Mid(buf2, 1, 2) = "10") Then
        If Mid(data1, response_num - 3, 2) <> "04" Then rcn1.Execute "update cr set cr_palt=" & palt_plc_write_asrs & ",cr_from='" & now1 & "',cr_to='" & nxt1 & "' where cr_mveq='" & mveq & "'"
      End If
      lib_buf = "err=" & buf2 & ",mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",command = " & wbuf
      ii = LL: Call syslog1("plc_write_asrs_err0", response_num, ii, Left(Trim(lib_buf), Len(Trim(lib_buf)) - 2))
    Else
      If ii <> Val(buf2) Then
'          rcn1.Execute "update cr set cr_err=" & Val(buf2) & " where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
        rcn1.Execute "update cw set cw_err=" & Val(buf2) & " where cw_awno='" & awno1 & "' and cw_mveq='" & mveq & "'"
      End If
      'If LL = response_num Then
        If Left(project_name, 2) = "SP" And (Mid(buf2, 1, 1) = "1" Or Mid(buf2, 2, 1) = "1") Then
          lib_buf = "err=" & buf2 & ",mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",command = " & wbuf
          ii = LL: Call syslog1("plc_write_asrs_err1", response_num, ii, Left(Trim(lib_buf), Len(Trim(lib_buf)) - 2))
          GoTo rtn13
        End If
      'End If
      'If Left(mveq, 2) = "#1" And Mid(wbuf, 11, 8) = "00000001" Then 'D001
        'lib_buf = "err=" & buf2 & ",mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",command = " & wbuf
        'ii = LL: Call syslog1("plc_write_asrs_err", response_num, ii, Left(Trim(lib_buf), Len(Trim(lib_buf)) - 2))
        If Mid(data1, response_num - 3, 2) <> "04" Then rcn1.Execute "update cr set cr_palt=" & palt_plc_write_asrs & ",cr_from='" & now1 & "',cr_to='" & nxt1 & "' where cr_mveq='" & mveq & "'"
      'End If
      lib_buf = "err=" & buf2 & ",mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",command = " & wbuf
      ii = LL: Call syslog1("plc_write_asrs_err**", response_num, ii, Left(Trim(lib_buf), Len(Trim(lib_buf)) - 2))
    End If
    plc_write_asrs_sp = False: Exit Function
  End If
rtn13:
  palt_plc_write_asrs_old = 0: cmd_plc_write_asrs_old = "00000000"
  If Mid(data1, response_num - 3, 2) <> "04" Then
    lib_buf = "ok,mveq=" & mveq & ",palt=" & palt_plc_write_asrs & ",command = " & wbuf
    ii = LL: Call syslog1("plc_write_asrs", response_num, ii, Left(Trim(lib_buf), Len(Trim(lib_buf)) - 2))
  End If
  plc_write_asrs_sp = True
  Exit Function
err1_rtn:
  Call err2_rtn("plc_write_asrs_sp")
End Function

'葉樹堅吊車
Public Function st_comd_cran_new(port_id As Integer, mveq As String, mode As String) As Boolean
Dim err1 As Integer, ii As Integer, jj As Integer, lib_buf As String * 300, palt_first As Integer, palt_first_flag As Boolean
Dim brs5 As New ADODB.Recordset, brs3 As New ADODB.Recordset, from1, to1
Dim LL As Long, wbuf As String
Dim awno1 As String * 1, array_buf
  On Error GoTo err1_rtn
  palt_first_flag = False: palt_first = 0: st_comd_cran_new = True
  awno1 = "1"
  brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='2' order by cw_mveq,cw_stat", rcn, adOpenKeyset, adLockOptimistic
  If Not brs5.EOF Then
     Set brs5 = Nothing: Exit Function
  End If
  Set brs5 = Nothing
  brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_now='(車上)' and cw_stat='1' order by cw_now", rcn, adOpenKeyset, adLockOptimistic
  If Not brs5.EOF Then GoTo rtn0
  Set brs5 = Nothing
'在下一行CRAN1及CRAN3同時發生
'異常碼: -2147467259,Your transaction (ProcessID#30及ProcessID#39) was deadlocked with another process and has been chosen as the deadlock victim.Rerun your transaction.
  brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_first='Y' and cw_stat='1' order by cw_sitm", rcn, adOpenKeyset, adLockOptimistic
  'brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_first='Y' and cw_stat='1' order by cw_palt", rcn, adOpenKeyset, adLockOptimistic
  If Not brs5.EOF Then
    palt_first = brs5!cw_palt: GoTo rtn0
  End If
rtn1:
  Set brs5 = Nothing
  If Left(mode, 3) = "B01" Then
     ii = ii
  End If
rtn2:
  If Left(mode, 3) = "000" Then '搬棧板至天車上並停住case
     brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_stno='0000' order by cw_mveq,cw_stat,cw_stno,cw_sitm", rcn, adOpenKeyset, adLockOptimistic
  Else
     brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_stno='" & Left(mode, STNO_NUM) & "' order by cw_sitm", rcn, adOpenKeyset, adLockOptimistic
     'brs5.Open "select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_stno='" & Left(mode, STNO_NUM) & "' order by cw_mveq,cw_stat,cw_stno,cw_sitm", rcn, adOpenKeyset, adLockOptimistic
  End If
  Do
     If brs5.EOF = True Then 'cw_first='Y的工作若無法執行時務必要將cw_first改為'N',否則所有cw_first='N'的工作都不會執行
'      If palt_first <> 0 Then rCn.Execute "update cw set cw_first='N' where cw_palt=" & palt_first
        Set brs5 = Nothing: st_comd_cran_new = False: Exit Do
     End If
rtn0:
     If brs5!cw_opno = TR_TEST Then GoTo rtn
     brs3.Open "select * from cw where cw_mveq = '" & brs5!cw_mveq & "' and cw_stat='2'", rcn1, adOpenKeyset, adLockOptimistic
     If Not brs3.EOF Then
       If brs5!cw_err <> 49 Then rcn.Execute "update cw set cw_err=49 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
       Set brs3 = Nothing: GoTo rtn
     End If
     Set brs3 = Nothing
     If (brs5!cw_err >= 100 And brs5!cw_err <= 499) Or (brs5!cw_err >= 600 And brs5!cw_err <= 699) Or (brs5!cw_err >= 3000 And brs5!cw_err <= 3999) Then GoTo rtn
     err1 = check_status(awno1, brs5!cw_mveq, brs5!cw_now, brs5!cw_nxt)
     If err1 > 0 Then
        If Not (err1 = 69 And Left(brs5!cw_now, 1) = "(") Then
           If brs5!cw_err <> err1 Then rcn.Execute "update cw set cw_err=" & err1 & " where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           GoTo rtn 'Exit Do
        End If
     End If
     'check是否有其它作業cw_now="(車上)"(表示東西在車上情況)
     If Left(brs5!cw_now, 1) <> "(" Then
        brs3.Open "select * from cw where cw_now='(車上)' and cw_awno='" & awno1 & "' and cw_palt <> " & brs5!cw_palt & " and cw_mveq='" & mveq & "'", rcn1, adOpenKeyset, adLockOptimistic
        If Not brs3.EOF Then
           If brs5!cw_err <> 2 Then rcn.Execute "update cw set cw_err=2 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           Set brs3 = Nothing: GoTo rtn 'Exit Do
        End If
        Set brs3 = Nothing
     End If
     If is_stno(awno1, brs5!cw_now) Or is_stno(awno1, brs5!cw_nxt) Then
       If Left(brs5!cw_now, 1) <> "(" Then
         If is_stno(awno1, brs5!cw_now) Then
           brs3.Open "select * from cw where cw_now like '" & Trim(Left(brs5!cw_now, STNO_NUM)) & "%' order by cw_sitm", rcn1, adOpenKeyset, adLockOptimistic
         Else
           brs3.Open "select * from cw where cw_nxt like '" & Trim(Left(brs5!cw_nxt, STNO_NUM)) & "%' order by cw_sitm", rcn1, adOpenKeyset, adLockOptimistic
         End If
         If Not brs3.EOF Then
           If brs5!cw_palt <> brs3!cw_palt Then
             If brs5!cw_err <> 48 Then rcn.Execute "update cw set cw_err=48 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
             Set brs3 = Nothing: GoTo rtn 'Exit Do
           End If
         End If
         Set brs3 = Nothing
       End If
     End If
     'wbuf = get_pfile("bypass_st")
     'If Left(wbuf, 1) = "Y" Then GoTo rtn20
     If is_stno(awno1, brs5!cw_now) And get_string_value1(awno1, Trim(brs5!cw_now), "st_stat") <> "V" Then
'起站須有載
        If get_string_value1(awno1, Trim(brs5!cw_now), "st_load") = NOLD Then
           If brs5!cw_err <> 3 Then rcn.Execute "update cw set cw_err=3 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           GoTo rtn
        End If
     End If
'迄站須無鎖定
     If is_stno(awno1, brs5!cw_nxt) And check_block(brs5!cw_nxt, " ", ST_WORK) = True Then
        If brs5!cw_err <> 4 Then rcn.Execute "update cw set cw_err=4 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        GoTo rtn 'Exit Do
     End If
     If is_stno(awno1, brs5!cw_nxt) And get_string_value1(awno1, Trim(brs5!cw_nxt), "st_stat") <> "V" Then
'迄站須無載
       If get_string_value1(awno1, Trim(brs5!cw_nxt), "st_load") <> "0" Then
         If palt_first <> 0 Then
            'Err_Translate = "此為優先執行之命令,但迄站為有載狀態,因之出庫命令無法下達,請注意這將會造成所有其它的入出庫命令都無法執行,因之請速排除有載狀態"
            If brs5!cw_err <> 9 Then rcn.Execute "update cw set cw_err=9 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
'palt_first_flag = True
         Else
            If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
         End If
         GoTo rtn
       End If
     End If
     lib_buf = ""
     If Left(project_name, 4) = "SP01" Then
        array_buf = Array("C05 ", "C01 ", "C25 ", "C21 ", "C41 ", "C91 ", "C15 ", "C11 ", "C31 ", "C81 ", "END")
     End If
     If Left(brs5!cw_now, 1) = "(" Then
        lib_buf = String(12, "0")
     ElseIf is_lono(awno1, brs5!cw_now) Then
        If (Val(Mid(brs5!cw_now, 1, 2)) Mod 2) = 0 Then Mid(lib_buf, 1, 4) = "0002" Else Mid(lib_buf, 1, 4) = "0001"
        Mid(lib_buf, 5, 8) = Format(Hex(Val(Mid(brs5!cw_now, 3, 3))), "0000") & Format(Hex(Val(Mid(brs5!cw_now, 6, 3))), "0000")
     Else
        For ii = 0 To 1000
           If Left(array_buf(ii), 3) = "END" Then Exit For
           If Left(array_buf(ii), STNO_NUM) = Mid(brs5!cw_now, 1, STNO_NUM) Then
              lib_buf = Right(String(12, "0") & (ii + 1), 12)
              Exit For
           End If
        Next ii
     End If
     If is_lono(awno1, brs5!cw_nxt) Then
        If (Val(Mid(brs5!cw_nxt, 1, 2)) Mod 2) = 0 Then Mid(lib_buf, 1 + 12, 4) = "0002" Else Mid(lib_buf, 1 + 12, 4) = "0001"
        Mid(lib_buf, 5 + 12, 6) = Format(Hex(Val(Mid(brs5!cw_nxt, 3, 3))), "0000") & Format(Hex(Val(Mid(brs5!cw_nxt, 6, 3))), "0000")
     ElseIf Left(mode, 3) = "000" Then '搬棧板至天車上並停住case
        Mid(lib_buf, 1 + 12, 8) = "00000000"
     ElseIf Left(brs5!cw_nxt, 2) = "00" Then '天車只有移位的情況
        Mid(lib_buf, 1, 16) = "00" & Mid(brs5!cw_nxt, 3, 6) & "00000000"
     Else
        For ii = 0 To 1000
           If Left(array_buf(ii), 3) = "END" Then Exit For
           If Left(array_buf(ii), STNO_NUM) = Mid(brs5!cw_nxt, 1, STNO_NUM) Then
              Mid(lib_buf, 1 + 12, 12) = Right(String(12, "0") & (ii + 1), 12)
              Exit For
           End If
        Next ii
     End If
     If plc_write(port_id, "D0160", "000A") = False Then GoTo rtn
     If plc_write(port_id, "D0150", Left(lib_buf, 24)) = False Then GoTo rtn
     If plc_write(port_id, "D0160", "000B") = False Then GoTo rtn
     rcn.Execute "update cw set cw_stat='2',cw_err=0 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
     rcn.Execute "update cr set cr_err=0,cr_palt=" & brs5!cw_palt & ",cr_from='" & brs5!cw_now & "',cr_to='" & brs5!cw_nxt & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
'虛擬站命令分裂
     If Trim(brs5!cw_nxt) <> Trim(brs5!cw_to) And get_string_value1(awno1, brs5!cw_nxt, "st_stat") = "V" Then
        rcn.Execute "update cw set cw_to='" & brs5!cw_nxt & "' where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        ii = put_cw(awno1, brs5!cw_srid, brs5!cw_nxt, brs5!cw_to, brs5!cw_name, "N")
     End If
rtn3:
     Call syslog("st_comd_cran_new", 0, 0, "st_comd:palt=" & brs5!cw_palt & ",now=" & brs5!cw_now & ",nxt=" & brs5!cw_nxt & ",palt_first=" & palt_first)
     st_comd_cran_new = False: palt_first_flag = True
     Exit Do
rtn:
     If Left(brs5!cw_now, 1) = "(" Then Exit Do
     brs5.MoveNext
  Loop
  Set brs5 = Nothing
'cw_first='Y的工作若無法執行時務必要將cw_first改為'N',否則所有cw_first='N'的工作都不會執行
  If palt_first <> 0 And palt_first_flag = False Then
     palt_first = 0: GoTo rtn2
'rCn.Execute "update cw set cw_first='N' where cw_palt=" & palt_first & " and cw_stat = '1'"
  End If
  Exit Function
err1_rtn:
  Set brs3 = Nothing: Set brs5 = Nothing
  If Err.Number = -2147467259 Then   'SQL deadlock case,或[DBNETLIB]ConnectionOpen: SQL Server不存在或拒絕存取,或連線失敗(for SO專案Save程式)
    Call syslog1("st_comd_cran_new", 999, 0, prog & "-st_comd_cran:系統異常,行號為:" & Erl & ",異常碼:" & Str(Err.Number) & Chr(13) & Err.Description)
    Call Wait(2)
  Else
    Call err2_rtn("st_comd_cran_new")
  End If
End Function

Public Function conveyor_rtn(model As String, port_id As Integer, stno As String, jj As Integer, mveq As String) As Boolean
  Dim awno1 As String * 1, cnt As Integer, clss As Integer, lib_buf As String * 300, door As String * 4, fork_out As Integer
  On Error GoTo err1_rtn
  awno1 = get_awno(port_id)
  conveyor_rtn = False
  If Left(stno, STNO_NUM) = "B09" Or Left(stno, STNO_NUM) = "B59" Then
    '輸送機B09,輸送機B59入出庫作業,控制自動門B13(dt_clss = 2),自動門B63打開(dt_clss = 4),然後控制自動門B13(dt_clss = 3),自動門B63關閉(dt_clss = 5)
    If Left(stno, STNO_NUM) = "B09" Then
      clss = 2: door = "B13 "
    Else
      clss = 4: door = "B63 "
    End If
    '控制自動門開
    rcn1.Execute "update dt set dt_valu=1 where dt_awno='" & awno1 & "' and dt_clss=" & clss
    rcn1.Execute "update dt set dt_valu=0 where dt_awno='" & awno1 & "' and dt_clss=" & (clss + 1)
    cnt = 0
    Do
      If get_string_value1(awno1, door, "st_load") = NOLD Then
        Call syslog("", 0, 0, "stno=" & stno & ",door=" & door & "open")
        Exit Do '確認自動門開
      End If
      If cnt = 40 Then
        Call syslog("", 0, 0, "stno=" & stno & ",door=" & door & "open timeout")
        If get_integer_value1(awno1, stno, "st_err") <> 51 Then rcn1.Execute "update st set st_err=51 where st_awno='" & awno1 & "' and st_stno='" & Left(stno, STNO_NUM) & "'"
        Exit Function
      End If
      cnt = cnt + 1: Call Wait(0.5)
    Loop
  End If
  If model = "1" Then
    '入庫作業,控制天車取載確認
    If cran_write_fx(port_id, 751 + jj - 17, 1, "0100", mveq) = False Then Exit Function
  Else
    '出庫作業,控制天車卸載確認
    If cran_write_fx(port_id, 751 + jj - 17, 1, "0200", mveq) = False Then Exit Function
  End If
  DoEvents 'Call wait1_100ms(port_id,1)
  If cran_write_fx(port_id, 751 + jj - 17, 1, "0000", mveq) = False Then Exit Function
  If Left(stno, STNO_NUM) = "B09" Or Left(stno, STNO_NUM) = "B59" Then
    cnt = 0: fork_out = 0
    Do
      '暫存器D606代表叉牙狀態
      If cran_read_fx(port_id, 606, 1, lib_buf, mveq) = False Then Exit Function
      'Call syslog("", 0, 0, "stno=" & stno & ",buf=" & Mid(lib_buf, 2, 4) & ",fork_out=" & fork_out)
      '確認叉牙伸出去
      If Mid(lib_buf, 3, 1) <> "0" Then fork_out = 1
      '確認叉牙收回中心
      If Mid(lib_buf, 3, 1) = "0" And fork_out = 1 Then
        Call syslog("", 0, 0, "stno=" & stno & ",fork draw in ok")
        Exit Do
      End If
      If cnt = 40 Then
        Call syslog("", 0, 0, "stno=" & stno & ",fork draw in timeout")
        If get_integer_value1(awno1, stno, "st_err") <> 52 Then rcn1.Execute "update st set st_err=52 where st_awno='" & awno1 & "' and st_stno='" & Left(stno, STNO_NUM) & "'"
        Exit Function
      End If
      cnt = cnt + 1: Call Wait(0.5)
    Loop
    rcn1.Execute "update dt set dt_valu=0 where dt_awno='" & awno1 & "' and dt_clss=" & clss
    '控制自動門關閉
    rcn1.Execute "update dt set dt_valu=1 where dt_awno='" & awno1 & "' and dt_clss=" & (clss + 1)
  End If
  conveyor_rtn = True
  Exit Function
err1_rtn:
  Call err2_rtn("conveyor_rtn")
End Function

Public Function st_comd_cran_fx(port_id As Integer, mveq As String, mode As String) As Boolean
  Dim err1 As Integer, ii As Integer, lib_buf As String * 300, palt_first As Integer, palt_first_flag As Boolean
  Dim brs5 As New ADODB.Recordset, brs3 As New ADODB.Recordset
  Dim LL As Long, wbuf As String
  Dim awno1 As String * 1, array_buf, load1 As String * 1
  On Error GoTo err1_rtn
  palt_first_flag = False: palt_first = 0: st_comd_cran_fx = True
  awno1 = get_awno(port_id)
  Set brs5 = rcn.OpenResultset("select * from cw where cw_mveq='" & mveq & "' and cw_stat='2' order by cw_mveq,cw_stat", rdOpenStatic)
  If Not brs5.EOF Then
    Set brs5 = Nothing: Exit Function
  End If
  Set brs5 = Nothing
  Set brs5 = rcn.OpenResultset("select * from cw where cw_mveq='" & mveq & "' and cw_now='(車上)' and cw_stat='1' order by cw_now", rdOpenStatic)
  If Not brs5.EOF Then GoTo rtn0
  Set brs5 = Nothing
  Set brs5 = rcn.OpenResultset("select * from cw where cw_mveq='" & mveq & "' and cw_first='Y' and cw_stat='1' order by cw_palt", rdOpenStatic)
  If Not brs5.EOF Then
    palt_first = brs5!cw_palt: GoTo rtn0
  End If
  Set brs5 = Nothing
  If Left(mode, 3) = "B01" Then
    ii = ii
  End If
rtn2:
  Set brs5 = rcn.OpenResultset("select * from cw where cw_mveq='" & mveq & "' and cw_stat='1' and cw_stno='" & Left(mode, STNO_NUM) & "' order by cw_mveq,cw_stat,cw_stno,cw_sitm", rdOpenStatic)
  Do
    If brs5.EOF = True Then
'      If palt_first <> 0 Then rCn.Execute "update cw set cw_first='N' where cw_palt=" & palt_first
      st_comd_cran_fx = False: Exit Do
    End If
rtn0:
    If brs5!cw_opno = TR_TEST Then GoTo rtn
    If (brs5!cw_err >= 100 And brs5!cw_err <= 499) Or (brs5!cw_err >= 600 And brs5!cw_err <= 699) Or (brs5!cw_err >= 3000 And brs5!cw_err <= 3999) Then GoTo rtn
    err1 = check_status(awno1, brs5!cw_mveq, brs5!cw_now, brs5!cw_nxt)
    If err1 > 0 Then
      If err1 = 69 Then '有載
        If Left(brs5!cw_now, 1) <> "(" Then
          If brs5!cw_err <> 69 Then rcn.Execute "update cw set cw_err=69 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
          GoTo rtn 'Exit Do
        End If
      Else
        If brs5!cw_err <> err1 Then rcn.Execute "update cw set cw_err=" & err1 & " where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        GoTo rtn 'Exit Do
      End If
    End If
    brs3.Open "select * from cr where cr_awno='" & awno1 & "' and cr_mveq='" & brs5!cw_mveq & "'", rcn1, adOpenKeyset, adLockOptimistic
    If brs3!cr_moti = "3" Then '異常
      If brs5!cw_err <> 680 Then rcn.Execute "update cw set cw_err=680 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
      Set brs3 = Nothing: GoTo rtn 'Exit Do
    End If
    Set brs3 = Nothing
    'check是否有其它作業cw_now="(車上)"(表示東西在車上情況)
    If Left(brs5!cw_now, 1) <> "(" Then
      Set brs3 = rcn.OpenResultset("select * from cw where cw_now='(車上)' and cw_awno='" & awno1 & "' and cw_palt <> " & brs5!cw_palt & " and cw_mveq='" & mveq & "'", rdOpenStatic)
      If Not brs3.EOF Then
        If brs5!cw_err <> 2 Then rcn.Execute "update cw set cw_err=2 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        Set brs3 = Nothing: GoTo rtn 'Exit Do
      End If
      Set brs3 = Nothing
    End If
    If is_stno(awno1, brs5!cw_now) And get_string_value1(awno1, Trim(brs5!cw_now), "st_stat") <> "V" Then
      '起站須有載
      If get_string_value1(awno1, Trim(brs5!cw_now), "st_load") = NOLD Then
        If brs5!cw_err <> 3 Then rcn.Execute "update cw set cw_err=3 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        GoTo rtn
      End If
    End If
    '迄站須無鎖定
    If is_stno(awno1, brs5!cw_nxt) And check_block(brs5!cw_nxt, " ", ST_WORK) = True Then
        If brs5!cw_err <> 4 Then rcn.Execute "update cw set cw_err=4 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        GoTo rtn 'Exit Do
    End If
    If is_stno(awno1, brs5!cw_nxt) And get_string_value1(awno1, Trim(brs5!cw_nxt), "st_stat") <> "V" Then
      '迄站須無載
      If get_string_value1(awno1, Trim(brs5!cw_nxt), "st_load") <> "0" Then
        If palt_first <> 0 Then
          'Err_Translate = "此為優先執行之命令,但迄站為有載狀態,因之出庫命令無法下達,請注意這將會造成所有其它的入出庫命令都無法執行,因之請速排除有載狀態"
          If brs5!cw_err <> 9 Then rcn.Execute "update cw set cw_err=9 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
          'palt_first_flag = True
        Else
          If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        End If
        GoTo rtn
      End If
    End If
    lib_buf = ""
    If Left(project_name, 2) = "SF" Then
      If awno1 = "2" And Left(brs5!cw_nxt, 3) >= "B11" And Left(brs5!cw_nxt, 3) <= "B61" Then
        If get_string_value1(awno1, Left("STV" & (Val(Mid(brs5!cw_nxt, 2, 1)) + 4) & "       ", 10), "cr_load") = "1" Then
          If brs5!cw_err <> 13 Then rcn.Execute "update cw set cw_err=13 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
          GoTo rtn
        End If
      End If
      If Left(brs5!cw_now, 3) = "B01" Or Left(brs5!cw_nxt, 3) = "B01" Then
        load1 = get_string_value1(awno1, "B01   ", "st_load")
        If load1 = "2" Or (Left(brs5!cw_now, 3) = "B01" And load1 = "0") Or (Left(brs5!cw_nxt, 3) = "B01" And load1 = "1") Then
          If brs5!cw_err <> 14 Then rcn.Execute "update cw set cw_err=14 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
          GoTo rtn
        End If
      End If
      If mveq = EQ_CRAN1 Then
        array_buf = Array("A31 ", "A21 ", "A11 ", "A41 ", "END")
      ElseIf mveq = EQ_CRAN2 Then
        array_buf = Array("B51 ", "B61 ", "B01 ", "B11 ", "B21 ", "B31 ", "B41 ", "END")
      End If
    ElseIf Left(project_name, 2) = "SG" Then
      If mveq = EQ_CRAN1 Then
        array_buf = Array("A12 ", "A11 ", "A62 ", "END")
      ElseIf mveq = EQ_CRAN2 Then
        array_buf = Array("B01 ", "B02 ", "B03 ", "B04 ", "B05 ", "B06 ", "B07 ", "B08 ", "B09 ", "B10 ", "B51 ", "B52 ", "B53 ", "B54 ", "B55 ", "B56 ", "B57 ", "B58 ", "B59 ", "B60 ", "END")
      End If
    End If
    If Left(brs5!cw_now, 1) = "(" Then
      lib_buf = "000000000000"
    ElseIf is_lono(awno1, brs5!cw_now) Then
      If (Val(Mid(brs5!cw_now, 1, 2)) Mod 2) = 0 Then Mid(lib_buf, 1, 4) = "0200" Else Mid(lib_buf, 1, 4) = "0100"
      '"13"必須改為"0D"
      Mid(lib_buf, 5, 4) = Right("00" & Hex(Mid(brs5!cw_now, 3, 2)), 2) & "00"
      Mid(lib_buf, 9, 4) = Right("00" & Hex(Mid(brs5!cw_now, 5, 2)), 2) & "00"
    Else
      For ii = 0 To 1000
        If Left(array_buf(ii), 3) = "END" Then Exit For
        If Left(array_buf(ii), STNO_NUM) = Mid(brs5!cw_now, 1, STNO_NUM) Then
          lib_buf = "00000000" & Right("00" & Hex(ii + 1), 2) & "00"
          Exit For
        End If
      Next ii
    End If
    If is_lono(awno1, brs5!cw_nxt) Then
      If (Val(Mid(brs5!cw_nxt, 1, 2)) Mod 2) = 0 Then Mid(lib_buf, 1 + 12, 4) = "0200" Else Mid(lib_buf, 1 + 12, 4) = "0100"
      Mid(lib_buf, 5 + 12, 4) = Right("00" & Hex(Mid(brs5!cw_nxt, 3, 2)), 2) & "00"
      Mid(lib_buf, 9 + 12, 4) = Right("00" & Hex(Mid(brs5!cw_nxt, 5, 2)), 2) & "00"
    Else
      For ii = 0 To 1000
        If Left(array_buf(ii), 3) = "END" Then Exit For
        If Left(array_buf(ii), STNO_NUM) = Mid(brs5!cw_nxt, 1, STNO_NUM) Then
          Mid(lib_buf, 1 + 12, 12) = "00000000" & Right("00" & Hex(ii + 1), 2) & "00"
          Exit For
        End If
      Next ii
    End If
      '一開始cran狀態moti="0"(無作業),pc丟命令到D650,cran開始搬運,最後cran搬運完成時狀態moti="9"(作業完成),
      '注意pc必須清D650後,cran狀態moti才會變為"0"(無作業),然後cran才會再去D650讀命令,
      '倘若cran狀態moti="9"(作業完成),僅管pc丟命令到D650,但cran是不會去D650讀命令
'      If cran_write_fx(port_id, 650, 6, "000000000000000000000000", mveq) = False Then Exit Do
    If cran_write_fx(port_id, 650, 6, lib_buf, mveq) = False Then GoTo rtn 'Exit Do
    rcn.Execute "update cw set cw_stat='2',cw_err=0 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
    rcn.Execute "update cr set cr_err=0,cr_palt=" & brs5!cw_palt & ",cr_from='" & brs5!cw_now & "',cr_to='" & brs5!cw_nxt & "' where cr_awno='" & awno1 & "' and cr_mveq='" & mveq & "'"
    If Left(project_name, 2) = "SF" And awno1 = "2" Then
      rcn.Execute "update dt set dt_valu=0 where dt_awno='" & awno1 & "' and dt_clss=2"
    End If
    '虛擬站命令分裂
    If Trim(brs5!cw_nxt) <> Trim(brs5!cw_to) And get_string_value1(awno1, brs5!cw_nxt, "st_stat") = "V" Then
      rcn.Execute "update cw set cw_to='" & brs5!cw_nxt & "' where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
      Call put_cw(awno1, brs5!cw_srid, brs5!cw_nxt, brs5!cw_to, brs5!cw_name, "N")
    End If
    Call syslog("st_comd_cran", 0, 0, "st_comd:palt=" & brs5!cw_palt & ",now=" & brs5!cw_now & ",nxt=" & brs5!cw_nxt & ",palt_first=" & palt_first)
    st_comd_cran_fx = False: palt_first_flag = True
    Exit Do
rtn:
    If Left(brs5!cw_now, 1) = "(" Then Exit Do
    brs5.MoveNext
  Loop
  Set brs5 = Nothing
'cw_first='Y'的工作若無法執行時務必要將cw_first改為'N',否則所有cw_first='N'的工作都不會執行
  If palt_first <> 0 And palt_first_flag = False Then
    palt_first = 0: GoTo rtn2
    'rCn.Execute "update cw set cw_first='N' where cw_palt=" & palt_first & " and cw_stat = '1'"
  End If
  Exit Function
err1_rtn:
  Call err2_rtn("st_comd_cran_fx")
End Function
Public Function conveyor_full(awno1 As String, stno As String, cnt As Integer, palt As Integer, err1 As Integer, err2 As Integer, Optional mode As String) As Boolean
  Dim brs4 As New ADODB.Recordset, cnt1 As Integer, cnt2 As Integer, i As Integer, stno1 As String
  conveyor_full = False
  brs4.Open "select count(*) from cw where cw_now ='" & Left(stno, STNO_NUM) & "     '", rcn1, adOpenKeyset, adLockOptimistic
  cnt1 = brs4(0): Set brs4 = Nothing
  If cnt1 >= cnt Then
    If err1 <> err2 Then rcn.Execute "update cw set cw_err=" & err2 & " where cw_awno='" & awno1 & "' and cw_palt=" & palt
    conveyor_full = True: Exit Function
  End If
  If err2 = 47 And (mode = "+" Or mode = "-") Then
    cnt2 = 0
    For i = 0 To (cnt - 2)
      If mode = "+" Then stno1 = Left(stno, 1) & CStr(Val(Mid(stno, 2, STNO_NUM - 1)) + i)
      If mode = "-" Then stno1 = Left(stno, 1) & CStr(Val(Mid(stno, 2, STNO_NUM - 1)) - i)
      If get_string_value1(awno1, stno1, "st_load") = "1" Then cnt2 = cnt2 + 1
    Next i
    If cnt1 <> cnt2 Then
      If err1 <> err2 Then rcn.Execute "update cw set cw_err=" & err2 & " where cw_awno='" & awno1 & "' and cw_palt=" & palt
      conveyor_full = True: Exit Function
    End If
  End If
  Exit Function
err1_rtn:
  Call err2_rtn("conveyor_full")
End Function

Public Function st_comd_plc(port_id As Integer, Optional awno2 As String) As Boolean
Dim brs3 As New ADODB.Recordset, brs5 As New ADODB.Recordset, from1, to1, conveyor_flag As Boolean
Dim err1 As Integer, ii As Integer, jj As Integer, lib_buf As String * 100, buf2 As String * 2
Dim LL As Long, wbuf As String, awno1 As String * 1, cnt As Integer, fork_out As Integer, load1 As String * 1
  On Error GoTo err1_rtn
  If awno2 >= "1" And awno2 <= CStr(AW_NUM) Then
    awno1 = awno2
  Else
    awno1 = get_awno(port_id)
  End If
  brs5.Open "select * from cw where cw_awno='" & awno1 & "' and cw_stat='1' and not(cw_mveq like '#%') order by cw_sitm", rcn, adOpenKeyset, adLockOptimistic
  Do
     If brs5.EOF = True Then Exit Do
     If brs5!cw_opno = TR_TEST Then GoTo rtn
     '一套倉庫中有兩個PLC或STV程式會呼叫到此副程式時就要作判斷
     If Left(project_name, 2) = "SI" Then
        If Left(prog, 3) = "STV" And Left(prog, 4) <> Left(brs5!cw_mveq, 4) Then GoTo rtn
        If Left(prog, 4) = "PLC4" And Not (Left(brs5!cw_mveq, 1) = "E" Or Left(brs5!cw_mveq, 1) = "H") Then GoTo rtn
        If Left(prog, 4) = "PLC5" And Not (Left(brs5!cw_mveq, 4) = "STV4" Or Left(brs5!cw_mveq, 4) = "STV5" Or Left(brs5!cw_mveq, 4) = "STV6" Or Left(brs5!cw_mveq, 4) = "STV7" Or Left(brs5!cw_mveq, 4) = "STV8" Or Left(brs5!cw_mveq, 1) = "G") Then GoTo rtn
     ElseIf Left(project_name, 4) = "SU01" Then
        If port_id = 3 And brs5!cw_mveq <> EQ_STV16 Then GoTo rtn
        If port_id = 4 And brs5!cw_mveq <> EQ_STV17 Then GoTo rtn
        If port_id = 5 And brs5!cw_mveq <> EQ_STV18 Then GoTo rtn
        If port_id = 6 And brs5!cw_mveq <> EQ_STV19 Then GoTo rtn
        If brs5!cw_mveq = EQ_STV16 Or brs5!cw_mveq = EQ_STV17 Then
          If Left(brs5!cw_now, 2) = "GL" Or Left(brs5!cw_now, 2) = "Gl" Then
            If brs5!cw_mveq = EQ_STV16 Then wbuf = EQ_STV17
            If brs5!cw_mveq = EQ_STV17 Then wbuf = EQ_STV16
            brs3.Open "select * from cr where cr_awno='" & awno1 & "' and cr_mveq='" & wbuf & "'", rcn1, adOpenKeyset, adLockOptimistic
            If (brs5!cw_mveq = EQ_STV16 And brs3!cr_loct = "000009") Or (brs5!cw_mveq = EQ_STV17 And brs3!cr_loct = "00000A") Then
              If brs5!cw_err <> 74 Then rcn.Execute "update cw set cw_err=74 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
              Set brs3 = Nothing: GoTo rtn
            End If
            Set brs3 = Nothing
          End If
          If Mid(brs5!cw_now, 2, 1) <> "M" And Mid(brs5!cw_now, 2, 1) <> "m" Then
            '棧板準備要進入STV16,STV17的情況
            If (Mid(brs5!cw_to, 2, 1) >= "a" And Mid(brs5!cw_to, 2, 1) >= "z") Or (Mid(brs5!cw_now, 2, 1) >= "a" And Mid(brs5!cw_now, 2, 1) >= "z") Then
              wbuf = Mid(brs5!cw_now, 1, 1) & "m": lib_buf = "2" 'Fm
            Else
              wbuf = Mid(brs5!cw_now, 1, 1) & "M": lib_buf = "1" 'FM
            End If
            brs3.Open "select * from st where st_awno='" & awno1 & "' and st_stno like '" & wbuf & "%' and st_cvar like '" & Left(lib_buf, 1) & "%'", rcn1, adOpenKeyset, adLockOptimistic
            If brs3!st_load = "1" Then
              If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
              Set brs3 = Nothing: GoTo rtn
            End If
            Set brs3 = Nothing
            If (brs5!cw_mveq = EQ_STV16 And mchno(brs5!cw_to) = 1) Or (brs5!cw_mveq = EQ_STV17 And mchno(brs5!cw_to) = 2) Then
              brs3.Open "select * from st where st_awno='" & awno1 & "' and st_stno like '" & Trim(brs5!cw_to) & "%' and st_cvar like '" & Left(lib_buf, 1) & "%'", rcn1, adOpenKeyset, adLockOptimistic
              If brs3!st_load = "1" Then
                If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                Set brs3 = Nothing: GoTo rtn
              End If
              Set brs3 = Nothing
            End If
          End If
        End If
        If (Mid(brs5!cw_now, 1, 2) = "GL" And Mid(brs5!cw_nxt, 1, 2) = "FL") Or (Mid(brs5!cw_now, 1, 2) = "FL" And Mid(brs5!cw_nxt, 1, 2) = "GL") Then GoTo rtn7
     ElseIf Left(project_name, 2) = "SK" Then
        If awno1 = "1" Then
           'STV3命令(A65-->A66)在mono程式,無STV4命令(已變成MONO命令)
           'If Left(brs5!cw_mveq, 4) = "STV3" Or Left(brs5!cw_mveq, 4) = "STV4" Then GoTo rtn10
           GoTo rtn
        ElseIf awno1 = "5" Then
           If Left(prog, 4) = "PLC5" And Not (Left(brs5!cw_mveq, 1) = "F" Or Left(brs5!cw_mveq, 5) = "STV13" Or Left(brs5!cw_mveq, 5) = "STV14") Then GoTo rtn
           If Left(prog, 4) = "PLC6" And Not (Left(brs5!cw_mveq, 1) = "E" Or Left(brs5!cw_mveq, 5) = "STV10") Then GoTo rtn
           If Left(brs5!cw_now, 3) = "E10" Then
             If conveyor_full(awno1, "E12", 2, brs5!cw_palt, brs5!cw_err, 10) Then GoTo rtn
           ElseIf Left(brs5!cw_now, 3) = "E22" Then
             wbuf = get_string_value1(awno1, EQ_STV11, "cr_loct")
             If wbuf <> "000001" And wbuf <> "000002" Then
               If brs5!cw_err <> 41 Then rcn.Execute "update cw set cw_err=41 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
           ElseIf Left(brs5!cw_now, 3) = "F01" Or Left(brs5!cw_now, 3) = "F10" Then
             If Left(brs5!cw_now, 3) = "F01" Then
                If conveyor_full(awno1, Left(brs5!cw_now, 3), 3, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
                If get_string_value1(awno1, "F10", "st_load") = "1" Or get_string_value1(awno1, "F11", "st_load") = "1" Then
                  If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                  GoTo rtn
                End If
                If get_string_value1(awno1, "F16", "st_load") = "1" Then
                  If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                  GoTo rtn
                End If
                If conveyor_full(awno1, "F11", 2, brs5!cw_palt, brs5!cw_err, 10) Then GoTo rtn
                'If check_block("F11     ", " ", ST_WORK) Then 'F12->F11
                brs3.Open "select cw_palt from cw where (cw_now='F12     ' and cw_stat='2') or (cw_now='F10     ' and (cw_stat='1' or cw_stat='2'))", rcn1, adOpenKeyset, adLockOptimistic
                If Not brs3.EOF Then
                  Set brs3 = Nothing
                  If brs5!cw_err <> 32 Then rcn.Execute "update cw set cw_err=32 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                  GoTo rtn
                End If
                Set brs3 = Nothing
             Else
                If conveyor_full(awno1, Left(brs5!cw_now, 3), 3, brs5!cw_palt, brs5!cw_err, 47, "+") Then GoTo rtn
                If conveyor_full(awno1, "F02", 2, brs5!cw_palt, brs5!cw_err, 10) Then GoTo rtn
             End If
             err1 = check_status(awno1, "STV12     ", brs5!cw_now, brs5!cw_nxt)
             If err1 > 0 Then
               If brs5!cw_err <> err1 Then rcn.Execute "update cw set cw_err=" & err1 & " where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn 'Exit Do
             End If
           ElseIf Left(brs5!cw_now, 3) = "F11" Then
             If conveyor_full(awno1, Left(brs5!cw_now, 3), 3, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
             If conveyor_full(awno1, "F19", 4, brs5!cw_palt, brs5!cw_err, 10) Then GoTo rtn
           ElseIf Left(brs5!cw_now, 3) = "F19" Then
             If conveyor_full(awno1, Left(brs5!cw_now, 3), 5, brs5!cw_palt, brs5!cw_err, 47) Then GoTo rtn
             brs3.Open "select cw_palt from cw where cw_mveq='CCL天車   '", rcn1, adOpenKeyset, adLockOptimistic
             If Not brs3.EOF Then
                Set brs3 = Nothing
                If brs5!cw_err <> 38 Then rcn.Execute "update cw set cw_err=38 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                GoTo rtn
             End If
             Set brs3 = Nothing
             brs3.Open "select cw_palt from cw where cw_now='F19     ' and cw_stat='2'", rcn1, adOpenKeyset, adLockOptimistic
             If Not brs3.EOF Then
                Set brs3 = Nothing
                If brs5!cw_err <> 28 Then rcn.Execute "update cw set cw_err=28 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                GoTo rtn
             End If
             Set brs3 = Nothing
             load1 = get_string_value1("5", EQ_STV14, "cr_load")
             If load1 = "1" Or load1 = "3" Then
               If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
           ElseIf Left(brs5!cw_nxt, 3) = "F15" Then '指F20-->F15 非 F20-->CCL
             If conveyor_full(awno1, "F12", 4, brs5!cw_palt, brs5!cw_err, 10) Then GoTo rtn
             load1 = get_string_value1("5", EQ_STV14, "cr_load")
             If load1 = "0" Or load1 = "1" Then
               If brs5!cw_err <> 3 Then rcn.Execute "update cw set cw_err=3 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
           ElseIf Left(brs5!cw_now, 3) = "F12" Then
              If get_string_value1(awno1, "F00", "st_load") = "1" Or get_string_value1(awno1, "F01", "st_load") = "1" Then
                'If brs5!cw_err <> 40 Then rcn.Execute "update cw set cw_err=40 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                'GoTo rtn
              End If
              If get_string_value1(awno1, "F10", "st_load") = "1" Or get_string_value1(awno1, "F11", "st_load") = "1" Then
                If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                GoTo rtn
              End If
              If get_string_value1(awno1, "F03", "st_load") = "1" Then
                If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                GoTo rtn
              End If
              If conveyor_full(awno1, Left(brs5!cw_now, 3), 5, brs5!cw_palt, brs5!cw_err, 47, "+") Then GoTo rtn
              If conveyor_full(awno1, "F10", 2, brs5!cw_palt, brs5!cw_err, 10) Then GoTo rtn
              brs3.Open "select cw_palt from cw where (cw_now='F01     ' and cw_stat='2') or (cw_now='F11     ' and (cw_stat='1' or cw_stat='2'))", rcn1, adOpenKeyset, adLockOptimistic
              If Not brs3.EOF Then
                 Set brs3 = Nothing
                 If brs5!cw_err <> 33 Then rcn.Execute "update cw set cw_err=33 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
                 GoTo rtn
              End If
              Set brs3 = Nothing
           ElseIf Left(brs5!cw_now, 3) = "F02" Then
             If conveyor_full(awno1, Left(brs5!cw_now, 3), 3, brs5!cw_palt, brs5!cw_err, 47, "+") Then GoTo rtn
           End If
        End If
     End If
     If brs5!cw_opno = TR_TEST Then GoTo rtn
     If is_stno(awno1, brs5!cw_now) Then
       brs3.Open "select * from cw where cw_now like '" & Trim(Left(brs5!cw_now, STNO_NUM)) & "%' order by cw_sitm", rcn1, adOpenKeyset, adLockOptimistic
       If Not brs3.EOF Then
         If brs5!cw_palt <> brs3!cw_palt Then
           If brs5!cw_err <> 48 Then rcn.Execute "update cw set cw_err=48 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           Set brs3 = Nothing: GoTo rtn 'Exit Do
         End If
       End If
       Set brs3 = Nothing
     End If
     brs3.Open "select * from cw where cw_mveq = '" & brs5!cw_mveq & "' and cw_stat='2'", rcn1, adOpenKeyset, adLockOptimistic
     If Not brs3.EOF Then
       If brs5!cw_err <> 49 Then rcn.Execute "update cw set cw_err=49 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
       Set brs3 = Nothing: GoTo rtn
     End If
     Set brs3 = Nothing
     err1 = check_status(awno1, brs5!cw_mveq, brs5!cw_now, brs5!cw_nxt)
     If err1 > 0 Then
       If Left(project_name, 2) = "SK" Then
         If Left(brs5!cw_now, 3) = "F19" And err1 = 69 Then 'STV14雙層台車F19-->F20,err1=69時有載狀態有3種(1,2,3),F19要走下層,STV14有載狀態0 or 2時應該允許下命令
           GoTo rtn8
         End If
       End If
       If brs5!cw_err <> err1 Then rcn.Execute "update cw set cw_err=" & err1 & " where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
       'Epoxy2,PLC5搬運命令包含G__->G__線及STV4-STV8,因之若某一筆失敗時,不可以Exit Do
       GoTo rtn 'Exit Do
     End If
rtn8:
     If get_string_value1(awno1, Trim(brs5!cw_now), "st_stat") <> "V" Then
        '起站須有載
       If Left(project_name, 2) = "SK" Then
         If Left(brs5!cw_now, 1) = "F" Then
           ii = 0
           Do
             If ii >= 5 Then Exit Do
             If get_string_value1(awno1, Left(brs5!cw_now, 3), "st_load") = NOLD Then
               If brs5!cw_err <> 3 Then rcn.Execute "update cw set cw_err=3 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
             ii = ii + 1: Call Wait(1)
           Loop
         End If
       Else
         If get_string_value1(awno1, Left(brs5!cw_now, 3), "st_load") = NOLD Then
           If brs5!cw_err <> 3 Then rcn.Execute "update cw set cw_err=3 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           GoTo rtn
         End If
       End If
     End If
     '迄站須無鎖定
     If check_block(brs5!cw_nxt, " ", ST_WORK) Then
        If brs5!cw_err <> 4 Then rcn.Execute "update cw set cw_err=4 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        GoTo rtn
     End If
     If get_string_value1(awno1, Trim(brs5!cw_nxt), "st_stat") <> "V" Then
        '迄站須無載
       If Left(project_name, 2) = "SK" Then
         If Left(brs5!cw_now, 1) = "F" Then
           ii = 0
           Do
             If ii >= 5 Then Exit Do
             If get_string_value1(awno1, Left(brs5!cw_nxt, STNO_NUM), "st_load") = "1" Then
               If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
               GoTo rtn
             End If
             ii = ii + 1: Call Wait(1)
           Loop
         End If
       Else
         If get_string_value1(awno1, Left(brs5!cw_nxt, STNO_NUM), "st_load") = "1" Then
           If brs5!cw_err <> 5 Then rcn.Execute "update cw set cw_err=5 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
           GoTo rtn
         End If
       End If
     End If
rtn7:
     palt_plc_write_asrs = brs5!cw_palt
     If Left(project_name, 2) = "SA" Then
        wbuf = "01A" & Mid(brs5!cw_now, 3, 1) & "A" & Mid(brs5!cw_nxt, 3, 1) & Chr(13) & Chr(10)
        If plc_write_asrs(port_id, 8, wbuf, brs5!cw_mveq, 4, awno1) = False Then GoTo rtn 'Exit Do
     ElseIf Left(project_name, 2) = "SF" Then
        'STV平常都在靠近天車的位置,要作入庫時需先在PLC操作盤上按一個叫車按鈕讓STV移出,然後再執行移入動作
        '不管STV是要作入或出,都下01B10B11,即先作入的動作
        wbuf = Chr(2) & "01" & Left(brs5!cw_now, 2) & "0" & Left(brs5!cw_nxt, 2) & "1" & Chr(13) & Chr(10)
        'wbuf = Chr(2) & "01" & Left(brs5!cw_now, 3) & Left(brs5!cw_nxt, 3) & Chr(13) & Chr(10)
        If plc_write_asrs(port_id, 11, wbuf, brs5!cw_mveq, 6, awno1) = False Then GoTo rtn 'Exit Do
     ElseIf Left(project_name, 2) = "SG" Then
        wbuf = Chr(2) & "01" & Left(brs5!cw_now, 3) & Left(brs5!cw_nxt, 3) & Chr(13) & Chr(10)
        If Mid(wbuf, 4, 6) = "C32C31" Then Mid(wbuf, 4, 6) = "C32C81"
        If plc_write_asrs(port_id, 11, wbuf, brs5!cw_mveq, 5, awno1) = False Then GoTo rtn 'Exit Do
     ElseIf Left(project_name, 4) = "SU01" Then
        Dim now2 As String * 2, nxt2 As String * 2
        now2 = Mid(brs5!cw_now, 1, 2): nxt2 = Mid(brs5!cw_nxt, 1, 2)
        If brs5!cw_mveq = EQ_STV16 Then
          If Mid(brs5!cw_now, 1, 2) = "GL" Then now2 = "FL"
          If Mid(brs5!cw_now, 1, 2) = "Gl" Then now2 = "Fl"
          If Mid(brs5!cw_nxt, 1, 2) = "GL" Then nxt2 = "FL"
          If Mid(brs5!cw_nxt, 1, 2) = "Gl" Then nxt2 = "Fl"
        End If
        wbuf = "01" & now2 & nxt2 & Chr(13) & Chr(10)
        If plc_write_asrs(port_id, 8, wbuf, brs5!cw_mveq, 4, awno1) = False Then GoTo rtn 'Exit Do
        If Left(brs5!cw_nxt, 2) = "GL" Or Left(brs5!cw_nxt, 2) = "Gl" Then
          If brs5!cw_mveq = EQ_STV16 Then
            Call put_cw(awno1, TR_MOVE_STNO, "GL      ", "FL      ", loginname, "N")
          End If
          If brs5!cw_mveq = EQ_STV17 Then
            Call put_cw(awno1, TR_MOVE_STNO, "FL      ", "GL      ", loginname, "N")
          End If
        End If
     Else
       For ii = 1 To 3
         If Left(project_name, 2) = "SK" And Left(brs5!cw_now, 3) = "F20" And Left(brs5!cw_nxt, 3) = "F15" Then
           wbuf = "01" & Left(brs5!cw_now, 3) & Left(brs5!cw_nxt, 3) & Right("00000" & CStr(brs5!cw_palt), 4) & Chr(13) & Chr(10)
           If plc_write_asrs(port_id, 14, wbuf, brs5!cw_mveq, 4, awno1) = True Then
             Exit For
           Else
             GoTo rtn 'Exit Do
           End If
         Else
           wbuf = "01" & Left(brs5!cw_now, 3) & Left(brs5!cw_nxt, 3) & Chr(13) & Chr(10)
  'Epoxy2,PLC5搬運命令包含G__->G__線及STV4-STV8,因之若某一筆失敗時,不可以Exit Do
           If plc_write_asrs(port_id, 10, wbuf, brs5!cw_mveq, 4, awno1) = True Then
             Exit For
           Else
             If Left(project_name, 2) = "SI" Then
               Call syslog1("ST_COMD_PLC", brs5!cw_palt, ii, Left(wbuf, 10) & ",send again")
               Wait (1)
             Else
               GoTo rtn 'Exit Do
             End If
           End If
         End If
       Next ii
     End If
     If (Left(project_name, 2) = "SI" And (Left(brs5!cw_nxt, 3) = "G26" Or Left(brs5!cw_nxt, 3) = "G04")) Or _
        (Left(project_name, 4) = "SU01" And ((Mid(brs5!cw_now, 1, 2) = "GL" And Mid(brs5!cw_nxt, 1, 2) = "FL") Or (Mid(brs5!cw_now, 1, 2) = "FL" And Mid(brs5!cw_nxt, 1, 2) = "GL"))) Then
        rcn.Execute "update cw set cw_stat='0',cw_err=0 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        Call Wait(5)
     Else
rtn10:
        rcn.Execute "update cw set cw_stat='2',cw_err=0 where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
     End If
     If Left(brs5!cw_mveq, 3) = "STV" Or Left(brs5!cw_mveq, 1) = "#" Then
        rcn.Execute "update cr set cr_err=0,cr_palt=" & brs5!cw_palt & ",cr_from='" & brs5!cw_now & "',cr_to='" & brs5!cw_nxt & "' where cr_awno='" & awno1 & "' and cr_mveq='" & brs5!cw_mveq & "'"
     Else
        rcn.Execute "update cr set cr_err=0,cr_palt=" & brs5!cw_palt & " where cr_awno='" & awno1 & "' and cr_mveq='" & brs5!cw_mveq & "'"
        rcn.Execute "update st set st_err=0,st_palt=" & brs5!cw_palt & " where st_awno='" & awno1 & "' and st_stno='" & Left(brs5!cw_mveq, STNO_NUM) & "'"
     End If
     If Left(project_name, 2) = "SG" And Left(brs5!cw_mveq, 4) = "STV1" Then
        rcn.Execute "update cr set cr_mvst='1' where cr_awno='" & awno1 & "' and cr_mveq='" & brs5!cw_mveq & "'"
     End If
     If Left(project_name, 2) = "SK" Then
        '在rd_data_plc中STV9搬運結束時才產生虛擬站D20命令
        If Left(brs5!cw_mveq, 4) = "STV9" And Left(brs5!cw_nxt, 3) = "D20" Then GoTo rtn3
        If Left(brs5!cw_nxt, 3) = "F20" Then GoTo rtn3
     End If
     '虛擬站命令分裂
     If Trim(brs5!cw_nxt) <> Trim(brs5!cw_to) And get_string_value1(awno1, brs5!cw_nxt, "st_stat") = "V" Then
  '原cw_srid改為TR_MOVE_STNO As String * 1 = "I"   '站間搬運
        rcn.Execute "update cw set cw_srid='I',cw_to='" & brs5!cw_nxt & "' where cw_awno='" & awno1 & "' and cw_palt=" & brs5!cw_palt
        ii = put_cw(awno1, brs5!cw_srid, brs5!cw_nxt, brs5!cw_to, brs5!cw_name, "N")
        If Left(project_name, 2) = "SI" Then
           rcn.Execute "update cw set cw_ptno='" & brs5!cw_ptno & "',cw_odno='" & brs5!cw_odno & "',cw_ltno='" & brs5!cw_ltno & "',cw_quty=" & brs5!cw_quty & ",cw_mark='" & brs5!cw_mark & "',cw_lot2='" & brs5!cw_lot2 & "',cw_quty2=" & brs5!cw_quty2 & ",cw_lot3='" & brs5!cw_lot3 & "',cw_quty3=" & brs5!cw_quty3 & " where cw_palt=" & ii
        End If
     End If
rtn3:
     Call syslog("st_comd_plc", 0, 0, "st_comd:palt=" & brs5!cw_palt & ",now=" & brs5!cw_now & ",nxt=" & brs5!cw_nxt)
     Exit Do
rtn:
     brs5.MoveNext
  Loop
  Set brs5 = Nothing
  Exit Function
err1_rtn:
  Call err2_rtn("st_comd_plc")
End Function

Public Sub mf2000_update(num_car1 As String, sts_ptn1 As String, lono As String)
  Dim brs3 As New ADODB.Recordset
  On Error GoTo err1_rtn
  Exit Sub
  If Left(project_name, 2) = "SJ" And Left(num_car1, 8) <> "00000000" Then
    If Trim(num_car1) <> "" Then
      If get_string_value1("1", "", "pa_every") = "Y" Then
      'If Not (UCase(Left(computername, 7)) = "TTY1A01" Or UCase(Left(computername, 7)) = "TTY1A10" Or UCase(Left(computername, 7)) = "TTY1A20" Or UCase(Left(computername, 7)) = "TTY1A30") Then
        err1_place = "mf2000_update"
        rdoEnvironments(0).CursorDriver = rdUseClientBatch: rdoEnvironments(0).UserName = "mf2000": rdoEnvironments(0).Password = "mf2000": Set rcn_bar = rdoEnvironments(0).OpenConnection("sqldb1", rdDriverNoPrompt, False)
        err1_place = ""
      End If
rtn10:
      If Left(lono, 1) = "0" Then
        rcn_bar.Execute "update prup set cod_loca='" & lono & "',sts_prup='" & Left(sts_ptn1, 2) & "' where num_car = '" & Trim(num_car1) & "' and sts_prup='10'"
      Else
        rcn_bar.Execute "update prup set sts_prup='" & Left(sts_ptn1, 2) & "' where num_car = '" & Trim(num_car1) & "' and sts_prup='10'"
      End If
      If get_string_value1("1", "", "pa_every") = "Y" Then
      'If Not (UCase(Left(computername, 7)) = "TTY1A01" Or UCase(Left(computername, 7)) = "TTY1A10" Or UCase(Left(computername, 7)) = "TTY1A20" Or UCase(Left(computername, 7)) = "TTY1A30") Then
        rcn_bar.Close
      End If
    End If
  End If
  Exit Sub
err1_rtn:
  If err1_place = "mf2000_update" Then
    err1_place = "": Call syslog1(err1_place, 0, 0, "mf2000_update"): Wait (2)
    rdoEnvironments(0).CursorDriver = rdUseClientBatch: rdoEnvironments(0).UserName = "mf2000": rdoEnvironments(0).Password = "mf2000"
    Set rcn_bar = rdoEnvironments(0).OpenConnection("sqldb1", rdDriverNoPrompt, False)
    GoTo rtn10
  Else
    MsgBox (prog & "-mf2000_update:聯合資訊ERP系統有問題無法連結,請檢查ERP系統狀況,行號=" & Erl)
  End If
  'If UCase(Left(computername, 7)) <> "TTY1A18" Then
  'End If
  'Call err2_rtn("mf2000_update")
End Sub
Public Function mchno(stno As String) As Integer
  mchno = 0
  If Mid(stno, 1, 1) = "F" Then mchno = 1
  If Mid(stno, 1, 1) = "G" And ((Mid(stno, 2, 1) >= "A" And Mid(stno, 2, 1) <= "Z") Or (Mid(stno, 2, 1) >= "a" And Mid(stno, 2, 1) <= "z")) Then mchno = 2
  If (Mid(stno, 1, 3) >= "C01" And Mid(stno, 1, 3) <= "C06") Or (Mid(stno, 1, 3) >= "D01" And Mid(stno, 1, 3) <= "D14") Then mchno = 3
  If (Mid(stno, 1, 3) >= "C07" And Mid(stno, 1, 3) <= "C12") Or (Mid(stno, 1, 3) >= "D15" And Mid(stno, 1, 3) <= "D28") Then mchno = 4
  If (Mid(stno, 1, 3) >= "G01" And Mid(stno, 1, 3) <= "G06") Or Mid(stno, 1, 3) = "G17" Or (Mid(stno, 1, 3) >= "E01" And Mid(stno, 1, 3) <= "E18") Then mchno = 5
End Function
Public Function Err_Translate_st(Err_Code As Integer) As String
  Err_Code = Abs(Err_Code)
'等待狀態,非異常
  If Err_Code = 0 Then
    Err_Translate_st = ""
  ElseIf Err_Code = 5 Then
    Err_Translate_st = "倉庫中空棧板數量已至下限,即將無法補充"
  ElseIf Err_Code = 6 Then
    Err_Translate_st = "需要補充空棧板但倉庫中沒有空棧板可以補充"
  ElseIf Err_Code = 7 Then
    Err_Translate_st = "倉庫中沒有空庫格可以入庫"
  ElseIf Err_Code = 10 Then
    Err_Translate_st = "變頻器異常檢知"
  ElseIf Err_Code = 20 Then
    Err_Translate_st = "輸送機馬達過載檢知"
  ElseIf Err_Code = 21 Then
    Err_Translate_st = "移載輸送機馬達過載檢知"
  ElseIf Err_Code = 22 Then
    Err_Translate_st = "油壓馬達過載檢知"
  ElseIf Err_Code = 30 Then
    Err_Translate_st = "輸送機運轉過久異常檢知"
  ElseIf Err_Code = 31 Then
    Err_Translate_st = "移載輸送機運轉過久異常檢知"
  ElseIf Err_Code = 40 Then
    Err_Translate_st = "防凸異常檢知"
  ElseIf Err_Code = 44 Then
    Err_Translate_st = "輸送機NOT READY狀態,無法下命令"
  ElseIf Err_Code = 50 Then
    Err_Translate_st = "極限異常檢知" '升降機上升極限
  ElseIf Err_Code = 51 Then
    Err_Translate_st = "升降機下降極限異常檢知"
  ElseIf Err_Code = 60 Then
    Err_Translate_st = "夾臂異常檢知"
  ElseIf Err_Code = 70 Then
    Err_Translate_st = "定位異常檢知"
  ElseIf Err_Code = 71 Then
    Err_Translate_st = "棧板分配器超高檢知"
  ElseIf Err_Code = 72 Then
    Err_Translate_st = "阻擋氣缸定位異常檢知"
  ElseIf Err_Code = 80 Then
    Err_Translate_st = "有序號無載異常檢知"
  ElseIf Err_Code = 81 Then
    Err_Translate_st = "有載無序號異常檢知"
  ElseIf Err_Code = 91 Then
    Err_Translate_st = "台車走行站號計數錯誤"
  ElseIf Err_Code = 92 Then
    Err_Translate_st = "台車走行前或後擋板動作"
  ElseIf Err_Code = 93 Then
    Err_Translate_st = "台車走行前或後極限動作"
  ElseIf Err_Code = 94 Then
    Err_Translate_st = "台車輸送左或右防凸異常"
  ElseIf Err_Code = 95 Then
    Err_Translate_st = "台車走行後極限動作"
  ElseIf Err_Code = 96 Then
    Err_Translate_st = "台車變頻器異常"
  ElseIf Err_Code = 97 Then
    Err_Translate_st = "台車走行馬達異常"
  ElseIf Err_Code = 98 Then
    Err_Translate_st = "台車緊急停止動作"
  ElseIf Err_Code = 99 Then
    Err_Translate_st = "台車交訊過久異常"
  ElseIf Err_Code = 100 Then
    Err_Translate_st = "台車派車命令錯誤"
  ElseIf Err_Code = 101 Then
    Err_Translate_st = "光資料傳送器異常"
  ElseIf Err_Code = 102 Then
    Err_Translate_st = "地面PLC要求緊急停止"
  ElseIf Err_Code = 103 Then
    Err_Translate_st = "地面PLC防凸動作"
  ElseIf Err_Code = 104 Then
    Err_Translate_st = "台車輸送馬達異常"
  ElseIf Err_Code = 105 Then
    Err_Translate_st = "台車走行馬達運轉過久異常"
  ElseIf Err_Code = 106 Then
    Err_Translate_st = "台車輸送馬達運轉過久異常"
  ElseIf Err_Code <> 0 Then
    Err_Translate_st = "異常說明未建檔"
  End If
End Function

Public Function Err_Translate(Err_Code As Integer) As String
  Err_Code = Abs(Err_Code)
'等待狀態,非異常
  If Err_Code = 0 Then
    Err_Translate = ""
  ElseIf Err_Code = 1 Then
    Err_Translate = "設備為搬運中狀態,因之命令暫時無法下達,請稍後"
  ElseIf Err_Code = 2 Then
    Err_Translate = "工作檔中正有(天車上有棧板或優先執行)的命令執行中,因之命令暫時無法下達,請稍後"
  ElseIf Err_Code = 3 Then
    Err_Translate = "起站為無載狀態,因之命令無法下達"
  ElseIf Err_Code = 4 Then
    Err_Translate = "工作檔中正有其它命令搬運棧板至下一站,因之命令暫時無法下達"
  ElseIf Err_Code = 5 Then
    Err_Translate = "迄站(或迄站前一站)為有載狀態,因之命令無法下達"
  ElseIf Err_Code = 6 Then
    Err_Translate = "需要補充空棧板但倉庫中沒有空棧板可以補充"
  ElseIf Err_Code = 7 Then
    Err_Translate = "倉庫中沒有空庫格可以入庫"
  ElseIf Err_Code = 8 Then
    Err_Translate = "STV為有載狀態,因之命令無法下達"
  ElseIf Err_Code = 9 Then
    Err_Translate = "此為優先執行之命令,但迄站為有載狀態,因之出庫命令無法下達,請注意這將會造成所有其它的入出庫命令都無法執行,因之請速排除有載狀態"
  ElseIf Err_Code = 10 Then
    Err_Translate = "工作檔中存在下一步的目前站命令數量已等於要移入的輸送機組節數,因之命令無法再下達"
  ElseIf Err_Code = 11 Then
    Err_Translate = "天車卸載完成,狀態轉移中"
  ElseIf Err_Code = 12 Then
    Err_Translate = "三截以上之輸送機組第一板命令無法下達,因之後續板命令也無法下達"
  ElseIf Err_Code = 13 Then
    Err_Translate = "STV有載或禁用,因之天車出庫命令無法下達"
  ElseIf Err_Code = 14 Then
    Err_Translate = "鐵料架收集分配器異常或分配時無載或收集時有載,因之天車命令無法下達"
'SK專案使用
  ElseIf Err_Code = 15 Then
    Err_Translate = "A66入庫旋轉輸送機或A69出庫旋轉輸送機作業中,因之天車命令無法下達"
  ElseIf Err_Code = 16 Then
    Err_Translate = "工作檔中有往STV3或STV4接近之天車命令,所以STV3或STV4命令無法下"
  ElseIf Err_Code = 17 Then
    Err_Translate = "STV3不在000002位置,所以STV3命令無法下"
  ElseIf Err_Code = 18 Then
    Err_Translate = "A65位置讀到bar code,但double check異常"
  ElseIf Err_Code = 19 Then
    Err_Translate = "A65位置沒讀到bar code異常"
  ElseIf Err_Code = 20 Then
    Err_Translate = "下A50-->A65命令,但MONO RAIL not ready"
  ElseIf Err_Code = 21 Then
    Err_Translate = "下A68-->A52命令,但MONO RAIL not ready"
  ElseIf Err_Code = 22 Then
    Err_Translate = "下A19-->A50(或A35)命令,但MONO PLC not ready"
  ElseIf Err_Code = 23 Then
    Err_Translate = "工作檔中有A66入庫旋轉輸送機或A69出庫旋轉輸送機命令,因之天車命令無法下達"
  ElseIf Err_Code = 24 Then
    Err_Translate = "命令下給MONO,但MONO逾時沒有回11"
  ElseIf Err_Code = 25 Then
    Err_Translate = "A94-A66不允許STV3下命令 或 A96-A66不允許天車下入庫命令"
  ElseIf Err_Code = 26 Then
    Err_Translate = "A95-A69不允許天車下出庫命令 或 A97-A69不允許STV4下命令"
  ElseIf Err_Code = 27 Then
    Err_Translate = "下給MONO RAIL的命令數已達MONO RAIL實際運轉數"
  ElseIf Err_Code = 28 Then
    Err_Translate = "工作檔中正有其它命令由此站搬出(或搬入)棧板,因之命令暫時無法下達"
  ElseIf Err_Code = 29 Then
    Err_Translate = "下給MONO的起站或迄站不正確命令無法下,請在工作檔中刪除此命令"
  ElseIf Err_Code = 30 Then
    Err_Translate = "天車在附近,所以STV3或STV4命令無法下"
  ElseIf Err_Code = 31 Then
    Err_Translate = "A50或STV4(A68站)上工作序號不為0,所以MONO命令無法下,請查看搬運設備或輸送機狀態查詢維護作業中A50或STV4(A68站)的工作序號"
'CCL地下道
  ElseIf Err_Code = 32 Then
    Err_Translate = "F12->F11線正入庫中,所以從F01出庫之STV13命令無法下"
  ElseIf Err_Code = 33 Then
    Err_Translate = "從F01出庫之STV13命令正出庫中,所以F12->F11線命令無法下"
  ElseIf Err_Code = 34 Then
    Err_Translate = "A50或A58或A56或A52或A68不是OK,無法下命令"
  ElseIf Err_Code = 35 Then
    Err_Translate = "CCL電腦回應本系統命令逾時異常"
  ElseIf Err_Code = 36 Then
    Err_Translate = "庫存資料不存在,無法產生命令給CCL倉庫"
  ElseIf Err_Code = 37 Then
    Err_Translate = "A50或A68站有一筆MONO命令已傳(或未傳)但MONO號卻未決定的命令"
  ElseIf Err_Code = 38 Then
    Err_Translate = "工作檔中已有CCL天車命令存在,因之命令暫時無法下達"
  ElseIf Err_Code = 39 Then
    Err_Translate = "F21 ON-允許天車下出庫至F00命令"
  ElseIf Err_Code = 40 Then
    Err_Translate = "F00或F01為有載狀態,必須等它搬完成無載狀態,因之命令暫時無法下達"
  ElseIf Err_Code = 41 Then
    Err_Translate = "STV11不在000001或000002位置,所以E22->E21命令無法下"
  ElseIf Err_Code = 42 Then
    Err_Translate = "命令下給PLC,但PLC回拒絕"
'
  ElseIf Err_Code = 43 Then
    Err_Translate = "輸送機組A19或A29或A35或45不是OK,無法下命令"
  ElseIf Err_Code = 44 Then
    Err_Translate = "輸送機NOT READY狀態,無法下命令"
  ElseIf Err_Code = 45 Then
    Err_Translate = "工作檔中有相同重覆的命令存在(命令已傳(或未傳)),請先處理掉"
  ElseIf Err_Code = 46 Then
    Err_Translate = "工作檔中正有其它命令向迄站方向移入,因之命令暫時無法下達"
  ElseIf Err_Code = 47 Then
    Err_Translate = "工作檔中存在的命令數量不等於本輸送機組上的棧板數,或目前站命令數量已大於本輸送機組節數,因之命令無法下達"
  ElseIf Err_Code = 48 Then
    Err_Translate = "工作檔中存在有投入工作檔時間較早(時間較小),必須先搬運或出庫的命令存在,因之命令暫時無法下達"
  ElseIf Err_Code = 49 Then
    Err_Translate = "工作檔中存在有此設備命令已傳,因之命令暫時無法下達"
  ElseIf Err_Code < 50 And Err_Code <> 0 Then
    Err_Translate = "等待狀態,命令稍候會自動執行"
'特別case異常碼
  ElseIf Err_Code = 51 Then
    Err_Translate = "警示:輸送機B09入庫或B59出庫作業,自動門B13或B63沒有打開,逾時異常"
  ElseIf Err_Code = 52 Then
    Err_Translate = "警示:STV或輸送機入出庫作業,天車叉牙沒有完成伸出再收回循環,逾時異常"
  ElseIf Err_Code = 53 Then
    Err_Translate = "警示:天車回00命令不接受異常碼,表示(1)禁用庫格或(2)命令超出範圍或(3)作業中或(4)手動"
  ElseIf Err_Code = 54 Then
    Err_Translate = "警示:BL入庫包裝檔案不存在(未傳入),無法執行入庫作業"
  ElseIf Err_Code = 55 Then
    Err_Translate = "警示:天車為禁用狀態,命令無法下達"
  ElseIf Err_Code = 56 Then
    Err_Translate = "警示:一期倉庫沒有空棧板可以補充或補充作業異常"
  ElseIf Err_Code = 57 Then
    Err_Translate = "警示:程式旗號鎖定後逾時,卻未解除鎖定異常"
  ElseIf Err_Code = 58 Then
    Err_Translate = "警示:天車取載要求,但STV(或輸送機)無載或STV不在靠近天車位置,無法下達取載同意"
  ElseIf Err_Code = 59 Then
    Err_Translate = "警示:天車卸載要求,但STV(或輸送機)有載或STV不在靠近天車位置,無法下達卸載同意"
'cw檔異常碼
  ElseIf Err_Code = 60 Then
    Err_Translate = "警示:設備代號不存在異常"
  ElseIf Err_Code = 61 Then
    Err_Translate = "警示:設備為禁用狀態,命令無法下達"
  ElseIf Err_Code = 62 Then
    Err_Translate = "警示:站號為禁用狀態,命令無法下達"
  ElseIf Err_Code = 63 Then
    Err_Translate = "警示:設備為手動狀態,命令無法下達"
  ElseIf Err_Code = 64 Then
    Err_Translate = "警示:輸入的起始站及終止站組合有問題,系統無法無法決定下一個搬運設備或下一站異常,請紀錄下此筆命令的詳細內容,反應給系統開發人員,然後按工作取消將此筆命令刪除"
  ElseIf Err_Code = 65 Then
    Err_Translate = "警示:通訊異常,連線設備回訊資料只有部份,不完整"
  ElseIf Err_Code = 66 Then
    Err_Translate = "警示:通訊異常,連線設備回訊資料檢查碼錯誤"
  ElseIf Err_Code = 67 Then
    Err_Translate = "警示:連線中斷異常,連線設備沒有回訊"
  ElseIf Err_Code = 68 Then
    Err_Translate = "警示:天車已卸載至輸送機但輸送機卻仍為無載狀態異常"
  ElseIf Err_Code = 69 Then
    Err_Translate = "警示:設備為有載狀態,命令無法下達"
  ElseIf Err_Code = 70 Then
    Err_Translate = "警示:站號為不能入出之狀態"
  ElseIf Err_Code = 71 Then
    Err_Translate = "搬運路徑已被另一個反向搬運之命令鎖住,暫時無法執行(較短路徑)"
  ElseIf Err_Code = 72 Then
    Err_Translate = "搬運路徑已被另一個同向搬運之命令鎖住,暫時無法執行(較短路徑)"
  ElseIf Err_Code = 73 Then
    Err_Translate = "搬運路徑已被另一個反向搬運之命令鎖住,暫時無法執行(較長路徑)"
  ElseIf Err_Code = 74 Then
    Err_Translate = "前往起站時會碰到前面的STV,所以命令暫時無法執行"
  ElseIf Err_Code = 75 Then
    Err_Translate = "PLC Timeout逾時"
  ElseIf Err_Code = 76 Then
    Err_Translate = "PLC Timeout逾時1"
  ElseIf Err_Code = 77 Then
    Err_Translate = "工作檔工作序號與輸送機工作序號不相同,所以命令無法執行"
  ElseIf Err_Code = 78 Then
    Err_Translate = "找不到可容納整櫃的空庫格天車可以入庫"
  ElseIf Err_Code = 79 Then
    Err_Translate = "該天車已被其他充填線入庫鎖定中"
  ElseIf Err_Code = 80 Then
    Err_Translate = "PLC不接受HOST命令(可能起站及迄站組合有問題),請刪除該命令"
'STV異常碼
  ElseIf Err_Code = 81 Then
    Err_Translate = "警示:STV或PLC入出庫命令型態不符合異常(Plc_check_crst),請取消STV入出庫命令,直接由天車作入出庫"
  ElseIf Err_Code = 82 Then
    Err_Translate = "棧板切換時暫停補充棧板"
  ElseIf Err_Code = 83 Then
    Err_Translate = "允入旗號必須為Y才能產生工作檔,若一直無法變Y,煩請先點選線別然後執行允入旗號改Y"
  ElseIf Err_Code = 84 Then
    Err_Translate = "警示:入庫命令,STV或PLC起站無載異常,搬運命令無法下達"
  ElseIf Err_Code = 85 Then
    Err_Translate = "警示:STV或PLC手動狀態異常,搬運命令無法下達"
  ElseIf Err_Code = 86 Then
    Err_Translate = "警示:STV或PLC未在起站位置異常,搬運命令無法下達"
  ElseIf Err_Code = 87 Then
    Err_Translate = "警示:STV或PLC緊急停止異常,搬運命令無法下達"
  ElseIf Err_Code = 88 Then
    Err_Translate = "警示:STV或PLC迄站有載異常,搬運命令無法下達"
  ElseIf Err_Code = 89 Then
    Err_Translate = "警示:出庫命令,STV或PLC起站有載異常,搬運命令無法下達"
  ElseIf Err_Code = 91 Then
    Err_Translate = "警示:棧板收集或分配器異常,搬運命令無法下達"
  ElseIf Err_Code = 90 Then
    Err_Translate = "警示:STV或PLC馬達過載,搬運命令無法下達"
  ElseIf Err_Code = 92 Then
    Err_Translate = "警示:前後安全bar碰到異常,搬運命令無法下達"
  ElseIf Err_Code = 93 Then
    Err_Translate = "警示:輸送過久異常,搬運命令無法下達"
  ElseIf Err_Code = 94 Then
    Err_Translate = "警示:走行過久異常,搬運命令無法下達"
  ElseIf Err_Code = 95 Then
    Err_Translate = "警示:兩側防凸異常,搬運命令無法下達"
  ElseIf Err_Code = 96 Then
    Err_Translate = "警示:與輸送機交訊過久異常,搬運命令無法下達"
  ElseIf Err_Code = 97 Then
    Err_Translate = "警示:變頻器異常,搬運命令無法下達"
  ElseIf Err_Code = 98 Then
    Err_Translate = "警示:前後走行極限異常,搬運命令無法下達"
  ElseIf Err_Code = 99 Then
    If Left(project_name, 2) = "SI" Then
      Err_Translate = "警示:PLC未準備好,命令暫不接受"
    Else
      Err_Translate = "警示:後走行極限異常,搬運命令無法下達"
    End If
  ElseIf Err_Code = 100 Then
    If Left(project_name, 4) = "SK02" Then
      Err_Translate = "異常:緊急停止裝置動作" 'awno="5",天車異常碼-Haushahn
    Else
      Err_Translate = "換車等待時間不足,搬運命令無法下達"
    End If
  ElseIf Err_Code = 101 Then
    Err_Translate = "AGV搬運命令被取消"
'天車異常碼-Haushahn
  ElseIf Err_Code = 102 Then
    Err_Translate = "異常:天車緊急停止"
  ElseIf Err_Code = 110 Then
    Err_Translate = "警示:禁用庫格"
  ElseIf Err_Code = 111 Then
    Err_Translate = "警示:命令超出設定庫格"
  ElseIf Err_Code = 114 Then
    Err_Translate = "警示:光資料傳送器異常"
  ElseIf Err_Code = 120 Then
    Err_Translate = "異常:前進L.S動作 走行超過前極限"
  ElseIf Err_Code = 121 Then
    Err_Translate = "異常:後退L.S動作 走行超過後極限"
  ElseIf Err_Code = 122 Then
    Err_Translate = "異常:上昇L.S動作 昇降超過上極限"
  ElseIf Err_Code = 123 Then
    Err_Translate = "異常:下降L.S動作 昇降超過下極限"
  ElseIf Err_Code = 124 Then
    Err_Translate = "異常:昇降鍊條異常"
  ElseIf Err_Code = 130 Then
    If Left(project_name, 4) = "SK02" Then
      Err_Translate = "異常:天車本體強迫開關被切在ON"
    Else
      Err_Translate = "異常:走行SENSOR異常"
    End If
  ElseIf Err_Code = 131 Then
    Err_Translate = "異常:走行鐵片計數異常"
  ElseIf Err_Code = 132 Then
    Err_Translate = "異常:走行SENSOR異常"
  ElseIf Err_Code = 133 Then
    Err_Translate = "異常:昇降SENSOR異常"
  ElseIf Err_Code = 140 Then
    Err_Translate = "異常:走行變頻器異常"
  ElseIf Err_Code = 141 Then
    Err_Translate = "異常:昇降變頻器異常"
  ElseIf Err_Code = 150 Then
    Err_Translate = "異常:緊急停止"
  ElseIf Err_Code = 151 Then
    Err_Translate = "警示:FROM作業有載"
  ElseIf Err_Code = 152 Then
    Err_Translate = "警示:TO作業無載"
  ElseIf Err_Code = 153 Then  '空出荷
    Err_Translate = "警示:空出荷(該庫位並無庫存可出庫)"
  ElseIf Err_Code = 154 Then  '先入品
    Err_Translate = "警示:先入品(該庫位已有物品無法入庫)"
  ElseIf Err_Code = 155 Then
    Err_Translate = "警示:交訊時間過久"
  ElseIf Err_Code = 156 Then
    Err_Translate = "警示:取載吊車有載"
  ElseIf Err_Code = 157 Then
    Err_Translate = "警示:取載該站無載"
  ElseIf Err_Code = 160 Then
    Err_Translate = "異常:走行站超過設定範圍"
  ElseIf Err_Code = 161 Then
    Err_Translate = "異常:走行時間過久"
  ElseIf Err_Code = 162 Then
    Err_Translate = "異常:叉牙未在中心"
  ElseIf Err_Code = 170 Then
    Err_Translate = "異常:昇降站超過設定範圍"
  ElseIf Err_Code = 171 Then
    Err_Translate = "異常:昇降時間過久"
  ElseIf Err_Code = 172 Then
    Err_Translate = "異常:叉牙未在中心"
  ElseIf Err_Code = 173 Then
    Err_Translate = "異常:上→←下定位時間過久"
  ElseIf Err_Code = 180 Then
    Err_Translate = "異常:叉牙站超過設定範圍"
  ElseIf Err_Code = 181 Then
    Err_Translate = "異常:叉牙時間過久"
  ElseIf Err_Code = 182 Then
    Err_Translate = "異常:取載未在下定位"
  ElseIf Err_Code = 183 Then
    Err_Translate = "異常:卸載未在上定位"
  ElseIf Err_Code = 184 Then
    Err_Translate = "異常:取載走行未在站上"
  ElseIf Err_Code = 185 Then
    Err_Translate = "異常:卸載走行未在站上"
  ElseIf Err_Code = 186 Then
    Err_Translate = "警示:光資料傳送器異常"
  ElseIf Err_Code = 187 Then
    Err_Translate = "異常:昇降馬達與Encoder方向不同異常"
  ElseIf Err_Code = 188 Then
    Err_Translate = "異常:Encoder計數異常"
  ElseIf Err_Code = 228 Then
    Err_Translate = "異常:昇降馬達超重"
  ElseIf Err_Code = 307 Then
    Err_Translate = "異常:天車煞車器過熱"
  ElseIf Err_Code = 408 Then
    Err_Translate = "異常:棧板規格錯誤"
  ElseIf Err_Code = 409 Then
    Err_Translate = "異常:棧板定位不良"
  ElseIf Err_Code = 410 Or Err_Code = 412 Then
    Err_Translate = "異常:叉牙伸出時Gap free sensor仍ON"
  ElseIf Err_Code = 414 Then
    Err_Translate = "異常:載物超邊"
  ElseIf Err_Code = 418 Then
    Err_Translate = "異常:走行期間叉牙Encoder值未在中心"
  ElseIf Err_Code = 419 Then
    Err_Translate = "異常:天車於取卸載過程中升降Encoder值變化過大"
  ElseIf Err_Code = 430 Then
    Err_Translate = "警示:左側庫格沒有棧板"
  ElseIf Err_Code = 431 Then
    Err_Translate = "警示:左側輸送機沒有棧板"
  ElseIf Err_Code = 432 Then
    Err_Translate = "警示:右側庫格沒有棧板"
  ElseIf Err_Code = 433 Then
    Err_Translate = "警示:右側輸送機沒有棧板"
  ElseIf Err_Code = 434 Then
    Err_Translate = "警示:左側靠近天車庫格已有棧板"
  ElseIf Err_Code = 435 Then
    Err_Translate = "警示:左側輸送機先入品"
  ElseIf Err_Code = 436 Then
    Err_Translate = "警示:右側靠近天車庫格已有棧板"
  ElseIf Err_Code = 437 Then
    Err_Translate = "警示:右側輸送機先入品"
  ElseIf Err_Code = 438 Then
    Err_Translate = "異常左側載物品歪斜"
  ElseIf Err_Code = 439 Then
    Err_Translate = "異常:右側載物品歪斜"
  ElseIf Err_Code = 440 Then
    Err_Translate = "異常:左側物品超寬"
  ElseIf Err_Code = 442 Then
    Err_Translate = "異常:右側物品超寬"
  ElseIf Err_Code = 444 Then
    Err_Translate = "異常:左側物品超高"
  ElseIf Err_Code = 450 Then
    Err_Translate = "異常:右側物品超高"
  ElseIf Err_Code = 452 Then
    Err_Translate = "警示:右側遠離天車庫格已有棧板"
  ElseIf Err_Code = 454 Then
    Err_Translate = "警示:左側遠離天車庫格已有棧板"
  ElseIf Err_Code = 455 Then
    Err_Translate = "警示:靠近天車庫格已有棧板;無法存取遠離天車庫格"
  ElseIf Err_Code = 460 Then
    Err_Translate = "警示:PLC沒有回應天車訊息"
  ElseIf Err_Code = 470 Then
    Err_Translate = "異常:天車於入庫站取載時,載物超邊並將布梱送回至輸送帶上"
  ElseIf Err_Code = 471 Then
    Err_Translate = "異常:天車於庫格內取載時載物超邊"
  ElseIf Err_Code >= 1400 And Err_Code <= 1499 Then
    Err_Translate = "異常:Interbus error 通訊失敗"
  ElseIf Err_Code = 9908 Then
    Err_Translate = "警示:庫格禁用,天車不接受命令"
'天車異常碼-王海朋
  ElseIf Err_Code = 600 Then
    Err_Translate = "警示:命令執行中無法接受新命令異常"
  ElseIf Err_Code = 601 Then
    Err_Translate = "警示:光資料傳送器異常"
  ElseIf Err_Code = 602 Then
    If Left(project_name, 4) <> "SU01" Then
      Err_Translate = "警示:禁用站異常"
    Else
      Err_Translate = "警示:From命令格式不符異常"
    End If
  ElseIf Err_Code = 603 Then
    If Left(project_name, 4) <> "SU01" Then
      Err_Translate = "警示:命令超出設定庫格異常"
    Else
      Err_Translate = "警示:to命令格式不符異常"
    End If
  ElseIf Err_Code = 604 Then '吊車作業中又下另一道命令
    Err_Translate = "異常:FROM作業吊車已有載異常"
  ElseIf Err_Code = 605 Then '空出荷
    Err_Translate = "警示:空出荷,取載該站無載異常"
    'Err_Translate = "異常:TO作業吊車無載異常"
  ElseIf Err_Code = 606 Then '空出荷
    Err_Translate = "警示:空出荷,取載該站無載異常"
  ElseIf Err_Code = 607 Then
    Err_Translate = "異常:卸載後吊車仍有載異常"
  ElseIf Err_Code = 608 Then
    If Left(project_name, 4) <> "SU01" Then
      Err_Translate = "警示:大棧板入小庫格異常"
      'Err_Translate = "異常:叉牙取載作業吊車已有載異常"
    Else
      Err_Translate = "警示:高載物欲入低庫位異常"
    End If
  ElseIf Err_Code = 609 Then
    Err_Translate = "異常:叉牙取載走行未在站上異常"
  ElseIf Err_Code = 610 Then
    Err_Translate = "異常:叉牙取載昇降未在下定位異常"
  ElseIf Err_Code = 611 Then
    Err_Translate = "異常:叉牙卸載作業吊車無載異常"
  ElseIf Err_Code = 612 Then
    Err_Translate = "異常:叉牙卸載走行未在站上異常"
  ElseIf Err_Code = 613 Then
    Err_Translate = "異常:叉牙卸載昇降未在上定位異常"
  ElseIf Err_Code = 614 Then '先入品,或吊車欲卸載輸送機而該站有載異常
    Err_Translate = "警示:先入品,或吊車欲卸載輸送機而該站有載異常"
  ElseIf Err_Code = 615 Then
    Err_Translate = "異常:緊急停止異常"
  ElseIf Err_Code = 616 Then
    Err_Translate = "警示:取載為禁用庫格異常"
  ElseIf Err_Code = 617 Then
    If Left(project_name, 4) <> "SU01" Then
      Err_Translate = "警示:取載為禁用庫格異常"
    Else
      Err_Translate = "警示:卸載為禁用庫格異常"
    End If
  ElseIf Err_Code = 618 Then
    If Left(project_name, 4) <> "SU01" Then
      Err_Translate = "警示:卸載禁用庫格異常"
    Else
      Err_Translate = "警示:取載為禁用站異常"
    End If
  ElseIf Err_Code = 619 Then
    Err_Translate = "異常:卸載為禁用站異常"
  ElseIf Err_Code = 620 Then
    Err_Translate = "異常:雷射測距值與設定值差異過大異常"
  ElseIf Err_Code = 621 Then
    Err_Translate = "異常:雷射測距儀異常,或走行計數異常"
  ElseIf Err_Code = 622 Then
    Err_Translate = "異常:走行超出前範圍異常"
  ElseIf Err_Code = 623 Then
    Err_Translate = "異常:走行超出後範圍異常"
  ElseIf Err_Code = 624 Then
    Err_Translate = "異常:走行前極限開關動作異常"
  ElseIf Err_Code = 625 Then
    Err_Translate = "異常:走行後極限開關動作異常"
  ElseIf Err_Code = 626 Then
    Err_Translate = "異常:走行變頻器異常"
  ElseIf Err_Code = 627 Then
    Err_Translate = "異常:雷射通訊異常,或雷射測距儀異常"
  ElseIf Err_Code = 628 Then
    Err_Translate = "異常:雷射傳送資料異常"
  ElseIf Err_Code = 629 Then
    Err_Translate = "異常:走行時間過久異常"
  ElseIf Err_Code = 630 Then
    Err_Translate = "異常:昇降鐵片異常"
  ElseIf Err_Code = 631 Then
    Err_Translate = "異常:昇降編碼器異常,昇降計數異常"
  ElseIf Err_Code = 632 Then
    Err_Translate = "異常:昇降超出上範圍異常"
  ElseIf Err_Code = 633 Then
    Err_Translate = "異常:昇降超出下範圍異常"
  ElseIf Err_Code = 634 Then
    Err_Translate = "異常:昇降上極限開關動作異常"
  ElseIf Err_Code = 635 Then
    Err_Translate = "異常:昇降下極限開關動作異常"
  ElseIf Err_Code = 636 Then
    Err_Translate = "異常:昇降變頻器異常"
  ElseIf Err_Code = 637 Then
    If Left(project_name, 4) <> "SU01" Then
      Err_Translate = "異常:上→←下定位時間過久異常"
    Else
      Err_Translate = "異常:上下定位昇降時間過久異常"
    End If
  ElseIf Err_Code = 638 Then
    Err_Translate = "異常:昇降鍊條異常"
  ElseIf Err_Code = 639 Then
    Err_Translate = "異常:昇降時間過久異常"
  ElseIf Err_Code = 641 Then
    Err_Translate = "異常:叉牙編碼器異常,叉牙計數異常"
  ElseIf Err_Code = 642 Then
    Err_Translate = "異常:叉牙超出左範圍異常"
  ElseIf Err_Code = 643 Then
    Err_Translate = "異常:叉牙超出右範圍異常"
  ElseIf Err_Code = 644 Then
    Err_Translate = "異常:叉牙未在中心異常"
  ElseIf Err_Code = 645 Then
    Err_Translate = "異常:左側物品超邊異常"
  ElseIf Err_Code = 646 Then
    Err_Translate = "異常:右側物品超邊異常"
  ElseIf Err_Code = 647 Then
    Err_Translate = "異常:左側卸載物品超高異常"
  ElseIf Err_Code = 648 Then
    If Left(project_name, 4) <> "SU01" Then
      Err_Translate = "異常:載物歪斜異常"
    Else
      Err_Translate = "警示:開關箱門被打開異常"
    End If
  ElseIf Err_Code = 649 Then
    Err_Translate = "異常:叉牙時間過久異常"
  ElseIf Err_Code = 650 Then
    Err_Translate = "警示:天車與STV交訊時間過久(取載或卸載時)"
  ElseIf Err_Code >= 651 And Err_Code <= 699 Then
    If Left(project_name, 4) <> "SU01" Then
      If Err_Code = 651 Then
        Err_Translate = "警示:D/A ERROR"
      ElseIf Err_Code = 652 Then
        Err_Translate = "警示:昇降煞車接點故障異常"
      ElseIf Err_Code = 653 Then
        Err_Translate = "警示:叉牙變頻器異常"
      ElseIf Err_Code = 654 Then
        Err_Translate = "警示:叉牙SERVOPACK Error"
      ElseIf Err_Code = 655 Then
        Err_Translate = "警示:記憶體無資料異常"
      End If
    Else
      If Err_Code = 651 Then
        Err_Translate = "異常:命令From站超出範圍"
      ElseIf Err_Code = 652 Then
        Err_Translate = "異常:命令From庫超出範圍異常"
      ElseIf Err_Code = 653 Then
        Err_Translate = "異常:命令to站超出範圍異常"
      ElseIf Err_Code = 654 Then
        Err_Translate = "異常:命令to庫超出範圍異常"
      ElseIf Err_Code = 655 Then
        Err_Translate = "異常:無載但棧板歪斜有載不匹配異常"
      ElseIf Err_Code = 656 Then
        Err_Translate = "異常:載物歪斜異常"
      ElseIf Err_Code = 657 Then
        Err_Translate = "異常:上下定位差超過限制異常"
      ElseIf Err_Code = 660 Then
        Err_Translate = "異常:昇降歸原點下降逾時異常"
      ElseIf Err_Code = 661 Then
        Err_Translate = "異常:昇降歸原點上升逾時異常"
      ElseIf Err_Code = 662 Then
        Err_Translate = "異常:叉牙歸原點向左逾時異常"
      ElseIf Err_Code = 663 Then
        Err_Translate = "異常:叉牙歸原點向右逾時異常"
      ElseIf Err_Code = 664 Then
        Err_Translate = "異常:出庫站為窄站異常"
      ElseIf Err_Code = 665 Then
        Err_Translate = "異常:出庫站為低站異常"
      ElseIf Err_Code = 666 Then
        Err_Translate = "異常:前檢查點異常"
      ElseIf Err_Code = 667 Then
        Err_Translate = "異常:後檢查點異常"
      ElseIf Err_Code = 680 Then
        Err_Translate = "異常:搬運狀態回異常"
      End If
    End If
'RS6000系統中自訂異常碼
  ElseIf Err_Code = 801 Or Err_Code = 1701 Then
    Err_Translate = "警示:設備為禁用狀態,請解除禁用狀態"
  ElseIf Err_Code = 802 Then
    Err_Translate = "警示:設備為手動狀態,請轉成自動狀態"
  ElseIf Err_Code = 803 Then
    Err_Translate = "警示:設備為搬運中狀態"
  ElseIf Err_Code = 1702 Then
    Err_Translate = "警示:設備為無載狀態,請檢查"
  ElseIf Err_Code = 804 Or Err_Code = 1703 Then
    Err_Translate = "警示:設備為有載狀態,請檢查"
  ElseIf Err_Code >= 910 And Err_Code <= 919 Then
    Err_Translate = "警示:設備連線中斷"
'天車異常碼-葉樹堅
  ElseIf Err_Code = 3000 Then '空出荷
    Err_Translate = "警示:空出荷,吊車至庫格載不到棧板"
  ElseIf Err_Code = 3001 Then '先入品
    Err_Translate = "警示:先入品,左庫格已有棧板"
  ElseIf Err_Code = 3002 Then '先入品
    Err_Translate = "警示:先入品,右庫格已有棧板"
  ElseIf Err_Code = 3003 Then
    Err_Translate = "警示:吊車已有棧板,無法再載棧板"
  ElseIf Err_Code = 3004 Then
    Err_Translate = "異常:吊車欲載棧板而不在下定位"
  ElseIf Err_Code = 3005 Then
    Err_Translate = "異常:吊車欲放棧板而不在上定位"
  ElseIf Err_Code = 3006 Then
    Err_Translate = "異常:叉牙伸縮時不在上定位"
  ElseIf Err_Code = 3007 Then
    Err_Translate = "異常:叉牙伸縮時不在下定位"
  ElseIf Err_Code = 3008 Then
    Err_Translate = "警示:吊車欲放棧板而感測無棧板"
  ElseIf Err_Code = 3009 Then
    Err_Translate = "異常:卸載後叉牙收回中心,仍感應有載"
  ElseIf Err_Code = 3010 Then
    Err_Translate = "異常:叉牙未在中心"
  ElseIf Err_Code = 3011 Then
    Err_Translate = "異常:吊車走行未在定位"
  ElseIf Err_Code = 3012 Then
    Err_Translate = "異常:吊車昇降未在上定位"
  ElseIf Err_Code = 3013 Then
    Err_Translate = "異常:吊車昇降未在下定位"
  ElseIf Err_Code = 3014 Then
    Err_Translate = "異常:吊車緊急停止"
  ElseIf Err_Code = 3015 Then
    Err_Translate = "異常:前柱安全裝置動作"
  ElseIf Err_Code = 3016 Then
    Err_Translate = "異常:後柱安全裝置動作"
  ElseIf Err_Code = 3017 Then
    Err_Translate = "異常:左右叉牙中心確認光電開關異常"
  ElseIf Err_Code = 3018 Then
    Err_Translate = "異常:走行馬達電流過載"
  ElseIf Err_Code = 3019 Then
    Err_Translate = "異常:昇降馬達電流過載"
  ElseIf Err_Code = 3020 Then
    Err_Translate = "異常:叉牙馬達電流過載"
  ElseIf Err_Code = 3021 Then
    Err_Translate = "異常:走行馬達溫度過熱"
  ElseIf Err_Code = 3022 Then
    Err_Translate = "異常:昇降馬達溫度過熱"
  ElseIf Err_Code = 3023 Then
    Err_Translate = "異常:叉牙馬達溫度過熱"
  ElseIf Err_Code = 3024 Then
    Err_Translate = "異常:變頻器異常"
  ElseIf Err_Code = 3025 Then
    Err_Translate = "警示:出庫時,載物尺寸超出出庫站尺寸"
  ElseIf Err_Code = 3026 Then
    Err_Translate = "異常:右叉牙時間過久"
  ElseIf Err_Code = 3027 Then
    Err_Translate = "異常:右叉牙時間過久"
  ElseIf Err_Code = 3028 Then
    Err_Translate = "異常:右叉牙時間過久"
  ElseIf Err_Code = 3029 Then
    Err_Translate = "異常:右叉牙時間過久"
  ElseIf Err_Code = 3030 Then
    Err_Translate = "異常:右叉牙時間過久"
  ElseIf Err_Code = 3031 Then
    Err_Translate = "異常:右叉牙時間過久"
  ElseIf Err_Code = 3032 Then
    Err_Translate = "異常:左叉牙時間過久"
  ElseIf Err_Code = 3033 Then
    Err_Translate = "異常:左叉牙時間過久"
  ElseIf Err_Code = 3034 Then
    Err_Translate = "異常:左叉牙時間過久"
  ElseIf Err_Code = 3035 Then
    Err_Translate = "異常:左叉牙時間過久"
  ElseIf Err_Code = 3036 Then
    Err_Translate = "異常:左叉牙時間過久"
  ElseIf Err_Code = 3037 Then
    Err_Translate = "異常:左叉牙時間過久"
  ElseIf Err_Code = 3038 Then
    Err_Translate = "異常:叉牙馬達運轉過久"
  ElseIf Err_Code = 3039 Then
    Err_Translate = "異常:走行馬達運轉過久"
  ElseIf Err_Code = 3040 Then
    Err_Translate = "異常:昇降馬達運轉過久"
  ElseIf Err_Code = 3041 Then
    Err_Translate = "異常:執行FROM TO CYCLE TIME過久"
  ElseIf Err_Code = 3042 Then
    Err_Translate = "異常:走行前進極限動作"
  ElseIf Err_Code = 3043 Then
    Err_Translate = "異常:走行後退極限動作"
  ElseIf Err_Code = 3044 Then
    Err_Translate = "異常:上升極限動作"
  ElseIf Err_Code = 3045 Then
    Err_Translate = "異常:下降極限動作"
  ElseIf Err_Code = 3046 Then
    Err_Translate = "警示:小棧板入中庫格(水平)"
  ElseIf Err_Code = 3047 Then
    Err_Translate = "警示:小棧板入大庫格(水平)"
  ElseIf Err_Code = 3048 Then
    Err_Translate = "警示:中棧板入小庫格(水平)"
  ElseIf Err_Code = 3049 Then
    Err_Translate = "警示:中棧板入大庫格(水平)"
  ElseIf Err_Code = 3050 Then
    Err_Translate = "警示:大棧板入小庫格(水平)"
  ElseIf Err_Code = 3051 Then
    Err_Translate = "警示:大棧板入中庫格(水平)"
  ElseIf Err_Code = 3052 Then
    Err_Translate = "警示:中棧板入低庫格(垂直)"
  ElseIf Err_Code = 3053 Then
    Err_Translate = "警示:高棧板入低庫格(垂直)"
  ElseIf Err_Code = 3054 Then
    Err_Translate = "警示:高棧板入中庫格(垂直)"
  ElseIf Err_Code = 3055 Then
    Err_Translate = "異常:由每經過水平定位鐵片檢測走行ENCODER數值是否誤差過大"
  ElseIf Err_Code = 3056 Then
    Err_Translate = "異常:由定時比對走行ENCODER數值是否與定位片相符"
  ElseIf Err_Code = 3057 Then
    Err_Translate = "異常:走行ENCODER超過FROM TO座標範圍"
  ElseIf Err_Code = 3058 Then
    Err_Translate = "異常:走行ENCODER超過前後極限座標"
  ElseIf Err_Code = 3059 Then
    Err_Translate = "異常:走行運轉中單位時間內ENCODER數值變化太大"
  ElseIf Err_Code = 3060 Then
    Err_Translate = "異常:走行運轉中單位時間內ENCODER數值均未變化"
  ElseIf Err_Code = 3061 Then
    Err_Translate = "異常:由每經過水平定位鐵片檢測昇降ENCODER數值是否誤差過大"
  ElseIf Err_Code = 3062 Then
    Err_Translate = "異常:由定時比對昇降ENCODER數值是否與定位片相符"
  ElseIf Err_Code = 3063 Then
    Err_Translate = "異常:昇降ENCODER超過FROM TO座標範圍"
  ElseIf Err_Code = 3064 Then
    Err_Translate = "異常:昇降ENCODER超過前後極限座標"
  ElseIf Err_Code = 3065 Then
    Err_Translate = "異常:昇降運轉中單位時間內ENCODER數值變化太大"
  ElseIf Err_Code = 3066 Then
    Err_Translate = "異常:昇降運轉中單位時間內ENCODER數值均未變化"
  ElseIf Err_Code = 3067 Then
    Err_Translate = "警示:前光資料傳送器被遮住"
  ElseIf Err_Code = 3068 Then
    Err_Translate = "警示:後光資料傳送器被遮住"
  ElseIf Err_Code = 3069 Then
    Err_Translate = "警示:載物超出昇降台右側"
  ElseIf Err_Code = 3070 Then
    Err_Translate = "警示:載物超出昇降台左側"
  ElseIf Err_Code = 3071 Then
    Err_Translate = "警示:載物超高檢出"
  ElseIf Err_Code = 3072 Then
    Err_Translate = "警示:棧板歪斜檢出"
  ElseIf Err_Code = 3073 Then
    Err_Translate = "異常:走行ENCODER異常"
  ElseIf Err_Code = 3074 Then
    Err_Translate = "異常:昇降ENCODER異常"
  ElseIf Err_Code = 3075 Then
    Err_Translate = "異常:叉牙ENCODER異常"
  ElseIf Err_Code = 3076 Then
    Err_Translate = "異常:上定位叉牙循環前感測無棧板"
  ElseIf Err_Code = 3077 Then
    Err_Translate = "警示:與出入庫站交訊過久"
  ElseIf Err_Code = 3078 Then
    Err_Translate = "異常:叉牙循環超過下定位"
  ElseIf Err_Code = 3079 Then
    Err_Translate = "異常:叉牙循環超過上定位"
  ElseIf Err_Code = 3080 Then
    Err_Translate = "警示:吊車未停止即切換操作模式"
  ElseIf Err_Code = 3081 Then
    Err_Translate = "異常:叉牙歸位異常"
  ElseIf Err_Code = 3082 Then
    Err_Translate = "異常:昇降歸位異常"
  ElseIf Err_Code = 3083 Then
    Err_Translate = "異常:走行歸位異常"
  ElseIf Err_Code = 3084 Then
    Err_Translate = "異常:變頻器周波數檢出異常"
  ElseIf Err_Code = 3085 Then
    Err_Translate = "異常:昇降台載物交叉確認異常"
  ElseIf Err_Code = 3086 Then
    Err_Translate = "異常:走行自動執行CYCLE TIME過久"
  ElseIf Err_Code = 3087 Then
    Err_Translate = "異常:昇降自動執行CYCLE TIME過久"
  ElseIf Err_Code = 3088 Then
    Err_Translate = "異常:叉牙自動執行CYCLE TIME過久"
  ElseIf Err_Code = 3089 Then
    Err_Translate = "異常:只作MOVE移動命令之前,吊車已有載"
  ElseIf Err_Code = 3090 Then
    Err_Translate = "異常:只作盤點命令之前,吊車已有載"
  ElseIf Err_Code = 3091 Then
    Err_Translate = "異常:FROM TO命令無法執行"
  ElseIf Err_Code = 3092 Then
    Err_Translate = "異常:叉牙運轉中單位時間內ENCODER數值變化太大"
  ElseIf Err_Code = 3093 Then
    Err_Translate = "異常:叉牙運轉中單位時間內ENCODER數值均未變化"
  ElseIf Err_Code = 3094 Then
    Err_Translate = "異常:走行計碼器定位不良"
  ElseIf Err_Code = 3095 Then
    Err_Translate = "異常:昇降計碼器定位不良"
  ElseIf Err_Code = 3096 Then
    Err_Translate = "異常:叉牙計碼器定位不良"
  ElseIf Err_Code = 3097 Then
    Err_Translate = "異常:叉牙運動中載物超出左右側"
  ElseIf Err_Code = 3098 Then
    Err_Translate = "異常:走行鐵片自我學習異常(至下一鐵片時間過久)"
  ElseIf Err_Code = 3099 Then
    Err_Translate = "異常:走行鐵片自我學習異常(過早碰觸定位鐵片)"
  ElseIf Err_Code = 3100 Then
    Err_Translate = "異常:昇降鐵片自我學習異常(至下一鐵片時間過久)"
  ElseIf Err_Code = 3101 Then
    Err_Translate = "異常:昇降鐵片自我學習異常(過早碰觸定位鐵片)"
  ElseIf Err_Code = 3102 Then
    Err_Translate = "警示:命令格式中FROM為禁用庫格"
  ElseIf Err_Code = 3103 Then
    Err_Translate = "警示:命令格式中TO為禁用庫格"
  ElseIf Err_Code = 3104 Then
    Err_Translate = "警示:只作TO命令,但無載"
  ElseIf Err_Code = 3105 Then
    Err_Translate = "異常:做FROM或TO時走行定位後目前鐵片與目標不合"
  ElseIf Err_Code = 3106 Then
    Err_Translate = "異常:做FROM時昇降定位後目前鐵片與目標不合"
  ElseIf Err_Code = 3107 Then
    Err_Translate = "異常:做FROM或TO時走行定位後目前ENCODER數值與目標不合"
  ElseIf Err_Code = 3108 Then
    Err_Translate = "異常:做FROM或TO時走行定位後目前ENCODER數值與目標不合"
  ElseIf Err_Code = 6000 Then
    Err_Translate = "警示:有殘留板在3截輸送機上,請處理"
  ElseIf Err_Code <> 0 Then
    Err_Translate = "異常說明未建檔"
  End If
End Function

Public Function InitialComPort(ByVal lComPort As Long, ByVal lBaudRate As Long, ByVal bParityCheck As Boolean, ByVal lDataBit As Long, ByVal lStopBit As Long) As Long
    Dim sComPortSetting As String, ComPort As Long
    
    On Error GoTo err1_rtn
        ComPort = lComPort
        'If Left(project_name, 2) = "SP" Then ComPort = lComPort + 11
        InitialComPort = -1
        
        ' Check Port No. is valid or not ? It should be between 1 & 8.
        If lComPort < 1 Or lComPort > 10 Then
            InitialComPort = 1
            GoTo err1_rtn
        End If
        
        ' Assemble COM Port Settings
        If bParityCheck = True Then
            sComPortSetting = CStr(lBaudRate) & ",Y," & CStr(lDataBit) & "," & CStr(lStopBit)
        ElseIf bParityCheck = False Then
            sComPortSetting = CStr(lBaudRate) & ",N," & CStr(lDataBit) & "," & CStr(lStopBit)
        End If
        
        ' Initial Port
        '''frmMain.mscLED(lComPort - 1).CommPort = ComPort
        '''frmMain.mscLED(lComPort - 1).Settings = sComPortSetting
        '''frmMain.mscLED(lComPort - 1).InputLen = 0
        '''frmMain.mscLED(lComPort - 1).PortOpen = True
        '''frmMain.mscLED(lComPort - 1).RThreshold = 1
        
        InitialComPort = 0
    Exit Function
err1_rtn:
    InitialComPort = Err.Number
    If InitialComPort = 8012 Then 'NPort沒接,err.Number=8012,此裝置未開啟
      Call syslog("", 0, 0, "ComPort=" & ComPort & ",err.Number=8012", "show")
    Else
      Call err2_rtn("InitialComPort")
    End If
End Function



' 說明 : 盈德LED資料傳送程式
Public Function SendMessage(ByVal sMessage As String, ByVal lComPort As Long, ByVal lID As Long, ByVal lCharNumType As Long) As Long
    Dim sMessageB As String
    Dim lMessageLength As Long
    Dim lPageCount As Long
    Dim bMessage() As Byte
    Dim bOutput() As Byte
    Dim bPageMessage() As Byte
    Dim lCount As Long
    Dim lTotalLength As Long
    Dim lCountType As Long
    
    On Error GoTo err1_rtn
        SendMessage = -1
        
        ' Check Port No. is valid or not ? It should be between 1 & 8.
        If lComPort < 1 Or lComPort > 10 Then
            SendMessage = 10
            GoTo err1_rtn
        End If
        ' Check Com Port is opened or not?
'        If frmMain.mscLED(lComPort - 1).PortOpen = False Then
'            SendMessage = 11
'            GoTo err1_rtn
'        End If
        
        ' Check LED ID No. is valid or not ? It should be between 1 & 127 OR 255.
        If (lID < 1 Or lID > 127) And lID <> 255 Then
            SendMessage = 20
            GoTo err1_rtn
        End If
        
        ' Check LED Display character number is valid or not ? It should be between 1 & 100.
        If (lCharNumType < 1 Or lCharNumType > 100) Then
            SendMessage = 25
            GoTo err1_rtn
        End If
        
        ' Check data length of Send Message
        If Len(sMessage) = 0 Then
            SendMessage = 30
            GoTo err1_rtn
        End If
        
        sMessageB = StrConv(sMessage, vbFromUnicode)
        lMessageLength = LenB(sMessageB)
        lPageCount = lMessageLength \ lCharNumType
        If (lMessageLength Mod lCharNumType) <> 0 Then
            lPageCount = lPageCount + 1
        End If
        If lPageCount < 1 Or lPageCount > 127 Then
            SendMessage = 31
            GoTo err1_rtn
        End If
        bMessage = StrConv(sMessage, vbFromUnicode)
        
        
        ReDim bOutput(14 + lPageCount * (2 + lCharNumType)) As Byte
        
        
        ' 中文12字資料傳送格式:
        ' ED ED ED 5C INT FF ID FF Sel Dis 00 Hbyte Lbyte Function TXT End
        '    1        2   3  4  5   6  7   8    9     10     11    12   13
        
        ' Leading Code --> ED ED ED 5C :  4 BYTES HEX. CODE.
        bOutput(0) = &HED
        bOutput(1) = &HED
        bOutput(2) = &HED
        bOutput(3) = &H5C
        
        ' INT : 中斷循環次數. (01H - 7FH). 不使用時設 01H .
        bOutput(4) = &H1
        
        ' FF :  1 BYTE HEX. CODE.
        bOutput(5) = &HFF
        
        ' Machine ID, 01H~7FH  -->  ID : 傳送資料給特定機號(ID.CODE = 01H - 7FH), 設為FFH時可傳送至每一台.
'        bOutput(6) = &H1
        bOutput(6) = CByte(lID)
        
        ' FF :  1 BYTE HEX. CODE.
        bOutput(7) = &HFF
        
        ' Display Mode, Interrupt Mode : FBH, Common Mode : 01H  -->  Sel : 中斷傳送時 Sel= FBH , 一般傳送時 Sel= 01H.
        bOutput(8) = &H1
        
        ' Display Screen Count, 01H~7FH  -->  Dis = 顯示內容之總"幕數", 即字幕機循環一次所出現之畫面總數.(01H - 7FH)
        bOutput(9) = CByte(lPageCount)
        
        ' 00 =  1 BYTE HEX. CODE
        bOutput(10) = &H0
        
        ' Hbyte = 所有 Function + TXT 之總長度的高位元組
        ' Lbyte = 所有 Function + TXT 之總長度的低位元組
        ' Total Message Length, High Byte + Low Byte
        lTotalLength = lPageCount * (2 + lCharNumType)
        
        ' Purpose : avoid overflow, 應該小於3302=127*(2+lCharNumType)
        If lTotalLength > 65536 Then
            MsgBox "訊息數目過多,造成資料無法顯示", vbOKOnly, "警告"
            GoTo err1_rtn
        End If
        bOutput(11) = CByte(lTotalLength \ 256)
        bOutput(12) = CByte(lTotalLength Mod 256)
        
        
'        Function = 依序為第一幕畫面顯示方式碼, 第一幕畫面停留時間碼, 第二幕畫面顯示
'                  方式碼, 第二幕畫面停留時間碼, 第三幕畫面顯示方式碼  ...........
'                  .至最後一幕畫面停留時間碼
'
'         (1) 畫面顯示方式碼 = 20H - 3FH (定義如下所示)
'
'             20 左移連接上幕出現     21 不作變化換幕出現     22 雪花方式換幕出現
'             23 由上往下覆蓋出現     24 由下往上覆蓋出現     25 由下往上旋轉出現
'             26 由上往下旋轉出現     27 由中間向上下展開     28 由中間向上下閉合
'             29 由右向左覆蓋出現     3A 由左向右覆蓋出現     2B 由右向左旋轉出現
'             2C 由左向右旋轉出現     2D 由中間向左右展開     2E 由中間向左右閉合
'             2F 波浪方式換幕出現     30 閃爍方式換幕出現     31 由下往上慢慢竄升
'             32 由左向右貼字出現     33 由右向左飛字出現     34 由左向右慢慢浮字
'             35 由右向左慢慢浮字     36 右移連接上幕出現     37 稍慢左移連接上幕
'             38 很慢左移連接上幕
'
'             ps.中斷傳送的畫面顯示方式碼固定為 20H, 若只有一幕畫面時,畫面會靜止不動.
'
'         (2) 畫面消失方式碼 = 畫面消失方式之設定值 + 停留時間之設定值
'
'             畫面消失方式碼定義如下所示: (中斷資料傳送的消失方式碼固定為 02H)
'             00 無(或左移用)  10 只作畫面清除  20 雪花方式消失  30 往下覆蓋消失
'             40 往上覆蓋消失  50 往下旋轉消失  60 往上旋轉消失  70 上下展開消失
'             80 上下閉合消失  90 往右覆蓋消失  A0 往左覆蓋消失  B0 往右旋轉消失
'             C0 往左旋轉消失  D0 顯示時間 4秒  E0 顯示時間 6秒  F0 顯示時間 8秒
'
'             "停留時間" 可設定 00 - 15 秒 , 依序代表 " 00H - 0FH "
'
'             ps. 畫面顯示方式碼 = 21H & 畫面顯示方式碼 = F0H & 只有一幕資料時,
'                 畫面只會顯示日期與時間, 並不做任何變化.
        ' Message Content
        For lCount = 0 To lPageCount - 1
'            bOutput(12 + 2 * lCount + 1) = &H2C
'            bOutput(12 + 2 * lCount + 2) = CByte((lCount * 16 + 34) Mod 256)
            bOutput(12 + 2 * lCount + 1) = &H20
            bOutput(12 + 2 * lCount + 2) = &H0
        Next
        
'        TXT = 依序為第一幕"畫面顯示內容", 第二幕"畫面顯示內容"........ 最後一幕
'             "畫面顯示內容".每一幕畫面顯示內容固定是 lCharNumType bytes文字資料 (文字資料
'             不足者補 20H),每一幕畫面顯示內容最後一個BYTE不可是BIG5碼的第一個碼.
        For lCount = 0 To lMessageLength - 1
            bOutput(12 + 2 * lPageCount + 1 + lCount) = bMessage(lCount)
        Next
        
        If (lMessageLength Mod lCharNumType) <> 0 Then
            lCount = lCharNumType - (lMessageLength Mod lCharNumType)
            ReDim bPageMessage(lCount) As Byte
            For lCountType = 0 To lCount - 1
                bPageMessage(lCountType) = &H20
                bOutput(12 + 2 * lPageCount + lMessageLength + 1 + lCountType) = &H20
            Next
        End If
        
        ' Ending Code  -->  END =  "00"  HEX. CODE
'        bOutput(13 + lPageCount * 42) = &H0
        bOutput(13 + lPageCount * (2 + lCharNumType)) = &H0
        
'        If mlInDerdLEDUsedPortNum >= 1 Then
''            If frmCYILEDDisplay.MSComm2.PortOpen = False Then
''                Trace "[InDerdLED!SendMessage] : Message > " & "MSComm2 Port was not opened!"
''                frmCYILEDDisplay.MSComm2.PortOpen = True
''            End If
'            frmCYILEDDisplay.MSComm2.Output = bOutput
'            Trace "[InDerdLED!SendMessage] : Message > " & "Message sent to #1InDerd LED is : " & sMessage
'        End If
'        If mlInDerdLEDUsedPortNum >= 2 Then
''            If frmCYILEDDisplay.MSComm3.PortOpen = False Then
''                Trace "[InDerdLED!SendMessage] : Message > " & "MSComm3 Port was not opened!"
''                frmCYILEDDisplay.MSComm3.PortOpen = True
''            End If
'            frmCYILEDDisplay.MSComm3.Output = bOutput
'            Trace "[InDerdLED!SendMessage] : Message > " & "Message sent to #2InDerd LED is : " & sMessage
'        End If
        '''frmMain.mscLED(lComPort - 1).Output = bOutput

        SendMessage = 0
    Exit Function
err1_rtn:
    Call err2_rtn("SendMessage")
End Function

'1. 系統損毀時,重新安裝系統的程序:
'   (1) for Server: 安裝 window2000 OS,電腦名稱為SERVER_
'       for Client: 安裝 window98 OS,電腦名稱為COM0_或TTY1___或TTY2___
'   (2) 安裝 VB5.0
'   (3) 產生目錄 c:\system\vb
'   (4) 找一台正常的PC,Copy此PC下c:\system\vb所有檔案到 c:\system\vb中
'   (5) for Server: 執行 c:\system\vb\Public_init.BAT SERVER
'       for Client: 執行 c:\system\vb\Public_init.BAT CLIENT
'   (6) StatusBar以Label999取代,所以COMCTL32.ocx不再使用
'       執行regedit以進入registry中,使用搜尋功能確認COMCTL32.ocx的版本
'       for Server: 版本必須為1.2
'       for Client: 版本必須為1.1
'   (7) 專案的名稱
'       for Server: SF000_server.vbp
'                   必須在Server上compile,compile後的執行檔叫作 "自動倉庫系統(Server版)"
'       for Client: SF000_client.vbp
'                   必須在Client上compile,compile後的執行檔叫作 "自動倉庫系統(Client版)"
'   (8) 在Server上安裝及規劃SQL SERVER資料庫程序
'       (A) 安裝SQL SERVER資料庫
'       (B) 執行Microsoft SQL Server -> Enterprise Manager -> Microsoft SQL Servers ->
'           SQL Server Group -> Server(for NT system) -> Databases ->
'           建一個名稱為sqldb空的資料庫
'       (C) 執行Microsoft SQL Server -> Query Analyzer ->
'           SQL Server: 選(local) , 選Use SQL Server authentication , Login Name: sa , Password: 空白
'           -> DB: 選sqldb
'           -> File -> Open -> 選SF_DB1.sql -> 按F5執行SF_DB1.sql,在sqldb空的資料庫中建立檔案結構
'       (D) 將之前備份的資料Restore回SQL SERVER資料庫中
'       (E) 建立ODBC
'           名稱(DSN): sqldb , 伺服器(Server): 選(local) , 變更預設資料庫到: sqldb
'       (F) 測試VB程式,是否可以連結SQL SERVER資料庫
'   (9) 在Client上建立ODBC
'       名稱 (DSN): sqldb , 伺服器(Server): 選server , 變更預設資料庫到: sqldb
'
'3. 設定使用元件:(.ocx,拉出使用)
'   regsvr32 comctl32.ocx 時就已註冊入register
'   (1) Apex True DBGrid Data Bound Grid
'   (2) Crystal Report Control 4.6(CRYSTL32.ocx)
'   (3) Microsoft Data Bound Grid control(DBGRID32.ocx)
'   (4) Microsoft Data Bound List controls 5.0(used by DBList,DBCombo)(DBLIST32.ocx)
'   (5) Microsoft Grid control(GRID32.ocx)
'   (6) Microsoft Masked Edit control 5.0(MSMASK32.ocx)
'   (7) Microsoft Remote Data control 2.0(MSRDC20.ocx)
'   (8) Microsoft Tabbed Dialog control 5.0(TABCTL32.ocx)
'   (9) Microsoft Windows Common controls 5.0(SP2)(COMCTL32.ocx 1.1版 or 1.3版)(used by SStatus)
'   (10) Pinnacle-BPS Graph Control
'   (11) Sheridan 3D controls(THREED32.ocx)
'   (12) Microsoft Calendar Control 9.0:瀏覽(找到MSCAL.ocx)-->開啟MSCAL.ocx,設定引用項目會自動出現
'   (13) FarPoint Spread 3.0(OLEDB):setup時必須給密碼
'   (14) QC-Capacity Control
'   (15) QC-Charter Quality Charting Tool
'   設定引用項目:(.oca or .dll,程式呼叫)
'   (1) Microsoft DAO3.5 Object Library(used by Database,Workspace,recordset)
'   (2) Microsoft Remote Data Object 2.0(used by rdoEnvironment,bcn,brs,Msrdo20.dll)
'   (3) Microsoft Calendar Control 9.0
'   (4) FarPoint Spread 3.0(OLEDB)
'   (5) QC-Capacity Control
'   (6) QC-Charter Quality Charting Tool
'
'4. 系統或程式注意事項:
'   (1) server使用c:\winnt\system32,client使用c:\windows\system32
'   (2) StatusBar以Label999取代,所以COMCTL32.ocx不再使用
'       COMCTL32.ocx版本在server為1.2,但在台北發展系統及client為1.1,因之在server上compiler後的.exe,不能拿到client上執行
'   (3) 呼叫Msgbox()後會造成brs time out,此時再呼叫任何brs.(eg. brs.EOF,brs.close)都會異常,Object不存在異常
'   (4) 呼叫Set brs = bcn2.OpenResultset後,要再呼叫Set brs = bcn2.OpenResultset或bcn2.execute,一定要先作一次bcn2.close,及Set bcn2 = rdoEnvironments(0).OpenConnection,否則會有[ODBC SQL Server Driver]連接正忙錄於另一個hstmt結果異常
'       呼叫Set brs = bcn2.OpenResultset後,再呼叫Set brs = rcn.OpenResultset或rcn.execute則不會有異常
'       連續呼叫rcn.execute則不會異常
'       因之要多open幾個bcn,或open後立即close,可以Set brs = bcn2.OpenResultset,rcn.execute,Set MSRDC1.Resultset = rcn1.OpenResultset各使用一組
'   (5) 呼叫Set brs = bcn2.OpenResultset後必須有bcn2.close動作,但可以不須作brs.close,作brs.close後很容易再去呼叫brs!,造成異常
'   (6) 呼叫Set MSRDC1.Resultset = rcn.OpenResultset前一定要呼叫Set MSRDC1.Resultset = Nothing,否則會造成下一次呼叫Set brs = bcn2.OpenResultset或Set MSRDC1.Resultset = bcn2.OpenResultset異常
'       Set MSRDC1.Resultset = Nothing
'       Set MSRDC1.Resultset = bcn2.OpenResultset("select * from pd", rdOpenKeyset, rdConcurRowver)
'
'   (7) S111中可以重覆呼叫下列程式碼但呼叫順序必須如下:
'       If Index = 0 Then
'         Set brs = bcn2.OpenResultset("Select count(*) from lo where lo_awno='" & awno & "' and lo_sgst='" & stat & "'", rdOpenKeyset, rdConcurRowver)
'         Text2 = brs(0)
'         bcn2.Close: Set bcn2 = rdoEnvironments(0).OpenConnection("sqldb", rdDriverNoPrompt, False)
          '此處若打bcn2.OpenResultset時,執行S360.Show時會出問題
'         (注意Set MSRDC1.Resultset = Nothing不可以放在Set MSRDC1.Resultset = bcn2.OpenResultset之後)
'         Set MSRDC1.Resultset = Nothing '必須加入此行否則第二次執行Set MSRDC1.Resultset = rcn.OpenResultset時會異常
'         Set MSRDC1.Resultset = bcn2.OpenResultset("select * from lo,lo1 where lo_awno='" & awno & "' and lo_sgst='" & stat & "' and lo_awno=lo1_awno and lo_lono=lo1_lono order by lo_lono", rdOpenKeyset, rdConcurRowver)
'       End If
'       If MSRDC1.Resultset.EOF Then MsgBox ("查無資料!")
        '錯誤 bcn2.Close: Set bcn2 = rdoEnvironments(0).OpenConnection("sqldb", rdDriverNoPrompt, False)
        '錯誤 Set MSRDC1.Resultset = Nothing '必須加入此行否則第二次執行Set MSRDC1.Resultset = rcn.OpenResultset時會異常
'       (注意MsgBox()不會造成Set MSRDC1.Resultset = Nothing異常)
'   (8) If MSRDC1.Resultset.EOF Then MsgBox ("查無資料!") '一定之前要有Set MSRDC1.Resultset = bcn2.OpenResultset,才可執行此行
'       但Set MSRDC1.Resultset = Nothing則可不須有
'   (9) 注意對於SQL Server而言,下面Set brs= 及Set MSRDC1 兩行指令巔倒放時會有異常,但對於Access則無影響
'       Set brs = bcn2.OpenResultset("Select count(*) from tr where tr_awno='" & awno & "' and tr_date>='" & beg1 & "' and tr_date<='" & end1 & "'", rdOpenKeyset, rdConcurRowver)
'       Set MSRDC1.Resultset = rcn.OpenResultset("select * from tr where tr_awno='" & awno & "' and tr_date>='" & beg1 & "' and tr_date<='" & end1 & "' order by tr_sitm", rdOpenKeyset, rdConcurRowver)
'       bcn2及rcn若設不同時就不會異常
'  (10) 假如first_prog.systemname = "宏和胚布自動倉庫系統",則If Left(first_prog.systemname, 4) = "宏和"會false,first_prog.systemname = "宏和",則If Left(first_prog.systemname, 4) = "宏和"會true
'  (11) 注意對於DBCombo不可以如此給值 ==> DBCombo1=lono_buf(此時DBCombo1並不是放lono_buf值),必須如下式:
'       MSRDC1.SQL = "select lo_lono from lo where lo_awno='" & awno & "' lo_lono='" & lono_buf & "'": MSRDC1.Refresh
'       (DBCombo1內容值為MSRDC1.SQL)
'  (12) DBCOMBO內容值只能來自MSRDC1.SQL,且只能用選的,當以鍵盤在DBCOMBO中改任何一字時,都會造成後續動作異常
'       另外將DBCOMBO內容當作Select之Where條件時會篩錯東西
'       因之不要用DBCOMBO
'  (13) 注意log檔稍大時(200行左右),當連線程式要寫一個訊息時,會讓連線程式變很慢(會idle近5秒),所以試車完時務必取消寫入log檔指令
'  (14) If Combo1.Text = "空庫格" Then stat = SG_EMPTY:If Combo1.Text = "在庫庫格" Then stat = SG_STOR
'       注意後半部不會執行
'  (15) 使用TDBGrid,TDBGrid1_RowColChange副程式必須作如下保護,且注意不要在TDBGrid1_RowColChange副程式外使用TDBGrid1.Columns(0).Text,可以使用Global變數帶出
'       If LastCol <> -1 Then
'         lono = TDBGrid1.Columns(0).Text
'       End If
'  (16) Client上.exe程式執行時若發生執行期錯誤-2147417848(80010108) Automation錯誤,此時試著在程式中隨便加一些syslog1,或隨便mark幾行,多compiler幾次就會過
'       因之Client上程式compile後請立即執行S111查詢功能鍵或S110庫存明細查詢,若有問題時Automation錯誤會立即產生,
'       此時請立即在S111中將Call syslog1("test", 3, 0, ""),Call syslog1("test", 4, 0, ""),Call syslog1("test", 5, 0, "")三行mark掉或恢復,之後再compile及執行就會正常
'       If Combo1.Text = "禁用庫格" Then
'         'Call syslog1("test", 3, 0, "")
'         Set brs = bcn2.OpenResultset("Select count(*) from lo where lo_awno='" & awno & "' and lo_fbst='X'", rdOpenKeyset, rdConcurRowver)
'         'Call syslog1("test", 4, 0, "")
'         Text2 = brs(0)
'         'Call syslog1("test", 5, 0, "")
'         Set MSRDC1.Resultset = rcn.OpenResultset("select * from lo where lo_awno='" & awno & "' and lo_fbst='X' order by lo_lono", rdOpenKeyset, rdConcurRowver)
'  (17) 透過網路作目錄copy時,必須確定沒有程式在執行,可以按Ctrl Alt Del看是否有Graphical server或SF000_Client存在
'  (19) 系統異常碼:429,ActiveX元件無法產生物件
'       MSRDO20.Dll版本修改日期必須為1997/1/16上午12:00,檔長368K(灌VB時WINNT\SYSTEM32中仍會是舊版本Msrdo20.dll修改日期為1997/7/19下午05:01,檔長368K
'       regsvr32 msrdo20.dll
'  (20) 資料庫改用ACCESS時注意事項:(不支援store procedure 及view及Create table)
'       (A) get_now(),setting()等store procedure必須作exit function處理
'       (B) View lo1,cw1無法運作
'  (21) SQL2000安裝時,其中會有一個螢幕問你驗證模式及sa之密碼設定
'       螢幕上會有2個Option驗證模式選項:
'       (A) WINDOW驗證模式(default)
'       (B) WINDOW及SQL混合式驗證模式
'       請改選(B),同時下方之sa密碼設定請勾選不設密碼
'       倘若不如此設時,之後將無法以sa去使用SQL
'  (22) COM01電腦要透過ODBC連到SERVER電腦時,針對ODBC中用戶端設定如下:
'       伺服器別名為:SERVER,網路程式庫選:具名管道(選TCPIP時,若沒有設定IP,會失敗),具名管道內容為:\\SERVER\pipe\sql\query
'       注意若資料在本機上時,則可以不管用戶端設定
'  (23) put_tr()中執行 values('" & awno1 & "'," & brs_cw!cw_palt & ",'" & lono & "','       ',0,0,'" & get_now() & "','" & get_now() & "',' ','          ','             ',' ','" & Left(get_now(), 10) & "','" & opno & "','" & srid & "','" & stno & "','" & name1 & "','0','" & brs_cw!cw_now & "->" & brs_cw!cw_nxt & "','" & get_srnm(srid) & "','        ','        ','        ','        ',0,'            ','       ','0','000000000000000000000000','000000000000000000000000','0000','    ',' ',' ','    ',' ')"
'       發生ODBC -SQL異常碼40002: 陳述式已經結束
'       上述異常描述不夠清楚,可以將下述指令加至err1_rtn中
'       Dim er As rdoError
'       Cran.Print Err, Error
'       For Each er In rdoErrors
'         Cran.Print er.Description, er.Number
'       Next er
'       執行後結果顯示"Primary Key違反條件約束'PK_tr_0D44F85C',無法在物件'tr'上插入重覆索引鍵"
'       注意按掉msgbox()後,程式會回到上述rcn.Execute "insert into tr指令處繼續執行
'  (24) (A) 在S998中2台天車分由Timercran(1)及Timercran(2)去控制時,並不會造成Timercran(1)及Timercran(2)正常的分時效應,而是如下方所示,其中一個Timer會佔住並連續執行很多次,呼叫Doevents時並不會將控制權立即交給對方
'           下面的測試情況為Timercran(2)連線正常,但Timercran(1)連線沒得到對方回應,作call wait(3)秒
'           Timercran(2)執行 --> 2sec後 --> Timercran(2)執行 --> 2sec後 --> Timercran(2)執行 --> 2sec後 --> Timercran(2)執行 --> 2sec後 --> Timercran(2)執行 --> 2sec後 -->
'           Timercran(1)執行 --> 5sec後 --> Timercran(1)執行 --> 5sec後 --> Timercran(1)執行 --> 5sec後 -->
'           再輪回Timercran(2),如此反覆執行,上述情況似乎是以接籠的方式在執行
'       (B) 上述Timercran(1) call wait(3)改為call wait_100ms(30)時,Timercran(2)會變成連續執行6-7次(原先為連續執行5次),效果更佳
'           注意絕不能使用call wait(),否則會造成純接籠的效果,一個Timercran作完後,另一個Timercran才接著作,call wait_100ms較不會(較不明顯),
'           觀察原因在於doevents次數不能太頻繁,call wait()中doevents非常頻繁(10ms一次),call wait_100ms()中doevents(100ms一次)
'           Timercran(ii).Interval = 400也不能設太小,至少必須400,否則Plc會搶不到
'       (C) 程式中call wait_100ms()的時間縮短時,整個運作的效能會改善很多,改為Doevents更好
'       (D) 將2台天車改由S998及S998_1兩個.frm分別去執行,效果同(A)
'       (E) 因之目前程式改採接籠的方式執行,call wait_100ms()的時間儘量縮短,連續發生的異常訊息不能寫入log檔中
'       (F) 再加入S999時情況類似上方
'  (25) insert COMBO into table時,由於COMBO沒有Maxlength的屬性,因之很容易會造成insert失敗,因之必須加上left(COMBO,)去作管制
'  (26) 發生ODBC SQL SERVER Drive 逾時異常:處理方式為進入用戶端設定:將TCP/IP改為具名管道(pipe)
'  (27) (A) for S210 : Dim mm as long
'           mm = (9960 - 240) * 12會產生溢位異常 , 但mm=9960-240,mm=mm*12時則正常
'       (B) 異常碼6:溢位,dim palt as integer,palt=dt_valu,但dt_valu>32767造成
'       (C) SG362中原先jj = (i * 100)/j,其中i=329
'           但因為jj=329*100時不管jj是什麼型態,都會發生溢位,jj=32900則正常,因此必須改寫成Dim jj as single jj=(i/j)*100
'  (28) NetBEUI要設,才能看到外面的電腦名稱,NetBEUI若沒設,IP仍能ping的到
'  (28) Client端ODBC異常:伺服器不支援封包大小變更,已使用預設值,發生原因為SQL Server connections數超過設定值
'       異常發生處Set rcn1 = rdoEnvironments(0).OpenConnection("sqldb", rdDriverNoPrompt, False)
'       TPNPCENG0232-->Properties-->Conections-->Maximum concurrent user connections設為0表示unlimited
'  (29) VB程式特殊語法:
'       (A) Set myrs = rcn.OpenResultset("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD HH24miss') FROM dual", rdOpenStatic)
'           date1 = Format(Now - 60, "YYYYMMDDhhmmss")
'       (B) MSRDC1.SQL="select * from cw":MSRDC1.Refresh
'           MSRDC1.Caption = Space(26) & MSRDC1.Resultset.ROWCOUNT & "  筆資料"
'           brs.Requery
'           注意MSRDC搭配DBGRID時純作顯示用,不要使用例如MSRDC1.Resultset.Movenext或MSRDC1.Resultset!sg_lono
'           eg. SB315
'       (C) update dt set dt_palt=dt_palt+1 where
'           update sg set sg_pqty=sg_quty where
'       (D) 系統錯誤行號ERL:APP.Path
'       (E) create view ff_33 as select distinct fi_cono,fi_ptno,fi_loct,sum(fi_quty) fi_quty,fi_name from ff where fi_opno='2' group by fi_cono,fi_ptno,fi_loct,fi_name;
'           eg. 用法請參考SB321PE
'           2. 合併單號出庫作業
'           (1) 螢幕左半部MRP檔ff(以合併單號cono1查詢,配料位置fi_loct='Z')
'               fi_cono   fi_ptno    fi_loct    fi_quty(預出量)    fi_opno    fi_name  fi_date
'                cono1     ptno1       Z          100                '2'        ' '
'                cono1     ptno1       Z          300                '2'        ' '
'               螢幕右半部選取待出庫資料檔ff3(ptno1之loct='Z'的ptno1併成一筆,fi_name變成'2')
'               fi_cono   fi_ptno    fi_loct    fi_quty(預出量)    fi_opno    fi_name  fi_date
'                cono1     ptno1       Z          400                '2'        '2'
'       (F) rcn.Execute "insert into sgbb select * from sg where sg_awno = '" & awno & "' and sg_lono = '" & sglono & "' and sg_ptno  <> '" & ptno_want_out & "'"
'           sgbb及sg的結構必須完全相同,sg. SB321PE,SB314EB1
'       (G) 整個系統rcn只需在SB000P內連一次即可,揚生系統
'           set brs=rcn.:set brs=nothing:set MSRDC1.Resultset=nothing
'           但public_empty.frm卻會在下列行If get_string_value1(awno, brs!sg_lono, "lo_sgst") = "F" And get_string_value1(awno, brs!sg_lono, "lo_fbst") = "0" Then
'           發生項次(4)[ODBC SQL Server Driver]連接正忙錄於另一個hstmt結果異常
'           Set brs = rcn.OpenResultset("select * from sg where sg_awno='" & awno & "' and sg_ptno like 'PALLET%' and sg_quty=" & Val(pallet), rdOpenKeyset, rdConcurRowver)
'           lono1 = "      "
'           Do
'             If brs.EOF Then Exit Do
'             If get_string_value1(awno, brs!sg_lono, "lo_sgst") = "F" And get_string_value1(awno, brs!sg_lono, "lo_fbst") = "0" Then
'               lono1 = brs!sg_lono: Exit Do
'             End If
'             brs.MoveNext
'           Loop
'           Set brs = Nothing
'           解決對策:
'             (1) 只要改成If get_string_value(awno, brs!sg_lono, "lo_sgst") = "F" And get_string_value(awno, brs!sg_lono, "lo_fbst") = "0" Then就正常
'             (2) get_string_value1不變,SG001中另產生rcn1,將Public.bas及SG900.bas中的rcn改為rcn1
'           以下的例子在get_string_value1中Set brs2 = rcn1.OpenResultset時發生連接正忙錄於另一個hstmt結果異常
'           Public Sub pt_insert()
'           Dim brs2 As New ADODB.Recordset, zone As String * 1
'           Set brs2 = rcn1.OpenResultset("Select sg_lono,sg_ptno from sg", rdOpenKeyset, rdConcurRowver)
'           Do
'             If brs2.EOF Then Exit Do
'             Set brs = rCn.OpenResultset("Select pt_zonename from pt where pt_ptno='" & brs2!sg_ptno & "'", rdOpenKeyset, rdConcurRowver)
'             If brs.EOF Then
'               zone = get_string_value1(awno, brs2!sg_lono, "lo_zone")
'       (H) MyPos = Instr(StartChar, SearchString, SearchChar, 1),eg. SB332
'       (I) Sb332.frm中List使用,可以取代DBGRID,eg. SB332
'           For i = 0 To List1.ListCount - 1
'             List1.ListIndex = i:tmp = Trim(List1.Text):ipos = InStr(2, tmp, "<")
'       (J) 最複雜case,eg. SB314EC
'       (K) Crystal report當查詢用,同一時間產生很多Crystal report收在螢幕下,之後user再去看,eg. SB321PE
'       (L) 只有SB000P always為放大狀態,其它Form不必是always放大狀態,如此user同一時間可以操作很多畫面
'       (M) DBGRID中按Control作多選處理方式:eg. SB321PE
'           '選2筆時DBGrid1.SelBookmarks.Count=2
'           Dim jj as string
'           For ii = 0 To DBGrid1.SelBookmarks.Count - 1
'             jj = DBGrid1.SelBookmarks.Item(ii): MSRDC1.Resultset.Bookmark = CLng(Asc(jj))
'             If Trim(MSRDC1.Resultset!fi_loct) < Trim(loct_start) Or Trim(MSRDC1.Resultset!fi_loct) > Trim(loct_end) Then GoTo rtn1
'       (N) 倘若DBGRID中只是作單選時則使用MSRDC1.resultset!fi_odno即可,eg. SB315
'           注意如Public240工作檔維護作業,螢幕中有refresh的情況下,一進入SSCommand副程式中,必須立即將MSRDC1.resultset!cw_palt保護起來
'       (O) Set myrs2 = rcn.OpenResultset("SELECT sum(sg_pqty) FROM sgbb WHERE SG_LONO='9'", rdOpenStatic)
'           倘若不存在合乎條件的資料,注意此時(1) brs2為非EOF (2) brs2(0)=NULL (3) ii=brs2(0)會當掉,eg. SB315
'           plan_quty = 0: calculate_plan_quty = 0
'           If Not myrs2.EOF Then
'             If Not IsNull(myrs2(0)) Then plan_quty = myrs2(0): calculate_plan_quty = myrs2(0)
'       (P) 不可只以1條指令刪除大量歷史資料,會咬住不作,必須如下例方式分次刪,eg. ff_mt1
'           date1 = Format(Now - 4, "YYYYMMDDhhmmss")
'           For ii = 15 To 30
'             date1 = Format(Now - ii, "YYYYMMDDhhmmss"): DoEvents
'             For jj = 0 To 24
'               rcn.Execute "delete tr where tr_sitm like '" & Left(date1, 8) & Right("00" & jj, 2) & "%'": DoEvents
'             Next jj
'           Next ii
'           rcn.Execute "delete tr where tr_sitm <= '" & Left(date1, 8) & "'"
'           刪除大量歷史資料時系統Performance會變很慢,所以可以使用pc所提供的排定工作軟體在凌晨作
'       (Q) 在同一台PC中,同一個MOXA port只能initial一次
'       (R) 在同一台PC中包含連線及螢幕兩部份時,最好分成兩個Projects,如IBS系統,硬布自動包裝系統
'           (1) 硬布自動包裝系統: 連線程式每1秒與自動包裝區Picker連線,因之螢幕程式SH610若要與自動包裝區Picker連線時,必須透過pa_labelflag要求作連線
'       (S) 針對雙層STV,請將雙層站號設在最後,eg. I99/I45,如此才不致影響其它PLC處理作業(會造成其它PLC站號中斷)
'       (U) RS6000, s__.log檔可達50M, log目錄可達1.5G
'       (V) DBGrid1.Columns(ii + 5).Visible = False: DBGrid1.Columns(ii + 12).caption = "False"
'       (W) VB Builder: user: CHIA-TAI CHEN ,password: BLD1-fuow-8113-gpyu-5741
'       (X) 英文小字請選擇Arial Narrow字型8
'       (Y) call wait最小單位為0.1 sec,eg. call wait(0.1)
'       (Z) Dim LL as long,lib_buf as string*100 注意lib_buf不可設為string,LL會回-7
'           LL = sio_read(port, lib_buf, 100)
'       (B) len("(車上)")=4,對VB而言,中文字算1byte
'       (C) CrystalReport1.Action = 1發生Unable to connect:incorrect log on parameters
'           (1) 以SQL作測試發現在CrystalReport設計環境中設計CrystalReport時若選擇資料庫時選取不設密碼的資料庫user則列印時沒問題,若選取設密碼的資料庫user則於VB啟動CrystalReport列印時會有問題(注意在CrystalReport設計環境中列印沒有問題)
'               SQL資料庫user可以不設密碼,但ORACLE資料庫user不可以不設密碼
'           (2) 若資料庫user要設密碼且可列印的條件為CrystalReport1.SelectionFormula = ""且DiscardSavedData=False
'           (3) 在CrystalReport設計環境中列印沒有問題
'           (4) 在SQL中設計的報表,無法在ORACLE使用,反之亦同,可以直接在Crystal report中Remove from report,再Add Database to report,然後insert Database field
'           (5) Visual Linking Expert設定pt(pt_ptno)-->sg(sg_ptno)的連結
'       (D) SJ系統,由於archive log mode on,Server上每天11:45必須自動執行\oracle\oradata\nanya\archive\del-archive.bat,自動清除%t_%s.dbf archive log資料
'       (E) 執行時發生ActiveX元件無法產生物件或執行到系統保留字String,er,Ucase等發生編譯錯誤:找不到專案或程式庫,解決方法為(1).在設定引用項目中將遺漏項移除(2).Public_init.bat重新執行一次
'       (F) loct = Mid(buf, 3, 3) & "   "
'           If brs5!cr_loct <> loct Then rCn.Execute "update cr set cr_loct='" & loct & "' where cr_mveq='" & mveq & "'"
'           buf內是亂碼,update時發生引號字串未以恰當方式終止異常
'       (G) 資料欄的插入值過大或The statement has been terminated異常,insert資料長度比欄位定義的長度大時發生
'           一般常發生在內含中文字欄位,內含中文字欄位最好加trim,eg: lono="(車上)",name1="測試人員"
'       (H) 設ODBC時,要連結的電腦最好是指定IP(不要給電腦名稱)
'       (I) 必須作電腦登入動作後才可以(1)設定net use i:(2)使用網路的資源,使用別人分享的資料
'       (J) CrystalReport1.SortFields(0) = "+{sg_lono}"
'       (K) -8 = sio_open(3),表port已被別人使用
'           -1 = sio_ioctl(3, B9600, BIT_8 Or STOP_1 Or P_NONE)
'       (L) 針對S510.rpt,S350.rpt,PUBLIC620.rpt,產品編號pt_ptno必須存在時,報表上才會看到資料
'       (M) PLC連線程式在讀狀態時不要把虛擬站號納入通訊協定中,SI Project犯了此錯誤
'       (N) 網路風暴會造成ODBC逾時或Name Pipe斷裂異常
'       (O) Epoxy2系統:(1)開機時顯示WORKGROUP-網路上有重覆的名稱,(2)點選靠近我的電腦時發生無法存取WORKGROUP,指定的檔案是唯讀屬性,(3)以IP要尋找WORKGROUP的其他電腦卻找不到(但卻Ping得到)
'                      解決方法:改為WORKGROUP-SI即可
'       (P) Epoxy2系統:Server IP改掉後Run自動倉庫系統.exe,於Set rCn = rdoEnvironments(0).OpenConnection("sqldb", rdDriverNoPrompt, False),發生Connect異常
'                      但(1)在VB開發環境卻Run正常(2)test ODBC,SQL Plus也都正常
'                      解決方法:恢復原IP
'       (Q) FileCopy "c:\system\si\si000.vbp", "\\bl\nanya.txt\si000.vbp"
'           Kill "c:\system\si\si000.vbp"
'       (R) Mkdir "c:\nanya" 目錄若已存在時會當掉
'       (S) 引號字串未以恰當方式終止,發生原因之一,如下'" & awno & "''多打一個'
'           rCn.Execute "insert into sg values('" & awno & "'','" & lono & "','" & ptno & "'," & Val(quty) & ",0,'" & get_now() & "','0','" & odno & "','" & lot1 & "','S','" & Left(wrap, 1) & "'," & (Val(quty) + Val(quty2) + Val(quty3)) & ",'" & lot1 & "'," & Val(quty2) & ",'" & lot2 & "'," & Val(quty3) & ",'" & lot3 & "')"
'       (T) S001中Timer1_Timer顯示通訊通道上出現EOF異常
'       (U) SJ360中出現msgbox,此時在執行階段SJ001中讀固定式條碼Timer2不會受影響,但在Debug階段SJ001中讀固定式條碼Timer2會被抑制
'       (V) 執行期錯誤,通訊通道上出現EOF或未與ORACLE相連
'           (1) 程式一開始即作Set rcn_bar = rdoEnvironments(0).OpenConnection("sqldb1", rdDriverNoPrompt, False)
'               由於系統在名皓一廠但要連接的ERP database在二廠,頻寬的因素造成Connection被關閉
'               因之執行右式時會異常Set brs2 = rcn_bar.OpenResultset("select * from prup where num_car = '" & Mid(lib_buf, 2, 8) & "'", rdOpenStatic)
'               因之若要連接遠端資料庫時,最好是要存取時才作連結,存取後立即關閉
'           (2) 名皓二廠出貨台PC,當馬達轉動時,要切換畫面call menu_enable_rtn,或作任何資料庫讀取動作都會咬注,然後異常跳出
'               經測試後是馬達干擾網路線,換一條由小電控室直拉的網路線時情況稍有改善,但仍會稍微咬住,表示整個電腦機櫃附近都有受到干擾
'       (W) 另外若發生驅動程式的 SQLSetConnectAttr失敗,表示執行Set rcn_bar = rdoEnvironments(0).OpenConnection時異常
'           針對ORA8i,TNSLSNR.EXE(listener.ora),ORACLE.EXE必須執行,OracleORAHomeTNSListener,OracleServiceNANYA(instance)必須為啟動自動狀態
'           否則VB rdoEnvironments(0).OpenConnection時會顯示SQLSetConnectAttr失敗
'       (W) SF:
'           SQL SERVER在作較大動作時,會產生ODBC逾時異常,此時必須先rcn.close再open一次
'           COM01會呼叫err2_rtn產生ODBC逾時異常,2003/07/27 09:52:54 [CRAN1-rd_data_cran_fx] (999,0) CRAN1          -rd_data_cran_fx:系統異常,行號為:5,異常碼: 40002 S1T00: [Microsoft][ODBC SQL Server Driver]逾時
'                                               2003/07/27 09:52:54 [PLC1-get_integer_value1] (999,0) PLC1           -get_integer_value1:系統異常,行號為:1,異常碼: 40002 S1T00: [Microsoft][ODBC SQL Server Driv
'
'           COM02會呼叫err2_rtn產生ODBC逾時異常,2003/07/27 09:52:26 [CRAN2-st_comd_cran_fx] (999,0) CRAN2          -st_comd_cran_fx:系統異常,行號為:13,異常碼: 40002 S1T00: [Microsoft][ODBC SQL Server Driver]逾
'                                               2003/07/27 09:52:26 [MAPATH-get_string_value1] (999,0) MAPATH         -get_string_value1:系統異常,行號為:4,異常碼: 40002 S1T00: [Microsoft][ODBC SQL Server Driv
'                                               2003/07/27 09:52:26 [PLC2-get_string_value1] (999,0) PLC2           -get_string_value1:系統異常,行號為:4,異常碼: 40002 S1T00: [Microsoft][ODBC SQL Server Driver'     (30) Mapath處理原則
'        () '87'=緊急停止,'90'=馬達過載,'92'=前後安全bar碰到異常,'96'=與輸送機交訊過久異常,'97'=變頻器異常,'98'=前後走行極限異常
'        () SJ993發生ORA-01013:使用者要求取消目前作業,原因為刪檔資料太多
'     (30) Mapath處理原則
'       由新的cw_now決定新的搬運設備cw_mveq
'       mveq = "          "
'       If brs_mapath!cw_nxt = brs_mapath!cw_to Then
'         mveq = EQ_NULL
'       End If
'       If mveq = "          " Then
'         逐站處理雙向道以上的站,會決定出兩個以上的搬運設備
'         If Mid(now1, 1, 3) = "A21" Then
'           If is_lono(brs_mapath!cw_awno, brs_mapath!cw_to) Then mveq = EQ_CRAN1 Else mveq = "A21->A20線"
'         End If
'       endif
'       If mveq = "          " Then '決定輸送機線的搬運設備
'         Set brs2 = rcn.OpenResultset("select * from cr where cr_awno='" & brs_mapath!cw_awno & "' and cr_mveq like '" & Mid(now1, 1, 3) & "%'", rdOpenStatic)
'         If not brs2.EOF Then mveq = brs2!cr_mveq
'         Set brs2 = Nothing
'       endif
'       If mveq = "          " Then
'         決定STV的搬運設備
'       endif
'       If mveq = "          " Then '決定天車搬運設備
'         If brs_mapath!cw_awno = "1" Then mveq = EQ_CRAN1
'         If brs_mapath!cw_awno = "2" Then mveq = EQ_CRAN2
'       End If
'       由新的搬運設備mveq,逐項設備決定新的下一站cw_nxt
'       nxt1 = "      "
'       If mveq = EQ_NULL Then
'         Call put_tr(brs_mapath!cw_awno, brs_mapath!cw_srid, brs_mapath!cw_palt, "      ")
'         Call syslog("mapath", 2, 0, "Complete:palt=" & brs_mapath!cw_palt & ",from=" & brs_mapath!cw_from & ",to=" & brs_mapath!cw_to)
'         Set brs_mapath = Nothing: Exit Sub
'       ElseIf Left(mveq, 1) = "#" Then 'mveq=天車-->A05-->STV(或輸送機)-->A03(to站)
'         必須考慮(車上)-->010203,(車上)-->000203,(車上)-->A05的情況
'         If is_lono(brs_mapath!cw_awno, brs_mapath!cw_to) or Left(brs_mapath!cw_to, 2) = "00" Then
'           nxt1 = brs_mapath!cw_to
'         Else
'           If Left(brs_mapath!cw_to, STNO_NUM) = "A03" Or Left(brs_mapath!cw_to, STNO_NUM) = "A05" Then nxt1 = "A05   "
'         End If
'       ElseIf Left(mveq, 3) = "STV" Then
'       Else 'mveq=輸送機線
'         nxt1 = Mid(mveq, 6, STNO_NUM) & "   "
'       End If
'     (31) QC
'        iRowCnt = brs.RowCount
'        Do Until brs.EOF
'            sQCDate = Trim(brs!sp_saveTm) '_880620173020
'            dTval = brs!sp_meaPV
'            iNewCount = .QCAddValue(dTval)
'            Count = Count + 1
'            brs.MoveNext
'        Loop
'        iRolCnt = .QCDataCount
'        brs.Close
'        sMesMean = .Mean
'        .Redraw = True
'        .Refresh
'     (32) sjcom.vbp作VB-Builder編行號,sjcom1.vbp作Compiler
'     (34) VB\Microsoft Visual Basic 5.0\應用程式安裝精靈
'     (35) 1. 右鍵-->內容-->設定值-->800*600像素,高彩(16位元)-->按確定鍵
'          2. 安裝ORACLE9i Client-->選擇Administrator Type-->選No,I will create net service names myself.The Assistant will help me create one now-->選Oracle8i or later database or service-->Service Name輸入nanya-->選TCP-->Hostname輸入Server電腦之IP,選Use the standard port number of 1521-->選Yes,perform a test-->按Next鍵-->選No,按Next鍵-->按Next鍵-->按Finish鍵-->按Exit鍵
'          3. 設定ODBC:
'             使用ORACLE8i Server時,ODBC請不要搭配Microsoft ODBC Driver for Oracle或Microsoft ODBC for Oracle,某些地方會有問題(例MSRDC)
'             ODBC可以搭配ORACLE9i Client(ODBC為Oracle in OraHome90)(麥寮電腦安裝ORACLE8i Client無法執行,ODBC為Oracle ODBC Driver),但注意必須於install時選擇Administrator Type(選Runtime Type時,在ODBC中Oracle in OraHome901不會出現,若不設ODBC時可以選此)
'             開始-->控制台-->系統管理工具-->資料來源(ODBC)-->系統資源來源名稱(在上方)-->按新增鍵-->選Oracle in OraHome90(在下方)-->按完成鍵-->Data Source Name輸入sqldb,TNS Service Name輸入nanya,User ID輸入nanya-->按TestConnection鍵-->Password輸入nanya-->按OK鍵-->會看到Connection successful-->按確認鍵-->按OK鍵-->按確定鍵
'          4. 尋找Server電腦,將Server電腦上\system\install目錄copy至本機\system\install目錄,執行\system\install\setup-->按確定鍵-->按更改目錄鍵-->請輸入或選擇目的目錄,路徑輸入c:\system\si-->按是鍵-->按鈕-->按確定鍵
'     (36) 由UNIX ftp至PC時,注意UNIX端檔案名稱是否有重覆(例同時存在SF997.FRM,SF997.frm)
'     (37) 10.3.76.211,Administrator/nanyaeng,Service name:nanyaeng scott/tiger,tools/crystal6.0
'     (38) 測試用10.3.77.80,10.3.77.81
'     (39) SQL backup job
'          1. LH backup_complete : Occurs every 1 day(s),at 下午02:00:00
'          2. LH backup_log : Occurs every 1 day(s),at 下午02:00:00
'          3. LH backup_differential : Occurs every 1 hour(s) between 上午01:00 and 下午11:59:59
'     (40) 椰城資料收集器需在TTY1A18電腦set up SmartSync,TTY1A18電腦一開機,SmartSync就會自動執行,且會抓到並鎖住COM01-COM06的狀態,
'          注意此時其它連線程式會open port失敗,sio_open=-8,必須close SmartSync後再執行連線程式,資料收集器使用COM02
'          資料收集器主畫面(1)資料收集 (2)資料查詢 (3)資料刪除 , 進入後再分(1)入庫 (2)出庫
'          連線傳輸時資料收集器in.dat or out.dat --> TTY1A18 SmartSync接收 --> 轉出in.txt or out.txt
'(33)
'[MACH_NAME]
'MACHNAME = RL07
'[ODBC]
'UserName = SLCCL
'Password = SLCCL
'ODBCDSN = SLCCL
'[GROUP_ID]
'GROUPID1 = 8113
'GROUPID2 = 8123
'GROUPID3 = 8133
'[OPERATION]
'mode = OFFLINE
'MODE=ONLINE
'[Plc]
'OMRON
'TYPE=O
'MITSUBISHI
'TYPE=M

'sUserName = IniRead("ODBC", "USERNAME", App.Path & "\IniFile\ccl2.ini")

'Public Function IniRead(lpAppName As String, lpKeyName As String, lpFileName As String) As String
'   Dim Size As Long '字串總長度
'   Dim valid As Long '傳回有值的長度
'   Dim resoult As String '傳回字串
'   'Dim lpFileName As String '路徑
'   Dim lpDefault As String '如果沒有 Key的話，傳回來的初始值
'   Dim lpReturnString As String '傳回的 Buffer
'   lpDefault = "KeyNull"
'   lpReturnString = Space$(1024)
'   Size = Len(lpReturnString)
'   If StrComp(Trim$(lpKeyName), "NULL", 1) = 0 Then
'     valid = GetPrivateProfileString(lpAppName, vbNullString, lpDefault, lpReturnString, Size, lpFileName)
'   Else
'     valid = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, lpReturnString, Size, lpFileName)
'   End If
'   resoult = Left$(lpReturnString, valid)
'   If resoult = "KeyNull" Then
'      'Err.Raise 3001, , lpFileName & " 檔找不到 [" & lpAppName & "] " & lpKeyName & "= 的描述"
'       xMsgBox lpFileName & " Cannot Find Description Of [" & lpAppName & "] " & lpKeyName & " = ", vbOKOnly, "Err message"
'       IniRead = ""
'       End
'   Else
'      IniRead = resoult
'   End If
'End Function




