Imports System
Imports System.Text

Public Class D49F4022
	Dim report As D99C2003
    Dim dt1 As DataTable
    Dim dtPeriod As DataTable

    Dim strFind1 As String = ""
    Dim strFind2 As String = ""
    Dim strFind3 As String = ""
    Dim strFind4 As String = ""
    Dim strFind5 As String = ""

    Dim sField1 As String = ""
    Dim sField2 As String = ""
    Dim sField3 As String = ""
    Dim sField4 As String = ""
    Dim sField5 As String = ""

    Private _moduleID As String = "49"
    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            _moduleID = Value
        End Set
    End Property


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub SetBackColorObligatory()
        tdbcDivisionID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcReportCode.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcPeriodFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcPeriodTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""

        'Load tdbcDivisionID
        LoadCboDivisionIDReport(tdbcDivisionID, D49, , gbUnicode)

        'Load TablePeriod
        dtPeriod = LoadTablePeriodReport(D49)

        LoadCboPeriodReport(tdbcPeriodFrom, tdbcPeriodTo, dtPeriod, gsDivisionID)

        sSQL = "SELECT 	ReportID, ReportCode, ReportName" & UnicodeJoin(gbUnicode) & " as ReportName, "
        sSQL &= " IsGeneral, IsCustomized, Negatives, DecimalOriginal, MethodID,"
        sSQL &= " Selection01, Selection02, Selection03, Selection04, Selection05,"
        sSQL &= " Sel01IDFrom, Sel02IDFrom, Sel03IDFrom, Sel04IDFrom, Sel05IDFrom,"
        sSQL &= " Sel01IDTo, Sel02IDTo, Sel03IDTo, Sel04IDTo, Sel05IDTo,"
        sSQL &= " EditSel01, EditSel02, EditSel03, EditSel04, EditSel05"
        sSQL &= " FROM D49T4020 WITH(NOLOCK) "
        sSQL &= " WHERE Disabled = 0"

        dt1 = ReturnDataTable(sSQL)
        LoadDataSource(tdbcReportCode, dt1, gbUnicode)

        EnableTDBCSel(tdbcSel01IDFrom, tdbcSel01IDTo, "-1", "Sel01IDFrom", "Sel01IDTo", "EditSel01")
        EnableTDBCSel(tdbcSel02IDFrom, tdbcSel02IDTo, "-1", "Sel02IDFrom", "Sel02IDTo", "EditSel02")
        EnableTDBCSel(tdbcSel03IDFrom, tdbcSel03IDTo, "-1", "Sel03IDFrom", "Sel03IDTo", "EditSel03")
        EnableTDBCSel(tdbcSel04IDFrom, tdbcSel04IDTo, "-1", "Sel04IDFrom", "Sel04IDTo", "EditSel04")
        EnableTDBCSel(tdbcSel05IDFrom, tdbcSel05IDTo, "-1", "Sel05IDFrom", "Sel05IDTo", "EditSel05")
    End Sub

    Private Sub D49F4022_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.NumPad1, Keys.D1
                    tdbcDivisionID.Focus()
                Case Keys.NumPad2, Keys.D2
                    tdbcReportCode.Focus()
                Case Keys.NumPad3, Keys.D3
                    tdbcSel01IDFrom.Focus()
                Case Keys.NumPad4, Keys.D4
                    chkIsTime.Focus()
            End Select
        End If
    End Sub

    Private Sub D49F4022_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        Me.Cursor = Cursors.WaitCursor
        InputbyUnicode(Me, gbUnicode)
        SetBackColorObligatory()
        Loadlanguage()

        LoadTDBCombo()
        LoadDefault()

        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        If _ModuleID = "27" Then
            Me.Text = rL3("Bao_cao_phan_tich_tuoi_no_-_D49F4022").Replace("D49F4022", "D27F4012") & UnicodeCaption(gbUnicode)
            Me.Name = "D27F4012"
        Else
            Me.Text = rL3("Bao_cao_phan_tich_tuoi_no_-_D49F4022") & UnicodeCaption(gbUnicode) 'BÀo cÀo ph¡n tÛch tuåi ní - D49F4022
        End If
        '================================================================ 
        lblNotes.Text = rl3("Ghi_chu") 'Ghi chú
        lblReportCode.Text = rl3("Bao_cao") 'Báo cáo
        lblteAsOfDate.Text = rl3("Ngay_tinh_no") 'Ngày tính nợ
        lblSelID01From.Text = rL3("Tieu_thuc") & " 1" 'Tieâu thöùc 1
        lblSelID02From.Text = rL3("Tieu_thuc") & " 2" 'Tieâu thöùc 2
        lblSelID03From.Text = rL3("Tieu_thuc") & " 3" 'Tieâu thöùc 3
        lblSelID04From.Text = rL3("Tieu_thuc") & " 4" 'Tieâu thöùc 4
        lblSelID05From.Text = rL3("Tieu_thuc") & " 5" 'Tieâu thöùc 5
        lblPeriod.Text = rl3("Ky") 'Kỳ
        lblDivisionID.Text = rl3("Don_vi") 'Đơn vị
        '================================================================ 
        btnPrint.Text = rl3("_In") '&In
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        chkIsTime.Text = rl3("Thoi_gian_phat_sinh") 'Thời gian phát sinh
        '================================================================ 
        GroupBox1.Text = "1. " & rl3("Don_vi") '1. Đơn vị
        GroupBox2.Text = "2. " & rl3("Bao_cao") '2. Báo cáo
        GroupBox3.Text = "3. " & rl3("Tieu_thuc_loc") '3. Tiêu thức lọc
        GroupBox4.Text = "4. " & rl3("Thoi_gian") '4. Thời gian
        '================================================================ 
        tdbcSel05IDTo.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel05IDTo.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel05IDFrom.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel05IDFrom.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel04IDTo.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel04IDTo.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel04IDFrom.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel04IDFrom.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel03IDTo.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel03IDTo.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel03IDFrom.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel03IDFrom.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel02IDTo.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel02IDTo.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel02IDFrom.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel02IDFrom.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel01IDFrom.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel01IDFrom.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcSel01IDTo.Columns("Code").Caption = rl3("Ma") 'Mã
        tdbcSel01IDTo.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcReportCode.Columns("ReportCode").Caption = rl3("Ma") 'Mã
        tdbcReportCode.Columns("ReportName").Caption = rl3("Ten") 'Tên
        tdbcDivisionID.Columns("DivisionID").Caption = rl3("Ma") 'Mã
        tdbcDivisionID.Columns("DivisionName").Caption = rl3("Ten") 'Tên
    End Sub

    Private Sub LoadDefault()
        'Load mặc định
        c1dateAsOfDate.Value = Date.Today

        tdbcDivisionID.SelectedValue = gsDivisionID
        tdbcPeriodFrom.Text = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        tdbcPeriodTo.Text = giTranMonth.ToString("00") & "/" & giTranYear.ToString

        chkIsTime_CheckedChanged(Nothing, Nothing)
    End Sub

    Private Sub EnableTDBCSel(ByVal tdbcSelFrom As C1.Win.C1List.C1Combo, ByVal tdbcSelTo As C1.Win.C1List.C1Combo, ByVal sSelection As String, ByVal sSelIDFrom As String, ByVal sSelIDTo As String, ByVal sEditSel As String)
        If sSelection = "-1" Then
            tdbcSelFrom.Enabled = True
            tdbcSelTo.Enabled = True

            LoadTDBCSel(tdbcSelFrom, tdbcSelTo, "-1")

            tdbcSelFrom.Text = ""
            tdbcSelTo.Text = ""
        Else
            If tdbcReportCode.Columns(sEditSel).Text = "1" Then
                tdbcSelFrom.Enabled = False
                tdbcSelTo.Enabled = False
                tdbcSelFrom.Text = tdbcReportCode.Columns(sSelIDFrom).Text
                tdbcSelTo.Text = tdbcReportCode.Columns(sSelIDTo).Text
            Else
                tdbcSelFrom.Enabled = True
                tdbcSelTo.Enabled = True

                LoadTDBCSel(tdbcSelFrom, tdbcSelTo, tdbcReportCode.Columns(sSelection).Text)

                tdbcSelFrom.Text = tdbcReportCode.Columns(sSelIDFrom).Text
                tdbcSelTo.Text = tdbcReportCode.Columns(sSelIDTo).Text
            End If
        End If
    End Sub

#Region "Events tdbcDivisionID with txtDivisionName"

    Private Sub tdbcDivisionID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDivisionID.Close
        If tdbcDivisionID.SelectedValue IsNot Nothing Then
            tdbcReportCode.Focus()
        End If
    End Sub

    Private Sub tdbcDivisionID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcDivisionID.GotFocus
        'Dùng phím Enter
        tdbcDivisionID.Tag = tdbcDivisionID.Text
    End Sub

    Private Sub tdbcDivisionID_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcDivisionID.MouseDown
        'Di chuyển chuột
        tdbcDivisionID.Tag = tdbcDivisionID.Text
    End Sub

    Private Sub tdbcDivisionID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcDivisionID.SelectedValueChanged

        If tdbcDivisionID.SelectedValue Is Nothing Then
            txtDivisionName.Text = ""
        Else
            txtDivisionName.Text = tdbcDivisionID.Columns(1).Value.ToString

        End If
    End Sub

    Private Sub tdbcDivisionID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcDivisionID.LostFocus
        If tdbcDivisionID.Tag.ToString = "" And tdbcDivisionID.Text = "" Then Exit Sub
        If tdbcDivisionID.Tag.ToString = tdbcDivisionID.Text And tdbcDivisionID.SelectedValue IsNot Nothing Then Exit Sub

        If tdbcDivisionID.FindStringExact(tdbcDivisionID.Text) = -1 OrElse tdbcDivisionID.SelectedValue Is Nothing Then
            tdbcDivisionID.Text = "%"
            LoadCboPeriodReport(tdbcPeriodFrom, tdbcPeriodTo, dtPeriod, "%")
            Exit Sub
        End If

        LoadCboPeriodReport(tdbcPeriodFrom, tdbcPeriodTo, dtPeriod, tdbcDivisionID.Text)

    End Sub
#End Region

#Region "Events tdbcReportCode with txtReportName1"

    Private Sub tdbcReportCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcReportCode.GotFocus
        'Dùng phím Enter
        tdbcReportCode.Tag = tdbcReportCode.Text
    End Sub

    Private Sub tdbcReportCode_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcReportCode.MouseDown
        'Di chuyển chuột
        tdbcReportCode.Tag = tdbcReportCode.Text
    End Sub

    Private Sub tdbcReportCode_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcReportCode.Close
        If tdbcReportCode.SelectedValue IsNot Nothing Then
            txtNotes.Focus()
        End If
    End Sub

    Private Sub tdbcReportCode_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcReportCode.SelectedValueChanged
        If tdbcReportCode.SelectedValue Is Nothing Then
            txtReportName.Text = ""
        Else
            txtReportName.Text = tdbcReportCode.Columns("ReportName").Value.ToString
        End If
    End Sub

    Private Sub tdbcReportCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcReportCode.LostFocus
        If tdbcReportCode.Tag.ToString = "" And tdbcReportCode.Text = "" Then Exit Sub
        If tdbcReportCode.Tag.ToString = tdbcReportCode.Text And tdbcReportCode.SelectedValue IsNot Nothing Then Exit Sub

        If tdbcReportCode.FindStringExact(tdbcReportCode.Text) = -1 Then
            tdbcReportCode.Text = ""
            txtReportName.Text = ""

            EnableTDBCSel(tdbcSel01IDFrom, tdbcSel01IDTo, "-1", "Sel01IDFrom", "Sel01IDTo", "EditSel01")
            EnableTDBCSel(tdbcSel02IDFrom, tdbcSel02IDTo, "-1", "Sel02IDFrom", "Sel02IDTo", "EditSel02")
            EnableTDBCSel(tdbcSel03IDFrom, tdbcSel03IDTo, "-1", "Sel03IDFrom", "Sel03IDTo", "EditSel03")
            EnableTDBCSel(tdbcSel04IDFrom, tdbcSel04IDTo, "-1", "Sel04IDFrom", "Sel04IDTo", "EditSel04")
            EnableTDBCSel(tdbcSel05IDFrom, tdbcSel05IDTo, "-1", "Sel05IDFrom", "Sel05IDTo", "EditSel05")

            lblSelID01From.Text = rL3("Tieu_thuc") & " 1" 'Tieâu thöùc 1
            lblSelID01From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
            lblSelID02From.Text = rL3("Tieu_thuc") & " 2" 'Tieâu thöùc 2
            lblSelID02From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
            lblSelID03From.Text = rL3("Tieu_thuc") & " 3" 'Tieâu thöùc 3
            lblSelID03From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
            lblSelID04From.Text = rL3("Tieu_thuc") & " 4" 'Tieâu thöùc 4
            lblSelID04From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
            lblSelID05From.Text = rL3("Tieu_thuc") & " 5" 'Tieâu thöùc 5
            lblSelID05From.Font = FontUnicode(gbUnicode, FontStyle.Underline)

            strFind1 = ""
            strFind2 = ""
            strFind3 = ""
            strFind4 = ""
            strFind5 = ""

            Exit Sub
        End If

        'Hoàng Long Update 17/06/2009
        EnableTDBCSel(tdbcSel01IDFrom, tdbcSel01IDTo, "Selection01", "Sel01IDFrom", "Sel01IDTo", "EditSel01")
        EnableTDBCSel(tdbcSel02IDFrom, tdbcSel02IDTo, "Selection02", "Sel02IDFrom", "Sel02IDTo", "EditSel02")
        EnableTDBCSel(tdbcSel03IDFrom, tdbcSel03IDTo, "Selection03", "Sel03IDFrom", "Sel03IDTo", "EditSel03")
        EnableTDBCSel(tdbcSel04IDFrom, tdbcSel04IDTo, "Selection04", "Sel04IDFrom", "Sel04IDTo", "EditSel04")
        EnableTDBCSel(tdbcSel05IDFrom, tdbcSel05IDTo, "Selection05", "Sel05IDFrom", "Sel05IDTo", "EditSel05")

        If tdbcReportCode.Columns("Selection01").Text = "" Then
            lblSelID01From.Text = rL3("Tieu_thuc") & " 1" 'Tieâu thöùc 1
            lblSelID01From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
        Else
            LoadSelCaption(tdbcReportCode.Columns("Selection01").Text, lblSelID01From, sField1)
        End If

        If tdbcReportCode.Columns("Selection02").Text = "" Then
            lblSelID02From.Text = rL3("Tieu_thuc") & " 2" 'Tieâu thöùc 1
            lblSelID02From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
        Else
            LoadSelCaption(tdbcReportCode.Columns("Selection02").Text, lblSelID02From, sField2)
        End If

        If tdbcReportCode.Columns("Selection03").Text = "" Then
            lblSelID03From.Text = rL3("Tieu_thuc") & " 3" 'Tieâu thöùc 1
            lblSelID03From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
        Else
            LoadSelCaption(tdbcReportCode.Columns("Selection03").Text, lblSelID03From, sField3)
        End If

        If tdbcReportCode.Columns("Selection04").Text = "" Then
            lblSelID04From.Text = rL3("Tieu_thuc") & " 4" 'Tieâu thöùc 1
            lblSelID04From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
        Else
            LoadSelCaption(tdbcReportCode.Columns("Selection04").Text, lblSelID04From, sField4)
        End If

        If tdbcReportCode.Columns("Selection05").Text = "" Then
            lblSelID05From.Text = rL3("Tieu_thuc") & " 5" 'Tieâu thöùc 1
            lblSelID05From.Font = FontUnicode(gbUnicode, FontStyle.Underline)
        Else
            LoadSelCaption(tdbcReportCode.Columns("Selection05").Text, lblSelID05From, sField5)
        End If

        strFind1 = ""
        strFind2 = ""
        strFind3 = ""
        strFind4 = ""
        strFind5 = ""
    End Sub

    Private Sub LoadSelCaption(ByVal sCode As String, ByVal lblSel As Label, ByRef sField As String)
        Dim sSQL As String = "Select Field, Description" & UnicodeJoin(gbUnicode) & " as Description "
        sSQL &= " From D49V4000 where Code = " & SQLString(sCode) & " And Language = " & SQLString(gsLanguage)

        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            lblSel.Text = dt.Rows(0).Item("Description").ToString
            lblSel.Font = FontUnicode(gbUnicode, FontStyle.Underline)
            sField = dt.Rows(0).Item("Field").ToString
        End If
    End Sub

    Private Sub LoadTDBCSel(ByVal tdbcSelIDFrom As C1.Win.C1List.C1Combo, ByVal tdbcSelIDTo As C1.Win.C1List.C1Combo, ByVal sCode As String)
        Dim sSQL As String = ""

        sSQL = "SELECT 	Code, Description" & UnicodeJoin(gbUnicode) & " as Description, 1 as DisplayOrder "
        sSQL &= " FROM D49V4001"
        sSQL &= " WHERE 	Type = " & SQLString(sCode) & vbCrLf
        sSQL &= " UNION ALL " & vbCrLf
        sSQL &= " SELECT " & AllCode & " AS Code, " & AllName & " AS Description ,0 as DisplayOrder "
        sSQL &= " ORDER BY DisplayOrder, Code "

        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        LoadDataSource(tdbcSelIDFrom, dt.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcSelIDTo, dt.DefaultView.ToTable, gbUnicode)
    End Sub
#End Region

    Private Function SQLStoreD49P4022() As String
        Dim sSQL As String = ""
        sSQL &= ("-- In bao cao" & vbCrLf)
        sSQL &= "Exec D49P4022 "
        sSQL &= SQLString(tdbcDivisionID.Text) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(tdbcReportCode.Text) & COMMA 'ReportCode, varchar[50], NOT NULL
        sSQL &= SQLDateSave(c1dateAsOfDate.Value) & COMMA 'ASOfDate, datetime, NOT NULL
        sSQL &= SQLString(sField1) & COMMA 'Sel01Type, varchar[50], NOT NULL
        If strFind1 = "" Then
            sSQL &= SQLString(tdbcSel01IDFrom.Text) & COMMA 'Sel01IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString(tdbcSel01IDTo.Text) & COMMA 'Sel01IDTo, varchar[20], NOT NULL
        Else
            sSQL &= SQLString("") & COMMA 'Sel01IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString("") & COMMA 'Sel01IDTo, varchar[20], NOT NULL
        End If

        sSQL &= SQLString(sField2) & COMMA 'Sel02Type, varchar[50], NOT NULL        
        If strFind2 = "" Then
            sSQL &= SQLString(tdbcSel02IDFrom.Text) & COMMA 'Sel02IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString(tdbcSel02IDTo.Text) & COMMA 'Sel02IDTo, varchar[20], NOT NULL
        Else
            sSQL &= SQLString("") & COMMA 'Sel02IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString("") & COMMA 'Sel02IDTo, varchar[20], NOT NULL
        End If

        sSQL &= SQLString(sField3) & COMMA 'Sel03Type, varchar[50], NOT NULL        
        If strFind3 = "" Then
            sSQL &= SQLString(tdbcSel03IDFrom.Text) & COMMA 'Sel03IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString(tdbcSel03IDTo.Text) & COMMA 'Sel03IDTo, varchar[20], NOT NULL
        Else
            sSQL &= SQLString("") & COMMA 'Sel03IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString("") & COMMA 'Sel03IDTo, varchar[20], NOT NULL
        End If

        sSQL &= SQLString(sField4) & COMMA 'Sel04Type, varchar[50], NOT NULL        
        If strFind4 = "" Then
            sSQL &= SQLString(tdbcSel04IDFrom.Text) & COMMA 'Sel04IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString(tdbcSel04IDTo.Text) & COMMA 'Sel04IDTo, varchar[20], NOT NULL
        Else
            sSQL &= SQLString("") & COMMA 'Sel04IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString("") & COMMA 'Sel04IDTo, varchar[20], NOT NULL
        End If

        sSQL &= SQLString(sField5) & COMMA 'Sel05Type, varchar[50], NOT NULL        
        If strFind5 = "" Then
            sSQL &= SQLString(tdbcSel05IDFrom.Text) & COMMA 'Sel05IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString(tdbcSel05IDTo.Text) & COMMA 'Sel05IDTo, varchar[20], NOT NULL
        Else
            sSQL &= SQLString("") & COMMA 'Sel05IDFrom, varchar[20], NOT NULL
            sSQL &= SQLString("") & COMMA 'Sel05IDTo, varchar[20], NOT NULL
        End If
        sSQL &= SQLNumber(chkIsTime.Checked) & COMMA 'IsTime, int, NOT NULL
        sSQL &= SQLNumber(tdbcPeriodFrom.Columns("TranMonth").Text) & COMMA 'FromMonth, int, NOT NULL
        sSQL &= SQLNumber(tdbcPeriodFrom.Columns("TranYear").Text) & COMMA 'FromYear, int, NOT NULL
        sSQL &= SQLNumber(tdbcPeriodTo.Columns("TranMonth").Text) & COMMA 'ToMonth, int, NOT NULL
        sSQL &= SQLNumber(tdbcPeriodTo.Columns("TranYear").Text) & COMMA 'ToYear, int, NOT NULL
        sSQL &= SQLString(strFind1) & COMMA 'strFind1, varchar[8000], NOT NULL
        sSQL &= SQLString(strFind2) & COMMA 'strFind2, varchar[8000], NOT NULL
        sSQL &= SQLString(strFind3) & COMMA 'strFind3, varchar[8000], NOT NULL
        sSQL &= SQLString(strFind4) & COMMA 'strFind4, varchar[8000], NOT NULL
        sSQL &= SQLString(strFind5) & COMMA 'strFind5, varchar[8000], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function

    Private Function AllowPrint() As Boolean
        If tdbcDivisionID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Don_vi"))
            tdbcDivisionID.Focus()
            Return False
        End If
        If tdbcReportCode.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Bao_cao"))
            tdbcReportCode.Focus()
            Return False
        End If

        If Not CheckValidPeriodFromTo(tdbcPeriodFrom, tdbcPeriodTo) Then
            Return False
        End If
      
        Return True
    End Function

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If Not AllowNewD99C2003(report, Me) Then Exit Sub
        If Not AllowPrint() Then Exit Sub
        btnPrint.Enabled = False
        ' GetSelection()

        Me.Cursor = Cursors.WaitCursor

        'Dim report As New D99C1003
        'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
        '************************************
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = tdbcReportCode.Columns("ReportID").Text
        Dim sSubReportName As String = "D91R0000"
        Dim sReportCaption As String = ""
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        sReportCaption = rL3("Bao_cao_phan_tich_tuoi_nof") & " - " & sReportName
        If tdbcReportCode.Columns("IsCustomized").Text = "0" Then
            '            If D49Options.ReportLanguage = 0 Then
            '                sPathReport = Application.StartupPath & "\XReports\" & sReportName & ".rpt"
            '            ElseIf D49Options.ReportLanguage = 1 Then
            '                sPathReport = Application.StartupPath & "\XReports\VE-XReports\" & sReportName & ".rpt"
            '            Else
            '                sPathReport = Application.StartupPath & "\XReports\E-XReports\" & sReportName & ".rpt"
            '            End If
            sPathReport = UnicodeGetReportPath(gbUnicode, D49Options.ReportLanguage, "") & sReportName & ".rpt"
        Else
            '  sPathReport = Application.StartupPath & "\XCustom\" & sReportName & ".rpt"
            sPathReport = UnicodeGetReportPath(gbUnicode, D49Options.ReportLanguage, ReturnValueC1Combo(tdbcReportCode, "ReportID")) & sReportName & ".rpt"
        End If

        sSQL = SQLStoreD49P4022()

        sSQLSub = "Select * from D91V0016 Where DivisionID  = " & SQLString(tdbcDivisionID.Text)
        UnicodeSubReport(sSubReportName, sSQLSub, tdbcDivisionID.Text, gbUnicode)

        With report
            .OpenConnection(conn)
            .AddParameter("Title", txtReportName.Text)
            .AddParameter("Notes", txtNotes.Text)
            Dim dtParameter As DataTable = ReturnDataTable("Select ThousandSeparator, DecimalSeparator, D90_ConvertedDecimals From D91T0025 WITH(NOLOCK) ")
            If dtParameter.Rows.Count > 0 Then
                .AddParameter("DecimalSeparator", IIf(IsNothing(dtParameter.Rows(0).Item("DecimalSeparator").ToString) Or Trim(dtParameter.Rows(0).Item("DecimalSeparator").ToString) = "", ",", dtParameter.Rows(0).Item("DecimalSeparator").ToString))
                .AddParameter("ThousandSeparator", IIf(IsNothing(dtParameter.Rows(0).Item("ThousandSeparator").ToString) Or Trim(dtParameter.Rows(0).Item("ThousandSeparator").ToString) = "", ".", dtParameter.Rows(0).Item("ThousandSeparator").ToString))
                .AddParameter("DecimalConverted", IIf(IsNothing(dtParameter.Rows(0).Item("D90_ConvertedDecimals").ToString) Or Trim(dtParameter.Rows(0).Item("D90_ConvertedDecimals").ToString) = "", 0, dtParameter.Rows(0).Item("D90_ConvertedDecimals").ToString))
            Else
                .AddParameter("DecimalSeparator", ",")
                .AddParameter("ThousandSeparator", ".")
                .AddParameter("DecimalConverted", 0)
            End If
            .AddParameter("Negatives", tdbcReportCode.Columns("Negatives").Text)
            .AddParameter("DecimalOriginal", tdbcReportCode.Columns("DecimalOriginal").Text)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(sSQL)
            .PrintReport(sPathReport, sReportCaption)
        End With
        strFind1 = ""
        strFind2 = ""
        strFind3 = ""
        strFind4 = ""
        strFind5 = ""
        Me.Cursor = Cursors.Default
        btnPrint.Enabled = True
    End Sub

#Region "Events tdbcPeriod1"

    Private Sub tdbcPeriod1_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcPeriodFrom.LostFocus
        If tdbcPeriodFrom.FindStringExact(tdbcPeriodFrom.Text) = -1 Then tdbcPeriodFrom.Text = ""
    End Sub

#End Region

#Region "Events tdbcPeriod2"

    Private Sub tdbcPeriod2_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcPeriodTo.LostFocus
        If tdbcPeriodTo.FindStringExact(tdbcPeriodTo.Text) = -1 Then tdbcPeriodTo.Text = ""
    End Sub

#End Region

    Private Sub CallKeyDownF2(ByVal sSelection As String, ByRef sResult As String)
        sResult = ""
        'Dim sKeyID As String = ""
        'Dim f As New D91F6020
        'With f
        '    .ModeSelect = "1"
        '    .SQLSelection = "Select Code As SelectionID, Description" & UnicodeJoin(gbUnicode) & " As SelectionName From D49V4001 Where Type = " & SQLString(sSelection) & " Order By Code"
        '    .ShowDialog()
        '    sKeyID = .OutPut01
        '    .Dispose()
        'End With
        'If sKeyID <> "" Then
        '    sResult = sKeyID
        'End If

        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "SQLSelection", "Select Code As SelectionID, Description" & UnicodeJoin(gbUnicode) & " As SelectionName From D49V4001 Where Type = " & SQLString(sSelection) & " Order By Code")
        '  SetProperties(arrPro, "FormIDPermission", formPer)
        SetProperties(arrPro, "ModeSelect", L3Byte("1"))

        Dim frm As Form = CallFormShowDialog("D91D0240", "D91F6020", arrPro)
        sResult = GetProperties(frm, "ReturnField").ToString
    End Sub

#Region "Events tdbcSel01IDFrom"
    Private Sub tdbcSel01IDFrom_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel01IDFrom.KeyDown

        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection01").Text, strFind1)

            If strFind1 <> "" Then
                If strFind1.Substring(0, 1) <> "(" Then
                    If strFind1.IndexOf(";") = -1 Then
                        tdbcSel01IDFrom.Text = strFind1
                        tdbcSel01IDTo.Text = strFind1
                    Else
                        tdbcSel01IDFrom.Text = strFind1.Substring(0, strFind1.IndexOf(";"))
                        tdbcSel01IDTo.Text = strFind1.Substring(strFind1.IndexOf(";") + 1)
                    End If

                    strFind1 = ""
                Else
                    tdbcSel01IDFrom.Text = "%"
                    tdbcSel01IDTo.Text = "%"
                End If
            End If

        End If
    End Sub

    Private Sub tdbcSel01IDFrom_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel01IDFrom.LostFocus
        If tdbcReportCode.Columns("Selection01").Text <> "DD" And tdbcReportCode.Columns("Selection01").Text <> "ND" And tdbcReportCode.Columns("Selection01").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel01IDFrom.FindStringExact(tdbcSel01IDFrom.Text) = -1 Then tdbcSel01IDFrom.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcSel02IDFrom"
    Private Sub tdbcSel02IDFrom_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel02IDFrom.KeyDown

        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection02").Text, strFind2)

            If strFind2 <> "" Then
                If strFind2.Substring(0, 1) <> "(" Then
                    If strFind2.IndexOf(";") = -1 Then
                        tdbcSel02IDFrom.Text = strFind2
                        tdbcSel02IDTo.Text = strFind2
                    Else
                        tdbcSel02IDFrom.Text = strFind2.Substring(0, strFind2.IndexOf(";"))
                        tdbcSel02IDTo.Text = strFind2.Substring(strFind2.IndexOf(";") + 1)
                    End If

                    strFind2 = ""
                Else
                    tdbcSel02IDFrom.Text = "%"
                    tdbcSel02IDTo.Text = "%"
                End If
            End If

        End If
    End Sub

    Private Sub tdbcSel02IDFrom_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel02IDFrom.LostFocus
        If tdbcReportCode.Columns("Selection02").Text <> "DD" And tdbcReportCode.Columns("Selection02").Text <> "ND" And tdbcReportCode.Columns("Selection02").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel02IDFrom.FindStringExact(tdbcSel02IDFrom.Text) = -1 Then tdbcSel02IDFrom.Text = ""
        End If
    End Sub
#End Region

#Region "Events tdbcSel03IDFrom"

    Private Sub tdbcSel03IDFrom_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel03IDFrom.KeyDown

        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection03").Text, strFind3)

            If strFind3 <> "" Then
                If strFind3.Substring(0, 1) <> "(" Then
                    If strFind3.IndexOf(";") = -1 Then
                        tdbcSel03IDFrom.Text = strFind3
                        tdbcSel03IDTo.Text = strFind3
                    Else
                        tdbcSel03IDFrom.Text = strFind3.Substring(0, strFind3.IndexOf(";"))
                        tdbcSel03IDTo.Text = strFind3.Substring(strFind3.IndexOf(";") + 1)
                    End If

                    strFind3 = ""
                Else
                    tdbcSel03IDFrom.Text = "%"
                    tdbcSel03IDTo.Text = "%"
                End If
            End If

        End If
    End Sub

    Private Sub tdbcSel03IDFrom_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel03IDFrom.LostFocus
        If tdbcReportCode.Columns("Selection03").Text <> "DD" And tdbcReportCode.Columns("Selection03").Text <> "ND" And tdbcReportCode.Columns("Selection03").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel03IDFrom.FindStringExact(tdbcSel03IDFrom.Text) = -1 Then tdbcSel03IDFrom.Text = ""
        End If
    End Sub
#End Region

#Region "Events tdbcSel04IDFrom"
    Private Sub tdbcSel04IDFrom_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel04IDFrom.KeyDown

        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection04").Text, strFind4)

            If strFind4 <> "" Then
                If strFind4.Substring(0, 1) <> "(" Then
                    If strFind4.IndexOf(";") = -1 Then
                        tdbcSel04IDFrom.Text = strFind4
                        tdbcSel04IDTo.Text = strFind4
                    Else
                        tdbcSel04IDFrom.Text = strFind4.Substring(0, strFind4.IndexOf(";"))
                        tdbcSel04IDTo.Text = strFind4.Substring(strFind4.IndexOf(";") + 1)
                    End If

                    strFind4 = ""
                Else
                    tdbcSel04IDFrom.Text = "%"
                    tdbcSel04IDTo.Text = "%"
                End If
            End If

        End If
    End Sub

    Private Sub tdbcSel04IDFrom_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel04IDFrom.LostFocus
        If tdbcReportCode.Columns("Selection04").Text <> "DD" And tdbcReportCode.Columns("Selection04").Text <> "ND" And tdbcReportCode.Columns("Selection04").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel04IDFrom.FindStringExact(tdbcSel04IDFrom.Text) = -1 Then tdbcSel04IDFrom.Text = ""
        End If
    End Sub
#End Region

#Region "Events tdbcSel05IDFrom"

    Private Sub tdbcSel05IDFrom_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel05IDFrom.KeyDown

        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection05").Text, strFind5)

            If strFind5 <> "" Then
                If strFind5.Substring(0, 1) <> "(" Then
                    If strFind5.IndexOf(";") = -1 Then
                        tdbcSel05IDFrom.Text = strFind5
                        tdbcSel05IDTo.Text = strFind5
                    Else
                        tdbcSel05IDFrom.Text = strFind5.Substring(0, strFind5.IndexOf(";"))
                        tdbcSel05IDTo.Text = strFind5.Substring(strFind5.IndexOf(";") + 1)
                    End If

                    strFind5 = ""
                Else
                    tdbcSel05IDFrom.Text = "%"
                    tdbcSel05IDTo.Text = "%"
                End If
            End If

        End If
    End Sub

    Private Sub tdbcSel05IDFrom_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel05IDFrom.LostFocus
        If tdbcReportCode.Columns("Selection05").Text <> "DD" And tdbcReportCode.Columns("Selection05").Text <> "ND" And tdbcReportCode.Columns("Selection05").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợ đồng
            If tdbcSel05IDFrom.FindStringExact(tdbcSel05IDFrom.Text) = -1 Then tdbcSel05IDFrom.Text = ""
        End If
    End Sub
#End Region

#Region "Events tdbcSel01IDTo"

    Private Sub tdbcSel01IDTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel01IDTo.KeyDown
        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection01").Text, strFind1)

            If strFind1 <> "" Then
                If strFind1.Substring(0, 1) <> "(" Then
                    If strFind1.IndexOf(";") = -1 Then
                        tdbcSel01IDFrom.Text = strFind1
                        tdbcSel01IDTo.Text = strFind1
                    Else
                        tdbcSel01IDFrom.Text = strFind1.Substring(0, strFind1.IndexOf(";"))
                        tdbcSel01IDTo.Text = strFind1.Substring(strFind1.IndexOf(";") + 1)
                    End If

                    strFind1 = ""
                Else
                    tdbcSel01IDFrom.Text = "%"
                    tdbcSel01IDTo.Text = "%"
                End If
            End If

        End If
    End Sub

    Private Sub tdbcSel01IDTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel01IDTo.LostFocus
        If tdbcReportCode.Columns("Selection01").Text <> "DD" And tdbcReportCode.Columns("Selection01").Text <> "ND" And tdbcReportCode.Columns("Selection01").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel01IDTo.FindStringExact(tdbcSel01IDTo.Text) = -1 Then tdbcSel01IDTo.Text = ""
        End If
    End Sub
#End Region

#Region "Events tdbcSel02IDTo"

    Private Sub tdbcSel02IDTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel02IDTo.KeyDown
        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection02").Text, strFind2)

            If strFind2 <> "" Then
                If strFind2.Substring(0, 1) <> "(" Then
                    If strFind2.IndexOf(";") = -1 Then
                        tdbcSel02IDFrom.Text = strFind2
                        tdbcSel02IDTo.Text = strFind2
                    Else
                        tdbcSel02IDFrom.Text = strFind2.Substring(0, strFind2.IndexOf(";"))
                        tdbcSel02IDTo.Text = strFind2.Substring(strFind2.IndexOf(";") + 1)
                    End If

                    strFind2 = ""
                Else
                    tdbcSel02IDFrom.Text = "%"
                    tdbcSel02IDTo.Text = "%"
                End If
            End If

        End If
    End Sub

    Private Sub tdbcSel02IDTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel02IDTo.LostFocus
        If tdbcReportCode.Columns("Selection02").Text <> "DD" And tdbcReportCode.Columns("Selection02").Text <> "ND" And tdbcReportCode.Columns("Selection02").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel02IDTo.FindStringExact(tdbcSel02IDTo.Text) = -1 Then tdbcSel02IDTo.Text = ""
        End If
    End Sub
#End Region

#Region "Events tdbcSel03IDTo"

    Private Sub tdbcSel03IDTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel03IDTo.KeyDown
        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection03").Text, strFind3)

            If strFind3 <> "" Then
                If strFind3.Substring(0, 1) <> "(" Then
                    If strFind3.IndexOf(";") = -1 Then
                        tdbcSel03IDFrom.Text = strFind3
                        tdbcSel03IDTo.Text = strFind3
                    Else
                        tdbcSel03IDFrom.Text = strFind3.Substring(0, strFind3.IndexOf(";"))
                        tdbcSel03IDTo.Text = strFind3.Substring(strFind3.IndexOf(";") + 1)
                    End If

                    strFind3 = ""
                Else
                    tdbcSel03IDFrom.Text = "%"
                    tdbcSel03IDTo.Text = "%"
                End If
            End If

        End If
    End Sub

    Private Sub tdbcSel03IDTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel03IDTo.LostFocus
        If tdbcReportCode.Columns("Selection03").Text <> "DD" And tdbcReportCode.Columns("Selection03").Text <> "ND" And tdbcReportCode.Columns("Selection03").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel03IDTo.FindStringExact(tdbcSel03IDTo.Text) = -1 Then tdbcSel03IDTo.Text = ""
        End If
    End Sub
#End Region

#Region "Events tdbcSel04IDTo"

    Private Sub tdbcSel04IDTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel04IDTo.KeyDown
        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection04").Text, strFind4)

            If strFind4 <> "" Then
                If strFind4.Substring(0, 1) <> "(" Then
                    If strFind4.IndexOf(";") = -1 Then
                        tdbcSel04IDFrom.Text = strFind4
                        tdbcSel04IDTo.Text = strFind4
                    Else
                        tdbcSel04IDFrom.Text = strFind4.Substring(0, strFind4.IndexOf(";"))
                        tdbcSel04IDTo.Text = strFind4.Substring(strFind4.IndexOf(";") + 1)
                    End If

                    strFind4 = ""
                Else
                    tdbcSel04IDFrom.Text = "%"
                    tdbcSel04IDTo.Text = "%"
                End If
            End If

        End If
    End Sub


    Private Sub tdbcSel04IDTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel04IDTo.LostFocus
        If tdbcReportCode.Columns("Selection04").Text <> "DD" And tdbcReportCode.Columns("Selection04").Text <> "ND" And tdbcReportCode.Columns("Selection04").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel04IDTo.FindStringExact(tdbcSel04IDTo.Text) = -1 Then tdbcSel04IDTo.Text = ""
        End If
    End Sub
#End Region

#Region "Events tdbcSel05IDTo"

    Private Sub tdbcSel05IDTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcSel05IDTo.KeyDown
        If e.KeyCode = Keys.F2 Then
            If tdbcReportCode.Text = "" Then Exit Sub

            CallKeyDownF2(tdbcReportCode.Columns("Selection05").Text, strFind5)

            If strFind5 <> "" Then
                If strFind5.Substring(0, 1) <> "(" Then
                    If strFind5.IndexOf(";") = -1 Then
                        tdbcSel05IDFrom.Text = strFind5
                        tdbcSel05IDTo.Text = strFind5
                    Else
                        tdbcSel05IDFrom.Text = strFind5.Substring(0, strFind5.IndexOf(";"))
                        tdbcSel05IDTo.Text = strFind5.Substring(strFind5.IndexOf(";") + 1)
                    End If

                    strFind5 = ""
                Else
                    tdbcSel05IDFrom.Text = "%"
                    tdbcSel05IDTo.Text = "%"
                End If
            End If

        End If
    End Sub


    Private Sub tdbcSel05IDTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSel05IDTo.LostFocus
        If tdbcReportCode.Columns("Selection05").Text <> "DD" And tdbcReportCode.Columns("Selection05").Text <> "ND" And tdbcReportCode.Columns("Selection05").Text <> "CD" Then 'Ngày đáo hạn, Ngày thông báo, Ngày hợp đồng
            If tdbcSel05IDTo.FindStringExact(tdbcSel05IDTo.Text) = -1 Then tdbcSel05IDTo.Text = ""
        End If
    End Sub
#End Region

    '    ' Kiểm tra thấy không sự dụng nên rem lại
    '    Dim strSelection As String = ""
    '    Dim strSelectionName As String = ""
    '    Private Sub GetSelection()
    '        strSelection = ""
    '        strSelectionName = ""
    '        '====selection 1
    '        If tdbcReportCode.Columns("Selection01").Text <> "" Then
    '            strSelection &= lblSelID01From.Text
    '            If tdbcSel01IDFrom.Text.Trim = "" And tdbcSel01IDTo.Text.Trim = "" Then
    '                strSelection &= ""
    '                strSelectionName &= ""
    '            Else
    '                If tdbcSel01IDFrom.Text.Trim = tdbcSel01IDTo.Text.Trim Then
    '                    strSelection &= ": " & tdbcSel01IDFrom.Text.Trim
    '                    strSelectionName &= tdbcSel01IDFrom.Columns(1).Text
    '                Else
    '                    strSelection &= ": " & tdbcSel01IDFrom.Text.Trim & "-" & tdbcSel01IDTo.Text.Trim
    '                    strSelectionName &= tdbcSel01IDFrom.Columns(1).Text & "-" & tdbcSel01IDTo.Columns(1).Text
    '                End If
    '            End If
    '        End If
    '        strSelection &= " ## "
    '        strSelectionName &= " ## "
    '
    '        '====selection2
    '        If tdbcReportCode.Columns("Selection02").Text <> "" Then
    '            strSelection &= lblSelID02From.Text
    '            If tdbcSel02IDFrom.Text.Trim = "" And tdbcSel02IDTo.Text.Trim = "" Then
    '                strSelection &= ""
    '                strSelectionName &= ""
    '            Else
    '                If tdbcSel02IDFrom.Text.Trim = tdbcSel02IDTo.Text.Trim Then
    '                    strSelection &= ": " & tdbcSel02IDFrom.Text.Trim
    '                    strSelectionName &= tdbcSel02IDFrom.Columns(1).Text
    '                Else
    '                    strSelection &= ": " & tdbcSel02IDFrom.Text.Trim & "-" & tdbcSel02IDTo.Text.Trim
    '                    strSelectionName &= tdbcSel02IDFrom.Columns(1).Text & "-" & tdbcSel02IDTo.Columns(1).Text
    '                End If
    '            End If
    '        End If
    '        strSelection &= " ## "
    '        strSelectionName &= " ## "
    '
    '        '====selection3
    '        If tdbcReportCode.Columns("Selection03").Text <> "" Then
    '            strSelection &= lblSelID03From.Text
    '            If tdbcSel03IDFrom.Text.Trim = "" And tdbcSel03IDTo.Text.Trim = "" Then
    '                strSelection &= ""
    '                strSelectionName &= ""
    '            Else
    '                If tdbcSel03IDFrom.Text.Trim = tdbcSel03IDTo.Text.Trim Then
    '                    strSelection &= ": " & tdbcSel03IDFrom.Text.Trim
    '                    strSelectionName &= tdbcSel03IDFrom.Columns(1).Text
    '                Else
    '                    strSelection &= ": " & tdbcSel03IDFrom.Text.Trim & "-" & tdbcSel03IDTo.Text.Trim
    '                    strSelectionName &= tdbcSel03IDFrom.Columns(1).Text & "-" & tdbcSel03IDTo.Columns(1).Text
    '                End If
    '            End If
    '        End If
    '        strSelection &= " ## "
    '        strSelectionName &= " ## "
    '
    '        '====selection4
    '        If tdbcReportCode.Columns("Selection04").Text <> "" Then
    '            strSelection &= lblSelID04From.Text
    '            If tdbcSel04IDFrom.Text.Trim = "" And tdbcSel04IDTo.Text.Trim = "" Then
    '                strSelection &= ""
    '                strSelectionName &= ""
    '            Else
    '                If tdbcSel04IDFrom.Text.Trim = tdbcSel04IDTo.Text.Trim Then
    '                    strSelection &= ": " & tdbcSel04IDFrom.Text.Trim
    '                    strSelectionName &= tdbcSel04IDFrom.Columns(1).Text
    '                Else
    '                    strSelection &= ": " & tdbcSel04IDFrom.Text.Trim & "-" & tdbcSel04IDTo.Text.Trim
    '                    strSelectionName &= tdbcSel04IDFrom.Columns(1).Text & "-" & tdbcSel04IDTo.Columns(1).Text
    '                End If
    '            End If
    '        End If
    '        strSelection &= " ## "
    '        strSelectionName &= " ## "
    '
    '        '====selection5
    '        If tdbcReportCode.Columns("Selection05").Text <> "" Then
    '            strSelection &= lblSelID05From.Text
    '            If tdbcSel05IDFrom.Text.Trim = "" And tdbcSel05IDTo.Text.Trim = "" Then
    '                strSelection &= ""
    '                strSelectionName &= ""
    '            Else
    '                If tdbcSel05IDFrom.Text.Trim = tdbcSel05IDTo.Text.Trim Then
    '                    strSelection &= ": " & tdbcSel05IDFrom.Text.Trim
    '                    strSelectionName &= tdbcSel05IDFrom.Columns(1).Text
    '                Else
    '                    strSelection &= ": " & tdbcSel05IDFrom.Text.Trim & "-" & tdbcSel05IDTo.Text.Trim
    '                    strSelectionName &= tdbcSel05IDFrom.Columns(1).Text & "-" & tdbcSel05IDTo.Columns(1).Text
    '                End If
    '            End If
    '        End If
    '    End Sub

    Public Function CheckPeriod(ByVal tdbcPeriodFrom As C1.Win.C1List.C1Combo, ByVal tdbcPeriodTo As C1.Win.C1List.C1Combo) As Boolean
        Dim iMonth1, iMonth2, iYear1, iYear2 As Integer
        iMonth1 = CInt(tdbcPeriodFrom.Columns("TranMonth").Text)
        iMonth2 = CInt(tdbcPeriodTo.Columns("TranMonth").Text)
        iYear1 = CInt(tdbcPeriodFrom.Columns("TranYear").Text)
        iYear2 = CInt(tdbcPeriodTo.Columns("TranYear").Text)
        If CInt(iYear1 * 100 + iMonth1) > CInt(iYear2 * 100 + iMonth2) Then
            Return False
        End If
        Return True
    End Function

    Private Sub chkIsTime_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkIsTime.CheckedChanged
        If chkIsTime.Checked Then
            UnReadOnlyControl(True, tdbcPeriodFrom, tdbcPeriodTo)
        Else
            ReadOnlyControl(tdbcPeriodFrom, tdbcPeriodTo)
        End If
    End Sub

End Class