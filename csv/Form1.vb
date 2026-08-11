Imports System.Data.OleDb
Imports Microsoft.VisualBasic.FileIO

Public Class Form1

    ' Main Database Connection String and Password
    Dim Aft As Integer = 0
    Dim connString As String = ""
    Dim dbPassword As String = ""

    ' ==========================================================
    '  DOUBLE CLICK PASSWORD UNLOCKER FOR MAPPING GRID
    ' ==========================================================
    Private Sub dgvMapping_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvMapping.CellDoubleClick
        If dgvMapping.Columns(1).ReadOnly = False Then Exit Sub

        Dim frmPass As New Form()
        frmPass.Text = "Security Check"
        frmPass.Size = New Size(380, 160)
        frmPass.StartPosition = FormStartPosition.CenterParent
        frmPass.FormBorderStyle = FormBorderStyle.FixedDialog
        frmPass.MaximizeBox = False
        frmPass.MinimizeBox = False
        frmPass.BackColor = Color.LightGray

        Dim lbl As New Label()
        lbl.Text = " Warning: Enter Admin Password to modify mapping:"
        lbl.Location = New Point(15, 15)
        lbl.AutoSize = True
        lbl.Font = New Font("Arial", 9, FontStyle.Bold)
        frmPass.Controls.Add(lbl)

        Dim txtPass As New TextBox()
        txtPass.Location = New Point(15, 45)
        txtPass.Width = 330
        txtPass.Font = New Font("Arial", 11)
        txtPass.PasswordChar = "*"c
        frmPass.Controls.Add(txtPass)

        Dim btnUnlock As New Button()
        btnUnlock.Text = "Unlock Grid"
        btnUnlock.Location = New Point(130, 85)
        btnUnlock.Size = New Size(110, 30)
        btnUnlock.Font = New Font("Arial", 9, FontStyle.Bold)
        btnUnlock.BackColor = Color.LightGreen
        btnUnlock.DialogResult = DialogResult.OK
        frmPass.Controls.Add(btnUnlock)
        frmPass.AcceptButton = btnUnlock

        If frmPass.ShowDialog() = DialogResult.OK Then
            If txtPass.Text.ToLower() = "idscsv" Then
                dgvMapping.Columns(1).ReadOnly = False
                MessageBox.Show("Grid Unlocked! You can now modify the column mapping.", "Access Granted", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Incorrect Password! Access Denied.", "Security Alert", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    ' ==========================================================
    ' BULLETPROOF DATE PARSER 
    ' ==========================================================
    Private Function ParseSafeDate(dateStr As String) As Object
        If String.IsNullOrEmpty(dateStr) Then Return DBNull.Value
        dateStr = dateStr.Trim().Replace(".", "/").Replace(" ", "").Replace(vbCr, "").Replace(vbLf, "")

        Dim parsedDate As DateTime
        Dim formats() As String = {
            "dd/MM/yyyy", "d/M/yyyy", "d/MM/yyyy", "dd-MM-yyyy", "yyyy-MM-dd",
            "dMMyyyy", "ddMMyyyy", "dMMyy", "ddMMyy",
            "MM/yyyy", "MM-yyyy", "M/yyyy", "M-yyyy",
            "MM/yy", "MM-yy", "M/yy", "M-yy",
            "MMyyyy", "MMyy"
        }

        If DateTime.TryParseExact(dateStr, formats, System.Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, parsedDate) Then
            Return parsedDate
        ElseIf DateTime.TryParse(dateStr, System.Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, parsedDate) Then
            Return parsedDate
        End If

        Return dateStr
    End Function

    ' ==========================================================
    ' SMART DATE FORMATTER 
    ' ==========================================================
    Private Function FormatCsvDate(dateStr As String) As String
        If String.IsNullOrEmpty(dateStr) Then Return ""

        Dim cleanStr = dateStr.Trim()
        Dim lowerStr = cleanStr.ToLower().Replace("/", "-").Replace(".", "-")
        Dim txtMonths() As String = {"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"}

        Dim hasTextMonth As Boolean = False
        For Each m In txtMonths
            If lowerStr.Contains(m) Then
                hasTextMonth = True
                Exit For
            End If
        Next

        If hasTextMonth Then
            Dim parts() As String = lowerStr.Split("-"c)
            If parts.Length = 2 Then
                Dim pt1 = parts(0).Trim()
                Dim pt2 = parts(1).Trim()

                Dim yearPart As String = ""
                Dim monthPart As String = ""

                If IsNumeric(pt1) Then
                    yearPart = pt1
                    monthPart = pt2
                ElseIf IsNumeric(pt2) Then
                    yearPart = pt2
                    monthPart = pt1
                End If

                If yearPart <> "" AndAlso monthPart <> "" Then
                    If yearPart.Length = 2 Then yearPart = "20" & yearPart
                    If monthPart.Length >= 3 Then monthPart = Char.ToUpper(monthPart(0)) & monthPart.Substring(1)
                    Return monthPart & "-" & yearPart
                End If
            End If
            Return dateStr
        End If

        Dim numStr = cleanStr.Replace("-", "").Replace("/", "").Replace(".", "")
        Dim isStrictlyDigits As Boolean = True
        For Each c As Char In numStr
            If Not Char.IsDigit(c) Then
                isStrictlyDigits = False
                Exit For
            End If
        Next

        If isStrictlyDigits Then
            If numStr.Length = 8 Then
                Return numStr.Substring(0, 2) & "/" & numStr.Substring(2, 2) & "/" & numStr.Substring(4, 4)
            ElseIf numStr.Length = 7 Then
                Return "0" & numStr.Substring(0, 1) & "/" & numStr.Substring(1, 2) & "/" & numStr.Substring(3, 4)
            ElseIf numStr.Length = 6 Then
                Return numStr.Substring(0, 2) & "/" & numStr.Substring(2, 4)
            ElseIf numStr.Length = 5 Then
                Return "0" & numStr.Substring(0, 1) & "/" & numStr.Substring(1, 4)
            ElseIf numStr.Length = 4 Then
                Return numStr.Substring(0, 2) & "/" & numStr.Substring(2, 2)
            End If
        End If

        Return dateStr
    End Function

    ' ==========================================================
    '  AUTO-CREATE TABLE LOGIC
    ' ==========================================================
    Private Sub CheckAndCreateFormatTable()
        Try
            Using conn As New OleDbConnection(connString)
                conn.Open()
                Dim dt As DataTable = conn.GetSchema("Tables")
                Dim tableExists As Boolean = False

                For Each row As DataRow In dt.Rows
                    If row("TABLE_NAME").ToString().Equals("CsvFormats", StringComparison.OrdinalIgnoreCase) Then
                        tableExists = True
                        Exit For
                    End If
                Next

                If Not tableExists Then
                    Dim createQuery As String = "CREATE TABLE CsvFormats ([FormatName] VARCHAR(255), [FieldName] VARCHAR(255), [ColIndex] VARCHAR(255))"
                    Using cmd As New OleDbCommand(createQuery, conn)
                        cmd.ExecuteNonQuery()
                    End Using
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Table Creation Error: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ==========================================================
    ' 2. HELPERS (Database, Dictionaries, A-Flag Mapper)
    ' ==========================================================
    Private Function GetSupplierId(name As String) As Integer
        Using conn As New OleDbConnection(connString)
            Dim cmd As New OleDbCommand("SELECT Id FROM AcName WHERE Name=?", conn)
            cmd.Parameters.AddWithValue("@Name", name.Trim())
            conn.Open()
            Dim r = cmd.ExecuteScalar()
            If r IsNot Nothing Then Return Convert.ToInt32(r)
        End Using
        Return 0
    End Function

    Private Function GetFirmStateCode() As String
        Using conn As New OleDbConnection(connString)
            Dim cmd As New OleDbCommand("SELECT FaxNo FROM FirmInfo", conn)
            Try
                conn.Open()
                Dim rd = cmd.ExecuteReader()
                If rd.Read() AndAlso Not IsDBNull(rd("FaxNo")) Then Return rd("FaxNo").ToString().Trim()
            Catch : End Try
        End Using
        Return ""
    End Function

    Private Function GetSupplierStateCode(name As String) As String
        Using conn As New OleDbConnection(connString)
            Dim cmd As New OleDbCommand("SELECT StCode FROM AcName WHERE Name=?", conn)
            cmd.Parameters.AddWithValue("@Name", name.Trim())
            Try
                conn.Open()
                Dim r = cmd.ExecuteScalar()
                If r IsNot Nothing AndAlso Not IsDBNull(r) Then Return r.ToString().Trim()
            Catch : End Try
        End Using
        Return ""
    End Function

    Private Function GetCompanyDictionary() As Dictionary(Of String, Integer)
        Dim dict As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Using conn As New OleDbConnection(connString)
            Dim cmd As New OleDbCommand("SELECT Id, Name FROM AcName Where Acid = 14", conn)
            conn.Open()
            Dim rd = cmd.ExecuteReader()
            While rd.Read()
                dict(rd("Name").ToString()) = Convert.ToInt32(rd("Id"))
            End While
        End Using
        Return dict
    End Function

    Private Function GetProductDictionary() As Dictionary(Of String, Integer)
        Dim dict As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Using conn As New OleDbConnection(connString)
            Dim cmd As New OleDbCommand("SELECT Id, Name FROM ProMaster", conn)
            conn.Open()
            Dim rd = cmd.ExecuteReader()
            While rd.Read()
                dict(rd("Name").ToString()) = Convert.ToInt32(rd("Id"))
            End While
        End Using
        Return dict
    End Function

    Private Function GetTaxDictionary() As Dictionary(Of String, Integer)
        Dim dict As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Using conn As New OleDbConnection(connString)
            Dim cmd As New OleDbCommand("SELECT Id, Caption, SI FROM TaxPlan", conn)
            conn.Open()
            Dim rd = cmd.ExecuteReader()
            While rd.Read()
                If Not IsDBNull(rd("Caption")) AndAlso Not IsDBNull(rd("SI")) Then
                    Dim key As String = rd("Caption").ToString().Trim() & "_" & rd("SI").ToString().Trim()
                    dict(key) = Convert.ToInt32(rd("Id"))
                End If
            End While
        End Using
        Return dict
    End Function

    Private Function GetHeaderMap(fieldName As String, ByRef isTranLine As Boolean) As Integer
        isTranLine = False
        For Each row As DataGridViewRow In dgvMapping.Rows
            If Convert.ToString(row.Cells(0).Value) = fieldName Then
                Dim valStr = Convert.ToString(row.Cells(1).Value).Trim().ToUpper()
                If valStr = "" Then Return -1

                If valStr.Contains("A") Then
                    isTranLine = True
                    valStr = valStr.Replace("A", "").Trim()
                End If
                If IsNumeric(valStr) Then Return CInt(valStr)
            End If
        Next
        Return -1
    End Function

    Private Function GetMappedIndex(fieldName As String) As Integer
        For Each row As DataGridViewRow In dgvMapping.Rows
            If Convert.ToString(row.Cells(0).Value) = fieldName Then
                Dim valStr = Convert.ToString(row.Cells(1).Value).Trim().ToUpper()
                valStr = valStr.Replace("A", "")
                If valStr <> "" AndAlso IsNumeric(valStr) Then Return CInt(valStr)
            End If
        Next
        Return -1
    End Function

    Private Function SafeGet(arr As String(), i As Integer) As String
        If i >= 0 AndAlso i < arr.Length Then Return arr(i).Trim()
        Return ""
    End Function

    Private Function ToD(val As Object) As Double
        If val Is Nothing OrElse IsDBNull(val) Then Return 0
        Dim d As Double = 0
        Double.TryParse(val.ToString(), d)
        Return d
    End Function

    ' ==========================================================
    ' SMART MAPPER POPUPS
    ' ==========================================================
    Private Function ResolveMissingCompany(missingName As String, compDict As Dictionary(Of String, Integer)) As String
        Dim frmMap As New Form()
        frmMap.Text = "Smart Company Mapper"
        frmMap.Size = New Size(650, 250)
        frmMap.StartPosition = FormStartPosition.CenterParent
        frmMap.FormBorderStyle = FormBorderStyle.FixedDialog
        frmMap.MaximizeBox = False
        frmMap.MinimizeBox = False
        frmMap.BackColor = Color.LavenderBlush

        Dim lblMsg As New Label()
        lblMsg.Text = "Attention! This Company was not found: '" & missingName & "'" & vbCrLf & "Please select or SEARCH for the correct Company:"
        lblMsg.Location = New Point(15, 15)
        lblMsg.AutoSize = True
        lblMsg.Font = New Font("Arial", 10, FontStyle.Bold)
        frmMap.Controls.Add(lblMsg)

        Dim cmbMatch As New ComboBox()
        cmbMatch.Location = New Point(15, 70)
        cmbMatch.Width = 600
        cmbMatch.Font = New Font("Arial", 11)
        cmbMatch.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        cmbMatch.AutoCompleteSource = AutoCompleteSource.ListItems

        For Each key As String In compDict.Keys : cmbMatch.Items.Add(key) : Next

        Dim searchStr As String = missingName.Trim()
        If searchStr.Length >= 3 Then searchStr = searchStr.Substring(0, 3).ToLower() Else searchStr = searchStr.ToLower()

        Dim bestMatchIndex As Integer = -1
        For i As Integer = 0 To cmbMatch.Items.Count - 1
            If cmbMatch.Items(i).ToString().ToLower().StartsWith(searchStr) Then
                bestMatchIndex = i
                Exit For
            End If
        Next
        If bestMatchIndex <> -1 Then cmbMatch.SelectedIndex = bestMatchIndex
        frmMap.Controls.Add(cmbMatch)

        Dim btnUpdate As New Button()
        btnUpdate.Text = "Update & Continue"
        btnUpdate.Location = New Point(250, 150)
        btnUpdate.Size = New Size(150, 40)
        btnUpdate.Font = New Font("Arial", 10, FontStyle.Bold)
        btnUpdate.BackColor = Color.LightGreen
        btnUpdate.DialogResult = DialogResult.OK
        frmMap.Controls.Add(btnUpdate)
        frmMap.AcceptButton = btnUpdate

        If frmMap.ShowDialog() = DialogResult.OK Then
            Dim selectedTxt As String = cmbMatch.Text.Trim()
            If compDict.ContainsKey(selectedTxt) Then
                Return selectedTxt
            ElseIf cmbMatch.SelectedItem IsNot Nothing Then
                Return cmbMatch.SelectedItem.ToString()
            End If
        End If
        Return ""
    End Function

    Private Function ResolveMissingProduct(missingName As String, prodDict As Dictionary(Of String, Integer)) As String
        Dim frmMap As New Form()
        frmMap.Text = "Smart Item Mapper"
        frmMap.Size = New Size(650, 250)
        frmMap.StartPosition = FormStartPosition.CenterParent
        frmMap.FormBorderStyle = FormBorderStyle.FixedDialog
        frmMap.MaximizeBox = False
        frmMap.MinimizeBox = False
        frmMap.BackColor = Color.LightYellow

        Dim lblMsg As New Label()
        lblMsg.Text = "Product : " & missingName & vbCrLf & " Not Found, Please select by searching:"
        lblMsg.ForeColor = Color.Red
        lblMsg.Location = New Point(15, 15)
        lblMsg.AutoSize = True
        lblMsg.Font = New Font("Arial", 10, FontStyle.Bold)
        frmMap.Controls.Add(lblMsg)

        Dim cmbMatch As New ComboBox()
        cmbMatch.Location = New Point(15, 70)
        cmbMatch.Width = 600
        cmbMatch.Font = New Font("Arial", 11)
        cmbMatch.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        cmbMatch.AutoCompleteSource = AutoCompleteSource.ListItems

        For Each key As String In prodDict.Keys : cmbMatch.Items.Add(key) : Next

        Dim searchStr As String = missingName.Trim()
        If searchStr.Length >= 3 Then searchStr = searchStr.Substring(0, 3).ToLower() Else searchStr = searchStr.ToLower()

        Dim bestMatchIndex As Integer = -1
        For i As Integer = 0 To cmbMatch.Items.Count - 1
            If cmbMatch.Items(i).ToString().ToLower().StartsWith(searchStr) Then
                bestMatchIndex = i
                Exit For
            End If
        Next
        If bestMatchIndex <> -1 Then cmbMatch.SelectedIndex = bestMatchIndex
        frmMap.Controls.Add(cmbMatch)

        Dim btnUpdate As New Button()
        btnUpdate.Text = "Update & Continue"
        btnUpdate.Location = New Point(240, 150)
        btnUpdate.Size = New Size(150, 40)
        btnUpdate.Font = New Font("Arial", 10, FontStyle.Bold)
        btnUpdate.BackColor = Color.LightGreen
        btnUpdate.DialogResult = DialogResult.OK
        frmMap.Controls.Add(btnUpdate)
        frmMap.AcceptButton = btnUpdate

        If frmMap.ShowDialog() = DialogResult.OK Then
            Dim selectedTxt As String = cmbMatch.Text.Trim()
            If prodDict.ContainsKey(selectedTxt) Then
                Return selectedTxt
            ElseIf cmbMatch.SelectedItem IsNot Nothing Then
                Return cmbMatch.SelectedItem.ToString()
            End If
        End If
        Return ""
    End Function

    ' ==========================================================
    ' 3. FORM LOAD
    ' ==========================================================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim configFilePath As String = Application.StartupPath & "\db_config.txt"
        Dim dbPath As String = ""

        If IO.File.Exists(configFilePath) Then
            dbPath = IO.File.ReadAllText(configFilePath).Trim()
        End If

        If dbPath = "" OrElse Not IO.File.Exists(dbPath) Then
            MessageBox.Show("Database not found! Please select your database.", "First Time Setup", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Dim ofd As New OpenFileDialog()
            ofd.Filter = "Access Database|*.mdb;*.accdb"
            ofd.Title = "Select Database File"

            If ofd.ShowDialog() = DialogResult.OK Then
                dbPath = ofd.FileName
                IO.File.WriteAllText(configFilePath, dbPath)
            Else
                Application.Exit()
                Return
            End If
        End If

        connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=" & dbPassword & ";Persist Security Info=False;"

        dgvMapping.Columns.Clear()
        dgvMapping.Columns.Add("Field", "Software Field")
        dgvMapping.Columns.Add("Index", "CSV Column No")
        dgvMapping.Columns(0).ReadOnly = True
        dgvMapping.Columns(1).ReadOnly = True
        dgvMapping.AllowUserToAddRows = False

        Dim fields = {"Invoice No", "Date", "Company Name", "Item Name", "Batch No", "Expiry", "Qty", "Free", "Deal", "MRP", "PRate", "SRate", "TRate", "DisRate", "DisAmt", "CGST Percentage", "SGST Percentage", "IGST Percentage", "HSN", "CS", "NetAmt", "Footer (Y/N)"}

        For Each f In fields
            dgvMapping.Rows.Add(f, "")
        Next

        CheckAndCreateFormatTable()
        LoadFormats()
        'LoadBillsForExport()
        If cmbBillType.Items.Count > 0 Then cmbBillType.SelectedIndex = 0
    End Sub

    ' ==========================================================
    ' 4. CSV LOADING & LIVE CALCULATION
    ' ==========================================================
    Private Sub btnSelectFile_Click(sender As Object, e As EventArgs) Handles btnSelectFile.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "CSV/TXT Files|*.csv;*.txt"
        If ofd.ShowDialog() <> DialogResult.OK Then Exit Sub

        txtFilePath.Text = ofd.FileName
        Dim dt As New DataTable()

        For Each r As DataGridViewRow In dgvMapping.Rows
            Dim fName = Convert.ToString(r.Cells(0).Value)
            If fName <> "" AndAlso fName <> "Invoice No" AndAlso fName <> "Date" AndAlso fName <> "Company Name" AndAlso fName <> "Footer (Y/N)" Then
                dt.Columns.Add(fName)
            End If
        Next

        Dim calcCols = {"Taxable", "CGST", "SGST", "IGST", "NRate", "AcRate"}
        For Each col In calcCols
            If Not dt.Columns.Contains(col) Then dt.Columns.Add(col)
        Next

        Try
            Dim compDict = GetCompanyDictionary()
            Dim firmStateCode As String = GetFirmStateCode()
            Dim supStateCode As String = ""
            Dim companyResolved As Boolean = False

            Dim invFound As Boolean = False
            Dim dateFound As Boolean = False
            Dim compFound As Boolean = False

            txtInvoiceNo.Text = ""
            txtDate.Text = ""
            txtCompanyName.Text = ""

            Dim footerMapVal As String = ""
            For i = 0 To dgvMapping.Rows.Count - 1
                If Convert.ToString(dgvMapping.Rows(i).Cells(0).Value) = "Footer (Y/N)" Then
                    footerMapVal = Convert.ToString(dgvMapping.Rows(i).Cells(1).Value).Trim().ToUpper()
                    Exit For
                End If
            Next
            Dim hasManualFooter As Boolean = (footerMapVal = "Y" OrElse footerMapVal = "YES")

            Dim delimiter As String = ","
            Dim firstLineStr As String = ""

            Using reader As New IO.StreamReader(ofd.FileName)
                firstLineStr = reader.ReadLine()
            End Using

            If firstLineStr IsNot Nothing AndAlso firstLineStr.Contains(vbTab) Then
                delimiter = vbTab
            End If

            Dim allLines As New List(Of String())
            Using parser As New TextFieldParser(ofd.FileName)
                parser.SetDelimiters(delimiter)
                While Not parser.EndOfData
                    Dim data = parser.ReadFields()
                    If data IsNot Nothing AndAlso data.Length > 0 Then
                        allLines.Add(data)
                    End If
                End While
            End Using

            Dim hasIndicators As Boolean = False
            If allLines.Count > 0 AndAlso allLines(0).Length > 0 Then
                Dim firstField = allLines(0)(0).Trim().ToUpper()
                If firstField = "H" OrElse firstField = "T" OrElse firstField = "F" Then
                    hasIndicators = True
                End If
            End If

            For lineIdx As Integer = 0 To allLines.Count - 1
                Dim data = allLines(lineIdx)
                Dim lineType As String = ""
                If data.Length > 0 Then lineType = data(0).Trim().ToUpper()

                If Not hasIndicators Then
                    If lineIdx = 0 Then
                        lineType = "H"
                    ElseIf lineIdx = allLines.Count - 1 AndAlso hasManualFooter Then
                        lineType = "F"
                    Else
                        lineType = "T"
                    End If
                End If

                Dim isInvTran As Boolean, isDateTran As Boolean, isCompTran As Boolean
                Dim idxInv = GetHeaderMap("Invoice No", isInvTran)
                Dim idxDate = GetHeaderMap("Date", isDateTran)
                Dim idxComp = GetHeaderMap("Company Name", isCompTran)

                ' --- HEADER PROCESSING ---
                If lineType = "H" Then
                    If Not isInvTran AndAlso Not invFound Then
                        If idxInv <> -1 AndAlso SafeGet(data, idxInv) <> "" Then
                            txtInvoiceNo.Text = SafeGet(data, idxInv)
                            invFound = True
                        ElseIf idxInv = -1 Then
                            txtInvoiceNo.Text = If(hasIndicators, SafeGet(data, 1), SafeGet(data, 0))
                            invFound = True
                        End If
                    End If

                    If Not isDateTran AndAlso Not dateFound Then
                        If idxDate <> -1 AndAlso SafeGet(data, idxDate) <> "" Then
                            txtDate.Text = FormatCsvDate(SafeGet(data, idxDate))
                            dateFound = True
                        ElseIf idxDate = -1 Then
                            txtDate.Text = FormatCsvDate(If(hasIndicators, SafeGet(data, 2), SafeGet(data, 1)))
                            dateFound = True
                        End If
                    End If

                    If Not isCompTran AndAlso Not compFound Then
                        If idxComp <> -1 AndAlso SafeGet(data, idxComp) <> "" Then
                            txtCompanyName.Text = SafeGet(data, idxComp)
                            compFound = True
                        ElseIf idxComp = -1 Then
                            txtCompanyName.Text = If(hasIndicators, SafeGet(data, 3), SafeGet(data, 2))
                            compFound = True
                        End If
                    End If
                End If

                ' --- TRANSACTION/ITEMS PROCESSING ---
                If lineType = "T" Then
                    If isInvTran AndAlso Not invFound Then
                        If idxInv <> -1 AndAlso SafeGet(data, idxInv) <> "" Then
                            txtInvoiceNo.Text = SafeGet(data, idxInv)
                            invFound = True
                        End If
                    End If

                    If isDateTran AndAlso Not dateFound Then
                        If idxDate <> -1 AndAlso SafeGet(data, idxDate) <> "" Then
                            txtDate.Text = FormatCsvDate(SafeGet(data, idxDate))
                            dateFound = True
                        End If
                    End If

                    If isCompTran AndAlso Not compFound Then
                        If idxComp <> -1 AndAlso SafeGet(data, idxComp) <> "" Then
                            txtCompanyName.Text = SafeGet(data, idxComp)
                            compFound = True
                        End If
                    End If

                    If Not companyResolved AndAlso txtCompanyName.Text.Trim() <> "" Then
                        Dim extComp = txtCompanyName.Text.Trim()
                        If Not compDict.ContainsKey(extComp) Then
                            Dim correctComp = ResolveMissingCompany(extComp, compDict)
                            If correctComp <> "" Then
                                txtCompanyName.Text = correctComp
                            End If
                        End If
                        supStateCode = GetSupplierStateCode(txtCompanyName.Text.Trim())
                        companyResolved = True
                    End If

                    Dim row(dt.Columns.Count - 1) As Object
                    Dim dtColIdx As Integer = 0

                    For i = 0 To dgvMapping.Rows.Count - 1
                        Dim fieldName = Convert.ToString(dgvMapping.Rows(i).Cells(0).Value)

                        If fieldName <> "" AndAlso fieldName <> "Invoice No" AndAlso fieldName <> "Date" AndAlso fieldName <> "Company Name" AndAlso fieldName <> "Footer (Y/N)" Then
                            Dim cellVal = SafeGet(data, GetMappedIndex(fieldName))
                            If fieldName = "Expiry" Then
                                row(dtColIdx) = FormatCsvDate(cellVal)
                            Else
                                row(dtColIdx) = cellVal
                            End If
                            dtColIdx += 1
                        End If
                    Next

                    Dim idxCGST = GetMappedIndex("CGST Percentage")
                    Dim idxSGST = GetMappedIndex("SGST Percentage")
                    Dim idxIGST = GetMappedIndex("IGST Percentage")

                    Dim cper As Double = 0
                    Dim sper As Double = 0
                    Dim iper As Double = 0

                    If idxCGST <> -1 AndAlso (idxCGST = idxSGST OrElse idxCGST = idxIGST) Then
                        Dim totalTaxPer = ToD(SafeGet(data, idxCGST))
                        If firmStateCode <> "" AndAlso firmStateCode.Equals(supStateCode, StringComparison.OrdinalIgnoreCase) Then
                            cper = totalTaxPer / 2
                            sper = totalTaxPer / 2
                            iper = 0
                        Else
                            cper = 0
                            sper = 0
                            iper = totalTaxPer
                        End If
                    Else
                        cper = ToD(SafeGet(data, idxCGST))
                        sper = ToD(SafeGet(data, idxSGST))
                        iper = ToD(SafeGet(data, idxIGST))
                    End If

                    If dt.Columns.Contains("CGST Percentage") Then row(dt.Columns.IndexOf("CGST Percentage")) = cper.ToString("0.##")
                    If dt.Columns.Contains("SGST Percentage") Then row(dt.Columns.IndexOf("SGST Percentage")) = sper.ToString("0.##")
                    If dt.Columns.Contains("IGST Percentage") Then row(dt.Columns.IndexOf("IGST Percentage")) = iper.ToString("0.##")

                    ' =======================================================
                    ' NEW ITEM-LEVEL DEAL CALCULATION LOGIC
                    ' =======================================================
                    Dim qty As Double = ToD(SafeGet(data, GetMappedIndex("Qty")))
                    Dim free As Double = ToD(SafeGet(data, GetMappedIndex("Free")))
                    Dim originalPRate As Double = ToD(SafeGet(data, GetMappedIndex("PRate")))
                    Dim disRate As Double = ToD(SafeGet(data, GetMappedIndex("DisRate")))
                    Dim dealVal As String = SafeGet(data, GetMappedIndex("Deal")).Trim()

                    ' AcRate me purana original rate save karenge
                    Dim acRate As Double = originalPRate
                    Dim pRate As Double = originalPRate

                    ' Step 1: Adjust PRate for 1 item if deal exists (e.g., 5+1)
                    If dealVal.Contains("+") Then
                        Dim parts() As String = dealVal.Split("+"c)
                        If parts.Length = 2 AndAlso IsNumeric(parts(0)) AndAlso IsNumeric(parts(1)) Then
                            Dim pre As Double = Val(parts(0))
                            Dim suf As Double = Val(parts(1))
                            If pre > 0 AndAlso (pre + suf) > 0 Then
                                ' Naya PRate Formula: (Original PRate * Pre) / (Pre + Suf)
                                pRate = (originalPRate * pre) / (pre + suf)
                            End If
                        End If
                    End If

                    ' Step 2: Base Amount (Naya sasta PRate * Qty)
                    Dim baseAmount As Double = pRate * qty

                    ' Step 3: Discount Amount 
                    Dim disAmt As Double = 0
                    If disRate > 0 Then
                        disAmt = baseAmount * (disRate / 100)
                    End If

                    ' Step 4: Taxable Amount
                    Dim taxable As Double = baseAmount - disAmt

                    ' Step 5: Tax Calculation
                    Dim cgst As Double = 0, sgst As Double = 0, igst As Double = 0
                    If iper > 0 Then
                        igst = taxable * (iper / 100)
                    Else
                        cgst = taxable * (cper / 100)
                        sgst = taxable * (sper / 100)
                    End If

                    ' Step 6: Net Amount
                    Dim taxAmt As Double = cgst + sgst + igst
                    Dim netAmt As Double = taxable + taxAmt

                    ' Step 7: Net Rate (Zero Division Safety ke sath)
                    Dim nRate As Double = 0
                    If (qty + free) > 0 Then
                        nRate = netAmt / (qty + free)
                    End If
                    ' =======================================================

                    ' Grid me update karna
                    If dt.Columns.Contains("PRate") Then row(dt.Columns.IndexOf("PRate")) = pRate.ToString("0.00") ' Naya adjusted PRate
                    If dt.Columns.Contains("AcRate") Then row(dt.Columns.IndexOf("AcRate")) = acRate.ToString("0.00") ' Purana original PRate
                    If dt.Columns.Contains("DisAmt") Then row(dt.Columns.IndexOf("DisAmt")) = disAmt.ToString("0.00")
                    If dt.Columns.Contains("DisRate") Then row(dt.Columns.IndexOf("DisRate")) = disRate.ToString("0.00")
                    If dt.Columns.Contains("NetAmt") Then row(dt.Columns.IndexOf("NetAmt")) = netAmt.ToString("0.00")

                    row(dt.Columns.IndexOf("Taxable")) = taxable.ToString("0.00")
                    row(dt.Columns.IndexOf("CGST")) = cgst.ToString("0.00")
                    row(dt.Columns.IndexOf("SGST")) = sgst.ToString("0.00")
                    row(dt.Columns.IndexOf("IGST")) = igst.ToString("0.00")
                    row(dt.Columns.IndexOf("NRate")) = nRate.ToString("0.00")

                    dt.Rows.Add(row)
                End If

                ' --- FOOTER PROCESSING ---
                If lineType = "F" Then
                    Dim totalVal As String = ""
                    If hasIndicators Then
                        totalVal = SafeGet(data, 1)
                    Else
                        totalVal = SafeGet(data, 0)
                        If totalVal = "" AndAlso data.Length > 1 Then totalVal = SafeGet(data, 1)
                    End If

                    Try
                        If Me.Controls.Find("lblTotalAmount", True).Length > 0 Then
                            Me.Controls("lblTotalAmount").Text = "Bill Total: ₹ " & totalVal
                        End If
                    Catch
                    End Try
                End If
            Next

            dgvPreview.DataSource = dt
            dgvPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            If dgvPreview.Columns.Contains("NRate") Then dgvPreview.Columns("NRate").DefaultCellStyle.BackColor = Color.LightCyan
            If dgvPreview.Columns.Contains("DisAmt") Then dgvPreview.Columns("DisAmt").DefaultCellStyle.BackColor = Color.LightYellow

        Catch ex As Exception
            MessageBox.Show("Load Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ==========================================================
    ' 5. DATABASE SAVE (Separate Sno Series & Dynamic Inv/Batch Suffix)
    ' ==========================================================
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        btnImport.Enabled = False
        Dim originalBtnText As String = btnImport.Text
        btnImport.Text = "Saving... Please Wait ⏳"
        Application.DoEvents()

        Dim billType As String = "I"
        If cmbBillType.SelectedItem IsNot Nothing Then
            billType = cmbBillType.SelectedItem.ToString().Trim()
        End If

        Try
            Dim companyName As String = txtCompanyName.Text.Trim()
            Dim supId As Integer = GetSupplierId(companyName)

            Dim compDict = GetCompanyDictionary()
            Dim prodDict = GetProductDictionary()
            Dim taxDict = GetTaxDictionary()

            If supId = 0 Then
                Dim correctCompName As String = ResolveMissingCompany(companyName, compDict)
                If correctCompName <> "" Then
                    txtCompanyName.Text = correctCompName
                    companyName = correctCompName
                    supId = compDict(correctCompName)
                Else
                    MessageBox.Show("Import Cancelled! You did not resolve the Company Name.", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            Dim invNo As String = txtInvoiceNo.Text.Trim()
            If invNo <> "" AndAlso billType = "E" Then
                If Not invNo.EndsWith("-E", StringComparison.OrdinalIgnoreCase) Then
                    invNo &= "-E"
                End If
            End If

            If invNo = "" Then
                MessageBox.Show("Invoice Number is blank! Please verify.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                Using connCheck As New OleDbConnection(connString)
                    connCheck.Open()
                    Dim qCheck = "SELECT COUNT(*) FROM PurchaseMaster WHERE SupId = ? AND InvNo = ? AND PurType = ?"
                    Using cmdCheck = New OleDbCommand(qCheck, connCheck)
                        cmdCheck.Parameters.AddWithValue("?", supId)
                        cmdCheck.Parameters.AddWithValue("?", invNo)
                        cmdCheck.Parameters.AddWithValue("?", billType)

                        Dim billCount = Convert.ToInt32(cmdCheck.ExecuteScalar())
                        If billCount > 0 Then
                            MessageBox.Show("Warning! Invoice number '" & invNo & "' for '" & companyName & "' already exists as Type '" & billType & "'." & vbCrLf & vbCrLf & "Duplicate uploads are not allowed!", "Duplicate Bill Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error occurred while checking for duplicates: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            Dim pendingUpdates As New Dictionary(Of String, String)

            For Each row As DataGridViewRow In dgvPreview.Rows
                If Not row.IsNewRow Then
                    Dim itemName = Convert.ToString(row.Cells("Item Name").Value).Trim()

                    If itemName <> "" AndAlso Not prodDict.ContainsKey(itemName) Then
                        If Not pendingUpdates.ContainsKey(itemName) Then
                            Dim correctName As String = ResolveMissingProduct(itemName, prodDict)
                            If correctName <> "" Then
                                pendingUpdates.Add(itemName, correctName)
                            Else
                                MessageBox.Show("Import Cancelled! You did not resolve '" & itemName & "'.", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Return
                            End If
                        End If
                    End If
                End If
            Next

            If pendingUpdates.Count > 0 Then
                Dim msg As String = "The following items have different names in the CSV:" & vbCrLf & vbCrLf
                For Each kvp In pendingUpdates
                    msg &= "Old Name (DB) : " & kvp.Value & vbCrLf
                    msg &= "New Name (CSV) : " & kvp.Key & vbCrLf
                    msg &= New String("-"c, 35) & vbCrLf
                Next
                msg &= vbCrLf & "Do you want to PERMANENTLY REPLACE these names in the database (ProMaster)?" & vbCrLf & "(Yes = Update DB and Save, No = Cancel Import)"

                Dim dialogRes = ScrollMessageBox.Show(msg, "Confirm Permanent Name Update", MessageBoxButtons.YesNo)

                If dialogRes = DialogResult.Yes Then
                    Try
                        Using connUpd As New OleDbConnection(connString)
                            connUpd.Open()
                            For Each kvp In pendingUpdates
                                Dim csvName = kvp.Key
                                Dim dbName = kvp.Value
                                Dim pId = prodDict(dbName)

                                Dim qUpdMaster = "UPDATE ProMaster SET [Name] = ? WHERE Id = ?"
                                Using cmdUpd = New OleDbCommand(qUpdMaster, connUpd)
                                    cmdUpd.Parameters.AddWithValue("?", csvName)
                                    cmdUpd.Parameters.AddWithValue("?", pId)
                                    cmdUpd.ExecuteNonQuery()
                                End Using

                                If prodDict.ContainsKey(dbName) Then prodDict.Remove(dbName)
                                If Not prodDict.ContainsKey(csvName) Then prodDict.Add(csvName, pId)
                            Next
                        End Using
                    Catch ex As Exception
                        MessageBox.Show("Error updating ProMaster: " & ex.Message, "Master Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try
                Else
                    MessageBox.Show("Import Cancelled! You can safely remap the products and try again.", "Import Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If
            End If

            Using conn As New OleDbConnection(connString)
                conn.Open()
                Dim transaction As OleDbTransaction = conn.BeginTransaction()

                Try
                    Dim nextSno As Integer = 1

                    ' STRICT SEPARATION LOGIC
                    Dim qSno = "SELECT MAX(Sno) FROM PurchaseMaster WHERE PurType = ?"
                    Using cmdSno = New OleDbCommand(qSno, conn, transaction)
                        cmdSno.Parameters.AddWithValue("?", billType)
                        Dim objSno = cmdSno.ExecuteScalar()
                        If objSno IsNot DBNull.Value AndAlso objSno IsNot Nothing Then
                            nextSno = Convert.ToInt32(objSno) + 1
                        End If
                    End Using

                    Dim tTaxable As Double = 0, tCGST As Double = 0, tSGST As Double = 0, tIGST As Double = 0, tNet As Double = 0
                    Dim sumDisAmt As Double = 0
                    Dim tax5 As Double = 0, tax12 As Double = 0, tax18 As Double = 0, tax28 As Double = 0

                    For Each row As DataGridViewRow In dgvPreview.Rows
                        If Not row.IsNewRow Then
                            Dim curTaxable = ToD(row.Cells("Taxable").Value)
                            Dim curCGST = ToD(row.Cells("CGST").Value)
                            Dim curSGST = ToD(row.Cells("SGST").Value)
                            Dim curIGST = ToD(row.Cells("IGST").Value)

                            tTaxable += curTaxable
                            tCGST += curCGST
                            tSGST += curSGST
                            tIGST += curIGST
                            tNet += ToD(row.Cells("NetAmt").Value)
                            sumDisAmt += ToD(row.Cells("DisAmt").Value)

                            Dim cper = ToD(row.Cells("CGST Percentage").Value)
                            Dim sper = ToD(row.Cells("SGST Percentage").Value)
                            Dim iper = ToD(row.Cells("IGST Percentage").Value)
                            Dim totalTaxPer As Double = If(iper > 0, iper, cper + sper)

                            Dim currentItemTaxAmt As Double = curCGST + curSGST + curIGST

                            Select Case totalTaxPer
                                Case 5 : tax5 += currentItemTaxAmt
                                Case 12 : tax12 += currentItemTaxAmt
                                Case 18 : tax18 += currentItemTaxAmt
                                Case 28 : tax28 += currentItemTaxAmt
                            End Select
                        End If
                    Next

                    Dim roundedGTotal As Double = Math.Round(tNet, 0, MidpointRounding.AwayFromZero)
                    Dim roundOffAmt As Double = roundedGTotal - tNet

                    Dim newMasterId As Integer = 0
                    Dim qPM = "INSERT INTO PurchaseMaster (Sno, InvNo, InvDate, [Date], SupId, Total, TaxableT, TSGST, TCGST, TIGST, TaxAmt, InvAmt, GTotal, PurType, Status, InvType, DisAmt, Roundoff, PaidAmt, CompId, Taxable5, Taxable12, Taxable18, Taxable28) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?, ?, ?, 'CREDIT', ?, ?, 0, 0, ?, ?, ?, ?)"
                    Using cmd = New OleDbCommand(qPM, conn, transaction)
                        cmd.Parameters.AddWithValue("?", nextSno)
                        cmd.Parameters.AddWithValue("?", invNo)
                        cmd.Parameters.AddWithValue("?", ParseSafeDate(txtDate.Text))
                        cmd.Parameters.AddWithValue("?", ParseSafeDate(txtDate.Text))
                        cmd.Parameters.AddWithValue("?", supId)
                        cmd.Parameters.AddWithValue("?", tTaxable)
                        cmd.Parameters.AddWithValue("?", tTaxable)
                        cmd.Parameters.AddWithValue("?", tSGST)
                        cmd.Parameters.AddWithValue("?", tCGST)
                        cmd.Parameters.AddWithValue("?", tIGST)
                        cmd.Parameters.AddWithValue("?", tSGST + tCGST + tIGST)
                        cmd.Parameters.AddWithValue("?", roundedGTotal)
                        cmd.Parameters.AddWithValue("?", roundedGTotal)

                        cmd.Parameters.AddWithValue("?", billType) ' PurType
                        cmd.Parameters.AddWithValue("?", billType) ' Status

                        cmd.Parameters.AddWithValue("?", sumDisAmt)
                        cmd.Parameters.AddWithValue("?", roundOffAmt)
                        cmd.Parameters.AddWithValue("?", tax5)
                        cmd.Parameters.AddWithValue("?", tax12)
                        cmd.Parameters.AddWithValue("?", tax18)
                        cmd.Parameters.AddWithValue("?", tax28)

                        cmd.ExecuteNonQuery()
                    End Using

                    Using cmdId = New OleDbCommand("SELECT MAX(Id) FROM PurchaseMaster", conn, transaction)
                        newMasterId = Convert.ToInt32(cmdId.ExecuteScalar())
                    End Using

                    For Each row As DataGridViewRow In dgvPreview.Rows
                        If Not row.IsNewRow Then
                            Dim itemName = Convert.ToString(row.Cells("Item Name").Value).Trim()
                            If itemName = "" Then Continue For

                            Dim pId = prodDict(itemName)
                            Dim originalBatch = row.Cells("Batch No").Value.ToString().Trim()

                            Dim saveBatch = originalBatch
                            If billType = "E" AndAlso saveBatch <> "" Then
                                If Not saveBatch.EndsWith("-E", StringComparison.OrdinalIgnoreCase) Then
                                    saveBatch &= "-E"
                                End If
                            End If

                            Dim curExp = row.Cells("Expiry").Value.ToString()
                            Dim curDeal = row.Cells("Deal").Value.ToString().Trim()
                            Dim qty = ToD(row.Cells("Qty").Value)
                            Dim free = ToD(row.Cells("Free").Value)
                            Dim pRate = ToD(row.Cells("PRate").Value)
                            Dim acRate = If(dgvPreview.Columns.Contains("AcRate"), ToD(row.Cells("AcRate").Value), pRate)
                            Dim disRate = ToD(row.Cells("DisRate").Value)
                            Dim disAmt = ToD(row.Cells("DisAmt").Value)

                            Dim curTaxable = ToD(row.Cells("Taxable").Value)
                            Dim curCGST = ToD(row.Cells("CGST").Value)
                            Dim curSGST = ToD(row.Cells("SGST").Value)
                            Dim curIGST = ToD(row.Cells("IGST").Value)
                            Dim nRate = ToD(row.Cells("NRate").Value)

                            Dim curHSN = Convert.ToString(row.Cells("HSN").Value).Trim()
                            Dim curCS = ToD(row.Cells("CS").Value)
                            Dim tRate = ToD(row.Cells("TRate").Value)

                            Dim cper = ToD(row.Cells("CGST Percentage").Value)
                            Dim sper = ToD(row.Cells("SGST Percentage").Value)
                            Dim iper = ToD(row.Cells("IGST Percentage").Value)

                            Dim isLocal As Boolean = (iper = 0)
                            Dim totalItemTaxPer As Double = If(isLocal, cper + sper, iper)
                            Dim si As String = If(isLocal, "LocalPur", "InterStatePur")

                            Dim taxKey As String = totalItemTaxPer.ToString() & "_" & si
                            Dim taxId As Integer = If(taxDict.ContainsKey(taxKey), taxDict(taxKey), 0)

                            Dim saveTxRate As Double = If(isLocal, cper, iper)
                            Dim afterDisAmt As Double = curTaxable - disAmt
                            Dim taxAmt As Double = curCGST + curSGST + curIGST

                            Dim batchId As Integer = 0

                            If originalBatch = "" Then
                                Dim qCheckEmpty = "SELECT MAX(Id) FROM BatchMaster WHERE PId=" & pId & " AND [Type]='" & billType & "' AND (Batch IS NULL OR Batch='')"
                                Dim existsObj = New OleDbCommand(qCheckEmpty, conn, transaction).ExecuteScalar()
                                If existsObj IsNot Nothing AndAlso existsObj IsNot DBNull.Value Then batchId = Convert.ToInt32(existsObj)
                            Else
                                Dim sNormal = originalBatch
                                If sNormal.EndsWith("-E", StringComparison.OrdinalIgnoreCase) Then
                                    sNormal = sNormal.Substring(0, sNormal.Length - 2)
                                End If
                                Dim sE = sNormal & "-E"

                                Dim qCheck As String
                                If billType = "E" Then
                                    qCheck = "SELECT MAX(Id) FROM BatchMaster WHERE PId=" & pId & " AND [Type]='" & billType & "' AND (Batch='" & sNormal & "' OR Batch='" & sE & "')"
                                Else
                                    qCheck = "SELECT MAX(Id) FROM BatchMaster WHERE PId=" & pId & " AND [Type]='" & billType & "' AND Batch='" & sNormal & "'"
                                End If

                                Dim existsObj = New OleDbCommand(qCheck, conn, transaction).ExecuteScalar()
                                If existsObj IsNot Nothing AndAlso existsObj IsNot DBNull.Value Then batchId = Convert.ToInt32(existsObj)
                            End If

                            If batchId > 0 Then
                                Dim qUpd = "UPDATE BatchMaster SET Batch=?, Qty=Qty+?, FreeQty=FreeQty+?, NRate=?, Deal=?, TRate=? WHERE Id=?"
                                Using cmdU = New OleDbCommand(qUpd, conn, transaction)
                                    cmdU.Parameters.AddWithValue("?", If(saveBatch = "", DBNull.Value, saveBatch))
                                    cmdU.Parameters.AddWithValue("?", qty + free)
                                    cmdU.Parameters.AddWithValue("?", free)
                                    cmdU.Parameters.AddWithValue("?", nRate)
                                    cmdU.Parameters.AddWithValue("?", curDeal)
                                    cmdU.Parameters.AddWithValue("?", tRate)
                                    cmdU.Parameters.AddWithValue("?", batchId)
                                    cmdU.ExecuteNonQuery()
                                End Using
                            Else
                                Dim qIns = "INSERT INTO BatchMaster (PId, Batch, [Exp], Qty, FreeQty, MRP, PRate, SRate, DisRate, Deal, TRate, [Type], BillDate, BillNo, PartyName, NRate) VALUES (?,?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?)"
                                Using cmdI = New OleDbCommand(qIns, conn, transaction)
                                    cmdI.Parameters.AddWithValue("?", pId)
                                    cmdI.Parameters.AddWithValue("?", If(saveBatch = "", DBNull.Value, saveBatch))
                                    cmdI.Parameters.AddWithValue("?", ParseSafeDate(curExp))
                                    cmdI.Parameters.AddWithValue("?", qty + free)
                                    cmdI.Parameters.AddWithValue("?", free)
                                    cmdI.Parameters.AddWithValue("?", ToD(row.Cells("MRP").Value))
                                    cmdI.Parameters.AddWithValue("?", pRate)
                                    cmdI.Parameters.AddWithValue("?", ToD(row.Cells("SRate").Value))
                                    cmdI.Parameters.AddWithValue("?", disRate)
                                    cmdI.Parameters.AddWithValue("?", curDeal)
                                    cmdI.Parameters.AddWithValue("?", tRate)
                                    cmdI.Parameters.AddWithValue("?", billType)
                                    cmdI.Parameters.AddWithValue("?", ParseSafeDate(txtDate.Text))
                                    cmdI.Parameters.AddWithValue("?", invNo)
                                    cmdI.Parameters.AddWithValue("?", companyName)
                                    cmdI.Parameters.AddWithValue("?", nRate)
                                    cmdI.ExecuteNonQuery()
                                End Using

                                Using cmdBatchId = New OleDbCommand("SELECT MAX(Id) FROM BatchMaster", conn, transaction)
                                    batchId = Convert.ToInt32(cmdBatchId.ExecuteScalar())
                                End Using
                            End If

                            ' FIX: Added 'Deal' to PurchaseDetails with a total of 32 parameters!
                            Dim qPD = "INSERT INTO PurchaseDetails (PMid, Pid, Batch, [Exp], Qty, [Free], tQty, MRP, PRate, SRate, DisRate, DisAmt, TxRate, Taxable, SGST, CGST, IGST, NetAmt, NRate, PurType, TaxId, HSN, BatchId, AcRate, TaxAmt, LB, CS, BoxQty, Scdl, AL, AftDisAmt, Deal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                            Using cmdPD = New OleDbCommand(qPD, conn, transaction)
                                cmdPD.Parameters.AddWithValue("PMid", newMasterId)
                                cmdPD.Parameters.AddWithValue("Pid", pId)
                                cmdPD.Parameters.AddWithValue("Batch", If(saveBatch = "", DBNull.Value, saveBatch))
                                cmdPD.Parameters.AddWithValue("Exp", ParseSafeDate(curExp))
                                cmdPD.Parameters.AddWithValue("Qty", qty)
                                cmdPD.Parameters.AddWithValue("Free", free)
                                cmdPD.Parameters.AddWithValue("tQty", qty + free)
                                cmdPD.Parameters.AddWithValue("MRP", ToD(row.Cells("MRP").Value))
                                cmdPD.Parameters.AddWithValue("PRate", pRate)
                                cmdPD.Parameters.AddWithValue("SRate", ToD(row.Cells("SRate").Value))
                                cmdPD.Parameters.AddWithValue("DisRate", disRate)
                                cmdPD.Parameters.AddWithValue("DisAmt", disAmt)
                                cmdPD.Parameters.AddWithValue("TxRate", saveTxRate)
                                cmdPD.Parameters.AddWithValue("Taxable", curTaxable)
                                cmdPD.Parameters.AddWithValue("SGST", curSGST)
                                cmdPD.Parameters.AddWithValue("CGST", curCGST)
                                cmdPD.Parameters.AddWithValue("IGST", curIGST)
                                cmdPD.Parameters.AddWithValue("NetAmt", ToD(row.Cells("NetAmt").Value))
                                cmdPD.Parameters.AddWithValue("NRate", nRate)

                                cmdPD.Parameters.AddWithValue("PurType", billType)

                                cmdPD.Parameters.AddWithValue("TaxId", taxId)
                                cmdPD.Parameters.AddWithValue("HSN", If(curHSN = "", DBNull.Value, curHSN))
                                cmdPD.Parameters.AddWithValue("BatchId", batchId)
                                cmdPD.Parameters.AddWithValue("AcRate", acRate)
                                cmdPD.Parameters.AddWithValue("TaxAmt", taxAmt)
                                cmdPD.Parameters.AddWithValue("LB", "L")
                                cmdPD.Parameters.AddWithValue("CS", curCS)
                                cmdPD.Parameters.AddWithValue("BoxQty", 1)
                                cmdPD.Parameters.AddWithValue("Scdl", "H")
                                cmdPD.Parameters.AddWithValue("AL", 0)
                                cmdPD.Parameters.AddWithValue("AftDisAmt", afterDisAmt)

                                Dim pDeal As New OleDbParameter("?", OleDbType.VarChar, 50)
                                pDeal.Value = If(curDeal = "", DBNull.Value, curDeal)
                                cmdPD.Parameters.Add(pDeal)

                                Try
                                    cmdPD.ExecuteNonQuery()
                                Catch exPD As Exception
                                    MessageBox.Show("PurchaseDetails insertion failed!" & vbCrLf & "Item: " & itemName & vbCrLf & "Actual Error: " & exPD.Message, "Insertion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Throw exPD
                                End Try
                            End Using
                        End If
                    Next

                    transaction.Commit()
                    MessageBox.Show("Success! Master, Details, and Batch records have been updated perfectly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    transaction.Rollback()
                    MessageBox.Show("Bill save failed due to an error. All entries have been rolled back!" & vbCrLf & vbCrLf & "Main Error Details: " & ex.Message, "Transaction Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using

        Catch exMain As Exception
            MessageBox.Show("Unexpected Error: " & exMain.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnImport.Enabled = True
            btnImport.Text = originalBtnText
        End Try
    End Sub

    ' ==========================================================
    ' 6. FORMAT FUNCTIONS & MENUSTRIP SETUP SELECTOR
    ' ==========================================================
    Private Sub LoadFormats()
        Try
            cmbSavedFormats.Items.Clear()
            Using conn As New OleDbConnection(connString)
                conn.Open()
                Dim rd = New OleDbCommand("SELECT DISTINCT FormatName FROM CsvFormats", conn).ExecuteReader()
                While rd.Read()
                    cmbSavedFormats.Items.Add(rd("FormatName").ToString())
                End While
            End Using
        Catch : End Try
    End Sub

    Private Sub btnSaveFormat_Click(sender As Object, e As EventArgs) Handles btnSaveFormat.Click
        Dim fname = txtFormatName.Text.Trim()
        If fname = "" Then MessageBox.Show("Please enter a format name!", "Validation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
        Try
            Using conn As New OleDbConnection(connString)
                conn.Open()
                Dim del = New OleDbCommand("DELETE FROM CsvFormats WHERE FormatName=?", conn)
                del.Parameters.AddWithValue("@f", fname)
                del.ExecuteNonQuery()

                For Each row As DataGridViewRow In dgvMapping.Rows
                    If row.Cells(1).Value IsNot Nothing AndAlso row.Cells(1).Value.ToString() <> "" Then
                        Dim ins = New OleDbCommand("INSERT INTO CsvFormats (FormatName,FieldName,ColIndex) VALUES (?,?,?)", conn)
                        ins.Parameters.AddWithValue("@f", fname)
                        ins.Parameters.AddWithValue("@n", row.Cells(0).Value.ToString())
                        ins.Parameters.AddWithValue("@i", row.Cells(1).Value.ToString().Trim())
                        ins.ExecuteNonQuery()
                    End If
                Next
            End Using
            MessageBox.Show("Format Saved Successfully!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dgvMapping.Columns(1).ReadOnly = True
            LoadFormats()
        Catch ex As Exception
            MessageBox.Show("Save Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnLoadFormat_Click(sender As Object, e As EventArgs) Handles btnLoadFormat.Click
        If cmbSavedFormats.Text = "" Then Exit Sub
        For Each r As DataGridViewRow In dgvMapping.Rows : r.Cells(1).Value = "" : Next
        Try
            Using conn As New OleDbConnection(connString)
                conn.Open()
                Dim rd = New OleDbCommand("SELECT FieldName,ColIndex FROM CsvFormats WHERE FormatName='" & cmbSavedFormats.Text & "'", conn).ExecuteReader()
                While rd.Read()
                    For Each r As DataGridViewRow In dgvMapping.Rows
                        If r.Cells(0).Value.ToString() = rd("FieldName").ToString() Then
                            r.Cells(1).Value = rd("ColIndex").ToString()
                        End If
                    Next
                End While
            End Using
        Catch : End Try
    End Sub

    Private Sub SelectDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectDatabaseToolStripMenuItem.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Access Database|*.mdb;*.accdb"
        ofd.Title = "Select New Database"

        If ofd.ShowDialog() = DialogResult.OK Then
            Dim dbPath As String = ofd.FileName
            Dim configFilePath As String = Application.StartupPath & "\db_config.txt"

            Try
                IO.File.WriteAllText(configFilePath, dbPath)
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=" & dbPassword & ";Persist Security Info=False;"
                CheckAndCreateFormatTable()
                LoadFormats()

                MessageBox.Show("Success! The database has been changed successfully." & vbCrLf & vbCrLf & "New Path: " & dbPath, "Setup Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error occurred while changing the database: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub EXT_Click(sender As Object, e As EventArgs) Handles EXT.Click
        Me.Close()
    End Sub






    '' ==========================================================
    '' EXPORT: LOAD BILLS IN COMBOBOX
    '' ==========================================================
    'Private Sub LoadBillsForExport()
    '    Try
    '        cmbExportBills.Items.Clear()
    '        ' Dictionary to store Bill Display Name as Key, and SalesMaster ID as Value
    '        Dim billDict As New Dictionary(Of String, Integer)()

    '        Using conn As New OleDbConnection(connString)
    '            conn.Open()
    '           
    '            Dim qLoad = "SELECT SM.Id, SM.InvNo, SM.[InvDate], A.Name AS PartyName " &
    '                        "FROM SaleMaster SM INNER JOIN AcName A ON SM.Partyid = A.id ORDER BY SM.Id DESC"

    '            Using cmd As New OleDbCommand(qLoad, conn)
    '                Using rd = cmd.ExecuteReader()
    '                    While rd.Read()
    '                        ' ComboBox me dikhane ke liye format: "INV101 - Sharma Traders (01/05/2026)"
    '                        'Updated rd("InvDate")
    '                        Dim billDate As String = Convert.ToDateTime(rd("InvDate")).ToString("dd/MM/yyyy")
    '                        Dim displayText As String = rd("InvNo").ToString() & " - " & rd("PartyName").ToString() & " (" & billDate & ")"

    '                        billDict.Add(displayText, Convert.ToInt32(rd("Id")))
    '                        cmbExportBills.Items.Add(displayText)
    '                    End While
    '                End Using
    '            End Using
    '        End Using

    '        ' Store the dictionary in ComboBox Tag so we can get the ID later
    '        cmbExportBills.Tag = billDict

    '        If cmbExportBills.Items.Count > 0 Then cmbExportBills.SelectedIndex = 0

    '    Catch ex As Exception
    '        MessageBox.Show("Bill load karne me error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    '' ==========================================================
    ''  EXPORT: BACKGROUND CSV GENERATOR (NO DGV NEEDED)
    '' ==========================================================
    'Private Sub btnGenerateCSV_Click(sender As Object, e As EventArgs)
    '    If cmbExportBills.SelectedItem Is Nothing Then
    '        MessageBox.Show("Pehle list me se ek Bill select karo!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '        Return
    '    End If

    '    ' Get the SalesMaster ID from the ComboBox Tag Dictionary
    '    Dim billDict As Dictionary(Of String, Integer) = CType(cmbExportBills.Tag, Dictionary(Of String, Integer))
    '    Dim selectedBillId As Integer = billDict(cmbExportBills.SelectedItem.ToString())

    '    Dim sfd As New SaveFileDialog()
    '    sfd.Filter = "CSV File|*.csv"
    '    sfd.Title = "Save Exported Bill as CSV"
    '    sfd.FileName = "ExportedBill_" & selectedBillId & ".csv"

    '    If sfd.ShowDialog() = DialogResult.OK Then
    '        Try
    '            Using conn As New OleDbConnection(connString)
    '                conn.Open()

    '                ' Updated SM.[InvDate] and proper JOIN syntax for MS Access
    '                Dim qExport = "SELECT SM.InvNo, SM.[InvDate], A.Name AS PartyName, P.Name AS ItemName, " &
    '                              "SD.Batch, SD.Exp, SD.QTY, SD.free, SD.deal, SD.MRP, SD.Rate, " &
    '                              "SD.DisRate, SD.DisRs, SD.TxRate, SD.IGST, SD.NetAmt " &
    '                              "FROM (((SaleMaster SM " &
    '                              "INNER JOIN SalesDetails SD ON SM.Id = SD.InvId) " &
    '                              "INNER JOIN AcName A ON SM.Partyid = A.id) " &
    '                              "INNER JOIN ProMaster P ON SD.Pid = P.id) " &
    '                              "WHERE SM.Id = ?"

    '                Using cmd As New OleDbCommand(qExport, conn)
    '                    cmd.Parameters.AddWithValue("?", selectedBillId)

    '                    Using rd = cmd.ExecuteReader()
    '                        Using writer As New IO.StreamWriter(sfd.FileName)
    '                            ' Standard Header Likhna (Jo tere import me match kare)
    '                            writer.WriteLine("Invoice No,Date,Company Name,Item Name,Batch No,Expiry,Qty,Free,Deal,MRP,PRate,DisRate,DisAmt,CGST Percentage,SGST Percentage,IGST Percentage,NetAmt")

    '                            While rd.Read()
    '                                ' Tax calculation logic
    '                                Dim txRate As Double = If(IsDBNull(rd("TxRate")), 0, Convert.ToDouble(rd("TxRate")))
    '                                Dim igstAmt As Double = If(IsDBNull(rd("IGST")), 0, Convert.ToDouble(rd("IGST")))

    '                                Dim cper As Double = 0, sper As Double = 0, iper As Double = 0
    '                                If igstAmt > 0 Then
    '                                    iper = txRate
    '                                Else
    '                                    cper = txRate / 2
    '                                    sper = txRate / 2
    '                                End If

    '                                ' Safe String Conversions
    '                                'Updated rd("InvDate")
    '                                Dim csvDate As String = Convert.ToDateTime(rd("InvDate")).ToString("dd/MM/yyyy")
    '                                Dim party As String = rd("PartyName").ToString().Replace(",", " ") ' Remove commas to avoid CSV break
    '                                Dim item As String = rd("ItemName").ToString().Replace(",", " ")
    '                                Dim deal As String = rd("deal").ToString()
    '                                Dim exp As String = rd("Exp").ToString()
    '                                Dim batch As String = rd("Batch").ToString()

    '                                ' CSV Line Format
    '                                Dim line As String = String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}",
    '                                    rd("InvNo"), csvDate, party, item, batch, exp,
    '                                    rd("QTY"), rd("free"), deal, rd("MRP"), rd("Rate"),
    '                                    rd("DisRate"), rd("DisRs"), cper, sper, iper, rd("NetAmt"))

    '                                writer.WriteLine(line)
    '                            End While
    '                        End Using
    '                    End Using
    '                End Using
    '            End Using

    '            MessageBox.Show("CSV File Successfully Generate ho gayi hai! Ab tu ise kisi ko bhi bhej sakta hai.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

    '        Catch ex As Exception
    '            MessageBox.Show("Export Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    End If
    'End Sub
End Class

' ==========================================================
' CUSTOM SCROLLABLE MESSAGE BOX
' ==========================================================
Public Class ScrollMessageBox
    Public Shared Function Show(msg As String, title As String, buttons As MessageBoxButtons) As DialogResult
        Dim frm As New Form()
        frm.Text = title
        frm.Size = New Size(550, 400) ' Perfect size choti screens ke liye
        frm.StartPosition = FormStartPosition.CenterScreen
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False
        frm.MinimizeBox = False
        frm.BackColor = Color.WhiteSmoke

        ' Scrollable TextBox for Message
        Dim txtBox As New TextBox()
        txtBox.Multiline = True
        txtBox.ReadOnly = True
        txtBox.ScrollBars = ScrollBars.Vertical ' Yahan Scroll aayega!
        txtBox.Text = msg
        txtBox.Font = New Font("Arial", 11)
        txtBox.Location = New Point(15, 15)
        txtBox.Size = New Size(500, 280)
        txtBox.BackColor = Color.White
        frm.Controls.Add(txtBox)

        ' Buttons setup
        If buttons = MessageBoxButtons.YesNo Then
            Dim btnYes As New Button()
            btnYes.Text = "Yes"
            btnYes.DialogResult = DialogResult.Yes
            btnYes.Location = New Point(295, 310)
            btnYes.Size = New Size(100, 35)
            btnYes.BackColor = Color.LightGreen
            btnYes.Font = New Font("Arial", 10, FontStyle.Bold)
            frm.Controls.Add(btnYes)

            Dim btnNo As New Button()
            btnNo.Text = "No"
            btnNo.DialogResult = DialogResult.No
            btnNo.Location = New Point(415, 310)
            btnNo.Size = New Size(100, 35)
            btnNo.BackColor = Color.LightCoral
            btnNo.Font = New Font("Arial", 10, FontStyle.Bold)
            frm.Controls.Add(btnNo)

            frm.AcceptButton = btnYes
            frm.CancelButton = btnNo
        Else
            Dim btnOk As New Button()
            btnOk.Text = "OK"
            btnOk.DialogResult = DialogResult.OK
            btnOk.Location = New Point(415, 310)
            btnOk.Size = New Size(100, 35)
            btnOk.BackColor = Color.LightSkyBlue
            btnOk.Font = New Font("Arial", 10, FontStyle.Bold)
            frm.Controls.Add(btnOk)
            frm.AcceptButton = btnOk
        End If

        Return frm.ShowDialog()
    End Function
End Class