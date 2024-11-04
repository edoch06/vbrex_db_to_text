Imports System.Data.Common
Imports System.Data.Odbc
Imports System.IO
Imports System.Text
Imports System.Data.SqlClient
Imports System.Configuration

Public Class Form1

    Dim con As Odbc.OdbcDataAdapter
    Dim dAdapt, da As Odbc.OdbcDataAdapter
    Dim dSet As DataSet
    Dim dSet1 As DataSet
    Dim ds As New DataSet
    Dim a As Integer = 0
    Dim dBind As New BindingSource
    Dim dBind1 As New BindingSource
    Dim dBind2 As New BindingSource

    Private Sub Connections()
        Dim connectionString As String

        Try
            If txtcstring.Text = "" Then

                MsgBox("Connection string empty", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Connection Error")
            Else

                connectionString = txtcstring.Text

                If cmbserver.Text = "SQL Server" Then

                    Dim con1 As New SqlConnection(connectionString)


                    con1.Open()
                    MessageBox.Show("Connection Opened Successfully")
                    lblconn.Visible = True
                    con1.Close()
                 

                ElseIf cmbserver.Text = "MySql Server" Then

                    connectionString = txtcstring.Text
                    Dim con As New Odbc.OdbcConnection(connectionString)


                    con.Open()
                    MessageBox.Show("Connection Opened Successfully")
                    lblconn.Visible = True
                    con.Close()

                End If

            End If
 






        Catch ex As Exception
            'MsgBox(ex.ToString())
            MsgBox("Error Connecting to Database", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Connection Error")
        End Try

    End Sub


    Private Sub extractdata()

        Try


        Catch ex As Exception
            MsgBox(ex.ToString())
            'MsgBox("No Database connection..." & vbCrLf & vbCrLf & "Contact administrator...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Database Connection Error")
        End Try

    End Sub

    Private Sub loaddata()

        Try
            Dim str As String
            Dim dt As New DataTable
            Dim connectionString As String

            connectionString = txtcstring.Text

            Dim con1 As New SqlConnection(connectionString)

            If RTxtSel.Text = "" Then

                str = "select * from " & txtdbname.Text & ""
                Dim dAdapt As New SqlDataAdapter(str, con1)
                dSet = New DataSet
                dAdapt.Fill(dSet, "tabledata")
                dAdapt.Fill(dt)
                dBind2.DataSource = dSet


                dBind2.DataMember = dSet.Tables(0).ToString()

                DataGridView3.DataSource = dBind2

            Else

                str = "select " & RTxtSel.Text & " from " & txtdbname.Text & ""
                Dim dAdapt As New SqlDataAdapter(str, con1)
                dSet = New DataSet
                dAdapt.Fill(dSet, "tabledata")
                dAdapt.Fill(dt)
                dBind2.DataSource = dSet


                dBind2.DataMember = dSet.Tables(0).ToString()

                DataGridView3.DataSource = dBind2

            End If






        Catch ex As Exception
            'MsgBox(ex.ToString())
            MsgBox("No Database Connection Selected", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Connection Error")
        End Try

    End Sub

    Private Sub loadcolumns()

        Try
            Dim str1 As String
            Dim dt1 As New DataTable
            Dim connectionString As String

            connectionString = txtcstring.Text

            Dim con1 As New SqlConnection(connectionString)

            str1 = "select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS IC where TABLE_NAME= '" & txtdbname.Text & "'"
            Dim dAdapt1 As New SqlDataAdapter(str1, con1)
            dSet1 = New DataSet
            dAdapt1.Fill(dSet1, "dbcolumns")
            dAdapt1.Fill(dt1)
            dBind1.DataSource = dSet1


            dBind1.DataMember = dSet1.Tables(0).ToString()

          

            DataGridView2.AutoGenerateColumns = False
            DataGridView2.ColumnCount = 1
            DataGridView2.Columns(0).Name = "COLUMN_NAME"
            DataGridView2.Columns(0).Width = 280
            DataGridView2.Columns(0).HeaderText = "colums"
            DataGridView2.Columns(0).DataPropertyName = "COLUMN_NAME"
            DataGridView2.Columns(0).ReadOnly = True



            'Add a CheckBox Column to the DataGridView at the first position.
            Dim chk As New DataGridViewCheckBoxColumn()
            DataGridView2.Columns.Add(chk)
            chk.Width = 57
            chk.HeaderText = "Select"
            chk.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            chk.Name = "chk"
            chk.TrueValue = CheckState.Checked
            chk.FalseValue = CheckState.Unchecked
            chk.ReadOnly = False

            DataGridView2.DataSource = dBind1
            DataGridView2.DataSource = dt1


        Catch ex As Exception
            'MsgBox(ex.ToString())
            'MsgBox("No Database connection..." & vbCrLf & vbCrLf & "Contact administrator...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Database Connection Error")
        End Try

    End Sub

    Private Sub loadsqltable()

        Try
            Dim str As String
            Dim dt As New DataTable
            Dim connectionString As String

            connectionString = txtcstring.Text

            Dim con1 As New SqlConnection(connectionString)

            str = "SELECT name FROM sys.tables"
            Dim dAdapt As New SqlDataAdapter(str, con1)
            dSet = New DataSet
            dAdapt.Fill(dSet, "dbtables")
            dAdapt.Fill(dt)
            dBind.DataSource = dSet


            dBind.DataMember = dSet.Tables(0).ToString()

            DataGridView1.DataSource = dBind


            Call Me.loadcolumns()

        Catch ex As Exception
            MsgBox(ex.ToString())
            'MsgBox("No Database connection..." & vbCrLf & vbCrLf & "Contact administrator...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Database Connection Error")
        End Try

    End Sub
    Private Sub writefile()

        Try
            Dim path As String = "C:\SLDataExtracts.txt"
            Dim fs As FileStream = File.Create(path)
            Dim info As Byte() = New UTF8Encoding(True).GetBytes("")
            fs.Write(info, 0, info.Length)
            fs.Close()

        Catch ex As Exception
            MsgBox(ex.ToString())
            'Call createdir()
        End Try

    End Sub
    Private Sub writefilecol()

        Try
            Dim path As String = "C:\SLDataExtractsCol.txt"
            Dim fs As FileStream = File.Create(path)
            Dim info As Byte() = New UTF8Encoding(True).GetBytes(RTxtSel.Text)
            fs.Write(info, 0, info.Length)
            fs.Close()

        Catch ex As Exception
            MsgBox(ex.ToString())
            'Call createdir()
        End Try

    End Sub
    Private Sub readcol()

        Try

            Using sr As New StreamReader("C:\SLDataExtractsCol.txt")
                Dim line As String
                line = sr.ReadToEnd()
                'txttillNo3.Text = line
                RTxtSel.Text = line
            End Using

        Catch ex As Exception
            Call Me.writefilecol()
            'MsgBox(ex.ToString())
        End Try

    End Sub

    Private Sub txtinitial()

        Try

            lblconn.Visible = False
            Call readcol()

        Catch ex As Exception
            MsgBox(ex.ToString())
            'Call createdir()
        End Try

    End Sub


    Public Function ReadLine(ByVal lineNumber As Integer, ByVal lines As List(Of String)) As String
        Return lines(lineNumber - 1)
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call txtinitial()
    End Sub

    Private Sub btnCStrings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCStrings.Click
        txtcstring.Text = ""
        Process.Start("C:\MyConnections.udl")
    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub btnExString_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExString.Click
        Try
            Dim connstring As String
            Dim newconnstring As String
            Dim reader As New System.IO.StreamReader("C:\MyConnections.udl")
            Dim allLines As List(Of String) = New List(Of String)
            Do While Not reader.EndOfStream
                allLines.Add(reader.ReadLine())
            Loop
            reader.Close()
            txtcstring.Text = ReadLine(3, allLines)


            connstring = txtcstring.Text
            newconnstring = connstring.Substring(connstring.IndexOf(";"c) + 1)





            If cmbserver.Text = "SQL Server" Then

                 txtcstring.Text = newconnstring


            ElseIf cmbserver.Text = "MySql Server" Then

                txtcstring.Text = newconnstring.Replace("Data Source", "DSN")

            End If



        Catch ex As Exception
            ' Let the user know what went wrong.
            Console.WriteLine("The file could not be read:")
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    Private Sub btnconnectdb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnconnectdb.Click
        Call Connections()

    End Sub

    Private Sub btnsplit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btndistables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndistables.Click
        If lblconn.Visible = True Then
            Call loadsqltable()

        Else
            MsgBox("No Connection", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Connection Error")
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Try
            Dim i As Integer


            If Not DataGridView1.SelectedRows.Count = 0 Then

                i = DataGridView1.CurrentRow.Index
                txtdbname.Text = DataGridView1.Item(0, i).Value


            End If



        Catch ex As Exception

            'MsgBox(ex.ToString())
            'DataGridView1.Rows(0).Selected = True
            DataGridView1.ClearSelection()

        End Try
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        Try
            Dim con As New Odbc.OdbcConnection("DSN=pointofsale")
            Dim con1 As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=Internal_Data;Integrated Security=True;Connect Timeout=30;User Instance=True")
            Dim strMyString As String
            Dim colString As String
            Dim ds3 As New DataSet

            Dim i As Integer


            txtcol.Clear()
            'DataGridViewIn.Item(5, i).Selected = True
            i = DataGridView2.CurrentRow.Index
            strMyString = RTxtSel.Text

            If DataGridView2.Item("chk", i).Selected = True And RTxtSel.Text = "" Then

                txtcol.Text = DataGridView2.Item(0, i).Value
                RTxtSel.Text = txtcol.Text


            Else

                txtcol.Text = ""
                txtcol.Text = DataGridView2.Item(0, i).Value

                colString = "," & txtcol.Text

                If strMyString.Contains(colString) Then

                    strMyString = strMyString.Replace(colString, "")
                    RTxtSel.Text = strMyString
                    Call writefilecol()

                Else
                    RTxtSel.AppendText(colString)
                    Call writefilecol()
                End If



            End If

            'DataGridView1.DataSource = Nothing
            'Call gridPOrder()
            'Call Me.countqnty()




        Catch ex As Exception

            'MsgBox(ex.ToString())
            'DataGridView1.Rows(0).Selected = True
            DataGridView1.ClearSelection()

        End Try
    End Sub


    Private Sub btnclr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclr.Click
        DataGridView2.DataSource = Nothing
        Call Me.loadcolumns()
        RTxtSel.Clear()
        Call writefilecol()
    
    End Sub

    Private Sub btnextract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnextract.Click
        Try


            'Dim result As String = ""
            ''go through all rows
            'For rowNumber As Integer = 0 To DataGridView2.Rows.Count - 1
            '    'this gets just column 0 (the first column)
            '    result &= DataGridView2.Rows(rowNumber).Cells("COLUMN_NAME").Value & vbCrLf
            'Next
            ''write out the string
            'File.WriteAllText("c:\SLDataExtracts.txt", result)

            Call writefile()
            Call loaddata()

            Dim sLine As String = ""
            Dim sLines As String = "" ' added
            For iRow As Integer = 0 To DataGridView3.RowCount - 1
                sLine = ""
                For iCol As Integer = 0 To DataGridView3.ColumnCount - 1
                    ' test header text: Device_LabelsDataGridView.Columns(iCol).HeaderText
                    'If iCol = 0 Or iCol = 1 Then
                    sLine &= DataGridView3(iCol, iRow).Value.ToString & "|"
                    'End If
                Next
                sLines &= sLine.Substring(0, sLine.Length - 1) & vbCrLf
            Next
            My.Computer.FileSystem.WriteAllText("c:\SLDataExtracts.txt", sLines, True)

            MsgBox("Extraction Complete", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Database Extraction")

            Process.Start("C:\SLDataExtracts.txt")


        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Sub

    Private Sub btnconfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call writefile()
    End Sub

    Private Sub txtdbname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtdbname.TextChanged
        DataGridView2.DataSource = Nothing
        Call loadcolumns()
    End Sub
End Class
