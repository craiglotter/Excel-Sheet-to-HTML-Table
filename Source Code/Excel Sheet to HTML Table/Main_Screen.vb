Imports System.IO
Imports System.Web.Mail
Imports Microsoft.Office.Interop



Public Class Main_Screen

    Private busyworking As Boolean = False
    Private AutoUpdate As Boolean = False

    Private SelectedWorksheet As String = ""
    Private SelectedRange As String = ""
    Private SelectedExcelDoc As String = ""
    Private SelectedHTMLDoc As String = ""

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.ToString
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            End If
            StatusLabel.Text = "Error Reported"
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error." & vbCrLf & vbCrLf & exc.ToString, MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Handler(ByVal message As String)
        Try
            Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            dir = Nothing
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & message)
            filewriter.WriteLine("")
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
            StatusLabel.Text = "Activity Logged"
        Catch ex As Exception
            Error_Handler(ex, "Activity Handler")
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.Text = My.Application.Info.ProductName & " (" & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ")"
            loadSettings()
            StatusLabel.Text = "Application Loaded"
        Catch ex As Exception
            Error_Handler(ex, "Application Loading")
        End Try
    End Sub

    Private Sub loadSettings()
        Try

            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then
                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)
                        If lineread.StartsWith("XLS_SourceFile=") Then
                            If My.Computer.FileSystem.FileExists(variablevalue) = True Then
                                OpenFileDialog1.FileName = variablevalue
                            End If
                        End If
                        If lineread.StartsWith("SelectedWorksheet=") Then
                            SelectedWorksheet = variablevalue
                        End If
                        If lineread.StartsWith("SelectedRange=") Then
                            SelectedRange = variablevalue
                        End If
                        If lineread.StartsWith("SelectedExcelDoc=") Then
                            SelectedExcelDoc = variablevalue
                        End If
                        If lineread.StartsWith("SelectedHTMLDoc=") Then
                            SelectedHTMLDoc = variablevalue
                        End If
                    End If
                End While
                reader.Close()
                reader = Nothing

                If SelectedRange = "" Then
                    SelectedRange = "A1:C10"
                End If

            End If
            StatusLabel.Text = "Application Settings Loaded"
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            Dim writer As StreamWriter = New StreamWriter(configfile, False)
            writer.WriteLine("XLS_SourceFile=" & OpenFileDialog1.FileName)
            writer.WriteLine("SelectedWorksheet=" & SelectedWorksheet)
            writer.WriteLine("SelectedRange=" & SelectedRange)
            writer.WriteLine("SelectedExcelDoc=" & SelectedExcelDoc)
            writer.WriteLine("SelectedHTMLDoc=" & SelectedHTMLDoc)
            writer.Flush()
            writer.Close()
            writer = Nothing
            StatusLabel.Text = "Application Settings Saved"
        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub

    Private Sub Main_Screen_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            SaveSettings()
            If AutoUpdate = True Then
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\AutoUpdate.exe").Replace("\\", "\")) = True Then
                    Dim startinfo As ProcessStartInfo = New ProcessStartInfo
                    startinfo.FileName = (Application.StartupPath & "\AutoUpdate.exe").Replace("\\", "\")
                    startinfo.Arguments = "force"
                    startinfo.CreateNoWindow = False
                    Process.Start(startinfo)
                End If
            End If
            StatusLabel.Text = "Application Shutting Down"
        Catch ex As Exception
            Error_Handler(ex, "Closing Application")
        End Try
    End Sub
  

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem1.Click
        Try
            HelpBox1.ShowDialog()
            StatusLabel.Text = "Help Dialog Viewed"
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub

    Private Sub AutoUpdateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoUpdateToolStripMenuItem.Click
        Try
            StatusLabel.Text = "AutoUpdate Requested"
            AutoUpdate = True
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "AutoUpdate")
        End Try
    End Sub

    Private Sub AboutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem1.Click
        Try
            AboutBox1.ShowDialog()
            StatusLabel.Text = "About Dialog Viewed"
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub Control_Enabler(ByVal IsEnabled As Boolean)
        Try
            Select Case IsEnabled
                Case True
                    Button1.Enabled = True
                    MenuStrip1.Enabled = True
                    Me.ControlBox = True
                    ProgressBar1.Enabled = False
                Case False
                    Button1.Enabled = False
                    MenuStrip1.Enabled = False
                    Me.ControlBox = False
                    ProgressBar1.Enabled = True
            End Select
            StatusLabel.Text = "Control Enabler Run"
        Catch ex As Exception
            Error_Handler(ex, "Control Enabler")
        End Try
    End Sub


   

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Try
                Dim excelApplication As Excel.ApplicationClass = New Excel.ApplicationClass()
                Dim excelWorkbook As Excel.Workbook = Nothing
                Try
                    excelWorkbook = excelApplication.Workbooks.Open(SelectedExcelDoc)
                    Dim excelSheet As Excel.Worksheet = excelWorkbook.Worksheets(SelectedWorksheet)
                    excelSheet.Activate()
                    Dim currentRange As Excel.Range = excelSheet.Range(SelectedRange.Split(":")(0), SelectedRange.Split(":")(1))
                    currentRange.Activate()

                    Dim cell As Excel.Range

                    Dim currentrow As Integer = 0
                    Dim writer As StreamWriter = New StreamWriter(SelectedHTMLDoc, False)
                    writer.WriteLine("<html>" & vbCrLf & "<body>")
                    writer.WriteLine("<table  border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""3"" style=""border-collapse: collapse"" bordercolor=""#C0C0C0"">")
                    StatusLabel.Text = "Running Extractions"
                    Dim counter As Integer = 0
                    For Each cell In currentRange.Cells
                        If cell.Row <> currentrow Then
                            If currentrow <> 0 Then
                                writer.WriteLine("</tr>" & vbCrLf & "<tr>")
                            Else
                                writer.WriteLine("<tr>")
                            End If
                            currentrow = cell.Row
                        End If
                        If cell.FormulaR1C1.ToString.Length > 0 Then
                            writer.WriteLine("<td align=""left"" valign=""top"">" & cell.FormulaR1C1.ToString & "</td>")
                        Else
                            writer.WriteLine("<td align=""left"" valign=""top"">&nbsp;</td>")
                        End If
                        counter = counter + 1
                        ProgressBar1.Value = Math.Round(((counter / currentRange.Cells.Count) * 100), 0)
                    Next
                    writer.WriteLine("</tr>")
                    writer.WriteLine("</table>")
                    writer.WriteLine("</body>" & vbCrLf & "</html>")
                    writer.Flush()
                    writer.Close()
                    writer = Nothing


                    e.Result = "Success"

                Catch ex As Exception
                    Error_Handler(ex, "Worksheet Conversion")
                    e.Cancel = True
                Finally
                    ' Close the workbook object.
                    If Not excelWorkbook Is Nothing Then
                        excelWorkbook.Close(False)
                        excelWorkbook = Nothing
                    End If

                    ' Quit Excel and release the ApplicationClass object.
                    If Not excelApplication Is Nothing Then
                        excelApplication.Quit()
                        excelApplication = Nothing
                    End If

                End Try
            Catch ex As Exception
                e.Cancel = True
                Error_Handler(ex, "Worksheet Conversion")
            End Try

        Catch ex As Exception
            e.Cancel = True
            Error_Handler(ex, "HTML Extraction Operation")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Control_Enabler(True)
            If e.Cancelled = False And e.Error Is Nothing Then
                If e.Result = "Success" Then
                    StatusLabel.Text = "HTML Extraction Complete"
                    If My.Computer.FileSystem.FileExists(XLS_GeneratedFile.Text) = True Then
                        Process.Start(XLS_GeneratedFile.Text)
                    Else
                        StatusLabel.Text = "HTML Extraction Failed"
                    End If
                End If
            Else
                StatusLabel.Text = "HTML Extraction Failed"
            End If
            busyworking = False
        Catch ex As Exception
            Error_Handler(ex, "HTML Extraction Complete")
        End Try
    End Sub

    Private Sub runworker()
        Try
            If busyworking = False Then
                ProgressBar1.Value = 0
                XLS_GeneratedFile.Text = ""
                XLS_Range.Text = ""
                XLS_SourceFile.Text = ""


                If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim excelApplication As Excel.ApplicationClass = New Excel.ApplicationClass()
                    Dim excelWorkbook As Excel.Workbook = Nothing
                    Dim exceloptions1 As ExcelOptions = New ExcelOptions
                    Me.Refresh()
                    excelWorkbook = excelApplication.Workbooks.Open(OpenFileDialog1.FileName)
                    Me.Refresh()
                    For Each sht As Excel.Worksheet In excelWorkbook.Worksheets
                        exceloptions1.SelectedWorksheet.Items.Add(sht.Name)
                    Next
                    If exceloptions1.SelectedWorksheet.Items.Count > 0 Then
                        exceloptions1.SelectedWorksheet.SelectedIndex = 0
                    End If
                    exceloptions1.SelectedRange.Text = SelectedRange
                    If Not excelWorkbook Is Nothing Then
                        excelWorkbook.Close(False)
                        excelWorkbook = Nothing
                    End If
                    If Not excelApplication Is Nothing Then
                        excelApplication.Quit()
                        excelApplication = Nothing
                    End If


                    If exceloptions1.ShowDialog = Windows.Forms.DialogResult.OK Then
                        busyworking = True
                        Control_Enabler(False)
                        StatusLabel.Text = "Initializing Extraction Operation"
                        SelectedExcelDoc = OpenFileDialog1.FileName
                        SelectedHTMLDoc = SelectedExcelDoc & ".htm"
                        SelectedRange = exceloptions1.SelectedRange.Text
                        SelectedWorksheet = exceloptions1.SelectedWorksheet.Items(exceloptions1.SelectedWorksheet.SelectedIndex)
                        XLS_SourceFile.Text = SelectedExcelDoc
                        XLS_Range.Text = SelectedRange & " (" & SelectedWorksheet & ")"
                        XLS_GeneratedFile.Text = SelectedHTMLDoc
                        BackgroundWorker1.RunWorkerAsync()
                    End If
                    exceloptions1 = Nothing
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Run Worker")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            runworker()
        Catch ex As Exception
            Error_Handler(ex, "Stop Timer Click")
        End Try
    End Sub

 
End Class
