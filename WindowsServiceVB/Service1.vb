
Imports System.IO
Imports System.Threading
Imports System.Configuration
Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class Service1

    Protected Overrides Sub OnStart(ByVal args() As String)
        Me.WriteToFile("Simple Service started at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Me.ScheduleService()
    End Sub

    Protected Overrides Sub OnStop()
        Me.WriteToFile("Simple Service stopped at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Me.Schedular.Dispose()
    End Sub

    Private Schedular As Timer

    Public Sub ScheduleService()
        Try
            Schedular = New Timer(New TimerCallback(AddressOf SchedularCallback))
            Dim mode As String = ConfigurationManager.AppSettings("Mode").ToUpper()
            Me.WriteToFile((Convert.ToString("Simple Service Mode: ") & mode) + " {0}")

            'Set the Default Time.
            Dim scheduledTime As DateTime = DateTime.MinValue

            If mode = "DAILY" Then
                'Get the Scheduled Time from AppSettings.
                scheduledTime = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings("ScheduledTime"))
                If DateTime.Now > scheduledTime Then
                    'If Scheduled Time is passed set Schedule for the next day.
                    scheduledTime = scheduledTime.AddDays(1)
                End If
            End If

            If mode.ToUpper() = "INTERVAL" Then
                'Get the Interval in Minutes from AppSettings.
                Dim intervalMinutes As Integer = Convert.ToInt32(ConfigurationManager.AppSettings("IntervalMinutes"))

                'Set the Scheduled Time by adding the Interval to Current Time.
                scheduledTime = DateTime.Now.AddMinutes(intervalMinutes)
                If DateTime.Now > scheduledTime Then
                    'If Scheduled Time is passed set Schedule for the next Interval.
                    scheduledTime = scheduledTime.AddMinutes(intervalMinutes)
                End If
            End If

            Dim timeSpan As TimeSpan = scheduledTime.Subtract(DateTime.Now)
            Dim schedule As String = String.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds)

            Me.WriteToFile((Convert.ToString("Simple Service scheduled to run after: ") & schedule) + " {0}")

            'Get the difference in Minutes between the Scheduled and Current Time.
            Dim dueTime As Integer = Convert.ToInt32(timeSpan.TotalMilliseconds)

            'Change the Timer's Due Time.
            Schedular.Change(dueTime, Timeout.Infinite)
        Catch ex As Exception
            WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace)

            'Stop the Windows Service.
            Using serviceController As New System.ServiceProcess.ServiceController("SimpleService")
                serviceController.[Stop]()
            End Using
        End Try
    End Sub

    Private Sub SchedularCallback(e As Object)
        Try
            Dim dt As New DataTable()
            Dim query As String = "SELECT Name, Email FROM Students WHERE DATEPART(DAY, BirthDate) = @Day AND DATEPART(MONTH, BirthDate) = @Month"
            Dim constr As String = ConfigurationManager.ConnectionStrings("constr").ConnectionString
            Using con As New SqlConnection(constr)
                Using cmd As New SqlCommand(query)
                    cmd.Connection = con
                    cmd.Parameters.AddWithValue("@Day", DateTime.Today.Day)
                    cmd.Parameters.AddWithValue("@Month", DateTime.Today.Month)
                    Using sda As New SqlDataAdapter(cmd)
                        sda.Fill(dt)
                    End Using
                End Using
            End Using
            For Each row As DataRow In dt.Rows
                Dim name As String = row("Name").ToString()
                Dim email As String = row("Email").ToString()
                WriteToFile("Trying to send email to: " & name & " " & email)

                Using mm As New MailMessage("sender@gmail.com", email)
                    mm.Subject = "Birthday Greetings"
                    mm.Body = String.Format("<b>Happy Birthday </b>{0}<br /><br />Many happy returns of the day.", name)

                    mm.IsBodyHtml = True
                    Dim smtp As New SmtpClient()
                    smtp.Host = "smtp.gmail.com"
                    smtp.EnableSsl = True
                    Dim credentials As New System.Net.NetworkCredential()
                    credentials.UserName = "sender@gmail.com"
                    credentials.Password = "<Password>"
                    smtp.UseDefaultCredentials = True
                    smtp.Credentials = credentials
                    smtp.Port = 587
                    smtp.Send(mm)
                    WriteToFile("Email sent successfully to: " & name & " " & email)
                End Using
            Next
            Me.ScheduleService()
        Catch ex As Exception
            WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace)

            'Stop the Windows Service.
            Using serviceController As New System.ServiceProcess.ServiceController("SimpleService")
                serviceController.[Stop]()
            End Using
        End Try
    End Sub

    Private Sub WriteToFile(text As String)
        Dim path As String = "C:\ServiceLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub
End Class
