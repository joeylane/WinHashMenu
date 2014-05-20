Imports System.Security.Cryptography
Imports System.Security.Principal
Imports System.IO
Imports System.Text

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'check to see if we passed any file paths in the arguments.  if not, run setup.
        If Environment.GetCommandLineArgs.Count > 1 Then
            Try
                Dim myFile As String = Environment.GetCommandLineArgs(1)
                Dim myMD5Hash As String = GetMD5(myFile)
                Dim mySHA1Hash As String = GetSHA1(myFile)
                Dim mySHA256Hash As String = GetSHA256(myFile)
                Label3.Text = myMD5Hash
                Label4.Text = mySHA1Hash
                Label7.Text = mySHA256Hash
                Label5.Text = myFile
            Catch
            End Try
        Else
            'check if we are running as administrator before proceeding.  if not, relaunch with UAC approval first.
            Dim identity = WindowsIdentity.GetCurrent()
            Dim principal = New WindowsPrincipal(identity)
            Dim isElevated As Boolean = principal.IsInRole(WindowsBuiltInRole.Administrator)
            If isElevated = True Then
                'if key already exists, ask user if we want to uninstall
                If Not My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\*\shell\Get file hashes\command", "", Nothing) Is Nothing Then
                    If MessageBox.Show("You have already installed WinHashMenu.  Would you like to uninstall it?  Select 'Yes' to uninstall completely, or 'No' to reinstall the shell extension.", "WinHashMenu", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                        My.Computer.Registry.ClassesRoot.DeleteSubKey("*\shell\Get file hashes\command")
                        My.Computer.Registry.ClassesRoot.DeleteSubKey("*\shell\Get file hashes")
                        MessageBox.Show("The shell extension has been successfully removed.", "WinHashMenu")
                        End
                    Else
                        install()
                    End If
                Else
                    install()
                End If
            Else
                MessageBox.Show("WinHashMenu requires administrator privileges to setup the shell extension.  Select 'Yes' when you see the UAC prompt to continue, or 'No' to abort.", "WinHashMenu Setup")
                'run as admin
                Try
                    Dim myProcess2 As New System.Diagnostics.Process()
                    Dim startInfo2 As New ProcessStartInfo
                    startInfo2.FileName = System.Reflection.Assembly.GetExecutingAssembly().Location
                    startInfo2.Verb = "runas"
                    myProcess2.StartInfo = startInfo2
                    Application.DoEvents()
                    myProcess2.Start()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "WinHashMenu")
                End Try
            End If
                End
        End If
    End Sub

    Function GetMD5(ByVal filePath As String)
        Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
        Dim f As FileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        f = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        md5.ComputeHash(f)
        f.Close()
        Dim hash As Byte() = md5.Hash
        Dim buff As StringBuilder = New StringBuilder
        Dim hashByte As Byte
        For Each hashByte In hash
            buff.Append(String.Format("{0:X2}", hashByte))
        Next
        Dim md5string As String
        md5string = buff.ToString()
        Return md5string
    End Function

    Function GetSHA1(ByVal filePath As String)
        Dim sha1 As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider
        Dim f As FileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        f = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        sha1.ComputeHash(f)
        f.Close()
        Dim hash As Byte() = sha1.Hash
        Dim buff As StringBuilder = New StringBuilder
        Dim hashByte As Byte
        For Each hashByte In hash
            buff.Append(String.Format("{0:X2}", hashByte))
        Next
        Dim md5string As String
        md5string = buff.ToString()
        Return md5string
    End Function

    Function GetSHA256(ByVal filePath As String)
        Dim sha256 As SHA256CryptoServiceProvider = New SHA256CryptoServiceProvider
        Dim f As FileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        f = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        sha256.ComputeHash(f)
        f.Close()
        Dim hash As Byte() = sha256.Hash
        Dim buff As StringBuilder = New StringBuilder
        Dim hashByte As Byte
        For Each hashByte In hash
            buff.Append(String.Format("{0:X2}", hashByte))
        Next
        Dim md5string As String
        md5string = buff.ToString()
        Return md5string
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Computer.Clipboard.SetText(Label3.Text)
        Label3.ForeColor = Color.Blue
        Timer1.Start()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        My.Computer.Clipboard.SetText(Label4.Text)
        Label4.ForeColor = Color.Blue
        Timer1.Start()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        My.Computer.Clipboard.SetText(Label7.Text)
        Label7.ForeColor = Color.Blue
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label3.ForeColor = Color.Black
        Label4.ForeColor = Color.Black
        Label7.ForeColor = Color.Black
        Timer1.Stop()
    End Sub

    Private Sub install()
        MessageBox.Show("WinHashMenu will now install the shell extension.  Note that it will configure itself to run from the current running directory:" & vbNewLine & vbNewLine & System.Reflection.Assembly.GetExecutingAssembly().Location & vbNewLine & vbNewLine & "If you would rather store this executable in a different directory, please move it to a more appropriate location first.  If you decide to relocate this file later, just run it again from the new directory to reconfigure the shell extension.  You may also run this executable again if you wish to uninstall the WinHashMenu shell extension.", "WinHashMenu Setup")
        My.Computer.Registry.ClassesRoot.CreateSubKey("*\shell\Get file hashes\command")
        Dim key = My.Computer.Registry.ClassesRoot.OpenSubKey("*\shell\Get file hashes\command", True)
        key.SetValue("", Chr(34) & System.Reflection.Assembly.GetExecutingAssembly().Location & Chr(34) & " " & Chr(34) & "%1" & Chr(34))
        key.Close()
        MessageBox.Show("Shell extension installed and configured successfully.", "WinHashMenu")
    End Sub

End Class
