Option Explicit On
Imports System.Diagnostics
Imports System.Net
Imports System.IO
Imports System.DirectoryServices
Imports System.Text

Public Class Form1
    Public clsProcess As Process
    Public username As String = Environment.UserName.ToString & "@tcicollege.edu"
    Dim fs As New FileStream("ADSIUsersAndTheirGroupsList.txt", FileMode.Create)
    Dim s As New StreamWriter(fs)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetLDAPUsers("")
    End Sub

    Public Function IsProcessRunning(ByVal name As String) As Boolean
        For Each Me.clsProcess In Process.GetProcesses()

            If Me.clsProcess.ProcessName.StartsWith(name) Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Sub GetClient()
        If IsProcessRunning("GoverLAN") = True Then
            MsgBox("Govelan is open")

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If IsProcessRunning("GoverLAN") = False Then
            Try
                Dim info As New ProcessStartInfo("C:\Program Files (x86)\GoverLAN v7\GoverLAN.exe")
                info.UseShellExecute = False
                info.RedirectStandardError = True
                info.RedirectStandardInput = True
                info.RedirectStandardOutput = True
                info.CreateNoWindow = True
                info.ErrorDialog = False
                info.WindowStyle = ProcessWindowStyle.Normal
                Dim process__1 As Process = Process.Start(info)
                createTicket()
            Catch ex1 As Exception
                Try
                    Dim info As New ProcessStartInfo("C:\Program Files\GoverLAN v7\GoverLAN.exe")
                    info.UseShellExecute = False
                    info.RedirectStandardError = True
                    info.RedirectStandardInput = True
                    info.RedirectStandardOutput = True
                    info.CreateNoWindow = True
                    info.ErrorDialog = False
                    info.WindowStyle = ProcessWindowStyle.Normal
                    Dim process__1 As Process = Process.Start(info)
                    createTicket()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Try

        End If
    End Sub

    Public Sub createTicket()
        Dim pageurl, requester As String
        requester = ComboBox1.SelectedItem.ToString
        Try
            pageurl = "http://admstatsrvr/GoveZen/GoverZen.php"

            PHP(pageurl, requester)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Public Sub PHP(ByVal pageurl As String, ByVal requester As String)

        Dim wc As New Net.WebClient 'the webclient
        Dim bt() As Byte 'the returned bytes
        Dim html As String 'the returned HTML text

        'setup fields to be sent
        Dim fields As New Specialized.NameValueCollection
        fields.Add("name", username) ' Assigned User
        fields.Add("subject", "Goverlan Ticket") 'subject
        fields.Add("problem", "A ticket was created from GoverLan by " & username) 'Description
        fields.Add("requester", requester)


        bt = wc.UploadValues(pageurl, fields) 'send fields and retrieve response
        html = System.Text.Encoding.ASCII.GetString(bt) 'convert to text
    End Sub

    Public Sub GetLDAPUsers(ByVal Query As String)
        Dim counter As Int16 = 0
        Dim searcher As New DirectorySearcher("")
        Try
            Cursor.Current = Cursors.WaitCursor
            If Query.Trim.Length = 0 Then
                s.WriteLine("search Filter =" + "(&(!(userAccountControl:1.2.840.113556.1.4.803:=2))(objectCategory=user))")
                searcher.Filter = "(&(!(userAccountControl:1.2.840.113556.1.4.803:=2))(objectCategory=user))"
            Else
                s.WriteLine("search Filter =" + "(&(!(userAccountControl:1.2.840.113556.1.4.803:=2))(objectCategory=user)(" + Query + "))")
                searcher.Filter = "(&(!(userAccountControl:1.2.840.113556.1.4.803:=2))(objectCategory=user)(" + Query + "))"
            End If
            searcher.SearchScope = SearchScope.Subtree

            s.WriteLine("")

            Dim FirstName, SurName, Email, UserName As String

            For Each result As SearchResult In searcher.FindAll()
                FirstName = ""
                SurName = ""
                Email = ""
                UserName = ""

                counter = counter + 1

                If Not (IsNothing(result)) Then
                    Dim myResultPropColl As ResultPropertyCollection
                    myResultPropColl = result.Properties

                    Dim myKey As String
                    For Each myKey In myResultPropColl.PropertyNames
                        Select Case myKey
                            Case "samaccountname"
                                Try
                                    UserName = myResultPropColl(myKey)(0)

                                Catch ex As Exception
                                    UserName = ""

                                End Try
                            Case "sn"
                                Try
                                    SurName = myResultPropColl(myKey)(0)

                                Catch ex As Exception
                                    SurName = ""

                                End Try
                            Case "givenname"
                                Try
                                    FirstName = myResultPropColl(myKey)(0)

                                Catch ex As Exception
                                    FirstName = ""

                                End Try
                            Case "mail"
                                Try
                                    Email = myResultPropColl(myKey)(0)

                                Catch ex As Exception
                                    Email = ""

                                End Try
                        End Select

                    Next
                End If
                If Email <> "" Then
                    ComboBox1.Items.Add(Email)
                End If
            Next
           
        Catch ex As Exception
            MessageBox.Show(ex.Message, " Error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'This will give a description of the error.
            Cursor.Current = Cursors.Default
            Exit Sub
        Finally
         
        End Try

    End Sub

    



End Class
