Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net

Public Class frmMain

#Region "Declarations"
    'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
#End Region

#Region "ReadOnly Variables"

    'Separator
    ReadOnly Separator As String = ";"

#End Region

#Region "Variables"

    'Engine of the app
    Dim motor As String

    'Source String var
    Dim sourceString As String

    'Matchcolletion for collect all regex matchs
    Dim matches As MatchCollection

    'Var for retrieve result of the first regex search
    Dim firstMatch As String

    'Var for retrieve result of the second regex search
    Dim secondMatch As String

    'Retrieve all cookies on this var
    Dim Cookies As New CookieContainer

    'Get string name for enter again
    Dim cookieString As String

    'Requests are here
    Dim request As System.Net.HttpWebRequest

    'Responses are here
    Dim response As System.Net.HttpWebResponse

    'For search cookie
    Dim cookieName As String

    'For read all content from Source code
    Dim sr As System.IO.StreamReader

    'Var for retrieve result of the third regex search
    Dim downloadLink As String

    'File name where are put all the links
    Dim fileName As String = "prueba-" & DateDiff(DateInterval.Second, #1/1/1970#, Date.Now) & ".txt"

    'Final text
    Dim finalText As String

    'Times
    Dim times As Integer

    'Index pattern
    Dim IndexPattern As Integer

    'Put song size here
    Dim songSize As String

    'Song size limit here, listed in bytes
    Dim songLimitSize As Integer = 300000

    'Determine if the actual song has less bytes than limit
    Dim retry As Boolean = False

    'Set the retries of the current song
    Dim retries As Integer

    'First regex search
    Dim regexSearch As String

    'Second regex search
    Dim secondregexMatch As String

    'Third and last regex search
    Dim regexLink As String

    'Get size
    Dim regexGetSize As String

    Dim id As Integer

#End Region

#Region "Properties"

    '[PROP] Set terms in search
    Private ReadOnly Property Terms As String()
        Get
            Terms = (TextBox1.Text & Separator).Split(Separator)
            Return Terms.Distinct().ToArray()
        End Get
    End Property

    Private ReadOnly Property SearchEngine As Integer
        Get
            Return DirectCast(ComboBox1.SelectedIndex, SearchEngines)
        End Get
    End Property

#End Region

#Region "Enums"

    Private Enum SearchEngines As Integer
        mrtzcmp3 = 0
        mp3skull = 1
    End Enum

#End Region

#Region "Funtions"

    '[FUNC] Search in all source code and extract download link
    Private Sub mrtzcmp3_Search(ByVal songName As String)

        Try

            Update_ToolStrip_Progress()
            id = IndexPattern

            'Put varaibles values
            cookieName = "haras"
            motor = "http://mrtzcmp3.net/"
            regexSearch = "D\?.+? _"
            regexLink = "http://m.mrtzcmp3.net/get.php.+?"""
            secondregexMatch = "MRTZC\?\w+"
            regexGetSize = "size=\d+"


            'Get first source
            request = CType(HttpWebRequest.Create(motor & songName & "_1s.html"), HttpWebRequest)
            request.CookieContainer = Cookies 'Request Cookies
            response = CType(request.GetResponse(), HttpWebResponse)

            'DO multiple things with the cookies
            For Each cookieValue As Cookie In response.Cookies
                If cookieValue.ToString.Substring(0, cookieValue.ToString.IndexOf("=")) = cookieName Then
                    cookieString = cookieValue.ToString.Replace(cookieName & "=", "") 'Get haras cookie value
                End If
                Cookies.Add(cookieValue) 'Add cookies to the container for use it again
            Next

            'Read Source String
            sr = New System.IO.StreamReader(response.GetResponseStream())
            sourceString = sr.ReadToEnd()

            'Check the result
            matches = Regex.Matches(sourceString, regexSearch)

            'If results = 0 stop sub
            If matches.Count = 0 Then
                'MsgBox("Se encontraron 0 resultados de esta canción.", MsgBoxStyle.Information, "Información")
                File.AppendAllText(fileName, String.Format("#{0} [{1}]: Se encontraron 0 resultados de esta canción." & Environment.NewLine, id, songName))
                Exit Sub
            End If

            'Get first song link
            firstMatch = matches.Item(0).Value

Retry:

            If retry Then
                If retries >= matches.Count Then
                    File.AppendAllText(fileName, String.Format("#{0} [{1}]: Todas las caciones de este artista no superaron el limite permanente de Bytes para ser aceptada." & Environment.NewLine, id, songName))
                    Exit Sub
                End If
                firstMatch = matches.Item(retries).Value
            End If

            'Enter with Fake Cookie
            request = CType(HttpWebRequest.Create(motor & firstMatch & cookieString), HttpWebRequest)
            request.CookieContainer = Cookies
            response = CType(request.GetResponse(), HttpWebResponse)

            'Read source String
            sr = New System.IO.StreamReader(response.GetResponseStream())
            sourceString = sr.ReadToEnd()

            'Get Download Link
            matches = Regex.Matches(sourceString, secondregexMatch)

            'If results = 0 stop sub
            If matches.Count = 0 Then
                'MsgBox("Hubo un error al obtener el link de descarga.", MsgBoxStyle.Information, "Información")
                File.AppendAllText(fileName, String.Format("#{0} [{1}]: Hubo un error al obtener el link de descarga." & Environment.NewLine, id, songName))
                Exit Sub
            End If

            'Get first song link
            secondMatch = matches.Item(0).Value

            'Final source Code get
            request = CType(HttpWebRequest.Create(motor & secondMatch), HttpWebRequest)
            request.CookieContainer = Cookies 'Still usign fake cookies
            response = CType(request.GetResponse(), HttpWebResponse)

            'Read source String
            sr = New System.IO.StreamReader(response.GetResponseStream())
            sourceString = sr.ReadToEnd()

            'Get direct download Link
            matches = Regex.Matches(sourceString, regexLink)

            'If results = 0 stop sub
            If matches.Count = 0 Then
                'MsgBox("Hubo un error al obtener el link de descarga directa.", MsgBoxStyle.Information, "Información")
                File.AppendAllText(fileName, String.Format("#{0} [{1}]: Hubo un error al obtener el link de descarga directa." & Environment.NewLine, id, songName))
                Exit Sub
            End If

            'Get first song link
            downloadLink = matches.Item(0).Value
            downloadLink = downloadLink.Replace("""", "")

            'Getting size
            matches = Regex.Matches(downloadLink, regexGetSize)
            songSize = matches.Item(0).Value.Replace("size=", "")

            'Check song Size
            If songSize < songLimitSize Then

                'Get first source
                request = CType(HttpWebRequest.Create(motor & songName & "_1s.html"), HttpWebRequest)
                request.CookieContainer = Cookies 'Request Cookies
                response = CType(request.GetResponse(), HttpWebResponse)

                'Read Source String
                sr = New System.IO.StreamReader(response.GetResponseStream())
                sourceString = sr.ReadToEnd()

                'Check the result
                matches = Regex.Matches(sourceString, regexSearch)

                retry = True
                retries += 1
                GoTo Retry

            End If

            'Set finalText.Text
            finalText = String.Format("#{0} [{1}]: {2}" & Environment.NewLine, id, songName, downloadLink)

            'Put Link on a Text File
            File.AppendAllText(fileName, finalText)

            'Set retries to 0 for future retries
            retries = 0

            'For advise that the function has finished
            'MsgBox("Done!")

        Catch ex As Exception
            File.AppendAllText(fileName, String.Format("#{0} [{1}]: {2}" & Environment.NewLine, id, songName, ex.Message))
        End Try

    End Sub

    Private Sub mp3skull_Search(ByVal songName As String)

        Try

            Update_ToolStrip_Progress()
            id = IndexPattern

            motor = "http://mp3skull.com/"
            regexSearch = "<a href="".+color:green"

            request = CType(HttpWebRequest.Create(motor & "mp3/" & songName & ".html"), HttpWebRequest)
            response = CType(request.GetResponse(), HttpWebResponse)

            sr = New System.IO.StreamReader(response.GetResponseStream())
            sourceString = sr.ReadToEnd()

            'Check the result
            matches = Regex.Matches(sourceString, regexSearch)

            'If results = 0 stop sub
            If matches.Count = 0 Then
                'MsgBox("Se encontraron 0 resultados de esta canción.", MsgBoxStyle.Information, "Información")
                File.AppendAllText(fileName, String.Format("#{0} [{1}]: Se encontraron 0 resultados de esta canción." & Environment.NewLine, id, songName))
                Exit Sub
            End If

            'Get first song link
            firstMatch = matches.Item(0).Value

            'Refinating link
            matches = Regex.Matches(firstMatch, "http://.+rel")
            firstMatch = matches.Item(0).Value

            'Setting download link
            downloadLink = firstMatch.Replace(""" rel", "")

            'Set finalText.Text
            finalText = String.Format("#{0} [{1}]: {2}" & Environment.NewLine, id, songName, downloadLink)

            'Put Link on a Text File
            File.AppendAllText(fileName, finalText)

            'For advise that the function has finished
            'MsgBox("Done!")

        Catch ex As Exception
            File.AppendAllText(fileName, String.Format("#{0} [{1}]: {2}" & Environment.NewLine, id, songName, ex.Message))
        End Try

    End Sub

    'Update second progress bar
    Private Sub Update_ToolStrip_Progress()

        Threading.Interlocked.Increment(IndexPattern)
        Threading.Interlocked.Increment(ProgressBar1.Value)

        Label2.Text = _
        String.Format("Descargando {0} de {1} canciones...", _
        IndexPattern.ToString, (Terms.Count - 1).ToString)

    End Sub

#End Region

#Region "Subs"

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Button1.Enabled = IIf(Terms.Count > 1, True, False)

        Label2.Text = _
        IIf(Terms.Count > 1, _
            String.Format("Preparado para buscar {0} canciones...", (Terms.Count - 1).ToString), _
            "Pulse el botón Buscar para empezar a obtener links...")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If String.IsNullOrEmpty(TextBox1.Text) Then
            MsgBox("Debes escribir la canción a buscar", MsgBoxStyle.Critical, "Error")
        Else
            Label2.Text = String.Format("Descargando 0 de {0} canciones...", (Terms.Count - 1).ToString)
            ProgressBar1.Maximum = Terms.Count
            For Each Match As String In Me.Terms
                Select Case SearchEngine

                    Case Is = SearchEngines.mrtzcmp3
                        mrtzcmp3_Search(Match)

                    Case Is = SearchEngines.mp3skull
                        mp3skull_Search(Match)

                End Select
            Next
            Label2.Text = "Terminado!"
            ProgressBar1.Style = ProgressBarStyle.Continuous
        End If

    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
    End Sub

#End Region

End Class
