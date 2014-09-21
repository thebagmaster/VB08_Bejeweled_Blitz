Public Class frmBlitzen
    Public l As Integer, t As Integer, color
    Dim pic(13, 13) As Windows.Forms.PictureBox
    Public xt As Integer, yt As Integer, xt2 As Integer, yt2 As Integer
    Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
    Public Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Integer) As Integer
    Public Declare Function GetDesktopWindow Lib "user32.dll" () As Integer
    Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Public Sub LeftDown()
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    End Sub
    Public Sub LeftUp()
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If yt = 8 Then
            yt = 0
            Timer2.Enabled = True
        End If
        sqrUp(yt)
        If yt < 8 Then
            yt = yt + 1
        End If
    End Sub
    Private Function choose(ByVal px As Integer, ByVal py As Integer) As String
        Dim dir As String = ""
        Dim pnt As Point
        pnt.X = px
        pnt.Y = py
        Dim dirray(3) As Integer
        Dim MaxIntegersIndex As Integer = 0
        dirray(0) = checkUp(pnt)
        dirray(1) = checkDown(pnt)
        dirray(2) = checkLeft(pnt)
        dirray(3) = checkRight(pnt)
        For i As Integer = 0 To UBound(dirray)
            If dirray(i) > dirray(MaxIntegersIndex) Then
                MaxIntegersIndex = i
            End If
        Next
        If Not dirray(MaxIntegersIndex) = 0 Then
            Select Case MaxIntegersIndex
                Case 0
                    dir = "u"
                Case 1
                    dir = "d"
                Case 2
                    dir = "l"
                Case 3
                    dir = "r"
            End Select
        End If
        Return dir
    End Function
    Private Function compare3(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal x3 As Integer, ByVal y3 As Integer)
        Dim boo As Boolean = False
        boo = (pic(x1, y1).BackColor = pic(x2, y2).BackColor) And (pic(x1, y1).BackColor = pic(x3, y3).BackColor)
        Return boo
    End Function
    Private Function compare5(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal x3 As Integer, ByVal y3 As Integer, ByVal x4 As Integer, ByVal y4 As Integer, ByVal x5 As Integer, ByVal y5 As Integer)
        Dim boo As Boolean = False
        boo = (pic(x1, y1).BackColor = pic(x2, y2).BackColor) And (pic(x1, y1).BackColor = pic(x3, y3).BackColor) And (pic(x1, y1).BackColor = pic(x4, y4).BackColor) And (pic(x1, y1).BackColor = pic(x5, y5).BackColor)
        Return boo
    End Function
    Private Function checkUp(ByVal pnt As Point)
        Dim score As Integer = 0
        If compare3(pnt.X - 1, pnt.Y - 1, pnt.X + 1, pnt.Y - 1, pnt.X, pnt.Y) Or _
           compare3(pnt.X - 1, pnt.Y - 1, pnt.X - 2, pnt.Y - 1, pnt.X, pnt.Y) Or _
           compare3(pnt.X + 1, pnt.Y - 1, pnt.X + 2, pnt.Y - 1, pnt.X, pnt.Y) Or _
           compare3(pnt.X, pnt.Y - 2, pnt.X, pnt.Y - 3, pnt.X, pnt.Y) Then
            score = 3
        Else
        End If
        If compare3(pnt.X - 1, pnt.Y - 1, pnt.X + 1, pnt.Y - 1, pnt.X, pnt.Y) And _
          (pic(pnt.X - 1, pnt.Y - 1).BackColor.Equals(pic(pnt.X - 2, pnt.Y - 1).BackColor) Or _
           pic(pnt.X - 1, pnt.Y - 1).BackColor.Equals(pic(pnt.X + 2, pnt.Y - 1).BackColor)) Then
            score = 4
        End If
        If compare5(pnt.X - 1, pnt.Y - 1, pnt.X + 1, pnt.Y - 1, pnt.X, pnt.Y, pnt.X - 2, pnt.Y - 1, pnt.X + 2, pnt.Y - 1) Then
            score = 5
        End If
        Return (score)
    End Function
    Private Function checkDown(ByVal pnt As Point)
        Dim score As Integer = 0
        If compare3(pnt.X - 1, pnt.Y + 1, pnt.X + 1, pnt.Y + 1, pnt.X, pnt.Y) Or _
           compare3(pnt.X - 1, pnt.Y + 1, pnt.X - 2, pnt.Y + 1, pnt.X, pnt.Y) Or _
           compare3(pnt.X + 1, pnt.Y + 1, pnt.X + 2, pnt.Y + 1, pnt.X, pnt.Y) Or _
           compare3(pnt.X, pnt.Y + 2, pnt.X, pnt.Y + 3, pnt.X, pnt.Y) Then
            score = 3
        Else
        End If
        If compare3(pnt.X - 1, pnt.Y + 1, pnt.X + 1, pnt.Y + 1, pnt.X, pnt.Y) And _
          (pic(pnt.X - 1, pnt.Y + 1).BackColor.Equals(pic(pnt.X - 2, pnt.Y + 1).BackColor) Or _
           pic(pnt.X - 1, pnt.Y + 1).BackColor.Equals(pic(pnt.X + 2, pnt.Y + 1).BackColor)) Then
            score = 4
        End If
        If compare5(pnt.X - 1, pnt.Y + 1, pnt.X + 1, pnt.Y + 1, pnt.X, pnt.Y, pnt.X - 2, pnt.Y + 1, pnt.X + 2, pnt.Y + 1) Then
            score = 5
        End If
        Return (score)
    End Function
    Private Function checkLeft(ByVal pnt As Point)
        Dim score As Integer = 0
        If compare3(pnt.X - 1, pnt.Y - 1, pnt.X - 1, pnt.Y + 1, pnt.X, pnt.Y) Or _
           compare3(pnt.X - 1, pnt.Y - 1, pnt.X - 1, pnt.Y - 2, pnt.X, pnt.Y) Or _
           compare3(pnt.X - 1, pnt.Y + 1, pnt.X - 1, pnt.Y + 2, pnt.X, pnt.Y) Or _
           compare3(pnt.X - 2, pnt.Y, pnt.X - 3, pnt.Y, pnt.X, pnt.Y) Then
            score = 3
        Else
        End If
        If compare3(pnt.X - 1, pnt.Y - 1, pnt.X - 1, pnt.Y + 1, pnt.X, pnt.Y) And _
          (pic(pnt.X - 1, pnt.Y - 1).BackColor.Equals(pic(pnt.X - 1, pnt.Y - 2).BackColor) Or _
           pic(pnt.X - 1, pnt.Y - 1).BackColor.Equals(pic(pnt.X - 1, pnt.Y + 2).BackColor)) Then
            score = 4
        End If
        If compare5(pnt.X - 1, pnt.Y - 1, pnt.X - 1, pnt.Y + 1, pnt.X, pnt.Y, pnt.X - 1, pnt.Y - 2, pnt.X - 1, pnt.Y + 2) Then
            score = 5
        End If
        Return (score)
    End Function
    Private Function checkRight(ByVal pnt As Point)
        Dim score As Integer = 0
        If compare3(pnt.X + 1, pnt.Y - 1, pnt.X + 1, pnt.Y + 1, pnt.X, pnt.Y) Or _
           compare3(pnt.X + 1, pnt.Y - 1, pnt.X + 1, pnt.Y - 2, pnt.X, pnt.Y) Or _
           compare3(pnt.X + 1, pnt.Y + 1, pnt.X + 1, pnt.Y + 2, pnt.X, pnt.Y) Or _
           compare3(pnt.X + 2, pnt.Y, pnt.X + 3, pnt.Y, pnt.X, pnt.Y) Then
            score = 3
        Else
        End If
        If compare3(pnt.X + 1, pnt.Y - 1, pnt.X + 1, pnt.Y + 1, pnt.X, pnt.Y) And _
          (pic(pnt.X + 1, pnt.Y - 1).BackColor.Equals(pic(pnt.X + 1, pnt.Y - 2).BackColor) Or _
           pic(pnt.X + 1, pnt.Y - 1).BackColor.Equals(pic(pnt.X + 1, pnt.Y + 2).BackColor)) Then
            score = 4
        End If
        If compare5(pnt.X + 1, pnt.Y - 1, pnt.X + 1, pnt.Y + 1, pnt.X, pnt.Y, pnt.X + 1, pnt.Y - 2, pnt.X + 1, pnt.Y + 2) Then
            score = 5
        End If
        Return (score)
    End Function
    Private Function sqrUp(ByVal yo As Integer)
        Dim winDc = GetWindowDC(GetDesktopWindow)
        For xTmp As Integer = 0 To 7
            pic(xTmp + 3, yo + 3).BackColor = System.Drawing.ColorTranslator.FromOle(GetPixel(winDc, l + 20 + 40 * xTmp, t + 20 + 40 * yo).ToString)
        Next
        Return 0
    End Function
    Private Sub frmBlitzen_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Up Then
            Me.Top = Me.Top - 1
        End If
        If e.KeyCode = Keys.Down Then
            Me.Top = Me.Top + 1
        End If
        If e.KeyCode = Keys.Left Then
            Me.Left = Me.Left - 1
        End If
        If e.KeyCode = Keys.Right Then
            Me.Left = Me.Left + 1
        End If
        Button1.Text = e.ToString
    End Sub
    Private Sub frmBlitzen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Randomize()
        Dim rand As New Random
        l = Me.Left
        t = Me.Top + 140
        For Me.xt = 0 To 13
            For Me.yt = 0 To 13
                pic(xt, yt) = New System.Windows.Forms.PictureBox()
                pic(xt, yt).Name = "p" & xt.ToString & yt.ToString
                pic(xt, yt).Left = xt * 10
                pic(xt, yt).Top = yt * 10
                pic(xt, yt).Height = 10
                pic(xt, yt).Width = 10
                pic(xt, yt).Visible = True
                Me.Controls.Add(pic(xt, yt))
                pic(xt, yt).BackColor = Drawing.Color.Black
                pic(xt, yt).BorderStyle = BorderStyle.FixedSingle
            Next
        Next
        yt = 0
        xt2 = 0
        yt2 = 0
        TextBox1.Focus()
    End Sub
    Private Function moveUp(ByVal xo As Integer, ByVal yo As Integer)
        xo = l + 19 + 40 * xo
        yo = t + 19 + 40 * yo
        SetCursorPos(xo, yo)
        LeftDown()
        SetCursorPos(xo, yo - 50)
        LeftUp()
        Return 0
    End Function
    Private Function moveDown(ByVal xo As Integer, ByVal yo As Integer)
        xo = l + 19 + 40 * xo
        yo = t + 19 + 40 * yo
        SetCursorPos(xo, yo)
        LeftDown()
        SetCursorPos(xo, yo + 50)
        LeftUp()
        Return 0
    End Function
    Private Function moveLeft(ByVal xo As Integer, ByVal yo As Integer)
        xo = l + 19 + 40 * xo
        yo = t + 19 + 40 * yo
        SetCursorPos(xo, yo)
        LeftDown()
        SetCursorPos(xo - 50, yo)
        LeftUp()
        Return 0
    End Function
    Private Function moveRight(ByVal xo As Integer, ByVal yo As Integer)
        xo = l + 19 + 40 * xo
        yo = t + 19 + 40 * yo
        SetCursorPos(xo, yo)
        LeftDown()
        SetCursorPos(xo + 50, yo)
        LeftUp()
        Return 0
    End Function
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        runThrough()
    End Sub
    Private Function runThrough()
        pic(xt2 + 3, yt2 + 3).BorderStyle = BorderStyle.None
        If xt2 = 7 And yt2 = 8 Then
            xt2 = 0
            yt2 = 0
        End If
        Select Case choose(xt2 + 3, yt2 + 3)
            Case "u"
                moveUp(xt2, yt2)
            Case "d"
                moveDown(xt2, yt2)
            Case "l"
                moveLeft(xt2, yt2)
            Case "r"
                moveRight(xt2, yt2)
        End Select
        If yt2 < 8 Then
            yt2 = yt2 + 1
        End If
        If yt2 = 8 And xt2 < 7 Then
            yt2 = 0
            xt2 = xt2 + 1
        End If
        pic(xt2 + 3, yt2 + 3).BorderStyle = BorderStyle.FixedSingle
        Return 0
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
    End Sub
    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Up Then
            Me.Top = Me.Top - 1
        End If
        If e.KeyCode = Keys.Down Then
            Me.Top = Me.Top + 1
        End If
        If e.KeyCode = Keys.Left Then
            Me.Left = Me.Left - 1
        End If
        If e.KeyCode = Keys.Right Then
            Me.Left = Me.Left + 1
        End If
    End Sub
End Class
