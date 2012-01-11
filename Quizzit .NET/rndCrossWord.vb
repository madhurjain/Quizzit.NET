Public Class rndCrossWord
    Inherits System.Windows.Forms.Form
    Dim i As Integer
    Dim j As Integer
    Dim x As Int16
    Dim y As Int16
    Dim tmp As Integer
    Dim lblArray(10, 10) As Label
    Dim QArray(10, 10) As Integer
    Dim XLoc As Integer
    Dim YLoc As Integer
    Dim Across_DownQ As String
    Dim UpdateScoreFlag As Boolean
    Dim BlinkFlag As Boolean



#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents QLabel As System.Windows.Forms.Label
    Friend WithEvents QButton As System.Windows.Forms.Button
    Friend WithEvents RevealButton As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TDButton As System.Windows.Forms.Button
    Friend WithEvents TCButton As System.Windows.Forms.Button
    Friend WithEvents TBButton As System.Windows.Forms.Button
    Friend WithEvents TAButton As System.Windows.Forms.Button
    Friend WithEvents TimeLabel As System.Windows.Forms.Label
    Friend WithEvents AxShockwaveFlash2 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents AxShockwaveFlash1 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents LabelA As System.Windows.Forms.Label
    Friend WithEvents BlinkTimer As System.Windows.Forms.Timer
    Friend WithEvents UpTimer As System.Windows.Forms.Timer
    Friend WithEvents MainAxShockwaveFlash As AxShockwaveFlashObjects.AxShockwaveFlash
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(rndCrossWord))
        Me.QLabel = New System.Windows.Forms.Label
        Me.QButton = New System.Windows.Forms.Button
        Me.RevealButton = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TDButton = New System.Windows.Forms.Button
        Me.TCButton = New System.Windows.Forms.Button
        Me.TBButton = New System.Windows.Forms.Button
        Me.TAButton = New System.Windows.Forms.Button
        Me.LabelA = New System.Windows.Forms.Label
        Me.TimeLabel = New System.Windows.Forms.Label
        Me.AxShockwaveFlash2 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.AxShockwaveFlash1 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.BlinkTimer = New System.Windows.Forms.Timer(Me.components)
        Me.UpTimer = New System.Windows.Forms.Timer(Me.components)
        Me.MainAxShockwaveFlash = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.Panel1.SuspendLayout()
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'QLabel
        '
        Me.QLabel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.QLabel.BackColor = System.Drawing.Color.Gray
        Me.QLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QLabel.Font = New System.Drawing.Font("Verdana", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.QLabel.Location = New System.Drawing.Point(32, 504)
        Me.QLabel.Name = "QLabel"
        Me.QLabel.Size = New System.Drawing.Size(500, 100)
        Me.QLabel.TabIndex = 0
        '
        'QButton
        '
        Me.QButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.QButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QButton.Location = New System.Drawing.Point(88, 625)
        Me.QButton.Name = "QButton"
        Me.QButton.Size = New System.Drawing.Size(128, 40)
        Me.QButton.TabIndex = 2
        Me.QButton.Text = "&Start"
        '
        'RevealButton
        '
        Me.RevealButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RevealButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.RevealButton.Location = New System.Drawing.Point(350, 625)
        Me.RevealButton.Name = "RevealButton"
        Me.RevealButton.Size = New System.Drawing.Size(128, 40)
        Me.RevealButton.TabIndex = 3
        Me.RevealButton.Text = "A&nswer"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.TDButton)
        Me.Panel1.Controls.Add(Me.TCButton)
        Me.Panel1.Controls.Add(Me.TBButton)
        Me.Panel1.Controls.Add(Me.TAButton)
        Me.Panel1.Controls.Add(Me.LabelA)
        Me.Panel1.Location = New System.Drawing.Point(800, 504)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 200)
        Me.Panel1.TabIndex = 4
        '
        'TDButton
        '
        Me.TDButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TDButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TDButton.Location = New System.Drawing.Point(30, 150)
        Me.TDButton.Name = "TDButton"
        Me.TDButton.Size = New System.Drawing.Size(125, 25)
        Me.TDButton.TabIndex = 3
        Me.TDButton.Text = "Team &D"
        '
        'TCButton
        '
        Me.TCButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TCButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TCButton.Location = New System.Drawing.Point(30, 110)
        Me.TCButton.Name = "TCButton"
        Me.TCButton.Size = New System.Drawing.Size(125, 25)
        Me.TCButton.TabIndex = 2
        Me.TCButton.Text = "Team &C"
        '
        'TBButton
        '
        Me.TBButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TBButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TBButton.Location = New System.Drawing.Point(30, 70)
        Me.TBButton.Name = "TBButton"
        Me.TBButton.Size = New System.Drawing.Size(125, 25)
        Me.TBButton.TabIndex = 1
        Me.TBButton.Text = "Team &B"
        '
        'TAButton
        '
        Me.TAButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TAButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TAButton.Location = New System.Drawing.Point(30, 30)
        Me.TAButton.Name = "TAButton"
        Me.TAButton.Size = New System.Drawing.Size(125, 25)
        Me.TAButton.TabIndex = 0
        Me.TAButton.Text = "Team &A"
        '
        'LabelA
        '
        Me.LabelA.BackColor = System.Drawing.SystemColors.Control
        Me.LabelA.Location = New System.Drawing.Point(16, 24)
        Me.LabelA.Name = "LabelA"
        Me.LabelA.Size = New System.Drawing.Size(149, 37)
        Me.LabelA.TabIndex = 17
        Me.LabelA.Visible = False
        '
        'TimeLabel
        '
        Me.TimeLabel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TimeLabel.BackColor = System.Drawing.Color.Black
        Me.TimeLabel.Font = New System.Drawing.Font("Verdana", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TimeLabel.Location = New System.Drawing.Point(800, 40)
        Me.TimeLabel.Name = "TimeLabel"
        Me.TimeLabel.Size = New System.Drawing.Size(200, 125)
        Me.TimeLabel.TabIndex = 16
        Me.TimeLabel.Text = "10"
        Me.TimeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AxShockwaveFlash2
        '
        Me.AxShockwaveFlash2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AxShockwaveFlash2.Enabled = True
        Me.AxShockwaveFlash2.Location = New System.Drawing.Point(32, 16)
        Me.AxShockwaveFlash2.Name = "AxShockwaveFlash2"
        Me.AxShockwaveFlash2.OcxState = CType(resources.GetObject("AxShockwaveFlash2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash2.Size = New System.Drawing.Size(574, 60)
        Me.AxShockwaveFlash2.TabIndex = 15
        '
        'AxShockwaveFlash1
        '
        Me.AxShockwaveFlash1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.AxShockwaveFlash1.Enabled = True
        Me.AxShockwaveFlash1.Location = New System.Drawing.Point(799, 250)
        Me.AxShockwaveFlash1.Name = "AxShockwaveFlash1"
        Me.AxShockwaveFlash1.OcxState = CType(resources.GetObject("AxShockwaveFlash1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash1.Size = New System.Drawing.Size(200, 200)
        Me.AxShockwaveFlash1.TabIndex = 14
        '
        'BlinkTimer
        '
        Me.BlinkTimer.Interval = 400
        '
        'UpTimer
        '
        Me.UpTimer.Interval = 1000
        '
        'MainAxShockwaveFlash
        '
        Me.MainAxShockwaveFlash.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MainAxShockwaveFlash.Enabled = True
        Me.MainAxShockwaveFlash.Location = New System.Drawing.Point(32, 112)
        Me.MainAxShockwaveFlash.Name = "MainAxShockwaveFlash"
        Me.MainAxShockwaveFlash.OcxState = CType(resources.GetObject("MainAxShockwaveFlash.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainAxShockwaveFlash.Size = New System.Drawing.Size(496, 360)
        Me.MainAxShockwaveFlash.TabIndex = 17
        '
        'rndCrossWord
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.CornflowerBlue
        Me.ClientSize = New System.Drawing.Size(1028, 746)
        Me.Controls.Add(Me.MainAxShockwaveFlash)
        Me.Controls.Add(Me.TimeLabel)
        Me.Controls.Add(Me.AxShockwaveFlash2)
        Me.Controls.Add(Me.AxShockwaveFlash1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.RevealButton)
        Me.Controls.Add(Me.QButton)
        Me.Controls.Add(Me.QLabel)
        Me.Name = "rndCrossWord"
        Me.Text = "CrossWord"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub rndCrossWord_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TeamBtn = ""
        RoundID = "CRW"
        UpdateScoreFlag = False
        A_Score = 0
        B_Score = 0
        C_Score = 0
        D_Score = 0
        XLoc = 100
        YLoc = 100
        Across_DownQ = "AcrossQ"
        ReadFile()
        ''''Dynamically Creating An Array Of Labels
        For i = 0 To 9
            For j = 0 To 9
                lblArray(i, j) = New Label
                lblArray(i, j).Name = "Label" & i & j
                lblArray(i, j).Size = New Size(35, 35)
                lblArray(i, j).ForeColor = Color.Black
                lblArray(i, j).Font = New System.Drawing.Font("Verdana", 16.0!, FontStyle.Regular)
                lblArray(i, j).BorderStyle = BorderStyle.Fixed3D
                lblArray(i, j).TextAlign = ContentAlignment.MiddleCenter
                lblArray(i, j).Location = New Point(XLoc, YLoc)
                lblArray(i, j).BackColor = Color.White
                If TxtArray(i, j) = " " Then
                    lblArray(i, j).BackColor = Color.Black
                Else
                    'lblArray(i, j).Text = TxtArray(i, j)
                End If
                XLoc = XLoc + 35
                AddHandler lblArray(i, j).MouseDown, New System.windows.Forms.MouseEventHandler(AddressOf Label_MouseDown)
                Me.Controls.Add(lblArray(i, j))
            Next
            XLoc = 100
            YLoc = YLoc + 35
        Next
        ''''Init. Q Array
        tmp = 1
        For i = 0 To 9
            For j = 0 To 9
                QArray(i, j) = tmp
                tmp = tmp + 1
            Next
        Next


        Try
            OpenConnection()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        AxShockwaveFlash2.Size = New Size(496, 65)
        AxShockwaveFlash2.Location = New Point(32, 16)
        LoadFlash()

        'AxShockwaveFlash1.Location = New Point(700, 250)
        'TimeLabel.Location = New Point(700, 20)
        'Panel1.Location = New Point(700, 400)
    End Sub
    Private Sub Label_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        ClearCRW()
        Dim DumLbl As Label = CType(sender, Label)
        i = Microsoft.VisualBasic.Val(DumLbl.Name.Chars(5))
        j = Microsoft.VisualBasic.Val(DumLbl.Name.Chars(6))
        x = i
        y = j
        If e.Button = MouseButtons.Right Then
            Across_DownQ = "DownQ"
            tmp = i
            For i = tmp To 9
                If TxtArray(i, j) <> " " Then
                    'lblArray(i, j).Text = TxtArray(i, j)
                    lblArray(i, j).BackColor = Color.Yellow
                Else
                    Exit For
                End If
            Next
        ElseIf e.Button = MouseButtons.Left Then
            Across_DownQ = "AcrossQ"
            tmp = j
            For j = tmp To 9
                If TxtArray(i, j) <> " " Then
                    'lblArray(i, j).Text = TxtArray(i, j)
                    lblArray(i, j).BackColor = Color.Yellow
                Else
                    Exit For
                End If
            Next
        Else
            ClearCRW()
        End If
        RevealButton.Text = "A&nswer"
        QLabel.Text = ""
        TeamBtn = ""
        StartPolling()
        TeamBlink()
        UpTimer.Stop()
        'i = Microsoft.VisualBasic.Val(DumLbl.Name.Chars(5))
        'j = Microsoft.VisualBasic.Val(DumLbl.Name.Chars(6))
        'MsgBox(DumLbl.Name + Chr(13) & i & ", " & j)
    End Sub
    Private Sub RevealButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RevealButton.Click
        If RevealButton.Text = "A&nswer" Then
            For i = 0 To 9
                For j = 0 To 9
                    If lblArray(i, j).BackColor.ToString = "Color [Yellow]" Then
                        lblArray(i, j).Text = TxtArray(i, j)
                        lblArray(i, j).BackColor = Color.GreenYellow
                    End If
                Next
            Next
            RevealButton.Text = "A&nswered"
            UpdateScoreFlag = True
        ElseIf RevealButton.Text = "A&nswered" Then
            RevealButton.Text = "Una&nswered"
        ElseIf RevealButton.Text = "Una&nswered" Then
            RevealButton.Text = "A&nswered"
        End If
        Poll = False
        UpTimer.Stop()
    End Sub
    Private Sub QButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QButton.Click
        If QButton.Text = "&Start" Then
            MainAxShockwaveFlash.StopPlay()
            MainAxShockwaveFlash.Visible = False
            QButton.Text = "&Question"
            Exit Sub
        End If
        RevealButton.Text = "A&nswer"
        Try
            Recordset("SELECT * FROM CrossWordz_T WHERE LabelID = " & QArray(x, y) & " AND MainRound = '" + MainRound + "'")
            'MsgBox("SELECT * FROM CrossWordz_T WHERE LabelID = " & QArray(x, y) & " AND MainRound = '" + MainRound + "'")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If Not MyRecordset.EOF Then
            QLabel.Text = MyRecordset.Fields(Across_DownQ).Value
            UpTime = 9
            UpTimer.Start()
        Else
            QLabel.Text = "No Question Found !"
        End If
        CloseRecordset()
    End Sub
    Private Sub ClearCRW()
        For i = 0 To 9
            For j = 0 To 9
                If lblArray(i, j).BackColor.ToString = "Color [Yellow]" Then
                    lblArray(i, j).BackColor = Color.White
                End If
            Next
        Next
    End Sub
    Private Sub TAButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TAButton.Click
        TeamBtn = "A"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If UpdateScoreFlag = True Then
            If RevealButton.Text = "A&nswered" Then
                A_Score = A_Score + 20
            ElseIf RevealButton.Text = "Una&nswered" Then
                A_Score = A_Score - 10
            End If
            QLabel.Text = ""
            UpdateScoreFlag = False
            TimeLabel.Text = ""
            UpTimer.Stop()
            UpTime = 10
        End If
    End Sub
    Private Sub TBButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBButton.Click
        TeamBtn = "B"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If UpdateScoreFlag = True Then
            If RevealButton.Text = "A&nswered" Then
                B_Score = B_Score + 20
            ElseIf RevealButton.Text = "Una&nswered" Then
                B_Score = B_Score - 10
            End If
            QLabel.Text = ""
            UpdateScoreFlag = False
            TimeLabel.Text = ""
            UpTimer.Stop()
            UpTime = 10
        End If
    End Sub
    Private Sub TCButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TCButton.Click
        TeamBtn = "C"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If UpdateScoreFlag = True Then
            If RevealButton.Text = "A&nswered" Then
                C_Score = C_Score + 20
            ElseIf RevealButton.Text = "Una&nswered" Then
                C_Score = C_Score - 10
            End If
            UpdateScoreFlag = False
            TimeLabel.Text = ""
            UpTimer.Stop()
            UpTime = 10
            QLabel.Text = ""
        End If
    End Sub
    Private Sub TDButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TDButton.Click
        TeamBtn = "D"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If UpdateScoreFlag = True Then
            If RevealButton.Text = "A&nswered" Then
                D_Score = D_Score + 20
            ElseIf RevealButton.Text = "Una&nswered" Then
                D_Score = D_Score - 10
            End If
            UpdateScoreFlag = False
            TimeLabel.Text = ""
            UpTimer.Stop()
            UpTime = 10
            QLabel.Text = ""
        End If
    End Sub
    Private Sub rndCrossWord_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Poll = False
        Try
            OpenScoreSet("SELECT * FROM Scores WHERE SubRound = '" + RoundID + "'")
            ScoreSet.Fields(1).Value = A_Score
            ScoreSet.Fields(2).Value = B_Score
            ScoreSet.Fields(3).Value = C_Score
            ScoreSet.Fields(4).Value = D_Score
            ScoreSet.UpdateBatch(ADODB.AffectEnum.adAffectAll)
            CloseScoreSet()
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error Updating Scores !!")
        End Try
        Try
            CloseRecordset()
            CloseConnection()
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error In Closing Connection/Recordset")
        End Try
    End Sub
    Private Sub BlinkTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BlinkTimer.Tick
        If BlinkFlag = True Then
            LabelA.BackColor = Color.Blue
            BlinkFlag = False
        Else
            LabelA.BackColor = Color.Tomato
            BlinkFlag = True
        End If
    End Sub
    Private Sub UpTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpTimer.Tick
        If UpTime > 0 Then
            TimeLabel.Text = UpTime
            UpTime = UpTime - 1
        ElseIf UpTime = 0 Then
            TimeLabel.Text = UpTime
            PlaySnd(ResourcePath + "\Tones\timer1.wav")
            UpTimer.Stop()
            Poll = False
        End If
    End Sub
    Private Sub TeamBlink()
        Select Case TeamBtn
            Case "A"
                LabelA.Location = New Point(17, 24)
                LabelA.Visible = True
                BlinkTimer.Start()
            Case "B"
                LabelA.Location = New Point(17, 64)
                LabelA.Visible = True
                BlinkTimer.Start()
            Case "C"
                LabelA.Location = New Point(17, 104)
                LabelA.Visible = True
                BlinkTimer.Start()
            Case "D"
                LabelA.Location = New Point(17, 144)
                LabelA.Visible = True
                BlinkTimer.Start()
        End Select
    End Sub
    Private Sub StopBlink()
        LabelA.BackColor = Color.Blue
        LabelA.Visible = False
        BlinkTimer.Stop()


    End Sub

    Private Sub LoadFlash()
        MainAxShockwaveFlash.Movie = ResourcePath + "\Extras\" + RoundID + "_main.swf"
        MainAxShockwaveFlash.BringToFront()
        MainAxShockwaveFlash.Play()
        AxShockwaveFlash1.Movie = ResourcePath + "\Extras\Main.swf"
        AxShockwaveFlash2.Movie = ResourcePath + "\Extras\" + RoundID + "_banner.swf"
        AxShockwaveFlash1.Play()
        AxShockwaveFlash2.Play()
        AxShockwaveFlash1.Loop = True
        AxShockwaveFlash2.Loop = True
    End Sub

End Class
