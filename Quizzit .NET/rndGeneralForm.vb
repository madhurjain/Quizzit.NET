Public Class rndGeneralForm
    Inherits System.Windows.Forms.Form
    Dim Qquery As String
    Dim CorrectAnswer As String
    Dim Passed As Boolean
    Dim UpdateScoreFlag As Boolean
    Dim BlinkFlag As Boolean
    Dim Passcount As Integer


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        Application.EnableVisualStyles()
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
    Friend WithEvents genQLabel As System.Windows.Forms.Label
    Friend WithEvents genAnsLabel As System.Windows.Forms.Label
    Friend WithEvents genTBButton As System.Windows.Forms.Button
    Friend WithEvents genTCButton As System.Windows.Forms.Button
    Friend WithEvents genTDButton As System.Windows.Forms.Button
    Friend WithEvents genAnsButton As System.Windows.Forms.Button
    Friend WithEvents genPassButton As System.Windows.Forms.Button
    Friend WithEvents genStartButton As System.Windows.Forms.Button
    Friend WithEvents genTAButton As System.Windows.Forms.Button
    Friend WithEvents genBtnPanel As System.Windows.Forms.Panel
    Friend WithEvents genTeamPanel As System.Windows.Forms.Panel
    Friend WithEvents AxShockwaveFlash1 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents TimeLabel As System.Windows.Forms.Label
    Friend WithEvents UpTimer As System.Windows.Forms.Timer
    Friend WithEvents BlinkTimer As System.Windows.Forms.Timer
    Friend WithEvents LabelA As System.Windows.Forms.Label
    Friend WithEvents MainAxShockwaveFlash As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents AxShockwaveFlash2 As AxShockwaveFlashObjects.AxShockwaveFlash
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(rndGeneralForm))
        Me.genTBButton = New System.Windows.Forms.Button()
        Me.genTCButton = New System.Windows.Forms.Button()
        Me.genTDButton = New System.Windows.Forms.Button()
        Me.genQLabel = New System.Windows.Forms.Label()
        Me.genAnsLabel = New System.Windows.Forms.Label()
        Me.genBtnPanel = New System.Windows.Forms.Panel()
        Me.genAnsButton = New System.Windows.Forms.Button()
        Me.genPassButton = New System.Windows.Forms.Button()
        Me.genStartButton = New System.Windows.Forms.Button()
        Me.genTAButton = New System.Windows.Forms.Button()
        Me.genTeamPanel = New System.Windows.Forms.Panel()
        Me.LabelA = New System.Windows.Forms.Label()
        Me.AxShockwaveFlash1 = New AxShockwaveFlashObjects.AxShockwaveFlash()
        Me.TimeLabel = New System.Windows.Forms.Label()
        Me.UpTimer = New System.Windows.Forms.Timer(Me.components)
        Me.BlinkTimer = New System.Windows.Forms.Timer(Me.components)
        Me.MainAxShockwaveFlash = New AxShockwaveFlashObjects.AxShockwaveFlash()
        Me.AxShockwaveFlash2 = New AxShockwaveFlashObjects.AxShockwaveFlash()
        Me.genBtnPanel.SuspendLayout()
        Me.genTeamPanel.SuspendLayout()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'genTBButton
        '
        Me.genTBButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.genTBButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genTBButton.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genTBButton.ForeColor = System.Drawing.Color.Black
        Me.genTBButton.Location = New System.Drawing.Point(30, 70)
        Me.genTBButton.Name = "genTBButton"
        Me.genTBButton.Size = New System.Drawing.Size(125, 25)
        Me.genTBButton.TabIndex = 1
        Me.genTBButton.Text = "Team &B"
        '
        'genTCButton
        '
        Me.genTCButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.genTCButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genTCButton.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genTCButton.ForeColor = System.Drawing.Color.Black
        Me.genTCButton.Location = New System.Drawing.Point(30, 110)
        Me.genTCButton.Name = "genTCButton"
        Me.genTCButton.Size = New System.Drawing.Size(125, 25)
        Me.genTCButton.TabIndex = 2
        Me.genTCButton.Text = "Team &C"
        '
        'genTDButton
        '
        Me.genTDButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.genTDButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genTDButton.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genTDButton.ForeColor = System.Drawing.Color.Black
        Me.genTDButton.Location = New System.Drawing.Point(30, 150)
        Me.genTDButton.Name = "genTDButton"
        Me.genTDButton.Size = New System.Drawing.Size(125, 25)
        Me.genTDButton.TabIndex = 3
        Me.genTDButton.Text = "Team &D"
        '
        'genQLabel
        '
        Me.genQLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.genQLabel.BackColor = System.Drawing.Color.Gray
        Me.genQLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genQLabel.Font = New System.Drawing.Font("Verdana", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genQLabel.Location = New System.Drawing.Point(25, 50)
        Me.genQLabel.Name = "genQLabel"
        Me.genQLabel.Size = New System.Drawing.Size(479, 300)
        Me.genQLabel.TabIndex = 5
        '
        'genAnsLabel
        '
        Me.genAnsLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.genAnsLabel.BackColor = System.Drawing.Color.Gray
        Me.genAnsLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genAnsLabel.Font = New System.Drawing.Font("Verdana", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genAnsLabel.Location = New System.Drawing.Point(25, 355)
        Me.genAnsLabel.Name = "genAnsLabel"
        Me.genAnsLabel.Size = New System.Drawing.Size(479, 100)
        Me.genAnsLabel.TabIndex = 6
        '
        'genBtnPanel
        '
        Me.genBtnPanel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.genBtnPanel.BackColor = System.Drawing.SystemColors.Control
        Me.genBtnPanel.Controls.Add(Me.genAnsButton)
        Me.genBtnPanel.Controls.Add(Me.genPassButton)
        Me.genBtnPanel.Controls.Add(Me.genStartButton)
        Me.genBtnPanel.Location = New System.Drawing.Point(61, 500)
        Me.genBtnPanel.Name = "genBtnPanel"
        Me.genBtnPanel.Size = New System.Drawing.Size(471, 64)
        Me.genBtnPanel.TabIndex = 7
        '
        'genAnsButton
        '
        Me.genAnsButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.genAnsButton.BackColor = System.Drawing.SystemColors.Control
        Me.genAnsButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genAnsButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genAnsButton.Location = New System.Drawing.Point(352, 16)
        Me.genAnsButton.Name = "genAnsButton"
        Me.genAnsButton.Size = New System.Drawing.Size(104, 32)
        Me.genAnsButton.TabIndex = 3
        Me.genAnsButton.Text = "A&nswer"
        Me.genAnsButton.UseVisualStyleBackColor = False
        '
        'genPassButton
        '
        Me.genPassButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.genPassButton.BackColor = System.Drawing.SystemColors.Control
        Me.genPassButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genPassButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genPassButton.Location = New System.Drawing.Point(191, 16)
        Me.genPassButton.Name = "genPassButton"
        Me.genPassButton.Size = New System.Drawing.Size(80, 32)
        Me.genPassButton.TabIndex = 2
        Me.genPassButton.Text = "&Pass"
        Me.genPassButton.UseVisualStyleBackColor = False
        '
        'genStartButton
        '
        Me.genStartButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.genStartButton.BackColor = System.Drawing.SystemColors.Control
        Me.genStartButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genStartButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genStartButton.Location = New System.Drawing.Point(16, 16)
        Me.genStartButton.Name = "genStartButton"
        Me.genStartButton.Size = New System.Drawing.Size(96, 32)
        Me.genStartButton.TabIndex = 0
        Me.genStartButton.Text = "&Start"
        Me.genStartButton.UseVisualStyleBackColor = False
        '
        'genTAButton
        '
        Me.genTAButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.genTAButton.BackColor = System.Drawing.Color.Silver
        Me.genTAButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.genTAButton.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.genTAButton.ForeColor = System.Drawing.Color.Black
        Me.genTAButton.Location = New System.Drawing.Point(30, 30)
        Me.genTAButton.Name = "genTAButton"
        Me.genTAButton.Size = New System.Drawing.Size(125, 25)
        Me.genTAButton.TabIndex = 0
        Me.genTAButton.Text = "Team &A"
        Me.genTAButton.UseVisualStyleBackColor = False
        '
        'genTeamPanel
        '
        Me.genTeamPanel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.genTeamPanel.BackColor = System.Drawing.SystemColors.Control
        Me.genTeamPanel.Controls.Add(Me.genTBButton)
        Me.genTeamPanel.Controls.Add(Me.genTCButton)
        Me.genTeamPanel.Controls.Add(Me.genTDButton)
        Me.genTeamPanel.Controls.Add(Me.genTAButton)
        Me.genTeamPanel.Controls.Add(Me.LabelA)
        Me.genTeamPanel.Location = New System.Drawing.Point(573, 365)
        Me.genTeamPanel.Name = "genTeamPanel"
        Me.genTeamPanel.Size = New System.Drawing.Size(200, 200)
        Me.genTeamPanel.TabIndex = 8
        '
        'LabelA
        '
        Me.LabelA.BackColor = System.Drawing.SystemColors.Control
        Me.LabelA.Location = New System.Drawing.Point(16, 82)
        Me.LabelA.Name = "LabelA"
        Me.LabelA.Size = New System.Drawing.Size(149, 37)
        Me.LabelA.TabIndex = 5
        Me.LabelA.Visible = False
        '
        'AxShockwaveFlash1
        '
        Me.AxShockwaveFlash1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.AxShockwaveFlash1.Enabled = True
        Me.AxShockwaveFlash1.Location = New System.Drawing.Point(573, 150)
        Me.AxShockwaveFlash1.Name = "AxShockwaveFlash1"
        Me.AxShockwaveFlash1.OcxState = CType(resources.GetObject("AxShockwaveFlash1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash1.Size = New System.Drawing.Size(200, 200)
        Me.AxShockwaveFlash1.TabIndex = 10
        '
        'TimeLabel
        '
        Me.TimeLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TimeLabel.BackColor = System.Drawing.Color.Black
        Me.TimeLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TimeLabel.Font = New System.Drawing.Font("Verdana", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TimeLabel.Location = New System.Drawing.Point(573, 56)
        Me.TimeLabel.Name = "TimeLabel"
        Me.TimeLabel.Size = New System.Drawing.Size(200, 72)
        Me.TimeLabel.TabIndex = 12
        Me.TimeLabel.Text = "20"
        Me.TimeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'UpTimer
        '
        Me.UpTimer.Interval = 1000
        '
        'BlinkTimer
        '
        Me.BlinkTimer.Interval = 400
        '
        'MainAxShockwaveFlash
        '
        Me.MainAxShockwaveFlash.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MainAxShockwaveFlash.Enabled = True
        Me.MainAxShockwaveFlash.Location = New System.Drawing.Point(23, 126)
        Me.MainAxShockwaveFlash.Name = "MainAxShockwaveFlash"
        Me.MainAxShockwaveFlash.OcxState = CType(resources.GetObject("MainAxShockwaveFlash.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainAxShockwaveFlash.Size = New System.Drawing.Size(481, 280)
        Me.MainAxShockwaveFlash.TabIndex = 13
        '
        'AxShockwaveFlash2
        '
        Me.AxShockwaveFlash2.Enabled = True
        Me.AxShockwaveFlash2.Location = New System.Drawing.Point(25, 10)
        Me.AxShockwaveFlash2.Name = "AxShockwaveFlash2"
        Me.AxShockwaveFlash2.OcxState = CType(resources.GetObject("AxShockwaveFlash2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash2.Size = New System.Drawing.Size(479, 40)
        Me.AxShockwaveFlash2.TabIndex = 14
        '
        'rndGeneralForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.RoyalBlue
        Me.ClientSize = New System.Drawing.Size(793, 630)
        Me.Controls.Add(Me.AxShockwaveFlash2)
        Me.Controls.Add(Me.MainAxShockwaveFlash)
        Me.Controls.Add(Me.TimeLabel)
        Me.Controls.Add(Me.AxShockwaveFlash1)
        Me.Controls.Add(Me.genBtnPanel)
        Me.Controls.Add(Me.genAnsLabel)
        Me.Controls.Add(Me.genQLabel)
        Me.Controls.Add(Me.genTeamPanel)
        Me.Name = "rndGeneralForm"
        Me.Text = "General"
        Me.genBtnPanel.ResumeLayout(False)
        Me.genTeamPanel.ResumeLayout(False)
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub rndGeneralForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TeamBtn = ""
        RoundID = "GEN"
        Passed = False
        UpTime = 20
        UpdateScoreFlag = True
        A_Score = 0
        B_Score = 0
        C_Score = 0
        D_Score = 0
        Passcount = 0
        Qquery = "SELECT * FROM " + MainRound + " WHERE RndID='" + RoundID + "'"
        Try
            OpenConnection()
            Recordset(Qquery)
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error Opening Connection/Recordset")
        End Try
        LoadFlash()
    End Sub
    Private Sub genTAButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles genTAButton.Click
        TeamBtn = "A"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If genAnsButton.Text = "A&nswered" And UpdateScoreFlag = True Then
            If Passed = False Then
                A_Score = A_Score + 20
            Else
                A_Score = A_Score + 5
                Passed = False
            End If
            UpdateScoreFlag = False
        End If
        If genAnsButton.Text <> "A&nswer" Then
            genAnsLabel.Text = ""
            genQLabel.Text = ""
            genPassButton.Enabled = True
            Passcount = 0
            genAnsButton.Text = "A&nswer"
            TimeLabel.Text = ""
        End If
    End Sub
    Private Sub genTBButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles genTBButton.Click
        TeamBtn = "B"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If genAnsButton.Text = "A&nswered" And UpdateScoreFlag = True Then
            If Passed = False Then
                B_Score = B_Score + 20
            Else
                B_Score = B_Score + 5
                Passed = False
            End If
            UpdateScoreFlag = False
        End If
        If genAnsButton.Text <> "A&nswer" Then
            genAnsLabel.Text = ""
            genQLabel.Text = ""
            genPassButton.Enabled = True
            Passcount = 0
            genAnsButton.Text = "A&nswer"
            TimeLabel.Text = ""
        End If
    End Sub
    Private Sub genTCButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles genTCButton.Click
        TeamBtn = "C"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If genAnsButton.Text = "A&nswered" And UpdateScoreFlag = True Then
            If Passed = False Then
                C_Score = C_Score + 20
            Else
                C_Score = C_Score + 5
                Passed = False
            End If
            UpdateScoreFlag = False
        End If
        If genAnsButton.Text <> "A&nswer" Then
            genAnsLabel.Text = ""
            genQLabel.Text = ""
            genPassButton.Enabled = True
            Passcount = 0
            genAnsButton.Text = "A&nswer"
            TimeLabel.Text = ""
        End If
    End Sub
    Private Sub genTDButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles genTDButton.Click
        TeamBtn = "D"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If genAnsButton.Text = "A&nswered" And UpdateScoreFlag = True Then
            If Passed = False Then
                D_Score = D_Score + 20
            Else
                D_Score = D_Score + 5
                Passed = False
            End If
            UpdateScoreFlag = False
        End If
        If genAnsButton.Text <> "A&nswer" Then
            genAnsLabel.Text = ""
            genQLabel.Text = ""
            genPassButton.Enabled = True
            Passcount = 0
            genAnsButton.Text = "A&nswer"
            TimeLabel.Text = ""
        End If
    End Sub
    Private Sub genStartButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles genStartButton.Click
        genAnsLabel.Text = ""
        Passcount = 0
        genPassButton.Enabled = True
        If genStartButton.Text = "&Start" Then
            genStartButton.Text = "&Question"
            MainAxShockwaveFlash.StopPlay()
            MainAxShockwaveFlash.Visible = False
        ElseIf genStartButton.Text = "&Question" Then
            UpTime = 19
            UpTimer.Start()
            Passed = False
            Try
                If Not MyRecordset.EOF Then
                    genQLabel.Text = MyRecordset.Fields(2).Value
                    CorrectAnswer = MyRecordset.Fields(3).Value
                    MyRecordset.MoveNext()
                Else
                    Stats2 = "Questions Over !!"
                End If
            Catch ex As Exception
                MsgBox(ex.Message + Chr(13) + "Recordset Error! Cannot Verify EOF")
            End Try
        End If
        genPassButton.Enabled = True
        genAnsButton.Text = "A&nswer"
        UpdateScoreFlag = True
    End Sub
    Private Sub genPassButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles genPassButton.Click
        If Passcount < 3 Then
            Passed = True
            Select Case TeamBtn
                Case "A"
                    LabelA.Location = New Point(17, 64)
                    TeamBtn = "B"
                Case "B"
                    LabelA.Location = New Point(17, 104)
                    TeamBtn = "C"
                Case "C"
                    LabelA.Location = New Point(17, 144)
                    TeamBtn = "D"
                Case "D"
                    LabelA.Location = New Point(17, 24)
                    TeamBtn = "A"
            End Select
            Passcount = Passcount + 1
            UpTimer.Stop()
        Else
            genPassButton.Enabled = False
            StopBlink()
        End If
    End Sub
    Private Sub genAnsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles genAnsButton.Click
        UpTimer.Stop()
        If genAnsButton.Text = "A&nswer" Then
            genAnsLabel.Text = CorrectAnswer
            genAnsButton.Text = "A&nswered"
        ElseIf genAnsButton.Text = "A&nswered" Then
            genAnsButton.Text = "Una&nswered"
        ElseIf genAnsButton.Text = "Una&nswered" Then
            genAnsButton.Text = "A&nswered"
        End If
    End Sub
    Private Sub rndGeneralForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
    Private Sub UpTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpTimer.Tick
        If UpTime > 0 Then
            TimeLabel.Text = UpTime
            UpTime = UpTime - 1
        ElseIf UpTime = 0 Then
            TimeLabel.Text = UpTime
            PlaySnd(ResourcePath + "\Tones\timer1.wav")
            UpTimer.Stop()
        End If
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
        MainAxShockwaveFlash.Play()
        AxShockwaveFlash1.Movie = ResourcePath + "\Extras\Main.swf"
        AxShockwaveFlash2.Movie = ResourcePath + "\Extras\" + RoundID + "_banner.swf"
        AxShockwaveFlash1.Play()
        AxShockwaveFlash2.Play()
        AxShockwaveFlash1.Loop = True
        AxShockwaveFlash2.Loop = True
    End Sub

End Class


