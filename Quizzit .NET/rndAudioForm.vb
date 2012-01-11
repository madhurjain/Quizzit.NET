Public Class rndAudioForm
    Inherits System.Windows.Forms.Form
    Dim AudWMP As New AxWMPLib.AxWindowsMediaPlayer
    Dim Qquery As String
    Dim CorrectAnswer As String
    Dim tmp As Array
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
    Friend WithEvents BtnPanel As System.Windows.Forms.Panel
    Friend WithEvents TAButton As System.Windows.Forms.Button
    Friend WithEvents TBButton As System.Windows.Forms.Button
    Friend WithEvents TCButton As System.Windows.Forms.Button
    Friend WithEvents TDButton As System.Windows.Forms.Button
    Friend WithEvents AnsLabel As System.Windows.Forms.Label
    Friend WithEvents QLabel As System.Windows.Forms.Label
    Friend WithEvents QButton As System.Windows.Forms.Button
    Friend WithEvents AnsButton As System.Windows.Forms.Button
    Friend WithEvents AxShockwaveFlash1 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents AxShockwaveFlash2 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents TimeLabel As System.Windows.Forms.Label
    Friend WithEvents UpTimer As System.Windows.Forms.Timer
    Friend WithEvents PollButton As System.Windows.Forms.Button
    Friend WithEvents LabelA As System.Windows.Forms.Label
    Friend WithEvents BlinkTimer As System.Windows.Forms.Timer
    Friend WithEvents MainAxShockwaveFlash As AxShockwaveFlashObjects.AxShockwaveFlash
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(rndAudioForm))
        Me.BtnPanel = New System.Windows.Forms.Panel
        Me.TDButton = New System.Windows.Forms.Button
        Me.TCButton = New System.Windows.Forms.Button
        Me.TBButton = New System.Windows.Forms.Button
        Me.TAButton = New System.Windows.Forms.Button
        Me.LabelA = New System.Windows.Forms.Label
        Me.AnsLabel = New System.Windows.Forms.Label
        Me.QLabel = New System.Windows.Forms.Label
        Me.QButton = New System.Windows.Forms.Button
        Me.AnsButton = New System.Windows.Forms.Button
        Me.AxShockwaveFlash1 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.AxShockwaveFlash2 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.TimeLabel = New System.Windows.Forms.Label
        Me.UpTimer = New System.Windows.Forms.Timer(Me.components)
        Me.PollButton = New System.Windows.Forms.Button
        Me.BlinkTimer = New System.Windows.Forms.Timer(Me.components)
        Me.MainAxShockwaveFlash = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.BtnPanel.SuspendLayout()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnPanel
        '
        Me.BtnPanel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnPanel.Controls.Add(Me.TDButton)
        Me.BtnPanel.Controls.Add(Me.TCButton)
        Me.BtnPanel.Controls.Add(Me.TBButton)
        Me.BtnPanel.Controls.Add(Me.TAButton)
        Me.BtnPanel.Controls.Add(Me.LabelA)
        Me.BtnPanel.Location = New System.Drawing.Point(500, 365)
        Me.BtnPanel.Name = "BtnPanel"
        Me.BtnPanel.Size = New System.Drawing.Size(200, 200)
        Me.BtnPanel.TabIndex = 0
        '
        'TDButton
        '
        Me.TDButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TDButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.TDButton.Location = New System.Drawing.Point(30, 150)
        Me.TDButton.Name = "TDButton"
        Me.TDButton.Size = New System.Drawing.Size(144, 24)
        Me.TDButton.TabIndex = 3
        Me.TDButton.Text = "Team &D"
        '
        'TCButton
        '
        Me.TCButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TCButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.TCButton.Location = New System.Drawing.Point(30, 110)
        Me.TCButton.Name = "TCButton"
        Me.TCButton.Size = New System.Drawing.Size(144, 24)
        Me.TCButton.TabIndex = 2
        Me.TCButton.Text = "Team &C"
        '
        'TBButton
        '
        Me.TBButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TBButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.TBButton.Location = New System.Drawing.Point(30, 70)
        Me.TBButton.Name = "TBButton"
        Me.TBButton.Size = New System.Drawing.Size(144, 24)
        Me.TBButton.TabIndex = 1
        Me.TBButton.Text = "Team &B"
        '
        'TAButton
        '
        Me.TAButton.BackColor = System.Drawing.SystemColors.Control
        Me.TAButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TAButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.TAButton.Location = New System.Drawing.Point(30, 30)
        Me.TAButton.Name = "TAButton"
        Me.TAButton.Size = New System.Drawing.Size(144, 24)
        Me.TAButton.TabIndex = 0
        Me.TAButton.Text = "Team &A"
        '
        'LabelA
        '
        Me.LabelA.BackColor = System.Drawing.SystemColors.Control
        Me.LabelA.Location = New System.Drawing.Point(24, 56)
        Me.LabelA.Name = "LabelA"
        Me.LabelA.Size = New System.Drawing.Size(168, 37)
        Me.LabelA.TabIndex = 4
        Me.LabelA.Visible = False
        '
        'AnsLabel
        '
        Me.AnsLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AnsLabel.BackColor = System.Drawing.Color.Gray
        Me.AnsLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.AnsLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AnsLabel.Font = New System.Drawing.Font("Verdana", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AnsLabel.Location = New System.Drawing.Point(25, 395)
        Me.AnsLabel.Name = "AnsLabel"
        Me.AnsLabel.Size = New System.Drawing.Size(450, 100)
        Me.AnsLabel.TabIndex = 3
        '
        'QLabel
        '
        Me.QLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.QLabel.BackColor = System.Drawing.Color.Gray
        Me.QLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.QLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QLabel.Font = New System.Drawing.Font("Verdana", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QLabel.Location = New System.Drawing.Point(25, 50)
        Me.QLabel.Name = "QLabel"
        Me.QLabel.Size = New System.Drawing.Size(450, 300)
        Me.QLabel.TabIndex = 2
        '
        'QButton
        '
        Me.QButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.QButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QButton.Location = New System.Drawing.Point(180, 521)
        Me.QButton.Name = "QButton"
        Me.QButton.Size = New System.Drawing.Size(128, 40)
        Me.QButton.TabIndex = 4
        Me.QButton.Text = "&Start"
        '
        'AnsButton
        '
        Me.AnsButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.AnsButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AnsButton.Location = New System.Drawing.Point(335, 521)
        Me.AnsButton.Name = "AnsButton"
        Me.AnsButton.Size = New System.Drawing.Size(128, 40)
        Me.AnsButton.TabIndex = 5
        Me.AnsButton.Text = "A&nswer"
        '
        'AxShockwaveFlash1
        '
        Me.AxShockwaveFlash1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.AxShockwaveFlash1.Enabled = True
        Me.AxShockwaveFlash1.Location = New System.Drawing.Point(500, 150)
        Me.AxShockwaveFlash1.Name = "AxShockwaveFlash1"
        Me.AxShockwaveFlash1.OcxState = CType(resources.GetObject("AxShockwaveFlash1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash1.Size = New System.Drawing.Size(200, 200)
        Me.AxShockwaveFlash1.TabIndex = 11
        '
        'AxShockwaveFlash2
        '
        Me.AxShockwaveFlash2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AxShockwaveFlash2.Enabled = True
        Me.AxShockwaveFlash2.Location = New System.Drawing.Point(25, 10)
        Me.AxShockwaveFlash2.Name = "AxShockwaveFlash2"
        Me.AxShockwaveFlash2.OcxState = CType(resources.GetObject("AxShockwaveFlash2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash2.Size = New System.Drawing.Size(450, 60)
        Me.AxShockwaveFlash2.TabIndex = 12
        '
        'TimeLabel
        '
        Me.TimeLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TimeLabel.BackColor = System.Drawing.Color.Black
        Me.TimeLabel.Font = New System.Drawing.Font("Verdana", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TimeLabel.Location = New System.Drawing.Point(500, 10)
        Me.TimeLabel.Name = "TimeLabel"
        Me.TimeLabel.Size = New System.Drawing.Size(200, 125)
        Me.TimeLabel.TabIndex = 13
        Me.TimeLabel.Text = "10"
        Me.TimeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'UpTimer
        '
        Me.UpTimer.Interval = 1000
        '
        'PollButton
        '
        Me.PollButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.PollButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.PollButton.Location = New System.Drawing.Point(20, 521)
        Me.PollButton.Name = "PollButton"
        Me.PollButton.Size = New System.Drawing.Size(128, 40)
        Me.PollButton.TabIndex = 14
        Me.PollButton.Text = "&Pollling"
        '
        'BlinkTimer
        '
        Me.BlinkTimer.Interval = 400
        '
        'MainAxShockwaveFlash
        '
        Me.MainAxShockwaveFlash.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MainAxShockwaveFlash.Enabled = True
        Me.MainAxShockwaveFlash.Location = New System.Drawing.Point(24, 72)
        Me.MainAxShockwaveFlash.Name = "MainAxShockwaveFlash"
        Me.MainAxShockwaveFlash.OcxState = CType(resources.GetObject("MainAxShockwaveFlash.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainAxShockwaveFlash.Size = New System.Drawing.Size(452, 424)
        Me.MainAxShockwaveFlash.TabIndex = 15
        '
        'rndAudioForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Firebrick
        Me.ClientSize = New System.Drawing.Size(720, 582)
        Me.Controls.Add(Me.MainAxShockwaveFlash)
        Me.Controls.Add(Me.PollButton)
        Me.Controls.Add(Me.TimeLabel)
        Me.Controls.Add(Me.AxShockwaveFlash2)
        Me.Controls.Add(Me.AxShockwaveFlash1)
        Me.Controls.Add(Me.AnsButton)
        Me.Controls.Add(Me.QButton)
        Me.Controls.Add(Me.AnsLabel)
        Me.Controls.Add(Me.QLabel)
        Me.Controls.Add(Me.BtnPanel)
        Me.Name = "rndAudioForm"
        Me.Text = "Audio"
        Me.BtnPanel.ResumeLayout(False)
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub rndAudioForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TeamBtn = ""
        RoundID = "AUD"
        UpdateScoreFlag = False
        UpTime = 10
        A_Score = 0
        B_Score = 0
        C_Score = 0
        D_Score = 0
        Qquery = "SELECT * FROM " + MainRound + " WHERE RndID='" + RoundID + "'"
        Try
            OpenConnection()
            Recordset(Qquery)
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error Opening Connection/Recordset")
        End Try
        AudWMP.Location = New Point(25, 393)
        AudWMP.Size = New Size(750, 45)
        AudWMP.Visible = True
        Me.Controls.Add(AudWMP)
        AudWMP.BringToFront()
        LoadFlash()
    End Sub
    Private Sub QButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QButton.Click
        If QButton.Text = "&Start" Then
            MainAxShockwaveFlash.StopPlay()
            MainAxShockwaveFlash.Visible = False
            QButton.Text = "&Question"
            Exit Sub
        End If
        If QButton.Text = "&Question" Then
            QButton.Text = "&Timer"
            AnsLabel.Text = ""
            If Not MyRecordset.EOF Then
                QLabel.Text = MyRecordset.Fields("Question").Value
                AnsButton.Text = "A&nswer"
                tmp = Split(MyRecordset.Fields("Answer").Value(), ";", -1)
                CorrectAnswer = tmp(0)
                AudWMP.Ctlcontrols.stop()
                AudWMP.URL = ResourcePath + tmp(1)
                TeamBtn = ""
                StartPolling()
                If AudWMP.playState = WMPLib.WMPPlayState.wmppsPlaying Then
                    AudWMP.Ctlcontrols.stop()
                End If
                UpTimer.Stop()
                TeamBlink()
                MyRecordset.MoveNext()
            Else
                QLabel.Text = "No Questions Found !"
            End If

        ElseIf QButton.Text = "&Timer" Then
            UpTime = 9
            UpTimer.Start()
            QButton.Text = "&Timer Stop"
        ElseIf QButton.Text = "&Timer Stop" Then
            UpTimer.Stop()
            Poll = False
            TimeLabel.Text = ""
            QButton.Text = "&Question"
        End If
    End Sub
    Private Sub AnsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AnsButton.Click
        UpTimer.Stop()
        If AnsButton.Text = "A&nswer" Then
            AnsButton.Text = "A&nswered"
            AnsLabel.Text = CorrectAnswer
            UpdateScoreFlag = True
        ElseIf AnsButton.Text = "A&nswered" Then
            AnsButton.Text = "Una&nswered"
        ElseIf AnsButton.Text = "Una&nswered" Then
            AnsButton.Text = "A&nswered"
        End If
    End Sub
    Private Sub TAButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TAButton.Click
        TeamBtn = "A"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        AnsLabel.Text = ""
        QLabel.Text = ""
        If UpdateScoreFlag = True Then
            If AnsButton.Text = "A&nswered" Then
                A_Score = A_Score + 20
            ElseIf AnsButton.Text = "Una&nswered" Then
                A_Score = A_Score - 10
            End If
            UpdateScoreFlag = False
        End If
    End Sub
    Private Sub TBButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBButton.Click
        TeamBtn = "B"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        AnsLabel.Text = ""
        QLabel.Text = ""
        If UpdateScoreFlag = True Then
            If AnsButton.Text = "A&nswered" Then
                B_Score = B_Score + 20
            ElseIf AnsButton.Text = "Una&nswered" Then
                B_Score = B_Score - 10
            End If
            UpdateScoreFlag = False
        End If
    End Sub
    Private Sub TCButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TCButton.Click
        TeamBtn = "C"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        AnsLabel.Text = ""
        QLabel.Text = ""
        If UpdateScoreFlag = True Then
            If AnsButton.Text = "A&nswered" Then
                C_Score = C_Score + 20
            ElseIf AnsButton.Text = "Una&nswered" Then
                C_Score = C_Score - 10
            End If
            UpdateScoreFlag = False
        End If
    End Sub
    Private Sub TDButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TDButton.Click
        TeamBtn = "D"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        AnsLabel.Text = ""
        QLabel.Text = ""
        If UpdateScoreFlag = True Then
            If AnsButton.Text = "A&nswered" Then
                D_Score = D_Score + 20
            ElseIf AnsButton.Text = "Una&nswered" Then
                D_Score = D_Score - 10
            End If
            UpdateScoreFlag = False
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
    Private Sub rndAudioForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
    Private Sub PollButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PollButton.Click
        TeamBtn = ""
        StartPolling()
        AudWMP.Ctlcontrols.stop()
        UpTimer.Stop()
        TeamBlink()
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
                LabelA.Visible = True
                BlinkTimer.Start()
                LabelA.Location = New Point(17, 24)
            Case "B"
                LabelA.Visible = True
                BlinkTimer.Start()
                LabelA.Location = New Point(17, 64)
            Case "C"
                LabelA.Visible = True
                BlinkTimer.Start()
                LabelA.Location = New Point(17, 104)
            Case "D"
                LabelA.Visible = True
                BlinkTimer.Start()
                LabelA.Location = New Point(17, 144)
        End Select
    End Sub
    Private Sub StopBlink()
        BlinkTimer.Stop()
        LabelA.Visible = False
        LabelA.BackColor = Color.Blue
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

