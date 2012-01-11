Public Class rndRapidFireForm
    Inherits System.Windows.Forms.Form
    Dim cnt As Integer
    Dim Score As Integer
    Dim Bonus As Integer
    Dim BlinkFlag As Boolean
    Dim FlashFlag As Boolean


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
    Friend WithEvents AxShockwaveFlash2 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents AxShockwaveFlash1 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents BtnPanel As System.Windows.Forms.Panel
    Friend WithEvents TDButton As System.Windows.Forms.Button
    Friend WithEvents TCButton As System.Windows.Forms.Button
    Friend WithEvents TBButton As System.Windows.Forms.Button
    Friend WithEvents TAButton As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TimeLabel As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents TickTimer As System.Windows.Forms.Timer
    Friend WithEvents SSButton As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ScoreTextBox As System.Windows.Forms.TextBox
    Friend WithEvents OkButton As System.Windows.Forms.Button
    Friend WithEvents BlinkTimer As System.Windows.Forms.Timer
    Friend WithEvents LabelA As System.Windows.Forms.Label
    Friend WithEvents MainAxShockwaveFlash As AxShockwaveFlashObjects.AxShockwaveFlash
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(rndRapidFireForm))
        Me.AxShockwaveFlash2 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.AxShockwaveFlash1 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.BtnPanel = New System.Windows.Forms.Panel
        Me.TDButton = New System.Windows.Forms.Button
        Me.TCButton = New System.Windows.Forms.Button
        Me.TBButton = New System.Windows.Forms.Button
        Me.TAButton = New System.Windows.Forms.Button
        Me.LabelA = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.TimeLabel = New System.Windows.Forms.Label
        Me.TickTimer = New System.Windows.Forms.Timer(Me.components)
        Me.SSButton = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.OkButton = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.ScoreTextBox = New System.Windows.Forms.TextBox
        Me.BlinkTimer = New System.Windows.Forms.Timer(Me.components)
        Me.MainAxShockwaveFlash = New AxShockwaveFlashObjects.AxShockwaveFlash
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BtnPanel.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxShockwaveFlash2
        '
        Me.AxShockwaveFlash2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AxShockwaveFlash2.Enabled = True
        Me.AxShockwaveFlash2.Location = New System.Drawing.Point(50, 25)
        Me.AxShockwaveFlash2.Name = "AxShockwaveFlash2"
        Me.AxShockwaveFlash2.OcxState = CType(resources.GetObject("AxShockwaveFlash2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash2.Size = New System.Drawing.Size(450, 40)
        Me.AxShockwaveFlash2.TabIndex = 15
        '
        'AxShockwaveFlash1
        '
        Me.AxShockwaveFlash1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AxShockwaveFlash1.Enabled = True
        Me.AxShockwaveFlash1.Location = New System.Drawing.Point(550, 25)
        Me.AxShockwaveFlash1.Name = "AxShockwaveFlash1"
        Me.AxShockwaveFlash1.OcxState = CType(resources.GetObject("AxShockwaveFlash1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash1.Size = New System.Drawing.Size(200, 200)
        Me.AxShockwaveFlash1.TabIndex = 14
        '
        'BtnPanel
        '
        Me.BtnPanel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnPanel.Controls.Add(Me.TDButton)
        Me.BtnPanel.Controls.Add(Me.TCButton)
        Me.BtnPanel.Controls.Add(Me.TBButton)
        Me.BtnPanel.Controls.Add(Me.TAButton)
        Me.BtnPanel.Controls.Add(Me.LabelA)
        Me.BtnPanel.Location = New System.Drawing.Point(546, 453)
        Me.BtnPanel.Name = "BtnPanel"
        Me.BtnPanel.Size = New System.Drawing.Size(200, 200)
        Me.BtnPanel.TabIndex = 13
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
        Me.LabelA.Location = New System.Drawing.Point(16, 82)
        Me.LabelA.Name = "LabelA"
        Me.LabelA.Size = New System.Drawing.Size(168, 37)
        Me.LabelA.TabIndex = 5
        Me.LabelA.Visible = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gray
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.TimeLabel)
        Me.Panel1.Location = New System.Drawing.Point(150, 224)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(354, 344)
        Me.Panel1.TabIndex = 16
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(10, 275)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(335, 20)
        Me.ProgressBar1.TabIndex = 1
        Me.ProgressBar1.Visible = False
        '
        'TimeLabel
        '
        Me.TimeLabel.BackColor = System.Drawing.Color.Black
        Me.TimeLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.TimeLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 72.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TimeLabel.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.TimeLabel.Location = New System.Drawing.Point(80, 120)
        Me.TimeLabel.Name = "TimeLabel"
        Me.TimeLabel.Size = New System.Drawing.Size(208, 112)
        Me.TimeLabel.TabIndex = 0
        Me.TimeLabel.Text = "60"
        Me.TimeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TickTimer
        '
        Me.TickTimer.Interval = 1000
        '
        'SSButton
        '
        Me.SSButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.SSButton.Location = New System.Drawing.Point(256, 608)
        Me.SSButton.Name = "SSButton"
        Me.SSButton.Size = New System.Drawing.Size(128, 40)
        Me.SSButton.TabIndex = 17
        Me.SSButton.Text = "&Start"
        '
        'Panel2
        '
        Me.Panel2.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Panel2.BackColor = System.Drawing.Color.Black
        Me.Panel2.Controls.Add(Me.OkButton)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.ScoreTextBox)
        Me.Panel2.Location = New System.Drawing.Point(552, 256)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(200, 168)
        Me.Panel2.TabIndex = 18
        '
        'OkButton
        '
        Me.OkButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.OkButton.Location = New System.Drawing.Point(72, 121)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(60, 30)
        Me.OkButton.TabIndex = 2
        Me.OkButton.Text = "&Ok"
        '
        'Label1
        '
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(184, 32)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Right Answers"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ScoreTextBox
        '
        Me.ScoreTextBox.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.ScoreTextBox.Location = New System.Drawing.Point(72, 80)
        Me.ScoreTextBox.MaxLength = 2
        Me.ScoreTextBox.Name = "ScoreTextBox"
        Me.ScoreTextBox.Size = New System.Drawing.Size(56, 20)
        Me.ScoreTextBox.TabIndex = 0
        Me.ScoreTextBox.Text = ""
        '
        'BlinkTimer
        '
        Me.BlinkTimer.Interval = 400
        '
        'MainAxShockwaveFlash
        '
        Me.MainAxShockwaveFlash.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MainAxShockwaveFlash.Enabled = True
        Me.MainAxShockwaveFlash.Location = New System.Drawing.Point(90, 241)
        Me.MainAxShockwaveFlash.Name = "MainAxShockwaveFlash"
        Me.MainAxShockwaveFlash.OcxState = CType(resources.GetObject("MainAxShockwaveFlash.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainAxShockwaveFlash.Size = New System.Drawing.Size(452, 285)
        Me.MainAxShockwaveFlash.TabIndex = 19
        '
        'rndRapidFireForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(816, 670)
        Me.Controls.Add(Me.MainAxShockwaveFlash)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.SSButton)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.AxShockwaveFlash2)
        Me.Controls.Add(Me.AxShockwaveFlash1)
        Me.Controls.Add(Me.BtnPanel)
        Me.Name = "rndRapidFireForm"
        Me.Text = "RapidFire"
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BtnPanel.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub rndRapidFireForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TeamBtn = ""
        RoundID = "RAP"
        A_Score = 0
        B_Score = 0
        C_Score = 0
        D_Score = 0
        Bonus = 0
        cnt = 60
        FlashFlag = True
        LoadFlash()
        MainAxShockwaveFlash.Location = New Point(75, 195)
        MainAxShockwaveFlash.Size = New Size(350, 400)
    End Sub
    Private Sub SSButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SSButton.Click
        If FlashFlag = True Then
            MainAxShockwaveFlash.StopPlay()
            MainAxShockwaveFlash.Visible = False
            FlashFlag = False
            Exit Sub
        End If
        If SSButton.Text = "&Start" Then
            TickTimer.Start()
            SSButton.Text = "&Stop"
            ProgressBar1.Visible = True
        ElseIf SSButton.Text = "&Stop" Then
            TickTimer.Stop()
            cnt = 60
            TimeLabel.Text = cnt
            ProgressBar1.Value = 0
            SSButton.Text = "&Start"
            ProgressBar1.Visible = False
        End If
    End Sub
    Private Sub TickTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TickTimer.Tick
        If cnt > 0 Then
            TimeLabel.Text = cnt
            ProgressBar1.Value = (60 - cnt) * 1.6
            cnt = cnt - 1
        ElseIf cnt = 0 Then
            TimeLabel.Text = cnt
            PlaySnd(ResourcePath + "\Tones\timer1.wav")
            SSButton.Text = "&Start"
            TickTimer.Stop()
            cnt = 60
            ProgressBar1.Value = 0
        End If
    End Sub
    Private Sub OkButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OkButton.Click
        Bonus = 0
        Try
            Score = CInt(ScoreTextBox.Text) * 10
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If Score >= 50 And Score <= 70 Then
            Bonus = 10
        ElseIf Score >= 80 And Score <= 90 Then
            Bonus = 20
        ElseIf Score = 100 Then
            Bonus = 30
        End If
        Score = Score + Bonus
        Select Case TeamBtn
            Case "A"
                A_Score = Score
            Case "B"
                B_Score = Score
            Case "C"
                C_Score = Score
            Case "D"
                D_Score = Score
        End Select
        ScoreTextBox.Clear()
        StopBlink()
    End Sub
    Private Sub TAButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TAButton.Click
        TeamBtn = "A"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
    End Sub
    Private Sub TBButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBButton.Click
        TeamBtn = "B"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
    End Sub
    Private Sub TCButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TCButton.Click
        TeamBtn = "C"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
    End Sub
    Private Sub TDButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TDButton.Click
        TeamBtn = "D"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
    End Sub
    Private Sub rndRapidFireForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            OpenConnection()
            OpenScoreSet("SELECT * FROM Scores WHERE SubRound = '" + RoundID + "'")
            ScoreSet.Fields(1).Value = A_Score
            ScoreSet.Fields(2).Value = B_Score
            ScoreSet.Fields(3).Value = C_Score
            ScoreSet.Fields(4).Value = D_Score
            ScoreSet.UpdateBatch(ADODB.AffectEnum.adAffectAll)
            CloseScoreSet()
            CloseConnection()
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error Updating Scores !!")
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
