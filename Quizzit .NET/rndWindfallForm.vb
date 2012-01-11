Public Class rndWindfallForm
    Inherits System.Windows.Forms.Form
    Dim HardQNo As Int16
    Dim MidQNo As Int16
    Dim EasyQNo As Int16
    Dim Level As Int16
    Dim CorrectAnswer As String
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
    Friend WithEvents QButton As System.Windows.Forms.Button
    Friend WithEvents LvlButton As System.Windows.Forms.Button
    Friend WithEvents QLabel As System.Windows.Forms.Label
    Friend WithEvents AnsLabel As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TDButton As System.Windows.Forms.Button
    Friend WithEvents TCButton As System.Windows.Forms.Button
    Friend WithEvents TBButton As System.Windows.Forms.Button
    Friend WithEvents TAButton As System.Windows.Forms.Button
    Friend WithEvents UpTimer As System.Windows.Forms.Timer
    Friend WithEvents TimeLabel As System.Windows.Forms.Label
    Friend WithEvents AxShockwaveFlash2 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents AxShockwaveFlash1 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents AnsButton As System.Windows.Forms.Button
    Friend WithEvents BlinkTimer As System.Windows.Forms.Timer
    Friend WithEvents LabelA As System.Windows.Forms.Label
    Friend WithEvents TrackBar1 As System.Windows.Forms.TrackBar
    Friend WithEvents MainAxShockwaveFlash As AxShockwaveFlashObjects.AxShockwaveFlash
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(rndWindfallForm))
        Me.QButton = New System.Windows.Forms.Button
        Me.LvlButton = New System.Windows.Forms.Button
        Me.QLabel = New System.Windows.Forms.Label
        Me.AnsLabel = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TDButton = New System.Windows.Forms.Button
        Me.TCButton = New System.Windows.Forms.Button
        Me.TBButton = New System.Windows.Forms.Button
        Me.TAButton = New System.Windows.Forms.Button
        Me.LabelA = New System.Windows.Forms.Label
        Me.UpTimer = New System.Windows.Forms.Timer(Me.components)
        Me.TimeLabel = New System.Windows.Forms.Label
        Me.AxShockwaveFlash2 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.AxShockwaveFlash1 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.AnsButton = New System.Windows.Forms.Button
        Me.BlinkTimer = New System.Windows.Forms.Timer(Me.components)
        Me.TrackBar1 = New System.Windows.Forms.TrackBar
        Me.MainAxShockwaveFlash = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.Panel1.SuspendLayout()
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'QButton
        '
        Me.QButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.QButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QButton.Location = New System.Drawing.Point(200, 500)
        Me.QButton.Name = "QButton"
        Me.QButton.Size = New System.Drawing.Size(128, 40)
        Me.QButton.TabIndex = 0
        Me.QButton.Text = "&Start"
        '
        'LvlButton
        '
        Me.LvlButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LvlButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.LvlButton.Location = New System.Drawing.Point(72, 500)
        Me.LvlButton.Name = "LvlButton"
        Me.LvlButton.Size = New System.Drawing.Size(104, 40)
        Me.LvlButton.TabIndex = 1
        Me.LvlButton.Text = "&Level"
        '
        'QLabel
        '
        Me.QLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.QLabel.BackColor = System.Drawing.Color.Gray
        Me.QLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QLabel.Font = New System.Drawing.Font("Verdana", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QLabel.Location = New System.Drawing.Point(25, 50)
        Me.QLabel.Name = "QLabel"
        Me.QLabel.Size = New System.Drawing.Size(450, 300)
        Me.QLabel.TabIndex = 2
        '
        'AnsLabel
        '
        Me.AnsLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AnsLabel.BackColor = System.Drawing.Color.Gray
        Me.AnsLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AnsLabel.Font = New System.Drawing.Font("Verdana", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AnsLabel.Location = New System.Drawing.Point(25, 355)
        Me.AnsLabel.Name = "AnsLabel"
        Me.AnsLabel.Size = New System.Drawing.Size(450, 100)
        Me.AnsLabel.TabIndex = 3
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.TDButton)
        Me.Panel1.Controls.Add(Me.TCButton)
        Me.Panel1.Controls.Add(Me.TBButton)
        Me.Panel1.Controls.Add(Me.TAButton)
        Me.Panel1.Controls.Add(Me.LabelA)
        Me.Panel1.Location = New System.Drawing.Point(500, 365)
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
        Me.LabelA.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.LabelA.Location = New System.Drawing.Point(17, 24)
        Me.LabelA.Name = "LabelA"
        Me.LabelA.Size = New System.Drawing.Size(160, 40)
        Me.LabelA.TabIndex = 4
        Me.LabelA.Visible = False
        '
        'UpTimer
        '
        Me.UpTimer.Interval = 1000
        '
        'TimeLabel
        '
        Me.TimeLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TimeLabel.BackColor = System.Drawing.Color.Black
        Me.TimeLabel.Font = New System.Drawing.Font("Verdana", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TimeLabel.Location = New System.Drawing.Point(500, 10)
        Me.TimeLabel.Name = "TimeLabel"
        Me.TimeLabel.Size = New System.Drawing.Size(200, 125)
        Me.TimeLabel.TabIndex = 16
        Me.TimeLabel.Text = "20"
        Me.TimeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me.AxShockwaveFlash2.TabIndex = 15
        '
        'AxShockwaveFlash1
        '
        Me.AxShockwaveFlash1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.AxShockwaveFlash1.Enabled = True
        Me.AxShockwaveFlash1.Location = New System.Drawing.Point(500, 150)
        Me.AxShockwaveFlash1.Name = "AxShockwaveFlash1"
        Me.AxShockwaveFlash1.OcxState = CType(resources.GetObject("AxShockwaveFlash1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash1.Size = New System.Drawing.Size(200, 200)
        Me.AxShockwaveFlash1.TabIndex = 14
        '
        'AnsButton
        '
        Me.AnsButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.AnsButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AnsButton.Location = New System.Drawing.Point(350, 500)
        Me.AnsButton.Name = "AnsButton"
        Me.AnsButton.Size = New System.Drawing.Size(128, 40)
        Me.AnsButton.TabIndex = 17
        Me.AnsButton.Text = "A&nswer"
        '
        'BlinkTimer
        '
        Me.BlinkTimer.Interval = 400
        '
        'TrackBar1
        '
        Me.TrackBar1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TrackBar1.BackColor = System.Drawing.SystemColors.Control
        Me.TrackBar1.Location = New System.Drawing.Point(16, 472)
        Me.TrackBar1.Maximum = 100
        Me.TrackBar1.Name = "TrackBar1"
        Me.TrackBar1.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.TrackBar1.Size = New System.Drawing.Size(45, 96)
        Me.TrackBar1.TabIndex = 18
        Me.TrackBar1.TickStyle = System.Windows.Forms.TickStyle.None
        '
        'MainAxShockwaveFlash
        '
        Me.MainAxShockwaveFlash.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MainAxShockwaveFlash.Enabled = True
        Me.MainAxShockwaveFlash.Location = New System.Drawing.Point(24, 71)
        Me.MainAxShockwaveFlash.Name = "MainAxShockwaveFlash"
        Me.MainAxShockwaveFlash.OcxState = CType(resources.GetObject("MainAxShockwaveFlash.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainAxShockwaveFlash.Size = New System.Drawing.Size(452, 280)
        Me.MainAxShockwaveFlash.TabIndex = 19
        '
        'rndWindfallForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.DarkViolet
        Me.ClientSize = New System.Drawing.Size(736, 582)
        Me.Controls.Add(Me.MainAxShockwaveFlash)
        Me.Controls.Add(Me.TrackBar1)
        Me.Controls.Add(Me.AnsButton)
        Me.Controls.Add(Me.TimeLabel)
        Me.Controls.Add(Me.AxShockwaveFlash2)
        Me.Controls.Add(Me.AxShockwaveFlash1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.AnsLabel)
        Me.Controls.Add(Me.QLabel)
        Me.Controls.Add(Me.LvlButton)
        Me.Controls.Add(Me.QButton)
        Me.Name = "rndWindfallForm"
        Me.Text = "Windfall"
        Me.Panel1.ResumeLayout(False)
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub rndHRHGForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TeamBtn = ""
        RoundID = "WIN"
        A_Score = 0
        B_Score = 0
        C_Score = 0
        D_Score = 0
        UpTime = 20
        HardQNo = 0
        MidQNo = 0
        EasyQNo = 0
        Level = 1
        Try
            OpenConnection()
            Recordset("SELECT * FROM HRHG ")

        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error Opening Connection/Recordset")
        End Try
        LoadFlash()
    End Sub
    Private Sub LvlButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LvlButton.Click
        Select Case LvlButton.Text
            Case "&Level"
                LvlButton.Text = "&Easy"
                Level = 1
                TrackBar1.BackColor = Color.LimeGreen
                TrackBar1.Value = 15
            Case "&Easy"
                LvlButton.Text = "&Medium"
                Level = 2
                TrackBar1.BackColor = Color.Orange
                TrackBar1.Value = 50
            Case "&Medium"
                LvlButton.Text = "&Hard"
                Level = 3
                TrackBar1.BackColor = Color.Red
                TrackBar1.Value = 100
            Case "&Hard"
                LvlButton.Text = "&Easy"
                Level = 1
                TrackBar1.BackColor = Color.LimeGreen
                TrackBar1.Value = 15
        End Select
    End Sub
    Private Sub QButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QButton.Click
        If QButton.Text = "&Start" Then
            MainAxShockwaveFlash.StopPlay()
            MainAxShockwaveFlash.Visible = False
            QButton.Text = "&Question"
            Exit Sub
        End If
        UpTime = 19
        UpTimer.Start()
        AnsLabel.Text = ""
        AnsButton.Text = "A&nswer"
        Select Case LvlButton.Text
            Case "&Easy"
                If EasyQNo < 5 Then
                    MyRecordset.Move(EasyQNo)
                    QLabel.Text = MyRecordset.Fields(Level).Value
                    CorrectAnswer = MyRecordset.Fields(Level + 3).Value
                    EasyQNo = EasyQNo + 1
                Else
                    QLabel.Text = "No Easy Question Found !"
                End If
            Case "&Medium"
                If MidQNo < 5 Then
                    MyRecordset.Move(MidQNo)
                    QLabel.Text = MyRecordset.Fields(Level).Value
                    CorrectAnswer = MyRecordset.Fields(Level + 3).Value
                    MidQNo = MidQNo + 1
                Else
                    QLabel.Text = "No Medium Question Found !"
                End If
            Case "&Hard"
                If HardQNo < 5 Then
                    MyRecordset.Move(HardQNo)
                    QLabel.Text = MyRecordset.Fields(Level).Value
                    CorrectAnswer = MyRecordset.Fields(Level + 3).Value
                    HardQNo = HardQNo + 1
                Else
                    QLabel.Text = "No Hard Question Found !"
                End If
        End Select
        MyRecordset.MoveFirst()
    End Sub
    Private Sub rndHRHGForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
    Private Sub AnsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AnsButton.Click
        TimeLabel.Text = ""
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
        If AnsButton.Text <> "A&nswer" Then
            QLabel.Text = ""
            AnsLabel.Text = ""
            TrackBar1.Value = 0
            TrackBar1.BackColor = Me.BackColor
            UpdateScore()
        End If
    End Sub
    Private Sub TBButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBButton.Click
        TeamBtn = "B"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If AnsButton.Text <> "A&nswer" Then
            QLabel.Text = ""
            AnsLabel.Text = ""
            TrackBar1.Value = 0
            TrackBar1.BackColor = Me.BackColor
            UpdateScore()
        End If
    End Sub
    Private Sub TCButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TCButton.Click
        TeamBtn = "C"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If AnsButton.Text <> "A&nswer" Then
            QLabel.Text = ""
            AnsLabel.Text = ""
            TrackBar1.Value = 0
            TrackBar1.BackColor = Me.BackColor
            UpdateScore()
        End If
    End Sub
    Private Sub TDButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TDButton.Click
        TeamBtn = "D"
        If BlinkTimer.Enabled Then
            StopBlink()
        Else
            TeamBlink()
        End If
        If AnsButton.Text <> "A&nswer" Then
            QLabel.Text = ""
            AnsLabel.Text = ""
            TrackBar1.Value = 0
            TrackBar1.BackColor = Me.BackColor
            UpdateScore()
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
    Private Sub BlinkTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BlinkTimer.Tick
        If BlinkFlag = True Then
            LabelA.BackColor = Color.Blue
            BlinkFlag = False
        Else
            LabelA.BackColor = Color.Tomato
            BlinkFlag = True
        End If
    End Sub
    Private Sub UpdateScore()
        Select Case TeamBtn
            Case "A"
                If Level <> 1 Then
                    If AnsButton.Text = "A&nswered" Then
                        A_Score = A_Score + 20 * (Level - 1)
                    ElseIf AnsButton.Text = "Una&nswered" Then
                        A_Score = A_Score - 10 * (Level - 1)
                    End If
                ElseIf Level = 1 Then
                    If AnsButton.Text = "A&nswered" Then
                        A_Score = A_Score + 10
                    End If
                End If
            Case "B"
                If Level <> 1 Then
                    If AnsButton.Text = "A&nswered" Then
                        B_Score = B_Score + 20 * (Level - 1)
                    ElseIf AnsButton.Text = "Una&nswered" Then
                        B_Score = B_Score - 10 * (Level - 1)
                    End If
                ElseIf Level = 1 Then
                    If AnsButton.Text = "A&nswered" Then
                        B_Score = B_Score + 10
                    End If
                End If
            Case "C"
                If Level <> 1 Then
                    If AnsButton.Text = "A&nswered" Then
                        C_Score = C_Score + 20 * (Level - 1)
                    ElseIf AnsButton.Text = "Una&nswered" Then
                        C_Score = C_Score - 10 * (Level - 1)
                    End If
                ElseIf Level = 1 Then
                    If AnsButton.Text = "A&nswered" Then
                        C_Score = C_Score + 10
                    End If
                End If
            Case "D"
                If Level <> 1 Then
                    If AnsButton.Text = "A&nswered" Then
                        D_Score = D_Score + 20 * (Level - 1)
                    ElseIf AnsButton.Text = "Una&nswered" Then
                        D_Score = D_Score - 10 * (Level - 1)
                    End If
                ElseIf Level = 1 Then
                    If AnsButton.Text = "A&nswered" Then
                        D_Score = D_Score + 10
                    End If
                End If
        End Select
        LvlButton.Text = "&Level"
        TrackBar1.Value = 0
        TrackBar1.BackColor = Me.BackColor
        AnsButton.Text = "A&nswer"
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
