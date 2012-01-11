Public Class rndImageForm
    Inherits System.Windows.Forms.Form
    Dim UpdateScoreFlag As Boolean
    Dim Qquery As String
    Dim tmp As Array
    Dim CorrectAnswer As String
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
    Friend WithEvents TimeLabel As System.Windows.Forms.Label
    Friend WithEvents AxShockwaveFlash2 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents AxShockwaveFlash1 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents BtnPanel As System.Windows.Forms.Panel
    Friend WithEvents TDButton As System.Windows.Forms.Button
    Friend WithEvents TCButton As System.Windows.Forms.Button
    Friend WithEvents TBButton As System.Windows.Forms.Button
    Friend WithEvents TAButton As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents QLabel As System.Windows.Forms.Label
    Friend WithEvents AnsLabel As System.Windows.Forms.Label
    Friend WithEvents QButton As System.Windows.Forms.Button
    Friend WithEvents AnsButton As System.Windows.Forms.Button
    Friend WithEvents ImgPictureBox As System.Windows.Forms.PictureBox
    Friend WithEvents UpTimer As System.Windows.Forms.Timer
    Friend WithEvents BlinkTimer As System.Windows.Forms.Timer
    Friend WithEvents LabelA As System.Windows.Forms.Label
    Friend WithEvents MainAxShockwaveFlash As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents ImageTimer As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(rndImageForm))
        Me.TimeLabel = New System.Windows.Forms.Label
        Me.AxShockwaveFlash2 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.AxShockwaveFlash1 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.BtnPanel = New System.Windows.Forms.Panel
        Me.TDButton = New System.Windows.Forms.Button
        Me.TCButton = New System.Windows.Forms.Button
        Me.TBButton = New System.Windows.Forms.Button
        Me.TAButton = New System.Windows.Forms.Button
        Me.LabelA = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ImgPictureBox = New System.Windows.Forms.PictureBox
        Me.QLabel = New System.Windows.Forms.Label
        Me.AnsLabel = New System.Windows.Forms.Label
        Me.QButton = New System.Windows.Forms.Button
        Me.AnsButton = New System.Windows.Forms.Button
        Me.UpTimer = New System.Windows.Forms.Timer(Me.components)
        Me.BlinkTimer = New System.Windows.Forms.Timer(Me.components)
        Me.MainAxShockwaveFlash = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.ImageTimer = New System.Windows.Forms.Timer(Me.components)
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BtnPanel.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TimeLabel
        '
        Me.TimeLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TimeLabel.BackColor = System.Drawing.Color.Black
        Me.TimeLabel.Font = New System.Drawing.Font("Verdana", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TimeLabel.Location = New System.Drawing.Point(750, 10)
        Me.TimeLabel.Name = "TimeLabel"
        Me.TimeLabel.Size = New System.Drawing.Size(200, 125)
        Me.TimeLabel.TabIndex = 17
        Me.TimeLabel.Text = "10"
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
        Me.AxShockwaveFlash2.Size = New System.Drawing.Size(610, 60)
        Me.AxShockwaveFlash2.TabIndex = 16
        '
        'AxShockwaveFlash1
        '
        Me.AxShockwaveFlash1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.AxShockwaveFlash1.Enabled = True
        Me.AxShockwaveFlash1.Location = New System.Drawing.Point(750, 200)
        Me.AxShockwaveFlash1.Name = "AxShockwaveFlash1"
        Me.AxShockwaveFlash1.OcxState = CType(resources.GetObject("AxShockwaveFlash1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash1.Size = New System.Drawing.Size(200, 200)
        Me.AxShockwaveFlash1.TabIndex = 15
        '
        'BtnPanel
        '
        Me.BtnPanel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnPanel.Controls.Add(Me.TDButton)
        Me.BtnPanel.Controls.Add(Me.TCButton)
        Me.BtnPanel.Controls.Add(Me.TBButton)
        Me.BtnPanel.Controls.Add(Me.TAButton)
        Me.BtnPanel.Controls.Add(Me.LabelA)
        Me.BtnPanel.Location = New System.Drawing.Point(750, 450)
        Me.BtnPanel.Name = "BtnPanel"
        Me.BtnPanel.Size = New System.Drawing.Size(200, 200)
        Me.BtnPanel.TabIndex = 14
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
        Me.Panel1.Controls.Add(Me.ImgPictureBox)
        Me.Panel1.Location = New System.Drawing.Point(32, 165)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(552, 384)
        Me.Panel1.TabIndex = 18
        '
        'ImgPictureBox
        '
        Me.ImgPictureBox.BackColor = System.Drawing.Color.White
        Me.ImgPictureBox.Image = CType(resources.GetObject("ImgPictureBox.Image"), System.Drawing.Image)
        Me.ImgPictureBox.Location = New System.Drawing.Point(34, 30)
        Me.ImgPictureBox.Name = "ImgPictureBox"
        Me.ImgPictureBox.Size = New System.Drawing.Size(488, 328)
        Me.ImgPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.ImgPictureBox.TabIndex = 0
        Me.ImgPictureBox.TabStop = False
        '
        'QLabel
        '
        Me.QLabel.BackColor = System.Drawing.Color.Gray
        Me.QLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QLabel.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.QLabel.Location = New System.Drawing.Point(32, 96)
        Me.QLabel.Name = "QLabel"
        Me.QLabel.Size = New System.Drawing.Size(552, 64)
        Me.QLabel.TabIndex = 19
        Me.QLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AnsLabel
        '
        Me.AnsLabel.BackColor = System.Drawing.Color.Gray
        Me.AnsLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AnsLabel.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.AnsLabel.Location = New System.Drawing.Point(32, 555)
        Me.AnsLabel.Name = "AnsLabel"
        Me.AnsLabel.Size = New System.Drawing.Size(552, 48)
        Me.AnsLabel.TabIndex = 20
        '
        'QButton
        '
        Me.QButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.QButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QButton.Location = New System.Drawing.Point(150, 650)
        Me.QButton.Name = "QButton"
        Me.QButton.Size = New System.Drawing.Size(128, 40)
        Me.QButton.TabIndex = 22
        Me.QButton.Text = "&Start"
        '
        'AnsButton
        '
        Me.AnsButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.AnsButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AnsButton.Location = New System.Drawing.Point(350, 650)
        Me.AnsButton.Name = "AnsButton"
        Me.AnsButton.Size = New System.Drawing.Size(128, 40)
        Me.AnsButton.TabIndex = 23
        Me.AnsButton.Text = "A&nswer"
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
        Me.MainAxShockwaveFlash.Location = New System.Drawing.Point(32, 165)
        Me.MainAxShockwaveFlash.Name = "MainAxShockwaveFlash"
        Me.MainAxShockwaveFlash.OcxState = CType(resources.GetObject("MainAxShockwaveFlash.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainAxShockwaveFlash.Size = New System.Drawing.Size(552, 384)
        Me.MainAxShockwaveFlash.TabIndex = 24
        '
        'ImageTimer
        '
        Me.ImageTimer.Interval = 1000
        '
        'rndImageForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(992, 670)
        Me.Controls.Add(Me.MainAxShockwaveFlash)
        Me.Controls.Add(Me.AnsButton)
        Me.Controls.Add(Me.QButton)
        Me.Controls.Add(Me.AnsLabel)
        Me.Controls.Add(Me.QLabel)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TimeLabel)
        Me.Controls.Add(Me.AxShockwaveFlash2)
        Me.Controls.Add(Me.AxShockwaveFlash1)
        Me.Controls.Add(Me.BtnPanel)
        Me.Name = "rndImageForm"
        Me.Text = "rndImageForm"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BtnPanel.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub rndImageForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TeamBtn = ""
        RoundID = "IMA"
        A_Score = 0
        B_Score = 0
        C_Score = 0
        D_Score = 0
        UpdateScoreFlag = False
        Qquery = "SELECT * FROM " + MainRound + " WHERE RndID='" + RoundID + "'"
        Try
            OpenConnection()
            Recordset(Qquery)
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error Opening Connection/Recordset")
        End Try
        LoadFlash()
    End Sub
    Private Sub QButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QButton.Click
        If QButton.Text = "&Start" Then
            MainAxShockwaveFlash.StopPlay()
            MainAxShockwaveFlash.Visible = False
            QButton.Text = "&Question"
            Exit Sub
        End If
        TeamBtn = ""
        ImageTimer.Start()
        StartPolling()
        If TeamBtn <> "" Then
            'MsgBox("Team " + TeamBtn + Chr(13) + "Please Release The Buzzer!!!", MsgBoxStyle.Information)
            TeamBlink()
            Exit Sub
        End If
        If Not MyRecordset.EOF Then
            QLabel.Text = MyRecordset.Fields("Question").Value
            AnsLabel.Text = ""
            TimeLabel.Text = "10"
            AnsButton.Text = "A&nswer"
            tmp = Split(MyRecordset.Fields("Answer").Value(), ";", -1)
            CorrectAnswer = tmp(0)
            Try
                ImgPictureBox.Image = Image.FromFile(ResourcePath + tmp(1))
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            MyRecordset.MoveNext()
            UpTime = 9
            UpTimer.Start()
            TeamBtn = ""
            StartPolling()
            TeamBlink()
            UpTimer.Stop()
        Else
            QLabel.Text = "No Questions Found !"
            AnsLabel.Text = ""
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
    Private Sub UpTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpTimer.Tick
        If UpTime > 0 Then
            TimeLabel.Text = UpTime
            UpTime = UpTime - 1
        ElseIf UpTime = 0 Then
            TimeLabel.Text = UpTime
            UpTime = UpTime - 1
            PlaySnd(ResourcePath + "\Tones\timer1.wav")
        Else
            UpTimer.Stop()
            Poll = False
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
        ImgPictureBox.Image = Image.FromFile(ResourcePath + "\image\blank.jpg")
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
        ImgPictureBox.Image = Image.FromFile(ResourcePath + "\image\blank.jpg")
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
        ImgPictureBox.Image = Image.FromFile(ResourcePath + "\image\blank.jpg")
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
        ImgPictureBox.Image = Image.FromFile(ResourcePath + "\image\blank.jpg")
        If UpdateScoreFlag = True Then
            If AnsButton.Text = "A&nswered" Then
                D_Score = D_Score + 20
            ElseIf AnsButton.Text = "Una&nswered" Then
                D_Score = D_Score - 10
            End If
            UpdateScoreFlag = False
        End If
    End Sub
    Private Sub rndImageForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        MainAxShockwaveFlash.Play()
        AxShockwaveFlash1.Movie = ResourcePath + "\Extras\Main.swf"
        AxShockwaveFlash2.Movie = ResourcePath + "\Extras\" + RoundID + "_banner.swf"
        AxShockwaveFlash1.Play()
        AxShockwaveFlash2.Play()
        AxShockwaveFlash1.Loop = True
        AxShockwaveFlash2.Loop = True
    End Sub

    Private Sub ImageTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageTimer.Tick
        ImageTimer.Stop()
        Poll = False
    End Sub
End Class
