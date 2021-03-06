Public Class rndHangManForm
    Inherits System.Windows.Forms.Form
    Dim LetLabel(26) As Label
    Dim i As Integer
    Dim j As Integer
    Dim ClueCount As Integer
    Dim XLoc As Integer
    Dim YLoc As Integer
    Dim TmpAns As String
    Dim Chararray() As Char
    Dim AnsArray(50) As Char
    Dim CorrectAnswer As String
    Dim UpdateScoreFlag As Boolean
    Dim PlusScore As Integer
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
    Friend WithEvents AnsLabel As System.Windows.Forms.Label
    Friend WithEvents QButton As System.Windows.Forms.Button
    Friend WithEvents TimeLabel As System.Windows.Forms.Label
    Friend WithEvents AxShockwaveFlash1 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TDButton As System.Windows.Forms.Button
    Friend WithEvents TCButton As System.Windows.Forms.Button
    Friend WithEvents TBButton As System.Windows.Forms.Button
    Friend WithEvents TAButton As System.Windows.Forms.Button
    Friend WithEvents AnsButton As System.Windows.Forms.Button
    Friend WithEvents ClueLabel As System.Windows.Forms.Label
    Friend WithEvents AxShockwaveFlash2 As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents BlinkTimer As System.Windows.Forms.Timer
    Friend WithEvents LabelA As System.Windows.Forms.Label
    Friend WithEvents MainAxShockwaveFlash As AxShockwaveFlashObjects.AxShockwaveFlash
    Friend WithEvents UpTimer As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(rndHangManForm))
        Me.QLabel = New System.Windows.Forms.Label
        Me.AnsLabel = New System.Windows.Forms.Label
        Me.QButton = New System.Windows.Forms.Button
        Me.AnsButton = New System.Windows.Forms.Button
        Me.TimeLabel = New System.Windows.Forms.Label
        Me.AxShockwaveFlash1 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TDButton = New System.Windows.Forms.Button
        Me.TCButton = New System.Windows.Forms.Button
        Me.TBButton = New System.Windows.Forms.Button
        Me.TAButton = New System.Windows.Forms.Button
        Me.LabelA = New System.Windows.Forms.Label
        Me.ClueLabel = New System.Windows.Forms.Label
        Me.AxShockwaveFlash2 = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.BlinkTimer = New System.Windows.Forms.Timer(Me.components)
        Me.MainAxShockwaveFlash = New AxShockwaveFlashObjects.AxShockwaveFlash
        Me.UpTimer = New System.Windows.Forms.Timer(Me.components)
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'QLabel
        '
        Me.QLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.QLabel.BackColor = System.Drawing.Color.Gray
        Me.QLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.QLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QLabel.Font = New System.Drawing.Font("Verdana", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QLabel.Location = New System.Drawing.Point(25, 85)
        Me.QLabel.Name = "QLabel"
        Me.QLabel.Size = New System.Drawing.Size(450, 300)
        Me.QLabel.TabIndex = 0
        '
        'AnsLabel
        '
        Me.AnsLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AnsLabel.BackColor = System.Drawing.Color.Gray
        Me.AnsLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.AnsLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AnsLabel.Font = New System.Drawing.Font("Verdana", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AnsLabel.Location = New System.Drawing.Point(25, 400)
        Me.AnsLabel.Name = "AnsLabel"
        Me.AnsLabel.Size = New System.Drawing.Size(450, 100)
        Me.AnsLabel.TabIndex = 1
        '
        'QButton
        '
        Me.QButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.QButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.QButton.Location = New System.Drawing.Point(104, 641)
        Me.QButton.Name = "QButton"
        Me.QButton.Size = New System.Drawing.Size(136, 32)
        Me.QButton.TabIndex = 2
        Me.QButton.Text = "&Start"
        '
        'AnsButton
        '
        Me.AnsButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.AnsButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.AnsButton.Location = New System.Drawing.Point(280, 641)
        Me.AnsButton.Name = "AnsButton"
        Me.AnsButton.Size = New System.Drawing.Size(128, 32)
        Me.AnsButton.TabIndex = 3
        Me.AnsButton.Text = "A&nswer"
        '
        'TimeLabel
        '
        Me.TimeLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TimeLabel.BackColor = System.Drawing.Color.Black
        Me.TimeLabel.Font = New System.Drawing.Font("Verdana", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TimeLabel.Location = New System.Drawing.Point(514, 10)
        Me.TimeLabel.Name = "TimeLabel"
        Me.TimeLabel.Size = New System.Drawing.Size(200, 125)
        Me.TimeLabel.TabIndex = 16
        Me.TimeLabel.Text = "20"
        Me.TimeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AxShockwaveFlash1
        '
        Me.AxShockwaveFlash1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AxShockwaveFlash1.Enabled = True
        Me.AxShockwaveFlash1.Location = New System.Drawing.Point(514, 144)
        Me.AxShockwaveFlash1.Name = "AxShockwaveFlash1"
        Me.AxShockwaveFlash1.OcxState = CType(resources.GetObject("AxShockwaveFlash1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxShockwaveFlash1.Size = New System.Drawing.Size(198, 192)
        Me.AxShockwaveFlash1.TabIndex = 14
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.TDButton)
        Me.Panel1.Controls.Add(Me.TCButton)
        Me.Panel1.Controls.Add(Me.TBButton)
        Me.Panel1.Controls.Add(Me.TAButton)
        Me.Panel1.Controls.Add(Me.LabelA)
        Me.Panel1.Location = New System.Drawing.Point(514, 461)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 200)
        Me.Panel1.TabIndex = 17
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
        Me.LabelA.Location = New System.Drawing.Point(16, 82)
        Me.LabelA.Name = "LabelA"
        Me.LabelA.Size = New System.Drawing.Size(149, 37)
        Me.LabelA.TabIndex = 5
        Me.LabelA.Visible = False
        '
        'ClueLabel
        '
        Me.ClueLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClueLabel.BackColor = System.Drawing.Color.Black
        Me.ClueLabel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.ClueLabel.Font = New System.Drawing.Font("Verdana", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClueLabel.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClueLabel.Location = New System.Drawing.Point(514, 350)
        Me.ClueLabel.Name = "ClueLabel"
        Me.ClueLabel.Size = New System.Drawing.Size(200, 78)
        Me.ClueLabel.TabIndex = 18
        Me.ClueLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me.AxShockwaveFlash2.TabIndex = 19
        '
        'BlinkTimer
        '
        Me.BlinkTimer.Interval = 400
        '
        'MainAxShockwaveFlash
        '
        Me.MainAxShockwaveFlash.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MainAxShockwaveFlash.Enabled = True
        Me.MainAxShockwaveFlash.Location = New System.Drawing.Point(24, 136)
        Me.MainAxShockwaveFlash.Name = "MainAxShockwaveFlash"
        Me.MainAxShockwaveFlash.OcxState = CType(resources.GetObject("MainAxShockwaveFlash.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainAxShockwaveFlash.Size = New System.Drawing.Size(452, 360)
        Me.MainAxShockwaveFlash.TabIndex = 20
        '
        'UpTimer
        '
        Me.UpTimer.Interval = 1000
        '
        'rndHangManForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(736, 686)
        Me.Controls.Add(Me.MainAxShockwaveFlash)
        Me.Controls.Add(Me.AxShockwaveFlash2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TimeLabel)
        Me.Controls.Add(Me.AxShockwaveFlash1)
        Me.Controls.Add(Me.AnsButton)
        Me.Controls.Add(Me.QButton)
        Me.Controls.Add(Me.AnsLabel)
        Me.Controls.Add(Me.QLabel)
        Me.Controls.Add(Me.ClueLabel)
        Me.Name = "rndHangManForm"
        Me.Text = "HangMan"
        CType(Me.AxShockwaveFlash1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        CType(Me.AxShockwaveFlash2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainAxShockwaveFlash, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub rndHangManForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RoundID = "HAN"
        A_Score = 0
        B_Score = 0
        C_Score = 0
        D_Score = 0
        UpdateScoreFlag = False
        PlusScore = 20
        XLoc = 100
        YLoc = 525
        For i = 0 To 25
            LetLabel(i) = New Label
            If i = 13 Then
                XLoc = 100
                YLoc = 575
            End If
            LetLabel(i).Name = "Label" + Chr(i + 65)
            LetLabel(i).TextAlign = ContentAlignment.MiddleCenter
            LetLabel(i).Font = New Font("Verdana", 22, FontStyle.Bold)
            LetLabel(i).BackColor = Color.Black
            LetLabel(i).Location = New Point(XLoc, YLoc)
            LetLabel(i).Size = New Size(35, 35)
            LetLabel(i).Text = Chr(i + 65)
            AddHandler LetLabel(i).Click, New System.EventHandler(AddressOf LetLabel_Click)
            Me.Controls.Add(LetLabel(i))
            XLoc = XLoc + 50
        Next
        Try
            OpenConnection()
            Recordset("SELECT * FROM FINAL WHERE RndID = '" + RoundID + "'")
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Select Main Round")
        End Try
        LoadFlash()
    End Sub
    Private Sub LetLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If ClueCount < 2 Then
            Dim DumLbl As Label = CType(sender, Label)
            DumLbl.BackColor = Color.Orange
            'MsgBox(DumLbl.Name)
            For i = 0 To Chararray.Length - 1
                If Chararray(i) = DumLbl.Text Then
                    AnsArray(i) = Chararray(i)
                End If
            Next
            AnsLabel.Text = AnsArray
            ClueLabel.Text = "Clue" + Chr(13) & (ClueCount + 1)
            ClueCount = ClueCount + 1
            PlusScore = PlusScore - 5
        End If
    End Sub
    Private Sub QButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QButton.Click
        If QButton.Text = "&Start" Then
            QButton.Text = "&Question"
            MainAxShockwaveFlash.StopPlay()
            MainAxShockwaveFlash.Visible = False
        ElseIf QButton.Text = "&Question" Then
            UpTime = 19
            UpTimer.Start()
            RefreshLetLabel()
            ClueCount = 0
            ClueLabel.Text = ""
            AnsButton.Text = "A&nswer"
            PlusScore = 20
            If Not MyRecordset.EOF Then
                QLabel.Text = MyRecordset.Fields("Answer").Value
                Chararray = MyRecordset.Fields("Question").Value
                CorrectAnswer = Chararray
                MyRecordset.MoveNext()
            Else
                MsgBox("No Questions Found !")
            End If
            'Chararray = TmpAns.ToCharArray()
            Array.Clear(AnsArray, 0, AnsArray.Length - 1)
            For i = 0 To Chararray.Length - 1
                If Chararray(i) <> " " Then
                    AnsArray(i) = "*"
                Else
                    AnsArray(i) = " "
                End If
            Next
            AnsLabel.Text = AnsArray
            'MsgBox(Chararray(TmpAns.Length - 1))

        End If
    End Sub
    Private Sub AnsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AnsButton.Click
        If AnsButton.Text = "A&nswer" Then
            AnsButton.Text = "A&nswered"
            AnsLabel.Text = CorrectAnswer
            Array.Clear(AnsArray, 0, AnsArray.Length - 1)
            UpdateScoreFlag = True
            TimeLabel.Text = "20"
            UpTimer.Stop()
        ElseIf AnsButton.Text = "A&nswered" Then
            AnsButton.Text = "Una&nswered"
            RefreshLetLabel()
        ElseIf AnsButton.Text = "Una&nswered" Then
            AnsButton.Text = "A&nswered"
            RefreshLetLabel()
        End If
    End Sub
    Private Sub RefreshLetLabel()
        For i = 0 To 25
            LetLabel(i).BackColor = Color.Black
        Next
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
        If UpdateScoreFlag = True And AnsButton.Text = "A&nswered" Then
            A_Score = A_Score + PlusScore
            UpdateScoreFlag = False
        End If
        If AnsButton.Text <> "A&nswer" Then
            RefreshLetLabel()
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
        If UpdateScoreFlag = True And AnsButton.Text = "A&nswered" Then
            B_Score = B_Score + PlusScore
            UpdateScoreFlag = False
        End If
        If AnsButton.Text <> "A&nswer" Then
            RefreshLetLabel()
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
        If UpdateScoreFlag = True And AnsButton.Text = "A&nswered" Then
            C_Score = C_Score + PlusScore
            UpdateScoreFlag = False
        End If
        If AnsButton.Text <> "A&nswer" Then
            RefreshLetLabel()
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
        If UpdateScoreFlag = True And AnsButton.Text = "A&nswered" Then
            D_Score = D_Score + PlusScore
            UpdateScoreFlag = False
        End If
        If AnsButton.Text <> "A&nswer" Then
            RefreshLetLabel()
        End If
    End Sub
    Private Sub rndHangManForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        LabelA.Visible = True
        BlinkTimer.Start()
        Select Case TeamBtn
            Case "A"
                LabelA.Location = New Point(17, 24)
            Case "B"
                LabelA.Location = New Point(17, 64)
            Case "C"
                LabelA.Location = New Point(17, 104)
            Case "D"
                LabelA.Location = New Point(17, 144)
        End Select
    End Sub
    Private Sub StopBlink()
        BlinkTimer.Stop()
        LabelA.Visible = False
        LabelA.BackColor = Color.Blue
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

