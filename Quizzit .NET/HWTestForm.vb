Public Class HWTestForm
    Inherits System.Windows.Forms.Form
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkButton As System.Windows.Forms.Button
    Friend WithEvents TAButton As System.Windows.Forms.Button
    Friend WithEvents TBButton As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents LabelA As System.Windows.Forms.Label
    Friend WithEvents BlinkTimer As System.Windows.Forms.Timer
    Friend WithEvents EditButton As System.Windows.Forms.Button
    Friend WithEvents TATextBox As System.Windows.Forms.TextBox
    Friend WithEvents TBTextBox As System.Windows.Forms.TextBox
    Friend WithEvents TCTextBox As System.Windows.Forms.TextBox
    Friend WithEvents TDTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PMonButton As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TATextBox = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.TBButton = New System.Windows.Forms.Button
        Me.TAButton = New System.Windows.Forms.Button
        Me.LabelA = New System.Windows.Forms.Label
        Me.ChkButton = New System.Windows.Forms.Button
        Me.BlinkTimer = New System.Windows.Forms.Timer(Me.components)
        Me.EditButton = New System.Windows.Forms.Button
        Me.TBTextBox = New System.Windows.Forms.TextBox
        Me.TCTextBox = New System.Windows.Forms.TextBox
        Me.TDTextBox = New System.Windows.Forms.TextBox
        Me.PMonButton = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TDTextBox)
        Me.GroupBox1.Controls.Add(Me.TCTextBox)
        Me.GroupBox1.Controls.Add(Me.TBTextBox)
        Me.GroupBox1.Controls.Add(Me.TATextBox)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.TBButton)
        Me.GroupBox1.Controls.Add(Me.TAButton)
        Me.GroupBox1.Controls.Add(Me.LabelA)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(56, 48)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(440, 312)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'TATextBox
        '
        Me.TATextBox.Location = New System.Drawing.Point(392, 48)
        Me.TATextBox.Name = "TATextBox"
        Me.TATextBox.Size = New System.Drawing.Size(40, 20)
        Me.TATextBox.TabIndex = 5
        Me.TATextBox.Text = ""
        Me.TATextBox.Visible = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Black
        Me.Button3.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button3.Location = New System.Drawing.Point(104, 231)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(256, 40)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Team D"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Black
        Me.Button2.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button2.Location = New System.Drawing.Point(104, 168)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(256, 40)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Team C"
        '
        'TBButton
        '
        Me.TBButton.BackColor = System.Drawing.Color.Black
        Me.TBButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.TBButton.Location = New System.Drawing.Point(104, 104)
        Me.TBButton.Name = "TBButton"
        Me.TBButton.Size = New System.Drawing.Size(256, 40)
        Me.TBButton.TabIndex = 1
        Me.TBButton.Text = "Team B"
        '
        'TAButton
        '
        Me.TAButton.BackColor = System.Drawing.Color.Black
        Me.TAButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.TAButton.Location = New System.Drawing.Point(104, 40)
        Me.TAButton.Name = "TAButton"
        Me.TAButton.Size = New System.Drawing.Size(256, 40)
        Me.TAButton.TabIndex = 0
        Me.TAButton.Text = "Team A"
        '
        'LabelA
        '
        Me.LabelA.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.LabelA.Location = New System.Drawing.Point(80, 27)
        Me.LabelA.Name = "LabelA"
        Me.LabelA.Size = New System.Drawing.Size(304, 64)
        Me.LabelA.TabIndex = 4
        Me.LabelA.Visible = False
        '
        'ChkButton
        '
        Me.ChkButton.BackColor = System.Drawing.Color.Black
        Me.ChkButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.ChkButton.Location = New System.Drawing.Point(206, 384)
        Me.ChkButton.Name = "ChkButton"
        Me.ChkButton.Size = New System.Drawing.Size(144, 40)
        Me.ChkButton.TabIndex = 1
        Me.ChkButton.Text = "&Check"
        '
        'BlinkTimer
        '
        Me.BlinkTimer.Interval = 400
        '
        'EditButton
        '
        Me.EditButton.BackColor = System.Drawing.Color.Black
        Me.EditButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.EditButton.Location = New System.Drawing.Point(360, 384)
        Me.EditButton.Name = "EditButton"
        Me.EditButton.Size = New System.Drawing.Size(144, 40)
        Me.EditButton.TabIndex = 2
        Me.EditButton.Text = "Edit"
        '
        'TBTextBox
        '
        Me.TBTextBox.Location = New System.Drawing.Point(392, 112)
        Me.TBTextBox.Name = "TBTextBox"
        Me.TBTextBox.Size = New System.Drawing.Size(40, 20)
        Me.TBTextBox.TabIndex = 6
        Me.TBTextBox.Text = ""
        Me.TBTextBox.Visible = False
        '
        'TCTextBox
        '
        Me.TCTextBox.Location = New System.Drawing.Point(392, 176)
        Me.TCTextBox.Name = "TCTextBox"
        Me.TCTextBox.Size = New System.Drawing.Size(40, 20)
        Me.TCTextBox.TabIndex = 7
        Me.TCTextBox.Text = ""
        Me.TCTextBox.Visible = False
        '
        'TDTextBox
        '
        Me.TDTextBox.Location = New System.Drawing.Point(392, 240)
        Me.TDTextBox.Name = "TDTextBox"
        Me.TDTextBox.Size = New System.Drawing.Size(40, 20)
        Me.TDTextBox.TabIndex = 8
        Me.TDTextBox.Text = ""
        Me.TDTextBox.Visible = False
        '
        'PMonButton
        '
        Me.PMonButton.BackColor = System.Drawing.Color.Black
        Me.PMonButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.PMonButton.Location = New System.Drawing.Point(51, 384)
        Me.PMonButton.Name = "PMonButton"
        Me.PMonButton.Size = New System.Drawing.Size(144, 40)
        Me.PMonButton.TabIndex = 3
        Me.PMonButton.Text = "Paramon"
        '
        'HWTestForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.DarkGray
        Me.ClientSize = New System.Drawing.Size(536, 470)
        Me.Controls.Add(Me.PMonButton)
        Me.Controls.Add(Me.EditButton)
        Me.Controls.Add(Me.ChkButton)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "HWTestForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "HWTestForm"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub ChkButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkButton.Click
        If ChkButton.Text = "&Check" Then
            Poll = True
            ChkButton.Text = "&Stop"
            TeamBtn = ""
            StartPolling()
            TeamBlink()
            TeamBtn = ""
        ElseIf ChkButton.Text = "&Stop" Then
            Poll = False
            StopBlink()
            ChkButton.Text = "&Check"
        End If
    End Sub


    Private Sub TeamBlink()
        Select Case TeamBtn
            Case "A"
                LabelA.Visible = True
                BlinkTimer.Start()
                LabelA.Location = New Point(80, 27)
            Case "B"
                LabelA.Visible = True
                BlinkTimer.Start()
                LabelA.Location = New Point(80, 92)
            Case "C"
                LabelA.Visible = True
                BlinkTimer.Start()
                LabelA.Location = New Point(80, 155)
            Case "D"
                LabelA.Visible = True
                BlinkTimer.Start()
                LabelA.Location = New Point(80, 219)
        End Select
    End Sub
    Private Sub StopBlink()
        BlinkTimer.Stop()
        LabelA.Visible = False
        LabelA.BackColor = Color.DeepSkyBlue
    End Sub

    Private Sub BlinkTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BlinkTimer.Tick
        If BlinkFlag = True Then
            LabelA.BackColor = Color.DeepSkyBlue
            BlinkFlag = False
        Else
            LabelA.BackColor = Color.Orange
            BlinkFlag = True
        End If
    End Sub

    Private Sub EditButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditButton.Click
        If EditButton.Text = "Edit" Then
            EditButton.Text = "Ok"
            TATextBox.Visible = True
            TBTextBox.Visible = True
            TCTextBox.Visible = True
            TDTextBox.Visible = True
            TATextBox.Text = TA_Code
            TBTextBox.Text = TB_Code
            TCTextBox.Text = TC_Code
            TDTextBox.Text = TD_Code
        Else
            TA_Code = TATextBox.Text
            TB_Code = TBTextBox.Text
            TC_Code = TCTextBox.Text
            TD_Code = TDTextBox.Text
            TATextBox.Visible = False
            TBTextBox.Visible = False
            TCTextBox.Visible = False
            TDTextBox.Visible = False
            EditButton.Text = "Edit"
        End If
    End Sub

    Private Sub PMonButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PMonButton.Click
        Process.Start(ResourcePath + "\ParaMon.exe")
    End Sub
End Class
