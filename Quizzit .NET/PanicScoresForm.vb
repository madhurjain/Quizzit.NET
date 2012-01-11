Public Class PanicScoresForm
    Inherits System.Windows.Forms.Form
    Dim ascr, bscr, cscr, dscr, i As Integer

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox13 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox17 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox21 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox25 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox29 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox14 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox15 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox16 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox18 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox19 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox20 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox22 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox23 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox24 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox26 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox27 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox28 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox30 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox31 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox32 As System.Windows.Forms.TextBox
    Friend WithEvents GetButton As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents UpdateButton As System.Windows.Forms.Button
    Friend WithEvents TextBox33 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox34 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox35 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox36 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.TextBox9 = New System.Windows.Forms.TextBox
        Me.TextBox13 = New System.Windows.Forms.TextBox
        Me.TextBox17 = New System.Windows.Forms.TextBox
        Me.TextBox21 = New System.Windows.Forms.TextBox
        Me.TextBox25 = New System.Windows.Forms.TextBox
        Me.TextBox29 = New System.Windows.Forms.TextBox
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.TextBox8 = New System.Windows.Forms.TextBox
        Me.TextBox10 = New System.Windows.Forms.TextBox
        Me.TextBox11 = New System.Windows.Forms.TextBox
        Me.TextBox12 = New System.Windows.Forms.TextBox
        Me.TextBox14 = New System.Windows.Forms.TextBox
        Me.TextBox15 = New System.Windows.Forms.TextBox
        Me.TextBox16 = New System.Windows.Forms.TextBox
        Me.TextBox18 = New System.Windows.Forms.TextBox
        Me.TextBox19 = New System.Windows.Forms.TextBox
        Me.TextBox20 = New System.Windows.Forms.TextBox
        Me.TextBox22 = New System.Windows.Forms.TextBox
        Me.TextBox23 = New System.Windows.Forms.TextBox
        Me.TextBox24 = New System.Windows.Forms.TextBox
        Me.TextBox26 = New System.Windows.Forms.TextBox
        Me.TextBox27 = New System.Windows.Forms.TextBox
        Me.TextBox28 = New System.Windows.Forms.TextBox
        Me.TextBox30 = New System.Windows.Forms.TextBox
        Me.TextBox31 = New System.Windows.Forms.TextBox
        Me.TextBox32 = New System.Windows.Forms.TextBox
        Me.GetButton = New System.Windows.Forms.Button
        Me.UpdateButton = New System.Windows.Forms.Button
        Me.TextBox33 = New System.Windows.Forms.TextBox
        Me.TextBox34 = New System.Windows.Forms.TextBox
        Me.TextBox35 = New System.Windows.Forms.TextBox
        Me.TextBox36 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.DarkGray
        Me.Label1.Location = New System.Drawing.Point(30, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 32)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "General"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.DarkGray
        Me.Label2.Location = New System.Drawing.Point(30, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 32)
        Me.Label2.TabIndex = 36
        Me.Label2.Text = "Hangman"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.DarkGray
        Me.Label3.Location = New System.Drawing.Point(30, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 32)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Image"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.DarkGray
        Me.Label4.Location = New System.Drawing.Point(30, 176)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 32)
        Me.Label4.TabIndex = 38
        Me.Label4.Text = "Audio"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.DarkGray
        Me.Label8.Location = New System.Drawing.Point(30, 336)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 32)
        Me.Label8.TabIndex = 42
        Me.Label8.Text = "Windfall"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.DarkGray
        Me.Label7.Location = New System.Drawing.Point(30, 296)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 32)
        Me.Label7.TabIndex = 41
        Me.Label7.Text = "Rapidfire"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.DarkGray
        Me.Label6.Location = New System.Drawing.Point(30, 256)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 32)
        Me.Label6.TabIndex = 40
        Me.Label6.Text = "Crossword"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.DarkGray
        Me.Label5.Location = New System.Drawing.Point(30, 216)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 32)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "Video"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.LightGray
        Me.Label9.Location = New System.Drawing.Point(504, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(96, 32)
        Me.Label9.TabIndex = 46
        Me.Label9.Text = "Team D"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.LightGray
        Me.Label10.Location = New System.Drawing.Point(392, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 32)
        Me.Label10.TabIndex = 45
        Me.Label10.Text = "Team C"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.LightGray
        Me.Label11.Location = New System.Drawing.Point(280, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(96, 32)
        Me.Label11.TabIndex = 44
        Me.Label11.Text = "Team B"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.LightGray
        Me.Label12.Location = New System.Drawing.Point(168, 8)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(96, 32)
        Me.Label12.TabIndex = 43
        Me.Label12.Text = "Team A"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(200, 60)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(40, 20)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = ""
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(300, 60)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(40, 20)
        Me.TextBox2.TabIndex = 2
        Me.TextBox2.Text = ""
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(415, 61)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(40, 20)
        Me.TextBox3.TabIndex = 3
        Me.TextBox3.Text = ""
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(525, 61)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(40, 20)
        Me.TextBox4.TabIndex = 4
        Me.TextBox4.Text = ""
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(200, 100)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(40, 20)
        Me.TextBox5.TabIndex = 5
        Me.TextBox5.Text = ""
        '
        'TextBox9
        '
        Me.TextBox9.Location = New System.Drawing.Point(200, 140)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(40, 20)
        Me.TextBox9.TabIndex = 9
        Me.TextBox9.Text = ""
        '
        'TextBox13
        '
        Me.TextBox13.Location = New System.Drawing.Point(200, 185)
        Me.TextBox13.Name = "TextBox13"
        Me.TextBox13.Size = New System.Drawing.Size(40, 20)
        Me.TextBox13.TabIndex = 13
        Me.TextBox13.Text = ""
        '
        'TextBox17
        '
        Me.TextBox17.Location = New System.Drawing.Point(200, 225)
        Me.TextBox17.Name = "TextBox17"
        Me.TextBox17.Size = New System.Drawing.Size(40, 20)
        Me.TextBox17.TabIndex = 17
        Me.TextBox17.Text = ""
        '
        'TextBox21
        '
        Me.TextBox21.Location = New System.Drawing.Point(200, 265)
        Me.TextBox21.Name = "TextBox21"
        Me.TextBox21.Size = New System.Drawing.Size(40, 20)
        Me.TextBox21.TabIndex = 21
        Me.TextBox21.Text = ""
        '
        'TextBox25
        '
        Me.TextBox25.Location = New System.Drawing.Point(200, 305)
        Me.TextBox25.Name = "TextBox25"
        Me.TextBox25.Size = New System.Drawing.Size(40, 20)
        Me.TextBox25.TabIndex = 25
        Me.TextBox25.Text = ""
        '
        'TextBox29
        '
        Me.TextBox29.Location = New System.Drawing.Point(200, 345)
        Me.TextBox29.Name = "TextBox29"
        Me.TextBox29.Size = New System.Drawing.Size(40, 20)
        Me.TextBox29.TabIndex = 29
        Me.TextBox29.Text = ""
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(300, 100)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(40, 20)
        Me.TextBox6.TabIndex = 6
        Me.TextBox6.Text = ""
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(416, 104)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(40, 20)
        Me.TextBox7.TabIndex = 7
        Me.TextBox7.Text = ""
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(525, 104)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(40, 20)
        Me.TextBox8.TabIndex = 8
        Me.TextBox8.Text = ""
        '
        'TextBox10
        '
        Me.TextBox10.Location = New System.Drawing.Point(300, 140)
        Me.TextBox10.Name = "TextBox10"
        Me.TextBox10.Size = New System.Drawing.Size(40, 20)
        Me.TextBox10.TabIndex = 10
        Me.TextBox10.Text = ""
        '
        'TextBox11
        '
        Me.TextBox11.Location = New System.Drawing.Point(416, 144)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.Size = New System.Drawing.Size(40, 20)
        Me.TextBox11.TabIndex = 11
        Me.TextBox11.Text = ""
        '
        'TextBox12
        '
        Me.TextBox12.Location = New System.Drawing.Point(525, 144)
        Me.TextBox12.Name = "TextBox12"
        Me.TextBox12.Size = New System.Drawing.Size(40, 20)
        Me.TextBox12.TabIndex = 12
        Me.TextBox12.Text = ""
        '
        'TextBox14
        '
        Me.TextBox14.Location = New System.Drawing.Point(300, 185)
        Me.TextBox14.Name = "TextBox14"
        Me.TextBox14.Size = New System.Drawing.Size(40, 20)
        Me.TextBox14.TabIndex = 14
        Me.TextBox14.Text = ""
        '
        'TextBox15
        '
        Me.TextBox15.Location = New System.Drawing.Point(416, 184)
        Me.TextBox15.Name = "TextBox15"
        Me.TextBox15.Size = New System.Drawing.Size(40, 20)
        Me.TextBox15.TabIndex = 15
        Me.TextBox15.Text = ""
        '
        'TextBox16
        '
        Me.TextBox16.Location = New System.Drawing.Point(525, 184)
        Me.TextBox16.Name = "TextBox16"
        Me.TextBox16.Size = New System.Drawing.Size(40, 20)
        Me.TextBox16.TabIndex = 16
        Me.TextBox16.Text = ""
        '
        'TextBox18
        '
        Me.TextBox18.Location = New System.Drawing.Point(300, 225)
        Me.TextBox18.Name = "TextBox18"
        Me.TextBox18.Size = New System.Drawing.Size(40, 20)
        Me.TextBox18.TabIndex = 18
        Me.TextBox18.Text = ""
        '
        'TextBox19
        '
        Me.TextBox19.Location = New System.Drawing.Point(416, 224)
        Me.TextBox19.Name = "TextBox19"
        Me.TextBox19.Size = New System.Drawing.Size(40, 20)
        Me.TextBox19.TabIndex = 19
        Me.TextBox19.Text = ""
        '
        'TextBox20
        '
        Me.TextBox20.Location = New System.Drawing.Point(525, 224)
        Me.TextBox20.Name = "TextBox20"
        Me.TextBox20.Size = New System.Drawing.Size(40, 20)
        Me.TextBox20.TabIndex = 20
        Me.TextBox20.Text = ""
        '
        'TextBox22
        '
        Me.TextBox22.Location = New System.Drawing.Point(300, 265)
        Me.TextBox22.Name = "TextBox22"
        Me.TextBox22.Size = New System.Drawing.Size(40, 20)
        Me.TextBox22.TabIndex = 22
        Me.TextBox22.Text = ""
        '
        'TextBox23
        '
        Me.TextBox23.Location = New System.Drawing.Point(416, 263)
        Me.TextBox23.Name = "TextBox23"
        Me.TextBox23.Size = New System.Drawing.Size(40, 20)
        Me.TextBox23.TabIndex = 23
        Me.TextBox23.Text = ""
        '
        'TextBox24
        '
        Me.TextBox24.Location = New System.Drawing.Point(525, 262)
        Me.TextBox24.Name = "TextBox24"
        Me.TextBox24.Size = New System.Drawing.Size(40, 20)
        Me.TextBox24.TabIndex = 24
        Me.TextBox24.Text = ""
        '
        'TextBox26
        '
        Me.TextBox26.Location = New System.Drawing.Point(300, 305)
        Me.TextBox26.Name = "TextBox26"
        Me.TextBox26.Size = New System.Drawing.Size(40, 20)
        Me.TextBox26.TabIndex = 26
        Me.TextBox26.Text = ""
        '
        'TextBox27
        '
        Me.TextBox27.Location = New System.Drawing.Point(416, 304)
        Me.TextBox27.Name = "TextBox27"
        Me.TextBox27.Size = New System.Drawing.Size(40, 20)
        Me.TextBox27.TabIndex = 27
        Me.TextBox27.Text = ""
        '
        'TextBox28
        '
        Me.TextBox28.Location = New System.Drawing.Point(525, 301)
        Me.TextBox28.Name = "TextBox28"
        Me.TextBox28.Size = New System.Drawing.Size(40, 20)
        Me.TextBox28.TabIndex = 28
        Me.TextBox28.Text = ""
        '
        'TextBox30
        '
        Me.TextBox30.Location = New System.Drawing.Point(300, 345)
        Me.TextBox30.Name = "TextBox30"
        Me.TextBox30.Size = New System.Drawing.Size(40, 20)
        Me.TextBox30.TabIndex = 30
        Me.TextBox30.Text = ""
        '
        'TextBox31
        '
        Me.TextBox31.Location = New System.Drawing.Point(415, 342)
        Me.TextBox31.Name = "TextBox31"
        Me.TextBox31.Size = New System.Drawing.Size(40, 20)
        Me.TextBox31.TabIndex = 31
        Me.TextBox31.Text = ""
        '
        'TextBox32
        '
        Me.TextBox32.Location = New System.Drawing.Point(525, 342)
        Me.TextBox32.Name = "TextBox32"
        Me.TextBox32.Size = New System.Drawing.Size(40, 20)
        Me.TextBox32.TabIndex = 32
        Me.TextBox32.Text = ""
        '
        'GetButton
        '
        Me.GetButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GetButton.Location = New System.Drawing.Point(216, 441)
        Me.GetButton.Name = "GetButton"
        Me.GetButton.Size = New System.Drawing.Size(120, 40)
        Me.GetButton.TabIndex = 33
        Me.GetButton.Text = "Get Scores"
        '
        'UpdateButton
        '
        Me.UpdateButton.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.UpdateButton.Location = New System.Drawing.Point(376, 440)
        Me.UpdateButton.Name = "UpdateButton"
        Me.UpdateButton.Size = New System.Drawing.Size(136, 40)
        Me.UpdateButton.TabIndex = 34
        Me.UpdateButton.Text = "Update Score"
        '
        'TextBox33
        '
        Me.TextBox33.Location = New System.Drawing.Point(200, 400)
        Me.TextBox33.Name = "TextBox33"
        Me.TextBox33.Size = New System.Drawing.Size(40, 20)
        Me.TextBox33.TabIndex = 47
        Me.TextBox33.Text = ""
        '
        'TextBox34
        '
        Me.TextBox34.Location = New System.Drawing.Point(300, 400)
        Me.TextBox34.Name = "TextBox34"
        Me.TextBox34.Size = New System.Drawing.Size(40, 20)
        Me.TextBox34.TabIndex = 48
        Me.TextBox34.Text = ""
        '
        'TextBox35
        '
        Me.TextBox35.Location = New System.Drawing.Point(415, 400)
        Me.TextBox35.Name = "TextBox35"
        Me.TextBox35.Size = New System.Drawing.Size(40, 20)
        Me.TextBox35.TabIndex = 49
        Me.TextBox35.Text = ""
        '
        'TextBox36
        '
        Me.TextBox36.Location = New System.Drawing.Point(525, 400)
        Me.TextBox36.Name = "TextBox36"
        Me.TextBox36.Size = New System.Drawing.Size(40, 20)
        Me.TextBox36.TabIndex = 50
        Me.TextBox36.Text = ""
        '
        'PanicScoresForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(728, 502)
        Me.Controls.Add(Me.TextBox36)
        Me.Controls.Add(Me.TextBox35)
        Me.Controls.Add(Me.TextBox34)
        Me.Controls.Add(Me.TextBox33)
        Me.Controls.Add(Me.UpdateButton)
        Me.Controls.Add(Me.GetButton)
        Me.Controls.Add(Me.TextBox32)
        Me.Controls.Add(Me.TextBox31)
        Me.Controls.Add(Me.TextBox30)
        Me.Controls.Add(Me.TextBox28)
        Me.Controls.Add(Me.TextBox27)
        Me.Controls.Add(Me.TextBox26)
        Me.Controls.Add(Me.TextBox24)
        Me.Controls.Add(Me.TextBox23)
        Me.Controls.Add(Me.TextBox22)
        Me.Controls.Add(Me.TextBox20)
        Me.Controls.Add(Me.TextBox19)
        Me.Controls.Add(Me.TextBox18)
        Me.Controls.Add(Me.TextBox16)
        Me.Controls.Add(Me.TextBox15)
        Me.Controls.Add(Me.TextBox14)
        Me.Controls.Add(Me.TextBox12)
        Me.Controls.Add(Me.TextBox11)
        Me.Controls.Add(Me.TextBox10)
        Me.Controls.Add(Me.TextBox8)
        Me.Controls.Add(Me.TextBox7)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.TextBox29)
        Me.Controls.Add(Me.TextBox25)
        Me.Controls.Add(Me.TextBox21)
        Me.Controls.Add(Me.TextBox17)
        Me.Controls.Add(Me.TextBox13)
        Me.Controls.Add(Me.TextBox9)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "PanicScoresForm"
        Me.Text = "PanicScoresForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub PanicScoresForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ascr = 0
        bscr = 0
        cscr = 0
        dscr = 0
        Try
            OpenConnection()
            OpenScoreSet("SELECT * FROM Scores")
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error Opening Connection/Recordset")
        End Try
    End Sub
    Private Sub PanicScoresForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            CloseScoreSet()
            CloseConnection()
        Catch ex As Exception
            MsgBox(ex.Message + Chr(13) + "Error In Closing Connection/Recordset")
        End Try
    End Sub
#Region "Crap Coding..."
    Private Sub GetButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetButton.Click
        ScoreSet.MoveFirst()
        Label1.Text = ScoreSet.Fields(0).Value
        TextBox1.Text = ScoreSet.Fields(1).Value
        TextBox2.Text = ScoreSet.Fields(2).Value
        TextBox3.Text = ScoreSet.Fields(3).Value
        TextBox4.Text = ScoreSet.Fields(4).Value
        ScoreSet.MoveNext()
        Label2.Text = ScoreSet.Fields(0).Value
        TextBox5.Text = ScoreSet.Fields(1).Value
        TextBox6.Text = ScoreSet.Fields(2).Value
        TextBox7.Text = ScoreSet.Fields(3).Value
        TextBox8.Text = ScoreSet.Fields(4).Value
        ScoreSet.MoveNext()
        Label3.Text = ScoreSet.Fields(0).Value
        TextBox9.Text = ScoreSet.Fields(1).Value
        TextBox10.Text = ScoreSet.Fields(2).Value
        TextBox11.Text = ScoreSet.Fields(3).Value
        TextBox12.Text = ScoreSet.Fields(4).Value
        ScoreSet.MoveNext()
        Label4.Text = ScoreSet.Fields(0).Value
        TextBox13.Text = ScoreSet.Fields(1).Value
        TextBox14.Text = ScoreSet.Fields(2).Value
        TextBox15.Text = ScoreSet.Fields(3).Value
        TextBox16.Text = ScoreSet.Fields(4).Value
        ScoreSet.MoveNext()
        Label5.Text = ScoreSet.Fields(0).Value
        TextBox17.Text = ScoreSet.Fields(1).Value
        TextBox18.Text = ScoreSet.Fields(2).Value
        TextBox19.Text = ScoreSet.Fields(3).Value
        TextBox20.Text = ScoreSet.Fields(4).Value
        ScoreSet.MoveNext()
        Label6.Text = ScoreSet.Fields(0).Value
        TextBox21.Text = ScoreSet.Fields(1).Value
        TextBox22.Text = ScoreSet.Fields(2).Value
        TextBox23.Text = ScoreSet.Fields(3).Value
        TextBox24.Text = ScoreSet.Fields(4).Value
        ScoreSet.MoveNext()
        Label7.Text = ScoreSet.Fields(0).Value
        TextBox25.Text = ScoreSet.Fields(1).Value
        TextBox26.Text = ScoreSet.Fields(2).Value
        TextBox27.Text = ScoreSet.Fields(3).Value
        TextBox28.Text = ScoreSet.Fields(4).Value
        ScoreSet.MoveNext()
        Label8.Text = ScoreSet.Fields(0).Value
        TextBox29.Text = ScoreSet.Fields(1).Value
        TextBox30.Text = ScoreSet.Fields(2).Value
        TextBox31.Text = ScoreSet.Fields(3).Value
        TextBox32.Text = ScoreSet.Fields(4).Value
        ScoreSet.MoveFirst()
        For i = 0 To ScoreSet.RecordCount - 1
            ascr = ascr + CInt(ScoreSet.Fields(1).Value)
            bscr = bscr + CInt(ScoreSet.Fields(2).Value)
            cscr = cscr + CInt(ScoreSet.Fields(3).Value)
            dscr = dscr + CInt(ScoreSet.Fields(4).Value)
            ScoreSet.MoveNext()
        Next
        TextBox33.Text = ascr
        TextBox34.Text = bscr
        TextBox35.Text = cscr
        TextBox36.Text = dscr
    End Sub
    Private Sub UpdateButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateButton.Click
        ScoreSet.MoveFirst()
        ScoreSet.Fields(1).Value = TextBox1.Text
        ScoreSet.Fields(2).Value = TextBox2.Text
        ScoreSet.Fields(3).Value = TextBox3.Text
        ScoreSet.Fields(4).Value = TextBox4.Text
        ScoreSet.MoveNext()
        ScoreSet.Fields(1).Value = TextBox5.Text
        ScoreSet.Fields(2).Value = TextBox6.Text
        ScoreSet.Fields(3).Value = TextBox7.Text
        ScoreSet.Fields(4).Value = TextBox8.Text
        ScoreSet.MoveNext()
        ScoreSet.Fields(1).Value = TextBox9.Text
        ScoreSet.Fields(2).Value = TextBox10.Text
        ScoreSet.Fields(3).Value = TextBox11.Text
        ScoreSet.Fields(4).Value = TextBox12.Text
        ScoreSet.MoveNext()
        ScoreSet.Fields(1).Value = TextBox13.Text
        ScoreSet.Fields(2).Value = TextBox14.Text
        ScoreSet.Fields(3).Value = TextBox15.Text
        ScoreSet.Fields(4).Value = TextBox16.Text
        ScoreSet.MoveNext()
        ScoreSet.Fields(1).Value = TextBox17.Text
        ScoreSet.Fields(2).Value = TextBox18.Text
        ScoreSet.Fields(3).Value = TextBox19.Text
        ScoreSet.Fields(4).Value = TextBox20.Text
        ScoreSet.MoveNext()
        ScoreSet.Fields(1).Value = TextBox21.Text
        ScoreSet.Fields(2).Value = TextBox22.Text
        ScoreSet.Fields(3).Value = TextBox23.Text
        ScoreSet.Fields(4).Value = TextBox24.Text
        ScoreSet.MoveNext()
        ScoreSet.Fields(1).Value = TextBox25.Text
        ScoreSet.Fields(2).Value = TextBox26.Text
        ScoreSet.Fields(3).Value = TextBox27.Text
        ScoreSet.Fields(4).Value = TextBox28.Text
        ScoreSet.MoveNext()
        ScoreSet.Fields(1).Value = TextBox29.Text
        ScoreSet.Fields(2).Value = TextBox30.Text
        ScoreSet.Fields(3).Value = TextBox31.Text
        ScoreSet.Fields(4).Value = TextBox32.Text
        Try
            ScoreSet.UpdateBatch(ADODB.AffectEnum.adAffectAll)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region
End Class
