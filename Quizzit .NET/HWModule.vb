Module HWModule
    Public Declare Sub PortOut Lib "IO.DLL" (ByVal Port As Integer, ByVal Data As Byte)
    Public Declare Sub PortWordOut Lib "IO.DLL" (ByVal Port As Integer, ByVal Data As Integer)
    Public Declare Sub PortDWordOut Lib "IO.DLL" (ByVal Port As Integer, ByVal Data As Long)
    Public Declare Function PortIn Lib "IO.DLL" (ByVal Port As Integer) As Byte
    Public Declare Function PortWordIn Lib "IO.DLL" (ByVal Port As Integer) As Integer
    Public Declare Function PortDWordIn Lib "IO.DLL" (ByVal Port As Integer) As Long
    Public Declare Sub SetPortBit Lib "IO.DLL" (ByVal Port As Integer, ByVal Bit As Byte)
    Public Declare Sub ClrPortBit Lib "IO.DLL" (ByVal Port As Integer, ByVal Bit As Byte)
    Public Declare Sub NotPortBit Lib "IO.DLL" (ByVal Port As Integer, ByVal Bit As Byte)
    Public Declare Function GetPortBit Lib "IO.DLL" (ByVal Port As Integer, ByVal Bit As Byte) As Boolean
    Public Declare Function RightPortShift Lib "IO.DLL" (ByVal Port As Integer, ByVal Val As Boolean) As Boolean
    Public Declare Function LeftPortShift Lib "IO.DLL" (ByVal Port As Integer, ByVal Val As Boolean) As Boolean
    Public Declare Function IsDriverInstalled Lib "IO.DLL" () As Boolean
    ''''
    Public TA_Code As Integer = 119
    Public TB_Code As Integer = 95
    Public TC_Code As Integer = 255
    Public TD_Code As Integer = 111
    ''''

    Public Sub StartPolling()
        'If IsDriverInstalled = True Then
        Poll = True
        Dim BzTeam As New BuzzTeamForm
        While (Poll)
            Application.DoEvents()
            Select Case PortIn(CInt(&H379))
                Case TA_Code
                    TeamBtn = "A"
                    Poll = False
                    BzTeam.Show()
                    'PlaySnd(ResourcePath + "\Tones\glass.wav")
                Case TB_Code
                    TeamBtn = "B"
                    Poll = False
                    BzTeam.Show()
                    'PlaySnd(ResourcePath + "\Tones\glass.wav")
                Case TC_Code
                    TeamBtn = "C"
                    Poll = False
                    BzTeam.Show()
                    'PlaySnd(ResourcePath + "\Tones\glass.wav")
                Case TD_Code
                    TeamBtn = "D"
                    Poll = False
                    BzTeam.Show()
                    'PlaySnd(ResourcePath + "\Tones\glass.wav")
            End Select
        End While
        'Else
        ' MsgBox("Hardware Driver Error!" + Chr(13) + "IO.DLL Not Found")
        'End If
    End Sub

End Module
