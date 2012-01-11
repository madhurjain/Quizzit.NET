Module PublicModule
    Dim basePath As String() = Environment.CurrentDirectory.Split(New Char() {"\"c})
    Public ResourcePath As String = basePath(0) + "\" + basePath(1) + "\Resources"
    Public Poll As Boolean
    Public A_Score As Integer
    Public B_Score As Integer
    Public C_Score As Integer
    Public D_Score As Integer
    Public UpTime As Integer
    Public TempScore As Integer
    Public MainRound As String
    Public SubRound As String
    Public RoundID As String
    Public Qz As String
    Public TeamBtn As String
    Public CurrTeam As String
    Public Stats2 As String
    Public Stats3 As String
    Public UpdateOnClose As Boolean

End Module
