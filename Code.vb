Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim RC As String
        Dim LL As Int16
        Dim a As String
        Dim L As Int64
        Dim n As String
        Dim RCL As String
        Dim NNC As Int16
        Dim M As String
        RC$ = ""
        LL = 1
        a$ = Clipboard.GetText
        L = Len(a$)
        For i = L To 1 Step -1
            n = Mid(a$, i, 1)
            Select Case n
                Case "A"
                    RC$ = RC$ + "T"
                Case "a"
                    RC$ = RC$ + "t"
                Case "T"
                    RC$ = RC$ + "A"
                Case "t"
                    RC$ = RC$ + "a"
                Case "G"
                    RC$ = RC$ + "C"
                Case "g"
                    RC$ = RC$ + "c"
                Case "C"
                    RC$ = RC$ + "G"
                Case "c"
                    RC$ = RC$ + "g"


                Case "M"
                    RC$ = RC$ + "K"
                Case "R"
                    RC$ = RC$ + "Y"
                Case "W"
                    RC$ = RC$ + "W"
                Case "S"
                    RC$ = RC$ + "S"
                Case "Y"
                    RC$ = RC$ + "R"
                Case "K"
                    RC$ = RC$ + "M"
                Case "V"
                    RC$ = RC$ + "B"
                Case "H"
                    RC$ = RC$ + "D"
                Case "D"
                    RC$ = RC$ + "H"
                Case "B"
                    RC$ = RC$ + "V"
                Case "N"
                    RC$ = RC$ + "N"

                Case "m"
                    RC$ = RC$ + "k"
                Case "r"
                    RC$ = RC$ + "y"
                Case "w"
                    RC$ = RC$ + "w"
                Case "s"
                    RC$ = RC$ + "s"
                Case "y"
                    RC$ = RC$ + "r"
                Case "k"
                    RC$ = RC$ + "m"
                Case "v"
                    RC$ = RC$ + "b"
                Case "h"
                    RC$ = RC$ + "d"
                Case "d"
                    RC$ = RC$ + "h"
                Case "b"
                    RC$ = RC$ + "v"
                Case "n"
                    RC$ = RC$ + "n"

                Case Strings.Chr(10)
                Case Strings.Chr(13)
                    LL = i - 1 : NNC = 0
                Case Else
                    NNC = NNC + 1
            End Select
        Next i
        RCL$ = ""
        LL = LL - NNC
        If LL <= 0 Then LL = 1
        For j = 1 To Len(RC$)
            M$ = Mid(RC$, j, 1)
            RCL$ = RCL$ + M
            If j / LL = Int(j / LL) And LL > 1 Then RCL$ = RCL$ + Strings.Chr(13) + Strings.Chr(10)
        Next j
        Clipboard.Clear()
        Clipboard.SetText(RCL$)
    End Sub
End Class
