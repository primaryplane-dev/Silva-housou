Option Explicit

Public Function 日付変換(ByVal i_変換前 As Variant) As String
    日付変換 = ""

    If IsNull(i_変換前) Then Exit Function
    If Val(i_変換前) = 0 Then Exit Function
    
    日付変換 = Format(i_変換前, "0000/00/00")

End Function




