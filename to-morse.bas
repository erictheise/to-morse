Function ToMorse(Optional x) As Variant
	Dim s$
	Dim i As Integer
	
	If IsMissing(x) Then
		s = "No Argument was passed"
	ElseIf NOT IsArray(x) Then
		For i = 1 to Len(x)
			char = Mid(x, i, 1)
			Select Case LCase(char)
				Case "a"
					s = s & "▄ ▄▄▄"
				Case "b"
					s = s & "▄▄▄ ▄ ▄ ▄"
				Case "c"
					s = s & "▄▄▄ ▄ ▄▄▄ ▄"
				Case "d"
					s = s & "▄▄▄ ▄ ▄"
				Case "e"
					s = s & "▄"
				Case "f"
					s = s & "▄ ▄ ▄▄▄ ▄"
				Case "g"
					s = s & "▄▄▄ ▄▄▄ ▄"
				Case "h"
					s = s & "▄ ▄ ▄ ▄"
				Case "i"
					s = s & "▄ ▄"
				Case "j"
					s = s & "▄ ▄▄▄ ▄▄▄ ▄▄▄"
				Case "k"
					s = s & "▄▄▄ ▄ ▄▄▄"
				Case "l"
					s = s & "▄ ▄▄▄ ▄ ▄"
				Case "m"
					s = s & "▄▄▄ ▄▄▄"
				Case "n"
					s = s & "▄▄▄ ▄"
				Case "o"
					s = s & "▄▄▄ ▄▄▄ ▄▄▄"
				Case "p"
					s = s & "▄ ▄▄▄ ▄▄▄ ▄"
				Case "q"
					s = s & "▄▄▄ ▄▄▄ ▄ ▄▄▄"
				Case "r"
					s = s & "▄ ▄▄▄ ▄"
				Case "s"
					s = s & "▄ ▄ ▄"
				Case "t"
					s = s & "▄▄▄"
				Case "u"
					s = s & "▄ ▄ ▄▄▄"
				Case "v"
					s = s & "▄ ▄ ▄ ▄▄▄"
				Case "w"
					s = s & "▄ ▄▄▄ ▄▄▄"
				Case "x"
					s = s & "▄▄▄ ▄ ▄ ▄▄▄"
				Case "y"
					s = s & "▄▄▄ ▄ ▄▄▄ ▄▄▄"
				Case "z"
					s = s & "▄▄▄ ▄▄▄ ▄ ▄"
				Case "1"
					s = s & "▄ ▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄"
				Case "2"
					s = s & "▄ ▄ ▄▄▄ ▄▄▄ ▄▄▄"
				Case "3"
					s = s & "▄ ▄ ▄ ▄▄▄ ▄▄▄"
				Case "4"
					s = s & "▄ ▄ ▄ ▄ ▄▄▄"
				Case "5"
					s = s & "▄ ▄ ▄ ▄ ▄"
				Case "6"
					s = s & "▄▄▄ ▄ ▄ ▄ ▄"
				Case "7"
					s = s & "▄▄▄ ▄▄▄ ▄ ▄ ▄"
				Case "8"
					s = s & "▄▄▄ ▄▄▄ ▄▄▄ ▄ ▄"
				Case "9"
					s = s & "▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄ ▄"
				Case "0"
					s = s & "▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄"
				Case " "
					s = s & "       "
			End Select
			s = s & "     "
		Next
	Else
		s = s & "Argument is an array"
 	End If
	ToMorse = s
End Function
