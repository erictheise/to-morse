REM to-morse.bas is a LibreOffice macro that translates text in a spreadsheet
REM cell to its representation in Morse Code.

REM https://github.com/erictheise/to-morse is its home.
REM Copyright (C) 2021 Eric Theise <erictheise@gmail.com>

REM This program is free software: you can redistribute it and/or modify
REM it under the terms of the GNU Affero General Public License as published
REM by the Free Software Foundation, either version 3 of the License, or
REM (at your option) any later version.

REM This program is distributed in the hope that it will be useful,
REM but WITHOUT ANY WARRANTY; without even the implied warranty of
REM MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
REM GNU Affero General Public License for more details.

REM You should have received a copy of the GNU Affero General Public License
REM along with this program.  If not, see <https://www.gnu.org/licenses/>.

Function ToMorse(Optional x) As Variant
  Dim s$
  Dim i As Integer

  For i = 1 to Len(x)
    char = Mid(x, i, 1)
    Select Case LCase(char)
      Case "a"
        s = s & "▄ ▄▄▄"
      Case "b"
        s = s & "▄▄▄ ▄ ▄ ▄"
      Case "c"
        s = s & "▄▄▄ ▄ ▄▄▄ ▄"
      Case "d"
        s = s & "▄▄▄ ▄ ▄"
      Case "e"
        s = s & "▄"
      Case "f"
        s = s & "▄ ▄ ▄▄▄ ▄"
      Case "g"
        s = s & "▄▄▄ ▄▄▄ ▄"
      Case "h"
        s = s & "▄ ▄ ▄ ▄"
      Case "i"
        s = s & "▄ ▄"
      Case "j"
        s = s & "▄ ▄▄▄ ▄▄▄ ▄▄▄"
      Case "k"
        s = s & "▄▄▄ ▄ ▄▄▄"
      Case "l"
        s = s & "▄ ▄▄▄ ▄ ▄"
      Case "m"
        s = s & "▄▄▄ ▄▄▄"
      Case "n"
        s = s & "▄▄▄ ▄"
      Case "o"
        s = s & "▄▄▄ ▄▄▄ ▄▄▄"
      Case "p"
        s = s & "▄ ▄▄▄ ▄▄▄ ▄"
      Case "q"
        s = s & "▄▄▄ ▄▄▄ ▄ ▄▄▄"
      Case "r"
        s = s & "▄ ▄▄▄ ▄"
      Case "s"
        s = s & "▄ ▄ ▄"
      Case "t"
        s = s & "▄▄▄"
      Case "u"
        s = s & "▄ ▄ ▄▄▄"
      Case "v"
        s = s & "▄ ▄ ▄ ▄▄▄"
      Case "w"
        s = s & "▄ ▄▄▄ ▄▄▄"
      Case "x"
        s = s & "▄▄▄ ▄ ▄ ▄▄▄"
      Case "y"
        s = s & "▄▄▄ ▄ ▄▄▄ ▄▄▄"
      Case "z"
        s = s & "▄▄▄ ▄▄▄ ▄ ▄"
      Case "1"
        s = s & "▄ ▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄"
      Case "2"
        s = s & "▄ ▄ ▄▄▄ ▄▄▄ ▄▄▄"
      Case "3"
        s = s & "▄ ▄ ▄ ▄▄▄ ▄▄▄"
      Case "4"
        s = s & "▄ ▄ ▄ ▄ ▄▄▄"
      Case "5"
        s = s & "▄ ▄ ▄ ▄ ▄"
      Case "6"
        s = s & "▄▄▄ ▄ ▄ ▄ ▄"
      Case "7"
        s = s & "▄▄▄ ▄▄▄ ▄ ▄ ▄"
      Case "8"
        s = s & "▄▄▄ ▄▄▄ ▄▄▄ ▄ ▄"
      Case "9"
        s = s & "▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄ ▄"
      Case "0"
        s = s & "▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄"
      Case "."
        s = s & "▄ ▄▄▄ ▄ ▄▄▄ ▄ ▄▄▄"
      Case ","
        s = s & "▄▄▄ ▄▄▄ ▄ ▄ ▄▄▄ ▄▄▄"
      REM Here we need to test for straight and "smart"
      REM opening & closing double quotation marks.
      Case """", "“", "”"
        s = s & "▄ ▄▄▄ ▄ ▄ ▄▄▄ ▄"
      Case "?"
        s = s & "▄ ▄ ▄▄▄ ▄▄▄ ▄ ▄"
      Case "'"
        s = s & "▄ ▄▄▄ ▄▄▄ ▄▄▄ ▄▄▄ ▄"
      Case "!"
        s = s & "▄▄▄ ▄ ▄▄▄ ▄ ▄▄▄ ▄▄▄"
      Case " "
        s = s & "       "
    End Select
    s = s & "     "
  Next
  ToMorse = s
End Function
