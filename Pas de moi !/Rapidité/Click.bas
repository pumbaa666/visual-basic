Attribute VB_Name = "Module1"
Option Explicit
Dim Lettre As String
Public Function DefLettre()

Dim NbLettre As Long

Randomize
NbLettre = Int((26 * Rnd) + 1)

If NbLettre = 1 Then Lettre = "a"
If NbLettre = 2 Then Lettre = "b"
If NbLettre = 3 Then Lettre = "c"
If NbLettre = 4 Then Lettre = "d"
If NbLettre = 5 Then Lettre = "e"
If NbLettre = 6 Then Lettre = "f"
If NbLettre = 7 Then Lettre = "g"
If NbLettre = 8 Then Lettre = "h"
If NbLettre = 9 Then Lettre = "i"
If NbLettre = 10 Then Lettre = "j"
If NbLettre = 11 Then Lettre = "k"
If NbLettre = 12 Then Lettre = "l"
If NbLettre = 13 Then Lettre = "m"
If NbLettre = 14 Then Lettre = "n"
If NbLettre = 15 Then Lettre = "o"
If NbLettre = 16 Then Lettre = "p"
If NbLettre = 17 Then Lettre = "q"
If NbLettre = 18 Then Lettre = "r"
If NbLettre = 19 Then Lettre = "s"
If NbLettre = 20 Then Lettre = "t"
If NbLettre = 21 Then Lettre = "u"
If NbLettre = 22 Then Lettre = "v"
If NbLettre = 23 Then Lettre = "w"
If NbLettre = 24 Then Lettre = "x"
If NbLettre = 25 Then Lettre = "y"
If NbLettre = 26 Then Lettre = "z"

Click.Label25.Caption = Lettre

End Function
