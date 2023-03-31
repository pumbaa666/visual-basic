Attribute VB_Name = "clsTrait"
Private Type Spot
 X As Single
 Y As Single
End Type
Private Type QueueDeComete
   lstPoint() As Spot
   Lng As Byte
End Type
Public lstQueue() As QueueDeComete

Public Type Trait
 M As Long   ' X minimum
 N As Long   ' X maximum
 Y1 As Long   ' Y1
 Y2 As Long   ' Y2
 a As Single ' Coef directeur
 b As Single ' Constante dans l'équation de droite "ax+b"
End Type

Public lstT() As Trait 'Liste de données des Traits


Public Sub DessinerTrait(ByRef Efface As Byte)
 If Efface Then frmSimulation.Plan.Cls
 Dim z As Byte
 For z = 1 To frmSimulation.nbrTrait + (frmSimulation.nbrTrait = 255)
  frmSimulation.Plan.Line (lstT(z).M, lstT(z).Y1)-(lstT(z).N, lstT(z).Y2), vbBlack
 Next z
 If frmSimulation.nbrTrait = 255 Then frmSimulation.Plan.Line (lstT(z).M, lstT(z).Y1)-(lstT(z).N, lstT(z).Y2), vbBlack
 frmSimulation.Plan.PSet (frmSimulation.X, frmSimulation.Y), vbRed
End Sub

Public Sub DrawComete(ByRef L, ByRef P)
 Call DessinerTrait(1)
 Dim z As Byte, t As Integer, i As Byte
 z = P
 Do
     i = i + 1
     z = z + 1 + L * (z + 1 > L)
     For t = 2 To lstQueue(z).lstPoint(0).X
         frmSimulation.Plan.Line (lstQueue(z).lstPoint(t - 1).X, lstQueue(z).lstPoint(t - 1).Y)- _
         (lstQueue(z).lstPoint(t).X, lstQueue(z).lstPoint(t).Y), RGB(255 - Int(Val(255 / L * i) + 0.5), 255 - Int(Val(255 / L * i) + 0.5), 255 - Int(Val(255 / L * i) + 0.5) + 100)
     Next t
 Loop Until z = P
 
End Sub




Public Sub CreerBord(ByRef nbrTrait)

 ReDim Preserve lstT(nbrTrait)

 frmSimulation.Plan.Line (0, 0)-(0, 480), vbBlack
 lstT(nbrTrait - 3).a = 0
 lstT(nbrTrait - 3).b = 0
 lstT(nbrTrait - 3).M = 0
 lstT(nbrTrait - 3).N = 0
 lstT(nbrTrait - 3).Y1 = 0
 lstT(nbrTrait - 3).Y2 = 480

 frmSimulation.Plan.Line (0, 480)-(600, 480), vbBlack
 lstT(nbrTrait - 2).a = 0
 lstT(nbrTrait - 2).b = 480
 lstT(nbrTrait - 2).M = 0
 lstT(nbrTrait - 2).N = 600
 lstT(2).Y1 = 480
 lstT(2).Y2 = 480
 
 frmSimulation.Plan.Line (600, 0)-(600, 480), vbBlack
 lstT(nbrTrait - 1).a = 0
 lstT(nbrTrait - 1).b = 0
 lstT(nbrTrait - 1).M = 600
 lstT(nbrTrait - 1).N = 600
 lstT(nbrTrait - 1).Y1 = 0
 lstT(nbrTrait - 1).Y2 = 480
 
 frmSimulation.Plan.Line (0, 0)-(600, 0), vbBlack
 lstT(nbrTrait).a = 0
 lstT(nbrTrait).b = 0
 lstT(nbrTrait).M = 0
 lstT(nbrTrait).N = 600
 lstT(nbrTrait).Y1 = 0
 lstT(nbrTrait).Y2 = 0
End Sub
