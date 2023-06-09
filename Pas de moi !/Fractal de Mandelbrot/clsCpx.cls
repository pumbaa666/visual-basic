VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCpx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PI As Double = 3.1415926535898

'Classe qui permet de g�rer les nombres complexes.
'Loin d'�tre exhaustive, elle permet un minimum d'op�rations...

Private mReal As Double 'Partie r�elle
Private mImag As Double 'Partie imaginaire

'***********************************************
'Property Let/Get de base

Public Property Get RealPart() As Double
RealPart = mReal
End Property

Public Property Let RealPart(nReal As Double)
mReal = nReal
End Property

Public Property Get ImagPart() As Double
ImagPart = mImag
End Property

Public Property Let ImagPart(nImag As Double)
mImag = nImag
End Property
'*****************************************************

Public Property Get ModuleCrr() As Double   'Retourne le module carr�
ModuleCrr = mReal ^ 2 + mImag ^ 2
End Property

Public Property Get Module() As Double  'Et le module (racine carr�e positive de la fonction d'en haut)
Module = Sqr(ModuleCrr)
End Property

Public Property Get Angle() As Double   'L'argument, appel� 'angle' suite � son interpr�tation g�om�trique
'C'est un truc que m'a refil� ma TI92+, je sais pas si �a marche mais �a m'a jamais pos� de probl�mes :)
Angle = Sgn(mImag) * PI / 2 - Atn(mReal / mImag)
End Property

Public Sub Negate()
'Inverse le complexe
mReal = -mReal
mImag = -mImag
End Sub

Public Sub Conj()
'Le conjugue
mImag = -mImag
End Sub

Public Sub Add(cpx As clsCpx)
'Lui ajoute un autre complexe
mReal = mReal + cpx.RealPart
mImag = mImag + cpx.ImagPart
End Sub

Public Sub AddVal(R As Double, i As Double)
'idem, mais directement avec les valeurs
mReal = mReal + R
mImag = mImag + i
End Sub

Public Sub Mult(cpx As clsCpx)
'Le multiplier par un autre complexe
Dim nReal As Double
Dim nImag As Double
nReal = mReal * cpx.RealPart - mImag * cpx.ImagPart
nImag = mReal * cpx.ImagPart + mImag * cpx.RealPart
mReal = nReal
mImag = nImag
End Sub

Public Sub MultVal(R As Double, i As Double)
'idem mais avec les valeurs
Dim nReal As Double
Dim nImag As Double
nReal = mReal * R - mImag * i
nImag = mReal * i + mImag * R
mReal = nReal
mImag = nImag
End Sub

Public Sub Square()
'L'�l�ve au carr�
Dim nReal As Double
Dim nImag As Double
nReal = mReal ^ 2 - mImag ^ 2
nImag = 2 * mReal * mImag
mReal = nReal
mImag = nImag
End Sub

Public Sub DefinePol(Rd As Double, Ang As Double)
'D�finit le complexe (lui attribut une nouvelle valeur) par coordonn�es polaires
mReal = Rd * Cos(Ang)
mImag = Rd * Sin(Ang)
End Sub

Public Sub DefineAlg(RP As Double, IP)
'Et par coordonn�es cart�siennes
mReal = RP
mImag = IP
End Sub

Public Sub Assert()
'Affiche simplement le complexe dans la fen�tre de D�buggage
Debug.Print mReal & "+i*" & mImag
End Sub
