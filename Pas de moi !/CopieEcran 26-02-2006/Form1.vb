'Les sources suivantes m'ont servis pour cette application 

''''''''''''''''''''''''''''''''''''''''''
'Merci à [HVB] pour sa source et sa fonction ShotScreenPart()
'http://www.vbfrance.fr/code.aspx?id=30267
''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''
'Merci à [pinje] pour sa source du rectangle lors de la selection
'http://www.vbfrance.fr/code.aspx?ID=30265
''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''
'Merci à [labout] pour sa source sur l'impression
'http://www.vbfrance.com/code.aspx?ID=18017
''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''
'Merci à [ soldier8514 ] pour sa source sur la récupération de la touche F7
'http://www.vbfrance.com/code.aspx?id=29269
''''''''''''''''''''''''''''''''''''''''''

'Le principe de cette appli est le suivant, lors du lancement de l'application 
'Le form est caché, lors du doucle-clic sur l'icone du menu systray ou de l'appui 
'sur la touche F7 , une copie de tout l'ecran est réalisé,le Form et le picturebox
'qu'il contient son redimensionné a la taille de l'ecran. Cette copie est collée 
'dans le picturebox et le form est activé et visible ce qui laisse paraitre que 
'l'ecran est figé. il est possible alors de selectionner une zone de l'ecran en
'faisant un cliqué-déplacé-relaché qui enregistre la position de la souris 
'puis copie la zone selectionnée dans le presse papier ou dans un fichier image.

'Modification le 25/10/2005
'Ajout du Cadre pendant la selection

'Modification le 30/10/2005
'Apercu avant impression de la selection

'Modification le 06/11/2005
'Changement de l'icone
'Ajout du Clic droit "Capturer"
'Appuie sur touche F7 pour commencer

'ANKOU22

'Merci de laisser vos commentaires si vous appréciez cette source ou si vous avez quelques idées 
'pour l'améliorer.



Imports System.Windows.Forms
Imports System
Imports System.Drawing
Imports System.IO
Imports System.IO.StreamWriter


Public Class Form1
    Inherits System.Windows.Forms.Form


#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem24 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Private WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.MenuItem17 = New System.Windows.Forms.MenuItem
        Me.MenuItem18 = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuItem21 = New System.Windows.Forms.MenuItem
        Me.MenuItem22 = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem24 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.MenuItem25 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "NotifyIcon1"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuItem1, Me.MenuItem4})
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem8, Me.MenuItem7, Me.MenuItem2, Me.MenuItem11, Me.MenuItem19})
        Me.MenuItem3.Text = "Selection vers"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 0
        Me.MenuItem8.Text = "Presse Papier"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Text = "Fichier Jpeg"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem9, Me.MenuItem10, Me.MenuItem12, Me.MenuItem13, Me.MenuItem15})
        Me.MenuItem2.Text = "un Fichier de type"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 0
        Me.MenuItem9.Text = "Bmp"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 1
        Me.MenuItem10.Text = "Gif"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 2
        Me.MenuItem12.Text = "Jpeg"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 3
        Me.MenuItem13.Text = "Png"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 4
        Me.MenuItem15.Text = "Tiff"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 3
        Me.MenuItem11.Text = "Aperçu avant Impression"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 4
        Me.MenuItem19.Text = "Imprimante"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem5, Me.MenuItem16, Me.MenuItem23, Me.MenuItem24})
        Me.MenuItem1.Text = "Ecran vers"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Text = "Presse Papier"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "Fichier Jpeg"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 2
        Me.MenuItem16.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem17, Me.MenuItem18, Me.MenuItem20, Me.MenuItem21, Me.MenuItem22})
        Me.MenuItem16.Text = "un Fichier de type"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 0
        Me.MenuItem17.Text = "Bmp"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 1
        Me.MenuItem18.Text = "Gif"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 2
        Me.MenuItem20.Text = "Jpeg"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 3
        Me.MenuItem21.Text = "Png"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 4
        Me.MenuItem22.Text = "Tiff"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 3
        Me.MenuItem23.Text = "Aperçu avant Impression"
        '
        'MenuItem24
        '
        Me.MenuItem24.Index = 4
        Me.MenuItem24.Text = "Imprimante"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "Annuler"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Cursor = System.Windows.Forms.Cursors.Cross
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(800, 600)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 80
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "jpg"
        '
        'ContextMenu2
        '
        Me.ContextMenu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem25, Me.MenuItem14})
        '
        'MenuItem25
        '
        Me.MenuItem25.Index = 0
        Me.MenuItem25.Text = "Capturer"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 1
        Me.MenuItem14.Text = "Quitter"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(800, 600)
        Me.ControlBox = False
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region


    '###################################################################
    'API utilisée pour la récupération de l'appui sur la touche F7
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Const vbKeyF7 = System.Windows.Forms.Keys.F7
    'Touche F7
    '###################################################################

    '###################################################################
    'API utilisée pour Faire la copie Ecran
    Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As IntPtr, ByVal x As Int32, _
          ByVal y As Int32, ByVal nWidth As Int32, ByVal nHeight As Int32, ByVal hSrcDC As IntPtr, _
          ByVal xSrc As Int32, ByVal ySrc As Int32, ByVal dwRop As Int32) As Int32

    Private Const SRCCOPY = &HCC0020

    'API Récupère le handle du bureau
    Private Declare Function GetDesktopWindow Lib "user32" () As IntPtr
    '###################################################################


    'Variable pour stocker les dimensions de l'ecran
    Dim ScreenX, ScreenY As Integer

    'Variable des images stockées
    Dim ImageFull As Bitmap
    Dim ImageCapture As Bitmap
    Dim ImageImprime As Bitmap

    '###################################################################
    'Constante et Variables pour la Selection et l'enregistrement des Images

    'dimenssions de la selection
    Dim DebutX, DebutY, FinX, FinY, Largeur, Hauteur As Integer

    'Variable indique si une capture est en cours
    Dim CaptureEnCours As Boolean = False

    'Constantes qui définissent l'action en cours
    Const ECRAN_PRESSEPAPIER = 1
    Const ECRAN_FICHIER = 2
    Const SELECTION_PRESSEPAPIER = 3
    Const SELECTION_FICHIER = 4
    Const SELECTION_APERCU = 5
    Const SELECTION_IMPRIMANTE = 6
    Const ECRAN_APERCU = 7
    Const ECRAN_IMPRIMANTE = 8

    Dim Type_Capture As Integer = 0

    'Constantes qui définissent le type de fichier en sortie
    Const BMP = 1
    Const GIF = 2
    Const ICO = 3
    Const JPEG = 4
    Const PNG = 5
    Const TIFF = 6

    Dim Type_Fichier As Integer = 0
    '###################################################################


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Recupère la taille de l'ecran
        ScreenX = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width
        ScreenY = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height

        'Ajout de l'icone dans le systray
        NotifyIcon1.ContextMenu = ContextMenu2
        NotifyIcon1.Visible = True
        NotifyIcon1.Text = "Double-Cliquez sur l'icone " & vbCr & "puis selectionnez la zone à copier"

    End Sub
    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        'Masque la Fenêtre si pas de capture
        'nécéssaire notament à la première ouverture de la Form
        If CaptureEnCours = False Then
            Me.Hide()
        End If
    End Sub
    Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        'Double click sur l'icone du menu systray
        Commence_Capture()

    End Sub

    Function Commence_Capture()

        CaptureEnCours = True

        'Position et Taille de la fenetre et du PictureBox
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Width = ScreenX
        Me.Height = ScreenY

        'Position et Taille de du PictureBox
        PictureBox1.Location = New System.Drawing.Point(0, 0)
        PictureBox1.Width = ScreenX
        PictureBox1.Height = ScreenY

        'Copie le Bureau en Totalité
        'Dim image As Bitmap
        ImageFull = CopieEcran(ScreenX, ScreenY)

        'le Form est visible et actif
        Me.Activate()
        Me.Visible = True

        'Copie l'image dans le PictureBox
        Me.PictureBox1.Image = ImageFull

    End Function

    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        'Bouton Gauche de la souris est enfoncé
        If e.Button = MouseButtons.Left Then
            'Recupère la position de la souris
            DebutX = e.X
            DebutY = e.Y
            FinX = e.X
            FinY = e.Y
            Dim r As Rectangle
            r = PictureBox1.RectangleToScreen(New Rectangle(DebutX, DebutY, FinX - DebutX, FinY - DebutY))
            ControlPaint.DrawReversibleFrame(r, Me.BackColor, FrameStyle.Dashed)

        End If
    End Sub
    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        'survole de la souris

        If e.Button = MouseButtons.Left Then
            'On efface l'ancien rectangle 
            Dim r As Rectangle
            r = PictureBox1.RectangleToScreen(New Rectangle(DebutX, DebutY, FinX - DebutX, FinY - DebutY))
            ControlPaint.DrawReversibleFrame(r, Me.BackColor, FrameStyle.Dashed)

            FinX = e.X
            FinY = e.Y

            'On redessine le rectangle avec les nouvelles coordonnes de la souris 
            r = PictureBox1.RectangleToScreen(New Rectangle(DebutX, DebutY, FinX - DebutX, FinY - DebutY))
            ControlPaint.DrawReversibleFrame(r, Me.BackColor, FrameStyle.Dashed)
        End If
    End Sub
    Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
        'Bouton de la souris relevé

        'si l'utilisateur a precedement double cliquer sur l'icone du menu systray
        If CaptureEnCours = True Then

            'On efface le dernier rectangle 
            Dim r As Rectangle
            r = PictureBox1.RectangleToScreen(New Rectangle(DebutX, DebutY, FinX - DebutX, FinY - DebutY))
            ControlPaint.DrawReversibleFrame(r, System.Drawing.Color.Blue, FrameStyle.Dashed)

            'Recupère la position de la souris
            FinX = e.X
            FinY = e.Y

            Me.PictureBox1.Refresh()


            'If DebutX > 0 And DebutY > 0 And FinX > 0 And FinY > 0 Then

            'Si l'utilisateur à fait la selection de droite à gauche
            If DebutX > FinX Then
                Dim TempFinX As Integer
                TempFinX = DebutX
                DebutX = FinX
                FinX = TempFinX
            End If

            'Si l'utilsateur à fais sa selection de bas en haut
            If DebutY > FinY Then
                Dim TempFinY As Integer
                TempFinY = DebutY
                DebutY = FinY
                FinY = TempFinY
            End If

            'Affiche le menu Contextuel à la position courante
            Dim pos As New System.Drawing.Point(e.X, e.Y)
            ContextMenu1.Show(Me, pos)

            'End If

            'La selection est terminée je Cache la Fenetre
            'Me.Visible = False
            CaptureEnCours = False
            Me.Visible = False
        End If

    End Sub

    Private Sub Capturer()
        On Error Resume Next

        SaveFileDialog1.FileName = ""

        'Deplace le menu Contextuel en dehors de l'ecran pour ne pas qu'il s'imprime
        ContextMenu1.SourceControl.Location = New System.Drawing.Point(3000, 3000)

        'Récupère les dimensions de la selection
        Hauteur = FinX - DebutX
        Largeur = FinY - DebutY

        'Selon le type de Fichier en sortie
        'initialise le Filtre du SaveFileDialog
        Select Case Type_Fichier
            Case BMP : SaveFileDialog1.Filter = "Fichier Bmp|*.Bmp"
            Case GIF : SaveFileDialog1.Filter = "Fichier Gif|*.Gif"
            Case JPEG : SaveFileDialog1.Filter = "Fichier Jpeg|*.jpg"
            Case PNG : SaveFileDialog1.Filter = "Fichier Png|*.Png"
            Case TIFF : SaveFileDialog1.Filter = "Fichier Tiff|*.Tiff"
            Case Else : SaveFileDialog1.Filter = "Fichier Jpeg|*.jpg"
        End Select

        'Selon le type de capture
        If Type_Capture = SELECTION_FICHIER Then
            'Selection Vers un Fichier

            'Capture de la Selection
            ImageCapture = CopieEcran(FinX - DebutX, FinY - DebutY, DebutX, DebutY)

            'Ouvre le selecteur de Fichier
            SaveFileDialog1.ShowDialog()

            'Ouvre le selecteur et enregistre le fichier
            Enregistre(ImageCapture)


        ElseIf Type_Capture = SELECTION_PRESSEPAPIER Then
            'Selection Vers le Presse Papier

            'Capture de la Selection
            ImageCapture = CopieEcran(FinX - DebutX, FinY - DebutY, DebutX, DebutY)

            'image vers Presse Papier
            Clipboard.SetDataObject(ImageCapture)


        ElseIf Type_Capture = ECRAN_FICHIER Then
            'Ecran Vers un Fichier

            'Capture de l'ecran
            ImageCapture = ImageFull

            'Ouvre le selecteur de Fichier
            SaveFileDialog1.FileName = ""
            SaveFileDialog1.ShowDialog()

            'Ouvre le selecteur et enregistre le fichier
            Enregistre(ImageCapture)

        ElseIf Type_Capture = ECRAN_PRESSEPAPIER Then
            'l'ecran vers le Presse Papier

            'Capture de l'ecran
            ImageCapture = ImageFull

            'image vers Presse Papier
            Clipboard.SetDataObject(ImageCapture)

        ElseIf Type_Capture = SELECTION_APERCU Then

            Dim m_pd As New Printing.PrintDocument

            'Capture de la zone
            ImageImprime = CopieEcran(FinX - DebutX, FinY - DebutY, DebutX, DebutY)

            AddHandler m_pd.PrintPage, AddressOf ImprimeImage

            Dim ppdlg As New ExtendedPrintPreviewDialog
            With ppdlg
                .PrintPreviewControl.Document = m_pd
                .WindowState = FormWindowState.Maximized
                .ShowDialog()
                .Dispose()
            End With

        ElseIf Type_Capture = SELECTION_IMPRIMANTE Then

            Dim m_pd As New Printing.PrintDocument

            'Capture de la zone
            ImageImprime = CopieEcran(FinX - DebutX, FinY - DebutY, DebutX, DebutY)

            AddHandler m_pd.PrintPage, AddressOf ImprimeImage

            m_pd.Print() 'impression directe sans aperçu

        ElseIf Type_Capture = ECRAN_APERCU Then

            Dim m_pd As New Printing.PrintDocument

            'Capture de la zone
            ImageImprime = ImageFull

            AddHandler m_pd.PrintPage, AddressOf ImprimeImage

            Dim ppdlg As New ExtendedPrintPreviewDialog
            With ppdlg
                .PrintPreviewControl.Document = m_pd
                .WindowState = FormWindowState.Maximized
                .ShowDialog()
                .Dispose()
            End With

        ElseIf Type_Capture = ECRAN_IMPRIMANTE Then

            Dim m_pd As New Printing.PrintDocument

            'Capture de la zone
            ImageImprime = CopieEcran(Largeur, Hauteur)

            AddHandler m_pd.PrintPage, AddressOf ImprimeImage

            m_pd.Print() 'impression directe sans aperçu

        End If


    End Sub

    Public Shared Function CopieEcran(ByVal TailleX As Integer, ByVal TailleY As Integer, Optional ByVal DebutX As Integer = 0, Optional ByVal DebutY As Integer = 0) As Bitmap

        On Error Resume Next
        'Création de l'Image bitmap cible
        Dim ImageCE As Bitmap = New Bitmap(TailleX, TailleY)
        'création de l'objet "graphics" à partir du handle du bureau 
        Dim SrcGraph As Graphics = Graphics.FromHwnd(GetDesktopWindow)
        'crée un objet graphics à partir du bitmap
        Dim BmpGraph As Graphics = Graphics.FromImage(ImageCE)
        'obtient le device context du bitmap 
        Dim bmpDC As IntPtr = BmpGraph.GetHdc()
        'obtient le device context du bureau 
        Dim hDC As IntPtr = SrcGraph.GetHdc()

        'copie chaque bits affichés dans le device context hDC dans le device context bmpDC du bitmap 
        BitBlt(bmpDC, 0, 0, TailleX, TailleY, hDC, DebutX, DebutY, &HCC0020)

        'Libère les ressources 
        SrcGraph.ReleaseHdc(hDC)
        BmpGraph.ReleaseHdc(bmpDC)
        SrcGraph.Dispose()
        BmpGraph.Dispose()
        'Renvoi l'Image
        Return ImageCE

    End Function

    Function Enregistre(ByVal Image As Bitmap)
        If SaveFileDialog1.FileName <> "" Then
            Dim Chemin As String = SaveFileDialog1.FileName

            Select Case Type_Fichier
                Case BMP : Image.Save(Chemin, System.Drawing.Imaging.ImageFormat.Bmp)
                Case GIF : Image.Save(Chemin, System.Drawing.Imaging.ImageFormat.Gif)
                Case JPEG : Image.Save(Chemin, System.Drawing.Imaging.ImageFormat.Jpeg)
                Case PNG : Image.Save(Chemin, System.Drawing.Imaging.ImageFormat.Png)
                Case TIFF : Image.Save(Chemin, System.Drawing.Imaging.ImageFormat.Tiff)
                Case Else : Image.Save(Chemin, System.Drawing.Imaging.ImageFormat.Jpeg)
            End Select
            SaveFileDialog1.Dispose()
        End If
    End Function

    'Options du Menu contextuel "Ecran"
    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Type_Capture = ECRAN_PRESSEPAPIER
        Capturer()
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Type_Capture = ECRAN_FICHIER
        Capturer()
    End Sub

    'Options du Menu contextuel "Selection"
    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Type_Capture = SELECTION_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        Type_Capture = SELECTION_PRESSEPAPIER
        Capturer()
    End Sub

    'Options du Menu contextuel "Annuler"
    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        'L'utilisteur a cliqué sur Annuler
        Me.Hide()
    End Sub

    'Options du Menu contextuel "Selection ... vers Fichier"
    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Type_Fichier = BMP
        Type_Capture = SELECTION_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Type_Fichier = GIF
        Type_Capture = SELECTION_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Type_Fichier = ICO
        Type_Capture = SELECTION_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        Type_Fichier = JPEG
        Type_Capture = SELECTION_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        Type_Fichier = PNG
        Type_Capture = SELECTION_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click
        Type_Fichier = TIFF
        Type_Capture = SELECTION_FICHIER
        Capturer()
    End Sub

    'Options du Menu contextuel "Ecran ... vers Fichier"
    Private Sub MenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem17.Click
        Type_Fichier = BMP
        Type_Capture = ECRAN_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem18.Click
        Type_Fichier = GIF
        Type_Capture = ECRAN_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Type_Fichier = ICO
        Type_Capture = ECRAN_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem20.Click
        Type_Fichier = JPEG
        Type_Capture = ECRAN_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem21.Click
        Type_Fichier = PNG
        Type_Capture = ECRAN_FICHIER
        Capturer()
    End Sub

    Private Sub MenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem22.Click
        Type_Fichier = TIFF
        Type_Capture = ECRAN_FICHIER
        Capturer()
    End Sub

    'Options du Menu systray "vers Apercu avant Impression"
    Private Sub MenuItem11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        'Selection vers apercu avant impression
        Type_Capture = SELECTION_APERCU
        Capturer()
    End Sub

    Private Sub MenuItem19_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem19.Click
        'Selection vers imprimante
        Type_Capture = SELECTION_IMPRIMANTE
        Capturer()
    End Sub

    Private Sub MenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem23.Click
        'Ecran vers apercu avant impression
        Type_Capture = ECRAN_APERCU
        Capturer()
    End Sub

    Private Sub MenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem24.Click
        'Ecran vers imprimante
        Type_Capture = ECRAN_IMPRIMANTE
        Capturer()
    End Sub

    'Options du Menu systray "Quitter"
    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click
        'Quitter
        Me.Close()
    End Sub

    Private Sub MenuItem25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem25.Click
        'Clic droit Capturer
        Commence_Capture()
    End Sub

    Private Sub ImprimeImage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        'Construction du document pour l'impression

        If Type_Capture = ECRAN_APERCU Or Type_Capture = ECRAN_IMPRIMANTE Then
            Largeur = ScreenY
            Hauteur = ScreenX
        End If

        'Calcul pour le centrage de l'image
        Dim HautPrint As Integer = ((720 - 50) * Largeur) / Hauteur
        Dim DebPrintY As Integer = 50 + (((1070 - 50) - HautPrint) / 2)
        Dim LargPrint As Integer = ((1070 - 50) * Hauteur) / Largeur
        Dim DebPrintX As Integer = 50 + (((720 - 50) - LargPrint) / 2)

        '
        If (DebPrintY > 50) Then
            'Ajustement sur la largeur
            e.Graphics.DrawImage(ImageImprime, 50, DebPrintY, 720, HautPrint)
        Else
            'Ajustement sur la Hauteur
            e.Graphics.DrawImage(ImageImprime, DebPrintX, 50, LargPrint, 1070)
        End If

        'plus de page
        e.HasMorePages = False

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Si appuie sur la touche F7
        If PresseToucheF7() Then
            Commence_Capture()
        End If
    End Sub

    Function PresseToucheF7() As Boolean
        'Fonction interroge pour savoir si touche F7 enfoncé

        Dim touche As Long

        touche = GetAsyncKeyState(vbKeyF7)
        If (touche And &H1) = &H1 Then
            Return True
        Else
            Return False
        End If

    End Function

End Class
