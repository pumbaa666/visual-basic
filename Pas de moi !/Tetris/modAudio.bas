Attribute VB_Name = "modAudio"
Option Explicit
Option Base 1

Public bMidi As Boolean
Public bWave As Boolean

Private AudioWave(4) As String
Private AudioMidi(2) As String

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
     ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' -----------------------------------------------------------------------------
' Nom  : InitialiseAudio
' -----------------------------------------------------------------------------
Public Sub InitialiseAudio()
Dim sPath As String

    sPath = App.Path
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    sPath = sPath & "Audio\"
    
    ' fichiers wave
    AudioWave(1) = sPath & "Toc.wav"
    AudioWave(2) = sPath & "Niveau.wav"
    AudioWave(3) = sPath & "Lignes.wav"
    AudioWave(4) = sPath & "Mouvement.wav"
    bWave = True
    
    ' fichiers midi
    AudioMidi(1) = sPath & "Tetris.mid"
    AudioMidi(2) = sPath & "Fin.mid"
    bMidi = True
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : PlayWave
' -----------------------------------------------------------------------------
Public Sub PlayWave(ByVal Wave As EWave)
    If bWave Then sndPlaySound AudioWave(Wave), 1
End Sub

' -----------------------------------------------------------------------------
' Nom  : PlayMidi
' -----------------------------------------------------------------------------
Public Sub PlayMidi(ByVal Midi As EMidi)

    If bMidi Then
        ' charge le fichier
        mciSendString "close mid", 0&, 0, 0
        mciSendString "open """ & AudioMidi(Midi) & _
            """ type sequencer alias mid", 0, 0, 0
        ' joue le fichier
        mciSendString "seek mid to start", 0, 0, 0
        mciSendString "play mid from 0", 0, 0, 0
   End If
   
End Sub

' -----------------------------------------------------------------------------
' Nom  : RepeatMidi
' -----------------------------------------------------------------------------
Public Sub RepeatMidi()
Dim sRetString As String * 128
Dim sPosition As String
Dim sTotal As String
   
    If bMidi Then
        mciSendString "status mid length", sRetString, 128, 0
        sTotal = Left(sRetString, InStr(sRetString, Chr(0)) - 1)
    
        mciSendString "status mid position", sRetString, 128, 0
        sPosition = Left(sRetString, InStr(sRetString, Chr(0)) - 1)
    
        If sTotal <> "" And sTotal = sPosition Then PlayMidi MID_TETRIS
        
    End If

End Sub

' -----------------------------------------------------------------------------
' Nom  : StopMidi
' -----------------------------------------------------------------------------
Public Sub StopMidi(Optional ByVal bClose As Boolean = False)
    mciSendString "stop mid", 0, 0, 0
    If bClose Then mciSendString "close mid", 0&, 0, 0
End Sub

' -----------------------------------------------------------------------------
' Nom  : TempoMidi
' Rem  : Changement dans le fichier midi :
'        "FF 51 03 tt tt tt" avec "tt tt tt" le tempo in microseconds
' -----------------------------------------------------------------------------
Public Sub TempoMidi(ByVal A As Byte, ByVal B As Byte, ByVal C As Byte) 'ByVal Tempo As Single)
Dim hFile As Long
'Dim A As Byte, B As Byte, C As Byte

    ' ferme le fichier en cours
    mciSendString "close mid", 0&, 0, 0
    
    ' change le tempo
    hFile = FreeFile
    'A = Int(Tempo): B = Int((Tempo - Int(Tempo)) * 256): C = 0
    Open AudioMidi(MID_TETRIS) For Binary As #hFile
    Seek #hFile, 41
    Put #hFile, , A: Put #hFile, , B: Put #hFile, , C
    Close #hFile
    
End Sub

