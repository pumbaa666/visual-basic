VERSION 5.00
Begin VB.Form Maine 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brick Game"
   ClientHeight    =   5505
   ClientLeft      =   4140
   ClientTop       =   3255
   ClientWidth     =   6255
   Icon            =   "Tetris_plato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Translater 
      Enabled         =   0   'False
      Left            =   6360
      Top             =   4800
   End
   Begin VB.Timer Compteur 
      Enabled         =   0   'False
      Left            =   6360
      Top             =   240
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   1
      Left            =   480
      Picture         =   "Tetris_plato.frx":0312
      Top             =   5040
      Width           =   2790
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   0
      Left            =   480
      Picture         =   "Tetris_plato.frx":3C34
      Top             =   120
      Width           =   2790
   End
   Begin VB.Image Image1 
      Height          =   5310
      Index           =   1
      Left            =   3240
      Picture         =   "Tetris_plato.frx":7556
      Top             =   120
      Width           =   390
   End
   Begin VB.Label ligne99 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   283
      Tag             =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.Label ligne101 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   282
      Tag             =   "0"
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label ligne102 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   281
      Tag             =   "0"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label ligne103 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   280
      Tag             =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label ligne100 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   279
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label ligne100 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   278
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label ligne100 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   277
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label ligne100 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   276
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label ligne100 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   275
      Tag             =   "0"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label ligne103 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   274
      Tag             =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label ligne103 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   273
      Tag             =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label ligne103 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   272
      Tag             =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label ligne103 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   271
      Tag             =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label ligne102 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   270
      Tag             =   "0"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label ligne102 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   269
      Tag             =   "0"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label ligne102 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   268
      Tag             =   "0"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label ligne102 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   267
      Tag             =   "0"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label ligne101 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   266
      Tag             =   "0"
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label ligne101 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   265
      Tag             =   "0"
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label ligne101 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   264
      Tag             =   "0"
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label ligne101 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   263
      Tag             =   "0"
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label ligne99 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   262
      Tag             =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.Label ligne99 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   261
      Tag             =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.Label ligne99 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   260
      Tag             =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.Label ligne99 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   259
      Tag             =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   5310
      Index           =   0
      Left            =   120
      Picture         =   "Tetris_plato.frx":E438
      Top             =   120
      Width           =   390
   End
   Begin VB.Label no_piecer 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   4560
      TabIndex        =   258
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label niv_lvl 
      BackColor       =   &H00FF8080&
      Caption         =   "0"
      Height          =   255
      Left            =   4920
      TabIndex        =   257
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Niveau:"
      Height          =   255
      Left            =   4200
      TabIndex        =   256
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label val_lgne 
      BackColor       =   &H00FF8080&
      Caption         =   "0"
      Height          =   255
      Left            =   4920
      TabIndex        =   255
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label NB_lignes 
      BackColor       =   &H00FF8080&
      Caption         =   "Lignes:"
      Height          =   255
      Left            =   4200
      TabIndex        =   254
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Score_val 
      BackColor       =   &H00FF8080&
      Caption         =   "0"
      Height          =   255
      Left            =   4920
      TabIndex        =   253
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Ligne0 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   252
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   251
      Tag             =   "-1"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   250
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   249
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   248
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   247
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   246
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   245
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   244
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   243
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   242
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne0 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   241
      Tag             =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   240
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   239
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   238
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   237
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   236
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   235
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   234
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   233
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   232
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   231
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   230
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne20 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   229
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne1 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   228
      Tag             =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne2 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   227
      Tag             =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne3 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   226
      Tag             =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne4 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   225
      Tag             =   "0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne5 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   224
      Tag             =   "0"
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne6 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   223
      Tag             =   "0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne7 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   222
      Tag             =   "0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne8 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   221
      Tag             =   "0"
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne10 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   220
      Tag             =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne11 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   219
      Tag             =   "0"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne12 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   218
      Tag             =   "0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne13 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   217
      Tag             =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne14 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   216
      Tag             =   "0"
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne15 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   215
      Tag             =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne16 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   214
      Tag             =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne17 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   213
      Tag             =   "0"
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne18 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   212
      Tag             =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne9 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   211
      Tag             =   "0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   210
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne1 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   209
      Tag             =   "-1"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne2 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   208
      Tag             =   "-1"
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne3 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   207
      Tag             =   "-1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne4 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   206
      Tag             =   "-1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne5 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   205
      Tag             =   "-1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne6 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   204
      Tag             =   "-1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne7 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   203
      Tag             =   "-1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne8 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   202
      Tag             =   "-1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne10 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   201
      Tag             =   "-1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne11 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   200
      Tag             =   "-1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne12 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   199
      Tag             =   "-1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne13 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   198
      Tag             =   "-1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne14 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   197
      Tag             =   "-1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne15 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   196
      Tag             =   "-1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne16 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   195
      Tag             =   "-1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne17 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   194
      Tag             =   "-1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne18 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   193
      Tag             =   "-1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne9 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   192
      Tag             =   "-1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   191
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   190
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   189
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   188
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   187
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   186
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   185
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   184
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   183
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   182
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne19 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   181
      Tag             =   "-1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   180
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   179
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   178
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   177
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   176
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   175
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   174
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   173
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   172
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   171
      Tag             =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   170
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   169
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   168
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   167
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   166
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   165
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   164
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   163
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   162
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   161
      Tag             =   "0"
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   160
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   159
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   158
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   157
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   156
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   155
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   154
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   153
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   152
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   151
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   150
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   149
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   148
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   147
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   146
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   145
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   144
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   143
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   142
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   141
      Tag             =   "0"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   140
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   139
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   138
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   137
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   136
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   135
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   134
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   133
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   132
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   131
      Tag             =   "0"
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   130
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   129
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   128
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   127
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   126
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   125
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   124
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   123
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   122
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   121
      Tag             =   "0"
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   120
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   119
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   118
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   117
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   116
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   115
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   114
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   113
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   112
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   111
      Tag             =   "0"
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   110
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   109
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   108
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   107
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   106
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   105
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   104
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   103
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   102
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   101
      Tag             =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   100
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   99
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   98
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   97
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   96
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   95
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   94
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   93
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   92
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   91
      Tag             =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   90
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   89
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   88
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   87
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   86
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   85
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   84
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   83
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   82
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   81
      Tag             =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   80
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   79
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   78
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   77
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   76
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   75
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   74
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   73
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   72
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   71
      Tag             =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   70
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   69
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   68
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   67
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   66
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   65
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   64
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   63
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   62
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   61
      Tag             =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   60
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   59
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   58
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   57
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   56
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   55
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   54
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   53
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   52
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   51
      Tag             =   "0"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   50
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   49
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   48
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   47
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   46
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   45
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   44
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   43
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   42
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   41
      Tag             =   "0"
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   40
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   39
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   38
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   37
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   36
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   35
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   34
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   33
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   32
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   31
      Tag             =   "0"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   30
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   29
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   28
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   27
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   26
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   25
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   24
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   23
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   22
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   21
      Tag             =   "0"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   20
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   19
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   18
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   17
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   16
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   14
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   13
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   12
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   11
      Tag             =   "0"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   8
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   7
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   6
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   5
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   4
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   3
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   2
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Ligne1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   1
      Tag             =   "0"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Score_ 
      BackColor       =   &H00FF8080&
      Caption         =   "Score:"
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   2460
      Left            =   3720
      Picture         =   "Tetris_plato.frx":1531A
      Top             =   120
      Width           =   2460
   End
   Begin VB.Image Image4 
      Height          =   2460
      Left            =   3720
      Picture         =   "Tetris_plato.frx":28E8C
      Top             =   3000
      Width           =   2460
   End
   Begin VB.Menu Fichier 
      Caption         =   "Fichier"
      Index           =   1
      Begin VB.Menu new 
         Caption         =   "Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu accès_Score 
         Caption         =   "Meilleurs Scores"
         Shortcut        =   ^S
      End
      Begin VB.Menu exit 
         Caption         =   "Quitter"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Index           =   2
      Begin VB.Menu def_tch 
         Caption         =   "Définir les touches"
         Shortcut        =   ^P
      End
      Begin VB.Menu Spec_pie 
         Caption         =   "Pièces spéciales"
      End
   End
   Begin VB.Menu Aide 
      Caption         =   "Aide?"
      Index           =   3
   End
End
Attribute VB_Name = "Maine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fin As Byte
Dim position_y As Byte
Dim position_x As Byte
Dim exiting As Byte
Dim col As Byte             '
Dim No_piece As Single      'Contien le N° de tag de piece en cours
Dim defi As Byte
Dim haut As Byte            'indice de hauteur d'une piece
Dim large As Byte           'indice de largeur d'une piece


Dim rot As Integer          'Définit la position de rotation
Dim differents As Byte      'définit le nombre de positions par piece
Dim type_piece As Integer   'contien la piece en cours
Dim pce_suivante As Byte    'contien la piece suivante
Dim pause As Boolean        ' dis si le jeu est déja en pause
Dim temps As Integer        'temps de chute d'une piece
Dim touche As String        'variable contenan le caractere de la touche pressés
'**********************************************************************
'   Description
'
'   Cette procédure affiche le tableau des scores
'
'**********************************************************************
Private Sub accès_Score_Click()
    Compteur.Enabled = False        ' arrete le timer qui fai descendre les pieces
    meilleurs_scores.Show           ' affiche le panneau des scores
    retour = True
    
End Sub
'**********************************************************************
'   Description
'
'   Cette procédure affiche l'aide lorsqu'on seletionne
'   aide dans le menu
'**********************************************************************
Private Sub Aide_Click(Index As Integer)
    Compteur.Enabled = False        ' arrete le timer qui fai descendre les pieces
    frmAbout.Show                   ' montre le formulaire " a propos de"
End Sub
'**********************************************************************
'   Description
'
'   Cette procédure affiche le paneau de configuration des touches
'
'**********************************************************************
Private Sub def_tch_Click()
    Compteur.Enabled = False        ' arrete le timer qui fai descendre les pieces
    def_touches.Show                ' affiche la page de config des touches
End Sub
'**********************************************************************
'   Description
'
'   Cette procédure permet d'enlencher le timer de translation
'   ou d'augmenter la vitesse du timer principal suivant la touche
'   pressée
'
'**********************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If LCase(Chr(KeyCode)) = def_touches.touche(2) Or KeyCode = vbKeyDown Then
    
        touche = def_touches.touche(2)
        Compteur.Interval = 50
        Score_val = Score_val + 10
        
    End If
    
    If LCase(Chr(KeyCode)) = def_touches.touche(0) Or KeyCode = vbKeyLeft Then
    
        touche = def_touches.touche(0)
        Translater.Enabled = True
        Translater.Interval = 50
        
    End If
    
    If LCase(Chr(KeyCode)) = def_touches.touche(1) Or KeyCode = vbKeyRight Then
    
        touche = def_touches.touche(1)
        Translater.Enabled = True
        Translater.Interval = 50
        
    End If
    
    '   Rotation de 90°
    
    If LCase(Chr(KeyCode)) = def_touches.touche(3).Text Or KeyCode = vbKeyUp Then
     
        If rot + 1 < differents Then
            Call erase_current
            rot = rot + 1
            Call piece(position_y, position_x, type_piece & rot, No_piece)
        Else
            rot = 0
            Call erase_current
            Call piece(position_y, position_x, type_piece & rot, No_piece)

        End If
        

    End If
    
        '   mise en pause
    If LCase(Chr(KeyCode)) = def_touches.touche(4) Or KeyCode = vbKeyPause Then
        If pause = False Then
            info_partie.MMControl1.Command = "pause"
            Compteur.Enabled = pause
            dessin_pause (True)
            pause = True
            info_partie.MMControl2.FileName = App.Path & "\pause.wav"
            info_partie.MMControl2.Command = "close"
            info_partie.MMControl2.Command = "open"
            info_partie.MMControl2.Command = "play"
            
            
            Exit Sub
        Else
            info_partie.MMControl1.Command = "play"
            Compteur.Enabled = pause
            dessin_pause (False)
            pause = False
            
        End If
    End If

End Sub
'**********************************************************************
'   Description
'
'   Cette procédure permet d' éteindre le timer de translation
'   ou de réinitialiser la vitesse du timer principal suivant la touche
'   relachée
'
'**********************************************************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If LCase(Chr(KeyCode)) = def_touches.touche(2) Or KeyCode = vbKeyDown Then
        Compteur.Interval = temps
    End If

    If LCase(Chr(KeyCode)) = def_touches.touche(0) Or KeyCode = vbKeyLeft Then
        Translater.Enabled = False
    End If
    
    If LCase(Chr(KeyCode)) = def_touches.touche(1) Or KeyCode = vbKeyRight Then
        Translater.Enabled = False
    End If
    
End Sub
'**********************************************************************
'   Description
'
'   Cette procédure permet de faire tourner une piece ou de mettre
'   le jeux en pause ou l'enlever losqu'on presseune touche definie
'
'**********************************************************************
Private Sub Form_KeyPress(KeyAscii As Integer)


    


End Sub
'**********************************************************************
'   Description
'
'   Cette fonction affiche "Pause" à l'écran lorsque le jeux est en
'   pause
'
'**********************************************************************
Function dessin_pause(dessin As Boolean)
Select Case dessin
    
    Case True
    
        Ligne9(4).Caption = "P"
        Ligne9(5).Caption = "A"
        Ligne9(6).Caption = "U"
        Ligne9(7).Caption = "S"
        Ligne9(8).Caption = "E"
        attend (0.3)
        Ligne9(4).Caption = ""
        Ligne9(5).Caption = ""
        Ligne9(6).Caption = ""
        Ligne9(7).Caption = ""
        Ligne9(8).Caption = ""
        attend (0.3)
        Ligne9(4).Caption = "P"
        Ligne9(5).Caption = "A"
        Ligne9(6).Caption = "U"
        Ligne9(7).Caption = "S"
        Ligne9(8).Caption = "E"

    Case False
    
        Ligne9(4).Caption = ""
        Ligne9(5).Caption = ""
        Ligne9(6).Caption = ""
        Ligne9(7).Caption = ""
        Ligne9(8).Caption = ""
    
End Select

End Function
'**********************************************************************
'   Fonction                Erase current
'
'   Description
'
'   Cette fonction efface la piece qui est en train de tomber
'   pour qu'il n'y aie pas de résidu lors de translation ou rotation
'
'**********************************************************************
Function erase_current()
Dim cpt As Byte
Dim coll As Byte

    For cpt = 1 To 18
        For col = 1 To 10
            If Controls("ligne" & cpt)(col).Tag = No_piece Then

                Controls("ligne" & cpt)(col).Tag = 0
                Controls("ligne" & cpt)(col).BackColor = &HFFFFFF
                Controls("ligne" & cpt)(col).BorderStyle = 0

            End If
        Next
    Next
End Function
Private Sub exit_Click()
    
    Unload Maine
    End

End Sub

Function piece_en_cours()
    Dim test As Byte

    Randomize
    
    Do Until test <> 0
    
        test = CByte(Rnd * 7)
    
    Loop

    type_piece = test
      
             
End Function
'**********************************************************************
'   Fonction                Attend
'
'   Description
'
'   Cette fonction permet d'attendre lors d'un effacement
'   de ligne par exemple
'
'   Entrée: on donne une valeur en seconde [s]
'
'**********************************************************************
Public Function attend(pausetime As Single)
    
    Dim start As Single
    
    start = Timer
        Do While Timer < start + pausetime
            DoEvents
        Loop
    
End Function
'**********************************************************************
'   Fonction                Form Activate
'
'   Description
'
'   Initialisation de la partie et des compteurs
'
'   Effacement du tableau
'
'   Difinition des temps de chute
'
'   Enclanchement du timer qui fait descendre les pieces
'
'**********************************************************************
Private Sub Form_Activate()
    Dim cpt As Byte
    Dim pce_suivante As Byte

    For cpt = 1 To 18
        For col = 1 To 10
            Controls("ligne" & cpt)(col).Tag = 0
            Controls("ligne" & cpt)(col).BackColor = &HFFFFFF
            Controls("ligne" & cpt)(col).BorderStyle = 0
        Next
    Next
    
    ' initialisation de la numérotation des piece de 1 jusqu'a l infini
    No_piece = 1
    Call piece_en_cours
    

    Call piece_suivante
    Call prevision(pce_suivante)


    position_x = 5
    position_y = 1
    defi = 5
    
    rot = 0
    
    If info_partie.op_niv_1.Value = True Then
        temps = 750
        niv_lvl = 1
    End If
    
    If info_partie.op_niv_2.Value = True Then
        temps = 700
        niv_lvl = 2
    End If
    
    If info_partie.op_niv_3.Value = True Then
        temps = 650
        niv_lvl = 3
    End If
    
    If info_partie.op_niv_4.Value = True Then
        temps = 600
        niv_lvl = 4
    End If
    
    If info_partie.op_niv_5.Value = True Then
        temps = 550
        niv_lvl = 5
    End If
    
    If info_partie.op_niv_6.Value = True Then
        temps = 500
        niv_lvl = 6
    End If
    
    If info_partie.op_niv_7.Value = True Then
        temps = 450
        niv_lvl = 7
    End If
    
    If info_partie.op_niv_8.Value = True Then
        temps = 400
        niv_lvl = 8
    End If
    
    If info_partie.op_niv_9.Value = True Then
        temps = 350
        niv_lvl = 9
    End If
       
                    
    fin = 0
    
    Compteur.Interval = temps
    
    Compteur.Enabled = True

End Sub
'**********************************************************************
'   Fonction                prévision
'
'   Description
'
'   Cette fonction dessine la pièce dans la fenêtre de prédiction
'
'**********************************************************************
Function prevision(typ_pce As Byte)
    Call piece(101, 2, typ_pce & 0, No_piece)
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
'**********************************************************************
'   Fonction                pièce suivante
'
'   Description
'
'   Cette fonction génére les numéreau de pieces
'
'**********************************************************************
Function piece_suivante()
    
    Randomize
    
    Do Until test <> 0
    
    test = CByte(Rnd * 7)
    Loop
       
    pce_suivante = test
            
    
End Function
Sub Compteur_Timer()
    
    If position_y + haut < 20 Then
        
        'position_x = defi        ' donne la position actuelle que la piece doit reprendre
                            
        
        Call piece(position_y, position_x, type_piece & rot, No_piece)
        Call reperage   'controle si la ligne d'en dessous est libre
                                            
        If fin = 0 Then
            position_y = position_y + 1 ' decale la piece d une ligne
            
        Else
            fin = 0     ' remise a zéro de la marque pour savoir si il continue a descendre ou pas
            ' No_piece = No_piece + 1
            GoTo suite
        End If
    Else
suite:
        Call check  'controle si une ligne est pleine
        Call fin_partie 'controle si le jeux est plein
            ' réinitialisation de la position en X intiale
        If exiting = 1 Then 'si il detect que le jeux est pline
            Compteur.Enabled = False
            Maine.Hide
            meilleurs_scores.Show
            exiting = 0
            
            Exit Sub    'il sort
        End If
        No_piece = No_piece + 1      'augmentation de la numerotation des pieces

        
        type_piece = pce_suivante    'donne la valeur pour la piece suivante
        Call efface_prevision        'efface la fenetre de vue a l'avance
        Call piece_suivante          'génére un numéro de piece
        no_piecer = pce_suivante     'affiche sur la forme le numéro de typ de piece suivante
        Call prevision(pce_suivante) 'affiche la piece suivant
        position_y = 2               'réinitialiation de la position verticale
        position_x = 5               'réinitialisation de la position horisontale
        rot = 0
        Exit Sub
        
    End If
End Sub
'**********************************************************************
'   Description
'
'   Cette fonction permet d'effacer l'image de la piee suivante
'
'**********************************************************************
Function efface_prevision()
Dim ligne As Byte
Dim coll As Byte
For ligne = 99 To 103
    For coll = 1 To 5
        Controls("ligne" & ligne)(coll).Tag = 0
        Controls("ligne" & ligne)(coll).BackColor = &HFFFFFF
        Controls("ligne" & ligne)(coll).BorderStyle = 0
    Next coll
Next ligne
End Function
'**********************************************************************
'   Description
'
'   Cette fonction permet de rechercher si les cases d'en dessous
'   des pièces sont libres ou non
'
'   retour:
'           fin = 1 si occupé
'**********************************************************************
Function reperage()
Dim ligner As Byte
'Pour chaques collone de la pièce
For curseur = 0 To large - 1

    ligner = 1
    'cherche le haut de la piece
    Do While Controls("Ligne" & ligner)(position_x + curseur).Tag = 0

        ligner = ligner + 1
    Loop
    ' cherche la longueur de la pièce
    Do While Controls("Ligne" & ligner)(position_x + curseur).Tag = No_piece
        ligner = ligner + 1
    Loop
    'recherche le contenu de la celule en dessous de la pièce
    If Controls("Ligne" & ligner)(position_x + curseur).Tag > 0 Then
    fin = 1
    End If
    If Controls("Ligne" & haut)(position_x + curseur).Tag <> No_piece And Controls("Ligne" & haut)(position_x + curseur).Tag <> 0 Then
    Call fin_partie
    End If
Next


End Function
'**********************************************************************
'   Description
'
'   Cette fonction permet d'effacer une ligne de brique et de
'   déplacer les briques d'en dessus de la ligne un étage plus bas
'
'**********************************************************************
Function check()
Dim col_2 As Byte
Dim cpt As Byte
Dim cpt_2 As Integer
    For cpt = 1 To 18
    
        For col = 1 To 10
        
            If Controls("ligne" & cpt)(col).Tag <> 0 Then
                cpt_pts = cpt_pts + 1
            End If
            
            If cpt_pts = 10 Then
                
                'if val_lgne.
                Score_val = Score_val + 40
                val_lgne = val_lgne + 1
                
                For cpt_2 = cpt To 1 Step -1
                        
                    For col_2 = 1 To 10
                        Controls("ligne" & cpt_2)(col_2).BorderStyle = Controls("ligne" & cpt_2 - 1)(col_2).BorderStyle
                        Controls("ligne" & cpt_2)(col_2).Tag = Controls("ligne" & cpt_2 - 1)(col_2).Tag
                        Controls("ligne" & cpt_2)(col_2).BackColor = Controls("ligne" & cpt_2 - 1)(col_2).BackColor
                        
                    Next
                     
                Next
                    info_partie.MMControl2.FileName = App.Path & "\ligne.wav"
                    info_partie.MMControl2.Command = "close"
                    info_partie.MMControl2.Command = "open"
                    info_partie.MMControl2.Command = "play"
            End If
        
        Next
     
        cpt_pts = 0
    Next
End Function
'**********************************************************************
'   Description
'
'   Cette fonction colorie la grille avec des case grise et marque
'   "vous avez perdu" sur le tableau
'
'**********************************************************************
Function fin_partie()

    Dim cpt_lg As Byte
    Dim col As Byte
    Dim couleur As ColorConstants
    couleur = &HFF8080
    
    For cpt_lg = 1 To 10
        If Ligne2(cpt_lg).Tag <> 0 Then
        
            Dim cpt As Integer
                info_partie.MMControl1.Command = "close"
                For cpt = 18 To 1 Step -1
                    For col = 1 To 10
    
                            Controls("ligne" & cpt)(col).Tag = 0
                            Controls("ligne" & cpt)(col).BackColor = couleur '&HC0C0C0
                            Controls("ligne" & cpt)(col).BorderStyle = 0
    
                    Next
                    attend (0.08)
                Next
                
                info_partie.MMControl2.FileName = App.Path & "\fin.mp3"
                info_partie.MMControl2.Command = "close"
                info_partie.MMControl2.Command = "open"
                info_partie.MMControl2.Command = "play"
   
                Compteur.Enabled = False
                Ligne5(3).Caption = "V"
                Ligne5(5).Caption = "O"
                Ligne5(7).Caption = "U"
                Ligne5(9).Caption = "S"
                Ligne8(3).Caption = "A"
                Ligne8(5).Caption = "V"
                Ligne8(7).Caption = "E"
                Ligne8(9).Caption = "Z"
                Ligne11(2).Caption = "P"
                Ligne11(4).Caption = "E"
                Ligne11(6).Caption = "R"
                Ligne11(8).Caption = "D"
                Ligne11(10).Caption = "U"
                attend (6)
                exiting = 1
                 
     End If
    Next

End Function
'**********************************************************************
'   Description
'
'   Cette fonction contient toute les pieces
'
'   position_y      transmetre une coordonnée en y
'
'   position_x      transettre une coordonnée en x
'
'   Num_piece       donne le type de la piece
'
'   ind_piece       donne la numérotation des pieces
'
'**********************************************************************
Function piece(position_y As Byte, position_x As Byte, Num_piece As Byte, ind_piece As Single)
'effacer les 2lignes d'en dessus de la pièces

Dim orange As ColorConstants
Dim jaune As ColorConstants
Dim bleu As ColorConstants
Dim violet As ColorConstants
Dim vert As ColorConstants
Dim rouge As ColorConstants
Dim rose As ColorConstants
Dim bordure As Byte

orange = &H80FF&
jaune = &HC0C0&
bleu = &HC00000
violet = &H4080&
vert = &HC000&
rouge = &HC0&
rose = &HFF00FF
bordure = 1
Select Case Num_piece

    '****************'  x x x
    '     elle       '  x
    '****************'
    Case 10
    
    'effacer les 2lignes d'en dessus de la pièces
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = bleu
    Maine.Controls("Ligne" & position_y)(position_x).Tag = No_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = bleu
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = No_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = bleu
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = No_piece
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = No_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    haut = 2
    large = 3
    differents = 4
    
    
    '****************'  x
    '     elle       '  x
    '****************'  x x
Case 11
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = 0
   
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = bleu
    Maine.Controls("Ligne" & position_y)(position_x).Tag = No_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = No_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 2)(position_x).Tag = No_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).Tag = No_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BorderStyle = bordure
 
    haut = 3
    large = 2
    differents = 4
    
    '****************'          x
    '     elle       '      x x x
    '****************'
Case 12
    
        'effacer les 2lignes d'en dessus de la pièces
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = No_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = No_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).Tag = No_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BorderStyle = bordure
    
    
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = bleu
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = No_piece
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = bordure
    
    haut = 2
    large = 3
    differents = 4
    
    '****************'      x x
    '     elle       '        x
    '****************'        x
    
    Case 13

    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = bleu
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = bleu
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BackColor = bleu
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BorderStyle = bordure
    
    haut = 3
    large = 2
    differents = 4
    '****************'      x x
    '     carré      '      x x
    '****************'
    Case 20
       
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = jaune
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = jaune
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = jaune
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = jaune
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    haut = 2
    large = 2
    differents = 1
    
    '******************'      x x
    '       esse       '    x x
    '******************'
    Case 30
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BorderStyle = 0

    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = vert
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = vert
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = vert
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = vert
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    haut = 2
    large = 3
    differents = 2
    
    '******************'    x
    '    esse  drt     '    x x
    '******************'      x
    Case 31
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = vert
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = vert
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = vert
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BackColor = vert
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BorderStyle = bordure
    
    haut = 3
    large = 2
    differents = 2
    
    '******************'    x x
    '      deux        '      x x
    '******************'
    Case 40
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = 0
    
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = rouge
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = rouge
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = rouge
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BackColor = rouge
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BorderStyle = bordure
        
    haut = 2
    large = 3
    differents = 2
    
    '******************'      x
    '      deux  drt   '    x x
    '******************'    x
    Case 41
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    

    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = rouge
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = rouge
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = rouge
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x).BackColor = rouge
    Maine.Controls("Ligne" & position_y + 2)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x).BorderStyle = bordure
        
    haut = 3
    large = 2
    differents = 2
    
    
    
    '******************'
    '    demi-crx      '    x x x
    '******************'      x
    Case 50
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BorderStyle = 0
  
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = rose
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = rose
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = rose
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = rose
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    haut = 2
    large = 3
    differents = 4
    
    '******************'    x
    '    demi-crx      '    x x
    '******************'    x
    Case 51
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = rose
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = rose
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = rose
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x).BackColor = rose
    Maine.Controls("Ligne" & position_y + 2)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x).BorderStyle = bordure
    haut = 3
    large = 2
    differents = 4
    
    '******************'      x
    '    demi-crx      '    x x x
    '******************'
    Case 52
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = rose
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = rose
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = rose
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BackColor = rose
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BorderStyle = bordure
    
    haut = 2
    large = 3
    differents = 4
    
    '******************'      x
    '    demi-crx      '    x x
    '******************'      x
    Case 53
    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0

    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = rose
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = rose
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = rose
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BackColor = rose
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BorderStyle = bordure
    
    haut = 3
    large = 2
    differents = 4
    
    
    '******************'
    '    elle_bar      '    x x x
    '******************'        x
    Case 60
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BorderStyle = 0

    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = orange
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = orange
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = orange
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BackColor = orange
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BorderStyle = bordure
    
    haut = 2
    large = 3
    differents = 4
    deb1 = 1
    fin1 = 3
    
    deb2 = 3
    fin2 = 3
    
    '******************'      x x
    '    elle_bar      '      x
    '******************'      x
    Case 61
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0

    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = orange
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = orange
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = orange
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x).BackColor = orange
    Maine.Controls("Ligne" & position_y + 2)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x).BorderStyle = bordure
    
    haut = 3
    large = 2
    differents = 4
        
    deb1 = 1
    fin1 = 2
    
    deb2 = 1
    fin2 = 1
    
    deb3 = 1
    fin3 = 1
    
    '******************'    x
    '    elle_bar      '    x x x
    '******************'
    Case 62
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = 0
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = 0

    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = orange
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = orange
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = orange
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BackColor = orange
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 2).BorderStyle = bordure
    
    haut = 2
    large = 3
    differents = 4
    
    deb2 = 1
    fin2 = 1
    
    '******************'    x
    '    elle_bar      '    x
    '******************'  x x
    Case 63
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0

    
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = orange
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BackColor = orange
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x + 1).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x).BackColor = orange
    Maine.Controls("Ligne" & position_y + 2)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BackColor = orange
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x + 1).BorderStyle = bordure
    
    haut = 3
    large = 2
    differents = 4

    
    
    '******************'
    '    elle_bar      '    x x x x
    '******************'
    Case 70
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    
   
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 1).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 2).BorderStyle = 0
    
    Maine.Controls("Ligne" & position_y - 1)(position_x + 3).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x + 3).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x + 3).BorderStyle = 0
    

    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = violet
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    Maine.Controls("Ligne" & position_y)(position_x + 1).BackColor = violet
    Maine.Controls("Ligne" & position_y)(position_x + 1).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 1).BorderStyle = bordure
    Maine.Controls("Ligne" & position_y)(position_x + 2).BackColor = violet
    Maine.Controls("Ligne" & position_y)(position_x + 2).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 2).BorderStyle = bordure
    Maine.Controls("Ligne" & position_y)(position_x + 3).BackColor = violet
    Maine.Controls("Ligne" & position_y)(position_x + 3).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x + 3).BorderStyle = bordure
    
    haut = 1
    large = 4
    differents = 2
    
    '                       x
    '******************'    x
    '    elle_bar      '    x
    '******************'    x
    Case 71
    
    Maine.Controls("Ligne" & position_y - 1)(position_x).BackColor = &HFFFFFF
    Maine.Controls("Ligne" & position_y - 1)(position_x).Tag = 0
    Maine.Controls("Ligne" & position_y - 1)(position_x).BorderStyle = 0
    

    
    Maine.Controls("Ligne" & position_y)(position_x).BackColor = violet
    Maine.Controls("Ligne" & position_y)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 1)(position_x).BackColor = violet
    Maine.Controls("Ligne" & position_y + 1)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 1)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 2)(position_x).BackColor = violet
    Maine.Controls("Ligne" & position_y + 2)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 2)(position_x).BorderStyle = bordure
    
    Maine.Controls("Ligne" & position_y + 3)(position_x).BackColor = violet
    Maine.Controls("Ligne" & position_y + 3)(position_x).Tag = ind_piece
    Maine.Controls("Ligne" & position_y + 3)(position_x).BorderStyle = bordure
    
    haut = 4
    large = 1
    differents = 2
    
    End Select
End Function
'**********************************************************************
'   Description
'
'   procédure pour un clic dans le menu fichier nouveau
'
'**********************************************************************
Private Sub new_Click()
    Compteur.Enabled = False
    Unload Maine
    info_partie.Show
End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub



'**********************************************************************
'   Description
'
'   Cette fonction permet de déplacer une piece à gauche ou a droite
'   suivant quelle touche à été pressée
'
'**********************************************************************
Private Sub Translater_Timer()
Dim curs As Byte
Dim lgn As Byte

    'translation à gauche

    If touche = def_touches.touche(0) Then
        
        If position_x - 1 > 0 Then
            For curs = position_y To position_y + haut - 1
              If Controls("ligne" & curs)(position_x - 1).Tag <> 0 And _
                Controls("ligne" & curs)(position_x - 1).Tag <> No_piece Then
                    Exit Sub
                End If
            Next
            
            Call erase_current
            position_x = position_x - 1
            Call piece(position_y, position_x, type_piece & rot, No_piece)
            
        Else
            Exit Sub
        End If
        
    End If
    
    If touche = def_touches.touche(1) Then
        
        If position_x + large + 1 < 12 Then

            For curs = position_y To position_y + haut - 1
                If Controls("ligne" & curs)(position_x + large).Tag <> 0 And _
                Controls("ligne" & curs)(position_x + large).Tag <> No_piece Then
                    Exit Sub
                End If
            Next

            Call erase_current
            position_x = position_x + 1
            Call piece(position_y, position_x, type_piece & rot, No_piece)

        End If
    End If
End Sub
