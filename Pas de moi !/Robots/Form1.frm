VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RobotFight (philcam@free.fr)"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   683
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin PicClip.PictureClip Pic 
      Left            =   9240
      Top             =   1320
      _ExtentX        =   3969
      _ExtentY        =   1588
      _Version        =   393216
      Rows            =   2
      Cols            =   5
      Picture         =   "Form1.frx":0000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8640
      Top             =   6840
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   120
      ScaleHeight     =   20
      ScaleMode       =   0  'User
      ScaleWidth      =   30
      TabIndex        =   0
      Top             =   120
      Width           =   9000
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   600
         Left            =   240
         Top             =   720
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   599
         Left            =   300
         Top             =   300
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   598
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   597
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   596
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   595
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   594
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   593
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   592
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   591
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   590
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   589
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   588
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   587
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   586
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   585
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   584
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   583
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   582
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   581
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   580
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   579
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   578
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   577
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   576
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   575
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   574
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   573
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   572
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   571
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   570
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   569
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   568
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   567
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   566
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   565
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   564
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   563
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   562
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   561
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   560
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   559
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   558
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   557
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   556
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   555
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   554
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   553
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   552
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   551
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   550
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   549
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   548
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   547
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   546
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   545
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   544
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   543
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   542
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   541
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   540
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   539
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   538
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   537
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   536
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   535
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   534
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   533
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   532
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   531
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   530
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   529
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   528
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   527
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   526
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   525
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   524
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   523
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   522
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   521
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   520
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   519
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   518
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   517
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   516
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   515
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   514
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   513
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   512
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   511
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   510
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   509
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   508
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   507
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   506
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   505
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   504
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   503
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   502
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   501
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   500
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   499
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   498
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   497
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   496
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   495
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   494
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   493
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   492
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   491
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   490
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   489
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   488
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   487
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   486
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   485
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   484
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   483
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   482
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   481
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   480
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   479
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   478
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   477
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   476
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   475
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   474
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   473
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   472
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   471
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   470
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   469
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   468
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   467
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   466
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   465
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   464
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   463
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   462
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   461
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   460
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   459
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   458
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   457
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   456
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   455
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   454
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   453
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   452
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   451
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   450
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   449
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   448
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   447
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   446
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   445
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   444
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   443
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   442
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   441
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   440
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   439
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   438
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   437
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   436
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   435
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   434
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   433
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   432
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   431
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   430
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   429
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   428
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   427
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   426
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   425
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   424
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   423
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   422
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   421
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   420
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   419
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   418
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   417
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   416
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   415
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   414
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   413
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   412
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   411
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   410
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   409
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   408
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   407
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   406
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   405
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   404
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   403
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   402
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   401
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   400
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   399
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   398
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   397
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   396
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   395
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   394
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   393
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   392
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   391
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   390
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   389
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   388
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   387
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   386
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   385
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   384
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   383
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   382
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   381
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   380
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   379
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   378
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   377
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   376
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   375
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   374
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   373
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   372
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   371
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   370
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   369
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   368
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   367
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   366
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   365
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   364
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   363
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   362
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   361
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   360
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   359
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   358
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   357
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   356
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   355
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   354
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   353
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   352
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   351
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   350
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   349
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   348
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   347
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   346
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   345
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   344
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   343
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   342
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   341
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   340
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   339
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   338
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   337
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   336
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   335
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   334
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   333
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   332
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   331
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   330
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   329
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   328
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   327
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   326
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   325
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   324
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   323
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   322
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   321
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   320
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   319
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   318
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   317
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   316
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   315
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   314
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   313
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   312
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   311
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   310
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   309
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   308
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   307
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   306
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   305
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   304
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   303
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   302
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   301
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   300
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   299
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   298
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   297
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   296
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   295
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   294
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   293
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   292
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   291
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   290
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   289
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   288
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   287
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   286
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   285
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   284
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   283
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   282
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   281
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   280
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   279
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   278
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   277
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   276
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   275
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   274
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   273
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   272
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   271
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   270
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   269
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   268
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   267
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   266
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   265
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   264
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   263
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   262
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   261
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   260
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   259
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   258
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   257
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   256
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   255
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   254
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   253
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   252
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   251
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   250
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   249
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   248
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   247
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   246
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   245
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   244
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   243
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   242
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   241
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   240
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   239
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   238
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   237
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   236
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   235
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   234
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   233
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   232
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   231
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   230
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   229
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   228
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   227
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   226
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   225
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   224
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   223
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   222
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   221
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   220
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   219
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   218
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   217
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   216
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   215
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   214
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   213
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   212
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   211
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   210
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   209
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   208
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   207
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   206
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   205
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   204
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   203
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   202
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   201
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   200
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   199
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   198
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   197
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   196
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   195
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   194
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   193
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   192
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   191
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   190
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   189
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   188
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   187
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   186
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   185
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   184
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   183
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   182
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   181
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   180
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   179
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   178
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   177
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   176
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   175
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   174
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   173
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   172
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   171
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   170
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   169
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   168
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   167
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   166
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   165
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   164
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   163
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   162
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   161
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   160
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   159
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   158
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   157
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   156
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   155
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   154
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   153
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   152
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   151
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   150
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   149
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   148
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   147
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   146
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   145
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   144
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   143
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   142
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   141
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   140
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   139
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   138
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   137
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   136
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   135
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   134
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   133
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   132
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   131
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   130
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   129
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   128
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   127
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   126
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   125
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   124
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   123
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   122
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   121
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   120
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   119
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   118
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   117
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   116
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   115
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   114
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   113
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   112
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   111
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   110
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   109
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   108
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   107
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   106
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   105
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   104
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   103
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   102
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   101
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   100
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   99
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   98
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   97
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   96
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   95
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   94
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   93
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   92
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   91
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   90
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   89
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   88
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   87
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   86
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   85
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   84
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   83
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   82
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   81
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   80
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   79
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   78
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   77
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   76
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   75
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   74
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   73
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   72
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   71
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   70
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   69
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   68
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   67
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   66
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   65
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   64
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   63
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   62
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   61
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   60
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   59
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   58
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   57
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   56
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   55
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   54
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   53
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   52
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   51
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   50
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   49
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   48
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   47
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   46
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   45
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   44
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   43
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   42
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   41
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   40
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   39
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   38
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   37
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   36
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   35
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   34
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   33
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   32
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   31
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   30
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   29
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   28
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   27
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   26
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   25
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   24
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   23
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   22
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   21
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   20
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   19
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   18
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   17
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   16
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   15
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   14
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   13
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   12
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   11
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   10
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   9
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   8
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   7
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   6
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   5
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   4
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   3
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   2
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Image Dalle 
         Height          =   300
         Index           =   1
         Left            =   600
         Top             =   600
         Width           =   300
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   6240
      Width           =   855
   End
   Begin VB.Image Im 
      Height          =   375
      Index           =   4
      Left            =   7380
      Top             =   6120
      Width           =   375
   End
   Begin VB.Image Im 
      Height          =   375
      Index           =   3
      Left            =   6300
      Top             =   6120
      Width           =   375
   End
   Begin VB.Image Im 
      Height          =   375
      Index           =   2
      Left            =   5220
      Top             =   6120
      Width           =   375
   End
   Begin VB.Image Im 
      Height          =   375
      Index           =   1
      Left            =   4140
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label LbPas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   9
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LbPas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   8
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LbPas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LbPas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   6
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label LbPV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label LbPV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   3
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label LbPV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   2
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label LbPV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   6480
      Width           =   975
   End
   Begin VB.Image Mine 
      Height          =   300
      Index           =   4
      Left            =   2520
      Picture         =   "Form1.frx":6A42
      Top             =   6960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Mine 
      Height          =   300
      Index           =   3
      Left            =   2160
      Picture         =   "Form1.frx":6F34
      Top             =   6960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Mine 
      Height          =   300
      Index           =   2
      Left            =   1800
      Picture         =   "Form1.frx":7426
      Top             =   6960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Mine 
      Height          =   300
      Index           =   1
      Left            =   1440
      Picture         =   "Form1.frx":7918
      Top             =   6960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Rob 
      Height          =   300
      Index           =   4
      Left            =   2520
      Picture         =   "Form1.frx":7E0A
      Top             =   6600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Rob 
      Height          =   300
      Index           =   3
      Left            =   2160
      Picture         =   "Form1.frx":82FC
      Top             =   6600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Rob 
      Height          =   300
      Index           =   2
      Left            =   1800
      Picture         =   "Form1.frx":87EE
      Top             =   6600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Rob 
      Height          =   300
      Index           =   1
      Left            =   1440
      Picture         =   "Form1.frx":8CE0
      Top             =   6600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Sol 
      Height          =   300
      Left            =   480
      Picture         =   "Form1.frx":91D2
      Top             =   6960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Mur 
      Height          =   300
      Left            =   960
      Picture         =   "Form1.frx":96C4
      Top             =   6960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Fond 
      Height          =   375
      Left            =   8640
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#######################################################################

'       La guerre des robots par philcam@free.fr

'   essai de traduction d'un programme en basic (affichage en mode texte)

'   paru dans le magazine Jeux & Stratégie en 198x (je ne sais plus)

'   il est améliorable, alors si vous y comprenez quelque chose

'   n'hésitez pas ! et tenez moi au courant quand même :)

'#######################################################################

'Option Explicit

Private Sub Form_Load()
Dim I As Long
Dim S

Picture1.Width = 900
Picture1.Height = 600

Picture1.ScaleWidth = 30
Picture1.ScaleHeight = 20

Pic.Picture = LoadPicture(App.Path & "\image1.bmp")

Mur.Picture = Pic.GraphicCell(0)
Sol.Picture = Pic.GraphicCell(5)
For I = 1 To 4
Rob(I).Picture = Pic.GraphicCell(I)
Mine(I).Picture = Pic.GraphicCell(I + 5)
Next I

DisRep = 4 'distance de repérage

NbR = 4

32
S = InputBox("Nombre d'instructions par robot ?", , "10")
If S < 10 Or S > 20 Then GoTo 32

ReDim Pas(1 To NbR, 0 To S) 'redimensionne tableau selon infos (10 pas pour l'instant)
'le pas "0" est le pas de secours
Nom(1) = "BersekOne"
Nom(2) = "HoraceAlpha"
Nom(3) = "DeepBlue "
Nom(4) = "Terminator"

For I = 1 To NbR
    PV(I) = 1500
    LbPV(I) = PV(I)
    LbPV(I).Top = 640
    LbPas(I).Top = 660
    Im(I).Top = 610
    Im(I) = Rob(I)
Next I

Form2.Show 1

Placement

Label1.Top = 640
Label2.Top = 640

Open App.Path & "\compterendu.txt" For Output As #2

Print #2, "Résultat du match du : " & Now
Print #2, ""

Form1.Timer1.Enabled = True
End Sub
Private Function CC(PX As Long, PY As Long) As Integer
CC = ((PY - 1) * 30) + PX 'fonction pour traduire des coordonnées en n° d'image
End Function
Private Sub MortRobot(NR)
Dim I As Long

Dalle(CC(X(NR), Y(NR))) = Sol 'le robot disparait
X(NR) = 100
Y(NR) = 100
PV(NR) = 0
LbPV(NR) = 0
LbPas(NR) = ""
For I = 0 To UBound(Pas, 2) 'on efface toutes ces instructions
Pas(NR, I) = ""
Next I
End Sub
Private Sub Placement()
Dim J As Long
Dim G As Long
Dim I As Long
Dim A As Long

'placement de toutes les dalles et attribution de leur pictures et positions
For I = 1 To 600
    Dalle(I) = Sol
    Dalle(I).Width = 1
    Dalle(I).Height = 1
Next I

G = 0
For J = 0 To 19
    For I = 0 To 29
        G = G + 1
        Dalle(G).Left = I
        Dalle(G).Top = J
    Next I
Next J

Randomize

'--------------------------murs horizontaux
For I = 1 To 30
   Dalle(CC(I, 1)) = Mur
   Dalle(CC(I, 20)) = Mur
Next I
'--------------------------murs verticaux
For I = 1 To 20
   Dalle(CC(1, I)) = Mur
   Dalle(CC(30, I)) = Mur
Next I
'--------------------------obstacles
For I = 1 To NbreBlocs 'nombre de blocs
A = Int(Rnd * 600) + 1
Dalle(A) = Mur
Next I
'--------------------------placement robots
For I = 1 To NbR
200 X(I) = Int(Rnd * 30) + 1
Y(I) = Int(Rnd * 20) + 1
If Dalle(CC(X(I), Y(I))) = Mur Then GoTo 200
Dalle(CC(X(I), Y(I))) = Rob(I)
Next I

End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #2
If AffichageCteRendu = True Then Shell "c:\windows\Notepad.exe compterendu.txt", vbMaximizedFocus
End Sub

Private Sub Timer1_Timer()
Dim Temp As String
Dim PasSiOui As String
Dim PasSiNon As String
Dim P As Integer
Dim I, Z, U As Long

Randomize
NumPas = NumPas + 1 'numero d'instruction incrémenté

Label1 = NumPas
Label2.Caption = Label2.Caption + 1
Print #2, "Tour : " & Label2.Caption & " Instruction : " & NumPas

For U = 1 To NbR
    If PV(U) < 1 Then GoTo 102
    
    Temp = Pas(U, NumPas)

    If Left$(Temp, 3) = "TT_" Then
        PasSiOui = Mid$(Temp, 4, 2)
        PasSiNon = Mid$(Temp, 7, 2)
        V = TestProx(U)
            If V = True Then
                Temp = PasSiOui
            Else
                Temp = PasSiNon
            End If
    End If
    
    If Temp = "TH" Then
    Print #2, Nom(U) & " tir horizontalement"
    TirHorizontal U, NumPas, 3
    End If
    
    If Temp = "TV" Then
    Print #2, Nom(U) & " tir verticalement"
    TirVertical U, NumPas, 3
    End If

    If Temp = "AL" Then
    Print #2, Nom(U) & " se déplace aléatoirement"
    P = Int(Rnd * 4) + 1 'choisi une direction au pif
    Deplacement U, NumPas, P, 1
    End If
    
    If Temp = "MI" Then
    Print #2, Nom(U) & " tente de placer une mine";
    PoseMine (U)
    End If
    
    If Left$(Temp, 3) = "DD_" Then
    Print #2, Nom(U) & " se déplace"
    Deplacement U, NumPas, Right$(Temp, 1), 5
    End If
    
    If Temp = "PS" Then
    V = Poursuite(U)
    Deplacement U, NumPas, V, 4
    End If
    
    If Temp = "FT" Then
    V = Fuite(U)
    Deplacement U, NumPas, V, 4
    End If
    
    If Temp = "IN" Then
    Print #2, Nom(U) & " est invisible"
    Dalle(CC(X(U), Y(U))) = Mur
    PV(U) = PV(U) - 20 'coût en energie
    End If
102 'numero de ligne

Next U

Print #2, vbCrLf

'mise à jour compteur PV et vérif mort robot
For I = 1 To NbR
LbPV(I) = PV(I)
LbPas(I) = Pas(I, NumPas)
If PV(I) < 0 Then MortRobot (I)
If PV(I) > 0 Then Print #2, Nom(I) & vbTab & vbTab & PV(I) & vbTab & X(I) & "," & Y(I)
Next I

Print #2, vbCrLf

If NumPas = UBound(Pas, 2) Then NumPas = 0 'remet le compteur d'instructions à 0

'arrête combat s'il ne reste qu'un robot
For I = 1 To NbR
If PV(I) = 0 Then Z = Z + 1
Next I
If Z = NbR - 1 Then Timer1.Enabled = False



End Sub
Private Sub Deplacement(NR, NPas, DI, CE)
Dim XX As Integer
Dim YY As Integer
Dim I As Long

If DI = 0 Then Exit Sub
If DI = 1 Then XX = 0: YY = -1 ' nord
If DI = 2 Then XX = 1: YY = 0 'est
If DI = 3 Then XX = 0: YY = 1 'sud
If DI = 4 Then XX = -1: YY = 0 'ouest
If DI = 12 Then XX = 1: YY = -1 'nord-est
If DI = 23 Then XX = 1: YY = 1 'sud-est
If DI = 34 Then XX = -1: YY = 1 'sud-ouest
If DI = 41 Then XX = -1: YY = -1 'nord-ouest

'pour pas marcher sur d'autres robots ou murs
If Dalle(CC(X(NR) + XX, Y(NR) + YY)) <> Mur And _
    Dalle(CC(X(NR) + XX, Y(NR) + YY)) <> Rob(1) And _
    Dalle(CC(X(NR) + XX, Y(NR) + YY)) <> Rob(2) And _
    Dalle(CC(X(NR) + XX, Y(NR) + YY)) <> Rob(3) And _
    Dalle(CC(X(NR) + XX, Y(NR) + YY)) <> Rob(4) Then
    'marche sur mine ?
    For I = 1 To 4
        If I <> NR Then 'si c'est son propre numéro de mine on passe
            If Dalle(CC(X(NR) + XX, Y(NR) + YY)) = Mine(I) Then
                PV(NR) = PV(NR) - 200 'perte PV
                Print #2, Nom(NR) & " marche sur une mine de " & Nom(I) & ", son pas de programme n°" & NPas & " (" & Pas(NR, NPas) & ") est remplacée par (" & Pas(NR, 0) & ")"
                Pas(NR, NPas) = Pas(NR, 0) 'remplace par instruction de secours
            End If
        End If
    Next I
    Dalle(CC(X(NR) + XX, Y(NR) + YY)) = Rob(NR) 'met image robot sur nouveau emplacement
    Dalle(CC(X(NR), Y(NR))) = Sol 'met image "sol" sur ancien emplacement
    Y(NR) = Y(NR) + YY 'change l'emplacement y du robot
    X(NR) = X(NR) + XX 'change l'emplacement x du robot
If PtEnMoinsSiPasPossible = False Then PV(NR) = PV(NR) - CE 'cout en energie

End If
If PtEnMoinsSiPasPossible = True Then PV(NR) = PV(NR) - CE 'cout en energie

End Sub

Private Sub PoseMine(NR)
Dim DI As Long
Dim XX As Integer
Dim YY As Integer

Randomize
DI = Int(Rnd * 4) + 1 'choisi une direction au pif
If DI = 1 Then XX = 0: YY = -1 ' nord
If DI = 2 Then XX = 1: YY = 0 'est
If DI = 3 Then XX = 0: YY = 1 'sud
If DI = 4 Then XX = -1: YY = 0 'ouest
'pose mine haut
If Dalle(CC(X(NR) + XX, Y(NR) + YY)) = Sol Then
        Dalle(CC(X(NR) + XX, Y(NR) + YY)) = Mine(NR)
        Print #2, " et réussi"
        If PtEnMoinsSiPasPossible = False Then PV(NR) = PV(NR) - 10 'cout en energie,si cette ligne est placée à l'intérieur de la boucle, robot perd energie que si la pose de mine est possible
    Else
        Print #2, " et échoue"
End If

If PtEnMoinsSiPasPossible = True Then PV(NR) = PV(NR) - 10 'cout en energie,si cette ligne est placée à l'intérieur de la boucle, robot perd energie que si la pose de mine est possible

End Sub

Function TestProx(NR) As Boolean
Dim GGtmp, NO As String
Dim PP As Integer
Dim I As Long

Dim M(1 To 4)
GGtmp = Nom(NR) & " utilise son radar et "

PP = 50

For I = 1 To 4
M(I) = Int(Sqr((X(NR) - X(I)) ^ 2 + (Y(NR) - Y(I)) ^ 2))
If M(I) < PP And M(I) <> 0 Then PP = M(I): NO = Nom(I)
Next I

If PP <= DisRep Then TestProx = True: Print #2, GGtmp & "repère " & NO & " à une distance de " & PP
If PP > DisRep Then TestProx = False: Print #2, GGtmp & "ne repère rien"
'call sndPlaySound(ByVal App.Path + "\sound\radar.wav", SND_ASYNC)
PV(NR) = PV(NR) - 4 'coût energie du test
End Function
Function Poursuite(NR)
Dim PP As Integer
Dim M(1 To 4)
Dim I As Long
Dim Dxx, Dyy, Dx, Dy As Long
Dim NO As String

PP = 50

For I = 1 To NbR
M(I) = Int(Sqr((X(NR) - X(I)) ^ 2 + (Y(NR) - Y(I)) ^ 2))
If M(I) < PP And M(I) <> 0 Then PP = M(I): Dxx = X(I): Dyy = Y(I): NO = Nom(I)
Next I

Print #2, Nom(NR) & " poursuit " & NO

Dx = Sgn(Dxx - X(NR))
Dy = Sgn(Dyy - Y(NR))

If Dx = 0 And Dy = -1 Then Poursuite = 1   ' nord
If Dx = 1 And Dy = 0 Then Poursuite = 2   'est
If Dx = 0 And Dy = 1 Then Poursuite = 3   'sud
If Dx = -1 And Dy = 0 Then Poursuite = 4   'ouest
If Dx = 1 And Dy = -1 Then Poursuite = 12 'nord-est
If Dx = 1 And Dy = 1 Then Poursuite = 23 'sud-est
If Dx = -1 And Dy = 1 Then Poursuite = 34  'sud-ouest
If Dx = -1 And Dy = -1 Then Poursuite = 41   'nord-ouest

End Function

Function Fuite(NR)
Dim PP As Integer
Dim M(1 To 4)
Dim I As Long
Dim Dxx, Dyy, Dx, Dy As Long
Dim NO As String

PP = 50

For I = 1 To 4
M(I) = Int(Sqr((X(NR) - X(I)) ^ 2 + (Y(NR) - Y(I)) ^ 2))
If M(I) < PP And M(I) <> 0 Then PP = M(I): Dxx = X(I): Dyy = Y(I): NO = Nom(I)
Next I

Print #2, Nom(NR) & " fuit " & NO

Dx = Sgn(Dxx - X(NR))
Dy = Sgn(Dyy - Y(NR))

If Dx = 0 And Dy = -1 Then Fuite = 4   ' nord
If Dx = 1 And Dy = 0 Then Fuite = 3   'est
If Dx = 0 And Dy = 1 Then Fuite = 2   'sud
If Dx = -1 And Dy = 0 Then Fuite = 1   'ouest
If Dx = 1 And Dy = -1 Then Fuite = 34 'nord-est
If Dx = 1 And Dy = 1 Then Fuite = 41 'sud-est
If Dx = -1 And Dy = 1 Then Fuite = 12  'sud-ouest
If Dx = -1 And Dy = -1 Then Fuite = 23   'nord-ouest

End Function

Private Sub TirHorizontal(NR, NPas, CE)
Dim I As Long

For I = X(NR) + 1 To 30
If Dalle(CC(I, Y(NR))) = Mur Then GoTo 500
If Dalle(CC(I, Y(NR))) = Mine(NR) Then GoTo 490
If Dalle(CC(I, Y(NR))) = Mine(1) Then Dalle(CC(I, Y(NR))) = Sol
If Dalle(CC(I, Y(NR))) = Mine(2) Then Dalle(CC(I, Y(NR))) = Sol
If Dalle(CC(I, Y(NR))) = Mine(3) Then Dalle(CC(I, Y(NR))) = Sol
If Dalle(CC(I, Y(NR))) = Mine(4) Then Dalle(CC(I, Y(NR))) = Sol
If Dalle(CC(I, Y(NR))) = Rob(1) Then PV(1) = PV(1) - 20: Print #2, Nom(1) & " est blessé par " & Nom(NR): GoTo 500
If Dalle(CC(I, Y(NR))) = Rob(2) Then PV(2) = PV(2) - 20: Print #2, Nom(2) & " est blessé par " & Nom(NR): GoTo 500
If Dalle(CC(I, Y(NR))) = Rob(3) Then PV(3) = PV(3) - 20: Print #2, Nom(3) & " est blessé par " & Nom(NR): GoTo 500
If Dalle(CC(I, Y(NR))) = Rob(4) Then PV(4) = PV(4) - 20: Print #2, Nom(4) & " est blessé par " & Nom(NR): GoTo 500
490
Next I
500
For I = X(NR) - 1 To 1 Step -1
If Dalle(CC(I, Y(NR))) = Mur Then GoTo 510
If Dalle(CC(I, Y(NR))) = Mine(NR) Then GoTo 505
If Dalle(CC(I, Y(NR))) = Mine(1) Then Dalle(CC(I, Y(NR))) = Sol
If Dalle(CC(I, Y(NR))) = Mine(2) Then Dalle(CC(I, Y(NR))) = Sol
If Dalle(CC(I, Y(NR))) = Mine(3) Then Dalle(CC(I, Y(NR))) = Sol
If Dalle(CC(I, Y(NR))) = Mine(4) Then Dalle(CC(I, Y(NR))) = Sol
If Dalle(CC(I, Y(NR))) = Rob(1) Then PV(1) = PV(1) - 20: Print #2, Nom(1) & " est blessé par " & Nom(NR): GoTo 510
If Dalle(CC(I, Y(NR))) = Rob(2) Then PV(2) = PV(2) - 20: Print #2, Nom(2) & " est blessé par " & Nom(NR): GoTo 510
If Dalle(CC(I, Y(NR))) = Rob(3) Then PV(3) = PV(3) - 20: Print #2, Nom(3) & " est blessé par " & Nom(NR): GoTo 510
If Dalle(CC(I, Y(NR))) = Rob(4) Then PV(4) = PV(4) - 20: Print #2, Nom(4) & " est blessé par " & Nom(NR): GoTo 510
505
Next I
510
PV(NR) = PV(NR) - CE
End Sub

Private Sub TirVertical(NR, NPas, CE)
Dim I As Long

For I = Y(NR) + 1 To 20
If Dalle(CC(X(NR), I)) = Mur Then GoTo 600
If Dalle(CC(I, Y(NR))) = Mine(NR) Then GoTo 590
If Dalle(CC(X(NR), I)) = Mine(1) Then Dalle(CC(X(NR), I)) = Sol
If Dalle(CC(X(NR), I)) = Mine(2) Then Dalle(CC(X(NR), I)) = Sol
If Dalle(CC(X(NR), I)) = Mine(3) Then Dalle(CC(X(NR), I)) = Sol
If Dalle(CC(X(NR), I)) = Mine(4) Then Dalle(CC(X(NR), I)) = Sol
If Dalle(CC(X(NR), I)) = Rob(1) Then PV(1) = PV(1) - 20: Print #2, Nom(1) & " est blessé par " & Nom(NR): GoTo 600
If Dalle(CC(X(NR), I)) = Rob(2) Then PV(2) = PV(2) - 20: Print #2, Nom(2) & " est blessé par " & Nom(NR): GoTo 600
If Dalle(CC(X(NR), I)) = Rob(3) Then PV(3) = PV(3) - 20: Print #2, Nom(3) & " est blessé par " & Nom(NR): GoTo 600
If Dalle(CC(X(NR), I)) = Rob(4) Then PV(4) = PV(4) - 20: Print #2, Nom(4) & " est blessé par " & Nom(NR): GoTo 600
590
Next I
600

For I = Y(NR) - 1 To 1 Step -1
If Dalle(CC(X(NR), I)) = Mur Then GoTo 610
If Dalle(CC(I, Y(NR))) = Mine(NR) Then GoTo 605
If Dalle(CC(X(NR), I)) = Mine(1) Then Dalle(CC(X(NR), I)) = Sol
If Dalle(CC(X(NR), I)) = Mine(2) Then Dalle(CC(X(NR), I)) = Sol
If Dalle(CC(X(NR), I)) = Mine(3) Then Dalle(CC(X(NR), I)) = Sol
If Dalle(CC(X(NR), I)) = Mine(4) Then Dalle(CC(X(NR), I)) = Sol
If Dalle(CC(X(NR), I)) = Rob(1) Then PV(1) = PV(1) - 20: Print #2, Nom(1) & " est blessé par " & Nom(NR): GoTo 610
If Dalle(CC(X(NR), I)) = Rob(2) Then PV(2) = PV(2) - 20: Print #2, Nom(2) & " est blessé par " & Nom(NR): GoTo 610
If Dalle(CC(X(NR), I)) = Rob(3) Then PV(3) = PV(3) - 20: Print #2, Nom(3) & " est blessé par " & Nom(NR): GoTo 610
If Dalle(CC(X(NR), I)) = Rob(4) Then PV(4) = PV(4) - 20: Print #2, Nom(4) & " est blessé par " & Nom(NR): GoTo 610
605
Next I
610

PV(NR) = PV(NR) - CE
End Sub

