VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   869
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   7380
      Picture         =   "Form1.frx":0342
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   68
      Top             =   3780
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox picTerrainM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   60
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   67
      Top             =   3180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picTerrain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   60
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   66
      Top             =   2580
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picSmallExplodeM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   0
      Left            =   7500
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   65
      Top             =   3960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   7
      Left            =   7500
      Picture         =   "Form1.frx":20C4
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   64
      Top             =   3480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   6
      Left            =   7500
      Picture         =   "Form1.frx":2A36
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   63
      Top             =   3000
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   5
      Left            =   7500
      Picture         =   "Form1.frx":33A8
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   62
      Top             =   2520
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   4
      Left            =   7500
      Picture         =   "Form1.frx":3D1A
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   61
      Top             =   2040
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   3
      Left            =   7500
      Picture         =   "Form1.frx":468C
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   60
      Top             =   1560
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   2
      Left            =   7500
      Picture         =   "Form1.frx":4FFE
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   59
      Top             =   1080
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   1
      Left            =   7500
      Picture         =   "Form1.frx":5970
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   58
      Top             =   600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   0
      Left            =   7500
      Picture         =   "Form1.frx":62E2
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picPodM 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      Picture         =   "Form1.frx":6C54
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   56
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picPod 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      Picture         =   "Form1.frx":7496
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   55
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer timFrame 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   6060
   End
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   54
      Top             =   660
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picExplodeM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   12000
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   52
      Top             =   4320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   14
      Left            =   12000
      Picture         =   "Form1.frx":7CD8
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   51
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   13
      Left            =   12000
      Picture         =   "Form1.frx":BCDA
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   50
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   12
      Left            =   12000
      Picture         =   "Form1.frx":FCDC
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   11
      Left            =   10980
      Picture         =   "Form1.frx":13CDE
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   48
      Top             =   4320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   10
      Left            =   10980
      Picture         =   "Form1.frx":17CE0
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   47
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   10980
      Picture         =   "Form1.frx":1BCE2
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   46
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   10980
      Picture         =   "Form1.frx":1FCE4
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   9960
      Picture         =   "Form1.frx":23CE6
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   44
      Top             =   4320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   9960
      Picture         =   "Form1.frx":27CE8
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   43
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   9960
      Picture         =   "Form1.frx":2BCEA
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   42
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   9960
      Picture         =   "Form1.frx":2FCEC
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   8940
      Picture         =   "Form1.frx":33CEE
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   40
      Top             =   4320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   8940
      Picture         =   "Form1.frx":37CF0
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   39
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   8940
      Picture         =   "Form1.frx":3BCF2
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   8940
      Picture         =   "Form1.frx":3FCF4
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picBossM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   7020
      Picture         =   "Form1.frx":43CF6
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   36
      Top             =   5760
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picBoss 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   6180
      Picture         =   "Form1.frx":45DE0
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   35
      Top             =   5760
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picBorgM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   5400
      Picture         =   "Form1.frx":47ECA
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   34
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBorg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4620
      Picture         =   "Form1.frx":48AAC
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   33
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picFalconM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   4380
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picXwingM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   0
      Left            =   2640
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   31
      Top             =   4920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   16
      Left            =   5880
      Picture         =   "Form1.frx":4968E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   30
      Top             =   4560
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   15
      Left            =   5880
      Picture         =   "Form1.frx":4B9F8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   29
      Top             =   4080
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   14
      Left            =   5880
      Picture         =   "Form1.frx":4DD62
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   28
      Top             =   3600
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   13
      Left            =   5880
      Picture         =   "Form1.frx":500CC
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   12
      Left            =   5880
      Picture         =   "Form1.frx":52436
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   26
      Top             =   2640
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   11
      Left            =   5880
      Picture         =   "Form1.frx":547A0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   10
      Left            =   5880
      Picture         =   "Form1.frx":56B0A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   9
      Left            =   5880
      Picture         =   "Form1.frx":58E74
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   8
      Left            =   5880
      Picture         =   "Form1.frx":5B1DE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   7
      Left            =   4380
      Picture         =   "Form1.frx":5D548
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   6
      Left            =   4380
      Picture         =   "Form1.frx":5F8B2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   5
      Left            =   4380
      Picture         =   "Form1.frx":61C1C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   4
      Left            =   4380
      Picture         =   "Form1.frx":63F86
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   4380
      Picture         =   "Form1.frx":662F0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   4380
      Picture         =   "Form1.frx":6865A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   4380
      Picture         =   "Form1.frx":6A9C4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   4380
      Picture         =   "Form1.frx":6CD2E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   13
      Left            =   2640
      Picture         =   "Form1.frx":6F098
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   12
      Left            =   2640
      Picture         =   "Form1.frx":71952
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   11
      Left            =   2640
      Picture         =   "Form1.frx":7420C
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   10
      Left            =   2640
      Picture         =   "Form1.frx":76AC6
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   9
      Left            =   2640
      Picture         =   "Form1.frx":79380
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   8
      Left            =   2640
      Picture         =   "Form1.frx":7BC3A
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   7
      Left            =   2640
      Picture         =   "Form1.frx":7E4F4
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":80DAE
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   5
      Left            =   1200
      Picture         =   "Form1.frx":83668
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   4
      Left            =   1200
      Picture         =   "Form1.frx":85F22
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   3
      Left            =   1200
      Picture         =   "Form1.frx":887DC
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   2
      Left            =   1200
      Picture         =   "Form1.frx":8B096
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   1
      Left            =   1200
      Picture         =   "Form1.frx":8D950
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   0
      Left            =   1200
      Picture         =   "Form1.frx":9020A
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFrames As Long
Dim lFramesDisp As Long
Dim LastFrameTime As Long
Sub DefineTerrain()
Dim z As Long
Dim j As Long
Dim MaxTerrainHeight As Long
    'set the terrain before play and scroll it
    
    MaxTerrainHeight = picBuffer.Height \ 5
    picTerrain.Move 0, 0, picBuffer.Width * 5, MaxTerrainHeight
    picTerrainM.Move 0, 0, picBuffer.Width * 5, MaxTerrainHeight
    
    
    With Terrain
        .X = picTerrain.Width \ 2
        .StartColor = RGB(60, 30, 30)
        .CurrentColor = .StartColor
        .ColorDirection = 1
        z = 0
        Do
            ReDim Preserve .Points(z)
            If z = 0 Then
                .Points(z).X = 0
            Else
                .Points(z).X = Int(Rnd * 100) + .Points(z - 1).X + 3
            End If
            If z > 0 Then .Points(z).Y = (Int(Rnd * 50) - 25) + .Points(z - 1).Y
            If .Points(z).Y < 0 Then .Points(z).Y = 1
            If .Points(z).Y > MaxTerrainHeight Then .Points(z).Y = MaxTerrainHeight - 2
            z = z + 1
        Loop Until .Points(z - 1).X >= picTerrain.Width
        .NumPoints = z - 2 'the last coordinates go over the picterrain width. scrub those, we'll join the landscape up.
    End With
    
End Sub
Sub DrawTerrainMask()
Dim z As Long
Dim SrcX As Long
Dim SrcY As Long
    
    SrcX = Terrain.Points(0).X
    SrcY = Terrain.Points(0).Y
    
    For z = 1 To Terrain.NumPoints
        DrawLineToTerrain SrcX, SrcY, Terrain.Points(z).X, Terrain.Points(z).Y, 0
        SrcX = Terrain.Points(z).X
        SrcY = Terrain.Points(z).Y
    Next
    DrawLineToTerrain SrcX, SrcY, picTerrain.Width, Terrain.Points(0).Y, 0
End Sub
Sub DrawTerrain()
Dim z As Long
Dim SrcX As Long
Dim SrcY As Long
    
    SrcX = Terrain.Points(0).X
    SrcY = Terrain.Points(0).Y
    
    For z = 1 To Terrain.NumPoints
        DrawLineToTerrain SrcX, SrcY, Terrain.Points(z).X, Terrain.Points(z).Y, Terrain.CurrentColor
        SrcX = Terrain.Points(z).X
        SrcY = Terrain.Points(z).Y
    Next
    DrawLineToTerrain SrcX, SrcY, picTerrain.Width, Terrain.Points(0).Y, Terrain.CurrentColor
    With Terrain
        If .CurrentColor = 7961085 Then
            .ColorDirection = 2 'dimmer
        ElseIf .CurrentColor = 1315884 Then
            .ColorDirection = 1 'brighter
            .CurrentColor = .StartColor
        End If
        If .ColorDirection = 1 Then
            .CurrentColor = AdjustBrightness(.CurrentColor, 3, False)
        Else
            .CurrentColor = AdjustBrightness(.CurrentColor, -3, False)
        End If
    End With
End Sub

Sub TerrainToBuffer()
    If Terrain.X + picBuffer.Width >= picTerrain.Width Then 'take end of terrain and start to scroll properly
        BitBlt picBuffer.hdc, 0, picBuffer.Height - picTerrain.Height - BottomOfScreen, picTerrain.Width - Terrain.X, picTerrain.Height, picTerrainM.hdc, Terrain.X, 0, vbSrcAnd
        BitBlt picBuffer.hdc, picTerrain.Width - Terrain.X + 1, picBuffer.Height - picTerrain.Height - BottomOfScreen, picBuffer.Width - (picTerrain.Width - Terrain.X), picTerrain.Height, picTerrainM.hdc, 0, 0, vbSrcAnd
        BitBlt picBuffer.hdc, 0, picBuffer.Height - picTerrain.Height - BottomOfScreen, picTerrain.Width - Terrain.X, picTerrain.Height, picTerrain.hdc, Terrain.X, 0, vbSrcPaint
        BitBlt picBuffer.hdc, picTerrain.Width - Terrain.X + 1, picBuffer.Height - picTerrain.Height - BottomOfScreen, picBuffer.Width - (picTerrain.Width - Terrain.X), picTerrain.Height, picTerrain.hdc, 0, 0, vbSrcPaint
    Else
        BitBlt picBuffer.hdc, 0, picBuffer.Height - picTerrain.Height - BottomOfScreen, picBuffer.Width, picTerrain.Height, picTerrainM.hdc, Terrain.X, 0, vbSrcAnd
        BitBlt picBuffer.hdc, 0, picBuffer.Height - picTerrain.Height - BottomOfScreen, picBuffer.Width, picTerrain.Height, picTerrain.hdc, Terrain.X, 0, vbSrcPaint
    End If
End Sub
Sub DrawPods()
Dim sAngle As Single
Dim z As Long
    For z = 0 To lNumPods - 1
        With Pods(z)
            If .State = Falling Then
            ElseIf .State = Captured Then
            ElseIf .State = Destroyed Then
            Else 'Free to roam
                If Terrain.X + picBuffer.Width <= picTerrain.Width Then 'no terrain overlap issues
                    If .X >= Terrain.X And .X <= (Terrain.X + picBuffer.Width) Then
                        BitBlt picBuffer.hdc, .X - Terrain.X, .Y, picPod.Width, picPod.Height, picPodM.hdc, 0, 0, vbSrcAnd
                        BitBlt picBuffer.hdc, .X - Terrain.X, .Y, picPod.Width, picPod.Height - 1, picPod.hdc, 0, 0, vbSrcPaint
                    End If
                Else 'add pods at the end and the start of the terrain map! arse!
                    If .X >= Terrain.X Then
                        BitBlt picBuffer.hdc, .X - Terrain.X, .Y, picPod.Width, picPod.Height, picPodM.hdc, 0, 0, vbSrcAnd
                        BitBlt picBuffer.hdc, .X - Terrain.X, .Y, picPod.Width, picPod.Height - 1, picPod.hdc, 0, 0, vbSrcPaint
                    ElseIf .X <= (picBuffer.Width - (picTerrain.Width - Terrain.X)) Then
                        BitBlt picBuffer.hdc, .X + (picTerrain.Width - Terrain.X), .Y, picPod.Width, picPod.Height, picPodM.hdc, 0, 0, vbSrcAnd
                        BitBlt picBuffer.hdc, .X + (picTerrain.Width - Terrain.X), .Y, picPod.Width, picPod.Height - 1, picPod.hdc, 0, 0, vbSrcPaint
                    End If
                End If
                'move if the right time LastTickCount
                If GetTickCount > .LastTickCount + .TimeTilNextPodMove Then
                    .LastTickCount = GetTickCount
                    'move the pods so they appear sentient
                    If Abs(.X - Terrain.Points(.TerrainPoint).X) < 15 Then 'near enough, set another terrain point
                        .TerrainPoint = .TerrainPoint + .TerrainPointDirection
                        If .TerrainPoint < 0 Or .TerrainPoint > Terrain.NumPoints Then
                            .TerrainPoint = IIf(.TerrainPoint <= 0, 0, Terrain.NumPoints - 1)
                            .X = IIf(.TerrainPoint = 0, 0, picTerrain.Width - 1)
                            .TerrainPointDirection = .TerrainPointDirection * -1
                        End If
                    End If
                    'move the pods toward their target
                    .X = .X + .TerrainPointDirection
                    If .Y > (Terrain.Points(.TerrainPoint).Y + (picBuffer.Height - picTerrain.Height - BottomOfScreen - 1)) Then 'going down
                        .Y = .Y - 1
                    Else
                        .Y = .Y + 1
                        If (.Y + picPod.Height) > (picBuffer.Height - BottomOfScreen - 1) Then .Y = picBuffer.Height - picPod.Height - BottomOfScreen - 1
                    End If
                End If
            End If
        End With
    Next
End Sub

Sub CreateLaser(X01 As Long, Y01 As Long, x02 As Long, Y02 As Long, LC As LaserColor)
Dim X As Long
Dim FS As Long
    FS = -1
    For X = 0 To UBound(Lasers)
        If Lasers(X).Free Then
            FS = X
            Exit For
        End If
    Next X
    If FS = -1 Then
        ReDim Preserve Lasers(UBound(Lasers) + 1)
        FS = UBound(Lasers)
    End If
    With Lasers(FS)
        .Free = False
        .Segment = BuildLaser
        .SegmentCount = 1
        .x1 = X01
        .x2 = x02
        .y1 = Y01
        .y2 = Y02
        .MainColor = LC
    End With
End Sub
Sub DrawLasers()
Dim z As Long
    For z = 0 To UBound(Lasers)
        If Lasers(z).Free = False Then
            With Lasers(z)
                If .Segment = BuildLaser Then
                    DrawLineToBuffer .x1, .y1, .x1 + (((.x2 - .x1) / 10) * .SegmentCount), .y2, .MainColor
                    .SegmentCount = .SegmentCount + 1
                    If .SegmentCount >= 11 Then
                        .Segment = HazeEffect
                        .SegmentCount = 1
                        .TimeTillNextAction = 20
                        .LastTickCount = GetTickCount
                        .CurrentColor = .MainColor
                    End If
                ElseIf .Segment = HazeEffect And GetTickCount > (.TimeTillNextAction + .LastTickCount) Then
                    DrawLineToBuffer .x1 + 1, .y1 - .SegmentCount, .x2 - 1, .y2 - .SegmentCount, .CurrentColor
                    DrawLineToBuffer .x1, .y1, .x2, .y2, .MainColor
                    DrawLineToBuffer .x1 + 1, .y1 + .SegmentCount, .x2 - 1, .y2 + .SegmentCount, .CurrentColor
                    .CurrentColor = AdjustBrightness(.CurrentColor, -25)
                    .SegmentCount = .SegmentCount + 1
                    If .SegmentCount >= 11 Then
                        .CurrentColor = .MainColor
                        .LastTickCount = GetTickCount
                        .TimeTillNextAction = 30
                        .Segment = DestroyLaser
                        .SegmentCount = 1
                        
                    End If
                ElseIf .Segment = DestroyLaser And GetTickCount > (.TimeTillNextAction + .LastTickCount) Then
    
                    .CurrentColor = AdjustBrightness(.CurrentColor, -20)
                    DrawLineToBuffer .x1 + (((.x2 - .x1) / 10) * .SegmentCount), .y1, .x2, .y2, .CurrentColor
                    
                    .SegmentCount = .SegmentCount + 1
                    If .SegmentCount >= 11 Then
                        .Free = True
                    End If
                End If
            End With
        End If
    Next
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        If Defender.ChangingDirection Then Exit Sub
        If Defender.Ship = falcon Then
            If Defender.Direction = DefenderRight Then
                CreateLaser Defender.X + picFalcon(0).Width, Defender.Y + (picFalcon(0).Height \ 2) + 3, Defender.X + picFalcon(0).Width + LaserLength, Defender.Y + (picFalcon(0).Height \ 2) + 3, greenlaser
            Else
                CreateLaser Defender.X, Defender.Y + (picFalcon(0).Height \ 2) + 3, Defender.X - LaserLength, Defender.Y + (picFalcon(0).Height \ 2) + 3, greenlaser
            End If
        Else
            If Defender.Direction = DefenderRight Then
                CreateLaser Defender.X + picXwing(0).Width - 30, Defender.Y + 2, Defender.X + picXwing(0).Width + LaserLength, Defender.Y + 2, redlaser
                CreateLaser Defender.X + picXwing(0).Width - 30, Defender.Y + 32, Defender.X + picXwing(0).Width + LaserLength, Defender.Y + 32, redlaser
            Else
                CreateLaser Defender.X + 30, Defender.Y + 2, Defender.X - LaserLength, Defender.Y + 2, redlaser
                CreateLaser Defender.X + 30, Defender.Y + 32, Defender.X - LaserLength, Defender.Y + 32, redlaser
            End If
        End If
    End If
End Sub
Sub CheckKeys()
    If Not Defender.ChangingDirection Then
        If GetAsyncKeyState(32) <> 0 Then
            Defender.ChangingDirection = True
        End If
    End If
    If GetAsyncKeyState(83) <> 0 Then
        Defender.Y = Defender.Y - Defender.yMovement
        If Defender.Y < (TopOfScreen + 1) Then Defender.Y = (TopOfScreen + 1)
    End If
    If GetAsyncKeyState(88) <> 0 Then
        Defender.Y = Defender.Y + Defender.yMovement
        If Defender.Ship = falcon Then
            If Defender.Y > (picBuffer.Height - picFalcon(0).Height - BottomOfScreen - 1) Then Defender.Y = picBuffer.Height - picFalcon(0).Height - BottomOfScreen - 1
        Else
            If Defender.Y > (picBuffer.Height - picXwing(0).Height - BottomOfScreen - 1) Then Defender.Y = picBuffer.Height - picXwing(0).Height - BottomOfScreen - 1
        End If
    End If
    If Defender.ChangingDirection Then Exit Sub
    
    If GetAsyncKeyState(65) <> 0 Then
        AdjustStarSpeed
        AdjustAnyLasers
        AdjustTerrain
    End If
End Sub
Sub AdjustTerrain()
    If Defender.Direction = DefenderRight Then
        Terrain.X = Terrain.X + 15
        If Terrain.X > picTerrain.Width Then Terrain.X = Terrain.X - picTerrain.Width
    Else
        Terrain.X = Terrain.X - 15
        If Terrain.X < 0 Then Terrain.X = Terrain.X + picTerrain.Width
    End If
End Sub
Sub AdjustAnyLasers()
Dim z As Long
    For z = 0 To UBound(Lasers)
        If Lasers(z).Free = False Then
            With Lasers(z)
                If Defender.Direction = DefenderRight Then
                    .x1 = .x1 - 3
                    .x2 = .x2 - 3
                Else
                    .x1 = .x1 + 3
                    .x2 = .x2 + 3
                End If
            End With
        End If
    Next
End Sub
Sub AdjustStarSpeed()
Dim z As Long
    For z = 0 To NumStars - 1
        With StarFieldStars(z)
            If Defender.Direction = DefenderRight Then
                .X = .X - .Speed
                If .X < 0 Then .X = picBuffer.Width
            Else
                .X = .X + .Speed
                If .X > picBuffer.Width Then .X = 0

            End If
        End With
    Next
End Sub
Private Sub Form_Load()
    'Me.Move 0, 0, Screen.Width, Screen.Height
    picBuffer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    picBlank.Move 0, 0, picBuffer.Width, picBuffer.Height
    
    Me.Show
    DoEvents
    
    ClearBuffer
    pTextOut "Initialising...", "Impact", 22, False, 20, 20, RGB(250, 250, 200)
    BufferToScreen

    
    Initialise
    
    'CreateExplosion 20, 20
    'CreateExplosion 100, 100
    
    timFrame.Enabled = True
    lFrames = 0
    LastFrameTime = GetTickCount
    
End Sub
Sub Initialise()

    Randomize Timer ^ CInt(Format(Now, "ss"))

    ReDim Explosion(0):         Explosion(0).Free = True
    ReDim SmallExplosion(0):    SmallExplosion(0).Free = True
    ReDim Lasers(0):            Lasers(0).Free = True
    
    CreateMasks

    With Defender
        .Ship = falcon
        .X = (picBuffer.Width \ 2) - (picFalcon(0).Width \ 2)
        .Y = (picBuffer.Height \ 2) - (picFalcon(0).Height \ 2)
        .Direction = DefenderRight
        .ChangingDirection = False
        .CurrentFrame = 0
        .yMovement = 2
    End With
    
    SetupStarField
    DefineTerrain
    DrawTerrainMask
    SetupPods
End Sub

Sub SetupPods()
Dim z As Long
Dim j As Long
Dim bTerrainPointLocated As Boolean

    'pods will move between the terrain coordinates
    lNumPods = 20
    ReDim Pods(lNumPods)
    
    Do 'make sure each pod has an anchor terrain point
        For z = 0 To lNumPods - 1
            With Pods(z)
                .X = Int(Rnd * (picTerrain.Width - (2 * picPod.Width))) + picPod.Width
                'find nearest terrain point
                bTerrainPointLocated = False
                For j = 0 To Terrain.NumPoints - 1
                    If Abs(Terrain.Points(j).X - .X) < 50 Then
                        bTerrainPointLocated = True
                        Exit For
                    End If
                Next
                If bTerrainPointLocated Then
                    .Y = (Terrain.Points(j).Y - (picPod.Height \ 2)) + (picBuffer.Height - picTerrain.Height) + BottomOfScreen + 2
                    .TerrainPoint = j
                    If .X > Terrain.Points(j).X Then
                        .TerrainPointDirection = -1
                    Else
                        .TerrainPointDirection = 1
                    End If
                Else
                    Exit For
                End If
                If .Y + picPod.Height > (picBuffer.Height + 1 - BottomOfScreen) Then .Y = picBuffer.Height - picPod.Height - 1 - BottomOfScreen
                .LastTickCount = GetTickCount
                .TimeTilNextPodMove = 10 + (Int(Rnd * 300))
            End With
        Next
    Loop Until bTerrainPointLocated
End Sub
Sub SetupStarField()
Dim X As Long
Dim RNDCOL As Long
    NumStars = 100

    ReDim StarFieldStars(NumStars)
    For X = 0 To NumStars - 1
        With StarFieldStars(X)
            .X = Int(Rnd * picBuffer.Width)
            .Y = Int(Rnd * (picBuffer.Height - (TopOfScreen + BottomOfScreen + 4))) + TopOfScreen
            .Speed = 2 + Int(Rnd * 12)
            RNDCOL = Int(Rnd * 250)
            .StartColor = RGB(RNDCOL, RNDCOL, RNDCOL)
            .CurrentColor = .StartColor
            .ColorDirection = 1 + Int(Rnd * 2)
        End With
    Next
End Sub
Sub CreateMasks()
Dim lF As Long
    For lF = 0 To picXwing.Count - 1
        If lF > 0 Then
            Load picXwingM(lF)
        End If
        CreateMask picXwing(lF).hdc, picXwingM(lF).hdc, picXwing(lF).Width, picXwing(lF).Height
    Next
    For lF = 0 To picFalcon.Count - 1
        If lF > 0 Then
            Load picFalconM(lF)
        End If
        CreateMask picFalcon(lF).hdc, picFalconM(lF).hdc, picFalcon(lF).Width, picFalcon(lF).Height
    Next
    For lF = 0 To picExplode.Count - 1
        If lF > 0 Then
            Load picExplodeM(lF)
        End If
        CreateMask picExplode(lF).hdc, picExplodeM(lF).hdc, picExplode(lF).Width, picExplode(lF).Height
    Next
    For lF = 0 To picSmallExplode.Count - 1
        If lF > 0 Then
            Load picSmallExplodeM(lF)
        End If
        CreateMask picSmallExplode(lF).hdc, picSmallExplodeM(lF).hdc, picSmallExplode(lF).Width, picSmallExplode(lF).Height
    Next
End Sub
Sub CreateMask(picInHDC As Long, picOutHDC As Long, lW As Long, lH As Long)
Dim X As Long
Dim Y As Long
    For X = 0 To lW
        For Y = 0 To lH
            If GetPixel(picInHDC, X, Y) <> 0 Then
                SetPixel picOutHDC, X, Y, 0
            End If
        Next
    Next
End Sub
Sub ClearBuffer()
    BitBlt picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub
Sub BufferToScreen()
    BitBlt Me.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBuffer.hdc, 0, 0, vbSrcCopy
    'Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timFrame.Enabled = False
End Sub

Private Sub timFrame_Timer()
    ClearBuffer
    CreateFrameToBuffer
    BufferToScreen
    CheckKeys
End Sub
Sub CreateFrameToBuffer()
    DrawStarField
    DrawTerrain
    TerrainToBuffer
    DrawPods
    'draw enemy
    'draw friendlys
    
    DrawLasers
    DrawDefenderToBuffer
    ExplosionsToBuffer
    lFrames = lFrames + 1
    TopOfScreenToBuffer
    BottomOfScreenToBuffer
    If GetTickCount >= LastFrameTime + 1000 Then
        lFramesDisp = lFrames
        lFrames = 0
        LastFrameTime = GetTickCount
    End If
    pTextOut lFramesDisp & "fps", "Arial", 8, True, 3, 3, RGB(100, 100, 100)
End Sub
Sub BottomOfScreenToBuffer()
    DrawLineToBuffer 0, picBuffer.Height - BottomOfScreen, picBuffer.Width, picBuffer.Height - BottomOfScreen, RGB(100, 50, 50)
    pTextOut "[A] Thrust  [S] Up  [X] Down  [ENTER] Fire  [SPACE] Change Direction", "Tahoma", 8, True, 5, picBuffer.Height - 16, RGB(51, 100, 51)
End Sub
Sub TopOfScreenToBuffer()
Dim Ratio As Single
Dim RectW As Long
Const RectH As Long = 34
Dim RectX As Long
Dim z As Long
Dim Xlen As Long
Dim Y As Long
Dim s As Long
Dim l As Long
Dim lS As Long
    BitBlt picBuffer.hdc, picBuffer.Width - picLogo.Width - 3, 9, picLogo.Width, picLogo.Height, picLogo.hdc, 0, 0, vbSrcCopy

    Ratio = (picBuffer.Height - (TopOfScreen + BottomOfScreen)) / RectH
    RectW = picTerrain.Width \ Ratio
    RectX = (picBuffer.Width \ 2) - (RectW \ 2)
    
    DrawRectangle RectX, 3, RectX + RectW, 3 + RectH, RGB(100, 50, 50)
    
    lS = 0
    l = RectX + (Terrain.X \ Ratio) + (picBuffer.Width \ Ratio)
    If l > RectX + RectW Then
        lS = l - (RectX + RectW)
        l = RectX + RectW
    End If
    s = RectX + (Terrain.X \ Ratio)
    DrawLineToBuffer s, 3, l, 3, RGB(200, 100, 100)
    DrawLineToBuffer s, 2 + RectH, l, 2 + RectH, RGB(200, 100, 100)
    SetPixel picBuffer.hdc, s, 4, RGB(200, 100, 100)
    SetPixel picBuffer.hdc, s, 1 + RectH, RGB(200, 100, 100)
    If lS > 0 Then
        DrawLineToBuffer RectX, 3, lS + RectX, 3, RGB(200, 100, 100)
        DrawLineToBuffer RectX, 2 + RectH, lS + RectX, 2 + RectH, RGB(200, 100, 100)
        SetPixel picBuffer.hdc, lS + RectX - 1, 4, RGB(200, 100, 100)
        SetPixel picBuffer.hdc, lS + RectX - 1, 1 + RectH, RGB(200, 100, 100)
        
    Else
        SetPixel picBuffer.hdc, l - 1, 4, RGB(200, 100, 100)
        SetPixel picBuffer.hdc, l - 1, 1 + RectH, RGB(200, 100, 100)
    End If
    
    Xlen = 2
    For Y = RectH + 2 To 3 Step -5
        DrawLineToBuffer RectX, Y, RectX - Xlen, Y, RGB(100, 50, 50)
        DrawLineToBuffer RectX + RectW, Y, RectX + RectW + Xlen, Y, RGB(100, 50, 50)
        Xlen = Xlen + 2
    Next
    
    DrawLineToBuffer 0, TopOfScreen, picBuffer.Width, TopOfScreen, RGB(100, 50, 50)
    
    '================= mini screen =====================

    'draw pods
    For z = 0 To lNumPods - 1
        With Pods(z)
        SetPixel picBuffer.hdc, RectX + (.X \ Ratio), (.Y \ Ratio), RGB(200, 255, 200)
        SetPixel picBuffer.hdc, RectX + (.X \ Ratio), (.Y \ Ratio) + 1, RGB(100, 155, 100)
        End With
    Next z
    'show ship
    l = RectX + (Terrain.X \ Ratio) + (picBuffer.Width \ 2 \ Ratio)
    If l > (RectX + RectW) Then
        l = RectX + (l - (RectX + RectW))
    End If
    SetPixel picBuffer.hdc, l - 1, (Defender.Y \ Ratio), RGB(190, 190, 190)
    SetPixel picBuffer.hdc, l, (Defender.Y \ Ratio), vbWhite
    SetPixel picBuffer.hdc, l + 1, (Defender.Y \ Ratio), vbWhite
    SetPixel picBuffer.hdc, l + 2, (Defender.Y \ Ratio), RGB(190, 190, 190)
End Sub
Sub DrawRectangle(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)
Dim hRPen As Long
    hRPen = CreatePen(0, 1, lColor)
    DeleteObject SelectObject(picBuffer.hdc, hRPen)
    Rectangle picBuffer.hdc, x1, y1, x2, y2
    DeleteObject hRPen
End Sub
Sub DrawLineToBuffer(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)
Dim Point As POINTAPI
Dim hPen As Long, hOldPen As Long
    hPen = CreatePen(0, 1, lColor)
    hOldPen = SelectObject(picBuffer.hdc, hPen)
    Point.X = x1: Point.Y = y1
    MoveToEx picBuffer.hdc, x1, y1, Point
    LineTo picBuffer.hdc, x2, y2
    DeleteObject SelectObject(hdc, hOldPen)

End Sub
Sub DrawStarField()
Dim z As Long
Dim Haze As Long
    For z = 0 To NumStars - 1
        With StarFieldStars(z)
            Haze = AdjustBrightness(.CurrentColor, -30, True)
            SetPixel picBuffer.hdc, .X, .Y, .CurrentColor
            SetPixel picBuffer.hdc, .X - 1, .Y, Haze
            SetPixel picBuffer.hdc, .X, .Y - 1, Haze
            SetPixel picBuffer.hdc, .X + 1, .Y, Haze
            SetPixel picBuffer.hdc, .X, .Y + 1, Haze
            
            If .CurrentColor = 16777215 Then
                .ColorDirection = 2 'dimmer
            ElseIf .CurrentColor = 0 Then
                .ColorDirection = 1 'brighter
                .CurrentColor = RGB(30, 30, 30)
            End If
            If .ColorDirection = 1 Then
                .CurrentColor = AdjustBrightness(.CurrentColor, 3, True)
            Else
                .CurrentColor = AdjustBrightness(.CurrentColor, -3, True)
            End If
        End With
    Next
End Sub
Sub DrawDefenderToBuffer()
    With Defender
        If .Ship = falcon Then
            BitBlt picBuffer.hdc, .X, .Y, picFalcon(0).Width, picFalcon(0).Height, picFalconM(.CurrentFrame).hdc, 0, 0, vbSrcAnd
            BitBlt picBuffer.hdc, .X, .Y, picFalcon(0).Width, picFalcon(0).Height, picFalcon(.CurrentFrame).hdc, 0, 0, vbSrcPaint
        Else
            BitBlt picBuffer.hdc, .X, .Y, picXwing(0).Width, picXwing(0).Height, picXwingM(.CurrentFrame).hdc, 0, 0, vbSrcAnd
            BitBlt picBuffer.hdc, .X, .Y, picXwing(0).Width, picXwing(0).Height, picXwing(.CurrentFrame).hdc, 0, 0, vbSrcPaint
        End If
        If .ChangingDirection Then
            If .Direction = DefenderRight Then
                .CurrentFrame = .CurrentFrame + 1
                If .Ship = falcon Then
                    If .CurrentFrame >= picFalcon.Count - 1 Then
                        .CurrentFrame = picFalcon.Count - 1
                        .ChangingDirection = False
                        .Direction = DefenderLeft
                    End If
                Else
                    If .CurrentFrame >= picXwing.Count - 1 Then
                        .CurrentFrame = picXwing.Count - 1
                        .ChangingDirection = False
                        .Direction = DefenderLeft
                    End If
                End If
            Else
                .CurrentFrame = .CurrentFrame - 1
                If .CurrentFrame <= 0 Then
                    .CurrentFrame = 0
                    .ChangingDirection = False
                    .Direction = DefenderRight
                End If
            End If
        End If
    End With
End Sub
Sub ExplosionsToBuffer()
Dim lExp As Long
    For lExp = 0 To UBound(Explosion)
        With Explosion(lExp)
            If Not .Free Then
                BitBlt picBuffer.hdc, .X, .Y, picExplode(.CurrentFrame).Width, picExplode(.CurrentFrame).Height, picExplodeM(.CurrentFrame).hdc, 0, 0, vbSrcAnd
                BitBlt picBuffer.hdc, .X, .Y, picExplode(.CurrentFrame).Width, picExplode(.CurrentFrame).Height, picExplode(.CurrentFrame).hdc, 0, 0, vbSrcPaint
                If (.LastFrameTime + .TimeTillNextFrame) <= GetTickCount Then
                    .LastFrameTime = GetTickCount
                    .CurrentFrame = .CurrentFrame + 1
                    If .CurrentFrame > picExplode.Count - 1 Then
                        .Free = True
                    End If
                End If
            End If
        End With
    Next
    For lExp = 0 To UBound(SmallExplosion)
        With SmallExplosion(lExp)
            If Not .Free Then
                BitBlt picBuffer.hdc, .X, .Y, picSmallExplode(.CurrentFrame).Width, picSmallExplode(.CurrentFrame).Height, picSmallExplodeM(.CurrentFrame).hdc, 0, 0, vbSrcAnd
                BitBlt picBuffer.hdc, .X, .Y, picSmallExplode(.CurrentFrame).Width, picSmallExplode(.CurrentFrame).Height, picSmallExplode(.CurrentFrame).hdc, 0, 0, vbSrcPaint
                If (.LastFrameTime + .TimeTillNextFrame) <= GetTickCount Then
                    .LastFrameTime = GetTickCount
                    .CurrentFrame = .CurrentFrame + 1
                    If .CurrentFrame > picSmallExplode.Count - 1 Then
                        .Free = True
                    End If
                End If
            End If
        End With
    Next

End Sub


Sub pTextOut(sIn As String, sFont As String, iFontSize As Integer, bFontBold As Boolean, xPos As Integer, yPos As Integer, lColor As Long)
    
    SetTextColor picBuffer.hdc, 0
    picBuffer.Font = sFont
    picBuffer.FontSize = iFontSize
    picBuffer.FontBold = bFontBold
    
    TextOut picBuffer.hdc, xPos + 1, yPos + 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos - 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos + 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos + 1, yPos - 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos + 1, yPos, sIn, Len(sIn)
    
    SetTextColor picBuffer.hdc, lColor
    TextOut picBuffer.hdc, xPos, yPos, sIn, Len(sIn)

End Sub

Sub DrawLineToTerrain(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)
Dim hRPen As Long
Dim Point As POINTAPI
    picTerrain.ForeColor = lColor
    Point.X = x1: Point.Y = y1
    MoveToEx picTerrain.hdc, x1, y1, Point
    LineTo picTerrain.hdc, x2, y2
End Sub
Sub ApplyBlend()
Dim Blend As BLENDFUNCTION
Dim BlendPtr As Long
    Blend.SourceConstantAlpha = 130
    
    CopyMemory BlendPtr, Blend, 4
    
    AlphaBlend picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, picBuffer.Width, picBuffer.Height, BlendPtr
End Sub
Sub CreateExplosion(xP As Long, yP As Long)
Dim bFreeSlot As Boolean
Dim a As Long
    bFreeSlot = False
    For a = 0 To UBound(Explosion)
        If Explosion(a).Free = True Then
            bFreeSlot = True
            Exit For
        End If
    Next
    If Not bFreeSlot Then
        ReDim Preserve Explosion(UBound(Explosion) + 1)
        a = UBound(Explosion)
    End If
    
    With Explosion(a)
        .Free = False
        .X = xP - 30 'adjust coords of explosion to appear in middle
        .Y = yP - 54
        .TimeTillNextFrame = 50
        .LastFrameTime = GetTickCount
        .CurrentFrame = 0
    End With
    
End Sub

Sub CreateSmallExplosion(xP As Long, yP As Long)
Dim bFreeSlot As Boolean
Dim a As Long
    bFreeSlot = False
    For a = 0 To UBound(SmallExplosion)
        If SmallExplosion(a).Free = True Then
            bFreeSlot = True
            Exit For
        End If
    Next
    If Not bFreeSlot Then
        ReDim Preserve SmallExplosion(UBound(SmallExplosion) + 1)
        a = UBound(SmallExplosion)
    End If
    
    With SmallExplosion(a)
        .Free = False
        .X = xP - 14 'adjust coords of explosion to appear in middle
        .Y = yP - 14
        .TimeTillNextFrame = 50
        .LastFrameTime = GetTickCount
        .CurrentFrame = 0
    End With
    
End Sub
