VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10785
   ClientLeft      =   2190
   ClientTop       =   360
   ClientWidth     =   10575
   FillColor       =   &H80000016&
   ForeColor       =   &H80000016&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Left            =   2160
      Top             =   8880
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   2040
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8640
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   9955
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Timer Timer7 
         Interval        =   1
         Left            =   8640
         Top             =   5400
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8640
         Top             =   4920
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8640
         Top             =   4440
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3960
         ScaleHeight     =   975
         ScaleWidth      =   4335
         TabIndex        =   9
         Top             =   8640
         Width           =   4335
         Begin VB.Timer Timer4 
            Interval        =   1
            Left            =   240
            Top             =   240
         End
      End
      Begin VB.Shape Shape78 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3240
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape77 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   360
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape76 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape75 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape74 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape73 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape72 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   720
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape71 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape70 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape69 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   8160
         Width           =   135
      End
      Begin VB.Shape Shape68 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3240
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape67 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   360
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape66 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape65 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape64 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape63 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape62 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   720
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape61 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape60 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape59 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   135
      End
      Begin VB.Shape Shape58 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3240
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape57 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   360
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape56 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape55 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape54 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape53 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape52 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   720
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape51 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape50 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape49 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Shape48 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape47 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape43 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape41 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape45 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape46 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape44 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape42 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape40 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   360
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape39 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3240
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIERS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   540
         Left            =   4890
         TabIndex        =   12
         Top             =   5640
         Width           =   2640
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   1080
         Left            =   240
         TabIndex        =   11
         Top             =   5160
         Width           =   3615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6840
         TabIndex        =   10
         Top             =   7080
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   1080
         Left            =   240
         TabIndex        =   8
         Top             =   7080
         Width           =   3615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   1080
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   1080
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Shape Shape31 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         FillColor       =   &H00004040&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INSTICK  MINI-GROCERY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   4185
         TabIndex        =   5
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REPORTS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   540
         Left            =   4440
         TabIndex        =   4
         Top             =   7560
         Width           =   3480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STOCKS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   540
         Left            =   4380
         TabIndex        =   3
         Top             =   3720
         Width           =   3540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSACTION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   480
         Left            =   4560
         TabIndex        =   2
         Top             =   1800
         Width           =   3165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   480
         Left            =   8880
         TabIndex        =   1
         ToolTipText     =   "Exit !"
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   7440
         Width           =   3495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Shape Shape22 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   8280
         Top             =   9000
         Width           =   615
      End
      Begin VB.Shape Shape21 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   1335
         Left            =   8880
         Top             =   7920
         Width           =   255
      End
      Begin VB.Shape Shape20 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   1335
         Left            =   8880
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape Shape19 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   8280
         Top             =   600
         Width           =   855
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H000000C0&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8280
         Top             =   8880
         Width           =   975
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H000000C0&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8280
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   3600
         Width           =   3495
      End
      Begin VB.Shape Shape25 
         BorderColor     =   &H00808000&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   4680
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Shape Shape26 
         BorderColor     =   &H00808000&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   4800
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Shape Shape27 
         BorderColor     =   &H00808000&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   4800
         Shape           =   3  'Circle
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Shape Shape24 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7080
         Top             =   7680
         Width           =   2055
      End
      Begin VB.Shape Shape23 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7800
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H000000C0&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   7320
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H000000C0&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   7560
         Width           =   2055
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H000000C0&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   8760
         Shape           =   4  'Rounded Rectangle
         Top             =   7560
         Width           =   495
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H000000C0&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   8760
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00000040&
         BorderWidth     =   8
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   6000
         Top             =   2640
         Width           =   375
      End
      Begin VB.Shape Shape30 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         FillColor       =   &H00004040&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6960
         Width           =   3855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         FillColor       =   &H00004040&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Shape Shape18 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   9000
         Width           =   3975
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   8880
         Width           =   3975
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   3960
         Top             =   240
         Width           =   4335
      End
      Begin VB.Shape Shape28 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004040&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   4080
         Top             =   360
         Width           =   4335
      End
      Begin VB.Shape Shape17 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   600
         Width           =   4095
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   4095
      End
      Begin VB.Shape Shape29 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004040&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   4080
         Top             =   8760
         Width           =   4335
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         FillColor       =   &H00004040&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5040
         Width           =   3855
      End
      Begin VB.Shape Shape32 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   5520
         Width           =   3495
      End
      Begin VB.Shape Shape33 
         BorderColor     =   &H00808000&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   4800
         Shape           =   3  'Circle
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Shape Shape34 
         BorderColor     =   &H00000040&
         BorderWidth     =   8
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   6000
         Top             =   4560
         Width           =   375
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00000040&
         BorderWidth     =   8
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   6000
         Top             =   6480
         Width           =   375
      End
      Begin VB.Shape Shape35 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3720
         Top             =   1920
         Width           =   855
      End
      Begin VB.Shape Shape36 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3720
         Top             =   3840
         Width           =   855
      End
      Begin VB.Shape Shape37 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3720
         Top             =   5760
         Width           =   855
      End
      Begin VB.Shape Shape38 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3720
         Top             =   7680
         Width           =   855
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String, I As Integer, sel As Boolean, Intcount As Integer, Z As Integer
Dim buf(100, 4), rel As Single
Dim x0 As Integer
Private Sub Form_Load()
    sel = True
    Me.BackColor = vbDesktop
    Frame1.BackColor = vbDesktop
    Frame1.Move (Screen.Width - Frame1.Width) / 2, (Screen.Height - Frame1.Height) / 2
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &HC0C0FF
    Label3.ForeColor = &HC0C0FF
    Label4.ForeColor = &HC0C0FF
    Label11.ForeColor = &HC0C0FF
    
    Shape2.FillColor = &H40
    Shape3.FillColor = &H40
    Shape4.FillColor = &H40
    Shape32.FillColor = &H40

    Shape2.BorderWidth = 6
    Shape3.BorderWidth = 6
    Shape4.BorderWidth = 6
    Shape32.BorderWidth = 6

    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    
    Shape39.FillColor = &H808000 'Lights of description
    Shape40.FillColor = &H808000
    Shape41.FillColor = &H808000
    Shape42.FillColor = &H808000
    Shape43.FillColor = &H808000
    Shape44.FillColor = &H808000
    Shape45.FillColor = &H808000
    Shape46.FillColor = &H808000
    Shape47.FillColor = &H808000
    Shape48.FillColor = &H808000
    
    Shape49.FillColor = &H808000
    Shape50.FillColor = &H808000
    Shape51.FillColor = &H808000
    Shape52.FillColor = &H808000
    Shape53.FillColor = &H808000
    Shape54.FillColor = &H808000
    Shape55.FillColor = &H808000
    Shape56.FillColor = &H808000
    Shape57.FillColor = &H808000
    Shape58.FillColor = &H808000
    
    Shape59.FillColor = &H808000
    Shape60.FillColor = &H808000
    Shape61.FillColor = &H808000
    Shape62.FillColor = &H808000
    Shape63.FillColor = &H808000
    Shape64.FillColor = &H808000
    Shape65.FillColor = &H808000
    Shape66.FillColor = &H808000
    Shape67.FillColor = &H808000
    Shape68.FillColor = &H808000
    
    Shape69.FillColor = &H808000
    Shape70.FillColor = &H808000
    Shape71.FillColor = &H808000
    Shape72.FillColor = &H808000
    Shape73.FillColor = &H808000
    Shape74.FillColor = &H808000
    Shape75.FillColor = &H808000
    Shape76.FillColor = &H808000
    Shape77.FillColor = &H808000
    Shape78.FillColor = &H808000
End Sub


Private Sub Label1_Click()
   Label1.Enabled = False
   Timer5.Enabled = True
End Sub


Private Sub Label11_Click()
Label11.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label11.ForeColor = vbBlack
    Label2.ForeColor = &HC0C0FF
    Label3.ForeColor = &HC0C0FF
    Label4.ForeColor = &HC0C0FF

    Shape32.FillColor = &HFFFFC0
    Shape32.BorderWidth = 12

    Shape3.FillColor = &H40
    Shape4.FillColor = &H40

    Shape3.BorderWidth = 6
    Shape4.BorderWidth = 6
    
    Label10.Caption = "Use this Button to Add and Modify SUPPLIERS"
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    
    Shape39.FillColor = &H808000 'Lights of description
    Shape40.FillColor = &H808000
    Shape41.FillColor = &H808000
    Shape42.FillColor = &H808000
    Shape43.FillColor = &H808000
    Shape44.FillColor = &H808000
    Shape45.FillColor = &H808000
    Shape46.FillColor = &H808000
    Shape47.FillColor = &H808000
    Shape48.FillColor = &H808000
    
    Shape49.FillColor = &H808000
    Shape50.FillColor = &H808000
    Shape51.FillColor = &H808000
    Shape52.FillColor = &H808000
    Shape53.FillColor = &H808000
    Shape54.FillColor = &H808000
    Shape55.FillColor = &H808000
    Shape56.FillColor = &H808000
    Shape57.FillColor = &H808000
    Shape58.FillColor = &H808000
    
    Shape59.FillColor = &HFFFF00
    Shape60.FillColor = &HFFFF00
    Shape61.FillColor = &HFFFF00
    Shape62.FillColor = &HFFFF00
    Shape63.FillColor = &HFFFF00
    Shape64.FillColor = &HFFFF00
    Shape65.FillColor = &HFFFF00
    Shape66.FillColor = &HFFFF00
    Shape67.FillColor = &HFFFF00
    Shape68.FillColor = &HFFFF00
    
    Shape69.FillColor = &H808000
    Shape70.FillColor = &H808000
    Shape71.FillColor = &H808000
    Shape72.FillColor = &H808000
    Shape73.FillColor = &H808000
    Shape74.FillColor = &H808000
    Shape75.FillColor = &H808000
    Shape76.FillColor = &H808000
    Shape77.FillColor = &H808000
    Shape78.FillColor = &H808000
End Sub

Private Sub Label2_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Label2.Enabled = False
    Timer5.Enabled = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = vbBlack
    Label3.ForeColor = &HC0C0FF
    Label4.ForeColor = &HC0C0FF
    Label11.ForeColor = &HC0C0FF

    Shape2.FillColor = &HFFFFC0
    Shape2.BorderWidth = 12

    Shape3.FillColor = &H40
    Shape4.FillColor = &H40

    Shape3.BorderWidth = 6
    Shape4.BorderWidth = 6

    Label7.Caption = "Use this Button to TRANSACT Customers PURCHASE"
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    
    Shape39.FillColor = &HFFFF00     'Lights of description
    Shape40.FillColor = &HFFFF00
    Shape41.FillColor = &HFFFF00
    Shape42.FillColor = &HFFFF00
    Shape43.FillColor = &HFFFF00
    Shape44.FillColor = &HFFFF00
    Shape45.FillColor = &HFFFF00
    Shape46.FillColor = &HFFFF00
    Shape47.FillColor = &HFFFF00
    Shape48.FillColor = &HFFFF00
    
    Shape49.FillColor = &H808000
    Shape50.FillColor = &H808000
    Shape51.FillColor = &H808000
    Shape52.FillColor = &H808000
    Shape53.FillColor = &H808000
    Shape54.FillColor = &H808000
    Shape55.FillColor = &H808000
    Shape56.FillColor = &H808000
    Shape57.FillColor = &H808000
    Shape58.FillColor = &H808000
    
    Shape59.FillColor = &H808000
    Shape60.FillColor = &H808000
    Shape61.FillColor = &H808000
    Shape62.FillColor = &H808000
    Shape63.FillColor = &H808000
    Shape64.FillColor = &H808000
    Shape65.FillColor = &H808000
    Shape66.FillColor = &H808000
    Shape67.FillColor = &H808000
    Shape68.FillColor = &H808000
    
    Shape69.FillColor = &H808000
    Shape70.FillColor = &H808000
    Shape71.FillColor = &H808000
    Shape72.FillColor = &H808000
    Shape73.FillColor = &H808000
    Shape74.FillColor = &H808000
    Shape75.FillColor = &H808000
    Shape76.FillColor = &H808000
    Shape77.FillColor = &H808000
    Shape78.FillColor = &H808000
    
    
End Sub

Private Sub Label3_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
     Label3.Enabled = False
     Timer5.Enabled = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = vbBlack
    Label2.ForeColor = &HC0C0FF
    Label4.ForeColor = &HC0C0FF
    Label11.ForeColor = &HC0C0FF

    Shape3.FillColor = &HFFFFC0
    Shape3.BorderWidth = 12

    Shape2.FillColor = &H40
    Shape32.FillColor = &H40
    
    Shape32.BorderWidth = 6
    Shape2.BorderWidth = 6

    Label8.Caption = "Use this Button to Add and Modify STOCKS"
    Label7.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    
    Shape39.FillColor = &H808000 'Lights of description
    Shape40.FillColor = &H808000
    Shape41.FillColor = &H808000
    Shape42.FillColor = &H808000
    Shape43.FillColor = &H808000
    Shape44.FillColor = &H808000
    Shape45.FillColor = &H808000
    Shape46.FillColor = &H808000
    Shape47.FillColor = &H808000
    Shape48.FillColor = &H808000
    
    Shape49.FillColor = &HFFFF00
    Shape50.FillColor = &HFFFF00
    Shape51.FillColor = &HFFFF00
    Shape52.FillColor = &HFFFF00
    Shape53.FillColor = &HFFFF00
    Shape54.FillColor = &HFFFF00
    Shape55.FillColor = &HFFFF00
    Shape56.FillColor = &HFFFF00
    Shape57.FillColor = &HFFFF00
    Shape58.FillColor = &HFFFF00
    
    Shape59.FillColor = &H808000
    Shape60.FillColor = &H808000
    Shape61.FillColor = &H808000
    Shape62.FillColor = &H808000
    Shape63.FillColor = &H808000
    Shape64.FillColor = &H808000
    Shape65.FillColor = &H808000
    Shape66.FillColor = &H808000
    Shape67.FillColor = &H808000
    Shape68.FillColor = &H808000
    
    Shape69.FillColor = &H808000
    Shape70.FillColor = &H808000
    Shape71.FillColor = &H808000
    Shape72.FillColor = &H808000
    Shape73.FillColor = &H808000
    Shape74.FillColor = &H808000
    Shape75.FillColor = &H808000
    Shape76.FillColor = &H808000
    Shape77.FillColor = &H808000
    Shape78.FillColor = &H808000
End Sub

Private Sub Label4_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Label4.Enabled = False
    Timer5.Enabled = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = vbBlack
    Label2.ForeColor = &HC0C0FF
    Label3.ForeColor = &HC0C0FF
    Label11.ForeColor = &HC0C0FF

    Shape4.FillColor = &HFFFFC0
    Shape4.BorderWidth = 12
    
    Shape32.FillColor = &H40
    Shape2.FillColor = &H40
    Shape3.FillColor = &H40

    Shape2.BorderWidth = 6
    Shape3.BorderWidth = 6
    Shape32.BorderWidth = 6

    Label9.Caption = "Use this Button to view and Generate REPORTS"
    Label7.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    
    Shape39.FillColor = &H808000 'Lights of description
    Shape40.FillColor = &H808000
    Shape41.FillColor = &H808000
    Shape42.FillColor = &H808000
    Shape43.FillColor = &H808000
    Shape44.FillColor = &H808000
    Shape45.FillColor = &H808000
    Shape46.FillColor = &H808000
    Shape47.FillColor = &H808000
    Shape48.FillColor = &H808000
    
    Shape49.FillColor = &H808000
    Shape50.FillColor = &H808000
    Shape51.FillColor = &H808000
    Shape52.FillColor = &H808000
    Shape53.FillColor = &H808000
    Shape54.FillColor = &H808000
    Shape55.FillColor = &H808000
    Shape56.FillColor = &H808000
    Shape57.FillColor = &H808000
    Shape58.FillColor = &H808000
    
    Shape59.FillColor = &H808000
    Shape60.FillColor = &H808000
    Shape61.FillColor = &H808000
    Shape62.FillColor = &H808000
    Shape63.FillColor = &H808000
    Shape64.FillColor = &H808000
    Shape65.FillColor = &H808000
    Shape66.FillColor = &H808000
    Shape67.FillColor = &H808000
    Shape68.FillColor = &H808000
    
    Shape69.FillColor = &HFFFF00
    Shape70.FillColor = &HFFFF00
    Shape71.FillColor = &HFFFF00
    Shape72.FillColor = &HFFFF00
    Shape73.FillColor = &HFFFF00
    Shape74.FillColor = &HFFFF00
    Shape75.FillColor = &HFFFF00
    Shape76.FillColor = &HFFFF00
    Shape77.FillColor = &HFFFF00
    Shape78.FillColor = &HFFFF00
End Sub

Private Sub Label5_Change()
Picture1.Cls
    For Intcount = 1 To 250
        Picture1.ForeColor = RGB(Intcount + 10, Intcount + 2, incount + 6)
            Picture1.CurrentX = Intcount
            Picture1.CurrentY = Intcount
        Picture1.Print Label5.Caption
    Next
End Sub

Private Sub Timer1_Timer()
    Label6.Visible = Label6.Visible Xor True
End Sub

Private Sub Timer2_Timer()
If Shape17.FillColor = &HC0C0FF Then
    Shape17.FillColor = &HFFFF&
    Shape19.FillColor = &HC0C0FF
ElseIf Shape19.FillColor = &HC0C0FF Then
    Shape19.FillColor = &HFFFF&
    Shape20.FillColor = &HC0C0FF
ElseIf Shape20.FillColor = &HC0C0FF Then
    Shape20.FillColor = &HFFFF&
    Shape23.FillColor = &HC0C0FF
ElseIf Shape23.FillColor = &HC0C0FF Then
    Shape23.FillColor = &HFFFF&
    Shape5.FillColor = &HC0C0FF
ElseIf Shape5.FillColor = &HC0C0FF Then
    Shape5.FillColor = &HFFFF&
    Shape34.FillColor = &HC0C0FF
ElseIf Shape34.FillColor = &HC0C0FF Then
    Shape34.FillColor = &HFFFF&
    Shape6.FillColor = &HC0C0FF
ElseIf Shape6.FillColor = &HC0C0FF Then
    Shape6.FillColor = &HFFFF&
    Shape24.FillColor = &HC0C0FF
ElseIf Shape24.FillColor = &HC0C0FF Then
    Shape24.FillColor = &HFFFF&
    Shape21.FillColor = &HC0C0FF
ElseIf Shape21.FillColor = &HC0C0FF Then
    Shape21.FillColor = &HFFFF&
    Shape22.FillColor = &HC0C0FF
ElseIf Shape22.FillColor = &HC0C0FF Then
    Shape22.FillColor = &HFFFF&
    Shape18.FillColor = &HC0C0FF
ElseIf Shape18.FillColor = &HC0C0FF Then
    Shape18.FillColor = &HFFFF&
Else
    Timer3.Interval = 50
    Timer2.Interval = 0
    
End If
    
End Sub

Private Sub Timer3_Timer()
If Shape17.FillColor = &HFFFF& Then
    Shape17.FillColor = &HC0C0FF
    Shape19.FillColor = &HFFFF&
ElseIf Shape19.FillColor = &HFFFF& Then
    Shape19.FillColor = &HC0C0FF
    Shape20.FillColor = &HFFFF&
ElseIf Shape20.FillColor = &HFFFF& Then
    Shape20.FillColor = &HC0C0FF
    Shape23.FillColor = &HFFFF&
ElseIf Shape23.FillColor = &HFFFF& Then
    Shape23.FillColor = &HC0C0FF
     Shape5.FillColor = &HFFFF&
ElseIf Shape5.FillColor = &HFFFF& Then
    Shape5.FillColor = &HC0C0FF
    Shape34.FillColor = &HFFFF&
ElseIf Shape34.FillColor = &HFFFF& Then
    Shape34.FillColor = &HC0C0FF
    Shape6.FillColor = &HFFFF&
ElseIf Shape6.FillColor = &HFFFF& Then
    Shape6.FillColor = &HC0C0FF
    Shape24.FillColor = &HFFFF&
ElseIf Shape24.FillColor = &HFFFF& Then
    Shape24.FillColor = &HC0C0FF
    Shape21.FillColor = &HFFFF&
ElseIf Shape21.FillColor = &HFFFF& Then
    Shape21.FillColor = &HC0C0FF
    Shape22.FillColor = &HFFFF&
ElseIf Shape22.FillColor = &HFFFF& Then
    Shape22.FillColor = &HC0C0FF
    Shape18.FillColor = &HFFFF&
ElseIf Shape18.FillColor = &HFFFF& Then
    Shape18.FillColor = &HC0C0FF
Else
        Timer2.Interval = 50
        Timer3.Interval = 0
End If
End Sub




Private Sub Form_Activate()
 Do While DoEvents()

   For Q = 0 To 80
           
    platos = buf(Q, 0)
    ph = buf(Q, 1)
    Y = buf(Q, 2)
    vel = buf(Q, 3)
    angle = buf(Q, 4)

    X = x0 + platos * Sin(ph * 7.28)
    Circle (X, Y), angle, vbDesktop
      
    Y = Y - vel
    ph = ph + rel / vel
    
    If Y < 0 Then
      buf(Q, 0) = Rnd(1) * 10000 + 5
      Y = ScaleHeight + 1000
      ph = Rnd(1) * 10000
      buf(Q, 3) = Rnd(1) * 10 + 1
      buf(Q, 4) = Rnd(1) * 300 + 1
    End If
             
    buf(Q, 1) = ph
    buf(Q, 2) = Y
    
    X = x0 + platos * Sin(ph * 7.28)
    Circle (X, Y), angle, vbWhite
      
    For T = 1 To 9: Next

xx:
    Next
   Loop

End Sub

Private Sub Form_Resize()

  x0 = ScaleWidth / 2    'position x  that bubbles appear
  rel = 0.0009999             'relation between bubble velocity and wave angle
 
Main.Cls
 
  For Q = 0 To 80
    buf(Q, 0) = Rnd(1) * 10000 + 5           'wave amplitude
    buf(Q, 1) = Rnd(1) * 1000               'wave angle
    buf(Q, 2) = ScaleHeight + 100          'current y position
    buf(Q, 3) = Rnd(1) * 10 + 1            'bubble velocity
    buf(Q, 4) = Rnd(1) * 300 + 1            'bubble radius
  Next

End Sub



Private Sub Timer4_Timer()
    Label5.Caption = "W E L C O M E !"
End Sub

Private Sub Timer5_Timer()
If Not Frame1.Height = 55 Then
    Frame1.Height = Frame1.Height - 900
    Frame1.Top = Frame1.Top + 900
Else
    If Label2.Enabled = False Then
        Timer5.Enabled = False
        Process.Timer8.Enabled = True
        Process.Show vbModal
        Label2.Enabled = True
    ElseIf Label3.Enabled = False Then
        Timer5.Enabled = False
        Stocks.Timer7.Enabled = True
        Stocks.Show vbModal
        Label3.Enabled = True
    ElseIf Label4.Enabled = False Then
        Timer5.Enabled = False
        Reports.Timer3.Enabled = True
        Reports.Show vbModal
        Label4.Enabled = True
    ElseIf Label11.Enabled = False Then
        Timer5.Enabled = False
        Suppliers.Timer4.Enabled = True
        Suppliers.Show vbModal
        Label11.Enabled = True
    Else
        Timer5.Enabled = False
        Label1.Enabled = True
        End
    End If
End If
End Sub

Private Sub Timer6_Timer()
If Not Frame1.Height = 9955 Then
    Frame1.Height = Frame1.Height + 900
    Frame1.Top = Frame1.Top - 900
Else
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer6.Enabled = False
End If
End Sub

