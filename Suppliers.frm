VERSION 5.00
Begin VB.Form Suppliers 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11025
   DrawWidth       =   3
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3120
         ScaleHeight     =   975
         ScaleWidth      =   6255
         TabIndex        =   28
         Top             =   6720
         Width           =   6255
      End
      Begin VB.Timer Timer6 
         Interval        =   100
         Left            =   1560
         Top             =   6360
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8880
         Top             =   5520
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8880
         Top             =   6000
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   5520
         Top             =   2160
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5280
         Top             =   5880
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8880
         Top             =   7200
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   1560
         TabIndex        =   1
         Top             =   2040
         Width           =   7095
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C000&
            Caption         =   "Search Supplier ID Here !"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   3960
            TabIndex        =   22
            Top             =   2520
            Width           =   3015
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00404000&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   345
               Left            =   240
               TabIndex        =   23
               Top             =   480
               Width           =   2535
            End
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   720
            Width           =   5055
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1320
            Width           =   5055
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1920
            Width           =   5055
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   2520
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Name:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier ID:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone Number:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   2640
            Width           =   1740
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax Number:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   3240
            Width           =   1455
         End
      End
      Begin VB.Shape Shape28 
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   840
         Shape           =   3  'Circle
         Top             =   7080
         Width           =   255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   4680
         TabIndex        =   27
         Top             =   5880
         Width           =   585
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   3960
         TabIndex        =   26
         Top             =   5880
         Width           =   660
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   4680
         TabIndex        =   25
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   3960
         TabIndex        =   24
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   600
         Left            =   3120
         TabIndex        =   21
         Top             =   7080
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape27 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   6
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   3000
         Top             =   6600
         Width           =   6495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   555
         Left            =   9240
         TabIndex        =   20
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   8760
         TabIndex        =   19
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UPDATE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   7080
         TabIndex        =   18
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAVE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   5640
         TabIndex        =   17
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANCEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   3720
         TabIndex        =   16
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EDIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8760
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   5400
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3720
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2040
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   1935
         Left            =   7680
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   1935
         Left            =   1320
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   4455
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   9000
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   7320
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   2280
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape19 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   8880
         Shape           =   3  'Circle
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Shape Shape18 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   9240
         Top             =   1080
         Width           =   255
      End
      Begin VB.Shape Shape24 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   8280
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape23 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6600
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape22 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4920
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape21 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3240
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape20 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1560
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape25 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   6255
         Left            =   840
         Top             =   1080
         Width           =   255
      End
      Begin VB.Shape Shape26 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   840
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   3240
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   3240
         Shape           =   2  'Oval
         Top             =   5040
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Suppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Con As New ADODB.Connection
Private Rs As New ADODB.Recordset
Private WithEvents Istext As TextBox
Attribute Istext.VB_VarHelpID = -1
Private WithEvents Bantam As TextBox
Attribute Bantam.VB_VarHelpID = -1
Dim X(50), Y(50), Pace(50), Size(50) As Integer

Private Sub Combo1_Click()
On Error Resume Next
Rs.Open "Select * from Suppliers where CompName='" & Combo1.Text & "'", Con, 3, 3
     If Not Rs.EOF Then
        Text1.Text = Rs("SuppID")
        Text6.Text = Rs("CompName")
        Text2.Text = Rs("ContactName")
        Text3.Text = Rs("Address")
        Text4.Text = Rs("PhoneNo")
        Text5.Text = Rs("FaxNo")
    Else
        Exit Sub
    End If
Rs.Close
End Sub

Private Sub Form_Activate()
For I = 0 To 50
    X1 = Rnd * Me.Width
    Y1 = Rnd * Me.Height
    
    pace1 = 500 - Rnd * 499
    size1 = Rnd * 16
    
    X(I) = X1: Y(I) = Y1
    Pace(I) = pace1: Size(I) = size1
Next

End Sub

Private Sub Form_Load()
    Label16.Caption = Date
    Frame1.Height = 95
    Frame1.Move (Screen.Width - Frame1.Width) / 2, (Screen.Height - Frame1.Height) / 1.17
    Frame1.BackColor = vbDesktop
    Me.BackColor = vbDesktop
With Con
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Grocery.mdb"
End With
    DisplayDB
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label7.ForeColor = vbYellow
    Shape8.FillColor = &H4000

    Shape9.FillColor = &H4000
    Label8.ForeColor = vbYellow

    Label9.ForeColor = vbYellow
    Shape10.FillColor = &H4000

    Shape11.FillColor = &H4000
    Label10.ForeColor = vbYellow

    Shape12.FillColor = &H4000
    Label11.ForeColor = vbYellow

    Shape13.FillColor = &H4000
    Label12.ForeColor = vbYellow

    Label13.ForeColor = vbYellow
    Shape19.FillColor = &H4000
End Sub

Private Sub Istext_LostFocus()
    Istext.BackColor = &H404000
    Istext.ForeColor = &HFFFFC0
End Sub

Private Sub Label10_Click()
If Text6.Text = "" Or Text3.Text = "" Or Text2.Text = "" Then
    MsgBox ("Please Fill in Form Correctly !"), vbOKOnly + vbCritical, "Cannot Be Saved !"
Else
With Con
    .BeginTrans
    .Execute "Insert into Suppliers (SuppID,CompName,ContactName,Address,PhoneNo,FaxNo) values ('" & Text1.Text & "','" & Text6.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "');"
    .CommitTrans
End With
    MsgBox ("Record is Saved !"), vbOKOnly + vbInformation, " Successful !"
    NoButtons 3
    LockText 1
    DisplayDB
End If
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape11.FillColor = &HFFFFC0
    Label10.ForeColor = vbBlue
End Sub

Private Sub Label11_Click()
With Con
    .BeginTrans
    .Execute "update Suppliers set CompName='" & Text6.Text & "',ContactName='" & Text2.Text & "',Address='" & Text3.Text & "',PhoneNo='" & Text4.Text & "',FaxNo='" & Text5.Text & "' where SuppID='" & Text1.Text & "'"
    .CommitTrans
End With
    MsgBox ("Record was update !"), vbOKOnly + vbInformation, "Successful !"
    NoButtons 3
    LockText 1
    DisplayDB
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape12.FillColor = &HFFFFC0
    Label11.ForeColor = vbBlue
End Sub

Private Sub Label12_Click()
Dim Response As Integer
Response = MsgBox("Are you sure you want to delete this Record?", vbYesNo + vbCritical, "Warning!")
If Response = vbYes Then
With Con
    .BeginTrans
    .Execute "delete * from Suppliers where SuppID='" & Text1.Text & "'"
    .CommitTrans
End With
    MsgBox ("Record was deleted !"), vbOKOnly + vbInformation, "Successful !"
    DisplayDB
Else
    Exit Sub
End If
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape13.FillColor = &HFFFFC0
    Label12.ForeColor = vbBlue
End Sub

Private Sub Label13_Click()
   Timer5.Enabled = True
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label13.ForeColor = vbBlue
    Shape19.FillColor = &HFFFFC0
End Sub

Private Sub Label14_Change()
For Intcount = 1 To 250
    Picture1.ForeColor = RGB(0, Intcount, 0)
    Picture1.CurrentX = Intcount
    Picture1.CurrentY = Intcount
    Picture1.Print "SUPPLIERS RECORD !"
Next
End Sub

Private Sub Label7_Click()
Dim M As Integer, Y As String
    NoButtons 1
    LockText 2
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""

Rs.Open "Select * from Suppliers", Con, 3, 2

    Rs.MoveLast
    
        Y = Rs("SuppID")
        Y = Right(Y, 1)
        M = Y
        
        If M <> 0 Then
            M = M + 1
        Else
            M = 1
        End If
        
        Text1.Text = "Sup-" & Format(M, "000")
        Text6.SetFocus
Rs.Close
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label7.ForeColor = vbBlue
    Shape8.FillColor = &HFFFFC0
End Sub

Private Sub Label8_Click()
    NoButtons 2
    LockText 2
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.Text)
    Text6.SetFocus
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape9.FillColor = &HFFFFC0
    Label8.ForeColor = vbBlue
End Sub

Private Sub Label9_Click()
    NoButtons 3
    LockText 1
    DisplayDB
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape10.FillColor = &HFFFFC0
    Label9.ForeColor = vbBlue
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
    Label14.Caption = "SUPPLIERS RECORD !"
End Sub

Public Sub DisplayDB()
    Combo1.Clear
Rs.Open "Select * from Suppliers", Con, 3, 2
    If Not Rs.EOF Then
        Text1.Text = Rs("SuppID")
        Text6.Text = Rs("CompName")
        Text2.Text = Rs("ContactName")
        Text3.Text = Rs("Address")
        Text4.Text = Rs("PhoneNo")
        Text5.Text = Rs("FaxNo")
    Else
        Exit Sub
    End If
    
    Do While Not Rs.EOF
        Combo1.AddItem Rs("CompName")
        Rs.MoveNext
    Loop
Rs.Close
End Sub

Public Sub NoButtons(n As Integer)

If n = 1 Then
    Label7.Enabled = False
    Label8.Enabled = False
    Label11.Enabled = False
    Label12.Enabled = False
    
    Label9.Enabled = True
    Label10.Enabled = True
    Combo1.Enabled = False
ElseIf n = 2 Then
    Label7.Enabled = False
    Label8.Enabled = False
    Label10.Enabled = False
    Label12.Enabled = False
    
    Label9.Enabled = True
    Label11.Enabled = True
    Combo1.Enabled = False
ElseIf n = 3 Then
    Label7.Enabled = True
    Label8.Enabled = True
    Label12.Enabled = True
    
    Label9.Enabled = False
    Label10.Enabled = False
    Label11.Enabled = False
    Combo1.Enabled = True
End If
End Sub

Function LockText(n As Integer)
If n = 1 Then
    Text2.Locked = True
    Text3.Locked = True
    Text4.Locked = True
    Text5.Locked = True
    Text6.Locked = True
ElseIf n = 2 Then
    Text2.Locked = False
    Text3.Locked = False
    Text4.Locked = False
    Text5.Locked = False
    Text6.Locked = False
End If
End Function

Private Sub Timer2_Timer()
    Label18.Caption = Time
End Sub

Private Sub Timer3_Timer()
Dim ret As Integer
If TypeOf Me.ActiveControl Is TextBox Then
    Set Istext = Me.ActiveControl
    Set Bantam = Me.ActiveControl
    Istext.ForeColor = vbBlue
    Istext.BackColor = &HC0C000
    ret = InStr(1, Bantam.Text, "'", vbTextCompare)
    If ret <> 0 Then
        Bantam.Text = Replace(Bantam.Text, "'", "", , , vbTextCompare)
        Bantam.SelStart = Len(Bantam.Text)
    End If
End If
End Sub


Private Sub Timer4_Timer()
If Not Frame1.Height = 8295 Then
    Frame1.Height = Frame1.Height + 410
    Frame1.Top = Frame1.Top - 410
Else
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
If Not Frame1.Height = 95 Then
    Frame1.Height = Frame1.Height - 410
    Frame1.Top = Frame1.Top + 410
Else
     Timer1.Enabled = False
     Timer2.Enabled = False
     Timer3.Enabled = False
      Timer5.Enabled = False
     Con.Close
     Unload Me
     Main.Timer6.Enabled = True

End If
End Sub

Private Sub Timer6_Timer()
For I = 0 To 50
    Circle (X(I), Y(I)), Size(I), BackColor
    Y(I) = Y(I) + Pace(I)
    If Y(I) >= Me.Height Then Y(I) = 0: X(I) = Rnd * Me.Width
    Circle (X(I), Y(I)), Size(I)
Next
End Sub
