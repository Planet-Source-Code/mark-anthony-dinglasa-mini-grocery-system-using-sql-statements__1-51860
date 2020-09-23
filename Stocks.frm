VERSION 5.00
Begin VB.Form Stocks 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   DrawWidth       =   3
   FillColor       =   &H00FFFFC0&
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Stock 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9600
         Top             =   4440
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9600
         Top             =   4920
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   2400
         ScaleHeight     =   1095
         ScaleWidth      =   5415
         TabIndex        =   25
         Top             =   7080
         Width           =   5415
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   360
         Top             =   1680
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   6120
         Top             =   4680
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3600
         Top             =   1680
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   360
         Top             =   5160
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   1560
         TabIndex        =   1
         Top             =   2400
         Width           =   7095
         Begin VB.TextBox Text5 
            BackColor       =   &H00004080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00004080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00004080&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   405
            Left            =   5400
            TabIndex        =   23
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00004080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   600
            Width           =   4815
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Search Product ID Here !"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   855
            Left            =   3600
            TabIndex        =   10
            Top             =   1560
            Width           =   3255
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00004080&
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
               Left            =   120
               TabIndex        =   11
               Top             =   360
               Width           =   3015
            End
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00004080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00004080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00004080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier ID :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   300
            Left            =   3600
            TabIndex        =   22
            Top             =   120
            Width           =   1635
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   1560
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stocks :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   240
            TabIndex        =   9
            Top             =   2040
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Order Level :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   3600
            TabIndex        =   8
            Top             =   1080
            Width           =   1740
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Price Per Unit :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   240
            TabIndex        =   7
            Top             =   1560
            Width           =   1530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit In Stock :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Number:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Top             =   120
            Width           =   1710
         End
      End
      Begin VB.Shape Shape37 
         BorderColor     =   &H0080C0FF&
         BorderWidth     =   3
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   2520
         Top             =   7200
         Width           =   5415
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&DELETE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   480
         Left            =   4290
         TabIndex        =   24
         Top             =   6000
         Width           =   1725
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&EXIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   480
         Left            =   7905
         TabIndex        =   19
         Top             =   6000
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&CANCEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   480
         Left            =   450
         TabIndex        =   18
         Top             =   6000
         Width           =   1725
      End
      Begin VB.Shape Shape19 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programmed by: M.A.D  2004"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   4920
         Width           =   3015
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instick Mini-Grocery STOCKS !"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   4440
         TabIndex        =   16
         Top             =   2040
         Width           =   3990
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&UPDATE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   480
         Left            =   7965
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&SAVE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   480
         Left            =   5280
         TabIndex        =   14
         Top             =   720
         Width           =   2265
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&EDIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   480
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&ADD"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   480
         Left            =   255
         TabIndex        =   12
         Top             =   720
         Width           =   2145
      End
      Begin VB.Shape Shape14 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   8520
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   615
      End
      Begin VB.Shape Shape13 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   1080
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   615
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   5040
         Shape           =   3  'Circle
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1560
         Top             =   4680
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   3960
         Shape           =   3  'Circle
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4560
         Top             =   1920
         Width           =   4095
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   7800
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   5280
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00C0E0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   8040
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   1695
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00C0E0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   5520
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   1695
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0E0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   1695
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00C0E0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   480
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   1695
      End
      Begin VB.Shape Shape21 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2280
         Top             =   840
         Width           =   615
      End
      Begin VB.Shape Shape23 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7320
         Top             =   840
         Width           =   615
      End
      Begin VB.Shape Shape22 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4800
         Top             =   840
         Width           =   615
      End
      Begin VB.Shape Shape18 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   2175
         Left            =   8760
         Top             =   1320
         Width           =   255
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   2175
         Left            =   1200
         Top             =   1320
         Width           =   255
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H0080C0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   840
         Shape           =   3  'Circle
         Top             =   3480
         Width           =   855
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H0080C0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   7800
         Shape           =   3  'Circle
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Shape Shape20 
         BorderColor     =   &H00C0E0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   480
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Shape Shape25 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   7800
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Shape Shape26 
         BorderColor     =   &H00C0E0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   8040
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Shape Shape27 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   8760
         Top             =   4320
         Width           =   255
      End
      Begin VB.Shape Shape24 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   1200
         Top             =   4320
         Width           =   255
      End
      Begin VB.Shape Shape28 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H00C0E0FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Shape Shape33 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6240
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Shape Shape32 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2280
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Shape Shape30 
         BorderColor     =   &H00C0E0FF&
         BorderWidth     =   2
         FillColor       =   &H00404080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   5040
         Top             =   6720
         Width           =   255
      End
   End
End
Attribute VB_Name = "Stocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Con As New ADODB.Connection
Private Rs As New ADODB.Recordset
Private Sup As New ADODB.Recordset
Private WithEvents Istext As TextBox
Attribute Istext.VB_VarHelpID = -1
Private WithEvents Bantam As TextBox
Attribute Bantam.VB_VarHelpID = -1
Dim s As String, Z As Integer, P As Integer
Dim X(50), Y(50), Pace(50), Size(50) As Integer

Private Sub Combo1_Click()
Rs.Open "select * from Products where ProdName='" & Combo1.Text & "'", Con, 3, 3
If Not Rs.EOF Then
            Text1.Text = Rs("ProdNo")
            Combo2.Text = Rs("SuppID")
            Text6.Text = Rs("ProdName")
            Text2.Text = Rs("StockUnit")
            Text3.Text = Rs("PriceUnit")
            Text5.Text = Rs("ReOrderLevel")
            Text4.Text = Rs("Stocks")
        Else
            Exit Sub
        End If
Rs.Close
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
    s = "Instick Mini-Grocery STOCKS !"
    Stock.BackColor = vbDesktop
    Me.BackColor = vbDesktop
    Stock.Height = 40
    Stock.Move (Screen.Width - Stock.Width) / 2, (Screen.Height - Stock.Height) / 1.14
With Con
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Grocery.mdb"
End With
    DisplayDB
End Sub

Private Sub Form_Resize()
For P = 1 To 250
    Picture1.ForeColor = RGB(P + 100, P + 1, 0)
    Picture1.CurrentX = P
    Picture1.CurrentY = P
    Picture1.Print "    S T O C K S "
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Con.Close
End Sub

Private Sub Istext_LostFocus()
    Istext.ForeColor = &HC0E0FF
    Istext.BackColor = &H4080
End Sub

Private Sub Label10_Click()
With Con
    .BeginTrans
    .Execute "Update Products set SuppID='" & Combo2.Text & "',ProdName='" & Text6.Text & "',StockUnit='" & Text2.Text & "',PriceUnit=" & Text3.Text & ",ReOrderLevel=" & Text5.Text & ",Stocks=" & Text4.Text & " where ProdNo='" & Text1.Text & "'"
    .CommitTrans
End With
    MsgBox ("Records Updated !"), vbOKOnly + vbInformation, "SUCCESSFULLY !"
    LockText 1
    NoButtons 3
    DisplayDB
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape6.FillColor = &HFFFF00
    Label10.ForeColor = vbBlue

    Shape3.FillColor = &H404080
    Shape4.FillColor = &H404080
    Shape5.FillColor = &H404080

    Label7.ForeColor = &HC0E0FF
    Label8.ForeColor = &HC0E0FF
    Label9.ForeColor = &HC0E0FF
End Sub

Private Sub Label13_Click()
    NoButtons 3
    DisplayDB
    LockText 1
    Text2.SelStart = 0
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label13.ForeColor = vbBlue
    Shape19.FillColor = &HFFFF00
End Sub

Private Sub Label14_Click()
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = True
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label14.ForeColor = vbBlue
    Shape25.FillColor = &HFFFF00
End Sub

Private Sub Label17_Click()
Dim Response As Integer
Response = MsgBox("Do you really want to delete this Record?", vbYesNo + vbCritical, "Warning !")

    If Response = vbYes Then
        

    Else
        Exit Sub
    End If
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label17.ForeColor = vbBlue
    Shape28.FillColor = &HFFFF00
End Sub

Private Sub Label18_Click()
  
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
End Sub



Private Sub Label6_Change()


End Sub

Private Sub Label7_Click()
Dim I As Integer, L As String
    Combo2.Clear
    NoButtons 1
    LockText 2
    Text2.Text = ""
    Text3.Text = ""
    Text5.Text = ""
    Text4.Text = ""
    Text6.Text = ""
Rs.Open "Select ProdNo from Products", Con, 3, 2
    Rs.MoveLast
        L = Rs("ProdNo")
        L = Right(L, 1)
        I = L
            If I <> 0 Then
                I = I + 1
            Else
                I = 1
            End If
            
            Text1.Text = "Prod-" & Format(I, "000")
            Text6.SetFocus
Rs.Close
Sup.Open "Select SuppID from Suppliers", Con, 3, 2
    Do While Not Sup.EOF
        Combo2.AddItem Sup("SuppId")
        Sup.MoveNext
    Loop
Sup.Close
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape3.FillColor = &HFFFF00
    Label7.ForeColor = vbBlue


    Shape4.FillColor = &H404080
    Shape5.FillColor = &H404080
    Shape6.FillColor = &H404080


    Label8.ForeColor = &HC0E0FF
    Label9.ForeColor = &HC0E0FF
    Label10.ForeColor = &HC0E0FF

End Sub

Private Sub Label8_Click()
    NoButtons 2
    LockText 2
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.Text)
    Text6.SetFocus
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape4.FillColor = &HFFFF00
    Label8.ForeColor = vbBlue

    Shape3.FillColor = &H404080
    Shape5.FillColor = &H404080
    Shape6.FillColor = &H404080

    Label7.ForeColor = &HC0E0FF
    Label9.ForeColor = &HC0E0FF
    Label10.ForeColor = &HC0E0FF
End Sub

Private Sub Label9_Click()
    If Text4.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text2.Text = "" Or Text6.Text = "" Then
            MsgBox ("Fill in Form Correctly !"), vbOKOnly + vbCritical, "Cannot Be Saved !"
    Else
       With Con
            .BeginTrans
            .Execute "Insert into Products(ProdNo,SuppID,ProdName,StockUnit,PriceUnit,ReOrderLevel,Stocks) values ('" & Text1.Text & "','" & Combo2.Text & "','" & Text6.Text & "','" & Text2.Text & "'," & Text3.Text & "," & Text5.Text & "," & Text4.Text & ");"
            .CommitTrans
        End With
            
            MsgBox ("Records Save !"), vbOKOnly + vbInformation, "SUCCESSFUL !"
            NoButtons 3
            LockText 1
            DisplayDB
      
    End If
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape5.FillColor = &HFFFF00
    Label9.ForeColor = vbBlue

    Shape3.FillColor = &H404080
    Shape4.FillColor = &H404080
    Shape6.FillColor = &H404080

    Label7.ForeColor = &HC0E0FF
    Label8.ForeColor = &HC0E0FF
    Label10.ForeColor = &HC0E0FF
End Sub

Private Sub Stock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Shape3.FillColor = &H404080
    Shape4.FillColor = &H404080
    Shape5.FillColor = &H404080
    Shape6.FillColor = &H404080
    Shape19.FillColor = &H404080
    Shape25.FillColor = &H404080
    Shape28.FillColor = &H404080

    

    Label7.ForeColor = &HC0E0FF
    Label8.ForeColor = &HC0E0FF
    Label9.ForeColor = &HC0E0FF
    Label10.ForeColor = &HC0E0FF
    Label13.ForeColor = &HC0E0FF
    Label14.ForeColor = &HC0E0FF
    Label17.ForeColor = &HC0E0FF
 
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    'Case Asc(vbCr)
     '   KeyAscii = 0
    Case 13
        Text3.Text = Format(Text3.Text, "#,###.00")
        Text4.SetFocus
    Case 8, 46
    Case 47 To 57
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub Text4_Change()
Dim message As String
On Error Resume Next
message = "Minimum Level Reach !" + vbCr
message = message & " Need to Order " & Text6.Text
If Val(Text5.Text) > Val(Text4.Text) Then
    MsgBox message, vbOKOnly + vbExclamation, "WARNING !"
Else
    Exit Sub
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
    Text5.SetFocus
    Case 8, 46
    Case 48 To 57
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc(vbCr)
        KeyAscii = 0
    Case 8, 46
    Case 48 To 57
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
For I = 0 To 50
    Circle (X(I), Y(I)), Size(I), BackColor
    Y(I) = Y(I) + Pace(I)
    If Y(I) >= Me.Height Then Y(I) = 0: X(I) = Rnd * Me.Width
    Circle (X(I), Y(I)), Size(I)
Next
End Sub

Private Sub Timer2_Timer()
Label11.Caption = Left(s, Z)
    Z = Z + 1
If Z > Len(s) Then
    Z = 0
        If Label11.Alignment = 0 Then
            Label11.Alignment = 1
        ElseIf Label11.Alignment = 1 Then
            Label11.Alignment = 2
        Else
            Label11.Alignment = 0
        End If
End If
End Sub

Private Sub Timer3_Timer()
    Label12.Visible = Label12.Visible Xor True
End Sub

Public Sub NoButtons(n As Integer)
If n = 1 Then
    Label7.Enabled = False
    Label8.Enabled = False
    Label14.Enabled = False
    Label10.Enabled = False
    Label17.Enabled = False
    
    Label9.Enabled = True
    Label13.Enabled = True
    
    Combo1.Enabled = False
    Combo2.Enabled = True
ElseIf n = 2 Then
    Label7.Enabled = False
    Label8.Enabled = False
    Label14.Enabled = False
    Label9.Enabled = False
    Label17.Enabled = False
    
    Label10.Enabled = True
    Label13.Enabled = True
    Combo1.Enabled = False
    Combo2.Enabled = True
ElseIf n = 3 Then
    Label9.Enabled = False
    Label10.Enabled = False
    Label13.Enabled = False
    Label17.Enabled = True
    
    Label7.Enabled = True
    Label8.Enabled = True
    Label14.Enabled = True
    Combo1.Enabled = True
    Combo2.Enabled = False
End If

End Sub



Public Sub LockText(n As Integer)
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
End Sub

Private Sub Timer4_Timer()
Dim ret As Integer
If TypeOf Me.ActiveControl Is TextBox Then
    Set Istext = Me.ActiveControl
    Set Bantam = Me.ActiveControl
        Istext.BackColor = &H80C0FF
        Istext.ForeColor = vbBlue
        ret = InStr(1, Bantam.Text, "'", vbTextCompare)
    If ret <> 0 Then
        Bantam.Text = Replace(Bantam.Text, "'", "")
        Bantam.SelStart = Len(Bantam.Text)
    End If
End If
End Sub

Private Sub Timer5_Timer()

End Sub

Private Sub Timer6_Timer()
If Not Stock.Height = 40 Then
    Stock.Height = Stock.Height - 800
    Stock.Top = Stock.Top + 800
Else
   
        Timer2.Enabled = True
        Timer3.Enabled = True
        Timer4.Enabled = True
        Label4.Enabled = True
        Timer6.Enabled = False
        Unload Me
        Main.Timer6.Enabled = True

End If
End Sub

Private Sub Timer7_Timer()
If Not Stock.Height = 8840 Then
    Stock.Height = Stock.Height + 800
    Stock.Top = Stock.Top - 800
Else
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer7.Enabled = False
End If
End Sub

Function DisplayDB()
Combo1.Clear
Rs.Open "Select * from Products", Con, 3, 2
    Rs.MoveFirst
        If Not Rs.EOF Then
            Text1.Text = Rs("ProdNo")
            Combo2.Text = Rs("SuppID")
            Text6.Text = Rs("ProdName")
            Text2.Text = Rs("StockUnit")
            Text3.Text = Rs("PriceUnit")
            Text5.Text = Rs("ReOrderLevel")
            Text4.Text = Rs("Stocks")
        Else
            Exit Function
        End If
        
        Do While Not Rs.EOF
            Combo1.AddItem Rs("ProdName")
            Rs.MoveNext
        Loop
Sup.Open "Select SuppID from Suppliers", Con, 3, 2
    Do While Not Sup.EOF
        Combo2.AddItem Sup("SuppId")
        Sup.MoveNext
    Loop
Sup.Close
Rs.Close
End Function
