VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Process 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   2010
   ClientTop       =   555
   ClientWidth     =   11970
   DrawWidth       =   3
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9690
      Left            =   120
      TabIndex        =   0
      Top             =   -240
      Width           =   11775
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   6495
         Left            =   1440
         TabIndex        =   48
         Top             =   1320
         Visible         =   0   'False
         Width           =   8895
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFC0FF&
            Height          =   5055
            Left            =   720
            ScaleHeight     =   4995
            ScaleWidth      =   7275
            TabIndex        =   49
            Top             =   720
            Width           =   7335
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00800080&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   65
               Top             =   2280
               Width           =   4815
            End
            Begin VB.TextBox Text16 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   5520
               Locked          =   -1  'True
               TabIndex        =   64
               Top             =   1560
               Width           =   1455
            End
            Begin VB.TextBox Text19 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   5520
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   3000
               Width           =   1455
            End
            Begin VB.TextBox Text18 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   3000
               Width           =   1455
            End
            Begin VB.PictureBox Picture5 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               Height          =   1095
               Left            =   0
               ScaleHeight     =   1095
               ScaleWidth      =   7575
               TabIndex        =   52
               Top             =   0
               Width           =   7575
               Begin VB.Shape Shape57 
                  BorderStyle     =   0  'Transparent
                  FillColor       =   &H00FFC0FF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Left            =   4680
                  Shape           =   3  'Circle
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Shape Shape56 
                  BorderStyle     =   0  'Transparent
                  FillColor       =   &H00FFC0FF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Left            =   120
                  Shape           =   3  'Circle
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   " D E L I V E R I E S"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   24
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0FF&
                  Height          =   540
                  Left            =   480
                  TabIndex        =   53
                  Top             =   120
                  Width           =   4095
               End
               Begin VB.Line Line2 
                  BorderColor     =   &H00FFC0FF&
                  BorderWidth     =   3
                  X1              =   0
                  X2              =   5400
                  Y1              =   600
                  Y2              =   600
               End
            End
            Begin VB.TextBox Text15 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Shape Shape52 
               BorderColor     =   &H00000000&
               Height          =   735
               Left            =   2640
               Shape           =   4  'Rounded Rectangle
               Top             =   3840
               Width           =   2055
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UPDATE"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   465
               Left            =   2880
               TabIndex        =   61
               Top             =   3960
               Width           =   1575
            End
            Begin VB.Shape Shape50 
               BorderColor     =   &H00000000&
               Height          =   735
               Left            =   240
               Shape           =   4  'Rounded Rectangle
               Top             =   3840
               Width           =   2055
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFF00&
               BackStyle       =   0  'Transparent
               Caption         =   "ADD"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   465
               Left            =   840
               TabIndex        =   60
               Top             =   3960
               Width           =   855
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   3840
               TabIndex        =   59
               Top             =   3120
               Width           =   975
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Product Name :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   240
               TabIndex        =   57
               Top             =   2400
               Width           =   1605
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delivered Date :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   240
               TabIndex        =   55
               Top             =   3120
               Width           =   1710
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Product Number :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   3720
               TabIndex        =   54
               Top             =   1680
               Width           =   1800
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delivery Number :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   240
               TabIndex        =   51
               Top             =   1680
               Width           =   1875
            End
            Begin VB.Shape Shape51 
               BorderColor     =   &H00FF00FF&
               BorderWidth     =   2
               FillColor       =   &H00800080&
               FillStyle       =   0  'Solid
               Height          =   735
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   3840
               Width           =   2295
            End
            Begin VB.Shape Shape53 
               BorderColor     =   &H00FF00FF&
               BorderWidth     =   2
               FillColor       =   &H00800080&
               FillStyle       =   0  'Solid
               Height          =   735
               Left            =   2520
               Shape           =   4  'Rounded Rectangle
               Top             =   3840
               Width           =   2295
            End
            Begin VB.Shape Shape54 
               BorderColor     =   &H00000000&
               Height          =   735
               Left            =   5040
               Shape           =   4  'Rounded Rectangle
               Top             =   3840
               Width           =   2055
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EXIT"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   465
               Left            =   5640
               TabIndex        =   62
               Top             =   3960
               Width           =   900
            End
            Begin VB.Shape Shape55 
               BorderColor     =   &H00FF00FF&
               BorderWidth     =   2
               FillColor       =   &H00800080&
               FillStyle       =   0  'Solid
               Height          =   735
               Left            =   4920
               Shape           =   4  'Rounded Rectangle
               Top             =   3840
               Width           =   2295
            End
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   555
            Left            =   8280
            TabIndex        =   66
            Top             =   360
            Width           =   315
         End
         Begin VB.Image Image2 
            Height          =   6375
            Left            =   0
            Picture         =   "Process.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   8895
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   6735
         Left            =   1320
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   9255
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFC0FF&
            Height          =   5175
            Left            =   840
            ScaleHeight     =   5115
            ScaleWidth      =   7515
            TabIndex        =   33
            Top             =   840
            Width           =   7575
            Begin VB.PictureBox Picture3 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               Height          =   1095
               Left            =   0
               ScaleHeight     =   1095
               ScaleWidth      =   7575
               TabIndex        =   44
               Top             =   0
               Width           =   7575
               Begin VB.Shape Shape59 
                  BorderStyle     =   0  'Transparent
                  FillColor       =   &H00FFC0FF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Left            =   4560
                  Shape           =   3  'Circle
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Shape Shape58 
                  BorderStyle     =   0  'Transparent
                  FillColor       =   &H00FFC0FF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Left            =   120
                  Shape           =   3  'Circle
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00FFC0FF&
                  BorderWidth     =   3
                  X1              =   0
                  X2              =   5400
                  Y1              =   600
                  Y2              =   600
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "P A Y M E N T S"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   24
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0FF&
                  Height          =   540
                  Left            =   720
                  TabIndex        =   47
                  Top             =   120
                  Width           =   3540
               End
            End
            Begin VB.TextBox Text14 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   1920
               TabIndex        =   42
               Top             =   3000
               Width           =   3855
            End
            Begin VB.TextBox Text13 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   5760
               TabIndex        =   40
               Top             =   2280
               Width           =   1575
            End
            Begin VB.TextBox Text12 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   1920
               TabIndex        =   38
               Top             =   2280
               Width           =   1575
            End
            Begin VB.TextBox Text11 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   5760
               TabIndex        =   36
               Top             =   1560
               Width           =   1575
            End
            Begin VB.TextBox Text10 
               BackColor       =   &H00800080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   420
               Left            =   1920
               TabIndex        =   34
               Top             =   1560
               Width           =   1575
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "P"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   5520
               TabIndex        =   69
               Top             =   2400
               Width           =   165
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "P"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   5520
               TabIndex        =   68
               Top             =   1680
               Width           =   165
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "P"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   1680
               TabIndex        =   67
               Top             =   2400
               Width           =   165
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "       EXIT"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   465
               Left            =   4440
               TabIndex        =   46
               Top             =   4080
               Width           =   2235
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFF00&
               BackStyle       =   0  'Transparent
               Caption         =   "       OK"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   465
               Left            =   960
               TabIndex        =   45
               Top             =   4080
               Width           =   2145
            End
            Begin VB.Shape Shape49 
               Height          =   735
               Left            =   4440
               Shape           =   4  'Rounded Rectangle
               Top             =   3960
               Width           =   2295
            End
            Begin VB.Shape Shape48 
               BorderColor     =   &H00000000&
               Height          =   735
               Left            =   840
               Shape           =   4  'Rounded Rectangle
               Top             =   3960
               Width           =   2295
            End
            Begin VB.Shape Shape47 
               BorderColor     =   &H00FF00FF&
               BorderWidth     =   2
               FillColor       =   &H00800080&
               FillStyle       =   0  'Solid
               Height          =   735
               Left            =   4320
               Shape           =   4  'Rounded Rectangle
               Top             =   3960
               Width           =   2535
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transact By:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   120
               TabIndex        =   43
               Top             =   3120
               Width           =   1320
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Change:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   3960
               TabIndex        =   41
               Top             =   2400
               Width           =   870
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amount Pay :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   120
               TabIndex        =   39
               Top             =   2400
               Width           =   1365
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grand Total"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   3960
               TabIndex        =   37
               Top             =   1680
               Width           =   1245
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Number :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400040&
               Height          =   240
               Left            =   120
               TabIndex        =   35
               Top             =   1680
               Width           =   1605
            End
            Begin VB.Shape Shape46 
               BorderColor     =   &H00FF00FF&
               BorderWidth     =   2
               FillColor       =   &H00800080&
               FillStyle       =   0  'Solid
               Height          =   735
               Left            =   720
               Shape           =   4  'Rounded Rectangle
               Top             =   3960
               Width           =   2535
            End
         End
         Begin VB.Image Image1 
            Height          =   6615
            Left            =   0
            Picture         =   "Process.frx":2312
            Stretch         =   -1  'True
            Top             =   120
            Width           =   9255
         End
      End
      Begin VB.Timer Timer9 
         Interval        =   100
         Left            =   960
         Top             =   4920
      End
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   10920
         Top             =   5760
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   10920
         Top             =   5160
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   1560
         Top             =   1560
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   360
         TabIndex        =   22
         Top             =   10560
         Visible         =   0   'False
         Width           =   10095
         Begin MSFlexGridLib.MSFlexGrid MS1 
            Height          =   2895
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            BackColor       =   -2147483639
            ForeColorFixed  =   16711680
            FormatString    =   $"Process.frx":4624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   8760
         Top             =   720
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   9720
         Top             =   8280
      End
      Begin MSFlexGridLib.MSFlexGrid MS2 
         Height          =   1935
         Left            =   2040
         TabIndex        =   25
         Top             =   6480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColor       =   8388736
         ForeColor       =   16622847
         BackColorFixed  =   16622847
         ForeColorFixed  =   8388736
         BackColorBkg    =   8388736
         Enabled         =   -1  'True
         FocusRect       =   2
         FormatString    =   "PRODUCT NO.|PRODUCT NAME                | PRICE           | QUANTITY|      TOTAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   2040
         TabIndex        =   1
         Top             =   2040
         Width           =   7695
         Begin VB.TextBox Text9 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   2640
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   250
            Left            =   5760
            Top             =   720
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   2640
            Width           =   1815
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00800080&
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
            ForeColor       =   &H00FFC0FF&
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " &NEW"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   405
            Left            =   4920
            TabIndex        =   75
            Top             =   3360
            Width           =   1995
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   300
            Left            =   4440
            TabIndex        =   74
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   300
            Left            =   5640
            TabIndex        =   73
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "G-Total :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   4440
            TabIndex        =   72
            Top             =   2640
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   390
            Left            =   5400
            TabIndex        =   71
            Top             =   2640
            Width           =   225
         End
         Begin VB.Shape Shape44 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7080
            Shape           =   3  'Circle
            Top             =   3360
            Width           =   855
         End
         Begin VB.Shape Shape43 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   -120
            Shape           =   3  'Circle
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min Reorder Level :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Top             =   1680
            Width           =   2085
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   390
            Left            =   2160
            TabIndex        =   27
            Top             =   2160
            Width           =   225
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   390
            Left            =   5400
            TabIndex        =   26
            Top             =   2160
            Width           =   225
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Enabled         =   0   'False
            Height          =   495
            Left            =   4320
            TabIndex        =   24
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Click Here !"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   4440
            TabIndex        =   20
            Top             =   720
            Width           =   1215
         End
         Begin VB.Shape Shape29 
            BorderColor     =   &H00800080&
            BorderWidth     =   2
            FillColor       =   &H00FF80FF&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   4320
            Shape           =   4  'Rounded Rectangle
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&PAYMENT"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   405
            Left            =   3000
            TabIndex        =   17
            Top             =   3360
            Width           =   1725
         End
         Begin VB.Shape Shape18 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00C000C0&
            FillStyle       =   0  'Solid
            Height          =   615
            Left            =   2880
            Shape           =   4  'Rounded Rectangle
            Top             =   3240
            Width           =   1935
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " &DELIVERY"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   405
            Left            =   720
            TabIndex        =   16
            Top             =   3360
            Width           =   1875
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00C000C0&
            FillStyle       =   0  'Solid
            Height          =   615
            Left            =   720
            Shape           =   4  'Rounded Rectangle
            Top             =   3240
            Width           =   2055
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   4440
            TabIndex        =   15
            Top             =   2160
            Width           =   660
         End
         Begin VB.Label Label7 
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
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   4440
            TabIndex        =   7
            Top             =   1680
            Width           =   810
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Top             =   2640
            Width           =   1005
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Price :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Top             =   2160
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   1200
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Number :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1770
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order Number :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   285
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1605
         End
         Begin VB.Shape Shape42 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   360
            Top             =   3480
            Width           =   615
         End
         Begin VB.Shape Shape60 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00C000C0&
            FillStyle       =   0  'Solid
            Height          =   615
            Left            =   4920
            Shape           =   4  'Rounded Rectangle
            Top             =   3240
            Width           =   2055
         End
         Begin VB.Shape Shape61 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   2520
            Top             =   3480
            Width           =   615
         End
         Begin VB.Shape Shape40 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   4440
            Top             =   3480
            Width           =   615
         End
         Begin VB.Shape Shape41 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   6840
            Top             =   3480
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
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
         Left            =   3000
         ScaleHeight     =   1095
         ScaleWidth      =   5775
         TabIndex        =   28
         Top             =   360
         Width           =   5775
         Begin VB.Line Line3 
            BorderColor     =   &H00FFC0FF&
            BorderWidth     =   3
            X1              =   0
            X2              =   5760
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACTIONS"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   630
            Left            =   240
            TabIndex        =   63
            Top             =   120
            Width           =   4440
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   2280
            TabIndex        =   29
            Top             =   1560
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin VB.Shape Shape38 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF80FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2040
         Top             =   6120
         Width           =   7695
      End
      Begin VB.Shape Shape21 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   360
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thank's ! and Come Again !"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   465
         Left            =   2520
         TabIndex        =   21
         Top             =   1560
         Width           =   4950
      End
      Begin VB.Shape Shape37 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF80FF&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   6600
         Shape           =   2  'Oval
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Shape Shape30 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF80FF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2040
         Top             =   1560
         Width           =   5295
      End
      Begin VB.Shape Shape34 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   11040
         Top             =   7320
         Width           =   255
      End
      Begin VB.Shape Shape33 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   480
         Top             =   7320
         Width           =   255
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&ORDER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   540
         Left            =   5160
         TabIndex        =   19
         Top             =   8760
         Width           =   1635
      End
      Begin VB.Shape Shape25 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   6
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   4680
         Shape           =   4  'Rounded Rectangle
         Top             =   8640
         Width           =   2655
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   435
         Left            =   11160
         TabIndex        =   18
         Top             =   2520
         Width           =   285
      End
      Begin VB.Shape Shape24 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   11160
         Top             =   3000
         Width           =   255
      End
      Begin VB.Shape Shape22 
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   10920
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   735
      End
      Begin VB.Shape Shape17 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9480
         Shape           =   3  'Circle
         Top             =   7800
         Width           =   975
      End
      Begin VB.Shape Shape15 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9480
         Shape           =   3  'Circle
         Top             =   6600
         Width           =   975
      End
      Begin VB.Shape Shape13 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9480
         Shape           =   3  'Circle
         Top             =   4440
         Width           =   975
      End
      Begin VB.Shape Shape12 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9480
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   975
      End
      Begin VB.Shape Shape11 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   7800
         Width           =   975
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   6600
         Width           =   975
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   975
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   975
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   1440
         Top             =   6840
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   1440
         Top             =   3360
         Width           =   735
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   9480
         Top             =   6840
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   9480
         Top             =   3480
         Width           =   855
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   9600
         Shape           =   3  'Circle
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   9600
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   960
         Shape           =   3  'Circle
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   3135
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Shape Shape23 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   10560
         Top             =   3960
         Width           =   735
      End
      Begin VB.Shape Shape28 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   8760
         Width           =   735
      End
      Begin VB.Shape Shape27 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   6960
         Shape           =   3  'Circle
         Top             =   8760
         Width           =   735
      End
      Begin VB.Shape Shape26 
         BorderColor     =   &H00FFC0FF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   5040
         Shape           =   4  'Rounded Rectangle
         Top             =   8520
         Width           =   1695
      End
      Begin VB.Shape Shape35 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   720
         Top             =   8880
         Width           =   3855
      End
      Begin VB.Shape Shape36 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7440
         Top             =   8880
         Width           =   3615
      End
      Begin VB.Shape Shape31 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   720
         Top             =   7320
         Width           =   495
      End
      Begin VB.Shape Shape32 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   10560
         Top             =   7320
         Width           =   495
      End
      Begin VB.Shape Shape20 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   2055
         Left            =   9840
         Top             =   4680
         Width           =   255
      End
      Begin VB.Shape Shape45 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   480
         Top             =   3840
         Width           =   735
      End
      Begin VB.Shape Shape39 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   3255
         Left            =   360
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape Shape19 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   2055
         Left            =   1680
         Top             =   4680
         Width           =   255
      End
   End
End
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Con As New ADODB.Connection
Private Ord As New ADODB.Recordset
Private Details As New ADODB.Recordset
Private Del As New ADODB.Recordset
Private Prod As New ADODB.Recordset
Private Sale As New ADODB.Recordset
Private WithEvents Istext As TextBox
Attribute Istext.VB_VarHelpID = -1
Dim s As String, I As Integer, sel As Boolean, Z As Integer, Q As Integer
Dim X(50), Y(50), Pace(50), Size(50) As Integer





Private Sub Combo1_Click()
If Combo1.Text = "" Then: Exit Sub
Prod.Open "Select ProdNo from Products where ProdName='" & Combo1.Text & "'", Con, 3, 3
    If Not Prod.EOF Then
        Text16.Text = Prod("ProdNo")
    Else
        Exit Sub
    End If
    
Text18.SelLength = Len(Text18.Text)
Text18.SetFocus
Prod.Close
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Activate()
          Randomize
For I = 1 To 50
          X1 = Rnd * Me.Width
          Y1 = Rnd * Me.Height
          pace1 = 500 - (Int(Rnd * 499))
          size1 = Rnd * 16
          X(I) = X1
          Y(I) = Y1
          Pace(I) = pace1
          Size(I) = size1
Next
End Sub

Private Sub Form_Load()
    sel = True
    Frame1.Height = 35
    MS2.SelectionMode = 1
    MS1.SelectionMode = 1
 
    Me.BackColor = vbDesktop
    Frame1.BackColor = vbDesktop
    Frame1.Move (Screen.Width - Frame1.Width) / 2, (Screen.Height - Frame1.Height) / 1.1
    s = "Thank's ! and Come Again !"
With Con
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Grocery.mdb"
End With
    Label43.Caption = Date
End Sub

Private Sub Form_Resize()
For Q = 1 To 250
    Picture1.ForeColor = RGB(Q + 1, 0, Q + 1)
    Picture1.CurrentX = Q
    Picture1.CurrentY = Q
    Picture1.Print " T R A N S A C T I O N "
Next
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.ForeColor = &HFFC0FF
    Label9.ForeColor = &HFFC0FF
    Shape1.FillColor = &HC000C0
    Shape18.FillColor = &HC000C0

    Shape25.FillColor = &H800080
    Label13.ForeColor = &HFFC0FF

    Shape18.FillColor = &HC000C0
    Shape19.FillColor = &H800080
    Shape20.FillColor = &H800080
    Shape23.FillColor = &H800080
    Shape24.FillColor = &H800080
    Shape31.FillColor = &H800080
    Shape32.FillColor = &H800080
    Shape33.FillColor = &H800080
    Shape34.FillColor = &H800080
    Shape35.FillColor = &H800080
    Shape36.FillColor = &H800080
    Shape21.FillColor = &H800080
    Shape39.FillColor = &H800080
    Shape45.FillColor = &H800080
    
    Shape1.BorderWidth = 2
    Shape18.BorderWidth = 2

    Timer1.Enabled = False
    
    Shape6.FillColor = &H800080
    Shape14.FillColor = &H800080
     Shape10.FillColor = &H800080
    Shape16.FillColor = &H800080
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.ForeColor = &HFFC0FF
    Label9.ForeColor = &HFFC0FF
    Label45.ForeColor = &HFFC0FF
    Shape60.FillColor = &HC000C0
    Shape60.BorderWidth = 2
    Shape1.FillColor = &HC000C0
    Shape18.FillColor = &HC000C0
    Shape19.FillColor = &H800080
    Shape20.FillColor = &H800080
    Shape23.FillColor = &H800080
    Shape24.FillColor = &H800080
    Shape31.FillColor = &H800080
    Shape32.FillColor = &H800080
    Shape33.FillColor = &H800080
    Shape34.FillColor = &H800080
    Shape35.FillColor = &H800080
    Shape36.FillColor = &H800080
    Shape21.FillColor = &H800080
    Shape39.FillColor = &H800080
    Shape45.FillColor = &H800080
    Shape1.BorderWidth = 2
    Shape18.BorderWidth = 2

End Sub

Private Sub Istext_LostFocus()
Istext.BackColor = &H800080
Istext.ForeColor = &HFFC0FF
End Sub

Private Sub Label10_Click()
On Error Resume Next
        Sale.Open "Select SaleNo from Sales", Con, 3, 3
            Sale.MoveLast
                Dump = Sale("SaleNo")
                Dump = Right(Dump, 1)
                Crash = Dump
                    If Crash <> 0 Then
                        Crash = Crash + 1
                    Else
                        Crash = 1
                    End If
                    Text10.Text = "SN-" & Format(Crash, "000")
                    
        Frame4.Visible = True
        Text12.SetFocus
        Sale.Close
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.ForeColor = vbBlue
    Shape18.FillColor = &HFFFF00
    Shape18.BorderWidth = 6
    Shape19.FillColor = vbWhite
    Shape20.FillColor = vbWhite
    Shape23.FillColor = vbWhite
    Shape24.FillColor = vbWhite
    Shape31.FillColor = vbWhite
    Shape32.FillColor = vbWhite
    Shape33.FillColor = vbWhite
    Shape34.FillColor = vbWhite
    Shape35.FillColor = vbWhite
    Shape36.FillColor = vbWhite
    Shape21.FillColor = vbWhite
    Shape39.FillColor = vbWhite
    Shape45.FillColor = vbWhite

    Label9.ForeColor = vbBlack
    Shape1.FillColor = &HC000C0
    Shape1.BorderWidth = 2

End Sub

Private Sub Label12_Click()
If MS2.Rows <> 1 Then
    MsgBox ("Cancel all Transaction first by deleting their Orders !"), vbOKOnly + vbCritical, "NOTICE !"
Else
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer5.Enabled = False
    Timer7.Enabled = True
End If
End Sub

Private Sub Label13_Click()
On Error Resume Next
Dim bat As Integer, What As String
Ord.Open "Select OrderID from Orders", Con, 3, 3
        Ord.MoveLast
            What = Ord("OrderID")
            What = Right(What, 1)
            bat = What
                If bat <> 0 Then
                    bat = bat + 1
                Else
                    bat = 1
                End If
                    Text1.Text = "Ord-" & Format(bat, "000")
Label13.Enabled = False
Text6.Locked = False
Label16.Enabled = True
    Call Label16_Click
    Text9.Text = 0
With Con
    .BeginTrans
    .Execute "Insert into Orders(OrderID) values ('" & Text1.Text & "');"
    .CommitTrans
End With
Ord.Close
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    Shape25.FillColor = &HFFFF00
    Label13.ForeColor = vbBlue
End Sub



Private Sub Label16_Click()
Prod.Open "Select * from Products", Con, 3, 3
        Do While Not Prod.EOF
            MS1.AddItem Prod("ProdNo") & vbTab & Prod("SuppID") & vbTab & Prod("ProdName") & vbTab & Prod("StockUnit") & vbTab & Prod("PriceUnit") & vbTab & Prod("ReOrderLevel") & vbTab & Prod("Stocks")
            Prod.MoveNext
        Loop
Prod.Close
    Label16.Enabled = False
    Frame3.Visible = True
End Sub

Private Sub Label26_Click()
If Val(Text12.Text) < Val(Text11.Text) Then
            MsgBox ("Amount Pay is Less than the Total Amount ! Therefore Cannot be Accepted!"), vbOKOnly + vbExclamation, "NOTE'S !"
            Text12.SelLength = Len(Text12.Text)
            Text12.SetFocus
Else
        With Con
            .BeginTrans
            .Execute "Insert into Sales(SaleNo,OrderID,Gtotal,TransactBy,SaleDate) values ('" & Text10.Text & "','" & Text1.Text & "'," & Text11.Text & ",'" & Text14.Text & "',#" & Label43.Caption & "#);"
            .CommitTrans
            
            .BeginTrans
            .Execute "Update Orders set Gtotal=" & Text11.Text & ",OrderDate=#" & Label43.Caption & "#  where OrderID='" & Text1.Text & "'"
            .CommitTrans
            
            .BeginTrans
            .Execute "Update Order_Details set SaleNo='" & Text10.Text & "' where OrderID='" & Text1.Text & "'"
            .CommitTrans
        End With
            MsgBox ("Thank you for Purchasing !"), vbOKOnly + vbInformation, "Come Again !"
            MS2.Rows = 1
            ClearText
            Label45.Enabled = True
            Label10.Enabled = True
            Frame4.Visible = False
End If
End Sub

Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label26.ForeColor = &HFFFF00
End Sub

Private Sub Label27_Click()
Frame4.Visible = False
End Sub

Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label27.ForeColor = &HFFFF00
End Sub

Private Sub Label36_Click()
Dim icount As Integer, bet As String
On Error Resume Next
Combo1.Clear
Prod.Open "Select ProdName from Products", Con, 3, 3
    Do While Not Prod.EOF
        Combo1.AddItem Prod("ProdName")
        Prod.MoveNext
    Loop
Prod.Close
Del.Open "Select DeliverNo from Deliveries", Con, 3, 3
        Del.MoveLast
            bet = Del("DeliverNo")
            bet = Right(bet, 1)
            icount = bet
                If icount <> 0 Then
                    icount = icount + 1
                Else
                    icount = 1
                End If
            Text15.Text = "DN-" & Format(icount, "000")
            Combo1.SetFocus
Label37.Enabled = True
Label36.Enabled = False
Label38.Enabled = False
Combo1.Locked = False
Text18.Locked = False
Text19.Locked = False
Del.Close
End Sub

Private Sub Label36_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.ForeColor = &HFFFF00
End Sub

Private Sub Label37_Click()
Dim Numb As Integer
If Text18.Text = "" Or Text19.Text = "" Or Combo1.Text = "" Then Exit Sub
Prod.Open "Select Stocks from Products where ProdNo='" & Text16.Text & "'", Con, 3, 3
    Numb = Prod("Stocks")
    Numb = Numb + Val(Text19.Text)
Prod.Close
With Con
    .BeginTrans
    .Execute "Insert into Deliveries(DeliverNo,ProdNo,DeliveryDate,ProdName,Quantity) values ('" & Text15.Text & "','" & Text16.Text & "',#" & Text18.Text & "#,'" & Combo1.Text & "'," & Text19.Text & ");"
    .CommitTrans
    
    .BeginTrans
    .Execute "Update Products set Stocks=" & Numb & " where ProdNo='" & Text16.Text & "'"
    .CommitTrans
End With
    MsgBox ("Records Updated !"), vbOKOnly + vbInformation, "SUCCESSFULLY !"
Label37.Enabled = False
Label36.Enabled = True
Label38.Enabled = True


Text19.Text = ""
Combo1.Text = ""
Text16.Text = ""
Text15.Text = ""
Text18.Text = ""

Combo1.Locked = True
Text18.Locked = True
Text19.Locked = True
End Sub

Private Sub Label37_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label37.ForeColor = &HFFFF00
End Sub

Private Sub Label38_Click()
Frame5.Visible = False
End Sub

Private Sub Label38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label38.ForeColor = &HFFFF00
End Sub

Private Sub Label39_Click()
Frame5.Visible = False
Label36.Enabled = True
Label37.Enabled = False
Label38.Enabled = True
Text18.Text = ""
Text19.Text = ""
Combo1.Text = ""
Text16.Text = ""
Text15.Text = ""
End Sub

Private Sub Label45_Click()
Call Label13_Click
MS2.Rows = 1
End Sub

Private Sub Label45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label45.ForeColor = vbBlue
    Shape60.FillColor = &HFFFF00
    Shape60.BorderWidth = 6
    Shape19.FillColor = vbWhite
    Shape20.FillColor = vbWhite
    Shape23.FillColor = vbWhite
    Shape24.FillColor = vbWhite
    Shape31.FillColor = vbWhite
    Shape32.FillColor = vbWhite
    Shape33.FillColor = vbWhite
    Shape34.FillColor = vbWhite
    Shape35.FillColor = vbWhite
    Shape36.FillColor = vbWhite
    Shape21.FillColor = vbWhite
    Shape39.FillColor = vbWhite
    Shape45.FillColor = vbWhite

    Label10.ForeColor = vbBlack
    Shape18.FillColor = &HC000C0
    Shape18.BorderWidth = 2
End Sub

Private Sub Label9_Click()
Frame5.Visible = True
Combo1.Clear
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label9.ForeColor = vbBlue
    Shape1.FillColor = &HFFFF00
    Shape1.BorderWidth = 6
    Shape19.FillColor = vbWhite
    Shape20.FillColor = vbWhite
    Shape23.FillColor = vbWhite
    Shape24.FillColor = vbWhite
    Shape31.FillColor = vbWhite
    Shape32.FillColor = vbWhite
    Shape33.FillColor = vbWhite
    Shape34.FillColor = vbWhite
    Shape35.FillColor = vbWhite
    Shape36.FillColor = vbWhite
    Shape21.FillColor = vbWhite
    Shape39.FillColor = vbWhite
    Shape45.FillColor = vbWhite
    Label10.ForeColor = vbBlack
    Shape18.FillColor = &HC000C0
    Shape18.BorderWidth = 2

End Sub

Private Sub MSFlexGrid2_Click()

End Sub

Private Sub MS1_Click()
    Text6.Locked = False
    Frame3.Visible = False
    Label16.Enabled = True
    Text3.Text = MS1.TextMatrix(MS1.Row, 0)
 
    Text4.Text = MS1.TextMatrix(MS1.Row, 2)
    Text7.Text = MS1.TextMatrix(MS1.Row, 6)
    Text2.Text = MS1.TextMatrix(MS1.Row, 5)
    Text5.Text = MS1.TextMatrix(MS1.Row, 4)
    
    Text6.SetFocus
    MS1.Rows = 1
End Sub

Private Sub MS2_DblClick()
Dim Tam As Integer
On Error GoTo Salbahis
If Not MS2.Rows = 1 Then
Prod.Open "Select Stocks from Products where ProdNo='" & MS2.TextMatrix(MS2.Row, 0) & "'", Con, 3, 3
        Tam = Prod("Stocks")
        Tam = Tam + Val(MS2.TextMatrix(MS2.Row, 3))
Prod.Close
    With Con
        .BeginTrans
        .Execute "Update Products set Stocks=" & Tam & " where ProdNo='" & MS2.TextMatrix(MS2.Row, 0) & "'"
        .CommitTrans
        
        .BeginTrans
        .Execute "Delete * from Order_Details where ProdName='" & MS2.TextMatrix(MS2.Row, 1) & "'"
        .CommitTrans
        
        Text9.Text = CCur(Text9.Text) - CCur(MS2.TextMatrix(MS2.Row, 4))
        MS2.RemoveItem MS2.Row
    End With
Else
   With Con
        .BeginTrans
        .Execute "delete * from Orders where OrderID='" & Text1.Text & "'"
        .CommitTrans
    End With
End If
Salbahis:
If Err = 30015 Then
    Label10.Enabled = False
    MS2.Rows = 1
Else
    Resume Next
End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label26.ForeColor = &HFFC0FF
Label27.ForeColor = &HFFC0FF
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label36.ForeColor = &HFFC0FF
Label37.ForeColor = &HFFC0FF
Label38.ForeColor = &HFFC0FF
End Sub

Private Sub Text12_Change()
On Error Resume Next
If Val(Text12.Text) >= Val(Text11.Text) Then
    Text13.Text = CCur(Text12.Text) - CCur(Text11.Text)
Else
    Text13.Text = "0.00"
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        If Val(Text12.Text) < Val(Text11.Text) Then
            MsgBox ("Amount Pay is Less than the Total Amount ! Therefore Cannot be Accepted!"), vbOKOnly + vbExclamation, "NOTE'S !"
            Text12.SelLength = Len(Text12.Text)
            Text12.SetFocus
        Else
        Text14.SetFocus
        End If
    Case 8, 46
    Case 47 To 58
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text19.SelLength = Len(Text19.Text)
    Text19.SetFocus
End If
End Sub


Private Sub Text2_Change()
Dim message As String
On Error Resume Next
message = "Minimum Level Reach !" + vbCr
message = message & " Need to Order " & Text4.Text
If Val(Text7.Text) <= Val(Text2.Text) Then
    MsgBox message, vbOKOnly + vbExclamation, "WARNING !"
End If
End Sub
Private Sub Text6_Change()
On Error Resume Next
 Text8.Text = CCur(Text5.Text) * CCur(Text6.Text)
 Text11.Text = Text9.Text
If Val(Text6.Text) > Val(Text7.Text) Then
    MsgBox (" We are out of Stock !"), vbOKOnly + vbExclamation, "Warning!"
    Text6.Text = ""
    Text8.Text = "0.00"
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Dim Crash As Integer
On Error GoTo Mark
Select Case KeyAscii
    Case 13
        If Text6.Text = "" Or Text6.Text = 0 Then Exit Sub
            Crash = Val(Text7.Text) - Val(Text6.Text)
                With Con
                    .BeginTrans
                    .Execute "Insert into Order_Details(OrderID,ProdName,Quantity,PriceUnit,Total,OrderDate) values ('" & Text1.Text & "','" & Text4.Text & "'," & Text6.Text & "," & Text5.Text & "," & Text8.Text & ",#" & Label43.Caption & "#);"
                    .CommitTrans
                    
                    .BeginTrans
                    .Execute "Update Products set Stocks=" & Crash & " where ProdNo='" & Text3.Text & "'"
                    .CommitTrans
                End With
        Text9.Text = CCur(Text9.Text) + CCur(Text8.Text)
            
        MS2.Rows = MS2.Rows + 1
        MS2.TextMatrix(MS2.Rows - 1, 0) = Text3.Text
        MS2.TextMatrix(MS2.Rows - 1, 1) = Text4.Text
        MS2.TextMatrix(MS2.Rows - 1, 2) = Text5.Text
        MS2.TextMatrix(MS2.Rows - 1, 3) = Text6.Text
        MS2.TextMatrix(MS2.Rows - 1, 4) = Text8.Text
    
        Text2.Text = ""
        Text7.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text8.Text = ""
        
        Label10.Enabled = True
        
    Case 8, 46
    Case 48 To 57
    Case Else
        KeyAscii = 0
End Select

Mark:
Exit Sub
End Sub

Private Sub Timer1_Timer()
If Shape31.FillColor = vbWhite Then
    Shape31.FillColor = &H800080
    Shape32.FillColor = &H800080
    Shape33.FillColor = &H800080
    Shape34.FillColor = &H800080
    Shape35.FillColor = &H800080
    Shape36.FillColor = &H800080
    Shape19.FillColor = &H800080
    Shape20.FillColor = &H800080
    Shape10.FillColor = &HFFC0FF
    Shape16.FillColor = &HFFC0FF
    Shape6.FillColor = &HFFC0FF
    Shape14.FillColor = &HFFC0FF
ElseIf Shape31.FillColor = &H800080 Then
    Shape31.FillColor = vbWhite
    Shape32.FillColor = vbWhite
    Shape33.FillColor = vbWhite
    Shape34.FillColor = vbWhite
    Shape35.FillColor = vbWhite
    Shape36.FillColor = vbWhite
    Shape19.FillColor = vbWhite
    Shape20.FillColor = vbWhite
    Shape10.FillColor = &H800080
    Shape16.FillColor = &H800080
    Shape6.FillColor = &H800080
    Shape14.FillColor = &H800080
End If
End Sub

Private Sub Timer2_Timer()
    Label14.Visible = Label14.Visible Xor True
End Sub

Private Sub Timer3_Timer()
On Error GoTo Mark
Picture1.Cls

If sel = True Then
        Label15.Caption = Left(s, I)
        I = I + 1
    If I > Len(s) Then
        I = 0
            If Label15.Alignment = 0 Then
                Label15.Alignment = 1
            ElseIf Label15.Alignment = 1 Then
                Label15.Alignment = 2
            Else
                Label15.Alignment = 0
                sel = False
            End If
    End If
Else
    Label15.Caption = Mid(Label15.Caption, 1, Len(Label15.Caption) - 1)
Mark:
If Err = 5 Then
    sel = True
Else
    Resume Next
End If
End If
End Sub

Private Sub Timer4_Timer()
Label23.Caption = "TRANSACTION"
End Sub

Private Sub Timer5_Timer()
If TypeOf Me.ActiveControl Is TextBox Then
    Set Istext = Me.ActiveControl
    Me.ActiveControl.BackColor = &HFFFFC0
    Me.ActiveControl.ForeColor = vbBlack
End If
End Sub

Private Sub Timer6_Timer()
Label23.Caption = "TRANSACTION"
End Sub

Private Sub Timer7_Timer()
If Not Frame1.Height = 35 Then
    Frame1.Height = Frame1.Height - 900
    Frame1.Top = Frame1.Top + 900
Else
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer5.Enabled = False
    Timer7.Enabled = False
    Con.Close
    Unload Me
    Main.Timer6.Enabled = True
End If
End Sub

Private Sub Timer8_Timer()
If Not Frame1.Height = 9935 Then
    Frame1.Height = Frame1.Height + 900
    Frame1.Top = Frame1.Top - 900
Else
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer5.Enabled = True
    Frame3.Move (Frame1.Width - Frame3.Width) / 2, (Frame1.Height - Frame3.Height) / 2
    Timer8.Enabled = False
End If

End Sub

Private Sub Timer9_Timer()
 For I = 1 To 50
          Circle (X(I), Y(I)), Size(I), BackColor
          Y(I) = Y(I) + Pace(I)
          If Y(I) >= Me.Height Then Y(I) = 0: X(I) = Rnd * Me.Width
          Circle (X(I), Y(I)), Size(I)
          Next
End Sub

Function ClearText()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text8.Text = ""
Text9.Text = ""
End Function
