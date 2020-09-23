VERSION 5.00
Begin VB.Form Reports 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   -195
   ClientTop       =   -390
   ClientWidth     =   12375
   DrawWidth       =   3
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Timer Timer4 
         Interval        =   100
         Left            =   5520
         Top             =   6240
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   10440
         Top             =   6240
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   10440
         Top             =   5520
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   3480
         ScaleHeight     =   1095
         ScaleWidth      =   5895
         TabIndex        =   8
         Top             =   360
         Width           =   5895
         Begin VB.Shape Shape52 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   5280
            Shape           =   3  'Circle
            Top             =   360
            Width           =   375
         End
         Begin VB.Shape Shape19 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   240
            Shape           =   3  'Circle
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "REPORTS"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   1035
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   5820
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2640
         Top             =   720
      End
      Begin VB.Shape Shape43 
         BorderColor     =   &H80000001&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   240
         Shape           =   3  'Circle
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   3240
         TabIndex        =   13
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DELIVERY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   960
         TabIndex        =   12
         Top             =   6240
         Width           =   1485
      End
      Begin VB.Shape Shape23 
         BorderColor     =   &H80000001&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   5880
         Shape           =   3  'Circle
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   8880
         TabIndex        =   11
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S A L E S"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6720
         TabIndex        =   10
         Top             =   4440
         Width           =   1200
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H0000FF00&
         Height          =   555
         Left            =   10800
         TabIndex        =   7
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORDERS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1080
         TabIndex        =   6
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIERS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6600
         TabIndex        =   5
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   960
         TabIndex        =   4
         Top             =   2640
         Width           =   1590
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   3240
         TabIndex        =   3
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   8880
         TabIndex        =   2
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   3240
         TabIndex        =   1
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Shape Shape33 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   6735
         Left            =   11280
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H00008000&
         BorderWidth     =   6
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   3360
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   6135
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H80000001&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   240
         Shape           =   3  'Circle
         Top             =   4080
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H80000001&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   5880
         Shape           =   3  'Circle
         Top             =   2280
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000001&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2280
         Width           =   735
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   1335
         Left            =   3720
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   5415
      End
      Begin VB.Shape Shape25 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   8760
         Shape           =   3  'Circle
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape24 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape32 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   9600
         Top             =   840
         Width           =   1935
      End
      Begin VB.Shape Shape26 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1560
         Top             =   840
         Width           =   1575
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   6240
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Shape Shape13 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Width           =   975
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   975
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   6840
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   975
      End
      Begin VB.Shape Shape15 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Shape Shape14 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Shape Shape27 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1335
         Left            =   1560
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1560
         Top             =   3360
         Width           =   255
      End
      Begin VB.Shape Shape20 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   3120
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Shape Shape31 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2400
         Top             =   2640
         Width           =   735
      End
      Begin VB.Shape Shape21 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   8760
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Shape Shape36 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   8040
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Shape Shape22 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   3120
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Shape Shape37 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2400
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Shape Shape30 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   6240
         Shape           =   4  'Rounded Rectangle
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Shape Shape34 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   6840
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Width           =   975
      End
      Begin VB.Shape Shape38 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Shape Shape39 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Shape Shape40 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   7200
         Top             =   3120
         Width           =   255
      End
      Begin VB.Shape Shape42 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   8760
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Shape Shape41 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   8040
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Shape Shape44 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Shape Shape45 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   975
      End
      Begin VB.Shape Shape47 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Shape Shape46 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Shape Shape48 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1560
         Top             =   5160
         Width           =   255
      End
      Begin VB.Shape Shape18 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   5160
         Top             =   2640
         Width           =   855
      End
      Begin VB.Shape Shape28 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   3
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   1560
         Top             =   6960
         Width           =   255
      End
      Begin VB.Shape Shape51 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   5160
         Top             =   4440
         Width           =   855
      End
      Begin VB.Shape Shape50 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   3120
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Shape Shape49 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2400
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Shape Shape35 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1560
         Top             =   7440
         Width           =   9975
      End
   End
End
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Con As New ADODB.Connection
Private Rs As New ADODB.Recordset
Dim O As Integer, T As Integer, Intcount As Integer
Dim X(50), Y(50), Pace(50), Size(50) As Integer
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
    Frame1.Height = 15
    Me.BackColor = vbDesktop
    Frame1.Move (Screen.Width - Frame1.Width) / 2, (Screen.Height - Frame1.Height) / 1.15
    Frame1.BackColor = vbDesktop
    T = 1
    O = Rnd * 100
With Con
    .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Grocery.mdb"
End With
DataEnvironment1.Connection1.Open Con
End Sub

Private Sub Form_Unload(Cancel As Integer)
Con.Close
DataEnvironment1.Connection1.Close
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape2.FillColor = vbBlack
    Shape7.FillColor = vbBlack
    Shape12.FillColor = vbBlack
    Shape23.FillColor = vbBlack
    Shape43.FillColor = vbBlack

    Label4.ForeColor = vbWhite
    Label5.ForeColor = vbWhite
    Label6.ForeColor = vbWhite
    Label9.ForeColor = vbWhite
    Label11.ForeColor = vbWhite

    Shape31.BorderWidth = 2
    Shape36.BorderWidth = 2
    Shape37.BorderWidth = 2
    Shape42.BorderWidth = 2
    Shape41.BorderWidth = 2
    Shape49.BorderWidth = 2
    Shape50.BorderWidth = 2
    Shape21.BorderWidth = 2
    Shape22.BorderWidth = 2
    Shape20.BorderWidth = 2
    

    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    
    Shape26.FillColor = &H4000
    Shape27.FillColor = &H4000
    Shape28.FillColor = &H4000
    Shape29.FillColor = &H4000
    Shape35.FillColor = &H4000
    Shape32.FillColor = &H4000
    Shape33.FillColor = &H4000
    Shape48.FillColor = &H4000
    Shape18.FillColor = &H4000
    Shape51.FillColor = &H4000

End Sub

Private Sub Label11_Click()
Dim Lat As String
'On Error GoTo Salbahis
Lat = InputBox(" What  Delivery Date you want to view?", " DELIVERY REPORT")
If Lat = "" Then Exit Sub
DataEnvironment1.Commands(3).CommandText = "Select * from Deliveries where DeliveryDate=#" & Lat & "#"
DataEnvironment1.Commands(3).Execute
Deliveries.Show vbModal
DataEnvironment1.rsCommand3.Close
'Salbahis:
'Exit Sub
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape43.FillColor = &HFFFF00
    Label11.ForeColor = &HFFFF00
    
    Shape49.BorderWidth = 6
    Shape50.BorderWidth = 6
    
    Shape26.FillColor = vbGreen
    Shape27.FillColor = vbGreen
    Shape28.FillColor = vbGreen
    Shape29.FillColor = vbGreen
    Shape35.FillColor = vbGreen
    Shape32.FillColor = vbGreen
    Shape33.FillColor = vbGreen
    Shape48.FillColor = vbGreen
    Shape18.FillColor = vbGreen
    Shape51.FillColor = vbGreen
    
    Label12.Caption = "Use this button to view and print Deliveries Report"
End Sub

Private Sub Label4_Click()
Dim ball As String
'On Error GoTo Salbahis
DataEnvironment1.Commands(4).CommandText = "Select * from Products"
DataEnvironment1.Commands(4).Execute
Products.Show vbModal
DataEnvironment1.rsCommand4.Close
'Salbahis:
'Exit Sub
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape2.FillColor = &HFFFF00
    Label4.ForeColor = &HFFFF00
    Shape31.BorderWidth = 6
    Shape20.BorderWidth = 6

    Shape26.FillColor = vbGreen
    Shape27.FillColor = vbGreen
    Shape28.FillColor = vbGreen
    Shape29.FillColor = vbGreen
    Shape35.FillColor = vbGreen
    Shape32.FillColor = vbGreen
    Shape33.FillColor = vbGreen
    Shape48.FillColor = vbGreen
    Shape18.FillColor = vbGreen
    Shape51.FillColor = vbGreen

    Label1.Caption = "Use this button to view and print Products Report"

End Sub

Private Sub Label5_Click()
'On Error GoTo Salbahis
DataEnvironment1.Commands(5).CommandText = "Select * from Suppliers"
DataEnvironment1.Commands(5).Execute
Supplier.Show vbModal
DataEnvironment1.rsCommand5.Close
'Salbahis:
'Exit Sub
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape7.FillColor = &HFFFF00
    Label5.ForeColor = &HFFFF00
    Shape36.BorderWidth = 6
    Shape21.BorderWidth = 6

    Shape26.FillColor = vbGreen
    Shape27.FillColor = vbGreen
    Shape28.FillColor = vbGreen
    Shape29.FillColor = vbGreen
    Shape35.FillColor = vbGreen
    Shape32.FillColor = vbGreen
    Shape33.FillColor = vbGreen
    Shape48.FillColor = vbGreen
    Shape18.FillColor = vbGreen
    Shape51.FillColor = vbGreen
    

    Label2.Caption = "Use this button to view and print Suppliers Report"
End Sub

Private Sub Label6_Click()
Dim Tat As String
'On Error GoTo Salbahis
Tat = InputBox("What Order Date you want to View?", "ORDER REPORT")
If Tat = "" Then Exit Sub
DataEnvironment1.Commands(2).CommandText = "Select * from Order_Details where OrderDate=#" & Tat & "#"
DataEnvironment1.Commands(2).Execute
Orders.Show vbModal
DataEnvironment1.rsCommand2.Close
Salbahis:
'Exit Sub
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape12.FillColor = &HFFFF00
    Label6.ForeColor = &HFFFF00
    Shape37.BorderWidth = 6
    Shape22.BorderWidth = 6

    Shape26.FillColor = vbGreen
    Shape27.FillColor = vbGreen
    Shape28.FillColor = vbGreen
    Shape29.FillColor = vbGreen
    Shape35.FillColor = vbGreen
    Shape32.FillColor = vbGreen
    Shape33.FillColor = vbGreen
    Shape48.FillColor = vbGreen
    Shape18.FillColor = vbGreen
    Shape51.FillColor = vbGreen

Label3.Caption = "Use this button to view and print Orders Report"
End Sub

Private Sub Label8_Click()
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Label9_Click()
Dim bat As String
'On Error GoTo Salbahis
bat = InputBox(" What  Sales Date would you like to View ? ", " SALES REPORT")
If bat = "" Then Exit Sub
DataEnvironment1.Commands(1).CommandText = "Select * from Sales where SaleDate=#" & bat & "#"
DataEnvironment1.Commands(1).Execute
Sales.Show vbModal
DataEnvironment1.rsCommand1.Close
'Salbahis:
'Exit Sub
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape42.BorderWidth = 6
Shape41.BorderWidth = 6

Label9.ForeColor = &HFFFF00
Shape23.FillColor = &HFFFF00

    Shape26.FillColor = vbGreen
    Shape27.FillColor = vbGreen
    Shape28.FillColor = vbGreen
    Shape29.FillColor = vbGreen
    Shape35.FillColor = vbGreen
    Shape32.FillColor = vbGreen
    Shape33.FillColor = vbGreen
    Shape48.FillColor = vbGreen
    Shape18.FillColor = vbGreen
    Shape51.FillColor = vbGreen

Label10.Caption = "Use this button to view and print Sales Report"
End Sub

Private Sub Timer1_Timer()
If T = 1 Then
    Picture1.Move Picture1.Left + O, Picture1.Top + O
    T = 2
ElseIf T = 2 Then
    Picture1.Move Picture1.Left - O, Picture1.Top + O
    T = 3
ElseIf T = 3 Then
    Picture1.Move Picture1.Left + O, Picture1.Top - O
    T = 4
Else
    Picture1.Move Picture1.Left - O, Picture1.Top - O
    T = 1
End If
End Sub

Private Sub Timer2_Timer()
If Not Frame1.Height = 15 Then
    Frame1.Height = Frame1.Height - 800
    Frame1.Top = Frame1.Top + 800
Else
    Timer1.Enabled = False
    Timer2.Enabled = False
    Main.Timer6.Enabled = True
    Unload Me
End If
End Sub

Private Sub Timer3_Timer()
If Not Frame1.Height = 8015 Then
    Frame1.Height = Frame1.Height + 800
    Frame1.Top = Frame1.Top - 800
Else
        Timer1.Enabled = True
        Timer3.Enabled = False

End If
End Sub

Private Sub Timer4_Timer()
For I = 0 To 50
    Circle (X(I), Y(I)), Size(I), BackColor
    Y(I) = Y(I) + Pace(I)
    If Y(I) >= Me.Height Then Y(I) = 0: X(I) = Rnd * Me.Width
    Circle (X(I), Y(I)), Size(I)
Next
End Sub
