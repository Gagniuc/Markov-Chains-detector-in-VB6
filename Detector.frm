VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo - Markov detector"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18735
   LinkTopic       =   "Form1"
   ScaleHeight     =   677
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1249
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton LoadEx 
      Caption         =   "Load example (16)"
      Height          =   375
      Index           =   15
      Left            =   16560
      TabIndex        =   40
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Preset experiments:"
      Height          =   9135
      Left            =   16320
      TabIndex        =   23
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (18)"
         Height          =   375
         Index           =   17
         Left            =   240
         TabIndex        =   42
         Top             =   8520
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (17)"
         Height          =   375
         Index           =   16
         Left            =   240
         TabIndex        =   41
         Top             =   8040
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (15)"
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   39
         Top             =   7080
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (14)"
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   37
         Top             =   6600
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (13)"
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   36
         Top             =   6120
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (12)"
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   35
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (11)"
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   34
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (10)"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   33
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (9)"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   32
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (8)"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   31
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (7)"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (6)"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (5)"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (4)"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (3)"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (2)"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton LoadEx 
         Caption         =   "Load example (1)"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CheckBox y_axis 
      Caption         =   "Reverse axis [Chart (-; red), (+; blue)]:"
      Height          =   350
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   3135
   End
   Begin VB.Frame Frame4 
      Caption         =   "Real time training vs recorded training"
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   7320
      Width           =   8655
      Begin VB.CommandButton Command2 
         Caption         =   "or load some precomputed matrices (+ and -)"
         Height          =   615
         Left            =   4680
         TabIndex        =   38
         Top             =   240
         Width           =   3855
      End
      Begin VB.CheckBox USEM1M2 
         Caption         =   "Extract transition probabilities from the (+) and (-) sequences,"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   1  'Checked
         Width           =   4815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parameters"
      Height          =   855
      Left            =   2760
      TabIndex        =   11
      Top             =   8400
      Width           =   6015
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1560
         Max             =   20
         Min             =   2
         TabIndex        =   14
         Top             =   360
         Value           =   9
         Width           =   3495
      End
      Begin VB.TextBox Window 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "9"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Sliding window:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output:"
      Height          =   6615
      Left            =   9000
      TabIndex        =   8
      Top             =   2640
      Width           =   7095
      Begin VB.TextBox WMatrix 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input models: "
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   8655
      Begin VB.TextBox Sec1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "Detector.frx":0000
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox Sec2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   4440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "Detector.frx":001B
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "The (+) model:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "The (-) model:"
         Height          =   255
         Left            =   4440
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   989
      TabIndex        =   2
      Top             =   360
      Width           =   14895
      Begin VB.Line Zero 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   992
         Y1              =   64
         Y2              =   64
      End
      Begin VB.Label down_label 
         BackStyle       =   0  'Transparent
         Caption         =   "Like the (-) model (red)"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label up_label 
         BackStyle       =   0  'Transparent
         Caption         =   "Like the (+) model (blue)"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3495
      End
      Begin VB.Shape Window_Shape 
         BorderColor     =   &H00808080&
         BorderStyle     =   2  'Dash
         Height          =   4095
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.TextBox secventata 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Detector.frx":0038
      Top             =   5640
      Width           =   8655
   End
   Begin VB.CommandButton Scan 
      Caption         =   "Scan"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -6480
      Picture         =   "Detector.frx":0145
      Top             =   9480
      Width           =   25290
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   1008
      X2              =   1008
      Y1              =   32
      Y2              =   16
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   1008
      X2              =   1008
      Y1              =   168
      Y2              =   152
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15240
      TabIndex        =   22
      Top             =   1200
      Width           =   375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   1008
      X2              =   904
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label down_val 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15240
      TabIndex        =   20
      Top             =   2310
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   1008
      X2              =   968
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1008
      X2              =   968
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label up_val 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15240
      TabIndex        =   19
      Top             =   270
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Find models in sequence:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ____________________
'  /    Demo - Markov detector      \________________________/       v1.00        |
' |                                                                               |
' |            Name:  Demo - Markov detector V1.0                                 |
' |        Category:  Open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |            Book:  Algorithms in Bioinformatics: Theory and Implementation     |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |  ____________________________________________________________________________ |
' |                                                                               |
' |    Date Created:  January 2014                                                |
' |          Update:  August 2021                                                 |
' |       Tested On:  WinXP, WinVista, Win7, Win8, Win10                          |
' |             Use:  Detection                                                   |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'

Dim M1(1 To 4, 1 To 4) As String
Dim M2(1 To 4, 1 To 4) As String
Dim LogM(1 To 4, 1 To 4) As String

Dim select_start_reminder As Integer
Dim lim_real_time As Integer

Function ver_alpha(tmp) As Boolean

    Dim re(1 To 4) As String
    
    re(1) = "A"
    re(2) = "T"
    re(3) = "G"
    re(4) = "C"
    
    For i = 1 To 4
        tmp = Replace(UCase(tmp), re(i), "")
    Next i
    
    If tmp <> "" Then ver_alpha = True Else ver_alpha = False

End Function



Private Sub Scan_Click()

    Dim Maxyp As Variant
    Dim Maxym As Variant
    Dim y_scale As Variant

    Sec1.Text = Replace(Sec1.Text, vbCrLf, "")
    Sec2.Text = Replace(Sec2.Text, vbCrLf, "")
    secventata.Text = Replace(secventata.Text, vbCrLf, "")
    secventata.Text = Replace(secventata.Text, Chr(13), "")

    If ver_alpha(Sec1.Text) = True Then
        MsgBox "The sequence from model (+) contains symbols outside the alphabet {A, T, G, C}"
        Exit Sub
    End If

    If ver_alpha(Sec2.Text) = True Then
        MsgBox "The sequence from model (-) contains symbols outside the alphabet {A, T, G, C}"
        Exit Sub
    End If
    
    If ver_alpha(secventata.Text) = True Then
        MsgBox "The main sequence contains symbols outside the alphabet {A, T, G, C}"
        Exit Sub
    End If

    
    If USEM1M2.Value = 1 Then
        Call Fill_Transition_Matrix(M1, Sec1.Text)
        Call Fill_Transition_Matrix(M2, Sec2.Text)
    End If
    
    Call Combine_M1M2(M1, M2)
    
    '-----------------------------------------------
    sTXT = DrowMatrix(4, 4, M1, "(+)", "Transition probabilities M1(+):")
    sTXT = sTXT & DrowMatrix(4, 4, M2, "(-)", "Transition probabilities M2(-):")
    sTXT = sTXT & DrowMatrix(4, 4, LogM, "LLR", "log-likelihood ratio Matrix (M1(+) / M2(-):")
    WMatrix.Text = sTXT
    '-----------------------------------------------
    
    lenwin = Val(Window.Text)
    'secventata.Text = Replace(secventata.Text, vbCrLf, "")
    'secventata.Text = Replace(secventata.Text, Chr(13), "")
    seq = UCase(secventata.Text)
    
    Picture1.Cls
    WMatrix.Text = WMatrix.Text & vbCrLf

    Maxyp = 0
    Maxym = 0
    
    For i = 1 To Len(seq) - lenwin + 1
    
        win = Mid(seq, i, lenwin)
        val_motif = Ask_logM(win)
        
        If val_motif > Maxyp Then Maxyp = val_motif
        If val_motif < Maxym Then Maxym = val_motif
        
    Next i
    
    
    If Maxyp = 0 Then Maxyp = 1
    If Maxym = 0 Then Maxym = -1
    
    'This reverses the scale (+) down and (-) up
    If y_axis.Value = 1 Then
    
        If Maxyp > Abs(Maxym) Then y_scale = -(Maxyp) Else y_scale = Maxym
        
        up_val.Caption = y_scale
        down_val.Caption = -y_scale
        
        up_label.Caption = "Like the (-) model (red)"
        down_label.Caption = "Like the (+) model (blue)"
    
    End If
    
    'This shows normal scale, (+) up and (-) down
    If y_axis.Value = 0 Then
    
        If Maxyp > Abs(Maxym) Then y_scale = Maxyp Else y_scale = Abs(Maxym)
        
        up_val.Caption = y_scale
        down_val.Caption = -y_scale
        
        up_label.Caption = "Like the (+) model (blue)"
        down_label.Caption = "Like the (-) model (red)"
    
    End If


    For i = 1 To Len(seq) - lenwin + 1
    
        win = Mid(seq, i, lenwin)
        
        val_motif = Ask_logM(win)
        
        If Len(secventata.Text) < lim_real_time Then WMatrix.Text = WMatrix.Text & "[" & i & "] " & win & " = " & val_motif & vbCrLf
        
        
        sliceX = (Picture1.ScaleWidth / (Len(seq) - lenwin + 1))
        sliceY = (Picture1.ScaleHeight / (y_scale * 2))
    
        OLD_val_motif = val_motif
        
        If OLD_val_motif <> 0 Then
            If val_motif < 0 Then
                Picture1.Line (sliceX * oldi, (Picture1.ScaleHeight / 2))-(sliceX * i, (Picture1.ScaleHeight / 2) - (sliceY * val_motif)), vbRed, BF
            Else
                Picture1.Line (sliceX * oldi, (Picture1.ScaleHeight / 2))-(sliceX * i, (Picture1.ScaleHeight / 2) - (sliceY * val_motif)), vbBlue, BF
            End If
        End If
        
        oldi = i
    
    Next i

End Sub


Function Ask_logM(ByVal s As String)
 
    For i = 1 To Len(s) - 1
    
            DI1 = Mid(s, i, 1)
            DI2 = Mid(s, i + 1, 1)
    
            If DI1 = "A" Then r = 1
            If DI1 = "C" Then r = 2
            If DI1 = "G" Then r = 3
            If DI1 = "T" Then r = 4
            
            If DI2 = "A" Then c = 1
            If DI2 = "C" Then c = 2
            If DI2 = "G" Then c = 3
            If DI2 = "T" Then c = 4
    
            plus = plus + Val(LogM(r, c))
    
    Next i
    
    Ask_logM = plus

End Function



Function Combine_M1M2(ByRef TM1() As String, ByRef TM2() As String)

    For i = 1 To 4
    
        For j = 1 To 4

            If Val(TM1(i, j)) <> 0 And Val(TM2(i, j)) <> 0 Then
                LogM(i, j) = Log(Val(TM1(i, j)) / Val(TM2(i, j))) / Log(2)
                LogM(i, j) = Round(LogM(i, j), 3)
            Else
                LogM(i, j) = 0
            End If
    
        Next j
    
    Next i

End Function



Function Fill_Transition_Matrix(ByRef M() As String, ByVal s As String)

    For i = 1 To 4
    
        For j = 1 To 4
    
            M(i, j) = 0.0000000001
    
        Next j
    
    Next i
    
    
    For i = 1 To Len(s)
    
        DI = Mid(s, i, 1)
    
        If DI = "A" Then nA = nA + 1
        If DI = "C" Then nC = nC + 1
        If DI = "G" Then nG = nG + 1
        If DI = "T" Then nT = nT + 1
    
    Next i
    
    
    For i = 1 To Len(s) - 1
    
            DI1 = Mid(s, i, 1)
            DI2 = Mid(s, i + 1, 1)
    
            If DI1 = "A" Then r = 1
            If DI1 = "C" Then r = 2
            If DI1 = "G" Then r = 3
            If DI1 = "T" Then r = 4
            
            If DI2 = "A" Then c = 1
            If DI2 = "C" Then c = 2
            If DI2 = "G" Then c = 3
            If DI2 = "T" Then c = 4
    
            M(r, c) = Val(M(r, c)) + 1
    
    Next i
    
    
    
    For i = 1 To 4
    
        For j = 1 To 4
    
            If nA = 0 Then nA = 1
            If nC = 0 Then nC = 1
            If nG = 0 Then nG = 1
            If nT = 0 Then nT = 1
            
            If i = 1 Then r = nA
            If i = 2 Then r = nC
            If i = 3 Then r = nG
            If i = 4 Then r = nT
            
            M(i, j) = Round(Val(M(i, j)) / r, 7)
    
        Next j
    
    Next i

End Function


Private Sub Command2_Click()

    USEM1M2.Value = 0

    Dim M1Load() As String
    Dim M2Load() As String
    
    pluss = "0.180 0.274 0.426 0.120 0.171 0.368 0.274 0.188 0.161 0.339 0.375 0.125 0.079 0.355 0.384 0.182"
    minus = "0.300 0.205 0.285 0.210 0.322 0.298 0.078 0.302 0.248 0.246 0.298 0.208 0.177 0.239 0.292 0.292"
    
    M1Load() = Split(pluss, " ")
    M2Load() = Split(minus, " ")
    
    a = 0
    For i = 1 To 4
        For j = 1 To 4
            a = a + 1
            M1(i, j) = M1Load(a - 1)
        Next j
    Next i
    
    a = 0
    For i = 1 To 4
        For j = 1 To 4
            a = a + 1
            M2(i, j) = M2Load(a - 1)
        Next j
    Next i
    
    'Transition probabilities
    '-----------------------------------------------
    sTXT = DrowMatrix(4, 4, M1, "(+)", "Ready to use transition probabilities M1(+):")
    sTXT = sTXT & DrowMatrix(4, 4, M2, "(-)", "Ready to use transition probabilities M2(-):")
    WMatrix.Text = sTXT
    '-----------------------------------------------
    
    '+  A      C     G     T    -   A     C     G     T
    'A 0.180 0.274 0.426 0.120  A 0.300 0.205 0.285 0.210
    'C 0.171 0.368 0.274 0.188  C 0.322 0.298 0.078 0.302
    'G 0.161 0.339 0.375 0.125  G 0.248 0.246 0.298 0.208
    'T 0.079 0.355 0.384 0.182  T 0.177 0.239 0.292 0.292

    Sec1.Enabled = False
    Sec2.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False

End Sub


Function DrowMatrix(ib, jb, ByVal M As Variant, ByVal model As String, ByVal msg As String) As String

    '------ Show Matrix in Text OBJ -------------------------------------------
    Y = "|_____|_____|_____|_____|_____|"
    
    ct = ct & vbCrLf & "_______________________________"
    ct = ct & vbCrLf & "| " & model & " |  A  |  C  |  G  |  T  |"
    ct = ct & vbCrLf & Y & vbCrLf
    
    For i = 1 To ib 'Rows
    
        For j = 1 To jb 'cols
        
        v = Round(M(i, j), 2)
        
            If Len(v) = 0 Then u = "|     "
            If Len(v) = 1 Then u = "|    "
            If Len(v) = 2 Then u = "|   "
            If Len(v) = 3 Then u = "|  "
            If Len(v) = 4 Then u = "| "
            If Len(v) = 5 Then u = "|"
            
            If j = jb Then o = "|" Else o = ""
            
            If j = 1 And i = 1 Then ct = ct & "|  A  "
            If j = 1 And i = 2 Then ct = ct & "|  C  "
            If j = 1 And i = 3 Then ct = ct & "|  G  "
            If j = 1 And i = 4 Then ct = ct & "|  T  "
            
            ct = ct & u & v & o
            
        Next j
    
    ct = ct & vbCrLf & Y & vbCrLf
    
    Next i
    '--------------------------------------------------------------------------
    DrowMatrix = msg & " M[" & Val(jb) & "," & Val(ib) & "]" & vbCrLf & ct & vbCrLf & vbCrLf
    '--------------------------------------------------------------------------

End Function


Private Sub Form_Load()

    LoadEx_Click (14)
    
    Label1.Caption = "The (+) model [length: " & Len(Sec1.Text) & "]"
    Label2.Caption = "The (-) model [length: " & Len(Sec2.Text) & "]"
    Label5.Caption = "Find models in sequence [length: " & Len(secventata.Text) & "]"
    
    Line3.Y1 = Picture1.Top + (Picture1.ScaleHeight / 2)
    Line3.Y2 = Line3.Y1
    Label6.Top = Line3.Y1 - (Label6.Height / 2) + 2
    
    Zero.Y1 = (Picture1.ScaleHeight / 2)
    Zero.Y2 = (Picture1.ScaleHeight / 2)
    
    lim_real_time = 200
    
End Sub


Private Sub HScroll1_Change()
    Window.Text = HScroll1.Value
    If Len(secventata.Text) < lim_real_time Then Scan_Click
End Sub


Private Sub HScroll1_Scroll()
    Window.Text = HScroll1.Value
    If Len(secventata.Text) < lim_real_time Then Scan_Click
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Window_Shape.Visible = True
    select_start_reminder = X
    Call k_real_time
    
End Sub


Private Sub Sec1_Change()
    Label1.Caption = "The (+) model [length: " & Len(Sec1.Text) & "]"
    If Len(secventata.Text) < lim_real_time Then Scan_Click
End Sub


Private Sub Sec2_Change()
    Label2.Caption = "The (-) model [length: " & Len(Sec2.Text) & "]"
    If Len(secventata.Text) < lim_real_time Then Scan_Click
End Sub


Private Sub secventata_Change()
    Label5.Caption = "Find models in sequence [length: " & Len(secventata.Text) & "]"
    If Len(secventata.Text) < lim_real_time Then Scan_Click
End Sub


Private Sub USEM1M2_Click()

    If USEM1M2.Value = 0 Then
        Sec1.Enabled = False
        Sec2.Enabled = False
        Label1.Enabled = False
        Label2.Enabled = False
    Else
        Sec1.Enabled = True
        Sec2.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
    End If
    
End Sub


Function k_real_time()

    X = select_start_reminder 'global
    
    '---------------------
    q = ((Len(secventata.Text) - Val(Window.Text) + 1) / Picture1.ScaleWidth) * X
    If q < 0 Then Exit Function
        
    secventata.SetFocus
    secventata.SelStart = Int(q)
        
    secventata.SelLength = Val(Window.Text)
    secventaADN = secventata.SelText
    
    If Len(secventata.Text) < 3 Then Exit Function
    
    Window_Shape.Width = (Picture1.ScaleWidth / (Len(secventata.Text) - Val(Window.Text) + 1)) '* Len(secventaADN)
    Window_Shape.Left = select_start_reminder
    '---------------------
        
    If Len(secventaADN) <= 2 Then Exit Function
    If (Len(secventaADN) - cols) <= 0 Then Exit Function

End Function


Private Sub y_axis_Click()
    Scan_Click
End Sub


Private Sub LoadEx_Click(Index As Integer)
    
    lim_real_time = 1
    
    If Index = 0 Then
        HScroll1.Value = 20
        Sec1.Text = "ATGGACTCCAACACTGTGTCAAGCTTTCAGGTAGACTGCTTTCTTTGGCATGTCCGCAAACGATTTGCAG" & _
                    "ACCAAGAACTGGGTGATGCCCCATTCCTTGACCGGCTTCGCCGAGACCAGAAGTCCCTAAGAGGAAGAGG" & _
                    "CAGCACTCTTGGTCTGGACATCGAGACAGCTACTCGTGCGGGAAAGCAAATAGTGGAGCGGATTCTGGGG" & _
                    "GAAGAATCTGATGAAGCACTTAAAATGAATATTGCTTCTGTACCGACTTCACGCTACCTAACTGACATGA" & _
                    "CTCTTGAAGAAATGTCAAGAGACTGGTTCATGCTCATGCCCAAGCAGAAAGTAGCAGGTTCTCTCTGCAT" & _
                    "CAAAATGGACCAGGCAATAATGGATAAAACCATCATACTGAAAGCAAATTTCAGTGTGATTTTTGATCGG" & _
                    "CTGGAAACCCTAATATTACTTAGAGCTTTCACAGAAGAAGGAGCAATTGTGGGAGAAATCTCACCATTAC" & _
                    "CTTCTCTTCCAGGACATACTGATGAGGATGTCAAAATTGCAATTGGGGTCCTCATCGGAGGGCTTGAATG"
        Sec2.Text = "ATTAAAGGTTTATACCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAAA" & _
                    "CGAACTTTAAAATCTGTGTGGCTGTCACTCGGCTGCATGCTTAGTGCACTCACGCAGTATAATTAATAAC" & _
                    "TAATTACTGTCGTTGACAGGACACGAGTAACTCGTCTATCTTCTGCAGGCTGCTTACGGTTTCGTCCGTG" & _
                    "TTGCAGCCGATCATCAGCACATCTAGGTTTCGTCCGGGTGTGACCGAAAGGTAAGATGGAGAGCCTTGTC" & _
                    "CCTGGTTTCAACGAGAAAACACACGTCCAACTCAGTTTGCCTGTTTTACAGGTTCGCGACGTGCTCGTAC" & _
                    "GTGGCTTTGGAGACTCCGTGGAGGAGGTCTTATCAGAGGCACGTCAACATCTTAAAGATGGCACTTGTGG" & _
                    "CTTAGTAGAAGTTGAAAAAGGCGTTTTGCCTCAACTTGAACAGCCCTATGTGTTCATCAAACGTTCGGAT" & _
                    "GCTCGAACTGCACCTCATGGTCATGTTATGGTTGAGCTGGTAGCAGAACTCGAAGGCATTCAGTACGGTC"
        secventata.Text = "ATGAGTCTTCTAACCGAGGTCGAAACGTACGTTCTCTCTATCATCCCGTCAGGCCCCCTCAAAGCCGAGA" & _
                          "TCGCGCAGAGACTTGAAGATGTCTTTGCAGGGAAAAACACCGATCTCGAGGCTCTCATGGAGTGGCTAAA" & _
                          "GACAAGACCAATCCTGTCACCTCTGACTAAAGGGATTTTGGGATTTGTGTTCACGCTCACCGTGCCCAGT" & _
                          "GAGCGAGGACTGCAGCGTAGACGCTTCGTCCAGAATGCCCTAAATGGAAATGGGGATCCAAATAATATGG" & _
                          "ATAAGGCAGTTAAGCTATATAAGAAGCTGAAAAGAGAGATAACATTCCATGGGGCTAAGGAGGTCGCACT" & _
                          "TAGCTACTCAACCGGTGCACTTGCCAGCTGCATGGGTCTCATATACAACAGGATGGGAACGGTGACTACA" & _
                          "GAAGTGGCTTTTGGCCTAGTGTGTGCCACTTGTGAGCAGATTGCAGATTCACAGCATCGGTCCCACAGAC" & _
                          "AGATGGCAACCATCACCAACCCATTAATCAGACATGAGAACAGAATGGTGCTGGCCAGCACTACAGCTAA"
    End If
    
    
    If Index = 1 Then
        HScroll1.Value = 20
        Sec1.Text = "ATGGACTCCAACACTGTGTCAAGCTTTCAGGTAGACTGCTTTCTTTGGCATGTCCGCAAACGATTTGCAG" & _
                    "ACCAAGAACTGGGTGATGCCCCATTCCTTGACCGGCTTCGCCGAGACCAGAAGTCCCTAAGAGGAAGAGG" & _
                    "CAGCACTCTTGGTCTGGACATCGAGACAGCTACTCGTGCGGGAAAGCAAATAGTGGAGCGGATTCTGGGG" & _
                    "GAAGAATCTGATGAAGCACTTAAAATGAATATTGCTTCTGTACCGACTTCACGCTACCTAACTGACATGA" & _
                    "CTCTTGAAGAAATGTCAAGAGACTGGTTCATGCTCATGCCCAAGCAGAAAGTAGCAGGTTCTCTCTGCAT" & _
                    "CAAAATGGACCAGGCAATAATGGATAAAACCATCATACTGAAAGCAAATTTCAGTGTGATTTTTGATCGG" & _
                    "CTGGAAACCCTAATATTACTTAGAGCTTTCACAGAAGAAGGAGCAATTGTGGGAGAAATCTCACCATTAC" & _
                    "CTTCTCTTCCAGGACATACTGATGAGGATGTCAAAATTGCAATTGGGGTCCTCATCGGAGGGCTTGAATG"
        Sec2.Text = "ACTAGTCTACTACGTAGTCATCTTCATCGTCATCATCATCAGTTCTCTCGACGCATCTAGTCTTCATGCC" & _
                    "GATCTATCATTATATATGCGCGGCGATATTATCATCTACGTACTGATGCGACGTACGTATCTACGAGTCA" & _
                    "TCTATCGACGTAGTCATCTACTATTGCCTATCATTCATCGTATCATCTACG"
        secventata.Text = "ATGAGTCTTCTAACCGAGGTCGAAACGTACGTTCTCTCTATCATCCCGTCAGGCCCCCTCAAAGCCGAGA" & _
                          "TCGCGCAGAGACTTGAAGATGTCTTTGCAGGGAAAAACACCGATCTCGAGGCTCTCATGGAGTGGCTAAA" & _
                          "GACAAGACCAATCCTGTCACCTCTGACTAAAGGGATTTTGGGATTTGTGTTCACGCTCACCGTGCCCAGT" & _
                          "GAGCGAGGACTGCAGCGTAGACGCTTCGTCCAGAATGCCCTAAATGGAAATGGGGATCCAAATAATATGG" & _
                          "ATAAGGCAGTTAAGCTATATAAGAAGCTGAAAAGAGAGATAACATTCCATGGGGCTAAGGAGGTCGCACT" & _
                          "TAGCTACTCAACCGGTGCACTTGCCAGCTGCATGGGTCTCATATACAACAGGATGGGAACGGTGACTACA" & _
                          "GAAGTGGCTTTTGGCCTAGTGTGTGCCACTTGTGAGCAGATTGCAGATTCACAGCATCGGTCCCACAGAC" & _
                          "AGATGGCAACCATCACCAACCCATTAATCAGACATGAGAACAGAATGGTGCTGGCCAGCACTACAGCTAA"
    End If
    
    
    If Index = 2 Then
        HScroll1.Value = 20
        Sec1.Text = "ATGGACTCCAACACTGTGTCAAGCTTTCAGGTAGACTGCTTTCTTTGGCATGTCCGCAAACGATTTGCAG" & _
                    "ACCAAGAACTGGGTGATGCCCCATTCCTTGACCGGCTTCGCCGAGACCAGAAGTCCCTAAGAGGAAGAGG" & _
                    "CAGCACTCTTGGTCTGGACATCGAGACAGCTACTCGTGCGGGAAAGCAAATAGTGGAGCGGATTCTGGGG" & _
                    "GAAGAATCTGATGAAGCACTTAAAATGAATATTGCTTCTGTACCGACTTCACGCTACCTAACTGACATGA" & _
                    "CTCTTGAAGAAATGTCAAGAGACTGGTTCATGCTCATGCCCAAGCAGAAAGTAGCAGGTTCTCTCTGCAT" & _
                    "CAAAATGGACCAGGCAATAATGGATAAAACCATCATACTGAAAGCAAATTTCAGTGTGATTTTTGATCGG" & _
                    "CTGGAAACCCTAATATTACTTAGAGCTTTCACAGAAGAAGGAGCAATTGTGGGAGAAATCTCACCATTAC" & _
                    "CTTCTCTTCCAGGACATACTGATGAGGATGTCAAAATTGCAATTGGGGTCCTCATCGGAGGGCTTGAATG"
        Sec2.Text = "ATGAGTCTTCTAACCGAGGTCGAAACGTACGTTCTCTCTATCATCCCGTCAGGCCCCCTCAAAGCCGAGA" & _
                    "TCGCGCAGAGACTTGAAGATGTCTTTGCAGGGAAAAACACCGATCTCGAGGCTCTCATGGAGTGGCTAAA" & _
                    "GACAAGACCAATCCTGTCACCTCTGACTAAAGGGATTTTGGGATTTGTGTTCACGCTCACCGTGCCCAGT" & _
                    "GAGCGAGGACTGCAGCGTAGACGCTTCGTCCAGAATGCCCTAAATGGAAATGGGGATCCAAATAATATGG" & _
                    "ATAAGGCAGTTAAGCTATATAAGAAGCTGAAAAGAGAGATAACATTCCATGGGGCTAAGGAGGTCGCACT" & _
                    "TAGCTACTCAACCGGTGCACTTGCCAGCTGCATGGGTCTCATATACAACAGGATGGGAACGGTGACTACA" & _
                    "GAAGTGGCTTTTGGCCTAGTGTGTGCCACTTGTGAGCAGATTGCAGATTCACAGCATCGGTCCCACAGAC" & _
                    "AGATGGCAACCATCACCAACCCATTAATCAGACATGAGAACAGAATGGTGCTGGCCAGCACTACAGCTAA"
        secventata.Text = "ATTAAAGGTTTATACCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAAA" & _
                          "CGAACTTTAAAATCTGTGTGGCTGTCACTCGGCTGCATGCTTAGTGCACTCACGCAGTATAATTAATAAC" & _
                          "TAATTACTGTCGTTGACAGGACACGAGTAACTCGTCTATCTTCTGCAGGCTGCTTACGGTTTCGTCCGTG" & _
                          "TTGCAGCCGATCATCAGCACATCTAGGTTTCGTCCGGGTGTGACCGAAAGGTAAGATGGAGAGCCTTGTC" & _
                          "CCTGGTTTCAACGAGAAAACACACGTCCAACTCAGTTTGCCTGTTTTACAGGTTCGCGACGTGCTCGTAC" & _
                          "GTGGCTTTGGAGACTCCGTGGAGGAGGTCTTATCAGAGGCACGTCAACATCTTAAAGATGGCACTTGTGG" & _
                          "CTTAGTAGAAGTTGAAAAAGGCGTTTTGCCTCAACTTGAACAGCCCTATGTGTTCATCAAACGTTCGGAT" & _
                          "GCTCGAACTGCACCTCATGGTCATGTTATGGTTGAGCTGGTAGCAGAACTCGAAGGCATTCAGTACGGTC"
    End If
    
    
    If Index = 3 Then
        HScroll1.Value = 20
        Sec1.Text = "ATGGACTCCAACACTGTGTCAAGCTTTCAGGTAGACTGCTTTCTTTGGCATGTCCGCAAACGATTTGCAG" & _
                    "ACCAAGAACTGGGTGATGCCCCATTCCTTGACCGGCTTCGCCGAGACCAGAAGTCCCTAAGAGGAAGAGG" & _
                    "CAGCACTCTTGGTCTGGACATCGAGACAGCTACTCGTGCGGGAAAGCAAATAGTGGAGCGGATTCTGGGG" & _
                    "GAAGAATCTGATGAAGCACTTAAAATGAATATTGCTTCTGTACCGACTTCACGCTACCTAACTGACATGA" & _
                    "CTCTTGAAGAAATGTCAAGAGACTGGTTCATGCTCATGCCCAAGCAGAAAGTAGCAGGTTCTCTCTGCAT" & _
                    "CAAAATGGACCAGGCAATAATGGATAAAACCATCATACTGAAAGCAAATTTCAGTGTGATTTTTGATCGG" & _
                    "CTGGAAACCCTAATATTACTTAGAGCTTTCACAGAAGAAGGAGCAATTGTGGGAGAAATCTCACCATTAC" & _
                    "CTTCTCTTCCAGGACATACTGATGAGGATGTCAAAATTGCAATTGGGGTCCTCATCGGAGGGCTTGAATG"
        Sec2.Text = "ATGAGTCTTCTAACCGAGGTCGAAACGTACGTTCTCTCTATCATCCCGTCAGGCCCCCTCAAAGCCGAGA" & _
                    "TCGCGCAGAGACTTGAAGATGTCTTTGCAGGGAAAAACACCGATCTCGAGGCTCTCATGGAGTGGCTAAA" & _
                    "GACAAGACCAATCCTGTCACCTCTGACTAAAGGGATTTTGGGATTTGTGTTCACGCTCACCGTGCCCAGT" & _
                    "GAGCGAGGACTGCAGCGTAGACGCTTCGTCCAGAATGCCCTAAATGGAAATGGGGATCCAAATAATATGG" & _
                    "ATAAGGCAGTTAAGCTATATAAGAAGCTGAAAAGAGAGATAACATTCCATGGGGCTAAGGAGGTCGCACT" & _
                    "TAGCTACTCAACCGGTGCACTTGCCAGCTGCATGGGTCTCATATACAACAGGATGGGAACGGTGACTACA" & _
                    "GAAGTGGCTTTTGGCCTAGTGTGTGCCACTTGTGAGCAGATTGCAGATTCACAGCATCGGTCCCACAGAC" & _
                    "AGATGGCAACCATCACCAACCCATTAATCAGACATGAGAACAGAATGGTGCTGGCCAGCACTACAGCTAA"
        secventata.Text = "ATGGCGTCTCAAGGCACCAAACGATCTTATGAACAGATGGAAACTGGTGGAGAACGCCAGAATGCCACTG"
    End If
    
    
    If Index = 4 Then
        HScroll1.Value = 20
        Sec1.Text = "ATAACATTATTTTTCGATTGGGAATGGCGCTTACCATTCCTGGAGCCAGACAATTAGTAAGCCATAGACA"
        Sec2.Text = "TATCCGCCTAAAAGAAAAACAAAGAGTACGTTTTCATTACGGACTTACAGAGCGACAATTACTTCAAT"
        secventata.Text = "ATGTCCCGTTATCGAGGACCTCGTTTTAAAAAAATACGCCGTCTGGGAGCTTTACCAGGACTCACTAGAA" & _
                          "AAACACCTAAATCCGGAAGTAATCTGAAAAAGAAATTCCATTCTGGGAAAAAAGAACAATATCGTATTCG" & _
                          "TCTTCAAGAAAAACAGAAATTGCGTTTTCATTATGGTCTGACAGAACGACAATTACTTAGATATGTACAT" & _
                          "ATCGCTGGAAAAGCAAAAAGTTCAACAGGTCAGGTTTTACTACAATTACTTGAAATGCGTTTGGATAATA" & _
                          "TCCTTTTTCGATTAGGTATGGCTTCAACCATTCCTGAGGCCCGGCAATTAGTTAACCATAGACATATTTT" & _
                          "AGTTAATGGTCGTATAGTCGATATACCAAGTTTTCGTTGCAAACCCCGAGATATTATTACTACGAAGGAT" & _
                          "AACCAAAGATCAAAACGTCTGGTTCAAAATTCTATTGCTTCATCCGATCCGGGGAAATTGCCAAAGCATT" & _
                          "TGACGATTGACACATTGCAATATAAAGGACTAGTAAAAAAAATCCTAGATAGGAAGTGGGTCGGTCTCAA" & _
                          "AATAAATGAGTTGTTAGTTGTAGAATAT"
    End If
    
    
    If Index = 5 Then
        HScroll1.Value = 20
        Sec1.Text = "GATCACAGGTCTATCACCCTATTAACCACTCACGGGAGCTCTCCATGCATTTGGT"
        Sec2.Text = "AAATTGTCATGCTCTGACAGCCCTGAGGATCCCCAAGAGATGACTTCGAAGAGTAGCCTACAGAAAGTTA"
        secventata.Text = "ATGTCCCGTTATCGAGGACCTCGTTTTAAAAAAATACGCCGTCTGGGAGCTTTACCAGGACTCACTAGAA" & _
                          "AAACACCTAAATCCGGAAGTAATCTGAAAAAGAAATTCCATTCTGGGAAAAAAGAACAATATCGTATTCG" & _
                          "TCTTCAAGAAAAACAGAAATTGCGTTTTCATTATGGTCTGACAGAACGACAATTACTTAGATATGTACAT" & _
                          "ATCGCTGGAAAAGCAAAAAGTTCAACAGGTCAGGTTTTACTACAATTACTTGAAATGCGTTTGGATAATA" & _
                          "TCCTTTTTCGATTAGGTATGGCTTCAACCATTCCTGAGGCCCGGCAATTAGTTAACCATAGACATATTTT" & _
                          "AGTTAATGGTCGTATAGTCGATATACCAAGTTTTCGTTGCAAACCCCGAGATATTATTACTACGAAGGAT" & _
                          "AACCAAAGATCAAAACGTCTGGTTCAAAATTCTATTGCTTCATCCGATCCGGGGAAATTGCCAAAGCATT" & _
                          "TGACGATTGACACATTGCAATATAAAGGACTAGTAAAAAAAATCCTAGATAGGAAGTGGGTCGGTCTCAA" & _
                          "AATAAATGAGTTGTTAGTTGTAGAATAT"
    End If
    
    
    If Index = 6 Then
        HScroll1.Value = 20
        Sec1.Text = "TTTATATAGAGGAGACAAGTCGTAACATGGTAAGTGTACTGGAAAGTGCACTTGGACGAACCAGAGTGTA"
        Sec2.Text = "TCCCAGCACAGAGAAAAGGTAGATCTGAATGCTGAACCCCTATATGGAAGAAGAAAACTGAACAAACAG"
        secventata.Text = "GATCACAGGTCTATCACCCTATTAACCACTCACGGGAGCTCTCCATGCATTTGGTATTTTCGTCTGGGGG" & _
                          "GTATGCACGCGATAGCATTGCGAGACGCTGGAGCCGGAGCACCCTATGTCGCAGTATCTGTCTTTGATTC" & _
                          "CTGCCTCATCCTATTATTTATCGCACCTACGTTCAATATTACAGGCGAACATACTTACTAAAGTGTGTTA" & _
                          "ATTAATTAATGCTTGTAGGACATAATAATAACAATTGAATGTCTGCACAGCCACTTTCCACACAGACATC" & _
                          "ATAACAAAAAATTTCCACCAAACCCCCCCTCCCCCGCTTCTGGCCACAGCACTTAAACACATCTCTGCCA" & _
                          "AACCCCAAAAACAAAGAACCCTAACACCAGCCTAACCAGATTTCAAATTTTATCTTTTGGCGGTATGCAC" & _
                          "TTTTAACAGTCACCCCCCAACTAACACATTATTTTCCCCTCCCACTCCCATACTACTAATCTCATCAATA"
    End If
    
    
    If Index = 7 Then
        HScroll1.Value = 20
        Sec1.Text = "ATCGATTCGATATCATACACGTAT"
        Sec2.Text = "CTCGACTAGTATGAAGTCCACGCTTG"
        secventata.Text = "ATTAAAGGTTTATACCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAAACGAACTTTAAAAT" & _
                          "CTGTGTGGCTGTCACTCGGCTGCATGCTTAGTGCACTCACGCAGTATAATTAATAACTAATTACTGTCGTTGACAGGACACGA" & _
                          "GTAACTCGTCTATCTTCTGCAGGCTGCTTACGGTTTCGTCCGTGTTGCAGCCGATCATCAGCACATCTAGGTTTCGTCCGGGT" & _
                          "GTGACCGAAAGGTAA"
    End If
    
    
    If Index = 8 Then
        HScroll1.Value = 20
        Sec1.Text = "ATCGATATCGATTCGATATCATACACGTATTCGATAATCGATTCGATATCATCGATTCGATATCAATCGATTCGATATCATACACGTAT" & _
                    "TACACGTATATACACGTATTCATACACGTAT"
        Sec2.Text = "CCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAAACGAACTTTAAAATCTGTGTGGCTGTCACTCGGC" & _
                    "TGCATGCTTAGTGCACTCACGCAGTATAATTAATAACTAATTACTGTCGTTGACAGGACACGAGTAACTCGTCTATCTTCTGC"
        secventata.Text = "ATTAAAGGTTTATACCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAAACGAACTTTAAAAT" & _
                          "CTGTGTGGCTGTCACTCGGCTGCATGCTTAGTGCACTCACGCAGTATAATTAATAACTAATTACTGTCGTTGACAGGACACGA" & _
                          "GTAACTCGTCTATCTTCTGCAGGCTGCTTACGGTTTCGTCCGTGTTGCAGCCGATCATCAGCACATCTAGGTTTCGTCCGGGT" & _
                          "GTGACCGAAAGGTAA"
    End If
    
    
    If Index = 9 Then
        HScroll1.Value = 20
        Sec1.Text = "AAAAATTTTTGGAAGCGCGGCATTGCAGGGGCCCCCC"
        Sec2.Text = "ATCGATCGATCGGATTTTTTTTTCGCTACTAGCGAGCTACTACACAC"
        secventata.Text = "AAAAAAAAATCGCAGCGACGAGCAGCGACGGATTTTTTTTTTTTTTTTTTT"
    End If
    
    
    If Index = 10 Then
        HScroll1.Value = 20
        Sec1.Text = "TATAACGGCATTATTACGCTTATCGCATTCACACTTTAACGTACGATCTACTATCTAGCCGCCAC"
        Sec2.Text = "TATAACGGCATTATTACGCTTATCGCATTCACACTTCCGATTACGCGATCGATAACGTACGATCTACTATCTAGCCGCCAC"
        secventata.Text = "AGCCCTCCAGGACAGGCTGCATCAGAAGAGGCCATCAAGCAGGTCTGTTCCAAGGGCCTTTGCGTCAGGTGGGCTCAGGATTCCAGGGTGGCTGGACCCCA" & _
                          "GGCCCCAGCTCTGCAGCAGGGAGGACGTGGCTGGGCTCGTGAAGCATGTGGGGGTGAGCCCAGGGGCCCCAAGGCAGGGCACCTGGCCTTCAGCCTGCCTC" & _
                          "AGCCCTGCCTGTCTCCCAGATCACTGTCCTTCTGCCATGGCCCTGTGGATGCGCCTCCTGCCCCTGCTGGCGCTGCTGGCCCTCTGGGGACCTGACCCAGC" & _
                          "CGCAGCCTTTGTGAACCAACACCTGTGCGGCTCACACCTGGTGGAAGCTCTCTACCTAGTGTGCGGGGAACGAGGCTTCTTCTACACACCCAAGACCCGCC" & _
                          "GGGAGGCAGAGGACCTGCAGGGTGAGCCAACTGCCCATTGCTGCCCCTGGCCGCCCCCAGCCACCCCCTGCTCCTGGCGCTCCCACCCAGCATGGGCAGAA" & _
                          "GGGGGCAGGAGGCTGCCACCCAGCAGGGGGTCAGGTGCACTTTTTTAAAAAGAAGTTCTCTTGGTCACGTCCTAAAAGTGACCAGCTCCCTGTGGCCCAGT" & _
                          "CAGAATCTCAGCCTGAGGACGGTGTTGGCTTCGGCAGCCCCGAGATACATCAGAGGGTGGGCACGCTCCTCCCTCCACTCGCCCCTCAAACAAATGCCCCG" & _
                          "CAGCCCATTTCTCCACCCTCATTTGATGACCGCAGATTCAAGTGTTTTGTTAAGTAAAGTCCTGGGTGACCTGGGGTCACAGGGTGCCCCACGCTGCCTGC" & _
                          "CTCTGGGCGAACACCCCATCACGCCCGGAGGAGGGCGTGGCTGCCTGCCTGAGTGGGCCAGACCCCTGTCGCCAGGCCTCACGGCAGCTCCATAGTCAGGA" & _
                          "GATGGGGAAGATGCTGGGGACAGGCCCTGGGGAGAAGTACTGGGATCACCTGTTCAGGCTCCCACTGTGACGCTGCCCCGGGGCGGGGGAAGGAGGTGGGA" & _
                          "CATGTGGGCGTTGGGGCCTGTAGGTCCACACCCAGTGTGGGTGACCCTCCCTCTAACCTGGGTCCAGCCCGGCTGGAGATGGGTGGGAGTGCGACCTAGGG" & _
                          "CTGGCGGGCAGGCGGGCACTGTGTCTCCCTGACTGTGTCCTCCTGTGTCCCTCTGCCTCGCCGCTGTTCCGGAACCTGCTCTGCGCGGCACGTCCTGGCAG" & _
                          "TGGGGCAGGTGGAGCTGGGCGGGGGCCCTGGTGCAGGCAGCCTGCAGCCCTTGGCCCTGGAGGGGTCCCTGCAGAAGCGTGGCATTGTGGAACAATGCTGT" & _
                          "ACCAGCATCTGCTCCCTCTACCAGCTGGAGAACTACTGCAACTAGACGCAGCCCGCAGGCAGCCCCACACCCGCCGCCTCCTGCACCGAGAGAGATGGAAT" & _
                          "AAAGCCCTTGAACCAGC"
    End If
    
    
    If Index = 11 Then
        HScroll1.Value = 20
        Sec1.Text = "ATATATATATATATGCGCATATATATATACTGCTATA"
        Sec2.Text = "GAGATACGTGAGAATCGGAGATTAGCGAGAGAGA"
        secventata.Text = "ATATATATATATGAGAGAGAGAGAGAGAGAATATATATATATA"
    End If
    
    
    If Index = 12 Then
        HScroll1.Value = 7
        Sec1.Text = "GGGGAAGATGCTGGGGACAGGCCCTGGGGAGAAGTACTGGGATCACCTGTTCAGGCTCCCACTGTGACGCTGCCCCGGGGCGGGGGAAGGAGGTGGG"
        Sec2.Text = "ATTAAAGGTTTATACCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAAACGAACTTTAAAAT"
        secventata.Text = "ATTTATTTATTTAATGCCGGGCATTCGGCGAGAGGCATGGGAATATATTATATATAA"
    End If
    
    
    If Index = 13 Then
        HScroll1.Value = 20
        Sec1.Text = "GGGGAAGATGCTGGGGACAGGCCCTGGGGAGAAGTACTGGGATCACCTGTTCAGGCTCCCACTGTGACGCTGCCCCGGGGCGGGGGAAGGAGGTGGG"
        Sec2.Text = "ATTAAAGGTTTATACCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAAACGAACTTTAAAAT"
        secventata.Text = "ATTTATTTATTTAATGCCGGGCATTCGGCGAGAGGCATGGGAATATATTATATATAA"
    End If
    
    
    If Index = 14 Then
        HScroll1.Value = 20
        Sec1.Text = "GGGAAAGGGAACGGAACCTGCTCTGAGGGAAAGGGAAAGGGAAA"
        Sec2.Text = "TTTAAATTTAAGGCACTGTGTCTCCCTGACTGTGATTTAAATTTAAA"
        secventata.Text = "AGCCCTCCAGGACAGGCTGCATCAGAAGAGGCCATCAAGCAGGTCTGTTCCAAGGGCCTTTGCGTCAGGTGGGCTCAGGATTCCAGGGTGGCTGGACCCCA" & _
                          "GGCCCCAGCTCTGCAGCAGGGAGGACGTGGCTGGGCTCGTGAAGCATGTGGGGGTGAGCCCAGGGGCCCCAAGGCAGGGCACCTGGCCTTCAGCCTGCCTC" & _
                          "AGCCCTGCCTGTCTCCCAGATCACTGTCCTTCTGCCATGGCCCTGTGGATGCGCCTCCTGCCCCTGCTGGCGCTGCTGGCCCTCTGGGGACCTGACCCAGC" & _
                          "CGCAGCCTTTGTGAACCAACACCTGTGCGGCTCACACCTGGTGGAAGCTCTCTACCTAGTGTGCGGGGAACGAGGCTTCTTCTACACACCCAAGACCCGCC" & _
                          "GGGAGGCAGAGGACCTGCAGGGTGAGCCAACTGCCCATTGCTGCCCCTGGCCGCCCCCAGCCACCCCCTGCTCCTGGCGCTCCCACCCAGCATGGGCAGAA" & _
                          "GGGGGCAGGAGGCTGCCACCCAGCAGGGGGTCAGGTGCACTTTTTTAAAAAGAAGTTCTCTTGGTCACGTCCTAAAAGTGACCAGCTCCCTGTGGCCCAGT" & _
                          "CAGAATCTCAGCCTGAGGACGGTGTTGGCTTCGGCAGCCCCGAGATACATCAGAGGGTGGGCACGCTCCTCCCTCCACTCGCCCCTCAAACAAATGCCCCG" & _
                          "CAGCCCATTTCTCCACCCTCATTTGATGACCGCAGATTCAAGTGTTTTGTTAAGTAAAGTCCTGGGTGACCTGGGGTCACAGGGTGCCCCACGCTGCCTGC" & _
                          "CTCTGGGCGAACACCCCATCACGCCCGGAGGAGGGCGTGGCTGCCTGCCTGAGTGGGCCAGACCCCTGTCGCCAGGCCTCACGGCAGCTCCATAGTCAGGA" & _
                          "GATGGGGAAGATGCTGGGGACAGGCCCTGGGGAGAAGTACTGGGATCACCTGTTCAGGCTCCCACTGTGACGCTGCCCCGGGGCGGGGGAAGGAGGTGGGA" & _
                          "CATGTGGGCGTTGGGGCCTGTAGGTCCACACCCAGTGTGGGTGACCCTCCCTCTAACCTGGGTCCAGCCCGGCTGGAGATGGGTGGGAGTGCGACCTAGGG" & _
                          "CTGGCGGGCAGGCGGGCACTGTGTCTCCCTGACTGTGTCCTCCTGTGTCCCTCTGCCTCGCCGCTGTTCCGGAACCTGCTCTGCGCGGCACGTCCTGGCAG" & _
                          "TGGGGCAGGTGGAGCTGGGCGGGGGCCCTGGTGCAGGCAGCCTGCAGCCCTTGGCCCTGGAGGGGTCCCTGCAGAAGCGTGGCATTGTGGAACAATGCTGT" & _
                          "ACCAGCATCTGCTCCCTCTACCAGCTGGAGAACTACTGCAACTAGACGCAGCCCGCAGGCAGCCCCACACCCGCCGCCTCCTGCACCGAGAGAGATGGAAT" & _
                          "AAAGCCCTTGAACCAGC"
    End If
    
    
    If Index = 15 Then
        HScroll1.Value = 20
        Sec1.Text = "AGCCGCCGCGTCGACGCGAACGTCGTACG"
        Sec2.Text = "TATAACGGCATTATTACGCTTATCGCATTCACACTTTA"
        secventata.Text = "AGCCCTCCAGGACAGGCTGCATCAGAAGAGGCCATCAAGCAGGTCTGTTCCAAGGGCCTTTGCGTCAGGTGGGCTCAGGATTCCAGGGTGGCTGGACCCCA" & _
                          "GGCCCCAGCTCTGCAGCAGGGAGGACGTGGCTGGGCTCGTGAAGCATGTGGGGGTGAGCCCAGGGGCCCCAAGGCAGGGCACCTGGCCTTCAGCCTGCCTC" & _
                          "AGCCCTGCCTGTCTCCCAGATCACTGTCCTTCTGCCATGGCCCTGTGGATGCGCCTCCTGCCCCTGCTGGCGCTGCTGGCCCTCTGGGGACCTGACCCAGC" & _
                          "CGCAGCCTTTGTGAACCAACACCTGTGCGGCTCACACCTGGTGGAAGCTCTCTACCTAGTGTGCGGGGAACGAGGCTTCTTCTACACACCCAAGACCCGCC" & _
                          "GGGAGGCAGAGGACCTGCAGGGTGAGCCAACTGCCCATTGCTGCCCCTGGCCGCCCCCAGCCACCCCCTGCTCCTGGCGCTCCCACCCAGCATGGGCAGAA" & _
                          "GGGGGCAGGAGGCTGCCACCCAGCAGGGGGTCAGGTGCACTTTTTTAAAAAGAAGTTCTCTTGGTCACGTCCTAAAAGTGACCAGCTCCCTGTGGCCCAGT" & _
                          "CAGAATCTCAGCCTGAGGACGGTGTTGGCTTCGGCAGCCCCGAGATACATCAGAGGGTGGGCACGCTCCTCCCTCCACTCGCCCCTCAAACAAATGCCCCG" & _
                          "CAGCCCATTTCTCCACCCTCATTTGATGACCGCAGATTCAAGTGTTTTGTTAAGTAAAGTCCTGGGTGACCTGGGGTCACAGGGTGCCCCACGCTGCCTGC" & _
                          "CTCTGGGCGAACACCCCATCACGCCCGGAGGAGGGCGTGGCTGCCTGCCTGAGTGGGCCAGACCCCTGTCGCCAGGCCTCACGGCAGCTCCATAGTCAGGA" & _
                          "GATGGGGAAGATGCTGGGGACAGGCCCTGGGGAGAAGTACTGGGATCACCTGTTCAGGCTCCCACTGTGACGCTGCCCCGGGGCGGGGGAAGGAGGTGGGA" & _
                          "CATGTGGGCGTTGGGGCCTGTAGGTCCACACCCAGTGTGGGTGACCCTCCCTCTAACCTGGGTCCAGCCCGGCTGGAGATGGGTGGGAGTGCGACCTAGGG" & _
                          "CTGGCGGGCAGGCGGGCACTGTGTCTCCCTGACTGTGTCCTCCTGTGTCCCTCTGCCTCGCCGCTGTTCCGGAACCTGCTCTGCGCGGCACGTCCTGGCAG" & _
                          "TGGGGCAGGTGGAGCTGGGCGGGGGCCCTGGTGCAGGCAGCCTGCAGCCCTTGGCCCTGGAGGGGTCCCTGCAGAAGCGTGGCATTGTGGAACAATGCTGT" & _
                          "ACCAGCATCTGCTCCCTCTACCAGCTGGAGAACTACTGCAACTAGACGCAGCCCGCAGGCAGCCCCACACCCGCCGCCTCCTGCACCGAGAGAGATGGAAT" & _
                          "AAAGCCCTTGAACCAGC"
    End If
    
    
    If Index = 16 Then
        HScroll1.Value = 9
        Sec1.Text = "CGGGTCGGAGTTAGCTCAAGCGGTTACCTCCTCATGCCGGACTTTCTATCTGTCCATCTCTGTGCTGGGGTTCGAGACCCGCGGGTGCTTACTGACCCTTTTATGCAA"
        Sec2.Text = "TATATATTTATTTTTAAAATTATTATTATTATATATCGTACAGC"
        secventata.Text = "ATTTATTTATTTAATGCCGGGCATTCGGCGAGAGCGCGATGCGCACGTTCCCGGCGGGAATATATTATATATAATATATATTTATTTTTAAAATTATTATT" & _
                          "ATTATATCGGGTCGGAGTTAGCTCAAGCGGTTACCTCCTCATGCCGGACTTTCTATCTGTCCATCTCTGTGCTGGGGTTCGAGACCCGCGGGTGCTTACTG" & _
                          "ACCCTTTTATGCAA"
    End If
    
    
    If Index = 17 Then
        HScroll1.Value = 20
        Sec1.Text = "GTTTATACCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAA"
        Sec2.Text = "TTTCGTCCGTGTTGCAGCCGATCATCAGCACATCTAGGTTTCGTCCG"
        secventata.Text = "ATTAAAGGTTTATACCTTCCCAGGTAACAAACCAACCAACTTTCGATCTCTTGTAGATCTGTTCTCTAAACGAACTTTAAAAT" & _
                          "CTGTGTGGCTGTCACTCGGCTGCATGCTTAGTGCACTCACGCAGTATAATTAATAACTAATTACTGTCGTTGACAGGACACGA" & _
                          "GTAACTCGTCTATCTTCTGCAGGCTGCTTACGGTTTCGTCCGTGTTGCAGCCGATCATCAGCACATCTAGGTTTCGTCCGGGT" & _
                          "GTGACCGAAAGGTAA"
    End If
    

    lim_real_time = 200
    Scan_Click
End Sub
