VERSION 5.00
Begin VB.Form frmBMI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Body Mass Index"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Convertor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   6600
      TabIndex        =   12
      Top             =   240
      Width           =   3015
      Begin VB.TextBox txtHInches 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   21
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtHMtr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtHPound 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtHKgs 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00F1F1F1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00F1F1F1&
         Caption         =   "Inches "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00F1F1F1&
         Caption         =   "Mtr"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00F1F1F1&
         Caption         =   "1.0 Meter = 39.34 Inches"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00F1F1F1&
         Caption         =   "Pound "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00F1F1F1&
         Caption         =   "Kgs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00F1F1F1&
         Caption         =   "1.0 Pound = 0.45 KG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5775
      Begin VB.Frame Frame1 
         BackColor       =   &H00F1F1F1&
         Caption         =   "Enter Your Weight in Kgs or Pounds"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   5295
         Begin VB.OptionButton opnKgs 
            BackColor       =   &H00F1F1F1&
            Caption         =   "  Kgs"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton opnPounds 
            BackColor       =   &H00F1F1F1&
            Caption         =   "  Pounds"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtWeight 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   0
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F1F1F1&
         Caption         =   "Enter Your Height in Mtr or Inches"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   5295
         Begin VB.TextBox txtHeight 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   1
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton opnInches 
            BackColor       =   &H00F1F1F1&
            Caption         =   " Inches"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton opnMtr 
            BackColor       =   &H00F1F1F1&
            Caption         =   "  Mtr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Label lblResult 
         BackColor       =   &H00F1F1F1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Too Much over weight, wake up, exercise & be on diet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   5295
      End
      Begin VB.Label lblBMI 
         BackColor       =   &H00F1F1F1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Your Body Mass Index (BMI) = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   5295
      End
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1F1F1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                                     H  E  L  P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   6120
      TabIndex        =   11
      Top             =   240
      Width           =   255
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Indx As Integer

Private Sub Form_Load()
On Error Resume Next
Me.Width = 6630
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

opnKgs.Value = True
opnMtr.Value = True
'txtWeight.SetFocus
lblBMI = ""
lblResult = ""
lblMsg = ""
ClearAll
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lblHelp.BackColor = &HF1F1F1
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lblHelp.BackColor = &HF1F1F1
End Sub

Private Sub lblHelp_Click()
On Error Resume Next
If Me.Width = 6630 Then
    Me.Width = 9930
Else
    Me.Width = 6630
End If
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lblHelp.BackColor = &H80FF&
End Sub

Private Sub opnInches_Click()
On Error Resume Next
txtHeight = ""
txtHeight.SetFocus
lblBMI = "Your Body Mass Index ( BMI ) ="
lblResult = "Result"
End Sub

Private Sub opnKgs_Click()
On Error Resume Next
txtWeight = ""
txtWeight.SetFocus
lblBMI = "Your Body Mass Index ( BMI ) ="
lblResult = "Result"
End Sub

Private Sub opnMtr_Click()
On Error Resume Next
txtHeight = ""
txtHeight.SetFocus
lblBMI = "Your Body Mass Index ( BMI ) ="
lblResult = "Result"
End Sub

Private Sub opnPounds_Click()
On Error Resume Next
txtWeight = ""
txtWeight.SetFocus
lblBMI = "Your Body Mass Index ( BMI ) ="
lblResult = "Result"
End Sub

Private Sub txtHeight_Change()
On Error Resume Next
If txtWeight = "" Then Exit Sub
If txtHeight = "" Then Exit Sub
lblBMI = ""
Index
lblBMI = "Your Body Mass Index ( BMI ) =" & Space(2) & Format(Indx, "0")
If Indx = 0 Then
    Rslt
    'Exit Sub
Else
    Rslt
End If
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
On Error Resume Next
''' number validation
    If KeyAscii = 13 Then SendKeys vbTab

    If KeyAscii = 8 Then Exit Sub
    
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
        If KeyAscii <> 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtHInches_KeyPress(KeyAscii As Integer)
On Error Resume Next
''' number validation
    'If KeyAscii = 13 Then SendKeys vbTab

    If KeyAscii = 8 Then Exit Sub
    
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
        If KeyAscii <> 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtHInches_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If txtHInches = "" Then
    txtHMtr = ""
    lblMsg = ""
    Exit Sub
ElseIf txtHInches < 39.34 Then
    txtHMtr = ""
    'txtHMtr = (1 * txtHInches) / 39.34
    lblMsg = "    Minimum Limit 39.34  Maximum No Limit"
Else
    txtHMtr = ""
    lblMsg = ""
    txtHMtr = Format((1 * txtHInches) / 39.34, "0.00")
End If
End Sub

Private Sub txtHKgs_KeyPress(KeyAscii As Integer)
On Error Resume Next
''' number validation
    If KeyAscii = 13 Then SendKeys vbTab

    If KeyAscii = 8 Then Exit Sub
    
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
        If KeyAscii <> 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtHKgs_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If txtHKgs = "" Then
    txtHPound = ""
    lblMsg = ""
    Exit Sub
Else
    txtHPound = ""
    lblMsg = ""
    txtHPound = Format((txtHKgs * 1) / 0.45, "0.00")
End If
End Sub

Private Sub txtHMtr_KeyPress(KeyAscii As Integer)
On Error Resume Next
''' number validation
    'If KeyAscii = 13 Then SendKeys vbTab

    If KeyAscii = 8 Then Exit Sub
    
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
        If KeyAscii <> 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtHMtr_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If txtHMtr = "" Then
    txtHInches = ""
    lblMsg = ""
    Exit Sub
Else
    txtHInches = ""
    lblMsg = ""
    txtHInches = Format((txtHMtr * 39.34) / 1, "0.00")
End If
End Sub

Private Sub txtHPound_KeyPress(KeyAscii As Integer)
On Error Resume Next
''' number validation
    'If KeyAscii = 13 Then SendKeys vbTab

    If KeyAscii = 8 Then Exit Sub
    
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
        If KeyAscii <> 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtHPound_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If txtHPound = "" Then
    txtHKgs = ""
    lblMsg = ""
    Exit Sub
Else
    txtHKgs = ""
    lblMsg = ""
    txtHKgs = Format((txtHPound * 0.45) / 1, "0.00")
End If
End Sub

Private Sub txtWeight_Change()
On Error Resume Next
If txtWeight = "" Then Exit Sub
If txtHeight = "" Then Exit Sub
lblBMI = ""
Index
lblBMI = "Your Body Mass Index ( BMI ) =" & Space(2) & Format(Indx, "0")
If Indx = 0 Then
    Exit Sub
Else
    Rslt
End If
End Sub

Private Sub txtWeight_KeyPress(KeyAscii As Integer)
On Error Resume Next
''' number validation
    If KeyAscii = 13 Then SendKeys vbTab

    If KeyAscii = 8 Then Exit Sub
    
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
        If KeyAscii <> 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub ClearAll()
On Error Resume Next
lblBMI = "Your Body Mass Index ( BMI )="
lblResult = "Result"
End Sub

Private Sub Index()
On Error Resume Next
If opnKgs.Value = True And opnInches.Value = True Then
    mtr = (1 * Val(txtHeight)) / 39.34
    Indx = Val(txtWeight) / (mtr * mtr)
ElseIf opnKgs.Value = True And opnMtr.Value = True Then
    Indx = Val(txtWeight) / (Val(txtHeight) * Val(txtHeight))
ElseIf opnPounds.Value = True And opnInches.Value = True Then
    kgs = (txtWeight * 0.45) / 1
    mtr = (1 * Val(txtHeight)) / 39.34
    Indx = Val(kgs) / (mtr * mtr)
ElseIf opnPounds.Value = True And opnMtr.Value = True Then
    kgs = (txtWeight * 0.45) / 1
    Indx = kgs / (Val(txtHeight) * Val(txtHeight))
Else
    Indx = 0
End If
End Sub

Private Sub Rslt()
On Error Resume Next
If opnKgs.Value = True Then
    hh = opnKgs.Caption
ElseIf opnPounds.Value = True Then
    hh = opnPounds.Caption
Else
    hh = ""
End If
lblResult = ""
If Indx <= 15 Then
    lblResult = "Result =" & Space(2) & "Too much under-weight, try to gain" & hh
ElseIf Indx = 16 Then
    lblResult = "Result =" & Space(2) & "Too much under-weight, try to gain" & hh
ElseIf Indx = 17 Then
    lblResult = "Result =" & Space(2) & "Under-weight, try to gain" & hh
ElseIf Indx = 18 Then
    lblResult = "Result =" & Space(2) & "Under-weight, try to gain" & hh
ElseIf Indx = 19 Then
    lblResult = "Result =" & Space(2) & "Under-weight, try to gain" & hh
ElseIf Indx = 20 Then
    lblResult = "Result =" & Space(2) & "Slim, Maintain your body, do exercise."
ElseIf Indx = 21 Then
    lblResult = "Result =" & Space(2) & "Good physique, keep it up."
ElseIf Indx = 22 Then
    lblResult = "Result =" & Space(2) & "Good physique, keep it up."
ElseIf Indx = 23 Then
    lblResult = "Result =" & Space(2) & "Medically fit, do not gain more" & hh
ElseIf Indx = 24 Then
    lblResult = "Result =" & Space(2) & "Slightly over-weight, try to reduce few" & hh
ElseIf Indx = 25 Then
    lblResult = "Result =" & Space(2) & "Over-weight, Work up your body."
ElseIf Indx = 26 Then
    lblResult = "Result =" & Space(2) & "Too much Over-weight, Wack up, Exercise & be on diet."
ElseIf Indx = 27 Then
    lblResult = "Result =" & Space(2) & "Too much Over-weight, Wack up, Exercise & be on diet."
ElseIf Indx = 28 Then
    lblResult = "Result =" & Space(2) & "Danger, Extremely Over-Weight : Do something NOW!"
ElseIf Indx >= 29 Then
    lblResult = "Result =" & Space(2) & "Most Dangerous, Extreme Over-Weight"
ElseIf Indx = 0 Then
    lblResult = "Result"
Else
    lblResult = "Result"
End If
End Sub
