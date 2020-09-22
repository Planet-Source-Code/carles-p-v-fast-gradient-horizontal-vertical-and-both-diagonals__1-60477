VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mGradient test"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPaint 
      Caption         =   "&Paint"
      Default         =   -1  'True
      Height          =   495
      Left            =   6570
      TabIndex        =   2
      Top             =   585
      Width           =   1605
   End
   Begin VB.TextBox txtIterations 
      Height          =   315
      Left            =   7515
      TabIndex        =   1
      Text            =   "1"
      Top             =   150
      Width           =   660
   End
   Begin VB.Label lblIterations 
      Caption         =   "Iterations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6555
      TabIndex        =   0
      Top             =   195
      Width           =   1020
   End
   Begin VB.Label lblTiming 
      Height          =   3105
      Left            =   6570
      TabIndex        =   3
      Top             =   1335
      Width           =   1605
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oTiming As cTiming

Private Sub Form_Load()

    If (App.LogMode <> 1) Then
        Call MsgBox("Absolutely recommended: compile first...")
    End If
    
    Set Me.Icon = Nothing
    Set m_oTiming = New cTiming
End Sub

Private Sub cmdPaint_Click()
  
  Dim clr1 As Long
  Dim clr2 As Long
  Dim i    As Long
  Dim it   As Long
    
    '-- Check iterations
    With txtIterations
        If (Not IsNumeric(.Text)) Then
            Call MsgBox("Please, enter a valid 'Iterations' number")
            Call .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
        it = Val(.Text)
    End With
    
    '-- Random colors
    Call Randomize(Timer)
    clr1 = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    clr2 = clr1 Xor vbWhite
    
    lblTiming = vbNullString
    
    '-- Test horizontal
    Call m_oTiming.Reset
    For i = 1 To it
        Call mGradient.PaintGradient(Me.hDC, 10, 10, 400, 100, clr1, clr2, [gdHorizontal])
    Next i
    lblTiming = lblTiming & "Horizontal:" & vbCrLf & Format$(m_oTiming.Elapsed / 1000, "0.0000 s") & vbCrLf & vbCrLf
    Call lblTiming.Refresh
    
    '-- Test vertical
    Call m_oTiming.Reset
    For i = 1 To it
        Call mGradient.PaintGradient(Me.hDC, 10, 120, 400, 100, clr1, clr2, [gdVertical])
    Next i
    lblTiming = lblTiming & "Vertical:" & vbCrLf & Format$(m_oTiming.Elapsed / 1000, "0.0000 s") & vbCrLf & vbCrLf
    Call lblTiming.Refresh
    
    '-- Test downward diagonal
    Call m_oTiming.Reset
    For i = 1 To it
        Call mGradient.PaintGradient(Me.hDC, 10, 230, 400, 100, clr1, clr2, [gdDownwardDiagonal])
    Next i
    lblTiming = lblTiming & "Downward diagonal:" & vbCrLf & Format$(m_oTiming.Elapsed / 1000, "0.0000 s") & vbCrLf & vbCrLf
    Call lblTiming.Refresh
    
    '-- Test upward diagonal
    Call m_oTiming.Reset
    For i = 1 To it
        Call mGradient.PaintGradient(Me.hDC, 10, 340, 400, 100, clr1, clr2, [gdUpwardDiagonal])
    Next i
    lblTiming = lblTiming & "Upward diagonal:" & vbCrLf & Format$(m_oTiming.Elapsed / 1000, "0.0000 s")
    Call lblTiming.Refresh
End Sub

