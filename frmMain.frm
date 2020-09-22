VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "The ENIGMA PRO Encryption & Decryption Machine"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExtra 
      Caption         =   "extra encryption - remove spaces"
      Height          =   315
      Left            =   5400
      TabIndex        =   9
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   7
      Top             =   3825
      Width           =   3930
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2940
      TabIndex        =   6
      Top             =   3330
      Width           =   2730
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   5
      Top             =   3330
      Width           =   2730
   End
   Begin VB.CommandButton btnAbout 
      Caption         =   "About Enigma Machine"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5745
      TabIndex        =   4
      Top             =   3315
      Width           =   2730
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   870
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1665
      Width           =   7590
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   870
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   105
      Width           =   7590
   End
   Begin VB.Label lblExtra 
      AutoSize        =   -1  'True
      Caption         =   "removing spaces from encoded message makes it harder to read, but also almost impossible to break"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2025
      TabIndex        =   10
      Top             =   4260
      Width           =   6270
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   3885
      Width           =   1005
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "result:"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "source:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The Enigma PRO Encryption Machine

' You can find description and other
'    technical information about ENIGMA machine here:
'
'  http://www.emu8086.com/vb/enigma.html

' This code is free to use for any purpose,
' you should keep the original copyright:

' Copyright (C) 2005 Free Code
' http://www.emu8086.com/vb/
' info@emu8086.com

' The same expression is can be encrypted
' differently, thus it is very hard to break down this code!

' This is PRO version of ENIGMA machine,
' to understand how ENIGMA machine works
' you can download less complicated version from
' my home page (it has visual wheels).

' If you would like to use ENIGMA encryption in your
' application, you may use Encrypt() and Decrypt()
' functions in "mEnigmaPRO.bas".

' Functions in "mEnigmaPRO.bas" are independent from
' code in this form.

' This Internet URL has some infomation
' about ENIGMA machine (some history...)
' http://www.emu8086.com/vb/enigma.html


Private Sub cmdEncrypt_Click()
On Error Resume Next
    txtResult.Text = Encrypt_PRO(txtSource.Text, txtPassword.Text, IIf(chkExtra.Value = 1, True, False))
End Sub

Private Sub cmdDecrypt_Click()
On Error Resume Next
    txtResult.Text = Decrypt_PRO(txtSource.Text, txtPassword.Text)
End Sub

' Open Help Window:
Private Sub btnAbout_Click()

On Error Resume Next

    Dim sURL As String
    sURL = "http://www.emu8086.com/vb/enigma.html"
    Shell "explorer " & sURL, vbNormalFocus
    
    
End Sub

Private Sub Form_Resize()

On Error GoTo err1


    lblExtra.Top = Me.ScaleHeight - lblExtra.Height - 50

    txtPassword.Top = lblExtra.Top - txtPassword.Height - 50
    lblPassword.Top = txtPassword.Top + 40
    
    cmdEncrypt.Top = txtPassword.Top - cmdEncrypt.Height - 50
    cmdDecrypt.Top = cmdEncrypt.Top
    btnAbout.Top = cmdEncrypt.Top
    
    chkExtra.Top = btnAbout.Top + btnAbout.Height + 20
    
    txtSource.Width = Me.ScaleWidth - txtSource.Left - 50
    txtResult.Width = txtSource.Width
    
    
    
    '' a bit more confusing :)
    Dim xJJ1 As Single
    
    xJJ1 = Me.ScaleHeight - (Me.ScaleHeight - cmdEncrypt.Top)
    xJJ1 = xJJ1 / 2 - 200
    
    txtSource.Height = xJJ1 - 100
    
    txtResult.Top = txtSource.Top + txtSource.Height + 100
    txtResult.Height = txtSource.Height
    
    lblResult.Top = txtResult.Top
    
    Exit Sub
err1:
    Debug.Print "resize: " & Err.Description
End Sub

Private Sub txtResult_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 65 And Shift = 2 Then
        txtResult.SelStart = 0
        txtResult.SelLength = Len(txtResult.Text)
    End If
End Sub


Private Sub txtSource_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 65 And Shift = 2 Then
        txtSource.SelStart = 0
        txtSource.SelLength = Len(txtSource.Text)
    End If
End Sub
