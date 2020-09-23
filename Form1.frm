VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "MatrixFall"
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "http://www.open-design.be"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":34CA
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    '
    Select Case KeyAscii
        '
        'Case vbKeyEscape
            '
             'bRunning = False
            '
        Case 102
            '
            AffFps = Not AffFps
            '
        Case 112
            '
            If PauseSz = False Then
                '
                PauseSz = True
                '
            Else
                '
                'on remet les compteurs à zéro
                FPS_NbrImg = 0
                lFpsTmp = GetTickCount - 1
                '
                PauseSz = False
                '
            End If
            '
        Case Else
            '
            bRunning = False
            '
        'Case vbKeyN
            '
            
            '
        'Case vbKeyZ
            '
            'CamDistance = CamDistance + 10
            '
        'Case vbKeyS
            '
            'CamDistance = CamDistance - 10
            '
        '
    End Select
    '
End Sub
'
'lancement de la procédure principale
Public Sub LancerProcP()
    '
    'on initialise quelques variables
    PauseSz = False
    '
    'on lance le programme
    Initialise Me, ModeAffSzX, ModeAffSzY
    '
    Unload Me
    '
End Sub

Private Sub Form_Load()
    '
    Me.WindowState = 0
    '
    'Me.Top = 0
    'Me.Left = 0
    'Me.Width = Screen.Width
    'Me.Height = Screen.Height
    '
End Sub

Private Sub Form_LostFocus()
    '
    bRunning = False
    '
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    bRunning = False
    '
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    'bRunning = False
    '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
    '
    bRunning = False
    '
End Sub
