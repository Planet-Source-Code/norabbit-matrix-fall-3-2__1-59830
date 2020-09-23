VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEvenements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE LES EVENEMENTS                                               *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'état de la classe
Public Etat As Long
'
'transitions
Public Tps1 As Long
Public tTps1 As Long
Public Tps2 As Long
Public tTps2 As Long
'
'tableau de chaines
Dim cTmp() As String
'
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    Etat = 1
    '
    CourCmdSz = 1
    '
End Sub
'
'ON EFFECTUE LES CHANGEMENTS
Public Sub Calcul()
    '
    If Etat = 0 Then Exit Sub
    '
    Select Case Etat
        '
        Case 1 'on affiche les lignes petit à petit
            '
            'on attend le temps nécessaire, ensuite on effectue l'opération en fonction de la demande
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on vérifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                tTps1 = Rnd() * 300
                Tps1 = 0
                '
                'on incrémente la variable contenant le nombre de lignes chargées
                LimiteLignesChargeSz = LimiteLignesChargeSz + 1
                '
                'on charge une ligne supplémentaire
                ReDim Preserve CLgn3D(0 To LimiteLignesChargeSz)
                '
                Set CLgn3D(LimiteLignesChargeSz) = New ClsLigne3D1
                '
                'on vérifie si on a chargé toutes les lignes
                If LimiteLignesAffSz = LimiteLignesChargeSz + 1 Then
                    '
                    'on change l'état de la classe
                    Etat = 2
                    '
                End If
                '
            End If
            '
        '
    End Select
    '
    'on exécute la liste des commandes en tenant compte de leur pose
    '
    'on attend le temps nécessaire, ensuite on effectue l'opération en fonction de la demande
    Tps2 = Tps2 + (1 * VitesseGenSz)
    '
    'on vérifie qu'on ne va pas trop loint
    If Tps2 >= tTps2 Then
        '
        'on exécute la commande
        Select Case CmdSz.Item(CourCmdSz)
            '
            Case "txt" 'titre #2
                '
                'on récupère les 2 titres
                cTmp() = Split(ValCmdSz.Item(CourCmdSz), "_")
                '
                If UBound(cTmp) >= 0 Then
                    '
                    'on ajoute le texte à la liste #1
                    TxtTitre1.Add cTmp(0)
                    '
                    'on ajoute le texte à la liste #2
                    TxtTitre2.Add cTmp(1)
                    '
                Else
                    '
                    'on ajoute un texte vide au titre #1
                    TxtTitre1.Add ""
                    '
                    'on ajoute le texte à la liste #2
                    TxtTitre2.Add cTmp(0)
                    '
                End If
                'on lance la classe si son état=0
                If cTitre2.Etat = 0 Then cTitre2.InitTitre2
                '
            '
        End Select
        '
        'on incrémente le pointeur
        CourCmdSz = CourCmdSz + 1
        '
        'on vérifie qu'il n'est pas trop grand sinon on le replace au début
        If CourCmdSz > CmdSz.Count Then CourCmdSz = 1
        '
        'on change les valeurs de la transition
        tTps2 = PauseCmdSz.Item(CourCmdSz)
        Tps2 = 0
        '
    End If
    '
End Sub

