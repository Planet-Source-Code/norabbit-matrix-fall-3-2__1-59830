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
'�tat de la classe
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
        Case 1 'on affiche les lignes petit � petit
            '
            'on attend le temps n�cessaire, ensuite on effectue l'op�ration en fonction de la demande
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                tTps1 = Rnd() * 300
                Tps1 = 0
                '
                'on incr�mente la variable contenant le nombre de lignes charg�es
                LimiteLignesChargeSz = LimiteLignesChargeSz + 1
                '
                'on charge une ligne suppl�mentaire
                ReDim Preserve CLgn3D(0 To LimiteLignesChargeSz)
                '
                Set CLgn3D(LimiteLignesChargeSz) = New ClsLigne3D1
                '
                'on v�rifie si on a charg� toutes les lignes
                If LimiteLignesAffSz = LimiteLignesChargeSz + 1 Then
                    '
                    'on change l'�tat de la classe
                    Etat = 2
                    '
                End If
                '
            End If
            '
        '
    End Select
    '
    'on ex�cute la liste des commandes en tenant compte de leur pose
    '
    'on attend le temps n�cessaire, ensuite on effectue l'op�ration en fonction de la demande
    Tps2 = Tps2 + (1 * VitesseGenSz)
    '
    'on v�rifie qu'on ne va pas trop loint
    If Tps2 >= tTps2 Then
        '
        'on ex�cute la commande
        Select Case CmdSz.Item(CourCmdSz)
            '
            Case "txt" 'titre #2
                '
                'on r�cup�re les 2 titres
                cTmp() = Split(ValCmdSz.Item(CourCmdSz), "_")
                '
                If UBound(cTmp) >= 0 Then
                    '
                    'on ajoute le texte � la liste #1
                    TxtTitre1.Add cTmp(0)
                    '
                    'on ajoute le texte � la liste #2
                    TxtTitre2.Add cTmp(1)
                    '
                Else
                    '
                    'on ajoute un texte vide au titre #1
                    TxtTitre1.Add ""
                    '
                    'on ajoute le texte � la liste #2
                    TxtTitre2.Add cTmp(0)
                    '
                End If
                'on lance la classe si son �tat=0
                If cTitre2.Etat = 0 Then cTitre2.InitTitre2
                '
            '
        End Select
        '
        'on incr�mente le pointeur
        CourCmdSz = CourCmdSz + 1
        '
        'on v�rifie qu'il n'est pas trop grand sinon on le replace au d�but
        If CourCmdSz > CmdSz.Count Then CourCmdSz = 1
        '
        'on change les valeurs de la transition
        tTps2 = PauseCmdSz.Item(CourCmdSz)
        Tps2 = 0
        '
    End If
    '
End Sub

