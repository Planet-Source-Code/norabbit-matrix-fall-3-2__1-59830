VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLigne3D1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LIGNE DE PLUSIEURS LETTRES                         *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classe gérant une lettre
Dim ClsL() As ClsLettre3D1
'
'nombre de caractères chargées
Public NbrCarSz As Integer
'
'nombre de caractères nécessaires
Public NbrCarNSz As Long
'
'nombre de caractères ayant écoulé leur durée de vie
Public NbrCarFin As Long
'
'opacité (en %) des caractères de la ligne en fonction de la distance
Public opDistSz As Long
'
'coordonnées
Public XSz As Single
Public YSz As Single
Public ZSz As Single
'Dim YTmp As Single
'
'transition
Public Tps1 As Long
Public tTps1 As Long
'
'etat de la ligne
Public Etat As Single '1 = active ; 0 = inactive
'
'nous prévient si la création de lettres est terminée
'Dim OkFinSz As Boolean
'
'index de cette instance
'Dim IndexTmpSz As Integer
'
'variables quelconques
Dim i As Integer
Dim prctSz As Single
Dim xC As Single
Dim zC As Single
'
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    'on initialise certaines variables
    ReDim ClsL(0 To 0)
    '
    NbrCarNSz = 30 '60 '+ Rnd() * 200
    tTps1 = NbrCarNSz * (5 + Rnd() * 10)
    YSz = -500 + Rnd() * 250
    XSz = CamVX + 1000 - Rnd() * 2000
    'ZSz = CamVZ + 1000 - Rnd() * 900
    ZSz = CamTZ - (Rnd() * (CamDistance - 1000))
    Etat = 1
    '
End Sub
'
'AJOUT DE LETTRE + CALCUL DE LEUR POSITION, OPACITE,...
Public Sub Calculer()
    '
    '
    'si l'état = 0, alors on en a fini avec cette ligne
    'If Etat = 0 Then Exit Sub
    '
    'on vérifie si on a déjà créé toutes les lettre, si oui : on passe directement au calcul des caractères
    If NbrCarNSz = NbrCarSz Then GoTo CALCUL_PROC
    '
    'on incrémente Tps1 en fonction de la vitesse générale
    Tps1 = Tps1 + (1 * VitesseGenSz)
    '
    'on vérifie qu'on ne va pas trop loint
    If Tps1 >= tTps1 Then
        '
        prctSz = 100
        '
    Else
        '
        'on calcule le pourcentage
        prctSz = (Tps1 / tTps1) * 100
        '
    End If
    '
    'dMsg = prctSz
    '
    'on calcule le nombre de nouveaux caractères à afficher en fonction du pourcentage
    NbrSz = ((NbrCarNSz / 100) * prctSz) - NbrCarSz
    '
    'dMsg = NbrSz
    '
    'on crée chaque caractère nécessaire
    For i = 1 To NbrSz
        '
        'on incrémente le nombre de lettres créées
        NbrCarSz = NbrCarSz + 1
        '
        'on crée une nouvelle lettre
        ReDim Preserve ClsL(0 To NbrCarSz)
        '
        Set ClsL(NbrCarSz) = New ClsLettre3D1
        '
        With ClsL(NbrCarSz)
            '
            .XSz = XSz
            .YSz = YSz
            .ZSz = ZSz
            '
        End With
        '
        'on incrément la position yZ
        YSz = YSz + HauteurLettreSz - (HauteurLettreSz / 5.8)
        '
    Next
    '
CALCUL_PROC:
    '
    'si aucun caractère n'a encore été créé, on quitte ici
    If UBound(ClsL) = 0 Then Exit Sub
    '
    'on calcule le % d'opacité (distance caméra / ligne)
    opDistSz = Sqr((XSz - CamVX) ^ 2 + (ZSz - CamVZ) ^ 2)
    '
    opDistSz = 100 - (((opDistSz - 500) / 1000) * 100)
    '
    'on vérifie qu'on ne dépasse pas 100
    If opDistSz > 100 Then opDistSz = 100
    '
    'on ne fait rien si l'opacité <= 0 (on utilise moins de ressources)
    'If opDistSz <= 0 Then Exit Sub
    '
    'dMsg = opDistSz
    '
    'on calcule les différentes données pour chaque caractères
    For i = 1 To UBound(ClsL)
        '
        ClsL(i).Calculer NbrCarFin, opDistSz
        '
    Next
    '
    'dMsg = NbrCarFin
    '
    'on vérifie si tous les caractères ont écoulé leur durée de vie
    If NbrCarFin = NbrCarNSz Then
        '
        'on ré-initialise certaines variables
        prctSz = 0
        Tps1 = 0
        '
        YSz = -750 + Rnd() * 250
        'XSz = CamTX + 300 - Rnd() * 600
        XSz = CamTX + 1000 - Rnd() * 2000
        'ZSz = CamTZ - (Rnd() * (CamTZ - CamVZ - 500))
        ZSz = CamTZ - (Rnd() * (CamDistance - 1000))
        '
        tTps1 = 800 + Rnd() * 800
        '
        NbrCarNSz = 40 + Rnd() * 40
        '
        NbrCarFin = 0
        NbrCarSz = 0
        ReDim ClsL(0 To 0)
        '
    End If
    '
End Sub
'
'AFFICHAGE DES LETTRES LES UNES APRES LES AUTRES
Public Sub Afficher()
    '
    'si aucune lettre n'a encore été créée, on quitte ici
    If UBound(ClsL) = 0 Then Exit Sub
    '
    'on vérifie que la ligne se trouve bien devant la caméra, sinon on ne l'affiche pas
    'If ZSz < CamVZ And CamVZ < CamTZ Then Exit Sub
    'If ZSz > CamVZ And CamVZ > CamTZ Then Exit Sub
    '
    'on vérifie qu'on est pas en-dessous de 100 (opacité générale de la ligne par rapport à la distance la séparant de la caméra), sinon, on n'affiche rien
    If opDistSz > 0 Then
        '
        'on affiche les caractères
        For i = 1 To UBound(ClsL)
            '
            ClsL(i).AfficherLettre
            '
        Next
        '
        NbrLignesAff = NbrLignesAff + 1
        '
    End If
    '
    '
End Sub

