VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLigneTitre1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LIGNE DE PLUSIEURS CARACTERES                      *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classe gérant un caractère
Dim ClsL() As ClsLettre1
'
'nombre de caractères chargées
Public NbrCarSz As Integer
'
'nombre de caractères nécessaires
Public NbrCarNSz As Long
'
'position (x) de la première lettre du titre #2
Public PosPremL As Long
'
'coordonnées
Public XSz As Single
Public YSz As Single
'
'transition
Dim Tps1 As Long
Public tTps1 As Long
'
'etat de la ligne
Public EtatSz As Long '1 = active ; 0 = inactive
'
'nous prévient si la création des caractères est terminée
Dim OkFinSz As Boolean
'
'variables quelconques (incrémentation, etc...)
Dim i As Integer
Dim NbrSz As Long
Dim prctSz As Single
'
'
'INITIALISATION DES PARAMETRES
Public Sub InitLigne()
    '
    'on initialise certaines variables
    ReDim ClsL(0)
    EtatSz = 1
    Tps1 = 0
    tTps1 = 500
    NbrCarSz = 0
    '
    'on détermine la position du premier caractère en fonction de la taille que les
    'caractères auront ainsi que de la position de la première lettre du titre #2
    XSz = PosPremL - ((PosPremL \ LargeurLettreSz + 1) * ((LargeurLettreSz / 18) * 19))
    '
    'on détermine le nombre de caractères nécessaires en fonction de la taille de ceux-ci, de
    'la position du premier caractère et de la taille de l'affichage choisi.
    NbrCarNSz = (DimL - XSz) \ ((LargeurLettreSz / 18) * 19) + 1
    '
End Sub
'
'AFFICHAGE DES CARACTERES LES UNS APRES LES AUTRES
Public Sub Afficher()
    '
    'si l'état = 0, alors on en a fini avec cette ligne
    If EtatSz = 0 Then Exit Sub
    '
    'on incrémente Tps1 en fonction de la vitesse générale
    Tps1 = Tps1 + (1 * VitesseGenSz)
    '
    'on vérifie qu'on ne va pas trop loint
    If Tps1 > tTps1 Then
        '
        Tps1 = tTps1
        '
        'on va directement à la procédure d'affichage
        GoTo AFF_PROC
        '
    End If
    '
    'on calcule le pourcentage
    prctSz = (Tps1 / tTps1) * 100
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
        Set ClsL(NbrCarSz) = New ClsLettre1
        '
        With ClsL(NbrCarSz)
            '
            .CarSz = Rnd() * 56
            .XSz = XSz
            .YSz = YSz
            '
        End With
        '
        'on incrément la position x
        XSz = XSz + ((LargeurLettreSz / 18) * 19)
        '
    Next
    '
AFF_PROC:
    '
    'si aucun caractère n'a encore été créé, on quitte ici
    If UBound(ClsL) = 0 Then Exit Sub
    '
    'on affiche les caractères
    For i = 1 To UBound(ClsL)
        '
        ClsL(i).AfficherCar
        '
    Next
    '
End Sub
