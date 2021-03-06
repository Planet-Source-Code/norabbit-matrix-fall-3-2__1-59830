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
'classe g�rant un caract�re
Dim ClsL() As ClsLettre1
'
'nombre de caract�res charg�es
Public NbrCarSz As Integer
'
'nombre de caract�res n�cessaires
Public NbrCarNSz As Long
'
'position (x) de la premi�re lettre du titre #2
Public PosPremL As Long
'
'coordonn�es
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
'nous pr�vient si la cr�ation des caract�res est termin�e
Dim OkFinSz As Boolean
'
'variables quelconques (incr�mentation, etc...)
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
    'on d�termine la position du premier caract�re en fonction de la taille que les
    'caract�res auront ainsi que de la position de la premi�re lettre du titre #2
    XSz = PosPremL - ((PosPremL \ LargeurLettreSz + 1) * ((LargeurLettreSz / 18) * 19))
    '
    'on d�termine le nombre de caract�res n�cessaires en fonction de la taille de ceux-ci, de
    'la position du premier caract�re et de la taille de l'affichage choisi.
    NbrCarNSz = (DimL - XSz) \ ((LargeurLettreSz / 18) * 19) + 1
    '
End Sub
'
'AFFICHAGE DES CARACTERES LES UNS APRES LES AUTRES
Public Sub Afficher()
    '
    'si l'�tat = 0, alors on en a fini avec cette ligne
    If EtatSz = 0 Then Exit Sub
    '
    'on incr�mente Tps1 en fonction de la vitesse g�n�rale
    Tps1 = Tps1 + (1 * VitesseGenSz)
    '
    'on v�rifie qu'on ne va pas trop loint
    If Tps1 > tTps1 Then
        '
        Tps1 = tTps1
        '
        'on va directement � la proc�dure d'affichage
        GoTo AFF_PROC
        '
    End If
    '
    'on calcule le pourcentage
    prctSz = (Tps1 / tTps1) * 100
    '
    'dMsg = prctSz
    '
    'on calcule le nombre de nouveaux caract�res � afficher en fonction du pourcentage
    NbrSz = ((NbrCarNSz / 100) * prctSz) - NbrCarSz
    '
    'dMsg = NbrSz
    '
    'on cr�e chaque caract�re n�cessaire
    For i = 1 To NbrSz
        '
        'on incr�mente le nombre de lettres cr��es
        NbrCarSz = NbrCarSz + 1
        '
        'on cr�e une nouvelle lettre
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
        'on incr�ment la position x
        XSz = XSz + ((LargeurLettreSz / 18) * 19)
        '
    Next
    '
AFF_PROC:
    '
    'si aucun caract�re n'a encore �t� cr��, on quitte ici
    If UBound(ClsL) = 0 Then Exit Sub
    '
    'on affiche les caract�res
    For i = 1 To UBound(ClsL)
        '
        ClsL(i).AfficherCar
        '
    Next
    '
End Sub
