VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLettre1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UN CARACTERE DE TYPE "KATANA" POUR L'AFFICHAGE DES TITRES    *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'stocke la g�om�trie d'un objet
Dim VBuffer As Direct3DVertexBuffer8
'
'dur�e des transitions (ms)
Dim Tps1 As Long
Public tTps1 As Long
Dim Tps2 As Long
Public tTps2 As Long
'
'opacit� du caract�re vert
Public OpaciteSz As Long
'
'opacit� du caract�re blanc
Public OpaciteSz2 As Long
'
'coordonn�es du caract�re
Public XSz As Single
Public YSz As Single
'
'couleur du caract�re
'Public RSz As Long
'Public VSz As Long
'Public BSz As Long
'
'�tat du caract�re
Public Etat As Integer
'
'caract�re
Public CarSz As Long
'
'variables quelconques
Dim i As Integer
Dim i2 As Single
Dim prctSz As Single
'
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    'on initialise les variables de base
    Etat = 1
    tTps1 = 30
    '
End Sub
'
'proc�dure permettant d'afficher une image
Private Sub AfficherCarSz1(NumSz As Long, X As Single, Y As Single, Optional CouleurSz As Long = vbWhite)
    '
    Dim cTmp As Long
    '
    'on demande au vertex shader d'utiliser le format de vertex appropri�
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on calcule les coordonn�es du caract�re
    cTmp = (NumSz - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (NumSz - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on d�termine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on g�n�re les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix(1) = CreateTLVertex(X + LargeurLettreSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix(3) = CreateTLVertex(X + LargeurLettreSz, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Blanc2
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
    'on incr�mente de 1 le nombre d'objets affich�s
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'proc�dure permettant d'afficher une image
Private Sub AfficherCarSz2(NumSz As Long, X As Single, Y As Single, Optional CouleurSz As Long = vbWhite)
    '
    Dim cTmp As Long
    '
    'on calcule les coordonn�es du caract�re
    cTmp = (NumSz - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (NumSz - 1) Mod 10
    CharX = 50 * cTmp
    '
    'on d�termine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz2) & Hex(CouleurSz))
    '
    'on demande au vertex shader d'utiliser le format de vertex appropri�
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on g�n�re les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix(1) = CreateTLVertex(X + LargeurLettreSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix(3) = CreateTLVertex(X + LargeurLettreSz, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Blanc
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
    'on incr�mente de 1 le nombre d'objets affich�s
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'proc�dure qui va afficher l'effet
Public Sub AfficherCar()
    '
    'dMsg = OpaciteSz & " - " & OpaciteSz2
    '
    'si l'�tat de l'effet = 0, on quitte la proc�dure (il n'y a rien � afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les op�rations en fontion de l'�tat de l'effet
    Select Case Etat
        '
        Case 1 'on augmente l'opacit� du caract�re vert jusqu'� 50%
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz = 2.55 * prctSz
            '
            'on affiche le caract�re vert
            AfficherCarSz1 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe � l'�tape suivante
            If prctSz = 50 Then
                '
                Etat = 2
                Tps1 = 0
                'tTps1 = 30
                '
            End If
            '
        Case 2 'on continue d'augmenter l'opacite du caract�re vert et on augment celle du blanc
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz = 2.55 * prctSz + 127.5
            OpaciteSz2 = 2.55 * prctSz
            '
            'on affiche les caract�res
            AfficherCarSz1 CarSz, XSz, YSz
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe � l'�tape suivante
            If prctSz = 50 Then
                '
                Etat = 3
                Tps1 = 0
                'tTps1 = 50
                '
            End If
            '
        Case 3 'on continue d'augmenter celle du blanc
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz2 = 2.55 * prctSz + 127.5
            '
            'on affiche les caract�res
            AfficherCarSz1 CarSz, XSz, YSz
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe � l'�tape suivante
            If prctSz = 50 Then
                '
                Etat = 4
                Tps1 = 0
                tTps1 = 10
                '
            End If
            '
        Case 4 'on maintient l'opacit� des caract�res
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on affiche le caract�re blanc
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 5
                Tps1 = 0
                tTps1 = 50
                '
            End If
            '
        Case 5 'on diminue l'opacit� du caract�re blanc jusqu'� 50%
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz2 = 255 - (2.55 * prctSz)
            '
            'on affiche les caract�res
            AfficherCarSz1 CarSz, XSz, YSz
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe � l'�tape suivante
            If prctSz = 50 Then
                '
                Etat = 6
                Tps1 = 0
                'tTps1 = 40
                '
            End If
            '
        Case 6 'on continue � diminuer l'opacit� du caract�re blanc et on commence � diminuer celle du caract�re vert
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz2 = 127.5 - (2.55 * prctSz)
            OpaciteSz = 255 - (2.55 * prctSz)
            '
            'on affiche les caract�res
            AfficherCarSz1 CarSz, XSz, YSz
            AfficherCarSz2 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on passe � l'�tape suivante
            If prctSz = 50 Then
                '
                Etat = 7
                Tps1 = 0
                'tTps1 = 30
                '
            End If
            '
        Case 7 'on r�duit l'opacit� du caract�re vert
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 50
            '
            OpaciteSz = 127.5 - (2.55 * prctSz)
            '
            'on affiche le caract�re vert
            AfficherCarSz1 CarSz, XSz, YSz
            '
            'si on a atteint les 50%, on revient � la premi�re �tape
            If prctSz = 50 Then
                '
                Etat = 8
                OpaciteSz = 0
                OpaciteSz2 = 0
                Tps1 = 0
                tTps1 = 500 + Rnd() * 1000
                '
            End If
            '
        Case 8 'on attend
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                Etat = 9
                Tps1 = 0
                tTps1 = 150 + Rnd() * 200
                '
            End If
            '
        Case 9 'on incr�mente l'opacit� du caract�re vert (40%)
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 40
            '
            OpaciteSz = 2.55 * prctSz
            '
            'on affiche le caract�re vert
            AfficherCarSz1 CarSz, XSz, YSz
            '
            'si on a atteint les 40%, on passe � l'�tape suivante
            If prctSz = 40 Then
                '
                Etat = 10
                Tps1 = 0
                'tTps1 = 30
                '
            End If
            '
        Case 10 'on r�duit l'opacit� des caract�res verts � partir de 30
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 30
            '
            OpaciteSz = 76.5 - (2.55 * prctSz)
            '
            'on affiche le caract�re vert
            AfficherCarSz1 CarSz, XSz, YSz
            '
            'si on a atteint les 30%, on revient � la premi�re �tape
            If prctSz = 30 Then
                '
                Etat = 0
                OpaciteSz = 0
                Tps1 = 0
                tTps1 = 30
                '
            End If
            '
        '
    End Select
    '
End Sub



