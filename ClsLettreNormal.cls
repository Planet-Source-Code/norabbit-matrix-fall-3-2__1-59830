VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLettreNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UN CARACTERE NORMAL POUR L'AFFICHAGE DES TITRES              *
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
'faut-il afficher le caract�re bizarre ou pas
Dim ChngSz As Integer
'
'reste lors du changement de caract�re
Dim ResteSz As Integer
'
'opacit� du caract�re normal
Public OpaciteSz As Long
'
'opacit� du caract�re "bizarre"
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
'caract�res
Public CarNorm As Long
Public CarBiz As Long
'
'variables quelconques
Dim prctSz As Single
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    tTps1 = 250
    tTps2 = 20 + Rnd() * 500
    'tTps2 = 60
    ChngSz = Rnd() * 1
    '
End Sub
'
'proc�dure permettant d'afficher une lettre
Private Sub AfficherLettre(Num�roLettre As Long, X As Single, Y As Single, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on demande au vertex shader d'utiliser le format de vertex appropri�
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on calcule les coordonn�es de la lettre
    cTmp = (Num�roLettre - 1) \ 10
    CharY = 58 * cTmp
    cTmp = (Num�roLettre - 1) Mod 10
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
    D3DDevice.SetTexture 0, MatrixTex_Normal
    '
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
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
    CouleurSz = Hex2Long(Hex(OpaciteSz2) & Hex(CouleurSz))
    '
    'on g�n�re les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * CharX, (1 / 512) * CharY)
    vertMatrix(1) = CreateTLVertex(X + LargeurLettreSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), (1 / 512) * CharY)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, (1 / 512) * CharX, ((1 / 512) * CharY) + ((1 / 512) * 58))
    vertMatrix(3) = CreateTLVertex(X + LargeurLettreSz, Y + HauteurLettreSz, 0, 1, CouleurSz, 0, ((1 / 512) * CharX) + ((1 / 512) * 50), ((1 / 512) * CharY) + ((1 / 512) * 58))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Blanc3
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
    'on incr�mente de 1 le nombre d'objets affich�s
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'proc�dure qui va afficher la lettre
Public Sub AfficherCar(NbrTmp As Integer)
    '
    'si l'�tat de l'effet = 0, on quitte la proc�dure (il n'y a rien � afficher)
    If Etat = 0 Then Exit Sub
    '
    'dMsg = OpaciteSz & " - " & OpaciteSz2
    '
    'on effectue les op�rations en fontion de l'�tat de l'effet
    Select Case Etat
        '
        Case 1 'on attend un peu
            '
            'on incr�mente la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie si le temps est �coul�
            If Tps1 >= tTps1 Then
                '
                Etat = 2
                Tps1 = 0
                tTps1 = 400
                '
            End If
            '
        Case 2 'on augmente l'opacit� du caract�re bizarre (ou du normal) et on le change
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 >= tTps1 Then
                '
                'on ajuste le pourcentage
                prctSz = 100
                '
                Etat = 3
                Tps1 = 0
                tTps1 = 30 + (Rnd() * 1000)
                ResteSz = tTps1
                '
            Else
                '
                'on calcule le pourcentage
                prctSz = (Tps1 / tTps1) * 100
                '
            End If
            '
            'on v�rifie si on doit afficher le caract�re bizarre
            If ChngSz = 1 Then
                '
                OpaciteSz2 = 2.55 * prctSz
                '
                'on affiche le caract�re bizarre
                AfficherCarSz2 CarBiz, XSz, YSz
                '
                'on change la lettre bizarre en fonction du temps sp�cifi� (tTps2)
                Tps2 = Tps2 + (1 * VitesseGenSz)
                '
                If Tps2 >= tTps2 Then
                    '
                    CarBiz = Rnd() * 56
                    Tps2 = 0
                    'tTps2 = (Rnd() * 250)
                    '
                End If
                '
            Else
                '
                OpaciteSz = 2.55 * prctSz
                '
                'on affiche le caract�re normal
                AfficherLettre CarNorm, XSz, YSz
                '
            End If
            '
        Case 3 'on attend un moment et on change le caract�re suivant la vitesse sp�cifi�e
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie si le temps est �coul�
            If Tps1 >= tTps1 Then
                '
                Etat = 4
                Tps1 = 0
                tTps1 = 600
                '
            End If
            '
            'on v�rifie si on doit afficher le caract�re bizarre
            If ChngSz = 1 Then
                '
                'on affiche le caract�re bizarre
                AfficherCarSz2 CarBiz, XSz, YSz
                '
                'on change la lettre bizarre en fonction du temps sp�cifi� (tTps2)
                Tps2 = Tps2 + (1 * VitesseGenSz)
                '
                If Tps2 >= tTps2 Then
                    '
                    CarBiz = Rnd() * 56
                    Tps2 = 0
                    'tTps2 = Rnd() * 250
                    '
                End If
                '
            Else
                '
                'on affiche le caract�re normal
                AfficherLettre CarNorm, XSz, YSz
                '
            End If
            '
        Case 4 'on augmente l'opacit� du caract�re blanc et on diminue celle du caract�re bizarre
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then
                '
                'on ajuste le pourcentage
                prctSz = 100
                '
                Etat = 5
                Tps1 = 0
                tTps1 = 3000 - ResteSz
                '
            Else
                '
                'on calcule le pourcentage
                prctSz = (Tps1 / tTps1) * 100
                '
            End If
            '
            'on v�rifie si on doit afficher le caract�re bizarre
            If ChngSz = 1 Then
                '
                OpaciteSz = 2.55 * prctSz
                OpaciteSz2 = 255 - (2.55 * prctSz)
                '
                'on affiche les caract�res
                AfficherLettre CarNorm, XSz, YSz
                AfficherCarSz2 CarBiz, XSz, YSz
                '
            Else
                '
                'on affiche le caract�re normal
                AfficherLettre CarNorm, XSz, YSz
                '
            End If
            '
        Case 5 'on attend
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on affiche le caract�re blanc
            AfficherLettre CarNorm, XSz, YSz
            '
            'on v�rifie si le temps est �coul�
            If Tps1 >= tTps1 Then
                '
                Etat = 6
                Tps1 = 0
                tTps1 = 500
                '
            End If
            '
        Case 6 'on diminue l'opacit� du caract�re blanc
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then
                '
                'on ajuste le pourcentage
                prctSz = 100
                '
                Etat = 0
                Tps1 = 0
                tTps1 = 250
                OpaciteSz = 0
                OpaciteSz2 = 0
                NbrTmp = NbrTmp + 1
                '
            Else
                '
                'on calcule le pourcentage
                prctSz = (Tps1 / tTps1) * 100
                '
            End If
            '
            OpaciteSz = 255 - (2.55 * prctSz)
            '
            'on affiche les caract�res
            AfficherLettre CarNorm, XSz, YSz
        '
    End Select
    '
End Sub
