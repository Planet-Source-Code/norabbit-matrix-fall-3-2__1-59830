VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEffet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE L'EFFET N� 2                                                 *
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
'
'opacit� de l'effet
Public OpaciteSz As Long
'
'coordonn�es de l'effet
Public XSz As Single
Public YSz As Single
Dim XTmp As Single
Dim YTmp As Single
'
'dimensions
Public HauteurSz As Long
Public LargeurSz As Long
Public HTmp As Long
Public LTmp As Long
'
'couleur de l'effet
'Public RSz As Long
'Public VSz As Long
'Public BSz As Long
'
'�tat de l'effet
Public Etat As Integer
'
'variables quelconques
Dim i As Integer
Dim i2 As Single
Dim prctSz As Single
'
'distance et direction du d�placement (haut ou bas)
Public DirectionSz As Single
'
'taille de l'effet (multiplicateur)
Public TailleSz As Integer
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    'on initialise les variables de base
    Etat = 1
    TailleSz = 1
    HTmp = (16 / 21) * HauteurLettreSz
    LTmp = (512 / 18) * LargeurLettreSz
    '
End Sub
'
'proc�dure permettant d'afficher une image
Private Sub AfficherTypeFX(xFX As Single, yFX As Single, X As Single, Y As Single, HSz As Long, LSz As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim cTmp As Long
    '
    'on demande au vertex shader d'utiliser le format de vertex appropri�
    D3DDevice.SetVertexShader FVF_TLVERTEX
    '
    'on d�termine la couleur (on ajoute la valeur alpha)
    CouleurSz = Hex2Long(Hex(OpaciteSz) & Hex(CouleurSz))
    '
    'on g�n�re les points
    vertMatrix(0) = CreateTLVertex(X, Y, 0, 1, CouleurSz, 0, (1 / 512) * xFX, (1 / 512) * yFX)
    vertMatrix(1) = CreateTLVertex(X + LargeurSz, Y, 0, 1, CouleurSz, 0, ((1 / 512) * xFX) + ((1 / 512) * LSz), (1 / 512) * yFX)
    vertMatrix(2) = CreateTLVertex(X, Y + HauteurSz, 0, 1, CouleurSz, 0, (1 / 512) * xFX, ((1 / 512) * yFX) + ((1 / 512) * HSz))
    vertMatrix(3) = CreateTLVertex(X + LargeurSz, Y + HauteurSz, 0, 1, CouleurSz, 0, ((1 / 512) * xFX) + ((1 / 512) * LSz), ((1 / 512) * yFX) + ((1 / 512) * HSz))
    '
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex_Effets
    '
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
    'on incr�mente de 1 le nombre d'objets affich�s
    NbrObjetsAff = NbrObjetsAff + 1
    '
End Sub
'
'proc�dure qui va afficher l'effet
Public Sub AfficherEffet()
    '
    'si l'�tat de l'effet = 0, on quitte la proc�dure (il n'y a rien � afficher)
    If Etat = 0 Then Exit Sub
    '
    'on effectue les op�rations en fontion de l'�tat de l'effet
    Select Case Etat
        '
        Case 1 'on fait apparaitre l'effet
            '
            'on augmente l'opacit� et la taille de l'effet en fonction de la vitesse g�n�rale
            If Tps1 < tTps1 Then
                '
                Tps1 = Tps1 + (1 * VitesseGenSz)
                If Tps1 > tTps1 Then Tps1 = tTps1
                '
                prctSz = (Tps1 / tTps1) * 100
                '
                OpaciteSz = 2.55 * prctSz
                '
                HauteurSz = (HTmp / 2) + ((HTmp / 200) * prctSz)
                LargeurSz = ((LTmp / 2) * TailleSz) + (((LTmp / 200) * TailleSz) * prctSz)
                '
                YTmp = YSz + (HTmp / 4) - ((HTmp / 400) * prctSz)
                XTmp = XSz + ((LTmp / 4) * TailleSz) - (((LTmp / 400) * TailleSz) * prctSz) - ((LTmp / 2) * (TailleSz - 1))
                '
            Else
                '
                OpaciteSz = 255
                Etat = 2
                Tps1 = 0
                tTps1 = tTps1 * 2
                '
            End If
            '
        Case 2 'on r�duit l'opacit� de l'effet
            '
            'on r�duit l'opacit� en fonction de la vitesse g�n�rale
            If Tps1 < tTps1 Then
                '
                Tps1 = Tps1 + (1 * VitesseGenSz)
                If Tps1 > tTps1 Then Tps1 = tTps1
                '
                prctSz = (Tps1 / tTps1) * 100
                prctSz = prctSz * (2 - (prctSz / 100))
                '
                OpaciteSz = 255 - (2.55 * prctSz)
                '
                YTmp = YSz + ((DirectionSz / 100) * prctSz)
                '
            Else
                '
                OpaciteSz = 0
                Etat = 0
                Tps1 = 0
                tTps1 = tTps1 / 2
                '
            End If
            '
        '
    End Select
    '
    'on affiche les effets
    AfficherTypeFX 0, 49, XTmp, YTmp, 16, 512
    '
End Sub


