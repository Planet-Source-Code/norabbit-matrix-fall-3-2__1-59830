VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEffet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE L'EFFET N� 3                                                 *
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
Dim tTps1 As Long
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
Dim prctSz As Long
'
'initialisation de la classe
Private Sub Class_Initialize()
    '
    'on initialise les variables de base
    Etat = 1
    tTps1 = 30
    OpaciteSz = 0
    HauteurSz = (HauteurLettreSz / 21) * 25
    LargeurSz = 760
    XSz = -760
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
        Case 1 'on attend un peu
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then
                '
                Etat = 2
                Tps1 = 0
                tTps1 = 600
                '
            End If
            '
        Case 2 'on d�place l'effet sur toute la longueur de l'�cran et on augmente son opacit�
            '
            'on incr�ment la variable Tps1 en fonction de la vitesse sp�cifi�e
            Tps1 = Tps1 + (1 * VitesseGenSz)
            '
            'on v�rifie qu'on ne va pas trop loint
            If Tps1 > tTps1 Then Tps1 = tTps1
            '
            'on calcule le pourcentage
            prctSz = (Tps1 / tTps1) * 100
            '
            XSz = ((DimL + 760) * prctSz / 100) - 760
            '
            OpaciteSz = 5 * (prctSz * (2 - (2 / 100 * prctSz)))
            '
            'si on a atteint les 50%, on passe � l'�tape suivante
            If prctSz = 100 Then
                '
                Etat = 0
                Tps1 = 0
                tTps1 = 30
                '
            End If
            '
        '
    End Select
    '
    'dMsg = XSz & " hum"
    '
    'on affiche l'effet
    'AfficherTypeFX 0, 125, XSz, YSz, 60, 512
    'AfficherTypeFX 0, 127, XSz, YSz, 57, 512
    '
    'AfficherTypeFX 0, 70, XSz - 500, YSz, 53, 512
    '
    AfficherTypeFX 0, 127, XSz, YSz, 57, 512
    '
End Sub


