Attribute VB_Name = "ModPrincipal"
'*************************************************************************
'*                                                                       *
'* MODULE DX8                                                            *
'*                                                                       *
'* DECLARATION DES VARIABLES PRINCIPALES + FONCTIONS                     *
'*                                                                       *
'* traduit et compl�t� par Thomas John (thomas.john@swing.be)            *
'*                                                                       *
'* source : http://216.5.163.53/DirectX4VB (DirectX 4 VB, Jack Hoxley)   *
'*                                                                       *
'*************************************************************************
'
'l'objet principal
Public Dx As DirectX8
'
'cet objet contr�le tout ce qui est 3D
Public D3D As Direct3D8
'
'cet objet repr�sente le "hardware" (la carte graphique) utilis� pour le rendu
Public D3DDevice As Direct3DDevice8
'
'une "librairie d'aide"
'D3DX8 est une classe d'aide qui contient une multitude de fonctions destin�es � faciliter la programmation en DX8
Public D3DX As D3DX8
'
'variable servant � d�tecter si le programme tourne ou pas
Public bRunning As Boolean
'
'description des diff�rents types de vertex
Public Const FVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
Public Const Lit_FVF = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
'
'cette structure repr�sente un vertex 2D (identique � la structure "D3DTLVERTEX" pour Dx7)
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tU As Single
    tV As Single
End Type
'
'structure d'un point (3D)
Public Type LITVERTEX
    X As Single
    Y As Single
    Z As Single
    Color As Long
    Specular As Long
    tU As Single
    tV As Single
End Type
'
'structure vertex pour chaque "sprite" 3D
Public vertMatrix3D(3) As LITVERTEX
'
'structure vertex pour chaque "sprite" 2D
Public vertMatrix(3) As TLVERTEX
'
'fonte
Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public TextRect As RECT
Public fnt As New StdFont
'
'Pi
Public Const pi As Single = 3.14159265358979
'
Public matWorld As D3DMATRIX '//How the vertices are positioned
'o� la cam�ra se trouve et vers o� pointe-t-elle
Public matView As D3DMATRIX
'comment la cam�ra projecte le monde 3D sur un �cran (surface) 2D
Public matProj As D3DMATRIX
'
'coordonn�es de la cam�ra
Public CamVX As Single
Public CamVY As Single
Public CamVZ As Single
Public CamTX As Single
Public CamTY As Single
Public CamTZ As Single
Public CamDistance As Single
'
'calcul du fps (images par seconde)
Public FPS_NbrFps As Long
Public FPS_NbrImg As Long
Public lFpsTmp As Long
'
'dimensions de l'affichage
Public DimH As Long
Public DimL As Long
'
'texture de fonte matrix
Public MatrixTex_Blanc As Direct3DTexture8
Public MatrixTex_Blanc2 As Direct3DTexture8
Public MatrixTex_Blanc3 As Direct3DTexture8
Public MatrixTex_Vert As Direct3DTexture8
Public MatrixTex_Trainee As Direct3DTexture8
Public MatrixTex_Normal As Direct3DTexture8
Public MatrixTex_Effets As Direct3DTexture8
'
'pause
Public PauseSz As Boolean
'
'affichage du nombre d'images par seconde
Public AffFps As Boolean
'
'permet de ralentir la vitesse d'affichage � un certain nombre d'images par seconde
Public FpsMod As Long
'
'liste contenant toutes les coordonn�es X utilis�es par les lignes
Public ListeCooX As New Collection
'
'dimensions des lettres
Public HauteurLettreSz As Long
Public LargeurLettreSz As Long
'
'mode d'affichage
Public ModeAffSzX As Long
Public ModeAffSzY As Long
Public ModeBit As Long
'
'acc�l�ration mat�rielle ou software
Public AccMatSoft As String
'
'carte choisie
Public CarteChoixSz As Long
'
'cycle de rendu en millisecondes (temps d'attente entre chaque rendu --> diminue le nombre d'image par seconde et donc la vitesse)
Public CycleRenduSz As Long
'
'limite de lignes � afficher
Public LimiteLignesAffSz As Long
'
'nombre de lignes charg�es
Public LimiteLignesChargeSz As Long
'
'classe se chargeant d'afficher les titres #2 (test)
Public cTitre2 As New ClsTitre2
'
'vitesse g�n�rale (multiplicateur)
Public VitesseGenSz As Long
'
'permet d'afficher des messages (debug seulement)
Public dMsg As String
Public dMsg2 As String
Public dMsg3 As String
'
'classe g�rant une lettre 3D
'Public CL3D As New ClsLettre3D1
'
'nombre d'objets affich�s
Public NbrObjetsAff As Long
'
'nombre de lignes affich�es
Public NbrLignesAff As Long
'
'classe qui g�re la cam�ra
Public Camera As New ClsCamera
'
'liste contenant toutes les coordonn�es Z utilis�es par les lignes
'cette liste va nous permettre d'organiser l'affichage des lignes en fonction de leur coordonn�e Z
Public ListeIndexClasses As New ClsCoordonnees
'
'un classe g�rant une ligne 3D
Public CLgn3D() As ClsLigne3D1
'
'liste des commandes
Public CmdSz As New Collection
'
'liste des valeurs des commandes
Public ValCmdSz As New Collection
'
'liste des pauses des commandes
Public PauseCmdSz As New Collection
'
'pointeur indiquant la commande courante � ex�cuter
Public CourCmdSz As Integer
'
'liste des textes � afficher (titre #1)
Public TxtTitre1 As New Collection
'
'liste des textes � afficher (titre #2)
Public TxtTitre2 As New Collection
'
'classe g�rant les �v�nements
Public cEvnmt As New ClsEvenements
'
'pause avant l'affichage de la premi�re ligne
Public PauseAffPremLgn As Long
'
'variables d'incr�mentation
Dim i As Integer
Dim i2 As Integer
Dim iTmp As Integer
'
'classe se chargeant d'afficher les caract�res sp�ciaux (katanas et autres...)
'Public cL1 As New ClsLettre1
'
'
'
'*******************************************************************
'* Initialise
'*******************************************************************
'
Public Function Initialise(FrmObjet As Form, DimLargeur As Long, DimHauteur As Long) As Boolean
    '
    On Error Resume Next
    '
    'd�crit notre mode d'affichage
    Dim DispMode As D3DDISPLAYMODE
    '
    'd�crit notre mode de vue
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    '
    'pour les filtreurs de texture
    Dim Caps As D3DCAPS8 '//For Texture Filters
    '
    'on cr�e notre objet principal
    Set Dx = New DirectX8
    '
    'on cr�e l'interface Direct3D via notre objet principal
    Set D3D = Dx.Direct3DCreate()
    '
    'on cr�e notre librairie d'aide
    Set D3DX = New D3DX8
    '
    'DispMode.Format = D3DFMT_X8R8G8B8
    'DispMode.Format = D3DFMT_A8R8G8B8
    DispMode.Format = D3DFMT_R5G6B5 'si ce mode ne fonctionne pas, utilisez celui juste au-dessus
    DispMode.Width = DimLargeur
    DispMode.Height = DimHauteur
    '
    DimL = DimLargeur
    DimH = DimHauteur
    '
    D3DWindow.BackBufferCount = 1 '1 BackBuffer
    D3DWindow.BackBufferWidth = DispMode.Width
    D3DWindow.BackBufferHeight = DispMode.Height
    D3DWindow.hDeviceWindow = FrmObjet.hWnd
    D3DWindow.EnableAutoDepthStencil = 1
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    D3DWindow.BackBufferFormat = DispMode.Format
    '
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
        '
        'on peut utiliser un Z-Buffer de 32-bit
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D32
        '
    Else
        '
        If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D24) = D3D_OK Then
            '
            'on peut utiliser un Z-Buffer de 24-bit
            D3DWindow.AutoDepthStencilFormat = D3DFMT_D24
            '
        Else
            '
            If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
                '
                'on peut utiliser un Z-Buffer de 16-bit
                D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
                '
            End If
            '
        End If
        '
    End If
    '
    'on montre notre feuille pour �tre s�r
    FrmObjet.Show
    '
    'on la met au-dessus de toutes
    SetWindowPos FrmObjet.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    '
    'cette ligne cr�e un "device" qui utilise la carte graphique ("hardware") pour effectuer les calculs si possible,
    'ou le processeur ("software") et utilise comme objet de r�ception notre feuille principale
    'on lance le mode hardware ou software selon les options charg�es
    Select Case AccMatSoft
        '
        Case "REF"
            '
            Set D3DDevice = D3D.CreateDevice(CarteChoixSz, D3DDEVTYPE_REF, FrmObjet.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
            '
        Case "HAL"
            '
            Set D3DDevice = D3D.CreateDevice(CarteChoixSz, D3DDEVTYPE_HAL, FrmObjet.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
            '
        '
    End Select
    '
    'nos points (vertices) n'ont pas besoin de lumi�re, donc on d�sactive cette option
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    '
    'ceci dit � dx d'afficher les triangles m�me s'ils ne sont pas en face de la cam�ra
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    '
    'd�clarations utiles pour le rendu de textures transparantes
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    '
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR)
    'Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_TFACTOR)
    '
    'filtrage de texture : donne un meilleur r�sultat lors d'un redimensionnement d'une texture
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    '
    'on active notre Z-Buffer
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    '
    'la matrice "World"
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    '
    'la matrice "view" (en gros, la cam�ra)
    D3DXMatrixLookAtLH matView, MakeVector(CamVX, CamVY, CamVZ), MakeVector(CamTX, CamTY, CamTZ), MakeVector(0, -1, 0)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    '
    'la matrice de projection (angle, distance de rendu)
    D3DXMatrixPerspectiveFovLH matProj, pi / 2, 1, 0.1, 100000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    '
    '
    'initialisation du rendu du texte
    fnt.Name = "Arial"
    fnt.Size = 10
    fnt.Bold = True
    Set MainFontDesc = fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    TextRect.Top = 1
    TextRect.Left = 1
    TextRect.bottom = DimH
    TextRect.Right = DimL
    '
    '**************************************
    '** chargement des textures          **
    '**************************************
    '
    Set MatrixTex_Blanc = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_blanches.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Blanc2 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_blanches2.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Blanc3 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_blanches3.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Vert = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_vertes.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Normal = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\fontes_normales.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Set MatrixTex_Effets = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\matrixfall\effets.png", 512, 512, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    '
    '**************************************
    '** fin du chargement des textures   **
    '**************************************
    '
    'on initialise le nombre de lignes charg�es (aucune donc -1 car les tableaux commencent toujours par 0)
    LimiteLignesChargeSz = -1
    '
    'on initialise le tableau des lignes 3d
    ReDim CLgn3D(0 To 0)
    '
    'on ajoute quelques commandes (affichage de textes)
    'CmdSz.Add "txt"
    'ValCmdSz.Add "matrix fall 3"
    'PauseCmdSz.Add "10000"
    '
    CmdSz.Add "txt"
    ValCmdSz.Add "realise par_THOMAS JOHN"
    PauseCmdSz.Add "10000"
    '
    CmdSz.Add "txt"
    ValCmdSz.Add "email_THOMAS.JOHN@OPEN-DESIGN.BE"
    PauseCmdSz.Add "1"
    '
    CmdSz.Add "txt"
    ValCmdSz.Add "website_WWW.OPEN-DESIGN.BE"
    PauseCmdSz.Add "1"
    '
    CmdSz.Add "noop"
    ValCmdSz.Add ""
    PauseCmdSz.Add "20000"
    '
    'on sp�cifie le temps de la premi�re pause de la classe g�rant les �v�nements
    cEvnmt.tTps2 = CLng(PauseCmdSz.Item(CourCmdSz))
    '
    'on cache le curseur
    ShowCursor False
    '
    'si on arrive jusqu'ici, c'est que tout s'est bien pass�
    Initialise = True
    '
    'on g�re les erreurs survenues durant l'initialisation ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors de l'initialisation :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
    Dim RotateAngle As Single
    Dim matTemp As D3DMATRIX 'contient des donn�es temporaires
    '
    bRunning = True
    '
    '-1 pour �viter la division par z�ro
    lFpsTmp = GetTickCount - 1
    '
    Do While bRunning
        '
        'on v�rifie si il y a une pause
        If PauseSz = False Then
            '
            'nombre d'images par secondes
            FPS_NbrFps = FPS_NbrImg / ((GetTickCount - lFpsTmp) / 1000)
            '
            'on v�rifie que la variable qui contient le nombre d'images rendues ne soit pas trop grande
            If FPS_NbrImg > 1000000 Then
                '
                FPS_NbrImg = 0
                '
                lFpsTmp = GetTickCount - 1
                '
            End If
            '
            'on rend la sc�ne si le nombre d'images par seconde a �t� atteint
            If FPS_NbrFps <= CycleRenduSz Then
                '
                'on incr�ment le nombre d'images rendues
                FPS_NbrImg = FPS_NbrImg + 1
                '
                'on met � jour la position de la cam�ra
                Camera.Affichage
                '
                'on modifie la matrice "view"
                D3DXMatrixLookAtLH matView, MakeVector(CamVX, CamVY, CamVZ), MakeVector(CamTX, CamTY, CamTZ), MakeVector(0, -1, 0)
                D3DDevice.SetTransform D3DTS_VIEW, matView
                '
                '*****************************************************************************************
                'on "rend" (dessine) la sc�ne
                '*****************************************************************************************
                '
                'on efface la surface et on lui donne la couleur noir
                D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
                '
                dMsg3 = ""
                '
                'on commence le rendu
                D3DDevice.BeginScene
                    '
                    'tous les appels de rendu doivent �tre fait entre "BeginScene" et "EndScene"
                    '
                    '�v�nements
                    cEvnmt.Calcul
                    '
                    'on calcule les lignes en 3D de test
                    i = 0
                    Do While i <= LimiteLignesChargeSz
                        '
                        CLgn3D(i).Calculer
                        '
                        i = i + 1
                        '
                    Loop
                    '
                    'on ajoute tous les index de chaque classe dans la liste vide
                    ListeIndexClasses.ToutEffacer
                    '
                    i = 0
                    Do While i <= LimiteLignesChargeSz 'LimiteLignesAffSz
                        '
                        ListeIndexClasses.Ajout i
                        '
                        i = i + 1
                        '
                    Loop
                    '
                    'on affiche les lignes en fonction de la position de la cam�ra
                    'il faut savoir que si les lignes ne sont pas affich�es dans le bon ordre
                    'la transparence des caract�res n'est pas g�r�e.
                    'il faut, par exemple, que la ligne se trouvant le plus pr�s de la cam�ra
                    'soit affich�e apr�s celle qui se trouve le plus loint de la cam�ra
                    'vous pouvez voir la diff�rence en inversant la condition ci-dessous --> CamVZ - CamTZ < 0
                    If CamVZ - CamTZ > 0 Then
                        '
                        Do While ListeIndexClasses.TotalElements > 0
                            '
                            iTmp = 1
                            '
                            For i = 1 To ListeIndexClasses.TotalElements
                                '
                                If CLgn3D(ListeIndexClasses.ElementIndex(i)).ZSz < CLgn3D(ListeIndexClasses.ElementIndex(iTmp)).ZSz Then iTmp = i
                                '
                            Next
                            '
                            'on affiche cet �l�ment
                            CLgn3D(ListeIndexClasses.ElementIndex(iTmp)).Afficher
                            '
                            'on enl�ve cet �l�ment de la liste
                            ListeIndexClasses.Enlever iTmp
                            '
                        Loop
                        '
                    Else
                        '
                        Do While ListeIndexClasses.TotalElements > 0
                            '
                            iTmp = 1
                            '
                            For i = 1 To ListeIndexClasses.TotalElements
                                '
                                If CLgn3D(ListeIndexClasses.ElementIndex(i)).ZSz > CLgn3D(ListeIndexClasses.ElementIndex(iTmp)).ZSz Then iTmp = i
                                '
                            Next
                            '
                            'on affiche cet �l�ment
                            CLgn3D(ListeIndexClasses.ElementIndex(iTmp)).Afficher
                            '
                            'on enl�ve cet �l�ment de la liste
                            ListeIndexClasses.Enlever iTmp
                            '
                        Loop
                        '
                    End If
                    '
                    'on affiche le titre de test
                    cTitre2.AfficherTitre
                    '
                    'on affiche le caract�re de test
                    'cL1.AfficherCar
                    '
                    dMsg = "Z1 cam�ra : " & CamVZ & vbCrLf
                    dMsg = dMsg & "Z2 cam�ra : " & CamTZ & vbCrLf
                    dMsg = dMsg & "pourcentage 1 : " & Camera.prctSz & vbCrLf
                    dMsg = dMsg & "pourcentage 2 : " & Camera.prctSz2 & vbCrLf
                    dMsg = dMsg & "a et b : " & Camera.aTmp & " - " & Camera.bTmp & vbCrLf
                    dMsg = dMsg & "lignes charg�es: " & LimiteLignesChargeSz + 1 & vbCrLf
                    dMsg = dMsg & "max lignes : " & LimiteLignesAffSz & vbCrLf
                    '
                    'i = 0
                    'Do While i <= LimiteLignesAffSz
                        '
                        'dMsg = dMsg & i & " : " & CLgn3D(i).ZSz & " ; " & CLgn3D(i).NbrCarFin & "/" & CLgn3D(i).NbrCarSz & " ; " & CLgn3D(i).tTps1 & vbCrLf
                        '
                        'i = i + 1
                        '
                    'Loop
                    '
                    'rendu du texte
                    If AffFps = True Then D3DX.DrawText MainFont, &HFFB30000, FPS_NbrFps & " fps (" & Time & ")" & vbCrLf & dMsg3 & vbCrLf & dMsg & vbCrLf & "Nombre d'objets affich�s : " & NbrObjetsAff & vbCrLf & "Nombre de lignes affich�es : " & NbrLignesAff & vbCrLf & dMsg2, TextRect, DT_LEFT        '
                    '
                D3DDevice.EndScene
                '
                'on met � jour l'image � l'�cran
                D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
                '
                'on remet � z�ro le nombre d'objets et de lignes affich�s
                NbrObjetsAff = 0
                NbrLignesAff = 0
                '
                '*****************************************************************************************
                'fin du rendu
                '*****************************************************************************************
                '
            End If
            '
        End If
        '
        'on laisse vb respirer
        DoEvents
        '
    Loop
    '
    'on g�re les erreurs survenues lors du rendu ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors du rendu :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
    'on affiche le curseur
    ShowCursor True
    '
    'la boucle s'est termin�e signifiant que le programme va se fermer
    'il faut avant tout d�charger les objets qu'on a charg� pr�c�demment
    '
    On Error Resume Next 'pour �tre s�r
    '
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set Dx = Nothing
    Set D3DX = Nothing
    '
    'on g�re les erreurs survenues lors du d�chargement des objets ici
    If Err Then
        '
        EcrireLog "### " & Time & " ###" & vbCrLf & "erreur lors du d�chargement des objets :" & vbCrLf & D3DX.GetErrorString(Err.Number) & vbCrLf
        '
    End If
    '
End Function
'
'fonction permettant de cr�er un vecteur en une ligne
Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    '
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z
    '
End Function
'
'fontion permettant de remplir un objet en une seule ligne
Public Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, Color As Long, Specular As Long, tU As Single, tV As Single) As TLVERTEX
    '
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
    '
    With CreateTLVertex
        '
        .X = X
        .Y = Y
        .Z = Z
        .rhw = rhw
        .Color = Color
        .Specular = Specular
        .tU = tU
        .tV = tV
        '
    End With
    '
End Function
'
'fontion permettant de remplir un objet en une seule ligne
Public Function CreateLitVertex(X As Single, Y As Single, Z As Single, Color As Long, Specular As Long, tU As Single, tV As Single) As LITVERTEX
    '
    With CreateLitVertex
        '
        .X = X
        .Y = Y
        .Z = Z
        .Color = Color
        .Specular = Specular
        .tU = tU
        .tV = tV
        '
    End With
    '
End Function
'
'convertit une donn�e hex en long
Public Function Hex2Long(hHex) As Long
    '
    Hex2Long = "&H" & hHex
    '
End Function
'
'lecture des options
Public Sub ChargerOptions()
    '
    On Error Resume Next
    '
    Dim FichSz As Integer
    Dim sTmp As String
    Dim sTmp2() As String
    Dim sTmp3() As String
    '
    FichSz = FreeFile
    '
    Open App.Path & "\matrixfall\matrixfall.ini" For Binary As #FichSz
    '
    sTmp = Space(LOF(FichSz))
    '
    Get #FichSz, , sTmp
    '
    Close #FichSz
    '
    'on r�cup�re chaque information en s�parant celles-ci par "VbCrLf"
    sTmp2() = Split(sTmp, vbCrLf)
    '
    'si le tableau ne contient aucune donn�e, on va directement � la fin de la proc�dure
    If UBound(sTmp2) = -1 Then GoTo FIN_PROC
    '
    'je r�utilise "FichSz" comme variable d'incr�mentation
    For FichSz = 0 To UBound(sTmp2)
        '
        If sTmp2(FichSz) = vbNullString Then GoTo FIN_PROC
        '
        'on r�cup�re les infos en fonction de leur nom
        Select Case Left$(sTmp2(FichSz), 4)
            '
            Case "vtgn" 'vitesse g�n�rale
                '
                VitesseGenSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "accm" 'type d'acc�l�ration (hardware ou software)
                '
                AccMatSoft = Right(sTmp2(FichSz), 3)
                '
            Case "mode" 'mode d'affichage
                '
                sTmp3() = Split(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7), "x")
                '
                'on v�rifie que toutes les infos sont l�
                If UBound(sTmp3) = 2 Then
                    '
                    ModeAffSzX = CLng(sTmp3(0))
                    ModeAffSzY = CLng(sTmp3(1))
                    ModeBit = CLng(sTmp3(2))
                    '
                Else
                    '
                    'sinon, on met le mode par d�faut
                    ModeAffSzX = 800
                    ModeAffSzY = 600
                    ModeBit = 16
                    '
                End If
                '
            Case "haut" 'hauteur des lettres
                '
                HauteurLettreSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "larg" 'largeur des lettres
                '
                LargeurLettreSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "cycl" 'nombre d'images par seconde � rendre
                '
                CycleRenduSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "cart" 'carte vid�o
                '
                CarteChoixSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "liml" 'nombre de lignes � afficher
                '
                LimiteLignesAffSz = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                '
            Case "plng" 'temps d'attente avant l'affichage de la premi�re ligne
                '
                cEvnmt.tTps1 = CLng(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7))
                PauseAffPremLgn = cEvnmt.tTps1
                '
            Case "cmds" 'commande
                '
                sTmp3() = Split(Mid(sTmp2(FichSz), 8, Len(sTmp2(FichSz)) - 7), ";")
                '
                'on v�rifie que toutes les infos sont l� sinon on ne fait rien
                If UBound(sTmp3) = 2 Then
                    '
                    CmdSz.Add sTmp3(0)
                    ValCmdSz.Add sTmp3(1)
                    PauseCmdSz.Add sTmp3(2)
                    '
                End If
                '
            '
        End Select
        '
    Next
    '
FIN_PROC:
    '
    'on v�rifie que les informations importantes sont pr�sentes
    If HauteurLettreSz <= 0 Then
        '
        HauteurLettreSz = 21
        '
    End If
    '
    If LargeurLettreSz <= 0 Then
        '
        LargeurLettreSz = (HauteurLettreSz / 21) * 18
        '
    End If
    '
    If ModeAffSzX < 800 Or ModeAffSzY < 600 Or ModeBit = 0 Then
        '
        ModeAffSzX = 800
        ModeAffSzY = 600
        ModeBit = 16
        '
    End If
    '
    If LimiteLignesAffSz <= 0 Then
        '
        LimiteLignesAffSz = 10
        '
    End If
    '
    'on v�rifie que la "vitesse g�n�rale" (multiplicateur) ne soit pas nul
    If VitesseGenSz = 0 Then VitesseGenSz = 17
    '
End Sub
'
'fonction inscrivant dans le fichier log_matrix_fall.txt les erreurs et autres
Public Sub EcrireLog(TexteSz As String)
    '
    Dim FichSz As Integer
    '
    FichSz = FreeFile
    '
    Open App.Path & "\log_matrix_fall.txt" For Binary As #FichSz
    Seek #FichSz, LOF(FichSz)
    Put #FichSz, , TexteSz & vbCrLf
    Close #FichSz
    '
End Sub
