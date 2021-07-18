VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmAssemble 
   BackColor       =   &H00000000&
   Caption         =   "WebBrowser"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TabStrip onglet 
      Height          =   6255
      Left            =   9360
      TabIndex        =   4
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   11033
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Classement par collection"
            Key             =   "collection"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Classement vertical"
            Key             =   "vertical"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vue d'ensemble"
            Key             =   "deux"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trv2 
      Height          =   5775
      Left            =   3720
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10186
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "il"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList il 
      Left            =   3120
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":0EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":104D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":1127
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":16C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":1C5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":1DB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":1E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":23E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":297E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":2DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":3222
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssemble.frx":353C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trv 
      Height          =   5775
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   10186
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "il"
      Appearance      =   1
   End
   Begin VB.TextBox txtUrl 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "dropper l'url ici"
      Top             =   0
      Width           =   9255
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2750
      Left            =   0
      TabIndex        =   0
      Top             =   6750
      Width           =   9255
      ExtentX         =   16325
      ExtentY         =   4851
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image imgBack 
      Height          =   375
      Left            =   10080
      Picture         =   "frmAssemble.frx":3856
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   9240
      Picture         =   "frmAssemble.frx":3EB1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "frmAssemble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim flag As Boolean

Private Sub btnGo_Click()
End Sub

Private Sub Form_Load()
onglet.Width = Me.Width - 150
onglet.Left = 20

trv.ZOrder 0
trv.Width = onglet.Width - 70
trv.Left = onglet.Left + 15
trv2.Width = onglet.Width - 70
trv2.Left = onglet.Left + 15
imgBack.Visible = False

End Sub

Private Sub Form_Resize()
On Error Resume Next

Me.Height = 10065



Select Case onglet.SelectedItem.Key
Case "collection"
trv.ZOrder 0
trv2.ZOrder 1
resize_unique

Case "vertical"
resize_unique
trv.ZOrder 1
trv2.ZOrder 0

Case "deux"
resize_deux

End Select
End Sub

Private Sub Image1_Click()
flag = True
If Left$(txtUrl, 7) <> "http://" Then txtUrl = "http://" & txtUrl


web.Navigate txtUrl

End Sub

Private Sub imgBack_Click()
web.GoBack
flag = True

End Sub

Private Sub onglet_Click()
Select Case onglet.SelectedItem.Key
    Case "collection"
        trv.ZOrder 0
        trv2.ZOrder 1
        resize_unique
    Case "vertical"
        trv2.ZOrder 0
        trv.ZOrder 1
        resize_unique
    Case "deux"
        
        placement
        
End Select
End Sub

Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
If Left$(Node.Key, 6) = "script" Then
    reste = Replace(Node.Key, "script", "")
    FrmScript.Visible = True
    FrmScript.rt.Text = web.Document.scripts(Val(reste)).Text
End If


End Sub

Private Sub trv2_NodeClick(ByVal Node As MSComctlLib.Node)

If InStr(Node.FullPath, "<FRAME>") > 0 Then
    leserver = trv2.Nodes(1).Text
    lapage = Node.Text
    leserver = Replace(leserver, "<server>", "")
    lapage = Replace(lapage, "src :", "")
    url_frame = Trim(leserver) & "/" & Trim(lapage)
    web.Navigate "http://" & url_frame
End If

End Sub

Private Sub txtUrl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo hdl
txt_url = Data.GetData(vbCFText)
flag = True
web.Navigate txt_url
txtUrl = txt_url
Exit Sub
hdl:
txt_url = Data.Files(1)

Resume Next

End Sub

Private Sub Init()
trv2.Nodes.Clear
trv.Nodes.Clear

End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Dim idoc As IHTMLDocument2
Dim img As HTMLImg
Dim lnk As HTMLLinkElement
Dim col As HTMLTable
Dim td As HTMLTableCol
Dim tr As HTMLTableCol
Dim HTMObj As IHTMLElement
Dim noeud As Node
Dim nn As Node
Dim rep As Node
Dim latable As Node
Dim colo As Node
Dim ligne As Node
Dim obj As Node
Dim lobj As Node
Dim labase As Node

If flag Then
If web.Document.Title <> "Cannot find server" Then
Init
Set idoc = web.Document
Set rep = trv.Nodes.Add(, tvwChild, "Images", "Les Images", 12, 12)

For i = 0 To idoc.images.length - 1
    Set img = idoc.images(i)
    Set noeud = trv.Nodes.Add(rep, tvwChild, , "idx : " & img.sourceIndex & ", num : " & i & " | " & img.src, 2, 2)
    Set nn = trv.Nodes.Add(noeud, tvwChild, "it" & i, "InnerText : " & img.innerText, 10, 10)
    Set nn = trv.Nodes.Add(noeud, tvwChild, "ialt" & i, "Alt : " & img.alt, 10, 10)
       
Next
Set rep = trv.Nodes.Add(, tvwChild, "Liens", "Les Liens", 12, 12)

For i = 0 To idoc.links.length - 1
    Set lnk = idoc.links(i)
    Set noeud = trv.Nodes.Add(rep, tvwChild, , "idx : " & lnk.sourceIndex & ", num : " & i & " | " & lnk, 2, 2)
    Set nn = trv.Nodes.Add(noeud, tvwChild, "lt" & i, "InnerText : " & lnk.innerText, 10, 10)
    'Set nn = trv.Nodes.Add(noeud, tvwChild, "ialt" & i, "Alt : " & img.alt, 10, 10)
       
Next

Set rep = trv.Nodes.Add(, tvwChild, "Puces", "Les puces", 12, 12)

For i = 0 To idoc.All.tags("li").length - 1
    Set pc = idoc.All.tags("li").Item(i)
    Set noeud = trv.Nodes.Add(rep, tvwChild, , "idx : " & pc.sourceIndex & ", num : " & i & " | " & "puce", 2, 2)
    Set nn = trv.Nodes.Add(noeud, tvwChild, , "InnerText : " & pc.innerText, 10, 10)
    'Set nn = trv.Nodes.Add(noeud, tvwChild, "ialt" & i, "Alt : " & img.alt, 10, 10)
       
Next

Set rep = trv.Nodes.Add(, tvwChild, "Tableaux", "Les Tables", 12)

For tb = 0 To idoc.All.tags("table").length - 1
    Set col = idoc.All.tags("table").Item(tb)
    Set latable = trv.Nodes.Add(rep, tvwChild, "tt" & tb, "Table : " & tb & ",SourceIndex : " & col.sourceIndex, 5, 5)
    
    For h = 0 To col.All.length - 1
        letext = col.All(h).innerText
        'letext = Replace(letext, vbCrLf, " ")
           
        If LCase$(col.All(h).tagName) = "tr" Then
            Set ligne = trv.Nodes.Add(latable, tvwChild, , "idx : " & col.All(h).sourceIndex & " | " & letext, 7)
            Set obj = ligne
        End If
        If LCase$(col.All(h).tagName) = "td" Then
            Set colo = trv.Nodes.Add(ligne, tvwChild, , "idx : " & col.All(h).sourceIndex & " | " & letext, 6)
            Set obj = colo
        End If
        
        
        If LCase$(col.All(h).tagName) <> "tr" And LCase$(col.All(h).tagName) <> "td" Then
           'Debug.Print col.All(h).tagName
           
           If obj Is Nothing Then Set obj = latable
           If LCase$("<" & col.All(h).tagName & ">") = "<img>" Then lemot = col.All(h).src
           If LCase$("<" & col.All(h).tagName & ">") = "<a>" Then lemot = col.All(h).href
           If lemot = "" Then lemot = letext
             
            
            
            Set lobj = trv.Nodes.Add(obj, tvwChild, , "<" & col.All(h).tagName & ">" & "idx : " & col.All(h).sourceIndex & " | " & lemot, 10)
            If LCase$("<" & col.All(h).tagName & ">") = "<img>" Then Set lobj = trv.Nodes.Add(lobj, tvwChild, , "ALT : " & col.All(h).alt, 11)
            If LCase$("<" & col.All(h).tagName & ">") = "<a>" Then Set lobj = trv.Nodes.Add(lobj, tvwChild, , "InnerText : " & letext, 11)
            If LCase$("<" & col.All(h).tagName & ">") = "<b>" Then Set lobj = trv.Nodes.Add(lobj, tvwChild, , "InnerText : " & letext, 11)
            If LCase$("<" & col.All(h).tagName & ">") = "<font>" Then Set lobj = trv.Nodes.Add(lobj, tvwChild, , "InnerText : " & letext, 11)
            If LCase$("<" & col.All(h).tagName & ">") = "<script>" Then Set lobj = trv.Nodes.Add(lobj, tvwChild, , "Text : " & col.All(h).Text, 11)
            
            lemot = ""
                
        End If
        
    Next
    
    
Next




Set rep = trv.Nodes.Add(, tvwChild, "Scripts", "Les scripts", 12, 12)

For i = 0 To idoc.scripts.length - 1
    Set HTMObj = idoc.scripts(i)
    Set noeud = trv.Nodes.Add(rep, tvwChild, , "idx : " & HTMObj.sourceIndex & ", num : " & i, 4, 4)
    Set nn = trv.Nodes.Add(noeud, tvwChild, "script" & i, "text : " & HTMObj.Text, 10, 10)
       
Next
Set rep = trv.Nodes.Add(, tvwChild, "remarques", "Les remarques", 12, 12)
For i = 0 To idoc.All.tags("!").length - 1
    Set pc = idoc.All.tags("!").Item(i)
    Set noeud = trv.Nodes.Add(rep, tvwChild, , "idx : " & pc.sourceIndex & ", num : " & i & " | " & "", 2, 2)
    Set nn = trv.Nodes.Add(noeud, tvwChild, , "InnerText : " & pc.innerText, 10, 10)
    'Set nn = trv.Nodes.Add(noeud, tvwChild, "ialt" & i, "Alt : " & img.alt, 10, 10)
       
Next



Dim ele  As IHTMLElement

Set labase = trv2.Nodes.Add(, tvwChild, , " <server>" & web.Document.domain, 13, 13)

For k = 0 To idoc.All.length - 1

Set noeud = trv2.Nodes.Add(labase, tvwChild, , "idx : " & k & " <" & idoc.All(k).tagName & ">", 10, 10)
Set ele = idoc.All(k)
Select Case LCase$(ele.tagName)
    Case "img"
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "InnerText : " & ele.innerText, 10, 10)
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "src : " & ele.src, 10, 10)
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "alt : " & ele.alt, 10, 10)
        
    Case "a"
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "InnerText : " & ele.innerText, 10, 10)
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "src : " & ele.href, 10, 10)
        'Set nn = trv.Nodes.Add(noeud, tvwChild, "aa" & i, "alt : " & ele.alt, 10, 10)
        
    Case "b"
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "InnerText : " & ele.innerText, 10, 10)
        
    Case "p"
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "InnerText : " & ele.innerText, 10, 10)
     Case "tr"
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "InnerText : " & ele.innerText, 10, 10)
    Case "td"
        Set nn = trv2.Nodes.Add(noeud, tvwChild, , "InnerText : " & ele.innerText, 10, 10)
    Case "frame"
        Set nn = trv2.Nodes.Add(noeud, tvwChild, "frame" & k, "src : " & ele.src, 10, 10)
End Select

Next
labase.Expanded = True
imgBack.Visible = True



End If
End If


End Sub
Private Sub resize_unique()
haut = 70
haut = 30
trv2.Left = 20

txtUrl.Width = Me.Width - 1800
Image1.Left = txtUrl.Left + txtUrl.Width + 20
imgBack.Left = Image1.Left + Image1.Width + 10

onglet.Width = Me.Width - 150
trv.Width = onglet.Width - 70
trv2.Width = onglet.Width - 70
web.Width = Me.Width - 150
'onglet.Height = (Me.Height - 370) * (haut / 100)
web.Top = 6840



End Sub
Private Sub resize_deux()
Dim gauche As Integer
Dim droite As Integer
gauche = 50
droite = 50
trv.Width = Me.Width * (gauche / 100) - 40
trv2.Width = Me.Width * (gauche / 100) - 40
trv2.Left = trv.Width + 10
trv.ZOrder 0
trv2.ZOrder 0
txtUrl.Width = Me.Width - 1800
Image1.Left = txtUrl.Left + txtUrl.Width + 20
imgBack.Left = Image1.Left + Image1.Width + 10

onglet.Width = Me.Width - 150
web.Width = Me.Width - 150
End Sub
Private Sub placement()
Dim gauche As Integer
Dim droite As Integer
gauche = 50
droite = 50
trv.Width = Me.Width * (gauche / 100) - 40
trv2.Width = Me.Width * (gauche / 100) - 40
trv2.Left = trv.Width + 10
trv.ZOrder 0
trv2.ZOrder 0


End Sub
