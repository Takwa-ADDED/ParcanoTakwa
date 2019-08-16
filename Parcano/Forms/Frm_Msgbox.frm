VERSION 5.00
Begin VB.Form Frm_Msgbox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Saisir article"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Lbl_Msgbox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voulez vous enregistrer ce nouveau article en tant que : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_Msgbox.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7215
   End
   Begin VB.Image Cmd_Annul 
      Height          =   615
      Left            =   5160
      Picture         =   "Frm_Msgbox.frx":3BB1A
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Cmd_cont 
      Height          =   615
      Left            =   3480
      Picture         =   "Frm_Msgbox.frx":4DC44
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Cmd_Lub 
      Height          =   615
      Left            =   1800
      Picture         =   "Frm_Msgbox.frx":5EEA6
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Cmd_Prod 
      Height          =   615
      Left            =   120
      Picture         =   "Frm_Msgbox.frx":707C0
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Frm_Msgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim design As String
Private Sub Cmd_Annul_Click()
    FrmSaisiePieceReparation.choix = False
    Unload Me
End Sub
Private Sub Cmd_cont_Click()
    FrmSaisiePieceReparation.choix = True
    Unload Me
End Sub
Private Sub Cmd_Lub_Click()
    FrmSaisiePieceReparation.choix = True
    Call FrmSaisiePieceReparation.Update_ProdLub(design, "Lubrifiant")
    Unload Me
End Sub
Private Sub Form_Load()
    design = FrmSaisiePieceReparation.Txt_Designation.text
End Sub
Private Sub Cmd_Prod_Click()
    FrmSaisiePieceReparation.choix = True
    Call FrmSaisiePieceReparation.Update_ProdLub(design, "Produit")
    Unload Me
End Sub
