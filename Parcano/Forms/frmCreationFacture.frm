VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreationFacture 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Création facture"
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14385
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreationFacture.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11880
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView Lsv_Totaux 
      Height          =   3975
      Left            =   6240
      TabIndex        =   37
      Top             =   7320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "HT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Remise"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "TVA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "TTC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView Lsv_Client 
      Height          =   3855
      Left            =   240
      TabIndex        =   36
      Top             =   3360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numéro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date op."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Energie/Vidange/Piece"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "N.Litre/Timbre"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valeur TTC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "P.U.HT / RP"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView Lsv_Details 
      Height          =   3855
      Left            =   7560
      TabIndex        =   35
      Top             =   3360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "a"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Qte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Libelle"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "P.U.HT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rem %"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "TVA %"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "TOTAL HT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker cda_Fin 
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   148570113
      CurrentDate     =   42875
   End
   Begin MSComCtl2.DTPicker cda_Debut 
      Height          =   255
      Left            =   1320
      TabIndex        =   34
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   148570113
      CurrentDate     =   42875
   End
   Begin MSComCtl2.DTPicker cda_opeartion 
      Height          =   375
      Left            =   9720
      TabIndex        =   33
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   148570113
      CurrentDate     =   42875
   End
   Begin VB.TextBox txt_Timbre 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   26
      Text            =   "00,600"
      Top             =   7800
      Width           =   2655
   End
   Begin VB.TextBox txt_nbc 
      Height          =   285
      Left            =   9720
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_Desected 
      Caption         =   "Désélectionner tout"
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton Cmd_Selected 
      Caption         =   "Sélectionner tout"
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   300
      Left            =   5160
      TabIndex        =   21
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txt_Numero 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   465
      Left            =   1560
      TabIndex        =   14
      Tag             =   "M"
      Top             =   1230
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   11535
      TabIndex        =   4
      Top             =   2400
      Width           =   11535
      Begin VB.TextBox txt_ville 
         Height          =   315
         Left            =   9480
         TabIndex        =   7
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox txt_adresse 
         Height          =   315
         Left            =   5520
         TabIndex        =   6
         Top             =   0
         Width           =   2895
      End
      Begin VB.TextBox txt_rsocial 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   0
         Width           =   2715
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ville :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8880
         TabIndex        =   10
         Top             =   0
         Width           =   435
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   9
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raison sociale :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1290
      End
   End
   Begin VB.TextBox txt_MatriculeStation 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Tag             =   "M"
      Top             =   1920
      Width           =   2655
   End
   Begin SToolBox.SCommand CmdFindStation 
      Height          =   360
      Left            =   4320
      TabIndex        =   11
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   635
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreationFacture.frx":0ECA
      ButtonType      =   1
   End
   Begin SToolBox.SCommand cmdFindNumero 
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   1320
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreationFacture.frx":121D
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   11760
      TabIndex        =   16
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreationFacture.frx":1570
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   10800
      TabIndex        =   17
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreationFacture.frx":16F2
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   11280
      TabIndex        =   18
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreationFacture.frx":1A45
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   10320
      TabIndex        =   19
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreationFacture.frx":1D98
   End
   Begin SToolBox.SCommand CmdPrint 
      Height          =   495
      Left            =   12240
      TabIndex        =   20
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreationFacture.frx":1F1A
   End
   Begin VB.Label cda_Create 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date création :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4440
      TabIndex        =   24
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Operation :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8280
      TabIndex        =   32
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label m 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factures"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   840
      TabIndex        =   30
      Top             =   360
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Image PicBox_Header 
      Height          =   1455
      Left            =   0
      Picture         =   "frmCreationFacture.frx":226D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
   Begin VB.Label nb_bon 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7440
      TabIndex        =   29
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NB.Lignes :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Timbre fiscal :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Station :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Au :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   2
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Période du :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   990
   End
End
Attribute VB_Name = "frmCreationFacture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim THT_MO As Double
Dim Tva_MO As Double
Dim TTC_TBR As Double
Dim TT_RmsPiece As Double
Dim RmsPiece As Double
Dim NbrMO As Integer

Private Sub AppCalcule()

Dim ii
Dim TTC_BV As Double
Dim TTC_BC As Double
Dim TTC_PR As Double
Dim TTC_BR As Double
Dim Timbre As Double

Dim Qte As Double
Dim TV As Double
Dim RM As Double

Dim TT As Double
Dim tva As Double
Dim Remise As Double
Dim TTC_MO As Double

Dim Tot_TVA As Double
Dim Tot_HT As Double
Dim Tot_REM As Double
Dim Tot_TTC As Double

Dim Tot_TVA_ As Double
Dim Tot_HT_ As Double
Dim Tot_REM_ As Double

'Init ttc
TTC_BV = 0
TTC_BC = 0
TTC_PR = 0
TTC_BR = 0

Tot_TVA = 0
Tot_HT = 0
Tot_REM = 0
Tot_TTC = 0

Tot_TVA_ = 0
Tot_HT_ = 0
Tot_REM_ = 0

TOT_TVA_BV = 0
TOT_REM_BV = 0
TOT_HT_BV = 0

TOT_REM_CAR = 0
TOT_TVA_CAR = 0
TOT_HT_CAR = 0

TOT_REM_PR = 0
TOT_TVA_PR = 0
TOT_HT_PR = 0

TOT_REM_BR = 0
TOT_TVA_BR = 0
TOT_HT_BR = 0

'Calcule Tot_HT , Tot_TVA et Tot_Remise
For ii = 1 To Lsv_Details.ListItems.Count
    TT = 0
    tva = 0
    Remise = 0
    
    Mont = Lsv_Details.ListItems(ii).SubItems(3)
    Qte = Lsv_Details.ListItems(ii).SubItems(1)
    TV = Lsv_Details.ListItems(ii).SubItems(5)
    RM = Lsv_Details.ListItems(ii).SubItems(4)
    
    TT = Qte * Mont
    Remise = (TT * RM) / 100
    tva = ((TT - Remise) * TV) / 100
    
    If Lsv_Details.ListItems(ii) = "V" Then
        TOT_TVA_CAR = TOT_TVA_CAR + tva
        TOT_HT_CAR = TOT_HT_CAR + TT
        TOT_REM_CAR = TOT_REM_CAR + Remise
    
    ElseIf Lsv_Details.ListItems(ii) = "A" Then
        TOT_TVA_BV = TOT_TVA_BV + tva
        TOT_HT_BV = TOT_HT_BV + TT
        TOT_REM_BV = TOT_REM_BV + Remise
        
    ElseIf Lsv_Details.ListItems(ii) = "P" Then
        TOT_TVA_PR = TOT_TVA_PR + tva
        TOT_HT_PR = TOT_HT_PR + TT
        TOT_REM_PR = TOT_REM_PR + Remise
    
    ElseIf Lsv_Details.ListItems(ii) = "X" Then
        TOT_TVA_BR = TOT_TVA_BR + tva
        TOT_HT_BR = TOT_HT_BR + TT
        TOT_REM_BR = TOT_REM_BR + Remise
    
    End If
    
   If Not (Lsv_Details.ListItems(ii) = "X") Then
        Tot_TVA = Tot_TVA + tva
        Tot_HT = Tot_HT + TT
        Tot_REM = Tot_REM + Remise
    Else
        Tot_TVA_ = Tot_TVA_ + tva
        Tot_HT_ = Tot_HT_ + TT
        Tot_REM_ = Tot_REM_ + Remise
   End If
Next

'calcule des totaux
'BC
TTC_BC = (TOT_HT_CAR - TOT_REM_CAR) + TOT_TVA_CAR
TTC_BC = CStr(Format(TTC_BC, "#,##0.000"))

'BV
TTC_BV = (TOT_HT_BV - TOT_REM_BV) + TOT_TVA_BV
TTC_BV = CStr(Format(TTC_BV, "#,##0.000"))

'PR
TOT_REM_PR = TOT_REM_PR + TT_RmsPiece
TOT_TVA_PR = TOT_TVA_PR + Tva_MO
TOT_HT_PR = TOT_HT_PR + THT_MO

For ii = 1 To Lsv_Client.ListItems.Count
    If (Lsv_Client.ListItems(ii) = "PR") Then
        If Lsv_Client.ListItems(ii).Checked = True Then TTC_PR = TTC_PR + Lsv_Client.ListItems(ii).ListSubItems(5)
   End If
Next
'TTC_PR = (TOT_HT_PR - TOT_REM_PR) + TOT_TVA_PR + TTC_TBR
TTC_PR = CStr(Format(TTC_PR, "#,##0.000"))

'BR
TTC_BR = (TOT_HT_BR - TOT_REM_BR) + TOT_TVA_BR
TTC_BR = CStr(Format(TTC_BR, "#,##0.000"))

'TOTALE
Timbre = CDbl(Replace(txt_Timbre.Text, ".", ","))
Tot_TVA = TOT_TVA_CAR + TOT_TVA_BV + TOT_TVA_PR - TOT_TVA_BR
Tot_HT = TOT_HT_CAR + TOT_HT_BV + TOT_HT_PR - TOT_HT_BR
Tot_REM = TOT_REM_CAR + TOT_REM_BV + TOT_REM_PR - TOT_REM_BR

'TTC_MO = THT_MO + Tva_MO
Tot_TTC = ((TTC_BC + TTC_BV + TTC_PR) - (TTC_BR)) + (Timbre)
Lsv_Totaux.ListItems.Clear

Set itmX = Lsv_Totaux.ListItems.Add(, , "Tot.BC")
    itmX.SubItems(1) = CStr(Format(TOT_HT_CAR, "#,##0.000"))
    itmX.SubItems(2) = CStr(Format(TOT_REM_CAR, "#,##0.000"))
    itmX.SubItems(3) = CStr(Format(TOT_TVA_CAR, "#,##0.000"))
    itmX.SubItems(4) = CStr(Format(TTC_BC, "#,##0.000"))
        
Set itmX = Lsv_Totaux.ListItems.Add(, , "Tot.BV")
    itmX.SubItems(1) = CStr(Format(TOT_HT_BV, "#,##0.000"))
    itmX.SubItems(2) = CStr(Format(TOT_REM_BV, "#,##0.000"))
    itmX.SubItems(3) = CStr(Format(TOT_TVA_BV, "#,##0.000"))
    itmX.SubItems(4) = CStr(Format(TTC_BV, "#,##0.000"))
        
Set itmX = Lsv_Totaux.ListItems.Add(, , "Tot.PR")
    itmX.SubItems(1) = CStr(Format(TOT_HT_PR, "#,##0.000"))
    itmX.SubItems(2) = CStr(Format(TOT_REM_PR, "#,##0.000"))
    itmX.SubItems(3) = CStr(Format(TOT_TVA_PR, "#,##0.000"))
    itmX.SubItems(4) = CStr(Format(TTC_PR, "#,##0.000"))
        
Set itmX = Lsv_Totaux.ListItems.Add(, , "Tot.BR")
    itmX.SubItems(1) = CStr(Format(TOT_HT_BR, "#,##0.000"))
    itmX.SubItems(2) = CStr(Format(TOT_REM_BR, "#,##0.000"))
    itmX.SubItems(3) = CStr(Format(TOT_TVA_BR, "#,##0.000"))
    itmX.SubItems(4) = CStr(Format(TTC_BR, "#,##0.000"))

Set itmX = Lsv_Totaux.ListItems.Add(, , "Total")
    itmX.SubItems(1) = CStr(Format(Tot_HT, "#,##0.000"))
    itmX.SubItems(2) = CStr(Format(Tot_REM, "#,##0.000"))
    itmX.SubItems(3) = CStr(Format(Tot_TVA, "#,##0.000"))
    itmX.SubItems(4) = CStr(Format(Tot_TTC, "#,##0.000"))

End Sub

Private Sub cda_Fin_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call AfficheDetails_PourCreation(txt_MatriculeStation.Text, cda_debut.Value, cda_fin.Value)
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Cmd_Desected_Click()

Dim T As Long
Dim i
On Error GoTo Err
MouseOn
    For T = 1 To Lsv_Client.ListItems.Count
        Lsv_Client.ListItems(T).Checked = False
    Next T
    
    For i = 1 To Lsv_Details.ListItems.Count
      Lsv_Details.ListItems(i).ListSubItems(1) = 0
    Next
    Call AppCalcule
MouseOff
    nb_bon.Caption = NbBonSelect(Lsv_Client)
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Cmd_Selected_Click()

On Error GoTo Err
MouseOn
    Command1 = True
MouseOff
    nb_bon.Caption = NbBonSelect(Lsv_Client)
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdAdd_Click()

On Error GoTo Err

If (CHECK_ACCES("InS_ff", LInt_UserId) = False) Then
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter Fiche Traffic", vbExclamation
        Exit Sub
    End If

If Lsv_Client.ListItems.Count > 0 Or txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If
Call ViderZone(frmCreationFacture)
txt_Numero.Text = "Auto"
cda_Create.Caption = Date
cda_opeartion.Value = Date
cda_debut.Value = "01/01/" & Year(Date)
cda_fin.Value = Date
txt_Timbre.Text = "00,000"
txt_Timbre.Text = Format(txt_Timbre.Text, "#,##0.600")
txt_MatriculeStation.Enabled = True
CmdFindStation.Enabled = True

txt_MatriculeStation.SetFocus
Lsv_Client.ListItems.Clear
Lsv_Totaux.ListItems.Clear
Lsv_Details.ListItems.Clear
Cmd_Selected.Enabled = True
Cmd_Desected.Enabled = True

'Crémentation
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()

Dim VCode As String
Dim LOBJ_Fact As Facture
Dim LOBJ_Stat As Station

On Error GoTo Err
If txt_Numero.Text = "" Then Exit Sub
If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If

If txt_Numero.Text <> "Auto" Then
If (CHECK_ACCES("supp_ff", LInt_UserId) = False) Then
        MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprimer Fiche Traffic", vbExclamation
        Exit Sub
    End If
End If
   
If MsgBox("Confirmez vous la suppression de cette " & vbNewLine & "facture", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

VCode = txt_Numero.Text
Set LOBJ_Fact = New Facture
Set LOBJ_Stat = New Station

Call LOBJ_Fact.Delete_Fact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

Call LOBJ_Stat.Update_NUMFCT(ErrNumber, ErrDescription, ErrSourceDetail, CNB, -1, txt_MatriculeStation.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

Call DeleteNumFact(VCode)

Call ViderZone(frmCreationFacture)
Lsv_Client.ListItems.Clear
Lsv_Details.ListItems.Clear
Lsv_Totaux.ListItems.Clear
txt_Numero.SetFocus
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub
Private Sub DeleteNumFact(ByVal VCode As String)

Dim LOBJ_Bv As BonVidange
Dim LOBJ_BC As BonCarburant
Dim LOBJ_PRep As PieceReparation

Set LOBJ_Bv = New BonVidange
Set LOBJ_BC = New BonCarburant
Set LOBJ_PRep = New PieceReparation

Call LOBJ_Bv.Update_NumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Call LOBJ_BC.Update_NumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Call LOBJ_PRep.Update_NumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

End Sub

Private Sub CmdFind_Click()

If Lsv_Client.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Unload FrmFind
With FrmFind
    .StrSource = "FactureCarburant"
    .Show vbModal
End With

End Sub

Private Sub cmdFindNumero_Click()

Unload FrmFind
With FrmFind
    .StrSource = "FactureCarburant"
    .Show vbModal
End With
End Sub

Private Sub CmdFindStation_Click()

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Stationfacture"
    .Show vbModal
End With
End Sub

Private Sub CmdPrint_Click()

Dim ttc As Double
Dim tht As Double
Dim tva As Double
Dim J
On Error GoTo Err

If txt_Numero.Text = "" Then Exit Sub
If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If

If MsgBox("Imprimer la facture en cours.?        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    ttc = 0
    tht = 0
    tva = 0
    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        For J = 1 To Lsv_Totaux.ListItems.Count
            If (Lsv_Totaux.ListItems(J)) = "Total" Then
                tht = Lsv_Totaux.ListItems(J).ListSubItems(1)
                tva = Lsv_Totaux.ListItems(J).ListSubItems(3)
                ttc = Lsv_Totaux.ListItems(J).ListSubItems(4)
            End If
        Next
        Call .PrintOutAndApercu_RECAP_Facture(0, tht, tva, ttc)
        .Show
    End With
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdSave_Click()
Dim LOBJ_Stat As Station

On Error GoTo Err

If Left(CheckMandatory(frmCreationFacture), 1) = 1 Then
   Exit Sub
End If

If Lsv_Client.ListItems.Count = 0 Then
    MsgBox "Veuillez saisir details ", vbInformation
    Exit Sub
End If

If txt_Numero.Text <> "Auto" Then
    If (CHECK_ACCES("Maj_ff", LInt_UserId) = False) Then
        MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier Fiche Traffic", vbExclamation
        Exit Sub
    End If
End If
    
If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

If txt_Numero.Text = "Auto" Then
    'Insertion enregistrement assiette
    Set LOBJ_Stat = New Station
    Call LOBJ_Stat.Update_NUMFCT(ErrNumber, ErrDescription, ErrSourceDetail, CNB, 1, txt_MatriculeStation.Text)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    txt_nbc.Text = Return_NBFact(txt_MatriculeStation.Text)
    Call Ajout_Fact
Else
    Call DeleteNumFact(txt_Numero.Text)
    Call Modif_Fact
End If
Call InsertNumFact(txt_Numero.Text)
'MAJ Tables BC BV et PieceRep : Ajout du numFact
   
MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
Call ViderZone(frmCreationFacture)
txt_Numero.Text = ""
cda_Create.Caption = Date
cda_opeartion.Value = Date
cda_debut.Value = "01/01/" & Year(Date)
cda_fin.Value = Date
txt_Timbre.Text = "00,000"
txt_Timbre.Text = Format(txt_Timbre.Text, "#,##0.600")
txt_MatriculeStation.Enabled = True
CmdFindStation.Enabled = True

txt_MatriculeStation.SetFocus
Lsv_Client.ListItems.Clear
Lsv_Details.ListItems.Clear
Lsv_Totaux.ListItems.Clear
Cmd_Selected.Enabled = True
Cmd_Desected.Enabled = True

Exit Sub

Err:
CNB.RollbackTrans
MsgBox Err.Description, vbInformation

End Sub

Private Sub Ajout_Fact()

Dim LOBJ_Fact As Facture
Dim LRs_NewRecord As New Recordset
Dim LInt_NumCompteur As Long

Dim ttc As Double
Dim TTC_BV As Double
Dim TTC_BC As Double
Dim TTC_PR As Double
Dim TTC_BR As Double
Dim Timbre As Double

TTC_BC = CDbl(Lsv_Totaux.ListItems(1).SubItems(4))
TTC_BV = CDbl(Lsv_Totaux.ListItems(2).SubItems(4))
TTC_PR = CDbl(Lsv_Totaux.ListItems(3).SubItems(4))
TTC_BR = CDbl(Lsv_Totaux.ListItems(4).SubItems(4))
Timbre = CDbl(Replace(txt_Timbre.Text, ".", ","))
ttc = (TTC_BC + TTC_BV + TTC_PR) - TTC_BR + Timbre

LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "FactureCarburant")
If ErrNumber <> 0 Then
   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
   ErrNumber = 0
   Exit Sub
End If
txt_Numero.Text = Format(LInt_NumCompteur, "00000")

Set LRs_NewRecord = CreateEmptyRS_Fact
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    .Fields("datedoc") = CDate(cda_Create.Caption)
    .Fields("Station") = txt_MatriculeStation.Text
    .Fields("Periodedu") = CDate(cda_debut.Value)
    .Fields("periodeAu") = CDate(cda_fin.Value)
    .Fields("ttc_bc") = TTC_BC
    .Fields("ttc_bv") = TTC_BV
    .Fields("ttc_pr") = TTC_PR
    .Fields("ttc_BR") = TTC_BR
    .Fields("ttc") = ttc
    .Fields("nbc") = Val(txt_nbc.Text)
    .Fields("dateOP") = CDate(cda_opeartion.Value)
    .Fields("Timbre") = CDbl(Replace(txt_Timbre.Text, ".", ","))
    .Fields("UserInsert") = LInt_UserId
End With

Set LOBJ_Fact = New Facture
Call LOBJ_Fact.Insert_Fact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

End Sub

Private Sub Modif_Fact()

Dim LOBJ_Fact As Facture
Dim LRs_NewRecord As New Recordset

Dim ttc As Double
Dim TTC_BV As Double
Dim TTC_BC As Double
Dim TTC_PR As Double
Dim TTC_BR As Double
Dim Timbre As Double

TTC_BC = CDbl(Lsv_Totaux.ListItems(1).SubItems(4))
TTC_BV = CDbl(Lsv_Totaux.ListItems(2).SubItems(4))
TTC_PR = CDbl(Lsv_Totaux.ListItems(3).SubItems(4))
TTC_BR = CDbl(Lsv_Totaux.ListItems(4).SubItems(4))
Timbre = CDbl(Replace(txt_Timbre.Text, ".", ","))
ttc = (TTC_BC + TTC_BV + TTC_PR) - TTC_BR + Timbre

Set LRs_NewRecord = CreateEmptyRS_Fact
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    .Fields("datedoc") = CDate(cda_Create.Caption)
    .Fields("Station") = txt_MatriculeStation.Text
    .Fields("Periodedu") = CDate(cda_debut.Value)
    .Fields("periodeAu") = CDate(cda_fin.Value)
    .Fields("ttc_bc") = TTC_BC
    .Fields("ttc_bv") = TTC_BV
    .Fields("ttc_pr") = TTC_PR
    .Fields("ttc_BR") = TTC_BR
    .Fields("ttc") = ttc
    .Fields("nbc") = Val(txt_nbc.Text)
    .Fields("dateOP") = CDate(cda_opeartion.Value)
    .Fields("Timbre") = CDbl(Replace(txt_Timbre.Text, ".", ","))
    .Fields("UserUpdate") = LInt_UserId
End With

Set LOBJ_Fact = New Facture
Call LOBJ_Fact.Update_Fact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

End Sub

'MAJ des Bons saisi par l'insertion du numero de facture
Private Sub InsertNumFact(ByVal NUM As String)

Dim LOBJ_PRep As PieceReparation
Dim LOBJ_Bv As BonVidange
Dim LOBJ_BC As BonCarburant
Dim ii As Long

Set LOBJ_PRep = New PieceReparation
Set LOBJ_Bv = New BonVidange
Set LOBJ_BC = New BonCarburant
   
For ii = 1 To Lsv_Client.ListItems.Count
    If Lsv_Client.ListItems(ii).Checked = True Then
        N = Lsv_Client.ListItems(ii).SubItems(1)
        If Lsv_Client.ListItems(ii) = "BC" Then
            Call LOBJ_BC.Insert_NumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, NUM, N)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        ElseIf Lsv_Client.ListItems(ii) = "BV" Then
            Call LOBJ_Bv.Insert_NumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, NUM, N)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        ElseIf Lsv_Client.ListItems(ii) = "PR" Or Lsv_Client.ListItems(ii) = "BR" Then
            Call LOBJ_PRep.Insert_NumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, NUM, N)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        End If
    End If
Next

Set LOBJ_PRep = Nothing
Set LOBJ_Bv = Nothing
Set LOBJ_BC = Nothing
End Sub

Private Function Return_NBFact(VCode As String) As Long

Dim rs As New Recordset
Dim LOBJ_Stat As Station

Return_NBFact = 0
Set LOBJ_Stat = New Station
Set rs = LOBJ_Stat.Get_NumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs.EOF Then
    Return_NBFact = rs("numFCT")
End If
rs.Close

End Function

'Charger les bons de ... à ....
Private Sub Command1_Click()

'On Error GoTo Err
Dim T

Lsv_Client.ListItems.Clear
Lsv_Totaux.ListItems.Clear
Lsv_Details.ListItems.Clear

THT_MO = 0
TTC_TBR = 0
TT_RmsPiece = 0
Tva_MO = 0
RmsPiece = 0
If cda_debut.Value > cda_fin.Value Then
    MsgBox "Vérifier dates saisies", vbInformation, App.ProductName
    Exit Sub
End If
If txt_Numero.Text = "" Then
    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "Numero de facture Vide", vbInformation
        Exit Sub
    End If
End If
    
If txt_Numero.Text = "Auto" Then
    Call AfficheDetails_PourCreation(txt_MatriculeStation.Text, cda_debut.Value, cda_fin.Value)
    For T = 1 To Lsv_Client.ListItems.Count
        Lsv_Client.ListItems(T).Checked = True
    Next T
    Call AppCalcule
Else
    Call AfficheDetails_N(txt_Numero.Text)
End If

nb_bon.Caption = NbBonSelect(Lsv_Client)

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub Form_Load()
Me.Height = 9210
Me.Width = 11715
Me.Move 0, 0
Me.WindowState = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)

'On Error GoTo erreur
'   Dim i As Integer
'   Dim Msg ' Déclare la variable.
'   ' Définit le texte du message.
'   Msg = "Voulez-vous vraiment quitter?"
'   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
'   If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'      Cancel = True
'   Else
'   Unload Me
'
'   End If
'
'   Exit Sub
'erreur:
'   MsgBox Err.Description, 48
End Sub

Private Sub Lsv_Client_DblClick()

Dim ii
Dim Numero As String

On Error GoTo Err
ii = Lsv_Client.SelectedItem.Index
Numero = Lsv_Client.ListItems(ii).ListSubItems(1)
If Lsv_Client.ListItems(ii) = "BC" Then
    With frmConsultBC
        .AfficheRow (Numero)
        .Show vbModal
    End With
ElseIf Lsv_Client.ListItems(ii) = "BV" Then
    With FrmConsultBV
        .AfficheRow (Numero)
        .Show vbModal
    End With
ElseIf Lsv_Client.ListItems(ii) = "PR" Then
    With FrmConsultPieceReception
        .AfficheRow (Numero)
        .Show vbModal
    End With
ElseIf Lsv_Client.ListItems(ii) = "BR" Then
    With FrmConsultPieceReception
        .AfficheRow (Numero)
        .Show vbModal
    End With
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Lsv_Client_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim ii
Dim Energie As String
Dim AA
Dim Qte As Double
Dim prixht As Double
Dim Jj
Dim QteTotal As Double
Dim TotHT As Double
Dim rQ As New Recordset
Dim LOBJ_BC As BonCarburant
Dim LOBJ_Bv As BonVidange
Dim LOBJ_PRep As PieceReparation

Dim PUHT As Double
Dim tva As Double
Dim ttc As Double
Dim Valeur As Double
Dim Numero As String
Dim Designation As String

On Error GoTo Err
ii = Item.Index

If Item.Checked = False Then
    AA = "-"
Else
    AA = "+"
End If

Set LOBJ_BC = New BonCarburant
Set LOBJ_Bv = New BonVidange
Set LOBJ_PRep = New PieceReparation

'produits carburant
If Lsv_Client.ListItems(ii) = "BC" Then
    Set rQ = LOBJ_BC.Get_DetBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lsv_Client.ListItems(ii).ListSubItems(1))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rQ.EOF Then
        While Not rQ.EOF
            Energie = rQ("Energie")
            Qte = CDbl(rQ("Litre"))
            prixht = CDbl(rQ("PrixHT"))
            
            For Jj = 1 To Lsv_Details.ListItems.Count
                QteTotal = 0
                TotHT = 0
                If Trim(Lsv_Details.ListItems(Jj).ListSubItems(2)) = Trim(Energie) And Val(Replace(Lsv_Details.ListItems(Jj).ListSubItems(3), ",", ".")) = Val(Replace(prixht, ",", ".")) Then
                QteTotal = CDbl(Lsv_Details.ListItems(Jj).ListSubItems(1))
                TotHT = CDbl(Lsv_Details.ListItems(Jj).ListSubItems(6))
                    If AA = "-" Then
                        Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal - Qte
                        TotHT = Lsv_Details.ListItems(Jj).ListSubItems(1) * prixht 'TotHT - (Qte * prixht)
                        Lsv_Details.ListItems(Jj).ListSubItems(6) = CStr(Format(TotHT, "#,##0.000"))
                    ElseIf AA = "+" Then
                         Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal + Qte
                         TotHT = CDbl(Lsv_Details.ListItems(Jj).ListSubItems(1)) * prixht 'TotHT + (Qte * prixht)
                        Lsv_Details.ListItems(Jj).ListSubItems(6) = CStr(Format(TotHT, "#,##0.000"))
                    End If
                End If
            Next
            rQ.MoveNext
        Wend
    End If
    rQ.Close
   
'Produits Vidange

ElseIf Lsv_Client.ListItems(ii) = "BV" Then
    Set rQ = LOBJ_Bv.Get_DetBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lsv_Client.ListItems(ii).ListSubItems(1))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rQ.EOF Then
        While Not rQ.EOF
            Qte = CDbl(rQ("Qte"))
            prixht = CDbl(rQ("THT"))
            For Jj = 1 To Lsv_Details.ListItems.Count
                If Trim(Lsv_Details.ListItems(Jj).ListSubItems(2)) = Trim(rQ("Libelle")) Then
                QteTotal = CDbl(Lsv_Details.ListItems(Jj).ListSubItems(1))
                TotHT = CDbl(Lsv_Details.ListItems(Jj).ListSubItems(6))
                    If AA = "-" Then
                        If QteTotal <> 0 Then
                            Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal - Qte
                            Lsv_Details.ListItems(Jj).ListSubItems(6) = TotHT - (Qte * prixht)
                        End If
                    ElseIf AA = "+" Then
                         Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal + Qte
                        Lsv_Details.ListItems(Jj).ListSubItems(6) = TotHT + (Qte * prixht)
                    End If
                End If
            Next
        rQ.MoveNext
        Wend
    End If
    rQ.Close
    
 'Pièces de rèception
Dim total_ht As Double
Dim rms As Double
Dim Tot_rms As Double
Tot_rms = 0
rms = 0
total_ht = 0

ElseIf Lsv_Client.ListItems(ii) = "PR" Then
    Numero = Lsv_Client.ListItems(ii).ListSubItems(1)
    Set rQ = LOBJ_PRep.Get_DetPRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rQ.EOF Then
        While Not rQ.EOF
            Designation = rQ("Designation")
            Qte = rQ("Qte")
            PUHT = rQ("PUHT")
            If Not IsNull(rQ("Remise")) Then rms = CDbl(rQ("Remise"))
            For Jj = 1 To Lsv_Details.ListItems.Count
                If (Lsv_Details.ListItems(Jj)) = "P" _
                    And Trim(Lsv_Details.ListItems(Jj).ListSubItems(2)) = Trim(Designation) _
                    And Val(Replace(Lsv_Details.ListItems(Jj).ListSubItems(3), ",", ".")) = Val(Replace(PUHT, ",", ".")) Then
                    QteTotal = Lsv_Details.ListItems(Jj).ListSubItems(1)
                    TotHT = Lsv_Details.ListItems(Jj).ListSubItems(6)
                    total_ht = total_ht + (Qte * PUHT) - (Qte * PUHT * rms / 100) ' total_ht de la piece de réparation
                    If AA = "-" Then
                        Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal - Qte
                        Lsv_Details.ListItems(Jj).ListSubItems(6) = TotHT - Qte * PUHT
                    ElseIf AA = "+" Then
                         Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal + Qte
                         Lsv_Details.ListItems(Jj).ListSubItems(6) = TotHT + Qte * PUHT
                    End If
                End If
            Next
            rQ.MoveNext
        Wend
    End If
    rQ.Close
    
    ' Bon Retour
ElseIf Lsv_Client.ListItems(ii) = "BR" Then
    Numero = Lsv_Client.ListItems(ii).ListSubItems(1)

    Set rQ = LOBJ_PRep.Get_DetBRetour(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rQ.EOF Then
        While Not rQ.EOF
            Designation = rQ("Designation")
            Qte = rQ("Qte")
            PUHT = rQ("PUHT")
        
            For Jj = 1 To Lsv_Details.ListItems.Count
                If (Lsv_Details.ListItems(Jj)) = "X" _
                    And Trim(Lsv_Details.ListItems(Jj).ListSubItems(2)) = Trim(Designation) _
                    And Val(Replace(Lsv_Details.ListItems(Jj).ListSubItems(3), ",", ".")) = Val(Replace(PUHT, ",", ".")) Then
                    QteTotal = Lsv_Details.ListItems(Jj).ListSubItems(1)
                    If AA = "-" Then
                        Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal - Qte
                    ElseIf AA = "+" Then
                         Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal + Qte
                    End If
                End If
        Next
        rQ.MoveNext
        Wend
    End If
    rQ.Close
End If

Call AppCalcule

Dim J
Dim tva_ As Double
Dim MainOeuvre As Double
Dim timb As Double

J = Lsv_Totaux.SelectedItem.Index

If Lsv_Client.ListItems(ii) = "PR" Then
    Numero = Lsv_Client.ListItems(ii).ListSubItems(1)
    Set rQ = LOBJ_PRep.Get_AssPieceReparation(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    MainOeuvre = 0
    tva_ = 0
    timb = 0
    If Not rQ.EOF Then
            If Not IsNull(rQ("timbre")) Then timb = CDbl(rQ("timbre"))
            If Not IsNull(rQ("PrixMOeuvre")) Then MainOeuvre = CDbl(rQ("PrixMOeuvre"))
            If Not IsNull(rQ("TVA_MOeuvre")) Then tva_ = CDbl(rQ("TVA_MOeuvre"))
            If Not IsNull(rQ("RemisePiece")) Then total_ht = (total_ht + MainOeuvre) * CDbl(rQ("RemisePiece")) / 100
            If MainOeuvre <> 0 Then
                For Jj = 1 To Lsv_Details.ListItems.Count
                    If (Lsv_Details.ListItems(Jj)) = "MO" Then
                        QteTotal = Lsv_Details.ListItems(Jj).ListSubItems(1)
                        PUHT = CDbl(Lsv_Details.ListItems(Jj).ListSubItems(3))
                        If AA = "-" Then
                            Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal - 1
                            Lsv_Details.ListItems(Jj).ListSubItems(3) = CStr(Format(PUHT - MainOeuvre, "#,##0.000"))
                            Lsv_Details.ListItems(Jj).ListSubItems(4) = CStr(Format(Lsv_Details.ListItems(Jj).ListSubItems(4) - (MainOeuvre * CDbl(rQ("RemisePiece")) / 100), "#,##0.000"))
                            Lsv_Details.ListItems(Jj).ListSubItems(5) = CStr(Format(Lsv_Details.ListItems(Jj).ListSubItems(5) - (MainOeuvre - (MainOeuvre * rQ("RemisePiece") / 100)) * tva_ / 100, "#,##0.000"))
                            Lsv_Details.ListItems(Jj).ListSubItems(6) = CStr(Format(PUHT - MainOeuvre, "#,##0.000"))
                            
                        ElseIf AA = "+" Then
                            Lsv_Details.ListItems(Jj).ListSubItems(1) = QteTotal + 1
                            Lsv_Details.ListItems(Jj).ListSubItems(3) = CStr(Format(PUHT + MainOeuvre, "#,##0.000"))
                            Lsv_Details.ListItems(Jj).ListSubItems(4) = CStr(Format(Lsv_Details.ListItems(Jj).ListSubItems(4) + (MainOeuvre * CDbl(rQ("RemisePiece")) / 100), "#,##0.000"))
                            Lsv_Details.ListItems(Jj).ListSubItems(5) = CStr(Format(Lsv_Details.ListItems(Jj).ListSubItems(5) + (MainOeuvre - (MainOeuvre * rQ("RemisePiece") / 100)) * tva_ / 100, "#,##0.000"))
                            Lsv_Details.ListItems(Jj).ListSubItems(6) = CStr(Format(PUHT + MainOeuvre, "#,##0.000"))
                        End If
                    End If
                Next
            End If
            
            For J = 1 To Lsv_Totaux.ListItems.Count
                 If (Lsv_Totaux.ListItems(J)) = "Tot.PR" Then
                    If AA = "-" Then
                        Lsv_Totaux.ListItems(J).ListSubItems(1) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(1) - MainOeuvre, "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(2) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(2) - total_ht, "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(3) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(3) - (MainOeuvre - (MainOeuvre * rQ("RemisePiece") / 100)) * tva_ / 100, "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(4) = Lsv_Totaux.ListItems(J).ListSubItems(4)
                        Lsv_Totaux.ListItems(J).ListSubItems(4) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(4), "#,##0.000"))
                    
                    ElseIf AA = "+" Then
                        Lsv_Totaux.ListItems(J).ListSubItems(1) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(1), "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(2) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(2), "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(3) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(3), "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(4) = Lsv_Totaux.ListItems(J).ListSubItems(4)
                        Lsv_Totaux.ListItems(J).ListSubItems(4) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(4), "#,##0.000"))
                    End If
                End If

                If (Lsv_Totaux.ListItems(J)) = "Total" Then
                    If AA = "-" Then
                            Lsv_Totaux.ListItems(J).ListSubItems(1) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(1) - MainOeuvre, "#,##0.000"))
                            Lsv_Totaux.ListItems(J).ListSubItems(2) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(2) - total_ht, "#,##0.000"))
                            Lsv_Totaux.ListItems(J).ListSubItems(3) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(3) - ((tva_ * MainOeuvre / 100)), "#,##0.000"))
                            Lsv_Totaux.ListItems(J).ListSubItems(4) = Lsv_Totaux.ListItems(J).ListSubItems(4)
                            Lsv_Totaux.ListItems(J).ListSubItems(4) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(4), "#,##0.000"))
                    ElseIf AA = "+" Then
                        Lsv_Totaux.ListItems(J).ListSubItems(1) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(1), "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(2) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(2), "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(3) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(3), "#,##0.000"))
                        Lsv_Totaux.ListItems(J).ListSubItems(4) = Lsv_Totaux.ListItems(J).ListSubItems(4)
                        Lsv_Totaux.ListItems(J).ListSubItems(4) = CStr(Format(Lsv_Totaux.ListItems(J).ListSubItems(4), "#,##0.000"))
                    End If
                End If

            Next
        rQ.MoveNext

    End If
    rQ.Close
End If
nb_bon.Caption = NbBonSelect(Lsv_Client)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation


End Sub

Private Sub txt_MatriculeStation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_MatriculeStation_LostFocus()
If Len(Trim(txt_MatriculeStation.Text)) > 0 Then Call AfficheRow_Station(txt_MatriculeStation.Text)
End Sub

Public Sub AfficheRow_Station(ByVal VCode As String)

Dim LOBJ_Stat As Station
Dim rs As New Recordset

Lsv_Totaux.ListItems.Clear
Lsv_Client.ListItems.Clear
Lsv_Details.ListItems.Clear
Set LOBJ_Stat = New Station
Set rs = LOBJ_Stat.GetStationByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
    
Else
    MsgBox "Code introuvable", vbInformation
    txt_MatriculeStation.SetFocus
    Exit Sub
End If

End Sub

Public Sub AfficheRow(ByVal VCode As String)

Dim LOBJ_Fact As Facture
Dim rs As New Recordset

Call ViderZone(frmCreationFacture)
Lsv_Client.ListItems.Clear
Lsv_Totaux.ListItems.Clear
Lsv_Details.ListItems.Clear

txt_MatriculeStation.Enabled = False
CmdFindStation.Enabled = False

Set LOBJ_Fact = New Facture
Set rs = LOBJ_Fact.Get_FactByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("STATION")) Then txt_MatriculeStation.Text = rs("STATION")
    If Not IsNull(rs("Datedoc")) Then cda_Create.Caption = rs("Datedoc")
    If Not IsNull(rs("dateOp")) Then cda_opeartion.Value = rs("dateOp")
    If Not IsNull(rs("PeriodeDu")) Then cda_debut.Value = rs("PeriodeDu")
    If Not IsNull(rs("PeriodeAu")) Then cda_fin.Value = rs("PeriodeAu")
    If Not IsNull(rs("NBC")) Then txt_nbc.Text = rs("NBC")
    txt_Timbre.Text = Format(rs("Timbre"), "#,##0.000")
            
    Call AfficheRow_Station(rs("STATION"))
    Call AfficheDetails_N(VCode)
    Cmd_Selected.Enabled = False
    Cmd_Desected.Enabled = False
    
    nb_bon.Caption = NbBonSelect(Lsv_Client)
Else
    MsgBox "Numéro facture introuvable", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If
rs.Close

End Sub

'Sub pour créer nouvelle facture
Public Sub AfficheDetails_PourCreation(vCodeSation As String, VdateD As Date, vDateF As Date)

Dim LOBJ_Fact As Facture
Dim LOBJ_BC As BonCarburant
Dim LOBJ_Bv As BonVidange
Dim LOBJ_PRep As PieceReparation
Dim rs As New Recordset

NbrMO = 0
THT_MO = 0
TTC_TBR = 0
TT_RmsPiece = 0
Tva_MO = 0
RmsPiece = 0

Set LOBJ_Fact = New Facture
Set rs = LOBJ_Fact.Get_Details_PourCreation(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCodeSation, VdateD, vDateF)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
'remplir LSV_Client
If Not rs.EOF Then
    While Not rs.EOF
            If Not IsNull(rs("type")) Then Set itmX = Lsv_Client.ListItems.Add(, , rs("type"))
            If Not IsNull(rs("Numero")) Then itmX.SubItems(1) = rs("Numero")
            If Not IsNull(rs("dateop")) Then itmX.SubItems(2) = rs("dateop")
            If Not IsNull(rs("Valeur")) Then itmX.SubItems(5) = CStr(Format(rs("Valeur"), "#,##0.000"))
        rs.MoveNext
    Wend
End If
rs.Close

' Produits carburant
Set LOBJ_BC = New BonCarburant
Set rs = LOBJ_BC.Get_ProdCarb(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCodeSation, VdateD, vDateF)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = Lsv_Details.ListItems.Add(, , "V")
            If Not IsNull(rs("Qte")) Then itmX.SubItems(1) = rs("Qte")
            If Not IsNull(rs("energie")) Then itmX.SubItems(2) = CStr(Format(rs("energie"), "#,##0.000"))
            If Not IsNull(rs("PRIXHT")) Then itmX.SubItems(3) = CStr(Format(rs("PRIXHT"), "#,##0.000"))
            If Not IsNull(rs("Remise")) Then itmX.SubItems(4) = rs("Remise")
            If Not IsNull(rs("TVA")) Then itmX.SubItems(5) = rs("TVA")
            If Not IsNull(rs("PRIXHT")) Then itmX.SubItems(6) = CStr(Format(rs("PRIXHT") * rs("Qte"), "#,##0.000"))
    rs.MoveNext
    Wend
End If
rs.Close

'Produits vidange
Set LOBJ_Bv = New BonVidange
Set rs = LOBJ_Bv.Get_ProdVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCodeSation, VdateD, vDateF)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = Lsv_Details.ListItems.Add(, , "A")
            itmX.SubItems(1) = rs("Qte")
            itmX.SubItems(2) = CStr(Format(rs("libelle"), "#,##0.000"))
            itmX.SubItems(3) = CStr(Format(rs("PRIXHT"), "#,##0.000"))
            If Not IsNull(rs("Remise")) Then itmX.SubItems(4) = rs("Remise")
            If Not IsNull(rs("TVA")) Then itmX.SubItems(5) = rs("TVA")
            itmX.SubItems(6) = CStr(Format(rs("PRIXHT") * rs("Qte"), "#,##0.000"))
        rs.MoveNext
    Wend
End If
rs.Close

'Details Piece de réception
Set LOBJ_PRep = New PieceReparation
Set rs = LOBJ_PRep.Get_ProdPRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCodeSation, VdateD, vDateF)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        Set itmX = Lsv_Details.ListItems.Add(, , "P")
            itmX.SubItems(1) = rs("Qte")
            itmX.SubItems(2) = CStr(Format(rs("Designation"), "#,##0.000"))
            itmX.SubItems(3) = CStr(Format(rs("PUHT"), "#,##0.000"))
            If Not IsNull(rs("Remise")) Then itmX.SubItems(4) = rs("Remise")
            If Not IsNull(rs("tva")) Then itmX.SubItems(5) = rs("tva")
            itmX.SubItems(6) = CStr(Format(rs("PUHT") * rs("Qte"), "#,##0.000"))
            If Not IsNull(rs("RemiseP")) Then
                RmsPiece = CDbl(rs("RemiseP"))
                TT_RmsPiece = TT_RmsPiece + (CDbl(itmX.SubItems(6)) - (CDbl(itmX.SubItems(6)) * rs("Remise") / 100)) * (CDbl(rs("RemiseP")) / 100)
            End If
        rs.MoveNext
    Wend
End If
rs.Close
    
Dim rmsMO As Double
rmsMO = 0
Set rs = LOBJ_PRep.Get_RmsMOTimbPRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCodeSation, VdateD, vDateF)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        If Not IsNull(rs("MOeuvre")) Then
            If rs("MOeuvre") <> 0 Then NbrMO = NbrMO + 1
            THT_MO = THT_MO + CDbl(rs("MOeuvre")) '- (CDbl(rs("MOeuvre")) * CDbl(rs("RemiseP")) / 100)
            If Not IsNull(rs("RemiseP")) Then
                rmsMO = rmsMO + (CDbl(rs("MOeuvre")) * CDbl(rs("RemiseP")) / 100)
                TT_RmsPiece = TT_RmsPiece + (CDbl(rs("MOeuvre")) * CDbl(rs("RemiseP")) / 100)
            End If
        End If
        If Not IsNull(rs("TVA_MOeuvre")) Then Tva_MO = Tva_MO + (CDbl(rs("MOeuvre")) - (CDbl(rs("MOeuvre")) * CDbl(rs("RemiseP")) / 100)) * (CDbl(rs("TVA_MOeuvre")) / 100)
        If Not IsNull(rs("timbre")) Then TTC_TBR = TTC_TBR + CDbl(rs("timbre"))
    rs.MoveNext
    Wend
End If
    rs.Close
    'Ajout ligne Main d'oeuvre dans la liste des détails
    Set itmX = Lsv_Details.ListItems.Add(, , "MO")
        itmX.SubItems(1) = NbrMO
        itmX.SubItems(2) = "Main d'oeuvre"
        itmX.SubItems(3) = CStr(Format(THT_MO, "#,##0.000"))
        itmX.SubItems(4) = CStr(Format(rmsMO, "#,##0.000"))
        itmX.SubItems(5) = CStr(Format(Tva_MO, "#,##0.000"))
        itmX.SubItems(6) = CStr(Format(THT_MO, "#,##0.000"))
    
'Detai Bon Retour Pièce de rèception
Set rs = LOBJ_PRep.Get_ProdBRetour(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCodeSation, VdateD, vDateF)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = Lsv_Details.ListItems.Add(, , "X")
            itmX.SubItems(1) = rs("Qte")
            itmX.SubItems(2) = CStr(Format(rs("Designation"), "#,##0.000"))
            itmX.SubItems(3) = CStr(Format(rs("PUHT"), "#,##0.000"))
            itmX.SubItems(4) = rs("Remise")
            itmX.SubItems(5) = rs("tva")
            itmX.SubItems(6) = CStr(Format(rs("PUHT") * rs("Qte"), "#,##0.000"))
        rs.MoveNext
    Wend
End If
rs.Close

End Sub

Public Sub AfficheDetails_N(vNumero)
'Cette procedure est pour afficher une facture deja créer
Dim LOBJ_Fact As Facture
Dim LOBJ_BC As BonCarburant
Dim LOBJ_Bv As BonVidange
Dim LOBJ_PRep As PieceReparation
Dim rs As New Recordset

Lsv_Client.ListItems.Clear
Lsv_Totaux.ListItems.Clear
Lsv_Details.ListItems.Clear

NbrMO = 0
THT_MO = 0
TTC_TBR = 0
TT_RmsPiece = 0
Tva_MO = 0
RmsPiece = 0

Set LOBJ_Fact = New Facture
Set rs = LOBJ_Fact.Get_Details_PourAffich(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vNumero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
'remplir LSV_Client
If Not rs.EOF Then
    While Not rs.EOF
        Set itmX = Lsv_Client.ListItems.Add(, , rs("type"))
        itmX.SubItems(1) = rs("Numero")
        itmX.SubItems(2) = rs("dateop")
        itmX.SubItems(5) = CStr(Format(rs("Valeur"), "#,##0.000"))
        rs.MoveNext
    Wend
End If
rs.Close

'Produits carburant
Set LOBJ_BC = New BonCarburant
Set rs = LOBJ_BC.Get_BCByNumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vNumero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        Set itmX = Lsv_Details.ListItems.Add(, , "V")
        itmX.SubItems(1) = CStr(rs("Qte"))
        itmX.SubItems(2) = CStr(rs("energie"))
        itmX.SubItems(3) = CStr(Format(rs("PRIXHT"), "#,##0.000"))
        itmX.SubItems(4) = rs("Remise")
        itmX.SubItems(5) = rs("TVA")
        itmX.SubItems(6) = CStr(Format(rs("PRIXHT") * rs("Qte"), "#,##0.000"))
    rs.MoveNext
    Wend
End If
rs.Close

'' Produits vidange
Set LOBJ_Bv = New BonVidange
Set rs = LOBJ_Bv.Get_BVByNumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vNumero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = Lsv_Details.ListItems.Add(, , "A")
            itmX.SubItems(1) = CStr(rs("Qte"))
            itmX.SubItems(2) = CStr(rs("libelle"))
            itmX.SubItems(3) = CStr(Format(rs("PRIXHT"), "#,##0.000"))
            itmX.SubItems(4) = rs("Remise")
            itmX.SubItems(5) = rs("TVA")
            itmX.SubItems(6) = CStr(Format(rs("PRIXHT") * rs("Qte"), "#,##0.000"))
        rs.MoveNext
    Wend
End If
rs.Close

'Pièces de rèception
Set LOBJ_PRep = New PieceReparation
Set rs = LOBJ_PRep.Get_PRByNumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vNumero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        Set itmX = Lsv_Details.ListItems.Add(, , "P")
        itmX.SubItems(1) = CStr(rs("Qte"))
        itmX.SubItems(2) = CStr(rs("Designation"))
        itmX.SubItems(3) = CStr(Format(rs("PUHT"), "#,##0.000"))
        If Not IsNull(rs("Remise")) Then itmX.SubItems(4) = rs("Remise")
        If Not IsNull(rs("tva")) Then itmX.SubItems(5) = rs("tva")
        itmX.SubItems(6) = CStr(Format(rs("PUHT") * rs("Qte"), "#,##0.000"))
        If Not IsNull(rs("RemisePiece")) Then
            RmsPiece = CDbl(rs("RemisePiece"))
            TT_RmsPiece = TT_RmsPiece + (CDbl(itmX.SubItems(6)) - (CDbl(itmX.SubItems(6)) * rs("Remise") / 100)) * (CDbl(rs("RemisePiece")) / 100)
        End If
        rs.MoveNext
    Wend
End If
rs.Close

Dim rmsMO As Double
rmsMO = 0

Set rs = LOBJ_PRep.Get_MOTimbPRepFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vNumero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        If Not IsNull(rs("MOeuvre")) Then
            THT_MO = THT_MO + CDbl(rs("MOeuvre"))
            If rs("MOeuvre") <> 0 Then NbrMO = NbrMO + 1
            If Not IsNull(rs("RemisePiece")) Then
                TT_RmsPiece = TT_RmsPiece + CDbl(rs("MOeuvre")) * CDbl(rs("RemisePiece")) / 100
                rmsMO = rmsMO + (CDbl(rs("MOeuvre")) * CDbl(rs("RemisePiece")) / 100)
            End If
        End If
        If Not IsNull(rs("TVA_MOeuvre")) Then Tva_MO = Tva_MO + (CDbl(rs("MOeuvre")) - (CDbl(rs("MOeuvre")) * CDbl(rs("RemisePiece")) / 100)) * (CDbl(rs("TVA_MOeuvre")) / 100)
        If Not IsNull(rs("timbre")) Then TTC_TBR = TTC_TBR + CDbl(rs("timbre"))
    rs.MoveNext
    Wend
End If
rs.Close
    
  'Ajout ligne Main d'oeuvre dans la liste des détails
    Set itmX = Lsv_Details.ListItems.Add(, , "MO")
        itmX.SubItems(1) = NbrMO
        itmX.SubItems(2) = "Main d'oeuvre"
        itmX.SubItems(3) = CStr(Format(THT_MO, "#,##0.000"))
        itmX.SubItems(4) = CStr(Format(rmsMO, "#,##0.000"))
        itmX.SubItems(5) = CStr(Format(Tva_MO, "#,##0.000"))
        itmX.SubItems(6) = CStr(Format(THT_MO, "#,##0.000"))
        
 'Detail Bon Retour Pièce de rèception
Set rs = LOBJ_PRep.Get_BRByNumFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vNumero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        Set itmX = Lsv_Details.ListItems.Add(, , "X")
        itmX.SubItems(1) = CStr(rs("Qte"))
        itmX.SubItems(2) = CStr(rs("Designation"))
        itmX.SubItems(3) = CStr(Format(rs("PUHT"), "#,##0.000"))
        itmX.SubItems(4) = rs("Remise")
        itmX.SubItems(5) = rs("tva")
        itmX.SubItems(6) = CStr(Format(rs("PUHT") * rs("Qte"), "#,##0.000"))
    rs.MoveNext
    Wend
End If
rs.Close

Dim T
For T = 1 To Lsv_Client.ListItems.Count
    Lsv_Client.ListItems(T).Checked = True
Next T
Call AppCalcule
'Call AfficheDetails_PourCreation(txt_MatriculeStation.Text, cda_Debut.Value, cda_Fin.Value)


End Sub

Private Sub txt_Timbre_LostFocus()

If Len(txt_Timbre.Text) > 0 Then
    txt_Timbre.Text = Format(txt_Timbre.Text, "#,##0.000")
Else
    txt_Timbre.Text = Format(0, "#,##0.000")
End If
Call AppCalcule
End Sub

Private Function NbBonSelect(ByVal Lsv As ListView) As Integer

Dim i As Integer
Dim Compteur As Integer

 Compteur = 0
If Lsv.ListItems.Count > 0 Then
   
    For i = 1 To Lsv.ListItems.Count
        If Lsv.ListItems(i).Selected = True Then
            Compteur = Compteur + 1
        End If
    Next
    
End If
NbBonSelect = Compteur
End Function

Private Sub txt_Numero_GotFocus()
Call ViderZone(frmCreationFacture)
Lsv_Client.ListItems.Clear
Lsv_Totaux.ListItems.Clear
Lsv_Details.ListItems.Clear
End Sub

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Numero_LostFocus()
On Error GoTo Err

If Len(Trim(txt_Numero.Text)) > 0 Then Call AfficheRow(txt_Numero.Text)

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub
