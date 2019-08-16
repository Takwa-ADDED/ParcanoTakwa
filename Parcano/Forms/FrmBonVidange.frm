VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBonVidange 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bon de Vidange"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   12990
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Station"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1935
      Left            =   120
      TabIndex        =   46
      Top             =   3960
      Width           =   5775
      Begin SToolBox.SBiCombo cbo_MatriculeStation 
         Height          =   315
         Left            =   1680
         TabIndex        =   65
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   4935
         TabIndex        =   47
         Top             =   720
         Width           =   4935
         Begin VB.TextBox txt_rsocial 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   16
            Top             =   0
            Width           =   2895
         End
         Begin VB.TextBox txt_adresse 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   17
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txt_ville 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   18
            Top             =   720
            Width           =   2895
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
            TabIndex        =   50
            Top             =   0
            Width           =   1290
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
            Left            =   0
            TabIndex        =   49
            Top             =   360
            Width           =   780
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
            Left            =   0
            TabIndex        =   48
            Top             =   720
            Width           =   435
         End
      End
      Begin SToolBox.SCommand CmdFindStation 
         Height          =   360
         Left            =   4680
         TabIndex        =   51
         Top             =   240
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
         Picture         =   "FrmBonVidange.frx":0000
         ButtonType      =   1
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Station "
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
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vehicule"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1935
      Left            =   120
      TabIndex        =   34
      Top             =   1920
      Width           =   11655
      Begin SToolBox.SBiCombo cbo_Matricule 
         Height          =   315
         Left            =   1680
         TabIndex        =   66
         Top             =   360
         Width           =   2240
         _ExtentX        =   3942
         _ExtentY        =   556
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox Cbo_Conducteur 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   11415
         TabIndex        =   36
         Top             =   720
         Width           =   11415
         Begin VB.TextBox txt_Type 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   11
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txt_libelle 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   10
            Top             =   120
            Width           =   2175
         End
         Begin VB.TextBox txt_Compteur 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   0
            Width           =   1455
         End
         Begin VB.TextBox txt_Energie 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9240
            MaxLength       =   50
            TabIndex        =   15
            Top             =   600
            Width           =   1935
         End
         Begin SToolBox.SDateBox cda_FinAssur 
            Height          =   285
            Left            =   6120
            TabIndex        =   12
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_FinVisite 
            Height          =   285
            Left            =   6120
            TabIndex        =   13
            Top             =   600
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date fin assurance :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4440
            TabIndex        =   42
            Top             =   120
            Width           =   1740
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date fin visite :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   4440
            TabIndex        =   41
            Top             =   600
            Width           =   1320
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   40
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Matricule :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   39
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "CPT. Traffic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8040
            TabIndex        =   38
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Energie :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   8040
            TabIndex        =   37
            Top             =   600
            Width           =   780
         End
      End
      Begin SToolBox.SCommand cmdFindMatricule 
         Height          =   345
         Left            =   3960
         TabIndex        =   35
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         Picture         =   "FrmBonVidange.frx":0353
         ButtonType      =   1
      End
      Begin SToolBox.SCommand CmdFindConducteur 
         Height          =   435
         Left            =   8880
         TabIndex        =   44
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
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
         Picture         =   "FrmBonVidange.frx":06A6
         ButtonType      =   1
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conducteur"
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
         Height          =   240
         Left            =   4560
         TabIndex        =   45
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Immatriculation"
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
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Vidange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Vidange"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1935
      Left            =   6240
      TabIndex        =   32
      Top             =   3960
      Width           =   5775
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   5175
         TabIndex        =   33
         Top             =   240
         Width           =   5175
         Begin VB.TextBox txt_Ncompteur 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "M"
            Text            =   "0"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.PictureBox Pic_derVdg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   4335
            TabIndex        =   54
            Top             =   120
            Width           =   4335
            Begin VB.TextBox txt_KlmVidange 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               MaxLength       =   50
               TabIndex        =   64
               Tag             =   "M"
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txt_DerCompteurV 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   55
               Tag             =   "M"
               Top             =   0
               Width           =   1695
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NB KM Vidange :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   0
               TabIndex        =   63
               Top             =   480
               Width           =   1470
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dernier.compteur vidange :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   390
               Left            =   0
               TabIndex        =   56
               Top             =   0
               Width           =   2325
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nouv compteur Vidange"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   0
            TabIndex        =   57
            Top             =   1080
            Width           =   2235
         End
      End
   End
   Begin MSComctlLib.ListView grid 
      Height          =   1455
      Left            =   600
      TabIndex        =   19
      Top             =   7200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2566
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
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Numero"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lubrifiant"
         Object.Width           =   4833
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "THT"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "TVA"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "PrixTTC"
         Object.Width           =   2470
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4680
      ScaleHeight     =   375
      ScaleWidth      =   6135
      TabIndex        =   30
      Top             =   1440
      Width           =   6135
      Begin MSComCtl2.DTPicker dateOp 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   148570113
         CurrentDate     =   42875
      End
      Begin VB.Label cda_Create 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4800
         TabIndex        =   62
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Operation:"
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
         TabIndex        =   58
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Créaion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   22
         Top             =   0
         Width           =   1185
      End
   End
   Begin VB.TextBox txt_Numero 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox PIC_NFACT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4680
      ScaleHeight     =   495
      ScaleWidth      =   4455
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ce bon est inseré dans une facture N° : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -720
         TabIndex        =   23
         Top             =   120
         Width           =   4380
      End
      Begin VB.Label LBL_NFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1250"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   25
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox PIC_2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   10815
      TabIndex        =   26
      Top             =   6120
      Width           =   10815
      Begin VB.OptionButton VdgFiltre 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vidange avec filtre"
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton VdgSimple 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vidange simple"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3585
         ScaleHeight     =   615
         ScaleWidth      =   2895
         TabIndex        =   27
         Top             =   2520
         Width           =   2895
         Begin VB.TextBox txt_Valeur 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1560
            TabIndex        =   28
            Tag             =   "M"
            Text            =   "0,000"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valeur TTC:"
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
            Height          =   240
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Détails Nouveau Vidange"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   210
         Left            =   150
         TabIndex        =   53
         Top             =   120
         Width           =   2265
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4680
      Top             =   240
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   9240
      TabIndex        =   6
      Top             =   120
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
      Picture         =   "FrmBonVidange.frx":09F9
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8520
      TabIndex        =   31
      Top             =   120
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
      Picture         =   "FrmBonVidange.frx":0B7B
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   120
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
      Picture         =   "FrmBonVidange.frx":0ECE
   End
   Begin SToolBox.SCommand cmdFindNumero 
      Height          =   495
      Left            =   3600
      TabIndex        =   21
      Top             =   1320
      Width           =   420
      _ExtentX        =   741
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
      Picture         =   "FrmBonVidange.frx":1221
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   8160
      TabIndex        =   8
      Top             =   120
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
      Picture         =   "FrmBonVidange.frx":1574
   End
   Begin SToolBox.SCommand CmdPrint 
      Height          =   495
      Left            =   9600
      TabIndex        =   9
      Top             =   120
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
      Picture         =   "FrmBonVidange.frx":16F6
   End
   Begin VB.Label Lbl_UserSaisi 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   11160
      TabIndex        =   61
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Lbl_user 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BV saisi par :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9840
      TabIndex        =   60
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bon de sortie vidange"
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
      Left            =   360
      TabIndex        =   59
      Top             =   360
      Width           =   3540
   End
   Begin VB.Image PicBox_Header 
      Height          =   1575
      Left            =   -120
      Picture         =   "FrmBonVidange.frx":1A49
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   12855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro bon"
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
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "FrmBonVidange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim thekey As Integer
Dim theshift As Integer

Private Sub cbo_conducteur_LostFocus()
    Call ExistDonneeCbo(cbo_conducteur)
End Sub

'Impression du bonVidange
Private Sub CmdPrint_Click()

Dim F As Form
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

If MsgBox("Imprimer ce bon        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        Call .PrintOutAndApercu_BV2(0)
        .Show
    End With
End If

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub Form_Load()

On Error GoTo Err
CmdDelete.Enabled = False
CmdSave.Enabled = False
CmdPrint.Enabled = False

Me.Width = 11715
Me.Height = 8625
cda_Create.Caption = Date
dateOp.Value = Date
Me.Move 500, 500
Call Affiche_Personnel_Combo(cbo_conducteur)
Call Affiche_Matricule_SBCombo(cbo_Matricule)
Call Affiche_StatCarb_SBCombo(cbo_MatriculeStation)
Me.WindowState = 2
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'Retourner le code du Conducteur selon son nom
Private Function RET_CODE_CONDUCTEUR(txt As String) As String

Dim LOBJ_Personnel As personnel
Dim rs As Recordset
' Initialisation
RET_CODE_CONDUCTEUR = ""
Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.GetCODE_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs.EOF Then
    RET_CODE_CONDUCTEUR = rs("Code")
End If
rs.Close
End Function

'Retourner prix d'energie suivant son libelle
Private Function RET_PRIX_ENERGIE(txt As String) As Double

Dim LOBJ_Energie As Energie
Dim rs As New Recordset
' Initialisation
RET_PRIX_ENERGIE = 0

Set LOBJ_Energie = New Energie
Set rs = LOBJ_Energie.Get_PrixEnergie(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt)
If Not rs.EOF Then
    RET_PRIX_ENERGIE = rs("prix")
End If
rs.Close
End Function

'Click sur bouton ADD pour l'insertion d'un nouveau bonVidange
Private Sub CmdAdd_Click()

On Error GoTo Err

Dim LOBJ_BonVidange As BonVidange
Dim LOBJ_Personnel As personnel

Set LOBJ_BonVidange = New BonVidange
Set LOBJ_Personnel = New personnel
'Verifier le droit d'accès pour insertion d'un bonVidange
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "INS_BV", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmBonVidange)
grid.ListItems.Clear
txt_Numero.Text = "Auto"
cda_Create.Caption = Date
dateOp.Value = Date
Lbl_UserSaisi.Caption = LStr_NameUser
LBL_NFact.Caption = ""
Timer1.Enabled = False

PIC_NFACT.Visible = False
PIC_2.Enabled = True
cbo_conducteur.Text = " "
cbo_Matricule.Text = " "
cbo_Matricule.Enabled = True
cbo_conducteur.Enabled = True
cbo_MatriculeStation.Enabled = True
txt_Ncompteur.Enabled = True
cmdFindMatricule.Enabled = True
cmdFindConducteur.Enabled = True
CmdFindStation.Enabled = True
cmdFindNumero.Visible = False
VdgFiltre.Value = False
VdgSimple.Value = False
cbo_Matricule.SetFocus
CmdSave.Enabled = True
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'Suppression d'un bonVidange
Private Sub CmdDelete_Click()

Dim LOBJ_BonVidange As BonVidange
Dim LOBJ_Station As Station
Dim LOBJ_Personnel As personnel
Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset
Dim VCode As String
Dim numMax
Dim derCmptV
On Error GoTo Err

Set LOBJ_BonVidange = New BonVidange

'si le bonV est déjà inséré dans une facture donc le MAJ ou la suppression est impossible
If PIC_NFACT.Visible = True Then
    MsgBox "Maj impossible", vbInformation
    Exit Sub
End If

If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If
Set LOBJ_Personnel = New personnel
'Vérifier le droit d'accès de l'utilisateur pour la suppression d'un bonvidange
If txt_Numero.Text <> "Auto" Then
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Supp_BV", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
End If

If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "bon de vidange", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    VCode = txt_Numero.Text
    
    Set rs = LOBJ_BonVidange.Get_NumMaxBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cbo_Matricule.Text)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
       numMax = rs("maxNum")
    End If
    rs.Close
    If Val(txt_Numero.Text) = Val(numMax) Then
   ' dernier Compteur Vidange du bon avant dernier
        Set rs = LOBJ_BonVidange.Get_AnComptVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Numero.Text, cbo_Matricule.Text)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        If Not rs.EOF Then
           derCmptV = rs("maxCpt")
        End If
        rs.Close
        Set Lobj_Vehicule = New VEHICULE
        Call Lobj_Vehicule.Update_ComptVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, derCmptV, cbo_Matricule.Text)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
    End If
    'Suppression du BV de la table Det_bonvidange
    Call LOBJ_BonVidange.DeleteBV_Lub(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    'Suppression du BV de la table Ass_BonVidange
    Call LOBJ_BonVidange.DeleteBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode, LInt_UserId)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Station = New Station
    ' MAJ du nombre des bon de la station : supprimer 1
    Call LOBJ_Station.UpdateNBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, -1, cbo_MatriculeStation.FirstValue)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    txt_Numero.SetFocus
End If
Call ViderZone(FrmBonVidange)
grid.ListItems.Clear
txt_Numero.SetFocus
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'recherche d'un bonvidange
Private Sub CmdFind_Click()

On Error GoTo Err

If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If
cmdFindNumero.Visible = True
Unload FrmFind
With FrmFind
    .StrSource = "BonVidange2"
    .Show vbModal
End With
CmdFind.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub cmdFindConducteur_Click()

On Error GoTo Err

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "PersonnelVidange2"
    .Show vbModal
End With

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub cmdFindMatricule_Click()

On Error GoTo Err

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "VéhiculeVidange2"
    .Show vbModal
End With
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub cmdFindNumero_Click()

On Error GoTo Err

Unload FrmFind
With FrmFind
    .StrSource = "BonVidange2"
    .Show vbModal
End With
cmdFindNumero.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

'Afficher la liste des stations de vidange
Private Sub CmdFindStation_Click()

On Error GoTo Err

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "StationVidange"
    .Show vbModal
End With
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

'Enregistrement de l'insertion d'un nouveau bonVidange ou MAJ d'un ancien BV
Private Sub CmdSave_Click()

Dim LOBJ_BonVidange As BonVidange
Dim LOBJ_Personnel As personnel
Dim rs As New Recordset
    
'On Error GoTo Err
If PIC_NFACT.Visible = True Or Timer1.Enabled = True Then
    MsgBox "Maj impossible", vbInformation
    Exit Sub
End If

If Left(CheckMandatory(FrmBonVidange), 1) = 1 Then
   Exit Sub
End If

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cbo_conducteur.Text)
If ErrNumber <> 0 Then
   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
   ErrNumber = 0
   Exit Sub
End If
If rs.EOF Then
    MsgBox "Ce conducteur n'existe pas ", vbInformation
    cbo_conducteur.SetFocus
    Exit Sub
    rs.Close
End If

If VdgSimple.Value = False And VdgFiltre.Value = False Then
    MsgBox "Choisir un type de vidange  ", vbInformation
    Exit Sub
End If

If grid.ListItems.Count = 0 And VdgFiltre.Value = True Then
    MsgBox "Véhicule n'a pas des lubrifiants de vidange. " & vbNewLine & "Choisir les détails de vidange avant de saisir le bon. ", vbInformation
    Exit Sub
End If

If txt_klmVidange.Text = "" Or Val(txt_klmVidange.Text) < 1000 Or Val(txt_klmVidange.Text) > 10000 Then
    MsgBox "Vérifier le nombre de kilomètre de vidange  ", vbInformation
    Exit Sub
End If

If txt_Ncompteur.Text = "" Or txt_Ncompteur.Text = "0" Then
    MsgBox "Saisir le nouveau compteur de vidange  ", vbInformation
    Exit Sub
End If

If cbo_conducteur.Text = "" Or cbo_conducteur.Text = " " Then
    MsgBox "Conducteur obligatoire      ", vbInformation
    Exit Sub
End If

If txt_libelle.Text = "" Or txt_libelle.Text = " " Then
    MsgBox "Vehicule obligatoire", vbInformation
    Exit Sub
End If

If txt_rsocial.Text = "" Or txt_rsocial.Text = " " Then
    MsgBox "Station obligatoire      ", vbInformation
    Exit Sub
End If

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

Set LOBJ_BonVidange = New BonVidange
Set LOBJ_Personnel = New personnel
'Vérification du droit de MAJ du bon
If txt_Numero.Text <> "Auto" And txt_Numero.Text <> "" Then
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "MAJ_BV", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
    Call ModifierBV
End If

If txt_Numero.Text = "Auto" Then
    Call AjouterBV
End If

    txt_DerCompteurV.Text = txt_Ncompteur.Text

If MsgBox("Enregistrement terminé avec succé  " & vbNewLine & "Imprimer ce bon        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        Call .PrintOutAndApercu_BV2(0)
        .Show
    End With
End If
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "MAJ_BV", LInt_UserId) Then
    Call ViderZone(FrmBonVidange)
End If
cmdFindNumero.Visible = True
Exit Sub
Err:
    MsgBox Err.Source, vbInformation
End Sub

'Insertion d'un nouveau Bonvidange et LubV
Private Sub AjouterBV()

Dim LOBJ_BonVidange As BonVidange
Dim LOBJ_Station As Station
Dim Lobj_Vehicule As VEHICULE
Dim LRs_NewRecord As New Recordset
Dim LInt_NumCompteur As Long
Dim i As Long
Dim NBV As Long
Dim VCode

On Error GoTo Err
'Incrementation du compteur du numero du bon dans t_BonVidange
LInt_NumCompteur = Return_CountBV() + 1
VCode = Format(LInt_NumCompteur, "00000")
txt_Numero.Text = VCode
Set LOBJ_BonVidange = New BonVidange
Set Lobj_Vehicule = New VEHICULE
Set LOBJ_Station = New Station
'Incrémenter le nombre des bons pour cet véhicule
CNB.BeginTrans
Call LOBJ_Station.UpdateNBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, 1, cbo_MatriculeStation.FirstValue)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

NBV = Return_NBV(cbo_MatriculeStation.FirstValue)

Set LRs_NewRecord = CreateEmptyRS_BV
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = VCode
    .Fields("DateDoc") = CDate(cda_Create.Caption)
    .Fields("Vehicule") = return_CDVehi(txt_libelle.Text)
    .Fields("Station") = cbo_MatriculeStation.FirstValue
    .Fields("Conducteur") = RET_CODE_CONDUCTEUR(cbo_conducteur.Text)
    .Fields("valeur") = CDbl(txt_Valeur.Text)
    .Fields("heure") = Format(Time, "hh:mm")
    .Fields("NBC") = NBV
    .Fields("dateOp") = CDate(dateOp.Value)
    .Fields("CompteurVidange") = txt_Ncompteur.Text
    .Fields("NBKLMvid") = txt_klmVidange.Text
    .Fields("UserInsert") = LInt_UserId
End With
Call LOBJ_BonVidange.Insert_BV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

'Insertion enregistrement details
If VdgFiltre.Value = True Then
    Set LRs_NewRecord = CreateEmptyRS_LubBV()
    For i = 1 To grid.ListItems.Count
        With LRs_NewRecord
            .AddNew
            .Fields("Numero") = VCode
            .Fields("Libelle") = grid.ListItems(i).SubItems(1)
            .Fields("THT") = CDbl(grid.ListItems(i).SubItems(2))
            .Fields("TVA") = CDbl(grid.ListItems(i).SubItems(3))
            .Fields("prixTTC") = CDbl(grid.ListItems(i).SubItems(4))
        End With
    Next
Else
    Set LRs_NewRecord = CreateEmptyRS_LubBV()
    With LRs_NewRecord
        .AddNew
        .Fields("Numero") = VCode
        .Fields("Libelle") = "Vidange simple gratuit"
        .Fields("THT") = 0
        .Fields("TVA") = 0
        .Fields("prixTTC") = 0
    End With
End If
Call LOBJ_BonVidange.Insert_LubBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

'Changer la valeur du dernier compteur pour ce véhicule
Call Lobj_Vehicule.Update_ComptVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Val(txt_Ncompteur.Text), return_CDVehi(txt_libelle.Text))
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
CNB.CommitTrans

Exit Sub
Err:
    MsgBox Err.Source, vbInformation
    CNB.RollbackTrans
End Sub

'Modification d'un BonVidange
Private Sub ModifierBV()

Dim LOBJ_BonVidange As BonVidange
Dim Lobj_Vehicule As VEHICULE
Dim LRs_NewRecord As New Recordset
Dim i As Long
Dim NBV As Long

CNB.BeginTrans
Set LOBJ_BonVidange = New BonVidange
Set Lobj_Vehicule = New VEHICULE
NBV = Return_NBV(cbo_MatriculeStation.FirstValue)
Set LRs_NewRecord = CreateEmptyRS_BV
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    .Fields("Conducteur") = RET_CODE_CONDUCTEUR(cbo_conducteur.Text)
    .Fields("valeur") = CDbl(txt_Valeur.Text)
    .Fields("heure") = Format(Time, "hh:mm")
    .Fields("NBC") = NBV
    .Fields("dateop") = Date
    .Fields("CompteurVidange") = txt_Ncompteur.Text
    .Fields("NBKLMvid") = txt_klmVidange.Text
    .Fields("UserUpdate") = LInt_UserId
End With
Call LOBJ_BonVidange.Update_BV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

Call LOBJ_BonVidange.Delete_DetBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Numero.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
'Insertion des nouveaux details dans T_LubV
If VdgFiltre.Value = True Then
    Set LRs_NewRecord = CreateEmptyRS_LubBV()
    For i = 1 To grid.ListItems.Count
        With LRs_NewRecord
            .AddNew
            .Fields("Numero") = txt_Numero.Text
            .Fields("Libelle") = grid.ListItems(i).SubItems(1)
            .Fields("THT") = CDbl(grid.ListItems(i).SubItems(2))
            .Fields("TVA") = CDbl(grid.ListItems(i).SubItems(3))
            .Fields("prixTTC") = CDbl(grid.ListItems(i).SubItems(4))
        End With
    Next
ElseIf VdgSimple.Value = True Then
    Set LRs_NewRecord = CreateEmptyRS_LubBV()
        With LRs_NewRecord
            .AddNew
            .Fields("Numero") = txt_Numero.Text
            .Fields("Libelle") = "Vidange simple gratuit"
            .Fields("THT") = 0
            .Fields("TVA") = 0
            .Fields("prixTTC") = 0
        End With
End If
Call LOBJ_BonVidange.Insert_LubBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

'Changer la valeur du dernier compteur pour ce véhicule
Call Lobj_Vehicule.Update_ComptVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Val(txt_Ncompteur.Text), return_CDVehi(txt_libelle.Text))
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
CNB.CommitTrans
End Sub

'Retourne le Numero Max du table Ass_BonVidange
Private Function Return_CountBV() As Long

Dim rs As New Recordset
Dim LOBJ_BonVidange As BonVidange

Return_CountBV = 0
Set LOBJ_BonVidange = New BonVidange
Set rs = LOBJ_BonVidange.Get_MaxNumBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB)

If Not rs.EOF Then
    Return_CountBV = CLng(rs("maxNum"))
End If
rs.Close
End Function

'Retourne le nombre des bon de vidange pour une station
Private Function Return_NBV(VCode As String) As Long

Dim LOBJ_Station As Station
Dim rs As New Recordset
' Initialisation
Return_NBV = 0

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.Get_NBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)

If Not rs.EOF Then
    Return_NBV = rs("numbvdg")
End If
rs.Close

End Function

'Appelé par AfficheRow
'Affiche les détails d'un véhicule ainsi que les détails du dernier vidange de ce véhicule
Public Sub AfficheRow_Vehicule(ByVal VCode As String)

Dim LOBJ_BonVidange As BonVidange
Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset
Dim rs1 As New Recordset

txt_libelle.Text = ""
txt_Type.Text = ""
txt_Energie.Text = ""
cda_FinAssur.Text = ""
cda_FinVisite.Text = ""

txt_DerCompteurV.Text = ""
    
Set LOBJ_BonVidange = New BonVidange
Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
'    cbo_Matricule.Text = rs("Code")
'    cbo_Matricule.SecondValue = rs("matricule")
    txt_Compteur.Text = CompteurVehicule(rs("Code"))
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
    If Not IsNull(rs("DAteFinAssur")) Then cda_FinAssur.Text = rs("DAteFinAssur")
    If Not IsNull(rs("DAteFinVisite")) Then cda_FinVisite.Text = rs("DAteFinVisite")
    If Not IsNull(rs("NbKlmVidange")) Then txt_klmVidange.Text = rs("NbKlmVidange")
    'Dernier bonVidange Max(Numero) pour un véhicule
    Set rs1 = LOBJ_BonVidange.Get_DerBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Code"))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs1.EOF Then
        If Not IsNull(rs1("CompteurVidange")) Then txt_DerCompteurV.Text = rs1("CompteurVidange")
    End If
    rs1.Close
    'Max entre txt_compteur et txt_DerCompteurV
    If Val(txt_DerCompteurV.Text) > Val(txt_Compteur.Text) Then
        txt_Ncompteur.Text = txt_DerCompteurV.Text
    Else
        txt_Ncompteur.Text = txt_Compteur.Text
    End If
    If txt_Numero.Text = "Auto" Then
'        cbo_conducteur.SetFocus
    End If
Else
    MsgBox "Code introuvable, vérifier votre saisie.", vbInformation, App.ProductName
    cbo_Matricule.Text = ""
End If
rs.Close

End Sub

Public Sub AfficheRow_Vehicule_sansPrix(ByVal VCode As String)

Dim LOBJ_BonVidange As BonVidange
Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset
Dim rs1 As New Recordset

Set LOBJ_BonVidange = New BonVidange
Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.GetVehiculeByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    cbo_Matricule.Text = rs("Code")
'    cbo_Matricule.SecondValue = rs("matricule")
    txt_Compteur.Text = CompteurVehicule(rs("Code"))
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
    If Not IsNull(rs("DAteFinAssur")) Then cda_FinAssur.Text = rs("DAteFinAssur")
    If Not IsNull(rs("DAteFinVisite")) Then cda_FinVisite.Text = rs("DAteFinVisite")
    If Not IsNull(rs("NbKlmVidange")) Then txt_klmVidange.Text = rs("NbKlmVidange")
    Set rs1 = LOBJ_BonVidange.Get_DerBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("code"))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs1.EOF Then
        If Not IsNull(rs1("NBKLMvid")) Then txt_klmVidange.Text = rs1("NBKLMvid")
        If Not IsNull(rs1("CompteurVidange")) Then txt_DerCompteurV.Text = rs1("CompteurVidange")
    End If
    'Max entre txt_compteur et txt_DerCompteurV
    If Val(txt_DerCompteurV.Text) > Val(txt_Compteur.Text) Then
        txt_Ncompteur.Text = txt_DerCompteurV.Text
    Else
        txt_Ncompteur.Text = txt_Compteur.Text
    End If
End If
rs.Close

'Call AfficheRow_Lubr(vcode)
'Call AppCalcul
End Sub
Private Sub AfficheRow_Lubr(ByVal VCode As String)

Dim Lobj_Vehicule As VEHICULE
Dim LOBJ_Lub As Produit_Lubrifiant
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim itmX
Dim prix As Double
    
 grid.ListItems.Clear
 
    Set LOBJ_Lub = New Produit_Lubrifiant
    Set Lobj_Vehicule = New VEHICULE
    Set rs = Lobj_Vehicule.Get_VehVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set Lobj_Vehicule = Nothing
        
    If Not rs.EOF Then
        While Not rs.EOF
          
            Set rs1 = LOBJ_Lub.Get_ProdLubBycode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Lubrifiant"))
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            If Not rs1.EOF Then
                Set itmX = grid.ListItems.Add(, , CStr(rs("Lubrifiant")))
                itmX.SubItems(1) = CStr(rs1("Libelle"))
                itmX.SubItems(2) = CStr(Format(rs1("prixht"), "#,##0.000"))
                itmX.SubItems(3) = CStr(Format(rs1("tva"), "#,##0.000"))
                prix = (rs1("prixht")) * (rs1("TVA")) / 100
                prix = prix + (rs1("prixht"))
                itmX.SubItems(4) = CStr(Format(prix, "#,##0.000"))
            End If
            rs1.Close
            
            rs.MoveNext
        Wend
    Else
        MsgBox "Véhicule n'a pas encore des détails de vidange, " & vbNewLine & "Veuillez choisir les lubrifiants de vidange. ", vbInformation
    End If
Set LOBJ_Lub = Nothing
Set Lobj_Vehicule = Nothing
rs.Close
End Sub

Public Sub AfficheRow_Station(ByVal VCode As String)

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.GetStatByCodeLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    cbo_MatriculeStation.Text = VCode
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
Else
    MsgBox "Code introuvable", vbInformation
    cbo_MatriculeStation.SetFocus
    Exit Sub
End If

End Sub
Private Sub EnabDis(ByVal bl As Boolean)

CmdDelete.Enabled = bl
CmdSave.Enabled = bl
CmdPrint.Enabled = bl
cbo_MatriculeStation.Enabled = bl
CmdFindStation.Enabled = bl
cmdFindConducteur.Enabled = bl
cbo_conducteur.Enabled = bl
cbo_Matricule.Enabled = bl
cmdFindMatricule.Enabled = bl
txt_Numero.Enabled = bl
txt_Ncompteur.Enabled = bl

End Sub

'Affiche les détails d'un vidange d'un véhicule
'appelé lors d'un DbClick sur le grid dans FrmFind
Public Sub AfficheRow(ByVal VCode As String)

Dim LOBJ_BonVidange As BonVidange
Dim rs As New Recordset

Set LOBJ_BonVidange = New BonVidange
Call ViderZone(FrmBonVidange)
grid.ListItems.Clear
PIC_NFACT.Visible = False
Timer1.Enabled = False

Set rs = LOBJ_BonVidange.Get_BV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
If rs("Supp") = "O" Then
    MsgBox "ce bon de vidange a été supprimé par " & Get_NameUserByCode(rs("UserDelete")), vbInformation
    Call EnabDis(False)
Else
    Call EnabDis(True)
    cbo_Matricule.Enabled = False
    cbo_MatriculeStation.Enabled = False
    cmdFindMatricule.Enabled = False
    CmdFindStation.Enabled = False
End If
    'Charge
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("VEHICULE")) Then txt_libelle.Text = rs("VEHICULE")
    If Not IsNull(rs("STATION")) Then txt_Type.Text = rs("STATION")
    If Not IsNull(rs("DATEDOC")) Then cda_Create.Caption = rs("DATEDOC")
    If Not IsNull(rs("dateop")) Then dateOp.Value = rs("dateop")
    If Not IsNull(rs("VALEUR")) Then txt_Valeur.Text = Format(rs("VALEUR"), "#,##0.000")
    cbo_Matricule.Text = rs("VEHICULE")
    Call AfficheRow_Vehicule(rs("VEHICULE"))
    Call AfficheRow_Station(rs("STATION"))
    Call AfficheRow_Conducteur(rs("CONDUCTEUR"))
    If Not IsNull(rs("CompteurVidange")) Then
        txt_DerCompteurV.Text = rs("CompteurVidange")
        txt_Ncompteur.Text = rs("CompteurVidange")
    End If
    If Not IsNull(rs("NBKLMvid")) Then txt_klmVidange.Text = rs("NBKLMvid")
    If Not IsNull(rs("UserInsert")) Then Lbl_UserSaisi.Caption = Get_NameUserByCode(rs("UserInsert"))
    If Val(txt_Valeur.Text) = 0 Then
        VdgSimple.Value = True
    Else
        VdgFiltre.Value = True
    End If
    
    Call AfficheRow_Lubrifiant_BV(txt_Numero.Text)
    
    If rs("Transf") = "O" Then
        LBL_NFact.Caption = rs("NumFact")
        PIC_NFACT.Visible = True
        Call Timer1_Timer
        Call EnabDis(False)
    Else
        PIC_NFACT.Visible = False
        Timer1.Enabled = False
    End If

End If
rs.Close

End Sub

Public Sub AfficheRow_Conducteur(ByVal VCode As String)

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Libelle")) Then cbo_conducteur.Text = rs("Libelle")
End If
rs.Close

End Sub

'Afficher les détails d'un bonvidange
Public Sub AfficheRow_Lubrifiant_BV(ByVal VCode As String)

Dim LOBJ_DetBV As BonVidange
Dim rs As New Recordset

grid.ListItems.Clear

Set LOBJ_DetBV = New BonVidange
Set rs = LOBJ_DetBV.Get_Lub_BV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        Set itmX = grid.ListItems.Add(, , CStr(rs("Numero")))
        itmX.SubItems(1) = CStr(rs("Libelle"))
        itmX.SubItems(2) = CStr(Format(rs("THT"), "#,##0.000"))
        itmX.SubItems(3) = CStr(Format(rs("TVA"), "#,##0.000"))
        itmX.SubItems(4) = CStr(Format(rs("prixTTC"), "#,##0.000"))
        rs.MoveNext
    Wend
End If
rs.Close
Call AppCalcul

End Sub

'Dernier compteur du véhicule en entrant (ficheTraffic)
Public Function CompteurVehicule(ByVal VCode As String) As String

Dim Lobj_Vehicule As VEHICULE
Dim rs1 As New Recordset
Dim Name_Tab As String

CompteurVehicule = "0"
Set Lobj_Vehicule = New VEHICULE
Name_Tab = "FicheTraffic"
Set rs1 = Lobj_Vehicule.Get_DerCompt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Name_Tab, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs1.EOF Then
    If Not IsNull(rs1("maxCpt")) Then CompteurVehicule = rs1("maxCpt")
End If
rs1.Close
    
End Function

Private Sub Form_Resize()
    Dim WidthForm As Integer
    WidthForm = Frm_Main.ACB_Main.Width
    PicBox_Header.Width = WidthForm - 1000
    CmdAdd.Left = WidthForm - 5500
    CmdDelete.Left = WidthForm - 5100
    CmdFind.Left = WidthForm - 4700
    CmdSave.Left = WidthForm - 4300
    CmdPrint.Left = WidthForm - 3900
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = True
Timer1.Interval = 600
If PIC_NFACT.Visible = True Then
    PIC_NFACT.Visible = False
Else
    PIC_NFACT.Visible = True
End If
End Sub

Private Sub Cbo_Conducteur_Change()

Dim i As Integer, start As Integer
Dim ShiftDown As Boolean
Dim CtrlDown As Boolean
Dim AltDown As Boolean

ShiftDown = (theshift And vbShiftMask) > 0
CtrlDown = (theshift And vbCtrlMask) > 0
AltDown = (theshift And vbAltMask) > 0
If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
 
Else
    start = Len(cbo_conducteur.Text)
    For i = 0 To cbo_conducteur.ListCount - 1
        If Left(cbo_conducteur.List(i), start) = cbo_conducteur.Text Then
            cbo_conducteur.Text = cbo_conducteur.List(i)
        End If
    Next
    cbo_conducteur.SelStart = start
    cbo_conducteur.SelLength = Len(cbo_conducteur.Text)
End If
End Sub

Private Sub Cbo_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Cbo_Conducteur_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub

Private Sub cbo_Matricule_Change()

Dim i As Integer, start As Integer
Dim ShiftDown As Boolean
Dim CtrlDown As Boolean
Dim AltDown As Boolean

ShiftDown = (theshift And vbShiftMask) > 0
CtrlDown = (theshift And vbCtrlMask) > 0
AltDown = (theshift And vbAltMask) > 0
If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then

Else
    start = Len(cbo_Matricule.Text)
    For i = 0 To cbo_Matricule.ListCount - 1
        If Left(cbo_Matricule.List(i), start) = cbo_Matricule.Text Then
            cbo_Matricule.Text = cbo_Matricule.List(i)
        End If
    Next

    cbo_Matricule.SetFocus
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

'On Error GoTo erreur
'Dim i As Integer
'Dim Msg ' Déclare la variable.
'' Définit le texte du message.
'Msg = "Voulez-vous vraiment quitter?"
'' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
'If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'   Cancel = True
'Else
'Unload Me
'End If
'
'Exit Sub
'erreur:
'   MsgBox Err.Description, 48
End Sub

Private Sub Cbo_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode =    vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Cbo_Matricule_LostFocus()

On Error GoTo Err

If Len(Trim(cbo_Matricule.Text)) > 0 Then
    If Trim(cbo_Matricule.FirstValue) <> "" Then Call AfficheRow_Vehicule(cbo_Matricule.FirstValue)
End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub cbo_MatriculeStation_LostFocus()
    If Len(Trim(cbo_MatriculeStation.Text)) > 0 Then
        If IsNumeric(cbo_MatriculeStation.FirstValue) Then Call AfficheRow_Station(cbo_MatriculeStation.FirstValue)
    End If
End Sub

Private Sub txt_Ncompteur_GotFocus()
On Error Resume Next
txt_Ncompteur.SelLength = Len(txt_Ncompteur.Text)
End Sub

Private Sub txt_Ncompteur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Ncompteur_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_Ncompteur_LostFocus()

On Error GoTo Err

    If Val(txt_Ncompteur.Text) < Val(txt_DerCompteurV.Text) Then
        MsgBox "Nouveau CompteurVidange invalid, veuillez vérifier le compteur saisi", vbInformation
        txt_Ncompteur.SetFocus
        Exit Sub
    End If

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_Numero_GotFocus()

On Error GoTo Err

Call ViderZone(FrmBonVidange)
grid.ListItems.Clear
PIC_2.Enabled = True
PIC_NFACT.Visible = False

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

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

Private Sub cbo_Matricule_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub

Private Sub cbo_MatriculeStation_Click()
If Len(Trim(cbo_MatriculeStation.Text)) > 0 Then
    If IsNumeric(cbo_MatriculeStation.FirstValue) Then Call AfficheRow_Station(cbo_MatriculeStation.FirstValue)
End If

End Sub

Private Sub cbo_Matricule_Click()
If Len(Trim(cbo_Matricule.Text)) > 0 Then
    If Trim(cbo_Matricule.FirstValue) <> "" Then Call AfficheRow_Vehicule(cbo_Matricule.FirstValue)
End If

End Sub

Private Sub VdgFiltre_Click()

If VdgFiltre.Value = True Then
    If cbo_Matricule.Text = "" And txt_libelle = "" Then
        MsgBox "Véhicule obligatoire ", vbInformation
        Exit Sub
    End If
    Call AfficheRow_Lubr(cbo_Matricule.FirstValue)
    Call AppCalcul

End If
End Sub

Private Sub VdgSimple_Click()
If VdgSimple.Value = True Then
    grid.ListItems.Clear
    txt_Valeur.Text = 0
    
    Set itmX = grid.ListItems.Add(, , CStr(txt_Numero.Text))
        itmX.SubItems(1) = CStr("Vidange simple")
        itmX.SubItems(2) = CStr(Format(0, "#,##0.000"))
        itmX.SubItems(3) = CStr(Format(0, "#,##0.000"))
        itmX.SubItems(4) = CStr(Format(0, "#,##0.000"))
        txt_Valeur.Text = Format(0, "#,##0.000")
End If
End Sub

Public Sub AppCalcul()

Dim i
Dim Valeur As Double
Dim TotTTC As Double

'Parcourir la liste pour calculer les sommes
For i = 1 To grid.ListItems.Count
    TotTTC = grid.ListItems(i).SubItems(4)
    Valeur = Valeur + TotTTC
Next
txt_Valeur.Text = Format(Valeur, "#,##0.000")

End Sub

Private Sub ExistData(ByVal cbo As ComboBox)

On Error GoTo Err

    Dim RCount As Integer, i As Integer, Existe As Boolean, tcbo As String
    
    RCount = cbo.ListCount
    tcbo = cbo.Text
    
    For i = 0 To RCount - 1
        cbo.ListIndex = i
        If tcbo = cbo.Text Then
            Existe = True
            Exit For
        Else
            Existe = False
        End If
    Next i
    If i = RCount Then
        If Existe = False Then
            MsgBox "Saisie non Valide!...     ", vbExclamation, App.ProductName
            cbo.Text = ""
            cbo.SetFocus
            txt_rsocial.Text = ""
            txt_adresse.Text = ""
            txt_ville.Text = ""
            Exit Sub
        End If
    End If
    Existe = False
    
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
