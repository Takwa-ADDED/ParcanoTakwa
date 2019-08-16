VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form Frm_GestionConge 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion des congés"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   12660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   12660
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab STab_conge 
      Height          =   6855
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Gestion des congés"
      TabPicture(0)   =   "Frm_GestionConge.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Recherche"
      TabPicture(1)   =   "Frm_GestionConge.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   -74880
         ScaleHeight     =   6315
         ScaleWidth      =   11235
         TabIndex        =   22
         Top             =   360
         Width           =   11295
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H00000000&
            Height          =   4815
            Left            =   2040
            ScaleHeight     =   4785
            ScaleWidth      =   6825
            TabIndex        =   28
            Top             =   1440
            Width           =   6855
            Begin SToolBox.SGrid grid 
               Height          =   4695
               Left            =   0
               TabIndex        =   29
               Top             =   0
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   8281
               RowMode         =   -1  'True
               BackgroundPictureHeight=   0
               BackgroundPictureWidth=   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   2
               DisableIcons    =   -1  'True
               MaxVisibleRows  =   0
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   0
            ScaleHeight     =   1215
            ScaleWidth      =   11295
            TabIndex        =   24
            Top             =   0
            Width           =   11295
            Begin SToolBox.SOptionButton Opt_Sup 
               Height          =   255
               Left            =   8760
               TabIndex        =   11
               Top             =   120
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
               BackStyle       =   0
               Caption         =   "SOptionButton1"
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
            Begin SToolBox.SOptionButton OptNSup 
               Height          =   255
               Left            =   6720
               TabIndex        =   10
               Top             =   120
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
               BackStyle       =   0
               Caption         =   "SOptionButton1"
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
            Begin SToolBox.SBiCombo Scbo_Conduc 
               Height          =   315
               Left            =   1680
               TabIndex        =   12
               Top             =   720
               Width           =   3015
               _ExtentX        =   5318
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
            Begin MSComCtl2.DTPicker cda_Deb 
               Height          =   375
               Left            =   1680
               TabIndex        =   8
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   14737632
               Format          =   149291009
               CurrentDate     =   42860
            End
            Begin MSComCtl2.DTPicker cda_au 
               Height          =   375
               Left            =   4200
               TabIndex        =   9
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   14737632
               Format          =   149291009
               CurrentDate     =   42860
            End
            Begin VB.Image CmdPrint 
               Height          =   495
               Left            =   9120
               Picture         =   "Frm_GestionConge.frx":0038
               Stretch         =   -1  'True
               Top             =   720
               Width           =   1935
            End
            Begin VB.Image Cmd_Find 
               Height          =   495
               Left            =   7080
               Picture         =   "Frm_GestionConge.frx":10C3A
               Stretch         =   -1  'True
               Top             =   720
               Width           =   1935
            End
            Begin VB.Label Lbl_Sup 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Supprimé"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   9000
               TabIndex        =   33
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Lbl_NSup 
               BackColor       =   &H8000000E&
               BackStyle       =   0  'Transparent
               Caption         =   "Non supprimé"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6960
               TabIndex        =   32
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Au :"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3720
               TabIndex        =   27
               Top             =   120
               Width           =   480
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Congé du :"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   240
               TabIndex        =   26
               Top             =   120
               Width           =   1230
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Conducteur"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   25
               Top             =   720
               Width           =   1335
            End
            Begin VB.Image Pic_Header 
               Height          =   1095
               Left            =   0
               Picture         =   "Frm_GestionConge.frx":2183C
               Stretch         =   -1  'True
               Top             =   0
               Width           =   11295
            End
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6375
         Left            =   120
         ScaleHeight     =   6345
         ScaleWidth      =   11265
         TabIndex        =   7
         Top             =   360
         Width           =   11295
         Begin VB.PictureBox Pict_cond 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2640
            ScaleHeight     =   615
            ScaleWidth      =   3735
            TabIndex        =   30
            Top             =   2160
            Width           =   3735
            Begin SToolBox.SBiCombo Cbo_Conducteur 
               Height          =   360
               Left            =   120
               TabIndex        =   1
               Top             =   120
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   635
               ForeColor       =   -2147483640
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin SToolBox.SCommand Cmd_FindConducteur 
               Height          =   345
               Left            =   3120
               TabIndex        =   31
               Top             =   120
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
               Picture         =   "Frm_GestionConge.frx":5D356
               ButtonType      =   1
            End
         End
         Begin VB.TextBox Txt_Observ 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   2760
            TabIndex        =   4
            Top             =   4560
            Width           =   3135
         End
         Begin VB.PictureBox Pict_Date 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   960
            ScaleHeight     =   495
            ScaleWidth      =   4815
            TabIndex        =   16
            Top             =   3000
            Width           =   4815
            Begin MSComCtl2.DTPicker cda_Debut 
               Height          =   375
               Left            =   1800
               TabIndex        =   2
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   14737632
               Format          =   144375809
               CurrentDate     =   42860
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Début Congé:"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   0
               TabIndex        =   17
               Top             =   120
               Width           =   1560
            End
         End
         Begin VB.PictureBox Pict_Num 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2640
            ScaleHeight     =   375
            ScaleWidth      =   1575
            TabIndex        =   15
            Top             =   1440
            Width           =   1575
            Begin VB.TextBox Txt_Num 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   0
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.PictureBox Pict_DateFin 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   960
            ScaleHeight     =   495
            ScaleWidth      =   4935
            TabIndex        =   13
            Top             =   3840
            Width           =   4935
            Begin MSComCtl2.DTPicker cda_Fin 
               Height          =   375
               Left            =   1800
               TabIndex        =   3
               Top             =   0
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   14737632
               Format          =   144375809
               CurrentDate     =   42860
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fin congé:"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   14
               Top             =   0
               Width           =   1305
            End
         End
         Begin VB.Image CmdSave 
            Height          =   495
            Left            =   9600
            Picture         =   "Frm_GestionConge.frx":5D690
            Stretch         =   -1  'True
            Top             =   600
            Width           =   1575
         End
         Begin VB.Image CmdDelete 
            Height          =   480
            Left            =   7920
            Picture         =   "Frm_GestionConge.frx":6F912
            Stretch         =   -1  'True
            Top             =   600
            Width           =   1575
         End
         Begin VB.Image CmdAdd 
            Height          =   495
            Left            =   6120
            Picture         =   "Frm_GestionConge.frx":8178C
            Stretch         =   -1  'True
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Lbl_info 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   6000
            TabIndex        =   23
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Lbl_Cond 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Conducteur:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   21
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Lbl_Obser 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Observation:"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   960
            TabIndex        =   20
            Top             =   4560
            Width           =   1695
         End
         Begin VB.Label Lbl_Num 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Numero :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   960
            TabIndex        =   19
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Lbl_Conge 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   6360
            TabIndex        =   18
            Top             =   3600
            Width           =   3615
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   0
            Picture         =   "Frm_GestionConge.frx":938B6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   11295
         End
      End
   End
   Begin VB.Label Lbl_Gestion 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion des congés"
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
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.Image PicBox_Header 
      Height          =   1000
      Left            =   0
      Picture         =   "Frm_GestionConge.frx":CF3D0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "Frm_GestionConge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim thekey As Integer
    Dim theshift As Integer

Private Sub cbo_conducteur_LostFocus()
Call ExistDonnee(cbo_conducteur)
End Sub

'Afficher la liste des conducteurs
Private Sub Cmd_FindConducteur_Click()
        On Error GoTo Err
    
    Unload FrmFind_Fils
    With FrmFind_Fils
        .StrSource = "PersoConge"
        .Show vbModal
    End With
    
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Suppression d'un congé qui ne doit pas être fini
Private Sub CmdDelete_Click()

Dim LOBJ_Personnel As personnel
Dim LOBJ_Conge As Conducteur

On Error GoTo Err

    If Txt_Num.Text <> "Auto" And Txt_Num.Text <> "" Then
        Set LOBJ_Personnel = New personnel
        If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Supp_Conge", LInt_UserId) Then
            MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-être pas des autorisations nécessaires pour Supprime un produit", vbExclamation, App.ProductName
            Exit Sub
        End If
    End If
    Set LOBJ_Personnel = Nothing

    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "Congé", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
    
    Set LOBJ_Conge = New Conducteur
    Call LOBJ_Conge.Delete_Conge(ErrNumber, ErrDescription, ErrSourceDetail, Txt_Num.Text, LInt_UserId, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Conge = Nothing
    MsgBox "Congé supprimer avec succées!...", vbInformation
    
    Call ViderZone(Frm_GestionConge)
    cbo_conducteur.Text = ""
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    cda_debut.Value = Date
    cda_fin.Value = Date

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'Ajouter un nouveau congé
Private Sub CmdAdd_Click()

Dim LOBJ_Personnel As personnel

On Error GoTo Err
'vérifier si l'utilisateur à le droit d'ajouter un nouveau congé
    Set LOBJ_Personnel = New personnel
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "INS_Conge", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
    Set LOBJ_Personnel = Nothing
    STab_conge.Tab = 0
    Call ViderZone(Frm_GestionConge)
    CmdDelete.Enabled = False
    Pict_cond.Enabled = True
    Pict_date.Enabled = True
    Pict_DateFin.Enabled = True
    CmdSave.Enabled = True
    cda_debut.Value = Date
    cda_fin.Value = Date
    Txt_Num.Text = "Auto"
    Lbl_Conge.Caption = ""
    Lbl_info.Caption = ""
    Txt_Observ.Enabled = True
    
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Chercher un conducteur par son code et afficher le code et le nom dans le comboBox s'il existe
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
    If Not IsNull(rs("Libelle")) Then cbo_conducteur.Text = rs("code") & "  -  " & rs("Libelle")
Else
    MsgBox "Code invalide, vérifier votre saisie.", vbExclamation, App.ProductName
    Exit Sub
End If
End Sub

'Impression des congés pour une période donnée
Private Sub CmdPrint_Click()

Dim F As Form
Dim vconduc As String
Dim Text As String
Dim sup As String

On Error GoTo Err
'tester si les dates entrées sont valides : date_debut < date_fin
If cda_Deb.Value > cda_au.Value Then
    MsgBox "Vérifier dates ", vbExclamation, App.ProductName
    Exit Sub
End If

vconduc = Scbo_Conduc.FirstValue

If Opt_Sup.Value = vbChecked Then
    sup = "O"
ElseIf OptNSup.Value = vbChecked Then
    sup = "N"
End If
'Charger grid avant l'impression
grid.ClearRows
Call Affiche_Conge(cda_Deb.Value, cda_au.Value, Scbo_Conduc.FirstValue, sup)
'Si grid vide : pas de données à imprimer --> exit
If grid.Rows = 0 Then Exit Sub

If MsgBox("Imprimer ce planning de congé   ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Call Frm_Rpt_Apercus.PrintOutAndApercu_Conge(0, cda_Deb.Value, cda_au.Value, vconduc, LStr_NameUser)
    Frm_Rpt_Apercus.Show
End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Enregistrement d'un nouveau congé ou du MAJ
Private Sub CmdSave_Click()

Dim LOBJ_Personnel As New personnel
Dim LOBJ_Conge As New Conducteur
Dim Lobj_PLNG As New PLANNING
Dim rs As New Recordset
Dim Text As String

STab_conge.Tab = 0
If Left(CheckMandatory(Frm_GestionConge), 1) = 1 Then
   Exit Sub
End If

If cbo_conducteur.Text = "" Or cbo_conducteur.ListIndex = 0 Or Len(Trim(cbo_conducteur.Text)) = 0 Then
        MsgBox "Conducteur obligatoire      ", vbInformation
        cbo_conducteur.SetFocus
        Exit Sub
End If

If Txt_Num.Text = "Auto" Then
    If cda_debut.Value > cda_fin.Value Or cda_debut.Value < Date Or cda_fin.Value < Date Then
        MsgBox "Vérifier les dates ! ", vbInformation
        Exit Sub
    End If
ElseIf Txt_Num.Text <> "Auto" And Txt_Num.Text <> "" Then
     If cda_debut.Value > cda_fin.Value Then
        MsgBox "Vérifier les dates ! ", vbInformation
        Exit Sub
    End If
    If cda_fin.Value < Date Then
        MsgBox "Congé fini vous ne pouvez pas ni le changer ni le supprimer.", vbInformation
        Exit Sub
    End If
End If

If Txt_Num.Text = "Auto" Then
    'Vérifier si ce conducteur à un congé dans cette période
    Set rs = LOBJ_Conge.Get_AllCongeConduc(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cda_debut.Value, cda_fin.Value, cbo_conducteur.FirstValue, "N")
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Conge = Nothing
    If Not rs.EOF Then
        MsgBox "Conducteur ayant un congé durant cette période ", vbInformation, App.ProductName
        grid.ClearRows
        Call Affiche_Conge(cda_debut.Value, cda_fin.Value, cbo_conducteur.FirstValue, "N")
        STab_conge.Tab = 1
        Scbo_Conduc.Text = cbo_conducteur.Text
        Exit Sub
    End If
    rs.Close
End If

'Vérifier si le conducteur à une tournée
Dim i As Date
Dim jour As String
Dim semainePLNG As Date
Dim text_tournee As String

For i = cda_debut.Value To cda_fin.Value
    jour = Format(i, "dddd")
    Set rs = Lobj_PLNG.Get_CondPLNG(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Frm_PLANNING.DatePlanning(i), jour, cbo_conducteur.FirstValue)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
        If text_tournee = "" Then
            text_tournee = rs("CONDUCTEUR") & " a une tournée " & vbCr & " - le " & i & " à " & rs("TOURNEE")
        Else
            text_tournee = text_tournee & vbCr & " - le " & i & " à " & rs("TOURNEE")
        End If
    End If
    rs.Close
    Set Lobj_PLNG = Nothing
Next
'Afficher tout les tournée auxquelles le conducteur est affecté durant la période du congé
If text_tournee <> "" Then
    text_tournee = text_tournee & vbCr & "Vérifier le Planning de cette période avant de confirmer le congé."
    MsgBox text_tournee, vbInformation, App.ProductName
    Exit Sub
End If

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

If Txt_Num.Text = "Auto" Then
    Call Insert_Conge
End If

If Txt_Num.Text <> "" And Txt_Num.Text <> "Auto" Then
    Set LOBJ_Personnel = New personnel
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Maj_Conge", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If

    Call Modif_Conge
End If

MsgBox "Enregistrement terminé avec succé.  ", vbQuestion, App.ProductName
Pict_cond.Enabled = False
Set rs = LOBJ_Conge.Get_MaxNumConge(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    Txt_Num.Text = rs("maxNum")
End If
rs.Close
Set LOBJ_Conge = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub
'Insertion d'un nouveau congé
Private Sub Insert_Conge()

Dim LRs_NewRecord As New Recordset
Dim LOBJ_Conge As Conducteur

Set LRs_NewRecord = CreateEmptyRS_Conge
With LRs_NewRecord
    .AddNew
    .Fields("Conducteur") = cbo_conducteur.FirstValue
    .Fields("DateDu") = cda_debut.Value
    .Fields("DateAu") = cda_fin.Value
    .Fields("Type") = "Congé"
    .Fields("Observation") = Txt_Observ.Text
    .Fields("UserInsert") = LInt_UserId
End With
Set LOBJ_Conge = New Conducteur
Call LOBJ_Conge.Insert_Conge(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
Set LOBJ_Conge = Nothing

End Sub
'Modification d'un congé : changé seulement la durée ou l'observation
Private Sub Modif_Conge()

Dim LRs_NewRecord As New Recordset
Dim LOBJ_Conge As Conducteur

Set LRs_NewRecord = CreateEmptyRS_Conge
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = Txt_Num.Text
    .Fields("Conducteur") = cbo_conducteur.FirstValue
    .Fields("DateDu") = cda_debut.Value
    .Fields("DateAu") = cda_fin.Value
    .Fields("Observation") = Txt_Observ.Text
    .Fields("UserUpdate") = LInt_UserId
End With
Set LOBJ_Conge = New Conducteur
Call LOBJ_Conge.Update_Conge(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
Set LOBJ_Conge = Nothing

End Sub
'Afficher détail d'un congé par code
Public Sub AfficheRow(ByVal VCode As Integer)

Dim LOBJ_Conge As Conducteur
Dim rs As New Recordset
Dim DateSorti As Date

Call ViderZone(Frm_GestionConge)
CmdSave.Enabled = True
CmdDelete.Enabled = True

Set LOBJ_Conge = New Conducteur
Set rs = LOBJ_Conge.Get_CongeCondByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

If Not rs.EOF Then
    Txt_Num.Text = rs("Numero")
    Call AfficheRow_Conducteur(rs("Conducteur"))

    If (Not (IsNull(rs("DateDu")))) Then cda_debut.Value = Format(rs("DateDu"), "dd/mm/yyyy")
    If (Not (IsNull(rs("DateAu")))) Then cda_fin.Value = Format(rs("DateAu"), "dd/mm/yyyy")
    
    If rs("Supp") = "O" Then            'si congé supprimé on ne peut ni le supprimé de nouveau ni le modifier
        Lbl_Conge.Caption = " Congé annulé par " & Get_NameUserByCode(rs("UserDelete")) & "  ! "
        Lbl_info.Caption = ""
        CmdSave.Enabled = False
        CmdDelete.Enabled = False
        Txt_Observ.Enabled = False
    Else
        Lbl_Conge.Caption = " "
        If rs("DateAu") < Date Then     'si congé fini on ne peut ni le modifier ni le supprimer
            Lbl_info.Caption = "Congé fini ! "
            Pict_date.Enabled = False
            Pict_DateFin.Enabled = False
            CmdSave.Enabled = False
            CmdDelete.Enabled = False
            Txt_Observ.Enabled = False
        Else                            'Modification ou suppression possible
            Lbl_info.Caption = ""
            Pict_date.Enabled = True
            Pict_DateFin.Enabled = True
            Txt_Observ.Enabled = True
        End If
    End If
    Txt_Observ.Text = rs("Observation")
    Pict_cond.Enabled = False
Else
    MsgBox "Code introuvable", vbInformation
End If
rs.Close
End Sub

Private Sub Form_Load()
    
    STab_conge.Tab = 0
    cda_debut.Value = Date
    cda_fin.Value = Date
    cbo_conducteur.AddItem "0000", "Conducteur"
    Call Affiche_Personnel_SBCombo(cbo_conducteur)
    cda_Deb.Value = "01/01/" & Year(Date)
    cda_au.Value = Date
    Scbo_Conduc.AddItem "0000", "Tous"
    Call Affiche_Personnel_SBCombo(Scbo_Conduc)
    Scbo_Conduc.ListIndex = 0
    OptNSup.Value = vbChecked
    Call Initgrid_Conge
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    Pict_cond.Enabled = False
End Sub

Private Sub Form_Resize()
    Dim WidthForm As Integer
        WidthForm = Frm_Main.ACB_Main.Width
        PicBox_Header.Width = WidthForm - 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo erreur
'    Dim i As Integer
'    Dim Msg
'    Msg = "Voulez-vous vraiment quitter?"
'    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'        Cancel = True
'    Else
'        Unload Me
'    End If
'
'Exit Sub
'erreur:
'   MsgBox Err.Description, 48
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
    End If
End Sub
Private Sub Cbo_Conducteur_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub

'===================================================================
'========================= Recherche des congés ====================

Private Sub Cmd_Find_Click()

Dim sup As String
Dim cond As String
Dim Text As String

If cda_Deb.Value > cda_au.Value Then
    MsgBox "Vérifier dates ", vbExclamation, App.ProductName
    Exit Sub
End If

cond = Scbo_Conduc.FirstValue

If Opt_Sup.Value = vbChecked Then
    sup = "O"                           'Congé supprimé
Else
    sup = "N"                           'Congé non supprimé
End If

grid.ClearRows
Call Affiche_Conge(cda_Deb.Value, cda_au.Value, Scbo_Conduc.FirstValue, sup)

End Sub
'Initialisation du grid
Public Sub Initgrid_Conge()
With grid
    .Redraw = False

    .HideGroupingBox = True
    .AllowGrouping = True

    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "Numero", "Numero", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Conducteur", "Conducteur", , , 140
    .AddColumn "DateDu", "DateDu", , , 110
    .AddColumn "DateAu", "DateAu", , , 110
    .AddColumn "Type", "Type", , , 70, False
    .AddColumn "Supp", "Supp", , , 40
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

'Chercher les congés durant une période donnée soit pour tout les conducteurs ou un seul conducteur (Congé supprimé ou non supprimé)
Public Sub Affiche_Conge(ByVal date_Deb As Date, ByVal date_fin As Date, ByVal Conduc As String, ByVal sup As String)

Dim LOBJ_Conge As Conducteur
Dim rs As New Recordset

Set LOBJ_Conge = New Conducteur
Set rs = LOBJ_Conge.Get_AllCongeConduc(ErrNumber, ErrDescription, ErrSourceDetail, CNB, date_Deb, date_fin, Conduc, sup)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    grid.Redraw = False
    While Not rs.EOF
        With grid
            .AddRow
            .CellDetails .Rows, 1, rs("Numero")
            .CellDetails .Rows, .ColumnIndex("Conducteur"), rs("Libelle")
            .CellDetails .Rows, .ColumnIndex("DateDu"), rs("DateDu")
            .CellDetails .Rows, .ColumnIndex("DateAu"), rs("DateAu")
            .CellDetails .Rows, .ColumnIndex("Type"), rs("Type")
            .CellDetails .Rows, .ColumnIndex("Supp"), rs("Supp")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
Else
    MsgBox "Pas de données à visualiser.", vbInformation, App.ProductName
    Exit Sub
End If
End Sub
'En appyant sur "Entrée" , afficher les détails du congé dans la table "Gestion des congés"
Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

Dim VCode
On Error GoTo Err

VCode = grid.CellText(grid.SelectedRow, 1)

 If KeyCode = vbKeyReturn Then
    STab_conge.Tab = 0
    Call AfficheRow(VCode)
 End If
  
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Par double click sur grid afficher les détails du congé dans la table "Gestion des congés"
Private Sub grid_DblClick(ByVal lRow As Long, ByVal lCol As Long)

Dim VCode

On Error GoTo Err

VCode = grid.CellText(lRow, 1)
STab_conge.Tab = 0
Call AfficheRow(VCode)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub Cbo_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'si click = Entrée , faire appel à la fonction Cmd_find_click
Private Sub Scbo_Conduc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call Cmd_Find_Click
End Sub

Private Sub cda_Debut_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
 If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cda_Fin_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Scbo_Conduc_LostFocus()
Call ExistDonnee(Scbo_Conduc)
End Sub

