VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHub 
   Appearance      =   0  'Flat
   Caption         =   "PT DC Hub x.x.x"
   ClientHeight    =   5640
   ClientLeft      =   2550
   ClientTop       =   2760
   ClientWidth     =   9270
   Icon            =   "frmHub.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9270
   Begin ComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   135
      Top             =   5385
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   8
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1588
            Text            =   "00:00:00"
            TextSave        =   "00:00:00"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "[M:Months](W:Weeks)(D:Days) Hours:Minutes:Seconds"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "3:48"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Time"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "27-04-2008"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "0"
            TextSave        =   "0"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Users Online"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "0 Bytes"
            TextSave        =   "0 Bytes"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Shared total"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "0"
            TextSave        =   "0"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Op Online"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Hub IP"
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3678
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "DSN Status"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBordTab 
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   9170
      ScaleHeight     =   260
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   137
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox picHideObj 
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   4035
      TabIndex        =   367
      Top             =   4680
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Timer tmrScriptTimer 
         Enabled         =   0   'False
         Index           =   0
         Left            =   2400
         Top             =   0
      End
      Begin VB.Timer tmrBackground 
         Enabled         =   0   'False
         Left            =   1440
         Top             =   0
      End
      Begin VB.Timer tmrSysInfo 
         Interval        =   1000
         Left            =   1920
         Top             =   0
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl 
         Index           =   0
         Left            =   2880
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin MSWinsockLib.Winsock wskRegister 
         Index           =   0
         Left            =   480
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wskLoop 
         Index           =   0
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wskListen 
         Index           =   0
         Left            =   960
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin ComctlLib.ImageList imlScripts 
         Left            =   3480
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   17
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":15162
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":154B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":15806
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":15B58
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":15EAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":161FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":1654E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":168A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":16BF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":16F44
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":17296
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":175E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":1793A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":17C8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":17FDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":18330
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":18682
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   0
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   208
      Top             =   420
      Width           =   9015
      Begin VB.PictureBox picLog 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H00C0C0C0&
         Height          =   2565
         Index           =   0
         Left            =   4560
         ScaleHeight     =   2565
         ScaleWidth      =   4095
         TabIndex        =   362
         Top             =   480
         Width           =   4095
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produced bY fLaSh"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   2
            Left            =   1080
            TabIndex        =   364
            Top             =   600
            Width           =   1785
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PT Direct Connect Hub"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   0
            Left            =   840
            TabIndex        =   363
            Top             =   360
            Width           =   2205
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PT Direct Connect Hub"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   855
            TabIndex        =   365
            Top             =   375
            Width           =   2205
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produced bY fLaSh"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   1095
            TabIndex        =   366
            Top             =   615
            Width           =   1785
         End
      End
      Begin VB.TextBox txtUpTime 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   209
         Text            =   "00:00:00"
         Top             =   3525
         Width           =   1815
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   36
         Left            =   1920
         TabIndex        =   7
         Tag             =   "RedirectAddress"
         ToolTipText     =   "Seperate addresses with a semicolon (;)"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   4
         Left            =   1920
         TabIndex        =   6
         Tag             =   "RegisterIP"
         ToolTipText     =   "Seperate addresses with a semicolon (;)"
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Tag             =   "Ports"
         ToolTipText     =   "Seperate ports with a semicolon (;) (First port in the list is the one used for registration purposes)"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Tag             =   "HubIP"
         ToolTipText     =   "The address for your hub"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   1
         Left            =   1920
         MaxLength       =   140
         TabIndex        =   3
         Tag             =   "HubDesc"
         ToolTipText     =   "A short description of your hub (140 characters max)"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   0
         Left            =   1920
         MaxLength       =   70
         TabIndex        =   2
         Tag             =   "HubName"
         ToolTipText     =   "Name of your hub (70 characters max)"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Start Server"
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   0
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   5
         X1              =   4440
         X2              =   4800
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   4
         X1              =   4440
         X2              =   8760
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   3
         X1              =   4440
         X2              =   4440
         Y1              =   3120
         Y2              =   360
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the world of"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   327
         Top             =   240
         Width           =   3735
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   11
         X1              =   6720
         X2              =   7080
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   10
         X1              =   6720
         X2              =   8760
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   9
         X1              =   6720
         X2              =   6720
         Y1              =   3960
         Y2              =   3360
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Hub Control"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   3
         Left            =   7200
         TabIndex        =   218
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   8
         X1              =   4440
         X2              =   4800
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   7
         X1              =   4440
         X2              =   6480
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   6
         X1              =   4440
         X2              =   4440
         Y1              =   3960
         Y2              =   3360
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "UpTime"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   217
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   2
         X1              =   120
         X2              =   480
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   120
         X2              =   120
         Y1              =   3960
         Y2              =   240
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   120
         X2              =   4200
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Redirect Address"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   99
         Left            =   240
         TabIndex        =   216
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Register Address"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   215
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Listening Ports"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   214
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   213
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   212
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   211
         Top             =   600
         Width           =   1575
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00C0C0C0&
         Height          =   1095
         Index           =   1
         Left            =   240
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00C0C0C0&
         Height          =   1095
         Index           =   2
         Left            =   240
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00C0C0C0&
         Height          =   1095
         Index           =   0
         Left            =   240
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Hub Settings"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   210
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   8
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   203
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   1
         Left            =   4680
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   4200
         TabIndex        =   207
         Top             =   60
         Width           =   4200
      End
      Begin VB.PictureBox picHelp 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   204
         Top             =   420
         Width           =   8715
         Begin VB.TextBox txtAbout 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   2640
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   543
            Top             =   120
            Width           =   5895
         End
         Begin VB.PictureBox picLog 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillColor       =   &H00C0C0C0&
            ForeColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   2
            Left            =   240
            MouseIcon       =   "frmHub.frx":189D4
            MousePointer    =   99  'Custom
            ScaleHeight     =   465
            ScaleWidth      =   930
            TabIndex        =   356
            Top             =   2850
            Width           =   930
         End
         Begin VB.PictureBox picLog 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillColor       =   &H00C0C0C0&
            ForeColor       =   &H00C0C0C0&
            Height          =   2565
            Index           =   1
            Left            =   120
            ScaleHeight     =   2565
            ScaleWidth      =   2475
            TabIndex        =   324
            Top             =   120
            Width           =   2475
         End
         Begin VB.Label LabelsURL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   0
            Left            =   1200
            MouseIcon       =   "frmHub.frx":18B26
            MousePointer    =   99  'Custom
            TabIndex        =   326
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label LabelsURL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Home Page"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   1
            Left            =   1200
            MouseIcon       =   "frmHub.frx":18C78
            MousePointer    =   99  'Custom
            TabIndex        =   325
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   645
            Index           =   4
            Left            =   120
            Top             =   2760
            Width           =   2465
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   285
            Index           =   23
            Left            =   2640
            Top             =   3120
            Width           =   5895
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direct Connect P2P Network"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Index           =   4
            Left            =   3960
            TabIndex        =   205
            Top             =   3120
            Width           =   3120
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direct Connect P2P Network"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   5
            Left            =   3975
            TabIndex        =   206
            Top             =   3135
            Width           =   3120
         End
      End
      Begin VB.PictureBox picHelp 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   314
         Top             =   420
         Visible         =   0   'False
         Width           =   8715
         Begin VB.TextBox txtNotePad 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   3285
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   136
            Top             =   120
            Width           =   8535
         End
      End
      Begin VB.PictureBox picHelp 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   540
         Top             =   480
         Visible         =   0   'False
         Width           =   8715
         Begin VB.ComboBox cmbChangeLog 
            Height          =   315
            ItemData        =   "frmHub.frx":18DCA
            Left            =   60
            List            =   "frmHub.frx":18DCC
            Style           =   2  'Dropdown List
            TabIndex        =   542
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox txtChangeLog 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   2925
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   541
            Top             =   480
            Width           =   8535
         End
      End
      Begin ComctlLib.TabStrip tbsHelp 
         Height          =   3945
         Left            =   60
         TabIndex        =   134
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "About PTDCH"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Version History"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "NotePad"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   6
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   202
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   6
         Left            =   8880
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   120
         TabIndex        =   322
         Top             =   60
         Width           =   120
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   328
         Top             =   430
         Width           =   8655
         Begin VB.TextBox txtStForm 
            Height          =   285
            Left            =   60
            MaxLength       =   40
            TabIndex        =   434
            Top             =   3075
            Width           =   1100
         End
         Begin VB.CommandButton cmdStSend 
            Caption         =   "Send"
            Height          =   495
            Left            =   5280
            TabIndex        =   433
            Top             =   2880
            Width           =   735
         End
         Begin VB.OptionButton optStSend 
            Caption         =   "Data"
            Height          =   195
            Index           =   1
            Left            =   6060
            TabIndex        =   432
            Top             =   3150
            Width           =   195
         End
         Begin VB.OptionButton optStSend 
            Caption         =   "Chat"
            Height          =   195
            Index           =   0
            Left            =   6060
            TabIndex        =   431
            Top             =   2925
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.TextBox txtStSend 
            Height          =   285
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   430
            Top             =   3075
            Width           =   3975
         End
         Begin VB.PictureBox picStInfo 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   5810
            ScaleHeight     =   1935
            ScaleWidth      =   2775
            TabIndex        =   332
            Top             =   810
            Visible         =   0   'False
            Width           =   2775
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "6 = Send PM To UnRegistered"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   423
               Top             =   1560
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "4 = Send PM To All"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   422
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "5 = Send PM To Op"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   337
               Top             =   1320
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "3 = Send Chat To UnRegistered"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   336
               Top             =   840
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "2 = Send Chat To Op"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   335
               Top             =   600
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "1 = Send Chat To All"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   334
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "------------- Send Chat -------------"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   333
               Top             =   120
               Width           =   2535
            End
         End
         Begin VB.ListBox lstStatus 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2700
            Index           =   0
            IntegralHeight  =   0   'False
            Left            =   60
            TabIndex        =   329
            Top             =   60
            Width           =   8535
         End
         Begin ComctlLib.Slider sldStatus 
            Height          =   315
            Left            =   6990
            TabIndex        =   435
            Tag             =   "PriorityVal"
            Top             =   2880
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   1
            Max             =   6
            SelectRange     =   -1  'True
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label lblOptStSend 
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            Height          =   255
            Index           =   1
            Left            =   6300
            TabIndex        =   440
            Top             =   3150
            Width           =   615
         End
         Begin VB.Label lblOptStSend 
            BackStyle       =   0  'Transparent
            Caption         =   "Chat"
            Height          =   255
            Index           =   0
            Left            =   6300
            TabIndex        =   439
            Top             =   2925
            Width           =   615
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Message or Data"
            Height          =   255
            Index           =   26
            Left            =   1200
            TabIndex        =   438
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Labels 
            BackStyle       =   0  'Transparent
            Caption         =   "Form"
            Height          =   255
            Index           =   9
            Left            =   60
            TabIndex        =   437
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "1    2    3    4    5    6"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   7040
            TabIndex        =   436
            Top             =   3240
            Width           =   1575
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   4
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   339
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin ComctlLib.ListView lvwUsers 
            Height          =   3255
            Left            =   3840
            TabIndex        =   340
            Top             =   120
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   5741
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Winsock Index"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Connected Since"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   300
            Index           =   24
            Left            =   120
            Top             =   120
            Width           =   3615
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Statistics"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   353
            Top             =   180
            Width           =   3135
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1095
            Index           =   22
            Left            =   120
            Top             =   1920
            Width           =   3615
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1095
            Index           =   21
            Left            =   120
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Connected users :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   45
            Left            =   240
            TabIndex        =   352
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Connected operators :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   47
            Left            =   240
            TabIndex        =   351
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Total shared :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   51
            Left            =   240
            TabIndex        =   350
            Top             =   1290
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Peak users :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   52
            Left            =   240
            TabIndex        =   349
            Top             =   2130
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Peak operators :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   53
            Left            =   240
            TabIndex        =   348
            Top             =   2370
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Peak shared :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   54
            Left            =   240
            TabIndex        =   347
            Top             =   2610
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   55
            Left            =   2160
            TabIndex        =   346
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   56
            Left            =   2160
            TabIndex        =   345
            Top             =   1050
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0 Bytes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   57
            Left            =   2160
            TabIndex        =   344
            Top             =   1290
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   58
            Left            =   2160
            TabIndex        =   343
            Top             =   2130
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   59
            Left            =   2160
            TabIndex        =   342
            Top             =   2370
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0 Bytes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   60
            Left            =   2160
            TabIndex        =   341
            Top             =   2610
            Width           =   1575
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   5
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   441
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   29
            X1              =   4320
            X2              =   7920
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   28
            X1              =   4320
            X2              =   7920
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Speed Recived:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   4320
            TabIndex        =   486
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Speed Send:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   4320
            TabIndex        =   485
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Recived:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   4320
            TabIndex        =   484
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Send:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   4320
            TabIndex        =   483
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0 Kbs"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   6600
            TabIndex        =   482
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0 Kbs"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   6600
            TabIndex        =   481
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0 bytes"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   6600
            TabIndex        =   480
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0 bytes"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   6600
            TabIndex        =   479
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "MyINFOs:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   4320
            TabIndex        =   478
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A/P Searchs:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   4320
            TabIndex        =   477
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Private messages:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   4320
            TabIndex        =   476
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Main Chat messages:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   4320
            TabIndex        =   475
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0 Bytes"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   6600
            TabIndex        =   474
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0 Bytes"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   6600
            TabIndex        =   473
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0 Bytes"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   6600
            TabIndex        =   472
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0 Bytes"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   6600
            TabIndex        =   471
            Top             =   480
            Width           =   1215
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   27
            X1              =   480
            X2              =   3240
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Requests:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   470
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Aborted Sockets:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   480
            TabIndex        =   469
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Failed Sockets:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   480
            TabIndex        =   468
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NetINFO:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   467
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "BotINFO:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   480
            TabIndex        =   466
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Redirects:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   480
            TabIndex        =   465
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kicks:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   464
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RevConnectMe:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   480
            TabIndex        =   463
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ConnectMe:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   480
            TabIndex        =   462
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NickList:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   480
            TabIndex        =   461
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   2760
            TabIndex        =   460
            Top             =   2880
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   2760
            TabIndex        =   459
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   2760
            TabIndex        =   458
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   2760
            TabIndex        =   457
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   456
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   455
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   454
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   453
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   452
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblStatistics 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   451
            Top             =   480
            Width           =   615
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   41
            X1              =   3960
            X2              =   8280
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   62
            X1              =   120
            X2              =   3600
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Trafic"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   16
            Left            =   4440
            TabIndex        =   443
            Top             =   120
            Width           =   3015
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   53
            X1              =   3960
            X2              =   3960
            Y1              =   3360
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   52
            X1              =   3960
            X2              =   4320
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   51
            X1              =   120
            X2              =   480
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   50
            X1              =   120
            X2              =   120
            Y1              =   3360
            Y2              =   240
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Protocol"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   15
            Left            =   600
            TabIndex        =   442
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   331
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.ListBox lstStatus 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3180
            Index           =   1
            IntegralHeight  =   0   'False
            ItemData        =   "frmHub.frx":18DCE
            Left            =   60
            List            =   "frmHub.frx":18DD0
            TabIndex        =   338
            Top             =   60
            Width           =   8535
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   3
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   330
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox txtLog 
            BackColor       =   &H8000000F&
            Height          =   3255
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   544
            Top             =   120
            Width           =   8415
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   354
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.ListBox lstStatus 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3180
            Index           =   2
            IntegralHeight  =   0   'False
            ItemData        =   "frmHub.frx":18DD2
            Left            =   60
            List            =   "frmHub.frx":18DD4
            TabIndex        =   355
            Top             =   60
            Width           =   8535
         End
      End
      Begin ComctlLib.TabStrip tbsStatus 
         Height          =   3945
         Left            =   60
         TabIndex        =   323
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2558
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   6
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Main Chat Log"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "PM Log"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Protocol Misc Log"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "System Log"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Status"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Satistics"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   7
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   201
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   5
         Left            =   3000
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   6045
         TabIndex        =   319
         Top             =   60
         Width           =   6045
      End
      Begin VB.PictureBox picInfo 
         BorderStyle     =   0  'None
         Height          =   3400
         Index           =   0
         Left            =   120
         ScaleHeight     =   3405
         ScaleWidth      =   8655
         TabIndex        =   317
         Top             =   480
         Width           =   8655
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "ReInstall All"
            Height          =   300
            Index           =   4
            Left            =   4920
            TabIndex        =   551
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "Refresh"
            Height          =   300
            Index           =   3
            Left            =   3720
            TabIndex        =   550
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "Reolad All"
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   2520
            TabIndex        =   496
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "Reolad"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1320
            TabIndex        =   495
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Enabled Plugins"
            Height          =   195
            Index           =   68
            Left            =   8280
            TabIndex        =   320
            Tag             =   "Plugins"
            ToolTipText     =   "This option requests restart the application.."
            Top             =   3050
            Width           =   195
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "Setup"
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   318
            Top             =   3000
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwPlugins 
            Height          =   2775
            Left            =   120
            TabIndex        =   549
            Top             =   120
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Use Events"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Version"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Author"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Description"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Release"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Comments"
               Object.Width           =   7056
            EndProperty
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enabled Plugins"
            Height          =   255
            Index           =   68
            Left            =   6600
            TabIndex        =   321
            ToolTipText     =   "This option requests restart the application.."
            Top             =   3030
            Width           =   1575
         End
      End
      Begin VB.PictureBox picInfo 
         BorderStyle     =   0  'None
         Height          =   3400
         Index           =   1
         Left            =   120
         ScaleHeight     =   3405
         ScaleWidth      =   8655
         TabIndex        =   316
         Top             =   480
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdLGApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   554
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtLanguages 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   553
            Text            =   "English.xml"
            Top             =   120
            Width           =   1695
         End
         Begin ComctlLib.ListView lvwLanguages 
            Height          =   1455
            Left            =   240
            TabIndex        =   552
            Top             =   480
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "International Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "National Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Translated By"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Author e-mail"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Operatores"
            Enabled         =   0   'False
            Height          =   735
            Index           =   8
            Left            =   6840
            TabIndex        =   450
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "All"
            Enabled         =   0   'False
            Height          =   375
            Index           =   7
            Left            =   4920
            TabIndex        =   449
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "UnRegistereds"
            Enabled         =   0   'False
            Height          =   375
            Index           =   9
            Left            =   4920
            TabIndex        =   448
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdConvDatabase 
            Caption         =   "Convert database"
            Height          =   375
            Left            =   6360
            TabIndex        =   445
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Check updates"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   420
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Detect IP"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   419
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Reload Settings"
            Height          =   375
            Index           =   4
            Left            =   2520
            TabIndex        =   417
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Save Settings"
            Height          =   375
            Index           =   3
            Left            =   2520
            TabIndex        =   416
            Top             =   2400
            Width           =   1935
         End
         Begin VB.ComboBox cmbInterface 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6480
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   414
            Tag             =   "Interface"
            ToolTipText     =   "Set Interface Language for DDCH"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   24
            X1              =   4800
            X2              =   5160
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   25
            X1              =   4800
            X2              =   4800
            Y1              =   3240
            Y2              =   2280
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   26
            X1              =   4800
            X2              =   8400
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Mas Messages"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   7
            Left            =   5280
            TabIndex        =   447
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Labels 
            BackStyle       =   0  'Transparent
            Caption         =   "Database (*.XML) of the YnHub or PtokaX."
            ForeColor       =   &H00808080&
            Height          =   735
            Index           =   6
            Left            =   6360
            TabIndex        =   446
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Convert Accounts"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   8
            Left            =   6720
            TabIndex        =   444
            Top             =   360
            Width           =   2175
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   18
            X1              =   120
            X2              =   480
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   19
            X1              =   120
            X2              =   120
            Y1              =   3240
            Y2              =   2280
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   20
            X1              =   120
            X2              =   2280
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "UpDate"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   421
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   15
            X1              =   2400
            X2              =   2760
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   16
            X1              =   2400
            X2              =   2400
            Y1              =   3240
            Y2              =   2280
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   17
            X1              =   2400
            X2              =   4680
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Settings"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   418
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   12
            X1              =   120
            X2              =   600
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   13
            X1              =   120
            X2              =   120
            Y1              =   2040
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   14
            X1              =   120
            X2              =   6120
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Interface"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   415
            Top             =   120
            Width           =   2295
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   47
            X1              =   6240
            X2              =   8400
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   46
            X1              =   6240
            X2              =   6600
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   45
            X1              =   6240
            X2              =   6240
            Y1              =   2040
            Y2              =   480
         End
      End
      Begin ComctlLib.TabStrip tbsInfo 
         Height          =   3975
         Left            =   60
         TabIndex        =   315
         Top             =   60
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         TabWidthStyle   =   2
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Plugins"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Others"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   5
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin ComctlLib.Toolbar tlbScript 
         Height          =   390
         Left            =   60
         TabIndex        =   380
         Top             =   60
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         ImageList       =   "imlScripts"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   22
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo"
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find"
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Replace"
               Object.ToolTipText     =   "Replace"
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "GoToLine"
               Object.ToolTipText     =   "Go To Line"
               Object.Tag             =   ""
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Save Only"
               Object.ToolTipText     =   "Save Only"
               Object.Tag             =   ""
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Save and Reset Script"
               Object.ToolTipText     =   "Save and Reset Script"
               Object.Tag             =   "9"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Clear"
               Object.ToolTipText     =   "Clear"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Hide Scripts"
               Object.ToolTipText     =   "Hide Scripts"
               Object.Tag             =   ""
               ImageIndex      =   9
               Style           =   1
            EndProperty
            BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Hide TabControl"
               Object.ToolTipText     =   "Hide TabControl"
               Object.Tag             =   ""
               ImageIndex      =   10
               Style           =   1
            EndProperty
            BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Show Debug Windows"
               Object.ToolTipText     =   "Show Debug Windows"
               Object.Tag             =   ""
               ImageIndex      =   14
            EndProperty
            BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               ImageIndex      =   17
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Enabled Tabs"
               Object.ToolTipText     =   "Enabled Tabs"
               Object.Tag             =   ""
               ImageIndex      =   11
               Style           =   1
            EndProperty
            BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "New"
               Object.ToolTipText     =   "New"
               Object.Tag             =   ""
               ImageIndex      =   16
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Menu"
               Object.ToolTipText     =   "Menu"
               Object.Tag             =   ""
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtScriptError 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   531
         Top             =   3720
         Width           =   8775
      End
      Begin MSComctlLib.ListView lvwScripts 
         Height          =   3495
         Left            =   7440
         TabIndex        =   425
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "State"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Script Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.PictureBox picSciMain 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Index           =   0
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   7155
         TabIndex        =   379
         Top             =   840
         Visible         =   0   'False
         Width           =   7155
      End
      Begin ComctlLib.TabStrip tbsScripts 
         Height          =   3135
         Left            =   60
         TabIndex        =   424
         Top             =   480
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   5530
         TabWidthStyle   =   2
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "New Script.vbs"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   4
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   178
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton cmdButton 
         Caption         =   "Redirect users"
         Enabled         =   0   'False
         Height          =   465
         Index           =   2
         Left            =   5520
         TabIndex        =   133
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "Allow ops to redirect (admins unaffected)"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   8
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   3720
         TabIndex        =   132
         Tag             =   "OpsCanRedirect"
         ToolTipText     =   "Check to allow operators to use the redirect ability (admins can always redirect)"
         Top             =   3000
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "All users"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   126
         ToolTipText     =   "Redirects all users to your redirect address"
         Top             =   1200
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Do not redirect"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   131
         ToolTipText     =   "Do not redirect anyone (disconnects if the hub is full)"
         Top             =   2400
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Only if full and the user is not registered"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   130
         ToolTipText     =   "Only redirect if the user is not registered and the hub is full"
         Top             =   2160
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Only if full and the user is not an op"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   129
         ToolTipText     =   "Only redirect if the user is an operator (or of a higher class) and the hub is full"
         Top             =   1920
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Only if full"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   128
         ToolTipText     =   "Only redirect users if the hub is full"
         Top             =   1680
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Unregistered users"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   127
         ToolTipText     =   "Only unregistered users are redirected (must have ""Allow only registered users"" in Security/Advanced checked)"
         Top             =   1440
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   32
         Left            =   1560
         TabIndex        =   107
         Tag             =   "ForTooOldDcppRedirectAddress"
         Top             =   2655
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "Too Old DC++"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   44
         Left            =   120
         TabIndex        =   120
         Tag             =   "RedirectFTooOldDCpp"
         Top             =   2640
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   28
         Left            =   1560
         TabIndex        =   108
         Tag             =   "ForTooOldNMDCRedirectAddress"
         Top             =   3015
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "Too Old NMDC"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   40
         Left            =   120
         TabIndex        =   121
         Tag             =   "RedirectFTooOldNMDC"
         Top             =   3000
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   31
         Left            =   1560
         TabIndex        =   106
         Tag             =   "ForSlotPerHubRedirectAddress"
         Top             =   2280
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Slot / Hub"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   43
         Left            =   120
         TabIndex        =   119
         Tag             =   "RedirectFSlotPerHub"
         Top             =   2260
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For KB / Slot"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   45
         Left            =   120
         TabIndex        =   122
         Tag             =   "RedirectFBWPerSlot"
         Top             =   3360
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   33
         Left            =   1560
         TabIndex        =   109
         Tag             =   "ForBWPerSlotRedirectAddress"
         Top             =   3375
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   35
         Left            =   1560
         TabIndex        =   110
         Tag             =   "ForFakeShareRedirectAddress"
         Top             =   3720
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   34
         Left            =   5160
         TabIndex        =   111
         Tag             =   "ForFakeTagRedirectAddress"
         Top             =   120
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Fake Tag"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   46
         Left            =   3720
         TabIndex        =   124
         Tag             =   "RedirectFFakeTag"
         Top             =   105
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Fake Share"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   47
         Left            =   120
         TabIndex        =   123
         Tag             =   "RedirectFFakeShare"
         Top             =   3705
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "Passive Mode"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   49
         Left            =   3720
         TabIndex        =   125
         Tag             =   "RedirectFPasMode"
         Top             =   465
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   24
         Left            =   5160
         TabIndex        =   112
         Tag             =   "ForPasModeRedirectAddress"
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   25
         Left            =   1560
         TabIndex        =   101
         Tag             =   "ForMaxShareRedirectAddress"
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   26
         Left            =   1560
         TabIndex        =   103
         Tag             =   "ForMaxSlotsRedirectAddress"
         Top             =   1200
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   29
         Left            =   1560
         TabIndex        =   104
         Tag             =   "ForMaxHubsRedirectAddress"
         Top             =   1560
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   30
         Left            =   1560
         TabIndex        =   105
         Tag             =   "ForNoTagRedirectAddress"
         Top             =   1920
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   27
         Left            =   1560
         TabIndex        =   102
         Tag             =   "ForMinSlotsRedirectAddress"
         Top             =   840
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   5
         Left            =   1560
         TabIndex        =   100
         Tag             =   "ForMinShareRedirectAddress"
         Top             =   120
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Max Share"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   37
         Left            =   120
         TabIndex        =   114
         Tag             =   "RedirectFMaxShare"
         Top             =   460
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Min Slots"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   39
         Left            =   120
         TabIndex        =   115
         Tag             =   "RedirectFMinSlots"
         Top             =   820
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Max Hubs"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   41
         Left            =   120
         TabIndex        =   117
         Tag             =   "RedirectFMaxHubs"
         Top             =   1540
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For No Tag"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   42
         Left            =   120
         TabIndex        =   118
         Tag             =   "RedirectFNoTag"
         Top             =   1900
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Max Slots"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   38
         Left            =   120
         TabIndex        =   116
         Tag             =   "RedirectFMaxSlots"
         Top             =   1180
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Min Share"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   113
         Tag             =   "RedirectFMinShare"
         Top             =   100
         Width           =   195
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Do not redirect"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   199
         ToolTipText     =   "Do not redirect anyone (disconnects if the hub is full)"
         Top             =   2400
         Width           =   4695
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Only if full and the user is not registered"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   198
         ToolTipText     =   "Only redirect if the user is not registered and the hub is full"
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Only if full and the user is not an op"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   197
         ToolTipText     =   "Only redirect if the user is an operator (or of a higher class) and the hub is full"
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Only if full"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   196
         ToolTipText     =   "Only redirect users if the hub is full"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered users"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   195
         ToolTipText     =   "Only unregistered users are redirected (must have ""Allow only registered users"" in Security/Advanced checked)"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "All users"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   194
         ToolTipText     =   "Redirects all users to your redirect address"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Allow ops to redirect (admins unaffected)"
         Height          =   375
         Index           =   30
         Left            =   3960
         TabIndex        =   193
         ToolTipText     =   "Check to allow operators to use the redirect ability (admins can always redirect)"
         Top             =   3000
         Width           =   4815
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Passive Mode"
         Height          =   375
         Index           =   49
         Left            =   3960
         TabIndex        =   192
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Fake Tag"
         Height          =   375
         Index           =   46
         Left            =   3960
         TabIndex        =   191
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Fake Share"
         Height          =   375
         Index           =   47
         Left            =   360
         TabIndex        =   190
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For KB / Slot"
         Height          =   375
         Index           =   45
         Left            =   360
         TabIndex        =   189
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Too Old NMDC"
         Height          =   495
         Index           =   40
         Left            =   360
         TabIndex        =   188
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Too Old DC++"
         Height          =   375
         Index           =   44
         Left            =   360
         TabIndex        =   187
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Slot / Hub"
         Height          =   375
         Index           =   43
         Left            =   360
         TabIndex        =   186
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For No Tag"
         Height          =   375
         Index           =   42
         Left            =   360
         TabIndex        =   185
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Max Hubs"
         Height          =   375
         Index           =   41
         Left            =   360
         TabIndex        =   184
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Max Slots"
         Height          =   375
         Index           =   38
         Left            =   360
         TabIndex        =   183
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Min Slots"
         Height          =   375
         Index           =   39
         Left            =   360
         TabIndex        =   182
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Max Share"
         Height          =   375
         Index           =   37
         Left            =   360
         TabIndex        =   181
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Min Share"
         Height          =   375
         Index           =   22
         Left            =   360
         TabIndex        =   180
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00C0C0C0&
         Height          =   1935
         Index           =   20
         Left            =   3720
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label lblHolder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Redirect Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   24
         Left            =   3840
         TabIndex        =   179
         Top             =   885
         Width           =   4815
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   3
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   138
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   3
         Left            =   7650
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   1395
         TabIndex        =   139
         Top             =   60
         Width           =   1400
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   161
         Top             =   430
         Width           =   8655
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Allow only registered users"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   8
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   81
            Tag             =   "RegOnly"
            ToolTipText     =   "Only registered users may connect to the hub"
            Top             =   3120
            Width           =   195
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   166
            Text            =   "0"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   165
            Text            =   "0"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   164
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   10
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   163
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Send to any user, including users below min class"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   240
            MaskColor       =   &H00000000&
            TabIndex        =   77
            Tag             =   "MinClsConnectSend"
            ToolTipText     =   "Strip all MyINFO before sending to unregistered users"
            Top             =   2160
            Width           =   195
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   12
            LargeChange     =   5
            Left            =   855
            Max             =   0
            Min             =   32000
            SmallChange     =   5
            TabIndex        =   73
            Tag             =   "MaxMessageLen"
            Top             =   840
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   11
            Left            =   855
            Max             =   0
            Min             =   11
            TabIndex        =   76
            Tag             =   "MinConnectCls"
            Top             =   1800
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   10
            Left            =   855
            Max             =   0
            Min             =   11
            TabIndex        =   74
            Tag             =   "MinSearchCls"
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Run in chat only mode"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   78
            Tag             =   "ChatOnly"
            ToolTipText     =   "Disables searching/connecting for all users; chatting in private messages and the main chat permitted"
            Top             =   2400
            Width           =   195
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   2
            Left            =   855
            Max             =   -1
            Min             =   99
            TabIndex        =   72
            Tag             =   "MinPassiveSearchLen"
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Send main chat messages to users in away mode"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   79
            Tag             =   "SendMessageAFK"
            ToolTipText     =   "All (NMDC) users who are in away mode will recieve main chat messages"
            Top             =   2640
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Hide MyINFOs to unregistered users"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   240
            TabIndex        =   80
            Tag             =   "HideMyinfos"
            Top             =   2880
            Width           =   195
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   855
            Max             =   1
            Min             =   32767
            TabIndex        =   71
            Tag             =   "MaxUsers"
            Top             =   120
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Send to all users, including users below min class"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   240
            MaskColor       =   &H00000000&
            TabIndex        =   75
            Tag             =   "MinClsSearchSend"
            ToolTipText     =   "Check to allow users above/equal to the min class to search users below the min class"
            Top             =   1560
            Width           =   195
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   9
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   162
            Text            =   "0"
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Allow only registered users"
            Height          =   255
            Index           =   10
            Left            =   480
            TabIndex        =   382
            ToolTipText     =   "Only registered users may connect to the hub"
            Top             =   3120
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show MyINFOs only to OPs"
            Height          =   255
            Index           =   36
            Left            =   480
            TabIndex        =   176
            Top             =   2880
            Width           =   6855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Send main chat messages to users in away mode"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   175
            ToolTipText     =   "All (NMDC) users who are in away mode will recieve main chat messages"
            Top             =   2640
            Width           =   7095
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Run in chat only mode"
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   174
            ToolTipText     =   "Disables searching/connecting for all users; chatting in private messages and the main chat permitted"
            Top             =   2400
            Width           =   7095
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Send to any user, including users below min class"
            Height          =   255
            Index           =   34
            Left            =   480
            TabIndex        =   173
            ToolTipText     =   "Strip all MyINFO before sending to unregistered users"
            Top             =   2160
            Width           =   7095
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Send to all users, including users below min class"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   172
            ToolTipText     =   "Check to allow users above/equal to the min class to search users below the min class"
            Top             =   1560
            Width           =   7095
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum main chat message length"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   88
            Left            =   1200
            TabIndex        =   171
            Top             =   840
            Width           =   6315
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum class required for downloading"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   85
            Left            =   1200
            TabIndex        =   170
            Top             =   1800
            Width           =   6195
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum class required for searching"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   1200
            TabIndex        =   169
            Top             =   1200
            Width           =   6315
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum passive search request length"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   31
            Left            =   1200
            TabIndex        =   168
            Top             =   480
            Width           =   6315
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Max Users"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   167
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   3
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   381
         Top             =   420
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "UpDate No-IP DNS at Starting PTDCH"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   70
            Left            =   480
            TabIndex        =   412
            Tag             =   "NoIPUpdateStartUp"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   120
            Width           =   195
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Force  Update "
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Index           =   6
            Left            =   120
            TabIndex        =   411
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   38
            Left            =   1680
            TabIndex        =   396
            Tag             =   "NoIPUser"
            ToolTipText     =   "No-IP DNS Service Account Name"
            Top             =   480
            Width           =   2484
         End
         Begin VB.TextBox txtData 
            Height          =   250
            IMEMode         =   3  'DISABLE
            Index           =   42
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   395
            Tag             =   "NoIPPass"
            ToolTipText     =   "No-IP DNS Service Password"
            Top             =   840
            Width           =   2484
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enable No-IP DNS Update(s)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   71
            Left            =   8040
            TabIndex        =   394
            Tag             =   "NoIPUpdateEna"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   120
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   41
            Left            =   4920
            TabIndex        =   393
            Tag             =   "NoIPDNS3"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   960
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   40
            Left            =   4920
            TabIndex        =   392
            Tag             =   "NoIPDNS2"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   600
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   39
            Left            =   1200
            TabIndex        =   391
            Tag             =   "NoIPDNS1"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   1320
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   37
            Left            =   4920
            TabIndex        =   390
            Tag             =   "NoIPDNS4"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   1320
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   43
            Left            =   4920
            TabIndex        =   389
            Tag             =   "DynDNS4"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   3000
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   44
            Left            =   1200
            TabIndex        =   388
            Tag             =   "DynDNS1"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   3000
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   45
            Left            =   4920
            TabIndex        =   387
            Tag             =   "DynDNS2"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   2280
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   46
            Left            =   4920
            TabIndex        =   386
            Tag             =   "DynDNS3"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   2640
            Width           =   2970
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enable Dyn DNS Update(s)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   69
            Left            =   8040
            TabIndex        =   385
            Tag             =   "DynDNSUpdateEna"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   1800
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Height          =   250
            IMEMode         =   3  'DISABLE
            Index           =   47
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   384
            Tag             =   "DynDNSPass"
            ToolTipText     =   "No-IP DNS Service Password"
            Top             =   2520
            Width           =   2484
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   48
            Left            =   1680
            TabIndex        =   383
            Tag             =   "DynDNSUser"
            ToolTipText     =   "No-IP DNS Service Account Name"
            Top             =   2160
            Width           =   2484
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "UpDate No-IP DNS at Starting PTDCH"
            Height          =   255
            Index           =   70
            Left            =   720
            TabIndex        =   413
            ToolTipText     =   "No-IP DNS Service"
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enable No-IP DNS Update(s)"
            Height          =   255
            Index           =   71
            Left            =   4200
            TabIndex        =   410
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 1"
            Height          =   255
            Index           =   30
            Left            =   360
            TabIndex        =   409
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 2"
            Height          =   255
            Index           =   34
            Left            =   4200
            TabIndex        =   408
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 3"
            Height          =   255
            Index           =   35
            Left            =   4200
            TabIndex        =   407
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 4"
            Height          =   255
            Index           =   36
            Left            =   4200
            TabIndex        =   406
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   29
            Left            =   720
            TabIndex        =   405
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   495
            Index           =   37
            Left            =   480
            TabIndex        =   404
            Top             =   840
            Width           =   1095
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   39
            X1              =   8280
            X2              =   8280
            Y1              =   1680
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   40
            X1              =   1200
            X2              =   8280
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enable Dyn DNS Update(s)"
            Height          =   255
            Index           =   69
            Left            =   4080
            TabIndex        =   403
            Top             =   1800
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 1"
            Height          =   255
            Index           =   38
            Left            =   360
            TabIndex        =   402
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 2"
            Height          =   255
            Index           =   39
            Left            =   4200
            TabIndex        =   401
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 3"
            Height          =   255
            Index           =   41
            Left            =   4200
            TabIndex        =   400
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 4"
            Height          =   255
            Index           =   42
            Left            =   4200
            TabIndex        =   399
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   43
            Left            =   720
            TabIndex        =   398
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   495
            Index           =   46
            Left            =   720
            TabIndex        =   397
            Top             =   2520
            Width           =   855
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   48
            X1              =   8280
            X2              =   8280
            Y1              =   3360
            Y2              =   1920
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   49
            X1              =   1200
            X2              =   8280
            Y1              =   3360
            Y2              =   3360
         End
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   152
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox txtData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   54
            Left            =   1680
            TabIndex        =   555
            Tag             =   "DBName"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton cmdDataBaseHelp 
            Caption         =   "Help"
            Height          =   330
            Left            =   240
            TabIndex        =   534
            Top             =   2400
            Width           =   1290
         End
         Begin VB.CommandButton cmdDataBaseApply 
            Caption         =   "Apply"
            Height          =   325
            Left            =   2040
            TabIndex        =   533
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cmbDataBase 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   532
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   50
            Left            =   240
            TabIndex        =   502
            Tag             =   "DBUserName"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtData 
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   51
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   501
            Tag             =   "DBPassword"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   52
            Left            =   240
            TabIndex        =   500
            Tag             =   "DBServerAddresse"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtData 
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   53
            Left            =   1680
            TabIndex        =   499
            Tag             =   "DBServerPort"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Register hub with public hub list"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   3600
            TabIndex        =   83
            Tag             =   "AutoRegister"
            ToolTipText     =   "Register the hub with the selected servers"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Compact user database on exit"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   3600
            TabIndex        =   84
            Tag             =   "CompactDBOnExit"
            ToolTipText     =   "Will compress the database to a smaller size on exit"
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Preload winsocks on start serving"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   3600
            TabIndex        =   82
            Tag             =   "PreloadWinsocks"
            ToolTipText     =   "Will preload sufficent (typically) connections assuming your entire hub was full (faster but uses more memory)"
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Update DynDNS / No-IP service"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   3600
            TabIndex        =   89
            Tag             =   "DynUpdate"
            Top             =   1680
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Check for updates on start up"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   3600
            TabIndex        =   85
            Tag             =   "AutoCheckUpdate"
            ToolTipText     =   "Will check for a new version of PTDCH on start up"
            Top             =   960
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Start serving on program start"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   3600
            TabIndex        =   86
            Tag             =   "AutoStart"
            ToolTipText     =   "Automatically start serving when PTDCH is opened"
            Top             =   1200
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Start minimized to the system tray"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   3600
            TabIndex        =   87
            Tag             =   "StartMinimized"
            ToolTipText     =   "When the hub is started, it will remain hidden in the system tray"
            Top             =   1440
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Start PTDCH at windows starting"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   3600
            MaskColor       =   &H00908675&
            TabIndex        =   88
            Tag             =   "StartWin"
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "DataBase Name"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   74
            Left            =   1680
            TabIndex        =   556
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   33
            X1              =   3480
            X2              =   3840
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Misc"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   10
            Left            =   3960
            TabIndex        =   508
            Top             =   0
            Width           =   2775
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   32
            X1              =   3480
            X2              =   8520
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   31
            X1              =   3480
            X2              =   3480
            Y1              =   3360
            Y2              =   120
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   30
            X1              =   120
            X2              =   3000
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Base Interface"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   225
            Index           =   9
            Left            =   600
            TabIndex        =   507
            Top             =   0
            Width           =   1995
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   21
            X1              =   120
            X2              =   3240
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   22
            X1              =   120
            X2              =   120
            Y1              =   3360
            Y2              =   120
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   23
            X1              =   120
            X2              =   480
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   70
            Left            =   240
            TabIndex        =   506
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   71
            Left            =   1680
            TabIndex        =   505
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Server (Host/IP)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   72
            Left            =   240
            TabIndex        =   504
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Remote Port"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   73
            Left            =   1680
            TabIndex        =   503
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Update DynDNS / No-IP service"
            Height          =   255
            Index           =   52
            Left            =   3840
            TabIndex        =   160
            Top             =   1680
            Width           =   4695
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Start minimized to the system tray"
            Height          =   255
            Index           =   26
            Left            =   3840
            TabIndex        =   159
            ToolTipText     =   "When the hub is started, it will remain hidden in the system tray"
            Top             =   1440
            Width           =   4695
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Start serving on program start"
            Height          =   255
            Index           =   9
            Left            =   3840
            TabIndex        =   158
            ToolTipText     =   "Automatically start serving when PTDCH is opened"
            Top             =   1200
            Width           =   4695
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Check for updates on start up"
            Height          =   255
            Index           =   28
            Left            =   3840
            TabIndex        =   157
            ToolTipText     =   "Will check for a new version of PTDCH on start up"
            Top             =   960
            Width           =   4695
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Register hub with public hub list"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   156
            ToolTipText     =   "Register the hub with the selected servers"
            Top             =   480
            Width           =   4695
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Compact user database on exit"
            Height          =   255
            Index           =   14
            Left            =   3840
            TabIndex        =   155
            ToolTipText     =   "Will compress the database to a smaller size on exit"
            Top             =   720
            Width           =   4695
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Start PTDCH at windows starting"
            Height          =   255
            Index           =   5
            Left            =   3840
            TabIndex        =   154
            Top             =   1920
            Width           =   3495
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Preload winsocks on start serving"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   153
            ToolTipText     =   "Will preload sufficent (typically) connections assuming your entire hub was full (faster but uses more memory)"
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   151
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on stoped serving"
            Height          =   195
            Index           =   72
            Left            =   240
            TabIndex        =   547
            Tag             =   "PopUpCoreError"
            Top             =   2160
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on stoped serving"
            Height          =   195
            Index           =   63
            Left            =   240
            TabIndex        =   521
            Tag             =   "PopUpStopedServing"
            Top             =   1920
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on started serving"
            Height          =   195
            Index           =   62
            Left            =   240
            TabIndex        =   520
            Tag             =   "PopUpStartedServing"
            Top             =   1680
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on user redirected"
            Height          =   195
            Index           =   60
            Left            =   240
            TabIndex        =   519
            Tag             =   "PopUpUserRedirected"
            Top             =   1440
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on user baned"
            Height          =   195
            Index           =   59
            Left            =   240
            TabIndex        =   518
            Tag             =   "PopUpUserBaned"
            Top             =   1200
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on user kicked"
            Height          =   195
            Index           =   58
            Left            =   240
            TabIndex        =   517
            Tag             =   "PopUpUserKick"
            Top             =   960
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on Op disconected"
            Height          =   195
            Index           =   57
            Left            =   240
            TabIndex        =   516
            Tag             =   "PopUpOpDisconected"
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on Op conected"
            Height          =   195
            Index           =   56
            Left            =   240
            TabIndex        =   515
            Tag             =   "PopUpOpConected"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on new user registed"
            Height          =   195
            Index           =   55
            Left            =   240
            TabIndex        =   514
            Tag             =   "PopUpNewReg"
            Top             =   240
            Width           =   195
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "Info"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   513
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "Warning"
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   512
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "Error"
            Height          =   375
            Index           =   2
            Left            =   3000
            TabIndex        =   511
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "App"
            Height          =   375
            Index           =   3
            Left            =   4320
            TabIndex        =   510
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "None"
            Height          =   375
            Index           =   4
            Left            =   5640
            TabIndex        =   509
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on core error"
            Height          =   255
            Index           =   72
            Left            =   480
            TabIndex        =   548
            Top             =   2160
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on new user registed"
            Height          =   255
            Index           =   55
            Left            =   480
            TabIndex        =   530
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on Op conected"
            Height          =   255
            Index           =   56
            Left            =   480
            TabIndex        =   529
            Top             =   480
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on Op disconected"
            Height          =   255
            Index           =   57
            Left            =   480
            TabIndex        =   528
            Top             =   720
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on user kicked"
            Height          =   255
            Index           =   58
            Left            =   480
            TabIndex        =   527
            Top             =   960
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on user baned"
            Height          =   255
            Index           =   59
            Left            =   480
            TabIndex        =   526
            Top             =   1200
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on user redirected"
            Height          =   255
            Index           =   60
            Left            =   480
            TabIndex        =   525
            Top             =   1440
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on started serving"
            Height          =   255
            Index           =   62
            Left            =   480
            TabIndex        =   524
            Top             =   1680
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on stoped serving"
            Height          =   255
            Index           =   63
            Left            =   480
            TabIndex        =   523
            Top             =   1920
            Width           =   5535
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   42
            X1              =   240
            X2              =   600
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   43
            X1              =   240
            X2              =   240
            Y1              =   3240
            Y2              =   2640
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   44
            X1              =   240
            X2              =   6960
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Poup test"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   14
            Left            =   720
            TabIndex        =   522
            Top             =   2520
            Width           =   1455
         End
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   4
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   140
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CheckBox chkData 
            Caption         =   "Enabled Skin"
            Height          =   195
            Index           =   66
            Left            =   4920
            TabIndex        =   94
            Tag             =   "blSkin"
            Top             =   480
            Width           =   195
         End
         Begin VB.ComboBox cmbSkin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   840
            Width           =   2535
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Allow usage"
            Height          =   195
            Index           =   65
            Left            =   480
            TabIndex        =   99
            Tag             =   "PriorityBl"
            Top             =   2040
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Move Form when clicking in any part"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   64
            Left            =   240
            TabIndex        =   93
            Tag             =   "MoveForm"
            Top             =   960
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enabled Magnetic Windows"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   54
            Left            =   240
            TabIndex        =   92
            Tag             =   "MagneticWin"
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Confirm exit"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   90
            Tag             =   "ConfirmExit"
            ToolTipText     =   "Ask if you want to exit the hub before unloading application"
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Minimize to system tray"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   240
            TabIndex        =   91
            Tag             =   "MinimizeTray"
            ToolTipText     =   "Check to have PTDCH minimize to the system tray"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Roundom skin at startup PTDCH"
            Enabled         =   0   'False
            Height          =   195
            Index           =   67
            Left            =   4920
            TabIndex        =   98
            Tag             =   "RndSkin"
            Top             =   1680
            Width           =   195
         End
         Begin VB.CommandButton cmdSkin 
            Caption         =   "<"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   4920
            TabIndex        =   96
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton cmdSkin 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   5280
            TabIndex        =   97
            Top             =   1200
            Width           =   375
         End
         Begin ComctlLib.Slider sldPriority 
            Height          =   315
            Left            =   360
            TabIndex        =   357
            Tag             =   "PriorityVal"
            Top             =   2280
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   3
            SelectRange     =   -1  'True
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Real Time"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   3
            Left            =   3480
            TabIndex        =   361
            Top             =   2580
            Width           =   810
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "High"
            Enabled         =   0   'False
            ForeColor       =   &H008080FF&
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   360
            Top             =   2580
            Width           =   420
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Normal"
            Enabled         =   0   'False
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   359
            Top             =   2580
            Width           =   495
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Idle"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   358
            Top             =   2580
            Width           =   270
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Enabled Skin"
            Height          =   255
            Index           =   66
            Left            =   5160
            TabIndex        =   150
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Skin"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   19
            Left            =   5160
            TabIndex        =   149
            Top             =   120
            Width           =   1455
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   59
            X1              =   4680
            X2              =   8520
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   58
            X1              =   4680
            X2              =   4680
            Y1              =   2040
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   57
            X1              =   4680
            X2              =   5040
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Allow usage"
            Height          =   255
            Index           =   65
            Left            =   720
            TabIndex        =   148
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "*Note: High and Real Time not recommended."
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   147
            Top             =   2955
            Width           =   3975
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "System Priority"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   18
            Left            =   720
            TabIndex        =   146
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   56
            X1              =   240
            X2              =   4320
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   55
            X1              =   240
            X2              =   240
            Y1              =   3240
            Y2              =   1800
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   54
            X1              =   240
            X2              =   600
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Move Form when clicking in any part"
            Height          =   255
            Index           =   64
            Left            =   480
            TabIndex        =   145
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Enabled Magnetic Windows"
            Height          =   255
            Index           =   54
            Left            =   480
            TabIndex        =   144
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimize to system tray"
            Height          =   255
            Index           =   35
            Left            =   480
            TabIndex        =   143
            ToolTipText     =   "Check to have PTDCH minimize to the system tray"
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm exit"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   142
            ToolTipText     =   "Ask if you want to exit the hub before unloading application"
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Roundom skin at startup PTDCH"
            Height          =   255
            Index           =   67
            Left            =   5160
            TabIndex        =   141
            Top             =   1680
            Width           =   3375
         End
      End
      Begin ComctlLib.TabStrip tabAdv 
         Height          =   3945
         Left            =   60
         TabIndex        =   177
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   5
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Hub/Users"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Application"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Notifications"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "DSN UpDate"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Miscllaneous"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   2
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   261
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   4
         Left            =   7680
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   1260
         TabIndex        =   262
         Top             =   60
         Width           =   1260
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   273
         Top             =   430
         Width           =   8715
         Begin VB.TextBox txtData 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   49
            Left            =   240
            TabIndex        =   488
            Tag             =   "BotEmail"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtData 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Index           =   6
            Left            =   2760
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   39
            Tag             =   "JoinMsg"
            ToolTipText     =   "Message that is sent when a user connects"
            Top             =   1680
            Width           =   5655
         End
         Begin VB.TextBox txtData 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   7
            Left            =   240
            TabIndex        =   34
            Tag             =   "BotName"
            ToolTipText     =   "Name of the bot (used for core messages)"
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton optJM 
            Appearance      =   0  'Flat
            Caption         =   "Do not send"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   2760
            TabIndex        =   37
            ToolTipText     =   "Do not send an on join message"
            Top             =   960
            Width           =   195
         End
         Begin VB.OptionButton optJM 
            Appearance      =   0  'Flat
            Caption         =   "Send as main chat message"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   2760
            TabIndex        =   36
            ToolTipText     =   "Sends to the user's main chat window (from bot name)"
            Top             =   720
            Width           =   195
         End
         Begin VB.OptionButton optJM 
            Appearance      =   0  'Flat
            Caption         =   "Send as private message"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   2760
            TabIndex        =   35
            ToolTipText     =   "Sends in a private message window (from bot name)"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Send messages in private messages (otherwise they are sent in main chat messages)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   2760
            TabIndex        =   38
            Tag             =   "SendMsgAsPrivate"
            ToolTipText     =   "Check to send reasons for failing various rules via a pm (from the default bot name)"
            Top             =   1200
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Bot name"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   240
            TabIndex        =   33
            Tag             =   "UseBotName"
            ToolTipText     =   "Check to have the bot listed in the user list"
            Top             =   240
            Width           =   195
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1335
            Index           =   16
            Left            =   120
            Top             =   120
            Width           =   2415
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   3255
            Index           =   13
            Left            =   2640
            Top             =   120
            Width           =   5895
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Bot E-mail"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   487
            ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Send messages in private messages (otherwise they are sent in main chat messages)"
            Height          =   435
            Index           =   27
            Left            =   3000
            TabIndex        =   279
            ToolTipText     =   "Check to send reasons for failing various rules via a pm (from the default bot name)"
            Top             =   1200
            Width           =   5385
         End
         Begin VB.Label lblOptJM 
            BackStyle       =   0  'Transparent
            Caption         =   "Do not send"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   278
            ToolTipText     =   "Do not send an on join message"
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblOptJM 
            BackStyle       =   0  'Transparent
            Caption         =   "Send as main chat message"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   277
            ToolTipText     =   "Sends to the user's main chat window (from bot name)"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblOptJM 
            BackStyle       =   0  'Transparent
            Caption         =   "Send as private message"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   276
            ToolTipText     =   "Sends in a private message window (from bot name)"
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Bot name"
            Height          =   255
            Index           =   21
            Left            =   480
            TabIndex        =   275
            ToolTipText     =   "Check to have the bot listed in the user list"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Welcome Message in here.(MOTD)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   7
            Left            =   3840
            TabIndex        =   274
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   3
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   263
         Top             =   430
         Visible         =   0   'False
         Width           =   8715
         Begin VB.TextBox txtTagRules 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   1035
            Left            =   3360
            MultiLine       =   -1  'True
            TabIndex        =   264
            Top             =   480
            Width           =   5175
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   9
            Left            =   6420
            TabIndex        =   66
            Tag             =   "NMDCMinVersion"
            ToolTipText     =   "Minimum client version for NMDC clients"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   17
            Left            =   6420
            TabIndex        =   67
            Tag             =   "DCMinVersion"
            ToolTipText     =   "0 = Disabled"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Deny all clients without a recognized tag"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   68
            Tag             =   "DenyNoTag"
            ToolTipText     =   "Disconnects all users without a <++ (or another supported) tag"
            Top             =   2400
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Prevent (certain) search tools from searching"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   3360
            TabIndex        =   69
            Tag             =   "PreventSearchBots"
            ToolTipText     =   "Prevents search tools such as MoGLO and DCSearch from searching (not connecting)"
            Top             =   2640
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Automatically kick MLDonkey clients"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   3360
            TabIndex        =   70
            Tag             =   "AutoKickMLDC"
            ToolTipText     =   "Automatically kick MLDonkey clients (recommended)"
            Top             =   2880
            Width           =   195
         End
         Begin VB.ListBox lstTagsEx 
            Height          =   2595
            Left            =   1800
            TabIndex        =   65
            Top             =   480
            Width           =   1455
         End
         Begin VB.ListBox lstTagsDef 
            BackColor       =   &H8000000F&
            Height          =   2595
            Left            =   120
            TabIndex        =   64
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Automatically kick MLDonkey clients"
            Height          =   255
            Index           =   8
            Left            =   3600
            TabIndex        =   272
            ToolTipText     =   "Automatically kick MLDonkey clients (recommended)"
            Top             =   2880
            Width           =   4935
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Prevent (certain) search tools from searching"
            Height          =   255
            Index           =   15
            Left            =   3600
            TabIndex        =   271
            ToolTipText     =   "Prevents search tools such as MoGLO and DCSearch from searching (not connecting)"
            Top             =   2640
            Width           =   4935
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Deny all clients without a recognized tag"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   270
            ToolTipText     =   "Disconnects all users without a <++ (or another supported) tag"
            Top             =   2400
            Width           =   4935
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum NMDC version"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   269
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum DC++ version"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   3480
            TabIndex        =   268
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Special handling that is built in"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   267
            Top             =   240
            Width           =   5235
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Accepted client tags"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   32
            Left            =   1800
            TabIndex        =   266
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Default client tags"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   33
            Left            =   120
            TabIndex        =   265
            Top             =   60
            Width           =   1455
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   5
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   306
         Top             =   430
         Visible         =   0   'False
         Width           =   8715
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   8
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   310
            Top             =   1440
            Width           =   855
         End
         Begin VB.VScrollBar vslData 
            Enabled         =   0   'False
            Height          =   255
            Index           =   13
            Left            =   3855
            Max             =   1
            Min             =   32555
            TabIndex        =   309
            Tag             =   "MinStrZBloc"
            Top             =   1440
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Log incoming"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   50
            Left            =   240
            TabIndex        =   308
            Tag             =   "LogIn"
            ToolTipText     =   "Check to send reasons for failing various rules via a pm (from the default bot name)"
            Top             =   2760
            Width           =   1572
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Log outgoing"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   51
            Left            =   240
            TabIndex        =   307
            Tag             =   "LogOut"
            ToolTipText     =   "Check to send reasons for failing various rules via a pm (from the default bot name)"
            Top             =   3000
            Width           =   1452
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "Length of string *2 (zbloc)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   90
            Left            =   2520
            TabIndex        =   312
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Reserve for SVN Build please no need to translate"
            Enabled         =   0   'False
            ForeColor       =   &H00808080&
            Height          =   495
            Left            =   5040
            TabIndex        =   311
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "SVN Debug Options"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   313
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   301
         Top             =   430
         Visible         =   0   'False
         Width           =   8715
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Scheduler"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   120
            TabIndex        =   63
            Tag             =   "EnabledScheduler"
            ToolTipText     =   "Check to enable sheduler"
            Top             =   1785
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   18
            Left            =   5520
            MaxLength       =   1
            TabIndex        =   61
            Tag             =   "CPrefix"
            ToolTipText     =   "Enter the prefix for which your built-in commands respond to (single character)"
            Top             =   1200
            Width           =   495
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Built-in command prefix (check to filter from main chat)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   5520
            TabIndex        =   60
            Tag             =   "FilterCPrefix"
            ToolTipText     =   "Filters messages starting with the prefix character from the main chat"
            Top             =   720
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   20
            Left            =   7680
            TabIndex        =   62
            Tag             =   "CSeperator"
            ToolTipText     =   "The seperator in which command params are seperated by"
            Top             =   1200
            Width           =   492
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enabled"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   5520
            TabIndex        =   59
            Tag             =   "EnabledCommands"
            ToolTipText     =   "Check to enable commands"
            Top             =   360
            Width           =   195
         End
         Begin ComctlLib.ListView lvwCommands 
            Height          =   1575
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   2778
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Command trigger"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Minimum class"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Enabled"
               Object.Width           =   1764
            EndProperty
         End
         Begin ComctlLib.ListView lvwPlan 
            Height          =   1335
            Left            =   120
            TabIndex        =   494
            Top             =   2040
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2355
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   9
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "User(s)"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Date/Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Enabled"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Command"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Parameter"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Increase"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Increase Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   7
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Description"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   8
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Status"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1575
            Index           =   14
            Left            =   5400
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Scheduler"
            Height          =   255
            Index           =   53
            Left            =   360
            TabIndex        =   305
            ToolTipText     =   "Check to enable sheduler"
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Built-in command prefix (check to filter from main chat)"
            Height          =   495
            Index           =   23
            Left            =   5760
            TabIndex        =   304
            ToolTipText     =   "Filters messages starting with the prefix character from the main chat"
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Enabled"
            Height          =   255
            Index           =   29
            Left            =   5760
            TabIndex        =   303
            ToolTipText     =   "Check to enable commands"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Seperator"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   48
            Left            =   6000
            TabIndex        =   302
            Top             =   1200
            Width           =   1695
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   280
         Top             =   430
         Visible         =   0   'False
         Width           =   8715
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   1
            ItemData        =   "frmHub.frx":18DD6
            Left            =   2520
            List            =   "frmHub.frx":18DE6
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Tag             =   "MinShareSize"
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   2
            ItemData        =   "frmHub.frx":18E00
            Left            =   2520
            List            =   "frmHub.frx":18E10
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Tag             =   "MaxShareSize"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   7920
            TabIndex        =   50
            Tag             =   "DCMaxHubs"
            ToolTipText     =   "0 = Disabled"
            Top             =   840
            Width           =   372
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   7920
            TabIndex        =   49
            Tag             =   "DCSlotsPerHub"
            ToolTipText     =   "0 = Disabled / Decimal values accepted (ex: 0.5 slots per hub)"
            Top             =   480
            Width           =   372
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   14
            Left            =   7440
            TabIndex        =   52
            Tag             =   "DCOSpeed"
            ToolTipText     =   "Grants the extra slot(s) when O: tag (if present) is equal or greater to this value"
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   4800
            TabIndex        =   51
            Tag             =   "DCOSlots"
            ToolTipText     =   "0 = Disabled"
            Top             =   1200
            Width           =   372
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   16
            Left            =   5160
            TabIndex        =   53
            Tag             =   "DCBandPerSlot"
            ToolTipText     =   "0 = Disabled / Decimal values accepted (ex: 4.5 kb/s per slot)"
            Top             =   1920
            Width           =   375
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Include hubs where user is an operator"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   4200
            TabIndex        =   54
            Tag             =   "DCIncludeOPed"
            ToolTipText     =   "In DC++ > 0.24 (among), include OPed hubs in hub count ?"
            Top             =   2280
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Deny Socks5 Connection"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   4200
            TabIndex        =   56
            Tag             =   "Denysocks5"
            ToolTipText     =   "Deny connection with socks5 (Validate tags must be enable)"
            Top             =   2760
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Deny Passive mode Connections"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   4200
            TabIndex        =   57
            Tag             =   "DenyPassive"
            ToolTipText     =   "Deny connection from client in Passive Mode (Validate tags must be enable)"
            Top             =   3000
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Validate tags (helps prevent fake tags)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   4200
            TabIndex        =   55
            Tag             =   "DCValidateTags"
            ToolTipText     =   "Kick client if anomaly in tags ? (such as H: or S: missing, wrong order, etc)"
            Top             =   2520
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   250
            Index           =   15
            Left            =   3360
            TabIndex        =   47
            Tag             =   "MinSlots"
            ToolTipText     =   "Minimum value for total slots of the client / 0 = Disabled"
            Top             =   2640
            Width           =   375
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   250
            Index           =   10
            Left            =   960
            LinkTimeout     =   0
            TabIndex        =   40
            Tag             =   "IMinShare"
            ToolTipText     =   "0 = Disabled / Minimum share size"
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Use mentoring system"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   240
            TabIndex        =   44
            Tag             =   "MentoringSystem"
            ToolTipText     =   "See Mentoring.txt for details - superior min share system"
            Top             =   1800
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Perform minor anti share faking checks"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   46
            Tag             =   "CheckFakeShare"
            ToolTipText     =   "Checks for traditional faking patterns (only the inexperienced are usually caught with this)"
            Top             =   2280
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   250
            Index           =   22
            Left            =   960
            LinkTimeout     =   0
            TabIndex        =   42
            Tag             =   "IMaxShare"
            ToolTipText     =   "0 = Disabled / Maximum share size"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   250
            Index           =   23
            Left            =   3360
            TabIndex        =   48
            Tag             =   "MaxSlots"
            ToolTipText     =   "Maximum value for total slots of the client / 0 = Disabled"
            Top             =   3000
            Width           =   375
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Ops/VIPs bypass all share and slot rules"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   45
            Tag             =   "OPBypass"
            ToolTipText     =   "Check to have Ops/VIPs bypass share and slot rules"
            Top             =   2040
            Width           =   195
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Deny Passive mode Connections"
            Height          =   255
            Index           =   48
            Left            =   4440
            TabIndex        =   300
            ToolTipText     =   "Deny connection from client in Passive Mode (Validate tags must be enable)"
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Deny Socks5 Connection"
            Height          =   255
            Index           =   25
            Left            =   4440
            TabIndex        =   299
            ToolTipText     =   "Deny connection with socks5 (Validate tags must be enable)"
            Top             =   2760
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Validate tags (helps prevent fake tags)"
            Height          =   255
            Index           =   11
            Left            =   4440
            TabIndex        =   298
            ToolTipText     =   "Kick client if anomaly in tags ? (such as H: or S: missing, wrong order, etc)"
            Top             =   2520
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Include hubs where user is an operator"
            Height          =   255
            Index           =   12
            Left            =   4440
            TabIndex        =   297
            ToolTipText     =   "In DC++ > 0.24 (among), include OPed hubs in hub count ?"
            Top             =   2280
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Perform minor anti share faking checks"
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   296
            ToolTipText     =   "Checks for traditional faking patterns (only the inexperienced are usually caught with this)"
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Ops/VIPs bypass all share and slot rules"
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   295
            ToolTipText     =   "Check to have Ops/VIPs bypass share and slot rules"
            Top             =   2040
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Use mentoring system"
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   294
            ToolTipText     =   "See Mentoring.txt for details - superior min share system"
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1695
            Index           =   19
            Left            =   120
            Top             =   1680
            Width           =   3855
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1455
            Index           =   18
            Left            =   120
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Tag options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   27
            Left            =   4200
            TabIndex        =   293
            Top             =   195
            Width           =   4215
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Hubs"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   5400
            TabIndex        =   292
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Slot/Hub"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   17
            Left            =   5400
            TabIndex        =   291
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "when upload speed <"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   22
            Left            =   4320
            TabIndex        =   290
            Top             =   1560
            Width           =   3015
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "KB/s"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   23
            Left            =   7920
            TabIndex        =   289
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "extra slot(s) if automated slot opens"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   21
            Left            =   5280
            TabIndex        =   288
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Grant"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   4320
            TabIndex        =   287
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Require"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   4200
            TabIndex        =   286
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "KB/s per slot (limiting upload)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   44
            Left            =   5640
            TabIndex        =   285
            Top             =   1920
            Width           =   2715
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum Slots"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   28
            Left            =   720
            TabIndex        =   284
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum share size"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   25
            Left            =   960
            TabIndex        =   283
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum share size"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   40
            Left            =   960
            TabIndex        =   282
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Slots"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   89
            Left            =   720
            TabIndex        =   281
            Top             =   3000
            Width           =   2535
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   3255
            Index           =   17
            Left            =   4080
            Top             =   120
            Width           =   4455
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   4
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   489
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdChatRom 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   1320
            TabIndex        =   493
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdChatRom 
            Caption         =   "Rem"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2520
            TabIndex        =   492
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdChatRom 
            Caption         =   "Add"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   491
            Top             =   3120
            Width           =   1215
         End
         Begin ComctlLib.ListView lvwChatRom 
            Height          =   2895
            Left            =   120
            TabIndex        =   490
            Top             =   120
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   5106
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   10
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Enabled"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Min Class"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Operator"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Share"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Connection"
               Object.Width           =   1765
            EndProperty
            BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   7
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "E-mail"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   8
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Tag"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   9
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Icon"
               Object.Width           =   882
            EndProperty
         End
      End
      Begin ComctlLib.TabStrip tbsInteractions 
         Height          =   3945
         Left            =   60
         TabIndex        =   32
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   5
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "General"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "User Controls"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Commands"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Clients Controls"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Chat Rooms"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   1
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   219
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   2
         Left            =   7640
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   1380
         TabIndex        =   220
         Top             =   60
         Width           =   1380
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleMode       =   0  'User
         ScaleWidth      =   8728.685
         TabIndex        =   260
         Top             =   430
         Width           =   8655
         Begin VB.CommandButton cmdDB 
            Caption         =   "Rename"
            Enabled         =   0   'False
            Height          =   300
            Index           =   8
            Left            =   3480
            TabIndex        =   378
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   300
            Index           =   6
            Left            =   2640
            TabIndex        =   376
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Rem"
            Enabled         =   0   'False
            Height          =   300
            Index           =   4
            Left            =   1800
            TabIndex        =   374
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Add"
            Height          =   300
            Index           =   2
            Left            =   960
            TabIndex        =   372
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Refresh"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   370
            Top             =   120
            Width           =   855
         End
         Begin VB.ComboBox cmbRegistered 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmHub.frx":18E2A
            Left            =   6360
            List            =   "frmHub.frx":18E2C
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   120
            Width           =   2175
         End
         Begin ComctlLib.ListView lvwRegistered 
            Height          =   2895
            Left            =   120
            TabIndex        =   368
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   5106
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "User Name"
               Object.Width           =   2624
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Password"
               Object.Width           =   2624
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Class"
               Object.Width           =   1050
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Class Name"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Reged By"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Reg Date"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Last Login"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   7
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Last IP"
               Object.Width           =   2519
            EndProperty
         End
         Begin VB.Label lblDBRegCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   6120
            TabIndex        =   535
            Top             =   200
            Width           =   90
         End
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleMode       =   0  'User
         ScaleWidth      =   8668.174
         TabIndex        =   252
         Top             =   430
         Visible         =   0   'False
         Width           =   8595
         Begin VB.PictureBox picBordTab 
            BorderStyle     =   0  'None
            Height          =   300
            Index           =   8
            Left            =   5682
            ScaleHeight     =   260
            ScaleMode       =   0  'User
            ScaleWidth      =   2940
            TabIndex        =   497
            Top             =   60
            Width           =   2940
         End
         Begin VB.TextBox txtBanFilter 
            Height          =   250
            Left            =   360
            TabIndex        =   17
            Top             =   2880
            Width           =   1575
         End
         Begin VB.OptionButton optBanFilter 
            Appearance      =   0  'Flat
            Caption         =   "Do not filter"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   16
            Top             =   2520
            Value           =   -1  'True
            Width           =   193
         End
         Begin VB.OptionButton optBanFilter 
            Appearance      =   0  'Flat
            Caption         =   "Begin with"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   13
            Top             =   1800
            Width           =   193
         End
         Begin VB.OptionButton optBanFilter 
            Appearance      =   0  'Flat
            Caption         =   "End in"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   15
            Top             =   2280
            Width           =   193
         End
         Begin VB.OptionButton optBanFilter 
            Appearance      =   0  'Flat
            Caption         =   "Contain"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   14
            Top             =   2040
            Width           =   193
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   0
            LargeChange     =   5
            Left            =   1320
            Max             =   -1
            Min             =   32767
            TabIndex        =   12
            Tag             =   "DefaultBanTime"
            Top             =   800
            Width           =   253
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   253
            Top             =   800
            Width           =   495
         End
         Begin ComctlLib.ListView lvwTempIPBan 
            Height          =   2835
            Left            =   2325
            TabIndex        =   11
            Top             =   480
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   5001
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "IP"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Expire"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Nick"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Banned By"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Reason"
               Object.Width           =   2519
            EndProperty
         End
         Begin ComctlLib.ListView lvwPermIPBan 
            Height          =   2835
            Left            =   2325
            TabIndex        =   10
            Top             =   480
            Visible         =   0   'False
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   5001
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "IP"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Nick"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Banned By"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Reason"
               Object.Width           =   2519
            EndProperty
         End
         Begin ComctlLib.TabStrip tabBansIPs 
            Height          =   3255
            Left            =   2280
            TabIndex        =   498
            Top             =   120
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5741
            TabWidthStyle   =   2
            TabFixedWidth   =   2973
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   2
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Temporary IP Bans"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Permanent IP Bans"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lblOptBanFilter 
            BackStyle       =   0  'Transparent
            Caption         =   "Do not filter"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   259
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblOptBanFilter 
            BackStyle       =   0  'Transparent
            Caption         =   "End in"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   258
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblOptBanFilter 
            BackStyle       =   0  'Transparent
            Caption         =   "Contain"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   257
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblOptBanFilter 
            BackStyle       =   0  'Transparent
            Caption         =   "Begin with"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   256
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   2055
            Index           =   8
            Left            =   120
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "IP Ban Filter"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   83
            Left            =   240
            TabIndex        =   255
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1095
            Index           =   7
            Left            =   120
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Default kick temp ban length (minutes)"
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   15
            Left            =   240
            TabIndex        =   254
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   3
         Left            =   120
         ScaleHeight     =   3495
         ScaleMode       =   0  'User
         ScaleWidth      =   8728.685
         TabIndex        =   224
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   7
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   231
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   230
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   229
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   228
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   227
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   226
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   225
            Top             =   1060
            Width           =   495
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enable flood wall"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   8
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   7920
            TabIndex        =   24
            Tag             =   "EnableFloodWall"
            ToolTipText     =   "Aids in preventing flooding via traditional means (ie MyINFO, nicklist, search, etc)"
            Top             =   600
            Width           =   193
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   3
            LargeChange     =   100
            Left            =   7575
            Max             =   1000
            Min             =   32000
            SmallChange     =   100
            TabIndex        =   25
            Tag             =   "FWInterval"
            Top             =   960
            Value           =   1000
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   4
            LargeChange     =   10
            Left            =   7215
            Max             =   -1
            Min             =   32767
            TabIndex        =   26
            Tag             =   "FWBanLength"
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   5
            LargeChange     =   100
            Left            =   8055
            Max             =   1
            Min             =   254
            TabIndex        =   29
            Tag             =   "FWMyINFO"
            Top             =   2520
            Value           =   254
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   6
            LargeChange     =   100
            Left            =   8055
            Max             =   1
            Min             =   254
            TabIndex        =   31
            Tag             =   "FWGetNickList"
            Top             =   2880
            Value           =   254
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   7
            LargeChange     =   100
            Left            =   6255
            Max             =   1
            Min             =   254
            TabIndex        =   28
            Tag             =   "FWActiveSearch"
            Top             =   2520
            Value           =   254
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   8
            LargeChange     =   100
            Left            =   6255
            Max             =   1
            Min             =   254
            TabIndex        =   30
            Tag             =   "FWPassiveSearch"
            Top             =   2880
            Value           =   254
            Width           =   255
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   19
            Left            =   7320
            MaxLength       =   10
            TabIndex        =   27
            Tag             =   "DataFragmentLen"
            Text            =   "2048"
            ToolTipText     =   "Limit messages and protocol commands length. Be careful not to set too low. Default is 2048"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Redirect users who give wrong password"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   240
            TabIndex        =   22
            Tag             =   "RedirectFGP"
            ToolTipText     =   "Redirects users who do not know the password to the redirect hub set in General settings"
            Top             =   2280
            Width           =   193
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Run in password mode"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Tag             =   "PasswordMode"
            ToolTipText     =   "Require all unregistered users to send the password specified below before logging in"
            Top             =   2040
            Width           =   193
         End
         Begin VB.TextBox txtData 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   250
            Index           =   21
            Left            =   2160
            TabIndex        =   23
            Tag             =   "HubPassword"
            ToolTipText     =   "Global password used in password mode"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Use descriptive ban messages"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   18
            Tag             =   "DescriptiveBanMsg"
            ToolTipText     =   "ex. ""Your IP is permanently banned."" versus ""Your IP is banned!"""
            Top             =   240
            Width           =   194
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Prevent brute force password guessing"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   8
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   19
            Tag             =   "PreventGuessPass"
            ToolTipText     =   "Limits the number of times you may attempt to log in before being temporarily banned"
            Top             =   600
            Width           =   194
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   9
            Left            =   1800
            Max             =   1
            Min             =   10
            TabIndex        =   20
            Tag             =   "MaxPassAttempts"
            Top             =   1080
            Value           =   10
            Width           =   255
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enable flood wall"
            Height          =   255
            Index           =   20
            Left            =   4440
            TabIndex        =   251
            ToolTipText     =   "Aids in preventing flooding via traditional means (ie MyINFO, nicklist, search, etc)"
            Top             =   600
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Redirect users who give wrong password"
            Height          =   495
            Index           =   31
            Left            =   480
            TabIndex        =   250
            ToolTipText     =   "Redirects users who do not know the password to the redirect hub set in General settings"
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Run in password mode"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   249
            ToolTipText     =   "Require all unregistered users to send the password specified below before logging in"
            Top             =   2040
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Prevent brute force password guessing"
            Height          =   375
            Index           =   24
            Left            =   480
            TabIndex        =   248
            ToolTipText     =   "Limits the number of times you may attempt to log in before being temporarily banned"
            Top             =   600
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Use descriptive ban messages"
            Height          =   375
            Index           =   17
            Left            =   480
            TabIndex        =   247
            ToolTipText     =   "ex. ""Your IP is permanently banned."" versus ""Your IP is banned!"""
            Top             =   240
            Width           =   3375
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   3255
            Index           =   11
            Left            =   4080
            Top             =   120
            Width           =   4455
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Flood wall options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   61
            Left            =   4320
            TabIndex        =   246
            Top             =   240
            Width           =   4035
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Flooding interval checks last"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   62
            Left            =   4200
            TabIndex        =   245
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "ms"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   63
            Left            =   7920
            TabIndex        =   244
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "MyINFO"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   64
            Left            =   6600
            TabIndex        =   243
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nicklist"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   65
            Left            =   6480
            TabIndex        =   242
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Active search"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   66
            Left            =   4080
            TabIndex        =   241
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Passive search"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   67
            Left            =   4080
            TabIndex        =   240
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ban user if flooding for"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   68
            Left            =   4200
            TabIndex        =   239
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "minutes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   69
            Left            =   7560
            TabIndex        =   238
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Number of permitted sendings during interval "
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   76
            Left            =   4200
            TabIndex        =   237
            Top             =   2160
            Width           =   4215
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Max messages and protocol length"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   4200
            TabIndex        =   236
            ToolTipText     =   "Limit messages and protocol commands length. Be careful not to set too low. Default is 2048"
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1335
            Index           =   10
            Left            =   120
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Global password mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   78
            Left            =   480
            TabIndex        =   235
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   77
            Left            =   480
            TabIndex        =   234
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1815
            Index           =   9
            Left            =   120
            Top             =   1560
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Permit"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   80
            Left            =   240
            TabIndex        =   233
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "attempts"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   81
            Left            =   2160
            TabIndex        =   232
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   4
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   426
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin ComctlLib.ListView lvwSqlExplorer 
            Height          =   2415
            Left            =   120
            TabIndex        =   545
            Top             =   840
            Visible         =   0   'False
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   4260
            View            =   3
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.PictureBox pciDBExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1080
            ScaleHeight     =   225
            ScaleWidth      =   465
            TabIndex        =   538
            Top             =   480
            Width           =   495
            Begin VB.Label lblQueryDB 
               Alignment       =   2  'Center
               Caption         =   "Cls"
               Height          =   255
               Index           =   1
               Left            =   0
               MouseIcon       =   "frmHub.frx":18E2E
               MousePointer    =   99  'Custom
               TabIndex        =   539
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.PictureBox pciDBExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   160
            ScaleHeight     =   225
            ScaleWidth      =   825
            TabIndex        =   536
            Top             =   480
            Width           =   855
            Begin VB.Label lblQueryDB 
               Alignment       =   2  'Center
               Caption         =   "Run"
               Height          =   255
               Index           =   0
               Left            =   0
               MouseIcon       =   "frmHub.frx":18F80
               MousePointer    =   99  'Custom
               TabIndex        =   537
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.PictureBox picBordTab 
            BorderStyle     =   0  'None
            Height          =   300
            Index           =   7
            Left            =   3000
            ScaleHeight     =   260
            ScaleMode       =   0  'User
            ScaleWidth      =   5580
            TabIndex        =   429
            Top             =   60
            Width           =   5580
         End
         Begin VB.TextBox txtSqlErr 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   427
            Top             =   480
            Width           =   6735
         End
         Begin VB.PictureBox picSqlSCI 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2415
            ScaleWidth      =   8295
            TabIndex        =   546
            Top             =   840
            Width           =   8295
         End
         Begin ComctlLib.TabStrip tbsDbManager 
            Height          =   3375
            Left            =   60
            TabIndex        =   428
            Top             =   60
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5953
            TabWidthStyle   =   2
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   2
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Query String"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Data Contents"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleMode       =   0  'User
         ScaleWidth      =   8728.685
         TabIndex        =   221
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdDB 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   300
            Index           =   7
            Left            =   2640
            TabIndex        =   377
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Rem"
            Enabled         =   0   'False
            Height          =   300
            Index           =   5
            Left            =   1800
            TabIndex        =   375
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Add"
            Height          =   300
            Index           =   3
            Left            =   960
            TabIndex        =   373
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Refresh"
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   371
            Top             =   120
            Width           =   855
         End
         Begin ComctlLib.ListView lvwBans 
            Height          =   1815
            Left            =   120
            TabIndex        =   369
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "User Name"
               Object.Width           =   3498
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Perm"
               Object.Width           =   1050
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Banned by"
               Object.Width           =   3498
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Reference Date"
               Object.Width           =   3848
            EndProperty
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   975
            Index           =   12
            Left            =   120
            Top             =   2400
            Width           =   8415
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reason :"
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
            Index           =   49
            Left            =   960
            TabIndex        =   223
            Top             =   2520
            Width           =   6615
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            DataField       =   "Reason"
            DataSource      =   "adoBans"
            Height          =   495
            Index           =   50
            Left            =   240
            TabIndex        =   222
            Top             =   2760
            Width           =   8175
         End
      End
      Begin ComctlLib.TabStrip tbsSecurity 
         Height          =   3945
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   5
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Registed Users"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Name Bans"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "IP Bans"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Advanced"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "BD Explorer"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip tbsMenu 
      Height          =   4575
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   8070
      TabWidthStyle   =   2
      TabFixedWidth   =   1785
      TabFixedHeight  =   489
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   9
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Security"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Interactions"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Redirections"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Scripts"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Status"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Full Help"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Tray Menu"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuTray 
         Caption         =   "Show"
         Index           =   0
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Hide"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Registered user list"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuRegistered 
         Caption         =   "Copy User Name"
         Index           =   0
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "Copy Password"
         Index           =   1
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "Copy Last IP"
         Index           =   2
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "Copy All"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Temp IP ban list"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Remove"
         Index           =   1
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Clear"
         Index           =   2
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Refresh list extract"
         Index           =   4
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Clear list extract"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Perm IP ban list"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Remove"
         Index           =   1
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Clear"
         Index           =   2
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Refresh list extract"
         Index           =   4
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Clear list extract"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Locked names"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnuLocked 
         Caption         =   "Copy User Name"
         Index           =   0
      End
      Begin VB.Menu mnuLocked 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLocked 
         Caption         =   "Copy All"
         Index           =   2
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Tags"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu mnuTags 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuTags 
         Caption         =   "Remove"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Status"
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu mnuStatus 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Clear"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Users"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu mnuUsers 
         Caption         =   "Send data (selected)"
         Index           =   0
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Send data (all)"
         Index           =   1
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Disconnect"
         Index           =   2
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Kick"
         Index           =   3
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Redirect"
         Index           =   4
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Ban"
         Index           =   5
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "(De)mute"
         Index           =   6
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Properties (selected)"
         Index           =   7
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Plan"
      Index           =   8
      Visible         =   0   'False
      Begin VB.Menu mnuPlan 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuPlan 
         Caption         =   "Edit"
         Index           =   1
      End
      Begin VB.Menu mnuPlan 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPlan 
         Caption         =   "Remove"
         Index           =   3
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Scripts"
      Index           =   9
      Visible         =   0   'False
      Begin VB.Menu mnuScripts 
         Caption         =   "Reset / Save"
         Index           =   0
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Stop"
         Index           =   2
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Stop All"
         Index           =   3
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Reolad (Checkeds)"
         Index           =   5
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Reolad Dir"
         Index           =   6
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Properties"
         Index           =   8
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "SCI"
      Index           =   10
      Visible         =   0   'False
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "View WhiteSpace"
         Index           =   0
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Line Number"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Misc"
         Index           =   3
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Insert Date/Time"
            Index           =   0
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Script Info.."
            Index           =   1
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Scripts Help"
            Index           =   2
            Begin VB.Menu mnuCodeRTB3 
               Caption         =   "VBScript Documentation"
               Index           =   0
            End
            Begin VB.Menu mnuCodeRTB3 
               Caption         =   "JScript Documentation"
               Index           =   1
            End
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Save as Script.."
            Index           =   4
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Clear Undo Buffer"
            Index           =   6
         End
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Help"
         Index           =   5
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Word Wrap"
         Index           =   7
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "ReadOnly"
         Index           =   8
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Plug-ins"
         Index           =   10
         Begin VB.Menu mnuPlugIn 
            Caption         =   "No Plug-Ins Found"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Compiler conditions
'
'PreConnectionRequest - Setting this value to true, turns on the optional "PreConnectionRequest"
'                       option; this adds a new Function to the mix called PreConnectionRequest
#Const PreConnectionRequest = True

'PredataArrival - Setting this value to true, turns on the optional "predataarrival"
'                 option; this adds a new sub to the mix called PreDataArrival
#Const PreDataArrival = True

'DataArrival - Setting this value to true, turns on the default event DataArrival.
'              It is a CPU intensive event, and if you are not using it in your scripts
'              I suggest you set this value to false
#Const DataArrival = True

'ObjectNotSet - Makes a check in wskLoop_DataArrival to make sure user object exists
#Const OBJECTNOTSET = True

'ColFreeSocks - Uses a collection to find free winsocks (otherwise it loops)
#Const COLFREESOCKS = True

'Status window - Setting this value to true turns on the Status / Admin panel.
'                Must be set in the Properties dialog (just included here for clarity)
#Const Status = True

'API calls
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
Private Const SW_SHOWMAXIMIZED = 3
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'//shellexecute constant
Private Const SW_NORMAL As Long = 1

'API calls MyIPTools.dll /////////////////////////////////////////////////////////////
Private Declare Function DetectIP Lib "MyIPTools.dll" () As Variant
Private Declare Function ResolveHost Lib "MyIPTools.dll" (sIP As Variant) As Variant
Private Declare Function UpdateDynDNS Lib "MyIPTools.dll" (User As Variant, Pass As Variant, Host As Variant, auto As Boolean, sIP As Variant, mail As Boolean) As Variant
Private Declare Function UpdateNoIP Lib "MyIPTools.dll" (User As Variant, Pass As Variant, Host As Variant, sIP As Variant) As Variant
Private Declare Function IPinRange Lib "MyIPTools.dll" (sIP As Variant, eIP As Variant, IP As Variant) As Boolean
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'Types
Private Type typBot
    Name As String
    MyINFO As String
    Operator As Boolean
End Type

'Private objects
Private WithEvents m_objDetectIP    As clsHTTPDownload
Attribute m_objDetectIP.VB_VarHelpID = -1
Private m_objLoopUser               As clsUser
Private m_wskLoopItem               As Winsock
Private m_colTags                   As Collection
Private m_colFailedReg              As Collection
Private m_colConnectAttempts        As Collection
Private m_colRevConnects            As Collection

#If COLFREESOCKS Then
    Private m_colFreeSocks          As Collection
#End If

'Array of API Timers Objects used by scripts or plugins
Private WithEvents m_objTimers      As clsTimersCol
Attribute m_objTimers.VB_VarHelpID = -1

'#If PREDATAARRIVAL Then
'    Private m_intPDIndex           As Integer
'#End If

'Cool FX m_ObjMagnetic Windows
Private m_ObjMagnetic               As New clsMagneticWnd

Private m_objSciExplorerSQL         As clsYScintilla
Attribute m_objSciExplorerSQL.VB_VarHelpID = -1

'Private vars
Private m_arrScriptEvents()         As Boolean
Private m_blnCommaDecimal           As Boolean
Private m_lngScriptEventsUB         As Long
Private m_lngRedirectUB             As Long
Private m_lngBotsUB                 As Long
Private m_lngBanFilter              As Long
Private m_arrRedirectIPs()          As String
Private m_arrBots()                 As typBot
Private m_datServingDate            As Date
Private m_datForceDNSUpdate         As Date

' NEW INTERFACE LANGUAGE /////////////////////////////////////////////////////////////
Private m_arrDynaCap(2)             As String
'limit and initialise for 51 clients tags
Private m_arrTagRules(50)           As String
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Const IPCheckInterval = 10  ' => every X minutes

Private lIntervalMin                As Long

Private Service(9)                  As String
Private Host(9)                     As String
Private User(9)                     As String
Private Pass(9)                     As String

'Private vars for sys tray
Private m_lHookID                   As Long
Private m_bSound                    As Boolean
'allows the use of 'tab' with in the SCI
Private m_TabsStop()                As Boolean

'Objects for data base
Private m_objRS                     As ADODB.Recordset
Private m_objPermaCon               As ADODB.Connection
Private m_DbType                    As enInterfaceDB

Public Sub APP_TERMINATE(Optional ByVal blnRestarEXE As Boolean = False)
    '------------------------------------------------------------------
    'Purpose:   This Sub Terminate the aplication
    '           Unload all objects and clear the reference in memory..
    'Params:    blnRestarEXE=True : On Exit, Auto Restart the application
    '
    'Returns:   none
    '
    '   Called by unload App, plugins or scripts
    '------------------------------------------------------------------

11:    If (G_GUI_IN_UNLOAD) Or (Not G_GUI_IS_LOADED) Then Exit Sub

13:    Dim frmLoop As Form
14:    Dim i       As Integer
15:    Dim lngTemp As Long
16:    Dim strTemp As String
       
       '***********************************************************************************
19:    On Error GoTo QueryUnload
       '***********************************************************************************

       'Hide hub Form
23:    frmHub.Hide

       'Close hub if it's still serving
26:    If G_SERVING Then SwitchServing
    
       'Save Scripts Value in XML file
29:    Call frmScript.XmlBooleanSave

       'Confirm the unloading
32:    G_GUI_IN_UNLOAD = True

34:    Call SaveSettings

36:    tmrSysInfo.Enabled = False
    
       'Call unload event
39:    Call SEvent_UnloadMain
    
      'Save plugin state to XML
42:    Call PlgXmlSave

       'Remove all stray bot names from the database
45:    If Not m_lngBotsUB = -1 Then
46:        For i = 0 To m_lngBotsUB
47:            g_objRegistered.Remove m_arrBots(i).Name
48:        Next
49:        Erase m_arrBots()
50:    End If

       'Close connection to database
53:    DBConnectionClose

       'Remove systray icon
56:    Call SysTrayRem
    
       '***********************************************************************************
59:    On Error GoTo Unload
       '***********************************************************************************

       'This is absolutly an imperative line
63:    For i = 1 To UBound(g_objSciLexer)
64:        Call g_objSciLexer(i).Detach(picSciMain(i))
65:        Set g_objSciLexer(i) = Nothing
66:    Next
67:    Call m_objSciExplorerSQL.Detach(picSqlSCI)
68:    Set m_objSciExplorerSQL = Nothing

       'Unload all scripts control
71:    For i = 1 To ScriptControl.UBound
72:        Unload ScriptControl(i)
73:    Next

       '***********************************************************************************
76:    On Error GoTo Terminate
       '***********************************************************************************

       'Terminate plugins
80:    Call PlgTerm

       'Compress database if needed (only for MS Access Data base interface)
83:    If g_objSettings.CompactDBOnExit And m_DbType = MsAccess Then
84:        Dim objEngine As JetEngine
        
86:        Set objEngine = New JetEngine
87:        objEngine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=.\DBs\userdb.mdb", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=.\DBs\tempdb.mdb"
88:        Set objEngine = Nothing
89:        Kill ".\DBs\userdb.mdb"
90:        Name ".\DBs\tempdb.mdb" As "DBs\userdb.mdb"
91:    End If

       'Public objects
94:    If Not g_objTimer Is Nothing Then Set g_objTimer = Nothing
95:    If Not g_objTimersCol Is Nothing Then Set g_objTimersCol = Nothing
96:    If Not g_objSCI Is Nothing Then Set g_objSCI = Nothing
97:    If Not g_objComDialog Is Nothing Then Set g_objComDialog = Nothing
98:    If Not g_objSQLite Is Nothing Then Set g_objSQLite = Nothing
99:    If Not g_colDictionary Is Nothing Then Set g_colDictionary = Nothing
       '
101:    If Not g_objChatRoom Is Nothing Then Set g_objChatRoom = Nothing
102:    If Not g_objScheduler Is Nothing Then Set g_objScheduler = Nothing
103:    If Not g_colUsers Is Nothing Then Set g_colUsers = Nothing
104:    If Not g_colIPBans Is Nothing Then Set g_colIPBans = Nothing
105:    If Not g_objRegistered Is Nothing Then Set g_objRegistered = Nothing
106:    If Not g_colCommands Is Nothing Then Set g_colCommands = Nothing
107:    If Not g_objRegExps Is Nothing Then Set g_objRegExps = Nothing
108:    If Not g_colLanguages Is Nothing Then Set g_colLanguages = Nothing
109:    If Not g_objHighlighter Is Nothing Then Set g_objHighlighter = Nothing
110:    If Not g_objActiveX Is Nothing Then Set g_objActiveX = Nothing
111:    If Not g_objAbout Is Nothing Then Set g_objAbout = Nothing
112:    If Not g_colSWinsocks Is Nothing Then Set g_colSWinsocks = Nothing
113:    If Not g_colSVariables Is Nothing Then Set g_colSVariables = Nothing
114:    If Not g_colToolTip Is Nothing Then Set g_colToolTip = Nothing
115:    If Not g_objFunctions Is Nothing Then Set g_objFunctions = Nothing
116:    If Not g_objFileAccess Is Nothing Then Set g_objFileAccess = Nothing
117:    If Not g_objSettings Is Nothing Then Set g_objSettings = Nothing
#If Status Then
119:    If Not g_objStatus Is Nothing Then Set g_objStatus = Nothing
#End If
        'Private objects
#If COLFREESOCKS Then
123:    If Not m_colFreeSocks Is Nothing Then Set m_colFreeSocks = Nothing
#End If
125:    If Not m_objLoopUser Is Nothing Then Set m_objLoopUser = Nothing
126:    If Not m_wskLoopItem Is Nothing Then Set m_wskLoopItem = Nothing
127:    If Not m_colTags Is Nothing Then Set m_colTags = Nothing
128:    If Not m_objDetectIP Is Nothing Then Set m_objDetectIP = Nothing
129:    If Not m_colFailedReg Is Nothing Then Set m_colFailedReg = Nothing
130:    If Not m_colConnectAttempts Is Nothing Then Set m_colConnectAttempts = Nothing
131:    If Not m_colRevConnects Is Nothing Then Set m_colRevConnects = Nothing
132:    If Not m_ObjMagnetic Is Nothing Then Set m_ObjMagnetic = Nothing
133:    If Not m_objRS Is Nothing Then Set m_objRS = Nothing
134:    If Not m_objPermaCon Is Nothing Then Set m_objPermaCon = Nothing
135:    If Not m_objTimers Is Nothing Then Set m_objTimers = Nothing

137:    Erase g_arrHighlighters()
138:    Erase g_objSciLexer()
139:    Erase g_arrToolTips()
140:    Erase g_arrHighlighters()
141:    Erase g_objSciLexer()

        'Unload any other forms left over
147:    For Each frmLoop In Forms
148:        If Not frmLoop.Name = Me.Name Then
149:            Call Unload(frmLoop)
150:            Set frmLoop = Nothing
151:        End If
152:    Next
        
154:    Call Unload(Me)
         
        'Close the error file
155:    Close G_ERRORFILE

156:    If blnRestarEXE Then
157:        GoTo RestartEXE
158:    End If
        
160:    Exit Sub
161:
QueryUnload:
162:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.APP_TERMINATE(QueryUnload)"
163:    Resume Next
    
165:    Exit Sub
166:
Unload:
167:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.APP_TERMINATE(Unload)"
168:    Resume Next
    
170:    Exit Sub
171:
Terminate:
172:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.APP_TERMINATE(Terminate)"
173:    Resume Next

175:    Exit Sub
176:
RestartEXE:
177:    On Error Resume Next
        'Save temp string to get in next load..
179:    SaveSetting "PTDCH", "Shell", "Parameter", "True"
180:    strTemp = App.Path & Chr$(92) & App.EXEName & ".exe"
181:    lngTemp = GetDesktopWindow()
182:    ShellExecute lngTemp, "Open", strTemp, vbNullString, vbNullString, SW_NORMAL
183:    End 'Hard end
End Sub
'------------------------------------------------------------------------------
'Start - Data Base Interace Subs/Functions
'------------------------------------------------------------------------------
Private Sub cmbRegistered_Click()
1:    On Error GoTo Err
2:    If cmbRegistered.ListCount = 4 Then
3:       Static IsLoaded As Boolean
4:       If IsLoaded Then 'Not duplic the load data at start up..
5:          Call DBGetRegRecord
6:       Else
7:          IsLoaded = True
8:       End If
9:    End If
10:   Exit Sub
Err:
13    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbRegistered_Click()"
End Sub
Private Sub DBGetBanRecord()
1:    Dim lvwItem       As ListItem
2:    Dim lvwItems      As ListItems
3:    Dim strQuery      As String
4:    On Error GoTo Err
 
6:    lvwBans.ListItems.Clear

8:    strQuery = "SELECT UsrClass.UserName, BanNames.Perm, BanNames.BannedBy, BanNames.RefDate, BanNames.Reason " & _
                 "FROM UsrClass INNER JOIN BanNames ON UsrClass.UserName = BanNames.UserName " & _
                 "ORDER BY BanNames.RefDate;"
    
12:   Set m_objRS = m_objPermaCon.Execute(strQuery)
13:   Set lvwItems = lvwBans.ListItems
        
15:   Do While Not m_objRS.EOF
        
17:        Set lvwItem = lvwItems.Add(, , CStr(m_objRS(0).Value)) 'UserName
        
19:        lvwItem.SubItems(1) = CBool(m_objRS(1).Value) 'Perm
20:        lvwItem.SubItems(2) = CStr(m_objRS(2).Value) 'BannedBy
21:        lvwItem.SubItems(3) = CStr(m_objRS(3).Value) 'RefDate
        
23:        If Not m_objRS(4).Value = Empty Then 'Reason
24:            lvwItem.Tag = CStr(m_objRS(4).Value)
25:        End If
        
27:        m_objRS.MoveNext
    
29:   Loop
    
31:   If Not m_objRS Is Nothing Then Set m_objRS = Nothing
    
33:   Exit Sub
34:
Err:
35:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.DBGetBanRecord()"
End Sub

Private Sub DBGetRegRecord()
1:    Dim lvwItem       As ListItem
2:    Dim lvwItems      As ListItems
3:    Dim strQuery      As String
4:    Dim strQueryAux   As String
5:    On Error GoTo Err

7:    lvwRegistered.ListItems.Clear

      Select Case cmbRegistered.ListIndex
        Case 0 'All Classes
9:            strQueryAux = "ORDER BY UsrClass.UserName;"
        Case 1 'Non-OPs only
10:           strQueryAux = "WHERE ((UsrClass.Class > 1) And (UsrClass.Class < 6)) ORDER BY UsrClass.UserName;"
        Case 2 'OPs and above
11:           strQueryAux = "WHERE ((UsrClass.Class) > 5) ORDER BY UsrClass.UserName;"
        Case 3 'Admins and above
12:           strQueryAux = "WHERE ((UsrClass.Class) > 9) ORDER BY UsrClass.UserName;"
13:    End Select

15:    strQuery = "SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP " & _
                  "FROM UsrStatic INNER JOIN (ClassTypes INNER JOIN (UsrClass INNER JOIN UsrDynamic ON UsrClass.UserName = UsrDynamic.UserName) ON ClassTypes.ID = UsrClass.Class) ON (UsrStatic.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrClass.UserName) " & _
                  strQueryAux

19:    Set m_objRS = m_objPermaCon.Execute(strQuery, , 1)
20:    Set lvwItems = lvwRegistered.ListItems
        
22:    Do Until m_objRS.EOF
            'Add listitens
24:         Set lvwItem = lvwItems.Add(, , CStr(m_objRS(0).Value)) 'UsrClass.UserName

26:         lvwItem.SubItems(1) = CStr(m_objRS(1).Value) 'UsrStatic.Pass
27:         lvwItem.SubItems(2) = CInt(m_objRS(2).Value) 'UsrClass.Class
28:         lvwItem.SubItems(3) = CStr(m_objRS(3).Value) 'ClassTypes.Name
29:         lvwItem.SubItems(4) = CStr(m_objRS(4).Value) 'UsrStatic.RegedBy
30:         lvwItem.SubItems(5) = CStr(m_objRS(5).Value) 'UsrStatic.RegDate
31:         If Not m_objRS(6).Value = Empty Then _
               lvwItem.SubItems(6) = CStr(m_objRS(6).Value) 'UsrDynamic.LastLogin
33:         If Not m_objRS(7).Value = Empty Then _
               lvwItem.SubItems(7) = CStr(m_objRS(7).Value) 'UsrDynamic.LastIP
                    
36:         m_objRS.MoveNext
37:    Loop
    
39:    lblDBRegCount.Caption = CInt(lvwRegistered.ListItems.Count)
    
41:    If Not m_objRS Is Nothing Then Set m_objRS = Nothing

43:    Exit Sub
44:
Err:
45:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.DBGetRegRecord()"
End Sub

Public Function DBConnectionOpen() As Boolean
1:    On Error GoTo Err
2:    Dim sConnString As String
    
      Select Case m_DbType
        Case MsAccess
4:            sConnString = "PROVIDER=MSDASQL;" & _
                            "DRIVER={Microsoft Access Driver (*.mdb)};" & _
                            "DBQ= " & G_APPPATH & "\DBs\userdb.mdb" & ";"
        Case MySQL
10:           sConnString = "Driver={Mysql ODBC 5.1 Driver};" & _
                            "Database=" & g_objSettings.DBName & ";" & _
                            "UID=" & g_objSettings.DBUserName & ";" & _
                            "PWD=" & g_objSettings.DBPassword & ";" & _
                            "PORT=" & g_objSettings.DBServerPort & ";" & _
                            "OemToAnsi=" & "No" & ";" & _
                            "SSL=" & "Yes" & ";" & _
                            "Server=" & g_objSettings.DBServerAddresse & ";" & _
                            "OPTION=" & "2084" & ";"
19:    End Select

       'Open database
22:    Set m_objPermaCon = New Connection
    
24:    m_objPermaCon.ConnectionTimeout = 10
25:    m_objPermaCon.CursorLocation = adUseClient
26:    m_objPermaCon.Mode = adModeReadWrite
27:    m_objPermaCon.Open sConnString

29:    DoEvents
        
      'Wait until it is open before continuing
32:    Do Until m_objPermaCon.State = adStateOpen
33:        DoEvents
34:    Loop
    
36:    If Not g_objRegistered Is Nothing Then
37:        Set g_objRegistered = Nothing
38:    End If

41:    Set g_objRegistered = New clsRegistered

43:    If DBConnectionTest Then
44:        DBConnectionOpen = True
45:    Else
46:        If Not m_DbType = MsAccess Then _
              Call DBSetDefaut
48:    End If
    
50:    Exit Function
51:
Err:
52:    If Not m_DbType = MsAccess Then
53:        MsgBox "Error when connecting to data base." & vbNewLine & _
                  "The data base interface was auto changed to MS Access." & vbNewLine & _
                  "Error Description:" & Err.Description, vbCritical
56:        Call DBSetDefaut
57:    Else
58:        HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.DBConnectionOpen()"
59:    End If
End Function
Public Sub DBConnectionClose()
1:    On Error GoTo Err
    
3:    If Not m_objPermaCon Is Nothing Then
4:        If Not m_objPermaCon.State = adStateClosed Then
5:            m_objPermaCon.Close
6:        End If
7:        Set m_objPermaCon = Nothing
8:    End If
    
10:   Exit Sub
11:
Err:
12:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.DBConnectionClose()"
End Sub

Public Function DBConnectionTest() As Boolean
1:    On Error GoTo Err
      'Check if exist all tables..
3:    m_objPermaCon.Execute "SELECT Count(*) FROM BanNames UNION " & _
                            "SELECT Count(*) FROM ClassTypes UNION " & _
                            "SELECT Count(*) FROM Messages UNION " & _
                            "SELECT Count(*) FROM UsrClass UNION " & _
                            "SELECT Count(*) FROM UsrDynamic UNION " & _
                            "SELECT Count(*) FROM UsrStatic;"
9:    DBConnectionTest = True
10:   Exit Function
11:
Err:
12:   MsgBox Err.Description, vbCritical
End Function

Public Function DBChangeType(ByVal iType As enInterfaceDB) As Boolean
1:    On Error GoTo Err
2:    m_DbType = iType
3:    g_objSettings.DBType = iType
      Select Case iType
        Case 0: cmbDataBase.Text = "MS Access"
        'Case 1: cmbDataBase.Text = "SQLite"
        Case 1: cmbDataBase.Text = "MySQL"
4:    End Select
5:    Exit Function
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.DBChangeType(" & iType & ")"
End Function

Public Property Get DBType() As enInterfaceDB
1:    DBType = m_DbType
End Property

Public Sub DBSetDefaut()
1:    On Error GoTo Err
2:    DBConnectionClose
3:    Call DBChangeType(MsAccess)
4:    Exit Sub
5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.DBSetDefaut()"
End Sub

Private Sub cmdDataBaseApply_Click()
1:    On Error GoTo Err

3:    DBConnectionClose
    
      Select Case cmbDataBase.ListIndex
        Case 0: Call DBChangeType(MsAccess)
        Case 1
5:          If LenB(txtData(54).Text) Then _
                 g_objSettings.DBName = txtData(54).Text _
            Else txtData(54).Text = g_objSettings.DBName
8:          Call DBChangeType(MySQL)
9:    End Select

11:   If DBConnectionOpen Then
12:        MsgBoxCenter Me, "Connection established with success.", vbInformation
13:   End If
    
15:   Exit Sub
16:
Err:
17:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdDataBaseApply_Click()"
End Sub
Private Sub cmdDataBaseHelp_Click()
1:    On Error GoTo Err
2:    MsgBoxCenter Me, _
        "PTDCH use MySQL Connector/ODBC 5.1 (Mysql ODBC 5.1 Driver) " & vbTwoLine & _
        "We chose to use the MySQL database server because" & vbNewLine & _
        "of its ease of installation, maintainability, configuration and speed." & vbNewLine & _
        "MySQL has also provided us with huge cost savings, " & vbNewLine & _
        "which we have been able to funnel into other resources." & vbTwoLine & _
        "For free download this driver goto:" & vbNewLine & _
        "http://dev.mysql.com/downloads/connector/odbc/5.1.html", vbInformation, "PTDCH - MySQL Interface"
9:   Exit Sub
10:
Err:
11:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdDataBaseHelp_Click()"
End Sub
Private Sub cmbDataBase_Click()
1:    On Error GoTo Err

3:    If cmbDataBase.ListIndex = 1 Then 'MySQL
4:        txtData(50).Enabled = True
5:        txtData(51).Enabled = True
6:        txtData(52).Enabled = True
7:        txtData(53).Enabled = True
8:        txtData(54).Enabled = True
9:    Else
10:       txtData(50).Enabled = False
11:       txtData(51).Enabled = False
12:       txtData(52).Enabled = False
13:       txtData(53).Enabled = False
14:       txtData(54).Enabled = False
15:   End If
    
17:   Exit Sub
18:
Err:
19:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbDataBase_Click()"
End Sub
'------------------------------------------------------------------------------
'End - Data Base Interace Subs/Functions
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'Form events
'------------------------------------------------------------------------------
Private Sub Form_Activate()
1:  If picTab(5).Visible Then SCI_Focus
End Sub
Private Sub Form_Initialize()
1:    On Error GoTo Err
     'Make sure people don't try to run multiple instances of the
     'same PTDCH exe; this will quite frankly frig everything up
     '(*** Note - Should not affect different PTDCH exes from running)
5:    If CBool(GetSetting("PTDCH", "Shell", "Parameter", "False")) = True Then
6:       DeleteSetting "PTDCH", "Shell", "Parameter"
7:    Else
8:       If App.PrevInstance Then
9:          MsgBox "There is already one instance of this particular PTDCH exe running.", vbExclamation, "PTDCH"
            End 'Hard end
11:      End If
12:   End If
      'Turn off nasty error messages which might lead to crashing (b/c of API calls)
14:   SetErrorMode &H1 Or &H2
15:   Exit Sub
16:
Err:
17:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.Form_Initialize()"
18:   Resume Next
End Sub
Private Sub Form_Load()
1:      On Error GoTo Err

3:      AddLog "Aplication started."
         
    #If SVN Then
6:     G_LOGPATH = G_APPPATH & "\Logs\MsgDebugLog.log"
    #End If

        'Open error handling file
10:     Open G_APPPATH & "\Logs\Error.log" For Append As G_ERRORFILE

        'Set comma status
13:     m_blnCommaDecimal = InStrB(1, CStr(0.1), ",")

        'Create our core global objects
16:     Set g_objFunctions = New clsFunctions
17:     Set g_colUsers = New clsHub
18:     Set g_colIPBans = New clsIPBans
19:     Set g_objSettings = New clsSettings
20:     Set g_colCommands = New clsCommands
21:     Set g_objRegExps = New clsRegExps
        'PLAN
23:     Set g_objScheduler = New clsPlan
        'Initialise PTDCH messages, language
25:     Set g_colMessages = New clsDictionary
        
        'USER LANGUAGE
28:     Set g_colLanguages = New Collection

        ' hook window for sizing control
        ' Disable the following line if you will be debugging form.
32:     Call HookWin(Me.hwnd, G_HbWnd)
        
    #If Status Then
35:     Set g_objStatus = New clsStatus
    #End If

        'Create local objects
39:     Set m_objDetectIP = New clsHTTPDownload
40:     Set m_colFailedReg = New Collection
41:     Set m_colConnectAttempts = New Collection
42:     Set m_colRevConnects = New Collection
43:     Set g_objChatRoom = New clsChatRoom

        'Set local vars
46:     m_lngBotsUB = -1

        'Add system tray icon
49:     Call SysTrayAdd
 
        'Load settings
52:     LoadDefaultSettings
53:     LoadSettings

55:     cmbDataBase.AddItem "MS Access"
56:     cmbDataBase.AddItem "MySQL"
         
        Select Case g_objSettings.DBType
             Case 0: Call DBChangeType(MsAccess)
             Case 1: Call DBChangeType(MySQL)
59:     End Select
         
61:     DBConnectionOpen
         
63:     If g_objSettings.MagneticWin Then _
            Call m_ObjMagnetic.AddWindow(frmHub.hwnd) 'Cool FX Windows
        
66:     tmrBackground.Interval = 60000 'Set background timer interval to every 20 mins

        'Prepare detect ip class
69:     m_objDetectIP.Host = "www.whatismyip.org"
70:     m_objDetectIP.Port = 80
    
        'Do extra actions
73:     If g_objSettings.AutoStart Then cmdButton_Click 1
    
        'Load DynamicIPServices for automatic IP updating
76:     lIntervalMin = IPCheckInterval 'to update services directly after start if neccessary
        
        'tmrUpdateIPs.Interval = 60 * 1000 'check every minute
80:     If g_objSettings.DynUpdate = True Then LoadDynIPs

        'UpDate No-IP DNS at Starting PTDCH
83:     If g_objSettings.NoIPUpdateStartUp Then
84:        m_datForceDNSUpdate = Now
85:        UpdateDNSs
86:     End If

87:     Set g_objHighlighter = New clsYHighlighter
88:     g_objHighlighter.LoadDirectory G_APPPATH & "\Settings"

90:     tlbScript.Buttons.Item(15).Value = tbrPressed
91:     tbsScripts.Visible = False

93:     Call IniDbExplorer

95:     Exit Sub

97:
Err:
98:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub_Load()"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
1:     On Error GoTo Err
       
3:     If Not G_GUI_IN_UNLOAD Then
           'Confirm exit if needed
5:         If g_objSettings.ConfirmExit And (Not UnloadMode = vbAppWindows) Then
6:            If MsgBoxCenter(Me, g_colMessages.Item("msgExitPTDCH"), vbYesNo Or vbQuestion Or vbDefaultButton2, g_colMessages.Item("msgConfirmExit")) = vbNo Then
7:               Cancel = 1
8:               Exit Sub
9:            End If
10:        End If
12:        Call APP_TERMINATE
13:    End If
14:    Exit Sub
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub_QueryUnload()"
End Sub
Private Sub Form_Terminate()
1:   On Error Resume Next
2:   Set frmHub = Nothing
End Sub
Public Sub Form_Resize()
1:   On Error GoTo Err
2:   Dim i As Integer

     Select Case Me.WindowState
       
           '***************************
           Case vbMinimized
           '***************************
          
8:            mnuTray(0).Enabled = True
9:            mnuTray(1).Enabled = False

             'Hide the form if selected
12:           If g_objSettings.MinimizeTray And Me.WindowState = vbMinimized Then _
                   Me.Hide: Exit Sub
           
           '***************************
           Case Else 'vbNormal Or vbMaximized
           '***************************
           
18:           mnuTray(0).Enabled = False
19:           mnuTray(1).Enabled = True
           
21:           With tbsMenu
22:              If Not Me.Width < 9390 Then
23:                      .Width = Me.Width - 240
24:              End If
25:              If Not Me.Height < 5445 Then
26:                      .Height = Me.Height - 850
27:              End If
28:           End With
           
30:           For i = 0 To picTab.Count - 1
31:              With picTab(i)
32:                   .Left = tbsMenu.Left + 80
33:                   .Width = tbsMenu.Width - 170
34:                   .Top = tbsMenu.Top + 360
35:                   .Height = tbsMenu.Height - 455
36:              End With
37:           Next

39:           picBordTab(0).Width = Me.Width
           
41:           With tbsSecurity
42:               If Not Me.Height < 5445 Then .Width = (Me.Width - 600)
43:               If Not Me.Height < 5445 Then .Height = (Me.Height - 1500)
44:           End With
           
46:           For i = 0 To picSTab.Count - 1
47:              With picSTab(i)
48:                   .Left = tbsSecurity.Left + 80
49:                   .Width = tbsSecurity.Width - 170
50:                   .Top = tbsSecurity.Top + 360
51:                   .Height = tbsSecurity.Height - 455
52:              End With
53:           Next

55:           picBordTab(2).Width = picTab(1).Width
           
57:           With tbsInteractions
58:               If Not Me.Height < 5445 Then .Width = (Me.Width - 600)
59:               If Not Me.Height < 5445 Then .Height = (Me.Height - 1500)
60:           End With
           
62:           For i = 0 To picITab.Count - 1
63:              With picITab(i)
64:                   .Left = tbsInteractions.Left + 80
65:                   .Width = tbsInteractions.Width - 170
66:                   .Top = tbsInteractions.Top + 360
67:                   .Height = tbsInteractions.Height - 455
68:              End With
69:           Next

71:           picBordTab(4).Width = picTab(2).Width
           
73:           With tabAdv
74:               If Not Me.Height < 5445 Then .Width = (Me.Width - 600)
75:               If Not Me.Height < 5445 Then .Height = (Me.Height - 1500)
76:           End With
           
78:           For i = 0 To picTabAdv.Count - 1
79:              With picTabAdv(i)
80:                   .Left = tabAdv.Left + 80
81:                   .Width = tabAdv.Width - 170
82:                   .Top = tabAdv.Top + 360
83:                   .Height = tabAdv.Height - 455
84:              End With
85:           Next

87:           picBordTab(3).Width = picTab(3).Width
              
89:           With tbsHelp
90:               If Not Me.Height < 5445 Then .Width = (Me.Width - 600)
91:               If Not Me.Height < 5445 Then .Height = (Me.Height - 1500)
92:           End With
              
94:           For i = 0 To picHelp.Count - 1
95:              With picHelp(i)
96:                   .Left = tbsHelp.Left + 80
97:                  .Width = tbsHelp.Width - 170
98:                  .Top = tbsHelp.Top + 360
99:                  .Height = tbsHelp.Height - 455
100:             End With
101:          Next

103:          picBordTab(1).Width = picTab(8).Width
              
105:          If lvwScripts.Visible Then
106:                 With tbsScripts
107:                      .Left = 60
108:                      .Width = (picTab(5).Width - lvwScripts.Width - 200)
109:                      .Top = 80 + tlbScript.Height
110:                      .Height = (picTab(5).Height - txtScriptError.Height - 230) - tlbScript.Height
111:                 End With
                  
113:                 With lvwScripts
114:                      .Left = (tbsScripts.Width + 120)
115:                      .Height = (picTab(5).Height - txtScriptError.Height - 230)
116:                      .Top = 60
117:                 End With
118:           Else
119:                 With tbsScripts
120:                      .Left = 60
121:                      .Width = (picTab(5).Width - 150)
122:                      .Top = 80 + tlbScript.Height
123:                      .Height = (picTab(5).Height - txtScriptError.Height - 230) - tlbScript.Height
124:                 End With
125:           End If

127:           If tbsScripts.Visible Then
128:               For i = 1 To picSciMain.Count - 1
129:                     With picSciMain(i)
130:                          .Left = tbsScripts.Left + 15
131:                          .Width = tbsScripts.Width - 60
132:                          .Top = tbsScripts.Top + 330
133:                          .Height = tbsScripts.Height - 375
134:                     End With
135:               Next
136:           Else
137:               For i = 1 To picSciMain.Count - 1
138:                     With picSciMain(i)
139:                          .Left = tbsScripts.Left
140:                          .Width = tbsScripts.Width
141:                          .Top = tbsScripts.Top
142:                          .Height = tbsScripts.Height
143:                     End With
144:               Next
145:           End If

147:           For i = 1 To UBound(g_objSciLexer)
148:                g_objSciLexer(i).SizeScintilla 0, 0, picSciMain(i).ScaleWidth / Screen.TwipsPerPixelX, (picSciMain(i).ScaleHeight / Screen.TwipsPerPixelY)
149:           Next
                
151:           With txtScriptError
152:                .Left = 60
153:                .Width = (picTab(5).Width - 150)
154:                .Top = (tbsScripts.Height) + 140 + tlbScript.Height
155:           End With

157:           With tlbScript
158:                .Left = 60
159:                .Width = tbsScripts.Width
160:           End With

162:           With txtNotePad
163:                .Left = 60
164:                .Top = 60
165:                .Height = (picHelp(2).Height - 140)
166:                    .Width = (picHelp(2).Width - 140)
167:           End With
            
169:           With txtChangeLog
170:                .Left = 60
171:                .Top = 60 + (cmbChangeLog.Height + 100)
172:                .Height = (picHelp(1).Height - 140) - (cmbChangeLog.Height + 100)
173:                .Width = (picHelp(2).Width - 140)
174:           End With

176:           With lvwRegistered
177:                .Width = (picSTab(0).Width - 180)
178:                .Height = (picSTab(0).Height - 600)
179:           End With
                
181:           With tbsDbManager
182:                .Top = 60
183:                .Left = 60
184:                .Width = (picSTab(0).Width - 140)
185:                .Height = (picSTab(0).Height - 140)
186:           End With
           
188:           With picSqlSCI
189:                .Left = tbsDbManager.Left + 60
190:                .Width = tbsDbManager.Width - 150
191:                .Top = 840
192:                .Height = tbsDbManager.Height - 840
193:           End With
                
195:           m_objSciExplorerSQL.SizeScintilla 0, 0, picSqlSCI.ScaleWidth / Screen.TwipsPerPixelX, (picSqlSCI.ScaleHeight / Screen.TwipsPerPixelY)
     
197:           With lvwSqlExplorer
198:                .Left = tbsDbManager.Left + 60
199:                .Width = tbsDbManager.Width - 150
200:                .Top = 440
201:                .Height = tbsDbManager.Height - 480
202:           End With
           
204:           With txtSqlErr
205:                .Left = 1680
206:                .Width = tbsDbManager.Width - 1750
207:           End With
           
209:           picBordTab(7).Width = picSTab(4).Width
                
211:           cmbRegistered.Left = (lvwRegistered.Width - cmbRegistered.Width + 100)
212:           lblDBRegCount.Left = (lvwRegistered.Width - cmbRegistered.Width - 150)
               
214:           With tabBansIPs
215:                .Left = Shape(7).Width + 200
216:                .Top = 60
217:                .Width = picSTab(2).Width - Shape(7).Width - 200
218:                .Height = picSTab(2).Height - 120
219:           End With

221:           With lvwTempIPBan
222:                .Left = tabBansIPs.Left + 80
223:                .Width = tabBansIPs.Width - 170
224:                .Top = tabBansIPs.Top + 360
225:                .Height = tabBansIPs.Height - 445
226:           End With
227:           With lvwPermIPBan
228:                .Left = tabBansIPs.Left + 80
229:                .Width = tabBansIPs.Width - 170
230:                .Top = tabBansIPs.Top + 360
231:                .Height = tabBansIPs.Height - 445
232:           End With
            
234:           picBordTab(8).Width = picSTab(2).Width
            
               'MOTD
237:           With txtData(6)
238:                .Top = 1680
239:                .Left = 2760
240:                .Width = picITab(0).Width - 3000
241:                .Height = picITab(0).Height - 1900
242:           End With

244:           Shape(13).Width = picITab(0).Width - 2750
245:           Shape(13).Height = picITab(0).Height - 200

#If Status Then
248:           With tbsStatus
249:                .Left = 60
250:                .Width = (picTab(6).Width - 150)
251:                .Top = 60
252:                .Height = (picTab(6).Height - 140)
253:           End With
           
255:           For i = 0 To picStatus.Count - 1
256:              With picStatus(i)
257:                   .Left = tbsStatus.Left + 80
258:                   .Width = tbsStatus.Width - 170
259:                   .Top = tbsStatus.Top + 360
260:                   .Height = tbsStatus.Height - 455
261:              End With
262:           Next
           
264:           picBordTab(6).Width = picTab(6).Width

266:           With lvwUsers
267:               .Left = 3840
268:               .Width = (picStatus(4).Width - 4000)
269:               .Top = 120
270:               .Height = (picStatus(4).Height - 280)
271:           End With

273:           With lstStatus(0)
274:                .Left = 60
275:                .Width = picStatus(0).Width - 160
276:                .Top = 60
277:                .Height = picStatus(0).Height - 700
278:           End With

280:           With lstStatus(1)
281:                .Left = 60
282:                .Width = picStatus(1).Width - 160
283:                .Top = 60
284:                .Height = picStatus(1).Height - 160
285:           End With

287:           With lstStatus(2)
288:                .Left = 60
289:                .Width = picStatus(2).Width - 160
290:                .Top = 60
291:                .Height = picStatus(2).Height - 160
292:           End With
           
294:           With txtLog
295:                .Left = 60
296:                .Width = picStatus(3).Width - 160
297:                .Top = 60
298:                .Height = picStatus(3).Height - 160
299:           End With
           
301:           Labels(9).Top = lstStatus(0).Height + 100
302:           lblHolder(26).Top = lstStatus(0).Height + 120
303:           cmdStSend.Top = lstStatus(0).Height + 120
           
305:           sldStatus.Top = lstStatus(0).Height + 120
306:           lblStatus(0).Top = lstStatus(0).Height + 420
           
308:           optStSend(0).Top = lstStatus(0).Height + 120
309:           optStSend(1).Top = lstStatus(0).Height + 350
310:           lblOptStSend(0).Top = lstStatus(0).Height + 140
311:           lblOptStSend(1).Top = lstStatus(0).Height + 370
           
313:           txtStForm.Top = lstStatus(0).Height + 300
314:           txtStSend.Top = lstStatus(0).Height + 300
           
316:           sldStatus.Left = lstStatus(0).Width - sldStatus.Width + 50
317:           lblStatus(0).Left = lstStatus(0).Width - sldStatus.Width + 70
           
319:           lblOptStSend(0).Left = lstStatus(0).Width - sldStatus.Width - 630
320:           lblOptStSend(1).Left = lstStatus(0).Width - sldStatus.Width - 630
           
322:           optStSend(0).Left = lblOptStSend(0).Left - 250
323:           optStSend(1).Left = lblOptStSend(1).Left - 250
           
325:           cmdStSend.Left = lblOptStSend(1).Left - cmdStSend.Width - 300
           
327:           txtStSend.Width = cmdStSend.Left - cmdStSend.Width - 500
#End If

330:           With tbsInfo
331:               .Left = 60
332:               .Width = (picTab(7).Width - 150)
333:               .Top = 60
334:               .Height = (picTab(7).Height - 140)
335:           End With
              
337:           For i = 0 To picInfo.Count - 1
338:                With picInfo(i)
339:                     .Left = tbsInfo.Left + 80
340:                     .Width = tbsInfo.Width - 170
341:                     .Top = tbsInfo.Top + 360
342:                     .Height = tbsInfo.Height - 455
343:                End With
344:           Next
              
346:           picBordTab(5).Width = picTab(7).Width
             
348:           With lvwPlugins
349:                .Left = 60
350:                .Width = picInfo(0).Width - 160
351:                .Top = 60
352:                .Height = picInfo(0).Height - 500
353:           End With

355:           For i = 0 To cmdPlugins.Count - 1
356:                cmdPlugins(i).Top = lvwPlugins.Height + 100
357:           Next
              
359:           chkData(68).Top = lvwPlugins.Height + 120
360:           chkData(68).Left = lvwPlugins.Width - 140
361:           lblCheck(68).Top = lvwPlugins.Height + 140
362:           lblCheck(68).Left = chkData(68).Left - lblCheck(68).Width - 60
              
364:           With lvwChatRom
365:                .Top = 60
366:                .Left = 60
367:                .Width = picITab(4).Width - 160
368:           End With
           
370:           With lvwPlan
371:                .Top = 2040
372:                .Left = 120
373:                .Width = picITab(2).Width - 260
374:                .Height = picITab(2).Height - lvwCommands.Height - 600
375:           End With
           
377:        End Select
   
379:  Exit Sub
380:
Err:
381:  HandleError Err.Number, Err.Description, Erl & "|frmHub.Form_Resize()"
382:  Resume Next
End Sub
'------------------------------------------------------------------------------
'End Form events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'SQL Explorer events
'------------------------------------------------------------------------------
Public Function SQLRemComments(ByRef strSqlCmd As String) As String
3:    Dim a As Integer, b As Integer, c As Integer
7:    Dim strTemp As String, strLine As String
9:    On Error GoTo Err

      If UBound(Split(strSqlCmd, vbCr)) = 0 Or _
         UBound(Split(strSqlCmd, vbCrLf)) = 0 Or _
         UBound(Split(strSqlCmd, vbLf)) = 0 Then _
            SQLRemComments = strSqlCmd: Exit Function

10:    b = 1
       strSqlCmd = vbNewLine & strSqlCmd & vbNewLine
       
       'Remove sql comments from string
13:    For a = 1 To Len(strSqlCmd)
14:        If Mid(strSqlCmd, a, 1) = Chr(10) Then
               'Check if is sql comment in this line
16:            strLine = Mid(strSqlCmd, b, (a - b))
17:            If strLine <> Chr(10) Then
18:                For c = 1 To Len(strLine)
19:                    If CStr(Mid(strLine, c, 2)) = "--" Then
20:                        Exit For
21:                    Else
22:                        If CStr(Mid(strLine, c, 1)) <> " " Then
23:                            strTemp = strTemp & Mid(strSqlCmd, b, (a - b))
24:                            Exit For
25:                        End If
26:                    End If
27:                Next
28:            End If
29:            b = a + 1
30:       End If
31:    Next a

       If strTemp = Empty Or strTemp = "" Then _
            SQLRemComments = Replace(strSqlCmd, Chr(10), " ") _
       Else SQLRemComments = Replace(strTemp, Chr(10), " ")

35:    Exit Function
36:
Err:
37:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SQLRemComments(" & strSqlCmd & ")"
End Function

Private Sub lblQueryDB_Click(Index As Integer)
1:    On Error GoTo Err
2:    Dim i As Integer, Y As Integer
3:    Dim strTemp           As String
4:    Dim strSqlCmd         As String
5:    Dim lvwItem           As ListItem
6:    Dim lvwItems          As ListItems

      Select Case Index

        'Run sql connection string
        '*******************************************************
        Case 0
        '*******************************************************
            'Check if slected text from SCI..
13:          strTemp = CStr(m_objSciExplorerSQL.GetSelText)
14:          If Not Len(strTemp) > 10 Then _
                strTemp = m_objSciExplorerSQL.Text

17:          strSqlCmd = SQLRemComments(strTemp)

19:          txtSqlErr.Text = ""
20:          lvwSqlExplorer.ListItems.Clear
21:          lvwSqlExplorer.ColumnHeaders.Clear

23:          On Error GoTo ErrSql

25:            Set m_objRS = m_objPermaCon.Execute(strSqlCmd)

27:            If Not m_objRS.EOF Then
                    'Add the colluns name
29:               For i = 0 To m_objRS.Fields.Count - 1
30:                   lvwSqlExplorer.ColumnHeaders.Add (i + 1), , m_objRS.Fields(i).Name
31:               Next
                  
33:               Set lvwItems = lvwSqlExplorer.ListItems

                  'Add the Rows
36:               For i = 0 To m_objRS.RecordCount - 1
37:                   Set lvwItem = lvwItems.Add((i + 1), , m_objRS(0).Value)
                      
39:                   For Y = 1 To m_objRS.Fields.Count - 1
40:                         If Not m_objRS(Y).Value = Empty Then
41:                             lvwItem.SubItems(Y) = m_objRS(Y).Value
42:                         Else
43:                             lvwItem.SubItems(Y) = "NULL"
44:                         End If
45:                   Next Y
                      
47:                   m_objRS.MoveNext
48:               Next i
49:          End If
         
51:          txtSqlErr.Text = "[" & Now & "] No syntax errors in SQL command."

53:          If Not m_objRS Is Nothing Then Set m_objRS = Nothing
             
        'Clear and add defaut sql string
        '*******************************************************
        Case 1
        '*******************************************************
        
59:          On Error GoTo Err
            
61:          Set m_objRS = m_objPermaCon.OpenSchema(adSchemaTables)

63:          strTemp = "-- Database Tables (userdb.mdb)" & vbNewLine
                  
65:          Do While Not m_objRS.EOF
66:              i = m_objRS.Fields.Count

68:              If UCase(Left(m_objRS.Fields("TABLE_NAME"), 4)) <> "MSYS" Then _
                    strTemp = strTemp & "--" & vbTab & m_objRS.Fields("TABLE_NAME") & vbNewLine

71:              m_objRS.MoveNext
72:          Loop
               
74:          strTemp = strTemp & vbNewLine & "-- Demo (Show all Registered Users) " & vbNewLine & _
                      "SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP " & vbNewLine & _
                      "FROM UsrStatic INNER JOIN (ClassTypes INNER JOIN (UsrClass INNER JOIN UsrDynamic ON UsrClass.UserName = UsrDynamic.UserName) ON ClassTypes.ID = UsrClass.Class) ON (UsrStatic.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrClass.UserName) " & vbNewLine & _
                      "-- Demo (Search User)" & vbNewLine & _
                      "-- WHERE UsrClass.UserName LIKE '%fLaSh%'" & vbNewLine & _
                      "ORDER BY UsrClass.UserName" & vbNewLine
                      
81:          m_objSciExplorerSQL.Text = strTemp
82:          m_objSciExplorerSQL.ClearUndoBuffer

84:         If Not m_objRS Is Nothing Then Set m_objRS = Nothing
            
86:   End Select
     
88:   m_objSciExplorerSQL.SetFocus
     
90:   Exit Sub
91:
ErrSql:
92:   txtSqlErr.Text = "[" & Now & "] Error: " & Err.Description
93:   Err.Clear
94:   m_objSciExplorerSQL.SetFocus
95:   Exit Sub
96:
Err:
97:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdSql_Click(" & Index & ")"
98:   m_objSciExplorerSQL.SetFocus
End Sub
Public Sub IniDbExplorer()

2:    On Error GoTo Err

      'Create code editor *****************************************************
5:    Set m_objSciExplorerSQL = New clsYScintilla
    
7:    m_objSciExplorerSQL.CreateScintilla picSqlSCI
8:    m_objSciExplorerSQL.SetFixedFont "Courier New", 10

      ' Give the scrollbar a nice long width to handle a long line which may
      ' occur.
12:   m_objSciExplorerSQL.ScrollWidth = 10000

      'This is absolutly an imperative line
15:   m_objSciExplorerSQL.Attach picSqlSCI

17:   m_objSciExplorerSQL.LineNumbers = True
18:   m_objSciExplorerSQL.AutoIndent = True

20:   m_objSciExplorerSQL.SetMarginWidth MarginLineNumbers, 50
   
22:   Call g_objHighlighter.SetHighlighterBasedOnExt(m_objSciExplorerSQL, "bdManager.sql")
      '************************************************************************

25:   Dim strTemp As String
        
34:   If g_objFileAccess.FileExists(G_APPPATH & "\Settings\bdManager.sql") Then
35:       strTemp = g_objFileAccess.ReadFile(G_APPPATH & "\Settings\bdManager.sql")
36:   End If
     
38:   If strTemp <> "" Then
39:       m_objSciExplorerSQL.Text = strTemp
40:       m_objSciExplorerSQL.ClearUndoBuffer
41:   Else
42:       Call lblQueryDB_Click(1)
43:   End If

     '************************************************************************

49:  Exit Sub
50:
Err:
52:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.IniDbExplorer()"
End Sub
Private Sub lblQueryDB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    lblQueryDB(Index).BackColor = &HC0C0C0
End Sub
Private Sub lblQueryDB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    lblQueryDB(Index).BackColor = &H8000000F
End Sub
'------------------------------------------------------------------------------
'End SQL Explorer events
'------------------------------------------------------------------------------

Private Sub SCI_Focus()
1:  Dim i As Integer
2:  For i = 1 To picSciMain.Count - 1
3:        If picSciMain(i).Visible Then Exit For
4:  Next
5:  g_objSciLexer(i).SetFocus
End Sub
Private Sub lvwBans_Click()
1:    On Error GoTo Err
2:    Dim lngSelected As Long

4:    lngSelected = IsListViewSelected(lvwBans)

6:    If lngSelected <> -1 Then
7:       cmdDB(5).Enabled = True
8:       cmdDB(7).Enabled = True
9:    Else
10:      cmdDB(5).Enabled = False
11:      cmdDB(7).Enabled = False
12:      lblHolder(50).Caption = ""
13:   End If

15:   Exit Sub
16:
Err:
17:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwBans_Click()"
End Sub

'------------------------------------------------------------------------------
'Chat Rom
'------------------------------------------------------------------------------
Private Sub cmdChatRom_Click(Index As Integer)
1:  On Error GoTo Err
2:  Dim frm As New frmChatRoom
   
    Select Case Index
        Case 0
6:           g_objChatRoom.ShowAddDialog Me
        Case 1, 2
7:           Dim lvwItem As ListItem
              '
9:           Set lvwItem = lvwChatRom.SelectedItem

11:           If lvwChatRom.ListItems.Count = 0 Then Exit Sub

13:           If lvwItem.Selected Then
14:               If Index = 2 Then
15:                  g_objChatRoom.ShowEditDialog _
                        lvwItem.Text, Me
17:               ElseIf Index = 1 Then
18:                  g_objChatRoom.RemChat _
                        lvwItem.Text
20:               End If
21:          End If
22:      End Select
    
24:   cmdChatRom(1).Enabled = False
25:   cmdChatRom(2).Enabled = False
    
27:   Exit Sub
28:
Err:
29:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwChatRom_Click()"
End Sub
Private Sub lvwChatRom_Click()
1:    On Error GoTo Err
2:    Dim lngSelected As Long

4:    lngSelected = IsListViewSelected(lvwChatRom)

6:    If lngSelected <> -1 Then
7:       cmdChatRom(1).Enabled = True
8:       cmdChatRom(2).Enabled = True
9:    Else
10:      cmdChatRom(1).Enabled = False
11:      cmdChatRom(2).Enabled = False
12:   End If

14:   Exit Sub
15:
Err:
16:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwChatRom_Click()"
End Sub
Private Sub lvwChatRom_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:  On Error GoTo Err
    
    'Sort listview by column clicked
4:   lvwChatRom.SortKey = ColumnHeader.Index - 1
5:   lvwChatRom.SortOrder = IIfLng(lvwChatRom.SortOrder, lvwAscending, lvwDescending)
6:   lvwChatRom.Sorted = True
    
8:   Exit Sub
9:
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwChatRom_ColumnClick()"
End Sub

Private Sub lvwPermIPBan_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:  On Error GoTo Err
    
    'Sort listview by column clicked
4:   lvwPermIPBan.SortKey = ColumnHeader.Index - 1
5:   lvwPermIPBan.SortOrder = IIfLng(lvwPermIPBan.SortOrder, lvwAscending, lvwDescending)
6:   lvwPermIPBan.Sorted = True
    
8:   Exit Sub
9:
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwPermIPBan_ColumnClick()"
End Sub

'------------------------------------------------------------------------------
'End Chat Rom
'------------------------------------------------------------------------------
Private Sub lvwRegistered_Click()
1:    On Error GoTo Err
2:    Dim lngSelected As Long

4:    lngSelected = IsListViewSelected(lvwRegistered)

6:    If lngSelected <> -1 Then
7:       cmdDB(4).Enabled = True
8:       cmdDB(6).Enabled = True
9:       cmdDB(8).Enabled = True
10:   Else
11:      cmdDB(4).Enabled = False
12:      cmdDB(6).Enabled = False
13:      cmdDB(8).Enabled = False
14:   End If

16:   Exit Sub
17:
Err:
18:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwRegistered_Click()"
End Sub

Private Sub lvwScripts_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'Check if we should reset or stop
2:    If Item.Checked Then _
           frmScript.SReset Item.Index, False, True _
      Else frmScript.SStop Item.Index
End Sub

Private Sub lvwScripts_DblClick()
1:    On Error GoTo Err
2:    Dim i As Integer
3:    For i = 1 To lvwScripts.ListItems.Count
4:        If lvwScripts.ListItems(i).Selected Then
5:            tbsScripts.Tabs(i).Selected = True
6:            Exit Sub
7:        End If
8:    Next
9:    Exit Sub
10:
Err:
11:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwScripts_DblClick()"
End Sub

Private Sub lvwScripts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'Popup menu if left button is pressed
2:    If Button = 2 Then PopupMenu mnuPopUp(9)
End Sub

Private Sub lvwTempIPBan_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:  On Error GoTo Err
    
    'Sort listview by column clicked
4:   lvwTempIPBan.SortKey = ColumnHeader.Index - 1
5:   lvwTempIPBan.SortOrder = IIfLng(lvwTempIPBan.SortOrder, lvwAscending, lvwDescending)
6:   lvwTempIPBan.Sorted = True
    
8:   Exit Sub
9:
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwTempIPBan_ColumnClick()"
End Sub

Private Sub mnuCodeRTB1_Click(Index As Integer)
1:    On Error GoTo Err
2:    Dim i As Integer
   
4:    For i = 1 To picSciMain.Count - 1
5:        If picSciMain(i).Visible Then Exit For
6:    Next
    
    Select Case Index
        Case 0 'Insert Date/Time
8:            g_objSciLexer(i).SelText = Now
        Case 1 'Script Info
9:              Dim strInfo, strSize, strName, strString As String

11:              strString = g_objSciLexer(i).Text
12:              strName = frmHub.tbsScripts.Tabs(i).Key
                 
14:              strInfo = "Script strInfo - " & strName & vbCrLf & vbCrLf & _
                           "Characters: " & Len(strString) & vbCrLf & _
                           "Lines: " & CharCount(strString, vbCrLf) + 1 & vbCrLf & _
                           "Words: " & CharCount(strString, " ") + 1

19:              strSize = Len(strString)

                  'Calculate Size
22:              If strSize > 1000 Then
23:                 strSize = FormatNumber(strSize / 1024, 2)

25:                 MsgBoxCenter Me, strInfo & vbCrLf & vbCrLf & _
                            "The size of the file is:" & vbCrLf & _
                            FormatNumber(strSize, 2) & " Kb.", vbOKOnly + vbInformation
28:              Else
29:                 MsgBoxCenter Me, strInfo & vbCrLf & vbCrLf & _
                            "The size of the file is:" & vbCrLf & _
                            FormatNumber(strSize, 2) & " bytes.", vbOKOnly + vbInformation
32:              End If

        Case 4 'Save as..
34:         Dim cD As New clsCommonDialog
35:         Dim sFile As String

37:         sFile = frmHub.tbsScripts.Tabs(i).Key

39:         If (cD.VBGetSaveFileName(sFile, _
                  Filter:="VBScript (*.script)|.script|VBScript (*.vbs)|.script|All Files (*.*)|*.*", _
                      DefaultExt:="htm", _
                         Owner:=Me.hwnd)) Then
                '
44:             g_objFileAccess.WriteFile sFile, g_objSciLexer(i).Text
45:         End If
        Case 6 'Clear Undo Buffer
46:         g_objSciLexer(i).ClearUndoBuffer
47:         lvwScripts.ListItems(i).SubItems(3) = CStr(g_objSciLexer(i).Modified)
48:      End Select

50:   SCI_Focus

52:   Exit Sub
53:
Err:
54:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuCodeRTB1_Click(" & Index & ")"
End Sub

Private Sub mnuCodeRTB3_Click(Index As Integer)
1:    On Error GoTo Err
      Select Case Index
          Case 0 'VBScript documentation
2:              g_objFunctions.ShellExec "http://msdn2.microsoft.com/en-us/library/t0aew7h6.aspx"
          Case 1 'JScript documentation
3:              g_objFunctions.ShellExec "http://msdn2.microsoft.com/en-us/library/hbxc2t98(vs.71).aspx"
4:      End Select
    
6:    Exit Sub
7:
Err:
8:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuCodeRTB3_Click(" & Index & ")"
End Sub

Private Sub tabBansIPs_Click()
1:    On Error GoTo Err
 
      Select Case tabBansIPs.SelectedItem.Index
        Case 1
3:            lvwTempIPBan.Visible = True
4:            lvwPermIPBan.Visible = False
        Case 2
5:            lvwTempIPBan.Visible = False
6:            lvwPermIPBan.Visible = True
7:    End Select
    
9:    Exit Sub
10:
Err:
11:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tabBansIPs_Click()"
12:   Resume Next
End Sub

Private Sub tbsDbManager_Click()
2:   On Error GoTo Err
 
     Select Case tbsDbManager.SelectedItem.Index
        Case 1
4:          lvwSqlExplorer.Visible = False
5:          pciDBExplorer(0).Visible = True
6:          pciDBExplorer(1).Visible = True
7:          txtSqlErr.Visible = True
8:          picSqlSCI.Visible = True
9:          m_objSciExplorerSQL.SetFocus
        Case 2
10:         lvwSqlExplorer.Visible = True
11:         pciDBExplorer(0).Visible = False
12:         pciDBExplorer(1).Visible = False
13:         txtSqlErr.Visible = False
14:         picSqlSCI.Visible = False
15:    End Select
    
17:    Exit Sub
18:
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsDbManager_Click()"
End Sub

Private Sub tbsScripts_Click()

2:   On Error GoTo Err
 
4:   Dim i, i2 As Integer
   
6:   i2 = Val(tbsScripts.SelectedItem.Index)
7:   If picSciMain(i2).Visible = True Then Exit Sub
    
9:   For i = 1 To picSciMain.Count - 1
10:     picSciMain(i).Visible = False
11:  Next i
   
13:  i = Val(tbsScripts.SelectedItem.Index)
14:  picSciMain(i).Visible = True

16:  If frmEditScintilla.Visible Then frmEditScintilla.Visible = False

18:  Exit Sub
19:
Err:
20:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsScripts_Click()"
21:     Resume Next
End Sub

Private Sub tlbScript_ButtonClick(ByVal Button As ComctlLib.Button)
1:      Dim i As Integer
2:      Dim intIndex As Integer

4:      On Error GoTo Err

6:      For i = 1 To picSciMain.Count - 1
7:            If picSciMain(i).Visible Then intIndex = i: Exit For
8:      Next
        
     Select Case CStr(Button.Key)
            Case "Undo"
10:              g_objSciLexer(intIndex).Undo
            Case "Redo"
11:              g_objSciLexer(intIndex).Redo
            Case "Find"
12:              g_objSciLexer(intIndex).EditSciText 1, Me
            Case "Replace"
13:             g_objSciLexer(intIndex).EditSciText 2, Me
            Case "GoToLine"
14:             g_objSciLexer(intIndex).EditSciText 3, Me
            Case "Save Only"
15:             Call frmScript.SSave(i)
            Case "Save and Reset Script"
16:             Call frmScript.SReset(i, True, True)
            Case "Clear"
17:             g_objSciLexer(intIndex).Text = ""
            Case "Hide Scripts"
18:             If Button.Value Then _
                     lvwScripts.Visible = False _
                Else lvwScripts.Visible = True
21:             Call Form_Resize
            Case "Hide TabControl"
22:             If Button.Value Then _
                     tbsScripts.Visible = False _
                Else tbsScripts.Visible = True
23:             Call Form_Resize
            Case "Show Debug Windows"
24:             Dim Modal As Byte
25:             frmDebugSC.Show Modal, Me
            Case "Enabled Tabs"
26:             If Button.Value Then
                    'Remove all tabs..
28:                  ReDim m_TabsStop(0 To Controls.Count - 1) As Boolean
29:                  For i = 0 To Controls.Count - 1
30:                     On Error Resume Next
31:                     m_TabsStop(i) = Controls(i).TabStop
32:                     Controls(i).TabStop = False
33:                  Next
34:                  On Error GoTo Err
35:             Else
                     'Add All Tabs..
37:                  For i = 0 To Controls.Count - 1
38:                     On Error Resume Next
39:                     Controls(i).TabStop = m_TabsStop(i)
40:                  Next
41:                  On Error GoTo Err
42:             End If
            Case "New"
43:             frmNewScript.Show vbModal, Me
            Case "Menu"
44:             PopupMenu frmHub.mnuPopUp(10), 0, 5700, 800
45:        End Select

         Select Case Button.Key
            Case "Find", "Replace", "GoToLine", "New"
            Case Else: SCI_Focus
47:      End Select

49:   Exit Sub
50:
Err:
51:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tlbScript_ButtonClick(" & Button.Key & ")"
End Sub

Private Sub cmbSkin_Click()
   
2:      On Error GoTo Err
   
4:    cmdSkin(0).Enabled = True
5:    cmdSkin(1).Enabled = True
   
      Select Case cmbSkin.Text
         Case "01-Defaut"
7:           g_objSettings.lngSkin = 1
8:          cmdSkin(0).Enabled = False
         Case "02-Cyan Blue"
9:          g_objSettings.lngSkin = 2
         Case "03-Cyan Green"
10:          g_objSettings.lngSkin = 3
         Case "04-Metallic"
11:          g_objSettings.lngSkin = 4
         Case "05-Metallic Blue"
12:          g_objSettings.lngSkin = 5
         Case "06-Metallic Green"
13:          g_objSettings.lngSkin = 6
         Case "07-Metallic Navy Blue"
14:          g_objSettings.lngSkin = 7
         Case "08-Metallic Oliver"
15:          g_objSettings.lngSkin = 8
         Case "09-Texture Grain"
16:          g_objSettings.lngSkin = 9
         Case "10-Texture Spater"
17:          g_objSettings.lngSkin = 10
         Case "11-Texture Tiles"
18:          g_objSettings.lngSkin = 11
         Case "12-Texture Toxedo"
19:          g_objSettings.lngSkin = 12
         Case "13-Blue Berry"
20:          g_objSettings.lngSkin = 13
         Case "14-Glace Table"
21:          g_objSettings.lngSkin = 14
         Case "15-Pink"
22:          g_objSettings.lngSkin = 15
         Case "16-Gun Blue"
23:          g_objSettings.lngSkin = 16
         Case "17-Gun Metal"
24:          g_objSettings.lngSkin = 17
25:          cmdSkin(1).Enabled = False
26:      End Select
   
28:   Dim i As Integer
29:   On Error Resume Next
      'Refresh all picture box .. very fast
31:   For i = 0 To picTab.Count - 1: picTab(i).Refresh: Next i
32:   For i = 0 To picSTab.Count - 1: picSTab(i).Refresh: Next i
33:   For i = 0 To picITab.Count - 1: picITab(i).Refresh: Next i
34:   For i = 0 To picTabAdv.Count - 1: picTabAdv(i).Refresh: Next i
35:   For i = 0 To picHelp.Count - 1: picHelp(i).Refresh: Next i
36:   For i = 0 To picBordTab.Count - 1: picBordTab(i).Refresh: Next i
37:   For i = 0 To picInfo.Count - 1: picInfo(i).Refresh: Next i
38:   For i = 0 To picStatus.Count - 1: picStatus(i).Refresh: Next i

40:   Call Form_Paint
41:   Me.Refresh
   
43:   Exit Sub

45:
Err:
46:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbSkin_Click()"
End Sub

Private Sub cmdButton_Click(Index As Integer)
1:    Dim strTemp(3) As String
2:    Dim frm As New frmMulti

4:    On Error GoTo Err

      Select Case Index
        Case 0 'Check for updates
6:             frmUpDate.Show vbModal, Me
        Case 1 'Start / stop serving
7:            SwitchServing
        Case 2 'Redirect users
8:            NextRedirect
9:                With frm
10:                    .Label1.Caption = g_colMessages.Item("msgEnterRedirUsersAddress")
11:                    .Caption = g_colMessages.Item("msgRedirUsers")
12:                    .txtStr.Text = g_objSettings.RedirectIP
13:                    .Show vbModal, Me
14:                    strTemp(0) = .txtStr.Text
15:                End With
16:                Set frm = Nothing
               'Check and see if they pressed cancel
18:            If LenB(strTemp(0)) Then
                Select Case MsgBoxCenter(Me, g_colMessages.Item("msgRedirAll"), vbYesNo Or vbQuestion, g_colMessages.Item("msgRedirUsers"))
                    Case vbYes
19:                        g_colUsers.RedirectAll strTemp(0)
                    Case vbNo
20:                        g_colUsers.RedirectNonOps strTemp(0)
                    Case Else
21:                        Exit Sub
22:                End Select
                'Raise script event
24:                SEvent_StartedRedirecting
25:            End If
        Case 3 'Save settings
26:            SaveSettings
        Case 4 'Reload settings
27:            LoadDefaultSettings
28:            LoadSettings
        Case 5 'Detect IP
29:            strTemp(0) = DetectHubIP
30:            With frm
31:                .Label1.Caption = g_colMessages.Item("msgDetectIP")
32:                .Caption = "IP"
33:                .cmdCancel.Visible = False
34:                If Not strTemp(0) = vbNullString Then
35:                     .txtStr.Text = DetectHubIP
36:                Else 'change message to "try again" ?
37:                     .txtStr.Text = g_colMessages.Item("msgGettingIP")
38:                End If
39:                .Show vbModal, Me
40:            End With
41:            Set frm = Nothing
        Case 6 'Force UpDate
42:            m_datForceDNSUpdate = Now
43:            UpdateDNSs
        Case 7, 8, 9 ' Mass Messages..
        
45:            strTemp(0) = g_colMessages.Item("msgEnterPM")

47:            If Index = 7 Then 'Mass Messages to All
48:                 strTemp(1) = g_colMessages.Item("msgMassMsg")
49:                 strTemp(2) = g_objSettings.MassMessage
50:            ElseIf Index = 8 Then 'Mass Messages to Ops
51:                 strTemp(1) = g_colMessages.Item("msgMassMsgOp")
52:                 strTemp(2) = g_objSettings.OpMassMessage
53:            ElseIf Index = 9 Then 'Mass Messages to UnReg
54:                 strTemp(1) = g_colMessages.Item("msgMassMsgUnReg")
55:                 strTemp(2) = g_objSettings.UnRegMassMessage
56:            End If
57:            With frm
58:                 .Label1.Caption = strTemp(0)
59:                 .Caption = strTemp(1)
60:                 .Height = 2280
61:                 .txtStr.Visible = False
62:                 .txtStrMultiLine.Visible = True
63:                 .cmdCancel.Top = 1440
64:                 .cmdOK.Top = 1440
65:                 .txtStrMultiLine = strTemp(2)
66:                 .Show vbModal, Me
67:                 strTemp(3) = .txtStrMultiLine.Text
68:            End With
69:            Set frm = Nothing

71:            If Not LenB(strTemp(3)) Then Exit Sub

73:            If Index = 7 Then 'Mass Messages to All
74:                 g_objSettings.MassMessage = strTemp(3)
75:                 g_colUsers.SendPrivateToAll g_objSettings.BotName, strTemp(3)
76:                 AddLog "Mass Messages To All: " & strTemp(3)
77:            ElseIf Index = 8 Then 'Mass Messages to Ops
78:                 g_objSettings.OpMassMessage = strTemp(3)
79:                 g_colUsers.SendPrivateToOps g_objSettings.BotName, strTemp(3)
80:                 AddLog "Mass Messages To Ops: " & strTemp(3)
81:            ElseIf Index = 9 Then 'Mass Messages to UnReg
82:                 g_objSettings.UnRegMassMessage = strTemp(3)
83:                 g_colUsers.SendPrivateToUnReg g_objSettings.BotName, strTemp(3)
84:                 AddLog "Mass Messages To UnReg: " & strTemp(3)
85:            End If

87:            SEvent_MassMessage strTemp(3)

89:    End Select

91:    Exit Sub
    
93:
Err:
94:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdButton_Click(" & Index & ")"
End Sub
Private Sub cmdConvDatabase_Click()
1:   frmCAccounts.Show vbModal, Me
End Sub

Private Sub cmdDB_Click(Index As Integer)
1:  Dim strClass    As String
2:  Dim lvwItem     As ListItem
3:  Dim lvwItems    As ListItems
4:  On Error GoTo Err
    
6:  If Index = 4 Or Index = 6 Or Index = 8 Then If lvwRegistered.ListItems.Count = 0 Then Exit Sub
7:  If Index = 5 Or Index = 7 Then If lvwBans.ListItems.Count = 0 Then Exit Sub
    
9:    If Index = 4 Or Index = 6 Or Index = 8 Then '---------------------------------
10:       Set lvwItem = lvwRegistered.SelectedItem
11:       Set lvwItems = lvwRegistered.ListItems
       'Check if selected
13:       If lvwItem.Selected Then
            Select Case CLng(lvwItems(lvwItem.Index).SubItems(2))
                    'Case -1: strClass = "-1 = Locked"
                    'Case 0: strClass = "0 = Unknown"
                    'Case 1: strClass = "1 = Regular"
                    'Case 2: strClass = "2 = Mentored"
                    Case 3: strClass = "3 = Registered"
                    Case 4: strClass = "4 = Invisible"
                    Case 5: strClass = "5 = VIP"
                    Case 6: strClass = "6 = Operator"
                    Case 7: strClass = "7 = Invisible Operator"
                    Case 8: strClass = "8 = Super Operator"
                    Case 9: strClass = "9 = Invisible Super Operator"
                    Case 10: strClass = "10 = Admin"
                    Case 11: strClass = "11 = Invisible Admin"
                    Case Else: strClass = "2 = Mentored"
18:                End Select
19:       End If
20:    End If '---------------------------------------------------------
    
    Select Case Index '----------------------------------------------
            Case 0:  Call DBGetRegRecord 'Refresh Reg
            Case 1:  Call DBGetBanRecord 'Refresh Ban
            Case 2 'Add Reg ---------------------------------------------
22:             With frmReg
23:                  Load frmReg
24:                 .Tag = "Add"
25:                 .cmbClass = "3 = Registered"
26:                 .InicializeReg 'Perpare Form
27:                 .Show vbModal, Me
28:                 Pause (500): 'DBGetRegRecord 'Refresh Reg
29:             End With
            Case 3 'Add Ban ---------------------------------------------
30:             With frmBanName
31:                 .Tag = "Add"
32:                 .InicializeBan
33:                 .Show vbModal, Me
34:                 Pause (1000): 'DBGetBanRecord 'Refresh Ban
35:             End With
            Case 4 'Rem Reg ---------------------------------------------
36:             Set lvwItem = lvwRegistered.SelectedItem
            'Check if selected
38:             If lvwItem.Selected Then _
                    g_objRegistered.Remove lvwItem.Text: _
                      lvwRegistered.ListItems.Remove CInt(lvwItem.Index)
            Case 5 'Rem Ban ---------------------------------------------
41:             Set lvwItem = lvwBans.SelectedItem
            'Check if selected
43:             If lvwItem.Selected Then _
                    g_objRegistered.Remove lvwItem.Text: _
                     lvwBans.ListItems.Remove CInt(lvwItem.Index)
            Case 6 ' Edit Reg -------------------------------------------
            'Check if selected
47:             If lvwItem.Selected Then
48:                With frmReg
49:                    Load frmReg
50:                    .Tag = "Edit"
51:                    .txtPass.Text = lvwItems(lvwItem.Index).SubItems(1)  'Pass
52:                   .txtName.Text = lvwItem.Text
53:                   .cmbClass = strClass
54:                   .InicializeReg
55:                   .Show vbModal, Me
56:                End With
57:             End If
            Case 7 'Rename Ban ------------------------------------------
58:             Set lvwItem = lvwBans.SelectedItem
            'Check if selected
60:             If lvwItem.Selected Then
61:                With frmBanName
62:                    .Tag = "Rename"
63:                    .txtName.Text = lvwItem.Text
64:                    .txtName.Tag = lvwItem.Text
65:                    .txtReason.Text = lblHolder(50).Caption 'Reason
66:                    .InicializeBan
67:                    .Show vbModal, Me
68:                End With
69:             End If
            Case 8 'Rename Reg ------------------------------------------
            'Check if selected
71:               If lvwItem.Selected Then
72:                  With frmReg
73:                       Load frmReg
74:                      .Tag = "Rename"
75:                      .txtPass.Text = lvwItems(lvwItem.Index).SubItems(1)  'Pass
76:                      .txtName.Text = lvwItem.Text
77:                      .txtName.Tag = lvwItem.Text 'Used in rename.. firts name
78:                      .cmbClass = strClass
79:                     .InicializeReg
80:                     .Show vbModal, Me
81:                End With
82:             End If
83:        End Select '-----------------------------------------------------

85: Exit Sub
    
87:
Err:
88:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdDB_Click(" & Index & ")"
End Sub


Private Sub cmdPopup_Click(Index As Integer)
1:   g_objFunctions.ShowBallon "PT Direct Connect Hub " & vbVersion, "Created by fLaSh", Index, True
End Sub

Private Sub cmdSkin_Click(Index As Integer)
1:   On Error Resume Next
     Select Case Index
        Case 0
2:         cmbSkin.Text = cmbSkin.List(cmbSkin.ListIndex - 1)
        Case 1
3:         cmbSkin.Text = cmbSkin.List(cmbSkin.ListIndex + 1)
4:     End Select
End Sub

Sub CreateDynIPsXML()
    Dim strTemp As String
    Dim intFF As Integer
    
    On Error GoTo Err
    
    strTemp = G_APPPATH & "\Settings\DynIPs.xml"

    intFF = FreeFile

       'Append to file
    Open strTemp For Output As intFF

    Print #intFF, "<DynIPs>"
    Print #intFF, vbTab & "<!-- Service,Host,User,Pass -->"
    Print #intFF, vbTab & "<!-- if file does not exist, updating is disabled -->"
    Print #intFF, vbTab & "<0></0>"
    Print #intFF, vbTab & "<1></1>"
    Print #intFF, vbTab & "<2></2>"
    Print #intFF, vbTab & "<3></3>"
    Print #intFF, vbTab & "<4></4>"
    Print #intFF, vbTab & "<5></5>"
    Print #intFF, vbTab & "<6></6>"
    Print #intFF, vbTab & "<7></7>"
    Print #intFF, vbTab & "<8></8>"
    Print #intFF, vbTab & "<9></9>"
    Print #intFF, vbTab & "<!-- More than 10 services will be ignored -->"
    Print #intFF, "</DynIPs>";

    Close intFF

    Exit Sub
Err:
    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.CreateDynIPsXML()"
End Sub

Private Sub LoadDynIPs()
' should go in loadsettings sub
    
3:    Dim objXML          As clsXMLParser
4:    Dim objNode         As clsXMLNode
5:    Dim colNodes        As Collection
6:    Dim colAttributes   As Collection
    
8:    Dim strTemp     As String
9:    Dim strValues() As String

11:    On Error Resume Next
    
13:    strTemp = G_APPPATH & "\Settings\DynIPs.xml"
        
15:    If Not (g_objFileAccess.FileExists(strTemp)) Then
16:            CreateDynIPsXML
17:            Exit Sub
18:        End If

20:    Set objXML = New clsXMLParser

22:    objXML.Data = g_objFileAccess.ReadFile(strTemp)
23:    objXML.Parse

25:    Set colNodes = objXML.Nodes(1).Nodes

'    'Just in case...
28:    On Error Resume Next

30:    For Each objNode In colNodes
31:        strTemp = objNode.Value
32:        If (strTemp <> "") And (Val(objNode.Name) < 10) Then
33:            strValues = Split(strTemp, ",")
34:            If UBound(strValues) = 3 Then
35:                Service(objNode.Name) = strValues(0)
36:                Host(objNode.Name) = strValues(1)
37:                User(objNode.Name) = strValues(2)
38:                Pass(objNode.Name) = strValues(3)
'                tmrUpdateIPs.Enabled = True
40:            End If
41:        End If
42:    Next

44:    strTemp = CStr(UBound(Service))
45:    objXML.Clear
46:    Set objNode = Nothing
47:    Set colNodes = Nothing
End Sub

Private Sub LabelsURL_Click(Index As Integer)
1:   On Error Resume Next
     Select Case Index
        Case 0 'Send e-mail
2:         On Error Resume Next
3:         ShellExecute Me.hwnd, "open", "mailto:carlosferreiracarlos@hotmail.com?subject=About the PT DC Hub V." & vbVersion & "...&body=I have tested the software and...", 0&, 0&, vbNormal
        Case 1 'GoTo HomePage
4:         On Error Resume Next
5:         ShellExecute Me.hwnd, "open", "http://HublistChecker.pt.vu/", "", "", 3   'SW_SHOWMAXIMIZED
6:   End Select
End Sub

Private Sub LabelsURL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'1: LabelsURL(Index).ForeColor = &HFFC0C0
End Sub

Private Sub lblHolder_Change(Index As Integer)
1:    On Error GoTo Err
2:    Dim strTmp(1) As String

      'Connected Users
5:   If Index = 55 Then
6:        If Len(g_objSettings.HubName) > 22 Then _
               strTmp(0) = Left(g_objSettings.HubName, 20) & ".." _
          Else strTmp(0) = g_objSettings.HubName
          
10:       If G_SERVING Then
11:             strTmp(1) = "PT DC Hub " & vbVersion & vbNewLine & vbNewLine & _
                                strTmp(0) & vbNewLine & _
                              lblHolder(45).Caption & lblHolder(55).Caption
14:       Else
15:             strTmp(1) = "PT DC Hub " & vbVersion & vbNewLine & vbNewLine & _
                            strTmp(0)
17:       End If
          
19:       stbMain.Panels(4).Text = lblHolder(55).Caption
20:       SysTrayUpDate strTmp(1)
21:   End If
    
      'Connected Op
24:   If Index = 56 Then _
            stbMain.Panels(6).Text = lblHolder(56).Caption
    
      'Shared Total
28:   If Index = 57 Then _
            stbMain.Panels(5).Text = lblHolder(57).Caption
       
31:   Exit Sub
32:
Err:
33:   HandleError Err.Number, Err.Description, Erl & "|frmHub.lblHolder_Change(" & Index & ")"
End Sub

Private Sub lvwBans_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:  On Error GoTo Err
    
    'Sort listview by column clicked
4:   lvwBans.SortKey = ColumnHeader.Index - 1
5:   lvwBans.SortOrder = IIfLng(lvwBans.SortOrder, lvwAscending, lvwDescending)
6:   lvwBans.Sorted = True
    
8:   Exit Sub
9:
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwBans_ColumnClick()"
End Sub

Private Sub lvwBans_ItemClick(ByVal Item As ComctlLib.ListItem)
1:   lblHolder(50).Caption = CStr(Item.Tag)
End Sub

Private Sub lvwBans_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(4)
End Sub

Private Sub lvwCommands_DblClick()
1:    Dim lvwItem As ListItem
2:    Dim strKey  As String
    
4:    On Error GoTo Err
    
6:    Set lvwItem = lvwCommands.SelectedItem
    
    'Make sure an item is selected
9:    If ObjPtr(lvwItem) Then
10:        Load frmCommand
           'Update GUI
13:        strKey = lvwItem.Text
14:        frmCommand.txtTrigger.Text = strKey
15:        frmCommand.Tag = lvwItem.Text
16:        frmCommand.cmbClass.Text = lvwItem.SubItems(1)
17:        frmCommand.txtDescription.Text = g_colCommands(strKey).Description
18:        frmCommand.chkEnabled.Value = Abs(CBool(lvwItem.SubItems(2)))
    
20:        frmCommand.Show vbModal, Me
21:    End If

25:    Exit Sub
    
27:
Err:
28:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwCommands_DblClick()"
End Sub

Private Sub lvwPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    Dim lngSelected As Long

4:    lngSelected = IsListViewSelected(lvwPlan)

5:    If lngSelected <> -1 Then _
           mnuPlan(1).Enabled = True: mnuPlan(3).Enabled = True _
      Else mnuPlan(1).Enabled = False: mnuPlan(3).Enabled = False

9:    If Button = 2 Then PopupMenu mnuPopUp(8)

11:   Exit Sub
12:
Err:
14:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwPlan_MouseDown()"
End Sub

'------------------------------------------------------------------------------
'Start Plugins Subs
'------------------------------------------------------------------------------
Public Function GetVariable(ByVal strExpr As String) As Variant
    '------------------------------------------------------------------
    'Purpose:   This Sub set Variable frmHub/modules
    '
    'Params:    strVariable: variable name
    '
    'Returns:   none
    '
    '   Called by the plugins functions or scripts
    '------------------------------------------------------------------
10:  On Error GoTo Err

     Select Case LCase$(strExpr)
        'Consts
        Case "vbsvnversion": GetVariable = CStr(vbSVNVersion)
        Case "vbversion": GetVariable = CStr(vbVersion)
        Case "vblock": GetVariable = CStr(vbLock)
        Case "vbwelcome": GetVariable = CStr(vbWelcome)
        'Strings publics
        Case "g_apppath": GetVariable = CStr(G_APPPATH)
        Case "g_errorfile": GetVariable = CStr(G_ERRORFILE)
        Case "G_GUI_IN_UNLOAD": GetVariable = CBool(G_GUI_IN_UNLOAD)
        'Vars.. module
        Case "m_blncommadecimal": GetVariable = CBool(m_blnCommaDecimal)
        Case "m_lngscripteventsub": GetVariable = CLng(m_lngScriptEventsUB)
        Case "m_lngredirectub": GetVariable = CLng(m_lngRedirectUB)
        Case "m_lngbotsub": GetVariable = CLng(m_lngBotsUB)
        Case "m_lngbanfilter": GetVariable = CDate(m_lngBanFilter)
        Case "m_datservingdate": GetVariable = CDate(m_datServingDate)
15:   End Select

17:   Exit Function
18:
Err:
19:   On Error Resume Next
20:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SetVariable(" & strExpr & ")"
End Function
Public Sub SetVariable(ByRef strVariable As String, ByRef varValue As Variant)
    '------------------------------------------------------------------
    'Purpose:   This Sub set Variable frmHub/modules
    '
    'Params:    strVariable: variable name
    '           varValue: value
    '
    'Returns:   none
    '
    '   Called by the plugins functions or scripts
    '------------------------------------------------------------------
11:   On Error GoTo Err

      Select Case LCase$(strVariable)
        Case "g_apppath": G_APPPATH = CStr(varValue)
        Case "g_errorfile": G_ERRORFILE = CStr(varValue)
        Case "G_GUI_IN_UNLOAD": G_GUI_IN_UNLOAD = CBool(varValue)
        Case "m_blncommadecimal": m_blnCommaDecimal = CBool(varValue)
        Case "m_lngredirectub": m_lngRedirectUB = CLng(varValue)
        Case "m_lngbotsub": m_lngBotsUB = CLng(varValue)
        Case "m_lngbanfilter": m_lngBanFilter = CLng(varValue)
        Case "m_datservingdate": m_datServingDate = CDate(varValue)
        Case "m_datforcednsupdate": m_datForceDNSUpdate = CDate(varValue)
13:   End Select

15:   Exit Sub
16:
Err:
17:   On Error Resume Next
18:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SetVariable(" & strVariable & ", " & varValue & ")"
End Sub
Public Function RunFunction(ParamArray varExpr() As Variant) As Variant
    '------------------------------------------------------------------
    'Purpose:   This Function runs the Subs/Function (in modules)
    '
    'Params:    ParamArray VarExpr()
    '
    'Returns:   Depends of the expr
    '
    '   Called by the plugins functions or scripts
    '------------------------------------------------------------------
    
11:   On Error GoTo Err

      Select Case LCase$(CStr(varExpr(0)))
        '**********************
        Case "msgbeep"
        '**********************
15:           If UBound(varExpr) = 1 Then _
                    MsgBeep CLng(varExpr(1))
        '**********************
        Case "msgboxcenter"
        '**********************
              Select Case UBound(varExpr)
                  Case 2: RunFunction = CLng(MsgBoxCenter(CVar(varExpr(1)), CStr(varExpr(2))))
                  Case 3: RunFunction = CLng(MsgBoxCenter(CVar(varExpr(1)), CStr(varExpr(2)), CLng(varExpr(3))))
                  Case 4: RunFunction = CLng(MsgBoxCenter(CVar(varExpr(1)), CStr(varExpr(2)), CLng(varExpr(3)), CStr(varExpr(4))))
19:           End Select
        '**********************
        Case "addlog"
        '**********************
22:          If UBound(varExpr) = 1 Then _
                 AddLog CStr(varExpr(1))

        '**********************
        Case "handleerror"
        '**********************
              Select Case UBound(varExpr)
                  Case 3: HandleError CLng(varExpr(1)), CStr(varExpr(2)), CLng(varExpr(4))
                  Case 4: HandleError CLng(varExpr(1)), CStr(varExpr(2)), CStr(varExpr(3)), CLng(varExpr(4))
27:           End Select
        '**********************
        Case "utcdate"
        '**********************
30:           If UBound(varExpr) = 1 Then _
                    RunFunction = CLng(UTCDate(CStr(varExpr(1))))
        '**********************
        Case "iiflng"
        '**********************
34:           If UBound(varExpr) = 3 Then _
                    RunFunction = CLng(IIfLng(CBool(varExpr(1)), CLng(varExpr(2)), CLng(varExpr(3))))
        '**********************
        Case "xmlunescape"
        '**********************
38:           If UBound(varExpr) = 1 Then _
                    RunFunction = CStr(XMLUnescape(CStr(varExpr(1))))
        '**********************
        Case "xmlescape"
        '**********************
42:           If UBound(varExpr) = 1 Then _
                    RunFunction = CStr(XMLEscape(CStr(varExpr(1))))
        '**********************
        Case "getbyte"
        '**********************
46:           If UBound(varExpr) = 1 Then _
                    RunFunction = CByte(GetByte(CLng(varExpr(1))))
        '**********************
        Case "debuguser"
        '**********************
50:           If UBound(varExpr) = 1 Then _
                    RunFunction = CStr(DebugUser(CVar(varExpr(1))))
        '**********************
        Case "settaskbarmsg"
        '**********************
54:           If UBound(varExpr) = 2 Then _
                    SetTaskbarMsg CLng(varExpr(1)), CLng(varExpr(2))
        '**********************
        Case "gentempfile"
        '**********************
58:           If UBound(varExpr) = 1 Then _
                    RunFunction = CStr(GenTempFile())
        '**********************
        Case "truetrim"
        '**********************
62:           If UBound(varExpr) = 1 Then _
                    RunFunction = CStr(TrueTrim(CStr(varExpr(1))))
        '**********************
        Case "validip"
        '**********************
66:           If UBound(varExpr) = 1 Then _
                    RunFunction = CStr(ValidIP(CStr(varExpr(1))))
        '**********************
        Case "charcount"
        '**********************
70:           If UBound(varExpr) = 2 Then _
                    RunFunction = CLng(CharCount(CStr(varExpr(1)), CStr(varExpr(2))))
        '**********************
        Case "pause"
        '**********************
74:           If UBound(varExpr) = 1 Then _
                    Pause CLng(varExpr(1))
        '**********************
        Case "frmmove"
        '**********************
78:           If UBound(varExpr) = 1 Then _
                    frmMove CVar(varExpr(1))
        '**********************
        Case "getappversion"
        '**********************
82:            If UBound(varExpr) = 0 Then _
                    RunFunction = CStr(GetAppVersion())
        '**********************
        Case "hubuptime"
        '**********************
86:           If UBound(varExpr) = 0 Then _
                    RunFunction = CStr(HubUpTime())
        '**********************
        Case "painttileformbackground"
        '**********************
90:           If UBound(varExpr) = 1 Then _
                    PaintTileFormBackground CVar(varExpr(1)), LoadImage(g_objSettings.lngSkin)
        '**********************
        Case "painttilepicbackground"
        '**********************
94:           If UBound(varExpr) = 1 Then _
                    PaintTilePicBackground CVar(varExpr(1)), LoadImage(g_objSettings.lngSkin)
        '**********************
        Case "printdebug"
        '**********************
98:          If UBound(varExpr) = 1 Then _
                    PrintDebug CStr(varExpr(1))

        Case "showtooltip"
             Select Case UBound(varExpr)
                 Case 3: ShowToolTip CLng(varExpr(1)), CStr(varExpr(2)), CStr(varExpr(3))
                 Case 4: ShowToolTip CLng(varExpr(1)), CStr(varExpr(2)), CStr(varExpr(3)), CLng(varExpr(4))
                 Case 5: ShowToolTip CLng(varExpr(1)), CStr(varExpr(2)), CStr(varExpr(3)), CLng(varExpr(4)), CLng(varExpr(5))
                 Case 6: ShowToolTip CLng(varExpr(1)), CStr(varExpr(2)), CStr(varExpr(3)), CLng(varExpr(4)), CLng(varExpr(5)), CLng(varExpr(6))
101:         End Select
            
103:   End Select
    
105:  Exit Function
106:
Err:
107:  On Error Resume Next
108:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.RunFunction(" & CStr(varExpr(0)) & ")"
End Function

Public Function PlgCheckIndexByName(ByVal strName As String) As Long
    '------------------------------------------------------------------
    'Purpose:   Useful to check array index of plugin object (g_objPlugin)
    '
    'Params:    strName: name of the plugin
    '
    'Returns:   Plugin index
    '
    '   CheckPlgIndexByName = -1
    '   Plugin name invalid
    '
    '   Ex: lngTemp = frmHub.CheckPlgIndexByName("PTDCH Plugin Template")
    '------------------------------------------------------------------
13:   Dim intIndex As Integer
14:   On Error GoTo Err

16:   For intIndex = LBound(g_objPlugin) To UBound(g_objPlugin)
17:       If g_objPlugin(intIndex).Name = strName Then
18:           PlgCheckIndexByName = intIndex
19:           Exit Function
20:       End If
21:   Next

23:   intIndex = -1

25:   Exit Function
26:
Err:
27:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.PlgCheckIndexByName(" & strName & ")"
End Function
Public Function PlgSwitch(ByVal varIndexOrName As Variant, ByVal blnState As Boolean) As Boolean
    '------------------------------------------------------------------
    'Purpose:   Useful to use for scripts
    '
    'Params:    varIndexOrName: plguin index of he g_objPlugin array point
    '                           or plugin name of the g_objPlugin array point
    '
    'Returns:   True = OK
    '           False = Failed operation
    '------------------------------------------------------------------
10:    On Error GoTo Err
11:    Dim intTemp As Integer
    
13:    If IsNumeric(varIndexOrName) Then
14:        If (CInt(varIndexOrName) > UBound(g_objPlugin)) Or (CInt(varIndexOrName) < 0) Or (Not g_PluginsFound) Then
15:            PlgSwitch = False
16:            Exit Function
17:        Else
18:            intTemp = CInt(varIndexOrName)
19:        End If
20:    Else
21:       intTemp = PlgCheckIndexByName(varIndexOrName)
          'Plugin name found?
23:       If intTemp = -1 Then
24:            PlgSwitch = False
25:            Exit Function
26:       End If
27:    End If

29:    If blnState Then
30:        g_objPlugin(intTemp).Object.RunEvent "Switch", True
31:        lvwPlugins.ListItems(intTemp + 1).Checked = True
32:    Else
33:        g_objPlugin(intTemp).Object.RunEvent "Switch", False
34:        lvwPlugins.ListItems(intTemp + 1).Checked = False
35:    End If
    
37:    PlgSwitch = True

39:    Exit Function
40:
Err:
41:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.PlgSwitch(" & varIndexOrName & ")"
End Function
Public Sub PlgAddMenu(sName As String)
1:    On Error GoTo Err
2:    Dim i As Integer
    
4:    i = mnuPlugIn.Count + 1
      
6:    Load mnuPlugIn(i)
      
8:    mnuPlugIn(i).Visible = True
9:    mnuPlugIn(i).Caption = sName
10:   mnuPlugIn(i).Enabled = True
11:   mnuPlugIn(0).Visible = False
      
13:   Exit Sub
    
15:
Err:
17:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.PlgAddMenu(" & sName & ")"
End Sub
Private Sub mnuPlugIn_Click(Index As Integer)
1:   Dim lngTemp As Long
2:   On Error GoTo Err

4:   If g_objSettings.Plugins Then
5:      lngTemp = PlgCheckIndexByName(mnuPlugIn(Index).Name)
7:      If Not lngTemp = -1 Then _
            g_objPlugin(lngTemp).Object.LoadForm
9:   End If
    
11:  Exit Sub
12:
Err:
13:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuPlugIn_Click(" & Index & ")"
End Sub
Private Sub lvwPlugins_ItemCheck(ByVal Item As MSComctlLib.ListItem)
1:  On Error GoTo Err
    'Check if we should reset or stop
3:  If Item.Checked Then _
           PlgSwitch (Item.Index - 1), True _
    Else PlgSwitch (Item.Index - 1), False
6:  Exit Sub
7:
Err:
8:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwPlugins_ItemCheck(" & Item.Index & ")"
End Sub
Private Sub lvwPlugins_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo Err
    If g_objPlugin(Item.Index - 1).UseSetup Then _
         cmdPlugins(0).Enabled = True _
    Else cmdPlugins(0).Enabled = False
    cmdPlugins(1).Enabled = True
    Exit Sub
Err:
    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwPlugins_ItemClick()"
End Sub
Private Sub lvwPlugins_Click()
1:    On Error GoTo Err
2:    Dim i As Integer
3:    Dim bIsSelected As Boolean
      
5:    For i = 1 To lvwPlugins.ListItems.Count
6:        If lvwPlugins.ListItems.Item(i).Selected Then
7:            bIsSelected = True: Exit For
8:        End If
9:    Next
      
11:   If bIsSelected Then
12:       cmdPlugins(0).Enabled = True
13:       cmdPlugins(1).Enabled = True
14:   Else
15:       cmdPlugins(0).Enabled = False
16:       cmdPlugins(1).Enabled = False
17:   End If

19:   Exit Sub
20:
Err:
21:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwPlugins_Click()"
End Sub
Private Sub cmdPlugins_Click(Index As Integer)
1:    On Error GoTo Err
2:    Dim intIndex As Integer
3:    Dim i As Integer
    
      Select Case Index
        Case 0, 1: intIndex = lvwPlugins.SelectedItem.Index - 1
      End Select
      
      Select Case Index
        Case 0 'Stup
7:            g_objPlugin(intIndex).Object.LoadForm
        Case 1 'Reolad
8:            g_objPlugin(intIndex).Object.RunEvent "Reload"
        Case 2 'Reolad All
9:            For i = LBound(g_objPlugin) To UBound(g_objPlugin)
10:                On Error Resume Next
11:                g_objPlugin(i).Object.RunEvent "Reload"
12:                On Error GoTo Err
13:           Next
        Case 3 'Refresh GUI
14:            PlgRefreshGUI
        Case 4 'ReInstall All
15:           Dim objPlgins As New clsPlugins
           
17:           Call PlgTerm 'Terminate..
           
19:           objPlgins.InstallPlugins True
           
21:           If g_PluginsFound Then
22:                PlgXmlLoad
23:                PlgRefreshGUI
24:           If UBound(g_objPlugin) > 0 Then _
                   cmdPlugins(2).Enabled = True _
              Else cmdPlugins(2).Enabled = False
27:           If UBound(g_objPlugin) >= 0 Then _
                   cmdPlugins(3).Enabled = True _
              Else cmdPlugins(3).Enabled = False
30:           End If
            
32:           Set objPlgins = Nothing
34:    End Select
    
36:    Exit Sub
37:
Err:
38:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdPlugins_Click(" & Index & ")"
End Sub
Public Sub PlgRefreshGUI()
1:    Dim i As Integer
2:    Dim lvwItem As Variant
3:    On Error GoTo Err

5:    If Not g_PluginsFound Then Exit Sub

      lvwPlugins.ListItems.Clear
      cmdPlugins(0).Enabled = False
      cmdPlugins(1).Enabled = False
      
      'Refrsh plugin variables to core..
8:    For i = LBound(g_objPlugin) To UBound(g_objPlugin)
           'Plugin settings
10:        g_objPlugin(i).UseSetup = g_objPlugin(i).Object.UseSetup
11:        g_objPlugin(i).UseEvents = g_objPlugin(i).Object.UseEvents
           'Plugin propertys
14:        g_objPlugin(i).Name = g_objPlugin(i).Object.Name
15:        g_objPlugin(i).Version = g_objPlugin(i).Object.Version
16:        g_objPlugin(i).Author = g_objPlugin(i).Object.Author
17:        g_objPlugin(i).Description = g_objPlugin(i).Object.Description
18:        g_objPlugin(i).ReleaseDate = g_objPlugin(i).Object.ReleaseDate
19:        g_objPlugin(i).Comments = g_objPlugin(i).Object.Comments
           'Add itens to listview
21:        Set lvwItem = lvwPlugins.ListItems.Add((i + 1), (i + 1) & "s", g_objPlugin(i).Name)
22:        lvwItem.SubItems(1) = CBool(g_objPlugin(i).UseEvents)
23:        lvwItem.SubItems(2) = g_objPlugin(i).Version
24:        lvwItem.SubItems(3) = g_objPlugin(i).Author
25:        lvwItem.SubItems(4) = g_objPlugin(i).Description
26:        lvwItem.SubItems(5) = g_objPlugin(i).ReleaseDate
27:        lvwItem.SubItems(6) = g_objPlugin(i).Comments
28:        lvwItem.Checked = CBool(g_objPlugin(i).Object.Enabled)
29:    Next
    
30:    Set lvwItem = Nothing
    
32:    Exit Sub
33:
Err:
34:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.PlgRefreshGUI()"
End Sub
Public Property Get PlgObjectByIndex(intIndex As Integer) As Object
1:    On Error Resume Next
      If g_PluginsFound Then _
           Set PlgObjectByIndex = g_objPlugin(intIndex).Object _
      Else Set PlgObjectByIndex = Nothing
End Property
Public Property Get PlgObjectByName(intIndex As Integer) As Object
1:    On Error Resume Next
      If g_PluginsFound Then _
           Set PlgObjectByName = g_objPlugin(intIndex).Object _
      Else Set PlgObjectByName = Nothing
End Property
Public Property Get PlgCount() As Integer
1:    On Error Resume Next
2:    PlgCount = (UBound(g_objPlugin) + 1)
End Property
Private Sub PlgTerm()
1:    Dim i As Integer
2:    On Error GoTo Err
    
     'Terminate plugins
8:    On Error Resume Next 'Run UnloadMain event and Term
9:    If g_PluginsFound And g_objSettings.Plugins Then
10:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
11:            If g_objPlugin(i).UseEvents Then
12:                 If g_objPlugin(i).Object.Enabled Then
13:                      g_objPlugin(i).Object.RunEvent "UnloadMain"
14:                      g_objPlugin(i).Object.Term
15:                 End If
16:            End If
17:            Set g_objPlugin(i).Object = Nothing
19:        Next
20:   End If

22:   Erase g_objPlugin
      
24:   Exit Sub
25:
Err:
26:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.PlgTerm()"
End Sub
Public Sub PlgXmlLoad()
1:   Dim objXML          As clsXMLParser
2:   Dim objNode         As clsXMLNode
3:   Dim colNodes        As Collection
4:   Dim colSubNodes     As Collection

6:   Dim strTemp         As String
7:   Dim i               As Integer
        
9:   On Error GoTo Err

11:    Set objXML = New clsXMLParser
      
13:    strTemp = G_APPPATH & "\Settings\Plugins.xml"

15:    If g_objFileAccess.FileExists(strTemp) And g_PluginsFound Then
         
17:       objXML.Data = g_objFileAccess.ReadFile(strTemp)
18:       objXML.Parse

20:       Set colNodes = objXML.Nodes(1).Nodes

22:       On Error Resume Next

24:       For Each objNode In colNodes
25:            Set colSubNodes = objNode.Attributes
26:            For i = LBound(g_objPlugin) To UBound(g_objPlugin)
27:                 If g_objPlugin(i).Name = CStr(colSubNodes("Name").Value) Then
28:                       g_objPlugin(i).Object.RunEvent "SubMain"
29:                       If CBool(colSubNodes("Value").Value) Then
30:                            PlgSwitch i, True
31:                       End If
32:                 End If
33:            Next
34:       Next

36:       On Error GoTo Err
    
38:       objXML.Clear
    
40:       Set objNode = Nothing
41:       Set colSubNodes = Nothing
42:       Set colNodes = Nothing

44:   End If

48:   Exit Sub
    
50:
Err:
51:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.PlgXmlLoad()"
End Sub
Private Sub PlgXmlSave()
1:    On Error GoTo Err
2:    Dim intFF       As Integer
3:    Dim strTemp     As String
4:    Dim i           As Integer

6:    strTemp = G_APPPATH & "\Settings\Plugins.xml"

8:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp
 
10:   intFF = FreeFile

12:    Open strTemp For Append As intFF
13:       Print #intFF, "<Plugins>"
14:           If g_PluginsFound Then
15:              For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:                   Print #intFF, vbTab & "<Plugin Name=""" & g_objPlugin(i).Name & _
                                           """" & " Value=""" & CBool(g_objPlugin(i).Object.Enabled) & """ />"
18:              Next
19:           End If
20:       Print #intFF, "</Plugins>";
21:    Close intFF
    
23:   Exit Sub
24:
Err:
26:   On Error Resume Next
      'In case of mistake it cancels all plugins!!
28:   If g_objFileAccess.FileExists(G_APPPATH & "\Plugins\Plugins.xml") Then
29:        g_objFileAccess.DeleteFile G_APPPATH & "\Plugins\Plugins.xml"
30:   End If
31:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.PlgXmlSave()"
End Sub
'------------------------------------------------------------------------------
'End Plugins Subs
'------------------------------------------------------------------------------

Private Sub lvwRegistered_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:  On Error GoTo Err
    
    'Sort listview by column clicked
4:   lvwRegistered.SortKey = ColumnHeader.Index - 1
5:   lvwRegistered.SortOrder = IIfLng(lvwRegistered.SortOrder, lvwAscending, lvwDescending)
6:   lvwRegistered.Sorted = True
    
8:   Exit Sub
9:
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwRegistered_ColumnClick()"
End Sub

Private Sub lvwRegistered_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(1)
End Sub

#If Status Then
    Private Sub lvwUsers_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:      On Error GoTo Err
    
        'Sort listview by column clicked
4:      lvwUsers.SortKey = ColumnHeader.Index - 1
5:      lvwUsers.SortOrder = IIfLng(lvwUsers.SortOrder, lvwAscending, lvwDescending)
6:      lvwUsers.Sorted = True
    
8:      Exit Sub
9:
Err:
10:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwTempIPBan_ColumnClick()"
    End Sub
    Private Sub lvwUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    Dim lngSelected As Long

4:    lngSelected = IsListViewSelected(lvwUsers)

6:    If lngSelected <> -1 Then
7:       mnuUsers(0).Enabled = True: mnuUsers(2).Enabled = True
8:       mnuUsers(3).Enabled = True: mnuUsers(4).Enabled = True
9:       mnuUsers(5).Enabled = True: mnuUsers(6).Enabled = True
10:      mnuUsers(7).Enabled = True
11:   Else
12:      mnuUsers(0).Enabled = False: mnuUsers(2).Enabled = False
13:      mnuUsers(3).Enabled = False: mnuUsers(4).Enabled = False
14:      mnuUsers(5).Enabled = False: mnuUsers(6).Enabled = False
15:      mnuUsers(7).Enabled = False
16:   End If

18:   If Button = 2 Then PopupMenu mnuPopUp(7)

20:   Exit Sub
21:
Err:
22:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwUsers_MouseDown()"
    End Sub
#End If

Private Sub lstTagsEx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(5)
End Sub

Private Sub lvwPermIPBan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(3)
End Sub

Private Sub lvwTempIPBan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(2)
End Sub

Private Sub lstTagsDef_Click()
' ------------------------ NEW MOD INTERFACE LANGUAGE ------------------------
2:    If lstTagsDef.ListIndex = -1 Then
3:        txtTagRules.Text = Replace(m_arrTagRules(10), "%[LF]", vbNewLine)
4:    Else
5:        txtTagRules.Text = Replace(m_arrTagRules(lstTagsDef.ListIndex), "%[LF]", vbNewLine)
6:    End If

'    Select Case lstTagsDef.ListIndex
        ' ++
'        Case 0: txtTagRules.Text = "'NoHello' even if it's not in $Supports statement." & vbNewLine & vbNewLine & "Tests for minimum DC++ version V:______"
        ' DC
'        Case 1: txtTagRules.Text = "Skips standard DC++ O:# tests." & vbNewLine & vbNewLine & "O:# is used for free/open slots."
        ' DCGUI
'        Case 2: txtTagRules.Text = "* in it's slot param (S:*) means unlimited slots." & vbNewLine & vbNewLine & "* in it's limiter param (L:*) means unlimited bandwidth." & vbNewLine & vbNewLine & "Reports bandwidth limit on a per slot basis, not total."
        ' DC:Pro
'        Case 5: txtTagRules.Text = "Uses F:#Down/#Up to report bandwidth limiting."
        ' SdDC++
'        Case 8: txtTagRules.Text = "slot param has the format S:#/#"
        ' Chat (Gadgets Flash Add-on)
'        Case 9: txtTagRules.Text = "If you are using this option then you can figure it out for yourself."
'        Case Else: txtTagRules.Text = "None"
'    End Select
' ----------------------- NEW MOD INTERFACE LANGUAGE END ----------------------
End Sub

Private Sub lstTagsDef_LostFocus()
1:    lstTagsDef.ListIndex = -1
' ------------------------ NEW MOD INTERFACE LANGUAGE ------------------------
3:    txtTagRules.Text = Replace(m_arrTagRules(10), "%[LF]", vbNewLine)
'    txtTagRules.Text = "Select a Default Tag to see if it has any special processing rules."
' ----------------------- NEW MOD INTERFACE LANGUAGE END ----------------------
End Sub

'------------------------------------------------------------------------------
'Detect IP events
'------------------------------------------------------------------------------
Private Sub m_objDetectIP_OnDownloaded(strHeader As String, strData As String)
1:    Dim strTemp As String
    
3:    On Error GoTo Err
    
    'Skip vbNewLine in front
6:    strTemp = MidB$(strData, 5)
    
8:    If MsgBox(Replace(g_colMessages.Item("msgYourIP"), "%[IP]", strTemp), vbYesNo Or vbQuestion, g_colMessages.Item("msgDetectIP")) = vbYes Then
9:        Clipboard.Clear
10:        Clipboard.SetText strTemp
11:    End If
        
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.m_objDetectIP_OnDownloaded(""" & strHeader & """, """ & strData & """)"
End Sub
Private Sub m_objDetectIP_OnError(ByVal lngNumber As Long, strDescription As String)
1:    If MsgBox(Replace(g_colMessages.Item("msgIPError"), "%[IP]", wskListen(0).LocalIP), vbYesNo, g_colMessages.Item("msgDetectIP")) = vbYes Then
2:        Clipboard.Clear
3:        Clipboard.SetText wskListen(0).LocalIP
4:    End If
End Sub
Private Sub mnuCodeRTB_Click(Index As Integer)
1:    On Error GoTo Err
2:    Dim i As Integer
3:    Dim Modal As Byte
     
5:    If Index = 5 Then
6:        frmHelp.Show Modal, Me
7:        Exit Sub
8:    End If
   
10:   For i = 1 To UBound(g_objSciLexer)
      Select Case Index
            Case 0 'View WhiteSpace
11:               If i = 1 Then mnuCodeRTB(0).Checked = Not mnuCodeRTB(0).Checked
12:               g_objSciLexer(i).ViewWhiteSpace = mnuCodeRTB(0).Checked
            Case 1 'Line Number
13:               If i = 1 Then mnuCodeRTB(1).Checked = Not mnuCodeRTB(1).Checked
14:               g_objSciLexer(i).LineNumbers = mnuCodeRTB(1).Checked
            Case 7 'Word Wrap
15:               If i = 1 Then mnuCodeRTB(7).Checked = Not mnuCodeRTB(7).Checked
16:               g_objSciLexer(i).WordWrap = mnuCodeRTB(7).Checked
            Case 8 'ReadOnly
17:               If i = 1 Then mnuCodeRTB(8).Checked = Not mnuCodeRTB(8).Checked
18:               g_objSciLexer(i).ReadOnly = mnuCodeRTB(8).Checked
19:         End Select
20:   Next

22:   If Index <> 5 Then SCI_Focus
    
24:   Exit Sub
25:
Err:
26:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuCodeRTB_Click(" & Index & ")"
End Sub

'------------------------------------------------------------------------------
'Menu events
'------------------------------------------------------------------------------
Private Sub mnuPlan_Click(Index As Integer)
1:  On Error GoTo Err

  Select Case Index
        Case 0: g_objScheduler.ShowAddDialog Me
        Case 1, 3
3:          Dim lvwItem     As ListItem
4:          Dim lvwItems    As ListItems
            '
6:          If lvwPlan.ListItems.Count = 0 Then Exit Sub
7:          Set lvwItem = lvwPlan.SelectedItem

9:         If lvwPlan.ListItems.Count = 0 Then Exit Sub
            '
11:         If lvwItem.Selected Then
12:            If Index = 3 Then
13:               g_objScheduler.RemPlan lvwItem.Text, _
                                         lvwItem.SubItems(3), _
                                         lvwItem.SubItems(4)
16:            ElseIf Index = 1 Then
17:               g_objScheduler.ShowEditDialog lvwItem.Text, _
                                                lvwItem.SubItems(3), _
                                                lvwItem.SubItems(4), Me
20:            End If
21:         End If
22:      End Select
     
24:   Exit Sub
25:
Err:
26:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuPlan_Click(" & Index & ")"
End Sub

Private Sub mnuLocked_Click(Index As Integer)
1:  Dim lvwItem     As ListItem
2:  Dim lvwItems    As ListItems
3:  Dim strTxt      As String
4:  Set lvwItem = lvwBans.SelectedItem
5:  Set lvwItems = lvwBans.ListItems
  
7:  On Error GoTo Err

9:  If lvwItems.Count = 0 Then Exit Sub
    'Check if selected
11:    If lvwItem.Selected Then
12:        Clipboard.Clear
           Select Case Index
               Case 0 'Copy User Name
13:                Clipboard.SetText (lvwItem.Text)
               Case 2 'Copy All
14:                strTxt = "PT DC Hub " & vbVersion & " - Ban Name" & vbNewLine & vbNewLine & _
                         "User Name: " & lvwItem.Text & vbNewLine & _
                         "Perm: " & lvwItem.SubItems(1) & vbNewLine & _
                         "Banned By: " & lvwItem.SubItems(2) & vbNewLine & _
                         "Reference Date: " & lvwItem.SubItems(3) & vbNewLine & _
                         "Reason: " & lblHolder(50).Caption
20:                Clipboard.SetText (strTxt)
21:          End Select
22:    End If
    
24:  Exit Sub
25:
Err:
26:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuLocked_Click(" & Index & ")"
End Sub

Private Sub mnuPermIPBan_Click(Index As Integer)
1:    Dim strIP       As String
2:    Dim lngMinutes  As Long
3:    Dim varLoop     As Variant
4:    Dim lvwItems    As ListItems
5:    Dim X           As Variant
6:    Dim colBans     As Collection
7:    Dim objTB       As clsIPBansData
      
9:    On Error GoTo Err
    
      Select Case Index
        Case 0 'Add
11:            frmBanPerm.Show vbModal, Me
12:            GoTo RefreshN
        Case 1 'Remove
13:            If ObjPtr(lvwPermIPBan.SelectedItem) Then
14:                strIP = lvwPermIPBan.SelectedItem.Text
15:                lvwPermIPBan.ListItems.Remove CStr(strIP & "s")
16:                g_colIPBans.Remove strIP, 2
'17:           Else
'18:               strIP = InputBox(g_colMessages.Item("msgEnterRemIP"), g_colMessages.Item("msgRemoveIP"))
'19:               If LenB(strIP) Then g_colIPBans.Remove strIP
20:            End If
21:            GoTo RefreshN
        Case 2 'Clear
22:            If MsgBoxCenter(Me, g_colMessages.Item("msgClearPermIPs"), vbYesNo Or vbExclamation, g_colMessages.Item("msgConfirmClear")) = vbYes Then
23:                 g_colIPBans.ClearPerm
24:                 GoTo RefreshN
25:            End If
        Case 4 'Refresh list extract
RefreshN:
27:            Set colBans = g_colIPBans.PermItems
28:            Set lvwItems = lvwPermIPBan.ListItems
        
               'Clear out items first
31:            lvwItems.Clear
            
               Select Case m_lngBanFilter
                  Case 0 'No filter
33:                    For Each objTB In colBans
34:                        Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
35:                        X.SubItems(1) = objTB.Nick
36:                        X.SubItems(2) = objTB.BannedBy
37:                        X.SubItems(3) = objTB.Reason
38:                    Next
                  Case 1 'End in
39:                    strIP = txtBanFilter.Text
40:                    lngMinutes = LenB(strIP)
41:                    For Each objTB In colBans
42:                        If RightB$(objTB.IP, lngMinutes) = strIP Then
43:                           Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
44:                           X.SubItems(1) = objTB.Nick
45:                           X.SubItems(2) = objTB.BannedBy
46:                           X.SubItems(3) = objTB.Reason
47:                        End If
48:                    Next
                  Case 2 'Contain
49:                    strIP = txtBanFilter.Text
50:                    For Each objTB In colBans
51:                        If InStrB(1, objTB.IP, strIP) Then
52:                           Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
53:                           X.SubItems(1) = objTB.Nick
54:                           X.SubItems(2) = objTB.BannedBy
55:                           X.SubItems(3) = objTB.Reason
56:                        End If
57:                    Next
                  Case 3 'Begin with
58:                    strIP = txtBanFilter.Text
59:                    lngMinutes = LenB(strIP)
60:                    For Each objTB In colBans
61:                        If LeftB$(objTB.IP, lngMinutes) = strIP Then
62:                           Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
63:                           X.SubItems(1) = objTB.Nick
64:                           X.SubItems(2) = objTB.BannedBy
65:                           X.SubItems(3) = objTB.Reason
66:                        End If
67:                    Next
68:            End Select
        Case 5 'Clear list extract
69:            lvwPermIPBan.ListItems.Clear
70:    End Select
    
72:    Exit Sub

74:
Err:
75:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuPermIPBan_Click(" & Index & ")"
End Sub
Public Sub mnuScripts_Click(Index As Integer)
1:  On Error GoTo Err
2:    Dim i As Integer

4:    For i = 1 To picSciMain.Count - 1
5:            If picSciMain(i).Visible Then Exit For
6:    Next

8:    With frmScript
        Select Case Index
                Case 0 'Save/Reset
9:                 Call .SReset(i, True, True)
                Case 2 'Stop
10:                 .SStop i
                Case 3 'Stop All
11:                 .SStop -2
                Case 5 'Reolad Checkeds
12:                 Call .SReset(-2, True, True)
                Case 6 'Reolad Dir
13:                 .XmlBooleanSave
14:                 .SLoadDir
15:                 .XmlBooleanLoad
16:                 .SReset -2, False, False
                Case 8 'Properties
17:                 .SProperties CStr(i & "s"), lvwScripts.ListItems(i).Text, 0
18:          End Select
19:  End With

21:  If Index <> 8 Then SCI_Focus

23:  Exit Sub
24:
Err:
25:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuScripts_Click(" & Index & ")"
End Sub

#If Status Then
    Private Sub lstStatus_DblClick(Index As Integer)
1:       On Error GoTo Err
2:       With frmMulti
3:           .cmdCancel.Visible = False
4:           .Label1.Visible = False
5:           .Caption = "Copy to Clipboard Text"
6:           .txtStr.Top = 120
7:           .cmdOK.Top = 520
8:           .Height = 1350
9:           .txtStr.Text = CStr(lstStatus(Index))
10:          .Show vbModal, Me
11:       End With
12:      Exit Sub
13:
Err:
14:      HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lstStatus_DblClick(" & Index & ")"
    End Sub
    Private Sub lstStatus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:        If Button = 2 Then PopupMenu mnuPopUp(6)
    End Sub
    Private Sub mnuStatus_Click(Index As Integer)
1:        Dim lngLoop     As Long
2:        Dim lngUB       As Long
3:        Dim strCopy     As String
4:        Dim i           As Integer
    
6:        On Error GoTo Err
          
8:           If picStatus(0).Visible Then
9:             i = 0
10:          ElseIf picStatus(1).Visible Then
11:             i = 1
12:          ElseIf picStatus(2).Visible Then
13:             i = 2
14:          End If
          
          Select Case Index
                Case 0 'Copy
16:                lngUB = lstStatus(i).ListCount - 1
            
                  'Clear the clipboard before we start
19:                Clipboard.Clear
                
                  'Loop through and find all selected items
22:                For lngLoop = 0 To lngUB
23:                    If lstStatus(i).Selected(lngLoop) Then strCopy = strCopy & lstStatus(i).List(lngLoop) & vbNewLine
24:                Next
                
                  'Set the clipboard to all selected text
27:                Clipboard.SetText strCopy
            
               Case 1 'Clear
29:                g_objStatus.MClear i
30:        End Select
    
32:        Exit Sub
    
34:
Err:
35:        HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuStatus_Click(" & Index & ")"
    End Sub
    Private Sub mnuUsers_Click(Index As Integer)
1:        Dim lvwItem     As ListItem
2:        Dim lvwItems    As ListItems
3:        Dim strOne      As String
4:        Dim strTwo      As String
5:        Dim intOne      As Integer
6:        Dim lngOne      As Long

8:        On Error GoTo Err
        
          'Get selected item
11:       Set lvwItem = lvwUsers.SelectedItem

          'Get listitem collection
14:       Set lvwItems = lvwUsers.ListItems

          Select Case Index
            'Send data (selected)
            Case 0

                   'Get message
19:                With frmMulti
20:                   .Caption = g_colMessages.Item("msgSendToSel")
21:                   .Label1.Caption = g_colMessages.Item("msgEnterDataToSel")
22:                   .Show vbModal, Me
23:                   strOne = .txtStr.Text
24:                   Set frmMulti = Nothing
25:                End With

27:                If LenB(strOne) Then
28:                    For Each lvwItem In lvwItems
                        'Send if selected
30:                        If lvwItem.Selected Then _
                                g_colUsers.ItemByName(lvwItem.Text).SendData strOne
32:                    Next
33:                End If

            'Send data (all)
            Case 1
            
                   'Get message
38:                With frmMulti
39:                   .Caption = g_colMessages.Item("msgSendToAll")
40:                   .Label1.Caption = g_colMessages.Item("msgEnterDataToAll")
41:                   .Show vbModal, Me
42:                   strOne = .txtStr.Text
43:                   Set frmMulti = Nothing
44:                End With

46:                If LenB(strOne) Then g_colUsers.SendToAll strOne

            'Disconnect
            Case 2
            
50:                If ObjPtr(lvwItem) = 0 Then Exit Sub

52:                wskLoop_Close CInt(lvwItem.SubItems(2))

            'Kick
            Case 3
            
56:                If ObjPtr(lvwItem) = 0 Then Exit Sub

                   'Get ban length
59:                With frmMulti
60:                   .Caption = g_colMessages.Item("msgKickSel")
61:                   .Label1.Caption = g_colMessages.Item("msgEnterLength")
62:                   .Show vbModal, Me
63:                   strTwo = .txtStr.Text
64:                   Set frmMulti = Nothing
65:                End With

67:                If LenB(strTwo) Then lngOne = CLng(Val(strTwo)) Else Exit Sub

                  'Get reason
70:                With frmMulti
71:                   .Caption = g_colMessages.Item("msgKick")
72:                   .Label1.Caption = g_colMessages.Item("msgKickReason")
73:                   .Show vbModal, Me
74:                   strOne = .txtStr.Text
75:                   Set frmMulti = Nothing
76:                End With

78:                If LenB(strOne) Then
79:                    Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))
80:                    m_objLoopUser.SendChat g_objSettings.BotName, strOne
81:                    DoEvents
82:                    m_objLoopUser.Kick lngOne, "Admin / GUI"

84:                    Set m_objLoopUser = Nothing
85:                End If

            'Redirect
            Case 4
            
89:                If ObjPtr(lvwItem) = 0 Then Exit Sub

                   'Get address
92:                With frmMulti
93:                   .Caption = g_colMessages.Item("msgRedir")
94:                   .Label1.Caption = g_colMessages.Item("msgEnterRedirAddress")
95:                   .Show vbModal, Me
96:                   strTwo = .txtStr.Text
97:                   Set frmMulti = Nothing
98:                End With

100:               If LenB(strTwo) = 0 Then Exit Sub

                   'Get reason
103:               With frmMulti
104:                   .Caption = g_colMessages.Item("msgRedir")
105:                   .Label1.Caption = g_colMessages.Item("msgRedirReason")
106:                   .Show vbModal, Me
107:                   strOne = .txtStr.Text
108:                   Set frmMulti = Nothing
109:               End With

111:               If LenB(strOne) Then
112:                   Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))

114:                   m_objLoopUser.SendChat g_objSettings.BotName, strOne
115:                   m_objLoopUser.Redirect strTwo

117:                   Set m_objLoopUser = Nothing
118:               End If
            'Ban
            Case 5
120:                If ObjPtr(lvwItem) = 0 Then Exit Sub

                    'Get message (if we're sending one)
123:                With frmMulti
124:                   .Caption = g_colMessages.Item("msgBan")
125:                   .Label1.Caption = g_colMessages.Item("msgEnterBanReason")
126:                   .Show vbModal, Me
127:                   strOne = .txtStr.Text
128:                   Set frmMulti = Nothing
129:                End With

131:                If LenB(strOne) Then
132:                    Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))

134:                    m_objLoopUser.SendChat g_objSettings.BotName, strOne
135:                    DoEvents
136:                    m_objLoopUser.Ban "Admin / GUI"

138:                    Set m_objLoopUser = Nothing
139:                End If

            '(De)mute
            Case 6
            
143:                If ObjPtr(lvwItem) = 0 Then Exit Sub

                'Swap mute status
146:                Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))
147:                m_objLoopUser.Mute = Not m_objLoopUser.Mute
148:                Set m_objLoopUser = Nothing

            'Properties (selected)
            Case 7
            
152:                For Each lvwItem In lvwItems
                    'Check if selected
154:                    If lvwItem.Selected Then
                        'Get object / settings
156:                        Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))

                        'Create property string
159:                        strOne = "Name : " & m_objLoopUser.sName & vbNewLine & _
                                 "Winsock Index : " & m_objLoopUser.iWinsockIndex & vbNewLine & _
                                 "IP : " & m_objLoopUser.IP & vbNewLine & _
                                 "Connected Since : " & m_objLoopUser.ConnectedSince & vbNewLine & _
                                 "Class : " & m_objLoopUser.Class & vbNewLine & _
                                 "Language : " & m_objLoopUser.sLanguageID & vbNewLine & _
                                 "Version : " & m_objLoopUser.iVersion & vbNewLine & _
                                 "Share : " & g_objFunctions.ShareSize(m_objLoopUser.iBytesShared) & vbNewLine & _
                                 "MyINFO : " & m_objLoopUser.sMyInfoString & vbNewLine & _
                                 "Supports : " & m_objLoopUser.Supports '& vbTwoLine

                        'Append to collection
171:                        strTwo = strTwo & strOne
172:                    End If
173:                Next

175:                Set m_objLoopUser = Nothing

177:                frmUserInfo.txtInfo = strTwo
178:                frmUserInfo.Show vbModal, Me
179:                Set frmUserInfo = Nothing
                    
181:        End Select

183:        RefreshGUI

185:        Exit Sub

187:
Err:
189:        HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuUsers_Click(" & Index & ")"
    End Sub
    Private Sub txtStForm_Change()
1:       If txtStForm.Text = "" Then txtStForm.Text = g_objSettings.BotName
    End Sub
    Private Sub txtStSend_KeyPress(KeyAscii As Integer)
1:      If KeyAscii = 13 Then _
            KeyAscii = 0: Call cmdStSend_Click
    End Sub
    Private Sub sldStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:       picStInfo.Top = lstStatus(0).Height - picStInfo.Height + 50
2:       picStInfo.Left = lstStatus(0).Width - picStInfo.Width + 50
3:       picStInfo.Visible = True
    End Sub
    Private Sub sldStatus_Change()
1:      If optStSend(1).Value And sldStatus.Value = 6 Then _
             cmdStSend.Enabled = False: txtStSend.Enabled = False _
        Else cmdStSend.Enabled = True: txtStSend.Enabled = True
4:      Dim i As Integer
5:      For i = 2 To 7
6:            lblStatus(i).FontUnderline = False
7:      Next
8:      lblStatus(sldStatus.Value + 1).FontUnderline = True
    End Sub
    Private Sub cmdStSend_Click()
1:        Dim strMsg   As String
2:        Dim srtForm  As String
     
4:        On Error GoTo Err
     
6:         strMsg = txtStSend.Text
7:         srtForm = txtStForm.Text
8:         txtStSend.Text = ""

10:        If optStSend(0).Value Then 'Send Chat
              Select Case sldStatus.Value
                Case 1 'Send Chat To All
11:                    g_colUsers.SendChatToAll srtForm, strMsg
12:                    AddLog "Send Chat To All :" & "<" & srtForm & "> " & strMsg
                Case 2 'Send Chat To Op
13:                    g_colUsers.SendChatToOps srtForm, strMsg
14:                    AddLog "Send Chat To Op: " & "<" & srtForm & "> " & strMsg
                Case 3 'Send Chat To UnRegistered
15:                    g_colUsers.SendChatToUnReg srtForm, strMsg
16:                    AddLog "Send Chat To UnRegistered: " & "<" & srtForm & "> " & strMsg
                Case 4 'Send PM To All
17:                    g_colUsers.SendPrivateToAll srtForm, strMsg
18:                    AddLog "Send PM To All: " & "< " & srtForm & " > " & strMsg
                Case 5 'Send PM To Op
19:                    g_colUsers.SendPrivateToOps srtForm, strMsg
20:                    AddLog "Send PM To Op: " & "<" & srtForm & ">" & strMsg
                Case 6 'Send PM To UnRegistered
21:                    g_colUsers.SendPrivateToUnReg srtForm, strMsg
22:                    AddLog "Send PM To UnRegistered: " & "<" & srtForm & ">" & strMsg
23:              End Select
24:           SEvent_MassMessage strMsg
              #If Status Then
26:                g_objStatus.MAdd "<" & srtForm & "> " & strMsg
              #End If
28:        ElseIf optStSend(1) Then 'Send Data
              Select Case sldStatus.Value
                Case 1 'Send Data To All
29:                    g_colUsers.SendToAll strMsg
30:                    AddLog "Send Data To All: " & strMsg
                Case 2 'Send Data To Op
31:                    g_colUsers.SendToOps strMsg
32:                    AddLog "Send Data To Op: " & strMsg
                Case 3 'Send Data To UnRegistered
33:                    g_colUsers.SendToUnReg strMsg
34:                    AddLog "Send Data To UnRegistered: " & strMsg
                Case 4 'Send Data No Away Mode
35:                    g_colUsers.SendToNA strMsg
36:                    AddLog "Send Data No Away Mode: " & strMsg
                Case 5 'Send Data Non Quick List Clients
37:                    g_colUsers.SendToNQ strMsg
38:                    AddLog "Send Data Non Quick List Clients: " & strMsg
                Case 6
                        '
                        '
41:              End Select
              #If Status Then
43:                g_objStatus.MAdd strMsg
              #End If
45:        End If
46:     Exit Sub
47:
Err:
48:        HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdStSend_Click()"
    End Sub
    Private Sub cmdStSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:       picStInfo.Visible = False
    End Sub
    Private Sub lblStatus_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:       picStInfo.Visible = False
    End Sub
    Private Sub optStSend_Click(Index As Integer)
1:       On Error GoTo Err
         Select Case Index
             Case 0
2:               lblStatus(1).Caption = "------------- Send Chat -------------"
3:               lblStatus(2).Caption = "1 = Send Chat To All"
4:               lblStatus(3).Caption = "2 = Send Chat To Op"
5:               lblStatus(4).Caption = "3 = Send Chat To UnRegistered"
6:               lblStatus(5).Caption = "4 = Send PM To All"
7:               lblStatus(6).Caption = "5 = Send PM To Op"
8:              lblStatus(7).Caption = "6 = Send PM To UnRegistered"
9:              txtStForm.Enabled = True
10:              txtStForm.BackColor = &H80000005
             Case 1
11:              lblStatus(1).Caption = "------------- Send Data -------------"
12:              lblStatus(2).Caption = "1 = Send Data To All"
13:              lblStatus(3).Caption = "2 = Send Data To Op"
14:              lblStatus(4).Caption = "3 = Send Data To UnRegistered"
15:              lblStatus(5).Caption = "4 = Send Data No Away Mode"
16:              lblStatus(6).Caption = "5 = Send Data Non QuickListClients"
17:              lblStatus(7).Caption = "6 = ----------------"
18:              txtStForm.Enabled = False
19:              txtStForm.BackColor = &H8000000F
20:         End Select
21:      sldStatus.Value = 1
22:      Exit Sub
23:
Err:
24:      HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.optStSend_Click(" & Index & ")"
    End Sub
    Private Sub optStSend_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:       picStInfo.Visible = False
    End Sub
#End If

Private Sub picLog_Click(Index As Integer)
   If Index = 2 Then
      On Error Resume Next
      ShellExecute Me.hwnd, "open", "https://www.paypal.com/xclick/business=carlosferreiracarlos%40hotmail.com&item_name=Carlos+Ferreira&currency_code=EUR", "", "", SW_SHOWMAXIMIZED
   End If
End Sub

Private Sub mnuTags_Click(Index As Integer)
1:    Dim lngLoop     As Long
2:    Dim lngUB       As Long
3:    Dim strTag      As String
4:    Dim objTag      As clsTag
    
6:    On Error GoTo Err
    
      Select Case Index
        Case 0 'Add
8:            With frmMulti
9:               .Caption = g_colMessages.Item("msgAddTag")
10:               .Label1.Caption = g_colMessages.Item("msgEnterTag")
11:               .Show vbModal, Me
12:               strTag = .txtStr.Text
13:               Set frmMulti = Nothing
14:            End With

16:            If LenB(strTag) Then
                'Make sure the tag isn't already in the list
18:                On Error Resume Next
19:                m_colTags.Item strTag
                
21:                If Err.Number Then
22:                    On Error GoTo Err
                
                    'Add to list
25:                    lstTagsEx.AddItem strTag
                    
                    'Add to collection
28:                    Set objTag = New clsTag
                    
30:                    objTag.Name = strTag
                    'If a user re-adds a default Tag give it the right default ID
                    Select Case objTag.Name
                        Case "++": objTag.ID = 1
                        Case "DC": objTag.ID = 2
                        Case "DCGUI": objTag.ID = 3
                        Case "oDC": objTag.ID = 4
                        Case "QuickDC": objTag.ID = 5
                        Case "DC:Pro": objTag.ID = 6
                        Case "SDC": objTag.ID = 7
                        Case "StrgDC++": objTag.ID = 10
                        Case "SdDC++": objTag.ID = 8
                        Case "Z++": objTag.ID = 11
                        Case "Chat": objTag.ID = 9
                        Case Else: objTag.ID = -1
32:                    End Select
                    
34:                    m_colTags.Add objTag, strTag
                    
36:                    Set objTag = Nothing
37:                Else
38:                    MsgBoxCenter Me, strTag & g_colMessages.Item("msgAlreadyAdded"), vbInformation, "PTDCH"
39:                End If
40:            End If
        Case 1 'Remove
            'Make sure some tags are selected before looping through
42:            If lstTagsEx.SelCount Then
43:                lngUB = lstTagsEx.ListCount - 1
            
45:                For lngLoop = 0 To lngUB
46:                    If lstTagsEx.Selected(lngLoop) Then
                        'Remove from collection/list
48:                        m_colTags.Remove lstTagsEx.List(lngLoop)
49:                        lstTagsEx.RemoveItem lngLoop
                        
51:                        Exit For
52:                    End If
53:                Next
54:            End If
55:    End Select
    
57:    Exit Sub
    
59:
Err:
60:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuTags_Click(" & Index & ")"
End Sub

Private Sub mnuTempIPBan_Click(Index As Integer)
1:    Dim strIP       As String
2:    Dim lngMinutes  As Long
3:    Dim lvwItems    As Variant
4:    Dim X           As Variant
5:    Dim colBans     As Collection
6:    Dim objTB       As clsIPBansData
    
8:    On Error GoTo Err
    
      Select Case Index
        Case 0 'Add
10:            With frmBanTemp
11:               .Caption = g_colMessages.Item("msgBanTempIP")
12:               .Labels(0).Caption = g_colMessages.Item("msgEnterBanLength")
13:               .Labels(1).Caption = g_colMessages.Item("msgEnterBanLength")
14:               .Show vbModal, Me
15:            End With
16:            GoTo RefreshN
        Case 1 'Remove
17:            If ObjPtr(lvwTempIPBan.SelectedItem) Then
18:               strIP = lvwTempIPBan.SelectedItem.Text
19:               lvwTempIPBan.ListItems.Remove CStr(strIP & "s")
20:               g_colIPBans.Remove strIP, 1
21:            Else
'22:                strIP = InputBox(g_colMessages.Item("msgEnterRemIP"), g_colMessages.Item("msgRemoveIP"))
'23:                If LenB(strIP) Then g_colIPBans.Remove strIP
24:            End If
25:            GoTo RefreshN
        Case 2 'Clear
26:            If MsgBoxCenter(Me, g_colMessages.Item("msgClearTempIPs"), vbYesNo Or vbExclamation, g_colMessages.Item("msgConfirmClear")) = vbYes Then
27:                    g_colIPBans.ClearTemp
28:                 GoTo RefreshN
29:            End If
        Case 4 'Refresh list extract
RefreshN:
31:            Set colBans = g_colIPBans.TempItems
32:            Set lvwItems = lvwTempIPBan.ListItems
            
               'Clear out items first
35:            lvwItems.Clear
               
               Select Case m_lngBanFilter
                Case 0 'No filter
37:                    For Each objTB In colBans
38:                        If DateDiff("n", Now, objTB.ExpDate) > 0 Then
39:                            Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
40:                            X.SubItems(1) = objTB.ExpDate
41:                            X.SubItems(2) = objTB.Nick
42:                            X.SubItems(3) = objTB.BannedBy
43:                            X.SubItems(4) = objTB.Reason
44:                        Else
45:                            g_colIPBans.Remove objTB.IP, 1
46:                        End If
47:                    Next
                Case 1 'End in
48:                    strIP = txtBanFilter.Text
49:                    lngMinutes = LenB(strIP)
50:                    For Each objTB In colBans
51:                        If RightB$(objTB.IP, lngMinutes) = strIP Then
52:                            If DateDiff("n", Now, objTB.ExpDate) > 0 Then
53:                                Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
54:                                X.SubItems(1) = objTB.ExpDate
55:                                X.SubItems(2) = objTB.Nick
56:                                X.SubItems(3) = objTB.BannedBy
57:                                X.SubItems(4) = objTB.Reason
58:                            Else
59:                                g_colIPBans.Remove objTB.IP, 1
60:                            End If
61:                        End If
62:                    Next
                Case 2 'Contain
63:                    strIP = txtBanFilter.Text
64:                    For Each objTB In colBans
65:                        If InStrB(1, objTB.IP, strIP) Then
66:                            If DateDiff("n", Now, objTB.ExpDate) > 0 Then
67:                                Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
68:                                X.SubItems(1) = objTB.ExpDate
69:                                X.SubItems(2) = objTB.Nick
70:                                X.SubItems(3) = objTB.BannedBy
71:                                X.SubItems(4) = objTB.Reason
72:                            Else
73:                                g_colIPBans.Remove objTB.IP, 1
74:                            End If
75:                        End If
76:                    Next
                Case 3 'Begin with
77:                    strIP = txtBanFilter.Text
78:                    lngMinutes = LenB(strIP)
79:                    For Each objTB In colBans
80:                        If LeftB$(objTB.IP, lngMinutes) = strIP Then
81:                            If DateDiff("n", Now, objTB.ExpDate) > 0 Then
82:                                Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
83:                                X.SubItems(1) = objTB.ExpDate
84:                                X.SubItems(2) = objTB.Nick
85:                                X.SubItems(3) = objTB.BannedBy
86:                                X.SubItems(4) = objTB.Reason
87:                            Else
88:                                g_colIPBans.Remove objTB.IP, 1
89:                            End If
90:                        End If
91:                    Next
92:            End Select
        Case 5 'Clear list extract
93:            lvwTempIPBan.ListItems.Clear
94:    End Select
    
96:    Exit Sub
    
98:
Err:
99:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuTempIPBan_Click(" & Index & ")"
End Sub

Private Sub mnuRegistered_Click(Index As Integer)
1:  Dim lvwItem     As ListItem
2:  Dim lvwItems    As ListItems
3:  Dim strTxt      As String
  
5:  Set lvwItem = lvwRegistered.SelectedItem
6:  Set lvwItems = lvwRegistered.ListItems
       
8:  On Error GoTo Err
9:  If lvwItems.Count = 0 Then Exit Sub
    'Check if selected
11:    If lvwItem.Selected Then
12:        Clipboard.Clear
           Select Case Index
               Case 0 'Copy User Name
13:                Clipboard.SetText (lvwItem.Text)
               Case 1 'Copy Password
14:                Clipboard.SetText (lvwItem.SubItems(1))
               Case 2 'Copy Last IP
15:                Clipboard.SetText (lvwItem.SubItems(7))
               Case 4 'Copy All
16:                strTxt = "PT DC Hub " & vbVersion & " - Reg Name" & vbNewLine & vbNewLine & _
                            "User Name: " & lvwItem.Text & vbNewLine & _
                            "Password: " & lvwItem.SubItems(1) & vbNewLine & _
                            "Class: " & lvwItem.SubItems(2) & "=" & lvwItem.SubItems(3) & vbNewLine & _
                            "Reged By: " & lvwItem.SubItems(4) & vbNewLine & _
                            "Reg Date: " & lvwItem.SubItems(5) & vbNewLine & _
                            "Last Login: " & lvwItem.SubItems(6) & vbNewLine & _
                            "Last IP: " & lvwItem.SubItems(7)
24:                Clipboard.SetText (strTxt)
25:           End Select
26:    End If
    
28:   Exit Sub
29:
Err:
30:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuRegistered_Click(" & Index & ")"
End Sub

Private Sub mnuTray_Click(Index As Integer)
1:    On Error GoTo Err
    
      Select Case Index
        Case 0 'Show
3:            WindowState = vbNormal
4:            Show
        Case 1 'Hide
5:            WindowState = vbMinimized
7:      End Select
      
8:    Exit Sub
9:
Err:
10:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuTray_Click(" & Index & ")"
End Sub

'------------------------------------------------------------------------------
'Changed settings events
'------------------------------------------------------------------------------
Private Sub cmbData_Click(Index As Integer)
1:    On Error GoTo Err

    Select Case Index
        Case 1, 2 'Share sizes
3:            CallByName g_objSettings, cmbData(Index).Tag, VbLet, CByte(cmbData(Index).ListIndex)
        Case Else
4:            CallByName g_objSettings, cmbData(Index).Tag, VbLet, cmbData(Index).Text
5:    End Select
    
7:    Exit Sub
    
9:
Err:
10:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbData_Click(" & Index & ")"
End Sub

' ------------------------ NEW INTERFACE LANGUAGE ------------------------
Private Sub LoadLanguagesFiles()
1:    Dim strFiles()      As String
2:    Dim strFileName     As String
3:    Dim strTemp(5)      As String
4:    Dim intLoop         As Integer
5:    Dim intAux          As Integer
6:    Dim lvwItem         As Variant
7:    Dim objXML          As clsXMLParser
8:    Dim objNode         As clsXMLNode
9:    Dim objSubNode      As clsXMLNode
10:   Dim colNodes        As Collection
11:   Dim colSubNodes     As Collection
12:   On Error GoTo Err

      'Get xml files to array
15:   strFiles = g_objFileAccess.ListFiles(G_APPPATH & "\Languages\*.xml")

17:   lvwLanguages.ListItems.Clear
    
      'Loop in array..
20:   For intLoop = LBound(strFiles) To UBound(strFiles)
21:        Set objXML = New clsXMLParser

23:        strFileName = G_APPPATH & "\Languages\" & strFiles(intLoop)
        
25:        objXML.Data = g_objFileAccess.ReadFile(strFileName)
26:        objXML.Parse
        
28:        Set colNodes = objXML.Nodes(1).Nodes
           'Just in case...
30:        On Error Resume Next
31:        For Each objNode In colNodes
32:            Set colSubNodes = objNode.Nodes
33:            If objNode.Name = "Intro" Then
34:                For Each objSubNode In colSubNodes
                    Select Case objSubNode.Name
                        Case "InternationalName": strTemp(0) = objSubNode.Value
                        Case "NationalName": strTemp(1) = objSubNode.Value
                        Case "Author": strTemp(2) = objSubNode.Value
                        Case "Email": strTemp(3) = objSubNode.Value
                        Case "ReleaseDate": strTemp(4) = objSubNode.Value
35:                    End Select
36:                Next
37:            End If
38:            Exit For
39:        Next
           '
41:        For intAux = LBound(strFiles) To UBound(strFiles)
42:            If strTemp(intAux) = "" Then strTemp(intAux) = "--"
43:        Next
           '
           'Add list itens to listview..
46:        Set lvwItem = lvwLanguages.ListItems.Add(, , strTemp(0))
47:        lvwItem.SubItems(1) = strTemp(1)
48:        lvwItem.SubItems(2) = strTemp(2)
49:        lvwItem.SubItems(3) = strTemp(3)
50:        lvwItem.SubItems(4) = strTemp(4)
51:        lvwItem.Tag = strFiles(intLoop)
52:        Set lvwItem = Nothing
           '
53:    Next

55:    On Error GoTo Err
    
57:    objXML.Clear
    
59:    Set objSubNode = Nothing
60:    Set objNode = Nothing
61:    Set colSubNodes = Nothing
62:    Set colNodes = Nothing

64:   Exit Sub
65:
Err:
66:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.LoadLanguagesFiles()"
End Sub
Private Sub SetLanguageInterface(strLngFileName As String)
1:    On Error GoTo Err

3:    Dim objXML          As clsXMLParser
4:    Dim objNode         As clsXMLNode
5:    Dim objSubNode      As clsXMLNode
6:    Dim colNodes        As Collection
7:    Dim colSubNodes     As Collection
8:    Dim colAttributes   As Collection
9:    Dim strTemp         As String
10:   Dim X               As Integer

12:   Set objXML = New clsXMLParser
       
14:    Call ClearTranslations

16:    g_objSettings.Interface = strLngFileName
       
       'Set new Interface Language
19:    strTemp = G_APPPATH & "\Languages\" & g_objSettings.Interface
    
21:    If Not g_objFileAccess.FileExists(strTemp) Then
22:        g_objSettings.Interface = "English.xml"
23:        strTemp = G_APPPATH & "\Languages\" & g_objSettings.Interface
24:        If Not g_objFileAccess.FileExists(strTemp) Then
25:            If Not g_objFileAccess.FileExists(G_APPPATH & "\Languages\English.xml") Then
                   'Create defaut language file .. if is not found
27:                LoadAndSaveXML enuXML.EGLanguage
28:                LoadLanguagesFiles 'Sccan files again
29:            End If
30:        End If
31:    End If

33:    objXML.Data = g_objFileAccess.ReadFile(strTemp)
34:    objXML.Parse
    
36:    Set colNodes = objXML.Nodes(1).Nodes
    
38:    On Error Resume Next 'Just in case...
    
40:    Set g_colToolTip = New clsDictionary
41:    ReDim g_arrToolTips(0)
      
43:    For Each objNode In colNodes
44:        Set colSubNodes = objNode.Nodes
           Select Case objNode.Name
                Case "DynamicCaptions"
45:                        For Each objSubNode In colSubNodes
46:                            m_arrDynaCap(X) = objSubNode.Value
47:                            X = X + 1
48:                        Next
                Case "Captions"
49:                        For Each objSubNode In colSubNodes
50:                            TranslateCtrlCaption objSubNode.Name, objSubNode.Value
51:                        Next
                Case "Texts"
52:                        For Each objSubNode In colSubNodes
53:                            TranslateTexts objSubNode.Name, objSubNode.Value
54:                        Next
                Case "TabSCaption"
55:                        For Each objSubNode In colSubNodes
56:                            TranslateTabSCaption objSubNode.Name, objSubNode.Value
57:                        Next
                Case "ToolTips"
58:                        For Each objSubNode In colSubNodes
59:                           If objSubNode.Exists("Message", vbString) Then
60:                                ReDim Preserve g_arrToolTips(LBound(g_arrToolTips) To UBound(g_arrToolTips) + 1) As typToolTips
                            
62:                                g_arrToolTips(UBound(g_arrToolTips)).sMessage = XMLUnescape(objSubNode.Attributes("Message").Value)
    
64:                                If objSubNode.Exists("Title", vbString) Then _
                                 g_arrToolTips(UBound(g_arrToolTips)).sTitle = XMLUnescape(objSubNode.Attributes("Title").Value) _
                           Else g_arrToolTips(UBound(g_arrToolTips)).sTitle = "PTDCH - Info"
                            
68:                                If objSubNode.Exists("Style", vbInteger) Then _
                                 g_arrToolTips(UBound(g_arrToolTips)).iStyle = objSubNode.Attributes("Style").Value _
                           Else g_arrToolTips(UBound(g_arrToolTips)).iStyle = Tip_Normal
                            
72:                                If objSubNode.Exists("Icon", vbInteger) Then _
                                 g_arrToolTips(UBound(g_arrToolTips)).iIcon = objSubNode.Attributes("Icon").Value _
                           Else g_arrToolTips(UBound(g_arrToolTips)).iIcon = Tip_Info
    
76:                                g_colToolTip.Add objSubNode.Name, UBound(g_arrToolTips)
77:                           End If
78:                        Next
                Case "ListView"
79:                        For Each objSubNode In colSubNodes
80:                            TranslateListViewCaption objSubNode.Name, objSubNode.Value
81:                        Next
                Case "Captions"
82:                        For Each objSubNode In colSubNodes
83:                            TranslateCtrlCaption objSubNode.Name, objSubNode.Value
84:                        Next
                Case "TagsHelp"
85:                        For Each objSubNode In colSubNodes
86:                            m_arrTagRules(objSubNode.Name) = objSubNode.Value
87:                        Next
                Case "HubStringDef"
88:                        For Each objSubNode In colSubNodes
89:                            g_colMessages.Item(objSubNode.Name) = objSubNode.Value
90:                        Next
                Case "ToolBar"
91:                        For Each objSubNode In colSubNodes
92:                            TranslateToolBar objSubNode.Name, objSubNode.Value
93:                        Next
94:        End Select
95:     Next
    
97:     On Error GoTo Err
    
99:     objXML.Clear
    
101:    Set objSubNode = Nothing
102:    Set objNode = Nothing
103:    Set colSubNodes = Nothing
104:    Set colNodes = Nothing

106:    txtTagRules.Text = Replace(m_arrTagRules(10), "%[LF]", vbNewLine)
    
108:    If G_SERVING = False Then
109:       cmdButton(1).Caption = m_arrDynaCap(0)
110:    Else
111:       cmdButton(1).Caption = m_arrDynaCap(1)
112:    End If

114:    txtLanguages.Text = g_objSettings.Interface
    
116:    Exit Sub
117:
Err:
118:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbInterface_Click()"
End Sub
Private Sub cmdLGApply_Click()
1:   On Error GoTo Err
2:   Dim lngSelected As Long
3:   lngSelected = IsListViewSelected(lvwLanguages)
4:   If Not lngSelected = -1 Then
5:        SetLanguageInterface lvwLanguages.ListItems.Item(lngSelected).Tag
6:        Call lvwLanguages_Click
7:   End If
8:   Exit Sub
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdLGApply_Click()"
End Sub
Private Sub lvwLanguages_Click()
1:    On Error GoTo Err
2:    Dim lngSelected As Long

4:    lngSelected = IsListViewSelected(lvwLanguages)

6:    If Not lngSelected = -1 Then
7:       If lvwLanguages.ListItems.Item(lngSelected).Tag = txtLanguages.Text Then
8:            cmdLGApply.Enabled = False
9:       Else
10:           cmdLGApply.Enabled = True
11:      End If
12:   Else
13:      cmdLGApply.Enabled = False
14:   End If

16:   Exit Sub
17:
Err:
18:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwLanguages_Click()"
End Sub
' ----------------------- NEW INTERFACE LANGUAGE END ----------------------

Private Sub chkData_Click(Index As Integer)
1:    Dim objUser     As Object
2:    Dim strData     As String
3:    Dim i           As Integer
4:    Dim i2          As Integer

6:    On Error GoTo Err

8:    CallByName g_objSettings, chkData(Index).Tag, VbLet, CBool(chkData(Index).Value)
    
10:    If Index = 52 Then
11:        If chkData(52).Value Then
14:            g_objSettings.DynUpdate = True
15:        Else
18:            g_objSettings.DynUpdate = False
19:        End If
20:    End If

22:    If Index = 53 Then
23:         If chkData(53).Value Then
24:             g_objSettings.EnabledScheduler = True
25:         Else
26:             g_objSettings.EnabledScheduler = False
27:         End If
28:    End If

30:    If Index = 41 Then
31:        If chkData(41).Value Then
32:            For Each objUser In g_colUsers
33:                strData = objUser.sMyInfoString
34:                objUser.sMyInfoFakeString = "$MyINFO $ALL " & g_objRegExps.CaptureSubStr(strData, GETNICK) & " $ $$$" & g_objRegExps.CaptureDbl(strData, GETSHARESIZE) & "$"
35:            Next
36:        End If
37:    End If
       
39:    If Index = 54 Then
40:       If chkData(54).Value Then
41:          m_ObjMagnetic.AddWindow frmHub.hwnd
42:       Else
43:          m_ObjMagnetic.RemoveWindow frmHub.hwnd
44:       End If
45:    End If
        
47:    If Index = 65 Then
48:       sldPriority.Value = g_objSettings.PriorityVal
49:       If chkData(65).Value = False Then
50:          sldPriority.Enabled = False
51:          lblPriority(0).Enabled = False
52:          lblPriority(1).Enabled = False
53:          lblPriority(2).Enabled = False
54:          lblPriority(3).Enabled = False
55:          SetPriorityLivel 1
56:       Else
57:          sldPriority.Enabled = True
58:          lblPriority(0).Enabled = True
59:          lblPriority(1).Enabled = True
60:          lblPriority(2).Enabled = True
61:          lblPriority(3).Enabled = True
62:          SetPriorityLivel (g_objSettings.PriorityVal)
63:       End If
64:    End If
      
66:    If Index = 66 Then
67:          If chkData(66).Value Then
68:             g_objSettings.blSkin = True
69:             Call cmbSkin_Click
70:             cmbSkin.Enabled = True
71:             chkData(67).Enabled = True
72:             cmdSkin(0).Enabled = True
73:             cmdSkin(1).Enabled = True
74:          Else
75:             If Not g_objSettings.lngSkin = 0 Then
76:                g_objSettings.blSkin = False
77:                On Error Resume Next
                  'Refresh all picture box .. very fast
79:                For i = 0 To picTab.Count - 1: picTab(i).Cls: Next i
80:                For i = 0 To picSTab.Count - 1: picSTab(i).Cls: Next i
81:                For i = 0 To picITab.Count - 1: picITab(i).Cls: Next i
82:                For i = 0 To picTabAdv.Count - 1: picTabAdv(i).Cls: Next i
83:                For i = 0 To picHelp.Count - 1: picHelp(i).Cls: Next i
84:                For i = 0 To picBordTab.Count - 1: picBordTab(i).Cls: Next i
85:                For i = 0 To picInfo.Count - 1: picInfo(i).Cls: Next i

87:                Call Form_Paint
88:                Me.Refresh
89:             End If
90:             On Error GoTo Err
91:             cmbSkin.Enabled = False
92:             chkData(67).Enabled = False
93:             cmdSkin(0).Enabled = False
94:             cmdSkin(1).Enabled = False
95:          End If
96:    End If
       'Plugins
98:    If Index = 68 And g_PluginsFound Then
99:         Static IsLoaded As Boolean
100:        If IsLoaded Then
101:             For i = LBound(g_objPlugin) To UBound(g_objPlugin)
102:                 If g_objPlugin(i).Object.Enabled Then
103:                      If Not chkData(68).Value Then
104:                           g_objPlugin(i).Object.RunEvent "Switch", False
105:                      Else
106:                           g_objPlugin(i).Object.RunEvent "Switch", True
107:                           'Call lvwPlugins_Click 'Refresh buttons state
108:                      End If
109:                 End If
110:             Next
111:             If Not chkData(68).Value Then
                     'Desabled all plugins buttons
113:                 For i2 = 0 To cmdPlugins.Count - 1
114:                       cmdPlugins(i2).Enabled = False
115:                 Next
116:             End If
117:             IsLoaded = True
118:        End If
119:   End If
       
121:  Exit Sub
122:
Err:
123:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.chkData_Click(" & Index & ")"
End Sub
Private Sub optBanFilter_Click(Index As Integer)
1:    m_lngBanFilter = Index
End Sub

Private Sub optJM_Click(Index As Integer)
1:    g_objSettings.SendJoinMsg = Index
End Sub

Private Sub optRedirect_Click(Index As Integer)
1:    On Error GoTo Err

    'Set previous option to false
    Select Case True
        Case g_objSettings.AutoRedirect: g_objSettings.AutoRedirect = False
        Case g_objSettings.AutoRedirectFull: g_objSettings.AutoRedirectFull = False
        Case g_objSettings.AutoRedirectFullNonReg: g_objSettings.AutoRedirectFullNonReg = False
        Case g_objSettings.AutoRedirectFullNonOps: g_objSettings.AutoRedirectFullNonOps = False
        Case g_objSettings.AutoRedirectNonReg: g_objSettings.AutoRedirectNonReg = False
4:    End Select
    
    'Set correct option to true
    Select Case Index
        Case 0: g_objSettings.AutoRedirect = True
        Case 1: g_objSettings.AutoRedirectNonReg = True
        Case 2: g_objSettings.AutoRedirectFull = True
        Case 3: g_objSettings.AutoRedirectFullNonReg = True
        Case 4: g_objSettings.AutoRedirectFullNonOps = True
7:    End Select
    
9:    Exit Sub
    
11:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.optRedirect_Click(" & Index & ")"
End Sub

Private Sub sldPriority_Scroll()
'set Process Priority Class
2:    SetPriorityLivel (sldPriority.Value)
3:    g_objSettings.PriorityVal = Val(sldPriority.Value)
End Sub

Private Sub tabAdv_Click()

2: On Error GoTo Err
 
4:   Dim i, i2 As Integer

6:   i2 = Val(tabAdv.SelectedItem.Index - 1)
7:   If picTabAdv(i2).Visible = True Then Exit Sub
      
9:   For i = 0 To picTabAdv.Count - 1
10:     picTabAdv(i).Visible = False
11:   Next i
   
13:   i = Val(tabAdv.SelectedItem.Index - 1)
14:   picTabAdv(i).Refresh
15:   picTabAdv(i).Visible = True
   
17: Exit Sub
18:
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsMenu_Click()"
20:    Resume Next
End Sub

Private Sub tbsHelp_Click()
1: On Error GoTo Err
 
3:   Dim i, i2 As Integer
   
5:   i2 = Val(tbsHelp.SelectedItem.Index - 1)
6:   If picHelp(i2).Visible = True Then Exit Sub
   
8:   For i = 0 To picHelp.Count - 1
9:     picHelp(i).Visible = False
10:   Next i
   
12:   i = Val(tbsHelp.SelectedItem.Index - 1)
   
14:   picHelp(i).Visible = True

16: Exit Sub
17:
Err:
18:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsHelp_Click()"
19:    Resume Next
End Sub

Private Sub tbsInfo_Click()
1: On Error GoTo Err
 
3:   Dim i, i2 As Integer
   
5:   i2 = Val(tbsInfo.SelectedItem.Index - 1)
6:   If picInfo(i2).Visible = True Then Exit Sub
   
8:   For i = 0 To picInfo.Count - 1
9:     picInfo(i).Visible = False
10:   Next i
   
12:   i = Val(tbsInfo.SelectedItem.Index - 1)
   
14:   picInfo(i).Visible = True

16: Exit Sub
17:
Err:
18:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsInfo_Click()"
19:    Resume Next
End Sub

Private Sub tbsInteractions_Click()
1: On Error GoTo Err
 
3:   Dim i, i2 As Integer
   
5:    i2 = Val(tbsInteractions.SelectedItem.Index - 1)
6:    If picITab(i2).Visible = True Then Exit Sub
   
8:    For i = 0 To picITab.Count - 1
9:      picITab(i).Visible = False
10:   Next i
   
12:    i = Val(tbsInteractions.SelectedItem.Index - 1)
13:    picITab(i).Visible = True

15:    Exit Sub
16:
Err:
17:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsMenu_Click()"
18:    Resume Next
End Sub

Private Sub tbsMenu_Click()
 
2: On Error GoTo Err

4:    Dim i, i2, i3 As Integer
      
6:    i2 = Val(tbsMenu.SelectedItem.Index - 1)
7:    If picTab(i2).Visible = True Then Exit Sub
      
9:   For i = 0 To picTab.Count - 1
10:      picTab(i).Visible = False
11:   Next i

13:   i = Val(tbsMenu.SelectedItem.Index - 1)
14:   picTab(i).Refresh
15:   picTab(i).Visible = True

17:   If i = 5 Then Form_Resize: SCI_Focus

#If Not Status Then
20:   If i = 6 Then ' if not status then
21:       If tbsStatus.Enabled Then _
                lstStatus(0).AddItem "Status desabled..": _
                tbsStatus.Enabled = False: _
                picTab(6).Enabled = False
25:   End If
#End If

28:   If frmEditScintilla.Visible Then frmEditScintilla.Visible = False

30:   Exit Sub
31:
Err:
32:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsMenu_Click()"
33:   Resume Next
End Sub

Private Sub tbsSecurity_Click()

2: On Error GoTo Err
 
4:   Dim i, i2 As Integer
   
6:   i2 = Val(tbsSecurity.SelectedItem.Index - 1)
7:   If picSTab(i2).Visible = True Then Exit Sub
    
9:   For i = 0 To picSTab.Count - 1
10:     picSTab(i).Visible = False
11:   Next i
   
13:   i = Val(tbsSecurity.SelectedItem.Index - 1)
14:   picSTab(i).Visible = True
   
16: Exit Sub
17:
Err:
18:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsSecurity_Click()"
19:       Resume Next
End Sub

Private Sub tbsStatus_Click()
1:    On Error GoTo Err
2:    Dim i, i2 As Integer
   
4:    i2 = Val(tbsStatus.SelectedItem.Index - 1)
5:    If picStatus(i2).Visible = True Then Exit Sub
   
7:    For i = 0 To picStatus.Count - 1
8:      picStatus(i).Visible = False
9:    Next i
   
11:   i = Val(tbsStatus.SelectedItem.Index - 1)
12:   picStatus(i).Visible = True

  #If Status Then
      'Update ststistics if serving is online
16:   If i = 5 And G_SERVING Then
17:      g_objStatus.RefreshLvw iTrafic
18:      g_objStatus.RefreshLvw iProtocol
19:   End If
   #End If

22:   Exit Sub
23:
Err:
24:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsMenu_Click()"
25:  Resume Next
End Sub
Private Sub tmrBackground_Timer()
1:    Static bytCount As Byte

3:    On Error GoTo Err

5:   If g_objSettings.DynUpdate Then UpdateIPs_Timer
' Comment: should'nt we set it at a fixed 10 minutes ? UpdateIPs_Timer can be call in sub load at hub startup...
   
'***PLAN***
  ' If g_objSettings.EnabledScheduler Then Plan_Timer
'   If g_objSettings.EnabledScheduler Then TriggerCmds
' Comment: isn't a 5 minutes accuracy be enough or it must really stay 1 minute ?
'***PLAN END***

14:    If DateDiff("n", Now, m_datForceDNSUpdate) > 0 Then
15:        m_datForceDNSUpdate = Empty
16:    End If

18:    bytCount = bytCount + 1

'   If (bytCount Mod 15) = 0 Then (call the sub to If see if update is needed)
'   if protection is added against possible propagation delay, it could go in (mod 10)

    'Check if we should do anything
24:    If (bytCount Mod 10) = 0 Then
        'Remove users logging in for more than 5 minute (even if we make the check
        '                                                every 10 minutes)
27:        g_colUsers.CheckExtendedLogIn
28:        UpdateDNSs

30:        If (bytCount Mod 20) = 0 Then
            'Register the hub if needed
32:            If g_objSettings.AutoRegister Then
33:                For Each m_wskLoopItem In wskRegister
34:                    If m_wskLoopItem.State Then m_wskLoopItem.Close: DoEvents
35:                    m_wskLoopItem.Connect
36:                Next
    #If SVN Then
38:        g_objFileAccess.AppendFile G_LOGPATH, "m_wskLoopItem.Connect: " & m_wskLoopItem.RemoteHost
    #End If
40:                Set m_wskLoopItem = Nothing
41:            End If

            'If greater then 59 (ie an hour), remove all outdated
            'hammer/password guess records
45:            If bytCount > 59 Then
46:                CheckOutdatedRecords
47:                bytCount = 0
48:            End If
49:        End If

51:        Call RefreshGUI(True)

53:    End If

55:    Exit Sub
    
57:
Err:
58:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tmrBackground_Timer(bytCount = " & bytCount & ")"
End Sub

Private Sub tmrSysInfo_Timer()
1:    On Error GoTo Err

       'Close sub for not to use memory ..
4:     If Me.Visible = False Or Me.WindowState = vbMinimized Then _
         Exit Sub
           
7:     Dim iMonths As Integer, iWeeks As Integer, iDays As Integer, iHours As Integer, iMinutes As Integer, iSeconds As Integer
8:     Dim currTime As Long
9:     Dim t As String

       'Hub UpTime
12:    If G_SERVING Then
          'Get date of the server started
14:       currTime = DateDiff("s", ServingDate, DateTime.Now)
          'Calc.. date iMinutes/iSeconds/iHours/iDays/iWeeks and iMonths
16:       iSeconds = currTime Mod 60
17:       iMinutes = (currTime \ 60) Mod 60
18:       iHours = (currTime \ 3600) Mod 24
19:       iDays = (currTime \ 86400) Mod 7
20:       iWeeks = currTime \ 604800 Mod 4
21:       iMonths = (currTime \ 2419200)

23:       If iMonths > 0 Then t = "[M:" & iMonths & "["
24:       If iWeeks > 0 Then t = "[W:" & iWeeks & "["
25:       If iDays > 0 Then t = t & "[D:" & iDays & "] "

27:       t = t & StrZero(iHours, 2) & ":"
28:       t = t & StrZero(iMinutes, 2) & ":"
29:       t = t & StrZero(iSeconds, 2)

31:       txtUpTime.Text = t
32:       stbMain.Panels(1).Text = t
33:    Else
34:       t = "00:00:00"
35:       txtUpTime.Text = t
36:       stbMain.Panels(1).Text = "00:00:00"
37:    End If
    
39:  Exit Sub
40:
Err:
41:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tmrSysInfo_Timer"
End Sub

Private Sub txtData_Change(Index As Integer)
'------------------------------------------------------------------
'Purpose:   Update settings variables from text boxs
'
'Params:        Index
'               Index of the text box in the text box collection
'
'Added: Former Dev
'
'Changed:       RTD svn ?
'               TheNOP svn 26
'
'Comment:       New Cases should not be added if an actual Case can be use.
'               Just add the index number to the proper Case.
'------------------------------------------------------------------
15:  On Error GoTo Err

     Select Case Index
        Case 18 'Prefix
17:            If LenB(txtData(Index).Text) Then _
                    g_objSettings.CPrefix = AscW(txtData(18).Text) _
               Else txtData(Index).Text = ChrW$(g_objSettings.CPrefix)
        Case 19 'Long values
20:            CallByName g_objSettings, txtData(Index).Tag, VbLet, CLng(txtData(Index).Text)
        Case 20
21:            If LenB(txtData(20).Text) Then _
                    g_objSettings.CSeperator = " " _
               Else txtData(20).Text = g_objSettings.CSeperator
        Case 15, 12, 13, 14 'Byte values
24:            CallByName g_objSettings, txtData(Index).Tag, VbLet, CByte(txtData(Index).Text)
'--ROLL NEW REDIRECT TXTDATA BOXES INDEXES----------------------
        Case 5
26:        g_objSettings.ForMinShareRedirectAddress = txtData(5).Text
        Case 25
27:        g_objSettings.ForMaxShareRedirectAddress = txtData(25).Text
        Case 26
28:        g_objSettings.ForMaxSlotsRedirectAddress = txtData(26).Text
        Case 27
29:        g_objSettings.ForMinSlotsRedirectAddress = txtData(27).Text
        Case 28
30:        g_objSettings.ForTooOldNMDCRedirectAddress = txtData(28).Text
        Case 29
31:        g_objSettings.ForMaxHubsRedirectAddress = txtData(29).Text
        Case 30
32:        g_objSettings.ForNoTagRedirectAddress = txtData(30).Text
        Case 31
33:        g_objSettings.ForSlotPerHubRedirectAddress = txtData(31).Text
        Case 32
34:        g_objSettings.ForTooOldDcppRedirectAddress = txtData(32).Text
        Case 33
35:        g_objSettings.ForBWPerSlotRedirectAddress = txtData(33).Text
        Case 34
36:        g_objSettings.ForFakeTagRedirectAddress = txtData(34).Text
        Case 35
37:        g_objSettings.ForFakeShareRedirectAddress = txtData(35).Text
        Case 24
38:        g_objSettings.ForPasModeRedirectAddress = txtData(24).Text
            
'------------------AND END HERE-----------------------------------------
            'm_arrRedirectIPs = Split(g_objSettings.RedirectAddress, ";")
            'm_lngRedirectUB = UBound(m_arrRedirectIPs)
        Case 36
43:            g_objSettings.RedirectAddress = txtData(36).Text
44:            m_arrRedirectIPs = Split(g_objSettings.RedirectAddress, ";")
45:            m_lngRedirectUB = UBound(m_arrRedirectIPs)
            'If UBound = -1, then set the UBound to 0 to prevent crashes
            'otherwise set the RedirectIP to the first one
48:             If m_lngRedirectUB = -1 Then _
                     m_lngRedirectUB = 0 _
                Else g_objSettings.RedirectIP = m_arrRedirectIPs(0)
        Case 10 'Min share
51:            g_objSettings.IMinShare = CDbl(txtData(Index).Text)
52:            g_objSettings.MinShare = g_objSettings.IMinShare * (1024 ^ g_objSettings.MinShareSize)
        Case 22 'Max share
53:            g_objSettings.IMaxShare = CDbl(txtData(Index).Text)
54:            g_objSettings.MaxShare = g_objSettings.IMaxShare * (1024 ^ g_objSettings.MaxShareSize)
        Case 9, 11, 16, 17 'Double values
55:            CallByName g_objSettings, txtData(Index).Tag, VbLet, CDbl(txtData(Index).Text)

        Case 54 'DataBase Name for MySQL
57:            'If LenB(txtData(54).Text) Then _
               '     g_objSettings.DBName = txtData(54).Text _
               'Else txtData(Index).Text = g_objSettings.DBName
58:            If Not LenB(txtData(54).Text) Then txtData(Index).Text = g_objSettings.DBName
        Case Else 'Regular strings
60:            CallByName g_objSettings, txtData(Index).Tag, VbLet, txtData(Index).Text
61:    End Select
    
63:    Exit Sub
    
65:
Err:
66:    txtData(Index).Text = CallByName(g_objSettings, txtData(Index).Tag, VbGet)
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
1:    On Error GoTo Err
    
    'For numeric settings in textboxes, only allow numbers and backspace
    '(as well as decimals where required)

    Select Case Index
        Case 12, 13, 14, 15, 19 'Longs / Integers / Bytes
            Select Case KeyAscii
                Case 48 To 57, 8
                Case Else
6:                    KeyAscii = 0
7:            End Select
        Case 9, 10, 11, 16, 17, 22 'Doubles
            Select Case KeyAscii
                Case 48 To 57, 8, 46, 44
                Case Else
8:                    KeyAscii = 0
9:            End Select
10:    End Select
    
12:    Exit Sub
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.txtData_KeyPress(" & Index & ")"
End Sub

Private Sub vslData_Change(Index As Integer)
1:    On Error GoTo Err

3:    CallByName g_objSettings, vslData(Index).Tag, VbLet, vslData(Index).Value

    'Update linked label caption
    Select Case Index
        Case 0: txtVSl(0).Text = g_objSettings.DefaultBanTime
        Case 1: txtVSl(9).Text = g_objSettings.MaxUsers
        Case 2: txtVSl(10).Text = g_objSettings.MinPassiveSearchLen
        Case 3: txtVSl(2).Text = g_objSettings.FWInterval
        Case 4: txtVSl(3).Text = g_objSettings.FWBanLength
        Case 5: txtVSl(5).Text = g_objSettings.FWMyINFO
        Case 6: txtVSl(7).Text = g_objSettings.FWGetNickList
        Case 7: txtVSl(4).Text = g_objSettings.FWActiveSearch
        Case 8: txtVSl(6).Text = g_objSettings.FWPassiveSearch
        Case 9: txtVSl(1).Text = g_objSettings.MaxPassAttempts
        Case 10: txtVSl(12).Text = g_objSettings.MinSearchCls
        Case 11: txtVSl(13).Text = g_objSettings.MinConnectCls
        Case 12: txtVSl(11).Text = g_objSettings.MaxMessageLen
6:    End Select
    
8:    Exit Sub
    
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.vslData_Change(" & Index & ")"
End Sub

'------------------------------------------------------------------------------
'Winsock events
'------------------------------------------------------------------------------
Private Sub wskListen_Close(Index As Integer)
1:    On Error Resume Next
    'Ignore error, make then listen again
3:    wskListen(Index).Close
4:    DoEvents
5:    wskListen(Index).Listen
End Sub
Private Sub wskListen_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
1:   On Error Resume Next
    'Ignore error, make then listen again
3:    wskListen(Index).Close
4:    DoEvents
5:    wskListen(Index).Listen
End Sub
Private Sub wskLoop_Close(Index As Integer)
1:    Dim curUser As clsUser

3:    On Error GoTo Err

      'Make sure winsock is closed, otherwise we'll get an endless loop of Close events
6:    wskLoop(Index).Close

8:    Set curUser = g_colUsers.ItemByWinsockIndex(Index)

       'Remove them from the collection as needed
11:    If ObjPtr(curUser) Then
12:        If ObjPtr(curUser.Winsock) Then
13:            g_colUsers.Remove Index

            #If Status Then
16:                g_objStatus.URemove Index
            #End If

            'Send out quit message
20:            If curUser.State = Logged_In Then
21:                If curUser.Visible Then g_colUsers.SendToAll "$Quit " & curUser.sName & "|"
22:            End If

              'Call the sub UserQuit()
25:            SEvent_UserQuit curUser

              'Show pupop notification ..
28:           If g_objSettings.PopUpOpDisconected And curUser.Class >= 6 Then _
                    g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, g_objSettings.HubName & vbNewLine & "Op Disconected" & vbNewLine & "Nick: " & curUser.sName, 0, True
                
31:            Set curUser.Winsock = Nothing
32:        End If
33:    End If

    #If COLFREESOCKS Then
36:        On Error Resume Next
37:        m_colFreeSocks.Add wskLoop(Index), CStr(Index)
    #End If
        
40:    Exit Sub

42:
Err:
43:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.wskLoop_Close()"
   #If Status Then
45:    g_objStatus.UpDateProtocol iErrors
   #End If
47:    Resume Next
End Sub
Private Sub wskListen_ConnectionRequest(Index As Integer, ByVal requestID As Long)
1:    Dim lngTick     As Long
2:    Dim intIndex    As Integer
3:    Dim blnFull     As Boolean
4:    Dim lng         As Long
5:    Dim blnLoaded   As Boolean
6:    Dim wskUser     As Winsock

8:    Static lngDTick As Long

10:   On Error GoTo Err

      'Check if the hub is full
13:   blnFull = (g_colUsers.Count >= g_objSettings.MaxUsers)

    #If COLFREESOCKS Then
           'Check for free socket in collection
17:        If m_colFreeSocks.Count Then
18:            Set wskUser = m_colFreeSocks(1)
19:            intIndex = wskUser.Index
20:            m_colFreeSocks.Remove CStr(intIndex)
21:        Else
              'Get an unused winsock
23:            intIndex = wskLoop.UBound + 1
              'Load new winsock
25:            Load wskLoop(intIndex)
26:            Set wskUser = wskLoop(intIndex)
27:        End If
    #Else
29:        intIndex = wskLoop.UBound
           'If it's full, we're more likely to find a free winsock at the end
31:        If blnFull Then
32:            For lng = intIndex To 0
33:                If wskLoop(lng).State = 0 Then intIndex = lng: blnLoaded = True: Exit For
34:            Next
35:        Else
36:            For lng = 0 To intIndex
37:                If wskLoop(lng).State = 0 Then intIndex = lng: blnLoaded = True: Exit For
38:            Next
39:        End If
           'Load new winsock object if it never found one
41:        If Not blnLoaded Then
42:            intIndex = intIndex + 1
43:            Load wskLoop(intIndex)
44:        End If
45:        Set wskUser = wskLoop(intIndex)
     #End If
    
     'Hell the shocks is free, set interface object and run it for plugins event
     #If PreConnectionRequest Then
50:        If SEvent_PreConnectionRequest(wskUser, requestID) = True Then
              'This shocks is aborted by plugin?
              'Note: request close connection by plugin interface!
53:           Exit Sub
54:        End If
     #End If
     
    'Accept the request
58:  wskUser.Accept requestID
    
    'Check if their IP is banned
61:  requestID = g_colIPBans.Check(wskUser.RemoteHostIP)
    
    Select Case requestID
        Case 0 'Not banned
            'Check for hammering
64:            If Not UpdateConnectAttempt(wskUser, False) Then

66:                lngTick = GetTickCount
                '1 second = 1 000 milliseconds, expected default setting (< 250 ?) ~ 500 milliseconds
68:                If Abs((lngTick - lngDTick)) > g_objSettings.ConDropInterval Then
69:                    lngDTick = lngTick
70:                Else 'moved here to fix a DoS possibility...
71:                    wskUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "|"
72:                    wskUser.SendData "<" & g_objSettings.BotName & "> If you have had problems to enter here: try again in ± 10 secs or port " & Replace(g_objSettings.Ports, ";", " or ") & "|"
73:                    DoEvents
74:                    wskUser.Close
                    
                    #If COLFREESOCKS Then
77:                        m_colFreeSocks.Add wskUser, CStr(intIndex)
                    #End If
                    
80:                    Set wskUser = Nothing
                    
82:                    Exit Sub
83:                End If
 
                'Redirect as needed
86:                If g_objSettings.AutoRedirect Then
87:                    NextRedirect

89:                    wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("RedirectedTo") & g_objSettings.RedirectIP & "|$ForceMove " & g_objSettings.RedirectIP & "|"
                    
91:                    DoEvents
92:                    wskUser.Close
                    
                    #If COLFREESOCKS Then
95:                        m_colFreeSocks.Add wskUser, CStr(intIndex)
                    #End If
                    
98:                    Set wskUser = Nothing
                    
100:                    Exit Sub
101:                Else
                    'If it's full, check if we need to redirect
103:                    If blnFull Then
                        'Certain redirect types must wait till the user sends their nick
105:                        If Not g_objSettings.AutoRedirectFullNonReg Then
106:                            If Not g_objSettings.AutoRedirectFullNonOps Then
                                'Redirect as needed
108:                                If g_objSettings.AutoRedirectFull Then
109:                                    NextRedirect
110:                                    wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("FullRedirTo") & g_objSettings.RedirectIP & "|" & "$ForceMove " & g_objSettings.RedirectIP & "|"
111:                                Else
112:                                    wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("Full") & "|"
113:                                End If
                    
115:                                DoEvents
116:                                wskUser.Close
                    
                                #If COLFREESOCKS Then
119:                                    m_colFreeSocks.Add wskUser, CStr(intIndex)
                                #End If
                    
122:                                Set wskUser = Nothing
                                
124:                                Exit Sub
125:                            End If
126:                        End If
127:                    End If
                    
                    'If we get this far, the user is not connected to the hub
130:                    Set m_objLoopUser = g_colUsers.Add(intIndex)
131:                    Set m_objLoopUser.Winsock = wskUser
                    
                    #If Status Then
134:                        g_objStatus.UAdd m_objLoopUser
                    #End If
                    
                    'Send lock
138:                    wskUser.SendData "$Lock " & vbLock & "|"
    


                #If SVN Then
143:                On Error Resume Next
144:                    g_objFileAccess.AppendFile G_LOGPATH, Now & " --> " & wskUser.RemoteHostIP & " - " & "$Lock " & vbLock & "|"
145:                On Error GoTo Err
                #End If
                


                    'Call the sub AttemptedConnection(sIP)
151:                    SEvent_AttemptedConnection m_objLoopUser.IP

153:                End If
154:            End If

        Case -1 'Perm banned
            'Use descriptive ban message if needed (gives length of ban)
157:            If g_objSettings.DescriptiveBanMsg Then
158:                wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("IPPermBan") & "|"
159:            Else
160:                wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("IPBanned") & "|"
161:            End If
            
163:            DoEvents
164:            wskUser.Close
            
            #If COLFREESOCKS Then
167:                m_colFreeSocks.Add wskUser, CStr(intIndex)
            #End If
        Case Else 'Temp banned
            'Use descriptive ban message if needed (gives length of ban)
170:            If g_objSettings.DescriptiveBanMsg Then
171:                wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("IPTempBan") & MinToDate(requestID) & ".|"
172:            Else
173:                wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("IPBanned") & "|"
174:            End If

176:            DoEvents
177:            wskUser.Close
            
            #If COLFREESOCKS Then
180:                m_colFreeSocks.Add wskUser, CStr(intIndex)
            #End If
182:    End Select
    
184:    Set m_objLoopUser = Nothing

186:    Exit Sub
    
188:
Err:
189:    On Error Resume Next

191:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.wskListen_ConnectionRequest()", Err.LastDllError
192:    Set m_objLoopUser = Nothing
    
    'Make sure that if the connection was accepted it is closed
195:    If ObjPtr(wskUser) Then
196:        If wskUser.State = 7 Then
197:            wskUser.Close
198:            If g_colUsers.Exists(intIndex) Then g_colUsers.Remove intIndex
199:        End If
        
201:        Set wskUser = Nothing
202:    End If
     #If Status Then
204:    g_objStatus.UpDateProtocol iErrors
     #End If
End Sub
Private Sub wskLoop_DataArrival(Index As Integer, ByVal bytesTotal As Long)
1:    Dim lngPos            As Long
2:    Dim strIP             As String
3:    Dim strData           As String
4:    Dim strKey            As String
5:    Dim strParts          As String
6:    Dim strCommand        As String
7:    Dim curUser           As clsUser
8:    Dim strObjIP          As String
9:    Dim strTMyinfosStr    As String
10:   Dim lngTemp           As Long
11:   Dim strPreCommand     As String
12:   Dim strTemp           As String

13:   On Error GoTo Err
    
    'Prepare object / data
14:    Set curUser = g_colUsers.ItemByWinsockIndex(Index)
    
    #If OBJECTNOTSET Then
17:        If ObjPtr(curUser) = 0 Then
18:            wskLoop_Close Index
19:            Exit Sub
20:        End If
    #End If
    
23:    wskLoop(Index).GetData strData, vbString

    'Concat fragmented data if any
26:    If LenB(curUser.DataFragment) > 0 Then
27:        strData = curUser.DataFragment & strData
28:        curUser.DataFragment = vbNullString
29:    End If

    #If SVN Then
32:        If LenB(strData) Then
33:            g_objFileAccess.AppendFile G_LOGPATH, Now & " <-- " & curUser.IP & " - " & curUser.sName & " - " & strData
34:        End If
    #End If

    'Using numbers (especially longs since they are 32 bit) is much faster
    'than strings. This is an optimized approach to the DC protocol and is
    'unique to DDCH (as far as I can tell from other open source hubs)
      
    'Rather than examining the protocol as a whole, it checks the first
    'character and the length, where possible, otherwise the last character
    '(LenB is faster than AscW(RightB$(strKey, 2))
    
    #If FLASHCHAT Then
46:        If curUser.NullCharSeparator Then _
                strData = Replace(strData, vbNullChar, "|")
    #End If
    
50:    lngPos = InStrB(1, strData, "|")
    
52:    On Error GoTo LoopErr
    
    'Do while there is a | in strData
55:    Do While lngPos
        
57:        strCommand = LeftB$(strData, lngPos - 1)
58:        strData = MidB$(strData, lngPos + 2)
        
        #If PreDataArrival Then
61:        If LenB(strCommand) Then
62:               strPreCommand = SEvent_PreDataArrival(curUser, strCommand)
63:               If Not LenB(strPreCommand) = 0 Then
64:                  strCommand = strPreCommand
65:               End If
66:        End If
        #End If
        
        'Don't process command if it's empty (ignore if it's a single char)
67:        If LenB(strCommand) > 2 Then
            'Find out type of command it is
            '   -- $ = Protocol command
            '   -- < = Main chat message
            
            #If Status Then
                'Add to listbox
74:                g_objStatus.MAdd strCommand
            #End If
            
            '#If PREDATAARRIVAL Then
            '    On Error GoTo AfterPD
            '
            '    'This runs the PreDataArrival event
            '    '
            '    '  -- Parameters : curUser (the current user's clsUser object)
            '    '                : strData (data that was sent)
            '    '  -- Format     : Function PreDataArrival(curUser, strData)
            '    '
            '    '  -- Called when a user sends data to the hub, but before the hub parses
            '    '     it
            '    '  -- It should return the string it should parse
            '
            '    If m_intPDIndex Then strCommand = ScriptControl(m_intPDIndex).Run("PreDataArrival", curUser, strCommand)
            '    If Not LenB(strCommand) > 2 Then GoTo NextLoop
            '
93:
AfterPD:    '
            '    On Error GoTo Err
            '#End If
            
            Select Case AscW(strCommand)
                Case 36 '$
                    'Check if there is a " "; if there is remove the key from
                    'data and seperate it's params
99:                    lngPos = InStrB(1, strCommand, " ")
                    
101:                    If lngPos Then
102:                        strKey = MidB$(strCommand, 3, lngPos - 3)
103:                        strParts = MidB$(strCommand, lngPos + 2)
104:                    Else
105:                        strKey = MidB$(strCommand, 3)
106:                    End If
                    
                    'Start parsing!
                    
                    'Notes -- This is structured in such a way so that the most
                    '         common messages are at the top of the Select Case;
                    '         this makes it even more efficent =)
                    '
                    '      -- Also due to this format, DC protocol commands ARE
                    '         CASE SENSITIVE; I will NEVER add support for bots/
                    '         etc which do not follow the protocol properly
                    '
                    '      -- In case some of you are wondering, this is a quite
                    '         inaccurate way of parsing messages...if DDCH does
                    '         not support a message, it might think it's another
                    '         unrelated message. For that reason, I support
                    '         all documented protocol extensions (at least to the
                    '         point of making sure it doesn't get confused)
                    
                    Select Case AscW(strKey)
                        Case 83 'S
                            'Possible messages :
                            '   -- Search
                            '   -- SR (passive search result; active is sent via UDP directly to the client)
                            '   -- Supports
                            
                            Select Case LenB(strKey)
                                Case 12
                                    'Search
                                    '
                                    '   -- Format   : $Search [<ip:port>//Hub:<name>] <T/F>?<T/F>?<size>?<type>?<search>|
                                    '   -- Response : N/A (clients will respond with $SR)
                                    '
                                    '   -- Standard protocol message
                                    '   -- It has either an ip and port or a name (active versus passive).
                                    '      First T/F toggles whether or not the size of files is restricted.
                                    '      Second T/F toggles whether limit is upper (T) or lower (F)
                                    '      <size> is the size in bytes of the file
                                    '      <type> is the type of file it can be (document, audio, etc)
                                    '      <search> is the string to search for (spaces are converted to $)
                                    
                                    'If the user has not sent their MyINFO string, check if bots are allowed
                                    'The second check is 95-99% effective, as I've yet to see a search tool
                                    'which sends GetNickList (you must request it to search)
                                
147:                                    If g_objSettings.PreventSearchBots Then
148:                                        If Not curUser.State = Logged_In Then wskLoop_Close Index: Exit Sub
149:                                        If Not curUser.QNL Then wskLoop_Close Index: Exit Sub
150:                                    End If
                                    

153:                                    If Not g_objSettings.ChatOnly Then
154:                                        If curUser.ChatOnly Then curUser.Kick g_objSettings.DefaultBanTime, "PTDCH / Core", "ChatOnly"
                                    
156:                                        If curUser.Class >= g_objSettings.MinSearchCls Then
157:                                            If AscW(strParts) = 72 Then
                                                'Make sure the user isn't flooding
159:                                                If g_objSettings.EnableFloodWall Then _
                                                        If curUser.FloodCheck(1) Then _
                                                        Exit Sub
                                                    
                                                'Set to passive
164:                                                curUser.Passive = True
                                            
                                                'Allow search is passive searches are not disabled
                                                Select Case g_objSettings.MinPassiveSearchLen
                                                    Case -1 'Disabled
                                                    Case 0, Is <= Len(Trim$(Mid$(strParts, InStrRev(strParts, "?") + 1)))
                                                        'Make sure they aren't faking their name
168:                                                        If curUser.sName = MidB$(strParts, 9, InStrB(1, strParts, " ") - 9) Then
169:                                                            lngPos = ObjPtr(curUser)
                                                        
171:                                                            If g_objSettings.MinClsSearchSend Then
                                                                'Don't send search to person who sent it
173:                                                                For Each m_objLoopUser In g_colUsers
174:                                                                    If Not m_objLoopUser.Passive Then _
                                                                            If m_objLoopUser.Visible Then _
                                                                            If Not ObjPtr(m_objLoopUser) = lngPos Then _
                                                                                m_objLoopUser.SendData strCommand & "|"
178:                                                                Next
179:                                                            Else
                                                                'Don't send search to person who sent it
181:                                                                For Each m_objLoopUser In g_colUsers
182:                                                                    If Not m_objLoopUser.Passive Then _
                                                                            If m_objLoopUser.Class >= g_objSettings.MinSearchCls Then _
                                                                            If m_objLoopUser.Visible Then _
                                                                                If Not ObjPtr(m_objLoopUser) = lngPos Then _
                                                                                    m_objLoopUser.SendData strCommand & "|"
187:                                                                Next
188:                                                            End If
                                                        
190:                                                            Set m_objLoopUser = Nothing
191:                                                        Else
192:                                                            wskLoop_Close Index
193:                                                            Exit Sub
194:                                                        End If
195:                                                End Select
196:                                            Else
                                                'Make sure they aren't flooding
198:                                                If g_objSettings.EnableFloodWall Then _
                                                    If curUser.FloodCheck(0) Then _
                                                        Exit Sub
                                                        
                                                'Set to active
203:                                                curUser.Passive = False
                                                
205:                                                strKey = curUser.IP
                                                
                                                'Find out if the IP is a local range (skip IP match check if it is)
                                                Select Case CByte(LeftB$(strKey, InStrB(1, strKey, ".") - 1))
                                                    Case 192, 127, 10
                                                    Case Else
208:                                                        strParts = LeftB$(strParts, InStrB(1, strParts, ":") - 1)
                                                        
                                                        'If IP doesn't match, then fix it
211:                                                        If Not strKey = strParts Then _
                                                                strCommand = Replace(strCommand, strParts, strKey, 1, 1)
213:                                                End Select
                                                
215:                                                lngPos = ObjPtr(curUser)
                                                        
217:                                                If g_objSettings.MinClsSearchSend Then
                                                    'Don't send search to person who sent it
219:                                                    For Each m_objLoopUser In g_colUsers
220:                                                        If m_objLoopUser.Visible Then _
                                                            If Not ObjPtr(m_objLoopUser) = lngPos Then _
                                                                m_objLoopUser.SendData strCommand & "|"
223:                                                    Next
224:                                                Else
                                                    'Don't send search to person who sent it
226:                                                    For Each m_objLoopUser In g_colUsers
227:                                                        If m_objLoopUser.Class >= g_objSettings.MinSearchCls Then _
                                                                If m_objLoopUser.Visible Then _
                                                                If Not ObjPtr(m_objLoopUser) = lngPos Then _
                                                                    m_objLoopUser.SendData strCommand & "|"
231:                                                    Next
232:                                                End If
                                                    
234:                                                Set m_objLoopUser = Nothing
235:                                            End If
                                            #If Status Then
237:                                            g_objStatus.UpDateTrafic bytesTotal, iSearchs
                                            #End If
239:                                        End If
240:                                    End If
                                Case 4
                                    'SR
                                    '
                                    '   -- Format   : $SR <from> <fpath><char5><fsize> <fslots>/<tslots><char5><hubname> (<hubip>[:<hubport>])<char5><to>|
                                    '               : $SR <from> <directory> <fslots>/<tslots><char5><hubname> (<hubip>[:<hubport>])<char5><to>|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message
                                    '   -- Result for passive searches
                                    '   -- Two forms; first is for files, second is for directories
                                    
                                    'Make sure we even need to bother with the check
                                    Select Case False
                                        Case g_objSettings.OPBypass, (curUser.Class >= Vip)
252:                                            lngPos = InStrB(1, strParts, " ")
                                            
                                            'Make sure their nickname matches the search result name
255:                                            If LeftB$(strParts, lngPos - 1) = curUser.sName Then
256:                                                strParts = MidB$(strParts, lngPos + 2)
                                            
258:                                                lngPos = InStrB(1, strParts, "/") + 2
259:                                                strParts = MidB$(strParts, lngPos, InStrB(lngPos, strParts, vbChar5) - lngPos)
                                            
                                                'If the total slots value isn't numerical, kick
262:                                                If IsNumeric(strParts) Then
263:                                                    lngPos = GetByte(CLng(strParts))
                                                    
                                                    'Check min slots
266:                                                    If g_objSettings.MinSlots Then
267:                                                        If lngPos < g_objSettings.MinSlots Then
268:                                                            FailedConf curUser, MinSlots
269:                                                            Exit Sub
270:                                                        End If
271:                                                    End If
                                                   
                                                    
                                                    'Check max slots
275:                                                    If g_objSettings.MaxSlots Then
276:                                                        If lngPos > g_objSettings.MaxSlots Then
277:                                                            FailedConf curUser, MaxSlots
278:                                                            Exit Sub
279:                                                        End If
280:                                                    End If
281:                                                Else
282:                                                    wskLoop_Close Index
283:                                                    Exit Sub
284:                                                End If
285:                                            Else
286:                                                curUser.Kick 60, "PTDCH / Core", "Parameter not valid when processing protolcol"
287:                                                Exit Sub
288:                                            End If
                                            
                                            'checking slot here is risky.(MyINFO are delayed a bit client side in order not to spam hubs.)
                                            'If it finds two vbChar5, then it is a directory, else a file
                                            'If (LenB(strParts) - LenB(Replace(strParts, vbChar5, vbNullString))) = 4 Then
                                            '    strKey = LeftB$(strParts, InStrB(1, strParts, vbChar5) - 1)
                                            '    lngPos = CLng(MidB$(strKey, LenB(strKey) - InStrB(1, StrReverse(strKey), "/") + 2))
                                            'Else
                                            '    strKey = MidB$(strParts, InStrB(InStrB(1, strParts, vbChar5), strParts, " "))
                                            '    lngPos = InStrB(1, strKey, "/") + 2
                                            '    lngPos = CLng(MidB$(strKey, lngPos, InStrB(lngPos, strKey, vbChar5) - lngPos))
                                            'End If
                                           '
                                           ' 'Check for fake slots
                                           ' If curUser.Slots Then _
                                           '     If Not curUser.Slots = lngPos Then _
                                           '         FailedConf curUser, FakeTag: Exit Sub

306:                                    End Select
                                    
                                    'Find out who the result should be sent to
309:                                    lngPos = InStrRev(strCommand, vbChar5)
310:                                    strParts = Mid$(strCommand, lngPos + 1)

                                    'If online, send result to client
313:                                    If g_colUsers.Online(strParts) Then
314:                                            g_colUsers.ItemByName(strParts).SendData Left$(strCommand, lngPos - 1) & "|"
                                        #If Status Then
316:                                            g_objStatus.UpDateTrafic bytesTotal, iSearchs
                                        #End If
318:                                        End If
                                Case 16
                                    'Supports
                                    '
                                    '   -- Format   : $Supports <ext> <ext_etc>|
                                    '   -- Response : $Supports <etc> <ext_etc>|
                                    '
                                    '   -- Protocol extension (in response to EXTENDEDPROTOCOL in Lock string)
                                    '   -- Allows client to extend abilities
                                    '   -- Only extensions which both the client and hub support
                                    '      should be sent back to the client
                                    
329:                                    strParts = strParts & " "
330:                                    lngPos = InStrB(1, strParts, " ")
                                    
332:                                    curUser.ZLine = False
333:                                    curUser.ZPipe = False
334:                                    curUser.QuickList = False
335:                                    curUser.NoHello = False
336:                                    curUser.UserCommand = False
337:                                    curUser.ChatOnly = False

                                    #If FLASHCHAT Then
340:                                        curUser.NullCharSeparator = False
                                    #End If
                                    
                                    'Find out which extensions both support
344:                                    Do While lngPos
345:                                        strKey = LeftB$(strParts, lngPos - 1)
346:                                        strParts = MidB$(strParts, lngPos + 2)

                                        Select Case strKey
                                            Case "QuickList"
348:                                                curUser.Supports = curUser.Supports & " QuickList"
349:                                                curUser.QuickList = True
350:                                                curUser.NoHello = True
                                            Case "UserCommand"
351:                                                curUser.Supports = curUser.Supports & " UserCommand"
352:                                                curUser.UserCommand = True
                                            Case "NoHello"
353:                                                curUser.Supports = curUser.Supports & " NoHello"
354:                                                curUser.NoHello = True
                                            Case "NoGetINFO", "TTHSearch", "UserIP2", "UserIP", "xKick", "BotINFO"
355:                                                curUser.Supports = curUser.Supports & " " & strKey
                                            Case "ZPipe"
356:                                                curUser.Supports = curUser.Supports & " ZPipe"
357:                                                curUser.ZPipe = True
                                            Case "ZLine"
358:                                                curUser.Supports = curUser.Supports & " ZLine"
359:                                                curUser.ZLine = True
                                            Case "ChatOnly"
360:                                                curUser.Supports = curUser.Supports & " ChatOnly"
361:                                                curUser.ChatOnly = True

                                        #If FLASHCHAT Then
                                            Case "NullCharSeparator"
364:                                                curUser.Supports = curUser.Supports & " NullCharSeparator"
365:                                                curUser.NullCharSeparator = True
                                        #End If

368:                                        End Select

370:                                        lngPos = InStrB(1, strParts, " ")
371:                                    Loop

                                    'Remote leading space
374:                                    strKey = LTrim$(curUser.Supports)
375:                                    curUser.Supports = strKey
                                    
                                    'If it supports UserCommand, then send the clear command
378:                                    If curUser.UserCommand Then _
                                        curUser.SendData "$Supports " & strKey & "|$UserCommand 255 7 |" _
                                    Else _
                                        curUser.SendData "$Supports " & strKey & "|"
382:                            End Select
                        Case 77 'M
                            'Possible messages :
                            '   -- MyINFO
                            '   -- MyPass
                            '   -- MultiConnectToMe (ignored)
                            '   -- MultiSearch (ignored)
                            '   -- MyIP (ignored)
                            
                            Select Case AscW(RightB$(strKey, 2))
                                Case 79
                                    'MyINFO
                                    '
                                    '   -- Format   : $MyINFO $ALL <name> <description>$ $<connection><char>$<email>$<share>$|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message
                                    '   -- Contains all the info of a user
                                    '   -- Once this is sent, and assuming it passes the rules,
                                    '      the client has logged in
                                    
                                    'See if we can limit MyINFOs sending when Hide MyINFOs is enabled.
                                    'only send to registered users if it has changed.
                                    
                                    'Check if the user is flooding
404:                                    If g_objSettings.EnableFloodWall Then _
                                            If curUser.FloodCheck(2) Then _
                                            Exit Sub

408:                                    If curUser.State = Logged_In Then
                                        'If the MyINFO string has changed, then continue
410:                                        If Not curUser.sMyInfoString = strCommand Then
                                            'If the name matches, process it
412:                                            If MidB$(strParts, 11, InStrB(13, strParts, " ") - 11) = curUser.sName Then
                                                'If it passes the rules, then send it out to all users
414:                                                If ProcessMyINFO(curUser, strParts) Then
                                                    'But only if the user is Visible
416:                                                    If curUser.Visible Then
417:                                                        For Each m_objLoopUser In g_colUsers
418:                                                            If curUser.State = Disconnected Then Exit Sub
                                                            '#If FLASHCHAT Then
420:                                                                If Not m_objLoopUser.ChatOnly Then
421:                                                                    If g_objSettings.HideMyinfos Then
422:                                                                        If m_objLoopUser.Class < g_objSettings.MinMyinfoFakeCls Then
423:                                                                            m_objLoopUser.SendData curUser.sMyInfoFakeString & "|"
424:                                                                        Else
425:                                                                            m_objLoopUser.SendData strCommand & "|"
426:                                                                        End If
427:                                                                    Else
428:                                                                        m_objLoopUser.SendData strCommand & "|"
429:                                                                    End If
430:                                                                End If
                                                            '#Else
                                                            '    If g_objSettings.HideMyinfos Then
                                                            '        If m_objLoopUser.Class < g_objSettings.MinMyinfoFakeCls Then
                                                            '            m_objLoopUser.SendData curUser.sMyInfoFakeString & "|"
                                                            '        Else
                                                            '            m_objLoopUser.SendData strCommand & "|"
                                                            '        End If
                                                            '    Else
                                                            '        m_objLoopUser.SendData strCommand & "|"
                                                            '    End If
                                                            '#End If
442:                                                        Next
                                                        
444:                                                        Set m_objLoopUser = Nothing
445:                                                    Else
446:                                                        If g_objSettings.HideMyinfos Then
447:                                                            If curUser.Class < g_objSettings.MinMyinfoFakeCls Then
448:                                                                curUser.SendData curUser.sMyInfoFakeString & "|"
449:                                                            Else
450:                                                                curUser.SendData strCommand & "|"
451:                                                            End If
452:                                                        Else
453:                                                            curUser.SendData strCommand & "|"
454:                                                        End If
455:                                                    End If
456:                                                Else
                        'fail hub's rule(s)
458:                                                    Exit Sub
459:                                                End If
460:                                            Else
                        'not same nick in his myinfo
462:                                                wskLoop_Close Index
463:                                                Exit Sub
464:                                            End If
465:                                        End If
466:                                    Else
                                        'If they are not logged in, and support QuickList, they
                                        'must be logging in; if not, then the handshake is almost done
469:                                        If curUser.QuickList Then
                                            'Discontinue processing if they fail nick validation
471:                                            If Not ValidateNick(curUser, MidB$(strParts, 11, InStrB(13, strParts, " ") - 11), strParts) Then _
                                                Exit Sub
473:                                        Else
                                            'Make sure they aren't faking the name
475:                                            If MidB$(strParts, 11, InStrB(13, strParts, " ") - 11) = curUser.sName Then
                                                'Should be waiting for the MyINFO string
477:                                                If curUser.State = Wait_Info Then
                                                    'Check to see if it passes the rules
479:                                                    If ProcessMyINFO(curUser, strParts) Then
480:                                                        g_colUsers.UpdateLogIn curUser
                                                        
                                                        'Call the right sub
                                                        Select Case curUser.Class
                                                            Case Normal: SEvent_UserConnected curUser
                                                            Case Mentored, Registered, Invisible, Vip: SEvent_RegConnected curUser
                                                            Case Else
483:                                                            SEvent_OpConnected curUser
                                                                'Show pupop notification ..
485:                                                            If g_objSettings.PopUpOpConected Then _
                                                                        g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, g_objSettings.HubName & vbNewLine & "Op Connected" & vbNewLine & "Nick: " & curUser.sName, 0, True
487:                                                        End Select
488:                                                    Else
                        'fail hub's rule(s)
490:                                                        Exit Sub
491:                                                    End If
492:                                                Else
                        'wrong handshake
494:                                                    wskLoop_Close Index
495:                                                    Exit Sub
496:                                                End If
497:                                            Else
                        'fail nick validation
499:                                                wskLoop_Close Index
500:                                                Exit Sub
501:                                            End If
                                            #If Status Then
503:                                            g_objStatus.UpDateTrafic bytesTotal, iMyINFOs
                                            #End If
505:                                        End If
506:                                    End If
                                Case 115
                                    'MyPass
                                    '
                                    '   -- Format   : $MyPass <password>|
                                    '   -- Response : $BadPass|  //  $LogedIn <name>|
                                    '
                                    '   -- Standard protocol message (registered users only)
                                    '   -- If the user is registered they send the password for their
                                    '      account. If it's wrong, send $BadPass, otherwise send
                                    '      $LogedIn
                                    
                                    'Make sure the user is supposed to send a password
                                    Select Case curUser.State
                                        Case Wait_Pass
                                            'User is registered
                                            
520:                                            strKey = curUser.sName
                                    
                                            'Check if password is correct
523:                                            lngPos = g_objRegistered.Check(strKey, strParts)
                                            
                                            'If it's a nonzero value, the password is correct
526:                                            If lngPos Then
527:                                                If g_objSettings.PreventGuessPass Then UpdateFailedReg curUser, True

529:                                                curUser.Class = lngPos
                                        
                                                'If there is a logged user already with the same name
                                                'disconnect them
533:                                                If g_colUsers.Online(strKey) = -1 Then _
                                                        wskLoop_Close g_colUsers.ItemByName(strKey).iWinsockIndex
                                        
                                                'We need their MyINFO string before logging them in,
                                                'so unless they are using QuickList, don't log them in
538:                                                If curUser.QuickList Then
539:                                                    If lngPos > Vip Then
540:                                                        curUser.SendData "$LogedIn " & strKey & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & ".|"
541:                                                    Else
542:                                                        curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & ".|"
543:                                                    End If
                                                    
                                                    'Check MyINFO
546:                                                    If ProcessMyINFO(curUser, curUser.sMyInfoString) Then
                                                        'Set this to true, because QuickList clients don't send $GetNickList
548:                                                        curUser.QNL = True
                                                
550:                                                        g_colUsers.UpdateLogIn curUser
                                            
                                                        'Raise script event
553:                                                        If lngPos > Vip Then
554:                                                            SEvent_OpConnected curUser
                                                                'Show pupop notification ..
556:                                                            If g_objSettings.PopUpOpConected Then _
                                                                        g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, g_objSettings.HubName & vbNewLine & "Op Connected" & vbNewLine & "Nick: " & curUser.sName, 0, True
558:                                                        Else
559:                                                            SEvent_RegConnected curUser
560:                                                        End If
561:                                                    Else
562:                                                        Exit Sub
563:                                                    End If
564:                                                Else
565:                                                    If lngPos > Vip Then
566:                                                        curUser.SendData "$Hello " & strKey & "|$LogedIn " & strKey & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & ".|"
567:                                                    Else
568:                                                        curUser.SendData "$Hello " & strKey & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & ".|"
569:                                                    End If
570:                                                    curUser.State = Wait_Info
571:                                                End If
                                        
                                                'Update log in status in database
574:                                                m_objPermaCon.Execute "UPDATE UsrDynamic Set LastLogin='" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "', LastIP='" & curUser.IP & "' WHERE UserName=" & SQLQuotes(curUser.sName), , 129
575:                                            Else
576:                                                If g_objSettings.PreventGuessPass Then UpdateFailedReg curUser, False

578:                                                curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("WrongPass") & "|$BadPass|"

580:                                                DoEvents
581:                                                wskLoop_Close Index
                                                
583:                                                Exit Sub
584:                                            End If
                                        Case Wait_PassPM
                                            'Not registered, but the hub is running in PM mode
                                            
                                            'Make sure the password is correct
588:                                            If strParts = g_objSettings.HubPassword Then
589:                                                curUser.Class = Normal

591:                                                curUser.SendData "$Hello " & curUser.sName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & "|"
                                                
                                                'We need their MyINFO string before logging them in,
                                                'so unless they are using QuickList, don't log them in
595:                                                If curUser.QuickList Then
596:                                                    g_colUsers.UpdateLogIn curUser
                                            
                                                    'Raise script event
599:                                                    SEvent_UserConnected curUser
600:                                                Else
601:                                                    curUser.State = Wait_Info
602:                                                End If
603:                                            Else
                                                'Send redirect request or just tell them they got it wrong
605:                                                If g_objSettings.RedirectFGP Then
606:                                                    NextRedirect
607:                                                    curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("WrongPassRedir") & g_objSettings.RedirectIP & "|$BadPass|$ForceMove " & g_objSettings.RedirectIP & "|"
608:                                                Else
609:                                                    curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("WrongPass") & "|$BadPass|"
610:                                                End If
                                                
612:                                                DoEvents
613:                                                wskLoop_Close Index
                                                
615:                                                Exit Sub
616:                                            End If
617:                                    End Select
                                'Case 101
                                '    'MultiConnectToMe
                                'Case 104
                                '    'MultiSearch
                                'Case 80
                                '    'MyIP
                                '
                                '    curUser.SendData "$YourIP " & curUser.IP & "|"
626:                            End Select
                        Case 67 'C
                            'Possible messages :
                            '   -- ConnectToMe
                            '   -- ClientID (ignored)
                            
631:                            If LenB(strKey) = 22 Then
                                'ConnectToMe
                                '
                                '   -- Format   : $ConnectToMe <name> <ip>:<port>|
                                '   -- Response : N/A
                                '
                                '   -- Standard protocol message
                                '   -- Active users send this for <name> to connect to their IP
                                '      on the specified port to intiate a file transfer connection
                                
641:                                If Not g_objSettings.ChatOnly Then
                      
643:                                    If curUser.State = Logged_In Then
                                            'They know why ; )
645:                                        If curUser.ChatOnly Then curUser.Kick g_objSettings.DefaultBanTime, "PTDCH / Core", "ChatOnly": Exit Sub

647:                                        curUser.Passive = False
                                    
                                        'If using the mentoring system, then we may have to disconnect the user
650:                                        If g_objSettings.MentoringSystem Then
                                            'Do not check users who are being mentored
                                            'or are an VIP/Op
653:                                            lngPos = curUser.Class
                                            
                                            Select Case curUser.Class
                                                Case 2, Is < Vip
                                                    'Make the min share check
656:                                                    If curUser.iBytesShared < g_objSettings.MinShare Then _
                                                            FailedConf curUser, MinShare: Exit Sub
658:                                            End Select
659:                                        End If
                                        
661:                                        lngPos = InStrB(1, strParts, " ")
662:                                        strKey = LeftB$(strParts, lngPos - 1)
                                        
                                        Select Case True
                                            Case QueuedConnect(strKey & "|" & curUser.sName), curUser.Class >= g_objSettings.MinConnectCls
                                                'Make the MLDC check if needed
665:                                                If g_objSettings.AutoKickMLDC Then
666:                                                    strParts = MidB$(strParts, lngPos + 2)
667:                                                    lngPos = InStrB(1, strParts, ":")
668:                                                    strIP = MidB$(strParts, lngPos + 2)
                                                    
                                                    'Make sure the port is numeric
671:                                                    If IsNumeric(strIP) Then
                                                        'If the client is listening on port 4444 for connections
                                                        'they are a MLDC client
674:                                                        If CLng(strIP) = 4444 Then
675:                                                            curUser.Kick 60, "PTDCH / Core", "The client is listening on port 4444 for connections they are a MLDC client"
676:                                                            Exit Sub
677:                                                        Else
678:                                                            strParts = LeftB$(strParts, lngPos - 1)
679:                                                        End If
680:                                                    Else
681:                                                        curUser.Kick 60, "PTDCH / Core", "The port listening is not numeric"
682:                                                        Exit Sub
683:                                                    End If
684:                                                Else
685:                                                    lngPos = lngPos + 2
686:                                                    strParts = MidB$(strParts, lngPos, InStrB(lngPos, strParts, ":") - lngPos)
687:                                                End If
                                                
689:                                                strIP = curUser.IP
                                                
                                                'Find out if the IP is a local range (replace IP if needed)
                                                ' Range 1: Class A - 10.0.0.0 through 10.255.255.255
                                                ' Range 2: Class B - 172.16.0.0 through 172.31.255.255
                                                ' Range 3: Class C - 192.168.0.0 through 192.168.255.255
                                                Select Case CByte(LeftB$(strIP, InStrB(1, strIP, ".") - 1))
                                                    Case 192, 10, 172
                                                        '$ConnectToMe TheNOP_log 64.228.81.77:3340
                                                        '$ConnectToMe <strKey> <strParts>:<Port>
697:                                                        If g_colUsers.Online(strKey) Then
698:                                                            strObjIP = g_colUsers.ItemByName(strKey).IP
                                                            'Find out if the IP that he want to connect to is also a LAN IP
                                                            Select Case CByte(LeftB$(strObjIP, InStrB(1, strObjIP, ".") - 1))
                                                                Case 192, 10, 172
                                                                    'is also LAN, If IP is not the same, then fix it
701:                                                                    If Not strIP = strParts Then
702:                                                                        strCommand = Replace(strCommand, strParts, strIP, 1, 1)
703:                                                                    End If
704:                                                            End Select
                                                        'Else
                                                            'attempt to get ride of possible ghosts
                                                            'curUser.SendData "$Quit " & strKey & "|"
708:                                                        End If
                                                    Case 127

                                                    Case Else
                                                        'If IP is not the same, then fix it
711:                                                        If Not strParts = strIP Then _
                                                                strCommand = Replace(strCommand, strParts, strIP, 1, 1)
713:                                                End Select
                            
715:                                                If g_colUsers.Online(strKey) Then
                            
                                                '#If FLASHCHAT Then
718:                                                    If g_colUsers.ItemByName(strKey).ChatOnly Then
719:                                                        curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("ChatMode"), "%[user]", strKey)
720:                                                    Else
                                                '#End If
                                                
723:                                                        If g_objSettings.MinClsConnectSend Then
724:                                                            g_colUsers.ItemByName(strKey).SendData strCommand & "|"
725:                                                        Else
726:                                                            Set m_objLoopUser = g_colUsers.ItemByName(strKey)
                                                    
728:                                                            If m_objLoopUser.Class >= g_objSettings.MinConnectCls Then
729:                                                                g_colUsers.ItemByName(strKey).SendData strCommand & "|"
730:                                                            End If

732:                                                            Set m_objLoopUser = Nothing
733:                                                        End If
                                                
                                                '#If FLASHCHAT Then
736:                                                    End If
                                                '#End If
                                                'Else
                                                    'attempt to get ride of a possible ghost
                                                    'curUser.SendData "$Quit " & strKey & "|"
741:                                                End If
742:                                        End Select
743:                                    End If
                                    #If Status Then
745:                                    g_objStatus.UpDateProtocol iConnectMe
                                    #End If
747:                                End If
748:                            End If
                        Case 82 'R
                            'Possible messages :
                            '   -- RevConnectToMe
                            
                            'RevConnectToMe
                            '
                            '   -- Format   : $RevConnectToMe <name> <othername>|
                            '   -- Response : $ConnectToMe <name> <otherip>| (from client)
                            '
                            '   -- Standard protocol message (passive clients)
                            '   -- For passive mode connecting; name is the passive client
                            '      and othername is the client it wants to connect to.
                            '      Assuming the other client is active, it will respond with
                            '      a ConnectToMe message so that the passive user can connect to
                            '      them

764:                            If Not g_objSettings.ChatOnly Then
                                'The user must be logged in to connect to other uses
766:                                If curUser.State = Logged_In Then
                                        'They know why ; )
768:                                    If curUser.ChatOnly Then curUser.Kick g_objSettings.DefaultBanTime, "PTDCH / Core", "ChatOnly": Exit Sub

770:                                    If curUser.Class >= g_objSettings.MinConnectCls Then
                                    
                                        'add a Myinfo check here, to see if tag is really showing passive...
773:                                        curUser.Passive = True
                                    
775:                                        lngPos = InStrB(1, strParts, " ")
                                        
                                        'Make sure the user isn't faking their name
778:                                        If LeftB$(strParts, lngPos - 1) = curUser.sName Then
779:                                            strKey = MidB$(strParts, lngPos + 2)
                                            
                                            'If user is online, forward the message to them
782:                                            If g_colUsers.Online(strKey) Then
783:                                                If g_colUsers.ItemByName(strKey).ChatOnly Then
784:                                                        curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("ChatMode"), "%[user]", strKey)
785:                                                Else
786:                                                    If g_objSettings.MinClsConnectSend Then
787:                                                        If g_objSettings.MinConnectCls Then
788:                                                            g_colUsers.ItemByName(strKey).SendData strCommand & "|"
                                                            
                                                            'On Error GoTo NextLoop
791:                        On Error Resume Next
                                                            
793:                                                            m_colRevConnects.Add Now, curUser.sName & "|" & strKey

795:                                                        If Err.Number = 457 Then
796:                            Err.Clear
797:                            On Error GoTo Err
798:                            GoTo NextLoop
799:                        End If

801:                                                        End If
802:                                                    Else
803:                                                        Set m_objLoopUser = g_colUsers.ItemByName(strKey)
                                                    
805:                                                        If m_objLoopUser.Class >= g_objSettings.MinConnectCls Then
806:                                                            If g_objSettings.MinConnectCls Then
807:                                                                m_objLoopUser.SendData strData & "|"
                                                                
                                                                'On Error GoTo NextLoop
810:                                  On Error Resume Next

812:                                                                m_colRevConnects.Add Now, curUser.sName & "|" & strKey

814:                                                        If Err.Number = 457 Then
815:                            Err.Clear
816:                            On Error GoTo Err
817:                            GoTo NextLoop
818:                        End If

820:                                                            End If
821:                                                        End If
                                                    
823:                                                        Set m_objLoopUser = Nothing
824:                                                    End If
825:                                                End If
                                            'Else
                                                'attempt to get ride of possible ghosts
                                                'curUser.SendData "$Quit " & strKey & "|"
829:                                            End If
830:                                        End If
831:                                    End If
                                    #If Status Then
833:                                    g_objStatus.UpDateProtocol iRevConnectToMe
                                    #End If
835:                                End If
836:                            End If
                        Case 71 'G
                            'Possible messages :
                            '   -- GetNickList
                            '   -- GetInfo (ignored)
                            
841:                            If LenB(strKey) = 22 Then
                                'GetNickList
                                '
                                '   -- Format   : $GetNickList|
                                '   -- Response : $NickList <name>$$[<name_etc>$$]|$OpList <name>$$[<name_etc>$$]|
                                '
                                '   -- Standard protocl message
                                '   -- Retrieves the list of users / ops connected to the hub
                                '   -- Traditionally, a GetINFO should be sent to get the MyINFOs
                                '      of the users, but DDCH sends all the MyINFOs with the nicklist
                                '      and ignores GetINFO requests
                                
                                'Delaying the nicklist is is a major bandwidth saving feature
                                'GetINFO, NickList and OpList are not sent until MyINFO is validated (and passes)
                                'Also MyINFO is not sent if it fails the checks (and neither is Hello, therefore not
                                'requiring a Quit message)
                                
                                'Check for flooding
859:                                If g_objSettings.EnableFloodWall Then _
                                        If curUser.FloodCheck(3) Then _
                                        Exit Sub

                                'If user isn't logged in, then queue the nicklist
864:                                If curUser.State = Logged_In Then
                                        'If the user is not visible, then we must add their nickname to the lists
866:                                    If curUser.Visible Then
                                        'Only send the oplist if the user is using QuickList;
                                        'otherwise we send both the nicklist and the oplist
869:                                        If Not curUser.NoHello Then curUser.SendData "$NickList " & g_colUsers.NickList & "|"
870:                                        curUser.SendData "$OpList " & g_colUsers.OpList & "|"
871:                                    Else
872:                                        strKey = curUser.sName
                                        
                                        'Add user's name to nicklist if we're sending
875:                                        If Not curUser.NoHello Then curUser.SendData "$NickList " & g_colUsers.NickList & strKey & "$$|"
                                        
                                        'Add user's name to oplist if they are an operator
878:                                        If curUser.bOperator Then _
                                            curUser.SendData "$OpList " & g_colUsers.OpList & strKey & "$$|" _
                                        Else _
                                            curUser.SendData "$OpList " & g_colUsers.OpList & "|"
                                            
                                        'Send their MyINFO string to themselves (since they are invisible)
884:                                        If g_objSettings.HideMyinfos Then
885:                                            If curUser.Class < g_objSettings.MinMyinfoFakeCls Then
886:                                                curUser.SendData curUser.sMyInfoFakeString & "|"
887:                                            Else
888:                                                curUser.SendData curUser.sMyInfoString & "|"
889:                                            End If
890:                                        Else
891:                                            curUser.SendData curUser.sMyInfoString & "|"
892:                                        End If
893:                                    End If

                                    #If FLASHCHAT Then
                                        'ChatOnly client, should not need to refresh Nicklist
897:                                     If Not curUser.NullCharSeparator Then
                                    #End If

                                    'Build MyINFO stream
901:                                    For Each m_objLoopUser In g_colUsers
902:                                        If m_objLoopUser.Visible Then
903:                                            If g_objSettings.HideMyinfos Then
904:                                                If curUser.Class < g_objSettings.MinMyinfoFakeCls Then
905:                                                    strTMyinfosStr = strTMyinfosStr & m_objLoopUser.sMyInfoFakeString & "|"
906:                                                Else
907:                                                    strTMyinfosStr = strTMyinfosStr & m_objLoopUser.sMyInfoString & "|"
908:                                                End If
909:                                            Else
910:                                                strTMyinfosStr = strTMyinfosStr & m_objLoopUser.sMyInfoString & "|"
911:                                            End If
912:                                        End If
913:                                    Next
                                    'Send MyINFO stream
915:                                    curUser.SendData strTMyinfosStr
                                    'TheNOP End

                                    'Send Bot MyINFOs
919:                                    UpdateBots curUser
                                  
                                    #If FLASHCHAT Then
922:                                     End If
                                    #End If
924:                                Else
925:                                    curUser.QNL = True
926:                                End If
                                #If Status Then
928:                                g_objStatus.UpDateProtocol iNickList
                                #End If
930:                            End If
                        Case 84 'T
                            'Possible messages :
                            '   -- To: (Private message)

                            'To:
                            '
                            '   -- Format   : $To: <name> From: <from> $<<from>> <message>|
                            '   -- Response : N/A
                            '
                            '   -- Standard protocol message
                            '   -- Sends a "private" message to another user (the hub owner can
                            '      actually read it so it isn't really private; but hub owners which
                            '      do that are quite lame and I wish I could add idiotic protection to
                            '      DDCH to prevent them from using it *ahem*)

                            'Check if the user is muted
946:                            If Not curUser.Mute Then

                            'Make sure the user isn't flooding
949:                            If g_objSettings.EnableFloodWall Then
                                'don't check >= vips
951:                                If curUser.Class < Vip Then
                                    'If = 0 then disable main chat flood check
953:                                    If g_objSettings.FWMainChat Then
954:                                        If curUser.FloodCheck(4) Then Exit Sub
955:                                    End If
956:                                End If
957:                            End If

959:                            strKey = LeftB$(strParts, InStrB(3, strParts, " ") - 1)

                                'If the name is either the bot name or op chat name, take special actions
                                Select Case strKey
                                    Case g_objSettings.BotName
                                        'PMs to the bot normally mean it's a command

964:                                        strKey = MidB$(strParts, InStrB(InStrB(1, strParts, "$"), strParts, " ") + 2)
965:                                        If LenB(strKey) Then _
                                                If AscW(strKey) = g_objSettings.CPrefix Then _
                                                If g_objSettings.EnabledCommands Then ProcessTrigger curUser, MidB$(strKey, 3), False
                        
                                    Case Else
                                    
                                            'This function return min class to use the chat rum
971:                                        lngTemp = g_objChatRoom.ProcessChat(strKey)
972:                                        If lngTemp <> -1 Then
                                                 'Check if only ops or vips can use the vip chat
974:                                             If Not curUser.Class >= lngTemp Then
975:                                                   GoTo NextLoop
976:                                             End If
        
978:                                             strKey = " From: " & strKey & " " & MidB$(strParts, InStrB(1, strParts, "$")) & "|"
979:                                             lngPos = ObjPtr(curUser)
        
981:                                             For Each m_objLoopUser In g_colUsers
982:                                                 If m_objLoopUser.Class >= lngTemp Then
983:                                                     If Not lngPos = ObjPtr(m_objLoopUser) Then
984:                                                          m_objLoopUser.SendData "$To: " & m_objLoopUser.sName & strKey
985:                                                     End If
986:                                                 End If
987:                                             Next
988:                                        Else
                                                'Normal user; check if they are online
                                                'svn 216
                                                'check for nick spoofing
992:                                            If curUser.sName = g_objRegExps.CaptureSubStr(strCommand, GETFROMNICKINPM) Then
993:                                                If curUser.sName = g_objRegExps.CaptureSubStr(strCommand, GETNICKINPMMSG) Then
994:                                                    If g_colUsers.Online(strKey) Then g_colUsers.ItemByName(strKey).SendData strCommand & "|"
995:                                                Else
996:                                                    curUser.Kick 60, "PTDCH / Core", "Detected nick spoofing"
997:                                                    Exit Sub
998:                                                End If
999:                                            Else
1000:                                               curUser.Kick 60, "PTDCH / Core", "Detected nick not online"
1001:                                               Exit Sub
1002:                                           End If
1003:                                       End If
1004:                                End Select
                                #If Status Then
1006:                               g_objStatus.UpDateTrafic bytesTotal, iMsgPM
                                #End If
1008:                            End If
                        Case 75 'K
                            'Possible messages :
                            '   -- Key
                            '   -- Kick
                            
                            Select Case LenB(strKey)
                                Case 6
                                    'Key
                                    '
                                    '   -- Format   : $Key <string>|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message (ignored)
                                    '   -- Sent in response to $Lock
                                    '   -- Originally it was a security check
                                    '      to prevent unauthorized DC clients from
                                    '      connecting to the hub; now it's just useless
                                    
                                    'If Not strParts = vbKey Then wskLoop_Close Index: Exit Sub
1025:                                    curUser.State = Wait_Validate
                                Case 8
                                    'Kick
                                    '
                                    '   -- Format   : $Kick <name>|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message (for ops only)
                                    '   -- Disconnects and temp bans the user using <name>
                                    
1034:                                    If curUser.bOperator Then
1035:                                        If g_colUsers.Online(strParts) Then
1036:                                            Set m_objLoopUser = g_colUsers.ItemByName(strParts)
                                                 'Cannot kick user above or equal to own class (unless the user is an admin)
                                                 Select Case curUser.Class
                                                    Case Admin, InvisibleAdmin, Is > m_objLoopUser.Class
                                                        'Get IP because we need it for message
1040:                                                    strKey = m_objLoopUser.IP
1041:                                                    DoEvents
1042:                                                    m_objLoopUser.Kick -1, curUser.sName
1043:                                                    g_colUsers.SendChatToOps g_objSettings.BotName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("KickedBy"), "%[user]", strParts), "%[op]", curUser.sName), "%[ip]", strKey)
                                                     #If Status Then
1045:                                                    g_objStatus.UpDateProtocol iKicks
                                                     #End If
1047:                                            End Select
                                            
1049:                                            Set m_objLoopUser = Nothing
1050:                                        End If
1051:                                    End If
1052:                            End Select
                        Case 86 'V
                            'Possible messages :
                            '   -- ValidateNick
                            '   -- Version
                            
                            Select Case LenB(strKey)
                                Case 24
                                    'ValidateNick
                                    '
                                    '   -- Format   : $ValidateNick <name>|
                                    '   -- Response : $ValidateDenide <name>|  //  $GetPass|  //  $Hello <name>|
                                    '
                                    '   -- Standard protocol message
                                    '   -- Sent in response to intial message $Lock
                                    '   -- Used to see if <name> is used; if it is, then send
                                    '      $ValidateDenide; if it's not taken the send $GetPass
                                    '      if it's registered, or send $Hello if not
                                    
                                    'Rout to another sub (if it returns false, not check any more data)
1069:                                    If Not ValidateNick(curUser, strParts) Then _
                                            Exit Sub
                                Case 14
                                    'Version
                                    '
                                    '   -- Format   : $Version <version>|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message (optional)
                                    '   -- <version> contains (in NMDC) the clients version
                                    '   -- Can be overridden in all other clients (virtually all)
                                    
                                    'Convert to proper decimal
1081:                                    If m_blnCommaDecimal Then
1082:                                        curUser.iVersion = StrToDbl(Replace(strParts, ".", ","))
1083:                                    Else
1084:                                        curUser.iVersion = Val(strParts)
1085:                                    End If
                                    
                                    Select Case False
                                        Case g_objSettings.OPBypass, curUser.Class >= Vip
                                            'NMDC min version check
1088:                                            If g_objSettings.NMDCMinVersion Then _
                                                If g_objSettings.NMDCMinVersion > curUser.iVersion Then _
                                                    FailedConf curUser, NMDCVersion: Exit Sub
1091:                                    End Select
1092:                            End Select
                        Case 79 'O
                            'Possible messages :
                            '   -- OpForceMove (Redirect)
                            
                            'OpForceMove
                            '
                            '   -- Format   : $OpForceMove $Who:<name>$Where:<address>$Msg:<message>|
                            '   -- Response : N/A
                            '
                            '   -- Standard protocol message (for ops only)
                            '   -- Redirects a user to <address> (private messages them <message>)
                            '   -- A point worthy of mention is that the client can choose
                            '      to ignore the message; so make sure they are disconnected
                            
1106:                            If curUser.bOperator Then
                                'Check if the user can redirect
1108:                                If Not g_objSettings.OpsCanRedirect Then _
                                        If curUser.Class < Admin Then _
                                        GoTo NextLoop

1112:                                strParts = MidB$(strParts, 11)
1113:                                lngPos = InStrB(1, strParts, "$")
1114:                                strKey = LeftB$(strParts, lngPos - 1)
                                
                                'Check if user is online
1117:                                If g_colUsers.Online(strKey) Then
1118:                                    Set m_objLoopUser = g_colUsers.ItemByName(strKey)
                                    'Cannot redirect user above or equal to own class (unless the user is an admin)
                                    Select Case curUser.Class
                                        Case Admin, InvisibleAdmin, Is > m_objLoopUser.Class
1120:                                            strParts = MidB$(strParts, lngPos + 14)
1121:                                            lngPos = InStrB(1, strParts, "$")
                                    
                                            'Private message user reason, then redirect
1124:                                            m_objLoopUser.SendPrivate curUser.sName, MidB$(strParts, lngPos + 10)
1125:                                            m_objLoopUser.Redirect LeftB$(strParts, lngPos - 1)
1126:                                    End Select
                                    
1128:                                    Set m_objLoopUser = Nothing
                                     #If Status Then
1130:                                    g_objStatus.UpDateProtocol iRedirects
                                     #End If
1132:                                End If
1133:                            End If
                        Case 78 'N
                            'Possible messages :
                            '   -- NetINFO
                            
                            'NetINFO
                            '
                            '   -- Format   : $NetINFO <slots>$<hubs>$<mode>|
                            '                 $NetINFO <slots>$<hubs>$<mode>$<bandwidth>|
                            '      Response : N/A
                            '
                            '   -- Protocol Extension
                            '   -- Supported by NMDC only (as a substitute for tags)
                            '   -- This is sent in response to a $GetNetInfo|
                            '   -- Two different forms; version 2.02 of DC includes the upload
                            '      bandwidth limit value
                            
1149:                            curUser.NetInfo = True

                            'Skip ops if necessary
                            Select Case False
                                Case g_objSettings.OPBypass, curUser.Class >= Vip
                                    'Extract slots, hubs and bandwidth if provided
1153:                                    lngPos = InStrB(1, strParts, "$")
1154:                                    bytesTotal = CLng(LeftB$(strParts, lngPos - 1))
1155:                                    strParts = MidB$(strParts, lngPos + 2)
                                    
1157:                                    lngPos = InStrB(1, strParts, "$")
1158:                                    strKey = LeftB$(strParts, lngPos - 1)
1159:                                    strParts = MidB$(strParts, lngPos + 2)
1160:                                    lngPos = CLng(strKey)
                                    
                                    'Find out if we need to get upload bandwidth
1163:                                    If Not LenB(strParts) = 2 Then _
                                        strParts = MidB$(strParts, 5) _
                                    Else _
                                        strParts = "0"
                                
                                    'Max hubs
1169:                                    If g_objSettings.DCMaxHubs Then _
                                        If lngPos > g_objSettings.DCMaxHubs Then _
                                            FailedConf curUser, MaxHubs: Exit Sub
                                    
                                    'Min slots
1174:                                    If g_objSettings.MinSlots Then _
                                            If bytesTotal < g_objSettings.MinSlots Then _
                                            FailedConf curUser, MinSlots: Exit Sub
                                            
                                    'Max slots
1179:                                    If g_objSettings.MaxSlots Then _
                                        If bytesTotal > g_objSettings.MaxSlots Then _
                                            FailedConf curUser, MaxSlots: Exit Sub
                                    
                                    'Hub/Slot ratio
1184:                                    If g_objSettings.DCSlotsPerHub Then _
                                        If (bytesTotal / lngPos) < g_objSettings.DCSlotsPerHub Then _
                                            FailedConf curUser, HSRatio: Exit Sub
                                    
                                    'Bandwidth/Slot ratio
1189:                                    If g_objSettings.DCBandPerSlot Then _
                                        If Not strParts = "0" Then _
                                            If (CLng(strParts) / bytesTotal) < g_objSettings.DCBandPerSlot Then _
                                                FailedConf curUser, BSRatio: Exit Sub
1193:                            End Select
                             #If Status Then
1195:                            g_objStatus.UpDateProtocol iNetINFO
                             #End If
                        Case 85 'U
                            'Possible messages :
                            '   -- UserIP
                            
                            'UserIP
                            '   -- Format   : $UserIP <name>[$$<name_etc>]|
                            '      Response : $UserIP <name> <ip>[$$<name_etc> <ip_etc>]|
                            '
                            '   -- Protocol Extension
                            '   -- Ops can get anyone's IP, while users can
                            '      only get their own
                                
1208:                            If curUser.bOperator Then
1209:                                strKey = "$UserIP "
1210:                                strParts = strParts & "$$"
1211:                                lngPos = InStrB(1, strParts, "$$")

                                'Loop to find all ip requests
1214:                                Do While lngPos
1215:                                    strIP = LeftB$(strParts, lngPos - 1)
1216:                                    strParts = MidB$(strParts, lngPos + 4)
                                        
                                    'If online add their ip to the list, otherwise add blank
1219:                                    If g_colUsers.Online(strIP) Then _
                                        strKey = strKey & strIP & " " & g_colUsers.ItemByName(strIP).IP & "$$" _
                                    Else _
                                        strKey = strKey & strIP & "  $$"
                                        
1224:                                    lngPos = InStrB(1, strParts, "$$")
1225:                                Loop
                                
                                'Remove last $$
1228:                                curUser.SendData LeftB$(strKey, LenB(strKey) - 4) & "|"
1229:                            Else
1230:                                curUser.SendData "$UserIP " & curUser.sName & " " & curUser.IP & "|"
1231:                            End If
                        Case 120 'x
                            'Possible messages :
                            '   -- xKick
                            '
                            '   -- Format   : $xKick <name>$<length>$<show>$<msg>|
                            '      Response : N/A
                            '
                            '   -- Protocol Extension
                            '   -- For operators only
                            '   -- Allows op to choose how long the user will be
                            '      banned for as well as whether or not the kick
                            '      will be seen the in the main chat
                            '   -- xKick also contains the kick reason
                            
1246:                            If curUser.bOperator Then
1247:                                lngPos = InStrB(1, strParts, "$")
1248:                                strKey = LeftB$(strParts, lngPos - 1)
                                
                                'Make sure user to kick is online
1251:                            If g_colUsers.Online(strKey) Then
1252:                                Set m_objLoopUser = g_colUsers.ItemByName(strKey)
                                     'Cannot kick user above or equal to own class (unless the user is an admin)
                                      Select Case curUser.Class
                                          Case Admin, InvisibleAdmin, Is > m_objLoopUser.Class
1254:                                            strParts = MidB$(strParts, lngPos + 2)
1255:                                            lngPos = InStrB(1, strParts, "$")
1256:                                            strTemp = MidB$(strParts, InStrB(1, strParts, "$") + 2)
                                                 
                                                 'Ban the IP for the correct length
1258:                                            g_colIPBans.Add m_objLoopUser.IP, CLng(LeftB$(strParts, lngPos - 1)), strKey, curUser.sName, strTemp
                                    
1260:                                            strParts = MidB$(strParts, lngPos + 2)
                                    
                                            'Check to see if it should be sent to the main chat
1263:                                            If AscW(strParts) = 84 Then
1264:                                                strParts = MidB$(strParts, InStrB(1, strParts, "$") + 2)
                                    
1266:                                                g_colUsers.SendChatToAll curUser.sName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("IsKicking"), "%[op]", curUser.sName), "%[user]", strKey), "%[reason]", strParts)
                                        
                                                'Send private message and disconnect user
1269:                                                m_objLoopUser.SendPrivate curUser.sName, curUser.GetCoreMsgStr("KickedBecause") & strParts
1270:                                                DoEvents
1271:                                                wskLoop_Close Index
                                        
                                                'Notify that the user has been disconnected
1274:                                                g_colUsers.SendChatToAll g_objSettings.BotName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("KickedBy"), "%[user]", strKey), "%[op]", curUser.sName), "%[ip]", m_objLoopUser.IP)
1275:                                            Else
1276:                                                strParts = MidB$(strParts, InStrB(1, strParts, "$") + 2)
1277:                                                g_colUsers.SendChatToOps curUser.sName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("IsKicking"), "%[op]", curUser.sName), "%[user]", strKey), "%[reason]", strParts)
                                        
                                                'Send private message and disconnect user
1280:                                                m_objLoopUser.SendPrivate curUser.sName, curUser.GetCoreMsgStr("KickedBecause") & strParts
1281:                                                DoEvents
1282:                                                wskLoop_Close Index
                                        
                                                'Notify that the user has been disconnected
1285:                                                g_colUsers.SendChatToOps g_objSettings.BotName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("KickedBy"), "%[user]", strKey), "%[op]", curUser.sName), "%[ip]", m_objLoopUser.IP)
1286:                                            End If
                                             #If Status Then
1288:                                            g_objStatus.UpDateProtocol iKicks
                                             #End If
1290:                                    End Select
                                    
1292:                                    Set m_objLoopUser = Nothing
1293:                                End If
1294:                            End If
                        Case 66 'B
                            'Possible messages :
                            '   -- BotINFO
                            '   -- BlackDC (ignored)
                            
1299:                           If AscW(RightB$(strKey, 2)) = 79 Then 'O
                                'BotINFO
                                '   -- Format   : $BotINFO|
                                '      Response : $HubINFO <name>$<ip>$<port>$<description>$<maxusers>$<minshare>$<minslots>$<maxhubs>$<extra>$|
                                '
                                '   -- Protocol Extension
                                '   -- Used by Gadget's hub pinger for www.hublist.org
                                '   -- Gives various extra information on hub such as
                                '      the min share/slots, max hubs, etc
                                
1309:                                curUser.SendData "$HubINFO " & g_objSettings.HubName & "$" & g_objSettings.HubIP & ":" & g_objSettings.Port & _
                                                      "$" & g_objSettings.HubDesc & "$" & g_objSettings.MaxUsers & "$" & g_objSettings.MinShare & _
                                                      "$" & g_objSettings.MinSlots & "$" & g_objSettings.DCMaxHubs & "$PTDCH " & vbVersion & " Built-In$|"
1312:                                AddLog "Public Hublist Pinger from nick: " & curUser.sName & " - IP: " & curUser.IP
                                 #If Status Then
1314:                                g_objStatus.UpDateProtocol iBotINFO
                                 #End If
1316:                            End If
                        Case 122 'z
                            'Possible messages :
                            '  -- zSearch
                            
                            
                        'Case Else
                            'Perhaps one day I'll do something with unknown messages
1323:                    End Select
                Case 60 '<
                    'Main chat message
                    '
                    '   -- Format   : <<name>> <message>|
                    '   -- Response : N/A
                    '
                    '   -- Standard protocol message
                    '   -- Sends a main chat message to all users

                    'Make sure the user is logged in
1333:                    If curUser.State = Logged_In Then
                        'Check if the user is muted
1335:                        If Not curUser.Mute Then
                                                            
                            ' TheNOP svn 159
                            'Make sure the user isn't flooding
1339:                            If g_objSettings.EnableFloodWall Then
                                'If = 0 then disable PM flood checking
1341:                                If g_objSettings.FWMainChat Then
                                    'don't check >= vips, kick raw can occure fast ;)
1343:                                    If curUser.Class < Vip Then
1344:                                        If curUser.FloodCheck(4) Then Exit Sub
1345:                                    End If
1346:                                End If
1347:                            End If

                            'Truncate message if necessary
1350:                            If g_objSettings.MaxMessageLen Then
1351:                                If curUser.Class < Vip Then
1352:                                    lngPos = Len(strCommand)
                                    'svn 216  getlang...
1354:                                    If lngPos > g_objSettings.MaxMessageLen Then
1355:                                            curUser.SendChat g_objSettings.BotName, "your message was to big to be sent to other users"
                                    'If lngPos > g_objSettings.MaxMessageLen Then strCommand = Left$(strCommand, g_objSettings.MaxMessageLen)
1357:                                            Exit Sub
1358:                                        End If
1359:                                End If
1360:                            End If
                        
1362:                            lngPos = InStrB(1, strCommand, "> ")
                        
                            'Make sure the client isn't trying to fake it's username
                            'If so, then replace the fake name with the real one
1366:                            If Not MidB$(strCommand, 3, lngPos - 3) = curUser.sName Then
                                'Replace
                                'strCommand = "<" & curUser.sName & "> " & MidB$(strCommand, lngPos + 4)
                                'lngPos = InStrB(1, strCommand, " ")
                                'svn 216
                                'Kick without message, they know why they are kicked...
1372:                                curUser.Kick 60
1373:                                Exit Sub
1374:                            End If
                        
                            'Get the message
1377:                            strKey = MidB$(strCommand, lngPos + 4)
                        
                            'Don't send if there is no string
1380:                            If LenB(strKey) Then
                                'If first character is the command prefix, then process it
1382:                                If AscW(strKey) = g_objSettings.CPrefix Then
1383:                                    If g_objSettings.EnabledCommands Then ProcessTrigger curUser, MidB$(strKey, 3), True
1384:                                    If g_objSettings.FilterCPrefix Then lngPos = 0 Else lngPos = -1
1385:                                End If
                            
                                'Send out main chat message
1388:                                If lngPos Then _
                                    If g_objSettings.SendMessageAFK Then _
                                        g_colUsers.SendToAll strCommand & "|" _
                                    Else _
                                        g_colUsers.SendToNA strCommand & "|"
1393:                            End If
                             #If Status Then
1395:                            g_objStatus.UpDateTrafic bytesTotal, iMsgMainChat
                             #End If
1397:                        End If
1398:                    End If
1399:            End Select
            
            'Call dataarrival event if necessary
            #If DataArrival Then
1403:                SEvent_DataArrival curUser, strCommand
            #End If
            
1406:       End If
    
1408:
NextLoop:
        'Find next pipe to parse the next message
1410:        lngPos = InStrB(1, strData, "|")
1411:    Loop

    'If there is any data left over, put the fragment into user's data var
1414:    If LenB(strData) Then
1415:        If curUser.bOperator Then
1416:            curUser.DataFragment = strData
1417:        Else
1418:            If LenB(curUser.DataFragment) > (g_objSettings.DataFragmentLen * 2) Then
1419:                curUser.DataFragment = vbNullString
1420:            Else
1421:                curUser.DataFragment = strData
1422:            End If
1423:        End If
1424:    End If

     #If Status Then
1427:    g_objStatus.UpDateTrafic bytesTotal, iTotalRecived
1428:    g_objStatus.UpDateProtocol iRequests
     #End If

1431:    Exit Sub

1433:
LoopErr:
    'Error occured when trying to parse message
1435:    HandleError Err.Number, Err.Description, Erl & "|" & "wskLoop_DataArrival() (Loop - strCommand = """ & strCommand & """; strData = """ & strData & """)"
     #If Status Then
1437:    g_objStatus.UpDateProtocol iErrors
     #End If
1439:    Exit Sub
1440:
Err:
    'Error occured before parsing occured
1442:    HandleError Err.Number, Err.Description, Erl & "|" & "wskLoop_DataArrival() (Preloop - " & strData & ")"
     #If Status Then
1444:    g_objStatus.UpDateProtocol iErrors
     #End If
End Sub
Private Sub wskLoop_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
1:    On Error Resume Next
2:    wskLoop_Close Index
  #If Status Then
4:    g_objStatus.UpDateProtocol iErrSockets
  #End If
End Sub
Private Sub wskLoop_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
1:    On Error Resume Next
  #If Status Then
3:    g_objStatus.UpDateTrafic bytesSent, iTotalSend
  #End If
End Sub
Private Sub wskRegister_Close(Index As Integer)
1:    On Error Resume Next
     'Just close it; don't really care about the errors here
3:    wskRegister(Index).Close
End Sub
Private Sub wskRegister_DataArrival(Index As Integer, ByVal bytesTotal As Long)
1:    Dim strData     As String
2:    Dim strCommand  As String
3:    Dim strLock     As String
4:    Dim lngPos      As Long
    
6:    On Error Resume Next
      'Get data
8:    wskRegister(Index).GetData strData, vbString

    #If SVN Then
11:        g_objFileAccess.AppendFile G_LOGPATH, Now & " <-- " & wskRegister(Index).RemoteHostIP & " - " & strData
    #End If
13:    On Error GoTo Err
    
15:    lngPos = InStrB(1, strData, "|")
16:    Do While lngPos
17:        strCommand = LeftB$(strData, lngPos - 1)
18:        strData = MidB$(strData, lngPos + 2)
    
        'Possible messages :
        '   -- Lock
    
23:        If LeftB$(strCommand, 10) = "$Lock" Then
            'Lock
            '
            '   -- Format   : $Lock <string> pk=<astring>|
            '   -- Response : $Key <string>|<name>|<ip>[:<port>]|<description>|<users>|<bytes>|
            '
            '   -- Standard protocol message
            '   -- The <string> from Lock is decoded into the <string> for key;
            '      The information which follows is various details about the hub
            '      like the hub name, address, description, users, and total shared
            '      bytes
        
35:            lngPos = wskRegister(Index).LocalPort
        
37:            strLock = "$Key " & LockToKey(MidB$(strCommand, 13, LenB(strCommand) - 14), _
                                            ((lngPos \ 256) + (lngPos And 255)) And 255) _
                                 & "|"
                             
            'Add char160 to the end of the hub name to prevent MoGLO, MoSearch and GLOSearch
42:            If g_objSettings.PreventSearchBots Then _
                    strLock = strLock & g_objSettings.HubName & vbChar160 & "|" _
            Else _
                strLock = strLock & g_objSettings.HubName & "|"
                
            'If the port is 411, then don't add it to the address
48:            If g_objSettings.Port = 411 Then _
                     strLock = strLock & g_objSettings.HubIP & "|" _
            Else _
                strLock = strLock & g_objSettings.HubIP & ":" & g_objSettings.Port & "|"
                
            'Add char160 to the end of the hub description to prevent MoGLO, MoSearch and GLOSearch
54:            If g_objSettings.PreventSearchBots Then _
                strLock = strLock & g_objSettings.HubDesc & vbChar160 & "|" _
            Else _
                strLock = strLock & g_objSettings.HubDesc & "|"
                
            'Add user count and total bytes
60:            strLock = strLock & g_colUsers.Count & "|" & g_colUsers.iTotalBytesShared & "|"
            
62:            On Error Resume Next
            
            #If SVN Then
65:                g_objFileAccess.AppendFile G_LOGPATH, Now & " --> " & wskRegister(Index).RemoteHostIP & " - " & strLock
            #End If
            
68:            On Error GoTo Err
            'Submit registration
70:            wskRegister(Index).SendData strLock
71:        End If
        
73:        lngPos = InStrB(1, strData, "|")
74:    Loop
    
        'The auto disconnect has been removed due to behaviour with NMDCH
        'After researching it with my own registration server, I found it doesn't
        'disconnect, but rather it is the server which ends the connection
      
        'This could have had an impact with registering with vandel405.dynip.com
        'but now I believe however it perfectly emulates NMDCH behaviour in this respect
      
        'DoEvents
        'wskRegister(Index).Close
    
86:    Exit Sub
87:
Err:
88:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.wskRegister_DataArrival()"
End Sub

Private Sub wskRegister_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Just close it; don't really care about the errors here
2:    On Error Resume Next
    
   #If SVN Then
5:     g_objFileAccess.AppendFile G_LOGPATH, "wskRegister_Error: " & Description & " | Scode: " & Scode & " | Index: " & Index
   #End If
7:     wskRegister(Index).Close
   #If Status Then
9:     g_objStatus.UpDateProtocol iErrors
   #End If
End Sub
'------------------------------------------------------------------------------
'Winsock events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Core related private/public methods
'------------------------------------------------------------------------------
Public Sub SwitchServing()
1:    Dim lngLoop         As Long
2:    Dim lngUB           As Long
3:    Dim lngPos          As Long
4:    Dim arrTemp()       As String
    
6:    On Error GoTo Err
    
    'Find out which state we're in
9:    If G_SERVING Then
        'Stop serving
11:        G_SERVING = False

        'GUI related
' ------------------------ NEW INTERFACE LANGUAGE ------------------------
15:        cmdButton(1).Caption = m_arrDynaCap(0) '"Start Serving"
' ----------------------- NEW INTERFACE LANGUAGE END ----------------------
17:        txtData(3).Enabled = True
18:        txtData(4).Enabled = True
19:        txtData(7).Enabled = True
20:        txtData(49).Enabled = True
21:        chkData(21).Enabled = True
22:        cmdButton(7).Enabled = False
23:        cmdButton(8).Enabled = False
24:        cmdButton(9).Enabled = False
25:        cmdButton(2).Enabled = False

27:        tmrBackground.Enabled = False
     
        'Clear out listening winsocks
30:        wskListen(0).Close
31:        lngUB = wskListen.UBound
        
33:        If lngUB Then
34:            For lngLoop = 1 To lngUB
35:                wskListen(lngLoop).Close
36:                Unload wskListen(lngLoop)
37:            Next
38:        End If
        
        'Clear out user winsocks
41:        If wskLoop(0).State Then wskLoop(0).Close
42:        lngUB = wskLoop.UBound
        
44:        If lngUB Then
45:            For lngLoop = 1 To lngUB
46:                If wskLoop(lngLoop).State Then wskLoop(lngLoop).Close
47:                Unload wskLoop(lngLoop)
48:            Next
49:        End If
        
        'Clear out registration winsocks
52:        If wskRegister(0).State Then wskRegister(0).Close
53:        lngUB = wskRegister.UBound
        
55:        If lngUB Then
56:            For lngLoop = 1 To lngUB
57:                If wskRegister(lngLoop).State Then wskRegister(lngLoop).Close
58:                Unload wskRegister(lngLoop)
59:            Next
60:        End If
        
        'Clear out  collections
        #If COLFREESOCKS Then
64:            Set m_colFreeSocks = Nothing
        #End If
    
67:        g_colUsers.Clear
        
        'Remove bot names, if used
70:        If g_objSettings.UseBotName Then UnregisterBotName g_objSettings.BotName
71:        Call g_objChatRoom.UnRegisterChat
            
        'Set serving date
74:        m_datServingDate = Now

        'Raise event
77:        SEvent_StoppedServing
          
        'Show Ballon notification
80:        If g_objSettings.PopUpStopedServing Then g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, g_objSettings.HubName & vbNewLine & "Server Stoped", 0, True
81:        AddLog "Server stoped."
          
83:    Else

       #If Status Then
           'Inicialize new statistic
87:        g_objStatus.IniStatistics
       #End If

        'Show Ballon notification
91:        If g_objSettings.PopUpStartedServing Then g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, g_objSettings.HubName & vbNewLine & "Server Started", 0, True
92:        AddLog "Server started."
93:        AddLog "Listening ports: " & g_objSettings.Ports

        'Start serving
96:        G_SERVING = True

        'GUI related
' ------------------------ NEW INTERFACE LANGUAGE ------------------------
100:        cmdButton(1).Caption = m_arrDynaCap(1)  '"Stop Serving"
' ----------------------- NEW INTERFACE LANGUAGE END ----------------------
102:       txtData(3).Enabled = False
103:       txtData(4).Enabled = False
104:       txtData(7).Enabled = False
105:       txtData(49).Enabled = False
106:       chkData(21).Enabled = False
107:       cmdButton(7).Enabled = True
108:       cmdButton(8).Enabled = True
109:       cmdButton(9).Enabled = True
110:       cmdButton(2).Enabled = True
   
        'Create objects
        #If COLFREESOCKS Then
117:        Set m_colFreeSocks = New Collection
        #End If
        
120:        tmrBackground.Enabled = True
        
        'Get listening ports
123:        arrTemp = Split(g_objSettings.Ports, ";")
124:        lngUB = UBound(arrTemp)

126:        For lngLoop = 0 To lngUB
127:            If IsNumeric(arrTemp(lngLoop)) Then
                'Load winsock as necessary
129:                If lngLoop Then _
                    Load wskListen(lngLoop) _
                Else _
                    g_objSettings.Port = CInt(arrTemp(0))
                    
                'Set port and listen
135:                wskListen(lngLoop).LocalPort = CLng(arrTemp(lngLoop))
                
137:                On Error Resume Next
138:                wskListen(lngLoop).Listen
139:                On Error GoTo Err
                
                'Check if there was an error
142:                If Err.Number = 10048 Then
143:                    MsgBoxCenter Me, Replace(g_colMessages.Item("msgPortInUse"), "%[port]", arrTemp(lngLoop)), vbCritical, g_colMessages.Item("msgStartServing")
144:                    Err.Clear
145:                End If
146:            End If
147:        Next
        
        'Get registration servers
150:        If LenB(g_objSettings.RegisterIP) Then
151:            arrTemp = Split(g_objSettings.RegisterIP, ";")
152:            lngUB = UBound(arrTemp)
        
154:            For lngLoop = 0 To lngUB
                'Load winsock
156:                If lngLoop Then Load wskRegister(lngLoop)
            
                'Get port or set to default 2501
159:                lngPos = InStrB(1, arrTemp(lngLoop), ":")
160:                If lngPos Then
161:                    wskRegister(lngLoop).RemoteHost = LeftB$(arrTemp(lngLoop), lngPos - 1)
162:                    arrTemp(lngLoop) = MidB$(arrTemp(lngLoop), lngPos + 2)
                    
                    'If not numeric, set to default
165:                    If IsNumeric(arrTemp(lngLoop)) Then _
                        wskRegister(lngLoop).RemotePort = CLng(arrTemp(lngLoop)) _
                    Else
168:                        wskRegister(lngLoop).RemotePort = 2501
169:                Else
170:                    wskRegister(lngLoop).RemoteHost = arrTemp(lngLoop)
171:                    wskRegister(lngLoop).RemotePort = 2501
172:                End If
173:            Next
174:        End If
        
        'Preload winsocks if necessary
177:        If g_objSettings.PreloadWinsocks Then
178:            lngUB = g_objSettings.MaxUsers + (((g_objSettings.MaxUsers \ 100) + 1) * 5)
            
180:            For lngLoop = 1 To lngUB
181:                Load wskLoop(lngLoop)
                
                #If COLFREESOCKS Then
184:                    m_colFreeSocks.Add wskLoop(lngLoop), CStr(lngLoop)
                #End If
186:            Next
187:        End If
        
        'Add bot names which were registered before serving was started
190:        If Not m_lngBotsUB = -1 Then
191:            For lngLoop = 0 To m_lngBotsUB
192:                g_colUsers.AppendNL m_arrBots(lngLoop).Name, m_arrBots(lngLoop).Operator
193:            Next
194:        End If
        
        'Register bot names
197:        If g_objSettings.UseBotName Then RegisterBotName g_objSettings.BotName, , , "Bot", , g_objSettings.BotEmail
198:        Call g_objChatRoom.RegisterBots
            
        'Set serving date
201:        m_datServingDate = Now
        
        'Raise event
204:        SEvent_StartedServing
205:    End If

        'Clear user listview on start serving
        #If Status Then
209:            g_objStatus.UClear
        #End If
    
212:    Exit Sub
    
214:
Err:
215:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SwitchServing()"
216:    Resume Next
End Sub
Private Sub FailedConf(ByRef curUser As clsUser, ByRef intType As enuAlert)
1:    Dim strMessage As String

3:    On Error GoTo Err

5:   If SEvent_FailedConf(curUser, intType) Then Exit Sub

    'Find out which message to send
    Select Case intType
'----------ROLL---USERS---REDIRECT--TO--RIGHT--ADDRESS-----------------
'-------------------------------
        Case MaxHubs
'-------------------------------
            'Redirect if necessary
12:            If g_objSettings.RedirectFMaxHubs Then
13:                NextRedirect
            
                'Send message as either private or in the main chat
16:                If g_objSettings.SendMsgAsPrivate Then
17:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxHubs"), "%[maxhubs]", g_objSettings.DCMaxHubs, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxHubsRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxHubsRedirectAddress
18:                Else
19:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxHubs"), "%[maxhubs]", g_objSettings.DCMaxHubs, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxHubsRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxHubsRedirectAddress
20:                End If
21:            Else
22:                strMessage = Replace(curUser.GetCoreMsgStr("MaxHubs"), "%[maxhubs]", g_objSettings.DCMaxHubs, 1, 1)
23:            End If
'-------------------------------
        Case MinSlots
'-------------------------------
            'Redirect if necessary
27:            If g_objSettings.RedirectFMinSlots Then
28:                NextRedirect
                'Send message as either private or in the main chat
30:                If g_objSettings.SendMsgAsPrivate Then
31:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MinSlots"), "%[minslots]", g_objSettings.MinSlots, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMinSlotsRedirectAddress & "|$ForceMove " & g_objSettings.ForMinSlotsRedirectAddress
32:                Else
33:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MinSlots"), "%[minslots]", g_objSettings.MinSlots, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMinSlotsRedirectAddress & "|$ForceMove " & g_objSettings.ForMinSlotsRedirectAddress
34:                End If
35:            Else
36:                strMessage = Replace(curUser.GetCoreMsgStr("MinSlots"), "%[minslots]", g_objSettings.MinSlots, 1, 1)
37:            End If
'-------------------------------
        Case MaxSlots
'-------------------------------
            'Redirect if necessary
41:            If g_objSettings.RedirectFMaxSlots Then
42:                NextRedirect
                'Send message as either private or in the main chat
44:                If g_objSettings.SendMsgAsPrivate Then
45:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxSlots"), "%[maxslots]", g_objSettings.MaxSlots, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxSlotsRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxSlotsRedirectAddress
46:                Else
47:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxSlots"), "%[maxslots]", g_objSettings.MaxSlots, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxSlotsRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxSlotsRedirectAddress
48:                End If
49:            Else
50:                strMessage = Replace(curUser.GetCoreMsgStr("MaxSlots"), "%[maxslots]", g_objSettings.MaxSlots, 1, 1)
51:            End If
'-------------------------------
        Case NMDCVersion
'-------------------------------
            'Redirect if necessary
55:            If g_objSettings.RedirectFTooOldNMDC Then
56:                NextRedirect
                'Send message as either private or in the main chat
58:                If g_objSettings.SendMsgAsPrivate Then
59:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("NMDCMinVersion"), "%[minversion]", g_objSettings.NMDCMinVersion, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForTooOldNMDCRedirectAddress & "|$ForceMove " & g_objSettings.ForTooOldNMDCRedirectAddress
60:                Else
61:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("NMDCMinVersion"), "%[minversion]", g_objSettings.NMDCMinVersion, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForTooOldNMDCRedirectAddress & "|$ForceMove " & g_objSettings.ForTooOldNMDCRedirectAddress
62:                End If
63:            Else
64:                strMessage = Replace(curUser.GetCoreMsgStr("NMDCMinVersion"), "%[minversion]", g_objSettings.NMDCMinVersion, 1, 1)
65:            End If
'-------------------------------
        Case DCppversion
'-------------------------------
            'Redirect if necessary
69:            If g_objSettings.RedirectFTooOldDCpp Then
70:                NextRedirect
                'Send message as either private or in the main chat
72:                If g_objSettings.SendMsgAsPrivate Then
73:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("DCppMinVersion"), "%[minversion]", g_objSettings.DCMinVersion, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForTooOldDcppRedirectAddress & "|$ForceMove " & g_objSettings.ForTooOldDcppRedirectAddress
74:                Else
75:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("DCppMinVersion"), "%[minversion]", g_objSettings.DCMinVersion, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForTooOldDcppRedirectAddress & "|$ForceMove " & g_objSettings.ForTooOldDcppRedirectAddress
76:                End If
77:            Else
78:                strMessage = Replace(curUser.GetCoreMsgStr("DCppMinVersion"), "%[minversion]", g_objSettings.DCMinVersion, 1, 1)
79:            End If
'-------------------------------
        Case HSRatio
'-------------------------------
            'Redirect if necessary
83:            If g_objSettings.RedirectFSlotPerHub Then
84:                NextRedirect
                'Send message as either private or in the main chat
86:                If g_objSettings.SendMsgAsPrivate Then
87:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("HSRatio"), "%[hsratio]", g_objSettings.DCSlotsPerHub, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForSlotPerHubRedirectAddress & "|$ForceMove " & g_objSettings.ForSlotPerHubRedirectAddress
88:                Else
89:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("HSRatio"), "%[hsratio]", g_objSettings.DCSlotsPerHub, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForSlotPerHubRedirectAddress & "|$ForceMove " & g_objSettings.ForSlotPerHubRedirectAddress
90:                End If
91:            Else
92:                strMessage = Replace(curUser.GetCoreMsgStr("HSRatio"), "%[hsratio]", g_objSettings.DCSlotsPerHub, 1, 1)
93:            End If
'-------------------------------
        Case BSRatio
'-------------------------------
        'Redirect if necessary
97:            If g_objSettings.RedirectFBWPerSlot Then
98:                NextRedirect
                'Send message as either private or in the main chat
100:                If g_objSettings.SendMsgAsPrivate Then
101:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("BSRatio"), "%[bsratio]", g_objSettings.DCBandPerSlot, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForBWPerSlotRedirectAddress & "|$ForceMove " & g_objSettings.ForBWPerSlotRedirectAddress
102:                Else
103:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("BSRatio"), "%[bsratio]", g_objSettings.DCBandPerSlot, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForBWPerSlotRedirectAddress & "|$ForceMove " & g_objSettings.ForBWPerSlotRedirectAddress
104:                End If
105:            Else
106:                strMessage = Replace(curUser.GetCoreMsgStr("BSRatio"), "%[bsratio]", g_objSettings.DCBandPerSlot, 1, 1)
107:            End If
'-------------------------------
        Case NoTag
'-------------------------------
            'Redirect if necessary
111:            If g_objSettings.RedirectFNoTag Then
112:                NextRedirect
                'Send message as either private or in the main chat
114:                If g_objSettings.SendMsgAsPrivate Then
115:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("DenyNoTag") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForNoTagRedirectAddress & "|$ForceMove " & g_objSettings.ForNoTagRedirectAddress
116:                Else
117:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("DenyNoTag") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForNoTagRedirectAddress & "|$ForceMove " & g_objSettings.ForNoTagRedirectAddress
118:                End If
119:            Else
120:                strMessage = curUser.GetCoreMsgStr("DenyNoTag")
121:           End If
'-------------------------------
        Case MaxShare
'-------------------------------
            'Redirect if necessary
125:            If g_objSettings.RedirectFMaxShare Then
126:                NextRedirect
                'Send message as either private or in the main chat
128:                If g_objSettings.SendMsgAsPrivate Then
129:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxShare"), "%[maxshare]", g_objFunctions.ShareSize(g_objSettings.MaxShare), 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxShareRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxShareRedirectAddress
130:                Else
131:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxShare"), "%[maxshare]", g_objFunctions.ShareSize(g_objSettings.MaxShare), 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxShareRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxShareRedirectAddress
132:                End If
133:            Else
134:                strMessage = Replace(curUser.GetCoreMsgStr("MaxShare"), "%[maxshare]", g_objFunctions.ShareSize(g_objSettings.MaxShare), 1, 1)
135:            End If
'-------------------------------
        Case FakeShare
'-------------------------------
            'Redirect if necessary
139:            If g_objSettings.RedirectFFakeShare Then
140:                NextRedirect
                'Send message as either private or in the main chat
142:                If g_objSettings.SendMsgAsPrivate Then
143:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("Faker") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForFakeShareRedirectAddress & "|$ForceMove " & g_objSettings.ForFakeShareRedirectAddress
144:                Else
145:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("Faker") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForFakeShareRedirectAddress & "|$ForceMove " & g_objSettings.ForFakeShareRedirectAddress
146:                End If
147:            Else
148:                If g_objSettings.SendMsgAsPrivate Then
149:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("Faker")
150:                Else
151:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("Faker")
152:                End If
                
154:                DoEvents
155:                g_colIPBans.Add curUser.IP, 180, curUser.sName, "PTDCH / Core", "Fake Share"
156:           End If
'-------------------------------
        Case FakeTag
'-------------------------------
            'Redirect if necessary
160:            If g_objSettings.RedirectFFakeTag Then
161:                NextRedirect
                'Send message as either private or in the main chat
163:                If g_objSettings.SendMsgAsPrivate Then
164:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("Faker") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForFakeTagRedirectAddress & "|$ForceMove " & g_objSettings.ForFakeTagRedirectAddress
165:                Else
166:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("Faker") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForFakeTagRedirectAddress & "|$ForceMove " & g_objSettings.ForFakeTagRedirectAddress
167:                End If
168:            Else
169:                If g_objSettings.SendMsgAsPrivate Then
170:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("Faker")
171:                Else
172:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("Faker")
173:                End If
                
175:                DoEvents
176:                g_colIPBans.Add curUser.IP, 180, curUser.sName, "PTDCH / Core", "Fake Tag"
177:           End If
'-------------------------------
        Case MinShare
'-------------------------------
            'Redirect if necessary
181:            If g_objSettings.RedirectFMinShare Then
182:                NextRedirect
            
                'Send message as either private or in the main chat
185:                If g_objSettings.SendMsgAsPrivate Then
186:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MinShare"), "%[minshare]", g_objFunctions.ShareSize(g_objSettings.MinShare), 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMinShareRedirectAddress & "|$ForceMove " & g_objSettings.ForMinShareRedirectAddress
187:                Else
188:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MinShare"), "%[minshare]", g_objFunctions.ShareSize(g_objSettings.MinShare), 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMinShareRedirectAddress & "|$ForceMove " & g_objSettings.ForMinShareRedirectAddress
189:                End If
190:            Else
191:                strMessage = Replace(curUser.GetCoreMsgStr("MinShare"), "%[minshare]", g_objFunctions.ShareSize(g_objSettings.MinShare), 1, 1)
192:            End If

'-------------------------------
        Case Socks5
'-------------------------------
            'g_objSettings.Socks5Msg
            'I think socks5 dont follow redirecting?(rarely)
198:            strMessage = curUser.GetCoreMsgStr("Socks5")

'-------------------------------
        Case PassiveMode
'-------------------------------
202:            If g_objSettings.RedirectFPasMode Then
203:                NextRedirect
                'Send message as either private or in the main chat
205:                If g_objSettings.SendMsgAsPrivate Then
206:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("PassiveMode") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForPasModeRedirectAddress & "|$ForceMove " & g_objSettings.ForPasModeRedirectAddress
207:                Else
208:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("PassiveMode") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForPasModeRedirectAddress & "|$ForceMove " & g_objSettings.ForPasModeRedirectAddress
209:                End If
210:            Else
211:                 strMessage = curUser.GetCoreMsgStr("PassiveMode")
212:           End If
           
        Case NoCOClients
214:            strMessage = curUser.GetCoreMsgStr("NoCOClients")

216:    End Select
'---ROLL-----------END---------------REDIRECT---------PART

    'If there is no message, don't send anything
220:    If LenB(strMessage) Then
        'Send message as either private or in the main chat
222:        If g_objSettings.SendMsgAsPrivate Then
223:            curUser.SendPrivate g_objSettings.BotName, strMessage
224:        Else
225:            curUser.SendChat g_objSettings.BotName, strMessage
226:        End If
227:    End If

    'Close winsock
230:    DoEvents
231:    wskLoop_Close curUser.iWinsockIndex

233:    Exit Sub

235:
Err:
236:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.FailedConf(, " & intType & ")"
End Sub

Private Function ProcessMyINFO(ByRef curUser As clsUser, ByRef strMyINFO As String) As Boolean
1:    Dim arrSplit()  As String
2:    Dim arrTag()    As String
3:    Dim lngLoop     As Long
4:    Dim lngUB       As Long
5:    Dim dblShare    As Double
6:    Dim dblVersion  As Double
7:    Dim lngSlots    As Long
8:    Dim lngHubs     As Long
9:    Dim lngO        As Long
10:    Dim intID       As Integer
11:    Dim strStatus      As String

13:    On Error GoTo Err
    'It is parsed upto the $ALL part of the string
    'Format : $ALL <name> <description>$ $<connection><chr_flag>$<email>$<share>$

    'Check if client is ChatOnly and/or in away mode
18:    strStatus = g_objRegExps.CaptureSubStr(strMyINFO, GETSTATUS)

20:    If LenB(strStatus) Then
        Select Case AscW(strStatus)
            Case 1, 4, 5, 8, 9
21:             curUser.isAFK = False
            Case 2, 3, 6, 7, 10, 11
22:             curUser.isAFK = True
            Case 12, 13
                'ChatOnly client
24:             If g_objSettings.ACOClients Then
25:                 curUser.ChatOnly = True
26:                 curUser.isAFK = False
27:             Else
28:                 If g_objSettings.OPBypass Then
29:                     If curUser.Class < Vip Then
30:                         FailedConf curUser, NoCOClients
31:                         Exit Function
32:                    Else
33:                        curUser.ChatOnly = True
34:                        curUser.isAFK = False
35:                     End If
36:                Else
37:                    FailedConf curUser, NoCOClients
38:                    Exit Function
39:                 End If
40:             End If
            Case 14, 15
                'ChatOnly client
42:             If g_objSettings.ACOClients Then
43:                 curUser.ChatOnly = True
44:                 curUser.isAFK = True
45:             Else
46:                 If g_objSettings.OPBypass Then
47:                     If curUser.Class < Vip Then
48:                         FailedConf curUser, NoCOClients
49:                         Exit Function
50:                    Else
51:                        curUser.ChatOnly = True
52:                        curUser.isAFK = True
53:                     End If
54:                Else
55:                    FailedConf curUser, NoCOClients
56:                    Exit Function
57:                 End If
58:             End If
            Case Else
                'unknown Status--> buggy client or fake tag
60:                 If g_objSettings.DCValidateTags Then
61:                        If g_objSettings.OPBypass Then
62:                            If curUser.Class < Vip Then
                                'Fake tag...
64:                             FailedConf curUser, FakeTag
65:                             Exit Function
66:                            Else
67:                                curUser.isAFK = False
68:                            End If
69:                        Else
                            'Fake tag...
71:                            FailedConf curUser, FakeTag
72:                            Exit Function
73:                        End If
74:                 Else
75:                    curUser.isAFK = False
76:                 End If
77:     End Select
78:    Else
            'Missing status flag
80:            If g_objSettings.DCValidateTags Then
81:                If g_objSettings.OPBypass Then
82:                    If curUser.Class < Vip Then
                        'Fake tag...
84:                        FailedConf curUser, FakeTag
85:                        Exit Function
86:                    Else
87:                        curUser.isAFK = False
88:                    End If
89:                Else
                    'Fake tag...
91:                    FailedConf curUser, FakeTag
92:                    Exit Function
93:                End If
94:            Else
95:                curUser.isAFK = False
96:            End If
97:    End If

99:    intID = 0

101:    arrSplit = Split(MidB$(strMyINFO, 11), "$")
    
    'Make sure we have the right number of params
104:    If UBound(arrSplit) = 5 Then
        'Get their share
106:        dblShare = CDbl(arrSplit(4))
    
        'Make sure the rules apply to this user
        '#If FLASHCHAT Then
110:            If Not curUser.ChatOnly Then
        '#End If
            
113:        If g_objSettings.OPBypass Then
114:            If curUser.Class < Vip Then intID = 1
115:        Else
116:            intID = 1
117:        End If
            
        '#If FLASHCHAT Then
120:            End If
        '#End If

        'If the ID is nonzero, then check rules
124:        If intID Then
            'Set back to zero
126:            intID = 0

            'Check to see if this is an MLDonkey client
            'They usually have a space in front of their share size
130:            If g_objSettings.AutoKickMLDC Then _
                If AscW(arrSplit(4)) = 32 Then _
                    curUser.Kick 60: Exit Function
            
            'Check if they are fake sharing
135:        If g_objSettings.CheckFakeShare Then
136:            If dblShare Then
                Select Case True
                    Case Round(Round(dblShare / 1073741824, 6) * 1073741824, 0) = dblShare
137:                        FailedConf curUser, FakeShare
138:                        Exit Function
                    Case g_objRegExps.TestStr(CStr(dblShare), DENYSHARESIZE)
139:                        FailedConf curUser, FakeShare
140:                        Exit Function
141:                End Select
142:            End If
143:        End If

            'Min share check
146:            If g_objSettings.MentoringSystem Then
                'Good will policy if using mentoring system (must share something)
148:                If dblShare <= 0 Then FailedConf curUser, MinShare: Exit Function
149:            Else
150:                If g_objSettings.MinShare > dblShare Or dblShare < 0 Then FailedConf curUser, MinShare: Exit Function
151:            End If
    
            'Max share check
154:            If g_objSettings.MaxShare Then _
                If g_objSettings.MaxShare < dblShare Then _
                    FailedConf curUser, MaxShare: Exit Function
        
            'Tag checks (get last "<")
159:            lngLoop = InStrB(1, StrReverse(arrSplit(0)), "<")

161:            If lngLoop Then
                'If found, check tag name
163:                lngLoop = LenB(arrSplit(0)) - lngLoop + 2
164:                lngUB = InStrB(lngLoop, arrSplit(0), ">")
                
                'If there is a ">" then check for supported tag names
167:                If lngUB Then
168:                    arrSplit(0) = MidB$(arrSplit(0), lngLoop, lngUB - lngLoop)
169:                    lngUB = InStrB(1, arrSplit(0), " ")
                    
171:                    If lngUB Then
                        'Get the ID
173:                        On Error Resume Next
174:                        intID = m_colTags(LeftB$(arrSplit(0), lngUB - 1)).ID
175:                        On Error GoTo Err
                        
                        'If the ID is nonzero, then it is a real tag
178:                        If intID Then
179:                            arrTag = Split(MidB$(arrSplit(0), lngUB + 2), ",")
            
181:                            lngUB = UBound(arrTag)
            
                            'If lngUB is less than 3, they must be faking their tag
184:                            If lngUB < 3 Then
185:                                If g_objSettings.DCValidateTags Then FailedConf curUser, FakeTag: Exit Function
186:                            Else
                                'Perform the rest of the validation checks if needed
188:                                If g_objSettings.DCValidateTags Then
                                    'First element must be V
190:                                    If AscW(arrTag(0)) = 86 Then arrTag(0) = MidB$(arrTag(0), 5) Else FailedConf curUser, FakeTag: Exit Function
                
                                    'Second element must be M
193:                                    If AscW(arrTag(1)) = 77 Then arrTag(1) = MidB$(arrTag(1), 5) Else FailedConf curUser, FakeTag: Exit Function
    
                                    Select Case g_objRegExps.CaptureSubStr(strMyINFO, GETDCMODE)
                                        Case "A"
195:                                            curUser.Passive = False
                                            
                                        Case "P"
197:                                            If g_objSettings.DenyPassive Then
198:                                                FailedConf curUser, PassiveMode: Exit Function
199:                                            End If
                                            
201:                                            curUser.Passive = True
                                            
                                        Case "5"
203:                                            If g_objSettings.DenySocks5 Then
204:                                                FailedConf curUser, Socks5: Exit Function
205:                                            End If

207:                                            curUser.Passive = True
                                            
                                        Case Else
                                            'more then 1 capture collection or no valid mode
210:                                            FailedConf curUser, FakeTag: Exit Function

212:                                    End Select

                                    'Third element must be H
215:                                    If AscW(arrTag(2)) = 72 Then arrTag(2) = MidB$(arrTag(2), 5) Else FailedConf curUser, FakeTag: Exit Function
                
                                    'Fourth element must be S
218:                                    If AscW(arrTag(3)) = 83 Then arrTag(3) = MidB$(arrTag(3), 5) Else FailedConf curUser, FakeTag: Exit Function
219:                                Else
                                    'Skip beginning "C:" (where C = character) in the beginning of each array element
221:                                    arrTag(0) = MidB$(arrTag(0), 5)
222:                                    arrTag(1) = MidB$(arrTag(1), 5)
223:                                    arrTag(2) = MidB$(arrTag(2), 5)
224:                                    arrTag(3) = MidB$(arrTag(3), 5)
225:                                End If
            
                                'Check the min DC++ version if it is a ++ tag
228:                                If intID = 1 Then
                                    'DC++ does not require $Hello to be sent before $MyINFO, so therefore
                                    'we can skip sending it, saving bandwidth while there are no needed
                                    'protocol changes which DC++ must recognize)
232:                                    curUser.NoHello = True
                                    
                                    'Extract version
235:                                    If m_blnCommaDecimal Then _
                                        dblVersion = StrToDbl(Replace(arrTag(0), ".", ",")) _
                                    Else _
                                        dblVersion = Val(arrTag(0))
                            
240:                                    If g_objSettings.DCMinVersion Then _
                                            If g_objSettings.DCMinVersion > dblVersion Then FailedConf curUser, DCppversion: Exit Function
242:                                End If
                    
                                'Check for exceptions in S:
                                Select Case intID
                                    Case 8
                                        'SdDC++ has the format S:#/# and not S:#
246:                                        lngSlots = GetByte(MidB$(arrTag(3), InStrB(3, arrTag(3), "/") + 2))
                                    Case 3
                                        'DCGUI sometimes has * in it's S: param to denote unlimited slots
                                        'At the moment, I'm content to let them bypass the min slot requirement
249:                                        If AscW(arrTag(3)) = 42 Then _
                                            lngSlots = g_objSettings.MinSlots + 1 _
                                        Else _
                                            lngSlots = CLng(arrTag(3))
                                    Case Else
253:                                        lngSlots = CLng(arrTag(3))
254:                                End Select
                                
                                'If the number of slots is 0, then kick if
                                'validating tags, else ignore and set to 1
                                '(for division purposes)
259:                                If lngSlots = 0 Then
260:                                    If g_objSettings.DCValidateTags Then
261:                                        FailedConf curUser, FakeTag: Exit Function
262:                                    Else
263:                                        lngSlots = 1
264:                                    End If
265:                                End If
                
                                'Check for tag extensions
268:                                If lngUB > 3 Then
269:                                    For lngLoop = 3 To lngUB
                                        'Find out if we support it
                                        Select Case AscW(arrTag(lngLoop))
                                            Case 79 'O
                                                'Format - O:#
                                                
                                                'DC tags (NMDC 2.0) use O: for free/open slots
274:                                                If Not intID = 2 Then
275:                                                    lngO = CLng(MidB$(arrTag(lngLoop), 5))
276:                                                    If lngO > g_objSettings.DCOSpeed Then lngSlots = lngSlots + g_objSettings.DCOSlots
277:                                                End If
                                            Case 76, 66, 85 'L, B, U
                                                'Format - L:#, B:#, U:#
                                        
                                                'Perform bandwidth/slot ratio check
281:                                                If g_objSettings.DCBandPerSlot Then
282:                                                    If intID = 3 Then
                                                        'DCGUI may have * in it's limiter param meaning
                                                        'it is not limiting
285:                                                        arrTag(lngLoop) = MidB$(arrTag(lngLoop), 5)
                                                        
287:                                                        If Not AscW(CStr(arrTag(lngLoop))) = 42 Then
288:                                                                If CLng(arrTag(lngLoop)) < g_objSettings.DCBandPerSlot Then
289:                                                                    FailedConf curUser, BSRatio
290:                                                                    Exit Function
291:                                                                End If
292:                                                            End If
293:                                                    Else
294:                                                        If (CLng(MidB$(arrTag(lngLoop), 5)) / lngSlots) < g_objSettings.DCBandPerSlot Then _
                                                                FailedConf curUser, BSRatio: Exit Function
296:                                                    End If
297:                                                End If
                                            Case 70 'F
                                                'Format - F:#/#
                                        
                                                'Perform bandwidth/slot ratio check after
                                                'extracting upload limit
302:                                                If g_objSettings.DCBandPerSlot Then
303:                                                    If (CLng(MidB$(arrTag(lngLoop), InStrB(1, arrTag(lngLoop), "/") + 2)) / lngSlots) < g_objSettings.DCBandPerSlot Then _
                                                            FailedConf curUser, BSRatio: Exit Function
305:                                                End If
306:                                        End Select
307:                                    Next
308:                                End If
                
                                'Min slot check
311:                                If g_objSettings.MinSlots > lngSlots Then FailedConf curUser, MinSlots: Exit Function
                                
                                
                                'Max slot check
315:                                If g_objSettings.MaxSlots Then
316:                                    If g_objSettings.MaxSlots < lngSlots Then FailedConf curUser, MaxSlots: Exit Function
317:                                End If
                                
                                'Split up the max hubs if using DC++ 0.24 / other clients
320:                                If InStrB(1, arrTag(2), "/") Then
                                    'If using DC++ and if their version is pre 0.24, then they are faking
322:                                    If intID = 1 Then _
                                        If g_objSettings.DCValidateTags Then _
                                            If dblVersion < 0.24 Then _
                                                FailedConf curUser, FakeTag: Exit Function
                                        
327:                                    arrTag = Split(arrTag(2), "/")
                
                                    'H: property MUST have 3 array elements; otherwise they are faking
330:                                    If UBound(arrTag) = 2 Then
                                        'Make sure all items are numerical
                                        Select Case False
                                            Case IsNumeric(arrTag(0)), IsNumeric(arrTag(1)), IsNumeric(arrTag(2)): FailedConf curUser, FakeTag: Exit Function
332:                                        End Select
                                    
                                        'Count up total - included hubs where opped if necessary
335:                                        If g_objSettings.DCIncludeOPed Then _
                                            lngHubs = CLng(arrTag(0)) + CLng(arrTag(1)) + CLng(arrTag(2)) _
                                        Else _
                                            lngHubs = CLng(arrTag(0)) + CLng(arrTag(1))
339:                                    Else
340:                                        If g_objSettings.DCValidateTags Then
341:                                            FailedConf curUser, FakeTag
342:                                            Exit Function
343:                                        End If
344:                                    End If
345:                                Else
                                    'If using DC++ and if their version is post 0.24, they are faking
347:                                    If intID = 1 Then _
                                        If g_objSettings.DCValidateTags Then _
                                            If dblVersion >= 0.24 Then _
                                                FailedConf curUser, FakeTag: Exit Function
                                            
352:                                    lngHubs = CLng(arrTag(2))
353:                                End If
                
                                'If the user is not registered and their hub count is zero, they are faking
                                'If they are registered, this prevents division by 0
357:                                If lngHubs = 0 Then
                                    Select Case True
                                        Case curUser.Class > Normal, g_objSettings.PasswordMode
358:                                            lngHubs = 1
                                        Case Else
                                            'If the user is using QuickList, and they are not registered,
                                            'their hub count has yet to increment, assuming they are logging
                                            'in
362:                                            If curUser.QuickList Then
363:                                                If curUser.State = Logged_In Then _
                                                    If g_objSettings.DCValidateTags Then FailedConf curUser, FakeTag: Exit Function
365:                                            Else
366:                                                If g_objSettings.DCValidateTags Then FailedConf curUser, FakeTag: Exit Function
367:                                            End If
368:                                    End Select
369:                                End If
                
                                'Max hub check
372:                                If g_objSettings.DCMaxHubs Then _
                                        If g_objSettings.DCMaxHubs < lngHubs Then FailedConf curUser, MaxHubs: Exit Function
                
                                'Slots per hub check
376:                                If g_objSettings.DCSlotsPerHub Then _
                                        If (lngSlots / lngHubs) < g_objSettings.DCSlotsPerHub Then _
                                        FailedConf curUser, HSRatio: Exit Function
                                
                                'TheNOP svn 40
                                'Set passive status (if active, then add slots for fake slot check)
                                'If AscW(arrTag(1)) = 65 Then
                                '    curUser.Passive = False
                                'Else
                                '    curUser.Passive = True
                                'End If
387:                            End If
            
389:                            Erase arrTag
390:                        Else
391:                            GoTo NoTag
392:                        End If
393:                    Else
394:                        GoTo NoTag
395:                    End If
396:                Else
397:                    GoTo NoTag
398:                End If
399:            Else
400:
NoTag:
                'If needed, disconnect the user since they have no tag
402:                If g_objSettings.DenyNoTag Then
403:                    FailedConf curUser, NoTag
404:                    Exit Function
405:                Else
                    'Make another MLDC check
407:                    If g_objSettings.AutoKickMLDC Then _
                            If RightB$(arrSplit(0), 26) = "donkey client" Or RightB$(arrSplit(0), 22) = "mldc client" Then _
                            curUser.Kick 60: Exit Function
                            
411:                    If curUser.State = Logged_In Then
412:                        If curUser.NetInfo Then _
                                curUser.SendData "$GetNetInfo|"
414:                    Else
                        'Send GetNetInfo if the user is using NMDC2 (first attempt)
416:                        dblVersion = curUser.iVersion
                        
418:                        If dblVersion = 1.0091 Then
419:                            curUser.SendData "$GetNetInfo|"
420:                        Else
421:                            If dblVersion >= 2 Then _
                                If dblVersion <= 3 Then _
                                    curUser.SendData "$GetNetInfo|"
424:                        End If
425:                    End If
426:                End If
427:            End If
428:        End If
    
        'Check if they are in away mode
        'Select Case AscW(RightB$(arrSplit(2), 2))
        '    Case 2, 3, 6, 7, 10, 11
        '        curUser.isAFK = True
        '    Case Else
        '        curUser.isAFK = False
        'End Select
    
        'Update settings, do not count ChatOnly...
439:        curUser.sMyInfoString = "$MyINFO " & strMyINFO
            
441:        If Not curUser.ChatOnly Then
                'Don't add invisible user's share
443:                If curUser.Visible Then
444:                g_colUsers.iTotalBytesShared = g_colUsers.iTotalBytesShared - curUser.iBytesShared + dblShare
445:                curUser.iBytesShared = dblShare
446:                End If
447:        End If
    
        'Passed all checks
450:        ProcessMyINFO = True
451:    Else
        'Not DC complient
453:        wskLoop_Close curUser.iWinsockIndex
454:    End If
    
456:    Exit Function

458:
Err:
459:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.ProcessMyINFO(, """ & strMyINFO & """)"

    'Something is wrong with their MyINFO string so we disconnect them
462:    wskLoop_Close curUser.iWinsockIndex
End Function

Private Function ValidateNick(ByRef curUser As clsUser, ByRef strName As String, Optional ByRef strMyINFO As String) As Boolean
1:    Dim i           As Integer
2:    Dim objTmp      As Object
3:    Dim objUser     As clsUser

5:    On Error GoTo Err

    'Cannot be longer than 40 chars
8:    If LenB(strName) > 80 Then
9:        curUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("NickLength") & "|"

11:        DoEvents
12:        wskLoop_Close curUser.iWinsockIndex

14:        Exit Function
15:    End If

    'Disallow certain characters, """|'|/|\s"
18:    If g_objRegExps.TestStr(strName, CHRSTODENYINNICK) Then
19:        curUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("ChrInNick") & "|"

21:        DoEvents
22:        wskLoop_Close curUser.iWinsockIndex

24:        Exit Function
25:    End If

27:    On Error Resume Next

    'Copy to sLanguageID property(user language preference)
30:    Set objTmp = m_objPermaCon.Execute("Select UsrStatic.i18n From UsrStatic Where UsrStatic.UserName=" & SQLQuotes(strName) & ";", , 1)
31:    If LenB(objTmp.Collect(0)) Then curUser.sLanguageID = objTmp.Collect(0)
32:    If curUser.sLanguageID = vbNullString Then curUser.sLanguageID = "En"

34:    On Error GoTo Err

    'Find out their registered status
    Select Case g_objRegistered.Registered(strName)
        Case Locked 'Nickname is banned
            'Determine message to send
38:            If g_objSettings.DescriptiveBanMsg Then
39:                curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("IPPermBan") & "|"
40:            Else
41:                curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("IPBanned") & "|"
42:            End If

44:            DoEvents
45:            wskLoop_Close curUser.iWinsockIndex
        Case Unknown 'Not registered
46:            If g_objSettings.PreventSearchBots Then
                'If not registered, then it could be a search tool
48:                If InStrB(1, strName, "search") Then
49:                    wskLoop_Close curUser.iWinsockIndex
50:                    Exit Function
51:                End If
52:            End If

            'If only registered users, get rid of them
55:            If g_objSettings.RegOnly Then
                'Redirect if necessary, otherwise disconnect
57:                If g_objSettings.AutoRedirectNonReg Then
58:                    NextRedirect
59:                    curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RegOnlyRedirTo") & g_objSettings.RedirectIP & "|$ForceMove " & g_objSettings.RedirectIP & "|"

61:                    DoEvents
62:                    wskLoop_Close curUser.iWinsockIndex
63:                Else
64:                    curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RegOnly") & "|"

66:                    DoEvents
67:                    wskLoop_Close curUser.iWinsockIndex
68:                End If

70:                Exit Function
71:            Else
                'If redirecting only non-registered or non-opped users if full, make the check
73:                If g_objSettings.AutoRedirectFullNonOps Or g_objSettings.AutoRedirectFullNonReg Then
74:                        If g_colUsers.Count >= g_objSettings.MaxUsers Then
75:                            NextRedirect

77:                            curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("FullRedirTo") & g_objSettings.RedirectIP & "|$ForceMove " & g_objSettings.RedirectIP & "|"

79:                            DoEvents
80:                            wskLoop_Close curUser.iWinsockIndex

82:                            Exit Function
83:                        End If
84:                End If

                'Check if their nickname is in use
87:                If g_colUsers.Online(strName) Then
88:                    i = g_colUsers.ItemByName(strName).iWinsockIndex

                    'If it is, then check if the user's winsock is closed
91:                    If wskLoop(i).State = 0 Then
92:                        wskLoop_Close i
93:                    Else
                        'If it is still open, then compare their IPs
                        'If they are the same, disconnect the ghost
96:                        If wskLoop(i).RemoteHostIP = curUser.IP Then
97:                            wskLoop_Close i
98:                        Else
99:                            curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("NickTaken") & "|$ValidateDenide|"

101:                            DoEvents
102:                            wskLoop_Close curUser.iWinsockIndex

104:                            Exit Function
105:                        End If
106:                    End If
107:                End If

109:                If LenB(strMyINFO) Then
110:                    If ProcessMyINFO(curUser, strMyINFO) Then
                        'Since we have their MyINFO string, they must be QuickList
                        'They are not registered so that means they are now fully logged in
                        'unless the hub is running in password mode
114:                        curUser.sName = strName
115:                        g_colUsers.UpdateName curUser

                        'If the hub is using password mode, then ask for it
118:                        If g_objSettings.PasswordMode Then
119:                            curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("PassMode") & "|$GetPass|"

121:                            curUser.State = Wait_PassPM
122:                        Else
123:                            curUser.Class = Normal

                            'Send hub message and hub name
126:                            curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|"

128:                            curUser.QNL = True
129:                            g_colUsers.UpdateLogIn curUser
130:                            SEvent_UserConnected curUser
131:                        End If
132:                    Else
133:                        Exit Function
134:                    End If
135:                Else
                    'Add to user name collection
137:                    curUser.sName = strName
138:                    g_colUsers.UpdateName curUser

                    'If the hub is using password mode, ask for the password, else
                    'wait for their MyINFO string
142:                    If g_objSettings.PasswordMode Then
143:                        curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("PassMode") & "|$GetPass|"

145:                        curUser.State = Wait_PassPM
146:                    Else
147:                        curUser.State = Wait_Info
148:                        curUser.Class = Normal
                            'Send welcome message / hub name / $Hello
150:                        curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|$Hello " & strName & "|"

152:                    End If
153:                End If
154:            End If
        Case Mentored, Invisible, Registered, Vip 'Registered - Non op
            'Redirect if necessary
156:            If g_objSettings.AutoRedirectFullNonOps Then
157:                If g_colUsers.Count >= g_objSettings.MaxUsers Then
158:                    NextRedirect
159:                    curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("FullRedirTo") & g_objSettings.RedirectIP & "|$ForceMove " & g_objSettings.RedirectIP & "|"

161:                    DoEvents
162:                    wskLoop_Close curUser.iWinsockIndex

164:                    Exit Function
165:                End If
166:            End If

            Select Case g_colUsers.Online(strName)
                Case 0
                Case -1
                'm_colNames
                    'i = g_colUsers.ItemByName(strName).iWinsockIndex
                    'If Not i = curUser.iWinsockIndex Then
                    '    If wskLoop(i).State = 0 Then wskLoop_Close i
                    'End If
                    
174:                    For Each objUser In g_colUsers
175:                        If objUser.sName = strName Then
176:                            i = objUser.iWinsockIndex
177:                            If Not i = curUser.iWinsockIndex Then
178:                                If wskLoop(i).State = 0 Then wskLoop_Close i
179:                            End If
180:                        End If
181:                    Next
                Case 1
                'm_colNLoggingIn
183:                    For Each objUser In g_colUsers
184:                        If objUser.sName = strName Then
185:                            If Not objUser.iWinsockIndex = curUser.iWinsockIndex Then
186:                                wskLoop_Close objUser.iWinsockIndex
187:                            End If
188:                        End If
189:                    Next
190:            End Select

192:            curUser.sName = strName
193:            g_colUsers.UpdateName curUser

            'Set MyINFO string if they are QuickList to check after log in
196:            If LenB(strMyINFO) Then curUser.sMyInfoString = strMyINFO

            'Send welcome / password request
199:            curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RegPass") & "|$GetPass|"

201:            curUser.State = Wait_Pass
        Case Else 'Registered - Op
            'If their nickname is taken, either disconnect currently logging in
            'user, or wait until they do a password check (if it isn't a ghost)
            Select Case g_colUsers.Online(strName)
                Case 0
                Case -1
                    'i = g_colUsers.ItemByName(strName).iWinsockIndex
                    'If Not i = curUser.iWinsockIndex Then
                    '    If wskLoop(i).State = 0 Then wskLoop_Close i
                    'End If
208:                    For Each objUser In g_colUsers
209:                        If objUser.sName = strName Then
210:                            i = objUser.iWinsockIndex
211:                            If Not i = curUser.iWinsockIndex Then
212:                                If wskLoop(i).State = 0 Then wskLoop_Close i
213:                            End If
214:                        End If
215:                    Next

                Case 1
217:                    For Each objUser In g_colUsers
218:                        If objUser.sName = strName Then
219:                            If Not objUser.iWinsockIndex = curUser.iWinsockIndex Then
220:                                wskLoop_Close objUser.iWinsockIndex
221:                            End If
222:                        End If
223:                    Next

225:            End Select

227:            curUser.sName = strName
228:            g_colUsers.UpdateName curUser

            'Set MyINFO string if they are QuickList to check after log in
231:            If LenB(strMyINFO) Then curUser.sMyInfoString = strMyINFO

            'Send welcome / password request
234:            curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RegPass") & "|$GetPass|"

236:            curUser.State = Wait_Pass
237:    End Select

239:    ValidateNick = True

241:    Exit Function

243:
Err:
244:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.ValidateNick(, """ & strName & """, """ & strMyINFO & """)"

    'Something is wrong with their nickname/etc so we disconnect them
247:    wskLoop_Close curUser.iWinsockIndex
End Function

Friend Sub ProcessTrigger(ByRef objUser As clsUser, ByRef strTrigger As String, ByRef blnMainChat As Boolean)
1:    Dim arrCommand()        As String
2:    Dim strIP               As String
3:    Dim strMsg              As String
4:    Dim arrTmp()            As String
5:    Dim strTmp              As String
6:    Dim lngTmp              As Long
7:    Dim intTmp              As Integer
8:    Dim objTmp              As Object
9:    Dim varTmp              As Variant
10:   Dim objCommand          As clsCommand

12:    On Error GoTo Err

14:    arrCommand = Split(strTrigger, g_objSettings.CSeperator, 3)

16:    If UBound(arrCommand) = -1 Then Exit Sub

18:    On Error Resume Next
19:    Set objCommand = g_colCommands(arrCommand(0))
20:    On Error GoTo Err

22:    If ObjPtr(objCommand) Then
        'Commands and their corresponding ID

        'reg = 1
        'admin = 2
        'ban = 3
        'banip = 4
        'banuser = 5
        'close = 6
        'info = 7
        'iplist = 8
        'listbanip = 9
        'ipscan = 10
        'listbanuser = 11
        'unbanip = 12
        'unbanuser = 13
        'help = 14
      
        'Make sure command is enabled
41:        If Not objCommand.Enabled Then Exit Sub
      
        'Make sure the user has permission to use the command
44:     If objUser.Class < objCommand.Class Then Exit Sub
      
        Select Case objCommand.ID
            'Case 1
            'Case 2
            'Case 3
            'Case 4
            'Case 5
            'Case 6
            'Case 7
            'Case 8
            'Case 9
            'Case 10
            'Case 11
            'Case 12
            'Case 13
            'Case 14
                
            Case Else
61:                If objCommand.ID > 50 Then SEvent_CustComArrival objUser, objCommand, strTrigger, blnMainChat
62:        End Select
  
    
65:    End If

67:    Exit Sub
  
69:
Err:
70:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.ProcessTrigger()"
End Sub

Private Sub Notify(ByRef strMessage As String)
1:    On Error GoTo Err

    'This will be changed later

5:    For Each m_objLoopUser In g_colUsers
6:        If m_objLoopUser.Class > InvisibleSuperOp Then m_objLoopUser.SendPrivate g_objSettings.BotName, strMessage
7:    Next
  
9:    Set m_objLoopUser = Nothing
  
11:    Exit Sub

13:
Err:
14:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.Notify()"
15:    Set m_objLoopUser = Nothing
End Sub

Private Function StrToDbl(ByVal strConvert As String) As Double
1:    Dim arrString()     As Byte
2:    Dim lngLoop         As Long
3:    Dim lngUB           As Long
4:    Dim blnDecimal      As Boolean
5:    Dim bytDecimal      As Byte
    
7:    On Error GoTo Err
    
    'Converts a string to a double value
    'This only works for comma decimal based systems (swap 46 and 44 for a
    'period decimal system; however you should use Val() for those systems)
    
    'Replace useless , or . (whichever is not the decimal)
14:    If m_blnCommaDecimal Then
15:        strConvert = Replace(strConvert, ".", vbNullString)
16:        bytDecimal = 44
17:    Else
18:        strConvert = Replace(strConvert, ",", vbNullString)
19:        bytDecimal = 46
20:    End If
    
22:    lngUB = LenB(strConvert) - 1
    
    'Make sure it isn't a zero length string
25:    If Not lngUB = -1 Then
        'Copy into array
27:        ReDim arrString(0 To lngUB) As Byte
28:        CopyMemory arrString(0), ByVal StrPtr(strConvert), lngUB
        
        'Loop through and find first non numeric char
31:        For lngLoop = 0 To lngUB Step 2
            Select Case arrString(lngLoop)
                Case 48 To 57
                Case bytDecimal: If blnDecimal Then lngLoop = lngLoop - 2: Exit For Else blnDecimal = True
                Case Else: Exit For
32:            End Select
33:        Next
34:    End If
    
    'If it wasn't the first character, then convert numerical characters to string
37:    If lngLoop Then StrToDbl = CDbl(LeftB$(strConvert, lngLoop))
    
39:    Exit Function
    
41:
Err:
42:    HandleError Err.Number, Err.Description, Erl & "|" & "frmMain.StrToDbl(" & strConvert & ")", Err.LastDllError
End Function

'------------------------------------------------------------------------------
' Setting related methods
'------------------------------------------------------------------------------
Public Sub LoadDefaultSettings()

2:      Dim lngLoop     As Long
3:        On Error GoTo Err

    #If FLASHCHAT Then
6:      Dim objTag(10)   As New clsTag
    #Else
8:      Dim objTag(9)   As New clsTag
    #End If
    
    
    'pre-defined TagsHelp
13:    m_arrTagRules(0) = "'NoHello' even if it's not in $Supports statement.%[LF]%[LF] Tests for minimum DC++ version V:______"
14:    m_arrTagRules(1) = "Skips standard DC++ O:# tests.%[LF]%[LF] O:# is used for free/open slots."
15:    m_arrTagRules(2) = "* in it's slot param (S:*) means unlimited slots.%[LF]%[LF] * in it's limiter param (L:*) means unlimited bandwidth.%[LF]%[LF] Reports bandwidth limit on a per slot basis, not total."
16:    m_arrTagRules(3) = "None"
17:    m_arrTagRules(4) = "None"
18:    m_arrTagRules(5) = "Uses F:#Down/#Up to report bandwidth limiting."
19:    m_arrTagRules(6) = "None"
20:    m_arrTagRules(7) = "None"
21:    m_arrTagRules(8) = "Slot param has the format S:#/#"
22:    m_arrTagRules(9) = "If you are using this option then you can figure it out for yourself."
23:    m_arrTagRules(10) = "Select a Default Tag to see if it has any special processing rules."
24:    m_arrTagRules(11) = "None"
    
26:    Set m_colTags = New Collection
27:    lstTagsDef.Clear
    
29:    objTag(0).Name = "++"
30:    objTag(0).ID = 1
31:    objTag(1).Name = "DC"
32:    objTag(1).ID = 2
33:    objTag(2).Name = "DCGUI"
34:    objTag(2).ID = 3
35:    objTag(3).Name = "oDC"
36:    objTag(3).ID = 4
37:    objTag(4).Name = "QuickDC"
38:    objTag(4).ID = 5
39:    objTag(5).Name = "DC:Pro"
40:    objTag(5).ID = 6
41:    objTag(6).Name = "SDC"
42:    objTag(6).ID = 7
43:    objTag(7).Name = "StrgDC++"
44:    objTag(7).ID = 10
45:    objTag(8).Name = "SdDC++"
46:    objTag(8).ID = 8
47:    objTag(9).Name = "Z++"
48:    objTag(9).ID = 11

    #If FLASHCHAT Then
51:    objTag(10).Name = "Chat"
52:    objTag(10).ID = 9
    #End If
    
    #If FLASHCHAT Then
56:    For lngLoop = 0 To 10
    #Else
58:    For lngLoop = 0 To 9
    #End If
60:        m_colTags.Add objTag(lngLoop), objTag(lngLoop).Name
61:        lstTagsDef.AddItem objTag(lngLoop).Name
62:    Next

64:    Call LoadDfsSettings
       
66:    DoEvents

68:  Exit Sub

70:
Err:
71:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.LoadDefaultSettings()"
End Sub
Public Sub LoadSettings()
1:     Dim objXML          As clsXMLParser
2:     Dim objNode         As clsXMLNode
3:     Dim objSubNode      As clsXMLNode
4:     Dim colNodes        As Collection
5:     Dim colSubNodes     As Collection
6:     Dim colAttributes   As Collection
7:     Dim colSupported    As Collection
8:     Dim m_colLangString As Collection
9:     Dim lvwItem         As ListItem
10:    Dim lvwItems        As ListItems
11:    Dim objTag          As clsTag
12:    Dim lngLoop         As Long
13:    Dim strTemp         As String
14:    Dim strATemp        As String
15:    Dim strSettVer      As String
17:    Dim X               As Integer
18:    Dim objCmd          As clsCommand
19:    Dim strIPBans(4)    As String
20:    On Error GoTo Err
    
22:    If g_objFileAccess.FileExists(G_APPPATH & "\PTDCH.xml") Then
23:        g_objFileAccess.CopyFile G_APPPATH & "\XML.xml", G_APPPATH & "\Settings\PTDCH.xml"
24:        g_objFileAccess.CopyFile G_APPPATH & "\Commands.xml", G_APPPATH & "\Settings\Commands.xml"
25:        g_objFileAccess.CopyFile G_APPPATH & "\PermIPBans.xml", G_APPPATH & "\Settings\PermIPBans.xml"
26:        g_objFileAccess.CopyFile G_APPPATH & "\TempIPBans.xml", G_APPPATH & "\Settings\TempIPBans.xml"
27:        g_objFileAccess.CopyFile G_APPPATH & "\DefaultProps.xml", G_APPPATH & "\Settings\DefaultProps.xml"
        'These are not parts of any previously released hubsoft. but just in case...
29:        g_objFileAccess.DeleteFile G_APPPATH & "\*.xml"
30:    End If

32:    Set objXML = New clsXMLParser

'---------------------------------------------------------------------------------
    'Load regular settings
36:    strTemp = G_APPPATH & "\Settings\PTDCH.xml"
    
38:    If g_objFileAccess.FileExists(strTemp) Then
39:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
40:        objXML.Parse
    
42:        Set colNodes = objXML.Nodes(1).Nodes
    
        'Just in case...
        'On Error Resume Next
        'Set the Version from the Settings File.
47:        strSettVer = objXML.Nodes(1).Attributes("Version").Value
    
49:        For Each objNode In colNodes
50:            Set colSubNodes = objNode.Nodes
            'Using the CallByName sub may be a bit slower, but it is VERY convient
            'The settings parser no longer needs mothering! Woot!
            Select Case objNode.Name
                Case "Long"
53:                    For Each objSubNode In colSubNodes
54:                        CallByName g_objSettings, objSubNode.Name, VbLet, CLng(objSubNode.Value)
55:                    Next
                Case "Integer"
56:                    For Each objSubNode In colSubNodes
57:                        CallByName g_objSettings, objSubNode.Name, VbLet, CInt(objSubNode.Value)
58:                    Next
                Case "Boolean"
59:                    For Each objSubNode In colSubNodes
60:                        CallByName g_objSettings, objSubNode.Name, VbLet, CBool(objSubNode.Value)
61:                    Next
                Case "Double"
62:                    For Each objSubNode In colSubNodes
63:                        CallByName g_objSettings, objSubNode.Name, VbLet, CDbl(objSubNode.Value)
64:                    Next
                Case "String"
65:                    For Each objSubNode In colSubNodes
66:                        CallByName g_objSettings, objSubNode.Name, VbLet, objSubNode.Value
67:                    Next
                Case "Byte"
68:                    For Each objSubNode In colSubNodes
69:                        CallByName g_objSettings, objSubNode.Name, VbLet, CByte(objSubNode.Value)
70:                    Next
                Case "Tags"
                    ' If we have a Settings Version then all accepted Tags are saved
                    ' so clear Collection created in LoadDefaultSettings.
73:                    If strSettVer >= "0.1.1" Then Set m_colTags = New Collection
74:                    For Each objSubNode In colSubNodes
75:                        Set objTag = New clsTag
76:                        Set colAttributes = objSubNode.Attributes
    
78:                       objTag.Name = colAttributes("Name").Value
                        ' If the loaded Tag is one of the defaults give it the right default ID
                           Select Case objTag.Name
                            Case "++": objTag.ID = 1
                            Case "DC": objTag.ID = 2
                            Case "DCGUI": objTag.ID = 3
                            Case "oDC": objTag.ID = 4
                            Case "QuickDC": objTag.ID = 5
                            Case "DC:Pro": objTag.ID = 6
                            Case "SDC": objTag.ID = 7
                           Case "StrgDC++": objTag.ID = 10
                           Case "SdDC++": objTag.ID = 8
                           Case "Z++": objTag.ID = 11
                           Case "Chat": objTag.ID = 9
                           Case Else: objTag.ID = -1
80:                      End Select
    
82:                      m_colTags.Add objTag, objTag.Name
83:                  Next
84:            End Select
85:        Next
    
87:        On Error GoTo Err
     
          'Set min share value
90:        g_objSettings.MinShare = g_objSettings.IMinShare * (1024 ^ g_objSettings.MinShareSize)
    
92:        objXML.Clear
    
94:        Set objSubNode = Nothing
95:        Set objNode = Nothing
96:        Set colSubNodes = Nothing
97:        Set colNodes = Nothing
98:     End If


'---------------------------------------------------------------------------------
    'Load perm IP bans
103:    strTemp = G_APPPATH & "\Settings\PermIPBans.xml"
        
105:    If g_objFileAccess.FileExists(strTemp) Then
106:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
107:        objXML.Parse
108:        Set colNodes = objXML.Nodes(1).Nodes
            'Make sure ban list is cleared if we are reloading settings
110:        g_colIPBans.ClearPerm
           'Just in case...
           'On Error Resume Next
113:        For Each objNode In colNodes
114:            If objNode.Name = "PermIPBans" Then
115:                Set colSubNodes = objNode.Nodes
116:                For Each objSubNode In colSubNodes
                    Select Case objSubNode.Name
                        Case "IP"
117:                            strIPBans(0) = CStr(objSubNode.Value)
                        Case "Nick"
118:                            strIPBans(1) = CStr(objSubNode.Value)
                        Case "BannedBy"
119:                            strIPBans(2) = CStr(objSubNode.Value)
                        Case "Reason"
120:                            strIPBans(3) = CStr(objSubNode.Value)
121:                    End Select
122:                Next
123:                If Not strIPBans(0) = "" Then
124:                     g_colIPBans.Add strIPBans(0), -1, strIPBans(1), strIPBans(2), strIPBans(3), True
125:                End If
126:                Erase strIPBans
127:            End If
128:        Next
129:        On Error GoTo Err
    
131:        objXML.Clear

133:        Set objNode = Nothing
134:        Set colNodes = Nothing
135:        Set colSubNodes = Nothing
136:    End If

'---------------------------------------------------------------------------------
        'Make sure ban list is cleared if we are reloading settings
140:    g_colIPBans.ClearTemp

        'Load temp IP bans
143:    strTemp = G_APPPATH & "\Settings\TempIPBans.xml"

145:    If g_objFileAccess.FileExists(strTemp) Then
146:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
147:        objXML.Parse
148:        Set colNodes = objXML.Nodes(1).Nodes
           'Just in case...
           'On Error Resume Next
151:        For Each objNode In colNodes
152:            If objNode.Name = "TempIPBans" Then
153:                Set colSubNodes = objNode.Nodes
154:                For Each objSubNode In colSubNodes
                        Select Case objSubNode.Name
                            Case "IP"
155:                            strIPBans(0) = CStr(objSubNode.Value)
                            Case "ExpDate"
156:                            strIPBans(1) = CStr(objSubNode.Value)
                            Case "Nick"
157:                            strIPBans(2) = CStr(objSubNode.Value)
                            Case "BannedBy"
158:                            strIPBans(3) = CStr(objSubNode.Value)
                            Case "Reason"
159:                            strIPBans(4) = CStr(objSubNode.Value)
160:                    End Select
161:                Next
162:                If Not strIPBans(0) = "" Or Not strIPBans(1) = "" Then
163:                    lngLoop = DateDiff("n", Now, CDate(strIPBans(1)))
                        'Make sure the date hasn't expired
165:                    If lngLoop > 0 Then _
                            g_colIPBans.Add strIPBans(0), lngLoop, strIPBans(2), strIPBans(3), strIPBans(4), True
167:                End If
168:                Erase strIPBans
169:            End If
170:        Next
171:        On Error GoTo Err
    
173:        objXML.Clear

175:        Set objNode = Nothing
176:        Set colNodes = Nothing
177:        Set colSubNodes = Nothing
178:    End If
'---------------------------------------------------------------------------------
    'Commands
181:    strTemp = G_APPPATH & "\Settings\Commands.xml"

183:    If g_objFileAccess.FileExists(strTemp) Then
        'Clear old commands / add defaults
185:        g_colCommands.Clear
        'shouldn't this be remove ??? except language.
        'g_colCommands.Add 1, "reg", "The register command panel", Admin, True
        'g_colCommands.Add 2, "admin", "The admin command panel", Admin, True
        'g_colCommands.Add 3, "ban", "Bans (aka locks) a username", SuperOp, True
        'g_colCommands.Add 4, "banip", "Bans an IP", SuperOp, True
        'g_colCommands.Add 5, "banuser", "Disconnects and perm bans a user (by IP)", SuperOp, True
        'g_colCommands.Add 6, "close", "Disconnects a user", Op, True
        'g_colCommands.Add 7, "info", "Retrieves information on about user", Op, True
        'g_colCommands.Add 8, "iplist", "Lists the IPs / Names of connected users", Op, True
        'g_colCommands.Add 9, "listbanip", "Lists the IPs currently banned", SuperOp, True
        'g_colCommands.Add 10, "ipscan", "Checks for users who are connected more than once (on the same IP)", SuperOp, True
        'g_colCommands.Add 11, "listbanuser", "Lists the banned (aka locked) usernames", SuperOp, True
        'g_colCommands.Add 12, "unbanip", "Unban an IP", SuperOp, True
        'g_colCommands.Add 13, "unbanuser", "Unban (aka unlock) a username", SuperOp, True
        'g_colCommands.Add 14, "help", "Description is created by all of the other commands.", Op, True
        'g_colCommands.Add 15, "language", "Changes language preference for scripts which have multi-language support.", Mentored, True
 
203:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
204:        objXML.Parse
    
        'On Error Resume Next
    
208:        Set colNodes = objXML.Nodes(1).Nodes
    
210:        For Each objNode In colNodes
211:            Set colSubNodes = objNode.Attributes
212:            g_colCommands.Add CInt(colSubNodes("ID").Value), colSubNodes("Trigger").Value, colSubNodes("Description").Value, CInt(colSubNodes("Class").Value), CBool(colSubNodes("Enabled").Value)
213:        Next
    
215:        On Error GoTo Err
    
217:        objXML.Clear
    
219:        Set objNode = Nothing
220:        Set colSubNodes = Nothing
221:        Set colNodes = Nothing

        '-----------------------------------------
        ' Unload the default Commands by ID in case some names are re-used.
        ' removing commands by name would render changes made in gui useless
        '-----------------------------------------
227:        For Each objCmd In g_colCommands
228:            If objCmd.ID < 51 Then g_colCommands.Remove (objCmd.Name)
229:        Next
        
231:        Set objCmd = Nothing
        'If g_colCommands.Exists("ban") Then g_colCommands.Remove ("ban")
        'If g_colCommands.Exists("banip") Then g_colCommands.Remove ("banip")
        'If g_colCommands.Exists("banuser") Then g_colCommands.Remove ("banuser")
        'If g_colCommands.Exists("close") Then g_colCommands.Remove ("close")
        'If g_colCommands.Exists("info") Then g_colCommands.Remove ("info")
        'If g_colCommands.Exists("iplist") Then g_colCommands.Remove ("iplist")
        'If g_colCommands.Exists("listbanip") Then g_colCommands.Remove ("listbanip")
        'If g_colCommands.Exists("ipscan") Then g_colCommands.Remove ("ipscan")
        'If g_colCommands.Exists("listbanuser") Then g_colCommands.Remove ("listbanuser")
        'If g_colCommands.Exists("unbanip") Then g_colCommands.Remove ("unbanip")
        'If g_colCommands.Exists("unbanuser") Then g_colCommands.Remove ("unbanuser")
        'If g_colCommands.Exists("help") Then g_colCommands.Remove ("help")
        'If g_colCommands.Exists("language") Then g_colCommands.Remove ("language")
  
246:    End If

'---------------------------------------------------------------------------------

250:    m_arrDynaCap(0) = "Start Serving"
251:    m_arrDynaCap(1) = "Stop Serving"

'---------------------------------------------------------------------------------

'---------------------------------------------------------------------------------
    'Load Core/Reason Messages, a default must exist restart the hub if it don't
257:    strTemp = G_APPPATH & "\Settings\UsersMessages.xml"

259:    If Not g_objFileAccess.FileExists(strTemp) Then
               'Create defaut users messages file .. if is not found
261:           LoadAndSaveXML enuXML.EGUsersMessages
262:    End If

264:    objXML.Data = g_objFileAccess.ReadFile(strTemp)
265:    objXML.Parse

267:    Set colNodes = objXML.Nodes(1).Nodes

    'Just in case...
270:     On Error Resume Next

272:    Set colSupported = New Collection

274:    For Each objNode In colNodes
275:        Set colSubNodes = objNode.Nodes

277:        Set m_colLangString = New Collection

279:        For Each objSubNode In colSubNodes
280:            m_colLangString.Add objSubNode.Value, objSubNode.Name
281:        Next

283:        colSupported.Add objNode.Name, objNode.Name
284:        g_colLanguages.Add m_colLangString, objNode.Name

286:        Set m_colLangString = Nothing
287:    Next

289:    g_colLanguages.Add colSupported, "Supported"

291:    On Error GoTo Err

293:    objXML.Clear

295:    Set colSupported = Nothing
296:    Set objSubNode = Nothing
297:    Set objNode = Nothing
298:    Set colSubNodes = Nothing
299:    Set colNodes = Nothing

        'Scann Language files..
398:    Call LoadLanguagesFiles
399:    Call SetLanguageInterface(g_objSettings.Interface)

412:    Exit Sub
413:
Err:
414:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.LoadSettings()|" & strTemp & "|"
415:    Resume Next
End Sub
Public Sub SaveSettings()
1:    Dim intFF       As Integer
2:    Dim strTemp     As String
3:    Dim varLoop     As Variant
4:    Dim objCommand  As clsCommand
5:    Dim objTag      As clsTag
6:    Dim objTB       As clsIPBansData
7:    Dim i           As Integer
8:    Dim lvwItems    As ListItems
     'Now before I get any emails about all the file appending, read this
     'Using string concation (&) is several times slower than using this append method
     '& reallocates the string in the memory each time it's used, while this method does not (well on a much smaller scale)
12:    On Error GoTo Err

14:    strTemp = G_APPPATH & "\Settings\PTDCH.xml"

    'If the settings file exists, delete it
17:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp

19:    Call SaveFormSize ' save form position now

21:    intFF = FreeFile
    'Append to PTDCH.xml
23:    Open strTemp For Append As intFF

25:    Print #intFF, "<Settings Version=""" & vbVersion & """>"
26:    Print #intFF, vbTab & "<String>"
27:        Print #intFF, vbTab & vbTab & "<frmHubPosition>" & g_objSettings.frmHubPosition & "</frmHubPosition>"
28:        Print #intFF, vbTab & vbTab & "<HubName>" & XMLEscape(g_objSettings.HubName) & "</HubName>"
29:        Print #intFF, vbTab & vbTab & "<HubDesc>" & XMLEscape(g_objSettings.HubDesc) & "</HubDesc>"
30:        Print #intFF, vbTab & vbTab & "<HubIP>" & g_objSettings.HubIP & "</HubIP>"
31:        Print #intFF, vbTab & vbTab & "<Ports>" & g_objSettings.Ports & "</Ports>"
32:        Print #intFF, vbTab & vbTab & "<RegisterIP>" & g_objSettings.RegisterIP & "</RegisterIP>"
        '-------------NEW REDIRECT ADDRESSES-----------------------------------------------------------------------------------------------
34:        Print #intFF, vbTab & vbTab & "<ForMinShareRedirectAddress>" & g_objSettings.ForMinShareRedirectAddress & "</ForMinShareRedirectAddress>"
35:        Print #intFF, vbTab & vbTab & "<ForMaxShareRedirectAddress>" & g_objSettings.ForMaxShareRedirectAddress & "</ForMaxShareRedirectAddress>"
36:        Print #intFF, vbTab & vbTab & "<ForMinSlotsRedirectAddress>" & g_objSettings.ForMinSlotsRedirectAddress & "</ForMinSlotsRedirectAddress>"
37:        Print #intFF, vbTab & vbTab & "<ForMaxSlotsRedirectAddress>" & g_objSettings.ForMaxSlotsRedirectAddress & "</ForMaxSlotsRedirectAddress>"
38:        Print #intFF, vbTab & vbTab & "<ForMaxHubsRedirectAddress>" & g_objSettings.ForMaxHubsRedirectAddress & "</ForMaxHubsRedirectAddress>"
39:        Print #intFF, vbTab & vbTab & "<ForNoTagRedirectAddress>" & g_objSettings.ForNoTagRedirectAddress & "</ForNoTagRedirectAddress>"
40:        Print #intFF, vbTab & vbTab & "<ForTooOldDcppRedirectAddress>" & g_objSettings.ForTooOldDcppRedirectAddress & "</ForTooOldDcppRedirectAddress>"
41:        Print #intFF, vbTab & vbTab & "<ForTooOldNMDCRedirectAddress>" & g_objSettings.ForTooOldNMDCRedirectAddress & "</ForTooOldNMDCRedirectAddress>"
42:        Print #intFF, vbTab & vbTab & "<ForSlotPerHubRedirectAddress>" & g_objSettings.ForSlotPerHubRedirectAddress & "</ForSlotPerHubRedirectAddress>"
43:        Print #intFF, vbTab & vbTab & "<ForBWPerSlotRedirectAddress>" & g_objSettings.ForBWPerSlotRedirectAddress & "</ForBWPerSlotRedirectAddress>"
44:        Print #intFF, vbTab & vbTab & "<ForFakeShareRedirectAddress>" & g_objSettings.ForFakeShareRedirectAddress & "</ForFakeShareRedirectAddress>"
45:        Print #intFF, vbTab & vbTab & "<ForFakeTagRedirectAddress>" & g_objSettings.ForFakeTagRedirectAddress & "</ForFakeTagRedirectAddress>"
46:        Print #intFF, vbTab & vbTab & "<ForPasModeRedirectAddress>" & g_objSettings.ForPasModeRedirectAddress & "</ForPasModeRedirectAddress>"
        '----------------STOP HERE----------------------------------------------------------------------------------------------------------
48:        Print #intFF, vbTab & vbTab & "<RedirectAddress>" & g_objSettings.RedirectAddress & "</RedirectAddress>"
49:        Print #intFF, vbTab & vbTab & "<BotName>" & g_objSettings.BotName & "</BotName>"
50:        Print #intFF, vbTab & vbTab & "<BotEmail>" & g_objSettings.BotEmail & "</BotEmail>"
51:        Print #intFF, vbTab & vbTab & "<CSeperator>" & g_objSettings.CSeperator & "</CSeperator>"
52:        Print #intFF, vbTab & vbTab & "<HubPassword>" & XMLEscape(g_objSettings.HubPassword) & "</HubPassword>"
53:        Print #intFF, vbTab & vbTab & "<MassMessage>" & XMLEscape(g_objSettings.MassMessage) & "</MassMessage>"
54:        Print #intFF, vbTab & vbTab & "<OpMassMessage>" & XMLEscape(g_objSettings.OpMassMessage) & "</OpMassMessage>"
55:        Print #intFF, vbTab & vbTab & "<UnRegMassMessage>" & XMLEscape(g_objSettings.UnRegMassMessage) & "</UnRegMassMessage>"
        ' ------------------------ NEW INTERFACE LANGUAGE ------------------------
57:        Print #intFF, vbTab & vbTab & "<Interface>" & g_objSettings.Interface & "</Interface>"
        ' ----------------------- NEW INTERFACE LANGUAGE END ----------------------
59:        Print #intFF, vbTab & vbTab & "<HammeringRd>" & g_objSettings.HammeringRd & "</HammeringRd>"
60:        Print #intFF, vbTab & vbTab & "<NoIPDNS1>" & g_objSettings.NoIPDNS1 & "</NoIPDNS1>"
61:        Print #intFF, vbTab & vbTab & "<NoIPDNS2>" & g_objSettings.NoIPDNS2 & "</NoIPDNS2>"
62:        Print #intFF, vbTab & vbTab & "<NoIPDNS3>" & g_objSettings.NoIPDNS3 & "</NoIPDNS3>"
63:        Print #intFF, vbTab & vbTab & "<NoIPDNS4>" & g_objSettings.NoIPDNS4 & "</NoIPDNS4>"
64:        Print #intFF, vbTab & vbTab & "<NoIPUser>" & XMLEscape(g_objSettings.NoIPUser) & "</NoIPUser>"
65:        Print #intFF, vbTab & vbTab & "<NoIPPass>" & XMLEscape(g_objSettings.NoIPPass) & "</NoIPPass>"
66:        Print #intFF, vbTab & vbTab & "<DynDNS1>" & g_objSettings.DynDNS1 & "</DynDNS1>"
67:        Print #intFF, vbTab & vbTab & "<DynDNS2>" & g_objSettings.DynDNS2 & "</DynDNS2>"
68:        Print #intFF, vbTab & vbTab & "<DynDNS3>" & g_objSettings.DynDNS3 & "</DynDNS3>"
69:        Print #intFF, vbTab & vbTab & "<DynDNS4>" & g_objSettings.DynDNS4 & "</DynDNS4>"
70:        Print #intFF, vbTab & vbTab & "<DynDNSUser>" & XMLEscape(g_objSettings.DynDNSUser) & "</DynDNSUser>"
71:        Print #intFF, vbTab & vbTab & "<DynDNSPass>" & XMLEscape(g_objSettings.DynDNSPass) & "</DynDNSPass>"
        'MySQL interface
73:        Print #intFF, vbTab & vbTab & "<DBName>" & XMLEscape(g_objSettings.DBName) & "</DBName>"
74:        Print #intFF, vbTab & vbTab & "<DBUserName>" & XMLEscape(g_objSettings.DBUserName) & "</DBUserName>"
75:        Print #intFF, vbTab & vbTab & "<DBPassword>" & XMLEscape(g_objSettings.DBPassword) & "</DBPassword>"
76:        Print #intFF, vbTab & vbTab & "<DBServerAddresse>" & XMLEscape(g_objSettings.DBServerAddresse) & "</DBServerAddresse>"
77:    Print #intFF, vbTab & "</String>"

79:    Print #intFF, vbTab & "<Boolean>"
80:        Print #intFF, vbTab & vbTab & "<DenySocks5>" & g_objSettings.DenySocks5 & "</DenySocks5>"
81:        Print #intFF, vbTab & vbTab & "<DenyPassive>" & g_objSettings.DenyPassive & "</DenyPassive>"
82:        Print #intFF, vbTab & vbTab & "<AutoCheckUpdate>" & g_objSettings.AutoCheckUpdate & "</AutoCheckUpdate>"
83:        Print #intFF, vbTab & vbTab & "<AutoKickMLDC>" & g_objSettings.AutoKickMLDC & "</AutoKickMLDC>"
84:        Print #intFF, vbTab & vbTab & "<AutoRegister>" & g_objSettings.AutoRegister & "</AutoRegister>"
85:        Print #intFF, vbTab & vbTab & "<AutoRedirect>" & g_objSettings.AutoRedirect & "</AutoRedirect>"
86:        Print #intFF, vbTab & vbTab & "<AutoRedirectFull>" & g_objSettings.AutoRedirectFull & "</AutoRedirectFull>"
87:        Print #intFF, vbTab & vbTab & "<AutoRedirectNonReg>" & g_objSettings.AutoRedirectNonReg & "</AutoRedirectNonReg>"
88:        Print #intFF, vbTab & vbTab & "<AutoRedirectFullNonReg>" & g_objSettings.AutoRedirectFullNonReg & "</AutoRedirectFullNonReg>"
89:        Print #intFF, vbTab & vbTab & "<AutoRedirectFullNonOps>" & g_objSettings.AutoRedirectFullNonOps & "</AutoRedirectFullNonOps>"
90:        Print #intFF, vbTab & vbTab & "<AutoStart>" & g_objSettings.AutoStart & "</AutoStart>"
91:        Print #intFF, vbTab & vbTab & "<CompactDBOnExit>" & g_objSettings.CompactDBOnExit & "</CompactDBOnExit>"
92:        Print #intFF, vbTab & vbTab & "<ConfirmExit>" & g_objSettings.ConfirmExit & "</ConfirmExit>"
93:        Print #intFF, vbTab & vbTab & "<DCValidateTags>" & g_objSettings.DCValidateTags & "</DCValidateTags>"
94:        Print #intFF, vbTab & vbTab & "<DCIncludeOPed>" & g_objSettings.DCIncludeOPed & "</DCIncludeOPed>"
95:        Print #intFF, vbTab & vbTab & "<OPBypass>" & g_objSettings.OPBypass & "</OPBypass>"
96:        Print #intFF, vbTab & vbTab & "<PreloadWinsocks>" & g_objSettings.PreloadWinsocks & "</PreloadWinsocks>"
97:        Print #intFF, vbTab & vbTab & "<SendMessageAFK>" & g_objSettings.SendMessageAFK & "</SendMessageAFK>"
98:        Print #intFF, vbTab & vbTab & "<RegOnly>" & g_objSettings.RegOnly & "</RegOnly>"
99:        Print #intFF, vbTab & vbTab & "<MentoringSystem>" & g_objSettings.MentoringSystem & "</MentoringSystem>"
100:        Print #intFF, vbTab & vbTab & "<PreventSearchBots>" & g_objSettings.PreventSearchBots & "</PreventSearchBots>"
101:        Print #intFF, vbTab & vbTab & "<DescriptiveBanMsg>" & g_objSettings.DescriptiveBanMsg & "</DescriptiveBanMsg>"
102:        Print #intFF, vbTab & vbTab & "<UseBotName>" & g_objSettings.UseBotName & "</UseBotName>"
103:        Print #intFF, vbTab & vbTab & "<Passive>" & g_objSettings.Passive & "</Passive>"
104:        Print #intFF, vbTab & vbTab & "<DynUpdate>" & g_objSettings.DynUpdate & "</DynUpdate>"
105:        Print #intFF, vbTab & vbTab & "<DynDNSUpdateEna>" & g_objSettings.DynDNSUpdateEna & "</DynDNSUpdateEna>"
106:        Print #intFF, vbTab & vbTab & "<NoIPUpdateEna>" & g_objSettings.NoIPUpdateEna & "</NoIPUpdateEna>"
107:        Print #intFF, vbTab & vbTab & "<NoIPUpdateStartUp>" & g_objSettings.NoIPUpdateStartUp & "</NoIPUpdateStartUp>"
               
        '-------------NEW REDIRECT CHECK BOXES-----------------------------------------------------------------------------------------------
110:        Print #intFF, vbTab & vbTab & "<RedirectFMinShare>" & g_objSettings.RedirectFMinShare & "</RedirectFMinShare>"
111:        Print #intFF, vbTab & vbTab & "<RedirectFMaxShare>" & g_objSettings.RedirectFMaxShare & "</RedirectFMaxShare>"
112:        Print #intFF, vbTab & vbTab & "<RedirectFMinSlots>" & g_objSettings.RedirectFMinSlots & "</RedirectFMinSlots>"
113:        Print #intFF, vbTab & vbTab & "<RedirectFMaxSlots>" & g_objSettings.RedirectFMaxSlots & "</RedirectFMaxSlots>"
114:        Print #intFF, vbTab & vbTab & "<RedirectFMaxHubs>" & g_objSettings.RedirectFMaxHubs & "</RedirectFMaxHubs>"
115:        Print #intFF, vbTab & vbTab & "<RedirectFSlotPerHub>" & g_objSettings.RedirectFSlotPerHub & "</RedirectFSlotPerHub>"
116:        Print #intFF, vbTab & vbTab & "<RedirectFNoTag>" & g_objSettings.RedirectFNoTag & "</RedirectFNoTag>"
117:        Print #intFF, vbTab & vbTab & "<RedirectFTooOldDCpp>" & g_objSettings.RedirectFTooOldDCpp & "</RedirectFTooOldDCpp>"
118:        Print #intFF, vbTab & vbTab & "<RedirectFTooOldNMDC>" & g_objSettings.RedirectFTooOldNMDC & "</RedirectFTooOldNMDC>"
119:        Print #intFF, vbTab & vbTab & "<RedirectFBWPerSlot>" & g_objSettings.RedirectFBWPerSlot & "</RedirectFBWPerSlot>"
120:        Print #intFF, vbTab & vbTab & "<RedirectFFakeShare>" & g_objSettings.RedirectFFakeShare & "</RedirectFFakeShare>"
121:        Print #intFF, vbTab & vbTab & "<RedirectFFakeTag>" & g_objSettings.RedirectFFakeTag & "</RedirectFFakeTag>"
122:        Print #intFF, vbTab & vbTab & "<RedirectFPasMode>" & g_objSettings.RedirectFPasMode & "</RedirectFPasMode>"
        '----------------STOP HERE----------------------------------------------------------------------------------------------------------
124:        Print #intFF, vbTab & vbTab & "<FilterCPrefix>" & g_objSettings.FilterCPrefix & "</FilterCPrefix>"
125:        Print #intFF, vbTab & vbTab & "<EnabledCommands>" & g_objSettings.EnabledCommands & "</EnabledCommands>"
126:        Print #intFF, vbTab & vbTab & "<ScriptSafeMode>" & g_objSettings.ScriptSafeMode & "</ScriptSafeMode>"
127:        Print #intFF, vbTab & vbTab & "<StartMinimized>" & g_objSettings.StartMinimized & "</StartMinimized>"
128:        Print #intFF, vbTab & vbTab & "<SendMsgAsPrivate>" & g_objSettings.SendMsgAsPrivate & "</SendMsgAsPrivate>"
129:        Print #intFF, vbTab & vbTab & "<PasswordMode>" & g_objSettings.PasswordMode & "</PasswordMode>"
130:        Print #intFF, vbTab & vbTab & "<WordWrap>" & g_objSettings.WordWrap & "</WordWrap>"
131:        Print #intFF, vbTab & vbTab & "<DenyNoTag>" & g_objSettings.DenyNoTag & "</DenyNoTag>"
132:        Print #intFF, vbTab & vbTab & "<HideFadeImg>" & g_objSettings.HideFadeImg & "</HideFadeImg>"
133:        Print #intFF, vbTab & vbTab & "<CheckFakeShare>" & g_objSettings.CheckFakeShare & "</CheckFakeShare>"
134:        Print #intFF, vbTab & vbTab & "<PreventGuessPass>" & g_objSettings.PreventGuessPass & "</PreventGuessPass>"
135:        Print #intFF, vbTab & vbTab & "<EnableFloodWall>" & g_objSettings.EnableFloodWall & "</EnableFloodWall>"
136:        Print #intFF, vbTab & vbTab & "<RedirectFGP>" & g_objSettings.RedirectFGP & "</RedirectFGP>"
137:        Print #intFF, vbTab & vbTab & "<OpsCanRedirect>" & g_objSettings.OpsCanRedirect & "</OpsCanRedirect>"
138:        Print #intFF, vbTab & vbTab & "<ChatOnly>" & g_objSettings.ChatOnly & "</ChatOnly>"
139:        Print #intFF, vbTab & vbTab & "<MinClsSearchSend>" & g_objSettings.MinClsSearchSend & "</MinClsSearchSend>"
140:        Print #intFF, vbTab & vbTab & "<MinClsConnectSend>" & g_objSettings.MinClsConnectSend & "</MinClsConnectSend>"
141:        Print #intFF, vbTab & vbTab & "<MinimizeTray>" & g_objSettings.MinimizeTray & "</MinimizeTray>"
142:        Print #intFF, vbTab & vbTab & "<HideMyinfos>" & g_objSettings.HideMyinfos & "</HideMyinfos>"
143:        Print #intFF, vbTab & vbTab & "<ACOClients>" & g_objSettings.ACOClients & "</ACOClients>"
144:        Print #intFF, vbTab & vbTab & "<EnabledScheduler>" & g_objSettings.EnabledScheduler & "</EnabledScheduler>"
145:        Print #intFF, vbTab & vbTab & "<PriorityBl>" & g_objSettings.PriorityBl & "</PriorityBl>"
146:        Print #intFF, vbTab & vbTab & "<PopUpNewReg>" & g_objSettings.PopUpNewReg & "</PopUpNewReg>"
147:        Print #intFF, vbTab & vbTab & "<PopUpOpConected>" & g_objSettings.PopUpOpConected & "</PopUpOpConected>"
148:        Print #intFF, vbTab & vbTab & "<PopUpOpDisconected>" & g_objSettings.PopUpOpDisconected & "</PopUpOpDisconected>"
149:        Print #intFF, vbTab & vbTab & "<PopUpUserKick>" & g_objSettings.PopUpUserKick & "</PopUpUserKick>"
150:        Print #intFF, vbTab & vbTab & "<PopUpUserBaned>" & g_objSettings.PopUpUserBaned & "</PopUpUserBaned>"
151:        Print #intFF, vbTab & vbTab & "<PopUpUserRedirected>" & g_objSettings.PopUpUserRedirected & "</PopUpUserRedirected>"
152:        Print #intFF, vbTab & vbTab & "<PopUpStartedServing>" & g_objSettings.PopUpStartedServing & "</PopUpStartedServing>"
153:        Print #intFF, vbTab & vbTab & "<PopUpUserBaned>" & g_objSettings.PopUpUserBaned & "</PopUpUserBaned>"
154:        Print #intFF, vbTab & vbTab & "<PopUpUserRedirected>" & g_objSettings.PopUpUserRedirected & "</PopUpUserRedirected>"
155:        Print #intFF, vbTab & vbTab & "<PopUpStopedServing>" & g_objSettings.PopUpStopedServing & "</PopUpStopedServing>"
156:        Print #intFF, vbTab & vbTab & "<PopUpCoreError>" & g_objSettings.PopUpCoreError & "</PopUpCoreError>"
157:        Print #intFF, vbTab & vbTab & "<MoveForm>" & g_objSettings.MoveForm & "</MoveForm>"
158:        Print #intFF, vbTab & vbTab & "<MagneticWin>" & g_objSettings.MagneticWin & "</MagneticWin>"
159:        Print #intFF, vbTab & vbTab & "<StartWin>" & g_objSettings.StartWin & "</StartWin>"
160:        Print #intFF, vbTab & vbTab & "<blSkin>" & g_objSettings.blSkin & "</blSkin>"
161:        Print #intFF, vbTab & vbTab & "<RndSkin>" & g_objSettings.RndSkin & "</RndSkin>"
162:        Print #intFF, vbTab & vbTab & "<Plugins>" & g_objSettings.Plugins & "</Plugins>"
163:        Print #intFF, vbTab & "</Boolean>"

165:        Print #intFF, vbTab & "<Byte>"
166:        Print #intFF, vbTab & vbTab & "<DCMaxHubs>" & g_objSettings.DCMaxHubs & "</DCMaxHubs>"
167:        Print #intFF, vbTab & vbTab & "<DCOSlots>" & g_objSettings.DCOSlots & "</DCOSlots>"
168:        Print #intFF, vbTab & vbTab & "<MinSlots>" & g_objSettings.MinSlots & "</MinSlots>"
169:        Print #intFF, vbTab & vbTab & "<MaxSlots>" & g_objSettings.MaxSlots & "</MaxSlots>"
170:        Print #intFF, vbTab & vbTab & "<MinShareSize>" & g_objSettings.MinShareSize & "</MinShareSize>"
171:        Print #intFF, vbTab & vbTab & "<MaxShareSize>" & g_objSettings.MaxShareSize & "</MaxShareSize>"
172:        Print #intFF, vbTab & vbTab & "<CPrefix>" & g_objSettings.CPrefix & "</CPrefix>"
173:        Print #intFF, vbTab & vbTab & "<DCOSpeed>" & g_objSettings.DCOSpeed & "</DCOSpeed>"
174:        Print #intFF, vbTab & vbTab & "<SendJoinMsg>" & g_objSettings.SendJoinMsg & "</SendJoinMsg>"
175:        Print #intFF, vbTab & vbTab & "<MaxPassAttempts>" & g_objSettings.MaxPassAttempts & "</MaxPassAttempts>"
176:        Print #intFF, vbTab & vbTab & "<FWMyINFO>" & g_objSettings.FWMyINFO & "</FWMyINFO>"
177:        Print #intFF, vbTab & vbTab & "<FWGetNickList>" & g_objSettings.FWGetNickList & "</FWGetNickList>"
178:        Print #intFF, vbTab & vbTab & "<FWActiveSearch>" & g_objSettings.FWActiveSearch & "</FWActiveSearch>"
179:        Print #intFF, vbTab & vbTab & "<FWPassiveSearch>" & g_objSettings.FWPassiveSearch & "</FWPassiveSearch>"
'182:        Print #intFF, vbTab & vbTab & "<MinMyinfoFakeCls>" & g_objSettings.MinMyinfoFakeCls & "</MinMyinfoFakeCls>"
        'TheNOP svn 159 , hidden, no interface setting.
182:        Print #intFF, vbTab & vbTab & "<FWMainchat>" & g_objSettings.FWMainChat & "</FWMainchat>"
        'Print #intFF, vbTab & vbTab & "<FWGlobal>" & g_objSettings.FWGlobal & "</FWGlobal>"
184:    Print #intFF, vbTab & "</Byte>"

186:    Print #intFF, vbTab & "<Integer>"
187:        Print #intFF, vbTab & vbTab & "<MinPassiveSearchLen>" & g_objSettings.MinPassiveSearchLen & "</MinPassiveSearchLen>"
188:        Print #intFF, vbTab & vbTab & "<FWInterval>" & g_objSettings.FWInterval & "</FWInterval>"
189:        Print #intFF, vbTab & vbTab & "<FWBanLength>" & g_objSettings.FWBanLength & "</FWBanLength>"
190:        Print #intFF, vbTab & vbTab & "<MinConnectCls>" & g_objSettings.MinConnectCls & "</MinConnectCls>"
191:        Print #intFF, vbTab & vbTab & "<MinSearchCls>" & g_objSettings.MinSearchCls & "</MinSearchCls>"
192:        Print #intFF, vbTab & vbTab & "<ZLINELENGHT>" & g_objSettings.ZLINELENGHT & "</ZLINELENGHT>"
193:        Print #intFF, vbTab & vbTab & "<PriorityVal>" & g_objSettings.PriorityVal & "</PriorityVal>"
            'MySQL interface
195:        Print #intFF, vbTab & vbTab & "<DBServerPort>" & g_objSettings.DBServerPort & "</DBServerPort>"
196:        Print #intFF, vbTab & vbTab & "<DBType>" & g_objSettings.DBType & "</DBType>"
197:    Print #intFF, vbTab & "</Integer>"

199:    Print #intFF, vbTab & "<Long>"
200:        Print #intFF, vbTab & vbTab & "<MaxUsers>" & g_objSettings.MaxUsers & "</MaxUsers>"
201:        Print #intFF, vbTab & vbTab & "<DefaultBanTime>" & g_objSettings.DefaultBanTime & "</DefaultBanTime>"
202:        Print #intFF, vbTab & vbTab & "<ScriptTimeout>" & g_objSettings.ScriptTimeout & "</ScriptTimeout>"
203:        Print #intFF, vbTab & vbTab & "<MaxMessageLen>" & g_objSettings.MaxMessageLen & "</MaxMessageLen>"
204:        Print #intFF, vbTab & vbTab & "<DataFragmentLen>" & g_objSettings.DataFragmentLen & "</DataFragmentLen>"
205:        Print #intFF, vbTab & vbTab & "<ConDropInterval>" & g_objSettings.ConDropInterval & "</ConDropInterval>"
206:        Print #intFF, vbTab & vbTab & "<FWDropMsgInterval>" & g_objSettings.FWDropMsgInterval & "</FWDropMsgInterval>"
207:        Print #intFF, vbTab & vbTab & "<lngSkin>" & g_objSettings.lngSkin & "</lngSkin>"

209:    Print #intFF, vbTab & "</Long>"

211:    Print #intFF, vbTab & "<Double>"
212:        Print #intFF, vbTab & vbTab & "<IMinShare>" & g_objSettings.IMinShare & "</IMinShare>"
213:        Print #intFF, vbTab & vbTab & "<IMaxShare>" & g_objSettings.IMaxShare & "</IMaxShare>"
214:        Print #intFF, vbTab & vbTab & "<DCSlotsPerHub>" & g_objSettings.DCSlotsPerHub & "</DCSlotsPerHub>"
215:        Print #intFF, vbTab & vbTab & "<DCBandPerSlot>" & g_objSettings.DCBandPerSlot & "</DCBandPerSlot>"
216:        Print #intFF, vbTab & vbTab & "<DCMinVersion>" & g_objSettings.DCMinVersion & "</DCMinVersion>"
217:        Print #intFF, vbTab & vbTab & "<NMDCMinVersion>" & g_objSettings.NMDCMinVersion & "</NMDCMinVersion>"
218:    Print #intFF, vbTab & "</Double>"

220:    Print #intFF, vbTab & "<Tags>"
221:        For Each objTag In m_colTags
222:            Print #intFF, vbTab & vbTab & "<Tag ";
223:             Print #intFF, "Name=""" & objTag.Name & """ ";
224:             Print #intFF, "/>"
225:        Next
226:    Print #intFF, vbTab & "</Tags>"

228:    Print #intFF, "</Settings>";

230:    Close intFF

'---------------------------------------------------------------------------------
    'Perm IP bans
234:    strTemp = G_APPPATH & "\Settings\PermIPBans.xml"

236:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp

238:    intFF = FreeFile

    'Append to file
241:    Open strTemp For Append As intFF

243:    Print #intFF, "<PTDCH>"

245:    For Each objTB In g_colIPBans.PermItems
246:            Print #intFF, vbTab & "<PermIPBans>"
247:            Print #intFF, vbTab & vbTab & "<IP>" & objTB.IP & "</IP>"
248:            Print #intFF, vbTab & vbTab & "<Nick>" & objTB.Nick & "</Nick>"
249:            Print #intFF, vbTab & vbTab & "<BannedBy>" & objTB.BannedBy & "</BannedBy>"
250:            Print #intFF, vbTab & vbTab & "<Reason>" & objTB.Reason & "</Reason>"
251:            Print #intFF, vbTab & "</PermIPBans>"
252:    Next

254:    Print #intFF, "</PTDCH>";

256:    Close intFF

'---------------------------------------------------------------------------------
    'Temp IP Bans
260:    strTemp = G_APPPATH & "\Settings\TempIPBans.xml"

262:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp

264:    intFF = FreeFile

    'Append to file
267:    Open strTemp For Append As intFF

269:    Print #intFF, "<PTDCH>"

271:    For Each objTB In g_colIPBans.TempItems
272:            If DateDiff("n", Now, objTB.ExpDate) > 0 Then
273:                Print #intFF, vbTab & "<TempIPBans>"
274:                Print #intFF, vbTab & vbTab & "<IP>" & objTB.IP & "</IP>"
275:                Print #intFF, vbTab & vbTab & "<ExpDate>" & objTB.ExpDate & "</ExpDate>"
276:                Print #intFF, vbTab & vbTab & "<Nick>" & objTB.Nick & "</Nick>"
277:                Print #intFF, vbTab & vbTab & "<BannedBy>" & objTB.BannedBy & "</BannedBy>"
278:                Print #intFF, vbTab & vbTab & "<Reason>" & objTB.Reason & "</Reason>"
279:                Print #intFF, vbTab & "</TempIPBans>"
280:            End If
281:    Next

283:    Print #intFF, "</PTDCH>";

285:    Close intFF

'---------------------------------------------------------------------------------
    'Commands
289:    strTemp = G_APPPATH & "\Settings\Commands.xml"

291:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp

293:    intFF = FreeFile

295:    Open strTemp For Append As intFF

297:    Print #intFF, "<Commands>"

    'Loop through command collection
300:    For Each objCommand In g_colCommands
301:        Print #intFF, vbTab & "<Command ID=""" & objCommand.ID & """ Trigger=""" & objCommand.Name & """ Class=""" & objCommand.Class & """ Enabled=""" & objCommand.Enabled & """ Description=""" & XMLEscape(objCommand.Description) & """ />"
302:    Next

304:    Print #intFF, "</Commands>";

306:    Close intFF

'---------------------------------------------------------------------------------
309:    strTemp = G_APPPATH & "\Settings\UsersMessages.xml"
    
    'Must exist but it must not be replace
312:    If Not g_objFileAccess.FileExists(strTemp) Then
313:            Call LoadAndSaveXML(enuXML.EGUsersMessages)
314:    End If

'---------------------------------------------------------------------------------
        'Add\Rem auto start up at windows
       
319:     If g_objSettings.StartWin Then _
            AddRegRun _
         Else RemRegRun
'---------------------------------------------------------------------------------
        'Save txtNotepad to text file
324:     g_objFileAccess.WriteFile G_APPPATH & "\Settings\notepad.txt", txtNotePad.Text

        'Save motd to text file
327:     g_objFileAccess.WriteFile G_APPPATH & "\Settings\motd.txt", g_objSettings.JoinMsg

        'Save sql commands to text file
330:     g_objFileAccess.WriteFile G_APPPATH & "\Settings\bdManager.sql", m_objSciExplorerSQL.Text

332:     AddLog "Save settings.."
333:  Exit Sub
    
335:
Err:
336:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SaveSettings()"
End Sub
'------------------------------------------------------------------------------
' End Setting related methods
'------------------------------------------------------------------------------

Public Sub RefreshGUI(Optional ByRef bRefreshRegBanIPs = False)

2:    Dim lng         As Long
3:    Dim objTag      As clsTag
4:    Dim lvwItems    As ListItems

6:    On Error Resume Next
    
      'Set text boxes
9:     For lng = 0 To txtData.UBound
10:        If lng = 18 Then _
              txtData(18).Text = ChrW$(CallByName(g_objSettings, txtData(lng).Tag, VbGet)) _
           Else _
              txtData(lng).Text = CallByName(g_objSettings, txtData(lng).Tag, VbGet)
14:    Next
    
       'Set check boxes
17:    For lng = 0 To chkData.UBound
18:        chkData(lng).Value = Abs(CallByName(g_objSettings, chkData(lng).Tag, VbGet))
19:    Next
    
       'Set scroll bars
22:    For lng = 0 To vslData.UBound
23:        vslData(lng).Value = CallByName(g_objSettings, vslData(lng).Tag, VbGet)
24:    Next
     
       'Set combo boxes
27:    cmbData(0).Text = CallByName(g_objSettings, cmbData(0).Tag, VbGet)
28:    cmbData(1).ListIndex = CallByName(g_objSettings, cmbData(1).Tag, VbGet)
29:    cmbData(2).ListIndex = CallByName(g_objSettings, cmbData(2).Tag, VbGet)

31:    lvwItems("minshare").SubItems(1) = g_objSettings.MinShareMsg
32:    lvwItems("dcppminversion").SubItems(1) = g_objSettings.DCppMinVersionMsg
33:    lvwItems("minslots").SubItems(1) = g_objSettings.MinSlotsMsg
34:    lvwItems("maxslots").SubItems(1) = g_objSettings.MaxSlotsMsg
35:    lvwItems("hsratio").SubItems(1) = g_objSettings.HSRatioMsg
36:    lvwItems("bsratio").SubItems(1) = g_objSettings.BSRatioMsg
37:    lvwItems("maxhubs").SubItems(1) = g_objSettings.MaxHubsMsg
38:    lvwItems("nmdcminversion").SubItems(1) = g_objSettings.NMDCMinVersionMsg
39:    lvwItems("denynotag").SubItems(1) = g_objSettings.DenyNoTagMsg
40:    lvwItems("maxshare").SubItems(1) = g_objSettings.MaxShareMsg
41:    lvwItems("fakeshare").SubItems(1) = g_objSettings.FakeShareMsg
42:    lvwItems("faketag").SubItems(1) = g_objSettings.FakeTagMsg
43:    lvwItems("socks5").SubItems(1) = g_objSettings.Socks5Msg
44:    lvwItems("passivemode").SubItems(1) = g_objSettings.PassiveModeMsg
45:    lvwItems("NoCOClients").SubItems(1) = g_objSettings.NoCOClientsMsg

47:    Set lvwItems = Nothing
        
       'Add tags to ListBox
50:    lstTagsEx.Clear
    
52:    For Each objTag In m_colTags
53:        lstTagsEx.AddItem objTag.Name
54:    Next
    
        'Set redirect option
        Select Case True
            Case g_objSettings.AutoRedirect: optRedirect(0).Value = True
            Case g_objSettings.AutoRedirectNonReg: optRedirect(1).Value = True
            Case g_objSettings.AutoRedirectFull: optRedirect(2).Value = True
            Case g_objSettings.AutoRedirectFullNonReg: optRedirect(3).Value = True
            Case g_objSettings.AutoRedirectFullNonOps: optRedirect(4).Value = True
            Case Else: optRedirect(5).Value = True
57:     End Select
    
        'Set join message option
60:     optJM(g_objSettings.SendJoinMsg).Value = True

62:     If bRefreshRegBanIPs Then
63:         Call DBGetRegRecord 'Refresh registered user list
64:         Call DBGetBanRecord 'Refresh banned nicknames
            'Refresh IP Ban Perm/Temp
66:         Call mnuTempIPBan_Click(4)
67:         Call mnuPermIPBan_Click(4)
68:         Me.Refresh
69:     End If
        
        'Refresh all controls..this process refresh the ram memory used, minimize ptdch and see task manager ;-)
72:     If Me.WindowState = vbMinimized Then
73:        LockWindowUpdate frmHub.hwnd
74:        Dim CTL As Control
75:        For Each CTL In frmHub.Controls
76:            CTL.Refresh
77:            DoEvents
78:        Next
79:        frmHub.Visible = True
80:        DoEvents
81:        frmHub.Refresh
82:        frmHub.Visible = False
83:        LockWindowUpdate 0
84:    End If
End Sub
'------------------------------------------------------------------------------
'Assorted support methods
'------------------------------------------------------------------------------
Private Sub UpdateFailedReg(ByRef curUser As clsUser, ByRef blnLoggedIn As Boolean)
1:    Dim objFR   As clsFailedReg
2:    Dim strKey  As String

4:    On Error GoTo Err
    'Create key
6:    strKey = LCase$(curUser.sName) & "|" & curUser.IP
    
8:    On Error Resume Next
    
    'If they logged in sucessfully, then remove failed attempts from collection
11:    If blnLoggedIn Then
12:        m_colFailedReg.Remove strKey
13:    Else
           'Get object
15:        Set objFR = m_colFailedReg(strKey)
        
17:        On Error GoTo Err
        
           'If they have never tried to log in, create object
20:        If ObjPtr(objFR) = 0 Then
21:            Set objFR = New clsFailedReg
22:            m_colFailedReg.Add objFR, strKey
23:        End If
        
          'Update check
26:        If objFR.Check(curUser) Then _
                m_colFailedReg.Remove strKey
28:    End If
    
30:    Exit Sub
    
32:
Err:
33:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateFailedReg(, " & blnLoggedIn & ")"
End Sub
Friend Function UpdateConnectAttempt(ByRef wskUser As Winsock, ByRef blnLoggedIn As Boolean) As Boolean
1:    Dim objCA   As clsConnectAttempt
2:    Dim strIP   As String
    
4:    On Error GoTo Err
    
6:    strIP = wskUser.RemoteHostIP
    
    'If logged in, then we can just
    'svn 223
10:    If Not blnLoggedIn Then
    'If blnLoggedIn Then
        'If removed we won't know if they are hammering...
        'm_colConnectAttempts.Remove strIP
    'Else
        'Attempt to retrieve object
16:        On Error Resume Next
17:        Set objCA = m_colConnectAttempts(strIP)
18:        On Error GoTo Err
        
        'If it doesn't exist, create a new one
21:        If ObjPtr(objCA) = 0 Then
22:            Set objCA = New clsConnectAttempt
23:            objCA.IP = strIP
            
25:            m_colConnectAttempts.Add objCA, strIP
26:        End If
        
        'Check if they are hammering; if so, delete the record (as they are banned
        'for 30 minutes and won't be needed)
30:        If objCA.Check(wskUser) Then
31:            UpdateConnectAttempt = True
32:            m_colConnectAttempts.Remove strIP
33:        End If
34:    End If
    
36:    Exit Function
    
38:
Err:
39:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateConnectAttempt(, " & blnLoggedIn & ")"
End Function
Private Sub CheckOutdatedRecords()
1:    On Error GoTo Err
    
3:    Dim objCA   As clsConnectAttempt
    
    'Loop through collection
6:    For Each objCA In m_colConnectAttempts
7:        If DateDiff("n", objCA.LastAttempt, Now) > 10 Then _
         m_colConnectAttempts.Remove objCA.IP
9:    Next
        
11:    Exit Sub
    
13:
Err:
14:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.CheckOutdatedRecords()"
End Sub
Public Function LockToKey(ByRef strLock As String, Optional ByVal n As Long = 5) As String
1:    Dim arrChar() As Integer
2:    Dim arrRet() As Integer
3:    Dim i As Long
4:    Dim ub As Long
  
6:    On Error GoTo Err

    'n = 5 for hub and client locks
  
    'The lock only continues to the first space (Pk= comes after)
11:    i = InStrB(1, strLock, " Pk=")
12:    If i Then strLock = LeftB$(strLock, i - 1)

    'Make sure the lock is longer than 2 characters
15:    ub = Len(strLock)
16:    If ub < 3 Then LockToKey = "Invalid lock": Exit Function
    
    'Create buffers to hold vars
19:    ReDim arrChar(1 To ub) As Integer
20:    ReDim arrRet(1 To ub) As Integer
21:    LockToKey = String$(ub * 10, vbNullChar)
    
    'Set first character of string
24:    arrChar(1) = AscW(strLock)
    
    'Set all others and Xor the current and the previous together
27:    For i = 2 To ub
28:        arrChar(i) = AscW(Mid$(strLock, i))
29:        arrRet(i) = arrChar(i) Xor arrChar(i - 1)
30:    Next
    
    'Create first character based on first, last, second last and n from lock
33:    arrRet(1) = arrChar(1) Xor arrChar(ub) Xor arrChar(ub - 1) Xor n
    
    'Delete lock array since it is no longer needed
36:    Erase arrChar
    
    'Set i to 1 so that it starts on the first character
39:    i = 1
    
    'Now loop through and fix all the characters
42:    For n = 1 To ub
43:        arrRet(n) = ((CLng(arrRet(n)) * 16) And 240) Or ((arrRet(n) \ 16) And 15)
        
        'Escape if needed (increment position by 10 if escape is used)
        Select Case arrRet(n)
            Case 0: Mid$(LockToKey, i, 10) = "/%DCN000%/": i = i + 10
            Case 5: Mid$(LockToKey, i, 10) = "/%DCN005%/": i = i + 10
            Case 36: Mid$(LockToKey, i, 10) = "/%DCN036%/": i = i + 10
            Case 96: Mid$(LockToKey, i, 10) = "/%DCN096%/": i = i + 10
            Case 124: Mid$(LockToKey, i, 10) = "/%DCN124%/": i = i + 10
            Case 126: Mid$(LockToKey, i, 10) = "/%DCN126%/": i = i + 10
            Case Else: Mid$(LockToKey, i, 1) = Chr$(arrRet(n)): i = i + 1
46:        End Select
47:    Next
    
    'Erase array containing Xor-ed values
50:    Erase arrRet
    
    'Trim off extra space in the buffer
53:    LockToKey = Left$(LockToKey, i - 1)

55:    Exit Function

57:
Err:
58:    HandleError Err.Number, Err.Description, Erl & "|" & "frmMain.LockToKey(" & strLock & ", " & n & ")"
End Function

'Closes the user's winsock
Public Sub CloseSocket(ByRef intIndex As Integer)
1:    wskLoop_Close intIndex
End Sub

'Calls the VB DoEvents function from VBS
Public Function DoEventsForMe() As Long
1:    DoEventsForMe = DoEvents
End Function

'Registers a bot name in the lists
Public Sub RegisterBotName(ByRef strName As String, Optional ByRef blnOperator As Boolean = True, Optional ByVal dblShare As Double, Optional ByRef strDescription As String, Optional ByRef strConnection As String, Optional ByRef strEmail As String, Optional ByVal lngIcon As Long = 1, Optional ByRef blnOverwrite As Boolean = True)
1:    Dim lngLoop     As Long
2:    Dim blnHold     As Boolean
3:    Dim strOne      As String
4:    Dim strTwo      As String
      
6:    On Error GoTo Err
    
    'Check if it has already been registered
9:    lngLoop = IsRegisteredBotName(strName)
    
    'Make sure we lock the bot name so nobody can log in with it
12:    g_objRegistered.Add strName, "Auto bot name locking system", Locked, "PTDCH / Core", , True
    
    'If -1, then it isn't, else it is
15:    If lngLoop = -1 Then
        'Fix description if needed
17:        If LenB(strDescription) Then
18:            If InStrB(1, strDescription, "$") Then strDescription = Replace(strDescription, "$", "_")
19:            If InStrB(1, strDescription, "|") Then strDescription = Replace(strDescription, "|", "_")
20:        End If
    
        'Resize array
23:        m_lngBotsUB = m_lngBotsUB + 1
24:        ReDim Preserve m_arrBots(0 To m_lngBotsUB) As typBot
        
        'Update new array element
27:        m_arrBots(m_lngBotsUB).Name = strName
28:        m_arrBots(m_lngBotsUB).MyINFO = "$MyINFO $ALL " & strName & " " & strDescription & "$ $" & strConnection & ChrW$(lngIcon) & "$" & strEmail & "$" & dblShare & "$|"
29:        m_arrBots(m_lngBotsUB).Operator = blnOperator
        
31:        If G_SERVING Then
            'Add to nicklist
33:            g_colUsers.AppendNL strName, blnOperator
        
            'Prepare buffers
36:            If blnOperator Then _
                strOne = m_arrBots(m_lngBotsUB).MyINFO & "$OpList " & strName & "$$|" _
            Else _
                strOne = m_arrBots(m_lngBotsUB).MyINFO
                
41:            strTwo = "$Hello " & strName & "|" & strOne
            
            'Send to all users
44:            For Each m_objLoopUser In g_colUsers
45:                If m_objLoopUser.NoHello Then _
                    m_objLoopUser.SendData strOne _
                Else _
                    m_objLoopUser.SendData strTwo
49:            Next
            
51:            Set m_objLoopUser = Nothing
52:        End If
53:    Else
        'Check if we should overwrite
55:        If blnOverwrite Then
            'Fix description if needed
57:            If LenB(strDescription) Then
58:                If InStrB(1, strDescription, "$") Then strDescription = Replace(strDescription, "$", "_")
59:                If InStrB(1, strDescription, "|") Then strDescription = Replace(strDescription, "|", "_")
60:            End If
        
            'Update array
63:            m_arrBots(lngLoop).MyINFO = "$MyINFO $ALL " & strName & " " & strDescription & "$ $" & strConnection & ChrW$(lngIcon) & "$" & strEmail & "$" & dblShare & "$|"
            
65:            blnHold = m_arrBots(lngLoop).Operator
            
            'Update oplist if necessary or else just myinfo string
68:            If blnHold = blnOperator Then
69:                If G_SERVING Then _
                        g_colUsers.SendToAll m_arrBots(lngLoop).MyINFO
71:            Else
72:                m_arrBots(lngLoop).Operator = blnOperator
                
                'Update only if serving
75:                If G_SERVING Then
76:                    g_colUsers.RemoveNL strName, blnHold
77:                    g_colUsers.AppendNL strName, blnOperator
78:                    g_colUsers.SendToAll "$OpList " & g_colUsers.OpList & "|" & m_arrBots(lngLoop).MyINFO
79:                End If
80:            End If
81:        End If
82:    End If
    
84:    Exit Sub
    
86:
Err:
87:    If ObjPtr(m_objLoopUser) Then Set m_objLoopUser = Nothing
88:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.RegisterBotName()"
End Sub

'Unregisters a bot name in the lists
Public Sub UnregisterBotName(ByRef strName As String)
1:    Dim lngLoop         As Long
    
3:    On Error GoTo Err

    'See if it is registered
6:    lngLoop = IsRegisteredBotName(strName)
    
    'Make sure we unregister the bot name
9:    g_objRegistered.Remove strName
    
    'If it is registered, then get to work!
12:    If Not lngLoop = -1 Then
13:        m_lngBotsUB = m_lngBotsUB - 1
        
        'Update nicklist (local and remote) if serving
16:        If G_SERVING Then
17:            g_colUsers.RemoveNL strName, m_arrBots(lngLoop).Operator
18:            g_colUsers.SendToAll "$Quit " & strName & "|"
19:        End If
        
        'If there are no bots left in the array, destroy it, otherwise resize it
22:        If m_lngBotsUB = -1 Then
23:            Erase m_arrBots
24:        Else
25:            m_arrBots(lngLoop) = m_arrBots(m_lngBotsUB + 1)
26:            ReDim Preserve m_arrBots(0 To m_lngBotsUB) As typBot
27:        End If
28:    End If

30:    Exit Sub

32:
Err:
33:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UnregisterBotName(""" & strName & """)"
End Sub

'Returns -1 is not registered, index in array otherwise
Public Function IsRegisteredBotName(ByRef strName As String) As Long
1:    Dim lngLoop As Long
    
3:    On Error GoTo Err
    
    'Make any needed character replacements in the nickname
6:    If InStrB(1, strName, " ") Then strName = Replace(strName, " ", "_")
7:    If InStrB(1, strName, "$") Then strName = Replace(strName, "$", "_")
8:    If InStrB(1, strName, "|") Then strName = Replace(strName, "|", "_")
    
    'Set to -1, meaning it hasn't found the bot name
11:    IsRegisteredBotName = -1
    
    'Make sure they are bots in the array first
14:    If Not m_lngBotsUB = -1 Then
        'Loop through and see if the name matches any; if it does, return array index
16:        For lngLoop = 0 To m_lngBotsUB
17:            If m_arrBots(lngLoop).Name = strName Then IsRegisteredBotName = lngLoop: Exit For
18:        Next
19:    End If
    
21:    Exit Function
    
23:
Err:
24:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.IsRegisteredBotName(""" & strName & """)"
End Function
'Switches to the next redirect address
Public Sub NextRedirect()
1:    Static lngIndex As Long
    
3:    On Error GoTo Err
    
    'If the UBound is zero, then there is only one IP
6:    If m_lngRedirectUB Then
        'Increment index
8:        lngIndex = lngIndex + 1
    
        'If index has surpassed the max index, then set it back to the beginning
11:        If lngIndex > m_lngRedirectUB Then lngIndex = 0
        
        'Set redirect IP in settings to the new IP
14:        g_objSettings.RedirectIP = m_arrRedirectIPs(lngIndex)
15:    End If
    
17:    Exit Sub
    
19:
Err:
20:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.NextRedirect()"
End Sub

'Sends bot MyINFOs to a user
Friend Sub UpdateBots(ByRef curUser As clsUser)
1:    Dim lngLoop     As Long

3:    On Error GoTo Err

      'Make sure there are bot names to send
6:    If Not m_lngBotsUB = -1 Then
         'Loop through and send MyINFO strings
8:       For lngLoop = 0 To m_lngBotsUB
9:            curUser.SendData m_arrBots(lngLoop).MyINFO
10:      Next
11:   End If

      'Update Chats Run
14:   g_objChatRoom.UpDate curUser
        
16:   Exit Sub

18:
Err:
19:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateBots()"
End Sub
'Converts minutes to it's equivalent length in years, months, weeks, days, hours and minutes
Public Function MinToDate(ByVal lngMinutes As Long) As String
1:    Dim lngYears As Long
2:    Dim lngMonths As Long
3:    Dim lngWeeks As Long
4:    Dim lngDays As Long
5:    Dim lngHours As Long
    
7:    On Error GoTo Err
    
    'If there are more than 59 minutes, there is at least 1 hour
10:    If lngMinutes > 59 Then
11:        lngHours = lngMinutes \ 60
12:        lngMinutes = lngMinutes Mod 60
        'If there are more than 23 hours, there is at least 1 day
14:        If lngHours > 23 Then
15:            lngDays = lngHours \ 24
16:            lngHours = lngHours Mod 24
            
            'If there are more than 29 days, there is at least 1 month
19:            If lngDays > 29 Then
20:                lngMonths = lngDays \ 30
21:                lngDays = lngDays Mod 30
            'If there are more than 7 days, there is at least 1 week
23:            Else
24:                If lngDays > 6 Then
25:                    lngWeeks = lngDays \ 7
26:                    lngDays = lngDays Mod 7
27:                End If
28:            End If
            'If there are more than 11 months, there is at least 1 year
30:            If lngMonths > 11 Then
31:                lngYears = lngMonths \ 12
32:                lngMonths = lngMonths Mod 12
33:            End If
34:        End If
35:    End If

    'Construct length in words
38:    If lngYears > 1 Then MinToDate = lngYears & " years, " _
    Else: If lngYears Then MinToDate = lngYears & " year, "
40:    If lngMonths > 1 Then MinToDate = MinToDate & lngMonths & " months, " _
    Else: If lngMonths Then MinToDate = MinToDate & lngMonths & " month, "
42:    If lngWeeks > 1 Then MinToDate = MinToDate & lngWeeks & " weeks, " _
    Else: If lngWeeks Then MinToDate = MinToDate & lngWeeks & " week, "
44:    If lngDays > 1 Then MinToDate = MinToDate & lngDays & " days, " _
    Else: If lngDays Then MinToDate = MinToDate & lngDays & " day, "
46:    If lngHours > 1 Then MinToDate = MinToDate & lngHours & " hours, " _
    Else: If lngHours Then MinToDate = MinToDate & lngHours & " hour, "
    
    'If there are no minutes to add, then we need to remove the extra space/comma
50:    If lngMinutes > 1 Then
51:        MinToDate = MinToDate & lngMinutes & " minutes"
52:        Exit Function
53:    Else
54:        If lngMinutes Then
55:            MinToDate = MinToDate & lngMinutes & " minute"
56:            Exit Function
57:        End If
58:    End If
59:    If LenB(MinToDate) Then
60:            MinToDate = LeftB$(MinToDate, LenB(MinToDate) - 4)
61:        Else
62:            MinToDate = "0 minutes"
63:    End If
    
65:    Exit Function
    
67:
Err:
68:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.MinToDate()"
End Function

Public Property Get oPermaCon() As Connection
1:    Set oPermaCon = m_objPermaCon
End Property

Public Property Get ServingDate() As Date
1:    ServingDate = m_datServingDate
End Property

Public Property Get IsServing() As Boolean
1:    IsServing = G_SERVING
End Property

'------------------------------------------------------------------------------
'Script related functions / events
'------------------------------------------------------------------------------
Private Sub tmrScriptTimer_Timer(Index As Integer)
1:    On Error GoTo Err
    
    'This runs the tmrScriptTimer_Timer event
    '
    '  -- Parameters : None
    '  -- Format     : Sub tmrScriptTimer_Timer()
    '
    '  -- Called whenever the alloted interval for the timer has gone
    '     off
    
11:    If m_arrScriptEvents(Index, vbStmrScriptTimer_Timer) Then _
          ScriptControl(Index).Run "tmrScriptTimer_Timer" _
       Else _
          tmrScriptTimer(Index).Enabled = False
16:
Err:
End Sub
Public Property Get oTimersAPI() As clsTimersCol
1:    Set oTimersAPI = m_objTimers
End Property
Public Sub NewTimersAPI()
1:    On Error GoTo Err

3:    If Not m_objTimers Is Nothing Then
4:        Set m_objTimers = Nothing
5:    End If
    
7:    Set m_objTimers = New clsTimersCol
    
9:    Exit Sub
Err:
12:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.NewTimersAPI()"
End Sub
Private Sub m_objTimers_Timer(ByVal Index As Integer)
1:    On Error GoTo Err
    'This runs the TimerAPI_Timer event
    '
    '  -- Parameters : Index = Timer ID basead in Index
    '  -- Format     : Sub TimerAPI_Timer()
    '
    '  -- Called whenever the alloted interval for the timer has gone
    '     off
    
15:    SEvent_API_Timer Index
16:
Err:
End Sub
Private Sub SEvent_API_Timer(ByRef intIndex As Integer)
1:    Dim lng As Long
2:    Dim i As Integer

    '  -- Parameters : Index = Timer ID basead in Index
    '  -- Format     : Sub tmrAPI_Timer(Index)
    '
    '  -- Called whenever the alloted interval for the timer has gone
    '     off
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "tmrAPI_Timer", intIndex
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbStmrAPI_Timer) Then _
          ScriptControl(lng).Run "tmrAPI_Timer", intIndex
28:    Next
End Sub
Private Sub ScriptControl_Error(Index As Integer)
1:    On Error GoTo Err
    
    'This runs the Error event
    '
    '  -- Parameters : lngLine (line the error occured on)
    '  -- Format     : Sub Error(lngLine)
    '
    '  -- Called when an error occurs and On Error Resume Next is not used

12:    If m_arrScriptEvents(Index, vbSError) Then _
         ScriptControl(Index).Run "Error", ScriptControl(Index).Error.Line
14:
Err:
End Sub

'Private Sub ScriptControl_Timeout(Index As Integer)
'    On Error Resume Next
'
'    'This runs the Timeout event
'    '
'    '  -- Parameters : None
'    '  -- Format     : Sub Timeout()
'    '
'    '  -- Run when the script code timeouts
'
'    If m_arrScriptEvents(Index, vbSTimeout) Then _
'        ScriptControl(Index).Run "Timeout"
'End Sub

Private Function SEvent_PreConnectionRequest(ByRef wskShock As Variant, ByRef requestID As Long) As Boolean
1:    Dim i As Integer
    
    'This runs the ConnectionRequest event
    '
    '  -- Parameters : wskShock (wskShock of user who is trying to connect)
    '  -- Format     : False = OK
    '                : True = Cancel requestID
    '
    '  -- Called when a new user pre connect shcoks
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      SEvent_PreConnectionRequest = CBool(g_objPlugin(i).Object.RunEvent("PreConnectionRequest", wskShock, requestID))
19:                      If SEvent_PreConnectionRequest = True Then Exit For
20:                 End If
21:            End If
22:        Next
23:    End If
End Function
Private Sub SEvent_AttemptedConnection(ByRef strIP As String)
1:    Dim lng As Long
2:    Dim i As Integer
    
    'This runs the AttemptedConnection event
    '
    '  -- Parameters : strIP (IP of user who is trying to connect)
    '  -- Format     : Sub AttemptedConnection(strIP)
    '
    '  -- Called when a new user tries to connect to the hub (before any
    '     messages are exchanged)
    
12:    On Error Resume Next
    
       'Run plugin events
15:    If g_PluginsFound And g_objSettings.Plugins Then
16:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
17:            If g_objPlugin(i).UseEvents Then
18:                 If g_objPlugin(i).Object.Enabled Then
19:                      g_objPlugin(i).Object.RunEvent "AttemptedConnection", strIP
20:                 End If
21:            End If
22:        Next
23:    End If
    
       'Run script envents
26:    For lng = 1 To m_lngScriptEventsUB
27:        If m_arrScriptEvents(lng, vbSAttemptedConnection) Then _
         ScriptControl(lng).Run "AttemptedConnection", strIP
29:    Next
End Sub
Private Sub SEvent_UserConnected(ByRef curUser As clsUser)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the UserConnected event
    '
    '  -- Parameters : curUser (user's clsUser object)
    '  -- Format     : Sub UserConnected(curUser)
    '
    '  -- Called when a new user logs in (sends MyINFO basically)
    '  -- There is a slight difference with this and NMDCH; DDCH calls this
    '     event after it recieves MyINFO and the like, while NMDCH calls this
    '     right after the ValidateNick message
    
14:    On Error Resume Next
    
       'Run plugin events
17:    If g_PluginsFound And g_objSettings.Plugins Then
18:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
19:            If g_objPlugin(i).UseEvents Then
20:                 If g_objPlugin(i).Object.Enabled Then
21:                      g_objPlugin(i).Object.RunEvent "UserConnected", curUser
22:                 End If
23:            End If
24:        Next
25:    End If
    
       'Run script envents
28:    For lng = 1 To m_lngScriptEventsUB
29:        If m_arrScriptEvents(lng, vbSUserConnected) Then _
                ScriptControl(lng).Run "UserConnected", curUser
31:    Next
End Sub
Private Sub SEvent_RegConnected(ByRef curUser As clsUser)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the RegConnected event
    '
    '  -- Parameters : curUser (user's clsUser object)
    '  -- Format     : Sub RegConnected(curUser)
    '
    '  -- Called when a new user logs in (sends MyINFO basically)
    '  -- There is a slight difference with this and NMDCH; NMDCH combines this
    '     event with OpConnected, which is wrong...they are not an op so people
    '     can get confused
    
14:    On Error Resume Next
    
       'Run plugin events
17:    If g_PluginsFound And g_objSettings.Plugins Then
18:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
19:            If g_objPlugin(i).UseEvents Then
20:                 If g_objPlugin(i).Object.Enabled Then
21:                      g_objPlugin(i).Object.RunEvent "RegConnected", curUser
22:                 End If
23:            End If
24:        Next
25:    End If
    
       'Run script envents
28:    For lng = 1 To m_lngScriptEventsUB
29:        If m_arrScriptEvents(lng, vbSRegConnected) Then _
            ScriptControl(lng).Run "RegConnected", curUser
31:    Next
End Sub
Private Sub SEvent_OpConnected(ByRef curUser As clsUser)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the OpConnected event
    '
    '  -- Parameters : curUser (user's clsUser object)
    '  -- Format     : Sub OpConnected(curUser)
    '
    '  -- Called when a new user logs in (sends MyINFO basically)
    '  -- There are two slight differences with this and NMDCH; DDCH calls this
    '     event after it recieves MyINFO and the like, while NMDCH calls this
    '     right after the ValidateNick message. Second DDCH only calls this event
    '     with OPERATORS and calls non-ops-but-registered users with RegConnected
    
15:    On Error Resume Next
    
       'Run plugin events
18:    If g_PluginsFound And g_objSettings.Plugins Then
19:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
20:            If g_objPlugin(i).UseEvents Then
21:                 If g_objPlugin(i).Object.Enabled Then
22:                      g_objPlugin(i).Object.RunEvent "OpConnected", curUser
23:                 End If
24:            End If
25:        Next
26:    End If
    
       'Run script envents
29:    For lng = 1 To m_lngScriptEventsUB
30:        If m_arrScriptEvents(lng, vbSOpConnected) Then _
             ScriptControl(lng).Run "OpConnected", curUser
32:    Next
End Sub
Private Sub SEvent_UserQuit(ByRef curUser As clsUser)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the UserQuit event
    '
    '  -- Parameters : curUser (user's clsUser object)
    '  -- Format     : Sub UserQuit(curUser)
    '
    '  -- Called when a user leaves the hub. After this sub is called
    '     the user's clsHub object is destroyed (however you must remove it
    '     from any collections that might contain it in the hub)
    
13:    On Error Resume Next
    
       'Run plugin events
16:    If g_PluginsFound And g_objSettings.Plugins Then
17:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
18:            If g_objPlugin(i).UseEvents Then
19:                 If g_objPlugin(i).Object.Enabled Then
20:                      g_objPlugin(i).Object.RunEvent "UserQuit", curUser
21:                 End If
22:            End If
23:        Next
24:    End If
    
       'Run script envents
27:    For lng = 1 To m_lngScriptEventsUB
28:        If m_arrScriptEvents(lng, vbSUserQuit) Then _
                ScriptControl(lng).Run "UserQuit", curUser
30:    Next
End Sub
Private Sub SEvent_StartedServing()
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the StartedServing event
    '
    '  -- Parameters : None
    '  -- Format     : Sub StartServing()
    '
    '  -- Run when the hub starts serving
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "StartedServing"
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbSStartedServing) Then _
                ScriptControl(lng).Run "StartedServing"
28:    Next
End Sub
Friend Sub SEvent_AddedRegisteredUser(ByRef strName As String, ByRef strPassword As String, ByRef intClass As Integer, ByRef strAdminName As String, ByRef lngMin As Long)
1:    Dim lng As Long
2:    Dim i As Integer

    '
    '  -- Parameters : strName (name of the user who was registered)
    '  -- Format     : Sub AddedRegisteredUser(strName)
    '
    '  -- Run when a new user is registered (via the clsRegistered.Add function)
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "AddedRegisteredUser", strName, strPassword, intClass, strAdminName, lngMin
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbSAddedRegisteredUser) Then _
         ScriptControl(lng).Run "AddedRegisteredUser", strName, strPassword, intClass, strAdminName, lngMin
28:    Next
End Sub
Friend Sub SEvent_AddedPermBan(ByRef strIP As String, ByRef strNick As String, ByRef strBannedBy As String, ByRef strReason As String)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the AddedPermBan event
    '
    '  -- Parameters : strIP (IP that was banned)
    '  -- Format     : Sub AddedPermBan(strIP, strNick, strBannedBy, strReason)
    '
    '  -- Called when a user perm bans an IP via the clsIPBans.Add method
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "AddedPermBan", strIP, strNick, strBannedBy, strReason
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbSAddedPermBan) Then _
         ScriptControl(lng).Run "AddedPermBan", strIP, strNick, strBannedBy, strReason
28:    Next
End Sub
Friend Sub SEvent_AddedTempBan(ByRef strIP As String, ByRef lngMinutes As Long, ByRef strNick As String, ByRef strBannedBy As String, ByRef strReason As String)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the AddedPermBan event
    '
    '  -- Parameters : strIP (IP that was banned)
    '  -- Format     : Sub AddedTempBan(strIP, lngMinutes, strNick, strBannedBy, strReason)
    '
    '  -- Called when a user perm bans an IP via the clsIPBans.Add method
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "AddedTempBan", strIP, lngMinutes, strNick, strBannedBy, strReason
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbSAddedTempBan) Then _
         ScriptControl(lng).Run "AddedTempBan", strIP, lngMinutes, strNick, strBannedBy, strReason
28:    Next
End Sub
Private Sub SEvent_StartedRedirecting()
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the StartedRedirecting event
    '
    '  -- Parameters : None
    '  -- Format     : Sub StartedRedirecting()
    '
    '  -- Called when the hub owner presses the Redirect All button (just before
    '     redirects are done)
    
12:    On Error Resume Next
    
       'Run plugin events
15:    If g_PluginsFound And g_objSettings.Plugins Then
16:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
17:            If g_objPlugin(i).UseEvents Then
18:                 If g_objPlugin(i).Object.Enabled Then
19:                      g_objPlugin(i).Object.RunEvent "StartedRedirecting"
20:                 End If
21:            End If
22:        Next
23:    End If
    
       'Run script envents
26:    For lng = 1 To m_lngScriptEventsUB
27:        If m_arrScriptEvents(lng, vbSStartedRedirecting) Then _
         ScriptControl(lng).Run "StartedRedirecting"
29:    Next
End Sub
Private Sub SEvent_StoppedServing()
1:    Dim lng As Long
2:    Dim i As Integer
    
    'This runs the StoppedServing event
    '
    '  -- Parameters : None
    '  -- Format     : Sub StoppedServing
    '
    '  -- Called when the hub owner stops serving
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "StoppedServing"
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbSStoppedServing) Then _
         ScriptControl(lng).Run "StoppedServing"
28:    Next
End Sub
Private Sub SEvent_MassMessage(ByRef strMessage As String)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the MassMessage event
    '
    '  -- Parameters : strMessage (message that was sent to all users)
    '  -- Format     : Sub MassMessage(strMessage)
    '
    '  -- Run when the hub owner presses the mass message button
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "MassMessage", strMessage
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbSMassMessage) Then _
         ScriptControl(lng).Run "MassMessage", strMessage
28:    Next
End Sub
Private Sub SEvent_UnloadMain()
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the UnloadMain event
    '
    '  -- Parameters : None
    '  -- Format     : Sub UnloadMain()
    '
    '  -- Called when the hub is closing up
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "UnloadMain"
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbSUnloadMain) Then _
         ScriptControl(lng).Run "UnloadMain"
28:    Next
End Sub
Friend Sub SEvent_RemovedRegisteredUser(ByRef strName As String)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the RemovedRegisteredUser event
    '
    '  -- Parameters : strName (name of the user who was unregistered)
    '  -- Format     : Sub RemovedRegisteredUser(strName)
    '
    '  -- Run when a user is unregistered (via the clsRegistered.Remove function)
    
11:    On Error Resume Next
    
       'Run plugin events
14:    If g_PluginsFound And g_objSettings.Plugins Then
15:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
16:            If g_objPlugin(i).UseEvents Then
17:                 If g_objPlugin(i).Object.Enabled Then
18:                      g_objPlugin(i).Object.RunEvent "RemovedRegisteredUser", strName
19:                 End If
20:            End If
21:        Next
22:    End If
    
       'Run script envents
25:    For lng = 1 To m_lngScriptEventsUB
26:        If m_arrScriptEvents(lng, vbSRemovedRegisteredUser) Then _
         ScriptControl(lng).Run "RemovedRegisteredUser", strName
28:    Next
End Sub
Private Sub SEvent_CustComArrival(ByRef curUser As clsUser, ByRef objCommand As clsCommand, ByRef strMessage As String, ByRef blnMainChat As Boolean)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the CustComArrival event
    '
    '   -- Parameters : curUser (current user's object)
    '                 : objCommand (current command object from collection)
    '                 : strMessage (command text sent by user)
    '   -- Format     : Sub CustComArrival(curUser, objCommand, strMessage)
    '
    '   -- Fired when a user sends a command which is in the command collection
    '      but not supported by the hub
    
14:    On Error Resume Next
    
       'Run plugin events
17:    If g_PluginsFound And g_objSettings.Plugins Then
18:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
19:            If g_objPlugin(i).UseEvents Then
20:                 If g_objPlugin(i).Object.Enabled Then
21:                      g_objPlugin(i).Object.RunEvent "CustComArrival", curUser, objCommand, strMessage, blnMainChat
22:                 End If
23:            End If
24:        Next
25:    End If
    
       'Run script envents
28:    For lng = 1 To m_lngScriptEventsUB
29:        If m_arrScriptEvents(lng, vbSCustComArrival) Then _
                ScriptControl(lng).Run "CustComArrival", curUser, objCommand, strMessage, blnMainChat
31:    Next
End Sub
Private Function SEvent_PreDataArrival(ByRef curUser As clsUser, ByRef strCommand As String) As String
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the PreDataArrival event
    '
    '  -- Parameters : curUser (the current user's clsUser object)
    '                : strData (data that was sent)
    '  -- Format     : Function PreDataArrival(curUser, strData)
    '
    '  -- Called when a user sends data to the hub, but before the hub parses
    '     it
    '  -- It should return the string it should parse
    '
    
15:    On Error Resume Next
    
       'Run plugin events
18:    If g_PluginsFound And g_objSettings.Plugins Then
19:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
20:            If g_objPlugin(i).UseEvents Then
21:                 If g_objPlugin(i).Object.Enabled Then
22:                      SEvent_PreDataArrival = CStr(g_objPlugin(i).Object.RunEvent("PreDataArrival", curUser, strCommand))
23:                      If LenB(SEvent_PreDataArrival) = 0 Then Exit For
24:                 End If
25:            End If
26:        Next
27:    End If
    
       'Run script envents
30:    For lng = 1 To m_lngScriptEventsUB
31:        If m_arrScriptEvents(lng, vbSPreDataArrival) Then
32:                SEvent_PreDataArrival = ScriptControl(lng).Run("PreDataArrival", curUser, strCommand)
33:                If LenB(SEvent_PreDataArrival) = 0 Then Exit For
34:        End If
35:    Next
End Function
Private Sub SEvent_DataArrival(ByRef curUser As clsUser, ByRef strCommand As String)
1:    Dim lng As Long
2:    Dim i As Integer

    'SEvent_DataArrival curUser, strCommand
    '
    'This runs the DataArrival event
    '
    '  -- Parameters : curUser (the current user's clsUser object)
    '                : strData (data that was sent)
    '  -- Format     : Sub DataArrival(curUser, strData)
    '
    '  -- Called when a user sends data to the hub
    '  -- Difference with NMDCH is DDCH sends ALL data to the script, while
    '     NMDCH does it selectively (nothing before and including ValidateNick)
    
16:    On Error Resume Next
    
       'Run plugin events
19:    If g_PluginsFound And g_objSettings.Plugins Then
20:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
21:            If g_objPlugin(i).UseEvents Then
22:                 If g_objPlugin(i).Object.Enabled Then
23:                      g_objPlugin(i).Object.RunEvent "DataArrival", curUser, strCommand
24:                 End If
25:            End If
26:        Next
27:    End If
    
       'Run script envents
30:    For lng = 1 To m_lngScriptEventsUB
31:                If m_arrScriptEvents(lng, vbSDataArrival) Then _
                        ScriptControl(lng).Run "DataArrival", curUser, strCommand
33:    Next
End Sub
Private Function SEvent_FailedConf(ByRef curUser As clsUser, ByRef intType As enuAlert) As Boolean
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the FailedConf event
    '
    '   -- Parameters : curUser (current user's object)
    '                 : intType As enuAlert enumerator
    '
    '   -- Return     : boolean
    '
    '   -- Fired when a user get rejected by the hub.(user fail hub's rules.)
    
13:    On Error Resume Next
    
       'Run plugin events
16:    If g_PluginsFound And g_objSettings.Plugins Then
17:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
18:            If g_objPlugin(i).UseEvents Then
19:                 If g_objPlugin(i).Object.Enabled Then
20:                      SEvent_FailedConf = CBool(g_objPlugin(i).Object.RunEvent("FailedConf", curUser, intType))
21:                      If SEvent_FailedConf Then Exit For
22:                 End If
23:            End If
24:        Next
25:    End If
    
       'Run script envents
28:    For lng = 1 To m_lngScriptEventsUB
29:        If m_arrScriptEvents(lng, vbSFailedConf) Then
30:            SEvent_FailedConf = ScriptControl(lng).Run("FailedConf", curUser, intType)
31:            If SEvent_FailedConf Then Exit For
32:        End If
33:    Next
End Function
Public Sub SEvent_ReloadPlugin(Optional ByRef lngIndex As Long = -1)
1:    Dim lng As Long
2:    Dim i As Integer
                
    'This runs the ReoladPlugin event
    '
    '  -- Parameters : lngIndex = -1 (Reolad All plugins)
    '                : lngIndex <> -1 (Reolad plugin by index)
    '                :
    '  -- Return     : none
    
13:    On Error Resume Next

15:    If lngIndex = -1 Then
           'Run plugin events
17:        If g_PluginsFound And g_objSettings.Plugins Then
18:            For i = LBound(g_objPlugin) To UBound(g_objPlugin)
19:                If g_objPlugin(i).UseEvents Then
20:                     If g_objPlugin(i).Object.Enabled Then
21:                          g_objPlugin(i).Object.RunEvent "Reload"
22:                     End If
23:                End If
24:            Next
25:        End If
26:    Else
27:        If g_objPlugin(lngIndex).UseEvents And g_objPlugin(i).Object.Enabled Then
28:             g_objPlugin(lngIndex).Object.RunEvent "Reload"
29:        End If
30:    End If
End Sub
Public Sub SEvent_SwitchPlugin(ByRef blnSate As Boolean, Optional ByRef lngIndex As Long = -1)
1:    Dim lng As Long
2:    Dim i As Integer
                
    'This runs the Switch event
    '
    '  -- Parameters : blnSate (True/False)
    '                : lngIndex = -1 (Switch All plugins)
    '                : lngIndex <> -1 (Switch plugin by index)
    '  -- Return     : none
    
13:    On Error Resume Next

15:    If lngIndex = -1 Then
           'Run plugin events
17:        If g_PluginsFound And g_objSettings.Plugins Then
18:            For i = LBound(g_objPlugin) To UBound(g_objPlugin)
19:                If g_objPlugin(i).UseEvents Then
20:                    g_objPlugin(i).Object.RunEvent "Switch", blnSate
21:                End If
22:            Next
23:        End If
24:    Else
25:        If g_objPlugin(lngIndex).UseEvents Then
26:             g_objPlugin(lngIndex).Object.RunEvent "ReoladPlugin", blnSate
27:        End If
28:    End If
End Sub
Public Sub SEvent_CoreError(ByRef strErrDesc As String)
1:    Dim lng As Long
2:    Dim i As Integer

    'This runs the CoreError event
    '
    '  -- Parameters :strErrDesc:
    '                 Error log format : 'Date-Time|Method|Number|DLLError|Description|Version|Beta|
    '
    '  -- Return     : none
    '  -- Called when a core error
    
16:    On Error Resume Next
    
       'Run plugin events
19:    If g_PluginsFound And g_objSettings.Plugins Then
20:        For i = LBound(g_objPlugin) To UBound(g_objPlugin)
21:            If g_objPlugin(i).UseEvents Then
22:                 If g_objPlugin(i).Object.Enabled Then
23:                      g_objPlugin(i).Object.RunEvent "CoreError", strErrDesc
24:                 End If
25:            End If
26:        Next
27:    End If
    
       'Run script envents
30:    For lng = 1 To m_lngScriptEventsUB
31:                If m_arrScriptEvents(lng, vbSCoreError) Then _
                        ScriptControl(lng).Run "CoreError", strErrDesc
33:    Next
End Sub

Friend Sub SFindEvents(ByRef intIndex As Integer)
1:    Dim lngProc     As Long
2:    Dim lngProcUB   As Long
3:    Dim objProc     As Procedure
4:    Dim objProcs    As Procedures
    
6:    On Error GoTo Err
    
    'Prepare vars
9:    Set objProcs = ScriptControl(intIndex).Procedures
10:   lngProcUB = objProcs.Count
    
    'Clear out array
13:    For lngProc = 0 To vbSFC
14:        m_arrScriptEvents(intIndex, lngProc) = False
15:    Next
    
    '#If PREDATAARRIVAL Then
    '    'Make sure that PreDataArrival is disabled if it was this script that was using it
    '    If m_intPDIndex = intIndex Then m_intPDIndex = 0
    '#End If

    'Loop through procedures
23:    For Each objProc In objProcs
        'Find out which procedure it is, and set it to True in the boolean array
        Select Case LCase$(objProc.Name)
            Case "main"
25:                If objProc.NumArgs = 0 Then _
                        m_arrScriptEvents(intIndex, vbSMain) = True
         
         #If DataArrival Then
            Case "dataarrival"
29:                If objProc.NumArgs = 2 Then _
                        m_arrScriptEvents(intIndex, vbSDataArrival) = True
         #End If
            
            Case "attemptedconnection"
33:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSAttemptedConnection) = True
            Case "userconnected"
35:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSUserConnected) = True
            Case "regconnected"
37:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSRegConnected) = True
            Case "opconnected"
39:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSOpConnected) = True
            Case "userquit"
41:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSUserQuit) = True
            Case "startedserving"
43:                If objProc.NumArgs = 0 Then _
                        m_arrScriptEvents(intIndex, vbSStartedServing) = True
            Case "addedregistereduser"
45:                If objProc.NumArgs = 5 Then _
                        m_arrScriptEvents(intIndex, vbSAddedRegisteredUser) = True
            Case "wskscript_close"
47:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSwskScript_Close) = True
            Case "wskscript_connect"
49:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSwskScript_Connect) = True
            Case "wskscript_connectionrequest"
51:                If objProc.NumArgs = 2 Then _
                        m_arrScriptEvents(intIndex, vbSwskScript_ConnectionRequest) = True
            Case "wskscript_dataarrival"
53:                If objProc.NumArgs = 2 Then _
                        m_arrScriptEvents(intIndex, vbSwskScript_DataArrival) = True
            Case "wskscript_error"
55:                If objProc.NumArgs = 3 Then _
                        m_arrScriptEvents(intIndex, vbSwskScript_Error) = True
            Case "tmrscripttimer_timer"
57:                If objProc.NumArgs = 0 Then _
                        m_arrScriptEvents(intIndex, vbStmrScriptTimer_Timer) = True
            Case "addedpermban"
59:                If objProc.NumArgs = 4 Then _
                        m_arrScriptEvents(intIndex, vbSAddedPermBan) = True
            Case "addedtempban"
61:                If objProc.NumArgs = 5 Then _
                        m_arrScriptEvents(intIndex, vbSAddedTempBan) = True
            Case "startedredirecting"
63:                If objProc.NumArgs = 0 Then _
                        m_arrScriptEvents(intIndex, vbSStartedServing) = True
            Case "stoppedserving"
65:                If objProc.NumArgs = 0 Then _
                        m_arrScriptEvents(intIndex, vbSStoppedServing) = True
            Case "massmessage"
67:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSMassMessage) = True
            Case "unloadmain"
69:                If objProc.NumArgs = 0 Then _
                        m_arrScriptEvents(intIndex, vbSUnloadMain) = True
            Case "error"
71:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSError) = True
            Case "timeout"
73:                If objProc.NumArgs = 0 Then _
                        m_arrScriptEvents(intIndex, vbSTimeout) = True
            Case "removedregistereduser"
75:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSRemovedRegisteredUser) = True
            Case "custcomarrival"
77:                If objProc.NumArgs = 4 Then _
                        m_arrScriptEvents(intIndex, vbSCustComArrival) = True
            Case "failedconf"
79:                If objProc.NumArgs = 2 Then _
                        If objProc.HasReturnValue Then _
                        m_arrScriptEvents(intIndex, vbSFailedConf) = True
            
            'Evaluate PreDataArrival if needed
            #If PreDataArrival Then
                Case "predataarrival"
85:                    If objProc.NumArgs = 2 Then _
                        If objProc.HasReturnValue Then _
                            m_arrScriptEvents(intIndex, vbSPreDataArrival) = True
            #End If

            Case "coreerror"
90:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbSCoreError) = True
                        
            Case "tmrapi_timer"
91:                If objProc.NumArgs = 1 Then _
                        m_arrScriptEvents(intIndex, vbStmrAPI_Timer) = True
                        
93:        End Select
94:    Next

       'Set winsock collection booleans
97:    g_colSWinsocks(CStr(intIndex)).SetBools m_arrScriptEvents(intIndex, vbSwskScript_Connect), _
                                               m_arrScriptEvents(intIndex, vbSwskScript_Close), _
                                               m_arrScriptEvents(intIndex, vbSwskScript_ConnectionRequest), _
                                               m_arrScriptEvents(intIndex, vbSwskScript_DataArrival), _
                                               m_arrScriptEvents(intIndex, vbSwskScript_Error)
        
103:    Exit Sub
    
105:
Err:
106:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SFindEvents(" & intIndex & ")"
End Sub

Friend Sub SClearEvents(ByRef intIndex As Integer)
1:    Dim lng As Long
    
3:    On Error GoTo Err
    
    'Loop through array and set all procedure enabled settings to false
6:    For lng = 0 To vbSFC
7:        m_arrScriptEvents(intIndex, lng) = False
8:    Next
    
    'Set all winsock vars to false
11:    g_colSWinsocks(CStr(intIndex)).SetBools False, False, False, False, False
    
    'Set PreDataArrival index to zero if it in use by this script
    '#If PREDATAARRIVAL Then
    '    If m_intPDIndex = intIndex Then m_intPDIndex = 0
    '#End If
    
18:    Exit Sub
    
20:
Err:
21:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SClearEvents(" & intIndex & ")"
End Sub
Friend Sub SResizeArrEvent(ByRef intSize As Integer, ByRef blnPreserve As Boolean)
1:    On Error GoTo Err
    
    'Preserve elements if needed
4:    If blnPreserve Then
5:        ReDim Preserve m_arrScriptEvents(1 To intSize, 0 To vbSFC) As Boolean
6:    Else
7:        ReDim m_arrScriptEvents(1 To intSize, 0 To vbSFC) As Boolean
8:    End If
        
       'Set new UBound
11:    m_lngScriptEventsUB = intSize
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SResizeArrEvent(" & intSize & ", " & blnPreserve & ")"
End Sub
'------------------------------------------------------------------------------
'End Script related functions / events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'End Update DSN/IP Subs
'------------------------------------------------------------------------------
Private Sub UpdateIPs_Timer()
1:    On Error GoTo Err
    
    'do check every IPCheckInterval
4:    If lIntervalMin < IPCheckInterval Then
5:        lIntervalMin = lIntervalMin + 1
6:        Exit Sub
7:    End If
8:    lIntervalMin = 0
    
10:    Dim successor
11:    Dim HubIP
12:    Dim HostIP
13:    Dim X

15:    HubIP = frmHub.DetectHubIP
16:    If Not HubIP <> "" Then Exit Sub
17:    DoEvents
    
       'check against local IPs
20:    If IPinRange("10.0.0.0", "10.255.255.255", HubIP) Then Exit Sub
21:    If IPinRange("127.0.0.0", "127.255.255.255", HubIP) Then Exit Sub
22:    If IPinRange("172.16.0.0", "172.31.255.255", HubIP) Then Exit Sub
23:    If IPinRange("192.168.0.0", "192.168.255.255", HubIP) Then Exit Sub

25:    For X = 0 To 9 'a maximum of 10 services should be enough
26:        If Service(X) = "" Then Exit Sub
        
28:        DoEvents
29:        HostIP = frmHub.ResolveHostName(CStr(Host(X)))
  
31:        If HubIP <> HostIP Then
32:            DoEvents
33:            successor = frmHub.UpdateIP(CStr(Service(X)), CStr(User(X)), CStr(Pass(X)), CStr(Host(X)))
34:            g_colUsers.SendPrivateToOps g_objSettings.BotName, "IP UPDATE " & CStr(Host(X)) & ": " & CStr(successor)
35:        End If
36:        AddLog "IP UpDate: " & CStr(Host(X)) & ": " & CStr(successor)
37:    Next

39:    Exit Sub
40:
Err:
41:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateIPs_Timer"
End Sub
Private Function QueuedConnect(ByRef strName As String) As Boolean
1:    Dim datTemp     As Date
    
3:    On Error GoTo DNE
    

6:    datTemp = m_colRevConnects(strName)
7:    m_colRevConnects.Remove strName
8:    QueuedConnect = (DateDiff("s", datTemp, Now) < 60)
    
10:    Exit Function
    
12:
DNE:
13:    QueuedConnect = False
End Function
Public Function UpdateIP(Service, UserName, Password, HostName)
1:    On Error Resume Next
2:    If LCase(Service) = "dyndns" Then _
         UpdateIP = UpdateDynDNS(UserName, Password, HostName, True, "", False)
4:    If LCase(Service) = "noip" Then _
         UpdateIP = UpdateNoIP(UserName, Password, HostName, "")
End Function
Private Sub UpdateDNSs()
1:    Dim strTemp As String
2:    Dim successor
3:    Dim HubIP
4:    Dim HostIP
    
6:    On Error GoTo Err
    
8:    HubIP = DetectHubIP
9:    DoEvents
    
11:    If HubIP = vbNullString Then Exit Sub
    
13:    If g_objSettings.NoIPUpdateEna Then
14:        If g_objSettings.NoIPDNS1 <> vbNullString Then
15:            HostIP = ResolveHostName(g_objSettings.NoIPDNS1)
16:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(7).Text
18:            If Not HubIP = HostIP Then
19:                successor = UpdNOIP(g_objSettings.NoIPUser, g_objSettings.NoIPPass, g_objSettings.NoIPDNS1)
20:                DoEvents
21:                strTemp = "IP UpDate: " & g_objSettings.NoIPDNS1 & " - " & successor: AddLog strTemp: stbMain.Panels(8).Text = strTemp
22:            End If
23:        End If
        
25:        If g_objSettings.NoIPDNS2 <> vbNullString Then
26:            HostIP = ResolveHostName(g_objSettings.NoIPDNS2)
27:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(7).Text
29:            If Not HubIP = HostIP Then
30:                successor = UpdNOIP(g_objSettings.NoIPUser, g_objSettings.NoIPPass, g_objSettings.NoIPDNS2)
31:                DoEvents
32:                strTemp = "IP UpDate : " & g_objSettings.NoIPDNS2 & " - " & successor: AddLog strTemp: stbMain.Panels(8).Text = strTemp
33:            End If
34:        End If
        
36:        If g_objSettings.NoIPDNS3 <> vbNullString Then
37:            HostIP = ResolveHostName(g_objSettings.NoIPDNS3)
38:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(7).Text
40:            If Not HubIP = HostIP Then
41:                successor = UpdNOIP(g_objSettings.NoIPUser, g_objSettings.NoIPPass, g_objSettings.NoIPDNS3)
42:                DoEvents
43:                strTemp = "IP UpDate: " & g_objSettings.NoIPDNS3 & " - " & successor: AddLog strTemp: stbMain.Panels(8).Text = strTemp
44:            End If
45:        End If
        
47:        If g_objSettings.NoIPDNS4 <> vbNullString Then
48:            HostIP = ResolveHostName(g_objSettings.NoIPDNS4)
49:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(7).Text
51:            If Not HubIP = HostIP Then
52:                successor = UpdNOIP(g_objSettings.NoIPUser, g_objSettings.NoIPPass, g_objSettings.NoIPDNS4)
53:                DoEvents
54:                strTemp = "IP UpDate : " & g_objSettings.NoIPDNS4 & " - " & successor: AddLog strTemp: stbMain.Panels(8).Text = strTemp
55:            End If
56:        End If
57:    End If

59:    If g_objSettings.DynDNSUpdateEna Then
60:        If g_objSettings.DynDNS1 <> vbNullString Then
61:            HostIP = ResolveHostName(g_objSettings.DynDNS1)
62:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(7).Text
64:            If Not HubIP = HostIP Then
65:                successor = UpdDynDNS(g_objSettings.DynDNSUser, g_objSettings.DynDNSPass, g_objSettings.DynDNS1)
66:                DoEvents
67:                strTemp = "IP UpDate: " & g_objSettings.DynDNS1 & " - " & successor: AddLog strTemp: stbMain.Panels(8).Text = strTemp
68:            End If
69:        End If
        
71:        If g_objSettings.DynDNS2 <> vbNullString Then
72:            HostIP = ResolveHostName(g_objSettings.DynDNS2)
73:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(7).Text
75:            If Not HubIP = HostIP <> vbNullString Then
76:                successor = UpdDynDNS(g_objSettings.DynDNSUser, g_objSettings.DynDNSPass, g_objSettings.DynDNS2)
77:                DoEvents
78:                strTemp = "IP UpDate: " & g_objSettings.DynDNS2 & " - " & successor: AddLog strTemp: stbMain.Panels(8).Text = strTemp
79:            End If
80:        End If
        
82:        If g_objSettings.DynDNS3 <> vbNullString Then
83:            HostIP = ResolveHostName(g_objSettings.DynDNS3)
84:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(7).Text
86:            If Not HubIP = HostIP Then
87:                successor = UpdDynDNS(g_objSettings.DynDNSUser, g_objSettings.DynDNSPass, g_objSettings.DynDNS3)
88:                DoEvents
89:                strTemp = "IP UpDate: " & g_objSettings.DynDNS3 & " - " & successor: AddLog strTemp: stbMain.Panels(8).Text = strTemp
90:            End If
91:        End If
        
93:        If g_objSettings.DynDNS4 <> vbNullString Then
94:            HostIP = ResolveHostName(g_objSettings.DynDNS4)
95:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(7).Text
97:            If Not HubIP = HostIP Then
98:                successor = UpdDynDNS(g_objSettings.DynDNSUser, g_objSettings.DynDNSPass, g_objSettings.DynDNS4)
99:                DoEvents
100:               strTemp = "IP UpDate: " & g_objSettings.DynDNS4 & " - " & successor: AddLog strTemp: stbMain.Panels(8).Text = strTemp
101:            End If
102:        End If
103:    End If

105:    Exit Sub
106:
Err:
107:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateDNSs"
End Sub
Public Function UpdNOIP(UserName, Password, HostName)
1:    On Error Resume Next
2:    UpdNOIP = UpdateNoIP(UserName, Password, HostName, "")
End Function
Public Function UpdDynDNS(UserName, Password, HostName)
1:    On Error Resume Next
2:    UpdDynDNS = UpdateDynDNS(UserName, Password, HostName, True, "", False)
End Function
Public Function ResolveHostName(HostName)
1:    On Error Resume Next
2:    ResolveHostName = ResolveHost(CStr(HostName))
End Function
Public Function DetectHubIP()
1:    On Error Resume Next
2:    Dim X: X = DetectIP()
3:    DetectHubIP = X
4:    If X = vbNullString Then X = "Connection refused by target machine"
5:    AddLog "Detect Hub IP: " & X
6:    stbMain.Panels(7).Text = X
End Function
'------------------------------------------------------------------------------
'End Update DSN/IP Subs
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'System Tray Subs
'------------------------------------------------------------------------------
Public Sub SysTrayUpDate(strToolTip As String)
1:    On Error GoTo Err
2:    ModifyTrayIcon Me, 111&, strToolTip
3:    Exit Sub
4:
Err:
5:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SysTrayUpDate(" & strToolTip & ")"
End Sub
Private Sub SysTrayAdd()
1:    On Error GoTo Err
2:    Dim strTmp As String
    
     'Set sounds to be enabled by default
5:    m_bSound = True
              
      'This gets us a globaly unique ID so that we can be sure the message
      'we use for getting our programs messages is unique
9:    WM_TRAYHOOK = RegisterWindowMessage(GetGUID())
    
      'This retrieves the window message for when the taskbar is created
      'since usually the application is run after the taskbar is created
      'it is safe to assume that if your program receives this message
      'any icon in the tray that was there is now gone and needs to be
      'recreated with a call to Shell_NotifyIcon(NIM_ADD, x)
16:    mTaskbarCreated = RegisterWindowMessage("PTDCH")
    
18:    If Len(g_objSettings.HubName) > 22 Then _
            strTmp = Left(g_objSettings.HubName, 20) & ".." _
       Else strTmp = g_objSettings.HubName
       
       'Create the tray icon
23:    CreateTrayIcon frmHub, 111&, "PT DC Hub " & vbVersion & vbNewLine & vbNewLine & strTmp
       
       'Start the message hook
26:    m_lHookID = InsertHook(frmHub)

28:    Exit Sub

30:
Err:
31:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SysTrayAdd()"
End Sub
Private Sub SysTrayRem()
1:    On Error GoTo Err

      'Remove system tray icon
4:    DeleteTrayIcon 111&
      'Remove the message hook  <=!!!IMPORTANT!!!
6:    RemoveHook frmHub, m_lHookID
      
8:  Exit Sub
    
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SysTrayRem()"
End Sub
Friend Function WindowProcSysTray( _
    ByVal shWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

7:    On Error GoTo Err

    'Friend Function for sys tray
    'This is our message handler
    
12:     If shWnd = Me.hwnd Then 'First we check to see if the message is for this window
           Select Case uMsg    'Then we look at the message
              Case mTaskbarCreated    'This message is for when the taskbar is created
                  'if the taskbar was created, chances are explorer.exe had crashed
14:                CreateTrayIcon Me, 111&, "PT Direct Connect Hub " & vbVersion, Me.Icon  'recreate the tray icon
        
               Case WM_TRAYHOOK 'Our user defined window message
                'if we get this we know that lParam carries the "event"
                'that occured on the tray icon
            
                Select Case lParam
                      Case WM_LBUTTONDBLCLK   'Left button dbl clicked
                    
20:                        If Me.WindowState = vbMinimized Then

22:                            SetForegroundWindow Me.hwnd
23:                            Me.WindowState = vbNormal
24:                            Me.Show
25:                        End If
                        
                      Case WM_RBUTTONUP   'Right button released
                    
28:                        SetForegroundWindow Me.hwnd
29:                        RemoveBalloon Me, 111&
30:                        PopupMenu Me.mnuPopUp(0)
                    
                    Case NIN_BALLOONUSERCLICK
                          'User clicked the balloon.
                          '
                    Case NIN_BALLOONTIMEOUT
                          'Balloon disapeared floated away, or was dismissed.
35:              End Select
    
37:        End Select
38:    End If

    'also pass them to VB
41:    WindowProcSysTray = CallWindowProc(m_lHookID, shWnd, uMsg, wParam, lParam)
    
43:  Exit Function

45:
Err:
46:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.WindowProcSysTray(""" & shWnd & """, """ & uMsg & """, """ & wParam & """, """ & lParam & """)"
End Function
'------------------------------------------------------------------------------
'End System Tray Subs
'------------------------------------------------------------------------------

Public Function IsListViewSelected(LVW As ListView) As Long
1:  Dim lvwItem     As ListItem
2:  Dim lvwItems    As ListItems
3:  On Error GoTo Err
4:    IsListViewSelected = -1

6:    If LVW.ListItems.Count > 0 Then
7:        Set lvwItem = LVW.SelectedItem
8:        Set lvwItems = LVW.ListItems

10:        If lvwItem.Selected Then
11:            IsListViewSelected = lvwItem.Index
12:        End If
13:    End If
    
15:   Exit Function
16:
Err:
17:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.IsListViewSelected()"
End Function

Private Sub cmbChangeLog_Click()
1:    On Error GoTo Err
2:    If cmbChangeLog.ListCount = 7 Then
3:       Static IsLoaded As Boolean
4:       If IsLoaded Then 'Not duplic the load data at start up..
            Select Case CInt(cmbChangeLog.ListIndex)
                Case 0: g_objAbout.SetVersion [All]
                Case 1: g_objAbout.SetVersion [0.x.x]
                Case 2: g_objAbout.SetVersion [1.0.x]
                Case 3: g_objAbout.SetVersion [1.1.x]
                Case 4: g_objAbout.SetVersion [1.2.x]
                Case 5: g_objAbout.SetVersion [1.3.x]
                Case 6: g_objAbout.SetVersion [1.4.x]
5:          End Select
6:       Else
7:          IsLoaded = True
8:       End If
9:    End If
10:   Exit Sub
11:
Err:
13    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbChangeLog_Click()"
End Sub

'------------------------------------------------------------------------------
'Paint events
'------------------------------------------------------------------------------
Private Sub Form_Paint()
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTileFormBackground Me, LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picBordTab_Paint(Index As Integer)
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picBordTab(Index), LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picHelp_Paint(Index As Integer)
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picHelp(Index), LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picInfo_Paint(Index As Integer)
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picInfo(Index), LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picITab_Paint(Index As Integer)
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picITab(Index), LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picSTab_Paint(Index As Integer)
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picSTab(Index), LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picStatus_Paint(Index As Integer)
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picStatus(Index), LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picStInfo_Paint()
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picStInfo, LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picTab_Paint(Index As Integer)
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picTab(Index), LoadImage(g_objSettings.lngSkin)
End Sub
Private Sub picTabAdv_Paint(Index As Integer)
1: On Error Resume Next
2: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picTabAdv(Index), LoadImage(g_objSettings.lngSkin)
End Sub
'------------------------------------------------------------------------------
'End Paint events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'End Move Form events
'------------------------------------------------------------------------------
Private Sub picBordTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub picHelp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub picInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub picITab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub picLog_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If Not Index = 2 Then _
        If g_objSettings.MoveForm Then _
              Call frmMove(Me)
End Sub
Private Sub picSTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub picStatus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub picStatus_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   picStInfo.Visible = False
End Sub
Private Sub picTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub picTabAdv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub
Private Sub stbMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub lblOptBanFilter_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub lblOptJM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub lblOptRedirect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub lblOptStSend_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
        Call frmMove(Me)
End Sub
Private Sub Labels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblCheck_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub
Private Sub lblHolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub
Private Sub lblShadowed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub
Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub
Private Sub lblStatistics_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1: On Error Resume Next
2:   If g_objSettings.MoveForm Then _
         Call frmMove(Me)
End Sub
'------------------------------------------------------------------------------
'End Move Form events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'Start ToolTipText
'------------------------------------------------------------------------------
Private Sub picLog_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error Resume Next
2:    If Index = 2 Then
3:      ShowToolTip picLog(2).hwnd, "If you would like to support development of PT DC Hub " & vbNewLine & _
                                    "You can use one or more of following proposals Sending arbitrary amount of money " & vbNewLine & _
                                    "Using PayPal to e-mail CarlosFerreiraCarlos@hotmail.com." & vbNewLine & vbNewLine & _
                                    "If you support PT DC Hub, you'll receive:" & vbNewLine & _
                                    "-Authorised access to PT DC Hub CVS." & vbNewLine & _
                                    "-Your name will be written in donators list in application and on PT DC Hub page. " & vbNewLine & vbNewLine & _
                                    "If you are interested in supporting of PT DC Hub, contact me on my e-mail." & vbNewLine & _
                                    "Thank you in advance for all support :)", "PT DC Hub - Donate", Tip_Normal, Tip_Info, &H0&, &H80FF80
10:   End If
End Sub
Private Sub lblCheck_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    If g_colToolTip.Exists("lblCheck(" & Index & ")") Then
3:         ShowTips g_arrToolTips(CInt(g_colToolTip.Item("lblCheck(" & Index & ")")))
4:    End If
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.txtData_MouseMove()"
End Sub
Private Sub lblHolder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    If g_colToolTip.Exists("lblHolder(" & Index & ")") Then
3:         ShowTips g_arrToolTips(CInt(g_colToolTip.Item("lblHolder(" & Index & ")")))
4:    End If
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lblCheck_MouseMove()"
End Sub
Private Sub lblOptStSend_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    If g_colToolTip.Exists("lblOptStSend(" & Index & ")") Then
3:         ShowTips g_arrToolTips(CInt(g_colToolTip.Item("lblOptStSend(" & Index & ")")))
4:    End If
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lblOptStSend_MouseMove()"
End Sub
Private Sub lblOptRedirect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    If g_colToolTip.Exists("lblOptRedirect(" & Index & ")") Then
3:         ShowTips g_arrToolTips(CInt(g_colToolTip.Item("lblOptRedirect(" & Index & ")")))
4:    End If
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lblOptRedirect_MouseMove()"
End Sub
Private Sub lblOptJM_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    If g_colToolTip.Exists("lblOptJM(" & Index & ")") Then
3:         ShowTips g_arrToolTips(CInt(g_colToolTip.Item("lblOptJM(" & Index & ")")))
4:    End If
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lblOptJM_MouseMove()"
End Sub
Private Sub lblOptBanFilter_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    If g_colToolTip.Exists("lblOptBanFilter(" & Index & ")") Then
3:         ShowTips g_arrToolTips(CInt(g_colToolTip.Item("lblOptBanFilter(" & Index & ")")))
4:    End If
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lblOptBanFilter_MouseMove()"
End Sub
Private Sub txtData_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    If g_colToolTip.Exists("txtData(" & Index & ")") Then
3:        ShowTips g_arrToolTips(CInt(g_colToolTip.Item("txtData(" & Index & ")"))), "txtData", Index
4:    End If
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.txtData_MouseMove()"
End Sub
Private Sub cmdButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    If g_colToolTip.Exists("cmdButton(" & Index & ")") Then
3:         ShowTips g_arrToolTips(CInt(g_colToolTip.Item("cmdButton(" & Index & ")"))), "cmdButton", Index
4:    End If
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdButton_MouseMove()"
End Sub
Private Function ShowTips(ByRef oToolTip As typToolTips, Optional ByVal sObject As String = Empty, Optional ByVal iIndex As Integer = -1)
3:    Dim lngHwnd     As Long
4:    Dim strText     As String
5:    Dim strTitle    As String
6:    Dim intStyle    As Integer
7:    Dim intIcon     As Integer
8:    On Error GoTo Err

    'Get parent control to set the rect
11:    If sObject = "txtData" Then
12:        lngHwnd = txtData(iIndex).hwnd
13:    ElseIf sObject = "cmdButton" Then
           lngHwnd = cmdButton(iIndex).hwnd
       Else
          Select Case True
            Case picTab(0).Visible: lngHwnd = picTab(0).hwnd
            Case picTab(1).Visible
                Select Case True
                    Case picSTab(0).Visible: lngHwnd = picSTab(0).hwnd
                    Case picSTab(1).Visible: lngHwnd = picSTab(1).hwnd
                    Case picSTab(2).Visible: lngHwnd = picSTab(2).hwnd
                    Case picSTab(3).Visible: lngHwnd = picSTab(3).hwnd
                    Case picSTab(4).Visible: lngHwnd = picSTab(4).hwnd
14:                End Select
            Case picTab(2).Visible
                Select Case True
                    Case picITab(0).Visible: lngHwnd = picITab(0).hwnd
                    Case picITab(1).Visible: lngHwnd = picITab(1).hwnd
                    Case picITab(2).Visible: lngHwnd = picITab(2).hwnd
                    Case picITab(3).Visible: lngHwnd = picITab(3).hwnd
                    Case picITab(4).Visible: lngHwnd = picITab(4).hwnd
                    Case picITab(5).Visible: lngHwnd = picITab(5).hwnd
15:                End Select
            Case picTab(3).Visible
                Select Case True
                    Case picTabAdv(0).Visible: lngHwnd = picTabAdv(0).hwnd
                    Case picTabAdv(1).Visible: lngHwnd = picTabAdv(1).hwnd
                    Case picTabAdv(2).Visible: lngHwnd = picTabAdv(2).hwnd
                    Case picTabAdv(3).Visible: lngHwnd = picTabAdv(3).hwnd
                    Case picTabAdv(4).Visible: lngHwnd = picTabAdv(4).hwnd
16:                End Select
            Case picTab(4).Visible: lngHwnd = picTab(4).hwnd
            Case picTab(5).Visible: lngHwnd = tlbScript.hwnd
            Case picTab(6).Visible
                Select Case True
                    Case picStatus(0).Visible: lngHwnd = picStatus(0).hwnd
                    Case picStatus(1).Visible: lngHwnd = picStatus(1).hwnd
                    Case picStatus(2).Visible: lngHwnd = picStatus(2).hwnd
                    Case picStatus(3).Visible: lngHwnd = picStatus(3).hwnd
                    Case picStatus(4).Visible: lngHwnd = picStatus(4).hwnd
                    Case picStatus(5).Visible: lngHwnd = picStatus(5).hwnd
17:                End Select
            Case picTab(7).Visible
                Select Case True
                    Case picInfo(0).Visible: lngHwnd = picInfo(0).hwnd
                    Case picInfo(1).Visible: lngHwnd = picInfo(1).hwnd
18:                End Select
            Case picTab(8).Visible
                Select Case True
                    Case picHelp(0).Visible: lngHwnd = picHelp(0).hwnd
                    Case picHelp(1).Visible: lngHwnd = picHelp(1).hwnd
                    Case picHelp(2).Visible: lngHwnd = picHelp(2).hwnd
19:                End Select
20:        End Select
21:    End If

       'Show tool tip text
23:    ShowToolTip lngHwnd, oToolTip.sMessage, oToolTip.sTitle, oToolTip.iStyle, oToolTip.iIcon
    
26:    Exit Function
27:
Err:
28:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.ShowTips()"
End Function
'------------------------------------------------------------------------------
'End ToolTipText
'------------------------------------------------------------------------------
