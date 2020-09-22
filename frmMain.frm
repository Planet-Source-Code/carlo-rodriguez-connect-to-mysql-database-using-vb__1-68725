VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MYSQL Connector"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Field(s) List:"
      Height          =   3465
      Left            =   9540
      TabIndex        =   20
      Top             =   2430
      Width           =   2385
      Begin VB.ListBox lstFields 
         Appearance      =   0  'Flat
         Height          =   2970
         Left            =   150
         TabIndex        =   21
         Top             =   270
         Width           =   2115
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Query Results"
      Height          =   3465
      Left            =   60
      TabIndex        =   17
      Top             =   2430
      Width           =   9435
      Begin MSDataGridLib.DataGrid dg 
         Height          =   2895
         Left            =   240
         TabIndex        =   18
         Top             =   300
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Table List:"
      Height          =   2295
      Left            =   9540
      TabIndex        =   12
      Top             =   90
      Width           =   2385
      Begin VB.ListBox Listtble 
         Appearance      =   0  'Flat
         Height          =   1710
         Left            =   150
         TabIndex        =   13
         Top             =   270
         Width           =   2115
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MYSQL Command:"
      Height          =   2295
      Left            =   3990
      TabIndex        =   11
      Top             =   90
      Width           =   5505
      Begin VB.CommandButton cmdExecute 
         Appearance      =   0  'Flat
         Caption         =   "Run Query"
         Height          =   375
         Left            =   4230
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1800
         Width           =   1155
      End
      Begin VB.TextBox txtQuery 
         Appearance      =   0  'Flat
         Height          =   1425
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   300
         Width           =   5205
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection Settings:"
      Height          =   2295
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   3885
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Go"
         Height          =   405
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1740
         Width           =   585
      End
      Begin VB.TextBox txtDBPort 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         TabIndex        =   10
         Text            =   "3306"
         Top             =   1740
         Width           =   825
      End
      Begin VB.TextBox txtDBPass 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2220
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1380
         Width           =   1515
      End
      Begin VB.TextBox txtDBU 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         TabIndex        =   8
         Top             =   1020
         Width           =   1515
      End
      Begin VB.TextBox txtDBName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         TabIndex        =   7
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         TabIndex        =   6
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MYSQL Port:"
         Height          =   210
         Left            =   300
         TabIndex        =   5
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MYSQL DB Password:"
         Height          =   210
         Left            =   300
         TabIndex        =   4
         Top             =   1440
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MYSQL DB Username:"
         Height          =   210
         Left            =   300
         TabIndex        =   3
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MYSQL Database Name:"
         Height          =   210
         Left            =   300
         TabIndex        =   2
         Top             =   690
         Width           =   1785
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server IP:"
         Height          =   210
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   705
      End
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   19
      Top             =   5970
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7038
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7038
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7038
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset

Private Sub cmdConnect_Click()
    OpenConn txtIP, txtDBName, txtDBU, txtDBPass, txtDBPort
    
    If conn.State <> 0 Then
        getTables Listtble
    End If
End Sub

Private Sub cmdExecute_Click()
On Error GoTo errH
    If rs.State <> 0 Then rs.Close
    
    rs.CursorLocation = adUseClient
    rs.Open txtQuery.Text, conn, adOpenKeyset, adLockOptimistic
    getFields rs, lstFields
    If rs.RecordCount > 0 Then
        Set dg.DataSource = rs
        dg.Refresh
        sb.Panels(1).Text = rs.RecordCount & " records found"
    End If
    
    Exit Sub
errH:
MsgBox Err.Description, vbCritical, "Error!"
    
End Sub

