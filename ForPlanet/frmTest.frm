VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDose 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4560
      Width           =   975
   End
   Begin VB.ComboBox cboUse 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtDrug 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtTest 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgOrder 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7011
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "You can Press Enter Key From Here I Prefer Not to Use Mouse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cx As Integer
Dim Cy As Integer
Dim Cz As Integer


Private Sub cboUse_GotFocus()
    BGFocus cboUse
    HFocus cboUse
End Sub

Private Sub cboUse_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mfgOrder.TextMatrix(mfgOrder.Row, 2) = cboUse.Text
        cboUse.Visible = False
        mfgOrder.Col = 3
        txtDose.Move mfgOrder.Left + mfgOrder.CellLeft - Cx, mfgOrder.Top + mfgOrder.CellTop - Cy, mfgOrder.CellWidth, mfgOrder.CellHeight - Cz
        txtDose.Text = mfgOrder.TextMatrix(mfgOrder.Row, 3)
        txtDose.Visible = True
        txtDose.SetFocus
    End If
End Sub

Private Sub cboUse_LostFocus()
    BGLoss cboUse
End Sub

Private Sub Form_Load()
    Cx = 10
    Cy = 10
    Cz = 20
    ' that only for correction
    With mfgOrder
        If .Rows < 2 Then .Rows = 2
        If .Cols < 4 Then .Cols = 4
    End With
    OrderGrid
    txtDrug.Visible = False
    cboUse.Visible = False
    txtDose.Visible = False
    
    With cboUse
        .AddItem "Test1"
        .AddItem "Test2"
        .AddItem "Other"
    End With
    
End Sub
Private Sub OrderGrid()
    With mfgOrder
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 4000
        
    End With
End Sub

Private Sub mfgOrder_Click()
    If txtDrug.Visible = True Then txtDrug.Visible = False
    If cboUse.Visible = True Then cboUse.Visible = False
    If txtDose.Visible = True Then txtDose.Visible = False
    With mfgOrder
        If .Col = 1 Then
            txtDrug.Move mfgOrder.Left + mfgOrder.CellLeft - Cx, mfgOrder.Top + mfgOrder.CellTop - Cy, mfgOrder.CellWidth, mfgOrder.CellHeight - Cz
            txtDrug.Text = mfgOrder.TextMatrix(mfgOrder.Row, 1)
            txtDrug.Visible = True
            txtDrug.SetFocus
        ElseIf .Col = 2 Then
            cboUse.Move mfgOrder.Left + mfgOrder.CellLeft - Cx, mfgOrder.Top + mfgOrder.CellTop - Cy, mfgOrder.CellWidth
            cboUse.Text = mfgOrder.TextMatrix(mfgOrder.Row, 2)
            cboUse.Visible = True
            cboUse.SetFocus
        ElseIf .Col = 3 Then
            txtDose.Move mfgOrder.Left + mfgOrder.CellLeft - Cx, mfgOrder.Top + mfgOrder.CellTop - Cy, mfgOrder.CellWidth, mfgOrder.CellHeight - Cz
            txtDose.Text = mfgOrder.TextMatrix(mfgOrder.Row, 3)
            txtDose.Visible = True
            txtDose.SetFocus
        End If
        
    End With
End Sub

Private Sub txtDose_GotFocus()
    HighLight
End Sub

Private Sub txtDose_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mfgOrder.TextMatrix(mfgOrder.Row, 3) = txtDose.Text
        txtDose.Visible = False
        mfgOrder.Rows = mfgOrder.Rows + 1
        mfgOrder.Row = mfgOrder.Rows - 1
        mfgOrder.Col = 1
        txtDrug.Move mfgOrder.Left + mfgOrder.CellLeft - Cx, mfgOrder.Top + mfgOrder.CellTop - Cy, mfgOrder.CellWidth, mfgOrder.CellHeight - Cz
        txtDrug.Text = mfgOrder.TextMatrix(mfgOrder.Row, 1)
        txtDrug.Visible = True
        txtDrug.SetFocus
    End If
End Sub

Private Sub txtDose_LostFocus()
    BGLoss txtDose
End Sub
Public Sub BGFocus(anyControl As Control)
    
    If anyControl.Locked = False Then
        anyControl.BackColor = &H80000013
    End If
End Sub
Public Sub BGLoss(anyControl As Control)
    
    If anyControl.Locked = False Then
        anyControl.BackColor = &H80000005
    End If
End Sub
Public Sub HFocus(ByRef sText As Variant)
    
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
    
End Sub
Private Sub txtDrug_GotFocus()
    HighLight
End Sub
Public Sub HighLight()
    With Screen.ActiveForm
        If (TypeOf .ActiveControl Is TextBox) Then
            .ActiveControl.SelStart = 0
            .ActiveControl.SelLength = Len(.ActiveControl)
            .ActiveControl.BackColor = &H80000013
        End If
    End With
End Sub
Private Sub txtDrug_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mfgOrder.TextMatrix(mfgOrder.Row, 1) = txtDrug.Text
        txtDrug.Visible = False
        mfgOrder.Col = 2
        cboUse.Move mfgOrder.Left + mfgOrder.CellLeft - Cx, mfgOrder.Top + mfgOrder.CellTop - Cy, mfgOrder.CellWidth
        cboUse.Text = mfgOrder.TextMatrix(mfgOrder.Row, 2)
        cboUse.Visible = True
        cboUse.SetFocus
    End If
End Sub

Private Sub txtDrug_LostFocus()
    BGLoss txtDrug
    
End Sub


Private Sub txtTest_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mfgOrder.Col = 1
        txtDrug.Move mfgOrder.Left + mfgOrder.CellLeft - Cx, mfgOrder.Top + mfgOrder.CellTop - Cy, mfgOrder.CellWidth, mfgOrder.CellHeight - Cz
        txtDrug.Text = mfgOrder.TextMatrix(mfgOrder.Row, 1)
        txtDrug.Visible = True
        txtDrug.SetFocus
    End If
End Sub
