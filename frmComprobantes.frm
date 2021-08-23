VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circuito de Comprobantes"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14865
   Icon            =   "frmComprobantes.frx":0000
   LinkTopic       =   "frmComprobantes"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   14865
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   7785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   13732
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "Comprobantes"
      TabPicture(0)   =   "frmComprobantes.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.OptionButton Option5 
         Height          =   195
         Left            =   -71190
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   810
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Height          =   195
         Left            =   -72210
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   810
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filtros"
         Height          =   885
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   330
         Width           =   14670
         Begin VB.CommandButton Command 
            Height          =   585
            Left            =   12930
            Picture         =   "frmComprobantes.frx":08E6
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   180
            Width           =   1245
         End
         Begin VB.ComboBox cmbUsuarios 
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   300
            Width           =   1725
         End
         Begin VB.ComboBox cmbCV 
            Height          =   315
            Left            =   2190
            TabIndex        =   10
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txtFecha 
            Height          =   375
            Left            =   10890
            TabIndex        =   14
            Top             =   300
            Width           =   1425
         End
         Begin VB.TextBox txtNumero 
            Height          =   375
            Left            =   8790
            TabIndex        =   13
            Top             =   300
            Width           =   1425
         End
         Begin VB.TextBox txtSucursal 
            Height          =   375
            Left            =   6570
            TabIndex        =   12
            Top             =   300
            Width           =   1425
         End
         Begin VB.TextBox txtComprobante 
            Height          =   375
            Left            =   4410
            TabIndex        =   11
            Top             =   300
            Width           =   1425
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo:"
            Height          =   375
            Left            =   1800
            TabIndex        =   17
            Top             =   330
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha:"
            Height          =   375
            Left            =   10320
            TabIndex        =   16
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label2 
            Caption         =   "Número:"
            Height          =   375
            Left            =   8100
            TabIndex        =   8
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal:"
            Height          =   375
            Left            =   5880
            TabIndex        =   7
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label 
            Caption         =   "Cod.Comprobante:"
            Height          =   375
            Left            =   3060
            TabIndex        =   6
            Top             =   360
            Width           =   1425
         End
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   6480
         Left            =   60
         TabIndex        =   4
         Top             =   1230
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   11430
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "circuito"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Comprobante ROOT"
            Object.Width           =   5203
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "1"
            Object.Width           =   5203
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "2"
            Object.Width           =   5203
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "3"
            Object.Width           =   5203
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "4"
            Object.Width           =   5203
         EndProperty
      End
      Begin VB.Label lblUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71310
         TabIndex        =   5
         Top             =   420
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmComprobantes
' DateTime  : 06/2016
' Author    : Juan José Sortino
' Purpose   : Observar el circuito de Comprobantes
'---------------------------------------------------------------------------------------
Option Explicit
'***********************************************************************
' Constantes Propias
'***********************************************************************
Private Const Si                  As String = "Sí"
Private Const No                  As String = "No"
Private Const NullString          As String = ""
Private Const UNKNOWN_ERRORSOURCE As String = "[Fuente de Error Desconocida]"
Private Const KNOWN_ERRORSOURCE   As String = "[Fuente de Error Conocida]"
'***********************************************************************
Private cnn  As ADODB.Connection
Private rstCircuito As ADODB.Recordset
Private rstCircuitoCopia As ADODB.Recordset
Private rst  As ADODB.Recordset
Private SQL  As String
Private sqlTemp As String
Private itmX As ListItem
Private iEmpresas As Integer
Private strEmpresa As String

Private ErrorLog            As ErrType

'-----------GetRegistryValue
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Enum REGRootTypesEnum
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL = &H1F0000

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_READ = ((READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

'-----------GetRegistryValue

Private mvarControlData     As DataShare.udtControlData         'información de control
Private strState            As String * 50000

Private tliApp              As Object
Private tliApp2             As Object
Private objObjeto           As Object
Private bCargando           As Boolean
Private bBuscando           As Boolean
Private sngCoordenadaX      As Single
Private sngCoordenadaY      As Single
'Private itmX                As ListItem
Private bStop               As Boolean

Private strCV               As String
Private strComprobante      As String
Private iSucursal           As Integer
Private lngNumero           As Long
Private dFecha              As Date
Private bFiltros            As Boolean

Private Sub Command_Click()
   On Error GoTo GestErr
   
   bCargando = True
   
   CargarRecordset
      
   bCargando = False
   
   Exit Sub

GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[Command_Click]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Form_Load()

   On Error GoTo GestErr
   
   Cargar_cmbCV
   CargarUsuarios
   Exit Sub

GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[Form_Load]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Cargar_cmbCV()
   cmbCV.AddItem "V"
   cmbCV.AddItem "C"
   cmbCV.ListIndex = 0
End Sub

Private Sub ListView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   sngCoordenadaX = X
   sngCoordenadaY = Y
End Sub

Private Sub ListView_Click()
   On Error Resume Next
   
   Set itmX = ListView.HitTest(sngCoordenadaX, sngCoordenadaY)
   If itmX Is Nothing Then Exit Sub
   
   Clipboard.Clear
   Clipboard.SetText itmX.SubItems(1) & " " & itmX.SubItems(2) & " " & itmX.SubItems(3)
End Sub
Private Sub CargarRecordset()

   On Error GoTo GestErr
   
   Set cnn = New ADODB.Connection
   cnn.ConnectionString = "Provider=MSDataShape;Data Provider=MSDAORA;Password=apfrms2001;User ID=" & cmbUsuarios.Text & ";Data Source=BASE"
   cnn.Open
   
   Set rstCircuito = New ADODB.Recordset
   rstCircuito.CursorLocation = adUseClient
   rstCircuito.LockType = adLockBatchOptimistic
   rstCircuito.CursorType = adOpenStatic
   
   Set rstCircuitoCopia = New ADODB.Recordset
   rstCircuitoCopia.CursorLocation = adUseClient
   rstCircuitoCopia.LockType = adLockBatchOptimistic
   rstCircuitoCopia.CursorType = adOpenStatic
   
   sqlTemp = " SELECT HCA_RENGLON AS CIRCUITO, HCA_RENGLON AS INSTANCIA, HCA_ARTICULO AS ORDEN, HCA_COMPRA_VENTA, HCA_CODIGO, HCA_SUCURSAL, HCA_NUMERO, HCA_FECHA, HCA_RENGLON "
   sqlTemp = sqlTemp & "  FROM HISTORICO_COMP_CV_ARTICULOS "
   sqlTemp = sqlTemp & " WHERE 0 = 1"
   rstCircuito.Open sqlTemp, cnn
   rstCircuitoCopia.Open sqlTemp, cnn

   ListView.ListItems.Clear

   CargarROOTs
   If rstCircuito.RecordCount > 0 Then
      CargarRecordsetHijos
      LlenarListView
   End If
   
   Exit Sub

GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[CargarRecordset]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub CargarROOTs()
Dim ix As Double

   On Error GoTo GestErr
   
   VerFiltros

   Set rst = New ADODB.Recordset
   rst.CursorLocation = adUseClient
   rst.LockType = adLockReadOnly
   rst.CursorType = adOpenStatic
   If Not bFiltros Then rst.MaxRecords = 100
   
   SQL = " SELECT HCA_COMPRA_VENTA, HCA_CODIGO, HCA_SUCURSAL, HCA_NUMERO, HCA_FECHA, HCA_RENGLON "
   SQL = SQL & "  FROM HISTORICO_COMP_CV_ARTICULOS "
   SQL = SQL & " WHERE HISTORICO_COMP_CV_ARTICULOS.HCA_CV_ORIGEN = HISTORICO_COMP_CV_ARTICULOS.HCA_COMPRA_VENTA "
   SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_COMP_ORIGEN = HISTORICO_COMP_CV_ARTICULOS.HCA_CODIGO "
   SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_SUC_ORIGEN = HISTORICO_COMP_CV_ARTICULOS.HCA_SUCURSAL "
   SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_NRO_ORIGEN = HISTORICO_COMP_CV_ARTICULOS.HCA_NUMERO "
   SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA_ORIGEN = HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA "
   SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_RENGLON_ORIGEN = HISTORICO_COMP_CV_ARTICULOS.HCA_RENGLON "
   
   'Aplicar filtros de pantalla
   SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_COMPRA_VENTA = '" & cmbCV.Text & "' "
   If Len(txtComprobante.Text) > 0 Then
      SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_CODIGO = '" & strComprobante & "' "
   End If
   If Len(txtSucursal.Text) > 0 Then
      SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_SUCURSAL = " & iSucursal
   End If
   If Len(txtNumero.Text) > 0 Then
      SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_NUMERO = " & lngNumero
   End If
   If Len(txtFecha.Text) > 0 Then
      SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA = TO_DATE ('" & dFecha & "', 'DD/MM/YYYY') "
   End If

   SQL = SQL & " ORDER BY HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA DESC "
   
   rst.Open SQL, cnn
   
   If rst.RecordCount > 0 Then
      rst.MoveFirst
      ix = 0
      Do While Not rst.EOF
         ix = ix + 1
         rstCircuito.AddNew
         
         rstCircuito("CIRCUITO").Value = ix
         rstCircuito("INSTANCIA").Value = 1
         rstCircuito("ORDEN").Value = 1
         rstCircuito("HCA_COMPRA_VENTA").Value = rst("HCA_COMPRA_VENTA").Value
         rstCircuito("HCA_CODIGO").Value = rst("HCA_CODIGO").Value
         rstCircuito("HCA_SUCURSAL").Value = rst("HCA_SUCURSAL").Value
         rstCircuito("HCA_NUMERO").Value = rst("HCA_NUMERO").Value
         rstCircuito("HCA_FECHA").Value = rst("HCA_FECHA").Value
         rstCircuito("HCA_RENGLON").Value = rst("HCA_RENGLON").Value
         rstCircuito.Update
           
         rst.MoveNext
      Loop
   End If
   rst.Close
   
   Exit Sub

GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[CargarROOTs]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub CargarRecordsetHijos()
Dim ix   As Double
Dim iz   As Double
Dim bContinuar As Boolean

   On Error GoTo GestErr
   
   bContinuar = True
   ix = 1
   Do While bContinuar
      rstCircuito.MoveFirst
      
      bContinuar = False
      
      rstCircuito.Filter = "INSTANCIA = " & ix
      ix = ix + 1
      
      Do While Not rstCircuito.EOF
         Set rst = New ADODB.Recordset
         rst.CursorLocation = adUseClient
         rst.LockType = adLockReadOnly
         rst.CursorType = adOpenStatic
   
         SQL = " SELECT HISTORICO_COMP_CV_ARTICULOS.HCA_COMPRA_VENTA, HISTORICO_COMP_CV_ARTICULOS.HCA_CODIGO, "
         SQL = SQL & "       HISTORICO_COMP_CV_ARTICULOS.HCA_SUCURSAL, HISTORICO_COMP_CV_ARTICULOS.HCA_NUMERO, "
         SQL = SQL & "       HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA, HISTORICO_COMP_CV_ARTICULOS.HCA_RENGLON "
         SQL = SQL & "  FROM HISTORICO_COMP_CV_ARTICULOS "
         SQL = SQL & " WHERE HISTORICO_COMP_CV_ARTICULOS.HCA_COMP_ORIGEN <> HISTORICO_COMP_CV_ARTICULOS.HCA_CODIGO "
         SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_SUC_ORIGEN <> HISTORICO_COMP_CV_ARTICULOS.HCA_SUCURSAL "
         SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_NRO_ORIGEN <> HISTORICO_COMP_CV_ARTICULOS.HCA_NUMERO "
         'SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA_ORIGEN <> HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA "
         SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_CV_ORIGEN = '" & rstCircuito("HCA_COMPRA_VENTA").Value & "' "
         SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_COMP_ORIGEN = '" & rstCircuito("HCA_CODIGO").Value & "' "
         SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_SUC_ORIGEN = " & rstCircuito("HCA_SUCURSAL").Value
         SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_NRO_ORIGEN = " & rstCircuito("HCA_NUMERO").Value
         SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA_ORIGEN = TO_DATE('" & Format(rstCircuito("HCA_FECHA").Value, "DD-MM-YYYY") & "', 'DD-MM-YYYY') "
         SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_RENGLON_ORIGEN = " & rstCircuito("HCA_RENGLON").Value
         rst.Open SQL, cnn
         
         If rst.RecordCount > 0 Then
            bContinuar = True
            iz = 0
            Do While Not rst.EOF
               iz = iz + 1
               rstCircuitoCopia.AddNew
               
               rstCircuitoCopia("CIRCUITO").Value = rstCircuito("CIRCUITO").Value
               rstCircuitoCopia("INSTANCIA").Value = ix
               rstCircuitoCopia("ORDEN").Value = rstCircuito("ORDEN").Value & "." & iz
               rstCircuitoCopia("HCA_COMPRA_VENTA").Value = rst("HCA_COMPRA_VENTA").Value
               rstCircuitoCopia("HCA_CODIGO").Value = rst("HCA_CODIGO").Value
               rstCircuitoCopia("HCA_SUCURSAL").Value = rst("HCA_SUCURSAL").Value
               rstCircuitoCopia("HCA_NUMERO").Value = rst("HCA_NUMERO").Value
               rstCircuitoCopia("HCA_FECHA").Value = rst("HCA_FECHA").Value
               rstCircuitoCopia("HCA_RENGLON").Value = rst("HCA_RENGLON").Value
               rstCircuitoCopia.Update
         
               rst.MoveNext
            Loop
         End If
         rst.Close
         rstCircuito.MoveNext
      Loop
      
      'act rstCircuito con rstCircuitoCopia
      If rstCircuitoCopia.RecordCount > 0 Then
         rstCircuitoCopia.MoveFirst
         Do While Not rstCircuitoCopia.EOF
            rstCircuito.AddNew
            
            rstCircuito("CIRCUITO").Value = rstCircuitoCopia("CIRCUITO").Value
            rstCircuito("INSTANCIA").Value = rstCircuitoCopia("INSTANCIA").Value
            rstCircuito("ORDEN").Value = rstCircuitoCopia("ORDEN").Value
            rstCircuito("HCA_COMPRA_VENTA").Value = rstCircuitoCopia("HCA_COMPRA_VENTA").Value
            rstCircuito("HCA_CODIGO").Value = rstCircuitoCopia("HCA_CODIGO").Value
            rstCircuito("HCA_SUCURSAL").Value = rstCircuitoCopia("HCA_SUCURSAL").Value
            rstCircuito("HCA_NUMERO").Value = rstCircuitoCopia("HCA_NUMERO").Value
            rstCircuito("HCA_FECHA").Value = rstCircuitoCopia("HCA_FECHA").Value
            rstCircuito("HCA_RENGLON").Value = rstCircuitoCopia("HCA_RENGLON").Value
            rstCircuito.Update
      
            rstCircuitoCopia.MoveNext
         Loop
      End If
      
      Set rstCircuitoCopia = New ADODB.Recordset
      rstCircuitoCopia.CursorLocation = adUseClient
      rstCircuitoCopia.LockType = adLockBatchOptimistic
      rstCircuitoCopia.CursorType = adOpenStatic
      rstCircuitoCopia.Open sqlTemp, cnn
   Loop
   
   Exit Sub

GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[CargarRecordsetHijos]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub LlenarListView()
Dim ix         As Double
Dim bContinuar As Boolean
Dim strItem    As String
Dim iUltimaCol As Integer

   On Error GoTo GestErr
   
   bContinuar = True
   ix = 1
   
   Do While bContinuar
      rstCircuito.MoveFirst
      
      bContinuar = False
      
      rstCircuito.Filter = "INSTANCIA = " & ix
      rstCircuito.Sort = "ORDEN, HCA_COMPRA_VENTA, HCA_CODIGO,HCA_SUCURSAL,HCA_NUMERO,HCA_FECHA,HCA_RENGLON asc"

      ix = ix + 1
      
      Do While Not rstCircuito.EOF
         If rstCircuito.RecordCount > 0 Then
            bContinuar = True
            
            strItem = rstCircuito("ORDEN").Value & " " & rstCircuito("HCA_CODIGO").Value & " " & Format(rstCircuito("HCA_SUCURSAL").Value, "0000") & "-" & Format(rstCircuito("HCA_NUMERO").Value, "00000000") & " " & rstCircuito("HCA_FECHA").Value & "-" & rstCircuito("HCA_RENGLON").Value
            Select Case rstCircuito("INSTANCIA").Value
               Case 1
                  Set itmX = ListView.ListItems.Add
                  itmX.Text = rstCircuito("CIRCUITO").Value
                  itmX.SubItems(1) = strItem
                  iUltimaCol = 1
               Case 2
                  For Each itmX In ListView.ListItems
                     If itmX.Text = rstCircuito("CIRCUITO").Value Then
                        If Len(itmX.SubItems(2)) > 0 Then
                           itmX.SubItems(2) = itmX.SubItems(2) & vbCrLf
                        End If
                        itmX.SubItems(2) = itmX.SubItems(2) & strItem
                        iUltimaCol = 2
                     End If
                  Next itmX
               Case 3
                  For Each itmX In ListView.ListItems
                     If itmX.Text = rstCircuito("CIRCUITO").Value Then
                        If Len(itmX.SubItems(3)) > 0 Then
                           itmX.SubItems(3) = itmX.SubItems(3) & vbCrLf
                        End If
                        itmX.SubItems(3) = itmX.SubItems(3) & strItem
                        iUltimaCol = 3
                     End If
                  Next itmX
               Case 4
                  For Each itmX In ListView.ListItems
                     If itmX.Text = rstCircuito("CIRCUITO").Value Then
                        If Len(itmX.SubItems(4)) > 0 Then
                           itmX.SubItems(4) = itmX.SubItems(4) & vbCrLf
                        End If
                        itmX.SubItems(4) = itmX.SubItems(4) & strItem
                        iUltimaCol = 4
                     End If
                  Next itmX
               Case 5
                  For Each itmX In ListView.ListItems
                     If itmX.Text = rstCircuito("CIRCUITO").Value Then
                        If Len(itmX.SubItems(5)) > 0 Then
                           itmX.SubItems(5) = itmX.SubItems(5) & vbCrLf
                        End If
                        itmX.SubItems(5) = itmX.SubItems(5) & strItem
                        iUltimaCol = 5
                     End If
                  Next itmX
            End Select
         End If
         rstCircuito.MoveNext
      Loop
   Loop
   
   Exit Sub

GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[LlenarListView]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub VerFiltros()
   
   On Error GoTo GestErr

   bFiltros = True
   
   Set rst = New ADODB.Recordset
   rst.CursorLocation = adUseClient
   rst.LockType = adLockReadOnly
   rst.CursorType = adOpenStatic
   
   SQL = " SELECT HISTORICO_COMP_CV_ARTICULOS.HCA_ROOT_COMPRA_VENTA, HISTORICO_COMP_CV_ARTICULOS.HCA_ROOT_CODIGO, "
   SQL = SQL & "       HISTORICO_COMP_CV_ARTICULOS.HCA_ROOT_SUCURSAL, HISTORICO_COMP_CV_ARTICULOS.HCA_ROOT_NUMERO, "
   SQL = SQL & "       HISTORICO_COMP_CV_ARTICULOS.HCA_ROOT_FECHA "
   SQL = SQL & "  FROM HISTORICO_COMP_CV_ARTICULOS "
   SQL = SQL & " WHERE 1 = 1 "
   SQL = SQL & "      AND HISTORICO_COMP_CV_ARTICULOS.HCA_COMPRA_VENTA = '" & cmbCV.Text & "' "
   If Len(txtComprobante.Text) > 0 Then
      SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_CODIGO = '" & txtComprobante.Text & "' "
   Else
      bFiltros = False
   End If
   If Len(txtSucursal.Text) > 0 Then
      SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_SUCURSAL = " & txtSucursal.Text
   Else
      bFiltros = False
   End If
   If Len(txtNumero.Text) > 0 Then
      SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_NUMERO = " & txtNumero.Text
   Else
      bFiltros = False
   End If
   If Len(txtFecha.Text) > 0 Then
      SQL = SQL & "   AND HISTORICO_COMP_CV_ARTICULOS.HCA_FECHA = TO_DATE ('" & txtFecha.Text & "', 'DD/MM/YYYY') "
   Else
      bFiltros = False
   End If
      
   rst.Open SQL, cnn
   
   If rst.RecordCount > 0 Then
      strCV = rst("HCA_ROOT_COMPRA_VENTA").Value
      If Len(txtComprobante.Text) > 0 Then
         strComprobante = rst("HCA_ROOT_CODIGO").Value
      End If
      If Len(txtSucursal.Text) > 0 Then
         iSucursal = rst("HCA_ROOT_SUCURSAL").Value
      End If
      If Len(txtNumero.Text) > 0 Then
         lngNumero = rst("HCA_ROOT_NUMERO").Value
      End If
      If Len(txtFecha.Text) > 0 Then
         dFecha = rst("HCA_ROOT_FECHA").Value
      End If
   Else
      MsgBox "No se encuentra el comprobante"
   End If
   rst.Close
   
   Exit Sub

GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[VerFiltros]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub CargarUsuarios()
Dim rstGlobal As ADODB.Recordset
Dim cnnG      As ADODB.Connection

   On Error GoTo GestErr
   
   Set cnnG = New ADODB.Connection
   cnnG.ConnectionString = GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Algoritmo\DatabaseSettings\ConnectionsStrings", "ALG", REG_SZ, "", False)
   If Len(cnnG.ConnectionString) = 0 Then cnnG.ConnectionString = "Provider=MSDataShape;Data Provider=MSDAORA;Password=apfrms2001;User ID=SYSADMIN_E66;Data Source=BASE"
   cnnG.Open
   
   Set rstGlobal = New ADODB.Recordset
   rstGlobal.CursorLocation = adUseClient
   rstGlobal.LockType = adLockReadOnly
   rstGlobal.CursorType = adOpenStatic
   
   SQL = " SELECT   DBA_USERS.USERNAME, DBA_USERS.CREATED, EMPRESAS.EMP_DESCRIPCION "
   SQL = SQL & "    FROM DBA_USERS, "
   SQL = SQL & "         EMPRESAS "
   SQL = SQL & "   WHERE USERNAME LIKE 'SYSADMIN%' "
   SQL = SQL & "     AND SUBSTR (DBA_USERS.USERNAME, 10, 3) = EMPRESAS.EMP_CODIGO_EMPRESA(+) "
   SQL = SQL & "ORDER BY USERNAME "
   rstGlobal.Open SQL, cnnG
      
   rstGlobal.MoveFirst
   Do While Not rstGlobal.EOF
      cmbUsuarios.AddItem IIf(IsNull(rstGlobal("USERNAME").Value), "", rstGlobal("USERNAME").Value)

      rstGlobal.MoveNext
   Loop
   
   cmbUsuarios.ListIndex = 0
   
   rstGlobal.Close
   
   Exit Sub
GestErr:
   Screen.MousePointer = vbNormal
   MsgBox "[CargarUsuarios]" & vbCrLf & Err.Description & Erl
End Sub
Private Function GetRegistryValue(ByVal hKey As REGRootTypesEnum, ByVal KeyName As String, ByVal ValueName As String, Optional ByVal KeyType As REGKeyTypesEnum, Optional DefaultValue As Variant, Optional ByVal Create As Boolean) As Variant
      Dim handle As Long, resLong As Long
      Dim resString As String, length As Long
      Dim resBinary() As Byte

10       On Error GoTo GestErr

20       If KeyType = 0 Then
30          KeyType = REG_SZ
40       End If

50       If IsMissing(DefaultValue) Then
60          Select Case KeyType
               Case REG_SZ
70                DefaultValue = ""
80             Case REG_DWORD
90                DefaultValue = 0
100            Case REG_BINARY
110               DefaultValue = 0
120         End Select
130      End If

         ' Prepare the default result.
140      GetRegistryValue = DefaultValue
         ' Open the key, exit if not found.
150      If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
160         If Create Then
               'si no exite la creo
170            If CreateRegistryKey(hKey, KeyName) Then Exit Function
180            If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
190         Else
200            Exit Function
210         End If
220      End If

230      Select Case KeyType
             Case REG_DWORD
                 ' Read the value, use the default if not found.
240              If RegQueryValueEx(handle, ValueName, 0, REG_DWORD, resLong, 4) = 0 Then
250                  GetRegistryValue = resLong
260               Else
270                  If Create Then
280                     SetRegistryValue hKey, KeyName, ValueName, REG_DWORD, DefaultValue
290                  End If
300              End If
310          Case REG_SZ
320              length = 1024: resString = Space$(length)
330              If RegQueryValueEx(handle, ValueName, 0, REG_SZ, ByVal resString, length) = 0 Then
                     ' If value is found, trim characters in excess.
340                  GetRegistryValue = Left$(resString, length - 1)
350               Else
360                  If Create Then
370                     SetRegistryValue hKey, KeyName, ValueName, REG_SZ, DefaultValue
380                  End If
390              End If
400          Case REG_BINARY
410              length = 4096
420              ReDim resBinary(length - 1) As Byte
430              If RegQueryValueEx(handle, ValueName, 0, REG_BINARY, resBinary(0), length) = 0 Then
440                  ReDim Preserve resBinary(length - 1) As Byte
450                  GetRegistryValue = resBinary()
460               Else
470                  If Create Then
480                     SetRegistryValue hKey, KeyName, ValueName, REG_BINARY, DefaultValue
490                  End If
500              End If
510          Case Else
520              Err.Raise 1001, , "Tipo de valor no soportado"
530      End Select

540      RegCloseKey handle

550      Exit Function

GestErr:
560      Me.MousePointer = vbNormal
570      MsgBox "[GetRegistryValue]" & vbCrLf & Err.Description & Erl
End Function

