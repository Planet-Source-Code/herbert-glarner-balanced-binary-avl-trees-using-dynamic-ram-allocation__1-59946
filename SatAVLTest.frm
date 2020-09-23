VERSION 5.00
Begin VB.Form SatAVLTest 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "SatAVL Test and comparison with Collection"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   525
      Left            =   90
      TabIndex        =   39
      Top             =   2100
      Width           =   5955
      Begin VB.Label lblStatus 
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test results"
      Height          =   3075
      Left            =   90
      TabIndex        =   12
      Top             =   2730
      Width           =   5925
      Begin VB.Label AVLDestroy 
         Caption         =   "0"
         Height          =   255
         Left            =   3810
         TabIndex        =   38
         Top             =   2760
         Width           =   1665
      End
      Begin VB.Label CollDestroy 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   37
         Top             =   2760
         Width           =   1665
      End
      Begin VB.Label Label28 
         Caption         =   "Destroy"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label AVLSeqData 
         Caption         =   "0"
         Height          =   255
         Left            =   3810
         TabIndex        =   35
         Top             =   2430
         Width           =   1665
      End
      Begin VB.Label Label26 
         Caption         =   "(omitted, too slow)"
         Height          =   255
         Left            =   1800
         TabIndex        =   34
         Top             =   2430
         Width           =   1665
      End
      Begin VB.Label Label25 
         Caption         =   "Sequential, Data"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label AVLSeqKey 
         Caption         =   "0"
         Height          =   255
         Left            =   3810
         TabIndex        =   32
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label Label23 
         Caption         =   "n.a."
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label Label22 
         Caption         =   "Sequential, Key"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Label AVLDown 
         Caption         =   "0"
         Height          =   255
         Left            =   3810
         TabIndex        =   29
         Top             =   1770
         Width           =   1665
      End
      Begin VB.Label Label20 
         Caption         =   "n.a."
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   1770
         Width           =   1665
      End
      Begin VB.Label Label19 
         Caption         =   "Highest/Lower"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   1770
         Width           =   1575
      End
      Begin VB.Label AVLUp 
         Caption         =   "0"
         Height          =   255
         Left            =   3810
         TabIndex        =   26
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label17 
         Caption         =   "n.a."
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label16 
         Caption         =   "Lowest/Higher"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label AVLDirDesc 
         Caption         =   "0"
         Height          =   255
         Left            =   3810
         TabIndex        =   23
         Top             =   1110
         Width           =   1665
      End
      Begin VB.Label CollDirDesc 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   1110
         Width           =   1665
      End
      Begin VB.Label Label13 
         Caption         =   "Direct descending"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label AVLDirAsc 
         Caption         =   "0"
         Height          =   255
         Left            =   3810
         TabIndex        =   20
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label CollDirAsc 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label Label10 
         Caption         =   "Direct ascending"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label AVLAdd 
         Caption         =   "0"
         Height          =   255
         Left            =   3810
         TabIndex        =   17
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label CollAdd 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label Label7 
         Caption         =   "Add"
         Height          =   165
         Left            =   180
         TabIndex        =   15
         Top             =   510
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "SatAVL"
         Height          =   225
         Left            =   3810
         TabIndex        =   14
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Collection"
         Height          =   225
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mass test settings"
      Height          =   1935
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5925
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   285
         Left            =   4710
         TabIndex        =   11
         Top             =   1470
         Width           =   1035
      End
      Begin VB.CheckBox chkDataLen 
         Caption         =   "Random 1-2048"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   1470
         Width           =   1545
      End
      Begin VB.TextBox txtDataLen 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1380
         TabIndex        =   9
         Text            =   "32"
         Top             =   1440
         Width           =   1005
      End
      Begin VB.CheckBox chkKeySorted 
         Caption         =   "Sorted ascending (unchecked=random)"
         Height          =   285
         Left            =   1380
         TabIndex        =   7
         Top             =   690
         Width           =   3255
      End
      Begin VB.CheckBox chkKeyLen 
         Caption         =   "Random 8-24"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txtKeyLen 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1380
         TabIndex        =   4
         Text            =   "8"
         Top             =   1050
         Width           =   1005
      End
      Begin VB.TextBox txtNum 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1380
         TabIndex        =   2
         Text            =   "100000"
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Data length"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "Key order"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Key length"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Elements"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   1125
      End
   End
End
Attribute VB_Name = "SatAVLTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oColl As Collection
Dim oSatAVL As SatAVL

Private Declare Sub QueryPerformanceCounter _
    Lib "kernel32" (lpPerformanceCount As Currency)
Private Declare Sub QueryPerformanceFrequency _
    Lib "kernel32" (lpFrequency As Currency)


Private Sub chkDataLen_Click()
    If chkDataLen.Value Then
        txtDataLen.Enabled = False
    Else
        txtDataLen.Enabled = True
    End If
End Sub

Private Sub chkKeyLen_Click()
    If chkKeyLen.Value Then
        txtKeyLen.Enabled = False
    Else
        txtKeyLen.Enabled = True
    End If
End Sub

Private Sub chkKeySorted_Click()
    If chkKeySorted.Value Then
        txtKeyLen.Enabled = False
        chkKeyLen.Enabled = False
    Else
        txtKeyLen.Enabled = True
        chkKeyLen.Enabled = True
    End If
End Sub

Private Sub cmdStart_Click()
    Const Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim lNum As Long
    Dim lCnt As Long
    Dim sAllKeys() As String
    Dim lK1 As Long, lK2 As Long, lK3 As Long, lK4 As Long, lK5 As Long
    Dim lKeyLen As Long
    Dim sKey As String
    Dim lKeyPos As Long
    Dim lDataLen As Long
    Dim bRndDataLen As Boolean
    Dim timStart As Double
    Dim timEnd As Double
    Dim timDuration As Double
    Dim lAddr As Long
    Dim sData As String
    Dim lIndex As Long
    
    'Get data from form
    lNum = CLng(txtNum.Text)
    If lNum > 2000000 Then
        'In VB, we can't allocate arrays with more than 2 mill elements
        MsgBox "Array size max 2 million under VB"
        Exit Sub
    End If
    ReDim sAllKeys(1 To lNum) As String
    
    cmdStart.Enabled = False
    cmdStart.Refresh
    
    'First the key generation overhead (outside of time measurement): Same keys
    'for Collection and SatAVL.
    lblStatus = "Generating and storing keys..."
    lblStatus.Refresh
    If chkKeySorted Then
        For lK1 = 1 To 26: For lK2 = 1 To 26: For lK3 = 1 To 26: For lK4 = 1 To 26: For lK5 = 1 To 26
            sKey = Mid$(Alpha, lK1, 1) & Mid$(Alpha, lK2, 1) & Mid$(Alpha, lK3, 1) & Mid$(Alpha, lK4, 1) & Mid$(Alpha, lK5, 1)
            lCnt = lCnt + 1: sAllKeys(lCnt) = sKey
            If lCnt = lNum Then GoTo KeysGenerated
        Next: Next: Next: Next: Next
KeysGenerated:
    Else
        lKeyLen = CLng(txtKeyLen.Text)
        For lCnt = 1 To lNum
            If chkKeyLen.Value Then lKeyLen = Int(Rnd() * 17) + 8
            'Generate random key
            sKey = Space$(lKeyLen)
            For lKeyPos = 1 To lKeyLen
                Mid$(sKey, lKeyPos, 1) = Mid$(Alpha, Int(Rnd() * 26 + 1), 1)
            Next lKeyPos
            'Storing the generated key to have something to retrieve
            sAllKeys(lCnt) = sKey
        Next lCnt
    End If
    
    bRndDataLen = (chkDataLen.Value = 1)
    lDataLen = CLng(txtDataLen.Text)
    'ALL DATA READY FOR TESTS (except data length creation when random)
    
    
    
    'C O L L E C T I O N
    Set oColl = New Collection
    
    'Add
    lblStatus = "Collection.Add..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        'Calculate random data length if wanted
        If bRndDataLen Then lDataLen = Rnd() * 2048& + 1&
        
        'Add the item as per key array, reserving data space as required
        oColl.Add Space$(lDataLen), sAllKeys(lCnt)
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    CollAdd.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    CollAdd.Refresh
    
    'Item (direct ascending)
    lblStatus = "Collection.Item (Direct Ascending)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        sData = oColl.Item(sAllKeys(lCnt))
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    CollDirAsc.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    CollDirAsc.Refresh
    
    'Item (direct descending)
    lblStatus = "Collection.Item (Direct Descending)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = lNum To 1 Step -1
        sData = oColl.Item(sAllKeys(lCnt))
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    CollDirDesc.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    CollDirDesc.Refresh
    
    'Destroy
    lblStatus = "Coll.Destroy..."
    lblStatus.Refresh
    timStart = QPTimer
    Set oColl = Nothing
    timEnd = QPTimer
    timDuration = timEnd - timStart
    CollDestroy.Caption = Format(timDuration, "0.000") & "s "
    CollDestroy.Refresh
    
    
    
    
    'S A T A V L
    Set oSatAVL = New SatAVL
    
    'Add
    lblStatus = "SatAVL.Add..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        'Calculate random data length if wanted
        If bRndDataLen Then lDataLen = Rnd() * 2048& + 1&
        
        'Add the item as per key array, reserving data space as required
        lAddr = oSatAVL.Add(sAllKeys(lCnt), lDataLen)
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLAdd.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLAdd.Refresh
    
    'Item (direct ascending)
    lblStatus = "SatAVL.Item (Direct Ascending)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 1 To lNum
        lIndex = oSatAVL.Item(sAllKeys(lCnt))
        If lIndex = -1 Then Stop
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDirAsc.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLDirAsc.Refresh
    
    'Item (direct descending)
    lblStatus = "SatAVL.Item (Direct Descending)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = lNum To 1 Step -1
        lIndex = oSatAVL.Item(sAllKeys(lCnt))
        If lIndex = -1 Then Stop
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDirDesc.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLDirDesc.Refresh
    
    'Lowest/higher
    lblStatus = "SatAVL.Sorted (Lowest/Higher)..."
    lblStatus.Refresh
    timStart = QPTimer
    lIndex = oSatAVL.Lowest
    Do While lIndex <> -1&
        lIndex = oSatAVL.Higher
    Loop
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLUp.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLUp.Refresh
    
    'Highest/Lower
    lblStatus = "SatAVL.Sorted (Highest/Lower)..."
    lblStatus.Refresh
    timStart = QPTimer
    lIndex = oSatAVL.Highest
    Do While lIndex <> -1&
        lIndex = oSatAVL.Lower
    Loop
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDown.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLDown.Refresh
    
    'Sequential (Key)
    lblStatus = "SatAVL.Sequential (Key)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 0 To lNum - 1
        sKey = oSatAVL.Key(lCnt)
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLSeqKey.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLSeqKey.Refresh
    
    'Sequential (Data)
    lblStatus = "SatAVL.Sequential (Data)..."
    lblStatus.Refresh
    timStart = QPTimer
    For lCnt = 0 To lNum - 1
        lAddr = oSatAVL.Address(lCnt)
    Next lCnt
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLSeqData.Caption = Format(timDuration, "0.000") & "s (" & Format$(lNum / timDuration, "0") & "/s)"
    AVLSeqData.Refresh
    
    'Destroy
    lblStatus = "SatAVL.Destroy..."
    lblStatus.Refresh
    timStart = QPTimer
    Set oSatAVL = Nothing
    timEnd = QPTimer
    timDuration = timEnd - timStart
    AVLDestroy.Caption = Format(timDuration, "0.000") & "s "
    AVLDestroy.Refresh
    
    lblStatus = "Ready"
    cmdStart.Enabled = True
End Sub




Private Sub Form_Unload(Cancel As Integer)
    'Make sure we destroy the objects
    Set oColl = Nothing
    Set oSatAVL = Nothing
End Sub


'Source for Timer: http://vb-tec.de/timer.htm (German)
Public Function QPTimer() As Double
  Static Takt As Currency
  Dim Dauer As Currency
  
  If Takt = 0 Then
    'einmal die Taktfrequenz bestimmen:
    QueryPerformanceFrequency Takt
  End If
  
  'aktuelle Zeit holen:
  QueryPerformanceCounter Dauer
  
  'Zeit in Sekunden umrechnen:
  QPTimer = Dauer / Takt
End Function




