VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "PNG Monsterous"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Checkzopflipng 
      Caption         =   "zopflipng"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox CheckPNGOut 
      Caption         =   "pngout"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox CheckAdvPNG 
      Caption         =   "advpng"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CheckBox Checkoptipngconsole 
      Caption         =   "optipngconsole"
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox CheckPNGCrush 
      Caption         =   "pngcrush"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox CheckPNGRewrite 
      Caption         =   "pngrewrite"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox LoopChk 
      Caption         =   "Loop (Very slow)"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Can get a few more bytes compressed but VERY slow!"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox LogTxt 
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmMain.frx":0BC2
      Top             =   1080
      Width           =   5415
   End
   Begin VB.CommandButton RunCmd 
      Caption         =   "Run!"
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox DirTxt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Options:"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Log:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Source Directory:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private InCount As Long
Private OutCount As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Log(ByVal LogText As String)

    LogTxt.Text = LogTxt.Text & vbNewLine & LogText

End Sub

Private Sub DirTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        RunCmd_Click
    End If

End Sub

Private Sub Form_Load()

    LogTxt.Text = vbNullString
    DirTxt.Text = App.Path & "\"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Me.Visible = False

End Sub

Private Sub LogTxt_Change()

    LogTxt.SelStart = Len(LogTxt.Text)

End Sub

Private Sub LogTxt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim s() As String
Dim j As Long

    InCount = 0
    OutCount = 0

    Log "You fed me " & Data.Files.Count & " files! *Munch munch munch!*"
    For i = 1 To Data.Files.Count
        If Not Me.Visible Then Exit For
        If ValidateImageSuffix(Data.Files(i)) Then
            Me.Caption = "PNG Monsterous - " & Round((i - 1) / Data.Files.Count, 2) * 100 & "%"
            s = Split(Data.Files(i), "\")
            Log "Digesting file " & Chr$(34) & s(UBound(s)) & Chr$(34)
            DoEvents
            OptimizePNG Data.Files(i)
            j = j + 1
        End If
    Next i
    If Not Me.Visible Then
        Unload Me
        End
    End If
    Log "All full! *burp*"
    Log vbNewLine & "I ate a total of " & j & " files... I'm still hungry. :("
    Log "Bytes before: " & InCount
    Log "Bytes after: " & OutCount
    Log "Compression ratio: " & Round(OutCount / InCount, 4) * 100 & "%"
    Me.Caption = "PNG Monsterous - 100%"

End Sub

Private Sub RunCmd_Click()
Dim s() As String
Dim Files() As String
Dim i As Long
Dim j As Long

    InCount = 0
    OutCount = 0
    LogTxt.Text = ""

    Files() = AllFilesInFolders(DirTxt.Text, True)
    Log "You fed me " & UBound(Files) + 1 & " files! *Munch munch munch!*"
    For i = 0 To UBound(Files)
        If Me.Visible = False Then Exit For
        If ValidateImageSuffix(Files(i)) Then
            Me.Caption = "PNG Monsterous - " & Round(i / (UBound(Files) + 1), 2) * 100 & "%"
            s = Split(Files(i), "\")
            Log "Digesting file " & Chr$(34) & s(UBound(s)) & Chr$(34)
            DoEvents
            OptimizePNG Files(i)
            j = j + 1
        End If
    Next i
    If Not Me.Visible Then
        Unload Me
        End
    End If
    Log "All full! *burp*"
    Log vbNewLine & "I ate a total of " & j & " files... I'm still hungry. :("
    Log "Bytes before: " & InCount
    Log "Bytes after: " & OutCount
    If Not InCount = 0 Then
        Log "Compression ratio: " & Round(OutCount / InCount, 4) * 100 & "%"
    End If
    Me.Caption = "PNG Monsterous - 100%"
    
End Sub

Private Function ValidateImageSuffix(ByVal File As String) As Boolean

    File = LCase$(File)
    If Right$(File, 4) = ".png" Then
        ValidateImageSuffix = True
    Else
        ValidateImageSuffix = False
    End If

End Function

Private Sub OptimizePNG(ByVal File As String)
Dim LastSize As Long
Dim CurrSize As Long
Dim FileNum As Byte

    FileNum = FreeFile

    Open File For Binary Access Read As #FileNum
        CurrSize = LOF(FileNum)
    Close #FileNum
    InCount = InCount + CurrSize

    If CheckPNGRewrite.Value > 0 Then
        LogTxt.Text = LogTxt.Text & vbNewLine & "Running pngrewrite.exe"
        CommandLine App.Path & "\pngrewrite.exe " & Chr$(34) & File & Chr$(34) & " " & Chr$(34) & File & Chr$(34)
    End If
    Do
        LastSize = CurrSize
        If CheckPNGCrush.Value > 0 Then
            LogTxt.Text = LogTxt.Text & vbNewLine & "Running pngcrush.exe"
            CommandLine App.Path & "\pngcrush.exe -rem gAMA -rem cHRM -rem iCCP -rem sRGB -brute -l 9 -max -reduce -m 0 -q " & Chr$(34) & File & Chr$(34) & " " & Chr$(34) & File & ".temp" & Chr$(34)
            If Dir$(File & ".temp") <> vbNullString Then
                Kill File
                Name File & ".temp" As File
            End If
        End If
        
        If Checkoptipngconsole.Value > 0 Then
            LogTxt.Text = LogTxt.Text & vbNewLine & "Running optipngconsole.exe"
            CommandLine App.Path & "\optipngconsole.exe -o7 -q " & Chr$(34) & File & Chr$(34)
        End If
        
        If CheckAdvPNG.Value > 0 Then
            LogTxt.Text = LogTxt.Text & vbNewLine & "Running advpng.exe"
            CommandLine App.Path & "\advpng.exe -z -4 " & Chr$(34) & File & Chr$(34)
        End If
        
        If CheckPNGOut.Value > 0 Then
            LogTxt.Text = LogTxt.Text & vbNewLine & "Running pngout.exe"
            CommandLine App.Path & "\pngout.exe /q /y /k0 /s0 " & Chr$(34) & File & Chr$(34) & " " & Chr$(34) & File & Chr$(34)
        End If
        
        If Checkzopflipng.Value > 0 Then
            LogTxt.Text = LogTxt.Text & vbNewLine & "Running zopflipng.exe"
            CommandLine App.Path & "\zopflipng.exe -y " & Chr$(34) & File & Chr$(34) & " " & Chr$(34) & File & Chr$(34)
        End If
        
        Open File For Binary Access Read As #FileNum
            CurrSize = LOF(FileNum)
        Close #FileNum
        
        If Not LoopChk.Value Then Exit Do
        
    Loop While CurrSize < LastSize
    
    OutCount = OutCount + CurrSize
    
End Sub

Private Sub CommandLine(ByVal CommandLineString As String)
Dim Start As STARTUPINFO
Dim Proc As PROCESS_INFORMATION

    Start.dwFlags = &H1
    Start.wShowWindow = 0
    CreateProcessA 0&, CommandLineString, 0&, 0&, False, &H20&, 0&, 0&, Start, Proc
    Do While WaitForSingleObject(Proc.hProcess, 0) = 258
        DoEvents
        Sleep 10
    Loop

End Sub
