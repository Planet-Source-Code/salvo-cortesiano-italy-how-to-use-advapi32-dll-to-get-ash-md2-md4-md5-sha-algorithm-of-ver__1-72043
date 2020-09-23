VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ash-Sha1 MD5 Encode Big File (Demo)"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Ash"
      Height          =   645
      Left            =   90
      TabIndex        =   17
      Top             =   3060
      Width           =   4890
      Begin VB.TextBox txtHash 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   240
         Width           =   4725
      End
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   195
      Left            =   825
      TabIndex        =   15
      Top             =   3840
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdStartConvert 
      Caption         =   "Get ASH"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5100
      TabIndex        =   14
      Top             =   3240
      Width           =   1170
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   270
      Left            =   5835
      TabIndex        =   13
      Top             =   2670
      Width           =   465
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2640
      Width           =   4620
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   1605
      TabIndex        =   8
      Top             =   720
      Width           =   4680
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   30
         ScaleHeight     =   1635
         ScaleWidth      =   4605
         TabIndex        =   9
         Top             =   135
         Width           =   4605
         Begin MSComDlg.CommonDialog CDialog 
            Left            =   3645
            Top             =   165
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtInfo 
            Height          =   1275
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   15
            Width           =   4530
         End
         Begin VB.Label lblSize 
            Caption         =   "n/a"
            Height          =   270
            Left            =   15
            TabIndex        =   19
            Top             =   1350
            Width           =   4470
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1605
         Left            =   30
         ScaleHeight     =   1605
         ScaleWidth      =   1395
         TabIndex        =   2
         Top             =   165
         Width           =   1395
         Begin VB.OptionButton OptEncodeMode 
            Caption         =   "SHA1"
            Height          =   270
            Index           =   3
            Left            =   90
            TabIndex        =   6
            Top             =   1230
            Width           =   1200
         End
         Begin VB.OptionButton OptEncodeMode 
            Caption         =   "MD5"
            Height          =   270
            Index           =   2
            Left            =   75
            TabIndex        =   5
            Top             =   915
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptEncodeMode 
            Caption         =   "MD4"
            Height          =   270
            Index           =   1
            Left            =   75
            TabIndex        =   4
            Top             =   585
            Width           =   1215
         End
         Begin VB.OptionButton OptEncodeMode 
            Caption         =   "MD2"
            Height          =   270
            Index           =   0
            Left            =   75
            TabIndex        =   3
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Encode Mode:"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   60
            TabIndex        =   7
            Top             =   15
            Width           =   1380
         End
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "© 2009 by Salvo Cortesiano. All Right Reserved!"
      Height          =   255
      Left            =   660
      TabIndex        =   20
      Top             =   4140
      Width           =   5700
   End
   Begin VB.Label Label4 
      Caption         =   "Read:"
      Height          =   210
      Left            =   105
      TabIndex        =   16
      Top             =   3825
      Width           =   690
   End
   Begin VB.Label Label3 
      Caption         =   "FileName:"
      Height          =   225
      Left            =   135
      TabIndex        =   11
      Top             =   2685
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "This Project use the (advapi32.dll) to determine the ASH of a big File -> 32Kb, and Get the MD2-MD4-MD5 and SHA Algorithm!"
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   6150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' .... Init Controls XP-Vista           ;) file res modifyed
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_USEREX_CLASSES = &H200

' .... AdvancedAPI
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long

' .... Constant
Private Const PROV_RSA_FULL = 1
Private Const CRYPT_NEWKEYSET = &H8
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA1 = 4

' .... Enum
Enum HashAlgorithm
   MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
   MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
   MD5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
   SHA1 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
End Enum

' .... Oter Constant
Private Const HP_HASHVAL = 2
Private Const HP_HASHSIZE = 4
Private Sub cmdOpen_Click()
    On Local Error Resume Next
    With CDialog
        .CancelError = True
        .DialogTitle = "Select the File:"
        .Filter = "All File(s) (*.*)|*.*"
        .DefaultExt = "*.*"
        .ShowOpen
        If .Filename = Empty Then
                cmdStartConvert.Enabled = False
            Exit Sub
        Else
            txtFileName.Text = .Filename
            cmdStartConvert.Enabled = True
        End If
        
        
        
    End With
End Sub

Private Sub cmdStartConvert_Click()
    If OptEncodeMode(0).Value Then
       txtHash.Text = HashFile(txtFileName.Text, MD2)
    ElseIf OptEncodeMode(1).Value Then
        txtHash.Text = HashFile(txtFileName.Text, MD4)
    ElseIf OptEncodeMode(2).Value Then
        txtHash.Text = HashFile(txtFileName.Text, MD5)
    ElseIf OptEncodeMode(3).Value Then
        txtHash.Text = HashFile(txtFileName.Text, SHA1)
    End If
End Sub

Private Sub Form_Initialize()
    Call InitControlsCtx
End Sub

Private Sub InitControlsCtx()
 On Local Error GoTo Init_Error
      Dim iccex As tagInitCommonControlsEx
      With iccex
          .lngSize = LenB(iccex)
          .lngICC = ICC_USEREX_CLASSES
      End With
      InitCommonControlsEx iccex
Exit Sub
Init_Error:
    Err.Clear
End Sub

Private Function HashFile(ByVal Filename As String, Optional ByVal Algorithm As HashAlgorithm = MD5) As String
    Dim hCtx As Long
    Dim hHash As Long
    Dim lFile As Long
    Dim lRes As Long
    Dim lLen As Long
    Dim lIdx As Long
    Dim abHash() As Byte
    
    txtInfo.Text = Empty
    
    ' .... Check if the file exists
    If Len(Dir$(Filename)) = 0 Then Err.Raise 53
   
    ' .... Get default provider context handle
    lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)
    If lRes = 0 And Err.LastDllError = &H80090016 Then
   
        ' .... There's no default keyset container
        ' .... Get the provider context and create a default keyset container
        lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    End If
    If lRes <> 0 Then

      ' .... Create the hash
      lRes = CryptCreateHash(hCtx, Algorithm, 0, 0, hHash)
      If lRes <> 0 Then

         ' .... Get a file handle
         lFile = FreeFile
         
         ' .... Open the file
         Open Filename For Binary As lFile
         
         If Err.Number = 0 Then
            
            ' .... Init Const Block Size = 32x32 Kb ;)
            Const BLOCK_SIZE As Long = 32 * 1024& ' 32K
            ReDim abBlock(1 To BLOCK_SIZE) As Byte
            Dim lCount As Long
            Dim lBlocks As Long
            Dim lLastBlock As Long
            
            ' .... Calculate how many full blocks the file contains
            lBlocks = LOF(lFile) \ BLOCK_SIZE
            
            txtInfo = txtInfo & "Block Size: " & lBlocks & vbCrLf
            
            ' .... Calculate the remaining data length
            lLastBlock = LOF(lFile) - lBlocks * BLOCK_SIZE
            
            ' .... Calculate total Size
            lblSize.Caption = "File Size: " & FormatSize(LOF(lFile))
            
            txtInfo = txtInfo & "Block Remaining: " & lLastBlock & vbCrLf
            
            PB.Max = lBlocks
            
            ' .... Hash the blocks
            For lCount = 1 To lBlocks
               Get lFile, , abBlock
         
               ' .... Add the chunk to the hash
               lRes = CryptHashData(hHash, abBlock(1), BLOCK_SIZE, 0)
            
               ' .... Stop the loop if CryptHashData fails
               If lRes = 0 Then Exit For
               DoEvents
               PB.Value = lCount
            Next

            ' .... Is there more data?
            If lLastBlock > 0 And lRes <> 0 Then
            
               ' .... Get the last block
               ReDim abBlock(1 To lLastBlock) As Byte
               Get lFile, , abBlock
               
               ' .... Hash the last block
               lRes = CryptHashData(hHash, abBlock(1), lLastBlock, 0)
            End If
            
            ' .... Close the file
            Close lFile
         End If

         If lRes <> 0 Then
            
            ' .... Get the hash lenght
            lRes = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, 4, 0)
            If lRes <> 0 Then

                ' .... Initialize the buffer
                ReDim abHash(0 To lLen - 1)

                ' .... Get the hash value
                lRes = CryptGetHashParam(hHash, HP_HASHVAL, abHash(0), lLen, 0)
                If lRes <> 0 Then
                    
                    PB.Max = UBound(abHash)
                    
                    ' .... Convert value to hex string
                    For lIdx = 0 To UBound(abHash)
                        HashFile = HashFile & Right$("0" & Hex$(abHash(lIdx)), 2)
                        DoEvents
                        
                        PB.Value = lIdx
                        txtInfo = txtInfo & "Hex String: " & Right$("0" & Hex$(abHash(lIdx)), 2) & vbCrLf
                    Next
                End If
            End If
         End If
         
         ' .... Release the hash handle
         CryptDestroyHash hHash
      End If
   End If

    ' .... Release the provider context
    CryptReleaseContext hCtx, 0
    
    PB.Value = 0
    
    txtInfo = txtInfo & "": txtInfo = txtInfo & "Finish!"
    
    ' .... Raise an error if lRes = 0
    If lRes = 0 Then MsgBox "Error: " & Err.LastDllError & vbCr & Err.Description, vbCritical, App.Title
    
End Function


Private Function FormatSize(size As Variant) As String
    On Local Error Resume Next
    If size >= 1073741824 And size <= 1099511627776# Then
            FormatSize = Format(((size / 1024) / 1024) / 1024, "#") & " GB"
        Exit Function
    End If


    If size >= 1048576 And size <= 1073741824 Then
            FormatSize = Format((size / 1024) / 1024, "#") & " MB"
        Exit Function
    End If

    If size >= 1024 And size <= 1048576 Then
            FormatSize = Format(size / 1024, "#") & " KB"
        Exit Function
    End If

    If size < 1024 Then
            FormatSize = size & " bytes"
        Exit Function
    End If
End Function
