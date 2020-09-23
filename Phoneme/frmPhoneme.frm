VERSION 5.00
Begin VB.Form frmPhoneme 
   Caption         =   "Phoneme Demo"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSoundex 
      Caption         =   "Compute Soundex Word"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Soundex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   4455
      Begin VB.TextBox txtInputSoundex 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Soundex"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblSoundex 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "String"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdLD 
      Caption         =   "Compute Levenshtein Distance"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Levenshtein Distance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtStr2 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtStr1 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Levenshtein Distance"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblLD 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "String 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   615
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "String 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.Label lblSoundexInfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   15
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblLDInfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4680
      TabIndex        =   14
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmPhoneme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// Returns the 4 letter soundex for an english word
Private Function Soundex(ByVal argWord As String)
Dim workStr As String, i As Long, replaceMask(5) As Boolean

    '// Capitalize it to remove ambiguity
    argWord = UCase$(argWord)
    
    '// 1. Retain the first letter of the string
    workStr = Left$(argWord, 1)
    
    '// 2. Replacement
    '   [a, e, h, i, o, u, w, y] = 0
    '   [b, f, p, v] = 1
    '   [c, g, j, k, q, s, x, z] = 2
    '   [d, t] = 3
    '   [l] = 4
    '   [m, n] = 5
    '   [r] = 6
    
    For i = 2 To Len(argWord)
        Select Case Mid$(argWord, i, 1)
            Case "B", "F", "P", "V"
                If replaceMask(0) = False Then
                    workStr = workStr & Chr$(49) '// 1
                    replaceMask(0) = True
                End If
                
            Case "C", "G", "J", "K", "Q", "S", "X", "Z"
                If replaceMask(1) = False Then
                    workStr = workStr & Chr$(50) '// 2
                    replaceMask(1) = True
                End If
            
            Case "D", "T"
                If replaceMask(2) = False Then
                    workStr = workStr & Chr$(51) '// 3
                    replaceMask(2) = True
                End If
            
            Case "L"
                If replaceMask(3) = False Then
                    workStr = workStr & Chr$(52) '// 4
                    replaceMask(3) = True
                End If
            
            Case "M", "N"
                If replaceMask(4) = False Then
                    workStr = workStr & Chr$(53) '// 5
                    replaceMask(4) = True
                End If
                
            Case "R"
                If replaceMask(5) = False Then
                    workStr = workStr & Chr$(56) '// 6
                    replaceMask(5) = True
                End If
            
            '// A, E, H, I, O, U, W, Y do nothing
        End Select
    Next i
    
    '// 5. Return the first four bytes padded with 0.
    If Len(workStr) > 4 Then
        Soundex = Left$(workStr, 4)
    Else
        Soundex = workStr & Space$(4 - Len(workStr))
    End If
End Function

Public Function min3(ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long) As Long
    min3 = n1
    If n2 < min3 Then min3 = n2
    If n3 < min3 Then min3 = n3
End Function

Private Function LevenshteinDistance(ByVal argStr1 As String, ByVal argStr2 As String) As Long
Dim m As Long, n As Long
Dim editMatrix() As Long, i As Long, j As Long, cost As Long
Dim str1_i As String, str2_j As String
Dim p() As Long, q() As Long, r As Long
Dim x As Long, y As Long

    n = Len(argStr1)
    m = Len(argStr2)
    
    If (n = 0) Or (m = 0) Then Exit Function
    ReDim editMatrix(n, m) As Long
    
    
    For i = 0 To n
        editMatrix(i, 0) = i
    Next
    
    For j = 0 To m
        editMatrix(0, j) = j
    Next
    
    For i = 1 To n
        str1_i = Mid$(argStr1, i, 1)
        For j = 1 To m
            str2_j = Mid$(argStr2, j, 1)
            If str1_i = str2_j Then
                cost = 0
            Else
                cost = 1
            End If
            
            editMatrix(i, j) = min3(editMatrix(i - 1, j) + 1, editMatrix(i, j - 1) + 1, editMatrix(i - 1, j - 1) + cost)
        Next j
    Next i
            
    LevenshteinDistance = editMatrix(n, m)
    Erase editMatrix
End Function

Private Sub cmdLD_Click()
    lblLD.Caption = LevenshteinDistance(txtStr1.Text, txtStr2.Text)
End Sub

Private Sub cmdSoundex_Click()
    If txtInputSoundex.Text <> vbNullString Then lblSoundex.Caption = Soundex(txtInputSoundex.Text)
End Sub

Private Sub Form_Load()
    lblLDInfo.Caption = "    From Wikipedia, the free encyclopedia. " & _
                        "In information theory and computer science, " & _
                        "the Levenshtein distance or edit distance between two " & _
                        "strings is given by the minimum number of operations needed " & _
                        "to transform one string into the other, where an operation " & _
                        "is an insertion, deletion, or substitution of a single character. " & _
                        "It is named after Vladimir Levenshtein, " & _
                        "who considered this distance in 1965. " & _
                        "It is useful in applications that need to " & _
                        "determine how similar two strings are, such as spell checkers."
                        
    lblSoundexInfo.Caption = "    Soundex is a phonetic algorithm for indexing " & _
                             "names by their sound when pronounced in English. " & _
                             "The basic aim is for names with the same pronunciation " & _
                             "to be encoded to the same string so that matching " & _
                             "can occur despite minor differences in spelling. " & _
                             "Soundex is the most widely known of all phonetic " & _
                             "algorithms and is often used (incorrectly) as a " & _
                             "synonym for phonetic algorithm."
                             
End Sub
