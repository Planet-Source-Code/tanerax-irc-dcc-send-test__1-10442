VERSION 5.00
Begin VB.Form FrmMirc 
   Caption         =   "Mirc DCC Test"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Return Code"
      Height          =   2535
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      Begin VB.ListBox LstReturn 
         Height          =   2205
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "To: "
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3615
      Begin VB.TextBox TxtFileTo 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "From: "
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.TextBox TxtFileFrom 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
End
Attribute VB_Name = "FrmMirc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Programming By: Tanerax [Tanerax@nbnet.nb.ca] [50298342]
'// Program: DCC Sending/Recieving Test v1.0
'// Comments: Complete
'// This Program is a Test Version For DCC Sending For Irc
'// This Takes The Files As Its Receiving It And Creates The
'// Proper Strings To Be Sent Back To The Sender
'// Is 4 Bit Network Order
'// ********************************************************
'// You May Notice That This Does Not Actually Use The Internet
'// This Is More Of Test Showing How It Can Be Done
'// Instead Of Reading From The File You Would Read The Packets
'// From The Sender. And Instead Of Adding The Return Code To A
'// ListBox Send Them To The Sender
'// ********************************************************
'// If You Use This At All In A Program Please Mention Me
'// If You Can Optimize This Please Send Me A Optimized Version
'// Via E-Mail
'// Thank You.

Sub CmdTest(Infile As String, OutFile As String)
    Dim hIn, fileLength, ret                       '// Declare Variables
    Dim temp As String                             '// Declare Variables
    Dim BlockSize As Long                          '// Declare Variables
    BlockSize = 1024                               '// Set Your Read Buffer Size here
       
    Open Infile For Binary Access Read As #1       '// Open File For Reading
    Open OutFile For Binary Access Write As #2     '// Close File For Reading
   
    fileLength = LOF(1)                            '// Declare FileLength as the
                                                   '// The File Length
    Do Until EOF(1)                                '// Begin Do..Loop Stop At EOF
        If fileLength - Loc(1) <= BlockSize Then   '// If The Block Size is Larger
                                                   '// Than the EOF Make Block Size
            BlockSize = fileLength - Loc(1)        '// Small Enough To Accomadate
        
        End If
         
        If BlockSize = 0 Then Exit Do              '// If Block Size Is Zero Then
                                                   '// Finished
        bytesent = bytesent + BlockSize            '// Add The Blocksize To the Bytes Sent
        temp = Space$(BlockSize)                   '// Allocate The Read Buffer
        Get 1, , temp                              '// Read a block of data
        DoEvents                                   '// Stop Program Until All Is Read
        ret = DoEvents()                           '// Check for cancel button event etc.
                
'//**************************************************************************
        Put 2, , temp                                '// Write Information To Other File
        retval = Hex(LOF(2))                         '// Take File Size And Change To Hex
        If Len(retval) = 3 Then                      '// Check To Make Sure The Numeber To
            retval = "0" & retval                    '// To Be Sent Is In 4 Bytes
        End If
         
        If Len(retval) > 4 Then                      '// This Allows For the Network
         sFirst = Right$(retval, 4)                  '// Order To Come In Takes The 8
         ssecond = Mid$(retval, 1, Len(retval) - 4)  '// Bytes of Characters and Changes
            If Len(ssecond) = 1 Then                 '// Them Into Sets Of 2. 2 Sets Is all
                ssecond = "000" + ssecond            '// You Should Need For IRC. 2 Sets Of
                                                     '// FFFF FFFF Would Transfer Files Up To
            ElseIf Len(ssecond) = 2 Then             '// 4,294,967,295 Bytes = Appx 4.1 Gigs
                ssecond = "00" + ssecond             '// More Than Enough For Irc Sending Of
            ElseIf Len(ssecond) = 3 Then             '// Single Files
                ssecond = "00" + ssecond
            End If
         retval = ssecond & "     " & sFirst
         End If
            
         LstReturn.AddItem retval                    '// Adds The Values To The List
                     
         sizeOfFileSent = sizeOfFileSent + BlockSize '// Adds The Block Size To Total
                                                     '// Size Sent
         Loop                                        '// End Loop
   
    Close #1                                         '// Closes hIn
    Close #2                                         '// Closes hOut
   
   End Sub
Private Sub CmdCopy_Click()
    CmdTest TxtFileFrom.Text, TxtFileTo.Text         '// Call the Sub Routine
End Sub

Private Sub CmdExit_Click()
    End                                              '//End Program
End Sub
