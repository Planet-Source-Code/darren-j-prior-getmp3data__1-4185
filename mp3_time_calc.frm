VERSION 5.00
Begin VB.Form frmmp3_time_calc 
   Caption         =   "mp3timecalc"
   ClientHeight    =   8028
   ClientLeft      =   1584
   ClientTop       =   1392
   ClientWidth     =   4596
   LinkTopic       =   "Form1"
   ScaleHeight     =   8028
   ScaleWidth      =   4596
   Begin VB.FileListBox File1 
      Height          =   2568
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   4092
   End
   Begin VB.DirListBox Dir1 
      Height          =   1368
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   4092
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   372
      Left            =   2040
      TabIndex        =   2
      Top             =   6840
      Width           =   972
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   372
      Left            =   3360
      TabIndex        =   1
      Top             =   6840
      Width           =   972
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4092
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1572
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   4092
   End
End
Attribute VB_Name = "frmmp3_time_calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bitrate_lookup(7, 15) As Integer
Public actual_bitrate As Long
   Public Function Getmp3data(MP3File As String)
     Dim dIN As String
     cr = Chr(10)
     Open MP3File For Binary As #1
     ' read in 1st 4k of .mp3 file to find a frame header
     dIN = Input(4096, #1)
     filesize = LOF(1) ' needed to calculate track duration
     Close #1
     
     ' frame header starts with 12 set bits [sync]
     ' NB this ignores MPEG-2.5 which is 11 set bits, 1 zero bit.
     
     ' my search for the sync bits only works on nibble boundaries,
     ' I'm not sure if it is necessary to search on bit boundaries -
     ' if so then this search will be 4* slower and require a rewrite
     ' of this search section and shift_those_bits.
     Do Until i = 4095
       i = i + 1
       d1 = Asc(Mid(dIN, i, 1))
       d2 = Asc(Mid(dIN, i + 1, 1))
       If d1 = &HFF And (d2 And &HF0) = &HF0 Then
         'Debug.Print "Found at"; i
         ' get 20 hdr bits - they are last 20 bits of next 3 bytes
         temp_string = Mid(dIN, i + 1, 3)
         mp3bits_string = shift_those_bits(Mid(dIN, i + 1, 3))
         Exit Do
       End If
       ' if we haven't found the sync yet then shift left by 4 bits
       dSHIFT = shift_those_bits(Mid(dIN, i, 3))
       dd1 = Asc(Left(dSHIFT, 1))
       dd2 = Asc(Right(dSHIFT, 1))
       If dd1 = &HFF And (dd2 And &HF0) = &HF0 Then
         'Debug.Print "Found at"; i; "& a nibble"
         ' get 20 hdr bits - they are first 20 bits of next 3 bytes
         mp3bits_string = Mid(dIN, i + 2, 3)
         Exit Do
       End If
     Loop
     
     ' 1st 20 bits of mp3bits_string are hdr info for this frame
     ' 1st bit is ID - 0=MPG-2, 1=MPG-1
     mp3_id = (&H80 And Asc(Left(mp3bits_string, 1))) / 128
     ' next 2 bits are Layer
     mp3_layer = (&H60 And Asc(Left(mp3bits_string, 1))) / 32
     ' next bit is Protection
     mp3_prot = &H10 And Asc(Left(mp3bits_string, 1))
     ' next 4 bits are bitrate
     mp3_bitrate = &HF And Asc(Left(mp3bits_string, 1))
     'next 2 bits are frequency
     mp3_freq = &HC0 And Asc(Mid(mp3bits_string, 2, 1))
     ' next bit is Padding
     mp3_pad = (&H20 And Asc(Mid(mp3bits_string, 2, 1))) / 2
     actual_bitrate = 1000 * CLng((bitrate_lookup((mp3_id * 4) Or mp3_layer, mp3_bitrate)))
     
     dat = "ID: "
     If mp3_id = 0 Then
       dat = dat + "MPEG-2"
     Else
       dat = dat + "MPEG-1"
     End If
     
     dat = dat + cr + "Layer: "
      Select Case mp3_layer
        Case 1
          dat = dat + "Layer III"
        Case 2
          dat = dat + "Layer II"
        Case 3
          dat = dat + "Layer I"
      End Select
      dat = dat + cr + "Bitrate: " + Str(actual_bitrate)
      
      Select Case (mp3_id * 4) Or mp3_freq
        Case 0
          sample_rate = 22050
        Case 1
          sample_rate = 24000
        Case 2
          sample_rate = 16000
        Case 4
          sample_rate = 44100
        Case 5
          sample_rate = 48000
        Case 6
          sample_rate = 32000
      End Select
      dat = dat + cr + "Sample rate: " + Str(sample_rate)
      
      ' calculate track time
      framesize = ((144 * actual_bitrate) / sample_rate) + mp3_pad
      total_frames = filesize / framesize
      track_length = total_frames / 38.5 '38.5 frames per sec.
      
      dat = dat + cr + "Frames: " + Str(Int(total_frames))
      dat = dat + cr + "Duration: " + Str(Int(track_length)) + "secs"
      
      'display all the info
      lblInfo.Caption = dat
   End Function
   Public Function shift_those_bits(dIN As String) As String
     ' need to left shift 4 bits losing most significant 4 bits
     Dim sd1, sd2, sd3, do1, do2 As Integer
     duff = Left(dIN, 1)
     duff2 = Asc(duff)
     sd1 = Asc(Left(dIN, 1))
     sd2 = Asc(Mid(dIN, 2, 1))
     sd3 = Asc(Right(dIN, 1))
     
     do1 = ((sd1 And &HF) * 16) Or ((sd2 And &HF0) / 16)
     do2 = ((sd2 And &HF) * 16) Or ((sd3 And &HF0) / 16)
     shift_those_bits = Chr(do1) + Chr(do2)
   End Function


Private Sub cmdGo_Click()
  fname = File1.Path + "\" + File1.FileName
  If UCase(Right(fname, 4)) = ".MP3" Then
    Getmp3data (File1.Path + "\" + File1.FileName)
  Else
    lblInfo.Caption = "not a .mp3 file"
  End If
  End Sub

Private Sub cmdQuit_Click()
  End
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
  
  ' setup array for mpeg bitrate info
  bitrate_data = "032,032,032,032,008,008,"
  bitrate_data = bitrate_data + "064,048,040,048,016,016,"
  bitrate_data = bitrate_data + "096,056,048,056,024,024,"
  bitrate_data = bitrate_data + "128,064,056,064,032,032,"
  bitrate_data = bitrate_data + "160,080,064,080,040,040,"
  bitrate_data = bitrate_data + "192,096,080,096,048,048,"
  bitrate_data = bitrate_data + "224,112,096,112,056,056,"
  bitrate_data = bitrate_data + "256,128,112,128,064,064,"
  bitrate_data = bitrate_data + "288,160,128,144,080,080,"
  bitrate_data = bitrate_data + "320,192,160,160,096,096,"
  bitrate_data = bitrate_data + "352,224,192,176,112,112,"
  bitrate_data = bitrate_data + "384,256,224,192,128,128,"
  bitrate_data = bitrate_data + "416,320,256,224,144,144,"
  bitrate_data = bitrate_data + "448,384,320,256,160,160,"
    
  For y = 1 To 14
    For x = 7 To 5 Step -1
      bitrate_lookup(x, y) = Left(bitrate_data, 3)
      bitrate_data = Right(bitrate_data, Len(bitrate_data) - 4)
    Next
    For x = 3 To 1 Step -1
      bitrate_lookup(x, y) = Left(bitrate_data, 3)
      bitrate_data = Right(bitrate_data, Len(bitrate_data) - 4)
    Next
  Next
End Sub

