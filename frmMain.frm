VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg3N.ocx"
Begin VB.Form frmMain 
   Caption         =   "KESDP"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   3930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin CCRProgressBar6.ccrpProgressBar Bar 
      Height          =   135
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   238
      FillColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   50
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gözat"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Parçala ve Yok Et"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":17002
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Alt 
      Alignment       =   2  'Center
      Caption         =   "Bu Yazýlým KESDP 1.00 Yöntemi Sayesinde Bilgilerinizi Elegeçirlenez, Yeniden Okunamaz Bir Þekilde GüvenleSiler..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3705
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo SON
Bar.Value = 0
Bar.Value = Bar.Value + 1
Bar.Visible = True
Command1.Caption = "Dosya Siliniyor..."
Command1.Enabled = False

Dosya_No = FreeFile
Dim Ekle_BiLgi As String
Ekle_BiLgi = "/ÿÿgfdgfdgfdgdgfdgdfsggggggggggggggggggggggoýpoýpggg*-1gfdg234567daspofðd890*-,qwertyuýodsfspðüsdfghjklgdfgfdgffdsfdsffdpoýpdsgfdgfdgdsfgsdfgfdþixcvbnmöç.>£#$[{8][\}]\sfgdgdgdsgfdsg13541€´¨;" & CLng(Rnd * 100000 + 1)
Open Text1.Text For Binary As Dosya_No
Put #Dosya_No, 23, Ekle_BiLgi
Bar.Value = Bar.Value + 1
Randomize
Ekle_BiLgi = "ÿÿ32323(/&(&/dgfdgdgfdgdfs(ggggggggggggdfdsd890*-,qwertyu(/&ýodsfspðüsdfgfdggfdgffdsf(/&(dsffdpoýpdsgfdgfdgdsfgsdfgfdþixcvbngfg>£#$[{8][\}]\sfgdgdgds741ðððððð´¨;" & CLng(Rnd * 100000 + 1)
Put #Dosya_No, 223, Ekle_BiLgi
Bar.Value = Bar.Value + 1
Put #Dosya_No, 223, Ekle_BiLgi
Bar.Value = Bar.Value + 1
Randomize
Ekle_BiLgi = "32323(/&(&/dgfdgÿÿdgfdgdfs(ggggggggggggdfdfdsfdüsdfg454dgffdsf(/&(dsffd32432432gdsfgsdfgfdþixcvbngfg>£#$[{8][\}]\sfgdg43241***99" & CLng(Rnd * 100000 + 1)
Put #Dosya_No, 2233, Ekle_BiLgi
Bar.Value = Bar.Value + 1
Put #Dosya_No, 243, Ekle_BiLgi

Put #Dosya_No, , Ekle_BiLgi
Bar.Value = Bar.Value + 1
Randomize
Ekle_BiLgi = "32323(/&(&/dgfdgdgfdgdfs(ggggggggggggdfdsd890*-,qwertyu(/&ýodsfspðüÿÿsdfgfdggfdgffdsf(/&(dsffdpoýpdsgfdgfdgdsfgs*-/545121lkjlj¨;" & CLng(Rnd * 100000 + 1)
Put #Dosya_No, LOF(Dosya_No) / 2, Ekle_BiLgi

Randomize
Ekle_BiLgi = "3fdsfdssfffffffÿÿfffffff452432432ðððððððððððððððððððððüüüüüüüüüüüüüüüüüüüüüüeeeeeeeeeeeeeeÿÿ;" & CLng(Rnd * 100000 + 1)
Put #Dosya_No, LOF(Dosya_No) / 3, Ekle_BiLgi
Bar.Value = Bar.Value + 1
Randomize
Ekle_BiLgi = "3fu97ssfffýffffffff452432432ðððððððððððÿÿpðpðððüüüüüüüüüüüüüüüüüüüoüeeee69eeeeeeeeee;" & CLng(Rnd * 100000 + 1)
Put #Dosya_No, LOF(Dosya_No) / 7, Ekle_BiLgi
Bar.Value = Bar.Value + 1
Randomize
Ekle_BiLgi = "3fu97ssfffýffffffff45243243432432423ððÿÿpðpðððüüüüüüüü212121üeeee69ee21eee;" & CLng(Rnd * 100000 + 1)
Put #Dosya_No, LOF(Dosya_No) / 5, Ekle_BiLgi

Randomize
Ekle_BiLgi = "3fu97ss12fff45243213243242432ÿÿpðpðððüü121üüüüü212121üe2121eee;" & CLng(Rnd * 100000 + 1)
Put #Dosya_No, LOF(Dosya_No) / 11, Ekle_BiLgi
Bar.Value = Bar.Value + 1
Randomize
Ekle_BiLgi = "vbaOnError    _adj_fdiv_m16i    __vbaObjSetAddref   _adj_fdivr_m16i   __vbaStrFixstr    __vbaBoolVar    __vbaBoolVarNull    _CIsin    __vbaVargVarMove    __vbaChkstk   __vbaFileClose    EVENT_SINK_AddRef   __vbaStrCmp   __vbaVarTstEq   __vbaI2I4   DllFunctionCall   __vbaVarOr    _adj_fpatan   __vbaR4Var    __vbaFixstrConstruct    __vbaLateIdCallLd   __vbaRecUniToAnsi   EVENT_SINK_Release" & CLng(Rnd * 100000 + 1)
Put #Dosya_No, LOF(Dosya_No) / 6, Ekle_BiLgi
Bar.Value = Bar.Value + 1
Randomize
Ekle_BiLgi = "KEKSDP BU DOSYAYI BOZMUÞ VE PARÇALAMIÞTIR.GEÇMÝÞ OLSUN@À   º ´Í!¸LÍ!This program cannot be run in DOS mode."
Put #Dosya_No, 3, Ekle_BiLgi
Bar.Value = Bar.Value + 1
'10 CUT

Close #Dosya_No
FileCopy Text1.Text, Text1.Text & "$$$"
Bar.Value = Bar.Value + 15
'1 COP
Kill Text1.Text
Bar.Value = Bar.Value + 5
'1 DEL
FileCopy Text1.Text & "$$$", Text1.Text
Bar.Value = Bar.Value + 15
'2 COP
Kill Text1.Text & "$$$"
Bar.Value = Bar.Value + 5
'2DEL
Name Text1.Text As Pth(Text1.Text) & "AAA.AAA.AAA"
Bar.Value = Bar.Value + 3
Name Pth(Text1.Text) & "AAA.AAA.AAA" As Pth(Text1.Text) & "CCC.CCC.CCC"
Bar.Value = Bar.Value + 3
Name Pth(Text1.Text) & "CCC.CCC.CCC" As Pth(Text1.Text) & "DDD.DDD.DDD"
Bar.Value = Bar.Value + 3
Name Pth(Text1.Text) & "DDD.DDD.DDD" As Pth(Text1.Text) & "XXX.XXX.XXX"
Bar.Value = Bar.Value + 3
Name Pth(Text1.Text) & "XXX.XXX.XXX" As Pth(Text1.Text) & "QQQ.QQQ.QQQ"
Bar.Value = Bar.Value + 3
Name Pth(Text1.Text) & "QQQ.QQQ.QQQ" As Pth(Text1.Text) & "WWW.WWW.WWW"
Bar.Value = Bar.Value + 3
Name Pth(Text1.Text) & "WWW.WWW.WWW" As Pth(Text1.Text) & "LLL.LLL.LLL"
Bar.Value = Bar.Value + 3
'7NAME
FileCopy Pth(Text1.Text) & "LLL.LLL.LLL", Pth(Text1.Text) & "$ZZ.ZZZ"
Bar.Value = Bar.Value + 15
'3COP
Kill Pth(Text1.Text) & "LLL.LLL.LLL"
Bar.Value = Bar.Value + 5
'3DEL
Name Pth(Text1.Text) & "$ZZ.ZZZ" As Pth(Text1.Text) & "$KE.SDK"
Bar.Value = Bar.Value + 3
'8NAME
Kill Pth(Text1.Text) & "$KE.SDK"
Bar.Value = Bar.Value + 5
'4 DEL
Close #Dosya_No
Bar.Visible = False
Command1.Caption = "Parçala ve Yok Et."
Command1.Enabled = True
Alt.Caption = StripPath(Text1.Text) & " baþarýyla silindi!"
SON:
         If Err.Number <> 0 Then
         Bar.Visible = False
         Command1.Caption = "Parçala ve Yok Et."
         Command1.Enabled = True
         Alt.Caption = "Bir Hata Oluþtu... Lütfen Dosya Konumunu Kontrol Ediniz. HATA NO: " & Err.Number
         End If
End Sub
Private Function StripPath(ref3$) As String

         Dim ref1%, ref2%

         StripPath$ = ref3$

         ref1% = InStr(ref3$, "\")

         Do While ref1%

                  ref2% = ref1%

                  ref1% = InStr(ref2% + 1, ref3$, "\")

         Loop

         If ref2% > 0 Then StripPath$ = Mid$(ref3$, ref2% + 1)

End Function
Private Function Pth(x As String)
Dim y As String
y = Len(StripPath(x))
y = Len(x) - y
Pth = Left(x, y)
End Function

Private Sub Command2_Click()
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
End Sub
