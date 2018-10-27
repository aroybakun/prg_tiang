VERSION 5.00
Begin VB.Form AAAA_LOGIN 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "LOGIN"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "AAAA_LOGIN.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "tombol update hpp tiang"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   4920
      Picture         =   "AAAA_LOGIN.frx":4464
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtpassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2790
      PasswordChar    =   "X"
      TabIndex        =   2
      Top             =   2050
      Width           =   2055
   End
   Begin VB.ComboBox txtbagian 
      Height          =   315
      ItemData        =   "AAAA_LOGIN.frx":4AAD
      Left            =   2790
      List            =   "AAAA_LOGIN.frx":4AC0
      TabIndex        =   1
      Top             =   1680
      Width           =   2025
   End
   Begin VB.TextBox txtnama 
      Height          =   285
      Left            =   2790
      TabIndex        =   0
      Top             =   1400
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lbl_download 
      BackColor       =   &H00FFFFFF&
      Caption         =   "https://www.dropbox.com/s/e7em0rcms91wtr6/TiangPancang.exe"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alamat Download :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   5880
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "   x"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "AAAA_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo buat_koneksi_Error
Dim rp1 As New ADODB.Recordset
Dim str_rp1, str_versi As String
'str_versi = "15.08.00"
'str_versi = "14.04.10"
'str_versi = "15.09.01"
'str_versi = "15.09.21"
'str_versi = "15.09.25"
'str_versi = "15.10.10.1"
'str_versi = "16.02.16"
str_versi = "16.09.14"
str_versi = "16.03.24"

Set rt = New ADODB.Recordset
Set rp1 = New ADODB.Recordset
rp1.CursorLocation = adUseClient
'str_rp1 = "CALL liat_versi"
str_rp1 = "SELECT * FROM versi_prg WHERE aktif='1' ORDER BY log_date DESC"
rp1.Open str_rp1, conn, 3, 1, 1

rt.CursorLocation = adUseClient
rt.Open "select nama,bagian,password from login where nama='" & txtnama.Text & "'and bagian='" & txtbagian.Text & "'and password='" & txtpassword.Text & "'", conn, adOpenStatic, adLockOptimistic, adCmdText
If rp1.RecordCount = 0 Then
Else
  If rp1!versi = str_versi Then
    If rt.BOF And rt.EOF Then
       MsgBox "silakan coba lagi", vbOKOnly + vbInformation, "PALU MAS SEJATI"
    Else
          If txtbagian = "ADMINISTRATOR" Then
             master_administrator.Show
             master_administrator.lblNama = txtnama.Text
             master_administrator.lblBagian = txtbagian.Text
             master_administrator.lbl_versi = "Prg. Versi : " & str_versi
             Unload Me
          ElseIf txtbagian = "PRODUKSI" Then
             layer_produksi.Show
             layer_produksi.lbl_versi = "Prg. Versi : " & str_versi
             layer_produksi!Label1.Caption = txtnama.Text
             Unload Me
          ElseIf txtbagian = "ADMINISTRASI" Then
             layer_administrasi.Show
             layer_administrasi.lbl_versi = "Prg. Versi : " & str_versi
             layer_administrasi!Label1.Caption = txtnama.Text
             Unload Me
          ElseIf txtbagian = "GUDANG" Then
             layer_gudang.Show
             layer_gudang.lbl_versi = "Prg. Versi : " & str_versi
             Unload Me
          Else
buat_koneksi_Error:                  MsgBox "Koneksi Ok, tetapi Ada kesalahan login, periksa apakah server sudah berjalan !", vbInformation, "Cek Server, Jaringan dan Client"
          End If
    End If
  Else
     MsgBox "maaf !!! program sudah ada yang baru, harap mendownload lagi", vbOKOnly, "Pms Group"
     Label2.Visible = True
     lbl_download.Visible = True
      lbl_download.Caption = "https://www.dropbox.com/s/e7em0rcms91wtr6/TiangPancang.exe"
  End If
End If
'rt.Close
End Sub

Private Sub Command2_Click()
koneksi
Dim rp1 As New ADODB.Recordset
Dim rp2 As New ADODB.Recordset
Dim hpp As New ADODB.Recordset
Dim str_rp1, str_rp2, str_ext, str_rt, str_hpp As String

Set rp1 = New ADODB.Recordset
rp1.CursorLocation = adUseClient

Set rp2 = New ADODB.Recordset
rp2.CursorLocation = adUseClient

Set hpp = New ADODB.Recordset
hpp.CursorLocation = adUseClient

str_rp2 = "SELECT tgl_1,tgl_2 FROM range_update_hpp_endang"
rp2.Open str_rp2, conn, 3, 1, 1
Dim tgl1, tgl2 As String
'tgl1 = rp2!tgl_1
'tgl2 = rp2!tgl_2
'cari data yang mw diupdate harganya
str_rp1 = "SELECT * FROM laporan_history WHERE tgl_transaksi>='" & Format(rp2!tgl_1, "yyyy-MM-dd") & "' AND tgl_transaksi<='" & Format(rp2!tgl_2, "yyyy-MM-dd") & "' AND LENGTH(kode_tiang)>8 AND jenis_transaksi='produksi'"
rp1.Open str_rp1, conn, 3, 1, 1
If rp1.RecordCount = 0 Then

Else
    Do While Not rp1.EOF
          Set rt = New ADODB.Recordset
            'Set rz = New ADODB.Recordset
            rt.CursorLocation = adUseClient
            'rz.CursorLocation = adUseClient
          'str_rt = "select no_transaksi,tgl_transaksi,nama_pemborong,kode_tiang,nama_tiang,jumlah,jenis_transaksi,operator,log_date,hpp from laporan_history where no_transaksi='" & rp1!no_transaksi & "' and tgl_transaksi='" & rp1!TGL_PRODUKSI & "' and nama_pemborong='" & rp1!nama_pemborong & "' and kode_tiang='" & rp1!kode_tiang & "'and jenis_transaksi='PRODUKSI' and jumlah='" & rp1!jumlah & "'"
          'str_rz = "select no_produksi,tgl_produksi,nama_pemborong,kode_tiang,nama_tiang,jumlah from master_tiang_muda where no_produksi='" & rs!no_produksi & "' and tgl_produksi='" & rs!TGL_PRODUKSI & "' and nama_pemborong='" & rs!NAMA_PEMBORONG & "' and kode_tiang='" & rs!kode_tiang & "' and jumlah='" & rs!jumlah & "'"
          'rt.Open str_rt, conn, adOpenStatic, adLockOptimistic, adCmdText
          'rz.Open str_rz, conn, adOpenStatic, adLockOptimistic, adCmdText
          
          'cari harga total from harga_pengiriman.`harga_bayar_tiang`.`harga_total`
          str_hpp = "SELECT harga_total FROM harga_pengiriman.harga_bayar_tiang "
            str_hpp = str_hpp & " WHERE tanggal_awal<='" & Format(rp1!tgl_transaksi, "yyyy-mm-dd") & "' AND tanggal_akhir>='" & Format(rp1!tgl_transaksi, "yyyy-mm-dd") & "' AND kode_tiang='" & rp1!kode_tiang & "'"
          If hpp.State = 1 Then
            hpp.Close
          End If
          
          hpp.Open str_hpp, conn, 3, 1, 1
          str_ext = "update laporan_history set hpp='" & hpp!harga_total & "' where no_transaksi='" & rp1!no_transaksi & "' and tgl_transaksi='" & Format(rp1!tgl_transaksi, "yyyy-mm-dd") & "' and nama_pemborong='" & rp1!NAMA_PEMBORONG & "' and kode_tiang='" & rp1!kode_tiang & "'"
          conn.Execute str_ext
          hpp.Close
        rp1.MoveNext
    Loop
End If
End Sub

Private Sub Form_Load()
setformattanggalawal
PeriksaTanggal
koneksi
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Do Until Me.Top <= -5000
    DoEvents
    Me.Move Me.Left, Me.Top - 0.6
    DoEvents
Loop
Unload Me
'unloadss (fLOGIN)
End Sub
Private Sub Label1_Click()
End
End Sub

Private Sub Label3_Click()
    Command2_Click
End Sub

Private Sub lbl_download_Click()
OpenURL "https://dl.dropboxusercontent.com/u/91238766/KERJA/Program%20PMS/PROGRAM%20TIANG/TiangPancang.exe", Me.hwnd
End Sub

Private Sub txtbagian_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
If KeyAscii = 13 Then
    txtpassword.SetFocus
End If
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    txtbagian.SetFocus
End If
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub
Sub PeriksaTanggal()
Dim CekTanggal As String
Ulangi:
  CekTanggal = Date  'Tampung tanggal dalam bentuk string
  'Lakukan pemeriksaan format tanggal tersebut...
  If CekTanggal <> Format(Date, "MM/dd/yyyy") Then
     If MsgBox("Format tanggal di komputer Anda tidak sama dengan" & vbCrLf & _
           "'MM/dd/yyyy'. Klik OK untuk mengganti melalui menu" & vbCrLf & _
           "Regional Settings pada tab Date di kotak isian" & vbCrLf & _
           "'Short Date Style'. Ganti menjadi format:" & vbCrLf & _
           "mm/dd/yyyy. Jika Anda tidak melakukannya, maka" & vbCrLf & _
           "program tidak dapat dijalankan!", _
           vbCritical + vbOKCancel, _
           "Format Tanggal Tidak Sama Dengan 'MM/dd/yyyy'") _
           = vbOK And CekTanggal <> Format(Date, "MM/dd/yyyy") Then
           
           Call Shell("rundll32.exe shell32.dll," & _
                   "Control_RunDLL INTL.CPL,,4", 1)
           MsgBox "Silakan rubah date sesuai ketentuan, jika telah selesai, klik ok", vbOKOnly + vbInformation, "penting"
                 If vbOK Then
                    End
                 Else
                    End
                 End If
           
        'Tampilkan Regional Settings dari program, dan
        'langsung ke tab Date (Tab indeks ke-4)...
        
     Else
        End
        Unload Me 'Jika tdk mau mengganti, langsung keluar program...
     End If
     
  End If
End Sub


