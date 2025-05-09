VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThongTinLoTrinh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private ngay As String
Private diaDiem As String
Private thoiGianBd As String
Private thoiGianKt As String
Private lamThemGio As Integer
Private soKmBd As Integer
Private soKmKt As Integer
Private quangDuong As Long
Private tongTienVe As Long
Private soLuongVe As Integer
Private bienSoXe As String
Private taiXe As String
Private khachHang As String
Private tuyenDuong As String

'------
'Constructor
' --- Ngay ---
Public Property Let ngay_(ByRef newNgay As String)
    ngay = newNgay
End Property

Public Property Get ngay_() As String
    ngay_ = ngay
End Property

' --- DiaDiem ---
Public Property Let diaDiem_(ByRef newDiaDiem As String)
    diaDiem = newDiaDiem
End Property

Public Property Get diaDiem_() As String
    diaDiem_ = diaDiem
End Property

' --- ThoiGianBd ---
Public Property Let thoiGianBd_(ByRef newTGBD As String)
    thoiGianBd = newTGBD
End Property

Public Property Get thoiGianBd_() As String
    thoiGianBd_ = thoiGianBd
End Property

' --- ThoiGianKt ---
Public Property Let thoiGianKt_(ByRef newTGKT As String)
    thoiGianKt = newTGKT
End Property

Public Property Get thoiGianKt_() As String
    thoiGianKt_ = thoiGianKt
End Property

' --- SoKmBd ---
Public Property Let soKmBd_(ByVal newKmBd As Long)
    soKmBd = newKmBd
End Property

Public Property Get soKmBd_() As Long
    soKmBd_ = soKmBd
End Property

' --- SoKmKt ---
Public Property Let soKmKt_(ByVal newKmKt As Long)
    soKmKt = newKmKt
End Property

Public Property Get soKmKt_() As Long
    soKmKt_ = soKmKt
End Property

' --- QuangDuong ---
Public Property Let quangDuong_(ByVal newQD As Long)
    quangDuong = newQD
End Property

Public Property Get quangDuong_() As Long
    quangDuong_ = quangDuong
End Property

' --- TongTienVe ---
Public Property Let tongTienVe_(ByVal newTien As Long)
    tongTienVe = newTien
End Property

Public Property Get tongTienVe_() As Long
    tongTienVe_ = tongTienVe
End Property

' --- SoLuongVe ---
Public Property Let soLuongVe_(ByVal newSLV As Integer)
    soLuongVe = newSLV
End Property

Public Property Get soLuongVe_() As Integer
    soLuongVe_ = soLuongVe
End Property

' --- BienSoXe ---
Public Property Let bienSoXe_(ByRef newBSX As String)
    bienSoXe = newBSX
End Property

Public Property Get bienSoXe_() As String
    bienSoXe_ = bienSoXe
End Property

' --- TaiXe ---
Public Property Let taiXe_(ByRef newTaiXe As String)
    taiXe = newTaiXe
End Property

Public Property Get taiXe_() As String
    taiXe_ = taiXe
End Property

' --- KhachHang ---
Public Property Let khachHang_(ByRef newKH As String)
    khachHang = newKH
End Property

Public Property Get khachHang_() As String
    khachHang_ = khachHang
End Property

' --- TuyenDuong ---
Public Property Let tuyenDuong_(ByRef newTD As String)
    tuyenDuong = newTD
End Property

Public Property Get tuyenDuong_() As String
    tuyenDuong_ = tuyenDuong
End Property

Public Property Let lamThemGio_(ByRef newLamThemGio As Integer)
    lamThemGio = newLamThemGio
End Property
Public Property Get lamThemGio_() As Integer
    lamThemGio_ = lamThemGio
End Property


