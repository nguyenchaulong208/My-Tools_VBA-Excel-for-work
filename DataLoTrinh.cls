VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataLoTrinh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private bienSoXe As String
Private ngayBatDau As Date
Private ngayKetThuc As Date
Private tenSheet As String

Public Property Let bienSoXe_(ByRef newBienSoXe As String)
    bienSoXe = newBienSoXe
End Property
Public Property Get bienSoXe_() As String
    bienSoXe_ = bienSoXe
End Property


Public Property Let ngayBatDau_(newNgayBatDau As Date)
    ngayBatDau = newNgayBatDau
End Property
Public Property Get ngayBatDau_() As Date
    ngayBatDau_ = ngayBatDau
End Property

Public Property Let ngayKetThuc_(newNgayKetThuc As Date)
    ngayKetThuc = newNgayKetThuc
End Property
Public Property Get ngayKetThuc_() As Date
    ngayKetThuc_ = ngayKetThuc
End Property

Public Property Let tenSheet_(newTenSheet As String)
    tenSheet = newTenSheet
End Property
Public Property Get tenSheet_() As String
    tenSheet_ = tenSheet
End Property
