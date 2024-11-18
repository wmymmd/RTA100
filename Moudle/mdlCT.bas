Attribute VB_Name = "mdlCT"
Option Explicit


Public Declare Sub CopyMemory Lib "kernel32" _
      Alias "RtlMoveMemory" _
      (Destination As Any, Source As Any, ByVal Length As Long)


Public SendBuffer(100) As Byte
Public ReadData(255) As Byte
Public CommPort As Integer
Public Modbus_Addr As Integer
Public CmdData(10, 7) As Byte
Public CmdDataLength(10) As Integer
Public cmdID As Integer
Public r_baudrateID As Integer, r_stopbit As Integer, r_ratio As Integer, r_PT As Integer, r_CT As Integer
Public Volt(3) As Double, Current(3) As Double, Power(3) As Double
Public kvar(3) As Double, kVA(3) As Double, PF(3) As Double
Public kWh(3) As Double, kvarh(3) As Double, kVAh(3) As Double
Public CRCTable_Low(255) As Byte, CRCTable_High(255) As Byte
Public CRC_Low, CRC_High As Byte
Public Net_ID As Byte





' Initial CRC table
Public Sub InitCRCTable()
    Dim i As Integer
    Dim CRC_HI
    Dim CRC_LO
    
    CRC_HI = Array( _
        &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
        &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
        &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
        &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
        &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
        &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
        &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
        &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
        &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
        &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
        &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
        &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
        &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
        &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
        &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
        &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40)
    CRC_LO = Array( _
        &H0, &HC0, &HC1, &H1, &HC3, &H3, &H2, &HC2, &HC6, &H6, &H7, &HC7, &H5, &HC5, &HC4, &H4, _
        &HCC, &HC, &HD, &HCD, &HF, &HCF, &HCE, &HE, &HA, &HCA, &HCB, &HB, &HC9, &H9, &H8, &HC8, _
        &HD8, &H18, &H19, &HD9, &H1B, &HDB, &HDA, &H1A, &H1E, &HDE, &HDF, &H1F, &HDD, &H1D, &H1C, &HDC, _
        &H14, &HD4, &HD5, &H15, &HD7, &H17, &H16, &HD6, &HD2, &H12, &H13, &HD3, &H11, &HD1, &HD0, &H10, _
        &HF0, &H30, &H31, &HF1, &H33, &HF3, &HF2, &H32, &H36, &HF6, &HF7, &H37, &HF5, &H35, &H34, &HF4, _
        &H3C, &HFC, &HFD, &H3D, &HFF, &H3F, &H3E, &HFE, &HFA, &H3A, &H3B, &HFB, &H39, &HF9, &HF8, &H38, _
        &H28, &HE8, &HE9, &H29, &HEB, &H2B, &H2A, &HEA, &HEE, &H2E, &H2F, &HEF, &H2D, &HED, &HEC, &H2C, _
        &HE4, &H24, &H25, &HE5, &H27, &HE7, &HE6, &H26, &H22, &HE2, &HE3, &H23, &HE1, &H21, &H20, &HE0, _
        &HA0, &H60, &H61, &HA1, &H63, &HA3, &HA2, &H62, &H66, &HA6, &HA7, &H67, &HA5, &H65, &H64, &HA4, _
        &H6C, &HAC, &HAD, &H6D, &HAF, &H6F, &H6E, &HAE, &HAA, &H6A, &H6B, &HAB, &H69, &HA9, &HA8, &H68, _
        &H78, &HB8, &HB9, &H79, &HBB, &H7B, &H7A, &HBA, &HBE, &H7E, &H7F, &HBF, &H7D, &HBD, &HBC, &H7C, _
        &HB4, &H74, &H75, &HB5, &H77, &HB7, &HB6, &H76, &H72, &HB2, &HB3, &H73, &HB1, &H71, &H70, &HB0, _
        &H50, &H90, &H91, &H51, &H93, &H53, &H52, &H92, &H96, &H56, &H57, &H97, &H55, &H95, &H94, &H54, _
        &H9C, &H5C, &H5D, &H9D, &H5F, &H9F, &H9E, &H5E, &H5A, &H9A, &H9B, &H5B, &H99, &H59, &H58, &H98, _
        &H88, &H48, &H49, &H89, &H4B, &H8B, &H8A, &H4A, &H4E, &H8E, &H8F, &H4F, &H8D, &H4D, &H4C, &H8C, _
        &H44, &H84, &H85, &H45, &H87, &H47, &H46, &H86, &H82, &H42, &H43, &H83, &H41, &H81, &H80, &H40)
        
    For i = 0 To 255
        CRCTable_Low(i) = CRC_LO(i)
        CRCTable_High(i) = CRC_HI(i)
    Next
    
End Sub

' calculate modbus CRC16
Public Sub Modbus_CRC16(iLen As Integer)
    Dim Hi_Byte As Byte
    Dim Lo_Byte  As Byte
    Dim idx As Byte
    Dim mStr As Byte
    Dim tStr As Byte
    Dim i As Integer
    
    Hi_Byte = &HFF
    Lo_Byte = &HFF
        
    For i = 0 To iLen - 1
       idx = Hi_Byte Xor SendBuffer(i)
       Hi_Byte = Lo_Byte Xor CRCTable_High(idx)
       Lo_Byte = CRCTable_Low(idx)
    Next i
   
   CRC_High = Hi_Byte
   CRC_Low = Lo_Byte

End Sub

Public Function Modbus_CRC(data() As Byte) As Byte()
    Dim Hi_Byte As Byte
    Dim Lo_Byte  As Byte
    Dim idx As Byte
    Dim i As Integer
    Dim crcBytes(1 To 2) As Byte
    
    Hi_Byte = &HFF
    Lo_Byte = &HFF
        
    For i = LBound(data) To UBound(data)
       idx = Hi_Byte Xor data(i)
       Hi_Byte = Lo_Byte Xor CRCTable_High(idx)
       Lo_Byte = CRCTable_Low(idx)
    Next i
   
   crcBytes(1) = Hi_Byte
   crcBytes(2) = Lo_Byte
   Modbus_CRC = crcBytes
End Function
