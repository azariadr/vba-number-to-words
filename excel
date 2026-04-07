Function TERBILANG(ByVal angka As Double) As String
    Dim satuan As Variant
    Dim unitPos As Variant
    Dim hasil As String
    Dim negatif As Boolean
    Dim unitIndex As Integer
    Dim chunk As Integer
    Dim tempAngka As Double

    ' Penanganan angka nol
    If angka = 0 Then
        TERBILANG = "NOL RUPIAH"
        Exit Function
    End If

    ' Penanganan angka negatif
    If angka < 0 Then
        negatif = True
        angka = Abs(angka)
    End If

    satuan = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan")
    unitPos = Array("", "Ribu", "Juta", "Miliar", "Triliun")
    
    tempAngka = Int(angka)
    hasil = ""
    unitIndex = 0

    Do While tempAngka > 0
        chunk = tempAngka Mod 1000
        If chunk > 0 Then
            Dim chunkWords As String
            chunkWords = ConvertChunk(chunk, satuan)
            
            ' Logika khusus untuk "Seribu"
            If chunk = 1 And unitIndex = 1 Then
                hasil = "Seribu " & hasil
            Else
                If unitIndex > 0 Then
                    hasil = chunkWords & " " & unitPos(unitIndex) & " " & hasil
                Else
                    hasil = chunkWords & " " & hasil
                End If
            End If
        End If
        
        tempAngka = Int(tempAngka / 1000)
        unitIndex = unitIndex + 1
    Loop

    If negatif Then hasil = "Minus " & hasil
    
    TERBILANG = UCase(Trim(hasil)) & " RUPIAH"
End Function

Private Function ConvertChunk(ByVal num As Integer, satuan As Variant) As String
    Dim ratus As Integer
    Dim sisaPuluhan As Integer
    Dim tempStr As String
    Dim puluhan As Variant
    Dim belasan As Variant
    
    belasan = Array("Sepuluh", "Sebelas", "Dua Belas", "Tiga Belas", "Empat Belas", "Lima Belas", "Enam Belas", "Tujuh Belas", "Delapan Belas", "Sembilan Belas")
    puluhan = Array("", "", "Dua Puluh", "Tiga Puluh", "Empat Puluh", "Lima Puluh", "Enam Puluh", "Tujuh Puluh", "Delapan Puluh", "Sembilan Puluh")
    
    ratus = Int(num / 100)
    sisaPuluhan = num Mod 100
    
    ' Ratusan
    If ratus > 0 Then
        If ratus = 1 Then
            tempStr = "Seratus "
        Else
            tempStr = satuan(ratus) & " Ratus "
        End If
    End If
    
    ' Puluhan dan Satuan
    If sisaPuluhan > 0 Then
        If sisaPuluhan >= 10 And sisaPuluhan <= 19 Then
            tempStr = tempStr & belasan(sisaPuluhan - 10)
        Else
            tempStr = tempStr & puluhan(Int(sisaPuluhan / 10))
            If (sisaPuluhan Mod 10) > 0 Then
                If puluhan(Int(sisaPuluhan / 10)) <> "" Then tempStr = tempStr & " "
                tempStr = tempStr & satuan(sisaPuluhan Mod 10)
            End If
        End If
    End If
    
    ConvertChunk = Trim(tempStr)
End Function
