Public Class CHelperFunctions

    Public Sub swapBytes(ByRef bytArray() As Byte, ByVal stByt As Integer)
        Dim tmp As Byte

        tmp = bytArray(stByt)
        bytArray(stByt) = bytArray(stByt + 1)
        bytArray(stByt + 1) = tmp
    End Sub

    'This function swaps a word (4 bytes)
    Public Sub swapWord(ByRef bytArray() As Byte, ByVal stByt As Integer)
        Dim tmp1 As Byte
        Dim tmp2 As Byte
        Dim tmp3 As Byte
        Dim tmp4 As Byte

        tmp1 = bytArray(stByt)
        tmp2 = bytArray(stByt + 1)
        tmp3 = bytArray(stByt + 2)
        tmp4 = bytArray(stByt + 3)
        bytArray(stByt) = tmp4
        bytArray(stByt + 1) = tmp3
        bytArray(stByt + 2) = tmp2
        bytArray(stByt + 3) = tmp1
    End Sub

    'This function swaps a word (8 bytes)
    Public Sub swap8(ByRef bytArray() As Byte, ByVal stByt As Integer)
        Dim tmp1 As Byte
        Dim tmp2 As Byte
        Dim tmp3 As Byte
        Dim tmp4 As Byte
        Dim tmp5 As Byte
        Dim tmp6 As Byte
        Dim tmp7 As Byte
        Dim tmp8 As Byte

        tmp1 = bytArray(stByt)
        tmp2 = bytArray(stByt + 1)
        tmp3 = bytArray(stByt + 2)
        tmp4 = bytArray(stByt + 3)
        tmp5 = bytArray(stByt + 4)
        tmp6 = bytArray(stByt + 5)
        tmp7 = bytArray(stByt + 6)
        tmp8 = bytArray(stByt + 7)
        bytArray(stByt) = tmp8
        bytArray(stByt + 1) = tmp7
        bytArray(stByt + 2) = tmp6
        bytArray(stByt + 3) = tmp5
        bytArray(stByt + 4) = tmp4
        bytArray(stByt + 5) = tmp3
        bytArray(stByt + 6) = tmp2
        bytArray(stByt + 7) = tmp1
    End Sub
End Class
