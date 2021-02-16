Attribute VB_Name = "Selekcja_zdjec"
Sub podzielenie_zdjec()
    Dim polygon(4, 1) As Double
    Dim input_fullpath_folder As String
    Dim output_fullpath_folder As String
    Dim pozostale As Boolean
    Dim photo_coord(0, 1) As Double
    Dim licznik_w_zakresie As Integer
    Dim licznik_poza_zakresem As Integer
    
    Dim MyFile As Object
    Dim MyFSO As Object
    Dim MyFolder As Object
    Set MyFSO = CreateObject("Scripting.FileSystemObject")
    
    If Not czy_zaczynac Then End ' kontrolne pytanie "czy zaczynamy?"
    
'''...........Pobranie danych z Excela........................
    ''' wczytanie wspó³rzednych naro¿y
    polygon(0, 0) = Range("C3").Value 'p1.x
    polygon(0, 1) = Range("C4").Value 'p1.y
    polygon(1, 0) = Range("I3").Value 'p2.x
    polygon(1, 1) = Range("I4").Value 'p2.y
    polygon(2, 0) = Range("I16").Value 'p3.x
    polygon(2, 1) = Range("I17").Value 'p3.y
    polygon(3, 0) = Range("C16").Value 'p4.x
    polygon(3, 1) = Range("C17").Value 'p4.y
    polygon(4, 0) = Range("C3").Value 'p1.x
    polygon(4, 1) = Range("C4").Value 'p2.x
    
    ''' wczytanie œcie¿ki
    input_fullpath_folder = Range("E20").Value
    output_fullpath_folder = Range("E21").Value

    ''' czy kopiowaæ zdjêcia spoaz obszaru do osobnego folderu
    If Range("E22").Value = "Tak" Then
        pozostale = True
    Else
        pozostale = False
    End If
    
    licznik_w_zakresie = 0
    licznik_poza_zakresem = 0
    If Not if_folder_exist(input_fullpath_folder) Then End ' kontrola czy folder z danymi istnieje
    If Not if_folder_exist(output_fullpath_folder) Then End ' kontrola czy folder z danymi istnieje
    If incorrect_data(polygon) Then End ' kontrola poprawnoœci danych

'''...........................................................
    
    '''----------utworzenie podfolderów wynikowych----------------
    On Error Resume Next
    MyFSO.CreateFolder (output_fullpath_folder & "\1. W obszarze")
    If pozostale = True Then
        MyFSO.CreateFolder (output_fullpath_folder & "\2. Poza obszarem")
    End If
    '''------------------------------------------------------------
    
    Set MyFolder = MyFSO.GetFolder(input_fullpath_folder)

    For Each MyFile In MyFolder.Files
        With CreateObject("WIA.ImageFile")
            .LoadFile MyFile
            '''---------pobranie wspó³rzêdnych zdjêcia---------
            With .Properties("GpsLatitude").Value
                photo_coord(0, 0) = .Item(1).Value + .Item(2).Value / 60 + .Item(3).Value / 3600 '''Latitude
            End With
            With .Properties("GpsLongitude").Value
                photo_coord(0, 1) = .Item(1).Value + .Item(2).Value / 60 + .Item(3).Value / 3600 '''Longitude
            End With
            '''-------------------------------------------------
            
            '''------kopiowani pod warunkiem ¿e w obszarze------
            If czy_w_zakresie(polygon, photo_coord) Then
                Call MyFSO.CopyFile(input_fullpath_folder & "\" & MyFile.Name, output_fullpath_folder & "\1. W obszarze\" & MyFile.Name, False)
                licznik_w_zakresie = licznik_w_zakresie + 1
            Else
                If pozostale = True Then       '''kopiowanie zdjêæ poza obszarem
                    Call MyFSO.CopyFile(input_fullpath_folder & "\" & MyFile.Name, output_fullpath_folder & "\2. Poza obszarem\" & MyFile.Name, False)
                End If
                licznik_poza_zakresem = licznik_poza_zakresem + 1
            End If
            '''-------------------------------------------------
        End With
    Next
    
    Call raport(licznik_w_zakresie, licznik_poza_zakresem, pozostale)
    
End Sub

Private Function czy_w_zakresie(polygon As Variant, photo_coord() As Double) As Boolean

    Dim aa As Double
    Dim bb As Double
    Dim cc As Double
    Dim dd As Double
    Dim test1 As Boolean
    Dim test2 As Boolean
    Dim i As Integer
    
    test1 = True
    test2 = True
   
    For i = 0 To 3
        aa = -(polygon(i + 1, 1) - polygon(i, 1))
        bb = polygon(i + 1, 0) - polygon(i, 0)
        cc = -(aa * polygon(i, 0) + bb * polygon(i, 1))
        dd = aa * photo_coord(0, 0) + bb * photo_coord(0, 1) + cc
        test1 = test1 And (dd >= 0)
        test2 = test2 And (dd <= 0)
    Next i
    
    czy_w_zakresie = test1 Or test2
   
End Function

Private Function czy_zaczynac() As Boolean
    Dim str_response As String
    
    str_response = MsgBox("Czy zaczynaæ? Mo¿e to chwilê potrwaæ..", vbQuestion + vbYesNo)
    If str_response = vbYes Then
        czy_zaczynac = True
    Else
        MsgBox "Przewano na ¿yczenie u¿ytkownika", , "Raport"
        czy_zaczynac = False
    End If
    
End Function

Private Function incorrect_data(polygon As Variant) As Boolean
    
        If polygon(0, 1) > polygon(1, 1) Or _
            polygon(3, 1) > polygon(2, 1) Or _
            polygon(3, 0) > polygon(0, 0) Or _
            polygon(2, 0) > polygon(1, 0) Then
        incorrect_data = True
        End If
    
End Function

Private Sub raport(licznik_w_zakresie As Integer, licznik_poza_zakresem As Integer, pozostale As Boolean)
    
    If pozostale = False Then
        MsgBox "Skopiowano " & licznik_w_zakresie & " zdjêæ bêd¹cych w zadanym obszarze z " & licznik_w_zakresie + licznik_poza_zakresem & " przeszukiwanych zdjêæ.", , "Raport"
    Else
        MsgBox "Skopiowano " & licznik_w_zakresie & " zdjêæ bêd¹cych w zadanym obszarze oraz " _
        & licznik_poza_zakresem & " zdjêæ bêd¹cych poza zadanym obszarzem z " _
        & licznik_w_zakresie + licznik_poza_zakresem & " przeszukiwanych zdjêæ.", , "Raport"
    End If
    
End Sub

Private Function if_folder_exist(fullpath_folder As String) As Boolean

    If Dir(fullpath_folder, vbDirectory) = "" Then
        if_folder_exist = False '"The selected folder doesn't exist"
        MsgBox fullpath_folder & vbCrLf & "Folder o podanej œcie¿ce nie istnieje." & vbCrLf & "Podaj poprawny (istniej¹cy) folder" & vbCrLf & "Program zakoñczy dzia³anie.", vbCritical, "Error"
    Else
        if_folder_exist = True ' "The selected folder exists"
    End If

End Function

