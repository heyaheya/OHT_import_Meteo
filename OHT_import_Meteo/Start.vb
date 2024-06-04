Imports System.IO
Imports System.Data.OracleClient
Imports DataTable = System.Data.DataTable
Imports System.Configuration

'Imports System.Data.OleDb
'Imports System.Data.Sql
'Imports System.Data.SqlClient

'************************************************************************************************************************************************
'
'Paramtery wywołania aplikacji
'   -v  działanie aplikacji w tle oraz podczyt danych z FTP po uruchomineniu aplikacji
'	-p1 / -p2   poziom logów
'		1 - poziom podst
'		2 - poziom serwisowy 
'
'************************************************************************************************************************************************

Module OHT

    'Public Class Form1
    Private Declare Auto Function ShowWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    Private Declare Auto Function GetConsoleWindow Lib "kernel32.dll" () As IntPtr
    Private Const SW_HIDE As Integer = 0

    Dim Tools As New OHT_import_Meteo.Tools_OHT

    Public katalog_pliku_log As String
    Public baza As String = "ERGH_TEST"  ' "ERGH_NEW"    'Dim baza As String = "ERGH"
    'Public conn As New OracleConnection(Tools.Get_ConnectionString(baza))
    Public conn As OracleConnection

    Public sciezka_zrodlo As String = "Y:\OHT\Dana Pomiarowe\"
    Public czy_zapisac_wszystkie_pliki As Boolean = False

    Public status_logu As Integer = 1
    '1 - poziom podst
    '2 - poziom serwisowy 
    '
    Public tryb_testowy As Boolean = False
    Public plik_do_importu As String = ""
    Public zrodlo_meteo As String = ""
    Public nazwa_pliku_log As String

    Sub Main()

        Dim StatusPodczytu As Boolean = False
        Dim arg_wartosc As String

        czy_zapisac_wszystkie_pliki = ConfigurationManager.AppSettings("czy_zapisac_wszystkie_pliki").ToString '.ToLower

        baza = ConfigurationManager.AppSettings("baza").ToString '.ToLower

        conn = New OracleConnection(Tools.Get_ConnectionString(baza))

        For Each arg As String In My.Application.CommandLineArgs
            If arg.Trim("-") = "v" Then
                StatusPodczytu = True
            End If

            If arg.Length > 2 Then

                If (arg.Substring(1, 2) = "p:") Then
                    If CInt(arg.Chars(3).ToString) < 3 Then
                        status_logu = CInt(arg.Substring(3, 1))
                    End If
                End If

                If (arg.Substring(1, 2).ToLower = "f:") Then
                    arg_wartosc = arg.Substring(3, arg.Length - 3)
                    plik_do_importu = arg_wartosc
                End If

                If (arg.Substring(1, 2).ToLower = "s:") Then
                    arg_wartosc = arg.Substring(3, arg.Length - 3)
                    zrodlo_meteo = arg_wartosc
                End If

                If (arg.Substring(1, 2).ToLower = "z:") Then
                    arg_wartosc = arg.Substring(3, arg.Length - 3)
                    czy_zapisac_wszystkie_pliki = arg_wartosc
                End If

            End If
        Next

        '/*
        '2021-12-05 14:14:45 Ilość argumentów wywołania aplikacji:5
        '2021-12-05 14:14:45 -v
        '2021-12-05 14:14:45 -p:1
        '2021-12-05 14:14:45 -f:y:\OHT\Dana Pomiarowe\Cumulus_3\20210818_2_Cumulus_prog_wiatr_sila.csv
        '2021-12-05 14:14:45 -s:Cumulus
        '2021-12-05 14:14:45 -z:true
        '*/


        sciezka_zrodlo = zrodlo_meteo & "_3\"

        Dim proces As Process
        proces = System.Diagnostics.Process.GetCurrentProcess

        If File.Exists(plik_do_importu) Then
            Dim plik = My.Computer.FileSystem.GetFileInfo(plik_do_importu)
            nazwa_pliku_log = plik.Name
        End If

        'nazwa_pliku_log = zrodlo_meteo & "_" & proces.Id.ToString & "_" & nazwa_pliku_log
        nazwa_pliku_log = zrodlo_meteo & "_" & nazwa_pliku_log

        DoLogu("Start")
        DoLogu(Tools.Check_internet)
        DoLogu("Ilość argumentów wywołania aplikacji:" & My.Application.CommandLineArgs.Count, 1)
        For Each arg As String In My.Application.CommandLineArgs
            DoLogu(arg.ToString)
        Next

        Const SW_HIDE As Integer = &H0

        If StatusPodczytu = True Then

            Dim hWndConsole As IntPtr
            hWndConsole = GetConsoleWindow()
            ShowWindow(hWndConsole, SW_HIDE)

            If File.Exists(plik_do_importu) Then
                Dim plik = My.Computer.FileSystem.GetFileInfo(plik_do_importu)

                DoLogu("Plik: " + plik_do_importu.ToString + " został znaleziony.")
                'zapis danych na bazę

                If plik_do_importu = "asdas" Then
                    Dim aaasd As String = 1
                End If


                Zapis_danych_na_baze(plik, zrodlo_meteo)
                DoLogu("Zakończono import danych.")

            Else
                DoLogu("Nie znaleziono pliku: " + plik_do_importu)

            End If

        Else


            Dim nCol As Integer = 0

        End If



    End Sub







    Private Sub DoLogu(str As String, Optional poziom As Integer = 1)
        Tools.WriteToFile3(str, nazwa_pliku_log, poziom)

    End Sub








    Sub Zapis_danych_na_baze(singleFile As IO.FileInfo, zrodlo_meteo As String)
        Dim data_s As String
        Dim wartosc As Double
        Dim wartosc_s As String
        Dim wersja As String
        Dim model As String
        Dim Ostatni_W As Integer
        Dim Ostatni_K As Integer
        Dim dData As Date
        Dim dData_temp As Date
        Dim dGodzina As Date
        Dim typPrognozy As String
        Dim nazwaMeteo As String
        Dim licznik As Integer = 0
        Dim dt As DataTable = Nothing
        Dim nazwaPliku As String = ""
        Dim status_podczytu As Boolean = False
        Dim odejmij As Integer = 0
        Dim czy_15 As String = ""
        Dim weryf As String

        'Dim Directory As New IO.DirectoryInfo(sciezka_zrodlo)
        'Dim allFiles As IO.FileInfo() = Directory.GetFiles("*_Cumulus*.csv")
        ' Dim singleFile As IO.FileInfo

        'petla po wszystkich plikach
        'For Each singleFile In allFiles

        licznik = licznik + 1

        Try
            conn.Open()
        Catch ex As OracleException ' catches only Oracle errors
            Select Case ex.ErrorCode
                Case 1
                    DoLogu("Error attempting to insert duplicate data.")
                Case 12560
                    DoLogu("Baza danych jest niedostępna !")
                Case Else
                    DoLogu("Database error: " + ex.Message.ToString())
            End Select
        Catch ex As Exception
            DoLogu(ex.Message.ToString())
        Finally
            DoLogu("Połączono z bazą:" & baza)
        End Try


        Try



            sciezka_zrodlo = singleFile.Directory.FullName

            nazwaPliku = singleFile.Name

            DoLogu("  * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * ")
            DoLogu("Parsowanie pliku " & nazwaPliku.ToString)
            dt = FileToTable(sciezka_zrodlo & "\" & nazwaPliku, ";", False)

            weryf = nazwaPliku.Substring(10)

            If weryf.Contains("15") Then
                czy_15 = "_15"
            End If

            'do usu
            'Dim stat As Boolean
            'stat = Spr_czy_plik_juz_zapisany(conn, nazwaPliku)


            'jezeli czy_zapisac_wszystkie_pliki = true

            If Spr_czy_plik_juz_zapisany(conn, nazwaPliku) = False Then
                status_podczytu = True
            End If

            If czy_zapisac_wszystkie_pliki = True Then
                status_podczytu = True
            End If

            If status_podczytu = True Then

                typPrognozy = dt.Rows(0).Item(0).ToString
                nazwaMeteo = Mid(typPrognozy, 3, Len(typPrognozy))

                If nazwaMeteo.ToLower.Contains("cumulus") _
                Or nazwaMeteo.ToLower.Contains("conwx") Then

                    If zrodlo_meteo.ToLower = "cumulus" Then
                        odejmij = 1
                    End If

                    Ostatni_W = dt.Rows.Count
                    Ostatni_K = dt.Columns.Count

                    wersja = typPrognozy.Substring(0, 1)
                    model = typPrognozy.Substring(2, 7) 'Cumulus
                    dData = CDate(dt.Rows(1).Item(0))

                    dGodzina = (dt.Rows(0).Item(1))

                    Dim nazwaProfilu_N As String
                    Dim nazwaProfilu_N1 As String
                    Dim nazwaProfilu_N2 As String
                    Dim nazwaLokalizacji As String = ""
                    Dim ID_profil_N As Long
                    Dim ID_profil_N1 As Long
                    Dim ID_profil_N2 As Long

                    'petla po wierszach lokalizacji
                    For w = 2 To Ostatni_W - 1

                        '*************************************************************************************************************************
                        'Do usu
                        If w = 3 And tryb_testowy = True Then
                            w = Ostatni_W - 1
                        End If


                        nazwaLokalizacji = (dt.Rows(w).Item(0))

                        nazwaProfilu_N = nazwaLokalizacji & "_" & nazwaMeteo & "_N"
                        nazwaProfilu_N1 = nazwaLokalizacji & "_" & nazwaMeteo & "_N1"
                        nazwaProfilu_N2 = nazwaLokalizacji & "_" & nazwaMeteo & "_N2"

                        ID_profil_N = Pobierz_ID_z_Name(conn, nazwaProfilu_N)
                        If ID_profil_N = 0 Then
                            DoLogu("Nie znaleziono ID profilu:" & nazwaProfilu_N)
                        Else
                            DoLogu("Znaleziono profil:" & nazwaProfilu_N & " / ID:" + ID_profil_N.ToString, 2)

                        End If

                        ID_profil_N1 = Pobierz_ID_z_Name(conn, nazwaProfilu_N1)
                        If ID_profil_N1 = 0 Then
                            DoLogu("Nie znaleziono ID profilu:" & nazwaProfilu_N1)
                        Else
                            DoLogu("Znaleziono profil:" & nazwaProfilu_N1 & " / ID:" + ID_profil_N1.ToString, 2)
                        End If

                        ID_profil_N2 = Pobierz_ID_z_Name(conn, nazwaProfilu_N2)
                        If ID_profil_N2 = 0 Then
                            DoLogu("Nie znaleziono ID profilu:" & nazwaProfilu_N2)
                        Else
                            DoLogu("Znaleziono profil:" & nazwaProfilu_N2 & " / ID:" + ID_profil_N2.ToString, 2)


                        End If



                        'If czy_15 = "" Then
                        '    dData_temp = dData.AddHours(dGodzina.Hour)
                        'Else

                        'End If

                        dData_temp = dData.AddHours(dGodzina.Hour)

                        'petla po kolumnach
                        For k = 1 To Ostatni_K - 1 - odejmij


                            wartosc_s = (dt.Rows(w).Item(k)).ToString

                            wartosc_s = wartosc_s.Replace(".", ",")

                            If wartosc_s = "" Then
                                wartosc_s = 0
                            Else
                                wartosc = CDbl(wartosc_s)
                                If czy_15 = "" Then
                                    wartosc /= 4
                                End If
                            End If

                            wartosc_s = wartosc.ToString
                            wartosc_s = wartosc_s.Replace(",", ".")


                            If czy_15 = "" Then
                                'zapis 15'
                                For g = 1 To 4

                                    data_s = "TO_DATE('" & Format(dData_temp, "yyyy-MM-dd HH:mm:ss") & "', 'YYYY-MM-DD HH24:MI:SS')"

                                    'zapis na profil N
                                    If ID_profil_N > 0 Then
                                        InsertRowEnergy15(conn, ID_profil_N, data_s, wartosc_s)
                                    End If

                                    'zapis na profil N1
                                    If Format(dData.AddDays(1), "yyyyMMdd") = Format(dData_temp, "yyyyMMdd") Then
                                        If ID_profil_N1 > 0 Then
                                            InsertRowEnergy15(conn, ID_profil_N1, data_s, wartosc_s)
                                        End If
                                    End If

                                    'zapis na profil N2
                                    If Format(dData.AddDays(2), "yyyyMMdd") = Format(dData_temp, "yyyyMMdd") Then
                                        If ID_profil_N2 > 0 Then
                                            InsertRowEnergy15(conn, ID_profil_N2, data_s, wartosc_s)
                                        End If
                                    End If

                                    dData_temp = dData_temp.AddMinutes(15)
                                Next

                            Else
                                data_s = "TO_DATE('" & Format(dData_temp, "yyyy-MM-dd HH:mm:ss") & "', 'YYYY-MM-DD HH24:MI:SS')"

                                'zapis na profil N
                                If ID_profil_N > 0 Then
                                    InsertRowEnergy15(conn, ID_profil_N, data_s, wartosc_s)
                                End If

                                'zapis na profil N1
                                If Format(dData.AddDays(1), "yyyyMMdd") = Format(dData_temp, "yyyyMMdd") Then
                                    If ID_profil_N1 > 0 Then
                                        InsertRowEnergy15(conn, ID_profil_N1, data_s, wartosc_s)
                                    End If
                                End If

                                'zapis na profil N2
                                If Format(dData.AddDays(2), "yyyyMMdd") = Format(dData_temp, "yyyyMMdd") Then
                                    If ID_profil_N2 > 0 Then
                                        InsertRowEnergy15(conn, ID_profil_N2, data_s, wartosc_s)
                                    End If
                                End If


                                dData_temp = dData_temp.AddMinutes(15)

                            End If




                        Next
                        'DoLogu("Zapisano dane dla lokalizacji: " & nazwaLokalizacji)

                    Next

                End If



                'dodanie logu do bazy
                InsertRowHarmonogram_log(conn, "Import_" & zrodlo_meteo, nazwaPliku.ToString)

                conn.Close()
                dt.Reset()
            Else
                DoLogu("Plik pominięty. Został już wcześniej zapisany.")
            End If


            Przenies_plik(singleFile)


        Catch err As Exception
            conn.Close()
            DoLogu("Błąd parsowania pliku: .... Błąd numer:" & err.ToString)
        End Try
        'Next

        conn.Close()

    End Sub
    Function FileToTable(ByVal fileName As String, ByVal separator As String, isFirstRowHeader As Boolean) As DataTable
        Dim result As DataTable = Nothing

        'dt = Nothing
        Try
            If Not System.IO.File.Exists(fileName) Then Throw New ArgumentException("fileName", String.Format("The file does not exist : {0}", fileName))
            Dim dt As New System.Data.DataTable
            Dim isFirstLine As Boolean = True
            Using sr As New System.IO.StreamReader(fileName)
                While Not sr.EndOfStream
                    Dim data() As String = sr.ReadLine.Split(CType(separator, Char()), StringSplitOptions.None)
                    If isFirstLine Then
                        If isFirstRowHeader Then
                            For Each columnName As String In data
                                dt.Columns.Add(New DataColumn(columnName, GetType(String)))
                            Next
                            isFirstLine = True ' Signal that this row is NOT to be considered as data.
                        Else
                            For i As Integer = 1 To data.Length
                                dt.Columns.Add(New DataColumn(String.Format("Column_{0}", i), GetType(String)))
                            Next
                            isFirstLine = False ' Signal that this row IS to be considered as data.
                        End If
                    End If
                    If Not isFirstLine Then
                        dt.Rows.Add(data.ToArray)
                    End If
                    isFirstLine = False ' All subsequent lines shall be considered as data.
                End While
                result = dt
            End Using
        Catch ex As Exception
            Throw New Exception(String.Format("{0}.CSVToDatatable Error", GetType(DataTable).FullName), ex)
        End Try
        Return result
    End Function




    Function Pobierz_ID_z_Name(cn As OracleConnection, name As String) As Long
        Dim dr As OracleDataReader
        Dim id As Long = 0
        Try
            Using cmd As OracleCommand = New OracleCommand()
                Dim sql As String = "select id from formula where lower (name) = lower( '" & name & "')"
                ' Const sql As String = "skome.SET_ENERGY15_INPUT( 1,3,TO_DATE ('2021-04-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS'),10.1012,1,1,3)"
                cmd.Connection = cn
                'cmd.Parameters.Add(New OracleParameter("var1", id))
                'cmd.Parameters.Add(New OracleParameter("var2", data))
                'cmd.Parameters.Add(New OracleParameter("var3", wartosc))
                cmd.CommandText = sql
                cmd.CommandType = CommandType.Text
                cmd.ExecuteNonQuery()
                dr = cmd.ExecuteReader()
                dr.Read()
                If dr.HasRows = True Then
                    id = dr.Item("id")
                End If
                dr.Close()
            End Using
        Catch err As Exception
            DoLogu("Błąd pobierania ID formuły z nazwy:" & name & ". Błąd numer:" & err.ToString)
            Return 0
        End Try
        Return id
    End Function

    Function Spr_czy_plik_juz_zapisany(cn As OracleConnection, nazwa_pliku As String) As Boolean
        Dim rezult As Boolean = False
        Dim dr As OracleDataReader
        Dim id As Long
        Try
            Using cmd As OracleCommand = New OracleCommand()
                Dim sql As String = "select id from OZEN.HARMONOGRAM_LOG where lower (INFO) = lower ('" & nazwa_pliku & "') and lower (ZADANIE) = lower ('Import_" & zrodlo_meteo & "') "
                cmd.Connection = cn
                'cmd.Parameters.Add(New OracleParameter("var1", id))
                'cmd.Parameters.Add(New OracleParameter("var2", data))
                'cmd.Parameters.Add(New OracleParameter("var3", wartosc))
                cmd.CommandText = sql
                cmd.CommandType = CommandType.Text
                cmd.ExecuteNonQuery()
                dr = cmd.ExecuteReader()
                dr.Read()

                If dr.HasRows = True Then
                    id = dr.Item(0)
                    rezult = True
                End If
                dr.Close()

            End Using
        Catch err As Exception
            DoLogu("Błąd pobierania ID formuły z nazwy:" & nazwa_pliku & ". Błąd numer:" & err.ToString)
            Return 0
        End Try
        Return rezult
    End Function


    Function Spr_czy_dane_sa_juz_zapisane(cn As OracleConnection, data As Date, zrodlo_meteo As String, kod_lokalizacji As String) As Boolean
        Dim rezult As Boolean = False
        Dim dr As OracleDataReader
        Dim id As Long
        Try
            Using cmd As OracleCommand = New OracleCommand()


                '------------------------------**********************
                ' do zrobienia


                Dim sql As String = "select id from OZEN.HARMONOGRAM_LOG where lower (INFO) = lower ('" & data & "') and lower (ZADANIE) = lower ('Import_Cumulus') "

                cmd.Connection = cn
                'cmd.Parameters.Add(New OracleParameter("var1", id))
                'cmd.Parameters.Add(New OracleParameter("var2", data))
                'cmd.Parameters.Add(New OracleParameter("var3", wartosc))
                cmd.CommandText = sql
                cmd.CommandType = CommandType.Text
                cmd.ExecuteNonQuery()
                dr = cmd.ExecuteReader()
                dr.Read()

                If dr.HasRows = True Then
                    id = dr.Item(0)
                    rezult = True
                End If
                dr.Close()

            End Using
        Catch err As Exception
            DoLogu("Błąd pobierania danych do weryfikacji:" & data & ". Błąd numer:" & err.ToString)
            Return 0
        End Try
        Return rezult
    End Function



    Private Sub InsertRowEnergy15(cn As OracleConnection, id As Long, data As String, wartosc As String)
        'Using cn As OracleConnection = New OracleConnection(connectionString)
        'cn.Open()
        Try
            Using cmd As OracleCommand = New OracleCommand()
                'Const sql As String = "Insert into test_table (val1, val2) values (:var1, :var2)"
                Dim sql As String = "SKOME.SET_ENERGY15_INPUT(1, " & id.ToString & "," & data & " , " & wartosc.ToString & ",1,1,3)"
                ' Const sql As String = "skome.SET_ENERGY15_INPUT( 1,3,TO_DATE ('2021-04-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS'),10.1012,1,1,3)"
                cmd.Connection = cn
                'cmd.Parameters.Add(New OracleParameter("var1", id))
                'cmd.Parameters.Add(New OracleParameter("var2", data))
                'cmd.Parameters.Add(New OracleParameter("var3", wartosc))
                cmd.CommandText = sql
                cmd.CommandType = CommandType.StoredProcedure
                cmd.ExecuteNonQuery()
            End Using
            'End Using
        Catch ex As Exception
            DoLogu("Błąd zapisu na bazę Energy15 -- " & Err.ToString)
        Finally
            DoLogu("Zapisano dane dla id:" + id.ToString, 2)
        End Try
    End Sub

    Private Sub InsertRowHarmonogram_log(cn As OracleConnection, zadanie As String, info As String)
        'Using cn As OracleConnection = New OracleConnection(connectionString)
        'cn.Open()
        Try
            Using cmd As OracleCommand = New OracleCommand()
                'Const sql As String = "Insert into test_table (val1, val2) values (:var1, :var2)"
                Dim sql As String = "INSERT INTO OZEN.HARMONOGRAM_LOG (CZAS_ZAPISU, INFO, ZADANIE)VALUES (TO_DATE (SYSDATE, 'YYYY-MM-DD HH24:MI:SS'),'" & info & "','" & zadanie & "')"
                cmd.Connection = cn
                'cmd.Parameters.Add(New OracleParameter("var1", id))
                'cmd.Parameters.Add(New OracleParameter("var2", data))
                'cmd.Parameters.Add(New OracleParameter("var3", wartosc))
                cmd.CommandText = sql
                cmd.CommandType = CommandType.Text
                cmd.ExecuteNonQuery()
            End Using
            'End Using
        Catch ex As Exception
            DoLogu("Błąd zapisu na bazę Energy15 -- " & Err.ToString)
        End Try
    End Sub





    Sub Przenies_plik(singleFile As IO.FileInfo)
        Try
            If File.Exists(singleFile.FullName) Then

                'spr czy istnieje katalog docelowy
                If Directory.Exists(singleFile.Directory.FullName & "/archiwum/" & Format(Now, "yyyyMM")) = False Then
                    Directory.CreateDirectory(singleFile.Directory.FullName & "/archiwum/" & Format(Now, "yyyyMM"))
                End If
                ' singleFile.MoveTo(singleFile.Directory.FullName & "/archiwum/" & Format(Now, "yyyyMM") & "/" & singleFile.Name)

                If File.Exists(singleFile.Directory.FullName & "/archiwum/" & Format(Now, "yyyyMM") & "/" & singleFile.Name) Then
                Else

                    Dim aa As String = singleFile.Directory.FullName & "/archiwum/" & Format(Now, "yyyyMM") & "/"
                    'singleFile.CopyTo(singleFile.Directory.FullName & "\archiwum\" & Format(Now, "yyyyMM") & "\")
                    File.Move(singleFile.FullName, singleFile.Directory.FullName & "\archiwum\" & Format(Now, "yyyyMM") & "\" & singleFile.Name)
                End If
                singleFile.Delete()
            End If
            DoLogu("Do archium:" & singleFile.FullName)
            ' Next
        Catch err As Exception
            DoLogu("Błąd przenoszenia pliku:" & singleFile.FullName & " do archiwum. Numer błędu:" & err.ToString)
        End Try


    End Sub










End Module
'End Class