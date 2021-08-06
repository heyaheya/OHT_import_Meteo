


Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Data.OracleClient
Imports DataTable = System.Data.DataTable
Imports System.Data.OleDb

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





    Dim FTP_import_Cumulus As New OHT_import_Meteo.FTP_import_Cumulus
    Dim Tools As New OHT_import_Meteo.Tools_OHT

    Public katalog_pliku_log As String
    Public baza As String = "ERGH_TEST"    'Dim baza As String = "ERGH"
    Public conn As New OracleConnection(Tools.Get_ConnectionString(baza))

    Public sciezka_zrodlo As String = "Y:\OHT\Dana Pomiarowe\Cumulus_3\"
    Public czy_zapisac_wszystkie_pliki As Boolean = False

    Public status_logu As Integer = 1
    '1 - poziom podst
    '2 - poziom serwisowy 
    '
    Public tryb_testowy As Boolean = False


    Sub Main()



        Dim StatusPodczytu As Boolean = False

        'spr sieci
        DoLogu("Start")
        DoLogu(Tools.Check_internet)

        DoLogu("Ilość argumentów wywołania aplikacji:" & My.Application.CommandLineArgs.Count, 1)


        'Dim aaa As String
        'Dim k As Integer


        For Each arg As String In My.Application.CommandLineArgs
            DoLogu("Argumenty wywołania aplikacji:" & arg.ToString, 1)
            If arg.Trim("-") = "v" Then
                StatusPodczytu = True
                DoLogu("Argument 'StatusPodczytu':" & arg.ToString, 1)
            End If

            'DoLogu(arg.Length)
            'For k = 0 To arg.Length - 1
            '	DoLogu("k" & k & "   " & arg.Chars(k).ToString)
            'Next

            If (arg.Length = 3) Then
                If (arg.Chars(1).ToString = "p") Then
                    'Dim MyChar() As Char = {"-", "p"}
                    'Dim str As String = arg.TrimStart(MyChar)
                    If CInt(arg.Chars(2).ToString) < 3 Then
                        status_logu = CInt(arg.Chars(2).ToString)
                        DoLogu("Poziom logu: " & status_logu.ToString, 1)
                        DoLogu("Argument 'Poziom logu':" & arg.ToString, 1)
                    End If
                End If
            End If

        Next



        If StatusPodczytu = True Then

            DoLogu("Start pobierania danych z FTP")
            FTP_import_Cumulus.PobierzDanezFTP(sciezka_zrodlo)
            DoLogu("Koniec pobierania danych")

            DoLogu("Zapis_danych_na_baze.")

            'zapis danych na bazę
            Zapis_danych_na_baze()
            DoLogu("Zakończono import danych.")


        Else


            Dim nCol As Integer = 0

        End If



    End Sub







    Private Sub DoLogu(str As String, Optional poziom As Integer = 1)
        Tools.WriteToFile2(str, poziom)

    End Sub








    Sub Zapis_danych_na_baze()
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

        Dim Directory As New IO.DirectoryInfo(sciezka_zrodlo)
        Dim allFiles As IO.FileInfo() = Directory.GetFiles("*_Cumulus*.csv")
        Dim singleFile As IO.FileInfo

        'petla po wszystkich plikach
        For Each singleFile In allFiles

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
                nazwaPliku = singleFile.Name

                DoLogu("  * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * ")
                DoLogu("Parsowanie pliku " & licznik & "\" & allFiles.Length + 1 & " - " & nazwaPliku.ToString)
                dt = FileToTable(sciezka_zrodlo & nazwaPliku, ";", False)

                If Spr_czy_juz_zapisany(conn, nazwaPliku) = False And czy_zapisac_wszystkie_pliki = False Then



                    typPrognozy = dt.Rows(0).Item(0).ToString
                    nazwaMeteo = Mid(typPrognozy, 3, Len(typPrognozy))

                    If nazwaMeteo = "Cumulus_prog_wiatr_sila" _
                        Or nazwaMeteo = "Cumulus_prog_wiatr_kierunek" _
                        Or nazwaMeteo = "Cumulus_prog_slonce_strumien" _
                        Or nazwaMeteo = "Cumulus_prog_slonce_zachm_calk" _
                        Or nazwaMeteo = "Cumulus_prog_slonce_zachm_konw" _
                        Or nazwaMeteo = "Cumulus_prog_slonce_zachm_nis" _
                        Or nazwaMeteo = "Cumulus_prog_slonce_zachm_sred" _
                        Or nazwaMeteo = "Cumulus_prog_slonce_zachm_wys" _
                        Or nazwaMeteo = "Cumulus_prog_wiatr_gestosc" _
                        Then


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
                            End If

                            ID_profil_N1 = Pobierz_ID_z_Name(conn, nazwaProfilu_N1)
                            If ID_profil_N1 = 0 Then
                                DoLogu("Nie znaleziono ID profilu:" & nazwaProfilu_N1)
                            End If

                            ID_profil_N2 = Pobierz_ID_z_Name(conn, nazwaProfilu_N2)
                            If ID_profil_N2 = 0 Then
                                DoLogu("Nie znaleziono ID profilu:" & nazwaProfilu_N2)
                            End If


                            'petla po kolumnach
                            For k = 1 To Ostatni_K - 2
                                dData_temp = dData.AddHours(dGodzina.Hour + k - 1)

                                wartosc_s = (dt.Rows(w).Item(k)).ToString
                                wartosc_s = wartosc_s.Replace(".", ",")
                                wartosc = CDbl(wartosc_s)
                                wartosc /= 4
                                wartosc_s = wartosc.ToString
                                wartosc_s = wartosc_s.Replace(",", ".")

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

                            Next
                            DoLogu("Zapisano dane dla lokalizacji: " & nazwaLokalizacji)

                        Next

                    End If



                    'dodanie logu do bazy
                    InsertRowHarmonogram_log(conn, "Import_Cumulus", nazwaPliku.ToString)

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
        Next

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

    Function Spr_czy_juz_zapisany(cn As OracleConnection, nazwa_pliku As String) As Boolean
        Dim rezult As Boolean = False
        Dim dr As OracleDataReader
        Dim id As Long
        Try
            Using cmd As OracleCommand = New OracleCommand()
                Dim sql As String = "select id from OZEN.HARMONOGRAM_LOG where lower (INFO) = lower ('" & nazwa_pliku & ".csv') and lower (ZADANIE) = lower ('Import_Cumulus') "
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
            DoLogu("Błąd pobierania ID formuły z nazwy:" & Name & ". Błąd numer:" & err.ToString)
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