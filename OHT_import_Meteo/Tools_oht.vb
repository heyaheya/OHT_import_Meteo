Imports System.IO


Namespace OHT_import_Meteo

    Public Class Tools_OHT



        Public Function Check_internet()
            If My.Computer.Network.IsAvailable() Then
                Check_internet = ("Sieć dostępna.")
            Else
                Check_internet = ("Sieć niedostępna !")
            End If
        End Function

        Public Sub WriteToFile(text As String)
            'Dim path As String = "C:\temp\" & DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss").ToString & "_ServiceLog.txt"if 
            'Dim text2 As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss").ToString
            Dim path As String = "C:\temp\meteo\" + DateTime.Now.ToString("yyyyMMdd").ToString + "_Podczyt_Cumulus_log.txt"
            text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss").ToString & " " & text

            If Directory.Exists("C:\temp\meteo\") = False Then
                Directory.CreateDirectory("C:\temp\meteo\")
            End If

            Using writer As New StreamWriter(path, True)
                writer.WriteLine(text)
                writer.Close()
            End Using


            If OHT.status_logu Then

            End If



        End Sub

        Public Sub WriteToFile2(text As String, Optional poziom As Integer = 1)
            'Dim path As String = "C:\temp\" & DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss").ToString & "_ServiceLog.txt"if 
            'Dim text2 As String = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss").ToString



            If OHT.status_logu >= poziom Then

                Dim path As String = "C:\temp\meteo\" + DateTime.Now.ToString("yyyyMMdd").ToString + "_Podczyt_Cumulus_log.txt"
                'text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss").ToString & " " & text

                Using writer As New StreamWriter(path, True)
                    writer.WriteLine(text)
                    writer.Close()
                End Using


            End If



        End Sub

        Public Sub WriteToFile3(text As String, nazwa As String, Optional poziom As Integer = 1)
            If OHT.status_logu >= poziom Then
                Dim path As String = "C:\temp\meteo\" + DateTime.Now.ToString("yyyyMMdd").ToString + "_import_" & nazwa & "_log.txt"
                text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss").ToString & " " & text
                Using writer As New StreamWriter(path, True)
                    writer.WriteLine(text)
                    writer.Close()
                End Using
            End If
        End Sub

        'uruchamianie palikacji i wysyłanie znaków klawiszy
        Sub wyslij()
            Dim ProcID As Integer
            ' Start the Calculator application, and store the process id.
            ProcID = Shell("CALC.EXE", AppWinStyle.NormalFocus)
            ' Activate the Calculator application.
            AppActivate(ProcID)
            ' Send the keystrokes to the Calculator application.
            My.Computer.Keyboard.SendKeys("22", True)
            My.Computer.Keyboard.SendKeys("*", True)
            My.Computer.Keyboard.SendKeys("44", True)
            My.Computer.Keyboard.SendKeys("=", True)
            ' The result is 22 * 44 = 968.
        End Sub

        Function Get_ConnectionString(baza As String) As String
            Get_ConnectionString = ""
            Select Case UCase(baza)
                Case "ERGH"
                    Get_ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=" _
                       + "(ADDRESS=(PROTOCOL=TCP)(HOST = SIPHA.energa.loc)(PORT = 1521)))" _
                       + "(CONNECT_DATA = (Server = DEDICATE)(SERVICE_NAME = SRV_ERGH)));" _
                       + "User Id=skome;Password=szafran1;"

                Case "ERGH_NEW"
                    Get_ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=" _
                       + "(ADDRESS=(PROTOCOL=TCP)(HOST = SIPHA.energa.loc)(PORT = 1521)))" _
                       + "(CONNECT_DATA = (Server = DEDICATE)(SERVICE_NAME = SRV_ERGH)));" _
                       + "User Id=skome;Password=szafran1;"

                Case "ERGH_TEST"
                    Get_ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=" _
                       + "(ADDRESS=(PROTOCOL=TCP)(HOST = SIPHA-DR.energa.loc)(PORT = 1521)))" _
                       + "(CONNECT_DATA = (Server = DEDICATE)(SERVICE_NAME = SRV_ERGHT)));" _
                       + "User Id=skome;Password=szafran1"';Pooling=False"

                Case "WIRE"
                    Get_ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=" _
                       + "(ADDRESS=(PROTOCOL=TCP)(HOST = SIPHA.energa.loc)(PORT = 1521)))" _
                       + "(CONNECT_DATA = (Server = DEDICATE)(SERVICE_NAME = SRV_WIRE)));" _
                       + "User Id=skome;Password=szafran1;"
                Case "HEEI"
                    Get_ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=" _
                       + "(ADDRESS=(PROTOCOL=TCP)(HOST = SIPHA.energa.loc)(PORT = 1521)))" _
                       + "(CONNECT_DATA = (Server = DEDICATE)(SERVICE_NAME = SRV_HEEI)));" _
                       + "User Id=skome;Password=szafran1;"
                Case "MM"
                    Get_ConnectionString = "server=eobx-s-00372;uid=oht;pwd=P@ssw0rd;database=mm"

                Case "ESP_ZYDOWO_372"
                    Get_ConnectionString = "server=eobx-s-00372;uid=oht;pwd=P@ssw0rd;database=esp_zydowo"

                Case "ESP_ZYDOWO_223"
                    Get_ConnectionString = "server=eobx-s-00223;uid=oht;pwd=P@ssw0rd;database=esp_zydowo"


                Case Else
                    WriteToFile2("W funkcji 'Get_ConnectionString' nie rozpoznano bazy: " & baza, 1)
            End Select
        End Function







    End Class
End Namespace