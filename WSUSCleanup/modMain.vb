Imports Microsoft.UpdateServices.Administration
Imports System.Environment
Module modMain
    Sub Main()
        Dim arrCommandLineArgs As String()

        Try
            arrCommandLineArgs = GetCommandLineArgs()

            If arrCommandLineArgs.GetUpperBound(0) > 0 Then
                Dim strCommandLineArg As String
                Dim blnRun As Boolean
                Dim blnCOC As Boolean
                Dim blnCOU As Boolean
                Dim blnCUCF As Boolean
                Dim blnCU As Boolean
                Dim blnDEU As Boolean
                Dim blnDSU As Boolean

                For Each strCommandLineArg In arrCommandLineArgs
                    Select Case strCommandLineArg.ToUpper
                        Case "COC"
                            blnCOC = True
                            blnRun = True
                        Case "COU"
                            blnCOU = True
                            blnRun = True
                        Case "CUCF"
                            blnCUCF = True
                            blnRun = True
                        Case "CU"
                            blnCU = True
                            blnRun = True
                        Case "DEU"
                            blnDEU = True
                            blnRun = True
                        Case "DSU"
                            blnDSU = True
                            blnRun = True
                    End Select
                Next

                If blnRun = True Then
                    Dim objAdminProxy As New AdminProxy
                    Dim objIUpdateServer As IUpdateServer
                    Dim objICleanupManager As ICleanupManager
                    Dim objCleanupScope As New CleanupScope

                    objIUpdateServer = objAdminProxy.GetUpdateServer
                    objICleanupManager = objIUpdateServer.GetCleanupManager()

                    objCleanupScope.CleanupObsoleteComputers = blnCOC          'COC
                    objCleanupScope.CleanupObsoleteUpdates = blnCOU            'COU
                    objCleanupScope.CleanupUnneededContentFiles = blnCUCF      'CUCF
                    objCleanupScope.CompressUpdates = blnCU                    'CU
                    objCleanupScope.DeclineExpiredUpdates = blnDEU             'DEU
                    objCleanupScope.DeclineSupersededUpdates = blnDSU          'DSU

                    Console.WriteLine("Working...")
                    objICleanupManager.PerformCleanup(objCleanupScope)
                    Console.WriteLine("Done")
                Else
                    ShowUsage()
                End If
            Else
                ShowUsage()
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
    Private Sub ShowUsage()
        Console.WriteLine("Performs a WSUS Server Cleanup")
        Console.WriteLine("")
        Console.WriteLine("CleanupObsoleteComputers = COC")
        Console.WriteLine("CleanupObsoleteUpdates = COU")
        Console.WriteLine("CleanupUnneededContentFiles = CUCF")
        Console.WriteLine("CompressUpdates = CU")
        Console.WriteLine("DeclineExpiredUpdates = DEU")
        Console.WriteLine("DeclineSupersededUpdates = DSU")
    End Sub

End Module
