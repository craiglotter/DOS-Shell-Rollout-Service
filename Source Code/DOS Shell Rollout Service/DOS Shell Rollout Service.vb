Imports System.Serviceprocess
Imports Microsoft.Win32
Imports System.IO


Public Class DOS_Shell_Rollout_Service
    Inherits System.Serviceprocess.ServiceBase

    Dim submitfolder As String = ""

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call
    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.Serviceprocess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.Serviccurrentprocessess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.Serviceprocess.ServiceBase() {New DOS_Shell_Rollout_Service}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    Friend WithEvents Main_Timer As System.Timers.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Main_Timer = New System.Timers.Timer
        CType(Me.Main_Timer, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Main_Timer
        '
        Me.Main_Timer.Enabled = True
        Me.Main_Timer.Interval = 60000
        '
        'DOS_Shell_Rollout_Service
        '
        Me.ServiceName = "DOS_Shell_Rollout"
        CType(Me.Main_Timer, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

#End Region

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try
            Activity_Logger("Service Successfully Started")
            submitfolder = GetSubmitFolder()
            If Not submitfolder = "" Then
                Code_Execute()
            End If
            Main_Timer.Start()
        Catch ex As Exception
            Error_Handler(ex, "OnStart")
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Try
            Main_Timer.Stop()
            Activity_Logger("Service Successfully Stopped")
        Catch ex As Exception
            Error_Handler(ex, "OnStop")
        End Try
    End Sub

    Function GetSubmitFolder() As String
        Dim result As String = ""
        Try
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine

            Dim oKey As RegistryKey
            oKey = oReg
            Dim subs() As String = ("Software\Microsoft\Windows\CurrentVersion\Run").Split("\")
            For Each stri As String In subs
                oKey = oKey.OpenSubKey(stri, True)
            Next

            If Not oKey Is Nothing Then
                str = oKey.GetValue("InvisibleAppStarter")
                If Not IsNothing(str) And Not (str = "") Then
                    result = str
                    result = result.Replace("""", "").Replace("\Application_Loader.exe", "\Monitor")
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "GetSubmitFolder")
        End Try
        Return result
    End Function

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "DOS Shell Rollout Service\Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "DOS Shell Rollout Service\Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
            filewriter.Flush()
            filewriter.Close()
        Catch exc As Exception
            Dim mylog As New EventLog
            If Not mylog.SourceExists("DOS Shell Rollout Service") Then
                mylog.CreateEventSource("DOS Shell Rollout Service", "DOS Shell Rollout Service Log")
            End If
            mylog.Source = "DOS Shell Rollout Service"
            mylog.WriteEntry("DOS Shell Rollout Service Log", "Error Handler Failure: " & exc.ToString, EventLogEntryType.Error)
            mylog.Close()
        End Try
    End Sub

    Public Sub Activity_Logger(ByVal message As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "DOS Shell Rollout Service\Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((System.Environment.SystemDirectory & "\").Replace("\\", "\") & "DOS Shell Rollout Service\Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & message)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity_Logger")
        End Try
    End Sub



    'Private Sub WorkerStatusMessageUpdate(ByVal message As String)
    '    Try
    '        Dim mylog As New EventLog
    '        If Not mylog.SourceExists("DOS Shell Rollout Service") Then
    '            mylog.CreateEventSource("DOS Shell Rollout Service", "DOS Shell Rollout Service Log")
    '        End If
    '        mylog.Source = "DOS Shell Rollout Service"
    '        mylog.WriteEntry("DOS Shell Rollout Service Log", message, EventLogEntryType.Information)
    '        mylog.Close()


    '    Catch ex As Exception
    '        Error_Handler(ex, "Worker Status Message Update")
    '    End Try
    'End Sub

    Private Sub Main_Timer_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Main_Timer.Elapsed
        Try
            submitfolder = GetSubmitFolder()
            If Not submitfolder = "" Then
                Code_Execute()
            End If
        Catch ex As Exception
            Error_Handler(ex, "Code_Execute")
        End Try
    End Sub

    Private Function Load_Registry_Values() As String
        Dim str As String = ""
        Try
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\DOS Shell Rollout") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\DOS Shell Rollout")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\DOS Shell Rollout", True)


            str = oKey.GetValue("executablePath")


            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex, "Load Registry Value")
        End Try
        Return str
    End Function

    Private Sub Code_Execute()
        Try
            Dim processname As String = "DOS Shell Rollout"
            Dim processexecutable As String = Load_Registry_Values()
            If processexecutable = "" Or processexecutable Is Nothing Then
                processexecutable = "C:\Program Files\DOS Shell Rollout\Application_Loader.exe"
            End If

            Dim existing As Process() = Process.GetProcesses
            Dim currentprocess As Process
            Dim foundproc As Boolean = False
            For Each currentprocess In existing
                Try
                    Dim currentprocessname As String = currentprocess.ProcessName
                    'Activity_Logger(currentprocessname)
                    If currentprocess.ProcessName.ToLower.Equals(processname.ToLower) = True Then
                        foundproc = True
                    End If
                Catch ex As Exception
                    Error_Handler(ex, "Testing Process Names")
                End Try
            Next


            If foundproc = False Then
                Activity_Logger("Process not Found. Submitting to Invisible Application Starter: " & processexecutable)
                If Not submitfolder = "" Then
                    Try

                        Dim testfile As FileInfo = New FileInfo(processexecutable.Replace("""", ""))
                        If testfile.Exists = True Then



                            Dim finfo As FileInfo = New FileInfo((submitfolder & "\DOS_Shell_Rollout_Service.txt").Replace("\\", "\"))
                            If finfo.Exists = False Then
                                Dim filewriter As StreamWriter = New StreamWriter((submitfolder & "\DOS_Shell_Rollout_Service.txt").Replace("\\", "\"), False)
                                filewriter.WriteLine(processexecutable)
                                filewriter.Flush()
                                filewriter.Close()
                                filewriter = Nothing
                            End If
                            finfo = Nothing
                        Else
                            Activity_Logger("Process Executable cannot be located.")
                        End If
                    Catch ex As Exception
                        Error_Handler(ex, "Trying Application Launch")
                    End Try

                Else
                    Activity_Logger("Unable to Determine Submit Folder")
                End If
            End If

        Catch ex As Exception
            Error_Handler(ex, "Code Execute")
        End Try
    End Sub


End Class
