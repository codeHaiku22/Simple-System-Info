Imports System.Management
Public Class frmMain

    Private mstrCurrentTargetServer As String
    Private mstrPreviousTargetServer As String
    Private mintSparkLineLength = 30
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        lblServerNameIP.Text = ""

    End Sub
    Private Sub cmdScan_Click(sender As Object, e As EventArgs) Handles cmdScan.Click

        If Len(txtServerNameIP.Text.Trim) = 0 Then Exit Sub

        mstrPreviousTargetServer = mstrCurrentTargetServer
        mstrCurrentTargetServer = txtServerNameIP.Text.Trim

        If mstrPreviousTargetServer = mstrCurrentTargetServer Then Exit Sub

        lblServerNameIP.Text = Get_HostName_Or_IPAddress(mstrCurrentTargetServer)

        lbxDetails.Items.Clear()

        GetSystemInfo()

    End Sub
    Private Function Get_HostName_Or_IPAddress(ByVal strHostNameorIPAddress As String) As String

        Dim intError As Integer

        Try

            Dim blnIsIPAddress As Boolean = IsNumeric(Replace(strHostNameorIPAddress, ".", ""))
            Dim strHostName As String
            Dim strIPAddress As String

            If blnIsIPAddress Then
                strHostName = System.Net.Dns.GetHostEntry(strHostNameorIPAddress).HostName.ToString
                Get_HostName_Or_IPAddress = strHostName
            Else
                For Each ipAddress As System.Net.IPAddress In System.Net.Dns.GetHostAddresses(strHostNameorIPAddress)
                    If ipAddress.AddressFamily.ToString = "InterNetwork" Then
                        strIPAddress = strIPAddress & "/" & ipAddress.ToString
                    End If
                Next
                If strIPAddress.Length > 0 Then strIPAddress = Replace(strIPAddress, "/", "", 1, 1)
                Get_HostName_Or_IPAddress = strIPAddress
            End If

        Catch ex As Exception

            intError = Err.Number

            MsgBox("Error retrieving Host Name and/or IP Address.", vbCritical, "SysInfo")

        End Try

    End Function
    Private Sub GetSystemInfo()

        Dim intError As Integer

        Try

            Dim strOSManufacturer As String
            Dim strOSName As String
            Dim strOSVersion As String
            Dim strOSArchitecture As String
            Dim strOSInstallDate As String
            Dim strOSLastBootUpTime As String
            Dim strWindowsDir As String
            Dim strComputerName As String
            Dim sngPhysicalMemoryFree As Single
            Dim sngPhysicalMemoryTotal As Single
            Dim sngPhysicalMemoryUsage As Single
            Dim strManufacturer As String
            Dim strModel As String
            Dim strSystemType As String
            Dim strProcessorsPhysical As String
            Dim strCPUManufacturer As String
            Dim strCPUName As String
            Dim strCPUCaption As String
            Dim strCPUClockSpeed As String
            Dim strCPUCores As String
            Dim strCPUProcessorsLogical As String
            Dim sngCPULoad As String
            Dim strDiskModel As String
            Dim lngDiskSize As Long
            Dim strLogicalDiskDeviceId As String
            Dim strLogicalDiskFS As String
            Dim strLogicalDiskName As String
            Dim sngLogicalDiskFree As Single
            Dim sngLogicalDiskSize As Single
            Dim strLogicalDiskVolumeName As String
            Dim i As Integer
            Dim strScope As String = "\\" & mstrCurrentTargetServer & "\root\cimv2"
            Dim objOS As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_OperatingSystem")
            Dim objCS As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_ComputerSystem")
            Dim objCPU As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_Processor")
            Dim objDisk As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_DiskDrive")
            Dim objLogicalDisk As New ManagementObjectSearcher(strScope, "SELECT * FROM Win32_LogicalDisk")
            Dim objMgmt As ManagementObject

            For Each objMgmt In objOS.Get
                strOSManufacturer = IIf(IsNothing(objMgmt("manufacturer")), "", objMgmt("manufacturer").ToString)
                strOSName = IIf(IsNothing(objMgmt("name")), "", objMgmt("name").ToString)
                strOSVersion = IIf(IsNothing(objMgmt("version")), "", objMgmt("version").ToString)
                strOSArchitecture = IIf(IsNothing(objMgmt("osarchitecture")), "", objMgmt("osarchitecture").ToString)
                strOSInstallDate = IIf(IsNothing(objMgmt("installdate")), "", objMgmt("installdate").ToString)
                strOSLastBootUpTime = IIf(IsNothing(objMgmt("lastbootuptime")), "", objMgmt("lastbootuptime").ToString)
                'strWindowsDir = IIf(IsNothing(objMgmt("windowsdirectory")), "", objMgmt("windowsdirectory").ToString)
                'strComputerName = IIf(IsNothing(objMgmt("csname")), "", objMgmt("csname").ToString)
                sngPhysicalMemoryFree = IIf(IsNothing(objMgmt("freephysicalmemory")), 0, (Convert.ToInt64(objMgmt("freephysicalmemory")))) '/ 1024000)
            Next

            For Each objMgmt In objCS.Get
                sngPhysicalMemoryTotal = IIf(IsNothing(objMgmt("totalphysicalmemory")), 0, (Convert.ToInt64(objMgmt("totalphysicalmemory")))) ' / 1024000000)
                strManufacturer = IIf(IsNothing(objMgmt("manufacturer")), "", objMgmt("manufacturer").ToString)
                strModel = IIf(IsNothing(objMgmt("model")), "", objMgmt("model").ToString)
                strSystemType = IIf(IsNothing(objMgmt("systemtype")), "", objMgmt("systemtype").ToString)
                strProcessorsPhysical = IIf(IsNothing(objMgmt("numberofprocessors")), "", objMgmt("numberofprocessors").ToString)
            Next

            For Each objMgmt In objCPU.Get
                strCPUManufacturer = IIf(IsNothing(objMgmt("manufacturer")), "", objMgmt("manufacturer").ToString)
                strCPUName = IIf(IsNothing(objMgmt("name")), "", objMgmt("name").ToString)
                strCPUCaption = IIf(IsNothing(objMgmt("caption")), "", objMgmt("caption").ToString)
                strCPUClockSpeed = IIf(IsNothing(objMgmt("maxclockspeed")), "", objMgmt("maxclockspeed").ToString)
                strCPUCores = IIf(IsNothing(objMgmt("numberofcores")), "", objMgmt("numberofcores").ToString)
                strCPUProcessorsLogical = IIf(IsNothing(objMgmt("numberoflogicalprocessors")), "", objMgmt("numberoflogicalprocessors").ToString)
                sngCPULoad = IIf(IsNothing(objMgmt("loadpercentage")), "", Convert.ToInt64(objMgmt("loadpercentage")))
            Next

            If InStr(strOSName, "|") > 0 Then
                Dim vArray() As String
                vArray = Split(strOSName, "|")
                strOSName = vArray(0)
            End If

            sngPhysicalMemoryUsage = sngPhysicalMemoryTotal - (sngPhysicalMemoryFree * 1024)

            lbxDetails.Items.Add("System Information:")
            lbxDetails.Items.Add(strManufacturer & " " & strModel & " (" & strSystemType & ")")
            lbxDetails.Items.Add(strCPUManufacturer & " " & strCPUName)
            lbxDetails.Items.Add(strCPUCaption)
            lbxDetails.Items.Add("Processors: " & strCPUProcessorsLogical & " (" & strCPUCores & " Cores) @" & strCPUClockSpeed & " MHz")
            lbxDetails.Items.Add("Total Physical Memory: " & CalculateBestByteSize(sngPhysicalMemoryTotal))
            lbxDetails.Items.Add(strOSManufacturer & " " & strOSName & " " & strOSArchitecture & " (" & strOSVersion & ")")
            lbxDetails.Items.Add("OS Installed: " & Mid(strOSInstallDate, 5, 2) & "/" & Mid(strOSInstallDate, 7, 2) & "/" & Strings.Left(strOSInstallDate, 4) &
                                 " @" & Mid(strOSInstallDate, 9, 2) & ":" & Mid(strOSInstallDate, 11, 2) & ":" & Mid(strOSInstallDate, 13, 2))
            lbxDetails.Items.Add("Last Boot: " & Mid(strOSLastBootUpTime, 5, 2) & "/" & Mid(strOSLastBootUpTime, 7, 2) & "/" & Strings.Left(strOSLastBootUpTime, 4) &
                                 " @" & Mid(strOSLastBootUpTime, 9, 2) & ":" & Mid(strOSLastBootUpTime, 11, 2) & ":" & Mid(strOSLastBootUpTime, 13, 2))
            lbxDetails.Items.Add("Uptime: " & CalculateSystemUpTime(strOSLastBootUpTime))

            For Each objMgmt In objDisk.Get
                strDiskModel = IIf(IsNothing(objMgmt("model")), "", objMgmt("model").ToString)
                lngDiskSize = IIf(IsNothing(objMgmt("size")), "", CLng(objMgmt("size")))
                lbxDetails.Items.Add("Physical Disk (" & i & "): " & strDiskModel & " [" & CalculateBestByteSize(lngDiskSize, True) & "]")
                i = i + 1
            Next

            lbxDetails.Items.Add(" ")
            lbxDetails.Items.Add("Resource Usage:")
            lbxDetails.Items.Add(("CPU").PadRight(15) & ": " & BuildSparkLine(sngCPULoad / 100))
            lbxDetails.Items.Add(" ")
            lbxDetails.Items.Add(("Memory").PadRight(15) & ": " & BuildSparkLine(sngPhysicalMemoryUsage / sngPhysicalMemoryTotal) & "  " & CalculateBestByteSize(sngPhysicalMemoryUsage) & "/" & CalculateBestByteSize(sngPhysicalMemoryTotal))
            lbxDetails.Items.Add(" ")

            i = 0
            For Each objMgmt In objLogicalDisk.Get
                sngLogicalDiskSize = IIf(IsNothing(objMgmt("size")), 0, Convert.ToInt64(objMgmt("size")))
                If sngLogicalDiskSize = 0 Then Continue For
                strLogicalDiskDeviceId = IIf(IsNothing(objMgmt("deviceid")), "", objMgmt("deviceid").ToString)
                strLogicalDiskFS = IIf(IsNothing(objMgmt("filesystem")), "", objMgmt("filesystem").ToString)
                strLogicalDiskName = IIf(IsNothing(objMgmt("name")), "", objMgmt("name").ToString)
                sngLogicalDiskFree = IIf(IsNothing(objMgmt("freespace")), 0, Convert.ToInt64(objMgmt("freespace")))
                strLogicalDiskVolumeName = IIf(IsNothing(objMgmt("volumename")), "", objMgmt("volumename").ToString)
                lbxDetails.Items.Add((Replace(strLogicalDiskDeviceId, ":", "") & " " & strLogicalDiskVolumeName).PadRight(15) & ": " & BuildSparkLine((sngLogicalDiskSize - sngLogicalDiskFree) / sngLogicalDiskSize) & "  " & (CalculateBestByteSize(sngLogicalDiskSize - sngLogicalDiskFree, True) & "/" & CalculateBestByteSize(sngLogicalDiskSize, True)).PadRight(13) & "[" & strLogicalDiskFS & "]")
                lbxDetails.Items.Add(" ")
                i = i + 1
            Next

        Catch ex As Exception

            intError = Err.Number

            MsgBox("Error retrieving WMI information.", vbCritical, "SysInfo")

        End Try

    End Sub
    Private Function CalculateSystemUpTime(ByVal strLastBootUpTime As String, Optional blnFriendlyUpTime As Boolean = True) As String

        Dim intError As Integer

        Try

            Dim dtmLastBootUpTime As DateTime = Mid(strLastBootUpTime, 5, 2) & "/" & Mid(strLastBootUpTime, 7, 2) & "/" & Strings.Left(strLastBootUpTime, 4) & " " & Mid(strLastBootUpTime, 9, 2) & ":" & Mid(strLastBootUpTime, 11, 2) & ":" & Mid(strLastBootUpTime, 13, 2)
            Dim dtmCurrent As DateTime = DateTime.Now
            Dim tsSystemUpTime As TimeSpan = dtmCurrent - dtmLastBootUpTime
            Dim strStystemUpTime As String

            If blnFriendlyUpTime Then
                Dim strTimeSpan As String = tsSystemUpTime.ToString
                Dim strDays As String = Mid(strTimeSpan, 1, InStr(strTimeSpan, ".") - 1) & " days "
                Dim strHours As String = Mid(strTimeSpan, InStr(strTimeSpan, ".") + 1, 2) & " hours "
                Dim strMinutes As String = Mid(strTimeSpan, InStr(strTimeSpan, ":") + 1, 2) & " minutes "
                Dim strSeconds As String = Mid(strTimeSpan, InStr(InStr(strTimeSpan, ":") + 1, strTimeSpan, ":") + 1, 2) & " seconds"
                strStystemUpTime = strDays & strHours & strMinutes & strSeconds
            Else
                strStystemUpTime = tsSystemUpTime.ToString
            End If

            CalculateSystemUpTime = strStystemUpTime

        Catch ex As Exception

            intError = Err.Number

            MsgBox("Error calculating system up time.", vbCritical, "SysInfo")

        End Try


    End Function
    Private Function CalculateBestByteSize(ByVal lngBytes As Long, Optional blnMetric As Boolean = False) As String

        Dim intError As Integer

        Try

            Dim strBestByteSize As String

            If blnMetric Then
                Select Case lngBytes
                    Case < 1000
                        strBestByteSize = lngBytes.ToString & "B"
                    Case < 1000000
                        strBestByteSize = Format(lngBytes / 1000, "#,###,###") & "KB"
                    Case < 1000000000
                        strBestByteSize = Format(lngBytes / 1000000, "#,###,###") & "MB"
                    Case < 1000000000000
                        strBestByteSize = Format(lngBytes / 1000000000, "#,###,###") & "GB"
                    Case < 1000000000000000
                        strBestByteSize = Format(lngBytes / 1000000000000, "#,###,###") & "TB"
                    Case < 1000000000000000000
                        strBestByteSize = Format(lngBytes / 1000000000000000, "#,###,###") & "PB"
                    Case Else
                        strBestByteSize = ""
                End Select
            Else
                Select Case lngBytes
                    Case < 1024
                        strBestByteSize = lngBytes.ToString & "B"
                    Case < 1048576
                        strBestByteSize = Format(lngBytes / 1024, "#,###,###") & "KB"
                    Case < 1073741824
                        strBestByteSize = Format(lngBytes / 1048576, "#,###,###") & "MB"
                    Case < 1099511627776
                        strBestByteSize = Format(lngBytes / 1073741824, "#,###,###") & "GB"
                    Case < 1125899906842624
                        strBestByteSize = Format(lngBytes / 1099511627776, "#,###,###") & "TB"
                    Case < 1152921504606846976
                        strBestByteSize = Format(lngBytes / 1125899906842624, "#,###,###") & "PB"
                    Case Else
                        strBestByteSize = ""
                End Select
            End If

            CalculateBestByteSize = strBestByteSize

        Catch ex As Exception

            intError = Err.Number

            MsgBox("Error calculating byte size conversion.", vbCritical, "SysInfo")

        End Try


    End Function
    Private Function BuildSparkLine(ByVal sngUsagePercentage As Single, Optional blnShowPercent As Boolean = True) As String

        Dim intError As Integer

        Try

            Dim strSparkLine As String
            Dim strSpark As String = ChrW(&H2588) 'H2588, H2589, H258A, H258B, H258C, H258D, H258E, H258F  Thick to thin
            Dim strBlank As String = " "
            Dim intRepeatSparks As Integer = sngUsagePercentage * mintSparkLineLength
            Dim intRepeatBlanks As Integer = mintSparkLineLength - intRepeatSparks

            strSparkLine = "[" & String.Concat(Enumerable.Repeat(strSpark, intRepeatSparks)) & String.Concat(Enumerable.Repeat(strBlank, intRepeatBlanks)) & "]"

            If blnShowPercent Then
                strSparkLine = strSparkLine & " " & Format(sngUsagePercentage * 100, "00") & "%"
            End If

            BuildSparkLine = strSparkLine

        Catch ex As Exception

            intError = Err.Number

            MsgBox("Error creating usage spark line.", vbCritical, "SysInfo")

        End Try

    End Function

End Class
'Parking Lot
'Only works for local PC
'Dim pc As PerformanceCounter = New PerformanceCounter("System", "System Up Time")
'pc.NextValue()
'Dim ts As TimeSpan = TimeSpan.FromSeconds(pc.NextValue)
'Dim strTimeSpan As String = ts.ToString
'Dim strDays As String = Mid(strTimeSpan, 1, InStr(strTimeSpan, ".") - 1) & "days "
'Dim strHours As String = Mid(strTimeSpan, InStr(strTimeSpan, ".") + 1, 2) & "hours "
'Dim strMinutes As String = Mid(strTimeSpan, InStr(strTimeSpan, ":") + 1, 2) & "minutes "
'Dim strSeconds As String = Mid(strTimeSpan, InStr(InStr(strTimeSpan, ":") + 1, strTimeSpan, ":") + 1, 2) & "seconds"