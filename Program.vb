Imports System.IO
Imports System.Net
Imports System.Web
Imports Newtonsoft.Json.Linq

Module Program
    Sub Main(args As String())
        'validateOnExport("D:\TamilVanan\Temp\OC259JOB.mdb")
        Dim rd As New ReadExcel
        'Console.Write("Before Read")
        'Dim outputArray() As String = rd.ImportExcel("D:\TamilVanan\Temp\OC271JOB 09-Apr-21 2-58-21 PM.xls")
        'Console.WriteLine("Planned: " + outputArray(0))
        'Console.WriteLine("Exported: " + outputArray(1))
        'Dim prop As New Properties
        'prop.Load("D:\TamilVanan\Temp\email_properties.properties")

        'Dim se As New SendEmail
        'Console.Write("Before Read")
        'se.sendEmail(prop, "Test Message", "D:\TamilVanan\Temp\OC271JOB 09-Apr-21 2-58-21 PM.xls")
        Dim filePath As String = "D:\TamilVanan\Temp\OC259JOB 23-Mar-21 7-23-20 PM.xls"
        Dim length = filePath.Length
        Dim startIndex = filePath.LastIndexOf("\") + 1
        Dim sunLen = length - filePath.LastIndexOf("\") - 1

        Dim fileName As String = filePath.Substring(startIndex, sunLen)
        Console.WriteLine(filePath)
        Console.WriteLine(fileName)
        Console.WriteLine("Sending")
        Dim query As String = "OC259JOB 23-Mar-21 7-23-20 PM.xls"
        query = HttpUtility.UrlEncode(query)

        Dim process As String = "01 Main Assembly"
        Dim lineNo As String = "Line_No_1"
        Dim machineName As String = "Column Assy"
        Dim machineDrwaing As String = "GRADE"
        Dim subFolder As String = "INVBOM"
        Dim partReferenceNumber As String = "112227  SAMBATH KUMAR S"
        Dim plannedPages As String = "10"

        Dim url As String = "http://127.0.0.1:8080/MappingServlet/EmailServlet?file_name=" + query + "&template_name=TestTemplate &process=" + HttpUtility.UrlEncode(process) +
                                                                                                    "&lineNo=" + HttpUtility.UrlEncode(lineNo) + "&machineName=" + HttpUtility.UrlEncode(machineName) +
                                                                                                    "&machineDrwaing=" + HttpUtility.UrlEncode(machineDrwaing) + "&subFolder=" + HttpUtility.UrlEncode(subFolder) +
                                                                                                    "&partReferenceNumber=" + HttpUtility.UrlEncode(partReferenceNumber) + "&plannedPages=" + HttpUtility.UrlEncode(plannedPages)

        Console.WriteLine(url)
        WebrequestWithPost(url, "D:\TamilVanan\Temp\OC259JOB 23-Mar-21 7-23-20 PM.xls", "application/x-www-form-urlencoded")
        Console.WriteLine("Sent")
    End Sub


    Public Function validateOnExport(ByVal sMDBPath As String)
        Dim mdbPathString() As String = sMDBPath.Split("\")
        Dim mdbFileName = mdbPathString(mdbPathString.Length - 1)
        Dim mdbFileNameArr() As String = mdbFileName.Split(".")
        Dim jobId As String = mdbFileNameArr(0)
        Dim indexOfSlash As Integer = sMDBPath.LastIndexOf("\")
        Dim dbPath As String = sMDBPath.Substring(0, indexOfSlash)
        Dim reportPath As String = dbPath & "\Report\"

        searchFile(reportPath, jobId)
    End Function

    Private Sub searchFile(reportPath As String, jobId As String)
        Dim Folder As New IO.DirectoryInfo(reportPath)
        Dim fileNames As List(Of String) = New List(Of String)
        Dim wildcard = jobId & "*.xls"
        Console.WriteLine(wildcard)

        For Each File As IO.FileInfo In Folder.GetFiles(wildcard, IO.SearchOption.TopDirectoryOnly)
            fileNames.Add(File.FullName)
        Next

        For Each fileName As String In fileNames
            Console.WriteLine("FileInReportPath: " & fileName)
        Next
    End Sub

    Public Function WebrequestWithPost(ByVal url As String, ByVal fileName As String, ByVal contentType As String) As String()
        Dim postDataAsByteArray As Byte() = FileToByteArray(fileName)
        Dim returnValue As String = String.Empty
        Dim outputArray() As String = New String(1) {"0", "0"}
        Try
            Dim webRequest As HttpWebRequest = webRequest.CreateHttp(url)
            If (Not (webRequest) Is Nothing) Then
                webRequest.AllowAutoRedirect = False
                webRequest.Method = "POST"
                webRequest.ContentType = contentType
                webRequest.ContentLength = postDataAsByteArray.Length
                Dim requestDataStream As Stream = webRequest.GetRequestStream
                requestDataStream.Write(postDataAsByteArray, 0, postDataAsByteArray.Length)
                requestDataStream.Close()

                Dim response As WebResponse = webRequest.GetResponse
                Dim responseDataStream As Stream = response.GetResponseStream
                If (Not (responseDataStream) Is Nothing) Then
                    Dim responseDataStreamReader As StreamReader = New StreamReader(responseDataStream)
                    returnValue = responseDataStreamReader.ReadToEnd
                    responseDataStreamReader.Close()
                    responseDataStream.Close()
                End If
                response.Close()
                requestDataStream.Close()
            End If
        Catch ex As WebException
            If (ex.Status = WebExceptionStatus.ProtocolError) Then
                Dim response As HttpWebResponse = CType(ex.Response, HttpWebResponse)
                'handle this your own way.
                Console.WriteLine("Webexception! Statuscode: {0}, Description: {1}", CType(response.StatusCode, Integer), response.StatusDescription)
            End If
        Catch ex As Exception
            'handle this your own way, something serious happened here.
            Console.WriteLine(ex.Message)
        End Try

        Dim jObect = JObject.Parse(returnValue)

        outputArray(0) = jObect.SelectToken("plannedPages")
        outputArray(1) = jObect.SelectToken("exportedPages")

        Return outputArray
    End Function

    Public Function FileToByteArray(ByVal _FileName As String) As Byte()
        Dim _Buffer() As Byte = Nothing

        Try
            ' Open file for reading
            Dim _FileStream As New System.IO.FileStream(_FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)

            ' attach filestream to binary reader
            Dim _BinaryReader As New System.IO.BinaryReader(_FileStream)

            ' get total byte length of the file
            Dim _TotalBytes As Long = New System.IO.FileInfo(_FileName).Length

            ' read entire file into buffer
            _Buffer = _BinaryReader.ReadBytes(CInt(Fix(_TotalBytes)))

            ' close file reader
            _FileStream.Close()
            _FileStream.Dispose()
            _BinaryReader.Close()
        Catch _Exception As Exception
            ' Error
            Console.WriteLine("Exception caught in process: {0}", _Exception.ToString())
        End Try

        Return _Buffer
    End Function
End Module
