Imports Newtonsoft.Json.Linq
Imports System.Web
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Exception
Imports System.Net
Imports Newtonsoft.Json
Public Class TencentUpload
    Dim appid As Long
    Dim COSAPI_CGI_URL As String = "http://web.file.myqcloud.com/files/v1/"
    Dim secretid As String
    Dim secretkey As String

    Sub New(appid As Long, secretid As String, secretkey As String)
        Me.appid = appid
        Me.secretid = secretid
        Me.secretkey = secretkey
    End Sub

    '腾讯云上传文件函数开始(下面几个函数全部是参照腾讯云对象存储服务的  .NET SDK(C#)来写的，虽然都是用的.net 但是C#和VB 的语法问题debug了好长时间)
    '产生unix 时间戳函数
    Function ToUnixTime(nowTime As DateTime) As Long
        Dim startTime As DateTime = TimeZone.CurrentTimeZone.ToLocalTime(New System.DateTime(1970, 1, 1, 0, 0, 0, 0))
        Return CLng(Math.Round((nowTime - startTime).TotalMilliseconds, MidpointRounding.AwayFromZero))
    End Function
    '生成签名
    Function Signature(appId As Integer, secretId As String, secretKey As String, expired As Long, fileId As String, bucketName As String) As String
        If secretId = "" Or secretKey = "" Then
            Return "-1"
        End If
        Dim current = Now
        Dim nowTime = CLng(ToUnixTime(current) / 1000)
        Randomize()
        Dim rand = New Random()
        Dim rdm = rand.Next(Int32.MaxValue)
        Dim plainText = "a=" + CStr(appId) + "&k=" + secretId + "&e=" + CStr(expired) + "&t=" + CStr(nowTime) + "&r=" + CStr(rdm) + "&b=" + bucketName + "&f=" + fileId
        Console.WriteLine(plainText)

        Using mac As HMACSHA1 = New HMACSHA1(Encoding.UTF8.GetBytes(secretKey))
            Dim hash = mac.ComputeHash(Encoding.UTF8.GetBytes(plainText))
            Dim pText = Encoding.UTF8.GetBytes(plainText)
            Dim all As Byte() = New Byte(hash.Length + pText.Length - 1) {}

            Array.Copy(hash, 0, all, 0, hash.Length)
            Array.Copy(pText, 0, all, hash.Length, pText.Length)
            Return Convert.ToBase64String(all)
        End Using

    End Function
    '规范化remotePath
    Function EncodeRemotePath(remotePath As String) As String
        If remotePath = "/" Then
            Return remotePath
        End If
        Dim endWith As Boolean = remotePath.EndsWith("/")
        Dim part As String() = remotePath.Split("/")
        remotePath = ""
        For Each s As String In part
            If s <> "" Then
                If remotePath <> "" Then
                    remotePath = remotePath + "/"
                End If
                remotePath = remotePath + HttpUtility.UrlEncode(s)
            End If
        Next
        If remotePath.StartsWith("/") Then
            remotePath = "" + remotePath
        Else
            remotePath = "/" + remotePath
        End If
        If endWith Then
            remotePath = remotePath + "/"
        Else
            remotePath = remotePath + ""
        End If
        Return remotePath

    End Function
    '上传文件接口函数
    Function UploadFile(bucketName As String, remotePath As String, localPath As String) As String
        Dim cos_path = EncodeRemotePath(remotePath)
        MsgBox(cos_path)
        Dim url = COSAPI_CGI_URL + CStr(appid) + "/" + bucketName + EncodeRemotePath(remotePath)
        Dim SHA1 = GetSHA1(localPath)
        Dim data = New Dictionary(Of String, String)
        data.Add("op", "upload")
        data.Add("sha", SHA1)
        Dim current = Now
        Dim expired = CLng(ToUnixTime(current) / 1000) + 60

        Dim sign = Signature(appid, secretid, secretkey, expired, "/" + CStr(appid) + "/" + bucketName + cos_path, bucketName)
        System.Console.WriteLine(sign)
        System.Console.WriteLine(url)
        Dim header = New Dictionary(Of String, String)
        header.Add("Authorization", sign)
        Return SendRequest(url, data, "post", header, 60000, localPath, -1, 0)
    End Function
    '获取文件的sha值
    Function GetSHA1(filePath As String) As String

        Dim strResult = ""
        Dim strHashData = ""
        Dim arrbytHashValue As Byte() = New Byte() {}
        Dim oFileStream As FileStream
        Dim osha1 As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider()

        Try
            oFileStream = New FileStream(filePath.Replace(Chr(34), ""), FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            arrbytHashValue = osha1.ComputeHash(oFileStream)
            oFileStream.Close()

            strHashData = BitConverter.ToString(arrbytHashValue)

            strHashData = strHashData.Replace("-", "")
            strResult = strHashData
        Catch ex As Exception
            Throw ex
        End Try


        Return strResult
    End Function
    '发送请求
    Function SendRequest(url As String, data As Dictionary(Of String, String), requestMethod As String,
             header As Dictionary(Of String, String), timeOut As Integer, localPath As String, offset As Long, sliceSize As Integer) As String
        '        System.Net.ServicePointManager.Expect100Continue = False
        Dim request As HttpWebRequest
        System.Console.WriteLine(url)
        Console.WriteLine(JsonConvert.SerializeObject(data))
        Console.WriteLine(JsonConvert.SerializeObject(header))
        Try
            System.Net.ServicePointManager.Expect100Continue = False
            If requestMethod = "get" Then
                Dim paramStr = ""
                Dim keys = data.Keys
                For Each key As String In data.Keys
                    paramStr += String.Format("{0}={1}&", key, HttpUtility.UrlEncode(data(key).ToString()))
                Next
                paramStr = paramStr.TrimEnd("&")
                If url.EndsWith("?") Then
                    url = url + "&" + paramStr
                Else
                    url = url + "?" + paramStr
                End If
            End If

            request = HttpWebRequest.Create(url)
            request.Accept = "*/*"
            request.KeepAlive = True

            request.UserAgent = ""
            request.Timeout = timeOut

            For Each hk As String In header.Keys
                If hk = "Content-Type" Then
                    request.ContentType = header(hk)
                Else
                    request.Headers.Add(hk, header(hk))
                End If
            Next

            If requestMethod = "post" Then

                request.Method = "POST"
                Dim memStream = New MemoryStream()

                If header.ContainsKey("Content-Type") Then
                    If header.Item("Content-Type") = "application/json" Then
                        Dim json = JsonConvert.SerializeObject(data)
                        Dim jsonByte = Encoding.GetEncoding("utf-8").GetBytes(json.ToString())
                        memStream.Write(jsonByte, 0, jsonByte.Length)
                    End If
                Else
                    Dim boundary = "---------------" + DateTime.Now.Ticks.ToString("x")
                    Dim beginBoundary = Encoding.ASCII.GetBytes(Chr(13) + Chr(10) + "--" + boundary + Chr(13) + Chr(10))
                    Dim endBoundary = Encoding.ASCII.GetBytes(Chr(13) + Chr(10) + "--" + boundary + "--" + Chr(13) + Chr(10))
                    request.ContentType = "multipart/form-data; boundary=" + boundary

                    Dim strBuf = New StringBuilder()
                    For Each dk As String In data.Keys
                        strBuf.Append(Chr(13) + Chr(10) + "--" + boundary + Chr(13) + Chr(10))
                        strBuf.Append("Content-Disposition: form-data; name=" + Chr(34) + dk + Chr(34) + Chr(13) + Chr(10) + Chr(13) + Chr(10))  'Chr(13) + Chr(10) + Chr(13) + Chr(10))+
                        strBuf.Append(data(dk).ToString())
                    Next

                    Dim paramsByte = Encoding.GetEncoding("utf-8").GetBytes(strBuf.ToString())
                    memStream.Write(paramsByte, 0, paramsByte.Length)

                    If localPath <> "" Then
                        memStream.Write(beginBoundary, 0, beginBoundary.Length)
                        Dim fileInfo = New FileInfo(localPath)
                        Dim fileStream = New FileStream(localPath, FileMode.Open, FileAccess.Read)
                        Const filePartHeader As String =
                            "Content-Disposition: form-data; name=" + Chr(34) + "fileContent" + Chr(34) + "; filename=" + Chr(34) + "{0}" + Chr(34) + Chr(13) + Chr(10) +
                            "Content-Type: application/octet-stream" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
                        Console.WriteLine(fileInfo.Name)
                        Dim headerText = String.Format(filePartHeader, fileInfo.Name)

                        Dim headerbytes = Encoding.UTF8.GetBytes(headerText)
                        memStream.Write(headerbytes, 0, headerbytes.Length)

                        If offset = -1 Then
                            Dim buffer = New Byte(1024) {}
                            Dim bytesRead As Integer = fileStream.Read(buffer, 0, buffer.Length)
                            While bytesRead <> 0
                                Dim txt As String = Encoding.UTF8.GetString(buffer)
                                memStream.Write(buffer, 0, bytesRead)
                                bytesRead = fileStream.Read(buffer, 0, buffer.Length)
                            End While
                        Else
                            Dim buffer = New Byte(1024) {}
                            Dim bytesRead As Integer
                            fileStream.Seek(offset, SeekOrigin.Begin)
                            bytesRead = fileStream.Read(buffer, 0, buffer.Length)
                            memStream.Write(buffer, 0, bytesRead)
                        End If
                        memStream.Write(endBoundary, 0, endBoundary.Length)
                    End If

                    request.ContentLength = memStream.Length

                    Dim requestStream = request.GetRequestStream()
                    memStream.Position = 0
                    Dim tempBuffer = New Byte(memStream.Length - 1) {}
                    memStream.Read(tempBuffer, 0, tempBuffer.Length)
                    Dim memtxt = Encoding.UTF8.GetString(tempBuffer)
                    Console.WriteLine(memtxt)
                    memStream.Close()

                    requestStream.Write(tempBuffer, 0, tempBuffer.Length)
                    requestStream.Close()
                End If
                Dim response = request.GetResponse()
                MsgBox(response.ContentType)
                Using s = response.GetResponseStream()
                    Dim reader = New StreamReader(s, Encoding.UTF8)
                    Return reader.ReadToEnd()
                End Using

            End If

        Catch we As WebException

            If we.Status = WebExceptionStatus.ProtocolError Then
                Using s = we.Response.GetResponseStream()
                    Dim reader = New StreamReader(s, Encoding.UTF8)
                    Return reader.ReadToEnd()
                End Using
            Else
                Throw we
            End If

        Catch e As Exception
            Throw e
        End Try

    End Function

End Class
