# VBSDKUploadFileForTencent

### 使用步骤
>1. 将TencentUpload.vb文件引入到自己的VB项目中

>2. 在要使用到改功能的文件中用Imports引入该文件

>3. 创建该类的一个实例，调用UploadFile方法

### Demo
```
Imports JX3.TencentUpload  '这里的JX3只是我当时的项目名字，使用的时候请替换
Dim tencentUpload As TencentUpload = New TencentUpload(appid, secretid, secretkey)  '这里的appid  secretid   secretkey是自己的腾讯云中的
Dim a As String = tencentUpload.UploadFile(bucket, "/jx3NamePic/aaaeee.txt", "D:\aa.txt")

'a 就是最后返回的json格式的字符串  类似于
{    
    "code":0,
    "message":"SUCCESS",
    "data":{
        "access_url":"http://gamepic-10053455.file.myqcloud.com/jx3NamePic/aaaeeeb.txt",
        "resource_path":"/jx3NamePic/aaaeeeb.txt",
        "source_url":"http://gamepic-10053455.cos.myqcloud.com/jx3NamePic/aaaeeeb.txt",
        "url":"http://web.file.myqcloud.com/files/v1/jx3NamePic/aaaeeeb.txt"
    }
}
```
