'������ֱ�ӽ�Ҫɸѡ��ͬ�е��ļ��ϵ����vbs�ļ��ϼ���
'���ߣ�sysdzw
'���䣺sysdzw@163.com
'QQ��171977759
'12:51 2009-7-12
Dim strFileSource, strFileResult,t1

On Error Resume Next
strFileSource = wscript.Arguments(0)
strFileResult = Left(strFileSource, InStrRev(strFileSource, ".") - 1) & "_out.txt"

If strFileSource <> "" Then
    t1=Time()
    Set fso = CreateObject("scripting.filesystemobject")
    Set stream = fso.opentextfile(strFileSource, 1, False)
    Set stream2 = fso.opentextfile(strFileResult, 2, True)

    Set dict = CreateObject("scripting.dictionary")


    While Not stream.atendofstream
        Line = stream.readline
        If Not dict.Exists(Line) Then
            Call dict.Add(Line, Null)
            Call stream2.writeline(Line)
        End If
    Wend
    
    stream.Close
    stream2.Close
    MsgBox "������ϣ��ܼƺ�ʱ " & DateDiff("s",t1,Time) & " �롣" & vbCrLf & vbCrLf & strFileResult, vbInformation, "Del Same Line QQ:171977759"
Else
    MsgBox "no file!", vbExclamation, "Del Same Line QQ:171977759"
End If