' modWMML.vbs
Option Explicit

' ���� JSON ģ��
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("modJSON.vbs").ReadAll()
On Error GoTo 0

' Global variables to replace form controls
Dim gJavaPath, gMemoryMB, gUseCustomMemory

' Initialize global variables (should be called before using other functions)
Sub InitializeGlobals(javaPath, memoryMB, useCustomMemory)
    gJavaPath = javaPath
    gMemoryMB = memoryMB
    gUseCustomMemory = useCustomMemory
End Sub

' �ú�������������� Minecraft
' (string)mcPath .minecraft�ļ���·��
' (string)versionName �汾����
' (string)playerName �������
Sub LaunchMinecraft(mcPath, versionName, playerName)
    'On Error Resume Next
    
    ' ��׼��·��
    If Right(mcPath, 1) <> "\" Then mcPath = mcPath & "\"
    
    ' ��ȡ�汾json�ļ�
    Dim versionJsonPath, jsonContent
    versionJsonPath = mcPath & "versions\" & versionName & "\" & versionName & ".json"
    jsonContent = ReadTextFile(versionJsonPath)
    
    If Err.Number <> 0 Then
        MsgBox "����Minecraftʱ����: �޷���ȡ�汾JSON�ļ� - " & Err.Description, vbCritical, "����"
        Exit Sub
    End If

    If jsonContent = "" Then
        MsgBox "�޷���ȡ�汾 JSON �ļ���������·��������ļ���", vbCritical, "����"
        Exit Sub
    End If
    
    ' ����JSON
    Dim versionJson
    Set versionJson = CreateObject("Scripting.Dictionary")
    ParseJSONString2 jsonContent, versionJson
    
    If Err.Number <> 0 Then
        MsgBox "����Minecraftʱ����: ����JSONʧ�� - " & Err.Description, vbCritical, "����"
        Exit Sub
    End If
    
    ' ��ȡ����
    Dim mainClass
    mainClass = versionJson("mainClass")
    
    ' ������·��
    Dim libraries
    libraries = BuildLibrariesPath(mcPath, versionJson)
    
    ' ������Ϸ����
    Dim gameArgs
    gameArgs = BuildGameArguments(mcPath, versionName, playerName, versionJson)
    
    ' ����Java����
    Dim javaCommand, configlx, conststr
    configlx = ""
    If Not gUseCustomMemory Then
        configlx = "-Xmx" & CInt(gMemoryMB) & "M -Xms" & CInt(gMemoryMB) & "M "
    End If
    
    ' Note: App.Major, App.Minor, App.Revision are not available in VBScript
    ' You'll need to define these as global variables or constants
    conststr = configlx & "-Dfile.encoding=GB18030 -Dsun.stdout.encoding=GB18030 -Dsun.stderr.encoding=GB18030 " & _
               "-Djava.rmi.server.useCodebaseOnly=true -Dcom.sun.jndi.rmi.object.trustURLCodebase=false " & _
               "-Dcom.sun.jndi.cosnaming.object.trustURLCodebase=false -Dlog4j2.formatMsgNoLookups=true " & _
               "-Dlog4j.configurationFile=.minecraft\versions\" & versionName & "\log4j2.xml " & _
               "-Dminecraft.client.jar=.minecraft\versions\" & versionName & "\" & versionName & ".jar " & _
               "-XX:+UnlockExperimentalVMOptions -XX:+UseG1GC -XX:G1NewSizePercent=20 -XX:G1ReservePercent=20 " & _
               "-XX:MaxGCPauseMillis=50 -XX:G1HeapRegionSize=32m -XX:-UseAdaptiveSizePolicy " & _
               "-XX:-OmitStackTraceInFastThrow -XX:-DontCompileHugeMethods -Dfml.ignoreInvalidMinecraftCertificates=true " & _
               "-Dfml.ignorePatchDiscrepancies=true -XX:HeapDumpPath=MojangTricksIntelDriversForPerformance_javaw.exe_minecraft.exe.heapdump " & _
               "-Djava.library.path=.minecraft\versions\" & versionName & "\natives-windows-x86_64 " & _
               "-Djna.tmpdir=.minecraft\versions\" & versionName & "\natives-windows-x86_64 " & _
               "-Dorg.lwjgl.system.SharedLibraryExtractPath=.minecraft\versions\" & versionName & "\natives-windows-x86_64 " & _
               "-Dio.netty.native.workdir=.minecraft\versions\" & versionName & "\natives-windows-x86_64 " & _
               "-Dminecraft.launcher.brand=WMML -Dminecraft.launcher.version=1.0.0" ' Replace with your version
    
    javaCommand = "cmd /K " & gJavaPath & " " & conststr & " -cp """ & libraries & """ " & mainClass & " " & gameArgs
    
    ' ִ������
    WScript.Echo "Generated command: " & javaCommand
    CreateObject("WScript.Shell").Run javaCommand, 1, False
    
    If Err.Number <> 0 Then
        MsgBox "����Minecraftʱ����: " & Err.Description, vbCritical, "����"
    End If
End Sub

' ������·��
Function BuildLibrariesPath(mcPath, versionJson)
    On Error Resume Next
    
    Dim libs, lib, libPath, result
    
    ' ������Ӱ汾jar�ļ�
    result = mcPath & "versions\" & versionJson("id") & "\" & versionJson("id") & ".jar"
    
    ' ������п��ļ�
    libs = versionJson("libraries")
    For Each lib In libs
        ' ������(�����������ڵ�ǰϵͳ�Ŀ�)
        If Not CheckLibraryRules(lib) Then 
            ' In VBScript, we use Continue For (but it's not available)
            ' So we just skip to next iteration
        Else
            ' ��ȡ��·��
            libPath = GetLibraryPath(mcPath, lib)
            
            ' ��ӵ����
            If libPath <> "" Then
                If result <> "" Then result = result & ";"
                result = result & libPath
            End If
        End If
    Next
    
    BuildLibrariesPath = result
End Function

' �������
Function CheckLibraryRules(lib)
    On Error Resume Next
    
    ' ���û��rules���֣������ǰ���
    If Not lib.Exists("rules") Then
        CheckLibraryRules = True
        Exit Function
    End If
    
    Dim rule, osName, osArch
    
    ' ��ȡ��ǰϵͳ��Ϣ
    osArch = "x86_64" ' Assuming 64-bit for VBScript version
    osName = "windows"
    
    ' ������й���
    For Each rule In lib("rules")
        ' ���action
        If rule("action") = "allow" Then
            ' ���û��os���֣�������
            If Not rule.Exists("os") Then
                CheckLibraryRules = True
                Exit Function
            End If
            
            ' ���os����
            If rule("os")("name") = osName Then
                ' �����arch���������arch
                If rule("os").Exists("arch") Then
                    If rule("os")("arch") = osArch Then
                        CheckLibraryRules = True
                        Exit Function
                    Else
                        CheckLibraryRules = False
                        Exit Function
                    End If
                Else
                    CheckLibraryRules = True
                    Exit Function
                End If
            Else
                CheckLibraryRules = False
                Exit Function
            End If
        ElseIf rule("action") = "disallow" Then
            ' ���û��os���֣�������
            If Not rule.Exists("os") Then
                CheckLibraryRules = False
                Exit Function
            End If
            
            ' ���os����
            If rule("os")("name") = osName Then
                CheckLibraryRules = False
                Exit Function
            End If
        End If
    Next
    
    ' Ĭ������
    CheckLibraryRules = True
End Function

' ��ȡ��·��
Function GetLibraryPath(mcPath, lib)
    On Error Resume Next
    
    Dim parts, artifactPath, nativePath, i, classifier
    
    ' ����������
    parts = Split(lib("name"), ":")
    
    ' ��������·��
    artifactPath = mcPath & "libraries\" & Replace(parts(0), ".", "\") & "\" & parts(1) & "\" & parts(2) & "\" & parts(1) & "-" & parts(2)
    
    ' ����Ƿ���natives
    If lib.Exists("natives") Then
        ' ��ȡwindowsƽ̨��native������
        If lib("natives").Exists("windows") Then
            classifier = lib("natives")("windows")
            classifier = Replace(classifier, "${arch}", "64") ' Assuming 64-bit
            
            ' ����native·��
            nativePath = artifactPath & "-" & classifier & ".jar"
            
            ' ����ļ��Ƿ����
            If FileExists(nativePath) Then
                GetLibraryPath = nativePath
                Exit Function
            End If
        End If
    End If
    
    ' ���û��natives���Ҳ���native�ļ���ʹ����ͨjar
    artifactPath = artifactPath & ".jar"
    If FileExists(artifactPath) Then
        GetLibraryPath = artifactPath
        Exit Function
    End If
    
    GetLibraryPath = ""
End Function

' ������Ϸ����
Function BuildGameArguments(mcPath, versionName, playerName, versionJson)
    On Error Resume Next
    
    Dim args, assetsPath, versionType, arg
    
    ' ����Ĭ��ֵ
    assetsPath = mcPath & "assets"
    versionType = "WMML"
    
    ' ���Դ�json��ȡassets��versionType
    If versionJson.Exists("assets") Then
        assetsPath = mcPath & "assets\"
    End If
    
    If versionJson.Exists("type") Then
        versionType = versionJson("type")
    End If
    
    ' ��������
    args = ""
    
    ' ����Ƿ���minecraftArguments(�ɰ汾)
    If versionJson.Exists("minecraftArguments") Then
        args = versionJson("minecraftArguments") & " " & args
    End If
    
    ' ����Ƿ���arguments(�°汾)
    If versionJson.Exists("arguments") Then
        For Each arg In versionJson("arguments")("game")
            If VarType(arg) = vbString Then
                args = args & " " & arg
            End If
        Next
    End If
    
    args = Replace(args, "${auth_player_name}", playerName)
    args = Replace(args, "${version_name}", versionName)
    args = Replace(args, "${game_directory}", mcPath)
    args = Replace(args, "${assets_root}", assetsPath)
    args = Replace(args, "${assets_index_name}", versionJson("assets"))
    args = Replace(args, "${auth_uuid}", "00000000-0000-0000-0000-000000000000")
    args = Replace(args, "${auth_access_token}", "00000000000000000000000000000000")
    args = Replace(args, "${user_type}", "legacy")
    args = Replace(args, "${version_type}", """WMML 0.1.26""")
    
    BuildGameArguments = args
End Function

' ��������: ��ȡ�ı��ļ�
Function ReadTextFile(filePath)
    On Error Resume Next
    
    Dim fso, file, content
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1)
        content = file.ReadAll
        file.Close
        ReadTextFile = content
    Else
        ReadTextFile = ""
    End If
End Function

' ��������: ����ļ��Ƿ����
Function FileExists(filePath)
    On Error Resume Next
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function

' Initialize global variables
InitializeGlobals "java.exe", 4096, False
'
' Launch Minecraft 1.20.1 for player "player123"
LaunchMinecraft ".minecraft\", "1.20.1", "player123"