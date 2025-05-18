' modWMML.vbs
Option Explicit

' 加载 JSON 模块
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

' 该函数生成命令并启动 Minecraft
' (string)mcPath .minecraft文件夹路径
' (string)versionName 版本名称
' (string)playerName 玩家名称
Sub LaunchMinecraft(mcPath, versionName, playerName)
    'On Error Resume Next
    
    ' 标准化路径
    If Right(mcPath, 1) <> "\" Then mcPath = mcPath & "\"
    
    ' 读取版本json文件
    Dim versionJsonPath, jsonContent
    versionJsonPath = mcPath & "versions\" & versionName & "\" & versionName & ".json"
    jsonContent = ReadTextFile(versionJsonPath)
    
    If Err.Number <> 0 Then
        MsgBox "启动Minecraft时出错: 无法读取版本JSON文件 - " & Err.Description, vbCritical, "错误"
        Exit Sub
    End If

    If jsonContent = "" Then
        MsgBox "无法读取版本 JSON 文件，可能是路径错误或文件损坏", vbCritical, "错误"
        Exit Sub
    End If
    
    ' 解析JSON
    Dim versionJson
    Set versionJson = CreateObject("Scripting.Dictionary")
    ParseJSONString2 jsonContent, versionJson
    
    If Err.Number <> 0 Then
        MsgBox "启动Minecraft时出错: 解析JSON失败 - " & Err.Description, vbCritical, "错误"
        Exit Sub
    End If
    
    ' 获取主类
    Dim mainClass
    mainClass = versionJson("mainClass")
    
    ' 构建库路径
    Dim libraries
    libraries = BuildLibrariesPath(mcPath, versionJson)
    
    ' 构建游戏参数
    Dim gameArgs
    gameArgs = BuildGameArguments(mcPath, versionName, playerName, versionJson)
    
    ' 构建Java命令
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
    
    ' 执行命令
    WScript.Echo "Generated command: " & javaCommand
    CreateObject("WScript.Shell").Run javaCommand, 1, False
    
    If Err.Number <> 0 Then
        MsgBox "启动Minecraft时出错: " & Err.Description, vbCritical, "错误"
    End If
End Sub

' 构建库路径
Function BuildLibrariesPath(mcPath, versionJson)
    On Error Resume Next
    
    Dim libs, lib, libPath, result
    
    ' 首先添加版本jar文件
    result = mcPath & "versions\" & versionJson("id") & "\" & versionJson("id") & ".jar"
    
    ' 添加所有库文件
    libs = versionJson("libraries")
    For Each lib In libs
        ' 检查规则(跳过不适用于当前系统的库)
        If Not CheckLibraryRules(lib) Then 
            ' In VBScript, we use Continue For (but it's not available)
            ' So we just skip to next iteration
        Else
            ' 获取库路径
            libPath = GetLibraryPath(mcPath, lib)
            
            ' 添加到结果
            If libPath <> "" Then
                If result <> "" Then result = result & ";"
                result = result & libPath
            End If
        End If
    Next
    
    BuildLibrariesPath = result
End Function

' 检查库规则
Function CheckLibraryRules(lib)
    On Error Resume Next
    
    ' 如果没有rules部分，则总是包含
    If Not lib.Exists("rules") Then
        CheckLibraryRules = True
        Exit Function
    End If
    
    Dim rule, osName, osArch
    
    ' 获取当前系统信息
    osArch = "x86_64" ' Assuming 64-bit for VBScript version
    osName = "windows"
    
    ' 检查所有规则
    For Each rule In lib("rules")
        ' 检查action
        If rule("action") = "allow" Then
            ' 如果没有os部分，则允许
            If Not rule.Exists("os") Then
                CheckLibraryRules = True
                Exit Function
            End If
            
            ' 检查os条件
            If rule("os")("name") = osName Then
                ' 如果有arch条件，检查arch
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
            ' 如果没有os部分，则不允许
            If Not rule.Exists("os") Then
                CheckLibraryRules = False
                Exit Function
            End If
            
            ' 检查os条件
            If rule("os")("name") = osName Then
                CheckLibraryRules = False
                Exit Function
            End If
        End If
    Next
    
    ' 默认允许
    CheckLibraryRules = True
End Function

' 获取库路径
Function GetLibraryPath(mcPath, lib)
    On Error Resume Next
    
    Dim parts, artifactPath, nativePath, i, classifier
    
    ' 解析库名称
    parts = Split(lib("name"), ":")
    
    ' 构建基本路径
    artifactPath = mcPath & "libraries\" & Replace(parts(0), ".", "\") & "\" & parts(1) & "\" & parts(2) & "\" & parts(1) & "-" & parts(2)
    
    ' 检查是否有natives
    If lib.Exists("natives") Then
        ' 获取windows平台的native分类器
        If lib("natives").Exists("windows") Then
            classifier = lib("natives")("windows")
            classifier = Replace(classifier, "${arch}", "64") ' Assuming 64-bit
            
            ' 构建native路径
            nativePath = artifactPath & "-" & classifier & ".jar"
            
            ' 检查文件是否存在
            If FileExists(nativePath) Then
                GetLibraryPath = nativePath
                Exit Function
            End If
        End If
    End If
    
    ' 如果没有natives或找不到native文件，使用普通jar
    artifactPath = artifactPath & ".jar"
    If FileExists(artifactPath) Then
        GetLibraryPath = artifactPath
        Exit Function
    End If
    
    GetLibraryPath = ""
End Function

' 构建游戏参数
Function BuildGameArguments(mcPath, versionName, playerName, versionJson)
    On Error Resume Next
    
    Dim args, assetsPath, versionType, arg
    
    ' 设置默认值
    assetsPath = mcPath & "assets"
    versionType = "WMML"
    
    ' 尝试从json获取assets和versionType
    If versionJson.Exists("assets") Then
        assetsPath = mcPath & "assets\"
    End If
    
    If versionJson.Exists("type") Then
        versionType = versionJson("type")
    End If
    
    ' 构建参数
    args = ""
    
    ' 检查是否有minecraftArguments(旧版本)
    If versionJson.Exists("minecraftArguments") Then
        args = versionJson("minecraftArguments") & " " & args
    End If
    
    ' 检查是否有arguments(新版本)
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

' 辅助函数: 读取文本文件
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

' 辅助函数: 检查文件是否存在
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