'################################################################################
'#                                                                              #
'#  Script para realizar atualiza��o de hotfix                                  #
'#  Autor: Leonardo La Rosa                                                     #
'#  L2R1 2018                                                                   #
'#  v0.01                                                                       #
'#                                                                              #
'#  - Consulta lista de Host (HostList.txt)- Possibilidade de consultar no SQL  #
'#  - Consulta lista de Path (Install.txt) - Possibilidade de consultar no SQL  #
'#  - Informa oa path onde est� os reposit�rios                                 #
'#  - Identifica o SO e busca o patch na pasta selecionada                      #
'#  - Realiza C�pia dos arquivos para a esta��o remota - SMB                    #
'#  - Realiza a instala��o remotamente (MSI e/ou EXE e/ou BAT)                  #
'#  - Grava log                                                                 #
'#  - Grava SQL                                                                 #
'#                                                                              #
'################################################################################

On Error Resume Next

Dim arq, strArquivo, strTexto, strTextFile, OS
Public gLog
arq = ""

Const ForReading = 1
Const ForWriting = 2

' Modo de usar - Ative para receber passagem de par�metro, do contr�rio, deixe desativado.
'If WScript.arguments.count <> 3 Then
'	WScript.echo "Utilize:" & WScript.scriptname & " <HostList.txt> <Install.txt> <LocalPathToPatches>" & vbCrLf & vbCrLf & _
'	"  <LocalPathToPatches> precisa ser informado o caminho completo" & vbCrLf & _
'	WScript.quit
'End If
'ipFile = WScript.arguments(0)
'strArquivo = WScript.arguments(1)
'localPathToPatches = WScript.arguments(2)

ipFile = "hostList.txt"
strArquivo = "install.txt"
localPathToPatches = "C:\Users\lrosa\Desktop\HotFix\repo"

Set onet = CreateObject("wscript.network")
Set ofs = CreateObject("scripting.filesystemobject")
Set FSO = CreateObject("scripting.filesystemobject")

' Valida arquivo de lista de Hostname.
Set oipFile = ofs.opentextfile(ipFile, 1, False)
If (Err.Number <> 0) Then
	'	WScript.echo "N�o foi poss�vel abrir o arquivo:" & ipFile
	gLog = "N�o foi poss�vel abrir o arquivo:" & ipFile
	Call gravaLog
	WScript.quit
End If

' Valida arquivo de lista de Patch.
Set ostrArquivo = ofs.opentextfile(strArquivo, 1, False)
If (Err.Number <> 0) Then
	'	WScript.echo "N�o foi poss�vel abrir o arquivo:" & strArquivo
	gLog = "N�o foi poss�vel abrir o arquivo:" & strArquivo
	Call gravaLog
	
	WScript.quit
End If


' Valida caminho de LocalPathToPatches, inserindo uma \ no final caso n�o haja.
If Right(localPathToPatches, 1) <> "\" Then
	localPathToPatches = localPathToPatches & "\"
End If

' O local de LocalPathToPatches precisa ser um reposit�rio local ou mapeado (n�o h� suporte para UNC path).
If Left(localPathToPatches, 2) = "\\" Then
	'	WScript.echo "<pathToExecutable> precisa ser um reposit�rio local ou mapeado localmente"
	gLog = "<pathToExecutable> precisa ser um reposit�rio local ou mapeado localmente"
	Call gravaLog
	WScript.quit
End If

Set osvcLocal = GetObject("winmgmts:root\cimv2")

'Verifica se os equipamentos na lista s�o v�lidos
Do While oipFile.atEndOfStream <> True
	ip = oipFile.ReadLine()
	'	WScript.echo vbCrLf & "Conectando a " & ip & "..."
	
	Err.Clear
	Set osvcRemote = GetObject("winmgmts:\\" & ip & "\root\cimv2")
	
	If (Err.Number <> 0) Then
		'		WScript.echo "Erro ao conectar-se a " & ip & "."
		gLog = "Erro ao conectar-se a " & ip & "."
		Call gravaLog
		
		
	Else
		
		Do While ostrArquivo.AtEndOfStream <> True
			exeCorrectPatch = ostrArquivo.ReadLine()
			
			
			'Identifica vers�o de build para aplicar patch espec�fico.			
			'Set oOSInfo = osvcRemote.InstancesOf("Win32_OperatingSystem")
			'For Each objOperatingSystem In oOSInfo
			' Define a vers�o do SO para atualiza��o adequada
			'	OS = objOperatingSystem.Version
			'Next
			
			localPathToPatches1  = localPathToPatches & "\" 
			ExecInstall = ofs.getfile(localPathToPatches1 + exeCorrectPatch).Name
			
			
			If (exeCorrectPatch <> "") Then
				'				WScript.echo "Intalando patch " & exeCorrectPatch & "..."
				onet.mapnetworkdrive "z:", "\\" & ip & "\C$"
				
				Set osourceFile = osvcLocal.Get("cim_datafile=""" & Replace(localPathToPatches1, "\", "\\") &exeCorrectPatch& """")
				'Copia os arquivos para o equipamento remoto (C:\)
				ret = osourceFile.Copy("z:\\temp\\"&exeCorrectPatch)
				
				
				If (ret <> 0 And ret <> 10) Then
					'Em caso de erro ao copiar o arquivo localmente:
					'					WScript.echo "Erro ao copiar para: " & ip & " - C�digo de erro: " & ret
					gLog = "Erro ao copiar para: " & ip & " - C�digo de erro: " & ret
					Call gravaLog
					
				Else
					'Do contr�rio, a instala��o continua.
					Set oprocess = osvcRemote.Get("win32_process")
					
					'Valida a extens�o do arquivo para realizar a instal���o silenciosa (Necess�rio definir os par�metros do MSI e EXE)
					
					strTextFile = Split(exeCorrectPatch,".")
					If strTextFile(1) = "msi" Then
						ret = oprocess.create("msiexec.exe /i c:\\temp\\"&exeCorrectPatch&" /qn")
					ElseIf strTextFile(1) = "bat" Then
						ret = oprocess.create("c:\\temp\\"&exeCorrectPatch)
					Else
					'presume que � execut�vel
						ret = oprocess.create("c:\\temp\\"&exeCorrectPatch&" -q")
					End If						
				End If						
				
				'Em caso de erro, retorna mensagem					
				If (ret <> 0) Then
					'WScript.echo "Erro ao iniciar o processo de instala��o em: " & ip & ": " & ret
					gLog = "Erro ao iniciar o processo de instala��o em: " & ip & ": " & ret
					Call gravaLog						
				Else
					Set odestFile = osvcLocal.Get("cim_datafile=""z:\\temp\\"&exeCorrectPatch&"""")
					
					'Se a instala��o for n�o assistida, aguarda o t�rmino
					For waitTime = 0 To 120     
						WScript.Sleep 1000 
						'Assim que o temp for liberado, o mesmo � deletado							
						If (odestFile.Delete() = 0) Then
							Exit For
						End If
					Next
					
					'WScript.echo "Installation successful."
					gLog = "Installation successful."
					Call gravaLog
				End If     
				'End If     
				onet.removenetworkdrive "z:", True
				'End If      
			End If 
			
		Loop
		
	End If 
	
Loop

oipFile.Close()
ostrArquivo.Close()

Sub gravaLog
	
	' inicia a grava��o do LOG
	Set WshNetwork = CreateObject("Wscript.Network")
	strUserName = WshNetwork.UserName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile("app.log", ForReading)
	strContents = objFile.ReadAll
	objFile.Close
	strFirstLine = Now&" - "&gLog
	strNewContents = strFirstLine & vbCrLf & strContents
	Set objFile = objFSO.OpenTextFile("app.log", ForWriting)
	objFile.WriteLine strNewContents 
End Sub

' remove o mapeamento de rede
If ofs.folderexists("z:\") Then
	onet.removenetworkdrive "z:", True
End If


'Grava��o no SQL (validada) 
'Set conn = CreateObject("ADODB.Connection")
'strConnection = "Provider=SQLOLEDB.1;Data Source=br001lab106;User ID=international\XXXXXX; Password=XXXXXXXXXXXXX;Initial Catalog=HotFix;Trusted_Connection=yes;"
'Set conn = CreateObject("ADODB.Connection")
'conn.Open strConnection
'query = "Select * from HotFix"
'query = "INSERT INTO HotFix (Name, Date) VALUES ('teste2', '2018/03/16 00:00')"
'Set rs = conn.Execute(query)
'dbResults = rs.GetString 
'WScript.echo dbResults
