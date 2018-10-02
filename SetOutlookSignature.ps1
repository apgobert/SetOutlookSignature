# ------------------------------------------------------------------------------------------------------------------------------------------------------------------- #
# • Author : Alexis Plassmann Gobert - Stagiaire														     														  #
# • Stäubli Company - 2014																	      																	  #
# • Usage : .\SetOutlookSignature [signature_name][default_mode]												      												  #
# • Config : PowerShell 1.0, Outlook 2010, Exchange 2013													      													  #
#																				      																				  #
# • Ce script PowerShell sera placé sur un Share, dans le dossier Netlogon.                                                                                           #
# • Un Dossier "mailsig" doit être placé dans Netlogon : il contiendra toutes les signatures (au format ".docx") voulant être mises en place.                         #
# • Le script permet de créer 3 formats de fichiers à partir du ".docx" hébergé sur le Share. Ces 3 formats sont générés en local dans "AppData\Microsoft\Signatures".#
# • Les 3 formats de fichiers constituent la signature Outlook. La signature n'est push que si la version locale est différente de la version serveur au lancement.   #
# • Un fichier de logs est créé dans "AppData\Roaming\SetOutlookSignatureLogs" : toutes les erreurs rencontrées à l'exécution du script y sont stockées.              # 
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------- #

# --------------------------------------------------------------- #
#                          "Globales"                             #
# --------------------------------------------------------------- #

$AppData = (Get-Item ENV:APPDATA).value
$LogsFile = "SetOutlookSignatureLogs.txt"
$LogsDir = $AppData + "\SetOutlookSignatureLogs"
$LogsPath = $LogsDir + '\' + $LogsFile

# --------------------------------------------------------------- #
#                          FUNCTIONS                              #
# --------------------------------------------------------------- #

### Ecrire dans le fichier de logs ###
Function WriteLogs ([String]$Logs)
{
	$Date = Get-Date -Format d
	$Time = Get-Date -Format T
	"$Date | $Time | $ENV:USERNAME | $ENV:COMPUTERNAME | $Logs" >> $LogsPath
}

### Affiche un message à l'écran ###
Function DisplayMessage ([String]$Message, [String]$Color)
{
	Write-Host $Message -ForegroundColor $Color
}

### Initialisation de l'application Word ###
Function InitApp ($Type)
{
	$App = New-Object -ComObject $Type
	$App.Visible = $False
	Return $App
}

### Quitter l'application Word ###
Function QuitApp ($App)
{
	$App.quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) | Out-Null
	Remove-Variable -Name App
	[gc]::collect()
	[gc]::WaitForPendingFinalizers()
}

### Ouvrir un document word, visibilité désactivée ###
Function OpenWordDocument ($App, $File)
{
	$GetLocalFile = Get-childitem $File
	$Doc = $App.documents.open($GetLocalFile.fullname)
	Return $Doc
}

### Fermer le document word ###
Function CloseDocument ($Document, $BuiltinProperties, [ref]$SaveOption)
{
	If ($SaveOption -eq $Null)
	{
		$Document.close()
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($BuiltinProperties) | Out-Null
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
		Remove-Variable -Name Document, BuiltinProperties
	}
	Else
	{
		$Document.close([ref]$SaveOption::wdDoNotSaveChanges)
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($BuiltinProperties) | Out-Null
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
		Remove-Variable -Name Document, BuiltinProperties
	}
}

### Outlook est-il installé ? ###
Function CheckOutlookInstall
{
	$OutlookLocation = ((Get-ItemProperty -Path HKLM:\SOFTWARE\Classes\Outlook.File.Pst.14\shell\Open\command).'(default)'-split '"')[1]
	If (!(Get-ChildItem $OutlookLocation))
	{
		WriteLogs "[Outlook 2010 location error] : `"$OutlookLocation`" not found"
		DisplayMessage "[Outlook 2010 location error] : `"$OutlookLocation`" not found" "Yellow"
		Exit
	}
}

### Obtenir le nom du profil outlook###
Function GetOutlookProfileName
{
	$ProfilePath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
	If (!(Test-Path $ProfilePath))
	{
		WriteLogs "[Outlook Profile Path] : `"$ProfilePath`" not found"
		DisplayMessage "[Outlook 2010 location error] : `"$ProfilePath`" not found" "Yellow"
		Exit
	}
	Else
	{
		$ProfileKey = "DefaultProfile"
		Try
		{
			$DefaultProfile = (Get-ItemProperty $ProfilePath).DefaultProfile
		}
		Catch [system.exception]
		{
			WriteLogs "[Outlook Default Profile] : `"$ProfileKey`" doesn't exist"
			DisplayMessage "[Outlook Default Profile] : `"$ProfileKey`" doesn't exist" "Yellow"
			Exit
		}
		Return $DefaultProfile
	}
}

### Verifier que Outlook est configuré ###
Function CheckOutlookConfig
{
	$OutlookProfileName = GetOutlookProfileName
	If ($OutlookProfileName -eq $Null)
	{
		WriteLogs "[Outlook User Profile] : Not configured"
		DisplayMessage "[Outlook User Profile] : Not configured" "Yellow"
		Exit
	}
}

### Création du dossier et du fichier de logs ###
Function CreateLogs
{
	If (!(Test-Path $LogsPath))
	{
		New-Item $LogsDir -ItemType Directory
		New-Item $LogsPath -ItemType File
	}
}

### Créer une clé de registre ###
Function CreateRegistryKey ($SignatureRegPath, $SignatureName)
{
	If (!(Test-Path $SignatureRegPath))
	{
		New-Item -Path "HKCU:\Software" -Name $SignatureName
	}

	If (!(Test-Path $SignatureRegPath"\Outlook Signature Settings"))
	{
		New-Item -Path $SignatureRegPath -Name "Outlook Signature Settings"
	}
}

### Obtenir la valeur d'une clé de registre ###
Function GetRegistryValue ($SignatureRegPath, $Value)
{
	Try
	{
		(Get-ItemProperty $SignatureRegPath"\Outlook Signature Settings").$Value
	}
	Catch [system.exception]
	{
		WriteLogs "[Get-ItemProperty] : Impossible to get $Value value for $SignatureRegPath"
		DisplayMessage "[Get-ItemProperty] : Impossible to get $Value value for $SignatureRegPath" "Yellow"
	}
}

### Obtenir le contenu du mot clé Tags (Propriétés > Détails) d'un .docx ###
Function GetSignatureVersion ($Keywords, $ServerSignaturePath)
{
	$App = InitApp "word.application"
	[ref]$SaveOption = "microsoft.office.interop.word.WdSaveOptions" -as [type]
	$Document = OpenWordDocument $App $ServerSignaturePath
	$BuiltinProperties = $Document.BuiltInDocumentProperties
	$Binding = "System.Reflection.BindingFlags" -as [type]
	$PropertiesType = $BuiltinProperties.GetType()
	Try
	{
		$BuiltInProperty = $PropertiesType.invokemember("item",$Binding::GetProperty,$Null,$BuiltinProperties,$Keywords)
		$BuiltInPropertyType = $BuiltInProperty.GetType()
		$Version = $BuiltInPropertyType.invokemember("value",$Binding::GetProperty,$Null,$BuiltInProperty,$Null)
	}
	Catch [system.exception]
	{
		WriteLogs "[GetKeywordsDocument] : Unable to get value for $ServerSignaturePath"
		DisplayMessage "[GetKeywordsDocument] : Unable to get value for $ServerSignaturePath" "Yellow"
		CloseDocument $Document $BuiltinProperties $SaveOption
		QuitApp $App
		Exit
	}
	CloseDocument $Document $BuiltinProperties $SaveOption
	QuitApp $App
	Return $Version
}

### Set le contenu du mot clé Tags (Propriétés > Détails >) d'un .docx ###
Function SetSignatureVersion ($Keywords, [String]$StockedVersion, $Path)
{
	DisplayMessage "Setting Document Version" "Green"
	$App = InitApp "word.application"
	$Binding = "System.Reflection.BindingFlags" -as [type]
	$Document = OpenWordDocument $App $Path
	$BuiltinProperties = $Document.BuiltInDocumentProperties
	$PropertiesType = $BuiltinProperties.GetType()
	Try
	{
		$BuiltInProperty = $PropertiesType.invokemember("item",$Binding::GetProperty,$Null,$BuiltinProperties, $Keywords)
		$BuiltInPropertyType = $BuiltInProperty.GetType()
		$BuiltInPropertyType.invokemember("value",$Binding::SetProperty,$Null,$BuiltInProperty,$StockedVersion)
	}
	Catch [system.exception]
	{
		WriteLogs "[SetKeywordsValue] : Unable to set value for $Keywords for $Path"
		DisplayMessage "[SetKeywordsValue] : Unable to set value for $Keywords for $Path" "Yellow"
		CloseDocument $Document $BuiltinProperties $Null
		QuitApp $App
		Exit
	}
	CloseDocument $Document $BuiltinProperties $Null
	QuitApp $App
}

### Sauvegarder en local un format de fichier (ex : .rtf, .htm, .txt) ###
Function SaveFormats ($OutlookSignaturePath, $SignatureName, [String]$Extension, [String]$Format, $MsWord)
{
	$SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], $Format);
	$Path = $OutlookSignaturePath+'\'+$SignatureName+$Extension
	$MsWord.ActiveDocument.saveas([ref]$Path, [ref]$SaveFormat)
}

### Création des 3 formats Outlook ###
Function CreateOutlookFormats ($MsWord, $OutlookSignaturePath, $SignatureName)
{
	SaveFormats $OutlookSignaturePath $SignatureName ".htm" "wdFormatHTML" $MsWord
	SaveFormats $OutlookSignaturePath $SignatureName ".rtf" "wdFormatRTF" $MsWord
	SaveFormats $OutlookSignaturePath $SignatureName ".txt" "wdFormatText" $MsWord
	$MsWord.ActiveDocument.Close()
	$MsWord.Quit()
}

### Remplacer la variable de l'AD par son contenu ###
Function InsertVariable ([String]$TextToReplace, [String]$ReplaceWith, $WordObject)
{
	[void]$WordObject.Selection.Find.Execute($TextToReplace, $False, $True, $False, $False, $False, $True, 1, $False, $ReplaceWith, 2)
}

### Listing des variables de l'AD voulues ###
Function InsertAllVariables ($OutlookSignaturePath, $SignatureName, $LocalSignaturePath)
{
	### Obtenir les valeurs de l'AD pour le current user ###
	$UserName = $ENV:USERNAME
	$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
	$Searcher = New-Object System.DirectoryServices.DirectorySearcher
	$Searcher.Filter = $Filter
	$ADUserPath = $Searcher.FindOne()
	$ADUser = $ADUserPath.GetDirectoryEntry()
	$ADEmailAddress = $ADUser.mail
	
	### Insertion variables de l'AD dans .rtf ###
	$MsWord = InitApp "word.application"
	Try
	{
		[void]$MsWord.Documents.Open($LocalSignaturePath)
	}
	Catch [system.exception]
	{
		WriteLogs "[Open Word Document] : Impossible to open $LocalSignaturePath"
		DisplayMessage "[Open Word Document] : Impossible to open $LocalSignaturePath" "Yellow"
	}
	InsertVariable "<DisplayName>" $ADUser.DisplayName.ToString() $MsWord
	InsertVariable "<Department>" $ADUser.Department.ToString() $MsWord
	InsertVariable "<Telephone>" $ADUser.telephoneNumber.ToString() $MsWord
	InsertVariable "<FirstName>" $ADUser.FirstName.ToString() $MsWord
	InsertVariable "<LastName>" $ADUser.LastName.ToString() $MsWord
	InsertVariable "<StreetAddress>" $ADUser.streetAddress.ToString() $MsWord
	InsertVariable "<City>" $ADUser.l.ToString() $MsWord
	InsertVariable "<CountryName>" $ADUser.c.ToString() $MsWord
	InsertVariable "<PostalCode>" $ADUser.postalCode $MsWord
	InsertVariable "<CompanyName>" $ADUser.company $MsWord
	InsertVariable "<MobilePhone>" $ADUser.mobile $MsWord
	InsertVariable "<Title>" $ADUser.title $MsWord
		
	If ($MsWord.Selection.Find.Execute("<Email>"))
	{
		[void]$MsWord.ActiveDocument.Hyperlinks.Add($MsWord.Selection.Range, "mailto:"+$ADEmailAddress.ToString(), $Missing, $Missing, $ADEmailAddress.ToString())
	}
	CreateOutlookFormats $MsWord $OutlookSignaturePath $SignatureName
}

### Signature opérationnelle dans Outlook ###
Function ChangeDefaultMod ($Type, $SignatureName)
{
    $MsWord = New-Object -ComObject "word.application"
    $EmailOptions = $MsWord.emailoptions
    $EmailSignature = $EmailOptions.emailsignature
    $EmailSignatureEntries = $EmailSignature.emailsignatureentries
    $EmailSignature.$Type = $SignatureName
    $MsWord.quit()
}

### Stockage des préférences de la signature dans les registres ###
Function SetRegistryValue ([int]$DefaultMod, $SignatureName, $SignatureRegPath, $ForcedSignatureNew)
{
	If ($DefaultMod -eq 1)
	{
		Set-ItemProperty $SignatureRegPath"\Outlook Signature Settings" -Name ForcedSignatureNew -Value 1
		Set-ItemProperty $SignatureRegPath"\Outlook Signature Settings" -Name ForcedSignatureReplyForward -Value 1
	}
	ElseIf ($DefaultMod -eq 0)
	{
		Set-ItemProperty $SignatureRegPath"\Outlook Signature Settings" -Name ForcedSignatureNew -Value 0
		Set-ItemProperty $SignatureRegPath"\Outlook Signature Settings" -Name ForcedSignatureReplyForward -Value 0
	}
}

Function CheckOutlookFiles ($OutlookSigPath, $SigName)
{
	$DocxFormat = $SigName + ".docx"
	$RtfFormat = $SigName + ".rtf"
	$HtmFormat = $SigName + ".htm"
	$TxtFormat = $SigName + ".txt"
	If ((Test-Path $OutlookSigPath\$DocxFormat) -and (Test-Path $OutlookSigPath\$RtfFormat) -and (Test-Path $OutlookSigPath\$HtmFormat) -and (Test-Path $OutlookSigPath\$TxtFormat))
	{
		Return $True
	}
	Else
	{
		Return $False
	}
}

# --------------------------------------------------------------- #
#                          MAIN                                   #
# --------------------------------------------------------------- #

CreateLogs
CheckOutlookInstall
CheckOutlookConfig

### Gestion des arguments ###
If ($args.Length -eq 2)
{
	[String]$FirstArg = $args[0]
	
	### Parsing nom firme + SignatureName ###
	If ($FirstArg.Contains("\"))
	{
		$Firm = $FirstArg.split("\")[0]
		$SignatureName = $FirstArg.split("\")[1]
		### $SigSource = Emplacement signatures serveur ###
		[String]$SigSource = "$ENV:LOGONSERVER\NETLOGON\mailsig\" + $Firm
	}
	Else
	{
		$SignatureName = $FirstArg
		[String]$SigSource = "$ENV:LOGONSERVER\NETLOGON\mailsig"
	}
	
	If (!(Test-Path $SigSource))
    {
        WriteLogs "[Wrong Path] : $SigSource doesn't exist"
		DisplayMessage "[Wrong Path] : $SigSource doesn't exist" "Yellow"
    }
	
	### Premier argument = default mode ###
	If ($args[1] -eq 1)
	{	
		$DefaultMod = 1
	}
	ElseIf ($args[1] -eq 0)
	{
		$DefaultMod = 0
	}
	Else
	{
        WriteLogs "[default_mode] : Second parameter, '1' to enable or '0' to disable default_mod"
		DisplayMessage "[default_mode] : Second parameter, '1' to enable or '0' to disable default_mod" "Yellow"
		Exit
	}
	
	[String]$Signature = $SignatureName + ".docx"
	[String]$ServerSignaturePath = $SigSource + '\' + $Signature
	If (!(Get-ChildItem $ServerSignaturePath))
    {
		WriteLogs "[Signature location error] : $SignatureName.docx doesn't exist in `"$SigSource`""
        DisplayMessage "[Signature location error] : $SignatureName.docx doesn't exist in `"$SigSource`"" "Yellow"
        Exit
    }

	[array]$Keywords = "Keywords"
	### Obtenir version de la signature serveur ###
	[String]$StockedVersion = GetSignatureVersion $Keywords $ServerSignaturePath
	If ($StockedVersion -eq "")
	{
		WriteLogs "[Server signature version] : not set, please update it"
        DisplayMessage "[Server signature version] : not set, please update it" "Yellow"
        Exit
	}

	$SigPath = "\Microsoft\Signatures"
	$OutlookSignaturePath = $AppData+$SigPath
	[String]$LocalSignaturePath = $OutlookSignaturePath + '\' + $Signature
    If (Test-Path $LocalSignaturePath)
    {
		### Obtenir la version de la signature locale ###
		[String]$LocalSignatureVersion = GetSignatureVersion $Keywords $LocalSignaturePath
    }
	
    ### Comparaison n°version ###
	### Si n° version stocké est différent du n° version de la signature ou plus de fichiers outlook -> Push Signatures ###
    If (($StockedVersion -ne $LocalSignatureVersion) -or ((CheckOutlookFiles $OutlookSignaturePath $SignatureName) -eq $False))
    {
		WriteLogs "Pushing signature [loading ...]"
		DisplayMessage "Pushing signature [loading ...]" "Green"
		Try
		{
		    ### Copie signature .docx serveur vers local ###
			If (Test-Path $OutlookSignaturePath)
			{
				Copy-Item "$SigSource\$Signature" $OutlookSignaturePath -Force
			}
			Else
			{
				New-Item $OutlookSignaturePath -ItemType Directory
				Copy-Item "$SigSource\$Signature" $OutlookSignaturePath -Force
			}
		}
		Catch [system.exception]
		{
			WriteLogs "[Copy Item] : Impossible to copy $Signature in $OutlookSignaturePath"
			DisplayMessage "[Copy Item] : Impossible to copy $Signature in $OutlookSignaturePath" "Yellow"
			Exit
		}
			
		### Variables AD + création 3 formats  : .rtf, .htm, .txt ###
		InsertAllVariables $OutlookSignaturePath $SignatureName $LocalSignaturePath

		### Set valeurs registre pour stockage préférences signatures ###
		[String]$SignatureRegPath = "HKCU:\Software\"+$SignatureName
		CreateRegistryKey $SignatureRegPath $SignatureName
		$ForcedSignatureNew = GetRegistryValue $SignatureRegPath "ForcedSignatureNew"
		$ForcedSignatureReplyForward = GetRegistryValue $SignatureRegPath "ForcedSignatureReplyForward"
		SetRegistryValue $DefaultMod $SignatureName $SignatureRegPath
		
		### Unset default mode si on passe la même signature d'un mode par défaut à n'étant pas par défaut ###
		If (((Test-Path $SignatureRegPath) -and ($ForcedSignatureNew -eq 1)) -and ($DefaultMod -eq 0))
		{
			WriteLogs "Unsetting `"$SignatureName`" default mode"
			DisplayMessage "Unsetting `"$SignatureName`" default mode" "Green"
			ChangeDefaultMod "NewMessageSignature" $null
			ChangeDefaultMod "ReplyMessageSignature" $null
		}
		
		### Set DefaultMod  pour les nouveaux messages  + réponses/transferts ###
		If ($DefaultMod -eq 1)
		{
			WriteLogs "Setting `"$SignatureName`" to default signature"
			DisplayMessage "Setting `"$SignatureName`" to default signature" "Green"
		    ChangeDefaultMod "NewMessageSignature" $SignatureName
			ChangeDefaultMod "ReplyMessageSignature" $SignatureName
		}
		
		### Ajout signature version local ###
		SetSignatureVersion $Keywords $StockedVersion $LocalSignaturePath
		WriteLogs "[End] : New signature available"
		DisplayMessage "[End] : New signature available" "Green"
    }
	### Si n° version stocké = n° version signature locale -> signature déjà en place ###
    ElseIf ($StockedVersion -eq $LocalSignatureVersion)
    {
		WriteLogs "[Signature Version] : already up to date"
		DisplayMessage "[Signature Version] : already up to date" "Green"
		Exit
    }
}
Else
{
	WriteLogs "[Usage] : .\script_name [folder\signature_name] [default_mode]"
	DisplayMessage "[Usage] : .\script_name [folder\signature_name] [default_mode]" "Yellow"
	Exit
}