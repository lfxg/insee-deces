# Rechercher une personne dans les fichiers des Deces locaux 
#  depuis l'annee fournie jusqu au dernier mois de l'an en cours
# fourni en parametre
#  1) -anneeStr de debut de la recherche ex :2020
#  2) -nom de naissance de la personne 
#  3) -prenom de de la personne
#  4) -exact option Boolean valeur defaut False
#  5) -autre option String valeur annee naissance ou code isee ville de naissance ou deces
#
#   Si un fichier est absent on le telecharge
#   Comme il y a des sosies on continue meme apres un succes 
#   pour cela la variable present reste sur false apres un succes
#
# Pour le  lancement
#    cd  $Home\OneDrive\Documents\Batch\Viager\
#    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
#    .\RechercheDeces-FichiersLocaux.ps1 2019 RONDA Juliette
#			Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope Process -Force
# 
param(
	[Parameter(Mandatory=$false)][Switch] $exact,
	[Parameter(Mandatory=$true)][String] $anneeStr,
	[Parameter(Mandatory=$true)][String] $prenom,
	[Parameter(Mandatory=$true)][String] $nom,
	[Parameter(Mandatory=$false)][String] $autre
)
BEGIN {
	$app = "RD4"
	$inseeAn0 = 1970
	$nom = $nom.ToUpper()
	$prenom = $prenom.ToUpper()
	if ($exact) {
		# les noms commencent par une guillemet " finissent par *
		# les noms cherchés sont egaux strictement aux noms du fichier
		$regName = '\"'+ "$nom\*$prenom"
	} else {
		# les noms cherchés sont inclus dans les noms du fichier
		$regName = '\"'+ "$nom[A-Z]*\*$prenom"
	}
	if ($autre.Length -ge 1 ) {
		$regname = $regname + '.*'+ $autre
	}
	$DestDir  = "$Home\OneDrive\OffLine\Administratif\Family\PascalZuch\Viager"
	$Log = "$DestDir/Recherches_deces_"+$app +".log"
	If ( $PSBoundParameters.Keys.Count -ge 3) {
		if ( -not ($anneeStr -match "^\d+$")) { 
		myLog "$Date Attention le premier parametre obligatoire est l annee de départ, fin procedure " 
		exit
		}
		$anDemEnt =  [convert]::ToInt32($anneeStr)
		if ( $anDemEnt  -lt $inseeAn0 ) { 
			myLog "$Date Archives INSEE commence en $inseeAn0, fin procedure " 
			exit
		}
	} else {
		myLog "$Date Attention 3 parametres obligatoire: annee de départ , nom, prenom , fin procedure " 
		exit
	}	
	# Preparation des variables indispensables
	# Nouveau Nom du fichier de Log
	$Log = "$DestDir/Recherches_deces_"+$app+"_"+$anneeStr+"_"+($prenom-replace '\\|\*','')+"_"+($nom -replace '\\|\*','') +".log"
	$Date = $((Get-Date).ToString("yyyyMMdd-HHmm"))
	# Internet Source file location
	# https://www.insee.fr/fr/statistiques/fichier/4190491/Deces_2022_M09.zip
	# $inseeDecActuelle = 'https://www.insee.fr/fr/statistiques/fichier/4190491'
	# $inseeDecPassee = 'https://www.insee.fr/fr/statistiques/fichier/4769950'
	# Source file location
	# Deces_AAAA_Mmm.zip exemple Deces_2022_M09.zip 
	#  Fichiers |
	$ext = ".zip"
	$ext2 = ".csv"
	# Fonction Log activity
	function myLog{
		[CmdletBinding()]
		param(
			[Parameter()]
			[String] $Text
		)
		write-output "$Text" | out-file -FilePath $Log -Encoding UTF8 -Append
	}
	#
	# Function  prepareMidNames
	#  MidNameZip MUST exist before function is launched
	#  prepareMidNames 1989 -MidZip ([Ref]$MidNameZip) 
	function  prepareMidNames{
		[CmdletBinding()]
		param(
			[Parameter()]
			[Int32] $annee,
			[Ref]$MidZip
		)
		# Calcul decennie date demandee
		$decDemEnt=([Math]::Floor($annee/10))*10
		#
		# On part de l'hypothèse que l'INSEE a un mois de retard / date actuelle
		$dateMAJ =  $(Get-Date).AddMonths(-1)
		# Dans notre contexte Actuel veut dir Actuellement disponible sur INSEE
		$moisActuel = $dateMAJ.ToString("MM")
		$anActuel = $dateMAJ.ToString("yyyy")
		# Calcul decennie en cours
		$decAct=([Math]::Floor($anActuel/10))*10
		#
		# Attention si la date MAJ est le mois de décembre 
		# Nous sommes janvier année N
		# La compilation de l'an precedent n'est pas faite
		# les fichiers mensuels sont renommés en ajoutant _annee avant le mois
		if (($annee -eq $anActuel) -and ($moisActuel -eq "12")) {
			$MidZip.Value = "$annee"
		}else {
			switch ($decDemEnt){
				{($PSItem -lt [Int32]$anActuel) -and ($PSItem -ge [Int32]$decAct) }
					{
					$MidZip.Value = "$annee" 
					}
				{$PSItem -lt [Int32]$decAct }
					{
					$MidZip.Value = "$decDemEnt" + "-" + "$($decDemEnt+9)" +"-csv"
					}
			}
		}	
	}
	function  preparePrefixes{
		[CmdletBinding()]
		param(
			[Parameter()]
			[Int32] $decennie,
			[Ref]$PrfxZip,
			[Ref]$PrfxCsv
		)
		#myLog "PrefixZip = $PrefixZip"
		$PrefixZip2 = "Deces_"  # Depuis 2020 
		$PrefixZip1 = "deces-"  # Avant 2020
		$PrefixCSV2 = "Deces_"   # Depuis 2010 
		$PrefixCSV1 = "deces-"   # Avant 2010 
		#
		$anBasculeZip = 2020
		$anBasculeCsv = 2010
		#
		switch ($decennie){
			{$PSItem -ge $anBasculeZip }
				{
				$PrfxZip.Value  =  $PrefixZip2
				$PrfxCsv.Value  =  $PrefixCsv2
				#myLog " PrefixZip = PrefixZip2 et PrefixCsv = PrefixCsv2"
				}
			{($PSItem -ge $anBasculeCsv) -And ($PSItem -lt $anBasculeZip)  }
				{
				$PrfxZip.Value  =  $PrefixZip1
				$PrfxCsv.Value  =  $PrefixCsv2
				#myLog " PrefixZip = PrefixZip1 et PrefixCsv = PrefixCsv2"
				}
			{$PSItem -lt $anBasculeCsv }
				{
				$PrfxZip.Value  =  $PrefixZip1
				$PrfxCsv.Value  =  $PrefixCsv1
				#myLog " PrefixZip = PrefixZip1 et PrefixCsv = PrefixCsv1"
				}
		}
	}
	# end Function prepare-Prefixes
	#
	### function  Prepare-url
	#  $insee string MUST exist before function is launched
	#  Prepare-url 1989 -url ([Ref]$insee) 
	Function  prepareUrl {
		[CmdletBinding()]
		param(
			[Parameter()]
			[Int32] $annee,
			[Ref] $url
		)
		# Calcul decennie date demandee
		$decDemEnt=([Math]::Floor($annee/10))*10
		#
		$anActuel = $((Get-Date).ToString("yyyy"))
		# Calcul decennie en cours
		$decAct=([Math]::Floor($anActuel/10))*10
		#
		switch ($decDemEnt){
			{$PSItem -eq $decAct }
				{
				$url.Value = 'https://www.insee.fr/fr/statistiques/fichier/4190491/'
				#myLog "url insee  décénnie en cours"
				} 
			{$PSItem -lt $decAct }
				{
				$url.Value = 'https://www.insee.fr/fr/statistiques/fichier/4769950/'
				#myLog "url insee  decennies passees"
				}
		}
	}
	# end Function prepare-url
	#
}
PROCESS {
	# ****************************** MAIN ********************************
	#
	# Calcul decennie année de début de recherche demandee
	$decDemEnt=([Math]::Floor($anneeStr/10))*10
	#
	# Calcul de la plus récente mise à jour
	# On part de l'hypothèse que l'INSEE a un mois de retard / date actuelle
	$dateMAJ =  $(Get-Date).AddMonths(-1)
	# Dans notre contexte Actuel veut dir ACtuellement disponible sur INSEE
	$moisActuel = $dateMAJ.ToString("MM")
	$anActuel = $dateMAJ.ToString("yyyy")
	# Calcul decennie en cours
	$decAct=([Math]::Floor($anActuel/10))*10
	# Il faut attendre le mois N° 2 pour que l'INSEE complie l'an precedent
	#
	$present = $false
	$PrefixZip = ""
	$PrefixCsv = ""
	$MidNameZip =""
	#$MidNameCsv = ""
	$insee = ""
	#
	if ( $anDemEnt -lt $anActuel ) {
		# recherche dans les années anterieures
		$indexAn = $anDemEnt
		#$indexDec = $decDemEnt
		while ( -not ($present) -and ($indexAn -lt $anActuel) ) {
			# test si le fichier csv existe localement
			preparePrefixes $indexAn -PrfxZip ([Ref]$PrefixZip) -PrfxCsv ([Ref]$PrefixCsv)
			$fileCsv = $PrefixCsv + $indexAn
			if (-not (Test-Path -Path $DestDir\$fileCsv$ext2 -PathType Leaf) ) {
				# Test si le fichier zip existe : un Zip par decennie
				prepareMidNames $indexAn -MidZip ([Ref]$MidNameZip)
				$fileZip = $PrefixZip + $MidNameZip
				if (-not (Test-Path -Path $DestDir\$fileZip$ext -PathType Leaf) ) {
						# sinon Download the zip file
						prepareUrl $indexAn  -url ([Ref]$insee) 
						$source = "$insee/$fileZip$ext"
						try
						{
							$Response = Invoke-WebRequest -Uri $source -OutFile $DestDir\$fileZip$ext
							# This will only execute if the Invoke-WebRequest is successful.
							myLog "$Date Le serveur connait le fichier $fileZip$ext nous avons une copie locale maintenant" 
							$StatusCode = $Response.StatusCode
						} catch {
							myLog "$Date Le serveur ne connait pas le fichier $fileZip$ext, fin de procedure" 
							$StatusCode = $_.Exception.Response.StatusCode.value__
							exit
						}
				}
				# Le fichier zip est ici > decompression desarchivage
				try
				{
					$Response = Expand-Archive -Path $DestDir\$fileZip$ext -DestinationPath $DestDir
					# This will only execute if the Expand-Archive is successful.
					myLog "$Date decompression archive du fichier $fileZip$ext OK" 
					$StatusCode = $Response.StatusCode
				} catch {
					myLog "$Date Erreur decompression archive du fichier $fileZip, fin de procedure" 
					myLog "$Date Code erreur $StatusCode = $_.Exception.Response.StatusCode.value__" 
					exit
				}
				# On Renomme les fichiers Deces car certains noms sont trop longs
				#  <Deces_1986(version rectifi‚e).csv>
				#  Je recupere dans $list les noms des fichiers de la decennie telechargee
				#
				$head = $PrefixCsv + $([String]$decDemEnt).substring(0,3)
				$list = Get-ChildItem -Path $DestDir -Name -Include "$head*$ext2" -Exclude "*.zip"
				#
				$expectedLen = $PrefixCsv.Length + ([String]$decDemEnt).Length + $ext2.Length
				foreach ($item in $list) {
						# myLog $item $item.Length $expectedLen
					if ( $item.Length -gt $expectedLen  ) {
						Rename-Item $DestDir/$item ($($item.substring(0,$expectedLen-$ext2.Length))+$ext2)
					}
				}
			}
			# on recherche le prenom et le nom
			$Search = Select-String -Path $DestDir\$fileCsv$ext2 -pattern $regName
			if ([string]::IsNullOrEmpty($Search)) {
				myLog "$Date Pas de $prenom $nom dans $fileCsv$ext2 " 
			} else { 
				$present = $false # on continue meme apres un succes  
				myLog "$Date $prenom $nom est present dans $fileCsv$ext2 ligne $((($Search -split (';'))[0] -split (':'))[2])" 
				Foreach ($line in $Search) {
					myLog "$Date Identite : $((($line -split (';'))[0] -split (':'))[3]) , naissance: $(($line -split (';'))[2]), deces $(($line -split (';'))[6]) " 		
				}
			}
			$indexAn++
		}
	} 
	$insee = ""
	$indexMois = 1
	while (	-not ($present) -and ($indexMois -le ([convert]::ToInt32($moisActuel)))) {
		preparePrefixes $anActuel -PrfxZip ([Ref]$PrefixZip) -PrfxCsv ([Ref]$PrefixCsv)
		$fileMoisCsv = $PrefixCSV + $anActuel+ "_M"+([String]$indexMois).PadLeft($moisActuel.Length,'0')
		# test si le fichier csv existe
		if (-not (Test-Path -Path $DestDir\$fileMoisCsv$ext2 -PathType Leaf) ) {
			# Test si le fichier zip existe
			$fileMoisZip = $PrefixZip + $anActuel+ "_M"+([String]$indexMois).PadLeft($moisActuel.Length,'0')
			if (-not (Test-Path -Path $DestDir\$fileMoisZip$ext -PathType Leaf) ) {
				# sinon Download the file
				prepareUrl $indexAn  -url ([Ref]$insee) 
				$source = "$insee/$fileMoisZip$ext"
				try
				{
					$Response = Invoke-WebRequest -Uri $source -OutFile $DestDir\$fileMoisZip$ext
					# This will only execute if the Invoke-WebRequest is successful.
					myLog "$Date Le serveur connait le fichier $fileMoisZip$ext nous avons une copie locale maintenant" 
					$StatusCode = $Response.StatusCode
				} catch {
					$StatusCode = $_.Exception.Response.StatusCode.value__
					if ($indexMois -ge ([convert]::ToInt32($moisActuel)-1)) {
						myLog "$Date Le fichier $fileMoisZip$ext n'est pas encore disponible, fin de procedure" 
					} else {
						myLog "$Date Le serveur devrait connaitre le fichier $fileMoisZip$ext, fin de procedure" 
					}
					exit
				}
			}
			# decompression
			Expand-Archive -Path $DestDir\$fileMoisZip$ext -DestinationPath $DestDir
			# On Renomme les fichiers Deces car certains noms sont trop longs
			#  <Deces_2022_M09 (version rectifi‚e).csv>
			$head = $PrefixCSV + $anActuel + "_M"
			$list = Get-ChildItem -Path $DestDir -Name -Include "$head*$ext2" -Exclude "*.zip"
			#
			$expectedLen = $head.Length + $moisActuel.Length + $ext2.Length
			foreach ($item in $list) {
					# myLog $item $item.Length $expectedLen
				if ( $item.Length -gt $expectedLen  ) {
					Rename-Item $DestDir/$item ($($item.substring(0,$expectedLen-$ext2.Length))+$ext2)
				}
			}
		}
		# on recherche le prenom et le nom
		$Search = Select-String -Path $DestDir\$fileMoisCsv$ext2 -pattern $regName
		if ([string]::IsNullOrEmpty($Search)) {
			myLog "$Date Pas de $prenom $nom dans $fileMoisCsv$ext2 " 
		} else { 
			$present = $false  # on continue meme apres un succes
			myLog "$Date $prenom $nom est presente dans $fileMoisCsv$ext2 ligne $((($Search -split (';'))[0] -split (':'))[2]) " 
			myLog "$Date Identite : $((($Search -split (';'))[0] -split (':'))[3]) , naissance: $(($Search -split (';'))[2]), deces $(($Search -split (';'))[6]) " 		
		}
		$indexMois++
	}
}
END {
myLog "$Date La recherche programme est finie " 
# end
}