#======================================== PARTIE 0 - FONCTIONS ========================================#

function CreateLabel #Fonction de création d'un label
{
    param ([string]$text,[int]$size1,[int]$size2,[int]$size3,[int]$size4)
    $label = New-Object system.Windows.Forms.Label
    $label.text = $text
    $label.size = New-Object System.Drawing.Size($size1,$size2)
    $label.location = New-Object System.Drawing.Size($size3,$size4)
    $label.Font = 'Microsoft Sans Serif,10'
    return $label
}

function CreateBox #Fonction de création d'une texteBox
{
    param ([int]$size1,[int]$size2,[int]$size3,[int]$size4)
    $box = New-Object System.Windows.Forms.Textbox
    $box.size = New-Object System.Drawing.Size($size1,$size2)
    $box.location = New-Object System.Drawing.Size($size3,$size4)
    $box.font = 'Microsoft Sans Serif,10'
    return $box
}

function CreateComboBox #Fonction de création d'une comboBox
{
    param ([int]$size1,[int]$size2,[int]$size3,[int]$size4,[int]$size5, $item1, $item2, $item3, $item4, $item5, $item6, $item7, $item8, $item9, $item10)
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboBox.Size = New-Object System.Drawing.Size($size1,$size2)
    $comboBox.Location = New-Object System.Drawing.Size($size3,$size4)
    [void] $comboBox.Items.Add($item1)
    if ($item2){[void] $comboBox.Items.Add($item2)}
    if ($item3){[void] $comboBox.Items.Add($item3)}
    if ($item4){[void] $comboBox.Items.Add($item4)}
    if ($item5){[void] $comboBox.Items.Add($item5)}
    if ($item6){[void] $comboBox.Items.Add($item6)}
    if ($item7){[void] $comboBox.Items.Add($item7)}
    if ($item8){[void] $comboBox.Items.Add($item8)}
    if ($item9){[void] $comboBox.Items.Add($item9)}
    if ($item10){[void] $comboBox.Items.Add($item10)}
    return $comboBox
}

function CreateButton #Fonction de création d'un button
{
    param ([int]$size1,[int]$size2,[int]$size3,[int]$size4, [string]$text)
    $button = New-Object System.Windows.Forms.Button
    $button.Size = New-Object System.Drawing.Size($size1,$size2)
    $button.Location = New-Object System.Drawing.Size($size3,$size4)
    $button.Text = $text
    $button.Add_Click({ Write-Host "test test" ; $formulaire.Close()})
    return $button
}

function Get-RandomPassword #Fonction de génération de mot de passe aléatoire
{
    param (
        [Parameter(Mandatory)]
        [int] $length,
        [int] $amountOfNonAlphanumeric = 1
    )
    Add-Type -AssemblyName 'System.Web'
    return [System.Web.Security.Membership]::GeneratePassword($length, $amountOfNonAlphanumeric)
}

#======================================== PARTIE 1 - CODE ========================================#

<# ----------------------------------------------- FENETRE DE BASE ---------------------------------------------- #>

Add-Type -AssemblyName System.Windows.Forms

###Création de la fenêtre
$formulaire = New-Object system.Windows.Forms.Form
$formulaire.ClientSize = '320, 375'
$formulaire.text =  "Aide au deploiement Micropole"
$formulaire.BackColor = "#E0DDEE"
$formulaire.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

###Création du texte de base
$title = New-Object system.Windows.Forms.Label
$title.text = "Gestion des informations :"
$title.size = New-Object System.Drawing.Size(1400,30)
$title.location = New-Object System.Drawing.Size(10,20)
$title.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Underline)
$formulaire.Controls.Add($title)

<# ------------------------------------------------ TITRE ----------------------------------------------- #>

###Création du label de "Titre"
$formulaire.Controls.Add((CreateLabel -text "Titre :" -size1 140 -size2 20 -size3 10 -size4 60)) #1

###Création de la combobox "Titre"
$formulaire.Controls.Add((CreateComboBox -size1 150 -size2 20 -size3 160 -size4 60 -item1 "Monsieur" -item2 "Madame" )) #2

<# ------------------------------------------------ NOM ----------------------------------------------- #>

###Création du label de "Nom"
$formulaire.Controls.Add((CreateLabel -text "Nom :" -size1 140 -size2 20 -size3 10 -size4 90)) #3

###Création de la textbox "Nom"
$formulaire.Controls.Add((CreateBox -size1 150 -size2 20 -size3 160 -size4 90)) #4

<# -------------------------------------------------- PRÉNOM -------------------------------------------------- #>

###Création du label de "Prénom"
$formulaire.Controls.Add((CreateLabel -text "Prenom :" -size1 140 -size2 20 -size3 10 -size4 120)) #5

###Création de la textbox "Prénom"
$formulaire.Controls.Add((CreateBox -size1 150 -size2 20 -size3 160 -size4 120)) #6

<# ------------------------------------------------ TYPE ----------------------------------------------- #>

###Création du label de "Type"
$formulaire.Controls.Add((CreateLabel -text "Type :" -size1 140 -size2 20 -size3 10 -size4 150)) #7

###Création de la combobox "Type"
$formulaire.Controls.Add((CreateComboBox -size1 150 -size2 20 -size3 160 -size4 150 -item1 "CDI" -item2 "CDD" -item3 "Stagiaire/Alternant" -item4 "Prestataire")) #8

<# ------------------------------------------------ UNITE ----------------------------------------------- #>

###Création du label de "Unite"
$formulaire.Controls.Add((CreateLabel -text "Unite :" -size1 140 -size2 20 -size3 10 -size4 180)) #9

###Création de la textbox "Unite"
$formulaire.Controls.Add((CreateBox -size1 150 -size2 20 -size3 160 -size4 180)) #10

<# ------------------------------------------------ LOCALISATION ----------------------------------------------- #>

###Création du label de "Localisation"
$formulaire.Controls.Add((CreateLabel -text "Localisation :" -size1 140 -size2 20 -size3 10 -size4 210)) #11

###Création de la combobox "Localisation"
$formulaire.Controls.Add((CreateComboBox -size1 150 -size2 20 -size3 160 -size4 210 -item1 "AIX" -item2 "LEVALLOIS" -item3 "LYON" -item4 "NANTES" -item5 "NIORT" `
                                                                                    -item6 "RENNES" -item7 "BRUXELLES" -item8 "LUXEMBOURG" -item9 "CHINE" -item10 "SUISSE" )) #12

<# ------------------------------------------------ M365 ----------------------------------------------- #>

###Création du label de "M365l"
$formulaire.Controls.Add((CreateLabel -text "Licence M365 :" -size1 140 -size2 20 -size3 10 -size4 240)) #13

###Création de la combobox "M365"
$formulaire.Controls.Add((CreateComboBox -size1 150 -size2 20 -size3 160 -size4 240 -item1 "F3" -item2 "E3" -item3 "E1" )) #14

<# ---------------------------------------------------- TÉLÉPHONE ---------------------------------------------------- #>

###Création du label de "Téléphone"
$formulaire.Controls.Add((CreateLabel -text "Telephone :" -size1 140 -size2 20 -size3 10 -size4 270)) #15

###Création de la textbox "Téléphone"
$formulaire.Controls.Add((CreateBox -size1 150 -size2 20 -size3 160 -size4 270)) #16



<#------------------------------------------------- BOUTONS & APPEL ------------------------------------------------- #>

###Création des buttons

$formulaire.Controls.Add((CreateButton -size1 100 -size2 30 -size3 110 -size4 320 -text 'Valider'))

###Affichage du formulaire
[void]$formulaire.ShowDialog()



<#---------------------------------------------------- VARIABLES ---------------------------------------------------- #>

#$checkbox1.Add_CheckStateChanged({ $checkbox1.Checked = $True})

$titre = $formulaire.Controls.Text[2]
$nom = $formulaire.Controls.Text[4]
$prenom = $formulaire.Controls.Text[6]
$type = $formulaire.Controls.Text[8]
$unite = $formulaire.Controls.Text[10]
$localisation = $formulaire.Controls.Text[12]
$M365 = $formulaire.Controls.Text[14]
$telephone = $formulaire.Controls.Text[16]

$password = Get-RandomPassword 8


#=================================== PARTIE 2 - CREATION DU COMPTE ==============================================#

#Setup username with the first firstname letter + name
$array = $prenom.ToCharArray()
$username = $array[0]+$nom

#Setup email address
$mail = "$username@AD.local"

#Setup OU path where account will be create
$OUpath = "OU=ORG_France,DC=AD,DC=local"

#Check if account already exist
$checkIfExist = [bool] (Get-ADUser -Filter { SamAccountName -eq $username }) # returns true if exist
if ($checkIfExist -eq $true) {Write-Host "Utilisateur déjà existant" -ForegroundColor Red;}

#Setup licence 
$GroupeM365 = if($M365 -eq "E3"){$GroupeM365 = "M365-E3"};if($M365 -eq "F3"){$GroupeM365 = "M365-F3"};if($M365 -eq "E1"){$GroupeM365 = "M365-E1"};



New-ADUser -Server AD-LAB -SurName $nom -GivenName $prenom -Name "$prenom $nom" -DisplayName "$prenom $nom" `
           -UserPrincipalName $mail -SamAccountName $username -Path $OUpath -EmailAddress $mail -Department 'Agence Centre Est' `
           -Company 'Micropole Rhône-Alpes' -StreetAddress '131 Boulevard Stalingrad' -PostalCode '69100' -City 'Villeurbanne' -Description 'Employee' ; `
           Set-ADAccountPassword -Server AD-LAB -Identity $username -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force); `
           Set-ADUser -Identity $username -Server AD-LAB -Enabled $true -ChangePasswordAtLogon $false `
           -Add @{extensionAttribute10='France';personalTitle='M';co='France'; c='FR'; proxyAddresses=$mail; mailNickname=$username};

           Add-ADGroupMember -Server AD-LAB -Identity $GroupeM365 -Members $username;		
