
🚀 VBA Automation for Interco and REC Files
📝 Description
Ce projet contient des scripts VBA pour automatiser des tâches répétitives dans Excel. Ces macros permettent :

Extraction de données :

Recherche d’un fichier dont le nom commence par "-RETOUR -interco".
Extraction des valeurs des cellules C5 et C6.
Mise à jour des fichiers cibles :

Recherche d’un fichier dont le nom commence par "REC".
Collage des valeurs extraites dans la feuille REC des fichiers correspondants.
Gestion de plusieurs formats de fichiers REC :
REC avec 12 lignes,
REC avec 13 lignes,
REC avec 17 lignes,
REC avec 22 lignes.
Ces scripts simplifient et accélèrent le processus tout en réduisant les erreurs manuelles.

📂 Structure des fichiers
Nom du fichier	Description
Create sheet REC	Script pour créer ou mettre à jour les feuilles REC.
REC 13 ligne.cls	Macro pour les fichiers REC contenant 13 lignes.
REC 17 ligne.cls	Macro pour les fichiers REC contenant 17 lignes.
REC 22 ligne.cls	Macro pour les fichiers REC contenant 22 lignes.
Recols12Ligne.cls	Macro pour les fichiers REC contenant 12 lignes.
🛠️ Instructions
Étape 1 : Préparation des fichiers
Placez les fichiers -RETOUR -interco et REC dans un répertoire commun.
Assurez-vous que les noms des fichiers respectent la structure attendue.
Étape 2 : Exécution des macros
Ouvrez Excel et activez l’exécution des macros.
Importez le fichier .cls correspondant dans l’éditeur VBA.
Lancez la macro adaptée selon le format de votre fichier REC.
Exemple de lancement d'une macro dans l'éditeur VBA :

vba
Copier
Modifier
Sub RunRE
