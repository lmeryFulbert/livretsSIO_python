# Livrets SIO

Ce projet a pour but de formatter les données scolaires issues du logiciel pronote dans un format intermédiaire json afin de faciliter la création des livrets scolaire nécessaire pour la délibération des jury de BTS SIO.

## Prérequis

Avant de pouvoir utiliser ce programme, vous devez installer les modules suivants :

* openpyxl (pour manipuler les fichiers Excel en Python)

Vous pouvez les installer en exécutant la commande suivante :

```
pip install -r requirements.txt
```

ou plus simplement:

```
pip install openpyxl
```

## Comment utiliser ce programme

1. Créer un dossier correspondant à l'année en cours dans le repertoire in 
1. (bis) ou Modifiez les variables:

        - `file_in_pronote_etudiants`,
        - `file_in_pronote_an1`,
        - `file_in_pronote_an2`, 
        - `file_in_disciplines` 
    pour spécifier le chemin relatif des fichiers correspondants.

3. Exécutez le programme en tapant la commande suivante dans le terminal :
   
   ```
   python pronote2json.py
   ```

Le programme va alors traiter les fichiers spécifiés et produire les résultats dans le fichier `json/data.json`.

## Structure des fichiers

* `pronote2json.py` : le script principal à exécuter pour générer le fichier json intermédiaire.
* `in/` : le dossier contenant les fichiers d'entrée pour le programme.
    * `xxxx/` : repertoire contenant les fichiers excel issus de pronote.
        * `etudiants.xlsx`
        * `an1.xlsx`
        * `an2.xlsx`
        * `discipline.xlsx`
* `out/` : le dossier contenant les résultats produits par le programme `json2livrets.py`.
* `json/` : le dossier contenant les données de tous les livrets au format json
* `requirements.txt` : le fichier contenant la liste des modules Python nécessaires pour exécuter ce programme.
* `README.md` : ce fichier.

## Préparation des fichiers sources

### etudiants.xlsx

Les données sont issus de Pronote

1. Se connecter à Pronote en tant qu'enseignant
2. Aller dans Ressources / choisir la classe SIO2 / Eleves / Trier les élèves dans l'ordre alphabétique sur le nom.
3. Copier la liste au format csv et la copier dans la Feuil1

- Nom: Nom de l'élève
- Prénom: Prénom de l'élève
- Né(e) le: Date de naissance de l'élève
- Prénom d'usage: Prénom d'usage de l'élève
- Sexe: Sexe de l'élève
- Classe: Classe de l'élève

4. Supprimer les autres colonnes
5. Ajouter celles ci:

- Nombre de Pix: Nombre de points Pix obtenus par l'élève  (liste obtenu par le gestionnaire PIX)
- groupe: Groupe de l'élève (SLAM ou SISR)
- AVIS: décision du conseil de classe du second semestre:

        - TF : Très Favorable
        - F : Favorable
        - P : Doit faire ses preuves à l'examen

6. Compter les nombres d'avis grace à la formule NB.SI

        - AVIS: code de l'avis du conseil de classe (Colonne K)
        - NB: Nombre d'étudiants avec cet avis (Colonne L)
        - POURC: Pourcentage correspondant (Colonne M)

### an1.xlsx

1. Créer 3 feuilles avec ces noms:

        * SEMESTRE1
        * SEMESTRE2
        * SIO1ANNEE

2. Se connecter à Pronote en tant qu'enseignant sur la base de l'année précédante.
3. Aller dans Résultats / choisir la classe SIO1 / Tableau des moyennes
4. Choisir le bon semestre et le copier dans l'onglet correspondant.
5. Choisir les donénes de l'année entière et les copier dans l'onglet correspondant.

Pour les blocs d'enseignements répartis sur de multiples enseignants il faut établir des calculs.

6. Créer une feuille CALCUL
7. Copier les données nécéssaires avec les identités dans le même ordre alphabétique et les notes des enseignements du bloc 1 et du bloc 2

        - La colonne E doit contenir la moyenne du bloc 1 pour le semestre 1
        - La colonne F doit contenir la moyenne du bloc 1 pour le semestre 2
        - La colonne M doit contenir la moyenne du bloc 2 (uniquement valable pour le semestre 2)

### an2.xlsx

1. aller dans pronote / résultats / livrets scolaire / classe SIO2
2. choisir chaque étudiant et copier les données au format csv et les coller dans une nouvelle feuille qui seront nommées de manière itérative Feuil1, Feuil2, etc...,Feuiln
3. Vérifier que les nombre de feuilles obtenu correspond bien aux nombre d'élèves présent dans le fichier etudiants.xlsx

La structure (en 2023) de chaque feuille est structurée ainsi (idcolonne.):

    A. Disciplines : La matière évaluée.
    B. Notation : La période de la notation avec:
        - Sem1
        - Sem2
        - Année
    C. Rang de l'élève : Le rang de l'élève dans la classe pour cette matière.  (inutile)
    D. Moyenne de l'élève : La moyenne de l'élève pour cette matière à cette période
    E. Moyenne de la classe : La moyenne de la classe pour cette matière.
    E. % des moyennes inférieures à 8 : Le pourcentage de moyennes de la classe inférieures à 8 pour cette matière. (inutile)
    G. % des moyennes comprises entre 8 et 12 : Le pourcentage de moyennes de la classe comprises entre 8 et 12 pour cette matière. (inutile)
    H.% des moyennes supérieures à 12 : Le pourcentage de moyennes de la classe supérieures à 12 pour cette matière. (inutile)
    I. Appréciations des professeurs : Les commentaires ou appréciations des professeurs pour l'élève pour cette matière à cette période.

On a ainsi les 3 commentaires: Les 2 des bulletins pour le semestre 1 et 2, et le commentaire annuel qu'il faut placer sur le livret.


### discipline.xlsx

les codes des enseignements étant différents des blocs il faut établir les correspondances.
Le script charge les modules d'enseignement de premère et seconde année afin de les associer à un code de bloc d'enseignement officiel présent sur le livret

        FRA	Culture générale et expression
        ANG	Expression et communication en langue anglaise
        MAT	Mathématiques pour l’informatique
        CEJM	Culture économique, juridique et managériale
        CEJMA	Culture économique, juridique et managériale appliquée
        B1	Bloc 1 : Support et mise à disposition de services informatiques
        B2SLAM	"Bloc 2 : Administration des systèmes et des réseaux (option SISR)
        Bloc 2 : Conception et développement d’applications (option SLAM)"
        B2SISR	"Bloc 2 : Administration des systèmes et des réseaux (option SISR)
        Bloc 2 : Conception et développement d’applications (option SLAM)"
        B3	Bloc 3 : Cybersécurite des services informatiques
        AP	Ateliers de professionnalisation (1)
        O1	Langue vivante 2
        O2	Mathématiques approfondies 
        O3	Parcours de certification complémentaire

3 feuilles:

    - livrets
    - SIO1ANNEE
    - SIO2ANNEE

1. Copier les code de discipline issus des fichiers `an1.xlsx` et `an2.xlsx` et coller les dans la colonne A de chaque feuille en éliminant les doublons.
2. Associer le code du bloc correspondant.

    Exemple de données pour SIO1:

        CGENX	FRA
        CEJUM	CEJM
        MATAP	O2
        MA-AI	MAT
        AGL1	ANG
        ESP2	O1
        SUPPOR.	B1
        CYBERS	B3
        SISR1	B2
        SISR2	B2
        SLAM1	B2
        SLAM2	B2
        TC PBD	B1
        TC PBD	B1
        CYBERS	B3
        TC SR	B1
        B1-1	B1
        B1-2	B1
        B1-3	B1

Exemple de données pour SIO2

    ANGLAIS LV1 Professeur A ANG
    B1 MATIERE A M. B B1
    B2 MATIERE B M. B B2
    B2 MATIERE B M. C B2
    B3 MATIERE C M. B B3
    CERTIFICATION M. B O3
    CULT.ECO JUR. MANAG. Professeur D CEJM
    CULTURE GENE.ET EXPR Professeur E FRA
    MATHS APPROFONDIES Mme F O2
    MATHS POUR INFORMATQ Mme F MAT
    ESPAGNOL Mme G O1
    B2 MATIERE D M. H B2
    B2 MATIERE D M. C B2
    B3 MATIERE C M. C B3

## Generation des livrets

1. Vérifier la cohérence du `fichierjson/data.json`
2. Exécutez le programme en tapant la commande suivante dans le terminal :
   
   ```
   python json2livrets.py
   ```

## Auteur

Ce programme a été écrit par Ludovic MERY, enseignant en BTS SIO au lycée Fulbert de Chartres
