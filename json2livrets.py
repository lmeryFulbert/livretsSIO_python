import openpyxl     #faire un pip install openpyxl pour installer la librairie
from openpyxl.styles import Alignment, Font
import datetime
import json

def traitematiere(linenumber, feuille, results):
    code='A' + linenumber
    cell_sem1 = feuille[code]
    cell_sem1.value = results['moyenne_sem1']
    
    code='B' + linenumber
    cell_sem2 = feuille[code]
    cell_sem2.value = results['moyenne_sem2']

    code='C' + linenumber
    cell_an1 = feuille[code]
    cell_an1.value = results['moyenne_an1']

def traitematierean2(linenumber, feuille, results):
    code='H' + linenumber
    cell_sem1 = feuille[code]
    cell_sem1.value = results['moy_sem1']
    
    code='I' + linenumber
    cell_sem2 = feuille[code]
    cell_sem2.value = results['moy_sem2']

    code='J' + linenumber
    cell_an1 = feuille[code]
    cell_an1.value = results['moy_an2']

    code='K' + linenumber
    cell_commentaire = feuille[code]
    cell_commentaire.value = results['commentaire']
    cell_commentaire.alignment = Alignment(wrap_text=True,vertical='top') 

    #si le commentaire est trop long, on agrandit la cellule
    if(results['commentaire'] is not None and len(results['commentaire'])>=120):
        cell_commentaire.font = Font(size=10)
        sheet.row_dimensions[cell_commentaire.row].height = 40
    else:
        cell_commentaire.font = Font(size=11)
        sheet.row_dimensions[cell_commentaire.row].height = 30


    code='N' + linenumber
    cell_moyclasse = feuille[code]
    cell_moyclasse.value = results['moy_classe']

#flush les commentaires et notes des étudiants ne suivants pas l'option LV2
def flushlivret(feuille):
    feuille['A16'].value = None
    feuille['B16'].value = None
    feuille['C16'].value = None
    feuille['H16'].value = None
    feuille['I16'].value = None
    feuille['J16'].value = None
    feuille['K16'].value = None


annee = datetime.datetime.now().year

# Spécifier le chemin relatif du livret vierge
file_in_livret = 'in/livret_vierge.xlsx'


# Spécifier le chemin relatif du dossier de sortie
directory_output = 'out'

# Ouvrir les fichiers excel
workbook_livret = openpyxl.load_workbook(file_in_livret)


# Ouverture du fichier json
f = open('json/data.json',encoding='utf-8')
  
# returns JSON object as 
data = json.load(f)

TRESFAV=data['statistiques']['TF']
FAV=data['statistiques']['F']
PREUVE=data['statistiques']['P']
  
# Boucle pour parcourir les données de chaque étudiants
for etudiant in data['livrets']:
    #print(etudiant)

    # # Accéder à une feuille spécifique du modèle
    sheet = workbook_livret['Recto']

    # # Accéder à la cellule contenant l'année du livret
    cell_annee = sheet['G1']
    cell_annee.value = annee

    # identité
    cell_nom = sheet['H1']
    cell_nom.value="NOM :" + etudiant['nom']

    cell_prenom = sheet['H2']
    cell_prenom.value="Prénom :" + etudiant['prenom']

    cell_date_naiss = sheet['K3']
    cell_date_naiss.value= etudiant['date_naiss']

    # specialité
    cell_spe = sheet['D3']
    if(etudiant['specialite']=='SISR'):
        cell_spe.value= '\u2611 Solutions d’infrastructure, systèmes et réseaux (SISR) \n \u2610 Solutions logicielles et applications métiers (SLAM)'
    else:
        cell_spe.value= '\u2610 Solutions d’infrastructure, systèmes et réseaux (SISR) \n \u2611 Solutions logicielles et applications métiers (SLAM)'

    for resultsan1 in etudiant['an1']:

        match resultsan1['discipline']:
            case 'FRA':
                traitematiere('6', sheet, resultsan1)
            
            case 'ANG':
                traitematiere('7', sheet, resultsan1)  

            case 'MAT':
                traitematiere('8', sheet, resultsan1)

            case 'CEJM':
                traitematiere('9', sheet, resultsan1)

            case 'CEJMA':
                traitematiere('10', sheet, resultsan1)  

            case 'B1':
                traitematiere('11', sheet, resultsan1)  

            case 'B2':
                traitematiere('12', sheet, resultsan1)  

            case 'B3':
                traitematiere('13', sheet, resultsan1)  

            case 'O1':
                traitematiere('16', sheet, resultsan1)
            
            case 'O2':
                traitematiere('17', sheet, resultsan1)

            #case 'O3':
            #    traitematiere('18', sheet, resultsan1)

            case _:
                print('Discipline non reconnue :' + str(resultsan1['discipline']))

    somsem1_b2=0
    somsem2_b2=0
    nbmodule_b2=0
    somclasse_b2=0
    comment_b2=''
    
    for resultsan2 in etudiant['an2']:

        match resultsan2['discipline']:
            case 'FRA':
                traitematierean2('6', sheet, resultsan2)
            
            case 'ANG':
                traitematierean2('7', sheet, resultsan2)  

            case 'MAT':
                traitematierean2('8', sheet, resultsan2)

            case 'CEJM':
                traitematierean2('9', sheet, resultsan2)
                traitematierean2('10', sheet, resultsan2)  

            case 'CEJMA':
                traitematierean2('10', sheet, resultsan2)  

            case 'O1':
                traitematierean2('16', sheet, resultsan2)
            
            case 'O2':
                traitematierean2('17', sheet, resultsan2)

            #case 'O3':
            #    traitematiere('18', sheet, resultsan2)

            case 'B1':
                traitematierean2('11', sheet, resultsan2)  

            case 'B3':
                traitematierean2('13', sheet, resultsan2)  

            case 'AP':
                traitematierean2('14', sheet, resultsan2) 
                code='K14'
                cell_commentaire = sheet[code]
                cell_commentaire.value = resultsan2['commentaire']
                cell_commentaire.alignment = Alignment(wrap_text=True,vertical='top') 
                #si le commentaire est trop long, on agrandit la cellule
                if(resultsan2['commentaire'] is not None and len(resultsan2['commentaire'])>=120):
                    cell_commentaire.font = Font(size=10)
                    sheet.row_dimensions[cell_commentaire.row].height = 40
                else:
                    cell_commentaire.font = Font(size=11)
                    sheet.row_dimensions[cell_commentaire.row].height = 30

            case 'B2':
                somsem1_b2=somsem1_b2 + resultsan2['moy_sem1'] 
                somsem2_b2=somsem2_b2 + resultsan2['moy_sem2']
                somclasse_b2=somclasse_b2 + resultsan2['moy_classe']
                nbmodule_b2=nbmodule_b2+1

                comment_b2=comment_b2 + str(resultsan2['commentaire']) + '\n'
            
            case _:
                print('Discipline non reconnue :' + str(resultsan2['discipline']))

    # traitement du bloc 2 par plusieurs enseignants
    #
    if (somsem1_b2!=0):
        moysem1_b2=somsem1_b2/nbmodule_b2
        moysem2_b2=somsem2_b2/nbmodule_b2
        moyclasse_b2=somclasse_b2/nbmodule_b2

        moyannuelle_b2=(moysem1_b2+moysem2_b2)/2

        code='H12'
        cell_sem1 = sheet[code]
        cell_sem1.value = moysem1_b2
        
        code='I12'
        cell_sem2 = sheet[code]
        cell_sem2.value = moysem2_b2

        code='J12'
        cell_an1 = sheet[code]
        cell_an1.value = moyannuelle_b2

        code='K12'
        cell_commentaire = sheet[code]
        cell_commentaire.value = comment_b2
        cell_commentaire.alignment = Alignment(wrap_text=True,vertical='top') 
        #si le commentaire est trop long, on agrandit la cellule
        if(comment_b2 is not None and len(comment_b2)>=120):
            cell_commentaire.font = Font(size=10)
            sheet.row_dimensions[cell_commentaire.row].height = 50
        else:
            cell_commentaire.font = Font(size=11)
            sheet.row_dimensions[cell_commentaire.row].height = 30

        code='N12'
        cell_moyclasse = sheet[code]
        cell_moyclasse.value = moyclasse_b2


    # Gestion de Pix
    nbpointspix=etudiant['pix']
    #\u2611 = case cochée
    sheet['A19'].value="\u2611   A obtenu la certification de compétences numériques totalisant "+ str(nbpointspix) +" Pix / 768 points"


    # Gestion de l'avis
    
    match etudiant['avis']:
        case 'TF':
            sheet['A22']='Très favorable'
        case 'F':
            sheet['A22']='Favorable'
        case 'P':
            sheet['A22']='Doit faire ses preuves à l\'examen'
        case _:
            print("erreur pour l'avis")

    sheet['A22'].alignment = Alignment(wrap_text=True,vertical='center') 
    sheet['A22'].font = Font(size=14)


    #Gestion des pourcentage
    sheet['E24']=TRESFAV
    sheet['F24']=FAV
    sheet['G24']=PREUVE


    # Enregistrement du fichier
    filename=etudiant['nom']+'_'+etudiant['prenom']+'.xlsx'

    # # Sauvegarder les modifications
    workbook_livret.save(directory_output +'/' + filename)

    # Flush LV2 line
    flushlivret(sheet)

  
# Closing file
f.close()



