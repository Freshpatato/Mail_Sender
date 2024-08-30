from spire.xls import * 
from spire.xls.common import * 

# Lire un fichier Excel dans un flux
excelStream = Stream("C:\\Users\\maxime.bourquard\\Downloads\\Classeur1.xlsx") 

# Instancier un objet Workbook
workbook = Workbook() 
# Charger le fichier Excel à partir du flux
workbook.LoadFromStream(excelStream) 

# Obtenir la première feuille de calcul
sheet = workbook.Worksheets[ 0 ] 

# Définir les options de conversion
options = HTMLOptions() 
options.ImageEmbedded = True 

# Enregistrer la feuille de calcul au format HTML
htmlStream = Stream("C:\\Users\\maxime.bourquard\Downloads\\Classeur1_modif.html") 
sheet.SaveToHtml(htmlStream, options) 

htmlStream.Close() 
workbook.Dispose()