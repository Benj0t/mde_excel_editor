from xlrd import *
from xlwt import *
import xlrd
from xlutils.copy import copy

print("Bienvenue dans l'éditeur de tableaux excel.")
print("Si le programme se ferme inopinément, c'est probablement dû à une erreur dans le nom du fichier, relancez le programme en vérifiant de bien taper le nom du fichier et que son extension est bien '.xls'")
test = input('\nEntrez le nom du ficher que vous souhaitez modifier:\n')

rb = open_workbook(test + ".xls")
wb = copy(rb)
xlWorkbook = xlrd.open_workbook("./" + test + ".xls")
print("Ouverture du fichier", test + ".xls", ": OK")

style = XFStyle()
style.num_format_str = '[h]:mm:ss'

i = 0
j = 1
l = 0
f = xlWorkbook.nsheets
total = 0
total2 = 0
print("Modification du fichier En cours...")
while (l < f):
	s = wb.get_sheet(l)
	xlSheet = xlWorkbook.sheet_by_index(l)
	nx = xlSheet.nrows
	ny = xlSheet.ncols
	l = l + 1	
	while (i < nx):
		n = 0
		n2 = 0
		while(j < ny):
			cell = xlSheet.cell_value(i,j)
			if (cell == 'F'):
                                n2 = n2 + 85
			if (cell == 1 and i % 2 == 1):
				n = n + 85
				s.write(i,j,"1:05:00", style)
			if (cell == 1 and i % 2 == 0):
				n = n + 85
				s.write(i,j,"2:00:00", style)
			j = j + 1
		h = n // 60
		m = n % 60
		h2 = (n + n2) // 60
		m2 = (n + n2) % 60
		total = total + n
		total2 = total + n2 + n
		if (h != 0 or m != 0):
			s.write(i,j, str(h) + ":" + str(m) + ":00", style)
		if (h2 != 0 or m2 != 0):
			s.write(i,j + 1, str(h2) + ":" + str(m2) + ":00", style)
		else:
			s.write(i,j, "0:00:00", style)
			s.write(i,j + 1, "0:00:00", style)
		tmp = j
		j = 1
		i = i + 1
	h = total // 60
	m = total % 60
	h2 = total2 // 60
	m2 = total2 % 60
	s.write(i,tmp, str(h) + ":" + str(m) + ":00", style)
	s.write(i,tmp + 1, str(h2) + ":" + str(m2) + ":00", style)
	total = 0
	i = 0
	j = 1

print("Modification du fichier Terminée\nFichier édité avec succès.")
wb.save(test + "_edit" + ".xls")
print("Sauvegarde du fichier OK")
input("\nAppuyez sur ENTRÉE pour quitter le programme !")
