from xlrd import *
from xlwt import *
import xlrd
from xlutils.copy import copy

print("Bienvenue dans l'éditeur de tableaux excel.")
print("Si le programme se ferme inopinément, c'est probablement dû à une erreur dans le nom du fichier, relancez le programme en vérifiant de bien taper le nom du fichier et que son extension est bien '.xls'")
filename = input('\nEntrez le nom du ficher que vous souhaitez modifier:\n')

rb = open_workbook(filename + ".xls")
wb = copy(rb)
xlWorkbook = xlrd.open_workbook("./" + filename + ".xls")

# workbook = self.Workbook(self.dataDir + 'Book1.xls')
# worksheet = workbook.getWorksheets().get(0)

style = XFStyle()
style.num_format_str='MM-DD-YYYY'

classes = ["PS", "MS", "GS", "CP", "CE1", "CE2", "CM1", "CM2"]
child = []
name = 1
classe = 2
date = 3
i = 0
j = 0
k = 0
k2 = 0
l = 0
ref = ""

nsheets = xlWorkbook.nsheets
wb.add_sheet("LISTING")
listing = wb.get_sheet(nsheets)
print("Modification du fichier En cours...")

def is_child(ref, classes):
	i = 0
	while (i < 8):
		if (ref == classes[i]):
			return 0
		i = i + 1
	return 1

def	get_class(classes, ref):
	i = 0
	while (i < 8):
		if (ref == classes[i]):
			if (i <= 2):
				return 0
			return 1
		i = i + 1
	return 2

def parser(childs, ref):
	l = len(childs)
	i = 0
	if (ref == '' or ref == "Total journée"):
		return 1
	while (i < l):
		if (childs[i] == ref):
			return 1
		i = i + 1
	childs.append(ref)
	return 0


while (l < nsheets):
	sheet = wb.get_sheet(l)
	xlSheet = xlWorkbook.sheet_by_index(l)
	nx = xlSheet.nrows
	ny = xlSheet.ncols
	while (i < nx):
		ref = xlSheet.cell_value(i, name)
		ref_date = xlSheet.cell_value(i, date)
		ref_class = xlSheet.cell_value(i, classe)
		if (is_child(ref_class, classes) == 0 and parser(child, ref) == 0):
			if (get_class(classes, ref_class) == 0):
				listing.write(k, 0, ref)
				listing.write(k, 1, ref_class)
				listing.write(k, 2, ref_date, style)
				k = k + 1
			if (get_class(classes, ref_class) == 1):
				listing.write(k2, 4, ref)
				listing.write(k2, 5, ref_class)
				listing.write(k2, 6, ref_date, style)
				k2 = k2 + 1
		i = i + 1
	l = l + 1
	i = 0
listing.write(k, 0, ("TOTAL : " + str(k)))
listing.write(k2, 4, ("TOTAL : " + str(k2)))
wb.save(filename + "_listing" + ".xls")
print("Modification du fichier Terminée\nFichier édité avec succès.")
print("\nFIN")
