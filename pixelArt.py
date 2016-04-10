# -*- coding: utf-8 -*-

import xlwt
import sys

#tranforme le format en entree en liste 
def parseData(str):
    size = len(str)
    #print size
    data = list()
    i = 0    
    while i < size:
        if str[i] == 'l':
            data.append(2)
        elif str[i] == '*':
            data.append(1)
        elif str[i] == '-':
            data.append(3)
        else:
            data.append(0)
        i = i + 1
    return data

#nombre de retour a la ligne
def nbLigne(str):
    nbLine = 1
    y = 0
    size = len(str)
    while y < size:
        if str[y] == 'l':
            nbLine += 1
        y += 1
    return nbLine




def main():
    #recuperer les arguments passé en paramètre
    if len(sys.argv) != 2:
        print "Not enought arg"
    else:
        string = sys.argv[1]
   #gestion des couleur
    #book = xlwt.Workbook()
    # add new colour to palette and set RGB colour value
   # xlwt.add_palette_colour("custom_colour", 0x21)
    #book.set_colour_RGB(0x21, 251, 228, 228)

    # now you can use the colour in styles
        style = xlwt.easyxf('pattern: pattern solid, fore_colour white')
        style2 = xlwt.easyxf('pattern: pattern solid, fore_colour black')
        style3 = xlwt.easyxf('pattern: pattern solid, fore_colour  yellow')

        name = "pixel.xls"
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Pixel Art")
        data = list()
        x = 0
        y = 0
        count = 0
        #nombre de ligne
        nbL = 0
        #convertir la chaine
        data = parseData(string)
        print data
        #boucle de modification
        i = 0
        while i < 10:
            #ws.row(count).height_mismatch = True
            #ws.row(count).height = 100
            #ws.col(nbL).width = 1000
            i += 1
        while count < len(data):
    
            if data[count] == 2:
                nbL += 1
                x = 0
            else:
                if data[count] == 3:
                    ws.write(nbL, x, ' ', style3)
                if data[count] == 1:
                    ws.write(nbL, x, ' ', style2)    
                if data[count] == 0:
                    ws.write(nbL, x, ' ', style)
                x += 1
            count += 1            
       
        #save folder
        wb.save(name)
main()
