"""
BMWS Product Feed Parser for Alex
AUTHOR: Kalpataru Mallick
EMAIL: kalpataru.mallick@gmail.com
Skype: kalpataru.mallick
TIMEZONE: GMT-8
Date: 10/22/2013
"""

# -*- coding: utf-8 -*-
import os
import sys
import logging
import argparse
import csv
import pandas as pd

from xlrd import open_workbook
from django.utils.encoding import smart_str, smart_unicode
from decimal import Decimal

logger = logging.getLogger("DISTRIBUTORNAME")

def parse_args():
    """Parses the standard args and returns a configuration object."""
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--feed", required=True, type=str,
        help="Path to distributor feed file.")
		
    parser.add_argument(
        "--image", required=True, type=str,
        help="Path to distributor image file.")

    parser.add_argument(
        "--export", required=True, type=str,
        help="Path to write the BMWS product feed to.")

    parser.add_argument(
        "--debug-level", required=False, type=str, default="WARNING",
        choices=["INFO", "DEBUG", "WARNING", "ERROR", "CRITICAL"],
        help="Path to write the BMWS product feed to.")

    args = parser.parse_args()

    try:
        logging.basicConfig(level=getattr(logging, args.debug_level))
    except AttributeError:
        logger.error("Invalid debug level specified, defaulting to WARNING.")
        logging.basicConfig(level=logging.WARNING)

    return args

def parse_infile(xlsxpath, imagepath):
	wb = open_workbook(xlsxpath, 'rb')
	with open(imagepath) as f:
		lines = f.readlines()
	data=[]
	for s in wb.sheets():
		group = ""
		main = ""
		main2 = ""
		color_link = ""
		size_link = ""
		group_material = ""
		group_color = ""
		for row in range(s.nrows):
			datarow = []
			values = []
			words = []
			skus = []
			length = 0
			j = 0
			if row > 0 :
				count = 0
				images = []
				for col in range(s.ncols):
					if col in range(0,2):
						mystring =\
						str(smart_str(s.cell(row, col).value))\
						.replace(".0","")
					else:
						mystring = str(smart_str(s.cell(row, col).value))
					values.append(mystring)
				for i in range(0, 121):	
					if i == 0: #sku
						if values[0]:
							datarow.append(values[0])
							group = values[0]
							color_link = group
						elif values[1]:
							datarow.append(values[1])
							color_link = ""
						if datarow[0].isdigit():
							datarow[0]=str(Decimal(datarow[0]))
						main = datarow[0]
						main2 = main.lower()
						#print main2
						temp1 = datarow[0]
						words = temp1.split('-')
						length = len(words)
						j = 0
					elif i == 1: #advirtise
						datarow.append("TRUE")
					elif i == 2: #group
						datarow.append(group)
					elif i == 3: #subsku1, #subsku2, #subsku3, #subsku4, 
                        #subsku5, #subsku6, #subsku7, #subsku8, #subsku9, 
                        #subsku10,
						count2 = 0
						gflag = 1
						x = 0
						if values[0]:
							for row2 in range(row+1, s.nrows):
								if gflag is 1 and count2<10:
									values2 = []
									for col2 in range(0, 2):
										mystring2 =\
										str(s.cell(row2, col2).value)
										values2.append(mystring2)
									if values2[1].strip() is not "":
										skus.append(values2[1])
										count2 = count2+1
									else:
										gflag = 0
							count3 = 0
							for x in range(0, len(skus)):
								datarow.append(skus[x])
								count3 = count3+1
							while(count3<10):
								datarow.append("")
								count3 = count3+1
						else:
							while(count2<10):
								datarow.append("")
								count2 = count2+1					
					elif i == 13:  #UPC
						values[2]=values[2].replace('.', '').split('e')[0]
						if not values[2]:
							values[2]=0
						datarow.append(values[2])
					elif i == 14:  #Collection
						datarow.append(values[3])
					elif i == 15:  #title
						datarow.append(values[4])
					elif i == 16:  # Color1
						if values[0].strip() is not "":
							temp = values[5]
							temp2 = temp.replace('/', ',')
							words = temp2.split(',')
							length = len(words)
							words[0].strip()
							datarow.append(words[0])
							group_color = values[5]
						elif values[5].strip() is not "":
							temp = values[5]
							temp2 = temp.replace('/', ',')
							words = temp2.split(',')
							length = len(words)
							words[0].strip()
							datarow.append(words[0])
						else :
							temp = group_color
							temp2 = temp.replace('/', ',')
							words = temp2.split(',')
							length = len(words)
							words[0].strip()
							datarow.append(words[0])
						j = 1
					elif i in range(17, 20):  #Color2 #Color3 #Color4
						if length>j:
							words[j].strip()
							datarow.append(words[j])
							j = j+1
						else:
							datarow.append("")
					elif i == 20:  #colorlink
						datarow.append("")
					elif i == 21:  #fabric1
						temp = values[6]
						words  =  temp.split(',')
						length = len(words)
						words[0].strip()
						datarow.append(words[0])
						j = 1
						k = 7
					elif i in range(22, 24): #fabric2,  #fabric3
						if length>j:
							words[j].strip()
							datarow.append(words[j])
							j = j+1
						else:
							datarow.append("")
					elif i in range(24, 26): #price
						if values[k].strip() is '':
							datarow.append(values[k])
							k = k+1
						else:
							datarow.append("$"+values[k])
							k = k+1	
					elif i in range(26, 31): #category,  #subcategory1,
                        #subcategory2, #subcategory3, #subcategory4
						datarow.append(values[k])
						k = k+1	
					elif i in range(31, 35): #weight, #height, #width, 
                        #depth
						datarow.append(values[k+1])
						k = k+1
					elif i in range(35, 36): #sizelink
						datarow.append("")
						k = k+1
					elif i in range(36, 37): #class
						datarow.append(values[k+1])
						k = k+1
					elif i in range(37, 38): #description
						tem = values[k+2]
						datarow.append(tem)
						k = k+1
					elif i in range(38, 42): #feature 1, 
                        #feature 2, #feature 3, #feature 4
						datarow.append(values[k+2])
						k = k+1
					elif i in range(42, 43): #feature 5, 
						datarow.append(values[k+2]+" "+values[k+3])
						k = k+2
					elif i in range(43, 44): #assembly
						datarow.append(values[k+2])
						k = k+1
					elif i == 44: #material1
						if values[0].strip() is not "":
							temp = values[30]
							words  =  temp.split(',')
							length = len(words)
							words[0].strip()
							datarow.append(words[0])
							group_material = values[30]
						elif values[30].strip() is not "":
							temp = values[30]
							words  =  temp.split(',')
							length = len(words)
							words[0].strip()
							datarow.append(words[0])
						else:
							temp = group_material
							words  =  temp.split(',')
							length = len(words)
							words[0].strip()
							datarow.append(words[0])			
						j = 1
					elif i in range(45, 48): #material2, #material3, 
                        #material4
						if length>j:
							words[j].strip()
							datarow.append(words[j])
							j = j+1
						else:
							datarow.append("")
						j = 31
					elif i in range(48,49): #wood1, #wood2, #wood3
						y=0
						if values[j].strip() is not "":
							temp = values[j]
							words  =  temp.split(',')
							length = len(words)
							words[0].strip()
							datarow.append(words[0])
							y=y+1
							while((y<length) and (y<3)):
								words[y].strip()
								datarow.append(words[y])
								y=y+1
							while(y<3):
								datarow.append("")
								y=y+1
						else:
							while(y<3):
								datarow.append("")
								y=y+1
						j = j+1
						
					elif i in range(49,90): #origin	shipmethod	
                        #boxcount	boxweight	boxheight	boxwidth	
                        #boxdepth	box2weight	box2height	box2width	
                        #box2depth	box3weight	box3height	box3width	
                        #box3depth	box4weight	box4height	box4width	
                        #box4depth	box5weight	box5height	box5width	
                        #box5depth	box6weight	box6height	box6width	
                        #box6depth	box7weight	box7height	box7width	
                        #box7depth	box8weight	box8height	box8width	
                        #box8depth	box9weight	box9height	box9width	
                        #box9depth	estimatedship	shipnote
						datarow.append(values[j])
						j = j+1
					elif i == 90: #mainimage1, #mainimage2, #mainimage3
						count = 0
						for line in lines:
							line3=line.strip()
							line = line.replace("hh","").replace(".jpg","")
							line = line.strip().replace(" ","")
							main2 = main2.strip().replace(" ","")
							if (main2 in line) and (line in main2)\
							and (count<3):
								temp="hh"+line+".jpg"
								datarow.append(line3)
								count = count+1
								images.append(temp)
						while count<3:
							datarow.append("")
							count = count+1
					elif i == 93: #setimage1, #setimage2, #setimage3
						count = 0
						for line in lines:	
							line3=line.strip()
							line = line.replace("hh","").replace(".jpg","")
							line = line.strip()
							main2 = main2.strip()
							if (main2 in line) and ('&' in line)\
							and (count<3): 
								temp="hh"+line+".jpg"
								datarow.append(line3)
								count = count+1
								images.append(temp)
						while count<3:
							datarow.append("")
							count = count+1							
					elif i == 96: #roomimage1, #roomimage2, #roomimage3
						count = 0
						tokens  =  ["room", "bedroom", "bed", "daybed", 
                        "headboard", "footboard"]
						for line in lines:	
							line3=line.strip()
							line = line.replace("hh","").replace(".jpg","")
							line = line.strip()
							main2 = main2.strip()						
							if (main2 in line) and \
							any((t in line for t in tokens)) and (count<3):
								temp="hh"+line+".jpg"
								datarow.append(line3)
								count = count+1
								images.append(temp)
						while count<3:
							datarow.append("")
							count = count + 1
					elif i == 99: #additionalimage1	additionalimage2	
                    #additionalimage3	additionalimage4	additionalimage5
                    #additionalimage6	additionalimage7	additionalimage8
                    #additionalimage9	additionalimage10
						count = 0
						tokens  =  ["detail", "top", "bottom", "side", 
                        "left", "right", "up", "down", "back", "front", 
                        "out", "corner", "with", "without", "open", "view",
                        "alt", "dresser", "chest", "nightstand", "finish", 
                        "height", "end"]
						for line in lines:
							line3=line.strip()
							line = line.replace("hh","").replace(".jpg","")
							line = line.strip()
							main2 = main2.strip()
							if (main2 in line) and ('&' not in line) and \
							any((t in line for t in tokens)) and count<10:
								temp="hh"+line+".jpg"
								datarow.append(line3)
								count = count+1
								images.append(temp)
						while count<10:
							datarow.append("")
							count = count+1		
					elif i == 109: #diagramimage1, #diagramimage2, 
                        #diagramimage3
						count = 0
						tokens = ["diagram", "sketch", "instructions", 
                        "drawing"]
						for line in lines:
							line3=line.strip()
							line = line.replace("hh","").replace(".jpg","")
							line = line.strip()
							main2 = main2.strip()
							if (main2 in line) and \
							any((t in line for t in tokens)) and (count<3):
								temp="hh"+line+".jpg"
								datarow.append(line3)
								count = count+1
								images.append(temp)
						while count<3:
							datarow.append("")
							count = count+1		
					elif i == 112: #swatchimage1	swatchimage2	
                        #swatchimage3 swatchimage4	swatchimage5
                        #swatchimage6 swatchimage7	swatchimage8
                        #swatchimage9 swatchimage10
						count = 0
						tokens  =  ["dark", "black", "white", "red", 
                        "brown", "light"]
						for line in lines:
							line3=line.strip()
							line = line.replace("hh","").replace(".jpg","")
							line = line.strip()
							main2 = main2.strip()
							if (main2 in line) and \
							any((t in line for t in tokens)) and (count<10):
								temp="hh"+line+".jpg"
								datarow.append(line3)
								count = count+1
								images.append(temp)
						while count<10:
							datarow.append("")
							count = count+1
				data.append(datarow)
			else :
				datarow=['sku', 'advertise', 'group', 'subsku1', 'subsku2',
                'subsku3', 'subsku4', 'subsku5', 'subsku6', 'subsku7', 
                'subsku8', 'subsku9', 'subsku10', 'upc', 'collection', 
                'title', 'color1', 'color2', 'color3', 'color4', 'colorlink',
                'fabric1', 'fabric2', 'fabric3', 'price', 'map', 'category',
                'subcategory1', 'subcategory2', 'subcategory3', 
                'subcategory4', 'weight', 'height', 'width', 'depth', 
                'sizelink', 'class', 'description', 'feature 1', 'feature 2',
                'feature 3' , 'feature 4', 'feature 5', 'assembly', 
                'material1', 'material2', 'material3', 'material4', 
                'wood1', 'wood2', 'wood3', 'origin', 'shipmethod', 
				'boxcount', 'boxweight', 
                'boxheight', 'boxwidth', 'boxdepth' , 'box2weight', 
                'box2height', 'box2width', 'box2depth' , 'box3weight', 
                'box3height', 'box3width', 'box3depth' , 'box4weight', 
                'box4height', 'box4width', 'box4depth' , 'box5weight', 
                'box5height', 'box5width', 'box5depth' , 'box6weight', 
                'box6height', 'box6width', 'box6depth' , 'box7weight', 
                'box7height', 'box7width', 'box7depth' , 'box8weight', 
                'box8height', 'box8width', 'box8depth', 'box9weight', 
                'box9height', 'box9width', 'box9depth', 'estimatedship', 
                'shipnote', 'mainimage1', 'mainimage2', 'mainimage3', 
                'setimage1', 'setimage2', 'setimage3', 'roomimage1', 
                'roomimage2', 'roomimage3', 'additionalimage1',  
                'additionalimage2', 'additionalimage3', 'additionalimage4', 
                'additionalimage5', 'additionalimage6', 'additionalimage7', 
                'additionalimage8', 'additionalimage9', 'additionalimage10',
                'diagramimage1', 'diagramimage2', 'diagramimage3', 
                'swatchimage1', 'swatchimage2', 'swatchimage3', 
                'swatchimage4', 'swatchimage5', 'swatchimage6', 
                'swatchimage7', 'swatchimage8', 'swatchimage9', 
                'swatchimage10']
				data.append(datarow)
	return data

def write_outfile(path, data):
	df = pd.DataFrame(data)
	with open(path, 'wb') as f:
		df.to_csv(f, header = False, skiprows = 1, index = False, 
        names = data[0], encoding="utf-8")

def main():
    """Entry point for the application."""
    args  =  parse_args()
    logger.debug("Started parsing file %s" % args.feed)
    data =  parse_infile(args.feed, args.image)
    logger.debug("Started writing file %s" % args.export)
    write_outfile(args.export, data)
    pass

if __name__ == "__main__":
    main()