import sys
import fitz
import shutil
import os
import argparse
import time
from openpyxl import Workbook, load_workbook
# If you need to get the column letter, also import this
from openpyxl.utils import get_column_letter
import unicodedata as ud
from sys import platform
from os import walk
import json
import xlwings as xw
import easyocr
from pathlib import Path
from pdf2image import convert_from_path
import numpy
import warnings
import amazon_lib as lib

warnings.filterwarnings("ignore", category=Warning)

def clearlist(*args):
    for varlist in args:
        varlist.clear()



def file_delimeter():
    delimeter = "/"    
    if platform == "win32":
        delimeter = "\\"
    return delimeter
    
def data_generator(xlsworksheet):
    shipmentlist = []
    shipreadylist = []
    maxrow = xlsworksheet.range('B' + str(xlsworksheet.cells.last_cell.row)).end('up').row
    for i in range(2, maxrow + 2):
        shipment_row = str(xlsworksheet['A{}'.format(i)].value)
        if shipment_row.find('Shipment') != -1:
            # print(shipment_row, i)
            startrow = i
            y = i
            shipment_empty = True
            while True:
                y += 1
                # skip if shipment_id was filled
                if ''.join(str(xlsworksheet['B{}'.format(y)].value)).strip() == 'Shipment ID':
                    # if ''.join(str(xlsworksheet['E{}'.format(y)].value)).strip() != 'None':
                    #     shipment_empty = False
                    boxes = ('E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P')
                    for box in boxes:
                        if ''.join(str(xlsworksheet['{}{}'.format(box, y)].value)).strip() != 'None':
                            shipment_empty = False
                            break

                if str(xlsworksheet['B{}'.format(y)].value) == 'Tracking Number':
                    endrow = y + 1
                    i = y + 1
                    break
            if shipment_empty == True:
                shipmentlist.append({'begin':startrow, 'end':endrow})
            else:
                shipreadylist.append({'begin':startrow, 'end':endrow})

    # print(json.dumps(shipmentlist))
    for index, shipmentdata in enumerate(shipmentlist):
        shipmentlist[index]['submitter'] = xlsworksheet['B{}'.format(shipmentdata['begin'])].value
        shipmentlist[index]['address'] = xlsworksheet['B{}'.format(shipmentdata['begin']+1)].value
        shipmentlist[index]['name'] = xlsworksheet['A{}'.format(shipmentdata['begin'])].value
        boxes = ('E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P')
        boxcount = 0
        for box in boxes:
            if xlsworksheet['{}{}'.format(box, shipmentdata['begin'])].value != None:
                boxcount += 1
            else:
                break
        if boxcount == 0:
            del shipmentlist[index]
            continue
        shipmentlist[index]['boxcount'] = boxcount
        start = shipmentdata['begin'] + 2
        shipmentlist[index]['weightboxes'] = []
        shipmentlist[index]['dimensionboxes'] = []
        shipmentlist[index]['nameboxes'] = []
        shipmentlist[index]['items'] = []

        # get weightboxes
        rowsearch = 0
        for i in range(start, shipmentdata['end']):
            if xlsworksheet['B{}'.format(i)].value == 'Weight':
                rowsearch = i
                break
        
        for ke, box in enumerate(boxes):
            if ke == boxcount:
                break
            shipmentlist[index]['weightboxes'].append(int(xlsworksheet['{}{}'.format(box, rowsearch)].value)) #UP

        # get dimensionboxes
        rowsearch = 0
        for i in range(start, shipmentdata['end']):
            if xlsworksheet['B{}'.format(i)].value == 'Dimensions':
                rowsearch = i
                break
        
        for ke, box in enumerate(boxes):
            if ke == boxcount:
                break
            shipmentlist[index]['dimensionboxes'].append(xlsworksheet['{}{}'.format(box, rowsearch)].value)

        #get nameboxes
        for ke, box in enumerate(boxes):
            if ke == boxcount:
                break
            shipmentlist[index]['nameboxes'].append(str(int(xlsworksheet['{}{}'.format(box, shipmentdata['begin'])].value)))

        ti = -1
        for i in range(start, shipmentdata['end']):
            ti += 1
            if xlsworksheet['A{}'.format(i)].value == None or str(xlsworksheet['A{}'.format(i)].value).strip() == '':
                break
            
            dict = {
                'id': xlsworksheet['A{}'.format(i)].value,
                'name': xlsworksheet['B{}'.format(i)].value,
                'total': int(xlsworksheet['C{}'.format(i)].value), #UP
                'expiry': str(xlsworksheet['D{}'.format(i)].value),
                'boxes':[],

            }

            shipmentlist[index]['items'].append(dict)
            for ke, box in enumerate(boxes):
                if ke == boxcount:
                    break
                if xlsworksheet['{}{}'.format(box, i)].value == None or str(xlsworksheet['{}{}'.format(box, i)].value).strip() == '':
                    shipmentlist[index]['items'][ti]['boxes'].append(0)
                else:                           
                    shipmentlist[index]['items'][ti]['boxes'].append(int(xlsworksheet['{}{}'.format(box, i)].value)) #UP
    shipids = []
    for shipmentdata in shipreadylist:
        boxes = ('E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P')
        boxcount = 0
        start = shipmentdata['begin'] + 2
        for box in boxes:
            
            if xlsworksheet['{}{}'.format(box, shipmentdata['begin'])].value != None:
                boxcount += 1
            else:
                break
        if boxcount == 0:
            del shipreadylist[index]
            continue

        rowsearch = 0
        for i in range(start, shipmentdata['end']):
            if xlsworksheet['B{}'.format(i)].value == 'Shipment ID':
                rowsearch = i
                break

        

        rowsearch2 = 0
        for i in range(start, shipmentdata['end']):
            if xlsworksheet['B{}'.format(i)].value == 'Tracking Number':
                rowsearch2 = i
                break

        for ke, box in enumerate(boxes):
            if ke == boxcount:
                break
            mdict = {
                'boxname':str(int(xlsworksheet['{}{}'.format(box, shipmentdata['begin'])].value)),
                'shipid': xlsworksheet['{}{}'.format(box, rowsearch)].value,
                'label': xlsworksheet['{}{}'.format(box, rowsearch2)].value

            }
            shipids.append(mdict)

    #cleansing
    idxdel = []
    for index, shipmentdata in enumerate(shipmentlist):
        try:
            cheat = shipmentdata['name']
        except:
            idxdel.append(index)
    
    for idx in idxdel:
        for index, shipmentdata in enumerate(shipmentlist):
            try:
                cheat = shipmentdata['name']
            except:
                del shipmentlist[index]
            
        # pass
    return shipids

def extract_pdf(download_folder, filename, datalist):
    pdffile = filename
    foldername = "{}{}combined".format(download_folder, file_delimeter() ) 
    isExist = os.path.exists(foldername)
    if not isExist:
        os.makedirs(foldername)

    white = fitz.utils.getColor("white")
    try:
        mfile = fitz.open(pdffile)
    except:
        return pdffile + " " + "file not found"
    reader = easyocr.Reader(['en'], gpu=False, verbose=False)

    tmpname = "{}{}{}.pdf".format(foldername, file_delimeter() , "tmp")

    for i in range(0, mfile.page_count):
        found = False
        pfound = 0
        dtrackid = ""
        dshipid = ""
        dboxname = ""
        page = mfile[i]
        page.set_rotation(90)
        for data in datalist:
            if data['shipid'] != None and len(data['shipid']) == 19:
                plist = page.search_for(data['shipid'])
                if len(plist) != 0:
                    found = True
                    pfound = i
                    dtrackid = data['label'].strip()
                    dshipid = data['shipid'].strip()
                    dboxname = data['boxname'].strip()
                    break

        if found:
            single = fitz.open()
            single.insert_pdf(mfile, from_page=pfound, to_page=pfound)
            single.save(tmpname)
            images = convert_from_path(tmpname)
            imgcrop = images[0].crop(box = (180,750,750,900))
            res = reader.readtext(numpy.array(imgcrop)  , detail = 0)
            tracking_id = res[1].strip().replace('TRACKING #','').replace(' ','').replace(":","")
            print(dshipid, dtrackid, "... ", end="", flush=True)
            if tracking_id == dtrackid:
                boxfilename = "{}{}{}.pdf".format(foldername, file_delimeter(),  dboxname)
                mfile2 = fitz.open(tmpname)
                page2 = mfile2[0]
                page2.insert_text((550.2469787597656, 100.38037109375), "Box:{}".format(str(dboxname)), rotate=90, color=white)
                page2.set_rotation(90)
                mfile2.save(boxfilename)
                mfile2.close()    
                print("Box", dboxname)
            else:
                print("not found")
                print(tracking_id, dtrackid, "not same")
        else:
            print(filename, "page:{}".format(i+1), "was not found in the excel file")
            continue
        try:
            os.remove(tmpname)
        except:
            pass

def main():
    parser = argparse.ArgumentParser(description="FDA PDF Extractor")
    parser.add_argument('-pdf', '--pdfinput', type=str,help="PDF File Input")
    parser.add_argument('-xls', '--xlsinput', type=str,help="XLSX File Input")
    parser.add_argument('-sname', '--sheetname', type=str,help="Sheet Name of XLSX file")
    parser.add_argument('-output', '--pdfoutput', type=str,help="PDF output folder")
    args = parser.parse_args()
        
    if not (args.xlsinput[-5:] == '.xlsx' or args.xlsinput[-5:] == '.xlsm'):
        input('input the right XLSX or XLSM file')
        sys.exit()

    filelist = args.pdfinput.replace("('", '').replace("')","").replace("',)","").split("', '")
    for idx, filename in enumerate(filelist):
        isExist = os.path.exists(filename.strip())
        if not isExist:
            input(filename.strip() + " does not exist")
            sys.exit()
        else:
            filelist[idx] = filename.strip()

    isExist = os.path.exists(args.xlsinput)
    if not isExist:
        input(args.xlsinput + " does not exist")
        sys.exit()
    isExist = os.path.exists(args.pdfoutput)
    if not isExist:
        input(args.pdfoutput + " folder does not exist")
        sys.exit()

    print('Opening the Source Excel File...', end=" ", flush=True)
    xlbook = xw.Book(args.xlsinput)
    xlsheet = xlbook.sheets[args.sheetname]
    print('OK')
    print("Data Mounting...", end=" ", flush=True)
    datalist = data_generator(xlsworksheet=xlsheet)
    print("OK")
    # input(datalist)
    allsavedfiles = []
    for filename in filelist:
        extract_pdf(download_folder=args.pdfoutput, filename=filename, datalist=datalist)
    resultfile = lib.join_pdfs(source_folder=args.pdfoutput + file_delimeter() + "combined" , output_folder = args.pdfoutput, tag='Labels')
    if resultfile != "":
        lib.add_page_numbers(resultfile)

    input("data generating completed...")

if __name__ == '__main__':
    main()


  