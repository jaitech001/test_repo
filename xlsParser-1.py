#!/usr/bin/python3

import json
import xlrd
import os.path

def xlsParse(xlsfile):
    store = []  #List to store namespace objects
    framenames = []
    namespaceMappings = {}
    book = xlrd.open_workbook(xlsfile)
    sh1 = book.sheet_by_index(0)

    for rx in range(1, sh1.nrows):
        if sh1.row(rx)[0].value not in framenames:
            framenames.append(sh1.row(rx)[0].value)
            frame = {"namespace": sh1.row(rx)[0].value,
               "attributes":[]
               }
            block1 = {
                sh1.row(rx)[0].value: {
                    "duration":300,
                    "dsSysnm" : "RESNAME",
                    "timestamp": "TIME",
                    "timeformat": "yyyy-mm-dd'T'HH:mm:ss.SSSz",
                    "metricMappings" : {}
                    }
                }
                
            store.append(frame)
            namespaceMappings.update(block1)

    
    attribute = {"attributeName":""}
    for frame in store:
        for rx in range(1, sh1.nrows):
            if frame["namespace"] == sh1.row(rx)[0].value:
                if attribute["attributeName"] != sh1.row(rx)[1].value:
                    key1 = sh1.row(rx)[0].value
                    key2 = sh1.row(rx)[1].value
                    attributeList = {}
                    attribute = {
                    "attributeName": sh1.row(rx)[1].value
                    }
                    frame["attributes"].append(attribute)

                    block2 = {
                        "tscoMetric": sh1.row(rx)[3].value,
                        "tscoScale": int(sh1.row(rx)[4].value),
                        "extended": sh1.row(rx)[5].value
                        }
 
                    attributeList[key2]=block2
                    namespaceMappings[key1]['metricMappings'].update(attributeList)
    
    out = json.dumps(store, indent=4)
    f = open( 'xlsdata.json', 'w')
    f.write(out)
    return namespaceMappings

def main():
    filename = input("Enter the path to the filename ==> ")
    if os.path.isfile(filename):
        xlsParse(filename)
        namespaceMappings = xlsParse(filename)
        output = \
        open((filename.replace("xlsx", "json")).replace("xls", "json"), "w")
        output.write(json.dumps(namespaceMappings, indent=4))
        output.close()
        print ("%s was created" %output.name)
    else:
        print ("Sorry, that was not a valid filename")


main()
