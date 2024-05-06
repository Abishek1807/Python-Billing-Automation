def centerStatement():
    import os
    from xlwt import Workbook

    directory = input("Folder Location: ")
    centerNameDate = input("Center name and bill Period (This will be the name of the Excel file): ")

    def readingData(fileName):
        #directory la iruka individual files ah read panni, extracting data
        with open(fileName) as file:
            lines = file.readlines()
            top = lines[5].split()
            bottom = lines[-1].split()
            code = top[1]
            name = top[3]
            amount = bottom[-1]
            return {'code':int(code), 'name':name, 'amount':float(amount)}
    # In which format we need data being set above
    # Storing data from all the files (in directory in this "overall list")

    def writeToExcelSheet(data,loc,centerName):
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet 1')
        sheet1.write(0,0,"Code")
        sheet1.write(0,1,"Name")
        sheet1.write(0,2,"Amount")

        for index, entry in enumerate(data):
            sheet1.write(index+1, 0, entry["code"])
            sheet1.write(index+1, 1, entry["name"])
            sheet1.write(index+1, 2, entry["amount"])
        wb.save(loc + "\XX" +  centerName + '.xls')

    # ----------- Statement is Done with here ---------------

    def delExtraLine(folderLoc):

        def deleteLine(fileName):
            with open(fileName) as file:
                lines = file.readlines()
                # Checking for duplicate line
                if lines[-3] == lines[-4]:
                    del lines[-3]
                with open(fileName, 'w') as file:
                    for line in lines:
                        file.write(line)

        for i in os.listdir(folderLoc):
            if f"{folderLoc}/{i}".endswith('.txt'):
                deleteLine(f"{folderLoc}/{i}")

    # -------------- Extra Line deletion done --------------

    def billFill(fileLoc):
        for i in os.listdir(fileLoc):
            if f"{fileLoc}/{i}".endswith('.txt'): 
                with open(f"{fileLoc}/{i}", 'a') as file:
                    file.write("--------------------------------------------------------------------------------")
                    file.write("\n")
                    file.write("\n       Additions                                         Deductions ")
                    file.write("\n   ------------------                              --------------------")
                    file.write("\n Prem/penalty.(+/-)  : 0.00                       Advance amount  : 0.00")
                    file.write("\n Vol incentive       : 0.00                       Bank loan       : 0.00")
                    file.write("\n Special incentive am: 0.00                       Dairy loan      : 0.00")
                    file.write("\n Cartage  amount     : 0.00                       Cattle feed     : 0.00")
                    file.write("\n Recovery from trans : 0.00                       Medicines       : 0.00")
                    file.write("\n                                                  Dairy products  : 0.00")
                    file.write("\n                   ------------                                 -----------")
                    file.write("\n additions  total    : 0.00                       deductions      : 0.00")
                    file.write("\n                   ------------                                 -----------")
                    file.write("\n")
                    file.write("\n")
                    file.write("\n")
                    file.write("\n Authorised  signature                        Farmer Representative signature")
    # ----------------- filling additions/ deductions to the bill done --------------------

    overall = []
    for i in os.listdir(directory):
        overall.append(readingData(f"{directory}/{i}"))

    writeToExcelSheet(overall, directory, centerNameDate)

    delExtraLine(directory)

    billFill(directory)
