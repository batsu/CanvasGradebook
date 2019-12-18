import xlrd
import xlwt
from scale import scale

filename = input("Name of Excel file (do not include .xlsx)? ")

loc = (filename+".xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
global col_list
col_list = []


#Assignment counter declaration
assn_count = 0

#Search Function
def search():
    colval = []
    lookfor = input("Assignment name? ")
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if (lookfor) in str(sheet.cell_value(row,col)):
                colval.append(col)
    #Checking for uniqueness of assignment
    if len(colval) > 1:
        print("Many assingments found with same name! Choose from the following.")
        for i in colval:
            print(i,":",sheet.cell_value(0,i))
        colval = [0]
        colval[0] = input("Select the number value for your assignment: ")
    #If string doesn't exists
    if colval == []:
        print("This assignment doesn't exist.")
        colval = []
        search()
    return(colval)

#Gather Students Function
def gather(period):
    rowlist = []
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if ("Period: "+period) in str(sheet.cell_value(row, col)):
                rowlist.append(row)
    return(rowlist)

#Writing Function
def create_scale(col_val,period,rowlist,scale_val,wb,total_val,assn_count):
    Sheet1 = wb.add_sheet("Per. "+period)

    Sheet1.write(0,0,"Students")
    Sheet1.write(0,1,"ID")

    #Make Counter for Printing Scores
    write_counter = int(assn_count)

    #Students Print
    x = 0
    for i in rowlist:
        Sheet1.write(1+x,0,sheet.cell_value(i,0))
        x += 1
    
    #ID Number
    x = 0
    for i in rowlist:
        Sheet1.write(1+x,1,sheet.cell_value(i,2))
        x += 1

    
    x = 0
    #Writing ALL SCORES
    for m in range(1, len(col_list)+1):
        #Raw Score
        x = 0
        for i in rowlist:
            Sheet1.write(1+x,2*m+1,sheet.cell_value(i,col_list[m-1]))
            x += 1
        #Missing Score
        x = 0
        for i in rowlist:
            if sheet.cell_value(i,col_list[m-1]) == '':
                Sheet1.write(1+x,2*m,'no score')
            else:
                #Scaled Score
                y = float(sheet.cell_value(i,col_list[m-1]))
                y = scale(y * 100 / int(scale_val[m-1]))
                y = float(y / 100 * float(total_val[m-1]))
                Sheet1.write(1+x,2*m,round(y, 1))
            x += 1
        Sheet1.write(0,2*m,"Scaled")
        Sheet1.write(0,2*m+1,sheet.cell_value(0,col_list[m-1]))
    
    

    x = 0



def all_together(assn_count, curve_val, total_val):


    #Search
    col_val = search()

    #Assignment Column Value
    col_use = int(col_val[0])
    col_list.append(col_use)

    #Points Possible
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if ("Points Possible") in str(sheet.cell_value(row,col)):
                pointposs = row
    curve_val.append(int(sheet.cell_value(pointposs,col_use)))

    #Curve setting
    print("Current points possible is: "+str(sheet.cell_value(pointposs,col_use)))
    choice = input("Is there a curve you want to set (Y or N)? ")
    if (choice == 'Y') or (choice == 'y'):
        curve_val[len(curve_val)-1] = int(input("New Curve Value: "))
    choice = input("Do you want a score other than 100 points possible for Q (Y or N)? ")
    if (choice == 'Y') or (choice == 'y'):
        total_val.append(int(input("New points possible: ")))
    else:
        total_val.append(100)
                      
    new_assn = input("Would you like to add another assignment (Y or N)? ")
    if (new_assn == "Y") or (new_assn == "y"):
        assn_count += 3
        wb = all_together(assn_count, curve_val, total_val)

    #Get students
    rowlist_p1 = gather("1")
    rowlist_p2 = gather("2")
    rowlist_p3 = gather("3")
    rowlist_p4 = gather("4")
    rowlist_p5 = gather("5")
    rowlist_p6 = gather("6")
    #Create files
    wb = xlwt.Workbook()
    create_scale(col_use,"1",rowlist_p1,curve_val,wb,total_val,assn_count)
    create_scale(col_use,"2",rowlist_p2,curve_val,wb,total_val,assn_count)
    create_scale(col_use,"3",rowlist_p3,curve_val,wb,total_val,assn_count)
    create_scale(col_use,"4",rowlist_p4,curve_val,wb,total_val,assn_count)
    create_scale(col_use,"5",rowlist_p5,curve_val,wb,total_val,assn_count)
    create_scale(col_use,"6",rowlist_p6,curve_val,wb,total_val,assn_count)
    wb.save("temp.xls")


    return(wb)


#Start the program!
curve_val = []
total_val = []
wb = all_together(assn_count, curve_val, total_val)

#Saving the final excel file
file_name = input("Name of output file: ")
wb.save(file_name+".xls")
print(file_name+".xls")
