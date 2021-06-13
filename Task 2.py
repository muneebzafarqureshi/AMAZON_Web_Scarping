import pandas as pd #Using Pandas library to show the selected category 

dviz = pd.read_excel("DVIZ Test.xlsx") #reading the excel file, you can change the location of the excel as in my case the python file and excel file both are in same location

#all the categories under COMPUTER section
print('Computer Accessories and Peripherals\nComputer Components\nComputers & Tablets\nData Storage\nLaptop Accessories\nMonitors\nNetworking Products\nPower Strips & Surge Protectors\nPrinters\nScanners\nServers\nTablet Accessories\nTablet Replacement Parts\nWarranties & Services\n')

print("***Enter Correct Spelling\'s of Category***") #instructions

item = input("Enter Product Category : ") #taking input from user,(the user have to enter the category name like Scanners, Monitors etc in correct form(case sensitive))

#by uncommenting the below line, you can view the selected category data in dataframe form
#dviz[dviz["Product Categories"] == str(item)] #showing data of selected product category

#saving the specific product category data in view variable
view = dviz[dviz["Product Categories"] == str(item)]

#converting the specific product category into list
con = view["Product\'s"]
to_list = con.values.tolist()
print("\nLength of LIST is : "+str(len(to_list))+"\n")
print(to_list)