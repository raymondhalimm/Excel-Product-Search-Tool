import pandas as pd


storedata = []
storename = []

def intro():
    global item
    item = input("Enter the name of the item: ")
    item.strip()
    
  
def timeframe():
    global time
    global months
    global year
    global check
    time = input("Do you want to check by Year or Month? ")
    time = time.upper()
    while time.strip() != "YEAR" and time.strip() != "MONTH":
        print("Please enter only Year or Month. ")
        time = input("Do you want to check by Year or Month? ")
        time = time.upper()
        
       
        
    if time == "MONTH":
        
        while True:
            try:
                year = int(input("Which year would you like to check? "))
                
            except ValueError:
                print("Please input a number.")
                continue
            
            if year > 2024 or year < 2019:
                print("Please choose between 2019 and 2023.")
                continue
            
            else:
                break
            
        while True:
            try:
                months = int(input("Which month do you want to check? "))
                
            except ValueError:
                print("Please input a number")
                continue
            
            if months > 12 or months < 1:
                print("Please choose between 1 to 12.")
                continue 
            
            else:
                break
            
        yearmonth(year)
        
    elif time == "YEAR":
        while True:
            try:
                years = int(input("Which year do you want to check? "))
                
            except ValueError:
                print("Please input a number")
                continue
            
            if years > 2024 or years < 2019:
                print("Please input between 2019 and 2023")
                continue
            
            
            else:
                break
            
        year(years)

def year(x):
    if x == 2020:
        for x in range (1,13):
            month2020(x)
    elif x == 2019:
        for x in range (7,13):
            month2019(x)
    elif x == 2021:
        for x in range (1,13):
            month2021(x)
    elif x == 2022:
        for x in range (1, 13):
            month2022(x)
    elif x == 2023:
        for x in range (1, 12):
            month2023(x)
    
def yearmonth(y):
    if y == 2019:
        month2019(months)
    elif y == 2020:
        month2020(months)
    elif y == 2021:
        month2021(months)
    elif y == 2022:
        month2022(months)
    elif y == 2023:
        month2023(months)
        
def month2019(m):
    global path2
    if m == 7:
        path2 = "/371-379 July 2019/Bon July 2019 (SSR,HBM,SIR,BTR).xlsx"
    elif m == 9:
        path2 = "/390-397 Sept 2019/Bon Sept 2019 (SIR,SSR,IAL).xlsx"
    elif m == 10:
        path2 = "/398-403 Oct 2019/Bon Oct 2019 (SIR,SSR,BJP).xlsx"
    elif m == 11: 
        path2 = "/404-410 Nov 2019/Bon Nov 2019 (SIR,SSR,BTR,BTKJ).xlsx"
    elif m == 12:
        path2 = "/411-416 Dec 2019/Bon Dec 2019 (HBM,SIR,SSR,SJI).xlsx"
    filepath(path2)
    sheetname()
        
def month2020(m):
    global path2
    if m == 1:
        path2 = "/417-422 January 2020/Bon January 2020 (SIR,SSR).xlsx"
    elif m == 2:
        path2 = "/423-428 February 2020/Bon February 2020 (SIR,SSR,BTR).xlsx"
    elif m == 3:
        path2 = "/429-435 March 2020/Bon March 2020 (SIR,SSR,GPH,HBM).xlsx"
    elif m == 4:
        path2 = "/436-443 April 2020/Bon April 2020 (SIR,GPH,SUNSAWIT,SSR,RETUR,STOCK).xlsx"
    elif m == 5:
        path2 = "/444-447 May & June 2020/Bon May 2020 (SIR).xlsx"
    elif m == 6:
        path2 = "/444-447 May & June 2020/Bon June 2020 (SIR,SSR,GPH).xlsx"
    elif m == 7:
        path2 = "/448-451 July 2020/Bon July 2020 (SIR,SJI,SSR).xlsx"
    elif m == 8:
        path2 = "/452-458 August 2020/Bon August 2020 (RAS,KWP,SIR,SJI).xlsx"
    elif m == 9:
        path2 = "/459-464 Sept 2020/Bon Sept 2020 (RAS,SJI,BTR,SSR).xlsx"
    elif m == 10:
        path2 = "/465-471 Oct 2020/Bon October 2020 (RAS,BTR,SSR,KWP,SJI).xlsx"
    elif m == 11: 
        path2 = "/472-479 Nov 2020/Bon Nov 2020 (RAS,GPH,KAS,KWP).xlsx"  
    elif m == 12:
        path2 = "/480-489 Dec 2020/Bon December 2020 (RAS,TAL,BTR,KAS,SSR,STOCK).xlsx"
    filepath(path2)
    sheetname()

def month2021(m):
    global path2
    if m == 1:
        path2 = "/490-498 Jan 2021/Bon January 2021 (RAS,GPH,SSR,KAS,TAL,SIR,BTR).xlsx"
    elif m == 2:
        path2 = "/499-505 Feb 2021/Bon February 2021 (GPH,RAS,BTR,KAS,STOCK).xlsx"
    elif m == 3:
        path2 = "/506-512 March 2021/Bon March 2021 (RAS,BOS,GPH,KAS).xlsx"
    elif m == 4:
        path2 = "/513-521 April 2021/Bon April 2021 (RAS,GPH,KAS,SUN,AP,PAS).xlsx"
    elif m == 5:
        path2 = "/522-527 May 2021/Bon May 2021 (KAS,RAS,KA,GPH,AP).xlsx"
    elif m == 6:
        path2 = "/528-532 June 2021/Bon June 2021 (KA,GPH,RAS,SSR).xlsx"
    elif m == 7:
        path2 = "/533-538 July 2021/Bon July 2021(HBM,RAS,GPH,SSR,STOCK,PTPN).xlsx"
    elif m == 8:
        path2 = "/539-546 August 2021/Bon August 2021(RAS,SSR,TUMPUAN,PAT,HBM).xlsx"
    elif m == 9:
        path2 = "/547-554 September 2021/Bon Sept 2021 (RAS,SSR,SUNSAWIT,PAT,TUMPUAN).xlsx"   
    elif m == 10:
        path2 = "/555-559 October 2021/Bon October 2021 (RAS,TUMPUAN,BOS).xlsx"
    elif m == 11:
        path2 = "/560-567 November 2021/Bon November 2021 (TMP,RAS,GPH,SS,BOS).xlsx"
    elif m == 12:
        path2 = "/568-574 December 2021/Bon December 2021 (RAS,TMP,HBM,GPH).xlsx"

    filepath(path2)
    sheetname()

def month2022(m):
    global path2
    if m == 1:
        path2 = "/575-581 January 2022/Bon January 2022 (RAS,GPH,HBM,SS,BOS,TMP).xlsx"
    elif m == 2:
        path2 = "/582 February 2022/Bon February 2022 (RAS, STOCK).xlsx"
    elif m == 3:
        path2 = "/583-586 March 2022/Bon March 2022 (PMG,RAS,GPH,BTKJ,STOCK).xlsx"
    elif m == 4 :
        path2 = "/587-593 April 2022/Bon April 2022(RAS,GPH,SS,PMG,SSR,BTKJ,STOCK).xlsx"
    elif m == 5 :
        path2 = "/594-597 May 2022/Bon May 2022 (PMG,TMP,RAS,UNIMAS).xlsx"
    elif m == 6 :
        path2 = "/598-604 June 2022/Bon June 2022(RAS,GPH,APL,TMP,BTKJ,STOCK).xlsx"
    elif m == 7 :
        path2 = "/605-607 July 2022/Bon July 2022(TMP,BTKJ,RAS,STOCK).xlsx"
    elif m == 8 :
        path2 = "/608-614 August 2022/Bon August 2022(TMP,BTKJ,HGE,RAS,BKP,PAT).xlsx"
    elif m == 9 :
        path2 = "/615-622 September 2022/Bon September 2022(RAS,TMP,BKP,SS,SSR,BTKJ).xlsx"
    elif m == 10 :
        path2 = "/623-630 October 2022/Bon October 2022(SS,BKP,TMP,BOS,BTKJ,HBM).xlsx"
    elif m == 11 :
        path2 = "/631-635 November 2022/Bon November(BTKJ,TMP,BKP,STOCK).xlsx"
    elif m == 12 :
        path2 = "/636-643 December 2022/Bon Dec 2022 (BOS,BKP,APL,TMP,BTKJ,GPH,STOCK).xlsx"
        
    filepath(path2)
    sheetname()

def month2023(m) :
    
    global path2
    if m == 1:
        path2 = "/644-650 January 2023/Bon Jan 2023(BTKJ,BKP,TMP,PMG).xlsx"
    elif m == 2:
        path2 = "/651-655 February 2023/Bon February 2023(BKP,TMP,BTKJ).xlsx"
    elif m == 3:
        path2 = "/656-664 March 2023/Bon Maret 2023(BKP,SBL,PMG,BTKJ,TMP).xlsx"
    elif m == 4:
        path2 = "/665-672 April & May 2023/Bon April dan May(SBL,BKP,TMP,BTKJ).xlsx"
    elif m == 6 : 
        path2 = "/673-680 June 2023/Bon June 2023(SBL,BKP,TMP,RP).xlsx"
    elif m == 7 :
        path2 = "/681-688 July 2023/Bon July 2023(BKP,SBL,TMP,RP).xlsx"
    elif m == 8 :
        path2 = "/689-695 August 2023/Bon August 2023(BKP,TMP,BTKJ,RP).xlsx"
    elif m == 9 :
        path2 = "/696-702 September 2023/Bon Sept 2023(BKP,SBL,RP).xlsx"
    elif m == 10 :
        path2 = "/703-713 October 2023/Bon October 2023(BKP,SBL,RP,GPH,RAS,TMP).xlsx"
    elif m == 11 :
        path2 = "/714-721 November 2023/Bon November 2023(BKP,TMP,SBL,RP,GUDANG,STOCK).xlsx"
        
    filepath(path2)
    sheetname()    
    
def filepath(a):
    global path
    path = "/Users/raymondhalim/Desktop/BJP"
    path = path + a
     
def sheetname():
    global sheet
    sheetlist = pd.ExcelFile(path)
    sheets = sheetlist.sheet_names
    sheets.pop(-1)
    go = "YES"
    
    if time.strip() == "YEAR":
        for x in sheets:
            sheet = x
            extractfile()
    
    else:
        while go.upper() == "YES":
            print(sheets)
            sheet = str(input("Which sheet do you want to check or all of them? "))
             
               
            if sheet.upper() == "ALL":
                for x in sheets:
                    sheet = x
                    extractfile()
                break
            
            while sheet not in sheets:
                print("Please choose from the given list.")
                sheet = str(input("Which sheet do you want to check or all of them? "))        
            
            extractfile()
            index = sheets.index(sheet)
            sheets.pop(index)
            
            if len(sheets) == 0:
                print("There is no more available sheet to check, you will be directed to the result.")
                go = "NO"
                
            else: 
                go = input("Do you want to check other sheets? ")
                while go.upper() != "YES" and go.upper() != "NO":
                    print("Please input just Yes or No.")
                    go = input("Do you want to check other sheets? ")
                    go.strip()
                
def extractfile(): 
    global bjpfile
    global itemlist
    cols = [2]
    bjpfile = pd.read_excel(path, header=5, nrows=31, usecols=cols, sheet_name=sheet, index_col=False)
    itemlist = bjpfile["Nama Barang"].tolist()
    for x in itemlist:
        if item.upper() in str(x).upper():
            storedata.append(x)
            storename.append(sheet)
            continue
        elif x != str:
            continue
    
def result():
    global count
    count = len(storedata)
    if count == 0 :
        if time == "MONTH":
            print("Item not found in " +str(months) + "/" +str(year))
    start = 0
    while start < count:
        print("Found " +storedata[start]+ " in " +storename[start]+ ".")
        start+=1
    
def goAgain():
    if time.strip() == "MONTH":
        again = "YES"

        again = input(("Do you want to check other month?: "))
        again = again.upper()
        while again.strip() == "YES":
            if count != 0:
                storedata.clear()
                storename.clear()
                
            month2 = int(input("Which month do you want to check? "))           
            ans = input("Is it in the same year? ")
            ans = ans.upper()
            
            if ans.strip() == "YES":
                if year == 2020:
                    month2020(month2)
                elif year == 2019:
                    month2019(month2)
                    
            else:   
                year2 = int(input("Which year would you like to check?"))
                if year2 == 2020:
                    month2020(month2)
                elif year2 == 2019:
                    month2019(month2)
                    
            result()
            again = input(("Do you want to check other month?: "))
            again = again.upper()        
            if again.strip() == "NO":
                break
                   
intro() 
timeframe()
result()
goAgain()