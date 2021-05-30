import pandas
# Create a workbook and add a worksheet.
import xlsxwriter
workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet()
PalletDoos=[]
b=0
BlikkenOpPallet=840


def totaal():
    Doos= input ("aantal dozen")
    BlikkenPerDoos= input ("aantal  blikken per doos")
    BlikkenOver= input ("aantal  blikken over")   
    import datetime
    now = datetime.datetime.now()
    datm = ("Datum "+now.strftime("%Y-%m-%d"))
    worksheet.write(10, 0, datm)
    worksheet.write(10, 2, '1e dag ')
    worksheet.write(10, 3, Doos)
    worksheet.write(9, 3, 'aantal dozen')
    worksheet.write(10, 4, 'X')
    worksheet.write(10, 5, BlikkenPerDoos)
    worksheet.write(9, 5, 'blikken per doos')
    worksheet.write(10, 6, '+')
    worksheet.write(10, 7, BlikkenOver)
    worksheet.write(9, 7, 'blikken over')
    worksheet.write(10, 8, 'pakjes per doos')
    worksheet.write(10, 10,'=SUM(D11*F11+H11)' )
    worksheet.write(10, 12, 'pakjes')
    workbook.close()
    Doos=int(Doos)
    BlikkenPerDoos=int(BlikkenPerDoos)
    BlikkenOver=int(BlikkenOver)
    totaal = Doos*BlikkenPerDoos+BlikkenOver
    print(f'{f"{datm} ":<12}{f"1e dag {Doos}":}{f"x{BlikkenPerDoos}":}{f"+({BlikkenOver}) pakjes per doos":}{f"={totaal} pakjes":<10}')
    return


            
    

   


    

def hello():
  a=0
  doosvol=0
  BlikkenPerPallet=0
  print("Hello World")
  aantal= input ("aantal  blikken maken")
  DoosOpPallet= input ("aantal dozen op pallet")
  BlikkenPerDoos= input ("aantal  blikken per doos")
  DoosOpPallet=int(DoosOpPallet)
  BlikkenPerDoos=int(BlikkenPerDoos)
  BlikkenPerPallet=DoosOpPallet * BlikkenPerDoos
  #hoeveel=DoosOpPallet*BlikkenPerDoos
  aantal=int(aantal)
  for c in range(aantal):
   a=a+1
   if a == BlikkenPerDoos :doosvol=doosvol + 1
   if a == BlikkenPerDoos :a=0
   
  print (a,'blikken over')
  print (doosvol,'dozen')
  #---------------------------
  b=0
  palletvol=0
  for c in range(doosvol):
   b=b+1
   if b == DoosOpPallet :palletvol=palletvol + 1
   if b == DoosOpPallet :b=0


  print (b,'dozen over')  
  print (palletvol,'pallets')
  
  print (BlikkenPerPallet,'blikken per pallet')

   # Some data we want to write to the worksheet.
  workbook = xlsxwriter.Workbook('Expenses01.xlsx')
  worksheet = workbook.add_worksheet()
    
  a=0
  nummer=0
  totaal=0
    
       
  for index in range(palletvol):
        nummer=nummer+1
        worksheet.write(nummer, 0,nummer)
        worksheet.write(nummer, 1,'e  pallet')
        worksheet.write(nummer, 2,BlikkenPerPallet,)
        worksheet.write(nummer, 3,'blikken')
        totaal=BlikkenPerPallet*nummer
    
  worksheet.write(1, 5,'blikken')
  worksheet.write(2, 5,totaal)
  worksheet.write(2, 6,'totaal')
        
  workbook.close()

def helloA():
  a=0
  doosvol=0
  BlikkenPerPallet=0
  print("Hello World")
  Doos= input ("aantal dozen")
  
  
  DoosOpPallet= input ("aantal dozen op pallet")
  BlikkenPerDoos= input ("aantal  blikken per doos")
  Doos=int(Doos)
  DoosOpPallet=int(DoosOpPallet)
  BlikkenPerDoos=int(BlikkenPerDoos)
  BlikkenPerPallet=DoosOpPallet * BlikkenPerDoos
  aantal=Doos * BlikkenPerDoos
  #hoeveel=DoosOpPallet*BlikkenPerDoos
  aantal=int(aantal)
  for c in range(aantal):
   a=a+1
   if a == BlikkenPerDoos :doosvol=doosvol + 1
   if a == BlikkenPerDoos :a=0
   
  print (a,'blikken over')
  print (doosvol,'dozen')
  #---------------------------
  b=0
  palletvol=0
  for c in range(doosvol):
   b=b+1
   if b == DoosOpPallet :palletvol=palletvol + 1
   if b == DoosOpPallet :b=0


  print (b,'dozen over')  
  print (palletvol,'pallets')
  
  print (BlikkenPerPallet,'blikken per pallet')


      
  return

def laatZien():
  excel_data_df = pandas.read_excel('Expenses01.xlsx')
  # print whole sheet data
  print(excel_data_df)



while True:
    a=0
    bb=0
    for index in range(len(PalletDoos)):
        a=a+BlikkenOpPallet
        bb=bb+1
       
        print(bb,"e pallet",BlikkenOpPallet," blikken")
        
        
    print ("totaal ",a)

    print ("(1) totaal (2) add pallet (3) remove pallet (4) save")
    
    Vraag= input ("wat wil je doen")
    if Vraag == "1":totaal()
    if Vraag == "2":pallets(PalletDoos,b)
    
    if Vraag == "5":hello()
    if Vraag == "6":helloA()
    if Vraag == "7":laatZien()
    
   




