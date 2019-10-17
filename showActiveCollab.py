import xlrd 
import psycopg2
from prettytable import PrettyTable
from datetime import time
import datetime
from datetime import timedelta  
from datetime import datetime

'''Creating Excel class '''
class Excel():
    ''' Initializing '''
    def __init__(self,sheet):   
        self.sheet = sheet

    ''' Function to convert excel_date to timestamp format '''
    def basicCalculationTimeOnly(self,excel_date):
        if(excel_date==0):
            return "null"
        dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(excel_date) - 2)
        hour, minute, second = self.floatHourToTime(excel_date % 1)
        dt = dt.replace(hour=hour, minute=minute, second=second)
        Created_On = "' %s'" % (dt.strftime("%m/%d/%Y"))
        return Created_On

    '''Function to convert excel_date to hh,mm,ss '''
    def floatHourToTime(self,fh):
        h, r = divmod(fh, 1)
        m, r = divmod(r*60, 1)
        return (
        int(h),
        int(m),
        int(r*60),
        )

    '''Initializing process to store excel sheet in Arraylist '''
    def process(self,sheet): 
        ''' Initializing an ArrayList '''
        ArrayList = []     
        ArrayList.clear()
        '''processing an arraylist to store null values at blank cells in excel sheet and then converting to string format '''
        for y in range(sheet.nrows):
            for z in range(sheet.ncols):
                if((sheet.cell(y,z).ctype)==(xlrd.XL_CELL_DATE)):
                    k=sheet.cell(y,z).value
                    sheet._cell_values[y][z] = self.basicCalculationTimeOnly(k)
                elif(sheet.cell_type(y,z)==0):
                    sheet._cell_types[y][z] = xlrd.XL_CELL_NUMBER
                    sheet._cell_values[y][z] = "null"
                    sheet._cell_values[y][z] = str(sheet._cell_values[y][z]) 
                elif((sheet.cell(y,z).ctype)==(xlrd.XL_CELL_NUMBER)):
                    sheet._cell_values[y][z] = str(sheet._cell_values[y][z])
                elif(sheet.cell_type(y,z)==1):
                    sheet._cell_values[y][z]="'%s'"  % sheet._cell_values[y][z]

        '''appending the processed sheet into Arraylist '''
        for l in range(1,sheet.nrows):
            ArrayList.append(sheet.row_values(l))
        ''' returning ArrayList '''
        return ArrayList   

'''Writing to database '''
class writeToDb():
    def __init__(self,ArrayList,table):
        self.ArrayList = ArrayList
        self.table = table
        ''' Starting a connection '''
        try:
            connection = psycopg2.connect(user = "postgres",
                                          password = "Click@123",
                                          host = "10.0.3.25",
                                          port = "5432",
                                          database = "activecollab")
            cursor  = connection.cursor()

            print ( connection.get_dsn_parameters(),"\n")
            iteratingIndex = 0
            for f in ArrayList: 
                print(iteratingIndex)   
                query=("INSERT INTO public."+table+" " "VALUES("+(",".join(f))+")") 
                print(query)
                iteratingIndex = iteratingIndex + 1
                cursor.execute(query)
                connection.commit()

        finally:
                '''closing database connection.'''
                if(connection):
                    cursor.close()
                    connection.close()
                    print("PostgreSQL connection is closed")

if __name__ == ""__main__":
filesToProcess={"C:/Users/RST014/Videos/sample/sample/billable_projects.xlsx" :
    { "Project Billing Status":'"project billing status"',
      "Project Default Tasks":'"project default tasks"'},
      "C:/Users/RST014/Videos/sample/sample/assetpurchase.xlsx" : {"assetpurchase":'"assetpurchase"'}
                }

ArrayList = []
for files, totalSheets in filesToProcess.items():
    print("\nPerson ID:", files)
    
    
    for key in totalSheets:
        '''Storing sheet in workbook '''
        workBook = xlrd.open_workbook(files)
        sheet_names = workBook.sheet_names()
        print(key + '=>' + totalSheets[key])
        sheet = workBook.sheet_by_name(key) 
        ''' Initializing an excel object  '''
        ExcelObj = Excel(sheet)
        ''' Calling process function in excel class '''
        ArrayList = ExcelObj.process(sheet)
        ''' Calling writeToDb class '''
        writeToDb(ArrayList,totalSheets[key])
        ''' Clearing an arrayList at the end of every iteration '''            
        ArrayList.clear()

