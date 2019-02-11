import tkinter
from tkinter import Checkbutton
import pyodbc
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from tkinter import Listbox
from tkinter import ttk
import PIL
from PIL import ImageTk
from PIL import Image
pyodbc.pooling = False
import os


def openDb():
    global dblocation, conn, tableNames, cursor,table_name,c,cc
    dblocation = askopenfilename(filetypes=[("Acces", "*.mdb")])
    lbPK=[]
    lPk=[]
    # print(dblocation)
    try:
        
        conn = pyodbc.connect(
            'DRIVER={Microsoft Access Driver (*.mdb)};UID=admin;UserCommitSync=Yes;Threads=3;SafeTransactions=0;PageTimeout=5;MaxScanRows=8;MaxBufferSize=2048;FIL={MS Access};DriverId=25;DBQ='+dblocation)
        cursor = conn.cursor()
        c = conn.cursor()
        cc= conn.cursor()

        fkcursor = conn.cursor()
        tableNames = [x[2] for x in cursor.tables().fetchall() if x[3] == 'TABLE']
        
        #for row in fkcursor.statistics(table='Register'):
        #    print(':::::::: |||| '+row.table_name)
        #    print(':::::::: '+str(row.non_unique))
        #    print(':::::::: '+str(row.index_qualifier))
        #    print(':::::::: '+str(row.index_name))
        #    print(':::::::: '+str(row.column_name))
        #crsr = conn.cursor()    
        #pk_cols = {row[7]: row[8] for row in crsr.statistics('Employee') if row[5]=='PrimaryKey'}
        #print(pk_cols)
        
        listTables.delete(0, 'end')
        listTablesPK.delete(0, 'end')
        
        for x in tableNames:
            table_columns_pk=[]
            table_columns_uniq=[]
            table_columns_datatype=[]
            
            i=0
            for row in c.statistics(table=x):
                try:
                    table_name=str(row.index_qualifier)
                    
                    if str(row.index_name) == 'PrimaryKey':
                        
                        #print('Nume Tabel:::::: '+str(row.table_name))
                        #print('NUME coloana PK:' +str(row.index_name))
                        
                        table_columns_pk.append(row.column_name)
                        #print(row.column_name)
                        if row.non_unique == 0:
                            table_columns_uniq.append('TRUE')
                        else:
                            table_columns_uniq.append('FALSE')
                        column_name= row.column_name
                        
                        for r in cc.columns():
                            if r.column_name == column_name:
                                table_columns_datatype.append(r.type_name)
                                #print(r.type_name)
                    else:
                        i+=1

                except Exception as e:
                    print(e)
                    messagebox.showwarning(
                        'Error', "Something wrong with DataBase!")
            
            table_description=table_name+' [PKC]: '
            i = len(table_columns_pk)
            ii=0
            addBoolean=0
            
            while ii<i:
                addBoolean=1
                #table_description += ' PKC: '+str(table_columns_pk[ii])+'[ UNIQUE: '+str(table_columns_uniq[ii])+', '+str(table_columns_datatype[ii])+']'
                table_description +=' '+str(table_columns_pk[ii])+' [ DATA_TYPE: '+str(table_columns_datatype[ii])+'],   '
                ii+=1
            #print('|||||||||||||:::::::::      '+table_description)
            if len(table_columns_pk) == 1 and str(table_columns_uniq[0]) == 'TRUE' and str(table_columns_datatype[0]) == 'COUNTER':
                
                #if addBoolean:
                #   table_description += ' )'
                   
                try:
                    cursor.execute('select * from '+x)
                    #print(cursor.description[0][0])
                    #column = cursor.description[0][0]
                    cursor.execute(
                    'ALTER TABLE '+table_name+' ADD CONSTRAINT PK_ID PRIMARY KEY('+str(table_columns_pk[0])+')')
                    conn.commit()
                    cursor.execute('ALTER TABLE '+table_name +
                               ' DROP CONSTRAINT PK_ID')
                    conn.commit()
               
                    #listTables.insert('end', x)
                    lbPK.append(x)
                except:
                    #listTablesPK.insert('end', table_description)
                    lPk.append(table_description)
            else:
                #if addBoolean:
                #    table_description += ' )'
                #listTables.insert('end', table_description)
                lbPK.append(table_description)
            
    except Exception as e:
        print(e)
        messagebox.showwarning(
            'Error', "Something wrong with DataBase!")
    lbPK.sort()
    lPk.sort()
    for x in lbPK:
        listTables.insert('end',x)
    for x in lPk:
        listTablesPK.insert('end',x)
        
def addPK():

    try:
        # conn.commit()
        # cursor = conn.cursor()
        tablename = listTables.get(listTables.curselection())
        cursor.execute('ALTER TABLE '+tablename+' ADD ID LONG INTEGER')
        conn.commit()
        cursor.execute('ALTER TABLE '+tablename +
                       ' ADD CONSTRAINT PK_ID PRIMARY KEY(ID)')
        conn.commit()
        cursor.execute('ALTER TABLE '+tablename +
                       ' ALTER COLUMN ID COUNTER')
        conn.commit()
        #print(listTables.get(listTables.curselection()))
    except Exception as e:
        # print(dblocation)
        print(e)
        messagebox.showwarning(
            'Error', "Please select a table from 'Tables without PK'!")


def refreshTables():
    lbPK=[]
    lPk=[]
    try:
        
        
        cursor = conn.cursor()
        c = conn.cursor()
        cc= conn.cursor()
        fkcursor = conn.cursor()
        tableNames = [x[2] for x in cursor.tables().fetchall() if x[3] == 'TABLE']
        
        #for row in fkcursor.statistics(table='Register'):
        #    print(':::::::: |||| '+row.table_name)
        #    print(':::::::: '+str(row.non_unique))
        #    print(':::::::: '+str(row.index_qualifier))
        #    print(':::::::: '+str(row.index_name))
        #    print(':::::::: '+str(row.column_name))
        #crsr = conn.cursor()    
        #pk_cols = {row[7]: row[8] for row in crsr.statistics('Employee') if row[5]=='PrimaryKey'}
        #print(pk_cols)
        
        listTables.delete(0, 'end')
        listTablesPK.delete(0, 'end')
        
        for x in tableNames:
            table_columns_pk=[]
            table_columns_uniq=[]
            table_columns_datatype=[]
            
            i=0
            for row in c.statistics(table=x):
                try:
                    table_name=str(row.index_qualifier)
                    
                    if str(row.index_name) == 'PrimaryKey':
                        
                        #print('Nume Tabel:::::: '+str(row.table_name))
                        #print('NUME coloana PK:' +str(row.index_name))
                        #print('Coloana: '+str(row.column_name))
                        table_columns_pk.append(row.column_name)
                        #print(row.column_name)
                        if row.non_unique == 0:
                            table_columns_uniq.append('TRUE')
                        else:
                            table_columns_uniq.append('FALSE')
                        column_name= row.column_name
                        
                        for r in cc.columns():
                            if r.column_name == column_name:
                                table_columns_datatype.append(r.type_name)
                                #print(r.type_name)
                    else:
                        i+=1

                except Exception as e:
                    print(e)
                    messagebox.showwarning(
                        'Error', "Something wrong with DataBase!")
            
            table_description=table_name
            i = len(table_columns_pk)
            ii=0
            addBoolean=0
            while ii<i:
                addBoolean=1
                table_description += ' PKC: '+str(table_columns_pk[ii])+'[ UNIQUE: '+str(table_columns_uniq[ii])+', '+str(table_columns_datatype[ii])+']'
                ii+=1
            #print('|||||||||||||:::::::::      '+table_description)
            if len(table_columns_pk) == 1 and str(table_columns_uniq[0]) == 'TRUE' and str(table_columns_datatype[0]) == 'COUNTER':
                
                if addBoolean:
                   table_description += ' )'
                   
                try:
                    cursor.execute('select * from '+x)
                    #print(cursor.description[0][0])
                    #column = cursor.description[0][0]
                    cursor.execute(
                    'ALTER TABLE '+table_name+' ADD CONSTRAINT PrimaryKey PRIMARY KEY('+str(table_columns_pk[0])+')')
                    conn.commit()
                    cursor.execute('ALTER TABLE '+table_name +
                               ' DROP CONSTRAINT PK_ID')
                    conn.commit()
               
                    #listTables.insert('end', x)
                    lbPK.append(x)
                except:
                    #listTablesPK.insert('end', table_description)
                    lPk.append(table_description)
            else:
                if addBoolean:
                    table_description += ' )'
                #listTables.insert('end', table_description)
                lbPK.append(table_description)
            
    except Exception as e:
        print(e)
        messagebox.showwarning(
            'Error', "Something wrong with DataBase!")
    lbPK.sort()
    lPk.sort()
    for x in lbPK:
        listTables.insert('end',x)
    for x in lPk:
        listTablesPK.insert('end',x)

def addColumnF(sent_addpkw,sent_tableFrame):
        try:
                cc = conn.cursor()
                ccc= conn.cursor()
                cursor = conn.cursor()
                auxrowi =0
                columns_table=[]
                columns_table_unique=[]
                columns_table_unique_columns=[]
                columns_table_datatype=[]
                length=0
                col= 1
                roww = 0
                bi=0
                bii=0
                cursor.execute('SELECT * INTO AUX FROM '+tablename+' WHERE 1 = 2;')
                conn.commit()
                cursor.execute('ALTER TABLE AUX ADD ID_'+tablename+' Integer;')
                conn.commit()
                cursor.execute('ALTER TABLE AUX ALTER COLUMN ID_'+tablename+' AUTOINCREMENT;')
                
                conn.commit()
                #Read rows and insert it manually
                cursor.execute('SELECT * FROM AUX')
                for x in cursor.description:
                    if auxrowi == 0:
                        auxrow =x[0]
                        auxrowi = 2
                    else:
                        if str(x[0]) != 'ID_'+tablename:
                            auxrow = auxrow+','+x[0]
                            #print(auxrow)
                        
                print(auxrow)
                cursor.execute('SELECT * FROM '+tablename)
                for row in cursor.fetchall():
                    #print (row)
                    cursor.execute("INSERT INTO AUX("+auxrow+") VALUES"+str(row))
                    conn.commit()
                #End
                cursor.execute('DROP TABLE '+tablename+';')
                conn.commit()
                cursor.execute('SELECT * INTO '+tablename+' FROM AUX;')
                cursor.execute('ALTER TABLE '+tablename+' ADD UNIQUE (ID_'+tablename+')')
                conn.commit()
                cursor.execute('DROP TABLE AUX;')
                conn.commit()
                
                col= 1
                roww = 0
                
                checkboxes.clear()
                checkvar.clear()
                checkboxTables.clear()
                for c in ccc.columns():
                    if str(c.table_name) == tablename: 
                        columns_table_datatype.append(str(c.type_name))
                        #print('Tipul de date: '+str(c.type_name))
                        if str(c.column_name) not in columns_table:
                            columns_table.append(str(c.column_name))
                        #    print('Nume coloana: '+str(c.column_name))
                for row in cc.statistics(table=tablename):
                    if str(row.column_name) != 'None':
                        if row.non_unique == 0 and str(row.column_name) not in columns_table_unique_columns:
                            columns_table_unique.append('TRUE')
                            columns_table_unique_columns.append(str(row.column_name))
                        #    print('Coloana + unique: '+str(row.column_name)+' True')
                        elif str(row.column_name) not in columns_table_unique_columns:
                            columns_table_unique.append('FALSE')
                            columns_table_unique_columns.append(str(row.column_name))
                        #    print('Coloana + unique: '+str(row.column_name)+' False')

                leng=len(columns_table_unique_columns)
                #print(columns_table)
                add=0
                for x in columns_table:
                    textCol=x+' ['
                    #print('Lungime : '+str(len(columns_table_unique))+' |||| '+str(len(columns_table_unique_columns)))
                    if x in columns_table_unique_columns:
                        textCol+='  UNIQUE: '+columns_table_unique[bi] + ', TYPE: '+columns_table_datatype[bii]+' ]'
                        bi+=1
                    if x not in columns_table_unique_columns:
                        textCol+=' UNIQUE: FALSE , TYPE: '+ columns_table_datatype[bii]+' ]'
                   
                    bii+=1
                    cv = tkinter.IntVar()
                    c = tkinter.Checkbutton(sent_tableFrame, text=textCol,variable =cv ,onvalue = 1, offvalue = 0)
                                        
                    c.grid(row = roww , column = col)
                    checkboxes.append(c)
                    checkvar.append(cv)
                    checkboxTables.append(x)
                    if col < 4:
                        col = col+1
                    else:
                        col = 1
                        roww = roww + 1
                    
                
        except Exception as e:
                cursor.execute('DROP TABLE AUX')
                conn.commit()
                # print(dblocation)
                print(e)
                if str(e) == "('HYS21', '[HYS21] [Microsoft][ODBC Microsoft Access Driver] Resultant table not allowed to have more than one AutoNumber field. (-1510) (SQLExecDirectW)')":
                        messagebox.showwarning('Error', "Can't add new column,you can't add more than one Autonumber field!")

                        cursor.execute('ALTER TABLE '+tablename+' DROP COLUMN ID_'+tablename)
                        conn.commit()
                        
                        print('Aux drop')
                else:
                        messagebox.showwarning('Error', "Can't add new column,something went wrong!")
               

def makePk():
        fields =''
        i = 0
        j = 0
        try:
            for x in checkvar:
                if x.get() == 1:
                    
                    if j ==0:
                        fields = fields + checkboxTables[i]
                        j = 1
                    else:
                        fields = fields+','+checkboxTables[i]
                    i= i + 1
                else:
                    i = i + 1

            msgBox = messagebox.askquestion ('Warning',"The fields "+fields+' will become Primary Key! Continue ?')
            if msgBox == 'yes':
                print('ALTER TABLE '+tablename+' ADD CONSTRAINT PrimaryKey PRIMARY KEY (' + fields + ');')
                cursor.execute('ALTER TABLE '+tablename+' ADD CONSTRAINT PrimaryKey PRIMARY KEY (' + fields + ');')
                conn.commit()
            refreshTables()
        except Exception as e:
            print(e)
            messagebox.showwarning('Error', "Can't add make PK!")

def autoPKC(tablename):
    add=[]
    ccname=''
    try:
        cur = conn.cursor()
        c = conn.cursor()
        cc = conn.cursor()
        for row in c.statistics(table=tablename):
                #print(row.non_unique)
                #add.append(row.non_unique)
                if row.non_unique == 0:
                    #print("OOOOOF")
                    for riot in cc.columns():
                        #print("OOOOOF")
                        if str(riot.type_name) == 'COUNTER' and str(riot.column_name) == row.column_name:
                            print('ALTER TABLE '+tablename+' ADD CONSTRAINT PrimaryKey Primary Key('+str(row.column_name)+')')
                            cur.execute('ALTER TABLE '+tablename+' ADD CONSTRAINT PrimaryKey Primary Key('+str(row.column_name)+')')
                            cur.commit()
                                
                            refreshTables()
                            break
               

        if 0 not in add:
            for riot in cc.columns():
                if riot.table_name == tablename and str(riot.type_name) == 'COUNTER':
                    ccname=str(riot.column_name)
                    cur.execute('ALTER TABLE '+tablename+' ADD UNIQUE ('+ccname+')')
                    cur.commit()
                    cur.execute('ALTER TABLE '+tablename+' ADD CONSTRAINT PrimaryKey Primary Key('+ccname+')')
                    cur.commit()
            

        refreshTables()
    except Exception as e:
            refreshTables()
            print(e)
            if str(e) == "('HY000', '[HY000] [Microsoft][ODBC Microsoft Access Driver] Primary key already exists. (-1402) (SQLExecDirectW)')":
                messagebox.showwarning('Error', "PrimaryKey allready exist in this table!")


        
   


def dropPKF(tablename):
    columns=[]
    columnsUnique=''
    try:
        c = conn.cursor()
        cc = conn.cursor()
        ccc = conn.cursor()
        
        for x in c.statistics(table=tablename):
            if str(x.index_name) == 'PrimaryKey':
                    columns.append(x.column_name)
        print(len(columns))
        if len(columns) == 0:
            cc.execute('ALTER TABLE '+tablename+' DROP CONSTRAINT PrimaryKey ')
            cc.commit()
            cc.execute('ALTER TABLE '+tablename+' ADD UNIQUE ('+columns[0]+')')
            cc.commit()
        else:
            for y in columns:
                columnsUnique +=y+','
            cUnique=columnsUnique[0:(len(columnsUnique))-1]
            print(cUnique)
            cc.execute('ALTER TABLE '+tablename+' DROP CONSTRAINT PrimaryKey ')
            cc.commit()
            cc.execute('ALTER TABLE '+tablename+' ADD CONSTRAINT uniq UNIQUE ('+str(cUnique)+')')
            cc.commit()

        refreshTables()
    except Exception as e:
            print(e)
            if "CHECK constraint 'PrimaryKey' does not exist." in str(e):
                messagebox.showwarning('Error', "This table doesn't have PrimaryKey!")
            else:
                messagebox.showwarning('Error', "Can't drop PrimaryKey!")
def addPKWindow():
        global tablename,checkboxes,checkvar,checkboxTables,addPKw
        
        # try:
        #     addPKw.destroy()
        # except Exception as e:
        #         # print(dblocation)
        #         print(e)
        
        try:

                cc = conn.cursor()
                ccc= conn.cursor()
                checkboxes = []
                checkvar = []
                checkboxTables = []
                textCol=''
                tablenamet = listTables.get(listTables.curselection())
                aux=tablenamet.split()
                tablename=aux[0]
                addPKw = tkinter.Toplevel()
                addPKw.title('Add Primary Key')
                columns_table=[]
                columns_table_unique=[]
                columns_table_unique_columns=[]
                columns_table_datatype=[]
                length=0
                col= 1
                roww = 0
                bi=0
                bii=0
                background_label = tkinter.Label(addPKw, image=ph)
                background_label.place(x=0, y=0, relwidth=1, relheight=1)


                
                tablesFrame = tkinter.LabelFrame(addPKw,text='Columns of the table '+tablename.upper())
                tablesFrame.grid(row=0,column=1,pady=5, padx=5)

                dropPK = tkinter.Button(addPKw, text='Drop Primary Key',width=20,command = lambda:dropPKF(tablename))
                dropPK.grid(row=0, column=0, pady=5, padx=5)

                autoPK = tkinter.Button(addPKw, text='Auto Primary Key',width=20,command = lambda:autoPKC(tablename))
                autoPK.grid(row=1, column=0, pady=5, padx=5)

                createPK = tkinter.Button(addPKw, text='Make Primary Key',width=20,command = lambda:makePk())
                createPK.grid(row=2, column=0, pady=5, padx=5)

                addColumn = tkinter.Button(addPKw, text='    Add Id Column   ',width=20,command= lambda:addColumnF(addPKw,tablesFrame))
                addColumn.grid(row=3, column=0, pady=5, padx=5)

                
                checkboxes.clear()
                checkvar.clear()
                checkboxTables.clear()
                for c in ccc.columns():
                    if str(c.table_name) == tablename: 
                        columns_table_datatype.append(str(c.type_name))
                        print('Tipul de date: '+str(c.type_name))
                        if str(c.column_name) not in columns_table:
                            columns_table.append(str(c.column_name))
                            print('Nume coloana: '+str(c.column_name))
                for row in cc.statistics(table=tablename):
                    if str(row.column_name) != 'None':
                        if row.non_unique == 0 and str(row.column_name) not in columns_table_unique_columns:
                            columns_table_unique.append('TRUE')
                            columns_table_unique_columns.append(str(row.column_name))
                            print('Coloana + unique: '+str(row.column_name)+' True')
                        elif str(row.column_name) not in columns_table_unique_columns:
                            columns_table_unique.append('FALSE')
                            columns_table_unique_columns.append(str(row.column_name))
                            print('Coloana + unique: '+str(row.column_name)+' False')

                leng=len(columns_table_unique_columns)
                print(columns_table)
                add=0
                for x in columns_table:
                    textCol=x+' ['
                    #print('Lungime : '+str(len(columns_table_unique))+' |||| '+str(len(columns_table_unique_columns)))
                    if x in columns_table_unique_columns:
                        textCol+='  UNIQUE: '+columns_table_unique[bi] + ', TYPE: '+columns_table_datatype[bii]+' ]'
                        bi+=1
                    if x not in columns_table_unique_columns:
                        textCol+=' UNIQUE: FALSE , TYPE: '+ columns_table_datatype[bii]+' ]'
                   
                    bii+=1
                    cv = tkinter.IntVar()
                    c = tkinter.Checkbutton(tablesFrame, text=textCol,variable =cv ,onvalue = 1, offvalue = 0)
                                        
                    c.grid(row = roww , column = col)
                    checkboxes.append(c)
                    checkvar.append(cv)
                    checkboxTables.append(x)
                    if col < 4:
                        col = col+1
                    else:
                        col = 1
                        roww = roww + 1                   
        
        except Exception as e:
                # print(dblocation)
                print(e)
                messagebox.showwarning(
                'Error', "Please select a table from 'Tables without PK'!")




        


if __name__ == "__main__":
    global ph
    root = tkinter.Tk()
    root.title(
        "Primary key Normalization by Olteanu Ionut Valentin [ OVIDIUS ]")
    
    base_folder = os.path.dirname(__file__)
    filename = os.path.join(base_folder, 'background.jpg')
    
    im = Image.open(filename)
    ph = ImageTk.PhotoImage(im)
    background_label = tkinter.Label(root , image=ph)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)

    openDBButton = tkinter.Button(text='Open DataBase', command=openDb)
    openDBButton.grid(row=0, column=1, pady=5)

# p = ttk.Progressbar(root, length=200, mode='determinate')
# p.grid(row=0,column=2,pady=5)

    addPK = tkinter.Button(text='Edit Primary Key', command=addPKWindow)
    addPK.grid(row=1, column=1, pady=5)

    refresh = tkinter.Button(text='Refresh', command=refreshTables)
    refresh.grid(row=2, column=1, pady=5)

    noPKLabel = tkinter.Label(text="Tables with bad PK")
    noPKLabel.grid(row=1, column=0, padx=20, pady=20)

    listTables = tkinter.Listbox(root, height=20, width=110)
    listTables.grid(row=2, column=0, padx=20, pady=20)

    withPKLabel = tkinter.Label(text="Tables with PK")
    withPKLabel.grid(row=1, column=2, padx=20, pady=20)

    listTablesPK = tkinter.Listbox(root, height=20, width=110)
    listTablesPK.grid(row=2, column=2, padx=20, pady=20)
    
    
   

    root.mainloop()
