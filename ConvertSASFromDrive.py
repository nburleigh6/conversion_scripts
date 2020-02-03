# Author: Nick Burleigh, nburleigh@gmail.com
# This is used to push a SAS data file from a local (or mounted) drive to a MySQL database.
# It can take a list of files to push, so it can work in batches.

from sas7bdat import SAS7BDAT
from sqlalchemy import create_engine

host_name = 'dg-solar-db.ciedqx9kizpw.us-east-1.rds.amazonaws.com:3306'
db_name = 'dg_solar_db'
username = 'admin'
password = 'dgsolaradmin'
connection_string = 'mysql://' + username + ':' + password + '@' + host_name + '/' + db_name + '?charset=utf8mb4'
myengine = create_engine(connection_string)

dir = 'L:\\Research\\Solar Research\\US_DG_SAS_Data\\2020Q1\\'

filenames = ['de', 'eia_final_nem_q12020', 'eia_final_nemstorage_q12020', 'ma', 'tx', 'wa']

wholedir = ''
for file in filenames:
    wholedir = dir + file + '.sas7bdat'
    with SAS7BDAT(wholedir) as reader:
        print('Now pushing: ' + file + '.sas7bdat')
        tbl = 'SAS_' + file.upper() + '_2020Q1'
        df = reader.to_data_frame()
    #print(df.head(1))
    df.to_sql(con=myengine, name=tbl, if_exists='replace', index=False, chunksize=50000)
print('Okay, check the DB.')
