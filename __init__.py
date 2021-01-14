import os
from pathlib import Path
import logging
import datetime

systemdrive = os.getenv('WINDIR')[0:3]
dbfolder = os.path.join(systemdrive,'Forms\DB')
#dbfolder = "D:\Company Projects\Form creator\DB"
State_forms = os.path.join(systemdrive,'Forms\State forms')
#State_forms = "D:\Company Projects\Form creator\State forms"
Statefolder = Path(State_forms)
logfolder = os.path.join(systemdrive,'Forms\logs')
#logfolder = "D:\Company Projects\Form creator\logs"


monthdict= {'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}
#inverse_monthdict = dict((v,k) for k,v in monthdict.items())

log_filename = datetime.datetime.now().strftime(os.path.join(logfolder,'logfile_%d_%m_%Y_%H_%M_%S.log'))
logging.basicConfig(filename=log_filename, level=logging.INFO)