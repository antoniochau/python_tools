import sys, os, shutil
import time
import pathlib
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
import warnings
from src.lib_email import send_email
warnings.simplefilter("ignore")

sys.path.insert(1, r'\\glpfileshare.macausjm-glp.com\dept\financeAcc\Department Share\FinancialSystem\PythonTools\Util'  )
from Common.config import get_CONFIG
   
#from src.lib_processing import load_data_files , gen_source_df , gen_html_report


####################
# Project Packages
####################
from src.lib_flash_util import flash_util, flash_config
from src.lib_email import send_email
from src.lib_dataload import checking, load_data_files
from src.lib_output_df import main_gen_output_df
from src.lib_excel import gen_excel_files
from src.lib_update_summary import update_summary

#from src.lib_output_df import RN_total_ggr_df, RN_rolling_df, RN_non_rolling_df, RN_EG_df , Hotel_df


def copy_files_to_share_folder ( CONFIG ) :
    
    config = CONFIG.copy()
    
    output_folder = CONFIG['GAMING_DAY_PATH'] .parent
    
    YYYYMMDD = CONFIG['GAMING_DAY']['YYYYMMDD']

    
    RN_output_file_name = f"SJM Daily Summary - {YYYYMMDD}.xlsx"
    RN_original_file = Path (  CONFIG['OUTPUT_PATH'] ) /  CONFIG['RN_OUTPUT_FILE']
    RN_destination_file = output_folder / RN_output_file_name
    
    VM_output_file_name = f"SJM Daily Summary - {YYYYMMDD}(VIPMASS).xlsx"
    VM_original_file = Path (  CONFIG['OUTPUT_PATH'] ) /  CONFIG['VM_OUTPUT_FILE']
    VM_destination_file = output_folder / VM_output_file_name    

    
    output_folder_in_log = str(output_folder).replace( str( CONFIG['SHARED_DATA_PATH'] ) , "" )
    rolling_file_in_log  = RN_destination_file.name 
    vip_file_in_log      = VM_destination_file.name
            
    v_msg1 = flash_util.format_paragraph( "Folder Path", 25, str( output_folder_in_log ), 100)
    v_msg2 = flash_util.format_paragraph( "Rolling/Non-Rolling", 25, str ( rolling_file_in_log )  , 100)
    v_msg3 = flash_util.format_paragraph( "VIP/Non-VIP", 25, str ( vip_file_in_log )  , 100)
    log_msg = f"""\n\nREPORT_OUTPUTS\n#######################\n{v_msg1}\n{v_msg2}\n{v_msg3}     """ 
    
    config = flash_util.append_log_msg ( config, log_msg )   
    
    try:
        shutil.copy(RN_original_file, RN_destination_file )  # For Python 3.8+.
    except:
        pass
    
    try:
        shutil.copy(VM_original_file, VM_destination_file )  # For Python 3.8+.
    except:
        pass

def main( RERUN = False ):

    
    today = datetime.today()
    today_string = today.strftime("%Y%m%d")
    
    #today_string = today.strftime("20260107")   # for debug
    
    print ( "... Start ....")
  ###############################
  # Construction CONFIG dictionary
  ################################
    CONFIG = {}
    CONFIG ['APP_PATH'] = Path ( __file__ ).parent 
    CONFIG ['REPORT_DATE'] = today_string
    CONFIG = get_CONFIG(CONFIG )    

    CONFIG = flash_config.amend_CONFIG ( CONFIG )
    
    print ( CONFIG['GAMING_DAY']  )
    print ( CONFIG['COUNT_DAY'] )


  #######################################################
  # Stop the program if outputfiles are already generated
  #######################################################
    YYYYMMDD = CONFIG['GAMING_DAY']['YYYYMMDD']
    output_folder = CONFIG['GAMING_DAY_PATH'] .parent
    RN_output_file_name = f"SJM Daily Summary - {YYYYMMDD}.xlsx"
    RN_destination_file = output_folder / RN_output_file_name 
    
    is_output_already_exist = False
    #CONFIG ['FORCE_EXECUTION'] = True # for debug
    if RN_destination_file.exists() :
        is_output_already_exist = True
        
    if is_output_already_exist :        
        if CONFIG.get ( "FORCE_EXECUTION", False ) == False :
            print ( "output files generated already ")
            sys.exit()
        else:
            print ( "output files exists. FORCE EXECUTION ")
            
  ###############################
  # Check Missing Files
  ################################
    CONFIG = checking.check_files_exist ( CONFIG ) 

    if CONFIG['MANDATORY_FILE_MISSING'] == True :
        CONFIG = flash_util.append_log_msg ( CONFIG, "One or more mandatory file is missing. Program Terminated" )  



  ###############################
  # Load Data Files into DuckDB
  ###############################



    load_data_files.load_data_files ( CONFIG )
    print ( "Finished load_data_files ")
    
    CONFIG = main_gen_output_df.gen_output_df ( CONFIG )
    print ( "Finished main_gen_output_df ")
        
    CONFIG = gen_excel_files.gen_RN_file ( CONFIG )
    print ( "Finished Generate Rolling Non Rolling file ")

    CONFIG = gen_excel_files.gen_VM_file ( CONFIG )
    print ( "Finished Generate VIP Mass File ")
    
    copy_files_to_share_folder ( CONFIG )   
    print ( "Finished Copy Output File to Sharefolder ")
    
    send_email.send_log ( CONFIG ) 
    print ( "Finished Email ")
    
    update_summary . update_summary_tables ( CONFIG )
    print ( "Finished update Summary  ")
    
    
if __name__ == '__main__':
    
    main ()
    
