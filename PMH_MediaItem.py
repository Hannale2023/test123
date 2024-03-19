import pandas as pd
import numpy as np
from datetime import date, datetime

import logging as logger
import logging.config
import pymysql
from sshtunnel import SSHTunnelForwarder
import win32con, win32api, os, openpyxl.cell._writer
ssh_host = '103.9.209.2'
ssh_password = 'RfHJUIyL@B7'
ssh_port = 329
localhost = '127.0.0.1'
localport = 3306
ssh_user = 'eplus'

# database variables
sql_username='root'
sql_password='Eplus@123?Q!'
sql_main_database='_c34c507d0e994817'

def executeSQL(query):
    conn = None
    """ access the database over the SSH tunnel and execute the query """
    with SSHTunnelForwarder(
        (ssh_host, ssh_port),
        ssh_username=ssh_user,
        ssh_password=ssh_password,
        remote_bind_address=(localhost, 3306)) as tunnel:
        try:
            conn = pymysql.connect(
                host='localhost', 
                user=sql_username,
                passwd=sql_password,
                db=sql_main_database,
                port=tunnel.local_bind_port,
                cursorclass=pymysql.cursors.DictCursor)
            
            mycursor = conn.cursor()
            mycursor.execute(query)
            result = mycursor.fetchall()
            data = pd.DataFrame(result)
            
        except ConnectionError as e:
            conn.close()
        finally:
            conn.close()
    return data
sql1 = """
SELECT 
DATE(creation) AS creation,
name,item_code, item_name, item_type,title,company, branch, branch_group,region, block , `position`,
category_class, from_date to_date, used_status, warning_status, workflow_state 
FROM `_c34c507d0e994817`.`tabMedia Item`
"""
Media_Item = executeSQL(sql1)
Media_Item.to_csv(r'C:\Users\DELL\OneDrive - eplusresearchvn\PMH_MediaItem.csv', index=False)