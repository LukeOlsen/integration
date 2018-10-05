import os
import logging

basedir = os.path.abspath(os.path.dirname(__file__))

LOGGING_LOCATION = 'sapb1adaptor.log'
LOGGING_LEVEL = logging.INFO
LOGGING_FORMAT = '%(asctime)s %(levelname)-8s %(message)s'

DIAPI = 'SAPbobsCOM90'
SERVER = 'Servidor-2'
LICENSE_SERVER = 'Servidor-2:30000'
LANGUAGE = 'ln_Spanish'
DBSERVERTYPE = 'dst_MSSQL2012'
DBUSERNAME = 'sa'
DBPASSWORD = 'Cesehsa2010'
COMPANYDB = 'Pruebas_Cesehsa'
B1USERNAME = 'manager'
B1PASSWORD = 'Pame079'
USE_TRUSTED = False
