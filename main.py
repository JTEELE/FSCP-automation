#Functionality code contained in '_functions.py'
import os
import re
from pathlib import Path
salaries_benefits_geo = 42
current_month = '2021-12-31'
import pandas as pd
from openpyxl import Workbook, load_workbook
from _functions import *

income_statement()