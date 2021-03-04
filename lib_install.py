import os
#import sys
#import json
#import requests
#import random,string
#from requests_toolbelt.multipart.encoder import MultipartEncoder
#import re
#import time


import pandas as pd
from pathlib import Path


libs = {'requests', 'random', 'requests_toolbelt', 'pandas', 'pathlib', 'openpyxl'}


for lib in libs:
	cmd = 'pip install {}'.format(lib)
	os.system(cmd)
	print('\n')