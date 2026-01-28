# -*- coding: utf-8 -*-
import os
import configparser

config = configparser.ConfigParser()

config.read('config.ini')

test_case_paths = ['testcases1.xlsx', # e.g. RfG Ranks: 1..44
                   'testcases2.xlsx', # e.g. RfG Ranks: 45..88
                   'testcases3.xlsx', # e.g. RfG Ranks: 89..132
                   'testcases4.xlsx', # e.g. RfG Ranks: 133..176
                   'testcases5.xlsx', # e.g. Custom Ranks: 3001..3044
                   'testcases6.xlsx'] # e.g. Custom Ranks: 3045..3087

for test_case_path in test_case_paths:
    print(f'\nUsing test case sheet: {test_case_path}:\n')
    config['General']['Casesheet path'] = test_case_path
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

    os.system('py execute_pscad.py')