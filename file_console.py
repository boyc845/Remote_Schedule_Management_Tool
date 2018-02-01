import os
from pathlib import Path
from collections import OrderedDict

def file_whether_exist(filename: str) -> bool:
    '''check the file whether exist'''
    return Path(filename).exists()

def save_data(filename: str, list_data: list):
    '''save data to the file line by line'''
    outfile = open(filename, 'w')
    for line in list_data:
        outfile.write(str(line))
        outfile.write('\n')
    outfile.close()

def read_data(filename:str) -> list:
    '''read data from file and apend to the list'''
    list_data = []
    infile = open(filename, 'r')
    for line in infile:
        list_data.append(line.strip())
    return list_data

def read_data_as_dict(filename: str) -> OrderedDict:
    '''read txt file and apend to a Dict'''
    dict_data = OrderedDict()
    infile = open(filename, 'r')
    for line in infile:
        if line.strip() != '':
            dict_data[line.split()[0]] = line.split()[1]
    return dict_data

def create_folder_current_directory(foldername: str) -> 'Path':
    '''create a folder in current directory if the folder is not exists'''
    path = Path.cwd()/ foldername
    if not file_whether_exist(path):
        path.mkdir()
    return path
