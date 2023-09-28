import os
import shutil

def main():

    os.remove('Counter.xlsx')
    os.remove('PRLU.xlsx')
    os.remove('Customer Care.xlsx') 
    os.remove('Book1.csv')
    shutil.rmtree('C:/Users/bugel049/AppData/Local/Temp/gen_py')

main()