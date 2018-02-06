import pandas as pd

DATA_PATH = r'C:\Users\penut\Documents\Futures Contract\Reference\txf_1min_2017.csv'

def __main__():
    td = pd.read_table(DATA_PATH, sep=',')
    print(td.head(5))

__main__()
