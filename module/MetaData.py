import pandas as pd

class MetaData:
    def __init__(self, path):
        self.data = pd.read_excel(path) 
    
    def getSups(self):
        return self.data.loc[self.data['FUNCAO'].str.lower() == 'supervisor', 'NOME'].values
    
    def isSup(self, worker):
        list_sup = self.getSups()
        return worker in list_sup
    
    def getMat(self, worker):
        worker = worker.replace('_', ' ')
        worker = worker.lower()
        try:
            return self.data.loc[self.data['NOME'].str.lower().str.strip() == worker, 'MAT'].values[0]
        except:
            return -9
    
if __name__ == '__main__':
    metaData_path = './assets/data.xlsx'

    metaData = MetaData(metaData_path)
    print(metaData.getMat('adilson_góes_silva'))
