import tabula
import pandas as pd

class Changes:
    def __init__(self, path):
        dfs = tabula.read_pdf(path)
        df_aux = dfs[0]
        if len(dfs)>1:
            for df in dfs:
                df_aux = pd.concat(df_aux, df)
        self.df = df_aux
        self.format()
    
    def format(self):
        self.df.columns = ['gp_wk', 'wk', 'gp_cg', 'cg', 'mat', 'road_map']
        self.df['gp_wk'] = self.df['gp_wk'].apply(lambda x: x.split(' ')[1])
        self.df['mat'] = self.df['road_map'].apply(lambda x: x.split(' ')[0])
        self.df['road_map'] = self.df['road_map'].apply(lambda x: x.split(' ')[1])
        self.df['wk'] = self.df['wk'].str.lower() 
        self.df['cg'] = self.df['cg'].str.lower() 
    
    def getChanges(self):
        return [wk.lower() for wk in self.df['wk'].values]
    
    def changeByWk(self, wk):
        try:
            mat = self.df.loc[self.df['wk'] == wk, 'cg'].values[0]
        except:
            mat = 'WK nao encontrado'
        return mat
    
    def getMatByWk(self, wk):
        try:
            mat = self.df.loc[self.df['wk'] == wk, 'mat'].values[0]
        except:
            mat = 'WK nao encontrado'
        return mat

if __name__ == '__main__':
    c = Changes('./permuta.pdf')
    print(c.getChanges())
    print("alexia reis tavares" in c.getChanges())
    print(c.changeByWk("alexia reis tavares"))