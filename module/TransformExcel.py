import pandas as pd


class TransformExcel:
    def __init__(self, path, sheet='sb'):
        self.path = path
        try:
            self.df = pd.read_excel(path, sheet_name=sheet)
        except:
            raise(f'Erro: zip; file {self.path}')
        self.run()

    def getDate(self):
        self.df['date'] = self.df.iloc[1, 5].strftime("%m/%d/%Y")

    def getGP(self):
        self.df['gp'] = self.df.iloc[0, 7]

    def fillLastLocal(self):
        self.df.reset_index(inplace=True)
        self.df.drop(['index'], axis=1, inplace=True)

        for i in range(1, len(self.df)):
            if pd.isna(self.df.loc[i, 'pb']):
                self.df.loc[i, 'pb'] = self.df.loc[i-1, 'pb']

            if pd.isna(self.df.loc[i, 'eq']):
                self.df.loc[i, 'eq'] = self.df.loc[i-1, 'eq']

            if pd.isna(self.df.loc[i, 'road']):
                self.df.loc[i, 'road'] = self.df.loc[i-1, 'road']

    def formatColumns(self):
        cols = list(self.df.columns)
        date_position = cols.index('date')
        gp_position = cols.index('gp')
        self.df = self.df.loc[self.df.iloc[:, 9] == 1, :].iloc[:, [
            0, 1, 2, 3, 4, 5, 7, 9, date_position, gp_position]]
        self.df.columns = ['road', 'roadByworker', 'begin',
                           'end', 'pb', 'eq', 'worker', 'rowTrue', 'date', 'gp']
        self.df = self.df[['gp', 'date', 'road', 'roadByworker',
                           'begin', 'end', 'pb', 'eq', 'worker', 'rowTrue']]

    def transform_to_str(self):
        for c in self.df.columns:
            self.df[c] = self.df[c].apply(lambda x: str(x).strip().lower() if x != None else None)
            self.df[c] = self.df[c].apply(lambda x: ' '.join(x.split('-')) if x != None else None)
            #self.df[c] = self.df[c].apply(lambda x: '_'.join(x.split(' ')) if x != None else None)

    def mapAreasbyRoad(self):
        self.df['areas'] = self.df[self.df['road'].str.len() >= 3]['road'].apply(lambda x: x[0] if x[0] in ['1', '2', '3', '4', '5', '6'] else x[:3])
        self.df.loc[self.df['road'].str.len() < 3, 'areas'] = self.df['pb']
        self.df.loc[(self.df['pb'].str.contains('selve'))&(self.df['pb'].str.contains(r'dom|sup')), 'areas'] = 'SELVE DOM AVELAR'
        self.df.loc[(self.df['pb'].str.contains('selve'))&(self.df['pb'].str.contains('retiro')), 'areas'] = 'SELVE RETIRO'
        self.df.loc[(self.df['pb'].str.contains('selve'))&(self.df['pb'].str.contains('orlando')), 'areas'] = 'SELVE ORLANDO GOMES'
        self.df.loc[self.df['pb'].str.contains(r'sevop|condutor_do_coordenador|atendimento___n.o.a.|sefit|serat'), 'areas'] = 'INTERNO'

    @staticmethod
    def getInternWorkers(df):
        inHouse = df.loc[df['areas'] == 'INTERNO', :]
        return list(inHouse.index)

    @staticmethod
    def getNoa(df):
        noa = df.loc[df['road'].str.len() < 3, :]
        noa = noa[noa['pb'].str.contains('n.o.a')]
        return list(noa.index)

    def createSuper(self):
        self.df['super'] = 0
        condition = (
            ((self.df['pb'].str.contains('supervisor')) |
            (self.df['eq'].str.contains('supervisor')))&
            (~self.df['pb'].str.contains('condutor'))
        )
        self.df.loc[condition, 'super'] = 1

    @staticmethod
    def save(df, nameCSV):
        df.to_csv(f'../files/db_format/{nameCSV}.csv', index=False)

    def formatAreas(self):
        self.df['areas'] = self.df['areas'].str.replace('_', ' ')
        self.df['areas'] = self.df['areas'].str.capitalize()
        self.df['areas'] = self.df['areas'].apply(lambda x: f'Area {x}' if x in ['1', '2', '3', '4', '5'] else x)

    def run(self):
        self.getDate()
        self.getGP()
        self.formatColumns()
        self.fillLastLocal()
        self.transform_to_str()
        self.mapAreasbyRoad()
        self.createSuper()
        self.formatAreas()

    def getDF(self):
        return self.df
    
    def getFrequency(self, gp):
        if gp == 'supervisor':
            response =  self.df.loc[self.df['super']==1, :]
        else:
            response =  self.df.loc[(self.df['areas'] == gp)&(self.df['super'] == 0), :]
        return sorted(list(response['worker'].values))
    
    def uniqueAreas(self):
        return list(self.df['areas'].unique())


if __name__ == '__main__':

    import os

    excelFiles = list(os.listdir('../files/raw/'))
    for fds in ['sb']:
        df = TransformExcel(f'../files/raw/{excelFiles[0]}', sheet=fds).getDF()
        for xlsx in excelFiles[1:]:
            try:
                print(f'RUN: {xlsx}')
                df = pd.concat([df, TransformExcel(xlsx, sheet=fds).getDF()])
                print(f'OK: {xlsx}')
            except ValueError:
                print(f'ERROR: {xlsx}')

        TransformExcel.save(df, f'escala_full_{fds}')

        df = df.reset_index(drop=True)
        index_inHouse = set(TransformExcel.getInternWorkers(df))
        index_noa = set(TransformExcel.getNoa(df))
        index_op = set(df.index.to_list())-(index_inHouse.union(index_noa))

        op = df.loc[index_op, :].drop(['roadByworker'], axis=1)
        inH = df.loc[index_inHouse, :].drop(
            ['road', 'roadByworker', 'eq', 'rowTrue', 'areas'], axis=1)
        noa = df.loc[index_noa, :].drop(
            ['road', 'roadByworker', 'pb', 'rowTrue', 'areas'], axis=1)
        TransformExcel.save(op, f'escala_op_{fds}')
        TransformExcel.save(df.loc[index_inHouse, :], f'escala_inHouse_{fds}')
        TransformExcel.save(df.loc[index_noa, :], f'escala_noa_{fds}')
