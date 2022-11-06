import pandas as pd
import json
from module.CreateFrequency import CreaFrequency
from module.TransformExcel import TransformExcel
from module.MetaData import MetaData
from module.Changes import Changes

escala_path = 'test.xlsx'
metaData_path = './assets/data.xlsx'
escala_sheet = 'sb'

metaData = MetaData(metaData_path)
escala = TransformExcel(escala_path, sheet=escala_sheet)

areas = escala.uniqueAreas()
changes = Changes('./permuta.pdf')

frequency = {}

for area in areas+['supervisor']:
    frequency[area] = {}
    workers = []
    changesList = []
    for wk in escala.getFrequency(area):
        workers.append((wk, metaData.getMat(wk)))
        if wk in changes.getChanges():
            changesList.append((changes.changeByWk(wk), changes.getMatByWk(wk)))
        else:
            changesList.append(('-', '-'))
    frequency[area]['workers'] = workers
    frequency[area]['changes'] = changesList


CreaFrequency('2022-11-11').run(frequency, './assets/')

with open('frequency.json', 'w', encoding='utf-8') as file:
    json.dump(frequency,file, ensure_ascii=False)
