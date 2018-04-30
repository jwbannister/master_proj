#!/usr/bin/env python

import pandas as pd
import numpy as np
from openpyxl import load_workbook
import csv
import datetime
from itertools import product
import types
#import openpyxl
from openpyxl import worksheet
from openpyxl.utils import range_boundaries


def patch_worksheet():
    """This monkeypatches Worksheet.merge_cells to remove cell deletion bug
    https://bitbucket.org/openpyxl/openpyxl/issues/365/styling-merged-cells-isnt-working
    Thank you to Sergey Pikhovkin for the fix
    """

    def merge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        """ Set merge on a cell range.  Range is a cell range (e.g. A1:E1)
        This is monkeypatched to remove cell deletion bug
        https://bitbucket.org/openpyxl/openpyxl/issues/365/styling-merged-cells-isnt-working
        """
        if not range_string and not all((start_row, start_column, end_row, end_column)):
            msg = "You have to provide a value either for 'coordinate' or for\
            'start_row', 'start_column', 'end_row' *and* 'end_column'"
            raise ValueError(msg)
        elif not range_string:
            range_string = '%s%s:%s%s' % (get_column_letter(start_column),
                                          start_row,
                                          get_column_letter(end_column),
                                          end_row)
        elif ":" not in range_string:
            if COORD_RE.match(range_string):
                return  # Single cell, do nothing
            raise ValueError("Range must be a cell range (e.g. A1:E1)")
        else:
            range_string = range_string.replace('$', '')

        if range_string not in self._merged_cells:
            self._merged_cells.append(range_string)


        # The following is removed by this monkeypatch:

        # min_col, min_row, max_col, max_row = range_boundaries(range_string)
        # rows = range(min_row, max_row+1)
        # cols = range(min_col, max_col+1)
        # cells = product(rows, cols)

        # all but the top-left cell are removed
        #for c in islice(cells, 1, None):
            #if c in self._cells:
                #del self._cells[c]

    # Apply monkey patch
    worksheet.Worksheet.merge_cells = merge_cells
patch_worksheet()

# path for excel files
file_path = "/home/john/airsci/owens/Master Project & Cost Benefit/"
npv_name = "MP NPV TEMPLATE.xlsx"
mp_name = "MP Workbook JWB 04-27-2018.xlsx"

# read reference tables from NPY Calculation workbook
input_file = pd.ExcelFile(file_path + npv_name)
# what years are MP steps to be implemented?
step_years = input_file.parse(sheet_name="Script Input", header=0, \
        usecols="A,B", converters={'year':int}).dropna(how='any')
step_dict = pd.Series(step_years.year.values, index=step_years.step)
# list of years for projection
year_range = range(int(step_years.year.min()), 2101, 1)
# DCM costs
dcm_costs = input_file.parse(sheet_name="Script Input", header=0, \
        usecols="D,E,F,G,H").dropna(how='any')
dcm_costs.dropna(inplace=True)
# mapping between MP habitats and DCMs
hab2dcm = input_file.parse(sheet_name="Script Input", header=0, \
        usecols="J,K,L,M").dropna(how='any')
hab_dict = pd.Series(hab2dcm.dust_dcm.values, index=hab2dcm.mp_name)

# read data from John Dickey's workbook
mp_file = pd.ExcelFile(file_path + mp_name)
# MP implementation schedule
mp_new_names = ["dca", "acres", "dcm_base", "dcm_dwm", "dcm_step0", \
        "dcm_step5", "mp_step"]
mp_new = mp_file.parse(sheet_name="MP_new", header=24, \
        usecols="A,B,D,E,F,G,H", names=mp_new_names, \
        converters={'Step':int}).dropna(how='any')
# pull water demand by DCA (in acre-ft/year)
mp_new_wd = mp_file.parse(sheet_name="MP_new", header=24, \
        usecols="A,H,W,AK,AY,BM", converters={'Step':int}, \
        names=[mp_new_names[i] for i in [0, 6, 2, 3, 4, 5]]).dropna(how='any')

# get dca areas
dca_areas = mp_new[['dca', 'acres']].copy().set_index('dca')
dca_areas['miles'] = dca_areas['acres'] * 0.0015625

mp_steps = mp_new.copy()
mp_steps_wd = mp_new_wd.copy()
# expand data to show all steps in Master Project
for obj in [mp_steps, mp_steps_wd]:
    obj.set_index('dca', inplace=True)
    for n in ['1', '2', '3', '4']:
        obj['dcm_step' + n] = ['X'] * len(obj)
        for i in obj.index:
            if obj.loc[i, 'mp_step'] > int(n):
                obj.loc[i, 'dcm_step' + n] = obj.loc[i, 'dcm_step0']
            else:
                obj.loc[i, 'dcm_step' + n] = obj.loc[i, 'dcm_step5']
    obj.drop(columns=[s for s in obj.columns if 'dcm' not in s], inplace=True)
    obj.columns=[s[4:] for s in obj.columns]

mp_years = mp_steps.copy()
mp_years_wd = mp_steps_wd.copy()
for obj2 in [mp_years, mp_years_wd]:
    obj2.columns = [step_dict[a] for a in obj2.columns]
    # expand table to include all years in range
    for yr in year_range:
        if yr in obj2.columns:
            obj2[yr] = obj2[yr]
        else:
            obj2[yr] = obj2[yr-1]

mp_years_dcm = mp_years[year_range].copy()
mp_years_dcm.replace(hab_dict, inplace=True)
# build costs array 
# dim 0 = total capital cost ($ million) 
# dim 1 = o&m cost ($ million) 
mp_costs = np.zeros((mp_years_dcm.shape[0], mp_years_dcm.shape[1], 2))
for i in range(0, len(mp_years_dcm.index), 1):
    mp_costs[i, 0, 1] = dcm_costs.loc[dcm_costs.dust_dcm==mp_years_dcm.iloc[i, 1], \
            'om'].item() * dca_areas['miles'][i]
    lifespan = dcm_costs.loc[dcm_costs.dust_dcm==mp_years_dcm.iloc[i, 1], \
            'lifespan'].item()
    if lifespan == 0: lifespan = float('inf')
    age = 0
    for j in range(1, len(year_range), 1):
        if mp_years_dcm.iloc[i, j-1] == mp_years_dcm.iloc[i, j]:
            mp_costs[i, j, 1] = dcm_costs.loc[dcm_costs.dust_dcm==\
                    mp_years_dcm.iloc[i, j], 'om'].item() * dca_areas['miles'][i]
            age += 1
            if age > lifespan:
                mp_costs[i, j, 1] = 0
                mp_costs[i, j, 0] = dcm_costs.loc[dcm_costs.dust_dcm==\
                        mp_years_dcm.iloc[i, j], 'replacement'].item() * \
                dca_areas['miles'][i]
                age = 0
        else:
            mp_costs[i, j, 0] = dcm_costs.loc[dcm_costs.dust_dcm==\
                    mp_years_dcm.iloc[i, j], 'capital'].item() * dca_areas['miles'][i]
            lifespan = dcm_costs.loc[dcm_costs.dust_dcm==mp_years_dcm.iloc[i, 1], \
                    'lifespan'].item()
            if lifespan == 0: lifespan = float('inf')
            age = 0

cost_summary = pd.DataFrame({\
        'year':year_range, \
        'capital':np.sum(mp_costs[:, :, 0], axis=0), \
        'om':np.sum(mp_costs[:, :, 1], axis=0), \
        'water_demand':mp_years_wd.sum(axis=0)
        })

sheet_dict = {'base':'No Change', 'dwm':'DWM', 'step0':'Step0', 'step1':'Step1', \
        'step2':'Step2', 'step3':'Step3', 'step4':'Step4', \
        'step5':'Full Project'}
wb = load_workbook(filename = file_path + npv_name)
for step in step_years.step:
    yr = int(step_years[step_years.step==step].year.item())
    ind = cost_summary.year.tolist().index(yr)
    capital_output = cost_summary.capital.tolist()
    capital_output[ind:] = [0]*(len(capital_output) - ind)
    om_output = cost_summary.om.tolist()
    om_output[ind:] = [om_output[ind]]*(len(om_output) - ind)
    wd_output = cost_summary.water_demand.tolist()
    wd_output[ind:] = [wd_output[ind]]*(len(wd_output) - ind)
    ws = wb.get_sheet_by_name(sheet_dict[step])
    for i in range(0, len(cost_summary), 1):
        offset = 12
        ws.cell(row=i+offset, column=3).value = capital_output[i]
        ws.cell(row=i+offset, column=4).value = om_output[i]
        ws.cell(row=i+offset, column=19).value = wd_output[i]

    mp_steps_dcm = pd.concat([mp_steps.replace(hab_dict), \
            dca_areas['acres']], axis=1)
    mp_steps_hab = pd.concat([mp_steps, dca_areas['acres']], axis=1)
    # write water demand summary tables
    ws = wb.get_sheet_by_name('Water Use Summary')
    wd_summary_dcm = pd.DataFrame({'wd':mp_steps_wd[step], \
            'dcm':mp_steps_dcm[step]}).groupby('dcm')['wd'].sum()
    wd_summary_dcm = wd_summary_dcm.reindex(dcm_costs.dust_dcm.tolist())
    wd_summary_dcm.fillna(0, inplace=True)
    wd_summary_dcm = \
            wd_summary_dcm.append(pd.Series({'total':wd_summary_dcm.sum()}))
    for i in range(0, len(wd_summary_dcm), 1):
        col = step_years.index[step_years.step==step].item()
        ws.cell(row=i+5, column=col+2).value = int(wd_summary_dcm[i].round())
    wd_summary_hab = pd.DataFrame({'wd':mp_steps_wd[step], \
            'hab':mp_steps[step]}).groupby('hab')['wd'].sum()
    wd_summary_hab = wd_summary_hab.reindex(hab2dcm.mp_name.tolist())
    wd_summary_hab.fillna(0, inplace=True)
    wd_summary_hab = \
            wd_summary_hab.append(pd.Series({'total':wd_summary_hab.sum()}))
    for i in range(0, len(wd_summary_hab), 1):
        col = step_years.index[step_years.step==step].item()
        ws.cell(row=i+5, column=col+12).value = int(wd_summary_hab[i].round())

    # write area summary tables
    ws = wb.get_sheet_by_name('Area Summary')
    area_summary_dcm = pd.DataFrame({'acres':mp_steps_dcm['acres'], \
            'dcm':mp_steps_dcm[step]}).groupby('dcm')['acres'].sum()
    area_summary_dcm = area_summary_dcm.reindex(dcm_costs.dust_dcm.tolist())
    area_summary_dcm.fillna(0, inplace=True)
    area_summary_dcm = \
            area_summary_dcm.append(pd.Series({'total':area_summary_dcm.sum()}))
    for i in range(0, len(area_summary_dcm), 1):
        col = step_years.index[step_years.step==step].item()
        ws.cell(row=i+5, column=col+2).value = int(area_summary_dcm[i].round())
    area_summary_hab = pd.DataFrame({'acres':mp_steps_dcm['acres'], \
            'hab':mp_steps_hab[step]}).groupby('hab')['acres'].sum()
    area_summary_hab = area_summary_hab.reindex(hab2dcm.mp_name.tolist())
    area_summary_hab.fillna(0, inplace=True)
    area_summary_hab = \
            area_summary_hab.append(pd.Series({'total':area_summary_hab.sum()}))
    for i in range(0, len(area_summary_hab), 1):
        col = step_years.index[step_years.step==step].item()
        ws.cell(row=i+5, column=col+12).value = int(area_summary_hab[i].round())

ws = wb.get_sheet_by_name('NPV Summary')
ws.cell(row=18, column=2).value = "Data read from Master Project workbook '" + \
        mp_name + "'"
ws.cell(row=19, column=2).value = 'NPV Analysis run on ' + \
        datetime.datetime.now().strftime('%m-%d-%Y %H:%M')
output_file = file_path + npv_name[:7] + \
        datetime.datetime.now().strftime('%m_%d_%y %H_%M') + '.xlsx'
wb.save(output_file)

book = load_workbook(filename=output_file)
writer = pd.ExcelWriter(output_file, engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
mp_new.to_excel(writer, sheet_name='MP Schedule', index=False)
mp_new_wd.to_excel(writer, sheet_name='MP Water', index=False)
writer.save()
