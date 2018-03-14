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
mp_name = "Master_proj_20180205.xlsm"

# read reference tables from NPY Calculation workbook
input_file = pd.ExcelFile(file_path + npv_name)
# what years are MP steps to be implemented?
step_years = input_file.parse(sheet_name="Script Input", header=0, \
        usecols="A,B").dropna(how='any')
# list of years for projection
year_range = range(int(step_years.year.min()), 2101, 1)
# DCM costs
dcm_costs = input_file.parse(sheet_name="Script Input", header=0, \
        usecols="D,E,F,G,H").dropna(how='any')
dcm_costs.dropna(inplace=True)
# mapping between MP habitats and DCMs
hab2dcm = input_file.parse(sheet_name="Script Input", header=0, \
        usecols="J,K,L,M").dropna(how='any')

# read data from John Dickey's workbook
mp_file = pd.ExcelFile(file_path + mp_name)
# MP implementation schedule
mp_new_names = ["dca", "acres", "dcm_base", "dcm_dwm", "dcm_step0", \
        "dcm_step5", "mp_step"]
mp_new = mp_file.parse(sheet_name="MP_new", header=1, \
        usecols="Q,R,U,V,W,X,Z", names=mp_new_names, \
        converters={'Step':int}).dropna(how='any')
mp_sched = mp_new.copy()
wd_hv_gen = mp_file.parse(sheet_name="WD_HV_gen", header=0, \
        usecols="C,I").dropna(how='any')
wd_src = wd_hv_gen.copy()
# expand mp_new data to show all steps in Master Project
for n in ['1', '2', '3', '4']:
    mp_new['dcm_step' + n] = ['X'] * len(mp_new)
    for i in mp_new.index:
        if mp_new.loc[i, 'mp_step'] > int(n):
            mp_new.loc[i, 'dcm_step' + n] = mp_new.loc[i, 'dcm_step0']
        else:
            mp_new.loc[i, 'dcm_step' + n] = mp_new.loc[i, 'dcm_step5']

# get dca areas
dcm_miles = np.array(mp_new.acres * 0.0015625)
dcm_acres = np.array(mp_new.acres)
# create clean mp_schedule
mp_new.drop(columns=['acres', 'mp_step'], inplace=True)
mp_new.columns = ['dca'] + step_years['step'].tolist()
mp_years = mp_new.copy()
mp_years.columns = ['dca'] + step_years['year'].tolist()
# expand mp_years table to include all years in range
for yr in year_range:
    if yr in mp_years.columns:
        mp_years[yr] = mp_years[yr]
    else:
        mp_years[yr] = mp_years[yr-1]
col_order = ['dca']
col_order.extend(year_range)
mp_years = mp_years[col_order]
mp_dcm = np.array(mp_years.copy())
mp_wd = np.zeros((mp_dcm.shape[0], mp_dcm.shape[1]-1, 10))
for name in hab2dcm.mp_name:
    mp_dcm[mp_dcm==name] = hab2dcm[hab2dcm.mp_name==name].dust_dcm
# build costs array 
# dim 0 = total capital cost ($ million) 
# dim 1 = o&m cost ($ million) 
# dim 2 = water demand (acre-ft/year)
mp_costs = np.zeros((mp_dcm.shape[0], mp_dcm.shape[1]-1, 11))
for i in range(0, len(mp_costs), 1):
    mp_costs[i, 0, 1] = dcm_costs.loc[dcm_costs.dust_dcm==mp_dcm[i, 1], \
            'om'].item() * dcm_miles[i]
    mp_costs[i, 0, 2] = wd_hv_gen.loc[wd_hv_gen.DCM==mp_years.iloc[i, 1], \
            'Water (f/y)'].item() * dcm_acres[i]
    lifespan = dcm_costs.loc[dcm_costs.dust_dcm==mp_dcm[i, 1], \
            'lifespan'].item()
    if lifespan == 0: lifespan = float('inf')
    age = 0
    for j in range(1, len(mp_costs[i, :]), 1):
        if mp_dcm[i, j+1] == mp_dcm[i, j]:
            mp_costs[i, j, 1] = dcm_costs.loc[dcm_costs.dust_dcm==\
                    mp_dcm[i, j+1], 'om'].item() * dcm_miles[i]
            age += 1
            if age > lifespan:
                mp_costs[i, j, 1] = 0
                mp_costs[i, j, 0] = dcm_costs.loc[dcm_costs.dust_dcm==\
                        mp_dcm[i, j+1], 'replacement'].item() * dcm_miles[i]
                age = 0
        else:
            mp_costs[i, j, 0] = dcm_costs.loc[dcm_costs.dust_dcm==\
                    mp_dcm[i, j+1], 'capital'].item() * dcm_miles[i]
            lifespan = dcm_costs.loc[dcm_costs.dust_dcm==mp_dcm[i, 1], \
                    'lifespan'].item()
            if lifespan == 0: lifespan = float('inf')
            age = 0
        mp_costs[i, j, 2] = wd_hv_gen.loc[wd_hv_gen.DCM==mp_years.iloc[i, j+1], \
            'Water (f/y)'].item() * dcm_acres[i]

npv_summary = pd.DataFrame({\
        'year':year_range, \
        'capital':np.sum(mp_costs[:, :, 0], axis=0), \
        'om':np.sum(mp_costs[:, :, 1], axis=0), \
        'water_demand':np.sum(mp_costs[:, :, 2], axis=0), \
        })

sheet_dict = {'base':'No Change', 'dwm':'DWM', 'step0':'Step0', 'step1':'Step1', \
        'step2':'Step2', 'step3':'Step3', 'step4':'Step4', \
        'step5':'Full Project'}
wb = load_workbook(filename = file_path + npv_name)
for step in step_years.step:
    yr = int(step_years[step_years.step==step].year.item())
    ind = npv_summary.year.tolist().index(yr)
    capital_output = npv_summary.capital.tolist()
    capital_output[ind:] = [0]*(len(capital_output) - ind)
    om_output = npv_summary.om.tolist()
    om_output[ind:] = [om_output[ind]]*(len(om_output) - ind)
    wd_output = npv_summary.water_demand.tolist()
    wd_output[ind:] = [wd_output[ind]]*(len(wd_output) - ind)
    ws = wb.get_sheet_by_name(sheet_dict[step])
    for i in range(0, len(npv_summary), 1):
        offset = 12
        ws.cell(row=i+offset, column=3).value = capital_output[i]
        ws.cell(row=i+offset, column=4).value = om_output[i]
        ws.cell(row=i+offset, column=19).value = wd_output[i]
ws = wb.get_sheet_by_name('NPV Summary')
ws.cell(row=18, column=2).value = 'Analysis run on ' + \
        datetime.datetime.now().strftime('%m-%d-%Y %H:%M')
ws.cell(row=19, column=2).value = "Data read from Master Project workbook '" + \
        mp_name + "'"
output_file = file_path + npv_name[:7] + \
        datetime.datetime.now().strftime('%m_%d_%y %H_%M') + '.xlsx'
wb.save(output_file)

book = load_workbook(filename=output_file)
writer = pd.ExcelWriter(output_file, engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
mp_sched.to_excel(writer, sheet_name='MP Schedule', index=False)
wd_src.to_excel(writer, sheet_name='MP Water', index=False)
writer.save()
