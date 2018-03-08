#!/usr/bin/env python
"""load_data.py: reads necessary data from Master Project excel files and \
        loads into Postgres tables for cost-benefit calculations"""

import pandas as pd
import numpy as np

# path for excel files
file_path = "/home/john/airsci/owens/Master Project & Cost Benefit/"

# read reference tables from NPY Calculation workbook
npv_file = pd.ExcelFile(file_path + "NPV Calculations JWB.xlsx")
# what years are MP steps to be implemented?
step_years = npv_file.parse(sheet_name="mp_schedule", header=0)
# DCM costs
dust_dcms = npv_file.parse(sheet_name="dust_dcms", header=0)
dust_dcms.dropna(inplace=True)
# mapping between MP habitats and DCMs
dcm_xmap = npv_file.parse(sheet_name="dcm_xmap", header=0, \
        usecols="A,B,C,D")
# calculation factors (escalation rate, etc.)
factors = npv_file.parse(sheet_name="factors", header=None, usecols=[0, 1], \
        index_col=0, convert_float=False).T
# list of years for projection
year_range = range(step_years.year.min(), factors['end_year']+1, 1)
npv_file.close()

# read data from John Dickey's workbook
mp_file = pd.ExcelFile(file_path + "Master_proj_20180205.xlsm")
# MP implementation schedule
mp_new_names = ["dca", "acres", "dcm_base", "dcm_dwm", "dcm_step0", \
        "dcm_step5", "mp_step"]
mp_new = mp_file.parse(sheet_name="MP_new", header=1, \
        usecols="Q,R,U,V,W,X,Z", names=mp_new_names, \
        converters={'Step':int})
wd_hv_gen = mp_file.parse(sheet_name="WD_HV_gen", header=0, \
        usecols="C,I")
mp_new.dropna(inplace=True)
wd_hv_gen.dropna(inplace=True)
# % area breakdown for habitats
#dcm_surface = mp_file.parse(sheet_name="DCM_surf", header=2, \
#        usecols="A,B,C,D,E,F,G,H,I")
#dcm_surface.dropna(inplace=True)
#dcm_surface['veg_total'] = dcm_surface.iloc[:, 4:-2].astype(float).sum(axis=1)
#dcm_surface = dcm_surface.iloc[:, [0, 1, 2, 3, 8, 9]]
#dcm_surface.columns = ['desc', 'habitat_dcm', 'saturated', 'ponded', 'dry', \
#        'veg']
mp_file.close()
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
mp_new = mp_new[['dca', 'dcm_base', 'dcm_dwm', 'dcm_step0', \
        'dcm_step1', 'dcm_step2', 'dcm_step3', 'dcm_step4', 'dcm_step5']]
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
for name in dcm_xmap.mp_name:
    mp_dcm[mp_dcm==name] = dcm_xmap[dcm_xmap.mp_name==name].dust_dcm
# build costs array 
# dim 0 = total capital cost ($ million) 
# dim 1 = o&m cost ($ million) 
# dim 2 = water demand (acre-ft/year)
mp_costs = np.zeros((mp_dcm.shape[0], mp_dcm.shape[1]-1, 11))
for i in range(0, len(mp_costs), 1):
    mp_costs[i, 0, 1] = dust_dcms.loc[dust_dcms.dust_dcm==mp_dcm[i, 1], \
            'om'].item() * dcm_miles[i]
    mp_costs[i, 0, 10] = wd_hv_gen.loc[wd_hv_gen.DCM==mp_years.iloc[i, 1], \
            'Water (f/y)'].item() * dcm_acres[i]
    lifespan = dust_dcms.loc[dust_dcms.dust_dcm==mp_dcm[i, 1], \
            'lifespan'].item()
    if lifespan == 0: lifespan = float('inf')
    age = 0
    for j in range(1, len(mp_costs[i, :]), 1):
        if mp_dcm[i, j+1] == mp_dcm[i, j]:
            mp_costs[i, j, 1] = dust_dcms.loc[dust_dcms.dust_dcm==\
                    mp_dcm[i, j+1], 'om'].item() * dcm_miles[i]
            age += 1
            if age > lifespan:
                mp_costs[i, j, 1] = 0
                mp_costs[i, j, 0] = dust_dcms.loc[dust_dcms.dust_dcm==\
                        mp_dcm[i, j+1], 'replacement'].item() * dcm_miles[i]
                age = 0
        else:
            mp_costs[i, j, 0] = dust_dcms.loc[dust_dcms.dust_dcm==\
                    mp_dcm[i, j+1], 'capital'].item() * dcm_miles[i]
            lifespan = dust_dcms.loc[dust_dcms.dust_dcm==mp_dcm[i, 1], \
                    'lifespan'].item()
            if lifespan == 0: lifespan = float('inf')
            age = 0
        mp_costs[i, j, 2] = wd_hv_gen.loc[wd_hv_gen.DCM==mp_years.iloc[i, j+1], \
            'Water (f/y)'].item() * dcm_acres[i]

time = np.array(year_range) - factors['projection_year'].item()
mwd_costs = np.array([factors['mwd_start_price'].item() * \
        (1 + factors['mwd_rate_increase'].item())**t \
        for t in range(0, len(year_range), 1)])
npv_summary = pd.DataFrame({\
        'year':year_range, \
        'capital':np.sum(mp_costs[:, :, 0], axis=0), \
        'o&m':np.sum(mp_costs[:, :, 1], axis=0), \
        'water_demand':np.sum(mp_costs[:, :, 10], axis=0), \
        })
npv_summary['capital_esc'] = npv_summary['capital'] *\
        ((1 + factors['cap_escalation'].item())**time)
npv_summary['o&m_esc'] = npv_summary['o&m'] *\
        ((1 + factors['om_escalation'].item())**time)
npv_summary['capital_cash'] = npv_summary['capital_esc'] *\
        (1 - factors['finance_percent'].item())
npv_summary['capital_finance'] = npv_summary['capital_esc'] *\
        factors['finance_percent'].item()
npv_summary['finance_payment'] = npv_summary['capital_finance'] * \
        factors['finance_rate'].item()/\
        (1 - ((1 + factors['finance_rate'].item())**\
            -(factors['finance_term'].item())))
npv_summary['finance_payment_cum'] = np.cumsum(npv_summary['finance_payment'])
npv_summary['pumped_water'] = \
        np.minimum(np.sum(npv_summary['water_demand'], axis=0), \
        np.repeat(factors['max_gw_usage'].item(), len(mwd_costs)))
npv_summary['pumping_cost'] = (npv_summary['pumped_water'] * \
        factors['pump_cost'].item())/1000000
npv_summary['pump_cost_esc'] = npv_summary['pumping_cost'] * \
        (1 + ((npv_summary['year'] - factors['start_year'].item())) * \
        factors['pump_escalation'].item())
npv_summary['purchased_water'] = npv_summary['water_demand'] - \
        npv_summary['pumped_water']
npv_summary['water_purchase_cost'] = (npv_summary['purchased_water'] * \
        mwd_costs)/1000000
npv_summary['cost'] = npv_summary[['o&m_esc', 'capital_cash', \
        'finance_payment_cum', 'pump_cost_esc']].sum(axis=1)
npv_summary['avoided_water_purchase'] = npv_summary['pumped_water'] + \
        (npv_summary['water_demand'][0] - npv_summary['water_demand'])
npv_summary['avoided_cost'] = (npv_summary['avoided_water_purchase'] * \
        mwd_costs)/1000000
npv_summary['net_cost'] = npv_summary['cost'] - npv_summary['avoided_cost']
npv_summary = npv_summary[['year', 'capital', 'o&m', 'capital_esc', \
         'o&m_esc', 'capital_cash', 'capital_finance', 'finance_payment', \
        'finance_payment_cum', 'water_demand', \
        'pumped_water', 'pumping_cost', 'pump_cost_esc', 'purchased_water', \
        'water_purchase_cost', 'cost', 'avoided_water_purchase', \
        'avoided_cost', 'net_cost']]










mp_transitions = mp_years.copy()
mp_transitions[year_range[0]] = False
for col in year_range[1:]:

breakdown = mp_years[2015]


