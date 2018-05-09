#! /usr/bin/env python
import pandas as pd
import numpy as np
import datetime
import os
from openpyxl import worksheet
from openpyxl import load_workbook

def evaluate_dca_change(dca_case, previous_case, previous_factors, custom_factors, \
        priority, dca_idx, dca_info, waterless_preferences):
    previous_case_factors = previous_factors.loc[previous_case.name]
    dca_name = dca_info.iloc[dca_idx].name
    dcm_name = custom_factors['dcm'].index.tolist()[dca_case.index(1)]
    if (dca_name, dcm_name) in [(x, y) for x, y in custom_factors['mp'].index.tolist()]:
       case_factors = custom_factors['mp'].loc[dca_name, dcm_name]
    else:
       case_factors = custom_factors['dcm'].iloc[dca_case.index(1)]
    if priority[1]=='water':
        smart = case_factors['water'] - previous_case_factors['water'] < 0
        benefit1 = previous_case_factors['water'] - case_factors['water']
        try:
            benefit2 = waterless_preferences[dcm_name]
        except:
            benefit2 = -6
    else:
        smart = case_factors[priority[1]] - previous_case_factors[priority[1]] > 0
        benefit1 = case_factors[priority[1]] - previous_case_factors[priority[1]]
        benefit2 = previous_case_factors['water'] - case_factors['water']
    return {'smart': smart, 'benefit1':benefit1, 'benefit2':benefit2}

def prioritize(value_percents, hab_minimums):
    hab_deficits = {x: value_percents[x] - hab_minimums[x] \
            for x in value_percents.index.tolist() \
            if x != 'water'}
    if any([x < 0 for x in hab_deficits.values()]):
        sort_deficits = sorted(hab_deficits.values())
        one = hab_deficits.values().index(sort_deficits[0])
        two = hab_deficits.values().index(sort_deficits[1])
        return {1: hab_deficits.keys()[one], 2: hab_deficits.keys()[two]}
    else:
        return {1: 'water', 2: value_percents[0:5].idxmin()}

def backfill(row, backfill_data, columns_list):
    """
    row = Series with index of columns_list and name (X, Y) where Y = backfill_factors.index
    backfill_factors = DataFrame with all columns_list columns
    columns_list = list
    """
    for col in columns_list:
        if np.isnan(row[col]):
            row[col] = backfill_data.loc[row.name[1], col]
    return row

def build_custom_steps(step, step_list, data):
    """
    step = string
    step_list = list
    data = DataFrame with column 'step'
    """
    factors = data.loc[data['step']==step_list[0], :].copy()
    sub_steps = [step_list[x] for x in range(1, step_list.index(step)+1)]
    for sub in sub_steps:
        sub_df = data.loc[data['step']==sub, :].copy()
        for idx in factors.index:
            if idx in sub_df.index:
                factors.drop(idx, inplace=True)
        factors = factors.append(sub_df)
    factors.drop('step', axis=1, inplace=True)
    return factors

def get_assignments(case, dca_list, dcm_list):
    assignments = pd.DataFrame([dcm_list[row.tolist().index(1)] \
            for index, row in case.iterrows()], index=dca_list, columns=['dcm'])
    return assignments

def build_case_factors(lake_case, custom_factors, stp):
    factors = pd.DataFrame()
    for idx in range(0, len(lake_case)):
        dca_name = lake_case.iloc[idx].name
        dca_case = lake_case.iloc[idx]
        dcm_name = dca_case[dca_case==1].index[0]
        dca_idx = [x for x, y in \
                enumerate(custom_factors[stp].index.get_level_values('dca')) \
                if y==dca_name]
        dcm_idx = [x for x, y in \
                enumerate(custom_factors[stp].index.get_level_values('dcm')) \
                if y==dcm_name]
        custom_idx = [x for x in dca_idx if x in dcm_idx]
        if len(custom_idx)>0:
            tmp = custom_factors[stp].iloc[custom_idx].copy()
            tmp['dca'] = lake_case.index.tolist()[idx]
            factors = factors.append(tmp)
        else:
            tmp = custom_factors['dcm'].loc[dcm_name].copy()
            tmp['dca'] = lake_case.index.tolist()[idx]
            factors = factors.append(tmp)
    factors.set_index('dca', inplace=True)
    return factors

def calc_totals(case, custom_factors, step, dca_info):
    dca_list = dca_info.index.tolist()
    factors = build_case_factors(case, custom_factors, step)
    return factors.multiply(np.array(dca_info['area_ac']), axis=0).sum()

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
#patch_worksheet()

# read data from original Master Project planning workbook
file_path = os.path.realpath(os.getcwd()) + "/"
file_name = "MP Workbook LAUNCHPAD.xlsx"
mp_file = pd.ExcelFile(file_path + file_name)

# generate habitat and water duty factor tables
dcm_factors = mp_file.parse(sheet_name="Generic HV & WD", header=2, \
        usecols="A,C,D,E,F,G,H", \
        names=["dcm", "bw", "mw", "pl", "ms", "md", "water"])[0:31]
dcm_factors.set_index('dcm', inplace=True)
# build up custom habitat and water factor tables
custom_info = mp_file.parse(sheet_name="Custom HV & WD", header=0, \
        usecols="A,B,C,D,E,F,G,H,I", \
        names=["dca", "dcm", "step", "bw", "mw", "pl", "ms", "md", "water"])
custom_info.set_index(['dca', 'dcm'], inplace=True)
custom_filled = custom_info.apply(backfill, axis=1, \
        backfill_data=dcm_factors, \
        columns_list=['bw', 'mw', 'pl', 'ms', 'md', 'water'])
custom_steps = ['base', 'dwm', 'step0', 'mp']
factors = {x: build_custom_steps(x, custom_steps, custom_filled) \
        for x in custom_steps}
factors['dcm'] = dcm_factors.copy()

dca_info = mp_file.parse(sheet_name="MP_new", header=None, skiprows=21, \
        usecols="A,B,C,D,F", \
        names=["dca", "area_ac", "area_sqmi", "base", "step0"])
dca_info.set_index('dca', inplace=True)

# known cases
lake_case = {'base': [], 'step0': []}
for case in lake_case.keys():
    assignments = [np.array(factors['dcm'].index.get_level_values('dcm')) \
            == dca_info[case][x] for x in range(0, len(dca_info))]
    assignments = [assignments[x].astype(int).tolist() \
            for x in range(0, len(assignments))]
    case_df = pd.DataFrame(assignments)
    case_df.index = dca_info.index
    case_df.columns = factors['dcm'].index.tolist()
    lake_case[case] = case_df

# read DCA-DCM constraints file
start_constraints = mp_file.parse(sheet_name="Constraints", header=8, \
        usecols="A:AF")

# define "soft" transition DCMs
soft_dcm_input = mp_file.parse(sheet_name="MP_new", header=5, \
        usecols="K").iloc[:, 0].tolist()
soft_dcms = [x for x in soft_dcm_input if x !=0][0:7]
soft_idx = [x for x, y in enumerate(factors['dcm'].index.tolist()) if y in soft_dcms]

# read limits and toggles
script_input = mp_file.parse(sheet_name="MP_new", header=None, \
        usecols="L")[0].tolist()
hard_limit = script_input[0]
soft_limit = script_input[1]
dcm_limits = {}
dcm_limits['Brine'] = script_input[2]
dcm_limits['Sand Fences'] = script_input[3]
hab_limit_input = mp_file.parse(sheet_name="MP_new", header=None, \
        usecols="P")[0].tolist()
hab_buffer_over = 0
hab_buffer_under = 0.03
hab_limit_input = [x + hab_buffer_over for x in hab_limit_input]
hab_limits = dict(zip(['bw', 'mw', 'pl', 'ms', 'md'], hab_limit_input))

# read and set preferences for waterless DCMs
pref_input = mp_file.parse(sheet_name="MP_new", header=None, \
        usecols="M")[0].tolist()[5:11]
pref_dict = {x:-y for x, y in zip(pref_input, range(1, 6))}

dcm_list = factors['dcm'].index.tolist()
dca_list = lake_case['base'].index.get_level_values('dca').tolist()
total = {}
for case in lake_case.keys():
    total[case] = calc_totals(lake_case[case], factors, case, dca_info)
step_info = {}
step_info['base'] = {'totals': total['base'],
        'percent_base': total['base']/total['base'], \
        'hard_transition': 0, 'soft_transition': 0, \
        'assignments': get_assignments(lake_case['base'], dca_list, dcm_list)}

# initialize ariables before loop - DO NOT MAKE CHANGES HERE
new_constraints = start_constraints.copy()
new_case = lake_case['step0'].copy()
new_assignments = get_assignments(new_case, dca_list, dcm_list)
case_factors = build_case_factors(new_case, factors, 'mp')
new_percent = total['step0']/total['base']
new_total = total['step0'].copy()
tracking = pd.DataFrame.from_items([('step', []), ('dca', []), ('from', []), \
        ('to', []), ('bw', []), ('mw', []), ('pl', []), ('ms', []), ('md', []), \
        ('water', []), ('hard', []), ('soft', []), ('brine', []), ('sand_fences', [])])
tracking.index.name = 'change'
hard_transition = 0
soft_transition = 0
dcm_area_tracking = {}
sand_fence_dcas = [x for x, y in enumerate(new_case['Sand Fences']) if y ==1]
sand_fence_area = sum([dca_info['area_sqmi'][x] for x in sand_fence_dcas])
dcm_area_tracking['Sand Fences'] = sand_fence_area
# set priorities for initial change
priority = prioritize(new_percent, hab_limits)
step_info[0] = {'totals': new_total, 'percent_base': new_percent, \
        'hard_transition': hard_transition, \
        'soft_transition': soft_transition, \
        'assignments': new_assignments}
change_counter = 0
for step in range(1, 6):
    print 'step ' + str(step)
    hard_transition = 0
    soft_transition = 0
    dcm_area_tracking['Brine'] = 0
    while hard_transition < hard_limit or soft_transition < soft_limit:
        change_counter += 1
        print "hard = " + str(round(hard_transition, 2)) + ", soft = " +\
                str(round(soft_transition, 2)) + ", " + str(priority[1]) + " = " +\
                str(round(new_percent[priority[1]], 3)) + ", " + str(priority[2]) +\
                " = " + str(round(new_percent[priority[2]], 3))
        constraints = new_constraints.copy()
        eval_case = new_case.copy()

        allowed_cases = []
        for dca in range(0, len(constraints)):
            tmp = constraints.iloc[dca].tolist()
            tmp_ind = [x for x, y in enumerate(tmp) if y == 1]
            a = []
            for dcm in tmp_ind:
                b = [0 for x in tmp]
                b[dcm] = 1
                a.append(b)
            allowed_cases.append(a)
        n_allowed = [len(x) for x in allowed_cases]

        smart_cases = {'soft': [], 'hard': []}
        smartest = {'soft': [], 'hard': []}
        for dca in range(0, len(allowed_cases)):
            benefit1 = {'soft': [], 'hard': []}
            benefit2 = {'soft': [], 'hard': []}
            dca_assigns = {'soft': [], 'hard': []}
            for case in range(0, len(allowed_cases[dca])):
                case_eval = evaluate_dca_change(allowed_cases[dca][case], \
                    eval_case.iloc[dca], case_factors, factors, \
                    priority, dca, dca_info, pref_dict)
                if allowed_cases[dca][case].index(1) in soft_idx:
                    flag = 'soft'
                else:
                    flag='hard'
                if case_eval['smart']:
                    dca_assigns[flag].append(allowed_cases[dca][case])
                    benefit1[flag].append(case_eval['benefit1'])
                    benefit2[flag].append(case_eval['benefit2'])
            for flag in ['soft', 'hard']:
                no_change = (0, 0, eval_case.iloc[dca].tolist(), dca, flag)
                if not benefit1[flag]:
                    best = no_change
                else:
                    best = [(x, y, z, dca, flag) for x,y,z \
                            in sorted(zip(benefit1[flag], benefit2[flag], \
                            dca_assigns[flag]), key=lambda x: (x[0], x[1]), \
                            reverse=True)][0]
                smart_cases[flag].append(best)
        n_smart = [len(smart_cases[x]) for x in smart_cases]
        for flag in ['soft', 'hard']:
            smartest[flag] = sorted(smart_cases[flag], key=lambda x: (x[0], x[1]), \
                    reverse=True)

        soft_nn = 0
        hard_nn = 0
        best_change = sorted([smartest['soft'][soft_nn], \
                smartest['hard'][hard_nn]], key=lambda x: (x[0], x[1]), \
                reverse=True)[0]
        try:
            while True:
                best_change = sorted([smartest['soft'][soft_nn], \
                        smartest['hard'][hard_nn]], key=lambda x: (x[0], x[1]), \
                        reverse=True)[0]
                test_case = eval_case.copy()
                test_case.iloc[best_change[3]] = np.array(best_change[2])
                case_factors = build_case_factors(test_case, factors, 'mp')
                test_total = case_factors.multiply(dca_info['area_ac'], axis=0).sum()
                test_percent = test_total/total['base']
                test_deficits = {x: hab_limits[x] - test_percent[x] \
                        for x in test_percent.index.tolist() \
                        if x != 'water'}
                if any([x > hab_buffer_under for x in test_deficits.values()]):
                    print "Buffer Exceeded!"
                    if best_change[4] == 'soft':
                        soft_nn += 1
                    else:
                        hard_nn += 1
                    continue
                dcm_type = dcm_list[best_change[2].index(1)]
                if dcm_type in dcm_limits.keys():
                    if dca_info.iloc[best_change[3]]['area_sqmi'] + \
                            dcm_area_tracking[dcm_type] > dcm_limits[dcm_type]:
                        soft_nn += 1
                        print lim + " area exceeded!"
                        continue
                if best_change[4] == 'soft':
                    if soft_transition + \
                        dca_info.iloc[best_change[3]]['area_sqmi'] > soft_limit:
                        soft_nn += 1
                        continue
                    else:
                        for lim in dcm_limits.keys():
                            if dcm_list[best_change[2].index(1)] == lim:
                                dcm_area_tracking[lim] += \
                                        dca_info.iloc[best_change[3]]['area_sqmi']
                        soft_transition += dca_info.iloc[best_change[3]]['area_sqmi']
                        break
                else:
                    if hard_transition + \
                        dca_info.iloc[best_change[3]]['area_sqmi'] > hard_limit:
                        hard_nn += 1
                        continue
                    else:
                        hard_transition += dca_info.iloc[best_change[3]]['area_sqmi']
                        break

        except:
            break

        if best_change[0] <= 0:
            break

        new_case = eval_case.copy()
        new_case.iloc[best_change[3]] = np.array(best_change[2])
        new_constraints = constraints.copy()
        new_constraints.iloc[best_change[3]] = np.array(best_change[2])
        new_assignments = get_assignments(new_case, dca_list, dcm_list)
        case_factors = build_case_factors(new_case, factors, 'mp')
        new_total = case_factors.multiply(dca_info['area_ac'], axis=0).sum()
        new_percent = new_total/total['base']
        priority = prioritize(new_percent, hab_limits)
        change = pd.Series({'step': step, \
                'dca': dca_info.index.tolist()[best_change[3]], \
                'from': eval_case.columns.tolist()[\
                            eval_case.iloc[best_change[3]].tolist().index(1)],
                'to': eval_case.columns.tolist()[best_change[2].index(1)],
                'hard': hard_transition, \
                'soft': soft_transition,
                'brine': dcm_area_tracking['Brine'],
                'sand_fences': dcm_area_tracking['Sand Fences']})
        change = change.append(new_percent)
        change.name = change_counter
        tracking = tracking.append(change)
    step_info[step] = {'totals': new_total, 'percent_base': new_percent, \
            'hard': hard_transition, \
            'soft': soft_transition, \
            'assignments': new_assignments}
total_water_savings = total['base']['water'] - new_total['water']
print 'Finished!'
print 'Total Water Savings = ' + str(total_water_savings) + ' acre-feet/year'
assignment_output = new_assignments.copy()
assignment_output['step'] = 0
for i in range(1, 6):
    changes = [x for x in tracking.loc[tracking['step']==i, 'dca']]
    flag = [x in changes for x in assignment_output.index.tolist()]
    assignment_output.loc[flag, 'step'] = i
assignment_output['step0'] = (get_assignments(lake_case['step0'], dca_list, dcm_list))
assignment_output.columns = ['mp', 'step', 'step0']
assignment_output.index.name = 'dca'
dca_changes = zip(assignment_output['step0'], assignment_output['step'], \
        assignment_output['mp'])
for i in ['1', '2', '3', '4', '5']:
    assignment_output['step'+i] = [x[2] if x[1] <= int(i) else x[0] \
            for x in dca_changes]
output_csv = file_path + "output/mp_steps " + \
        datetime.datetime.now().strftime('%m_%d_%y %H_%M') + '.csv'
assignment_output.to_csv(output_csv)

summary_df = assignment_output.join(dca_info['area_sqmi'])
hab_ponds = [dcm_list[x] for x in [8, 9, 10, 11, 12, 14, 15, 16, 17, 18, 19, 29]]
summary_melt = pd.melt(summary_df, id_vars=['area_sqmi'], \
        value_vars=['step'+str(i) for i in range(0, 6)], \
        var_name='step', value_name='dcm')
for i in range(0, 6):
    empty = pd.DataFrame.from_dict({'area_sqmi': 0, 'step':'step'+str(i), \
            'dcm':dcm_list})
    summary_melt = summary_melt.append(empty)
summary = summary_melt.groupby(['step', 'dcm']).sum().unstack(level=[1])
summary['total'] = summary.sum(axis=1)

# write results into output workbook
wb = load_workbook(filename = file_path + file_name)
ws = wb['MP_new']
for i in range(0, len(assignment_output), 1):
    offset = 22
    ws.cell(row=i+offset, column=7).value = assignment_output['mp'][i]
    ws.cell(row=i+offset, column=8).value = assignment_output['step'][i]
rw = 3
for j in ['base', 0, 1, 2, 3, 4, 5]:
    for k in range(2, 8):
        ws.cell(row=rw, column=k).value = step_info[j]['totals'][k-2]
    rw += 1
output_excel = file_path + "output/" +file_name[:12] + \
        datetime.datetime.now().strftime('%m_%d_%y %H_%M') + '.xlsx'
wb.save(output_excel)
book = load_workbook(filename=output_excel)
writer = pd.ExcelWriter(output_excel, engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
tracking.to_excel(writer, sheet_name='Script Output - DCA Changes')
summary.to_excel(writer, sheet_name='Script Output - DCM Areas')
writer.save()
