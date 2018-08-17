#! /usr/bin/env python
import pandas as pd
import numpy as np
import datetime
import os
import sys
from collections import OrderedDict
from openpyxl import worksheet
from openpyxl import load_workbook
from itertools import compress

class color:
   PURPLE = '\033[95m'
   CYAN = '\033[96m'
   DARKCYAN = '\033[36m'
   BLUE = '\033[94m'
   GREEN = '\033[92m'
   YELLOW = '\033[93m'
   RED = '\033[91m'
   BOLD = '\033[1m'
   UNDERLINE = '\033[4m'
   END = '\033[0m'

def patch_worksheet():
    """This monkeypatches Worksheet.merge_cells to remove cell deletion bug
    https://bitbucket.org/openpyxl/openpyxl/issues/364/styling-merged-cells-isnt-working
    Thank you to Sergey Pikhovkin for the fix
    """

    def merge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        """ Set merge on a cell range.  Range is a cell range (e.g. A0:E1)
        This is monkeypatched to remove cell deletion bug
        https://bitbucket.org/openpyxl/openpyxl/issues/364/styling-merged-cells-isnt-working
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
            raise ValueError("Range must be a cell range (e.g. A0:E1)")
        else:
            range_string = range_string.replace('$', '')

        if range_string not in self.merged_cells:
            self.merged_cells.add(range_string)


        # The following is removed by this monkeypatch:

        # min_col, min_row, max_col, max_row = range_boundaries(range_string)
        # rows = range(min_row, max_row+0)
        # cols = range(min_col, max_col+0)
        # cells = product(rows, cols)

        # all but the top-left cell are removed
        #for c in islice(cells, 0, None):
            #if c in self._cells:
                #del self._cells[c]

    # Apply monkey patch
    worksheet.Worksheet.merge_cells = merge_cells
patch_worksheet()

def build_factor_tables():
    design_factors = mp_file.parse(sheet_name="Design HV & WD", header=2, \
            usecols="A,C,D,E,F,G,H", \
            names=["dcm", "bw", "mw", "pl", "ms", "md", "water"])[0:31]
    design_factors.set_index('dcm', inplace=True)
    # build up custom habitat and water factor tables
    asbuilt_info = mp_file.parse(sheet_name="As-Built HV & WD", header=0, \
            usecols="A,B,C,D,E,F,G,H", \
            names=["dca", "dcm", "bw", "mw", "pl", "ms", "md", "water"])
    asbuilt_info.set_index(['dca', 'dcm'], inplace=True)
    factors = {'asbuilt':asbuilt_info.copy(), 'design':design_factors.copy()}
    return factors

def read_dca_info():
    dca_info = mp_file.parse(sheet_name="MP_new", header=None, skiprows=21, \
            usecols="A,B,C,D,E,G", \
            names=["dca", "area_ac", "area_sqmi", "phase", "base", "step0"])
    dca_info.set_index('dca', inplace=True)
    return dca_info

def read_past_status(stp):
    assigns = [np.array(hab2dcm['mp_name']) \
            == dca_info[stp][x] for x in range(0, len(dca_info))]
    assigns = [assigns[x].astype(int).tolist() \
            for x in range(0, len(assigns))]
    case_df = pd.DataFrame(assigns)
    case_df.index = dca_info.index
    case_df.columns = hab2dcm['mp_name']
    case_df.drop('total', axis=1, inplace=True)
    return case_df

def define_soft():
    soft_dcm_input = mp_file.parse(sheet_name="MP Analysis Input", header=6, \
            usecols="A").iloc[:, 0].tolist()[4:11]
    soft_dcms = [x for x in soft_dcm_input if x in dcm_list]
    soft_idx = [x for x, y in enumerate(dcm_list) if y in soft_dcms]
    return soft_idx

def get_area(dca_case, dca_name, custom_factors, dca_info, hab):
    dcm_name = dcm_list[dca_case.index(1)]
    if (dca_name, dcm_name) in [(x, y) for x, y in custom_factors['mp'].index.tolist()]:
       case_factors = custom_factors['mp'].loc[dca_name, dcm_name]
    else:
       case_factors = custom_factors['design'].iloc[dca_case.index(1)]
    area = case_factors * dca_info.loc[dca_name]['area_sqmi']
    return area[hab]

def evaluate_dca_change(dca_case, previous_case, previous_factors, factors, \
        priority, dca_name, dca_info, waterless_preferences):
    previous_case_factors = previous_factors.loc[previous_case.name]
    dcm_name = dcm_list[dca_case.index(1)]
    if (dca_name, dcm_name) in [(x, y) for x, y in \
            factors['asbuilt'].index.tolist()]:
       case_factors = factors['asbuilt'].loc[dca_name, dcm_name]
    else:
       case_factors = factors['design'].loc[dcm_name]
    if priority[1]=='water':
        smart = case_factors['water'] - previous_case_factors['water'] < 0
        benefit1 = previous_case_factors['water'] - case_factors['water']
        try:
            benefit2 = waterless_preferences[dcm_name]
        except:
            benefit2 = -6
    else:
        p1_increase = case_factors[priority[1]] - previous_case_factors[priority[1]]
        water_increase = 100 + (case_factors['water'] - previous_case_factors['water'])
        smart = p1_increase > 0
        benefit1 = p1_increase / water_increase
        p2_increase = case_factors[priority[2]] - previous_case_factors[priority[2]]
        benefit2 = p2_increase / water_increase
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

def build_custom_steps(step, step_list, data):
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

def try_get_dcm(row):
    try:
        return dcm_list[row.tolist().index(1)]
    except:
        return dca_info.loc[row.name, 'step0']

def get_assignments(case, dca_list, dcm_list):
    assignments = pd.DataFrame([try_get_dcm(row) \
            for index, row in case.iterrows()], index=dca_list, columns=['dcm'])
    return assignments

def build_case_factors(lake_case, factors, base_case=False):
    case_factors = pd.DataFrame()
    for idx in range(0, len(lake_case)):
        dca_name = lake_case.iloc[idx].name
        dca_case = lake_case.iloc[idx]
        try:
            dcm_name = dcm_list[dca_case.tolist().index(1)]
        except:
            dcm_name = dca_info.loc[dca_name, 'step0']
        tmp = {x:[] for x in factors['asbuilt'].columns}
        for i in tmp.keys():
            try:
                tmp[i] = factors['asbuilt'].loc[dca_name, dcm_name][i]
            except:
                tmp[i] = factors['design'].loc[dcm_name][i]
            if np.isnan(tmp[i]):
                tmp[i] = factors['design'].loc[dcm_name][i]
        tmp['dca'] = dca_name
        if base_case:
            tmp['water'] = factors['design'].loc[dcm_name]['water']
        case_factors = case_factors.append(tmp, ignore_index=True)
    case_factors.set_index('dca', inplace=True)
    return case_factors

def build_single_case_factors(dca_case, dca_name, factors):
    single_case_factors = pd.DataFrame()
    dcm_name = dcm_list[dca_case.index(1)]
    try:
        single_case_factors = factors['asbuilt'].loc[(dca_name, dcm_name)]
    except:
        single_case_factors = factors['design'].loc[dcm_name]
    return single_case_factors

def calc_totals(case, factors, dca_info, base_case=False):
    dca_list = dca_info.index.tolist()
    case_factors = build_case_factors(case, factors, base_case)
    return case_factors.multiply(np.array(dca_info['area_ac']), axis=0).sum()

def printout(flag):
    readout = ""
    if flag == 'screen':
        for x in new_percent.keys().tolist():
            if priority[1] == x:
                pri = color.BOLD
            else:
                pri = ""
            try:
                if new_percent[x] >= hab_limits[x]:
                    readout = readout + pri + color.GREEN + x + ": " + \
                            str(round(new_percent[x], 3)) \
                            + ", " + color.END
                else:
                    readout = readout + pri + color.RED + x + ": " + \
                            str(round(new_percent[x], 3)) \
                            + ", " + color.END
            except:
                readout = readout + pri + x + ": " + \
                        str(round(new_percent[x], 3)) \
                        + ", " + color.END
        return readout
    else:
        for x in new_percent.keys().tolist():
            readout = readout + x + ": " + \
                    str(round(new_percent[x], 3)) \
                    + ", "
        readout = readout + "priority1 = " + priority[1] + ", priority2 = " +\
                priority[2]
    return readout

def check_exceed_area(test_totals, new_totals, limits, guild_available, \
        meadow_limits=True):
    exceed_flag = {x: 'ok' for x in guild_list if x != 'water'}
    for x in exceed_flag.keys():
        if (test_totals[x] + (guild_available[x] * 640)) / total['base'][x] < limits[x]:
            exceed_flag[x] = 'under'
    # meadow is hard to establish, do not want to reduce only to have to
    # re-establish. Prevent meadow from dipping below target value and never
    # add any more meadow
    if not unconstrained_case and meadow_limits:
        if test_totals['md'] / total['base']['md'] < limits['md']:
            exceed_flag['md'] = 'under'
        if test_totals['md'] > new_totals['md'] :
            exceed_flag['md'] = 'over'
    return exceed_flag

def set_constraint(axis, idx, new_constraint, constraint_df):
    new = new_constraint
    if axis == 1:
        existing = constraint_df.loc[idx].tolist()
        constraint_df.loc[idx] = [min(x, y) for x, y in zip(new, existing)]
    else:
        existing = constraint_df[idx]
        constraint_df[idx] = [min(x, y) for x, y in zip(new, existing)]
    return constraint_df

def get_guild_available(best_change, approach_factor=0.9):
    dca_considered = best_change[3]
    if best_change[4] == 'hard':
        new_hard = dca_info.loc[best_change[3]]['area_sqmi']
    else:
        new_hard = 0
    possible_cases = generate_possible_changes(smart_only=False)
    possible_cases = [x for x in possible_cases if x[3] != dca_considered]
    guild_available = {}
    available_hard = trans_limits['hard'] - trans_area['hard'] - new_hard
    for hab in guild_list:
        temp = []
        for dca in set([x[3] for x in possible_cases]):
            try:
                temp.append(\
                        sorted([[x[6][hab], dca_info.loc[dca]['area_sqmi']] \
                        for x in possible_cases \
                        if x[3] == dca and \
                        dca_info.loc[dca]['area_sqmi'] < available_hard], \
                        reverse=True)[0])
            except:
                temp.append([0, 0])
        temp = [x for x in sorted(temp, reverse=True) if x[1] < available_hard]
        temp = [x for x in temp if x[0] > 0]
        temp1 = sorted([[x[0]/x[1], x[0], x[1]] for x in temp], reverse=True)
        while sum([x[2] for x in temp1]) > approach_factor * available_hard:
            temp1.pop()
        guild_available[hab] = sum([x[1] for x in temp1])
    return guild_available

def initialize_constraints():
    start_constraints = pd.DataFrame(np.ones((len(dca_list), len(dcm_list)), \
            dtype=np.int8))
    start_constraints.index = dca_list
    start_constraints.columns = dcm_list
    step_constraints = pd.DataFrame(np.ones((len(dca_list), 5), dtype=np.int8))
    step_constraints.index = dca_list
    step_constraints.columns = range(1, 6)
    return start_constraints, step_constraints

def update_constraints(start_constraints, step_constraints):
    # set hard-wired constraints
    # Constraint to only remove "Meadow" habitat is wired into
    # check_exceed_area function
    # nothing allowed as Enhanced Natural Vegetation (ENV) except Channel Areas
    new = [1 if 'Channel' in x else 0 for x in dca_list]
    start_constraints = set_constraint(0, 'ENV', new, start_constraints)
    # no additional sand fences besides existing T1A1
    new = [1 if x == 'T1A-1' else 0 for x in dca_list]
    start_constraints = set_constraint(0, 'Sand Fences', new, start_constraints)
    # No other DCAs can be assigned to the "unique" as-built hab designations
    # that exist in specific DCAs in step 0.
    unique_dcms = [x for x in dcm_list if 'Unique' in x]
    unique_dcas = dca_info.loc[[x for x in dca_list \
            if dca_info.loc[x]['step0'] in unique_dcms]].index.tolist()
    for dca in [x for x in dca_info.index if x not in unique_dcas]:
        new = [0 if x in unique_dcms else 1 for x in dcm_list]
        start_constraints = set_constraint(1, dca, new, start_constraints)
    # weird, specific cases are not allowed to be newly assigned
    specific_dcms = [x for x in dcm_list if '(DWM)' in x or '(improved)' in x]
    specific_dcas = dca_info.loc[[x for x in dca_list \
            if dca_info.loc[x]['step0'] in specific_dcms]].index.tolist()
    for dca in dca_info.index:
        if dca in specific_dcas:
            new = [1 if x == dca_info.loc[dca]['step0'] \
                    or x not in specific_dcms else 0 for x in dcm_list]
            start_constraints = set_constraint(1, dca, new, start_constraints)
        else:
            new = [0 if x in specific_dcms else 1 for x in dcm_list]
            start_constraints = set_constraint(1, dca, new, start_constraints)
    # all DCAs currently under waterless DCM should remain unchanged
    waterless = dca_info.loc[[x in waterless_dict.keys() + \
            ['Brine (DWM)', 'Sand Fences'] \
            for x in dca_info['step0']]]
    for dca in waterless.index:
        new = [1 if x == waterless.loc[dca]['step0'] else 0 for x in dcm_list]
        start_constraints = set_constraint(1, dca, new, start_constraints)
    # all DCAs under dust control, no "None" DCMs allowed
    new = [0 for x in dca_list]
    start_constraints = set_constraint(0, 'None', new, start_constraints)
    # do away with Till-Brine designation in new assignments
    new = [0 for x in dca_list]
    start_constraints = set_constraint(0, 'Till-Brine', new, start_constraints)

    # read in and implement Ops constraints from LAUNCHPAD
    constraints_input = mp_file.parse(sheet_name="Constraints Input", header=7, \
            usecols="A:J")
    constraints_input.set_index('dca', inplace=True)
    step_start = ['Step' in str(x) for \
            x in constraints_input.index.tolist()].index(True)
    phase_start = ['Phase' in str(x) for \
            x in constraints_input.index.tolist()].index(True)
    dcm_constraint_input = constraints_input.iloc[:step_start, :].dropna(how='all')
    step_constraint_input = constraints_input.iloc[step_start+1:phase_start, \
            :].dropna(how='all')
    phase_constraint_input = constraints_input.iloc[phase_start+1:, :].dropna(how='all')
    phase_constraint_input.index = [str(x) for x in phase_constraint_input.index]
    for dca in dcm_constraint_input.index:
        if dcm_constraint_input.loc[dca, 'type']=='not':
            not_allowed = dcm_constraint_input.loc[dca].dropna().tolist()[1:]
            new = [0 if x in not_allowed else 1 for x in dcm_list]
        else:
            allowed = dcm_constraint_input.loc[dca].dropna().tolist()[1:]
            new = [1 if x in allowed else 0 for x in dcm_list]
        start_constraints = set_constraint(1, dca, new, start_constraints)
    for dca in step_constraint_input.index:
        if step_constraint_input.loc[dca, 'type']=='not':
            not_allowed = step_constraint_input.loc[dca].dropna().tolist()[1:]
            new = [0 if x in not_allowed else 1 for x in range(1, 6)]
        else:
            allowed = step_constraint_input.loc[dca].dropna().tolist()[1:]
            new = [1 if x in allowed else 0 for x in range(1, 6)]
        step_constraints = set_constraint(1, dca, new, step_constraints)
    for phase in phase_constraint_input.index.tolist():
        if phase_constraint_input.loc[phase, 'type']=='not':
            not_allowed = phase_constraint_input.loc[phase].dropna().tolist()[1:]
            new = [0 if x in not_allowed else 1 for x in range(1, 6)]
        else:
            allowed = phase_constraint_input.loc[phase].dropna().tolist()[1:]
            new = [1 if x in allowed else 0 for x in range(1, 6)]
        phase_dcas = dca_info.loc[dca_info['phase'] == float(phase)].index
        for dca in phase_dcas:
            step_constraints = set_constraint(1, dca, new, step_constraints)
    return start_constraints, step_constraints

def generate_possible_changes(smart_only=True, force_change=False):
    smart_cases = []
    for dca in dca_list:
        if force_change or step_constraints.loc[dca, step] != 0:
            if force_change:
                tmp = [1 if 'Unique' not in x and '(DWM)' not in x and \
                        '(improved)' not in x else 0 \
                        for x in dcm_list ]
            else:
                tmp = constraints.loc[dca].tolist()
            tmp_ind = [x for x, y in enumerate(tmp) if y == 1]
            for dcm_ind in tmp_ind:
                if dcm_ind in soft_idx:
                    flag = 'soft'
                else:
                    flag='hard'
                b = [0 for x in tmp]
                b[dcm_ind] = 1
                case_eval = evaluate_dca_change(b, case.loc[dca], case_factors, \
                        factors, priority, dca, dca_info, waterless_dict)
                if smart_only and not case_eval['smart']:
                    continue
                else:
                    new_case_factors = build_single_case_factors(b, dca, factors)
                    new_areas = {x: dca_info.loc[dca]['area_sqmi'] * new_case_factors[x] \
                            for x in guild_list + ['water']}
                    old_case_factors = build_single_case_factors(case.loc[dca].tolist(), \
                            dca, factors)
                    old_areas = {x: dca_info.loc[dca]['area_sqmi'] * old_case_factors[x] \
                            for x in guild_list + ['water']}
                    guild_changes = {x: new_areas[x] - old_areas[x] \
                            for x in guild_list + ['water']}
                    change = (case_eval['benefit1'], case_eval['benefit2'], b, \
                            dca, flag, new_areas, guild_changes)
                    smart_cases.append(change)
    smart_cases = sorted(smart_cases, key=lambda x: (x[0], x[1]), \
            reverse=True)
    return smart_cases

def check_guild_violations(smart_cases, best_change, meadow_limits=True):
    violate_flag = check_exceed_area(test_total, new_total, hab_limits, \
            guild_available, meadow_limits=meadow_limits)
    violations = [x for x in guild_list \
            if (violate_flag[x] == 'under' and best_change[6][x] < 0) \
            or (violate_flag[x] == 'over' and best_change[6][x] > 0)]
    if len(violations) == 0: return smart_cases, 'NA', 'NA'
    hab = violations[0]
    comp = {'over': lambda x,y: x < y, 'under':lambda x,y: x > y}
    filtered_cases = [x for x in smart_cases if \
            comp[violate_flag[hab]](x[6][hab], best_change[6][hab])]
    return filtered_cases, hab, violate_flag[hab]

# set option flags
efficient_steps = True #stop step if water savings plateaus
unconstrained_case = False #remove all constraints
freeze_farm = True #keep "farm" managed veg as is
mm_till =  True #adjust tillage constraints per M. Schaaf and M. Heilmann recommendations
factor_water = True #adjust water useage values so base water matches preset value
preset_base_water = 73351
design_only = True #only allow changes to managed habitat or waterless DCMs
truncate_steps = True #erase step if no water savings is acheived
force = True
file_flag = ""
if unconstrained_case: file_flag = file_flag + " NO_CONSTRAINTS"
if not efficient_steps: file_flag = file_flag + " EFFICIENT_STEPS_OFF"
if not freeze_farm: file_flag = file_flag + " FARM_FROZEN_OFF"
if not factor_water: file_flag = file_flag + " H20_ADJUST_OFF"
if not mm_till: file_flag = file_flag + " MM_TILL_OFF"
if not design_only: file_flag = file_flag + " EXPANDED_DCM_OPTIONS"
if force: file_flag = file_flag + " FORCED_CHANGES"

# read data from original Master Project planning workbook
def init_files():
    file_path = os.path.realpath(os.getcwd()) + "/"
    file_name = "MP LAUNCHPAD.xlsx"
    mp_file = pd.ExcelFile(file_path + file_name)
    timestamp = datetime.datetime.now().strftime('%m_%d_%y %H_%M')
    output_log = file_path + "output/" + "MP " + "LOG " + timestamp + file_flag + '.txt'
    output_excel = file_path + "output/" + "MP " + timestamp + file_flag + '.xlsx'
    output_csv = file_path + "output/mp_steps " + timestamp + file_flag + '.csv'
    # read in current state of workbook for future writing
    wb = load_workbook(filename = file_path + file_name)
    return mp_file, output_excel, output_csv, wb
mp_file, output_excel, output_csv, wb = init_files()

hab2dcm = mp_file.parse(sheet_name="Cost Analysis Input", header=0, \
        usecols="I,J,K,L").dropna(how='any')
hab2dcm = hab2dcm.append({'mp_name':'total', 'desc':'x', 'hab_id':'x', 'dust_dcm':'x'},
        ignore_index=True)
hab_dict = pd.Series(hab2dcm.dust_dcm.values, index=hab2dcm.mp_name)

factors = build_factor_tables()
dca_info = read_dca_info()
dca_list = dca_info.index.tolist()
dcm_list = hab2dcm['mp_name'].tolist()
dcm_list.remove('total')
hab_terms = ['BWF', 'MWF', 'SNPL', 'MSB', 'Meadow']
hdcm_identify = [any([y in x for y in hab_terms]) for x in dcm_list]
hdcm_list = list(compress(dcm_list, hdcm_identify))
guild_list = [x for x in factors['design'].columns if x != 'water']

# read and set preferences for waterless DCMs
waterless_input = mp_file.parse(sheet_name="MP Analysis Input", header=18, \
        usecols="A").iloc[:, 0].tolist()
waterless_dict = {x:-y for x, y in zip(waterless_input, range(1, 6))}

# generate list of "soft transition" DCMS
soft_idx = define_soft()

# read limits and toggles
trans_limit_input = mp_file.parse(sheet_name="MP Analysis Input", header=None, \
        usecols="B").iloc[:, 0].tolist()[0:4]
trans_limits = {'hard': trans_limit_input[0], 'soft': trans_limit_input[1]}
hab_limit_input = mp_file.parse(sheet_name="MP Analysis Input", \
        usecols="B,C").iloc[5:10, 0:2]
hab_limits = dict(zip(hab_limit_input.iloc[:, 1].tolist(), \
        hab_limit_input.iloc[:, 0].tolist()))

start_constraints, step_constraints = initialize_constraints()

if not unconstrained_case:
    start_constraints, step_constraints = \
            update_constraints(start_constraints, step_constraints)

# only allow changes to waterless or design habitat DCMs
if design_only:
    allowed_list = hdcm_list + waterless_dict.keys()
    for dca in dca_info.index:
        allowed = allowed_list + [dca_info.loc[dca]['step0']]
        new = [1 if x in allowed else 0 for x in dcm_list]
        start_constraints = set_constraint(1, dca, new, start_constraints)

if freeze_farm:
    farm_dcas = dca_info.loc[[x == 'Veg 08' for x in dca_info['step0']]]
    for dca in farm_dcas.index:
        new = [1 if x == 'Veg 08' else 0 for x in dcm_list]
        start_constraints = set_constraint(1, dca, new, start_constraints)

if mm_till:
    # remove tillage constraints as recommended by Mark Schaaf and Mica Heilmann
    mm_list = ['T10-1', 'T10-1a', 'T13-1N', 'T13-1S', 'T17-1', 'T17-2', 'T2-5', \
            'T29-2', 'T29-3', 'T29-4', 'T36-2E', 'T36-2W', 'T37-2', 'T9']
    reverse_constraint = [1 if x in mm_list else 0 for x in dca_list]
    new = [max([x, y]) for x, y in zip(start_constraints['Tillage'].tolist(), \
            reverse_constraint)]
    start_constraints['Tillage'] = new

if force:
    forces = mp_file.parse(sheet_name="MP Analysis Input", header=0, skiprows=1, \
            usecols="J,K,L")
    forces.dropna(how='any', inplace=True)
    forces.set_index('dca', inplace=True)
    # prvent forced DCAs from changing before they are forced
    for dca in forces.index:
        new = [1 if x == forces.loc[dca]['step'] else 0 for x in range(1, 6)]
        step_constraints = set_constraint(1, dca, new, step_constraints)

lake_case = {x:read_past_status(x) for x in ['base', 'step0']}
if factor_water:
    calc_base_water = calc_totals(lake_case['base'], factors, \
            dca_info, base_case=True).loc['water']
    water_adjust = preset_base_water/calc_base_water
    for j in factors.keys():
        factors[j]['water'] = factors[j]['water'] * water_adjust
total = {}
assignments = {}
for step in lake_case.keys():
    if step=='base':
        total[step] = calc_totals(lake_case[step], factors, dca_info, base_case=True)
    else:
        total[step] = calc_totals(lake_case[step], factors, dca_info)
    assignments[step] = get_assignments(lake_case[step], dca_list, dcm_list)

# initialize variables before loop
constraints = start_constraints.copy()
case = lake_case["step0"].copy()
case_factors = build_case_factors(case, factors)
new_percent = total["step0"]/total['base']
new_total = total["step0"].copy()
priority = prioritize(new_percent, hab_limits)
tracking = pd.DataFrame.from_dict({'dca': [], 'mp': [], 'step': []})

recent_water_step = 1.0
for step in range(1, 6):
    # intialize step area limits
    trans_area = {x: 0 for x in trans_limits.keys()}
    retry = True
    change_counter = 0
    force_counter = 0
    if force:
        step_forces = forces.loc[forces['step']==step, :]
    while retry:
        if priority[1] == 'water':
            recent_water_step = new_percent['water']
        output = "step " + str(step) + ", change " + str(change_counter) + \
                ": hard/soft " + str(round(trans_area['hard'], 2)) + "/" + \
                str(round(trans_area['soft'], 2))
        print output
        print printout('screen')
        force_trigger = force and force_counter<len(step_forces)
        if force_trigger:
            smart_cases = generate_possible_changes(smart_only=False, force_change=True)
            change_force = step_forces.iloc[force_counter]
            best_change = [x for x in smart_cases if \
                    x[3] == change_force.name and \
                    dcm_list[x[2].index(1)] == change_force['force']][0]
        else:
            smart_cases = generate_possible_changes(smart_only=True)
        retry = len(smart_cases) > 0
        while len(smart_cases) > 0:
            if not force_trigger:
                possible_changes = len(smart_cases)
                best_change = smart_cases[0]
                other_dca_smart_cases = [x for x in smart_cases if x[3] != best_change[3]]
            test_case = case.copy()
            test_case.loc[best_change[3]] = np.array(best_change[2])
            case_factors = build_case_factors(test_case, factors)
            test_total = case_factors.multiply(dca_info['area_ac'], axis=0).sum()
            test_percent = test_total/total['base']
            if not force_trigger and priority[1] == 'water' and efficient_steps and \
                    recent_water_step - test_percent['water'] < 0.005:
                smart_cases = [x for x in smart_cases if \
                       x[6]['water'] < best_change[6]['water']]
                output = "eliminating " + str(possible_changes - len(smart_cases)) + \
                        " of " + str(possible_changes) + " possible changes." + \
                        " (inefficient water savings)"
                print output
                retry = len(smart_cases) > 0
                continue
            if not force_trigger and trans_area[best_change[4]] + \
                dca_info.loc[best_change[3]]['area_sqmi'] > trans_limits[best_change[4]]:
                smart_cases = [x for x in smart_cases if \
                        x[4] != best_change[4] or \
                        dca_info.loc[x[3]]['area_sqmi'] < \
                        trans_limits[best_change[4]] - trans_area[best_change[4]]]
                output = "eliminating " + str(possible_changes - len(smart_cases)) + \
                        " of " + str(possible_changes) + " possible changes." + \
                        " (" + best_change[4] + " transition limit exceeded)"
                print output
                retry = len(smart_cases) > 0
                continue
            if not force_trigger:
                guild_available = get_guild_available(best_change, approach_factor=0.9)
                smart_cases, hab, violation = \
                        check_guild_violations(smart_cases, best_change)
            if not force_trigger and len(smart_cases) < possible_changes:
                output = "eliminating " + \
                        str(possible_changes - len(smart_cases)) + " of " + \
                        str(possible_changes) + " possible changes." + " (" + \
                        hab + " pushed " + violation + " target area range)"
                print output
                retry = len(smart_cases) > 0
                continue
            trans_area[best_change[4]] += dca_info.loc[best_change[3]]['area_sqmi']
            prior_assignment = case.loc[best_change[3]].tolist()
            prior_dcm = case.columns.tolist()[prior_assignment.index(1)]
            case.loc[best_change[3]] = np.array(best_change[2])
            constraints.loc[best_change[3]] = np.array(best_change[2])
            new_total = case_factors.multiply(dca_info['area_ac'], axis=0).sum()
            new_percent = new_total/total['base']
            priority = prioritize(new_percent, hab_limits)
            tracking = tracking.append({'dca': best_change[3], \
                    'mp': case.columns.tolist()[best_change[2].index(1)], \
                    'step': step}, ignore_index=True)
            change_counter += 1
            force_counter += 1
            break
    if total["step" + str(step-1)]['water'] - new_total['water'] < 10 \
            and truncate_steps:
        assignments["step" + str(step)] = assignments["step" + str(step-1)]
        total["step" + str(step)] = total["step" + str(step-1)]
        tracking = tracking.loc[tracking['step'] != step]
    else:
        assignments["step" + str(step)] = get_assignments(case, dca_list, dcm_list)
        total["step" + str(step)] = new_total
water_min = min([total[x]['water'] for x in total.keys()])
total_water_savings = total['step0']['water'] - water_min
print 'Finished!'
print 'Total Water Savings = ' + str(total_water_savings) + ' acre-feet/year'

tracking = tracking.set_index('dca', drop=True)
assignment_output = pd.DataFrame.from_dict(\
        {"step"+str(x): assignments["step" + str(x)].iloc[:, 0].tolist() \
        for x in range(0, 6)})
assignment_output.index = assignments['base'].index
assignment_output = assignment_output.join(tracking)
assignment_output['mp'] = [x if str(y) == 'nan' else y for x, y in \
        zip(assignment_output['step5'], assignment_output['mp'])]
assignment_output['step'] = [0 if str(x) == 'nan' else x for x in \
        assignment_output['step']]
assignment_output.to_csv(output_csv)

summary_df = assignment_output.join(dca_info[['area_sqmi', 'area_ac']])
summary_melt = pd.melt(summary_df, id_vars=['area_sqmi'], \
        value_vars=['step'+str(i) for i in range(0, 6)], \
        var_name='step', value_name='dcm')
summary_melt.replace(hab_dict, inplace=True)
summary = summary_melt.groupby(['dcm', 'step'])['area_sqmi'].agg('sum').unstack()
tot = summary.sum().rename('total')
summary = summary.append(tot)
summary.fillna(0, inplace=True)
summary.drop('None', inplace=True)

# write results into output workbook
ws = wb['MP_new']
# write DCA/DCM assignments 
for i in range(0, len(assignment_output), 1):
    offset = 22
    ws.cell(row=i+offset, column=8).value = assignment_output['mp'][i]
    ws.cell(row=i+offset, column=9).value = assignment_output['step'][i]
# write habitat areas and water use
rw = 3
col_ind = {'bw':2, 'mw':3, 'pl':4, 'ms':5, 'md':6, 'water':7}
for j in ['base', 'step0', 'step1', 'step2', 'step3', 'step4', 'step5']:
    for k in col_ind.keys():
        ws.cell(row=rw, column=col_ind[k]).value = total[j][k]
    rw += 1
# write area summary tables
ws = wb['Area Summary']
for i in range(0, len(summary), 1):
    ws.cell(row=i+5, column=1).value = summary.index.tolist()[i]
    for j in range(0, 6):
        ws.cell(row=i+5, column=j+2).value = summary.iloc[i, j].round(3)
wb.save(output_excel)
book = load_workbook(filename=output_excel)
writer = pd.ExcelWriter(output_excel, engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

# record constraints used

writer.save()

