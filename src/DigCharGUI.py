# -*- coding: utf-8 -*-
"""
Created on Mon Apr  1 10:04:18 2019.

@author: joslaton
"""
import PySimpleGUI as sg
from pandas import read_csv, DataFrame, concat, ExcelWriter
from openpyxl import load_workbook
from copy import deepcopy
from subprocess import run
from win32com.client import DispatchEx
from shutil import copyfile
from pathlib import Path
from sys import exit

sg.theme('TanBlue')

def val_getter(table, col, condition):
    """

    val_getter is used to get the mean value of a test.

    Parameters
    ----------
    table : Pandas Dataframe
        table is the whole df.
    col : str
        col is the name of the column where the result is found.
    condition : str
        The name of the test in cond_test.

    Returns
    -------
    val : float
        The mean value of results of a given test in a given df.

    """
    val = table[col].loc[table[' COND_TEST'] == condition].mean()
    return val


def header_finder(pth):
    """

    header_finder is used to automatically identify the header for a given CSV.

    Parameters
    ----------
    pth : str
        The path to the given file.

    Returns
    -------
    header_line : int
        The line on which the header falls.

    """
    i = 0
    header_line = 0
    with open(pth, 'r') as f:
        for line in f.readlines():
            if (line[0] == '#'):
                header_line = i
            i += 1
    header_line += 1
    return header_line


def dub_or_norm(df):
    """
    dub_or_norm returns the better of double and single speed results.

    Parameters
    ----------
    df : pandas dataframe
        df is the current df.

    Returns
    -------
    use_mhz : float
        The mean result in MHz.
    use_ns : float
        The mean result in nanoseconds.

    """
    if 'min_write_sclk_period_double' in df[' COND_TEST'].values:
        dub_mhz = val_getter(df, ' MEAS_MIN_WRITE_SCLK_PERIOD_MHZ',
                             'min_write_sclk_period_double')
        norm_mhz = val_getter(df, ' MEAS_MIN_WRITE_SCLK_PERIOD_MHZ',
                              'min_write_sclk_period_normal')
        dub_ns = val_getter(df, ' MEAS_MIN_WRITE_SCLK_PERIOD_NS',
                            'min_write_sclk_period_double')
        norm_ns = val_getter(df, ' MEAS_MIN_WRITE_SCLK_PERIOD_NS',
                             'min_write_sclk_period_normal')
        if ((dub_mhz > 1000 and dub_ns > 1000) or (dub_mhz < norm_mhz)):
            use_mhz = norm_mhz
            use_ns = norm_ns
        else:
            use_mhz = dub_mhz
            use_ns = dub_ns
    else:
        use_mhz = val_getter(df, ' MEAS_MIN_WRITE_SCLK_PERIOD_MHZ',
                             'min_write_sclk_period_normal')
        use_ns = val_getter(df, ' MEAS_MIN_WRITE_SCLK_PERIOD_NS',
                            'min_write_sclk_period_normal')
    return (use_mhz, use_ns)


def make_jsl(func_path, infile, outfile):
    lines = []
    lines += ['//!\n']
    lines += [f'include("{func_path}");\n']
    lines += [f'x=interior_funct("{infile}", "{outfile}");\n']
    outpath = str(Path(func_path).parent.resolve().as_posix())
    outpath += '/TMP.jsl'
    with open(outpath, 'w+') as f:
        f.writelines(lines)
    return outpath

def get_settings(settings_file, template_file):
    try:
        with open(settings_file, 'r') as f:
            settings = f.readlines()
            settings = {el.split(';')[0]: el.split(';')[1].rstrip() for el in settings}
    except Exception as e:
        layout_settings = [
            [sg.Text('JMP Location', size=(15, 1)), sg.InputText(key='-JMP-'), sg.FileBrowse() ],
            [sg.OK(size=(15, 1)), sg.Cancel(size=(15,1))]
        ]
        window = sg.Window('Initial Settings', layout_settings)
        event, values = window.read()
        if event in (None, 'Cancel'):
            exit()
        jmp_file = Path(values['-JMP-'])
        if jmp_file.is_file() and values['-JMP-'][-7:]=='jmp.exe':
            lines = []
            lines += [f'JMP Location;{values["-JMP-"]}\n']
            lines += ['Derate;7.0\n']
            lines += [f'Template;{template_file}\n']
            with open(settings_file, 'w+') as f:
                f.writelines(lines)
        else:
            sg.popup('Invalid Input. Please select the jmp.exe file.')
        return get_settings(settings_file, template_file)
    return settings['JMP Location'], settings['Template'], float(settings['Derate'])

def set_settings(settings_file, jmp, template, derate):
    lines = []
    lines += [f'JMP Location;{jmp}\n']
    lines += [f'Derate;{derate}\n']
    lines += [f'Template;{template}\n']
    with open(settings_file, 'w+') as f:
        f.writelines(lines)

def get_input(derate, template):
    layout = [
        [sg.Text('Input File (CSV)', size=(15, 1)), sg.InputText(key='-IN-', size=(70,1)), sg.FileBrowse(file_types=(("CSV Files","*.csv"),))],
        [sg.Text('Derating Factor', size=(15, 1)), sg.InputText(default_text=derate, size=(70,1), key='-DERATE-')],
        [sg.Text('Template File', size=(15, 1)), sg.InputText(default_text=template, size=(70,1), key='-TMP-')],
        [sg.Stretch(), sg.Ok(size=(10, 1)), sg.Cancel(size=(10, 1)), sg.CB('Generate PowerPoint Report', key='-PPT-', default=True), sg.Stretch()]
    ]
    window = sg.Window('Select File', layout)
    event, values = window.read()
    window.close()
    if event in (None,'Cancel'):
        exit()
    
    return values['-IN-'], values['-TMP-'], float(values['-DERATE-']), values['-PPT-']


pwd = str(Path(__file__).parent.resolve().as_posix())
src_template_path = pwd + "/Templates/TEMPLATE_DPA_1p8V_Copy.xlsx"
src_settings_path = pwd + "/Settings/Settings.txt"
src_jsl_func_path =  pwd + '/JMP/Digital Char Process-Jon.jsl'

jmp_path, src_template_path, derate = get_settings(src_settings_path, src_template_path)

src_data_path, src_template_path, derate, ppt = get_input(derate, src_template_path)

output_path = src_data_path[:-4] + '_Parametrics_Analysis.xlsx'
output_path_ppt = src_data_path[:-4] + '_Report.pptx'

set_settings(src_settings_path, jmp_path, src_template_path, derate)

# Here we look at our template and create a copy
copyfile(src_template_path, output_path)
header_rslt = header_finder(src_data_path)
raw_zorro_df = read_csv(src_data_path,
                           header=header_rslt).dropna(axis=1, how='all')


header_lst = ['DUT',                                    # 0
              'Temperature',                            # 1
              'VIO (V)',                                # 2
              'WriteSCLK_MaxFreq (MHz)',                # 3
              'WriteSCLK_MinPeriod (ns)',               # 4
              'ReadSCLK_MaxFreq (MHz)',                 # 5
              'ReadSCLK_MaxFreq (Derated) (MHz)',       # 6
              'ReadSCLK_MinPeriod (ns)',                # 7
              'ReadSCLK_MinPeriod (Derated) (ns)',      # 8
              'Setup_Time (ns)',                        # 9
              'Hold_Time (ns)',                         # 10
              'SCLK min high time (ns)',                # 11
              'SCLK min high time (Derated) (ns)',      # 12
              'SCLK min low time (ns)',                 # 13
              'Dstab time (ns)',                        # 14
              'Dstab time (Derated) (ns)',              # 15
              'SDATA VOH 1mA (V)',                      # 16
              'SDATA VOL 1mA (V)',                      # 17
              'SDATA VOH 2mA (V)',                      # 18
              'SDATA VOL 2mA (V)',                      # 19
              'SDATA IIH (uA)',                         # 20
              'SDATA IIL (uA)',                         # 21
              'SCLK IIH (uA)',                          # 22
              'SCLK IIL (uA)',                          # 23
              'VIH (SDATA) (V)',                        # 24
              'VIH (SCLK) (V)',                         # 25
              'VIL (SDATA) (V)',                        # 26
              'VIL (SCLK) (V)',                         # 27
              'VIO current (uA)']                       # 28
uniq_dut_lst = raw_zorro_df['COND_SERIAL_NUMBER'].unique().tolist()
uniq_temperature_lst = raw_zorro_df[' COND_TEMP_C'].unique().tolist()
uniq_vio_lst = raw_zorro_df[' COND_VCTRL_HI_V'].unique().tolist()
# tst_lst = df[' COND_TEST'].unique().tolist()

# %%
row_cnt = len(uniq_dut_lst)*len(uniq_temperature_lst)*len(uniq_vio_lst)
jmp_format_df = DataFrame(index=range(0, row_cnt), columns=header_lst)


i = 0
for dut in uniq_dut_lst:
    for temp_c in uniq_temperature_lst:
        for vio in uniq_vio_lst:
            tmp_df = raw_zorro_df.loc[
                (raw_zorro_df['COND_SERIAL_NUMBER'] == dut) &
                (raw_zorro_df[' COND_TEMP_C'] == temp_c) &
                (raw_zorro_df[' COND_VCTRL_HI_V'] == vio)]
            jmp_format_df.iloc[i, 0] = dut
            jmp_format_df.iloc[i, 1] = temp_c
            jmp_format_df.iloc[i, 2] = vio
            mhz, ns = dub_or_norm(tmp_df)
            jmp_format_df.iloc[i, 3] = mhz
            jmp_format_df.iloc[i, 4] = ns
            jmp_format_df.iloc[i, 5] = val_getter(tmp_df,
                                                  ' MEAS_MIN_READ_SCLK_' +
                                                  'PERIOD_MHZ',
                                                  'min_read_sclk_period')
            jmp_format_df.iloc[i, 7] = val_getter(tmp_df,
                                                  ' MEAS_MIN_READ_SCLK_' +
                                                  'PERIOD_NS',
                                                  'min_read_sclk_period')
            jmp_format_df.iloc[i, 8] = jmp_format_df.iloc[i, 7]-derate
            jmp_format_df.iloc[i, 6] = 1000.0/jmp_format_df.iloc[i, 8]
            jmp_format_df.iloc[i, 9] = val_getter(tmp_df,
                                                  ' MEAS_SDATA_SETUP',
                                                  'sdata_setup_time')
            jmp_format_df.iloc[i, 10] = val_getter(tmp_df,
                                                   ' MEAS_SDATA_HOLD',
                                                   'sdata_hold_time')
            jmp_format_df.iloc[i, 11] = val_getter(tmp_df,
                                                   ' MEAS_SCLK_MIN_HIGH_TIME',
                                                   'sclk_min_high_time')
            jmp_format_df.iloc[i, 12] = jmp_format_df.iloc[i, 11]-derate
            jmp_format_df.iloc[i, 13] = val_getter(tmp_df,
                                                   ' MEAS_SCLK_MIN_LOW_TIME',
                                                   'sclk_min_low_time')
            jmp_format_df.iloc[i, 14] = val_getter(tmp_df,
                                                   ' MEAS_SDATA_STABILIZATION',
                                                   'sdata_stabilization_time')
            jmp_format_df.iloc[i, 15] = jmp_format_df.iloc[i, 14]-derate
            jmp_format_df.iloc[i, 16] = val_getter(tmp_df,
                                                   ' MEAS_VOH',
                                                   'SDATA_VOH_VOL_1mA')
            jmp_format_df.iloc[i, 17] = val_getter(tmp_df,
                                                   ' MEAS_VOL',
                                                   'SDATA_VOH_VOL_1mA')
            jmp_format_df.iloc[i, 18] = val_getter(tmp_df,
                                                   ' MEAS_VOH',
                                                   'SDATA_VOH_VOL_2mA')
            jmp_format_df.iloc[i, 19] = val_getter(tmp_df,
                                                   ' MEAS_VOL',
                                                   'SDATA_VOH_VOL_2mA')
            jmp_format_df.iloc[i, 20] = val_getter(tmp_df,
                                                   ' MEAS_IIH_uA',
                                                   'SDATA_IIH_IIL')
            jmp_format_df.iloc[i, 21] = val_getter(tmp_df,
                                                   ' MEAS_IIL_uA',
                                                   'SDATA_IIH_IIL')
            jmp_format_df.iloc[i, 22] = val_getter(tmp_df,
                                                   ' MEAS_IIH_uA',
                                                   'SCLK_IIH_IIL')
            jmp_format_df.iloc[i, 23] = val_getter(tmp_df,
                                                   ' MEAS_IIL_uA',
                                                   'SCLK_IIH_IIL')
            jmp_format_df.iloc[i, 24] = val_getter(tmp_df,
                                                   ' MEAS_VIH_V',
                                                   'VIH_SDATA')
            jmp_format_df.iloc[i, 25] = val_getter(tmp_df,
                                                   ' MEAS_VIH_V',
                                                   'VIH_SCLK')
            jmp_format_df.iloc[i, 26] = val_getter(tmp_df,
                                                   ' MEAS_VIL_V',
                                                   'VIL_SDATA')
            jmp_format_df.iloc[i, 27] = val_getter(tmp_df,
                                                   ' MEAS_VIL_V',
                                                   'VIL_SCLK')
            jmp_format_df.iloc[i, 28] = val_getter(tmp_df,
                                                   ' MEAS_ICTRL_RF_OFF_uA',
                                                   'VIH_SCLK')
            i += 1


# %%
tabulated_df = deepcopy(jmp_format_df)
tabulated_df = tabulated_df.set_index(['VIO (V)', 'Temperature'])
tabulated_df = tabulated_df.drop(columns=['DUT',
                                          'ReadSCLK_MaxFreq (MHz)',
                                          'ReadSCLK_MinPeriod (ns)',
                                          'SCLK min high time (ns)',
                                          'Dstab time (ns)']
                                 ).astype(float)
# %%
a = tabulated_df.groupby(['VIO (V)', 'Temperature']).mean()
a['Calc'] = '0 Mean'
b = tabulated_df.groupby(['VIO (V)', 'Temperature']).std()
b['Calc'] = '1 Std Dev'
b = b.fillna(0.0)
c = tabulated_df.groupby(['VIO (V)', 'Temperature']).min()
c['Calc'] = '2 Min'
d = tabulated_df.groupby(['VIO (V)', 'Temperature']).max()
d['Calc'] = '3 Max'
tabulated_df = concat([a, b, c, d]).set_index('Calc', append=True)
# %%

tabulated_df = tabulated_df.stack()
tabulated_df = tabulated_df.swaplevel(1, 2)
tabulated_df = tabulated_df.unstack(level=[2, 1])
tabulated_df = tabulated_df.stack().sort_index(axis=1)
tabulated_df = tabulated_df.unstack()
# %%
# jmp_format_df.rename(columns={'DUT':''},inplace=True)

writer = ExcelWriter(output_path, engine='openpyxl')
book = load_workbook(output_path)
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
# print(writer.sheets)

raw_zorro_df.to_excel(writer, sheet_name='RawZorro', index=False)
jmp_format_df.to_excel(writer, sheet_name='JMP Format', index=False)
tabulated_df.to_excel(writer, sheet_name='Tabulated JMP')
writer.save()

xl = DispatchEx("Excel.Application")
wb = xl.workbooks.open(output_path)
xl.Visible = False
wb.Close(SaveChanges=1)
xl.Quit()

if ppt:
    tmp_path = make_jsl(src_jsl_func_path, output_path, output_path_ppt)
    run([jmp_path, tmp_path])