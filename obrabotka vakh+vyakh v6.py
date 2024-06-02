import time
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
import scipy
from lmfit import Model
import os
from lmfit import Model
from scipy.stats import norm
import statistics

path = 'C:/Users/Анастасия/Downloads/analiz/975tcg integr_time 2'


def files_search(path: str, file_extension: str) -> list:
    files_list = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(file_extension):
                # if file != 'result.xlsx' and file != 'result1.xlsx' and file != 'result 975tcg.xlsx' and file != 'result with calibrate.xlsx' and file != 'result with calibrate 1.xlsx':
                    path = os.path.join(root, file)
                    files_list.append(path)

    return files_list


def file_name(dir: str) -> str:
    slash_index = []
    for n in range(len(dir)):
        if dir[n] == str('\\'):
            slash_index.append(n)
    return dir[slash_index[-1]+1:-5]


def drive_voltage_current(current_data: list, voltage_data: list):
    for i in range(len(data)):
        integrate_current = np.trapz(y=current_data[0:i], x=voltage_data[0:i], dx=0.01)

        if integrate_current != 0:
            return data['Voltage, V'][i-1], i-1


def drive_voltage_intensity(data: list, voltage_data:list):
    for n in range(len(data)):
        wavelengths = data.columns[5:]
        intensity_wavelengths = data.iloc[n, 5:]

        integrate_spectre = np.trapz(y=intensity_wavelengths, x=wavelengths, dx=0.01)

        if integrate_spectre > 5000:
            return voltage_data[n-1], n-1


display_char = {
    'coordinates': [],
    'diode': [],
    'driving_voltage_intensity': [],
    'driving_voltage_current': [],
    'driving_current_intensity': [],
    'driving_current_current': [],
    'FoM': []
}

data_calibrate_curve = pd.read_excel('Cd calibrate.xlsx')
files_list = files_search(path=path, file_extension='.xlsx')

for file in range(len(files_list)):
    data = pd.read_excel(files_list[file])

    display_number = file_name(files_list[file])

    linear_mod = Model(lambda x, slope, intercept: x * slope + intercept)
    params = linear_mod.make_params(slope=20, intercept=-40)

    res_cur_lum = linear_mod.fit(
        data['Intensity'], params,
        x=data['Current, mA'])
    koef, none = res_cur_lum.params['slope'].value, res_cur_lum.params[
        'intercept'].value

    try:
        float(display_number[0])
    except:
        None
    else:

        on_voltage_current, index_on_voltage_current = drive_voltage_current(data['Current, mA'], data['Voltage, V'])
        on_voltage_intensity, index_on_voltage_intensity = drive_voltage_intensity(data, data['Voltage, V'])

        brightness_array = []
        for n in range(len(data)):
            wavelengths = data.columns[5:]
            intensity_wavelengths = data.iloc[n, 5:]

            calibrating_intensity = []
            for i in range(len(intensity_wavelengths)):
                calibrating_intensity.append(intensity_wavelengths[i] * data_calibrate_curve['Lum eff'][i])

            brightness_array.append(
                np.trapz(y=calibrating_intensity, x=wavelengths, dx=0.01) / 5.51219E7 * 443)

            # brightness_array.append(
            #     np.trapz(y=intensity_wavelengths, x=wavelengths, dx=0.01))

        FoM_brightness_current = []
        for a in range(len(data)):
            if data['Current, mA'][a] != 0:
                FoM_brightness_current.append(brightness_array[a] / data['Current, mA'][a])

        display_char['coordinates'].append(display_number)

        if on_voltage_intensity - on_voltage_current > 1.5:
            display_char['diode'].append('NO')
        else:
            display_char['diode'].append('YES')

        display_char['driving_voltage_intensity'].append(on_voltage_intensity)
        display_char['driving_voltage_current'].append(on_voltage_current)
        display_char['driving_current_intensity'].append(data['Current, mA'][index_on_voltage_intensity])
        display_char['driving_current_current'].append(data['Current, mA'][index_on_voltage_current])
        display_char['FoM'].append(np.max(FoM_brightness_current))

        print('**********************************************')
        print(f"Display: {display_char['coordinates'][-1]}")
        print(f"Diode: {display_char['diode'][-1]}")
        print(f"FoM: {round(display_char['FoM'][-1], 0)}")
        print(f"Vdc, Idc: {display_char['driving_voltage_current'][-1]} V, {display_char['driving_current_current'][-1]} mA")
        print(f"Vdi, Idi: {display_char['driving_voltage_intensity'][-1]} V, {display_char['driving_current_intensity'][-1]} mA")
        print('**********************************************')

        plt.figure(figsize=(15, 6))
        plt.subplots_adjust(wspace=0.5)

        plt.subplot(1, 3, 1)
        plt.title(f'Вольт-амперная характеристика')
        plt.plot(data['Voltage, V'], data['Current, mA'], 'o', markersize=9, mfc='none', label='Linear current',
                 color='#0080FF')
        plt.tick_params(axis='y', labelcolor='#0080FF')
        plt.xlabel('Voltage, V')
        plt.ylabel('Linear current, mA', color='#0080FF')
        plt.ylim([0, np.max(data['Current, mA'] * 1.1)])

        plt.twinx()

        plt.axvline(x=on_voltage_current, color='black')
        plt.plot(data['Voltage, V'], data['Current, mA'], 'o', markersize=9, mfc='none', label='Log current',
                 color='#800000')

        plt.tick_params(axis='y', labelcolor='#800000')
        plt.semilogy()
        plt.xlim([0, 5])
        plt.xlabel('Voltage, V')
        plt.ylabel('Log current, mA', color='#800000')

        plt.subplot(1, 3, 2)
        plt.plot(data['Current, mA'], brightness_array, 'o', markersize=9, mfc='none', color='#0080FF')

        plt.xlabel('Current, mA')
        plt.ylabel('Brightness, Cd/m^2')
        plt.xlim([0, np.max(data['Current, mA'])])
        plt.ylim([0, np.max(brightness_array)])
        plt.legend([f'FoM: {int(np.max(FoM_brightness_current))}'])

        plt.subplot(1, 3, 3)
        plt.title('Спектр при управляющем напряжении')
        plt.plot(data.columns[5:], data.iloc[index_on_voltage_current, 5:], '-', markersize=5, mfc='none', color='#0080FF',
                 alpha=0.2)
        plt.plot(data.columns[5:],
                 scipy.signal.savgol_filter(data.iloc[index_on_voltage_current, 5:], window_length=100, polyorder=5), '-',
                 linewidth=3, color='#0080FF')
        plt.tick_params(axis='y', labelcolor='#0080FF')
        plt.xlabel('Wavelengths, nm')
        plt.ylabel('Min Intensity', color='#0080FF')
        plt.ylim([0, np.max(data.iloc[index_on_voltage_current, 5:]) * 1.1])

        plt.twinx()

        plt.plot(data.columns[5:], data.iloc[len(data) - 1, 5:], '-', markersize=5, mfc='none', color='#800000',
                 linewidth=3)
        plt.tick_params(axis='y', labelcolor='#800000')
        plt.xlabel('Wavelengths, nm')
        plt.ylabel('Max Intensity', color='#800000')
        plt.ylim([0, np.max(data.iloc[len(data) - 1, 5:]) * 1.1])

        plt.show()

df = pd.DataFrame(list(zip(display_char['driving_voltage_current'],
                           display_char['driving_voltage_intensity'],
                           display_char['driving_current_current'],
                           display_char['driving_current_intensity'],
                           display_char['FoM'],
                           display_char['diode'],
                           display_char['coordinates'])), columns=[
    'Vd_cur',
    'Vd_int',
    'Id_cur',
    'Id_int',
    'FoM',
    'diode',
    'name'
])

df.to_excel(path + '/' + 'result with calibrate 2.xlsx', index=False)

