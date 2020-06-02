import os
from math import ceil
import numpy as np
import matplotlib.pyplot as plt
import xlrd
from scipy.stats import kendalltau
import matplotlib.gridspec as gridspec

NB_GROUPS = 3
NB_SERIES = 4

def import_data_from_xlxs(wb):
    # Extract the name (used for graph saving)
    wb_name = ''.join(wb.split('.')[:-1]).replace(' ', '_')
    print(wb)
    # Open the workbook, then the desired sheet
    workbook = xlrd.open_workbook(wb)
    worksheet = workbook.sheet_by_name('Performances')

    # We extract the means
    # Group row col
    data = np.ndarray((3, 15, 4), dtype='float32')

    # For each row in the sheet
    for row in range(worksheet.nrows):
        # If the cells contains the word "Groupe", it's a new group (duh)
        if 'Groupe' in str(worksheet.cell_value(row, 0)):
            # We get the group number
            group_number = int(worksheet.cell_value(row, 0).split(' ')[-1]) - 1
            # In the arrays, the values starts 3 rows below the Group name cell
            # and one column to the right of the Group name cell
            base_row = row + 3
            base_col = 1

            current_row = base_row
            current_col = base_col

            # Indexes use for the data array
            col_idx = 0
            row_idx = 0

            # While we haven't reach the bottom of the group's values
            while worksheet.cell_value(current_row, current_col) != '':
                data[group_number, row_idx, col_idx] = worksheet.cell_value(current_row, current_col)

                col_idx += 1
                # Each mean column is separated by the std column, so we're
                # taking one out of two columns
                current_col += 2
                # If we're further than the 7th column, we go to the next row
                if current_col > 7:
                    current_row += 1
                    if current_row == worksheet.nrows:
                        break
                    current_col = 1
                    row_idx += 1
                    col_idx = 0

    return data, wb_name

def boxplot():
    # Take all the Excel workbooks in the current folder
    workbooks = [x for x in os.listdir('.') if 'xlsx' in x]

    # For each Excel workbook
    for wb in workbooks:
        data, wb_name = import_data_from_xlxs(wb)

        # Reshape into dict
        # GX : s1, s2, s3, s4
        data_dict = {'g1': {}, 'g2': {}, 'g3': {}}
        keys = list(data_dict.keys())

        for i in range(len(keys)):
            # number of col
            for j in range(data.shape[2]):
                data_dict[keys[i]][f's{j+1}'] = data[i, :, j]

        # /---------------\
        # |INTERGROUP DATA|
        # \---------------/
        intergroup_data = [[data_dict[f'g{x+1}'][f's{y+1}'] for x in range(NB_GROUPS)] for y in range(NB_SERIES)]

        fig = plt.figure(figsize=[12.8, 9.6])

        # Used to scale the y axes
        y_min = 0

        # Set the max value of the graph to the nearest 5 OR 2 OR 1 OR 0.1
        # (depending on the max value of the data). Could be done in a better way
        round_to = 10
        if np.max(data) < 30:
            round_to = 5
            if np.max(data) < 20:
                round_to = 2
                if np.max(data) < 10:
                    round_to = 1
                    if np.max(data) < 1:
                        round_to = 0.1


        y_max = round_to * ceil(np.max(data) / round_to)

        colours = ['lightcoral', 'yellowgreen', 'darkturquoise']

        for i in range(NB_SERIES):
            ls = fig.add_subplot(2, 2, i + 1)
            ls.yaxis.grid()
            bp = ls.boxplot(intergroup_data[i], patch_artist=True, showmeans=True)
            ls.set_ylim([y_min, y_max])

            plt.setp(bp['fliers'], markerfacecolor='r')
            plt.setp(bp['medians'], color='midnightblue')
            plt.setp(bp['means'], markerfacecolor='black', marker='d', markeredgecolor='black')

            # Set box filling colour for each box
            for patch, colour in zip(bp['boxes'], colours):
                patch.set_facecolor(colour)

            plt.xlabel('Groupe', fontsize=10)
            plt.ylabel('Distance', fontsize=10)
            plt.title(f'Moyenne intergroupe série {i+1}', fontsize=10)

        plt.savefig(f'{wb_name}_intergroupe.png', bbox_inches='tight')

        # /---------------\
        # |INTRAGROUP DATA|
        # \---------------/
        intragroup_data = [[data_dict[f'g{y+1}'][f's{x+1}'] for x in range(NB_SERIES)] for y in range(NB_GROUPS)]
        fig = plt.figure(figsize=[12.8, 9.6])
        y_min = 0

        round_to = 10
        if np.max(data) < 30:
            round_to = 5
            if np.max(data) < 20:
                round_to = 2
                if np.max(data) < 10:
                    round_to = 1
                    if np.max(data) < 1:
                        round_to = 0.1


        y_max = round_to * ceil(np.max(data) / round_to)

        colours = ['lightcoral', 'yellowgreen', 'darkturquoise', 'gold']

        gs = gridspec.GridSpec(2, 4)
        for i in range(NB_GROUPS):
            if i != 2:
                ls = fig.add_subplot(gs[0, i * 2 : i * 2 + 2])
            else:
                ls = fig.add_subplot(gs[1, 1 : 3])
            ls.yaxis.grid()
            bp = ls.boxplot(intragroup_data[i], patch_artist=True, showmeans=True)
            ls.set_ylim([y_min, y_max])

            plt.setp(bp['fliers'], markerfacecolor='r')
            plt.setp(bp['medians'], color='midnightblue')
            plt.setp(bp['means'], markerfacecolor='black', marker='d', markeredgecolor='black')

            for patch, colour in zip(bp['boxes'], colours):
                patch.set_facecolor(colour)

            plt.xlabel('Série', fontsize=10)
            plt.ylabel('Distance', fontsize=10)
            plt.title(f'Moyenne intragroupe groupe {i+1}', fontsize=10)

        plt.savefig(f'{wb_name}_intragroupe.png', bbox_inches='tight')


def correlation():
    # Take all the Excel workbooks in the current folder
    workbooks = [x for x in os.listdir('.') if 'xlsx' in x]

    data = []
    for wb in workbooks:
        data.append(import_data_from_xlxs(wb))

    data_dict = {'g1': {'precision': None, 'align_arm': None, 'elbow_move': None, 'javelin': None, 'leaning': None, 'pca': None},
                 'g2': {'precision': None, 'align_arm': None, 'elbow_move': None, 'javelin': None, 'leaning': None, 'pca': None},
                 'g3': {'precision': None, 'align_arm': None, 'elbow_move': None, 'javelin': None, 'leaning': None, 'pca': None}}

    group_keys = list(data_dict.keys())
    flaw_keys = list(data_dict['g1'].keys())

    # For each flaw
    for nb_k, fk in enumerate(flaw_keys):
        # get the data
        data_to_format = [d for d in data if fk in d[1].lower()][0]

        # for each group
        for i in range(NB_GROUPS):
            formated_data = np.zeros(data_to_format[0][i].shape[0] * data_to_format[0][i].shape[1])
            # for each serie, append values to 1D array
            for j in range(NB_SERIES):
                formated_data[j * data_to_format[0][i].shape[0]:
                              j * data_to_format[0][i].shape[0] + data_to_format[0][i].shape[0]] = data_to_format[0][i][:, j]
            # put the values of the flaw for the group into the dict
            data_dict[group_keys[i]][fk] = formated_data.copy()

    # row, col, other
    # flaw, flaw, groups
    correlations = np.ndarray((6, 6, 3))
    for k, gk in enumerate(data_dict.keys()):
        flaw_keys = list(data_dict[gk].keys())

        print(flaw_keys)

        for i in range(len(flaw_keys)):
            for j in range(len(flaw_keys)):
                corr = kendalltau(data_dict[gk][flaw_keys[i]], data_dict[gk][flaw_keys[j]])
                if i == j:
                    correlations[i, j, k] = -1
                correlations[i, j, k] = corr.correlation

        np.savetxt(f'{gk}_correlation_corr.csv', correlations[:,:,k], delimiter=' ', fmt='%1.4f')

if __name__ == '__main__':
    # boxplot()
    correlation()