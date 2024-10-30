import os
import time
import numpy as np
import pandas
import pandas as pd
import settings as s
import shutil
from pathlib import Path
from analyse_all_programs_report import AnalyseAllPrograms

# init objects
analyse_all_programs_report = AnalyseAllPrograms()
    
# instructions for user
print("--- FMS Cycle Time Analysis --- " + s.version)
print("Determines if each program is running at its shortest cycle.")
print("Export FMS Cycles from TTM for any date range, check 'include complete cycles only'.")
print("Copy exported files to C:\\programcycles")
print()

while(True):
    input("Press Enter to process files...")
    print("Reading exported files...")

    # delete existing reports folder
    if os.path.exists(s.reports_folder):
        shutil.rmtree(s.reports_folder)

    # pandas settings
    pd.set_option('display.max_columns', 18)
    pd.set_option('display.width', 400)

    file_data_appended = []

    file_list = os.listdir(s.program_cycles_folder)
    for file in file_list:
        data = pd.read_excel(s.program_cycles_folder + file)
        file_data_appended.append(data)

    df = pd.concat(file_data_appended, ignore_index=True)
    df = df.drop_duplicates(keep='first')
    df = df.sort_values(by=[s.cycle_start], ascending=False)
    df = df[['Machine','Program','Pallet','PCount_Actual',s.cycle_start,s.cycle_time]]
    df[s.cycle_start] = pd.to_datetime(df[s.cycle_start]).dt.date
    # print(df)

    df_programs = dict()
    for k, v in df.groupby('Program'):
        v = v.reset_index(drop=True)
        v['Cycle_Minutes'] = pd.to_datetime(v[s.cycle_time],format='%H:%M:%S', errors='coerce')
        v['Cycle_Minutes'] = (
            round(v['Cycle_Minutes'].dt.hour*60 +
                  v['Cycle_Minutes'].dt.minute +
                  v['Cycle_Minutes'].dt.second/60,2))
        df_programs[k] = v


    df_programs_keys = list(df_programs.keys())
    programs_found = len(df_programs_keys)
    print(" Programs Found:", programs_found)

    programs_processed = 0
    df_program_groups_dict = dict()
    df_all_programs_report = pd.DataFrame(
        columns=['Program', 'Current Group Cycle', s.current_group_date, 'Shortest Group Cycle', s.shortest_group_date, 'Longest Group Cycle'])

    for program in df_programs_keys:
        # progress count
        programs_processed = programs_processed + 1
        print('\r', "Processing: " + str(int(100 * round(programs_processed/programs_found,2))) + "%", end='')
        time.sleep(.01)

        cycle_count = len(df_programs[program].index)
        newest_cycle = df_programs[program].iloc[0][s.cycle_start]
        oldest_cycle = df_programs[program].iloc[cycle_count-1][s.cycle_start]
        # print("Program: " + program)
        # print("Cycle Count:", cycle_count)
        # print("Newest Cycle:", newest_cycle)
        # print("Oldest Cycle:", oldest_cycle)
        # print()
        # print(df_programs[program])

        # create df for cycle summary and report
        df_program_groups = pd.DataFrame(columns=['Cycle Group Start Time', 'Cycle Group End Time', 'Median Length', 'Matching Cycles', 'Start Index', 'End Index'])


        program_times = dict()
        #cycle_index
        i = 0

        while i < cycle_count:
            base_cycle = df_programs[program].iloc[i]['Cycle_Minutes']
            matches = 0
            matches_dict = dict()
            test_list = list()
            median_test_cycle = 0

            # add the current cycle and next 4 to a list and find the median
            for j in range(0,cycle_count-i):
                test_cycle = df_programs[program].iloc[i+j]['Cycle_Minutes']
                test_list.append(test_cycle)

                if len(test_list) >4:
                    median_test_cycle = np.median(test_list)
                    break;

            # test if all cycles in the list are within range of the median
            for k in range(0, len(test_list)):
                if test_list[k] < median_test_cycle * 1.05 and test_list[k] > median_test_cycle * 0.95:
                    matches_dict[i+k] = test_list[k]
                    matches += 1

           # if a pattern is found, add subsequent matching cycles and record information in the df
            if matches > 4:
                for j in range(5, cycle_count - i - 5):
                    test_cycle = df_programs[program].iloc[i + j]['Cycle_Minutes']
                    if test_cycle < median_test_cycle * 1.05 and test_cycle > median_test_cycle * 0.95:
                        matches_dict[i + j] = test_cycle
                        matches += 1
                    else:
                        break

                # summarize information and add to df
                program_times[i] = matches_dict
                cycle_group_start_time = df_programs[program].iloc[i+matches][s.cycle_start]
                cycle_group_end_time = df_programs[program].iloc[i][s.cycle_start]
                median_length = round(np.median(list(matches_dict.values())),2)
                matching_cycles = matches + 1
                start_index = i
                end_index = i + j
                df_program_groups.loc[i] = (cycle_group_start_time,
                                         cycle_group_end_time,
                                         median_length,
                                         matching_cycles,
                                         start_index,
                                         end_index)

                # print all match groups
                # print("Cycle Group Start: " + df_programs[program].iloc[i+matches][s.cycle_start])
                # print("Cycle Group End: " + df_programs[program].iloc[i][s.cycle_start])
                # print("Median Length: " + str(df_programs[program].iloc[i]['Cycle_Minutes']))
                # print("Matching Cycles:", matches)
                # print()

                # skip current matched cycles
                i = i + matches - 1

            else:
                i += 1

        # print(df_programs[program].iloc[0])
        df_program_groups.reset_index(drop=True, inplace=True)
        if len(df_program_groups.index) > 0:
            current_group_cycle = df_program_groups.iloc[0]['Median Length']
            current_group_cycle_date = df_program_groups.iloc[0]['Cycle Group Start Time']
            shortest_group_cycle = df_program_groups.loc[df_program_groups['Median Length'].idxmin()]['Median Length']
            shortest_group_cycle_date = df_program_groups.loc[df_program_groups['Median Length'].idxmin()]['Cycle Group Start Time']
            longest_group_cycle = df_program_groups.loc[df_program_groups['Median Length'].idxmax()]['Median Length']

        else:
            current_group_cycle = 0
            shortest_group_cycle = 0
            longest_group_cycle = 0

        df_all_programs_report.loc[i] = (program, current_group_cycle, current_group_cycle_date, shortest_group_cycle, shortest_group_cycle_date, longest_group_cycle)

        # print()
        # print(df_program_groups)
        # input("Press Enter to continue...")

    df_all_programs_report.reset_index(drop = True, inplace=True)
    df_longer_cycles_report = analyse_all_programs_report.find_longer_cycles(df_all_programs_report)
    df_no_groups_programs = df_all_programs_report[df_all_programs_report['Current Group Cycle'] == 0].copy()

    # print("All Programs Report")
    # print(df_all_programs_report)
    # print()
    # print("Longer Cycles Report")
    # print(df_longer_cycles_report)
    # print()
    # print("Programs With No Cycle Groups Found")
    # print(df_no_groups_programs)
    # print()

    report_dict = {'longer cycles': df_longer_cycles_report, 'all programs': df_all_programs_report, 'no groups': df_no_groups_programs}
    Path("C:/programcycles/reports").mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(s.reports_folder + s.reports_file, engine = 'xlsxwriter') as writer:

        for sheetname, df in report_dict.items():  # loop through `dict` of dataframes
            df.to_excel(writer, sheet_name=sheetname)  # send df to writer
            worksheet = writer.sheets[sheetname]  # pull worksheet object
            worksheet.autofit()
            # for idx, col in enumerate(df):  # loop through all columns
            #     series = df[col]
            #     max_len = max((
            #         series.astype(str).map(len).max(),  # len of largest item
            #         len(str(series.name))  # len of column name/header
            #         )) + 1  # adding a little extra space
            #     worksheet.set_column(idx, idx, max_len)  # set column width


    print()
    print("Report saved to: " + s.reports_folder)
    print()
    #df_all_programs_report.to_excel(s.reports_folder + s.all_report_output)
    #df_longer_cycles_report.to_excel(s.longer_cycles_output)
    #df_no_groups_programs.to_excel(s.no_groups_output)
