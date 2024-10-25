import os
import numpy as np
import pandas as pd
import settings as s

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
df = df.sort_values(by=['CycleStart_Internal'], ascending=False)
df = df[['Machine','Program','Pallet','PCount_Actual','CycleStart_Internal','CycleTime_Internal']]
# print(df)

df_programs = dict()
for k, v in df.groupby('Program'):
    v = v.reset_index(drop=True)
    v['Cycle_Minutes'] = pd.to_datetime(v['CycleTime_Internal'],format='%H:%M:%S')
    v['Cycle_Minutes'] = (
        round(v['Cycle_Minutes'].dt.hour*60 +
              v['Cycle_Minutes'].dt.minute +
              v['Cycle_Minutes'].dt.second/60,2))
    df_programs[k] = v


df_programs_keys = list(df_programs.keys())
print("Programs Found:", len(df_programs_keys))

for program in df_programs_keys:
    cycle_count = len(df_programs[program].index)
    newest_cycle = df_programs[program].iloc[0]['CycleStart_Internal']
    oldest_cycle = df_programs[program].iloc[cycle_count-1]['CycleStart_Internal']
    print("Program: " + program)
    print("Cycle Count:", cycle_count)
    print("Newest Cycle:", newest_cycle)
    print("Oldest Cycle:", oldest_cycle)
    print()
    # print(df_programs[program])

    # create df for cycle summary
    columns = ['Cycle Group Start Time', 'Cycle Group End Time', 'Median Length', 'Matching Cycles', 'Start Index', 'End Index']
    df_program_summary = pd.DataFrame(columns=columns)

    program_times = dict()
    #cycle_index
    i = 0
    while i < cycle_count:
        base_cycle = df_programs[program].iloc[i]['Cycle_Minutes']
        matches = 0
        matches_dict = dict()

        for j in range(1,cycle_count-i):
            test_cycle = df_programs[program].iloc[i+j]['Cycle_Minutes']

            if test_cycle < base_cycle * 1.05 and test_cycle > base_cycle * 0.95:
                matches_dict[i+j] = test_cycle
                matches += 1
            else:
                break

        if matches > 3:
            # summarize information and add to df
            program_times[i+j] = matches_dict
            cycle_group_start_time = df_programs[program].iloc[i+matches]['CycleStart_Internal']
            cycle_group_end_time = df_programs[program].iloc[i]['CycleStart_Internal']
            median_length = round(np.median(list(matches_dict.values())),2)
            matching_cycles = matches + 1
            start_index = i
            end_index = i + j
            df_program_summary.loc[i] = (cycle_group_start_time,
                                     cycle_group_end_time,
                                     median_length,
                                     matching_cycles,
                                     start_index,
                                     end_index)

            # print all match groups
            # print("Cycle Group Start: " + df_programs[program].iloc[i+matches]['CycleStart_Internal'])
            # print("Cycle Group End: " + df_programs[program].iloc[i]['CycleStart_Internal'])
            # print("Median Length: " + str(df_programs[program].iloc[i]['Cycle_Minutes']))
            # print("Matching Cycles:", matches)
            # print()

            # skip current matched cycles
            i = i + matches - 1

        else:
            i += 1

    # print(df_programs[program].iloc[0])
    df_program_summary.reset_index(drop=True, inplace=True)
    print("Current Group Cycle Time: " + str(df_program_summary.iloc[0]['Median Length']))
    print("Shortest Group Cycle Time: " + str(df_program_summary.loc[df_program_summary['Median Length'].idxmin()]['Median Length']))
    print("Longest Group Cycle Time: " + str(df_program_summary.loc[df_program_summary['Median Length'].idxmax()]['Median Length']))


    print()
    print(df_program_summary)


    #input("Press Enter to continue...")
