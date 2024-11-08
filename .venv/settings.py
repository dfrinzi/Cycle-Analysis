version = "Version 1.0 - 10/30/2024"

program_cycles_folder  = "C:/programcycles/"
reports_folder = "C:/programcycles/reports/"
reports_file = "CycleReports.xlsx"
all_report_output = "all_programs.xlsx"
longer_cycles_output = "longer_cycles.xlsx"
no_groups_output = "no_program_groups.xlsx"

# range for matching cycles into a group
high_match_limit = 1.05
low_match_limit = 0.95
# limit for reporting that the current cycle is longer than the shortest
current_to_short_limit = 1.03

# column names
program = "Program"
machine = "Machine"
pallet = "Pallet"
part_count = "PCount_Actual"
cycle_length_minutes = "Cycle Minutes"
cycle_start = "CycleStart_Full"
cycle_time = "CycleTime_Full"

current_group_cycle = "Current Group Cycle"
current_group_date = "Current Date"
current_part_count = "Current Part Count"
shortest_group_cycle = "Shortest Group Cycle"
shortest_group_date = "Shortest Date"
shortest_part_count = "Shortest Part Count"
longest_group_cycle = "Longest Group Cycle"
longest_group_date = "Longest Date"
longest_part_count = "Longest  Part Count"
cycle_group_start_time = "Cycle Group Start Time"
cycle_group_end_time = "Cycle Group End Time"
median_length = "Median Length"
matching_cycles = "Matching Cycles"
start_index = "Start Index"
end_index = "End Index"

