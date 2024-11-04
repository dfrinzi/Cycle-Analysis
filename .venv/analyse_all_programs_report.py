import pandas as pd
import settings as s

class AnalyseAllPrograms:
    
    def __init__(self):
        pass

    def find_longer_cycles(self, df_all_programs_report):
        df_longer_cycles_report = (
            df_all_programs_report[df_all_programs_report[s.current_group_cycle] >
            s.current_to_short_limit * df_all_programs_report[s.shortest_group_cycle]].copy())
        df_longer_cycles_report = df_longer_cycles_report[df_longer_cycles_report[s.current_part_count] ==
                                                          df_longer_cycles_report[s.shortest_part_count]].copy()
        return df_longer_cycles_report