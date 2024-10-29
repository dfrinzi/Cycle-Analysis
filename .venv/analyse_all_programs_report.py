import pandas as pd
import settings as s

class AnalyseAllPrograms:
    
    def __init__(self):
        pass

    def find_longer_cycles(self, df_all_programs_report):
        df_longer_cycles_report = pd.DataFrame(
            columns=['Program', 'Current Group Cycle', 'Shortest Group Cycle', 'Longest Group Cycle'])

        df_longer_cycles_report = df_all_programs_report[df_all_programs_report['Current Group Cycle'] > 1.05 * df_all_programs_report['Shortest Group Cycle']].copy()

        return df_longer_cycles_report