import os
import time

from gen_xlsx_egis1 import export_subsystem_to_excel_egis1
from gen_xlsx_egis2 import export_subsystem_to_excel_egis2
from utils import init_conn

EGIS2_sys = 'СЕРВ'
TO_DO = [1, 2]

EGIS_2_sys_from_EGIS_1 = {
    'КОХР': ['АСКМ-А', 'ПАС-ПА', 'ПАС-ПВ', 'ПАС-ПЖ', 'ПАС-ПМ', 'ПБД-П', 'ПВАХ-ПД', 'ПВБ-П', 'ПВИ-А', 'ПВИ-В', 'ПВИ-Ж', 'ПВИ-ЗРП', 'ПВИ-М', 'ПЗП-А', 'ПЗП-В', 'ПИВ-КИАСК', 'ПИВ-КИСМО', 'ПИВ-ОРВД', 'ПМР-А', 'ПМР-В', 'ПМР-Ж', 'ПМР-М', 'ПОА-П', 'ПОД', 'ПОП', 'ПОШ-П', 'ПРП-П', 'ПСХ-МР', 'ПСХ-ОП', 'ПУД', 'ША', 'ШИ', 'ШМП', 'ШФП', 'ШЭ-НП'],
    'ОНСИ': ['АСКМ-А', 'ПБДТ-П'],
    'ПАНА': ['МС', 'МС-КПП', 'МС-КПП-В', 'МС-КПР', 'МС-КПР-В', 'МС-КПР-Ж', 'МС-КПР-М', 'МС-ПД', 'МС-ПД-В', 'МС-ПД-Ж', 'МС-ПД-М', 'ОПКП-В', 'ПД-РАТ', 'ПКП', 'ПКП-ПИ', 'ПКП-РТН', 'ПОА-П', 'ПОП', 'ПОС-П', 'ПРП-П', 'ПРС-П', 'ПСТ-П', 'ПУВ'],
    'ПУДО': ['АРПИ-П', 'ПЛК', 'ПУП-П'],
    'РВКП': ['ПЗП-А', 'ПЗП-В', 'ПЗП-Е', 'ПОШ-П'],
    'СЕРВ': ['ПАУ-П', 'ПДН-П', 'ПИБ-П', 'ПМО', 'ПРЗ-П', 'ПРЗ-П-2', 'ПРО-П', 'ПСП-П', 'ПТП-П']
}

SN_E1 = EGIS_2_sys_from_EGIS_1[EGIS2_sys]
OUTPUT_PATH_E1 = f"{os.path.abspath(os.curdir)}/out_dir/egis1/"

SN_E2 = [EGIS2_sys]
OUTPUT_PATH_E2 = f"{os.path.abspath(os.curdir)}/out_dir/egis2/"

if __name__ == "__main__":
    start_time = time.time()

    if 1 in TO_DO:
        for sys in SN_E1:
            export_subsystem_to_excel_egis1(sys, OUTPUT_PATH_E1)

    if 2 in TO_DO:
        for sys in SN_E2:
            export_subsystem_to_excel_egis2(sys, OUTPUT_PATH_E2)

    print("Время формирования файлов: {:.2f}s".format(time.time() - start_time))


