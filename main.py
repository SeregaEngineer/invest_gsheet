import pars_Dohod
import dividend_Dohod
import pasr_RBK
import time_update



ru_eu = 26
ru = 22

range_clear = f'F3:AS{ru_eu}'  # Все зона для очистки
range_RBK = f'D3:D{ru_eu}' # Данные с РБК как Американские так и Русские
range_dohod_rait = f'C3:C{ru_eu}' # Рейтинг с доход по акциям РУС и Амер
range_dohod_analyze = f'F3:F{ru}'  # Прогноз по выплатам дивидендов только  РУС




if __name__ == '__main__':
    pasr_RBK.clear_area(range_clear)
    pasr_RBK.main(range_RBK)
    pars_Dohod.main(range_dohod_rait)
    dividend_Dohod.main(range_dohod_analyze)
    time_update.send_time()
