import xlrd, xlwt
import matplotlib.pyplot as plt
import numpy as np

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    for num in range(1, 6):
        # открываем файл
        rb = xlrd.open_workbook(f'data/data_{num}.xls', formatting_info=True)
        # выбираем активный лист
        sheet = rb.sheet_by_index(0)

        # получаем список значений из всех записей
        # sheet.nrows некоректно
        vals = np.array([sheet.row_values(rownum) for rownum in range(sheet.nrows)])
        vals = vals[:np.where(vals[:, 0] == '')[0][0] - 1]

        freq = np.array([float(val[0]) for val in vals[1:]])
        channel_1 = np.array([float(val[1]) for val in vals[1:]])
        channel_set = np.array([float(val[4]) for val in vals[1:]])
        amplitude = np.array([channel_set[i] / channel_1[i] for i in range(len(channel_set))])

        plt.cla()
        plt.clf()
        plt.grid()
        FREQ = 1500
        FREQ_NUM = np.where(freq >= FREQ)[0][0]
        plt.plot(freq[:FREQ_NUM], amplitude[:FREQ_NUM])
        plt.yscale('log')
        plt.savefig(f'res/data_{num}.png', dpi=200)

        np.savetxt(f'res/data_{num}.txt', np.array([[it1, it2] for it1, it2 in zip(freq, amplitude)]))
        np.save(f'res/data_{num}', np.array([[it1, it2] for it1, it2 in zip(freq, amplitude)]))
