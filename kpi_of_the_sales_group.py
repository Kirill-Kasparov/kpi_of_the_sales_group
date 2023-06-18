import pandas as pd    # by Kirill Kasparov 2023
import datetime
import calendar
import numpy as np
import dash
from dash import dcc as dcc
from dash import html as html
import plotly.express as px
import plotly.graph_objs as go

def kpi_of_the_sales_group():
    # Собираем базы
    def workind_days():
        # Получаем дату из файла
        creation_time = pd.read_excel('month_otgr.xls', sheet_name='Лист ТРП', header=1, nrows=1)
        creation_time = creation_time.columns[0].replace('Отчётная дата: ', '').split()
        creation_time = creation_time[0].split('.')
        ddf = datetime.date(int(creation_time[2]), int(creation_time[1]), int(creation_time[0]))

        # Список праздничных дней
        holidays = pd.read_excel('holidays.xlsx')
        # Конвертируем дату из Excel формата в datetime формат
        for col in holidays.columns:
            holidays[col] = pd.to_datetime(holidays[col]).dt.date

        # Начальная и конечная дата текущего месяца
        start_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')), 1)
        now_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')), int(ddf.strftime('%d')))
        end_date = datetime.date(int(ddf.strftime('%Y')), int(ddf.strftime('%m')),
                                 calendar.monthrange(now_date.year, now_date.month)[1])

        # Вычисляем количество рабочих дней
        now_working_days = 0
        end_working_days = 0

        current_date = start_date
        while current_date <= end_date:
            # Если текущий день не является выходным или праздничным, увеличиваем счетчик рабочих дней
            if current_date in list(holidays['working_days']):  # рабочие дни исключения
                now_working_days += 1
                end_working_days += 1
            elif current_date.weekday() < 5 and current_date not in list(holidays[int(ddf.strftime('%Y'))]):
                if current_date <= now_date:
                    now_working_days += 1
                end_working_days += 1
            current_date += datetime.timedelta(days=1)

        # Получаем данные прошлого года
        start_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')), 1)
        now_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')), int(ddf.strftime('%d')))
        end_date_2 = datetime.date(int(ddf.strftime('%Y')) - 1, int(ddf.strftime('%m')),
                                   calendar.monthrange(now_date_2.year, now_date_2.month)[1])
        now_working_days_2 = 0
        end_working_days_2 = 0

        current_date = start_date_2
        while current_date <= end_date_2:
            # Если текущий день не является выходным или праздничным, увеличиваем счетчик рабочих дней
            if current_date in list(holidays['working_days']):  # рабочие дни исключения
                now_working_days_2 += 1
                end_working_days_2 += 1
            elif current_date.weekday() < 5 and current_date not in list(holidays[int(ddf.strftime('%Y')) - 1]):
                if current_date < now_date_2:
                    now_working_days_2 += 1
                end_working_days_2 += 1
            current_date += datetime.timedelta(days=1)
        date_lst = [start_date, start_date_2, now_date, now_date_2, end_date, end_date_2, now_working_days,
                    now_working_days_2, end_working_days, end_working_days_2]
        # print('Начало месяца:', start_date, 'в прошлом году', start_date_2)
        # print('Отчетная дата:', now_date, 'в прошлом году', now_date_2)
        # print('Конец месяца', end_date, 'в прошлом году', end_date_2)
        # print('Рабочих дней сейчас', now_working_days, 'в прошлом году', now_working_days_2)
        # print('Всего рабочих дней', end_working_days, 'в прошлом году', end_working_days_2)
        return date_lst
    def data_trp():
        # Загружаем базу 'Лист ТРП'
        df_trp = pd.read_excel('month_otgr.xls', sheet_name='Лист ТРП', header=14)
        df_trp.drop(df_trp.tail(1).index, inplace=True)  # удаляем строку Итого
        # получаем список месяцев от даты отчета
        col_for_total_result = [now_date.strftime('%d.%m.%Y')]  # вместо list(df.columns[13:0:-1])
        for i in range(1, 13):
            prev_month = now_date - datetime.timedelta(days=28 * i)
            while prev_month.strftime('%m.%Y') in col_for_total_result or prev_month.strftime(
                    '%m.%Y') == now_date.strftime(
                    '%m.%Y'):  # поправка на дни
                prev_month = prev_month - datetime.timedelta(days=1)
            col_for_total_result.append(prev_month.strftime('%m.%Y'))
        # добавляем месяца, вместо столбцов Т, -1, -2...
        count = 0
        for i in df_trp.columns[4:29:2]:
            df_trp[col_for_total_result[count]] = df_trp[i]
            count += 1
        # добавляем ВП, вместо столбцов КТН
        count = 0
        for i in df_trp.columns[5:30:2]:
            df_trp['ВП ' + str(col_for_total_result[count])] = df_trp.iloc[:, 32 + count] - (df_trp.iloc[:, 32 + count] / df_trp[i])
            df_trp['ВП ' + str(col_for_total_result[count])] = df_trp['ВП ' + str(col_for_total_result[count])].fillna(0)
            df_trp['ВП ' + str(col_for_total_result[count])] = df_trp['ВП ' + str(col_for_total_result[count])].replace(np.inf, 0)
            count += 1
        # Удаляем лишние столбцы с 4 по 31 включительно
        df_trp = df_trp.iloc[:, :4].join(df_trp.iloc[:, 32:], how='outer')
        # Группируем по столбцу "Название ТС", отсекаем столбцы до него
        df_trp = df_trp.groupby('Название ТС')[df_trp.columns[4:]].sum().reset_index()
        # Добавляем строку суммы всех колонок
        sums = df_trp.select_dtypes(include=['number']).sum()    # Считаем сумму
        totals = pd.DataFrame([['Итого'] + sums.tolist()], columns=df_trp.columns)    # собираем ДФ к строке Итого
        df_trp = pd.concat([df_trp, totals], ignore_index=True)    # добавляем строку в общий ДФ
        # Сортируем список по сумме текущего месяца
        df_trp = df_trp.sort_values(by=now_date.strftime('%d.%m.%Y'), ascending=False)
        # Добавляем КТН
        count = 0
        for i in df_trp.columns[1:14]:
            df_trp['КТН ' + str(col_for_total_result[count])] = round(df_trp[i] / (df_trp[i] - df_trp.iloc[:, 14 + count]), 3)
            df_trp['КТН ' + str(col_for_total_result[count])] = df_trp['КТН ' + str(col_for_total_result[count])].fillna(0)
            df_trp['КТН ' + str(col_for_total_result[count])] = df_trp['КТН ' + str(col_for_total_result[count])].replace(np.inf, 0)
            count += 1
        # Добавляем доли отгрузок
        df_trp['Доля на ' + df_trp.columns[1]] = round(df_trp[df_trp.columns[1]] / df_trp.iloc[0, 1], 4)
        df_trp['Доля на ' + df_trp.columns[13]] = round(df_trp[df_trp.columns[13]] / df_trp.iloc[0, 13], 4)
        df_trp['Отклонение долей'] = df_trp['Доля на ' + df_trp.columns[1]] - df_trp['Доля на ' + df_trp.columns[13]]
        # Добавляем прогноз прироста
        df_trp['Прогноз выполнения'] = round(df_trp[df_trp.columns[1]] / now_working_days * end_working_days, 2)
        df_trp['Прогноз прироста в руб.'] = df_trp['Прогноз выполнения'] - round((df_trp[df_trp.columns[13]] / end_working_days_2 * end_working_days), 2)
        df_trp['Прогноз прироста в %'] = round((df_trp['Прогноз выполнения'] / (df_trp[df_trp.columns[13]] / end_working_days_2 * end_working_days) - 1) * 100, 2)
        df_trp['Отклонение отгрузок по ССП(3 мес)'] = round(df_trp['Прогноз выполнения'] - (df_trp.iloc[:, 2] + df_trp.iloc[:, 3] + df_trp.iloc[:, 4]) / 3)
        df_trp['Прогноз выполнения ВП'] = round(df_trp[df_trp.columns[14]] / now_working_days * end_working_days, 2)
        df_trp['Прогноз прироста ВП в руб.'] = df_trp['Прогноз выполнения ВП'] - round(
            (df_trp[df_trp.columns[26]] / end_working_days_2 * end_working_days), 2)
        df_trp['Прогноз прироста ВП в %'] = round((df_trp['Прогноз выполнения ВП'] / (
                    df_trp[df_trp.columns[26]] / end_working_days_2 * end_working_days) - 1) * 100, 2)
        df_trp['Прирост КТН'] = df_trp.iloc[:, 27] - df_trp.iloc[:, 39]
        df_trp.to_excel('Лист ТРП.xlsx', index=False)
        return df_trp

    start_date, start_date_2, now_date, now_date_2, end_date, end_date_2, now_working_days, now_working_days_2, end_working_days, end_working_days_2 = workind_days()

    # Для листа ТРП
    df_trp = data_trp()
    # Формируем графики
    options_df_trp = [{'label': i, 'value': i} for i in df_trp['Название ТС'].unique()]
    def fig_bar_trp_otgr():
        x_bar_trp_otgr_tr = df_trp.columns[1:14]
        y_bar_trp_otgr_tr = [round(i, 2) for i in list(df_trp.iloc[0, 1:14] / 1000000)]    # округляем до двух знаков
        total_result = {'Месяц': x_bar_trp_otgr_tr[::-1], 'Отгрузка, млн. руб.': y_bar_trp_otgr_tr[::-1]}
        # строим график
        fig1 = px.line(total_result, x='Месяц', y='Отгрузка, млн. руб.', title='Динамика оборота по месяцам', text='Отгрузка, млн. руб.')
        fig1.update_layout(font=dict(size=18))  # увеличиваем шрифт
        return fig1
    def fig_bar_trp_vp():
        x_bar_trp_vp = df_trp.columns[14:27]
        y_bar_trp_vp = [round(i, 2) for i in list(df_trp.iloc[0, 14:27] / 1000000)]    # округляем до двух знаков
        total_result = {'Месяц': x_bar_trp_vp[::-1], 'Валовая прибыль, млн. руб.': y_bar_trp_vp[::-1]}
        # строим график
        fig1 = px.line(total_result, x='Месяц', y='Валовая прибыль, млн. руб.', title='Динамика валовой прибыли по месяцам', text='Валовая прибыль, млн. руб.')
        fig1.update_layout(font=dict(size=18))  # увеличиваем шрифт
        return fig1
    def fig_bar_trp_otgr_tr():
        if len(df_trp['Название ТС'].iloc[1:]) > 15:
            top = 12
        else:
            top = len(df_trp['Название ТС'].iloc[1:])
        x_bar_trp_otgr_tr = df_trp['Название ТС'].iloc[1:top]
        y_bar_trp_otgr_tr = df_trp['Прогноз выполнения'].iloc[1:top]
        y_bar_trp_otgr_tr2 = df_trp[df_trp.columns[13]].iloc[1:top]
        fig_bar_trp_otgr_tr = px.bar()
        fig_bar_trp_otgr_tr.add_trace(go.Bar(x=x_bar_trp_otgr_tr, y=y_bar_trp_otgr_tr, name=str(now_date.strftime('%m.%Y')), text=round(y_bar_trp_otgr_tr / 1000000, 2)))
        fig_bar_trp_otgr_tr.add_trace(go.Bar(x=x_bar_trp_otgr_tr, y=y_bar_trp_otgr_tr2, name=str(df_trp.columns[13]), text=round(y_bar_trp_otgr_tr2 / 1000000, 2)))
        fig_bar_trp_otgr_tr.update_layout(barmode='group')
        fig_bar_trp_otgr_tr.update_layout(font=dict(size=18))  # увеличиваем шрифт
        fig_bar_trp_otgr_tr.update_layout(xaxis_title="Названия ТР", yaxis_title="Сумма отгрузки")
        fig_bar_trp_otgr_tr.update_layout(height=600)
        return fig_bar_trp_otgr_tr
    def fig_bar_trp_otgr_tr2():
        if len(df_trp['Название ТС'].iloc[1:]) > 20:
            top = 22
        else:
            top = len(df_trp['Название ТС'].iloc[12:])
        x_bar_trp_otgr_tr = df_trp['Название ТС'].iloc[12:top]
        y_bar_trp_otgr_tr = df_trp['Прогноз выполнения'].iloc[12:top]
        y_bar_trp_otgr_tr2 = df_trp[df_trp.columns[13]].iloc[12:top]
        fig_bar_trp_otgr_tr = px.bar()
        fig_bar_trp_otgr_tr.add_trace(go.Bar(x=x_bar_trp_otgr_tr, y=y_bar_trp_otgr_tr, name=str(now_date.strftime('%m.%Y')), text=round(y_bar_trp_otgr_tr / 1000000, 2)))
        fig_bar_trp_otgr_tr.add_trace(go.Bar(x=x_bar_trp_otgr_tr, y=y_bar_trp_otgr_tr2, name=str(df_trp.columns[13]), text=round(y_bar_trp_otgr_tr2 / 1000000, 2)))
        fig_bar_trp_otgr_tr.update_layout(barmode='group')
        fig_bar_trp_otgr_tr.update_layout(font=dict(size=18))  # увеличиваем шрифт
        fig_bar_trp_otgr_tr.update_layout(xaxis_title="Названия ТР", yaxis_title="Сумма отгрузки")
        fig_bar_trp_otgr_tr.update_layout(height=600)
        return fig_bar_trp_otgr_tr
    def fig_pie_trp_otgr_tr_now():
        pie_values = df_trp['Доля на ' + df_trp.columns[1]].iloc[1:]
        pie_names = df_trp['Название ТС'].iloc[1:]
        pie_labels = df_trp['Отклонение долей'].iloc[1:]
        # группируем значения меньше 1%
        sum_other = 0
        for i in range(len(pie_values)):
            if pie_values[i] < 0.03:
                sum_other += pie_values[i]
                del pie_names[i]
                del pie_values[i]
                del pie_labels[i]
        # Добавляем групповое значенние
        pie_values = list(pie_values)
        pie_values.append(sum_other)
        pie_names = list(pie_names)
        pie_names.append('Другие ТР менее 3%')
        pie_labels = list(pie_labels)
        pie_labels.append(0)
        # Добавляем приросты к названию рынка
        pie_names_and_labels = []
        for i in range(len(pie_names)):
            if pie_labels[i] > 0:
                pie_names_and_labels.append(pie_names[i] + ' (+' + str(round(pie_labels[i] * 100, 2)) + '%)')
            elif pie_labels[i] < 0:
                pie_names_and_labels.append(pie_names[i] + ' (' + str(round(pie_labels[i] * 100, 2)) + '%)')
            else:
                pie_names_and_labels.append(pie_names[i])
        pie_names = pie_names_and_labels
        # выводим график
        fig = px.pie(values=pie_values, names=pie_names, title='Доля по ТР на ' + str(now_date))
        fig.update_traces(textinfo='label+percent',
                          hovertemplate='<b>%{label}</b><br>%{value:.2f} тн (%{percent:.1%})',
                          text=pie_names)
        fig.update_layout(font=dict(size=18))  # увеличиваем шрифт
        fig.update_layout(height=700)
        return fig
    # Формируем комменты
    def text_ob():
        text_ob1 = 'Оборот отгрузок текущего месяца: ' + str(round(df_trp.iloc[0, 1] / 1000000, 2)) + ' млн. руб.'
        text_ob2 = 'Оборот отгрузок в прошлом году: ' + str(round(df_trp.iloc[0, 13] / 1000000, 2)) + ' млн. руб.'
        text_ob3 = 'Прошло рабочих дней: ' + str(now_working_days) + " из " + str(end_working_days) + ". Всего раб. дней в прошлом году: " + str(end_working_days_2)
        text_ob4 = 'Прогноз оборота: ' + str(round(list(df_trp['Прогноз выполнения'])[0] / 1000000, 2)) + ' млн. руб.'
        text_ob5 = 'Прогноз прироста по обороту (ср-дн): ' + str(round(list(df_trp['Прогноз прироста в руб.'])[0] / 1000000, 2)) + ' млн. руб., ' + str(list(df_trp['Прогноз прироста в %'])[0]) + '%'
        text_ob = html.P(children=text_ob1 + '\n' + text_ob2 + '\n' + text_ob3 + '\n' + text_ob4 + '\n' + text_ob5,
            className="header-description", style={'whiteSpace': 'pre-line'})
        return text_ob
    def text_vp():
        text_vp1 = 'Оборот ВП текущего месяца: ' + str(round(df_trp.iloc[0, 14] / 1000000, 2)) + ' млн. руб.'
        text_vp2 = 'Оборот ВП в прошлом году: ' + str(round(df_trp.iloc[0, 26] / 1000000, 2)) + ' млн. руб.'
        text_vp3 = 'Прогноз ВП: ' + str(round(list(df_trp['Прогноз выполнения ВП'])[0] / 1000000, 2)) + ' млн. руб.'
        text_vp4 = 'Прогноз прироста по ВП (ср-дн): ' + str(round(list(df_trp['Прогноз прироста ВП в руб.'])[0] / 1000000, 2)) + ' млн. руб., ' + str(list(df_trp['Прогноз прироста ВП в %'])[0]) + '%'
        text_vp = html.P(children=text_vp1 + '\n' + text_vp2 + '\n' + text_vp3 + '\n' + text_vp4,
            className="header-description", style={'whiteSpace': 'pre-line'})
        return text_vp
    def text_tr():
        df_tr = df_trp.copy().sort_values(by='Прогноз прироста в руб.')
        df_tr = df_tr[df_tr['Название ТС'] != 'Итого']
        text_tr1 = '1. ' + str(list(df_tr['Название ТС'])[-1]) + ': ' + str(round(list(df_tr['Прогноз прироста в руб.'])[-1] / 1000000, 2)) + ' млн. руб.'
        text_tr2 = '2. ' + str(list(df_tr['Название ТС'])[-2]) + ': ' + str(round(list(df_tr['Прогноз прироста в руб.'])[-2] / 1000000, 2)) + ' млн. руб.'
        text_tr3 = '3. ' + str(list(df_tr['Название ТС'])[-3]) + ': ' + str(round(list(df_tr['Прогноз прироста в руб.'])[-3] / 1000000, 2)) + ' млн. руб.'
        text_tr4 = '1. ' + str(list(df_tr['Название ТС'])[0]) + ' ' + str(round(list(df_tr['Прогноз прироста в руб.'])[0] / 1000000, 2)) + ' млн. руб.'
        text_tr5 = '2. ' + str(list(df_tr['Название ТС'])[1]) + ' ' + str(round(list(df_tr['Прогноз прироста в руб.'])[1] / 1000000, 2)) + ' млн. руб.'
        text_tr6 = '3. ' + str(list(df_tr['Название ТС'])[2]) + ' ' + str(round(list(df_tr['Прогноз прироста в руб.'])[2] / 1000000, 2)) + ' млн. руб.'
        text_tr = html.P(children='По приросту:\n' + text_tr1 + '\n' + text_tr2 + '\n' + text_tr3 + '\n\nПо оттоку: \n' + text_tr4 + '\n' + text_tr5 + '\n' + text_tr6,
                         className="header-description", style={'whiteSpace': 'pre-line'})
        return text_tr
    def text_market_share():
        df_market = df_trp.sort_values('Отклонение долей')
        text_negative = ''
        for i in range(3):
            if list(df_market['Отклонение долей'])[i] < (-0.01):
                text_negative += list(df_market['Название ТС'])[i] + ': ' + str(round(list(df_market['Отклонение долей'])[i] * 100, 2)) + '%\n'
        text_positive = ''
        for i in range(-1, -4, -1):
            if list(df_market['Отклонение долей'])[i] > 0.01:
                text_positive += list(df_market['Название ТС'])[i] + ': ' + str(round(list(df_market['Отклонение долей'])[i] * 100, 2)) + '%\n'

        text_market_share = html.P(children='Ключевые изменения долей рынка:\n' + text_positive + '\n' + text_negative,
            className="header-description", style={'whiteSpace': 'pre-line'})
        return text_market_share
    def text_recom():
        text_rec = ''
        # Отгрузка - снижение среднемесячного оборота
        mask = (df_trp['Доля на ' + df_trp.columns[1]] > 0.01) & (df_trp['Отклонение отгрузок по ССП(3 мес)'] < -100000) & ((df_trp['Отклонение отгрузок по ССП(3 мес)'].abs() / df_trp['Прогноз выполнения']) > 0.25)
        df_low_sales = df_trp[mask].sort_values('Отклонение отгрузок по ССП(3 мес)')
        if len(df_low_sales['Название ТС']) > 0:
            text_rec += '\nСигнал: отклонение прогноза отгрузки от ССП за 3 мес \n\n Мероприятия: \n'
            for i in range(len(df_low_sales['Название ТС'])):
                text_rec += 'Прогноз оттока по ' + str(list(df_low_sales['Название ТС'])[i]) + ' на сумму ' + str(round(list(df_low_sales['Отклонение отгрузок по ССП(3 мес)'])[i])) + '\n'
        # Снижение КТН
        if df_trp['Прирост КТН'].iloc[0] < -0.01:
            text_rec += '\nСигнал: снижение КТН на ' + str(df_trp['Прирост КТН'].iloc[0]) + '\n\nМероприятия:\n'
            mask = (df_trp.iloc[:, 27] > df_trp.iloc[0, 27]) & (df_trp['Доля на ' + df_trp.columns[1]] > 0.005) & (df_trp['Прогноз прироста в руб.'] < 0)
            df_ktn = df_trp[mask].sort_values('Прогноз прироста в руб.')
            if len(df_ktn['Название ТС']) > 0:
                text_rec += 'Проработать отток по ' + str(list(df_ktn['Название ТС'])[0]) + ' на сумму: ' + str(round(list(df_ktn['Прогноз прироста в руб.'])[0])) + ' руб.' + ', КТН ' + str(df_ktn.iloc[0, 39]) + '\n'
            mask = (df_trp.iloc[:, 27] < df_trp.iloc[0, 27]) & (df_trp['Доля на ' + df_trp.columns[1]] > 0.005) & (df_trp['Прогноз прироста в руб.'] > 0) & (df_trp['Прирост КТН'] < 0)
            df_ktn = df_trp[mask].sort_values('Прогноз прироста в руб.')
            if len(df_ktn['Название ТС']) > 0:
                text_rec += 'Поднять КТН по ' + str(list(df_ktn['Название ТС'])[-1]) + ' с ' + str(df_ktn.iloc[-1, 27]) + ' до уровня ' + str(df_ktn.iloc[-1, 39]) + '\n'
        # Угрозы прошлого года
        middle_sales = (df_trp['Прогноз выполнения'] + df_trp.iloc[:, 2] + df_trp.iloc[:, 3]) / 3
        mask = ((middle_sales * 1.5) < df_trp.iloc[:, 12]) & (df_trp['Доля на ' + df_trp.columns[1]] > 0.01)
        df_risk = df_trp[mask].sort_values(df_trp.columns[12])
        if len(df_risk['Название ТС']) > 0:
            text_rec += '\n' + 'Сигнал: выявлен аномальный оборот по ТР на ' + str(df_risk.columns[12]) + '\n\n' + 'Мероприятия:' + '\n'
            for i in range(len(df_risk['Название ТС'])):
                middle_risk = (list(df_risk['Прогноз выполнения'])[i] + df_risk.iloc[i, 2] + df_risk.iloc[i, 3]) / 3
                text_rec += 'Проработать оборот по ' + str(list(df_risk['Название ТС'])[i]) + ' на сумму ' + str(round(list(df_risk[df_risk.columns[12]])[i])) + ' руб. Прогноз оттока: ' + str(round(middle_risk - list(df_risk[df_risk.columns[12]])[i])) + '\n'
        # Отношение оборота к ВП

        text_rec = html.P(children=text_rec, className="header-description", style={'whiteSpace': 'pre-line'})
        return text_rec




    # Создаем Dash
    app = dash.Dash(__name__)
    app.layout = html.Div(style={'background-image': 'url("/assets/bg.jpg")',
           'background-repeat': 'repeat', 'background-position': 'right top',
           'background-size': '1920px 1152px'}, children=[
        # дашборд по листу ТРП
        html.H1('Анализ отчета \n"Помесячный оборот по отгрузке, КТН по клиентам и ТС"', className="header-title", style={'whiteSpace': 'pre-line', 'text-align': 'center'}),
        html.Br(),
        html.Br(),
        html.H1('Выполнение по обороту', className="header-title"),
        text_ob(),
        html.Br(),
        dcc.Graph(figure=fig_bar_trp_otgr()),
        html.Br(),
        html.H1('Выполнение по валовой прибыли', className="header-title"),
        text_vp(),
        html.Br(),
        dcc.Graph(figure=fig_bar_trp_vp()),
        html.Br(),
        html.H1('Прогноз прироста по ТР', className="header-title"),
        text_tr(),
        html.Br(),
        dcc.Graph(figure=fig_bar_trp_otgr_tr()),
        dcc.Graph(figure=fig_bar_trp_otgr_tr2()),
        html.Br(),
        html.H1('Доли отгрузки по ТР', className="header-title"),
        text_market_share(),
        html.Br(),
        dcc.Graph(id='graph-pie', figure=fig_pie_trp_otgr_tr_now()),
        html.Br(),
        html.H1('Рекомендации по развитию', className="header-title"),
        text_recom(),
        html.Br(),

        # График отгрузок по ТР
        html.H1('Динамика отгрузок по ТР', className="header-title"),
        dcc.Dropdown(id='ts-dropdown', options=options_df_trp, style={'font-size': '24px', 'font-weight': 'bold'},  # передаем список значений в Dropdown
            value=df_trp['Название ТС'].iloc[0]),  # задаем значение по умолчанию
        dcc.Graph(id='graph'),
        dcc.Graph(id='graph2'),
        dcc.Graph(id='graph3'),
    ])

    # Создаем callback-функцию
    @app.callback(
        dash.dependencies.Output('graph', 'figure'),
        [dash.dependencies.Input('ts-dropdown', 'value')]
    )
    def update_graph(selected_ts):
        # определяем оси
        x_plot_df_trp = df_trp.columns[13:0:-1]
        y_plot_df_trp = df_trp[df_trp['Название ТС'] == selected_ts].iloc[0, 13:0:-1].values
        y_plot_txt = list(map(lambda x: str(round((x / 1000000), 2)), y_plot_df_trp))

        fig = px.line(x=x_plot_df_trp, y=y_plot_df_trp, title=f'Динамика отгрузок по полю "{selected_ts}" в млн. руб.', text=y_plot_txt)
        fig.update_layout(font=dict(size=18))  # увеличиваем шрифт
        fig.update_layout(xaxis_title="Период отгрузки", yaxis_title="Сумма отгрузки")
        fig.update_traces(textposition='bottom center')    # позиция текста
        return fig

    @app.callback(
        dash.dependencies.Output('graph2', 'figure'),
        [dash.dependencies.Input('ts-dropdown', 'value')]
    )
    def update_graph(selected_ts):
        # определяем оси
        x_plot_df_trp = df_trp.columns[26:13:-1]
        y_plot_df_trp = df_trp[df_trp['Название ТС'] == selected_ts].iloc[0, 26:13:-1].values
        y_plot_txt = list(map(lambda x: str(round((x / 1000000), 2)), y_plot_df_trp))

        fig = px.line(x=x_plot_df_trp, y=y_plot_df_trp, title=f'Динамика ВП по полю "{selected_ts}" в млн. руб.',
                      text=y_plot_txt)
        fig.update_layout(font=dict(size=18))  # увеличиваем шрифт
        fig.update_layout(xaxis_title="Период отгрузки", yaxis_title="Сумма отгрузки")
        fig.update_traces(textposition='bottom center')  # позиция текста
        return fig

    @app.callback(
        dash.dependencies.Output('graph3', 'figure'),
        [dash.dependencies.Input('ts-dropdown', 'value')]
    )
    def update_graph(selected_ts):
        # определяем оси
        x_plot_df_trp = df_trp.columns[39:26:-1]
        y_plot_df_trp = df_trp[df_trp['Название ТС'] == selected_ts].iloc[0, 39:26:-1].values

        fig = px.line(x=x_plot_df_trp, y=y_plot_df_trp, title=f'Динамика КТН по полю "{selected_ts}"', text=y_plot_df_trp)
        fig.update_layout(font=dict(size=18))  # увеличиваем шрифт
        fig.update_layout(xaxis_title="Период отгрузки", yaxis_title="Значение КТН")
        fig.update_traces(textposition='bottom center')  # позиция текста
        return fig

    # Запускаем сервер
    if __name__ == '__main__':
        app.run_server(debug=True)



kpi_of_the_sales_group()