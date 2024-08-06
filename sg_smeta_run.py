# проба git
# проба git windows
# проба git debian

import PySimpleGUI as sg
import functions_and_classes as fc
import pathlib
import pickle
import docxtpl
import copy #####


DB_FILE = 'prices/DB.xls'
DB_FILE_SHEET_DEFAULT = "Default settings"

default = fc.FromXLS(DB_FILE, DB_FILE_SHEET_DEFAULT).default()

'''
default['dir_rec_files'] = 'record_files'
default['tpl_file'] = 'templates/smeta_tpl.docx'
default['pablic_gd_url_db_file'] = 'https://docs.google.com/spreadsheets/d/1KlRFMJgW-Dm53Qbjh9Dv6W067BhVeyVv/edit?usp=drivesdk&ouid=101603067907582645147&rtpof=true&sd=true'
default['db_file_sheet_contractor'] = "Contractor"
default['db_file_sheet_rab'] = "Jobs names"
default['db_file_sheet_mat'] = "Insulation materials"
default['db_file_sheet_about'] = "About"
default['sg_theme_style'] = 'Dark'
default['max_len_line'] = 34
default['ps_default'] = 'Аванс - 50%\nГарантия на работы - 1 год\nСрок выполнения - по-согласованию Сторон\nПодпись  ____________________'
default['distance_default'] = 10
default['delivery_default'] = 3
default['zp_default'] = 3500
default['kzp_default'] = 1.78
default['df_default'] = 1225
default['kdf_default'] = 1.5
default['sut_default']= 550
default['ksut_default'] = 2.11
default['proezd_default'] = 10
default['kproezd_default'] = 1.4
default['arenda_default'] = 51
default['karenda_default'] = 1.4
default['cargo_default'] = 100
default['kcargo_default'] = 1.2
default['prochie_default'] = 5
default['kprochie_default'] = 1.4
default['itr_default'] = 0
default['kitr_default'] = 1.5
default['proiz_default'] = 3
default['exit_yes_1'] = 'Уверены,\nчто хотите выйти\nиз программы?'
default['exit_yes_2'] = 'Сохранили\nвведенные данные?'
default['vsego_st_smeta_error'] = "\nСорян, чувак,\nНе определен ни один раздел!!!\n"
default['export_txt_ok'] = "\nЭкспорт сметы в ТХТ-формат\nуспешно завершен!\n"
default['export_txt_error'] = "\nЭкспорт сметы в ТХТ-формат!\n\nНЕУДАЧНО!\n\n Нажмите сначала •Расчитать смету•,\n Затем попробуйте снова\n"
default['export_docx_ok'] = "\nЭкспорт сметы в DOCX-формат\nуспешно завершен!\n"
default['export_docx_error'] = "\nЭкспорт сметы в DOCX-формат!\n\nНЕУДАЧНО!\n\n Нажмите сначала •Расчитать смету•,\n Затем попробуйте снова\n"
default['db_update_ok'] = "\nОбновление базы\nуспешно загружено\nс удаленного сервера!\nДля начала работы\nтребуется перезапуск\nпрограммы.\n"
default['db_update_error'] = "\nСорян, чувак!\nЧто-то пошло не так!\n\nИли интернета нет,\n или удаленная база\nне найдена!!!\n"
'''

sg.theme(default['sg_theme_style'])
cnfg_1 = dict(font='Courier 12', size=(15, 2), text_color='yellow')
cnfg_2 = dict(font='Courier 12', size=(15, 2), text_color='grey')

dict_contractor = fc.FromXLS(DB_FILE, default['db_file_sheet_contractor']).dict
dict_rab = fc.FromXLS(DB_FILE, default['db_file_sheet_rab']).dict
dict_mat = fc.FromXLS(DB_FILE, default['db_file_sheet_mat']).dict
dict_about_programm = fc.FromXLS(DB_FILE, default['db_file_sheet_about']).dict

contractor_val = [value[0] for value in dict_contractor.values()]
r_combo = [value[0] for value in dict_rab.values()]
#m_listbox = [f'-{value[0]} = {round(value[2], 2)}р/м2' for value in dict_mat.values()]
m_listbox = dict_mat.values()
about_programm = [value[0] for value in dict_about_programm.values()]

if not pathlib.Path.exists(pathlib.Path.cwd()/default['dir_rec_files']):
    pathlib.Path.mkdir(pathlib.Path.cwd()/default['dir_rec_files'])

menu_def = [
    ['&Файл', ['&Открыть', '&Сохранить как', '&Выход']],
    ['&Обновление', ['&Обновить базу', ], ],
    ['&Экспорт', ['&Экспортировать в ТХТ', '&Экспорт в DOCX']],
    ['&Помощь', ['&О программе']],
    ]
right_click_menu_def = [
    [],
    ['&Помощь', '&Выход'],
    ]

common_layout = [
    [sg.Multiline('Общие данные сметы', font='Courier 12', text_color='darkred', background_color='lightblue', justification='center', expand_x=True, disabled=True, no_scrollbar=True, size=(0, 1)), ],
    [sg.Text('Файл: ', **cnfg_1, ), sg.Multiline('', font='Courier 10', size=(0, 3), expand_x=True, disabled=True, no_scrollbar=True, key='-FILE-', ), ],
    [sg.Text('Дата: ', **cnfg_1, ),
        sg.Input(font='Courier 10', readonly=False, expand_x=True, size=(22, 1), key='-DATE-', ),
        sg.CalendarButton('Календарь', title='Выберите дату', target='-DATE-', font='Courier 10', size=(20, 1), format='%d-%m-20%y', ), ],
    [sg.Text('Подрядчик: ', **cnfg_1,), sg.Combo(values=contractor_val, default_value='', font='Courier 10', readonly=False, expand_x=True, key='-CONTRACTOR-', ), ],
    [sg.Text('Заказчик: ', **cnfg_1, ), sg.Multiline(font='Courier 10', size=(0, 3), expand_x=True, key='-CLIENT-', metadata='Вводи'), ],
    [sg.Text('Стройка: ', **cnfg_1, ), sg.Multiline(font='Courier 10', size=(0, 4), expand_x=True, key='-CONSTRUCTION-', ), ],
    [sg.Text('Объект: ', **cnfg_1, ), sg.Multiline(font='Courier 10', size=(0, 4), expand_x=True, key='-OBJECT-', ), ],
    [sg.Text('Условия: ', **cnfg_1, ), sg.Multiline(default['ps_default'], font='Courier 10', size=(0, 6), expand_x=True, key='-PS-', ), ],
    ]

logistic_layout = [
    [sg.Multiline('Расстояние до объекта, количество доставок оборудования и материалов', font='Courier 12', text_color='darkred', background_color='lightblue', justification='center', expand_x=True, disabled=True, no_scrollbar=True, size=(0, 1)), ],
    [sg.Text('Расстояние,\nкм: ', **cnfg_1,), sg.Slider(range=(0, 10000), default_value=default['distance_default'], font='Courier 10', resolution=10, tick_interval=2000, orientation='h', expand_x=True, key='-DISTANCE-'), ],
    [sg.Text('Доставка,\nшт: ', **cnfg_1, ), sg.Slider(range=(0, 20), default_value=default['delivery_default'], font='Courier 10', resolution=1, tick_interval=5, orientation='h', expand_x=True, key='-DELIVERY-'), ],
    ]

ot_layout = [
    [sg.Multiline('Ставка оплаты труда рабочего "на руки" и чистый доход фирмы с одного рабочего', font='Courier 12', text_color='darkred', background_color='lightblue', justification='center', expand_x=True, disabled=True, no_scrollbar=True, size=(0, 1)), ],
    [sg.StatusBar('Оплата труда рабочего', font='Courier 10', text_color='darkred', background_color='grey', justification='center', pad=(5, 20), visible=True,)],
    [sg.Text('З/П,\nр/ч.д.: ', **cnfg_1,), sg.Slider(range=(0, 15000), default_value=default['zp_default'], font='Courier 10', resolution=50, tick_interval=5000, orientation='h', expand_x=True, key='-ZP-'), ],
    [sg.Text('К-т к З/П: ', **cnfg_2,), sg.Slider(range=(1, 3), default_value=default['kzp_default'], font='Courier 8', text_color='grey', resolution=0.01, tick_interval=1, orientation='h', expand_x=True, key='-KZP-'), ],
    [sg.StatusBar('Доход фирмы', font='Courier 10', text_color='darkred', background_color='grey', justification='center', pad=(5, 20), visible=True,)],
    [sg.Text('Доход,\nр/ч.д.: ', **cnfg_1,), sg.Slider(range=(0, 15000), default_value=default['df_default'], font='Courier 10', resolution=5, tick_interval=5000, orientation='h', expand_x=True, key='-DF-'), ],
    [sg.Text('К-т к Доходу: ', **cnfg_2,), sg.Slider(range=(1, 3), default_value=default['kdf_default'], font='Courier 8', text_color='grey', resolution=0.01, tick_interval=1, orientation='h', expand_x=True, key='-KDF-'), ],
    ]

mob_layout = [
    [sg.Multiline('Тарифы (без учёта налогов и отчислений)', font='Courier 12', text_color='darkred', background_color='lightblue', justification='center', expand_x=True, disabled=True, no_scrollbar=True, size=(0, 1)), ],
    [sg.StatusBar('Суточные, доставка персонала, аренда жилья, доставка и возврат материалов и оборудования', font='Courier 10', text_color='darkred', background_color='grey', justification='center', pad=(5, 20), visible=True,)],
    [sg.Text('Суточные,\nр/ч.д.: ', **cnfg_1, ), sg.Slider(range=(0, 4500), default_value=default['sut_default'], font='Courier 10', resolution=1, tick_interval=1500, orientation='h', expand_x=True, key='-SUT-'), ],
    [sg.Text('К-т к Суточным: ', **cnfg_2,), sg.Slider(range=(1, 3), default_value=default['ksut_default'], font='Courier 8', text_color='grey', resolution=0.01, tick_interval=1, orientation='h', expand_x=True, key='-KSUT-'), ],
    [sg.Text('Проезд,\nр/км: ', **cnfg_1, ), sg.Slider(range=(0, 300), default_value=default['proezd_default'], font='Courier 10', resolution=1, tick_interval=100, orientation='h', expand_x=True, key='-PROEZD-'), ],
    [sg.Text('К-т к Проезду: ', **cnfg_2,), sg.Slider(range=(1, 3), default_value=default['kproezd_default'], font='Courier 8', text_color='grey', resolution=0.01, tick_interval=1, orientation='h', expand_x=True, key='-KPROEZD-'), ],
    [sg.Text('Жилье на 3чел,\nтыс.р/месяц: ', **cnfg_1, ), sg.Slider(range=(0, 300), default_value=default['arenda_default'], font='Courier 10', resolution=1, tick_interval=100, orientation='h', expand_x=True, key='-ARENDA-'), ],
    [sg.Text('К-т к Жилью: ', **cnfg_2,), sg.Slider(range=(1, 3), default_value=default['karenda_default'], font='Courier 8', text_color='grey', resolution=0.01, tick_interval=1, orientation='h', expand_x=True, key='-KARENDA-'), ],
    [sg.Text('Грузоперевозка,\nр/км: ', **cnfg_1,), sg.Slider(range=(0, 300), default_value=default['cargo_default'], font='Courier 10', resolution=1, tick_interval=100, orientation='h', expand_x=True, key='-CARGO-'), ],
    [sg.Text('К-т к\nГрузоперевозке: ', **cnfg_2,), sg.Slider(range=(1, 3), default_value=default['kcargo_default'], font='Courier 8', text_color='grey', resolution=0.01, tick_interval=1, orientation='h', expand_x=True, key='-KCARGO-'), ],
    ]

tarif_layout = [[sg.TabGroup([[
    sg.Tab(' Оплата труда \n', ot_layout),
    sg.Tab('\n Мобилизация ', mob_layout),
    ]],
    key='-TAB GROUP_3-',
    font='Courier 12',
    tab_background_color='grey', selected_background_color='pink',
    selected_title_color='blue',
    pad=(5, 10),
    tab_location='top', expand_x=True, )
    ]]

other_layout = [
    [sg.Multiline('Дополнительные расходы на объекте "на руки"', font='Courier 12', text_color='darkred', background_color='lightblue', justification='center', expand_x=True, disabled=True, no_scrollbar=True, size=(0, 1)), ],
    [sg.StatusBar('Прочие расходы, вагончик, ... (всего)', font='Courier 10', text_color='darkred', background_color='grey', justification='center', pad=(5, 20), visible=True,)],
    [sg.Text('Прочие,\nтыс.р: ', **cnfg_1,), sg.Slider(range=(0, 1500), default_value=default['prochie_default'], font='Courier 10', resolution=1, tick_interval=500, orientation='h', expand_x=True, key='-PROCHIE-'), ],
    [sg.Text('К-т к Прочим: ', **cnfg_2,), sg.Slider(range=(1, 3), default_value=default['kprochie_default'], font='Courier 8', text_color='grey', resolution=0.01, tick_interval=1, orientation='h', expand_x=True, key='-KPROCHIE-'), ],
    [sg.StatusBar('ИТР - З/П, жилье, проезд, ... (всего)', font='Courier 10', text_color='darkred', background_color='grey', justification='center', pad=(5, 20), visible=True,)],
    [sg.Text('ИТР,\nтыс.р: ', **cnfg_1, ), sg.Slider(range=(0, 1500), default_value=default['itr_default'], font='Courier 10', resolution=1, tick_interval=500, orientation='h', expand_x=True, key='-ITR-'), ],
    [sg.Text('К-т к ИТР: ', **cnfg_2,), sg.Slider(range=(1, 3), default_value=default['kitr_default'], font='Courier 8', text_color='grey', resolution=0.01, tick_interval=1, orientation='h', expand_x=True, key='-KITR-'), ],
    ]


def chapter_layout(num):
    chapter_layout = [
        [sg.Text(f'РАЗДЕЛ №{num}', font='Courier 12', text_color='darkred', background_color='lightblue', justification='center', expand_x=True, ), ],
        [sg.Text('Описание раздела: ', font='Courier 12', text_color='yellow', justification='center', expand_x=True), ],
        [sg.Multiline(font='Courier 10', size=(60, 3), expand_x=True, key=f'-OPPOS{num}-'), ],
        [sg.Text('Наименование работы: ', font='Courier 12', text_color='yellow', justification='center', expand_x=True, ), ],
        [sg.Combo(values=r_combo, default_value='', font='Courier 10', readonly=False, expand_x=True, k=f'-JOB{num}-'), ],
        [sg.Text('Колич-во работ,\nм2: ', **cnfg_1,), sg.Slider(range=(0, 4500), default_value=0, font='Courier 10', resolution=1, tick_interval=1500, orientation='h', expand_x=True, key=f'-V{num}-'), ],
        [sg.Text('Производит-ть,\nм2/ч.д.: ', **cnfg_1,), sg.Slider(range=(0, 45), default_value=default['proiz_default'], font='Courier 10', resolution=0.1, tick_interval=15, orientation='h', expand_x=True, key=f'-P{num}-'), ],
        [sg.Text('Наименование материалов: ', font='Courier 12', text_color='yellow', justification='center', expand_x=True, ), ],
        [sg.Listbox(values=m_listbox, font='Courier 10', select_mode='multiple', size=(20, 10), expand_x=True, key=f'-MAT{num}-')],
        ]
    return chapter_layout


chapters = [sg.Tab(f'Раздел {i}\n', chapter_layout(i), visible=True, ) for i in range(1, 6)]

chapters_layout = [[sg.TabGroup(
    [chapters],
    key='-TAB GROUP_2-',
    font='Courier 12',
    tab_background_color='grey', selected_background_color='pink',
    selected_title_color='blue',
    pad=(5, 10),
    tab_location='top', expand_x=True,
    )]]

layout = [
    [sg.Menu(menu_def, font='Courier 12', key='-MENU-')],
    [sg.TabGroup([
        [
            sg.Tab(' ОБЩЕЕ \n', common_layout),
            sg.Tab('\n ЛОГИСТИКА ', logistic_layout),
            sg.Tab(' ТАРИФЫ \n', tarif_layout),
            sg.Tab('\n ДОПРАСХОДЫ ', other_layout),
            sg.Tab(' РАЗДЕЛЫ \n', chapters_layout), ]],
        key='-TAB GROUP_1-',
        font='Courier 12',
        tab_background_color='grey',
        selected_background_color='darkorange',
        selected_title_color='blue',
        pad=(5, 10),
        tab_location='top',
        expand_x=True,
        )],
    [sg.HorizontalSeparator()],
    [sg.Text('ИТОГИ', font='Courier 10', text_color='darkred', background_color='lightblue', justification='center', expand_x=True), ],
    [sg.Multiline(size=(60, 6), font='Courier 10', expand_x=True, disabled=True, key='-ITOG-', visible=False, text_color='darkred', background_color='lightblue', border_width=3), ],
    [sg.Text('Скидка на мат-лы, %: ', **cnfg_1,), sg.Slider(range=(0, 50), default_value=0, font='Courier 10', resolution=1, tick_interval=5, orientation='h', expand_x=True, key='-DISCOUNT_MAT-', text_color='red', background_color='lightgrey'), ],
    [sg.Button('Расчитать смету', expand_x=True, expand_y=True, button_color='green', size=(20, 1), font='Courier 20')],
    [sg.StatusBar('Powered by Vitaly Arkhipov', enable_events=False, font='Courier 10', text_color='darkred', background_color='grey', justification='center', key=None, expand_x=False, expand_y=False, visible=True)],
    [sg.Sizegrip(background_color='darkred', key=None)],
    ]


window = sg.Window(
    'Smeta',
    layout,
    right_click_menu=right_click_menu_def,
    enable_close_attempted_event=True,
    size=(1180, 950),
    resizable=True,
    location=(0, 0),
    finalize=True,
    # auto_size_buttons=False,
    )
# window.Maximize()


while True:

    event, values = window.read(timeout=0)

    smf_file_data = copy.deepcopy(values) #####

    for num in range(1, 6):
       smf_file_data[f'-MAT{num}-'] = window[f'-MAT{num}-'].get_indexes()

    if event == sg.WIN_CLOSE_ATTEMPTED_EVENT or event == 'Exit' or event == 'Выход':
        if sg.popup_yes_no(default['exit_yes_1'], title='Завершение работы', modal=True) == 'Yes':
            if sg.popup_yes_no(default['exit_yes_2'], title='Завершение работы', modal=True) == 'Yes':
                break

    elif event == 'Расчитать смету':
        filename = smf_file_data['-FILE-']
        date = smf_file_data['-DATE-']
        isp = smf_file_data['-CONTRACTOR-']
        zak = smf_file_data['-CLIENT-']
        strojka = smf_file_data['-CONSTRUCTION-']
        object = smf_file_data['-OBJECT-']
        ps = smf_file_data['-PS-']
        zp = smf_file_data['-ZP-']
        kzp = smf_file_data['-KZP-']
        df = smf_file_data['-DF-']
        kdf = smf_file_data['-KDF-']
        sut = smf_file_data['-SUT-']
        c_arenda = smf_file_data['-ARENDA-']
        distance = smf_file_data['-DISTANCE-']
        c_leg = smf_file_data['-PROEZD-']
        ksut = smf_file_data['-KSUT-']
        kproezd = smf_file_data['-KPROEZD-']
        karenda = smf_file_data['-KARENDA-']
        c_gruz = smf_file_data['-CARGO-']
        kgruz = smf_file_data['-KCARGO-']
        kol_rejs_gruz = smf_file_data['-DELIVERY-']
        prochie = smf_file_data['-PROCHIE-']
        kprochie = smf_file_data['-KPROCHIE-']
        itr = smf_file_data['-ITR-']
        kitr = smf_file_data['-KITR-']
        discount_mat = (100-smf_file_data['-DISCOUNT_MAT-'])/100
        pl = []
        number = 0
        fc.Rabota.vsego_v = 0
        fc.Rabota.vsego_st = 0
        fc.Rabota.vsego_zp = 0
        fc.Rabota.vsego_df = 0
        fc.Material.vsego_st = 0
        fc.Trudoemkost.vsego_chd = 0
        fc.Position.vsego_st = 0
        for num in range(1, 6):
            oppos = smf_file_data[f'-OPPOS{num}-']
            name_rab = smf_file_data[f'-JOB{num}-']
            v_rab = smf_file_data[f'-V{num}-']
            proiz = smf_file_data[f'-P{num}-']
            selected_indexes_mat = smf_file_data[f'-MAT{num}-']
            if name_rab and v_rab:
                number += 1
                rabota = fc.Rabota(dict_rab, name_rab, zp, df, v_rab, proiz, default['max_len_line'], kzp, kdf)
                material = fc.Material(dict_mat, v_rab, selected_indexes_mat, default['max_len_line'], discount_mat)
                trudoemkost = fc.Trudoemkost(v_rab, proiz)
                mobilization = fc.Mobilization(c_arenda, distance, c_leg, c_gruz, sut, kol_rejs_gruz, prochie, itr, trudoemkost, ksut, kproezd, karenda, kgruz, kprochie, kitr)
                position = fc.Position(number, oppos, rabota, material)
                pl.append(position)
        try:
            vsego_st_smeta = (position.vsego_st + mobilization.vsego_st) / 1.2
            vsego_st_smeta_nds = vsego_st_smeta * 0.2
            vsego_smeta = vsego_st_smeta + vsego_st_smeta_nds
            smeta_txt, smeta_dict = fc.get_smeta(smf_file_data, pl, rabota, material, mobilization, trudoemkost, vsego_st_smeta, vsego_st_smeta_nds, vsego_smeta)
            smeta_txt = smeta_txt + '\n\n' + str(selected_indexes_mat) + str(smf_file_data) + '\n\n' + str(values) ###########
        except Exception:
            sg.popup_error(default['vsego_st_smeta_error'], title="Ошибка ввода разделов")
            smeta_txt = 'Скорректируй ввод разделов!'
            smeta_txt = smeta_txt + '\n\n' + str(dict_mat) + str(smf_file_data) + '\n\n' + str(values) ###########
        window['-ITOG-'].Update(smeta_txt)
        ##################################
        layout_out = [[sg.Multiline(smeta_txt, size=(85, 50), font='Courier 10', expand_x=True, disabled=True, key='-OUT-', visible=True, text_color='darkred', background_color='lightblue', border_width=1)]]
        window1 = sg.Window('Форма вывода сметы', layout_out, location=(30, 30), modal=True, finalize=True)
        # window1.Maximize()

    elif event == 'Открыть':
        open_file = sg.popup_get_file(message='Открыть', title='Открытие из файла настроек сметы', default_path='', default_extension='', save_as=False, multiple_files=False, file_types=(('SGF', '*.sgf'),), no_window=True, size=(35, 20), button_color=None, background_color=None, text_color=None, icon=None, font='Courier 6', no_titlebar=False, grab_anywhere=False, keep_on_top=False, location=(None, None), initial_folder=pathlib.Path.cwd()/default['dir_rec_files'], image=None, files_delimiter=';', modal=True)
        if open_file:
            with open(open_file, 'rb') as f:
                smf_file_data = pickle.load(f)
        smf_file_data.pop('-MENU-', )   ##########
        smf_file_data.pop('Календарь', )    ###########
        sg.fill_form_with_values(window, smf_file_data)
        for num in range(1, 6):
            window[f'-MAT{num}-'].Update(values=m_listbox, set_to_index=smf_file_data[f'-MAT{num}-'])
        window['-ITOG-'].Update('')

    elif event == 'Сохранить как':
        save_file = sg.popup_get_file('', save_as=True, file_types=(('SGF', '*.sgf'),), no_window=True, initial_folder=str(pathlib.Path.cwd()/default['dir_rec_files']), )
        smf_file_data['-FILE-'] = save_file
        window['-FILE-'](save_file)
        if save_file:
            with open(save_file, 'wb') as f:
                pickle.dump(smf_file_data, f)

    elif event == 'Экспортировать в ТХТ':
        try:
            fc.export_txt(filename, default['dir_rec_files'], content=smeta_txt, )
            sg.popup(default['export_txt_ok'], title='Результат экспорта сметы в ТХТ-формат')
        except Exception:
            sg.popup(default['export_txt_error'], title='Результат экспорта сметы в ТХТ-формат')

    elif event == 'Экспорт в DOCX':
        try:
            fc.export_docxtpl(filename, default['dir_rec_files'], docxtpl, pathlib, default['tpl_file'], content=smeta_dict)
            sg.popup(default['export_docx_ok'], title='Подтверждение экспорта сметы в DOCX-формат')
        except Exception:
            sg.popup(default['export_docx_error'], title='Ошибка экспорта сметы в DOCX-формат')

    elif event == 'О программе':
        sg.popup_scrolled(*about_programm, title='Сведения о программе', size=(30, 10), font=None, modal=True)

    elif event == 'Обновить базу':
        try:
            fc.db_update(default['pablic_gd_url_db_file'], DB_FILE)
            sg.popup(default['db_update_ok'], title='Результат обновления базы')
            break
        except Exception:
            sg.popup_error(default['db_update_error'], title="Ошибка обновления базы")
            break

window.close()
