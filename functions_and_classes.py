#!/usr/bin/python3
# -*- coding: utf-8 -*-

import prettytable
import pathlib
import xlrd
import urllib.request


class Cell:
    """
    Деление передаваемой строки на подстроки
    """

    def __init__(self, line, max_len_line):
        self.line = line
        self.max_len_line = max_len_line
        self.ls_line = list(self.line)

    def sep_line(self):
        """
        Передаваемая строка делится на подстроки (посимвольно) длинами не более max_len_line символов
        """

        n = 0
        while n < len(self.ls_line) - self.max_len_line:  # - len , чтобы в конце не добавлялась пустая строка
            n += self.max_len_line + 1
            self.ls_line.insert(n - 1, '\n')
        return ''.join(self.ls_line)

    def wrap(self):
        '''Передаваемая строка делится на подстроки (пословно) длинами не более max_len_line символов, которое как минимум на 1 больше длины самого длинного слова в строке
'''
        split_line = self.line.split(' ')
        c = 0
        a = 0
        wrap_line = []
        w = []
        while a <= len(split_line) - 1:
            w.append(split_line[a])
            c += len(split_line[a])
            if c <= self.max_len_line - len(w):
                wrap_line.append(split_line[a])
                a += 1
            else:
                wrap_line.append('\n')
                c = 0
                w = []
        return ' '.join(wrap_line)


class Table:
    '''
	Шапка таблицы (длина th = [] должна быть кратна длине td = []) ==> Table(th, td)
	'''

    def __init__(self, th, td):
        self.th = th
        self.td = td

    def __repr__(self):
        columns = len(self.th)
        table = prettytable.PrettyTable(self.th)
        # table.align = 'c' # l, r, c или table.align[self.th[2]] = 'r' ...
        table.align[self.th[0]] = 'l'
        # table.vertical_char = '|'
        table.horizontal_char = '—'
        table.junction_char = '+'
        table.padding_width = 1
        # table.hrules = prettytable.ALL
        while self.td:
            table.add_row(self.td[:columns])
            self.td = self.td[columns:]
        return str(table)


class Rabota:
    '''
	Расчет работ
	'''

    vsego_v = 0
    vsego_st = 0
    vsego_zp = 0
    vsego_df = 0

    def __init__(self, dict_rab, name_rab, zp, df, v_rab, proiz, max_len_line, kzp=1.83, kdf=1.5):
        self.dict_rab = dict_rab
        self.name_rab = name_rab
        self.zp = zp
        self.df = df
        self.v_rab = v_rab
        self.proiz = proiz
        self.max_len_line = max_len_line
        self.kzp = kzp
        self.kdf = kdf

        Rabota.vsego_v += self.v_rab
        Rabota.vsego_st += self.get_itogo_st()
        Rabota.vsego_zp += self.get_itogo_zp()
        Rabota.vsego_df += self.get_itogo_df()

    def get_name_rab(self):
        return Cell(self.name_rab, self.max_len_line).wrap()

    def get_crm2(self):
        return round(((self.zp * self.kzp + self.df * self.kdf) / self.proiz), 2)

    def get_itogo_st(self):
        return round((self.get_crm2() * self.v_rab), 2)

    def get_td_pos_rab(self):
        return [self.get_name_rab(), self.v_rab, self.get_crm2(), self.get_itogo_st()]

    def get_itogo_zp(self):
        return round((self.zp * self.v_rab / self.proiz), 2)

    def get_itogo_df(self):
        return round((self.df * self.v_rab / self.proiz), 2)

    def get_table_pos_rab(self):
        return Table(th=["........Наименование работ........", "Кол-во, м2", "Цена, р", "Стоимость, р"],
                     td=self.get_td_pos_rab())

    def __repr__(self):
        return f'{self.get_table_pos_rab()}\nИтого стоимость работ по этому разделу......................: {self.get_itogo_st()} р'


class Material:
    '''
	Расчет материалов
	'''
    vsego_pos = 0
    vsego_st = 0

    def __init__(self, dict_mat, v_rab, selected_indexes_mat, max_len_line, discount_mat):
        self.dict_mat = dict_mat
        self.v_rab = v_rab
        self.selected_indexes_mat = selected_indexes_mat
        self.discount_mat = discount_mat
        self.max_len_line = max_len_line  # namemat_ent.curselection() --> (0, 1, 2...) предст-ет кортеж номеров выбранных в Listbox позиций

        Material.vsego_pos += 1
        Material.vsego_st += self.get_itogo_st()

    def win_accept_mat(self):
        td_pos_mat = []
        itogo_st = 0
        pos_mat = []  # Выбранные позиции материалов
        name = []
        cena = []
        for pos in self.selected_indexes_mat:
            selected_name_krash_cena = self.dict_mat[pos + 1]
            pos_mat.append(selected_name_krash_cena)
            name.append(selected_name_krash_cena[0])
            cena.append(selected_name_krash_cena[2] * self.discount_mat)
            kol_mat = round(self.v_rab * selected_name_krash_cena[1], 2)
            st_mat = round(kol_mat * selected_name_krash_cena[2] * self.discount_mat, 2)
            name_kolmat_cena_stmat = [Cell(selected_name_krash_cena[0], self.max_len_line).wrap(), kol_mat,
                                      round(selected_name_krash_cena[2] * self.discount_mat, 2), st_mat]
            td_pos_mat += name_kolmat_cena_stmat
            itogo_st += st_mat
        name_change_mat = '\n• '.join(name)
        return td_pos_mat, itogo_st, name_change_mat

    def get_td_pos_mat(self):
        return self.win_accept_mat()[0]

    def get_itogo_st(self):
        return round(self.win_accept_mat()[1], 2)

    def get_table_pos_mat(self):
        return Table(th=['.....Наименование  материалов.....', 'Кол-во, м2', 'Цена, р', 'Стоимость, р'],
                     td=self.get_td_pos_mat())

    def __repr__(self):
        return f'{self.get_table_pos_mat()}\nИтого стоимость материалов по этому разделу.................: {self.get_itogo_st()} р'


class Trudoemkost:
    '''
	Определяет трудоемкость работ
	'''
    vsego_chd = 0

    def __init__(self, v_rab, proiz):
        self.v_rab = v_rab
        self.proiz = proiz
        Trudoemkost.vsego_chd += self.get_itogo_chd()

    def get_itogo_chd(self):
        return round((self.v_rab / self.proiz), 2)

    def __repr__(self):
        return f'Трудоемкость по этому разделу: {self.get_itogo_chd()} чел.дней\n{Trudoemkost.vsego_chd}'


class Position:
    '''
	Позиция
	'''
    vsego_st = 0

    def __init__(self, num, oppos, rabota, material):
        self.num = num
        self.oppos = oppos
        self.rabota = rabota
        self.material = material

        Position.vsego_st += self.get_itogo_st()

    def get_itogo_st(self):
        return round((self.rabota.get_itogo_st() + self.material.get_itogo_st()), 2)

    def __repr__(self):
        return f'Раздел № {self.num}: {self.oppos}\n{self.rabota}\n{self.material}\n'


class Mobilization:
    '''
	Определяет мобилизационные расходы
	'''
    vsego_st = 0

    def __init__(self, c_arenda, distance, c_leg, c_gruz, sut, kol_rejs_gruz, prochie, itr, trudoemkost, ksut=2.11, kproezd=1.4, karenda=1.4, kgruz=1.2, kprochie=1.4, kitr=1.4):
        self.ksut = ksut
        self.kproezd = kproezd
        self.karenda = karenda
        self.kgruz = kgruz
        self.kprochie = kprochie
        self.kitr = kitr
        self.c_arenda = c_arenda * 1000 * self.karenda
        #self.c_arenda = c_arenda * self.karenda
        self.distance = distance
        self.c_leg = c_leg * self.kproezd
        self.c_gruz = c_gruz * self.kgruz
        self.sut = sut * self.ksut
        self.kol_rejs_gruz = kol_rejs_gruz
        self.prochie = prochie * 1000 * self.kprochie
        self.itr = itr * 1000 * self.kitr
        self.trudoemkost = trudoemkost

        self.komand = self.trudoemkost.vsego_chd / 24 * 31 * self.sut

        self.kol_proezd = self.trudoemkost.vsego_chd / 24
        if self.kol_proezd < 5:
            self.kol_proezd = 5
        self.proezd = 2 * self.distance * self.c_leg * self.kol_proezd

        self.kol_mes_arenda = int((self.trudoemkost.vsego_chd / 3 / 24) + 0.85)
        if self.kol_mes_arenda == 0:
            self.kol_mes_arenda = 0.5
        self.arenda = self.kol_mes_arenda * self.c_arenda
        
        #self.kol_day_arenda = int((self.trudoemkost.vsego_chd / 24 * 31))
        #if self.kol_mes_arenda == 0:
            #self.kol_mes_arenda = 0.5
        #self.arenda = self.trudoemkost.vsego_chd / 24 * 31 * self.c_arenda

        self.dostavka = 2 * self.distance * self.c_gruz * self.kol_rejs_gruz

        if self.trudoemkost.vsego_chd == 0:
            self.komand = 0
            self.proezd = 0
            self.arenda = 0
            self.dostavka = 0
            self.prochie = 0
            self.itr = 0

        if self.distance < 25:
            Mobilization.vsego_st = self.prochie + self.itr
            self.komand = 0
            self.proezd = 0
            self.arenda = 0
            self.dostavka = 0
        else:
            Mobilization.vsego_st = self.komand + self.proezd + self.arenda + self.dostavka + self.prochie + self.itr

    def __repr__(self):
        return 'Мобилизация всего: {} р, в том числе:\n\t-Командировочные: {} р\n\t-Проезд: {} р\n\t-Аренда жилья: {} р\n\t-Доставка материалов: {}р\n\t-Прочие расходы: {} р\n\t-Расходы на ИТР: {} р'.format(
            round(Mobilization.vsego_st, 2), round(self.komand, 2), round(self.proezd, 2), round(self.arenda, 2),
            round(self.dostavka, 2), round(self.prochie, 2), round(self.itr, 2))


class FromXLS:

    def __init__(self, file_excel, sheet_file):
        self.file_excel = file_excel
        self.sheet_file = sheet_file
        self.read_db = xlrd.open_workbook(self.file_excel)  # a,formatting_info=True) # Открывает файл Excel
        self.sheet = self.read_db.sheet_by_name(self.sheet_file)  # Открывает нужный лист
        self.dict = {}
        self.number_dict = 1

        for rownum in range(1, self.sheet.nrows):
            self.row = self.sheet.row_values(rownum)
            if self.row[0]:
                self.dict.update({self.number_dict: self.row[
                                                    0:3]})  # заносит в словарь только 3 первых столбца, при условии, что первый столбец в этом ряду не пустой
                self.number_dict += 1

    def default(self):
        return {value[0]: value[1] for value in self.dict.values()}


##########

def db_update(url, file):
    file_id = url.split('/')[-2]
    dwn_url = 'https://drive.google.com/uc?export=download&id=' + file_id
    urllib.request.urlretrieve(dwn_url,
                               file)  # Сохраняет файл с Гугл диска в текущей папке с программой под именем 'DB.xls'


def export_txt(filename, dir_rec_files, content):
    name_file_txt = filename.split('/')[-1].split('.')[-2] + '.txt'
    sourcefile = pathlib.Path.open((pathlib.Path.cwd() / dir_rec_files / name_file_txt), 'w', encoding='utf-8')
    sourcefile.write(content)
    sourcefile.close()


######
def export_docxtpl(filename, dir_rec_files, docxtpl, pathlib, tpl_file, content):
    '''
	Сохраняет данные в .docx шаблон
	'''
    doc = docxtpl.DocxTemplate(tpl_file)
    # подставляем контент в шаблон
    doc.render(content)
    # сохраняем в файл *.docx в папку /dir_rec_files
    name_file_docx = filename.split('/')[-1].split('.')[-2] + '.docx'
    sourcefile = pathlib.Path.cwd() / dir_rec_files / name_file_docx
    doc.save(sourcefile)


######
def get_smeta(smf_file_data, pl, rabota, material, mobilization, trudoemkost, vsego_st_smeta, vsego_st_smeta_nds,
              vsego_smeta):
    filename = smf_file_data['-FILE-'].split('/')[-1]
    isp = smf_file_data['-CONTRACTOR-']
    zak = smf_file_data['-CLIENT-']
    date = smf_file_data['-DATE-']
    strojka = smf_file_data['-CONSTRUCTION-']
    object = smf_file_data['-OBJECT-']
    pos_txt = ''
    for i in range(len(pl)):
        pos_txt += f'\n {pl[i]}'
    vsego_st_rab = round(rabota.vsego_st, 2)
    vsego_st_mat = round(material.vsego_st, 2)
    vsego_st_mob = round(mobilization.vsego_st, 2)
    vsego_st_smeta = round(vsego_st_smeta, 2)
    vsego_st_smeta_nds = round(vsego_st_smeta_nds, 2)
    vsego_smeta = round(vsego_smeta, 2)
    ps = smf_file_data['-PS-']
    finish = '\n\n [Program finished]\n'
    vsego_v_rab = rabota.vsego_v
    vsego_te = round(trudoemkost.vsego_chd, 2)
    mob = mobilization
    vsego_zp = rabota.vsego_zp
    vsego_df = rabota.vsego_df
    discount_mat = smf_file_data['-DISCOUNT_MAT-']
    middle_proiz = round(vsego_v_rab / vsego_te, 2)

    smeta_dict = {
        'filename': filename,
        'isp': isp,
        'zak': zak,
        'date': date,
        'strojka': strojka,
        'object': object,
        'pos_txt': pos_txt,
        'vsego_st_rab': vsego_st_rab,
        'vsego_st_mat': vsego_st_mat,
        'vsego_st_mob': vsego_st_mob,
        'vsego_st_smeta': '{0:,}'.format(vsego_st_smeta).replace(',', ' '),  # с разделением разрядов
        'vsego_st_smeta_nds': '{0:,}'.format(vsego_st_smeta_nds).replace(',', ' '),  # с разделением разрядов
        'vsego_smeta': '{0:,}'.format(vsego_smeta).replace(',', ' '),  # с разделением разрядов
        'ps': ps,
        'finish': '\n\n [Program finished]\n',
        'vsego_v_rab': vsego_v_rab,
        'vsego_te': vsego_te,
        'mob': mob,
        'vsego_zp': vsego_zp,
        'vsego_df': vsego_df,
        'discount_mat': discount_mat,
        'middle_proiz': middle_proiz
    }

    ################

    filename_txt = f" Расчет записан в файл: {smf_file_data['-FILE-'].split('/')[-1]}"
    isp_txt = f"\n\n\n Подрядчик: {smf_file_data['-CONTRACTOR-']}"
    zak_txt = f"\n Заказчик: {smf_file_data['-CLIENT-']}"
    date_txt = f"\n\n                     Локальный сметный расчет от {smf_file_data['-DATE-']} года"
    strojka_txt = f"\n\n Стройка: {smf_file_data['-CONSTRUCTION-']}"
    object_txt = f"\n Об`ект: {smf_file_data['-OBJECT-']}\n"
    pos_txt = ''
    for i in range(len(pl)):
        pos_txt += f'\n {pl[i]}'
    vsego_st_rab_txt = f'\n Всего стоимость работ: {round(rabota.vsego_st, 2)} p'
    vsego_st_mat_txt = f'\n Всего стоимость материалов: {round(material.vsego_st, 2)} p'
    vsego_st_mob_txt = f'\n Всего стоимость мобилизации: {round(mobilization.vsego_st, 2)} p'
    # vsego_st_smeta_txt = f'\n\n ВСЕГО ПО СМЕТЕ.............................................: {round(vsego_st_smeta, 2)} p'
    vsego_st_smeta_txt = f'\n\n Всего по смете без НДС.....................................: {smeta_dict["vsego_st_smeta"]} p'
    vsego_st_smeta_nds_txt = f'\n кроме того НДС-20%.........................................: {smeta_dict["vsego_st_smeta_nds"]} p'
    vsego_smeta_txt = f'\n\n Всего по смете с НДС-20%...................................: {smeta_dict["vsego_smeta"]} p'
    ps_txt = f"\n\n{smf_file_data['-PS-']}"
    finish_txt = '\n\n [Program finished]\n'
    vsego_v_rab_txt = f'\n\n • Всего об`емов работ: {rabota.vsego_v} м2'
    vsego_te_txt = f'\n • Трудоемкость работ: {round(trudoemkost.vsego_chd, 2)} чел.д.'
    middle_proiz_txt = f'\n • Сред. произ-ть работ: {round(middle_proiz, 2)} м2/чел.д.'
    discount_mat_txt = f'\n • Скидка на материалы: {round(discount_mat, 2)} %'
    mob_txt = f'\n\n\t• {mobilization}'
    vsego_zp_txt = f'\n\n • ФОТ чистыми всего: {rabota.vsego_zp} р'
    vsego_df_txt = f'\n • Доход фирмы чистыми всего: {rabota.vsego_df} р'

    smeta_txt = (
            filename_txt +
            isp_txt +
            zak_txt +
            date_txt +
            strojka_txt +
            object_txt +
            pos_txt +
            vsego_st_rab_txt +
            vsego_st_mat_txt +
            vsego_st_mob_txt +
            vsego_st_smeta_txt +
            vsego_st_smeta_nds_txt +
            vsego_smeta_txt +
            ps_txt +
            finish_txt +
            vsego_v_rab_txt +
            vsego_te_txt +
            middle_proiz_txt +
            discount_mat_txt +
            mob_txt +
            vsego_zp_txt +
            vsego_df_txt
    )

    return smeta_txt, smeta_dict
