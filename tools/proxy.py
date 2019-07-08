#!/usr/bin/env python3

import sys
from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import TimeoutError

# Инициализация и запуск 
# todo вынести path 
app = Application(backend='uia').start('C:\Program Files (x86)\Declaration\Dec2018\Decl2018.exe')
main_window = app.window(title='Без имени - Декларация 2018')
main_window.set_focus()

def _select_tab(num: int):
    step = int(main_window.rectangle().height() / 9)
    offset = step * num
    navigation_panel.click_input(coords=(0, offset))

def _select_tab_toolpanel(num: int):
    step = int(main_window.rectangle().height() / 9)
    offset = step * num
    tool_panel.click_input(coords=(0, offset))

def _select_radio(group: str, v: str):
    group = main_window.window(title='  ' + group)
    c = group.window(title=v)
    c.click_input()

def _type_to_onlyone_input_by_index(group: str, text: str, class_name: str):
	parent = main_window.window(title=group)
	child = parent.child_window()
	child.set_text(text)
	
def _type_to_input_by_index(group: str, ind: int, text: str, class_name: str):
    inpts = main_window.window(title='  ' + group).descendants(class_name=class_name)
    inpts[ind].set_text(text)
	
def _type_to_input_by_index_mask(group: str, ind: int, text: str, class_name: str):
    inpts = main_window.window(title='  ' + group).descendants(class_name=class_name)
    inpts[0].set_text(text) 	

def _set_spinner(group: str, ind: int, v: int):
    group = main_window.window(title='  ' + group)
    bs = group.descendants()
    b = bs[ind]
    b.set_text(v)

def _select_from_dict(group: str, win_title: int, v: str):
    group = main_window.window(title='  ' + group)
    b = group.window(title='...')
    b.click_input()

    dw = app.window(title=win_title)
    l = dw.window(class_name="TListView")
    l.set_focus()
    vl = l.window(title=v)

    for _ in range(1000):
        try:
            #if vl.is_visible(timeout=0.5):
             vl.click_input()
             break
        except ElementNotFoundError:
            send_keys('{PGDN}')

    ok_b = dw.window(title='Да')
    ok_b.click_input()
	
def _check_checkbox(group: str, name: str):
    cb = main_window.window(title=name)
    cb.click_input()


labels2indTab1 = {
    'Номер инспекции': ('Общая информация', 'Справочник инспекций', 0), 
    'Номер корректировки': ('Общая информация', 4, 0), 
    'ОКТМО': ('Общая информация', 0, 0), 

    'Индивидуальный предприниматель': ('Признак налогоплательщика', None, 0), 
    'Глава фермерского хозяйства': ('Признак налогоплательщика', None, 0), 
    'Частный нотариус': ('Признак налогоплательщика', None, 0), 
    'Адвокат': ('Признак налогоплательщика', None, 0), 
    'Иное физическое лицо': ('Признак налогоплательщика', None, 0), 
    'Арбитражный управляющий': ('Признак налогоплательщика', None, 0), 

    'Учитываемые "справками о доходах физического лица", доходы по договорам': ('Имеются доходы', None, 0), 
    'В иностранной валюте': ('Имеются доходы', None, 0), 
    'От предпринимательской деятельности': ('Имеются доходы', None, 0), 

    'Лично': ('Достоверность подтверждается', None, 0), 
    'Представителем - ФЛ': ('Достоверность подтверждается', None, 0), 
    'ФИО': ('Достоверность подтверждается', 1, 0), 
    'Документ': ('Достоверность подтверждается', 0, 0),
    'Фамилия': ('Ф.И.О.', 4, 1),
    'Имя': ('Ф.И.О.', 3, 1),	
    'Отчество': ('Ф.И.О.', 2, 1),
	'ИНН': ('Ф.И.О.', 1, 1),
	'Дата рождения': ('Ф.И.О.', 0, 1),
    'Место рождения': ('Ф.И.О.', 0, 1),
	'Вид документа': ('Сведения о документе, удостоверяющем личность', 'Справочник видов документов', 1),
	'Серия и номер': ('Сведения о документе, удостоверяющем личность', 0, 1),
	'Дата выдачи': ('Сведения о документе, удостоверяющем личность', 0, 1),
	'Кем выдан': ('Сведения о документе, удостоверяющем личность', 1, 1),
	'Контактный телефон': ('Контактный телефон', 0, 1),
	'Наименование источника выплаты': ('Наименование источника выплаты', 0, 0),
	'ИНН Источника выплаты': ('ИНН', 1, 0),
	'КПП Источника выплаты': ('КПП', 2, 0),
	'ОКТМО Источника выплаты': ('ОКТМО', 3, 0),
	'Чекбокс Источник выплаты1': ('Расчет стандартных вычетов вести по этому источнику', None, 0),
	'Чекбокс Источника выплаты2': ('Источник является инвестиционным товариществом', None, 0),
	'Код дохода': ('Сведения о полученном доходе', 'Справочник видов доходов', 2)	
}

def select_from_dict(name: str, v: str):
    group, ind, tb = labels2indTab1[name]
    _select_tab(tb) 
    _select_from_dict(group, ind, v)

def select_from_dict_from_modal_window(name: str, v: str):
    group, ind, tb = labels2indTab1[name]
    _select_from_dict_from_modal_window(group, ind, v)	
	

def set_spinner(name: str, v: int):
    group, ind, tb = labels2indTab1[name]
    _select_tab(tb) 
    _set_spinner(group, ind, v)

def fill_only_one_input(name: str, v: str, class_name: str):
    group, ind, tb = labels2indTab1[name]
    _select_tab(tb) 
    _type_to_onlyone_input_by_index(group, v, class_name)	

def fill_input(name: str, v: str, class_name: str):
    group, ind, tb = labels2indTab1[name]
    _select_tab(tb) 
    _type_to_input_by_index(group, ind, v, class_name)

def select_radio(name: str):
    group, _, tb = labels2indTab1[name]
    _select_tab(tb) 
    _select_radio(group, name)

def select_checkbox(name: str):
    group, _, tb = labels2indTab1[name]
    _select_tab(tb) 
    _check_checkbox(group, name)

def fill_input_mask(name: str, v: str, class_name: str):
    group, ind, tb = labels2indTab1[name]
    _select_tab(tb) 
    _type_to_input_by_index_mask(group, ind, v, class_name)	

# Начало заполнения 
# 1.Блок Задание условий 
navigation_panel = main_window.descendants(class_name="TPanel")[1]

# RETURN THIS select_from_dict('Номер инспекции', '2411')
select_from_dict('Номер инспекции', '0101')
set_spinner('Номер корректировки', 0)
fill_input('ОКТМО', '04701000', 'TMaskedEdit')
select_radio('Иное физическое лицо')
#первое устанавливается по-умолчанию
#select_checkbox('Учитываемые "справками о доходах физического лица", доходы по договорам')
select_radio('Лично')

# 2.Сведения о декларанте
_select_tab(1) 
fill_input('Фамилия', 'ИВАНОВА', 'TMaskedEdit')
fill_input('Имя', 'ЮЛИЯ','TMaskedEdit')
fill_input('Отчество', 'ВАЛЕРЬЕВНА', 'TMaskedEdit')
fill_input('ИНН', '780204893183', 'TMaskedEdit')
fill_input('Место рождения', 'Г.РОСТОВ-НА-ДОНУ', 'TMaskedEdit')
fill_input('Дата рождения', '02.06.1982', 'TTestMaskD')

fill_only_one_input('Контактный телефон', '89135555555', 'TMaskedEdit')
select_from_dict('Вид документа', '21')
fill_input_mask('Серия и номер', '04 09 934566', 'TMaskedDocEdit')
fill_input_mask('Кем выдан', 'ТП №75 ОТДЕЛА УФМС РОССИИ ПО САНКТ-ПЕТЕРБУРГУ И ЛЕНИНГРАДСКОЙ ОБЛ. В ФРУНЗЕНСКОМ Р-НЕ ГОР.САНКТ-ПЕТЕРБУРГА КОД ПОДРАЗДЕЛЕНИЯ 780-075', 'TMaskedEdit')
fill_input('Дата выдачи', '02.03.2002', 'TTestMaskD')

# 3.Доходы, полученные в РФ
_select_tab(2)
## Только ли 13 всегда ?? 
## Клик по кнопке Добавить источник выплат  
add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
add_source_panel[8].click_input(coords=(10, 55))
## Заполнение модального окна Источник выплат

dw = app.window(title='Источник выплаты')
inner_edit = dw.window(class_name="TPanel").descendants();
inner_edit[1].type_keys('04701000')
inner_edit[3].type_keys('2464001001')
inner_edit[4].type_keys('2464231675')

inner_edit[5].type_keys('IKEA')

inner_checkbox = dw.window(class_name="TPanel").descendants(class_name="TCheckBox")
inner_checkbox[1].click_input()

ok_b = dw.window(title='Да')
ok_b.click_input()

edit_form = add_money_source_allpanel.descendants(class_name='TMaskedEdit')
edit_form[1].set_text('14316')

## Заполнение модального окна Сведения о доходе 

add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
add_source_panel[8].click_input(coords=(10, 155))


finance_buttons = app.window(title='Сведения о доходе').descendants(class_name="TButton")
finance_buttons[3].click_input();

## Справочник видов доходов
dw = app.window(title='Справочник видов доходов')
l = dw.window(class_name="TListView")
l.set_focus()
vl = l.window(title='1011')

for _ in range(1000):
    try:
        #if vl.is_visible(timeout=0.5):
        vl.click_input()
        break
    except ElementNotFoundError:
        send_keys('{PGDN}')

ok_b = dw.window(title='Да')
ok_b.click_input()

## Справочник видов доходов 

edit_buttons = app.window(title='Сведения о доходе').descendants(class_name="TMaskedEdit")
edit_buttons[2].set_text('10000')
edit_buttons[1].set_text('1')

app.window(title='Сведения о доходе').window(title='Да').click_input()

## Конец заполнения модального окна Сведения о доходе 

add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
add_source_panel[8].click_input(coords=(10, 155))


finance_buttons = app.window(title='Сведения о доходе').descendants(class_name="TButton")
finance_buttons[3].click_input();

## Справочник видов доходов
dw = app.window(title='Справочник видов доходов')
l = dw.window(class_name="TListView")
l.set_focus()
vl = l.window(title='1011')

for _ in range(1000):
    try:
        #if vl.is_visible(timeout=0.5):
        vl.click_input()
        break
    except ElementNotFoundError:
        send_keys('{PGDN}')

ok_b = dw.window(title='Да')
ok_b.click_input()

## Справочник видов доходов 

edit_buttons = app.window(title='Сведения о доходе').descendants(class_name="TMaskedEdit")
edit_buttons[2].set_text('10000')
edit_buttons[1].set_text('2')

app.window(title='Сведения о доходе').window(title='Да').click_input()


#### 3 month 
add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
add_source_panel[8].click_input(coords=(10, 155))


finance_buttons = app.window(title='Сведения о доходе').descendants(class_name="TButton")
finance_buttons[3].click_input();

## Справочник видов доходов
dw = app.window(title='Справочник видов доходов')
l = dw.window(class_name="TListView")
l.set_focus()
vl = l.window(title='1011')

for _ in range(1000):
    try:
        #if vl.is_visible(timeout=0.5):
        vl.click_input()
        break
    except ElementNotFoundError:
        send_keys('{PGDN}')

ok_b = dw.window(title='Да')
ok_b.click_input()

## Справочник видов доходов 

edit_buttons = app.window(title='Сведения о доходе').descendants(class_name="TMaskedEdit")
edit_buttons[2].set_text('10000')
edit_buttons[1].set_text('3')

app.window(title='Сведения о доходе').window(title='Да').click_input()

### 3 month 

add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
add_source_panel[8].click_input(coords=(10, 455))

finance_buttons = app.window(title='Вычеты, указанные в разделе 3 справки 2-ндфл').descendants(class_name="TButton")
finance_buttons[2].click_input();

## Справочник других доходов
dw = app.window(title='Справочник видов вычетов')
l = dw.window(class_name="TListView")
l.set_focus()
vl = l.window(title='126')

for _ in range(1000):
    try:
        #if vl.is_visible(timeout=0.5):
        vl.click_input()
        break
    except ElementNotFoundError:
        send_keys('{PGDN}')

ok_b = dw.window(title='Да')
ok_b.click_input()

## Справочник других доходов
edit_buttons = app.window(title='Вычеты, указанные в разделе 3 справки 2-ндфл').descendants(class_name='TMaskedEdit')
edit_buttons[0].set_text('16800')
app.window(title='Вычеты, указанные в разделе 3 справки 2-ндфл').window(title='Да').click_input()

# 3.Вычеты 
_select_tab(5) 

check_boxes = main_window.descendants(class_name='TCheckBox')
check_boxes[0].click_input()


add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
add_source_panel[0].click_input(coords=(10, 55))

## Заполнение формы по данным объекта

combo_boxes = app.window(title='Данные объекта').descendants(class_name='TComboBox')
# Код номера объекта  Кадастровый номер 
combo_boxes[0].click_input()
cmb1 = combo_boxes[0].descendants()
cmb1[1].click_input()
# 1 - Кадастровый номер 

# Признак налогоплательщика Собственник объекта 
combo_boxes[1].click_input()
cmb1 = combo_boxes[1].descendants()
cmb1[1].click_input()
# 0 - Родитель несовершеннолетнего собственника объекта
# 1 - Собственник объекта 


# Наименование объекта 
combo_boxes[2].click_input()
cmb1 = combo_boxes[2].descendants()
cmb1[2].click_input()
# 0 - Земельный участок для индивидуального жилищного строительства 
# 1 - Жилой дом
# 2 - Квартира

masked_edits = app.window(title='Данные объекта').descendants(class_name='TMaskedEdit')
masked_edits[5].set_text('24:50:0400417:1739')
masked_edits[4].set_text('660135 Красноярский край, г.Красноярск, ул. Тополей, д.5, кв. 1')
masked_edits[1].set_text('2000000')
masked_edits[0].set_text('204090')

masked_date_edits = app.window(title='Данные объекта').descendants(class_name='TTestMaskD')
masked_date_edits[1].set_text('17.06.2014')

app.window(title='Данные объекта').window(title='Да').click_input()

### Сохранить декларацию
manage_btns = main_window.descendants(class_name='TToolBar')
manage_btns[1].click_input() # Сохранить 
###

## Выбор папки 



dw = app.window(title='Обзор папок')
l = dw.descendants()[1]
item = l.get_item([r'\Рабочий стол\dec'])
item.select()
#item.click_input()

dw.descendants(class_name='Button')[0].click_input()

### Last Step
app.window(title='Декларация 2018').descendants(class_name='Button')[0].click_input()

app.window(title='Декларация 2018').descendants(class_name='Button')[0].click_input()



main_window.close()


#main_window.print_control_identifiers()
