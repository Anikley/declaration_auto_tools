#!/usr/bin/env python3
# -*- coding: utf8 -*-
#!flask/bin/python
import sys
import json
from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import TimeoutError
from flask import Flask, jsonify, abort, request, make_response, url_for
from flask_httpauth import HTTPBasicAuth
import os
import base64

apl = Flask(__name__, static_url_path = "")
path1 = (r'‪C:\\Users\yota_\Desktop').replace('\u202a','')

def clear_folder():	
	for file in os.listdir(path1):
		if file.endswith(".xml"):
			path_delete = os.path.join(path1, file)
			os.remove(path_delete)

def result_in_base64():	
	for file in os.listdir(path1):
		if file.endswith(".xml"):
			path_delete = os.path.join(path1, file)			
			with open(path_delete, "rb") as image_file:
				encoded_string = base64.b64encode(image_file.read())
				return encoded_string
	

def run_declaration_app(path:str, data_json: str):
	print("Read input params")
	print('First arg - path to Declaration 2018') 	
	print('Second arg - data json for Declaration 2018')
	
	## todo проверка входных параметров 
	
	data = json.loads(data_json)
	number_i= data['declaration']['block1']['number_i'] # 0101
	number_k= int(data['declaration']['block1']['number_k']) # 0
	oktmo = data['declaration']['block1']['oktmo'] # '04701000'
	radio_btn1 = data['declaration']['block1']['radio_btn1'] # 'Иное физическое лицо'
	radio_btn2 = data['declaration']['block1']['radio_btn2'] # 'Лично'

	last_name = data['declaration']['block2']['lastname'] # 'ИВАНОВА'
	first_name = data['declaration']['block2']['firstname'] # 'ЮЛИЯ'
	sir_name = data['declaration']['block2']['sirname'] #  ВАЛЕРЬЕВНА
	inn = data['declaration']['block2']['inn'] # '780204893183'
	birth_place = data['declaration']['block2']['birth_place'] # 'Г.РОСТОВ-НА-ДОНУ' 
	birth_date = data['declaration']['block2']['birth_date'] # '02.06.1982'
	phone = data['declaration']['block2']['phone'] # '89135555555'
	type_doc = data['declaration']['block2']['type_doc'] # '21'
	doc_number = data['declaration']['block2']['doc_number'] #  '04 09 934566'
	doc_from = data['declaration']['block2']['doc_from'] # 'ТП №75 ОТДЕЛА УФМС РОССИИ ПО САНКТ-ПЕТЕРБУРГУ И ЛЕНИНГРАДСКОЙ ОБЛ. В ФРУНЗЕНСКОМ Р-НЕ ГОР.САНКТ-ПЕТЕРБУРГА КОД ПОДРАЗДЕЛЕНИЯ 780-075'
	doc_date = data['declaration']['block2']['doc_date'] # '02.03.2002'

	money_source_company = data['declaration']['block3']['money_source_company'] 
	#   '04701000'
	#   '2464001001'
	#   '2464231675'
	#    PRAVOCARD
	money_dec = data['declaration']['block3']['money_dec']  # '14316'
	money_source = data['declaration']['block3']['money_source']

	money_source_bottom_table = data['declaration']['block3']['money_source_bottom_table']

	code_number_object = data['declaration']['block5']['code_number_object'] # 1
	sign_taxplayer = data['declaration']['block5']['sign_taxplayer'] # 1
	object_name = data['declaration']['block5']['object_name'] # 2 
	object_data = data['declaration']['block5']['object_data']
	app = Application(backend='uia').start(path)
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

	print('Start new decl filling....')
	print('Block №1 - Задание условий')

	navigation_panel = main_window.descendants(class_name="TPanel")[1]
	
	select_from_dict('Номер инспекции', number_i)
	set_spinner('Номер корректировки', number_k)
	fill_input('ОКТМО', oktmo, 'TMaskedEdit')
	select_radio(radio_btn1)

	## todo проверка - добавляется количество заполняемых данных 
	## первое устанавливается по-умолчанию
	## select_checkbox('Учитываемые "справками о доходах физического лица", доходы по договорам')

	select_radio(radio_btn2)
	print('Block №1 - Сведения о декларанте')
	_select_tab(1) 
	fill_input('Фамилия', last_name, 'TMaskedEdit')
	fill_input('Имя', first_name,'TMaskedEdit')
	fill_input('Отчество', sir_name, 'TMaskedEdit')
	fill_input('ИНН', inn , 'TMaskedEdit')
	fill_input('Место рождения', birth_place, 'TMaskedEdit')
	fill_input('Дата рождения', birth_date, 'TTestMaskD')

	fill_only_one_input('Контактный телефон', phone, 'TMaskedEdit')
	select_from_dict('Вид документа', type_doc)
	fill_input_mask('Серия и номер', doc_number , 'TMaskedDocEdit')
	fill_input_mask('Кем выдан', doc_from , 'TMaskedEdit')
	fill_input('Дата выдачи', doc_date, 'TTestMaskD')

	print('Block №1 - Доходы, полученные в РФ')

	_select_tab(2)
	## todo пока клик настроен только на кнопку 13 в панели
	## она устанавливается по-умолчанию 
	print('Add money source')
	add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
	add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
	add_source_panel[8].click_input(coords=(10, 55))

	print('Source money Modal window filling')
	dw = app.window(title='Источник выплаты')
	inner_edit = dw.window(class_name="TPanel").descendants();
	inner_edit[1].type_keys(money_source_company[0])
	inner_edit[3].type_keys(money_source_company[1])
	inner_edit[4].type_keys(money_source_company[2])

	## todo проверить варианты ввода на русском языке 
	## пока проблема 
	inner_edit[5].type_keys(money_source_company[3])

	inner_checkbox = dw.window(class_name="TPanel").descendants(class_name="TCheckBox")
	inner_checkbox[1].click_input()
	ok_b = dw.window(title='Да')
	ok_b.click_input()
	edit_form = add_money_source_allpanel.descendants(class_name='TMaskedEdit')
	edit_form[1].set_text(money_dec)

	print('Info money source Modal Window filling')

	for money_source_item in money_source:
		add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
		add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
		add_source_panel[8].click_input(coords=(10, 155))
		finance_buttons = app.window(title='Сведения о доходе').descendants(class_name="TButton")
		finance_buttons[3].click_input();

		print('Open dictionary of the type of money')
		dw = app.window(title='Справочник видов доходов')
		l = dw.window(class_name="TListView")
		l.set_focus()
		vl = l.window(title= money_source_item['code'])
		for _ in range(1000):
			try:
				#if vl.is_visible(timeout=0.5):
				vl.click_input()
				break
			except ElementNotFoundError:
				send_keys('{PGDN}')

		ok_b = dw.window(title='Да')
		ok_b.click_input()

		print('Dictionary of the type of money source')
		edit_buttons = app.window(title='Сведения о доходе').descendants(class_name="TMaskedEdit")
		edit_buttons[2].set_text( money_source_item['money_sum'])
		edit_buttons[1].set_text( money_source_item['month'])

		app.window(title='Сведения о доходе').window(title='Да').click_input()
	
	print('Bottom table filling')
	add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
	add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
	add_source_panel[8].click_input(coords=(10, 455))

	finance_buttons = app.window(title='Вычеты, указанные в разделе 3 справки 2-ндфл').descendants(class_name="TButton")
	finance_buttons[2].click_input()

	dw = app.window(title='Справочник видов вычетов')
	l = dw.window(class_name="TListView")
	l.set_focus()
	vl = l.window(title= money_source_bottom_table['code'])

	for _ in range(1000):
		try:
			#if vl.is_visible(timeout=0.5):
			vl.click_input()
			break
		except ElementNotFoundError:
			send_keys('{PGDN}')

	ok_b = dw.window(title='Да')
	ok_b.click_input()

	print('Get dictionary another money source')
	edit_buttons = app.window(title='Вычеты, указанные в разделе 3 справки 2-ндфл').descendants(class_name='TMaskedEdit')
	edit_buttons[0].set_text(money_source_bottom_table['money']) # 16800
	app.window(title='Вычеты, указанные в разделе 3 справки 2-ндфл').window(title='Да').click_input()

	print("Block 3")

	_select_tab(5) 
	check_boxes = main_window.descendants(class_name='TCheckBox')
	check_boxes[0].click_input()

	add_money_source_allpanel = main_window.descendants(class_name='TPanel')[0]
	add_source_panel = add_money_source_allpanel.descendants(class_name='TPanel')
	add_source_panel[0].click_input(coords=(10, 55))

	print('Objects Data Modal Window filling')

	combo_boxes = app.window(title='Данные объекта').descendants(class_name='TComboBox')
	# Код номера объекта  Кадастровый номер 
	combo_boxes[0].click_input()
	cmb1 = combo_boxes[0].descendants()
	cmb1[int(code_number_object)].click_input()
	# 1 - Кадастровый номер 

	# Признак налогоплательщика Собственник объекта 
	combo_boxes[1].click_input()
	cmb1 = combo_boxes[1].descendants()
	cmb1[int(sign_taxplayer)].click_input()
	# 0 - Родитель несовершеннолетнего собственника объекта
	# 1 - Собственник объекта 

	# Наименование объекта 
	combo_boxes[2].click_input()
	cmb1 = combo_boxes[2].descendants()
	cmb1[int(object_name)].click_input()
	# 0 - Земельный участок для индивидуального жилищного строительства 
	# 1 - Жилой дом
	# 2 - Квартира

	masked_edits = app.window(title='Данные объекта').descendants(class_name='TMaskedEdit')

	masked_edits[5].set_text(object_data['object_number'])
	masked_edits[4].set_text(object_data['object_address']) 
	masked_edits[1].set_text(object_data['object_sum'])
	masked_edits[0].set_text(object_data['object_del'])

	masked_date_edits = app.window(title='Данные объекта').descendants(class_name='TTestMaskD')
	masked_date_edits[1].set_text(object_data['object_date']) 

	app.window(title='Данные объекта').window(title='Да').click_input()

	### Сохранить декларацию
	manage_btns = main_window.descendants(class_name='TToolBar')
	manage_btns[1].click_input() 
	print('Choose folder to save...')

	dw = app.window(title='Обзор папок')
	l = dw.descendants()[1]
	#todo перенести в параметры
	item = l.get_item([r'\Рабочий стол\dec'])
	item.select()

	dw.descendants(class_name='Button')[0].click_input()

	### Last Step
	app.window(title='Декларация 2018').descendants(class_name='Button')[0].click_input()

	## todo проверка на ошибку - куча всплывающих окон может быть 
	#app.window(title='Декларация 2018').descendants(class_name='Button')[0].click_input()

	print('Declaration is Ready!')
	app.kill()


@apl.route('/declaration/api/v1.0/get', methods = ['POST'])
def create_declaration():
	clear_folder()
	run_declaration_app("C:\\Program Files (x86)\\Declaration\\Dec2018\\Decl2018.exe", request.json)
	result_base_64 = result_in_base64()
	clear_folder()
	return result_base_64 , 201
  
if __name__ == '__main__':
    apl.run(debug = True)
	

