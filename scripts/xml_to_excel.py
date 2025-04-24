# Путь в проекте: project/scripts/xml_to_excel.py
import sys
import os
import pandas as pd
from lxml import etree
import xml.etree.ElementTree as ET


def xml_to_excel(input_file, output_file):
    # Инициализация данных для листов Excel
    podpisant_data = []
    sv_prod_per_data = []
    sv_sch_fakt_data = []
    inf_pol_fhzh1_data = []
    products = []

    # Попытка парсинга XML
    print("Попытка парсинга XML...")
    root = None
    content_str = None
    try:
        # Читаем файл как UTF-8, игнорируя заявленную кодировку
        with open(input_file, 'rb') as file:
            content = file.read()
            if content.startswith(b'\xef\xbb\xbf'):  # Удаляем BOM, если есть
                content = content[3:]
            content_str = content.decode('utf-8', errors='replace')

        # Парсим с lxml
        parser = etree.XMLParser(recover=True, remove_blank_text=True, encoding='utf-8')
        tree = etree.fromstring(content_str.encode('utf-8'), parser=parser)
        root = tree
        print(f"lxml: Корневой элемент найден: {root.tag}")
    except etree.LxmlError as e:
        print(f"lxml: Ошибка парсинга: {e}")
        # Попытка парсинга с xml.etree.ElementTree
        print("\nПопытка парсинга с xml.etree.ElementTree...")
        try:
            tree = ET.fromstring(content_str)
            print(f"ElementTree: Корневой элемент найден: {tree.tag}")
            root = tree
        except ET.ParseError as e:
            print(f"ElementTree: Ошибка парсинга: {e}")
            print("\nНе удалось распарсить XML ни одним методом.")
            print("Первые 200 символов файла:")
            print(repr(content_str[:200]))
            print("Последние 200 символов файла:")
            print(repr(content_str[-200:]))
        except Exception as e:
            print(f"Ошибка при альтернативном парсинге XML: {e}")
    except FileNotFoundError:
        print(f"Файл {input_file} не найден.")
        return
    except Exception as e:
        print(f"Произошла ошибка при парсинге XML: {e}")

    # Если root всё ещё None, сообщаем об ошибке, но продолжаем выполнение
    if root is None:
        print("Не удалось распарсить XML. Файл будет создан с пустыми данными.")
    else:
        # Извлечение данных Подписант
        try:
            podpisant = root.find('.//Подписант')
            if podpisant is not None:
                fio = podpisant.find('ФИО')
                podpisant_entry = {
                    'Должн': podpisant.get('Должн', ''),
                    'СпосПодтПолном': podpisant.get('СпосПодтПолном', ''),
                    'Фамилия': fio.get('Фамилия', '') if fio is not None else '',
                    'Имя': fio.get('Имя', '') if fio is not None else '',
                    'Отчество': fio.get('Отчество', '') if fio is not None else ''
                }
                podpisant_data.append(podpisant_entry)
        except Exception as e:
            print(f"Ошибка при извлечении данных Подписант: {e}")

        # Извлечение данных СвПродПер
        try:
            sv_prod_per = root.find('.//СвПродПер')
            if sv_prod_per is not None:
                sv_per = sv_prod_per.find('СвПер')
                osn_per = sv_per.find('ОснПер') if sv_per is not None else None
                bez_dok = sv_per.find('БезДокОснПер') if sv_per is not None else None
                sv_prod_per_entry = {
                    'ДатаПер': sv_per.get('ДатаПер', '') if sv_per is not None else '',
                    'СодОпер': sv_per.get('СодОпер', '') if sv_per is not None else '',
                    'РеквДатаДок': osn_per.get('РеквДатаДок', '') if osn_per is not None else '',
                    'РеквНаимДок': osn_per.get('РеквНаимДок', '') if osn_per is not None else '',
                    'РеквНомерДок': osn_per.get('РеквНомерДок', '') if osn_per is not None else '',
                    'БезДокОснПер': bez_dok.text if bez_dok is not None else ''
                }
                sv_prod_per_data.append(sv_prod_per_entry)
        except Exception as e:
            print(f"Ошибка при извлечении данных СвПродПер: {e}")

        # Извлечение данных СвСчФакт
        try:
            sv_sch_fakt = root.find('.//СвСчФакт')
            if sv_sch_fakt is not None:
                sv_prod = sv_sch_fakt.find('СвПрод')
                id_sv_prod = sv_prod.find('ИдСв') if sv_prod is not None else None
                sv_yul_uch_prod = id_sv_prod.find('СвЮЛУч') if id_sv_prod is not None else None
                sv_ip_prod = id_sv_prod.find('СвИП') if id_sv_prod is not None else None
                fio_prod = sv_ip_prod.find('ФИО') if sv_ip_prod is not None else None
                adres_prod = sv_prod.find('Адрес') if sv_prod is not None else None
                adr_inf_prod = adres_prod.find('АдрИнф') if adres_prod is not None else None
                adr_rf_prod = adres_prod.find('АдрРФ') if adres_prod is not None else None
                kontakt_prod = sv_prod.find('Контакт') if sv_prod is not None else None

                gruz_ot = sv_sch_fakt.find('ГрузОт')
                on_zhe = gruz_ot.find('ОнЖе') if gruz_ot is not None else None
                gruz_otpr = gruz_ot.find('ГрузОтпр') if gruz_ot is not None else None
                id_sv_gruz_ot = gruz_otpr.find('ИдСв') if gruz_otpr is not None else None
                sv_yul_uch_gruz_ot = id_sv_gruz_ot.find('СвЮЛУч') if id_sv_gruz_ot is not None else None
                adres_gruz_ot = gruz_otpr.find('Адрес') if gruz_otpr is not None else None
                adr_inf_gruz_ot = adres_gruz_ot.find('АдрИнф') if adres_gruz_ot is not None else None
                kontakt_gruz_ot = gruz_otpr.find('Контакт') if gruz_otpr is not None else None

                gruz_poluch = sv_sch_fakt.find('ГрузПолуч')
                id_sv_gruz_poluch = gruz_poluch.find('ИдСв') if gruz_poluch is not None else None
                sv_ip_gruz_poluch = id_sv_gruz_poluch.find('СвИП') if id_sv_gruz_poluch is not None else None
                fio_gruz_poluch = sv_ip_gruz_poluch.find('ФИО') if sv_ip_gruz_poluch is not None else None
                adres_gruz_poluch = gruz_poluch.find('Адрес') if gruz_poluch is not None else None
                adr_rf_gruz_poluch = adres_gruz_poluch.find('АдрРФ') if adres_gruz_poluch is not None else None

                sv_pokup = sv_sch_fakt.find('СвПокуп')
                id_sv_pokup = sv_pokup.find('ИдСв') if sv_pokup is not None else None
                sv_ip_pokup = id_sv_pokup.find('СвИП') if id_sv_pokup is not None else None
                fio_pokup = sv_ip_pokup.find('ФИО') if sv_ip_pokup is not None else None
                adres_pokup = sv_pokup.find('Адрес') if sv_pokup is not None else None
                adr_inf_pokup = adres_pokup.find('АдрИнф') if adres_pokup is not None else None

                dok_podtv = sv_sch_fakt.find('ДокПодтвОтгрНом')
                den_izm = sv_sch_fakt.find('ДенИзм')

                sv_sch_fakt_entry = {
                    'ДатаДок': sv_sch_fakt.get('ДатаДок', ''),
                    'НомерДок': sv_sch_fakt.get('НомерДок', ''),
                    'ОКПО_СвПрод': sv_prod.get('ОКПО', '') if sv_prod is not None else '',
                    'ИННЮЛ_СвПрод': sv_yul_uch_prod.get('ИННЮЛ', '') if sv_yul_uch_prod is not None else '',
                    'КПП_СвПрод': sv_yul_uch_prod.get('КПП', '') if sv_yul_uch_prod is not None else '',
                    'НаимОрг_СвПрод': sv_yul_uch_prod.get('НаимОрг', '') if sv_yul_uch_prod is not None else '',
                    'ИННФЛ_СвПрод': sv_ip_prod.get('ИННФЛ', '') if sv_ip_prod is not None else '',
                    'Фамилия_СвПрод': fio_prod.get('Фамилия', '') if fio_prod is not None else '',
                    'Имя_СвПрод': fio_prod.get('Имя', '') if fio_prod is not None else '',
                    'Отчество_СвПрод': fio_prod.get('Отчество', '') if fio_prod is not None else '',
                    'АдрТекст_СвПрод': adr_inf_prod.get('АдрТекст', '') if adr_inf_prod is not None else '',
                    'КодСтр_СвПрод': adr_inf_prod.get('КодСтр', '') if adr_inf_prod is not None else '',
                    'НаимСтран_СвПрод': adr_inf_prod.get('НаимСтран', '') if adr_inf_prod is not None else '',
                    'КодРегион_СвПрод': adr_rf_prod.get('КодРегион', '') if adr_rf_prod is not None else '',
                    'НаимРегион_СвПрод': adr_rf_prod.get('НаимРегион', '') if adr_rf_prod is not None else '',
                    'Индекс_СвПрод': adr_rf_prod.get('Индекс', '') if adr_rf_prod is not None else '',
                    'Город_СвПрод': adr_rf_prod.get('Город', '') if adr_rf_prod is not None else '',
                    'Улица_СвПрод': adr_rf_prod.get('Улица', '') if adr_rf_prod is not None else '',
                    'Дом_СвПрод': adr_rf_prod.get('Дом', '') if adr_rf_prod is not None else '',
                    'Тлф_СвПрод': kontakt_prod.find('Тлф').text if kontakt_prod is not None and kontakt_prod.find(
                        'Тлф') is not None else '',
                    'ЭлПочта_СвПрод': kontakt_prod.find(
                        'ЭлПочта').text if kontakt_prod is not None and kontakt_prod.find(
                        'ЭлПочта') is not None else '',
                    'ОнЖе_ГрузОт': on_zhe.text if on_zhe is not None else '',
                    'ОКПО_ГрузОт': gruz_otpr.get('ОКПО', '') if gruz_otpr is not None else '',
                    'ИННЮЛ_ГрузОт': sv_yul_uch_gruz_ot.get('ИННЮЛ', '') if sv_yul_uch_gruz_ot is not None else '',
                    'КПП_ГрузОт': sv_yul_uch_gruz_ot.get('КПП', '') if sv_yul_uch_gruz_ot is not None else '',
                    'НаимОрг_ГрузОт': sv_yul_uch_gruz_ot.get('НаимОрг', '') if sv_yul_uch_gruz_ot is not None else '',
                    'АдрТекст_ГрузОт': adr_inf_gruz_ot.get('АдрТекст', '') if adr_inf_gruz_ot is not None else '',
                    'КодСтр_ГрузОт': adr_inf_gruz_ot.get('КодСтр', '') if adr_inf_gruz_ot is not None else '',
                    'НаимСтран_ГрузОт': adr_inf_gruz_ot.get('НаимСтран', '') if adr_inf_gruz_ot is not None else '',
                    'Тлф_ГрузОт': kontakt_gruz_ot.find(
                        'Тлф').text if kontakt_gruz_ot is not None and kontakt_gruz_ot.find('Тлф') is not None else '',
                    'ЭлПочта_ГрузОт': kontakt_gruz_ot.find(
                        'ЭлПочта').text if kontakt_gruz_ot is not None and kontakt_gruz_ot.find(
                        'ЭлПочта') is not None else '',
                    'ОКПО_ГрузПолуч': gruz_poluch.get('ОКПО', '') if gruz_poluch is not None else '',
                    'СокрНаим_ГрузПолуч': gruz_poluch.get('СокрНаим', '') if gruz_poluch is not None else '',
                    'ИННФЛ_ГрузПолуч': sv_ip_gruz_poluch.get('ИННФЛ', '') if sv_ip_gruz_poluch is not None else '',
                    'СвГосРегИП_ГрузПолуч': sv_ip_gruz_poluch.get('СвГосРегИП',
                                                                  '') if sv_ip_gruz_poluch is not None else '',
                    'Фамилия_ГрузПолуч': fio_gruz_poluch.get('Фамилия', '') if fio_gruz_poluch is not None else '',
                    'Имя_ГрузПолуч': fio_gruz_poluch.get('Имя', '') if fio_gruz_poluch is not None else '',
                    'Отчество_ГрузПолуч': fio_gruz_poluch.get('Отчество', '') if fio_gruz_poluch is not None else '',
                    'Дом_ГрузПолуч': adr_rf_gruz_poluch.get('Дом', '') if adr_rf_gruz_poluch is not None else '',
                    'Индекс_ГрузПолуч': adr_rf_gruz_poluch.get('Индекс', '') if adr_rf_gruz_poluch is not None else '',
                    'КодРегион_ГрузПолуч': adr_rf_gruz_poluch.get('КодРегион',
                                                                  '') if adr_rf_gruz_poluch is not None else '',
                    'НаимРегион_ГрузПолуч': adr_rf_gruz_poluch.get('НаимРегион',
                                                                   '') if adr_rf_gruz_poluch is not None else '',
                    'Улица_ГрузПолуч': adr_rf_gruz_poluch.get('Улица', '') if adr_rf_gruz_poluch is not None else '',
                    'ОКПО_СвПокуп': sv_pokup.get('ОКПО', '') if sv_pokup is not None else '',
                    'СокрНаим_СвПокуп': sv_pokup.get('СокрНаим', '') if sv_pokup is not None else '',
                    'ИННФЛ_СвПокуп': sv_ip_pokup.get('ИННФЛ', '') if sv_ip_pokup is not None else '',
                    'СвГосРегИП_СвПокуп': sv_ip_pokup.get('СвГосРегИП', '') if sv_ip_pokup is not None else '',
                    'Фамилия_СвПокуп': fio_pokup.get('Фамилия', '') if fio_pokup is not None else '',
                    'Имя_СвПокуп': fio_pokup.get('Имя', '') if fio_pokup is not None else '',
                    'Отчество_СвПокуп': fio_pokup.get('Отчество', '') if fio_pokup is not None else '',
                    'АдрТекст_СвПокуп': adr_inf_pokup.get('АдрТекст', '') if adr_inf_pokup is not None else '',
                    'КодСтр_СвПокуп': adr_inf_pokup.get('КодСтр', '') if adr_inf_pokup is not None else '',
                    'НаимСтран_СвПокуп': adr_inf_pokup.get('НаимСтран', '') if adr_inf_pokup is not None else '',
                    'РеквДатаДок_ДокПодтв': dok_podtv.get('РеквДатаДок', '') if dok_podtv is not None else '',
                    'РеквНаимДок_ДокПодтв': dok_podtv.get('РеквНаимДок', '') if dok_podtv is not None else '',
                    'РеквНомерДок_ДокПодтв': dok_podtv.get('РеквНомерДок', '') if dok_podtv is not None else '',
                    'КодОКВ': den_izm.get('КодОКВ', '') if den_izm is not None else '',
                    'НаимОКВ': den_izm.get('НаимОКВ', '') if den_izm is not None else ''
                }
                sv_sch_fakt_data.append(sv_sch_fakt_entry)
        except Exception as e:
            print(f"Ошибка при извлечении данных СвСчФакт: {e}")

        # Извлечение данных ИнфПолФХЖ1
        try:
            inf_pol_fhzh1 = root.find('.//ИнфПолФХЖ1')
            if inf_pol_fhzh1 is not None:
                for text_inf in inf_pol_fhzh1.findall('ТекстИнф'):
                    inf_pol_entry = {
                        'Идентиф': text_inf.get('Идентиф', ''),
                        'Значен': text_inf.get('Значен', '')
                    }
                    inf_pol_fhzh1_data.append(inf_pol_entry)
        except Exception as e:
            print(f"Ошибка при извлечении данных ИнфПолФХЖ1: {e}")

        # Извлечение данных о товарах
        try:
            for item in root.findall('.//ТаблСчФакт/СведТов'):
                product = {
                    'Номер строки': item.get('НомСтр', ''),
                    'Наименование': item.get('НаимТов', ''),
                    'Количество': item.get('КолТов', ''),
                    'Ед. измерения': item.get('НаимЕдИзм', ''),
                    'Цена': item.get('Ц ollнаТов', ''),
                    'Стоимость без НДС': item.get('СтТовБезНДС', ''),
                    'Сумма НДС': item.find('СумНал/СумНал').text if item.find('СумНал/СумНал') is not None else '',
                    'Стоимость с НДС': item.get('СтТовУчНал', ''),
                    'Ставка НДС': item.get('НалСт', ''),
                    'Код товара': item.find('ДопСведТов').get('КодТов', '') if item.find(
                        'ДопСведТов') is not None else '',
                    'ОКЕИ_Тов': item.get('ОКЕИ_Тов', ''),
                    'GTIN': '',
                    'КодПокупателя': '',
                    'НазваниеПокупателя': '',
                    'КИЗ': ''
                }

                # Извлекаем все ИнфПолФХЖ2
                for inf in item.findall('.//ИнфПолФХЖ2'):
                    ident = inf.get('Идентиф', '')
                    value = inf.get('Значен', '')
                    if ident == 'GTIN':
                        product['GTIN'] = value
                    elif ident == 'КодПокупателя':
                        product['КодПокупателя'] = value
                    elif ident == 'НазваниеПокупателя':
                        product['НазваниеПокупателя'] = value

                # Добавляем КИЗ если они есть
                kiz_list = [kiz.text for kiz in item.findall('.//КИЗ') if kiz.text is not None]
                product['КИЗ'] = '; '.join(kiz_list) if kiz_list else ''

                products.append(product)
        except Exception as e:
            print(f"Ошибка при обработке данных товаров: {e}")

    # Создаем DataFrame для товаров
    columns_order = [
        'Номер строки', 'Наименование', 'Код товара', 'Количество', 'Ед. измерения',
        'Цена', 'Стоимость без НДС', 'Ставка НДС', 'Сумма НДС', 'Стоимость с НДС',
        'ОКЕИ_Тов', 'GTIN', 'КодПокупателя', 'НазваниеПокупателя', 'КИЗ'
    ]
    try:
        if products:
            df_products = pd.DataFrame(products)
            # Убедимся, что все столбцы присутствуют
            for col in columns_order:
                if col not in df_products.columns:
                    df_products[col] = ''
            df_products = df_products[columns_order]
        else:
            # Если товаров нет, создаём пустой DataFrame с нужными столбцами
            df_products = pd.DataFrame(columns=columns_order)
    except Exception as e:
        print(f"Ошибка при создании DataFrame для товаров: {e}")
        df_products = pd.DataFrame(columns=columns_order)

    # Сохраняем все данные в Excel с несколькими листами
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Лист с товарами
            df_products.to_excel(writer, sheet_name='Товары', index=False)

            # Лист Подписант
            df_podpisant = pd.DataFrame(podpisant_data if podpisant_data else [{}])
            df_podpisant.to_excel(writer, sheet_name='Подписант', index=False)

            # Лист СвПродПер
            df_sv_prod_per = pd.DataFrame(sv_prod_per_data if sv_prod_per_data else [{}])
            df_sv_prod_per.to_excel(writer, sheet_name='СвПродПер', index=False)

            # Лист СвСчФакт
            df_sv_sch_fakt = pd.DataFrame(sv_sch_fakt_data if sv_sch_fakt_data else [{}])
            df_sv_sch_fakt.to_excel(writer, sheet_name='СвСчФакт', index=False)

            # Лист ИнфПолФХЖ1
            df_inf_pol = pd.DataFrame(inf_pol_fhzh1_data if inf_pol_fhzh1_data else [{}])
            df_inf_pol.to_excel(writer, sheet_name='ИнфПолФХЖ1', index=False)

        print(
            f"Таблица успешно сохранена в {output_file} с листами: Товары, Подписант, СвПродПер, СвСчФакт, ИнфПолФХЖ1")
    except Exception as e:
        print(f"Ошибка при сохранении Excel: {e}")
