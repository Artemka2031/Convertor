import xml.etree.ElementTree as ET
from collections import defaultdict

import pandas as pd
from lxml import etree

from utils import setup_logging

logger = setup_logging()

EXPECTED_TAGS = {
    'Файл', 'Подписант', 'ФИО', 'СвПродПер', 'СвПер', 'ОснПер', 'БезДокОснПер',
    'СвСчФакт', 'СвПрод', 'ИдСв', 'СвЮЛУч', 'СвИП', 'Адрес', 'АдрИнф', 'АдрРФ',
    'Контакт', 'ГрузОт', 'ОнЖе', 'ГрузОтпр', 'ГрузПолуч', 'СвПокуп', 'ДокПодтвОтгрНом',
    'ДенИзм', 'Тлф', 'ЭлПочта', 'ИнфПолФХЖ1', 'ТекстИнф', 'ТаблСчФакт', 'СведТов',
    'ДопСведТов', 'СвДТ', 'СумНал', 'БезНДС', 'ИнфПолФХЖ2', 'КИЗ', 'КрНаимСтрПр',
    'Акциз', 'БезАкциз', 'ВсегоОпл', 'Документ', 'НомСредИдентТов', 'СумНалВсего',
    'БанкРекв', 'СвБанк'
}

SECTIONS = {
    'Подписант': './/Подписант',
    'СвПродПер': './/СвПродПер',
    'СвСчФакт': './/СвСчФакт',
    'ИнфПолФХЖ1': './/ИнфПолФХЖ1',
    'ТаблСчФакт': './/ТаблСчФакт'
}


def detect_encoding(content: bytes) -> str:
    if content.startswith(b'\xef\xbb\xbf'):
        logger.debug("Обнаружен BOM: кодировка UTF-8")
        return 'utf-8-sig'
    try:
        header = content[:100].decode('ascii', errors='ignore')
        if 'encoding="' in header:
            start = header.find('encoding="') + 10
            end = header.find('"', start)
            encoding = header[start:end].lower()
            logger.debug(f"Кодировка из заголовка XML: {encoding}")
            return encoding if encoding in ['windows-1251', 'utf-8'] else 'windows-1251'
        logger.debug("Кодировка не указана, предполагается Windows-1251")
        return 'windows-1251'
    except Exception as e:
        logger.error(f"Ошибка при определении кодировки: {e}, предполагается Windows-1251")
        return 'windows-1251'


def collect_tags(element, tag_dict=None, parent_path=""):
    if tag_dict is None:
        tag_dict = defaultdict(list)
    current_path = f"{parent_path}/{element.tag}" if parent_path else element.tag
    tag_dict[element.tag].append(current_path)
    for child in element:
        collect_tags(child, tag_dict, current_path)
    return tag_dict


def classify_tag_location(tag, paths, root):
    locations = defaultdict(int)
    for path in paths:
        for section_name, xpath in SECTIONS.items():
            section_node = root.find(xpath)
            if section_node is not None and any(section_node.find(f".//{tag}") is not None for p in paths):
                locations[section_name] += 1
                break
        else:
            locations['Другие'] += 1
    return locations


def log_new_tags(root):
    tag_dict = collect_tags(root)
    new_tags = {tag for tag in tag_dict.keys() if tag not in EXPECTED_TAGS}
    if new_tags:
        logger.info("Обнаружены новые теги:")
        for tag in sorted(new_tags):
            count = len(tag_dict[tag])
            locations = classify_tag_location(tag, tag_dict[tag], root)
            location_str = ", ".join([f"{section}: {count}" for section, count in locations.items()])
            logger.info(f"- Тег: {tag}, Количество: {count}, Местоположение: {location_str}")
    else:
        logger.info("Новых тегов не обнаружено.")

def xml_to_excel(input_file, output_file):
    podpisant_data = []
    sv_prod_per_data = []
    sv_sch_fakt_data = []
    inf_pol_fhzh1_data = []
    products = []
    total_payment = None
    total_tax = None

    logger.debug(f"Попытка парсинга XML из {input_file}...")
    root = None
    try:
        with open(input_file, 'rb') as file:
            content = file.read()
            logger.debug(f"Первые 50 байтов файла: {repr(content[:50])}")
            encoding = detect_encoding(content)

        parser = etree.XMLParser(recover=True, remove_blank_text=True)
        tree = etree.parse(input_file, parser=parser)
        root = tree.getroot()
        logger.debug(f"lxml: Корневой элемент найден: {root.tag}")
        log_new_tags(root)
    except etree.LxmlError as e:
        logger.error(f"lxml: Ошибка парсинга: {e}")
        try:
            with open(input_file, 'rb') as file:
                content = file.read()
                encoding = detect_encoding(content)
                content_str = content.decode(encoding, errors='replace')
            tree = ET.fromstring(content_str)
            root = tree
            logger.debug(f"ElementTree: Корневой элемент найден: {root.tag}")
            log_new_tags(root)
        except ET.ParseError as e:
            logger.error(f"ElementTree: Ошибка парсинга: {e}")
            logger.error("Не удалось распарсить XML ни одним методом.")
    except Exception as e:
        logger.error(f"Произошла ошибка при парсинге XML: {e}")
        raise

    if root is None:
        logger.error("Не удалось распарсить XML. Файл будет создан с пустыми данными.")
    else:
        # Документ
        doc = root.find('.//Документ')
        if doc is not None:
            sv_sch_fakt_data.append({
                'ДатаИнфПр': doc.get('ДатаИнфПр', ''),
                'ВремИнфПр': doc.get('ВремИнфПр', ''),
                'НаимЭконСубСост': doc.get('НаимЭконСубСост', '')
            })

        # Подписант
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
            logger.debug(f"Найдено записей Подписант: {len(podpisant_data)}")

        # СвПродПер
        sv_prod_per = root.find('.//СвПродПер')
        if sv_prod_per is not None:
            sv_per = sv_prod_per.find('СвПер')
            osn_per = sv_per.find('ОснПер') if sv_per is not None else None
            bez_dok = sv_per.find('БезДокОснПер') if sv_per is not None else None
            sv_prod_per_entry = {
                'ДатаПер': sv_per.get('ДатаПер', ''),
                'СодОпер': sv_per.get('СодОпер', ''),
                'РеквДатаДок': osn_per.get('РеквДатаДок', '') if osn_per is not None else '',
                'РеквНаимДок': osn_per.get('РеквНаимДок', '') if osn_per is not None else '',
                'РеквНомерДок': osn_per.get('РеквНомерДок', '') if osn_per is not None else '',
                'БезДокОснПер': bez_dok.text if bez_dok is not None else ''
            }
            sv_prod_per_data.append(sv_prod_per_entry)
            logger.debug(f"Найдено записей СвПродПер: {len(sv_prod_per_data)}")

        # СвСчФакт
        sv_sch_fakt = root.find('.//СвСчФакт')
        if sv_sch_fakt is not None:
            sv_prod = sv_sch_fakt.find('СвПрод')
            id_sv_prod = sv_prod.find('ИдСв') if sv_prod is not None else None
            sv_yul_uch_prod = id_sv_prod.find('СвЮЛУч') if id_sv_prod is not None else None
            sv_ip_prod = id_sv_prod.find('СвИП') if id_sv_prod is not None else None
            fio_prod = sv_ip_prod.find('ФИО') if sv_ip_prod is not None else None
            adres_prod = sv_prod.find('Адрес') if sv_prod is not None else None
            adr_rf_prod = adres_prod.find('АдрРФ') if adres_prod is not None else None
            adr_inf_prod = adres_prod.find('АдрИнф') if adres_prod is not None else None
            kontakt_prod = sv_prod.find('Контакт') if sv_prod is not None else None
            bank_rekv_prod = sv_prod.find('БанкРекв') if sv_prod is not None else None
            sv_bank_prod = bank_rekv_prod.find('СвБанк') if bank_rekv_prod is not None else None

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
            sv_ip_pokup = id_sv_pokup.find('СвИП') if sv_pokup is not None else None
            fio_pokup = sv_ip_pokup.find('ФИО') if sv_ip_pokup is not None else None
            adres_pokup = sv_pokup.find('Адрес') if sv_pokup is not None else None
            adr_inf_pokup = adres_pokup.find('АдрИнф') if adres_pokup is not None else None
            bank_rekv_pokup = sv_pokup.find('БанкРекв') if sv_pokup is not None else None
            sv_bank_pokup = bank_rekv_pokup.find('СвБанк') if bank_rekv_pokup is not None else None

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
                'СвГосРегИП_СвПрод': sv_ip_prod.get('СвГосРегИП', '') if sv_ip_prod is not None else '',
                'Фамилия_СвПрод': fio_prod.get('Фамилия', '') if fio_prod is not None else '',
                'Имя_СвПрод': fio_prod.get('Имя', '') if fio_prod is not None else '',
                'Отчество_СвПрод': fio_prod.get('Отчество', '') if fio_prod is not None else '',
                'Индекс_СвПрод': adr_rf_prod.get('Индекс', '') if adr_rf_prod is not None else '',
                'КодРегион_СвПрод': adr_rf_prod.get('КодРегион', '') if adr_rf_prod is not None else '',
                'НаимРегион_СвПрод': adr_rf_prod.get('НаимРегион', '') if adr_rf_prod is not None else '',
                'Город_СвПрод': adr_rf_prod.get('Город', '') if adr_rf_prod is not None else '',
                'АдрТекст_СвПрод': adr_inf_prod.get('АдрТекст', '') if adr_inf_prod is not None else '',
                'КодСтр_СвПрод': adr_inf_prod.get('КодСтр', '') if adr_inf_prod is not None else '',
                'НаимСтран_СвПрод': adr_inf_prod.get('НаимСтран', '') if adr_inf_prod is not None else '',
                'Тлф_СвПрод': kontakt_prod.find('Тлф').text if kontakt_prod is not None and kontakt_prod.find(
                    'Тлф') is not None else '',
                'ЭлПочта_СвПрод': kontakt_prod.find('ЭлПочта').text if kontakt_prod is not None and kontakt_prod.find(
                    'ЭлПочта') is not None else '',
                'НомерСчёта_СвПрод': bank_rekv_prod.get('НомерСчёта', '') if bank_rekv_prod is not None else '',
                'БИК_СвПрод': sv_bank_prod.get('БИК', '') if sv_bank_prod is not None else '',
                'НаимБанк_СвПрод': sv_bank_prod.get('НаимБанк', '') if sv_bank_prod is not None else '',
                'КорСчет_СвПрод': sv_bank_prod.get('КорСчет', '') if sv_bank_prod is not None else '',
                'ОнЖе_ГрузОт': on_zhe.text if on_zhe is not None else '',
                'ОКПО_ГрузОт': gruz_otpr.get('ОКПО', '') if gruz_otpr is not None else '',
                'ИННЮЛ_ГрузОт': sv_yul_uch_gruz_ot.get('ИННЮЛ', '') if sv_yul_uch_gruz_ot is not None else '',
                'КПП_ГрузОт': sv_yul_uch_gruz_ot.get('КПП', '') if sv_yul_uch_gruz_ot is not None else '',
                'НаимОрг_ГрузОт': sv_yul_uch_gruz_ot.get('НаимОрг', '') if sv_yul_uch_gruz_ot is not None else '',
                'АдрТекст_ГрузОт': adr_inf_gruz_ot.get('АдрТекст', '') if adr_inf_gruz_ot is not None else '',
                'КодСтр_ГрузОт': adr_inf_gruz_ot.get('КодСтр', '') if adr_inf_gruz_ot is not None else '',
                'НаимСтран_ГрузОт': adr_inf_gruz_ot.get('НаимСтран', '') if adr_inf_gruz_ot is not None else '',
                'Тлф_ГрузОт': kontakt_gruz_ot.find('Тлф').text if kontakt_gruz_ot is not None and kontakt_gruz_ot.find(
                    'Тлф') is not None else '',
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
                'НомерСчёта_СвПокуп': bank_rekv_pokup.get('НомерСчёта', '') if bank_rekv_pokup is not None else '',
                'БИК_СвПокуп': sv_bank_pokup.get('БИК', '') if sv_bank_pokup is not None else '',
                'НаимБанк_СвПокуп': sv_bank_pokup.get('НаимБанк', '') if sv_bank_pokup is not None else '',
                'КорСчет_СвПокуп': sv_bank_pokup.get('КорСчет', '') if sv_bank_pokup is not None else '',
                'РеквДатаДок_ДокПодтв': dok_podtv.get('РеквДатаДок', '') if dok_podtv is not None else '',
                'РеквНаимДок_ДокПодтв': dok_podtv.get('РеквНаимДок', '') if dok_podtv is not None else '',
                'РеквНомерДок_ДокПодтв': dok_podtv.get('РеквНомерДок', '') if dok_podtv is not None else '',
                'КодОКВ': den_izm.get('КодОКВ', '') if den_izm is not None else '',
                'НаимОКВ': den_izm.get('НаимОКВ', '') if den_izm is not None else ''
            }
            sv_sch_fakt_data.append(sv_sch_fakt_entry)
            logger.debug(f"Найдено записей СвСчФакт: {len(sv_sch_fakt_data)}")

        # ИнфПолФХЖ1
        inf_pol_fhzh1 = root.findall('.//ИнфПолФХЖ1/ТекстИнф')
        if inf_pol_fhzh1:
            for text_inf in inf_pol_fhzh1:
                inf_pol_entry = {
                    'Идентиф': text_inf.get('Идентиф', ''),
                    'Значен': text_inf.get('Значен', '')
                }
                inf_pol_fhzh1_data.append(inf_pol_entry)
            logger.debug(f"Найдено записей ИнфПолФХЖ1: {len(inf_pol_fhzh1_data)}")

        # ТаблСчФакт и товары
        tab_sch_fakt = root.find('.//ТаблСчФакт')
        if tab_sch_fakt is not None:
            total_payment = tab_sch_fakt.get('ВсегоОпл', '')
            total_tax = tab_sch_fakt.find('.//СумНалВсего/СумНал').text if tab_sch_fakt.find(
                './/СумНалВсего/СумНал') is not None else ''
            for item in root.findall('.//ТаблСчФакт/СведТов'):
                dop_sved = item.find('ДопСведТов')
                sv_dt = item.find('СвДТ')
                product = {
                    'Номер строки': item.get('НомСтр', ''),
                    'Наименование': item.get('НаимТов', ''),
                    'Количество': item.get('КолТов', ''),
                    'Ед. измерения': item.get('НаимЕдИзм', ''),
                    'Цена': item.get('ЦенаТов', ''),
                    'Стоимость без НДС': item.get('СтТовБезНДС', ''),
                    'Ставка НДС': item.get('НалСт', ''),
                    'Стоимость с НДС': item.get('СтТовУчНал', ''),
                    'ОКЕИ_Тов': item.get('ОКЕИ_Тов', ''),
                    'Код товара': dop_sved.get('КодТов', '') if dop_sved is not None else '',
                    'ГТИН': dop_sved.get('ГТИН', '') if dop_sved is not None else '',
                    'КрНаимСтрПр': dop_sved.find('КрНаимСтрПр').text if dop_sved is not None and dop_sved.find(
                        'КрНаимСтрПр') is not None else '',
                    'КодПроисх': sv_dt.get('КодПроисх', '') if sv_dt is not None else '',
                    'НомерДТ': sv_dt.get('НомерДТ', '') if sv_dt is not None else '',
                    'GTIN': '',
                    'КодПокупателя': '',
                    'НазваниеПокупателя': '',
                    'КИЗ': '',
                    'НомСредИдентТов': item.get('НомСредИдентТов', '')
                }

                sum_nal = item.find('СумНал')
                if sum_nal is not None:
                    if sum_nal.find('СумНал') is not None:
                        product['Сумма НДС'] = sum_nal.find('СумНал').text
                    elif sum_nal.find('БезНДС') is not None:
                        product['Сумма НДС'] = 'без НДС'
                else:
                    product['Сумма НДС'] = ''

                excise = item.find('Акциз')
                if excise is not None:
                    if excise.find('СумАкциз') is not None:
                        product['Сумма Акциза'] = excise.find('СумАкциз').text
                    elif excise.find('БезАкциз') is not None:
                        product['Сумма Акциза'] = 'без акциза'
                else:
                    product['Сумма Акциза'] = ''

                for inf in item.findall('.//ИнфПолФХЖ2'):
                    ident = inf.get('Идентиф', '')
                    value = inf.get('Значен', '')
                    if ident == 'GTIN':
                        product['GTIN'] = value
                    elif ident == 'КодПокупателя':
                        product['КодПокупателя'] = value
                    elif ident == 'НазваниеПокупателя':
                        product['НазваниеПокупателя'] = value

                kiz_list = [kiz.text for kiz in item.findall('.//КИЗ') if kiz.text is not None]
                product['КИЗ'] = '; '.join(kiz_list) if kiz_list else ''

                products.append(product)
            logger.debug(f"Найдено товаров: {len(products)}")

    columns_order = [
        'Номер строки', 'Наименование', 'Код товара', 'Количество', 'Ед. измерения',
        'Цена', 'Стоимость без НДС', 'Ставка НДС', 'Сумма НДС', 'Сумма Акциза', 'Стоимость с НДС',
        'ОКЕИ_Тов', 'GTIN', 'ГТИН', 'КодПокупателя', 'НазваниеПокупателя', 'КИЗ',
        'КрНаимСтрПр', 'КодПроисх', 'НомерДТ', 'НомСредИдентТов'
    ]
    try:
        if products:
            df_products = pd.DataFrame(products)
            for col in columns_order:
                if col not in df_products.columns:
                    df_products[col] = ''
            df_products = df_products[columns_order]
        else:
            df_products = pd.DataFrame(columns=columns_order)
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_products.to_excel(writer, sheet_name='Товары', index=False)
            df_podpisant = pd.DataFrame(podpisant_data if podpisant_data else [{}])
            df_podpisant.to_excel(writer, sheet_name='Подписант', index=False)
            df_sv_prod_per = pd.DataFrame(sv_prod_per_data if sv_prod_per_data else [{}])
            df_sv_prod_per.to_excel(writer, sheet_name='СвПродПер', index=False)
            df_sv_sch_fakt = pd.DataFrame(sv_sch_fakt_data if sv_sch_fakt_data else [{}])
            df_sv_sch_fakt.to_excel(writer, sheet_name='СвСчФакт', index=False)
            df_inf_pol = pd.DataFrame(inf_pol_fhzh1_data if inf_pol_fhzh1_data else [{}])
            df_inf_pol.to_excel(writer, sheet_name='ИнфПолФХЖ1', index=False)
            df_totals = pd.DataFrame({'ВсегоОпл': [total_payment], 'СумНалВсего': [total_tax]})
            df_totals.to_excel(writer, sheet_name='Итоги', index=False)
        logger.info(
            f"Таблица успешно сохранена в {output_file} с листами: Товары, Подписант, СвПродПер, СвСчФакт, ИнфПолФХЖ1, Итоги")
    except Exception as e:
        logger.error(f"Ошибка при сохранении Excel: {e}")
