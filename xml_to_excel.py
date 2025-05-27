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

STRING_FIELDS = [
    'ОКПО_СвПрод', 'ИННФЛ_СвПрод', 'Индекс_СвПрод', 'КодРегион_СвПрод',
    'ОКПО_ГрузОт', 'ОКПО_ГрузПолуч', 'ИННФЛ_ГрузПолуч', 'Индекс_ГрузПолуч',
    'КодРегион_ГрузПолуч', 'ОКПО_СвПокуп', 'ИННФЛ_СвПокуп', 'КодОКВ', 'БИК_СвПрод',
    'БИК_СвПокуп', 'НомерСчета_СвПрод', 'НомерСчета_СвПокуп'
]

SV_SCH_FAKT_COLUMNS = [
    'ДатаИнфПр', 'ВремИнфПр', 'НаимЭконСубСост', 'НаимДокОпр', 'ДатаДок', 'НомерДок',
    'ОКПО_СвПрод', 'ИННЮЛ_СвПрод', 'КПП_СвПрод', 'НаимОрг_СвПрод', 'ИННФЛ_СвПрод',
    'СвГосРегИП_СвПрод', 'Фамилия_СвПрод', 'Имя_СвПрод', 'Отчество_СвПрод', 'Адрес_СвПрод',
    'Тлф_СвПрод', 'ЭлПочта_СвПрод', 'НомерСчета_СвПрод', 'БИК_СвПрод', 'НаимБанк_СвПрод',
    'КорСчет_СвПрод', 'ОнЖе_ГрузОт', 'ОКПО_ГрузОт', 'ИННЮЛ_ГрузОт', 'КПП_ГрузОт',
    'НаимОрг_ГрузОт', 'Адрес_ГрузОт', 'Тлф_ГрузОт', 'ЭлПочта_ГрузОт', 'ОКПО_ГрузПолуч',
    'СокрНаим_ГрузПолуч', 'ИННФЛ_ГрузПолуч', 'СвГосРегИП_ГрузПолуч', 'Фамилия_ГрузПолуч',
    'Имя_ГрузПолуч', 'Отчество_ГрузПолуч', 'Адрес_ГрузПолуч', 'Тлф_ГрузПолуч', 'ЭлПочта_ГрузПолуч',
    'НомерСчета_ГрузПолуч', 'БИК_ГрузПолуч', 'НаимБанк_ГрузПолуч', 'КорСчет_ГрузПолуч',
    'ОКПО_СвПокуп', 'СокрНаим_СвПокуп', 'ИННФЛ_СвПокуп', 'СвГосРегИП_СвПокуп',
    'Фамилия_СвПокуп', 'Имя_СвПокуп', 'Отчество_СвПокуп', 'Адрес_СвПокуп', 'Тлф_СвПокуп',
    'ЭлПочта_СвПокуп', 'НомерСчета_СвПокуп', 'БИК_СвПокуп', 'НаимБанк_СвПокуп', 'КорСчет_СвПокуп',
    'РеквДатаДок_ДокПодтв', 'РеквНаимДок_ДокПодтв', 'РеквНомерДок_ДокПодтв', 'КодОКВ', 'НаимОКВ'
]

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


def format_address(adr_rf, adr_inf):
    if adr_rf is not None:
        parts = []
        for key in ['Индекс', 'КодРегион', 'НаимРегион', 'Город', 'Улица', 'Дом']:
            if adr_rf.get(key):
                parts.append(f"{key}={adr_rf.get(key)}")
        return "; ".join(parts) if parts else ""
    elif adr_inf is not None:
        parts = []
        for key in ['АдрТекст', 'КодСтр', 'НаимСтран']:
            if adr_inf.get(key):
                parts.append(f"{key}={adr_inf.get(key)}")
        return "; ".join(parts) if parts else ""
    return ""

def xml_to_excel(input_file, output_file):
    podpisant_data = []
    sv_prod_per_data = []
    sv_sch_fakt_data = {col: "" for col in SV_SCH_FAKT_COLUMNS}
    inf_pol_fhzh1_data = []
    products = []
    total_payment = None
    total_tax = None
    file_data = {}

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
        file_data = {
            'ИдФайл': root.get('ИдФайл', ''),
            'ВерсФорм': root.get('ВерсФорм', ''),
            'ВерсПрог': root.get('ВерсПрог', '')
        }

        doc = root.find('.//Документ')
        if doc is not None:
            sv_sch_fakt_data.update({
                'ДатаИнфПр': doc.get('ДатаИнфПр', ''),
                'ВремИнфПр': doc.get('ВремИнфПр', ''),
                'НаимЭконСубСост': doc.get('НаимЭконСубСост', ''),
                'НаимДокОпр': doc.get('НаимДокОпр', '')
            })

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
            kontakt_gruz_poluch = gruz_poluch.find('Контакт') if gruz_poluch is not None else None
            bank_rekv_gruz_poluch = gruz_poluch.find('БанкРекв') if gruz_poluch is not None else None
            sv_bank_gruz_poluch = bank_rekv_gruz_poluch.find('СвБанк') if bank_rekv_gruz_poluch is not None else None

            sv_pokup = sv_sch_fakt.find('СвПокуп')
            id_sv_pokup = sv_pokup.find('ИдСв') if sv_pokup is not None else None
            sv_ip_pokup = id_sv_pokup.find('СвИП') if id_sv_pokup is not None else None
            fio_pokup = sv_ip_pokup.find('ФИО') if sv_ip_pokup is not None else None
            adres_pokup = sv_pokup.find('Адрес') if sv_pokup is not None else None
            adr_rf_pokup = adres_pokup.find('АдрРФ') if adres_pokup is not None else None
            adr_inf_pokup = adres_pokup.find('АдрИнф') if adres_pokup is not None else None
            kontakt_pokup = sv_pokup.find('Контакт') if sv_pokup is not None else None
            bank_rekv_pokup = sv_pokup.find('БанкРекв') if sv_pokup is not None else None
            sv_bank_pokup = bank_rekv_pokup.find('СвБанк') if bank_rekv_pokup is not None else None

            dok_podtv = sv_sch_fakt.find('ДокПодтвОтгрНом')
            den_izm = sv_sch_fakt.find('ДенИзм')

            # Обновление данных с явной проверкой и преобразованием в скалярные значения
            sv_sch_fakt_data.update({
                'ДатаДок': str(sv_sch_fakt.get('ДатаДок', '')),
                'НомерДок': str(sv_sch_fakt.get('НомерДок', '')),
                'ОКПО_СвПрод': str(sv_prod.get('ОКПО', '')) if sv_prod is not None else '',
                'ИННЮЛ_СвПрод': str(sv_yul_uch_prod.get('ИННЮЛ', '')) if sv_yul_uch_prod is not None else '',
                'КПП_СвПрод': str(sv_yul_uch_prod.get('КПП', '')) if sv_yul_uch_prod is not None else '',
                'НаимОрг_СвПрод': str(sv_yul_uch_prod.get('НаимОрг', '')) if sv_yul_uch_prod is not None else '',
                'ИННФЛ_СвПрод': str(sv_ip_prod.get('ИННФЛ', '')) if sv_ip_prod is not None else '',
                'СвГосРегИП_СвПрод': str(sv_ip_prod.get('СвГосРегИП', '')) if sv_ip_prod is not None else '',
                'Фамилия_СвПрод': str(fio_prod.get('Фамилия', '')) if fio_prod is not None else '',
                'Имя_СвПрод': str(fio_prod.get('Имя', '')) if fio_prod is not None else '',
                'Отчество_СвПрод': str(fio_prod.get('Отчество', '')) if fio_prod is not None else '',
                'Адрес_СвПрод': format_address(adr_rf_prod, adr_inf_prod),
                'Тлф_СвПрод': str(kontakt_prod.find('Тлф').text) if kontakt_prod is not None and kontakt_prod.find(
                    'Тлф') is not None else '',
                'ЭлПочта_СвПрод': str(
                    kontakt_prod.find('ЭлПочта').text) if kontakt_prod is not None and kontakt_prod.find(
                    'ЭлПочта') is not None else '',
                'НомерСчета_СвПрод': str(bank_rekv_prod.get('НомерСчета', '')) if bank_rekv_prod is not None else '',
                'БИК_СвПрод': str(sv_bank_prod.get('БИК', '')) if sv_bank_prod is not None else '',
                'НаимБанк_СвПрод': str(sv_bank_prod.get('НаимБанк', '')) if sv_bank_prod is not None else '',
                'КорСчет_СвПрод': str(sv_bank_prod.get('КорСчет', '')) if sv_bank_prod is not None else '',
                'ОнЖе_ГрузОт': str(on_zhe.text) if on_zhe is not None else '',
                'ОКПО_ГрузОт': str(gruz_otpr.get('ОКПО', '')) if gruz_otpr is not None else '',
                'ИННЮЛ_ГрузОт': str(sv_yul_uch_gruz_ot.get('ИННЮЛ', '')) if sv_yul_uch_gruz_ot is not None else '',
                'КПП_ГрузОт': str(sv_yul_uch_gruz_ot.get('КПП', '')) if sv_yul_uch_gruz_ot is not None else '',
                'НаимОрг_ГрузОт': str(sv_yul_uch_gruz_ot.get('НаимОрг', '')) if sv_yul_uch_gruz_ot is not None else '',
                'Адрес_ГрузОт': format_address(None, adr_inf_gruz_ot),
                'Тлф_ГрузОт': str(
                    kontakt_gruz_ot.find('Тлф').text) if kontakt_gruz_ot is not None and kontakt_gruz_ot.find(
                    'Тлф') is not None else '',
                'ЭлПочта_ГрузОт': str(
                    kontakt_gruz_ot.find('ЭлПочта').text) if kontakt_gruz_ot is not None and kontakt_gruz_ot.find(
                    'ЭлПочта') is not None else '',
                'ОКПО_ГрузПолуч': str(gruz_poluch.get('ОКПО', '')) if gruz_poluch is not None else '',
                'СокрНаим_ГрузПолуч': str(gruz_poluch.get('СокрНаим', '')) if gruz_poluch is not None else '',
                'ИННФЛ_ГрузПолуч': str(sv_ip_gruz_poluch.get('ИННФЛ', '')) if sv_ip_gruz_poluch is not None else '',
                'СвГосРегИП_ГрузПолуч': str(
                    sv_ip_gruz_poluch.get('СвГосРегИП', '')) if sv_ip_gruz_poluch is not None else '',
                'Фамилия_ГрузПолуч': str(fio_gruz_poluch.get('Фамилия', '')) if fio_gruz_poluch is not None else '',
                'Имя_ГрузПолуч': str(fio_gruz_poluch.get('Имя', '')) if fio_gruz_poluch is not None else '',
                'Отчество_ГрузПолуч': str(fio_gruz_poluch.get('Отчество', '')) if fio_gruz_poluch is not None else '',
                'Адрес_ГрузПолуч': format_address(adr_rf_gruz_poluch, None),
                'Тлф_ГрузПолуч': str(kontakt_gruz_poluch.find(
                    'Тлф').text) if kontakt_gruz_poluch is not None and kontakt_gruz_poluch.find(
                    'Тлф') is not None else '',
                'ЭлПочта_ГрузПолуч': str(kontakt_gruz_poluch.find(
                    'ЭлПочта').text) if kontakt_gruz_poluch is not None and kontakt_gruz_poluch.find(
                    'ЭлПочта') is not None else '',
                'НомерСчета_ГрузПолуч': str(
                    bank_rekv_gruz_poluch.get('НомерСчета', '')) if bank_rekv_gruz_poluch is not None else '',
                'БИК_ГрузПолуч': str(sv_bank_gruz_poluch.get('БИК', '')) if sv_bank_gruz_poluch is not None else '',
                'НаимБанк_ГрузПолуч': str(
                    sv_bank_gruz_poluch.get('НаимБанк', '')) if sv_bank_gruz_poluch is not None else '',
                'КорСчет_ГрузПолуч': str(
                    sv_bank_gruz_poluch.get('КорСчет', '')) if sv_bank_gruz_poluch is not None else '',
                'ОКПО_СвПокуп': str(sv_pokup.get('ОКПО', '')) if sv_pokup is not None else '',
                'СокрНаим_СвПокуп': str(sv_pokup.get('СокрНаим', '')) if sv_pokup is not None else '',
                'ИННФЛ_СвПокуп': str(sv_ip_pokup.get('ИННФЛ', '')) if sv_ip_pokup is not None else '',
                'СвГосРегИП_СвПокуп': str(sv_ip_pokup.get('СвГосРегИП', '')) if sv_ip_pokup is not None else '',
                'Фамилия_СвПокуп': str(fio_pokup.get('Фамилия', '')) if fio_pokup is not None else '',
                'Имя_СвПокуп': str(fio_pokup.get('Имя', '')) if fio_pokup is not None else '',
                'Отчество_СвПокуп': str(fio_pokup.get('Отчество', '')) if fio_pokup is not None else '',
                'Адрес_СвПокуп': format_address(adr_rf_pokup, adr_inf_pokup),
                'Тлф_СвПокуп': str(kontakt_pokup.find('Тлф').text) if kontakt_pokup is not None and kontakt_pokup.find(
                    'Тлф') is not None else '',
                'ЭлПочта_СвПокуп': str(
                    kontakt_pokup.find('ЭлПочта').text) if kontakt_pokup is not None and kontakt_pokup.find(
                    'ЭлПочта') is not None else '',
                'НомерСчета_СвПокуп': str(bank_rekv_pokup.get('НомерСчета', '')) if bank_rekv_pokup is not None else '',
                'БИК_СвПокуп': str(sv_bank_pokup.get('БИК', '')) if sv_bank_pokup is not None else '',
                'НаимБанк_СвПокуп': str(sv_bank_pokup.get('НаимБанк', '')) if sv_bank_pokup is not None else '',
                'КорСчет_СвПокуп': str(sv_bank_pokup.get('КорСчет', '')) if sv_bank_pokup is not None else '',
                'РеквДатаДок_ДокПодтв': str(dok_podtv.get('РеквДатаДок', '')) if dok_podtv is not None else '',
                'РеквНаимДок_ДокПодтв': str(dok_podtv.get('РеквНаимДок', '')) if dok_podtv is not None else '',
                'РеквНомерДок_ДокПодтв': str(dok_podtv.get('РеквНомерДок', '')) if dok_podtv is not None else '',
                'КодОКВ': str(den_izm.get('КодОКВ', '')) if den_izm is not None else '',
                'НаимОКВ': str(den_izm.get('НаимОКВ', '')) if den_izm is not None else ''
            })

            # Преобразование всех значений в строки для обеспечения единообразия
            for key in sv_sch_fakt_data:
                sv_sch_fakt_data[key] = str(sv_sch_fakt_data[key]) if sv_sch_fakt_data[key] is not None else ''

            # Применение логики для строковых полей
            for key, value in sv_sch_fakt_data.items():
                if key in STRING_FIELDS and value:
                    try:
                        sv_sch_fakt_data[key] = str(int(float(value)))
                    except (ValueError, TypeError):
                        sv_sch_fakt_data[key] = str(value)

            if sv_sch_fakt_data['КодОКВ']:
                sv_sch_fakt_data['КодОКВ'] = '643'
                sv_sch_fakt_data['НаимОКВ'] = 'Российский рубль'

            logger.debug(f"Данные СвСчФакт после обработки: {sv_sch_fakt_data}")

        inf_pol_fhzh1 = root.findall('.//ИнфПолФХЖ1/ТекстИнф')
        if inf_pol_fhzh1:
            for text_inf in inf_pol_fhzh1:
                inf_pol_entry = {
                    'Идентиф': text_inf.get('Идентиф', ''),
                    'Значен': text_inf.get('Значен', '')
                }
                inf_pol_fhzh1_data.append(inf_pol_entry)
            logger.debug(f"Найдено записей ИнфПолФХЖ1: {len(inf_pol_fhzh1_data)}")

        tab_sch_fakt = root.find('.//ТаблСчФакт')
        total_with_tax = 0
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
                    'АртикулТов': dop_sved.get('АртикулТов', '') if dop_sved is not None else '',
                    'Код товара': dop_sved.get('КодТов', '') if dop_sved is not None else '',
                    'КодВидТов': dop_sved.get('КодВидТов', '') if dop_sved is not None else '',
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

                try:
                    total_with_tax += float(item.get('СтТовУчНал', 0))
                except (ValueError, TypeError):
                    pass

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

                for key, value in product.items():
                    if isinstance(value, str) and value.lower() == 'nan':
                        product[key] = ''

                products.append(product)
            logger.debug(f"Найдено товаров: {len(products)}")

    columns_order = [
        'Номер строки', 'Наименование', 'АртикулТов', 'Код товара', 'КодВидТов', 'Количество', 'Ед. измерения',
        'Цена', 'Стоимость без НДС', 'Ставка НДС', 'Сумма НДС', 'Сумма Акциза', 'Стоимость с НДС',
        'ОКЕИ_Тов', 'GTIN', 'ГТИН', 'КодПокупателя', 'НазваниеПокупателя', 'КИЗ',
        'КрНаимСтрPr', 'КодПроисх', 'НомерДТ', 'НомСредИдентТов'
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
            df_sv_sch_fakt = pd.DataFrame([sv_sch_fakt_data])
            df_sv_sch_fakt = df_sv_sch_fakt[SV_SCH_FAKT_COLUMNS]
            df_sv_sch_fakt.to_excel(writer, sheet_name='СвСчФакт', index=False)
            df_inf_pol = pd.DataFrame(inf_pol_fhzh1_data if inf_pol_fhzh1_data else [{}])
            df_inf_pol.to_excel(writer, sheet_name='ИнфПолФХЖ1', index=False)
            df_file = pd.DataFrame([file_data])
            df_file.to_excel(writer, sheet_name='Файл', index=False)
            # Исправление: унифицируем длину списков
            data = {}
            if total_payment or total_tax or total_with_tax:
                data['ВсегоОпл'] = [str(total_payment)] if total_payment else ['']
                data['СумНалВсего'] = [str(total_tax)] if total_tax else ['']
                data['СтТовУчНалВсего'] = [str(total_with_tax)] if total_with_tax else ['']
            else:
                data = {'ВсегоОпл': [''], 'СумНалВсего': [''], 'СтТовУчНалВсего': ['']}
            df_totals = pd.DataFrame(data)
            df_totals.to_excel(writer, sheet_name='Итоги', index=False)
        logger.info(
            f"Таблица успешно сохранена в {output_file} с листами: Товары, Подписант, СвПродПер, СвСчФакт, ИнфПолФХЖ1, Файл, Итоги")
    except Exception as e:
        logger.error(f"Ошибка при сохранении Excel: {e}")
        raise
