from pathlib import Path

import pandas as pd
from lxml import etree

from utils import setup_logging

logger = setup_logging()


def parse_address(address_str):
    if not address_str:
        return None, None
    parts = dict(part.split('=') for part in address_str.split('; ') if '=' in part)
    if 'Индекс' in parts or 'КодРегион' in parts:
        return 'АдрРФ', parts
    elif 'АдрТекст' in parts or 'КодСтр' in parts:
        return 'АдрИнф', parts
    return None, None


def excel_to_xml(input_file, output_file):
    output_path = Path(output_file)
    output_dir = output_path.parent
    if not output_dir.exists():
        logger.info(f"Папка {output_dir} не найдена. Создаю...")
        output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"Папка {output_dir} создана.")

    try:
        df = pd.read_excel(input_file, sheet_name='Товары', dtype=str)
        logger.debug(f"Успешно прочитан лист Товары из {input_file}")
    except FileNotFoundError:
        logger.error(f"Ошибка: файл {input_file} не найден.")
        return
    except Exception as e:
        logger.error(f"Ошибка при чтении листа Товары из Excel: {e}")
        df = pd.DataFrame()

    try:
        podpisant_df = pd.read_excel(input_file, sheet_name='Подписант', dtype=str)
        podpisant = podpisant_df.iloc[0].to_dict() if not podpisant_df.empty else {}
        logger.debug(f"Найдено записей Подписант: {len(podpisant_df)}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа Подписант: {e}")
        podpisant = {}

    try:
        sv_prod_per_df = pd.read_excel(input_file, sheet_name='СвПродПер', dtype=str)
        sv_prod_per = sv_prod_per_df.iloc[0].to_dict() if not sv_prod_per_df.empty else {}
        logger.debug(f"Найдено записей СвПродПер: {len(sv_prod_per_df)}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа СвПродПер: {e}")
        sv_prod_per = {}

    try:
        sv_sch_fakt_df = pd.read_excel(input_file, sheet_name='СвСчФакт', dtype=str)
        sv_sch_fakt = sv_sch_fakt_df.iloc[0].to_dict() if not sv_sch_fakt_df.empty else {}
        logger.debug(f"Найдено записей СвСчФакт: {len(sv_sch_fakt_df)}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа СвСчФакт: {e}")
        sv_sch_fakt = {}

    try:
        inf_pol_fhzh1_df = pd.read_excel(input_file, sheet_name='ИнфПолФХЖ1', dtype=str)
        inf_pol_fhzh1 = inf_pol_fhzh1_df.to_dict('records') if not inf_pol_fhzh1_df.empty else []
        logger.debug(f"Найдено записей ИнфПолФХЖ1: {len(inf_pol_fhzh1)}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа ИнфПолФХЖ1: {e}")
        inf_pol_fhzh1 = []

    try:
        file_df = pd.read_excel(input_file, sheet_name='Файл', dtype=str)
        file_data = file_df.iloc[0].to_dict() if not file_df.empty else {}
        logger.debug(f"Найдены данные Файл: {file_data}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа Файл: {e}")
        file_data = {}

    try:
        totals_df = pd.read_excel(input_file, sheet_name='Итоги', dtype=str)
        totals = totals_df.iloc[0].to_dict() if not totals_df.empty else {}
        logger.debug(f"Найдены данные Итоги: {totals}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа Итоги: {e}")
        totals = {}

    sv_sch_fakt['КодОКВ'] = '643'
    sv_sch_fakt['НаимОКВ'] = 'Российский рубль'

    root = etree.Element("Файл",
                         ВерсПрог=str(file_data.get('ВерсПрог', '')),
                         ВерсФорм=str(file_data.get('ВерсФорм', '')),
                         ИдФайл=str(file_data.get('ИдФайл', '')))
    logger.debug("Создан корневой элемент XML: Файл")

    doc = etree.SubElement(root, "Документ", КНД="1115131",
                           Функция="СЧФДОП",
                           ПоФактХЖ="Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (документ об оказании услуг)",
                           НаимДокОпр=sv_sch_fakt.get('НаимДокОпр', ''),
                           ДатаИнфПр=sv_sch_fakt.get('ДатаИнфПр', ''),
                           ВремИнфПр=sv_sch_fakt.get('ВремИнфПр', ''),
                           НаимЭконСубСост=sv_sch_fakt.get('НаимЭконСубСост', ''))
    logger.debug("Создан элемент Документ")

    sv_sch_fakt_elem = etree.SubElement(doc, "СвСчФакт",
                                        ДатаДок=str(sv_sch_fakt.get('ДатаДок', '')),
                                        НомерДок=str(sv_sch_fakt.get('НомерДок', '')))

    sv_prod = etree.SubElement(sv_sch_fakt_elem, "СвПрод", ОКПО=str(sv_sch_fakt.get('ОКПО_СвПрод', '')))
    id_sv = etree.SubElement(sv_prod, "ИдСв")
    if sv_sch_fakt.get('ИННФЛ_СвПрод'):
        sv_ip = etree.SubElement(id_sv, "СвИП",
                                 ИННФЛ=str(sv_sch_fakt.get('ИННФЛ_СвПрод', '')),
                                 СвГосРегИП=str(sv_sch_fakt.get('СвГосРегИП_СвПрод', '')))
        etree.SubElement(sv_ip, "ФИО",
                         Фамилия=str(sv_sch_fakt.get('Фамилия_СвПрод', '')),
                         Имя=str(sv_sch_fakt.get('Имя_СвПрод', '')),
                         Отчество=str(sv_sch_fakt.get('Отчество_СвПрод', '')))
    adres = etree.SubElement(sv_prod, "Адрес")
    adres_type, adres_parts = parse_address(sv_sch_fakt.get('Адрес_СвПрод', ''))
    if adres_type == 'АдрРФ':
        etree.SubElement(adres, "АдрРФ",
                         Индекс=str(adres_parts.get('Индекс', '')),
                         КодРегион=str(adres_parts.get('КодРегион', '')),
                         НаимРегион=str(adres_parts.get('НаимРегион', '')),
                         Город=str(adres_parts.get('Город', '')))
    elif adres_type == 'АдрИнф':
        etree.SubElement(adres, "АдрИнф",
                         АдрТекст=str(adres_parts.get('АдрТекст', '')),
                         КодСтр=str(adres_parts.get('КодСтр', '')),
                         НаимСтран=str(adres_parts.get('НаимСтран', '')))
    if sv_sch_fakt.get('НомерСчета_СвПрод') or sv_sch_fakt.get('БИК_СвПрод') or sv_sch_fakt.get('НаимБанк_СвПрод'):
        bank_rekv = etree.SubElement(sv_prod, "БанкРекв", НомерСчета=str(sv_sch_fakt.get('НомерСчета_СвПрод', '')))
        etree.SubElement(bank_rekv, "СвБанк",
                         БИК=str(sv_sch_fakt.get('БИК_СвПрод', '')),
                         НаимБанк=str(sv_sch_fakt.get('НаимБанк_СвПрод', '')),
                         КорСчет=str(sv_sch_fakt.get('КорСчет_СвПрод', '')))
    if sv_sch_fakt.get('Тлф_СвПрод') or sv_sch_fakt.get('ЭлПочта_СвПрод'):
        kontakt = etree.SubElement(sv_prod, "Контакт")
        if sv_sch_fakt.get('Тлф_СвПрод'):
            etree.SubElement(kontakt, "Тлф").text = str(sv_sch_fakt.get('Тлф_СвПрод', ''))
        if sv_sch_fakt.get('ЭлПочта_СвПрод'):
            etree.SubElement(kontakt, "ЭлПочта").text = str(sv_sch_fakt.get('ЭлПочта_СвПрод', ''))

    gruz_ot = etree.SubElement(sv_sch_fakt_elem, "ГрузОт")
    if sv_sch_fakt.get('ОнЖе_ГрузОт'):
        etree.SubElement(gruz_ot, "ОнЖе").text = str(sv_sch_fakt.get('ОнЖе_ГрузОт', ''))
    else:
        gruz_otpr = etree.SubElement(gruz_ot, "ГрузОтпр", ОКПО=str(sv_sch_fakt.get('ОКПО_ГрузОт', '')))
        id_sv = etree.SubElement(gruz_otpr, "ИдСв")
        sv_yul_uch = etree.SubElement(id_sv, "СвЮЛУч",
                                      ИННЮЛ=str(sv_sch_fakt.get('ИННЮЛ_ГрузОт', '')),
                                      КПП=str(sv_sch_fakt.get('КПП_ГрузОт', '')),
                                      НаимОрг=str(sv_sch_fakt.get('НаимОрг_ГрузОт', '')))
        adres = etree.SubElement(gruz_otpr, "Адрес")
        adres_type, adres_parts = parse_address(sv_sch_fakt.get('Адрес_ГрузОт', ''))
        if adres_type == 'АдрИнф':
            etree.SubElement(adres, "АдрИнф",
                             АдрТекст=str(adres_parts.get('АдрТекст', '')),
                             КодСтр=str(adres_parts.get('КодСтр', '')),
                             НаимСтран=str(adres_parts.get('НаимСтран', '')))
        if sv_sch_fakt.get('НомерСчета_ГрузОт') or sv_sch_fakt.get('БИК_ГрузОт') or sv_sch_fakt.get('НаимБанк_ГрузОт'):
            bank_rekv = etree.SubElement(gruz_otpr, "БанкРекв",
                                         НомерСчета=str(sv_sch_fakt.get('НомерСчета_ГрузОт', '')))
            etree.SubElement(bank_rekv, "СвБанк",
                             БИК=str(sv_sch_fakt.get('БИК_ГрузОт', '')),
                             НаимБанк=str(sv_sch_fakt.get('НаимБанк_ГрузОт', '')),
                             КорСчет=str(sv_sch_fakt.get('КорСчет_ГрузОт', '')))
        kontakt = etree.SubElement(gruz_otpr, "Контакт")
        etree.SubElement(kontakt, "Тлф").text = str(sv_sch_fakt.get('Тлф_ГрузОт', ''))
        etree.SubElement(kontakt, "ЭлПочта").text = str(sv_sch_fakt.get('ЭлПочта_ГрузОт', ''))

    if sv_sch_fakt.get('ОКПО_ГрузПолуч'):
        gruz_poluch_attrs = {'ОКПО': str(sv_sch_fakt.get('ОКПО_ГрузПолуч', ''))}
        if sv_sch_fakt.get('СокрНаим_ГрузПолуч') and sv_sch_fakt.get('СокрНаим_ГрузПолуч') != 'nan':
            gruz_poluch_attrs['СокрНаим'] = str(sv_sch_fakt.get('СокрНаим_ГрузПолуч', ''))
        gruz_poluch = etree.SubElement(sv_sch_fakt_elem, "ГрузПолуч", **gruz_poluch_attrs)
        id_sv = etree.SubElement(gruz_poluch, "ИдСв")
        sv_ip = etree.SubElement(id_sv, "СвИП",
                                 ИННФЛ=str(sv_sch_fakt.get('ИННФЛ_ГрузПолуч', '')),
                                 СвГосРегИП=str(sv_sch_fakt.get('СвГосРегИП_ГрузПолуч', '')))
        fio = etree.SubElement(sv_ip, "ФИО",
                               Имя=str(sv_sch_fakt.get('Имя_ГрузПолуч', '')),
                               Отчество=str(sv_sch_fakt.get('Отчество_ГрузПолуч', '')),
                               Фамилия=str(sv_sch_fakt.get('Фамилия_ГрузПолуч', '')))
        adres = etree.SubElement(gruz_poluch, "Адрес")
        adres_type, adres_parts = parse_address(sv_sch_fakt.get('Адрес_ГрузПолуч', ''))
        if adres_type == 'АдрРФ':
            adr_rf_attrs = {
                'Индекс': str(adres_parts.get('Индекс', '')),
                'КодРегион': str(adres_parts.get('КодРегион', '')),
                'НаимРегион': str(adres_parts.get('НаимРегион', ''))
            }
            if adres_parts.get('Город'):
                adr_rf_attrs['Город'] = str(adres_parts.get('Город', ''))
            if adres_parts.get('Улица'):
                adr_rf_attrs['Улица'] = str(adres_parts.get('Улица', ''))
            if adres_parts.get('Дом'):
                adr_rf_attrs['Дом'] = str(adres_parts.get('Дом', ''))
            etree.SubElement(adres, "АдрРФ", **adr_rf_attrs)
        if sv_sch_fakt.get('НомерСчета_ГрузПолуч') or sv_sch_fakt.get('БИК_ГрузПолуч') or sv_sch_fakt.get(
                'НаимБанк_ГрузПолуч'):
            bank_rekv = etree.SubElement(gruz_poluch, "БанкРекв",
                                         НомерСчета=str(sv_sch_fakt.get('НомерСчета_ГрузПолуч', '')))
            etree.SubElement(bank_rekv, "СвБанк",
                             БИК=str(sv_sch_fakt.get('БИК_ГрузПолуч', '')),
                             НаимБанк=str(sv_sch_fakt.get('НаимБанк_ГрузПолуч', '')),
                             КорСчет=str(sv_sch_fakt.get('КорСчет_ГрузПолуч', '')))
        if sv_sch_fakt.get('Тлф_ГрузПолуч') or sv_sch_fakt.get('ЭлПочта_ГрузПолуч'):
            kontakt = etree.SubElement(gruz_poluch, "Контакт")
            if sv_sch_fakt.get('Тлф_ГрузПолуч'):
                etree.SubElement(kontakt, "Тлф").text = str(sv_sch_fakt.get('Тлф_ГрузПолуч', ''))
            if sv_sch_fakt.get('ЭлПочта_ГрузПолуч'):
                etree.SubElement(kontakt, "ЭлПочта").text = str(sv_sch_fakt.get('ЭлПочта_ГрузПолуч', ''))

    etree.SubElement(sv_sch_fakt_elem, "ДокПодтвОтгрНом",
                     РеквДатаДок=str(sv_sch_fakt.get('РеквДатаДок_ДокПодтв', '')),
                     РеквНаимДок=str(sv_sch_fakt.get('РеквНаимДок_ДокПодтв', '')),
                     РеквНомерДок=str(sv_sch_fakt.get('РеквНомерДок_ДокПодтв', '')))

    sv_pokup_attrs = {'ОКПО': str(sv_sch_fakt.get('ОКПО_СвПокуп', ''))}
    if sv_sch_fakt.get('СокрНаим_СвПокуп') and sv_sch_fakt.get('СокрНаим_СвПокуп') != 'nan':
        sv_pokup_attrs['СокрНаим'] = str(sv_sch_fakt.get('СокрНаим_СвПокуп', ''))
    sv_pokup = etree.SubElement(sv_sch_fakt_elem, "СвПокуп", **sv_pokup_attrs)
    id_sv = etree.SubElement(sv_pokup, "ИдСв")
    sv_ip = etree.SubElement(id_sv, "СвИП",
                             ИННФЛ=str(sv_sch_fakt.get('ИННФЛ_СвПокуп', '')),
                             СвГосРегИП=str(sv_sch_fakt.get('СвГосРегИП_СвПокуп', '')))
    fio = etree.SubElement(sv_ip, "ФИО",
                           Имя=str(sv_sch_fakt.get('Имя_СвПокуп', '')),
                           Отчество=str(sv_sch_fakt.get('Отчество_СвПокуп', '')),
                           Фамилия=str(sv_sch_fakt.get('Фамилия_СвПокуп', '')))
    adres = etree.SubElement(sv_pokup, "Адрес")
    adres_type, adres_parts = parse_address(sv_sch_fakt.get('Адрес_СвПокуп', ''))
    if adres_type == 'АдрРФ':
        etree.SubElement(adres, "АдрРФ",
                         Индекс=str(adres_parts.get('Индекс', '')),
                         КодРегион=str(adres_parts.get('КодРегион', '')),
                         НаимРегион=str(adres_parts.get('НаимРегион', '')))
    elif adres_type == 'АдрИнф':
        etree.SubElement(adres, "АдрИнф",
                         АдрТекст=str(adres_parts.get('АдрТекст', '')),
                         КодСтр=str(adres_parts.get('КодСтр', '')),
                         НаимСтран=str(adres_parts.get('НаимСтран', '')))
    if sv_sch_fakt.get('НомерСчета_СвПокуп') or sv_sch_fakt.get('БИК_СвПокуп') or sv_sch_fakt.get('НаимБанк_СвПокуп'):
        bank_rekv = etree.SubElement(sv_pokup, "БанкРекв",
                                     НомерСчета=str(sv_sch_fakt.get('НомерСчета_СвПокуп', '')))
        etree.SubElement(bank_rekv, "СвБанк",
                         БИК=str(sv_sch_fakt.get('БИК_СвПокуп', '')),
                         НаимБанк=str(sv_sch_fakt.get('НаимБанк_СвПокуп', '')),
                         КорСчет=str(sv_sch_fakt.get('КорСчет_СвПокуп', '')))
    if sv_sch_fakt.get('Тлф_СвПокуп') or sv_sch_fakt.get('ЭлПочта_СвПокуп'):
        kontakt = etree.SubElement(sv_pokup, "Контакт")
        if sv_sch_fakt.get('Тлф_СвПокуп'):
            etree.SubElement(kontakt, "Тлф").text = str(sv_sch_fakt.get('Тлф_СвПокуп', ''))
        if sv_sch_fakt.get('ЭлПочта_СвПокуп'):
            etree.SubElement(kontakt, "ЭлПочта").text = str(sv_sch_fakt.get('ЭлПочта_СвПокуп', ''))

    etree.SubElement(sv_sch_fakt_elem, "ДенИзм",
                     КодОКВ=str(sv_sch_fakt.get('КодОКВ', '')),
                     НаимОКВ=str(sv_sch_fakt.get('НаимОКВ', '')))

    tabl = etree.SubElement(doc, "ТаблСчФакт")

    if not df.empty:
        for _, row in df.iterrows():
            if not any(pd.notna(row[col]) for col in ['Номер строки', 'Наименование', 'Количество']):
                continue

            row = row.replace('nan', '')

            sved_tov_attrs = {
                'НомСтр': str(row['Номер строки']) if pd.notna(row['Номер строки']) else "",
                'НаимТов': str(row['Наименование']) if pd.notna(row['Наименование']) else "",
                'КолТов': str(row['Количество']) if pd.notna(row['Количество']) else "",
                'НаимЕдИзм': str(row['Ед. измерения']) if pd.notna(row['Ед. измерения']) else "",
                'ЦенаТов': str(row['Цена']) if pd.notna(row['Цена']) else "",
                'СтТовБезНДС': str(row['Стоимость без НДС']) if pd.notna(row['Стоимость без НДС']) else "",
                'НалСт': str(row['Ставка НДС']) if pd.notna(row['Ставка НДС']) else "",
                'СтТовУчНал': str(row['Стоимость с НДС']) if pd.notna(row['Стоимость с НДС']) else ""
            }
            if pd.notna(row.get('ОКЕИ_Тов')):
                sved_tov_attrs['ОКЕИ_Тов'] = str(row['ОКЕИ_Тов'])
            if pd.notna(row.get('НомСредИдентТов')):
                sved_tov_attrs['НомСредИдентТов'] = str(row['НомСредИдентТов'])

            sved_tov = etree.SubElement(tabl, "СведТов", **sved_tov_attrs)

            if pd.notna(row.get('КодПроисх')) or pd.notna(row.get('НомерДТ')):
                etree.SubElement(sved_tov, "СвДТ",
                                 КодПроисх=str(row['КодПроисх']) if pd.notna(row.get('КодПроисх')) else "",
                                 НомерДТ=str(row['НомерДТ']) if pd.notna(row.get('НомерДТ')) else "")

            dop_sved_attrs = {
                'КодТов': str(row['Код товара']) if pd.notna(row.get('Код товара')) else "",
                'ПрТовРаб': "1"
            }
            if pd.notna(row.get('АртикулТов')):
                dop_sved_attrs['АртикулТов'] = str(row['АртикулТов'])
            if pd.notna(row.get('КодВидТов')):
                dop_sved_attrs['КодВидТов'] = str(row['КодВидТов'])
            if pd.notna(row.get('ГТИН')):
                dop_sved_attrs['ГТИН'] = str(row['ГТИН'])

            dop_sved = etree.SubElement(sved_tov, "ДопСведТов", **dop_sved_attrs)

            if pd.notna(row.get('КрНаимСтрPr')):
                etree.SubElement(dop_sved, "КрNaимСтрPr").text = str(row['КрНаимСтрPr'])

            if pd.notna(row.get('КИЗ')) and row['КИЗ']:
                nom_sred = etree.SubElement(dop_sved, "НомСредИдентТов")
                for kiz in row['КИЗ'].split('; '):
                    etree.SubElement(nom_sred, "КИЗ").text = kiz

            if pd.notna(row.get('Сумма Акциза')):
                akciz = etree.SubElement(sved_tov, "Акциз")
                if str(row['Сумма Акциза']).lower() == 'без акциза':
                    etree.SubElement(akciz, "БезАкциз").text = "без акциза"
                else:
                    etree.SubElement(akciz, "СумАкциз").text = str(row['Сумма Акциза'])

            sum_nal = etree.SubElement(sved_tov, "СумНал")
            sum_nal_value = str(row['Сумма НДС']) if pd.notna(row['Сумма НДС']) else ""
            if sum_nal_value.lower() in ['без ндс', 'безндс']:
                etree.SubElement(sum_nal, "БезНДС").text = "без НДС"
            elif sum_nal_value:
                etree.SubElement(sum_nal, "СумНал").text = sum_nal_value

            if pd.notna(row.get('КодПокупателя')) and row['КодПокупателя']:
                etree.SubElement(sved_tov, "ИнфПолФХЖ2", Идентиф="КодПокупателя", Значен=str(row['КодПокупателя']))
            if pd.notna(row.get('НазваниеПокупателя')) and row['НазваниеПокупателя']:
                etree.SubElement(sved_tov, "ИнфПолФХЖ2", Идентиф="НазваниеПокупателя",
                                 Значен=str(row['НазваниеПокупателя']))
            if pd.notna(row.get('GTIN')) and row['GTIN'] and not pd.notna(row.get('ГТИН')):
                etree.SubElement(sved_tov, "ИнфПолФХЖ2", Идентиф="GTIN", Значен=str(row['GTIN']))

    # Создание блока ВсегоОпл после всех товаров
    if totals.get('СумНалВсего') or totals.get('СтТовУчНалВсего'):
        vsego_opl_attrs = {}

        # Вычисляем СтТовБезНДСВсего, если его нет в totals
        if 'СтТовБезНДСВсего' not in totals or pd.isna(totals.get('СтТовБезНДСВсего')) or totals.get(
                'СтТовБезНДСВсего') == '':
            if totals.get('СтТовУчНалВсего') and totals.get('СумНалВсего'):
                try:
                    st_tov_with_tax = float(totals.get('СтТовУчНалВсего'))
                    tax_total = float(totals.get('СумНалВсего'))
                    st_tov_without_tax = st_tov_with_tax - tax_total
                    vsego_opl_attrs['СтТовБезНДСВсего'] = f"{st_tov_without_tax:.2f}"
                except (ValueError, TypeError) as e:
                    logger.warning(f"Не удалось вычислить СтТовБезНДСВсего: {e}")
            else:
                logger.warning(
                    "Недостаточно данных для вычисления СтТовБезНДСВсего (СтТовУчНалВсего или СумНалВсего отсутствуют).")
        else:
            vsego_opl_attrs['СтТовБезНДСВсего'] = str(totals.get('СтТовБезНДСВсего'))

        # Добавляем СтТовУчНалВсего
        if totals.get('СтТовУчНалВсего'):
            total_with_tax = totals.get('СтТовУчНалВсего')
            if pd.isna(total_with_tax) or str(total_with_tax).lower() == 'nan':
                logger.warning("Значение СтТовУчНалВсего содержит nan, будет проигнорировано.")
            else:
                vsego_opl_attrs['СтТовУчНалВсего'] = str(total_with_tax)

        # Создаём элемент ВсегоОпл с атрибутами
        vsego_opl = etree.SubElement(tabl, "ВсегоОпл", **vsego_opl_attrs)

        # Добавляем вложенный элемент СумНалВсего
        if totals.get('СумНалВсего'):
            sum_nal_vsego = etree.SubElement(vsego_opl, "СумНалВсего")
            etree.SubElement(sum_nal_vsego, "СумНал").text = str(totals.get('СумНалВсего'))

    sv_prod_per_elem = etree.SubElement(doc, "СвПродПер")
    sv_per = etree.SubElement(sv_prod_per_elem, "СвПер",
                              ДатаПер=str(sv_prod_per.get('ДатаПер', '')),
                              СодОпер=str(sv_prod_per.get('СодОпер', '')))
    if sv_prod_per.get('БезДокОснПер'):
        etree.SubElement(sv_per, "БезДокОснПер").text = str(sv_prod_per.get('БезДокОснПер', ''))
    else:
        etree.SubElement(sv_per, "ОснПер",
                         РеквДатаДок=str(sv_prod_per.get('РеквДатаДок', '')),
                         РеквНаимДок=str(sv_prod_per.get('РеквНаимДок', '')),
                         РеквНомерДок=str(sv_prod_per.get('РеквНомерДок', '')))
    logger.debug("Создан элемент СвПродПер")

    podpisant_elem = etree.SubElement(doc, "Подписант",
                                      Должн=str(podpisant.get('Должн', '')),
                                      СпосПодтПолном=str(podpisant.get('СпосПодтПолном', '')))
    etree.SubElement(podpisant_elem, "ФИО",
                     Имя=str(podpisant.get('Имя', '')),
                     Отчество=str(podpisant.get('Отчество', '')),
                     Фамилия=str(podpisant.get('Фамилия', '')))
    logger.debug("Создан элемент Подписант")

    tree = etree.ElementTree(root)
    with open(output_file, 'wb') as f:
        f.write(etree.tostring(tree, pretty_print=True, encoding='windows-1251', xml_declaration=True))
    logger.info(f"XML успешно сохранён в {output_file}")