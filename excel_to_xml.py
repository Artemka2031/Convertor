from pathlib import Path

import pandas as pd
from lxml import etree

from utils import setup_logging

logger = setup_logging()

def excel_to_xml(input_file, output_file):
    output_path = Path(output_file)
    output_dir = output_path.parent
    if not output_dir.exists():
        logger.info(f"Папка {output_dir} не найдена. Создаю...")
        output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"Папка {output_dir} создана.")

    # Чтение всех листов
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
        podpisant_df = pd.read_excel(input_file, sheet_name='Подписант')
        podpisant = podpisant_df.iloc[0].to_dict() if not podpisant_df.empty else {}
        logger.debug(f"Найдено записей Подписант: {len(podpisant_df)}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа Подписант: {e}")
        podpisant = {}

    try:
        sv_prod_per_df = pd.read_excel(input_file, sheet_name='СвПродПер')
        sv_prod_per = sv_prod_per_df.iloc[0].to_dict() if not sv_prod_per_df.empty else {}
        logger.debug(f"Найдено записей СвПродПер: {len(sv_prod_per_df)}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа СвПродПер: {e}")
        sv_prod_per = {}

    try:
        sv_sch_fakt_df = pd.read_excel(input_file, sheet_name='СвСчФакт')
        # Объединяем данные из нескольких строк
        sv_sch_fakt = {}
        if not sv_sch_fakt_df.empty:
            for index, row in sv_sch_fakt_df.iterrows():
                sv_sch_fakt.update({k: str(v) if pd.notna(v) else '' for k, v in row.items()})
        logger.debug(f"Найдено записей СвСчФакт: {len(sv_sch_fakt_df)}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа СвСчФакт: {e}")
        sv_sch_fakt = {}

    try:
        inf_pol_fhzh1_df = pd.read_excel(input_file, sheet_name='ИнфПолФХЖ1')
        inf_pol_fhzh1 = inf_pol_fhzh1_df.to_dict('records') if not inf_pol_fhzh1_df.empty else []
        logger.debug(f"Найдено записей ИнфПолФХЖ1: {len(inf_pol_fhzh1)}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа ИнфПолФХЖ1: {e}")
        inf_pol_fhzh1 = []

    try:
        totals_df = pd.read_excel(input_file, sheet_name='Итоги')
        totals = totals_df.iloc[0].to_dict() if not totals_df.empty else {}
        logger.debug(f"Найдены данные Итоги: {totals}")
    except Exception as e:
        logger.error(f"Ошибка при чтении листа Итоги: {e}")
        totals = {}

    try:
        root = etree.Element("Файл", ВерсПрог="СБиС3", ВерсФорм="5.03")
        logger.debug("Создан корневой элемент XML: Файл")
    except Exception as e:
        logger.error(f"Ошибка при создании корневого элемента XML: {e}")
        return

    try:
        doc = etree.SubElement(root, "Документ", КНД="1115131",
                               НаимДокОпр="Универсальный передаточный документ",
                               ПоФактХЖ="Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (документ об оказании услуг)",
                               Функция="СЧФДОП",
                               ДатаИнфПр=sv_sch_fakt.get('ДатаИнфПр', ''),
                               ВремИнфПр=sv_sch_fakt.get('ВремИнфПр', ''),
                               НаимЭконСубСост=sv_sch_fakt.get('НаимЭконСубСост', ''))
        logger.debug("Создан элемент Документ")
    except Exception as e:
        logger.error(f"Ошибка при создании элемента Документ: {e}")
        return

    try:
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
        if sv_sch_fakt.get('АдрТекст_СвПрод'):
            etree.SubElement(adres, "АдрИнф",
                             АдрТекст=str(sv_sch_fakt.get('АдрТекст_СвПрод', '')),
                             КодСтр=str(sv_sch_fakt.get('КодСтр_СвПрод', '')),
                             НаимСтран=str(sv_sch_fakt.get('НаимСтран_СвПрод', '')))
        else:
            etree.SubElement(adres, "АдрРФ",
                             Индекс=str(sv_sch_fakt.get('Индекс_СвПрод', '')),
                             КодРегион=str(sv_sch_fakt.get('КодРегион_СвПрод', '')),
                             НаимРегион=str(sv_sch_fakt.get('НаимРегион_СвПрод', '')),
                             Город=str(sv_sch_fakt.get('Город_СвПрод', '')))
        if sv_sch_fakt.get('Тлф_СвПрод') or sv_sch_fakt.get('ЭлПочта_СвПрод'):
            kontakt = etree.SubElement(sv_prod, "Контакт")
            if sv_sch_fakt.get('Тлф_СвПрод'):
                etree.SubElement(kontakt, "Тлф").text = str(sv_sch_fakt.get('Тлф_СвПрод', ''))
            if sv_sch_fakt.get('ЭлПочта_СвПрод'):
                etree.SubElement(kontakt, "ЭлПочта").text = str(sv_sch_fakt.get('ЭлПочта_СвПрод', ''))

        if sv_sch_fakt.get('НомерСчёта_СвПрод') or sv_sch_fakt.get('БИК_СвПрод') or sv_sch_fakt.get('НаимБанк_СвПрод'):
            bank_rekv = etree.SubElement(sv_prod, "БанкРекв", НомерСчёта=str(sv_sch_fakt.get('НомерСчёта_СвПрод', '')))
            etree.SubElement(bank_rekv, "СвБанк",
                             БИК=str(sv_sch_fakt.get('БИК_СвПрод', '')),
                             НаимБанк=str(sv_sch_fakt.get('НаимБанк_СвПрод', '')),
                             КорСчет=str(sv_sch_fakt.get('КорСчет_СвПрод', '')))

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
            etree.SubElement(adres, "АдрИнф",
                             АдрТекст=str(sv_sch_fakt.get('АдрТекст_ГрузОт', '')),
                             КодСтр=str(sv_sch_fakt.get('КодСтр_ГрузОт', '')),
                             НаимСтран=str(sv_sch_fakt.get('НаимСтран_ГрузОт', '')))
            kontakt = etree.SubElement(gruz_otpr, "Контакт")
            etree.SubElement(kontakt, "Тлф").text = str(sv_sch_fakt.get('Тлф_ГрузОт', ''))
            etree.SubElement(kontakt, "ЭлПочта").text = str(sv_sch_fakt.get('ЭлПочта_ГрузОт', ''))

        if sv_sch_fakt.get('ОКПО_ГрузПолуч') or sv_sch_fakt.get('СокрНаим_ГрузПолуч'):
            gruz_poluch = etree.SubElement(sv_sch_fakt_elem, "ГрузПолуч",
                                           ОКПО=str(sv_sch_fakt.get('ОКПО_ГрузПолуч', '')),
                                           СокрНаим=str(sv_sch_fakt.get('СокрНаим_ГрузПолуч', '')))
            id_sv = etree.SubElement(gruz_poluch, "ИдСв")
            sv_ip = etree.SubElement(id_sv, "СвИП",
                                     ИННФЛ=str(sv_sch_fakt.get('ИННФЛ_ГрузПолуч', '')),
                                     СвГосРегИП=str(sv_sch_fakt.get('СвГосРегИП_ГрузПолуч', '')))
            fio = etree.SubElement(sv_ip, "ФИО",
                                   Имя=str(sv_sch_fakt.get('Имя_ГрузПолуч', '')),
                                   Отчество=str(sv_sch_fakt.get('Отчество_ГрузПолуч', '')),
                                   Фамилия=str(sv_sch_fakt.get('Фамилия_ГрузПолуч', '')))
            adres = etree.SubElement(gruz_poluch, "Адрес")
            etree.SubElement(adres, "АдрРФ",
                             Дом=str(sv_sch_fakt.get('Дом_ГрузПолуч', '')),
                             Индекс=str(sv_sch_fakt.get('Индекс_ГрузПолуч', '')),
                             КодРегион=str(sv_sch_fakt.get('КодРегион_ГрузПолуч', '')),
                             НаимРегион=str(sv_sch_fakt.get('НаимРегион_ГрузПолуч', '')),
                             Улица=str(sv_sch_fakt.get('Улица_ГрузПолуч', '')))

        etree.SubElement(sv_sch_fakt_elem, "ДокПодтвОтгрНом",
                         РеквДатаДок=str(sv_sch_fakt.get('РеквДатаДок_ДокПодтв', '')),
                         РеквНаимДок=str(sv_sch_fakt.get('РеквНаимДок_ДокПодтв', '')),
                         РеквНомерДок=str(sv_sch_fakt.get('РеквНомерДок_ДокПодтв', '')))

        sv_pokup = etree.SubElement(sv_sch_fakt_elem, "СвПокуп",
                                    ОКПО=str(sv_sch_fakt.get('ОКПО_СвПокуп', '')),
                                    SокрНаим=str(sv_sch_fakt.get('СокрНаим_СвПокуп', '')))
        id_sv = etree.SubElement(sv_pokup, "ИдСв")
        sv_ip = etree.SubElement(id_sv, "СвИП",
                                 ИННФЛ=str(sv_sch_fakt.get('ИННФЛ_СвПокуп', '')),
                                 СвГосРегИП=str(sv_sch_fakt.get('СвГосРегИП_СвПокуп', '')))
        fio = etree.SubElement(sv_ip, "ФИО",
                               Имя=str(sv_sch_fakt.get('Имя_СвПокуп', '')),
                               Отчество=str(sv_sch_fakt.get('Отчество_СвПокуп', '')),
                               Фамилия=str(sv_sch_fakt.get('Фамилия_СвПокуп', '')))
        adres = etree.SubElement(sv_pokup, "Адрес")
        etree.SubElement(adres, "АдрРФ",
                         Индекс=str(sv_sch_fakt.get('Индекс_ГрузПолуч', '')),
                         КодРегион=str(sv_sch_fakt.get('КодРегион_ГрузПолуч', '')),
                         НаимРегион=str(sv_sch_fakt.get('НаимРегион_ГрузПолуч', '')),
                         Город=str(sv_sch_fakt.get('Город_ГрузПолуч', '')))

        if sv_sch_fakt.get('НомерСчёта_СвПокуп') or sv_sch_fakt.get('БИК_СвПокуп') or sv_sch_fakt.get(
                'НаимБанк_СвПокуп'):
            bank_rekv = etree.SubElement(sv_pokup, "БанкРекв",
                                         НомерСчёта=str(sv_sch_fakt.get('НомерСчёта_СвПокуп', '')))
            etree.SubElement(bank_rekv, "СвБанк",
                             БИК=str(sv_sch_fakt.get('БИК_СвПокуп', '')),
                             НаимБанк=str(sv_sch_fakt.get('НаимБанк_СвПокуп', '')),
                             КорСчет=str(sv_sch_fakt.get('КорСчет_СвПокуп', '')))

        etree.SubElement(sv_sch_fakt_elem, "ДенИзм",
                         КодОКВ=str(sv_sch_fakt.get('КодОКВ', '')),
                         НаимОКВ=str(sv_sch_fakt.get('НаимОКВ', '')))
    except Exception as e:
        logger.error(f"Ошибка при создании элемента СвСчФакт: {e}")

    try:
        tabl = etree.SubElement(doc, "ТаблСчФакт",
                                ВсегоОпл=str(totals.get('ВсегоОпл', '')),
                                СумНалВсего=str(totals.get('СумНалВсего', '')))
        if not df.empty:
            for _, row in df.iterrows():
                if not any(pd.notna(row[col]) for col in ['Номер строки', 'Наименование', 'Количество']):
                    continue

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
                if pd.notna(row.get('ГТИН')):
                    dop_sved_attrs['ГТИН'] = str(row['ГТИН'])

                dop_sved = etree.SubElement(sved_tov, "ДопСведТов", **dop_sved_attrs)

                if pd.notna(row.get('КрНаимСтрПр')):
                    etree.SubElement(dop_sved, "КрНаимСтрPr").text = str(row['КрНаимСтрPr'])

                if pd.notna(row['КИЗ']) and row['КИЗ']:
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

        if totals.get('ВсегоОпл') or totals.get('СумНалВсего'):
            vsego_opl = etree.SubElement(tabl, "ВсегоОпл",
                                         СтТовБезНДСВсего=str(totals.get('ВсегоОпл', '')) if totals.get(
                                             'ВсегоОпл') else '')
            if totals.get('СумНалВсего'):
                sum_nal_vsego = etree.SubElement(vsego_opl, "СумНалВсего")
                etree.SubElement(sum_nal_vsego, "СумНал").text = str(totals.get('СумНалВсего', ''))
    except Exception as e:
        logger.error(f"Ошибка при создании элемента ТаблСчФакт: {e}")

    try:
        sv_prod_per_elem = etree.SubElement(doc, "СвПродПер")
        sv_per = etree.SubElement(sv_prod_per_elem, "СвПер",
                                  ДатаПер=str(sv_prod_per.get('ДатаПер', '')),
                                  СодОпер=str(sv_prod_per.get('СодОпер', '')))
        if sv_prod_per.get('БезДокОснПер'):
            etree.SubElement(sv_per, "БезДокОснПер").text = str(sv_prod_per.get('БезДокОснПер', ''))
        else:
            osn_per = etree.SubElement(sv_per, "ОснПер",
                                       РеквДатаДок=str(sv_prod_per.get('РеквДатаДок', '')),
                                       РеквНаимДок=str(sv_prod_per.get('РеквНаимДок', '')),
                                       РеквНомерДок=str(sv_prod_per.get('РеквНомерДок', '')))
        logger.debug("Создан элемент СвПродПер")
    except Exception as e:
        logger.error(f"Ошибка при создании элемента СвПродПер: {e}")

    try:
        podpisant_elem = etree.SubElement(doc, "Подписант",
                                          Должн=str(podpisant.get('Должн', '')),
                                          СпосПодтПолном=str(podpisant.get('СпосПодтПолном', '')))
        etree.SubElement(podpisant_elem, "ФИО",
                         Имя=str(podpisant.get('Имя', '')),
                         Отчество=str(podpisant.get('Отчество', '')),
                         Фамилия=str(podpisant.get('Фамилия', '')))
        logger.debug("Создан элемент Подписант")
    except Exception as e:
        logger.error(f"Ошибка при создании элемента Подписант: {e}")

    try:
        tree = etree.ElementTree(root)
        with open(output_file, 'wb') as f:
            f.write(etree.tostring(tree, pretty_print=True, encoding='windows-1251', xml_declaration=True))
        logger.info(f"XML успешно сохранён в {output_file}")
    except Exception as e:
        logger.error(f"Ошибка при сохранении XML: {e}")
