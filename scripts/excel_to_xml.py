# Путь в проекте: project/scripts/excel_to_xml.py
import sys
import os
from pathlib import Path
import pandas as pd
from lxml import etree


def excel_to_xml(input_file, output_file):
    # Убедимся, что директория для output_file существует
    output_path = Path(output_file)
    output_dir = output_path.parent
    if not output_dir.exists():
        print(f"Папка {output_dir} не найдена. Создаю...")
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"Папка {output_dir} создана.")

    # Читаем Excel-файл с данными о товарах
    try:
        df = pd.read_excel(input_file, sheet_name='Товары', dtype=str)
    except FileNotFoundError:
        print(f"Ошибка: файл {input_file} не найден.")
        return
    except Exception as e:
        print(f"Ошибка при чтении листа Товары из Excel: {e}")
        df = pd.DataFrame()

    # Читаем служебные данные из других листов
    try:
        podpisant_df = pd.read_excel(input_file, sheet_name='Подписант')
        podpisant = podpisant_df.iloc[0].to_dict() if not podpisant_df.empty else {}
    except Exception as e:
        print(f"Ошибка при чтении листа Подписант: {e}")
        podpisant = {}

    try:
        sv_prod_per_df = pd.read_excel(input_file, sheet_name='СвПродПер')
        sv_prod_per = sv_prod_per_df.iloc[0].to_dict() if not sv_prod_per_df.empty else {}
    except Exception as e:
        print(f"Ошибка при чтении листа СвПродПер: {e}")
        sv_prod_per = {}

    try:
        sv_sch_fakt_df = pd.read_excel(input_file, sheet_name='СвСчФакт')
        sv_sch_fakt = sv_sch_fakt_df.iloc[0].to_dict() if not sv_sch_fakt_df.empty else {}
    except Exception as e:
        print(f"Ошибка при чтении листа СвСчФакт: {e}")
        sv_sch_fakt = {}

    try:
        inf_pol_fhzh1_df = pd.read_excel(input_file, sheet_name='ИнфПолФХЖ1')
        inf_pol_fhzh1 = inf_pol_fhzh1_df.to_dict('records') if not inf_pol_fhzh1_df.empty else []
    except Exception as e:
        print(f"Ошибка при чтении листа ИнфПолФХЖ1: {e}")
        inf_pol_fhzh1 = []

    # Создаем корневой элемент XML
    try:
        root = etree.Element("Файл", ВерсПрог="СБиС3", ВерсФорм="5.03")
    except Exception as e:
        print(f"Ошибка при создании корневого элемента XML: {e}")
        return

    # Добавляем элемент Документ
    try:
        doc = etree.SubElement(root, "Документ", КНД="1115131",
                               НаимДокОпр="Универсальный передаточный документ",
                               ПоФактХЖ="Документ об отгрузке товаров (выполнении работ), передаче имущественных прав (документ об оказании услуг)",
                               Функция="СЧФДОП")
    except Exception as e:
        print(f"Ошибка при создании элемента Документ: {e}")
        return

    # Добавляем СвСчФакт
    try:
        sv_sch_fakt_elem = etree.SubElement(doc, "СвСчФакт",
                                            ДатаДок=str(sv_sch_fakt.get('ДатаДок', '')),
                                            НомерДок=str(sv_sch_fakt.get('НомерДок', '')))

        # СвПрод
        sv_prod = etree.SubElement(sv_sch_fakt_elem, "СвПрод", ОКПО=str(sv_sch_fakt.get('ОКПО_СвПрод', '')))
        id_sv = etree.SubElement(sv_prod, "ИдСв")
        if sv_sch_fakt.get('ИННЮЛ_СвПрод'):
            sv_yul_uch = etree.SubElement(id_sv, "СвЮЛУч",
                                          ИННЮЛ=str(sv_sch_fakt.get('ИННЮЛ_СвПрод', '')),
                                          КПП=str(sv_sch_fakt.get('КПП_СвПрод', '')),
                                          НаимОрг=str(sv_sch_fakt.get('НаимОрг_СвПрод', '')))
        else:
            sv_ip = etree.SubElement(id_sv, "СвИП",
                                     ИННФЛ=str(sv_sch_fakt.get('ИННФЛ_СвПрод', '')))
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
                             КодРегион=str(sv_sch_fakt.get('КодРегион_СвПрод', '')),
                             НаимРегион=str(sv_sch_fakt.get('НаимРегион_СвПрод', '')),
                             Индекс=str(sv_sch_fakt.get('Индекс_СвПрод', '')),
                             Город=str(sv_sch_fakt.get('Город_СвПрод', '')),
                             Улица=str(sv_sch_fakt.get('Улица_СвПрод', '')),
                             Дом=str(sv_sch_fakt.get('Дом_СвПрод', '')))
        if sv_sch_fakt.get('Тлф_СвПрод') or sv_sch_fakt.get('ЭлПочта_СвПрод'):
            kontakt = etree.SubElement(sv_prod, "Контакт")
            if sv_sch_fakt.get('Тлф_СвПрод'):
                etree.SubElement(kontakt, "Тлф").text = str(sv_sch_fakt.get('Тлф_СвПрод', ''))
            if sv_sch_fakt.get('ЭлПочта_СвПрод'):
                etree.SubElement(kontakt, "ЭлПочта").text = str(sv_sch_fakt.get('ЭлПочта_СвПрод', ''))

        # ГрузОт
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

        # ГрузПолуч
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

        # ДокПодтвОтгрНом
        etree.SubElement(sv_sch_fakt_elem, "ДокПодтвОтгрНом",
                         РеквДатаДок=str(sv_sch_fakt.get('РеквДатаДок_ДокПодтв', '')),
                         РеквНаимДок=str(sv_sch_fakt.get('РеквНаимДок_ДокПодтв', '')),
                         РеквНомерДок=str(sv_sch_fakt.get('РеквНомерДок_ДокПодтв', '')))

        # СвПокуп
        sv_pokup = etree.SubElement(sv_sch_fakt_elem, "СвПокуп",
                                    ОКПО=str(sv_sch_fakt.get('ОКПО_СвПокуп', '')),
                                    СокрНаим=str(sv_sch_fakt.get('СокрНаим_СвПокуп', '')))
        id_sv = etree.SubElement(sv_pokup, "ИдСв")
        sv_ip = etree.SubElement(id_sv, "СвИП",
                                 ИННФЛ=str(sv_sch_fakt.get('ИННФЛ_СвПокуп', '')),
                                 СвГосРегИП=str(sv_sch_fakt.get('СвГосРегИП_СвПокуп', '')))
        fio = etree.SubElement(sv_ip, "ФИО",
                               Имя=str(sv_sch_fakt.get('Имя_СвПокуп', '')),
                               Отчество=str(sv_sch_fakt.get('Отчество_СвПокуп', '')),
                               Фамилия=str(sv_sch_fakt.get('Фамилия_СвПокуп', '')))
        adres = etree.SubElement(sv_pokup, "Адрес")
        etree.SubElement(adres, "АдрИнф",
                         АдрТекст=str(sv_sch_fakt.get('АдрТекст_СвПокуп', '')),
                         КодСтр=str(sv_sch_fakt.get('КодСтр_СвПокуп', '')),
                         НаимСтран=str(sv_sch_fakt.get('НаимСтран_СвПокуп', '')))

        # ДенИзм
        etree.SubElement(sv_sch_fakt_elem, "ДенИзм",
                         КодОКВ=str(sv_sch_fakt.get('КодОКВ', '')),
                         НаимОКВ=str(sv_sch_fakt.get('НаимОКВ', '')))

        # ИнфПолФХЖ1
        if inf_pol_fhzh1:
            inf_pol_fhzh1_elem = etree.SubElement(sv_sch_fakt_elem, "ИнфПолФХЖ1")
            for entry in inf_pol_fhzh1:
                etree.SubElement(inf_pol_fhzh1_elem, "ТекстИнф",
                                 Идентиф=str(entry.get('Идентиф', '')),
                                 Значен=str(entry.get('Значен', '')))
    except Exception as e:
        print(f"Ошибка при создании элемента СвСчФакт: {e}")

    # Добавляем ТаблСчФакт
    try:
        tabl = etree.SubElement(doc, "ТаблСчФакт")
        if not df.empty:
            for _, row in df.iterrows():
                # Пропускаем строки, где все ключевые поля пустые
                if not any(pd.notna(row[col]) for col in ['Номер строки', 'Наименование', 'Количество']):
                    continue

                # Формируем атрибуты для СведТов
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

                sved_tov = etree.SubElement(tabl, "СведТов", **sved_tov_attrs)

                # СвДТ
                if pd.notna(row.get('КодПроисх')) or pd.notna(row.get('НомерДТ')):
                    etree.SubElement(sved_tov, "СвДТ",
                                     КодПроисх=str(row['КодПроисх']) if pd.notna(row.get('КодПроисх')) else "",
                                     НомерДТ=str(row['НомерДТ']) if pd.notna(row.get('НомерДТ')) else "")

                # ДопСведТов
                dop_sved_attrs = {
                    'КодТов': str(row['Код товара']) if pd.notna(row.get('Код товара')) else "",
                    'ПрТовРаб': "1"
                }
                # Если есть ГТИН, добавляем его как атрибут
                if pd.notna(row.get('ГТИН')):
                    dop_sved_attrs['ГТИН'] = str(row['ГТИН'])

                dop_sved = etree.SubElement(sved_tov, "ДопСведТов", **dop_sved_attrs)

                # КрНаимСтрПр
                if pd.notna(row.get('КрНаимСтрПр')):
                    etree.SubElement(dop_sved, "КрНаимСтрПр").text = str(row['КрНаимСтрПр'])

                # КИЗ
                if pd.notna(row['КИЗ']) and row['КИЗ']:
                    nom_sred = etree.SubElement(dop_sved, "НомСредИдентТов")
                    for kiz in row['КИЗ'].split('; '):
                        etree.SubElement(nom_sred, "КИЗ").text = kiz

                # Акциз
                akciz = etree.SubElement(sved_tov, "Акциз")
                etree.SubElement(akciz, "БезАкциз").text = "без акциза"

                # СумНал
                sum_nal = etree.SubElement(sved_tov, "СумНал")
                sum_nal_value = str(row['Сумма НДС']) if pd.notna(row['Сумма НДС']) else ""
                if sum_nal_value.lower() in ['без ндс', 'безндс']:
                    etree.SubElement(sum_nal, "БезНДС").text = "без НДС"
                elif sum_nal_value:
                    etree.SubElement(sum_nal, "СумНал").text = sum_nal_value

                # ИнфПолФХЖ2
                if pd.notna(row.get('КодПокупателя')) and row['КодПокупателя']:
                    etree.SubElement(sved_tov, "ИнфПолФХЖ2", Идентиф="КодПокупателя", Значен=str(row['КодПокупателя']))
                if pd.notna(row.get('НазваниеПокупателя')) and row['НазваниеПокупателя']:
                    etree.SubElement(sved_tov, "ИнфПолФХЖ2", Идентиф="НазваниеПокупателя",
                                     Значен=str(row['НазваниеПокупателя']))
                # Используем GTIN из столбца GTIN, если ГТИН уже не добавлен как атрибут
                if pd.notna(row.get('GTIN')) and row['GTIN'] and not pd.notna(row.get('ГТИН')):
                    etree.SubElement(sved_tov, "ИнфПолФХЖ2", Идентиф="GTIN", Значен=str(row['GTIN']))
    except Exception as e:
        print(f"Ошибка при создании элемента ТаблСчФакт: {e}")

    # Добавляем СвПродПер
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
    except Exception as e:
        print(f"Ошибка при создании элемента СвПродПер: {e}")

    # Добавляем Подписант
    try:
        podpisant_elem = etree.SubElement(doc, "Подписант",
                                          Должн=str(podpisant.get('Должн', '')),
                                          СпосПодтПолном=str(podpisant.get('СпосПодтПолном', '')))
        etree.SubElement(podpisant_elem, "ФИО",
                         Имя=str(podpisant.get('Имя', '')),
                         Отчество=str(podpisant.get('Отчество', '')),
                         Фамилия=str(podpisant.get('Фамилия', '')))
    except Exception as e:
        print(f"Ошибка при создании элемента Подписант: {e}")

    # Сохраняем XML
    try:
        tree = etree.ElementTree(root)
        with open(output_file, 'wb') as f:
            f.write(etree.tostring(tree, pretty_print=True, encoding='windows-1251', xml_declaration=True))
        print(f"XML успешно сохранён в {output_file}")
    except Exception as e:
        print(f"Ошибка при сохранении XML: {e}")
