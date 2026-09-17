"""
Автоматизация заказа пропусков для ММТС-9 v2
Заполняет шаблон Excel заявки на основе входных данных
"""

import openpyxl
from datetime import datetime
import json
import os
import sys

class PassRequestAutomation:
    """Класс для автоматизации заполнения заявок на пропуск"""
    
    def __init__(self, template_file=None):
        if template_file:
            template_candidates = [template_file]
        else:
            template_candidates = [
                'Пример пропуска ТехРу (актуальный).xlsx',
                '04-08.1242.xlsx',
                '04-08.1221.xlsx',
            ]

        resolved_template = None
        for candidate in template_candidates:
            resolved_template = self._resolve_template_path(candidate)
            if resolved_template:
                break

        if not resolved_template:
            raise FileNotFoundError(
                "Не найден шаблон Excel 'Пример пропуска ТехРу (актуальный).xlsx'. "
                "Поместите файл шаблона рядом с программой и повторите попытку."
            )

        self.template_file = resolved_template
        self.workbook = None
        self.data_sheet = None
        self.visitors_sheet = None
        self.equipment_sheet = None
        self.vehicles_sheet = None
        
        # Карта ячеек для основных данных (определена путем ручного анализа Excel)
        self.data_cells = {
            'organization': 'C2',
            'inn': 'C3',
            'contract_number': 'C4',
            'contract_date': 'E4',
            'contract_valid_until': 'C5',
            'rooms': 'C6',
            'work_description': 'C7',
            'work_start_date': 'C8',
            'work_end_date': 'C9',
            'purpose': 'C10',
            'responsible_person': 'C11',
            'responsible_name': 'D11',
            'request_number': 'C12',
            'request_date': 'E12',
            'contact_phone': 'C13',
            'object_address': 'C14',
        }
        
        # Начальные строки для динамических данных
        self.visitor_start_row = 3
        self.visitor_columns = {
            'number': 2,
            'surname': 3,
            'name': 4,
            'patronymic': 5,
            'birth_date': 6,
            'citizenship': 7,
            'passport_series': 8,
            'passport_number': 9,
            'passport_issue_date': 10,
            'passport_issued_by': 11,
            'organization': 12,
            'electrical_safety_group': 13,
            'electrical_safety_certificate': 14,
            'note': 15
        }
        
        self.equipment_start_row = 3
        # Лист «Оборудование»: B№ Cимя Dмарка Eмодель Fсерийный Gкол-во Hед. Iвнос
        # Jпримечание Kпомещение L«ряд , место» Mпричина Nряд Oместо
        self.equipment_columns = {
            'number': 2,
            'name': 3,
            'brand': 4,
            'model': 5,
            'serial_number': 6,
            'quantity': 7,
            'unit': 8,
            'action': 9,
            'note': 10,
            'room': 11,
            'row_place': 12,
            'reason': 13,
            'row': 14,
            'place': 15,
        }
        
        self.vehicle_start_row = 4
        self.vehicle_columns = {
            'number': 2,
            'license_plate': 4,
            'driver_surname': 5,
            'driver_name': 6,
            'driver_patronymic': 7,
            'driver_birth_date': 8,
            'driver_citizenship': 9,
            'passport_series': 10,
            'passport_number': 11,
            'passport_issue_date': 12,
            'passport_issued_by': 13,
            'driver_organization': 14,
        }
    
    def _resolve_template_path(self, template_file):
        """Определение пути к файлу шаблона с учетом запуска из EXE"""
        if not template_file:
            return None

        if os.path.exists(template_file):
            return template_file

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        candidate = os.path.join(base_path, template_file)
        if os.path.exists(candidate):
            return candidate

        parent_path = os.path.dirname(base_path)
        candidate = os.path.join(parent_path, template_file)
        if os.path.exists(candidate):
            return candidate

        return None

    def parse_date(self, date_str):
        """Парсинг строки даты в объект datetime"""
        if not date_str or date_str == '':
            return None
        try:
            # Пробуем разные форматы
            for fmt in ['%d.%m.%Y', '%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d.%m.%y', '%d/%m/%y']:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
            return None
        except Exception:
            return None
        
    def load_template(self):
        """Загрузка шаблона Excel"""
        print(f"Загрузка шаблона: {self.template_file}")
        self.workbook = openpyxl.load_workbook(self.template_file, data_only=False)
        self.data_sheet = self.workbook['Данные заявки']
        self.visitors_sheet = self.workbook['Посетители']
        self.equipment_sheet = self.workbook['Оборудование']
        self.vehicles_sheet = self.workbook['Автотранспорт']
        print("Шаблон успешно загружен")
        
    def fill_request_data(self, data):
        """Заполнение основных данных заявки"""
        print("Заполнение основных данных заявки...")

        for key, cell_address in self.data_cells.items():
            if key not in data:
                continue
            value = data[key]
            if key == 'rooms':
                rooms = value if isinstance(value, list) else [
                    part.strip() for part in str(value).replace(';', ',').split(',') if part.strip()
                ]
                for offset in range(6):
                    cell = self.data_sheet.cell(6, 3 + offset)
                    cell.value = rooms[offset] if offset < len(rooms) else None
                    if offset < len(rooms):
                        print(f"  room[{offset}]: {rooms[offset]}")
                continue

            cell = self.data_sheet[cell_address]
            if ('date' in key.lower() or key == 'contract_valid_until') and isinstance(value, str):
                date_obj = self.parse_date(value)
                if date_obj:
                    value = date_obj
            cell.value = value
            print(f"  {key}: {data[key]}")

        print("Основные данные заполнены")
    
    def _clear_block(self, sheet, start_row, end_row, columns):
        """Очистка ранее заполненных значений, чтобы в шаблоне не оставался мусор."""
        for row in range(start_row, end_row + 1):
            for col in columns:
                if col == 'number':
                    continue
                sheet.cell(row, columns[col]).value = None

    def fill_visitors(self, visitors):
        """Заполнение данных посетителей"""
        print(f"\nЗаполнение данных посетителей ({len(visitors)} чел.)...")
        self._clear_block(self.visitors_sheet, self.visitor_start_row, self.visitor_start_row + 19, self.visitor_columns)

        for idx, visitor in enumerate(visitors[:20], start=0):  # Максимум 20
            row = self.visitor_start_row + idx
            
            for key, col in self.visitor_columns.items():
                if key == 'number':
                    self.visitors_sheet.cell(row, col).value = idx + 1
                elif key in visitor:
                    value = visitor[key]
                    
                    # Преобразуем даты в datetime объекты
                    if 'date' in key.lower() and isinstance(value, str):
                        date_obj = self.parse_date(value)
                        if date_obj:
                            value = date_obj
                    
                    self.visitors_sheet.cell(row, col).value = value
            
            print(f"  Посетитель {idx + 1}: {visitor.get('surname', '')} {visitor.get('name', '')} {visitor.get('patronymic', '')}")
        
        print(f"Данные {len(visitors)} посетителей заполнены")
    
    def fill_equipment(self, equipment):
        """Заполнение данных об оборудовании"""
        print(f"\nЗаполнение данных оборудования ({len(equipment)} позиций)...")
        self._clear_block(self.equipment_sheet, self.equipment_start_row, self.equipment_start_row + 39, self.equipment_columns)

        for idx, item in enumerate(equipment[:40], start=0):  # Максимум 40
            row = self.equipment_start_row + idx
            payload = dict(item)
            row_value = str(payload.get('row', '') or '').strip()
            place_value = str(payload.get('place', '') or '').strip()
            if row_value or place_value:
                payload['row_place'] = " , ".join(part for part in (row_value, place_value) if part)
            if not payload.get('unit'):
                payload['unit'] = 'шт.'
            if not payload.get('quantity'):
                payload['quantity'] = '1'

            for key, col in self.equipment_columns.items():
                if key == 'number':
                    self.equipment_sheet.cell(row, col).value = idx + 1
                elif key in payload:
                    value = payload[key]
                    
                    # Преобразуем даты в datetime объекты
                    if 'date' in key.lower() and isinstance(value, str):
                        date_obj = self.parse_date(value)
                        if date_obj:
                            value = date_obj
                    
                    self.equipment_sheet.cell(row, col).value = value
            
            print(f"  Оборудование {idx + 1}: {item.get('name', '')} ({item.get('quantity', '')} {item.get('unit', '')})")
        
        print(f"Данные {len(equipment)} позиций оборудования заполнены")
    
    def fill_vehicles(self, vehicles):
        """Заполнение данных автотранспорта"""
        print(f"\nЗаполнение данных автотранспорта ({len(vehicles)} машин)...")
        self._clear_block(self.vehicles_sheet, self.vehicle_start_row, self.vehicle_start_row + 19, self.vehicle_columns)

        for idx, vehicle in enumerate(vehicles[:20], start=0):  # Максимум 20
            row = self.vehicle_start_row + idx
            
            for key, col in self.vehicle_columns.items():
                if key == 'number':
                    self.vehicles_sheet.cell(row, col).value = idx + 1
                elif key in vehicle:
                    value = vehicle[key]
                    
                    # Преобразуем даты в datetime объекты
                    if 'date' in key.lower() and isinstance(value, str):
                        date_obj = self.parse_date(value)
                        if date_obj:
                            value = date_obj
                    
                    self.vehicles_sheet.cell(row, col).value = value
            
            print(f"  Машина {idx + 1}: {vehicle.get('license_plate', '')} - {vehicle.get('driver_surname', '')} {vehicle.get('driver_name', '')}")
        
        print(f"Данные {len(vehicles)} машин заполнены")
    
    def save(self, output_file=None):
        """Сохранение заполненной заявки"""
        if output_file is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"Заявка_пропуск_{timestamp}.xlsx"
        
        print(f"\nСохранение файла: {output_file}")
        self.workbook.save(output_file)
        print("Файл успешно сохранен")
        return output_file
    
    def validate(self):
        """Проверка заполненных данных на ошибки (с предупреждением о формулах)"""
        print("\nПроверка данных...")
        print("ВНИМАНИЕ: Excel файл содержит формулы валидации. ")
        print("Откройте файл в Excel для полной проверки на ошибки.")

def load_from_json(json_file):
    """Загрузка данных из JSON файла"""
    print(f"Загрузка данных из {json_file}...")
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

def load_from_config():
    """Загрузка данных из config.py файла"""
    try:
        import config
        print("Загрузка данных из config.py...")
        data = {
            'request_data': config.REQUEST_DATA,
            'visitors': config.VISITORS,
            'equipment': config.EQUIPMENT,
            'vehicles': config.VEHICLES
        }
        return data
    except ImportError:
        print("config.py не найден, используйте config_example.py как образец")
        return None
    except AttributeError as e:
        print(f"Ошибка в config.py: отсутствует {e}")
        return None

def create_example_json():
    """Создание примера JSON файла для ввода данных"""
    example_data = {
        "request_data": {
            "organization": "ООО \"Тестовая компания\"",
            "inn": "1234567890",
            "contract_number": "123",
            "contract_date": "10.11.2024",
            "contract_valid_until": "31.12.2025",
            "rooms": "Линия входных турникетов",
            "work_description": "зал 1, ряд 5, место 10",
            "work_start_date": "15.01.2025",
            "work_end_date": "20.01.2025",
            "purpose": "установка дополнительного оборудования",
            "responsible_person": "Генеральный директор",
            "responsible_name": "Иванов Иван Иванович",
            "request_number": "1245",
            "request_date": "10.11.2024",
            "contact_phone": "7(999)999-99-99"
        },
        "visitors": [
            {
                "surname": "Иванов",
                "name": "Иван",
                "patronymic": "Иванович",
                "birth_date": "01.01.1990",
                "citizenship": "",
                "passport_series": "1234",
                "passport_number": "567890",
                "passport_issue_date": "01.01.2015",
                "passport_issued_by": "Отделом УФМС России по г. Москве",
                "organization": "ООО \"Тестовая компания\"",
                "electrical_safety_group": "",
                "electrical_safety_certificate": "",
                "note": ""
            },
            {
                "surname": "Петров",
                "name": "Петр",
                "patronymic": "Петрович",
                "birth_date": "15.05.1985",
                "citizenship": "",
                "passport_series": "5678",
                "passport_number": "901234",
                "passport_issue_date": "15.05.2010",
                "passport_issued_by": "Отделом УФМС России по г. Москве",
                "organization": "ООО \"Тестовая компания\"",
                "electrical_safety_group": "",
                "electrical_safety_certificate": "",
                "note": ""
            }
        ],
        "equipment": [
            {
                "name": "Сервер",
                "brand": "HP",
                "model": "ProLiant DL380",
                "serial_number": "SN123456789",
                "quantity": "1",
                "unit": "шт.",
                "action": "Внос",
                "note": ""
            },
            {
                "name": "Коммутатор",
                "brand": "Cisco",
                "model": "Catalyst 2960",
                "serial_number": "FCW1234A567",
                "quantity": "2",
                "unit": "шт.",
                "action": "Внос",
                "note": ""
            }
        ],
        "vehicles": [
            {
                "license_plate": "A123BC777",
                "driver_surname": "Сидоров",
                "driver_name": "Сергей",
                "driver_patronymic": "Сергеевич",
                "driver_birth_date": "20.08.1988",
                "driver_citizenship": "",
                "passport_series": "9012",
                "passport_number": "345678",
                "passport_issue_date": "20.08.2013",
                "passport_issued_by": "Отделом УФМС России по г. Москве"
            }
        ]
    }
    
    with open('example_data.json', 'w', encoding='utf-8') as f:
        json.dump(example_data, f, ensure_ascii=False, indent=2)
    
    print("Создан файл example_data.json с примером данных")

def example_usage():
    """Пример использования"""
    print("=" * 80)
    print("АВТОМАТИЗАЦИЯ ЗАКАЗА ПРОПУСКОВ ММТС-9")
    print("=" * 80)
    
    # Пытаемся загрузить из config.py
    data = load_from_config()
    
    # Если config.py нет, пытаемся загрузить из JSON
    if data is None:
        # Создаем пример JSON если его нет
        if not os.path.exists('example_data.json'):
            create_example_json()
            print("\nТеперь отредактируйте example_data.json и запустите скрипт снова")
            return
        
        # Загружаем данные из JSON
        try:
            data = load_from_json('example_data.json')
        except FileNotFoundError:
            print("Файл example_data.json не найден")
            return
    
    # Создаем экземпляр класса и заполняем
    automation = PassRequestAutomation()
    automation.load_template()
    automation.fill_request_data(data['request_data'])
    automation.fill_visitors(data['visitors'])
    automation.fill_equipment(data['equipment'])
    automation.fill_vehicles(data['vehicles'])
    automation.validate()
    
    # Сохраняем
    output_file = automation.save()
    
    print("\n" + "=" * 80)
    print(f"[OK] ГОТОВО! Файл сохранен: {output_file}")
    print("=" * 80)
    print("\nСледующие шаги:")
    print("1. Откройте сохраненный файл в Excel")
    print("2. Проверьте данные на наличие ошибок (они будут подсвечены)")
    print("3. Отправьте заявку согласно процедуре ММТС-9")


if __name__ == '__main__':
    example_usage()

