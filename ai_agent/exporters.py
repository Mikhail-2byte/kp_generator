"""
Модуль для экспорта результатов аналитики в различные форматы.
Поддерживает экспорт в CSV, Excel и PNG (для графиков).
"""

import base64
import logging
import os
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from PIL import Image

logger = logging.getLogger(__name__)


class DataExporter:
    """Класс для экспорта данных в различные форматы."""
    
    # Директория для временных файлов
    TEMP_DIR = Path(tempfile.gettempdir()) / "kp_generator_exports"
    
    def __init__(self):
        """Инициализирует экспортер."""
        # Создаем директорию для временных файлов, если её нет
        self.TEMP_DIR.mkdir(parents=True, exist_ok=True)
    
    def export_to_csv(
        self,
        df: pd.DataFrame,
        filename: Optional[str] = None,
        include_index: bool = False
    ) -> Dict[str, Any]:
        """
        Экспортирует DataFrame в CSV файл.
        
        Args:
            df: DataFrame для экспорта
            filename: Имя файла (если None, генерируется автоматически)
            include_index: Включать ли индекс в экспорт
            
        Returns:
            Словарь с информацией о файле: {'path': str, 'filename': str, 'size': int}
        """
        try:
            if df.empty:
                raise ValueError("DataFrame пуст, экспорт невозможен")
            
            # Генерируем имя файла, если не указано
            if not filename:
                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"export_{timestamp}.csv"
            
            # Убеждаемся, что расширение .csv
            if not filename.endswith('.csv'):
                filename += '.csv'
            
            # Полный путь к файлу
            file_path = self.TEMP_DIR / filename
            
            # Экспортируем в CSV
            df.to_csv(
                file_path,
                index=include_index,
                encoding='utf-8-sig',  # UTF-8 с BOM для корректного отображения в Excel
                sep=','
            )
            
            file_size = file_path.stat().st_size
            
            logger.info(f"Экспортировано {len(df)} записей в CSV: {file_path}")
            
            return {
                'path': str(file_path),
                'filename': filename,
                'size': file_size,
                'format': 'csv',
                'records_count': len(df)
            }
        
        except Exception as e:
            logger.error(f"Ошибка при экспорте в CSV: {e}", exc_info=True)
            raise
    
    def export_to_excel(
        self,
        df: pd.DataFrame,
        filename: Optional[str] = None,
        sheet_name: str = 'Данные',
        include_index: bool = False
    ) -> Dict[str, Any]:
        """
        Экспортирует DataFrame в Excel файл.
        
        Args:
            df: DataFrame для экспорта
            filename: Имя файла (если None, генерируется автоматически)
            sheet_name: Название листа
            include_index: Включать ли индекс в экспорт
            
        Returns:
            Словарь с информацией о файле: {'path': str, 'filename': str, 'size': int}
        """
        try:
            if df.empty:
                raise ValueError("DataFrame пуст, экспорт невозможен")
            
            # Проверяем наличие openpyxl
            try:
                import openpyxl
            except ImportError:
                logger.error("Для экспорта в Excel требуется библиотека openpyxl. Установите: pip install openpyxl")
                raise ImportError("openpyxl не установлен. Установите: pip install openpyxl")
            
            # Генерируем имя файла, если не указано
            if not filename:
                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"export_{timestamp}.xlsx"
            
            # Убеждаемся, что расширение .xlsx
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'
            
            # Полный путь к файлу
            file_path = self.TEMP_DIR / filename
            
            # Экспортируем в Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=include_index,
                    engine='openpyxl'
                )
            
            file_size = file_path.stat().st_size
            
            logger.info(f"Экспортировано {len(df)} записей в Excel: {file_path}")
            
            return {
                'path': str(file_path),
                'filename': filename,
                'size': file_size,
                'format': 'xlsx',
                'records_count': len(df),
                'sheet_name': sheet_name
            }
        
        except Exception as e:
            logger.error(f"Ошибка при экспорте в Excel: {e}", exc_info=True)
            raise
    
    def export_chart_to_png(
        self,
        chart_base64: str,
        filename: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Экспортирует график (base64) в PNG файл.
        
        Args:
            chart_base64: Base64 строка изображения
            filename: Имя файла (если None, генерируется автоматически)
            
        Returns:
            Словарь с информацией о файле: {'path': str, 'filename': str, 'size': int}
        """
        try:
            if not chart_base64:
                raise ValueError("Base64 строка пуста")
            
            # Убираем префикс data URL, если есть
            if chart_base64.startswith('data:image'):
                # Извлекаем base64 часть
                chart_base64 = chart_base64.split(',', 1)[1] if ',' in chart_base64 else chart_base64
            
            # Декодируем base64
            image_data = base64.b64decode(chart_base64)
            
            # Генерируем имя файла, если не указано
            if not filename:
                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"chart_{timestamp}.png"
            
            # Убеждаемся, что расширение .png
            if not filename.endswith('.png'):
                filename += '.png'
            
            # Полный путь к файлу
            file_path = self.TEMP_DIR / filename
            
            # Сохраняем изображение
            image = Image.open(BytesIO(image_data))
            image.save(file_path, 'PNG')
            
            file_size = file_path.stat().st_size
            
            logger.info(f"Экспортирован график в PNG: {file_path}")
            
            return {
                'path': str(file_path),
                'filename': filename,
                'size': file_size,
                'format': 'png',
                'width': image.width,
                'height': image.height
            }
        
        except Exception as e:
            logger.error(f"Ошибка при экспорте графика в PNG: {e}", exc_info=True)
            raise
    
    def export_multiple_charts(
        self,
        charts: List[Dict[str, Any]],
        base_filename: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Экспортирует несколько графиков в отдельные PNG файлы.
        
        Args:
            charts: Список словарей с графиками {'title': str, 'image': str (base64)}
            base_filename: Базовое имя файла (если None, генерируется автоматически)
            
        Returns:
            Словарь с информацией о файлах: {'files': List[Dict], 'total_size': int}
        """
        try:
            if not charts:
                raise ValueError("Список графиков пуст")
            
            # Генерируем базовое имя, если не указано
            if not base_filename:
                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                base_filename = f"charts_{timestamp}"
            
            exported_files = []
            total_size = 0
            
            for i, chart in enumerate(charts, 1):
                title = chart.get('title', f'chart_{i}')
                image = chart.get('image', '')
                
                if not image:
                    continue
                
                # Создаем безопасное имя файла из заголовка
                safe_title = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in title)
                safe_title = safe_title.replace(' ', '_')[:50]  # Ограничиваем длину
                
                filename = f"{base_filename}_{safe_title}_{i}.png"
                
                result = self.export_chart_to_png(image, filename)
                exported_files.append(result)
                total_size += result['size']
            
            logger.info(f"Экспортировано {len(exported_files)} графиков")
            
            return {
                'files': exported_files,
                'total_size': total_size,
                'count': len(exported_files)
            }
        
        except Exception as e:
            logger.error(f"Ошибка при экспорте нескольких графиков: {e}", exc_info=True)
            raise
    
    def cleanup_old_files(self, max_age_hours: int = 24) -> int:
        """
        Удаляет старые временные файлы экспорта.
        
        Args:
            max_age_hours: Максимальный возраст файлов в часах
            
        Returns:
            Количество удаленных файлов
        """
        try:
            import datetime
            import time
            
            if not self.TEMP_DIR.exists():
                return 0
            
            cutoff_time = time.time() - (max_age_hours * 3600)
            deleted_count = 0
            
            for file_path in self.TEMP_DIR.iterdir():
                if file_path.is_file():
                    file_mtime = file_path.stat().st_mtime
                    if file_mtime < cutoff_time:
                        try:
                            file_path.unlink()
                            deleted_count += 1
                        except Exception as e:
                            logger.warning(f"Не удалось удалить файл {file_path}: {e}")
            
            if deleted_count > 0:
                logger.info(f"Удалено {deleted_count} старых файлов экспорта")
            
            return deleted_count
        
        except Exception as e:
            logger.error(f"Ошибка при очистке старых файлов: {e}", exc_info=True)
            return 0
    
    def get_file_info(self, file_path: str) -> Optional[Dict[str, Any]]:
        """
        Получает информацию о файле экспорта.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            Словарь с информацией о файле или None
        """
        try:
            path = Path(file_path)
            if not path.exists():
                return None
            
            stat = path.stat()
            return {
                'path': str(path),
                'filename': path.name,
                'size': stat.st_size,
                'created': stat.st_ctime,
                'modified': stat.st_mtime,
                'exists': True
            }
        except Exception as e:
            logger.error(f"Ошибка при получении информации о файле: {e}")
            return None
    
    def delete_file(self, file_path: str) -> bool:
        """
        Удаляет файл экспорта.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            True если файл удален успешно
        """
        try:
            path = Path(file_path)
            if path.exists() and path.is_file():
                path.unlink()
                logger.info(f"Удален файл экспорта: {file_path}")
                return True
            return False
        except Exception as e:
            logger.error(f"Ошибка при удалении файла: {e}")
            return False

