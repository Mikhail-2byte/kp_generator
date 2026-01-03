#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Скрипт для управления Flask приложением (запуск, остановка, перезапуск, статус).

Использование:
    python scripts/manage_app.py status    # Проверить статус
    python scripts/manage_app.py stop      # Остановить приложение
    python scripts/manage_app.py restart   # Перезапустить приложение
    python scripts/manage_app.py start     # Запустить приложение
"""

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

# Настраиваем кодировку UTF-8 для Windows
if sys.platform == 'win32':
    try:
        import codecs
        if sys.stdout.encoding != 'utf-8':
            sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
            sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')
    except Exception:
        pass

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_PORT = 5000


def find_process_on_port(port: int) -> list:
    """Находит процессы, использующие указанный порт."""
    try:
        result = subprocess.run(
            ['netstat', '-ano'],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        processes = []
        for line in result.stdout.split('\n'):
            if f':{port}' in line and 'LISTENING' in line:
                parts = line.split()
                if len(parts) >= 5:
                    pid = parts[-1]
                    processes.append(pid)
        return processes
    except Exception as e:
        print(f"[ERROR] Ошибка при поиске процессов: {e}")
        return []


def get_process_info(pid: str) -> dict:
    """Получает информацию о процессе по PID."""
    try:
        result = subprocess.run(
            ['tasklist', '/FI', f'PID eq {pid}', '/FO', 'CSV', '/NH'],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        if result.stdout.strip():
            parts = result.stdout.strip().split('","')
            if len(parts) >= 2:
                return {
                    'pid': pid,
                    'name': parts[0].strip('"'),
                    'memory': parts[4].strip('"') if len(parts) > 4 else 'N/A'
                }
    except Exception:
        pass
    
    return {'pid': pid, 'name': 'Unknown', 'memory': 'N/A'}


def stop_app(port: int = DEFAULT_PORT) -> bool:
    """Останавливает приложение на указанном порту."""
    processes = find_process_on_port(port)
    
    if not processes:
        print(f"[INFO] Приложение не запущено на порту {port}")
        return True
    
    print(f"[INFO] Найдено процессов на порту {port}: {len(processes)}")
    
    for pid in processes:
        info = get_process_info(pid)
        print(f"[INFO] Остановка процесса {pid} ({info['name']})...")
        
        try:
            subprocess.run(['taskkill', '/F', '/PID', pid], check=True, capture_output=True)
            print(f"[OK] Процесс {pid} остановлен")
        except subprocess.CalledProcessError as e:
            print(f"[ERROR] Не удалось остановить процесс {pid}: {e}")
            return False
        except Exception as e:
            print(f"[ERROR] Ошибка при остановке процесса {pid}: {e}")
            return False
    
    # Ждем немного, чтобы порт освободился
    time.sleep(1)
    
    # Проверяем, что порт освобожден
    remaining = find_process_on_port(port)
    if remaining:
        print(f"[WARNING] Порт {port} все еще занят процессами: {remaining}")
        return False
    
    print(f"[OK] Приложение остановлено, порт {port} освобожден")
    return True


def start_app(port: int = DEFAULT_PORT) -> bool:
    """Запускает приложение."""
    # Проверяем, не занят ли порт
    existing = find_process_on_port(port)
    if existing:
        print(f"[ERROR] Порт {port} уже занят процессами: {existing}")
        print("[INFO] Используйте 'restart' для перезапуска или 'stop' для остановки")
        return False
    
    print(f"[INFO] Запуск приложения на порту {port}...")
    print(f"[INFO] Рабочая директория: {BASE_DIR}")
    
    # Проверяем наличие app.py
    app_file = BASE_DIR / 'app.py'
    if not app_file.exists():
        print(f"[ERROR] Файл app.py не найден в {BASE_DIR}")
        return False
    
    try:
        # Запускаем в фоновом режиме
        os.chdir(BASE_DIR)
        
        # Для Windows используем start для запуска в новом окне
        # или можно использовать subprocess.Popen для фонового запуска
        print("[INFO] Запуск приложения...")
        print("[INFO] Приложение будет запущено в новом окне терминала")
        print("[INFO] Для остановки используйте: python scripts/manage_app.py stop")
        
        # Запускаем в новом окне (Windows)
        if sys.platform == 'win32':
            subprocess.Popen(
                ['python', 'app.py'],
                cwd=str(BASE_DIR),
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
        else:
            # Для Linux/Mac
            subprocess.Popen(
                ['python', 'app.py'],
                cwd=str(BASE_DIR),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        
        # Ждем немного и проверяем
        time.sleep(2)
        processes = find_process_on_port(port)
        
        if processes:
            print(f"[OK] Приложение запущено! PID: {processes[0]}")
            print(f"[INFO] Откройте в браузере: http://localhost:{port}")
            return True
        else:
            print("[WARNING] Приложение запущено, но порт еще не открыт (может быть в процессе запуска)")
            return True
            
    except Exception as e:
        print(f"[ERROR] Ошибка при запуске приложения: {e}")
        return False


def status_app(port: int = DEFAULT_PORT) -> bool:
    """Проверяет статус приложения."""
    processes = find_process_on_port(port)
    
    if not processes:
        print(f"[INFO] Приложение не запущено на порту {port}")
        return False
    
    print(f"[INFO] Приложение запущено на порту {port}")
    print(f"[INFO] Найдено процессов: {len(processes)}")
    
    for pid in processes:
        info = get_process_info(pid)
        print(f"  PID: {info['pid']}")
        print(f"  Имя: {info['name']}")
        print(f"  Память: {info['memory']}")
        print()
    
    print(f"[INFO] Приложение доступно по адресу: http://localhost:{port}")
    return True


def restart_app(port: int = DEFAULT_PORT) -> bool:
    """Перезапускает приложение."""
    print("[INFO] Перезапуск приложения...")
    
    # Останавливаем
    if not stop_app(port):
        print("[ERROR] Не удалось остановить приложение")
        return False
    
    # Ждем немного
    time.sleep(1)
    
    # Запускаем
    if not start_app(port):
        print("[ERROR] Не удалось запустить приложение")
        return False
    
    print("[OK] Приложение перезапущено")
    return True


def main():
    """Главная функция."""
    parser = argparse.ArgumentParser(
        description="Управление Flask приложением KP Generator"
    )
    parser.add_argument(
        'action',
        choices=['start', 'stop', 'restart', 'status'],
        help='Действие: start (запуск), stop (остановка), restart (перезапуск), status (статус)'
    )
    parser.add_argument(
        '--port',
        type=int,
        default=DEFAULT_PORT,
        help=f'Порт приложения (по умолчанию: {DEFAULT_PORT})'
    )
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("  KP Generator - Управление приложением")
    print("=" * 60)
    print()
    
    if args.action == 'status':
        success = status_app(args.port)
        sys.exit(0 if success else 1)
    elif args.action == 'stop':
        success = stop_app(args.port)
        sys.exit(0 if success else 1)
    elif args.action == 'start':
        success = start_app(args.port)
        sys.exit(0 if success else 1)
    elif args.action == 'restart':
        success = restart_app(args.port)
        sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()

