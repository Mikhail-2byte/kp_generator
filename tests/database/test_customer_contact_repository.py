"""Тесты для CustomerContactRepository."""

import sys
from pathlib import Path

# Добавляем корень проекта в путь для возможности прямого запуска
project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))

import pytest

from app.services.repositories import customer_contact_repository


class TestCustomerContactRepository:
    """Тесты для CustomerContactRepository."""
    
    def test_create_contact(self):
        """Тест создания контакта."""
        contact_id = customer_contact_repository.create(
            company_name='Тестовая компания',
            contact_person='Иван Иванов',
            phone='+7 999 123-45-67',
            email='test@example.com',
            address='Москва, ул. Тестовая, 1',
            notes='Тестовые заметки'
        )
        
        assert contact_id is not None
        assert isinstance(contact_id, int)
    
    def test_get_by_id(self):
        """Тест получения контакта по ID."""
        # Создаем контакт
        contact_id = customer_contact_repository.create(
            company_name='Компания для получения',
            contact_person='Петр Петров'
        )
        
        # Получаем контакт
        contact = customer_contact_repository.get_by_id(contact_id)
        
        assert contact is not None
        assert contact['id'] == contact_id
        assert contact['company_name'] == 'Компания для получения'
    
    def test_get_by_id_not_found(self):
        """Тест получения несуществующего контакта."""
        contact = customer_contact_repository.get_by_id(999999)
        assert contact is None
    
    def test_get_all(self):
        """Тест получения всех контактов."""
        # Создаем несколько контактов
        customer_contact_repository.create(company_name='Компания 1')
        customer_contact_repository.create(company_name='Компания 2')
        
        contacts = customer_contact_repository.get_all()
        
        assert isinstance(contacts, list)
        assert len(contacts) >= 2
    
    def test_get_all_with_search(self):
        """Тест поиска контактов."""
        import time
        # Создаем контакт с уникальным названием
        unique_name = f'Уникальная компания {int(time.time())}'
        customer_contact_repository.create(
            company_name=unique_name,
            contact_person='Тестовый контакт'
        )
        
        contacts = customer_contact_repository.get_all(search='Уникальная')
        
        assert len(contacts) > 0
        assert any(c['company_name'] == unique_name for c in contacts)
    
    def test_update_contact(self):
        """Тест обновления контакта."""
        # Создаем контакт
        contact_id = customer_contact_repository.create(
            company_name='Компания для обновления',
            phone='+7 111 111-11-11'
        )
        
        # Обновляем контакт
        success = customer_contact_repository.update(
            contact_id,
            phone='+7 222 222-22-22',
            email='updated@example.com'
        )
        
        assert success is True
        
        # Проверяем обновление
        contact = customer_contact_repository.get_by_id(contact_id)
        assert contact['phone'] == '+7 222 222-22-22'
        assert contact['email'] == 'updated@example.com'
    
    def test_update_contact_not_found(self):
        """Тест обновления несуществующего контакта."""
        success = customer_contact_repository.update(999999, phone='+7 999 999-99-99')
        assert success is False
    
    def test_delete_contact(self):
        """Тест удаления контакта."""
        # Создаем контакт
        contact_id = customer_contact_repository.create(
            company_name='Компания для удаления'
        )
        
        # Удаляем контакт
        success = customer_contact_repository.delete(contact_id)
        
        assert success is True
        
        # Проверяем, что контакт удален
        contact = customer_contact_repository.get_by_id(contact_id)
        assert contact is None
    
    def test_delete_contact_not_found(self):
        """Тест удаления несуществующего контакта."""
        success = customer_contact_repository.delete(999999)
        assert success is False


if __name__ == "__main__":
    # Запуск тестов при прямом выполнении файла
    pytest.main([__file__, "-v"])

