from flask_wtf import FlaskForm
from wtforms import (
    StringField,
    PasswordField,
    SubmitField,
    BooleanField,
    HiddenField,
    TextAreaField,
    DecimalField,
    IntegerField,
    SelectField
)
from wtforms.validators import DataRequired, Length, EqualTo, Optional, NumberRange


class LoginForm(FlaskForm):
    """Форма авторизации пользователей на сайте."""
    username = StringField('Логин', validators=[DataRequired(), Length(min=3, max=50)])
    password = PasswordField('Пароль', validators=[DataRequired(), Length(min=6, max=128)])
    remember_me = BooleanField('Запомнить меня')
    submit_login = SubmitField('Войти')


class RegistrationForm(FlaskForm):
    """Форма регистрации с обязательным вводом профиля и пароля."""
    username = StringField('Логин', validators=[DataRequired(), Length(min=3, max=50)])
    last_name = StringField('Фамилия', validators=[DataRequired(), Length(min=2, max=100)])
    first_name = StringField('Имя', validators=[DataRequired(), Length(min=2, max=100)])
    password = PasswordField('Пароль', validators=[DataRequired(), Length(min=6, max=128)])
    confirm_password = PasswordField(
        'Подтвердите пароль',
        validators=[DataRequired(), EqualTo('password', message='Пароли должны совпадать')]
    )
    submit_register = SubmitField('Зарегистрироваться')


class ProfileUpdateForm(FlaskForm):
    """Форма редактирования данных профиля и смены пароля."""
    username = StringField('Логин', validators=[DataRequired(), Length(min=3, max=50)])
    last_name = StringField('Фамилия', validators=[DataRequired(), Length(min=2, max=100)])
    first_name = StringField('Имя', validators=[DataRequired(), Length(min=2, max=100)])
    new_password = PasswordField('Новый пароль', validators=[Optional(), Length(min=6, max=128)])
    confirm_new_password = PasswordField(
        'Подтвердите новый пароль',
        validators=[Optional(), EqualTo('new_password', message='Пароли должны совпадать')]
    )
    submit_update = SubmitField('Сохранить изменения')


class ContactInfoForm(FlaskForm):
    """Форма редактирования контактной информации пользователя."""
    contact_info = TextAreaField('Контактная информация', validators=[Optional(), Length(max=1000)])
    submit = SubmitField('Сохранить')


class DeleteAccountForm(FlaskForm):
    """Подтверждение удаления аккаунта со всеми данными."""
    confirm_delete = BooleanField(
        'Я понимаю, что удаление аккаунта необратимо',
        validators=[DataRequired()]
    )
    submit_delete = SubmitField('Удалить аккаунт')


class DutyItemForm(FlaskForm):
    """Добавление новой ставки пошлины в справочник."""
    product = StringField('Наименование товара', validators=[DataRequired(), Length(min=1, max=200)])
    category = StringField('Категория', validators=[DataRequired(), Length(min=1, max=200)])
    duty_percent = DecimalField(
        'Пошлина, %', places=2, rounding=None, validators=[DataRequired(), NumberRange(min=0)]
    )
    action = HiddenField(default='add_duty')
    submit = SubmitField('Добавить позицию')


class DutyDeleteForm(FlaskForm):
    """Удаление существующей записи о пошлине."""
    action = HiddenField(default='delete_duty', validators=[DataRequired()])
    index = HiddenField(validators=[DataRequired()])
    submit = SubmitField('Удалить')


class GBMaterialForm(FlaskForm):
    """Добавление аналога материала по стандарту GB."""
    russian = StringField('Наименование (RU)', validators=[DataRequired(), Length(min=1, max=200)])
    gb = StringField('Наименование (GB)', validators=[DataRequired(), Length(min=1, max=200)])
    notes = StringField('Описание', validators=[Optional(), Length(max=500)])
    composition = TextAreaField('Состав (формат: элемент: значение, каждое с новой строки)', validators=[Optional()])
    action = HiddenField(default='add_gb')
    submit = SubmitField('Добавить аналог')


class GBMaterialDeleteForm(FlaskForm):
    """Удаление аналога материала из справочника."""
    action = HiddenField(default='delete_gb', validators=[DataRequired()])
    index = HiddenField(validators=[DataRequired()])
    submit = SubmitField('Удалить')


class LogisticsCityForm(FlaskForm):
    """Добавление города с тарифами перевозки."""
    name = StringField('Город', validators=[DataRequired(), Length(min=1, max=200)])
    region = StringField('Регион', validators=[Optional(), Length(max=200)])
    truck_price = DecimalField(
        'Цена фуры, руб', places=2, rounding=None, validators=[DataRequired(), NumberRange(min=0)]
    )
    trail_price = DecimalField(
        'Цена трала, руб', places=2, rounding=None, validators=[Optional(), NumberRange(min=0)]
    )
    is_main_route = BooleanField('Основной маршрут', default=False)
    allows_trail = BooleanField('Доступен трал', default=False)
    distance_from_ekb_km = IntegerField(
        'Расстояние от ЕКБ, км', validators=[Optional(), NumberRange(min=0)]
    )
    action = HiddenField(default='add_city')
    submit = SubmitField('Добавить город')


class MainCityForm(FlaskForm):
    """Форма для основных городов и городов в радиусе 300км."""
    name = StringField('Город', validators=[DataRequired(), Length(min=1, max=200)])
    region = StringField('Регион', validators=[Optional(), Length(max=200)])
    truck_price = DecimalField(
        'Цена фуры, руб', places=2, rounding=None, validators=[DataRequired(), NumberRange(min=0)]
    )
    is_main_route = BooleanField('Основной город', default=False)
    main_city = StringField(
        'Привязан к основному городу', 
        validators=[Optional(), Length(max=200)],
        description='Укажите название основного города, если это город в радиусе 300км'
    )
    action = HiddenField(default='add_main_city')
    submit = SubmitField('Добавить город')


class EkbRfCityForm(FlaskForm):
    """Форма для городов за 300км (алгоритм ЕКБ+РФ)."""
    name = StringField('Город', validators=[DataRequired(), Length(min=1, max=200)])
    region = StringField('Регион', validators=[Optional(), Length(max=200)])
    distance_from_ekb_km = IntegerField(
        'Расстояние от ЕКБ, км', validators=[Optional(), NumberRange(min=0)]
    )
    action = HiddenField(default='add_ekb_rf_city')
    submit = SubmitField('Добавить город')


class TrailCityForm(FlaskForm):
    """Форма для справочника трала."""
    name = StringField('Город', validators=[DataRequired(), Length(min=1, max=200)])
    region = StringField('Регион', validators=[Optional(), Length(max=200)])
    trail_price = DecimalField(
        'Цена трала, руб', places=2, rounding=None, validators=[DataRequired(), NumberRange(min=0)]
    )
    action = HiddenField(default='add_trail_city')
    submit = SubmitField('Добавить город')


class LogisticsCityDeleteForm(FlaskForm):
    """Удаление города из справочника логистики."""
    action = HiddenField(default='delete_city', validators=[DataRequired()])
    index = HiddenField(validators=[DataRequired()])
    submit = SubmitField('Удалить')


class MainCityDeleteForm(FlaskForm):
    """Удаление города из справочника основных городов."""
    action = HiddenField(default='delete_main_city', validators=[DataRequired()])
    index = HiddenField(validators=[DataRequired()])
    submit = SubmitField('Удалить')


class EkbRfCityDeleteForm(FlaskForm):
    """Удаление города из справочника ЕКБ+РФ."""
    action = HiddenField(default='delete_ekb_rf_city', validators=[DataRequired()])
    index = HiddenField(validators=[DataRequired()])
    submit = SubmitField('Удалить')


class TrailCityDeleteForm(FlaskForm):
    """Удаление города из справочника трала."""
    action = HiddenField(default='delete_trail_city', validators=[DataRequired()])
    index = HiddenField(validators=[DataRequired()])
    submit = SubmitField('Удалить')


class AdminUserForm(FlaskForm):
    """Форма создания/редактирования пользователя администратором."""
    username = StringField('Логин', validators=[DataRequired(), Length(min=3, max=50)])
    last_name = StringField('Фамилия', validators=[Optional(), Length(max=100)])
    first_name = StringField('Имя', validators=[Optional(), Length(max=100)])
    contact_info = TextAreaField('Контактная информация', validators=[Optional(), Length(max=1000)])
    role = SelectField('Роль', choices=[('user', 'Пользователь'), ('admin', 'Администратор')], validators=[DataRequired()])
    password = PasswordField('Пароль', validators=[Optional(), Length(min=6, max=128)])
    confirm_password = PasswordField(
        'Подтвердите пароль',
        validators=[Optional(), EqualTo('password', message='Пароли должны совпадать')]
    )
    submit = SubmitField('Сохранить')


class AdminUserDeleteForm(FlaskForm):
    """Форма удаления пользователя."""
    user_id = HiddenField(validators=[DataRequired()])
    confirm_delete = BooleanField(
        'Подтвердить удаление',
        validators=[DataRequired()]
    )
    submit = SubmitField('Удалить пользователя')


class AdminResetPasswordForm(FlaskForm):
    """Форма сброса пароля пользователя."""
    user_id = HiddenField(validators=[DataRequired()])
    new_password = PasswordField('Новый пароль', validators=[DataRequired(), Length(min=6, max=128)])
    confirm_password = PasswordField(
        'Подтвердите пароль',
        validators=[DataRequired(), EqualTo('new_password', message='Пароли должны совпадать')]
    )
    submit = SubmitField('Сбросить пароль')


class CustomerContactForm(FlaskForm):
    """Форма для создания и редактирования контакта заказчика."""
    company_name = StringField('Название компании', validators=[DataRequired(), Length(min=1, max=200)])
    contact_person = StringField('Контактное лицо', validators=[Optional(), Length(max=200)])
    phone = StringField('Телефон', validators=[Optional(), Length(max=50)])
    email = StringField('Email', validators=[Optional(), Length(max=200)])
    address = TextAreaField('Адрес', validators=[Optional(), Length(max=500)])
    notes = TextAreaField('Заметки', validators=[Optional(), Length(max=1000)])
    submit = SubmitField('Сохранить')


class CustomerContactDeleteForm(FlaskForm):
    """Форма удаления контакта заказчика."""
    contact_id = HiddenField(validators=[DataRequired()])
    confirm_delete = BooleanField(
        'Подтвердить удаление',
        validators=[DataRequired()]
    )
    submit = SubmitField('Удалить контакт')