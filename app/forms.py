from flask_wtf import FlaskForm
from wtforms import (
    StringField,
    PasswordField,
    SubmitField,
    BooleanField,
    HiddenField,
    TextAreaField,
    DecimalField
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
    contact_info = TextAreaField('Контактная информация', validators=[Optional(), Length(max=1000)])
    new_password = PasswordField('Новый пароль', validators=[Optional(), Length(min=6, max=128)])
    confirm_new_password = PasswordField(
        'Подтвердите новый пароль',
        validators=[Optional(), EqualTo('new_password', message='Пароли должны совпадать')]
    )
    submit_update = SubmitField('Сохранить изменения')


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
        'Цена трала, руб', places=2, rounding=None, validators=[DataRequired(), NumberRange(min=0)]
    )
    action = HiddenField(default='add_city')
    submit = SubmitField('Добавить город')


class LogisticsCityDeleteForm(FlaskForm):
    """Удаление города из справочника логистики."""
    action = HiddenField(default='delete_city', validators=[DataRequired()])
    index = HiddenField(validators=[DataRequired()])
    submit = SubmitField('Удалить')
