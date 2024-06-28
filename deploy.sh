#!/bin/bash

# Убедитесь, что скрипт остановится при любой ошибке
set -e

# Инициализация Git репозитория (если не сделано ранее)
if [ ! -d .git ]; then
    git init
fi

# Добавление всех файлов и создание коммита
git add .
git commit -m "Deploy commit"

# Переименование ветки master в main (если не сделано ранее)
if ! git show-ref --quiet refs/heads/main; then
    git branch -M main
fi

# Вход в Heroku
heroku login

# Создание нового приложения на Heroku (если не сделано ранее)
APP_NAME="fluck"
if ! heroku apps | grep -q "$APP_NAME"; then
    heroku create "$APP_NAME"
fi

# Загрузка приложения на Heroku
git push heroku main

# Открытие приложения в браузере
heroku open
