#!/bin/bash

# Остановить выполнение скрипта при ошибке
set -e

# Укажите имя приложения (замените на свое)
APP_NAME="fluck"

# Проверка, вошли ли в Heroku
echo "Logging into Heroku..."
heroku whoami &>/dev/null || heroku login

# Удаляем старый удаленный репозиторий Heroku (если он есть)
if git remote | grep -q "heroku"; then
    git remote remove heroku
fi

# Проверяем, существует ли приложение с таким именем
if ! heroku apps | grep -q "$APP_NAME"; then
    echo "Creating new Heroku app: $APP_NAME"
    heroku create "$APP_NAME"
fi

# Добавляем Heroku-репозиторий
git remote add heroku https://git.heroku.com/$APP_NAME.git

# Проверяем, инициализирован ли Git
if [ ! -d .git ]; then
    echo "Initializing Git repository..."
    git init
fi

# Добавляем файлы и делаем коммит
git add .
git commit -m "Deploying to Heroku"

# Убедимся, что находимся в правильной ветке
if ! git show-ref --quiet refs/heads/main; then
    echo "Renaming branch master to main..."
    git branch -M main
fi

# Отправляем код в Heroku
echo "Deploying to Heroku..."
git push heroku main

# Открываем приложение в браузере
echo "Opening the application..."
heroku open