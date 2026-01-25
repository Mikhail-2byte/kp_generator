# Скрипт для проверки и запуска Redis через Docker

Write-Host "Проверка Docker Desktop..." -ForegroundColor Cyan

# Проверяем, запущен ли Docker
$dockerRunning = docker ps 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "Docker Desktop запущен!" -ForegroundColor Green
    
    # Проверяем, запущен ли контейнер Redis
    $redisContainer = docker ps -a --filter "name=redis" --format "{{.Names}}" 2>$null
    if ($redisContainer -eq "redis") {
        Write-Host "Контейнер Redis найден. Проверяем статус..." -ForegroundColor Yellow
        $redisRunning = docker ps --filter "name=redis" --format "{{.Names}}" 2>$null
        if ($redisRunning -eq "redis") {
            Write-Host "Redis уже запущен!" -ForegroundColor Green
        } else {
            Write-Host "Запускаем контейнер Redis..." -ForegroundColor Yellow
            docker start redis
        }
    } else {
        Write-Host "Создаем и запускаем контейнер Redis..." -ForegroundColor Yellow
        docker run -d -p 6379:6379 --name redis redis:latest
    }
    
    # Проверяем подключение
    Start-Sleep -Seconds 2
    $redisCheck = docker exec redis redis-cli ping 2>$null
    if ($redisCheck -eq "PONG") {
        Write-Host "Redis работает! Ответ: $redisCheck" -ForegroundColor Green
    } else {
        Write-Host "Redis контейнер запущен, но не отвечает. Подождите несколько секунд." -ForegroundColor Yellow
    }
} else {
    Write-Host "Docker Desktop не запущен или не установлен." -ForegroundColor Red
    Write-Host ""
    Write-Host "Варианты решения:" -ForegroundColor Yellow
    Write-Host "1. Запустите Docker Desktop вручную и подождите 30-60 секунд"
    Write-Host "2. Используйте приложение без Redis (fallback на cookie-based сессии)"
    Write-Host "3. Установите Redis через WSL2 или Memurai (см. docs/REDIS_SETUP.md)"
    Write-Host ""
    Write-Host "Приложение будет работать без Redis, но с предупреждениями о размере cookie." -ForegroundColor Cyan
}

