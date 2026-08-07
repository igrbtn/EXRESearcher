# EXRESearcher

## Overview
GUI для поиска содержимого почтовых ящиков Exchange и массового удаления сообщений (фишинг/малвара). Search-Mailbox, In-Place eDiscovery, org-wide delete. Все асинхронно через runspaces.

## Quick Start
```powershell
# GUI
.\EXRESearcher.ps1

# Тесты
Invoke-Pester -Path ./tests/
```

## Architecture

### Модульная структура (lib/):
- `Core.ps1` — Exchange-функции: подключение, Search-Mailbox, eDiscovery, org-wide, статистика, folder cleanup, dumpster purge, дубликаты
- `Settings.ps1` — настройки, кэш, аудит-лог оператора
- `AsyncRunner.ps1` — async framework: runspaces, job tracker, progress bar, job console

### GUI (EXRESearcher.ps1) — 6 вкладок:
1. **Mailbox Search** — Search-Mailbox с KQL-фильтрами: subject, from, to, keywords, attachment, messageid, даты. Действия: Estimate / Log / Copy / Delete
2. **Org-Wide Delete** — поиск и удаление по ВСЕМ ящикам организации (батчами). Двойное подтверждение, safety check
3. **eDiscovery** — In-Place eDiscovery (New-MailboxSearch): создание, мониторинг, управление compliance-поисками
4. **Mailboxes** — браузер ящиков: фильтрация, статистика, folder stats, быстрая передача в Search
5. **Folder Cleanup** - поиск/удаление из папок с фильтрами (возраст, отправитель, тема, размер, вложения), purge dumpster, обнаружение дубликатов (EWS, группировка Subject+From+дата), backup/backup+delete через EWS CopyItem/MoveItem в бэкап-папку ТОГО ЖЕ ящика (Search-Mailbox не умеет folder-scope, EWS не умеет кросс-ящичное копирование)
6. **Audit Log** — лог поисков и лог операций оператора

### Async:
- Все Exchange-операции в runspaces (Start-AsyncJob), init (snap-in/remote session) внутри BeginInvoke
- Remote-режим: каждый job создаёт свою PSSession + Import-PSSession (прокси-командлеты живут per-runspace), сессия закрывается после job
- `$script:Session` - только маркер `@{ IsEMS; Server }`, живая сессия в GUI-runspace не хранится
- Callbacks (OnComplete/OnError) не видят локальные переменные завершившегося click-хендлера - контекст передаётся через `$script:JobCtx`
- Job console + progress bar внизу
- Comprehensive try/catch - ядро не падает

### EWS:
- FindItem не поддерживает Deep traversal - поиск по всему ящику идёт через FindFolder Deep + явный список папок в ParentFolderIds
- Порядок попыток Invoke-EwsRequest: явный Mailbox (Full Access) -> impersonation. Вариант "свой ящик" - только когда Mailbox не задан
- TLS: колбэк скомпилирован (Add-Type), self-signed доверяется по умолчанию; `EXRE_STRICT_TLS=1` включает строгую проверку

## Configuration
- `EXCHANGE_SERVER` — env variable для авто-заполнения
- Settings в `%APPDATA%\EXRESearcher\settings.json`

## Testing
- Framework: Pester v5
- `tests/EXRESearcher.Tests.ps1`

## Versioning
- Текущая версия: 1.4.0
