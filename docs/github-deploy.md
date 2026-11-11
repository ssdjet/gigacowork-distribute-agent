# Автоматическая загрузка из GitHub в GigaCowork без архива

Скрипт `deploy_to_gigacowork.py` скачивает актуальные файлы агента напрямую из
GitHub, заполняет форму создания агента и загружает `SKILL.md`, `scripts` и
`references` через стандартное файловое поле GigaCowork. ZIP не создаётся.

> Проверено скачивание из GitHub и сквозной сценарий на модели формы GigaCowork.
> Реальный запуск в рабочем пространстве требует вашей авторизации и отдельной
> проверки: из внешней тестовой среды рабочая ссылка возвращает 403 Forbidden.

## Что потребуется

- Компьютер с Google Chrome.
- Python 3.10 или новее. В установщике Windows включите **Add python.exe to PATH**.
- Git.
- Доступ к рабочему пространству GigaCowork.

Ссылка на создание агента уже записана в скрипте:

```text
https://way.cowork.ru/workspaces/64d8061e-0468-4522-b884-69e225f39caf/agents/new
```

## Пошаговый запуск

### 1. Откройте терминал

В Windows нажмите **Win + R**, введите `powershell` и нажмите Enter.

Проверьте программы:

```powershell
python --version
git --version
```

Если команда `python` не найдена, закройте и заново откройте терминал после
установки Python. Git можно установить с https://git-scm.com/download/win.

### 2. Получите проект из GitHub

```powershell
git clone https://github.com/ssdjet/gigacowork-distribute-agent.git
cd gigacowork-distribute-agent
```

Архив скачивать и распаковывать не нужно.

### 3. Установите браузерную зависимость

```powershell
python -m pip install -r requirements-deploy.txt
```

Рабочие зависимости агента (`openpyxl`) на компьютер для загрузки не требуются.

### 4. Запустите загрузчик

```powershell
python deploy_to_gigacowork.py
```

Скрипт автоматически:

1. Скачает из ветки `main` четыре файла:
   - `skill/SKILL.md`;
   - `skill/scripts/distribute.py`;
   - `skill/references/rules.md`;
   - `system_prompt.txt`.
2. Проверит GitHub blob SHA каждого файла.
3. Откроет Chrome и страницу создания агента.
4. Заполнит название, описание и инструкцию.
5. Нажмёт «+» в блоке «Навыки» и передаст три файла навыка без архива.
6. Нажмёт «Сохранить».

### 5. Один раз войдите в GigaCowork

При первом запуске войдите в открывшемся Chrome и дождитесь формы создания
агента. Затем вернитесь в PowerShell и нажмите Enter.

Сессия сохраняется в локальном профиле:

```text
%USERPROFILE%\.gigacowork-deploy\chrome-profile
```

Профиль не отправляется в GitHub. Не передавайте эту папку другим людям:
в ней может быть активная сессия GigaCowork. Логин и пароль в код и логи
скрипт не записывает.

### 6. Проверьте результат

Успешное завершение:

```text
DEPLOY=OK: агент сохранён в GigaCowork.
```

Откройте созданного агента «Распределение ФУЛ между логистами», приложите три
Excel-файла и отправьте пустое сообщение. Проверка должна подтвердить пять
листов в `result.xlsx` и счётчики из `docs/installation.md`.

Каждый запуск со ссылкой `/agents/new` создаёт нового агента. Для повторного
теста удалите неудачный черновик в интерфейсе или используйте другое название:

```powershell
python deploy_to_gigacowork.py --agent-name "Распределение ФУЛ — тест 2"
```

## Обновление файлов перед загрузкой

```powershell
cd gigacowork-distribute-agent
git pull origin main
python -m pip install -r requirements-deploy.txt
python deploy_to_gigacowork.py
```

Можно загрузить конкретную версию по коммиту:

```powershell
python deploy_to_gigacowork.py --ref 4d3aadb1568b4cbc0ef2bd22f66d71426013d44d
```

## Полезные режимы

Только скачать и проверить файлы, не открывая браузер:

```powershell
python deploy_to_gigacowork.py --dry-run
```

Только сохранить сессию GigaCowork:

```powershell
python deploy_to_gigacowork.py --login-only
```

Оставить браузер открытым после сохранения:

```powershell
python deploy_to_gigacowork.py --keep-open
```

## Если загрузка не удалась

- **403 Forbidden:** проверьте, что ссылка открывается в вашем Chrome и у вас
  есть доступ именно к этому рабочему пространству.
- **Chrome не найден:** установите Chrome или выполните
  `python -m playwright install chromium`, затем повторите команду.
- **Форма принимает только один файл:** скрипт остановится намеренно. Загружать
  только `SKILL.md` нельзя — навык останется без скрипта и справочника. Сделайте
  снимок файловой формы для адаптации загрузчика.
- **Кнопка «Сохранить» неактивна:** дождитесь завершения загрузки файлов и
  проверьте обязательные поля.
- **Изменилась разметка страницы:** скрипт сохранит `gigacowork-deploy-error.png`
  в текущей папке и прекратит работу, не пытаясь нажимать элементы вслепую.

После `DEPLOY=FAILED` агент не считается установленным. Проверьте сообщение и
снимок ошибки; пароли и токены публиковать вместе со снимком не нужно.
