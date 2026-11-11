#!/usr/bin/env python3
"""Автоматическая загрузка агента из GitHub в GigaCowork без ZIP-архива.

Перед запуском: python -m pip install -r requirements-deploy.txt
При первом запуске потребуется войти в GigaCowork в открывшемся Chrome.
"""

import argparse
import base64
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import urllib.parse
import urllib.request
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

REPO_API = "https://api.github.com/repos/ssdjet/gigacowork-distribute-agent/contents/"
DEFAULT_URL = (
    "https://way.cowork.ru/workspaces/"
    "64d8061e-0468-4522-b884-69e225f39caf/agents/new"
)
AGENT_NAME = "Распределение ФУЛ между логистами"
AGENT_DESCRIPTION = (
    "Распределяет новые ФУЛ поставщиков между логистами, учитывает историю Cloud, "
    "загрузку и доступность сотрудников и формирует result.xlsx."
)
AGENT_FILES = (
    "skill/SKILL.md",
    "skill/scripts/distribute.py",
    "skill/references/rules.md",
    "system_prompt.txt",
)


def download_files(destination, ref="main"):
    """Скачать отдельные файлы через GitHub API и проверить их blob SHA."""
    downloaded = []
    for relative in AGENT_FILES:
        url = REPO_API + urllib.parse.quote(relative) + "?ref=" + urllib.parse.quote(ref)
        request = urllib.request.Request(
            url,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "gigacowork-github-deploy",
            },
        )
        with urllib.request.urlopen(request, timeout=45) as response:
            payload = json.load(response)
        if payload.get("type") != "file" or payload.get("encoding") != "base64":
            raise RuntimeError("GitHub вернул неожиданный ответ для " + relative)
        data = base64.b64decode(payload["content"])
        actual_sha = hashlib.sha1(
            b"blob " + str(len(data)).encode("ascii") + b"\0" + data
        ).hexdigest()
        if actual_sha != payload.get("sha"):
            raise RuntimeError("Не совпала контрольная сумма файла " + relative)
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        downloaded.append(target)
        print("Скачан и проверен: " + relative)
    return downloaded


def validate_files(root):
    skill_md = (root / "skill/SKILL.md").read_text(encoding="utf-8")
    if not skill_md.startswith("---") or "name: distribute-orders" not in skill_md:
        raise RuntimeError("SKILL.md не содержит ожидаемый заголовок distribute-orders")
    if "description:" not in skill_md:
        raise RuntimeError("В SKILL.md отсутствует description")
    if not (root / "system_prompt.txt").read_text(encoding="utf-8").strip():
        raise RuntimeError("Системный промпт пуст")


def first_label_field(page, label):
    return page.get_by_label(re.compile(r"^\s*" + re.escape(label), re.I)).first


def fill_field(page, label, value, placeholder=None):
    field = first_label_field(page, label)
    try:
        field.wait_for(state="visible", timeout=2500)
        field.fill(value)
        return
    except PlaywrightTimeoutError:
        pass
    if placeholder:
        fallback = page.locator(
            "input[placeholder*='" + placeholder + "'], "
            "textarea[placeholder*='" + placeholder + "']"
        ).first
        fallback.wait_for(state="visible", timeout=5000)
        fallback.fill(value)
        return
    raise RuntimeError("Не найдено поле «" + label + "»")


def wait_for_form(page, workspace_url):
    page.goto(workspace_url, wait_until="domcontentloaded", timeout=60000)
    title = page.title()
    body = page.locator("body").inner_text(timeout=5000)
    if "403 Forbidden" in title or "403 Forbidden" in body:
        raise RuntimeError(
            "GigaCowork вернул 403 Forbidden. Откройте ссылку в своей сети "
            "и проверьте доступ к рабочему пространству."
        )
    try:
        first_label_field(page, "Название").wait_for(state="visible", timeout=12000)
        return
    except PlaywrightTimeoutError:
        pass
    print("\nТребуется вход в GigaCowork.")
    print("Войдите в открывшемся Chrome и дождитесь формы создания агента.")
    if not sys.stdin.isatty():
        raise RuntimeError("Вход не выполнен, а интерактивный терминал недоступен")
    input("Когда форма откроется, вернитесь в терминал и нажмите Enter…")
    if "/agents/new" not in page.url:
        page.goto(workspace_url, wait_until="domcontentloaded", timeout=60000)
    first_label_field(page, "Название").wait_for(state="visible", timeout=30000)


def skills_add_button(page):
    heading = page.get_by_text("Навыки", exact=True).first
    heading.wait_for(state="visible", timeout=10000)
    block = heading.locator("xpath=ancestor::*[.//button][1]")
    buttons = block.locator("button")
    count = buttons.count()
    if not count:
        raise RuntimeError("В блоке «Навыки» не найдена кнопка добавления")
    for index in range(count):
        button = buttons.nth(index)
        metadata = " ".join(
            filter(None, [button.get_attribute("aria-label"), button.get_attribute("title")])
        )
        if re.search(r"добав|создат|add|plus", metadata, re.I):
            return button
    # На текущей форме «+» — первая из двух кнопок, вторая сворачивает блок.
    return buttons.first


def upload_skill(page, files_root):
    add_button = skills_add_button(page)
    try:
        with page.expect_file_chooser(timeout=8000) as chooser_info:
            add_button.click()
        chooser = chooser_info.value
        element = chooser.element
        is_directory = element.get_attribute("webkitdirectory") is not None
        if is_directory:
            chooser.set_files(str(files_root / "skill"))
        elif chooser.is_multiple():
            chooser.set_files(
                [str(files_root / relative) for relative in AGENT_FILES[:3]]
            )
        else:
            raise RuntimeError(
                "Форма принимает только один файл. Безопасная загрузка навыка "
                "без архива невозможна: нужны одновременно SKILL.md, scripts и references."
            )
    except PlaywrightTimeoutError:
        inputs = page.locator("input[type=file]")
        if not inputs.count():
            raise RuntimeError("Кнопка «+» не открыла файловое поле")
        file_input = inputs.first
        is_directory = file_input.get_attribute("webkitdirectory") is not None
        is_multiple = file_input.get_attribute("multiple") is not None
        if is_directory:
            file_input.set_input_files(str(files_root / "skill"))
        elif is_multiple:
            file_input.set_input_files(
                [str(files_root / relative) for relative in AGENT_FILES[:3]]
            )
        else:
            raise RuntimeError("Файловое поле не поддерживает загрузку папки навыка")
    page.wait_for_timeout(2500)
    errors = page.get_by_text(re.compile(r"ошиб|не удалось|неверн", re.I))
    if errors.count() and errors.first.is_visible():
        raise RuntimeError("GigaCowork сообщил об ошибке загрузки: " + errors.first.inner_text())


def save_agent(page):
    save = page.get_by_role("button", name=re.compile(r"^Сохранить$", re.I)).first
    save.wait_for(state="visible", timeout=10000)
    deadline = 30
    while save.is_disabled() and deadline:
        page.wait_for_timeout(1000)
        deadline -= 1
    if save.is_disabled():
        raise RuntimeError(
            "Кнопка «Сохранить» осталась неактивной. Проверьте обязательные поля "
            "и завершение загрузки навыка."
        )
    save.click()
    try:
        page.wait_for_url(lambda url: "/agents/new" not in url, timeout=30000)
    except PlaywrightTimeoutError:
        validation = page.get_by_text(re.compile(r"обязатель|ошиб|не удалось", re.I))
        if validation.count() and validation.first.is_visible():
            raise RuntimeError("Агент не сохранён: " + validation.first.inner_text())
        print("Сохранение отправлено; страница не изменила адрес за 30 секунд.")


def launch_context(playwright, profile):
    options = {
        "user_data_dir": str(profile),
        "headless": False,
        "no_viewport": True,
        "accept_downloads": True,
        "args": ["--start-maximized"],
    }
    try:
        return playwright.chromium.launch_persistent_context(channel="chrome", **options)
    except PlaywrightError as error:
        print("Google Chrome не найден штатно; пробую Chromium.")
        system_chromium = shutil.which("chromium") or shutil.which("chromium-browser")
        if system_chromium:
            options["executable_path"] = system_chromium
        try:
            return playwright.chromium.launch_persistent_context(**options)
        except PlaywrightError:
            raise RuntimeError(
                "Не удалось запустить браузер. Установите Chrome или выполните: "
                "python -m playwright install chromium"
            ) from error


def parse_args():
    parser = argparse.ArgumentParser(
        description="Загрузить агента из GitHub в GigaCowork без ZIP-архива"
    )
    parser.add_argument("--workspace-url", default=DEFAULT_URL)
    parser.add_argument("--ref", default="main", help="Ветка или коммит GitHub")
    parser.add_argument("--agent-name", default=AGENT_NAME)
    parser.add_argument("--agent-description", default=AGENT_DESCRIPTION)
    parser.add_argument(
        "--profile",
        default=str(Path.home() / ".gigacowork-deploy" / "chrome-profile"),
        help="Локальный профиль браузера; не хранится в репозитории",
    )
    parser.add_argument(
        "--dry-run", action="store_true", help="Только скачать и проверить файлы"
    )
    parser.add_argument(
        "--login-only", action="store_true", help="Только сохранить сессию браузера"
    )
    parser.add_argument(
        "--keep-open", action="store_true", help="Не закрывать браузер после завершения"
    )
    return parser.parse_args()


def main():
    args = parse_args()
    profile = Path(args.profile).expanduser().resolve()
    profile.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="gigacowork-agent-") as temporary:
        root = Path(temporary)
        download_files(root, args.ref)
        validate_files(root)
        if args.dry_run:
            print("DRY_RUN=OK: файлы GitHub скачаны и проверены, браузер не открывался.")
            return 0
        with sync_playwright() as playwright:
            context = launch_context(playwright, profile)
            page = context.pages[0] if context.pages else context.new_page()
            try:
                wait_for_form(page, args.workspace_url)
                if args.login_only:
                    print("LOGIN=OK: сессия сохранена в " + str(profile))
                    return 0
                print("Заполняю название, описание и инструкцию…")
                fill_field(page, "Название", args.agent_name, "Например")
                fill_field(page, "Описание", args.agent_description, "Кратко опишите")
                fill_field(
                    page,
                    "Инструкция",
                    (root / "system_prompt.txt").read_text(encoding="utf-8"),
                    "Опишите, как агент",
                )
                print("Загружаю SKILL.md, scripts и references…")
                upload_skill(page, root)
                print("Сохраняю агента…")
                save_agent(page)
                print("DEPLOY=OK: агент сохранён в GigaCowork.")
                print("Проверьте его на трёх Excel-файлах перед использованием.")
                if args.keep_open and sys.stdin.isatty():
                    input("Нажмите Enter, чтобы закрыть браузер…")
            except Exception as error:
                screenshot = Path.cwd() / "gigacowork-deploy-error.png"
                try:
                    page.screenshot(path=str(screenshot), full_page=True)
                    print("Снимок ошибки: " + str(screenshot), file=sys.stderr)
                except Exception:
                    pass
                print("DEPLOY=FAILED: " + str(error), file=sys.stderr)
                print("Учётные данные в сообщения скрипта не записываются.", file=sys.stderr)
                return 1
            finally:
                if not args.keep_open:
                    context.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
