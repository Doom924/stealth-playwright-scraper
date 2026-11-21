import asyncio
import ipaddress
import socket
from urllib.parse import urlparse
from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font


def is_safe_url(url: str) -> bool:
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        return False
    try:
        ip = ipaddress.ip_address(socket.gethostbyname(parsed.hostname))
        return not (ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved)
    except socket.gaierror:
        return False


async def scrape_dynamic_quotes(target_url: str, max_pages: int = 3) -> list[dict]:
    if not is_safe_url(target_url):
        raise ValueError(f"Небезопасный или недопустимый URL: {target_url}")

    results = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 768},
        )
        page = await context.new_page()

        # Блокируем загрузку тяжёлых картинок и шрифтов для ускорения парсинга
        await page.route(
            "**/*",
            lambda route: route.abort()
            if route.request.resource_type in ("image", "media", "font")
            else route.continue_(),
        )

        await page.goto(target_url, wait_until="networkidle", timeout=15000)

        for _ in range(max_pages):
            await page.wait_for_selector(".quote", timeout=10000)
            cards = await page.query_selector_all(".quote")

            for card in cards:
                text_el = await card.query_selector(".text")
                author_el = await card.query_selector(".author")
                tag_els = await card.query_selector_all(".tags .tag")

                text = (await text_el.inner_text()).strip("“”") if text_el else ""
                author = (await author_el.inner_text()).strip() if author_el else ""
                tags = ", ".join([(await t.inner_text()).strip() for t in tag_els])

                results.append({"author": author, "text": text, "tags": tags})

            next_btn = await page.query_selector("li.next a")
            if not next_btn:
                break
            await next_btn.click()
            await page.wait_for_load_state("networkidle")

        await browser.close()
    return results


def save_to_excel(items: list[dict], filename: str = "dynamic_export.xlsx") -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "JS Scrape Data"

    headers = ["Автор", "Цитата", "Теги"]
    ws.append(headers)
    for cell in ws:
        cell.font = Font(bold=True)

    for item in items:
        ws.append([item["author"], item["text"], item["tags"]])

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col.column_letter].width = min(max_len + 4, 70)

    wb.save(filename)


async def main() -> None:
    url = "https://toscrape.com"
    print(f"[*] Запуск headless-браузера для: {url}")
    data = await scrape_dynamic_quotes(url, max_pages=3)
    save_to_excel(data)
    print(f"[+] Готово! Собрано записей: {len(data)}. Сохранено в dynamic_export.xlsx")


if __name__ == "__main__":
    asyncio.run(main())
