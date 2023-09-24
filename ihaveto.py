import asyncio
import aiohttp
import openpyxl
import json
from bs4 import BeautifulSoup
from arsenic import get_session, browsers, services

file_path = r"C:\Users\koush\Downloads\persons.xlsx"
workbook = openpyxl.load_workbook(file_path)
worksheet = workbook.active


def get_value(key, data):
    if "playerInfo" in data and key in data["playerInfo"]:
        return data["playerInfo"][key]
    else:
        return f"Not Available"


def to_add(data, i):
    try:
        string_list = []
        nickname = get_value("nickname", data)
        level = get_value("level", data)
        signature = get_value("signature", data)
        name_card_id = get_value("nameCardId", data)
        finish_achievement_num = get_value("finishAchievementNum", data)
        string_list.extend(
            [i, nickname, level, signature, name_card_id, finish_achievement_num]
        )
        print(f"UID: {i}")
        print(f"Nickname: {nickname}")
        print(f"Level: {level}")
        print(f"Signature: {signature}")
        print(f"Name Card ID: {name_card_id}")
        print(f"Finish Achievement Num: {finish_achievement_num}")
        worksheet.append(string_list)
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


y = "http://62.109.31.192:20000"
#                    f"--proxy-server={y}",


async def astre(url, limit, i):
    try:
        service = services.Chromedriver(
            binary=r"C:\Users\koush\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"
        )
        browser = browsers.Chrome()
        browser.capabilities = {
            "goog:chromeOptions": {
                "args": [
                    "--headless",
                    "--disable-gpu",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ]
            }
        }
        async with limit:
            async with get_session(service, browser) as session:
                await session.get(url)
                await asyncio.sleep(2)
                html = await session.get_page_source()
                soup = BeautifulSoup(html, "html.parser")
                title = soup.body.string.strip()
                data = json.loads(title)
                to_add(data, i)
    except aiohttp.ClientError as e:
        print(f"HTTP request error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


x = 5


async def main():
    try:
        limit = asyncio.Semaphore(x)
        tasks = [
            astre(f"https://enka.network/api/uid/{i}?info", limit, i)
            for i in range(618285856, 618285858)
        ]
        await asyncio.gather(*tasks)
    except Exception as e:
        print(f"An error occurred in the main function: {e}")
    workbook.save(file_path)
    workbook.close()


if __name__ == "__main__":
    asyncio.run(main())
