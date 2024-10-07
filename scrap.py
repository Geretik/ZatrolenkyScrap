from bs4 import BeautifulSoup
import pandas as pd
import requests
import os


SCRAP_URL = 'https://www.zatrolene-hry.cz/klub/klub-deskovych-her-doupe-olomouc-58/'
TMP_FILE_NAME = 'tmp_webpage.html'


def web_page_download(url: str, file_name: str) -> str | None:
    try:
        response = requests.get(url)
        with open(file_name, 'w', encoding='utf-8') as file:
            file.write(response.text)
            print("HTML downloaded")
        return file_name
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


def html_scrap_games(html_name: str) -> pd.ExcelFile:
    try:
        with open(html_name, 'r', encoding='utf-8') as file:
            html_content = file.read()

    except FileNotFoundError:
        print("The file was not found.")
        return
    except PermissionError:
        print("You do not have permission to access this file.")
        return
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return

    # Parse the HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Extracting data
    games = []
    for game in soup.find_all('div', class_='row list-item mb-3 pb-3'):
        # Extract the game name and link    
        game_info = game.find('h3')
        game_name = game_info.a.text if game_info.a else ''
        game_link = game_info.a['href'] if game_info.a else ''

        # Extract the commentary
        commentary_div = game.find('div', class_='card-body')
        commentary = commentary_div.text.strip() if commentary_div else ''

        # Append to list
        games.append([game_name, game_link, commentary])

    print("Scrap done")
    # Create DataFrame (no need for Link column)
    df = pd.DataFrame(games, columns=['Název hry', 'Link', 'Komentář'])

    # Save to Excel with clickable hyperlinks for Game Name
    with pd.ExcelWriter('Seznam_her.xlsx', engine='xlsxwriter') as writer:
        df[['Název hry', 'Komentář']].to_excel(writer, index=False, sheet_name='hry')

        # Access the XlsxWriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['hry']
        print("XLSX created")

        # Iterate through DataFrame and add hyperlinks
        for row_num, (name, link) in enumerate(zip(df['Název hry'], df['Link']), start=1):
            worksheet.write_url(f'A{row_num + 1}', link, string=name)  # A is for the 'Game Name' column
    print("URLS mapped")
    


def clean(file_path: str) -> bool:
    try:
        os.remove(file_path)
        print(f"{file_path} has been removed successfully.")
        return True
    except FileNotFoundError:
        print("File not found!")
        return False
    except PermissionError:
        print("Permission denied!")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False
        

if __name__ == "__main__":
    name = web_page_download(SCRAP_URL, TMP_FILE_NAME)
    html_scrap_games(name)
    clean(name)
