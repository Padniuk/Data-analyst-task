from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from structures import Game, game_genres, platforms
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException
import time
import os

def style_chriteria(genre):
    styles = [game_genres[style] for style in genre]
    realism_percentage = (len(list(filter(lambda x: x == 'Realism', styles))) / len(styles))
    if realism_percentage > 0.8:
        return 'Realism'
    elif realism_percentage > 0.2:
        return 'Semi-Realism'
    else:
        return 'Stylization'


def desired_games(driver, max_value=1000):
    url = 'https://rawg.io/games/strategy'
    driver.get(url)

    for i in range(0,4):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

    games = []
    div_tags = driver.find_elements(By.CSS_SELECTOR, "div[class='game-card-medium__info']")

    for div_tag in div_tags:
        a_tag = div_tag.find_element(By.TAG_NAME, 'a')
        span_tag = div_tag.find_element(By.CSS_SELECTOR, "span[class='game-card-button__inner']")
        downloads = span_tag.text
        href = a_tag.get_attribute('href')
        games.append((href,downloads))

    if os.path.exists("games.xlsx"):
        os.remove("games.xlsx")
        
    for game in games[:150]:
        game_processor(game)


def game_processor(data):
 
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)

    game = Game()
    game.downloads = data[1]
    driver.get(data[0])
    time.sleep(2)
    h1_tag = driver.find_element(By.CSS_SELECTOR, "h1[class='heading heading_1 game__title']")
    game.name = h1_tag.get_attribute("textContent")
    div_tags = driver.find_elements(By.CSS_SELECTOR,"div[class='game__meta-block']")

    image_tag = driver.find_element(By.CSS_SELECTOR, "img[class='responsive-image game__screenshot-image']")
    game.image = image_tag.get_attribute("src")

    for div in div_tags:
        text = div.text.split('\n')
        name = text[0]
        value = text[1]
        match name:
            case 'Platforms':
                platforms_list = value.split(', ')
                if len(platforms_list) == 1:
                    game.platform = platforms_list[0]
                else:
                    common_values = set(platforms[platform] for platform in platforms_list)
                    if len(common_values) == 1:
                        game.platform = common_values.pop()
                    else:
                        game.platform = 'Switch'
    
            case 'Metascore':
                game.rating = float(value[0])/10.
            case 'Genre':
                game.genre = value
                genres_list = game.genre.split(', ')
                game.style = style_chriteria(genres_list)
            case 'Release date':
                game.release = value
            case 'Developer':
                game.developer = value
            case 'Publisher':
                game.publisher = value

    game.save()