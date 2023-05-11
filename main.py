import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from parsing import desired_games


def main():
    logging.basicConfig(level=logging.CRITICAL)

    chrome_options = Options()
    chrome_options.add_argument("--headless")

    with webdriver.Chrome(options=chrome_options) as driver:
        desired_games(driver)
        
if __name__ == '__main__':
    main()