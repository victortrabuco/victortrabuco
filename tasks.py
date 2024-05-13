from robocorp.tasks import task
from RPA.Robocorp.WorkItems import WorkItems
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from time import sleep
from os.path import abspath
import urllib
import datetime
import re
import logging


class excel:
    def __init__(self):
        self.excel_app = Files()
        datetime_now = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
        file_name = abspath(f"data/Result {datetime_now}.xlsx")
        self.excel_app.create_workbook(file_name)

    def set_header(self, header):
        self.header = {}
        for count, name in enumerate(header):
            self.header[name] = count + 1
            self.excel_app.set_cell_value(row=1, column=count + 1, value=name)

    def save_to_workbook(self, data):
        row = self.excel_app.find_empty_row()
        for column_name in data:
            self.excel_app.set_cell_value(row=row,
                                          column=self.header[column_name],
                                          value=data[column_name])
        self.excel_app.save_excel()

    def exit_excel(self):
        self.excel_app.close_document(save_changes=True)


class web:
    def __init__(self):
        self.browser = Selenium()
        self.browser.open_available_browser()

    def navigate(self):
        self.browser.maximize_browser_window()
        self.browser.go_to(url="https://www.latimes.com/")

    def exit_browser(self):
        self.browser.close_browser()

    def search(self, search_phrase):
        element_locator = "//button[@data-element='search-button']"
        self.browser.wait_until_element_is_visible(element_locator)
        self.browser.scroll_element_into_view(element_locator)
        self.browser.click_element(element_locator)
        element_locator = "//input[@data-element='search-form-input']"
        self.browser.wait_until_element_is_visible(element_locator)
        self.browser.press_keys(element_locator, search_phrase)
        element_locator = "//button[@data-element='search-submit-button']"
        self.browser.click_element(element_locator)
        try:
            element_locator = "//ul[@class='search-results-module-results-menu']"
            self.browser.wait_until_element_is_visible(element_locator)
        except:
            return False
        element_locator = "//ul[@class='search-results-module-results-menu']//li"
        if not len(self.browser.find_elements(element_locator)):
            return False
        return True

    def select_category(self, categories_or_sections):
        if not len(categories_or_sections):
            return
        # Loop on 'Topics'
        load_all = True
        for name in categories_or_sections:
            if load_all:
                element_locator = "//div[@class='search-filter'][1]//button"
                self.browser.wait_until_element_is_visible(element_locator)
                sleep(1)
                see_all_button = self.browser.find_elements("//div[@class='search-filter'][1]//button")
                self.browser.click_element(see_all_button[0])
            first_filter = "//div[@class='search-filter'][1]//li"
            topics_lines = self.browser.find_elements()
            for index, topic_line in enumerate(topics_lines, 1):
                if name.lower() in self.browser.get_text(topic_line).lower():
                    self.browser.scroll_element_into_view(f"{first_filter}[{index}]//input")
                    self.browser.click_element(f"{first_filter}[{index}]//input")
                    sleep(2)
                    load_all = True
                    break
                else:
                    load_all = False
        # Loop on 'Type'
        load_all = True
        for name in categories_or_sections:
            if load_all:
                element_locator = "//div[@class='search-filter'][2]//button"
                self.browser.wait_until_element_is_visible(element_locator)
                sleep(1)
                see_all_button = self.browser.find_elements(element_locator)
                self.browser.click_element(see_all_button[0])
            second_filter = "//div[@class='search-filter'][2]//li"
            topics_lines = self.browser.find_elements(second_filter)
            for index, topic_line in enumerate(topics_lines, 1):
                if name.lower() in self.browser.get_text(topic_line).lower():
                    self.browser.scroll_element_into_view(f"{second_filter}[{index}]//input")
                    self.browser.click_element(f"{second_filter}[{index}]//input")
                    sleep(2)
                    load_all = True
                    break
                else:
                    load_all = False

    def select_newest(self):
        element_locator = "//div[@class='search-results-module-sorts']//select"
        select_list = self.browser.find_elements(element_locator)[0]
        self.browser.select_from_list_by_label(select_list, "Newest")
        sleep(2)

    def get_news(self, month_range, search_phrase, excel: excel):
        if month_range <= 1:
            month_range = 0
        else:
            month_range -= 1
        datetime_now = datetime.datetime.now()
        count = 1
        finished = False
        while not finished:
            lines_locator = "//ul[@class='search-results-module-results-menu']//li"
            lines = self.browser.find_elements(lines_locator)
            for _ in range(count, len(lines)):
                try:
                    pub_time = self.browser.get_element_attribute(f"{lines_locator}[{count}]//p[@class='promo-timestamp']", "data-timestamp")
                except:
                    count += 1
                    continue
                pub_date = datetime.datetime.fromtimestamp(int(pub_time) / 1000, datetime.UTC)
                months_diff = (datetime_now.year - pub_date.year) * 12 + datetime_now.month - pub_date.month
                if months_diff > month_range:
                    finished = True
                    break
                pub_title = self.browser.get_text(f"{lines_locator}[{count}]//h3[@class='promo-title']")
                pub_description = self.browser.get_text(f"{lines_locator}[{count}]//p[@class='promo-description']")
                try:
                    pub_img_src = self.browser.get_element_attribute(f"{lines_locator}[{count}]//img", "src")
                    pub_img_path = f"data/news_pics/{datetime.datetime.now().strftime('%d%m%Y%H%M%S%f')}.png"
                    urllib.request.urlretrieve(pub_img_src, pub_img_path)
                except:
                    pub_img_path = ""
                count_phrase = str(len(re.findall(search_phrase.lower(), pub_title.lower())) + len(re.findall(search_phrase.lower(), pub_description.lower())))
                currency_pattern = r'(?:\$[0-9,]+(?:\.[0-9]+)?|\b\d+\sdollars\b|\b\d+\sUSD\b)'
                if re.search(currency_pattern, pub_title) or re.search(currency_pattern, pub_description):
                    currency = "True"
                else:
                    currency = "False"
                excel.save_to_workbook({   "Date": pub_date.strftime("%Y-%m-%d"),
                                                "Title": pub_title,
                                                "Description": pub_description,
                                                "Image file name": pub_img_path,
                                                "Phrase count": count_phrase,
                                                "Currency on title or description": currency})

                count += 1
            if not finished:
                show_more_button = "//div[@class='search-results-module-next-page']"
                try:
                    self.browser.scroll_element_into_view(show_more_button)
                    self.browser.click_element(f"{show_more_button}//a")
                except:
                    finished = True


@task
def capture_news():
    logger = logging.getLogger(__name__)
    logging.basicConfig(filename="applog.log", level=logging.INFO)
    logger.info("Staring WorkItems")
    work_items = WorkItems()
    logger.info("Staring Web")
    web_obj = web()
    logger.info("Staring Excel")
    excel_obj = excel()
    logger.info("Setting Excel header")
    excel_obj.set_header(["Date",
                          "Title",
                          "Description",
                          "Image file name",
                          "Phrase count",
                          "Currency on title or description"])
    while True:
        try:
            logger.info("Getting input work intem")
            work_items.get_input_work_item()
        except:
            logger.info("No input work item found")
            break
        try:
            logger.info("Input work item found")
            web_obj.navigate()
            search_phrase = work_items.get_work_item_variable("search_phrase")
            logger.info("Input work item search phrase: %s", search_phrase)
            if not web_obj.search(search_phrase):
                logger.info(f"No news found for '{search_phrase}'")
            news_category = work_items.get_work_item_variable("news_category")
            logger.info("Input work item news category: %s", "news_category")
            web_obj.select_category(news_category)
            sleep(1.5)
            web_obj.select_newest()
            sleep(1.5)
            number_months = work_items.get_work_item_variable("number_months")
            logger.info("Input work item news age in months: %s", str(number_months))
            web_obj.get_news(number_months, search_phrase, excel_obj)
        except Exception as error_message:
            logger.warning(error_message)
            continue
    web_obj.exit_browser()
    excel_obj.exit_excel()
