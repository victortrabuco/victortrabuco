from RPA.Robocorp.Process   import Process
from RPA.Robocorp.WorkItems import WorkItems
from RPA.Browser.Selenium   import Selenium
from RPA.Excel.Application  import Application
from time                   import sleep
from os.path                import abspath
import urllib
import datetime
import re


class excel:
    def __init__(self):
        self.excel_app = Application()
        self.excel_app.open_application(visible=True)
        self.file_name = abspath(f"data/Result {datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')}.xlsx")
        self.header = {}

    def set_new_workbook(self, header):
        self.excel_app.add_new_workbook()
        self.excel_app.save_excel_as(self.file_name)
        self.excel_app.close_document()
        self.excel_app.open_workbook(self.file_name)
        for count, name in enumerate(header):
            self.header[name] = count + 1
            self.excel_app.write_to_cells(row=1, column=count + 1, value=name)

    def save_to_workbook(self, data):
        row, column = self.excel_app.find_first_available_row()
        for column_name in data:
            self.excel_app.write_to_cells(row=row, column=self.header[column_name], value=data[column_name])
        self.excel_app.save_excel()


class web:
    def __init__(self):
        self.browser = Selenium()
        self.browser.open_available_browser()
        self.excel = excel()
        self.excel.set_new_workbook([   "Date",
                                        "Title",
                                        "Description",
                                        "Image file name",
                                        "Phrase count",
                                        "Currency on title or description"])

    def navigate(self):
        self.browser.maximize_browser_window()
        self.browser.go_to(url="https://www.latimes.com/")

    def get_element(self, locator, parent=None, frame_index=-1):
        try:
            if parent:
                self.browser.wait_until_page_contains_element(locator=parent)
            else:
                self.browser.wait_until_page_contains_element(locator=locator)
            return self.browser.find_elements(locator=locator, parent=parent)
        except:
            frame_index += 1
            self.browser.select_frame(self.browser.find_elements("//iframe")[frame_index])
            self.get_element(locator=locator, parent=parent, frame_index=frame_index)

    def search(self, search_phrase):
        self.browser.wait_until_element_is_visible("//button[@data-element='search-button']")
        self.browser.scroll_element_into_view("//button[@data-element='search-button']")
        self.browser.click_element("//button[@data-element='search-button']")
        self.browser.wait_until_element_is_visible("//input[@data-element='search-form-input']")
        self.browser.press_keys("//input[@data-element='search-form-input']", search_phrase)
        self.browser.click_element("//button[@data-element='search-submit-button']")
        try:
            self.browser.wait_until_element_is_visible("//ul[@class='search-results-module-results-menu']")
        except:
            return False
        if not len(self.browser.find_elements("//ul[@class='search-results-module-results-menu']//li")):
            return False
        return True

    def select_category(self, categories_or_sections):
        if not len(categories_or_sections):
            return
        # Loop on 'Topics'
        load_all = True
        for name in categories_or_sections:
            if load_all:
                self.browser.wait_until_element_is_visible("//div[@class='search-filter'][1]//button")
                sleep(1)
                see_all_button = self.browser.find_elements("//div[@class='search-filter'][1]//button")
                self.browser.click_element(see_all_button[0])
            topics_lines = self.browser.find_elements("//div[@class='search-filter'][1]//li")
            for index, topic_line in enumerate(topics_lines, 1):
                if name.lower() in self.browser.get_text(topic_line).lower():
                    self.browser.scroll_element_into_view(f"//div[@class='search-filter'][1]//li[{index}]//input")
                    self.browser.click_element(f"//div[@class='search-filter'][1]//li[{index}]//input")
                    sleep(2)
                    load_all = True
                    break
                else:
                    load_all = False
        # Loop on 'Type'
        load_all = True
        for name in categories_or_sections:
            if load_all:
                self.browser.wait_until_element_is_visible("//div[@class='search-filter'][2]//button")
                sleep(1)
                see_all_button = self.browser.find_elements("//div[@class='search-filter'][2]//button")
                self.browser.click_element(see_all_button[0])
            topics_lines = self.browser.find_elements("//div[@class='search-filter'][2]//li")
            for index, topic_line in enumerate(topics_lines, 1):
                if name.lower() in self.browser.get_text(topic_line).lower():
                    self.browser.scroll_element_into_view(f"//div[@class='search-filter'][2]//li[{index}]//input")
                    self.browser.click_element(f"//div[@class='search-filter'][2]//li[{index}]//input")
                    sleep(2)
                    load_all = True
                    break
                else:
                    load_all = False

    def select_newest(self):
        select_list = self.browser.find_elements("//div[@class='search-results-module-sorts']//select")[0]
        self.browser.select_from_list_by_label(select_list, "Newest")
        sleep(2)

    def get_news(self, month_range, search_phrase):
        if month_range <= 1:
            month_range = 0
        else:
            month_range -= 1
        datetime_now = datetime.datetime.now()
        count = 1
        finished = False
        while not finished:
            lines = self.browser.find_elements("//ul[@class='search-results-module-results-menu']//li")
            for _ in range(count, len(lines)):
                try:
                    pub_time = self.browser.get_element_attribute(f"//ul[@class='search-results-module-results-menu']//li[{count}]//p[@class='promo-timestamp']", "data-timestamp")
                except:
                    count+=1
                    continue
                pub_date = datetime.datetime.fromtimestamp(int(pub_time) / 1000, datetime.UTC)
                months_diff = (datetime_now.year - pub_date.year) * 12 + datetime_now.month - pub_date.month
                if months_diff > month_range:
                    finished = True
                    break
                pub_title = self.browser.get_text(f"//ul[@class='search-results-module-results-menu']//li[{count}]//h3[@class='promo-title']")
                pub_description = self.browser.get_text(f"//ul[@class='search-results-module-results-menu']//li[{count}]//p[@class='promo-description']")
                try:
                    pub_img_src = self.browser.get_element_attribute(f"//ul[@class='search-results-module-results-menu']//li[{count}]//img", "src")
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
                self.excel.save_to_workbook({   "Date": pub_date.strftime("%Y-%m-%d"),
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


class main:
    def execute(self):
        work_items = WorkItems()
        web_obj = web()
        while True:
            try:
                work_items.get_input_work_item()
            except:
                break
            try:
                web_obj.navigate()
                search_phrase = work_items.get_work_item_variable("search_phrase")
                if not web_obj.search(search_phrase):
                    continue
                news_category = work_items.get_work_item_variable("news_category")
                web_obj.select_category(news_category)
                sleep(1.5)
                web_obj.select_newest()
                sleep(1.5)
                number_months = work_items.get_work_item_variable("number_months")
                web_obj.get_news(number_months, search_phrase)
            except:
                continue

if __name__ == "__main__":
    Process().start_process()
    main().execute()
