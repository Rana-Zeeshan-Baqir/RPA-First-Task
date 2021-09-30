import time
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Desktop import Desktop


class Title:
    def __init__(self, keyword, base_link):
        self.desktop = Desktop()
        self.browser = Selenium()
        self.lib = Files()
        self.keyword = keyword
        self.url_link = base_link
        self.open_browser = self.browser.open_available_browser(base_link)
        self.browser.input_text('//input[@title="Search"]', keyword)
        time.sleep(1)
        self.desktop.press_keys("enter")
        self.web_link = []

    def loop_browser(self):
        while True:
            for ele in self.browser.find_elements('//div[@class="yuRUbf"]'):
                link = ele.find_element_by_tag_name("a").get_attribute("href")
                title = ele.find_element_by_tag_name("h3").text
                if title and link:
                    self.web_link.append({"title": title, "link": link})
            try:
                self.browser.wait_until_page_contains_element('//*[@id="pnnext"]/span[2]')
                self.browser.find_element('//*[@id="pnnext"]/span[2]').click()
            except:
                break

    def workbook(self):
        file = self.lib.create_workbook("workbook.xlsx")
        file.append_worksheet("Sheet", content=self.web_link, header=True)
        file.save()


user_input = input("Enter what you want to search:")
title_obj = Title(user_input, "https://www.google.com")
title_obj.loop_browser()
print(title_obj.web_link)
title_obj.workbook()
