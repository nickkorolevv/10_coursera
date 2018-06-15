from lxml import etree
import requests
from bs4 import BeautifulSoup
import json
from openpyxl import Workbook
import random


def get_courses_list(page_xml):
    response_content = requests.get(page_xml).content
    tree = etree.fromstring(response_content)
    urls = [url.text for url in tree.iter("{*}loc")]
    return urls


def get_course_info(url):
    html_doc = requests.get(url)
    soup = BeautifulSoup(html_doc.text, "lxml")
    course_mark = soup.find("div", {"class": "ratings-text bt3-visible-xs"})
    if course_mark is not None:
        course_mark = course_mark.text
    course_name = soup.find("title").text
    course_language = soup.find("div", {"class": "rc-Language"})
    if course_language is not None:
        course_language = soup.find("div", {"class": "rc-Language"}).text
    course_duration = len(soup.find_all("div", {"class": "week"}))
    json_course = soup.find("script", {"type": "application/ld+json"}).text
    try:
        near_course = json.loads(json_course)["hasCourseInstance"]["startDate"]
    except KeyError:
        near_course = "Нет доступных предстоящих сессий"
    course_info_dict = {}
    course_info_dict.update(
        {"Название курса": course_name,
         "Язык курса": course_language,
         "Длительность курса": course_duration,
         "Оценка": course_mark,
         "Ближайший курс": near_course}
    )
    return course_info_dict


def output_courses_info_to_xlsx(courses_info):
    courses_workbook = Workbook()
    active_sheet = courses_workbook.active
    head_line = ["Название курса",
                 "Язык курса",
                 "Длительность курса",
                 "Оценка",
                 "Ближайший курс"]
    active_sheet.append(head_line)

    for course in courses_info:
        active_sheet.append([
            course["Название курса"],
            course["Язык курса"],
            course["Длительность курса"],
            course["Оценка"],
            course["Ближайший курс"]
        ])
    return courses_workbook


if __name__ == "__main__":
    page_xml = "https://www.coursera.org/sitemap~www~courses.xml"
    all_urls_list = get_courses_list(page_xml)
    number_of_courses = 20
    urls_list = random.sample(
        all_urls_list,
        number_of_courses
    )
    courses_info_list = []
    for url in urls_list:
        course_info = get_course_info(url)
        courses_info_list.append(course_info)
    courses_workbook = output_courses_info_to_xlsx(courses_info_list)
    courses_workbook.save("example.xls")
