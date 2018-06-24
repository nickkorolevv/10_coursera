from lxml import etree
import requests
from bs4 import BeautifulSoup
import json
from openpyxl import Workbook
import random
import argparse


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=str, default="example.xls")
    return parser


def get_html_code(url):
    html_doc = requests.get(url).content
    return html_doc


def get_courses_urls(xml_from_coursera):
    tree = etree.fromstring(xml_from_coursera)
    urls = [url.text for url in tree.iter("{*}loc")]
    return urls


def get_course_info(html_doc):
    soup = BeautifulSoup(html_doc, "lxml")
    course_mark = soup.find("div", {"class": "ratings-text bt3-visible-xs"})
    if course_mark is not None:
        course_mark = course_mark.text
    course_name = soup.find("title").text
    course_lang = soup.find("div", {"class": "rc-Language"})
    if course_lang is not None:
        course_lang = soup.find("div", {"class": "rc-Language"}).text
    course_duration = len(soup.find_all("div", {"class": "week"}))
    json_course = soup.find("script", {"type": "application/ld+json"})
    if json_course is not None:
        json_course = json_course.text
    try:
        near_course = json.loads(json_course)["hasCourseInstance"]["startDate"]
    except KeyError:
        near_course = None
    course_info_dict = {
        "Название курса": course_name,
        "Язык курса": course_lang,
        "Длительность курса": course_duration,
        "Оценка": course_mark,
        "Ближайший курс": near_course
    }
    return course_info_dict


def output_courses_info_to_xlsx(courses_info):
    courses_workbook = Workbook()
    active_sheet = courses_workbook.active
    head_line = [
        "Название курса",
        "Язык курса",
        "Длительность курса",
        "Оценка",
        "Ближайший курс"
    ]
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
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    parser = create_parser()
    parser_args = parser.parse_args()
    html_doc = get_html_code(url)
    all_urls_list = get_courses_urls(html_doc)
    number_of_courses = 3
    urls_list = random.sample(
        all_urls_list,
        number_of_courses
    )
    courses_info_list = []
    for url in urls_list:
        html_doc = get_html_code(url)
        course_info = get_course_info(html_doc)
        courses_info_list.append(course_info)
    courses_workbook = output_courses_info_to_xlsx(courses_info_list)
    output_filepath = parser_args.output
    courses_workbook.save(output_filepath)
