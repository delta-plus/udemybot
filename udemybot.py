#!/usr/bin/python

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import itertools
import xlwt
import sys  

reload(sys)  
sys.setdefaultencoding('utf8')
driver = webdriver.PhantomJS()
driver.set_window_size(1120, 550)
driver.get(sys.argv[1])
book = xlwt.Workbook(encoding='latin-1')
sheetCount = itertools.count()
check = 1
while check:
    try:
        sheet = book.add_sheet(sys.argv[2].partition('.xls')[0] + 'Page ' + str(next(sheetCount)), cell_overwrite_ok=True)
        sheet.write(1, 0, 'URL')
        sheet.write(2, 0, 'Instructors')
        sheet.write(3, 0, 'Stars')
        sheet.write(4, 0, '# of Ratings')
        sheet.write(5, 0, 'Students in Course')
        sheet.write(6, 0, 'Last Updated')
        sheet.write(7, 0, 'Price')
        sheet.write(8, 0, 'Sale')
        sheet.write(9, 0, 'Hours of Video')
        sheet.write(10, 0, 'Articles')
        sheet.write(11, 0, 'Other Resources')
        sheet.write(12, 0, 'Lectures')
        sheet.write(13, 0, 'Instructor Ratings')
        sheet.write(14, 0, '# of Instructor Reviews')
        sheet.write(15, 0, 'Students Taught')
        sheet.write(16, 0, 'Courses Taught')
        print 'Searching', driver.current_url
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'search-course-card--card__title--2xzHX'))
        )
        count = driver.find_elements_by_class_name('search-course-card--card__title--2xzHX')
        for x in range(0, len(count)):
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'search-course-card--card__title--2xzHX'))
            )
            articles = driver.find_elements_by_class_name('search-course-card--card__title--2xzHX')
            sheet.write(0, x + 1, articles[x].text)
            articles[x].click()
            soup = BeautifulSoup(driver.page_source, 'lxml')
            text = soup.get_text().encode('utf-8').replace(' ', '')
            courseRatings = text[text.find('ratings)') - 10:text.find('ratings)')]
            numRatings = courseRatings.partition('(')[2]
            stars = courseRatings.partition('(')[0].strip()
            instructors = driver.find_element_by_xpath('.//*[@id=\'udemy\']/div[*]/div/div[2]/div[1]/div/div/div[3]/div[1]/div/span').text
            studentsInCourse = text[text.find('studentsenrolled') - 10:text.find('studentsenrolled')].strip()
            lastUpdated = text[text.find('Lastupdated') + 11:text.find('Lastupdated') + 19].strip()
            oldPrice = driver.find_element_by_class_name('price-text').text
            hoursOfVideo = driver.find_element_by_xpath('.//*[@id=\'udemy\']/div[*]/div/div[2]/div[2]/div/div[2]/div[2]/div[1]/ul/li[1]/span').text.partition(' ')[0]
            lectures = driver.find_element_by_class_name('collapsed-text').text
            stats = driver.find_elements_by_class_name('instructor__stat')
            counter = 0
            rating = ''
            reviews = ''
            students = ''
            courses = ''
            for stat in stats:
              if counter % 4 == 0:
                if counter is not 0:
                  rating += ',' + ' ' + stat.text
                else:
                  rating = stat.text
              if counter % 4 == 1:
                if counter is not 1:
                  reviews += ',' + ' ' + stat.text
                else:
                  reviews = stat.text   
              if counter % 4 == 2:
                 if counter is not 2:
                   students += ',' + ' ' + stat.text
                 else:
                   students = stat.text
              if counter % 4 == 3:
                if counter is not 3:
                  courses += ',' + ' ' + stat.text
                else:
                  courses = stat.text
              counter += 1
            sheet.write(1, x + 1, driver.current_url)
            sheet.write(2, x + 1, instructors[11:len(instructors)])
            sheet.write(3, x + 1, stars)
            sheet.write(4, x + 1, numRatings)
            sheet.write(5, x + 1, studentsInCourse)
            sheet.write(6, x + 1, lastUpdated)
            sheet.write(7, x + 1, '$10')
            sheet.write(8, x + 1, oldPrice[0:len(oldPrice)])
            sheet.write(9, x + 1, hoursOfVideo)
            includes = driver.find_elements_by_class_name('incentives__text')
            for item in includes:
              if item.text.find('Article') is not -1:
                sheet.write(10, x + 1, item.text[0:item.text.find('Article')])
              if item.text.find('Supplemental') is not -1:
                sheet.write(11, x + 1, item.text[0:item.text.find('Supplemental')])
            sheet.write(12, x + 1, lectures[11:len(lectures)].partition(' ')[0])
            sheet.write(13, x + 1, rating)
            sheet.write(14, x + 1, reviews)
            sheet.write(15, x + 1, students)
            sheet.write(16, x + 1, courses)
            driver.back()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, './/*[@id=\'udemy\']/ui-view/ui-view/div[1]/div[2]/div/div/ul[3]/li[7]/a'))
        )
        nextPage = driver.find_element_by_xpath('.//*[@id=\'udemy\']/ui-view/ui-view/div[1]/div[2]/div/div/ul[3]/li[7]/a')
        endCheck = driver.current_url
        nextPage.click()
        if endCheck == driver.current_url:
          check = 0
    except NoSuchElementException:
        check = 0
book.save(sys.argv[2])
