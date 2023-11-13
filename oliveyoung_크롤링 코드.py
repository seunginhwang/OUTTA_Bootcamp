#!/usr/bin/env python
# coding: utf-8

# In[14]:


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

import requests
from bs4 import BeautifulSoup
import time
import openpyxl


#엑셀파일 생성
wb = openpyxl.Workbook("index11.xlsx")        
ws = wb.create_sheet("시트명")             
ws.append(['브랜드','상품명','카테고리','정가','할인가','아이디','별점','피부정보','피부타입','피부고민','자극도'])  #컬럼명 제공

driver = webdriver.Chrome()
chrome_options = Options()
chrome_options.add_experimental_option('detach',True)
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

main_url = 'https://www.oliveyoung.co.kr/store/main/getBestList.do?dispCatNo=900000100100001&fltDispCatNo=10000010001&pageIdx=1&rowsPerPage=8&t_page=%EC%B9%B4%ED%85%8C%EA%B3%A0%EB%A6%AC%EA%B4%80&t_click=%EB%9E%AD%ED%82%B9BEST%EC%83%81%ED%92%88%EB%B8%8C%EB%9E%9C%EB%93%9C_%EC%9D%B8%EA%B8%B0%EC%83%81%ED%92%88%EB%8D%94%EB%B3%B4%EA%B8%B0'
response = requests.get(main_url)
html = response.text
soup = BeautifulSoup(html, 'html.parser')
links = soup.select('a.prd_thumb.goodsList') 


#for반복문을 활용하여 5위까지 제품별 상세링크 sub_url에 저장

sub_url= []

for link in links[:5]:
    href = link.attrs['href']
    sub_url.append(href)

time.sleep(1)


def customer_info(x):
    rev1 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(1)')
    rev2 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(2)')
    rev3 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(3)')
    rev4 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(4)')
    rev5 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(5)')
    rev6 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(6)')
    rev7 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(7)')
    rev8 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(8)')
    rev9 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(9)')
    rev10 = x.find_elements(By.CSS_SELECTOR, '#gdasList > li:nth-child(10)')
    rev_li = [rev1[0], rev2[0], rev3[0], rev4[0], rev5[0], rev6[0], rev7[0], rev8[0], rev9[0], rev10[0], ]
    
    for j in range(0,10,1):
        #브랜드
        try:
            brand = x.find_element(By.CSS_SELECTOR, '#moveBrandShop')
            brand = brand.text
        except:
            brand ="없음"

        #상품명
        try:
            p_name = x.find_element(By.CSS_SELECTOR, 'p.prd_name')
            p_name = p_name.text
        except:
            p_name="없음"

        #카테고리
        try:
            p_category = x.find_element(By.CSS_SELECTOR, '#dtlCatNm')
            p_category = p_category.text
        except:
            p_category="없음"

        #정가
        try:
            price = x.find_elements(By.CSS_SELECTOR, '#Contents > div.prd_detail_box.renew > div.right_area > div > div.price > span.price-1 > strike')
            price = price[0].text
        except:
            price=0

        #할인가
        try:
            discount = x.find_elements(By.CSS_SELECTOR, '#Contents > div.prd_detail_box.renew > div.right_area > div > div.price > span.price-2 > strong')
            discount = discount[0].text
        except:
            discount=0

            
            
        ###########################리뷰정보 수집###########################
        
        #id
        try:
            _id = rev_li[j].find_element(By.CSS_SELECTOR, 'a.id')
            _id = _id.text
        except:
            _id ="없음"
        #별점
        try:
            _star = rev_li[j].find_elements(By.CSS_SELECTOR, 'span.review_point > span')
            _star = _star[0].text
        except:
            _star="없음"
            
        #고객 피부 정보
        try:
            _info = rev_li[j].find_elements(By.CSS_SELECTOR, 'div.info > div > p.tag')
            _info = _info[0].text
        except:
            _info = "없음"


        #피부 타입
        try:
            skin_type = rev_li[j].find_elements(By.CSS_SELECTOR, 'dd > span')
            skin_type = skin_type[0].text
        except:
            skin_type = "없음"
        #피부 고민
        try:
            skin_trouble = rev_li[j].find_elements(By.CSS_SELECTOR, 'dd > span')
            skin_trouble = skin_trouble[1].text
        except:
            skin_trouble = "없음"
        #자극도
        try:
            skin_irritation = rev_li[j].find_elements(By.CSS_SELECTOR, 'dd > span')
            skin_irritation = skin_irritation[2].text
        except:
            skin_irritation = "없음"
            
        #엑셀 데이터 쌓기
        print(f"\n{_id} {_star} {_info}")
        ws.append([brand,p_name,p_category,price,discount,_id, _star, _info, skin_type, skin_trouble, skin_irritation])
        time.sleep(1)


import re

#웹페이지 해당 주소 이동
for i in range(0,5):          #전체 제품을 한번에 크롤링하지 않고 나눠서 크롤링 할 경우 sub_url 인덱스 고려해서 range숫자 변경
    print(f"\n현재 링크 index: {i}")
    driver.implicitly_wait(5)  #웹페이지 로딩 될때까지 5초는 기다림
    driver.maximize_window()   #화면 최대화
    driver.get(sub_url[i])          
    time.sleep(3)
    lp = driver.find_element(By.CSS_SELECTOR, '#reviewInfo > a > span')
    lp = re.sub(r'[^0-9]', '', lp.text)
    last_page = int(lp)//10 + 1
    
    #리뷰버튼 클릭
    rv = driver.find_element(By.CSS_SELECTOR, 'a.goods_reputation')
    rv.click()
    time.sleep(3)
    
    #리뷰 하단 끝까지 스크롤하는 함수 (빈칸없음.그대로 코드 사용가능)
    before_h = driver.execute_script("return window.scrollY")         #스크롤 전 높이
    #화면 맨아래까지 스크롤
    while True:
        driver.find_element(By.CSS_SELECTOR,"body").send_keys(Keys.END)
        time.sleep(1)
        #스크롤 후 높이
        after_h = driver.execute_script("return window.scrollY")

        #스크롤 값이 같으면 스크롤 멈춤
        if after_h == before_h:
            break
        before_h = after_h   
    
    for k in range(1,101):  #100페이지 크롤링 한다고 했을 때 (상품당 최대 100페이지까지 있음) k=현재 페이지
        print(f"\n{k}페이지")
        time.sleep(3)
        #마지막 페이지면 종료
        if k > last_page:
            break
        #페이지 숫자 10이하 일 때
        if k<11:
            try:
                if k != 10:
                    customer_info(driver)
                    next_p = "#gdasContentsArea > div > div.pageing > a:nth-child(" + str(k+1) + ")" #k+1 page 버튼 클릭
                    driver.find_element(By.CSS_SELECTOR, next_p).click()
                elif k == 10: #10번째 페이지 크롤링 한 후에 다음페이지로 가는 화살표(>>) 버튼 클릭
                    customer_info(driver)
                    driver.find_element(By.CSS_SELECTOR, "a.next").click()  
                    print(">>")
            except:
                pass
                   
                    
       #페이지 숫자 11이상 일 때  (규칙을 찾아 각 페이지 크롤링 후 다음 페이지로 이동하도록 코드 작성)        
        elif k>=11:
            try:
                if k%10 != 0: #11p~19p, 21p~29p, ...
                    customer_info(driver)
                    next_p = "#gdasContentsArea > div > div.pageing > a:nth-child(" + str(k%10+2) + ")" #k+1 page 버튼 클릭
                    driver.find_element(By.CSS_SELECTOR, next_p).click()
                
                elif k%10 == 0:  #20p, 30p, ... 일 땐 10x페이지 크롤링 한 후에 다음페이지로 가는 화살표(>>) 버튼 클릭
                    customer_info(driver)
                    driver.find_element(By.CSS_SELECTOR, "a.next").click()
                    print(">>")
            except:
                pass



driver.quit()


# ### 크롤링한 결과를 엑셀에 저장 
wb.save("review_data.xlsx")       






