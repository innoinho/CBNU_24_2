import requests
from bs4 import BeautifulSoup
import time
import datetime

def check_new_posts(url, last_post_number):
    try:
        response = requests.get(url)
        response.raise_for_status()  # HTTP 오류 발생 시 예외 발생
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 테이블 찾기
        table = soup.find('table', class_='board_list')
        if not table:
            print("테이블을 찾을 수 없습니다.")
            return last_post_number

        # tbody 찾기
        tbody = table.find('tbody')
        if not tbody:
            print("tbody를 찾을 수 없습니다.")
            return last_post_number

        # 첫 번째 행 찾기
        latest_post = tbody.find('tr')
        if not latest_post:
            print("게시물을 찾을 수 없습니다.")
            return last_post_number

        # 가장 최근 공고의 번호 찾기
        number_cell = latest_post.find('td')
        if not number_cell:
            print("게시물 번호를 찾을 수 없습니다.")
            return last_post_number

        latest_post_number = int(number_cell.text.strip())
        
        if latest_post_number > last_post_number:
            new_posts = []
            for row in tbody.find_all('tr'):
                post_number = int(row.find('td').text.strip())
                if post_number > last_post_number:
                    title = row.find('td', class_='subject').text.strip()
                    new_posts.append(f"새로운 공고: {post_number}. {title}")
            
            print("\n".join(new_posts))
            return latest_post_number
        return last_post_number
    except requests.RequestException as e:
        print(f"요청 중 오류 발생: {e}")
        return last_post_number
    except Exception as e:
        print(f"예상치 못한 오류 발생: {e}")
        return last_post_number

url = "https://www.bizinfo.go.kr/web/lay1/bbs/S1T122C128/AS/74/list.do"
last_post_number = 0

while True:
    print(f"확인 시간: {datetime.datetime.now()}")
    last_post_number = check_new_posts(url, last_post_number)
    time.sleep(30)  # 300초(5분)마다 확인

