"""
author : jwk0214

"""

import sys
import os
import pyexcel as px
import random
import time

print("process Start")

start_time = time.time()

# 파괴하려는 엑셀파일이 저장된 폴더이름
directory = sys.argv[1]

# 몇퍼센트의 데이터를 파괴할건가요?
percent = float(sys.argv[2])/100

# 폴더안에 파일 목록 받아오기
files = os.listdir(directory)

# 원래 있던 자료 대신에 집어 넣을 약오르는 단어들을 모아줍니다.
TEERROR = ["고양이","야옹", "야옹이", "미야옹", "팀장님 사랑해요"]

# for 문을 돌면서 파일을 하나씩 읽어온다.
for filename in files:
    # 엑셀 파일이 아닌 경우 건너 뛴다.
    if not filename.endswith(".xlsx"):
        continue

    file_array = px.get_array(file_name=directory + "/" + filename)

    # 엑셀 파일을 위에서부터 한 줄씩 불러옵니다.
    for i in range(len(file_array)):
        # 엑셀 파일을 왼쪽에서부터 한 개씩 불러옵니다.
        for j in range(len(file_array[0])):
            # 확률게임. precent 확률로 당첨
            if random.random() < percent:
                #엑셀 파일 내용물 바꿔치기
                file_array[i][j] = random.choice(TEERROR)
    # 수정이 끝난 파일
    px.save_as(array=file_array, dest_file_name=directory + "/" + filename)

print("process Done")
end_time = time.time()
print("The Job Took" + str(end_time - start_time ) + "second.")