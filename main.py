from deal_data import deal_data
import threading
from queue import Queue
from to_excel import to_excel
from time import time
from tqdm import tqdm

t1 = time()

# user修改區
student_Category = '碩士'  # '碩士' or '博士'
start_year = 99  # int
end_year = 109  # int
advisor_name_ = ['xxx', 'xxx']  # [str, ]
excel_name = 'paperFinder.xlsx'  # str.xlsx


# 資料處理區
threads = []
q = Queue()
advisor_name = []

## 去重複
for a in advisor_name_:
    if a not in advisor_name:
        advisor_name.append(a)

## 多線程
for i in range(len(advisor_name)):
    threads.append(threading.Thread(target=deal_data,
                                    args=(start_year, end_year, student_Category, advisor_name[i], q)))
    threads[i].start()

print('抓取資料中...')
for t in tqdm(threads):
    t.join()

## 將佇列資料調整成原始順序
results = [None for _ in range(len(advisor_name))]
for _ in range(len(advisor_name)):
    temp = q.get()
    results[advisor_name.index(temp[5])] = temp

## 寫入excel
to_excel(results, excel_name, advisor_name)
t2 = time()
print('完成作業')
print('耗時：' + str(round(t2 - t1, 2)) + 's')
