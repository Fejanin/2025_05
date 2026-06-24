import os
import shutil


res = list(filter(lambda x: '.txt' == x[-4:], os.listdir()))
for i in res:
    new_dir = i[:-4]
    print(new_dir)
    os.mkdir(new_dir)
    for i in filter(lambda x: x[-4:] in ('.txt', '.pdf'), os.listdir()):
        print(i)
        if new_dir in i:
            shutil.move(i, f'{new_dir}/{i}')
