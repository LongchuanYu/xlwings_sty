import os,sys,json,time
from shutil import copy,copytree,rmtree

ROOT_DIR = '\\\\liyang' + '\\' + 'Media\\1' + '\\'
VER_FILE_NAME = 'Ver.txt'
WHITE_NAME_LIST = ['Result','update.py','update.exe']

def get_loacal_dir():
    return os.path.abspath('.')

def est_error(msg):
    print(msg)

def need_update():
    ver_remote = ''
    ver_local = ''
    with open(os.path.join(get_loacal_dir(),'Ver.txt'),'r') as f:
        ver_local = json.load(f)
    with open(os.path.join(ROOT_DIR,VER_FILE_NAME),'r') as fr:
        ver_remote = json.load(fr)
    if not 'Ver' in ver_local or not 'Ver' in ver_remote:
        est_error('版本检查失败...\n1.远程文件/本地文件校验出错...')
        return False
    if ver_remote['Ver'] == ver_local['Ver']:
        return False
    return True

def copy_from_remote(path):
    files = os.listdir(path)
    p=0
    for file in files:
        remote = os.path.join(path,file)
        local = os.path.join(get_loacal_dir(),file)
        full_dir = os.path.join(path,file)
        if file in WHITE_NAME_LIST:
            continue
        if os.path.isdir(full_dir):
            # 如果是文件夹
            copytree(remote,local)
        elif os.path.isfile(full_dir):
            # 如果是文件
            copy(remote,local)
        p+=1
        processing(
            p/len(files)*(50) + 50
        )

def delete_local_files(path):
    files = os.listdir(path)
    p = 0
    for file in files:
        if file in WHITE_NAME_LIST:
            continue
        full_dir = os.path.join(path,file)
        if os.path.isdir(full_dir):
            # 如果是文件夹
            rmtree(full_dir)
        elif os.path.isfile(full_dir):
            # 如果是文件
            os.remove(full_dir)
        p+=1
        processing((p/len(files))*50)


def handle_error(msg):
    print(msg)
    time.sleep(5)
    exit(0)
def processing(num):
    num = int(num)
    sys.stdout.write('升级中...{}% \r'.format(num))
    sys.stdout.flush()
def main():
    if not need_update() :
        print('不需要升级，bye...\n')
        time.sleep(5)
        return
    print('升级准备中...')
    time.sleep(5)
    try:
        delete_local_files(get_loacal_dir())
    except Exception as ex:
        handle_error('升级失败...\n'+ex)
    time.sleep(5)
    try:
        print('获取远程文件中...\n')
        copy_from_remote(ROOT_DIR)
    except Exception as ex:
        handle_error('获取远程文件失败...\n'+ex)
    
    time.sleep(2)
    print('\n\n搞定！！')
    time.sleep(5)

main()