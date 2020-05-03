import xlwings as xw
import threading as trd
import time,win32api,win32con,os
import pyperclip as pcp
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse
# ----------------------------------------------
# 共通
SH1_EXCEL_NAME = '1.xlsx'
SH1_1_NAME = '1'
FLAG_NAME = '%end%'
DELETE_SING = '－'
# sh1 全局变量
R_DATA = 7      # 数据航
C_C7 = 25        # C7的位置
C_COUNT = 26     # 项目总数的位置
C_SHUKINOU = 2
C_DAI = 4       # 大项目
C_CHU = 6       # 。。。
C_SHO = 8

# sh2 全局变量
SH2_EXCEL_NAME = '2.xlsm'
SH2_NAMAE_DATA_NAME = '生ダーた_P用'
SH2_TOCASE_NAME = 'To Case'
SH2_COPY_TITLE_NAME = '大機能'
SH2_SET_TITLE_NAME = ''
R_DATA_SH2 = 3
R_DATA_SH2_1 = 3
C_COUNT_SH2 = 8
C_DATA_SH2_1 = 9
C_CASE_SH2_1 = 12

# sh3 全局
SH3_CHECK_TITLE_NAME = '確認'
SH3_AUTHOR = 'liyang'
CH3_GEN_TITLE_NAME = 'Microsoft Excel'
SH3_EXCEL_NAME = '3.xlsm'
R_DATA_SH3 = 3
C_CASE_SH3 = 17

#--------------------------------------------------

class Mywork():
    def __init__(self):
        self.tools = Tools()

        self.books = self.tools.get_AllBooks()
        for book in self.books:
            if book.name == SH1_EXCEL_NAME:
                self.b1 = book
            elif book.name == SH2_EXCEL_NAME:
                self.b2 = book
            elif book.name == SH3_EXCEL_NAME:
                self.b3 = book
        self.sh1 = self.b1.sheets['1']
        self.sh2 = self.b2.sheets[SH2_NAMAE_DATA_NAME]
        self.sh3 = self.b3.sheets[0]
        self.sh2_1 = self.b2.sheets[SH2_TOCASE_NAME]
        self.sh2_1_daikinou_name = ''
        self.log = self.tools.log_save
        self.clickWindow = False
        self.sh3CheckButtonFlag = False
        self.sh3GenButtonFlag = False
        self.author_name = ''
    def main(self):
        
        errors = {}
        print(self.tools.doc_print())
        ret = input('输入你的名字，按回车键继续:')
        if ret=='':
            print('\n再见，你什么都没输入！')
            time.sleep(3)
            return
        else:
            self.author_name = ret
        # 清理工作簿，删除C7非对应的行
        # max_row = self.tools.get_MaxRowByEndFlag(self.sh1)
        # for row in range(max_row , R_DATA,-1):
        #     t = self.sh1.range((row,C_C7))
        #     if self.sh1.range((row,C_C7)).value == '－':
        #         self.sh1.range('{id}:{id}'.format(id=row)).delete()
        # 检查是否有end标志
        if FLAG_NAME not in self.sh1.range((1,1),(self.tools.get_MaxRowBySheet(self.sh1),1)).value:
            return self.emit_error('没有结束标志符')

        # 加入线程监听
        # 监听sh2 的 copy
        trd.Thread(target=self.set_window_top_and_sendkey_1).start()
        # 监听sh3的 Check Button
        trd.Thread(target=self.set_window_top_and_sendkey_2).start()
        # 监听sh3的 Gen Button
        trd.Thread(target=self.set_window_top_and_sendkey_3).start()

        # 复制
        pre_row = R_DATA
        for row in range(R_DATA+1,10000):
            val = self.sh1.range((row,C_SHUKINOU)).value

            if val or self.sh1.range((row,1)).value == FLAG_NAME:
                
                self.sh2_1_daikinou_name = (self.sh1.range((pre_row,C_SHUKINOU-1)).value or '') + ' ' +\
                    (self.sh1.range((pre_row,C_SHUKINOU)).value or '')
                # self.sh1.range(('A1')).value = self.sh2_1_daikinou_name
                print('----------------执行'+self.sh2_1_daikinou_name+'------------------')
                ### 
                #   Core 
                ###
                #---------------------------------------------  1.xlsx
                # 选择生ダーた_P用表单，执行clear命令
                self.b2.activate(steal_focus=True)
                self.sh2.select()
                self.b2.macro('Clear2.Clear2')()
                # 执行复制
                self.do_task_1_copy(start=pre_row,end=row-1)
                time.sleep(2)
                #----------------------------------------------- 2.xlsm
                # 跳转到ToCase表单并执行clear命令
                
                self.sh2_1.select()
                self.b2.macro('Clear2.Clear2')()
                time.sleep(2)
                self.log('[{}] 2.xlsm 清理完毕...'.format(self.sh2_1_daikinou_name))
                # 执行copy命令
                self.clickWindow= True 
                self.b2.macro('Copy1.Copy1')()
                self.log('[{}] 2.xlsm copy完毕...'.format(self.sh2_1_daikinou_name))
                time.sleep(2)
                # 执行整理命令
                self.b2.macro('seiri.seiri')()
                self.log('[{}] 2.xlsm 整理完毕...'.format(self.sh2_1_daikinou_name))
                time.sleep(5)
                # 仕向设定
                self.b2.macro('kyoutsu.kyoutsu')()
                self.log('[{}] 2.xlsm 仕向设定完毕...'.format(self.sh2_1_daikinou_name))
                time.sleep(2)
                
                # 执行tocase
                self.b2.macro('tocase.to_case')()
                self.log('[{}] 2.xlsm ToCase完毕...'.format(self.sh2_1_daikinou_name))
                time.sleep(2)
                
                #----------------------------------------------- 3.xlsm

                # 复制
                self.b3.activate(steal_focus=True)
                self.do_task_sh3_copy()
                self.log('[{}]  3.xlsm copy完毕...'.format(self.sh2_1_daikinou_name))
                # macro Check
                self.sh3CheckButtonFlag = True
                self.b3.macro('keisen_Check')()
                self.log('[{}]  3.xlsm 网挂Check执行完毕...'.format(self.sh2_1_daikinou_name))
                time.sleep(2)
                # 名字入力 self.author_name
                rng = self.sh3.range((R_DATA_SH3,C_CASE_SH3+1),(self.tools.get_MaxRowBySheet(self.sh3),C_CASE_SH3+1))
                rng.columns[0].value = self.author_name
                self.log('[{}] 3.xlsm 作者名字入力成功'.format(self.sh2_1_daikinou_name))
                time.sleep(1)

                # macro Gen
                self.sh3GenButtonFlag = True
                self.b3.macro('GenWithOutCol')()
                self.log('[{}]  Gen执行完毕...'.format(self.sh2_1_daikinou_name))
                time.sleep(2)
                
                #--------------------------------------------文件操作
                # root = os.path.abspath('.')
                root = os.path.join(os.path.abspath('.'),'Result')
                dir_ori_file = os.path.join(root,'3_TestCase.xml')
                dir_tar_file = os.path.join(root,'{}.xml'.format(self.sh2_1_daikinou_name or 'null'))
                if not os.path.exists(root):
                    os.makedirs(root)
                if not os.path.exists(dir_tar_file):
                    os.rename(dir_ori_file,dir_tar_file)
                else:
                    os.rename(
                        dir_ori_file , 
                        os.path.join(root,'{}.xml'.format(
                            self.sh2_1_daikinou_name+str(time.strftime("%H%M%S", time.localtime()))
                        ))
                    )
                self.log('[{}]  文件创建成功...'.format(self.sh2_1_daikinou_name))

                ### 
                #   End Core 
                ###
            
                pre_row = row
            if self.sh1.range((row,1)).value == FLAG_NAME:
                break
        input('已完成') 

    def do_task_1_copy(self,**kwargs):
        '''
        kwargs:{
            start:int,
            end:int
        }
        '''
        start = kwargs['start']
        end = kwargs['end']
        # 复制中、小、详细
        rng = [
            (start,C_DAI-1),
            (end,C_SHO)
        ]
        count_rng = [
            (start,C_COUNT),
            (end,C_COUNT)
        ]
        self.sh1.range(*rng).copy(destination=self.sh2.range((R_DATA_SH2,1)))
        self.sh1.range(*count_rng).copy()
        self.sh2.range(R_DATA_SH2,C_COUNT_SH2).paste(paste='values_and_number_formats')

    def do_task_sh3_copy(self):

        self.sh3.range(
            (R_DATA_SH3,1),(self.tools.get_MaxRowBySheet(self.sh3),100)
        ).delete()

        self.sh2_1.range(
            (R_DATA_SH2_1,1),(self.tools.get_MaxRowBySheet(self.sh2_1),C_DATA_SH2_1)
        ).copy(
            destination=self.sh3.range((R_DATA_SH3,1))
        )
        self.sh2_1.range(
            (R_DATA_SH2_1,C_CASE_SH2_1),(max(self.tools.get_MaxRowBySheet(self.sh2_1),R_DATA_SH2_1),C_CASE_SH2_1)
        ).copy()
        self.sh3.range((R_DATA_SH3,C_CASE_SH3)).paste(paste='values_and_number_formats')


    def set_window_top_and_sendkey_3(self):
        '''把sh3的网线check focus并sendkey'''
        while 1:
            if self.sh3GenButtonFlag:
                try:
                    wd2 = Application().connect(title_re=CH3_GEN_TITLE_NAME,timeout=1)
                    x = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                    y = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
                    mouse.click(button='left',coords=(x-100,y-200))
                    # wd1.top_window().set_focus()
                    time.sleep(1)
                    
                    wd2.top_window().set_focus()
                    time.sleep(1)
                    send_keys('~') 
                    self.sh3GenButtonFlag = False
                except:
                    print('3.xlsm窗口置顶等待中...')
            time.sleep(1)  

    def set_window_top_and_sendkey_2(self):
        '''把sh3的网线check focus并sendkey'''
        while 1:
            if self.sh3CheckButtonFlag:
                try:
                    wd2 = Application().connect(title_re=SH3_CHECK_TITLE_NAME,timeout=1)
                    x = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                    y = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
                    mouse.click(button='left',coords=(x-100,y-200))
                    # wd1.top_window().set_focus()
                    time.sleep(1)
                    
                    wd2.top_window().set_focus()
                    time.sleep(1)
                    send_keys('n') 
                    self.sh3CheckButtonFlag = False
                except:
                    print('3.xlsm窗口置顶等待中...')
            time.sleep(1)   
    def set_window_top_and_sendkey_1(self):
        '''把sh2的copy focus并sendkey'''
        while 1:
            if self.clickWindow:
                try:
                    # wd1 = Application().connect(title_re=SH2_EXCEL_NAME,timeout=1)
                    wd2 = Application().connect(title_re=SH2_COPY_TITLE_NAME,timeout=1)

                    pcp.copy(self.sh2_1_daikinou_name)
                    time.sleep(1)
                    x = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                    y = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
                    mouse.click(button='left',coords=(x-100,y-200))
                    # wd1.top_window().set_focus()
                    time.sleep(1)
                    
                    dlg = wd2.top_window()
                    dlg.set_focus()
                    time.sleep(1)
                    dlg.Edit.set_edit_text(self.sh2_1_daikinou_name)
                    dlg.Button0.click()
                    self.clickWindow = False
                except:
                    print('[OK] Excel - 2.xlsm 窗口置顶等待中...')
            time.sleep(1)


    def test(self):
        # self.do_copy(sh1=1,sh2=2)
        # app = Application().connect(title_re = SH2_COPY_TITLE_NAME,timeout=1)
        # dlg = app.window(best_match=SH2_COPY_TITLE_NAME)
        # dlg['F3 Server 53e900002'].draw_outline(colour = 'red')
        # self.clickWindow = True
        # self.clickWindowName=SH2_EXCEL_NAME
        # self.sh2_1_daikinou_name = 'testetestes'
        # trd.Thread(target=self.set_window_top_and_sendkey_1).start()
        # wb = xw.books.active
        # wb.macro('Copy1.Copy1')()
        # self.b2.macro('tocase.to_case')()
        # z=1
        # wd2 = Application().connect(title_re='大機能',timeout=1).top_window()
        # wd2.set_focus()
        # # mouse.move((x-100,y-200))
        root = os.path.abspath('.')
        dir_result = os.path.join(root,'Result')
        if not os.path.exists(dir_result):
            os.makedirs(dir_result)
        
class Tools():
    def get_AllBooks(self):
        '''获取所有已打开的工作簿'''
        books = xw.apps.active.books
        return books
    def get_MaxRowByEndFlag(self,sh):
        max_row = sh1.range('A1:A{}'.format(self.tools.get_MaxRowBySheet(sh1))).value.index(FLAG_NAME)
        return maxrow
    def get_MaxRowBySheet(self,sh):
        '''获取sheets的最大行'''
        t = int(
            sh.used_range.address.split('$')[-1]
        )
        return t
    def is_RowBlank(self,sh,row):
        '''判断指定sheet的row是否空'''
        return any(sh.range('{id}:{id}'.format(id=row)).value)
    def emit_error(self,msg):
        print(msg)
    def log_save(self,msg):
        print(msg)
    def doc_print(self):
        msg = '''
        注意点：
        -----------------------------------------
        1. 准备三个excel，按照顺序分别命名为
            1.xlsx （需要把sheet名改为：1）
            2.xlsm
            3.xlsm
        2. 打开这三个Excel,全部放到窗口右边
        3. 1.xlsx 添加结束标记：%end% 
        4. 筛选出C7不支持的行并删除！
        '''
        return msg



def test():
    pass



try:
    m = Mywork()
    m.main()
except Exception as ex:
    print('\n\n异常...5秒后退出\n')
    print(ex)
    time.sleep(5)
