## 复制单元格或者Range
ori = [(4,1),(10,11)]
tar = [(5,1)]
sh1.range(*ori).copy(destination=sh2.range(*tar))

def thread_sh2_copy(self,**kwargs):
    '''处理macro弹出框'''
    while 1:
        try:
            app = Application().connect(title_re = SH2_COPY_TITLE_NAME,timeout=1)
            dlg = app.Diaglog
            dlg.Edit.set_edit_text('hello')
            time.sleep(1)
            dlg.Button.click()
        except:
            self.log('[OK] 线程驻留,窗口检测中...')
        time.sleep(1)