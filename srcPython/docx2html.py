from win32com import client as wc  

#使用win32com
#win32com因更新問題會出現 ImportError: DLL load failed while importing win32api: 找不到指定的模組。
#1. 安裝 pip install pywin32
#2. 使用系統管理員權限開啟cmd
#3. cd至安裝目錄 C:\ProgramData\Anaconda3\Scripts
#4. 用python安裝腳本 python pywin32_postinstall.py -install

#編譯
#1. 安裝編譯套件 pip install pyinstaller
#2. 編譯 pyinstaller -F docx2html.py

#win32com教學
#https://zhuanlan.zhihu.com/p/67543981


def getError():
    import sys

    #exc_info
    type, message, traceback = sys.exc_info()

    #es
    es=[]
    while traceback:
        e={
            'name':traceback.tb_frame.f_code.co_name,
            'filename':traceback.tb_frame.f_code.co_filename,
        }
        es.append(e)
        traceback = traceback.tb_next

    #err
    err={
        'type':type,
        'message':message,
        'traceback':es,
    }

    return err


def j2o(v):
    #json轉物件
    import json
    return json.loads(v)


def o2j(v):
    #物件轉json
    import json
    return json.dumps(v, ensure_ascii=False)


def str2b64(v):
    #字串轉base64字串
    import base64
    v=base64.b64encode(v.encode('utf-8'))
    return str(v,'utf-8')
    

def b642str(v):
    #base64字串轉字串
    import base64
    return base64.b64decode(v).decode('utf-8')


def readText(fn):
    #讀取檔案fn內文字
    import codecs
    with codecs.open(fn,'r',encoding='utf8') as f:
        return f.read()

    
def writeText(fn,str):
    #寫出文字str至檔案fn
    import codecs
    with codecs.open(fn,'w',encoding='utf8') as f:
        f.write(str)


def docx2html(fpIn, fpOut, opt):

    #Dispatch
    app = wc.Dispatch('Word.Application')

    #正式版須隱藏
    app.Visible = False  

    #不詢問使用者
    app.DisplayAlerts = False 

    docIn = None
    prev_view_type = None
    try:

        # Open
        docIn = app.Documents.Open(fpIn)

        # 記錄原本 View.Type
        try:
            aw = docIn.ActiveWindow
            prev_view_type = aw.View.Type
        except:
            prev_view_type = None

        # rng
        rng = docIn.Range()

        # execFontGrow
        execFontGrow = opt.get('execFontGrow', None)

        # Font Grow, 直接格式與樣式皆會被更改
        if isinstance(execFontGrow, int) and execFontGrow > 0:
            for i in range(execFontGrow):
                rng.Font.Grow()  

        # SaveAs2
        WdSaveFormat = 10 # wdFormatFilteredHTML
        docIn.SaveAs2(fpOut, WdSaveFormat)

        # 轉完後恢復 View.Type
        if prev_view_type is not None:
            try:
                docIn.ActiveWindow.View.Type = prev_view_type
            except:
                pass

    finally:

        # Close input doc & quit word
        try:
            if docIn is not None:
                WdDoNotSaveChanges = 0
                docIn.Close(WdDoNotSaveChanges)
        except:
            pass

        try:
            WdSaveOptions = 0  # wdDoNotSaveChanges
            app.Quit(WdSaveOptions)
        except:
            pass


def core(b64):
    state=''

    try:

        #b642str
        s=b642str(b64)

        #j2o
        o=j2o(s)

        #params
        fpIn=o['fpIn']
        fpOut=o['fpOut']

        #opt
        opt={}
        opt['execFontGrow']=o['execFontGrow']

        #docx2html
        docx2html(fpIn, fpOut, opt)

        state='success'
    except:
        err=getError()
        state='error: '+str(err["message"])

    return state


def run():
    import sys

    #由外部程序呼叫或直接給予檔案路徑
    state=''
    argv=sys.argv
    #argv=['','']
    if len(argv)==2:
        
        #b64
        b64=sys.argv[1]
        
        #core
        state=core(b64)
        
    else:
        #print(sys.argv)
        state='error: invalid length of argv'
    
    #print & flush
    print(state)
    sys.stdout.flush()


if True:
    #正式版
    
    #run
    run()
    
    
if False:
    #產生測試輸入b64
    
    #inp
    inp={
        'fpIn':'D:\\- 006 -        開源\\開源-JS-008-4-w-docx2html\\w-docx2html\\test\\docin.docx',
        'fpOut':'D:\\- 006 -        開源\\開源-JS-008-4-w-docx2html\\w-docx2html\\test\\docout.html',
        'execFontGrow': 1,
    }
    # print(o2j(inp))
    
    #str2b64
    b64=str2b64(o2j(inp))
    print(b64)

    #core
    state=core(b64)

    print(state)

