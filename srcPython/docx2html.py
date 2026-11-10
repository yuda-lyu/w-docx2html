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


def isSamePath(fp1, fp2):
    #兩路徑是否為同一檔(大小寫, 相對路徑, 符號連結與短檔名皆正規化); 無法判定時視為不同
    import os
    try:
        return os.path.normcase(os.path.realpath(fp1)) == os.path.normcase(os.path.realpath(fp2))
    except:
        return False


def countOtherDocs(app, docsLeft):
    #實例內非本行程之文件數, 含受保護檢視之視窗(其不在Documents內); docsLeft為本行程關閉失敗而仍開著者, 以COM物件身分比對(不以路徑: 他人可能重開同一路徑)
    #回傳(數量, 狀態): 狀態'ok'為已判定, 'gone'為實例已不存在(如已被他人Quit), 'unknown'為無法判定(忙碌或讀取失敗)
    import time
    import pythoncom

    #實例已不存在之錯誤: RPC_S_SERVER_UNAVAILABLE, RPC_S_CALL_FAILED, RPC_E_DISCONNECTED
    hrsGone = (-2147023174, -2147023170, -2147417848)

    for i in range(10):
        try:
            try:
                n = app.ProtectedViewWindows.Count
            except AttributeError:
                n = 0 #無此集合之Word版本
            for d in app.Documents:
                if not any(d == o for o in docsLeft):
                    n += 1
            return n, 'ok'
        except pythoncom.com_error as e:
            if getattr(e, 'hresult', None) in hrsGone:
                return 0, 'gone'
        except:
            pass
        time.sleep(0.5) #他人呼叫中(忙碌)等暫時失敗, 稍候重試
    return None, 'unknown'


def closeOwnWord(app, docs):
    #關閉本行程開啟之文件, 再結束本行程所開之Word實例; 實例內另有他人之文件或無法判定時不結束
    #why: DispatchEx所開之實例仍登錄於ROT, 他行程以Dispatch(先GetActiveObject)可連上並於其中開文件(實測), 此時Quit會關掉對方之文件且不儲存;
    #  連上者(舊寫法)用畢自會Quit, 故保留實例; 無法判定時亦保留(寧可留下實例, 不關他人之文件); 盤點與Quit之間他人仍可連入, 此窗口無法消除
    import sys

    #Close, 關閉失敗者仍屬本行程
    WdDoNotSaveChanges = 0 #wdDoNotSaveChanges
    docsLeft = []
    for doc in reversed(docs):
        try:
            doc.Close(WdDoNotSaveChanges)
        except:
            docsLeft.append(doc)

    #nOther
    nOther, st = countOtherDocs(app, docsLeft)
    if st == 'gone':
        return
    if st != 'ok' or nOther > 0:
        print('Word instance is kept since ' + (f'{nOther} document(s) opened by another client' if st == 'ok' else 'its documents can not be checked'), file=sys.stderr)
        return

    #Quit
    WdSaveOptions = 0 #wdDoNotSaveChanges=0(不儲存擱置中變更)
    try:
        app.Quit(WdSaveOptions)
    except:
        pass


def docx2html(fpIn, fpOut, opt):

    #DispatchEx, 每次另開獨立之Word實例(CoCreateInstanceEx, 實測已有其他Word時亦另起WINWORD)
    #why: Dispatch會先以GetActiveObject連上已在執行之Word(他行程之轉檔或使用者開啟者), 轉完Quit即關掉對方之文件且不儲存, 兩邊同時轉檔則互相失敗
    app = wc.DispatchEx('Word.Application')

    #docs, 本行程開啟之文件, 供finally以closeOwnWord關閉
    docs = []

    docIn = None
    prev_view_type = None

    #任一步驟(含下列設定)例外時亦須關閉已開啟之文件並結束本行程所開之Word, 否則殘留之WINWORD會鎖住輸入檔與輸出檔
    try:

        #正式版須隱藏
        app.Visible = False

        #不詢問使用者
        app.DisplayAlerts = False

        # Open, 以唯讀開啟(Open之參數依序為FileName, ConfirmConversions, ReadOnly): 本轉檔器只SaveAs2至fpOut, 不存回輸入檔;
        #  惟fpOut即輸入檔時須可寫開啟, 與改前行為相同
        #why: 各轉檔各自一個Word實例, 兩個轉檔同時以可寫方式開同一輸入檔時, 後開者跳出「檔案使用中」對話框(可見)而停住
        docIn = app.Documents.Open(fpIn, False, not isSamePath(fpIn, fpOut))
        docs.append(docIn)

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

        # Close input doc & quit word(只關本行程之文件, 只結束本行程所開之實例; 實例內有他人之文件或無法判定時不結束)
        closeOwnWord(app, docs)


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

