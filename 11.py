import win32com.client
from pycatia import CATIA

def initialize_array():
    att_usp_Names = ["cm", "iBodys", "iMaterial", "iDensity", "iMass", "iThickness"]
    return att_usp_Names
def my_prop(rw):
    try:
        catia = CATIA()
        o_doc = catia.active_document
        root_prd = o_doc.product
        if not root_prd:
            print("请打开CATIA产品，再运行本程序")
            return
        catia.visible = True
        xl_app = win32com.client.Dispatch("Excel.Application")
        xl_sht = xl_app.ActiveSheet
        xl_sht.Columns(2).NumberFormatLocal = "0.000"
        xl_sht.Columns(1).NumberFormatLocal = "0.00"
        xl_sht.Rows(1).NumberFormatLocal = "0"
        xl_app.visible = True
        start_row = 3
        start_col = 3
        if rw == 0:
            b_dict = {}
            ini_prd(root_prd, b_dict)
        elif rw == 1:
            prd2wt = whois2rv()
            if not prd2wt:
                print("没有要读取的产品")
                return
            Prd2Read = prd2wt
            rng = xl_sht.Range(xl_sht.Cells(3, 1), xl_sht.Cells(50, 14))
            rng.ClearContents
            curr_row = start_row
            Arry2sht(info_prd(Prd2Read), xl_sht, curr_row)
            children = Prd2Read.Products
            for i in range(1, children.Count + 1):
                curr_row = i + start_row
                Arry2sht(info_prd(children.Item(i)), xl_sht, curr_row)
            Prd2Read = None
        elif rw == 2:
            if not prd2wt:
                my_prop(1)
                print("请填写修改信息并重新按修改按钮")
                return
            print(f"你要修改的分总成是 {prd2wt.PartNumber} 是否继续")
            Prd2Rv = prd2wt
            curr_row = start_row
            AttModify(Prd2Rv, sht2Arry(xl_sht, curr_row))
            children = Prd2Rv.Products
            for i in range(1, children.Count + 1):
                curr_row = i + start_row
                AttModify(children.Item(i), sht2Arry(xl_sht, curr_row))
            Prd2Rv = None
        elif rw == 3:
            Assmass(root_prd)
        elif rw == 4:
            xl_sht = getsht(xl_app)
            curr_row = start_row
            LV = 1
            recurPrd(root_prd, xl_sht, curr_row, LV)
            LvMg(xl_sht)
        print("已执行，0.5s 后自动关闭对话框")
    except Exception as e:
        print(f"程序执行异常，请排查错误后重新执行，0.5s后将退出: {e}")
# 以下是其他函数的定义，需要根据具体功能进行实现
def ini_prd(root_prd, b_dict):
    pass
def whois2rv():
    pass
def info_prd(prd):
    pass
def Arry2sht(data, xl_sht, curr_row):
    pass
def AttModify(prd, data):
    pass
def Assmass(root_prd):
    pass
def getsht(xl_app):
    pass
def recurPrd(root_prd, xl_sht, curr_row, LV):
    pass
def LvMg(xl_sht):
    pass
if __name__ == "__main__":
    att_usp_Names = initialize_array()
    my_prop(0)  # 示例调用，根据需要修改参数
