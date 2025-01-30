# 
#

import win32com.client
import pycatia


    def CBM(self, rw):
        try:
                # 获取当前启动的 CATIA 应用程序实例
                catia = pycatia.CATIA()      
                # 在这里添加处理按钮点击事件的代码
                # 例如，你可以访问 CATIA 的文档、零件等
                document = catia.active_document
                print(f"当前活动文档: {document.name}")           
        except Exception as e:
                print(f"获取 CATIA 实例时出错: {e}")

    if __name__ == "__main__":
    # start_catia()
