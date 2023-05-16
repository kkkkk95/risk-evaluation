import os
import pandas as pd
import requests
from PIL import Image
import streamlit as st
from streamlit_lottie import st_lottie
import shutil
from datetime import date
import subprocess
import sys
import platform
import webbrowser
class home:
    def __init__(self,options,selected_options):
        #self.img_sphere = Image.open("images/sphere.jpg")
        #self.img_phase_separation = Image.open("images/phase_separation.jpg")
        #self.img_nano = Image.open("images/nano.jpg")
        # 定义要打开的Word文档路径
        self.doc_path_program = r'database/新开航线风险分析评价及4D_15能力评估工作程序.docx'
        self.doc_path_workflow = r'database/自动化生成流程.docx'
        self.report_path = r'templet/风险评价报告表模版.xlsx'
        self.dangerlist_path = r'templet/危险源清单模板.xlsx'
        self.sysrecord_path = r'templet/系统与工作分析记录表模板.xlsx'
        self.database_path = r'database/航班动态监控室危险源数据库（对应公司三层级、中心部门级危险源数据库）.xlsx'
        self.resultfile = os.path.join(os.getcwd(), 'result')
        self.database = None
        self.options=options
        self.selected_options=selected_options

    def empty_dir(self, dir_path):
        if platform.system() == 'Windows':
            os.system('del /q ' + dir_path + '\\*')
        elif platform.system() == 'Linux':
            os.system('rm -rf ' + dir_path + '/*')
        else:
            # For macOS
            os.system('rm -rf ' + dir_path + '/*')

    def open_file(self,file_path):
        if platform.system() == 'Windows':
            os.startfile(file_path)
        elif platform.system() == 'Linux':
            subprocess.run(["xdg-open", file_path])
        else:
            # For macOS
            subprocess.run(["open", file_path])

    def load_lottieurl(self, url):
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()

    # Use local CSS
    def local_css(self, file_name):
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

    def run(self):
        self.local_css(r"style/style.css")
        # ---- LOAD ASSETS ----
        self.lottie_coding = self.load_lottieurl("https://assets1.lottiefiles.com/packages/lf20_2xoRs4A4MD.json")
        # ---- WHAT I DO ----

        with st.container():
            st.write("---")
            left_column, right_column = st.columns(2)
            with left_column:
                st.header("What I do")
                st.write("新开航线风险分析评价及4D/15能力评估工作程序")
                st.write(
                    """
                    系统与工作分析记录表：\n
                    系统与工作分析记录表中内容需结合航线风险源清单中所涉风险源信息及控制措施对应填写。\n
                    风险评价报告表：\n
                    风险评价报告表中内容需结合航线风险源清单中所涉风险源信息及控制措施对应填写。\n
                    部门风险管控实施检查单：\n
                    新开航线风险源清单、系统与工作分析记录表和风险评价报告表完成后，需向部门负责人、航线风险分析牵头人及一线生产员工征求修改意见及建议，修改完成且一致通过后需填写纸质部门风险管控实施检查单并留存记录。\n
                    """
                )
            with right_column:
                st_lottie(self.lottie_coding, height=300, key="coding")
                # 创建一个链接以打开Word文档
                if st.button("查看工作程序详情"):
                    self.open_file(os.path.abspath(self.doc_path_program))
                if st.button("查看流程"):
                    self.open_file(os.path.abspath(self.doc_path_workflow))
        # 导入数据
        
        st.write("---")
        select_column,left_column, right_column = st.columns(3)   
        with select_column:
            n=1
            key=0
            while n<len(self.options):
                selected = st.selectbox('请选择一个数据库', self.options,key=str(key+1))
                self.selected_options.append(selected)
                n+=1
                key+=1

        with right_column:
            if st.button('查看数据库'):
                if '监控数据库' in self.selected_options:
                    database_abs_path = os.path.abspath(self.database_path)
                    #self.open_file(database_abs_path)
                    webbrowser.open('https://github.com/kkkkk95/Risk_Evaluate/raw/main/database/%E8%88%AA%E7%8F%AD%E5%8A%A8%E6%80%81%E7%9B%91%E6%8E%A7%E5%AE%A4%E5%8D%B1%E9%99%A9%E6%BA%90%E6%95%B0%E6%8D%AE%E5%BA%93%EF%BC%88%E5%AF%B9%E5%BA%94%E5%85%AC%E5%8F%B8%E4%B8%89%E5%B1%82%E7%BA%A7%E3%80%81%E4%B8%AD%E5%BF%83%E9%83%A8%E9%97%A8%E7%BA%A7%E5%8D%B1%E9%99%A9%E6%BA%90%E6%95%B0%E6%8D%AE%E5%BA%93%EF%BC%89.xlsx')
                else:
                    st.warning('未选择数据库')
        with left_column:
            if st.button('导入数据库和模板'):
                with st.spinner('正在处理数据，请稍等...'):
                    # 在每次复制前清空目标文件夹
                    self.empty_dir(self.resultfile)
                    if '监控数据库' in self.selected_options:
                        database_abs_path = os.path.abspath(self.database_path)
                        self.database = pd.read_excel(database_abs_path, header=0, skiprows=1)
                        st.session_state.database=self.database
                        st.success('监控数据库导入成功！')
                        
                    #待加入其他数据库
                    #elif '' in selected_option:

                    else:
                        st.warning('请选择正确的数据库！')

if __name__ == "__main__":
    st.set_page_config(page_title="new_line_analyze", page_icon="🏠")

    # 初始化全局配置
    if 'first_visit' not in st.session_state:
        st.session_state.first_visit=True
        st.session_state.flight_type=''
        st.balloons()
          
    else:
        st.session_state.first_visit=False
        
    home=home(['监控数据库','#待加入'],[])
    home.run()
