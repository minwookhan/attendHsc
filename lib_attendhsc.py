
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import pandas as pd
import openpyxl as pXLS
import time
import pdb

class AttendHSC(webdriver.Firefox, webdriver.Chrome, webdriver.Ie):

    
    sym_a = '○'
    sym_na = '/'
    sym_la = ' '
    filename =''

    def __init__(self, browser):
        if browser.lower() == "ie":
            webdriver.Ie.__init__(self)
        elif browser.lower() == "chrome":
            webdriver.Chrome.__init__(self)
        else:
            webdriver.Firefox.__init__(self)
        


    def __get_sbj_info__(self, _fname='info.xlsx', _sh='Sheet1'):
        _tmp = pd.read_excel(_fname, sheetname = _sh)
        _time = int(_tmp.iloc[4,54]) # 배당 시간
        _title_sbj = _tmp.iloc[4,15] # 과목명

        return {"Time_lec": _time, "Title_sbj": _title_sbj}

    def findS_xpath(self, _kwd):
        return self.find_elements_by_xpath(_kwd)

    def find_xpath(self, _kwd):
        return self.find_element_by_xpath(_kwd)

    def modal_close(self):
        self.find_element_by_xpath("//button[@id='modelClose']").click()

    def log_in(self, _id, _pwd):
        self.find_xpath("//input[@id='user_id']").send_keys('2012070')
        self.find_xpath("//input[@id='user_pwd']").send_keys('1154914')
        self.find_xpath("//button").click()

    def apply_selected(self, _stats):
        ''''
        '출석', '지각', '결석', 현재 선택된 학생들 대상 일괄변경 버튼 누르기

        '''
        _stat_dic = {'attnd':'10', 'lattnd':'20', 'nattnd':'30'}
        print('APPLY_:', _stats)
        self.find_element_by_xpath("//button[@name='mChange']").click()
        time.sleep(3)

        #self.find_element_by_xpath("//option[@value='%s']"%_stat_dic[_stats]).click()
        select = Select(self.find_element_by_id('atd_typ'))
        select.select_by_value(_stat_dic[_stats])
        self.find_element_by_xpath("//button[text()='저장하기']").click()
        print("PRESSED SAVE")
        try:
            print('SAVE BUTTON PRESSED')
            WebDriverWait(self,10).until(EC.alert_is_present()) 
            alert = self.switch_to.alert
            alert.accept()
            print("Select a SAVE OPTION")
        except TimeoutException:
            print('No Alert')

        try:
            time.sleep(3)
            print("Verification Button Pressed")
            WebDriverWait(self,10).until(EC.alert_is_present())
            alert = self.switch_to.alert
            alert.accept()
            print('상태변경 완료')
        except TimeoutException:
            print('No Alert')



    def mk_year_day(self, _year, _day):
        return str(_year)+'-'+_day[:2]+'-'+_day[2:]

    def mk_url(self,_cps_cd='01' , _year='2017', _term='2', _ls_cd='', _dp_cd='', _grde='', _crse_div='', _s_ls_cd='', _ls_date='2017-09-20'):
                '''
                _cps_cd: 학년
                _year: 년도
                _ls_date: 날짜 2017-08-30
                '''
                url = "http://attend.hsc.ac.kr/lecturer/attendance?search_cps_cd=%s&search_year=%s&search_term=%s&search_ls_cd=%s&search_dp_cd=%s&search_grde=%s&search_crse_div=%s&search_s_ls_cd=%s&search_ls_date=%s" % (_cps_cd, _year, _term, _ls_cd, _dp_cd, _grde, _crse_div, _s_ls_cd, _ls_date)
                return url

    def get_times_sbj(self, _title):
        return self.find_elements_by_xpath("//a[contains(@href, '%s')]"% _title)


    def select_all_students(self):
        _t = self.find_elements_by_xpath("//input[@id='trd_all'] |//input[@id='atd_all'] | //input[@id='abs_all']")
        for i in _t:
            i.click()

    def select_Id_of_stats(self, _stat, _lst_ids ):
        '''
        출석, 지각, 결석각각 테이블에 있는 학생중에서 _lst_ids 에 있는 학번을 체크 한다.
        _stat: 'atd_chk' 출석, 'trd_chk' 지각, 'abs_chk' 결석 테이블 선택
        _lst_ids : 적용할 학번 리스트
        '''
        if len(_lst_ids) == 0 :
            return None
    
        try:
            for i in _lst_ids:
                print(i, " : ", _stat)
                _std = ''
                while _std =='':
                    _std = self.find_element_by_xpath("//input[@name='%s' and @value='%s']" %(_stat, i))
                while not(_std.is_selected()):
                    _std.click()
                del _std
                time.sleep(1)
        except:
            pass
        return True


    
class AttendXLS:

    sym_a = '○'
    sym_na = '/'
    sym_la = ' '

    def __init__(self, _fn ='2018_info.xlsx', _sheet='Sheet1'):
        self.xlsFile = _fn
        self.nameSheet = _sheet
        self.info_sbj = self._get_Info_sbj(_fn, _sheet)
        self.DF = self.__read_xls__( _fn, _sheet)



    def _mk_num_idx(self, x,  _time):
        _t = x%_time
        return _t if( _t!=0) else _time


    def print_file_info(self):
        
        _tmp = pd.read_excel(self.xlsFile,  self.nameSheet)
        _time = int(_tmp.iloc[4,54]) # 배당 시간
        _title_sbj = _tmp.ilo[4,15] # 과목명

        print(self.xlsFile, self.nameSheet, "\n")
        _info = {'Title_Subject': _title_sbj, 'times':_time }
        for key in _info:
            print(key,_info[key])

        return _info
 


    def __read_xls__(self,_fname='info.xlsx', _sh='Sheet1'):
        ''' _fname : 엑셀파일
        _sh : sheet name
        returns:
        rsh2: 전체 출석지각 정보 테이블
        lst_attnd: 출석 학번리스트
        lst_nattnd4 결석 학번 리스트
        lst_lattnd: 지각 학번 리스ㅌ
        _title_sbj: 과목명
        '''
        sh =  pd.read_excel(_fname, sheet_name=_sh,index_col=0, header=7, usecols='G:BN', skipfooter=3, skiprows=1).dropna(how='all', axis=1)
        _time = self.info_sbj['total_time']

        rsh = sh.dropna(how='all', axis=0)
        rsh2 = rsh.rename(columns = {"성명":0})
        idx_lst = rsh2.index.tolist() # student num id 
        idx_lst[0] = 0.0

        idx_lst = [ str(x)[:-2] for x in idx_lst]

        rsh2.index = idx_lst

        a = list(rsh2.index.map(str))
        b = [ x[:-2] for x in a]
        lst_col = rsh2.columns.tolist()
        #y = list(map(mk_num_idx, lst_col, [_time]*len(lst_col)))


        #  시간표 번호를 컬럼에 추가
        import pdb
     #   pdb.set_trace()


        lec_time  =[str(x) for x in  list(map(self._mk_num_idx, lst_col, [_time]*len(lst_col)))]
        
        rsh2.loc['lec_time']= lec_time
        lst_idx = rsh2.index.tolist()
        lst_idx[0]='dates'
        rsh2.index = lst_idx
#        pdb.set_trace()
        rsh2.loc['dates'] = [ x if isinstance(x,str) else str(x) for x in rsh2.loc['dates']]
        rsh2.columns = (rsh2.loc['lec_time']+rsh2.loc['dates'])  # 
        return rsh2

    def Alert_diag(self,_ans):
        '''
        팝업 대화상자 처리
        '''
        if _ans.lower() =="yes":
            Alert(self).accept()
        else:
            Alert(self).dismiss()




    def get_Id_student_of(self,_stat , _day, _t):
        '''
        해당날짜의 출결지 학생의 학번을 리스트로 리턴
        _stat: 'attd', 'nattd', 'lattd'
        '''
        _lst_stat =['attnd', 'nattnd', 'lattnd']
        assert _stat in _lst_stat, print("ERROR: stat code is not in list")
        # lst_attnd = DF[DF[6] == self.sym_a].index.tolist()
        # lst_nattnd = DF[DF[6] == self.sym_na].index.tolist()
        # lst_lattnd = DF[DF[6] == self.sym_la].index.tolist()

     #   pdb.set_trace()
        if _stat == 'attnd':
            return self.DF[self.DF[str(_t)+str(_day)]== self.sym_a].index.tolist()

        elif _stat == 'nattnd':
            return self.DF[self.DF[str(_t)+str(_day)]== self.sym_na].index.tolist()

        else :
            return self.DF[self.DF[str(_t)+str(_day)]== self.sym_la].index.tolist()
    def _get_Info_sbj(self, _fname= 'info.xlsx', _sh= 'Sheet1'):

        Term = {
            'title_sbj':'과목이름',
            'year': '년도',
            'term': '학기',
            'code_sbj': '과목코드',
            'div_class': '분반',
            'pt_cls': '학점',
            'time_lec': '강의시간',
            'time_prac': '실습시간',
            'num_pop': '수강인원',
            'id_prof': '교수ID',
            'name_prof': '교수이름',
            '1st_day' : '첫수업날짜' # 첫 화면으로 이동하기 위해 필요
        }

        inf_cell_addr = {
            'year': 'B4',
            'term': 'H4',
            'code_sbj': 'K7',
            'div_class': 'AR7',
            'pt_cls': 'BA7',
            'time_lec': 'BI8',
            'time_prac': 'BN8',
            'num_pop': 'BS8',
            'id_prof': 'CA4',
            'name_prof': 'CA4',
            '1st_day' : 'L10',
            'title_sbj': 'N7'
        }

        df = pXLS.load_workbook(_fname)
        sh = df.get_sheet_by_name(_sh)
        _info_sbj = {}

        for key in inf_cell_addr:
            _info_sbj[key] = sh[inf_cell_addr[key]].value

        # import pdb
        # pdb.set_trace()
        _info_sbj['total_time'] = _info_sbj['time_lec']+_info_sbj['time_prac']
        return _info_sbj


    




