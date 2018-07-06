from lib_attendhsc import AttendHSC, AttendXLS
from selenium.webdriver.common.alert import Alert
from optparse import OptionParser
import time, datetime, pdb


parser = OptionParser()
(options, args) = parser.parse_args()

HOME_URL ='http://attend.hsc.ac.kr'
ID_NUM = "2012070"
PWD_NUM = "1154914"
FILE_NAME = args[0] 

ahsc = AttendHSC('Firefox')
axls = AttendXLS(FILE_NAME)
INFO_SBJ = axls.info_sbj


ahsc.implicitly_wait(10)
ahsc.get(HOME_URL)
ahsc.log_in(ID_NUM, PWD_NUM)

#강의 첫날로 이동한다.
_1st_year_day = ahsc.mk_year_day(INFO_SBJ['year'], INFO_SBJ['1st_day'])
opening_url = ahsc.mk_url(_year=INFO_SBJ['year'], _term=INFO_SBJ['term'], _ls_date=_1st_year_day)

ahsc.get(opening_url)


# waiting_y=''
# while waiting_y !='y':
#     print("해당과목 출석부 페이지로 이동하세요.\n  이동 후 'y' 입력하세요")
#     waiting_y = input().lower()


# sbj_url = aclass.current_url


# 하루에 몇시간들었는지 리스트로 저장
lst_sbj_times = ahsc.get_times_sbj(INFO_SBJ['title_sbj'])

#날짜를 한번씩만 반복되도로 정리한다.

for key in INFO_SBJ:
    print(key, INFO_SBJ[key])


lst_days = set(axls.DF.loc['dates'])
lst_days = list({ x for x in lst_days if x ==x }) # remove NaN
lst_days.sort()

for _d in lst_days: #날짜 단위
    _year_day = ahsc.mk_year_day(INFO_SBJ['year'], _d)
    _url = ahsc.mk_url(_year=INFO_SBJ['year'], _term=INFO_SBJ['term'], _ls_date=_year_day)
#    print(_year_day)
    ahsc.get(_url)
    lst_times_sbj = ahsc.get_times_sbj(INFO_SBJ['title_sbj'])
    _t_lec = len(lst_times_sbj)

    for count, _time in enumerate(lst_times_sbj): #시간 단위반복
        _time.click() # 해당 시간으로 으로 이동
        time.sleep(2)
        ahsc.modal_close() # 대화창 닫기
        time.sleep(2)
        ahsc.select_all_students() #모든 학생을 출석 처리 하기 위해 전체학생 선택
        ahsc.apply_selected('attnd') #선택된 학생을 출석으로 처리

        #  해당 날짜+ 시간 으로 출결지, 학생 목록을 리스트로 저장.
        _lst_attnd = axls.get_Id_student_of('attnd', _d, count+1)
        _lst_nattnd = axls.get_Id_student_of('nattnd', _d, count+1)
        _lst_lattnd = axls.get_Id_student_of('lattnd', _d, count+1)
        print("-----------------------------------------------")
        print('{} : {}: {}'.format(_year_day, count+1, INFO_SBJ['title_sbj']) )
        print("출석: %s 결석: %s 지각: %s"%(len(_lst_attnd), len(_lst_nattnd), len(_lst_lattnd)))
        if len(_lst_nattnd) != 0:
            _selected = False
            while(_selected is False):
                ahsc.select_Id_of_stats('atd_chk', _lst_nattnd)
                _std = ahsc.find_elements_by_xpath("//input[@name='%s' and @value='%s']" %('atd_chk', _lst_nattnd[-1]))
                if _std[0].is_selected():
                    _selected = True

            ahsc.apply_selected('nattnd')

        if len(_lst_lattnd) != 0:
            _selected = False
            while(_selected is False):
                ahsc.select_Id_of_stats('atd_chk', _lst_lattnd)
                _std = ahsc.find_elements_by_xpath("//input[@name='%s' and @value='%s']" %('atd_chk', _lst_lattnd[-1]))
                if _std[0].is_selected():
                    _selected = True

            ahsc.apply_selected('lattnd')

        time.sleep(2)




















































#lst_year_days in format '2017-08-30' 
# year =2017
# lst_year_days = list(map(lambda x : datetime.date(int(year), int(x[0:2]),int(x[2:])), docs_info['dates'][1:]))



