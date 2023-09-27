#############################################     라이브러리 호출     ###################################################
import gspread as gs
import pandas as pd
import openpyxl

##############################################     필요 함수 정의     ###################################################
# 시트 호출 함수 정의
def load_sheets(url, sheet):
    df_sheet = url.worksheet(sheet).get_all_values()
    df_return = pd.DataFrame(df_sheet[1:], columns=df_sheet[0])
    return df_return

# 사이버캠퍼스 명단 과정명 수정 함수 정의
def change_course(row):
    if '경제' in row['과정명']:
        return '경제톡톡'
    elif 'WM' in row['과정명']:
        return 'WM센터'
    else:
        return row['과정명']

# 한 개만 신청한 인원들 명단 정리 함수 정의
def modify_single(df, course):
    df_modify = df[df['과정명'].str.contains(course)]
    df_modify = df_modify.drop(columns=['횟수','과정명'])
    df_modify['과정명'] = course
    return df_modify

###########################################     입과입력 과정코드 입력     ################################################
# 입과입력 양식 제작을 위한 과정코드 입력
input_all = input('전체과정 타겟홍보 과정코드를 입력하세요 : ')
input_eco = input('경제톡톡 타겟홍보 과정코드를 입력하세요 : ')
input_wm = input('WM센터 타겟홍보 과정코드를 입력하세요 : ')
input_theme = input('보험사테마과정 타겟홍보 과정코드를 입력하세요 : ')
# 입력받은 과정코드로 데이터프레임 제작
input_data = {'과정명' : ['전체', '경제톡톡', 'WM센터', '테마과정'], '과정코드' : [input_all, input_eco, input_wm, input_theme]}
df_input = pd.DataFrame(input_data)

#######################################     2023 교육관리 데이터베이스 호출     ############################################
gc = gs.service_account(filename='clever-seat-398104-4fa3ad6684b5.json')
sheets = gc.open_by_url('https://docs.google.com/spreadsheets/d/1AG89W1nwRzZxYreM6i1qmwS6APf-8GM2K_HDyX7REG4/edit#gid=0')
df_bfr_ytb = load_sheets(sheets, '[전월]유튜브') # 이번달 유튜브 교육 신청자 시트 호출
df_nxt_ytb = load_sheets(sheets, '[매일]신청현황') # 다음달 유튜브 교육 신청자 시트 호출
df_bfr_cmp = load_sheets(sheets, '[전월]사캠') # 이번달 사이버캠퍼스(경제톡톡, WM센터) 교육시청자 시트 호출

################################################     시트 정규화     ######################################################
# '[전월]유튜브' 시트 정리
df_bfr_ytb = df_bfr_ytb[(df_bfr_ytb['과정명'].str.contains('유튜브') & ~df_bfr_ytb['과정명'].str.contains('교육부') & ~df_bfr_ytb['파트너'].str.contains('인카본사'))]
df_bfr_ytb = df_bfr_ytb.drop(columns=['번호','소속부문','소속총괄','소속부서','파트너','성함','IMO신청여부','수료현황','비고'])

# '[매일]신청현황' 시트 정리
df_nxt_ytb = df_nxt_ytb[(df_nxt_ytb['과정명'].str.contains('유튜브') & ~df_nxt_ytb['과정명'].str.contains('교육부') & ~df_nxt_ytb['파트너'].str.contains('인카본사'))]
df_today = df_nxt_ytb.drop(df_nxt_ytb[df_nxt_ytb.iloc[:,11] != df_nxt_ytb.iloc[-1,11]].index) # 오늘 기준 이번달 신청자 표시
df_today = df_today.groupby(['사원번호'])['사원번호'].size().reset_index(name='신청여부')
df_today['신청여부'] = '신청'

# '[전월]사캠' 시트 정리
df_bfr_cmp['아이디'] = df_bfr_cmp['아이디'].str.replace('incar','') # 사번 정리
staff = df_bfr_cmp['아이디'].str[:3].str.contains('161') # 직원 삭제
df_bfr_cmp = df_bfr_cmp[~staff]
df_bfr_cmp['과정명'] = df_bfr_cmp.apply(change_course, axis=1) # 과정명 수정
df_bfr_cmp = df_bfr_cmp.drop(columns=['NO.','이름','사번','소속1','소속2','소속3','소속4','소속5','휴대폰','이메일','결제일','최종학습일시','유선전화','카테고리','주문코드','과정코드','과정형태','수강기간','진도율','제출일']) # 필요없는 컬럼 삭제
df_bfr_cmp = df_bfr_cmp.rename(columns={'아이디':'사원번호'})

# '[전월]유튜브' 시트와 '[전월]사캠' 시트 병합
df_before = pd.concat([df_bfr_ytb, df_bfr_cmp])
df_result = pd.DataFrame()

###############################################     데이터 처리작업     ####################################################
# 사원번호 카운트 (신청횟수)
df_counts = df_before.groupby(['사원번호'])['사원번호'].size().reset_index(name='횟수')

# 저번달 유튜브 교육 2개 이상 신청 (전체과정 안내예정)
df_multi = df_counts[df_counts['횟수'] > 1].drop(columns='횟수')
df_multi['과정명'] = '전체'
df_result = pd.concat([df_result, df_multi])

# 저번달 유튜브 교육 1개만 신청 (개별과정 안내예정)
df_single = df_counts[df_counts['횟수'] == 1]
df_course = pd.merge(df_single, df_before, on='사원번호', how='inner')
# 경제톡톡
df_eco = modify_single(df_course, '경제톡톡')
df_result = pd.concat([df_result, df_eco])
# WM센터
df_wm = modify_single(df_course, 'WM센터')
df_result = pd.concat([df_result, df_wm])
# 보테과
df_theme = modify_single(df_course, '테마과정')
df_result = pd.concat([df_result, df_theme])

# 정리한 데이터와 이번달 교육신청 데이터 병합
df_applied = pd.merge(df_result, df_today, on='사원번호', how='left')
df_applied['신청여부'] = df_applied['신청여부'].astype(str)
df_target = df_applied[~df_applied['신청여부'].str.contains('신청')]

# 입과입력 양식에 맞게 필요 컬럼 추가
df_target['손보'] = ''
df_target['생보'] = ''
df_target['승인여부'] = 2
# 컬럼 정리
df_target = pd.merge(df_target, df_input, on='과정명', how='inner')
df_target = df_target.drop(columns=['과정명','신청여부'])
df_target = df_target[['과정코드','사원번호','손보','생보','승인여부','비고']]

#############################################     엑셀 파일로 저장     ####################################################
df_target.to_excel('target.xlsx', index=False)