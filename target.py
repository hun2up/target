import gspread as gs
import pandas as pd

# 입력
input_all = input('전체과정 타겟홍보 과정코드를 입력하세요 : ')
input_eco = input('경제톡톡 타겟홍보 과정코드를 입력하세요 : ')
input_wm = input('WM센터 타겟홍보 과정코드를 입력하세요 : ')
input_theme = input('보험사테마과정 타겟홍보 과정코드를 입력하세요 : ')

input_data = {'과정명' : ['전체', '경제톡톡', 'WM센터', '테마과정'], '과정코드' : [input_all, input_eco, input_wm, input_theme]}
df_input = pd.DataFrame(input_data)

gc = gs.service_account(filename='clever-seat-398104-4fa3ad6684b5.json')
before = gc.open_by_url('https://docs.google.com/spreadsheets/d/1hV7_tGjtrVQ0QHtF-tu6wYdiWZiIUeHnoYeW1sFdHH8/edit#gid=0')
sheets_before = before.worksheet('전월').get_all_values()
df_before = pd.DataFrame(sheets_before[1:], columns=sheets_before[0])
# 과정명 컬럼에 유튜브는 포함하고, 교육부는 포함하지 않으며, 파트너 컬럼에 인카본사를 포함하지 않는 명단
youtube_before = df_before[(df_before['과정명'].str.contains('유튜브') & ~df_before['과정명'].str.contains('교육부') & ~df_before['파트너'].str.contains('인카본사'))]

already = gc.open_by_url('https://docs.google.com/spreadsheets/d/1AG89W1nwRzZxYreM6i1qmwS6APf-8GM2K_HDyX7REG4/edit#gid=216302834')
sheets_already = already.worksheet('[매일]신청현황').get_all_values()
df_already = pd.DataFrame(sheets_already[1:], columns=sheets_already[0])
# 과정명 컬럼에 유튜브는 포함하고, 교육부는 포함하지 않으며, 파트너 컬럼에 인카본사를 포함하지 않는 명단
youtube_already = df_already[(df_already['과정명'].str.contains('유튜브') & ~df_already['과정명'].str.contains('교육부') & ~df_already['파트너'].str.contains('인카본사'))]

# 이번달 신청인원
df_today = youtube_already.drop(youtube_already[youtube_already.iloc[:,11] != youtube_already.iloc[-1,11]].index).reset_index()
df_today = df_today.groupby(['사원번호'])['사원번호'].size().reset_index(name='신청여부')
df_today['신청여부'] = '신청'

df_result = pd.DataFrame()

# 사원번호 카운트 (신청횟수)
df_all = youtube_before.groupby(['사원번호'])['사원번호'].size().reset_index(name='횟수')

# 2개 이상
df_multiple = df_all[df_all['횟수'] > 1].drop(columns=['횟수'])
df_multiple['과정명'] = '전체'
df_result = pd.concat([df_result, df_multiple])

# 1개
df_single = df_all[df_all['횟수'] == 1]
df_merge = pd.merge(df_single, youtube_before, on='사원번호', how='inner')


# 경제톡톡
df_eco = df_merge[df_merge['과정명'].str.contains('경제톡톡')]
df_eco = df_eco.drop(columns=['횟수','번호','과정명','소속부문','소속총괄','소속부서','파트너','성함','IMO신청여부','수료현황','비고'])
df_eco['과정명'] = '경제톡톡'
df_result = pd.concat([df_result, df_eco])

# Wm센터
df_wm = df_merge[df_merge['과정명'].str.contains('WM센터')]
df_wm = df_wm.drop(columns=['횟수','번호','과정명','소속부문','소속총괄','소속부서','파트너','성함','IMO신청여부','수료현황','비고'])
df_wm['과정명'] = 'WM센터'
df_result = pd.concat([df_result, df_wm])

# 보험사테마과정
df_theme = df_merge[df_merge['과정명'].str.contains('테마과정')]
df_theme = df_theme.drop(columns=['횟수','번호','과정명','소속부문','소속총괄','소속부서','파트너','성함','IMO신청여부','수료현황','비고'])
df_theme['과정명'] = '테마과정'
df_result = pd.concat([df_result, df_theme])

# 결과랑 이번달 결합
df_applied = pd.merge(df_result, df_today, on='사원번호', how='left')
df_applied['신청여부'] = df_applied['신청여부'].astype(str)
df_target = df_applied[~df_applied['신청여부'].str.contains('신청')]

df_target['손보'] = ''
df_target['생보'] = ''
df_target['승인여부'] = ''
df_target['비고'] = 2

df_target = pd.merge(df_target, df_input, on='과정명', how='inner')
df_target = df_target.drop(columns=['과정명','신청여부'])
df_target = df_target[['과정코드','사원번호','손보','생보','승인여부','비고']]


print(df_target)