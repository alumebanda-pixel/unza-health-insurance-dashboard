from analyze import *
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col
from scipy.stats import chi2_contingency
from datetime import datetime

OUT=ROOT/'UNZA_Health_Insurance_Dashboard.xlsx'
wb=xlsxwriter.Workbook(OUT,{'use_future_functions':True})
wb.set_properties({'title':'UNZA Health Insurance & NHIMA Dashboard','subject':'Ridgeway Campus student survey, 358 responses','author':'UNZA survey analysis','comments':'Based on the supplied response workbook and analysis report. Provisional instrument-derived scoring key.'})
wb.set_calc_mode('auto')
TEAL='#1A5F5F'; DARK='#163D42'; LIGHT='#D3E5E5'; BG='#F4F8F8'; GRAY='#647578'; GOLD='#D49A35'; WHITE='#FFFFFF'; LINE='#D9E3E3'
fmt={}
def f(name,**kwargs):
    fmt[name]=wb.add_format({'font_name':'Calibri','font_size':11,'font_color':DARK,'valign':'vcenter',**kwargs});return fmt[name]
f('title',font_name='Georgia',font_size=22,bold=True,font_color=WHITE,bg_color=TEAL)
f('subtitle',font_size=10,font_color=GRAY)
f('section',font_name='Georgia',font_size=13,bold=True,font_color=TEAL,bg_color=LIGHT)
f('header',font_name='Georgia',font_size=10,bold=True,font_color=WHITE,bg_color=TEAL,text_wrap=True,bottom=2,bottom_color=TEAL,align='center')
f('text',text_wrap=True,bottom=1,bottom_color=LINE)
f('center',align='center',bottom=1,bottom_color=LINE)
f('int',num_format='#,##0',align='right',bottom=1,bottom_color=LINE)
f('dec',num_format='0.0',align='right',bottom=1,bottom_color=LINE)
f('pct',num_format='0.0%',align='right',bottom=1,bottom_color=LINE)
f('p',num_format='0.0000',align='right',bottom=1,bottom_color=LINE)
f('note',font_size=10,font_color=GRAY,text_wrap=True)
f('warning',font_size=10,font_color='#815F22',bg_color='#FFF3DA',text_wrap=True)
f('input',font_size=11,bold=True,font_color=TEAL,bg_color='#FFFDE7',text_wrap=True,align='center',border=1,border_color='#B6C8C8')
f('label',font_size=10,bold=True,font_color=TEAL)
f('klabel',font_size=10,bold=True,bg_color=LIGHT,align='center')
f('knum',font_size=28,bold=True,bg_color=WHITE,align='center',num_format='#,##0')
f('kpct',font_size=28,bold=True,bg_color=WHITE,align='center',num_format='0.0%')
f('kdec',font_size=28,bold=True,bg_color=WHITE,align='center',num_format='0.0')
f('kfoot',font_size=9,font_color=GRAY,bg_color=WHITE,align='center')
f('link',font_size=10,bold=True,font_color=TEAL,underline=1)
f('date',num_format='dd mmm yyyy hh:mm',align='center',bottom=1,bottom_color=LINE)
f('key',font_size=11,bold=True,bg_color='#FFFDE7',align='center',border=1,border_color=LINE)

names=['Overview','Cohort analysis','Item results','Associations','Responses','Codebook','Notes']
sheets={name:wb.add_worksheet(name) for name in names}
for ws in sheets.values():
    ws.hide_gridlines(2); ws.set_tab_color(TEAL); ws.set_column('A:A',3); ws.set_default_row(20)
    ws.set_landscape();ws.set_paper(9);ws.fit_to_pages(1,0);ws.set_margins(.25,.25,.35,.35)
    ws.set_footer('&LUNZA | Ridgeway Campus&CHealth insurance survey&RPage &P of &N',{'margin':.15})
    if ws.name!='Overview':ws.write_url('B1',"internal:'Overview'!B2",fmt['link'],'Back to dashboard')

def merge(ws,r,c1,c2,text,style='note',height=None):
    ws.merge_range(r,c1,r,c2,text,fmt[style])
    if height:ws.set_row(r,height)
def heading(ws,title,subtitle,last=8):
    merge(ws,1,1,last,title,'title',36);merge(ws,2,1,last,subtitle,'subtitle',28)
def section(ws,row,title,last=8):merge(ws,row,1,last,title,'section',27)
def header(ws,row,labels,start=1):
    ws.set_row(row,32)
    for i,label in enumerate(labels):ws.write(row,start+i,label,fmt['header'])
def formula(ws,r,c,expr,cache,style='int'):
    ws.write_formula(r,c,expr,fmt[style],cache)

# -------------------- Editable codebook --------------------
cb=sheets['Codebook'];cb.set_column('B:B',10);cb.set_column('C:C',23);cb.set_column('D:D',72);cb.set_column('E:E',43);cb.set_column('F:F',16)
heading(cb,'Scoring key & question dictionary','Yellow cells are editable. The answer key is instrument-derived and must be validated against the study protocol.',5)
header(cb,4,['Scale','Threshold','Maximum','Missing-response treatment'],1)
thresholds=[('Health insurance',6,10,'All 10 B items answered in supplied sample.'),('NHIMA',6,9,'Score and adequacy exclude incomplete C records.'),('Overall',12,19,'Missing items contribute zero to overall score; all respondents retained to reproduce the report.')]
for i,(label,t,m,note) in enumerate(thresholds,5):
    cb.write(i,1,label,fmt['text']);cb.write(i,2,t,fmt['key']);cb.write(i,3,m,fmt['int']);merge(cb,i,4,5,note,'note',38)
    cb.data_validation(i,2,i,2,{'validate':'integer','criteria':'between','minimum':0,'maximum':m})
merge(cb,9,1,5,'Correct = 1; incorrect / Not sure = 0; blank = missing. Benefits (D) and limitations (E) are reported separately from the 19-item overall score.','warning',37)
header(cb,11,['Item','Domain','Full survey question','Chart label','Correct answer'])
for j,(code,key,short) in enumerate(zip(CODES,KEYS,SHORT)):
    domain={'B':'Health insurance','C':'NHIMA awareness','D':'Covered benefits','E':'Limitations'}[code[0]]
    for c,v in enumerate([code,domain,HEADERS[6+j],short,key],1):cb.write(12+j,c,v,fmt['key'] if c==5 else fmt['text'])
    cb.set_row(12+j,52 if len(HEADERS[6+j])>110 else 40)
    cb.data_validation(12+j,5,12+j,5,{'validate':'list','source':['True','False'] if j<19 else ['Covered','Not covered']})
cb.autofilter(11,1,44,5);cb.freeze_panes(12,3);cb.print_area(1,1,44,5);cb.repeat_rows(11)

# -------------------- Dashboard controls --------------------
ov=sheets['Overview'];ov.set_column('B:Q',9);ov.set_column('R:R',3);ov.set_zoom(80)
for r in range(58):ov.set_row(r,18)
heading(ov,'UNZA | Health Insurance & NHIMA','Ridgeway Campus student survey  •  21 Aug – 29 Sep 2026  •  358 original responses  •  Report dated 02 Oct 2026',16)
filter_specs=[(1,'Sex',['All','Female','Male']),(5,'Year of study',['All','2nd year','3rd year','4th year','5th year or above']),(9,'Programme',['All','Medicine','Nursing Sciences','Other Ridgeway Health Sciences programme','Pharmacy','Public Health']),(13,'Insurance registration',['All','Yes','No','Not sure'])]
for c,label,options in filter_specs:
    merge(ov,4,c,c+3,label,'label');merge(ov,5,c,c+3,'All','input')
    ov.data_validation(5,c,5,c+3,{'validate':'list','source':options,'input_title':label,'input_message':'Choose All to reset this filter.','error_title':'Use a listed category','error_message':'Please select a category from the list.'})
ov.set_row(5,36)
merge(ov,6,1,16,'Change the four yellow dropdowns to explore a cohort. All cards, dashboard charts and detailed summaries recalculate.','note',20)

# -------------------- Responses / calculations --------------------
ds=sheets['Responses'];ds.set_zoom(85)
display_headers=['ID','Timestamp','Age as reported','Sex','Year of study','Programme','Insurance registered','Health score /10','NHIMA score /9','Overall score /19','Adequate overall (1/0)','B/C complete (1/0)','Benefits score /8','Limitations score /6','Benefits + limits /14','Wants more information','Most useful source (normalised)','Selected cohort (1/0)']+CODES+['Information sources (multi-select)','Most useful source (original)']+SOURCE_SHORT+['No information source']+['Score '+code for code in CODES]
# Zero-based columns: input items R:S... begin at 18; flags begin 53; helpers 62.
ITEM_START=18;F2COL=51;F3COL=52;FLAG_START=53;SCORE_START=62
assert len(display_headers)==95
heading(ds,'Survey responses & derived scores','Original file preserved. Text trimmed; Boolean true/false cells normalised to text. Filter this table to inspect individual records.',16)
header(ds,4,display_headers)
ds.set_column('B:B',7);ds.set_column('C:C',21);ds.set_column('D:D',16);ds.set_column('E:F',18);ds.set_column('G:G',30);ds.set_column('H:H',16);ds.set_column('I:P',13);ds.set_column('Q:R',24);ds.set_column('S:S',13,None,{'hidden':True});ds.set_column(18,50,13);ds.set_column(51,52,65);ds.set_column(53,94,13,None,{'hidden':True})
# ID starts in column B; internal indices shifted +1 when written.
START_ROW=5;END_ROW=START_ROW+len(RECORDS)-1
# Column helper for zero-based display_header index.
def dc(i):return col(i+1)
def rr(i):return f"'Responses'!${dc(i)}${START_ROW+1}:${dc(i)}${END_ROW+1}"
sel=rr(17)
for n,rec in enumerate(RECORDS):
    row=START_ROW+n;er=row+1;raw=rec['raw']
    vals=[n+1,raw[0],raw[1],raw[2],raw[3],raw[4],raw[5],None,None,None,None,None,None,None,None,raw[41],rec['useful'],None]+raw[6:39]+[raw[39],raw[40]]+rec['sources']+[int('not received information' in str(raw[39]).lower())]+[None]*33
    for i,value in enumerate(vals):
        c=i+1
        if value is None:ds.write_blank(row,c,None,fmt['center'])
        elif isinstance(value,datetime):ds.write_datetime(row,c,value,fmt['date'])
        elif isinstance(value,(int,float)):ds.write_number(row,c,value,fmt['int'])
        else:ds.write_string(row,c,str(value),fmt['text'] if i in [5,16,51,52] else fmt['center'])
    for j,key in enumerate(KEYS):
        ic=f'{dc(ITEM_START+j)}{er}';kr=f"'Codebook'!$F${13+j}"
        cache='' if rec['sc'][j] is None else rec['sc'][j]
        formula(ds,row,SCORE_START+j+1,f'=IF({ic}="","",--({ic}={kr}))',cache)
    bstart=f'{dc(SCORE_START)}{er}';bend=f'{dc(SCORE_START+9)}{er}';cstart=f'{dc(SCORE_START+10)}{er}';cend=f'{dc(SCORE_START+18)}{er}';dstart=f'{dc(SCORE_START+19)}{er}';dend=f'{dc(SCORE_START+26)}{er}';estart=f'{dc(SCORE_START+27)}{er}';eend=f'{dc(SCORE_START+32)}{er}'
    formula(ds,row,8,f'=SUM({bstart}:{bend})',rec['b'])
    formula(ds,row,9,f'=IF(COUNT({cstart}:{cend})=9,SUM({cstart}:{cend}),"")',rec['c'] if rec['c'] is not None else '')
    formula(ds,row,10,f'=SUM({bstart}:{cend})',rec['total'])
    formula(ds,row,11,f'=--({dc(9)}{er}>=\'Codebook\'!$C$8)',rec['adequate'])
    formula(ds,row,12,f'=--(COUNT({bstart}:{cend})=19)',rec['complete'])
    formula(ds,row,13,f'=SUM({dstart}:{dend})',rec['d'])
    formula(ds,row,14,f'=SUM({estart}:{eend})',rec['e'])
    formula(ds,row,15,f'={dc(12)}{er}+{dc(13)}{er}',rec['d']+rec['e'])
    terms=[]
    for rawidx,filtercell in [(3,'B6'),(4,'F6'),(5,'J6'),(6,'N6')]:
        terms.append(f'--OR(Overview!${filtercell[0]}${filtercell[1:] }="All",{dc(rawidx)}{er}=Overview!${filtercell[0]}${filtercell[1:]})')
    formula(ds,row,18,'='+'*'.join(terms),1)
    for j,source in enumerate(SOURCES):
        # Match complete option text, not a comma split (social-media examples include commas).
        formula(ds,row,FLAG_START+j+1,f'=--ISNUMBER(SEARCH("{source}",{dc(F2COL)}{er}))',rec['sources'][j])
    formula(ds,row,FLAG_START+8+1,f'=--ISNUMBER(SEARCH("not received information",{dc(F2COL)}{er}))',vals[61])
    ds.set_row(row,24)
ds.add_table(4,1,END_ROW,len(display_headers),{'name':'SurveyResponses','style':'Table Style Medium 2','columns':[{'header':h} for h in display_headers],'autofilter':True})
ds.freeze_panes(5,8);ds.repeat_rows(4);ds.print_area(1,1,END_ROW,16)
ds.write_comment(4,10,'Overall score: B1–B10 + C2–C10. The two missing C blocks contribute zero, matching 100/358 and mean 9.7/19 in the report.')
ds.write_comment(4,9,'NHIMA score is blank unless all nine C items were answered. Valid NHIMA denominator = 356 in the full sample.')

# -------------------- Cohort summary engine --------------------
co=sheets['Cohort analysis'];co.set_column('B:B',42);co.set_column('C:C',16);co.set_column('D:D',17);co.set_column('E:H',20)
heading(co,'Cohort analysis','All results on this sheet follow the four dashboard dropdowns. Percentages use explicit valid-response denominators.',7)
header(co,4,['Metric','Value','Denominator','Definition'],1)
metric_defs=[
('Selected respondents',f'=SUM({sel})',len(RECORDS),'int',len(RECORDS),'Filter-matched survey records'),
('Adequate overall',f'=SUMIFS({rr(10)},{sel},1)',RESULTS['overall_adequate'],'int',len(RECORDS),'Overall score meets editable threshold (default 12/19)'),
('Overall adequacy rate','=IFERROR(C7/C6,0)',RESULTS['overall_adequate']/len(RECORDS),'pct',len(RECORDS),'All selected respondents, matching report convention'),
('Mean overall score',f'=IFERROR(SUMIFS({rr(9)},{sel},1)/C6,0)',RESULTS['overall_mean'],'dec',len(RECORDS),'Score out of 19; missing C items contribute zero'),
('Adequate health insurance',f'=COUNTIFS({rr(7)},">="&Codebook!$C$6,{sel},1)',RESULTS['b_adequate'],'int',len(RECORDS),'B1–B10 threshold (default 6/10)'),
('Health-insurance adequacy rate','=IFERROR(C10/C6,0)',RESULTS['b_adequate']/len(RECORDS),'pct',len(RECORDS),'All selected respondents'),
('Mean health-insurance score',f'=IFERROR(SUMIFS({rr(7)},{sel},1)/C6,0)',RESULTS['b_mean'],'dec',len(RECORDS),'Score out of 10'),
('Valid NHIMA scores',f'=SUMIFS({rr(11)},{sel},1)',RESULTS['c_valid'],'int',len(RECORDS),'Complete B/C records (B complete in source)'),
('Adequate NHIMA',f'=COUNTIFS({rr(8)},">="&Codebook!$C$7,{sel},1,{rr(11)},1)',RESULTS['c_adequate'],'int',RESULTS['c_valid'],'C2–C10 threshold (default 6/9)'),
('NHIMA adequacy rate','=IFERROR(C14/C13,0)',RESULTS['c_adequate']/RESULTS['c_valid'],'pct',RESULTS['c_valid'],'Only respondents with complete NHIMA blocks'),
('Mean NHIMA score',f'=IFERROR(SUMIFS({rr(8)},{sel},1)/C13,0)',RESULTS['c_mean'],'dec',RESULTS['c_valid'],'Score out of 9; missing blocks excluded'),
('Mean benefits + limitations',f'=IFERROR(SUMIFS({rr(14)},{sel},1)/C6,0)',RESULTS['de_mean'],'dec',len(RECORDS),'Score out of 14; not part of overall adequacy'),
('Registered with insurance',f'=COUNTIFS({rr(6)},"Yes",{sel},1)',138,'int',len(RECORDS),'Any scheme, not necessarily NHIMA'),
('Insurance registration rate','=IFERROR(C18/C6,0)',138/len(RECORDS),'pct',len(RECORDS),'Yes / all selected respondents'),
('Social media selected',f'=SUMIFS({rr(FLAG_START+2)},{sel},1)',252,'int',len(RECORDS),'Multi-select F2 source question'),
('Social media reach','=IFERROR(C20/C6,0)',252/len(RECORDS),'pct',len(RECORDS),'May overlap with other information sources'),
('Wants more information',f'=COUNTIFS({rr(15)},"Yes",{sel},1)',158,'int',354,'F4 response = Yes'),
('Valid information-interest responses',f'=COUNTIFS({sel},1,{rr(15)},"<>")',354,'int',len(RECORDS),'Excludes four missing F4 responses'),
('Desire for more information rate','=IFERROR(C22/C23,0)',158/354,'pct',354,'Yes / nonmissing F4 responses'),
('Incomplete NHIMA blocks','=C6-C13',2,'int',len(RECORDS),'Missing C2–C10 responses'),
]
METRIC_ROWS={}
for i,(label,expr,cache,style,den,note) in enumerate(metric_defs,5):
    METRIC_ROWS[label]=i+1;co.write(i,1,label,fmt['text']);formula(co,i,2,expr,cache,style)
    denformula='=C13' if i in [13,14,15] else '=C23' if i in [21,23] else '=C6'
    formula(co,i,3,denformula,den)
    merge(co,i,4,7,note,'note',25)
co.freeze_panes(5,2)
extra_metrics=[
    ('Overall score SD',f'=IFERROR(SQRT(MAX(0,(SUMPRODUCT({sel},{rr(9)},{rr(9)})-C6*C9^2)/(C6-1))),0)',RESULTS['overall_sd'],'Sample standard deviation; zero shown if fewer than two respondents.'),
    ('Overall score median',None,RESULTS['overall_median'],'Median of selected overall scores.'),
    ('Mean covered-benefit score',f'=IFERROR(SUMIFS({rr(12)},{sel},1)/C6,0)',float(np.mean([rec['d'] for rec in RECORDS])),'D1–D8 score out of 8.'),
    ('Mean limitations score',f'=IFERROR(SUMIFS({rr(13)},{sel},1)/C6,0)',float(np.mean([rec['e'] for rec in RECORDS])),'E1–E6 score out of 6.')
]
for r,(label,expr,cache,note) in enumerate(extra_metrics,26):
    co.write(r,1,label,fmt['text'])
    if expr:formula(co,r,2,expr,cache,'dec')
    else:co.write_array_formula(r,2,r,2,f'=IFERROR(MEDIAN(IF({sel}=1,{rr(9)})),0)',fmt['dec'],cache)
    formula(co,r,3,'=C6',358);merge(co,r,4,7,note,'note',25)
section(co,31,'Domain adequacy',5);header(co,32,['Domain','Adequate','Valid n','Adequate %'])
domainrows=[('Health insurance',RESULTS['b_adequate'],358,RESULTS['b_adequate']/358,'C10','C6'),('NHIMA awareness',74,356,74/356,'C14','C13'),('Overall knowledge / awareness',100,358,100/358,'C7','C6')]
for i,(label,n,d,p,nref,dref) in enumerate(domainrows,33):
    co.write(i,1,label,fmt['text']);formula(co,i,2,f'={nref}',n);formula(co,i,3,f'={dref}',d);formula(co,i,4,f'=IFERROR(C{i+1}/D{i+1},0)',p,'pct')
section(co,39,'Adequacy by programme',5);header(co,40,['Programme','Selected n','Adequate n','Adequate %'])
PROGRAMMES=['Public Health','Medicine','Nursing Sciences','Pharmacy','Other Ridgeway Health Sciences programme']
for i,g in enumerate(PROGRAMMES,41):
    sub=[r for r in RECORDS if r['raw'][4]==g];n=len(sub);ad=sum(r['adequate'] for r in sub)
    co.write(i,1,g,fmt['text']);formula(co,i,2,f'=COUNTIFS({rr(5)},B{i+1},{sel},1)',n);formula(co,i,3,f'=SUMIFS({rr(10)},{rr(5)},B{i+1},{sel},1)',ad);formula(co,i,4,f'=IFERROR(D{i+1}/C{i+1},0)',ad/n,'pct');co.set_row(i,30)
section(co,49,'Information sources | multi-select',5);header(co,50,['Source','Selected n','Respondents','Reach %'])
sourceorder=[2,0,3,4,5,1,6,7]
for i,j in enumerate(sourceorder,51):
    n=RESULTS['sources'][SOURCE_SHORT[j]];co.write(i,1,SOURCE_SHORT[j],fmt['text']);formula(co,i,2,f'=SUMIFS({rr(FLAG_START+j)},{sel},1)',n);formula(co,i,3,'=C6',358);formula(co,i,4,f'=IFERROR(C{i+1}/D{i+1},0)',n/358,'pct')
merge(co,60,1,7,'Multi-select sources are counted once per respondent per option; percentages can add to more than 100%. Social-media option contains commas and was matched as a complete label, not split on commas.','note',38)
section(co,63,'Cohort breakdowns',5)
r=64
for label,idx,groups in [('Sex',3,['Female','Male']),('Year',4,['2nd year','3rd year','4th year','5th year or above']),('Insurance registered',6,['Yes','No','Not sure'])]:
    header(co,r,[label,'Selected n','Adequate n','Adequate %']);r+=1
    rawidx={3:2,4:3,6:5}[idx]
    for g in groups:
        sub=[rec for rec in RECORDS if rec['raw'][rawidx]==g];n=len(sub);ad=sum(rec['adequate'] for rec in sub)
        co.write(r,1,g,fmt['text']);formula(co,r,2,f'=COUNTIFS({rr(idx)},B{r+1},{sel},1)',n);formula(co,r,3,f'=SUMIFS({rr(10)},{rr(idx)},B{r+1},{sel},1)',ad);formula(co,r,4,f'=IFERROR(D{r+1}/C{r+1},0)',ad/n,'pct');r+=1
    r+=2
co.conditional_format('E34:E36',{'type':'data_bar','bar_color':TEAL,'min_type':'num','min_value':0,'max_type':'num','max_value':1})
co.conditional_format('E42:E46',{'type':'data_bar','bar_color':TEAL,'min_type':'num','min_value':0,'max_type':'num','max_value':1})
co.conditional_format('E52:E59',{'type':'data_bar','bar_color':TEAL,'min_type':'num','min_value':0,'max_type':'num','max_value':1})
co.print_area(1,1,r,7);co.repeat_rows(4)

# -------------------- Item results --------------------
it=sheets['Item results'];it.set_column('B:B',9);it.set_column('C:C',22);it.set_column('D:D',44);it.set_column('E:E',14);it.set_column('F:M',12)
heading(it,'Item-level knowledge gaps','Follows the dashboard filters. Correct-answer rates exclude missing responses; incorrect and Not sure remain in the denominator.',12)
merge(it,4,1,12,'Priority: use the lowest-scoring items to focus NHIMA orientation on benefit scope, accreditation and common insurance misconceptions. All correct-answer labels are provisional.','warning',33)
labels=['Item','Domain','Question / concept','Key','Selected n','Answered n','Correct n','Correct %','Incorrect n','Not sure n','Missing n','Not sure %'];header(it,6,labels)
for j,(code,key,short) in enumerate(zip(CODES,KEYS,SHORT)):
    r=7+j;er=r+1;domain={'B':'Health insurance','C':'NHIMA awareness','D':'Covered benefits','E':'Limitations'}[code[0]];sc=rr(SCORE_START+j);inputr=rr(ITEM_START+j)
    answered=sum(rec['sc'][j] is not None for rec in RECORDS);correct=sum(rec['sc'][j] or 0 for rec in RECORDS);ns=sum(rec['raw'][6+j]=='Not sure' for rec in RECORDS)
    it.write(r,1,code,fmt['center']);it.write(r,2,domain,fmt['text']);it.write(r,3,short,fmt['text']);formula(it,r,4,f'=Codebook!F{13+j}',key,'center')
    formula(it,r,5,"='Cohort analysis'!C6",358)
    formula(it,r,6,f'=COUNTIFS({sel},1,{inputr},"<>")',answered)
    formula(it,r,7,f'=SUMIFS({sc},{sel},1)',correct)
    formula(it,r,8,f'=IFERROR(H{er}/G{er},0)',correct/answered,'pct')
    formula(it,r,9,f'=G{er}-H{er}-K{er}',answered-correct-ns)
    formula(it,r,10,f'=COUNTIFS({inputr},"Not sure",{sel},1)',ns)
    formula(it,r,11,f'=F{er}-G{er}',358-answered)
    formula(it,r,12,f'=IFERROR(K{er}/G{er},0)',ns/answered,'pct')
    it.set_row(r,31)
it.add_table(6,1,39,12,{'name':'ItemResults','style':'Table Style Medium 2','columns':[{'header':v} for v in labels]})
it.conditional_format('I8:I40',{'type':'3_color_scale','min_type':'num','min_value':0,'min_color':'#F9DECF','mid_type':'num','mid_value':.5,'mid_color':'#FFF6DE','max_type':'num','max_value':1,'max_color':'#C8E1D6'})
it.freeze_panes(7,4);it.repeat_rows(6);it.print_area(1,1,39,12)
# Fixed baseline-priority items; chart changes rates but not membership/order.
header(it,44,['Priority item','Correct %'],1)
priority=[16,1,15,8,2]
for i,j in enumerate(priority,45):
    it.write(i,1,CODES[j]+' | '+SHORT[j],fmt['text']);correct=sum(rec['sc'][j] or 0 for rec in RECORDS);valid=sum(rec['sc'][j] is not None for rec in RECORDS)
    formula(it,i,2,f'=I{8+j}',correct/valid,'pct')

# -------------------- Full-sample association tests --------------------
asw=sheets['Associations'];asw.set_column('B:B',29);asw.set_column('C:G',15);asw.set_column('H:H',35)
heading(asw,'Factors associated with overall adequacy','Full-sample tests: independent of the dashboard filters. Formula-linked counts and p-values update if the scoring key changes.',7)
merge(asw,4,1,7,'Associations are unadjusted, not causal. Sex uses Yates continuity correction to reproduce the report; the other core tests use uncorrected Pearson chi-square.','warning',36)
header(asw,6,['Factor','Chi-square','df','p-value','Significant < .05','Min expected','Method'])
tests=[('Sex',3,['Female','Male'],True),('Year of study',4,['2nd year','3rd year','4th year','5th year or above'],False),('Programme',5,sorted(set(x['raw'][4] for x in RECORDS)),False),('Insurance registration',6,['No','Not sure','Yes'],False)]
tr=14
for summary_i,(label,idx,groups,yates) in enumerate(tests,7):
    section(asw,tr,label+' | observed and expected counts',7);tr+=1
    header(asw,tr,['Category','Inadequate','Adequate','Group total','Expected inadequate','Expected adequate','Chi-square component']);tr+=1
    first=tr;observed=[];rawidx={3:2,4:3,5:4,6:5}[idx]
    for g in groups:
        sub=[rec for rec in RECORDS if rec['raw'][rawidx]==g];ad=sum(rec['adequate'] for rec in sub);inad=len(sub)-ad;observed.append([inad,ad]);er=tr+1
        asw.write(tr,1,g,fmt['text']);formula(asw,tr,2,f'=COUNTIFS({rr(idx)},B{er},{rr(10)},0)',inad);formula(asw,tr,3,f'=COUNTIFS({rr(idx)},B{er},{rr(10)},1)',ad);formula(asw,tr,4,f'=SUM(C{er}:D{er})',len(sub))
        e0=len(sub)*(358-100)/358;e1=len(sub)*100/358
        formula(asw,tr,5,f'=E{er}*COUNTIF({rr(10)},0)/COUNTA({rr(0)})',e0,'dec');formula(asw,tr,6,f'=E{er}*COUNTIF({rr(10)},1)/COUNTA({rr(0)})',e1,'dec')
        expr=f'=IFERROR((MAX(0,ABS(C{er}-F{er})-0.5)^2/F{er})+(MAX(0,ABS(D{er}-G{er})-0.5)^2/G{er}),0)' if yates else f'=IFERROR((C{er}-F{er})^2/F{er}+(D{er}-G{er})^2/G{er},0)'
        component=(max(0,abs(inad-e0)-(.5 if yates else 0))**2/e0)+(max(0,abs(ad-e1)-(.5 if yates else 0))**2/e1)
        formula(asw,tr,7,expr,component,'dec');asw.set_row(tr,38);tr+=1
    stat,p,dof,expected=chi2_contingency(observed,correction=yates)
    asw.write(summary_i,1,label,fmt['text']);formula(asw,summary_i,2,f'=SUM(H{first+1}:H{tr})',stat,'dec');asw.write_number(summary_i,3,dof,fmt['int']);formula(asw,summary_i,4,f'=CHISQ.DIST.RT(C{summary_i+1},D{summary_i+1})',p,'p');formula(asw,summary_i,5,f'=IF(E{summary_i+1}<0.05,"Yes","No")','Yes' if p<.05 else 'No','center');formula(asw,summary_i,6,f'=MIN(F{first+1}:G{tr})',expected.min(),'dec');asw.write(summary_i,7,'Yates-corrected' if yates else 'Pearson (uncorrected)',fmt['text'])
    merge(asw,tr,1,7,'All expected cells exceed 5 in the supplied full-sample table. Review validity if you change the scoring key.','note',24);tr+=3
asw.conditional_format('E8:E11',{'type':'cell','criteria':'<','value':.05,'format':wb.add_format({'bg_color':'#D3E5E5','font_color':TEAL,'bold':True,'num_format':'0.0000'})})
asw.freeze_panes(7,2);asw.print_area(1,1,tr-1,7)

# -------------------- Notes, reconciliation and navigation --------------------
nt=sheets['Notes'];nt.set_column('B:B',28);nt.set_column('C:H',16)
heading(nt,'How to use & interpretation notes','A reproducible dashboard built from the supplied survey workbook, with the supplied report used as the reconciliation reference.',7)
notes=[
('START HERE','Open Overview in Excel and change the four yellow dropdowns. Reset all four to All for the report view. Cards, charts, Cohort analysis and Item results update together. A no-match selection displays zero values; read them as no data, not as evidence of low knowledge.'),
('NAVIGATION','Overview: executive dashboard. Cohort analysis: filter-linked counts and subgroup summaries. Item results: all 33 scored items. Associations: full-sample unadjusted tests. Responses: filterable record-level data. Codebook: question wording, provisional keys and editable thresholds.'),
('SCORING','Health insurance = B1–B10, maximum 10, adequate at 6. NHIMA = C2–C10, maximum 9, adequate at 6. Overall = B + C, maximum 19, adequate at 12. Benefits D1–D8 and limitations E1–E6 form a separate 14-item score. Incorrect and Not sure answers score zero.'),
('MISSING RESPONSES','Two respondents have all nine C items missing. Item percentages and NHIMA score statistics exclude them (valid n = 356). Overall totals treat missing C items as zero and retain all 358 respondents; this reproduces the report mean 9.7 and 100/358 adequate. This is not a complete-case overall analysis.'),
('DENOMINATORS','Full-sample health-insurance adequacy: 212/358 = 59.2%. NHIMA adequacy: 74/356 = 20.8%. Overall adequacy: 100/358 = 27.9%. Social-media source: 252/358 = 70.4%. Interest in more information: 158/354 = 44.6%; four responses missing.'),
('SOURCE QUESTIONS','F2 is multi-select. The eight named options are matched as complete strings because commas also occur within the social-media label. Source percentages may exceed 100% in total. F3 is nominally single-source but contains multiple-source and free-text responses; the original is retained and a transparent normalised grouping is provided.'),
('F3 NORMALISATION','Case-insensitive keyword matching: radio/TV/television; newspaper/print; social/socil; family; friend; school/lecture; NHIMA; health facility. Two or more matched groups = Multiple sources; non/none/not received = No information; unmatched text = Other / unclear; blank = Missing. This may differ from the report normalisation, so F3 association tests are not recreated.'),
('DEMOGRAPHIC CAVEATS','Age entries mix individual ages, ranges and multiple ranges. The original age response is retained; no midpoint or continuous-age conversion was imposed. Registration refers to any health-insurance scheme, not specifically NHIMA.'),
('STATISTICAL INTERPRETATION','Programme (p = 0.0067) and registration (p = 0.0008) show full-sample unadjusted associations with overall adequacy. Sex and year are not significant at 5%. No causal inference or generalisation to all UNZA students is warranted from this campus sample.'),
('SEX TEST RECONCILIATION','The report lists sex chi-square 0.66, p = 0.4153. Those values are reproduced by Yates-corrected chi-square (0.6637, p = 0.4153). Uncorrected Pearson gives 0.8697, p = 0.3511. The report methods call the tests Pearson; this workbook labels the correction explicitly.'),
('SUPPLEMENTARY REPORT TESTS','The report also gives age response group: chi-square 35.62, df 14, p = 0.0012; most useful source (normalised): chi-square 30.88, df 9, p = 0.0003. These are report-only findings with sparse cells and are not recreated in the core association table. The exploratory adjusted logistic model mentioned in the PDF was not supplied and is not reconstructed here.'),
('DATA QUALITY','358 records; one source worksheet. No exact duplicate records; timestamps unique. Demographics and B/D/E responses complete. Two missing responses per C item; eight missing F3 and four missing F4. Blank inputs remain blank; text is trimmed; Boolean Excel cells are normalised to True/False labels without changing meaning.'),
('POLICY VALIDATION','The answer key was inferred from the questionnaire wording and reconciles to the report. It has not been externally verified against current NHIMA rules. Confirm policy-specific items against the protocol and benefit package before thesis submission or policy use. Benefits may depend on accreditation, referrals, contribution compliance, authorisation and benefit limits.'),
('UPDATING','Yellow Codebook cells allow answer-key and threshold edits; formula-linked scores, summaries, charts and core tests recalculate in Excel. F3 normalisation is a fixed cleaned field. The workbook covers the 358 supplied records: adding new survey rows requires extending formulas and summary ranges. It contains no macros or external data links.'),
('SOURCE FILES',SOURCE.name+'; UNZA_health_insurance_analysis_report.pdf (analysis date 2 October 2026). Original attachments were not modified. Workbook created 6 October 2026. Survey timestamps: 21 August to 29 September 2026.'),
('PRACTICAL ACTIONS','Prioritise a short NHIMA orientation at Ridgeway Campus. Use social media to clarify benefit scope, accredited-facility access and exclusions. Correct specific misconceptions rather than relying only on broad awareness messages. Repeat the survey after an intervention using a validated, consistent scoring key.'),
]
row=5
for label,text in notes:
    merge(nt,row,1,7,label,'section',26);row+=1
    merge(nt,row,1,7,text,'note',62 if len(text)>320 else 48);row+=2
nt.freeze_panes(5,1);nt.print_area(1,1,row-1,7)

# -------------------- Dashboard cards and charts --------------------
def card(c,row,label,expr,cache,style,footexpr,footcache):
    merge(ov,row,c,c+3,label,'klabel',22)
    ov.merge_range(row+1,c,row+2,c+3,expr,fmt[style]);ov.write_formula(row+1,c,expr,fmt[style],cache)
    ov.merge_range(row+3,c,row+3,c+3,footexpr,fmt['kfoot']);ov.write_formula(row+3,c,footexpr,fmt['kfoot'],footcache)
card(1,8,'RESPONDENTS',"='Cohort analysis'!C6",358,'knum','="Selected cohort | total 358"','Selected cohort | total 358')
card(5,8,'ADEQUATE OVERALL',"='Cohort analysis'!C8",100/358,'kpct','="n = "&\'Cohort analysis\'!C7&" | threshold "&Codebook!C8&"/19"','n = 100 | threshold 12/19')
card(9,8,'MEAN OVERALL SCORE',"='Cohort analysis'!C9",RESULTS['overall_mean'],'kdec','="Out of 19 correct answers"','Out of 19 correct answers')
card(13,8,'ADEQUATE NHIMA',"='Cohort analysis'!C15",74/356,'kpct','="n = "&\'Cohort analysis\'!C14&" / "&\'Cohort analysis\'!C13&" valid"','n = 74 / 356 valid')
# Compact secondary statistics with meaningful denominator labels.
secondary=[('HEALTH-INSURANCE ADEQUACY',"='Cohort analysis'!C11",212/358,'pct'),('REGISTERED WITH INSURANCE',"='Cohort analysis'!C19",138/358,'pct'),('SOCIAL-MEDIA REACH',"='Cohort analysis'!C21",252/358,'pct'),('BENEFITS + LIMITATIONS',"='Cohort analysis'!C17",RESULTS['de_mean'],'dec')]
for c,(label,expr,value,style) in zip([1,5,9,13],secondary):
    merge(ov,13,c,c+3,label,'klabel',22)
    st=wb.add_format({'font_name':'Calibri','font_size':18,'bold':True,'font_color':TEAL,'align':'center','valign':'vcenter','bg_color':WHITE,'num_format':'0.0%' if style=='pct' else '0.0" /14"'})
    ov.merge_range(14,c,14,c+3,expr,st);ov.write_formula(14,c,expr,st,value);ov.set_row(14,27)

def bar_chart(title,sheet,catcol,start,end,valcol,colour=TEAL,label=True,shortlabels=None):
    chart=wb.add_chart({'type':'bar'})
    chart.add_series({'name':title,'categories':[sheet,start,catcol,end,catcol],'values':[sheet,start,valcol,end,valcol], 'fill':{'color':colour},'border':{'none':True},'data_labels':{'value':label,'num_format':'0.0%','position':'outside_end','font':{'name':'Calibri','size':9,'color':DARK}}})
    chart.set_title({'name':title,'name_font':{'name':'Georgia','size':12,'bold':True,'color':TEAL}})
    chart.set_legend({'none':True});chart.set_x_axis({'min':0,'max':1,'major_unit':.25,'num_format':'0%','num_font':{'name':'Calibri','size':9,'color':GRAY},'line':{'none':True},'major_gridlines':{'visible':True,'line':{'color':'#E8EEEE'}}})
    chart.set_y_axis({'reverse':True,'interval_unit':1,'num_font':{'name':'Calibri','size':9,'color':DARK},'line':{'none':True},'major_tick_mark':'none'})
    chart.set_chartarea({'fill':{'color':WHITE},'border':{'color':LINE}});chart.set_plotarea({'fill':{'color':WHITE},'border':{'none':True},'layout':{'x':.33,'y':.30,'width':.59,'height':.60}})
    chart.set_size({'width':497,'height':244});chart.set_style(10);chart.set_table() if False else None
    return chart
# Auxiliary short labels keep chart text readable, stored off-print on Cohort analysis.
co.set_column('K:M',18,None,{'hidden':True})
for i,(short,n,d,p,_,_) in enumerate(domainrows,33):co.write(i,10,short.replace('Overall knowledge / awareness','Overall'),fmt['text'])
for i,g in enumerate(PROGRAMMES,41):co.write(i,10,g.replace('Nursing Sciences','Nursing').replace('Other Ridgeway Health Sciences programme','Other health sciences'),fmt['text'])
for i,j in enumerate(priority,45):it.write(i,14,CODES[j]+' | '+['Curative-only coverage','Reimbursement only','Any facility access','All schemes accepted','For sick people only'][i-45],fmt['text'])
it.set_column('O:O',26,None,{'hidden':True})
# Charts include hidden category-label helpers.
chart1=bar_chart('Adequacy by knowledge domain','Cohort analysis',10,33,35,4);chart1.show_hidden_data()
chart2=bar_chart('Overall adequacy by programme','Cohort analysis',10,41,45,4);chart2.show_hidden_data()
chart3=bar_chart('Five priority misconceptions | % correct','Item results',14,45,49,2,GOLD);chart3.show_hidden_data()
chart4=bar_chart('Where students get information | reach','Cohort analysis',1,51,58,4)
for chart,anchor in [(chart1,'B17'),(chart2,'J17'),(chart3,'B31'),(chart4,'J31')]:ov.insert_chart(anchor,chart,{'x_offset':0,'y_offset':4})
merge(ov,29,1,16,'Rates use selected valid respondents. Programme rates are descriptive; full-sample tests are on Associations.','note',18)
merge(ov,43,1,16,'Priority items are the five lowest B/C items in the full sample; cohort filters update their rates, not the selected item set. F2 sources are multi-select.','note',24)
section(ov,45,'Interpretation & next steps',16)
insight='="Full sample: 27.9% overall adequate; 20.8% NHIMA adequate. Current cohort: "&TEXT(\'Cohort analysis\'!C8,"0.0%")&" overall adequate. Prioritise NHIMA orientation and targeted misconceptions."'
merge(ov,46,1,16,insight,'note',32);ov.write_formula(46,1,insight,fmt['note'],'Full sample: 27.9% overall adequate; 20.8% NHIMA adequate. Current cohort: 27.9% overall adequate. Prioritise NHIMA orientation and targeted misconceptions.')
merge(ov,48,1,16,'Provisional scoring key: validate policy-specific NHIMA answers before final use. NHIMA valid n = 356; overall n = 358 retains two incomplete C blocks with missing items scored zero.','warning',32)
for c,name in [(1,'Cohort analysis'),(5,'Item results'),(9,'Associations'),(13,'Notes')]:
    ov.merge_range(50,c,50,c+3,name,fmt['link']);ov.write_url(50,c,f"internal:'{name}'!B2",fmt['link'],name)
for c,name in [(1,'Responses'),(5,'Codebook')]:ov.write_url(52,c,f"internal:'{name}'!B2",fmt['link'],name)
merge(ov,54,1,16,'Sources: supplied response workbook + UNZA analysis report dated 02 Oct 2026. Created 06 Oct 2026. All charts and formulas are native Excel; no macros.','note',23)
ov.freeze_panes(7,1);ov.print_area(1,1,54,16);ov.fit_to_pages(1,1);ov.set_portrait();ov.set_paper(9);ov.set_header('&RUNZA • Ridgeway Campus',{'margin':.15});ov.activate();ov.set_first_sheet();ov.set_selection('B6')
wb.close()
print('Created',OUT)
