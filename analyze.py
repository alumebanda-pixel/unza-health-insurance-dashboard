from pathlib import Path
import os
from openpyxl import load_workbook
from collections import Counter
import numpy as np
from scipy.stats import chi2_contingency
import json
ROOT=Path(__file__).resolve().parent
SOURCE=Path(os.environ.get('UNZA_SURVEY_FILE', ROOT/'source_responses.xlsx'))
if not SOURCE.exists():
    raise SystemExit(f'Set UNZA_SURVEY_FILE to the supplied survey responses workbook: {SOURCE}')
w=load_workbook(SOURCE,read_only=True,data_only=True)
raw=list(w.active.values); HEADERS=list(raw[0]); RAW=[list(r) for r in raw[1:] if any(v is not None for v in r)];w.close()
def norm(v):
    if isinstance(v,bool): return 'True' if v else 'False'
    if isinstance(v,str): return v.strip()
    return v
ROWS=[[norm(v) for v in row] for row in RAW]
KEYS=['True','False','False','True','True','False','True','True','False','True','True','True','False','True','True','False','False','True','False']+['Covered']*8+['Not covered']*6
CODES=['B'+str(i) for i in range(1,11)]+['C'+str(i) for i in range(2,11)]+['D'+str(i) for i in range(1,9)]+['E'+str(i) for i in range(1,7)]
SOURCES=['Radio or television','Newspapers or print media','Social media (e.g. Facebook, WhatsApp, YouTube, X/Twitter)','Family members','Friends or fellow students','School coursework or lectures','NHIMA outreach activities or NHIMA offices','Health facility staff']
SOURCE_SHORT=['Radio / TV','Print media','Social media','Family','Friends / students','Coursework / lectures','NHIMA outreach / offices','Health facility staff']
SHORT=['Meaning of health insurance','Reimbursement-only misconception','Insurance for sick people only','Monthly contributions','Salary deductions','Eligibility outside formal employment','Reduced out-of-pocket costs','Membership proof at facility','Universal facility acceptance','Registration before claims','National scheme management','Established by Act of Parliament','Formal-employment-only registration','Employer matching contributions','Membership cards','Access at non-accredited facilities','Curative-only coverage misconception','Online status checks','NHIMA as the only insurance option','Doctor consultation','Hospital admission','Basic laboratory tests','Prescribed medicines','Maternity / childbirth care','Emergency care','Chronic-condition follow-up','Rehabilitative care','Non-accredited facility treatment','Purely cosmetic procedures','Specialist care without referral','Overseas care without authorisation','Lapsed contributions','Care beyond benefit limits']
def useful(v):
    if v is None or v=='':return 'Missing'
    t=str(v).strip().lower()
    hits=[]
    rules=[('radio','Radio / TV'),('tv','Radio / TV'),('television','Radio / TV'),('newspaper','Print media'),('print media','Print media'),('social','Social media'),('socil','Social media'),('family','Family'),('friend','Friends / students'),('school','Coursework / lectures'),('lecture','Coursework / lectures'),('nhima','NHIMA outreach / offices'),('health facility','Health facility staff')]
    for k,label in rules:
        if k in t and label not in hits:hits.append(label)
    if len(hits)>1:return 'Multiple sources'
    if len(hits)==1:return hits[0]
    if t in ['non','none'] or 'not received' in t:return 'No information'
    return 'Other / unclear'
RECORDS=[]
for row in ROWS:
    sc=[None if row[6+i] is None or row[6+i]=='' else int(row[6+i]==k) for i,k in enumerate(KEYS)]
    bs=sum(x or 0 for x in sc[:10]);cs=sum(x or 0 for x in sc[10:19]);total=bs+cs; complete=all(x is not None for x in sc[:19])
    RECORDS.append(dict(raw=row,sc=sc,b=bs,c=cs if complete else None,total=total,adequate=int(total>=12),complete=int(complete),d=sum(x or 0 for x in sc[19:27]),e=sum(x or 0 for x in sc[27:]),sources=[int(s.lower() in str(row[39]).lower()) for s in SOURCES],useful=useful(row[40])))
a=np.array([r['total'] for r in RECORDS]);cv=[r['c'] for r in RECORDS if r['c'] is not None]
RESULTS={'n':len(ROWS),'b_mean':np.mean([r['b'] for r in RECORDS]),'b_adequate':sum(r['b']>=6 for r in RECORDS),'c_valid':len(cv),'c_mean':np.mean(cv),'c_adequate':sum(v>=6 for v in cv),'overall_mean':a.mean(),'overall_sd':a.std(ddof=1),'overall_median':float(np.median(a)),'overall_adequate':int(sum(a>=12)),'de_mean':np.mean([r['d']+r['e'] for r in RECORDS]),'sources':dict(zip(SOURCE_SHORT,np.sum([r['sources'] for r in RECORDS],axis=0).tolist())),'tests':{}}
for label,idx in [('Sex',2),('Year of study',3),('Programme',4),('Insurance registration',5)]:
    groups=sorted(set(row[idx] for row in ROWS));table=np.array([[sum(r['raw'][idx]==g and r['adequate']==val for r in RECORDS) for val in [0,1]] for g in groups]);stat,p,df,expected=chi2_contingency(table,correction=False)
    RESULTS['tests'][label]={'chi2':stat,'p':p,'df':int(df),'min_expected':expected.min(),'cells_under5':int((expected<5).sum()),'cells_total':int(expected.size),'groups':groups,'table':table.tolist()}
if __name__=='__main__':
    print(json.dumps(RESULTS,indent=2)); (ROOT/'analysis_checks.json').write_text(json.dumps(RESULTS,indent=2)+'\n')
