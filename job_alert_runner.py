"""
job_alert_runner.py
Used by BOTH Option 2 (Make.com webhook) and Option 3 (GitHub Actions)
Upload this file to your GitHub repo root.
"""
import os, json, smtplib, sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from apify_client import ApifyClient
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config from environment variables (set as GitHub Secrets) ────────────────
APIFY_TOKEN     = os.environ["APIFY_API_TOKEN"]
GMAIL_ADDRESS   = os.environ["GMAIL_ADDRESS"]
GMAIL_PASSWORD  = os.environ["GMAIL_APP_PASSWORD"]
RECIPIENT_EMAIL = os.environ.get("RECIPIENT_EMAIL", GMAIL_ADDRESS)
SEEN_FILE       = "seen_jobs.json"   # persisted via GitHub Actions cache

TIER_COLORS = {
    "🏆 EXCELLENT": ("1A7340","FFFFFF"), "✅ STRONG": ("2D6A2D","FFFFFF"),
    "🟡 GOOD": ("7F6000","FFFFFF"),     "🟠 STRETCH": ("8B3A0F","FFFFFF"),
    "❌ LOW MATCH": ("5A0000","FFFFFF"),
}

def score_job(j):
    title = j.get("title","").lower(); desc = str(j.get("descriptionText","")).lower()
    comb = title + " " + desc; score = 0; matches = []
    if "salesforce" in comb:    score+=25; matches.append("Salesforce")
    if "financial services cloud" in comb or " fsc" in comb: score+=10; matches.append("FSC")
    if "marketing cloud" in comb: score+=10; matches.append("Marketing Cloud")
    if "ncino" in comb:          score+=10; matches.append("nCino")
    if "sales cloud" in comb:    score+=7;  matches.append("Sales Cloud")
    if "soql" in comb:           score+=6;  matches.append("SOQL")
    if "crm" in comb:            score+=5;  matches.append("CRM")
    if any(w in comb for w in ["qa","quality assurance","quality engineer","tester"]):
        score+=8; matches.append("QA/Testing")
    if any(w in comb for w in ["functional testing","regression","integration testing"]):
        score+=6; matches.append("Functional/Regression")
    if "uat" in comb:            score+=4;  matches.append("UAT")
    if "postman" in comb:        score+=5;  matches.append("Postman/API")
    if "selenium" in comb:       score+=4;  matches.append("Selenium")
    if "jira" in comb:           score+=3;  matches.append("Jira")
    if "agile" in comb:          score+=3;  matches.append("Agile/Scrum")
    if any(w in comb for w in ["banking","financial","fintech","lending","wealth","insurance"]):
        score+=8; matches.append("Financial/Banking")
    if any(w in title for w in ["senior","lead","manager","principal"]):
        score+=3; matches.append("Senior/Lead")
    if score>=65:   tier="🏆 EXCELLENT"; iv="Very High (85-95%)"
    elif score>=45: tier="✅ STRONG";    iv="High (70-85%)"
    elif score>=25: tier="🟡 GOOD";      iv="Medium (50-70%)"
    elif score>=12: tier="🟠 STRETCH";   iv="Low-Medium (30-50%)"
    else:           tier="❌ LOW MATCH"; iv="Low (<30%)"
    return dict(score=min(score,100),tier=tier,interview=iv,matches=", ".join(matches) or "General QA")

def scrape_jobs():
    client = ApifyClient(APIFY_TOKEN)
    all_jobs = {}
    searches = [
        ("Salesforce+QA","Alberta%2C+Canada"), ("Salesforce+Quality+Assurance","Alberta%2C+Canada"),
        ("Salesforce+QA","Canada"), ("nCino+QA","Canada"), ("Salesforce+QA+Analyst","Canada"),
    ]
    urls = [f"https://www.linkedin.com/jobs/search/?keywords={kw}&location={loc}&f_TPR=r21600&position=1&pageNum=0"
            for kw,loc in searches]
    try:
        run = client.actor("curious_coder/linkedin-jobs-scraper").call(
            run_input={"urls": urls, "count": 100, "scrapeCompany": False})
        for item in client.dataset(run["defaultDatasetId"]).iterate_items():
            jid = str(item.get("id",""))
            if jid and jid not in all_jobs:
                item["source"] = "LinkedIn"; all_jobs[jid] = item
    except Exception as e: print(f"LinkedIn error: {e}")
    for query in ["Salesforce QA","nCino QA","Salesforce quality assurance"]:
        for loc in ["Alberta","Canada"]:
            try:
                run = client.actor("borderline/indeed-scraper").call(
                    run_input={"country":"ca","query":query,"location":loc,"maxRows":30,"sort":"date","enableUniqueJobs":True})
                for item in client.dataset(run["defaultDatasetId"]).iterate_items():
                    t = item.get("title",""); co = item.get("company","")
                    key = f"indeed_{t}_{co}"
                    if key not in all_jobs and t:
                        lc = item.get("location",{}); lc = lc.get("formattedAddressShort","") if isinstance(lc,dict) else str(lc)
                        sal = item.get("salary","") or ""; sal = sal.get("salaryText","") if isinstance(sal,dict) else str(sal)
                        all_jobs[key] = {"id":key,"title":t,"companyName":co,"location":str(lc),
                            "postedAt":item.get("date",""),"salary":str(sal),
                            "descriptionText":str(item.get("description","") or item.get("snippet","")),
                            "link":item.get("url","") or item.get("externalApplyLink",""),"source":"Indeed"}
            except Exception as e: print(f"Indeed error ({query}/{loc}): {e}")
    return list(all_jobs.values())

def get_new_jobs(jobs):
    seen = set()
    if os.path.exists(SEEN_FILE):
        with open(SEEN_FILE) as f: seen = set(json.load(f))
    new_jobs = []; new_keys = set()
    for j in jobs:
        k = f"{j.get('title','').lower().strip()}|{j.get('companyName','').lower().strip()}"
        if k not in seen: new_jobs.append(j); new_keys.add(k)
    seen.update(new_keys)
    with open(SEEN_FILE,"w") as f: json.dump(list(seen),f)
    return new_jobs

def build_excel(jobs, timestamp):
    wb = Workbook(); ws = wb.active; ws.title = "New Job Alerts"
    def tb(c):
        s=Side(style="thin",color="CCCCCC"); c.border=Border(left=s,right=s,top=s,bottom=s)
    ws.merge_cells("A1:K1")
    ws["A1"] = f"🆕 SALESFORCE QA JOB ALERTS — {timestamp} | {len(jobs)} new matches"
    ws["A1"].fill=PatternFill("solid",start_color="0D1B2A")
    ws["A1"].font=Font(bold=True,color="C9A84C",name="Arial",size=12)
    ws["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height=28
    hdrs=["#","Score","Tier","Interview %","Job Title","Company","Location","Source","Posted","Salary","Apply Link"]
    ws.row_dimensions[2].height=32
    for ci,h in enumerate(hdrs,1):
        c=ws.cell(row=2,column=ci,value=h)
        c.fill=PatternFill("solid",start_color="1E3A5F"); c.font=Font(bold=True,color="C9A84C",name="Arial",size=9)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); tb(c)
    for i,j in enumerate(jobs,1):
        a=score_job(j); tier=a["tier"]; bg,fg=TIER_COLORS.get(tier,("E0E0E0","000000"))
        rb="F5F5F5" if i%2==0 else "FFFFFF"
        vals=[i,f"{a['score']}/100",tier,a["interview"],j.get("title",""),j.get("companyName",""),
              j.get("location",""),j.get("source",""),j.get("postedAt",""),j.get("salary","") or "N/A",j.get("link","") or ""]
        for ci,val in enumerate(vals,1):
            c=ws.cell(row=2+i,column=ci,value=str(val) if val else ""); tb(c)
            c.font=Font(name="Arial",size=9); c.alignment=Alignment(vertical="center",wrap_text=True)
            if ci==1:
                c.fill=PatternFill("solid",start_color="1E3A5F"); c.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
                c.alignment=Alignment(horizontal="center",vertical="center")
            elif ci in [2,3]:
                c.fill=PatternFill("solid",start_color=bg); c.font=Font(name="Arial",size=9,bold=True,color=fg)
                c.alignment=Alignment(horizontal="center",vertical="center")
            elif ci==4:
                c.fill=PatternFill("solid",start_color=bg); c.font=Font(name="Arial",size=9,color=fg)
                c.alignment=Alignment(horizontal="center",vertical="center")
            elif ci==11:
                c.font=Font(name="Arial",size=9,color="1155CC",underline="single"); c.fill=PatternFill("solid",start_color=rb)
            else: c.fill=PatternFill("solid",start_color=rb)
        ws.row_dimensions[2+i].height=45
    for i,w in enumerate([4,9,16,18,40,26,26,10,11,16,55],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A3"
    fname=f"Salesforce_QA_Jobs_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(fname); return fname

def send_email(filepath, jobs):
    msg=MIMEMultipart(); msg["From"]=GMAIL_ADDRESS; msg["To"]=RECIPIENT_EMAIL
    msg["Subject"]=f"🆕 {len(jobs)} New Salesforce QA Jobs — {datetime.now().strftime('%b %d, %I:%M %p')}"
    rows=""
    for j in jobs[:5]:
        a=score_job(j)
        rows+=f"<tr><td style='padding:8px;border:1px solid #ddd;color:#1A7340;font-weight:bold'>{a['score']}/100</td><td style='padding:8px;border:1px solid #ddd'><b>{j.get('title','')}</b></td><td style='padding:8px;border:1px solid #ddd'>{j.get('companyName','')}</td><td style='padding:8px;border:1px solid #ddd'>{j.get('location','')}</td><td style='padding:8px;border:1px solid #ddd'><a href='{j.get('link','')}' style='color:#1155CC'>Apply →</a></td></tr>"
    html=f"""<html><body style="font-family:Arial,sans-serif">
    <div style="background:#0D1B2A;padding:20px;border-radius:8px 8px 0 0">
      <h2 style="color:#C9A84C;margin:0">🆕 {len(jobs)} New Salesforce QA Job Alerts</h2>
      <p style="color:#fff;margin:5px 0 0">{datetime.now().strftime('%A, %B %d %Y at %I:%M %p')}</p>
    </div>
    <div style="background:#f8f9fa;padding:20px;border:1px solid #dee2e6">
      <p>Hi Sukalyani! Here are your latest job matches:</p>
      <table style="width:100%;border-collapse:collapse">
        <tr style="background:#1E3A5F;color:#C9A84C">
          <th style="padding:10px;border:1px solid #ddd">Score</th><th style="padding:10px;border:1px solid #ddd">Title</th>
          <th style="padding:10px;border:1px solid #ddd">Company</th><th style="padding:10px;border:1px solid #ddd">Location</th>
          <th style="padding:10px;border:1px solid #ddd">Apply</th></tr>{rows}
      </table>
      <p style="color:#666;margin-top:15px">📎 Full Excel attached ({len(jobs)} jobs). Runs every 6 hours automatically via GitHub Actions.</p>
    </div></body></html>"""
    msg.attach(MIMEText(html,"html"))
    with open(filepath,"rb") as f:
        part=MIMEBase("application","octet-stream"); part.set_payload(f.read()); encoders.encode_base64(part)
        part.add_header("Content-Disposition",f"attachment; filename={os.path.basename(filepath)}"); msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com",465) as s:
        s.login(GMAIL_ADDRESS,GMAIL_PASSWORD); s.sendmail(GMAIL_ADDRESS,RECIPIENT_EMAIL,msg.as_string())
    print(f"✅ Email sent to {RECIPIENT_EMAIL}")

if __name__=="__main__":
    ts=datetime.now().strftime("%b %d %Y, %I:%M %p")
    print(f"🔍 Running job alert: {ts}")
    jobs=scrape_jobs()
    print(f"Total scraped: {len(jobs)}")
    new=get_new_jobs(jobs)
    print(f"New jobs: {len(new)}")
    scored=[{**j,**score_job(j)} for j in new if score_job(j)["score"]>=30]
    scored.sort(key=lambda x:-x["score"])
    print(f"High-match new jobs: {len(scored)}")
    if not scored:
        print("No new matching jobs — skipping email."); sys.exit(0)
    xl=build_excel(scored,ts)
    send_email(xl,scored)
    os.remove(xl)
    print(f"✅ Done! Sent {len(scored)} jobs.")
