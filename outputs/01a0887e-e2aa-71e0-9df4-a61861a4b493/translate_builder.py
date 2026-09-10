import json
from pathlib import Path
p=Path(__file__).resolve().parent
labels=json.loads((p/'english_labels.json').read_text(encoding='utf-8'))
s=(p/'build.mjs').read_text(encoding='utf-8')
s=s.replace('Ticket_Labour_Invoice_月度核对.xlsx','Ticket_Labour_Invoice_Monthly_Reconciliation_EN.xlsx')
for k in sorted(labels,key=len,reverse=True):
    s=s.replace(k,labels[k])
s=s.replace('const wb=Workbook.create();',"""const labels=JSON.parse(await fs.readFile(path.join(dir,'english_labels.json'),'utf8'));
for(const rows of [d.tickets,d.confirmations,d.billing,d.expected])for(const row of rows)for(let i=0;i<row.length;i++)if(typeof row[i]==='string' && labels[row[i]])row[i]=labels[row[i]];
const wb=Workbook.create();""")
s=s.replace('q.format.rowHeight=32','q.format.rowHeight=36')
s=s.replace("name+'.png'","name+'_en.png'")
s=s.replace("'validation.json'","'validation_en.json'")
(p/'build.mjs').write_text(s,encoding='utf-8')
print('English builder saved')
