// Focused regression for keep-together: build a synthetic statements object with
// MANY multi-line opex categories so Operations spans >1 page, then assert no
// category group is split (header and its "Total" land on the same page).
const fin = require('../server/financials.js');
const pdfParse = require('pdf-parse');

// Two-line categories chosen so their headers naturally fall near page bottoms.
// Each maps to real PL_EXPENSE_MAP codes so plExpenseCategory groups them.
const CATS = [
  { title: 'Professional Services', codes: [['63000','Accounting'],['63025','Professional Fees']] },
  { title: 'Administrative & Other', codes: [['60210','Travel'],['60500','Meals'],['67100','Dues & Subscriptions'],['67200','Office Expense'],['67400','Advertising & Marketing']] },
  { title: 'Personnel / Payroll', codes: [['60000','Salaries & Wages'],['60002','Payroll Taxes'],['60005','Health Insurance'],['60012','RRB Taxes'],['63042','Offsite Staff']] },
  { title: 'Equipment & Rolling Stock', codes: [['61000','Locomotive Rent'],['61005','Vehicle Rent'],['61053','Equipment Supplies'],['61054','Locomotive Repair']] },
  { title: 'Fuel & Utilities', codes: [['61150','Utilities'],['61152','Water'],['61164','Fuel']] },
  { title: 'Contracted Services', codes: [['61056','Landscape Maintenance'],['61064','Pest Control Services']] },
  { title: 'Insurance', codes: [['65000','Insurance - Liability'],['68055','Property Insurance']] },
  { title: 'Taxes & Assessments', codes: [['68000','Tax & License'],['68050','Property Tax'],['68060','State and Local Taxes']] },
];

// Build opex lines + opexGroups exactly as buildStatements would.
let opex = [], opexGroups = [];
for (const c of CATS) {
  const lines = c.codes.map(([code,name],i)=>({ code, name, cur: 100+i, pri: 50+i, ytd: 200+i }));
  opex = opex.concat(lines);
  const sum = k => lines.reduce((a,l)=>a+l[k],0);
  opexGroups.push({ title: c.title, lines, subtotal: { cur: sum('cur'), pri: sum('pri'), ytd: sum('ytd') } });
}
// Pad with extra synthetic multi-line groups to force a page overflow.
for (let p=0;p<6;p++){
  const lines=[];
  for(let j=0;j<6;j++) lines.push({code:'9'+p+j+'00',name:'Pad Expense '+p+'-'+j+' (professional fee)',cur:10+j,pri:5+j,ytd:20+j});
  opex=opex.concat(lines);
  const sum=k=>lines.reduce((a,l)=>a+l[k],0);
  opexGroups.push({title:'Pad Group '+p,lines,subtotal:{cur:sum('cur'),pri:sum('pri'),ytd:sum('ytd')}});
}
const sumAll = k => opex.reduce((a,l)=>a+l[k],0);
const totOpex = { cur: sumAll('cur'), pri: sumAll('pri'), ytd: sumAll('ytd') };

const s = {
  meta: { entityName: 'County Line SRN', asOf: '2026-04-30', priorDate: '2026-03-31',
          longDate: 'April 30, 2026', priorLongDate: 'March 31, 2026',
          monthsEnded: 'For the Four Months Ended April 30, 2026', period: 'monthly',
          periodLabel: 'For the Month Ended April 30, 2026', colLabel: 'Month Ended' },
  balanceSheet: { assetSections: [], liabSections: [], equityRows: [], retainedRows: [],
    totalAssets:{cur:0,pri:0}, totalLiab:{cur:0,pri:0}, totalContribEquity:{cur:0,pri:0},
    niLine:{cur:0,pri:0}, totalEquity:{cur:0,pri:0}, totalLiabEquity:{cur:0,pri:0} },
  operations: { revenue: [{code:'40000',name:'Revenue',cur:500,pri:400,ytd:2000}],
    cogs: [], opex, opexGroups,
    totRev:{cur:500,pri:400,ytd:2000}, totCogs:{cur:0,pri:0,ytd:0},
    grossProfit:{cur:500,pri:400,ytd:2000}, totOpex,
    netIncome:{cur:500-totOpex.cur,pri:400-totOpex.pri,ytd:2000-totOpex.ytd} },
  cashFlow: { netIncome:0, amortization:0, changeAR:0, changePrepaidOther:0, changeAP:0,
    changeAccrued:0, changeIntercompany:0, capex:0, ltInvest:0, equityContrib:0, debtChange:0,
    cashBeg:0, cashEnd:0, netOperating:0, netInvesting:0, netFinancing:0, netChange:0,
    actualCashChange:0, tieOut:0 },
  equity: { rows: [], totals: { beginning:0, contributions:0, distributions:0, netIncome:0, ending:0 } },
  checks: { balanceSheetTies:true, balanceSheetDiff:0, cashFlowTies:true, cashFlowDiff:0, niAgrees:true },
};

(async () => {
  const bytes = await fin.renderStatementsPdf(s, []);
  const pages = [];
  await pdfParse(Buffer.from(bytes), { pagerender: pd => pd.getTextContent().then(tc => { const t = tc.items.map(i=>i.str).join(' '); pages.push(t); return t; }) });

  // Find the Operations pages.
  const opsPages = pages.map((t,i)=>({i,t})).filter(p=>/Operating Expenses|Total Operating Expenses|Total Contracted Services/.test(p.t));
  const multiPage = pages.filter(t=>/Statements of Operations/.test(t)).length;
  console.log('Operations spans '+multiPage+' page(s); total pages '+pages.length);

  let problems = 0, checked = 0;
  for (const g of opexGroups) {
    if (g.lines.length < 2) continue;
    const totalLabel = 'Total ' + g.title;
    const totalPage = pages.findIndex(t => t.includes(totalLabel));
    // header page = first page containing the title as a non-"Total" occurrence
    let headerPage = -1;
    for (let i=0;i<pages.length;i++){
      const t = pages[i];
      const stripped = t.split(totalLabel).join('');
      if (stripped.includes(g.title)) { headerPage = i; break; }
    }
    checked++;
    if (headerPage !== totalPage || headerPage === -1) {
      console.log('SPLIT: '+g.title+'  header p'+(headerPage+1)+'  total p'+(totalPage+1));
      problems++;
    }
  }
  if (multiPage < 2) console.log('WARN: Operations did not span multiple pages — test not exercising a break.');
  console.log(problems===0 ? ('PASS — '+checked+' groups checked, none split across pages') : (problems+' SPLIT group(s)'));
  process.exit((problems===0 && multiPage>=2) ? 0 : 1);
})();
