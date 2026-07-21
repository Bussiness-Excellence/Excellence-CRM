import React, { useEffect, useMemo, useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../supabaseClient';
import { useAuth } from '../contexts/AuthContext';
import './Dashboard.css';

// ── i18n ──────────────────────────────────────────────────────────────────────
const T = {
  en: {
    brand:'EXCELLENCE', signOut:'Sign out', adminPanel:'Admin Panel',
    lastMonth:'Last Month', recent:'Recent  1–15',
    allTeams:'All teams', allUsers:'All reps', search:'Search name or territory…',
    export:'Export', loading:'Loading…', noData:'No data for this period.',
    shiftAll:'Both', shiftAM:'AM', shiftPM:'PM',
    people: n=>`${n} rep${n!==1?'s':''}`,
    tabs:{ summary:'Summary', specialty:'Specialty', products:'Products', coaching:'Coaching' },
    roleView:{ MR:'My Results', Supervisor:'My Team', 'Area Manager':'My Area', BLM:'Full Team', Admin:'All Teams' },
    avg:'Avg', sum:'Sum', teamSummary:'Team Summary',
    kpiGroups:[
      { label:'Field Activity', keys:['working_days','complete_field_days','am_shift_days','pm_shift_days','double_visit_days','office_work_days'] },
      { label:'Doctor Calls',   keys:['am_calls','am_call_rate','pm_calls','pm_call_rate'] },
      { label:'Coverage',       keys:['total_am_covered','total_pm_covered','amcenter_covered','hospital_covered','clinic_covered','polyclinic_covered'] },
      { label:'Pharmacy',       keys:['pharmacies_visited','pharmacies_covered'] },
      { label:'Products',       keys:['total_product_calls','distinct_products'] },
      { label:'Coaching',       keys:['coaching_days'] },
      { label:'Timing',         keys:['avg_am_start_time','avg_am_shift_hm','avg_pm_shift_hm'] },
    ],
    kpi:{
      working_days:'Working Days', complete_field_days:'Field Days',
      am_shift_days:'AM Days', pm_shift_days:'PM Days',
      am_calls:'AM Calls', pm_calls:'PM Calls',
      am_call_rate:'AM Call Rate', pm_call_rate:'PM Call Rate',
      total_am_covered:'AM Covered', total_pm_covered:'PM Covered',
      amcenter_covered:'AM Center', hospital_covered:'Hospital',
      clinic_covered:'Clinic', polyclinic_covered:'Poly Clinic',
      double_visit_days:'Double Visits', coaching_days:'Coaching Days',
      office_work_days:'Office Work',
      pharmacies_visited:'Pharm. Visits', pharmacies_covered:'Pharm. Covered',
      total_product_calls:'Product Calls', distinct_products:'Products',
      avg_am_start_time:'AM Start Time', avg_am_shift_hm:'AM Duration', avg_pm_shift_hm:'PM Duration',
    },
  },
  ar: {
    brand:'إكسيلنس', signOut:'خروج', adminPanel:'لوحة الإدارة',
    lastMonth:'الشهر الماضي', recent:'الأحدث  1–15',
    allTeams:'كل الفرق', allUsers:'كل المندوبين', search:'بحث باسم أو منطقة…',
    export:'تصدير', loading:'جارٍ التحميل…', noData:'لا توجد بيانات.',
    shiftAll:'الكل', shiftAM:'AM', shiftPM:'PM',
    people: n=>`${n} مندوب`,
    tabs:{ summary:'الملخص', specialty:'التخصص', products:'المنتجات', coaching:'التوجيه' },
    roleView:{ MR:'نتائجي', Supervisor:'فريقي', 'Area Manager':'منطقتي', BLM:'الفريق', Admin:'الكل' },
    avg:'متوسط', sum:'مجموع', teamSummary:'ملخص الفريق',
    kpiGroups:[
      { label:'النشاط الميداني', keys:['working_days','complete_field_days','am_shift_days','pm_shift_days','double_visit_days','office_work_days'] },
      { label:'الزيارات',        keys:['am_calls','am_call_rate','pm_calls','pm_call_rate'] },
      { label:'التغطية',         keys:['total_am_covered','total_pm_covered','amcenter_covered','hospital_covered','clinic_covered','polyclinic_covered'] },
      { label:'الصيدليات',       keys:['pharmacies_visited','pharmacies_covered'] },
      { label:'المنتجات',        keys:['total_product_calls','distinct_products'] },
      { label:'التوجيه',         keys:['coaching_days'] },
      { label:'التوقيت',         keys:['avg_am_start_time','avg_am_shift_hm','avg_pm_shift_hm'] },
    ],
    kpi:{
      working_days:'أيام العمل', complete_field_days:'أيام الميدان',
      am_shift_days:'أيام AM', pm_shift_days:'أيام PM',
      am_calls:'زيارات AM', pm_calls:'زيارات PM',
      am_call_rate:'معدل AM', pm_call_rate:'معدل PM',
      total_am_covered:'تغطية AM', total_pm_covered:'تغطية PM',
      amcenter_covered:'مراكز AM', hospital_covered:'مستشفيات',
      clinic_covered:'عيادات', polyclinic_covered:'مراكز صحية',
      double_visit_days:'زيارات مزدوجة', coaching_days:'أيام التوجيه',
      office_work_days:'مكتب',
      pharmacies_visited:'زيارات صيدليات', pharmacies_covered:'تغطية صيدليات',
      total_product_calls:'مكالمات منتج', distinct_products:'منتجات',
      avg_am_start_time:'بدء AM', avg_am_shift_hm:'مدة AM', avg_pm_shift_hm:'مدة PM',
    },
  },
};



const NUMERIC_KPI_KEYS = [
  'working_days','complete_field_days','am_shift_days','pm_shift_days','double_visit_days','office_work_days',
  'am_calls','am_call_rate','pm_calls','pm_call_rate',
  'total_am_covered','total_pm_covered','amcenter_covered','hospital_covered','clinic_covered','polyclinic_covered',
  'pharmacies_visited','pharmacies_covered',
  'total_product_calls','distinct_products','coaching_days',
  'avg_am_shift_hm', 'avg_pm_shift_hm'
];

function fmtVal(v, key) {
  if (v===null||v===undefined||v==='') return '—';
  if (key?.includes('rate')) return Number(v).toFixed(1);
  if (typeof v==='number') return Number.isInteger(v)?v:Number(v).toFixed(1);
  return v;
}

function sortSummary(rows) {
  return [...rows].sort((a,b)=>{
    const tc=(a.team||'').localeCompare(b.team||'');
    if(tc) return tc;
    if(a.is_manager!==b.is_manager) return a.is_manager?1:-1;
    return (a.user_name||'').localeCompare(b.user_name||'');
  });
}

// ── Compute team/level aggregate stats ───────────────────────────────────────
function computeAggregates(rows) {
  // Only include non-manager reps for avg/sum (exclude managers from field stats)
  const reps = rows.filter(r => !r.is_manager);
  const agg = {};
  NUMERIC_KPI_KEYS.forEach(key => {
    const vals = reps.map(r => Number(r[key])||0).filter(v => v > 0);
    agg[key] = {
      sum: vals.reduce((s,v)=>s+v, 0),
      avg: vals.length ? (vals.reduce((s,v)=>s+v,0)/vals.length) : 0,
    };
  });
  return { agg, repCount: reps.length };
}

function TeamKpiHover({ rows, teamLabel }) {
  const { agg, repCount } = useMemo(() => computeAggregates(rows), [rows]);
  if (!rows.length) return null;

  const d_am_days = agg['am_shift_days']?.avg || 0;
  const d_pm_days = agg['pm_shift_days']?.avg || 0;
  const d_am_doc = agg['total_am_covered']?.avg || 0;
  const d_pm_doc = agg['total_pm_covered']?.avg || 0;
  const d_am_dur = agg['avg_am_shift_hm']?.avg || 0;
  const d_pm_dur = agg['avg_pm_shift_hm']?.avg || 0;

  return (
    <div className="team-kpi-btn">
      <span className="tk-name">{teamLabel}</span>
      <span className="tk-count">{repCount} reps</span>
      <div className="tk-dropdown">
        <div className="tk-dd-title">{teamLabel} Averages</div>
        <div className="tk-dd-grid">
          <div>Working Days</div><div>AM: {fmtVal(d_am_days)} | PM: {fmtVal(d_pm_days)}</div>
          <div>Covered Doctors</div><div>AM: {fmtVal(d_am_doc)} | PM: {fmtVal(d_pm_doc)}</div>
          <div>Shift Duration</div><div>AM: {fmtVal(d_am_dur)}h | PM: {fmtVal(d_pm_dur)}h</div>
        </div>
      </div>
    </div>
  );
}

// ── Pivot summary banner ──────────────────────────────────────────────────────
function PivotSummaryBanner({ rows, valueKey, rowKey, shift, t }) {
  // Per-team totals and per-user totals for the pivot data
  const filtered = useMemo(() => shift==='all'?rows:rows.filter(r=>r.shift===shift), [rows,shift]);

  const byTeam = useMemo(() => {
    const m = {};
    filtered.forEach(r => {
      const team = r.team || 'Unknown';
      if(!m[team]) m[team] = { total:0, users:new Set() };
      m[team].total += (r[valueKey]||0);
      m[team].users.add(r.user_name);
    });
    return m;
  }, [filtered, valueKey]);

  const grandTotal = useMemo(() => filtered.reduce((s,r)=>s+(r[valueKey]||0),0), [filtered,valueKey]);
  const allUsers   = useMemo(() => new Set(filtered.map(r=>r.user_name)).size, [filtered]);
  const teamList   = Object.entries(byTeam).sort((a,b)=>a[0].localeCompare(b[0]));
  if(!filtered.length) return null;

  return (
    <div className="pivot-banner">
      <div className="pivot-banner-total">
        <span className="pb-label">Grand Total</span>
        <span className="pb-val">{grandTotal.toLocaleString()}</span>
        <span className="pb-sub">{allUsers} reps</span>
      </div>
      {teamList.map(([team,d]) => (
        <div key={team} className="pivot-banner-team">
          <span className="pb-team">{team}</span>
          <span className="pb-val">{d.total.toLocaleString()}</span>
          <span className="pb-sub">{d.users.size} reps · avg {d.users.size?Math.round(d.total/d.users.size):0}</span>
        </div>
      ))}
    </div>
  );
}

// ── ShiftToggle ───────────────────────────────────────────────────────────────
function ShiftToggle({ value, onChange, t }) {
  return (
    <div className="shift-toggle">
      {['all','AM','PM'].map(s=>(
        <button key={s}
          className={`stoggle${value===s?' on':''} ${s==='AM'?'am':s==='PM'?'pm':''}`}
          onClick={()=>onChange(s)}>
          {s==='all'?t.shiftAll:s}
        </button>
      ))}
    </div>
  );
}

// ── Pivot table ───────────────────────────────────────────────────────────────
function PivotTable({ rows, rowKey, valueKey, shiftFilter, userFilter, searchFilter, lang }) {
  const filtered = useMemo(()=> rows.filter(r=>{
    if(shiftFilter!=='all' && r.shift!==shiftFilter) return false;
    if(userFilter && userFilter!=='all' && r.user_name!==userFilter) return false;
    if(searchFilter && !r[rowKey]?.toLowerCase().includes(searchFilter.toLowerCase())
       && !r.user_name?.toLowerCase().includes(searchFilter.toLowerCase())) return false;
    return true;
  }), [rows, rowKey, shiftFilter, userFilter, searchFilter]);

  const users = useMemo(()=>[...new Set(filtered.map(r=>r.user_name))].sort(),[filtered]);
  const rowKeys = useMemo(()=>[...new Set(filtered.map(r=>r[rowKey]))].sort(),[filtered,rowKey]);
  const cells = useMemo(()=>{
    const c={};
    filtered.forEach(r=>{
      const k=r[rowKey]; if(!c[k])c[k]={};
      c[k][r.user_name]=(c[k][r.user_name]||0)+(r[valueKey]||0);
    });
    return c;
  },[filtered,rowKey,valueKey]);

  const colTotals = useMemo(()=>{
    const ct={};
    users.forEach(u=>ct[u]=filtered.filter(r=>r.user_name===u).reduce((s,r)=>s+(r[valueKey]||0),0));
    return ct;
  },[users,filtered,valueKey]);

  if(!filtered.length) return <div className="dash-empty">{lang==='ar'?'لا توجد بيانات':'No data'}</div>;

  return (
    <div className="pivot-wrap">
      <table className="pivot-tbl">
        <thead>
          <tr>
            <th className="s-col">{rowKey==='specialty'?(lang==='ar'?'التخصص':'Specialty'):(lang==='ar'?'المنتج':'Product')}</th>
            {users.map(u=><th key={u} title={u}>{u.split(' ').slice(0,2).join(' ')}</th>)}
            <th className="t-col">Σ Total</th>
          </tr>
          {/* Avg row in header */}
          <tr className="avg-row">
            <th className="s-col avg-lbl">⌀ Avg / rep</th>
            {users.map(u=>{
              const uTotal = colTotals[u]||0;
              const uRows  = rowKeys.filter(k=>cells[k]?.[u]).length;
              return <th key={u} className="avg-cell">{uRows>0?Math.round(uTotal/uRows):0}</th>;
            })}
            <th className="t-col avg-cell">
              {(() => {
                const gt = filtered.reduce((s,r)=>s+(r[valueKey]||0),0);
                return users.length>0?Math.round(gt/users.length):0;
              })()}
            </th>
          </tr>
        </thead>
        <tbody>
          {rowKeys.map(k=>{
            const rowTotal=users.reduce((s,u)=>s+(cells[k]?.[u]||0),0);
            return (
              <tr key={k}>
                <td className="s-col">{k}</td>
                {users.map(u=>{
                  const v=cells[k]?.[u];
                  return <td key={u} className={v?'has-v':'nil'}>{v||''}</td>;
                })}
                <td className="t-col">{rowTotal}</td>
              </tr>
            );
          })}
          <tr className="tot-row">
            <td className="s-col">Σ Total</td>
            {users.map(u=><td key={u}>{colTotals[u]||0}</td>)}
            <td className="t-col">{filtered.reduce((s,r)=>s+(r[valueKey]||0),0)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ── Main Dashboard ────────────────────────────────────────────────────────────
export default function Dashboard() {
  const { profile, visibleCodes, signOut } = useAuth();
  const [lang, setLang]       = useState(profile?.preferred_lang||'en');
  const [period, setPeriod]   = useState('recent');
  const [team, setTeam]       = useState('all');
  const [shift, setShift]     = useState('all');
  const [search, setSearch]   = useState('');
  const [userFilter, setUser] = useState('all');
  const [tab, setTab]         = useState('summary');
  const [summary, setSummary]       = useState([]);
  const [specialty, setSpecialty]   = useState([]);
  const [products, setProducts]     = useState([]);
  const [coaching, setCoaching]     = useState([]);
  const [loading, setLoading]       = useState(true);
  const [error, setError]           = useState('');

  const t   = T[lang]||T.en;
  const rtl = lang==='ar';
  const isMgr = profile?.role && profile.role!=='MR';
  const periodLabel = period==='last_month'?'Last Month':'Recent';

  const load = useCallback(async()=>{
    if(!visibleCodes?.length){setLoading(false);return;}
    setLoading(true); setError('');
    const codes=visibleCodes;
    const cf=codes.map(c=>`manager_code.eq.${c}`).concat(codes.map(c=>`rep_code.eq.${c}`)).join(',');
    const queries=[
      supabase.from('summaries').select('*').eq('period',periodLabel).in('employee_code',codes),
      supabase.from('specialty_classification').select('*').eq('period',periodLabel).in('employee_code',codes),
      supabase.from('product_calls').select('*').eq('period',periodLabel).in('employee_code',codes),
    ];
    if(isMgr) queries.push(supabase.from('coaching_days').select('*').eq('period',periodLabel).or(cf));
    const [s,sp,pr,co]=await Promise.all(queries);
    if(s.error) setError(s.error.message);
    setSummary(s.data||[]);
    setSpecialty(sp.data||[]);
    setProducts(pr.data||[]);
    setCoaching(isMgr?(co?.data||[]):[]);
    setLoading(false);
  },[periodLabel,visibleCodes,isMgr]);

  useEffect(()=>{load();},[load]);

  const teams=useMemo(()=>[...new Set(summary.map(r=>r.team).filter(Boolean))].sort(),[summary]);
  const byTeam=useCallback(rows=>{
    if(team==='all') return rows;
    const teamUserNames = new Set(summary.filter(r=>r.team===team).map(r=>r.user_name));
    return rows.filter(r => (r.team !== undefined && r.team !== null && r.team !== '') ? r.team === team : teamUserNames.has(r.user_name));
  },[team, summary]);

  const fSummary=useMemo(()=>{
    let r=sortSummary(byTeam(summary));
    if(search) r=r.filter(x=>x.user_name?.toLowerCase().includes(search.toLowerCase())||x.territory?.toLowerCase().includes(search.toLowerCase()));
    if(userFilter!=='all') r=r.filter(x=>x.user_name===userFilter);
    return r;
  },[summary,byTeam,search,userFilter]);

  const fSpecialty = useMemo(()=>byTeam(specialty),[specialty,byTeam]);
  const fProducts  = useMemo(()=>byTeam(products),[products,byTeam]);
  const fCoaching  = useMemo(()=>{
    let r=byTeam(coaching);
    if(search) r=r.filter(x=>x.manager_name?.toLowerCase().includes(search.toLowerCase())||x.rep_name?.toLowerCase().includes(search.toLowerCase()));
    return r;
  },[coaching,byTeam,search]);

  const allUsers=useMemo(()=>[...new Set(byTeam(summary).map(r=>r.user_name))].sort(),[summary,byTeam]);
  const teamCount=new Set(fSummary.map(r=>r.team)).size;

  // Group summary rows by team for aggregate cards (when viewing all teams)
  const teamGroups = useMemo(()=>{
    if(team!=='all') return [{ label: team||'Team', rows: fSummary }];
    const groups = {};
    fSummary.forEach(r=>{
      const tm = r.team||'Unknown';
      if(!groups[tm]) groups[tm]=[];
      groups[tm].push(r);
    });
    return Object.entries(groups).sort((a,b)=>a[0].localeCompare(b[0])).map(([label,rows])=>({label,rows}));
  },[fSummary,team]);

  const visibleTabs=useMemo(()=>{
    const all=Object.entries(t.tabs);
    return isMgr?all:all.filter(([k])=>k!=='coaching');
  },[t.tabs,isMgr]);

  function doExport(){
    const wb=XLSX.utils.book_new();
    const allKpiKeys=t.kpiGroups.flatMap(g=>g.keys);
    const sh=[['Team','User','Territory','Manager',...allKpiKeys.map(k=>t.kpi[k]||k)]];
    fSummary.forEach(r=>sh.push([r.team,r.user_name,r.territory,r.is_manager?'✓':'',...allKpiKeys.map(k=>r[k]??'')]));
    // Add aggregate sheet
    const aggRows=[['Team','KPI','Sum','Avg']];
    teamGroups.forEach(({label,rows})=>{
      const {agg}=computeAggregates(rows);
      NUMERIC_KPI_KEYS.forEach(k=>{
        if(agg[k]) aggRows.push([label,t.kpi[k]||k,agg[k].sum,+agg[k].avg.toFixed(2)]);
      });
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sh),'Summary');
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aggRows),'Team Averages');
    XLSX.writeFile(wb,`excellence_${periodLabel.replace(' ','_')}_${Date.now()}.xlsx`);
  }

  return (
    <div className={`dash${rtl?' rtl':''}`} dir={rtl?'rtl':'ltr'}>

      {/* HEADER */}
      <header className="dash-hdr">
        <div className="dash-hdr-l">
          <span className="dash-brand">{t.brand}</span>
          <div className="dash-sep"/>
          <span className="dash-view">{t.roleView[profile?.role]||''}</span>
        </div>
        <div className="dash-hdr-r">
          {profile?.role==='Admin'&&<a className="hbtn hbtn-outline" href="#/admin">{t.adminPanel}</a>}
          <button className="hbtn hbtn-lang" onClick={()=>setLang(lang==='en'?'ar':'en')}>
            {lang==='en'?'عربي':'EN'}
          </button>
          <div className="dash-user">
            <div className="du-name">{profile?.employee_name}</div>
            <div className="du-role">{profile?.role} · {profile?.employee_code}</div>
          </div>
          <button className="hbtn hbtn-outline" onClick={signOut}>{t.signOut}</button>
        </div>
      </header>

      {/* CONTROL BAR */}
      <div className="ctrl-bar">
        <div className="ctrl-row">
          <div className="ctrl-group">
            <span className="ctrl-lbl">{rtl?'الفترة':'Period'}</span>
            <div className="shift-toggle">
              <button className={`stoggle${period==='recent'?' on':''}`} onClick={()=>setPeriod('recent')}>{t.recent}</button>
              <button className={`stoggle${period==='last_month'?' on':''}`} disabled style={{opacity:0.5, cursor:'not-allowed'}}>{t.lastMonth}</button>
            </div>
          </div>
          <div className="ctrl-group">
            <span className="ctrl-lbl">{rtl?'الوردية':'Shift'}</span>
            <ShiftToggle value={shift} onChange={setShift} t={t}/>
          </div>
          {teams.length>1&&(
            <div className="ctrl-group">
              <span className="ctrl-lbl">{rtl?'الفريق':'Team'}</span>
              <select className="ctrl-sel" value={team} onChange={e=>{setTeam(e.target.value);setUser('all');}}>
                <option value="all">{t.allTeams}</option>
                {teams.map(tm=><option key={tm} value={tm}>{tm}</option>)}
              </select>
            </div>
          )}
          {isMgr&&allUsers.length>1&&(
            <div className="ctrl-group">
              <span className="ctrl-lbl">{rtl?'المندوب':'Rep'}</span>
              <select className="ctrl-sel" value={userFilter} onChange={e=>setUser(e.target.value)}>
                <option value="all">{t.allUsers}</option>
                {allUsers.map(u=><option key={u} value={u}>{u}</option>)}
              </select>
            </div>
          )}
          <div className="ctrl-group ctrl-search">
            <div className="search-box">
              <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="9" cy="9" r="6"/><path d="M15 15l-3.5-3.5"/>
              </svg>
              <input className="search-inp" placeholder={t.search}
                value={search} onChange={e=>setSearch(e.target.value)}/>
              {search&&<button className="search-clear" onClick={()=>setSearch('')}>✕</button>}
            </div>
          </div>
          <div className="ctrl-end">
            <span className="ctrl-stat">{t.people(fSummary.length)}{teamCount>1&&` · ${teamCount} teams`}</span>
            <button className="hbtn hbtn-primary" onClick={doExport}>↓ {t.export}</button>
          </div>
        </div>
      </div>

      {/* TABS */}
      <nav className="dash-tabs">
        {visibleTabs.map(([k,label])=>(
          <button key={k} className={`dtab${tab===k?' on':''}`} onClick={()=>setTab(k)}>{label}</button>
        ))}
      </nav>

      {error&&<div className="dash-err">{error}</div>}

      {loading?(
        <div className="dash-empty">{t.loading}</div>
      ):(
        <div className="dash-body">

          {/* SUMMARY TAB */}
          {tab==='summary'&&(
            fSummary.length===0?<div className="dash-empty">{t.noData}</div>:(
              <>
                {/* Aggregate cards — one per team (or one overall) pinned at top */}
                {isMgr&&userFilter==='all'&&(
                  <div className="agg-cards-row">
                    {teamGroups.map(({label,rows})=>(
                      <TeamKpiHover key={label} rows={rows} teamLabel={label}/>
                    ))}
                  </div>
                )}
                {/* Individual user cards */}
                <div className="cards-grid">
                  {fSummary.map((r,i)=>(
                    <div key={r.id||i} className={`ucard${r.is_manager?' mgr':''}`}>
                      <div className="ucard-hdr">
                        <div className="ucard-info">
                          <div className="ucard-name">{r.user_name}</div>
                          <div className="ucard-meta">{r.team||''}{r.is_manager?' · Manager':''}</div>
                          {r.territory&&<div className="ucard-terr" title={r.territory}>{r.territory}</div>}
                        </div>
                        {r.is_manager&&<span className="mgr-pip">MGR</span>}
                      </div>
                      {t.kpiGroups.map(g=>{
                        const keys=g.keys.filter(k=>{
                          if(shift==='AM') return !['pm_calls','pm_call_rate','pm_shift_days','total_pm_covered','clinic_covered','polyclinic_covered','avg_pm_shift_hm'].includes(k);
                          if(shift==='PM') return !['am_calls','am_call_rate','am_shift_days','total_am_covered','amcenter_covered','hospital_covered','avg_am_shift_hm','avg_am_start_time'].includes(k);
                          return true;
                        });
                        if(g.keys.includes('coaching_days')&&!isMgr) return null;
                        const kpiRows=keys.map(k=>({k,v:r[k]})).filter(x=>x.v!==null&&x.v!==undefined&&x.v!=='');
                        if(!kpiRows.length) return null;
                        return (
                          <div key={g.label} className="kpi-sec">
                            <div className="kpi-sec-hd">{g.label}</div>
                            {kpiRows.map(({k,v})=>(
                              <div key={k} className="kpi-row">
                                <span className="kpi-lbl">{t.kpi[k]||k}</span>
                                <span className={`kpi-v${k.includes('rate')?' rate':''}`}>{fmtVal(v,k)}</span>
                              </div>
                            ))}
                          </div>
                        );
                      })}
                      {r.product_calls_detail&&shift!=='AM'&&(
                        <div className="kpi-sec">
                          <div className="kpi-sec-hd">{rtl?'تفاصيل المنتج':'Product Detail'}</div>
                          <div className="prod-det">{r.product_calls_detail}</div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </>
            )
          )}

          {/* SPECIALTY TAB */}
          {tab==='specialty'&&(
            <>
              <PivotSummaryBanner rows={fSpecialty} valueKey="call_count" rowKey="specialty" shift={shift} t={t}/>
              <PivotTable rows={fSpecialty} rowKey="specialty" valueKey="call_count"
                shiftFilter={shift} userFilter={userFilter} searchFilter={search} lang={lang}/>
            </>
          )}

          {/* PRODUCTS TAB */}
          {tab==='products'&&(
            <>
              <PivotSummaryBanner rows={fProducts} valueKey="call_count" rowKey="product" shift={shift} t={t}/>
              <PivotTable rows={fProducts} rowKey="product" valueKey="call_count"
                shiftFilter={shift} userFilter={userFilter} searchFilter={search} lang={lang}/>
            </>
          )}

          {/* COACHING TAB */}
          {tab==='coaching'&&isMgr&&(
            fCoaching.length===0?<div className="dash-empty">{t.noData}</div>:(
              <div className="pivot-wrap">
                <table className="pivot-tbl">
                  <thead>
                    <tr>
                      <th className="s-col">{rtl?'المدير':'Manager'}</th>
                      <th>{rtl?'المندوب':'Rep'}</th>
                      <th>{rtl?'التاريخ':'Date'}</th>
                      <th>{rtl?'الفريق':'Team'}</th>
                      <th>{rtl?'زيارات AM':'AM Visits'}</th>
                      <th>{rtl?'مرافقة AM':'AM Acc.'}</th>
                      <th>AM %</th>
                      <th>{rtl?'زيارات PM':'PM Visits'}</th>
                      <th>{rtl?'مرافقة PM':'PM Acc.'}</th>
                      <th>PM %</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...fCoaching].sort((a,b)=>(a.manager_name||'').localeCompare(b.manager_name||'')).map((r,i)=>(
                      <tr key={r.id||i}>
                        <td className="s-col">{r.manager_name}</td>
                        <td>{r.rep_name}</td>
                        <td>{r.coaching_date}</td>
                        <td>{r.team||'—'}</td>
                        <td>{r.am_visits||0}</td>
                        <td>{r.am_accompanied||0}</td>
                        <td>{r.am_visits ? Math.round((r.am_accompanied/r.am_visits)*100)+'%' : '-'}</td>
                        <td>{r.pm_visits||0}</td>
                        <td>{r.pm_accompanied||0}</td>
                        <td>{r.pm_visits ? Math.round((r.pm_accompanied/r.pm_visits)*100)+'%' : '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
          )}
        </div>
      )}
    </div>
  );
}
