import React, { useEffect, useMemo, useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../supabaseClient';
import { useAuth } from '../contexts/AuthContext';
import { useToast } from '../contexts/ToastContext';
import ChartBuilder from '../components/ChartBuilder';
import { SkeletonCardGrid, SkeletonTable } from '../components/Skeleton';
import './Dashboard.css';

// ─── i18n ────────────────────────────────────────────────────────────────────
const T = {
  en: {
    brandMain:'Excellence - CRM', brandSub:'web app', signOut:'Sign out', adminPanel:'Admin Panel',
    lastMonth:'Last Month', recent:'Recent  1–15',
    allTeams:'All teams', allUsers:'All reps', search:'Search name or territory…',
    export:'Export', loading:'Loading…', noData:'No data for this period.',
    shiftAll:'Both', shiftAM:'AM', shiftPM:'PM',
    people: n=>`${n} rep${n!==1?'s':''}`,
    tabs:{ summary:'Summary', specialty:'Specialty', products:'Products', coaching:'Coaching' },
    roleView:{ MR:'My Results', Supervisor:'My Team', 'Area Manager':'My Area', BLM:'Full Team', Admin:'All Teams' },
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
    brandMain:'إكسيلنس - CRM', brandSub:'تطبيق ويب', signOut:'خروج', adminPanel:'لوحة الإدارة',
    lastMonth:'الشهر الماضي', recent:'الأحدث  1–15',
    allTeams:'كل الفرق', allUsers:'كل المندوبين', search:'بحث باسم أو منطقة…',
    export:'تصدير', loading:'جارٍ التحميل…', noData:'لا توجد بيانات.',
    shiftAll:'الكل', shiftAM:'AM', shiftPM:'PM',
    people: n=>`${n} مندوب`,
    tabs:{ summary:'الملخص', specialty:'التخصص', products:'المنتجات', coaching:'التوجيه' },
    roleView:{ MR:'نتائجي', Supervisor:'فريقي', 'Area Manager':'منطقتي', BLM:'الفريق', Admin:'الكل' },
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

function fmtVal(v, key) {
  if (v===null||v===undefined||v==='') return '—';
  if (key?.includes('rate')) return Number(v).toFixed(1);
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

// ─── ShiftToggle ─────────────────────────────────────────────────────────────
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

// ─── Pivot table ─────────────────────────────────────────────────────────────
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

  if(!filtered.length) return <div className="empty-state">{lang==='ar'?'لا توجد بيانات':'No data'}</div>;

  return (
    <div className="table-scroll-container" style={{ overflowX: 'auto', WebkitOverflowScrolling: 'touch', width: '100%' }}>
      <table className="pivot-tbl">
        <thead>
          <tr>
            <th className="s-col">{rowKey==='specialty'?(lang==='ar'?'التخصص':'Specialty'):(lang==='ar'?'المنتج':'Product')}</th>
            {users.map(u=><th key={u} title={u}>{u.split(' ').slice(0,2).join(' ')}</th>)}
            <th className="t-col">Total</th>
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
            <td className="s-col">Total</td>
            {users.map(u=><td key={u}>{filtered.filter(r=>r.user_name===u).reduce((s,r)=>s+(r[valueKey]||0),0)}</td>)}
            <td className="t-col">{filtered.reduce((s,r)=>s+(r[valueKey]||0),0)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ─── Main Dashboard ───────────────────────────────────────────────────────────
export default function Dashboard() {
  const { profile, visibleCodes, signOut } = useAuth();
  const { success: toastSuccess, error: toastError } = useToast();
  const [lang, setLang]       = useState(profile?.preferred_lang||'en');
  const [period, setPeriod]   = useState('last_month');
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

  // ── Load data ──────────────────────────────────────────────────────────────
  const load = useCallback(async()=>{
    if(!visibleCodes?.length){setLoading(false);return;}
    setLoading(true); setError('');
    const codes=visibleCodes;
    const queries=[
      supabase.from('summaries').select('*').eq('period',periodLabel).in('employee_code',codes),
      supabase.from('specialty_classification').select('*').eq('period',periodLabel).in('employee_code',codes),
      supabase.from('product_calls').select('*').eq('period',periodLabel).in('employee_code',codes),
    ];
    // No manual manager_code/rep_code filter here — the RLS policy on
    // coaching_days already enforces the exact same visibility rule at the
    // database level. Building an .or() filter listing every visible code
    // client-side produced URLs tens of thousands of characters long for
    // high-visibility roles (Admin/BLM), which silently failed — RLS makes
    // that redundant and lets this stay a plain period-scoped query.
    if(isMgr) queries.push(supabase.from('coaching_days').select('*').eq('period',periodLabel));
    const [s,sp,pr,co]=await Promise.all(queries);
    if(s.error) setError(s.error.message);
    setSummary(s.data||[]);
    setSpecialty(sp.data||[]);
    setProducts(pr.data||[]);
    setCoaching(isMgr?(co?.data||[]):[]);
    setLoading(false);
  },[periodLabel,visibleCodes,isMgr]);

  useEffect(()=>{load();},[load]);

  // ── Filter chains ──────────────────────────────────────────────────────────
  const teams=useMemo(()=>[...new Set(summary.map(r=>r.team).filter(Boolean))].sort(),[summary]);

  const byTeam=useCallback(rows=>team==='all'?rows:rows.filter(r=>r.team===team),[team]);

  const fSummary=useMemo(()=>{
    let raw = sortSummary(byTeam(summary));
    if(search) raw = raw.filter(x=>x.user_name?.toLowerCase().includes(search.toLowerCase())||x.territory?.toLowerCase().includes(search.toLowerCase()));
    if(userFilter!=='all') raw = raw.filter(x=>x.user_name===userFilter);
    
    // Aggregate duplicates by employee_code
    const aggregated = new Map();
    raw.forEach(r => {
      if (!r.employee_code) return; // Skip if no code
      if (!aggregated.has(r.employee_code)) {
        aggregated.set(r.employee_code, { ...r });
      } else {
        const existing = aggregated.get(r.employee_code);
        // Sum numeric fields
        Object.keys(r).forEach(k => {
          if (typeof r[k] === 'number') {
            existing[k] = (existing[k] || 0) + r[k];
          }
        });
        // Merge text fields (if different, just append or keep first)
        // Recalculate rates since we summed the raw counts
        if (existing.field_days) {
            existing.am_call_rate = existing.am_calls / existing.field_days;
            existing.pm_call_rate = existing.pm_calls / existing.field_days;
        }
      }
    });
    
    return Array.from(aggregated.values());
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

  // ── Tabs (hide Coaching & Charts for MR) ────────────────────────────────────────────
  const visibleTabs=useMemo(()=>{
    const all=Object.entries(t.tabs);
    return isMgr?all:all.filter(([k])=>k!=='coaching');
  },[t.tabs,isMgr]);

  // ── Export ─────────────────────────────────────────────────────────────────
  function doExport(){
    try {
      const wb=XLSX.utils.book_new();
      const allKpiKeys=t.kpiGroups.flatMap(g=>g.keys);
      const sh=[['Team','User','Territory','Manager',...allKpiKeys.map(k=>t.kpi[k]||k)]];
      fSummary.forEach(r=>sh.push([r.team,r.user_name,r.territory,r.is_manager?'✓':'',...allKpiKeys.map(k=>r[k]??'')]));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sh),'Summary');
      const filename=`excellence_${periodLabel.replace(' ','_')}_${Date.now()}.xlsx`;
      XLSX.writeFile(wb,filename);
      toastSuccess(`✓ Exported ${fSummary.length} rows to Excel`);
    } catch(e) {
      toastError('Export failed: ' + e.message);
    }
  }

  // ── Render ─────────────────────────────────────────────────────────────────
  return (
    <div className={`dash${rtl?' rtl':''}`} dir={rtl?'rtl':'ltr'}>

      {/* ── HEADER ── */}
      <header className="dash-hdr">
        <div className="dash-hdr-l">
          <div className="dash-brand-wrap">
            <span className="dash-brand">{t.brandMain}</span>
            <span className="dash-brand-sub">{t.brandSub}</span>
          </div>
          <div className="dash-sep"/>
          <span className="dash-view">{t.roleView[profile?.role]||''}</span>
        </div>
        <div className="dash-hdr-r">
          {profile?.role==='Admin'&&<a className="hbtn hbtn-outline" href="#/admin">{t.adminPanel}</a>}
          <button className="hbtn hbtn-outline" style={{padding: '8px 12px', fontSize: '16px'}} onClick={() => {
            const isLight = document.documentElement.classList.toggle('light');
            localStorage.setItem('theme', isLight ? 'light' : 'dark');
          }} title="Toggle Theme">
            ◐
          </button>
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

      {/* ── CONTROL BAR ── */}
      <div className="ctrl-bar">
        {/* Row 1: Period + Shift + Team + User + Search */}
        <div className="ctrl-row">
          {/* Period */}
          <div className="ctrl-group">
            <span className="ctrl-lbl">{rtl?'الفترة':'Period'}</span>
            <div className="seg-ctrl">
              {[{k:'last_month',l:t.lastMonth},{k:'recent',l:t.recent}].map(p=>(
                <button key={p.k} className={`seg${period===p.k?' on':''}`} onClick={()=>setPeriod(p.k)}>{p.l}</button>
              ))}
            </div>
          </div>

          {/* Shift */}
          <div className="ctrl-group">
            <span className="ctrl-lbl">{rtl?'الوردية':'Shift'}</span>
            <ShiftToggle value={shift} onChange={setShift} t={t}/>
          </div>

          {/* Team */}
          {teams.length>1&&(
            <div className="ctrl-group">
              <span className="ctrl-lbl">{rtl?'الفريق':'Team'}</span>
              <select className="ctrl-sel" value={team} onChange={e=>{setTeam(e.target.value);setUser('all');}}>
                <option value="all">{t.allTeams}</option>
                {teams.map(tm=><option key={tm} value={tm}>{tm}</option>)}
              </select>
            </div>
          )}

          {/* User */}
          {isMgr&&allUsers.length>1&&(
            <div className="ctrl-group">
              <span className="ctrl-lbl">{rtl?'المندوب':'Rep'}</span>
              <select className="ctrl-sel" value={userFilter} onChange={e=>setUser(e.target.value)}>
                <option value="all">{t.allUsers}</option>
                {allUsers.map(u=><option key={u} value={u}>{u}</option>)}
              </select>
            </div>
          )}

          {/* Search */}
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

          {/* Stats + Export */}
          <div className="ctrl-end">
            <span className="ctrl-stat">{t.people(fSummary.length)}{teamCount>1&&` · ${teamCount} teams`}</span>
            <button className="hbtn hbtn-primary" onClick={doExport}>↓ {t.export}</button>
          </div>
        </div>
      </div>

      {/* ── TABS ── */}
      <nav className="dash-tabs">
        {visibleTabs.map(([k,label])=>(
          <button key={k} className={`dtab${tab===k?' on':''}`} onClick={()=>setTab(k)}>
            {label}
          </button>
        ))}
      </nav>

      {error&&<div className="dash-err">{error}</div>}

      {/* ── CONTENT ── */}
      {loading?(
        <div className="dash-body" style={{maxWidth:1600,margin:'0 auto',padding:'24px 32px'}}>
          {tab==='summary' && <SkeletonCardGrid count={6}/>}
          {(tab==='specialty'||tab==='products'||tab==='coaching') && <SkeletonTable rows={8} cols={5}/>}
        </div>
      ):(
        <div className="dash-main-layout">
          <div className="dash-body">

          {/* ── SUMMARY: vertical KPI cards ── */}
          {tab==='summary'&&(
            fSummary.length===0?<div className="dash-empty">{t.noData}</div>:(
              <div className="cards-grid">
                {fSummary.map((r,i)=>(
                  <div key={r.id||i} className={`ucard${r.is_manager?' mgr':''}`}>
                    {/* Card header */}
                    <div className="ucard-hdr">
                      <div className="ucard-info">
                        <div className="ucard-name">{r.user_name}</div>
                        <div className="ucard-meta">{r.team||''}{r.is_manager?' · Manager':''}</div>
                        {r.territory&&<div className="ucard-terr" title={r.territory}>{r.territory}</div>}
                      </div>
                      {r.is_manager&&<span className="mgr-pip">MGR</span>}
                    </div>
                    {/* KPI groups — filter by shift */}
                    {t.kpiGroups.map(g=>{
                      // Filter keys by shift selection
                      const keys=g.keys.filter(k=>{
                        if(shift==='all') return true;
                        if(shift==='AM') return !['pm_calls','pm_call_rate','pm_shift_days','total_pm_covered','clinic_covered','polyclinic_covered','avg_pm_shift_hm'].includes(k);
                        if(shift==='PM') return !['am_calls','am_call_rate','am_shift_days','total_am_covered','amcenter_covered','hospital_covered','avg_am_shift_hm','avg_am_start_time'].includes(k);
                        return true;
                      });
                      // Skip coaching section for MR
                      if(g.keys.includes('coaching_days')&&!isMgr) return null;
                      const rows=keys.map(k=>({k,v:r[k]})).filter(x=>x.v!==null&&x.v!==undefined&&x.v!=='');
                      if(!rows.length) return null;
                      return (
                        <div key={g.label} className="kpi-sec">
                          <div className="kpi-sec-hd">{g.label}</div>
                          {rows.map(({k,v})=>(
                            <div key={k} className="kpi-row">
                              <span className="kpi-lbl">{t.kpi[k]||k}</span>
                              <span className={`kpi-v${k.includes('rate')?' rate':''}`}>{fmtVal(v,k)}</span>
                            </div>
                          ))}
                        </div>
                      );
                    })}
                    {/* Product detail */}
                    {r.product_calls_detail&&shift!=='AM'&&(
                      <div className="kpi-sec">
                        <div className="kpi-sec-hd">{rtl?'تفاصيل المنتج':'Product Detail'}</div>
                        <div className="prod-det">{r.product_calls_detail}</div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )
          )}

          {/* ── SPECIALTY pivot ── */}
          {tab==='specialty'&&(
            <PivotTable rows={fSpecialty} rowKey="specialty" valueKey="call_count"
              shiftFilter={shift} userFilter={userFilter} searchFilter={search} lang={lang}/>
          )}

          {/* ── PRODUCTS pivot ── */}
          {tab==='products'&&(
            <PivotTable rows={fProducts} rowKey="product" valueKey="call_count"
              shiftFilter={shift} userFilter={userFilter} searchFilter={search} lang={lang}/>
          )}

          {/* ── COACHING ── (managers only, already filtered via visibleTabs) */}
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
                      <th style={{textAlign: 'center'}}>AM Visits</th>
                      <th style={{textAlign: 'center'}}>AM Acc</th>
                      <th style={{textAlign: 'center'}}>AM %</th>
                      <th style={{textAlign: 'center'}}>PM Visits</th>
                      <th style={{textAlign: 'center'}}>PM Acc</th>
                      <th style={{textAlign: 'center'}}>PM %</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...fCoaching].sort((a,b)=>(a.manager_name||'').localeCompare(b.manager_name||'')).map((r,i)=>{
                      const amPct = r.am_visits > 0 ? Math.round((r.am_accompanied / r.am_visits) * 100) : 0;
                      const pmPct = r.pm_visits > 0 ? Math.round((r.pm_accompanied / r.pm_visits) * 100) : 0;
                      return (
                      <tr key={r.id||i}>
                        <td className="s-col">{r.manager_name}</td>
                        <td>{r.rep_name}</td>
                        <td>{r.coaching_date}</td>
                        <td>{r.team||'—'}</td>
                        <td style={{textAlign: 'center'}}>{r.am_visits || 0}</td>
                        <td style={{textAlign: 'center'}}>{r.am_accompanied || 0}</td>
                        <td style={{textAlign: 'center'}}>{amPct}%</td>
                        <td style={{textAlign: 'center'}}>{r.pm_visits || 0}</td>
                        <td style={{textAlign: 'center'}}>{r.pm_accompanied || 0}</td>
                        <td style={{textAlign: 'center'}}>{pmPct}%</td>
                      </tr>
                    )})}
                  </tbody>
                </table>
              </div>
            )
          )}

          </div>

          {/* ── CHARTS SIDE PANEL ── */}
          {isMgr&&(
            <aside className="dash-side-panel">
              <ChartBuilder data={fSummary} isManager={isMgr} />
            </aside>
          )}
        </div>
      )}
    </div>
  );
}