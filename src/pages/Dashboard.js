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

const PIE_COLORS = [
  '#3b82f6','#8b5cf6','#10b981','#f59e0b','#ef4444',
  '#06b6d4','#ec4899','#84cc16','#f97316','#6366f1',
  '#14b8a6','#e11d48','#a855f7','#0ea5e9','#eab308',
];

// ── Helpers ──────────────────────────────────────────────────────────────────
function fmtDuration(decimalHours) {
  const h = Math.floor(decimalHours);
  const m = Math.round((decimalHours - h) * 60);
  if (h === 0 && m === 0) return '—';
  if (h === 0) return `${m}m`;
  if (m === 0) return `${h}h`;
  return `${h}h ${m}m`;
}

function fmtVal(v, key) {
  if (v===null||v===undefined||v==='') return '—';
  if (key==='avg_am_shift_hm'||key==='avg_pm_shift_hm') return fmtDuration(Number(v));
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

function computeAggregates(rows) {
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

function daysUntil15th() {
  const now = new Date();
  let target = new Date(now.getFullYear(), now.getMonth(), 15, 23, 59, 59);
  if (now > target) target = new Date(now.getFullYear(), now.getMonth() + 1, 15, 23, 59, 59);
  return Math.max(0, Math.ceil((target - now) / (1000 * 60 * 60 * 24)));
}

// ── PieChart (SVG donut) ─────────────────────────────────────────────────────
// ── KPI targets for progress indicators ─────────────────────────────────────
const KPI_TARGETS = {
  working_days: 22,
  complete_field_days: 20,
  am_calls: 120,
  pm_calls: 120,
  total_am_covered: 80,
  total_pm_covered: 80,
  pharmacies_visited: 40,
  coaching_days: 4,
};

// ── PieChart (SVG donut) ─────────────────────────────────────────────────────
function PieChart({ data, title, size = 140, thickness = 22, onSelect, activeFilters = new Set() }) {
  const total = data.reduce((s, d) => s + d.value, 0);
  if (!total) return <div className="pie-empty">No data</div>;
  const center = size / 2;
  const radius = (size - thickness) / 2;
  const circumference = 2 * Math.PI * radius;
  let accumulated = 0;
  const segments = data.map((d) => {
    const pct = d.value / total;
    const arc = pct * circumference;
    const offset = accumulated;
    accumulated += arc;
    return { ...d, pct, arc, offset };
  });

  const hasSelections = activeFilters && activeFilters.size > 0;

  return (
    <div className="pie-chart">
      {title && <div className="pie-title">{title}</div>}
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} className="pie-svg">
        <circle cx={center} cy={center} r={radius} fill="none" stroke="rgba(255,255,255,.06)" strokeWidth={thickness}/>
        {segments.map((seg, i) => {
          const isSelected = activeFilters?.has(seg.label);
          const fade = hasSelections && !isSelected;
          return (
            <circle key={i} cx={center} cy={center} r={radius} fill="none"
              stroke={seg.color} strokeWidth={isSelected ? thickness + 4 : thickness}
              strokeDasharray={`${seg.arc} ${circumference - seg.arc}`}
              strokeDashoffset={-seg.offset}
              transform={`rotate(-90 ${center} ${center})`}
              className="pie-segment" strokeLinecap="butt"
              style={{
                cursor: onSelect ? 'pointer' : 'default',
                opacity: fade ? 0.25 : 1,
                transition: 'opacity 0.2s, stroke-width 0.2s, stroke 0.2s',
              }}
              onClick={() => onSelect && onSelect(seg.label)}
            />
          );
        })}
        <text x={center} y={center-6} textAnchor="middle" dominantBaseline="central" className="pie-center-val">{total}</text>
        <text x={center} y={center+10} textAnchor="middle" dominantBaseline="central" className="pie-center-lbl">calls</text>
      </svg>
      <div className="pie-legend">
        {segments.slice(0, 6).map((seg, i) => {
          const isSelected = activeFilters?.has(seg.label);
          const fade = hasSelections && !isSelected;
          return (
            <div key={i} 
              className={`pie-leg-item ${isSelected ? 'selected' : ''}`}
              style={{
                cursor: onSelect ? 'pointer' : 'default',
                opacity: fade ? 0.4 : 1,
                background: isSelected ? 'rgba(200, 168, 75, 0.12)' : 'none',
                border: isSelected ? '1px solid rgba(200, 168, 75, 0.3)' : '1px solid transparent',
                padding: '4px 6px',
                borderRadius: '4px',
                transition: 'opacity 0.2s, background 0.2s, border-color 0.2s',
              }}
              onClick={() => onSelect && onSelect(seg.label)}
            >
              <span className="pie-dot" style={{ background: seg.color }}/>
              <span className="pie-leg-label">{seg.label}</span>
              <span className="pie-leg-val">{Math.round(seg.pct * 100)}%</span>
            </div>
          );
        })}
        {segments.length > 6 && <div className="pie-leg-more">+{segments.length - 6} more</div>}
      </div>
    </div>
  );
}

// ── TeamKpiHover ─────────────────────────────────────────────────────────────
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
          <div>Shift Duration</div><div>AM: {fmtDuration(d_am_dur)} | PM: {fmtDuration(d_pm_dur)}</div>
        </div>
      </div>
    </div>
  );
}

// ── PivotSummaryBanner ───────────────────────────────────────────────────────
function PivotSummaryBanner({ rows, valueKey, rowKey, shift, t, selectedTeam, onSelectTeam, userTeamMap }) {
  const filtered = useMemo(() => shift==='all'?rows:rows.filter(r=>r.shift===shift), [rows,shift]);
  const byTeam = useMemo(() => {
    const m = {};
    filtered.forEach(r => {
      const team = (r.team && r.team !== 'Unknown') ? r.team : (userTeamMap && userTeamMap[r.user_name]) || 'Other';
      if(!m[team]) m[team] = { total:0, users:new Set() };
      m[team].total += (r[valueKey]||0);
      m[team].users.add(r.user_name);
    });
    return m;
  }, [filtered, valueKey, userTeamMap]);
  const grandTotal = useMemo(() => filtered.reduce((s,r)=>s+(r[valueKey]||0),0), [filtered,valueKey]);
  const allUsers   = useMemo(() => new Set(filtered.map(r=>r.user_name)).size, [filtered]);
  const teamList   = Object.entries(byTeam).sort((a,b)=>a[0].localeCompare(b[0]));
  if(!filtered.length) return null;
  return (
    <div className="pivot-banner">
      <div 
        className={`pivot-banner-total ${selectedTeam === 'all' ? 'active' : ''}`}
        onClick={() => onSelectTeam && onSelectTeam('all')}
        style={{ cursor: onSelectTeam ? 'pointer' : 'default' }}
      >
        <span className="pb-label">Grand Total</span>
        <span className="pb-val">{grandTotal.toLocaleString()}</span>
        <span className="pb-sub">{allUsers} reps</span>
      </div>
      {teamList.map(([team,d]) => (
        <div key={team} 
          className={`pivot-banner-team ${selectedTeam === team ? 'active' : ''}`}
          onClick={() => onSelectTeam && onSelectTeam(selectedTeam === team ? 'all' : team)}
          style={{ cursor: onSelectTeam ? 'pointer' : 'default' }}
        >
          <span className="pb-team">{team}</span>
          <span className="pb-val">{d.total.toLocaleString()}</span>
          <span className="pb-sub">{d.users.size} reps · avg {d.users.size?Math.round(d.total/d.users.size):0}</span>
        </div>
      ))}
    </div>
  );
}

// ── ShiftToggle ──────────────────────────────────────────────────────────────
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

// ── PivotTable ───────────────────────────────────────────────────────────────
function PivotTable({ rows, rowKey, valueKey, shiftFilter, userFilter, searchFilter, lang }) {
  const filtered = useMemo(()=>rows.filter(r=>{
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
  const [shift, setShift]     = useState('PM'); // Default is PM
  const [search, setSearch]   = useState('');
  const [userFilter, setUser] = useState('all');
  const [tab, setTab]         = useState('summary');
  const [summary, setSummary]       = useState([]);
  const [specialty, setSpecialty]   = useState([]);
  const [products, setProducts]     = useState([]);
  const [coaching, setCoaching]     = useState([]);
  const [loading, setLoading]       = useState(true);
  const [error, setError]           = useState('');

  // Sidebar states
  const [selectedRep, setSelectedRep]         = useState(null);
  const [specialtyFilter, setSpecialtyFilter] = useState(new Set());
  const [productFilter, setProductFilter]     = useState(new Set());
  const [classificationFilter, setClassificationFilter] = useState(new Set());
  const [selectedManager, setSelectedManager] = useState(null);
  const [sidebarOpen, setSidebarOpen]         = useState(false);

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

  // ── Filtering ──────────────────────────────────────────────────────────────
  const teams=useMemo(()=>[...new Set(summary.map(r=>r.team).filter(Boolean))].sort(),[summary]);
  const byTeam=useCallback(rows=>{
    if(team==='all') return rows;
    const teamUserNames = new Set(summary.filter(r=>r.team===team).map(r=>r.user_name));
    return rows.filter(r => (r.team !== undefined && r.team !== null && r.team !== '') ? r.team === team : teamUserNames.has(r.user_name));
  },[team, summary]);

  const fSummary=useMemo(()=>{
    let r=byTeam(summary);
    if(search) r=r.filter(x=>x.user_name?.toLowerCase().includes(search.toLowerCase())||x.territory?.toLowerCase().includes(search.toLowerCase()));
    if(userFilter!=='all') r=r.filter(x=>x.user_name===userFilter);
    const m=new Map();
    r.forEach(x=>{
      const k=x.user_name;
      if(!m.has(k)){
        m.set(k,{...x,_am_days_sum:x.am_shift_days||0,_pm_days_sum:x.pm_shift_days||0,_am_dur_sum:(x.avg_am_shift_hm||0)*(x.am_shift_days||0),_pm_dur_sum:(x.avg_pm_shift_hm||0)*(x.pm_shift_days||0)});
      }else{
        const existing=m.get(k);
        const sumKeys=['working_days','complete_field_days','am_shift_days','pm_shift_days','double_visit_days','office_work_days','am_calls','pm_calls','total_am_covered','total_pm_covered','amcenter_covered','hospital_covered','clinic_covered','polyclinic_covered','pharmacies_visited','pharmacies_covered','total_product_calls','distinct_products','coaching_days'];
        sumKeys.forEach(sk=>existing[sk]=(existing[sk]||0)+(x[sk]||0));
        if(x.territory && !existing.territory?.includes(x.territory)){
          existing.territory=existing.territory?`${existing.territory}; ${x.territory}`:x.territory;
        }
        existing._am_days_sum+=(x.am_shift_days||0);
        existing._pm_days_sum+=(x.pm_shift_days||0);
        existing._am_dur_sum+=(x.avg_am_shift_hm||0)*(x.am_shift_days||0);
        existing._pm_dur_sum+=(x.avg_pm_shift_hm||0)*(x.pm_shift_days||0);
        existing.am_call_rate=existing.am_shift_days?(existing.am_calls/existing.am_shift_days):0;
        existing.pm_call_rate=existing.pm_shift_days?(existing.pm_calls/existing.pm_shift_days):0;
        existing.avg_am_shift_hm=existing._am_days_sum>0?(existing._am_dur_sum/existing._am_days_sum):0;
        existing.avg_pm_shift_hm=existing._pm_days_sum>0?(existing._pm_dur_sum/existing._pm_days_sum):0;
      }
    });
    return sortSummary(Array.from(m.values()));
  },[summary,byTeam,search,userFilter]);

  const fSpecialty = useMemo(()=>byTeam(specialty),[specialty,byTeam]);
  const fProducts  = useMemo(()=>byTeam(products),[products,byTeam]);
  const fCoaching  = useMemo(()=>{
    let r=byTeam(coaching);
    if(search) r=r.filter(x=>x.manager_name?.toLowerCase().includes(search.toLowerCase())||x.rep_name?.toLowerCase().includes(search.toLowerCase()));
    return r;
  },[coaching,byTeam,search]);

  const companyAverages = useMemo(() => {
    const reps = summary.filter(r => !r.is_manager);
    const avgs = {};
    NUMERIC_KPI_KEYS.forEach(key => {
      const vals = reps.map(r => Number(r[key])||0).filter(v => v > 0);
      avgs[key] = vals.length ? (vals.reduce((s,v)=>s+v, 0) / vals.length) : 0;
    });
    return avgs;
  }, [summary]);

  const userTeamMap = useMemo(() => {
    const map = {};
    summary.forEach(r => {
      if (r.user_name && r.team) {
        map[r.user_name] = r.team;
      }
    });
    return map;
  }, [summary]);

  const allUsers=useMemo(()=>[...new Set(byTeam(summary).map(r=>r.user_name))].sort(),[summary,byTeam]);
  const teamCount=new Set(fSummary.map(r=>r.team)).size;

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

  // ── Sidebar computed data ──────────────────────────────────────────────────
  const selectedRepData = useMemo(() => {
    if (!selectedRep) return null;
    return fSummary.find(r => r.user_name === selectedRep) || null;
  }, [fSummary, selectedRep]);

  // Specialty pie charts (shift-filtered & rep-filtered)
  const shiftFilteredSpecialty = useMemo(() => {
    let list = fSpecialty;
    if (selectedRep) {
      list = list.filter(r => r.user_name === selectedRep);
    }
    return shift==='all' ? list : list.filter(r => r.shift===shift);
  }, [fSpecialty, shift, selectedRep]);

  const specialtyPieData = useMemo(() => {
    const m = {};
    shiftFilteredSpecialty.forEach(r => { const s=r.specialty||'Other'; m[s]=(m[s]||0)+(r.call_count||0); });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,10)
      .map(([label,value],i) => ({ label, value, color: PIE_COLORS[i%PIE_COLORS.length] }));
  }, [shiftFilteredSpecialty]);

  const classificationPieData = useMemo(() => {
    const m = {};
    shiftFilteredSpecialty.forEach(r => { const c=r.classification||'Unclassified'; m[c]=(m[c]||0)+(r.call_count||0); });
    return Object.entries(m).sort((a,b)=>b[1]-a[1])
      .map(([label,value],i) => ({ label, value, color: PIE_COLORS[i%PIE_COLORS.length] }));
  }, [shiftFilteredSpecialty]);

  const allSpecialties = useMemo(() =>
    [...new Set(fSpecialty.map(r => r.specialty).filter(Boolean))].sort()
  , [fSpecialty]);

  const filteredSpecialty = useMemo(() => {
    let res = fSpecialty;
    if (specialtyFilter.size > 0) {
      res = res.filter(r => specialtyFilter.has(r.specialty));
    }
    if (classificationFilter.size > 0) {
      res = res.filter(r => classificationFilter.has(r.classification));
    }
    return res;
  }, [fSpecialty, specialtyFilter, classificationFilter]);

  const filteredProducts = useMemo(() => {
    if (productFilter.size === 0) return fProducts;
    return fProducts.filter(r => productFilter.has(r.product));
  }, [fProducts, productFilter]);

  const allProducts = useMemo(() =>
    [...new Set(fProducts.map(r => r.product).filter(Boolean))].sort()
  , [fProducts]);

  // Products pie chart (shift-filtered & rep-filtered)
  const shiftFilteredProducts = useMemo(() => {
    let list = fProducts;
    if (selectedRep) {
      list = list.filter(r => r.user_name === selectedRep);
    }
    return shift==='all' ? list : list.filter(r => r.shift===shift);
  }, [fProducts, shift, selectedRep]);

  const productPieData = useMemo(() => {
    const m = {};
    shiftFilteredProducts.forEach(r => { const p=r.product||'Other'; m[p]=(m[p]||0)+(r.call_count||0); });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,10)
      .map(([label,value],i) => ({ label, value, color: PIE_COLORS[i%PIE_COLORS.length] }));
  }, [shiftFilteredProducts]);

  const topProducts = useMemo(() => {
    const m = {};
    shiftFilteredProducts.forEach(r => { const p=r.product||'Other'; m[p]=(m[p]||0)+(r.call_count||0); });
    const sorted = Object.entries(m).sort((a,b)=>b[1]-a[1]);
    const max = sorted[0]?.[1]||1;
    return sorted.slice(0,8).map(([name,count]) => ({ name, count, pct: Math.round(count/max*100) }));
  }, [shiftFilteredProducts]);

  // Coaching manager groups
  const managerGroups = useMemo(() => {
    const m = {};
    fCoaching.forEach(r => {
      const mgr = r.manager_name||'Unknown';
      if(!m[mgr]) m[mgr] = { name: mgr, team: r.team||'', dates: new Set(), reps: new Set() };
      m[mgr].dates.add(r.coaching_date);
      m[mgr].reps.add(r.rep_name);
    });
    return Object.values(m).sort((a,b) => a.team.localeCompare(b.team) || a.name.localeCompare(b.name))
      .map(g => ({ ...g, dayCount: g.dates.size, repCount: g.reps.size }));
  }, [fCoaching]);

  const filteredCoaching = useMemo(() => {
    if (!selectedManager) return fCoaching;
    return fCoaching.filter(r => r.manager_name === selectedManager);
  }, [fCoaching, selectedManager]);

  // ── Actions ────────────────────────────────────────────────────────────────
  function toggleSpecialty(s) {
    setSpecialtyFilter(prev => {
      const next = new Set(prev);
      if (next.has(s)) next.delete(s); else next.add(s);
      return next;
    });
  }

  function toggleProduct(p) {
    setProductFilter(prev => {
      const next = new Set(prev);
      if (next.has(p)) next.delete(p); else next.add(p);
      return next;
    });
  }

  function toggleClassification(c) {
    setClassificationFilter(prev => {
      const next = new Set(prev);
      if (next.has(c)) next.delete(c); else next.add(c);
      return next;
    });
  }

  const handleSelectRep = (repName) => {
    if (selectedRep === repName) {
      setSelectedRep(null);
      setUser('all');
    } else {
      setSelectedRep(repName);
      setUser(repName);
    }
  };

  function changeTab(k) {
    setTab(k);
    setSidebarOpen(false);
  }

  function doExport(){
    const wb=XLSX.utils.book_new();
    const allKpiKeys=t.kpiGroups.flatMap(g=>g.keys);
    const sh=[['Team','User','Territory','Manager',...allKpiKeys.map(k=>t.kpi[k]||k)]];
    fSummary.forEach(r=>sh.push([r.team,r.user_name,r.territory,r.is_manager?'✓':'',...allKpiKeys.map(k=>r[k]??'')]));
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

  // ── Render ─────────────────────────────────────────────────────────────────
  const countdown = daysUntil15th();

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

      {/* LAYOUT: SIDEBAR + CONTENT */}
      <div className="dash-with-sidebar">

        {/* ── SIDEBAR ────────────────────────────────────────────── */}
        {tab !== 'roadmap' && (
          <aside className={`dash-sidebar${sidebarOpen?' open':''}`}>
            <button className="sb-close" onClick={()=>setSidebarOpen(false)}>✕</button>

            {/* ─── SUMMARY SIDEBAR ─────────────────────────────── */}
            {tab==='summary' && (
              <div className="sb-panel">
              {selectedRepData ? (
                <div className="sb-rep-detail">
                  <button className="sb-back" onClick={()=>setSelectedRep(null)}>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>
                    {rtl?'رجوع':'Back to list'}
                  </button>
                  <div className="sb-rep-hdr">
                    <div className="sb-rep-name-lg">{selectedRepData.user_name}</div>
                    <div className="sb-rep-team">{selectedRepData.team}{selectedRepData.is_manager?' · Manager':''}</div>
                    {selectedRepData.territory&&<div className="sb-rep-terr">{selectedRepData.territory}</div>}
                  </div>
                  {t.kpiGroups.map(g => {
                    const keys = g.keys.filter(k => {
                      if(shift==='AM') return !['pm_calls','pm_call_rate','pm_shift_days','total_pm_covered','clinic_covered','polyclinic_covered','avg_pm_shift_hm'].includes(k);
                      if(shift==='PM') return !['am_calls','am_call_rate','am_shift_days','total_am_covered','amcenter_covered','hospital_covered','avg_am_shift_hm','avg_am_start_time'].includes(k);
                      return true;
                    });
                    if(g.keys.includes('coaching_days')&&!isMgr) return null;
                    const kpiRows = keys.map(k=>({k,v:selectedRepData[k]})).filter(x=>x.v!==null&&x.v!==undefined&&x.v!=='');
                    if(!kpiRows.length) return null;
                    return (
                      <div key={g.label} className="sb-kpi-sec">
                        <div className="sb-kpi-hd">{g.label}</div>
                        {kpiRows.map(({k,v})=>(
                          <div key={k} className="sb-kpi-row">
                            <span>{t.kpi[k]||k}</span>
                            <span className="sb-kpi-val">{fmtVal(v,k)}</span>
                          </div>
                        ))}
                      </div>
                    );
                  })}
                </div>
              ) : (
                <>
                  <div className="sb-section-hd">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
                    {rtl?'زيارات PM حسب المندوب':'PM Visits by Rep'}
                  </div>
                  {teamGroups.map(({label, rows}) => {
                    const reps = [...rows].filter(r=>!r.is_manager).sort((a,b)=>(b.pm_calls||0)-(a.pm_calls||0));
                    if(!reps.length) return null;
                    return (
                      <div key={label} className="sb-team-group">
                        <div className="sb-team-label">{label}</div>
                        {reps.map(r => (
                          <div key={r.user_name}
                            className={`sb-rep-row${selectedRep===r.user_name?' active':''}`}
                            onClick={()=>setSelectedRep(r.user_name)}>
                            <span className="sb-rep-name">{r.user_name}</span>
                            <div className="sb-rep-bar-wrap">
                              <div className="sb-rep-bar" style={{width:`${Math.min(100, (r.pm_calls||0) / Math.max(1,...reps.map(x=>x.pm_calls||1)) * 100)}%`}}/>
                            </div>
                            <span className="sb-rep-val">{r.pm_calls||0}</span>
                          </div>
                        ))}
                      </div>
                    );
                  })}
                </>
              )}
            </div>
          )}

          {/* ─── SPECIALTY SIDEBAR ───────────────────────────── */}
          {tab==='specialty' && (
            <div className="sb-panel">
              <div className="sb-section-hd">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v8M8 12h8"/></svg>
                {rtl?'تغطية التخصصات':'Specialty Coverage'}
              </div>
              <PieChart 
                data={specialtyPieData} 
                title={rtl?'حسب التخصص':'By Specialty'}
                onSelect={toggleSpecialty}
                activeFilters={specialtyFilter}
              />

              <div className="sb-divider"/>

              <PieChart 
                data={classificationPieData} 
                title={rtl?'حسب التصنيف':'By Classification'}
                onSelect={toggleClassification}
                activeFilters={classificationFilter}
              />

              <div className="sb-divider"/>

              {/* Specialty Slicer */}
              <div className="sb-slicer">
                <div className="sb-slicer-hd">
                  <span>{rtl?'فلتر التخصص':'Filter Specialty'}</span>
                  {(specialtyFilter.size > 0 || classificationFilter.size > 0) && (
                    <button className="sb-slicer-clear" onClick={()=>{setSpecialtyFilter(new Set()); setClassificationFilter(new Set());}}>
                      {rtl?'مسح':'Clear'}
                    </button>
                  )}
                </div>
                <div className="sb-slicer-pills">
                  {allSpecialties.map(s => (
                    <button key={s}
                      className={`slicer-pill${specialtyFilter.has(s)?' on':''}`}
                      onClick={()=>toggleSpecialty(s)}>
                      {s}
                    </button>
                  ))}
                </div>
              </div>

              <div className="sb-divider"/>

              {/* Countdown Card */}
              <div className="sb-countdown">
                <div className="sb-countdown-ring">
                  <svg width="80" height="80" viewBox="0 0 80 80">
                    <circle cx="40" cy="40" r="34" fill="none" stroke="rgba(255,255,255,.08)" strokeWidth="5"/>
                    <circle cx="40" cy="40" r="34" fill="none" stroke="var(--gold)"
                      strokeWidth="5" strokeLinecap="round"
                      strokeDasharray={`${(1 - countdown/30) * 213.6} 213.6`}
                      transform="rotate(-90 40 40)"/>
                  </svg>
                  <div className="sb-countdown-num">{countdown}</div>
                </div>
                <div className="sb-countdown-label">{rtl?'أيام متبقية':'Days Remaining'}</div>
                <div className="sb-countdown-sub">{rtl?'حتى نهاية الفترة':'Until period closes'}</div>
              </div>
            </div>
          )}

          {/* ─── PRODUCTS SIDEBAR ────────────────────────────── */}
          {tab==='products' && (
            <div className="sb-panel">
              <div className="sb-section-hd">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
                {rtl?'مساهمة المنتجات':'Product Contribution'}
              </div>
              <PieChart 
                data={productPieData} 
                title={rtl?'حسب المنتج':'By Product'}
                onSelect={toggleProduct}
                activeFilters={productFilter}
              />

              <div className="sb-divider"/>

              {/* Product Slicer */}
              <div className="sb-slicer">
                <div className="sb-slicer-hd">
                  <span>{rtl?'فلتر المنتج':'Filter Product'}</span>
                  {productFilter.size > 0 && (
                    <button className="sb-slicer-clear" onClick={()=>setProductFilter(new Set())}>
                      {rtl?'مسح':'Clear'}
                    </button>
                  )}
                </div>
                <div className="sb-slicer-pills">
                  {allProducts.map(p => (
                    <button key={p}
                      className={`slicer-pill${productFilter.has(p)?' on':''}`}
                      onClick={()=>toggleProduct(p)}>
                      {p}
                    </button>
                  ))}
                </div>
              </div>

              <div className="sb-section-hd" style={{marginTop:0}}>
                {rtl?'أعلى المنتجات':'Top Products'}
              </div>
              <div className="sb-top-list">
                {topProducts.map((p,i) => (
                  <div key={p.name} className="sb-top-item">
                    <span className="sb-top-rank">#{i+1}</span>
                    <div className="sb-top-info">
                      <div className="sb-top-name">{p.name}</div>
                      <div className="sb-top-bar-wrap">
                        <div className="sb-top-bar" style={{width:`${p.pct}%`}}/>
                      </div>
                    </div>
                    <span className="sb-top-count">{p.count}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ─── COACHING SIDEBAR ────────────────────────────── */}
          {tab==='coaching' && isMgr && (
            <div className="sb-panel">
              <div className="sb-section-hd">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>
                {rtl?'المديرون':'Managers'}
              </div>
              {selectedManager && (
                <button className="sb-back" onClick={()=>setSelectedManager(null)}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>
                  {rtl?'عرض الكل':'Show all'}
                </button>
              )}
              {managerGroups.map(mgr => (
                <div key={mgr.name}
                  className={`sb-mgr-card${selectedManager===mgr.name?' active':''}`}
                  onClick={()=>setSelectedManager(selectedManager===mgr.name?null:mgr.name)}>
                  <div className="sb-mgr-avatar">
                    {mgr.name.split(' ').map(w=>w[0]).slice(0,2).join('')}
                  </div>
                  <div className="sb-mgr-info">
                    <div className="sb-mgr-name">{mgr.name}</div>
                    <div className="sb-mgr-meta">{mgr.team}</div>
                  </div>
                  <div className="sb-mgr-stats">
                    <div className="sb-mgr-stat">{mgr.dayCount}<small> days</small></div>
                    <div className="sb-mgr-stat">{mgr.repCount}<small> reps</small></div>
                  </div>
                </div>
              ))}
            </div>
            )}
          </aside>
        )}

        {/* Mobile backdrop */}
        {sidebarOpen && <div className="sb-backdrop" onClick={()=>setSidebarOpen(false)}/>}

        {/* ── MAIN CONTENT ───────────────────────────────────── */}
        <div className="dash-content">

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
                  <select className="ctrl-sel" value={userFilter} onChange={e=>{
                    const val = e.target.value;
                    setUser(val);
                    setSelectedRep(val==='all'?null:val);
                  }}>
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
            {tab!=='roadmap' && (
              <button className="sidebar-toggle" onClick={()=>setSidebarOpen(!sidebarOpen)} title="Toggle panel">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18"/>
                </svg>
              </button>
            )}
            {visibleTabs.map(([k,label])=>(
              <button key={k} className={`dtab${tab===k?' on':''}`} onClick={()=>changeTab(k)}>{label}</button>
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
                    {isMgr&&userFilter==='all'&&(
                      <div className="agg-cards-row">
                        {teamGroups.map(({label,rows})=>(
                          <TeamKpiHover key={label} rows={rows} teamLabel={label}/>
                        ))}
                      </div>
                    )}
                    <div className="cards-grid">
                      {fSummary.map((r,i)=>(
                        <div key={r.id||i} className={`ucard${r.is_manager?' mgr':''}${selectedRep===r.user_name?' ucard-selected':''}`}
                          onClick={()=>handleSelectRep(r.user_name)}>
                          <div className="ucard-hdr">
                            <div className="ucard-info">
                              <div className="ucard-name">{r.user_name}</div>
                              <div className="ucard-meta">{r.team||''}{r.is_manager?' · Manager':''}</div>
                              {r.territory&&<div className="ucard-terr" title={r.territory}>{r.territory}</div>}
                              {(r.avg_am_shift_hm||r.avg_pm_shift_hm)&&(
                                <div className="ucard-dur">
                                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>
                                  {r.avg_am_shift_hm?<span className="dur-am">AM {fmtDuration(r.avg_am_shift_hm)}</span>:null}
                                  {r.avg_pm_shift_hm?<span className="dur-pm">PM {fmtDuration(r.avg_pm_shift_hm)}</span>:null}
                                </div>
                              )}
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
                              <div key={g.label} className={`kpi-sec${g.keys.includes('avg_am_shift_hm')?' kpi-timing':''}`}>
                                <div className="kpi-sec-hd">{g.label}</div>
                                {kpiRows.map(({k,v})=>{
                                  const target = KPI_TARGETS[k];
                                  const avgVal = companyAverages[k] || 0;
                                  const numVal = Number(v) || 0;
                                  let benchmarkClass = '';
                                  if (avgVal > 0 && typeof numVal === 'number' && !k.includes('time') && !k.includes('hm')) {
                                    if (numVal > avgVal * 1.05) benchmarkClass = 'above-avg';
                                    else if (numVal < avgVal * 0.95) benchmarkClass = 'below-avg';
                                  }
                                  const pct = target ? Math.min(100, Math.round((numVal / target) * 100)) : null;
                                  return (
                                    <div key={k} className="kpi-row-wrapper">
                                      <div className="kpi-row">
                                        <span className="kpi-lbl">
                                          {t.kpi[k]||k}
                                          {benchmarkClass === 'above-avg' && <span className="bench-arrow up" title="Above company average">▲</span>}
                                          {benchmarkClass === 'below-avg' && <span className="bench-arrow down" title="Below company average">▼</span>}
                                        </span>
                                        <span className={`kpi-v ${k.includes('rate')?'rate':''} ${benchmarkClass}`}>{fmtVal(v,k)}</span>
                                      </div>
                                      {pct !== null && (
                                        <div className="kpi-card-progress" title={`${pct}% of target (${target})`}>
                                          <div className="kpi-card-progress-bar" style={{ width: `${pct}%`, backgroundColor: pct >= 100 ? '#10b981' : pct >= 70 ? '#3b82f6' : '#ef4444' }} />
                                        </div>
                                      )}
                                    </div>
                                  );
                                })}
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
                  <PivotSummaryBanner 
                    rows={filteredSpecialty} 
                    valueKey="call_count" 
                    rowKey="specialty" 
                    shift={shift} 
                    t={t}
                    selectedTeam={team}
                    onSelectTeam={setTeam}
                    userTeamMap={userTeamMap}
                  />
                  <PivotTable rows={filteredSpecialty} rowKey="specialty" valueKey="call_count"
                    shiftFilter={shift} userFilter={userFilter} searchFilter={search} lang={lang}/>
                </>
              )}

              {/* PRODUCTS TAB */}
              {tab==='products'&&(
                <>
                  <PivotSummaryBanner 
                    rows={filteredProducts} 
                    valueKey="call_count" 
                    rowKey="product" 
                    shift={shift} 
                    t={t}
                    selectedTeam={team}
                    onSelectTeam={setTeam}
                    userTeamMap={userTeamMap}
                  />
                  <PivotTable rows={filteredProducts} rowKey="product" valueKey="call_count"
                    shiftFilter={shift} userFilter={userFilter} searchFilter={search} lang={lang}/>
                </>
              )}

              {/* COACHING TAB */}
              {tab==='coaching'&&isMgr&&(
                filteredCoaching.length===0?(
                  <div className="dash-empty">
                    {selectedManager
                      ? (rtl?'لا توجد بيانات لهذا المدير':'No coaching data for this manager')
                      : t.noData}
                  </div>
                ):(
                  <>
                    {selectedManager && (
                      <div className="coaching-selected-hdr">
                        <span className="coaching-sel-name">{selectedManager}</span>
                        <span className="coaching-sel-meta">
                          {filteredCoaching.length} {rtl?'جلسة':'session'}{filteredCoaching.length!==1?'s':''}
                        </span>
                      </div>
                    )}
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
                          {[...filteredCoaching].sort((a,b)=>(a.manager_name||'').localeCompare(b.manager_name||'')||(a.coaching_date||'').localeCompare(b.coaching_date||'')).map((r,i)=>(
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
                  </>
                )
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
