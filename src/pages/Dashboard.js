import React, { useEffect, useMemo, useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../supabaseClient';
import { useAuth } from '../contexts/AuthContext';
import './Dashboard.css';

// ── i18n ──────────────────────────────────────────────────────────────────────
const T = {
  en: {
    brand: 'EXCELLENCE', signOut: 'Sign out',
    periodLastMonth: 'Last Month', periodRecent: 'Recent (1–15)',
    team: 'Team', allTeams: 'All teams', download: '⬇ Download Excel',
    loading: 'Loading…', noData: 'No data for this period yet.',
    people: (n) => `${n} ${n===1?'person':'people'}`,
    acrossTeams: (n) => ` across ${n} teams`,
    tabs: { summary:'Summary', specialty:'Specialty × Class', products:'Product Calls', coaching:'Coaching Days' },
    roles: { MR:'My', Supervisor:"My team's", 'Area Manager':"My area's", BLM:"My team's", Admin:'All teams' },
    cols: {
      team:'Team', user_name:'User', territory:'Territory', is_manager:'Mgr',
      working_days:'Working Days', complete_field_days:'Field Days',
      am_shift_days:'AM Days', pm_shift_days:'PM Days',
      double_visit_days:'Double', coaching_days:'Coaching', office_work_days:'Office',
      no_events:'Events', total_am_covered:'AM Covered', total_pm_covered:'PM Covered',
      clinic_covered:'Clinic', polyclinic_covered:'Poly Clinic',
      amcenter_covered:'AM Center', hospital_covered:'Hospital',
      am_calls:'AM Calls', am_call_rate:'AM Rate',
      pm_calls:'PM Calls', pm_call_rate:'PM Rate',
      avg_am_start_time:'AM Start', pharmacies_visited:'Pharm.Visits',
      pharmacies_covered:'Pharm.Covered', total_product_calls:'Prod.Calls',
      distinct_products:'Products', avg_am_shift_hm:'AM Dur.',
      avg_pm_shift_hm:'PM Dur.', avg_field_overall_hm:'Field Dur.',
      specialty:'Specialty', classification:'Class', shift:'Shift',
      call_count:'Calls', unique_doctors:'Unique Drs',
      product:'Product',
      manager_name:'Manager', rep_name:'Rep', coaching_date:'Date',
    },
  },
  ar: {
    brand:'إكسيلنس', signOut:'خروج',
    periodLastMonth:'الشهر الماضي', periodRecent:'الأحدث (1–15)',
    team:'الفريق', allTeams:'كل الفرق', download:'⬇ تنزيل Excel',
    loading:'جاري التحميل…', noData:'لا توجد بيانات لهذه الفترة.',
    people:(n)=>`${n} ${n===1?'شخص':'أشخاص'}`,
    acrossTeams:(n)=>` في ${n} فرق`,
    tabs:{ summary:'الملخص', specialty:'التخصص × الفئة', products:'المنتجات', coaching:'التوجيه' },
    roles:{ MR:'بياناتي', Supervisor:'فريقي', 'Area Manager':'منطقتي', BLM:'فريقي', Admin:'كل الفرق' },
    cols:{
      team:'الفريق', user_name:'المستخدم', territory:'المنطقة', is_manager:'مدير',
      working_days:'أيام العمل', complete_field_days:'أيام الميدان',
      am_shift_days:'أيام AM', pm_shift_days:'أيام PM',
      double_visit_days:'مزدوج', coaching_days:'توجيه', office_work_days:'مكتب',
      no_events:'فعاليات', total_am_covered:'تغطية AM', total_pm_covered:'تغطية PM',
      clinic_covered:'عيادات', polyclinic_covered:'مراكز', amcenter_covered:'AM Center', hospital_covered:'مستشفيات',
      am_calls:'زيارات AM', am_call_rate:'معدل AM', pm_calls:'زيارات PM', pm_call_rate:'معدل PM',
      avg_am_start_time:'بدء AM', pharmacies_visited:'صيدليات (زيارات)', pharmacies_covered:'صيدليات (تغطية)',
      total_product_calls:'مكالمات', distinct_products:'منتجات',
      avg_am_shift_hm:'مدة AM', avg_pm_shift_hm:'مدة PM', avg_field_overall_hm:'مدة الميدان',
      specialty:'التخصص', classification:'الفئة', shift:'الفترة', call_count:'الزيارات', unique_doctors:'أطباء',
      product:'المنتج', manager_name:'المدير', rep_name:'المندوب', coaching_date:'التاريخ',
    },
  },
};

const SUMMARY_COLS = [
  'team','user_name','territory','is_manager',
  'working_days','complete_field_days',
  'am_shift_days','pm_shift_days','double_visit_days','coaching_days','office_work_days','no_events',
  'total_am_covered','amcenter_covered','hospital_covered',
  'total_pm_covered','clinic_covered','polyclinic_covered',
  'am_calls','am_call_rate','pm_calls','pm_call_rate','avg_am_start_time',
  'pharmacies_visited','pharmacies_covered',
  'total_product_calls','distinct_products',
  'avg_am_shift_hm','avg_pm_shift_hm','avg_field_overall_hm',
];
const SPECIALTY_COLS = ['team','user_name','specialty','classification','shift','call_count','unique_doctors'];
const PRODUCT_COLS   = ['team','user_name','product','shift','specialty','call_count','unique_doctors'];
const COACHING_COLS  = ['team','manager_name','rep_name','coaching_date'];

const PERIODS = [
  { key: 'last_month', en: 'Last Month',    ar: 'الشهر الماضي' },
  { key: 'recent',     en: 'Recent (1–15)', ar: 'الأحدث (1–15)' },
];

function fmt(val, key) {
  if (val === null || val === undefined || val === '') return '—';
  if (key === 'is_manager') return val ? '✓' : '';
  if (key === 'am_call_rate' || key === 'pm_call_rate') return Number(val).toFixed(1);
  return val;
}

function sortSummary(rows) {
  return [...rows].sort((a, b) => {
    const tc = (a.team||'').localeCompare(b.team||'');
    if (tc !== 0) return tc;
    // managers last within team
    const am = a.is_manager ? 1 : 0, bm = b.is_manager ? 1 : 0;
    if (am !== bm) return am - bm;
    return (a.user_name||'').localeCompare(b.user_name||'');
  });
}

function downloadExcel(tabs, lang) {
  const wb = XLSX.utils.book_new();
  const t = T[lang];
  tabs.forEach(({ name, cols, rows }) => {
    const headers = cols.map(c => t.cols[c] || c);
    const data = rows.map(r => cols.map(c => {
      const v = r[c];
      if (v === null || v === undefined || v === '') return '';
      if (c === 'is_manager') return v ? 'Yes' : '';
      return v;
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...data]), name);
  });
  XLSX.writeFile(wb, `excellence_report_${Date.now()}.xlsx`);
}

export default function Dashboard() {
  const { profile, visibleCodes, signOut } = useAuth();
  const [lang, setLang]         = useState(profile?.preferred_lang || 'en');
  const [period, setPeriod]     = useState('last_month');
  const [selectedTeam, setTeam] = useState('all');
  const [activeTab, setTab]     = useState('summary');
  const [summaryRows, setSummary]   = useState([]);
  const [specialtyRows, setSpec]    = useState([]);
  const [productRows, setProd]      = useState([]);
  const [coachingRows, setCoaching] = useState([]);
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState('');

  const t   = T[lang] || T.en;
  const rtl = lang === 'ar';

  // Load data when period or visible codes change
  const loadData = useCallback(async () => {
    if (!visibleCodes.length) { setLoading(false); return; }
    setLoading(true); setError('');

    // Map period key to label stored in DB
    const periodLabel = period === 'last_month' ? 'Last Month' : 'Recent';

    const codeStr = visibleCodes.join(',');
    const [s, sp, pr, co] = await Promise.all([
      supabase.from('summaries').select('*')
        .eq('period', periodLabel).in('employee_code', visibleCodes),
      supabase.from('specialty_classification').select('*')
        .eq('period', periodLabel).in('employee_code', visibleCodes),
      supabase.from('product_calls').select('*')
        .eq('period', periodLabel).in('employee_code', visibleCodes),
      supabase.from('coaching_days').select('*').eq('period', periodLabel)
        .or(visibleCodes.map(c=>`manager_code.eq.${c}`).concat(visibleCodes.map(c=>`rep_code.eq.${c}`)).join(',')),
    ]);

    if (s.error) setError(s.error.message);
    setSummary(s.data || []);
    setSpec(sp.data || []);
    setProd(pr.data || []);
    setCoaching(co.data || []);
    setLoading(false);
  }, [period, visibleCodes]);

  useEffect(() => { loadData(); }, [loadData]);

  // Teams list from summary rows
  const teams = useMemo(() =>
    [...new Set(summaryRows.map(r => r.team).filter(Boolean))].sort(),
    [summaryRows]
  );

  // Filter by team
  const filter = useCallback((rows) =>
    selectedTeam === 'all' ? rows : rows.filter(r => r.team === selectedTeam),
    [selectedTeam]
  );

  const filtered = useMemo(() => ({
    summary:   sortSummary(filter(summaryRows)),
    specialty: filter(specialtyRows),
    products:  filter(productRows),
    coaching:  filter(coachingRows),
  }), [summaryRows, specialtyRows, productRows, coachingRows, filter]);

  const activeRows = filtered[
    activeTab === 'summary' ? 'summary' :
    activeTab === 'specialty' ? 'specialty' :
    activeTab === 'products' ? 'products' : 'coaching'
  ];
  const activeCols = activeTab === 'summary' ? SUMMARY_COLS :
    activeTab === 'specialty' ? SPECIALTY_COLS :
    activeTab === 'products' ? PRODUCT_COLS : COACHING_COLS;

  const teamCount = new Set(filtered.summary.map(r=>r.team)).size;

  return (
    <div className={`dashboard${rtl?' rtl':''}`} dir={rtl?'rtl':'ltr'}>
      {/* Header */}
      <header className="dash-header">
        <div className="dash-header-left">
          <div className="dash-logo">
            <img src="/logo192.png" alt="logo" className="dash-logo-img" onError={e=>e.target.style.display='none'}/>
          </div>
          <div>
            <div className="dash-brand">{t.brand}</div>
            <div className="dash-role">{t.roles[profile?.role] || ''} VIEW</div>
          </div>
        </div>
        <div className="dash-header-right">
          <button className="btn-lang" onClick={()=>setLang(lang==='en'?'ar':'en')}>
            {lang==='en'?'عربي':'EN'}
          </button>
          <div className="dash-user">
            <div className="dash-user-name">{profile?.employee_name}</div>
            <div className="dash-user-role">{profile?.role} · {profile?.employee_code}</div>
          </div>
          <button className="btn-signout" onClick={signOut}>{t.signOut}</button>
        </div>
      </header>

      {/* Toolbar */}
      <div className="dash-toolbar">
        {/* Period toggle — just two buttons */}
        <div className="period-toggle">
          {PERIODS.map(p => (
            <button
              key={p.key}
              className={`period-btn${period===p.key?' active':''}`}
              onClick={()=>setPeriod(p.key)}
            >{lang==='ar'?p.ar:p.en}</button>
          ))}
        </div>

        {/* Team filter */}
        {teams.length > 1 && (
          <div className="team-filter">
            <label>{t.team}</label>
            <select value={selectedTeam} onChange={e=>setTeam(e.target.value)}>
              <option value="all">{t.allTeams}</option>
              {teams.map(tm=><option key={tm} value={tm}>{tm}</option>)}
            </select>
          </div>
        )}

        <div className="dash-stats">
          {t.people(filtered.summary.length)}
          {teamCount > 1 && t.acrossTeams(teamCount)}
        </div>

        <button className="btn-download" onClick={()=>downloadExcel([
          {name:t.tabs.summary,   cols:SUMMARY_COLS,   rows:filtered.summary},
          {name:t.tabs.specialty, cols:SPECIALTY_COLS, rows:filtered.specialty},
          {name:t.tabs.products,  cols:PRODUCT_COLS,   rows:filtered.products},
          {name:t.tabs.coaching,  cols:COACHING_COLS,  rows:filtered.coaching},
        ], lang)}>{t.download}</button>
      </div>

      {/* Tabs */}
      <div className="dash-tabs">
        {Object.entries(t.tabs).map(([key,label])=>(
          <button key={key}
            className={`dash-tab${activeTab===key?' active':''}`}
            onClick={()=>setTab(key)}>{label}
          </button>
        ))}
      </div>

      {error && <div className="dash-error">{error}</div>}

      {/* Table */}
      {loading ? (
        <div className="dash-empty">{t.loading}</div>
      ) : activeRows.length === 0 ? (
        <div className="dash-empty">{t.noData}</div>
      ) : (
        <div className="table-wrap">
          <table className="data-table">
            <thead>
              <tr>{activeCols.map(c=><th key={c}>{t.cols[c]||c}</th>)}</tr>
            </thead>
            <tbody>
              {activeRows.map((r,i)=>(
                <tr key={r.id||i} className={r.is_manager?'row-mgr':''}>
                  {activeCols.map(c=><td key={c}>{fmt(r[c],c)}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
