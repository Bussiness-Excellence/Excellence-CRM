import React, { useEffect, useMemo, useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../supabaseClient';
import { useAuth } from '../contexts/AuthContext';
import './Dashboard.css';

// ── i18n ─────────────────────────────────────────────────────────────────────
const T = {
  en: {
    signOut: 'Sign out', period: 'Period', team: 'Team',
    allTeams: 'All teams', user: 'User', download: 'Download Excel',
    loading: 'Loading…', noData: 'No data for this period yet.',
    people: (n) => `${n} ${n === 1 ? 'person' : 'people'}`,
    acrossTeams: (n) => ` across ${n} teams`,
    tabs: { summary: 'Summary', specialty: 'Specialty × Class', products: 'Product Calls', coaching: 'Coaching Days' },
    roles: { MR: 'My', Supervisor: "My team's", 'Area Manager': "My area's", BLM: "My team's", Admin: 'All teams' },
    appTitle: 'Excellence-CRM web app',
    cols: {
      team: 'Team', user_name: 'User', territory: 'Territory', is_manager: 'Manager',
      working_days: 'Working Days', complete_field_days: 'Field Days',
      am_shift_days: 'AM Days', pm_shift_days: 'PM Days',
      double_visit_days: 'Double Visits', coaching_days: 'Coaching Days',
      office_work_days: 'Office Work', no_events: 'Events',
      total_am_covered: 'AM Covered', total_pm_covered: 'PM Covered',
      clinic_covered: 'Clinic', polyclinic_covered: 'Poly Clinic',
      amcenter_covered: 'AM Center', hospital_covered: 'Hospital',
      am_calls: 'AM Calls', am_call_rate: 'AM Call Rate',
      pm_calls: 'PM Calls', pm_call_rate: 'PM Call Rate',
      avg_am_start_time: 'AM Start', pharmacies_visited: 'Pharm. Visited',
      pharmacies_covered: 'Pharm. Covered', total_product_calls: 'Product Calls',
      distinct_products: 'Products', avg_am_shift_hm: 'AM Duration',
      avg_pm_shift_hm: 'PM Duration', avg_field_overall_hm: 'Field Duration',
      // specialty tab
      specialty: 'Specialty', classification: 'Class', shift: 'Shift',
      call_count: 'Calls', unique_doctors: 'Unique Drs',
      // product tab
      product: 'Product',
      // coaching tab
      manager_name: 'Manager', rep_name: 'Rep', coaching_date: 'Date',
    },
  },
  ar: {
    signOut: 'تسجيل خروج', period: 'الفترة', team: 'الفريق',
    allTeams: 'كل الفرق', user: 'المستخدم', download: 'تنزيل Excel',
    loading: 'جاري التحميل…', noData: 'لا توجد بيانات لهذه الفترة بعد.',
    people: (n) => `${n} ${n === 1 ? 'شخص' : 'أشخاص'}`,
    acrossTeams: (n) => ` في ${n} فرق`,
    tabs: { summary: 'الملخص', specialty: 'التخصص × الفئة', products: 'المنتجات', coaching: 'التوجيه' },
    roles: { MR: 'بياناتي', Supervisor: 'بيانات فريقي', 'Area Manager': 'بيانات منطقتي', BLM: 'بيانات فريقي', Admin: 'كل الفرق' },
    appTitle: 'تطبيق Excellence-CRM',
    cols: {
      team: 'الفريق', user_name: 'المستخدم', territory: 'المنطقة', is_manager: 'مدير',
      working_days: 'أيام العمل', complete_field_days: 'أيام الميدان الكاملة',
      am_shift_days: 'أيام AM', pm_shift_days: 'أيام PM',
      double_visit_days: 'زيارات مزدوجة', coaching_days: 'أيام التوجيه',
      office_work_days: 'عمل مكتبي', no_events: 'فعاليات',
      total_am_covered: 'تغطية AM', total_pm_covered: 'تغطية PM',
      clinic_covered: 'عيادات', polyclinic_covered: 'مراكز صحية',
      amcenter_covered: 'مراكز AM', hospital_covered: 'مستشفيات',
      am_calls: 'زيارات AM', am_call_rate: 'معدل AM',
      pm_calls: 'زيارات PM', pm_call_rate: 'معدل PM',
      avg_am_start_time: 'بدء AM', pharmacies_visited: 'صيدليات (زيارات)',
      pharmacies_covered: 'صيدليات (تغطية)', total_product_calls: 'مكالمات المنتج',
      distinct_products: 'منتجات', avg_am_shift_hm: 'مدة AM',
      avg_pm_shift_hm: 'مدة PM', avg_field_overall_hm: 'مدة الميدان',
      specialty: 'التخصص', classification: 'الفئة', shift: 'الفترة',
      call_count: 'الزيارات', unique_doctors: 'أطباء فريدون',
      product: 'المنتج',
      manager_name: 'المدير', rep_name: 'المندوب', coaching_date: 'التاريخ',
    },
  },
};

// ── column definitions per tab ────────────────────────────────────────────────
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

// ── helpers ────────────────────────────────────────────────────────────────────
function fmt(val, key) {
  if (val === null || val === undefined || val === '') return '—';
  if (key === 'is_manager') return val ? '✓' : '';
  if (key === 'am_call_rate' || key === 'pm_call_rate') return Number(val).toFixed(1);
  return val;
}

function sortRows(rows, managers) {
  // Within each team: reps first (alphabetical), then managers
  const managerSet = new Set(managers);
  return [...rows].sort((a, b) => {
    const teamCmp = (a.team || '').localeCompare(b.team || '');
    if (teamCmp !== 0) return teamCmp;
    const aMgr = managerSet.has(a.user_name) ? 1 : 0;
    const bMgr = managerSet.has(b.user_name) ? 1 : 0;
    if (aMgr !== bMgr) return aMgr - bMgr;
    return (a.user_name || '').localeCompare(b.user_name || '');
  });
}

function downloadExcel(tabs, lang) {
  const wb = XLSX.utils.book_new();
  tabs.forEach(({ name, cols, rows, t }) => {
    const headers = cols.map(c => t.cols[c] || c);
    const data = rows.map(r => cols.map(c => {
      const v = r[c];
      if (v === null || v === undefined || v === '') return '';
      if (c === 'is_manager') return v ? 'Yes' : '';
      return v;
    }));
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  XLSX.writeFile(wb, `excellence_${Date.now()}.xlsx`);
}

// ── main component ─────────────────────────────────────────────────────────────
export default function Dashboard() {
  const { profile, visibleCodes, signOut, hierarchy } = useAuth();
  const [lang, setLang]           = useState(profile?.preferred_lang || 'en');
  const [periods, setPeriods]     = useState([]);
  const [selectedPeriod, setSelectedPeriod] = useState('');
  const [selectedTeam, setSelectedTeam]     = useState('all');
  const [activeTab, setActiveTab] = useState('summary');
  const [summaryRows, setSummaryRows]       = useState([]);
  const [specialtyRows, setSpecialtyRows]   = useState([]);
  const [productRows, setProductRows]       = useState([]);
  const [coachingRows, setCoachingRows]     = useState([]);
  const [loading, setLoading]     = useState(true);
  const [error, setError]         = useState('');

  const t = T[lang] || T.en;
  const isRtl = lang === 'ar';

  // Manager names set (for sorting)
  const managerNames = useMemo(
    () => new Set(summaryRows.filter(r => r.is_manager).map(r => r.user_name)),
    [summaryRows]
  );

  // Available teams from summary rows
  const teams = useMemo(
    () => ['all', ...new Set(summaryRows.map(r => r.team).filter(Boolean))].sort(),
    [summaryRows]
  );

  // Load available periods
  useEffect(() => {
    supabase.from('summaries').select('period').then(({ data }) => {
      const unique = [...new Set((data || []).map(r => r.period))].filter(Boolean).sort().reverse();
      setPeriods(unique);
      if (unique.length) setSelectedPeriod(unique[0]);
    });
  }, []);

  // Load all data for selected period
  const loadData = useCallback(async () => {
    if (!selectedPeriod || !visibleCodes.length) {
      setSummaryRows([]); setSpecialtyRows([]); setProductRows([]); setCoachingRows([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    setError('');

    const [s, sp, pr, co] = await Promise.all([
      supabase.from('summaries').select('*')
        .eq('period', selectedPeriod).in('employee_code', visibleCodes),
      supabase.from('specialty_classification').select('*')
        .eq('period', selectedPeriod).in('employee_code', visibleCodes),
      supabase.from('product_calls').select('*')
        .eq('period', selectedPeriod).in('employee_code', visibleCodes),
      supabase.from('coaching_days').select('*')
        .eq('period', selectedPeriod)
        .or(`manager_code.in.(${visibleCodes.join(',')}),rep_code.in.(${visibleCodes.join(',')})`),
    ]);

    if (s.error) setError(s.error.message);
    // Enrich summary with team from hierarchy
    const hierarchyMap = Object.fromEntries((hierarchy || []).map(h => [h.employee_code, h]));
    const enriched = (s.data || []).map(r => ({
      ...r,
      team: r.team || hierarchyMap[r.employee_code]?.division_name || '',
    }));
    setSummaryRows(enriched);

    // Enrich specialty/product rows with team name
    const teamByCode = Object.fromEntries(enriched.map(r => [r.employee_code, r.team]));
    setSpecialtyRows((sp.data || []).map(r => ({ ...r, team: teamByCode[r.employee_code] || r.team || '' })));
    setProductRows((pr.data || []).map(r => ({ ...r, team: teamByCode[r.employee_code] || r.team || '' })));
    setCoachingRows(co.data || []);
    setLoading(false);
  }, [selectedPeriod, visibleCodes, hierarchy]);

  useEffect(() => { loadData(); }, [loadData]);

  // Filter by team
  const filtered = useMemo(() => {
    const filter = (rows) => selectedTeam === 'all' ? rows : rows.filter(r => r.team === selectedTeam);
    return {
      summary:   sortRows(filter(summaryRows), managerNames),
      specialty: filter(specialtyRows).sort((a,b) => (a.user_name||'').localeCompare(b.user_name||'')),
      products:  filter(productRows).sort((a,b) => (a.user_name||'').localeCompare(b.user_name||'')),
      coaching:  selectedTeam === 'all' ? coachingRows : coachingRows.filter(r => r.team === selectedTeam),
    };
  }, [summaryRows, specialtyRows, productRows, coachingRows, selectedTeam, managerNames]);

  const activeRows = filtered[activeTab === 'summary' ? 'summary' : activeTab === 'specialty' ? 'specialty' : activeTab === 'products' ? 'products' : 'coaching'];
  const activeCols = activeTab === 'summary' ? SUMMARY_COLS : activeTab === 'specialty' ? SPECIALTY_COLS : activeTab === 'products' ? PRODUCT_COLS : COACHING_COLS;

  function handleDownload() {
    downloadExcel([
      { name: t.tabs.summary,   cols: SUMMARY_COLS,   rows: filtered.summary,   t },
      { name: t.tabs.specialty, cols: SPECIALTY_COLS, rows: filtered.specialty, t },
      { name: t.tabs.products,  cols: PRODUCT_COLS,   rows: filtered.products,  t },
      { name: t.tabs.coaching,  cols: COACHING_COLS,  rows: filtered.coaching,  t },
    ], lang);
  }

  return (
    <div className={`dashboard${isRtl ? ' rtl' : ''}`} dir={isRtl ? 'rtl' : 'ltr'}>
      {/* Header */}
      <header className="dashboard-header">
        <div className="dashboard-brand">
          <div className="header-logos-pill">
            <img src="/eipico-logo.png" alt="EIPICO" className="header-logo header-logo-company" />
            <div className="header-logos-divider" />
            <img src="/dept-logo.png" alt="Excellence Department" className="header-logo header-logo-dept" />
          </div>
          <div className="dashboard-brand-text">
            <div className="dashboard-app-title">{t.appTitle}</div>
            <div className="dashboard-role-label">{(t.roles[profile?.role] || '')} {isRtl ? '' : 'view'}</div>
          </div>
        </div>
        <div className="dashboard-header-right">
          <button className="lang-toggle" onClick={() => setLang(lang === 'en' ? 'ar' : 'en')}>
            {lang === 'en' ? 'عربي' : 'EN'}
          </button>
          <div className="dashboard-user">
            <div className="dashboard-user-name">{profile?.employee_name}</div>
            <div className="dashboard-user-role">{profile?.role} · {profile?.employee_code}</div>
          </div>
          <button className="signout-button" onClick={signOut}>{t.signOut}</button>
        </div>
      </header>

      {/* Toolbar */}
      <div className="dashboard-toolbar">
        <label>{t.period}</label>
        <select value={selectedPeriod} onChange={e => setSelectedPeriod(e.target.value)}>
          {periods.map(p => <option key={p} value={p}>{p}</option>)}
        </select>

        {teams.length > 2 && (
          <>
            <label>{t.team}</label>
            <select value={selectedTeam} onChange={e => setSelectedTeam(e.target.value)}>
              <option value="all">{t.allTeams}</option>
              {teams.filter(t => t !== 'all').map(tm => <option key={tm} value={tm}>{tm}</option>)}
            </select>
          </>
        )}

        <div className="dashboard-stats">
          {t.people(filtered.summary.length)}
          {new Set(filtered.summary.map(r=>r.team)).size > 1 && t.acrossTeams(new Set(filtered.summary.map(r=>r.team)).size)}
        </div>

        <button className="download-button" onClick={handleDownload}>⬇ {t.download}</button>
      </div>

      {/* Tabs */}
      <div className="tabs">
        {Object.entries(t.tabs).map(([key, label]) => (
          <button
            key={key}
            className={`tab${activeTab === key ? ' tab-active' : ''}`}
            onClick={() => setActiveTab(key)}
          >{label}</button>
        ))}
      </div>

      {error && <div className="dashboard-error">{error}</div>}

      {/* Table */}
      {loading ? (
        <div className="centered-message">{t.loading}</div>
      ) : activeRows.length === 0 ? (
        <div className="centered-message">{t.noData}</div>
      ) : (
        <div className="table-wrap">
          <table className="summary-table">
            <thead>
              <tr>
                {activeCols.map(c => <th key={c}>{t.cols[c] || c}</th>)}
              </tr>
            </thead>
            <tbody>
              {activeRows.map((r, i) => (
                <tr key={r.id || i} className={r.is_manager ? 'row-manager' : ''}>
                  {activeCols.map(c => <td key={c}>{fmt(r[c], c)}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
