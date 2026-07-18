import React, { useEffect, useMemo, useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../supabaseClient';
import { useAuth } from '../contexts/AuthContext';
import './Dashboard.css';

// ── i18n ───────────────────────────────────────────────────────────────────
const T = {
  en: {
    appName: 'EXCELLENCE', signOut: 'Sign out', admin: 'Admin Panel',
    periodLastMonth: 'Last Month', periodRecent: 'Recent  1–15',
    allTeams: 'All teams', download: 'Export', loading: 'Loading…',
    noData: 'No data available for this period.',
    people: n => `${n} ${n===1?'rep':'reps'}`,
    acrossTeams: n => ` · ${n} teams`,
    tabs: { summary:'Summary', specialty:'Specialty', products:'Products', coaching:'Coaching' },
    roleView: { MR:'My Results', Supervisor:'My Team', 'Area Manager':'My Area', BLM:'Full Team', Admin:'All Teams' },
    kpi: {
      working_days:'Working Days', complete_field_days:'Field Days',
      am_shift_days:'AM Days', pm_shift_days:'PM Days',
      am_calls:'AM Calls', pm_calls:'PM Calls',
      am_call_rate:'AM Rate', pm_call_rate:'PM Rate',
      total_am_covered:'AM Covered', total_pm_covered:'PM Covered',
      double_visit_days:'Double Visits', coaching_days:'Coaching Days',
      office_work_days:'Office Work', pharmacies_visited:'Pharm. Visits',
      pharmacies_covered:'Pharm. Covered', total_product_calls:'Product Calls',
      distinct_products:'Products', avg_am_start_time:'AM Start',
      avg_am_shift_hm:'AM Duration', avg_pm_shift_hm:'PM Duration',
    },
  },
  ar: {
    appName: 'إكسيلنس', signOut: 'خروج', admin: 'لوحة الإدارة',
    periodLastMonth: 'الشهر الماضي', periodRecent: 'الأحدث  1–15',
    allTeams: 'كل الفرق', download: 'تصدير', loading: 'جارٍ التحميل…',
    noData: 'لا توجد بيانات لهذه الفترة.',
    people: n => `${n} مندوب`,
    acrossTeams: n => ` · ${n} فرق`,
    tabs: { summary:'الملخص', specialty:'التخصص', products:'المنتجات', coaching:'التوجيه' },
    roleView: { MR:'نتائجي', Supervisor:'فريقي', 'Area Manager':'منطقتي', BLM:'الفريق كاملاً', Admin:'كل الفرق' },
    kpi: {
      working_days:'أيام العمل', complete_field_days:'أيام الميدان',
      am_shift_days:'أيام AM', pm_shift_days:'أيام PM',
      am_calls:'زيارات AM', pm_calls:'زيارات PM',
      am_call_rate:'معدل AM', pm_call_rate:'معدل PM',
      total_am_covered:'تغطية AM', total_pm_covered:'تغطية PM',
      double_visit_days:'زيارات مزدوجة', coaching_days:'أيام التوجيه',
      office_work_days:'مكتب', pharmacies_visited:'صيدليات (زيارات)',
      pharmacies_covered:'صيدليات (تغطية)', total_product_calls:'مكالمات المنتج',
      distinct_products:'منتجات', avg_am_start_time:'بدء AM',
      avg_am_shift_hm:'مدة AM', avg_pm_shift_hm:'مدة PM',
    },
  },
};

// KPI groups for vertical card layout
const KPI_GROUPS = [
  { label: { en: 'Field Activity', ar: 'النشاط الميداني' }, keys: ['working_days','complete_field_days','am_shift_days','pm_shift_days','double_visit_days','office_work_days'] },
  { label: { en: 'Calls', ar: 'الزيارات' }, keys: ['am_calls','am_call_rate','pm_calls','pm_call_rate'] },
  { label: { en: 'Coverage', ar: 'التغطية' }, keys: ['total_am_covered','total_pm_covered','pharmacies_visited','pharmacies_covered'] },
  { label: { en: 'Products', ar: 'المنتجات' }, keys: ['total_product_calls','distinct_products'] },
  { label: { en: 'Coaching', ar: 'التوجيه' }, keys: ['coaching_days'] },
  { label: { en: 'Timing', ar: 'التوقيت' }, keys: ['avg_am_start_time','avg_am_shift_hm','avg_pm_shift_hm'] },
];

function fmtVal(v, key) {
  if (v === null || v === undefined || v === '') return '—';
  if (key?.includes('rate')) return Number(v).toFixed(1);
  return v;
}

function sortSummary(rows) {
  return [...rows].sort((a,b) => {
    const tc = (a.team||'').localeCompare(b.team||'');
    if (tc) return tc;
    if (a.is_manager !== b.is_manager) return a.is_manager ? 1 : -1;
    return (a.user_name||'').localeCompare(b.user_name||'');
  });
}

// Pivot a list of detail rows into a table:
// rows[user][rowKey] = value
// Returns { users[], rowKeys[], cells{user}{rowKey} }
function pivotRows(rows, userKey, rowKey, valueKey, secondaryKey) {
  const users = [...new Set(rows.map(r => r[userKey]))].sort();
  const keys = [...new Set(rows.map(r => secondaryKey ? `${r[rowKey]} (${r[secondaryKey]})` : r[rowKey]))].sort();
  const cells = {};
  users.forEach(u => { cells[u] = {}; });
  rows.forEach(r => {
    const u = r[userKey];
    const k = secondaryKey ? `${r[rowKey]} (${r[secondaryKey]})` : r[rowKey];
    cells[u][k] = (cells[u][k] || 0) + (r[valueKey] || 0);
  });
  return { users, keys, cells };
}

export default function Dashboard() {
  const { profile, visibleCodes, signOut } = useAuth();
  const [lang, setLang]      = useState(profile?.preferred_lang || 'en');
  const [period, setPeriod]  = useState('last_month');
  const [team, setTeam]      = useState('all');
  const [tab, setTab]        = useState('summary');
  const [summary, setSummary]     = useState([]);
  const [specialty, setSpecialty] = useState([]);
  const [products, setProducts]   = useState([]);
  const [coaching, setCoaching]   = useState([]);
  const [loading, setLoading]     = useState(true);
  const [error, setError]         = useState('');

  const t   = T[lang] || T.en;
  const rtl = lang === 'ar';
  const periodLabel = period === 'last_month' ? 'Last Month' : 'Recent';

  const loadData = useCallback(async () => {
    if (!visibleCodes?.length) { setLoading(false); return; }
    setLoading(true); setError('');
    const codes = visibleCodes;

    const coachingFilter = codes.map(c => `manager_code.eq.${c}`).concat(codes.map(c => `rep_code.eq.${c}`)).join(',');

    const [s, sp, pr, co] = await Promise.all([
      supabase.from('summaries').select('*').eq('period', periodLabel).in('employee_code', codes),
      supabase.from('specialty_classification').select('*').eq('period', periodLabel).in('employee_code', codes),
      supabase.from('product_calls').select('*').eq('period', periodLabel).in('employee_code', codes),
      supabase.from('coaching_days').select('*').eq('period', periodLabel).or(coachingFilter),
    ]);

    if (s.error) setError(s.error.message);
    setSummary(s.data || []);
    setSpecialty(sp.data || []);
    setProducts(pr.data || []);
    setCoaching(co.data || []);
    setLoading(false);
  }, [periodLabel, visibleCodes]);

  useEffect(() => { loadData(); }, [loadData]);

  const teams = useMemo(() => [...new Set(summary.map(r=>r.team).filter(Boolean))].sort(), [summary]);

  const filterByTeam = useCallback(rows =>
    team === 'all' ? rows : rows.filter(r => r.team === team), [team]);

  const fSummary   = useMemo(() => sortSummary(filterByTeam(summary)), [summary, filterByTeam]);
  const fSpecialty = useMemo(() => filterByTeam(specialty), [specialty, filterByTeam]);
  const fProducts  = useMemo(() => filterByTeam(products), [products, filterByTeam]);
  const fCoaching  = useMemo(() => filterByTeam(coaching), [coaching, filterByTeam]);

  const teamCount = new Set(fSummary.map(r=>r.team)).size;

  function handleExport() {
    const wb = XLSX.utils.book_new();
    // Summary sheet (horizontal for export)
    const sHeaders = ['Team','User','Territory','Manager',...KPI_GROUPS.flatMap(g=>g.keys)];
    const sData = fSummary.map(r => [
      r.team, r.user_name, r.territory, r.is_manager?'✓':'',
      ...KPI_GROUPS.flatMap(g=>g.keys.map(k=>r[k]??''))
    ]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([sHeaders,...sData]), 'Summary');

    // Specialty pivot
    const { users:su, keys:sk, cells:sc } = pivotRows(fSpecialty,'user_name','specialty','call_count','shift');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['User',...sk], ...su.map(u=>[u,...sk.map(k=>sc[u][k]||0)])
    ]), 'Specialty');

    // Products pivot
    const { users:pu, keys:pk, cells:pc } = pivotRows(fProducts,'user_name','product','call_count','shift');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['User',...pk], ...pu.map(u=>[u,...pk.map(k=>pc[u][k]||0)])
    ]), 'Products');

    XLSX.writeFile(wb, `excellence_${periodLabel.replace(' ','_')}_${Date.now()}.xlsx`);
  }

  return (
    <div className={`app${rtl?' rtl':''}`} dir={rtl?'rtl':'ltr'}>
      {/* ── Header ── */}
      <header className="hdr">
        <div className="hdr-left">
          <span className="hdr-brand">{t.appName}</span>
          <span className="hdr-view">{t.roleView[profile?.role] || ''}</span>
        </div>
        <div className="hdr-right">
          {profile?.role === 'Admin' && (
            <a className="btn btn-ghost" href="/admin">{t.admin}</a>
          )}
          <button className="btn btn-lang" onClick={()=>setLang(lang==='en'?'ar':'en')}>
            {lang==='en'?'عربي':'EN'}
          </button>
          <div className="hdr-user">
            <div className="hdr-name">{profile?.employee_name}</div>
            <div className="hdr-role">{profile?.role} · {profile?.employee_code}</div>
          </div>
          <button className="btn btn-ghost" onClick={signOut}>{t.signOut}</button>
        </div>
      </header>

      {/* ── Toolbar ── */}
      <div className="toolbar">
        <div className="period-pills">
          {[
            {k:'last_month', label:t.periodLastMonth},
            {k:'recent',     label:t.periodRecent},
          ].map(p=>(
            <button key={p.k} className={`pill${period===p.k?' pill-on':''}`}
              onClick={()=>setPeriod(p.k)}>{p.label}</button>
          ))}
        </div>

        {teams.length > 1 && (
          <select className="team-sel" value={team} onChange={e=>setTeam(e.target.value)}>
            <option value="all">{t.allTeams}</option>
            {teams.map(tm=><option key={tm} value={tm}>{tm}</option>)}
          </select>
        )}

        <span className="toolbar-stat">
          {t.people(fSummary.length)}{teamCount>1 && t.acrossTeams(teamCount)}
        </span>

        <button className="btn btn-primary ml-auto" onClick={handleExport}>
          ↓ {t.download}
        </button>
      </div>

      {/* ── Tabs ── */}
      <nav className="tabs">
        {Object.entries(t.tabs).map(([k,label])=>(
          <button key={k} className={`tab-btn${tab===k?' on':''}`} onClick={()=>setTab(k)}>
            {label}
          </button>
        ))}
      </nav>

      {error && <div className="banner-error">{error}</div>}

      {loading ? (
        <div className="empty-state">{t.loading}</div>
      ) : (
        <div className="content">
          {/* ── SUMMARY: vertical KPI card per user ── */}
          {tab === 'summary' && (
            fSummary.length === 0 ? <div className="empty-state">{t.noData}</div> : (
              <div className="cards-grid">
                {fSummary.map((r,i) => (
                  <div key={r.id||i} className={`user-card${r.is_manager?' mgr':''}`}>
                    <div className="card-header">
                      <div>
                        <div className="card-name">{r.user_name}</div>
                        <div className="card-meta">{r.team}{r.is_manager?' · Manager':''}</div>
                        {r.territory && <div className="card-terr">{r.territory}</div>}
                      </div>
                      {r.is_manager && <span className="mgr-badge">MGR</span>}
                    </div>
                    {KPI_GROUPS.map(group => (
                      <div key={group.label.en} className="kpi-group">
                        <div className="kpi-group-label">{group.label[lang] || group.label.en}</div>
                        {group.keys.map(key => {
                          const val = r[key];
                          if (val === null || val === undefined || val === '') return null;
                          return (
                            <div key={key} className="kpi-row">
                              <span className="kpi-label">{t.kpi[key] || key}</span>
                              <span className="kpi-val">{fmtVal(val, key)}</span>
                            </div>
                          );
                        })}
                      </div>
                    ))}
                    {r.product_calls_detail && (
                      <div className="kpi-group">
                        <div className="kpi-group-label">{lang==='ar'?'تفاصيل المنتج':'Product Detail'}</div>
                        <div className="prod-detail">{r.product_calls_detail}</div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )
          )}

          {/* ── SPECIALTY: pivoted table — rows = specialties, cols = users ── */}
          {tab === 'specialty' && (() => {
            if (!fSpecialty.length) return <div className="empty-state">{t.noData}</div>;
            // Pivot: specialty+shift rows, user cols
            const users = [...new Set(fSpecialty.map(r=>r.user_name))].sort();
            const specKeys = [...new Set(fSpecialty.map(r=>`${r.specialty||'?'} · ${r.shift||'?'}`))].sort();
            const cells = {};
            fSpecialty.forEach(r => {
              const k = `${r.specialty||'?'} · ${r.shift||'?'}`;
              if (!cells[k]) cells[k] = {};
              cells[k][r.user_name] = (cells[k][r.user_name]||0) + (r.call_count||0);
            });
            return (
              <div className="pivot-wrap">
                <table className="pivot-table">
                  <thead>
                    <tr>
                      <th className="sticky-col">{lang==='ar'?'التخصص / الفترة':'Specialty / Shift'}</th>
                      {users.map(u => <th key={u}>{u}</th>)}
                      <th className="total-col">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {specKeys.map(k => {
                      const rowTotal = users.reduce((s,u)=>s+(cells[k][u]||0),0);
                      return (
                        <tr key={k}>
                          <td className="sticky-col row-head">{k}</td>
                          {users.map(u=><td key={u} className={cells[k][u]?'has-val':'empty-val'}>{cells[k][u]||''}</td>)}
                          <td className="total-col">{rowTotal}</td>
                        </tr>
                      );
                    })}
                    <tr className="total-row">
                      <td className="sticky-col">Total</td>
                      {users.map(u=><td key={u}>{fSpecialty.filter(r=>r.user_name===u).reduce((s,r)=>s+(r.call_count||0),0)}</td>)}
                      <td className="total-col">{fSpecialty.reduce((s,r)=>s+(r.call_count||0),0)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            );
          })()}

          {/* ── PRODUCTS: pivoted table — rows = products, cols = users ── */}
          {tab === 'products' && (() => {
            if (!fProducts.length) return <div className="empty-state">{t.noData}</div>;
            const users = [...new Set(fProducts.map(r=>r.user_name))].sort();
            const prodKeys = [...new Set(fProducts.map(r=>`${r.product||'?'} · ${r.shift||'?'}`))].sort();
            const cells = {};
            fProducts.forEach(r => {
              const k = `${r.product||'?'} · ${r.shift||'?'}`;
              if (!cells[k]) cells[k]={};
              cells[k][r.user_name] = (cells[k][r.user_name]||0) + (r.call_count||0);
            });
            return (
              <div className="pivot-wrap">
                <table className="pivot-table">
                  <thead>
                    <tr>
                      <th className="sticky-col">{lang==='ar'?'المنتج / الفترة':'Product / Shift'}</th>
                      {users.map(u=><th key={u}>{u}</th>)}
                      <th className="total-col">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {prodKeys.map(k => {
                      const rowTotal = users.reduce((s,u)=>s+(cells[k][u]||0),0);
                      return (
                        <tr key={k}>
                          <td className="sticky-col row-head">{k}</td>
                          {users.map(u=><td key={u} className={cells[k][u]?'has-val':'empty-val'}>{cells[k][u]||''}</td>)}
                          <td className="total-col">{rowTotal}</td>
                        </tr>
                      );
                    })}
                    <tr className="total-row">
                      <td className="sticky-col">Total</td>
                      {users.map(u=><td key={u}>{fProducts.filter(r=>r.user_name===u).reduce((s,r)=>s+(r.call_count||0),0)}</td>)}
                      <td className="total-col">{fProducts.reduce((s,r)=>s+(r.call_count||0),0)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            );
          })()}

          {/* ── COACHING ── */}
          {tab === 'coaching' && (
            fCoaching.length === 0 ? <div className="empty-state">{t.noData}</div> : (
              <div className="pivot-wrap">
                <table className="pivot-table">
                  <thead>
                    <tr>
                      <th>Team</th><th>Manager</th><th>Rep</th><th>Date</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...fCoaching].sort((a,b)=>(a.manager_name||'').localeCompare(b.manager_name||'')).map((r,i)=>(
                      <tr key={r.id||i}>
                        <td>{r.team||'—'}</td>
                        <td className="row-head">{r.manager_name}</td>
                        <td>{r.rep_name}</td>
                        <td>{r.coaching_date}</td>
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
