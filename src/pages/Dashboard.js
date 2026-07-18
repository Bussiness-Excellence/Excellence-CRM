import React, { useEffect, useMemo, useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, ComposedChart, Line,
  Cell, LabelList, PieChart, Pie,
} from 'recharts';
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
    searchPlaceholder: 'Search employees…',
    people: n => `${n} ${n===1?'rep':'reps'}`,
    acrossTeams: n => ` · ${n} teams`,
    tabs: { summary:'Summary', specialty:'Specialty', products:'Products', coaching:'Coaching', charts:'Charts' },
    roleView: { MR:'My Results', Supervisor:'My Team', 'Area Manager':'My Area', BLM:'Full Team', Admin:'All Teams' },
    charts: {
      fieldActivity: 'Field Activity',
      calls: 'Calls & Rates',
      coverage: 'Coverage',
      productsChart: 'Product Calls',
      timing: 'Timing',
      specialtyShare: 'Calls by Specialty',
      productShare: 'Calls by Product',
    },
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
    searchPlaceholder: 'ابحث عن موظف…',
    people: n => `${n} مندوب`,
    acrossTeams: n => ` · ${n} فرق`,
    tabs: { summary:'الملخص', specialty:'التخصص', products:'المنتجات', coaching:'التوجيه', charts:'مخططات' },
    roleView: { MR:'نتائجي', Supervisor:'فريقي', 'Area Manager':'منطقتي', BLM:'الفريق كاملاً', Admin:'كل الفرق' },
    charts: {
      fieldActivity: 'النشاط الميداني',
      calls: 'الزيارات والمعدلات',
      coverage: 'التغطية',
      productsChart: 'مكالمات المنتج',
      timing: 'التوقيت',
      specialtyShare: 'الزيارات حسب التخصص',
      productShare: 'الزيارات حسب المنتج',
    },
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

// Chart color palette
const COLORS = {
  navy:   '#16243e',
  navy2:  '#253668',
  gold:   '#c8a84b',
  gold2:  '#e0c06a',
  teal:   '#2a9d8f',
  coral:  '#e76f51',
  blue:   '#3b82f6',
  green:  '#2d9e6b',
  purple: '#7c3aed',
  rose:   '#fda4af',
  slate:  '#64748b',
};

const PIE_PALETTE = [
  COLORS.navy,
  COLORS.gold,
  COLORS.teal,
  COLORS.blue,
  COLORS.coral,
  COLORS.purple,
  COLORS.green,
  COLORS.slate
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

function pivotRows(rows, userKey, rowKey, valueKey, secondaryKey) {
  const users = [...new Set(rows.map(r => r[userKey]))].sort();
  const keys  = [...new Set(rows.map(r => secondaryKey ? `${r[rowKey]} (${r[secondaryKey]})` : r[rowKey]))].sort();
  const cells = {};
  users.forEach(u => { cells[u] = {}; });
  rows.forEach(r => {
    const u = r[userKey];
    const k = secondaryKey ? `${r[rowKey]} (${r[secondaryKey]})` : r[rowKey];
    cells[u][k] = (cells[u][k] || 0) + (r[valueKey] || 0);
  });
  return { users, keys, cells };
}

// ── Skeleton loader ────────────────────────────────────────────────────────
function Skeleton() {
  return (
    <div className="skeleton-grid">
      {[...Array(6)].map((_,i) => (
        <div key={i} className="skeleton-card">
          <div className="skeleton-header" />
          {[...Array(5)].map((_,j) => (
            <div key={j} className="skeleton-row">
              <div className="skeleton-bar" style={{ width: `${45+Math.random()*40}%` }} />
              <div className="skeleton-bar short" />
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}

// ── Custom tooltip for charts ──────────────────────────────────────────────
function ChartTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div className="chart-tooltip">
      <div className="chart-tooltip-label">{label}</div>
      {payload.map((p, i) => (
        <div key={i} className="chart-tooltip-row">
          <span className="chart-tooltip-dot" style={{ background: p.color }} />
          <span className="chart-tooltip-name">{p.name}:</span>
          <span className="chart-tooltip-val">{typeof p.value === 'number' ? p.value.toLocaleString(undefined, {maximumFractionDigits: 1}) : p.value}</span>
        </div>
      ))}
    </div>
  );
}

// ── Chart section wrapper ──────────────────────────────────────────────────
function ChartSection({ title, children, halfWidth }) {
  return (
    <div className={`chart-section${halfWidth ? ' half' : ''}`}>
      <div className="chart-section-title">{title}</div>
      <div className="chart-section-body">{children}</div>
    </div>
  );
}

// ── Charts Tab ────────────────────────────────────────────────────────────
function ChartsTab({ data, specialty, products, lang, t }) {
  // Specialty Pie Chart Data (Top 7 + Others)
  const specialtyPieData = useMemo(() => {
    const counts = {};
    specialty.forEach(r => {
      const name = r.specialty || 'Other';
      counts[name] = (counts[name] || 0) + (r.call_count || 0);
    });
    const items = Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a,b) => b.value - a.value);
    
    if (items.length <= 7) return items;
    const top = items.slice(0, 6);
    const otherSum = items.slice(6).reduce((s, x) => s + x.value, 0);
    top.push({ name: lang === 'ar' ? 'أخرى' : 'Other Specialties', value: otherSum });
    return top;
  }, [specialty, lang]);

  // Product Pie Chart Data (Top 7 + Others)
  const productPieData = useMemo(() => {
    const counts = {};
    products.forEach(r => {
      const name = r.product || 'Other';
      counts[name] = (counts[name] || 0) + (r.call_count || 0);
    });
    const items = Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a,b) => b.value - a.value);

    if (items.length <= 7) return items;
    const top = items.slice(0, 6);
    const otherSum = items.slice(6).reduce((s, x) => s + x.value, 0);
    top.push({ name: lang === 'ar' ? 'أخرى' : 'Other Products', value: otherSum });
    return top;
  }, [products, lang]);

  if (!data.length) return <div className="empty-state">{t.noData}</div>;

  // Abbreviate long names for chart labels
  const abbr = name => {
    if (!name) return '';
    const parts = name.trim().split(/\s+/);
    if (parts.length === 1) return name.slice(0,12);
    return parts.length <= 2 ? name : `${parts[0]} ${parts[parts.length-1]}`;
  };

  const sorted = [...data].sort((a,b) => (b.complete_field_days||0) - (a.complete_field_days||0));

  // Build chart datasets
  const fieldData = sorted.map(r => ({
    name: abbr(r.user_name),
    fullName: r.user_name,
    [t.kpi.complete_field_days]: r.complete_field_days || 0,
    [t.kpi.am_shift_days]:       r.am_shift_days       || 0,
    [t.kpi.pm_shift_days]:       r.pm_shift_days       || 0,
    [t.kpi.double_visit_days]:   r.double_visit_days   || 0,
  }));

  const callData = sorted.map(r => ({
    name: abbr(r.user_name),
    fullName: r.user_name,
    [t.kpi.am_calls]:     r.am_calls     || 0,
    [t.kpi.pm_calls]:     r.pm_calls     || 0,
    [t.kpi.am_call_rate]: parseFloat(r.am_call_rate) || 0,
    [t.kpi.pm_call_rate]: parseFloat(r.pm_call_rate) || 0,
  }));

  const coverageData = sorted.map(r => ({
    name: abbr(r.user_name),
    fullName: r.user_name,
    [t.kpi.pharmacies_visited]:  r.pharmacies_visited  || 0,
    [t.kpi.pharmacies_covered]:  r.pharmacies_covered  || 0,
    [t.kpi.total_am_covered]:    r.total_am_covered    || 0,
    [t.kpi.total_pm_covered]:    r.total_pm_covered    || 0,
  }));

  const productData = [...data]
    .sort((a,b) => (b.total_product_calls||0) - (a.total_product_calls||0))
    .map(r => ({
      name: abbr(r.user_name),
      fullName: r.user_name,
      [t.kpi.total_product_calls]: r.total_product_calls || 0,
      [t.kpi.distinct_products]:   r.distinct_products   || 0,
    }));

  const timingData = sorted
    .filter(r => r.avg_am_shift_hm || r.avg_pm_shift_hm)
    .map(r => ({
      name: abbr(r.user_name),
      fullName: r.user_name,
      [t.kpi.avg_am_shift_hm]: parseFloat(r.avg_am_shift_hm) || 0,
      [t.kpi.avg_pm_shift_hm]: parseFloat(r.avg_pm_shift_hm) || 0,
    }));

  const chartH = Math.max(320, sorted.length * 38);

  return (
    <div className="charts-container">

      {/* ── Pie Charts Row (Specialty & Product Distribution) ── */}
      <div className="charts-pie-row">
        <ChartSection title={t.charts.specialtyShare} halfWidth>
          <div className="pie-wrap">
            <ResponsiveContainer width="100%" height={260}>
              <PieChart>
                <Pie
                  data={specialtyPieData}
                  dataKey="value"
                  nameKey="name"
                  cx="50%"
                  cy="50%"
                  outerRadius={80}
                  innerRadius={30}
                  paddingAngle={2}
                  label={({ name, percent }) => `${name} (${(percent * 100).toFixed(0)}%)`}
                >
                  {specialtyPieData.map((_, index) => (
                    <Cell key={index} fill={PIE_PALETTE[index % PIE_PALETTE.length]} />
                  ))}
                </Pie>
                <Tooltip content={<ChartTooltip />} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </ChartSection>

        <ChartSection title={t.charts.productShare} halfWidth>
          <div className="pie-wrap">
            <ResponsiveContainer width="100%" height={260}>
              <PieChart>
                <Pie
                  data={productPieData}
                  dataKey="value"
                  nameKey="name"
                  cx="50%"
                  cy="50%"
                  outerRadius={80}
                  innerRadius={30}
                  paddingAngle={2}
                  label={({ name, percent }) => `${name} (${(percent * 100).toFixed(0)}%)`}
                >
                  {productPieData.map((_, index) => (
                    <Cell key={index} fill={PIE_PALETTE[index % PIE_PALETTE.length]} />
                  ))}
                </Pie>
                <Tooltip content={<ChartTooltip />} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </ChartSection>
      </div>

      {/* ── Field Activity ── */}
      <ChartSection title={t.charts.fieldActivity}>
        <ResponsiveContainer width="100%" height={chartH}>
          <BarChart data={fieldData} layout="vertical" margin={{ left: 16, right: 40, top: 8, bottom: 8 }}>
            <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e3e8f0" />
            <XAxis type="number" tick={{ fontSize: 11, fill: '#6b7a9a' }} />
            <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 11, fill: '#16243e', fontWeight: 600 }} />
            <Tooltip content={<ChartTooltip />} />
            <Legend wrapperStyle={{ fontSize: 11, paddingTop: 8 }} />
            <Bar dataKey={t.kpi.complete_field_days} fill={COLORS.navy}  radius={[0,3,3,0]} barSize={9}>
              <LabelList dataKey={t.kpi.complete_field_days} position="right" style={{ fontSize: 10, fill: '#16243e', fontWeight: 700 }} />
            </Bar>
            <Bar dataKey={t.kpi.am_shift_days}       fill={COLORS.teal}  radius={[0,3,3,0]} barSize={9} />
            <Bar dataKey={t.kpi.pm_shift_days}       fill={COLORS.gold}  radius={[0,3,3,0]} barSize={9} />
            <Bar dataKey={t.kpi.double_visit_days}   fill={COLORS.coral} radius={[0,3,3,0]} barSize={9} />
          </BarChart>
        </ResponsiveContainer>
      </ChartSection>

      {/* ── Calls & Rates ── */}
      <ChartSection title={t.charts.calls}>
        <ResponsiveContainer width="100%" height={chartH}>
          <ComposedChart data={callData} layout="vertical" margin={{ left: 16, right: 60, top: 8, bottom: 8 }}>
            <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e3e8f0" />
            <XAxis type="number" tick={{ fontSize: 11, fill: '#6b7a9a' }} />
            <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 11, fill: '#16243e', fontWeight: 600 }} />
            <Tooltip content={<ChartTooltip />} />
            <Legend wrapperStyle={{ fontSize: 11, paddingTop: 8 }} />
            <Bar dataKey={t.kpi.am_calls} fill={COLORS.navy} radius={[0,3,3,0]} barSize={10}>
              <LabelList dataKey={t.kpi.am_calls} position="right" style={{ fontSize: 10, fill: '#16243e', fontWeight: 700 }} />
            </Bar>
            <Bar dataKey={t.kpi.pm_calls} fill={COLORS.teal} radius={[0,3,3,0]} barSize={10} />
            <Line dataKey={t.kpi.am_call_rate} stroke={COLORS.gold}   strokeWidth={2} dot={{ r: 3 }} />
            <Line dataKey={t.kpi.pm_call_rate} stroke={COLORS.coral}  strokeWidth={2} dot={{ r: 3 }} />
          </ComposedChart>
        </ResponsiveContainer>
      </ChartSection>

      {/* ── Coverage ── */}
      <ChartSection title={t.charts.coverage}>
        <ResponsiveContainer width="100%" height={chartH}>
          <BarChart data={coverageData} layout="vertical" margin={{ left: 16, right: 40, top: 8, bottom: 8 }}>
            <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e3e8f0" />
            <XAxis type="number" tick={{ fontSize: 11, fill: '#6b7a9a' }} />
            <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 11, fill: '#16243e', fontWeight: 600 }} />
            <Tooltip content={<ChartTooltip />} />
            <Legend wrapperStyle={{ fontSize: 11, paddingTop: 8 }} />
            <Bar dataKey={t.kpi.pharmacies_visited}  fill={COLORS.navy}   radius={[0,3,3,0]} barSize={9}>
              <LabelList dataKey={t.kpi.pharmacies_visited} position="right" style={{ fontSize: 10, fill: '#16243e', fontWeight: 700 }} />
            </Bar>
            <Bar dataKey={t.kpi.pharmacies_covered}  fill={COLORS.blue}   radius={[0,3,3,0]} barSize={9} />
            <Bar dataKey={t.kpi.total_am_covered}    fill={COLORS.teal}   radius={[0,3,3,0]} barSize={9} />
            <Bar dataKey={t.kpi.total_pm_covered}    fill={COLORS.gold}   radius={[0,3,3,0]} barSize={9} />
          </BarChart>
        </ResponsiveContainer>
      </ChartSection>

      {/* ── Product Calls ── */}
      <ChartSection title={t.charts.productsChart}>
        <ResponsiveContainer width="100%" height={chartH}>
          <BarChart data={productData} layout="vertical" margin={{ left: 16, right: 50, top: 8, bottom: 8 }}>
            <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e3e8f0" />
            <XAxis type="number" tick={{ fontSize: 11, fill: '#6b7a9a' }} />
            <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 11, fill: '#16243e', fontWeight: 600 }} />
            <Tooltip content={<ChartTooltip />} />
            <Legend wrapperStyle={{ fontSize: 11, paddingTop: 8 }} />
            <Bar dataKey={t.kpi.total_product_calls} fill={COLORS.purple} radius={[0,3,3,0]} barSize={12}>
              <LabelList dataKey={t.kpi.total_product_calls} position="right" style={{ fontSize: 10, fill: '#16243e', fontWeight: 700 }} />
              {productData.map((_, i) => (
                <Cell key={i} fill={i === 0 ? COLORS.gold : COLORS.purple} />
              ))}
            </Bar>
            <Bar dataKey={t.kpi.distinct_products} fill={COLORS.teal} radius={[0,3,3,0]} barSize={12} />
          </BarChart>
        </ResponsiveContainer>
      </ChartSection>

      {/* ── Timing ── */}
      {timingData.length > 0 && (
        <ChartSection title={t.charts.timing}>
          <ResponsiveContainer width="100%" height={Math.max(280, timingData.length * 38)}>
            <BarChart data={timingData} layout="vertical" margin={{ left: 16, right: 40, top: 8, bottom: 8 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e3e8f0" />
              <XAxis type="number" tick={{ fontSize: 11, fill: '#6b7a9a' }} unit="h" />
              <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 11, fill: '#16243e', fontWeight: 600 }} />
              <Tooltip content={<ChartTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, paddingTop: 8 }} />
              <Bar dataKey={t.kpi.avg_am_shift_hm} fill={COLORS.teal}  radius={[0,3,3,0]} barSize={12}>
                <LabelList dataKey={t.kpi.avg_am_shift_hm} position="right" style={{ fontSize: 10, fill: '#16243e', fontWeight: 700 }} />
              </Bar>
              <Bar dataKey={t.kpi.avg_pm_shift_hm} fill={COLORS.gold}  radius={[0,3,3,0]} barSize={12} />
            </BarChart>
          </ResponsiveContainer>
        </ChartSection>
      )}

    </div>
  );
}

// ── Main Dashboard ─────────────────────────────────────────────────────────
export default function Dashboard() {
  const { profile, visibleCodes, signOut } = useAuth();
  const [lang, setLang]      = useState(profile?.preferred_lang || 'en');
  const [availablePeriods, setAvailablePeriods] = useState([]); // real period strings, e.g. "July 2026"
  const [period, setPeriod]  = useState(null);
  const [team, setTeam]      = useState('all');
  const [tab, setTab]        = useState('summary');
  const [search, setSearch]  = useState('');
  const [summary, setSummary]     = useState([]);
  const [specialty, setSpecialty] = useState([]);
  const [products, setProducts]   = useState([]);
  const [coaching, setCoaching]   = useState([]);
  const [loading, setLoading]     = useState(true);
  const [error, setError]         = useState('');

  const t   = T[lang] || T.en;
  const rtl = lang === 'ar';

  // Load the real, distinct periods that actually exist in the data, sorted
  // most-recent-first. "Last Month"/"Recent" as fixed literal strings never
  // matched any real row (periods are things like "July 2026"), which is
  // why the dashboard always showed "No data available" — this replaces
  // that with the actual values so the period pills always match something.
  useEffect(() => {
    supabase.from('summaries').select('period').then(({ data, error }) => {
      if (error) { setError(error.message); return; }
      const uniq = [...new Set((data || []).map(r => r.period).filter(Boolean))];
      uniq.sort((a, b) => new Date(`1 ${b}`) - new Date(`1 ${a}`));
      setAvailablePeriods(uniq);
    });
  }, []);

  // Default to the most recent period once periods have loaded.
  useEffect(() => {
    if (availablePeriods.length && !period) setPeriod(availablePeriods[0]);
  }, [availablePeriods, period]);

  const periodLabel = period; // the real DB value — no more string mismatch

  const loadData = useCallback(async () => {
    if (!visibleCodes?.length || !periodLabel) { setLoading(false); return; }
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

  // Search filter — applies across all tabs
  const filterBySearch = useCallback(rows => {
    const q = search.trim().toLowerCase();
    if (!q) return rows;
    return rows.filter(r =>
      (r.user_name  || '').toLowerCase().includes(q) ||
      (r.territory  || '').toLowerCase().includes(q) ||
      (r.team       || '').toLowerCase().includes(q)
    );
  }, [search]);

  const fSummary   = useMemo(() => sortSummary(filterBySearch(filterByTeam(summary))),   [summary, filterByTeam, filterBySearch]);
  const fSpecialty = useMemo(() => filterBySearch(filterByTeam(specialty)), [specialty, filterByTeam, filterBySearch]);
  const fProducts  = useMemo(() => filterBySearch(filterByTeam(products)),  [products, filterByTeam, filterBySearch]);
  const fCoaching  = useMemo(() => {
    const q = search.trim().toLowerCase();
    let rows = filterByTeam(coaching);
    if (q) rows = rows.filter(r =>
      (r.manager_name || '').toLowerCase().includes(q) ||
      (r.rep_name     || '').toLowerCase().includes(q) ||
      (r.team         || '').toLowerCase().includes(q)
    );
    return rows;
  }, [coaching, filterByTeam, search]);

  const teamCount = new Set(fSummary.map(r=>r.team)).size;

  function handleExport() {
    const wb = XLSX.utils.book_new();
    const sHeaders = ['Team','User','Territory','Manager',...KPI_GROUPS.flatMap(g=>g.keys)];
    const sData = fSummary.map(r => [
      r.team, r.user_name, r.territory, r.is_manager?'✓':'',
      ...KPI_GROUPS.flatMap(g=>g.keys.map(k=>r[k]??''))
    ]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([sHeaders,...sData]), 'Summary');

    const { users:su, keys:sk, cells:sc } = pivotRows(fSpecialty,'user_name','specialty','call_count','shift');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['User',...sk], ...su.map(u=>[u,...sk.map(k=>sc[u][k]||0)])
    ]), 'Specialty');

    const { users:pu, keys:pk, cells:pc } = pivotRows(fProducts,'user_name','product','call_count','shift');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['User',...pk], ...pu.map(u=>[u,...pk.map(k=>pc[u][k]||0)])
    ]), 'Products');

    XLSX.writeFile(wb, `excellence_${(periodLabel||'export').replace(' ','_')}_${Date.now()}.xlsx`);
  }

  return (
    <div className={`app${rtl?' rtl':''}`} dir={rtl?'rtl':'ltr'}>

      {/* ── Header ── */}
      <header className="hdr">
        <div className="hdr-left">
          <div className="hdr-brand-wrap">
            <img src="/eipico-logo.png" alt="EIPICO" className="hdr-logo" />
            <div className="hdr-brand-divider" />
            <img src="/dept-logo.png" alt="Excellence" className="hdr-logo hdr-logo-dept" />
          </div>
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
          {availablePeriods.map(p => (
            <button key={p} className={`pill${period===p?' pill-on':''}`}
              onClick={()=>setPeriod(p)}>{p}</button>
          ))}
        </div>

        {teams.length > 1 && (
          <select className="team-sel" value={team} onChange={e=>setTeam(e.target.value)}>
            <option value="all">{t.allTeams}</option>
            {teams.map(tm=><option key={tm} value={tm}>{tm}</option>)}
          </select>
        )}

        {/* ── Search bar ── */}
        <div className="search-wrap">
          <svg className="search-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="9" cy="9" r="6"/><path d="M15 15l3 3" strokeLinecap="round"/>
          </svg>
          <input
            className="search-input"
            type="search"
            placeholder={t.searchPlaceholder}
            value={search}
            onChange={e=>{ setSearch(e.target.value); }}
          />
          {search && (
            <button className="search-clear" onClick={()=>setSearch('')} title="Clear">✕</button>
          )}
        </div>

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
        <div className="content"><Skeleton /></div>
      ) : (
        <div className="content">

          {/* ── SUMMARY ── */}
          {tab === 'summary' && (
            fSummary.length === 0 ? <div className="empty-state">{search ? `No results for "${search}"` : t.noData}</div> : (
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

          {/* ── SPECIALTY ── */}
          {tab === 'specialty' && (() => {
            if (!fSpecialty.length) return <div className="empty-state">{search ? `No results for "${search}"` : t.noData}</div>;
            const users    = [...new Set(fSpecialty.map(r=>r.user_name))].sort();
            const specKeys = [...new Set(fSpecialty.map(r=>`${r.specialty||'?'} · ${r.shift||'?'}`))].sort();
            const cells = {};
            fSpecialty.forEach(r => {
              const k = `${r.specialty||'?'} · ${r.shift||'?'}`;
              if (!cells[r.user_name]) cells[r.user_name] = {};
              cells[r.user_name][k] = (cells[r.user_name][k]||0) + (r.call_count||0);
            });
            return (
              <div className="pivot-wrap">
                <table className="pivot-table">
                  <thead>
                    <tr>
                      <th className="sticky-col">{lang==='ar'?'المستخدم':'User'}</th>
                      {specKeys.map(k => <th key={k}>{k}</th>)}
                      <th className="total-col">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {users.map(u => {
                      const rowTotal = specKeys.reduce((s,k)=>s+(cells[u][k]||0),0);
                      return (
                        <tr key={u}>
                          <td className="sticky-col row-head">{u}</td>
                          {specKeys.map(k=><td key={k} className={cells[u][k]?'has-val':'empty-val'}>{cells[u][k]||''}</td>)}
                          <td className="total-col">{rowTotal}</td>
                        </tr>
                      );
                    })}
                    <tr className="total-row">
                      <td className="sticky-col">Total</td>
                      {specKeys.map(k=><td key={k}>{users.reduce((s,u)=>s+(cells[u][k]||0),0)}</td>)}
                      <td className="total-col">{fSpecialty.reduce((s,r)=>s+(r.call_count||0),0)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            );
          })()}

          {/* ── PRODUCTS ── */}
          {tab === 'products' && (() => {
            if (!fProducts.length) return <div className="empty-state">{search ? `No results for "${search}"` : t.noData}</div>;
            const users    = [...new Set(fProducts.map(r=>r.user_name))].sort();
            const prodKeys = [...new Set(fProducts.map(r=>`${r.product||'?'} · ${r.shift||'?'}`))].sort();
            const cells = {};
            fProducts.forEach(r => {
              const k = `${r.product||'?'} · ${r.shift||'?'}`;
              if (!cells[r.user_name]) cells[r.user_name]={};
              cells[r.user_name][k] = (cells[r.user_name][k]||0) + (r.call_count||0);
            });
            return (
              <div className="pivot-wrap">
                <table className="pivot-table">
                  <thead>
                    <tr>
                      <th className="sticky-col">{lang==='ar'?'المستخدم':'User'}</th>
                      {prodKeys.map(k=><th key={k}>{k}</th>)}
                      <th className="total-col">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {users.map(u => {
                      const rowTotal = prodKeys.reduce((s,k)=>s+(cells[u][k]||0),0);
                      return (
                        <tr key={u}>
                          <td className="sticky-col row-head">{u}</td>
                          {prodKeys.map(k=><td key={k} className={cells[u][k]?'has-val':'empty-val'}>{cells[u][k]||''}</td>)}
                          <td className="total-col">{rowTotal}</td>
                        </tr>
                      );
                    })}
                    <tr className="total-row">
                      <td className="sticky-col">Total</td>
                      {prodKeys.map(k=><td key={k}>{users.reduce((s,u)=>s+(cells[u][k]||0),0)}</td>)}
                      <td className="total-col">{fProducts.reduce((s,r)=>s+(r.call_count||0),0)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            );
          })()}

          {/* ── COACHING ── */}
          {tab === 'coaching' && (
            fCoaching.length === 0 ? <div className="empty-state">{search ? `No results for "${search}"` : t.noData}</div> : (
              <div className="pivot-wrap">
                <table className="pivot-table">
                  <thead>
                    <tr>
                      <th>{lang==='ar'?'الفريق':'Team'}</th>
                      <th>{lang==='ar'?'المدير':'Manager'}</th>
                      <th>{lang==='ar'?'المندوب':'Rep'}</th>
                      <th>{lang==='ar'?'التاريخ':'Date'}</th>
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

          {/* ── CHARTS ── */}
          {tab === 'charts' && (
            <ChartsTab data={fSummary} specialty={fSpecialty} products={fProducts} lang={lang} t={t} />
          )}

        </div>
      )}
    </div>
  );
}