import React, { useState, useEffect, createContext, useContext } from 'react';
import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from 'recharts';

// Supabase setup
const supabaseUrl = process.env.REACT_APP_SUPABASE_URL;
const supabaseKey = process.env.REACT_APP_SUPABASE_ANON_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

// Auth Context
const AuthContext = createContext();

function useAuth() {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error('useAuth must be used within AuthProvider');
  }
  return context;
}

// Main App
export default function App() {
  const [auth, setAuth] = useState(null);
  const [loading, setLoading] = useState(false);

  const login = async (employeeCode, password) => {
    setLoading(true);
    try {
      // Fetch employee from hierarchy
      const { data, error } = await supabase
        .from('hierarchy')
        .select('*')
        .eq('employee_code', employeeCode)
        .single();

      if (error || !data) {
        throw new Error('Employee not found');
      }

      // Verify password (using employee name as password)
      if (password !== data.employee_name) {
        throw new Error('Invalid password');
      }

      // Map role
      const roleMap = {
        'MR': 'rep',
        'Supervisor': 'supervisor',
        'Area Manager': 'area_manager',
        'BLM': 'blm',
        'Admin': 'admin'
      };

      const userRole = roleMap[data.role] || 'rep';

      setAuth({
        employeeCode,
        employeeName: data.employee_name,
        role: userRole,
        team: data.role,
        fullUser: data
      });
    } catch (err) {
      alert(`Login failed: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const logout = () => {
    setAuth(null);
  };

  if (!auth) {
    return <LoginPage onLogin={login} loading={loading} />;
  }

  return (
    <AuthContext.Provider value={{ auth, logout }}>
      <MainApp />
    </AuthContext.Provider>
  );
}

// Login Page
function LoginPage({ onLogin, loading }) {
  const [code, setCode] = useState('');
  const [password, setPassword] = useState('');

  const handleSubmit = (e) => {
    e.preventDefault();
    onLogin(code, password);
  };

  return (
    <div style={{
      minHeight: '100vh',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", "Roboto", sans-serif'
    }}>
      <div style={{
        background: 'white',
        padding: '3rem 2rem',
        borderRadius: '12px',
        boxShadow: '0 20px 25px rgba(0,0,0,0.15)',
        width: '100%',
        maxWidth: '420px'
      }}>
        <h1 style={{
          textAlign: 'center',
          marginBottom: '0.5rem',
          color: '#1f2937',
          fontSize: '2rem',
          fontWeight: '600'
        }}>
          Excellence CRM
        </h1>
        <p style={{
          textAlign: 'center',
          marginBottom: '2rem',
          color: '#6b7280',
          fontSize: '0.875rem'
        }}>
          Role-Based Sales Dashboard
        </p>

        <form onSubmit={handleSubmit}>
          <div style={{ marginBottom: '1.5rem' }}>
            <label style={{
              display: 'block',
              marginBottom: '0.5rem',
              fontWeight: '600',
              color: '#374151',
              fontSize: '0.95rem'
            }}>
              Employee Code
            </label>
            <input
              type="text"
              value={code}
              onChange={(e) => setCode(e.target.value)}
              placeholder="e.g., 13512"
              style={{
                width: '100%',
                padding: '0.75rem',
                border: '1px solid #d1d5db',
                borderRadius: '8px',
                fontSize: '1rem',
                boxSizing: 'border-box',
                fontFamily: 'inherit'
              }}
              disabled={loading}
              required
            />
          </div>

          <div style={{ marginBottom: '2rem' }}>
            <label style={{
              display: 'block',
              marginBottom: '0.5rem',
              fontWeight: '600',
              color: '#374151',
              fontSize: '0.95rem'
            }}>
              Password (Your Name)
            </label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Your full name"
              style={{
                width: '100%',
                padding: '0.75rem',
                border: '1px solid #d1d5db',
                borderRadius: '8px',
                fontSize: '1rem',
                boxSizing: 'border-box',
                fontFamily: 'inherit'
              }}
              disabled={loading}
              required
            />
          </div>

          <button
            type="submit"
            disabled={loading}
            style={{
              width: '100%',
              padding: '0.875rem',
              background: loading ? '#9ca3af' : '#667eea',
              color: 'white',
              border: 'none',
              borderRadius: '8px',
              fontSize: '1rem',
              fontWeight: '600',
              cursor: loading ? 'not-allowed' : 'pointer',
              transition: 'background 0.2s',
              fontFamily: 'inherit'
            }}
            onMouseOver={(e) => {
              if (!loading) e.target.style.background = '#5568d3';
            }}
            onMouseOut={(e) => {
              if (!loading) e.target.style.background = '#667eea';
            }}
          >
            {loading ? 'Logging in...' : 'Login'}
          </button>
        </form>

        <hr style={{ margin: '2rem 0', border: 'none', borderTop: '1px solid #e5e7eb' }} />

        <p style={{ fontSize: '0.85rem', color: '#6b7280', lineHeight: '1.5' }}>
          <strong>Demo:</strong> Use any employee code from your organization. Password is the employee's full name exactly as it appears in the hierarchy.
        </p>
      </div>
    </div>
  );
}

// Main App Layout
function MainApp() {
  const { auth, logout } = useAuth();
  const [currentPage, setCurrentPage] = useState('dashboard');

  const pages = [
    { id: 'dashboard', label: 'Dashboard', icon: '📊' },
    { id: 'summary', label: 'Summary', icon: '📋' },
    { id: 'visits', label: 'Visits', icon: '👥' },
    { id: 'coaching', label: 'Coaching', icon: '🎯' },
    { id: 'export', label: 'Export', icon: '⬇️' }
  ];

  return (
    <div style={{
      display: 'flex',
      minHeight: '100vh',
      background: '#f9fafb',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", "Roboto", sans-serif'
    }}>
      {/* Sidebar */}
      <div style={{
        width: '280px',
        background: '#1f2937',
        color: 'white',
        padding: '2rem 1.5rem',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
        display: 'flex',
        flexDirection: 'column'
      }}>
        <div style={{ marginBottom: '3rem' }}>
          <h2 style={{
            fontSize: '1.5rem',
            fontWeight: '700',
            marginBottom: '0.5rem'
          }}>
            Excellence CRM
          </h2>
          <p style={{ fontSize: '0.85rem', color: '#9ca3af' }}>
            Sales Analytics
          </p>
        </div>

        <nav style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
          {pages.map(page => (
            <button
              key={page.id}
              onClick={() => setCurrentPage(page.id)}
              style={{
                background: currentPage === page.id ? '#667eea' : 'transparent',
                color: 'white',
                border: 'none',
                padding: '0.875rem 1rem',
                borderRadius: '6px',
                cursor: 'pointer',
                textAlign: 'left',
                fontSize: '1rem',
                fontFamily: 'inherit',
                transition: 'all 0.2s',
                fontWeight: currentPage === page.id ? '600' : '500'
              }}
              onMouseOver={(e) => {
                if (currentPage !== page.id) {
                  e.target.style.background = '#374151';
                }
              }}
              onMouseOut={(e) => {
                if (currentPage !== page.id) {
                  e.target.style.background = 'transparent';
                }
              }}
            >
              <span style={{ marginRight: '0.5rem' }}>{page.icon}</span>
              {page.label}
            </button>
          ))}
        </nav>

        <div style={{
          paddingTop: '2rem',
          borderTop: '1px solid #374151'
        }}>
          <div style={{ marginBottom: '1rem' }}>
            <p style={{
              fontSize: '0.85rem',
              color: '#9ca3af',
              marginBottom: '0.25rem'
            }}>
              Logged in as:
            </p>
            <p style={{ fontWeight: '600', marginBottom: '0.25rem' }}>
              {auth.employeeName}
            </p>
            <p style={{ fontSize: '0.8rem', color: '#6b7280' }}>
              Role: <strong>{auth.role.toUpperCase()}</strong>
            </p>
          </div>
          <button
            onClick={logout}
            style={{
              width: '100%',
              padding: '0.625rem',
              background: '#ef4444',
              color: 'white',
              border: 'none',
              borderRadius: '6px',
              cursor: 'pointer',
              fontSize: '0.9rem',
              fontWeight: '500',
              fontFamily: 'inherit',
              transition: 'background 0.2s'
            }}
            onMouseOver={(e) => e.target.style.background = '#dc2626'}
            onMouseOut={(e) => e.target.style.background = '#ef4444'}
          >
            Logout
          </button>
        </div>
      </div>

      {/* Main Content */}
      <div style={{ flex: 1, padding: '2rem', overflowY: 'auto' }}>
        {currentPage === 'dashboard' && <DashboardPage />}
        {currentPage === 'summary' && <SummaryPage />}
        {currentPage === 'visits' && <VisitsPage />}
        {currentPage === 'coaching' && <CoachingPage />}
        {currentPage === 'export' && <ExportPage />}
      </div>
    </div>
  );
}

// Dashboard Page
function DashboardPage() {
  const { auth } = useAuth();
  const [kpis, setKpis] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      const codes = await getVisibleCodes();
      const { data: summaries } = await supabase
        .from('summaries')
        .select('*')
        .in('employee_code', codes);

      if (summaries && summaries.length > 0) {
        const workingDays = summaries.reduce((sum, s) => sum + (s.working_days || 0), 0);
        const activities = summaries.reduce((sum, s) => sum + (s.no_activities || 0), 0);
        const coachingDays = summaries.reduce((sum, s) => sum + (s.coaching_days || 0), 0);

        setKpis({
          workingDays,
          activities,
          coachingDays,
          avgPerDay: (activities / Math.max(workingDays, 1)).toFixed(1)
        });
      }
    } catch (err) {
      console.error('Error:', err);
    } finally {
      setLoading(false);
    }
  };

  const roleDescriptions = {
    'rep': '👤 You see only your own data: your visits, activities, coaching days, and performance metrics.',
    'supervisor': '👥 You see your direct reports plus your own data.',
    'area_manager': '🌍 You see all supervisors, their teams, and all reps under your area.',
    'blm': '🏢 You see the full data for your entire team.',
    'admin': '🔑 You have full access to all data across all 12 teams.'
  };

  if (loading) {
    return (
      <div style={{ padding: '2rem', textAlign: 'center', color: '#6b7280' }}>
        Loading dashboard...
      </div>
    );
  }

  const cards = [
    { label: 'Working Days', value: kpis?.workingDays || 0, color: '#667eea' },
    { label: 'Total Activities', value: kpis?.activities || 0, color: '#764ba2' },
    { label: 'Coaching Days', value: kpis?.coachingDays || 0, color: '#f59e0b' },
    { label: 'Avg/Day', value: kpis?.avgPerDay || 0, color: '#10b981' }
  ];

  return (
    <div>
      <h1 style={{ fontSize: '2rem', fontWeight: '600', marginBottom: '2rem', color: '#1f2937' }}>
        Dashboard
      </h1>

      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
        gap: '1.5rem',
        marginBottom: '2rem'
      }}>
        {cards.map((card, i) => (
          <div
            key={i}
            style={{
              background: 'white',
              padding: '1.5rem',
              borderRadius: '12px',
              boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
              borderLeft: `4px solid ${card.color}`
            }}
          >
            <p style={{ color: '#6b7280', fontSize: '0.9rem', marginBottom: '0.5rem' }}>
              {card.label}
            </p>
            <p style={{ fontSize: '2.5rem', fontWeight: '700', color: card.color }}>
              {typeof card.value === 'number' ? card.value.toLocaleString() : card.value}
            </p>
          </div>
        ))}
      </div>

      <div style={{
        background: 'white',
        padding: '2rem',
        borderRadius: '12px',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
      }}>
        <h2 style={{ fontSize: '1.25rem', fontWeight: '600', marginBottom: '1rem', color: '#1f2937' }}>
          Your Access Level
        </h2>
        <p style={{ color: '#6b7280', lineHeight: '1.6', fontSize: '1rem' }}>
          {roleDescriptions[auth.role]}
        </p>
      </div>
    </div>
  );
}

// Summary Page
function SummaryPage() {
  const { auth } = useAuth();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [sortBy, setSortBy] = useState('employee_name');
  const [filterTeam, setFilterTeam] = useState('');
  const [teams, setTeams] = useState([]);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      const codes = await getVisibleCodes();
      const { data: summaries } = await supabase
        .from('summaries')
        .select('*')
        .in('employee_code', codes);

      if (summaries) {
        setData(summaries);
        const uniqueTeams = [...new Set(summaries.map(s => s.team))].filter(Boolean);
        setTeams(uniqueTeams.sort());
      }
    } catch (err) {
      console.error('Error:', err);
    } finally {
      setLoading(false);
    }
  };

  const filtered = data
    .filter(item => !filterTeam || item.team === filterTeam)
    .sort((a, b) => {
      const aVal = a[sortBy] || '';
      const bVal = b[sortBy] || '';
      if (typeof aVal === 'number') return aVal - bVal;
      return String(aVal).localeCompare(String(bVal));
    });

  if (loading) {
    return <div style={{ padding: '2rem', color: '#6b7280' }}>Loading...</div>;
  }

  return (
    <div>
      <h1 style={{ fontSize: '2rem', fontWeight: '600', marginBottom: '2rem', color: '#1f2937' }}>
        Summary Table
      </h1>

      <div style={{ display: 'flex', gap: '1rem', marginBottom: '2rem' }}>
        <select
          value={filterTeam}
          onChange={(e) => setFilterTeam(e.target.value)}
          style={{
            padding: '0.75rem',
            border: '1px solid #d1d5db',
            borderRadius: '8px',
            fontSize: '1rem',
            fontFamily: 'inherit',
            flex: 1,
            maxWidth: '250px'
          }}
        >
          <option value="">All Teams ({data.length})</option>
          {teams.map(team => {
            const count = data.filter(d => d.team === team).length;
            return <option key={team} value={team}>{team} ({count})</option>;
          })}
        </select>

        <select
          value={sortBy}
          onChange={(e) => setSortBy(e.target.value)}
          style={{
            padding: '0.75rem',
            border: '1px solid #d1d5db',
            borderRadius: '8px',
            fontSize: '1rem',
            fontFamily: 'inherit'
          }}
        >
          <option value="employee_name">Sort by Name</option>
          <option value="working_days">Sort by Working Days</option>
          <option value="no_activities">Sort by Activities</option>
          <option value="coaching_days">Sort by Coaching</option>
        </select>
      </div>

      <div style={{
        background: 'white',
        borderRadius: '12px',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
        overflow: 'hidden'
      }}>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.95rem' }}>
            <thead style={{ background: '#f3f4f6', borderBottom: '2px solid #e5e7eb' }}>
              <tr>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Name</th>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Team</th>
                <th style={{ padding: '1rem', textAlign: 'center', fontWeight: '600', color: '#1f2937' }}>Working Days</th>
                <th style={{ padding: '1rem', textAlign: 'center', fontWeight: '600', color: '#1f2937' }}>Activities</th>
                <th style={{ padding: '1rem', textAlign: 'center', fontWeight: '600', color: '#1f2937' }}>Coaching</th>
                <th style={{ padding: '1rem', textAlign: 'center', fontWeight: '600', color: '#1f2937' }}>Field Days</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row, i) => (
                <tr
                  key={i}
                  style={{
                    borderBottom: '1px solid #e5e7eb',
                    background: i % 2 === 0 ? 'white' : '#f9fafb'
                  }}
                >
                  <td style={{ padding: '1rem', color: '#1f2937', fontWeight: '500' }}>{row.user_name}</td>
                  <td style={{ padding: '1rem', color: '#6b7280' }}>{row.team}</td>
                  <td style={{ padding: '1rem', textAlign: 'center', color: '#1f2937' }}>{row.working_days || 0}</td>
                  <td style={{ padding: '1rem', textAlign: 'center', color: '#1f2937' }}>{row.no_activities || 0}</td>
                  <td style={{ padding: '1rem', textAlign: 'center', color: '#1f2937' }}>{row.coaching_days || 0}</td>
                  <td style={{ padding: '1rem', textAlign: 'center', color: '#1f2937' }}>{row.complete_field_days || 0}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// Visits Page
function VisitsPage() {
  const { auth } = useAuth();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filterSpecialty, setFilterSpecialty] = useState('');
  const [specialties, setSpecialties] = useState([]);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      const codes = await getVisibleCodes();
      const { data: visits } = await supabase
        .from('visits')
        .select('*')
        .in('employee_code', codes);

      if (visits) {
        setData(visits);
        const uniqueSpecs = [...new Set(visits.map(v => v.specialty))].filter(Boolean).sort();
        setSpecialties(uniqueSpecs);
      }
    } catch (err) {
      console.error('Error:', err);
    } finally {
      setLoading(false);
    }
  };

  const filtered = filterSpecialty
    ? data.filter(v => v.specialty === filterSpecialty)
    : data;

  const specialty = filtered.reduce((acc, v) => {
    const key = `${v.specialty} - ${v.classification}`;
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  const chartData = Object.entries(specialty)
    .map(([name, value]) => ({ name: name.substring(0, 20), value }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);

  const COLORS = ['#667eea', '#764ba2', '#f59e0b', '#10b981', '#ef4444', '#06b6d4', '#8b5cf6', '#ec4899'];

  if (loading) {
    return <div style={{ padding: '2rem', color: '#6b7280' }}>Loading...</div>;
  }

  return (
    <div>
      <h1 style={{ fontSize: '2rem', fontWeight: '600', marginBottom: '2rem', color: '#1f2937' }}>
        Visits Breakdown
      </h1>

      <div style={{ marginBottom: '2rem' }}>
        <label style={{ display: 'block', marginBottom: '0.5rem', fontWeight: '600', color: '#374151' }}>
          Filter by Specialty:
        </label>
        <select
          value={filterSpecialty}
          onChange={(e) => setFilterSpecialty(e.target.value)}
          style={{
            padding: '0.75rem',
            border: '1px solid #d1d5db',
            borderRadius: '8px',
            fontSize: '1rem',
            fontFamily: 'inherit',
            maxWidth: '300px'
          }}
        >
          <option value="">All Specialties ({data.length} visits)</option>
          {specialties.map(spec => {
            const count = data.filter(v => v.specialty === spec).length;
            return <option key={spec} value={spec}>{spec} ({count})</option>;
          })}
        </select>
      </div>

      {chartData.length > 0 && (
        <div style={{
          background: 'white',
          padding: '2rem',
          borderRadius: '12px',
          boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
          marginBottom: '2rem'
        }}>
          <h2 style={{ fontSize: '1.25rem', fontWeight: '600', marginBottom: '1rem', color: '#1f2937' }}>
            Specialty × Classification
          </h2>
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie
                data={chartData}
                cx="50%"
                cy="50%"
                labelLine={false}
                label={({ name, value }) => `${name}: ${value}`}
                outerRadius={80}
                fill="#8884d8"
                dataKey="value"
              >
                {chartData.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </div>
      )}

      <div style={{
        background: 'white',
        borderRadius: '12px',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
        overflow: 'hidden'
      }}>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.95rem' }}>
            <thead style={{ background: '#f3f4f6', borderBottom: '2px solid #e5e7eb' }}>
              <tr>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Date</th>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Doctor</th>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Specialty</th>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Classification</th>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>User</th>
              </tr>
            </thead>
            <tbody>
              {filtered.slice(0, 100).map((row, i) => (
                <tr
                  key={i}
                  style={{
                    borderBottom: '1px solid #e5e7eb',
                    background: i % 2 === 0 ? 'white' : '#f9fafb'
                  }}
                >
                  <td style={{ padding: '1rem', color: '#1f2937' }}>{row.visit_date}</td>
                  <td style={{ padding: '1rem', color: '#6b7280' }}>{row.doctor_name}</td>
                  <td style={{ padding: '1rem', color: '#6b7280', fontSize: '0.9rem' }}>{row.specialty}</td>
                  <td style={{ padding: '1rem', color: '#6b7280', fontSize: '0.9rem' }}>{row.classification}</td>
                  <td style={{ padding: '1rem', color: '#6b7280', fontSize: '0.9rem' }}>{row.user}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// Coaching Page
function CoachingPage() {
  const { auth } = useAuth();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      const codes = await getVisibleCodes();
      const { data: coaching } = await supabase
        .from('coaching_days')
        .select('*')
        .in('rep_code', codes);

      if (coaching) {
        setData(coaching);
      }
    } catch (err) {
      console.error('Error:', err);
    } finally {
      setLoading(false);
    }
  };

  if (loading) {
    return <div style={{ padding: '2rem', color: '#6b7280' }}>Loading...</div>;
  }

  return (
    <div>
      <h1 style={{ fontSize: '2rem', fontWeight: '600', marginBottom: '2rem', color: '#1f2937' }}>
        Coaching Days
      </h1>

      <div style={{
        background: 'white',
        borderRadius: '12px',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
        overflow: 'hidden'
      }}>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.95rem' }}>
            <thead style={{ background: '#f3f4f6', borderBottom: '2px solid #e5e7eb' }}>
              <tr>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Date</th>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Manager</th>
                <th style={{ padding: '1rem', textAlign: 'left', fontWeight: '600', color: '#1f2937' }}>Rep</th>
                <th style={{ padding: '1rem', textAlign: 'center', fontWeight: '600', color: '#1f2937' }}>AM Visits</th>
                <th style={{ padding: '1rem', textAlign: 'center', fontWeight: '600', color: '#1f2937' }}>AM Accompanied</th>
                <th style={{ padding: '1rem', textAlign: 'center', fontWeight: '600', color: '#1f2937' }}>PM Visits</th>
                <th style={{ padding: '1rem', textAlign: 'center', fontWeight: '600', color: '#1f2937' }}>PM Accompanied</th>
              </tr>
            </thead>
            <tbody>
              {data.map((row, i) => (
                <tr
                  key={i}
                  style={{
                    borderBottom: '1px solid #e5e7eb',
                    background: i % 2 === 0 ? 'white' : '#f9fafb'
                  }}
                >
                  <td style={{ padding: '1rem', color: '#1f2937' }}>{row.coaching_date}</td>
                  <td style={{ padding: '1rem', color: '#6b7280' }}>{row.manager_name}</td>
                  <td style={{ padding: '1rem', color: '#6b7280' }}>{row.rep_name}</td>
                  <td style={{ padding: '1rem', textAlign: 'center', color: '#1f2937' }}>{row.am_visits || 0}</td>
                  <td style={{ padding: '1rem', textAlign: 'center', color: '#1f2937' }}>{row.am_accompanied || 0}</td>
                  <td style={{ padding: '1rem', textAlign: 'center', color: '#1f2937' }}>{row.pm_visits || 0}</td>
                  <td style={{ padding: '1rem', textAlign: 'center', color: '#1f2937' }}>{row.pm_accompanied || 0}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// Export Page
function ExportPage() {
  const { auth } = useAuth();
  const [exporting, setExporting] = useState(false);

  const handleExport = async () => {
    setExporting(true);
    try {
      const codes = await getVisibleCodes();

      const [
        { data: summaries },
        { data: visits },
        { data: coaching }
      ] = await Promise.all([
        supabase.from('summaries').select('*').in('employee_code', codes),
        supabase.from('visits').select('*').in('employee_code', codes),
        supabase.from('coaching_days').select('*').in('rep_code', codes)
      ]);

      const workbook = XLSX.utils.book_new();
      if (summaries && summaries.length > 0) {
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaries), 'Summary');
      }
      if (visits && visits.length > 0) {
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(visits.slice(0, 5000)), 'Visits');
      }
      if (coaching && coaching.length > 0) {
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(coaching), 'Coaching');
      }

      const date = new Date().toISOString().split('T')[0];
      XLSX.writeFile(workbook, `Excellence-CRM-${auth.role}-${date}.xlsx`);
    } catch (err) {
      alert(`Export failed: ${err.message}`);
    } finally {
      setExporting(false);
    }
  };

  return (
    <div>
      <h1 style={{ fontSize: '2rem', fontWeight: '600', marginBottom: '2rem', color: '#1f2937' }}>
        Export Data
      </h1>

      <div style={{
        background: 'white',
        padding: '2rem',
        borderRadius: '12px',
        boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
        maxWidth: '500px'
      }}>
        <p style={{ color: '#6b7280', marginBottom: '1.5rem', lineHeight: '1.6', fontSize: '1rem' }}>
          Download all data you have access to based on your role ({auth.role.toUpperCase()}). Includes:
        </p>
        <ul style={{ color: '#6b7280', marginBottom: '2rem', paddingLeft: '1.5rem', lineHeight: '1.8' }}>
          <li>Summary statistics (working days, activities, coaching)</li>
          <li>Visit records (specialty, classification, doctor)</li>
          <li>Coaching day sessions</li>
        </ul>

        <button
          onClick={handleExport}
          disabled={exporting}
          style={{
            padding: '0.875rem 1.75rem',
            background: exporting ? '#9ca3af' : '#667eea',
            color: 'white',
            border: 'none',
            borderRadius: '8px',
            fontSize: '1rem',
            fontWeight: '600',
            cursor: exporting ? 'not-allowed' : 'pointer',
            fontFamily: 'inherit',
            transition: 'background 0.2s'
          }}
          onMouseOver={(e) => {
            if (!exporting) e.target.style.background = '#5568d3';
          }}
          onMouseOut={(e) => {
            if (!exporting) e.target.style.background = '#667eea';
          }}
        >
          {exporting ? '⏳ Exporting...' : '⬇️ Download Excel'}
        </button>
      </div>
    </div>
  );
}

// Helper: Get visible employee codes based on role
async function getVisibleCodes() {
  const auth = useContext(AuthContext);
  if (!auth) return [];

  try {
    const { data: hierarchy } = await supabase.from('hierarchy').select('*');
    if (!hierarchy) return [auth.auth.employeeCode];

    const userRow = hierarchy.find(h => h.employee_code === auth.auth.employeeCode);
    if (!userRow) return [auth.auth.employeeCode];

    switch (auth.auth.role) {
      case 'rep':
        return [auth.auth.employeeCode];

      case 'supervisor':
        return [
          auth.auth.employeeCode,
          ...hierarchy
            .filter(h => h.supervisor_name === userRow.employee_name)
            .map(h => h.employee_code)
        ];

      case 'area_manager':
        return hierarchy
          .filter(h =>
            h.supervisor_name === userRow.employee_name ||
            h.employee_code === auth.auth.employeeCode
          )
          .map(h => h.employee_code);

      case 'blm':
        return hierarchy
          .filter(h => h.team === userRow.role)
          .map(h => h.employee_code);

      case 'admin':
        return hierarchy.map(h => h.employee_code);

      default:
        return [auth.auth.employeeCode];
    }
  } catch (err) {
    console.error('Error getting visible codes:', err);
    return [auth.auth.employeeCode];
  }
}
