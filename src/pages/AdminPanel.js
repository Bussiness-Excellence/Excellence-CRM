import React, { useEffect, useState, useCallback } from 'react';
import { supabase } from '../supabaseClient';
import { useAuth } from '../contexts/AuthContext';
import './AdminPanel.css';

const ROLES = ['Admin','BLM','Area Manager','Supervisor','MR'];
const PER_PAGE = 30;

export default function AdminPanel() {
  const { profile, signOut } = useAuth();
  const [users, setUsers]       = useState([]);
  const [teams, setTeams]       = useState([]);
  const [loading, setLoading]   = useState(true);
  const [search, setSearch]     = useState('');
  const [filterTeam, setFilterTeam]   = useState('all');
  const [filterRole, setFilterRole]   = useState('all');
  const [page, setPage]         = useState(0);
  const [editing, setEditing]   = useState(null); // user being edited
  const [editRole, setEditRole] = useState('');
  const [newPass, setNewPass]   = useState('');
  const [saving, setSaving]     = useState(false);
  const [msg, setMsg]           = useState({ text:'', type:'' });

  const [periodStats, setPeriodStats] = useState([]);

  const showMsg = (text, type='ok') => {
    setMsg({text,type});
    setTimeout(()=>setMsg({text:'',type:''}), 4000);
  };

  const loadUsers = useCallback(async () => {
    setLoading(true);
    const [{ data:u }, { data:t }, { data:s }] = await Promise.all([
      supabase.from('app_users').select('*,teams(name)').order('employee_name'),
      supabase.from('teams').select('*').order('name'),
      supabase.from('summaries').select('period,team,employee_code'),
    ]);
    setUsers(u || []);
    setTeams(t || []);
    const periodsUniq = [...new Set((s||[]).map(r=>r.period))].sort().reverse();
    // Count users with data per period
    const stats = periodsUniq.map(p => ({
      period: p,
      count: (s||[]).filter(r=>r.period===p).length,
      teams: new Set((s||[]).filter(r=>r.period===p).map(r=>r.team)).size,
    }));
    setPeriodStats(stats);
    setLoading(false);
  }, []);

  useEffect(()=>{ loadUsers(); }, [loadUsers]);

  // ── Filtered user list ──
  const filtered = users.filter(u => {
    const q = search.toLowerCase();
    const matchQ = !q || u.employee_name?.toLowerCase().includes(q) || u.employee_code?.includes(q);
    const matchT = filterTeam==='all' || u.team_id === Number(filterTeam);
    const matchR = filterRole==='all' || u.role === filterRole;
    return matchQ && matchT && matchR;
  });
  const paged = filtered.slice(page*PER_PAGE, page*PER_PAGE+PER_PAGE);
  const totalPages = Math.ceil(filtered.length/PER_PAGE);

  // ── Save role ──
  async function saveRole() {
    if (!editing || !editRole) return;
    setSaving(true);
    const { error } = await supabase.from('app_users')
      .update({ role: editRole }).eq('id', editing.id);
    if (error) showMsg('Failed: '+error.message, 'err');
    else { showMsg('Role updated'); await loadUsers(); setEditing(null); }
    setSaving(false);
  }

  // ── Reset password ──
  async function resetPassword() {
    if (!editing || !newPass || newPass.length < 6) {
      showMsg('Password must be at least 6 characters', 'err'); return;
    }
    setSaving(true);
    // Use admin API via Vercel serverless function
    const resp = await fetch('/api/reset-password', {
      method:'POST',
      headers:{
        'Content-Type':'application/json',
        'Authorization':`Bearer ${(await supabase.auth.getSession()).data.session?.access_token}`,
      },
      body: JSON.stringify({ user_id: editing.id, password: newPass }),
    });
    if (resp.ok) {
      showMsg(`Password reset for ${editing.employee_name}`);
      setNewPass(''); setEditing(null);
    } else {
      const err = await resp.json().catch(()=>({}));
      showMsg('Failed: '+(err.error||resp.statusText), 'err');
    }
    setSaving(false);
  }

  // ── Toggle active ──
  async function toggleActive(u) {
    await supabase.from('app_users').update({ is_active: !u.is_active }).eq('id', u.id);
    await loadUsers();
    showMsg(u.is_active ? `${u.employee_name} deactivated` : `${u.employee_name} activated`);
  }

  return (
    <div className="admin">
      {/* Header */}
      <header className="admin-hdr">
        <div className="admin-hdr-left">
          <span className="admin-brand">EXCELLENCE</span>
          <span className="admin-title">Admin Panel</span>
        </div>
        <div className="admin-hdr-right">
          <a className="abtn abtn-primary" href="/admin/upload" style={{ marginRight: '16px' }}>↑ Upload CRM Data</a>
          <a className="abtn abtn-ghost" href="/dashboard">← Dashboard</a>
          <span className="admin-who">{profile?.employee_name}</span>
          <button className="abtn abtn-ghost" onClick={signOut}>Sign out</button>
        </div>
      </header>

      {msg.text && (
        <div className={`admin-msg ${msg.type==='err'?'msg-err':'msg-ok'}`}>{msg.text}</div>
      )}

      <div className="admin-body">
        {/* ── Left: Data periods panel ── */}
        <aside className="admin-side">
          <div className="aside-section">
            <div className="aside-title">Data Periods</div>
            {periodStats.length === 0
              ? <div className="aside-empty">No periods uploaded yet</div>
              : periodStats.map(s => (
                <div key={s.period} className="period-card">
                  <div className="period-name">{s.period}</div>
                  <div className="period-meta">{s.count} summaries · {s.teams} teams</div>
                </div>
              ))
            }
          </div>

          <div className="aside-section">
            <div className="aside-title">Teams</div>
            {teams.map(t => {
              const cnt = users.filter(u=>u.team_id===t.id).length;
              return (
                <div key={t.id} className="team-stat">
                  <span>{t.name}</span>
                  <span className="team-cnt">{cnt}</span>
                </div>
              );
            })}
          </div>

          <div className="aside-section">
            <div className="aside-title">By Role</div>
            {ROLES.map(role => {
              const cnt = users.filter(u=>u.role===role).length;
              return (
                <div key={role} className="team-stat">
                  <span>{role}</span>
                  <span className="team-cnt">{cnt}</span>
                </div>
              );
            })}
          </div>
        </aside>

        {/* ── Main: user table ── */}
        <main className="admin-main">
          <div className="admin-filters">
            <input
              className="filter-input" placeholder="Search name or code…"
              value={search} onChange={e=>{setSearch(e.target.value);setPage(0);}}
            />
            <select className="filter-sel" value={filterTeam} onChange={e=>{setFilterTeam(e.target.value);setPage(0);}}>
              <option value="all">All teams</option>
              {teams.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}
            </select>
            <select className="filter-sel" value={filterRole} onChange={e=>{setFilterRole(e.target.value);setPage(0);}}>
              <option value="all">All roles</option>
              {ROLES.map(r=><option key={r} value={r}>{r}</option>)}
            </select>
            <span className="filter-count">{filtered.length} users</span>
          </div>

          {loading ? (
            <div className="admin-loading">Loading users…</div>
          ) : (
            <>
              <table className="admin-table">
                <thead>
                  <tr>
                    <th>Code</th><th>Name</th><th>Team</th><th>Role</th>
                    <th>Active</th><th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {paged.map(u => (
                    <tr key={u.id} className={!u.is_active?'inactive':''}>
                      <td className="code-cell">{u.employee_code}</td>
                      <td className="name-cell">{u.employee_name}</td>
                      <td>{u.teams?.name || '—'}</td>
                      <td><span className={`role-badge role-${u.role?.replace(' ','-').toLowerCase()}`}>{u.role}</span></td>
                      <td>
                        <button
                          className={`toggle-btn ${u.is_active?'active':'inactive'}`}
                          onClick={()=>toggleActive(u)}
                          title={u.is_active?'Deactivate':'Activate'}
                        >{u.is_active?'✓ Active':'✗ Off'}</button>
                      </td>
                      <td>
                        <button className="abtn abtn-edit" onClick={()=>{setEditing(u);setEditRole(u.role);setNewPass('');}}>
                          Edit
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {totalPages > 1 && (
                <div className="pagination">
                  <button className="abtn abtn-ghost" disabled={page===0} onClick={()=>setPage(p=>p-1)}>← Prev</button>
                  <span>Page {page+1} of {totalPages}</span>
                  <button className="abtn abtn-ghost" disabled={page===totalPages-1} onClick={()=>setPage(p=>p+1)}>Next →</button>
                </div>
              )}
            </>
          )}
        </main>
      </div>

      {/* ── Edit modal ── */}
      {editing && (
        <div className="modal-backdrop" onClick={()=>setEditing(null)}>
          <div className="modal" onClick={e=>e.stopPropagation()}>
            <div className="modal-header">
              <div>
                <div className="modal-name">{editing.employee_name}</div>
                <div className="modal-code">{editing.employee_code} · {editing.teams?.name}</div>
              </div>
              <button className="modal-close" onClick={()=>setEditing(null)}>✕</button>
            </div>

            <div className="modal-section">
              <label className="modal-label">Role</label>
              <select className="filter-sel w100" value={editRole} onChange={e=>setEditRole(e.target.value)}>
                {ROLES.map(r=><option key={r} value={r}>{r}</option>)}
              </select>
              <button className="abtn abtn-primary mt8" onClick={saveRole} disabled={saving}>
                {saving?'Saving…':'Save Role'}
              </button>
            </div>

            <div className="modal-divider"/>

            <div className="modal-section">
              <label className="modal-label">Reset Password</label>
              <input
                className="filter-input w100" type="password"
                placeholder="New password (min 6 chars)"
                value={newPass} onChange={e=>setNewPass(e.target.value)}
                onKeyDown={e=>e.key==='Enter'&&resetPassword()}
              />
              <button className="abtn abtn-danger mt8" onClick={resetPassword} disabled={saving||newPass.length<6}>
                {saving?'Resetting…':'Reset Password'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}