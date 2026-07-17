import React, { createContext, useContext, useEffect, useState, useCallback } from 'react';
import { supabase, employeeCodeToEmail } from '../supabaseClient';

const AuthContext = createContext(null);

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used inside <AuthProvider>');
  return ctx;
}

/**
 * Given the logged-in person's app_users row and the full hierarchy table,
 * compute which employee_codes they're allowed to see. This mirrors what
 * the RLS policies will eventually enforce server-side — having the same
 * logic here means the UI behaves correctly even before those policies
 * exist, and RLS becomes a second, server-enforced layer of the same rule
 * rather than the only layer.
 *
 *   MR            → only themselves
 *   Supervisor     → themselves + every rep whose supervisor_name matches theirs
 *   Area Manager   → themselves + their supervisors + every rep under those supervisors
 *   BLM            → everyone on their team
 *   Admin          → everyone, across all teams
 */
export function computeVisibleEmployeeCodes(profile, hierarchyRows) {
  if (!profile) return [];
  const { role, employee_name: myName, team_id: myTeamId, employee_code: myCode } = profile;

  if (role === 'Admin') {
    return hierarchyRows.map((h) => h.employee_code).filter(Boolean);
  }

  if (role === 'BLM') {
    return hierarchyRows
      .filter((h) => h.team_id === myTeamId)
      .map((h) => h.employee_code)
      .filter(Boolean);
  }

  if (role === 'Area Manager') {
    // Reps/supervisors whose chain leads up through this area manager.
    // The hierarchy rows store each person's direct supervisor_name; an
    // Area Manager oversees the supervisors reporting through them, plus
    // every rep under those supervisors.
    const supervisorNames = new Set(
      hierarchyRows
        .filter((h) => h.role === 'Supervisor' && h.team_id === myTeamId)
        .map((h) => h.employee_name)
    );
    const visible = hierarchyRows
      .filter((h) => h.team_id === myTeamId &&
        (h.employee_name === myName || supervisorNames.has(h.supervisor_name)))
      .map((h) => h.employee_code);
    return [...new Set([myCode, ...visible])].filter(Boolean);
  }

  if (role === 'Supervisor') {
    const visible = hierarchyRows
      .filter((h) => h.supervisor_name === myName || h.employee_name === myName)
      .map((h) => h.employee_code);
    return [...new Set([myCode, ...visible])].filter(Boolean);
  }

  // MR (default): only their own data
  return [myCode].filter(Boolean);
}

export function AuthProvider({ children }) {
  const [session, setSession] = useState(null);
  const [profile, setProfile] = useState(null);       // app_users row
  const [hierarchy, setHierarchy] = useState([]);      // full hierarchy table
  const [visibleCodes, setVisibleCodes] = useState([]);
  const [loading, setLoading] = useState(true);

  const loadProfileAndHierarchy = useCallback(async (userId) => {
    const [{ data: profileRow, error: profileErr },
           { data: hierarchyRows, error: hierErr }] = await Promise.all([
      supabase.from('app_users').select('*').eq('id', userId).single(),
      supabase.from('hierarchy').select('*'),
    ]);

    if (profileErr) {
      // eslint-disable-next-line no-console
      console.error('Failed to load profile:', profileErr.message);
      setProfile(null);
      setVisibleCodes([]);
      return;
    }
    if (hierErr) {
      // eslint-disable-next-line no-console
      console.error('Failed to load hierarchy:', hierErr.message);
    }

    setProfile(profileRow);
    setHierarchy(hierarchyRows || []);
    setVisibleCodes(computeVisibleEmployeeCodes(profileRow, hierarchyRows || []));
  }, []);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      if (session) {
        loadProfileAndHierarchy(session.user.id).finally(() => setLoading(false));
      } else {
        setLoading(false);
      }
    });

    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      if (session) {
        setLoading(true);
        loadProfileAndHierarchy(session.user.id).finally(() => setLoading(false));
      } else {
        setProfile(null);
        setHierarchy([]);
        setVisibleCodes([]);
      }
    });

    return () => listener.subscription.unsubscribe();
  }, [loadProfileAndHierarchy]);

  const signInWithEmployeeCode = useCallback(async (employeeCode, password) => {
    const email = employeeCodeToEmail(employeeCode);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    return { error };
  }, []);

  const signOut = useCallback(() => supabase.auth.signOut(), []);

  const value = {
    session,
    profile,
    hierarchy,
    visibleCodes,
    loading,
    signInWithEmployeeCode,
    signOut,
  };

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}
