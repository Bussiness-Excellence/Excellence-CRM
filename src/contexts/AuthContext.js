import React, { createContext, useContext, useEffect, useState, useCallback } from 'react';
import { supabase, employeeCodeToEmail } from '../supabaseClient';

const AuthContext = createContext(null);

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used inside <AuthProvider>');
  return ctx;
}

/**
 * Compute which employee_codes this user can see, based on the hierarchy chain.
 *
 * How the hierarchy table works:
 *   - Every row has: employee_name, employee_code, role, supervisor_name, team_id
 *   - MRs:           supervisor_name = their supervisor's name
 *   - Supervisors:   supervisor_name = their own name (self-referential)
 *   - Area Managers: supervisor_name = their own name
 *   - BLMs:          supervisor_name = null
 *
 * Visibility rules:
 *   MR           → own code only
 *   Supervisor   → own code + codes of all MRs whose supervisor_name = my name
 *   Area Manager → own code + codes of supervisors in my team +
 *                  codes of all MRs under those supervisors
 *                  BUT only supervisors that appear under ME in the hierarchy,
 *                  detected by: MRs whose supervisor_name matches a supervisor
 *                  who is in the same team and whose area_manager_name = my name
 *                  (fallback: ALL supervisors in team if area_manager_name not set)
 *   BLM          → everyone in my team
 *   Admin        → everyone
 */
export function computeVisibleEmployeeCodes(profile, hierarchyRows) {
  if (!profile) return [];
  const { role, employee_name: myName, team_id: myTeamId, employee_code: myCode } = profile;

  if (role === 'Admin') {
    return hierarchyRows.map(h => h.employee_code).filter(Boolean);
  }

  if (role === 'BLM') {
    return hierarchyRows
      .filter(h => h.team_id === myTeamId)
      .map(h => h.employee_code)
      .filter(Boolean);
  }

  if (role === 'Area Manager') {
    const myTeamRows = hierarchyRows.filter(h => h.team_id === myTeamId);

    // Find supervisors that report to me via area_manager_name (if populated)
    // OR fall back to: supervisors whose MRs' supervisor_name chains include me
    const hasMappedAMs = myTeamRows.some(h => h.area_manager_name);

    let mySupervisorNames;
    if (hasMappedAMs) {
      // Use the area_manager_name column directly
      mySupervisorNames = new Set(
        myTeamRows
          .filter(h => h.area_manager_name === myName && h.role === 'Supervisor')
          .map(h => h.employee_name)
      );
    } else {
      // Fallback: all supervisors in the team (old behaviour — shows too much but
      // better than showing nothing while area_manager_name is being populated)
      mySupervisorNames = new Set(
        myTeamRows.filter(h => h.role === 'Supervisor').map(h => h.employee_name)
      );
    }

    // Include self, all my supervisors, and all MRs under my supervisors
    const visible = myTeamRows
      .filter(h =>
        h.employee_name === myName ||
        mySupervisorNames.has(h.employee_name) ||
        mySupervisorNames.has(h.supervisor_name)
      )
      .map(h => h.employee_code);

    return [...new Set([myCode, ...visible])].filter(Boolean);
  }

  if (role === 'Supervisor') {
    // My own rows + all MRs whose supervisor_name = my name
    const visible = hierarchyRows
      .filter(h => h.employee_name === myName || h.supervisor_name === myName)
      .map(h => h.employee_code);
    return [...new Set([myCode, ...visible])].filter(Boolean);
  }

  // MR default
  return [myCode].filter(Boolean);
}

export function AuthProvider({ children }) {
  const [session, setSession]         = useState(null);
  const [profile, setProfile]         = useState(null);
  const [hierarchy, setHierarchy]     = useState([]);
  const [visibleCodes, setVisibleCodes] = useState([]);
  const [loading, setLoading]         = useState(true);

  const loadProfileAndHierarchy = useCallback(async (userId) => {
    const [{ data: profileRow, error: profileErr }, { data: hierarchyRows }] =
      await Promise.all([
        supabase.from('app_users').select('*').eq('id', userId).single(),
        supabase.from('hierarchy').select('*'),
      ]);

    if (profileErr) {
      console.error('Profile load failed:', profileErr.message);
      setProfile(null); setVisibleCodes([]);
      return;
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
        setProfile(null); setHierarchy([]); setVisibleCodes([]);
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

  const updatePassword = useCallback(async (newPassword) => {
    const { error } = await supabase.auth.updateUser({ password: newPassword });
    return { error };
  }, []);

  return (
    <AuthContext.Provider value={{
      session, profile, hierarchy, visibleCodes, loading,
      signInWithEmployeeCode, signOut, updatePassword,
    }}>
      {children}
    </AuthContext.Provider>
  );
}
