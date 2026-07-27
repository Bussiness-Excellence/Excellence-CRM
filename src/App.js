import React from 'react';
import { HashRouter, Routes, Route, Navigate } from 'react-router-dom';
import { AuthProvider, useAuth } from './contexts/AuthContext';
import { ToastProvider } from './contexts/ToastContext';
import ProtectedRoute from './components/ProtectedRoute';
import Login from './components/Login';
import Dashboard from './pages/Dashboard';
import AdminPanel from './pages/AdminPanel';
import ChangePassword from './components/ChangePassword';

import './App.css';

function LoginRoute() {
  const { session, profile, loading } = useAuth();
  if (loading) return <div className="app-loading"><div className="app-spinner"/></div>;
  if (session && profile) return <Navigate to="/dashboard" replace />;
  return <Login />;
}

function RootRoute() {
  const { session, profile, loading } = useAuth();
  if (loading) return <div className="app-loading"><div className="app-spinner"/></div>;
  if (!session || !profile) return <Navigate to="/login" replace />;
  // All roles (including Admin) land on the dashboard by default
  return <Navigate to="/dashboard" replace />;
}

function AppContent() {
  const { session, profile, loading } = useAuth();
  
  if (loading) {
    return (
      <div className="app-loading" style={{display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh'}}>
        <div className="app-spinner" style={{
          width: '40px', height: '40px', borderRadius: '50%', border: '3px solid rgba(0,0,0,0.1)', borderTopColor: '#3b82f6', animation: 'spin 1s linear infinite'
        }}/>
      </div>
    );
  }

  // If logged in but needs to change default password
  if (session && profile?.is_default_password) {
    return (
      <HashRouter>
        <Routes>
          <Route path="*" element={<ChangePassword />} />
        </Routes>
      </HashRouter>
    );
  }

  return (
    <HashRouter>
      <Routes>
        <Route path="/login" element={<LoginRoute />} />
        <Route path="/dashboard" element={
          <ProtectedRoute>
            <Dashboard />
          </ProtectedRoute>
        } />
        <Route path="/admin" element={
          <ProtectedRoute adminOnly={true}>
            <AdminPanel />
          </ProtectedRoute>
        } />
        <Route path="/" element={<RootRoute />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </HashRouter>
  );
}

export default function App() {
  return (
    <ToastProvider>
      <AuthProvider>
        <AppContent />
      </AuthProvider>
    </ToastProvider>
  );
}
