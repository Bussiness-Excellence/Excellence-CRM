import React from 'react';
import { HashRouter, Routes, Route, Navigate } from 'react-router-dom';
import { AuthProvider, useAuth } from './contexts/AuthContext';
import { ToastProvider } from './contexts/ToastContext';
import Login from './components/Login';
import ChangePassword from './components/ChangePassword';
import ProtectedRoute from './components/ProtectedRoute';
import Dashboard from './pages/Dashboard';
import AdminPanel from './pages/AdminPanel';
import './App.css';

function LoginRoute() {
  const { session, loading } = useAuth();
  if (loading) return <div className="app-loading"><div className="app-spinner"/></div>;
  if (session) return <Navigate to="/" replace />;
  return <Login />;
}

function RootRoute() {
  const { profile, loading } = useAuth();
  if (loading) return <div className="app-loading"><div className="app-spinner"/></div>;
  if (!profile) return <Navigate to="/login" replace />;
  // All roles (including Admin) land on the dashboard by default
  return <Navigate to="/dashboard" replace />;
}

/**
 * Wraps the entire app — if the logged-in user still has a default password,
 * force them to the ChangePassword screen regardless of which route they hit.
 */
function AppContent() {
  const { session, profile, loading } = useAuth();
  if (loading) return <div className="app-loading"><div className="app-spinner"/></div>;
  if (session && profile?.is_default_password) return <ChangePassword />;
  return (
    <Routes>
      <Route path="/login" element={<LoginRoute />} />
      <Route path="/" element={<ProtectedRoute><RootRoute /></ProtectedRoute>} />
      <Route path="/dashboard" element={<ProtectedRoute><Dashboard /></ProtectedRoute>} />
      <Route path="/admin" element={<ProtectedRoute adminOnly><AdminPanel /></ProtectedRoute>} />
    </Routes>
  );
}

export default function App() {
  React.useEffect(() => {
    const theme = localStorage.getItem('theme') || 'light';
    if (theme === 'dark') {
      document.documentElement.classList.add('dark');
      document.documentElement.classList.remove('light');
    } else {
      document.documentElement.classList.add('light');
      document.documentElement.classList.remove('dark');
    }
  }, []);

  return (
    <AuthProvider>
      <ToastProvider>
        <HashRouter>
          <AppContent />
        </HashRouter>
      </ToastProvider>
    </AuthProvider>
  );
}
