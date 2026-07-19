import React, { useState, useMemo } from 'react';
import { 
  BarChart, Bar, LineChart, Line, AreaChart, Area, 
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer 
} from 'recharts';
import './ChartBuilder.css';

const METRICS = [
  { key: 'am_calls', label: 'AM Calls' },
  { key: 'pm_calls', label: 'PM Calls' },
  { key: 'working_days', label: 'Working Days' },
  { key: 'complete_field_days', label: 'Field Days' },
  { key: 'double_visit_days', label: 'Double Visits' },
  { key: 'coaching_days', label: 'Coaching Days' },
  { key: 'office_work_days', label: 'Office Work' },
  { key: 'total_am_covered', label: 'AM Covered (Centers/Hospitals)' },
  { key: 'total_pm_covered', label: 'PM Covered (Clinics)' },
  { key: 'pharmacies_visited', label: 'Pharmacies Visited' },
  { key: 'total_product_calls', label: 'Total Product Calls' },
];

const DIMENSIONS = [
  { key: 'user_name', label: 'By Rep' },
  { key: 'team', label: 'By Team' },
  { key: 'territory', label: 'By Territory' }
];

const CHART_TYPES = ['Bar', 'Area', 'Line'];

export default function ChartBuilder({ data, isManager }) {
  const [metric, setMetric] = useState('am_calls');
  const [dimension, setDimension] = useState('user_name');
  const [chartType, setChartType] = useState('Bar');

  // Aggregate data based on dimension
  const chartData = useMemo(() => {
    if (!data || data.length === 0) return [];
    
    const map = {};
    data.forEach(row => {
      let dimValue = row[dimension];
      if (!dimValue || dimValue === '') dimValue = 'Unknown';
      
      if (!map[dimValue]) {
        map[dimValue] = { name: dimValue };
      }
      
      const val = row[metric];
      if (typeof val === 'number') {
        map[dimValue][metric] = (map[dimValue][metric] || 0) + val;
      } else {
        map[dimValue][metric] = (map[dimValue][metric] || 0) + (parseFloat(val) || 0);
      }
    });

    // Convert to array and sort by metric descending
    return Object.values(map)
      .sort((a, b) => (b[metric] || 0) - (a[metric] || 0))
      .slice(0, 30); // limit to top 30 to avoid over-crowded charts
  }, [data, dimension, metric]);

  const activeMetricLabel = METRICS.find(m => m.key === metric)?.label || metric;

  const renderChart = () => {
    if (chartData.length === 0) {
      return <div className="chart-empty">No data available for the selected filters.</div>;
    }

    const commonProps = {
      data: chartData,
      margin: { top: 20, right: 30, left: 20, bottom: 60 }
    };

    const axes = (
      <>
        <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" vertical={false} />
        <XAxis 
          dataKey="name" 
          stroke="#94a3b8" 
          tick={{fill: '#94a3b8', fontSize: 12}} 
          angle={-45} 
          textAnchor="end" 
          height={80} 
        />
        <YAxis 
          stroke="#94a3b8" 
          tick={{fill: '#94a3b8', fontSize: 12}} 
          tickFormatter={(value) => value.toLocaleString()}
        />
        <Tooltip 
          contentStyle={{ backgroundColor: 'rgba(15,23,42,0.9)', border: '1px solid rgba(255,255,255,0.1)', borderRadius: '8px', color: '#fff' }}
          itemStyle={{ color: '#60a5fa', fontWeight: 'bold' }}
        />
        <Legend wrapperStyle={{ paddingTop: '20px' }} />
      </>
    );

    switch (chartType) {
      case 'Line':
        return (
          <LineChart {...commonProps}>
            {axes}
            <Line type="monotone" dataKey={metric} name={activeMetricLabel} stroke="#a78bfa" strokeWidth={3} dot={{ r: 4, fill: '#a78bfa', strokeWidth: 2 }} activeDot={{ r: 6 }} />
          </LineChart>
        );
      case 'Area':
        return (
          <AreaChart {...commonProps}>
            <defs>
              <linearGradient id="colorMetric" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor="#60a5fa" stopOpacity={0.8}/>
                <stop offset="95%" stopColor="#60a5fa" stopOpacity={0}/>
              </linearGradient>
            </defs>
            {axes}
            <Area type="monotone" dataKey={metric} name={activeMetricLabel} stroke="#60a5fa" strokeWidth={2} fillOpacity={1} fill="url(#colorMetric)" />
          </AreaChart>
        );
      case 'Bar':
      default:
        return (
          <BarChart {...commonProps}>
            {axes}
            <Bar dataKey={metric} name={activeMetricLabel} fill="#60a5fa" radius={[4, 4, 0, 0]} />
          </BarChart>
        );
    }
  };

  return (
    <div className="cb-container">
      <div className="cb-controls">
        <div className="cb-group">
          <label>Metric (Y-Axis)</label>
          <select value={metric} onChange={(e) => setMetric(e.target.value)}>
            {METRICS.map(m => <option key={m.key} value={m.key}>{m.label}</option>)}
          </select>
        </div>

        <div className="cb-group">
          <label>Dimension (X-Axis)</label>
          <select value={dimension} onChange={(e) => setDimension(e.target.value)}>
            {DIMENSIONS.map(d => <option key={d.key} value={d.key}>{d.label}</option>)}
          </select>
        </div>

        <div className="cb-group">
          <label>Chart Type</label>
          <div className="cb-type-toggles">
            {CHART_TYPES.map(type => (
              <button 
                key={type} 
                className={`cb-btn ${chartType === type ? 'active' : ''}`}
                onClick={() => setChartType(type)}
              >
                {type}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div className="cb-chart-area">
        <ResponsiveContainer width="100%" height={400}>
          {renderChart()}
        </ResponsiveContainer>
      </div>
    </div>
  );
}
