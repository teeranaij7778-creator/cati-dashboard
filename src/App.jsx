import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell, PieChart, Pie
} from 'recharts';
import { 
  Users, CheckCircle, AlertTriangle, XCircle, Search, 
  FileText, BarChart2, MessageSquare, Calendar, TrendingUp, Database, Link, RefreshCw, Trash2, Globe, FilterX, PlayCircle, UserCheck, Settings, AlertCircle, Info, ChevronRight, ExternalLink, User, ChevronDown, CheckSquare, Square, X, Briefcase, Lock, LogIn, Activity, Filter, Check, Clock, ListChecks, Award
} from 'lucide-react';

/** * CATI CES 2026 Analytics Dashboard - INTAGE BLACK & RED Edition
 * ระบบวิเคราะห์ผลการตรวจ QC งานสัมภาษณ์ (CATI) พร้อมระบบ Auto-Refresh ทุก 5 นาที
 * เพิ่มการแสดงผลการประเมิน 13 หัวข้อ (P:AB)
 */

const DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSHePu18q6f93lQqVW5_JNv6UygyYRGNjT5qOq4nSrROCnGxt1pkdgiPT91rm-_lVpku-PW-LWs-ufv/pub?gid=470556665&single=true&output=csv"; 

const COLORS = {
  'ดีเยี่ยม': '#3B82F6',     
  'ผ่านเกณฑ์': '#10B981',   
  'ควรปรับปรุง': '#EF4444', 
  'พบข้อผิดพลาด': '#EF4444', 
  'ไม่ผ่านเกณฑ์': '#B91C1C', 
};

const RESULT_ORDER = ['ดีเยี่ยม', 'ผ่านเกณฑ์', 'ควรปรับปรุง', 'พบข้อผิดพลาด', 'ไม่ผ่านเกณฑ์'];

const IntageLogo = ({ className = "h-8" }) => (
  <div className={`flex items-center gap-2 ${className}`}>
    <div className="w-4 h-4 rounded-full bg-red-600 animate-pulse"></div>
    <span className="font-black tracking-[0.15em] text-white italic text-lg">INTAGE</span>
  </div>
);

const parseCSV = (text) => {
  const result = [];
  let row = [];
  let cell = '';
  let inQuotes = false;
  if (!text) return [];
  
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];
    if (inQuotes) {
      if (char === '"' && nextChar === '"') { cell += '"'; i++; }
      else if (char === '"') inQuotes = false;
      else cell += char;
    } else {
      if (char === '"') inQuotes = true;
      else if (char === ',') { row.push(cell); cell = ''; }
      else if (char === '\r' || char === '\n') {
        row.push(cell);
        if (row.length > 1 || row[0] !== '') result.push(row);
        row = []; cell = '';
        if (char === '\r' && nextChar === '\n') i++;
      } else cell += char;
    }
  }
  if (cell || row.length > 0) { row.push(cell); result.push(row); }
  return result;
};

const App = () => {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [inputUser, setInputUser] = useState('');
  const [inputPass, setInputPass] = useState('');
  const [loginError, setLoginError] = useState('');
  const [data, setData] = useState([]);
  const [evalHeaders, setEvalHeaders] = useState([]); // เก็บหัวข้อการประเมิน 13 ข้อ
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [sheetUrl, setSheetUrl] = useState(localStorage.getItem('qc_sheet_url') || DEFAULT_SHEET_URL);
  const [searchTerm, setSearchTerm] = useState('');
  const [isFilterSidebarOpen, setIsFilterSidebarOpen] = useState(false);
  const [lastUpdated, setLastUpdated] = useState(null);
  
  const [filterSup, setFilterSup] = useState('All');
  const [selectedResults, setSelectedResults] = useState([]);
  const [selectedTypes, setSelectedTypes] = useState([]);
  const [selectedAgents, setSelectedAgents] = useState([]); 
  const [selectedMonth, setSelectedMonth] = useState('All');
  const [showSync, setShowSync] = useState(!localStorage.getItem('qc_sheet_url') && !DEFAULT_SHEET_URL);
  
  const [activeCell, setActiveCell] = useState({ agent: null, resultType: null });
  const [expandedCaseId, setExpandedCaseId] = useState(null); // ID ของเคสที่คลิกดูรายละเอียด

  useEffect(() => {
    if (!document.getElementById('thai-font-link')) {
      const fontLink = document.createElement('link');
      fontLink.id = 'thai-font-link';
      fontLink.rel = 'stylesheet';
      fontLink.href = 'https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,100;0,200;0,300;0,400;0,500;0,600;0,700;0,800;1,100&display=swap';
      document.head.appendChild(fontLink);
    }
    if (!document.getElementById('tailwind-script')) {
      const script = document.createElement('script');
      script.id = 'tailwind-script';
      script.src = "https://cdn.tailwindcss.com";
      script.async = true;
      const configScript = document.createElement('script');
      configScript.innerHTML = `
        tailwind.config = {
          theme: {
            extend: {
              fontFamily: {
                sans: ['Sarabun', 'sans-serif'],
              }
            }
          }
        }
      `;
      document.head.appendChild(configScript);
      document.head.appendChild(script);
    }
  }, []);

  // --- Auto-Refresh Logic (Every 5 Minutes) ---
  useEffect(() => {
    let intervalId;
    if (isAuthenticated && sheetUrl && sheetUrl.includes('http')) {
      fetchFromSheet(sheetUrl);
      intervalId = setInterval(() => {
        fetchFromSheet(sheetUrl);
      }, 300000);
    }
    return () => {
      if (intervalId) clearInterval(intervalId);
    };
  }, [sheetUrl, isAuthenticated]);

  const handleLogin = (e) => {
    e.preventDefault();
    if (inputUser === 'Admin' && inputPass === '1234') {
      setIsAuthenticated(true);
      setLoginError('');
    } else {
      setLoginError('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    }
  };

  const fetchFromSheet = async (urlToFetch) => {
    let finalUrl = urlToFetch.trim();
    if (finalUrl.includes('docs.google.com/spreadsheets/d/') && !finalUrl.includes('pub?')) {
        const idMatch = finalUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (idMatch) {
            finalUrl = `https://docs.google.com/spreadsheets/d/e/${idMatch[1]}/pub?output=csv`;
        }
    }
    setLoading(true);
    setError(null);
    try {
      const fetchUrl = `${finalUrl}&t=${new Date().getTime()}`;
      const response = await fetch(fetchUrl);
      if (!response.ok) throw new Error("ไม่สามารถเข้าถึงไฟล์ได้");
      const csvText = await response.text();
      const allRows = parseCSV(csvText);
      let headerIdx = allRows.findIndex(row => row.some(cell => cell.toString().toLowerCase().includes("interviewer") || cell.toString().includes("สรุปผล")));
      
      if (headerIdx === -1) throw new Error("ไม่พบคอลัมน์ข้อมูลที่กำหนด");

      const headers = allRows[headerIdx].map(h => h.trim());
      
      // เก็บหัวข้อการประเมิน 13 ข้อ (P:AB คือ index 15 ถึง 27)
      const evaluationsList = headers.slice(15, 28);
      setEvalHeaders(evaluationsList);

      const getIdx = (name) => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
      const idx = {
        year: getIdx("Year"), month: getIdx("เดือน"), date: getIdx("วันที่สัมภาษณ์"), touchpoint: getIdx("TOUCH_POINT"), type: getIdx("AC / BC"), sup: getIdx("Supervisor"), agent: getIdx("Interviewer"),
        agentName: headers.findIndex(h => h.toLowerCase().includes("interviewer name") || h.toLowerCase().includes("ชื่อ-นามสกุล") || h.toLowerCase().includes("ชื่อพนักงาน")),
        audio: getIdx("ไฟล์เสียง"), result: getIdx("สรุปผลการสัมภาษณ์"), comment: getIdx("Comment")
      };
      
      const parsedData = allRows.slice(headerIdx + 1).filter(row => {
          const agentCode = row[idx.agent]?.toString().trim() || "";
          if (agentCode === "" || agentCode.includes("#N/A")) return false;
          return true;
        })
        .map((row, index) => {
          let rawResult = row[idx.result]?.toString().trim() || "N/A";
          let cleanResult = "N/A";
          if (rawResult.includes("ไม่ผ่านเกณฑ์")) cleanResult = "ไม่ผ่านเกณฑ์";
          else if (rawResult.includes("ดีเยี่ยม")) cleanResult = "ดีเยี่ยม";
          else if (rawResult.includes("ผ่านเกณฑ์") || rawResult.includes("ผ่าน%")) cleanResult = "ผ่านเกณฑ์";
          else if (rawResult.includes("ควรปรับปรุง")) cleanResult = "ควรปรับปรุง";
          else if (rawResult.includes("พบข้อผิดพลาด")) cleanResult = "พบข้อผิดพลาด";
          
          const agentId = row[idx.agent]?.toString().trim() || 'Unknown';
          const agentName = (idx.agentName !== -1) ? row[idx.agentName]?.toString().trim() : '';
          let displayAgent = agentName && agentName !== agentId ? `${agentId} : ${agentName}` : agentId;
          
          // ดึงคะแนน 13 หัวข้อ
          const evaluations = evaluationsList.map((header, i) => ({
            label: header,
            value: row[15 + i] || '-'
          }));

          return {
            id: index, year: row[idx.year] || 'N/A', month: row[idx.month] || 'N/A', date: row[idx.date] || 'N/A', touchpoint: row[idx.touchpoint] || 'N/A', type: row[idx.type] || 'N/A', supervisor: row[idx.sup] || 'N/A', agent: displayAgent, audio: row[idx.audio] || '', result: cleanResult, comment: row[idx.comment] || '',
            evaluations: evaluations
          };
        });
      setData(parsedData);
      setLastUpdated(new Date().toLocaleTimeString('th-TH'));
      localStorage.setItem('qc_sheet_url', finalUrl);
      setShowSync(false);
    } catch (err) { setError(err.message); } finally { setLoading(false); }
  };

  const clearConnection = () => {
    if(window.confirm("ต้องการยกเลิกการเชื่อมต่อ?")) {
      localStorage.removeItem('qc_sheet_url'); setSheetUrl(""); setData([]); setShowSync(true);
    }
  };

  const availableMonths = useMemo(() => [...new Set(data.map(d => d.month).filter(m => m && m !== 'N/A'))], [data]);
  const availableSups = useMemo(() => [...new Set(data.map(d => d.supervisor).filter(s => s && s !== 'N/A'))].sort(), [data]);
  const availableTypes = useMemo(() => [...new Set(data.map(d => d.type).filter(t => t && t !== 'N/A' && t !== ''))].sort(), [data]);
  const availableAgents = useMemo(() => {
    let filtered = data;
    if (filterSup !== 'All') filtered = filtered.filter(d => d.supervisor === filterSup);
    return [...new Set(filtered.map(d => d.agent).filter(a => a && a !== 'Unknown'))].sort();
  }, [data, filterSup]);

  const filteredData = useMemo(() => {
    return data.filter(item => {
      const matchesSearch = item.agent.toLowerCase().includes(searchTerm.toLowerCase()) || item.comment.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesResult = selectedResults.length === 0 || selectedResults.includes(item.result);
      const matchesACBC = selectedTypes.length === 0 || selectedTypes.includes(item.type);
      const matchesSup = filterSup === 'All' || item.supervisor === filterSup;
      const matchesAgent = selectedAgents.length === 0 || selectedAgents.includes(item.agent);
      const matchesMonth = selectedMonth === 'All' || item.month === selectedMonth;
      return matchesSearch && matchesResult && matchesACBC && matchesSup && matchesAgent && matchesMonth;
    });
  }, [data, searchTerm, selectedResults, selectedTypes, filterSup, selectedAgents, selectedMonth]);

  const agentSummary = useMemo(() => {
    const summaryMap = {};
    filteredData.forEach(item => {
      if (!summaryMap[item.agent]) {
        summaryMap[item.agent] = { name: item.agent, 'ดีเยี่ยม': 0, 'ผ่านเกณฑ์': 0, 'ควรปรับปรุง': 0, 'พบข้อผิดพลาด': 0, 'ไม่ผ่านเกณฑ์': 0, total: 0 };
      }
      if (summaryMap[item.agent][item.result] !== undefined) summaryMap[item.agent][item.result] += 1;
      summaryMap[item.agent].total += 1;
    });
    return Object.values(summaryMap).sort((a, b) => b.total - a.total);
  }, [filteredData]);

  const chartData = useMemo(() => {
    const total = filteredData.length;
    const counts = { 'ดีเยี่ยม': 0, 'ผ่านเกณฑ์': 0, 'ควรปรับปรุง': 0, 'พบข้อผิดพลาด': 0, 'ไม่ผ่านเกณฑ์': 0 };
    filteredData.forEach(d => { if(counts[d.result] !== undefined) counts[d.result]++; });
    return Object.keys(counts).map(key => ({ name: key, count: counts[key], percent: total > 0 ? ((counts[key] / total) * 100).toFixed(1) : 0, color: COLORS[key] }));
  }, [filteredData]);

  const detailLogs = useMemo(() => (activeCell.agent && activeCell.resultType) ? filteredData.filter(d => d.agent === activeCell.agent && d.result === activeCell.resultType) : filteredData, [filteredData, activeCell]);
  const passRate = useMemo(() => filteredData.length === 0 ? 0 : ((filteredData.filter(d => ['ดีเยี่ยม', 'ผ่านเกณฑ์'].includes(d.result)).length / filteredData.length) * 100).toFixed(1), [filteredData]);

  const handleMatrixClick = (agentName, type) => {
    if (activeCell.agent === agentName && activeCell.resultType === type) setActiveCell({ agent: null, resultType: null });
    else { setActiveCell({ agent: agentName, resultType: type }); document.getElementById('detail-section')?.scrollIntoView({ behavior: 'smooth' }); }
  };

  const hasActiveFilters = filterSup !== 'All' || selectedAgents.length > 0 || selectedResults.length > 0 || selectedTypes.length > 0 || selectedMonth !== 'All';

  // --- Login ---
  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-zinc-950 flex items-center justify-center p-4 font-sans text-white">
        <div className="bg-zinc-900 p-8 md:p-12 rounded-[2.5rem] shadow-2xl w-full max-w-md border border-zinc-800 relative overflow-hidden">
          <div className="absolute top-0 left-0 w-full h-1 bg-red-600"></div>
          <div className="flex justify-center mb-10">
            <IntageLogo className="scale-150" />
          </div>
          <h2 className="text-xl font-black text-center text-zinc-100 uppercase tracking-[0.2em] mb-2 italic">CATI CES 2026</h2>
          <p className="text-center text-red-500 text-xs font-bold uppercase tracking-widest mb-10">QC Report</p>
          <form onSubmit={handleLogin} className="space-y-4">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest ml-2 pl-2">Username</label>
              <div className="relative">
                <User className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-300" size={18} />
                <input type="text" value={inputUser} onChange={(e) => setInputUser(e.target.value)} className="w-full pl-12 pr-4 py-4 bg-zinc-800 border border-zinc-700 rounded-2xl font-bold text-white outline-none focus:ring-2 focus:ring-red-600 transition-all" placeholder="Admin" />
              </div>
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest ml-2 pl-2">Password</label>
              <div className="relative">
                <Lock className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-300" size={18} />
                <input type="password" value={inputPass} onChange={(e) => setInputPass(e.target.value)} className="w-full pl-12 pr-4 py-4 bg-zinc-800 border border-zinc-700 rounded-2xl font-bold text-white outline-none focus:ring-2 focus:ring-red-600 transition-all" placeholder="••••" />
              </div>
            </div>
            {loginError && <div className="flex items-center gap-2 text-red-500 text-xs font-bold bg-red-600/10 p-3 rounded-xl border border-red-600/20"><AlertCircle size={14} /> {loginError}</div>}
            <button type="submit" className="w-full py-4 bg-red-600 hover:bg-red-700 text-white rounded-2xl font-black uppercase tracking-wider shadow-lg shadow-red-900/20 transition-all mt-4 flex items-center justify-center gap-2 group">
              เข้าสู่ระบบ <LogIn size={18} className="group-hover:translate-x-1 transition-transform"/>
            </button>
          </form>
          <p className="mt-10 text-center text-[9px] text-zinc-300 font-bold uppercase tracking-tighter italic opacity-60">Authorized Personnel Only • Powered by INTAGE</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-zinc-950 p-4 md:p-8 font-sans text-zinc-100 relative overflow-x-hidden">
      
      {/* Sidebar Panel */}
      <div className={`fixed inset-0 z-40 bg-black/60 backdrop-blur-sm transition-opacity duration-300 ${isFilterSidebarOpen ? 'opacity-100' : 'opacity-0 pointer-events-none'}`} onClick={() => setIsFilterSidebarOpen(false)} />
      <aside className={`fixed inset-y-0 right-0 z-50 w-80 bg-zinc-900 shadow-2xl transform transition-transform duration-300 ease-in-out ${isFilterSidebarOpen ? 'translate-x-0' : 'translate-x-full'} overflow-y-auto border-l border-zinc-800`}>
          <div className="p-6">
             <div className="flex items-center justify-between mb-8 pb-4 border-b border-zinc-800">
                <div className="flex items-center gap-2">
                   <div className="p-2 bg-red-600 rounded-lg text-white shadow-lg shadow-red-900/20"><Filter size={20} /></div>
                   <div>
                       <h3 className="font-black text-white uppercase italic tracking-tight">ตัวกรองข้อมูล</h3>
                       <p className="text-[10px] text-zinc-100 font-bold uppercase tracking-widest">Filter Settings</p>
                   </div>
                </div>
                <button onClick={() => setIsFilterSidebarOpen(false)} className="p-2 hover:bg-zinc-800 rounded-full text-zinc-100 transition-colors"><X size={20} /></button>
             </div>
             <div className="space-y-6">
                <button onClick={() => { setFilterSup('All'); setSelectedAgents([]); setSelectedResults([]); setSelectedTypes([]); setSelectedMonth('All'); setActiveCell({ agent: null, resultType: null }); }} className="w-full py-2 text-xs font-black text-red-500 bg-red-600/10 hover:bg-red-600/20 rounded-xl transition-colors border border-red-600/20">RESET ALL FILTERS</button>
                
                <div className="space-y-2">
                    <label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest pl-2">เดือน (Month)</label>
                    <div className="relative">
                      <select className="w-full px-4 py-3 bg-zinc-800 border border-zinc-700 rounded-2xl text-xs font-bold text-white outline-none focus:ring-2 focus:ring-red-600 appearance-none" value={selectedMonth} onChange={(e)=>setSelectedMonth(e.target.value)}>
                          <option value="All">ทุกเดือน</option>
                          {availableMonths.map(m => <option key={m} value={m}>{m}</option>)}
                      </select>
                      <ChevronDown size={14} className="absolute right-4 top-1/2 -translate-y-1/2 text-zinc-400 pointer-events-none" />
                    </div>
                </div>

                <div className="space-y-2">
                    <label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest pl-2">Supervisor</label>
                    <div className="relative">
                      <select className="w-full px-4 py-3 bg-zinc-800 border border-zinc-700 rounded-2xl text-xs font-bold text-white outline-none focus:ring-2 focus:ring-red-600 appearance-none" value={filterSup} onChange={(e) => { setFilterSup(e.target.value); setSelectedAgents([]); }}>
                          <option value="All">ทุก Supervisor</option>
                          {availableSups.map(s => <option key={s} value={s}>{s}</option>)}
                      </select>
                      <ChevronDown size={14} className="absolute right-4 top-1/2 -translate-y-1/2 text-zinc-400 pointer-events-none" />
                    </div>
                </div>

                {/* Multi-Select Result Filter */}
                <div className="space-y-2">
                    <div className="flex items-center justify-between pl-2">
                      <label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest">ผลการสัมภาษณ์</label>
                      <div className="flex gap-2">
                        <button onClick={() => setSelectedResults(RESULT_ORDER)} className="text-[9px] font-bold text-zinc-100 hover:text-white transition-colors">เลือกทั้งหมด</button>
                        <button onClick={() => setSelectedResults([])} className="text-[9px] font-bold text-red-500/70 hover:text-red-500 transition-colors">ล้าง</button>
                      </div>
                    </div>
                    <div className="bg-zinc-800 border border-zinc-700 rounded-2xl p-2 max-h-48 overflow-y-auto custom-scrollbar">
                        {RESULT_ORDER.map(res => {
                        const isSelected = selectedResults.includes(res);
                        return (
                            <div key={res} onClick={() => setSelectedResults(prev => isSelected ? prev.filter(r => r !== res) : [...prev, res])} className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-xs font-bold transition-colors mb-1 ${isSelected ? 'bg-zinc-700 ring-1 ring-zinc-600' : 'hover:bg-zinc-800/50 text-zinc-300'}`}>
                            {isSelected ? <CheckSquare size={16} className="text-red-500" /> : <Square size={16} className="text-zinc-500" />}
                            <span style={{ color: isSelected ? COLORS[res] : undefined }}>{res}</span>
                            </div>
                        );
                        })}
                    </div>
                </div>

                {/* Multi-Select Type Filter */}
                <div className="space-y-2">
                    <div className="flex items-center justify-between pl-2">
                      <label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest">ประเภทงาน</label>
                      <div className="flex gap-2">
                        <button onClick={() => setSelectedTypes(availableTypes)} className="text-[9px] font-bold text-zinc-100 hover:text-white transition-colors">เลือกทั้งหมด</button>
                        <button onClick={() => setSelectedTypes([])} className="text-[9px] font-bold text-red-500/70 hover:text-red-500 transition-colors">ล้าง</button>
                      </div>
                    </div>
                    <div className="bg-zinc-800 border border-zinc-700 rounded-2xl p-2 max-h-40 overflow-y-auto custom-scrollbar">
                        {availableTypes.length > 0 ? availableTypes.map(type => {
                        const isSelected = selectedTypes.includes(type);
                        return (
                            <div key={type} onClick={() => setSelectedTypes(prev => isSelected ? prev.filter(t => t !== type) : [...prev, type])} className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-xs font-bold mb-1 ${isSelected ? 'bg-red-600 text-white shadow-lg shadow-red-900/20' : 'hover:bg-zinc-800 text-zinc-300'}`}>
                            {isSelected ? <CheckSquare size={16} /> : <Square size={16} className="text-zinc-500" />} {type}
                            </div>
                        );
                        }) : <div className="p-2 text-center text-xs text-zinc-300">ไม่พบข้อมูล</div>}
                    </div>
                </div>

                {/* Multi-Select Agent Filter */}
                <div className="space-y-2">
                    <div className="flex items-center justify-between pl-2">
                      <label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest">พนักงาน ({availableAgents.length})</label>
                      <div className="flex gap-2">
                        <button onClick={() => setSelectedAgents(availableAgents)} className="text-[9px] font-bold text-zinc-100 hover:text-white transition-colors">เลือกทั้งหมด</button>
                        <button onClick={() => setSelectedAgents([])} className="text-[9px] font-bold text-red-500/70 hover:text-red-500 transition-colors">ล้าง</button>
                      </div>
                    </div>
                    <div className="bg-zinc-800 border border-zinc-700 rounded-2xl p-2 max-h-60 overflow-y-auto custom-scrollbar">
                        {availableAgents.length > 0 ? availableAgents.map(agent => {
                        const isSelected = selectedAgents.includes(agent);
                        return (
                            <div key={agent} onClick={() => setSelectedAgents(prev => isSelected ? prev.filter(a => a !== agent) : [...prev, agent])} className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-xs font-bold mb-1 ${isSelected ? 'bg-zinc-100 text-black shadow-md' : 'hover:bg-zinc-800 text-zinc-300'}`}>
                            {isSelected ? <CheckSquare size={16} /> : <Square size={16} className="text-zinc-500" />} <span className="truncate">{agent}</span>
                            </div>
                        );
                        }) : <div className="p-2 text-center text-xs text-zinc-100 italic font-bold">ไม่พบพนักงานในเงื่อนไขนี้</div>}
                    </div>
                </div>
             </div>
          </div>
       </aside>

      <div className="max-w-7xl mx-auto space-y-6">
        {/* Header */}
        <header className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 bg-zinc-900 p-6 rounded-[2.5rem] shadow-xl border border-zinc-800">
          <div className="flex items-center gap-6">
            <div className="p-3 bg-zinc-800 rounded-2xl border border-zinc-700">
                <IntageLogo />
            </div>
            <div className="h-10 w-[1px] bg-zinc-800 hidden lg:block opacity-40"></div>
            <div>
              <h1 className="text-xl font-black text-white tracking-tight flex items-center gap-2 uppercase italic">
                QC REPORT 2026
                {loading && <RefreshCw size={18} className="animate-spin text-red-600" />}
              </h1>
              <div className="flex flex-col gap-1 mt-1">
                <div className="text-white text-[10px] font-black flex items-center gap-2 uppercase tracking-widest">
                  {data.length > 0 ? <><div className="w-1.5 h-1.5 rounded-full bg-red-600 animate-pulse"></div> CONNECTED: {data.length} CASES</> : "WAITING FOR CONNECTION"}
                </div>
                {lastUpdated && (
                  <div className="text-zinc-100 text-[9px] font-bold flex items-center gap-1 uppercase tracking-widest">
                    <Clock size={10} className="text-red-600" /> อัปเดตล่าสุด: {lastUpdated} (Autoทุก 5 นาที)
                  </div>
                )}
              </div>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={() => setIsFilterSidebarOpen(true)} className={`flex items-center gap-2 px-5 py-3 rounded-2xl text-xs font-black shadow-sm transition-all border ${hasActiveFilters ? 'bg-red-600 border-red-500 text-white shadow-lg shadow-red-900/20' : 'bg-zinc-800 border-zinc-700 text-white hover:bg-zinc-700'}`}>
              <Filter size={16} /> ตัวกรอง {hasActiveFilters && <span className="w-1.5 h-1.5 rounded-full bg-white animate-bounce"></span>}
            </button>
            <button onClick={() => setShowSync(true)} className="flex items-center gap-2 px-5 py-3 bg-white text-black rounded-2xl text-xs font-black hover:bg-zinc-200 transition-all border-b-2 border-zinc-300 shadow-xl shadow-white/5"><Settings size={14} /> ตั้งค่า</button>
            <button onClick={() => setIsAuthenticated(false)} className="p-3 bg-zinc-800 text-zinc-100 rounded-2xl hover:text-red-500 border border-zinc-700 transition-colors" title="ออกจากระบบ"><User size={20} /></button>
          </div>
        </header>

        {data.length > 0 ? (
          <div className="space-y-6">
            {/* KPI Section */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
            {[
                { label: 'งานที่ตรวจทั้งหมด', value: filteredData.length, icon: FileText, color: 'text-white', bg: 'bg-zinc-900 border-zinc-800' },
                { label: 'อัตราผ่านเกณฑ์', value: `${passRate}%`, icon: CheckCircle, color: 'text-emerald-500', bg: 'bg-emerald-500/10 border-emerald-900/20 shadow-lg shadow-emerald-900/5' },
                { label: 'ควรปรับปรุง', value: filteredData.filter(d=>d.result==='ควรปรับปรุง').length, icon: AlertTriangle, color: 'text-red-500', bg: 'bg-red-500/10 border-red-900/20' },
                { label: 'พบข้อผิดพลาด', value: filteredData.filter(d=>d.result==='พบข้อผิดพลาด').length, icon: XCircle, color: 'text-red-700', bg: 'bg-zinc-900 border-zinc-800' }
            ].map((kpi, i) => (
                <div key={i} className={`p-6 rounded-[2.5rem] border shadow-sm ${kpi.bg}`}>
                  <div className={`w-8 h-8 rounded-xl flex items-center justify-center mb-3 bg-zinc-800/50 ${kpi.color}`}><kpi.icon size={16} /></div>
                  <p className="text-[9px] font-black text-white uppercase tracking-widest">{kpi.label}</p>
                  <h2 className={`text-3xl font-black ${kpi.color} tracking-tighter mt-1 uppercase`}>{kpi.value}</h2>
                </div>
            ))}
            </div>

            {/* Visual Analytics Chart */}
            <div className="bg-zinc-900 p-8 rounded-[3rem] shadow-xl border border-zinc-800">
                <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
                    <div className="flex items-center gap-3">
                        <div className="p-3 bg-red-600 rounded-xl text-white shadow-lg shadow-red-900/30"><BarChart2 size={20}/></div>
                        <h3 className="font-black text-white uppercase italic tracking-tight text-lg">Performance Distribution</h3>
                    </div>
                    <div className="flex items-center gap-2 flex-wrap">
                        {chartData.map((item) => (
                            <div key={item.name} className="flex flex-col items-center bg-zinc-800/50 px-4 py-2 rounded-2xl min-w-[90px] border border-zinc-800/50">
                                <span className="text-[9px] font-black text-white uppercase tracking-widest">{item.name}</span>
                                <span className="text-sm font-black" style={{color: item.color}}>{item.percent}%</span>
                            </div>
                        ))}
                    </div>
                </div>
                <div className="h-72 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#3f3f46" opacity={0.5} />
                            <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 10, fontWeight: 900, fill: '#e4e4e7', fontFamily: 'Sarabun'}} dy={10} />
                            <YAxis axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#e4e4e7', fontFamily: 'Sarabun'}} />
                            <Tooltip cursor={{fill: '#18181b', opacity: 0.5}} content={({ active, payload, label }) => {
                              if (active && payload?.length) {
                                return (
                                  <div className="bg-zinc-800 p-4 rounded-2xl shadow-2xl border border-zinc-700 text-white font-sans">
                                    <p className="font-black text-xs uppercase mb-1 tracking-widest">{label}</p>
                                    <p className="text-zinc-100 font-black text-xl" style={{color: payload[0].payload.color}}>{payload[0].value} <span className="text-white text-xs font-bold uppercase">Cases</span></p>
                                    <p className="text-zinc-100 text-[10px] font-bold mt-1">PERCENT: {payload[0].payload.percent}%</p>
                                  </div>
                                );
                              }
                              return null;
                            }} />
                            <Bar dataKey="count" radius={[12, 12, 12, 12]} barSize={55}>
                                {chartData.map((entry, index) => <Cell key={index} fill={entry.color} />)}
                            </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>
            </div>

            {/* Matrix Table */}
            <div className="bg-zinc-900 rounded-[3rem] shadow-2xl border border-zinc-800 overflow-hidden">
                <div className="p-8 border-b border-zinc-800 bg-zinc-900/50 flex items-center justify-between">
                    <div>
                    <h3 className="font-black text-white flex items-center gap-2 italic text-lg uppercase tracking-tight">
                        <TrendingUp size={24} className="text-red-600" /> สรุปคุณภาพพนักงาน x ผลสัมภาษณ์
                    </h3>
                    <p className="text-[9px] text-zinc-100 font-black uppercase mt-1 italic tracking-widest underline decoration-red-600/30 flex items-center gap-1"><ChevronRight size={12}/> คลิกที่ตัวเลข เพื่อดูรายละเอียดคอมเมนต์ด้านล่าง</p>
                    </div>
                </div>
                <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                    <table className="w-full text-left text-sm border-separate border-spacing-0">
                    <thead className="sticky top-0 bg-zinc-900 z-20 font-black text-white text-[10px] uppercase tracking-widest border-b border-zinc-800">
                        <tr>
                            <th rowSpan="2" className="px-8 py-6 border-b border-zinc-800 border-r border-zinc-800 bg-zinc-900">พนักงานสัมภาษณ์</th>
                            <th colSpan={RESULT_ORDER.length} className="px-4 py-4 text-center border-b border-zinc-800 bg-zinc-900/40 text-red-600 text-[11px] font-black italic uppercase">สรุปผลการสัมภาษณ์</th>
                            <th rowSpan="2" className="px-8 py-6 text-center bg-zinc-800 text-white border-b border-zinc-800 border-l border-zinc-800">รวม</th>
                        </tr>
                        <tr className="bg-zinc-900/80">
                            {RESULT_ORDER.map(type => <th key={type} className="px-4 py-3 text-center border-b border-zinc-800 border-r border-zinc-800 text-white">{type}</th>)}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-800/40">
                        {agentSummary.map((agent, i) => (
                        <tr key={i} className="hover:bg-zinc-800/50 transition-colors group">
                            <td className="px-8 py-5 font-black text-white border-r border-zinc-800">{agent.name}</td>
                            {RESULT_ORDER.map(type => {
                            const val = agent[type];
                            const isActive = activeCell.agent === agent.name && activeCell.resultType === type;
                            return (
                                <td key={type} className={`px-4 py-5 text-center border-r border-zinc-800 transition-all ${val > 0 ? 'cursor-pointer hover:bg-zinc-800/60 shadow-inner' : ''} ${isActive ? 'bg-zinc-800 ring-2 ring-inset ring-red-600 shadow-lg' : ''}`} onClick={() => val > 0 && handleMatrixClick(agent.name, type)}>
                                    <span className={`text-sm font-black ${val > 0 ? '' : 'text-zinc-900'}`} style={{ color: val > 0 ? COLORS[type] : undefined }}>{val || '-'}</span>
                                </td>
                            );
                            })}
                            <td className="px-8 py-5 text-center bg-zinc-800/20 font-black text-white border-l border-zinc-800 group-hover:bg-zinc-800 transition-colors">{agent.total}</td>
                        </tr>
                        ))}
                    </tbody>
                    </table>
                </div>
            </div>

            {/* Detailed Case Log */}
            <div id="detail-section" className="bg-zinc-900 rounded-[3rem] shadow-2xl border border-zinc-800 overflow-hidden scroll-mt-6">
                <div className="p-8 border-b border-zinc-800 flex flex-col md:flex-row items-center justify-between gap-6">
                    <div className="space-y-1">
                        <h3 className="font-black text-white uppercase tracking-widest text-xs flex items-center gap-2 italic">
                            <MessageSquare size={16} className="text-red-600" /> ข้อมูลรายเคสแบบละเอียด
                        </h3>
                        <p className="text-[9px] text-zinc-500 font-bold uppercase tracking-widest italic ml-6">คลิกที่แถวเพื่อดูผลการประเมิน 13 หัวข้อ</p>
                        {activeCell.agent && (
                            <div className="flex items-center gap-2 mt-2">
                                <span className="text-[9px] font-black px-3 py-1 bg-red-600 text-white rounded-lg shadow-lg shadow-red-900/30 uppercase italic tracking-tighter animate-pulse">
                                    FILTERING: {activeCell.agent} ({activeCell.resultType})
                                </span>
                                <button onClick={() => setActiveCell({ agent: null, resultType: null })} className="p-1 hover:bg-zinc-800 rounded text-zinc-100 hover:text-red-500 transition-colors"><X size={12}/></button>
                            </div>
                        )}
                    </div>
                    <div className="relative w-full md:w-80">
                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-400" size={16} />
                        <input type="text" placeholder="ค้นหาพนักงาน หรือ คอมเมนต์..." className="w-full pl-12 pr-6 py-4 bg-zinc-800/50 border border-zinc-800 rounded-2xl text-sm font-bold outline-none focus:ring-2 focus:ring-red-600 text-white placeholder-zinc-500 transition-all shadow-inner" value={searchTerm} onChange={(e)=>setSearchTerm(e.target.value)} />
                    </div>
                </div>
                <div className="overflow-auto max-h-[800px] custom-scrollbar">
                    <table className="w-full text-left text-xs font-medium border-separate border-spacing-0">
                    <thead className="sticky top-0 bg-zinc-900 shadow-md z-10 border-b border-zinc-800 font-black text-white uppercase tracking-widest">
                        <tr>
                            <th className="px-8 py-5 border-r border-zinc-800/30">วันที่ / ประเภท</th>
                            <th className="px-8 py-5 border-r border-zinc-800/30">พนักงานสัมภาษณ์</th>
                            <th className="px-4 py-5 text-center border-r border-zinc-800/30">ผลการสัมภาษณ์</th>
                            <th className="px-8 py-5">QC Comment & Audio</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-800/30">
                        {detailLogs.length > 0 ? detailLogs.slice(0, 150).map((item, idx) => {
                          const isExpanded = expandedCaseId === item.id;
                          return (
                            <React.Fragment key={item.id}>
                              <tr 
                                onClick={() => setExpandedCaseId(isExpanded ? null : item.id)}
                                className={`transition-all group cursor-pointer ${isExpanded ? 'bg-zinc-800/80 shadow-inner' : 'hover:bg-zinc-800/40'}`}
                              >
                                  <td className="px-8 py-6 border-r border-zinc-800/20">
                                      <div className="font-black text-white">{item.date}</div>
                                      <div className="text-[9px] font-black text-red-600 uppercase mt-1 tracking-tighter shadow-red-900/10 italic">{item.type} &bull; SUP: {item.supervisor}</div>
                                  </td>
                                  <td className="px-8 py-6 border-r border-zinc-800/20">
                                      <div className="font-black text-white text-sm group-hover:text-red-500 transition-colors flex items-center gap-2">
                                        {item.agent}
                                        {isExpanded ? <ChevronDown size={14} className="text-zinc-500" /> : <ChevronRight size={14} className="text-zinc-700" />}
                                      </div>
                                      <div className="text-[9px] text-zinc-300 font-bold uppercase tracking-widest mt-0.5 italic opacity-80">{item.touchpoint}</div>
                                  </td>
                                  <td className="px-4 py-6 text-center border-r border-zinc-800/20">
                                      <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[9px] font-black border uppercase shadow-sm" style={{ backgroundColor: `${COLORS[item.result]}10`, color: COLORS[item.result], borderColor: `${COLORS[item.result]}25` }}>
                                          <div className="w-1 h-1 rounded-full" style={{backgroundColor: COLORS[item.result]}}></div>
                                          {item.result}
                                      </span>
                                  </td>
                                  <td className="px-8 py-6">
                                      <div className="flex flex-col gap-2">
                                          <p className="text-zinc-300 font-semibold italic max-w-sm leading-relaxed group-hover:text-white transition-colors">
                                              {item.comment ? `"${item.comment}"` : '-'}
                                          </p>
                                          {item.audio && item.audio.includes('http') && (
                                          <a href={item.audio} target="_blank" rel="noopener noreferrer" onClick={(e)=>e.stopPropagation()} className="flex items-center gap-1.5 text-red-500 hover:text-red-400 font-black text-[10px] uppercase transition-all hover:translate-x-1">
                                              <PlayCircle size={14} /> Listen Recording
                                          </a>
                                          )}
                                      </div>
                                  </td>
                              </tr>
                              {isExpanded && (
                                <tr className="bg-zinc-950/40 animate-in slide-in-from-top-2 duration-300">
                                  <td colSpan={4} className="p-8 border-b border-zinc-800">
                                    <div className="flex items-center gap-4 mb-6">
                                      <div className="p-2 bg-red-600/10 rounded-lg text-red-500"><Award size={20} /></div>
                                      <div>
                                        <h4 className="font-black text-white uppercase italic tracking-widest text-sm">รายละเอียดการประเมิน 13 หัวข้อ</h4>
                                        <p className="text-[9px] text-zinc-500 font-bold uppercase tracking-widest">Detailed Quality Assessment (P:AB Columns)</p>
                                      </div>
                                    </div>
                                    
                                    <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3">
                                      {item.evaluations.map((evalItem, eIdx) => (
                                        <div key={eIdx} className="bg-zinc-900/80 border border-zinc-800 p-3 rounded-2xl hover:border-red-600/30 transition-colors group/item">
                                          <p className="text-[9px] font-black text-zinc-500 uppercase tracking-tighter mb-2 line-clamp-1 group-hover/item:text-zinc-300 transition-colors" title={evalItem.label}>
                                            {evalItem.label}
                                          </p>
                                          <div className={`text-sm font-black italic tracking-widest ${evalItem.value === '1' || evalItem.value === '100' ? 'text-emerald-500' : evalItem.value === '0' ? 'text-red-500' : 'text-zinc-200'}`}>
                                            {evalItem.value}
                                          </div>
                                        </div>
                                      ))}
                                    </div>

                                    {item.comment && (
                                      <div className="mt-6 p-4 bg-zinc-800/30 rounded-2xl border border-zinc-800/50 border-l-4 border-l-red-600">
                                        <p className="text-[10px] font-black text-red-500 uppercase tracking-widest mb-1 italic">QC Full Comment</p>
                                        <p className="text-sm text-zinc-100 font-medium italic">"{item.comment}"</p>
                                      </div>
                                    )}
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                        }) : (
                        <tr><td colSpan={4} className="px-8 py-24 text-center text-white font-black uppercase italic tracking-widest text-lg opacity-20">No Matching Analysis Data</td></tr>
                        )}
                    </tbody>
                    </table>
                </div>
            </div>
          </div>
        ) : (
          !loading && (
            <div className="bg-zinc-900 rounded-[4rem] border-2 border-dashed border-zinc-800 py-32 text-center shadow-inner">
                <div className="bg-zinc-800 w-24 h-24 rounded-[3rem] flex items-center justify-center mx-auto mb-8 shadow-2xl border border-zinc-700">
                    <Database size={40} className="text-zinc-600" />
                </div>
                <h2 className="text-2xl font-black text-white uppercase italic tracking-widest">Waiting for Signal</h2>
                <p className="text-zinc-100 text-xs mt-3 max-w-sm mx-auto font-black uppercase tracking-widest leading-relaxed">กรุณากดปุ่ม "ตั้งค่าการเชื่อมต่อ" ด้านบน และวางลิงก์ CSV จาก Google Sheets</p>
            </div>
          )
        )}
      </div>

      {/* Sync Modal */}
      {(showSync || error) && (
        <div className="fixed inset-0 z-[100] bg-black/90 backdrop-blur-xl flex items-center justify-center p-4">
            <div className="bg-zinc-900 w-full max-w-2xl rounded-[3rem] p-10 border border-zinc-800 shadow-2xl animate-in zoom-in duration-300 relative overflow-hidden font-sans">
                <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-red-600 via-red-900 to-zinc-900"></div>
                <div className="flex items-center justify-between mb-10">
                    <h3 className="text-xl font-black flex items-center gap-3 text-white uppercase italic tracking-tight">
                        {error ? <AlertCircle className="text-red-500" /> : <Settings className="text-red-600" />}
                        {error ? 'System Error' : 'Connection Settings'}
                    </h3>
                    {data.length > 0 && <button onClick={() => {setShowSync(false); setError(null);}} className="text-zinc-600 hover:text-white transition-colors"><X size={28} /></button>}
                </div>
                <div className="space-y-8">
                    <div className="bg-zinc-950/50 p-6 rounded-[2rem] border border-zinc-800/50">
                        <p className="text-[10px] font-black text-white uppercase tracking-widest mb-4 flex items-center gap-2 italic underline decoration-red-600 decoration-2">Instructions</p>
                        <ol className="text-xs text-zinc-100 space-y-3 font-medium">
                            <li className="flex gap-2"><span>1.</span> ไปที่ Google Sheets { ' > ' } File { ' > ' } Publish to web</li>
                            <li className="flex gap-2"><span>2.</span> เลือกชีต <span className="text-red-600 font-bold underline italic tracking-tight">"ACQC"</span> และเปลี่ยนเป็นรูปแบบ <span className="text-white font-bold underline italic tracking-tight">"CSV"</span></li>
                            <li className="flex gap-2"><span>3.</span> คัดลอกลิงก์ที่ได้มาวางในช่องด้านล่าง</li>
                        </ol>
                    </div>
                    <div className="flex flex-col gap-5">
                        <label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest pl-2">CSV URL Link</label>
                        <input type="text" placeholder="https://docs.google.com/spreadsheets/d/e/..." className="w-full px-8 py-5 bg-zinc-950 border border-zinc-800 rounded-3xl text-sm font-bold text-white outline-none focus:ring-2 focus:ring-red-600 transition-all shadow-inner" value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)} />
                        <div className="flex gap-4">
                            <button onClick={() => fetchFromSheet(sheetUrl)} disabled={loading || !sheetUrl} className="flex-1 py-5 bg-red-600 hover:bg-red-700 text-white rounded-3xl font-black uppercase transition-all shadow-xl shadow-red-900/30 disabled:opacity-20 text-sm tracking-widest italic">
                                {loading ? 'SYNCING...' : 'CONNECT & ANALYZE'}
                            </button>
                            {data.length > 0 && (
                                <button onClick={clearConnection} className="p-5 bg-zinc-800 text-zinc-100 hover:text-red-500 rounded-3xl border border-zinc-700 hover:border-red-900/30 transition-all"><Trash2 size={24}/></button>
                            )}
                        </div>
                        {error && <p className="text-center text-xs text-red-500 font-black mt-2 uppercase italic tracking-tighter animate-pulse bg-red-600/5 p-3 rounded-xl border border-red-600/10 shadow-lg">{error}</p>}
                    </div>
                </div>
            </div>
        </div>
      )}
    </div>
  );
};

export default App;