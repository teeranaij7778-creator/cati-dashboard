import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell, PieChart, Pie
} from 'recharts';
import { 
  Users, CheckCircle, AlertTriangle, XCircle, Search, 
  FileText, BarChart2, MessageSquare, Calendar, TrendingUp, Database, Link, RefreshCw, Trash2, Globe, FilterX, PlayCircle, UserCheck, Settings, AlertCircle, Info, ChevronRight, ExternalLink, User, ChevronDown, CheckSquare, Square, X, Briefcase, Lock, LogIn, Activity, Filter, Check, Clock, ListChecks, Award, Save, Edit2, Hash, Star
} from 'lucide-react';

/** * CATI CES 2026 Analytics Dashboard - INTAGE BLACK & RED Edition
 * ระบบวิเคราะห์ผลการตรวจ QC พร้อมระบบแก้ไขข้อมูล (Edit Mode)
 * แก้ไข: เพิ่มการดึง Tailwind CDN และ Google Fonts เพื่อให้รันใน VS Code ได้ทันที
 */

const DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSHePu18q6f93lQqVW5_JNv6UygyYRGNjT5qOq4nSrROCnGxt1pkdgiPT91rm-_lVpku-PW-LWs-ufv/pub?gid=470556665&single=true&output=csv"; 
const DEFAULT_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzDXUfKHXtjv2ZEMOOS5-aB9ApOb3hT4vvO-9jZ3g9dgOAtQQ2pikzY_gRq5edmD01s/exec";

const RESULT_ORDER = [
  'ดีเยี่ยม: ครบถ้วนตามมาตรฐาน (พนักงานทำได้ดีทุกข้อ น้ำเสียงเป็นมืออาชีพ ข้อมูลแม่นยำ 100%)',
  'ผ่านเกณฑ์: ปฏิบัติงานได้ถูกต้อง (ทำได้ตามมาตรฐาน มีข้อผิดพลาดเล็กน้อยที่ไม่กระทบคุณภาพข้อมูลหลัก)',
  'ควรปรับปรุง: มารยาทและน้ำเสียง (มีคำฟุ่มเฟือย/หัวเราะขณะสัมภาษณ์ หรือสนิทสนมกับผู้ตอบมากเกินไป)',
  'ควรปรับปรุง: การอ่านคำถาม/ตัวเลือก (อ่านไม่ครบถ้วน อ่านข้ามตัวเลือก หรือรวบรัดคำถามตามความเข้าใจตนเอง)',
  'ควรปรับปรุง: การจดบันทึก Open-end (จดบันทึกไม่ละเอียด ไม่ครบทุกคำ หรือซักคำถามปลายเปิดไม่เพียงพอ)',
  'พบข้อผิดพลาด: มีการชี้นำคำตอบ (แสดงความเห็นส่วนตัว แนะนำคำตอบ หรือพูดแทรกเพื่อเร่งรัดการสัมภาษณ์)',
  'พบข้อผิดพลาด: ข้อมูลไม่ตรงตามจริง (บันทึกคำตอบผิดจากที่ตอบจริง หรือทำผิดเงื่อนไข Logic ของแบบสอบถาม)',
  'ไม่ผ่านเกณฑ์: ต้องอบรมใหม่ทันที (มีข้อผิดพลาดรุนแรงในจุดสำคัญหลายข้อซึ่งส่งผลเสียต่อคุณภาพงานวิจัย)'
];

const SCORE_OPTIONS = ['5', '4', '3', '2', '1', '-'];
const SCORE_LABELS = {
  '5': '5.ดี',
  '4': '4.ค่อนข้างดี',
  '3': '3.ปานกลาง',
  '2': '2.ไม่ค่อยดี',
  '1': '1.ไม่ดีเลย',
  '-': '-'
};

const formatResultDisplay = (text) => (text ? text.split('(')[0].trim() : '-');

const COLORS = {
  'ดีเยี่ยม': '#3B82F6',     
  'ผ่านเกณฑ์': '#10B981',   
  'ควรปรับปรุง': '#F59E0B', 
  'พบข้อผิดพลาด': '#EF4444', 
  'ไม่ผ่านเกณฑ์': '#B91C1C', 
};

const getResultColor = (fullText) => {
  if (fullText.startsWith('ดีเยี่ยม')) return COLORS['ดีเยี่ยม'];
  if (fullText.startsWith('ผ่านเกณฑ์')) return COLORS['ผ่านเกณฑ์'];
  if (fullText.startsWith('ควรปรับปรุง')) return COLORS['ควรปรับปรุง'];
  if (fullText.startsWith('พบข้อผิดพลาด')) return COLORS['พบข้อผิดพลาด'];
  if (fullText.startsWith('ไม่ผ่านเกณฑ์')) return COLORS['ไม่ผ่านเกณฑ์'];
  return '#71717a';
};

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
  const [loading, setLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState(null);
  
  const getStorage = (key, fallback) => {
    try { return localStorage.getItem(key) || fallback; } catch(e) { return fallback; }
  };

  const [sheetUrl, setSheetUrl] = useState(getStorage('qc_sheet_url', DEFAULT_SHEET_URL));
  const [appsScriptUrl, setAppsScriptUrl] = useState(getStorage('apps_script_url', DEFAULT_APPS_SCRIPT_URL));
  const [searchTerm, setSearchTerm] = useState('');
  const [isFilterSidebarOpen, setIsFilterSidebarOpen] = useState(false);
  const [lastUpdated, setLastUpdated] = useState(null);
  
  const [filterSup, setFilterSup] = useState('All');
  const [selectedResults, setSelectedResults] = useState([]);
  const [selectedAgents, setSelectedAgents] = useState([]); 
  const [selectedMonth, setSelectedMonth] = useState('All');
  const [showSync, setShowSync] = useState(getStorage('qc_sheet_url', '') === '' && DEFAULT_SHEET_URL === '');
  
  const [activeCell, setActiveCell] = useState({ agent: null, resultType: null });
  const [expandedCaseId, setExpandedCaseId] = useState(null);
  const [editingCase, setEditingCase] = useState(null); 

  // --- Ensure Tailwind and Font are loaded for VS Code / Local ---
  useEffect(() => {
    // 1. Load Thai Font
    if (!document.getElementById('thai-font-link')) {
      const link = document.createElement('link');
      link.id = 'thai-font-link';
      link.rel = 'stylesheet';
      link.href = 'https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;700;800&display=swap';
      document.head.appendChild(link);
    }

    // 2. Load Tailwind CDN (Safety for environments without Tailwind)
    if (!document.getElementById('tailwind-cdn')) {
      const script = document.createElement('script');
      script.id = 'tailwind-cdn';
      script.src = 'https://cdn.tailwindcss.com';
      document.head.appendChild(script);
    }

    // 3. Inject Custom Styles
    const styleId = 'custom-dashboard-styles';
    if (!document.getElementById(styleId)) {
      const style = document.createElement('style');
      style.id = styleId;
      style.innerHTML = `
        body { 
          font-family: 'Sarabun', sans-serif !important; 
          background-color: #09090b; /* zinc-950 */
          margin: 0;
        }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #3f3f46; border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #52525b; }
      `;
      document.head.appendChild(style);
    }
  }, []);

  useEffect(() => {
    let intervalId;
    if (isAuthenticated && sheetUrl && sheetUrl.includes('http')) {
      fetchFromSheet(sheetUrl);
      intervalId = setInterval(() => fetchFromSheet(sheetUrl), 300000);
    }
    return () => clearInterval(intervalId);
  }, [sheetUrl, isAuthenticated]);

  const fetchFromSheet = async (urlToFetch) => {
    setLoading(true);
    setError(null);
    try {
      const fetchUrl = `${urlToFetch.trim()}&t=${new Date().getTime()}`;
      const response = await fetch(fetchUrl);
      if (!response.ok) throw new Error("ไม่สามารถเข้าถึงไฟล์ได้");
      const csvText = await response.text();
      const allRows = parseCSV(csvText);
      let headerIdx = allRows.findIndex(row => row.some(cell => cell.toString().toLowerCase().includes("interviewer") || cell.toString().includes("สรุปผล")));
      
      if (headerIdx === -1) throw new Error("ไม่พบคอลัมน์ข้อมูลที่กำหนด");

      const headers = allRows[headerIdx].map(h => h.trim());
      const evaluationsList = headers.slice(15, 28);

      const getIdx = (name) => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
      let qNoIdx = headers.findIndex(h => h.toLowerCase().includes("questionnaire") || h.toLowerCase().includes("no.") || h.toLowerCase().includes("เลขชุด"));
      if (qNoIdx === -1) qNoIdx = 3; 

      const idx = {
        month: getIdx("เดือน"), date: getIdx("วันที่สัมภาษณ์"), touchpoint: getIdx("TOUCH_POINT"), 
        type: getIdx("AC / BC"), sup: getIdx("Supervisor"), agent: getIdx("Interviewer"),
        questionnaireNo: qNoIdx, result: getIdx("สรุปผลการสัมภาษณ์"), comment: getIdx("Comment"),
        audio: getIdx("ไฟล์เสียง") 
      };
      
      const parsedData = allRows.slice(headerIdx + 1).filter(row => {
          const agentCode = row[idx.agent]?.toString().trim() || "";
          return agentCode !== "" && !agentCode.includes("#N/A");
        })
        .map((row, index) => {
          let rawResult = row[idx.result]?.toString().trim() || "N/A";
          let cleanResult = rawResult;
          const matchedResult = RESULT_ORDER.find(opt => rawResult.includes(opt.split(':')[0].trim()));
          if (matchedResult) cleanResult = matchedResult;
          
          return {
            id: index, rowIndex: index + headerIdx + 2, 
            month: row[idx.month] || 'N/A', date: row[idx.date] || 'N/A', 
            agent: row[idx.agent] || 'Unknown', questionnaireNo: row[idx.questionnaireNo] || '-', 
            result: cleanResult, comment: row[idx.comment] || '',
            audio: row[idx.audio] || '', 
            touchpoint: row[idx.touchpoint] || 'N/A', supervisor: row[idx.sup] || 'N/A',
            type: row[idx.type] || 'N/A',
            evaluations: evaluationsList.map((header, i) => ({ label: header, value: row[15 + i] || '-' }))
          };
        });
      setData(parsedData);
      setLastUpdated(new Date().toLocaleTimeString('th-TH'));
      try { localStorage.setItem('qc_sheet_url', urlToFetch); } catch(e) {}
      setShowSync(false);
    } catch (err) { setError(err.message); } finally { setLoading(false); }
  };

  const handleUpdateCase = async () => {
    if (!appsScriptUrl) return alert("กรุณาตั้งค่า Apps Script URL");
    setIsSaving(true);
    try {
      const updateData = {
        rowIndex: editingCase.rowIndex,
        result: editingCase.result,
        evaluations: editingCase.evaluations.map(e => e.value),
        comment: editingCase.comment
      };
      await fetch(appsScriptUrl, { method: 'POST', mode: 'no-cors', cache: 'no-cache', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(updateData) });
      setTimeout(() => { setIsSaving(false); setEditingCase(null); fetchFromSheet(sheetUrl); }, 2000);
    } catch (err) { alert(err.message); setIsSaving(false); }
  };

  const availableMonths = useMemo(() => [...new Set(data.map(d => d.month).filter(m => m !== 'N/A'))], [data]);
  const availableSups = useMemo(() => [...new Set(data.map(d => d.supervisor).filter(s => s !== 'N/A'))].sort(), [data]);
  const availableAgents = useMemo(() => {
    let filtered = data;
    if (filterSup !== 'All') filtered = filtered.filter(d => d.supervisor === filterSup);
    return [...new Set(filtered.map(d => d.agent).filter(a => a !== 'Unknown'))].sort();
  }, [data, filterSup]);

  const filteredData = useMemo(() => {
    return data.filter(item => {
      const matchesSearch = item.agent.toLowerCase().includes(searchTerm.toLowerCase()) || 
                           item.comment.toLowerCase().includes(searchTerm.toLowerCase()) ||
                           item.questionnaireNo.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesResult = selectedResults.length === 0 || selectedResults.includes(item.result);
      const matchesSup = filterSup === 'All' || item.supervisor === filterSup;
      const matchesAgent = selectedAgents.length === 0 || selectedAgents.includes(item.agent);
      const matchesMonth = selectedMonth === 'All' || item.month === selectedMonth;
      return matchesSearch && matchesResult && matchesSup && matchesAgent && matchesMonth;
    });
  }, [data, searchTerm, selectedResults, filterSup, selectedAgents, selectedMonth]);

  const agentSummary = useMemo(() => {
    const summaryMap = {};
    filteredData.forEach(item => {
      if (!summaryMap[item.agent]) {
        summaryMap[item.agent] = { name: item.agent, total: 0 };
        RESULT_ORDER.forEach(r => summaryMap[item.agent][r] = 0);
      }
      if (summaryMap[item.agent][item.result] !== undefined) summaryMap[item.agent][item.result] += 1;
      summaryMap[item.agent].total += 1;
    });
    return Object.values(summaryMap).sort((a, b) => b.total - a.total);
  }, [filteredData]);

  const chartData = useMemo(() => {
    const total = filteredData.length;
    return RESULT_ORDER.map(key => ({ 
      name: key, displayLabel: formatResultDisplay(key), 
      count: filteredData.filter(d => d.result === key).length, 
      percent: total > 0 ? ((filteredData.filter(d => d.result === key).length / total) * 100).toFixed(1) : 0, 
      color: getResultColor(key) 
    }));
  }, [filteredData]);

  const detailLogs = useMemo(() => (activeCell.agent && activeCell.resultType) ? filteredData.filter(d => d.agent === activeCell.agent && d.result === activeCell.resultType) : filteredData, [filteredData, activeCell]);
  const passRate = useMemo(() => {
    if (filteredData.length === 0) return 0;
    const passCount = filteredData.filter(d => d.result.startsWith('ดีเยี่ยม') || d.result.startsWith('ผ่านเกณฑ์')).length;
    return ((passCount / filteredData.length) * 100).toFixed(1);
  }, [filteredData]);

  const handleMatrixClick = (agentName, type) => {
    if (activeCell.agent === agentName && activeCell.resultType === type) setActiveCell({ agent: null, resultType: null });
    else { setActiveCell({ agent: agentName, resultType: type }); document.getElementById('detail-section')?.scrollIntoView({ behavior: 'smooth' }); }
  };

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-zinc-950 flex items-center justify-center p-4 text-white">
        <div className="bg-zinc-900 p-10 rounded-[2.5rem] border border-zinc-800 w-full max-w-sm text-center shadow-2xl">
          <div className="flex justify-center mb-8"><IntageLogo className="scale-125" /></div>
          <h2 className="text-white font-black uppercase text-xs mb-8 tracking-widest italic">CATI CES 2026 DASHBOARD</h2>
          <form onSubmit={(e) => { e.preventDefault(); if(inputUser==='Admin'&&inputPass==='1234') setIsAuthenticated(true); else setLoginError('Invalid Creds'); }} className="space-y-4 text-left">
            <div className="space-y-2"><label className="text-[10px] font-black uppercase tracking-widest pl-2">Username</label><input type="text" value={inputUser} onChange={e=>setInputUser(e.target.value)} className="w-full p-4 bg-zinc-800 border border-zinc-700 rounded-2xl outline-none focus:ring-2 focus:ring-red-600 text-white" placeholder="Admin" /></div>
            <div className="space-y-2"><label className="text-[10px] font-black uppercase tracking-widest pl-2">Password</label><input type="password" value={inputPass} onChange={e=>setInputPass(e.target.value)} className="w-full p-4 bg-zinc-800 border border-zinc-700 rounded-2xl outline-none focus:ring-2 focus:ring-red-600 text-white" placeholder="••••" /></div>
            {loginError && <p className="text-red-500 text-xs font-bold text-center">รหัสผ่านไม่ถูกต้อง</p>}
            <button type="submit" className="w-full py-4 bg-red-600 hover:bg-red-700 text-white rounded-2xl font-black uppercase tracking-widest transition-all">SIGN IN</button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-zinc-950 p-4 md:p-8 text-zinc-100">
      
      {/* Sidebar Panel */}
      <div className={`fixed inset-0 z-40 bg-black/60 backdrop-blur-sm transition-opacity duration-300 ${isFilterSidebarOpen ? 'opacity-100' : 'opacity-0 pointer-events-none'}`} onClick={() => setIsFilterSidebarOpen(false)} />
      <aside className={`fixed inset-y-0 right-0 z-50 w-80 bg-zinc-900 shadow-2xl transform transition-transform duration-300 ${isFilterSidebarOpen ? 'translate-x-0' : 'translate-x-full'} overflow-y-auto border-l border-zinc-800 p-6`}>
          <div className="flex items-center justify-between mb-8 pb-4 border-b border-zinc-800">
             <div className="flex items-center gap-2"><div className="p-2 bg-red-600 rounded-lg text-white"><Filter size={20} /></div><h3 className="font-black text-white uppercase italic tracking-tight">ตัวกรองข้อมูล</h3></div>
             <button onClick={() => setIsFilterSidebarOpen(false)}><X size={20} /></button>
          </div>
          <div className="space-y-6">
             <button onClick={() => { setFilterSup('All'); setSelectedAgents([]); setSelectedResults([]); setSelectedMonth('All'); setActiveCell({ agent: null, resultType: null }); }} className="w-full py-2 text-xs font-black text-red-500 bg-red-600/10 hover:bg-red-600/20 rounded-xl border border-red-600/20 uppercase tracking-widest transition-colors">Reset Filters</button>
             <div className="space-y-2"><label className="text-[10px] font-black uppercase tracking-widest pl-2">เดือน (Month)</label><select className="w-full px-4 py-3 bg-zinc-800 border border-zinc-700 rounded-2xl text-xs font-bold outline-none text-white" value={selectedMonth} onChange={(e)=>setSelectedMonth(e.target.value)}><option value="All">ทุกเดือน</option>{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}</select></div>
             <div className="space-y-2"><label className="text-[10px] font-black uppercase tracking-widest pl-2">Supervisor</label><select className="w-full px-4 py-3 bg-zinc-800 border border-zinc-700 rounded-2xl text-xs font-bold outline-none text-white" value={filterSup} onChange={(e) => { setFilterSup(e.target.value); setSelectedAgents([]); }}><option value="All">ทุก Supervisor</option>{availableSups.map(s => <option key={s} value={s}>{s}</option>)}</select></div>
             <div className="space-y-2">
                 <div className="flex justify-between pl-2 mb-1"><label className="text-[10px] font-black uppercase tracking-widest">พนักงานสัมภาษณ์</label><button onClick={() => setSelectedAgents(availableAgents)} className="text-[9px] text-zinc-500 hover:text-white">เลือกทั้งหมด</button></div>
                 <div className="bg-zinc-800 border border-zinc-700 rounded-2xl p-2 max-h-48 overflow-y-auto custom-scrollbar">
                     {availableAgents.map(agent => (
                         <div key={agent} onClick={() => setSelectedAgents(prev => prev.includes(agent) ? prev.filter(a => a !== agent) : [...prev, agent])} className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-[10px] font-bold mb-1 ${selectedAgents.includes(agent) ? 'bg-zinc-700 text-white' : 'hover:bg-zinc-800/50 text-zinc-400'}`}>
                           {selectedAgents.includes(agent) ? <CheckSquare size={14} className="text-red-500" /> : <Square size={14} />} <span className="truncate">{agent}</span>
                         </div>
                     ))}
                 </div>
             </div>
             <div className="space-y-2">
                 <label className="text-[10px] font-black uppercase tracking-widest pl-2">ผลการสัมภาษณ์</label>
                 <div className="bg-zinc-800 border border-zinc-700 rounded-2xl p-2 max-h-48 overflow-y-auto custom-scrollbar">
                     {RESULT_ORDER.map(res => (
                         <div key={res} onClick={() => setSelectedResults(prev => prev.includes(res) ? prev.filter(r => r !== res) : [...prev, res])} className={`flex items-start gap-2 p-2 rounded-xl cursor-pointer text-[10px] font-bold mb-1 ${selectedResults.includes(res) ? 'bg-zinc-700 text-white' : 'hover:bg-zinc-800/50 text-zinc-400'}`}>
                           {selectedResults.includes(res) ? <CheckSquare size={14} className="text-red-500 mt-0.5" /> : <Square size={14} className="mt-0.5" />}
                           <span style={{ color: selectedResults.includes(res) ? getResultColor(res) : undefined }}>{formatResultDisplay(res)}</span>
                         </div>
                     ))}
                 </div>
             </div>
          </div>
       </aside>

      <div className="max-w-7xl mx-auto space-y-6">
        <header className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 bg-zinc-900 p-6 rounded-[2.5rem] shadow-xl border border-zinc-800">
          <div className="flex items-center gap-6"><div className="p-3 bg-zinc-800 rounded-2xl border border-zinc-700"><IntageLogo /></div>
            <div>
              <h1 className="text-xl font-black text-white tracking-tight flex items-center gap-2 uppercase italic">QC REPORT 2026 {loading && <RefreshCw size={18} className="animate-spin text-red-600" />}</h1>
              <div className="text-white text-[10px] font-black uppercase tracking-widest mt-1 flex items-center gap-2 italic">{data.length > 0 ? <><div className="w-1.5 h-1.5 rounded-full bg-red-600 animate-pulse"></div> CONNECTED: {data.length} CASES</> : "WAITING FOR CONNECTION"} {lastUpdated && <span className="ml-4 opacity-50"><Clock size={10} className="inline mr-1" />{lastUpdated}</span>}</div>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={() => setIsFilterSidebarOpen(true)} className={`flex items-center gap-2 px-5 py-3 rounded-2xl text-xs font-black shadow-sm transition-all border ${selectedResults.length > 0 || filterSup !== 'All' || selectedAgents.length > 0 ? 'bg-red-600 border-red-500 text-white shadow-lg' : 'bg-zinc-800 border-zinc-700 hover:bg-zinc-700'}`}><Filter size={16} /> ตัวกรอง</button>
            <button onClick={() => setShowSync(true)} className="flex items-center gap-2 px-5 py-3 bg-white text-black rounded-2xl text-xs font-black hover:bg-zinc-200 transition-all shadow-xl font-bold"><Settings size={14} /> ตั้งค่า</button>
            <button onClick={() => setIsAuthenticated(false)} className="p-3 bg-zinc-800 rounded-2xl hover:text-red-500 transition-colors"><User size={20} /></button>
          </div>
        </header>

        {data.length > 0 ? (
          <div className="space-y-6">
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                {[
                    { label: 'งานที่ตรวจทั้งหมด', value: filteredData.length, icon: FileText, color: 'text-white', bg: 'bg-zinc-900 border-zinc-800' },
                    { label: 'อัตราผ่านเกณฑ์', value: `${passRate}%`, icon: CheckCircle, color: 'text-emerald-500', bg: 'bg-emerald-500/10 border-emerald-900/20 shadow-emerald-900/5' },
                    { label: 'ควรปรับปรุง', value: filteredData.filter(d=>d.result.startsWith('ควรปรับปรุง')).length, icon: AlertTriangle, color: 'text-amber-500', bg: 'bg-amber-500/10 border-amber-900/20' },
                    { label: 'พบข้อผิดพลาด', value: filteredData.filter(d=>d.result.startsWith('พบข้อผิดพลาด')).length, icon: XCircle, color: 'text-red-600', bg: 'bg-red-600/10 border-red-900/20' }
                ].map((kpi, i) => (
                    <div key={i} className={`p-6 rounded-[2.5rem] border shadow-sm ${kpi.bg}`}>
                      <div className={`w-8 h-8 rounded-xl flex items-center justify-center mb-3 bg-zinc-800/50 ${kpi.color}`}><kpi.icon size={16} /></div>
                      <p className="text-[9px] font-black text-white uppercase tracking-widest">{kpi.label}</p>
                      <h2 className={`text-3xl font-black ${kpi.color} tracking-tighter mt-1 uppercase`}>{kpi.value}</h2>
                    </div>
                ))}
            </div>

            {/* Matrix Table */}
            <div className="bg-zinc-900 rounded-[3rem] shadow-2xl border border-zinc-800 overflow-hidden">
                <div className="p-8 border-b border-zinc-800 bg-zinc-900/50 flex items-center justify-between">
                    <div><h3 className="font-black text-white flex items-center gap-2 italic text-lg uppercase tracking-tight"><TrendingUp size={24} className="text-red-600" /> สรุปคุณภาพพนักงาน x ผลสัมภาษณ์</h3><p className="text-[9px] text-zinc-100 font-black uppercase mt-1 italic tracking-widest flex items-center gap-1 underline decoration-red-600/30">คลิกที่ตัวเลข เพื่อดูรายละเอียดด้านล่าง</p></div>
                </div>
                <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                    <table className="w-full text-left text-sm border-separate border-spacing-0 min-w-[1200px]">
                        <thead className="sticky top-0 bg-zinc-900 z-20 font-black text-white text-[10px] uppercase tracking-widest border-b border-zinc-800 shadow-md">
                            <tr>
                                <th rowSpan="2" className="px-8 py-6 border-b border-zinc-800 border-r border-zinc-800 bg-zinc-900 w-64">พนักงานสัมภาษณ์</th>
                                <th colSpan={RESULT_ORDER.length} className="px-4 py-4 text-center border-b border-zinc-800 bg-zinc-900/40 text-red-600 text-[11px] font-black italic uppercase">สรุปผลการสัมภาษณ์ละเอียด</th>
                                <th rowSpan="2" className="px-8 py-6 text-center bg-zinc-800 text-white border-b border-zinc-800 border-l border-zinc-800">รวม</th>
                            </tr>
                            <tr className="bg-zinc-900/80 text-white">
                                {RESULT_ORDER.map(type => <th key={type} className="px-4 py-3 text-center border-b border-zinc-800 border-r border-zinc-800 max-w-[180px]"><span className="line-clamp-2" title={type}>{formatResultDisplay(type)}</span></th>)}
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-zinc-800/40 font-bold">
                            {agentSummary.map((agent, i) => (
                            <tr key={i} className="hover:bg-zinc-800/50 transition-colors group">
                                <td className="px-8 py-5 text-white border-r border-zinc-800">{agent.name}</td>
                                {RESULT_ORDER.map(type => {
                                const val = agent[type];
                                const isActive = activeCell.agent === agent.name && activeCell.resultType === type;
                                return (
                                    <td key={type} className={`px-4 py-5 text-center border-r border-zinc-800 transition-all ${val > 0 ? 'cursor-pointer hover:bg-zinc-800/60 shadow-inner' : ''} ${isActive ? 'bg-red-600/20 ring-2 ring-inset ring-red-600' : ''}`} onClick={() => val > 0 && handleMatrixClick(agent.name, type)}>
                                        <span className={`text-sm font-black ${val > 0 ? '' : 'text-zinc-900'}`} style={{ color: val > 0 ? getResultColor(type) : undefined }}>{val || '-'}</span>
                                    </td>
                                );
                                })}
                                <td className="px-8 py-5 text-center bg-zinc-800/20 text-white border-l border-zinc-800">{agent.total}</td>
                            </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>

            {/* List and Details */}
            <div id="detail-section" className="bg-zinc-900 rounded-[3rem] shadow-2xl border border-zinc-800 overflow-hidden scroll-mt-6">
                <div className="p-8 border-b border-zinc-800 flex flex-col md:flex-row items-center justify-between gap-6">
                    <div className="space-y-1">
                        <h3 className="font-black text-white uppercase tracking-widest text-xs flex items-center gap-2 italic"><MessageSquare size={16} className="text-red-600" /> ข้อมูลรายเคสแบบละเอียด</h3>
                        {activeCell.agent && (
                            <div className="flex items-center gap-2 mt-2"><span className="text-[9px] font-black px-3 py-1 bg-red-600 text-white rounded-lg shadow-lg uppercase italic animate-pulse">FILTERING: {activeCell.agent}</span><button onClick={() => setActiveCell({ agent: null, resultType: null })}><X size={12}/></button></div>
                        )}
                    </div>
                    <div className="relative w-full md:w-80"><Search className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-400" size={16} /><input type="text" placeholder="ค้นหาพนักงาน หรือ เลขชุด..." className="w-full pl-12 pr-6 py-4 bg-zinc-800/50 border border-zinc-800 rounded-2xl text-sm font-bold focus:ring-2 focus:ring-red-600 outline-none text-white" value={searchTerm} onChange={(e)=>setSearchTerm(e.target.value)} /></div>
                </div>
                <div className="overflow-auto max-h-[1000px] custom-scrollbar">
                    <table className="w-full text-left text-xs font-medium border-separate border-spacing-0">
                    <thead className="sticky top-0 bg-zinc-900 shadow-md z-10 border-b border-zinc-800 font-black text-white uppercase tracking-widest">
                        <tr><th className="px-8 py-5 border-r border-zinc-800/30">วันที่ / เลขชุด</th><th className="px-8 py-5 border-r border-zinc-800/30">พนักงาน</th><th className="px-4 py-5 text-center border-r border-zinc-800/30">ผลสรุป</th><th className="px-8 py-5">QC Comment & Recording</th></tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-800/30">
                        {detailLogs.length > 0 ? detailLogs.slice(0, 150).map((item) => {
                          const isExpanded = expandedCaseId === item.id;
                          const isEditing = editingCase && editingCase.id === item.id;
                          return (
                            <React.Fragment key={item.id}>
                              <tr onClick={() => !isEditing && setExpandedCaseId(isExpanded ? null : item.id)} className={`transition-all group cursor-pointer ${isExpanded ? 'bg-zinc-800' : 'hover:bg-zinc-800/40'}`}>
                                  <td className="px-8 py-6 border-r border-zinc-800/20"><div className="font-black text-white">{item.date}</div><div className="flex items-center gap-1.5 mt-1 bg-zinc-950 px-2 py-0.5 rounded border border-zinc-800 w-fit"><Hash size={10} className="text-red-600" /><span className="text-[11px] font-black text-white">{item.questionnaireNo}</span></div></td>
                                  <td className="px-8 py-6 border-r border-zinc-800/20"><div className="font-black text-white text-sm group-hover:text-red-500 transition-colors flex items-center gap-2">{item.agent} {isExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}</div><div className="text-[9px] text-zinc-500 font-bold mt-0.5 italic">SUP: {item.supervisor} &bull; {item.touchpoint}</div></td>
                                  <td className="px-4 py-6 text-center border-r border-zinc-800/20"><span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[9px] font-black border uppercase" style={{ backgroundColor: `${getResultColor(item.result)}10`, color: getResultColor(item.result), borderColor: `${getResultColor(item.result)}30` }}>{formatResultDisplay(item.result)}</span></td>
                                  <td className="px-8 py-6">
                                      <p className="text-zinc-400 italic max-w-sm truncate group-hover:text-white transition-colors">{item.comment ? `"${item.comment}"` : '-'}</p>
                                      {item.audio && item.audio.includes('http') && (
                                        <a href={item.audio} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()} className="mt-2 inline-flex items-center gap-1 text-red-500 hover:text-red-400 font-black text-[9px] uppercase transition-all hover:translate-x-1">
                                            <PlayCircle size={14} /> Listen Recording
                                        </a>
                                      )}
                                  </td>
                              </tr>
                              {isExpanded && (
                                <tr className="bg-zinc-950/40 animate-in slide-in-from-top-2 duration-300">
                                  <td colSpan={4} className="p-8 border-b border-zinc-800 text-white">
                                    <div className="flex items-center justify-between mb-8">
                                      <div className="flex items-center gap-4">
                                        <div className="p-3 bg-red-600/10 rounded-2xl text-red-600 shadow-inner"><Award /></div>
                                        <div><h4 className="font-black uppercase italic tracking-widest text-sm text-white">Assessment Detail & Recording (ID: {item.questionnaireNo})</h4><p className="text-[9px] text-zinc-500 font-bold uppercase tracking-widest leading-relaxed italic">ข้อมูลการประเมิน 13 หัวข้อ (P:AB) และ สรุปผลละเอียด (M)</p></div>
                                      </div>
                                      <div className="flex gap-2">
                                        {!isEditing ? (<button onClick={() => setEditingCase({...item})} className="flex items-center gap-2 px-6 py-3 bg-zinc-800 hover:bg-zinc-700 text-white rounded-2xl text-[10px] font-black uppercase border border-zinc-700 transition-all"><Edit2 size={12} className="text-red-500"/> แก้ไขข้อมูล</button>) : (
                                          <div className="flex gap-2"><button disabled={isSaving} onClick={handleUpdateCase} className="flex items-center gap-2 px-6 py-3 bg-red-600 hover:bg-red-700 text-white rounded-2xl text-[10px] font-black uppercase shadow-lg shadow-red-900/20">{isSaving ? <RefreshCw className="animate-spin" size={14}/> : <Save size={14}/>} บันทึกลง Sheet</button><button onClick={() => setEditingCase(null)} className="px-6 py-3 bg-zinc-800 text-white rounded-2xl text-[10px] font-black uppercase transition-all">ยกเลิก</button></div>
                                        )}
                                      </div>
                                    </div>

                                    {item.audio && item.audio.includes('http') && (
                                        <div className="mb-8 p-6 bg-zinc-950/50 border border-zinc-800 rounded-[2rem] flex items-center justify-between shadow-inner">
                                            <div className="flex items-center gap-3">
                                                <div className="p-3 bg-red-600 text-white rounded-2xl animate-pulse"><PlayCircle size={20} /></div>
                                                <div>
                                                    <p className="text-[10px] font-black uppercase text-zinc-500 tracking-widest">Audio Recording File</p>
                                                    <p className="text-xs font-black text-white">คลิกปุ่มขวาเพื่อฟังไฟล์เสียงที่บันทึกไว้</p>
                                                </div>
                                            </div>
                                            <a href={item.audio} target="_blank" rel="noopener noreferrer" className="px-6 py-3 bg-red-600 hover:bg-red-700 text-white rounded-2xl text-[10px] font-black uppercase transition-all flex items-center gap-2 shadow-lg shadow-red-900/20">
                                                <ExternalLink size={14} /> OPEN AUDIO PLAYER
                                            </a>
                                        </div>
                                    )}

                                    <div className="mb-8 p-6 bg-zinc-900/80 border border-zinc-800 rounded-[2rem] shadow-inner">
                                      <div className="flex items-center gap-2 mb-4 text-amber-500 font-black text-[10px] uppercase italic tracking-widest"><Star size={16} /> สรุปผลการสัมภาษณ์แบบละเอียด (Column M)</div>
                                      {isEditing ? (<div className="relative"><select className="w-full p-4 bg-zinc-950 border border-red-600/30 rounded-2xl text-[11px] font-black text-white focus:ring-2 focus:ring-red-600 outline-none appearance-none shadow-lg shadow-red-900/10" value={editingCase.result} onChange={e => setEditingCase({...editingCase, result: e.target.value})}>{RESULT_ORDER.map(opt => <option key={opt} value={opt}>{opt}</option>)}</select><ChevronDown size={14} className="absolute right-4 top-1/2 -translate-y-1/2 text-zinc-600 pointer-events-none" /></div>) : (<p className="text-sm font-black italic text-zinc-100 bg-zinc-950/50 p-4 rounded-xl border border-zinc-800">{item.result}</p>)}
                                    </div>
                                    <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3">
                                      {(isEditing ? editingCase : item).evaluations.map((evalItem, eIdx) => (
                                        <div key={eIdx} className={`bg-zinc-900 border p-3 rounded-2xl transition-all ${isEditing ? 'border-red-600/50 ring-1 ring-red-600/20 shadow-md' : 'border-zinc-800'}`}>
                                          <p className="text-[9px] font-black text-zinc-500 uppercase tracking-tighter mb-2 truncate" title={evalItem.label}>{evalItem.label}</p>
                                          {isEditing ? (
                                            <div className="relative">
                                                <select className="w-full bg-zinc-950 border border-zinc-800 rounded-lg px-2 py-1.5 text-[10px] font-black text-white appearance-none" value={evalItem.value} onChange={(e) => { const newEvals = [...editingCase.evaluations]; newEvals[eIdx].value = e.target.value; setEditingCase({...editingCase, evaluations: newEvals}); }}>{SCORE_OPTIONS.map(opt => (<option key={opt} value={opt}>{SCORE_LABELS[opt]}</option>))}</select>
                                                <ChevronDown size={10} className="absolute right-1 top-1/2 -translate-y-1/2 text-zinc-600 pointer-events-none" />
                                            </div>
                                          ) : (<div className={`text-sm font-black italic tracking-widest ${evalItem.value === '5' || evalItem.value === '4' ? 'text-emerald-500' : (evalItem.value === '1' || evalItem.value === '2') ? 'text-red-500' : 'text-zinc-200'}`}>{SCORE_LABELS[evalItem.value] || evalItem.value}</div>)}
                                        </div>
                                      ))}
                                    </div>
                                    <div className="mt-8">
                                      <p className="text-[10px] font-black text-red-600 uppercase mb-2 italic tracking-widest flex items-center gap-2"><MessageSquare size={12}/> QC Full Comment (Column AC)</p>
                                      {isEditing ? (<textarea className="w-full bg-zinc-950 border border-zinc-800 rounded-[1.5rem] p-6 text-sm italic text-white outline-none min-h-[120px] focus:ring-2 focus:ring-red-600 shadow-inner" value={editingCase.comment} onChange={e => setEditingCase({...editingCase, comment: e.target.value})}/>) : (<div className="p-4 bg-zinc-800/30 rounded-2xl border border-zinc-800 border-l-4 border-l-red-600 shadow-sm"><p className="text-sm text-zinc-100 font-medium italic leading-relaxed">"{item.comment || 'ไม่มีคอมเมนต์'}"</p></div>)}
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                        }) : (<tr><td colSpan={4} className="px-8 py-24 text-center text-white font-black uppercase italic tracking-widest text-lg opacity-20">No Matching Analysis Data</td></tr>)}
                    </tbody>
                    </table>
                </div>
            </div>
          </div>
        ) : (!loading && <div className="bg-zinc-900 rounded-[4rem] border-2 border-dashed border-zinc-800 py-32 text-center shadow-inner"><Database size={40} className="text-zinc-600 mx-auto mb-8" /><h2 className="text-2xl font-black text-white uppercase italic tracking-widest">Waiting for Signal</h2><p className="text-zinc-100 text-xs mt-3 max-w-sm mx-auto font-black uppercase tracking-widest italic leading-relaxed">กรุณากดปุ่ม "ตั้งค่า" ด้านบน และวางลิงก์ CSV จาก Google Sheets</p></div>)}
      </div>

      {/* Sync Modal */}
      {(showSync || error) && (
        <div className="fixed inset-0 z-[100] bg-black/95 flex items-center justify-center p-4">
            <div className="bg-zinc-900 w-full max-w-2xl rounded-[3rem] p-10 border border-zinc-800 shadow-2xl relative animate-in zoom-in duration-300">
                <div className="absolute top-0 left-0 w-full h-1 bg-red-600"></div>
                <div className="flex items-center justify-between mb-10"><h3 className="text-xl font-black flex items-center gap-3 text-white uppercase italic tracking-tight">{error ? 'Connection Error' : 'System Configuration'}</h3>{data.length > 0 && <button onClick={() => {setShowSync(false); setError(null);}} className="text-zinc-600 hover:text-white transition-colors"><X size={28} /></button>}</div>
                <div className="space-y-6">
                    <div className="space-y-2"><label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest pl-2 flex items-center gap-2 italic"><Database size={12}/> Google Sheets CSV Link</label><input type="text" className="w-full px-6 py-4 bg-zinc-950 border border-zinc-800 rounded-2xl text-xs font-bold text-white outline-none focus:ring-2 focus:ring-red-600 shadow-inner font-sans" value={sheetUrl} onChange={e=>setSheetUrl(e.target.value)} /></div>
                    <div className="space-y-2"><label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest pl-2 flex items-center gap-2 italic"><Save size={12}/> Apps Script API Link</label><input type="text" className="w-full px-6 py-4 bg-zinc-950 border border-zinc-800 rounded-2xl text-xs font-bold text-white outline-none focus:ring-2 focus:ring-red-600 shadow-inner font-sans" value={appsScriptUrl} onChange={e=>{setAppsScriptUrl(e.target.value); localStorage.setItem('apps_script_url', e.target.value);}} /></div>
                    <div className="flex gap-4 pt-4"><button onClick={() => fetchFromSheet(sheetUrl)} disabled={loading || !sheetUrl} className="flex-1 py-5 bg-red-600 hover:bg-red-700 text-white rounded-[2.5rem] font-black uppercase shadow-xl shadow-red-900/30 text-sm tracking-widest italic flex items-center justify-center gap-2 transition-all">{loading ? <RefreshCw className="animate-spin" size={18}/> : <RefreshCw size={18}/>} {loading ? 'SYNCING...' : 'CONNECT & ANALYZE'}</button></div>
                    {error && <p className="text-center text-[10px] text-red-500 font-black mt-2 uppercase italic tracking-tighter animate-pulse bg-red-600/5 p-4 rounded-2xl border border-red-600/10 shadow-lg">{error}</p>}
                </div>
            </div>
        </div>
      )}
    </div>
  );
};

export default App;