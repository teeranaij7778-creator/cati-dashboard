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
 * แก้ไข: เพิ่มตัวกรองพนักงานสัมภาษณ์ที่หายไปใน Sidebar
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

// Helper: ดึงข้อความนอกวงเล็บ
const formatResultDisplay = (text) => {
  if (!text) return '-';
  return text.split('(')[0].trim();
};

const COLORS = {
  'ดีเยี่ยม': '#3B82F6',     
  'ผ่านเกณฑ์': '#10B981',   
  'ควรปรับปรุง': '#F59E0B', 
  'พบข้อผิดพลาด': '#EF4444', 
  'ไม่ผ่านเกณฑ์': '#B91C1C', 
};

// Helper: ดึงสีตามหมวดหมู่
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
  const [evalHeaders, setEvalHeaders] = useState([]); 
  const [loading, setLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState(null);
  const [sheetUrl, setSheetUrl] = useState(localStorage.getItem('qc_sheet_url') || DEFAULT_SHEET_URL);
  const [appsScriptUrl, setAppsScriptUrl] = useState(localStorage.getItem('apps_script_url') || DEFAULT_APPS_SCRIPT_URL);
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
  const [expandedCaseId, setExpandedCaseId] = useState(null);
  const [editingCase, setEditingCase] = useState(null); 

  useEffect(() => {
    if (!document.getElementById('thai-font-link')) {
      const fontLink = document.createElement('link');
      fontLink.id = 'thai-font-link';
      fontLink.rel = 'stylesheet';
      fontLink.href = 'https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,100;0,200;0,300;0,400;0,500;0,600;0,700;0,800;1,100&display=swap';
      document.head.appendChild(fontLink);
    }
  }, []);

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
      const evaluationsList = headers.slice(15, 28);
      setEvalHeaders(evaluationsList);

      const getIdx = (name) => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
      let qNoIdx = headers.findIndex(h => h.toLowerCase().includes("questionnaire") || h.toLowerCase().includes("no.") || h.toLowerCase().includes("เลขชุด"));
      if (qNoIdx === -1) qNoIdx = 3; 

      const idx = {
        year: getIdx("Year"), month: getIdx("เดือน"), date: getIdx("วันที่สัมภาษณ์"), touchpoint: getIdx("TOUCH_POINT"), type: getIdx("AC / BC"), sup: getIdx("Supervisor"), agent: getIdx("Interviewer"),
        questionnaireNo: qNoIdx,
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
          let cleanResult = rawResult;
          const matchedResult = RESULT_ORDER.find(opt => rawResult.includes(opt.split(':')[0].trim()));
          if (matchedResult) cleanResult = matchedResult;
          
          const agentId = row[idx.agent]?.toString().trim() || 'Unknown';
          const agentName = (idx.agentName !== -1) ? row[idx.agentName]?.toString().trim() : '';
          let displayAgent = agentName && agentName !== agentId ? `${agentId} : ${agentName}` : agentId;
          
          const evaluations = evaluationsList.map((header, i) => ({
            label: header,
            value: row[15 + i] || '-'
          }));

          return {
            id: index, rowIndex: index + headerIdx + 2, 
            year: row[idx.year] || 'N/A', month: row[idx.month] || 'N/A', date: row[idx.date] || 'N/A', touchpoint: row[idx.touchpoint] || 'N/A', type: row[idx.type] || 'N/A', supervisor: row[idx.sup] || 'N/A', agent: displayAgent, questionnaireNo: row[idx.questionnaireNo] || '-', audio: row[idx.audio] || '', result: cleanResult, comment: row[idx.comment] || '',
            evaluations: evaluations
          };
        });
      setData(parsedData);
      setLastUpdated(new Date().toLocaleTimeString('th-TH'));
      localStorage.setItem('qc_sheet_url', finalUrl);
      setShowSync(false);
    } catch (err) { setError(err.message); } finally { setLoading(false); }
  };

  const handleUpdateCase = async () => {
    if (!appsScriptUrl) {
      alert("กรุณาตั้งค่า Apps Script URL ในเมนูตั้งค่าก่อนบันทึก");
      return;
    }
    setIsSaving(true);
    try {
      const updateData = {
        rowIndex: editingCase.rowIndex,
        result: editingCase.result,
        evaluations: editingCase.evaluations.map(e => e.value),
        comment: editingCase.comment
      };
      await fetch(appsScriptUrl, {
        method: 'POST',
        mode: 'no-cors', 
        cache: 'no-cache',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(updateData)
      });
      setTimeout(() => {
        setIsSaving(false);
        setEditingCase(null);
        fetchFromSheet(sheetUrl);
      }, 2000);
    } catch (err) {
      alert("เกิดข้อผิดพลาดในการบันทึก: " + err.message);
      setIsSaving(false);
    }
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
      const matchesSearch = item.agent.toLowerCase().includes(searchTerm.toLowerCase()) || 
                           item.comment.toLowerCase().includes(searchTerm.toLowerCase()) ||
                           item.questionnaireNo.toLowerCase().includes(searchTerm.toLowerCase());
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
      name: key, 
      displayLabel: formatResultDisplay(key), 
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
      <div className="min-h-screen bg-zinc-950 flex items-center justify-center p-4 font-sans text-white">
        <div className="bg-zinc-900 p-8 md:p-12 rounded-[2.5rem] shadow-2xl w-full max-w-sm border border-zinc-800 relative overflow-hidden">
          <div className="absolute top-0 left-0 w-full h-1 bg-red-600"></div>
          <div className="flex justify-center mb-10"><IntageLogo className="scale-150" /></div>
          <h2 className="text-xl font-black text-center text-zinc-100 uppercase tracking-[0.2em] mb-2 italic">CATI CES 2026</h2>
          <form onSubmit={handleLogin} className="space-y-4">
            <input type="text" value={inputUser} onChange={(e) => setInputUser(e.target.value)} className="w-full px-6 py-4 bg-zinc-800 border border-zinc-700 rounded-2xl font-bold text-white outline-none focus:ring-2 focus:ring-red-600" placeholder="Username" />
            <input type="password" value={inputPass} onChange={(e) => setInputPass(e.target.value)} className="w-full px-6 py-4 bg-zinc-800 border border-zinc-700 rounded-2xl font-bold text-white outline-none focus:ring-2 focus:ring-red-600" placeholder="Password" />
            {loginError && <p className="text-red-500 text-xs font-bold text-center">{loginError}</p>}
            <button type="submit" className="w-full py-4 bg-red-600 hover:bg-red-700 text-white rounded-2xl font-black uppercase tracking-wider transition-all mt-4">เข้าสู่ระบบ</button>
          </form>
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
                   <h3 className="font-black text-white uppercase italic tracking-tight">ตัวกรองข้อมูล</h3>
                </div>
                <button onClick={() => setIsFilterSidebarOpen(false)} className="p-2 hover:bg-zinc-800 rounded-full text-zinc-100 transition-colors"><X size={20} /></button>
             </div>
             <div className="space-y-6">
                <button onClick={() => { setFilterSup('All'); setSelectedAgents([]); setSelectedResults([]); setSelectedTypes([]); setSelectedMonth('All'); setActiveCell({ agent: null, resultType: null }); }} className="w-full py-2 text-xs font-black text-red-500 bg-red-600/10 hover:bg-red-600/20 rounded-xl transition-colors border border-red-600/20 uppercase tracking-widest">Reset All Filters</button>
                
                <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest pl-2">เดือน (Month)</label>
                    <select className="w-full px-4 py-3 bg-zinc-800 border border-zinc-700 rounded-2xl text-xs font-bold" value={selectedMonth} onChange={(e)=>setSelectedMonth(e.target.value)}>
                        <option value="All">ทุกเดือน</option>{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}
                    </select>
                </div>
                <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest pl-2">Supervisor</label>
                    <select className="w-full px-4 py-3 bg-zinc-800 border border-zinc-700 rounded-2xl text-xs font-bold" value={filterSup} onChange={(e) => { setFilterSup(e.target.value); setSelectedAgents([]); }}>
                        <option value="All">ทุก Supervisor</option>{availableSups.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                </div>

                {/* พนักงานสัมภาษณ์ Filter */}
                <div className="space-y-2">
                    <div className="flex items-center justify-between pl-2">
                      <label className="text-[10px] font-black uppercase tracking-widest">พนักงานสัมภาษณ์</label>
                      <div className="flex gap-2">
                        <button onClick={() => setSelectedAgents(availableAgents)} className="text-[9px] font-bold text-zinc-400 hover:text-white">เลือกทั้งหมด</button>
                        <button onClick={() => setSelectedAgents([])} className="text-[9px] font-bold text-red-500/70 hover:text-red-500">ล้าง</button>
                      </div>
                    </div>
                    <div className="bg-zinc-800 border border-zinc-700 rounded-2xl p-2 max-h-48 overflow-y-auto custom-scrollbar">
                        {availableAgents.map(agent => {
                          const isSelected = selectedAgents.includes(agent);
                          return (
                            <div key={agent} onClick={() => setSelectedAgents(prev => isSelected ? prev.filter(a => a !== agent) : [...prev, agent])} className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-[10px] font-bold transition-colors mb-1 ${isSelected ? 'bg-zinc-700 text-white shadow-sm' : 'hover:bg-zinc-800/50 text-zinc-400'}`}>
                              {isSelected ? <CheckSquare size={14} className="text-red-500 shrink-0" /> : <Square size={14} className="text-zinc-600 shrink-0" />}
                              <span className="truncate">{agent}</span>
                            </div>
                          )
                        })}
                    </div>
                </div>

                <div className="space-y-2">
                    <div className="flex items-center justify-between pl-2">
                        <label className="text-[10px] font-black uppercase tracking-widest">ผลการสัมภาษณ์</label>
                        <div className="flex gap-2">
                            <button onClick={() => setSelectedResults(RESULT_ORDER)} className="text-[9px] font-bold text-zinc-400 hover:text-white">เลือกทั้งหมด</button>
                            <button onClick={() => setSelectedResults([])} className="text-[9px] font-bold text-red-500/70 hover:text-red-500">ล้าง</button>
                        </div>
                    </div>
                    <div className="bg-zinc-800 border border-zinc-700 rounded-2xl p-2 max-h-48 overflow-y-auto custom-scrollbar">
                        {RESULT_ORDER.map(res => (
                            <div key={res} onClick={() => setSelectedResults(prev => prev.includes(res) ? prev.filter(r => r !== res) : [...prev, res])} className={`flex items-start gap-2 p-2 rounded-xl cursor-pointer text-[10px] font-bold transition-colors mb-1 ${selectedResults.includes(res) ? 'bg-zinc-700 ring-1 ring-zinc-600 shadow-md' : 'hover:bg-zinc-800/50'}`}>
                              {selectedResults.includes(res) ? <CheckSquare size={14} className="text-red-500 mt-0.5 shrink-0" /> : <Square size={14} className="text-zinc-500 mt-0.5 shrink-0" />}
                              <span style={{ color: selectedResults.includes(res) ? getResultColor(res) : undefined }}>{formatResultDisplay(res)}</span>
                            </div>
                        ))}
                    </div>
                </div>
             </div>
          </div>
       </aside>

      <div className="max-w-7xl mx-auto space-y-6">
        <header className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 bg-zinc-900 p-6 rounded-[2.5rem] shadow-xl border border-zinc-800">
          <div className="flex items-center gap-6">
            <div className="p-3 bg-zinc-800 rounded-2xl border border-zinc-700"><IntageLogo /></div>
            <div>
              <h1 className="text-xl font-black text-white tracking-tight flex items-center gap-2 uppercase italic">QC REPORT 2026 {loading && <RefreshCw size={18} className="animate-spin text-red-600" />}</h1>
              <div className="text-white text-[10px] font-black flex items-center gap-2 uppercase tracking-widest mt-1">
                  {data.length > 0 ? <><div className="w-1.5 h-1.5 rounded-full bg-red-600 animate-pulse"></div> CONNECTED: {data.length} CASES</> : "WAITING FOR CONNECTION"}
                  {lastUpdated && <span className="ml-4 text-zinc-500 italic lowercase tracking-tight"><Clock size={10} className="inline mr-1" />{lastUpdated}</span>}
              </div>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={() => setIsFilterSidebarOpen(true)} className={`flex items-center gap-2 px-5 py-3 rounded-2xl text-xs font-black shadow-sm transition-all border ${selectedResults.length > 0 || filterSup !== 'All' || selectedAgents.length > 0 ? 'bg-red-600 border-red-500 text-white shadow-lg shadow-red-900/20' : 'bg-zinc-800 border-zinc-700 text-white hover:bg-zinc-700'}`}>
              <Filter size={16} /> ตัวกรอง { (selectedResults.length > 0 || filterSup !== 'All' || selectedAgents.length > 0) && <span className="w-1.5 h-1.5 rounded-full bg-white animate-bounce"></span>}
            </button>
            <button onClick={() => setShowSync(true)} className="flex items-center gap-2 px-5 py-3 bg-white text-black rounded-2xl text-xs font-black hover:bg-zinc-200 transition-all shadow-xl shadow-white/5 font-bold"><Settings size={14} /> ตั้งค่า</button>
            <button onClick={() => setIsAuthenticated(false)} className="p-3 bg-zinc-800 text-zinc-100 rounded-2xl hover:text-red-500 border border-zinc-700 transition-colors" title="ออกจากระบบ"><User size={20} /></button>
          </div>
        </header>

        {data.length > 0 ? (
          <div className="space-y-6">
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                {[
                    { label: 'งานที่ตรวจทั้งหมด', value: filteredData.length, icon: FileText, color: 'text-white', bg: 'bg-zinc-900 border-zinc-800' },
                    { label: 'อัตราผ่านเกณฑ์', value: `${passRate}%`, icon: CheckCircle, color: 'text-emerald-500', bg: 'bg-emerald-500/10 border-emerald-900/20 shadow-lg shadow-emerald-900/5' },
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

            {/* Matrix Table Summary */}
            <div className="bg-zinc-900 rounded-[3rem] shadow-2xl border border-zinc-800 overflow-hidden">
                <div className="p-8 border-b border-zinc-800 bg-zinc-900/50 flex items-center justify-between">
                    <div>
                        <h3 className="font-black text-white flex items-center gap-2 italic text-lg uppercase tracking-tight"><TrendingUp size={24} className="text-red-600" /> สรุปคุณภาพพนักงาน x ผลสัมภาษณ์</h3>
                        <p className="text-[9px] text-zinc-100 font-black uppercase mt-1 italic tracking-widest flex items-center gap-1 underline decoration-red-600/30">คลิกที่ตัวเลข เพื่อดูรายละเอียดด้านล่าง</p>
                    </div>
                </div>
                <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                    <table className="w-full text-left text-sm border-separate border-spacing-0 min-w-[1200px]">
                        <thead className="sticky top-0 bg-zinc-900 z-20 font-black text-white text-[10px] uppercase tracking-widest border-b border-zinc-800 shadow-md">
                            <tr>
                                <th rowSpan="2" className="px-8 py-6 border-b border-zinc-800 border-r border-zinc-800 bg-zinc-900 w-64">พนักงานสัมภาษณ์</th>
                                <th colSpan={RESULT_ORDER.length} className="px-4 py-4 text-center border-b border-zinc-800 bg-zinc-900/40 text-red-600 text-[11px] font-black italic uppercase">สรุปผลการสัมภาษณ์ละเอียด</th>
                                <th rowSpan="2" className="px-8 py-6 text-center bg-zinc-800 text-white border-b border-zinc-800 border-l border-zinc-800">รวม</th>
                            </tr>
                            <tr className="bg-zinc-900/80">
                                {RESULT_ORDER.map(type => <th key={type} className="px-4 py-3 text-center border-b border-zinc-800 border-r border-zinc-800 text-white max-w-[180px]"><span className="line-clamp-2" title={type}>{formatResultDisplay(type)}</span></th>)}
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

            {/* Case List with Detail Section */}
            <div id="detail-section" className="bg-zinc-900 rounded-[3rem] shadow-2xl border border-zinc-800 overflow-hidden scroll-mt-6">
                <div className="p-8 border-b border-zinc-800 flex flex-col md:flex-row items-center justify-between gap-6">
                    <div className="space-y-1">
                        <h3 className="font-black text-white uppercase tracking-widest text-xs flex items-center gap-2 italic"><MessageSquare size={16} className="text-red-600" /> ข้อมูลรายเคสแบบละเอียด</h3>
                        {activeCell.agent && (
                            <div className="flex items-center gap-2 mt-2">
                                <span className="text-[9px] font-black px-3 py-1 bg-red-600 text-white rounded-lg shadow-lg shadow-red-900/20 uppercase italic tracking-tighter animate-pulse">FILTERING: {activeCell.agent} ({formatResultDisplay(activeCell.resultType)})</span>
                                <button onClick={() => setActiveCell({ agent: null, resultType: null })} className="p-1 hover:bg-zinc-800 rounded text-zinc-100 hover:text-red-500 transition-colors"><X size={12}/></button>
                            </div>
                        )}
                    </div>
                    <div className="relative w-full md:w-80">
                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-400" size={16} />
                        <input type="text" placeholder="ค้นหาเลขชุด, พนักงาน..." className="w-full pl-12 pr-6 py-4 bg-zinc-800/50 border border-zinc-800 rounded-2xl text-sm font-bold focus:ring-2 focus:ring-red-600 outline-none" value={searchTerm} onChange={(e)=>setSearchTerm(e.target.value)} />
                    </div>
                </div>
                <div className="overflow-auto max-h-[1000px] custom-scrollbar">
                    <table className="w-full text-left text-xs font-medium border-separate border-spacing-0">
                    <thead className="sticky top-0 bg-zinc-900 shadow-md z-10 border-b border-zinc-800 font-black text-white uppercase tracking-widest">
                        <tr>
                            <th className="px-8 py-5 border-r border-zinc-800/30">วันที่ / เลขชุด</th>
                            <th className="px-8 py-5 border-r border-zinc-800/30">พนักงาน</th>
                            <th className="px-4 py-5 text-center border-r border-zinc-800/30">ผลสรุป</th>
                            <th className="px-8 py-5">QC Comment</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-800/30">
                        {detailLogs.length > 0 ? detailLogs.slice(0, 150).map((item) => {
                          const isExpanded = expandedCaseId === item.id;
                          const isCurrentlyEditing = editingCase && editingCase.id === item.id;
                          return (
                            <React.Fragment key={item.id}>
                              <tr onClick={() => !isCurrentlyEditing && setExpandedCaseId(isExpanded ? null : item.id)} className={`transition-all group cursor-pointer ${isExpanded ? 'bg-zinc-800' : 'hover:bg-zinc-800/40'}`}>
                                  <td className="px-8 py-6 border-r border-zinc-800/20">
                                      <div className="font-black text-white">{item.date}</div>
                                      <div className="flex items-center gap-1.5 mt-1 bg-zinc-950 px-2 py-0.5 rounded border border-zinc-800 w-fit shadow-lg shadow-red-900/5">
                                          <Hash size={10} className="text-red-600" /><span className="text-[11px] font-black text-white">{item.questionnaireNo}</span>
                                      </div>
                                  </td>
                                  <td className="px-8 py-6 border-r border-zinc-800/20">
                                      <div className="font-black text-white text-sm group-hover:text-red-500 transition-colors flex items-center gap-2">{item.agent} {isExpanded ? <ChevronDown size={14} className="text-zinc-500" /> : <ChevronRight size={14} className="text-zinc-700" />}</div>
                                      <div className="text-[9px] text-zinc-500 font-bold uppercase mt-0.5 italic">{item.type} &bull; {item.touchpoint} &bull; SUP: {item.supervisor}</div>
                                  </td>
                                  <td className="px-4 py-6 text-center border-r border-zinc-800/20">
                                      <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[9px] font-black border uppercase shadow-sm" style={{ backgroundColor: `${getResultColor(item.result)}10`, color: getResultColor(item.result), borderColor: `${getResultColor(item.result)}30` }}>
                                          {formatResultDisplay(item.result)}
                                      </span>
                                  </td>
                                  <td className="px-8 py-6">
                                      <p className="text-zinc-400 font-semibold italic max-w-sm leading-relaxed truncate group-hover:text-white transition-colors">{item.comment ? `"${item.comment}"` : '-'}</p>
                                  </td>
                              </tr>
                              {isExpanded && (
                                <tr className="bg-zinc-950/40 animate-in slide-in-from-top-2 duration-300">
                                  <td colSpan={4} className="p-8 border-b border-zinc-800">
                                    <div className="flex items-center justify-between mb-8">
                                      <div className="flex items-center gap-4">
                                        <div className="p-3 bg-red-600/10 rounded-2xl text-red-600 shadow-inner"><Award /></div>
                                        <div><h4 className="font-black text-white uppercase italic tracking-widest text-sm">Assessment Detail & Edit Mode (ID: {item.questionnaireNo})</h4><p className="text-[9px] text-zinc-500 font-bold uppercase tracking-widest leading-relaxed">ข้อมูลการประเมิน 13 หัวข้อ (P:AB) และ สรุปผลละเอียด (M)</p></div>
                                      </div>
                                      <div className="flex gap-2">
                                        {!isCurrentlyEditing ? (
                                          <button onClick={() => setEditingCase({...item})} className="flex items-center gap-2 px-6 py-3 bg-zinc-800 hover:bg-zinc-700 text-white rounded-2xl text-[10px] font-black uppercase transition-all border border-zinc-700"><Edit2 size={12} className="text-red-500"/> แก้ไขข้อมูล</button>
                                        ) : (
                                          <div className="flex gap-2">
                                            <button disabled={isSaving} onClick={handleUpdateCase} className="flex items-center gap-2 px-6 py-3 bg-red-600 hover:bg-red-700 text-white rounded-2xl text-[10px] font-black uppercase transition-all shadow-lg shadow-red-900/20">
                                              {isSaving ? <RefreshCw className="animate-spin" size={14}/> : <Save size={14}/>} บันทึกลง Sheet
                                            </button>
                                            <button onClick={() => setEditingCase(null)} className="px-6 py-3 bg-zinc-800 text-white rounded-2xl text-[10px] font-black uppercase transition-all">ยกเลิก</button>
                                          </div>
                                        )}
                                      </div>
                                    </div>

                                    {/* Edit Result (M) */}
                                    <div className="mb-8 p-6 bg-zinc-900/80 border border-zinc-800 rounded-[2rem] shadow-inner">
                                      <div className="flex items-center gap-2 mb-4 text-amber-500"><Star size={16} /><span className="text-[10px] font-black uppercase tracking-widest italic leading-relaxed">สรุปผลการสัมภาษณ์แบบละเอียด (Column M)</span></div>
                                      {isCurrentlyEditing ? (
                                        <div className="relative"><select className="w-full p-4 bg-zinc-950 border border-red-600/30 rounded-2xl text-[11px] font-black text-white focus:ring-2 focus:ring-red-600 outline-none appearance-none shadow-lg shadow-red-900/10" value={editingCase.result} onChange={e => setEditingCase({...editingCase, result: e.target.value})}>{RESULT_ORDER.map(opt => <option key={opt} value={opt}>{opt}</option>)}</select><ChevronDown size={14} className="absolute right-4 top-1/2 -translate-y-1/2 text-zinc-600 pointer-events-none" /></div>
                                      ) : (<p className="text-sm font-black italic text-zinc-100 bg-zinc-950/50 p-4 rounded-xl border border-zinc-800 leading-relaxed shadow-sm">{item.result}</p>)}
                                    </div>

                                    {/* Evaluations (P-AB) */}
                                    <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3">
                                      {(isCurrentlyEditing ? editingCase : item).evaluations.map((evalItem, eIdx) => (
                                        <div key={eIdx} className={`bg-zinc-900/80 border p-3 rounded-2xl transition-all group/item ${isCurrentlyEditing ? 'border-red-600/50 ring-1 ring-red-600/20 shadow-lg shadow-red-900/10' : 'border-zinc-800'}`}>
                                          <p className="text-[9px] font-black text-zinc-500 uppercase tracking-tighter mb-2 line-clamp-1 group-hover/item:text-zinc-300 transition-colors" title={evalItem.label}>{evalItem.label}</p>
                                          {isCurrentlyEditing ? (
                                            <div className="relative">
                                                <select className="w-full bg-zinc-950 border border-zinc-800 rounded-lg px-2 py-1.5 text-[10px] font-black text-white outline-none appearance-none" value={evalItem.value} onChange={(e) => { const newEvals = [...editingCase.evaluations]; newEvals[eIdx].value = e.target.value; setEditingCase({...editingCase, evaluations: newEvals}); }}>
                                                    {SCORE_OPTIONS.map(opt => (<option key={opt} value={opt}>{SCORE_LABELS[opt]}</option>))}
                                                </select>
                                                <ChevronDown size={10} className="absolute right-1 top-1/2 -translate-y-1/2 text-zinc-600 pointer-events-none"/>
                                            </div>
                                          ) : (<div className={`text-sm font-black italic tracking-widest ${evalItem.value === '5' || evalItem.value === '4' ? 'text-emerald-500' : (evalItem.value === '1' || evalItem.value === '2') ? 'text-red-500' : 'text-zinc-200'}`}>{SCORE_LABELS[evalItem.value] || evalItem.value}</div>)}
                                        </div>
                                      ))}
                                    </div>

                                    {/* Comment (AC) */}
                                    <div className="mt-8">
                                      <p className="text-[10px] font-black text-red-600 uppercase mb-2 italic tracking-widest flex items-center gap-2"><MessageSquare size={12}/> QC Full Comment (Column AC)</p>
                                      {isCurrentlyEditing ? (<textarea className="w-full bg-zinc-950 border border-zinc-800 rounded-[1.5rem] p-6 text-sm italic text-white outline-none min-h-[120px] focus:ring-2 focus:ring-red-600 transition-all shadow-inner" value={editingCase.comment} onChange={e => setEditingCase({...editingCase, comment: e.target.value})}/>) : (
                                        <div className="p-4 bg-zinc-800/30 rounded-2xl border border-zinc-800 border-l-4 border-l-red-600 shadow-sm"><p className="text-sm text-zinc-100 font-medium italic leading-relaxed">"{item.comment || 'ไม่มีคอมเมนต์'}"</p></div>
                                      )}
                                    </div>
                                    {isCurrentlyEditing && <div className="mt-4 flex items-center gap-2 p-3 bg-red-600/5 border border-red-600/10 rounded-xl text-[10px] text-zinc-400 italic leading-relaxed"><Info size={14} className="text-red-500 shrink-0" />บันทึกข้อมูลจะทำการอัปเดตบรรทัดที่ {item.rowIndex} ในชีต {sheetUrl.split('gid=')[1]?.split('&')[0]} ผ่าน Apps Script API</div>}
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
                <div className="bg-zinc-800 w-24 h-24 rounded-[3rem] flex items-center justify-center mx-auto mb-8 shadow-2xl border border-zinc-700"><Database size={40} className="text-zinc-600" /></div>
                <h2 className="text-2xl font-black text-white uppercase italic tracking-widest">Waiting for Signal</h2>
                <p className="text-zinc-100 text-xs mt-3 max-w-sm mx-auto font-black uppercase tracking-widest leading-relaxed">กรุณากดปุ่ม "ตั้งค่า" ด้านบน และวางลิงก์ CSV จาก Google Sheets</p>
            </div>
          )
        )}
      </div>

      {/* Connection Modal */}
      {(showSync || error) && (
        <div className="fixed inset-0 z-[100] bg-black/90 backdrop-blur-xl flex items-center justify-center p-4">
            <div className="bg-zinc-900 w-full max-w-2xl rounded-[3rem] p-10 border border-zinc-800 shadow-2xl relative overflow-hidden font-sans animate-in zoom-in duration-300">
                <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-red-600 via-red-900 to-zinc-900"></div>
                <div className="flex items-center justify-between mb-10"><h3 className="text-xl font-black flex items-center gap-3 text-white uppercase italic tracking-tight">{error ? <AlertCircle className="text-red-500" /> : <Settings className="text-red-600" />}{error ? 'Connection Error' : 'System Configuration'}</h3>{data.length > 0 && <button onClick={() => {setShowSync(false); setError(null);}} className="text-zinc-600 hover:text-white transition-colors"><X size={28} /></button>}</div>
                <div className="space-y-6">
                    <div className="space-y-2"><label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest pl-2">Google Sheets CSV Link</label><input type="text" className="w-full px-6 py-4 bg-zinc-950 border border-zinc-800 rounded-2xl text-xs font-bold text-white outline-none focus:ring-2 focus:ring-red-600" value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)} /></div>
                    <div className="space-y-2"><label className="text-[10px] font-black text-zinc-100 uppercase tracking-widest pl-2">Apps Script API Link</label><input type="text" className="w-full px-6 py-4 bg-zinc-950 border border-zinc-800 rounded-2xl text-xs font-bold text-white outline-none focus:ring-2 focus:ring-red-600" value={appsScriptUrl} onChange={(e) => { setAppsScriptUrl(e.target.value); localStorage.setItem('apps_script_url', e.target.value); }} /></div>
                    <div className="flex gap-4 pt-4"><button onClick={() => fetchFromSheet(sheetUrl)} disabled={loading || !sheetUrl} className="flex-1 py-5 bg-red-600 hover:bg-red-700 text-white rounded-[2rem] font-black uppercase shadow-xl shadow-red-900/30 disabled:opacity-20 text-sm tracking-widest italic">{loading ? 'SYNCING...' : 'CONNECT & ANALYZE'}</button>{data.length > 0 && <button onClick={clearConnection} className="p-5 bg-zinc-800 text-zinc-100 hover:text-red-500 rounded-[2rem] border border-zinc-700 transition-all shadow-lg"><Trash2 size={24}/></button>}</div>
                    {error && <p className="text-center text-[10px] text-red-500 font-black mt-2 uppercase italic tracking-tighter animate-pulse bg-red-600/5 p-3 rounded-xl border border-red-600/10">{error}</p>}
                </div>
            </div>
        </div>
      )}
    </div>
  );
};

export default App;