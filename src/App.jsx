import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell, PieChart, Pie
} from 'recharts';
import { 
  Users, CheckCircle, AlertTriangle, XCircle, Search, 
  FileText, BarChart2, MessageSquare, Calendar, TrendingUp, Database, Link, RefreshCw, Trash2, Globe, FilterX, PlayCircle, UserCheck, Settings, AlertCircle, Info, ChevronRight, ExternalLink, User, ChevronDown, CheckSquare, Square, X, Briefcase, Lock, LogIn, Activity
} from 'lucide-react';

/** * CATI CES 2026 Analytics Dashboard
 * ระบบวิเคราะห์ผลการตรวจ QC งานสัมภาษณ์ (CATI)
 */

/** * 🚩 SETUP สำคัญ: นำลิงก์จาก Google Sheets มาวางที่บรรทัดด้านล่างนี้ (ในเครื่องหมายคำพูด)
 * วิธีเอาลิงก์: File > Share > Publish to web > เลือกชีต ACQC > เลือกเป็น CSV > Copy ลิงก์มาวาง
 */
const DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSHePu18q6f93lQqVW5_JNv6UygyYRGNjT5qOq4nSrROCnGxt1pkdgiPT91rm-_lVpku-PW-LWs-ufv/pub?gid=470556665&single=true&output=csv"; 

const COLORS = {
  'ดีเยี่ยม': '#10B981',
  'ผ่านเกณฑ์': '#3B82F6',
  'ควรปรับปรุง': '#F59E0B',
  'พบข้อผิดพลาด': '#EF4444',
  'ไม่ผ่านเกณฑ์': '#7F1D1D',
};

const RESULT_ORDER = ['ดีเยี่ยม', 'ผ่านเกณฑ์', 'ควรปรับปรุง', 'พบข้อผิดพลาด', 'ไม่ผ่านเกณฑ์'];

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
  // --- Login State ---
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [inputUser, setInputUser] = useState('');
  const [inputPass, setInputPass] = useState('');
  const [loginError, setLoginError] = useState('');

  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [sheetUrl, setSheetUrl] = useState(localStorage.getItem('qc_sheet_url') || DEFAULT_SHEET_URL);
  const [searchTerm, setSearchTerm] = useState('');
  
  // Filters State
  const [filterSup, setFilterSup] = useState('All');
  
  // Multi-select state for Result
  const [selectedResults, setSelectedResults] = useState([]);
  const [isResultDropdownOpen, setIsResultDropdownOpen] = useState(false);

  // Multi-select state for Type (AC/BC)
  const [selectedTypes, setSelectedTypes] = useState([]);
  const [isTypeDropdownOpen, setIsTypeDropdownOpen] = useState(false);
  
  // Multi-select state for agents
  const [selectedAgents, setSelectedAgents] = useState([]); 
  const [isAgentDropdownOpen, setIsAgentDropdownOpen] = useState(false);

  const [selectedYear, setSelectedYear] = useState('All');
  const [selectedMonth, setSelectedMonth] = useState('All');
  const [showSync, setShowSync] = useState(!localStorage.getItem('qc_sheet_url') && !DEFAULT_SHEET_URL);
  
  const [activeCell, setActiveCell] = useState({ agent: null, resultType: null });

  // --- ส่วนที่เติมเข้าไป: เพื่อให้แสดงผลสวยงามเมื่อรันใน Local/VS Code ---
  useEffect(() => {
    // 1. โหลดฟอนต์ภาษาไทย (Sarabun)
    if (!document.getElementById('thai-font-link')) {
      const fontLink = document.createElement('link');
      fontLink.id = 'thai-font-link';
      fontLink.rel = 'stylesheet';
      fontLink.href = 'https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,100;0,200;0,300;0,400;0,500;0,600;0,700;0,800;1,100&display=swap';
      document.head.appendChild(fontLink);
    }

    // 2. โหลด Tailwind CSS
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
  // -----------------------------------------------------------

  useEffect(() => {
    if (isAuthenticated && sheetUrl && sheetUrl.includes('http')) {
      fetchFromSheet(sheetUrl);
    }
  }, [sheetUrl, isAuthenticated]); // Add isAuthenticated dependency

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
      const response = await fetch(finalUrl);
      if (!response.ok) throw new Error("ไม่สามารถเข้าถึงไฟล์ได้ (404 หรือ ลิ้งก์หมดอายุ)");
      
      const csvText = await response.text();
      if (csvText.includes('<!DOCTYPE html>')) throw new Error("ลิ้งก์ที่ส่งมาไม่ใช่ CSV กรุณา Publish to Web เป็น CSV");
      
      const allRows = parseCSV(csvText);
      if (allRows.length < 2) throw new Error("ไฟล์ไม่มีข้อมูล");

      let headerIdx = allRows.findIndex(row => row.some(cell => {
        const c = cell.toString().toLowerCase();
        return c.includes("interviewer") || c.includes("สรุปผล") || c.includes("วันที่สัมภาษณ์");
      }));
      
      if (headerIdx === -1) throw new Error("หาหัวตารางไม่เจอ กรุณาตรวจสอบชื่อคอลัมน์");
      
      const headers = allRows[headerIdx].map(h => h.trim());
      const getIdx = (name) => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
      
      const idx = {
        year: getIdx("Year"),
        month: getIdx("เดือน"),
        date: getIdx("วันที่สัมภาษณ์"),
        touchpoint: getIdx("TOUCH_POINT"),
        type: getIdx("AC / BC"),
        sup: getIdx("Supervisor"),
        agent: getIdx("Interviewer"), // Usually gets 'Interviewer ID'
        agentName: headers.findIndex(h => h.toLowerCase().includes("interviewer name") || h.toLowerCase().includes("ชื่อ-นามสกุล") || h.toLowerCase().includes("ชื่อพนักงาน")), // Try to find Name
        audio: getIdx("ไฟล์เสียง"),
        result: getIdx("สรุปผลการสัมภาษณ์"),
        comment: getIdx("Comment")
      };

      if (idx.agent === -1 || idx.result === -1) throw new Error("ไม่พบคอลัมน์ 'Interviewer' หรือ 'สรุปผลการสัมภาษณ์'");

      const parsedData = allRows.slice(headerIdx + 1)
        .filter(row => {
          const agentCode = row[idx.agent]?.toString().trim() || "";
          
          // 1. Filter invalid Agent ID (Empty, #N/A, or Header text)
          if (agentCode === "" || 
              agentCode.includes("#N/A") || 
              agentCode.toLowerCase().includes("interviewer")) {
            return false;
          }

          // 2. Filter invalid Agent Name (#N/A)
          if (idx.agentName !== -1) {
            const agentNameValue = row[idx.agentName]?.toString().trim() || "";
            if (agentNameValue.includes("#N/A")) {
              return false; // Skip this row if Name is #N/A
            }
          }

          return true;
        })
        .map((row, index) => {
          let rawResult = row[idx.result]?.toString().trim() || "N/A";
          let cleanResult = "N/A";
          if (rawResult.includes("ดีเยี่ยม")) cleanResult = "ดีเยี่ยม";
          else if (rawResult.includes("ผ่านเกณฑ์")) cleanResult = "ผ่านเกณฑ์";
          else if (rawResult.includes("ควรปรับปรุง")) cleanResult = "ควรปรับปรุง";
          else if (rawResult.includes("พบข้อผิดพลาด")) cleanResult = "พบข้อผิดพลาด";
          else if (rawResult.includes("ไม่ผ่านเกณฑ์")) cleanResult = "ไม่ผ่านเกณฑ์";

          // Agent Logic: Combine ID + Name
          const agentId = row[idx.agent]?.toString().trim() || 'Unknown';
          const agentName = (idx.agentName !== -1) ? row[idx.agentName]?.toString().trim() : '';
          
          let displayAgent = agentId;
          // Check if name exists, is valid, is not same as ID
          if (agentName && 
              agentName !== agentId && 
              !agentId.includes(agentName)
          ) {
            displayAgent = `${agentId} : ${agentName}`;
          }

          return {
            id: index,
            year: idx.year !== -1 ? row[idx.year]?.toString().trim() : 'N/A',
            month: idx.month !== -1 ? row[idx.month]?.toString().trim() : 'N/A',
            date: idx.date !== -1 ? row[idx.date]?.toString().trim() : 'N/A',
            touchpoint: idx.touchpoint !== -1 ? row[idx.touchpoint]?.toString().trim() : 'N/A',
            type: idx.type !== -1 ? row[idx.type]?.toString().trim() : 'N/A',
            supervisor: idx.sup !== -1 ? row[idx.sup]?.toString().trim() : 'N/A',
            agent: displayAgent,
            audio: idx.audio !== -1 ? row[idx.audio]?.toString().trim() : '',
            result: cleanResult,
            comment: idx.comment !== -1 ? row[idx.comment]?.toString().trim() : ''
          };
        });

      if (parsedData.length === 0) throw new Error("ไม่พบข้อมูลพนักงานในไฟล์");

      setData(parsedData);
      localStorage.setItem('qc_sheet_url', finalUrl);
      setShowSync(false);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const clearConnection = () => {
    if(window.confirm("ต้องการยกเลิกการเชื่อมต่อ? ค่าทั้งหมดจะถูกลบออกจากเบราว์เซอร์นี้")) {
      localStorage.removeItem('qc_sheet_url');
      setSheetUrl("");
      setData([]);
      setShowSync(true);
    }
  };

  const availableMonths = useMemo(() => {
    const months = data.map(d => d.month).filter(m => m && m !== 'N/A');
    return [...new Set(months)];
  }, [data]);

  const availableSups = useMemo(() => {
    const sups = data.map(d => d.supervisor).filter(s => s && s !== 'N/A');
    return [...new Set(sups)].sort();
  }, [data]);

  const availableTypes = useMemo(() => {
    const types = data.map(d => d.type).filter(t => t && t !== 'N/A' && t !== '');
    return [...new Set(types)].sort();
  }, [data]);

  const availableAgents = useMemo(() => {
    let filtered = data;
    // Filter agents based on selected Supervisor so the list is relevant
    if (filterSup !== 'All') {
      filtered = filtered.filter(d => d.supervisor === filterSup);
    }
    const agents = filtered.map(d => d.agent).filter(a => a && a !== 'Unknown');
    return [...new Set(agents)].sort();
  }, [data, filterSup]);

  const filteredData = useMemo(() => {
    return data.filter(item => {
      const matchesSearch = item.agent.toLowerCase().includes(searchTerm.toLowerCase()) || 
                            item.comment.toLowerCase().includes(searchTerm.toLowerCase());
      
      const matchesResult = selectedResults.length === 0 || selectedResults.includes(item.result);
      const matchesACBC = selectedTypes.length === 0 || selectedTypes.includes(item.type);
      const matchesSup = filterSup === 'All' || item.supervisor === filterSup;
      
      // Multi-select Agent Logic
      const matchesAgent = selectedAgents.length === 0 || selectedAgents.includes(item.agent);
      
      const matchesYear = selectedYear === 'All' || item.year === selectedYear;
      const matchesMonth = selectedMonth === 'All' || item.month === selectedMonth;
      return matchesSearch && matchesResult && matchesACBC && matchesSup && matchesAgent && matchesYear && matchesMonth;
    });
  }, [data, searchTerm, selectedResults, selectedTypes, filterSup, selectedAgents, selectedYear, selectedMonth]);

  const agentSummary = useMemo(() => {
    const summaryMap = {};
    filteredData.forEach(item => {
      if (!summaryMap[item.agent]) {
        summaryMap[item.agent] = { name: item.agent, 'ดีเยี่ยม': 0, 'ผ่านเกณฑ์': 0, 'ควรปรับปรุง': 0, 'พบข้อผิดพลาด': 0, 'ไม่ผ่านเกณฑ์': 0, total: 0 };
      }
      if (summaryMap[item.agent][item.result] !== undefined) {
        summaryMap[item.agent][item.result] += 1;
      }
      summaryMap[item.agent].total += 1;
    });
    return Object.values(summaryMap).sort((a, b) => b.total - a.total);
  }, [filteredData]);

  // Chart Data Preparation
  const chartData = useMemo(() => {
    const total = filteredData.length;
    const counts = { 'ดีเยี่ยม': 0, 'ผ่านเกณฑ์': 0, 'ควรปรับปรุง': 0, 'พบข้อผิดพลาด': 0, 'ไม่ผ่านเกณฑ์': 0 };
    filteredData.forEach(d => {
        if(counts[d.result] !== undefined) counts[d.result]++;
    });
    return Object.keys(counts).map(key => ({
        name: key,
        count: counts[key],
        percent: total > 0 ? ((counts[key] / total) * 100).toFixed(1) : 0, 
        color: COLORS[key]
    }));
  }, [filteredData]);

  const detailLogs = useMemo(() => {
    let result = filteredData;
    if (activeCell.agent && activeCell.resultType) {
      result = result.filter(d => d.agent === activeCell.agent && d.result === activeCell.resultType);
    }
    return result;
  }, [filteredData, activeCell]);

  const passRate = useMemo(() => {
    if (filteredData.length === 0) return 0;
    const passed = filteredData.filter(d => ['ดีเยี่ยม', 'ผ่านเกณฑ์'].includes(d.result)).length;
    return ((passed / filteredData.length) * 100).toFixed(1);
  }, [filteredData]);

  const handleMatrixClick = (agentName, type) => {
    if (activeCell.agent === agentName && activeCell.resultType === type) {
      setActiveCell({ agent: null, resultType: null });
    } else {
      setActiveCell({ agent: agentName, resultType: type });
      const el = document.getElementById('detail-section');
      if (el) el.scrollIntoView({ behavior: 'smooth' });
    }
  };

  const toggleAgentSelection = (agent) => {
    setSelectedAgents(prev => {
      if (prev.includes(agent)) {
        return prev.filter(a => a !== agent);
      } else {
        return [...prev, agent];
      }
    });
  };

  const toggleTypeSelection = (type) => {
    setSelectedTypes(prev => {
      if (prev.includes(type)) {
        return prev.filter(t => t !== type);
      } else {
        return [...prev, type];
      }
    });
  };

  const toggleResultSelection = (result) => {
    setSelectedResults(prev => {
      if (prev.includes(result)) {
        return prev.filter(r => r !== result);
      } else {
        return [...prev, result];
      }
    });
  };

  const CustomTooltip = ({ active, payload, label }) => {
    if (active && payload && payload.length) {
      const data = payload[0].payload;
      return (
        <div className="bg-white p-4 rounded-2xl shadow-xl border border-slate-100 font-sans">
          <p className="font-black text-slate-800 text-sm mb-1">{label}</p>
          <div className="flex items-center gap-2 text-indigo-600 font-bold text-lg">
            <span>{data.count} เคส</span>
            <span className="text-slate-300">|</span>
            <span>{data.percent}%</span>
          </div>
        </div>
      );
    }
    return null;
  };

  // --- LOGIN SCREEN ---
  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4 font-sans text-slate-900">
        <div className="bg-white p-8 md:p-12 rounded-[2.5rem] shadow-2xl w-full max-w-md border border-white">
          <div className="flex justify-center mb-8">
            <div className="p-4 bg-indigo-600 rounded-3xl shadow-lg shadow-indigo-200">
              <Database size={40} className="text-white" />
            </div>
          </div>
          
          <h2 className="text-2xl font-black text-center text-slate-800 uppercase tracking-tight mb-2">
            CATI CES 2026
          </h2>
          <p className="text-center text-slate-400 text-sm font-bold uppercase tracking-widest mb-8">
            QC Report
          </p>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-black text-slate-400 uppercase tracking-widest mb-2 pl-2">
                Username
              </label>
              <div className="relative">
                <User className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                <input 
                  type="text" 
                  value={inputUser}
                  onChange={(e) => setInputUser(e.target.value)}
                  className="w-full pl-12 pr-4 py-4 bg-slate-50 border border-slate-200 rounded-2xl font-bold text-slate-700 outline-none focus:ring-2 focus:ring-indigo-500 transition-all"
                  placeholder="Enter Username"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-black text-slate-400 uppercase tracking-widest mb-2 pl-2">
                Password
              </label>
              <div className="relative">
                <Lock className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                <input 
                  type="password" 
                  value={inputPass}
                  onChange={(e) => setInputPass(e.target.value)}
                  className="w-full pl-12 pr-4 py-4 bg-slate-50 border border-slate-200 rounded-2xl font-bold text-slate-700 outline-none focus:ring-2 focus:ring-indigo-500 transition-all"
                  placeholder="Enter Password"
                />
              </div>
            </div>

            {loginError && (
              <div className="flex items-center gap-2 text-red-500 text-xs font-bold bg-red-50 p-3 rounded-xl">
                <AlertCircle size={14} /> {loginError}
              </div>
            )}

            <button 
              type="submit" 
              className="w-full py-4 bg-indigo-600 hover:bg-indigo-700 text-white rounded-2xl font-black uppercase tracking-wider shadow-lg shadow-indigo-200 transition-all mt-4 flex items-center justify-center gap-2 group"
            >
              เข้าสู่ระบบ <LogIn size={18} className="group-hover:translate-x-1 transition-transform"/>
            </button>
          </form>

          <div className="mt-8 text-center">
            <p className="text-[10px] text-slate-300 font-bold uppercase">
              Secure Access • Authorized Personnel Only
            </p>
          </div>
        </div>
      </div>
    );
  }

  // --- MAIN DASHBOARD ---
  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans text-slate-900">
      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* Header */}
        <header className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-white p-6 rounded-[2.5rem] shadow-sm border border-slate-200">
          <div className="flex items-center gap-4">
            <div className={`p-3 rounded-[1.25rem] text-white shadow-lg transition-all duration-500 ${data.length > 0 ? 'bg-indigo-600 scale-100' : 'bg-slate-300 scale-95 opacity-50'}`}>
              <Database size={28} />
            </div>
            <div>
              <h1 className="text-2xl font-black text-slate-800 tracking-tight flex items-center gap-2 uppercase">
                CATI CES 2026 ANALYTICS
                {loading && <RefreshCw size={20} className="animate-spin text-blue-500" />}
              </h1>
              <div className="text-slate-500 text-xs font-bold flex items-center gap-1 uppercase tracking-widest">
                {data.length > 0 ? (
                  <><div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></div> เชื่อมต่อสำเร็จ: {data.length} เคส</>
                ) : (
                  <><div className="w-2 h-2 rounded-full bg-red-400"></div> ยังไม่ได้เชื่อมต่อข้อมูล</>
                )}
              </div>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={() => setShowSync(!showSync)} className="flex items-center gap-2 px-6 py-3 bg-indigo-600 text-white rounded-2xl text-sm font-black hover:bg-indigo-700 shadow-lg shadow-indigo-100 transition-all">
              <Settings size={14} /> {data.length > 0 ? 'ตั้งค่าการเชื่อมต่อ' : 'เชื่อมต่อ GOOGLE SHEET'}
            </button>
            {data.length > 0 && (
              <button onClick={clearConnection} className="p-3 bg-red-50 text-red-500 rounded-2xl hover:bg-red-100 border border-red-100 transition-colors">
                <Trash2 size={20} />
              </button>
            )}
             <button onClick={() => setIsAuthenticated(false)} className="p-3 bg-slate-100 text-slate-500 rounded-2xl hover:bg-slate-200 border border-slate-200 transition-colors" title="ออกจากระบบ">
                <User size={20} />
            </button>
          </div>
        </header>

        {/* Sync Panel / Error Display */}
        {(showSync || error) && (
          <div className={`p-8 rounded-[2.5rem] shadow-xl animate-in fade-in zoom-in-95 duration-300 border-2 ${error ? 'bg-red-50 border-red-200' : 'bg-white border-indigo-200'}`}>
            <div className="flex items-center justify-between mb-6">
                <h3 className={`text-xl font-black flex items-center gap-2 italic ${error ? 'text-red-700' : 'text-slate-800'}`}>
                    {error ? <AlertCircle className="text-red-500" /> : <Globe className="text-indigo-500" />}
                    {error ? 'เกิดข้อผิดพลาด' : 'ขั้นตอนการเชื่อมต่อข้อมูล'}
                </h3>
                {data.length > 0 && <button onClick={() => {setShowSync(false); setError(null);}} className="text-slate-400 font-bold">ปิดหน้าต่างนี้</button>}
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div className="space-y-4">
                    <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100">
                        <p className="text-sm font-bold text-slate-700 mb-2 flex items-center gap-2"><Info size={16} className="text-blue-500"/> วิธีเอาลิงก์ CSV จาก Google Sheets</p>
                        <ol className="text-xs text-slate-500 space-y-2 list-decimal list-inside leading-relaxed">
                            <li>เปิด Google Sheets ไฟล์งานของคุณ</li>
                            <li>ไปที่ <span className="text-indigo-600 font-bold">ไฟล์ (File) {">"} แชร์ (Share) {">"} เผยแพร่ไปยังเว็บ (Publish to web)</span></li>
                            <li>เลือกชีต <span className="text-indigo-600 font-bold">"ACQC"</span> และเปลี่ยนรูปแบบเป็น <span className="text-indigo-600 font-bold">"Comma-separated values (.csv)"</span></li>
                            <li>กด <span className="text-indigo-600 font-bold">เผยแพร่ (Publish)</span> แล้วคัดลอกลิงก์มาวาง</li>
                        </ol>
                    </div>
                </div>

                <div className="flex flex-col justify-center gap-3">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">CSV URL จาก Google Sheets</label>
                    <input 
                        type="text" placeholder="วางลิงก์ https://docs.google.com/spreadsheets/d/e/... ที่นี่" 
                        className={`px-5 py-4 bg-white border rounded-[1.25rem] text-sm font-medium focus:ring-2 outline-none shadow-sm ${error ? 'border-red-300 focus:ring-red-500' : 'border-slate-200 focus:ring-indigo-500'}`}
                        value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)}
                    />
                    <button 
                        onClick={() => fetchFromSheet(sheetUrl)} 
                        disabled={loading || !sheetUrl}
                        className={`px-8 py-4 text-white rounded-[1.25rem] text-sm font-black transition-all shadow-lg ${loading ? 'bg-slate-400 cursor-not-allowed' : 'bg-slate-900 hover:bg-black'}`}
                    >
                        {loading ? 'กำลังโหลดข้อมูล...' : 'ตกลงและเริ่มวิเคราะห์'}
                    </button>
                    {error && <p className="text-xs text-red-600 font-bold mt-2 flex items-center gap-1"><AlertCircle size={14}/> {error}</p>}
                </div>
            </div>
          </div>
        )}

        {data.length > 0 ? (
          <>
            {/* KPI Section */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              {[
                { label: 'งานที่ตรวจทั้งหมด', value: filteredData.length, icon: FileText, color: 'text-indigo-600', bg: 'bg-indigo-100' },
                { label: 'อัตราผ่านเกณฑ์', value: `${passRate}%`, icon: CheckCircle, color: 'text-emerald-600', bg: 'bg-emerald-100' },
                { label: 'ควรปรับปรุง', value: filteredData.filter(d=>d.result==='ควรปรับปรุง').length, icon: AlertTriangle, color: 'text-orange-600', bg: 'bg-orange-100' },
                { label: 'พบข้อผิดพลาด', value: filteredData.filter(d=>d.result==='พบข้อผิดพลาด').length, icon: XCircle, color: 'text-red-600', bg: 'bg-red-100' }
              ].map((kpi, i) => (
                <div key={i} className="bg-white p-6 rounded-[2rem] border border-slate-200 shadow-sm transition-all hover:shadow-md">
                  <div className={`w-8 h-8 ${kpi.bg} ${kpi.color} rounded-xl flex items-center justify-center mb-3`}><kpi.icon size={16} /></div>
                  <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{kpi.label}</p>
                  <h2 className={`text-3xl font-black ${kpi.color}`}>{kpi.value}</h2>
                </div>
              ))}
            </div>

            {/* Visual Analytics Chart */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div className="lg:col-span-3 bg-white p-8 rounded-[2.5rem] shadow-sm border border-slate-200">
                    <div className="flex flex-col md:flex-row md:items-center justify-between mb-6 gap-4">
                        <div className="flex items-center gap-3">
                            <div className="p-3 bg-indigo-50 rounded-xl text-indigo-600"><BarChart2 size={20}/></div>
                            <h3 className="font-black text-slate-800 uppercase italic tracking-tight text-lg">Performance Distribution (with %)</h3>
                        </div>
                        <div className="flex items-center gap-2">
                            {chartData.map((item) => (
                                <div key={item.name} className="flex flex-col items-center bg-slate-50 px-3 py-2 rounded-xl min-w-[80px]">
                                    <span className="text-[10px] font-bold text-slate-400">{item.name}</span>
                                    <span className="text-sm font-black" style={{color: item.color}}>{item.percent}%</span>
                                </div>
                            ))}
                        </div>
                    </div>
                    <div className="h-64 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={chartData} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 12, fontWeight: 700, fill: '#64748b', fontFamily: 'Sarabun'}} dy={10} />
                                <YAxis axisLine={false} tickLine={false} tick={{fontSize: 12, fill: '#94a3b8', fontFamily: 'Sarabun'}} />
                                <Tooltip cursor={{fill: '#f8fafc'}} content={<CustomTooltip />} />
                                <Bar dataKey="count" radius={[8, 8, 8, 8]} barSize={50}>
                                    {chartData.map((entry, index) => (
                                        <Cell key={`cell-${index}`} fill={entry.color} />
                                    ))}
                                </Bar>
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                </div>
            </div>

            {/* Matrix Table */}
            <div className="bg-white rounded-[3rem] shadow-sm border border-slate-200 overflow-hidden">
              <div className="p-8 border-b border-slate-100 bg-slate-50/50 flex flex-col md:flex-row items-center justify-between gap-6">
                <div>
                  <h3 className="font-black text-slate-800 flex items-center gap-2 italic text-lg uppercase tracking-tight">
                    <TrendingUp size={24} className="text-indigo-500" />
                    สรุปคุณภาพพนักงาน x ผลสัมภาษณ์
                  </h3>
                  <p className="text-[10px] text-slate-400 font-bold uppercase mt-1 italic flex items-center gap-1"><ChevronRight size={12}/> คลิกที่ตัวเลข เพื่อดูรายละเอียดคอมเมนต์ด้านล่าง</p>
                </div>
                <div className="flex flex-wrap gap-2">

                  {/* Multi-Select Result Filter */}
                  <div className="relative">
                    <button 
                      onClick={() => setIsResultDropdownOpen(!isResultDropdownOpen)}
                      className={`flex items-center gap-2 border px-3 py-1.5 rounded-xl text-[10px] font-black outline-none shadow-sm transition-all ${selectedResults.length > 0 ? 'bg-indigo-50 border-indigo-200 text-indigo-700' : 'bg-white border-slate-200'}`}
                    >
                      <Activity size={14} className={selectedResults.length > 0 ? "text-indigo-500" : "text-indigo-500"} />
                      <span className="truncate max-w-[120px]">
                        {selectedResults.length === 0 ? 'ผลการสัมภาษณ์' : `เลือกแล้ว ${selectedResults.length}`}
                      </span>
                      <ChevronDown size={12} className={`transition-transform ${isResultDropdownOpen ? 'rotate-180' : ''}`} />
                    </button>

                    {isResultDropdownOpen && (
                      <>
                        <div className="fixed inset-0 z-10" onClick={() => setIsResultDropdownOpen(false)}></div>
                        <div className="absolute top-full mt-2 left-0 min-w-[200px] bg-white rounded-2xl shadow-xl z-20 border border-slate-100 overflow-hidden animate-in fade-in zoom-in-95 duration-200">
                          <div className="p-2 border-b border-slate-50 flex items-center justify-between bg-slate-50">
                            <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-2">เลือกผลงาน</span>
                             <div className="flex gap-1">
                                {selectedResults.length > 0 && (
                                    <button onClick={() => setSelectedResults([])} className="text-[10px] text-red-500 font-bold px-2 py-1 hover:bg-red-50 rounded-lg">Clear</button>
                                )}
                                <button onClick={() => setIsResultDropdownOpen(false)} className="text-slate-400 hover:text-slate-600 p-1"><X size={14}/></button>
                            </div>
                          </div>
                          <div className="max-h-60 overflow-y-auto p-1 custom-scrollbar">
                            {RESULT_ORDER.map(res => {
                              const isSelected = selectedResults.includes(res);
                              return (
                                <div 
                                  key={res} 
                                  onClick={() => toggleResultSelection(res)}
                                  className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-xs font-bold transition-colors ${isSelected ? 'bg-indigo-50 text-indigo-700' : 'hover:bg-slate-50 text-slate-600'}`}
                                >
                                  {isSelected ? <CheckSquare size={16} className="text-indigo-500" /> : <Square size={16} className="text-slate-300" />}
                                  <span style={{ color: COLORS[res] }}>{res}</span>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </>
                    )}
                  </div>
                  
                  {/* Multi-Select Type Filter (AC/BC) */}
                  <div className="relative">
                    <button 
                      onClick={() => setIsTypeDropdownOpen(!isTypeDropdownOpen)}
                      className={`flex items-center gap-2 border px-3 py-1.5 rounded-xl text-[10px] font-black outline-none shadow-sm transition-all ${selectedTypes.length > 0 ? 'bg-orange-50 border-orange-200 text-orange-700' : 'bg-white border-slate-200'}`}
                    >
                      <Briefcase size={14} className={selectedTypes.length > 0 ? "text-orange-500" : "text-orange-500"} />
                      <span className="truncate max-w-[120px]">
                        {selectedTypes.length === 0 ? 'ประเภทงาน' : `เลือกแล้ว ${selectedTypes.length}`}
                      </span>
                      <ChevronDown size={12} className={`transition-transform ${isTypeDropdownOpen ? 'rotate-180' : ''}`} />
                    </button>

                    {isTypeDropdownOpen && (
                      <>
                        <div className="fixed inset-0 z-10" onClick={() => setIsTypeDropdownOpen(false)}></div>
                        <div className="absolute top-full mt-2 left-0 min-w-[200px] bg-white rounded-2xl shadow-xl z-20 border border-slate-100 overflow-hidden animate-in fade-in zoom-in-95 duration-200">
                          <div className="p-2 border-b border-slate-50 flex items-center justify-between bg-slate-50">
                            <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-2">เลือกประเภทงาน</span>
                             <div className="flex gap-1">
                                {selectedTypes.length > 0 && (
                                    <button onClick={() => setSelectedTypes([])} className="text-[10px] text-red-500 font-bold px-2 py-1 hover:bg-red-50 rounded-lg">Clear</button>
                                )}
                                <button onClick={() => setIsTypeDropdownOpen(false)} className="text-slate-400 hover:text-slate-600 p-1"><X size={14}/></button>
                            </div>
                          </div>
                          <div className="max-h-60 overflow-y-auto p-1 custom-scrollbar">
                            {availableTypes.length > 0 ? availableTypes.map(type => {
                              const isSelected = selectedTypes.includes(type);
                              return (
                                <div 
                                  key={type} 
                                  onClick={() => toggleTypeSelection(type)}
                                  className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-xs font-bold transition-colors ${isSelected ? 'bg-orange-50 text-orange-700' : 'hover:bg-slate-50 text-slate-600'}`}
                                >
                                  {isSelected ? <CheckSquare size={16} className="text-orange-500" /> : <Square size={16} className="text-slate-300" />}
                                  {type}
                                </div>
                              );
                            }) : (
                                <div className="p-4 text-center text-xs text-slate-400 font-bold">ไม่พบข้อมูล</div>
                            )}
                          </div>
                        </div>
                      </>
                    )}
                  </div>

                  {/* Supervisor Filter */}
                  <div className="flex items-center gap-2 bg-white border border-slate-200 px-3 py-1.5 rounded-xl">
                    <UserCheck size={14} className="text-indigo-500" />
                    <select 
                      className="bg-transparent text-[10px] font-black outline-none max-w-[120px]" 
                      value={filterSup} 
                      onChange={(e) => {
                        setFilterSup(e.target.value);
                        setSelectedAgents([]); // Reset agent selection when Sup changes
                      }}
                    >
                      <option value="All">ทุก Supervisor</option>
                      {availableSups.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>

                  {/* Multi-Select Agent Filter */}
                  <div className="relative">
                    <button 
                      onClick={() => setIsAgentDropdownOpen(!isAgentDropdownOpen)}
                      className={`flex items-center gap-2 border px-3 py-1.5 rounded-xl text-[10px] font-black outline-none shadow-sm transition-all ${selectedAgents.length > 0 ? 'bg-emerald-50 border-emerald-200 text-emerald-700' : 'bg-white border-slate-200'}`}
                    >
                      <User size={14} className={selectedAgents.length > 0 ? "text-emerald-500" : "text-emerald-500"} />
                      <span className="truncate max-w-[120px]">
                        {selectedAgents.length === 0 ? 'พนักงานทุกคน' : `เลือกแล้ว ${selectedAgents.length} คน`}
                      </span>
                      <ChevronDown size={12} className={`transition-transform ${isAgentDropdownOpen ? 'rotate-180' : ''}`} />
                    </button>

                    {isAgentDropdownOpen && (
                      <>
                        <div className="fixed inset-0 z-10" onClick={() => setIsAgentDropdownOpen(false)}></div>
                        <div className="absolute top-full mt-2 left-0 min-w-[240px] bg-white rounded-2xl shadow-xl z-20 border border-slate-100 overflow-hidden animate-in fade-in zoom-in-95 duration-200">
                          <div className="p-2 border-b border-slate-50 flex items-center justify-between bg-slate-50">
                            <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-2">เลือกพนักงาน</span>
                            <div className="flex gap-1">
                                {selectedAgents.length > 0 && (
                                    <button onClick={() => setSelectedAgents([])} className="text-[10px] text-red-500 font-bold px-2 py-1 hover:bg-red-50 rounded-lg">Clear</button>
                                )}
                                <button onClick={() => setIsAgentDropdownOpen(false)} className="text-slate-400 hover:text-slate-600 p-1"><X size={14}/></button>
                            </div>
                          </div>
                          <div className="max-h-60 overflow-y-auto p-1 custom-scrollbar">
                            {availableAgents.length > 0 ? availableAgents.map(agent => {
                              const isSelected = selectedAgents.includes(agent);
                              return (
                                <div 
                                  key={agent} 
                                  onClick={() => toggleAgentSelection(agent)}
                                  className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-xs font-bold transition-colors ${isSelected ? 'bg-indigo-50 text-indigo-700' : 'hover:bg-slate-50 text-slate-600'}`}
                                >
                                  {isSelected ? <CheckSquare size={16} className="text-indigo-600" /> : <Square size={16} className="text-slate-300" />}
                                  {agent}
                                </div>
                              );
                            }) : (
                              <div className="p-4 text-center text-xs text-slate-400 font-bold">ไม่พบพนักงาน</div>
                            )}
                          </div>
                        </div>
                      </>
                    )}
                  </div>

                  <select className="bg-white border px-3 py-1.5 rounded-xl text-[10px] font-black outline-none shadow-sm" value={selectedMonth} onChange={(e)=>setSelectedMonth(e.target.value)}>
                    <option value="All">ทุกเดือน</option>
                    {availableMonths.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                </div>
              </div>
              <div className="overflow-x-auto max-h-[500px]">
                <table className="w-full text-left text-sm border-separate border-spacing-0">
                  <thead className="sticky top-0 bg-white z-20 font-black text-slate-400 text-[10px] uppercase tracking-widest border-b border-slate-100 shadow-sm">
                    <tr>
                      <th rowSpan="2" className="px-8 py-6 border-b-2 border-slate-100 border-r border-slate-50 bg-white">พนักงานสัมภาษณ์</th>
                      <th colSpan={RESULT_ORDER.length} className="px-4 py-4 text-center border-b-2 border-slate-100 bg-slate-50 text-indigo-600 text-[11px] font-black italic uppercase">สรุปผลการสัมภาษณ์</th>
                      <th rowSpan="2" className="px-8 py-6 text-center bg-slate-50 text-slate-800 border-b-2 border-slate-100 border-l border-slate-100">รวม</th>
                    </tr>
                    <tr className="bg-white shadow-sm">
                      {RESULT_ORDER.map(type => (
                        <th key={type} className="px-4 py-3 text-center border-b border-slate-100 border-r border-slate-50 text-slate-500">
                          {type}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50 font-medium">
                    {agentSummary.map((agent, i) => (
                      <tr key={i} className="hover:bg-slate-50 transition-colors group">
                        <td className="px-8 py-4 font-black text-slate-700 border-r border-slate-50">{agent.name}</td>
                        {RESULT_ORDER.map(type => {
                          const isActive = activeCell.agent === agent.name && activeCell.resultType === type;
                          const val = agent[type];
                          return (
                            <td 
                              key={type} 
                              className={`px-4 py-4 text-center border-r border-slate-50 transition-all ${val > 0 ? 'cursor-pointer hover:bg-slate-100' : ''} ${isActive ? 'bg-indigo-50 ring-2 ring-inset ring-indigo-500 z-10' : ''}`}
                              onClick={() => val > 0 && handleMatrixClick(agent.name, type)}
                            >
                              <span className={`text-sm font-black ${val > 0 ? '' : 'text-slate-100'}`} style={{ color: val > 0 ? COLORS[type] : undefined }}>
                                {val || '-'}
                              </span>
                            </td>
                          );
                        })}
                        <td className="px-8 py-4 text-center bg-slate-50/50 font-black text-slate-900 border-l border-slate-100 group-hover:bg-indigo-50">{agent.total}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Detailed Case Log */}
            <div id="detail-section" className="bg-white rounded-[3rem] shadow-sm border border-slate-200 overflow-hidden scroll-mt-6">
              <div className="p-8 border-b border-slate-100 flex flex-col md:flex-row items-center justify-between gap-6">
                <div className="space-y-1">
                  <h3 className="font-black text-slate-800 uppercase tracking-widest text-xs flex items-center gap-2">
                    <MessageSquare size={16} className="text-indigo-500" /> ข้อมูลรายเคสแบบละเอียด
                  </h3>
                  {activeCell.agent && (
                    <div className="flex items-center gap-2 animate-bounce">
                      <span className="text-[10px] font-black px-2 py-1 bg-indigo-100 text-indigo-700 rounded-lg flex items-center gap-1 shadow-sm border border-indigo-200">
                        กำลังกรอง: {activeCell.agent} ({activeCell.resultType})
                        <button onClick={() => setActiveCell({ agent: null, resultType: null })} className="ml-1 hover:text-red-500 transition-colors">
                          <FilterX size={12} />
                        </button>
                      </span>
                    </div>
                  )}
                </div>
                <div className="relative w-full md:w-80">
                  <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                  <input 
                    type="text" placeholder="ค้นหาพนักงาน หรือ คอมเมนต์..." 
                    className="w-full pl-12 pr-6 py-3 bg-slate-50 border border-slate-200 rounded-[1.25rem] text-sm font-medium outline-none focus:ring-2 focus:ring-indigo-500 transition-all shadow-inner"
                    value={searchTerm} onChange={(e)=>setSearchTerm(e.target.value)}
                  />
                </div>
              </div>
              <div className="overflow-auto max-h-[600px]">
                <table className="w-full text-left text-xs font-medium">
                  <thead className="sticky top-0 bg-white shadow-sm z-10 border-b border-slate-100 font-black text-slate-400 uppercase tracking-widest">
                    <tr>
                      <th className="px-8 py-5">วันที่ / ประเภท</th>
                      <th className="px-8 py-5">พนักงานสัมภาษณ์</th>
                      <th className="px-4 py-5 text-center">ผลการสัมภาษณ์</th>
                      <th className="px-8 py-5">QC Comment & Audio</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {detailLogs.length > 0 ? detailLogs.slice(0, 150).map((item, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/80 transition-all group">
                        <td className="px-8 py-5">
                          <div className="font-black text-slate-800">{item.date}</div>
                          <div className="text-[9px] text-slate-400 uppercase font-bold mt-1 tracking-tighter">Sup: {item.supervisor}</div>
                          <div className={`text-[9px] font-black inline-block px-2 py-0.5 rounded-lg mt-1 shadow-sm ${item.type === 'AC' ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-white'}`}>{item.type}</div>
                        </td>
                        <td className="px-8 py-5">
                          <div className="font-black text-slate-800 text-sm group-hover:text-indigo-600 transition-colors">{item.agent}</div>
                          <div className="text-[10px] text-slate-400 font-bold uppercase tracking-tighter italic opacity-60">{item.touchpoint}</div>
                        </td>
                        <td className="px-4 py-5 text-center">
                          <span className="inline-flex items-center gap-1.5 px-4 py-1.5 rounded-full text-[9px] font-black shadow-sm" style={{ backgroundColor: `${COLORS[item.result]}15`, color: COLORS[item.result], border: `1px solid ${COLORS[item.result]}25` }}>
                            <div className="w-1.5 h-1.5 rounded-full" style={{ backgroundColor: COLORS[item.result] }}></div>
                            {item.result}
                          </span>
                        </td>
                        <td className="px-8 py-5">
                          <div className="flex flex-col gap-2">
                            <p className="text-slate-500 font-semibold italic line-clamp-2 group-hover:text-slate-900 group-hover:line-clamp-none transition-all">
                              {item.comment ? `"${item.comment}"` : '-'}
                            </p>
                            {item.audio && item.audio.includes('http') && (
                              <a 
                                href={item.audio} 
                                target="_blank" 
                                rel="noopener noreferrer"
                                className="flex items-center gap-1.5 text-indigo-600 hover:text-indigo-800 font-black text-[10px] uppercase transition-colors"
                              >
                                <PlayCircle size={14} /> Listen Recording
                              </a>
                            )}
                          </div>
                        </td>
                      </tr>
                    )) : (
                      <tr>
                        <td colSpan={4} className="px-8 py-20 text-center text-slate-400 font-bold italic opacity-50">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        ) : (
          !loading && (
            <div className="bg-white rounded-[4rem] border-4 border-dashed border-slate-200 py-32 text-center shadow-inner">
                <div className="bg-slate-100 w-24 h-24 rounded-[2.5rem] flex items-center justify-center mx-auto mb-8 shadow-sm">
                <Database size={44} className="text-slate-300" />
                </div>
                <h2 className="text-3xl font-black text-slate-300 tracking-tighter uppercase italic">Waiting for Connection</h2>
                <p className="text-slate-400 text-sm mt-3 max-w-sm mx-auto font-bold uppercase tracking-widest leading-relaxed px-4">
                กรุณากดปุ่ม "ตั้งค่าการเชื่อมต่อ" ด้านบน และวางลิงก์ CSV จาก Google Sheets เพื่อเริ่มการทำงาน
                </p>
            </div>
          )
        )}

      </div>
    </div>
  );
};

export default App;