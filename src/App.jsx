import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell, PieChart, Pie
} from 'recharts';
import { 
  Users, CheckCircle, AlertTriangle, XCircle, Search, 
  FileText, BarChart2, MessageSquare, Calendar, TrendingUp, Database, Link, RefreshCw, Trash2, Globe, FilterX, PlayCircle, UserCheck, Settings, AlertCircle, Info, ChevronRight, ExternalLink, User, ChevronDown, CheckSquare, Square, X, Briefcase, Lock, LogIn, Activity, Filter, Check, Clock, ListChecks, Award, Save, Edit2, Hash, Star, Zap
} from 'lucide-react';

/** * CATI CES 2026 Analytics Dashboard - MASTER VERSION (JSON API METHOD)
 * ระบบวิเคราะห์ผลการตรวจ QC พร้อมระบบแก้ไขข้อมูล (High Performance)
 * - อัปเดต: เปลี่ยนระบบดึงข้อมูลจาก CSV เป็น JSON ผ่าน Apps Script (Method 1)
 * - อัปเดต: เพิ่มความยืดหยุ่นในการหา Header Row (แก้ปัญหาหาคอลัมน์ไม่เจอ)
 * - อัปเดต: รวมช่อง Link เหลือช่องเดียวเพื่อความ Seamless
 */

// ค่า Default Apps Script URL ที่ User ต้องเอามาใส่ (ถ้ามีของเดิมให้ใส่ของเดิม)
const DEFAULT_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwpee9kgf_wdEgrHIsTD3ECiaSwCevocNL79vWqCi5s6XrvdI1sQc8ekhu_D4j77hz0/exec";

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
const SCORE_LABELS = { '5': '5.ดี', '4': '4.ค่อนข้างดี', '3': '3.ปานกลาง', '2': '2.ไม่ค่อยดี', '1': '1.ไม่ดีเลย', '-': '-' };

const formatResultDisplay = (text) => (text ? text.split('(')[0].trim() : '-');

const COLORS = {
  'ดีเยี่ยม': '#6366f1',     // Indigo-500
  'ผ่านเกณฑ์': '#10B981',   // Emerald-500
  'ควรปรับปรุง': '#F59E0B', // Amber-500
  'พบข้อผิดพลาด': '#f43f5e', // Rose-500
  'ไม่ผ่านเกณฑ์': '#be123c', // Rose-700
};

const getResultColor = (fullText) => {
  if (fullText.startsWith('ดีเยี่ยม')) return COLORS['ดีเยี่ยม'];
  if (fullText.startsWith('ผ่านเกณฑ์')) return COLORS['ผ่านเกณฑ์'];
  if (fullText.startsWith('ควรปรับปรุง')) return COLORS['ควรปรับปรุง'];
  if (fullText.startsWith('พบข้อผิดพลาด')) return COLORS['พบข้อผิดพลาด'];
  if (fullText.startsWith('ไม่ผ่านเกณฑ์')) return COLORS['ไม่ผ่านเกณฑ์'];
  return '#94a3b8'; // Slate-400
};

const IntageLogo = ({ className = "h-8" }) => (
  <div className={`flex items-center gap-2 ${className}`}>
    <div className="w-4 h-4 rounded-full bg-indigo-500 animate-pulse"></div>
    <span className="font-black tracking-[0.15em] text-white italic text-lg">INTAGE</span>
  </div>
);

const App = () => {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [inputUser, setInputUser] = useState('');
  const [inputPass, setInputPass] = useState('');
  const [loginError, setLoginError] = useState('');
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState(null);
  
  const getStorage = (key, fallback) => { try { return localStorage.getItem(key) || fallback; } catch(e) { return fallback; } };

  const [appsScriptUrl, setAppsScriptUrl] = useState(getStorage('apps_script_url', DEFAULT_APPS_SCRIPT_URL));
  const [searchTerm, setSearchTerm] = useState('');
  const [isFilterSidebarOpen, setIsFilterSidebarOpen] = useState(false);
  const [lastUpdated, setLastUpdated] = useState(null);
  
  const [selectedMonths, setSelectedMonths] = useState([]);
  const [selectedSups, setSelectedSups] = useState([]);
  const [selectedResults, setSelectedResults] = useState([]);
  const [selectedAgents, setSelectedAgents] = useState([]); 
  const [selectedTypes, setSelectedTypes] = useState([]); 
  
  const [showSync, setShowSync] = useState(getStorage('apps_script_url', '') === '' && DEFAULT_APPS_SCRIPT_URL === '');
  
  const [activeCell, setActiveCell] = useState({ agent: null, resultType: null });
  const [expandedCaseId, setExpandedCaseId] = useState(null);
  const [editingCase, setEditingCase] = useState(null); 

  useEffect(() => {
    if (!document.getElementById('thai-font-link')) {
      const link = document.createElement('link');
      link.id = 'thai-font-link'; link.rel = 'stylesheet';
      link.href = 'https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;700;800&display=swap';
      document.head.appendChild(link);
    }
    if (!document.getElementById('tailwind-cdn')) {
      const script = document.createElement('script');
      script.id = 'tailwind-cdn'; script.src = 'https://cdn.tailwindcss.com';
      document.head.appendChild(script);
    }
  }, []);

  useEffect(() => {
    let intervalId;
    if (isAuthenticated && appsScriptUrl && appsScriptUrl.includes('http')) {
      fetchFromAppsScript(appsScriptUrl);
      intervalId = setInterval(() => fetchFromAppsScript(appsScriptUrl), 60000); // 1 minute (Faster refresh)
    }
    return () => clearInterval(intervalId);
  }, [appsScriptUrl, isAuthenticated]);

  const fetchFromAppsScript = async (urlToFetch) => {
    setLoading(true); setError(null);
    try {
      // 1. Fetch JSON from Apps Script (Method 1)
      const response = await fetch(urlToFetch);
      if (!response.ok) throw new Error("ไม่สามารถเชื่อมต่อกับ Apps Script ได้");
      
      const allRows = await response.json(); // ได้ข้อมูลเป็น Array of Arrays ทันที
      
      if (!Array.isArray(allRows) || allRows.length === 0) throw new Error("ไม่พบข้อมูล หรือ Format ไม่ถูกต้อง");

      // 2. Logic ในการหา Header และ Map ข้อมูล (UPDATED for Robustness)
      let headerIdx = allRows.findIndex(row => 
        Array.isArray(row) && row.some(cell => 
          (cell && cell.toString().toLowerCase().includes("interviewer")) || 
          (cell && cell.toString().includes("สรุปผล")) ||
          (cell && cell.toString().toLowerCase().includes("name"))
        )
      );

      // Fallback Strategy 1: Look for Date/Time
      if (headerIdx === -1) {
         headerIdx = allRows.findIndex(row => 
            Array.isArray(row) && row.some(cell => 
              (cell && cell.toString().toLowerCase().includes("date")) || 
              (cell && cell.toString().includes("วันที่"))
            )
         );
      }

      // Fallback Strategy 2: Default to row 0 if headers look like strings
      if (headerIdx === -1 && allRows.length > 0) {
          console.warn("Could not find specific header keywords. Defaulting to row 0.");
          headerIdx = 0;
      }

      if (headerIdx === -1) throw new Error("ไม่พบคอลัมน์ข้อมูลที่กำหนด (กรุณาตรวจสอบ Header Row)");

      const headers = allRows[headerIdx].map(h => h ? h.toString().trim() : "");
      
      // Helper function to safely find column index
      const getIdx = (keywords) => {
        if (!Array.isArray(keywords)) keywords = [keywords];
        return headers.findIndex(h => {
            if (!h) return false;
            const lowerH = h.toLowerCase();
            return keywords.some(k => lowerH.includes(k.toLowerCase()));
        });
      };

      // Column mapping with multiple keywords
      const idx = {
        month: getIdx(["เดือน", "month"]), 
        date: getIdx(["วันที่สัมภาษณ์", "date", "timestamp"]), 
        touchpoint: getIdx(["TOUCH_POINT", "touchpoint"]), 
        type: getIdx(["AC / BC", "type", "ac/bc"]), 
        sup: getIdx(["Supervisor", "sup"]), 
        agent: getIdx(["Interviewer", "name", "ชื่อ"]),
        questionnaireNo: getIdx(["questionnaire", "no.", "เลขชุด", "id"]), 
        result: getIdx(["สรุปผลการสัมภาษณ์", "result", "grade"]), 
        comment: getIdx(["Comment", "ความคิดเห็น", "ข้อเสนอแนะ"]),
        audio: getIdx(["ไฟล์เสียง", "audio", "record"]) 
      };

      // Check essential columns
      if (idx.agent === -1) {
          // If agent column not found by name, try generic indices or fallback
          // Assuming column 1 or 2 often has names if headers are messed up
          // For now, let's proceed but warn
          console.warn("Agent column ambiguous.");
      }
      
      // Determine evaluations columns (Starting from 'Introduction' or similar)
      let evalStartIdx = headers.findIndex(h => h && (h.includes("Introduction") || h.includes("น้ำเสียง") || h.includes("1.")));
      if (evalStartIdx === -1) evalStartIdx = 15; // Default fallback

      const parsedData = allRows.slice(headerIdx + 1)
        .map((row, i) => ({ row, actualRowNumber: i + headerIdx + 2 })) // +1 header, +1 for 1-based index
        .filter(({ row }) => {
          // Check bounds
          if (!row || !Array.isArray(row)) return false;
          // Use the detected agent index, or fallback to a likely column (e.g. 10 or 2)
          const agentColIndex = idx.agent !== -1 ? idx.agent : 10; 
          if (row.length <= agentColIndex) return false;
          
          const agentCode = row[agentColIndex]?.toString().trim() || "";
          return agentCode !== "" && !agentCode.includes("#N/A");
        })
        .map(({ row, actualRowNumber }, index) => {
          let rawResult = (idx.result !== -1 && row[idx.result]) ? row[idx.result].toString().trim() : "N/A";
          let cleanResult = rawResult;
          const matchedResult = RESULT_ORDER.find(opt => rawResult.includes(opt.split(':')[0].trim()));
          if (matchedResult) cleanResult = matchedResult;
          
          const agentColIndex = idx.agent !== -1 ? idx.agent : 10;
          const agentId = row[agentColIndex]?.trim() || 'Unknown';
          
          // Try to find full name if agentId is just a code, assume name is next column or specific logic
          // Simplified: Just use what we found in agent column
          const displayAgent = agentId;
          
          const rawType = (idx.type !== -1 && row[idx.type]) ? row[idx.type].toString().trim() : "";
          const cleanType = (rawType === "" || rawType === "N/A") ? "ยังไม่ได้ตรวจ" : rawType;

          // Evaluations
          const evaluations = [];
          if (evalStartIdx > -1) {
              for (let i = 0; i < 13; i++) { // Approx 13 criteria
                  if (headers[evalStartIdx + i]) {
                      evaluations.push({
                          label: headers[evalStartIdx + i],
                          value: row[evalStartIdx + i] || '-'
                      });
                  }
              }
          }

          return {
            id: index, rowIndex: actualRowNumber, 
            month: (idx.month !== -1 && row[idx.month]) ? row[idx.month] : 'N/A', 
            date: (idx.date !== -1 && row[idx.date]) ? row[idx.date] : 'N/A', 
            agent: displayAgent, 
            questionnaireNo: (idx.questionnaireNo !== -1 && row[idx.questionnaireNo]) ? row[idx.questionnaireNo] : '-', 
            result: cleanResult, 
            comment: (idx.comment !== -1 && row[idx.comment]) ? row[idx.comment] : '', 
            audio: (idx.audio !== -1 && row[idx.audio]) ? row[idx.audio] : '', 
            touchpoint: (idx.touchpoint !== -1 && row[idx.touchpoint]) ? row[idx.touchpoint] : 'N/A', 
            supervisor: (idx.sup !== -1 && row[idx.sup]) ? row[idx.sup] : 'N/A',
            type: cleanType,
            evaluations: evaluations
          };
        });

      setData(parsedData); 
      setLastUpdated(new Date().toLocaleTimeString('th-TH'));
      localStorage.setItem('apps_script_url', urlToFetch); // Save success URL
      setShowSync(false);

    } catch (err) { 
        setError(err.message); 
        console.error("Fetch Error:", err);
    } finally { 
        setLoading(false); 
    }
  };

  const handleUpdateCase = async () => {
    if (!appsScriptUrl) return alert("กรุณาตั้งค่า Apps Script URL");
    
    if (editingCase.type === "ยังไม่ได้ตรวจ") {
      alert("กรุณาเลือกประเภทงาน (AC หรือ BC) ก่อนบันทึก");
      return;
    }

    setIsSaving(true);
    const backupData = [...data];

    try {
      const updateData = {
        rowIndex: editingCase.rowIndex, 
        result: editingCase.result,
        type: editingCase.type,
        evaluations: editingCase.evaluations.map(e => e.value), 
        comment: editingCase.comment
      };
      
      // Send POST to the SAME URL (Apps Script handles both GET/POST)
      await fetch(appsScriptUrl, { 
        method: 'POST', 
        mode: 'no-cors', 
        cache: 'no-cache', 
        headers: { 'Content-Type': 'application/json' }, 
        body: JSON.stringify(updateData) 
      });
      
      // INSTANT UI UPDATE (Optimistic Update)
      setData(prevData => prevData.map(item => {
        if (item.id === editingCase.id) {
          return { ...editingCase };
        }
        return item;
      }));

      setTimeout(() => { 
        setIsSaving(false); 
        setEditingCase(null); 
        // Fetch new data via JSON immediately (Real-time check)
        fetchFromAppsScript(appsScriptUrl); 
      }, 1500);

    } catch (err) { 
      alert("เกิดข้อผิดพลาดในการบันทึก: " + err.message); 
      setData(backupData);
      setIsSaving(false); 
    }
  };

  // --- Analytical Calculations ---
  const availableMonths = useMemo(() => [...new Set(data.map(d => d.month).filter(m => m !== 'N/A'))].sort(), [data]);
  const availableSups = useMemo(() => [...new Set(data.map(d => d.supervisor).filter(s => s !== 'N/A'))].sort(), [data]);
  const availableTypes = useMemo(() => [...new Set(data.map(d => d.type).filter(t => t !== 'N/A' && t !== ''))].sort(), [data]);
  
  const availableAgents = useMemo(() => {
    let filtered = data;
    if (selectedSups.length > 0) filtered = filtered.filter(d => selectedSups.includes(d.supervisor));
    if (selectedMonths.length > 0) filtered = filtered.filter(d => selectedMonths.includes(d.month));
    if (selectedTypes.length > 0) filtered = filtered.filter(d => selectedTypes.includes(d.type));
    return [...new Set(filtered.map(d => d.agent).filter(a => a !== 'Unknown'))].sort();
  }, [data, selectedSups, selectedMonths, selectedTypes]);

  const filteredData = useMemo(() => {
    return data.filter(item => {
      const matchesSearch = item.agent.toLowerCase().includes(searchTerm.toLowerCase()) || item.comment.toLowerCase().includes(searchTerm.toLowerCase()) || item.questionnaireNo.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesResult = selectedResults.length === 0 || selectedResults.includes(item.result);
      const matchesSup = selectedSups.length === 0 || selectedSups.includes(item.supervisor);
      const matchesAgent = selectedAgents.length === 0 || selectedAgents.includes(item.agent);
      const matchesMonth = selectedMonths.length === 0 || selectedMonths.includes(item.month);
      const matchesType = selectedTypes.length === 0 || selectedTypes.includes(item.type);
      return matchesSearch && matchesResult && matchesSup && matchesAgent && matchesMonth && matchesType;
    });
  }, [data, searchTerm, selectedResults, selectedSups, selectedAgents, selectedMonths, selectedTypes]);

  const totalWorkByMonthOnly = useMemo(() => {
    if (selectedMonths.length === 0) return data.length;
    return data.filter(item => selectedMonths.includes(item.month)).length;
  }, [data, selectedMonths]);

  const agentSummary = useMemo(() => {
    const summaryMap = {};
    filteredData.forEach(item => {
      if (!summaryMap[item.agent]) { summaryMap[item.agent] = { name: item.agent, total: 0 }; RESULT_ORDER.forEach(r => summaryMap[item.agent][r] = 0); }
      if (summaryMap[item.agent][item.result] !== undefined) summaryMap[item.agent][item.result] += 1;
      summaryMap[item.agent].total += 1;
    });
    return Object.values(summaryMap).sort((a, b) => b.total - a.total);
  }, [filteredData]);

  const chartData = useMemo(() => {
    const total = filteredData.length;
    return RESULT_ORDER.map(key => ({ 
      name: formatResultDisplay(key), full: key, count: filteredData.filter(d => d.result === key).length, 
      percent: total > 0 ? ((filteredData.filter(d => d.result === key).length / total) * 100).toFixed(1) : 0, 
      color: getResultColor(key) 
    }));
  }, [filteredData]);

  // --- New Monthly Performance Chart Data (Updated to %) ---
  const monthlyPerformanceData = useMemo(() => {
    return availableMonths.map(month => {
        const monthData = data.filter(d => d.month === month);
        const total = monthData.length;
        // Logic for "Audited": type is not 'ยังไม่ได้ตรวจ', 'N/A', or ''
        const audited = monthData.filter(d => d.type !== 'ยังไม่ได้ตรวจ' && d.type !== 'N/A' && d.type !== '').length;
        const percent = total > 0 ? parseFloat(((audited / total) * 100).toFixed(1)) : 0;
        return { name: month, audited, total, percent };
    });
  }, [data, availableMonths]);

  const detailLogs = useMemo(() => (activeCell.agent && activeCell.resultType) ? filteredData.filter(d => d.agent === activeCell.agent && d.result === activeCell.resultType) : filteredData, [filteredData, activeCell]);
  const passRate = useMemo(() => filteredData.length === 0 ? 0 : ((filteredData.filter(d => d.result.startsWith('ดีเยี่ยม') || d.result.startsWith('ผ่านเกณฑ์')).length / filteredData.length) * 100).toFixed(1), [filteredData]);

  const totalAuditedFiltered = useMemo(() => {
    return filteredData.filter(d => d.type !== 'ยังไม่ได้ตรวจ' && d.type !== 'N/A' && d.type !== '').length;
  }, [filteredData]);

  const handleMatrixClick = (agentName, type) => {
    if (activeCell.agent === agentName && activeCell.resultType === type) setActiveCell({ agent: null, resultType: null });
    else { setActiveCell({ agent: agentName, resultType: type }); document.getElementById('detail-section')?.scrollIntoView({ behavior: 'smooth' }); }
  };

  const handleToggleFilter = (item, selectedList, setSelectedFn) => {
    selectedList.includes(item) ? setSelectedFn(selectedList.filter(i => i !== item)) : setSelectedFn([...selectedList, item]);
  };

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-zinc-950 flex items-center justify-center p-4 text-white font-sans relative overflow-hidden">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-indigo-600/10 blur-[120px] rounded-full"></div>
        <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-indigo-900/10 blur-[120px] rounded-full"></div>

        <div className="bg-zinc-900/80 backdrop-blur-xl p-8 rounded-[2rem] border border-slate-800 w-full max-w-[360px] text-center shadow-2xl relative z-10 animate-in fade-in zoom-in duration-500">
          <div className="flex justify-center mb-6">
            <div className="p-4 bg-zinc-800 rounded-2xl border border-slate-700 shadow-inner">
                <IntageLogo className="scale-110" />
            </div>
          </div>
          <div className="space-y-1 mb-8">
            <h2 className="text-white font-black uppercase text-xs tracking-[0.3em] italic">CATI CES 2026</h2>
            <p className="text-slate-500 text-[9px] font-bold uppercase tracking-widest">Analytics & Quality Control System</p>
          </div>
          
          <form onSubmit={(e) => { e.preventDefault(); if(inputUser==='Admin'&&inputPass==='1234') setIsAuthenticated(true); else setLoginError('Login Failed'); }} className="space-y-4 text-left">
            <div className="space-y-1.5">
                <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest pl-2">Authorized Username</label>
                <div className="relative">
                    <User className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-500" size={16} />
                    <input type="text" value={inputUser} onChange={e=>setInputUser(e.target.value)} className="w-full pl-11 pr-5 py-3 bg-zinc-950 border border-slate-800 rounded-xl outline-none focus:ring-1 focus:ring-indigo-600 focus:border-transparent text-white transition-all text-sm font-bold placeholder:text-slate-700" placeholder="Username" />
                </div>
            </div>

            <div className="space-y-1.5">
                <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest pl-2">Security Password</label>
                <div className="relative">
                    <Lock className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-500" size={16} />
                    <input type="password" value={inputPass} onChange={e=>setInputPass(e.target.value)} className="w-full pl-11 pr-5 py-3 bg-zinc-950 border border-slate-800 rounded-xl outline-none focus:ring-1 focus:ring-indigo-600 focus:border-transparent text-white transition-all text-sm font-bold placeholder:text-slate-700" placeholder="••••••••" />
                </div>
            </div>

            {loginError && (
                <div className="bg-rose-500/10 border border-rose-500/20 py-2 rounded-lg flex items-center justify-center gap-2">
                    <AlertCircle size={12} className="text-rose-500" />
                    <p className="text-rose-500 text-[9px] font-black uppercase tracking-widest">Authentication Failed</p>
                </div>
            )}

            <button type="submit" className="w-full py-3.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-black uppercase tracking-[0.2em] transition-all shadow-lg shadow-indigo-900/40 flex items-center justify-center gap-2 active:scale-95 italic text-xs">
                <LogIn size={16} />
                Access System
            </button>
          </form>

          <div className="mt-8 pt-6 border-t border-slate-800/50">
            <p className="text-[8px] text-slate-600 font-bold uppercase tracking-widest">© 2026 INTAGE (Thailand) Co., Ltd.</p>
          </div>
        </div>
      </div>
    );
  }

  const FilterSection = ({ title, items, selectedItems, onToggle, onSelectAll, onClear, maxH = "max-h-40" }) => (
    <div className="space-y-2">
        <div className="flex items-center justify-between pl-2">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">{title}</label>
            <div className="flex gap-2">
                <button onClick={onSelectAll} className="text-[9px] font-bold text-slate-500 hover:text-indigo-400 transition-colors">เลือกทั้งหมด</button>
                <button onClick={onClear} className="text-[9px] font-bold text-slate-500 hover:text-indigo-400 transition-colors">ล้าง</button>
            </div>
        </div>
        <div className={`bg-slate-800/50 border border-slate-700/50 rounded-2xl p-2 overflow-y-auto custom-scrollbar ${maxH}`}>
            {items.map(item => (
                <div key={item} onClick={() => onToggle(item)} className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-[10px] font-bold mb-1 transition-all ${selectedItems.includes(item) ? 'bg-indigo-600/20 text-indigo-100 shadow-lg border border-indigo-500/30' : 'hover:bg-slate-700/30 text-slate-400'}`}>
                    {selectedItems.includes(item) ? <CheckSquare size={14} className="text-indigo-400 shrink-0" /> : <Square size={14} className="shrink-0" />}
                    <span className="truncate">{formatResultDisplay(item)}</span>
                </div>
            ))}
        </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-zinc-950 p-4 md:p-8 text-slate-100 font-sans custom-scrollbar">
      
      <div className={`fixed inset-0 z-40 bg-black/60 backdrop-blur-sm transition-opacity duration-300 ${isFilterSidebarOpen ? 'opacity-100' : 'opacity-0 pointer-events-none'}`} onClick={() => setIsFilterSidebarOpen(false)} />
      <aside className={`fixed inset-y-0 right-0 z-50 w-80 bg-zinc-900 shadow-2xl transform transition-transform duration-300 ${isFilterSidebarOpen ? 'translate-x-0' : 'translate-x-full'} overflow-y-auto border-l border-slate-800 p-6`}>
          <div className="flex items-center justify-between mb-8 pb-4 border-b border-slate-800">
             <div className="flex items-center gap-2"><div className="p-2 bg-indigo-600 rounded-lg text-white shadow-lg shadow-indigo-900/20"><Filter size={20} /></div><h3 className="font-black text-white uppercase italic tracking-tight">ตัวกรอง</h3></div>
             <button onClick={() => setIsFilterSidebarOpen(false)} className="text-slate-400 hover:text-white"><X size={20} /></button>
          </div>
          <div className="space-y-8">
             <button onClick={() => { setSelectedMonths([]); setSelectedSups([]); setSelectedAgents([]); setSelectedResults([]); setSelectedTypes([]); setActiveCell({ agent: null, resultType: null }); }} className="w-full py-2.5 text-xs font-black text-indigo-400 bg-indigo-600/10 hover:bg-indigo-600/20 rounded-xl border border-indigo-600/20 uppercase tracking-widest transition-all">ล้างตัวกรองทั้งหมด</button>
             <FilterSection title="เดือน" items={availableMonths} selectedItems={selectedMonths} onToggle={(item) => handleToggleFilter(item, selectedMonths, setSelectedMonths)} onSelectAll={() => setSelectedMonths(availableMonths)} onClear={() => setSelectedMonths([])} />
             <FilterSection title="Supervisor" items={availableSups} selectedItems={selectedSups} onToggle={(item) => handleToggleFilter(item, selectedSups, setSelectedSups)} onSelectAll={() => setSelectedSups(availableSups)} onClear={() => setSelectedSups([])} />
             <FilterSection title="ประเภทงาน (AC / BC)" items={availableTypes} selectedItems={selectedTypes} onToggle={(item) => handleToggleFilter(item, selectedTypes, setSelectedTypes)} onSelectAll={() => setSelectedTypes(availableTypes)} onClear={() => setSelectedTypes([])} />
             <FilterSection title="พนักงาน (ID : ชื่อ)" items={availableAgents} selectedItems={selectedAgents} onToggle={(item) => handleToggleFilter(item, selectedAgents, setSelectedAgents)} onSelectAll={() => setSelectedAgents(availableAgents)} onClear={() => setSelectedAgents([])} maxH="max-h-60" />
             <FilterSection title="ผลการสัมภาษณ์" items={RESULT_ORDER} selectedItems={selectedResults} onToggle={(item) => handleToggleFilter(item, selectedResults, setSelectedResults)} onSelectAll={() => setSelectedResults(RESULT_ORDER)} onClear={() => setSelectedResults([])} maxH="max-h-60" />
          </div>
       </aside>

      <div className="max-w-7xl mx-auto space-y-6">
        <header className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 bg-zinc-900 p-6 rounded-[2.5rem] shadow-xl border border-slate-800">
          <div className="flex items-center gap-6"><div className="p-3 bg-slate-800 rounded-2xl border border-slate-700 shadow-inner"><IntageLogo /></div>
            <div>
              <h1 className="text-xl font-black text-white tracking-tight flex items-center gap-2 uppercase italic">QC REPORT 2026 {loading && <RefreshCw size={18} className="animate-spin text-indigo-500" />}</h1>
              <div className="text-slate-400 text-[10px] font-black uppercase tracking-widest mt-1 flex items-center gap-2 italic">{data.length > 0 ? <><div className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-pulse"></div> REAL-TIME CONNECTED: {data.length} รายการ</> : "OFFLINE"} {lastUpdated && <span className="ml-4 opacity-50"><Clock size={10} className="inline mr-1" />{lastUpdated}</span>}</div>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {/* ปุ่มรีเฟรชข้อมูลใหม่ */}
            <button 
              onClick={() => fetchFromAppsScript(appsScriptUrl)} 
              disabled={loading}
              className={`flex items-center gap-2 px-5 py-3 rounded-2xl text-xs font-black shadow-sm transition-all border ${loading ? 'bg-slate-800 border-slate-700 text-slate-500' : 'bg-zinc-800 border-slate-700 hover:bg-slate-700 text-indigo-400 hover:text-indigo-300 shadow-indigo-900/10'}`}
            >
              <RefreshCw size={16} className={loading ? 'animate-spin' : ''} />
              {loading ? 'SYNC DATA...' : 'SYNC DATA'}
            </button>

            <button onClick={() => setIsFilterSidebarOpen(true)} className={`flex items-center gap-2 px-5 py-3 rounded-2xl text-xs font-black shadow-sm transition-all border ${selectedResults.length > 0 || selectedSups.length > 0 || selectedMonths.length > 0 || selectedAgents.length > 0 || selectedTypes.length > 0 ? 'bg-indigo-600 border-indigo-500 text-white shadow-lg shadow-indigo-900/40' : 'bg-slate-800 border-slate-700 hover:bg-slate-700 text-slate-200'}`}><Filter size={16} /> ตัวกรอง</button>
            <button onClick={() => setShowSync(true)} className="flex items-center gap-2 px-5 py-3 bg-white text-black rounded-2xl text-xs font-black hover:bg-slate-200 transition-all shadow-xl font-bold"><Settings size={14} /> ตั้งค่า</button>
            <button onClick={() => setIsAuthenticated(false)} className="p-3 bg-slate-800 rounded-2xl hover:text-indigo-400 text-slate-400 transition-colors border border-slate-700"><User size={20} /></button>
          </div>
        </header>

        {/* KPI Cards Grid */}
        <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
            {[
                { label: 'จำนวนงานทั้งหมด (กรองแค่เดือน)', value: totalWorkByMonthOnly, icon: FileText, color: 'text-white', bg: 'bg-zinc-900 border-slate-800' },
                { 
                  label: 'จำนวนที่ตรวจแล้ว (AC/BC ตาม Filter)', 
                  value: `${totalAuditedFiltered} (${totalWorkByMonthOnly > 0 ? ((totalAuditedFiltered / totalWorkByMonthOnly) * 100).toFixed(1) : 0}%)`, 
                  icon: Database, 
                  color: 'text-indigo-400', 
                  bg: 'bg-indigo-950/20 border-indigo-900/30' 
                },
                { label: 'อัตราผ่านเกณฑ์', value: `${passRate}%`, icon: CheckCircle, color: 'text-emerald-500', bg: 'bg-emerald-500/10 border-emerald-900/20 shadow-emerald-900/5' },
                { label: 'ควรปรับปรุง', value: filteredData.filter(d=>d.result.startsWith('ควรปรับปรุง')).length, icon: AlertTriangle, color: 'text-amber-500', bg: 'bg-amber-500/10 border-amber-900/20' },
                { label: 'พบข้อผิดพลาด', value: filteredData.filter(d=>d.result.startsWith('พบข้อผิดพลาด')).length, icon: XCircle, color: 'text-rose-500', bg: 'bg-rose-500/10 border-rose-900/20' }
            ].map((kpi, i) => (
                <div key={i} className={`p-6 rounded-[2.5rem] border shadow-sm ${kpi.bg}`}>
                  <div className={`w-8 h-8 rounded-xl flex items-center justify-center mb-3 bg-slate-800/50 ${kpi.color}`}><kpi.icon size={16} /></div>
                  <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">{kpi.label}</p>
                  <h2 className={`text-3xl font-black ${kpi.color} tracking-tighter mt-1 uppercase`}>{kpi.value}</h2>
                </div>
            ))}
        </div>

        {/* --- VISUAL INTELLIGENCE SECTION --- */}
        <div className="flex flex-col lg:flex-row gap-6">
            <div className="bg-zinc-900 p-8 rounded-[3rem] shadow-2xl border border-slate-800 flex-1 flex flex-col min-h-[300px]">
                <h3 className="font-black text-white flex items-center gap-2 italic text-sm uppercase mb-6"><PieChart size={16} className="text-indigo-500" /> Case Composition Summary</h3>
                <div className="flex-1 w-full min-h-[200px]">
                    <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                            <Pie data={chartData} dataKey="count" innerRadius={60} outerRadius={85} paddingAngle={5}>
                                {chartData.map((entry, index) => <Cell key={index} fill={entry.color} />)}
                            </Pie>
                            <Tooltip contentStyle={{backgroundColor: '#0f172a', border: '1px solid #1e293b', borderRadius: '12px', fontSize: '10px', color: '#f8fafc'}} />
                        </PieChart>
                    </ResponsiveContainer>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mt-6">
                    {chartData.map(c => (
                        <div key={c.full} className="flex items-center gap-2">
                            <div className="w-2 h-2 rounded-full shrink-0" style={{backgroundColor: c.color}}></div>
                            <span className="text-[9px] text-slate-500 font-bold truncate uppercase">{c.name}</span>
                        </div>
                    ))}
                </div>
            </div>

             {/* --- NEW: Monthly Performance Chart --- */}
             <div className="bg-zinc-900 p-8 rounded-[3rem] shadow-2xl border border-slate-800 flex-1 flex flex-col min-h-[300px]">
                <h3 className="font-black text-white flex items-center gap-2 italic text-sm uppercase mb-6"><BarChart2 size={16} className="text-emerald-500" /> Monthly Audit Progress (%)</h3>
                <div className="flex-1 w-full min-h-[200px]">
                    <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={monthlyPerformanceData}>
                            <defs>
                                <linearGradient id="barGradient" x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%" stopColor="#10B981" stopOpacity={1}/>
                                    <stop offset="100%" stopColor="#059669" stopOpacity={0.6}/>
                                </linearGradient>
                            </defs>
                            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                            <XAxis dataKey="name" stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                            <YAxis stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} unit="%" />
                            <Tooltip 
                                cursor={{fill: '#1e293b', opacity: 0.4}}
                                content={({ active, payload, label }) => {
                                    if (active && payload && payload.length) {
                                    return (
                                        <div className="bg-slate-900 border border-slate-700 p-3 rounded-xl shadow-xl">
                                            <p className="text-slate-300 text-[10px] font-bold uppercase tracking-widest mb-1">{label}</p>
                                            <p className="text-emerald-400 text-lg font-black">{payload[0].value}%</p>
                                            <p className="text-slate-500 text-[10px] font-bold">Audited: {payload[0].payload.audited} / {payload[0].payload.total}</p>
                                        </div>
                                    );
                                    }
                                    return null;
                                }}
                            />
                            <Bar dataKey="percent" fill="url(#barGradient)" radius={[6, 6, 0, 0]} barSize={40} animationDuration={1500}>
                                {monthlyPerformanceData.map((entry, index) => (
                                    <Cell key={`cell-${index}`} fill={entry.percent >= 100 ? '#3b82f6' : 'url(#barGradient)'} />
                                ))}
                            </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>
            </div>
        </div>

        <div className="bg-zinc-900 rounded-[3rem] shadow-2xl border border-slate-800 overflow-hidden">
            <div className="p-8 border-b border-slate-800 bg-zinc-900/50 flex items-center justify-between">
                <div><h3 className="font-black text-white flex items-center gap-2 italic text-lg uppercase tracking-tight"><Users size={20} className="text-indigo-500" /> สรุปพนักงาน (ID : ชื่อ) x ผลการตรวจ</h3></div>
            </div>
            <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                <table className="w-full text-left text-sm border-separate border-spacing-0 min-w-[1200px]">
                    <thead className="sticky top-0 bg-zinc-900 z-20 font-black text-slate-200 text-[10px] uppercase tracking-widest border-b border-slate-800 shadow-md">
                        <tr>
                            <th rowSpan="2" className="px-8 py-6 border-b border-slate-800 border-r border-slate-800 bg-zinc-900 w-64 text-slate-100 italic">Interviewer (ID : Name)</th>
                            <th colSpan={RESULT_ORDER.length} className="px-4 py-4 text-center border-b border-slate-800 bg-zinc-900/40 text-indigo-400 text-[11px] font-black italic uppercase">QC Result Distribution</th>
                            <th rowSpan="2" className="px-8 py-6 text-center bg-slate-800 text-white border-b border-slate-800 border-l border-slate-800">Total</th>
                        </tr>
                        <tr className="bg-zinc-900/80 text-white">
                            {RESULT_ORDER.map(type => <th key={type} className="px-4 py-3 text-center border-b border-slate-800 border-r border-slate-800 max-w-[180px] text-slate-400"><span className="line-clamp-2" title={type}>{formatResultDisplay(type)}</span></th>)}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800 font-bold">
                        {agentSummary.map((agent, i) => (
                        <tr key={i} className="hover:bg-indigo-950/20 transition-colors group">
                            <td className="px-8 py-5 text-white border-r border-slate-800 font-medium">{agent.name}</td>
                            {RESULT_ORDER.map(type => {
                            const val = agent[type]; const isActive = activeCell.agent === agent.name && activeCell.resultType === type;
                            return (
                                <td key={type} className={`px-4 py-5 text-center border-r border-slate-800 transition-all ${val > 0 ? 'cursor-pointer hover:bg-slate-700/50 shadow-inner' : ''} ${isActive ? 'bg-indigo-600/20 ring-2 ring-inset ring-indigo-500' : ''}`} onClick={() => val > 0 && handleMatrixClick(agent.name, type)}>
                                    <span className={`text-sm font-black ${val > 0 ? '' : 'text-slate-800'}`} style={{ color: val > 0 ? getResultColor(type) : undefined }}>{val || '-'}</span>
                                </td>
                            );
                            })}
                            <td className="px-8 py-5 text-center bg-slate-800/20 text-white border-l border-slate-800">{agent.total}</td>
                        </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>

        <div id="detail-section" className="bg-zinc-900 rounded-[3rem] shadow-2xl border border-slate-800 overflow-hidden scroll-mt-6">
            <div className="p-8 border-b border-slate-800 flex flex-col md:flex-row items-center justify-between gap-6">
                <div className="space-y-1">
                    <h3 className="font-black text-white uppercase tracking-widest text-xs flex items-center gap-2 italic"><MessageSquare size={16} className="text-indigo-500" /> รายละเอียดรายเคส</h3>
                    {activeCell.agent && <div className="flex items-center gap-2 mt-2"><span className="px-3 py-1 bg-indigo-600 text-white text-[9px] font-black rounded-lg uppercase italic animate-pulse">กำลังแสดง: {activeCell.agent}</span><button onClick={() => setActiveCell({ agent: null, resultType: null })} className="text-slate-500 hover:text-white"><X size={12}/></button></div>}
                </div>
                <div className="relative w-full md:w-80"><Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500" size={16} /><input type="text" placeholder="ค้นหาพนักงาน หรือ เลขชุด..." className="w-full pl-12 pr-6 py-4 bg-slate-800/30 border border-slate-700 rounded-2xl text-sm font-bold focus:ring-2 focus:ring-indigo-600 outline-none text-white shadow-inner" value={searchTerm} onChange={(e)=>setSearchTerm(e.target.value)} /></div>
            </div>
            <div className="overflow-auto max-h-[1000px] custom-scrollbar">
                <table className="w-full text-left text-xs font-medium border-separate border-spacing-0">
                <thead className="sticky top-0 bg-zinc-900 shadow-md z-10 border-b border-slate-800 font-black text-slate-300 uppercase tracking-widest">
                    <tr><th className="px-8 py-5 border-r border-slate-800/50">วันที่ / เลขชุด</th><th className="px-8 py-5 border-r border-slate-800/50">พนักงาน (ID : ชื่อ)</th><th className="px-4 py-5 text-center border-r border-slate-800/50">ผลสรุป</th><th className="px-8 py-5">QC Comment & Audio</th></tr>
                </thead>
                <tbody className="divide-y divide-slate-800/50">
                    {detailLogs.length > 0 ? detailLogs.slice(0, 150).map((item) => {
                        const isExpanded = expandedCaseId === item.id; const isEditing = editingCase && editingCase.id === item.id;
                        const isNewAudit = item.type === "ยังไม่ได้ตรวจ";

                        return (
                        <React.Fragment key={item.id}>
                            <tr onClick={() => !isEditing && setExpandedCaseId(isExpanded ? null : item.id)} className={`transition-all group cursor-pointer ${isExpanded ? 'bg-indigo-950/20 shadow-inner' : 'hover:bg-slate-800/40'}`}>
                                <td className="px-8 py-6 border-r border-slate-800/20"><div className="font-black text-white">{item.date}</div><div className="flex items-center gap-1.5 mt-1 bg-zinc-950 px-2 py-0.5 rounded border border-zinc-800 w-fit"><Hash size={10} className="text-indigo-400" /><span className="text-[11px] font-black text-slate-300">{item.questionnaireNo}</span></div></td>
                                <td className="px-8 py-6 border-r border-slate-800/20"><div className="font-black text-white text-sm group-hover:text-indigo-400 transition-colors flex items-center gap-2">{item.agent} {isExpanded ? <ChevronDown size={14} className="text-slate-500" /> : <ChevronRight size={14} className="text-slate-700" />}</div><div className="text-[9px] text-slate-500 font-bold mt-0.5 italic uppercase font-sans tracking-wider">TYPE: {item.type} &bull; SUP: {item.supervisor}</div></td>
                                <td className="px-4 py-6 text-center border-r border-slate-800/20"><span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[9px] font-black border uppercase shadow-sm" style={{ backgroundColor: `${getResultColor(item.result)}10`, color: getResultColor(item.result), borderColor: `${getResultColor(item.result)}30` }}>{formatResultDisplay(item.result)}</span></td>
                                <td className="px-8 py-6">
                                    <p className="text-slate-400 italic max-w-sm truncate group-hover:text-slate-200 transition-colors font-sans leading-relaxed">{item.comment ? `"${item.comment}"` : '-'}</p>
                                    {item.audio && item.audio.includes('http') && (<a href={item.audio} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()} className="mt-2 inline-flex items-center gap-1 text-indigo-400 hover:text-indigo-300 font-black text-[9px] uppercase transition-all hover:translate-x-1"><PlayCircle size={14} /> Listen Recording</a>)}
                                </td>
                            </tr>
                            {isExpanded && (
                            <tr className="bg-slate-900/40 animate-in slide-in-from-top-2 duration-300">
                                <td colSpan={4} className="p-8 border-b border-slate-800 text-white">
                                <div className="flex items-center justify-between mb-8">
                                    <div className="flex items-center gap-4"><div className="p-3 bg-indigo-500/10 rounded-2xl text-indigo-400 shadow-inner"><Award /></div><div><h4 className="font-black uppercase italic tracking-widest text-sm text-white">{isNewAudit ? "START AUDIT SESSION" : "ASSESSMENT DETAIL"} (ID: {item.questionnaireNo})</h4><p className="text-[9px] text-slate-500 font-bold uppercase tracking-widest italic leading-relaxed">{isNewAudit ? "กรุณากรอกคะแนนและผลสรุปเพื่อบันทึกงานใหม่" : "แก้ไขคะแนน P:AB และ สรุปผล M"}</p></div></div>
                                    <div className="flex gap-2">
                                    {!isEditing ? (
                                        <button onClick={() => setEditingCase({...item})} className={`flex items-center gap-2 px-6 py-3 rounded-2xl text-[10px] font-black uppercase border transition-all ${isNewAudit ? 'bg-indigo-600 hover:bg-indigo-500 text-white border-indigo-500 shadow-lg shadow-indigo-900/20' : 'bg-slate-800 hover:bg-slate-700 text-slate-200 border-slate-700'}`}>
                                            <Edit2 size={12} className={isNewAudit ? "text-white" : "text-indigo-400"}/> 
                                            {isNewAudit ? "เริ่มตรวจงานนี้" : "แก้ไขข้อมูล"}
                                        </button>
                                    ) : (
                                        <div className="flex gap-2"><button disabled={isSaving} onClick={handleUpdateCase} className="flex items-center gap-2 px-6 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-2xl text-[10px] font-black uppercase shadow-lg shadow-indigo-900/20">{isSaving ? <RefreshCw className="animate-spin" size={14}/> : <Save size={14}/>} {isNewAudit ? "บันทึกผลการตรวจ" : "บันทึกการแก้ไข"}</button><button onClick={() => setEditingCase(null)} className="px-6 py-3 bg-slate-800 text-white rounded-2xl text-[10px] font-black uppercase transition-all border border-slate-700">ยกเลิก</button></div>
                                    )}
                                    </div>
                                </div>

                                {isEditing && (
                                    <div className="mb-8 p-6 bg-indigo-950/20 border border-indigo-500/20 rounded-[2rem] shadow-inner">
                                        <div className="flex items-center gap-2 mb-4 text-indigo-400 font-black text-[10px] uppercase italic tracking-widest"><Info size={16} /> กำหนดประเภทงาน (AC / BC)</div>
                                        <div className="flex gap-4">
                                            {['AC', 'BC'].map(t => (
                                                <button key={t} onClick={() => setEditingCase({...editingCase, type: t})} className={`px-8 py-3 rounded-xl text-[10px] font-black uppercase transition-all border ${editingCase.type === t ? 'bg-indigo-600 text-white border-indigo-500 shadow-lg shadow-indigo-900/40' : 'bg-zinc-950 text-slate-500 border-slate-800 hover:border-slate-700'}`}>
                                                    {t} MODE
                                                </button>
                                            ))}
                                            {editingCase.type === "ยังไม่ได้ตรวจ" && <p className="text-rose-500 text-[10px] font-black uppercase self-center animate-pulse">*** กรุณาเลือก AC หรือ BC เพื่อบันทึกงาน</p>}
                                        </div>
                                    </div>
                                )}

                                {item.audio && item.audio.includes('http') && (
                                    <div className="mb-8 p-6 bg-slate-950/50 border border-slate-800 rounded-[2rem] flex items-center justify-between shadow-inner">
                                        <div className="flex items-center gap-3">
                                            <div className="p-3 bg-indigo-600 text-white rounded-2xl animate-pulse"><PlayCircle size={20} /></div>
                                            <div><p className="text-[10px] font-black uppercase text-slate-500 tracking-widest">Audio Recording File</p><p className="text-xs font-black text-slate-200">ไฟล์เสียงบันทึกการสัมภาษณ์</p></div>
                                        </div>
                                        <a href={item.audio} target="_blank" rel="noopener noreferrer" className="px-6 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-2xl text-[10px] font-black uppercase flex items-center gap-2 transition-all shadow-lg shadow-indigo-900/20"><ExternalLink size={14} /> OPEN PLAYER</a>
                                    </div>
                                )}
                                <div className="mb-8 p-6 bg-slate-900/80 border border-slate-800 rounded-[2rem] shadow-inner">
                                    <div className="flex items-center gap-2 mb-4 text-indigo-400 font-black text-[10px] uppercase italic tracking-widest"><Star size={16} /> สรุปผลการสัมภาษณ์แบบละเอียด (Column M)</div>
                                    {isEditing ? (<div className="relative"><select className="w-full p-4 bg-slate-950 border border-indigo-600/30 rounded-2xl text-[11px] font-black text-white focus:ring-2 focus:ring-indigo-600 outline-none appearance-none shadow-lg shadow-indigo-900/10" value={editingCase.result} onChange={e => setEditingCase({...editingCase, result: e.target.value})}>{RESULT_ORDER.map(opt => <option key={opt} value={opt}>{opt}</option>)}</select><ChevronDown size={14} className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-600 pointer-events-none" /></div>) : (<p className="text-sm font-black italic text-slate-200 bg-slate-950/50 p-4 rounded-xl border border-slate-800 leading-relaxed shadow-sm">{item.result}</p>)}
                                </div>
                                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3">
                                    {(isEditing ? editingCase : item).evaluations.map((evalItem, eIdx) => (
                                    <div key={eIdx} className={`bg-slate-900 border p-3 rounded-2xl transition-all ${isEditing ? 'border-indigo-600/50 ring-1 ring-indigo-600/20' : 'border-slate-800'}`}>
                                        <p className="text-[9px] font-black text-slate-500 uppercase tracking-tighter mb-2 truncate" title={evalItem.label}>{evalItem.label}</p>
                                        {isEditing ? (<div className="relative"><select className="w-full bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-[10px] font-black text-white appearance-none outline-none focus:ring-1 focus:ring-indigo-600" value={evalItem.value} onChange={(e) => { const newEvals = [...editingCase.evaluations]; newEvals[eIdx].value = e.target.value; setEditingCase({...editingCase, evaluations: newEvals}); }}>{SCORE_OPTIONS.map(opt => (<option key={opt} value={opt}>{SCORE_LABELS[opt]}</option>))}</select><ChevronDown size={10} className="absolute right-1 top-1/2 -translate-y-1/2 text-zinc-600 pointer-events-none" /></div>) : (<div className={`text-sm font-black italic tracking-widest ${evalItem.value === '5' || evalItem.value === '4' ? 'text-emerald-500' : (evalItem.value === '1' || evalItem.value === '2') ? 'text-rose-500' : 'text-slate-300'}`}>{SCORE_LABELS[evalItem.value] || evalItem.value}</div>)}
                                    </div>
                                    ))}
                                </div>
                                <div className="mt-8">
                                    <p className="text-[10px] font-black text-indigo-400 uppercase mb-2 italic tracking-widest flex items-center gap-2"><MessageSquare size={12}/> QC Full Comment (Column AC)</p>
                                    {isEditing ? (<textarea className="w-full bg-slate-950 border border-slate-800 rounded-[1.5rem] p-6 text-sm italic text-white outline-none min-h-[120px] focus:ring-2 focus:ring-indigo-600 shadow-inner font-sans" value={editingCase.comment} onChange={e => setEditingCase({...editingCase, comment: e.target.value})}/>) : (<div className="p-4 bg-slate-800/30 rounded-2xl border border-slate-800 border-l-4 border-l-indigo-500 shadow-sm"><p className="text-sm text-slate-200 font-medium italic leading-relaxed font-sans">"{item.comment || 'ไม่มีคอมเมนต์'}"</p></div>)}
                                </div>
                                </td>
                            </tr>
                            )}
                        </React.Fragment>
                        );
                    }) : (<tr><td colSpan={4} className="px-8 py-24 text-center text-slate-500 font-black uppercase italic tracking-widest text-lg opacity-20 italic">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</td></tr>)}
                </tbody>
                </table>
            </div>
        </div>
      </div>

      {(showSync || error) && (
        <div className="fixed inset-0 z-[100] bg-black/95 flex items-center justify-center p-4">
            <div className="bg-zinc-900 w-full max-w-2xl rounded-[3rem] p-10 border border-slate-800 shadow-2xl relative animate-in zoom-in duration-300">
                <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600"></div>
                <div className="flex items-center justify-between mb-10"><h3 className="text-xl font-black flex items-center gap-3 text-white uppercase italic tracking-tight">{error ? 'Connection Error' : 'ตั้งค่าระบบ'}</h3>{data.length > 0 && <button onClick={() => {setShowSync(false); setError(null);}} className="text-slate-500 hover:text-white transition-colors"><X size={28} /></button>}</div>
                <div className="space-y-6">
                    <div className="p-4 bg-indigo-900/30 border border-indigo-500/30 rounded-2xl mb-4">
                      <p className="text-xs text-indigo-300 font-bold mb-1 flex items-center gap-2"><Zap size={14}/> SYSTEM UPGRADE (JSON API MODE)</p>
                      <p className="text-[10px] text-slate-400">ระบบได้รับการอัปเกรดเป็นโหมดประสิทธิภาพสูง กรุณาใส่ Web App URL ของ Apps Script ลงในช่องด้านล่างเพียงช่องเดียว</p>
                    </div>
                    <div className="space-y-2"><label className="text-[10px] font-black text-slate-300 uppercase tracking-widest pl-2 italic">Apps Script Web App URL (Read & Write)</label><input type="text" className="w-full px-6 py-4 bg-slate-950 border border-slate-800 rounded-2xl text-xs font-bold text-white outline-none focus:ring-2 focus:ring-indigo-600 shadow-inner" value={appsScriptUrl} onChange={e=>{setAppsScriptUrl(e.target.value); localStorage.setItem('apps_script_url', e.target.value);}} /></div>
                    <div className="flex gap-4 pt-4"><button onClick={() => fetchFromAppsScript(appsScriptUrl)} disabled={loading || !appsScriptUrl} className="flex-1 py-5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-[2.5rem] font-black uppercase shadow-xl shadow-indigo-900/30 text-sm tracking-widest italic flex items-center justify-center gap-2 transition-all">{loading ? <RefreshCw className="animate-spin" size={18}/> : <RefreshCw size={18}/>} CONNECT & SYNC</button></div>
                    {error && <p className="text-center text-[10px] text-rose-500 font-black mt-2 uppercase italic animate-pulse">{error}</p>}
                </div>
            </div>
        </div>
      )}
    </div>
  );
};

export default App;