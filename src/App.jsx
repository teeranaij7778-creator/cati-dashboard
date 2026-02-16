import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell, PieChart, Pie, ComposedChart, Line
} from 'recharts';
import { 
  Users, CheckCircle, AlertTriangle, XCircle, Search, 
  FileText, BarChart2, MessageSquare, Calendar, TrendingUp, Database, Link, RefreshCw, Trash2, Globe, FilterX, PlayCircle, UserCheck, Settings, AlertCircle, Info, ChevronRight, ExternalLink, User, ChevronDown, CheckSquare, Square, X, Briefcase, Lock, LogIn, Activity, Filter, Check, Clock, ListChecks, Award, Save, Edit2, Hash, Star, Zap, MousePointerClick, ShieldCheck, UserPlus, MapPin
} from 'lucide-react';

/** * CATI CES 2026 Analytics Dashboard - MASTER VERSION (V2.5 LIGHT THEME + AGENT TREND)
 * - THEME: LIGHT MODE (Clean White/Slate)
 * - FEATURE: "Agent Performance Trend" Chart (Idea #2)
 * - Shows Total Cases (Bar) vs Pass Rate % (Line) per month when an agent is selected.
 * - FEATURE: Grand Total Row included with Vertical %
 * - FEATURE: Row Total Column includes % share
 * - FEATURE: Added TOUCH_POINT (Column F) Display & Filter
 * - UPDATE: Changed 'Pass Rate %' KPI to 'Pass Count'
 * - USER ROLE: Admin (Full), QC (Edit Only), INV (View Only)
 */

const DEFAULT_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbytB3UpN0xv7kNcuk-XqFmwoB6LWekjIEj0B9b8H5Me25mQ0ozy69NniuRvM_uNjWD5/exec";

// ลำดับการแสดงผล (ใช้ในการ Match ค่าจาก Sheet)
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

const SUPERVISOR_OPTIONS = ['เสกข์พลกฤต', 'ศรัณยกร', 'นิตยา', 'มณีรัตน์', 'Gallup'];

const formatResultDisplay = (text) => (text ? text.split('(')[0].trim() : '-');

const COLORS = {
  'ดีเยี่ยม': '#6366f1',      // Indigo-500
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
    <div className="w-4 h-4 rounded-full bg-indigo-600 animate-pulse"></div>
    <span className="font-black tracking-[0.15em] text-slate-800 italic text-lg">INTAGE</span>
  </div>
);

const App = () => {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [userRole, setUserRole] = useState(null); // 'Admin', 'QC', or 'INV'
  const [inputUser, setInputUser] = useState('');
  const [inputPass, setInputPass] = useState('');
  const [loginError, setLoginError] = useState('');
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState(null);
  
  const recentEdits = useRef(new Map());

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
  const [selectedTouchpoints, setSelectedTouchpoints] = useState([]);
  
  const [activeKpiFilter, setActiveKpiFilter] = useState(null); // 'audited', 'pass', 'improve', 'error'
  
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
      intervalId = setInterval(() => fetchFromAppsScript(appsScriptUrl), 60000); 
    }
    return () => clearInterval(intervalId);
  }, [appsScriptUrl, isAuthenticated]);

  const fetchFromAppsScript = async (urlToFetch) => {
    setLoading(true); setError(null);
    try {
      const fetchUrl = `${urlToFetch}${urlToFetch.includes('?') ? '&' : '?'}t=${Date.now()}`;
      const response = await fetch(fetchUrl);
      
      if (!response.ok) throw new Error(`Server returned ${response.status} ${response.statusText}`);

      const text = await response.text();
      let allRows;
      try { allRows = JSON.parse(text); } 
      catch (jsonErr) {
        const preview = text.substring(0, 50);
        throw new Error(text.trim().startsWith('<') ? `URL ส่งกลับมาเป็น HTML: ${preview}...` : `ได้รับข้อมูลที่ไม่ใช่ JSON: "${preview}..."`);
      }
      
      if (allRows.error) throw new Error(`Apps Script Error: ${allRows.error}`);
      if (!Array.isArray(allRows)) throw new Error("Format ข้อมูลไม่ถูกต้อง (ต้องเป็น Array)");
      if (allRows.length === 0) throw new Error("เชื่อมต่อสำเร็จ แต่ Sheet ว่างเปล่า!");

      let headerIdx = allRows.findIndex(row => 
        Array.isArray(row) && row.some(cell => 
          (cell && cell.toString().toLowerCase().includes("interviewer")) || 
          (cell && cell.toString().includes("สรุปผล")) ||
          (cell && cell.toString().toLowerCase().includes("name"))
        )
      );

      if (headerIdx === -1) {
         headerIdx = allRows.findIndex(row => Array.isArray(row) && row.some(cell => (cell && cell.toString().toLowerCase().includes("date")) || (cell && cell.toString().includes("วันที่"))));
      }

      if (headerIdx === -1 && allRows.length > 0) headerIdx = 0;
      if (headerIdx === -1) throw new Error("ไม่พบคอลัมน์ข้อมูลที่กำหนด (กรุณาตรวจสอบ Header Row)");

      const headers = allRows[headerIdx].map(h => h ? h.toString().trim() : "");
      
      const getIdx = (keywords) => {
        if (!Array.isArray(keywords)) keywords = [keywords];
        return headers.findIndex(h => {
            if (!h) return false;
            const lowerH = h.toLowerCase();
            return keywords.some(k => lowerH.includes(k.toLowerCase()));
        });
      };

      const idx = {
        month: getIdx(["เดือน", "month"]), 
        date: getIdx(["วันที่สัมภาษณ์", "date", "timestamp"]), 
        type: getIdx(["AC / BC", "type", "ac/bc"]), 
        sup: getIdx(["Supervisor", "sup"]), 
        questionnaireNo: getIdx(["questionnaire", "no.", "เลขชุด", "id"]), 
        audio: getIdx(["ไฟล์เสียง", "audio", "record"]) 
      };

      const COL_TOUCHPOINT = 5; 
      const COL_INTERVIEWER_ID = 9;  
      const COL_INTERVIEWER_NAME = 10; 
      const COL_RESULT = 12; 
      const COL_CATI_SUPERVISOR = 7; 
      const COL_QC_COMMENT = 13; 
      const EVAL_START_INDEX = 15; 
      const EVAL_COUNT = 13; 

      let parsedData = allRows.slice(headerIdx + 1)
        .map((row, i) => ({ row, actualRowNumber: i + headerIdx + 2 }))
        .filter(({ row }) => {
          if (!row || !Array.isArray(row)) return false;
          const agentCode = row[COL_INTERVIEWER_NAME]?.toString().trim() || "";
          return agentCode !== "" && !agentCode.includes("#N/A");
        })
        .map(({ row, actualRowNumber }, index) => {
          let rawResult = (row[COL_RESULT]) ? row[COL_RESULT].toString().trim() : "N/A";
          let cleanResult = rawResult;
          const matchedResult = RESULT_ORDER.find(opt => {
              const prefix = opt.split(':')[0].trim();
              return rawResult.startsWith(prefix);
          });
          if (matchedResult) cleanResult = matchedResult;
          
          const interviewerId = row[COL_INTERVIEWER_ID] ? row[COL_INTERVIEWER_ID].toString().trim() : '-';
          const agentId = row[COL_INTERVIEWER_NAME] ? row[COL_INTERVIEWER_NAME].toString().trim() : 'Unknown';
          const displayAgent = `${interviewerId} : ${agentId}`;
          const supervisorVal = (row[COL_CATI_SUPERVISOR]) ? row[COL_CATI_SUPERVISOR].toString() : '';
          const commentVal = (row[COL_QC_COMMENT]) ? row[COL_QC_COMMENT].toString() : '';
          const touchpointVal = (row[COL_TOUCHPOINT]) ? row[COL_TOUCHPOINT].toString() : 'N/A';
          const rawType = (idx.type !== -1 && row[idx.type]) ? row[idx.type].toString().trim() : "";
          const cleanType = (rawType === "" || rawType === "N/A") ? "ยังไม่ได้ตรวจ" : rawType;

          const evaluations = [];
          for (let i = 0; i < EVAL_COUNT; i++) {
              const currentIdx = EVAL_START_INDEX + i;
              const label = headers[currentIdx] || `Criteria ${i+1}`;
              const value = (row[currentIdx] !== undefined && row[currentIdx] !== null) ? row[currentIdx].toString() : '-';
              evaluations.push({ label: label, value: value });
          }

          return {
            id: index, rowIndex: actualRowNumber, 
            month: (idx.month !== -1 && row[idx.month]) ? row[idx.month] : 'N/A', 
            date: (idx.date !== -1 && row[idx.date]) ? row[idx.date] : 'N/A', 
            agent: displayAgent, 
            rawName: agentId,    
            interviewerId: interviewerId, 
            questionnaireNo: (idx.questionnaireNo !== -1 && row[idx.questionnaireNo]) ? row[idx.questionnaireNo] : '-', 
            result: cleanResult, 
            supervisor: supervisorVal, 
            comment: commentVal, 
            audio: (idx.audio !== -1 && row[idx.audio]) ? row[idx.audio] : '', 
            touchpoint: touchpointVal, 
            supervisorFilter: (idx.sup !== -1 && row[idx.sup]) ? row[idx.sup] : 'N/A',
            type: cleanType,
            evaluations: evaluations
          };
        });

      parsedData = parsedData.map(item => {
        const cached = recentEdits.current.get(item.id);
        if (cached && Date.now() - cached.timestamp < 60000) return cached.data;
        else if (cached) recentEdits.current.delete(item.id);
        return item;
      });

      setData(parsedData); 
      setLastUpdated(new Date().toLocaleTimeString('th-TH'));
      localStorage.setItem('apps_script_url', urlToFetch);
      setShowSync(false);
    } catch (err) { setError(err.message); console.error("Fetch Error:", err); } finally { setLoading(false); }
  };

  const handleUpdateCase = async () => {
    if (!appsScriptUrl) return alert("กรุณาตั้งค่า Apps Script URL");
    if (editingCase.type === "ยังไม่ได้ตรวจ") return alert("กรุณาเลือกประเภทงาน (AC หรือ BC) ก่อนบันทึก");

    setIsSaving(true);
    const backupData = [...data];
    const updatedItem = { ...editingCase };
    recentEdits.current.set(updatedItem.id, { data: updatedItem, timestamp: Date.now() });

    setData(prevData => prevData.map(item => item.id === editingCase.id ? updatedItem : item));

    try {
      const updateData = {
        rowIndex: editingCase.rowIndex, result: editingCase.result, type: editingCase.type,
        supervisor: String(editingCase.supervisor || ''), evaluations: editingCase.evaluations.map(e => e.value), comment: editingCase.comment
      };
      await fetch(appsScriptUrl, { method: 'POST', mode: 'no-cors', cache: 'no-cache', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(updateData) });
      setEditingCase(null); setIsSaving(false);
    } catch (err) { 
      alert("เกิดข้อผิดพลาดในการบันทึก: " + err.message); 
      recentEdits.current.delete(editingCase.id);
      setData(backupData); setIsSaving(false); 
    }
  };

  const availableMonths = useMemo(() => [...new Set(data.map(d => d.month).filter(m => m !== 'N/A'))].sort(), [data]);
  const availableSups = useMemo(() => [...new Set(data.map(d => d.supervisorFilter).filter(s => s !== 'N/A'))].sort(), [data]);
  const availableTypes = useMemo(() => [...new Set(data.map(d => d.type).filter(t => t !== 'N/A' && t !== ''))].sort(), [data]);
  const availableTouchpoints = useMemo(() => [...new Set(data.map(d => d.touchpoint).filter(t => t !== 'N/A' && t !== ''))].sort(), [data]);

  const availableAgents = useMemo(() => {
    let filtered = data;
    if (selectedSups.length > 0) filtered = filtered.filter(d => selectedSups.includes(d.supervisorFilter));
    if (selectedMonths.length > 0) filtered = filtered.filter(d => selectedMonths.includes(d.month));
    if (selectedTypes.length > 0) filtered = filtered.filter(d => selectedTypes.includes(d.type));
    if (selectedTouchpoints.length > 0) filtered = filtered.filter(d => selectedTouchpoints.includes(d.touchpoint));
    return [...new Set(filtered.map(d => d.agent).filter(a => a !== 'Unknown'))].sort();
  }, [data, selectedSups, selectedMonths, selectedTypes, selectedTouchpoints]);

  const baseFilteredData = useMemo(() => {
    return data.filter(item => {
      const matchesSearch = item.agent.toLowerCase().includes(searchTerm.toLowerCase()) || item.comment.toLowerCase().includes(searchTerm.toLowerCase()) || item.questionnaireNo.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesResult = selectedResults.length === 0 || selectedResults.includes(item.result);
      const matchesSup = selectedSups.length === 0 || selectedSups.includes(item.supervisorFilter);
      const matchesAgent = selectedAgents.length === 0 || selectedAgents.includes(item.agent);
      const matchesMonth = selectedMonths.length === 0 || selectedMonths.includes(item.month);
      const matchesType = selectedTypes.length === 0 || selectedTypes.includes(item.type);
      const matchesTouchpoint = selectedTouchpoints.length === 0 || selectedTouchpoints.includes(item.touchpoint);
      return matchesSearch && matchesResult && matchesSup && matchesAgent && matchesMonth && matchesType && matchesTouchpoint;
    });
  }, [data, searchTerm, selectedResults, selectedSups, selectedAgents, selectedMonths, selectedTypes, selectedTouchpoints]);

  const finalFilteredData = useMemo(() => {
    if (!activeKpiFilter) return baseFilteredData;
    return baseFilteredData.filter(item => {
        if (activeKpiFilter === 'audited') return item.type !== 'ยังไม่ได้ตรวจ' && item.type !== 'N/A' && item.type !== '';
        if (activeKpiFilter === 'pass') return item.result.startsWith('ดีเยี่ยม') || item.result.startsWith('ผ่านเกณฑ์');
        if (activeKpiFilter === 'improve') return item.result.startsWith('ควรปรับปรุง');
        if (activeKpiFilter === 'error') return item.result.startsWith('พบข้อผิดพลาด');
        return true;
    });
  }, [baseFilteredData, activeKpiFilter]);

  const totalWorkByMonthOnly = useMemo(() => {
    if (selectedMonths.length === 0) return data.length;
    return data.filter(item => selectedMonths.includes(item.month)).length;
  }, [data, selectedMonths]);

  const agentSummary = useMemo(() => {
    const summaryMap = {};
    finalFilteredData.forEach(item => {
      if (!summaryMap[item.agent]) { summaryMap[item.agent] = { name: item.agent, total: 0 }; RESULT_ORDER.forEach(r => summaryMap[item.agent][r] = 0); }
      if (summaryMap[item.agent][item.result] !== undefined) summaryMap[item.agent][item.result] += 1;
      summaryMap[item.agent].total += 1;
    });
    return Object.values(summaryMap).sort((a, b) => b.total - a.total);
  }, [finalFilteredData]);

  const totalSummary = useMemo(() => {
    const totals = { total: 0 };
    RESULT_ORDER.forEach(r => totals[r] = 0);
    agentSummary.forEach(agent => { totals.total += agent.total; RESULT_ORDER.forEach(r => { totals[r] += (agent[r] || 0); }); });
    return totals;
  }, [agentSummary]);

  const chartData = useMemo(() => {
    const total = finalFilteredData.length;
    return RESULT_ORDER.map(key => ({ 
      name: formatResultDisplay(key), full: key, count: finalFilteredData.filter(d => d.result === key).length, 
      percent: total > 0 ? ((finalFilteredData.filter(d => d.result === key).length / total) * 100).toFixed(1) : 0, 
      color: getResultColor(key) 
    }));
  }, [finalFilteredData]);

  const passCount = useMemo(() => baseFilteredData.filter(d => d.result.startsWith('ดีเยี่ยม') || d.result.startsWith('ผ่านเกณฑ์')).length, [baseFilteredData]);
  const passRate = useMemo(() => baseFilteredData.length === 0 ? 0 : ((passCount / baseFilteredData.length) * 100).toFixed(1), [baseFilteredData, passCount]);
  
  const totalAuditedFiltered = useMemo(() => baseFilteredData.filter(d => d.type !== 'ยังไม่ได้ตรวจ' && d.type !== 'N/A' && d.type !== '').length, [baseFilteredData]);

  const detailLogs = useMemo(() => (activeCell.agent && activeCell.resultType) ? finalFilteredData.filter(d => d.agent === activeCell.agent && d.result === activeCell.resultType) : finalFilteredData, [finalFilteredData, activeCell]);

  const monthlyPerformanceData = useMemo(() => {
    return availableMonths.map(month => {
        const monthData = data.filter(d => d.month === month);
        const total = monthData.length;
        const audited = monthData.filter(d => d.type !== 'ยังไม่ได้ตรวจ' && d.type !== 'N/A' && d.type !== '').length;
        const percent = total > 0 ? parseFloat(((audited / total) * 100).toFixed(1)) : 0;
        return { name: month, audited, total, percent };
    });
  }, [data, availableMonths]);

  // NEW IDEA #2: Agent Trend Data (Based on Base Data: Result & Month)
  const selectedAgentTrendData = useMemo(() => {
    if (!activeCell.agent) return [];
    
    // Use availableMonths to ensure chronological order (assuming availableMonths is sorted correctly in base)
    return availableMonths.map(month => {
        const monthData = data.filter(d => d.agent === activeCell.agent && d.month === month);
        const total = monthData.length;
        const pass = monthData.filter(d => d.result.startsWith('ดีเยี่ยม') || d.result.startsWith('ผ่านเกณฑ์')).length;
        const rate = total > 0 ? ((pass / total) * 100).toFixed(1) : 0;
        
        return {
            name: month,
            total: total,
            passRate: parseFloat(rate),
            passCount: pass
        };
    }).filter(d => d.total > 0); // Show only months with data for cleaner chart
  }, [activeCell.agent, data, availableMonths]);


  const handleMatrixClick = (agentName, type) => {
    if (activeCell.agent === agentName && activeCell.resultType === type) setActiveCell({ agent: null, resultType: null });
    else { setActiveCell({ agent: agentName, resultType: type }); }
  };

  const handleToggleFilter = (item, selectedList, setSelectedFn) => {
    selectedList.includes(item) ? setSelectedFn(selectedList.filter(i => i !== item)) : setSelectedFn([...selectedList, item]);
  };

  const handleKpiClick = (filterType) => {
      if (activeKpiFilter === filterType) { setActiveKpiFilter(null); } 
      else { setActiveKpiFilter(filterType); setTimeout(() => { document.getElementById('detail-section')?.scrollIntoView({ behavior: 'smooth' }); }, 100); }
  };

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4 text-slate-800 font-sans relative overflow-hidden">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-indigo-200/30 blur-[120px] rounded-full"></div>
        <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-indigo-200/30 blur-[120px] rounded-full"></div>
        <div className="bg-white/90 backdrop-blur-xl p-8 rounded-[2rem] border border-slate-200 w-full max-w-[360px] text-center shadow-2xl relative z-10 animate-in fade-in zoom-in duration-500">
          <div className="flex justify-center mb-6"><div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 shadow-inner"><IntageLogo className="scale-110" /></div></div>
          <div className="space-y-1 mb-8"><h2 className="text-slate-800 font-black uppercase text-xs tracking-[0.3em] italic">CATI CES 2026</h2><p className="text-slate-500 text-[9px] font-bold uppercase tracking-widest">Analytics & Quality Control System</p></div>
          <form onSubmit={(e) => { e.preventDefault(); if(inputUser==='Admin'&&inputPass==='1234') { setIsAuthenticated(true); setUserRole('Admin'); } else if(inputUser==='QC'&&inputPass==='1234') { setIsAuthenticated(true); setUserRole('QC'); } else if(inputUser==='INV'&&inputPass==='1234') { setIsAuthenticated(true); setUserRole('INV'); } else { setLoginError('Login Failed'); } }} className="space-y-4 text-left">
            <div className="space-y-1.5"><label className="text-[9px] font-black text-slate-400 uppercase tracking-widest pl-2">Authorized Username</label><div className="relative"><User className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={16} /><input type="text" value={inputUser} onChange={e=>setInputUser(e.target.value)} className="w-full pl-11 pr-5 py-3 bg-white border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-indigo-600 focus:border-transparent text-slate-800 transition-all text-sm font-bold placeholder:text-slate-300" placeholder="Username" /></div></div>
            <div className="space-y-1.5"><label className="text-[9px] font-black text-slate-400 uppercase tracking-widest pl-2">Security Password</label><div className="relative"><Lock className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={16} /><input type="password" value={inputPass} onChange={e=>setInputPass(e.target.value)} className="w-full pl-11 pr-5 py-3 bg-white border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-indigo-600 focus:border-transparent text-slate-800 transition-all text-sm font-bold placeholder:text-slate-300" placeholder="••••••••" /></div></div>
            {loginError && (<div className="bg-rose-50 border border-rose-200 py-2 rounded-lg flex items-center justify-center gap-2"><AlertCircle size={12} className="text-rose-500" /><p className="text-rose-500 text-[9px] font-black uppercase tracking-widest">Authentication Failed</p></div>)}
            <button type="submit" className="w-full py-3.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-black uppercase tracking-[0.2em] transition-all shadow-lg shadow-indigo-900/20 flex items-center justify-center gap-2 active:scale-95 italic text-xs"><LogIn size={16} /> Access System</button>
          </form>
          <div className="mt-8 pt-6 border-t border-slate-100"><p className="text-[8px] text-slate-400 font-bold uppercase tracking-widest">© 2026 INTAGE (Thailand) Co., Ltd.</p></div>
        </div>
      </div>
    );
  }

  const FilterSection = ({ title, items, selectedItems, onToggle, onSelectAll, onClear, maxH = "max-h-40" }) => (
    <div className="space-y-2">
        <div className="flex items-center justify-between pl-2"><label className="text-[10px] font-black uppercase tracking-widest text-slate-400">{title}</label><div className="flex gap-2"><button onClick={onSelectAll} className="text-[9px] font-bold text-slate-400 hover:text-indigo-500 transition-colors">เลือกทั้งหมด</button><button onClick={onClear} className="text-[9px] font-bold text-slate-400 hover:text-indigo-500 transition-colors">ล้าง</button></div></div>
        <div className={`bg-slate-50 border border-slate-200 rounded-2xl p-2 overflow-y-auto custom-scrollbar ${maxH}`}>{items.map(item => (<div key={item} onClick={() => onToggle(item)} className={`flex items-center gap-2 p-2 rounded-xl cursor-pointer text-[10px] font-bold mb-1 transition-all ${selectedItems.includes(item) ? 'bg-indigo-50 text-indigo-700 shadow-sm border border-indigo-200' : 'hover:bg-slate-100 text-slate-500'}`}>{selectedItems.includes(item) ? <CheckSquare size={14} className="text-indigo-600 shrink-0" /> : <Square size={14} className="shrink-0" />}<span className="truncate">{formatResultDisplay(item)}</span></div>))}</div>
    </div>
  );

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 text-slate-800 font-sans custom-scrollbar">
      <div className={`fixed inset-0 z-40 bg-slate-900/20 backdrop-blur-sm transition-opacity duration-300 ${isFilterSidebarOpen ? 'opacity-100' : 'opacity-0 pointer-events-none'}`} onClick={() => setIsFilterSidebarOpen(false)} />
      <aside className={`fixed inset-y-0 right-0 z-50 w-80 bg-white shadow-2xl transform transition-transform duration-300 ${isFilterSidebarOpen ? 'translate-x-0' : 'translate-x-full'} overflow-y-auto border-l border-slate-200 p-6`}>
          <div className="flex items-center justify-between mb-8 pb-4 border-b border-slate-100"><div className="flex items-center gap-2"><div className="p-2 bg-indigo-600 rounded-lg text-white shadow-lg shadow-indigo-900/10"><Filter size={20} /></div><h3 className="font-black text-slate-800 uppercase italic tracking-tight">ตัวกรอง</h3></div><button onClick={() => setIsFilterSidebarOpen(false)} className="text-slate-400 hover:text-slate-600"><X size={20} /></button></div>
          <div className="space-y-8">
             <button onClick={() => { setSelectedMonths([]); setSelectedSups([]); setSelectedAgents([]); setSelectedResults([]); setSelectedTypes([]); setSelectedTouchpoints([]); setActiveCell({ agent: null, resultType: null }); setActiveKpiFilter(null); }} className="w-full py-2.5 text-xs font-black text-indigo-600 bg-indigo-50 hover:bg-indigo-100 rounded-xl border border-indigo-200 uppercase tracking-widest transition-all">ล้างตัวกรองทั้งหมด</button>
             <FilterSection title="เดือน" items={availableMonths} selectedItems={selectedMonths} onToggle={(item) => handleToggleFilter(item, selectedMonths, setSelectedMonths)} onSelectAll={() => setSelectedMonths(availableMonths)} onClear={() => setSelectedMonths([])} />
             <FilterSection title="Touchpoint (จุดสัมผัส)" items={availableTouchpoints} selectedItems={selectedTouchpoints} onToggle={(item) => handleToggleFilter(item, selectedTouchpoints, setSelectedTouchpoints)} onSelectAll={() => setSelectedTouchpoints(availableTouchpoints)} onClear={() => setSelectedTouchpoints([])} />
             <FilterSection title="Supervisor" items={availableSups} selectedItems={selectedSups} onToggle={(item) => handleToggleFilter(item, selectedSups, setSelectedSups)} onSelectAll={() => setSelectedSups(availableSups)} onClear={() => setSelectedSups([])} />
             <FilterSection title="ประเภทงาน (AC / BC)" items={availableTypes} selectedItems={selectedTypes} onToggle={(item) => handleToggleFilter(item, selectedTypes, setSelectedTypes)} onSelectAll={() => setSelectedTypes(availableTypes)} onClear={() => setSelectedTypes([])} />
             <FilterSection title="พนักงาน (ID : ชื่อ)" items={availableAgents} selectedItems={selectedAgents} onToggle={(item) => handleToggleFilter(item, selectedAgents, setSelectedAgents)} onSelectAll={() => setSelectedAgents(availableAgents)} onClear={() => setSelectedAgents([])} maxH="max-h-60" />
             <FilterSection title="ผลการสัมภาษณ์" items={RESULT_ORDER} selectedItems={selectedResults} onToggle={(item) => handleToggleFilter(item, selectedResults, setSelectedResults)} onSelectAll={() => setSelectedResults(RESULT_ORDER)} onClear={() => setSelectedResults([])} maxH="max-h-60" />
          </div>
       </aside>

      <div className="max-w-7xl mx-auto space-y-6">
        <header className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 bg-white p-6 rounded-[2.5rem] shadow-xl border border-slate-200">
          <div className="flex items-center gap-6"><div className="p-3 bg-slate-50 rounded-2xl border border-slate-100 shadow-inner"><IntageLogo /></div><div><h1 className="text-xl font-black text-slate-800 tracking-tight flex items-center gap-2 uppercase italic">QC REPORT 2026 {loading && <RefreshCw size={18} className="animate-spin text-indigo-500" />}</h1><div className="text-slate-400 text-[10px] font-black uppercase tracking-widest mt-1 flex items-center gap-2 italic">{data.length > 0 ? <><div className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-pulse"></div> REAL-TIME CONNECTED: {data.length} รายการ</> : "OFFLINE"} {lastUpdated && <span className="ml-4 opacity-50"><Clock size={10} className="inline mr-1" />{lastUpdated}</span>}</div></div></div>
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={() => fetchFromAppsScript(appsScriptUrl)} disabled={loading} className={`flex items-center gap-2 px-5 py-3 rounded-2xl text-xs font-black shadow-sm transition-all border ${loading ? 'bg-slate-100 border-slate-200 text-slate-400' : 'bg-slate-50 border-slate-200 hover:bg-slate-100 text-indigo-500 hover:text-indigo-600 shadow-indigo-200/50'}`}><RefreshCw size={16} className={loading ? 'animate-spin' : ''} />{loading ? 'SYNC DATA...' : 'SYNC DATA'}</button>
            <button onClick={() => setIsFilterSidebarOpen(true)} className={`flex items-center gap-2 px-5 py-3 rounded-2xl text-xs font-black shadow-sm transition-all border ${selectedResults.length > 0 || selectedSups.length > 0 || selectedMonths.length > 0 || selectedAgents.length > 0 || selectedTypes.length > 0 || selectedTouchpoints.length > 0 ? 'bg-indigo-600 border-indigo-500 text-white shadow-lg shadow-indigo-900/20' : 'bg-slate-50 border-slate-200 hover:bg-slate-100 text-slate-600'}`}><Filter size={16} /> ตัวกรอง</button>
            {userRole === 'Admin' && <button onClick={() => setShowSync(true)} className="flex items-center gap-2 px-5 py-3 bg-slate-800 text-white rounded-2xl text-xs font-black hover:bg-slate-700 transition-all shadow-xl font-bold"><Settings size={14} /> ตั้งค่า</button>}
            <button onClick={() => setIsAuthenticated(false)} className="p-3 bg-slate-50 rounded-2xl hover:text-indigo-500 text-slate-400 transition-colors border border-slate-200"><User size={20} /></button>
          </div>
        </header>

        <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
            {[
                { id: 'total', label: 'จำนวนงานทั้งหมด (กรองแค่เดือน)', value: totalWorkByMonthOnly, icon: FileText, color: 'text-slate-800', bg: 'bg-white border-slate-200', activeBg: 'bg-slate-100 ring-2 ring-slate-300' },
                { id: 'audited', label: 'จำนวนที่ตรวจแล้ว (AC/BC)', value: `${totalAuditedFiltered} (${totalWorkByMonthOnly > 0 ? ((totalAuditedFiltered / totalWorkByMonthOnly) * 100).toFixed(1) : 0}%)`, icon: Database, color: 'text-indigo-600', bg: 'bg-indigo-50 border-indigo-100', activeBg: 'bg-indigo-100 ring-2 ring-indigo-300' },
                { id: 'pass', label: 'จำนวนผ่านเกณฑ์', value: passCount, icon: CheckCircle, color: 'text-emerald-600', bg: 'bg-emerald-50 border-emerald-100', activeBg: 'bg-emerald-100 ring-2 ring-emerald-300' },
                { id: 'improve', label: 'ควรปรับปรุง', value: baseFilteredData.filter(d=>d.result.startsWith('ควรปรับปรุง')).length, icon: AlertTriangle, color: 'text-amber-500', bg: 'bg-amber-50 border-amber-100', activeBg: 'bg-amber-100 ring-2 ring-amber-300' },
                { id: 'error', label: 'พบข้อผิดพลาด', value: baseFilteredData.filter(d=>d.result.startsWith('พบข้อผิดพลาด')).length, icon: XCircle, color: 'text-rose-500', bg: 'bg-rose-50 border-rose-100', activeBg: 'bg-rose-100 ring-2 ring-rose-300' }
            ].map((kpi) => {
                const isActive = activeKpiFilter === kpi.id || (kpi.id === 'total' && activeKpiFilter === null);
                const isTotal = kpi.id === 'total';
                return (
                  <button key={kpi.id} onClick={() => handleKpiClick(isTotal ? null : kpi.id)} className={`text-left p-6 rounded-[2.5rem] border shadow-sm transition-all duration-200 active:scale-95 group relative overflow-hidden ${isActive && !isTotal ? kpi.activeBg : kpi.bg} ${!isActive ? 'hover:border-slate-300' : ''}`}>
                    {isActive && !isTotal && <div className="absolute top-3 right-4 text-[10px] font-black uppercase text-white bg-slate-800 px-2 py-0.5 rounded-full flex items-center gap-1"><MousePointerClick size={10}/> Filtering</div>}
                    <div className={`w-8 h-8 rounded-xl flex items-center justify-center mb-3 bg-white border border-slate-100 shadow-sm ${kpi.color}`}><kpi.icon size={16} /></div>
                    <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">{kpi.label}</p>
                    <h2 className={`text-3xl font-black ${kpi.color} tracking-tighter mt-1 uppercase`}>{kpi.value}</h2>
                  </button>
                );
            })}
        </div>

        <div className="flex flex-col lg:flex-row gap-6">
            <div className="bg-white p-8 rounded-[3rem] shadow-2xl border border-slate-200 flex-1 flex flex-col min-h-[300px]">
                <h3 className="font-black text-slate-800 flex items-center gap-2 italic text-sm uppercase mb-6"><PieChart size={16} className="text-indigo-500" /> Case Composition Summary</h3>
                <div className="flex-1 w-full min-h-[200px]">
                    <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                            <Pie data={chartData} dataKey="count" innerRadius={60} outerRadius={85} paddingAngle={5}>
                                {chartData.map((entry, index) => <Cell key={index} fill={entry.color} />)}
                            </Pie>
                            <Tooltip contentStyle={{backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '12px', fontSize: '10px', color: '#1e293b', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                        </PieChart>
                    </ResponsiveContainer>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mt-6">{chartData.map(c => (<div key={c.full} className="flex items-center gap-2"><div className="w-2 h-2 rounded-full shrink-0" style={{backgroundColor: c.color}}></div><span className="text-[9px] text-slate-500 font-bold truncate uppercase">{c.name}</span></div>))}</div>
            </div>

             <div className="bg-white p-8 rounded-[3rem] shadow-2xl border border-slate-200 flex-1 flex flex-col min-h-[300px]">
                <h3 className="font-black text-slate-800 flex items-center gap-2 italic text-sm uppercase mb-6"><BarChart2 size={16} className="text-emerald-500" /> Monthly Audit Progress (%)</h3>
                <div className="flex-1 w-full min-h-[200px]">
                    <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={monthlyPerformanceData}>
                            <defs><linearGradient id="barGradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#10B981" stopOpacity={1}/><stop offset="100%" stopColor="#059669" stopOpacity={0.6}/></linearGradient></defs>
                            <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                            <XAxis dataKey="name" stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                            <YAxis stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} unit="%" />
                            <Tooltip cursor={{fill: '#f1f5f9', opacity: 0.8}} content={({ active, payload, label }) => { if (active && payload && payload.length) { return (<div className="bg-white border border-slate-200 p-3 rounded-xl shadow-xl"><p className="text-slate-400 text-[10px] font-bold uppercase tracking-widest mb-1">{label}</p><p className="text-emerald-600 text-lg font-black">{payload[0].value}%</p><p className="text-slate-500 text-[10px] font-bold">Audited: {payload[0].payload.audited} / {payload[0].payload.total}</p></div>); } return null; }} />
                            <Bar dataKey="percent" fill="url(#barGradient)" radius={[6, 6, 0, 0]} barSize={40} animationDuration={1500}>{monthlyPerformanceData.map((entry, index) => (<Cell key={`cell-${index}`} fill={entry.percent >= 100 ? '#3b82f6' : 'url(#barGradient)'} />))}</Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>
            </div>
        </div>

        <div className="bg-white rounded-[3rem] shadow-2xl border border-slate-200 overflow-hidden">
            <div className="p-8 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
                <div><h3 className="font-black text-slate-800 flex items-center gap-2 italic text-lg uppercase tracking-tight"><Users size={20} className="text-indigo-500" /> สรุปพนักงาน (ID : ชื่อ) x ผลการตรวจ</h3></div>
            </div>
            <div className="overflow-x-auto max-h-[500px] custom-scrollbar">
                <table className="w-full text-left text-sm border-separate border-spacing-0 min-w-[1200px]">
                    <thead className="sticky top-0 bg-white z-20 font-black text-slate-700 text-[10px] uppercase tracking-widest border-b border-slate-200 shadow-sm">
                        <tr>
                            <th rowSpan="2" className="px-8 py-6 border-b border-slate-200 border-r border-slate-100 bg-white w-64 text-slate-700 italic">Interviewer (ID : Name)</th>
                            <th colSpan={RESULT_ORDER.length} className="px-4 py-4 text-center border-b border-slate-200 bg-slate-50 text-indigo-500 text-[11px] font-black italic uppercase">QC Result Distribution</th>
                            <th rowSpan="2" className="px-8 py-6 text-center bg-slate-100 text-slate-700 border-b border-slate-200 border-l border-slate-200">Total</th>
                        </tr>
                        <tr className="bg-slate-50 text-slate-600">
                            {RESULT_ORDER.map(type => <th key={type} className="px-4 py-3 text-center border-b border-slate-200 border-r border-slate-200 max-w-[180px] text-slate-500"><span className="line-clamp-2" title={type}>{formatResultDisplay(type)}</span></th>)}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-bold">
                        {agentSummary.map((agent, i) => (
                        <tr key={i} className="hover:bg-indigo-50 transition-colors group">
                            <td className="px-8 py-5 text-slate-700 border-r border-slate-100 font-medium">{agent.name}</td>
                            {RESULT_ORDER.map(type => {
                            const val = agent[type]; const isActive = activeCell.agent === agent.name && activeCell.resultType === type;
                            const percent = agent.total > 0 ? ((val / agent.total) * 100).toFixed(1) : 0;
                            return (
                                <td key={type} className={`px-4 py-5 text-center border-r border-slate-100 transition-all ${val > 0 ? 'cursor-pointer hover:bg-white shadow-sm' : ''} ${isActive ? 'bg-indigo-50 ring-2 ring-inset ring-indigo-200' : ''}`} onClick={() => val > 0 && handleMatrixClick(agent.name, type)}>
                                    <span className={`text-sm font-black ${val > 0 ? '' : 'text-slate-300'}`} style={{ color: val > 0 ? getResultColor(type) : undefined }}>
                                        {val > 0 ? (
                                            <div className="flex flex-col items-center">
                                                <span>{val}</span>
                                                <span className="text-[9px] opacity-60">({percent}%)</span>
                                            </div>
                                        ) : '-'}
                                    </span>
                                </td>
                            );
                            })}
                            <td className="px-8 py-5 text-center bg-slate-50 text-slate-700 border-l border-slate-200">
                                <div className="flex flex-col items-center">
                                    <span className="font-black text-slate-700">{agent.total}</span>
                                    <span className="text-[9px] text-slate-400 font-bold">({totalSummary.total > 0 ? ((agent.total / totalSummary.total) * 100).toFixed(1) : 0}%)</span>
                                </div>
                            </td>
                        </tr>
                        ))}
                        <tr className="bg-slate-100 text-slate-800 font-black border-t-2 border-slate-300 sticky bottom-0 z-20 shadow-lg">
                            <td className="px-8 py-5 border-r border-slate-300 text-indigo-600 italic uppercase">GRAND TOTAL</td>
                            {RESULT_ORDER.map(type => {
                                const val = totalSummary[type];
                                const percent = totalSummary.total > 0 ? ((val / totalSummary.total) * 100).toFixed(1) : 0;
                                return (
                                    <td key={type} className="px-4 py-5 text-center border-r border-slate-300">
                                        <div className="flex flex-col items-center"><span>{val}</span><span className="text-[9px] text-slate-500">({percent}%)</span></div>
                                    </td>
                                );
                            })}
                            <td className="px-8 py-5 text-center border-l border-slate-300 text-indigo-600 bg-slate-200"><div className="flex flex-col items-center"><span>{totalSummary.total}</span><span className="text-[9px] opacity-60">(100%)</span></div></td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        {/* --- NEW IDEA 2: Agent Performance Trend Chart (Shows when Agent is selected) --- */}
        {activeCell.agent && selectedAgentTrendData.length > 0 && (
            <div className="bg-white p-8 rounded-[3rem] shadow-2xl border border-slate-200 animate-in slide-in-from-top-4 duration-500 scroll-mt-6">
                 <div className="flex items-center justify-between mb-6">
                    <h3 className="font-black text-slate-800 flex items-center gap-3 italic text-lg uppercase tracking-tight">
                        <TrendingUp size={24} className="text-indigo-500" /> 
                        Performance Trend: <span className="text-indigo-600 border-b-2 border-indigo-200">{activeCell.agent}</span>
                    </h3>
                    <div className="flex items-center gap-2">
                        <div className="flex items-center gap-1.5 bg-indigo-50 px-3 py-1 rounded-full"><div className="w-3 h-3 bg-indigo-500 rounded-sm"></div><span className="text-[10px] font-black text-indigo-800 uppercase">Total Cases</span></div>
                        <div className="flex items-center gap-1.5 bg-emerald-50 px-3 py-1 rounded-full"><div className="w-3 h-3 bg-emerald-500 rounded-full"></div><span className="text-[10px] font-black text-emerald-800 uppercase">Pass Rate %</span></div>
                    </div>
                 </div>
                 <div className="w-full h-[300px]">
                    <ResponsiveContainer width="100%" height="100%">
                        <ComposedChart data={selectedAgentTrendData} margin={{top: 20, right: 20, bottom: 20, left: 20}}>
                            <CartesianGrid stroke="#f1f5f9" vertical={false} strokeDasharray="3 3"/>
                            <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fill: '#64748b', fontSize: 11, fontWeight: 'bold'}} />
                            <YAxis yAxisId="left" axisLine={false} tickLine={false} tick={{fill: '#64748b', fontSize: 11}} label={{ value: 'Total Cases', angle: -90, position: 'insideLeft', style: {textAnchor: 'middle', fill: '#cbd5e1', fontSize: 10, fontWeight: 'bold'} }} />
                            <YAxis yAxisId="right" orientation="right" axisLine={false} tickLine={false} tick={{fill: '#10B981', fontSize: 11, fontWeight: 'bold'}} unit="%" domain={[0, 100]} />
                            <Tooltip contentStyle={{backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '12px', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} cursor={{fill: '#f8fafc'}} />
                            <Bar yAxisId="left" dataKey="total" name="จำนวนตรวจ (เคส)" fill="#818cf8" radius={[6, 6, 0, 0]} barSize={40} fillOpacity={0.8} />
                            <Line yAxisId="right" type="monotone" dataKey="passRate" name="% ผ่านเกณฑ์" stroke="#10B981" strokeWidth={3} dot={{r: 4, fill: '#10B981', strokeWidth: 2, stroke: '#fff'}} activeDot={{r: 6}} />
                        </ComposedChart>
                    </ResponsiveContainer>
                 </div>
            </div>
        )}

        <div id="detail-section" className="bg-white rounded-[3rem] shadow-2xl border border-slate-200 overflow-hidden scroll-mt-6">
            <div className="p-8 border-b border-slate-100 flex flex-col md:flex-row items-center justify-between gap-6">
                <div className="space-y-1">
                    <h3 className="font-black text-slate-800 uppercase tracking-widest text-xs flex items-center gap-2 italic"><MessageSquare size={16} className="text-indigo-500" /> รายละเอียดรายเคส</h3>
                    <div className="flex flex-wrap items-center gap-2 mt-2">
                        {activeCell.agent && <span className="px-3 py-1 bg-indigo-600 text-white text-[9px] font-black rounded-lg uppercase italic animate-pulse flex items-center gap-2">กำลังแสดง: {activeCell.agent} <button onClick={() => setActiveCell({ agent: null, resultType: null })} className="hover:text-slate-200"><X size={10}/></button></span>}
                        {activeKpiFilter && <span className="px-3 py-1 bg-emerald-600 text-white text-[9px] font-black rounded-lg uppercase italic animate-pulse flex items-center gap-2">Filter: {activeKpiFilter} <button onClick={() => setActiveKpiFilter(null)} className="hover:text-slate-200"><X size={10}/></button></span>}
                    </div>
                </div>
                <div className="relative w-full md:w-80"><Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={16} /><input type="text" placeholder="ค้นหาพนักงาน หรือ เลขชุด..." className="w-full pl-12 pr-6 py-4 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold focus:ring-2 focus:ring-indigo-600 outline-none text-slate-800 shadow-inner placeholder:text-slate-400" value={searchTerm} onChange={(e)=>setSearchTerm(e.target.value)} /></div>
            </div>
            <div className="overflow-auto max-h-[1000px] custom-scrollbar">
                <table className="w-full text-left text-xs font-medium border-separate border-spacing-0">
                <thead className="sticky top-0 bg-white shadow-sm z-10 border-b border-slate-200 font-black text-slate-500 uppercase tracking-widest">
                    <tr><th className="px-8 py-5 border-r border-slate-100">วันที่ / เลขชุด / TOUCH_POINT</th><th className="px-8 py-5 border-r border-slate-100">พนักงาน (ID : ชื่อ)</th><th className="px-4 py-5 text-center border-r border-slate-100">ผลสรุป</th><th className="px-8 py-5">QC Full Comment (Column N) & Audio</th></tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                    {detailLogs.length > 0 ? detailLogs.slice(0, 150).map((item) => {
                        const isExpanded = expandedCaseId === item.id; const isEditing = editingCase && editingCase.id === item.id;
                        const isNewAudit = item.type === "ยังไม่ได้ตรวจ";

                        return (
                        <React.Fragment key={item.id}>
                            <tr onClick={() => !isEditing && setExpandedCaseId(isExpanded ? null : item.id)} className={`transition-all group cursor-pointer ${isExpanded ? 'bg-indigo-50 shadow-inner' : 'hover:bg-slate-50'}`}>
                                <td className="px-8 py-6 border-r border-slate-100">
                                    <div className="font-black text-slate-800">{item.date}</div>
                                    <div className="flex flex-wrap gap-2 mt-1">
                                        <div className="flex items-center gap-1.5 bg-white px-2 py-0.5 rounded border border-slate-200 w-fit"><Hash size={10} className="text-indigo-400" /><span className="text-[11px] font-black text-slate-500">{item.questionnaireNo}</span></div>
                                        <div className="flex items-center gap-1.5 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-100 w-fit"><MapPin size={10} className="text-indigo-500" /><span className="text-[10px] font-black text-indigo-600">{item.touchpoint}</span></div>
                                    </div>
                                </td>
                                <td className="px-8 py-6 border-r border-slate-100"><div className="font-black text-slate-700 text-sm group-hover:text-indigo-600 transition-colors flex items-center gap-2">{item.agent} {isExpanded ? <ChevronDown size={14} className="text-slate-400" /> : <ChevronRight size={14} className="text-slate-400" />}</div><div className="text-[9px] text-slate-400 font-bold mt-0.5 italic uppercase font-sans tracking-wider">TYPE: {item.type} &bull; SUP: {item.supervisorFilter}</div></td>
                                <td className="px-4 py-6 text-center border-r border-slate-100"><span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[9px] font-black border uppercase shadow-sm" style={{ backgroundColor: `${getResultColor(item.result)}10`, color: getResultColor(item.result), borderColor: `${getResultColor(item.result)}30` }}>{formatResultDisplay(item.result)}</span></td>
                                <td className="px-8 py-6">
                                    <p className="text-slate-500 italic max-w-sm truncate group-hover:text-slate-700 transition-colors font-sans leading-relaxed">{item.comment ? `"${item.comment}"` : '-'}</p>
                                    {item.audio && item.audio.includes('http') && (<a href={item.audio} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()} className="mt-2 inline-flex items-center gap-1 text-indigo-500 hover:text-indigo-600 font-black text-[9px] uppercase transition-all hover:translate-x-1"><PlayCircle size={14} /> Listen Recording</a>)}
                                </td>
                            </tr>
                            {isExpanded && (
                            <tr className="bg-slate-50 animate-in slide-in-from-top-2 duration-300">
                                <td colSpan={4} className="p-8 border-b border-slate-200 text-slate-800">
                                <div className="flex items-center justify-between mb-8">
                                    <div className="flex items-center gap-4"><div className="p-3 bg-white border border-slate-200 rounded-2xl text-indigo-600 shadow-sm"><Award /></div><div><h4 className="font-black uppercase italic tracking-widest text-sm text-slate-700">{isNewAudit ? "START AUDIT SESSION" : "ASSESSMENT DETAIL"} (ID: {item.interviewerId} : {item.rawName})</h4><p className="text-[9px] text-slate-400 font-bold uppercase tracking-widest italic leading-relaxed">{isNewAudit ? "กรุณากรอกคะแนนและผลสรุปเพื่อบันทึกงานใหม่" : "แก้ไขคะแนน P:AB และ สรุปผล M"}</p></div></div>
                                    <div className="flex gap-2">
                                    {userRole === 'Admin' || userRole === 'QC' ? (
                                        !isEditing ? (
                                            <button onClick={() => setEditingCase({...item})} className={`flex items-center gap-2 px-6 py-3 rounded-2xl text-[10px] font-black uppercase border transition-all ${isNewAudit ? 'bg-indigo-600 hover:bg-indigo-700 text-white border-indigo-600 shadow-lg shadow-indigo-900/20' : 'bg-white hover:bg-slate-50 text-slate-600 border-slate-200'}`}>
                                                <Edit2 size={12} className={isNewAudit ? "text-white" : "text-indigo-500"}/> 
                                                {isNewAudit ? "เริ่มตรวจงานนี้" : "แก้ไขข้อมูล"}
                                            </button>
                                        ) : (
                                            <div className="flex gap-2"><button disabled={isSaving} onClick={handleUpdateCase} className="flex items-center gap-2 px-6 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-2xl text-[10px] font-black uppercase shadow-lg shadow-indigo-900/20">{isSaving ? <RefreshCw className="animate-spin" size={14}/> : <Save size={14}/>} {isNewAudit ? "บันทึกผลการตรวจ" : "บันทึกการแก้ไข"}</button><button onClick={() => setEditingCase(null)} className="px-6 py-3 bg-white text-slate-600 rounded-2xl text-[10px] font-black uppercase transition-all border border-slate-200 hover:bg-slate-50">ยกเลิก</button></div>
                                        )
                                    ) : (
                                        <div className="px-4 py-2 border border-slate-200 rounded-xl text-slate-400 text-[10px] font-black uppercase tracking-widest italic opacity-50 select-none">Read Only View</div>
                                    )}
                                    </div>
                                </div>

                                {isEditing && (
                                    <div className="mb-8 p-6 bg-indigo-50 border border-indigo-100 rounded-[2rem] shadow-inner">
                                        <div className="flex items-center gap-2 mb-4 text-indigo-600 font-black text-[10px] uppercase italic tracking-widest"><Info size={16} /> กำหนดประเภทงาน (AC / BC) & Supervisor (H)</div>
                                        <div className="flex flex-col md:flex-row gap-4">
                                            <div className="flex gap-2">
                                                {['AC', 'BC'].map(t => (
                                                    <button key={t} onClick={() => setEditingCase({...editingCase, type: t})} className={`px-6 py-3 rounded-xl text-[10px] font-black uppercase transition-all border ${editingCase.type === t ? 'bg-indigo-600 text-white border-indigo-600 shadow-lg shadow-indigo-900/40' : 'bg-white text-slate-500 border-slate-200 hover:border-slate-300'}`}>{t} MODE</button>
                                                ))}
                                            </div>
                                            <div className="flex-1 bg-white border border-slate-200 rounded-xl p-2 flex items-center gap-3 pl-4">
                                                <UserPlus size={16} className="text-indigo-500"/>
                                                <div className="flex-1 relative">
                                                    <p className="text-[9px] text-slate-400 font-bold uppercase tracking-wider mb-0.5">CATI Supervisor (Column H)</p>
                                                    <select value={editingCase.supervisor || ''} onChange={e=>setEditingCase({...editingCase, supervisor: e.target.value})} className="w-full bg-transparent text-slate-800 text-xs font-bold outline-none appearance-none"><option value="" className="text-slate-400">ระบุชื่อ Supervisor...</option>{SUPERVISOR_OPTIONS.map(opt => (<option key={opt} value={opt} className="text-slate-800">{opt}</option>))}</select><ChevronDown size={12} className="absolute right-0 top-1/2 translate-y-0 text-slate-400 pointer-events-none" />
                                                </div>
                                            </div>
                                            {editingCase.type === "ยังไม่ได้ตรวจ" && <p className="text-rose-500 text-[10px] font-black uppercase self-center animate-pulse">*** กรุณาเลือก AC หรือ BC เพื่อบันทึกงาน</p>}
                                        </div>
                                    </div>
                                )}

                                {item.audio && item.audio.includes('http') && (
                                    <div className="mb-8 p-6 bg-white border border-slate-200 rounded-[2rem] flex items-center justify-between shadow-sm">
                                        <div className="flex items-center gap-3"><div className="p-3 bg-indigo-50 text-indigo-600 rounded-2xl animate-pulse"><PlayCircle size={20} /></div><div><p className="text-[10px] font-black uppercase text-slate-400 tracking-widest">Audio Recording File</p><p className="text-xs font-black text-slate-700">ไฟล์เสียงบันทึกการสัมภาษณ์</p></div></div>
                                        <a href={item.audio} target="_blank" rel="noopener noreferrer" className="px-6 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-2xl text-[10px] font-black uppercase flex items-center gap-2 transition-all shadow-lg shadow-indigo-900/20"><ExternalLink size={14} /> OPEN PLAYER</a>
                                    </div>
                                )}
                                <div className="mb-8 p-6 bg-slate-100 border border-slate-200 rounded-[2rem] shadow-inner">
                                    <div className="flex items-center gap-2 mb-4 text-indigo-500 font-black text-[10px] uppercase italic tracking-widest"><Star size={16} /> สรุปผลการสัมภาษณ์แบบละเอียด (Column M)</div>
                                    {isEditing ? (<div className="relative"><select className="w-full p-4 bg-white border border-indigo-200 rounded-2xl text-[11px] font-black text-slate-800 focus:ring-2 focus:ring-indigo-600 outline-none appearance-none shadow-sm" value={editingCase.result} onChange={e => setEditingCase({...editingCase, result: e.target.value})}>{RESULT_ORDER.map(opt => <option key={opt} value={opt}>{opt}</option>)}</select><ChevronDown size={14} className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" /></div>) : (<p className="text-sm font-black italic text-slate-600 bg-white p-4 rounded-xl border border-slate-200 leading-relaxed shadow-sm">{item.result}</p>)}
                                </div>
                                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3">
                                    {(isEditing ? editingCase : item).evaluations.map((evalItem, eIdx) => (<div key={eIdx} className={`bg-white border p-3 rounded-2xl transition-all ${isEditing ? 'border-slate-200 opacity-60 cursor-not-allowed' : 'border-slate-200 shadow-sm'}`}><p className="text-[9px] font-black text-slate-400 uppercase tracking-tighter mb-2 truncate" title={evalItem.label}>{evalItem.label}</p><div className={`text-sm font-black italic tracking-widest ${evalItem.value === '5' || evalItem.value === '4' ? 'text-emerald-500' : (evalItem.value === '1' || evalItem.value === '2') ? 'text-rose-500' : 'text-slate-300'}`}>{SCORE_LABELS[evalItem.value] || evalItem.value}</div></div>))}
                                </div>
                                <div className="mt-8">
                                    <p className="text-[10px] font-black text-indigo-500 uppercase mb-2 italic tracking-widest flex items-center gap-2"><MessageSquare size={12}/> QC Full Comment (Column N)</p>
                                    {isEditing ? (<textarea className="w-full bg-white border border-slate-200 rounded-[1.5rem] p-6 text-sm italic text-slate-700 outline-none min-h-[120px] focus:ring-2 focus:ring-indigo-600 shadow-inner font-sans" value={editingCase.comment} onChange={e => setEditingCase({...editingCase, comment: e.target.value})}/>) : (<div className="p-4 bg-white rounded-2xl border border-slate-200 border-l-4 border-l-indigo-500 shadow-sm"><p className="text-sm text-slate-600 font-medium italic leading-relaxed font-sans">"{item.comment || 'ไม่มีคอมเมนต์'}"</p></div>)}
                                </div>
                                </td>
                            </tr>
                            )}
                        </React.Fragment>
                        );
                    }) : (<tr><td colSpan={4} className="px-8 py-24 text-center text-slate-400 font-black uppercase italic tracking-widest text-lg opacity-40 italic">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</td></tr>)}
                </tbody>
                </table>
            </div>
        </div>
      </div>

      {(showSync || error) && (
        <div className="fixed inset-0 z-[100] bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-white w-full max-w-2xl rounded-[3rem] p-10 border border-slate-200 shadow-2xl relative animate-in zoom-in duration-300">
                <div className="absolute top-0 left-0 w-full h-1 bg-indigo-600"></div>
                <div className="flex items-center justify-between mb-10"><h3 className="text-xl font-black flex items-center gap-3 text-slate-800 uppercase italic tracking-tight">{error ? 'Connection Error' : 'ตั้งค่าระบบ'}</h3>{data.length > 0 && <button onClick={() => {setShowSync(false); setError(null);}} className="text-slate-400 hover:text-slate-600 transition-colors"><X size={28} /></button>}</div>
                <div className="space-y-6">
                    <div className="p-4 bg-indigo-50 border border-indigo-100 rounded-2xl mb-4"><p className="text-xs text-indigo-600 font-bold mb-1 flex items-center gap-2"><Zap size={14}/> SYSTEM UPGRADE (JSON API MODE)</p><p className="text-[10px] text-slate-500">ระบบได้รับการอัปเกรดเป็นโหมดประสิทธิภาพสูง กรุณาใส่ Web App URL ของ Apps Script ลงในช่องด้านล่างเพียงช่องเดียว</p></div>
                    <div className="space-y-2"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-2 italic">Apps Script Web App URL (Read & Write)</label><input type="text" className="w-full px-6 py-4 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-bold text-slate-800 outline-none focus:ring-2 focus:ring-indigo-600 shadow-inner" value={appsScriptUrl} onChange={e=>{setAppsScriptUrl(e.target.value); localStorage.setItem('apps_script_url', e.target.value);}} /></div>
                    <div className="flex gap-4 pt-4"><button onClick={() => fetchFromAppsScript(appsScriptUrl)} disabled={loading || !appsScriptUrl} className="flex-1 py-5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-[2.5rem] font-black uppercase shadow-xl shadow-indigo-200 text-sm tracking-widest italic flex items-center justify-center gap-2 transition-all">{loading ? <RefreshCw className="animate-spin" size={18}/> : <RefreshCw size={18}/>} CONNECT & SYNC</button></div>
                    {error && <p className="text-center text-[10px] text-rose-500 font-black mt-2 uppercase italic animate-pulse">{error}</p>}
                </div>
            </div>
        </div>
      )}
    </div>
  );
};

export default App;