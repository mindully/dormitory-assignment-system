import React, { useState, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import * as xlsx from 'xlsx';
import {
  Upload,
  CheckCircle2,
  AlertTriangle,
  FileSpreadsheet,
  Users,
  Download,
  Search,
  Building2,
  RefreshCw,
  TrendingUp,
  UserCheck,
  UserX,
  Sparkles,
  Info,
  ChevronRight,
  HelpCircle
} from 'lucide-react';
import {
  generateAndDownloadSampleStudents,
  generateAndDownloadSampleRooms
} from './utils/sampleData';
import {
  Student,
  Room,
  AssignmentResult,
  FailedAssignment,
  AssignmentDashboardStats,
  AssignmentResponse,
  Diagnostics
} from './types';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  PieChart,
  Pie,
  Cell
} from 'recharts';

export default function App() {
  // File upload states
  const [studentFileName, setStudentFileName] = useState<string>('');
  const [studentBase64, setStudentBase64] = useState<string>('');
  const [studentCount, setStudentCount] = useState<number | null>(null);
  const [studentDragActive, setStudentDragActive] = useState<boolean>(false);

  const [roomFileName, setRoomFileName] = useState<string>('');
  const [roomBase64, setRoomBase64] = useState<string>('');
  const [roomCount, setRoomCount] = useState<number | null>(null);
  const [roomDragActive, setRoomDragActive] = useState<boolean>(false);

  // App processing states
  const [loading, setLoading] = useState<boolean>(false);
  const [loadingStep, setLoadingStep] = useState<string>('');
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  // Result data states
  const [assignedList, setAssignedList] = useState<AssignmentResult[]>([]);
  const [failedList, setFailedList] = useState<FailedAssignment[]>([]);
  const [stats, setStats] = useState<AssignmentDashboardStats | null>(null);
  const [diagnostics, setDiagnostics] = useState<Diagnostics | null>(null);
  const [showDiagnostics, setShowDiagnostics] = useState<boolean>(false);

  // Search & Filter state
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'assigned' | 'failed'>('all');
  const [resultsTab, setResultsTab] = useState<'dashboard' | 'records'>('dashboard');

  // Input references for trigger clicks
  const studentInputRef = useRef<HTMLInputElement>(null);
  const roomInputRef = useRef<HTMLInputElement>(null);

  // Toast helper
  const showToast = (msg: string, type: 'success' | 'error') => {
    if (type === 'success') {
      setSuccessMsg(msg);
      setTimeout(() => setSuccessMsg(null), 4000);
    } else {
      setErrorMsg(msg);
      setTimeout(() => setErrorMsg(null), 5000);
    }
  };

  // Local helper to read file as Base64 and inspect row counts
  const handleFileRead = (file: File, type: 'student' | 'room') => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const result = e.target?.result;
        if (!result) return;

        // Extract base64
        let base64Str = '';
        if (typeof result === 'string') {
          base64Str = result.split(',')[1] || result;
        }

        // Convert base64 to binary to inspect row count locally for feedback
        const binaryString = atob(base64Str);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        
        const workbook = xlsx.read(bytes, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const rawRows = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        if (type === 'student') {
          setStudentFileName(file.name);
          setStudentBase64(base64Str);
          setStudentCount(rawRows.length);
          showToast(`설문 데이터 ${rawRows.length}건 확인 완료`, 'success');
        } else {
          setRoomFileName(file.name);
          setRoomBase64(base64Str);
          setRoomCount(rawRows.length);
          showToast(`기숙사 방 현황 데이터 ${rawRows.length}건 확인 완료`, 'success');
        }
      } catch (err) {
        showToast('엑셀 파일을 읽는 도중 오류가 발생했습니다. 올바른 Excel 통합 문서(.xlsx) 형식인지 확인해 주세요.', 'error');
      }
    };
    reader.readAsDataURL(file);
  };

  // Drag and drop handlers
  const handleDrag = (e: React.DragEvent, type: 'student' | 'room', active: boolean) => {
    e.preventDefault();
    e.stopPropagation();
    if (type === 'student') {
      setStudentDragActive(active);
    } else {
      setRoomDragActive(active);
    }
  };

  const handleDrop = (e: React.DragEvent, type: 'student' | 'room') => {
    e.preventDefault();
    e.stopPropagation();
    if (type === 'student') {
      setStudentDragActive(false);
    } else {
      setRoomDragActive(false);
    }

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        showToast('엑셀 파일(.xlsx, .xls)만 업로드할 수 있습니다.', 'error');
        return;
      }
      handleFileRead(file, type);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>, type: 'student' | 'room') => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      handleFileRead(file, type);
    }
  };

  // Execute algorithm by sending raw Base64 data to our Node.js Express server
  const runAutoAssignment = async () => {
    if (!studentBase64 || !roomBase64) {
      showToast('설문 결과 엑셀과 기숙사 현황 엑셀을 모두 업로드해 주세요.', 'error');
      return;
    }

    setLoading(true);
    setStats(null);
    setAssignedList([]);
    setFailedList([]);

    // Staggered status messages to represent algorithm progression
    const steps = [
      '데이터 파싱 및 데이터 품질 검수 중...',
      '학생별 및 방별 성별 필터링 구조 분할 중...',
      '학번 기반 단방향/양방향 룸메이트 그래프 분석 및 결합 그룹 생성 중...',
      '입주일자(Move-in Date) 오름차순 우선순위 대기열 정렬 중...',
      '동일 조건 대상자 랜덤 가중치 배치 난수 배정 중...',
      '선호 기숙사 잔여 정원 실시간 매칭 계산 중...',
      '결과 데이터 정리 및 대시보드 인덱싱 빌드 중...'
    ];

    let currentStepIndex = 0;
    setLoadingStep(steps[currentStepIndex]);

    const stepInterval = setInterval(() => {
      if (currentStepIndex < steps.length - 1) {
        currentStepIndex++;
        setLoadingStep(steps[currentStepIndex]);
      }
    }, 600);

    try {
      const response = await fetch('/api/assign', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          studentFile: studentBase64,
          roomFile: roomBase64
        })
      });

      clearInterval(stepInterval);

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || '배정 프로세스 중 알 수 없는 오류가 발생했습니다.');
      }

      const resData: AssignmentResponse = await response.json();
      
      setStats(resData.stats);
      setAssignedList(resData.assigned_list);
      setFailedList(resData.failed_list);
      setDiagnostics(resData.diagnostics);
      setShowDiagnostics(true);
      setResultsTab('dashboard');
      showToast(`기숙사 자동 배정이 완료되었습니다! (성공률: ${resData.stats.success_rate}%)`, 'success');

    } catch (err: any) {
      clearInterval(stepInterval);
      showToast(err.message || '배정 서버 통신 실패', 'error');
    } finally {
      setLoading(false);
    }
  };

  // Trigger Excel file download by posting JSON list back to Express spreadsheet writer
  const downloadExcelResult = async () => {
    if (assignedList.length === 0 && failedList.length === 0) {
      showToast('배정 결과가 없습니다. 먼저 배정을 실행해 주세요.', 'error');
      return;
    }

    try {
      const response = await fetch('/api/download', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          assigned_list: assignedList,
          failed_list: failedList
        })
      });

      if (!response.ok) {
        throw new Error('엑셀 다운로드 파일 생성에 실패했습니다.');
      }

      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `기숙사_자동배정_최종결과_${new Date().toISOString().slice(0,10)}.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(downloadUrl);
      showToast('멀티탭 결과 엑셀 파일을 다운로드했습니다.', 'success');
    } catch (err: any) {
      showToast(err.message || '엑셀 다운로드 오류', 'error');
    }
  };

  // Reset current state
  const handleReset = () => {
    setStudentFileName('');
    setStudentBase64('');
    setStudentCount(null);
    setRoomFileName('');
    setRoomBase64('');
    setRoomCount(null);
    setAssignedList([]);
    setFailedList([]);
    setStats(null);
    setSearchQuery('');
    setStatusFilter('all');
    setResultsTab('dashboard');
    showToast('데이터와 상태가 정상적으로 초기화되었습니다.', 'success');
  };

  // Search/Filter function for student lists
  const filteredRecords = () => {
    const list: { id: string; name: string; gender: string; status: '성공' | '실패'; detail: string; movein: string; pref: string; roommateReq: string; roommateMatched?: string }[] = [];
    
    assignedList.forEach(item => {
      list.push({
        id: item.student_id,
        name: item.student_name,
        gender: item.gender,
        status: '성공',
        detail: `${item.dorm_type} / ${item.room_no}호`,
        movein: item.movein_date,
        pref: item.dorm_type,
        roommateReq: item.roommate_id_requested || '없음',
        roommateMatched: item.roommate_matched_id
      });
    });

    failedList.forEach(item => {
      list.push({
        id: item.student_id,
        name: item.student_name,
        gender: item.gender,
        status: '실패',
        detail: item.reason,
        movein: item.movein_date,
        pref: item.dorm_pref,
        roommateReq: item.roommate_id || '없음'
      });
    });

    return list.filter(item => {
      // Filter by status
      if (statusFilter === 'assigned' && item.status !== '성공') return false;
      if (statusFilter === 'failed' && item.status !== '실패') return false;

      // Filter by query
      if (searchQuery.trim() !== '') {
        const q = searchQuery.toLowerCase();
        return (
          item.id.toLowerCase().includes(q) ||
          item.name.toLowerCase().includes(q) ||
          item.detail.toLowerCase().includes(q) ||
          item.movein.toLowerCase().includes(q) ||
          item.pref.toLowerCase().includes(q)
        );
      }
      return true;
    });
  };

  // Constants for Pie chart colors
  const COLORS = ['#ef4444', '#f97316', '#f59e0b', '#3b82f6', '#10b981', '#8b5cf6'];

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans" id="dorm-app-root">
      
      {/* Toast notifications */}
      <AnimatePresence>
        {successMsg && (
          <motion.div
            initial={{ opacity: 0, y: -50 }}
            animate={{ opacity: 1, y: 16 }}
            exit={{ opacity: 0, y: -50 }}
            className="fixed top-4 left-1/2 -translate-x-1/2 z-50 bg-emerald-500 text-white px-6 py-3 rounded-xl shadow-lg flex items-center gap-3 font-medium text-sm border border-emerald-400"
            id="toast-success"
          >
            <CheckCircle2 className="w-5 h-5" />
            <span>{successMsg}</span>
          </motion.div>
        )}
        {errorMsg && (
          <motion.div
            initial={{ opacity: 0, y: -50 }}
            animate={{ opacity: 1, y: 16 }}
            exit={{ opacity: 0, y: -50 }}
            className="fixed top-4 left-1/2 -translate-x-1/2 z-50 bg-rose-500 text-white px-6 py-3 rounded-xl shadow-lg flex items-center gap-3 font-medium text-sm border border-rose-400"
            id="toast-error"
          >
            <AlertTriangle className="w-5 h-5" />
            <span>{errorMsg}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Header Bar */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-40 px-6 py-4 shadow-sm" id="main-header">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2.5 rounded-xl text-white shadow-md shadow-blue-200">
              <Building2 className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-bold font-display text-slate-900 tracking-tight flex items-center gap-2">
                외국인 유학생을 위한 기숙사 배치 시스템
                <span className="text-xs bg-blue-100 text-blue-700 font-mono px-2 py-0.5 rounded-full font-semibold">MVP</span>
              </h1>
              <p className="text-xs text-slate-500 mt-0.5">설문조사 및 호실 데이터를 활용한 알고리즘 기반 자동 배정 도구</p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={generateAndDownloadSampleStudents}
              className="bg-white hover:bg-slate-50 border border-slate-200 text-slate-700 text-xs px-3 py-2 rounded-lg flex items-center gap-1.5 transition-colors font-medium shadow-xs"
              id="download-sample-students"
              title="유학생 설문 데이터 200명 무작위 생성"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span>샘플 설문 결과 (200명)</span>
            </button>
            <button
              onClick={generateAndDownloadSampleRooms}
              className="bg-white hover:bg-slate-50 border border-slate-200 text-slate-700 text-xs px-3 py-2 rounded-lg flex items-center gap-1.5 transition-colors font-medium shadow-xs"
              id="download-sample-rooms"
              title="기숙사 호실 현황 정보(정원 포함) 생성"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span>샘플 기숙사 현황</span>
            </button>
            {(studentBase64 || roomBase64 || stats) && (
              <button
                onClick={handleReset}
                className="bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs px-3 py-2 rounded-lg flex items-center gap-1.5 transition-colors font-medium"
                id="reset-all"
              >
                <RefreshCw className="w-3.5 h-3.5 text-slate-500" />
                <span>초기화</span>
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-6 space-y-6" id="main-content">
        
        {/* Banner with guides and explanation of constraints */}
        <section className="bg-blue-50 border border-blue-100 rounded-2xl p-5 flex flex-col md:flex-row gap-4 items-start" id="intro-guide">
          <div className="bg-blue-100 p-2 rounded-lg text-blue-600 shrink-0">
            <Info className="w-5 h-5" />
          </div>
          <div className="space-y-1">
            <h3 className="font-semibold text-slate-900 text-sm md:text-base">기숙사 자동 배정 알고리즘 작동 가이드</h3>
            <p className="text-xs text-slate-600 leading-relaxed">
              본 시스템은 개인 정보 보호를 위해 데이터를 파일 형태로만 처리하며 데이터베이스에 영구 저장하지 않습니다.
              상단의 <strong>[샘플 결과 다운로드]</strong>를 통해 테스트 파일을 즉시 준비하여 업로드해보세요!
            </p>
            <ul className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2 text-[11px] text-slate-500 pt-2 font-mono">
              <li className="flex items-center gap-1">
                <span className="text-blue-500 font-bold">•</span>
                성별 필터 자동 분리 매칭
              </li>
              <li className="flex items-center gap-1">
                <span className="text-blue-500 font-bold">•</span>
                룸메이트 상호/단방향 우선그룹화
              </li>
              <li className="flex items-center gap-1">
                <span className="text-blue-500 font-bold">•</span>
                입사예정일 순 우선 대기열 배치
              </li>
              <li className="flex items-center gap-1">
                <span className="text-blue-500 font-bold">•</span>
                동률 발생 시 무작위로 우선순위 결정
              </li>
            </ul>
          </div>
        </section>

        {/* Upload Zone Panel */}
        <section className="grid grid-cols-1 md:grid-cols-2 gap-6" id="upload-section">
          {/* Box 1: Student Survey Results */}
          <div
            className={`bg-white rounded-2xl border-2 p-6 transition-all relative ${
              studentDragActive ? 'border-blue-500 bg-blue-50/20' : 'border-slate-200'
            } ${studentBase64 ? 'border-dashed border-emerald-500/50 bg-emerald-50/5' : 'border-dashed'}`}
            onDragOver={(e) => handleDrag(e, 'student', true)}
            onDragLeave={(e) => handleDrag(e, 'student', false)}
            onDrop={(e) => handleDrop(e, 'student')}
            id="dropzone-students"
          >
            <input
              type="file"
              ref={studentInputRef}
              onChange={(e) => handleFileChange(e, 'student')}
              accept=".xlsx, .xls"
              className="hidden"
            />
            
            <div className="flex flex-col items-center justify-center text-center space-y-4 py-4">
              <div className={`p-4 rounded-full transition-colors ${studentBase64 ? 'bg-emerald-100 text-emerald-600' : 'bg-slate-100 text-slate-400'}`}>
                {studentBase64 ? (
                  <CheckCircle2 className="w-8 h-8" />
                ) : (
                  <FileSpreadsheet className="w-8 h-8" />
                )}
              </div>
              
              <div className="space-y-1">
                <p className="text-sm font-semibold text-slate-800">
                  {studentBase64 ? '설문 데이터 업로드 성공' : '유학생 설문 결과 엑셀 파일 (.xlsx)'}
                </p>
                <p className="text-xs text-slate-500">
                  {studentBase64 ? studentFileName : '이곳에 파일을 드래그 앤 드롭하거나 클릭하여 선택하세요'}
                </p>
              </div>

              {studentCount !== null ? (
                <div className="inline-flex items-center gap-1.5 bg-emerald-100 text-emerald-800 text-xs px-3 py-1 rounded-full font-semibold font-mono">
                  <Users className="w-3.5 h-3.5" />
                  <span>설문 데이터 {studentCount}건 확인 완료</span>
                </div>
              ) : (
                <button
                  onClick={() => studentInputRef.current?.click()}
                  className="bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold px-4 py-2 rounded-xl transition-colors"
                >
                  학생 파일 찾기
                </button>
              )}
            </div>
          </div>

          {/* Box 2: Dormitory Status */}
          <div
            className={`bg-white rounded-2xl border-2 p-6 transition-all relative ${
              roomDragActive ? 'border-blue-500 bg-blue-50/20' : 'border-slate-200'
            } ${roomBase64 ? 'border-dashed border-emerald-500/50 bg-emerald-50/5' : 'border-dashed'}`}
            onDragOver={(e) => handleDrag(e, 'room', true)}
            onDragLeave={(e) => handleDrag(e, 'room', false)}
            onDrop={(e) => handleDrop(e, 'room')}
            id="dropzone-rooms"
          >
            <input
              type="file"
              ref={roomInputRef}
              onChange={(e) => handleFileChange(e, 'room')}
              accept=".xlsx, .xls"
              className="hidden"
            />
            
            <div className="flex flex-col items-center justify-center text-center space-y-4 py-4">
              <div className={`p-4 rounded-full transition-colors ${roomBase64 ? 'bg-emerald-100 text-emerald-600' : 'bg-slate-100 text-slate-400'}`}>
                {roomBase64 ? (
                  <CheckCircle2 className="w-8 h-8" />
                ) : (
                  <Building2 className="w-8 h-8" />
                )}
              </div>
              
              <div className="space-y-1">
                <p className="text-sm font-semibold text-slate-800">
                  {roomBase64 ? '기숙사 현황 데이터 업로드 성공' : '기숙사 호실 현황 엑셀 파일 (.xlsx)'}
                </p>
                <p className="text-xs text-slate-500">
                  {roomBase64 ? roomFileName : '이곳에 기숙사 현황 파일을 드래그 앤 드롭하거나 클릭하여 선택하세요'}
                </p>
              </div>

              {roomCount !== null ? (
                <div className="inline-flex items-center gap-1.5 bg-emerald-100 text-emerald-800 text-xs px-3 py-1 rounded-full font-semibold font-mono">
                  <Building2 className="w-3.5 h-3.5" />
                  <span>방 설정 데이터 {roomCount}건 확인 완료</span>
                </div>
              ) : (
                <button
                  onClick={() => roomInputRef.current?.click()}
                  className="bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold px-4 py-2 rounded-xl transition-colors"
                >
                  기숙사 파일 찾기
                </button>
              )}
            </div>
          </div>
        </section>

        {/* Central Action Bar */}
        <section className="bg-white rounded-2xl border border-slate-200 p-6 flex flex-col items-center justify-center text-center shadow-xs space-y-4" id="action-panel">
          {(!studentBase64 || !roomBase64) ? (
            <div className="space-y-2 py-2">
              <p className="text-sm font-medium text-slate-500">
                기숙사 자동 배정을 시작하려면 위의 두 개 파일을 모두 업로드해 주셔야 합니다.
              </p>
              <div className="flex justify-center gap-4 text-xs text-slate-400 font-mono">
                <span className={studentBase64 ? 'text-emerald-600 font-semibold' : ''}>
                  {studentBase64 ? '✓ 설문 파일 완료' : '✗ 설문 파일 대기'}
                </span>
                <span className={roomBase64 ? 'text-emerald-600 font-semibold' : ''}>
                  {roomBase64 ? '✓ 기숙사 파일 완료' : '✗ 기숙사 파일 대기'}
                </span>
              </div>
            </div>
          ) : (
            <div className="w-full max-w-md space-y-4 py-2">
              {loading ? (
                <div className="space-y-4">
                  {/* Custom Spinner */}
                  <div className="flex items-center justify-center gap-3">
                    <div className="w-8 h-8 border-4 border-blue-600 border-t-transparent rounded-full animate-spin"></div>
                    <span className="font-bold text-slate-800 text-sm">기숙사 매칭 알고리즘 가동 중...</span>
                  </div>
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs text-slate-500 font-mono transition-all">
                    {loadingStep}
                  </div>
                </div>
              ) : (
                <div className="space-y-3">
                  <p className="text-xs text-slate-500 leading-normal">
                    기숙사 배정 준비가 완료되었습니다. 데이터 보안 매칭 정책에 따라 성별 분할, 입주일 우선순위 및 룸메이트 쌍 분석을 실행합니다.
                  </p>
                  <button
                    onClick={runAutoAssignment}
                    className="w-full bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-bold py-3 px-6 rounded-xl transition-all shadow-md shadow-blue-200 flex items-center justify-center gap-2 cursor-pointer"
                    id="run-assignment-btn"
                  >
                    <Sparkles className="w-4 h-4" />
                    <span>자동 기숙사 배정 시작</span>
                  </button>
                </div>
              )}
            </div>
          )}
        </section>

        {/* Dashboard Results (displayed when results are available) */}
        {stats && (
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="space-y-6"
            id="results-container"
          >
            {/* Top Stat Summary Cards */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6" id="stats-grid">
              
              {/* Card 1: Total Applicants */}
              <div className="bg-white rounded-2xl border border-slate-200 p-6 flex items-center gap-5 shadow-xs">
                <div className="p-4 rounded-xl bg-slate-100 text-slate-700">
                  <Users className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-xs text-slate-500 font-medium">총 신청 사생</p>
                  <p className="text-2xl font-black font-mono text-slate-900 mt-1">{stats.total_students}명</p>
                  <div className="flex items-center gap-1 text-[10px] text-slate-400 mt-1.5">
                    <span className="font-semibold text-slate-500">엑셀 원본 행 수 기준</span>
                  </div>
                </div>
              </div>

              {/* Card 2: Successfully Assigned */}
              <div className="bg-white rounded-2xl border border-slate-200 p-6 flex items-center gap-5 shadow-xs">
                <div className="p-4 rounded-xl bg-emerald-100 text-emerald-700">
                  <UserCheck className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-xs text-slate-500 font-medium">배정 성공 인원</p>
                  <p className="text-2xl font-black font-mono text-emerald-600 mt-1">{stats.assigned_count}명</p>
                  <div className="flex items-center gap-1 text-[10px] text-emerald-500 mt-1.5">
                    <span>수용 완료 100% 안착</span>
                  </div>
                </div>
              </div>

              {/* Card 3: Failed / Waitlist */}
              <div className="bg-white rounded-2xl border border-slate-200 p-6 flex items-center gap-5 shadow-xs">
                <div className="p-4 rounded-xl bg-rose-100 text-rose-700">
                  <UserX className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-xs text-slate-500 font-medium">배정 실패 사생</p>
                  <p className="text-2xl font-black font-mono text-rose-600 mt-1">{stats.failed_count}명</p>
                  <div className="flex items-center gap-1 text-[10px] text-rose-500 mt-1.5">
                    <span>정원 부족 / 성별 불일치</span>
                  </div>
                </div>
              </div>

              {/* Card 4: Success rate */}
              <div className="bg-white rounded-2xl border border-slate-200 p-6 flex items-center gap-5 shadow-xs">
                <div className="p-4 rounded-xl bg-blue-100 text-blue-700">
                  <TrendingUp className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-xs text-slate-500 font-medium">배정 성공률</p>
                  <p className="text-2xl font-black font-mono text-blue-600 mt-1">{stats.success_rate}%</p>
                  <div className="flex items-center gap-1 text-[10px] text-blue-500 mt-1.5">
                    <span>기존 잔여 정원 대비 산출</span>
                  </div>
                </div>
              </div>

            </div>

            {/* Results Action Tabs Control */}
            <div className="border-b border-slate-200 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 pt-4" id="results-nav">
              <div className="flex gap-2">
                <button
                  onClick={() => setResultsTab('dashboard')}
                  className={`px-4 py-2 text-sm font-semibold border-b-2 transition-all cursor-pointer ${
                    resultsTab === 'dashboard'
                      ? 'border-blue-600 text-blue-600'
                      : 'border-transparent text-slate-500 hover:text-slate-800'
                  }`}
                  id="tab-dashboard"
                >
                  종합 대시보드 시각화
                </button>
                <button
                  onClick={() => setResultsTab('records')}
                  className={`px-4 py-2 text-sm font-semibold border-b-2 transition-all cursor-pointer ${
                    resultsTab === 'records'
                      ? 'border-blue-600 text-blue-600'
                      : 'border-transparent text-slate-500 hover:text-slate-800'
                  }`}
                  id="tab-records"
                >
                  전체 배정 명단 및 상세 검색
                </button>
              </div>

              {/* Excel Download Call-to-Action */}
              <button
                onClick={downloadExcelResult}
                className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2 rounded-xl flex items-center gap-1.5 transition-all shadow-md shadow-emerald-100 cursor-pointer self-start sm:self-auto"
                id="download-results-excel"
              >
                <Download className="w-4 h-4" />
                <span>최종 배정 결과 엑셀 파일 다운로드 (멀티탭)</span>
              </button>
            </div>

            {/* Diagnostics / Column Parsing Debug */}
            {diagnostics && !showDiagnostics && (
              <div className="flex justify-end">
                <button onClick={() => setShowDiagnostics(true)} className="text-amber-600 text-xs underline hover:no-underline cursor-pointer">
                  🔍 컬럼 파싱 진단 다시 보기
                </button>
              </div>
            )}
            {diagnostics && showDiagnostics && (
              <div className="bg-amber-50 border border-amber-200 rounded-2xl p-4 text-xs space-y-2" id="column-diagnostics">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <Info className="w-4 h-4 text-amber-600" />
                    <span className="font-bold text-amber-800">엑셀 컬럼 파싱 진단</span>
                    {diagnostics.all_columns_matched
                      ? <span className="bg-emerald-100 text-emerald-700 px-2 py-0.5 rounded-full text-[10px] font-semibold">모든 컬럼 정상 매칭</span>
                      : <span className="bg-rose-100 text-rose-700 px-2 py-0.5 rounded-full text-[10px] font-semibold">일부 컬럼 매칭 실패</span>
                    }
                  </div>
                  <button onClick={() => setShowDiagnostics(false)} className="text-amber-400 hover:text-amber-600 cursor-pointer">✕</button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="bg-white rounded-xl p-3 border border-amber-100">
                    <p className="font-semibold text-amber-700 mb-1">📋 설문결과 첫 행 데이터</p>
                    <table className="w-full text-[10px] font-mono border-collapse">
                      <thead><tr className="border-b border-amber-200"><th className="text-left pr-2 py-1 text-amber-600">컬럼명</th><th className="text-left py-1 text-amber-600">값</th><th className="text-left pl-2 py-1 text-amber-600">매칭</th></tr></thead>
                      <tbody>
                        {Object.entries(diagnostics.student_raw_preview).map(([col, val]) => {
                          // Find which targetKey this column matched to
                          let matchedTo = '❌';
                          for (const [k, v] of Object.entries(diagnostics.student_parsed)) {
                            if (v === val) { matchedTo = k; break; }
                          }
                          // Check if any other column matched to the same value
                          if (matchedTo === '❌' && val === '') matchedTo = '-';
                          return <tr key={col} className="border-b border-amber-50"><td className="pr-2 py-0.5 text-slate-700">{col}</td><td className="py-0.5 text-slate-500 max-w-[120px] truncate">{val}</td><td className="pl-2 py-0.5">{matchedTo}</td></tr>;
                        })}
                      </tbody>
                    </table>
                    <p className="text-[10px] text-slate-400 mt-1">파싱 결과: {Object.keys(diagnostics.student_parsed).length > 0
                      ? Object.entries(diagnostics.student_parsed).map(([k, v]) => `${k}="${v}"`).join(', ')
                      : '❌ 매칭된 컬럼 없음'}
                    </p>
                  </div>
                  <div className="bg-white rounded-xl p-3 border border-amber-100">
                    <p className="font-semibold text-amber-700 mb-1">🏠 기숙사 첫 행 데이터</p>
                    <table className="w-full text-[10px] font-mono border-collapse">
                      <thead><tr className="border-b border-amber-200"><th className="text-left pr-2 py-1 text-amber-600">컬럼명</th><th className="text-left py-1 text-amber-600">값</th><th className="text-left pl-2 py-1 text-amber-600">매칭</th></tr></thead>
                      <tbody>
                        {Object.entries(diagnostics.room_raw_preview).map(([col, val]) => {
                          let matchedTo = '❌';
                          for (const [k, v] of Object.entries(diagnostics.room_parsed)) {
                            if (v === val) { matchedTo = k; break; }
                          }
                          if (matchedTo === '❌' && val === '') matchedTo = '-';
                          return <tr key={col} className="border-b border-amber-50"><td className="pr-2 py-0.5 text-slate-700">{col}</td><td className="py-0.5 text-slate-500 max-w-[120px] truncate">{val}</td><td className="pl-2 py-0.5">{matchedTo}</td></tr>;
                        })}
                      </tbody>
                    </table>
                    <p className="text-[10px] text-slate-400 mt-1">파싱 결과: {Object.keys(diagnostics.room_parsed).length > 0
                      ? Object.entries(diagnostics.room_parsed).map(([k, v]) => `${k}="${v}"`).join(', ')
                      : '❌ 매칭된 컬럼 없음'}
                    </p>
                  </div>
                </div>

                {diagnostics.warnings.length > 0 && (
                  <div className="bg-white rounded-xl p-3 border border-amber-100 space-y-1">
                    <p className="font-semibold text-rose-600">⚠️ 경고</p>
                    {diagnostics.warnings.map((w, i) => <p key={i} className="text-rose-500">{w}</p>)}
                  </div>
                )}
                <div className="bg-blue-50 rounded-xl p-3 border border-blue-100">
                  <p className="text-[10px] text-blue-700 font-semibold">💡 진단 팁</p>
                  <p className="text-[10px] text-blue-600 mt-0.5">
                    컬럼명 옆 ❌ = 해당 컬럼 미인식. <code className="bg-blue-100 px-1 rounded">server.log</code>에서 상세 로그 확인 가능.<br />
                    <strong>서버 재시작:</strong> 반드시 <code className="bg-blue-100 px-1 rounded">stop-server.bat</code> 실행 후 <code className="bg-blue-100 px-1 rounded">start-server.bat</code> 실행.<br />
                    <strong>버전 확인:</strong> http://localhost:3000/api/health 에서 <code className="bg-blue-100 px-1 rounded">build_id</code> 값 변경 확인 (재시작해야 변경됨).
                  </p>
                </div>
              </div>
            )}

            {/* Visualizer Charts view */}
            {resultsTab === 'dashboard' && (
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6" id="dashboard-charts-view">
                
                {/* Chart 1: Dormitory Capacity vs Occupancy */}
                <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-xs flex flex-col space-y-4" id="chart-dorm-occupancy">
                  <div>
                    <h3 className="font-bold text-slate-950 text-sm md:text-base">기숙사 유형별 가용 정원 대비 수용 현황</h3>
                    <p className="text-xs text-slate-500 mt-0.5">각 타입 기숙사의 가용 총 정원 대비 최종 배정 완료된 사생 현황</p>
                  </div>
                  
                  <div className="h-72 w-full pt-2">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart
                        data={stats.dorm_stats}
                        margin={{ top: 10, right: 10, left: -15, bottom: 0 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" vertical={false} />
                        <XAxis dataKey="type" stroke="#94a3b8" fontSize={11} tickLine={false} />
                        <YAxis stroke="#94a3b8" fontSize={11} tickLine={false} />
                        <Tooltip
                          contentStyle={{ background: '#ffffff', borderRadius: '12px', border: '1px solid #e2e8f0', fontSize: '12px' }}
                          formatter={(value, name) => [value, name === 'capacity' ? '총 정원(명)' : '배정된 인원(명)']}
                        />
                        <Legend verticalAlign="top" height={36} iconType="circle" fontSize={11} wrapperStyle={{ fontSize: '11px' }} />
                        <Bar dataKey="capacity" name="capacity" fill="#cbd5e1" radius={[4, 4, 0, 0]} barSize={36} />
                        <Bar dataKey="occupied" name="occupied" fill="#2563eb" radius={[4, 4, 0, 0]} barSize={36} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                  
                  <div className="bg-slate-50 rounded-xl p-3 flex flex-wrap gap-4 text-[11px] text-slate-600 font-mono">
                    {stats.dorm_stats.map((dorm, i) => (
                      <div key={i} className="flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-blue-600"></span>
                        <strong>{dorm.type}</strong>: {dorm.occupied} / {dorm.capacity} 명 ({dorm.occupancy_rate}%)
                      </div>
                    ))}
                  </div>
                </div>

                {/* Chart 2: Failed Reasons */}
                <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-xs flex flex-col space-y-4" id="chart-failed-reasons">
                  <div>
                    <h3 className="font-bold text-slate-950 text-sm md:text-base">배정 불가 사유 분석</h3>
                    <p className="text-xs text-slate-500 mt-0.5">정원 초과 등으로 인해 배정이 제한된 유학생의 주원인 통계</p>
                  </div>

                  {stats.failed_count === 0 ? (
                    <div className="flex-1 flex flex-col items-center justify-center text-center p-8 space-y-3">
                      <div className="bg-emerald-100 p-4 rounded-full text-emerald-600">
                        <Sparkles className="w-10 h-10 animate-bounce" />
                      </div>
                      <p className="text-sm font-semibold text-slate-800">모든 유학생 100% 배정 완료!</p>
                      <p className="text-xs text-slate-500 max-w-xs">가용 기숙사 객실 정원이 넉넉하여 배정에 실패한 사생이 전혀 없습니다.</p>
                    </div>
                  ) : (
                    <div className="flex-1 flex flex-col justify-between space-y-4">
                      <div className="h-48 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie
                              data={stats.failed_reasons}
                              cx="50%"
                              cy="50%"
                              innerRadius={50}
                              outerRadius={75}
                              paddingAngle={4}
                              dataKey="count"
                              nameKey="reason"
                            >
                              {stats.failed_reasons.map((entry, index) => (
                                <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                              ))}
                            </Pie>
                            <Tooltip
                              contentStyle={{ background: '#ffffff', borderRadius: '12px', border: '1px solid #e2e8f0', fontSize: '12px' }}
                              formatter={(value) => [`${value}명`]}
                            />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>

                      <div className="space-y-2">
                        {stats.failed_reasons.map((reason, index) => (
                          <div key={index} className="flex items-center justify-between border-b border-slate-100 pb-2 text-xs">
                            <div className="flex items-center gap-2">
                              <span className="w-2.5 h-2.5 rounded-full block shrink-0" style={{ backgroundColor: COLORS[index % COLORS.length] }}></span>
                              <span className="text-slate-700 font-medium">{reason.reason}</span>
                            </div>
                            <div className="font-mono text-slate-600">
                              <strong>{reason.count}명</strong> ({reason.percentage}%)
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

              </div>
            )}

            {/* Records Explorer view */}
            {resultsTab === 'records' && (
              <div className="bg-white border border-slate-200 rounded-2xl shadow-xs overflow-hidden flex flex-col" id="records-table-view">
                
                {/* Search Bar and Filters */}
                <div className="p-4 border-b border-slate-100 bg-slate-50 flex flex-col sm:flex-row gap-4 justify-between items-center" id="table-controls">
                  
                  {/* Status buttons */}
                  <div className="flex gap-1 bg-white p-1 rounded-xl border border-slate-200">
                    <button
                      onClick={() => setStatusFilter('all')}
                      className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all cursor-pointer ${
                        statusFilter === 'all'
                          ? 'bg-slate-800 text-white'
                          : 'text-slate-500 hover:text-slate-800'
                      }`}
                    >
                      전체 ({stats.total_students})
                    </button>
                    <button
                      onClick={() => setStatusFilter('assigned')}
                      className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all cursor-pointer ${
                        statusFilter === 'assigned'
                          ? 'bg-emerald-600 text-white'
                          : 'text-slate-500 hover:text-emerald-600'
                      }`}
                    >
                      배정 성공 ({stats.assigned_count})
                    </button>
                    <button
                      onClick={() => setStatusFilter('failed')}
                      className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all cursor-pointer ${
                        statusFilter === 'failed'
                          ? 'bg-rose-600 text-white'
                          : 'text-slate-500 hover:text-rose-600'
                      }`}
                    >
                      배정 실패 ({stats.failed_count})
                    </button>
                  </div>

                  {/* Search query input */}
                  <div className="relative w-full sm:max-w-xs">
                    <Search className="absolute left-3 top-2.5 w-4 h-4 text-slate-400" />
                    <input
                      type="text"
                      placeholder="학번, 이름, 방유형 등으로 검색..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full pl-9 pr-4 py-1.5 rounded-xl border border-slate-200 text-xs bg-white focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all font-medium"
                    />
                  </div>

                </div>

                {/* Table list */}
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-slate-700 text-xs font-medium">
                    <thead className="bg-slate-50 border-b border-slate-100 text-slate-400 font-bold uppercase tracking-wider">
                      <tr>
                        <th className="px-6 py-3">학번 (ID)</th>
                        <th className="px-6 py-3">성명</th>
                        <th className="px-6 py-3">성별</th>
                        <th className="px-6 py-3">신청 기숙사</th>
                        <th className="px-6 py-3">배정 상태</th>
                        <th className="px-6 py-3">배정 호실 / 실패 사유</th>
                        <th className="px-6 py-3">신청 룸메이트</th>
                        <th className="px-6 py-3">매칭 룸메이트</th>
                        <th className="px-6 py-3">입사 예정일</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredRecords().length === 0 ? (
                        <tr>
                          <td colSpan={9} className="px-6 py-12 text-center text-slate-400">
                            검색 조건과 일치하는 사생 내역이 없습니다.
                          </td>
                        </tr>
                      ) : (
                        filteredRecords().map((row, index) => (
                          <tr key={index} className="hover:bg-slate-50/50 transition-colors">
                            <td className="px-6 py-3 font-mono font-semibold text-slate-900">{row.id}</td>
                            <td className="px-6 py-3 font-bold">{row.name}</td>
                            <td className="px-6 py-3">
                              <span className={`px-2 py-0.5 rounded-md font-semibold text-[10px] ${row.gender === '남' ? 'bg-blue-50 text-blue-700' : 'bg-pink-50 text-pink-700'}`}>
                                {row.gender}
                              </span>
                            </td>
                            <td className="px-6 py-3">{row.pref}</td>
                            <td className="px-6 py-3">
                              <span className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full font-bold text-[10px] ${
                                row.status === '성공'
                                  ? 'bg-emerald-100 text-emerald-800'
                                  : 'bg-rose-100 text-rose-800'
                              }`}>
                                {row.status === '성공' ? (
                                  <>
                                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                                    배정성공
                                  </>
                                ) : (
                                  <>
                                    <span className="w-1.5 h-1.5 rounded-full bg-rose-500"></span>
                                    배정실패
                                  </>
                                )}
                              </span>
                            </td>
                            <td className="px-6 py-3">
                              <span className={row.status === '성공' ? 'font-mono font-bold text-blue-700 bg-blue-50 px-2 py-1 rounded-lg' : 'text-rose-600 font-semibold'}>
                                {row.detail}
                              </span>
                            </td>
                            <td className="px-6 py-3 font-mono text-[10px] text-slate-500">{row.roommateReq}</td>
                            <td className="px-6 py-3 font-medium text-slate-700">
                              {row.roommateMatched ? (
                                <span className="inline-flex items-center gap-1 bg-emerald-50 text-emerald-800 px-2 py-1 rounded-lg text-[10px] border border-emerald-100 font-semibold">
                                  <Users className="w-3 h-3" />
                                  {row.roommateMatched}
                                </span>
                              ) : (
                                <span className="text-slate-400 font-normal">-</span>
                              )}
                            </td>
                            <td className="px-6 py-3 font-mono text-slate-500">{row.movein}</td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Bottom info on table filter */}
                <div className="p-3 border-t border-slate-100 bg-slate-50 text-[10px] text-slate-400 text-right">
                  필터링 및 검색된 레코드: <strong>{filteredRecords().length}</strong> 건
                </div>

              </div>
            )}

          </motion.div>
        )}

      </main>

      {/* Footer */}
      <footer className="bg-slate-900 text-slate-400 py-8 px-6 mt-12 border-t border-slate-800" id="main-footer">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-center gap-4 text-xs">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-slate-200">기숙사 자동 배정 프로그램 (MVP)</span>
            <span className="text-slate-600">|</span>
            <span>데이터가 서버 메모리에서 즉시 처리되어 안전합니다.</span>
          </div>
          <div className="text-slate-500 text-center md:text-right">
            © 2026 대학 기숙사 자동 배정 시스템 구축. All rights reserved.
          </div>
        </div>
      </footer>

    </div>
  );
}
