import express from 'express';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import * as xlsx from 'xlsx';
import { Student, Room, AssignmentResult, FailedAssignment, AssignmentDashboardStats } from './src/types';

const BUILD_ID = new Date().toISOString().slice(0, 16).replace('T', ' ');

const studentSynonyms = {
  id: ['학번', 'id', 'studentid', '학부생학번', '등록번호', '사번', '학적번호', 'studentnumber', 'studentno', '학생번호', '번호', '아이디'],
  name: ['이름', '성명', 'name', '학생명', '학생이름', 'studentname', 'fullname', '성함', '한글이름', 'fullname'],
  gender: ['성별', 'gender', 'sex', '남여', '남녀', '구분', '성별구분', '성구분', '성별(남/여)'],
  dorm_pref: ['선호기숙사', '선호기숙사유형', '기숙사선호', '선호타입', 'dormpref', 'preference', '희망기숙사', '기숙사유형', 'dorm', 'preferreddorm', 'prefdorm', 'dormitory', '희망기숙사유형', '선호건물', '건물', '희망건물', '선호', '기숙사타입', '선호기숙사타입', '희망기숙사타입'],
  roommate_id: ['룸메이트', '룸메이트학번', '룸메이트id', '희망룸메이트', '룸메학번', '룸메id', 'roommateid', 'roommate', '같이살고싶은사람', '룸메이트(학번)'],
  movein_date: ['입사일', '입주일', '입사일자', '입주일자', 'moveindate', 'movein', 'move-in', '입사예정일', '입주예정일', 'moveinday']
};

const roomSynonyms = {
  room_no: ['방번호', '호실', '호수', 'roomno', 'roomnumber', '방', '호', 'room', 'roomname', 'roomid', '호실번호', '호실명', '객실번호'],
  type: ['기숙사유형', '기숙사종류', '방종류', '유형', 'type', 'roomtype', 'dormtype', 'dormname', '방유형', '숙소유형', '건물', '동', '건물명', '기숙사', '기숙사타입', '방타입', 'category', 'accommodation', 'accom', 'building', 'build', 'bldg'],
  gender_type: ['성별구분', '방성별', '성별', 'gendertype', 'gender', 'sex', '구분', '남녀구분', '성구분', '성별(남/여)', '성별구분(남/여)'],
  capacity: ['정원', '수용인원', '인원', 'capacity', 'cap', 'size', 'max', 'spots', '수용정원', '최대인원']
};

const RE_INVISIBLE = /[\u00AD\u034F\u061C\u115F\u1160\u17B4\u17B5\u180B-\u180F\u200B-\u200F\u2028\u2029\u202A-\u202F\u205F\u2060-\u2069\u3000\u3164\uFE00-\uFE0F\uFEFF\uFFF9-\uFFFB]/g;
const RE_WS = /[\s_]/g;

function normalizeColumnName(raw: string): string {
  return raw.trim().replace(RE_INVISIBLE, '').toLowerCase().replace(RE_WS, '');
}

function mapRowKeys(row: any, keyMap: Record<string, string[]>): Record<string, any> {
  const result: Record<string, any> = {};
  const cols = Object.keys(row).map(raw => ({ raw, normalized: normalizeColumnName(raw) }));
  interface Match { colIndex: number; targetKey: string; synLen: number; isExact: boolean }
  const allMatches: Match[] = [];
  const targetOrder = Object.keys(keyMap);
  for (let ci = 0; ci < cols.length; ci++) {
    const nk = cols[ci].normalized;
    for (const [targetKey, synonyms] of Object.entries(keyMap)) {
      for (const syn of synonyms) {
        if (nk === syn) allMatches.push({ colIndex: ci, targetKey, synLen: syn.length, isExact: true });
        else if (nk.includes(syn) && syn.length > 0) allMatches.push({ colIndex: ci, targetKey, synLen: syn.length, isExact: false });
        else if ((nk.length >= 2 || /^[가-힣]$/.test(nk)) && syn.includes(nk)) allMatches.push({ colIndex: ci, targetKey, synLen: nk.length, isExact: false });
      }
    }
  }
  allMatches.sort((a, b) => {
    if (a.isExact !== b.isExact) return a.isExact ? -1 : 1;
    if (a.synLen !== b.synLen) return b.synLen - a.synLen;
    return targetOrder.indexOf(a.targetKey) - targetOrder.indexOf(b.targetKey);
  });
  const usedCols = new Set<number>();
  const usedTargets = new Set<string>();
  for (const m of allMatches) {
    if (usedCols.has(m.colIndex)) continue;
    if (usedTargets.has(m.targetKey)) continue;
    result[m.targetKey] = row[cols[m.colIndex].raw];
    usedCols.add(m.colIndex);
    usedTargets.add(m.targetKey);
  }
  return result;
}

function normalizeDormType(t: any): string {
  if (!t) return '';
  const val = String(t).trim();
  const map: Record<string, string> = {
    'a': 'A', 'a타입': 'A', 'a형': 'A', 'a관': 'A', 'a동': 'A', '기숙사a': 'A',
    'typea': 'A', 'type-a': 'A', 'type_a': 'A', 'atype': 'A', 'a(더블)': 'A', 'a(2인실)': 'A', 'a실': 'A',
    'b': 'B', 'b타입': 'B', 'b형': 'B', 'b관': 'B', 'b동': 'B', '기숙사b': 'B',
    'typeb': 'B', 'type-b': 'B', 'type_b': 'B', 'btype': 'B', 'b(더블)': 'B', 'b(2인실)': 'B', 'b실': 'B',
    'c': 'C', 'c타입': 'C', 'c형': 'C', 'c관': 'C', 'c동': 'C', '기숙사c': 'C',
    'typec': 'C', 'type-c': 'C', 'type_c': 'C', 'ctype': 'C', 'c(싱글)': 'C', 'c(1인실)': 'C', 'c실': 'C',
    '싱글': '1인실', '1인실': '1인실', 'single': '1인실', '1인': '1인실',
    '더블': '2인실', '2인실': '2인실', 'double': '2인실', '트윈': '2인실', 'twin': '2인실', '2인': '2인실',
  };
  const key = val.toLowerCase().replace(RE_INVISIBLE, '').replace(RE_WS, '').replace(/[-]/g, '').replace(/\(.*?\)/g, '');
  return map[key] || val;
}

function normalizeGender(g: any): string {
  if (!g) return '';
  const val = String(g).trim().toLowerCase();
  if (val.includes('남') || val.startsWith('m') || val.includes('male') || val === '1') return 'M';
  if (val.includes('여') || val.startsWith('f') || val.includes('female') || val.includes('woman') || val === '2') return 'F';
  return val.toUpperCase();
}

function parseExcelDate(val: any): string {
  if (!val) return '';
  if (typeof val === 'number') {
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  let str = String(val).trim().replace(/[\.\/]/g, '-');
  const match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
  return str;
}

async function start() {
  const app = express();
  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ limit: '50mb', extended: true }));

  app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', build_id: BUILD_ID, timestamp: new Date().toISOString() });
  });

  app.post('/api/assign', (req, res) => {
    try {
      const { studentFile, roomFile } = req.body;
      if (!studentFile || !roomFile) return res.status(400).json({ error: '두 개의 엑셀 파일(설문 결과 및 기숙사 현황)이 필요합니다.' });
      const studentWorkbook = xlsx.read(studentFile, { type: 'base64' });
      const studentRaw = xlsx.utils.sheet_to_json(studentWorkbook.Sheets[studentWorkbook.SheetNames[0]]);
      const roomWorkbook = xlsx.read(roomFile, { type: 'base64' });
      const roomRaw = xlsx.utils.sheet_to_json(roomWorkbook.Sheets[roomWorkbook.SheetNames[0]]);
      if (studentRaw.length === 0) return res.status(400).json({ error: '설문결과 파일에 데이터가 없습니다.' });
      if (roomRaw.length === 0) return res.status(400).json({ error: '기숙사현황 파일에 데이터가 없습니다.' });

      const students: Student[] = studentRaw.map((row: any) => {
        const mapped = mapRowKeys(row, studentSynonyms);
        return { id: mapped.id ? String(mapped.id).trim() : '', name: mapped.name ? String(mapped.name).trim() : '이름없음', gender: normalizeGender(mapped.gender), dorm_pref: mapped.dorm_pref ? String(mapped.dorm_pref).trim() : '미지정', roommate_id: mapped.roommate_id ? String(mapped.roommate_id).trim() : undefined, movein_date: parseExcelDate(mapped.movein_date) };
      }).filter(s => s.id !== '');
      const rooms: Room[] = roomRaw.map((row: any) => {
        const mapped = mapRowKeys(row, roomSynonyms);
        const capacityVal = parseInt(String(mapped.capacity || 2), 10);
        return { room_no: mapped.room_no ? String(mapped.room_no).trim() : '', type: mapped.type ? String(mapped.type).trim() : '일반', gender_type: normalizeGender(mapped.gender_type), capacity: isNaN(capacityVal) ? 2 : capacityVal, current_occupancy: 0 };
      }).filter(r => r.room_no !== '');
      if (students.length === 0) return res.status(400).json({ error: '유효한 학생 데이터를 파싱할 수 없습니다.' });
      if (rooms.length === 0) return res.status(400).json({ error: '유효한 방 데이터를 파싱할 수 없습니다.' });

      const studentsWithoutGender = students.filter(s => !s.gender);
      const studentsWithoutDormPref = students.filter(s => s.dorm_pref === '미지정');
      const roomsWithoutGenderType = rooms.filter(r => !r.gender_type);

      const studentMap = new Map<string, Student>();
      students.forEach(s => studentMap.set(s.id, s));
      const adjList = new Map<string, Set<string>>();
      students.forEach(s => adjList.set(s.id, new Set<string>()));
      students.forEach(student => {
        const rId = student.roommate_id;
        if (rId && rId !== student.id) {
          const target = studentMap.get(rId);
          if (target && target.gender === student.gender) { adjList.get(student.id)!.add(rId); adjList.get(rId)!.add(student.id); }
        }
      });

      const visited = new Set<string>();
      interface StudentGroup { id: string; students: Student[]; gender: string; earliest_date: string; dorm_pref: string; randomWeight: number; }
      const rawGroups: StudentGroup[] = [];
      students.forEach(student => {
        if (!visited.has(student.id)) {
          const groupStudents: Student[] = [];
          const queue: string[] = [student.id];
          visited.add(student.id);
          while (queue.length > 0) {
            const currId = queue.shift()!;
            const currStudent = studentMap.get(currId);
            if (currStudent) groupStudents.push(currStudent);
            const neighbors = adjList.get(currId);
            if (neighbors) for (const neighbor of neighbors) { if (!visited.has(neighbor)) { visited.add(neighbor); queue.push(neighbor); } }
          }
          let earliestDate = '9999-12-31';
          groupStudents.forEach(s => { if (s.movein_date && s.movein_date < earliestDate) earliestDate = s.movein_date; });
          if (earliestDate === '9999-12-31') earliestDate = '2026-09-01';
          rawGroups.push({ id: groupStudents[0].id, students: groupStudents, gender: groupStudents[0].gender, earliest_date: earliestDate, dorm_pref: groupStudents[0].dorm_pref, randomWeight: Math.random() });
        }
      });
      rawGroups.sort((a, b) => {
        if (a.earliest_date < b.earliest_date) return -1;
        if (a.earliest_date > b.earliest_date) return 1;
        return a.randomWeight - b.randomWeight;
      });

      const assignedList: AssignmentResult[] = [];
      const failedList: FailedAssignment[] = [];
      const roomOccupancy = new Map<string, number>();
      rooms.forEach(r => roomOccupancy.set(r.room_no, 0));

      function pushAssignment(student: Student, room: Room, groupStudents: Student[]) {
        const otherGroupMembers = groupStudents.filter(o => o.id !== student.id).map(o => `${o.name}(${o.id})`).join(', ');
        assignedList.push({ student_id: student.id, student_name: student.name, gender: student.gender === 'M' ? '남' : student.gender === 'F' ? '여' : student.gender, movein_date: student.movein_date, room_no: room.room_no, dorm_type: room.type, roommate_id_requested: student.roommate_id, roommate_matched_id: otherGroupMembers || undefined });
      }
      function pushFailed(student: Student, reason: string) {
        failedList.push({ student_id: student.id, student_name: student.name, gender: student.gender === 'M' ? '남' : student.gender === 'F' ? '여' : student.gender, dorm_pref: student.dorm_pref, movein_date: student.movein_date, roommate_id: student.roommate_id, reason });
      }

      rawGroups.forEach(group => {
        const groupSize = group.students.length;
        const matchedRooms = rooms.filter(r => normalizeDormType(r.type) === normalizeDormType(group.dorm_pref) && r.gender_type === group.gender);
        if (matchedRooms.length === 0) { group.students.forEach(s => pushFailed(s, '성별 방 불일치 또는 미보유 기숙사 유형')); return; }
        let assignedRoom: Room | null = null;
        for (const r of matchedRooms) { if (r.capacity - (roomOccupancy.get(r.room_no) || 0) >= groupSize) { assignedRoom = r; break; } }
        if (assignedRoom) {
          roomOccupancy.set(assignedRoom.room_no, (roomOccupancy.get(assignedRoom.room_no) || 0) + groupSize);
          group.students.forEach(student => pushAssignment(student, assignedRoom!, group.students)); return;
        }
        const sortedRooms = [...matchedRooms].sort((a, b) => (b.capacity - (roomOccupancy.get(b.room_no) || 0)) - (a.capacity - (roomOccupancy.get(a.room_no) || 0)));
        const unplaced: Student[] = [];
        const placed: Map<string, Student[]> = new Map();
        for (const student of group.students) {
          let placedStudent = false;
          for (const r of sortedRooms) {
            if (r.capacity - (roomOccupancy.get(r.room_no) || 0) >= 1) {
              roomOccupancy.set(r.room_no, (roomOccupancy.get(r.room_no) || 0) + 1);
              if (!placed.has(r.room_no)) placed.set(r.room_no, []);
              placed.get(r.room_no)!.push(student); placedStudent = true; break;
            }
          }
          if (!placedStudent) unplaced.push(student);
        }
        for (const [roomNo, roomStudents] of placed.entries()) {
          const room = sortedRooms.find(r => r.room_no === roomNo)!;
          for (const student of roomStudents) pushAssignment(student, room, roomStudents);
        }
        if (unplaced.length > 0) {
          const reason = groupSize > 1 && unplaced.length === groupSize ? '기숙사 정원 부족 (그룹 동실 배정 불가)' : '기숙사 정원 부족';
          unplaced.forEach(s => pushFailed(s, reason));
        }
      });

      const stats: AssignmentDashboardStats = {
        total_students: students.length, assigned_count: assignedList.length, failed_count: failedList.length,
        success_rate: students.length > 0 ? parseFloat(((assignedList.length / students.length) * 100).toFixed(1)) : 0,
        dorm_stats: Array.from(new Map(rooms.map(r => [r.type, { capacity: 0, occupied: 0 }])).entries()).map(([type]) => {
          let cap = 0, occ = 0;
          rooms.filter(r => r.type === type).forEach(r => { cap += r.capacity; occ += (roomOccupancy.get(r.room_no) || 0); });
          return { type, capacity: cap, occupied: occ, occupancy_rate: cap > 0 ? parseFloat(((occ / cap) * 100).toFixed(1)) : 0 };
        }),
        failed_reasons: []
      };
      const failedReasonsMap = new Map<string, number>();
      failedList.forEach(item => failedReasonsMap.set(item.reason, (failedReasonsMap.get(item.reason) || 0) + 1));
      stats.failed_reasons = Array.from(failedReasonsMap.entries()).map(([reason, count]) => ({ reason, count, percentage: failedList.length > 0 ? parseFloat(((count / failedList.length) * 100).toFixed(1)) : 0 }));

      const warnings: string[] = [];
      if (studentsWithoutGender.length > 0) warnings.push(`${studentsWithoutGender.length}명의 학생이 성별 정보가 없습니다. (예: ${studentsWithoutGender[0].id})`);
      if (studentsWithoutDormPref.length > 0) warnings.push(`${studentsWithoutDormPref.length}명의 학생이 선호 기숙사 정보가 없습니다. (기본값 처리)`);
      if (roomsWithoutGenderType.length > 0) warnings.push(`${roomsWithoutGenderType.length}개의 방이 성별 구분 정보가 없습니다.`);

      const firstStudentRaw = studentRaw[0] || {};
      const firstStudentMapped = mapRowKeys(firstStudentRaw, studentSynonyms);
      const firstRoomRaw = roomRaw[0] || {};
      const firstRoomMapped = mapRowKeys(firstRoomRaw, roomSynonyms);
      const studentMatched: Record<string, string> = {};
      for (const key of Object.keys(studentSynonyms)) { if (firstStudentMapped[key] !== undefined) studentMatched[key] = String(firstStudentMapped[key]).substring(0, 50); }
      const roomMatched: Record<string, string> = {};
      for (const key of Object.keys(roomSynonyms)) { if (firstRoomMapped[key] !== undefined) roomMatched[key] = String(firstRoomMapped[key]).substring(0, 50); }

      return res.json({
        build_id: BUILD_ID, success: true, stats, assigned_list: assignedList, failed_list: failedList,
        diagnostics: { student_parsed: studentMatched, room_parsed: roomMatched, warnings, all_columns_matched: studentsWithoutGender.length === 0 && studentsWithoutDormPref.length === 0 && roomsWithoutGenderType.length === 0 }
      });
    } catch (err: any) {
      return res.status(500).json({ error: `배정 프로세스 중 오류가 발생했습니다: ${err.message || err}` });
    }
  });

  app.post('/api/download', (req, res) => {
    try {
      const { assigned_list, failed_list } = req.body;
      const wb = xlsx.utils.book_new();
      const wsAssigned = xlsx.utils.json_to_sheet((assigned_list || []).map((item: any) => ({
        '방 번호': item.room_no, '기숙사 유형': item.dorm_type, '학생 성명': item.student_name,
        '학번 (ID)': item.student_id, '성별': item.gender, '입사 예정일': item.movein_date,
        '희망 룸메이트 학번': item.roommate_id_requested || '없음', '매칭된 룸메이트': item.roommate_matched_id || '없음'
      })));
      xlsx.utils.book_append_sheet(wb, wsAssigned, '배정 결과 목록');
      const wsFailed = xlsx.utils.json_to_sheet((failed_list || []).map((item: any) => ({
        '학번 (ID)': item.student_id, '학생 성명': item.student_name, '성별': item.gender,
        '희망 기숙사 유형': item.dorm_pref, '입사 예정일': item.movein_date,
        '신청 룸메이트 학번': item.roommate_id || '없음', '배정 실패 사유': item.reason
      })));
      xlsx.utils.book_append_sheet(wb, wsFailed, '배정 실패자 목록');
      const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="dormitory_assignment_results.xlsx"');
      return res.end(buffer);
    } catch (err: any) {
      return res.status(500).json({ error: `엑셀 생성 중 오류가 발생했습니다: ${err.message || err}` });
    }
  });

  const vite = await createViteServer({
    server: { middlewareMode: true },
    appType: 'spa',
  });
  app.use(vite.middlewares);

  const PORT = 3000;
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

start();
