import * as xlsx from 'xlsx';

// Korean names to generate realistic mock students
const firstNamesM = ['민수', '영수', '철수', '준우', '현우', '동현', '지훈', '서준', '도윤', '하준', '우진', '건우', '도현', '재원', '시우', '지호', '태민', '상현', '준영', '찬우'];
const firstNamesF = ['영희', '지영', '수지', '지원', '서연', '민서', '하은', '서현', '지민', '채원', '윤아', '다은', '은서', '예은', '지우', '소율', '아린', '유나', '수민', '민지'];
const lastNames = ['김', '이', '박', '최', '정', '강', '조', '윤', '장', '임', '한', '오', '서', '신', '권', '황', '안', '송', '전', '홍'];

function getRandomElement<T>(arr: T[]): T {
  return arr[Math.floor(Math.random() * arr.length)];
}

export function generateAndDownloadSampleStudents() {
  const students: any[] = [];
  const count = 200;
  
  // Create student basic data first so we can assign valid roommate requests
  const tempStudents: { id: string; name: string; gender: string }[] = [];
  
  for (let i = 1; i <= count; i++) {
    const id = String(20260000 + i);
    const gender = i % 2 === 0 ? '남' : '여';
    const lastName = getRandomElement(lastNames);
    const firstName = gender === '남' ? getRandomElement(firstNamesM) : getRandomElement(firstNamesF);
    
    tempStudents.push({
      id,
      name: `${lastName}${firstName}`,
      gender
    });
  }
  
  // Assign preferences, roommate requests and dates
  for (let i = 0; i < count; i++) {
    const current = tempStudents[i];
    
    // Preference: 50% A, 30% B, 20% C
    const randPref = Math.random();
    const dorm_pref = randPref < 0.5 ? 'A타입' : randPref < 0.8 ? 'B타입' : 'C타입';
    
    // Roommate requests:
    // - 15% reciprocal (A <-> B)
    // - 10% unidirectional (A -> B, B doesn't request A)
    // - 5% invalid (A -> B but gender mismatch or B doesn't exist)
    // - 70% individual (none)
    let roommate_id = '';
    const groupRand = Math.random();
    
    if (groupRand < 0.15) {
      // Reciprocal: link to the next/prev student of same gender
      const step = current.gender === '남' ? 2 : 2;
      const targetIndex = (i + step) % count;
      const target = tempStudents[targetIndex];
      if (target && target.gender === current.gender && target.id !== current.id) {
        roommate_id = target.id;
        // Make it mutual sometimes
        if (Math.random() < 0.7) {
          // Set reciprocal link
          tempStudents[targetIndex] = { ...target }; // trigger write if needed (already stored in array)
        }
      }
    } else if (groupRand < 0.25) {
      // Unidirectional: point to another student of same gender
      const targetIndex = (i + 13) % count;
      const target = tempStudents[targetIndex];
      if (target && target.gender === current.gender && target.id !== current.id) {
        roommate_id = target.id;
      }
    } else if (groupRand < 0.30) {
      // Invalid: gender mismatch (pointing to student with different gender)
      const targetIndex = (i + 1) % count;
      const target = tempStudents[targetIndex];
      if (target && target.gender !== current.gender) {
        roommate_id = target.id;
      }
    }
    
    // Move-in Dates: Distributed around late August / early September
    const dateChoices = ['2026-08-25', '2026-08-26', '2026-08-27', '2026-08-28', '2026-08-29'];
    const movein_date = getRandomElement(dateChoices);
    
    students.push({
      '학번': current.id,
      '학생 성명': current.name,
      '성별': current.gender,
      '선호 기숙사 유형': dorm_pref,
      '희망 룸메이트 학번': roommate_id || '',
      '입주일': movein_date
    });
  }
  
  // Write Excel
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet(students);
  xlsx.utils.book_append_sheet(wb, ws, '외국인 유학생 설문 결과');
  xlsx.writeFile(wb, '설문결과_샘플_200명.xlsx');
}

export function generateAndDownloadSampleRooms() {
  const rooms: any[] = [];
  
  // We need enough rooms for 200 students.
  // Suppose A타입 has 50% capacity, B타입 has 30%, C타입 has 20%.
  // We'll create:
  // - A타입 (정원 2명 / 4명): 40개 객실 (정원 총 120명)
  // - B타입 (정원 2명): 30개 객실 (정원 총 60명)
  // - C타입 (정원 2명 / 1명): 20개 객실 (정원 총 30명)
  // Total capacity ~ 210, nicely balanced with demand!
  
  let idM = 101;
  let idF = 201;
  
  // A타입 Rooms
  for (let i = 1; i <= 30; i++) {
    const isMale = i % 2 === 0;
    const roomNo = isMale ? `A-${idM++}` : `A-${idF++}`;
    rooms.push({
      '방 번호': roomNo,
      '기숙사 유형': 'A타입',
      '성별 구분': isMale ? '남' : '여',
      '수용 정원': i % 5 === 0 ? 4 : 2 // some 4-person rooms, mostly 2-person
    });
  }
  
  // B타입 Rooms
  idM = 101;
  idF = 201;
  for (let i = 1; i <= 25; i++) {
    const isMale = i % 2 === 0;
    const roomNo = isMale ? `B-${idM++}` : `B-${idF++}`;
    rooms.push({
      '방 번호': roomNo,
      '기숙사 유형': 'B타입',
      '성별 구분': isMale ? '남' : '여',
      '수용 정원': 2
    });
  }
  
  // C타입 Rooms
  idM = 101;
  idF = 201;
  for (let i = 1; i <= 15; i++) {
    const isMale = i % 2 === 0;
    const roomNo = isMale ? `C-${idM++}` : `C-${idF++}`;
    rooms.push({
      '방 번호': roomNo,
      '기숙사 유형': 'C타입',
      '성별 구분': isMale ? '남' : '여',
      '수용 정원': i % 4 === 0 ? 1 : 2 // some single rooms, mostly doubles
    });
  }
  
  // Write Excel
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet(rooms);
  xlsx.utils.book_append_sheet(wb, ws, '기숙사 호실 현황');
  xlsx.writeFile(wb, '기숙사현황_샘플_호실목록.xlsx');
}
