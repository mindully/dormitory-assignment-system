export interface Student {
  id: string; // 학번 / ID
  name: string; // 이름
  gender: 'M' | 'F' | string; // 성별 ("남" / "여", "M" / "F", etc.)
  dorm_pref: string; // 선호 기숙사 유형 (e.g., "A", "B", "C" or "Single", "Double")
  roommate_id?: string; // 룸메이트 희망 학생 학번 (optional)
  movein_date: string; // 입주일 (YYYY-MM-DD)
}

export interface Room {
  room_no: string; // 방 번호
  type: string; // 기숙사 유형 (e.g., "A", "B", "C" or "Single", "Double")
  gender_type: 'M' | 'F' | string; // 수용 성별
  capacity: number; // 정원
  current_occupancy: number; // 현재 배정된 인원
}

export interface AssignmentResult {
  student_id: string;
  student_name: string;
  gender: string;
  movein_date: string;
  room_no: string;
  dorm_type: string;
  roommate_id_requested?: string;
  roommate_matched_id?: string; // 매칭된 룸메이트 학번 (실제 매칭된 경우)
}

export interface FailedAssignment {
  student_id: string;
  student_name: string;
  gender: string;
  dorm_pref: string;
  movein_date: string;
  roommate_id?: string;
  reason: string; // 실패 사유 (e.g., "성별 불일치", "정원 초과", "친구 정보 오류" 등)
}

export interface AssignmentDashboardStats {
  total_students: number;
  assigned_count: number;
  failed_count: number;
  success_rate: number; // %
  
  // 기숙사 유형별 수용율
  dorm_stats: {
    type: string;
    capacity: number;
    occupied: number;
    occupancy_rate: number;
  }[];
  
  // 실패 사유 통계
  failed_reasons: {
    reason: string;
    count: number;
    percentage: number;
  }[];
}

export interface Diagnostics {
  student_parsed: Record<string, string>;
  room_parsed: Record<string, string>;
  raw_student_columns: string[];
  raw_room_columns: string[];
  student_raw_preview: Record<string, string>;
  room_raw_preview: Record<string, string>;
  warnings: string[];
  all_columns_matched: boolean;
}

export interface AssignmentResponse {
  success: boolean;
  build_id: string;
  stats: AssignmentDashboardStats;
  assigned_list: AssignmentResult[];
  failed_list: FailedAssignment[];
  diagnostics: Diagnostics;
}
