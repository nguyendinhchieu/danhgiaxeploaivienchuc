import { EvaluationRecord, Staff, TargetType, XepLoai } from '../types';
import { STAFF_LIST } from '../data/staffData';
import { getCriteriaForTarget, calculateXepLoai } from '../data/criteriaData';
import * as XLSX from 'xlsx';

const STORAGE_KEY_EVALUATIONS = 'ndc_evaluations_v1';
const STORAGE_KEY_CURRENT_USER = 'ndc_current_user_v1';
const STORAGE_KEY_CURRENT_PERIOD = 'ndc_current_period_v1';

export function getInitialEvaluations(): Record<string, EvaluationRecord> {
  const records: Record<string, EvaluationRecord> = {};
  const currentPeriod = 'thang_10_2026';

  STAFF_LIST.forEach((staff, index) => {
    const key = `${staff.id}_${currentPeriod}`;
    const criteria = getCriteriaForTarget(staff.targetType);
    
    // Seed high-performing, realistic score profiles
    // High performers ~ 15-20% get >= 90 with exemplary products
    const isExemplary = index % 5 === 0; // ~20%
    const isAboveAverage = index % 2 === 0;

    let baseDeduction = 0;
    if (isExemplary) {
      baseDeduction = (index % 3); // 0, 1, or 2 pts deducted -> total 92 - 96
    } else if (isAboveAverage) {
      baseDeduction = 7 + (index % 6); // total 80 - 87
    } else {
      baseDeduction = 12 + (index % 8); // total 75 - 82
    }

    const scores: Record<string, any> = {};
    let selfTotal = 0;
    let evalTotal = 0;
    let bghTotal = 0;

    criteria.forEach((c, cIdx) => {
      // Allocate slight deduction to some criteria
      let cSelf = c.maxScore;
      if (cIdx === 1 && baseDeduction > 2) cSelf = Math.max(c.maxScore - 1, 0);
      if (cIdx === 4 && baseDeduction > 5) cSelf = Math.max(c.maxScore - 2, 0);
      if (cIdx === 6 && baseDeduction > 10) cSelf = Math.max(c.maxScore - 3, 0);

      // Evaluator score might match or differ by 1 point
      let cEval = cSelf;
      if (!isExemplary && cIdx === 2) {
        cEval = Math.max(cSelf - (index % 2), 0);
      }

      // BGH score
      let cBgh = cEval;

      selfTotal += cSelf;
      evalTotal += cEval;
      bghTotal += cBgh;

      let evidence = '';
      if (c.id === 'TCC_GV_2') {
        evidence = 'Đã hoàn thành ký số 100% Học bạ số đợt 1; ứng dụng AI thiết kế bài giảng.';
      } else if (c.id === 'TCC_GV_3') {
        evidence = 'Đã triển khai 01 bài học STEM/tháng; có kế hoạch dạy học song ngữ.';
      } else if (c.id === 'TCC_GV_1') {
        evidence = 'Giáo án đầy đủ, tích hợp địa chỉ GDQPAN chuẩn theo quy định Bộ GD&ĐT.';
      } else if (c.id === 'TCC_NV_1') {
        evidence = 'Hoàn thành 100% khối lượng công việc chuyên môn đúng hạn.';
      } else if (c.id === 'TCC_QL_1') {
        evidence = '100% tổ viên trong tổ hoàn thành tốt nhiệm vụ, không có đơn thư khiếu nại.';
      }

      scores[c.id] = {
        criterionId: c.id,
        selfScore: cSelf,
        evaluatorScore: cEval,
        bghScore: cBgh,
        evidence: evidence,
      };
    });

    const hasProd = isExemplary;
    const selfXepLoai = calculateXepLoai(selfTotal, hasProd);
    const evalXepLoai = calculateXepLoai(evalTotal, hasProd);
    const bghXepLoai = calculateXepLoai(bghTotal, hasProd);

    // Some are approved, some evaluated, some submitted
    let status: EvaluationRecord['status'] = 'APPROVED';
    if (index > 55) {
      status = 'SUBMITTED';
    } else if (index > 40) {
      status = 'EVALUATED';
    }

    records[key] = {
      id: key,
      staffId: staff.id,
      period: currentPeriod,
      scores,
      selfTotal,
      evaluatorTotal: status !== 'SUBMITTED' ? evalTotal : undefined,
      bghTotal: status === 'APPROVED' ? bghTotal : undefined,
      selfXepLoai,
      evaluatorXepLoai: status !== 'SUBMITTED' ? evalXepLoai : undefined,
      bghXepLoai: status === 'APPROVED' ? bghXepLoai : undefined,
      hasExemplaryProduct: hasProd,
      exemplaryDetail: hasProd ? 'Sản phẩm STEM tiêu biểu cấp trường; hướng dẫn HS tham gia KHKT / Bồi dưỡng HSG đạt kết quả tốt.' : '',
      status,
      submittedAt: '2026-10-24 08:30',
      evaluatedAt: status !== 'SUBMITTED' ? '2026-10-25 14:15' : undefined,
      evaluatedBy: staff.evaluatorName,
      approvedAt: status === 'APPROVED' ? '2026-10-26 16:00' : undefined,
      approvedBy: 'Nguyễn Hồ Bắc (Hiệu trưởng)',
      tựNhậnXetƯuĐiểm: 'Chấp hành nghiêm đường lối của Đảng, pháp luật Nhà nước. Tích cực đổi mới phương pháp giảng dạy, hoàn thành tốt kế hoạch bài dạy và ứng dụng CNTT, Học bạ số.',
      tựNhậnXetKhuyếtĐiểm: 'Cần tiếp tục đẩy mạnh nghiên cứu tài liệu chuyên sâu môn học và bồi dưỡng tiếng Anh chuyên ngành.',
      ýKiếnCuộcHọpTổ: 'Tập thể tổ chuyên môn nhất trí với kết quả tự chấm và đề xuất xếp loại thi đua.',
      ýKiếnCấpỦy: 'Đồng ý với đề xuất của tổ chuyên môn và kết quả theo dõi của nhà trường.',
    };
  });

  return records;
}

export function loadEvaluations(): Record<string, EvaluationRecord> {
  try {
    const data = localStorage.getItem(STORAGE_KEY_EVALUATIONS);
    if (!data) {
      const initial = getInitialEvaluations();
      saveEvaluations(initial);
      return initial;
    }
    return JSON.parse(data);
  } catch (err) {
    console.error('Error loading evaluations:', err);
    return getInitialEvaluations();
  }
}

export function saveEvaluations(data: Record<string, EvaluationRecord>) {
  try {
    localStorage.setItem(STORAGE_KEY_EVALUATIONS, JSON.stringify(data));
  } catch (err) {
    console.error('Error saving evaluations:', err);
  }
}

export function loadCurrentUser(): Staff {
  try {
    const data = localStorage.getItem(STORAGE_KEY_CURRENT_USER);
    if (data) {
      const parsed = JSON.parse(data);
      const found = STAFF_LIST.find((s) => s.id === parsed.id);
      if (found) return found;
    }
  } catch (err) {
    console.error(err);
  }
  // Default to Hiệu trưởng Nguyễn Hồ Bắc for full visibility
  return STAFF_LIST[0];
}

export function saveCurrentUser(staff: Staff) {
  try {
    localStorage.setItem(STORAGE_KEY_CURRENT_USER, JSON.stringify(staff));
  } catch (err) {
    console.error(err);
  }
}

export function loadCurrentPeriod(): string {
  try {
    const data = localStorage.getItem(STORAGE_KEY_CURRENT_PERIOD);
    if (data) return data;
  } catch (err) {
    console.error(err);
  }
  return 'thang_10_2026';
}

export function saveCurrentPeriod(periodId: string) {
  try {
    localStorage.setItem(STORAGE_KEY_CURRENT_PERIOD, periodId);
  } catch (err) {
    console.error(err);
  }
}

const STORAGE_KEY_PASSWORDS = 'ndc_custom_passwords_v1';

export function getCustomPasswords(): Record<string, string> {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_PASSWORDS);
    if (raw) return JSON.parse(raw);
  } catch (err) {
    console.error(err);
  }
  return {};
}

export function getStaffPassword(staffId: string): string {
  const custom = getCustomPasswords();
  if (custom[staffId]) return custom[staffId];
  const staff = STAFF_LIST.find((s) => s.id === staffId);
  return staff ? staff.defaultPassword : '123';
}

export function setStaffPassword(staffId: string, newPassword: string): void {
  const custom = getCustomPasswords();
  custom[staffId] = newPassword;
  localStorage.setItem(STORAGE_KEY_PASSWORDS, JSON.stringify(custom));
}

export function authenticateUser(loginId: string, passwordInput: string): Staff | null {
  const cleanId = loginId.trim().toLowerCase();
  const staff = STAFF_LIST.find(
    (s) =>
      s.username.toLowerCase() === cleanId ||
      s.sdt.trim() === cleanId ||
      s.cccd.trim() === cleanId ||
      s.id.toLowerCase() === cleanId
  );
  if (!staff) return null;

  const validPassword = getStaffPassword(staff.id);
  if (validPassword === passwordInput.trim()) {
    return staff;
  }
  return null;
}

// Strict Permission Handlers:
export function canAccessBghTab(staff: Staff): boolean {
  // Only Principal and Vice Principals can access BGH Approval Tab
  return staff.role === 'HIEU_TRUONG' || staff.role === 'PHO_HIEU_TRUONG';
}

export function canAccessGradingTab(staff: Staff): boolean {
  // Evaluators who have assigned subordinates to grade
  if (staff.role === 'HIEU_TRUONG' || staff.role === 'PHO_HIEU_TRUONG') return true;
  return STAFF_LIST.some((s) => s.evaluatorId === staff.id);
}

export function getAuthorizedStaffToGrade(evaluator: Staff): Staff[] {
  if (evaluator.role === 'HIEU_TRUONG') {
    // Hiệu trưởng can evaluate all or assigned PHT and TT
    return STAFF_LIST.filter((s) => s.evaluatorId === evaluator.id || s.evaluatorId === 'UBND_XA');
  }
  // Strictly restricted to personnel where evaluatorId === evaluator.id
  return STAFF_LIST.filter((s) => s.evaluatorId === evaluator.id);
}

export function canGradeSpecificStaff(evaluator: Staff, targetStaff: Staff): boolean {
  if (evaluator.role === 'HIEU_TRUONG') return true;
  return targetStaff.evaluatorId === evaluator.id;
}

export function canEditSelfEvaluation(currentUser: Staff, targetStaffId: string): boolean {
  // Strictly restricted: each person can ONLY edit their own self-evaluation score
  return currentUser.id === targetStaffId;
}

export function resetAllDataToDefault() {
  const initial = getInitialEvaluations();
  saveEvaluations(initial);
  saveCurrentUser(STAFF_LIST[0]);
  saveCurrentPeriod('thang_10_2026');
  localStorage.removeItem(STORAGE_KEY_PASSWORDS);
  return initial;
}

// Export evaluation table to Excel
export function exportEvaluationsToExcel(
  records: Record<string, EvaluationRecord>,
  periodLabel: string,
  staffList: Staff[]
) {
  const rows = staffList.map((s, idx) => {
    const key = `${s.id}_${loadCurrentPeriod()}`;
    const rec = records[key];

    return {
      'STT': s.stt,
      'Cơ sở': s.coSo,
      'Tổ CM/Bộ phận': s.toCM,
      'Họ và tên': s.hoTen,
      'Vị trí việc làm': s.viTri,
      'Môn dạy': s.monDay,
      'Người phụ trách chấm': s.evaluatorName,
      'Cá nhân tự chấm (Đ/100)': rec ? rec.selfTotal : 'Chưa nộp',
      'Tổ chấm (Đ/100)': rec && rec.evaluatorTotal !== undefined ? rec.evaluatorTotal : 'Chưa chấm',
      'BGH duyệt (Đ/100)': rec && rec.bghTotal !== undefined ? rec.bghTotal : 'Chưa duyệt',
      'Xếp loại BGH': rec && rec.bghXepLoai ? rec.bghXepLoai : (rec?.evaluatorXepLoai || rec?.selfXepLoai || 'Chưa đánh giá'),
      'Sản phẩm vượt mức/STEM/HSG': rec?.hasExemplaryProduct ? 'Có (' + (rec.exemplaryDetail || '') + ')' : 'Không',
      'Trạng thái': rec ? (rec.status === 'APPROVED' ? 'Đã duyệt' : rec.status === 'EVALUATED' ? 'Tổ đã chấm' : 'Đã nộp tự chấm') : 'Chưa nộp',
      'Số điện thoại': s.sdt,
      'CCCD': s.cccd,
      'Đảng viên': s.isDangVien ? 'Đảng viên' : '',
    };
  });

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Ket_Qua_Danh_Gia');

  // Generate file name
  const safePeriod = periodLabel.replace(/[\/\s-]/g, '_');
  const fileName = `Danh_Gia_Vien_Chuc_NguyenDinhChieu_${safePeriod}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
