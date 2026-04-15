const XLSX = require('xlsx');
const path = require('path');

// ─── 학원 데이터 (학원 관리.png 기반) ───────────────────────────────────
const classes = [
  { id: 'cT5',  name: 'T5반',    level: '고급', teacher: '이선생', schedule: '월·수·금 오후 5시', max: 10 },
  { id: 'cT3',  name: 'T3반',    level: '중급', teacher: '이선생', schedule: '화·목 오후 5시',   max: 6  },
  { id: 'cT2',  name: 'T2반',    level: '초급', teacher: '박선생', schedule: '월·수 오후 4시',   max: 6  },
  { id: 'cSAT', name: '수능대비반', level: '특급', teacher: '박선생', schedule: '월~금 오후 7시',   max: 5  },
  { id: 'c3',   name: '3반',     level: '초급', teacher: '박선생', schedule: '화·목 오후 3시',   max: 5  },
];

const students = [
  // T5반
  { name: '강훈',   phone: '010-4922-3857', parent: '010-4922-3857', class: 'T5반',    level: '고급', status: '재원', date: '2025-03-01', memo: '형제: 강예준' },
  { name: '강예준', phone: '010-4922-3857', parent: '010-4922-3857', class: 'T5반',    level: '고급', status: '재원', date: '2025-03-01', memo: '형제: 강훈' },
  { name: '강현서', phone: '010-9885-7222', parent: '010-9885-7222', class: 'T5반',    level: '고급', status: '재원', date: '2025-03-01', memo: '형제: 강현준' },
  { name: '강현준', phone: '010-9885-7222', parent: '010-9885-7222', class: 'T5반',    level: '고급', status: '재원', date: '2025-03-01', memo: '형제: 강현서' },
  { name: '고윤우', phone: '010-9881-8549', parent: '010-9881-8549', class: 'T5반',    level: '고급', status: '재원', date: '2025-03-01', memo: '' },
  { name: '권수아', phone: '010-7174-9729', parent: '010-7174-9729', class: 'T5반',    level: '고급', status: '재원', date: '2025-03-01', memo: '' },
  { name: '김가원', phone: '010-5225-1773', parent: '010-5225-1773', class: 'T5반',    level: '고급', status: '재원', date: '2025-03-01', memo: '형제: 김규현' },
  { name: '김규현', phone: '010-5225-1773', parent: '010-5225-1773', class: 'T5반',    level: '고급', status: '재원', date: '2025-03-01', memo: '형제: 김가원' },
  // 수능대비반
  { name: '김규원', phone: '010-8642-9852', parent: '010-8642-9852', class: '수능대비반', level: '특급', status: '재원', date: '2025-03-01', memo: '형제: 김도우' },
  { name: '김도우', phone: '010-8642-9852', parent: '010-8642-9852', class: '수능대비반', level: '특급', status: '재원', date: '2025-03-01', memo: '형제: 김규원' },
  // T3반
  { name: '김도영', phone: '010-4760-5149', parent: '010-4760-5149', class: 'T3반',    level: '중급', status: '재원', date: '2025-03-01', memo: '' },
  { name: '김민찬', phone: '010-3771-9286', parent: '010-3771-9286', class: 'T3반',    level: '중급', status: '재원', date: '2025-03-01', memo: '' },
  // T2반
  { name: '김서이', phone: '010-3395-7957', parent: '010-3395-7957', class: 'T2반',    level: '초급', status: '재원', date: '2025-03-01', memo: '' },
  { name: '김소정', phone: '010-4741-1957', parent: '010-4741-1957', class: 'T2반',    level: '초급', status: '재원', date: '2025-03-01', memo: '' },
  // 3반
  { name: '김예슬', phone: '010-4557-7972', parent: '010-4557-7972', class: '3반',     level: '초급', status: '재원', date: '2025-03-01', memo: '' },
  { name: '김주원', phone: '010-8559-6250', parent: '010-8559-6250', class: '3반',     level: '초급', status: '재원', date: '2025-03-01', memo: '' },
];

// 원비 데이터 (2026-03 기준, 원 단위)
const fees = [
  { name: '강훈',   class: 'T5반',    month: '2026-03', amount: 270000, status: '미납',    paidDate: '',           note: '' },
  { name: '강예준', class: 'T5반',    month: '2026-03', amount: 270000, status: '미납',    paidDate: '',           note: '' },
  { name: '강현서', class: 'T5반',    month: '2026-03', amount: 270000, status: '미납',    paidDate: '',           note: '' },
  { name: '강현준', class: 'T5반',    month: '2026-03', amount: 280000, status: '납부완료', paidDate: '2026-03-23', note: '' },
  { name: '고윤우', class: 'T5반',    month: '2026-03', amount: 280000, status: '납부완료', paidDate: '2026-03-20', note: '' },
  { name: '권수아', class: 'T5반',    month: '2026-03', amount: 280000, status: '납부완료', paidDate: '2026-03-20', note: '' },
  { name: '김가원', class: 'T5반',    month: '2026-03', amount: 280000, status: '납부완료', paidDate: '2026-03-24', note: '' },
  { name: '김규현', class: 'T5반',    month: '2026-03', amount: 260000, status: '납부완료', paidDate: '2026-03-24', note: '교재비 별도 19,000' },
  { name: '김규원', class: '수능대비반', month: '2026-03', amount: 290000, status: '미납',    paidDate: '',           note: '형제: 김도우' },
  { name: '김도우', class: '수능대비반', month: '2026-03', amount: 280000, status: '미납',    paidDate: '',           note: '형제: 김규원' },
  { name: '김도영', class: 'T3반',    month: '2026-03', amount: 230000, status: '납부완료', paidDate: '2026-03-23', note: '교재비 별도 35,000' },
  { name: '김민찬', class: 'T3반',    month: '2026-03', amount: 270000, status: '납부완료', paidDate: '2026-03-05', note: '교재비 38,000 입금' },
  { name: '김서이', class: 'T2반',    month: '2026-03', amount: 300000, status: '납부완료', paidDate: '2026-03-25', note: '' },
  { name: '김소정', class: 'T2반',    month: '2026-03', amount: 300000, status: '미납',    paidDate: '',           note: '' },
  { name: '김예슬', class: '3반',     month: '2026-03', amount: 280000, status: '납부완료', paidDate: '2026-03-25', note: '' },
  { name: '김주원', class: '3반',     month: '2026-03', amount: 270000, status: '미납',    paidDate: '',           note: '교재비 별도 16,000' },
];

// 성적 데이터
const scores = [
  { name: '강훈',   class: 'T5반',    exam: '3월 월말 평가', date: '2026-03-28', score: 82, grade: 'B', note: '' },
  { name: '강예준', class: 'T5반',    exam: '3월 월말 평가', date: '2026-03-28', score: 88, grade: 'B', note: '' },
  { name: '강현서', class: 'T5반',    exam: '3월 월말 평가', date: '2026-03-28', score: 76, grade: 'C', note: '' },
  { name: '강현준', class: 'T5반',    exam: '3월 월말 평가', date: '2026-03-28', score: 91, grade: 'A', note: '' },
  { name: '고윤우', class: 'T5반',    exam: '3월 월말 평가', date: '2026-03-28', score: 85, grade: 'B', note: '' },
  { name: '권수아', class: 'T5반',    exam: '3월 월말 평가', date: '2026-03-28', score: 79, grade: 'C', note: '' },
  { name: '김가원', class: 'T5반',    exam: '3월 월말 평가', date: '2026-03-28', score: 93, grade: 'A', note: '' },
  { name: '김규현', class: 'T5반',    exam: '3월 월말 평가', date: '2026-03-28', score: 88, grade: 'B', note: '' },
  { name: '김규원', class: '수능대비반', exam: '3월 모의고사', date: '2026-03-28', score: 94, grade: 'A', note: '' },
  { name: '김도우', class: '수능대비반', exam: '3월 모의고사', date: '2026-03-28', score: 87, grade: 'B', note: '' },
  { name: '김도영', class: 'T3반',    exam: '3월 월말 평가', date: '2026-03-28', score: 74, grade: 'C', note: '' },
  { name: '김민찬', class: 'T3반',    exam: '3월 월말 평가', date: '2026-03-28', score: 81, grade: 'B', note: '' },
  { name: '김서이', class: 'T2반',    exam: '3월 월말 평가', date: '2026-03-28', score: 68, grade: 'D', note: '' },
  { name: '김소정', class: 'T2반',    exam: '3월 월말 평가', date: '2026-03-28', score: 72, grade: 'C', note: '' },
  { name: '김예슬', class: '3반',     exam: '3월 월말 평가', date: '2026-03-28', score: 77, grade: 'C', note: '' },
  { name: '김주원', class: '3반',     exam: '3월 월말 평가', date: '2026-03-28', score: 83, grade: 'B', note: '' },
];

// ─── 헬퍼: 셀 스타일 ──────────────────────────────────────────────────────
function headerStyle(bgColor) {
  return {
    font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
    fill: { fgColor: { rgb: bgColor }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top:    { style: 'thin', color: { rgb: 'BBCCDD' } },
      bottom: { style: 'thin', color: { rgb: 'BBCCDD' } },
      left:   { style: 'thin', color: { rgb: 'BBCCDD' } },
      right:  { style: 'thin', color: { rgb: 'BBCCDD' } },
    }
  };
}
function cellStyle(bold, align) {
  return {
    font: { bold: !!bold, sz: 10 },
    alignment: { horizontal: align || 'left', vertical: 'center' },
    border: {
      top:    { style: 'thin', color: { rgb: 'DDDDDD' } },
      bottom: { style: 'thin', color: { rgb: 'DDDDDD' } },
      left:   { style: 'thin', color: { rgb: 'DDDDDD' } },
      right:  { style: 'thin', color: { rgb: 'DDDDDD' } },
    }
  };
}
function paidStyle(status) {
  if (status === '납부완료') return { font: { color: { rgb: '1A7A4A' }, sz: 10 }, fill: { fgColor: { rgb: 'E6F5EE' }, patternType: 'solid' }, alignment: { horizontal: 'center' } };
  return { font: { color: { rgb: 'C03040' }, sz: 10 }, fill: { fgColor: { rgb: 'FDECEA' }, patternType: 'solid' }, alignment: { horizontal: 'center' } };
}

function makeCell(value, style) {
  return { v: value, t: typeof value === 'number' ? 'n' : 's', s: style };
}

// ─── Sheet 1: 학원 관리 (원비 납부 현황) ──────────────────────────────────
function createFeeSheet() {
  const ws = {};
  const headers = ['반', '이름', '전화번호', '청구월', '원비(원)', '납부상태', '납부일', '비고'];
  const colWidths = [10, 8, 14, 10, 12, 10, 12, 20];

  // 헤더
  headers.forEach((h, i) => {
    const cell = String.fromCharCode(65 + i) + '1';
    ws[cell] = makeCell(h, headerStyle('3B5FC0'));
  });

  // 데이터
  fees.forEach((f, ri) => {
    const row = ri + 2;
    ws[`A${row}`] = makeCell(f.class,   cellStyle(false));
    ws[`B${row}`] = makeCell(f.name,    cellStyle(true));
    ws[`C${row}`] = makeCell(f.phone || students.find(s => s.name === f.name)?.phone || '', cellStyle(false));
    ws[`D${row}`] = makeCell(f.month,   cellStyle(false, 'center'));
    ws[`E${row}`] = makeCell(f.amount,  { ...cellStyle(false, 'right'), numFmt: '#,##0' });
    ws[`F${row}`] = makeCell(f.status,  paidStyle(f.status));
    ws[`G${row}`] = makeCell(f.paidDate || '-', cellStyle(false, 'center'));
    ws[`H${row}`] = makeCell(f.note,    cellStyle(false));
  });

  // 합계 행
  const sumRow = fees.length + 2;
  ws[`A${sumRow}`] = makeCell('합계', { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center' } });
  ws[`E${sumRow}`] = makeCell(fees.reduce((a, f) => a + f.amount, 0), { font: { bold: true, color: { rgb: '3B5FC0' }, sz: 10 }, alignment: { horizontal: 'right' }, numFmt: '#,##0' });
  const paidSum = fees.filter(f => f.status === '납부완료').reduce((a, f) => a + f.amount, 0);
  ws[`F${sumRow}`] = makeCell(`납부 ${paidSum.toLocaleString()}원`, { font: { bold: true, color: { rgb: '1A7A4A' }, sz: 10 }, alignment: { horizontal: 'center' } });

  ws['!ref'] = `A1:H${sumRow}`;
  ws['!cols'] = colWidths.map(w => ({ wch: w }));
  ws['!rows'] = [{ hpt: 20 }];
  return ws;
}

// ─── Sheet 2: 학생 명단 ─────────────────────────────────────────────────
function createStudentSheet() {
  const ws = {};
  const headers = ['반', '이름', '레벨', '학생 연락처', '학부모 연락처', '등록일', '상태', '메모'];
  const colWidths = [10, 8, 6, 14, 14, 12, 6, 20];

  headers.forEach((h, i) => {
    ws[String.fromCharCode(65 + i) + '1'] = makeCell(h, headerStyle('2E7D32'));
  });

  students.forEach((s, ri) => {
    const row = ri + 2;
    ws[`A${row}`] = makeCell(s.class,  cellStyle(false, 'center'));
    ws[`B${row}`] = makeCell(s.name,   cellStyle(true));
    ws[`C${row}`] = makeCell(s.level,  cellStyle(false, 'center'));
    ws[`D${row}`] = makeCell(s.phone,  cellStyle(false));
    ws[`E${row}`] = makeCell(s.parent, cellStyle(false));
    ws[`F${row}`] = makeCell(s.date,   cellStyle(false, 'center'));
    ws[`G${row}`] = makeCell(s.status, cellStyle(false, 'center'));
    ws[`H${row}`] = makeCell(s.memo,   cellStyle(false));
  });

  ws['!ref'] = `A1:H${students.length + 1}`;
  ws['!cols'] = colWidths.map(w => ({ wch: w }));
  ws['!rows'] = [{ hpt: 20 }];
  return ws;
}

// ─── Sheet 3: 수업 관리 ─────────────────────────────────────────────────
function createClassSheet() {
  const ws = {};
  const headers = ['반 이름', '레벨', '담당 강사', '수업 시간', '정원', '현재 인원', '비고'];
  const colWidths = [12, 8, 10, 16, 6, 8, 20];

  headers.forEach((h, i) => {
    ws[String.fromCharCode(65 + i) + '1'] = makeCell(h, headerStyle('6A1E8A'));
  });

  classes.forEach((c, ri) => {
    const row = ri + 2;
    const enrolled = students.filter(s => s.class === c.name && s.status === '재원').length;
    ws[`A${row}`] = makeCell(c.name,     cellStyle(true));
    ws[`B${row}`] = makeCell(c.level,    cellStyle(false, 'center'));
    ws[`C${row}`] = makeCell(c.teacher,  cellStyle(false));
    ws[`D${row}`] = makeCell(c.schedule, cellStyle(false));
    ws[`E${row}`] = makeCell(c.max,      cellStyle(false, 'center'));
    ws[`F${row}`] = makeCell(enrolled,   cellStyle(false, 'center'));
    ws[`G${row}`] = makeCell('',         cellStyle(false));
  });

  ws['!ref'] = `A1:G${classes.length + 1}`;
  ws['!cols'] = colWidths.map(w => ({ wch: w }));
  ws['!rows'] = [{ hpt: 20 }];
  return ws;
}

// ─── Sheet 4: 성적 관리 ─────────────────────────────────────────────────
function createScoreSheet() {
  const ws = {};
  const headers = ['반', '이름', '시험명', '시험일', '점수', '등급', '메모'];
  const colWidths = [10, 8, 16, 12, 6, 6, 20];

  headers.forEach((h, i) => {
    ws[String.fromCharCode(65 + i) + '1'] = makeCell(h, headerStyle('B45309'));
  });

  scores.forEach((s, ri) => {
    const row = ri + 2;
    const gradeColor = s.score >= 90 ? '1A7A4A' : s.score >= 80 ? '1A56A8' : s.score >= 70 ? 'B45309' : 'C03040';
    ws[`A${row}`] = makeCell(s.class, cellStyle(false, 'center'));
    ws[`B${row}`] = makeCell(s.name,  cellStyle(true));
    ws[`C${row}`] = makeCell(s.exam,  cellStyle(false));
    ws[`D${row}`] = makeCell(s.date,  cellStyle(false, 'center'));
    ws[`E${row}`] = makeCell(s.score, { font: { bold: true, color: { rgb: gradeColor }, sz: 10 }, alignment: { horizontal: 'center' } });
    ws[`F${row}`] = makeCell(s.grade, { font: { bold: true, color: { rgb: gradeColor }, sz: 10 }, alignment: { horizontal: 'center' } });
    ws[`G${row}`] = makeCell(s.note,  cellStyle(false));
  });

  const avg = Math.round(scores.reduce((a, s) => a + s.score, 0) / scores.length);
  const sumRow = scores.length + 2;
  ws[`D${sumRow}`] = makeCell('평균', { font: { bold: true, sz: 10 }, alignment: { horizontal: 'right' } });
  ws[`E${sumRow}`] = makeCell(avg, { font: { bold: true, color: { rgb: '3B5FC0' }, sz: 10 }, alignment: { horizontal: 'center' } });

  ws['!ref'] = `A1:G${sumRow}`;
  ws['!cols'] = colWidths.map(w => ({ wch: w }));
  ws['!rows'] = [{ hpt: 20 }];
  return ws;
}

// ─── Workbook 1: DS_English_학원관리.xlsx ──────────────────────────────
const wb1 = XLSX.utils.book_new();
wb1.Props = {
  Title: 'DS English 학원 관리',
  Subject: '원비·학생·수업·성적 관리',
  Author: 'DS English',
  CreatedDate: new Date(),
};
XLSX.utils.book_append_sheet(wb1, createFeeSheet(),     '원비 납부 현황');
XLSX.utils.book_append_sheet(wb1, createStudentSheet(), '학생 명단');
XLSX.utils.book_append_sheet(wb1, createClassSheet(),   '수업 관리');
XLSX.utils.book_append_sheet(wb1, createScoreSheet(),   '성적 관리');

const out1 = path.join(__dirname, 'DS_English_학원관리.xlsx');
XLSX.writeFile(wb1, out1, { bookType: 'xlsx', type: 'buffer' });
console.log('✅ 생성 완료:', out1);

// ─── Workbook 2: DS_English_데이터입력양식.xlsx (입력 템플릿) ──────────
function createInputTemplate() {
  const wb2 = XLSX.utils.book_new();

  // 학생 입력 템플릿
  const studentTpl = XLSX.utils.aoa_to_sheet([
    ['이름', '수업반', '레벨', '학생 연락처', '학부모 연락처', '등록일 (YYYY-MM-DD)', '상태', '메모'],
    ...students.map(s => [s.name, s.class, s.level, s.phone, s.parent, s.date, s.status, s.memo]),
    // 빈 입력 행 5개
    ...Array(5).fill(['', '', '', '', '', '', '재원', '']),
  ]);
  studentTpl['!cols'] = [8, 12, 6, 14, 14, 18, 8, 20].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb2, studentTpl, '학생 입력');

  // 원비 입력 템플릿
  const feeTpl = XLSX.utils.aoa_to_sheet([
    ['학생 이름', '수업반', '청구월 (YYYY-MM)', '금액 (원)', '납부상태', '납부일 (YYYY-MM-DD)', '비고'],
    ...fees.map(f => [f.name, f.class, f.month, f.amount, f.status, f.paidDate, f.note]),
    ...Array(5).fill(['', '', '', '', '미납', '', '']),
  ]);
  feeTpl['!cols'] = [10, 12, 18, 12, 10, 18, 20].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb2, feeTpl, '원비 입력');

  // 수업 입력 템플릿
  const classTpl = XLSX.utils.aoa_to_sheet([
    ['반 이름', '레벨', '담당 강사', '수업 시간', '정원', '교재'],
    ...classes.map(c => [c.name, c.level, c.teacher, c.schedule, c.max, '']),
    ...Array(3).fill(['', '', '', '', 10, '']),
  ]);
  classTpl['!cols'] = [12, 8, 10, 18, 6, 20].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb2, classTpl, '수업 입력');

  // 성적 입력 템플릿
  const scoreTpl = XLSX.utils.aoa_to_sheet([
    ['학생 이름', '수업반', '시험명', '시험일 (YYYY-MM-DD)', '점수 (0-100)', '메모'],
    ...scores.map(s => [s.name, s.class, s.exam, s.date, s.score, s.note]),
    ...Array(5).fill(['', '', '', '', '', '']),
  ]);
  scoreTpl['!cols'] = [10, 12, 16, 18, 14, 20].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb2, scoreTpl, '성적 입력');

  const out2 = path.join(__dirname, 'DS_English_데이터입력양식.xlsx');
  XLSX.writeFile(wb2, out2, { bookType: 'xlsx', type: 'buffer' });
  console.log('✅ 생성 완료:', out2);
}

createInputTemplate();
