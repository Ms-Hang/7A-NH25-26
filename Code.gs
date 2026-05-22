/**
 * PHIẾU BÁO KẾT QUẢ HỌC TẬP HỌC SINH - LỚP 7A
 * Trường THCS Yên Lương - Năm học 2025 - 2026
 * 
 * HƯỚNG DẪN SỬ DỤNG:
 * 1. Upload file Excel "so_diem_tong_ket_lop_7a.xlsx" lên Google Drive
 * 2. Mở file đó bằng Google Sheets (File > Save as Google Sheets)
 * 3. Copy ID của Google Sheet (phần giữa /d/ và /edit trong URL)
 * 4. Dán ID vào biến SPREADSHEET_ID bên dưới
 * 5. Tạo Apps Script project mới (Extensions > Apps Script)
 * 6. Copy Code.gs và Index.html vào project
 * 7. Deploy > New deployment > Web app
 */

// ===== CẤU HÌNH =====
const SPREADSHEET_ID = '1DKLt7Q2DOMGnXuVmTBWtqKFa6Q1Itd9oVZC5ADlSJwQ'; // <-- Thay bằng ID Google Sheet của bạn
const SHEET_HK1 = 'Hoc ki 1';
const SHEET_HK2 = 'Hoc ki 2';
const SHEET_CN  = 'Ca nam';
const DATA_START_ROW = 8; // Dữ liệu bắt đầu từ hàng 8
const HEADER_ROW = 6;     // Hàng tiêu đề

// Thông tin trường
const SCHOOL_INFO = {
  ubanName: 'ỦY BAN NHÂN DÂN XÃ VẠN THẮNG',
  schoolName: 'TRƯỜNG THCS YÊN LƯƠNG',
  yearLabel: 'Năm học: 2025 - 2026',
  gradeLabel: 'Khối: 7',
  classLabel: 'Lớp: 7A',
  gvcn: 'Đỗ Thị Hằng'
};

/**
 * Serve the web app
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tra cứu điểm - Lớp 7A - THCS Yên Lương')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Lấy danh sách tên học sinh để autocomplete
 */
function getStudentList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CN);
  const lastRow = sheet.getLastRow();
  const students = [];

  for (let r = DATA_START_ROW; r <= lastRow; r++) {
    const stt = sheet.getRange(r, 1).getValue();
    if (stt === '' || isNaN(stt)) break;

    const name = sheet.getRange(r, 4).getValue();
    const maHS = sheet.getRange(r, 2).getValue();
    students.push({
      stt: stt,
      name: name.toString().trim(),
      maHS: maHS.toString().trim()
    });
  }

  return students;
}

/**
 * Lấy thông tin trường
 */
function getSchoolInfo() {
  return SCHOOL_INFO;
}

/**
 * Tìm kiếm học sinh theo tên hoặc mã
 */
function searchStudent(keyword) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const kw = keyword.toString().trim().toLowerCase();

  // Lấy dữ liệu từ 3 sheet
  const hk1Data = getSheetData(ss, SHEET_HK1);
  const hk2Data = getSheetData(ss, SHEET_HK2);
  const cnData  = getSheetData(ss, SHEET_CN);

  // Tìm học sinh trong sheet Cả năm
  let found = null;
  let foundIndex = -1;

  for (let i = 0; i < cnData.length; i++) {
    const row = cnData[i];
    const name = row.name.toLowerCase();
    const maHS = row.maHS.toLowerCase();

    if (name === kw || maHS === kw || name.includes(kw)) {
      found = row;
      foundIndex = i;
      break;
    }
  }

  if (!found) {
    return { success: false, message: 'Không tìm thấy học sinh với từ khóa: ' + keyword };
  }

  // Tìm dữ liệu HK1, HK2 tương ứng (theo mã HS)
  const hk1Row = hk1Data.find(r => r.maHS === found.maHS);
  const hk2Row = hk2Data.find(r => r.maHS === found.maHS);

  // Tính hạng lớp cho từng kỳ
  const hk1Rank = calculateRank(hk1Data, found.maHS);
  const hk2Rank = calculateRank(hk2Data, found.maHS);
  const cnRank  = calculateRank(cnData, found.maHS);

  const result = {
    success: true,
    schoolInfo: SCHOOL_INFO,
    student: {
      name: found.name,
      maHS: found.maHS,
      ngaySinh: found.ngaySinh
    },
    hk1: hk1Row ? formatPeriodData(hk1Row, hk1Rank) : null,
    hk2: hk2Row ? formatPeriodData(hk2Row, hk2Rank) : null,
    cn:  formatCaNamData(found, cnRank),
    nghiHoc: {
      coPhep: found.nghiP || 0,
      khongPhep: found.nghiK || 0
    },
    ketQuaCuoiNam: found.ghiChu || '',
    danhHieu: found.danhHieu || ''
  };

  return result;
}

/**
 * Đọc dữ liệu từ sheet
 */
function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const data = [];
  const isCaNam = (sheetName === SHEET_CN);

  for (let r = DATA_START_ROW; r <= lastRow; r++) {
    const stt = sheet.getRange(r, 1).getValue();
    if (stt === '' || isNaN(stt)) break;

    const rowData = {
      stt: stt,
      maHS: sheet.getRange(r, 2).getValue().toString().trim(),
      soDinhDanh: sheet.getRange(r, 3).getValue().toString().trim(),
      name: sheet.getRange(r, 4).getValue().toString().trim(),
      ngaySinh: formatDate(sheet.getRange(r, 5).getValue()),
      toan: sheet.getRange(r, 6).getValue(),
      lsdl: sheet.getRange(r, 7).getValue(),
      khtn: sheet.getRange(r, 8).getValue(),
      tin: sheet.getRange(r, 9).getValue(),
      van: sheet.getRange(r, 10).getValue(),
      tiengAnh: sheet.getRange(r, 11).getValue(),
      gdcd: sheet.getRange(r, 12).getValue(),
      congNghe: sheet.getRange(r, 13).getValue(),
      gdtc: sheet.getRange(r, 14).getValue().toString().trim(),
      ngheThuat: sheet.getRange(r, 15).getValue().toString().trim(),
      ndgdcdp: sheet.getRange(r, 16).getValue().toString().trim(),
      hdtnhn: sheet.getRange(r, 17).getValue().toString().trim(),
      kqht: sheet.getRange(r, 18).getValue().toString().trim(),
      kqrl: sheet.getRange(r, 19).getValue().toString().trim(),
    };

    if (isCaNam) {
      rowData.nghiP = sheet.getRange(r, 21).getValue();
      rowData.nghiK = sheet.getRange(r, 22).getValue();
      rowData.danhHieu = sheet.getRange(r, 24).getValue().toString().trim();
      rowData.ghiChu = sheet.getRange(r, 25).getValue().toString().trim();
    } else {
      rowData.nghiP = sheet.getRange(r, 20).getValue();
      rowData.nghiK = sheet.getRange(r, 21).getValue();
    }

    // Tính điểm trung bình các môn số
    const numericSubjects = [rowData.toan, rowData.lsdl, rowData.khtn, rowData.tin,
                             rowData.van, rowData.tiengAnh, rowData.gdcd, rowData.congNghe];
    const validScores = numericSubjects.filter(s => !isNaN(s) && s !== '');
    rowData.avg = validScores.length > 0 
      ? validScores.reduce((a, b) => a + parseFloat(b), 0) / validScores.length 
      : 0;

    data.push(rowData);
  }

  return data;
}

/**
 * Tính hạng (ranking) theo điểm trung bình
 */
function calculateRank(allData, targetMaHS) {
  // Sort theo avg giảm dần
  const sorted = [...allData].sort((a, b) => b.avg - a.avg);
  const rank = sorted.findIndex(s => s.maHS === targetMaHS) + 1;
  const total = sorted.length;
  return { rank: rank, total: total };
}

/**
 * Format dữ liệu một kỳ
 */
function formatPeriodData(row, rankInfo) {
  return {
    toan: formatScore(row.toan),
    lsdl: formatScore(row.lsdl),
    khtn: formatScore(row.khtn),
    tin: formatScore(row.tin),
    van: formatScore(row.van),
    tiengAnh: formatScore(row.tiengAnh),
    gdcd: formatScore(row.gdcd),
    congNghe: formatScore(row.congNghe),
    gdtc: row.gdtc,
    ngheThuat: row.ngheThuat,
    ndgdcdp: row.ndgdcdp,
    hdtnhn: row.hdtnhn,
    kqht: row.kqht,
    kqrl: row.kqrl,
    hangLop: rankInfo.rank,
    hangKhoi: '' // Không có dữ liệu khối
  };
}

/**
 * Format dữ liệu cả năm
 */
function formatCaNamData(row, rankInfo) {
  return {
    toan: formatScore(row.toan),
    lsdl: formatScore(row.lsdl),
    khtn: formatScore(row.khtn),
    tin: formatScore(row.tin),
    van: formatScore(row.van),
    tiengAnh: formatScore(row.tiengAnh),
    gdcd: formatScore(row.gdcd),
    congNghe: formatScore(row.congNghe),
    gdtc: row.gdtc,
    ngheThuat: row.ngheThuat,
    ndgdcdp: row.ndgdcdp,
    hdtnhn: row.hdtnhn,
    kqht: row.kqht,
    kqrl: row.kqrl,
    hangLop: rankInfo.rank,
    hangKhoi: '', // Không có dữ liệu khối
    danhHieu: row.danhHieu,
    ghiChu: row.ghiChu
  };
}

/**
 * Format điểm số
 */
function formatScore(score) {
  if (score === '' || score === null || score === undefined) return '';
  const num = parseFloat(score);
  if (isNaN(num)) return score.toString();
  return num % 1 === 0 ? num.toFixed(1) : parseFloat(num.toFixed(1)).toString();
}

/**
 * Format ngày tháng
 */
function formatDate(dateVal) {
  if (!dateVal) return '';
  if (dateVal instanceof Date) {
    const d = dateVal.getDate().toString().padStart(2, '0');
    const m = (dateVal.getMonth() + 1).toString().padStart(2, '0');
    const y = dateVal.getFullYear();
    return d + '/' + m + '/' + y;
  }
  return dateVal.toString();
}
