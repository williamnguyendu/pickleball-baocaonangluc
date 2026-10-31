const SHEET_AUTH_ID = '1I_x7-ovdbjWGTkDiLv-VZsj7zdx4LLtQDRlJqbUkTHY';
const SHEET_DATA_ID = '1lzyBxzx6kHkmz0gsjlWJ4MTENACo-D60n9cytERV374';

function doPost(e) {
  try {
    let request = JSON.parse(e.postData.contents);
    let action = request.action;
    let data = request.data;
    
    if (action === 'login') return response(handleLogin(data));
    if (action === 'getReports') return response(getReports());
    if (action === 'getReportById') return response(getReportById(data.id));
    if (action === 'saveReport') return response(saveReport(data));
    if (action === 'deleteReport') return response(deleteReport(data.id));
    
    return response({ success: false, message: 'Action not found' });
  } catch (error) {
    return response({ success: false, message: error.toString() });
  }
}

function response(res) {
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function handleLogin(data) {
  let sheet = SpreadsheetApp.openById(SHEET_AUTH_ID).getSheetByName('TaiKhoan');
  let rows = sheet.getDataRange().getValues();
  let username = data.username;
  let md5Pass = data.password; // Đã mã hoá từ client
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === username) {
      let status = rows[i][4];
      let failedAttempts = rows[i][5] || 0; // Cột 6 lưu số lần sai (cần tạo cột này trong sheet TaiKhoan)
      
      if (status === 'Tạm khoá') return { success: false, message: 'Tài khoản đã bị tạm khoá.' };
      if (status === 'Không hoạt động') return { success: false, message: 'Tài khoản không hoạt động.' };
      
      if (rows[i][3] === md5Pass) {
        sheet.getRange(i + 1, 6).setValue(0); // Reset số lần sai
        return { success: true, user: { maTK: rows[i][0], tenTK: rows[i][1] } };
      } else {
        failedAttempts++;
        sheet.getRange(i + 1, 6).setValue(failedAttempts);
        if (failedAttempts >= 5) {
          sheet.getRange(i + 1, 5).setValue('Tạm khoá');
          return { success: false, message: 'Nhập sai 5 lần. Tài khoản đã bị khoá.' };
        }
        return { success: false, message: `Sai mật khẩu. Còn ${5 - failedAttempts} lần thử.` };
      }
    }
  }
  return { success: false, message: 'Tên đăng nhập không tồn tại.' };
}

function getReports() {
  let sheet = SpreadsheetApp.openById(SHEET_DATA_ID).getSheetByName('BaoCaoNangLuc');
  let rows = sheet.getDataRange().getDisplayValues();
  let result = [];
  for (let i = 1; i < rows.length; i++) {
    result.push({
      MaBaoCao: rows[i][0],
      TenHocVien: rows[i][1],
      SoDienThoai: rows[i][2],
      NgayDanhGia: rows[i][3]
    });
  }
  return { success: true, data: result.reverse() };
}

function getReportById(id) {
  let sheet = SpreadsheetApp.openById(SHEET_DATA_ID).getSheetByName('BaoCaoNangLuc');
  let rows = sheet.getDataRange().getDisplayValues();
  let headers = rows[0];
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      let report = {};
      headers.forEach((h, index) => { report[h] = rows[i][index]; });
      return { success: true, data: report };
    }
  }
  return { success: false, message: 'Không tìm thấy báo cáo' };
}

function saveReport(data) {
  let sheet = SpreadsheetApp.openById(SHEET_DATA_ID).getSheetByName('BaoCaoNangLuc');
  let rows = sheet.getDataRange().getValues();
  
  let rowData = [
    data.MaBaoCao, data.TenHocVien, data.SoDienThoai, data.NgayDanhGia, data.ThoiGianDaChoi,
    data.KinhNghiemTheThao, data.TanSuatHienTai, data.TayThuan, data.HinhThucChoi, data.MucTieuHocVien,
    data.ThoiLuongBuoiDanhGia, data.MucTieuUuTienHocVien, data.MucTieuKhac, data.TQ_NenTang,
    data.TQ_DiemManh, data.TQ_DiemYeu, data.CS_KyThuatCoBan, data.CS_DoOnDinh, data.CS_FootworkThangBang,
    data.CS_KiemSoatBongMem, data.CS_KhaNangTanCong, data.CS_PhongThuReset, data.CS_TuDuyChienThuat,
    data.CS_PhoiHopDanhDoi, data.CS_TamLyVaKiemSoatLoi, data.CS_KhaNangTiepThu
  ];

  if (data.MaBaoCao) {
    // Update
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.MaBaoCao) {
        sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        return { success: true, id: data.MaBaoCao, message: 'Cập nhật thành công' };
      }
    }
  } else {
    // Create new
    let lastIdNum = 0;
    if (rows.length > 1) {
      let lastId = rows[rows.length - 1][0]; // format BCxxxx
      lastIdNum = parseInt(lastId.replace('BC', '')) || 0;
    }
    let newId = 'BC' + ('000' + (lastIdNum + 1)).slice(-4);
    rowData[0] = newId;
    sheet.appendRow(rowData);
    return { success: true, id: newId, message: 'Tạo mới thành công' };
  }
}

function deleteReport(id) {
  let sheet = SpreadsheetApp.openById(SHEET_DATA_ID).getSheetByName('BaoCaoNangLuc');
  let rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Đã xoá báo cáo' };
    }
  }
  return { success: false, message: 'Không tìm thấy báo cáo' };
}