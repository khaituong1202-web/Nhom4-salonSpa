// NGÂN Spa & Salon - Contact Form Management System

// Data storage key
const DATA_KEY = 'salon_bookings';

// Initialize sample data if empty
function initializeSampleData() {
  const existingData = localStorage.getItem(DATA_KEY);
  
  if (!existingData) {
    const sampleData = [
      {
        id: 1,
        name: "Nguyễn Thị A",
        email: "nguyenthia@email.com",
        phone: "0912345678",
        service: "Cắt tóc",
        notes: "Muốn cắt tóc ngắn",
        date: new Date().toISOString(),
        status: "Chờ xác nhận"
      },
      {
        id: 2,
        name: "Trần Văn B",
        email: "tranvanb@email.com",
        phone: "0987654321",
        service: "Massage",
        notes: "Đau vai gáy",
        date: new Date().toISOString(),
        status: "Đã xác nhận"
      }
    ];
    
    localStorage.setItem(DATA_KEY, JSON.stringify(sampleData));
  }
}

// Get all bookings from localStorage
function getBookings() {
  const data = localStorage.getItem(DATA_KEY);
  return data ? JSON.parse(data) : [];
}

// Save booking to localStorage
function saveBooking(formData) {
  const bookings = getBookings();
  
  const newBooking = {
    id: bookings.length > 0 ? Math.max(...bookings.map(b => b.id)) + 1 : 1,
    name: formData.name,
    email: formData.email,
    phone: formData.phone,
    service: formData.service,
    notes: formData.notes || '',
    date: new Date().toISOString(),
    status: 'Chờ xác nhận'
  };
  
  bookings.push(newBooking);
  localStorage.setItem(DATA_KEY, JSON.stringify(bookings));
  return newBooking;
}

// Export bookings to Excel
function exportToExcel() {
  const bookings = getBookings();
  
  if (bookings.length === 0) {
    alert('Không có dữ liệu để xuất!');
    return;
  }
  
  // Prepare data for Excel
  const exportData = bookings.map(booking => ({
    'ID': booking.id,
    'Họ và tên': booking.name,
    'Email': booking.email,
    'Số điện thoại': booking.phone,
    'Dịch vụ': booking.service,
    'Ghi chú': booking.notes,
    'Ngày đặt': new Date(booking.date).toLocaleString('vi-VN'),
    'Trạng thái': booking.status
  }));
  
  // Create workbook and worksheet
  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Bookings');
  
  // Style the header row
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const address = XLSX.utils.encode_col(C) + '1';
    if (!ws[address]) continue;
    ws[address].s = {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: 'B5564A' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
  }
  
  // Auto-adjust column widths
  const colWidths = [];
  for (let C = range.s.c; C <= range.e.c; ++C) {
    colWidths.push({ wch: 20 });
  }
  ws['!cols'] = colWidths;
  
  // Generate filename
  const filename = `NGAN_Spa_Bookings_${new Date().toISOString().split('T')[0]}.xlsx`;
  
  // Save file
  XLSX.writeFile(wb, filename);
}

// Clear all data
function clearAllData() {
  if (confirm('Bạn chắc chắn muốn xóa tất cả dữ liệu? Hành động này không thể hoàn tác!')) {
    localStorage.removeItem(DATA_KEY);
    initializeSampleData();
    showMessage('Dữ liệu đã được xóa và khôi phục dữ liệu mẫu', 'info');
    document.getElementById('contactForm').reset();
  }
}

// Show message
function showMessage(msg, type) {
  const messageEl = document.getElementById('message');
  if (!messageEl) return;
  
  messageEl.textContent = msg;
  messageEl.className = `form-message form-message-${type}`;
  messageEl.style.display = 'block';
  
  if (type === 'success') {
    setTimeout(() => {
      messageEl.style.display = 'none';
    }, 3000);
  }
}

// Initialize form
function initializeForm() {
  const form = document.getElementById('contactForm');
  if (!form) return;
  
  form.addEventListener('submit', function(e) {
    e.preventDefault();
    
    const formData = {
      name: document.getElementById('name').value.trim(),
      email: document.getElementById('email').value.trim(),
      phone: document.getElementById('phone').value.trim(),
      service: document.getElementById('service').value,
      notes: document.getElementById('notes').value.trim()
    };
    
    // Validation
    if (!formData.name || !formData.email || !formData.phone || !formData.service) {
      showMessage('⚠️ Vui lòng điền đầy đủ các trường bắt buộc!', 'error');
      return;
    }
    
    if (!isValidEmail(formData.email)) {
      showMessage('⚠️ Email không hợp lệ!', 'error');
      return;
    }
    
    if (!isValidPhone(formData.phone)) {
      showMessage('⚠️ Số điện thoại không hợp lệ!', 'error');
      return;
    }
    
    // Save booking
    try {
      saveBooking(formData);
      showMessage('✅ Đặt lịch thành công! Chúng tôi sẽ liên hệ với bạn sớm.', 'success');
      form.reset();
    } catch (error) {
      showMessage('❌ Có lỗi xảy ra. Vui lòng thử lại!', 'error');
      console.error('Error:', error);
    }
  });
  
  // Export button
  const exportBtn = document.getElementById('exportBtn');
  if (exportBtn) {
    exportBtn.addEventListener('click', exportToExcel);
  }
  
  // Clear button
  const clearBtn = document.getElementById('clearBtn');
  if (clearBtn) {
    clearBtn.addEventListener('click', clearAllData);
  }
}

// Email validation
function isValidEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

// Phone validation (Vietnamese phone format)
function isValidPhone(phone) {
  const re = /^(0|\+84)[0-9]{9,10}$/;
  return re.test(phone);
}

// Initialize newsletter
function initializeNewsletter() {
  // Newsletter functionality - can be extended later
  console.log('📰 Newsletter feature ready');
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
  // Initialize sample data
  initializeSampleData();
  
  // Initialize form
  initializeForm();

  // Initialize newsletter
  initializeNewsletter();
  
  console.log('✅ NGÂN Spa & Salon - Hệ thống quản lý đặt lịch đã khởi động!');
  console.log('📊 Dữ liệu hiện tại:', getBookings());
});
