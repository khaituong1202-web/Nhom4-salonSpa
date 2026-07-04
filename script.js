// NGÂN Spa & Salon - Contact Form Management System

// Data storage key
const DATA_KEY = 'salon_bookings';

// Toast System
function showToast(message, type = 'info', duration = 4000) {
    // Remove existing toast container if any
    let container = document.querySelector('.toast-container');
    if (!container) {
        container = document.createElement('div');
        container.className = 'toast-container';
        document.body.appendChild(container);
    }

    // Create toast element
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    // Icons mapping
    const icons = {
        success: '✅',
        error: '❌',
        warning: '⚠️',
        info: 'ℹ️'
    };
    
    const titles = {
        success: 'Thành công!',
        error: 'Lỗi!',
        warning: 'Cảnh báo!',
        info: 'Thông báo'
    };

    toast.innerHTML = `
        <span class="toast-icon">${icons[type] || 'ℹ️'}</span>
        <div class="toast-content">
            <div class="toast-title">${titles[type] || 'Thông báo'}</div>
            <div class="toast-message">${message}</div>
        </div>
        <button class="toast-close" onclick="this.parentElement.remove()">✕</button>
        <div class="toast-progress"></div>
    `;

    container.appendChild(toast);

    // Auto remove after duration
    setTimeout(() => {
        toast.classList.add('hide');
        setTimeout(() => {
            toast.remove();
            // Remove container if empty
            if (container.children.length === 0) {
                container.remove();
            }
        }, 500);
    }, duration);
}

// Initialize sample data if empty
function initializeSampleData() {
  const existingData = localStorage.getItem(DATA_KEY);
  
  if (!existingData) {
    const sampleData = [
      {
        id: 1,
        name: "Nguyễn Thị A",
        email: "nguyena@email.com",
        phone: "0912345678",
        service: "Cắt tóc nữ",
        notes: "Muốn cắt ngang vai",
        date: new Date().toISOString(),
        status: "Chờ xác nhận"
      },
      {
        id: 2,
        name: "Trần Văn B",
        email: "tranb@email.com",
        phone: "0987654321",
        service: "Gội đầu dưỡng sinh",
        notes: "",
        date: new Date(Date.now() - 86400000).toISOString(),
        status: "Đã xác nhận"
      }
    ];
    
    localStorage.setItem(DATA_KEY, JSON.stringify(sampleData));
  }
}

// Get all bookings from localStorage
function getBookings() {
  try {
    const data = localStorage.getItem(DATA_KEY);
    return data ? JSON.parse(data) : [];
  } catch (error) {
    console.error('Error reading data:', error);
    return [];
  }
}

// Save booking to localStorage
function saveBooking(formData) {
  const bookings = getBookings();
  
  const newBooking = {
    id: bookings.length > 0 ? Math.max(...bookings.map(b => b.id)) + 1 : 1,
    name: formData.name.trim(),
    email: formData.email.trim(),
    phone: formData.phone.trim(),
    service: formData.service,
    notes: formData.notes ? formData.notes.trim() : '',
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
    showToast('Không có dữ liệu để xuất!', 'warning');
    return;
  }
  
  try {
    // Prepare data for Excel
    const exportData = bookings.map(booking => ({
      'ID': booking.id,
      'Họ và tên': booking.name,
      'Email': booking.email,
      'Số điện thoại': booking.phone,
      'Dịch vụ': booking.service,
      'Ghi chú': booking.notes || '',
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
    showToast('Xuất file Excel thành công!', 'success');
  } catch (error) {
    console.error('Export error:', error);
    showToast('Lỗi xuất file Excel. Vui lòng thử lại!', 'error');
  }
}

// Clear all data
function clearAllData() {
  if (confirm('Bạn chắc chắn muốn xóa tất cả dữ liệu? Hành động này không thể hoàn tác!')) {
    localStorage.removeItem(DATA_KEY);
    initializeSampleData();
    showToast('Dữ liệu đã được xóa và khôi phục dữ liệu mẫu', 'info');
    const form = document.getElementById('contactForm');
    if (form) form.reset();
  }
}

// Initialize form
function initializeForm() {
  const form = document.getElementById('contactForm');
  
  if (!form) {
    console.warn('Contact form not found');
    return;
  }
  
  form.addEventListener('submit', function(e) {
    e.preventDefault();
    
    const name = document.getElementById('name');
    const email = document.getElementById('email');
    const phone = document.getElementById('phone');
    const service = document.getElementById('service');
    const notes = document.getElementById('notes');
    
    const formData = {
      name: name ? name.value.trim() : '',
      email: email ? email.value.trim() : '',
      phone: phone ? phone.value.trim() : '',
      service: service ? service.value : '',
      notes: notes ? notes.value.trim() : ''
    };
    
    // Validation
    if (!formData.name || !formData.email || !formData.phone || !formData.service) {
      showToast('Vui lòng điền đầy đủ các trường bắt buộc!', 'error');
      return;
    }
    
    if (!isValidEmail(formData.email)) {
      showToast('Email không hợp lệ! Vui lòng nhập đúng định dạng (example@email.com)', 'error');
      return;
    }
    
    if (!isValidPhone(formData.phone)) {
      showToast('Số điện thoại không hợp lệ! Vui lòng nhập số điện thoại Việt Nam (10-11 số)', 'error');
      return;
    }
    
    // Save booking
    try {
      saveBooking(formData);
      showToast('Đặt lịch thành công! Chúng tôi sẽ liên hệ với bạn sớm.', 'success');
      form.reset();
    } catch (error) {
      showToast('Có lỗi xảy ra khi lưu dữ liệu. Vui lòng thử lại!', 'error');
      console.error('Save error:', error);
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
  // Supports: 09xxxxxxxx, 08xxxxxxxx, 07xxxxxxxx, 03xxxxxxxx, 05xxxxxxxx, +849xxxxxxxx
  const re = /^(0[3|5|7|8|9][0-9]{8}|(\+84)[0-9]{9,10})$/;
  return re.test(phone);
}

// Initialize back-to-top button
function initializeBackToTop() {
  const backToTopBtn = document.getElementById('backToTop');
  if (backToTopBtn) {
    window.addEventListener('scroll', function() {
      if (window.scrollY > 300) {
        backToTopBtn.style.display = 'flex';
      } else {
        backToTopBtn.style.display = 'none';
      }
    });
    
    backToTopBtn.addEventListener('click', function() {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }
}

// Initialize newsletter
function initializeNewsletter() {
  const newsletterForm = document.getElementById('newsletterForm');
  if (newsletterForm) {
    newsletterForm.addEventListener('submit', function(e) {
      e.preventDefault();
      const emailInput = this.querySelector('input[type="email"]');
      if (emailInput && isValidEmail(emailInput.value)) {
        showToast('Đăng ký nhận tin thành công!', 'success');
        emailInput.value = '';
      } else {
        showToast('Vui lòng nhập email hợp lệ!', 'error');
      }
    });
  }
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
  try {
    // Initialize sample data
    initializeSampleData();
    
    // Initialize form
    initializeForm();
    
    // Initialize back-to-top button
    initializeBackToTop();
    
    // Initialize newsletter
    initializeNewsletter();
    
    console.log('✅ NGÂN Spa & Salon - Hệ thống quản lý đặt lịch đã khởi động!');
    console.log('📊 Dữ liệu hiện tại:', getBookings());
  } catch (error) {
    console.error('Initialization error:', error);
  }
});
