// QPapers — main.js

// Auto-dismiss alerts after 5s
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.qp-alert').forEach(alert => {
    setTimeout(() => {
      const bsAlert = bootstrap.Alert.getOrCreateInstance(alert);
      if (bsAlert) bsAlert.close();
    }, 5000);
  });
});

// Password toggle
function togglePw(inputId, btn) {
  const input = document.getElementById(inputId);
  const icon = btn.querySelector('i');
  if (input.type === 'password') {
    input.type = 'text';
    icon.className = 'bi bi-eye-slash';
  } else {
    input.type = 'password';
    icon.className = 'bi bi-eye';
  }
}

// Confirm delete helper (used inline as well)
function confirmDelete(msg) {
  return confirm(msg || 'Are you sure you want to delete this item?');
}
