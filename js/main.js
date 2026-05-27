// ========================================
// 台北市鍋爐壓力容器協會 - 主要 JavaScript
// ========================================

(function () {
  'use strict';

  // ---- Mobile Menu ----
  const menuToggle = document.getElementById('menuToggle');
  const mainNav = document.getElementById('mainNav');
  const navClose = document.getElementById('navClose');

  if (menuToggle && mainNav) {
    menuToggle.addEventListener('click', () => {
      mainNav.classList.toggle('open');
      document.body.style.overflow = mainNav.classList.contains('open') ? 'hidden' : '';
    });
  }
  if (navClose) {
    navClose.addEventListener('click', () => {
      mainNav.classList.remove('open');
      document.body.style.overflow = '';
    });
  }

  // ---- Scroll to Top ----
  const topBtn = document.getElementById('floatTopBtn');
  window.addEventListener('scroll', () => {
    if (topBtn) topBtn.classList.toggle('show', window.scrollY > 400);
  });
  if (topBtn) topBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));

  // ---- Registration Modal ----
  const regModal = document.getElementById('regModal');
  const regBtns = document.querySelectorAll('[data-action="register"]');
  const modalClose = document.getElementById('modalClose');
  const modalCancel = document.getElementById('modalCancel');
  const courseSelect = document.getElementById('modalCourseSelect');

  function openModal(courseName) {
    if (!regModal) return;
    if (courseName && courseSelect) {
      // Try to select matching option
      const opts = courseSelect.options;
      for (let i = 0; i < opts.length; i++) {
        if (opts[i].text.includes(courseName) || opts[i].value.includes(courseName)) {
          courseSelect.selectedIndex = i;
          break;
        }
      }
    }
    regModal.classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  function closeModal() {
    if (!regModal) return;
    regModal.classList.remove('active');
    document.body.style.overflow = '';
  }

  regBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const courseName = btn.dataset.course || '';
      openModal(courseName);
    });
  });

  if (modalClose) modalClose.addEventListener('click', closeModal);
  if (modalCancel) modalCancel.addEventListener('click', closeModal);
  if (regModal) {
    regModal.addEventListener('click', (e) => {
      if (e.target === regModal) closeModal();
    });
  }

  // ---- Registration Form Submit (LINE integration) ----
  const regForm = document.getElementById('regForm');
  if (regForm) {
    regForm.addEventListener('submit', function (e) {
      e.preventDefault();
      const name = document.getElementById('regName').value.trim();
      const phone = document.getElementById('regPhone').value.trim();
      const company = document.getElementById('regCompany').value.trim();
      const course = courseSelect ? courseSelect.options[courseSelect.selectedIndex].text : '';
      const lineId = document.getElementById('regLineId').value.trim();

      if (!name || !phone || !course) {
        showToast('請填寫必填欄位（姓名、電話、課程）', 'error');
        return;
      }

      // Build LINE message
      const msg = `【課程報名】\n姓名：${name}\n電話：${phone}\n公司：${company || '未填寫'}\nLINE ID：${lineId || '未填寫'}\n課程：${course}`;
      const encoded = encodeURIComponent(msg);
      const lineUrl = `https://line.me/R/oaMessage/@tbpva/?${encoded}`;

      // Show success then redirect to LINE
      showToast('報名資料已準備，即將轉至 LINE 官方帳號完成報名', 'success');
      setTimeout(() => {
        window.open(lineUrl, '_blank');
        closeModal();
        regForm.reset();
      }, 1500);
    });
  }

  // ---- Quick LINE Enroll (floating & buttons) ----
  const floatLine = document.getElementById('floatLineBtn');
  if (floatLine) {
    floatLine.addEventListener('click', () => {
      window.open('https://line.me/R/ti/p/@tbpva', '_blank');
    });
  }

  // ---- Toast Notification ----
  function showToast(message, type = 'info') {
    let toast = document.getElementById('siteToast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'siteToast';
      toast.style.cssText = `
        position:fixed;bottom:32px;left:50%;transform:translateX(-50%) translateY(24px);
        background:${type === 'success' ? '#06C755' : type === 'error' ? '#c0392b' : '#1a3a5c'};
        color:#fff;padding:12px 24px;border-radius:6px;font-size:0.88rem;font-weight:600;
        box-shadow:0 8px 32px rgba(0,0,0,0.2);z-index:9999;opacity:0;
        transition:all 0.3s ease;max-width:90vw;text-align:center;font-family:'Noto Sans TC',sans-serif;
      `;
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.style.background = type === 'success' ? '#06C755' : type === 'error' ? '#c0392b' : '#1a3a5c';
    toast.style.opacity = '1';
    toast.style.transform = 'translateX(-50%) translateY(0)';
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(-50%) translateY(24px)';
    }, 3500);
  }

  // ---- Smooth Anchor Scroll ----
  document.querySelectorAll('a[href^="#"]').forEach(a => {
    a.addEventListener('click', function (e) {
      const target = document.querySelector(this.getAttribute('href'));
      if (target) {
        e.preventDefault();
        // Close mobile menu if open
        if (mainNav) { mainNav.classList.remove('open'); document.body.style.overflow = ''; }
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  });

  // ---- Counter Animation ----
  function animateCounter(el) {
    const target = parseFloat(el.dataset.target);
    const duration = 1800;
    const start = performance.now();
    const isFloat = target % 1 !== 0;

    function step(now) {
      const progress = Math.min((now - start) / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3);
      const val = target * eased;
      el.textContent = isFloat ? val.toFixed(1) : Math.floor(val).toLocaleString();
      if (progress < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  }

  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const nums = entry.target.querySelectorAll('[data-target]');
        nums.forEach(animateCounter);
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.3 });

  const statsSection = document.querySelector('.stats-grid');
  if (statsSection) observer.observe(statsSection);

  // ---- Active nav highlight ----
  const sections = document.querySelectorAll('section[id]');
  const navLinks = document.querySelectorAll('.main-nav a[href^="#"]');

  const sectionObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        navLinks.forEach(link => {
          link.parentElement.classList.toggle('active', link.getAttribute('href') === '#' + entry.target.id);
        });
      }
    });
  }, { threshold: 0.4 });

  sections.forEach(s => sectionObserver.observe(s));

  // ---- Inline quick-register form (LINE section) ----
  const quickForm = document.getElementById('quickRegForm');
  if (quickForm) {
    quickForm.addEventListener('submit', function (e) {
      e.preventDefault();
      const name = document.getElementById('qName').value.trim();
      const phone = document.getElementById('qPhone').value.trim();
      const course = document.getElementById('qCourse').value;

      if (!name || !phone || !course) {
        showToast('請填寫姓名、電話及選擇課程', 'error');
        return;
      }

      const courseText = document.getElementById('qCourse').options[document.getElementById('qCourse').selectedIndex].text;
      const msg = `【快速報名】\n姓名：${name}\n電話：${phone}\n課程：${courseText}`;
      const encoded = encodeURIComponent(msg);
      const lineUrl = `https://line.me/R/oaMessage/@tbpva/?${encoded}`;

      showToast('準備轉至 LINE 完成報名…', 'success');
      setTimeout(() => window.open(lineUrl, '_blank'), 1200);
    });
  }

})();
