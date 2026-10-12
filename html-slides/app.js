const slides = [...document.querySelectorAll('.slide')];
const counter = document.getElementById('counter');
const progress = document.getElementById('progressBar');
const overview = document.getElementById('overview');
const overviewGrid = document.getElementById('overviewGrid');
let current = Math.max(0, Math.min(slides.length - 1, Number(location.hash.slice(1)) - 1 || 0));
let initialized = false;
let transitioning = false;
let transitionTimer = null;

function pad(number) { return String(number).padStart(2, '0'); }

function show(index, updateHash = true, direction) {
  const next = (index + slides.length) % slides.length;

  if (initialized && next === current) return;

  if (transitioning) {
    clearTimeout(transitionTimer);
    slides.forEach(slide => slide.classList.remove(
      'active', 'is-leaving', 'enter-forward', 'enter-backward', 'leave-forward', 'leave-backward'
    ));
    slides[current].classList.add('active');
    transitioning = false;
  }

  if (!initialized) {
    slides.forEach((slide, i) => slide.classList.toggle('active', i === next));
    initialized = true;
  } else {
    transitioning = true;
    const previous = slides[current];
    const upcoming = slides[next];
    const movingForward = direction ? direction === 'forward' : next > current;
    const enterClass = movingForward ? 'enter-forward' : 'enter-backward';
    const leaveClass = movingForward ? 'leave-forward' : 'leave-backward';

    clearTimeout(transitionTimer);
    previous.classList.remove('active', 'enter-forward', 'enter-backward');
    previous.classList.add('is-leaving', leaveClass);
    upcoming.classList.remove('is-leaving', 'leave-forward', 'leave-backward');
    upcoming.classList.add('active', enterClass);

    transitionTimer = setTimeout(() => {
      previous.classList.remove('is-leaving', leaveClass);
      upcoming.classList.remove(enterClass);
      transitioning = false;
    }, 680);
  }

  current = next;
  counter.textContent = `${pad(current + 1)} / ${slides.length}`;
  progress.style.width = `${((current + 1) / slides.length) * 100}%`;
  document.title = `${pad(current + 1)} · ${slides[current].dataset.title}`;
  document.querySelectorAll('.thumb').forEach((thumb, i) => thumb.classList.toggle('current', i === current));
  if (updateHash) history.replaceState(null, '', `#${current + 1}`);
  restartSlideEffects(slides[current]);
}

function restartSlideEffects(slide) {
  slide.querySelectorAll('.count-up').forEach(counter => {
    const target = Number(counter.dataset.count || counter.textContent);
    const started = performance.now();
    const decimals = Number(counter.dataset.decimals || 0);
    const duration = 900;
    const tick = now => {
      const progress = Math.min(1, (now - started) / duration);
      counter.textContent = (target * (1 - Math.pow(1 - progress, 3))).toFixed(decimals);
      if (progress < 1) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  });
}

function toggleOverview(force) {
  const open = typeof force === 'boolean' ? force : !overview.classList.contains('open');
  overview.classList.toggle('open', open);
  overview.setAttribute('aria-hidden', String(!open));
}

slides.forEach((slide, index) => {
  const button = document.createElement('button');
  button.className = 'thumb';
  button.innerHTML = `<b>${pad(index + 1)}</b><span>${slide.dataset.title}</span><small>${pad(index + 1)} / ${slides.length}</small>`;
  button.addEventListener('click', () => { show(index, true, index >= current ? 'forward' : 'backward'); toggleOverview(false); });
  overviewGrid.appendChild(button);
});

document.getElementById('prevBtn').addEventListener('click', () => show(current - 1, true, 'backward'));
document.getElementById('nextBtn').addEventListener('click', () => show(current + 1, true, 'forward'));
document.getElementById('overviewBtn').addEventListener('click', () => toggleOverview());
document.getElementById('closeOverview').addEventListener('click', () => toggleOverview(false));
document.getElementById('fullscreenBtn').addEventListener('click', () => {
  if (!document.fullscreenElement) document.documentElement.requestFullscreen?.();
  else document.exitFullscreen?.();
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && overview.classList.contains('open')) return toggleOverview(false);
  if (['ArrowRight', 'ArrowDown', 'PageDown', ' '].includes(event.key)) { event.preventDefault(); show(current + 1, true, 'forward'); }
  if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(event.key)) { event.preventDefault(); show(current - 1, true, 'backward'); }
  if (event.key.toLowerCase() === 'o') toggleOverview();
  if (event.key.toLowerCase() === 'f') document.getElementById('fullscreenBtn').click();
  if (event.key === 'Home') show(0, true, 'backward');
  if (event.key === 'End') show(slides.length - 1, true, 'forward');
});

let wheelDistance = 0;
let wheelLocked = false;
let wheelResetTimer = null;
document.addEventListener('wheel', event => {
  if (overview.classList.contains('open')) return;
  const delta = Math.abs(event.deltaY) >= Math.abs(event.deltaX) ? event.deltaY : event.deltaX;
  if (!delta) return;
  event.preventDefault();
  if (wheelLocked) return;

  wheelDistance += delta;
  clearTimeout(wheelResetTimer);
  wheelResetTimer = setTimeout(() => { wheelDistance = 0; }, 140);
  if (Math.abs(wheelDistance) < 48) return;

  const direction = wheelDistance > 0 ? 'forward' : 'backward';
  wheelDistance = 0;
  wheelLocked = true;
  show(current + (direction === 'forward' ? 1 : -1), true, direction);
  setTimeout(() => { wheelLocked = false; }, 760);
}, { passive: false });

let touchStart = null;
document.addEventListener('touchstart', event => {
  const touch = event.changedTouches[0];
  touchStart = { x: touch.clientX, y: touch.clientY };
}, { passive: true });
document.addEventListener('touchend', event => {
  if (touchStart === null) return;
  if (overview.classList.contains('open')) { touchStart = null; return; }
  const touch = event.changedTouches[0];
  const deltaX = touch.clientX - touchStart.x;
  const deltaY = touch.clientY - touchStart.y;
  if (Math.max(Math.abs(deltaX), Math.abs(deltaY)) > 60) {
    const delta = Math.abs(deltaX) >= Math.abs(deltaY) ? deltaX : deltaY;
    const direction = delta < 0 ? 'forward' : 'backward';
    show(current + (direction === 'forward' ? 1 : -1), true, direction);
  }
  touchStart = null;
}, { passive: true });

window.addEventListener('hashchange', () => {
  const next = Number(location.hash.slice(1)) - 1;
  show(next, false, next >= current ? 'forward' : 'backward');
});
show(current, false);
