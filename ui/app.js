'use strict';
let bridge = null;
let current = null;
const groupSignatures = {};
const byId = id => document.getElementById(id);

function setTheme(light) {
  document.documentElement.dataset.theme = light ? 'light' : 'dark';
  try { localStorage.setItem('hyper-theme', light ? 'light' : 'dark'); } catch (_) {}
}

function render(payload) {
  const state = typeof payload === 'string' ? JSON.parse(payload) : payload;
  current = state;
  byId('connection').hidden = true;
  byId('workspace').hidden = false;
  document.body.classList.toggle('compact', state.compact);
  byId('setup').hidden = state.compact;
  byId('activity').hidden = !state.compact;
  byId('collapse-icon').hidden = !state.running;
  byId('adas-title').textContent = state.adasTitle;
  byId('mode-caption').textContent = state.controls.mode_switch.checked ? 'Repair SI' : 'ADAS SI';
  setTheme(state.controls.theme_toggle.checked);
  byId('theme-label').textContent = state.controls.theme_toggle.text;
  for (const [name, items] of Object.entries(state.groups)) {
    const signature = JSON.stringify(items);
    // Keep DOM nodes stable during polling so keyboard focus and scroll survive.
    if (groupSignatures[name] !== signature) {
      const container = byId(name);
      if (container.children.length !== items.length || Array.from(container.children).some((row, i) => row.dataset.text !== items[i].text)) {
        const fragment = document.createDocumentFragment();
        items.forEach((item, index) => {
          const row = document.createElement('label');
          row.className = 'check-row';
          row.dataset.text = item.text;
          const input = document.createElement('input');
          input.type = 'checkbox';
          input.addEventListener('change', () => bridge.select(name, index, input.checked));
          const text = document.createElement('span');
          text.textContent = item.text;
          row.append(input, text);
          fragment.append(row);
        });
        container.textContent = '';
        container.appendChild(fragment);
      }
      items.forEach((item, index) => {
        const row = container.children[index];
        row.classList.toggle('selected', item.checked);
        row.classList.toggle('disabled', !item.enabled);
        row.firstChild.checked = item.checked;
        row.firstChild.disabled = !item.enabled;
        row.title = item.text;
      });
      groupSignatures[name] = signature;
    }
    byId(`${name}-count`).textContent = `${items.filter(item => item.checked).length} / ${items.length}`;
    const card = byId(`${name}-card`);
    if (card) card.classList.toggle('inactive', items.every(item => !item.enabled));
  }
  const files = byId('files');
  if (files.dataset.signature !== JSON.stringify(state.files)) {
    files.textContent = '';
    state.files.map(text => {
      const row = document.createElement('div'); row.className = 'file-row'; row.textContent = text; return row;
    }).forEach(row => files.appendChild(row));
    files.dataset.signature = JSON.stringify(state.files);
  }
  document.querySelectorAll('[data-toggle]').forEach(input => {
    const control = state.controls[input.dataset.toggle];
    input.checked = control.checked; input.disabled = !control.enabled;
  });
  document.querySelectorAll('[data-control]').forEach(group => {
    const control = state.controls[group.dataset.control];
    group.querySelectorAll('button').forEach(button => {
      const active = (button.dataset.value === 'true') === control.checked;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', String(active));
      button.disabled = !control.enabled;
    });
  });
  document.querySelectorAll('[data-action]').forEach(button => {
    const original = state.buttons[button.dataset.action];
    button.disabled = !original.enabled;
    if (button.dataset.action === 'start_button') {
      byId('start-text').textContent = original.text;
      button.classList.toggle('stopped', original.text === 'Stop Automation');
    } else if (button.dataset.action === 'pause_button') byId('pause-text').textContent = original.text;
  });
  byId('transparency').value = state.opacity;
  byId('opacity').textContent = `${state.opacity}%`;
  if (!byId('progress').children.length) {
    state.progress.forEach(() => {
      const item = document.createElement('div');
      const info = document.createElement('div'); info.className = 'progress-info';
      info.append(document.createElement('span'), document.createElement('span'));
      info.lastChild.className = 'percent';
      const track = document.createElement('div'); track.className = 'progress-track'; track.setAttribute('role', 'progressbar');
      const fill = document.createElement('div'); fill.className = 'progress-fill'; track.append(fill);
      item.append(info, track); byId('progress').append(item);
    });
  }
  state.progress.forEach((progress, i) => {
    const item = byId('progress').children[i];
    const percent = progress.maximum > 0 ? Math.min(100, Math.floor(progress.value / progress.maximum * 100)) : 0;
    item.firstChild.firstChild.textContent = progress.text;
    item.firstChild.lastChild.textContent = progress.maximum === 0 ? '…' : `${percent}%`;
    const track = item.lastChild;
    track.classList.toggle('stopped', progress.stopped);
    track.classList.toggle('busy', progress.maximum === 0);
    track.setAttribute('aria-label', progress.text);
    track.setAttribute('aria-valuemin', '0');
    track.setAttribute('aria-valuemax', String(progress.maximum));
    track.setAttribute('aria-valuenow', String(progress.value));
    track.firstChild.style.width = `${percent}%`;
  });
  const log = byId('log');
  if (log.textContent !== state.log) {
    const atEnd = log.scrollHeight - log.scrollTop - log.clientHeight < 40;
    log.textContent = state.log;
    if (atEnd) log.scrollTop = log.scrollHeight;
  }
}

document.querySelectorAll('[data-action]').forEach(button => button.addEventListener('click', () => bridge && bridge.click(button.dataset.action)));
document.querySelectorAll('[data-toggle]').forEach(input => input.addEventListener('change', () => bridge && bridge.toggle(input.dataset.toggle, input.checked)));
document.querySelectorAll('[data-control] button').forEach(button => button.addEventListener('click', () => bridge && bridge.toggle(button.parentElement.dataset.control, button.dataset.value === 'true')));
byId('transparency').addEventListener('input', event => bridge && bridge.transparency(Number(event.target.value)));
byId('collapse-icon').addEventListener('click', () => bridge && bridge.collapse());
const login = location.search === '?login';
if (login) { byId('login').hidden = false; byId('connection').hidden = true; }
try { setTheme(localStorage.getItem('hyper-theme') === 'light'); } catch (_) {}

if (typeof qt !== 'undefined' && typeof QWebChannel !== 'undefined') {
  new QWebChannel(qt.webChannelTransport, channel => {
    bridge = channel.objects.hyper;
    if (login) {
      bridge.theme(document.documentElement.dataset.theme === 'light');
      byId('signin').addEventListener('submit', event => {
        event.preventDefault();
        bridge.signIn(byId('username').value, byId('password').value);
        byId('password').value = '';
      });
      byId('cancel').addEventListener('click', () => bridge.cancel());
      byId('username').focus();
    } else {
      bridge.changed.connect(render);
      let light = false;
      try { light = localStorage.getItem('hyper-theme') === 'light'; } catch (_) {}
      bridge.toggle('theme_toggle', light, () => bridge.getState(render));
    }
  });
} else {
  byId('connection').textContent = 'Open HyperWeb.py to connect this interface to Hyper.';
  byId('connection').hidden = false;
  document.querySelectorAll('button,input').forEach(control => { control.disabled = true; });
}
