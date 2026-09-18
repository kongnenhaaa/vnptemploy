(() => {
  'use strict';

  const MENU_ID = 11077;
  const DEFAULT_DELAY_SECONDS = 20;
  const ALLOWED_PACKAGE_CODES = new Set([
    'KM_DT30',
    'MI_YOLO100M',
    'SODA155',
    'MI_YOLO125V',
  ]);
  const state = {
    initialized: false,
    initializing: null,
    permissionOk: false,
    subscriber: null,
    lookupPhone: '',
    packages: [],
    selectedPackageId: '',
    rows: [],
    running: false,
    stopRequested: false,
    runOutputPath: '',
    totalOutputPath: '',
  };

  const el = id => document.getElementById(id);
  const wait = ms => new Promise(resolve => window.setTimeout(resolve, ms));

  function apiData(response) {
    const body = apiBodyObject(response);
    return body && typeof body === 'object' && Object.prototype.hasOwnProperty.call(body, 'data')
      ? body.data
      : body;
  }

  function normalizePhone(value) {
    let digits = String(value || '').replace(/\D/g, '');
    if (digits.startsWith('0084')) digits = digits.slice(2);
    if (digits.length === 10 && digits.startsWith('0')) digits = `84${digits.slice(1)}`;
    else if (digits.length === 9) digits = `84${digits}`;
    if (!/^84\d{9}$/.test(digits)) {
      throw new Error(`SĐT không đúng định dạng: ${value || '(trống)'}`);
    }
    return digits;
  }

  function nationalPhone(value) {
    const digits = String(value || '').replace(/\D/g, '');
    return /^84\d{9}$/.test(digits) ? `0${digits.slice(2)}` : digits;
  }

  function durationLabel(value) {
    const days = Number(value || 0);
    const labels = {30: '1T · 30 ngày', 90: '3T · 90 ngày', 180: '6T · 180 ngày', 360: '12T · 360 ngày'};
    return labels[days] || (days ? `${days} ngày` : '—');
  }

  function basePackageCode(value) {
    return String(value || '')
      .trim()
      .toUpperCase()
      .replace(/_(3|6|12)M$/, '');
  }

  function priceValue(value) {
    const number = Number(String(value ?? '').replace(/[^0-9.-]/g, ''));
    return Number.isFinite(number) ? number : 0;
  }

  function priceLabel(value) {
    const number = priceValue(value);
    return number ? `${new Intl.NumberFormat('vi-VN').format(number)} đ` : '0 đ';
  }

  function message(text, type = '') {
    const target = el('package-registration-message');
    if (!target) return;
    target.textContent = text || '';
    target.className = `package-registration-message ${type}`.trim();
  }

  function batchMessage(text, type = '') {
    const target = el('package-registration-batch-message');
    if (!target) return;
    target.textContent = text || '';
    target.className = `package-registration-message ${type}`.trim();
  }

  function setPermission(kind, text) {
    const badge = el('package-registration-permission');
    if (!badge) return;
    badge.className = `package-registration-badge ${kind}`;
    badge.textContent = text;
  }

  function publicPackage(item, index) {
    const code = String(item?.serviceCode ?? item?.SERVICE_CODE ?? item?.serviceId ?? item?.SERVICE_ID ?? '').trim();
    const systemCode = String(item?.systemCode ?? item?.SYSTEM_CODE ?? 'SPS').trim() || 'SPS';
    const duration = String(item?.duration ?? item?.DURATION ?? '').trim();
    const price = priceValue(item?.price ?? item?.PRICE ?? 0);
    const name = String(item?.serviceName ?? item?.SERVICE_NAME ?? item?.packageName ?? item?.PACKAGE_NAME ?? item?.name ?? code).trim() || code;
    const description = String(item?.desc ?? item?.description ?? item?.DESCRIPTION ?? '').trim();
    return {
      id: `${systemCode}|${code}|${duration}|${price}|${index}`,
      code,
      systemCode,
      duration,
      price,
      name,
      description,
    };
  }

  function packageMatches(left, right) {
    return String(left.code).toUpperCase() === String(right.code).toUpperCase()
      && String(left.systemCode).toUpperCase() === String(right.systemCode).toUpperCase()
      && String(left.duration) === String(right.duration)
      && Number(left.price) === Number(right.price);
  }

  function selectedPackage() {
    return state.packages.find(item => item.id === state.selectedPackageId) || null;
  }

  function currentListPhones() {
    const rawLines = String(el('package-registration-list')?.value || '').split(/[\r\n,;\t]+/);
    const phones = [];
    const invalid = [];
    const seen = new Set();
    rawLines.forEach((raw, index) => {
      if (!String(raw).trim()) return;
      try {
        const phone = normalizePhone(raw);
        if (seen.has(phone)) return;
        seen.add(phone);
        phones.push(phone);
      } catch (_) {
        invalid.push(index + 1);
      }
    });
    if (invalid.length) {
      throw new Error(`SĐT sai định dạng tại dòng ${invalid.slice(0, 10).join(', ')}`);
    }
    return phones;
  }

  function updateButtons() {
    const selected = selectedPackage();
    const hasSingle = Boolean(state.lookupPhone && selected);
    const pending = state.rows.filter(row => row.status === 'pending').length;
    const delay = Number(el('package-registration-delay')?.value || 0);
    const validDelay = Number.isInteger(delay) && delay >= 5 && delay <= 3600;
    const allowed = state.permissionOk && !state.running;
    el('package-registration-save-btn').disabled = !allowed || !hasSingle;
    el('package-registration-run-btn').disabled = !allowed || !selected || !pending || !validDelay;
    el('package-registration-stop-btn').disabled = !state.running;
    el('package-registration-lookup-btn').disabled = state.running || !state.permissionOk;
    el('package-registration-open-run').disabled = state.running || !state.runOutputPath;
    el('package-registration-open-total').disabled = state.running || !state.totalOutputPath;
    el('package-registration-delay').disabled = state.running;
    el('package-registration-list').disabled = state.running;
    el('package-registration-file').disabled = state.running;
  }

  function renderSubscriber(data, phone) {
    const target = el('package-registration-subscriber');
    if (!target) return;
    const activePackages = Array.isArray(data?.GOI_DATA) ? data.GOI_DATA : [];
    const values = [
      ['Số thuê bao', nationalPhone(data?.SO_TB || phone)],
      ['Khách hàng', data?.TEN_KH || '—'],
      ['Loại thuê bao', data?.LOAI_TB || '—'],
      ['Ngày kích hoạt', data?.NGAY_KH || '—'],
      ['Trạng thái', String(data?.TRANG_THAI ?? '—')],
      ['Gói Data hiện có', activePackages.map(item => item?.PACKAGE_NAME || item?.SERVICES).filter(Boolean).join(', ') || 'Không có'],
    ];
    target.innerHTML = values.map(([label, value]) => (
      `<div><span>${escHtml(label)}</span><strong>${escHtml(String(value))}</strong></div>`
    )).join('');
    target.hidden = false;
  }

  function renderPackages() {
    const tbody = el('package-registration-package-body');
    if (!tbody) return;
    const filter = String(el('package-registration-filter')?.value || '').trim().toLowerCase();
    const duration = String(el('package-registration-duration')?.value || '30');
    const visible = state.packages.filter(item => {
      const haystack = `${item.code} ${item.name} ${item.description}`.toLowerCase();
      const durationMatches = String(item.duration || '30') === duration;
      return durationMatches && (!filter || haystack.includes(filter));
    });
    el('package-registration-package-count').textContent = `${visible.length} gói`;
    if (!visible.length) {
      tbody.innerHTML = '<tr class="package-registration-empty-row"><td colspan="6" class="package-registration-empty">OneBSS không trả một trong 4 gói đã chọn ở thời hạn này cho thuê bao.</td></tr>';
      updateButtons();
      return;
    }
    tbody.innerHTML = visible.map(item => `
      <tr class="${item.id === state.selectedPackageId ? 'selected' : ''}" data-package-id="${escHtml(item.id)}">
        <td><input type="radio" name="package-registration-package" value="${escHtml(item.id)}" ${item.id === state.selectedPackageId ? 'checked' : ''}></td>
        <td class="code">${escHtml(item.code)}</td>
        <td>${escHtml(item.name)}</td>
        <td>${escHtml(durationLabel(item.duration))}</td>
        <td class="price">${escHtml(priceLabel(item.price))}</td>
        <td class="description">${escHtml(item.description || '—')}</td>
      </tr>`).join('');
    tbody.querySelectorAll('tr[data-package-id]').forEach(row => {
      row.addEventListener('click', () => {
        state.selectedPackageId = row.dataset.packageId || '';
        const item = selectedPackage();
        el('package-registration-selected').textContent = item
          ? `Đã chọn: ${item.code} · ${durationLabel(item.duration)} · ${priceLabel(item.price)}`
          : 'Chưa chọn gói.';
        renderPackages();
        updateButtons();
      });
    });
    updateButtons();
  }

  function renderBatch() {
    const tbody = el('package-registration-batch-body');
    if (!tbody) return;
    if (!state.rows.length) {
      tbody.innerHTML = '<tr><td colspan="8" class="package-registration-empty">Chưa có danh sách.</td></tr>';
    } else {
      tbody.innerHTML = state.rows.map(row => `
        <tr class="package-registration-row-${escHtml(row.status)}">
          <td>${row.index}</td><td class="code">${escHtml(nationalPhone(row.phone))}</td>
          <td>${escHtml({pending:'Chờ chạy',running:'Đang chạy',success:'Thành công',error:'Lỗi'}[row.status] || row.status)}</td>
          <td class="code">${escHtml(row.packageCode || selectedPackage()?.code || '—')}</td>
          <td>${escHtml(row.duration ? durationLabel(row.duration) : '—')}</td>
          <td class="price">${escHtml(row.price !== '' && row.price != null ? priceLabel(row.price) : '—')}</td>
          <td class="description">${escHtml(row.result || '—')}</td><td>${escHtml(row.completedAt || '—')}</td>
        </tr>`).join('');
    }
    const total = state.rows.length;
    const complete = state.rows.filter(row => ['success', 'error'].includes(row.status)).length;
    const success = state.rows.filter(row => row.status === 'success').length;
    const failed = state.rows.filter(row => row.status === 'error').length;
    const pending = state.rows.filter(row => row.status === 'pending').length;
    el('package-registration-progress-fill').style.width = `${total ? Math.round(complete * 100 / total) : 0}%`;
    el('package-registration-progress-text').textContent = `${complete}/${total}`;
    el('package-registration-counts').textContent = `Thành công ${success} · Lỗi ${failed} · Còn lại ${pending}`;
    updateButtons();
  }

  function prepareRows() {
    if (state.running) return;
    try {
      const phones = currentListPhones();
      state.rows = phones.map((phone, index) => ({
        id: `${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
        index: index + 1,
        phone,
        status: 'pending',
        packageCode: '',
        duration: '',
        price: '',
        result: '',
        completedAt: '',
      }));
      state.runOutputPath = '';
      batchMessage(phones.length ? `Đã nạp ${phones.length} SĐT.` : 'Chưa có danh sách.');
    } catch (error) {
      state.rows = [];
      batchMessage(error.message, 'error');
    }
    renderBatch();
  }

  async function rawLookup(phone) {
    const subscriberResponse = await callApi('/ccbs/oneBss/app_tb_tc_thongtin', 'POST', {
      so_tb: phone,
      service: 'SIM4G',
      menu_id: MENU_ID,
    });
    if (!apiSucceeded(subscriberResponse)) {
      throw new Error(apiFailureMessage(subscriberResponse, 'Không tra cứu được thuê bao'));
    }
    const subscriber = apiData(subscriberResponse) || {};

    const packageResponse = await callApi('/ccbs/goicuoc/ds-goi-duoc-dang-ky', 'POST', {
      so_tb: phone,
      chuong_trinh_ban_goi: 'SPS',
      menu_id: MENU_ID,
    });
    if (!apiSucceeded(packageResponse)) {
      throw new Error(apiFailureMessage(packageResponse, 'Không tải được danh sách gói'));
    }
    const rawPackages = apiData(packageResponse);
    const list = Array.isArray(rawPackages)
      ? rawPackages
      : (Array.isArray(rawPackages?.items) ? rawPackages.items : []);
    const packages = list
      .map(publicPackage)
      .filter(item => item.code && ALLOWED_PACKAGE_CODES.has(basePackageCode(item.code)));
    return {subscriber, packages};
  }

  async function lookup(phoneValue) {
    await initialize();
    if (!state.permissionOk) throw new Error('Tài khoản không có quyền Đăng ký gói cước.');
    const phone = normalizePhone(phoneValue || el('package-registration-phone')?.value);
    const button = el('package-registration-lookup-btn');
    button.disabled = true;
    button.textContent = 'Đang tải…';
    message(`Đang tra cứu ${nationalPhone(phone)} và tải danh sách gói…`);
    try {
      const result = await rawLookup(phone);
      state.lookupPhone = phone;
      state.subscriber = result.subscriber;
      state.packages = result.packages;
      state.selectedPackageId = '';
      el('package-registration-phone').value = nationalPhone(phone);
      el('package-registration-selected').textContent = 'Chưa chọn gói.';
      el('package-registration-filter').value = '';
      el('package-registration-duration').value = '30';
      renderSubscriber(result.subscriber, phone);
      renderPackages();
      message(`Đã tải ${result.packages.length} gói OneBSS cho ${nationalPhone(phone)}.`, 'success');
      return result;
    } finally {
      button.textContent = '＋ Tra cứu gói';
      updateButtons();
    }
  }

  function existingDataPackages(subscriber) {
    const groups = [subscriber?.GOI_DATA, subscriber?.GOI_CUOC, subscriber?.GOI_CUOC_TS];
    return groups.flatMap(value => Array.isArray(value) ? value : []);
  }

  async function executeOne(phone, selection) {
    let lookupResult;
    try {
      lookupResult = await rawLookup(phone);
    } catch (error) {
      return {ok: false, uncertain: false, message: error.message};
    }
    const candidate = lookupResult.packages.find(item => packageMatches(item, selection));
    if (!candidate) {
      return {
        ok: false,
        uncertain: false,
        message: `Thuê bao không còn được OneBSS cấp đúng gói ${selection.code}, ${durationLabel(selection.duration)}, ${priceLabel(selection.price)}.`,
      };
    }
    const alreadyActive = existingDataPackages(lookupResult.subscriber).some(item => {
      const code = String(item?.PACKAGE_NAME ?? item?.SERVICES ?? item?.serviceCode ?? '').trim().toUpperCase();
      return code === String(selection.code).trim().toUpperCase();
    });
    if (alreadyActive) {
      return {
        ok: false,
        uncertain: false,
        message: `Thuê bao đang có gói ${selection.code}; tool không gửi yêu cầu đăng ký trùng.`,
      };
    }

    const response = await callApi('/ccbs/goicuoc/dangky', 'POST', {
      msisdn: phone,
      serviceCode: candidate.code,
      systemCode: candidate.systemCode,
      menu_id: MENU_ID,
    });
    if (apiSucceeded(response)) {
      const body = apiBodyObject(response) || {};
      return {
        ok: true,
        uncertain: false,
        message: String(body.message || body.data || `Đăng ký ${candidate.code} thành công.`),
      };
    }
    return {
      ok: false,
      uncertain: isBusinessMutationUncertain(response),
      message: apiFailureMessage(response, 'OneBSS từ chối đăng ký gói'),
    };
  }

  async function startOutput() {
    const response = await fetch('/api/package-registration/start-output', {
      method: 'POST', headers: {'Content-Type': 'application/json'}, body: '{}',
    });
    const data = await response.json();
    if (!response.ok || !data.ok) throw new Error(data.error || `HTTP ${response.status}`);
    state.runOutputPath = data.run_output_path || '';
    state.totalOutputPath = data.cumulative_output_path || state.totalOutputPath;
    renderOutputPaths();
  }

  function outputRow(row) {
    return [
      row.index,
      nationalPhone(row.phone),
      row.packageCode,
      row.systemCode,
      row.duration,
      row.price,
      row.status === 'success' ? 'Thành công' : 'Lỗi',
      row.result,
      row.completedAt,
      '',
    ];
  }

  async function appendOutput(row) {
    const response = await fetch('/api/package-registration/append-output', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        run_output_path: state.runOutputPath,
        rows: [outputRow(row)],
        entry_ids: [row.id],
      }),
    });
    const data = await response.json();
    if (!response.ok || !data.ok) {
      const error = new Error(data.error || `HTTP ${response.status}`);
      error.runSaved = Boolean(data.run_saved);
      throw error;
    }
    state.totalOutputPath = data.cumulative_output_path || state.totalOutputPath;
    renderOutputPaths();
  }

  function renderOutputPaths() {
    const target = el('package-registration-output-paths');
    if (!target) return;
    const parts = [];
    if (state.runOutputPath) parts.push(`Output phiên: ${state.runOutputPath}`);
    if (state.totalOutputPath) parts.push(`Output tổng: ${state.totalOutputPath}`);
    target.textContent = parts.join(' · ');
    updateButtons();
  }

  async function runOneRow(row, selection) {
    row.status = 'running';
    row.packageCode = selection.code;
    row.systemCode = selection.systemCode;
    row.duration = selection.duration;
    row.price = selection.price;
    renderBatch();
    const result = await executeOne(row.phone, selection);
    row.status = result.ok ? 'success' : 'error';
    row.result = result.message;
    row.uncertain = Boolean(result.uncertain);
    row.completedAt = new Date().toLocaleString('vi-VN');
    renderBatch();
    await appendOutput(row);
    return result;
  }

  async function saveSingle() {
    const selection = selectedPackage();
    if (!selection || !state.lookupPhone || state.running) return;
    const confirmText = [
      `Xác nhận đăng ký gói ${selection.code}`,
      `Thuê bao: ${nationalPhone(state.lookupPhone)}`,
      `Thời hạn: ${durationLabel(selection.duration)}`,
      `Giá: ${priceLabel(selection.price)}`,
      '',
      'Thao tác có thể phát sinh cước và không thể tự hoàn tác.',
    ].join('\n');
    if (!window.confirm(confirmText)) return;

    state.running = true;
    state.stopRequested = false;
    const row = {
      id: `${Date.now()}-single-${Math.random().toString(16).slice(2)}`,
      index: 1,
      phone: state.lookupPhone,
      status: 'pending',
      packageCode: '', systemCode: '', duration: '', price: '', result: '', completedAt: '',
    };
    const showManualInBatch = state.rows.length === 0;
    if (showManualInBatch) {
      state.rows = [row];
      renderBatch();
    }
    batchMessage(`Đang đăng ký ${selection.code} cho ${nationalPhone(row.phone)}…`);
    try {
      await startOutput();
      const result = await runOneRow(row, selection);
      batchMessage(result.message, result.ok ? 'success' : 'error');
    } catch (error) {
      batchMessage(`Đã dừng: ${error.message}`, 'error');
    } finally {
      state.running = false;
      renderBatch();
    }
  }

  async function runBatch() {
    if (state.running) return;
    const selection = selectedPackage();
    const pending = state.rows.filter(row => row.status === 'pending');
    const delaySeconds = Number(el('package-registration-delay')?.value || 0);
    if (!selection || !pending.length) return;
    if (!Number.isInteger(delaySeconds) || delaySeconds < 5 || delaySeconds > 3600) {
      batchMessage('Delay phải là số nguyên từ 5 đến 3600 giây.', 'error');
      return;
    }
    const confirmText = [
      `Xác nhận đăng ký gói ${selection.code} cho ${pending.length} SĐT còn lại?`,
      `Thời hạn: ${durationLabel(selection.duration)}`,
      `Giá: ${priceLabel(selection.price)}`,
      `Chạy 1 luồng; lưu xong một SĐT rồi chờ ${delaySeconds} giây.`,
      '',
      'Thao tác có thể phát sinh cước. Tool không tự gửi lại kết quả không xác định.',
    ].join('\n');
    if (!window.confirm(confirmText)) return;

    state.running = true;
    state.stopRequested = false;
    renderBatch();
    try {
      if (!state.runOutputPath) await startOutput();
      for (const row of state.rows) {
        if (row.status !== 'pending') continue;
        if (state.stopRequested) break;
        batchMessage(`Đang xử lý ${nationalPhone(row.phone)} với gói ${selection.code}…`);
        let result;
        try {
          result = await runOneRow(row, selection);
        } catch (error) {
          batchMessage(`Đã dừng sau ${nationalPhone(row.phone)}: ${error.message}`, 'error');
          state.stopRequested = true;
          break;
        }
        if (result.uncertain) {
          batchMessage(`Kết quả ${nationalPhone(row.phone)} chưa xác định; đã lưu và dừng, không tự gửi lại.`, 'error');
          state.stopRequested = true;
          break;
        }
        const hasMore = state.rows.some(item => item.status === 'pending');
        if (!hasMore) break;
        for (let remaining = delaySeconds; remaining > 0 && !state.stopRequested; remaining--) {
          batchMessage(`Đã lưu ${nationalPhone(row.phone)}. Chờ ${remaining} giây trước SĐT tiếp theo.`);
          await wait(1000);
        }
      }
      if (state.stopRequested) {
        const remaining = state.rows.filter(row => row.status === 'pending').length;
        if (!el('package-registration-batch-message').classList.contains('error')) {
          batchMessage(`Đã dừng. Còn ${remaining} SĐT chưa chạy.`);
        }
      } else {
        batchMessage('Đã chạy hết danh sách; từng kết quả đã được ghi vào output phiên và output tổng.', 'success');
      }
    } catch (error) {
      batchMessage(`Không bắt đầu được danh sách: ${error.message}`, 'error');
    } finally {
      state.running = false;
      renderBatch();
    }
  }

  function stop() {
    if (!state.running) return;
    state.stopRequested = true;
    batchMessage('Đang dừng sau khi lưu xong SĐT hiện tại…');
    updateButtons();
  }

  async function openOutput(kind) {
    const response = await fetch('/api/package-registration/open', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({type: kind, path: kind === 'run' ? state.runOutputPath : ''}),
    });
    const data = await response.json();
    if (!response.ok || !data.ok) {
      batchMessage(data.error || `HTTP ${response.status}`, 'error');
      return;
    }
    batchMessage(`Đã mở ${data.path}.`, 'success');
  }

  async function loadFile(input) {
    const file = input?.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      el('package-registration-list').value = text;
      el('package-registration-file-name').textContent = file.name;
      prepareRows();
      const first = state.rows[0]?.phone;
      if (first) {
        el('package-registration-phone').value = nationalPhone(first);
        await lookup(first);
      }
    } catch (error) {
      batchMessage(error.message, 'error');
    } finally {
      input.value = '';
    }
  }

  async function initialize() {
    if (state.initialized) return;
    if (state.initializing) return state.initializing;
    state.initializing = (async () => {
      setPermission('loading', 'Đang kiểm tra quyền 11077…');
      try {
        const filesResponse = await fetch('/api/package-registration/files');
        const files = await filesResponse.json();
        if (filesResponse.ok && files.ok) {
          state.totalOutputPath = files.output_path || '';
          renderOutputPaths();
        }

        // Ghi nhận đúng chức năng đang dùng; lỗi log không được che kết quả
        // kiểm tra quyền nghiệp vụ.
        await callApi('/quantri/user/log_sudung_chucnang', 'POST', {
          p_menu_id: MENU_ID,
          p_device_string: 'VNPT Employ Desktop',
          menu_id: MENU_ID,
        }).catch(() => null);

        const permission = await callApi(
          '/app-banhang/luong_didong_moi/mhddm_kiemtra_maquyen',
          'POST',
          {ma_quyen: 'BANGOICUOCDIDONG', menu_id: MENU_ID},
        );
        const permissionData = apiData(permission) || {};
        state.permissionOk = apiSucceeded(permission) && String(permissionData.status || '') === '1';
        if (!state.permissionOk) {
          throw new Error(permissionData.message || apiFailureMessage(permission, 'Không có quyền BANGOICUOCDIDONG'));
        }
        setPermission('ok', 'Có quyền truy cập');
        state.initialized = true;
      } catch (error) {
        state.permissionOk = false;
        setPermission('error', error.message || 'Không có quyền');
        message(error.message || 'Không kiểm tra được quyền.', 'error');
        throw error;
      } finally {
        state.initializing = null;
        updateButtons();
      }
    })();
    return state.initializing;
  }

  el('package-registration-list')?.addEventListener('input', prepareRows);
  el('package-registration-delay')?.addEventListener('input', updateButtons);
  el('package-registration-duration')?.addEventListener('change', () => {
    state.selectedPackageId = '';
    el('package-registration-selected').textContent = 'Chưa chọn gói.';
    renderPackages();
  });

  window.packageRegistrationInitialize = initialize;
  window.packageRegistrationLookup = () => lookup().catch(error => message(error.message, 'error'));
  window.packageRegistrationRenderPackages = renderPackages;
  window.packageRegistrationSaveSingle = saveSingle;
  window.packageRegistrationRunBatch = runBatch;
  window.packageRegistrationStop = stop;
  window.packageRegistrationOpenOutput = openOutput;
  window.packageRegistrationLoadFile = loadFile;
})();
