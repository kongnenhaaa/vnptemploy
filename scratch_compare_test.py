import sys, json, requests
sys.path.insert(0, r'c:\Users\congn\Pictures\vnptemploy\employ')
sys.stdout.reconfigure(encoding='utf-8')

from app import (
    app, _try_restore_persistent_session, _device_auth_onebss_post,
    DEVICE_AUTH_MENU_ID, _device_auth_clean_and_crop_portrait,
    _device_auth_sdk_settings, _device_auth_access_token,
    _device_auth_client_session, _get_account_context
)

with app.app_context():
    with app.test_request_context():
        _try_restore_persistent_session()
        account_id = ''
        context = _get_account_context(account_id)
        cs = _device_auth_client_session(context)
        tp = _device_auth_onebss_post('/app-com/Config/token_ekyc', {'menu_id': int(DEVICE_AUTH_MENU_ID)}, account_id)
        cfg = _device_auth_onebss_post('/app-com/Config/app_config', {'menu_id': int(DEVICE_AUTH_MENU_ID)}, account_id)
        sdk = _device_auth_sdk_settings(cfg)
        ch_code = sdk['challenge_code']
        headers = {
            'Authorization': _device_auth_access_token(tp),
            'Token-id': sdk['token_id'],
            'Token-key': sdk['token_key'],
            'User-Agent': 'okhttp/4.11.0'
        }
        
        # Upload portrait 0814397781
        with open(r'c:\Users\congn\Pictures\vnptemploy\employ\anh\0814397781.jpg', 'rb') as f:
            b = f.read()
        clean_b = _device_auth_clean_and_crop_portrait(b)
        up_url = f"{sdk['base_url']}/file-service/v1/addFile?challenge_code={ch_code}"
        r = requests.post(up_url, headers=headers, files={'file': ('portrait.jpg', clean_b, 'image/jpeg')}, data={'title': 'portrait', 'description': 'portrait'}, timeout=15)
        h = r.json().get('object', {}).get('hash')
        print('Uploaded hash h:', h)
        
        # 1. Test /ai/v2/face/compare with img_front=h and img_face=h
        cmp_url = f"{sdk['base_url']}/ai/v2/face/compare?challenge_code={ch_code}"
        cmp_r = requests.post(cmp_url, headers={**headers, 'Content-Type': 'application/json'}, json={
            'img_front': h,
            'img_face': h,
            'client_session': cs,
            'token': '8928skjhfa89298jahga1771vbvb',
            'step_id': 1
        }, timeout=15)
        print('compare v2 status:', cmp_r.status_code, cmp_r.text)

        # 2. Test /ai/v2/face/compare-general
        gen_url = f"{sdk['base_url']}/ai/v2/face/compare-general?challenge_code={ch_code}"
        gen_r = requests.post(gen_url, headers={**headers, 'Content-Type': 'application/json'}, json={
            'img_face1': h,
            'img_face2': h,
            'client_session': cs
        }, timeout=15)
        print('compare-general v2 status:', gen_r.status_code, gen_r.text)
