const fs = require('fs');
const FormData = require('form-data');
const axios = require('axios');

// =====================================================================
// BƯỚC 1: BẠN HÃY COPY 3 GIÁ TRỊ NÀY TỪ HTTP TOOLKIT VÀ DÁN VÀO ĐÂY
// =====================================================================
const AUTHORIZATION = "Bearer eyJhbGciOi..."; 
const TOKEN_ID = "04c0a953-7fb8-5461-e063-62199f0aeda6"; 
const TOKEN_KEY = "MFwwDQYJKoZIhvcNAQEBBQADSwAwSAJBAKjy7FK9SegSCW0cuUIbEDUsbRZOCoxijNPLMfvgX+8/XA7HebHXMN4/PO5c5mwK31Yk31RKuMXYLLp6X6oZPDKcAwEAAQ==";
// =====================================================================

async function upload() {
    const form = new FormData();
    // Đảm bảo file ảnh portrait_full.jpg có sẵn trong thư mục hiện tại
    form.append('file', fs.createReadStream('portrait_full.jpg'));

    try {
        console.log("Đang tải ảnh lên máy chủ VNPT bằng phiên của điện thoại...");
        const res = await axios.post('https://api.idg.vnpt.vn/file-service/v1/addFile', form, {
            headers: {
                ...form.getHeaders(),
                'Authorization': AUTHORIZATION,
                'Token-id': TOKEN_ID,
                'Token-key': TOKEN_KEY
            }
        });
        
        console.log("\n\n✅ TẢI ẢNH THÀNH CÔNG! HÃY COPY TOÀN BỘ ĐOẠN JSON DƯỚI ĐÂY VÀ DÁN VÀO HTTP TOOLKIT:\n");
        console.log(JSON.stringify(res.data, null, 4));
        console.log("\n===============================================================================\n");
    } catch (e) {
        console.error("Lỗi:", e.response ? e.response.data : e.message);
    }
}

upload();
