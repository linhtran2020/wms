# GHN WMS — Inbound • Putaway • Inventory

Hệ thống quản lý Nhập – Cất kho – Tồn kho cho hàng bánh tươi (GHN Fulfillment).

## Luồng nghiệp vụ
1. **Inbound**: Upload Excel (sheet `IB`: STT, SKU, Tên Bánh, Order uom, SL BÁNH CÓ HÓA ĐƠN) → đơn ở trạng thái **Đã tạo đơn hàng nhập**.
2. **Nhận hàng**: hàng vào vị trí mặc định **InboundQC** → trạng thái **Đã nhận hàng**, được phép Putaway.
3. **Putaway**: 1 SKU có thể chia vào nhiều vị trí; chỉ cho phép **PL.Tam.01 → PL.Tam.100**.
4. **Inventory**: tồn kho theo SKU và từng vị trí, kèm thẻ kho.

## Chạy local
```bash
npm install
npm start   # http://localhost:3000, dữ liệu lưu ở data/db.json
```

## Deploy Vercel
- **Production URL**: [https://wms-ghnlogs.vercel.app](https://wms-ghnlogs.vercel.app)
- UI tĩnh trong `public/`, API Express qua `api/index.js` (xem `vercel.json`).
- Vercel không ghi file được → cần **Upstash Redis**:
  - Vercel Dashboard → Project → Storage → Marketplace → **Upstash for Redis** → Connect
    (tự thêm `KV_REST_API_URL`, `KV_REST_API_TOKEN`), hoặc
  - tự thêm `UPSTASH_REDIS_REST_URL`, `UPSTASH_REDIS_REST_TOKEN` trong Environment Variables.
- Redeploy sau khi thêm biến môi trường.

> File Excel không được commit (.gitignore). Trên Vercel hãy dùng nút **Upload File Excel**.
