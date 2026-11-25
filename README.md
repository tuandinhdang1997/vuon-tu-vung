# 🌱 Vườn Từ

Web app học từ vựng TOEIC theo buổi: thẻ từ có phiên âm IPA, nghe phát âm, câu hỏi trắc nghiệm / nghe / gõ từ, và ôn tập ngắt quãng ("tưới cây").

Trang tĩnh thuần HTML/CSS/JavaScript, không cần máy chủ hay bước build, nên chạy được thẳng trên GitHub Pages.

## Nội dung

- 19 buổi học chính, mỗi buổi 30 từ chủ điểm (15 buổi đầu có thêm 20 từ thông thuộc).
- 2 bài bổ trợ: 21 cụm từ hay gặp trong bài thi và 63 liên từ / giới từ.
- Tổng cộng 954 mục từ.

## Hướng dẫn cho người mới

Nút **📖 Hướng dẫn** (ở trang chủ, và trên thanh công cụ khi màn hình rộng) chạy một vòng giới thiệu 12 bước: trang chủ → danh sách từ → thẻ từ → câu hỏi → thanh công cụ. Mỗi bước làm sáng đúng phần đang được giải thích; có thể bấm **Bỏ qua**, dùng phím ← → hoặc Esc. Phần "xem thử" thẻ từ và câu hỏi không ghi gì vào tiến độ học.

Nội dung các bước nằm trong mảng `TOUR` của `js/app.js`.

## Chạy thử trên máy

Mở tệp `index.html` bằng Chrome, Edge hoặc Safari.

## Đưa lên GitHub Pages

1. Tạo một repository **Public** trên GitHub (ví dụ `vuon-tu-vung`).
2. Tải toàn bộ nội dung thư mục này lên nhánh `main` (kéo-thả ở mục **Add file → Upload files**, hoặc dùng `git push`).
3. Vào **Settings → Pages**, mục **Build and deployment** chọn **Deploy from a branch**, nhánh `main`, thư mục `/ (root)`, rồi bấm **Save**.
4. Sau khoảng 1–2 phút trang có ở địa chỉ `https://<tên-tài-khoản>.github.io/<tên-repository>/`.

## Cấu trúc

| Tệp | Vai trò |
|---|---|
| `index.html` | Khung trang |
| `css/style.css` | Giao diện |
| `js/data.js` | Dữ liệu từ vựng (sinh tự động, đừng sửa tay) |
| `js/audio.js` | Nhạc nền, hiệu ứng âm thanh, giọng đọc |
| `js/app.js` | Màn hình, phiên học, lịch ôn tập |
| `tools/build_data.py` | Sinh `js/data.js` từ file Excel |

## Cập nhật từ vựng

Sửa file Excel rồi chạy (cần Python và `openpyxl`):

```
python -I tools/build_data.py "đường/dẫn/tới/file.xlsx"
```

Các chỗ sửa lỗi so với file gốc (phiên âm, nghĩa, chính tả) nằm trong bảng `FIXES` ở đầu script; emoji minh hoạ nằm trong bảng `EMOJI`.

## Ghi chú

- **Phát âm** dùng giọng đọc có sẵn của trình duyệt (Web Speech API), nên chất lượng phụ thuộc thiết bị. Nghe hay nhất trên Edge, Chrome và Safari; có thể chọn giọng và tốc độ trong ⚙️ Cài đặt.
- **Tiến độ** lưu trong trình duyệt của từng người (localStorage), không đồng bộ giữa các thiết bị.
- **Nhạc nền** được tạo trực tiếp bằng Web Audio, không dùng tệp nhạc bên ngoài.
