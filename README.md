# DocxTool – Gộp & tách file Word giữ nguyên định dạng (macOS)

Ứng dụng có giao diện (và dòng lệnh) để **gộp nhiều file .docx thành một** hoặc **tách một file .docx thành nhiều file**, xử lý trực tiếp cấu trúc XML của Word nên không làm mất style, khổ giấy, header/footer, đánh số, ảnh, bảng, chú thích.

## Cài đặt

Cần Python 3.9+ có Tkinter:
- Python từ [python.org](https://www.python.org/downloads/macos/) đã có sẵn Tkinter, hoặc
- Homebrew: `brew install python python-tk`

**Cách 1 – Đóng gói thành ứng dụng .app (khuyên dùng)**
```bash
cd DocxMergeSplit
chmod +x build_app.sh run.command
./build_app.sh
```
Kết quả: `dist/DocxTool.app` → kéo vào Applications. Lần đầu mở: chuột phải → **Open** (app chưa ký số nên Gatekeeper sẽ hỏi).

**Cách 2 – Chạy trực tiếp:** nhấp đúp `run.command` (lần đầu tự cài `lxml`).

## Gộp file

1. Tab **Gộp file** → *Thêm file…*, sắp xếp thứ tự bằng *Lên/Xuống*.
2. File đầu tiên là file gốc (giữ cài đặt chung của tài liệu: theme, cài đặt tài liệu).
3. *Gộp và lưu…*

Quy tắc style khi gộp:

| Trường hợp | Kết quả |
|---|---|
| Cùng tên, định dạng **giống hệt** | Dùng chung một style (không nhân bản) |
| Cùng tên, định dạng **khác** | Style mới tên `Tên_1`, `Tên_2`… (file thứ 3 có style giống `Tên_1` sẽ dùng lại `Tên_1`) |
| Tên chưa có | Chép sang nguyên vẹn |

"Giống hệt" được so theo **định dạng hiệu lực**: tính cả style cha (basedOn), định dạng mặc định của tài liệu (docDefaults), font/màu theme, và nội dung danh sách đánh số gắn với style. Cặp style liên kết (vd. `Heading 1` ↔ `Heading 1 Char`) luôn được xử lý cùng nhau.

Ngoài ra khi gộp: mỗi file thành một section riêng (giữ khổ giấy, hướng trang, lề, header/footer); danh sách đánh số của từng file độc lập; ảnh trùng nhau chỉ lưu một lần; chú thích cuối trang, endnote, comment, bookmark, hyperlink, biểu đồ, SmartArt, đối tượng nhúng đều được chuyển theo và đổi id để không xung đột.

## Tách file

Tab **Tách file** → chọn file → chọn cách tách → *Xem trước* → *Tách và lưu…*. Kết quả nằm trong thư mục `<tên file> (đã tách)`.

| Cách tách | Ý nghĩa |
|---|---|
| Theo đề mục | Cắt trước mỗi đề mục cấp 1..N (Heading hoặc đoạn có outline level). Phần trước đề mục đầu tiên là "Phần mở đầu". |
| Theo ngắt section | Mỗi section một file |
| Theo ngắt trang thủ công | Cắt tại Ctrl/⌘+Enter, "Page break before", ngắt section sang trang mới |
| Theo số trang | Dùng thông tin phân trang Word lưu trong file. Có thể đặt "N trang mỗi file" hoặc khoảng trang `1-3, 4-10` |

- **Số phần mỗi file**: gộp N phần liên tiếp vào một file.
- **Chọn phần** `1-3, 5, 7-9`: mỗi khoảng thành một file (số phần xem ở bảng Xem trước). Phần không chọn sẽ bị bỏ.
- **Tên file**: `{so}`, `{tieude}`, `{ten}`, `{trang}`.
- **Giữ số thứ tự**: danh sách / đề mục có đánh số trong file con tiếp tục số như file gốc (vd. file chương 3 bắt đầu "3." chứ không phải "1.").

Mỗi file con giữ toàn bộ style, theme, cài đặt, header/footer đang có hiệu lực ở đoạn đó; ảnh/chú thích/comment không dùng tới được dọn đi; bookmark và field bị cắt dở được sửa cho hợp lệ.

## Dòng lệnh

```bash
python3 cli.py merge  ket_qua.docx a.docx b.docx c.docx            # --continuous, --no-headers
python3 cli.py preview sach.docx --mode heading --level 1
python3 cli.py split  sach.docx ./ra --mode heading --level 1 --ranges "2-4,5"
python3 cli.py split  sach.docx ./ra --mode page --group 10 --name "{ten} trang {trang}"
```

## Giới hạn cần biết

- **Tách theo số trang** phụ thuộc thông tin phân trang mà Word ghi lúc lưu. File tạo bởi Google Docs/LibreOffice/phần mềm khác có thể không có — hãy mở và lưu lại bằng Word trước. Một đoạn văn kéo dài qua hai trang được giữ nguyên ở trang bắt đầu; bảng không bị cắt giữa chừng.
- Khi gộp, cài đặt cấp tài liệu (theme, chế độ header trang chẵn/lẻ, tương thích) lấy theo file đầu tiên. Nếu các file khác theme, font và màu theme của file sau được đổi sang giá trị cụ thể nên chữ vẫn hiển thị như cũ (hình vẽ/biểu đồ dùng màu theme có thể đổi tông).
- Mục lục (TOC) và tham chiếu chéo cần cập nhật lại sau khi gộp/tách (trong Word: chọn tất cả → F9).
- Không hỗ trợ file `.doc` cũ hay file có mật khẩu — mở bằng Word và lưu thành `.docx` trước.
