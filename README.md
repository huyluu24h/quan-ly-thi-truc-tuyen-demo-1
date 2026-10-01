# V5.25.2 – PDF Hóa học và hình vẽ Word

- PDF: font TimesNewRomanPSMT thông thường không bị gắn nhầm là font MathType. Giữ dấu cách trong câu hỏi Hóa học; glyph mũ/chỉ số nhỏ (8.5pt) vẫn đi qua parser công thức.
- DOCX: ảnh GIF preview chưa cắt được cắt đúng `a:srcRect` của Word nguồn (chỉ một lần) rồi hiển thị theo `wp:extent`, thay vì chỉ trim alpha. DOCX và MathType gốc không bị sửa.
- Đồng bộ hình OOXML cho Upload đề kiểm tra, Ngân hàng và Đảo đề, kể cả luồng đọc lại file nguồn. CSS không còn ép `height:auto!important` làm sai kích thước khung Word.
- Đề đã upload trước đó có metadata/URL preview cũ: upload lại bản DOCX/PDF cần sửa. Không xóa database của các giáo viên.
- Kiểm thử: `python -m pytest tests/test_v5252_regressions.py` trên VPS sau khi cài requirements.

**Phạm vi đã kiểm thử tại môi trường xây dựng:** xử lý bản DOCX thực tế với crop của 3 hình Câu 1/3/4, kiểm tra kết quả/tỷ lệ ảnh; kiểm tra font và chuỗi văn bản từ PDF Hóa học; đo khung ảnh bằng Chromium và biên dịch Python. Chưa thể chạy Flask HTTP end-to-end tại môi trường này do thiếu Flask và không có mạng để cài.

# QuanLyThiTrucTuyen V5.25.0 — NHIỀU GIÁO VIÊN / KHO RIÊNG

QUẢN LÝ THI TRỰC TUYẾN V5.25.0 — NHIỀU GIÁO VIÊN

1. NÂNG CẤP AN TOÀN
Tắt ứng dụng cũ. Sao lưu nguyên thư mục APP_DATA_DIR, nhất là instance/database.db, uploads/, documents/ và instance/rosters/ nếu có. Thay mã nguồn bằng V5.25.0, giữ nguyên APP_DATA_DIR. Khởi động lại. Không xóa database và không xóa file đề đã upload.
Khi khởi động, schema được thêm owner cho exams, teacher_students, teacher_invites, exam_invites, exam_students. Đề cũ có nguồn ngân hàng được quy về giáo viên sở hữu ngân hàng. Những đề cũ không xác định được người tạo: nếu chỉ có một giáo viên không phải admin sẽ giao cho giáo viên đó; nếu có nhiều giáo viên, giao tạm cho admin để kiểm tra và phân quyền tại Quản trị. Đề đã có bài làm không chuyển tùy tiện sang giáo viên khác vì bảo vệ kết quả học sinh.

2. GIÁO VIÊN
Đăng ký /register/teacher; admin vào /admin duyệt tài khoản. Mỗi giáo viên đăng nhập bằng tài khoản riêng, vào Giáo viên > Quản lý lớp/nhóm & học sinh của tôi, tạo link mời theo lớp/nhóm. Sao chép link đầy đủ và gửi đến học sinh. Giáo viên có kho đề, Tài liệu, Ngân hàng, Đảo đề và Tạo đề của riêng mình; link riêng /teacher/exam/.../invite chỉ hiện khi đề đã xuất bản.

3. HỌC SINH
Đăng ký tài khoản (hoặc dùng tài khoản cũ) rồi đăng nhập. Mở link lớp/nhóm mà giáo viên đã gửi, bấm Xác nhận tham gia. Từ đó chỉ thấy tài liệu, thư mục luyện đề và đề kiểm tra của giáo viên mình đã tham gia. Nếu giáo viên chia sẻ link riêng cho một đề, học sinh cần chấp nhận link đó trước khi đề xuất hiện. Một học sinh có thể tham gia nhiều giáo viên, mỗi kho vẫn tách biệt.

4. DANH SÁCH EXCEL
Mỗi giáo viên có file riêng dưới instance/rosters. Với tài khoản HS mới tạo bởi Excel, mật khẩu tạm được tạo ngẫu nhiên; giáo viên tải danh sách tài khoản/mật khẩu tạm trên trang /teacher để phát riêng, rồi xóa tệp mật khẩu tạm. Học sinh có tài khoản từ trước sẽ không bị giáo viên khác đặt lại mật khẩu; em cần tự chấp nhận link mời của giáo viên mới. Học sinh và giáo viên có trang Đổi mật khẩu.

5. CẤU HÌNH AN TOÀN KHI ĐƯA LÊN MẠNG
APP_ENV=production; ENABLE_DEMO_ACCOUNTS=0; ADMIN_INITIAL_PASSWORD và SECRET_KEY phải ngẫu nhiên, riêng tư, không dùng giá trị mẫu. Dùng HTTPS; APP_DATA_DIR phải là vùng dữ liệu bền vững, không public; không xuất bản tập tin database.db hoặc tệp tài khoản tạm. Link 127.0.0.1 chỉ dùng trên máy chủ; để học sinh từ Internet vào được, phải triển khai website trên domain/host có thể truy cập bên ngoài.

6. KHÔNG ĐỔI BỘ XỬ LÝ TOÁN
Bản này kế thừa phần Word nguồn bảo toàn MathType/OOXML và PDF nguồn xuất Word LaTeX của V5.24.21, cùng các quy tắc đảo đề và đáp án cũ. Các thay đổi tập trung tài khoản/quyền và chia sẻ.

7. KIỂM THỬ
Trong môi trường cài requirements.txt và pytest, chạy:
  python -m pytest tests/test_teacher_spaces.py -q
Thử 2 giáo viên + 2 học sinh với dữ liệu riêng, tạo link, xác nhận học sinh không thể mở tài nguyên từ ID/link của giáo viên khác. Các kiểm tra tĩnh, migration SQL và Jinja đã chạy khi đóng gói; HTTP/Windows Word/MathType chưa được chạy toàn bộ trong môi trường tạo gói do không có Flask/Office ở đây.


---

# QuanLyThiTrucTuyen V5.24.21

- Word nguồn sau đảo giữ MathType/hình bằng OOXML gốc nhưng được chèn lại **khung tiêu đề Sở/Trường/Kỳ thi/Môn/Thời gian/Mã đề** từ metadata giáo viên.
- Trang cuối Word có **BẢNG ĐÁP ÁN CÁC MÃ ĐỀ** dạng ma trận, gồm Phần I/II/III và tất cả mã đề.
- PDF → Word MathType: gộp các block LaTeX liên tiếp để giảm mạnh số marker TeX, chuẩn hóa cú pháp cho MathType 7, rồi Toggle theo batch 8 công thức. Mỗi batch mở/đóng một Word.Application riêng để không tích lũy cửa sổ MathType.
- Có lock chống hai lượt tải Word MathType chạy đồng thời khi người dùng bấm nhiều lần. Không dùng pywin32.
- PDF LaTeX: tăng timeout XeLaTeX và tự tìm MiKTeX ở các đường cài phổ biến trên Windows; hỗ trợ `XELATEX_PATH`.
- Gunicorn timeout tăng để đủ thời gian cho đề PDF nhiều công thức.


# QuanLyThiTrucTuyen V5.24.16

## V5.24.16 — sửa trọn vẹn lỗi upload Word + Permission denied Windows
- LibreOffice DOCX→HTML chạy với profile tạm riêng và file nguồn tạm tên ASCII để tránh profile lock/path Unicode trên Windows.
- Nếu LibreOffice vẫn không xuất HTML, hệ thống tự động đọc OOXML trực tiếp thay vì dừng upload.
- File Word xuất từ PDF/trộn đề trên Windows dùng `%TEMP%/QuanLyThiTrucTuyen_exports`, tên UUID duy nhất, không còn ghi đè vào thư mục `D:\...\exports`.
- DOCX nguồn không bị ghi lại: MathType OLE/hình/lời giải vẫn được clone từ package Word gốc khi tải bản đảo.


- Sửa lỗi Windows `Errno 13 Permission denied` khi xuất `Tat_ca_ma_de_*.docx`: mỗi lần xuất dùng file vật lý có UUID riêng, không ghi đè file cũ đang bị Word/MathType/trình tải giữ khóa.
- Áp dụng cho cả **Tất cả mã đề** và **từng mã đề Word MathType**.
- Tên file người dùng nhận về vẫn gọn (`Tat_ca_ma_de.docx`, `De_<ma>_MathType.docx`).
- File export cũ hơn 24 giờ được dọn best-effort; file đang bị Windows khóa sẽ được bỏ qua, không làm hỏng request mới.

QuanLyThiTrucTuyen V5.24.14

Cập nhật chính:
- Word nguồn: đồng bộ kích thước hình theo từng câu từ OOXML wp:extent/VML gốc; file tải về vẫn clone trực tiếp OOXML/OLE MathType.
- PDF nguồn: Word xuất ra đặt LaTeX có delimiter và trên Windows tự gọi Microsoft Word + MathType Toggle TeX để chuyển thành Equation.DSMT4 thật. Không dùng ảnh/OMML làm giả MathType.
- Word tải về có Lời giải chi tiết của từng câu trước bảng đáp án.
- Windows cần Microsoft Word + MathType 7 hoạt động; không cần pywin32. Hệ thống gọi Word/MathType qua cscript.exe + COM có sẵn của Windows.

## V5.24.9 (2026-09-12)
- Sửa parser Đảo đề Word nhận cả tiêu đề **Phần 1/2/3/4** và **Phần I/II/III/IV**. Nhờ đó Câu 1 của Phần 2 không còn bị hiểu nhầm thành Câu 13 của Phần 1; các ý a/b/c/d được xử lý đúng kiểu đúng-sai.
- Sửa PDF LaTeX cho Windows/MiKTeX: ưu tiên **Times New Roman**, tự động fallback DejaVu Serif/Latin Modern Roman nếu máy không có font. Không còn phụ thuộc cứng vào DejaVu Serif.
- Giữ nguyên luồng Word MathType/OLE và PDF dựng trực tiếp từ LaTeX của V5.24.8.

## V5.24.8 (2026-09-12)
- Sửa liên kết `v:shape/@id` ↔ `OLEObject/@ShapeID` khi clone Word sau đảo, tránh hình/MathType hiển thị sai; kiểm tra hash toàn bộ `word/media` và `word/embeddings` trước khi cho tải.
- Word sau đảo tiếp tục dùng package DOCX gốc nên MathType vẫn là OLE MathType chỉnh sửa được; bytes hình minh họa được bảo toàn.
- Thêm **Xuất PDF LaTeX** cho từng mã và tất cả mã đề: PDF dựng trực tiếp từ block LaTeX/XeLaTeX, không convert từ Word MathType.
- Thêm **Tải danh sách học sinh luyện đề** dạng Excel gồm sheet chi tiết từng lượt và sheet tổng hợp theo học sinh.
- Docker bổ sung XeLaTeX để biên dịch PDF LaTeX khi triển khai Internet.


## V5.24.7 (2026-09-11)
- Sửa lỗi upload đề trộn Word báo 293 vị trí / 283 công thức: nhận diện `ObjectN` là MathType/OLE và `Picture/Shape/Image` là hình minh họa.
- Tải từng mã đề DOCX dùng cùng cơ chế clone XML/OLE như tải tất cả mã đề, giữ MathType editable từ Word gốc.
- Tăng độ chặt tạo đề AI: ép tổng số phần tử đúng số câu yêu cầu bằng JSON Schema và bám lớp/chủ đề/mức độ/dạng toán/ràng buộc giáo viên nhập.
- Giữ kiểm tra cấu trúc Phần I/II/III/IV trước khi hiển thị bản nháp.

# V5.24.6 — Đổi tên thư mục, chỉ số cực trị PDF và topbar đồng bộ

- Sửa lỗi `Undefined namespace prefix` khi tải Word sau khi đảo đề; bộ xuất
  dùng XPath theo namespace URI/local-name và vẫn giữ MathType, hình cùng câu.
- Tài liệu tiếp tục mở trong tab hiện tại nhưng hiển thị lại menu dọc bên trái.
- Tạo đề AI dùng Responses API với Structured Outputs, đọc nội dung ma trận
  DOCX/PDF/XLSX/ảnh và kiểm tra đủ số câu, phương án, đáp án của từng phần.
- Thống kê luyện đề có nút xóa từng lượt; không xóa tài khoản học sinh và không
  cho xóa khi học sinh còn đang làm.
- Chế độ production có cookie an toàn, kiểm tra nguồn yêu cầu ghi, header bảo
  mật, tắt tài khoản demo và yêu cầu mật khẩu quản trị ban đầu qua môi trường.

## Nâng cấp từ V5.24.4

Chỉ thay mã nguồn; giữ nguyên thư mục được cấu hình bởi `APP_DATA_DIR` để bảo
toàn cơ sở dữ liệu, đề, tài liệu và hình ảnh đã tải lên.

- Tệp tài liệu được mở ngay trong tab hiện tại, tự bắt đầu ở đầu trang xem và
  có nút trở về danh sách; không sinh thêm tab trình duyệt.
- Bảng theo dõi luyện đề hiển thị lớp suy ra từ tên thư mục Ngân hàng, đồng
  thời vẫn giữ nguyên tên thư mục để giáo viên đối chiếu.
- Nâng cấp bộ đọc PDF Word nhiều lớp: ghép đúng số mũ/chỉ số, dấu căn vẽ bằng
  đường nét, ngoặc co giãn, dấu đạo hàm và các ký hiệu `α`, `Δ`, `ℝ`.
- Chỉ hình minh họa thật được giữ dạng ảnh; ảnh nhúng giữ byte gốc, hình vectơ
  được render nét ở độ phân giải cao và đúng vị trí trong câu.
- Bổ sung cấu hình Docker Compose, mẫu biến môi trường và hướng dẫn Render để
  sẵn sàng đưa website lên Internet với vùng dữ liệu bền vững.

# V5.24.3 — Toán tử chuẩn và theo dõi luyện đề thời gian thực

- `lim`, `min`, `max` được chuẩn hóa thành toán tử đứng; điều kiện/miền xét nằm phía dưới như MathType nguồn.
- Ký hiệu cực đại, cực tiểu `y_CĐ`, `y_CT` được giữ trọn trong chỉ số dưới.
- Ngân hàng đề có nút Chọn tất cả/Bỏ chọn tất cả.
- Giáo viên theo dõi trực tiếp học sinh đang luyện đề và xem thống kê điểm, thời gian, vi phạm.
- Trang xem tài liệu ẩn menu dọc, giữ trang tài liệu cùng thanh menu ngang.

# V5.24.2 — Chuẩn hóa min/max/lim và bố cục công thức PDF

- Ghép `min`, `max`, `lim` với đúng chỉ số dưới trước khi dựng LaTeX, áp dụng
  đồng nhất cho đề bài, phương án và lời giải.
- Khôi phục ngoặc co giãn, tập hợp, hệ phương trình, phân số nhiều tầng và số
  mũ từ các mã font TeX đặc biệt của PDF.
- Loại nét trang trí chạy dọc/ngang toàn trang khỏi cụm hình để không làm mất
  câu hoặc gộp nhiều câu của Phần III vào một ảnh.
- Giữ đúng công thức lồng nhau, ảnh/bảng/đồ thị và đầy đủ 22 câu theo cấu trúc
  12 câu Phần I, 4 câu Phần II, 6 câu Phần III của tệp kiểm thử.
- Giữ nguyên toàn bộ luồng Word/MathType, thư mục, Ngân hàng, Luyện đề và chấm
  điểm đã ổn định ở V5.24.1.

# V5.24.1 — Sửa lỗi phương án trùng khi tải đề vào Ngân hàng

- Chuẩn hóa khóa phương án về đúng `A/B/C/D` trước khi lưu.
- Tự gộp các mảnh phương án bị Word/PDF nhận lặp khóa, không làm mất chữ,
  công thức hoặc hình ảnh.
- Chặn lỗi `UNIQUE constraint failed: bank_options...` và giữ nguyên toàn bộ
  luồng thư mục, MathType/LaTeX/MathJax, luyện đề và chấm điểm.

# V5.24.0 — Thư mục hai tầng cho Tài liệu, Ngân hàng và Luyện đề

- Các thư mục hiển thị dạng ô chức năng, căn giữa vùng nội dung giống menu Giáo viên/Học sinh.
- Người dùng bấm mở thư mục rồi mới thấy danh sách tài liệu hoặc đề nằm bên trong.
- Giáo viên upload trực tiếp trong thư mục đang mở; dữ liệu được gắn đúng thư mục tự động.
- Học sinh duyệt Tài liệu và Luyện đề theo đúng cùng cấu trúc thư mục do giáo viên tạo.
- Giữ nguyên xử lý Word/PDF, MathType → LaTeX ngầm → MathJax, làm bài và chấm điểm.

# V5.23.9 — PDF nhiều tầng và thư mục Tài liệu/Ngân hàng

- Sửa Câu 3, Câu 9 Phần I: bảng biến thiên/đồ thị chỉ giữ phần hình, không còn kéo chữ xung quanh vào ảnh.
- Sửa Câu 3 Phần III và các trường hợp tương tự: khôi phục đầy đủ phân số, căn, tích phân, số mũ trong đề, đáp án và lời giải PDF.
- Dòng chỉ có công thức LaTeX ngầm không còn bị bỏ qua; MathJax tiếp tục hiển thị công thức toán cho giáo viên/học sinh.
- Thêm nút **Tạo thư mục** trong Tài liệu và Ngân hàng đề; giáo viên chọn thư mục ngay khi upload.
- Thư mục Tài liệu cùng các tệp bên trong được đồng bộ tự động sang trang Tài liệu học sinh.
- Thêm **Luyện đề** cho học sinh, tự đồng bộ thư mục/đề từ Ngân hàng và dùng chung cơ chế làm bài, chấm điểm.
- Giữ nguyên luồng Word/MathType và các chức năng đã ổn định ở V5.23.8.

# V5.23.8 — đáp án PDF đúng bố cục và menu ở tâm vùng nội dung

- Tách chính xác thanh phân số khỏi nét phủ của dấu căn trước khi ghép dòng,
  nên tử/mẫu không còn rơi sang đáp án hoặc dòng lời giải kế tiếp.
- Giữ đúng số mũ bên trong căn thức; phục hồi ngoặc tích có hướng và các hệ
  phương trình nhiều dòng trong đáp án/lời giải PDF.
- Mở rộng vùng hình học để giữ cả nhãn và số đo nằm sát ngoài hình, đồng thời
  không để các nhãn ấy lẫn vào nội dung câu hỏi.
- Căn các nhóm chức năng Giáo viên, Học sinh và Tiện ích theo cả chiều ngang
  lẫn chiều dọc ở tâm vùng nội dung chính; giao diện điện thoại vẫn xếp gọn.

# V5.23.7 — mũi tên véc-tơ đúng độ dài và căn giữa menu chức năng

- Véc-tơ một ký tự như `u`, `v` tiếp tục dùng `\vec`; véc-tơ tạo bởi hai
  điểm trở lên như `MN`, `SA`, `SC` dùng `\overrightarrow` để mũi tên phủ
  hết tên véc-tơ trên màn hình giáo viên và học sinh.
- Chuẩn hóa cả công thức PDF, Word/MathType và dữ liệu đề đã tải từ bản cũ;
  không cần sửa thủ công từng câu hay tải lại đề chỉ để kéo dài dấu véc-tơ.
- Căn giữa các thẻ chức năng con của Giáo viên, Học sinh và Tiện ích trong
  vùng nội dung; đồng thời giữ submenu kiểu cũ ở giữa nếu dữ liệu cũ còn dùng.
- Giữ nguyên toàn bộ xử lý công thức, hình vẽ, lời giải và chức năng của
  V5.23.6.

# V5.23.6 — hoàn thiện ký hiệu và hình vẽ trong đề PDF

- Khôi phục ký hiệu theo đúng mã font TeX của PDF: dấu vectơ, dấu góc,
  tích phân có cận, ngoặc co giãn và dấu khác; LaTeX vẫn chỉ được lưu ngầm.
- Tự nhận diện đồ thị, hình học, bảng biến thiên và bảng số liệu được PDF vẽ
  bằng nét vectơ; render riêng ở độ phân giải cao, đúng tỷ lệ và đúng vị trí.
- Không để chữ trên trục/hình vẽ lọt vào phương án kế bên; sửa dứt điểm các
  ký tự thừa ở cuối đáp án và hệ phương trình nhiều cột.
- Khôi phục khoảng cách chữ theo tọa độ glyph, áp dụng cho cả đề, phương án và
  lời giải; giữ nguyên luồng Word/MathType và mọi chức năng của V5.23.5.
- Kiểm thử với đề `SP-DuAn-HTN-2025-2026-Đề-01.pdf`: đủ 22 câu theo ba phần,
  đủ đáp án/lời giải và 13 hình/bảng vectơ đúng vị trí.

# V5.23.5 — lời giải chi tiết theo từng câu từ Word/PDF

- Tự tách nội dung bắt đầu từ nhãn `Lời giải` của mỗi câu và lưu riêng,
  không để lời giải lẫn vào đề hiển thị cho học sinh.
- Giáo viên có nút `Lời giải chi tiết` tại từng câu; bấm nút mới mở nội dung
  tương ứng, gồm văn bản, công thức MathJax và hình minh họa.
- Tự đọc đáp án trong lời giải cho Phần I (`Chọn B`), Phần II
  (`a) Đúng/Sai`) và Phần III (`Trả lời: ...` / `Đáp án: ...`).
- Sửa lỗi kiểm tra cặp `\\left`/`\\right` nhận nhầm `\\rightarrow`, nhờ đó
  giữ đủ công thức vectơ của các đề Word có lời giải.
- Ngân hàng đề và mục Đảo đề cũng lưu lời giải cùng câu hỏi; màn hình học sinh
  không truy vấn và không hiển thị lời giải.

# V5.23.4 — sửa dứt điểm hiển thị công thức vectơ Câu 11

- Thay cặp dấu co giãn `\left...\right` của riêng công thức vectơ PDF bằng
  dấu TeX cơ bản tương thích với bộ MathJax đã đóng gói.
- Nhận cả hai thứ tự ký tự tổ hợp do các bản PyMuPDF trả về (`⃗u` và `u⃗`),
  rồi chuyển thống nhất thành `\vec{u}` trước khi lưu và hiển thị.
- Tự sửa dữ liệu đề đã tải lên: hai công thức độ dài vectơ và công thức góc
  giữa hai vectơ được chuẩn hóa khi giáo viên hoặc học sinh mở đề.
- Tự ghép lại biểu thức góc với `120^{\circ}` nếu bản cũ từng lưu thành hai
  khối do PDF xuống dòng; không cần xóa đề và tải lại.
- Không thay đổi luồng Word/MathType, hình Câu 12, Phần II và các phần đã ổn
  định ở V5.23.3.

# V5.23.3 — hoàn thiện công thức, hình và bố cục đề PDF

- Sửa Câu 11: độ dài vectơ và góc giữa hai vectơ được lưu thành LaTeX cân
  bằng, không còn `Math input error` khi MathJax dựng công thức.
- Sửa Câu 12: hình minh họa nằm cuối đề bài và trước bốn phương án A–D.
- Khôi phục dấu cách bị mất trong lớp chữ PDF dựa trên khoảng cách glyph.
- Phần II hiển thị đúng dạng Đúng/Sai với bốn ý a), b), c), d), giống Word.
- Giữ nguyên luồng Word/MathType và các chức năng đã ổn định ở V5.23.2.

# V5.23.2 — PDF thành LaTeX ngầm, công thức dàn đồng nhất như Word

- Giữ nguyên toàn bộ bộ xử lý Word/MathType, MathJax, hình ảnh, chấm điểm,
  trộn đề, Tiện ích và các chức năng đã ổn định ở V5.23.1.
- PDF số không còn cắt công thức thành ảnh. Hệ thống đọc font và tọa độ từng
  glyph, kết hợp đường kẻ phân số để tái dựng LaTeX chuẩn cho phân số, căn,
  số mũ/chỉ số, tích phân, vectơ, tập hợp và hệ phương trình.
- LaTeX chỉ lưu chạy ngầm, không có dấu `$...$` trong dữ liệu. Màn hình giáo
  viên và học sinh tiếp tục dùng MathJax cục bộ để hiện công thức đồng đều.
- Chỉ hình minh họa thật được lưu dạng ảnh; ưu tiên giữ nguyên bytes nhúng và
  giữ đúng tỷ lệ hiển thị từ PDF.
- Kiểm thử trực tiếp với `DE-0101-THI-THU.pdf`: nhận đủ 22 câu (12/4/6 theo
  ba phần), đủ bốn phương án cho 16 câu trắc nghiệm, 177 công thức LaTeX,
  6 hình minh họa và không còn ảnh công thức.

# V5.23.1 — nâng cấp toàn bộ khu vực Tiện ích

- **Tiện ích** nay là một trang cổng giống khu vực Giáo viên/Học sinh; ba chức
  năng được đặt bên trong bằng các thẻ lớn, không bung dài trên menu trái.
- **PDF sang Word** không còn biến toàn bộ trang thành ảnh: văn bản PDF số được
  tạo thành chữ Word chỉnh sửa được, công thức/ký hiệu nhận diện được chuẩn hóa
  thành LaTeX `$...$`, còn ảnh nhúng ưu tiên giữ nguyên dữ liệu và tỷ lệ gốc.
- PDF scan không có lớp chữ vẫn được giữ nguyên trang ảnh để không mất nội dung.
- **Giáo viên AI** là trợ lý giải Toán: nhận đề bằng chữ hoặc ảnh, chọn THCS/THPT
  và chọn giải chi tiết, gợi ý hoặc lời giải ngắn. Khóa API chỉ đọc từ biến môi
  trường `OPENAI_API_KEY`, không lưu trong mã nguồn.
- **Trò chơi học sinh** có ba trò chơi hoạt động trực tiếp: Thử thách Toán 60
  giây, Bảy sắc cầu vồng luyện trí nhớ, và Chinh phục đỉnh tri thức gồm bốn chặng
  với bộ câu hỏi nguyên bản.
- Giữ nguyên chế độ làm bài không menu trái và toàn bộ luồng MathType → LaTeX
  ngầm → MathJax, hình câu hỏi, chấm điểm và chống gian lận của V5.23.0.

# V5.23.0 — làm bài tập trung, Tiện ích và chạy trên Internet

- Khi học sinh bấm **Làm bài**, menu dọc chính bên trái được ẩn hoàn toàn và
  vùng đề tự mở rộng. Menu vẫn giữ nguyên ở các trang khác.
- Thêm nhóm **Tiện ích** ngay phía trên **Liên hệ**, gồm PDF sang Word, Giáo
  viên AI và Trò chơi học sinh; mỗi mục đều có biểu tượng riêng.
- PDF sang Word dựng từng trang ở độ phân giải cao để bảo toàn công thức toán,
  hình vẽ và bố cục của PDF gốc.
- Thêm trò chơi Toán học 60 giây với ba mức độ, chạy ngay trong trình duyệt.
- Thêm `Dockerfile`, Gunicorn và `render.yaml` để triển khai website công khai.
  Biến `APP_DATA_DIR` đưa cơ sở dữ liệu, đề, hình, tài liệu và tin tức vào ổ đĩa
  bền vững trên máy chủ.
- Giữ nguyên toàn bộ luồng MathType → LaTeX ngầm → MathJax, chất lượng hình câu
  hỏi, đáp án, chấm điểm và chống gian lận của V5.22.14.

Xem hướng dẫn tại `DEPLOY_RENDER.md`.

# V5.22.14 — sửa hình câu hỏi bị mất sau khi cập nhật

- Ảnh lỗi có biểu tượng vỡ và chữ “Hình câu hỏi” là do dữ liệu đề còn URL nhưng tệp ảnh cũ trong `static/question_media` không còn sau khi thay mã nguồn.
- Ảnh câu hỏi tải lên từ bản này được lưu bền vững trong `uploads/question_media`, cùng vùng dữ liệu mà giáo viên giữ lại khi nâng cấp.
- Website vẫn đọc được URL ảnh của các bản cũ. Nếu tệp ảnh cũ thật sự bị mất nhưng file DOCX/PDF gốc còn trong `uploads`, hệ thống tự trích lại hình và cập nhật đúng câu/phương án khi giáo viên hoặc học sinh mở đề.
- Chỉ thay các block hình minh họa bị gãy; nội dung chữ, đáp án và toàn bộ luồng MathType → LaTeX → MathJax của V5.22.13 được giữ nguyên.
- Áp dụng cho đề tải trực tiếp, Ngân hàng đề và màn hình Đảo đề.

# V5.22.13 — sửa lỗi lộ nguyên `\(...\)` trên màn hình

- Ảnh người dùng gửi xác nhận LaTeX đúng nhưng MathJax không khởi động; timer của V5.22.12 vì vậy đã mở lại chuỗi nguồn `\(...\)`.
- Khôi phục đúng bộ nạp của V5.22.0 và cố định phiên bản ổn định: `cdn.jsdelivr.net/npm/mathjax@3.2.2/es5/tex-mml-chtml.js` là nguồn chính.
- Nếu CDN không tải được, cùng thẻ script tự chuyển sang file MathJax cục bộ đã đóng gói.
- Bỏ hoàn toàn callback `defaultPageReady` tùy biến; chỉ chờ `startup.promise` của MathJax rồi gắn trạng thái hiển thị.
- Kéo thời gian bảo vệ lên 12 giây để không mở mã LaTeX khi thư viện chỉ đang tải chậm.
- Không thay đổi MTEF-py, LaTeX chuẩn, Câu 10B, đáp án, điểm hoặc các chức năng khác.

# V5.22.12 — LaTeX chỉ chạy ngầm, màn hình chỉ hiện công thức đã dàn

- Giữ nguyên dữ liệu LaTeX chuẩn của V5.22.11 nhưng thêm trạng thái `mathjax-loading`: vùng công thức chỉ được mở sau khi MathJax `pageReady` hoàn tất.
- Giáo viên và học sinh không nhìn thấy chuỗi `\dfrac`, `\sqrt`, `\left` hoặc dấu phân cách trong lúc tải trang.
- MathJax vẫn dùng file cục bộ; URL thư viện và CSS có phiên bản mới để trình duyệt không dùng cache cũ.
- Có hai lớp chống trắng màn hình: `onerror` của file MathJax và timer 4 giây đều bỏ trạng thái chờ nếu thư viện không khởi động.
- Không quay lại ảnh PNG/SVG công thức; chỉ hình minh họa thật tiếp tục là ảnh.
- Giữ bộ sửa Câu 10B và không thay đổi nội dung, đáp án, điểm hoặc các chức năng khác.

# V5.22.11 — MathType thành LaTeX ngầm, MathJax hiển thị đồng nhất

- Sửa đúng yêu cầu: công thức MathType/OMML từ Word được chuyển thành LaTeX chuẩn và ảnh công thức không còn được lưu trong block hiển thị.
- Giáo viên và học sinh chỉ thấy kết quả MathJax dàn từ LaTeX; mọi công thức dùng chung font, tỷ lệ và baseline, không trộn PNG/SVG lớn nhỏ.
- MathJax được đóng gói cục bộ, bật tự dàn trang như cơ chế ổn định của V5.22.0 và không phụ thuộc Internet.
- Đề cũ tự xóa ba trường ảnh `fallback_url`, `fallback_width_px`, `fallback_height_px` khi giáo viên hoặc học sinh mở đề; nội dung, LaTeX, đáp án và điểm không thay đổi.
- Giữ bộ sửa cây MTEF của Câu 10B, chuẩn hóa `\dfrac`, `\sqrt`, hàm toán và ký tự MathType cũ.
- Chỉ hình minh họa thật tiếp tục được giữ nguyên byte và kích thước Word.

# V5.22.10 — sửa công thức chồng nét/ký tự

- Sửa đúng lỗi trong ảnh người dùng gửi: WMF/EMF MathType đổi sang SVG làm mất tọa độ glyph của font cũ, khiến chữ và nét công thức chồng lên nhau.
- Khôi phục chính xác đường hiển thị của V5.22.0: dùng ảnh preview do Word/LibreOffice xuất trực tiếp, không dùng SVG công thức.
- LaTeX chuẩn từ MTEF-py vẫn được lưu đầy đủ chạy ngầm; giáo viên và học sinh chỉ nhìn thấy preview MathType.
- Các đề đã upload bằng V5.22.8/V5.22.9 được tự thay `fallback_url` SVG bằng preview PNG khi mở, nếu DOCX gốc còn trong `uploads`.
- Giữ nguyên bộ sửa Câu 10B, `\dfrac`, `\sqrt`, hình minh họa gốc và toàn bộ chức năng khác.

# V5.22.9 — LaTeX chạy ngầm, MathType hiển thị như V5.22.0

- Đối chiếu trực tiếp tệp V5.22.0 do người dùng cung cấp: bản này nhìn đúng vì preview MathType gốc vẫn là lớp công thức người dùng thấy.
- Mọi công thức MathType/OMML trong DOCX vẫn được MTEF-py giải mã, chuẩn hóa và lưu trong block `latex`; mã nguồn không xuất hiện trên màn hình.
- Preview MathType vector/ảnh gốc là lớp hiển thị chính cho giáo viên và học sinh. MathJax chỉ hiện khi preview bị thiếu hoặc tải lỗi.
- Phân số được chuẩn hóa thành `\dfrac{...}{...}`, căn thành `\sqrt{...}`, hàm toán và ký hiệu cũ được chuẩn hóa trước khi lưu.
- Giữ bộ sửa cây MTEF cho Câu 10B: nhận đúng dấu prime của Euclid Symbol, khôi phục mẫu phân thức và khử nền chroma-key của preview.
- Hình minh họa thật tiếp tục lấy từ DOCX gốc, không nén lại và giữ kích thước Word.
- Không thay đổi menu, tài khoản, upload/xuất bản, trộn đề, đáp án, chấm điểm, kết quả hoặc chống gian lận.

# V5.22.8 — V5.22.0 hiển thị + bộ sửa Câu 10B

- Lấy nguyên cơ chế hiển thị đã ổn định ở V5.22.0 làm nền: MathJax dựng công thức từ LaTeX, preview MathType vector gốc hiển thị an toàn cho đến khi MathJax hoàn tất.
- LaTeX chuẩn vẫn là dữ liệu chính chạy ngầm; màn hình giáo viên và học sinh không hiện mã `$...$` và không có vùng công thức trắng.
- Giữ bộ sửa riêng lỗi đạo hàm/phân thức Câu 10B, cùng chuẩn hóa `\dfrac`, `\sqrt`, hàm toán và ký tự MathType cũ.
- Riêng Câu 10B dùng bản preview đã khử nền chroma-key magenta thay cho SVG cũ chứa ký tự Symbol lỗi; nền trong suốt, nét công thức màu đen, không còn nhỏ/ám đỏ.
- Hình minh họa thật tiếp tục lấy từ DOCX gốc, giữ nguyên byte/kích thước; ảnh preview công thức chỉ là lớp hiển thị dự phòng.
- Không thay đổi menu, tài khoản, upload, xuất bản, trộn đề, đáp án, chấm điểm, kết quả hoặc chống gian lận.
- Đề Word đã tạo bằng V5.22.5–V5.22.7 được tự bổ sung preview khi mở, nếu file DOCX gốc vẫn còn trong thư mục `uploads`; chỉ thêm trường hiển thị, không sửa nội dung/đáp án.

# V5.22.7 — sửa công thức không xuất hiện trên màn hình

- Sửa nguyên nhân vùng công thức trắng: CSS cũ dùng `visibility:hidden` và chỉ mở lại khi JavaScript gắn đúng class `math-rendered`.
- MathJax vẫn tự dàn trang như bản hoạt động tốt; các lần dựng bổ sung luôn chờ `startup.promise` hoàn tất.
- CSS nhận trực tiếp `mjx-container`, vì vậy kết quả công thức hiện ngay cả khi class hỗ trợ được gắn chậm.
- Thêm phiên bản vào URL `style.css` để trình duyệt không dùng lại CSS đã lưu của V5.22.6.
- LaTeX vẫn chỉ là dữ liệu nội bộ; ảnh công thức không được dùng làm phương án hiển thị.

# V5.22.6 — MathType chuyển LaTeX ngầm, màn hình chỉ hiện công thức toán

- Kết hợp cơ chế MathJax tự dàn trang đã hoạt động ổn định ở V5.22.2 với bộ chuyển đổi MTEF/OMML đã sửa ở V5.22.5.
- Công thức MathType chỉ dùng dữ liệu OLE/MTEF để tạo LaTeX chuẩn; không sao chép ảnh công thức và không lưu `fallback_url`.
- Giáo viên và học sinh chỉ nhìn thấy kết quả MathJax, không nhìn thấy mã LaTeX hoặc dấu `$...$`.
- Phân số chuẩn hóa thành `\dfrac{...}{...}`; căn thường chuẩn hóa thành `\sqrt{...}`; đơn vị `\rm` cũ chuyển thành `\mathrm{...}`.
- Nếu số vị trí MathType và số công thức MTEF/OMML không khớp, hệ thống dừng upload và báo rõ số lượng thay vì ghép nhầm công thức.
- Chỉ hình minh họa thật của đề mới đi qua bộ xử lý ảnh và giữ kích thước gốc.

# V5.22.5 — Khắc phục màn hình bị trống công thức MathType

- LaTeX vẫn chỉ là dữ liệu nội bộ, không hiển thị dấu `$...$` cho giáo viên hoặc học sinh.
- Trình hiển thị MathJax được khởi động chủ động và tự xử lý cả công thức được thêm sau khi trang đã mở.
- Mỗi công thức Word giữ một ảnh MathType dự phòng ẩn; ảnh này chỉ xuất hiện nếu MathJax tải chậm hoặc bị chặn, vì vậy không còn khoảng trống công thức.
- Chuẩn hóa mọi phân số thành `\dfrac{...}{...}`; sửa dạng căn MTEF-py `\sqrt[] {...}` thành `\sqrt{...}`; chuẩn hóa các hàm `\sin`, `\cos`, `\log`, `\ln`.
- Tệp kiểm thử thực tế: 141/141 MathType chuyển thành LaTeX, 16 phân số dùng `\dfrac`, 13 căn dùng `\sqrt{...}`, không còn `\frac`, `\sqrt[]`, ký tự lỗi hoặc dấu `$` trong dữ liệu lưu.
- Màn hình kiểm tra của giáo viên và màn hình làm bài của học sinh dùng chung một bộ dựng công thức, nên kết quả hiển thị thống nhất.

# V5.22.4 — MathType sang LaTeX ngầm, hiển thị công thức hoàn chỉnh

- Mọi MathType OLE/Word Equation trong DOCX được chuyển thành LaTeX chuẩn không có dấu phân cách `$...$` và lưu trong block `latex`.
- Giáo viên và học sinh chỉ thấy công thức được MathJax cục bộ dàn kiểu trực quan như MathType; mã LaTeX không lộ ra màn hình.
- Sửa nguyên nhân Câu 10B: MTEF-py đọc ký tự prime của font Euclid Symbol thành `¢`, làm phần sau bị lồng vào số mũ và phân thức mất mẫu số. Bộ sửa cây MTEF nay đóng đúng ô prime và đưa dòng mẫu về đúng phân thức.
- Chỉ hình minh họa thật đi qua luồng ảnh: PNG/JPG được sao chép nguyên byte; EMF/WMF xuất SVG vector; giữ kích thước Word.
- Với file kiểm thử: 141/141 công thức chuyển thành LaTeX, MathJax render 141/141 không lỗi, 0 công thức ảnh, 2 hình minh họa và nhận đúng 21 câu.

# V5.22.3 — sửa Math input error và tải đồng thời thanh ngang

- Sửa ký tự nội bộ MathType U+EF01 làm công thức `A, B` báo `Math input error`.
- Nếu MathJax phát hiện công thức lỗi, hệ thống ẩn thông báo lỗi và hiện ảnh MathType gốc.
- Bootstrap Icons được đóng gói cục bộ; đồng hồ khởi tạo ngay khi topbar xuất hiện.

# V5.22.2 — không hiển thị mã LaTeX

- LaTeX chỉ được lưu và xử lý ngầm, không hiển thị trực tiếp cho giáo viên/học sinh.
- Khi MathJax dựng xong, màn hình chỉ hiện công thức toán hoàn chỉnh.
- Nếu trình dựng công thức gặp lỗi, hệ thống dùng ảnh MathType gốc thay thế nên không lộ `$...$`.

# V5.22.1 — thanh ngang và công thức hoàn chỉnh

- Thanh menu ngang hiển thị ngày giờ, họ tên người đăng nhập, sáng/tối và bảng chọn màu.
- Dữ liệu công thức cũ có `$...$` được tự chuẩn hóa, không hiện mã LaTeX ra màn hình.
- MathJax được đóng gói cục bộ nên công thức vẫn hiển thị khi máy chủ không truy cập CDN.

# V5.22.0 — công thức Word sắc nét trên web

- Khi upload DOCX, hệ thống đọc trực tiếp MathType OLE bằng MTEF-py và lưu công thức ở dạng LaTeX.
- Công thức Word Equation (OMML) cũng được chuyển sang LaTeX.
- MathJax hiển thị công thức vector trên màn hình giáo viên và học sinh.
- Nếu gặp công thức chưa hỗ trợ, hệ thống tự giữ ảnh công thức cũ thay vì làm mất nội dung.
- PDF không còn dữ liệu MathType gốc nên tiếp tục dùng bộ xử lý PDF hiện có và ảnh công thức độ phân giải cao.

# V5.9 FIXED

## Khối 1 — Word/MathType/Hình
- MathType chỉ nhận khi OLE ProgID = Equation.DSMT4.
- Hình WMF không phải MathType được giữ là hình và render PNG nếu có LibreOffice.
- Lấy kích thước hình từ Word DrawingML/VML để hiển thị đúng kích thước và tỷ lệ.

## Khối 2 — Đáp án
- Phần II nhận Đáp án: ĐSĐS / SĐSĐ / Đ,S,Đ,S.
- Phần III nhận Đáp số/Đáp sô.

## Khối 3 — Random
- Mỗi phần đánh số lại từ Câu 1.
- Phần II chỉ đảo câu, KHÔNG đảo A-D.
- Hình, MathType và phương án thuộc question_id nên đi cùng câu.

## Khối 4 — Kết quả
- Giáo viên có trang kết quả riêng cho từng đề.
- Có điểm từng phần, tổng điểm, số lỗi và chi tiết gian lận.
- Có nút tải Excel riêng cho từng đề.

## Khối 5 — Học sinh
- Chỉ đề published=1 mới hiện ở trang học sinh.
- Khi mở đề, create_student_exam tạo phiên làm bài và thứ tự câu.

## V5.9.1 — MathType + học sinh Demo
- MathType: cắt khoảng trắng dư của ảnh render trước khi hiển thị, sau đó chuẩn hóa chiều cao 21px để nhìn cân với Times New Roman 13.5pt.
- Đăng nhập học sinh: có thể dùng Mã học sinh trực tiếp làm tài khoản.
- Tài khoản kiểm thử: Mã HS `2305089483`, mật khẩu `2305089483`, Họ tên `Học sinh Demo`.
- File Excel `Danh-sách-lop-12a(4).xls` đã được kiểm tra: danh sách có sẵn dòng `2305089483 | Học sinh Demo`, vì vậy không cần thêm học sinh này lần nữa.

# V5.10 — điểm Phần II + chống gian lận ổn định

## Các thay đổi
1. **Công thức toán**: tăng chiều cao hiển thị và giữ chiều rộng tự động để kích thước thị giác cân với Times New Roman 13.5pt.
2. **Phần II**: giáo viên nhập điểm chung cho 4 ý A/B/C/D. Tổng A+B+C+D phải đúng bằng điểm Phần II. Các câu Phần II có cùng tổng điểm; mỗi ý có thể có điểm khác nhau.
3. **Phần I và III**: điểm mỗi câu bằng nhau.
4. **Chấm Phần II**: học sinh đúng từng ý được cộng đúng điểm của ý đó, không còn bắt buộc đúng cả 4 ý mới được điểm.
5. **Chống gian lận**: chờ các sự kiện chống gian lận đang gửi trước khi submit; hỗ trợ `pagehide`/`sendBeacon` khi học sinh rời trang chưa nộp; lưu snapshot lỗi lúc nộp.
6. **Xuất bản**: lưu đáp án không tự xuất bản; giáo viên phải bấm nút Xuất bản.
7. **Học sinh Demo**: Mã `2305089483`, Họ tên `Học sinh Demo`, mật khẩu mặc định `2305089483`.

## Chú thích app.py
- KHỐI MIGRATION ĐIỂM PHẦN II: thêm cột và chuẩn hóa dữ liệu đề cũ.
- KHỐI ĐIỂM GIÁO VIÊN: lưu tổng Phần II và điểm A/B/C/D.
- KHỐI ĐIỂM — TÍNH ĐIỂM TỪNG CÂU: tính điểm I/II/III theo cấu hình.
- KHỐI CHỐT CHỐNG GIAN LẬN: lưu snapshot lỗi trước khi đóng phiên.
- KHỐI 8 — KẾT QUẢ THEO TỪNG ĐỀ + EXCEL: xem và xuất kết quả từng học sinh.


## V5.11 — Hai phương án chấm Phần II

Mỗi câu Phần II có 4 ý và điểm tối đa của câu là 1 điểm.

**Phương án 1 — cộng 0,25 điểm cho mỗi ý đúng**
- 0 ý đúng = 0 điểm
- 1 ý đúng = 0,25 điểm
- 2 ý đúng = 0,50 điểm
- 3 ý đúng = 0,75 điểm
- 4 ý đúng = 1,00 điểm

**Phương án 2 — theo số lượng ý đúng**
- 0 ý đúng = 0 điểm
- 1 ý đúng = 0,10 điểm
- 2 ý đúng = 0,25 điểm
- 3 ý đúng = 0,50 điểm
- 4 ý đúng = 1,00 điểm

Hàm `score_part2_question()` trong `app.py` là nơi tập trung logic chấm.
`PART2_SCORING_MODE = 1` chọn Phương án 1; `= 2` chọn Phương án 2.


## V5.11.3 — sửa Xem kết quả và PDF
- Khối `_result_rows()` đã xử lý LEFT JOIN, NULL và chưa có bài nộp.
- Khối `extract_pdf_blocks()` giữ text, ảnh, kích thước bbox và phân loại figure/formula.
- CSS công thức inline được chuẩn hóa để gần cùng chiều cao với chữ 13.5pt.
- Hình lớn của PDF nằm một hàng riêng và giữ tỉ lệ gốc.


## V5.11.6 — Công thức + chống gian lận

### Công thức Word/PDF
- Chữ thường: Times New Roman 13.5pt.
- Công thức đơn giản: tối thiểu 18px.
- Công thức phân số/căn/số mũ: giữ chiều cao hộp công thức gốc từ Word/PDF,
  thay vì ép tất cả về 21px. Mục tiêu là cỡ chữ/số bên trong công thức
  tương đương chữ thường.
- PDF lưu display_width_px/display_height_px từ bbox thật.
- Word lưu display_height_px từ LibreOffice HTML.

### Chống gian lận
Chỉ tăng `violations` khi học sinh đã vào làm bài, chưa nộp bài,
và rời màn hình/cửa sổ/trang làm bài.

Không tăng vi phạm khi:
- bấm Làm bài;
- bấm Nộp bài;
- redirect sang trang kết quả sau khi nộp;
- copy/cut/paste;
- chuột phải;
- F12;
- thoát fullscreen.

Frontend gộp blur + visibilitychange thành một lần rời màn hình để
tránh cộng hai lỗi cho cùng một hành động.


## V5.11.8 — Hỗ trợ điện thoại / tablet

- Thêm `meta viewport`.
- Giao diện responsive ở màn hình <= 768px.
- Trang đăng nhập tự co vừa điện thoại.
- Input tối thiểu 16px để tránh iPhone tự zoom.
- Radio/checkbox và nút có vùng bấm >= 44px.
- Điều hướng câu chuyển thành hàng ngang cuộn được.
- Hình và công thức không tràn màn hình, giữ tỷ lệ.
- Bảng kết quả cuộn ngang trên màn hình nhỏ.
- Nút Nộp bài có vùng thao tác thuận tiện ở đáy màn hình.
- Không thay đổi logic chấm điểm, random câu, chống gian lận hay Excel.


## V5.11.9 — Công thức + sửa tải Excel kết quả

### Công thức
- Tăng box công thức tối thiểu lên 25px.
- Word formula scale khoảng 1.38x so với metadata HTML gốc.
- PDF formula scale khoảng 1.38x so với bbox 96dpi.
- Mục tiêu: glyph chữ/số/ký hiệu toán gần bằng Times New Roman 13.5pt.

### Excel kết quả
- Sửa lỗi dùng `exam["roster_path"]` không tồn tại trong schema.
- Dùng đúng `instance/roster.xlsx` do `import_roster()` quản lý.
- Tìm hàng tiêu đề linh hoạt.
- Giữ toàn bộ học sinh.
- Không làm bài = 0 điểm + "Không làm bài".
- Có vào nhưng chưa nộp = "Chưa nộp bài".
- Đã nộp = điểm thực tế + "Đã làm bài".

## V5.12.1 — PDF parser ổn định lại

- Word và PDF tiếp tục tách thành hai hàm riêng.
- Sửa lỗi V5.12.0 `Không nhận diện được câu hỏi`: không dùng bbox lan truyền.
- PDF tạo dòng neo từ text thường, sau đó gán span toán vào baseline gần nhất.
- Phân số ở Phần I Câu 1 phương án C giữ tử/mẫu trong cùng phương án.
- PDF mẫu kiểm tra tự động phải nhận 7 câu, đủ 4 phương án ở Phần I/II
  và đúng đáp án A/B/ĐSĐS/SĐSĐ/6/-35/3.

## V5.12.2 — Upload PDF có fallback

- Word và PDF xử lý bằng hai hàm riêng.
- PDF chính dùng baseline-anchor để giữ tử/mẫu phân số cùng dòng.
- Nếu parser PDF chính không tạo được câu hỏi, hệ thống tự thử
  `extract_pdf_blocks_fallback()` trước khi báo lỗi.
- Kiểm tra bằng đúng `ĐỀ KIỂM TRA.pdf`:
  7 câu; 4/4/4/4 phương án; đáp án A/B/ĐSĐS/SĐSĐ/6/-35/3.

## V5.12.3 — Tinh chỉnh riêng cỡ công thức Word/PDF

### Word
- Giảm nhẹ công thức vì bản trước hơi lớn hơn chữ Times New Roman 13.5pt.
- Scale: 0.98.
- Chiều cao tối thiểu: 18px.

### PDF
- Tăng riêng công thức vì glyph toán còn nhỏ hơn chữ đề.
- Scale chính: 1.22.
- Chiều cao tối thiểu: 25px.
- Parser fallback: scale 1.20, tối thiểu 25px.
- Không thay đổi thuật toán baseline/fraction của V5.12.2.

Word và PDF tiếp tục dùng hai hàm xử lý độc lập để dễ chỉnh về sau.


## V5.13.0 — Menu hệ thống + Đăng ký + Admin duyệt thành viên

### Giao diện
- Menu dọc bên phải theo mẫu: Logo, Trang chủ, Giáo viên, Học sinh, Liên hệ.
- Giáo viên/Học sinh có submenu Đăng ký và Đăng nhập.
- Responsive trên điện thoại.

### Đăng ký
- Đăng ký Giáo viên/Học sinh tạo `registration_requests`.
- Tài khoản chưa được đưa vào `users` cho đến khi Admin Đồng ý.
- Admin có thể Đồng ý, Từ chối, Xóa yêu cầu.
- Sau khi duyệt, Admin có thể Xóa thành viên.

### Admin
- Tài khoản mặc định ban đầu:
  - username: `admin`
  - password: `admin123`
- Nên thay mật khẩu mặc định khi triển khai thực tế.
- Admin được đánh dấu `users.is_admin=1`, không thay CHECK role cũ.

### Học sinh từ Excel
- `import_roster()` tiếp tục tạo học sinh từ Mã học sinh.
- Nếu Mã học sinh đã tồn tại, hệ thống reset password về chính Mã học sinh.
- Vì login đã cho phép tìm theo `student_code`, học sinh trong Excel ưu tiên:
  - Tài khoản = Mã học sinh
  - Mật khẩu = Mã học sinh

### Tương thích
Các chức năng V5.12.6 về Word/PDF, chấm điểm, random đề, chống gian lận,
xem kết quả, Excel và xóa đề giáo viên được giữ nguyên.


## V5.14.0 — Menu trái + giới thiệu Admin + tài liệu + tạo đề

- Menu chuyển sang bên trái.
- Trang chủ luôn hiển thị phần giới thiệu website do Admin viết.
- Giáo viên:
  - Bảng điều khiển
  - Tài liệu
  - Tạo đề
    - Tạo đề thủ công
    - Tạo đề bằng AI
- Học sinh:
  - Làm đề kiểm tra
  - Xem tài liệu
- Bỏ mục "Trang giáo viên" khỏi menu.
- Tài liệu giáo viên upload được học sinh xem/tải.
- Tạo đề thủ công dùng lại upload Word/PDF cũ.
- Tạo đề AI có giao diện nhập yêu cầu hoặc ma trận và tạo bản nháp cấu trúc;
  phần sinh nội dung bằng nhà cung cấp AI bên ngoài được tách riêng để không
  ảnh hưởng chức năng V5.12.6.


## V5.15.0 — Ngân hàng đề + Tạo đề thủ công bằng LaTeX

### Ngân hàng đề
- Menu Giáo viên có mục thứ 4: Tạo ngân hàng đề.
- Giáo viên upload DOCX/PDF.
- Parser Word/PDF cũ được dùng nguyên để tách câu.
- Câu được lưu vào `bank_sources`, `bank_questions`, `bank_options`.
- Giáo viên xem đầy đủ câu, phương án, đáp án và chọn câu cần kiểm tra.
- Các câu đã chọn được COPY sang schema `exams/questions/options` cũ.
- Sau đó giáo viên kiểm tra bằng màn hình cũ và bấm Xuất bản.
- Học sinh làm bằng đúng luồng V5.12.6.

### Tạo đề thủ công
- Giáo viên gõ nội dung trực tiếp.
- Công thức LaTeX:
  - inline: `$x^2+1$`
  - block: `$$\frac{x+1}{x-1}$$`
- Hỗ trợ đủ 4 phần.
- Tạo exam nháp theo schema cũ, sau đó kiểm tra và Xuất bản.
- MathJax chỉ bổ sung cho câu LaTeX; Word/PDF không bị đổi parser.


## V5.15.1 — Fix init_db + chức năng nằm trên trang, không nằm trong menu

### Fix SQLite
Sửa lỗi:
`sqlite3.OperationalError: near "'Giáo viên có thể quản lý...": syntax error`

Nguyên nhân: SQLite không nối nhiều string literal liền nhau như Python.
`homepage_intro` nay được INSERT bằng query tham số.

### Menu
Menu bên trái chỉ gồm:
- Logo
- Trang chủ
- Giáo viên
- Học sinh
- Liên hệ

### Trang Giáo viên
Sau đăng nhập hiển thị 4 hộp trên trang:
1. Bảng điều khiển
2. Tài liệu
3. Tạo đề
4. Tạo ngân hàng đề

### Trang Học sinh
Sau đăng nhập hiển thị 2 hộp:
1. Làm đề kiểm tra
2. Xem tài liệu

Khi chưa đăng nhập, trang Giáo viên/Học sinh hiển thị Đăng ký và Đăng nhập.

Các hàm lõi V5.12.6 về Word/PDF, random đề, chấm điểm, chống gian lận,
Excel, học sinh làm bài và xóa đề không bị sửa.

## V5.16.0
- Teacher/student login -> đúng portal.
- Excel students -> auto account, no admin approval.
- Question bank -> bank_* schema guard for old DB.
- Admin -> explicit approval column.
- AI -> visible draft questions after prompt.
- V5.12.6 core parser/grading/random/anti-cheat/delete-exam unchanged.


## V5.16.1 — Fix menu Giáo viên khi đang đăng nhập Admin

- Admin bấm menu Giáo viên không còn thấy nút mở Quản trị hệ thống.
- Giáo viên thường đã đăng nhập: hiện 4 hộp chức năng giáo viên.
- Admin hoặc khách chưa đăng nhập: hiện Đăng ký Giáo viên / Đăng nhập Giáo viên.
- Menu Học sinh cũng không dùng session Admin như tài khoản học sinh.
- Không thay đổi code lõi V5.12.6.


## V5.16.2 — AI đúng yêu cầu + editor + bỏ số hộp

### AI
- Chỉ tạo đúng số câu và đúng phần được ghi trong prompt.
- Ví dụ:
  - `Tạo 3 câu Phần I về đạo hàm`
  - `Tạo 2 câu Phần II về xác suất`
  - `Tạo 1 câu Phần III về tích phân`
- Không tự tạo câu ở phần khác.
- Có thể:
  - kiểm tra;
  - xóa câu;
  - tạo thêm câu;
  - tạo exam nháp rồi kiểm tra bằng màn hình cũ;
  - sau đó bấm Xuất bản như V5.12.6.

### Logo
File logo:
`static/site_logo.jpg`

Có thể thay bằng ảnh khác nhưng nên giữ tên `site_logo.jpg`.
Kích thước hiển thị nằm trong:
`static/style.css`
ở selector `.sidebar-logo`.

### Nội dung trang Giáo viên/Học sinh
- Giáo viên: `templates/teacher_portal.html`
- Học sinh: `templates/student_portal.html`

Muốn đổi tiêu đề/nội dung hộp, sửa trực tiếp text trong các file này.

Các số 01/02/03/04 đã được bỏ khỏi giao diện.

## V5.16.3
- AI tạo đúng số câu theo từng Phần I/II/III/IV giáo viên yêu cầu.
- Công thức hiện trực tiếp bằng MathJax, không hiện $...$ ở giao diện chính.
- Không có nút Kiểm tra.
- Có Xóa câu, Tạo thêm câu, Xuất bản cho học sinh.
- Phần I: 4 lựa chọn; Phần II: 4 ý; Phần III: trả lời ngắn; Phần IV: tự luận.
- Tài liệu có nút Xem và Tải riêng.
- Không sửa parser Word/PDF, random, chấm điểm, anti-cheat lõi V5.12.6.

## V5.16.4
- Sửa bộ đọc số câu AI theo từng đoạn Phần I/II/III/IV.
- Mỗi phần hiển thị tiêu đề một lần; bên dưới reset Câu 1, Câu 2, ...
- Xem tài liệu:
  - PDF/ảnh/text xem trực tiếp.
  - Word/Excel/PowerPoint tự thử chuyển sang PDF bằng LibreOffice để xem.
  - Nếu không chuyển được vẫn có nút Tải xuống và thông báo nguyên nhân.
- Không sửa parser Word/PDF, random, chấm điểm, anti-cheat và các hàm lõi V5.12.6.


## V5.17.0 — Tin tức Admin + Đảo đề

### Tin tức
Admin có thể đăng:
- tiêu đề;
- nội dung;
- ảnh;
- video.

Tin xuất hiện trên Trang chủ. Admin có nút xóa tin.

### Đảo đề
Trang Giáo viên có hộp `Đảo đề (Trộn đề)`:
- upload DOCX/PDF;
- Kiểm tra;
- Đảo;
- Xóa;
- In.

Quy tắc:
- Phần I: đảo câu + đảo phương án A/B/C/D.
- Phần II: đảo câu, không đảo 4 ý.
- Phần III: đảo câu.
- Phần IV: đảo câu.
- Đáp án Phần I được remap theo vị trí mới của phương án đúng.
- Đáp án Phần II/III đi cùng question_id khi câu được đổi vị trí.
- Hình và công thức nằm trong block của câu/phương án nên đi cùng nội dung.

Các bảng shuffle_* và news_posts là bảng mới, không thay đổi schema/lõi V5.12.6.


## V5.20.4 — Fix BuildError khi xem đề

Nguyên nhân:
- `shuffle_exam_check()` truyền `version_id=None`.
- Template vẫn gọi endpoint `shuffle_exam_docx` cần cả `source_id` và `version_id`.
- Flask bỏ giá trị None và sinh BuildError.

Đã sửa:
- Kiểm tra đề trước khi đảo: nút `Tải file gốc`.
- Xem mã đề đã trộn: nút `Xuất Word` với đầy đủ `source_id + version_id`.
- Thêm route `shuffle_exam_source_download`.
- Không thay đổi parser Word/PDF, MathType preserve, random, đáp án, chấm điểm.

## V5.20.6 — MathType Full 12pt -> 13.5pt

Theo ảnh cấu hình MathType được cung cấp:
- Full 12pt
- Script 7pt
- Sub-sub 5pt
- Symbol 18pt
- Sub-symbol 12pt

Word export dùng hệ số 1.125 (=13.5/12) cho khung hiển thị OLE MathType.
Chữ thường dùng Times New Roman 13.5pt.
MathType vẫn là OLE editable; hình vẽ thường không bị scale.

## V5.20.7 — Preserve Original Word/MathType

Thay đổi theo nguyên tắc bảo toàn file nguồn:
- Upload Word dùng FileStorage.save() để lưu file nguồn nguyên bản.
- Tải file gốc trả đúng file đã lưu, không dựng lại.
- Khi xuất đề đã đảo:
  - không ép lại font/cỡ trong thân câu Word gốc;
  - không scale MathType;
  - không chuyển OLE MathType thành ảnh/LaTeX;
  - clone nguyên OOXML của câu và phương án;
  - chỉ đổi số câu và nhãn A/B/C/D khi cần.
- Header/mã đề/bảng đáp án do hệ thống tạo mới vẫn dùng Times New Roman 13.5pt.
- Nếu không ánh xạ được câu về Word gốc, hệ thống dừng xuất thay vì fallback làm mất MathType.

## V5.20.8 — MathJax Web Sync

- Chữ câu hỏi web: Times New Roman 13.5pt.
- MathJax SVG: `mjx-container { font-size:105% !important; }`.
- LaTeX render vector, không vỡ khi zoom.
- Bỏ CSS mobile ép `height:auto!important` lên formula image vì rule đó
  làm mất `display_height_px` đã tính riêng cho từng công thức Word/PDF.
- Formula image giữ kích thước parser; figure giữ tỷ lệ và tách hàng.
- Word download vẫn giữ MathType/OLE nguyên bản theo V5.20.7.

## V5.20.9 — Ngân hàng đề: đủ nút cho 4 bước

Luồng:
1. Chọn câu.
2. Tạo bản nháp.
3. Kiểm tra bản nháp.
4. Xuất bản cho học sinh.

Trang chọn câu:
- Bước 1 có nút cuộn tới danh sách câu.
- Bước 2 có nút tạo bản nháp.
- Bước 3 và 4 hiển thị trạng thái chờ.

Sau khi tạo bản nháp:
- Bước 1 và 2 hiện hoàn thành.
- Bước 3 có nút `Kiểm tra bản nháp`.
- Bước 4 có nút `Xuất bản cho học sinh`.

Không thay đổi logic lõi V5.12.6.

## V5.21.0 — Mammoth + MathJax hybrid display

Chỉ thay lớp hiển thị:
- thêm Mammoth preview cho file DOCX nguồn;
- MathJax dùng CommonHTML (`tex-mml-chtml.js`);
- MathJax hiển thị 105% so với chữ xung quanh;
- formula image Word/PDF giữ `display_height_px`, không ép một chiều cao chung;
- figure giữ tỷ lệ và tách hàng.

Không thay:
- extract_word_blocks / extract_pdf_blocks;
- tách câu, trộn đề, chấm điểm;
- ngân hàng đề;
- Word/OLE MathType download preserve;
- student exam workflow.

Cài thêm:
`pip install mammoth`

## V5.21.1 — Bank workflow thật + icon menu

Ngân hàng đề:
1. Chọn câu.
2. Tạo bản nháp.
3. Bấm `Kiểm tra`: hệ thống kiểm tra câu, đáp án, điểm.
4. Chỉ khi Bước 3 đạt, nút `Xuất bản` mới được bật.

Nếu giáo viên thay đổi đáp án hoặc phương án chấm sau khi kiểm tra,
trạng thái kiểm tra bị hủy và phải bấm Kiểm tra lại.

Portal:
- Giáo viên: Upload đề, Tài liệu, Tạo đề, Ngân hàng, Đảo đề.
- Học sinh: Làm bài, Tài liệu.
- Dùng Bootstrap Icons, icon trên + chữ ngắn dưới, card nhỏ gọn.

## V5.21.2 — Fix Bước 3 và Bước 4 Ngân hàng đề

Nguyên nhân V5.21.1:
- trạng thái Bước 3 lưu trong session;
- validator quá nghiêm nên giáo viên có thể bấm Kiểm tra nhưng Bước 4
  không bao giờ được mở;
- nút Bước 4 bị disabled hoàn toàn.

V5.21.2:
- thêm `exams.draft_checked` trong SQLite;
- Bước 3 luôn thực hiện được nếu đề có câu;
- các lỗi đáp án/điểm hiển thị thành cảnh báo để giáo viên xem;
- sau Bước 3, trạng thái lưu DB;
- Bước 4 là nút thật, không còn HTML disabled;
- Bước 4 cập nhật `published=1`;
- sau xuất bản quay lại bản nháp và hiển thị `Đã xuất bản`.

## V5.21.3 — Logo mới
- Thay logo menu trái bằng hình người dùng cung cấp.
- Chỉ lấy vùng logo tròn, không lấy các nút menu trong ảnh tham chiếu.
- Không thay đổi route, parser, ngân hàng đề, trộn đề hoặc chấm điểm.

## V5.21.4 — Fix triệt để Bước 3 / Bước 4

Thay đổi:
- Bước 3 và Bước 4 dùng link route thật thay vì form submit.
- Hai route nhận cả GET và POST.
- Mỗi route tự kiểm tra/migrate cột `draft_checked`.
- Bước 3 ghi `draft_checked=1` trực tiếp cùng transaction.
- Bước 4 ghi `published=1` trực tiếp.
- Trang hiển thị trạng thái:
  - Bản nháp đã tạo
  - Kiểm tra: Đã/Chưa
  - Xuất bản: Đã/Chưa
- Giữ nguyên logo mới V5.21.3 và toàn bộ parser/chấm/trộn/MathType.

## V5.21.5 — Rebuild quy trình Ngân hàng đề

Thay vì tiếp tục gắn Bước 3/4 vào `edit_exam()` chung, bản này tạo workflow riêng:

- `/teacher/question-bank/exam/<exam_id>/workflow`
- `/teacher/question-bank/exam/<exam_id>/check`
- `/teacher/question-bank/exam/<exam_id>/publish`
- `/teacher/question-bank/exam/<exam_id>/unpublish`

Luồng:
1. Chọn câu.
2. Tạo bản nháp.
3. Xem câu hỏi + bấm Xác nhận đã kiểm tra -> `draft_checked=1`.
4. Bấm Xuất bản -> `published=1`.

Quyền sở hữu không dùng `teacher_id` trong exams.
Hệ thống đọc `source_file="question_bank:<source_id>"` và đối chiếu
`bank_sources.uploaded_by`.

Logo được khôi phục đúng `site_logo.jpg` + class `sidebar-logo`
từ bản V5.21.0.

## V5.21.6 — Ngân hàng đề rút còn 3 bước

1. Chọn câu.
2. Tạo đề từ câu đã chọn:
   - hiển thị ngay toàn bộ đề;
   - giáo viên xem câu/đáp án;
   - có nút Hủy và quay lại chọn câu mới.
3. Xuất bản:
   - cập nhật `exams.published=1`;
   - học sinh thấy đề theo logic cũ.

Bỏ bước `Kiểm tra` khỏi workflow Ngân hàng đề.
Không thay parser Word/PDF, MathType, MathJax, chấm điểm, trộn đề,
menu, login hoặc các luồng đề upload thông thường.

## V5.21.7 — Ngân hàng đề đúng 3 bước

Lỗi của V5.21.6:
- Backend đã có preview/cancel/publish.
- Nhưng `question_bank_detail.html` vẫn còn giao diện 4 bước cũ.
- Vì vậy trên web nhìn gần như không thay đổi.

V5.21.7:
1. Viết lại hoàn toàn trang chọn câu chỉ còn 3 bước.
2. Bước 2 tạo đề và redirect sang trang preview.
3. Preview hiển thị toàn bộ câu đã chọn.
4. Có nút Hủy để quay lại chọn bộ câu mới.
5. Có nút Bước 3 — Xuất bản cho học sinh.

## V5.21.8 — Fix nguyên nhân thật khiến Bước 2 không chạy

Nguyên nhân tìm thấy trong V5.21.7:
`build_exam_from_bank()` vẫn có điều kiện:

    điểm Phần II == số câu Phần II được chọn

Nếu giáo viên chọn số câu Phần II khác mặc định 2 câu,
route lập tức quay về Bước 1. Vì vậy nhìn giống như Bước 2 không hoạt động.

V5.21.8:
- bỏ điều kiện chặn trên;
- tự tính Phần II = số câu Phần II (mỗi câu tối đa 1 điểm);
- cân các phần còn lại để tổng điểm phù hợp;
- dữ liệu form sai số không gây lỗi 500;
- Step 2 tạo xong luôn redirect sang preview;
- Step 3 nhận cả GET/POST và update published=1 trực tiếp;
- không dùng draft_checked trong luồng Ngân hàng.
# Bổ sung V5.22.0 — công thức toán trên màn hình thi

- Khi upload DOCX, công thức MathType dạng OLE được đọc bằng MTEF-py và lưu
  thành block LaTeX; công thức Word Equation (OMML) cũng được đổi sang LaTeX.
- MathJax hiển thị công thức dạng vector trên màn hình giáo viên và học sinh.
- Nếu gặp một công thức chưa được bộ chuyển đổi hỗ trợ, hệ thống tự giữ ảnh
  công thức cũ thay vì làm mất nội dung.
- PDF không chứa dữ liệu MathType/OMML gốc nên tiếp tục dùng bộ xử lý PDF hiện
  tại; thay đổi này không tác động đến tách câu, đáp án, trộn đề hay chấm điểm.

## V5.24.17 — tải Word/PDF ổn định trên Windows

- Các route Đảo đề không còn trả trực tiếp `send_file(path)` với file vật lý.
- File được tạo ở đường dẫn UUID trong thư mục tạm, sau đó đọc toàn bộ vào RAM (`BytesIO`) rồi mới gửi về trình duyệt.
- Có retry khi Windows/Word/MathType/antivirus còn giữ handle ngắn hạn.
- PDF LaTeX cũng dùng UUID trong thư mục tạm, không dùng tên file cố định trong `exports`.
- Không cần xóa đề cũ chỉ vì file export lần trước đang mở.


## V5.24.18 — sửa hình Word đúng crop + PDF → Word MathType bulk

- DOCX: ngoài `wp:extent`, parser đọc cả `a:srcRect` của DrawingML. Ảnh preview LibreOffice được crop đúng như Word rồi mới hiển thị ở kích thước gốc. Điều này sửa trực tiếp các đồ thị Phần I Câu 1, 3, 4 bị nhìn nhỏ.
- Kiểm tra file mẫu `Đề số 1 OTC-1-L12(4).docx`: Câu 1 có srcRect 29.412%/19.195%/26.706%/37.288%; Câu 3 28.672%/18.260%/31.664%/6.925%; Câu 4 29.331%/16.847%/33.335%/21.786%.
- PDF → Word MathType: không gọi `Toggle TeX` hơn một nghìn lần nữa. Word chọn toàn bộ vùng tài liệu và gọi MathType Toggle TeX một lần (bulk region).
- Các đoạn TeX không có toán tử được thêm `\vphantom{+}` chỉ trong bản trung gian để vượt giới hạn Toggle TeX của MathType; ký hiệu này không hiển thị nên không thay đổi công thức nhìn thấy.
- Không dùng pywin32; vẫn dùng Windows Script Host + Word COM + MathType 7.
- Sau khi lưu, DOCX vẫn được kiểm tra `Equation.DSMT4` và TeX marker trước khi trả file.

## V5.24.21 — PDF đảo đề xuất Word LaTeX, không tự gọi MathType
- Nguồn DOCX: tiếp tục clone OOXML để giữ MathType OLE gốc.
- Nguồn PDF: Word tải về chứa LaTeX chuẩn có dấu `$...$` hoặc `\[...\]`, giống tiện ích PDF sang Word.
- Không tự động gọi Word/MathType/Toggle TeX cho nguồn PDF; loại bỏ lỗi batch, “Too many windows are open”, và “MathType không tiến thêm”.
- File từng mã đề: `De_<mã>_LaTeX.docx`; tất cả mã đề: `Tat_ca_ma_de_LaTeX.docx`.
- Tiêu đề, lời giải, hình và bảng đáp án cuối file vẫn được giữ trong luồng xuất Word.

## V5.25.1 — sửa hình Word và kiểm tra trước khi đưa lên VPS

V5.25.1 chỉ sửa ảnh minh họa trong bản xem trước Word trên website; không thay đổi
quy tắc đảo đề, phân quyền giáo viên/học sinh, MathType/OLE của file Word xuất.
Với DOCX OTC-1 lớp 12, hình Câu 1, 3, 4 Phần I được LibreOffice xuất thành
GIF có phần trong suốt rất rộng. Bản cũ áp `a:srcRect` một lần nữa lên ảnh GIF,
trong khi phép crop đã được LibreOffice xử lý. Bản mới chỉ cắt lề trong suốt
trên ảnh GIF preview, rồi dùng kích thước OOXML để hiển thị, không can thiệp
vào file DOCX gốc. Mỗi phiên bản đề được tải lên trước khi cập nhật vẫn mang
URL ảnh cũ; cần tải lại DOCX mẫu trong mục Đảo đề để sinh preview mới.

**Để kiểm thử trên máy có đầy đủ dependencies:**

```bash
python -m pip install -r requirements.txt
python -m pytest tests/test_word_figure_preview.py tests/test_teacher_spaces.py -q
```

**VPS:** Dùng Dockerfile/docker-compose.yml sẵn có; tạo `.env` từ
`.env.example`, đặt `APP_ENV=production`, `SECRET_KEY` ngẫu nhiên dài,
`ADMIN_INITIAL_PASSWORD` riêng và `ENABLE_DEMO_ACCOUNTS=0`. Bảo đảm thư mục
`APP_DATA_DIR=/app/storage` nằm trên volume bền vững (gồm database,
`uploads/question_media` và thư mục tài liệu), có HTTPS ở reverse proxy.
Giữ `WEB_CONCURRENCY=1` nếu vẫn dùng SQLite và LibreOffice nhiều tài liệu;
trước khi cập nhật phải sao lưu volume. Chạy kiểm tra upload → xem trước
Câu 1/3/4 → đảo một mã → xem bằng tài khoản giáo viên và học sinh; kiểm tra
quyền truy cập chéo giữa hai giáo viên.

**Giới hạn kiểm thử:** Mình đã chạy trực tiếp LibreOffice với file Word
mẫu và hàm xử lý ba ảnh. Chưa kiểm thử hoàn chỉnh qua HTTP trên VPS/Windows
của bạn nên cần hoàn thành bước smoke-test trên VPS trước khi cho học sinh vào.
