# Dùng AI để tạo bảng tra cứu từ vựng nhanh thì vì gõ hoặc chép tay

<img width="2544" height="309" alt="image" src="https://github.com/user-attachments/assets/b2e16b37-d04e-4436-acdc-8755b43a2d07" />

> [!TIP]
> Bạn chỉ cần paste từ vào ô đầu tiên, còn lại để api và AI lo
  

# Hướng dẫn cài đặt

## Yêu cầu
- Tài khoản Google
- Tài khoản Groq (free): [console.groq.com](https://console.groq.com)

---

## Bước 1 - Tạo Google Sheet

1. Vào [sheets.google.com](https://sheets.google.com) → tạo sheet mới
2. Đặt tên sheet tab ở dưới là **`Trang tính1`** (mặc định đã đúng nếu dùng tiếng Việt)
3. Điền header ở **hàng 1** như sau:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| Từ vựng | Loại từ | Phát âm Anh Mỹ | Nghĩa tiếng Việt | Đồng nghĩa | Trái nghĩa | Ví dụ |
<img width="2536" height="1242" alt="image" src="https://github.com/user-attachments/assets/0c43028c-6975-4b53-b08e-5f02854dcc60" />


---

## Bước 2 - Lấy Groq API Key
> [!NOTE]
> Lý do mình dùng mô hình AI bên dịch vụ này vì nó nhả token nhanh nhất, thật ra cho 1 câu ví dụ thì đa số các top model hiện nay nhanh gần như nhau nhưng con này free, nhanh và số lượt dùng mỗi ngày nhiều  
> Các bạn dùng api của con AI nào cũng được, có thể nhờ AI như claude, gemini, chatgpt, ... viết lại code lấy đúng theo api các bạn dùng là được

1. Vào [console.groq.com](https://console.groq.com) → đăng nhập bằng Google
2. Xong vào [console.groq.com/keys](https://console.groq.com/keys)
3. Nhấn **Create API Key** → đặt tên tùy ý → nhấn **Submit**
<div align="center">
 <img width="1802" height="176" alt="image" src="https://github.com/user-attachments/assets/15a389be-8dbb-45d9-853a-ff9b83533b0f" />
<img width="1422" height="773" alt="image" src="https://github.com/user-attachments/assets/55c989f3-ce2c-4fb5-85a0-ebfb522dc687" />
</div>

4. **Copy key ngay** (bắt đầu bằng `gsk_...`)
<img width="1587" height="632" alt="image" src="https://github.com/user-attachments/assets/841a1082-0fc1-4c5c-b829-12f6119f6f81" />

---

## Bước 3 - Cài đặt Script
<img width="2536" height="1242" alt="image" src="https://github.com/user-attachments/assets/567ba5d2-9e5b-4107-9238-c0e7469f760d" />

1. Trong Google Sheet, vào menu **Extensions → Apps Script**
2. Xóa toàn bộ code mặc định trong editor
3. Paste đoạn script sau vào:

<details>
  <summary>Script (nhấn vào đây để mở ra)</summary>

```javascript
// https://console.groq.com/keys
const GROQ_API_KEY = "";  // thay key vào đây
const SHEET_NAME = "Trang tính1";

function tuDongGoiAI(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  if (sheet.getName() !== SHEET_NAME) return;
  if (range.getColumn() !== 1 || range.getRow() <= 1) return;

  var row = range.getRow();
  var word = String(range.getValue()).trim();

  if (word === "") {
    sheet.getRange(row, 2, 1, 6).clearContent();
    return;
  }

  sheet.getRange(row, 2).setValue("Đang tra cứu...");

  var responses = UrlFetchApp.fetchAll([
    {
      url: "https://api.dictionaryapi.dev/api/v2/entries/en/" + encodeURIComponent(word),
      method: "get",
      muteHttpExceptions: true
    },
    {
      url: "https://api.groq.com/openai/v1/chat/completions",
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + GROQ_API_KEY },
      payload: JSON.stringify({
        // model nhanh nhất, nếu bạn tìm đuợc model nào nhanh hơn thì có thê thay
        model: "llama-3.1-8b-instant",
        temperature: 0,
        max_tokens: 120,
        response_format: { type: "json_object" },
        messages: [
          {
            role: "system",
            // prompt để AI viết câu ví dụ
            content: 'You are an English dictionary. Always reply with a JSON object with exactly these keys: "partOfSpeech", "phonetic", "vietnamese", "synonyms", "antonyms", "example". Values are strings. synonyms and antonyms are comma-separated, max 3 each. example is one short sentence. No markdown, no extra keys.'
          },
          {
            role: "user",
            content: 'Analyze the word: "' + word + '"'
          }
        ]
      }),
      muteHttpExceptions: true
    }
  ]);

  var dictRes = responses[0];
  var groqRes = responses[1];
  var dictOk = dictRes.getResponseCode() === 200;
  var ai = {};
  if (groqRes.getResponseCode() === 200) {
    try {
      var groqJson = JSON.parse(groqRes.getContentText());
      ai = JSON.parse(groqJson.choices[0].message.content);
    } catch (_) {}
  }
  
  var partOfSpeech, phonetic, vietnamese, synonyms, antonyms, example;
  example = (ai.example || "").replace(/^["'`]|["'`]$/g, "").trim();

  if (dictOk) {
    try {
      var data = JSON.parse(dictRes.getContentText())[0];

      partOfSpeech = data.meanings && data.meanings.length
        ? data.meanings[0].partOfSpeech : (ai.partOfSpeech || "");

      phonetic = data.phonetic || "";
      if (!phonetic && data.phonetics) {
        for (var i = 0; i < data.phonetics.length; i++) {
          if (data.phonetics[i].text) { phonetic = data.phonetics[i].text; break; }
        }
      }

      if (!phonetic) phonetic = ai.phonetic || "";

      var syns = [], ants = [];
      data.meanings.forEach(function(m) {
        if (m.synonyms) syns = syns.concat(m.synonyms);
        if (m.antonyms) ants = ants.concat(m.antonyms);
        (m.definitions || []).forEach(function(d) {
          if (d.synonyms) syns = syns.concat(d.synonyms);
          if (d.antonyms) ants = ants.concat(d.antonyms);
        });
      });

      synonyms = syns.length
        ? [...new Set(syns)].slice(0, 3).join(", ")
        : (ai.synonyms || "");

      antonyms = ants.length
        ? [...new Set(ants)].slice(0, 3).join(", ")
        : (ai.antonyms || "");

    } catch (_) {
      // nếu từ bạn viết không có trong dictionary, fallback với AI
      partOfSpeech = ai.partOfSpeech || "";
      phonetic     = ai.phonetic    || "";
      synonyms     = ai.synonyms    || "";
      antonyms     = ai.antonyms    || "";
    }

    try { vietnamese = LanguageApp.translate(word, "en", "vi"); } catch (_) {}
    if (!vietnamese) vietnamese = ai.vietnamese || "";

  } else {
    partOfSpeech = ai.partOfSpeech || "";
    phonetic     = ai.phonetic     || "";
    vietnamese   = ai.vietnamese   || "";
    synonyms     = ai.synonyms     || "";
    antonyms     = ai.antonyms     || "";
  }

  sheet.getRange(row, 2, 1, 6).setValues([[
    partOfSpeech,
    phonetic,
    vietnamese,
    synonyms,
    antonyms,
    example
  ]]);
}
```

</details>

> hoặc là copy từ file [./script.js](script.js)

4. Thay `gsk_...` ở dòng đầu bằng key vừa copy ở Bước 2
5. Nhấn **Save** (Ctrl+S)

<img width="2557" height="1395" alt="image" src="https://github.com/user-attachments/assets/8984c41a-628d-4bf7-bfde-bf103cc3cab0" />


---

## Bước 4 - Gắn Trigger
<img width="862" height="888" alt="image" src="https://github.com/user-attachments/assets/5c54b2e9-f5a7-42fc-b1ff-f4760efba7d1" />

1. Trong Apps Script, nhấn biểu tượng **⏰ Triggers** ở sidebar trái (hoặc menu **Triggers**)
2. Nhấn **+ Add Trigger** góc dưới phải
3. Cấu hình như sau:

| Trường | Giá trị |
|---|---|
| Choose which function to run | `tuDongGoiAI` |
| Choose which deployment to run | `Head` |
| Select event source | `From spreadsheet` |
| Select event type | `On edit` |

4. Nhấn **Save**
5. Google sẽ hỏi cấp quyền → chọn tài khoản Google của bạn → nhấn **Allow**

<img width="1229" height="1179" alt="image" src="https://github.com/user-attachments/assets/7f6dfa43-1160-40b3-9b8c-943111c0692f" />

<div align="center"><i>Như hình là được</i></div>

---

## Bước 5 - Dùng thử

Quay lại Google Sheet, gõ một từ tiếng Anh vào **cột A** (từ hàng 2 trở xuống) rồi nhấn Enter - các ô B đến G sẽ tự động điền sau 2-3 giây.

**Ví dụ:**
- `happy` → noun/adjective, IPA, nghĩa TV, đồng/trái nghĩa, câu ví dụ
- `variants` → AI tự phân tích vì không có trong từ điển chuẩn

---

## Lưu ý

- **Chỉ nhập 1 từ mỗi ô** - cụm từ như "index variant" sẽ cho kết quả không chính xác
- **Xóa từ** trong cột A → các ô B–G tự động xóa theo
- Groq free tier cho **14,400 request/ngày** - đủ dùng thoải mái cho việc học vocab
