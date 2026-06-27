// https://console.groq.com/keys
const GROQ_API_KEY = "";
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
