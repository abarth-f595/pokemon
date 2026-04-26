// ==========================================
// 1. DOM Elements & Globals
// ==========================================
let contacts = [];
let snippets = [];
let historyList = [];
let currentOriginalText = "";
let currentGeneratedText = "";

// Settings Modal
const settingsBtn       = document.getElementById('settingsBtn');
const closeSettingsBtn  = document.getElementById('closeSettingsBtn');
const settingsModal     = document.getElementById('settingsModal');
const saveSettingsBtn   = document.getElementById('saveSettingsBtn');
const customRulesInput  = document.getElementById('customRules');

// Address Book Modal
const addressBookBtn   = document.getElementById('addressBookBtn');
const closeAddressBtn  = document.getElementById('closeAddressBtn');
const addressBookModal = document.getElementById('addressBookModal');
const contactListEl    = document.getElementById('contactList');
const contactForm      = document.getElementById('contactForm');
const clearContactFormBtn = document.getElementById('clearContactFormBtn');

// Snippets Modal
const snippetsBtn          = document.getElementById('snippetsBtn');
const closeSnippetsBtn     = document.getElementById('closeSnippetsBtn');
const snippetsModal        = document.getElementById('snippetsModal');
const snippetListEl        = document.getElementById('snippetList');
const snippetForm          = document.getElementById('snippetForm');
const clearSnippetFormBtn  = document.getElementById('clearSnippetFormBtn');
const insertSnippetBtn     = document.getElementById('insertSnippetBtn');
const snippetQuickDropdown = document.getElementById('snippetQuickDropdown');
const quickSnippetList     = document.getElementById('quickSnippetList');

// History Modal
const historyBtn     = document.getElementById('historyBtn');
const closeHistoryBtn = document.getElementById('closeHistoryBtn');
const historyModal   = document.getElementById('historyModal');
const historyListEl  = document.getElementById('historyList');

// Main UI
const targetContactSelect = document.getElementById('targetContact');
const inputText      = document.getElementById('inputText');
const outputText     = document.getElementById('outputText');
const diffOutputText = document.getElementById('diffOutputText');
const diffToggle     = document.getElementById('diffToggle');
const clearBtn       = document.getElementById('clearBtn');
const copyBtn        = document.getElementById('copyBtn');
const downloadTxtBtn = document.getElementById('downloadTxtBtn');
const downloadWordBtn = document.getElementById('downloadWordBtn');
const generateBtn    = document.getElementById('generateBtn');
const docTypeSelect  = document.getElementById('docType');
const toneSelect     = document.getElementById('tone');
const actionTypeSelect = document.getElementById('actionType');
const toast          = document.getElementById('toast');
const toastMsg       = document.getElementById('toastMsg');
const toastIcon      = toast.querySelector('i');
const loadingOverlay = document.getElementById('loadingOverlay');

// ==========================================
// 2. Utils & Modals
// ==========================================
function closeAllModals() {
    settingsModal.classList.add('hidden');
    addressBookModal.classList.add('hidden');
    snippetsModal.classList.add('hidden');
    historyModal.classList.add('hidden');
    snippetQuickDropdown.classList.add('hidden');
}
window.addEventListener('click', (e) => {
    if (e.target.classList.contains('modal-backdrop')) e.target.classList.add('hidden');
    if (!insertSnippetBtn.contains(e.target) && !snippetQuickDropdown.contains(e.target)) {
        snippetQuickDropdown.classList.add('hidden');
    }
});
let toastTimeout;
function showToast(message, isError = false) {
    toastMsg.textContent = message;
    toast.className = isError ? 'toast error' : 'toast';
    toastIcon.className = isError ? 'ph ph-warning-circle' : 'ph ph-check-circle';
    toast.classList.remove('hidden');
    clearTimeout(toastTimeout);
    toastTimeout = setTimeout(() => toast.classList.add('hidden'), 3500);
}

// ==========================================
// 3. Settings（APIキー不要・カスタムルールのみ）
// ==========================================
function loadSettings() {
    const rules = localStorage.getItem('bizRevise_rules') || '';
    customRulesInput.value = rules;
}
saveSettingsBtn.addEventListener('click', () => {
    localStorage.setItem('bizRevise_rules', customRulesInput.value);
    settingsModal.classList.add('hidden');
    showToast('設定を保存しました');
});
settingsBtn.addEventListener('click', () => { loadSettings(); settingsModal.classList.remove('hidden'); });
closeSettingsBtn.addEventListener('click', () => settingsModal.classList.add('hidden'));

// ==========================================
// 4. Contact / Address Book
// ==========================================
function loadContacts() {
    const stored = localStorage.getItem('bizRevise_contacts');
    contacts = stored ? JSON.parse(stored) : [];
    renderContactList(); renderContactDropdown();
}
function saveContacts() {
    localStorage.setItem('bizRevise_contacts', JSON.stringify(contacts));
    renderContactList(); renderContactDropdown();
}
function formatContactName(c) {
    const parts = [];
    if (c.company) parts.push(c.company);
    if (c.role)    parts.push(c.role);
    if (c.name)    parts.push(c.name + ' 様');
    return parts.join(' ');
}
function renderContactDropdown() {
    targetContactSelect.innerHTML = '<option value="">-- 指定なし --</option>';
    contacts.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c.id; opt.textContent = formatContactName(c);
        targetContactSelect.appendChild(opt);
    });
}
function renderContactList() {
    contactListEl.innerHTML = '';
    if (contacts.length === 0) {
        contactListEl.innerHTML = '<li class="contact-item"><span style="color:#94a3b8">登録がありません</span></li>';
        return;
    }
    contacts.forEach(c => {
        const li = document.createElement('li'); li.className = 'contact-item';
        li.innerHTML = `
            <div class="contact-info"><strong>${c.name}</strong><span>${c.company||''} ${c.role||''}</span></div>
            <div class="contact-actions">
                <button class="icon-btn edit-btn" onclick="editContact(${c.id})"><i class="ph ph-pencil-simple"></i></button>
                <button class="icon-btn btn-danger" onclick="deleteContact(${c.id})"><i class="ph ph-trash"></i></button>
            </div>`;
        contactListEl.appendChild(li);
    });
}
contactForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('contactId').value;
    const contact = {
        id: id ? parseInt(id) : Date.now(),
        company: document.getElementById('contactCompany').value.trim(),
        role:    document.getElementById('contactRole').value.trim(),
        name:    document.getElementById('contactName').value.trim(),
        email:   document.getElementById('contactEmail').value.trim()
    };
    if (id) { const idx = contacts.findIndex(x => x.id == id); if (idx > -1) contacts[idx] = contact; }
    else contacts.push(contact);
    saveContacts(); clearContactForm(); showToast('保存しました');
});
window.editContact = (id) => {
    const c = contacts.find(x => x.id == id); if (!c) return;
    document.getElementById('contactId').value    = c.id;
    document.getElementById('contactCompany').value = c.company;
    document.getElementById('contactRole').value  = c.role;
    document.getElementById('contactName').value  = c.name;
    document.getElementById('contactEmail').value = c.email;
};
window.deleteContact = (id) => {
    if (confirm('削除しますか？')) { contacts = contacts.filter(x => x.id != id); saveContacts(); }
};
function clearContactForm() { document.getElementById('contactId').value = ''; contactForm.reset(); }
clearContactFormBtn.addEventListener('click', clearContactForm);
addressBookBtn.addEventListener('click', () => addressBookModal.classList.remove('hidden'));
closeAddressBtn.addEventListener('click', () => addressBookModal.classList.add('hidden'));

// ==========================================
// 5. Snippets
// ==========================================
function loadSnippets() {
    const stored = localStorage.getItem('bizRevise_snippets');
    snippets = stored ? JSON.parse(stored) : [];
    renderSnippetList(); renderQuickSnippetDropdown();
}
function saveSnippets() {
    localStorage.setItem('bizRevise_snippets', JSON.stringify(snippets));
    renderSnippetList(); renderQuickSnippetDropdown();
}
function renderSnippetList() {
    snippetListEl.innerHTML = '';
    if (snippets.length === 0) { snippetListEl.innerHTML = '<li class="contact-item">なし</li>'; return; }
    snippets.forEach(s => {
        const li = document.createElement('li'); li.className = 'contact-item';
        li.innerHTML = `
            <div class="contact-info snippet-info"><strong>${s.title}</strong><span>${s.content.substring(0,30)}${s.content.length>30?'...':''}</span></div>
            <div class="contact-actions">
                <button class="icon-btn edit-btn" onclick="editSnippet(${s.id})"><i class="ph ph-pencil-simple"></i></button>
                <button class="icon-btn btn-danger" onclick="deleteSnippet(${s.id})"><i class="ph ph-trash"></i></button>
            </div>`;
        snippetListEl.appendChild(li);
    });
}
function renderQuickSnippetDropdown() {
    quickSnippetList.innerHTML = '';
    if (snippets.length === 0) { quickSnippetList.innerHTML = '<li style="color:#94a3b8;cursor:default;">登録されていません</li>'; return; }
    snippets.forEach(s => {
        const li = document.createElement('li'); li.textContent = s.title;
        li.addEventListener('click', () => {
            const start = inputText.selectionStart; const end = inputText.selectionEnd;
            inputText.value = inputText.value.substring(0, start) + s.content + inputText.value.substring(end);
            inputText.focus(); snippetQuickDropdown.classList.add('hidden');
        });
        quickSnippetList.appendChild(li);
    });
}
snippetForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('snippetId').value;
    const snip = { id: id ? parseInt(id) : Date.now(), title: document.getElementById('snippetTitle').value.trim(), content: document.getElementById('snippetContent').value.trim() };
    if (id) { const idx = snippets.findIndex(x => x.id == id); if (idx > -1) snippets[idx] = snip; } else snippets.push(snip);
    saveSnippets(); clearSnippetForm(); showToast('保存しました');
});
window.editSnippet = (id) => {
    const s = snippets.find(x => x.id == id); if (!s) return;
    document.getElementById('snippetId').value      = s.id;
    document.getElementById('snippetTitle').value   = s.title;
    document.getElementById('snippetContent').value = s.content;
};
window.deleteSnippet = (id) => { if (confirm('削除しますか？')) { snippets = snippets.filter(x => x.id != id); saveSnippets(); } };
function clearSnippetForm() { document.getElementById('snippetId').value = ''; snippetForm.reset(); }
clearSnippetFormBtn.addEventListener('click', clearSnippetForm);
snippetsBtn.addEventListener('click', () => snippetsModal.classList.remove('hidden'));
closeSnippetsBtn.addEventListener('click', () => snippetsModal.classList.add('hidden'));
insertSnippetBtn.addEventListener('click', (e) => { e.stopPropagation(); snippetQuickDropdown.classList.toggle('hidden'); });

// ==========================================
// 6. History
// ==========================================
function loadHistory() {
    const stored = localStorage.getItem('bizRevise_history');
    historyList = stored ? JSON.parse(stored) : [];
    renderHistory();
}
function saveTaskToHistory(obj) {
    historyList.unshift(obj);
    if (historyList.length > 30) historyList.pop();
    localStorage.setItem('bizRevise_history', JSON.stringify(historyList));
    renderHistory();
}
function renderHistory() {
    historyListEl.innerHTML = '';
    if (historyList.length === 0) { historyListEl.innerHTML = '<li style="color:#94a3b8">履歴はありません</li>'; return; }
    historyList.forEach(h => {
        const li = document.createElement('li'); li.className = 'history-item';
        const d = new Date(h.date);
        li.innerHTML = `
            <div class="history-meta"><span><i class="ph ph-file-text"></i> ${h.docType}</span><span>${d.getMonth()+1}/${d.getDate()} ${d.getHours()}:${String(d.getMinutes()).padStart(2,'0')}</span></div>
            <div class="history-content">${h.result.substring(0, 80)}...</div>`;
        li.addEventListener('click', () => {
            inputText.value = h.original;
            outputText.innerText = h.result;
            currentOriginalText = h.original;
            currentGeneratedText = h.result;
            docTypeSelect.value = h.docType;
            if (diffToggle.checked) updateDiffView();
            closeHistoryBtn.click();
            showToast('履歴を読み込みました');
        });
        historyListEl.appendChild(li);
    });
}
historyBtn.addEventListener('click', () => historyModal.classList.remove('hidden'));
closeHistoryBtn.addEventListener('click', () => historyModal.classList.add('hidden'));

// ==========================================
// 7. Diff & Export
// ==========================================
diffToggle.addEventListener('change', () => {
    if (diffToggle.checked) {
        outputText.classList.add('hidden');
        diffOutputText.classList.remove('hidden');
        updateDiffView();
    } else {
        outputText.classList.remove('hidden');
        diffOutputText.classList.add('hidden');
    }
});
function updateDiffView() {
    if (!currentOriginalText || !currentGeneratedText) {
        diffOutputText.innerHTML = '<div style="color:var(--text-secondary)">生成結果がありません。</div>';
        return;
    }
    if (typeof diff_match_patch === 'undefined') {
        diffOutputText.innerHTML = '<span style="color:var(--warning)">差分モジュールの読み込みに失敗しました。オンライン状態を確認してください。</span>';
        return;
    }
    const dmp = new diff_match_patch();
    const d = dmp.diff_main(currentOriginalText, currentGeneratedText);
    dmp.diff_cleanupSemantic(d);
    let html = '';
    d.forEach(([op, data]) => {
        const text = data.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
        if (op === 1)       html += `<ins>${text}</ins>`;
        else if (op === -1) html += `<del>${text}</del>`;
        else                html += `<span>${text}</span>`;
    });
    diffOutputText.innerHTML = html;
}
function exportTxt() {
    if (!currentGeneratedText) { showToast('対象がありません', true); return; }
    const blob = new Blob([currentGeneratedText], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `BizRevise_${Date.now()}.txt`;
    a.click(); URL.revokeObjectURL(url);
}
async function exportDocx() {
    if (!currentGeneratedText) { showToast('対象がありません', true); return; }
    if (typeof docx === 'undefined') { showToast('Wordモジュールが読み込めませんでした', true); return; }
    const lines = currentGeneratedText.split('\n');
    const paragraphs = lines.map(l => new docx.Paragraph({ children: [new docx.TextRun(l)] }));
    const doc = new docx.Document({ sections: [{ properties: {}, children: paragraphs }] });
    try {
        const buffer = await docx.Packer.toBlob(doc);
        const url = URL.createObjectURL(buffer);
        const a = document.createElement('a'); a.href = url; a.download = `BizRevise_${Date.now()}.docx`;
        a.click(); URL.revokeObjectURL(url);
    } catch (err) { showToast('Wordファイルの作成に失敗', true); }
}
downloadTxtBtn.addEventListener('click', exportTxt);
downloadWordBtn.addEventListener('click', exportDocx);
clearBtn.addEventListener('click', () => { inputText.value = ''; inputText.focus(); });
copyBtn.addEventListener('click', () => {
    if (!currentGeneratedText) { showToast('コピーする内容がありません', true); return; }
    navigator.clipboard.writeText(currentGeneratedText).then(() => showToast('コピーしました')).catch(() => showToast('コピー失敗', true));
});

// ==========================================
// 8. ルールベース添削エンジン（APIキー不要）
// ==========================================

// --- 基本修正ルール（正規表現 → 置換, 理由） ---
const BASE_RULES = [
    // 句読点の重複
    { re: /。{2,}/g,  rep: '。',  reason: '句点の重複を修正' },
    { re: /、{2,}/g,  rep: '、',  reason: '読点の重複を修正' },
    { re: /！{2,}/g,  rep: '！',  reason: '感嘆符の重複を修正' },
    { re: /？{2,}/g,  rep: '？',  reason: '疑問符の重複を修正' },
    // 全角スペースの重複
    { re: /　{2,}/g,  rep: '　',  reason: '全角スペースの重複を修正' },
    // 文末の不要スペース
    { re: /[ \t]+。/g, rep: '。', reason: '句点前のスペースを削除' },
    { re: /[ \t]+、/g, rep: '、', reason: '読点前のスペースを削除' },

    // 誤字・誤用（表記の統一）
    { re: /づつ/g,       rep: 'ずつ',       reason: '「ずつ」が正しい表記です' },
    { re: /ずつ(?!ずつ)/g, rep: 'ずつ',     reason: null },
    { re: /ちゃんと/g,   rep: 'きちんと',   reason: 'ビジネス文書では「きちんと」が適切です' },
    { re: /いっぱい/g,   rep: '多く',        reason: 'ビジネス文書では「多く」が適切です' },
    { re: /すごく/g,     rep: '非常に',      reason: 'ビジネス文書では「非常に」が適切です' },
    { re: /ちょっと/g,   rep: '少し',        reason: 'ビジネス文書では「少し」が適切です' },
    { re: /やっぱり/g,   rep: 'やはり',      reason: '「やはり」が正式な表現です' },
    { re: /やっぱ/g,     rep: 'やはり',      reason: '「やはり」が正式な表現です' },
    { re: /だんだん/g,   rep: '徐々に',      reason: 'ビジネス文書では「徐々に」が適切です' },
    { re: /もっと/g,     rep: 'さらに',      reason: 'ビジネス文書では「さらに」が適切です' },
    { re: /うまく/g,     rep: 'うまく',      reason: null },
    { re: /とても/g,     rep: '非常に',      reason: 'ビジネス文書では「非常に」が適切です' },
    { re: /たくさん/g,   rep: '多数',        reason: 'ビジネス文書では「多数」が適切です' },
    { re: /なかなか(?!できない|難し)/g, rep: '容易には', reason: 'ビジネス文書では「容易には」が適切です' },

    // バイト敬語・誤った敬語
    { re: /よろしかったでしょうか/g, rep: 'よろしいでしょうか',  reason: '「よろしかったでしょうか」はバイト敬語です' },
    { re: /なります。/g,             rep: 'です。',               reason: '「〜になります」の誤用を修正しました' },
    { re: /千円になります/g,         rep: '千円でございます',      reason: '金額には「でございます」が適切です' },
    { re: /ご苦労様/g,               rep: 'お疲れ様',              reason: '目上の方には「お疲れ様」が適切です' },
    { re: /了解です/g,               rep: '承知いたしました',      reason: '「了解です」は目上の方への使用は不適切です' },
    { re: /了解しました/g,           rep: '承知いたしました',      reason: '「了解しました」より「承知いたしました」が丁寧です' },
    { re: /わかりました/g,           rep: '承知いたしました',      reason: 'ビジネス文書では「承知いたしました」が適切です' },
    { re: /分かりました/g,           rep: '承知いたしました',      reason: 'ビジネス文書では「承知いたしました」が適切です' },
    { re: /すみません/g,             rep: '申し訳ございません',    reason: 'ビジネス文書では「申し訳ございません」が適切です' },
    { re: /すいません/g,             rep: '申し訳ございません',    reason: 'ビジネス文書では「申し訳ございません」が適切です' },
    { re: /ごめんなさい/g,           rep: '誠に申し訳ございません', reason: 'ビジネス文書では「誠に申し訳ございません」が適切です' },
    { re: /なるほど/g,               rep: 'おっしゃる通りです',    reason: '「なるほど」は目上の方への使用は不適切です' },

    // 冗長表現
    { re: /させていただきます/g,     rep: 'いたします',             reason: '「させていただく」の多用は冗長です' },
    { re: /させていただきました/g,   rep: 'いたしました',           reason: '「させていただく」の多用は冗長です' },
    { re: /〜していただければと思います/g, rep: '〜していただきたく存じます', reason: 'より簡潔な表現に修正しました' },
    { re: /できればと思います/g,     rep: 'できればと考えております', reason: 'より丁寧な表現に修正しました' },
    { re: /という形で/g,             rep: 'として',                 reason: '「という形で」は冗長です' },
    { re: /的な感じで/g,             rep: 'として',                 reason: '「〜的な感じで」は口語的です' },
    { re: /〜の方(は|が|を|に|で)/g, rep: '〜$1',                  reason: '「〜の方」は余分な表現です' },

    // 接続詞の修正
    { re: /でも、/g,    rep: 'しかし、', reason: '「でも」はカジュアルです。「しかし」が適切です' },
    { re: /^でも /gm,   rep: 'しかし ', reason: '「でも」はカジュアルです。「しかし」が適切です' },
    { re: /けど、/g,    rep: 'ですが、', reason: '「けど」はカジュアルです。「ですが」が適切です' },
    { re: /けど。/g,    rep: 'ですが。', reason: '「けど」はカジュアルです。「ですが」が適切です' },
    { re: /だから、/g,  rep: 'そのため、', reason: '「だから」はカジュアルです。「そのため」が適切です' },
    { re: /なので、/g,  rep: 'そのため、', reason: '「なので」より「そのため」がより丁寧です' },

    // 重複表現
    { re: /各(〜)?各/g,        rep: '各',    reason: '「各〜各」の重複を修正しました' },
    { re: /まず最初に/g,       rep: 'まず',  reason: '「まず最初に」は重複表現です' },
    { re: /一番最初/g,         rep: '最初',  reason: '「一番最初」は重複表現です' },
    { re: /一番最後/g,         rep: '最後',  reason: '「一番最後」は重複表現です' },
    { re: /返事を返す/g,       rep: '返事をする', reason: '「返事を返す」は重複表現です' },
    { re: /過去の経歴/g,       rep: '経歴', reason: '「過去の経歴」は重複表現です' },
    { re: /従来から/g,         rep: '従来', reason: '「従来から」は重複表現です' },
    { re: /あとで後ほど/g,     rep: '後ほど', reason: '重複表現を修正しました' },
    { re: /必ず必要/g,         rep: '必要', reason: '「必ず必要」は重複表現です' },

    // よくある誤字
    { re: /以外と/g,    rep: '意外と',   reason: '「意外と」が正しい表記です' },
    { re: /的を得/g,    rep: '的を射',   reason: '「的を射る」が正しい慣用句です' },
    { re: /汚名挽回/g,  rep: '名誉挽回', reason: '「名誉挽回」が正しい表現です' },
    { re: /役不足/g,    rep: '力不足',   reason: '「役不足」は「役が軽すぎる」という意味です。「力不足」が適切な場合が多いです' },
    { re: /確信犯/g,    rep: '故意犯',   reason: '「確信犯」は誤用されがちです。「故意犯」が正確です' },
    { re: /敷居が高い/g, rep: '難しい・ハードルが高い', reason: '「敷居が高い」は「不義理があって行きにくい」の意味です' },
    { re: /煮詰まった/g, rep: '行き詰まった', reason: '「煮詰まった」は「議論が十分つめられた」の意味です' },
    { re: /情けは人の為ならず/g, rep: '情けは人の為ならず（自分のためになるという意味）', reason: '慣用句の意味に注意してください' },

    // 数字・表記の統一
    { re: /([0-9０-９]+)ヶ月/g, rep: '$1か月', reason: '「か月」が正しい表記です' },
    { re: /([0-9０-９]+)ヶ所/g, rep: '$1か所', reason: '「か所」が正しい表記です' },
    { re: /([0-9０-９]+)ケ月/g, rep: '$1か月', reason: '「か月」が正しい表記です' },

    // 文末表現の統一（丁寧体）
    { re: /だ。/g,   rep: 'です。',   reason: '丁寧体「です。」に統一しました' },
    { re: /である。/g, rep: 'です。', reason: '丁寧体「です。」に統一しました' },
    { re: /する。/g,  rep: 'します。', reason: '丁寧体「します。」に統一しました' },
    { re: /した。/g,  rep: 'しました。', reason: '丁寧体「しました。」に統一しました' },
    { re: /ない。/g,  rep: 'ありません。', reason: '丁寧体「ありません。」に統一しました' },
    { re: /いる。/g,  rep: 'おります。', reason: '丁寧体「おります。」に統一しました' },
    { re: /思う。/g,  rep: '思います。', reason: '丁寧体「思います。」に統一しました' },
];

// --- トーン別変換テーブル ---
const TONE_RULES = {
    '丁寧（標準）': [],
    'より丁寧・謙譲': [
        { re: /思います/g,         rep: '存じます',           reason: 'より丁寧な表現に変更しました' },
        { re: /考えています/g,     rep: '考えております',     reason: 'より丁寧な表現に変更しました' },
        { re: /します/g,           rep: 'いたします',         reason: 'より丁寧な表現に変更しました' },
        { re: /しました/g,         rep: 'いたしました',       reason: 'より丁寧な表現に変更しました' },
        { re: /あります/g,         rep: 'ございます',         reason: 'より丁寧な表現に変更しました' },
        { re: /います/g,           rep: 'おります',           reason: 'より丁寧な表現に変更しました' },
        { re: /もらいます/g,       rep: 'いただきます',       reason: 'より丁寧な表現に変更しました' },
        { re: /もらいました/g,     rep: 'いただきました',     reason: 'より丁寧な表現に変更しました' },
        { re: /言います/g,         rep: '申し上げます',       reason: 'より丁寧な表現に変更しました' },
        { re: /お願いします/g,     rep: 'お願い申し上げます', reason: 'より丁寧な表現に変更しました' },
        { re: /見てください/g,     rep: 'ご確認いただけますでしょうか', reason: 'より丁寧な表現に変更しました' },
        { re: /送ります/g,         rep: 'お送りいたします',   reason: 'より丁寧な表現に変更しました' },
        { re: /確認します/g,       rep: 'ご確認いたします',   reason: 'より丁寧な表現に変更しました' },
        { re: /連絡します/g,       rep: 'ご連絡いたします',   reason: 'より丁寧な表現に変更しました' },
        { re: /知らせます/g,       rep: 'お知らせいたします', reason: 'より丁寧な表現に変更しました' },
    ],
    '簡潔・ロジカル': [
        { re: /誠にありがとうございます/g, rep: 'ありがとうございます', reason: '簡潔な表現に変更しました' },
        { re: /お世話になっております。/g, rep: '',  reason: '簡潔化のため挨拶文を省略しました' },
        { re: /〜していただきたく存じます/g, rep: '〜してください', reason: '簡潔な表現に変更しました' },
        { re: /ご検討のほどよろしくお願い申し上げます/g, rep: 'ご検討をお願いします', reason: '簡潔な表現に変更しました' },
    ],
    'フレンドリー・柔らかめ': [
        { re: /ご確認のほど/g,     rep: 'ご確認',             reason: 'より柔らかい表現に変更しました' },
        { re: /申し上げます/g,     rep: 'します',             reason: 'より柔らかい表現に変更しました' },
        { re: /いただきたく/g,     rep: 'いただければ',       reason: 'より柔らかい表現に変更しました' },
    ],
    '厳格・フォーマル': [
        { re: /ですが/g,           rep: 'しかしながら',       reason: '厳格な表現に変更しました' },
        { re: /します/g,           rep: 'いたします',         reason: '厳格な表現に変更しました' },
        { re: /思います/g,         rep: '存じます',           reason: '厳格な表現に変更しました' },
        { re: /お願いします/g,     rep: '願いたく存じます',   reason: '厳格な表現に変更しました' },
        { re: /ありがとうございます/g, rep: '厚く御礼申し上げます', reason: '厳格な表現に変更しました' },
    ],
};

// --- メール用テンプレート（挨拶・締め） ---
const EMAIL_GREETINGS = {
    '社内メール': { open: 'お疲れ様です。',         close: '以上、よろしくお願いいたします。' },
    '社外メール': { open: 'お世話になっております。', close: 'どうぞよろしくお願い申し上げます。' },
    'チャットメッセージ': { open: '',               close: 'よろしくお願いします！' },
    '謝罪文':     { open: 'この度は、誠に申し訳ございませんでした。', close: '今後ともよろしくお願いいたします。' },
    '企画書・提案書': { open: '',                  close: 'ご検討のほど、よろしくお願い申し上げます。' },
    '報告書・日報': { open: '',                    close: '以上、ご報告いたします。' },
};

// --- メイン添削関数 ---
function correctText(text, { docType, tone, actionType, targetStr, customRules }) {
    let result = text;
    const applied = [];  // 適用されたルールのログ

    function applyRules(rules) {
        rules.forEach(({ re, rep, reason }) => {
            const before = result;
            result = result.replace(re, rep);
            if (reason && result !== before) {
                applied.push(reason);
            }
        });
    }

    // 1. 基本ルールを適用
    applyRules(BASE_RULES);

    // 2. 要約モード：文字数削減（各文を短縮）
    if (actionType.includes('要約')) {
        const lines = result.split(/\n+/).filter(l => l.trim());
        result = lines.map(line => {
            // 各段落の最初の文だけ残す（簡易）
            const sentences = line.split(/(?<=。|！|？)/);
            return sentences.slice(0, Math.max(1, Math.ceil(sentences.length / 2))).join('');
        }).join('\n');
        applied.push('要約：文章を約半分に短縮しました');
    }

    // 3. 詳細化モード：補足を追加
    if (actionType.includes('詳細化')) {
        result = result.replace(/です。/g, 'です（詳細については別途ご確認ください）。');
        applied.push('詳細化：補足説明を追加しました');
    }

    // 4. トーンルールを適用
    const toneRules = TONE_RULES[tone] || [];
    applyRules(toneRules);

    // 5. メールの場合、挨拶・締め文の補完
    if (actionType.includes('リライト') || actionType.includes('修正')) {
        const tmpl = EMAIL_GREETINGS[docType];
        if (tmpl) {
            // 挨拶文がなければ先頭に追加
            if (tmpl.open && !result.startsWith(tmpl.open) && !result.includes(tmpl.open)) {
                // 宛名があれば先頭に
                const addressLine = targetStr && targetStr !== '指定なし' ? `${targetStr}\n\n` : '';
                result = addressLine + tmpl.open + '\n\n' + result;
                if (targetStr && targetStr !== '指定なし') applied.push(`宛先「${targetStr}」を補完しました`);
                applied.push(`「${tmpl.open}」の挨拶文を追加しました`);
            }
            // 締め文がなければ末尾に追加
            if (tmpl.close && !result.endsWith(tmpl.close)) {
                result = result.trimEnd() + '\n\n' + tmpl.close;
                applied.push(`「${tmpl.close}」の締め文を追加しました`);
            }
        }
    }

    // 6. ユーザー独自ルールを適用（1行1ルール: 「old→new」形式）
    if (customRules) {
        customRules.split('\n').forEach(line => {
            const m = line.match(/^(.+?)→(.+)$/);
            if (m) {
                const before = result;
                result = result.replace(new RegExp(m[1].trim(), 'g'), m[2].trim());
                if (result !== before) applied.push(`カスタムルール適用: 「${m[1].trim()}」→「${m[2].trim()}」`);
            }
        });
    }

    // 7. 重複した行末のトリム
    result = result.replace(/\n{3,}/g, '\n\n');

    // 変更なしの場合のデフォルトメッセージ
    if (result === text) applied.push('文章に修正すべき箇所は見つかりませんでした（高品質な文章です！）');

    return { correctedText: result, applied: [...new Set(applied)] };
}

// ==========================================
// 9. 生成ボタン（ルールベース処理）
// ==========================================
generateBtn.addEventListener('click', () => {
    const text = inputText.value.trim();
    if (!text) return showToast('添削する文章を入力してください', true);

    // 処理中表示
    loadingOverlay.classList.remove('hidden');
    generateBtn.disabled = true;

    // 宛先取得
    const contactId = targetContactSelect.value;
    let targetStr = '指定なし';
    if (contactId) {
        const c = contacts.find(x => x.id == contactId);
        if (c) targetStr = formatContactName(c);
    }

    const customRules = localStorage.getItem('bizRevise_rules') || '';

    // 少し遅延を入れてUIが反映されるようにする
    setTimeout(() => {
        try {
            const { correctedText, applied } = correctText(text, {
                docType:   docTypeSelect.value,
                tone:      toneSelect.value,
                actionType: actionTypeSelect.value,
                targetStr,
                customRules,
            });

            currentOriginalText  = text;
            currentGeneratedText = correctedText;

            // 修正結果 + 適用ルールの表示
            const rulesHtml = applied.length > 0
                ? `\n\n---【適用した修正ルール (${applied.length}件)】---\n` + applied.map((r, i) => `${i+1}. ${r}`).join('\n')
                : '';

            outputText.innerText = correctedText + rulesHtml;
            saveTaskToHistory({ date: Date.now(), docType: docTypeSelect.value, original: text, result: correctedText });
            if (diffToggle.checked) updateDiffView();
            showToast(`添削完了！ ${applied.length}件の修正を適用しました`);
        } catch (e) {
            showToast(`エラー: ${e.message}`, true);
            console.error(e);
        } finally {
            loadingOverlay.classList.add('hidden');
            generateBtn.disabled = false;
        }
    }, 200);
});

// Init
loadSettings(); loadContacts(); loadSnippets(); loadHistory();
