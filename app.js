// ==========================================
// 1. DOM Elements & Globals
// ==========================================
let contacts = [];
let snippets = [];
let historyList = [];
let currentOriginalText = "";
let currentGeneratedText = "";
// Settings Modal
const settingsBtn = document.getElementById('settingsBtn');
const closeSettingsBtn = document.getElementById('closeSettingsBtn');
const settingsModal = document.getElementById('settingsModal');
const saveSettingsBtn = document.getElementById('saveSettingsBtn');
const apiProviderSelect = document.getElementById('apiProvider');
const apiKeyInput = document.getElementById('apiKey');
const modelSelect = document.getElementById('modelSelect');
const customRulesInput = document.getElementById('customRules');
const claudeCorsWarning = document.getElementById('claudeCorsWarning');
// Address Book Modal
const addressBookBtn = document.getElementById('addressBookBtn');
const closeAddressBtn = document.getElementById('closeAddressBtn');
const addressBookModal = document.getElementById('addressBookModal');
const contactListEl = document.getElementById('contactList');
const contactForm = document.getElementById('contactForm');
const clearContactFormBtn = document.getElementById('clearContactFormBtn');
// Snippets Modal
const snippetsBtn = document.getElementById('snippetsBtn');
const closeSnippetsBtn = document.getElementById('closeSnippetsBtn');
const snippetsModal = document.getElementById('snippetsModal');
const snippetListEl = document.getElementById('snippetList');
const snippetForm = document.getElementById('snippetForm');
const clearSnippetFormBtn = document.getElementById('clearSnippetFormBtn');
const insertSnippetBtn = document.getElementById('insertSnippetBtn');
const snippetQuickDropdown = document.getElementById('snippetQuickDropdown');
const quickSnippetList = document.getElementById('quickSnippetList');
// History Modal
const historyBtn = document.getElementById('historyBtn');
const closeHistoryBtn = document.getElementById('closeHistoryBtn');
const historyModal = document.getElementById('historyModal');
const historyListEl = document.getElementById('historyList');
// Main UI Elements
const targetContactSelect = document.getElementById('targetContact');
const inputText = document.getElementById('inputText');
const outputText = document.getElementById('outputText');
const diffOutputText = document.getElementById('diffOutputText');
const diffToggle = document.getElementById('diffToggle');
const clearBtn = document.getElementById('clearBtn');
const copyBtn = document.getElementById('copyBtn');
const downloadTxtBtn = document.getElementById('downloadTxtBtn');
const downloadWordBtn = document.getElementById('downloadWordBtn');
const generateBtn = document.getElementById('generateBtn');
const docTypeSelect = document.getElementById('docType');
const toneSelect = document.getElementById('tone');
const actionTypeSelect = document.getElementById('actionType');
const toast = document.getElementById('toast');
const toastMsg = document.getElementById('toastMsg');
const toastIcon = toast.querySelector('i');
const loadingOverlay = document.getElementById('loadingOverlay');
// AI Configs
const AI_MODELS = {
    openai: [
        { id: 'gpt-4o',        name: 'GPT-4o (推奨)' },
        { id: 'gpt-4o-mini',   name: 'GPT-4o mini (高速・安価)' },
        { id: 'gpt-3.5-turbo', name: 'GPT-3.5 Turbo (安価)' }
    ],
    gemini: [
        { id: 'gemini-2.0-flash',    name: 'Gemini 2.0 Flash (推奨・最新)' },
        { id: 'gemini-1.5-pro',      name: 'Gemini 1.5 Pro' },
        { id: 'gemini-1.5-flash',    name: 'Gemini 1.5 Flash (高速)' }
    ],
    claude: [
        { id: 'claude-3-5-sonnet-20241022', name: 'Claude 3.5 Sonnet (推奨・最新)' },
        { id: 'claude-3-5-haiku-20241022',  name: 'Claude 3.5 Haiku (高速・安価)' },
        { id: 'claude-3-opus-20240229',     name: 'Claude 3 Opus (最高精度)' }
    ]
};
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
    toastTimeout = setTimeout(() => toast.classList.add('hidden'), 3000);
}
// ==========================================
// 3. Settings (API, Custom Rules)
// ==========================================
function updateModelOptions(provider, selectedModel = null) {
    modelSelect.innerHTML = '';
    const models = AI_MODELS[provider] || [];
    models.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.id; opt.textContent = m.name;
        modelSelect.appendChild(opt);
    });
    if (selectedModel && models.find(m => m.id === selectedModel)) modelSelect.value = selectedModel;
    if (provider === 'claude') claudeCorsWarning.classList.remove('hidden');
    else claudeCorsWarning.classList.add('hidden');
}
apiProviderSelect.addEventListener('change', (e) => updateModelOptions(e.target.value));
function loadSettings() {
    const provider = localStorage.getItem('bizRevise_provider') || 'openai';
    const key = localStorage.getItem(`bizRevise_apiKey_${provider}`);
    const model = localStorage.getItem('bizRevise_model');
    const rules = localStorage.getItem('bizRevise_rules') || '';
    apiProviderSelect.value = provider;
    updateModelOptions(provider, model);
    apiKeyInput.value = key ? key : '';
    customRulesInput.value = rules;
}
saveSettingsBtn.addEventListener('click', () => {
    const provider = apiProviderSelect.value;
    const key = apiKeyInput.value.trim();
    
    localStorage.setItem('bizRevise_provider', provider);
    localStorage.setItem('bizRevise_model', modelSelect.value);
    localStorage.setItem('bizRevise_rules', customRulesInput.value);
    if (key) localStorage.setItem(`bizRevise_apiKey_${provider}`, key);
    
    settingsModal.classList.add('hidden');
    showToast('設定を保存しました');
});
settingsBtn.addEventListener('click', () => { loadSettings(); settingsModal.classList.remove('hidden'); });
closeSettingsBtn.addEventListener('click', () => settingsModal.classList.add('hidden'));
// ==========================================
// 4. Contact / Address Book Management
// ==========================================
function loadContacts() {
    const stored = localStorage.getItem('bizRevise_contacts');
    contacts = stored ? JSON.parse(stored) : [];
    renderContactList(); renderContactDropdown();
}
function saveContacts() { localStorage.setItem('bizRevise_contacts', JSON.stringify(contacts)); renderContactList(); renderContactDropdown(); }
function formatContactName(c) {
    const parts = [];
    if(c.company) parts.push(c.company);
    if(c.role) parts.push(c.role);
    if(c.name) parts.push(c.name + " 様");
    return parts.join(" ");
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
    if(contacts.length === 0) { contactListEl.innerHTML = '<li class="contact-item"><span style="color:#94a3b8">登録がありません</span></li>'; return; }
    contacts.forEach(c => {
        const li = document.createElement('li'); li.className = 'contact-item';
        li.innerHTML = `
            <div class="contact-info"><strong>${c.name}</strong><span>${c.company||''} ${c.role||''}</span></div>
            <div class="contact-actions">
                <button class="icon-btn edit-btn" onclick="editContact(${c.id})"><i class="ph ph-pencil-simple"></i></button>
                <button class="icon-btn btn-danger" onclick="deleteContact(${c.id})"><i class="ph ph-trash"></i></button>
            </div>
        `;
        contactListEl.appendChild(li);
    });
}
contactForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('contactId').value;
    const contact = {
        id: id ? parseInt(id) : Date.now(), company: document.getElementById('contactCompany').value.trim(),
        role: document.getElementById('contactRole').value.trim(), name: document.getElementById('contactName').value.trim(),
        email: document.getElementById('contactEmail').value.trim()
    };
    if (id) { const idx = contacts.findIndex(x => x.id == id); if(idx>-1) contacts[idx] = contact; } else { contacts.push(contact); }
    saveContacts(); clearContactForm(); showToast('保存しました');
});
window.editContact = (id) => {
    const c = contacts.find(x => x.id == id); if(!c) return;
    document.getElementById('contactId').value = c.id; document.getElementById('contactCompany').value = c.company;
    document.getElementById('contactRole').value = c.role; document.getElementById('contactName').value = c.name;
    document.getElementById('contactEmail').value = c.email;
};
window.deleteContact = (id) => { if(confirm('削除しますか？')){ contacts = contacts.filter(x => x.id != id); saveContacts(); } };
function clearContactForm() { document.getElementById('contactId').value = ''; contactForm.reset(); }
clearContactFormBtn.addEventListener('click', clearContactForm);
addressBookBtn.addEventListener('click', () => addressBookModal.classList.remove('hidden'));
closeAddressBtn.addEventListener('click', () => addressBookModal.classList.add('hidden'));
// ==========================================
// 5. Snippets Management
// ==========================================
function loadSnippets() {
    const stored = localStorage.getItem('bizRevise_snippets');
    snippets = stored ? JSON.parse(stored) : [];
    renderSnippetList(); renderQuickSnippetDropdown();
}
function saveSnippets() { localStorage.setItem('bizRevise_snippets', JSON.stringify(snippets)); renderSnippetList(); renderQuickSnippetDropdown(); }
function renderSnippetList() {
    snippetListEl.innerHTML = '';
    if(snippets.length === 0) { snippetListEl.innerHTML = '<li class="contact-item">なし</li>'; return; }
    snippets.forEach(s => {
        const li = document.createElement('li'); li.className = 'contact-item';
        li.innerHTML = `
            <div class="contact-info snippet-info"><strong>${s.title}</strong><span>${s.content.substring(0,30)}${s.content.length>30?'...':''}</span></div>
            <div class="contact-actions">
                <button class="icon-btn edit-btn" onclick="editSnippet(${s.id})"><i class="ph ph-pencil-simple"></i></button>
                <button class="icon-btn btn-danger" onclick="deleteSnippet(${s.id})"><i class="ph ph-trash"></i></button>
            </div>
        `;
        snippetListEl.appendChild(li);
    });
}
function renderQuickSnippetDropdown() {
    quickSnippetList.innerHTML = '';
    if(snippets.length === 0) { quickSnippetList.innerHTML = '<li style="color:#94a3b8;cursor:default;">登録されていません</li>'; return; }
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
    if (id) { const idx = snippets.findIndex(x => x.id == id); if(idx>-1) snippets[idx] = snip; } else { snippets.push(snip); }
    saveSnippets(); clearSnippetForm(); showToast('保存しました');
});
window.editSnippet = (id) => {
    const s = snippets.find(x => x.id == id); if(!s) return;
    document.getElementById('snippetId').value = s.id; document.getElementById('snippetTitle').value = s.title; document.getElementById('snippetContent').value = s.content;
};
window.deleteSnippet = (id) => { if(confirm('削除しますか？')){ snippets = snippets.filter(x => x.id != id); saveSnippets(); } };
function clearSnippetForm() { document.getElementById('snippetId').value = ''; snippetForm.reset(); }
clearSnippetFormBtn.addEventListener('click', clearSnippetForm);
snippetsBtn.addEventListener('click', () => snippetsModal.classList.remove('hidden'));
closeSnippetsBtn.addEventListener('click', () => snippetsModal.classList.add('hidden'));
insertSnippetBtn.addEventListener('click', (e) => {
    e.stopPropagation(); snippetQuickDropdown.classList.toggle('hidden');
});
// ==========================================
// 6. History Management
// ==========================================
function loadHistory() {
    const stored = localStorage.getItem('bizRevise_history');
    historyList = stored ? JSON.parse(stored) : [];
    renderHistory();
}
function saveTaskToHistory(historyObj) {
    historyList.unshift(historyObj);
    if(historyList.length > 30) historyList.pop();
    localStorage.setItem('bizRevise_history', JSON.stringify(historyList));
    renderHistory();
}
function renderHistory() {
    historyListEl.innerHTML = '';
    if(historyList.length === 0) { historyListEl.innerHTML = '<li style="color:#94a3b8">履歴はありません</li>'; return; }
    historyList.forEach(h => {
        const li = document.createElement('li'); li.className = 'history-item';
        const d = new Date(h.date);
        li.innerHTML = `
            <div class="history-meta"><span><i class="ph ph-file-text"></i> ${h.docType}</span><span>${d.getMonth()+1}/${d.getDate()} ${d.getHours()}:${String(d.getMinutes()).padStart(2,'0')}</span></div>
            <div class="history-content">${h.result.substring(0, 80)}...</div>
        `;
        li.addEventListener('click', () => {
            inputText.value = h.original;
            outputText.innerText = h.result;
            currentOriginalText = h.original;
            currentGeneratedText = h.result;
            docTypeSelect.value = h.docType;
            if(diffToggle.checked) updateDiffView();
            closeHistoryBtn.click();
            showToast('履歴を読み込みました');
        });
        historyListEl.appendChild(li);
    });
}
historyBtn.addEventListener('click', () => historyModal.classList.remove('hidden'));
closeHistoryBtn.addEventListener('click', () => historyModal.classList.add('hidden'));
// ==========================================
// 7. Diff Highlight & Export Features
// ==========================================
diffToggle.addEventListener('change', () => {
    if(diffToggle.checked) {
        outputText.classList.add('hidden');
        diffOutputText.classList.remove('hidden');
        updateDiffView();
    } else {
        outputText.classList.remove('hidden');
        diffOutputText.classList.add('hidden');
    }
});
function updateDiffView() {
    if(!currentOriginalText || !currentGeneratedText) {
        diffOutputText.innerHTML = '<div style="color:var(--text-secondary)">生成結果がありません。</div>';
        return;
    }
    // Check if CDN is loaded
    if(typeof diff_match_patch === 'undefined') {
        diffOutputText.innerHTML = '<span style="color:var(--warning)">差分モジュールの読み込みに失敗しました。オンライン状態を確認してください。</span>';
        return;
    }
    
    var dmp = new diff_match_patch();
    var d = dmp.diff_main(currentOriginalText, currentGeneratedText);
    dmp.diff_cleanupSemantic(d);
    
    // Custom HTML rendering for diff
    let html = "";
    for (var i = 0; i < d.length; i++) {
        let op = d[i][0]; let data = d[i][1];
        let text = data.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        text = text.replace(/\n/g, '<br>');
        if (op === 1) html += "<ins>" + text + "</ins>";
        else if (op === -1) html += "<del>" + text + "</del>";
        else html += "<span>" + text + "</span>";
    }
    diffOutputText.innerHTML = html;
}
function exportTxt() {
    if(!currentGeneratedText) { showToast("対象がありません", true); return; }
    const blob = new Blob([currentGeneratedText], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `BizRevise_${Date.now()}.txt`;
    a.click(); URL.revokeObjectURL(url);
}
async function exportDocx() {
    if(!currentGeneratedText) { showToast("対象がありません", true); return; }
    if(typeof docx === 'undefined') { showToast("Wordモジュールが読み込めませんでした", true); return; }
    
    const lines = currentGeneratedText.split('\n');
    const paragraphs = lines.map(line => new docx.Paragraph({ children: [new docx.TextRun(line)] }));
    
    const doc = new docx.Document({
        sections: [{ properties: {}, children: paragraphs }]
    });
    
    try {
        const buffer = await docx.Packer.toBlob(doc);
        const url = URL.createObjectURL(buffer);
        const a = document.createElement('a'); a.href = url; a.download = `BizRevise_${Date.now()}.docx`;
        a.click(); URL.revokeObjectURL(url);
    } catch(err) {
        showToast("Wordファイルの作成に失敗", true); console.error(err);
    }
}
downloadTxtBtn.addEventListener('click', exportTxt);
downloadWordBtn.addEventListener('click', exportDocx);
// ==========================================
// 8. AI Core Logic
// ==========================================
clearBtn.addEventListener('click', () => { inputText.value = ''; inputText.focus(); });
copyBtn.addEventListener('click', () => {
    if (!currentGeneratedText) { showToast('コピーする内容がありません', true); return; }
    navigator.clipboard.writeText(currentGeneratedText).then(() => showToast('コピーしました')).catch(() => showToast('コピー失敗', true));
});
async function callAPI(provider, apiKey, model, systemPrompt, userText) {
    if (provider === 'openai') {
        const res = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
            body: JSON.stringify({
                model: model || 'gpt-4o',
                messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userText }],
                temperature: 0.3
            })
        });
        if (!res.ok) { const e = await res.json(); throw new Error(e.error?.message || `OpenAI Error ${res.status}`); }
        return (await res.json()).choices[0].message.content;

    } else if (provider === 'gemini') {
        const geminiModel = model || 'gemini-2.0-flash';
        const res = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${apiKey}`,
            {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ role: 'user', parts: [{ text: userText }] }],
                    systemInstruction: { role: 'system', parts: [{ text: systemPrompt }] },
                    generationConfig: { temperature: 0.3 }
                })
            }
        );
        if (!res.ok) { const e = await res.json(); throw new Error(e.error?.message || `Gemini Error ${res.status}`); }
        const data = await res.json();
        const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
        if (!text) throw new Error('Geminiから応答を取得できませんでした');
        return text;

    } else if (provider === 'claude') {
        const claudeModel = model || 'claude-3-5-sonnet-20241022';
        const res = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': apiKey,
                'anthropic-version': '2023-06-01',
                'anthropic-dangerously-allow-browser': 'true'
            },
            body: JSON.stringify({
                model: claudeModel,
                system: systemPrompt,
                messages: [{ role: 'user', content: userText }],
                max_tokens: 4096,
                temperature: 0.3
            })
        });
        if (!res.ok) { const e = await res.json(); throw new Error(e.error?.message || `Claude Error ${res.status}`); }
        const data = await res.json();
        const text = data?.content?.[0]?.text;
        if (!text) throw new Error('Claudeから応答を取得できませんでした');
        return text;
    }
    throw new Error('不明なプロバイダー: ' + provider);
}
generateBtn.addEventListener('click', async () => {
    const text = inputText.value.trim();
    if (!text) return showToast('添削する文章を入力してください', true);
    const provider = localStorage.getItem('bizRevise_provider') || 'openai';
    const apiKey = localStorage.getItem(`bizRevise_apiKey_${provider}`);
    const model = localStorage.getItem('bizRevise_model');
    const customRules = localStorage.getItem('bizRevise_rules') || '';
    if (!apiKey) {
        showToast(`APIキーが未設定です (${provider.toUpperCase()})`, true);
        return settingsModal.classList.remove('hidden');
    }
    loadingOverlay.classList.remove('hidden'); generateBtn.disabled = true;
    // Contact Logic
    const contactId = targetContactSelect.value;
    let targetStr = "指定なし";
    if (contactId) { const c = contacts.find(x => x.id == contactId); if(c) targetStr = formatContactName(c); }
    const actionType = actionTypeSelect.value;
    let actionInstruction = '';
    if (actionType.includes('誤字脱字') || actionType.includes('修正のみ')) {
        actionInstruction = `
【添削の具体的な指示】
- 誤字・脱字を発見し必ず修正すること
- 助詞の誤り（は/が/を/に等）を修正すること
- 句読点の位置が不自然な箇所を修正すること
- 二重否定・意味の通らない表現を修正すること
- 敬語の誤り（尊敬語・謙譲語の混同等）を修正すること
- ビジネス文書として不自然な表現を全て改善すること
- 修正箇所がなくても必ずより自然な表現に改善すること`;
    } else if (actionType.includes('リライト')) {
        actionInstruction = `
【リライトの指示】
- 全体の構成を整理し読みやすい順序に並べ直すこと
- 冗長な表現を簡潔にすること
- 誤字脱字・敬語の誤りも同時に修正すること
- 段落分けを適切に行うこと`;
    } else if (actionType.includes('要約')) {
        actionInstruction = `【要約の指示】重要な情報を漏らさず短くまとめること。箇条書きを活用して読みやすくすること。`;
    } else if (actionType.includes('詳細化')) {
        actionInstruction = `【詳細化の指示】根拠・理由・具体例を追加して内容を充実させること。曖昧な表現を具体的にすること。`;
    }
    const systemPrompt = `あなたは日本語ビジネス文書の専門校正者です。入力された文章を以下の条件で必ず修正・リライトしてください。

【基本条件】
- 文章の種類: ${docTypeSelect.value}
- トーン・スタイル: ${toneSelect.value}
- 処理内容: ${actionType}
- 宛先情報: ${targetStr}
${actionInstruction}

【出力ルール】
1. 修正後の文章のみを出力すること（説明・前置き・コメントは一切不要）
2. 宛名が必要な文書で宛先情報がある場合は冒頭に補完すること
3. 元の文章から何も変更しない場合でも、必ず改善点を見つけて修正すること
4. 「以下に修正しました」等の説明文は含めないこと
${customRules ? `\n【ユーザー固有の必須ルール】\n${customRules}` : ''}`;
    try {
        const result = await callAPI(provider, apiKey, model, systemPrompt, text);
        currentOriginalText = text;
        currentGeneratedText = result;
        outputText.innerText = result;
        
        saveTaskToHistory({ date: Date.now(), docType: docTypeSelect.value, original: text, result: result });
        
        if (diffToggle.checked) updateDiffView();
        showToast('リライト完了！');
    } catch (e) {
        showToast(`エラー: ${e.message}`, true);
    } finally {
        loadingOverlay.classList.add('hidden'); generateBtn.disabled = false;
    }
});
// Init
loadSettings(); loadContacts(); loadSnippets(); loadHistory();
