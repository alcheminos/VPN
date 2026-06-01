let userData = {};

document.addEventListener("DOMContentLoaded", () => {
    try {
        document.getElementById('btnOpenSettings').addEventListener('click', openSettings);
        document.getElementById('btnSaveSettings').addEventListener('click', saveSettings);
        
        document.getElementById('btnLoadIp1').addEventListener('click', () => loadCurrentIpToInput('setIp1', 'btnLoadIp1'));
        document.getElementById('btnLoadIp2').addEventListener('click', () => loadCurrentIpToInput('setIp2', 'btnLoadIp2'));
        
        document.getElementById('tabBtnExtend').addEventListener('click', (e) => switchTab(e, 'tabExtend'));
        document.getElementById('tabBtnNewAccount').addEventListener('click', (e) => switchTab(e, 'tabNewAccount'));
        
        document.getElementById('btnExtend').addEventListener('click', processExtendVpn);
        document.getElementById('btnNewAccount').addEventListener('click', processNewAccount);
    } catch(e) {
        console.error("이벤트 바인딩 실패:", e);
    }

    try {
        if (typeof flatpickr !== 'undefined') {
            flatpickr("#datePicker", { mode: "multiple", dateFormat: "Y-m-d", locale: "ko" });
        }
    } catch(e) { console.error("달력 로드 에러:", e); }
    
    try {
        chrome.storage.local.get(['vpnUserData'], (result) => {
            if(result.vpnUserData) {
                userData = result.vpnUserData;
                updateProfileUI();
                showView('mainWorkspaceView');
            }
        });
    } catch(e) { console.error("스토리지 로드 에러:", e); }
});

function showView(viewId) {
    document.getElementById('settingsView').classList.add('hidden');
    document.getElementById('mainWorkspaceView').classList.add('hidden');
    document.getElementById(viewId).classList.remove('hidden');
}

function openSettings() { 
    if(userData.name) document.getElementById('setName').value = userData.name;
    if(userData.id) document.getElementById('setId').value = userData.id.replace(/^skb/i, '');
    if(userData.dept) document.getElementById('setDept').value = userData.dept;
    if(userData.ip1) document.getElementById('setIp1').value = userData.ip1;
    if(userData.ip2) document.getElementById('setIp2').value = userData.ip2;
    if(userData.jiraId) document.getElementById('setJiraId').value = userData.jiraId;
    if(userData.phone) document.getElementById('setPhone').value = userData.phone;
    showView('settingsView'); 
}

function switchTab(event, tabId) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    event.target.classList.add('active');
    document.getElementById(tabId).classList.add('active');
}

function updateProfileUI() {
    document.getElementById('dispName').textContent = `${userData.name} (${userData.id})`;
    document.getElementById('dispDept').textContent = userData.dept;
    let ipOptions = `<option value="${userData.ip1}">IP 주소 1: ${userData.ip1}</option>`;
    if(userData.ip2) ipOptions += `<option value="${userData.ip2}">IP 주소 2: ${userData.ip2}</option>`;
    document.querySelectorAll('.ip-selector').forEach(s => s.innerHTML = ipOptions);
}

function saveSettings() {
    // 💡 저장할 때는 입력된 숫자에 무조건 'skb'를 붙여서 저장
    let rawId = document.getElementById('setId').value.trim().replace(/^skb/i, '');

    userData = {
        name: document.getElementById('setName').value.trim(),
        id: 'skb' + rawId, 
        jiraId: document.getElementById('setJiraId').value.trim(),
        dept: document.getElementById('setDept').value.trim(),
        phone: document.getElementById('setPhone').value.trim(), 
        ip1: document.getElementById('setIp1').value.trim(),
        ip2: document.getElementById('setIp2').value.trim()
    };
    if(!userData.name || !rawId || !userData.jiraId || !userData.phone || !userData.ip1) 
        return alert("이름, 사번, Jira ID, 연락처, IP 주소 1은 필수입니다."); 
    
    chrome.storage.local.set({vpnUserData: userData}, () => {
        updateProfileUI();
        showView('mainWorkspaceView');
    });
}

async function loadCurrentIpToInput(inputId, btnId) {
    const btn = document.getElementById(btnId);
    btn.textContent = "조회 중...";
    btn.disabled = true;

    try {
        const res = await fetch('https://api.ipify.org?format=json');
        if (!res.ok) throw new Error("API 통신 실패");
        const data = await res.json();
        document.getElementById(inputId).value = data.ip;
        btn.textContent = "입력 완료";
    } catch (e) {
        alert("IP를 불러오지 못했습니다. 사내망 차단 여부를 확인하세요.");
        btn.textContent = "현재 위치 IP 입력";
    }
    
    setTimeout(() => {
        btn.textContent = "현재 위치 IP 입력";
        btn.disabled = false;
    }, 2000);
}

const logger = {
    box: document.getElementById('statusLog'),
    clear: function() { this.box.innerHTML = ''; this.box.classList.remove('hidden'); },
    log: function(msg) { this.box.innerHTML += `<div>> ${msg}</div>`; this.box.scrollTop = this.box.scrollHeight; }
};

async function processExtendVpn() {
    const datesStr = document.getElementById("datePicker").value;
    const btn = document.getElementById("btnExtend");
    if (!datesStr) return alert("날짜 선택 필수");
    const dates = datesStr.split(", ").sort();
    
    btn.disabled = true;
    logger.clear();
    logger.log("Jira 일괄 생성 프로세스 시작...");

    const extendData = {
        ip: document.getElementById("extendVpnIp").value,
        reason: document.getElementById("reason").value,
        startTime: document.getElementById("usageStartTime").value,
        endTime: document.getElementById("usageEndTime").value,
        user: userData
    };

    for (let date of dates) {
        logger.log(`⏳ [${date}] 요청 전송 중...`);
        try {
            const res = await new Promise((resolve, reject) => {
                chrome.runtime.sendMessage({
                    action: "EXTEND_VPN",
                    data: { ...extendData, date }
                }, response => {
                    if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
                    else resolve(response);
                });
            });
            if(res.error) throw new Error(res.error);
            logger.log(`✅ [${res.issueKey}] 접수 완료!`);
        } catch(e) {
            logger.log(`❌ [${date}] 실패: ${e.message}`);
        }
    }
    btn.disabled = false;
}

async function processNewAccount() {
    // 💡 체크박스 중 '기타' 선택 시 직접 입력한 텍스트를 배열에 넣는 로직 추가
    let hasError = false;
    const systems = [];
    document.querySelectorAll('input[name="targetSystem"]:checked').forEach(cb => {
        if(cb.value === '기타') {
            const otherVal = document.getElementById('otherSystemInput').value.trim();
            if(otherVal) systems.push(otherVal);
            else hasError = true;
        } else {
            systems.push(cb.value);
        }
    });

    if (hasError) return alert("기타 시스템 이름을 직접 입력해주세요.");
    if (systems.length === 0) return alert("대상 시스템을 1개 이상 선택해주세요.");

    const btn = document.getElementById("btnNewAccount");
    btn.disabled = true;
    btn.textContent = "Jira API 호출 중...";

    const newAccountData = {
        systems, // 💡 직접 입력된 텍스트가 포함된 systems 배열 전달
        ip: document.getElementById("newAccountIp").value,
        user: userData
    };

    try {
        const res = await new Promise((resolve, reject) => {
            chrome.runtime.sendMessage({
                action: "CREATE_NEW_ACCOUNT",
                data: newAccountData
            }, response => {
                if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
                else resolve(response);
            });
        });
        if(res.error) throw new Error(res.error);
        alert(`✅ 신규 계정 신청 완료 (${res.issueKey})\n엑셀 자동 첨부됨.`);
    } catch(e) {
        alert(`API 에러: ${e.message}`);
    }
    
    btn.disabled = false;
    btn.textContent = "Jira 자동 생성 (엑셀 첨부)";
}

document.addEventListener('DOMContentLoaded', () => {
    const phoneInput = document.getElementById('setPhone');
    
    if (phoneInput) {
        phoneInput.addEventListener('input', function (e) {
            // 숫자 이외의 문자는 모두 제거
            let val = this.value.replace(/[^0-9]/g, '');
            
            // 길이에 따라 하이픈(-) 자동 삽입
            if (val.length < 4) {
                this.value = val;
            } else if (val.length < 10) {
                // 9자리 이하 (예: 02-123-4567 또는 010-123-456)
                this.value = val.substring(0, 3) + '-' + val.substring(3, 6) + '-' + val.substring(6);
            } else {
                // 10~11자리 정상 휴대폰 번호 (예: 010-1234-5678)
                this.value = val.substring(0, 3) + '-' + val.substring(3, 7) + '-' + val.substring(7);
            }
        });
    }
    const chkOther = document.getElementById('chkOther');
    if (chkOther) {
        chkOther.addEventListener('change', function() {
            const input = document.getElementById('otherSystemInput');
            if (this.checked) {
                input.classList.remove('hidden');
                input.focus();
            } else {
                input.classList.add('hidden');
                input.value = '';
            }
        });
    }
});
