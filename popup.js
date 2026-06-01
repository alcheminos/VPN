let userData = {};
let additionalMembers = []; // 💡 동료 명단 배열 (전역 변수)

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

        // 💡 동료 추가 버튼 이벤트 리스너
        document.getElementById('btnAddMember').addEventListener('click', addMember);
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

// 💡 동료 추가 기능 함수들
function addMember() {
    const name = document.getElementById('addMemberName').value.trim();
    const id = document.getElementById('addMemberId').value.trim();
    
    if(!name || id.length !== 4) return alert("이름과 사번 4자리를 정확히 입력해주세요.");
    if(additionalMembers.some(m => m.id === 'skb' + id)) return alert("이미 추가된 사번입니다.");

    additionalMembers.push({ name, id: 'skb' + id });
    renderMemberList();
    
    document.getElementById('addMemberName').value = '';
    document.getElementById('addMemberId').value = '';
}

function renderMemberList() {
    const container = document.getElementById('memberList');
    container.innerHTML = additionalMembers.map((m, index) => 
        `<span style="background: #ebecf0; padding: 3px 8px; border-radius: 10px; display: flex; align-items: center;">
            ${m.name}(${m.id}) <span style="margin-left:6px; cursor:pointer; font-weight:bold; color:#ff5630;" onclick="removeMember(${index})">×</span>
        </span>`
    ).join('');
}

window.removeMember = (index) => {
    additionalMembers.splice(index, 1);
    renderMemberList();
};
// ----------------------------

async function processExtendVpn() {
    const datesStr = document.getElementById("datePicker").value;
    const btn = document.getElementById("btnExtend");
    if (!datesStr) return alert("날짜 선택 필수");
    const dates = datesStr.split(", ").sort();
    
    btn.disabled = true;
    logger.clear();
    logger.log("VPN 활성화 프로세스 시작...");

    const baseData = {
        ip: document.getElementById("extendVpnIp").value,
        reason: document.getElementById("reason").value,
        startTime: document.getElementById("usageStartTime").value,
        endTime: document.getElementById("usageEndTime").value
    };

    // 💡 신청 대상을 '본인 + 추가된 동료'로 합침
    const targetUsers = [
        { name: userData.name, id: userData.id, dept: userData.dept, jiraId: userData.jiraId }, 
        ...additionalMembers.map(m => ({ ...m, dept: userData.dept, jiraId: userData.jiraId }))
    ];

    for (let user of targetUsers) {
        for (let date of dates) {
            logger.log(`⏳ [${user.name} / ${date}] 요청 중...`);
            try {
                const res = await new Promise((resolve, reject) => {
                    chrome.runtime.sendMessage({
                        action: "EXTEND_VPN",
                        data: { ...baseData, date, user } 
                    }, response => {
                        if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
                        else resolve(response);
                    });
                });
                if(res.error) throw new Error(res.error);
                logger.log(`✅ [${res.issueKey}] ${user.name} 완료`);
            } catch(e) {
                logger.log(`❌ [${user.name} / ${date}] 실패: ${e.message}`);
            }
        }
    }
    btn.disabled = false;
}

async function processNewAccount() {
    let hasError = false;
    const systems = [];
    document.querySelectorAll('input[name="targetSystem"]:checked').forEach(cb => {
        if(cb.value === '기타') {
            const oIp = document.getElementById('otherIp').value.trim();
            const oPort = document.getElementById('otherPort').value.trim();
            const oUsage = document.getElementById('otherUsage').value.trim();
            
            if(oIp && oPort && oUsage) {
                systems.push({ type: 'other', ip: oIp, port: oPort, usage: oUsage });
            } else {
                hasError = true;
            }
        } else {
            systems.push(cb.value);
        }
    });

    if (hasError) return alert("기타 시스템의 IP, Port, 용도를 모두 입력해주세요.");
    if (systems.length === 0) return alert("대상 시스템을 1개 이상 선택해주세요.");

    const btn = document.getElementById("btnNewAccount");
    btn.disabled = true;
    btn.textContent = "Jira API 호출 중...";

    const newAccountData = {
        systems, 
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

// 💡 이벤트 리스너 묶음 (전화번호 하이픈 & 체크박스 연동)
document.addEventListener('DOMContentLoaded', () => {
    // 1. 전화번호 하이픈
    const phoneInput = document.getElementById('setPhone');
    if (phoneInput) {
        phoneInput.addEventListener('input', function (e) {
            let val = this.value.replace(/[^0-9]/g, '');
            if (val.length < 4) {
                this.value = val;
            } else if (val.length < 10) {
                this.value = val.substring(0, 3) + '-' + val.substring(3, 6) + '-' + val.substring(6);
            } else {
                this.value = val.substring(0, 3) + '-' + val.substring(3, 7) + '-' + val.substring(7);
            }
        });
    }

    // 2. 기타 체크박스 활성화 시 숨김 메뉴 오픈
    const chkOther = document.getElementById('chkOther');
    if (chkOther) {
        chkOther.addEventListener('change', function() {
            const inputDiv = document.getElementById('otherSystemInputs');
            if (this.checked) {
                inputDiv.classList.remove('hidden');
                document.getElementById('otherIp').focus();
            } else {
                inputDiv.classList.add('hidden');
                document.getElementById('otherIp').value = '';
                document.getElementById('otherPort').value = '';
                document.getElementById('otherUsage').value = '';
            }
        });
    }
    
    // 3. 트리 체크박스 연동(모두 선택 / 하위 메뉴 연동)
    const chkSelectAll = document.getElementById('chkSelectAll');
    const grpCbs = document.querySelectorAll('.grp-cb');
    const subCbs = document.querySelectorAll('.sub-cb');

    if (chkSelectAll) {
        chkSelectAll.addEventListener('change', function() {
            const isChecked = this.checked;
            grpCbs.forEach(cb => cb.checked = isChecked);
            subCbs.forEach(cb => cb.checked = isChecked);
        });
    }

    grpCbs.forEach(grp => {
        grp.addEventListener('change', function() {
            const isChecked = this.checked;
            const targetClass = this.getAttribute('data-target');
            const children = document.querySelectorAll(`.sub-cb.${targetClass}`);
            children.forEach(cb => cb.checked = isChecked);
            updateSelectAllState();
        });
    });

    subCbs.forEach(sub => {
        sub.addEventListener('change', function() {
            const classes = Array.from(this.classList);
            const groupClass = classes.find(c => c !== 'sub-cb' && document.querySelector(`.grp-cb[data-target="${c}"]`));
            
            if (groupClass) {
                const parentGrp = document.querySelector(`.grp-cb[data-target="${groupClass}"]`);
                const siblings = document.querySelectorAll(`.sub-cb.${groupClass}`);
                const allChecked = Array.from(siblings).every(c => c.checked);
                const someChecked = Array.from(siblings).some(c => c.checked);
                
                parentGrp.checked = allChecked;
                parentGrp.indeterminate = !allChecked && someChecked; 
            }
            updateSelectAllState();
        });
    });

    function updateSelectAllState() {
        if (!chkSelectAll) return;
        const allSubCbs = Array.from(document.querySelectorAll('.sub-cb'));
        const allChecked = allSubCbs.every(c => c.checked);
        const someChecked = allSubCbs.some(c => c.checked);
        
        chkSelectAll.checked = allChecked;
        chkSelectAll.indeterminate = !allChecked && someChecked;
    }
});
