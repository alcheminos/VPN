importScripts('xlsx.full.min.js');

const JIRA_BASE_URL = "https://jira.skbroadband.com";
const PROJECT_KEY = "BTVMKT"; 
const TRANSITION_ID_RECEIPT = "4"; 

const SYSTEM_DESTINATIONS = {
    "EUXP": { ip: "1.255.152.40", port: "TCP 8080, 8443", usage: "EUXP 접속용" },
    "통합ES": { ip: "1.255.140.10", port: "5601", usage: "Web(Kibana 등)" },
    "ACS": { ip: "1.255.152.80", port: "80", usage: "Web" },
    "ECDN": { ip: "121.125.63.21", port: "80", usage: "Web" },
    "상용DB": { ip: "114.202.130.168", port: "9090", usage: "HTTP(DB접근)" }
};

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function fetchJira(endpoint, method, body = null, isFormData = false) {
    let url = `${JIRA_BASE_URL}${endpoint}`;
    let options = {
        method: method,
        credentials: 'include',
        headers: { 'X-Atlassian-Token': 'no-check' }
    };
    if (body) {
        if (isFormData) { options.body = body; } 
        else { options.headers['Content-Type'] = 'application/json'; options.body = JSON.stringify(body); }
    }

    for (let i = 0; i < 3; i++) {
        const res = await fetch(url, options);
        if (res.status === 429) { await sleep((res.headers.get('Retry-After') || 2) * 1000); continue; }
        if (!res.ok && res.status !== 204) {
            // 👇 Jira가 보낸 상세 에러 메시지를 추출하도록 수정
            const errText = await res.text(); 
            throw new Error(`Jira API 에러 (${res.status}): ${errText}`);
        }
        if (res.status === 204) return true;
        return await res.json();
    }
    throw new Error("서버 응답 지연");
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === "CREATE_NEW_ACCOUNT") {
        handleNewAccount(message.data).then(sendResponse).catch(err => sendResponse({error: err.message}));
        return true; 
    } 
    else if (message.action === "EXTEND_VPN") {
        handleExtendVpn(message.data).then(sendResponse).catch(err => sendResponse({error: err.message}));
        return true;
    }
});

async function handleNewAccount(data) {
    const { systems, ip, usagePeriod, user } = data;
    const userEmail = `${user.jiraId}@sk.com`; // 👈 이메일은 Jira ID 기준
    
    const descriptionText = `B tv 큐레이션/편성 업무 목적 신규 VPN 계정 발급 및 관련 어드민 접속 허용을 요청드립니다.
${user.id} / ${user.name} / ${user.dept} / ${userEmail}`;

    const payload = {
        fields: {
            project: { key: PROJECT_KEY },
            summary: `[신규 신청] ${user.name} - VPN 계정 발급 및 접속지 추가 요청 (${systems.join(', ')})`,
            description: descriptionText,
            issuetype: { name: "Task" }, 
            reporter: { name: user.jiraId },
            assignee: { name: "hs3986" } // 👈 담당자 강제 지정 추가
            
        }
    };
    
    const createRes = await fetchJira('/rest/api/2/issue', 'POST', payload);
    const issueKey = createRes.key;

    const today = new Date().toISOString().split('T')[0];
    const excelAoA = [
        ["VPN(SSL VPN) 작업요청서", "", "", "", "", "", "", ""],
        ["요청일시 및 담당자", "", "", "", "", "", "", ""],
        ["제 목(요청사유)", "", `${systems.join(', ')} 접속을 위한 VPN 신규 생성 및 접속지 추가`, "", "", "", "", ""],
        ["작업 신청일", "", today, "", "작업 구분", "", "신규생성", ""],
        ["작업 요청자", "부  서", user.dept, "", "실 사용자명\n(요청자 동일시 작성 불 필요)", "업체명", "SK브로드밴드", ""],
        ["", "담당자/사번", `${user.name}/${user.id}`, "", "이  름", user.name, "", ""],
        ["", "연락처", "", "", "연락처", "", "", ""],
        ["", "", "", "", "", "", "", ""],
        ["VPN 작업 요청사항", "", "", "", "", "", "", ""],
        ["1. 계정 정보", "", "", "", "", "", "", ""],
        ["계정명\n(사번기준)", "", "패스워드\n(초기 패스워드)", "", "", "사용기간(PJT기간)", "", "비고"],
        [user.id, "", "media123@", "", "", usagePeriod, "", ""],
        ["2. 접속지 추가 / 변경", "", "", "", "", "", "", ""],
        ["Source IP", "", "", "Destination IP", "Service Port\n(TCP/UDP)", "포트용도", "작업구분\n(신규,삭제)", "사용기간"],
        ["접속지 위치", "", "IP Address", "IP Address", "", "", "", ""]
    ];

    systems.forEach(sysName => {
        const target = SYSTEM_DESTINATIONS[sysName];
        if(target) excelAoA.push(["재택근무", "", ip, target.ip, target.port, target.usage, "신규", usagePeriod]);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(excelAoA);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "신청서");
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const excelBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const formData = new FormData();
    formData.append("file", excelBlob, `SSLVPN_신청서_${user.name}.xlsx`);
    
    await fetchJira(`/rest/api/2/issue/${issueKey}/attachments`, 'POST', formData, true);
    await fetchJira(`/rest/api/2/issue/${issueKey}/transitions`, 'POST', { transition: { id: TRANSITION_ID_START_WORK } });

    return { success: true, issueKey: issueKey };
}

async function handleExtendVpn(data) {
    const { date, ip, reason, startTime, endTime, user } = data;
    const [yyyy, mm, dd] = date.split('-'); 
    
    const tableDescription = [
        "반드시 아래 양식에 맞게 입력 부탁드립니다.",
        "아래 양식 이외 신청 건은 반려처리됩니다.",
        "",
        "|구분|SKB 담당자|사용자 소속|사용자 이름|VPN 계정|신청일자(년)|신청일자(월)|신청일자(일)|사용시간(시작)|사용시간(종료)|접속사유|",
        "|---|---|---|---|---|---|---|---|---|---|---|",
        `|1|${user.name}|${user.dept}|${user.name}|${user.id}|${yyyy}|${mm}|${dd}|${startTime}|${endTime}|${reason}|`
    ].join('\n');

    const payload = {
        fields: {
            project: { key: PROJECT_KEY },
            summary: `[활성화 연장] ${user.name} - ${date} VPN 사용 요청`,
            description: tableDescription,
            issuetype: { name: "Task" }, 
            reporter: { name: user.jiraId },
            assignee: { name: "hs3986" } // 👈 담당자 강제 지정 추가
        }
    };

    const createRes = await fetchJira('/rest/api/2/issue', 'POST', payload);
    const issueKey = createRes.key;

    const commentPayload = { body: `[재택 접속 정보 자동 기입]\n해당 인원 재택 근무로 인한 접속 IP 추가 공유합니다.\n*접속 IP:* ${ip}` };
    await fetchJira(`/rest/api/2/issue/${issueKey}/comment`, 'POST', commentPayload);
    await fetchJira(`/rest/api/2/issue/${issueKey}/transitions`, 'POST', { transition: { id: TRANSITION_ID_RECEIPT } });

    return { success: true, issueKey: issueKey };
}