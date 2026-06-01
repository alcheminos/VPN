importScripts('xlsx.min.js');

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
    const { systems, ip, user } = data;
    const userEmail = `${user.jiraId}@sk.com`;
    
    const numericId = user.id.replace(/^skb/i, '');

    const todayObj = new Date();
    const startY = todayObj.getFullYear();
    const startM = String(todayObj.getMonth() + 1).padStart(2, '0');
    const startD = String(todayObj.getDate()).padStart(2, '0');
    const startDate = `${startY}-${startM}-${startD}`;

    const endObj = new Date(todayObj);
    endObj.setFullYear(endObj.getFullYear() + 1);
    endObj.setDate(endObj.getDate() - 1);
    const endY = endObj.getFullYear();
    const endM = String(endObj.getMonth() + 1).padStart(2, '0');
    const endD = String(endObj.getDate()).padStart(2, '0');
    const endDate = `${endY}-${endM}-${endD}`;
    const exactUsagePeriod = `${startDate} ~ ${endDate}`;

    // 💡 systems 배열에서 '기타'인 경우(객체)는 usage 속성 텍스트를 뽑아 이름(문자열) 배열로 변환
    const systemNames = systems.map(s => typeof s === 'object' ? s.usage : s);

    const descriptionText = `B tv 큐레이션/편성 업무 목적 신규 VPN 계정 발급 및 관련 어드민 접속 허용을 요청드립니다.
${user.id} / ${user.name} / ${user.dept} / ${userEmail}`;

    const payload = {
        fields: {
            project: { key: PROJECT_KEY },
            // 💡 제목에 직접 입력한 용도 이름이 포함되도록 수정
            summary: `[신규 신청] ${user.name} - VPN 계정 발급 및 접속지 추가 요청 (${systemNames.join(', ')})`,
            description: descriptionText,
            issuetype: { name: "Task" }, 
            reporter: { name: user.jiraId },
            assignee: { name: "hs3986" } 
        }
    };
    
    const createRes = await fetchJira('/rest/api/2/issue', 'POST', payload);
    const issueKey = createRes.key;

    const excelAoA = [
        ["VPN(SSL VPN) 작업요청서", "", "", "", "", "", "", ""],
        ["요청일시 및 담당자", "", "", "", "", "", "", ""],
        // 💡 엑셀 제목에도 직접 입력한 이름이 포함되도록 수정
        ["제 목(요청사유)", "", `${systemNames.join(', ')} 접속을 위한 VPN 신규 생성 및 접속지 추가`, "", "", "", "", ""],
        ["작업 신청일", "", startDate, "", "작업 구분", "", "신규생성", ""],
        ["작업 요청자", "부  서", user.dept, "", "실 사용자명\n(요청자 동일시 작성 불 필요)", "업체명", "SK브로드밴드", ""],
        ["", "담당자/사번", `${user.name} / ${numericId}`, "", "", "이  름", "", ""], 
        ["", "연락처", user.phone, "", "", "연락처", "", ""], 
        ["", "", "", "", "", "", "", ""],
        ["VPN 작업 요청사항", "", "", "", "", "", "", ""],
        ["1. 계정 정보", "", "", "", "", "", "", ""],
        ["계정명\n(사번기준)", "", "패스워드\n(초기 패스워드)", "", "", "사용기간(PJT기간)", "", "비고"],
        [user.id, "", "media123@", "", "", exactUsagePeriod, "", ""], 
        ["2. 접속지 추가 / 변경", "", "", "", "", "", "", ""],
        ["Source IP", "", "Destination IP", "", "Service Port\n(TCP/UDP)", "포트용도", "작업구분\n(신규,삭제)", "사용기간"],
        ["접속지 위치", "IP Address", "IP Address", "", "", "", "", ""] 
    ];

    let currentRow = 15;
    systems.forEach(sysItem => {
        let target;
        // 💡 넘어온 데이터가 객체(기타)인 경우, 사용자가 입력한 IP, Port, 용도를 그대로 할당
        if (typeof sysItem === 'object' && sysItem.type === 'other') {
            target = { ip: sysItem.ip, port: sysItem.port, usage: sysItem.usage };
        } else {
            // 기존 고정 시스템인 경우
            target = SYSTEM_DESTINATIONS[sysItem];
        }
        
        // 엑셀 양식의 각 열(Column)에 맞게 데이터 배치
        excelAoA.push(["재택근무", ip, target.ip, "", target.port, target.usage, "신규", "1년"]); 
        currentRow++;
    });

    const worksheet = XLSX.utils.aoa_to_sheet(excelAoA);

    worksheet['!cols'] = [
        {wch: 15}, {wch: 15}, {wch: 18}, {wch: 20}, {wch: 18}, {wch: 15}, {wch: 12}, {wch: 25}
    ];

    worksheet['!merges'] = [
        { s: {r:0, c:0}, e: {r:0, c:7} }, 
        { s: {r:1, c:0}, e: {r:1, c:7} }, 
        { s: {r:2, c:0}, e: {r:2, c:1} }, { s: {r:2, c:2}, e: {r:2, c:7} }, 
        { s: {r:3, c:0}, e: {r:3, c:1} }, { s: {r:3, c:2}, e: {r:3, c:3} }, { s: {r:3, c:4}, e: {r:3, c:5} }, { s: {r:3, c:6}, e: {r:3, c:7} }, 
        { s: {r:4, c:0}, e: {r:6, c:0} }, 
        { s: {r:4, c:4}, e: {r:6, c:4} }, 
        { s: {r:4, c:2}, e: {r:4, c:3} }, { s: {r:4, c:6}, e: {r:4, c:7} }, 
        { s: {r:5, c:2}, e: {r:5, c:3} }, { s: {r:5, c:6}, e: {r:5, c:7} }, 
        { s: {r:6, c:2}, e: {r:6, c:3} }, { s: {r:6, c:6}, e: {r:6, c:7} }, 
        { s: {r:8, c:0}, e: {r:8, c:7} }, 
        { s: {r:9, c:0}, e: {r:9, c:7} }, 
        { s: {r:10, c:0}, e: {r:10, c:1} }, { s: {r:10, c:2}, e: {r:10, c:4} }, { s: {r:10, c:5}, e: {r:10, c:6} }, 
        { s: {r:11, c:0}, e: {r:11, c:1} }, { s: {r:11, c:2}, e: {r:11, c:4} }, { s: {r:11, c:5}, e: {r:11, c:6} }, 
        { s: {r:12, c:0}, e: {r:12, c:7} }, 
        { s: {r:13, c:0}, e: {r:13, c:1} }, 
        { s: {r:13, c:2}, e: {r:13, c:3} }, 
        { s: {r:13, c:4}, e: {r:14, c:4} }, 
        { s: {r:13, c:5}, e: {r:14, c:5} }, 
        { s: {r:13, c:6}, e: {r:14, c:6} }, 
        { s: {r:13, c:7}, e: {r:14, c:7} }, 
        { s: {r:14, c:2}, e: {r:14, c:3} }  
    ];

    let mergeRow = 15;
    systems.forEach(() => {
        worksheet['!merges'].push({ s: {r:mergeRow, c:2}, e: {r:mergeRow, c:3} });
        mergeRow++;
    });

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = XLSX.utils.encode_cell({c: C, r: R});
            if (!worksheet[cellAddress]) continue;

            if ([0, 9, 12].includes(R)) {
                worksheet[cellAddress].s = { 
                    font: { bold: true, sz: (R === 0 ? 16 : 11) }, 
                    alignment: { vertical: "center", horizontal: (R === 0 ? "center" : "left") }
                };
                continue;
            }
            if (R === 7) continue; 

            let cellStyle = {
                alignment: { horizontal: "center", vertical: "center", wrapText: true },
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                }
            };

            const isGray = 
                (R === 1) || 
                (R === 8) || 
                (R === 2 && C === 0) || 
                (R === 3 && (C === 0 || C === 4)) || 
                (R === 4 && (C === 0 || C === 1 || C === 4 || C === 5)) || 
                (R === 5 && (C === 1 || C === 5)) || 
                (R === 6 && (C === 1 || C === 5)) || 
                (R === 10 && (C === 0 || C === 2 || C === 5 || C === 7)) || 
                (R === 13) || (R === 14);

            if (isGray) {
                cellStyle.fill = { fgColor: { rgb: "FFF2F2F2" } };
                cellStyle.font = { bold: true };
            }

            worksheet[cellAddress].s = cellStyle;
        }
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "신청서");
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const excelBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const formData = new FormData();
    formData.append("file", excelBlob, `SSLVPN_신청서_${user.name}.xlsx`);
    
    await fetchJira(`/rest/api/2/issue/${issueKey}/attachments`, 'POST', formData, true);
    await fetchJira(`/rest/api/2/issue/${issueKey}/transitions`, 'POST', { transition: { id: TRANSITION_ID_RECEIPT } });

    return { success: true, issueKey: issueKey };
}
