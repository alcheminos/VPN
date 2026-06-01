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
    const userEmail = `${user.jiraId}@sk.com`;
    
    const descriptionText = `B tv 큐레이션/편성 업무 목적 신규 VPN 계정 발급 및 관련 어드민 접속 허용을 요청드립니다.
${user.id} / ${user.name} / ${user.dept} / ${userEmail}`;

    const payload = {
        fields: {
            project: { key: PROJECT_KEY },
            summary: `[신규 신청] ${user.name} - VPN 계정 발급 및 접속지 추가 요청 (${systems.join(', ')})`,
            description: descriptionText,
            issuetype: { name: "Task" }, 
            reporter: { name: user.jiraId },
            assignee: { name: "hs3986" } 
        }
    };
    
    const createRes = await fetchJira('/rest/api/2/issue', 'POST', payload);
    const issueKey = createRes.key;

    const today = new Date().toISOString().split('T')[0];
    
    // CSV 구조와 동일한 2D 배열 생성
    const excelAoA = [
        ["VPN(SSL VPN) 작업요청서", "", "", "", "", "", "", ""],
        ["요청일시 및 담당자", "", "", "", "", "", "", ""],
        ["제 목(요청사유)", "", `${systems.join(', ')} 접속을 위한 VPN 신규 생성 및 접속지 추가`, "", "", "", "", ""],
        ["작업 신청일", "", today, "", "작업 구분", "", "신규생성", ""],
        ["작업 요청자", "부  서", user.dept, "", "실 사용자명\n(요청자 동일시 작성 불 필요)", "업체명", "SK브로드밴드", ""],
        ["", "담당자/사번", `${user.name}/${user.id}`, "", "이  름", user.name, "", ""],
        ["", "연락처", "", "", "연락처", "", "", ""],
        ["", "", "", "", "", "", "", ""], // 빈줄 (8행)
        ["VPN 작업 요청사항", "", "", "", "", "", "", ""], // 9행
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

    // 열 너비 지정
    worksheet['!cols'] = [
        {wch: 15}, {wch: 15}, {wch: 18}, {wch: 20}, {wch: 18}, {wch: 15}, {wch: 12}, {wch: 15}
    ];

    // 셀 병합(Merge) - 구조 완벽 일치 적용
    worksheet['!merges'] = [
        { s: {r:0, c:0}, e: {r:0, c:7} }, // 타이틀
        { s: {r:1, c:0}, e: {r:1, c:7} }, // 요청일시
        { s: {r:2, c:0}, e: {r:2, c:1} }, { s: {r:2, c:2}, e: {r:2, c:7} }, // 제목
        { s: {r:3, c:0}, e: {r:3, c:1} }, { s: {r:3, c:2}, e: {r:3, c:3} }, { s: {r:3, c:4}, e: {r:3, c:5} }, { s: {r:3, c:6}, e: {r:3, c:7} }, // 신청일/구분
        { s: {r:4, c:0}, e: {r:7, c:0} }, // 작업 요청자(세로 병합)
        { s: {r:4, c:4}, e: {r:7, c:4} }, // 실 사용자명(세로 병합)
        { s: {r:4, c:2}, e: {r:4, c:3} }, { s: {r:4, c:6}, e: {r:4, c:7} }, // 부서, 업체명
        { s: {r:5, c:2}, e: {r:5, c:3} }, { s: {r:5, c:6}, e: {r:5, c:7} }, // 담당자, 이름
        { s: {r:6, c:2}, e: {r:6, c:3} }, { s: {r:6, c:6}, e: {r:6, c:7} }, // 연락처
        { s: {r:8, c:0}, e: {r:8, c:7} }, // 작업 요청사항
        { s: {r:9, c:0}, e: {r:9, c:7} }, // 계정 정보
        { s: {r:10, c:0}, e: {r:10, c:1} }, { s: {r:10, c:2}, e: {r:10, c:4} }, { s: {r:10, c:5}, e: {r:10, c:6} }, 
        { s: {r:11, c:0}, e: {r:11, c:1} }, { s: {r:11, c:2}, e: {r:11, c:4} }, { s: {r:11, c:5}, e: {r:11, c:6} }, 
        { s: {r:12, c:0}, e: {r:12, c:7} }, // 접속지 추가
        { s: {r:13, c:0}, e: {r:14, c:1} }, // Source IP
        { s: {r:13, c:2}, e: {r:14, c:2} }, { s: {r:13, c:3}, e: {r:14, c:3} }, { s: {r:13, c:4}, e: {r:14, c:4} },
        { s: {r:13, c:5}, e: {r:14, c:5} }, { s: {r:13, c:6}, e: {r:14, c:6} }, { s: {r:13, c:7}, e: {r:14, c:7} }
    ];

    let currentRow = 15;
    systems.forEach(() => {
        worksheet['!merges'].push({ s: {r:currentRow, c:0}, e: {r:currentRow, c:1} });
        currentRow++;
    });

    // --- 스타일 적용 로직 시작 (xlsx-js-style 전용) ---
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    
    // 기본 테두리 및 정렬 스타일
    const baseStyle = {
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
        }
    };
    
    // 회색 배경 및 굵은 글씨 (헤더용)
    const headerStyle = {
        ...baseStyle,
        fill: { fgColor: { rgb: "FFF2F2F2" } },
        font: { bold: true }
    };

    // 스타일 맵핑 (회색 배경이 들어가야 하는 주요 헤더 셀들의 좌표)
    const grayCells = [
        "A3", "A4", "E4", "A5", "B5", "E5", "F5", "B6", "F6", "B7", "F7",
        "A11", "C11", "F11", "H11", "A14", "D14", "E14", "F14", "G14", "H14", "C15", "D15"
    ];

    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = XLSX.utils.encode_cell({c: C, r: R});
            if (!worksheet[cellAddress]) continue;

            // 타이틀 부분(1,2,9,10,13행)은 테두리 없이 굵게
            if ([0, 1, 8, 9, 12].includes(R)) {
                worksheet[cellAddress].s = { font: { bold: true, sz: (R === 0 ? 14 : 11) }, alignment: { vertical: "center" } };
                continue;
            }
            
            // 빈 셀(8행 등) 처리
            if (R === 7) continue;

            // 지정된 셀은 회색 헤더 스타일, 나머지는 기본(흰색) 스타일 적용
            if (grayCells.includes(cellAddress)) {
                worksheet[cellAddress].s = headerStyle;
            } else {
                worksheet[cellAddress].s = baseStyle;
            }
        }
    }
    // --- 스타일 적용 로직 끝 ---

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
