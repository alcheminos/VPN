importScripts('xlsx.min.js');

const JIRA_BASE_URL = "https://jira.skbroadband.com";
const PROJECT_KEY = "BTVMKT"; 
const TRANSITION_ID_RECEIPT = "4"; 

const SYSTEM_DESTINATIONS = {
    "EUXP 상용": { ip: "1.255.152.40", port: "TCP 8080, 8443", usage: "EUXP" },
    "EUXP Pre": { ip: "1.255.128.22", port: "TCP 8443", usage: "EUXP" },
    "EUXP Stg": { ip: "1.255.86.188", port: "TCP 8443", usage: "EUXP" },
    "NCMS 상용": { ip: "1.255.152.19", port: "TCP 8001", usage: "NCMS" },
    "NCMS Pre": { ip: "1.255.128.37", port: "TCP 8443", usage: "NCMS" },
    "NCMS Stg": { ip: "175.113.150.43", port: "TCP 8443", usage: "NCMS" },
    "ACS": { ip: "114.202.130.191\n114.202.130.40", port: "TCP 9093", usage: "ACS" },
    "통합ES": { ip: "1.255.140.10", port: "TCP 5601", usage: "KIBANA" },
    "수유 빌드 키바나": { ip: "1.255.152.46", port: "TCP 5601", usage: "KIBANA" },
    "성수 빌드 키바나": { ip: "1.255.152.174", port: "TCP 5601", usage: "KIBANA" },
    "metainfo 키바나": { ip: "116.126.69.77", port: "TCP 5601", usage: "KIBANA" },
    // 💡 미디어디스커버리는 NDM으로 변경
    "미디어디스커버리 1": { ip: "221.140.123.144", port: "TCP 8080, 7070", usage: "NDM" },
    "미디어디스커버리 2": { ip: "221.140.123.143", port: "TCP 8080", usage: "NDM" },
    "미디어디스커버리 3": { ip: "221.140.123.78", port: "TCP 8080", usage: "NDM" },
    "미디어디스커버리 4": { ip: "1.255.113.183", port: "TCP 8080", usage: "NDM" },
    "추천 Admin": { ip: "1.255.152.46", port: "TCP 5630", usage: "NDM" },
    "ECDN": { ip: "121.125.63.21", port: "TCP 80", usage: "ETC" },
    "상용DB": { ip: "114.202.130.168", port: "TCP 9090", usage: "ETC" }
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
    
    // 사번에서 'skb'를 제거하여 숫자만 추출
    const numericId = user.id.replace(/^skb/i, '');

    // 오늘부터 1년(내년 전일) 날짜 자동 계산
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

    // 시스템 이름 배열 (객체로 넘어온 기타 항목 처리)
    const systemNames = systems.map(s => typeof s === 'object' ? s.usage : s);

    const descriptionText = `B tv 큐레이션/편성 업무 목적 신규 VPN 계정 발급 및 관련 어드민 접속 허용을 요청드립니다.
${user.id} / ${user.name} / ${user.dept} / ${userEmail}`;

    const payload = {
        fields: {
            project: { key: PROJECT_KEY },
            // 💡 3) Jira 제목 변경: VPN 발급 및 접속지 추가 요청 (이름)
            summary: `[신규 신청] VPN 발급 및 접속지 추가 요청 (${user.name})`,
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
        ["제 목(요청사유)", "", "업무 수행을 위한 VPN 신규 생성 및 접속지 추가", "", "", "", "", ""],
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
        // '기타'로 사용자가 직접 입력한 객체가 들어왔을 경우 처리
        if (typeof sysItem === 'object' && sysItem.type === 'other') {
            target = { ip: sysItem.ip, port: sysItem.port, usage: sysItem.usage };
        } else {
            target = SYSTEM_DESTINATIONS[sysItem];
        }
        
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

// 👇 제가 실수로 누락했던 handleExtendVpn 함수입니다!
async function handleExtendVpn(data) {
    const { date, ip, reason, startTime, endTime, user } = data;
    const [yyyy, mm, dd] = date.split('-'); 
    
    const tableDescription = `
<p>반드시 아래 양식에 맞게 입력 부탁드립니다.<br>
아래 양식 이외 신청 건은 반려처리됩니다.</p>
<table border="1" style="border-collapse: collapse; width: 100%; text-align: center;">
    <tbody>
        <tr style="background-color: #f4f5f7;">
            <td rowspan="2" style="padding: 5px;">구분</td>
            <td rowspan="2" style="padding: 5px;">SKB 담당자</td>
            <td rowspan="2" style="padding: 5px;">사용자 소속</td>
            <td rowspan="2" style="padding: 5px;">사용자 이름</td>
            <td rowspan="2" style="padding: 5px;">VPN 계정</td>
            <td colspan="3" style="padding: 5px;">신청일자</td>
            <td colspan="2" style="padding: 5px;">사용시간</td>
            <td rowspan="2" style="padding: 5px;">접속사유</td>
        </tr>
        <tr style="background-color: #f4f5f7;">
            <td style="padding: 5px;">년</td>
            <td style="padding: 5px;">월</td>
            <td style="padding: 5px;">일</td>
            <td style="padding: 5px;">시작</td>
            <td style="padding: 5px;">종료</td>
        </tr>
        <tr>
            <td style="padding: 5px;">1</td>
            <td style="padding: 5px;">${user.name}</td>
            <td style="padding: 5px;">${user.dept}</td>
            <td style="padding: 5px;">${user.name}</td>
            <td style="padding: 5px;">${user.id}</td>
            <td style="padding: 5px;">${yyyy}</td>
            <td style="padding: 5px;">${mm}</td>
            <td style="padding: 5px;">${dd}</td>
            <td style="padding: 5px;">${startTime}</td>
            <td style="padding: 5px;">${endTime}</td>
            <td style="padding: 5px;">${reason}</td>
        </tr>
    </tbody>
</table>
`;

    const payload = {
        fields: {
            project: { key: PROJECT_KEY },
            // 💡 4) 활성화 제목 변경: [활성화] 날짜 VPN 사용 요청
            summary: `[활성화] ${date} VPN 사용 요청`,
            description: tableDescription,
            issuetype: { name: "Task" }, 
            reporter: { name: user.jiraId },
            assignee: { name: "hs3986" }
        }
    };

    const createRes = await fetchJira('/rest/api/2/issue', 'POST', payload);
    const issueKey = createRes.key;

    const commentPayload = { body: `[재택 접속 정보 자동 기입]\n해당 인원 재택 근무로 인한 접속 IP 추가 공유합니다.\n*접속 IP:* ${ip}` };
    await fetchJira(`/rest/api/2/issue/${issueKey}/comment`, 'POST', commentPayload);
    await fetchJira(`/rest/api/2/issue/${issueKey}/transitions`, 'POST', { transition: { id: TRANSITION_ID_RECEIPT } });

    return { success: true, issueKey: issueKey };
}
