[시간표검증기_42_1.html](https://github.com/user-attachments/files/31054603/_42_1.html)
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>수업 시간표 검증기</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=DM+Mono:wght@400;500&display=swap');
:root{
  --bg:#F7F6F2;--surface:#fff;--surface2:#F0EFE9;--border:#E2E0D8;
  --text:#1A1916;--text2:#6B6860;--text3:#9B9890;
  --accent:#2D5A3D;--accent-light:#E8F0EA;
  --danger:#C0392B;--danger-light:#FBEAE8;
  --warn:#9C6A00;--warn-light:#FDF5E0;
  --ok:#1E6B3A;--ok-light:#E8F5EC;
  --info:#1A4A7A;--info-light:#E8EFF8;
  --mono:'DM Mono',monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Noto Sans KR',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;line-height:1.6}
header{background:var(--accent);padding:20px 32px;display:flex;align-items:center;gap:16px}
.logo{width:36px;height:36px;background:rgba(255,255,255,.15);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px}
header h1{font-size:17px;font-weight:500;color:#fff;letter-spacing:-.3px}
header p{font-size:12px;color:rgba(255,255,255,.6);margin-top:1px}
.container{max-width:1000px;margin:0 auto;padding:28px 20px}
.upload-card{background:var(--surface);border:1.5px dashed var(--border);border-radius:14px;padding:48px 32px;text-align:center;cursor:pointer;transition:border-color .2s,background .2s;margin-bottom:24px}
.upload-card:hover,.upload-card.drag{border-color:var(--accent);background:var(--accent-light)}
.upload-icon{width:56px;height:56px;background:var(--surface2);border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:26px;margin:0 auto 16px}
.upload-card h3{font-size:15px;font-weight:500;margin-bottom:6px}
.upload-card p{font-size:13px;color:var(--text2)}
.upload-card input{display:none}
.date-info{background:var(--info-light);border:.5px solid var(--info);border-radius:10px;padding:10px 16px;margin-bottom:16px;font-size:13px;color:var(--info);display:none}
.stats{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:24px}
.stat{background:var(--surface);border:.5px solid var(--border);border-radius:12px;padding:14px 16px}
.stat-label{font-size:10px;color:var(--text3);font-weight:500;letter-spacing:.5px;text-transform:uppercase;margin-bottom:6px}
.stat-value{font-size:22px;font-weight:700;color:var(--text);font-family:var(--mono)}
.stat-value.danger{color:var(--danger)}
.stat-value.warn{color:var(--warn)}
.stat-value.ok{color:var(--ok)}
.tabs{display:flex;gap:4px;background:var(--surface2);border-radius:10px;padding:4px;margin-bottom:20px;flex-wrap:wrap}
.tab{padding:7px 16px;border-radius:8px;border:none;font-family:'Noto Sans KR',sans-serif;font-size:13px;font-weight:400;color:var(--text2);cursor:pointer;background:transparent;transition:all .15s}
.tab.active{background:var(--surface);color:var(--text);font-weight:500;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.section{display:none}
.section.active{display:block}
.issue-card{background:var(--surface);border:.5px solid var(--border);border-radius:10px;padding:14px 16px;margin-bottom:10px}
.issue-card.error{border-left:3px solid var(--danger)}
.issue-card.warn-card{border-left:3px solid var(--warn)}
.issue-card.ok-card{border-left:3px solid var(--ok)}
.issue-title{font-size:13px;font-weight:500;color:var(--text);margin-bottom:4px;display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.issue-desc{font-size:12px;color:var(--text2);margin-top:8px;line-height:1.8}
.time-range{font-family:var(--mono);font-size:12px;background:var(--surface2);padding:2px 8px;border-radius:6px}
.amount-box{font-family:var(--mono);font-size:12px;padding:2px 8px;border-radius:6px;display:inline-block}
.amount-expect{background:var(--ok-light);color:var(--ok)}
.amount-actual{background:var(--danger-light);color:var(--danger)}
.badge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:500}
.badge-error{background:var(--danger-light);color:var(--danger)}
.badge-ok{background:var(--ok-light);color:var(--ok)}
.badge-warn{background:var(--warn-light);color:var(--warn)}
.badge-info{background:var(--info-light);color:var(--info)}
.badge-보강{background:#FDF5E0;color:#9C6A00}
.chip{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:500}
.chip-심리{background:#E8EFF8;color:#1A4A7A}
.chip-정서{background:#EEEDFE;color:#534AB7}
.chip-비전{background:#E1F5EE;color:#0F6E56}
.chip-오감{background:#FDF5E0;color:#9C6A00}
.chip-발달{background:#FAECE7;color:#993C1D}
.chip-해피{background:#FBEAF0;color:#993356}
.chip-과몰입{background:#F5E6F8;color:#7B2FA0}
.chip-합주{background:#EFEDE6;color:#6B6860}
.stat-link{cursor:pointer;transition:transform .1s,box-shadow .1s;}
.stat-link:hover{transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,.1);border-color:var(--accent)!important;}
.teacher-block{margin-bottom:20px}
.teacher-name{font-size:12px;font-weight:500;color:var(--text2);text-transform:uppercase;letter-spacing:.5px;padding-bottom:8px;margin-bottom:8px;border-bottom:.5px solid var(--border);display:flex;align-items:center;gap:8px}
.filter-row{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:center}
.filter-row label{font-size:12px;color:var(--text2)}
.filter-row select{font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:6px 12px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);cursor:pointer;outline:none}
.table-wrap{overflow-x:auto;border-radius:10px;border:.5px solid var(--border)}
table{width:100%;border-collapse:collapse;font-size:13px;background:var(--surface)}
thead{background:var(--surface2)}
th{padding:10px 14px;text-align:left;font-size:11px;font-weight:500;color:var(--text2);letter-spacing:.3px;border-bottom:.5px solid var(--border);white-space:nowrap}
td{padding:10px 14px;border-bottom:.5px solid var(--border);color:var(--text);vertical-align:middle}
tr:last-child td{border-bottom:none}
.mono{font-family:var(--mono);font-size:12px}
.empty{text-align:center;padding:40px 20px;color:var(--text3);font-size:13px}
.legend{display:flex;gap:16px;flex-wrap:wrap;font-size:12px;color:var(--text2);padding:12px 16px;background:var(--surface2);border-radius:8px;margin-bottom:16px;align-items:center}
.section-title{font-size:13px;font-weight:500;color:var(--text2);margin:16px 0 8px;padding-bottom:6px;border-bottom:.5px solid var(--border)}
@media(max-width:700px){.stats{grid-template-columns:repeat(3,1fr)}}
@media(max-width:500px){.stats{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<header>
  <div class="logo">📋</div>
  <div>
    <h1>수업 시간표 검증기</h1>
    <p>결제내역 엑셀 파일로 시간 겹침 · 결제 누락 · 금액 오류를 자동 확인합니다</p>
  </div>
</header>

<!-- 시간표 설정 패널 -->
<div class="container" style="padding-bottom:0">
  <div style="background:var(--surface);border:.5px solid var(--border);border-radius:14px;padding:16px 20px;margin-bottom:12px;">
    <div style="display:flex;align-items:center;justify-content:space-between;cursor:pointer;" onclick="togglePdfPanel()">
      <div style="display:flex;align-items:center;gap:8px;">
        <span style="font-size:15px;">📋</span>
        <span style="font-size:14px;font-weight:500;color:var(--text)">시간표 설정</span>
        <span id="scheduleStatus" style="font-size:12px;color:var(--text2)">— 스프레드시트에서 엑셀을 다운받아 업로드 해주세요</span>
      </div>
      <span id="pdfToggleIcon" style="font-size:12px;color:var(--text3)">▼ 펼치기</span>
    </div>
    <div id="pdfPanel" style="display:none;margin-top:14px;border-top:.5px solid var(--border);padding-top:14px;">
      <div style="display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap;align-items:center;">
        <label style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:7px 16px;border-radius:8px;border:.5px solid var(--accent);background:var(--accent);color:#fff;cursor:pointer;">
          📊 엑셀 시간표 업로드
          <input type="file" id="schedXlsxInput" accept=".xlsx,.xls" style="display:none" onchange="handleScheduleXlsx(event)">
        </label>
        <label style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:7px 16px;border-radius:8px;border:.5px solid var(--border);background:transparent;color:var(--text2);cursor:pointer;">
          📎 PDF 업로드
          <input type="file" id="pdfInput" accept=".pdf" style="display:none" onchange="handlePdf(event)">
        </label>
        <button onclick="clearSchedule()" style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:7px 16px;border-radius:8px;border:.5px solid var(--border);background:transparent;color:var(--danger);cursor:pointer;">🗑 저장된 시간표 삭제</button>
        <span id="pdfStatusMsg" style="font-size:12px;color:var(--text2);"></span>
      </div>
      <div id="schedulePreview" style="display:none;">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:8px;">
          <span style="font-size:13px;font-weight:500;color:var(--text)">파싱 결과 확인</span>
          <div style="display:flex;gap:8px;">
            <select id="previewDow" style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:5px 10px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);" onchange="renderPreview()">
              <option value="1">월요일</option>
              <option value="2">화요일</option>
              <option value="3">수요일</option>
              <option value="4">목요일</option>
              <option value="5" selected>금요일</option>
              <option value="6">토요일</option>
            </select>
            <button onclick="saveSchedule()" style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:7px 16px;border-radius:8px;border:.5px solid var(--accent);background:var(--accent);color:#fff;cursor:pointer;">💾 저장</button>
          </div>
        </div>
        <div id="previewTable" style="overflow-x:auto;border-radius:10px;border:.5px solid var(--border);max-height:300px;overflow-y:auto;"></div>
      </div>
    </div>
  </div>
</div>

<!-- 부재 설정 패널 -->
<div class="container" style="padding-bottom:0">
  <div style="background:var(--surface);border:.5px solid var(--border);border-radius:14px;padding:16px 20px;margin-bottom:12px;">
    <div style="display:flex;align-items:center;justify-content:space-between;cursor:pointer;" onclick="toggleAbsence()">
      <div style="display:flex;align-items:center;gap:8px;">
        <span style="font-size:15px;">🚫</span>
        <span style="font-size:14px;font-weight:500;color:var(--text)">부재 설정</span>
        <span id="absenceSummary" style="font-size:12px;color:var(--text2)">— 부재 중 결제되면 <span style="color:var(--danger);font-weight:500;">과오결제</span>로 표시됩니다</span>
      </div>
      <span id="absenceToggleIcon" style="font-size:12px;color:var(--text3)">▼ 펼치기</span>
    </div>
    <div id="absencePanel" style="display:none;margin-top:14px;border-top:.5px solid var(--border);padding-top:14px;">
      <div style="display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:flex-end;padding:14px;background:var(--surface2);border-radius:10px;">
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px;">구분</div>
          <select id="absenceType" style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:6px 12px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);">
            <option value="선생님">선생님</option>
            <option value="학생">학생</option>
          </select>
        </div>
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px;">이름</div>
          <input id="absenceName" type="text" placeholder="이름 입력"
            style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:6px 12px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);width:120px;"
            onkeydown="if(event.key==='Enter')addAbsence()">
        </div>
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px;">사유 (선택)</div>
          <input id="absenceReason" type="text" placeholder="예: 해외, 입원"
            style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:6px 12px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);width:110px;"
            onkeydown="if(event.key==='Enter')addAbsence()">
        </div>
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px;">시작일</div>
          <input id="absenceStart" type="date"
            style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:6px 10px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);cursor:pointer;"
            oninput="onAbsenceStartChange()">
        </div>
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px;">종료일</div>
          <input id="absenceEnd" type="date"
            style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:6px 10px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);cursor:pointer;">
        </div>
        <button onclick="addAbsence()"
          style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:7px 16px;border-radius:8px;border:.5px solid var(--accent);background:var(--accent);color:#fff;cursor:pointer;margin-bottom:1px;">
          + 추가
        </button>
      </div>
      <div id="absenceList">
        <span style="font-size:12px;color:var(--text3);">부재 목록이 없습니다</span>
      </div>
    </div>
  </div>
</div>
<div class="container">
  <label for="fileInput" class="upload-card" id="dropZone" style="display:block;cursor:pointer;">
    <input type="file" id="fileInput" accept=".xlsx,.xls" onchange="handleFile(event)" style="display:none;position:absolute;">
    <div class="upload-icon">📊</div>
    <h3>결제내역 엑셀 파일 업로드</h3>
    <p>클릭하거나 파일을 여기로 드래그하세요 · .xlsx / .xls 형식</p>
  </label>

  <div style="background:var(--surface);border:.5px solid var(--border);border-radius:14px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
    <span style="font-size:13px;color:var(--text2);">📅 검증 기준 날짜</span>
    <input type="date" id="targetDate" style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:5px 10px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);cursor:pointer;"
      onchange="onDateChange()">
    <span id="dowBadge" style="font-size:12px;font-weight:500;padding:3px 10px;border-radius:20px;background:var(--info-light);color:var(--info);"></span>
    <div style="display:flex;gap:0;border:.5px solid var(--border);border-radius:8px;overflow:hidden;margin-left:4px;">
      <button id="modeBtn진행중" onclick="setMode('진행중')"
        style="font-family:'Noto Sans KR',sans-serif;font-size:12px;padding:5px 14px;border:none;cursor:pointer;background:var(--accent);color:#fff;font-weight:500;">
        ⏳ 진행 중
      </button>
      <button id="modeBtn완료" onclick="setMode('완료')"
        style="font-family:'Noto Sans KR',sans-serif;font-size:12px;padding:5px 14px;border:none;cursor:pointer;background:var(--surface2);color:var(--text2);">
        ✅ 결제 완료
      </button>
    </div>
    <span style="font-size:12px;color:var(--text3);">진행 중 → 미결제는 결제 예정 · 완료 → 미결제는 누락</span>
  </div>

  <div id="results" style="display:none">
    <div class="stats">
      <div class="stat"><div class="stat-label">총 결제건</div><div class="stat-value" id="totalCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('overlap')"><div class="stat-label">시간 겹침</div><div class="stat-value danger" id="overlapCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('missing')"><div class="stat-label">결제 누락</div><div class="stat-value danger" id="missingCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('pending')"><div class="stat-label">결제 예정</div><div class="stat-value warn" id="pendingCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('amount')"><div class="stat-label">금액 오류</div><div class="stat-value danger" id="amountErrCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('paymismatch')"><div class="stat-label">결제구분 불일치</div><div class="stat-value danger" id="payMismatchCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('wrong')"><div class="stat-label">과오결제</div><div class="stat-value danger" id="wrongCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('dup')"><div class="stat-label">중복결제</div><div class="stat-value danger" id="dupCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('delay')"><div class="stat-label">지연결제</div><div class="stat-value warn" id="delayCount">—</div></div>
      <div class="stat stat-link" onclick="goTab('missing')"><div class="stat-label">시간표 외 결제</div><div class="stat-value warn" id="extraCount">—</div></div>
      <div class="stat"><div class="stat-label">정상</div><div class="stat-value ok" id="okCount">—</div></div>
    </div>

    <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;flex-wrap:wrap;">
      <span style="font-size:13px;color:var(--text2);">⏱ 지연 기준</span>
      <input type="number" id="delayThreshold" value="10" min="1" max="60"
        style="font-family:'Noto Sans KR',sans-serif;font-size:13px;padding:5px 10px;border-radius:8px;border:.5px solid var(--border);background:var(--surface);color:var(--text);width:70px;text-align:center;"
        onchange="if(allData.length)rerun()">
      <span style="font-size:13px;color:var(--text2);">분 이상이면 지연결제</span>
    </div>

    <div class="tabs">
      <button class="tab active" data-tab="overlap">⚠ 시간 겹침</button>
      <button class="tab" data-tab="missing">❌ 결제 누락</button>
      <button class="tab" data-tab="pending">🕐 결제 예정</button>
      <button class="tab" data-tab="amount">💰 금액 오류</button>
      <button class="tab" data-tab="paymismatch">🔄 결제구분 불일치</button>
      <button class="tab" data-tab="wrong">🚫 과오결제</button>
      <button class="tab" data-tab="dup">🔁 중복결제</button>
      <button class="tab" data-tab="delay">⏱ 지연결제</button>
      <button class="tab" data-tab="schedule">👩‍🏫 선생님별 시간표</button>
      <button class="tab" data-tab="all">📋 전체 목록</button>
    </div>
    <div id="tab-overlap" class="section active"><div id="overlapContent"></div></div>
    <div id="tab-missing" class="section"><div id="missingContent"></div></div>
    <div id="tab-pending" class="section"><div id="pendingContent"></div></div>
    <div id="tab-amount" class="section"><div id="amountContent"></div></div>
    <div id="tab-paymismatch" class="section"><div id="payMismatchContent"></div></div>
    <div id="tab-wrong" class="section"><div id="wrongContent"></div></div>
    <div id="tab-dup" class="section"><div id="dupContent"></div></div>
    <div id="tab-delay" class="section"><div id="delayContent"></div></div>
    <div id="tab-schedule" class="section">
      <div class="filter-row">
        <label>선생님</label>
        <select id="teacherFilter" onchange="renderSchedule()"><option value="">전체</option></select>
      </div>
      <div id="scheduleContent"></div>
    </div>
    <div id="tab-all" class="section">
      <div class="filter-row">
        <label>유형</label>
        <select id="typeFilter" onchange="renderAll()"><option value="">전체</option></select>
        <label>선생님</label>
        <select id="teacherFilter2" onchange="renderAll()"><option value="">전체</option></select>
        <label>결제구분</label>
        <select id="payFilter" onchange="renderAll()">
          <option value="">전체</option>
          <option value="정상결제">정상결제</option>
          <option value="보강결제">보강결제</option>
        </select>
      </div>
      <div id="allContent"></div>
    </div>
  </div>
</div>

<script>
/* ============================================================================
   ★★★ 규칙 설정 (여기만 수정하면 됩니다) ★★★
   ============================================================================

   【규칙 1】 수업 유형별 기본 수업시간 → 아래 SERVICE_MAP의 mins 값
   【규칙 2】 합주 수업시간              → RULE_합주분 (기본 20분)
              셀에 "합주" + "보강"이 같이 있으면 유형 무관(정서 포함) 합주 처리
   【규칙 3】 과몰입+부모상담+보강결제   → 합주와 동일하게 20분 처리 (코드 자동)
   【규칙 4】 금액 해석: 셀의 숫자가 100 이상이면 ×100 (405→40,500)
              100 미만이면 ×1000 (36→36,000)
   【규칙 5】 시간 해석: 시(時)가 10 미만이면 오후로 간주 (+12시간)
              예) "7시" → 19:00, "3시30분" → 15:30. "*" 붙은 시간이 최우선
   【규칙 6】 어두운 배경 셀(밝기<RULE_어두운셀밝기)은 파싱에서 제외 (삭제 표시용)
   【규칙 7】 같은 선생님·유형·합주여부 결제가 RULE_그룹묶음분 이내 간격이면
              한 수업 그룹으로 묶음
   【규칙 8】 시간 겹침이 RULE_겹침허용분 이하면 정상으로 간주 (쉬는시간 오차)
   【규칙 9】 소급결제는 시간 겹침·지연 검사에서 제외
   【규칙10】 합주 수업과 학생이 완전히 겹치는 일반 수업은 겹침 아님
   【규칙11】 지연결제 기준: 화면의 "지연 기준" 입력값 (기본 10분)
              지연 = 결제 시각 − 시간표상 종료 시각 (음수면 선결제)
   【규칙12】 승인취소 건은 무시
   【규칙13】 승인일시는 10분 단위 내림 처리 (16:07 → 16:00)
   【규칙14】 RULE_시간표면제_선생님 목록은 시간표에 없어도 정상 처리
              (시간 겹침만 검사)
   【규칙15】 오감(*오감)은 제공인력 명단 중 누가 결제해도 정상
   【규칙16】 부재 설정된 선생님/학생의 결제 → 과오결제, 미결제 → 검사 제외
   【규칙17】 시트 선택: 현재 년월(YY.MM) 이름의 시트, 없으면 첫 번째 시트
   【규칙18】 "진행 중" 모드: 아직 종료 전 수업의 미결제는 "결제 예정"
              "결제 완료" 모드: 모든 미결제는 "누락"
   【규칙19】 시간표는 하드코딩 없음. 엑셀 업로드 → 저장 → localStorage에서
              자동 로드. "저장된 시간표 삭제" 시 빈 시간표가 됩니다.
   ============================================================================ */

const RULE_합주분 = 20;           // 규칙2: 합주 수업시간(분)
const RULE_그룹묶음분 = 15;        // 규칙7: 같은 그룹으로 묶는 결제 간격(분)
const RULE_겹침허용분 = 10;        // 규칙8: 이 시간 이하 겹침은 정상(분)
const RULE_어두운셀밝기 = 180;     // 규칙6: RGB 평균이 이 값 미만이면 제외
const RULE_시간표면제_선생님 = ['전미라','김정희'];  // 규칙14

// 규칙1: 수업 유형 → 화면표시명 + 기본 수업시간(분)
const SERVICE_MAP = {
  '아동청소년심리지원서비스':{name:'심리',mins:50},
  '아동정서발달 및 치유 서비스':{name:'정서',mins:60},
  '아동비전형성지원서비스(통합형)':{name:'비전',mins:90},
  '오감톡톡상상놀이터(감성놀이 지원서비스)':{name:'오감',mins:100},
  '발달재활':{name:'발달',mins:50},
  '아동청소년행복지원서비스':{name:'해피',mins:60},
  '해피패밀리 프로젝트':{name:'해피',mins:60},
  '인터넷 과몰입 아동 청소년 치유 서비스':{name:'과몰입',mins:50},
};

// 시간표 파싱 시 학생 이름으로 오인하면 안 되는 단어들
const SKIP_WORDS = ['보강결제','정상결제','소급결제','합주','부모상담','카드들고','입원중',
  '주소변','선생님','제공인력','해외','태국','흰색통','자혜원','천양원',
  '정림원','카드재발급','숲센터','발달하고','전산에','확인후','변경되었는지'];

/* ============================================================================
   ★ 시간표 데이터 (규칙19: 하드코딩 없음)
   - 엑셀 업로드 → 저장하면 localStorage에 보관되고 접속 시 자동 로드됩니다.
   - 예전처럼 코드에 시간표를 직접 넣지 마세요. 옛 데이터가 섞이는 원인이 됩니다.
   ============================================================================ */
const SCHEDULE = {1:[],2:[],3:[],4:[],5:[],6:[]};

// 합주 수업 목록 (요일: 1=월 ... 6=토) — 시간표 엑셀 파싱 시 자동 채워짐
const HAPJU = {1:[],2:[],3:[],4:[],5:[],6:[]};

let allData=[], groups=[];

// ── 부재 관리 ──
let absenceList = [];
const ABSENCE_STORAGE_KEY = 'timetable_absence_v2';

function toggleAbsence(){
  const panel=document.getElementById('absencePanel');
  const icon=document.getElementById('absenceToggleIcon');
  const open=panel.style.display==='none';
  panel.style.display=open?'block':'none';
  icon.textContent=open?'▲ 접기':'▼ 펼치기';
}

function addAbsence(){
  const type=document.getElementById('absenceType').value;
  const name=document.getElementById('absenceName').value.trim();
  const reason=document.getElementById('absenceReason').value.trim();
  const startDate=document.getElementById('absenceStart').value;
  const endDate=document.getElementById('absenceEnd').value;
  if(!name) return;
  absenceList.push({type, name, reason, startDate, endDate});
  document.getElementById('absenceName').value='';
  document.getElementById('absenceReason').value='';
  document.getElementById('absenceStart').value='';
  document.getElementById('absenceEnd').value='';
  saveAbsenceList();
  renderAbsenceList();
  if(allData.length) rerun();
}

function removeAbsence(idx){
  absenceList.splice(idx,1);
  saveAbsenceList();
  renderAbsenceList();
  if(allData.length) rerun();
}

function saveAbsenceList(){
  localStorage.setItem(ABSENCE_STORAGE_KEY, JSON.stringify(absenceList));
}

function loadAbsenceList(){
  try{
    const saved=localStorage.getItem(ABSENCE_STORAGE_KEY);
    if(saved){ absenceList=JSON.parse(saved); renderAbsenceList(); }
  }catch(e){}
}

function isAbsenceActive(a, checkDateStr){
  if(!a.startDate && !a.endDate) return true;
  if(a.startDate && checkDateStr < a.startDate) return false;
  if(a.endDate && checkDateStr > a.endDate) return false;
  return true;
}

function fmtDate(d){
  if(!d) return '';
  const [,m,day]=d.split('-');
  return `${parseInt(m)}/${parseInt(day)}`;
}

function absDateYMD(d){
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function initAbsenceDates(){
  const today = nowKST();
  const end = new Date(today);
  end.setDate(end.getDate() + 6);
  document.getElementById('absenceStart').value = absDateYMD(today);
  document.getElementById('absenceEnd').value = absDateYMD(end);
}

function onAbsenceStartChange(){}

function renderAbsenceList(){
  const el=document.getElementById('absenceList');
  const todayStr=nowKST().toISOString().slice(0,10);

  if(!absenceList.length){
    el.innerHTML='<span style="font-size:12px;color:var(--text3);">부재 목록이 없습니다</span>';
    updateAbsenceSummary();
    return;
  }

  el.innerHTML=`<table style="width:100%;border-collapse:collapse;font-size:13px;">
    <thead>
      <tr style="background:var(--surface2);">
        <th style="padding:7px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);font-weight:500;font-size:11px;">구분</th>
        <th style="padding:7px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);font-weight:500;font-size:11px;">이름</th>
        <th style="padding:7px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);font-weight:500;font-size:11px;">사유</th>
        <th style="padding:7px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);font-weight:500;font-size:11px;">기간</th>
        <th style="padding:7px 12px;text-align:center;border-bottom:.5px solid var(--border);color:var(--text2);font-weight:500;font-size:11px;">상태</th>
        <th style="padding:7px 12px;border-bottom:.5px solid var(--border);"></th>
      </tr>
    </thead>
    <tbody>
      ${absenceList.map((a,i)=>{
        const active=isAbsenceActive(a, todayStr);
        const expired=a.endDate && todayStr > a.endDate;
        const period=a.startDate||a.endDate
          ? `${a.startDate?fmtDate(a.startDate):'시작일 없음'} ~ ${a.endDate?fmtDate(a.endDate):'종료일 없음'}`
          : '무기한';
        const statusBadge=expired
          ? `<span class="badge" style="background:var(--surface2);color:var(--text3);">종료됨</span>`
          : `<span class="badge badge-error">부재 중</span>`;
        const rowStyle=expired?'opacity:0.5;':'';
        return `<tr style="${rowStyle}border-bottom:.5px solid var(--border);">
          <td style="padding:8px 12px;color:var(--text3);font-size:12px;">${a.type}</td>
          <td style="padding:8px 12px;font-weight:500;color:var(--text);">${a.name}</td>
          <td style="padding:8px 12px;color:var(--text2);">${a.reason||'—'}</td>
          <td style="padding:8px 12px;font-family:var(--mono);font-size:12px;color:var(--text2);">${period}</td>
          <td style="padding:8px 12px;text-align:center;">${statusBadge}</td>
          <td style="padding:8px 12px;text-align:center;">
            <button onclick="removeAbsence(${i})" style="background:none;border:none;cursor:pointer;color:var(--text3);font-size:16px;padding:0;line-height:1;" title="삭제">×</button>
          </td>
        </tr>`;
      }).join('')}
    </tbody>
  </table>`;
  updateAbsenceSummary();
}

function updateAbsenceSummary(){
  const todayStr=nowKST().toISOString().slice(0,10);
  const active=absenceList.filter(a=>isAbsenceActive(a,todayStr));
  const el=document.getElementById('absenceSummary');
  if(active.length){
    el.innerHTML=`— <span style="color:var(--danger);font-weight:500;">현재 부재 ${active.length}명</span> · 부재 중 결제되면 <span style="color:var(--danger);font-weight:500;">과오결제</span>로 표시됩니다`;
  } else {
    el.innerHTML='— 부재 중 결제되면 <span style="color:var(--danger);font-weight:500;">과오결제</span>로 표시됩니다';
  }
}

function isAbsent(teacher, student){
  const todayStr=nowKST().toISOString().slice(0,10);
  const checkStr=lastFileDate
    ? `${lastFileDate.getFullYear()}-${String(lastFileDate.getMonth()+1).padStart(2,'0')}-${String(lastFileDate.getDate()).padStart(2,'0')}`
    : todayStr;
  return absenceList.some(a=>
    isAbsenceActive(a, checkStr) &&
    ((a.type==='선생님' && a.name===teacher) ||
     (a.type==='학생' && a.name===student))
  );
}

let lastDow=5, lastFileDate=null;
let paymentMode='진행중';

function setMode(mode){
  paymentMode=mode;
  document.getElementById('modeBtn진행중').style.cssText=
    mode==='진행중'
      ? 'font-family:\'Noto Sans KR\',sans-serif;font-size:12px;padding:5px 14px;border:none;cursor:pointer;background:var(--accent);color:#fff;font-weight:500;'
      : 'font-family:\'Noto Sans KR\',sans-serif;font-size:12px;padding:5px 14px;border:none;cursor:pointer;background:var(--surface2);color:var(--text2);';
  document.getElementById('modeBtn완료').style.cssText=
    mode==='완료'
      ? 'font-family:\'Noto Sans KR\',sans-serif;font-size:12px;padding:5px 14px;border:none;cursor:pointer;background:var(--ok);color:#fff;font-weight:500;'
      : 'font-family:\'Noto Sans KR\',sans-serif;font-size:12px;padding:5px 14px;border:none;cursor:pointer;background:var(--surface2);color:var(--text2);';
  if(allData.length) rerun();
}

function rerun(){ if(allData.length) renderResults(lastDow, lastFileDate); }

function updateDowBadge(dow){
  const el=document.getElementById('dowBadge');
  if(el) el.textContent = DOW_NAMES[dow]+'요일';
}

function onDateChange(){
  const val = document.getElementById('targetDate').value;
  if(!val) return;
  const d = new Date(val+'T00:00:00');
  const dow = d.getDay();
  lastDow = dow;
  lastFileDate = d;
  updateDowBadge(dow);
  if(allData.length) renderResults(dow, d);
}

/* ============================================================================
   시간표 엑셀 파싱 (규칙 4·5·6 적용)
   ============================================================================ */
function handleScheduleXlsx(e){
  const file=e.target.files[0]; if(!file) return;
  setPdfStatus('엑셀 읽는 중...');
  const reader=new FileReader();
  reader.onload=evt=>{
    try{
      const wb=XLSX.read(evt.target.result,{type:'binary', cellStyles:true});
      // 규칙17: 현재 년월(YY.MM) 시트 우선, 없으면 첫 번째 시트
      const now=nowKST();
      const sheetName=`${String(now.getFullYear()).slice(2)}.${String(now.getMonth()+1).padStart(2,'0')}`;
      const ws=wb.Sheets[sheetName]||wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      const result=parseScheduleSheet(raw, ws);
      if(result.error){
        setPdfStatus('❌ 파싱 실패: '+result.error);
        return;
      }
      parsedSchedule=result.schedule;
      parsedHapju=result.hapju;
      const total=Object.values(result.schedule).reduce((s,a)=>s+a.length,0);
      // 요일별 건수 함께 표시 → 특정 요일 파싱 실패를 바로 확인 가능
      const perDow=[1,2,3,4,5,6].map(d=>`${'월화수목금토'[d-1]}${(result.schedule[d]||[]).length}`).join(' ');
      setPdfStatus(`✓ ${total}건 파싱 완료 (${perDow}) (시트: ${wb.Sheets[sheetName]?sheetName:wb.SheetNames[0]}) — 확인 후 저장`);
      document.getElementById('schedulePreview').style.display='block';
      renderPreview();
    }catch(err){
      setPdfStatus('❌ 오류: '+err.message);
      console.error(err);
    }
    e.target.value='';
  };
  reader.readAsBinaryString(file);
}

function parseScheduleSheet(raw, ws){
  const DOW_KR=['월','화','수','목','금','토'];
  const DOW_NUM={월:1,화:2,수:3,목:4,금:5,토:6};
  let headerRowIdx=-1;
  let colPos={};
  // 요일 헤더 찾기: 앞뒤 공백 제거 후 비교 ("목 " 같은 셀도 인식)
  for(let r=0;r<Math.min(5,raw.length);r++){
    const row=raw[r].map(v=>String(v).trim());
    const found=DOW_KR.filter(d=>row.includes(d));
    if(found.length>=4){
      headerRowIdx=r;
      DOW_KR.forEach(d=>{
        const idx=row.indexOf(d);
        if(idx>=0) colPos[d]=idx;
      });
      break;
    }
  }
  if(headerRowIdx<0) return{error:'요일 헤더를 찾을 수 없습니다'};

  // 요일별 컬럼 범위 (헤더에 없는 요일은 건너뜀)
  const presentDows=DOW_KR.filter(d=>colPos[d]!==undefined);
  const colRanges={};
  presentDows.forEach((d,i)=>{
    const start=colPos[d];
    const next=presentDows[i+1];
    const end=next!==undefined?colPos[next]:raw[headerRowIdx].length;
    colRanges[d]={start,end};
  });

  const SKIP=new Set(SKIP_WORDS);

  function parseAmt(n){ return n>=100?n*100:n*1000; }  // 규칙4

  function parseEndTime(text){
    // 규칙5: "*" 붙은 시간 최우선
    const starMatch = text.match(/(\d+)시\s*(\d+)분\s*\*|(\d+)시\s*(\d+)\s*\*|(\d+)시\s*\*/);
    if(starMatch){
      const h=+(starMatch[1]||starMatch[3]||starMatch[5]);
      const mn=+(starMatch[2]||starMatch[4]||0);
      return (h<10?h+12:h)*60+mn;
    }
    const lines=text.split('\n');
    const times=[];
    for(const line of lines){
      const l=line.trim();
      let m;
      m=l.match(/(\d+)시\s*(\d+)분/); if(m){let h=+m[1],mn=+m[2];if(h<10)h+=12;times.push(h*60+mn);continue;}
      m=l.match(/(\d+)시\s*(\d+)(?!\d)/); if(m){let h=+m[1],mn=+m[2];if(h<10)h+=12;times.push(h*60+mn);continue;}
      m=l.match(/(\d+)시$/); if(m){let h=+m[1];if(h<10)h+=12;times.push(h*60);}
    }
    return times.length?times[times.length-1]:null;
  }

  // 규칙6: 어두운 배경 셀 제외
  function isDarkCell(cellObj){
    if(!cellObj||!cellObj.s) return false;
    const style=cellObj.s;
    const fg=style.fgColor||style.bgColor||(style.fill&&style.fill.fgColor);
    if(!fg) return false;
    const rgb=fg.rgb||fg.argb||'';
    if(rgb.length<6) return false;
    const hex=rgb.length===8?rgb.slice(2):rgb;
    const r=parseInt(hex.slice(0,2),16);
    const g=parseInt(hex.slice(2,4),16);
    const b=parseInt(hex.slice(4,6),16);
    const brightness=(r+g+b)/3;
    if(hex.length<6) return false;
    const alpha=rgb.length===8?rgb.slice(0,2):'FF';
    if(alpha==='00') return false;
    return brightness<RULE_어두운셀밝기;
  }

  const schedule={1:[],2:[],3:[],4:[],5:[],6:[]};
  const hapju={1:[],2:[],3:[],4:[],5:[],6:[]};

  presentDows.forEach(dow=>{
    const {start,end}=colRanges[dow];
    const dowNum=DOW_NUM[dow];
    const seen=new Set();
    const hapjuSeen=new Set();

    for(let r=headerRowIdx+1;r<raw.length;r++){
      for(let c=start;c<end;c++){
        const cellVal=raw[r][c];
        if(cellVal===undefined||cellVal===null||!String(cellVal).trim()) continue;

        if(ws){
          const cellAddr=XLSX.utils.encode_cell({r,c});
          const cellObj=ws[cellAddr];
          if(isDarkCell(cellObj)) continue;
        }
        const cell=String(cellVal).trim();
        const lines=cell.split('\n').map(l=>l.trim()).filter(Boolean);
        if(!lines.length) continue;

        const isHapjuCell=cell.includes('합주');
        const isBumoSangdam=cell.includes('부모상담');
        const isBogang=cell.includes('보강');  // '보강' 또는 '보강결제' 모두 인식
        const endTime=parseEndTime(cell);

        // 형식1: 유형:선생님 또는 유형 선생님
        const m1=lines[0].match(/^(심리|정서|비전|오감|발달|해피|과몰입)\s*[：:\s]\s*([가-힣]{2,4})/);
        if(m1){
          const typeKey=m1[1];
          const teacher=m1[2].replace(/\(.*?\)/g,'').trim().replace(/\s+/g,'');
          // 규칙3: 과몰입+부모상담+보강결제 → 합주 취급
          const effectiveHapju = isHapjuCell || (typeKey==='과몰입' && isBumoSangdam && isBogang);
          const students=[];
          for(const line of lines.slice(1)){
            if(/^\(/.test(line)) continue;
            for(const sm of line.matchAll(/([가-힣]{2,5})(\d{2,3})/g)){
              if(!SKIP.has(sm[1])) students.push([sm[1],parseAmt(+sm[2])]);
            }
          }
          if(effectiveHapju && isBogang && students.length){
            const hk=teacher+'|'+students.map(s=>s[0]).sort().join('|');
            if(!hapjuSeen.has(hk)){
              hapjuSeen.add(hk);
              hapju[dowNum].push({선생님:teacher,학생들:students.map(s=>s[0])});
            }
          }
          for(const [name,amt] of students){
            const k=`${teacher}|${name}|${typeKey}|${isBogang?'보강':'정상'}|${effectiveHapju}`;
            if(seen.has(k)) continue;
            seen.add(k);
            const item={선생님:teacher,대상자:name,유형:typeKey,금액:amt,
              합주:effectiveHapju,부모상담:(typeKey==='과몰입'&&isBumoSangdam&&isBogang),결제구분:isBogang?'보강결제':'정상결제'};
            if(endTime) item.종료=endTime;
            schedule[dowNum].push(item);
          }
          continue;
        }

        // 형식2: 오감 (제공인력 형식) — 규칙15
        if(lines[0].startsWith('오감')){
          const students=[];
          const teachers=[];
          let inTeacher=false;
          for(const line of lines.slice(1)){
            if(line.includes('제공인력')){ inTeacher=true; continue; }
            if(inTeacher){
              for(const n of line.match(/[가-힣]{2,4}/g)||[]){
                if(!SKIP.has(n)) teachers.push(n);
              }
            } else {
              for(const sm of line.matchAll(/([가-힣]{2,4})(\d{2,3})/g)){
                if(!SKIP.has(sm[1])) students.push([sm[1],parseAmt(+sm[2])]);
              }
            }
          }
          if(!students.length) continue;
          for(const [name,amt] of students){
            const k=`*오감|${name}|오감|${isBogang?'보강':'정상'}`;
            if(seen.has(k)) continue;
            seen.add(k);
            const item={선생님:'*오감',대상자:name,유형:'오감',금액:amt,
              합주:false,결제구분:isBogang?'보강결제':'정상결제',
              제공인력:teachers};
            if(endTime) item.종료=endTime;
            schedule[dowNum].push(item);
          }
        }
      }
    }
  });

  return{schedule,hapju};
}

let parsedSchedule = {};
let parsedHapju = null;
const STORAGE_KEY = 'timetable_schedule_v1';
const HAPJU_STORAGE_KEY = 'timetable_hapju_v1';

const TYPE_MAP_PDF = {
  '심리':'아동청소년심리지원서비스',
  '정서':'아동정서발달 및 치유 서비스',
  '비전':'아동비전형성지원서비스(통합형)',
  '오감':'오감톡톡상상놀이터(감성놀이 지원서비스)',
  '발달':'발달재활',
  '해피':'아동청소년행복지원서비스',
  '과몰입':'과몰입',
};

function parseAmount(n){
  if(n>=100) return n*100;
  return n*1000;
}

function togglePdfPanel(){
  const panel=document.getElementById('pdfPanel');
  const icon=document.getElementById('pdfToggleIcon');
  const open=panel.style.display==='none';
  panel.style.display=open?'block':'none';
  icon.textContent=open?'▲ 접기':'▼ 펼치기';
}

function setPdfStatus(msg){ document.getElementById('pdfStatusMsg').textContent=msg; }

async function handlePdf(e){
  const file=e.target.files[0]; if(!file) return;
  setPdfStatus('PDF 로딩 중...');
  try {
    if(!window.pdfjsLib){
      await loadPdfJs();
    }
    const ab=await file.arrayBuffer();
    const pdf=await pdfjsLib.getDocument({data:ab}).promise;
    let fullText='';
    for(let p=1;p<=pdf.numPages;p++){
      const page=await pdf.getPage(p);
      const tc=await page.getTextContent();
      fullText+=tc.items.map(i=>i.str).join('\n')+'\n';
    }
    setPdfStatus('텍스트 파싱 중...');
    parsedSchedule=parsePdfText(fullText);
    const total=Object.values(parsedSchedule).reduce((s,a)=>s+a.length,0);
    setPdfStatus(`✓ ${total}건 파싱 완료 — 확인 후 저장하세요`);
    document.getElementById('schedulePreview').style.display='block';
    renderPreview();
  } catch(err){
    setPdfStatus('❌ 파싱 실패: '+err.message);
    console.error(err);
  }
  e.target.value='';
}

function loadPdfJs(){
  return new Promise((res,rej)=>{
    const s=document.createElement('script');
    s.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    s.onload=()=>{
      pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      res();
    };
    s.onerror=rej;
    document.head.appendChild(s);
  });
}

function parsePdfText(text){
  const lines=text.split('\n').map(l=>l.trim()).filter(Boolean);
  const result={1:[],2:[],3:[],4:[],5:[],6:[]};

  let curType='', curTeacher='', curHapju=false, curBogang=false;
  const schedItems=[];
  let i=0;

  while(i<lines.length){
    const line=lines[i];
    const blockMatch=line.match(/^(심리|정서|비전|오감|발달|해피|과몰입)\s*[：:]\s*(.+)/);
    if(blockMatch){
      curType=blockMatch[1];
      curTeacher=blockMatch[2].trim()
        .replace(/\s*\(.*?\)\s*/g,'')
        .trim();
      curHapju=false; curBogang=false;
      i++;

      const students=[];
      let bogang날짜='';

      while(i<lines.length){
        const nl=lines[i];
        if(nl.match(/^(심리|정서|비전|오감|발달|해피|과몰입)\s*[：:]/)) break;
        if(nl.includes('합주')) curHapju=true;
        if(nl.includes('보강')) curBogang=true;  // '보강' 또는 '보강결제' 모두 인식
        const dateMatch=nl.match(/(\d{2})\.(\d{2})\.(\d{2})/);
        if(dateMatch) bogang날짜=`20${dateMatch[1]}-${dateMatch[2]}-${dateMatch[3]}`;
        const stuMatches=[...nl.matchAll(/([가-힣]{2,4})(\d{2,3})/g)];
        stuMatches.forEach(m=>students.push({이름:m[1],금액raw:parseInt(m[2])}));
        i++;
      }

      students.forEach(s=>{
        const item={
          선생님:curTeacher,
          대상자:s.이름,
          유형명:curType,
          사업유형:TYPE_MAP_PDF[curType]||curType,
          금액:parseAmount(s.금액raw),
          합주:curHapju,
          결제구분:curBogang?'보강결제':'정상결제',
          보강날짜:bogang날짜,
        };
        schedItems.push(item);
      });
      continue;
    }
    i++;
  }

  schedItems.forEach(item=>{
    if(item.보강날짜){
      const d=new Date(item.보강날짜+'T00:00:00');
      const dow=d.getDay();
      if(dow>=1&&dow<=6){
        if(!result[dow]) result[dow]=[];
        result[dow].push({...item});
      }
    } else {
      for(let d=1;d<=6;d++){
        result[d].push({...item});
      }
    }
  });

  for(const dow of Object.keys(result)){
    const seen=new Set();
    result[dow]=result[dow].filter(s=>{
      const k=`${s.선생님}|${s.대상자}|${s.유형명}|${s.결제구분}`;
      if(seen.has(k)) return false;
      seen.add(k); return true;
    });
  }

  return result;
}

function renderPreview(){
  const dow=parseInt(document.getElementById('previewDow').value);
  const items=(parsedSchedule[dow]||[]);
  const el=document.getElementById('previewTable');
  if(!items.length){
    el.innerHTML='<div style="padding:20px;text-align:center;color:var(--text3);font-size:13px;">해당 요일 데이터 없음</div>';
    return;
  }
  let html=`<table style="width:100%;border-collapse:collapse;font-size:12px;background:var(--surface);">
    <thead style="background:var(--surface2);">
      <tr>
        <th style="padding:8px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);">선생님</th>
        <th style="padding:8px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);">학생</th>
        <th style="padding:8px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);">유형</th>
        <th style="padding:8px 12px;text-align:right;border-bottom:.5px solid var(--border);color:var(--text2);">금액</th>
        <th style="padding:8px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);">종료</th>
        <th style="padding:8px 12px;text-align:left;border-bottom:.5px solid var(--border);color:var(--text2);">결제구분</th>
        <th style="padding:8px 12px;text-align:center;border-bottom:.5px solid var(--border);color:var(--text2);">합주</th>
      </tr>
    </thead><tbody>`;
  items.forEach(s=>{
    const typeName=s.유형||s.유형명||'';
    const endStr=s.종료?`${String(Math.floor(s.종료/60)).padStart(2,'0')}:${String(s.종료%60).padStart(2,'0')}`:'—';
    html+=`<tr>
      <td style="padding:7px 12px;border-bottom:.5px solid var(--border);color:var(--text);">${s.선생님}</td>
      <td style="padding:7px 12px;border-bottom:.5px solid var(--border);color:var(--text);font-weight:500;">${s.대상자}</td>
      <td style="padding:7px 12px;border-bottom:.5px solid var(--border);"><span class="chip chip-${typeName}">${typeName}</span></td>
      <td style="padding:7px 12px;border-bottom:.5px solid var(--border);text-align:right;font-family:var(--mono);font-size:12px;">${s.금액.toLocaleString()}</td>
      <td style="padding:7px 12px;border-bottom:.5px solid var(--border);font-family:var(--mono);font-size:12px;">${endStr}</td>
      <td style="padding:7px 12px;border-bottom:.5px solid var(--border);"><span class="badge ${(s.결제구분||'정상결제')==='보강결제'?'badge-보강':'badge-info'}">${(s.결제구분||'정상결제').replace('결제','')}</span></td>
      <td style="padding:7px 12px;border-bottom:.5px solid var(--border);text-align:center;">${hapjuChip(s)}</td>
    </tr>`;
  });
  html+='</tbody></table>';
  el.innerHTML=html;
}

function saveSchedule(){
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(parsedSchedule));
    if(parsedHapju) localStorage.setItem(HAPJU_STORAGE_KEY, JSON.stringify(parsedHapju));
    applyParsedSchedule(parsedSchedule, parsedHapju);
    setPdfStatus(`✅ 저장 완료! 다음 접속 시 자동 로드됩니다`);
    updateScheduleStatus();
    if(allData.length) rerun();
  } catch(err){
    setPdfStatus('저장 실패: '+err.message);
  }
}

function clearSchedule(){
  if(!confirm('저장된 시간표를 삭제할까요? 시간표가 비워지며, 새 엑셀을 업로드해야 검증할 수 있습니다.')) return;
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(HAPJU_STORAGE_KEY);
  parsedSchedule={};
  parsedHapju=null;
  // 규칙19: 하드코딩 없음 → 삭제 시 시간표는 완전히 비워짐
  for(let d=1;d<=6;d++){ SCHEDULE[d]=[]; HAPJU[d]=[]; }
  document.getElementById('schedulePreview').style.display='none';
  updateScheduleStatus();
  setPdfStatus('삭제됨. 새 시간표 엑셀을 업로드해주세요.');
  if(allData.length) rerun();
}

function applyParsedSchedule(sched, hapju){
  // ★ 전체 교체: 파싱 결과에 없는 요일도 비움 (옛 데이터가 섞이는 것 방지)
  for(let d=1;d<=6;d++){
    SCHEDULE[d]=(sched && sched[d]) ? sched[d] : [];
    HAPJU[d]=(hapju && hapju[d]) ? hapju[d] : [];
  }
}

function updateScheduleStatus(){
  const saved=localStorage.getItem(STORAGE_KEY);
  const el=document.getElementById('scheduleStatus');
  if(saved){
    const s=JSON.parse(saved);
    const total=Object.values(s).reduce((a,b)=>a+b.length,0);
    const perDow=[1,2,3,4,5,6].map(d=>`${'월화수목금토'[d-1]}${(s[d]||[]).length}`).join(' ');
    el.innerHTML=`— <span style="color:var(--ok);font-weight:500;">저장된 시간표 적용 중 (${total}건 · ${perDow})</span>`;
  } else {
    el.innerHTML='— <span style="color:var(--danger);font-weight:500;">시간표 없음</span> · 스프레드시트에서 엑셀을 다운받아 업로드 해주세요';
  }
}

function loadSavedSchedule(){
  try {
    const saved=localStorage.getItem(STORAGE_KEY);
    if(saved){
      parsedSchedule=JSON.parse(saved);
      const savedHapju=localStorage.getItem(HAPJU_STORAGE_KEY);
      parsedHapju=savedHapju?JSON.parse(savedHapju):null;
      applyParsedSchedule(parsedSchedule, parsedHapju);
    }
    updateScheduleStatus();
  } catch(e){ console.error('시간표 로드 실패',e); }
}

function nowKST(){
  const now=new Date();
  const kstOffset=9*60;
  const localOffset=now.getTimezoneOffset();
  return new Date(now.getTime()+(kstOffset+localOffset)*60000);
}

function todayKST(){
  const k=nowKST();
  return new Date(k.getFullYear(),k.getMonth(),k.getDate());
}

window.addEventListener('load', function(){
  loadSavedSchedule();
  loadAbsenceList();
  initAbsenceDates();
  const today = nowKST();
  const ymd = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  const el = document.getElementById('targetDate');
  if(el){ el.value = ymd; }
  lastDow = today.getDay();
  lastFileDate = today;
  updateDowBadge(today.getDay());

  document.querySelectorAll('.tab[data-tab]').forEach(btn=>{
    btn.addEventListener('click', function(){
      showTab(this.dataset.tab, this);
    });
  });
  const dz=document.getElementById('dropZone');
  if(dz){
    dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag')});
    dz.addEventListener('dragleave',()=>dz.classList.remove('drag'));
    dz.addEventListener('drop',e=>{
      e.preventDefault();dz.classList.remove('drag');
      const f=e.dataTransfer.files[0];
      if(f) handleFile({target:{files:[f]}});
    });
  }
});

function handleFile(e){
  const file=e.target.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=evt=>{
    try{
      const wb=XLSX.read(evt.target.result,{type:'binary',cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{raw:false});
      processData(rows);
    }catch(err){
      alert('파일 처리 오류: '+err.message);
      console.error(err);
    }
  };
  reader.readAsBinaryString(file);
}

function getSvcInfo(t){
  if(!t)return null;
  for(const[k,v]of Object.entries(SERVICE_MAP)){if(t.includes(k.slice(0,8)))return v;}
  return null;
}

function parseTime(s){
  if(!s)return null;
  const d=new Date(s);
  if(isNaN(d))return null;
  // 규칙13: 승인일시 10분 단위 내림
  const floored=Math.floor(d.getMinutes()/10)*10;
  d.setMinutes(floored,0,0);
  return d;
}

function toHHMM(d){
  if(!d)return'--:--';
  return String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');
}
function subMins(d,m){return new Date(d.getTime()-m*60000);}
function addMins(d,m){return new Date(d.getTime()+m*60000);}
function fmtMoney(n){return Number(n).toLocaleString()+'원';}
function hapjuChip(s){
  if(!s.합주) return '';
  if(s.부모상담) return '<span class="chip chip-합주">부모상담</span>';
  return '<span class="chip chip-합주">합주</span>';
}
function hapjuChipG(g){
  if(!g.합주) return '';
  if(g.부모상담) return '<span class="chip chip-합주">부모상담</span>';
  return '<span class="chip chip-합주">합주</span>';
}

function fmtApprovalTime(d){
  const t=d&&(d.원본승인일시||d.승인일시);
  return t?`<span style="font-family:var(--mono);font-size:11px;color:var(--text3);margin-left:6px;">🕐 ${toHHMM(t)}</span>`:'';
}

// 규칙3: 합주 여부 확인 — 보강결제 + HAPJU 목록 포함 시 합주
function isHapju(dow, teacher, student, payType, svcName){
  if(payType !== '보강결제') return false;
  if(svcName === '과몰입') return true;
  const list = HAPJU[dow]||[];
  return list.some(h => h.선생님===teacher && h.학생들.includes(student));
}

const DOW_NAMES=['일','월','화','수','목','금','토'];

function processData(rows){
  allData=[];
  let fileDate=null;
  rows.forEach(r=>{
    if(!r['제공인력명']||!r['사업유형'])return;
    // 규칙12: 승인취소 무시
    if((r['결제구분']||'').trim()==='승인취소')return;
    const svcInfo=getSvcInfo(r['사업유형']);
    if(!svcInfo)return;
    const t=parseTime(r['승인일시']);
    if(!t)return;
    const rawT=(()=>{const d=new Date(r['승인일시']);return isNaN(d)?t:d;})();
    if(!fileDate)fileDate=t;
    const dow=t.getDay();
    const payType=(r['결제구분']||'').trim();
    const hapju=isHapju(dow,(r['제공인력명']||'').trim(),(r['대상자명']||'').trim(),payType,svcInfo.name);
    const mins=hapju?RULE_합주분:svcInfo.mins;  // 규칙2
    const 종료=t;
    const 시작=subMins(종료, mins);
    const teacher=(r['제공인력명']||'').trim();
    const student=(r['대상자명']||'').trim();
    const payDow=t.getDay();
    const daySched=SCHEDULE[payDow]||[];
    const schedItem=daySched.find(s=>
      (s.선생님===teacher||s.선생님==='*오감') &&
      s.대상자===student &&
      (s.결제구분||'정상결제')===payType
    );
    // 규칙11: 지연 = 결제 시각 − 시간표 종료 시각
    let 지연분=0;
    if(schedItem && schedItem.종료){
      const schedEnd=new Date(rawT);
      schedEnd.setHours(Math.floor(schedItem.종료/60), schedItem.종료%60, 0, 0);
      지연분=Math.round((rawT.getTime()-schedEnd.getTime())/60000);
    } else {
      지연분=Math.floor((rawT.getTime()-종료.getTime())/60000);
    }
    allData.push({
      승인일시:t, 원본승인일시:rawT, 결제구분:payType,
      사업유형:r['사업유형']||'',유형명:svcInfo.name,수업분:mins,합주:hapju,
      대상자:student,선생님:teacher,
      총금액:parseInt(String(r['총금액']).replace(/,/g,''))||0,
      종료, 시작, 지연분,
      스케쥴종료: schedItem?.종료 || null,
    });
  });
  const dow=fileDate?fileDate.getDay():new Date().getDay();
  lastDow=dow; lastFileDate=fileDate||new Date();
  if(fileDate){
    const ymd=`${fileDate.getFullYear()}-${String(fileDate.getMonth()+1).padStart(2,'0')}-${String(fileDate.getDate()).padStart(2,'0')}`;
    document.getElementById('targetDate').value=ymd;
    updateDowBadge(dow);
  }
  renderResults(dow, lastFileDate);
}

/* ============================================================================
   그룹핑 · 겹침 검사 (규칙 7·8·9·10 적용)
   ============================================================================ */
function makeGroup(cur){
  const rep=cur.items[0];
  const 종료=rep.종료;
  const 시작=rep.시작;
  const 결제구분list=cur.items.map(i=>i.결제구분);
  const 소급=결제구분list.some(p=>p.includes('소급'));  // 규칙9
  return{
    선생님:rep.선생님,유형명:rep.유형명,수업분:rep.수업분,합주:rep.합주,소급,
    시작,종료,
    학생들:cur.items.map(i=>i.대상자),
    결제구분:결제구분list,
    금액들:cur.items.map(i=>i.총금액),
    raw:cur.items,
  };
}

function groupSessions(extraKeys){
  const byKey={};
  allData.forEach(d=>{
    if(extraKeys && extraKeys.has(d.선생님+'|'+d.대상자+'|'+d.결제구분)) return;
    const key=d.선생님+'|'+d.유형명+'|'+(d.합주?'합주':'일반');
    if(!byKey[key])byKey[key]=[];
    byKey[key].push(d);
  });
  groups=[];
  for(const items of Object.values(byKey)){
    items.sort((a,b)=>a.종료-b.종료);
    let cur=null;
    items.forEach(item=>{
      if(!cur){cur={items:[item]};return;}
      const diff=(item.종료-cur.items[0].종료)/60000;
      if(diff<=RULE_그룹묶음분)cur.items.push(item);  // 규칙7
      else{groups.push(makeGroup(cur));cur={items:[item]};}
    });
    if(cur)groups.push(makeGroup(cur));
  }
  groups.sort((a,b)=>a.종료-b.종료);
}

function findOverlaps(extraKeys){
  groupSessions(extraKeys||new Set());
  const issues=[];
  const seen=new Set();

  function checkPair(a, b, label, labelVal){
    const os=Math.max(a.시작.getTime(),b.시작.getTime());
    const oe=Math.min(a.종료.getTime(),b.종료.getTime());
    // 규칙8: 허용 시간 이하 겹침은 정상
    if(oe-os <= RULE_겹침허용분*60000) return;

    // 규칙10: 합주와 학생 완전 포함 관계면 정상
    if(a.합주 !== b.합주){
      const sa=new Set(a.학생들), sb=new Set(b.학생들);
      if(b.학생들.every(s=>sa.has(s)) || a.학생들.every(s=>sb.has(s))) return;
    }

    // 규칙9: 소급결제 제외
    if(a.소급 || b.소급) return;

    const key=[a,b].map(g=>g.선생님+'|'+g.시작.getTime()+'|'+g.유형명).sort().join('::');
    if(seen.has(key)) return;
    seen.add(key);

    issues.push({
      label, labelVal,
      a, b,
      overlapMins: Math.round((oe-os)/60000)
    });
  }

  // ① 선생님 기준 겹침
  const byTeacher={};
  groups.forEach(g=>{if(!byTeacher[g.선생님])byTeacher[g.선생님]=[];byTeacher[g.선생님].push(g);});
  for(const[t,gs]of Object.entries(byTeacher)){
    for(let i=0;i<gs.length;i++)for(let j=i+1;j<gs.length;j++){
      checkPair(gs[i],gs[j],'teacher',t);
    }
  }

  // ② 학생 기준 겹침 (다른 선생님 수업 간)
  const byStudent={};
  groups.forEach(g=>{
    g.학생들.forEach(s=>{
      if(!byStudent[s])byStudent[s]=[];
      byStudent[s].push(g);
    });
  });
  for(const[s,gs]of Object.entries(byStudent)){
    for(let i=0;i<gs.length;i++)for(let j=i+1;j<gs.length;j++){
      if(gs[i].선생님===gs[j].선생님) continue;
      checkPair(gs[i],gs[j],'student',s);
    }
  }

  return issues;
}

function getOverlapKeys(overlaps){
  const s=new Set();
  overlaps.forEach(o=>{
    s.add(o.a.선생님+'|'+o.a.시작.getTime());
    s.add(o.b.선생님+'|'+o.b.시작.getTime());
  });
  return s;
}

/* ============================================================================
   누락 · 금액 · 결제구분 · 과오 · 중복 검사 (규칙 14·15·16·18 적용)
   ============================================================================ */
function checkMissingAndAmount(dow){
  const sched=SCHEDULE[dow]||[];

  const now=nowKST();
  const fileDay=lastFileDate?new Date(lastFileDate):now;
  const isToday=fileDay.toDateString()===now.toDateString();
  const nowMins = now.getHours()*60 + now.getMinutes();

  const excelMap={};
  allData.forEach(d=>{
    const k=`${d.선생님}|${d.대상자}|${d.합주?'Y':'N'}|${d.결제구분}`;
    if(!excelMap[k])excelMap[k]=[];
    excelMap[k].push(d.총금액);
  });

  const schedKeys=new Set(sched.map(s=>`${s.선생님}|${s.대상자}|${s.합주?'Y':'N'}|${s.결제구분||'정상결제'}`));
  const schedKeysNoPayType=new Map();
  sched.forEach(s=>{
    const k=`${s.선생님}|${s.대상자}|${s.합주?'Y':'N'}`;
    schedKeysNoPayType.set(k, s.결제구분||'정상결제');
  });

  const missing=[], pending=[], extra=[], amountErr=[], wrongPayment=[], payTypeMismatch=[], duplicate=[];

  // 중복 결제 감지
  const excelCountMap={};
  allData.forEach(d=>{
    const k=`${d.선생님}|${d.대상자}|${d.합주?'Y':'N'}|${d.결제구분}`;
    excelCountMap[k]=(excelCountMap[k]||0)+1;
  });
  Object.entries(excelCountMap).forEach(([k,cnt])=>{
    if(cnt>1){
      const [teacher,student,hapju,pay]=k.split('|');
      const items=allData.filter(d=>
        d.선생님===teacher&&d.대상자===student&&
        (d.합주?'Y':'N')===hapju&&d.결제구분===pay
      );
      duplicate.push({teacher,student,hapju:hapju==='Y',결제구분:pay,count:cnt,items});
    }
  });

  // 규칙15: 오감 그룹
  const ogamExcelMap={};
  allData.forEach(d=>{
    if(d.유형명==='오감'){
      const k=`${d.대상자}|${d.결제구분}`;
      if(!ogamExcelMap[k]) ogamExcelMap[k]=[];
      ogamExcelMap[k].push({선생님:d.선생님, 금액:d.총금액, 시각:d.원본승인일시||d.승인일시});
    }
  });

  const ogamSchedKeys=new Set();
  sched.forEach(s=>{
    if(s.선생님==='*오감'){
      ogamSchedKeys.add(`${s.대상자}|${s.결제구분||'정상결제'}`);
    }
  });

  sched.forEach(s=>{
    if(s.선생님==='*오감'){
      const providers = s.제공인력 || [];
      const anyProviderAbsent = providers.some(t => isAbsent(t, s.대상자));
      if(isAbsent('', s.대상자) || anyProviderAbsent) return;
      const ok=`${s.대상자}|${s.결제구분||'정상결제'}`;
      if(!ogamExcelMap[ok]){
        const endMins=s.종료||0;
        const stillPending=isToday&&paymentMode==='진행중'&&endMins>nowMins;  // 규칙18
        if(stillPending) pending.push(s); else missing.push(s);
      } else {
        ogamExcelMap[ok].forEach(e=>{
          if(e.금액!==s.금액) amountErr.push({...s, 선생님:e.선생님, 실제금액:e.금액, 실제시각:e.시각});
        });
      }
      return;
    }

    const k=`${s.선생님}|${s.대상자}|${s.합주?'Y':'N'}|${s.결제구분||'정상결제'}`;
    if(!excelMap[k]){
      if(isAbsent(s.선생님, s.대상자)) return;  // 규칙16
      const endMins = s.종료 || 0;
      const stillPending = isToday && paymentMode==='진행중' && endMins > nowMins;  // 규칙18
      if(stillPending){
        pending.push(s);
      } else {
        missing.push(s);
      }
    } else {
      excelMap[k].forEach((amt,idx)=>{
        if(amt!==s.금액){
          const matched=allData.filter(d=>d.선생님===s.선생님&&d.대상자===s.대상자&&(d.합주?'Y':'N')===(s.합주?'Y':'N')&&d.결제구분===(s.결제구분||'정상결제'));
          const 실제시각=matched[idx]?.원본승인일시||matched[idx]?.승인일시||null;
          amountErr.push({...s, 실제금액:amt, 실제시각});
        }
      });
    }
  });

  // 규칙14: 시간표 면제 선생님
  const EXEMPT_TEACHERS = new Set(RULE_시간표면제_선생님);

  allData.forEach(d=>{
    const k=`${d.선생님}|${d.대상자}|${d.합주?'Y':'N'}|${d.결제구분}`;

    const payDate = d.승인일시 instanceof Date ? d.승인일시 : new Date(d.승인일시);
    const payDow = isNaN(payDate) ? dow : payDate.getDay();
    const daySched = SCHEDULE[payDow] || [];
    const daySchedKeys = new Set(daySched.map(s=>`${s.선생님}|${s.대상자}|${s.합주?'Y':'N'}|${s.결제구분||'정상결제'}`));
    const dayOgamKeys = new Set(daySched.filter(s=>s.선생님==='*오감').map(s=>`${s.대상자}|${s.결제구분||'정상결제'}`));
    const daySchedNoPayType = new Map();
    daySched.forEach(s=>{ if(s.선생님!=='*오감') daySchedNoPayType.set(`${s.선생님}|${s.대상자}|${s.합주?'Y':'N'}`, s.결제구분||'정상결제'); });

    if(d.유형명==='오감' && dayOgamKeys.has(`${d.대상자}|${d.결제구분}`)) return;
    if(EXEMPT_TEACHERS.has(d.선생님)) return;

    if(!daySchedKeys.has(k)){
      if(isAbsent(d.선생님, d.대상자)){
        wrongPayment.push(d);  // 규칙16
      } else {
        const kNoType=`${d.선생님}|${d.대상자}|${d.합주?'Y':'N'}`;
        const schedPayType=daySchedNoPayType.get(kNoType);
        if(schedPayType && schedPayType!==d.결제구분){
          payTypeMismatch.push({...d, 시간표결제구분:schedPayType, 실제결제구분:d.결제구분});
        } else {
          extra.push(d);
        }
      }
    }
  });

  // 결제 예정 → 누락 전환 (시간표 외 결제 건수만큼)
  const extraCountMap={};
  extra.forEach(d=>{
    const k=`${d.선생님}|${d.유형명}|${d.결제구분}`;
    extraCountMap[k]=(extraCountMap[k]||0)+1;
  });

  const finalPending=[], finalMissing=[...missing];
  const usedCount={};
  pending.forEach(p=>{
    const k=`${p.선생님}|${p.유형}|${p.결제구분||'정상결제'}`;
    const maxConvert=extraCountMap[k]||0;
    const used=usedCount[k]||0;
    if(maxConvert>0 && used<maxConvert){
      finalMissing.push(p);
      usedCount[k]=(used+1);
    } else {
      finalPending.push(p);
    }
  });

  allData.forEach(d=>{
    if(isAbsent(d.선생님, d.대상자)){
      const k=`${d.선생님}|${d.대상자}|${d.합주?'Y':'N'}|${d.결제구분}`;
      if(schedKeys.has(k) && !wrongPayment.some(w=>w.선생님===d.선생님&&w.대상자===d.대상자&&w.승인일시===d.승인일시)){
        wrongPayment.push(d);
      }
    }
  });

  return{missing:finalMissing,pending:finalPending,extra,amountErr,wrongPayment,payTypeMismatch,duplicate};
}

/* ============================================================================
   화면 렌더링
   ============================================================================ */
function renderResults(dow, fileDate){
  document.getElementById('results').style.display='block';
  document.getElementById('totalCount').textContent=allData.length;

  const{missing,pending,extra,amountErr,wrongPayment,payTypeMismatch,duplicate}=checkMissingAndAmount(dow);
  const extraKeys=new Set(extra.map(d=>d.선생님+'|'+d.대상자+'|'+d.결제구분));
  const overlaps=findOverlaps(extraKeys);
  const threshold=parseInt(document.getElementById('delayThreshold')?.value)||10;
  const delayed=allData.filter(d=>Math.abs(d.지연분)>=threshold && !d.결제구분.includes('소급'));

  const overlapPayCount = overlaps.reduce((sum,o)=>sum+o.a.raw.length+o.b.raw.length, 0);
  document.getElementById('overlapCount').textContent=overlapPayCount;
  document.getElementById('missingCount').textContent=missing.length;
  document.getElementById('pendingCount').textContent=pending.length;
  document.getElementById('amountErrCount').textContent=amountErr.length;
  document.getElementById('payMismatchCount').textContent=payTypeMismatch.length;
  document.getElementById('wrongCount').textContent=wrongPayment.length;
  document.getElementById('dupCount').textContent=duplicate.length;
  document.getElementById('delayCount').textContent=delayed.length;
  document.getElementById('extraCount').textContent=extra.length;

  const errInPayments=new Set();
  const payKey=(d)=>{
    const t=d.승인일시 instanceof Date ? d.승인일시.getTime() : String(d.승인일시);
    return t+'|'+d.선생님+'|'+d.대상자;
  };
  overlaps.forEach(o=>{
    o.a.raw.forEach(r=>errInPayments.add(payKey(r)));
    o.b.raw.forEach(r=>errInPayments.add(payKey(r)));
  });
  amountErr.forEach(e=>{
    allData.filter(d=>d.선생님===e.선생님&&d.대상자===e.대상자&&d.총금액===e.실제금액&&d.결제구분===(e.결제구분||'정상결제'))
      .forEach(d=>errInPayments.add(payKey(d)));
  });
  payTypeMismatch.forEach(e=>errInPayments.add(payKey(e)));
  wrongPayment.forEach(e=>errInPayments.add(payKey(e)));
  duplicate.forEach(e=>e.items.forEach(r=>errInPayments.add(payKey(r))));
  extra.forEach(e=>errInPayments.add(payKey(e)));
  document.getElementById('okCount').textContent=Math.max(0,allData.length-errInPayments.size);

  renderOverlap(overlaps);
  renderMissing(missing,extra);
  renderPending(pending);
  renderAmount(amountErr);
  renderPayTypeMismatch(payTypeMismatch);
  renderWrong(wrongPayment);
  renderDuplicate(duplicate);
  renderDelay(delayed, threshold);

  const teachers=[...new Set(groups.map(g=>g.선생님))].sort();
  ['teacherFilter','teacherFilter2'].forEach(id=>{
    const sel=document.getElementById(id);
    sel.innerHTML='<option value="">전체</option>';
    teachers.forEach(t=>sel.appendChild(new Option(t,t)));
  });
  const types=[...new Set(groups.map(g=>g.유형명))].sort();
  const tf=document.getElementById('typeFilter');
  tf.innerHTML='<option value="">전체</option>';
  types.forEach(t=>tf.appendChild(new Option(t,t)));

  renderSchedule();
  renderAll();
}

function renderOverlap(overlaps){
  const el=document.getElementById('overlapContent');
  if(!overlaps.length){
    el.innerHTML='<div class="issue-card ok-card"><div class="issue-title"><span class="badge badge-ok">✓ 이상 없음</span> 시간 겹침이 감지되지 않았습니다</div></div>';return;
  }
  const teacherOv = overlaps.filter(o=>o.label==='teacher');
  const studentOv = overlaps.filter(o=>o.label==='student');
  const totalPayCount = overlaps.reduce((s,o)=>s+o.a.raw.length+o.b.raw.length,0);
  let html=`<div class="legend">
    <span style="font-weight:500;color:var(--text)">겹침 ${totalPayCount}건 발견 (${overlaps.length}쌍)</span>
    <span class="badge badge-error">선생님 겹침 ${teacherOv.length}쌍</span>
    <span class="badge badge-warn">학생 겹침 ${studentOv.length}쌍</span>
  </div>`;

  function renderCard(o){
    const os=toHHMM(new Date(Math.max(o.a.시작.getTime(),o.b.시작.getTime())));
    const oe=toHHMM(new Date(Math.min(o.a.종료.getTime(),o.b.종료.getTime())));
    const isTeacher = o.label==='teacher';
    const titleText = isTeacher
      ? `<strong>${o.labelVal}</strong> 선생님 시간 겹침`
      : `학생 <strong>${o.labelVal}</strong> 수업 시간 겹침`;
    const cardClass = isTeacher ? 'error' : 'warn-card';
    const badgeClass = isTeacher ? 'badge-error' : 'badge-warn';
    const badgeText = isTeacher ? '선생님 겹침' : '학생 겹침';
    html+=`<div class="issue-card ${cardClass}">
      <div class="issue-title">
        <span class="badge ${badgeClass}">${badgeText}</span>${titleText}
        <span style="font-family:var(--mono);font-size:12px;background:var(--danger-light);color:var(--danger);padding:2px 10px;border-radius:20px;">⏱ ${o.overlapMins}분 (${os}~${oe})</span>
      </div>
      <div class="issue-desc">
        <div style="margin-bottom:6px">
          <span class="chip chip-${o.a.유형명}">${o.a.유형명}</span>
          ${hapjuChip(o.a)}
          <span class="time-range">${toHHMM(o.a.시작)}~${toHHMM(o.a.종료)}</span>
          <span style="color:var(--text2);font-size:12px;">${o.a.선생님}</span> · ${o.a.학생들.join(', ')}
        </div>
        <div>
          <span class="chip chip-${o.b.유형명}">${o.b.유형명}</span>
          ${hapjuChip(o.b)}
          <span class="time-range">${toHHMM(o.b.시작)}~${toHHMM(o.b.종료)}</span>
          <span style="color:var(--text2);font-size:12px;">${o.b.선생님}</span> · ${o.b.학생들.join(', ')}
        </div>
      </div>
    </div>`;
  }

  if(teacherOv.length){
    html+=`<div class="section-title">선생님 시간 겹침 (${teacherOv.length}건)</div>`;
    teacherOv.forEach(renderCard);
  }
  if(studentOv.length){
    html+=`<div class="section-title" style="margin-top:${teacherOv.length?'16px':'0'}">학생 수업 시간 겹침 (${studentOv.length}건)</div>`;
    studentOv.forEach(renderCard);
  }
  el.innerHTML=html;
}

function renderMissing(missing,extra){
  const el=document.getElementById('missingContent');
  let html='';
  if(missing.length){
    html+=`<div class="section-title">❌ 시간표에 있는데 결제 누락 (${missing.length}건)</div>`;
    missing.forEach(s=>{
      const timeStr = s.종료
        ? `<span style="font-family:var(--mono);font-size:12px;color:var(--text3);">🕐 ${String(Math.floor(s.종료/60)).padStart(2,'0')}:${String(s.종료%60).padStart(2,'0')}</span>`
        : '';
      html+=`<div class="issue-card error">
        <div class="issue-title"><span class="badge badge-error">누락</span><strong>${s.선생님==='*오감'?'오감그룹':s.선생님}</strong> · <strong>${s.대상자}</strong> ${hapjuChip(s)} ${timeStr}</div>
        <div class="issue-desc"><span class="chip chip-${s.유형}">${s.유형}</span> <span class="badge ${(s.결제구분||'정상결제')==='보강결제'?'badge-보강':'badge-info'}">${(s.결제구분||'정상결제').replace('결제','')}</span> 시간표 기준 <span class="amount-box amount-expect">${fmtMoney(s.금액)}</span></div>
      </div>`;
    });
  }
  if(extra.length){
    html+=`<div class="section-title" style="margin-top:${missing.length?'20px':'0'}">⚠ 시간표에 없는데 결제됨 (${extra.length}건)</div>`;
    extra.forEach(d=>{
      html+=`<div class="issue-card warn-card">
        <div class="issue-title"><span class="badge badge-warn">시간표 외</span><strong>${d.선생님}</strong> · <strong>${d.대상자}</strong> ${hapjuChip(d)}
          <span style="font-family:var(--mono);font-size:12px;color:var(--text2);">🕐 ${fmtTime(d.원본승인일시||d.승인일시)}</span>
        </div>
        <div class="issue-desc"><span class="chip chip-${d.유형명}">${d.유형명}</span> 결제금액 <span class="amount-box amount-actual">${fmtMoney(d.총금액)}</span> · ${d.결제구분}</div>
      </div>`;
    });
  }
  if(!missing.length&&!extra.length){
    html='<div class="issue-card ok-card"><div class="issue-title"><span class="badge badge-ok">✓ 이상 없음</span> 결제 누락 또는 시간표 외 결제가 없습니다</div></div>';
  }
  el.innerHTML=html;
}

function renderPending(pending){
  const el=document.getElementById('pendingContent');
  if(!pending.length){
    el.innerHTML='<div class="issue-card ok-card"><div class="issue-title"><span class="badge badge-ok">✓ 이상 없음</span> 결제 예정 수업이 없습니다</div></div>';return;
  }
  const now=nowKST();
  const hhmm=`${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;
  let html=`<div class="legend">
    <span style="font-weight:500;color:var(--text)">결제 예정 ${pending.length}건</span>
    <span style="font-size:12px;color:var(--text2);">현재 시각 ${hhmm} 기준 · 수업 종료 후 결제 예정</span>
  </div>`;
  const byTeacher={};
  pending.forEach(s=>{
    if(!byTeacher[s.선생님]) byTeacher[s.선생님]=[];
    byTeacher[s.선생님].push(s);
  });
  for(const[t,items] of Object.entries(byTeacher)){
    html+=`<div class="issue-card" style="border-left:3px solid var(--info);">
      <div class="issue-title">
        <span class="badge badge-info">결제 예정</span>
        <strong>${t==='*오감'?'오감그룹':t}</strong> 선생님 · ${items.length}건
      </div>
      <div class="issue-desc" style="margin-top:6px;">
        ${items.map(s=>`
          <div style="margin-bottom:4px;">
            <span class="chip chip-${s.유형}">${s.유형}</span>
            ${hapjuChip(s)}
            <strong>${s.대상자}</strong>
            <span class="amount-box amount-expect" style="margin-left:6px;">${fmtMoney(s.금액)}</span>
            <span style="font-size:11px;color:var(--text3);margin-left:4px;">${s.결제구분||'정상결제'}</span>
            ${s.종료?`<span style="font-family:var(--mono);font-size:11px;color:var(--text3);margin-left:4px;">🕐 ${String(Math.floor(s.종료/60)).padStart(2,'0')}:${String(s.종료%60).padStart(2,'0')}</span>`:''}
          </div>`).join('')}
      </div>
    </div>`;
  }
  el.innerHTML=html;
}

function fmtTime(d){
  if(!d) return '';
  const dt = d instanceof Date ? d : new Date(d);
  if(isNaN(dt)) return '';
  return `${String(dt.getHours()).padStart(2,'0')}:${String(dt.getMinutes()).padStart(2,'0')}`;
}

function renderAmount(amountErr){
  const el=document.getElementById('amountContent');
  if(!amountErr.length){
    el.innerHTML='<div class="issue-card ok-card"><div class="issue-title"><span class="badge badge-ok">✓ 이상 없음</span> 금액 오류가 없습니다</div></div>';return;
  }
  let html=`<div class="legend"><span style="font-weight:500;color:var(--text)">금액 오류 ${amountErr.length}건 발견</span></div>`;
  amountErr.forEach(e=>{
    const matched=allData.find(d=>d.선생님===e.선생님&&d.대상자===e.대상자&&d.총금액===e.실제금액&&d.결제구분===(e.결제구분||'정상결제'));
    const t=matched?(matched.원본승인일시||matched.승인일시):null;
    const timeStr=t?`<span style="font-family:var(--mono);font-size:11px;color:var(--text3);margin-left:6px;">🕐 ${toHHMM(t)}</span>`:'';
    html+=`<div class="issue-card error">
      <div class="issue-title">
        <span class="badge badge-error">금액오류</span>
        <strong>${e.선생님}</strong> · <strong>${e.대상자}</strong>
        ${hapjuChip(e)}
        ${timeStr}
      </div>
      <div class="issue-desc">
        <span class="chip chip-${e.유형}">${e.유형}</span>
        <span class="badge ${(e.결제구분||'정상결제')==='보강결제'?'badge-보강':'badge-info'}">${(e.결제구분||'정상결제').replace('결제','')}</span>
        시간표 기준 <span class="amount-box amount-expect">${fmtMoney(e.금액)}</span>
        → 실제 결제 <span class="amount-box amount-actual">${fmtMoney(e.실제금액)}</span>
      </div>
    </div>`;
  });
  el.innerHTML=html;
}

function renderPayTypeMismatch(list){
  const el=document.getElementById('payMismatchContent');
  if(!list.length){
    el.innerHTML='<div class="issue-card ok-card"><div class="issue-title"><span class="badge badge-ok">✓ 이상 없음</span> 결제구분 불일치가 없습니다</div></div>';return;
  }
  let html=`<div class="legend">
    <span style="font-weight:500;color:var(--text)">결제구분 불일치 ${list.length}건</span>
    <span style="font-size:12px;color:var(--text2);">시간표 기준과 실제 결제구분이 다른 항목입니다</span>
  </div>`;
  list.forEach(d=>{
    const schedBadge=d.시간표결제구분==='보강결제'?'badge-보강':'badge-info';
    const realBadge=d.실제결제구분==='보강결제'?'badge-보강':'badge-info';
    html+=`<div class="issue-card error">
      <div class="issue-title">
        <span class="badge badge-error">결제구분 불일치</span>
        <strong>${d.선생님}</strong> · <strong>${d.대상자}</strong>
        ${hapjuChip(d)}
        <span style="font-family:var(--mono);font-size:12px;color:var(--text2);">🕐 ${fmtTime(d.원본승인일시||d.승인일시)}</span>
      </div>
      <div class="issue-desc" style="margin-top:8px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
        <span class="chip chip-${d.유형명}">${d.유형명}</span>
        <span style="font-size:12px;color:var(--text2);">시간표</span>
        <span class="badge ${schedBadge}">${d.시간표결제구분}</span>
        <span style="font-size:16px;color:var(--text3);">→</span>
        <span style="font-size:12px;color:var(--text2);">실제 결제</span>
        <span class="badge ${realBadge}">${d.실제결제구분}</span>
        <span class="amount-box amount-actual" style="margin-left:4px;">${fmtMoney(d.총금액)}</span>
      </div>
    </div>`;
  });
  el.innerHTML=html;
}

function renderWrong(wrongPayment){
  const el=document.getElementById('wrongContent');
  if(!wrongPayment.length){
    el.innerHTML='<div class="issue-card ok-card"><div class="issue-title"><span class="badge badge-ok">✓ 이상 없음</span> 과오결제가 없습니다</div></div>';return;
  }
  function getReason(teacher, student){
    const a=absenceList.find(a=>(a.type==='선생님'&&a.name===teacher)||(a.type==='학생'&&a.name===student));
    return a?.reason||'부재';
  }
  let html=`<div class="legend"><span style="font-weight:500;color:var(--danger)">과오결제 ${wrongPayment.length}건 — 부재 중 결제된 내역입니다. 취소가 필요합니다.</span></div>`;
  wrongPayment.forEach(d=>{
    const reason=getReason(d.선생님,d.대상자);
    const absType=absenceList.find(a=>a.name===d.선생님)?'선생님':absenceList.find(a=>a.name===d.대상자)?'학생':'';
    html+=`<div class="issue-card error">
      <div class="issue-title">
        <span class="badge badge-error">과오결제</span>
        <strong>${d.선생님}</strong> · <strong>${d.대상자}</strong>
        <span style="background:var(--danger-light);color:var(--danger);font-size:11px;padding:2px 8px;border-radius:20px;">${absType} 부재 (${reason})</span>
        <span style="font-family:var(--mono);font-size:12px;color:var(--text2);">🕐 ${fmtTime(d.원본승인일시||d.승인일시)}</span>
      </div>
      <div class="issue-desc">
        <span class="chip chip-${d.유형명}">${d.유형명}</span>
        ${hapjuChip(d)}
        결제금액 <span class="amount-box amount-actual">${fmtMoney(d.총금액)}</span>
        · ${d.결제구분}
        · ${toHHMM(d.시작)}~${toHHMM(d.종료)}
      </div>
    </div>`;
  });
  el.innerHTML=html;
}

function renderDuplicate(duplicate){
  const el=document.getElementById('dupContent');
  if(!duplicate.length){
    el.innerHTML='<div class="issue-card ok-card"><div class="issue-title"><span class="badge badge-ok">✓ 이상 없음</span> 중복 결제가 없습니다</div></div>';return;
  }
  let html=`<div class="legend"><span style="font-weight:500;color:var(--danger)">중복결제 ${duplicate.length}건 — 동일 학생이 같은 결제구분으로 2회 이상 결제됐습니다</span></div>`;
  duplicate.forEach(d=>{
    html+=`<div class="issue-card error">
      <div class="issue-title">
        <span class="badge badge-error">중복 ${d.count}회</span>
        <strong>${d.teacher}</strong> · <strong>${d.student}</strong>
        <span class="badge ${d.결제구분==='보강결제'?'badge-보강':'badge-info'}">${d.결제구분}</span>
      </div>
      <div class="issue-desc" style="margin-top:6px;">
        ${d.items.map(item=>`
          <div style="margin-bottom:3px;font-size:12px;font-family:var(--mono);">
            ${toHHMM(item.원본승인일시||item.승인일시)} · ${fmtMoney(item.총금액)}
          </div>`).join('')}
      </div>
    </div>`;
  });
  el.innerHTML=html;
}

function renderDelay(delayed, threshold){
  const el=document.getElementById('delayContent');
  if(!delayed.length){
    el.innerHTML=`<div class="issue-card ok-card"><div class="issue-title"><span class="badge badge-ok">✓ 이상 없음</span> ${threshold}분 이상 지연된 결제가 없습니다</div></div>`;return;
  }
  const sorted=[...delayed].sort((a,b)=>Math.abs(b.지연분)-Math.abs(a.지연분));
  let html=`<div class="legend">
    <span style="font-weight:500;color:var(--text)">지연/선결제 ${delayed.length}건 (기준: ${threshold}분 이상)</span>
  </div>`;
  sorted.forEach(d=>{
    const abs분=Math.abs(d.지연분);
    const isEarly=d.지연분<0;
    const 원본시각=`${String(d.원본승인일시.getHours()).padStart(2,'0')}:${String(d.원본승인일시.getMinutes()).padStart(2,'0')}`;
    const 스케쥴종료시각 = d.스케쥴종료
      ? `${String(Math.floor(d.스케쥴종료/60)).padStart(2,'0')}:${String(d.스케쥴종료%60).padStart(2,'0')}`
      : toHHMM(d.종료);
    const color = abs분>=30 ? 'var(--danger)' : abs분>=20 ? 'var(--warn)' : '#888';
    const bgColor = abs분>=30 ? 'var(--danger-light)' : abs분>=20 ? 'var(--warn-light)' : 'var(--surface2)';
    const badgeLabel = isEarly ? `선결제 ${abs분}분` : `지연 ${abs분}분`;
    html+=`<div class="issue-card ${abs분>=30?'error':'warn-card'}">
      <div class="issue-title">
        <span class="badge ${abs분>=30?'badge-error':'badge-warn'}">${badgeLabel}</span>
        <strong>${d.선생님}</strong> · <strong>${d.대상자}</strong>
      </div>
      <div class="issue-desc" style="margin-top:6px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
        <span class="chip chip-${d.유형명}">${d.유형명}</span>
        ${hapjuChip(d)}
        <span style="font-family:var(--mono);font-size:12px;">
          종료 <strong>${스케쥴종료시각}</strong> → 결제 <strong>${원본시각}</strong>
        </span>
        <span style="font-size:12px;font-family:var(--mono);background:${bgColor};color:${color};padding:2px 10px;border-radius:20px;font-weight:500;">
          ${isEarly?'-':'+'}${abs분}분
        </span>
        <span style="font-size:12px;color:var(--text3);">${d.결제구분}</span>
      </div>
    </div>`;
  });
  el.innerHTML=html;
}

function renderSchedule(){
  const filter=document.getElementById('teacherFilter').value;
  const overlaps=findOverlaps();
  const ovKeys=getOverlapKeys(overlaps);
  const byT={};
  groups.forEach(g=>{if(filter&&g.선생님!==filter)return;if(!byT[g.선생님])byT[g.선생님]=[];byT[g.선생님].push(g);});
  const el=document.getElementById('scheduleContent');
  if(!Object.keys(byT).length){el.innerHTML='<div class="empty">데이터 없음</div>';return;}
  let html='';
  for(const[t,gs]of Object.entries(byT)){
    const hasOv=gs.some(g=>ovKeys.has(t+'|'+g.시작.getTime()));
    html+=`<div class="teacher-block"><div class="teacher-name">${t} ${hasOv?'<span class="badge badge-error">겹침 있음</span>':'<span class="badge badge-ok">정상</span>'}</div>`;
    gs.forEach(g=>{
      const isOv=ovKeys.has(t+'|'+g.시작.getTime());
      const hasB=g.결제구분.some(x=>x==='보강결제');
      html+=`<div class="issue-card ${isOv?'error':''}">
        <div class="issue-title">
          <span class="chip chip-${g.유형명}">${g.유형명} ${g.수업분}분</span>
          ${hapjuChipG(g)}
          <span class="time-range">${toHHMM(g.시작)}~${toHHMM(g.종료)}</span>
          ${isOv?'<span class="badge badge-error">겹침</span>':'<span class="badge badge-ok">정상</span>'}
          ${hasB?'<span class="badge badge-보강">보강</span>':''}
        </div>
        <div class="issue-desc" style="margin-top:6px">${g.학생들.map((s,i)=>`${s}<span style="font-size:10px;color:var(--text3);margin-left:2px">(${(g.결제구분[i]||'').replace('결제','')})</span>`).join(' · ')}</div>
      </div>`;
    });
    html+='</div>';
  }
  el.innerHTML=html;
}

function renderAll(){
  const tf=document.getElementById('typeFilter').value;
  const tt=document.getElementById('teacherFilter2').value;
  const pf=document.getElementById('payFilter').value;
  const overlaps=findOverlaps();
  const ovKeys=getOverlapKeys(overlaps);
  const filtered=groups.filter(g=>(!tf||g.유형명===tf)&&(!tt||g.선생님===tt)&&(!pf||g.결제구분.some(x=>x===pf)));
  const el=document.getElementById('allContent');
  if(!filtered.length){el.innerHTML='<div class="empty">해당 조건의 데이터 없음</div>';return;}
  let html=`<div class="table-wrap"><table><thead><tr><th>시간</th><th>선생님</th><th>유형</th><th>수업시간</th><th>학생</th><th>결제</th><th>상태</th></tr></thead><tbody>`;
  filtered.forEach(g=>{
    const isOv=ovKeys.has(g.선생님+'|'+g.시작.getTime());
    const payTypes=[...new Set(g.결제구분)];
    html+=`<tr>
      <td class="mono">${toHHMM(g.시작)}~${toHHMM(g.종료)}</td>
      <td>${g.선생님}</td>
      <td><span class="chip chip-${g.유형명}">${g.유형명}</span>${hapjuChipG(g)?(' '+hapjuChipG(g)):''}</td>
      <td class="mono">${g.수업분}분</td>
      <td style="font-size:12px">${g.학생들.join(', ')}</td>
      <td>${payTypes.map(p=>`<span class="badge ${p==='보강결제'?'badge-보강':'badge-info'}">${p.replace('결제','')}</span>`).join(' ')}</td>
      <td>${isOv?'<span class="badge badge-error">겹침</span>':'<span class="badge badge-ok">정상</span>'}</td>
    </tr>`;
  });
  html+='</tbody></table></div>';
  el.innerHTML=html;
}

function goTab(name){
  const btn=document.querySelector(`.tab[data-tab="${name}"]`);
  if(btn){ showTab(name,btn); }
  document.getElementById('results').scrollIntoView({behavior:'smooth',block:'start'});
}

function showTab(name,btn){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('tab-'+name).classList.add('active');
}
</script>
</body>
</html>
