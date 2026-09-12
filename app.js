/* ===================== 설정 ===================== */
const SHEET_MAIN = "ENG시트(Checklist)";
const SHEET_PHOTO = "공문 및 사진";
const TEMPLATE_URL = "template.xlsx";
const DRAFT_KEY = "engSheetDraft_v1";
const MAX_PHOTOS = 9;

// 공문 및 사진 시트의 사진 삽입 영역(9칸) + 슬롯당 제목/유형/설명 입력 셀
const PHOTO_SLOTS = [
  { range: "B41:I76",   titleCell: "D77",  typeCell: "H77",  descCell: "D78" },
  { range: "K41:R76",   titleCell: "M77",  typeCell: "Q77",  descCell: "M78" },
  { range: "T41:AA76",  titleCell: "V77",  typeCell: "Z77",  descCell: "V78" },
  { range: "B80:I115",  titleCell: "D116", typeCell: "H116", descCell: "D117" },
  { range: "K80:R115",  titleCell: "M116", typeCell: "Q116", descCell: "M117" },
  { range: "T80:AA115", titleCell: "V116", typeCell: "Z116", descCell: "V117" },
  { range: "B119:I154", titleCell: "D155", typeCell: "H155", descCell: "D156" },
  { range: "K119:R154", titleCell: "M155", typeCell: "Q155", descCell: "M156" },
  { range: "T119:AA154",titleCell: "V155", typeCell: "Z155", descCell: "V156" },
];

function pair(cCell, dCell, label){
  return [
    { cell:cCell, label:label, type:"select", options:["없음","있음"] },
    { cell:dCell, label:"상세내역", type:"text" },
  ];
}

const SECTIONS = [
  { title:"기본정보", fields:[
    { cell:"B3", label:"공사명", type:"text", required:true, placeholder:"예: 26년_취약개선_다분기접속함체개선_...", full:true },
    { cell:"B4", label:"협력사명 / 설계자(BP)", type:"text", placeholder:"예: 동서통신 / 홍길동" },
    { cell:"D4", label:"사업구분", type:"select", options:["RISK","민원해소(SKO CV)","한전위해개소","다분기/불량함체"] },
    { cell:"F4", label:"적출유형", type:"text", listId:"list-jeokchul", placeholder:"예: 01. 불량함체_함체불량" },
    { cell:"B5", label:"검토자", type:"text" },
    { cell:"D5", label:"공사방안", type:"text", placeholder:"예: 절체" },
    { cell:"F5", label:"이설요청 공문번호", type:"text", placeholder:"RM, CV, 공문번호, 한전적출번호 등" },
    { cell:"B6", label:"승인자", type:"text" },
    { cell:"D6", label:"요청주체", type:"text", placeholder:"예: SKTNS" },
    { cell:"F6", label:"현장주소", type:"text", full:true },
    { cell:"B7", label:"연계 ENG 존재유무", type:"select", options:["N","Y"] },
  ]},
  { title:"설계 요약", fields:[
    { cell:"C11", label:"이설 방법", type:"text", placeholder:"예: 절체" },
    { cell:"C12", label:"절체사유 (단순이설 불가사유)", type:"text" },
    { cell:"C13", label:"이설루트 요약", type:"text" },
    { cell:"C14", label:"이설 전 루트구성", type:"text" },
    { cell:"C15", label:"이설 후 루트구성", type:"text" },
    { cell:"C16", label:"케이블 신설 정보", type:"text", placeholder:"예: 신설 : 1조 0.625Km" },
    { cell:"D16", label:"케이블 철거 정보", type:"text", placeholder:"예: 철거 : 5조 1.942Km" },
    { cell:"C17", label:"함체 신설 정보", type:"text", placeholder:"예: 신설 : 3" },
    { cell:"D17", label:"함체 철거 정보", type:"text", placeholder:"예: 철거 : 1" },
    { cell:"C18", label:"접속코어 수 (주간)", type:"text", placeholder:"예: 주간 : 43C" },
    { cell:"D18", label:"접속코어 수 (야간)", type:"text", placeholder:"예: 야간 : 297C" },
  ]},
  { title:"공사배경 및 목적 (현장개요)", fields:[
    { type:"radiogroup", label:"요청 주체", groupCells:{
        "한전":"C21", "CV기반(민원)":"C22", "RM기반(위해)":"C23",
        "지자체":"C24", "공공기관":"C25", "기타":"C26" } },
    { cell:"D23", label:"RM등록번호 (RM기반 선택 시)", type:"text" },
    { cell:"C27", label:"병행공사 여부", type:"select", options:["없음","있음"] },
    { cell:"D27", label:"병행공사 상세 (주관사/참여사)", type:"text", placeholder:"예: 주관사: LG, 참여사: SKT" },
    { cell:"C28", label:"원인자 판정조서 첨부", type:"select", options:["첨부","미첨부"] },
    { cell:"D28", label:"원인자 불가사유", type:"text" },
  ]},
  { title:"현장실사결과", fields:[
    ...pair("C31","D31","GIS 불일치 관로·전주 존재"),
    ...pair("C32","D32","GIS 불일치 케이블·함체 존재"),
    ...pair("C33","D33","GIS 상 그 외 특이사항"),
    ...pair("C34","D34","자가주·관로 존재여부"),
    ...pair("C35","D35","타사주·관로 존재여부"),
    ...pair("C36","D36","자가주 인허가 가능여부"),
    ...pair("C37","D37","폭탄함체·다분기 함체 존재"),
  ]},
  { title:"기설정보 (작업대상)", fields:[
    { cell:"D40", label:"전주 정보", type:"text", placeholder:"예: 한전주 20본 자가주 0 타사주 16본" },
    { cell:"D41", label:"케이블 정보", type:"textarea", placeholder:"[기설케이블#1 000M] 가입자망 =0000= 00C/00C(00%)" },
    { cell:"D42", label:"함체 정보", type:"textarea", placeholder:"[기설함체#1] 한전전산번호 / SK관리번호 / 접속00C" },
    { cell:"D43", label:"기타", type:"textarea" },
  ]},
  { title:"설계 이설루트", fields:[
    { type:"checkbox", cell:"C46", label:"기설 통신주/관로 루트사용" },
    { cell:"D46", label:"상세 (기설 루트활용)", type:"text" },
    { type:"checkbox", cell:"C47", label:"병행 관로/전주 신설" },
    { type:"checkbox", cell:"C48", label:"단독 관로/전주 신설" },
    { type:"checkbox", cell:"C49", label:"한전주 이설" },
  ]},
  { title:"설계 신설정보", fields:[
    { cell:"C50", label:"절체이설 대상 조수", type:"text", placeholder:"예: 1조" },
    { cell:"D50", label:"신설케이블 정보", type:"textarea", placeholder:"[신설케이블#1 000m] 지선망 Dry000C/00C (00%)[기설#..]" },
    { cell:"C51", label:"용량증설 대상 여부", type:"select", options:["해당없음","해당"] },
    { cell:"D51", label:"용량증설 상세", type:"text" },
    { cell:"C52", label:"다대화 대상 여부", type:"select", options:["해당없음","해당"] },
    { cell:"D52", label:"다대화 상세", type:"text" },
    { cell:"C53", label:"T&B 공동투자", type:"bool", trueLabel:"공동투자", falseLabel:"SKT 단독공사" },
    { cell:"D53", label:"공동투자 상세", type:"text" },
    { cell:"C54", label:"함체 신설 대상", type:"bool", trueLabel:"있음", falseLabel:"없음" },
    { cell:"D54", label:"신설함체 정보", type:"textarea", placeholder:"[신설함체#1] 한전전산번호 접속00C_지선망 사용코아 접속" },
  ]},
  { title:"이설 후 RM여부", fields:[
    { cell:"C55", label:"6차선 횡단 구간 존재", type:"bool", trueLabel:"존재", falseLabel:"해소" },
    { cell:"D55", label:"상세 (자가주/타사주 등)", type:"text" },
    { cell:"C56", label:"임의횡단·지선없는 종말주·지상고 미달", type:"bool", trueLabel:"존재", falseLabel:"해소" },
    { cell:"D56", label:"상세/사유", type:"text" },
    { cell:"C57", label:"배전설비접촉·여유장과다·저압선이격 등", type:"bool", trueLabel:"존재", falseLabel:"해소" },
    { cell:"D57", label:"상세/사유", type:"text" },
    { cell:"C58", label:"코아링/케이블링 발생여부", type:"bool", trueLabel:"발생", falseLabel:"미발생" },
    { cell:"D58", label:"상세/사유", type:"text" },
  ]},
  { title:"지역 별 추가기준", fields:[
    { cell:"C59", label:"재활용 케이블 사용여부", type:"bool", trueLabel:"사용", falseLabel:"미사용" },
    { cell:"D59", label:"상세", type:"text" },
    { cell:"C60", label:"코아 접속수량 적정성", type:"bool", trueLabel:"부적정", falseLabel:"적정" },
    { cell:"D60", label:"상세", type:"text" },
    { cell:"C61", label:"회선분산설계 적용여부", type:"bool", trueLabel:"적용", falseLabel:"미적용" },
    { cell:"D61", label:"상세", type:"text" },
    { cell:"C62", label:"기타 지역기준 해당여부", type:"bool", trueLabel:"해당", falseLabel:"해당없음" },
    { cell:"D62", label:"상세", type:"text" },
  ]},
  { title:"특이사항", fields:[
    { cell:"B63", label:"특이사항", type:"textarea", full:true },
  ]},
];

const JEOKCHUL_LIST = [
  "01. 불량함체_함체불량","01.6차선횡단","02.불량맨홀","03.불량케이블","04.관로파손/복구/간섭등",
  "05.지상고미달(도로횡단)","06.자가주불량","07.하천횡단/교량첨가","08.단순정비",
  "02.지상고미달(도로횡단)","03.케이블민원","04.맨홀민원","05.관로민원","06.하천횡단/교량첨가",
  "07.자가주민원","08.동일지번이설","09.하천횡단/하천정비","10.사유지민원","11.건물신축민원",
  "01.행거/밴드불량","02.지상고미달","03.지장전주미이설","04.지선없는종말주","05.함체불량",
  "06.배전설비접촉/이격미달","07.한전_배전루트변경,한전시설물정비등한전기타사유","08.6차선횡단",
  "09.여유장과다","10.합동정비,소규모,단순공사등","11.기타",
];

/* ===================== 상태 ===================== */
const cellMeta = {};   // cell -> {kind:'text'|'bool'|'mark'}
const state = { cells:{}, photos:[], lastSaved:null };

/* ===================== 렌더링 ===================== */
function el(tag, cls, html){
  const e = document.createElement(tag);
  if(cls) e.className = cls;
  if(html !== undefined) e.innerHTML = html;
  return e;
}

function registerCell(cell, kind, def){
  cellMeta[cell] = kind;
  if(!(cell in state.cells)){
    state.cells[cell] = kind === "bool" ? false : (kind === "mark" ? false : "");
  }
}

function renderSections(){
  const root = document.getElementById("sections");
  root.innerHTML = "";
  let dl = document.getElementById("list-jeokchul");
  if(!dl){
    dl = document.createElement("datalist");
    dl.id = "list-jeokchul";
    JEOKCHUL_LIST.forEach(v=>{
      const o = document.createElement("option"); o.value = v; dl.appendChild(o);
    });
    document.body.appendChild(dl);
  }

  SECTIONS.forEach((sec, idx)=>{
    const secEl = el("div", "section" + (idx===0 ? " open" : ""));
    const head = el("div", "section-head");
    head.appendChild(el("h2", null, sec.title));
    head.appendChild(el("span", "chev", "▼"));
    head.onclick = ()=> secEl.classList.toggle("open");
    secEl.appendChild(head);

    const body = el("div", "section-body");
    sec.fields.forEach(f=> body.appendChild(renderField(f)));
    secEl.appendChild(body);
    root.appendChild(secEl);
  });
}

function renderField(f){
  const wrap = el("div", "field");

  if(f.type === "radiogroup"){
    Object.values(f.groupCells).forEach(c=> registerCell(c, "mark"));
    const lab = el("label", null, f.label);
    wrap.appendChild(lab);
    const row = el("div", "pillrow");
    Object.entries(f.groupCells).forEach(([name, cell])=>{
      const pill = el("div", "pill", name);
      pill.onclick = ()=>{
        Object.values(f.groupCells).forEach(c=> state.cells[c] = false);
        state.cells[cell] = true;
        row.querySelectorAll(".pill").forEach(p=>p.classList.remove("selected"));
        pill.classList.add("selected");
        scheduleSave();
      };
      if(state.cells[cell]) pill.classList.add("selected");
      row.appendChild(pill);
    });
    wrap.appendChild(row);
    return wrap;
  }

  if(f.type === "checkbox"){
    registerCell(f.cell, "mark");
    const row = el("label", "checkrow");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!state.cells[f.cell];
    cb.onchange = ()=>{ state.cells[f.cell] = cb.checked; scheduleSave(); };
    row.appendChild(cb);
    row.appendChild(el("span", null, f.label));
    wrap.appendChild(row);
    return wrap;
  }

  if(f.type === "bool"){
    registerCell(f.cell, "bool");
    wrap.appendChild(el("label", null, f.label));
    const seg = el("div", "seg");
    const bTrue = el("button", null, f.trueLabel);
    const bFalse = el("button", null, f.falseLabel);
    function refresh(){
      bTrue.className = state.cells[f.cell] ? "active a" : "";
      bFalse.className = !state.cells[f.cell] ? "active b" : "";
    }
    bTrue.onclick = ()=>{ state.cells[f.cell] = true; refresh(); scheduleSave(); };
    bFalse.onclick = ()=>{ state.cells[f.cell] = false; refresh(); scheduleSave(); };
    refresh();
    seg.appendChild(bTrue); seg.appendChild(bFalse);
    wrap.appendChild(seg);
    return wrap;
  }

  // text / textarea / select
  registerCell(f.cell, "text");
  const lab = el("label", null, f.required ? "" : "");
  lab.textContent = f.label;
  if(f.required) lab.classList.add("req");
  wrap.appendChild(lab);

  let input;
  if(f.type === "textarea"){
    input = document.createElement("textarea");
  } else if(f.type === "select"){
    input = document.createElement("select");
    const blank = document.createElement("option");
    blank.value = ""; blank.textContent = "선택";
    input.appendChild(blank);
    f.options.forEach(o=>{
      const op = document.createElement("option");
      op.value = o; op.textContent = o;
      input.appendChild(op);
    });
  } else {
    input = document.createElement("input");
    input.type = "text";
    if(f.listId) input.setAttribute("list", f.listId);
  }
  if(f.placeholder) input.placeholder = f.placeholder;
  input.value = state.cells[f.cell] || "";
  input.oninput = ()=>{ state.cells[f.cell] = input.value; scheduleSave(); };
  if(f.type === "select"){
    input.onchange = ()=>{ state.cells[f.cell] = input.value; scheduleSave(); };
  }
  wrap.appendChild(input);
  return wrap;
}

/* ===================== 저장/불러오기 ===================== */
let saveTimer = null;
function scheduleSave(){
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveDraft, 500);
}

function saveDraft(){
  state.lastSaved = new Date().toISOString();
  try{
    localStorage.setItem(DRAFT_KEY, JSON.stringify(state));
    showToast("임시저장됨");
    updateSaveStatus();
    updateSummary();
  }catch(e){
    showToast("저장 공간 부족: 사진을 줄여주세요");
  }
}

function loadDraft(){
  const raw = localStorage.getItem(DRAFT_KEY);
  if(!raw) return;
  try{
    const parsed = JSON.parse(raw);
    Object.assign(state.cells, parsed.cells || {});
    state.photos = parsed.photos || [];
    state.lastSaved = parsed.lastSaved || null;
  }catch(e){ /* ignore corrupt draft */ }
}

function updateSaveStatus(){
  const s = document.getElementById("saveStatus");
  if(state.lastSaved){
    const d = new Date(state.lastSaved);
    s.textContent = "마지막 임시저장: " + d.toLocaleString("ko-KR");
  } else {
    s.textContent = "";
  }
}

function showToast(msg){
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(showToast._timer);
  showToast._timer = setTimeout(()=> t.classList.remove("show"), 1600);
}

/* ===================== 사진 ===================== */
function resizeImage(file, maxDim){
  return new Promise((resolve, reject)=>{
    const img = new Image();
    const reader = new FileReader();
    reader.onload = ()=>{
      img.onload = ()=>{
        let { width, height } = img;
        if(width > height && width > maxDim){
          height = Math.round(height * (maxDim/width)); width = maxDim;
        } else if(height >= width && height > maxDim){
          width = Math.round(width * (maxDim/height)); height = maxDim;
        }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", 0.82));
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function renderPhotos(){
  const list = document.getElementById("photoList");
  list.innerHTML = "";
  document.getElementById("photoCount").textContent =
    `${state.photos.length} / ${MAX_PHOTOS}장 첨부됨`;
  document.getElementById("photoAddBtn").style.display =
    state.photos.length >= MAX_PHOTOS ? "none" : "flex";

  state.photos.forEach((p, i)=>{
    const item = el("div", "photo-item");
    const img = document.createElement("img");
    img.src = p.dataUrl;
    item.appendChild(img);
    const body = el("div", "pbody");

    const t = document.createElement("input");
    t.type = "text"; t.placeholder = "제목"; t.value = p.title || "";
    t.oninput = ()=>{ p.title = t.value; scheduleSave(); };

    const row1 = el("div", "prow");
    const typ = document.createElement("input");
    typ.type = "text"; typ.placeholder = "사진 유형 (예: 근거사진)"; typ.value = p.ptype || "";
    typ.oninput = ()=>{ p.ptype = typ.value; scheduleSave(); };
    const ref = document.createElement("input");
    ref.type = "text"; ref.placeholder = "관련 근거"; ref.value = p.ref || "";
    ref.oninput = ()=>{ p.ref = ref.value; scheduleSave(); };
    row1.appendChild(typ); row1.appendChild(ref);

    body.appendChild(t);
    body.appendChild(row1);

    const rm = el("button", "premove", "삭제");
    rm.onclick = ()=>{
      state.photos.splice(i,1);
      renderPhotos();
      scheduleSave();
    };
    body.appendChild(rm);
    item.appendChild(body);
    list.appendChild(item);
  });
}

document.getElementById("photoAddBtn").addEventListener("click", ()=>{
  document.getElementById("photoInput").click();
});

document.getElementById("photoInput").addEventListener("change", async (e)=>{
  const files = Array.from(e.target.files || []);
  for(const file of files){
    if(state.photos.length >= MAX_PHOTOS){
      showToast(`사진은 최대 ${MAX_PHOTOS}장까지 첨부할 수 있어요`);
      break;
    }
    try{
      const dataUrl = await resizeImage(file, 1400);
      state.photos.push({ dataUrl, title:"", ptype:"", ref:"" });
    }catch(err){ /* skip file on error */ }
  }
  e.target.value = "";
  renderPhotos();
  scheduleSave();
});

/* ===================== 요약/탭 ===================== */
function updateSummary(){
  document.getElementById("sumName").textContent = state.cells["B3"] || "-";
  document.getElementById("sumAddr").textContent = state.cells["F6"] || "-";
  document.getElementById("sumPhotos").textContent = state.photos.length + "장";
  document.getElementById("sumSaved").textContent = state.lastSaved
    ? new Date(state.lastSaved).toLocaleString("ko-KR") : "-";
}

document.querySelectorAll(".navbar button").forEach(btn=>{
  btn.addEventListener("click", ()=>{
    document.querySelectorAll(".navbar button").forEach(b=>b.classList.remove("active"));
    btn.classList.add("active");
    document.querySelectorAll(".tabpanel").forEach(p=>p.classList.remove("active"));
    document.getElementById(btn.dataset.tab).classList.add("active");
    if(btn.dataset.tab === "tab-done") updateSummary();
  });
});

/* ===================== 초기화 ===================== */
document.getElementById("btnSaveNow").addEventListener("click", saveDraft);

document.getElementById("btnReset").addEventListener("click", ()=>{
  if(!confirm("입력한 모든 내용과 사진이 삭제됩니다. 새 건으로 시작할까요?")) return;
  Object.keys(cellMeta).forEach(c=>{
    state.cells[c] = cellMeta[c] === "text" ? "" : false;
  });
  state.photos = [];
  state.lastSaved = null;
  localStorage.removeItem(DRAFT_KEY);
  renderSections();
  renderPhotos();
  updateSummary();
  updateSaveStatus();
  showToast("초기화되었습니다");
});

/* ===================== 엑셀 내보내기 ===================== */
function dataUrlToArrayBuffer(dataUrl){
  const base64 = dataUrl.split(",")[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for(let i=0;i<binary.length;i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

async function exportExcel(){
  if(!state.cells["B3"]){
    alert("공사명을 입력해 주세요.");
    document.querySelector('.navbar button[data-tab="tab-input"]').click();
    return;
  }
  showToast("엑셀 생성 중...");
  try{
    const resp = await fetch(TEMPLATE_URL);
    const buf = await resp.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);

    const wsMain = wb.getWorksheet(SHEET_MAIN);
    Object.entries(cellMeta).forEach(([cell, kind])=>{
      const val = state.cells[cell];
      const target = wsMain.getCell(cell);
      if(kind === "mark"){
        target.value = val ? "O" : "";
      } else if(kind === "bool"){
        target.value = !!val;
      } else {
        target.value = val || "";
      }
    });

    if(state.photos.length){
      const wsPhoto = wb.getWorksheet(SHEET_PHOTO);
      state.photos.slice(0, MAX_PHOTOS).forEach((p, i)=>{
        const slot = PHOTO_SLOTS[i];
        const imgId = wb.addImage({ base64: p.dataUrl, extension: "jpeg" });
        wsPhoto.addImage(imgId, slot.range);
        wsPhoto.getCell(slot.titleCell).value = p.title || "";
        wsPhoto.getCell(slot.typeCell).value = p.ptype || "";
        wsPhoto.getCell(slot.descCell).value = p.ref || "";
      });
    }

    const outBuf = await wb.xlsx.writeBuffer();
    const blob = new Blob([outBuf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const safeName = (state.cells["B3"] || "현장ENG").replace(/[\\/:*?"<>|]/g, "_").slice(0, 60);
    const today = new Date().toISOString().slice(0,10).replace(/-/g,"");
    const filename = `${safeName}_${today}.xlsx`;

    await deliverFile(blob, filename);
  }catch(err){
    console.error(err);
    alert("엑셀 생성 중 오류가 발생했습니다: " + err.message);
  }
}

/* 네이버/카카오 등 인앱 브라우저는 blob 다운로드를 막는 경우가 많아
   여러 방식을 순서대로 시도하고, 항상 수동으로 열 수 있는 링크도 남겨둔다. */
async function deliverFile(blob, filename){
  const url = URL.createObjectURL(blob);

  // 수동 폴백 링크는 성공 여부와 무관하게 항상 준비해둔다.
  const fb = document.getElementById("exportFallback");
  if(fb){
    fb.innerHTML = "";
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.target = "_blank";
    link.textContent = "⬇ 다운로드가 안 되면 여기를 눌러주세요";
    link.style.display = "inline-block";
    link.style.marginTop = "4px";
    link.style.color = "#0F2A4A";
    link.style.fontWeight = "700";
    link.style.textDecoration = "underline";
    fb.appendChild(link);
  }

  // 1) Web Share API (파일 공유 지원 브라우저) - 인앱 브라우저에서 가장 안정적
  try{
    const file = new File([blob], filename, { type: blob.type });
    if(navigator.canShare && navigator.canShare({ files:[file] })){
      await navigator.share({ files:[file], title: filename });
      showToast("공유 시트에서 저장을 완료해주세요");
      return;
    }
  }catch(shareErr){
    // 사용자가 공유를 취소한 경우 등은 무시하고 다음 방식 시도
  }

  // 2) 표준 a[download] 클릭
  try{
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    showToast("엑셀 다운로드를 시작했어요");
  }catch(clickErr){
    // 3) 새 탭에서 열기 (여기서도 막히면 위의 수동 링크를 눌러야 함)
    try{
      window.open(url, "_blank");
      showToast("새 탭에서 파일을 열었어요");
    }catch(openErr){
      showToast("자동 다운로드 실패 - 아래 링크를 눌러주세요");
    }
  }
  setTimeout(()=> URL.revokeObjectURL(url), 60000);
}

document.getElementById("btnExport").addEventListener("click", exportExcel);

/* ===================== 시작 ===================== */
loadDraft();
renderSections();
renderPhotos();
updateSummary();
updateSaveStatus();

if("serviceWorker" in navigator){
  navigator.serviceWorker.register("service-worker.js").catch(()=>{});
}
