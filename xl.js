/* 원드라이브 엑셀 연동 (Microsoft Graph)
   - 로그인: 개인 Microsoft 계정 (MSAL, 리디렉션 방식: 아이폰 홈화면 앱에서도 동작)
   - 읽기: 퇴직연금 / 배당내역 / 입금내역 / ETF시세 시트
   - 쓰기: 매매·배당·입금 행 추가, 잔여 금액, 시세, 보유/매도 종목 목록 자동 정리
   개인 데이터는 이 파일이나 저장소에 저장되지 않음 (브라우저에서 Graph로 직접 주고받음) */
const XL = (() => {
  const CLIENT_ID = 'c37e15f3-c0d2-47d2-b031-ccbc894a7a8d';
  const SCOPES = ['Files.ReadWrite', 'User.Read'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const INTEREST = '예수금 이자'; // 배당내역 시트에 이 종목명으로 이자를 기록
  const S = { trade: '퇴직연금', div: '배당내역', dep: '입금내역', price: 'ETF시세' };
  const store = {
    get(k, d) { try { return localStorage.getItem(k) ?? d; } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch {} },
    del(k) { try { localStorage.removeItem(k); } catch {} }
  };
  let app, account, itemId, meta = {};

  /* ---------- 날짜: 엑셀 일련번호 <-> YYYY-MM-DD ---------- */
  const EPOCH = Date.UTC(1899, 11, 30);
  const toSerial = d => { const [y, m, dd] = d.split('-').map(Number); return (Date.UTC(y, m - 1, dd) - EPOCH) / 864e5; };
  const fromSerial = n => new Date(EPOCH + Math.floor(n + 1e-6) * 864e5).toISOString().slice(0, 10);
  const isNum = v => typeof v === 'number' && !isNaN(v);
  const colOf = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  /* ---------- 로그인 ---------- */
  async function init() {
    if (!window.msal) throw new Error('로그인 모듈을 불러오지 못했어요. 인터넷 연결을 확인해 주세요.');
    app = new msal.PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/consumers', redirectUri: location.origin + location.pathname.replace(/index\.html$/, '') },
      cache: { cacheLocation: 'localStorage' }
    });
    await app.initialize();
    const res = await app.handleRedirectPromise();
    account = res?.account || app.getAllAccounts()[0] || null;
    if (account) app.setActiveAccount(account);
    return !!account;
  }
  const login = () => app.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
  const logout = () => { store.del('pensionCache'); store.del('xlItemId'); return app.logoutRedirect({ account }); };
  const user = () => account ? (account.name || account.username) : '';
  async function token() {
    try { return (await app.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { await app.acquireTokenRedirect({ scopes: SCOPES, account }); }
      throw e;
    }
  }

  /* ---------- Graph 호출 ---------- */
  async function g(method, url, body, session) {
    const headers = { Authorization: 'Bearer ' + await token(), 'Content-Type': 'application/json' };
    if (session) headers['workbook-session-id'] = session;
    for (let attempt = 0; attempt < 3; attempt++) {
      const r = await fetch(url.startsWith('http') ? url : GRAPH + url, { method, headers, body: body ? JSON.stringify(body) : undefined });
      const retriable = r.status === 429 || (method === 'GET' && (r.status === 503 || r.status === 504));
      if (retriable && attempt < 2) { await new Promise(ok => setTimeout(ok, 1500 * (attempt + 1))); continue; }
      if (!r.ok) {
        let msg = r.status + ''; try { const j = await r.json(); msg = j.error?.message || msg; } catch {}
        if (r.status === 504 || r.status === 503) msg = '마이크로소프트 서버 응답이 늦어요. 새로고침(↻)으로 저장됐는지 확인한 뒤 다시 시도해 주세요.';
        if (r.status === 404) msg = '엑셀 파일을 찾지 못했어요. 설정에서 파일 위치를 확인해 주세요. (' + msg + ')';
        if (r.status === 423 || /locked/i.test(msg)) msg = '엑셀 파일이 잠겨 있어요. PC에서 파일을 닫거나 잠시 후 다시 시도해 주세요.';
        throw new Error(msg);
      }
      return r.status === 204 ? null : r.json().catch(() => null);
    }
    throw new Error('마이크로소프트 서버가 바빠요. 잠시 후 다시 시도해 주세요.');
  }
  const path = () => store.get('xlPath', '주식/0.신_매매관리.xlsx');
  const setPath = p => { store.set('xlPath', p.replace(/^\/+/, '')); store.del('xlItemId'); itemId = null; };
  async function item() {
    if (itemId) return itemId;
    itemId = store.get('xlItemId', null);
    if (itemId) return itemId;
    const enc = path().split('/').map(encodeURIComponent).join('/');
    const j = await g('GET', `/me/drive/root:/${enc}?$select=id,name`);
    itemId = j.id; store.set('xlItemId', itemId);
    return itemId;
  }
  const wb = async () => `/me/drive/items/${await item()}/workbook`;
  const rng = async (sheet, addr) => `${await wb()}/worksheets/${encodeURIComponent(sheet)}/range(address='${addr}')`;
  const read = async (sheet, addr, sel = 'values') => (await g('GET', `${await rng(sheet, addr)}?$select=${sel}`));
  const patch = async (sheet, addr, body, ss) => g('PATCH', await rng(sheet, addr), body, ss);
  const insert = async (sheet, addr, ss) => g('POST', `${await rng(sheet, addr)}/insert`, { shift: 'Down' }, ss);
  const remove = async (sheet, addr, ss) => g('POST', `${await rng(sheet, addr)}/delete`, { shift: 'Up' }, ss);
  async function withSession(fn) {
    const s = await g('POST', `${await wb()}/createSession`, { persistChanges: true });
    try { return await fn(s.id); }
    finally { await g('POST', `${await wb()}/closeSession`, null, s.id).catch(() => {}); }
  }

  /* ---------- 읽기 ---------- */
  // 퇴직연금 시트 T~AN열 블록: 보유종목(T4~), 매도내역(미보유), 요약표(입금 금액, 잔여 금액 등)
  function parseBlock(vals, forms) {
    const T = r => String(vals[r - 1]?.[0] ?? '').trim();
    const cell = (r, col) => vals[r - 1]?.[col - 20];
    const rowOf = label => { for (let r = 1; r <= vals.length; r++) if (T(r) === label) return r; return 0; };
    const hold = []; let r = 4;
    for (; r <= vals.length && T(r) && !T(r).startsWith('매도내역'); r++) hold.push(r);
    const soldHeader = (() => { for (let i = r; i <= vals.length; i++) if (T(i).startsWith('매도내역')) return i; return 0; })();
    const sold = []; if (soldHeader) for (let i = soldHeader + 1; i <= vals.length && T(i); i++) sold.push(i);
    const depRow = rowOf('입금 금액'), cashRow = rowOf('잔여 금액');
    const stock = rr => ({ n: T(rr), fee: cell(rr, 21), target: isNum(cell(rr, 23)) ? cell(rr, 23) : 0, price: isNum(cell(rr, 28)) ? cell(rr, 28) : 0, row: rr, aa: String(forms[rr - 1]?.[7] ?? ''), ah: String(forms[rr - 1]?.[14] ?? ''), aj: String(forms[rr - 1]?.[16] ?? '') });
    return {
      hold: hold.map(stock), sold: sold.map(stock), soldHeader,
      holdStart: 4, holdEnd: hold.length ? hold[hold.length - 1] : 3, soldEnd: sold.length ? sold[sold.length - 1] : soldHeader,
      depRow, depFormula: depRow ? forms[depRow - 1]?.[1] : '', cashRow, cash: cashRow ? cell(cashRow, 21) : 0
    };
  }
  async function load() {
    try { return await loadOnce(); }
    catch (e) { if (!store.get('xlItemId', null)) throw e; store.del('xlItemId'); itemId = null; return loadOnce(); }
  }
  async function loadOnce() {
    const [tr, blk, dv, dp, px] = await Promise.all([
      read(S.trade, 'B4:H600'),
      read(S.trade, 'T1:AN60', 'values,formulas'),
      read(S.div, 'B4:H1500'),
      read(S.dep, 'B4:D300'),
      read(S.price, 'A2:C2500')
    ]);
    const codeOf = {}, priceRow = {};
    let lastPriceRow = 1;
    px.values.forEach((v, i) => { if (v[1]) { lastPriceRow = i + 2; codeOf[String(v[1]).trim()] = String(v[0]).trim(); priceRow[String(v[0]).trim()] = { row: i + 2, c: v[2] }; } });
    const tx = []; let lastTx = 3;
    tr.values.forEach((v, i) => {
      if (isNum(v[0]) && v[2]) { lastTx = i + 4; tx.push({ row: i + 4, d: fromSerial(v[0]), t: v[1] === '매도' ? 'sell' : 'buy', n: String(v[2]).trim(), q: +v[3] || 0, a: +v[5] || 0, pl: +v[6] || 0 }); }
    });
    const div = [], int = []; let lastDiv = 3, lastPensionDiv = 0;
    dv.values.forEach((v, i) => {
      if (isNum(v[0]) || v[2]) lastDiv = i + 4;
      if (isNum(v[0]) && v[1] === '퇴직연금' && String(v[2]).trim() === INTEREST) int.push({ row: i + 4, d: fromSerial(v[0]), n: INTEREST, a: +v[6] || 0 });
      else if (isNum(v[0]) && v[1] === '퇴직연금') { lastPensionDiv = i + 4; div.push({ row: i + 4, d: fromSerial(v[0]), n: String(v[2]).trim(), q: +v[3] || 0, cost: +v[4] || 0, y: isNum(v[5]) ? v[5] : null, a: +v[6] || 0 }); }
    });
    const dep = []; let lastDep = 3;
    dp.values.forEach((v, i) => { if (isNum(v[0]) && isNum(v[2])) { lastDep = i + 4; dep.push({ row: i + 4, d: fromSerial(v[0]), a: v[2] }); } });
    const b = parseBlock(blk.values, blk.formulas);
    const addCode = s => ({ ...s, code: codeOf[s.n] || null });
    meta = { lastTx, lastDiv, lastPensionDiv, lastDep, block: b, priceRow, codeOf, lastPriceRow };
    const data = {
      asof: new Date(Date.now() + 9 * 3600e3).toISOString().slice(0, 10), cash: +b.cash || 0,
      hold: b.hold.map(addCode), sold: b.sold.map(addCode), tx, div, int, dep,
      etfNames: Object.keys(codeOf), live: true, loadedAt: new Date().toISOString()
    };
    store.set('pensionCache', JSON.stringify(data));
    return data;
  }
  const cached = () => { try { return JSON.parse(store.get('pensionCache', 'null')); } catch { return null; } };

  /* ---------- 쓰기 ---------- */
  /* 엑셀 수식 자동 정비 (필요할 때만 1회씩). 바꾼 내용을 문장 목록으로 돌려줌
     1) 입금 금액 → 입금내역 합계에 연결
     2) 매수 총액(AA열): "가장 최근 날짜의 누적값 중 최댓값" → "그 종목의 마지막 행 누적값"
        (같은 날 같은 종목을 여러 번 매매해도 맞게 나옴. 앱이 매매를 날짜 순서로 넣기 때문에 마지막 행 = 최신 상태)
     3) 배당내역 총매수(F열): 같은 방식으로, 지급일까지의 마지막 행 누적값
     4) 종목별 수익율(AJ열): 총수익 ÷ 현재 매수총액 → 총수익 ÷ 누적 매수금액
        (일부를 판 종목이 부풀려지지 않고, 전량 매도 종목도 계산됨) */
  const AA_NEW = r => `=IFERROR(LOOKUP(2,1/($D$4:$D$5000=$T${r}),$I$4:$I$5000),0)`;
  const DIVF_NEW = r => `=IFERROR(LOOKUP(2,1/((퇴직연금!$D$4:$D$5000=$D${r})*(퇴직연금!$B$4:$B$5000<=$B${r})),퇴직연금!$I$4:$I$5000),0)`;
  const AJ_NEW = r => `=IFERROR(AH${r}/SUMIFS($G:$G,$D:$D,$T${r},$C:$C,"매수"),0)`;
  async function ensureFormulas() {
    const done = [], b = meta.block; if (!b) return done;
    if (b.depRow && String(b.depFormula).replace(/\s/g, '') !== '=입금내역!D3') {
      await patch(S.trade, 'U' + b.depRow, { formulas: [['=입금내역!D3']] });
      done.push('입금 금액을 입금내역 합계에 연결했어요');
    }
    if (b.hold.some(h => /MAXIFS/i.test(h.aa))) {
      await withSession(async ss => {
        for (const h of b.hold) if (/MAXIFS/i.test(h.aa)) await patch(S.trade, `AA${h.row}`, { formulas: [[AA_NEW(h.row)]] }, ss);
      });
      done.push('매수 총액 수식(AA열)을 마지막 행 기준으로 바꿨어요');
    }
    const ajOld = h => /^=\s*AH\d+\s*\/\s*AA\d+\s*$/i.test(h.aj);
    const ajHold = b.hold.filter(ajOld), ajSold = b.sold.filter(h => (!h.aj || ajOld(h)) && h.ah);
    if (ajHold.length || ajSold.length) {
      await withSession(async ss => {
        for (const h of ajHold) await patch(S.trade, `AJ${h.row}`, { formulas: [[AJ_NEW(h.row)]] }, ss);
        if (ajSold.length && b.hold.length) {
          const nf = (await sget(S.trade, `AJ${b.hold[0].row}`, 'numberFormat', ss)).numberFormat;
          for (const h of ajSold) await patch(S.trade, `AJ${h.row}`, { formulas: [[AJ_NEW(h.row)]], numberFormat: nf }, ss);
        }
      });
      done.push('종목별 수익율 수식(AJ열)을 누적 매수금액 기준으로 바꿨어요');
    }
    if (meta.lastPensionDiv && store.get('divFormulaOk', '') !== '1') {
      const last = meta.lastDiv, cf = (await read(S.div, `C4:F${last}`, 'formulas')).formulas;
      const col = cf.map((r, i) => [r[0] === '퇴직연금' && /MAXIFS/i.test(String(r[3])) ? DIVF_NEW(i + 4) : r[3]]);
      if (col.some((r, i) => r[0] !== cf[i][3])) {
        await patch(S.div, `F4:F${last}`, { formulas: col });
        done.push('배당내역 총매수 수식(F열)을 마지막 행 기준으로 바꿨어요');
      }
      store.set('divFormulaOk', '1');
    }
    return done;
  }
  async function setCash(v) {
    if (!meta.block?.cashRow) throw new Error('엑셀에서 "잔여 금액" 칸을 찾지 못했어요.');
    await patch(S.trade, 'U' + meta.block.cashRow, { values: [[v]] });
  }
  async function copyFormat(sheet, fromAddr) { return (await read(sheet, fromAddr, 'numberFormat')).numberFormat; }

  /* 저장 직전에 엑셀의 실제 마지막 행을 다시 읽음 (앱을 켠 뒤 PC에서 엑셀에 행을 추가했어도 덮어쓰지 않도록) */
  const sget = async (sheet, addr, sel, ss) => g('GET', `${await rng(sheet, addr)}?$select=${sel}`, null, ss);
  async function freshTrades(ss) {
    const v = (await sget(S.trade, 'B4:D600', 'values', ss)).values; let last = 3; const rows = [];
    v.forEach((r, i) => { if (isNum(r[0]) && r[2]) { last = i + 4; rows.push({ row: i + 4, s: Math.floor(r[0] + 1e-6) }); } });
    return { last, rows };
  }
  async function freshDiv(ss) {
    const v = (await sget(S.div, 'B4:D1500', 'values', ss)).values; let last = 3, lastPension = 0;
    v.forEach((r, i) => { if (isNum(r[0]) || r[2]) last = i + 4; if (isNum(r[0]) && r[1] === '퇴직연금' && String(r[2]).trim() !== INTEREST) lastPension = i + 4; });
    return { last, lastPension };
  }
  async function freshDep(ss) {
    const v = (await sget(S.dep, 'B4:D300', 'values', ss)).values; let last = 3;
    v.forEach((r, i) => { if (isNum(r[0]) && isNum(r[2])) last = i + 4; });
    return last;
  }
  const tradeRow = (x, r) => [toSerial(x.d), x.t === 'sell' ? '매도' : '매수', x.n, x.q, `=G${r}/E${r}`, x.a, x.t === 'sell' ? x.pl : ''];

  // 매매 한 줄 넣기: 날짜 순서에 맞는 자리에 넣음.
  // (엑셀의 매수총액은 "그 종목의 가장 최근 날짜 행"의 누적값을 읽으므로, 지난 날짜 기록을 맨 아래에 붙이면 계산에서 빠짐)
  async function insertTrade(x, ss) {
    const { last, rows } = await freshTrades(ss);
    const after = rows.find(r => r.s > toSerial(x.d));
    const pos = after ? after.row : last + 1, prev = Math.max(4, last), end = last + 1;
    const nf = (await sget(S.trade, `B${prev}:H${prev}`, 'numberFormat', ss)).numberFormat;
    if (after) {
      const f = (await sget(S.trade, `B${pos}:H${last}`, 'formulasR1C1', ss)).formulasR1C1;
      await patch(S.trade, `B${pos + 1}:H${end}`, { formulasR1C1: f }, ss);
      await patch(S.trade, `B${end}:H${end}`, { numberFormat: nf }, ss);
      await patch(S.trade, `B${pos}:H${pos}`, { formulas: [tradeRow(x, pos)] }, ss);
    } else {
      await patch(S.trade, `B${pos}:H${pos}`, { formulas: [tradeRow(x, pos)], numberFormat: nf }, ss);
    }
    // A열(종목별 순번)·I열(매수총액 누적) 수식이 준비된 행을 넘어가면 윗줄 수식을 이어 붙임
    if (end > 4 && !(await sget(S.trade, `I${end}`, 'formulas', ss)).formulas[0][0]) {
      for (const col of ['A', 'I']) {
        const f = (await sget(S.trade, `${col}${end - 1}`, 'formulasR1C1', ss)).formulasR1C1;
        if (f[0][0]) await patch(S.trade, `${col}${end}`, { formulasR1C1: f }, ss);
      }
    }
  }
  async function addTrade(x, D) {
    await withSession(async ss => {
      if (x.t === 'buy') await ensureEtfRow(x.n, ss);
      await insertTrade(x, ss);
      await reconcileLists([...D.tx, x], [x.n], ss);
    });
  }

  // 보유종목 / 매도내역 목록을 매매내역에 맞게 정리 (한 번에 하나씩 고치고, 행 번호가 바뀌므로 매번 다시 읽음)
  //  - 보유 수량이 있는데 보유 목록에 없으면 → 추가
  //  - 보유 목록에 있는데 수량이 0이면 → 매도내역으로 이동 (매매 기록이 아예 없으면 목록에서 제거)
  //  - 이번에 기록을 넣거나 고친 종목만 대상. 엑셀에 미리 적어둔 다른 종목은 건드리지 않음
  async function reconcileLists(txAfter, touchedNames, ss) {
    const touched = new Set(touchedNames);
    const qty = {}, has = {};
    txAfter.forEach(t => { qty[t.n] = (qty[t.n] || 0) + (t.t === 'buy' ? t.q : -t.q); has[t.n] = true; });
    const R1 = async addr => (await g('GET', `${await rng(S.trade, addr)}?$select=formulasR1C1`, null, ss)).formulasR1C1;
    for (let guard = 0; guard < 8; guard++) {
      const blk = await g('GET', `${await rng(S.trade, 'T1:AN60')}?$select=values,formulas`, null, ss);
      const b = parseBlock(blk.values, blk.formulas);
      const inHold = n => b.hold.find(s => s.n === n), inSold = n => b.sold.find(s => s.n === n);
      if (!b.hold.length) break; // 보유 목록 형태를 알 수 없으면 손대지 않음
      const toAdd = [...touched].find(n => qty[n] > 0 && !inHold(n));
      const zero = b.hold.find(s => touched.has(s.n) && !(qty[s.n] > 0));
      const ghost = b.sold.find(s => touched.has(s.n) && !has[s.n]);
      if (toAdd) {
        const old = inSold(toAdd);
        if (old) { await remove(S.trade, `T${old.row}:AN${old.row}`, ss); continue; } // 다음 바퀴에서 보유 목록에 추가
        // 보유 목록 마지막 행 위치에 칸 삽입 (합계 수식 범위 안쪽이라 SUM·드롭다운 범위가 자동으로 늘어남)
        const at = b.holdEnd;
        await insert(S.trade, `T${at}:AN${at}`, ss);
        const row = (await R1(`T${at + 1}:AN${at + 1}`))[0].slice();
        row[0] = toAdd; row[3] = meta.block?.sold.find(s => s.n === toAdd)?.target || 0; // T=종목명, W=관리비중(새 종목은 0)
        await patch(S.trade, `T${at}:AN${at}`, { formulasR1C1: [row] }, ss);
      } else if (zero) {
        if (has[zero.n] && b.soldHeader) {
          // 전량 매도 → "매도내역 (미보유)" 끝에 추가한 뒤 보유 목록에서 삭제
          const at = b.soldEnd + 1;
          await insert(S.trade, `T${at}:AN${at}`, ss);
          const src = b.sold.length ? b.soldEnd : zero.row;
          const row = (await R1(`T${src}:AN${src}`))[0].slice();
          row[0] = zero.n; row[3] = zero.target || 0;
          await patch(S.trade, `T${at}:AN${at}`, { formulasR1C1: [row] }, ss);
        }
        if (b.hold.length > 1) await remove(S.trade, `T${zero.row}:AN${zero.row}`, ss); else break;
      } else if (ghost) {
        await remove(S.trade, `T${ghost.row}:AN${ghost.row}`, ss);
      } else break;
    }
  }

  /* ---------- 수정 / 삭제 ---------- */
  // 저장 전에 엑셀의 그 행이 앱이 알고 있는 기록과 같은지 확인 (다른 곳에서 엑셀을 고친 경우 엉뚱한 행을 건드리지 않도록)
  async function verify(sheet, addr, expect, ss) {
    const v = (await g('GET', `${await rng(sheet, addr)}?$select=values`, null, ss)).values[0];
    const same = expect.every((e, i) => e === undefined || (isNum(e) ? Math.abs((+v[i] || 0) - e) < 0.5 : String(v[i] ?? '').trim() === e));
    if (!same) throw new Error('엑셀 내용이 앱과 달라요. 오른쪽 위 새로고침(↻) 후 다시 시도해 주세요.');
  }
  // 한 행 삭제: 칸을 지우지 않고 아래 내용을 한 줄씩 위로 올림 (다른 칸의 수식 참조가 깨지지 않음)
  async function shiftUp(sheet, c1, c2, row, last, ss) {
    if (last > row) {
      const f = (await g('GET', `${await rng(sheet, `${c1}${row + 1}:${c2}${last}`)}?$select=formulasR1C1`, null, ss)).formulasR1C1;
      await patch(sheet, `${c1}${row}:${c2}${last - 1}`, { formulasR1C1: f }, ss);
    }
    await g('POST', `${await rng(sheet, `${c1}${last}:${c2}${last}`)}/clear`, { applyTo: 'Contents' }, ss);
  }
  const txExpect = o => [toSerial(o.d), o.t === 'sell' ? '매도' : '매수', o.n, o.q, undefined, o.a];
  async function editTrade(o, x, D) {
    await withSession(async ss => {
      await verify(S.trade, `B${o.row}:G${o.row}`, txExpect(o), ss);
      if (x.d === o.d) await patch(S.trade, `B${o.row}:H${o.row}`, { formulas: [tradeRow(x, o.row)] }, ss);
      else { // 날짜가 바뀌면 빼냈다가 날짜 순서에 맞는 자리에 다시 넣음
        await shiftUp(S.trade, 'B', 'H', o.row, (await freshTrades(ss)).last, ss);
        await insertTrade(x, ss);
      }
      if (x.t === 'buy') await ensureEtfRow(x.n, ss);
      await reconcileLists(D.tx.map(t => t.row === o.row ? { ...t, ...x } : t), [o.n, x.n], ss);
    });
  }
  async function deleteTrade(o, D) {
    await withSession(async ss => {
      await verify(S.trade, `B${o.row}:G${o.row}`, txExpect(o), ss);
      await shiftUp(S.trade, 'B', 'H', o.row, (await freshTrades(ss)).last, ss);
      await reconcileLists(D.tx.filter(t => t.row !== o.row), [o.n], ss);
    });
  }
  const divExpect = o => [toSerial(o.d), '퇴직연금', o.n, undefined, undefined, undefined, o.a];
  async function editDividend(o, x) {
    await withSession(async ss => {
      await verify(S.div, `B${o.row}:H${o.row}`, divExpect(o), ss);
      await patch(S.div, `B${o.row}`, { values: [[toSerial(x.d)]] }, ss);
      await patch(S.div, `D${o.row}`, { values: [[x.n]] }, ss);
      await patch(S.div, `H${o.row}`, { values: [[x.a]] }, ss);
    });
  }
  async function deleteDividend(o) {
    await withSession(async ss => {
      await verify(S.div, `B${o.row}:H${o.row}`, divExpect(o), ss);
      await shiftUp(S.div, 'B', 'I', o.row, (await freshDiv(ss)).last, ss);
    });
  }
  const depExpect = o => [toSerial(o.d), undefined, o.a];
  async function editDeposit(o, x) {
    await withSession(async ss => {
      await verify(S.dep, `B${o.row}:D${o.row}`, depExpect(o), ss);
      await patch(S.dep, `B${o.row}`, { values: [[toSerial(x.d)]] }, ss);
      await patch(S.dep, `D${o.row}`, { values: [[x.a]] }, ss);
    });
  }
  async function deleteDeposit(o) {
    await withSession(async ss => {
      await verify(S.dep, `B${o.row}:D${o.row}`, depExpect(o), ss);
      await shiftUp(S.dep, 'B', 'D', o.row, await freshDep(ss), ss); // E·F열(연평균 수익률 계산용 수식)은 그대로 둠
    });
  }
  async function addDividend(x) {
    await withSession(async ss => {
      const fd = await freshDiv(ss), r = fd.last + 1, tplRow = fd.lastPension || fd.last;
      const [nf, f] = await Promise.all([copyFormat(S.div, `B${tplRow}:H${tplRow}`), read(S.div, `E${tplRow}:G${tplRow}`, 'formulasR1C1')]);
      await patch(S.div, `B${r}:D${r}`, { values: [[toSerial(x.d), '퇴직연금', x.n]], numberFormat: [nf[0].slice(0, 3)] }, ss);
      await patch(S.div, `E${r}:G${r}`, { formulasR1C1: f.formulasR1C1, numberFormat: [nf[0].slice(3, 6)] }, ss);
      await patch(S.div, `H${r}`, { values: [[x.a]], numberFormat: [[nf[0][6]]] }, ss);
    });
  }
  // 이자: 배당내역 시트에 종목명 '예수금 이자'로 한 줄 (보유수량·총매수·배당률 칸은 비움)
  async function addInterest(x) {
    await withSession(async ss => {
      const fd = await freshDiv(ss), r = fd.last + 1, tplRow = fd.lastPension || fd.last;
      const nf = (await sget(S.div, `B${tplRow}:H${tplRow}`, 'numberFormat', ss)).numberFormat;
      await patch(S.div, `B${r}:D${r}`, { values: [[toSerial(x.d), '퇴직연금', INTEREST]], numberFormat: [nf[0].slice(0, 3)] }, ss);
      await patch(S.div, `H${r}`, { values: [[x.a]], numberFormat: [[nf[0][6]]] }, ss);
    });
  }
  // 잔여 금액을 delta만큼 더하거나 뺌. 보유 목록이 바뀌면 칸 위치가 달라지므로 위치를 다시 찾음
  async function adjustCash(delta) {
    if (!delta) return;
    await withSession(async ss => {
      const blk = await sget(S.trade, 'T1:AN60', 'values,formulas', ss), b = parseBlock(blk.values, blk.formulas);
      if (!b.cashRow) throw new Error('엑셀에서 "잔여 금액" 칸을 찾지 못했어요.');
      if (String(blk.formulas[b.cashRow - 1][1]).startsWith('=')) return; // 수식이 들어 있으면 건드리지 않음
      await patch(S.trade, 'U' + b.cashRow, { values: [[(+b.cash || 0) + delta]] }, ss);
    });
  }
  async function addDeposit(x) {
    await withSession(async ss => {
      const lastDep = await freshDep(ss), r = lastDep + 1, prev = Math.max(4, lastDep);
      const nf = await copyFormat(S.dep, `B${prev}:D${prev}`);
      await patch(S.dep, `B${r}:D${r}`, { values: [[toSerial(x.d), '퇴직연금', x.a]], numberFormat: nf }, ss);
      // 연평균 수익률(XIRR) 계산용 E·F열 수식이 다음 행까지 있어야 함
      const ef = (await read(S.dep, `E${r}:F${r + 1}`, 'formulas')).formulas;
      if (!ef[0][0] || !ef[1][0]) {
        const tpl = (await read(S.dep, `E${prev}:F${prev}`, 'formulasR1C1')).formulasR1C1;
        await patch(S.dep, `E${r}:F${r + 1}`, { formulasR1C1: [tpl[0], tpl[0]] }, ss);
      }
    });
  }
  /* KRX 시세를 ETF시세 시트에 반영: 종가(C열)와 시가총액(L열), 시트에 있는 모든 종목. 시세가 새로 나왔을 때 하루 1번 */
  async function writePrices(px) {
    if (!px?.items || !meta.lastPriceRow || meta.lastPriceRow < 2) return 0;
    const key = px.updated + '|' + meta.lastPriceRow;
    if (store.get('pxWritten2', '') === key) return 0;
    const last = meta.lastPriceRow;
    const cur = (await read(S.price, `A2:L${last}`)).values;
    let changed = 0;
    const C = [], L = [];
    cur.forEach(v => {
      const it = px.items[String(v[0]).trim()];
      const c = it && it.c ? it.c : v[2], mc = it && it.mc ? it.mc : v[11];
      if (c !== v[2] || mc !== v[11]) changed++;
      C.push([c]); L.push([mc]);
    });
    if (changed) await withSession(async ss => {
      await patch(S.price, `C2:C${last}`, { values: C }, ss);
      await patch(S.price, `L2:L${last}`, { values: L }, ss);
    });
    store.set('pxWritten2', key);
    return changed;
  }
  /* 처음 사는 종목이 ETF시세 시트에 없으면 KRX 시세 목록에서 찾아 맨 아래에 한 줄 추가 (코드·이름·종가·시가총액) */
  let pxRef = null;
  const usePrices = px => { pxRef = px; };
  const pxCodeOf = name => { if (!pxRef?.items) return null; for (const c in pxRef.items) if (pxRef.items[c].n === name) return c; return null; };
  async function ensureEtfRow(name, ss) {
    if (meta.codeOf?.[name]) return false;
    const code = pxCodeOf(name); if (!code) return false;
    const names = (await sget(S.price, 'B2:B2600', 'values', ss)).values;
    let last = 1; let exists = false;
    names.forEach((v, i) => { if (v[0]) { last = i + 2; if (String(v[0]).trim() === name) exists = true; } });
    if (exists) return false;
    const r = last + 1, it = pxRef.items[code];
    await patch(S.price, `A${r}`, { numberFormat: [['@']] }, ss); // 종목코드는 글자로 (앞자리 0 유지)
    await patch(S.price, `A${r}:C${r}`, { values: [[code, name, it.c || '']] }, ss);
    if (it.mc) await patch(S.price, `L${r}`, { values: [[it.mc]] }, ss);
    return true;
  }

  return { init, login, logout, user, load, cached, path, setPath, ensureFormulas, usePrices, setCash, addTrade, addDividend, addDeposit, addInterest, adjustCash, INTEREST, editTrade, deleteTrade, editDividend, deleteDividend, editDeposit, deleteDeposit, writePrices, colOf };
})();
