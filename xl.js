/* 원드라이브 엑셀 연동 (Microsoft Graph)
   - 로그인: 개인 Microsoft 계정 (MSAL, 리디렉션 방식: 아이폰 홈화면 앱에서도 동작)
   - 읽기: 퇴직연금 / 배당내역 / 입금내역 / ETF시세 시트
   - 쓰기: 매매·배당·입금 행 추가, 잔여 금액, 시세, 보유/매도 종목 목록 자동 정리
   개인 데이터는 이 파일이나 저장소에 저장되지 않음 (브라우저에서 Graph로 직접 주고받음) */
const XL = (() => {
  const CLIENT_ID = 'c37e15f3-c0d2-47d2-b031-ccbc894a7a8d';
  const SCOPES = ['Files.ReadWrite', 'User.Read'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
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
  const fromSerial = n => new Date(EPOCH + Math.round(n) * 864e5).toISOString().slice(0, 10);
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
      if (r.status === 429 || r.status === 503 || r.status === 504) { await new Promise(ok => setTimeout(ok, 1500 * (attempt + 1))); continue; }
      if (!r.ok) {
        let msg = r.status + ''; try { const j = await r.json(); msg = j.error?.message || msg; } catch {}
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
    finally { g('POST', `${await wb()}/closeSession`, null, s.id).catch(() => {}); }
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
    const stock = rr => ({ n: T(rr), fee: cell(rr, 21), target: isNum(cell(rr, 23)) ? cell(rr, 23) : 0, price: isNum(cell(rr, 28)) ? cell(rr, 28) : 0, row: rr });
    return {
      hold: hold.map(stock), sold: sold.map(stock), soldHeader,
      holdStart: 4, holdEnd: hold.length ? hold[hold.length - 1] : 3, soldEnd: sold.length ? sold[sold.length - 1] : soldHeader,
      depRow, depFormula: depRow ? forms[depRow - 1]?.[1] : '', cashRow, cash: cashRow ? cell(cashRow, 21) : 0
    };
  }
  async function load() {
    const [tr, blk, dv, dp, px] = await Promise.all([
      read(S.trade, 'B4:H600'),
      read(S.trade, 'T1:AN60', 'values,formulas'),
      read(S.div, 'B4:H1500'),
      read(S.dep, 'B4:D300'),
      read(S.price, 'A2:C2500')
    ]);
    const codeOf = {}, priceRow = {};
    px.values.forEach((v, i) => { if (v[1]) { codeOf[String(v[1]).trim()] = String(v[0]).trim(); priceRow[String(v[0]).trim()] = { row: i + 2, c: v[2] }; } });
    const tx = []; let lastTx = 3;
    tr.values.forEach((v, i) => {
      if (isNum(v[0]) && v[2]) { lastTx = i + 4; tx.push({ row: i + 4, d: fromSerial(v[0]), t: v[1] === '매도' ? 'sell' : 'buy', n: String(v[2]).trim(), q: +v[3] || 0, a: +v[5] || 0, pl: +v[6] || 0 }); }
    });
    const div = []; let lastDiv = 3, lastPensionDiv = 0;
    dv.values.forEach((v, i) => {
      if (isNum(v[0]) || v[2]) lastDiv = i + 4;
      if (isNum(v[0]) && v[1] === '퇴직연금') { lastPensionDiv = i + 4; div.push({ row: i + 4, d: fromSerial(v[0]), n: String(v[2]).trim(), q: +v[3] || 0, cost: +v[4] || 0, y: isNum(v[5]) ? v[5] : null, a: +v[6] || 0 }); }
    });
    const dep = []; let lastDep = 3;
    dp.values.forEach((v, i) => { if (isNum(v[0]) && isNum(v[2])) { lastDep = i + 4; dep.push({ row: i + 4, d: fromSerial(v[0]), a: v[2] }); } });
    const b = parseBlock(blk.values, blk.formulas);
    const addCode = s => ({ ...s, code: codeOf[s.n] || null });
    meta = { lastTx, lastDiv, lastPensionDiv, lastDep, block: b, priceRow };
    const data = {
      asof: new Date(Date.now() + 9 * 3600e3).toISOString().slice(0, 10), cash: +b.cash || 0,
      hold: b.hold.map(addCode), sold: b.sold.map(addCode), tx, div, dep,
      etfNames: Object.keys(codeOf), live: true, loadedAt: new Date().toISOString()
    };
    store.set('pensionCache', JSON.stringify(data));
    return data;
  }
  const cached = () => { try { return JSON.parse(store.get('pensionCache', 'null')); } catch { return null; } };

  /* ---------- 쓰기 ---------- */
  // 입금 금액(U18 등)을 입금내역 합계에 연결 (최초 1회)
  async function ensureDepositLink() {
    const b = meta.block; if (!b?.depRow) return false;
    if (String(b.depFormula).replace(/\s/g, '') === '=입금내역!D3') return false;
    await patch(S.trade, 'U' + b.depRow, { formulas: [['=입금내역!D3']] });
    return true;
  }
  async function setCash(v) {
    if (!meta.block?.cashRow) throw new Error('엑셀에서 "잔여 금액" 칸을 찾지 못했어요.');
    await patch(S.trade, 'U' + meta.block.cashRow, { values: [[v]] });
  }
  async function copyFormat(sheet, fromAddr) { return (await read(sheet, fromAddr, 'numberFormat')).numberFormat; }

  async function addTrade(x, D) {
    await withSession(async ss => {
      const r = meta.lastTx + 1, prev = Math.max(4, meta.lastTx);
      const nf = await copyFormat(S.trade, `B${prev}:H${prev}`);
      await patch(S.trade, `B${r}:H${r}`, { formulas: [tradeRow(x, r)], numberFormat: nf }, ss);
      await reconcileLists([...D.tx, x], ss);
    });
  }
  const tradeRow = (x, r) => [toSerial(x.d), x.t === 'sell' ? '매도' : '매수', x.n, x.q, `=G${r}/E${r}`, x.a, x.t === 'sell' ? x.pl : ''];

  // 보유종목 / 매도내역 목록을 매매내역에 맞게 정리 (한 번에 하나씩 고치고, 행 번호가 바뀌므로 매번 다시 읽음)
  //  - 보유 수량이 있는데 보유 목록에 없으면 → 추가
  //  - 보유 목록에 있는데 수량이 0이면 → 매도내역으로 이동 (매매 기록이 아예 없으면 목록에서 제거)
  async function reconcileLists(txAfter, ss) {
    const qty = {}, has = {};
    txAfter.forEach(t => { qty[t.n] = (qty[t.n] || 0) + (t.t === 'buy' ? t.q : -t.q); has[t.n] = true; });
    const R1 = async addr => (await g('GET', `${await rng(S.trade, addr)}?$select=formulasR1C1`, null, ss)).formulasR1C1;
    for (let guard = 0; guard < 8; guard++) {
      const blk = await g('GET', `${await rng(S.trade, 'T1:AN60')}?$select=values,formulas`, null, ss);
      const b = parseBlock(blk.values, blk.formulas);
      const inHold = n => b.hold.find(s => s.n === n), inSold = n => b.sold.find(s => s.n === n);
      const toAdd = Object.keys(qty).find(n => qty[n] > 0 && !inHold(n));
      const zero = b.hold.find(s => !(qty[s.n] > 0));
      const ghost = b.sold.find(s => !has[s.n]);
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
      await patch(S.trade, `B${o.row}:H${o.row}`, { formulas: [tradeRow(x, o.row)] }, ss);
      await reconcileLists(D.tx.map(t => t.row === o.row ? { ...t, ...x } : t), ss);
    });
  }
  async function deleteTrade(o, D) {
    await withSession(async ss => {
      await verify(S.trade, `B${o.row}:G${o.row}`, txExpect(o), ss);
      await shiftUp(S.trade, 'B', 'H', o.row, meta.lastTx, ss);
      await reconcileLists(D.tx.filter(t => t.row !== o.row), ss);
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
      await shiftUp(S.div, 'B', 'I', o.row, meta.lastDiv, ss);
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
      await shiftUp(S.dep, 'B', 'D', o.row, meta.lastDep, ss); // E·F열(연평균 수익률 계산용 수식)은 그대로 둠
    });
  }
  async function addDividend(x) {
    await withSession(async ss => {
      const r = meta.lastDiv + 1, tplRow = meta.lastPensionDiv || meta.lastDiv;
      const [nf, f] = await Promise.all([copyFormat(S.div, `B${tplRow}:H${tplRow}`), read(S.div, `E${tplRow}:G${tplRow}`, 'formulasR1C1')]);
      await patch(S.div, `B${r}:D${r}`, { values: [[toSerial(x.d), '퇴직연금', x.n]], numberFormat: [nf[0].slice(0, 3)] }, ss);
      await patch(S.div, `E${r}:G${r}`, { formulasR1C1: f.formulasR1C1, numberFormat: [nf[0].slice(3, 6)] }, ss);
      await patch(S.div, `H${r}`, { values: [[x.a]], numberFormat: [[nf[0][6]]] }, ss);
    });
  }
  async function addDeposit(x) {
    await withSession(async ss => {
      const r = meta.lastDep + 1, prev = Math.max(4, meta.lastDep);
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
  // KRX 시세를 ETF시세 시트 C열(종가)에 반영 (보유·매도 종목만)
  async function writePrices(px, D) {
    if (!px?.items || store.get('pxWritten', '') === px.updated) return 0;
    const codes = [...D.hold, ...D.sold].map(s => s.code).filter(Boolean);
    const todo = codes.map(c => ({ c, p: px.items[c]?.c, pr: meta.priceRow[c] })).filter(o => o.p && o.pr && o.pr.c !== o.p);
    if (todo.length) await withSession(async ss => { for (const o of todo) await patch(S.price, `C${o.pr.row}`, { values: [[o.p]] }, ss); });
    store.set('pxWritten', px.updated);
    return todo.length;
  }

  return { init, login, logout, user, load, cached, path, setPath, ensureDepositLink, setCash, addTrade, addDividend, addDeposit, editTrade, deleteTrade, editDividend, deleteDividend, editDeposit, deleteDeposit, writePrices, colOf };
})();
