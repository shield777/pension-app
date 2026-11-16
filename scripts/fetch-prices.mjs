// KRX ETF 전종목 시세를 받아 prices.json 으로 저장 (Node 20, 외부 패키지 없음)
// 1순위: KRX 정보데이터시스템 / 2순위(실패 시): 네이버 금융 ETF 전종목
// 전종목을 통째로 저장하므로, 어떤 종목을 보유했는지는 드러나지 않음
import { writeFileSync } from 'node:fs';

const kst = () => new Date(Date.now() + 9 * 3600e3);
const ymd = d => d.toISOString().slice(0, 10).replace(/-/g, '');
const num = v => Number(String(v ?? '').replace(/,/g, '')) || 0;
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36';

async function fromKRX() {
  for (let back = 0; back < 10; back++) {
    const d = kst(); d.setUTCDate(d.getUTCDate() - back);
    const res = await fetch('https://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Referer': 'https://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC0201030101',
        'User-Agent': UA, 'X-Requested-With': 'XMLHttpRequest'
      },
      body: new URLSearchParams({ bld: 'dbms/MDC/STAT/standard/MDCSTAT04301', locale: 'ko_KR', trdDd: ymd(d), share: '1', money: '1', csvxls_isNo: 'false' })
    });
    if (!res.ok) throw new Error('KRX HTTP ' + res.status);
    const text = await res.text();
    let j; try { j = JSON.parse(text); } catch { throw new Error('KRX 응답이 JSON이 아님 (로그인 필요 가능성): ' + text.slice(0, 120)); }
    const rows = j.output || j.OutBlock_1 || [];
    if (rows.length && num(rows[0].TDD_CLSPRC)) {
      const items = {};
      for (const r of rows) items[r.ISU_SRT_CD] = { n: r.ISU_ABBRV, c: num(r.TDD_CLSPRC), mc: num(r.MKTCAP) };
      return { source: 'KRX', asof: ymd(d).replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3'), items };
    }
  }
  throw new Error('KRX: 최근 10일 데이터 없음');
}

async function fromNaver() {
  const res = await fetch('https://finance.naver.com/api/sise/etfItemList.nhn?etfType=0&targetColumn=market_sum&sortOrder=desc', { headers: { 'User-Agent': UA, 'Referer': 'https://finance.naver.com/sise/etfList.naver' } });
  if (!res.ok) throw new Error('Naver HTTP ' + res.status);
  const buf = await res.arrayBuffer();
  // 네이버는 EUC-KR로 응답하는 경우가 있어, 깨진 글자(�)가 보이면 EUC-KR로 다시 읽음
  let text = new TextDecoder('utf-8').decode(buf);
  if (text.includes('\uFFFD')) text = new TextDecoder('euc-kr').decode(buf);
  const list = JSON.parse(text).result.etfItemList;
  const items = {};
  for (const r of list) items[r.itemcode] = { n: r.itemname, c: num(r.nowVal), mc: num(r.marketSum) * 1e8 };
  return { source: 'NAVER', asof: kst().toISOString().slice(0, 10), items };
}

let out, errors = [];
for (const f of [fromKRX, fromNaver]) {
  try { out = await f(); break; } catch (e) { errors.push(e.message); console.log('실패:', e.message); }
}
if (!out) { console.error(errors.join('\n')); process.exit(1); }
out.updated = new Date().toISOString();
if (errors.length) out.note = errors.join(' / ');
writeFileSync('prices.json', JSON.stringify(out));
console.log(`${out.source} ${out.asof} ${Object.keys(out.items).length}종목 저장`);
