"""BOSU 출고분석 페이지 빌드.

사용법:
  python3 build_bosu.py <ERP 출고현황.xlsx> <입고원가.json> <비밀번호> [출력 경로]

입고원가.json 예: {"Bosu Professional": 161058, "Bosu PODS (2PACK)": 29215, "Bosu Dome": 83944}
원가·출고 데이터는 저장소에 올리지 말 것. 결과 HTML 안에는 비밀번호로 암호화되어 들어간다.
"""
import base64, json, os, sys
from datetime import datetime, timezone

import pandas as pd
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

HERE = os.path.dirname(os.path.abspath(__file__))
ITER = 600000
A = '출고금액(공급가액)'


def load(path):
    raw = pd.read_excel(path, header=0)
    total_row = raw[pd.to_numeric(raw['No'], errors='coerce').isna()]
    d = raw[pd.to_numeric(raw['No'], errors='coerce').notna()].copy()
    d['출고일'] = pd.to_datetime(d['출고일'])
    d['수주일'] = pd.to_datetime(d['수주일'], errors='coerce')
    for c in ['출고수량', A, '부가세액', '합계액', '단가(부가세포함)']:
        d[c] = pd.to_numeric(d[c])
    d['판매처명'] = d['판매처명'].astype(str).str.strip()
    return d, total_row


def ds(t):
    return t.strftime('%Y-%m-%d')


def issues_for(d, total_row, cost):
    out = []
    paid = d[d[A] > 0]

    # 1. 합계 검증
    if len(total_row):
        tq = pd.to_numeric(str(total_row.iloc[0]['출고수량']).replace(',', ''), errors='coerce')
        ta = pd.to_numeric(str(total_row.iloc[0][A]).replace(',', ''), errors='coerce')
        ok = tq == d['출고수량'].sum() and ta == d[A].sum()
        out.append(dict(level='ok' if ok else 'chk', title='엑셀 합계행과 내역 합계 일치' if ok else '엑셀 합계행과 내역 합계 불일치',
                        detail=f"엑셀 맨 아래 합계행(수량 {int(tq):,}개, 공급가 {int(ta):,}원)을 빼고 {len(d)}행을 분석했습니다. 내역을 더한 값은 수량 {int(d['출고수량'].sum()):,}개, 공급가 {int(d[A].sum()):,}원입니다.",
                        cols=[], rows=[]))

    # 2. 금액 정합
    bad = d[((d['부가세액'] - (d[A] * 0.1).round()).abs() > 2) | ((d['합계액'] - d[A] - d['부가세액']).abs() > 1)
            | ((d['단가(부가세포함)'] * d['출고수량'] - d['합계액']).abs() > d['출고수량'] + 1)]
    out.append(dict(level='ok' if bad.empty else 'chk', title=f"단가·공급가·부가세·합계 계산 {'이상 없음' if bad.empty else f'불일치 {len(bad)}건'}",
                    detail='모든 행에서 공급가 × 10% = 부가세, 공급가 + 부가세 = 합계, 단가 × 수량 = 합계가 맞는지 확인했습니다.',
                    cols=['출고일', '판매처', '품목', '수량', '공급가', '합계'],
                    rows=[[ds(r.출고일), r.판매처명, r.품목명, int(r.출고수량), int(r[A]), int(r.합계액)] for _, r in bad.iterrows()]))

    # 3. 반품 미포함
    ret_cols = ['출고취소번호', '미매출이월출고번호', '반품수주번호', '반품지시번호', '반품출고번호']
    n_ret = int(d[ret_cols].notna().sum().sum())
    out.append(dict(level='chk', title='반품·출고취소가 들어 있지 않음 (출고 기준이지 순매출이 아님)',
                    detail=f"취소·반품 관련 열 {len(ret_cols)}개가 {'모두 비어 있습니다' if n_ret == 0 else f'{n_ret}칸만 채워져 있습니다'}. 이 파일은 판매출고만 담고 있어서 불량 교환, 크레딧 처리, 반품으로 돌아온 금액은 빠져 있습니다. 실제 순매출은 반품 내역을 따로 받아 차감해야 합니다.",
                    cols=[], rows=[]))

    # 4. 0원 출고
    z = d[d[A] == 0]
    zs = z[z['수주유형'].str.startswith('정상')]
    by_type = z.groupby('수주유형')['출고수량'].agg(['size', 'sum'])
    desc = ', '.join(f"{k} {int(v['size'])}건 {int(v['sum'])}개" for k, v in by_type.iterrows())
    out.append(dict(level='fix', title=f'0원 출고 {len(z)}건 ({int(z["출고수량"].sum())}개)은 매출 분석에서 제외',
                    detail=f"{desc}. 금액이 없어서 매출·단가·마진 계산에서는 뺐고, 거래처 상세 내역에서는 보입니다.",
                    cols=['출고일', '판매처', '품목', '수량', '수주유형', '담당', '비고'],
                    rows=[[ds(r.출고일), r.판매처명, r.품목명, int(r.출고수량), r.수주유형, str(r.영업담당자명 or ''), str(r.납품지시비고) if pd.notna(r.납품지시비고) else ''] for _, r in z.iterrows()]))
    if len(zs):
        out.append(dict(level='chk', title=f'정상매출인데 금액이 0원인 출고 {len(zs)}건',
                        detail='수주유형은 정상매출인데 공급가가 0원입니다. 덤(같은 출고번호에 유상 품목과 함께 나간 1~2개)인지, 금액 입력이 빠진 건지 확인이 필요합니다.',
                        cols=['출고일', '출고번호', '판매처', '품목', '수량', '담당'],
                        rows=[[ds(r.출고일), r.출고번호, r.판매처명, r.품목명, int(r.출고수량), str(r.영업담당자명 or '')] for _, r in zs.iterrows()]))

    # 5. 같은 출고번호·품목 두 줄
    dup = d[d.duplicated(['출고번호', '품목'], keep=False)].sort_values(['출고번호'])
    same = d[d.duplicated(['출고번호', '품목', '출고수량', A], keep=False)]
    same = same[same[A] > 0]
    if len(dup):
        out.append(dict(level='chk' if len(same) else 'info',
                        title=f'같은 출고번호에 같은 품목이 두 줄 이상: {dup["출고번호"].nunique()}건' + (f' (수량·금액까지 똑같은 줄 {same["출고번호"].nunique()}건)' if len(same) else ''),
                        detail='대부분 유상 출고와 0원 덤을 나눠 적었거나 단가가 다른 분할 출고입니다. 다만 수량·금액이 완전히 같은 줄은 중복 입력일 수 있으니 ERP에서 확인해 주세요: '
                               + ', '.join(sorted({f"{r.판매처명} {ds(r.출고일)}" for _, r in same.iterrows()})) + '.' if len(same) else '대부분 유상 출고와 0원 덤을 나눠 적었거나 단가가 다른 분할 출고입니다.',
                        cols=['출고일', '출고번호', '판매처', '품목', '수량', '공급가'],
                        rows=[[ds(r.출고일), r.출고번호, r.판매처명, r.품목명, int(r.출고수량), int(r[A])] for _, r in dup.iterrows()]))

    # 6. 영업그룹 누락
    ng = d[d['영업그룹'].isna()]
    if len(ng):
        out.append(dict(level='fix', title=f'영업그룹이 비어 있는 출고 {len(ng)}건 보정',
                        detail='같은 담당자의 다른 출고에 적힌 팀으로 채웠습니다: ' + ', '.join(sorted({f"{r.판매처명}({r.영업담당자명})" for _, r in ng.iterrows()})) + '.',
                        cols=['출고일', '판매처', '품목', '공급가', '담당'],
                        rows=[[ds(r.출고일), r.판매처명, r.품목명, int(r[A]), str(r.영업담당자명)] for _, r in ng.iterrows()]))

    # 7. 팀 표기 변경
    x = d.dropna(subset=['영업담당자명', '영업그룹']).groupby(['영업담당자명', '영업그룹'])['출고일'].agg(['min', 'max', 'size']).reset_index()
    multi = x[x['영업담당자명'].isin(x['영업담당자명'].value_counts()[lambda s: s > 1].index)].sort_values(['영업담당자명', 'min'])
    y = d.dropna(subset=['영업그룹']).groupby('판매처명')['영업그룹'].nunique()
    moved = d[d['판매처명'].isin(y[y > 1].index)].groupby(['판매처명', '영업그룹', '영업담당자명'])['출고일'].agg(['min', 'max', 'size']).reset_index().sort_values(['판매처명', 'min'])
    out.append(dict(level='info', title=f'시기에 따라 팀 표기가 바뀐 담당자 {multi["영업담당자명"].nunique()}명, 담당 팀이 바뀐 거래처 {moved["판매처명"].nunique()}곳',
                    detail='ERP는 출고 당시 영업그룹을 기록합니다. 26년 초 팀 표기 변경과 거래처 담당 이관(' + ', '.join(sorted(set(moved['판매처명'])))[:200] + ' 등) 때문에 팀별 전년비는 실제 영업 성과보다 크게 오르내려 보입니다. 거래처별·담당자별 비교를 같이 보세요.',
                    cols=['구분', '이름', '팀', '담당', '첫 출고', '마지막 출고', '건수'],
                    rows=[['담당자', r.영업담당자명, r.영업그룹, '', ds(r['min']), ds(r['max']), str(int(r['size']))] for _, r in multi.iterrows()]
                         + [['거래처', r.판매처명, r.영업그룹, r.영업담당자명, ds(r['min']), ds(r['max']), str(int(r['size']))] for _, r in moved.iterrows()]))

    # 8. 단가 특이값 (온라인 소비자가 제외 후)
    rows = []
    for p, g in paid.groupby('품목명'):
        g = g.assign(u=g[A] / g['출고수량'])
        b2b = g[g['세무구분'] != '건별매출']
        q1, q3 = b2b.u.quantile([.25, .75])
        lo, hi = q1 - 1.5 * (q3 - q1), q3 + 1.5 * (q3 - q1)
        for _, r in b2b[(b2b.u < lo) | (b2b.u > hi)].iterrows():
            rows.append([ds(r.출고일), r.판매처명, p, int(r.출고수량), int(round(r.u)), str(r.영업담당자명 or '')])
    out.append(dict(level='info', title=f'B2B 단가가 유난히 높거나 낮은 출고 {len(rows)}건',
                    detail='온라인몰 건별매출(소비자가로 잡힘)을 뺀 B2B 출고에서 같은 품목의 일반적인 단가 범위를 크게 벗어난 건입니다. 클레임 보상, 특가, 개인 판매 등 사유가 있는지 참고하세요. 온라인몰 단가가 27~32만원으로 높은 건 소비자가라서 정상입니다.',
                    cols=['출고일', '판매처', '품목', '수량', '개당 공급가', '담당'], rows=rows))

    # 9. 거래처명
    odd = [n for n in sorted(d['판매처명'].unique()) if n.startswith('$') or 'XXX' in n or '클레임' in n or '가계정' in n]
    out.append(dict(level='info', title='매출처로 보기 애매한 거래처명',
                    detail='개인·내부 계정($), 말소 표기(XXX), 제품 클레임 보상 계정, 온라인몰 가계정이 판매처로 잡혀 있습니다. 클레임 보상 계정의 유상 출고는 매출에 포함되어 있습니다.',
                    cols=['판매처', '유상 건수', '공급가'],
                    rows=[[n, str(int((paid['판매처명'] == n).sum())), int(paid.loc[paid['판매처명'] == n, A].sum())] for n in odd]))

    # 10. 날짜
    early = d[d['출고일'] < d['수주일']]
    out.append(dict(level='info', title=f'출고일이 수주일보다 빠른 건 {len(early)}건',
                    detail='대부분 하루 차이로 수주 등록이 늦은 경우입니다. 분석에는 출고일을 씁니다.',
                    cols=['출고일', '수주일', '판매처', '품목', '공급가'],
                    rows=[[ds(r.출고일), ds(r.수주일), r.판매처명, r.품목명, int(r[A])] for _, r in early.iterrows()]))

    mx = d['출고일'].max()
    out.append(dict(level='info', title=f'{mx.year}년 {mx.month}월은 {mx.day}일까지만 들어 있음',
                    detail=f'그래서 기본 비교는 두 해 모두 1월 1일~{mx.month}월 {mx.day}일 같은 기간으로 맞췄습니다. 25년 연간과 비교하려면 위의 "연간" 버튼을 누르세요.', cols=[], rows=[]))
    out.append(dict(level='info', title='마진은 현재 입고원가로 계산한 추정치',
                    detail='SKU별 현재 입고원가를 과거 출고에도 똑같이 적용했습니다. 25년 당시 실제 원가(환율 등)와 다를 수 있고, 물류비·판매장려금 같은 비용은 빠져 있습니다.', cols=[], rows=[]))
    order = {'chk': 0, 'fix': 1, 'info': 2, 'ok': 3}
    return sorted(out, key=lambda i: order[i['level']])


def main():
    xlsx, cost_path, pw = sys.argv[1:4]
    out_path = sys.argv[4] if len(sys.argv) > 4 else os.path.join(HERE, '..', 'index.html')
    cost = json.load(open(cost_path, encoding='utf-8'))
    d, total_row = load(xlsx)
    missing = set(d['품목명']) - set(cost)
    if missing:
        sys.exit(f'입고원가가 없는 품목: {missing}')
    issues = issues_for(d, total_row, cost)
    # 영업그룹 보정: 같은 담당자의 가장 흔한 팀
    rep_team = d.dropna(subset=['영업그룹']).groupby('영업담당자명')['영업그룹'].agg(lambda s: s.mode().iloc[0])
    team = d['영업그룹'].fillna(d['영업담당자명'].map(rep_team)).fillna('미지정')
    rows = [[ds(r.출고일), r.품목명, int(r.출고수량), int(r[A]), r.판매처명,
             '' if pd.isna(r.거래처대분류) else r.거래처대분류, '' if pd.isna(r.거래처중분류) else r.거래처중분류,
             t, '' if pd.isna(r.영업담당자명) else r.영업담당자명, r.수주유형]
            for (_, r), t in zip(d.iterrows(), team)]
    payload = dict(rows=rows, cost=cost, issues=issues,
                   meta=dict(minDate=ds(d['출고일'].min()), maxDate=ds(d['출고일'].max()), nRows=len(d),
                             source=os.path.basename(xlsx), built=datetime.now(timezone.utc).isoformat()))
    pt = json.dumps(payload, ensure_ascii=False, separators=(',', ':')).encode()
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode())
    ct = AESGCM(key).encrypt(iv, pt, None)
    E = dict(s=base64.b64encode(salt).decode(), i=base64.b64encode(iv).decode(), c=base64.b64encode(ct).decode(), n=ITER)
    tpl = open(os.path.join(HERE, 'template.html'), encoding='utf-8').read()
    html = tpl.replace('/*__PAYLOAD__*/null', json.dumps(E))
    open(out_path, 'w', encoding='utf-8').write(html)
    print(f'{out_path}: {len(rows)} rows, {len(issues)} issues, {len(html):,} bytes')


if __name__ == '__main__':
    main()
