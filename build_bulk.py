# -*- coding: utf-8 -*-
"""루덴 남성복 12종 스마트스토어 일괄등록 엑셀 생성"""
import json, os, shutil
from openpyxl import load_workbook

W = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(W)
PRODUCTS = json.load(open(os.path.join(W, 'products.json'), encoding='utf-8'))

IMG_BASE = 'https://khcho0402-sys.github.io/luden-shop/'
CAT = {
    'coat-trench':      ('50021699', '트렌치코트', '남성 코트'),
    'cardigan-rabbit':  ('50021599', '카디건',     '남성 가디건'),
    'knit-halfzip':     ('50021579', '풀오버',     '남성 니트'),
    'jacket-corduroy':  ('50021640', '재킷',       '남성 재킷'),
    'pants-jogger':     ('50000836', '바지',       '남성 팬츠'),
    'hoodie-heavy':     ('50000830', '티셔츠',     '남성 후드티'),
    'fleece-zip':       ('50021620', '후드집업',   '남성 집업'),
    'knit-crew':        ('50021579', '풀오버',     '남성 니트'),
    'knit-turtleneck':  ('50021600', '터틀넥',     '남성 니트'),
    'knit-halfneck':    ('50021579', '풀오버',     '남성 니트'),
    'knit-vneck-muscle':('50021579', '풀오버',     '남성 니트'),
    'knit-polo':        ('50021579', '풀오버',     '남성 니트'),
}
PCODE = {
    'coat-trench': 'LD-TRENCH-COAT', 'cardigan-rabbit': 'LD-ZIP-CARDIGAN',
    'knit-halfzip': 'LD-HALFZIP-KNIT', 'jacket-corduroy': 'LD-CORDUROY-JACKET',
    'pants-jogger': 'LD-STRAIGHT-JOGGER', 'hoodie-heavy': 'LD-HEAVY-HOODIE',
    'fleece-zip': 'LD-FLEECE-ZIP', 'knit-crew': 'LD-CREW-KNIT',
    'knit-turtleneck': 'LD-TURTLENECK-KNIT', 'knit-halfneck': 'LD-HALFNECK-KNIT',
    'knit-vneck-muscle': 'LD-VNECK-MUSCLE', 'knit-polo': 'LD-POLO-KNIT',
}

INK, SUB, LINE, SOFT, SALE = '#111111', '#8A8A8E', '#EFEFEF', '#F5F5F5', '#D6402F'
FF = "-apple-system,BlinkMacSystemFont,'Apple SD Gothic Neo','Noto Sans KR','Malgun Gothic',sans-serif"


def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
             .replace('"', '&quot;'))


def size_table(p):
    st = p.get('sizeTable')
    if not st:
        return ('<p style="margin:10px 0 0;font-size:13px;line-height:1.8;color:#3A3A3E;">'
                + esc(p.get('sizeNote', '')) + '</p>')
    head = st['head']
    rows = st['rows']
    n = len(head)
    th = ''.join(
        '<th style="padding:10px 6px;font-size:11.5px;font-weight:600;color:#3A3A3E;'
        'background:%s;border-bottom:1px solid %s;white-space:nowrap;text-align:%s;">%s</th>'
        % (SOFT, LINE, 'left' if i == 0 else 'center', esc(h)) for i, h in enumerate(head))
    body = ''
    for r in rows:
        tds = ''
        for i, c in enumerate(r):
            tds += ('<td style="padding:10px 6px;font-size:12.5px;border-bottom:1px solid %s;'
                    'color:%s;text-align:%s;%s">%s</td>'
                    % (LINE, INK if i == 0 else '#3A3A3E', 'left' if i == 0 else 'center',
                       'font-weight:600;' if i == 0 else '', esc(str(c))))
        body += '<tr>' + tds + '</tr>'
    note = st.get('note', '')
    return ('<table style="width:100%%;border-collapse:collapse;margin-top:14px;table-layout:auto;">'
            '<thead><tr>%s</tr></thead><tbody>%s</tbody></table>'
            '<p style="margin:12px 0 0;font-size:11.5px;line-height:1.8;color:%s;">%s</p>'
            % (th, body, SUB, note))


def chip_row(colors):
    out = ''
    for c in colors:
        out += ('<span style="display:inline-block;margin:0 14px 10px 0;font-size:12px;color:#3A3A3E;">'
                '<span style="display:inline-block;width:13px;height:13px;border:1px solid #D8D8DC;'
                'vertical-align:-2px;margin-right:6px;background:%s;"></span>%s</span>'
                % (c['c'], esc(c['n'])))
    return out


def gosi_table(p, item, fabric_txt, care):
    rows = [
        ('품명', item),
        ('소재', fabric_txt),
        ('색상', ', '.join(c['n'] for c in p['colors'])),
        ('치수', ', '.join(p['sizes']) + ' (상세 실측표 참조)'),
        ('제조자', '루덴 협력사'),
        ('제조국', '중국'),
        ('제조연월', '2026. 10'),
        ('세탁방법', care + ' 제품 라벨의 취급표시를 함께 확인해 주세요.'),
        ('품질보증기준', '관련 법 및 소비자분쟁해결기준에 따름'),
        ('A/S 책임자', '조규희 · 010-4404-7637'),
    ]
    tr = ''
    for k, v in rows:
        tr += ('<tr><th style="width:88px;padding:8px 0;border-top:1px solid %s;font-size:11.5px;'
               'font-weight:500;color:%s;text-align:left;vertical-align:top;white-space:nowrap;">%s</th>'
               '<td style="padding:8px 0;border-top:1px solid %s;font-size:11.5px;line-height:1.7;color:#3A3A3E;">%s</td></tr>'
               % (LINE, SUB, k, LINE, esc(v)))
    return ('<table style="width:100%%;border-collapse:collapse;margin-top:12px;">%s</table>' % tr)


def detail_html(p):
    item = CAT[p['id']][2]
    rate = round((1 - p['price'] / p['listPrice']) * 100)
    imgs = [IMG_BASE + i for i in p['imgs']]
    points = p['points']
    pt_rows = ''
    for i, pt in enumerate(points):
        pt_rows += ('<tr><td style="width:30px;padding:14px 0;border-top:1px solid %s;vertical-align:top;'
                    'font-size:10.5px;letter-spacing:.08em;color:#B5B5B8;">%02d</td>'
                    '<td style="padding:14px 0;border-top:1px solid %s;font-size:13px;line-height:1.75;'
                    'color:#26262A;">%s</td></tr>' % (LINE, i + 1, LINE, pt))
    fab_rows = ''
    for k, v in p['fabric']:
        fab_rows += ('<tr><th style="width:76px;padding:10px 0;border-top:1px solid %s;font-size:11.5px;'
                     'font-weight:500;color:%s;text-align:left;vertical-align:top;white-space:nowrap;">%s</th>'
                     '<td style="padding:10px 0;border-top:1px solid %s;font-size:12.5px;line-height:1.7;'
                     'color:#26262A;">%s</td></tr>' % (LINE, SUB, esc(k), LINE, esc(v)))
    fabric_txt = ' / '.join('%s %s' % (k, v) for k, v in p['fabric'][:3])
    desc = '<br>'.join(esc(d) for d in p.get('desc', []))
    color_note = ('<p style="margin:12px 0 0;font-size:11.5px;color:%s;line-height:1.8;">'
                  '모니터 설정에 따라 실제 색상과 차이가 있을 수 있습니다.</p>' % SUB)

    html = f'''<div style="max-width:848px;margin:0 auto;padding:0 2px;font-family:{FF};font-size:13px;line-height:1.65;color:{INK};word-break:keep-all;">

<div style="border-top:2px solid {INK};padding:18px 0 0;">
<div style="font-size:10px;letter-spacing:.3em;color:{SUB};">LUDEN &nbsp;·&nbsp; MEN</div>
<h2 style="margin:10px 0 8px;font-size:23px;line-height:1.32;font-weight:700;letter-spacing:-.02em;">{esc(p['name'])}</h2>
<div style="font-size:12.5px;color:#5B5B60;">{esc(p['summary'])}</div>
<div style="margin-top:14px;font-size:12px;color:{SUB};">
<span style="color:{SALE};font-weight:700;font-size:16px;">{p['price']:,}원</span>
<span style="text-decoration:line-through;margin-left:8px;">{p['listPrice']:,}원</span>
<span style="margin-left:8px;color:{SALE};">{rate}%</span>
</div>
</div>

<div style="margin-top:22px;"><img src="{imgs[0]}" style="display:block;width:100%;border:0;" alt="{esc(p['name'])}"></div>

<div style="padding:34px 0 4px;">
<div style="font-size:10px;letter-spacing:.26em;color:{SUB};">CHECK POINT</div>
<table style="width:100%;border-collapse:collapse;margin-top:10px;">{pt_rows}</table>
</div>

<div style="margin-top:26px;"><img src="{imgs[1]}" style="display:block;width:100%;border:0;" alt=""></div>
<div style="margin-top:6px;"><img src="{imgs[2]}" style="display:block;width:100%;border:0;" alt=""></div>
<div style="margin-top:6px;"><img src="{imgs[3]}" style="display:block;width:100%;border:0;" alt=""></div>

<div style="padding:34px 0 4px;">
<div style="font-size:10px;letter-spacing:.26em;color:{SUB};">FABRIC &amp; DETAIL</div>
<table style="width:100%;border-collapse:collapse;margin-top:10px;">{fab_rows}</table>
<p style="margin:14px 0 0;font-size:12.5px;line-height:1.85;color:#3A3A3E;">{desc}</p>
</div>

<div style="padding:34px 0 4px;">
<div style="font-size:10px;letter-spacing:.26em;color:{SUB};">SIZE</div>
{size_table(p)}
</div>

<div style="padding:30px 0 4px;">
<div style="font-size:10px;letter-spacing:.26em;color:{SUB};">COLOR</div>
<div style="margin-top:14px;">{chip_row(p['colors'])}</div>
{color_note}
</div>

<div style="margin-top:30px;padding:20px 18px;background:{SOFT};">
<div style="font-size:10px;letter-spacing:.26em;color:{SUB};margin-bottom:12px;">DELIVERY &amp; EXCHANGE</div>
<table style="width:100%;border-collapse:collapse;font-size:12px;line-height:1.9;color:#3A3A3E;">
<tr><th style="width:76px;text-align:left;font-weight:600;color:{INK};vertical-align:top;padding:2px 0;">배송</th><td style="padding:2px 0;">결제 확인 후 영업일 기준 <b>2~3일 이내</b> 출고 · 택배 발송 (주말·공휴일 제외)</td></tr>
<tr><th style="text-align:left;font-weight:600;color:{INK};vertical-align:top;padding:2px 0;">배송비</th><td style="padding:2px 0;">기본 3,000원 · 70,000원 이상 구매 시 무료</td></tr>
<tr><th style="text-align:left;font-weight:600;color:{INK};vertical-align:top;padding:2px 0;">교환·반품</th><td style="padding:2px 0;">수령일부터 <b>7일 이내</b> 접수 · 단순 변심은 왕복 배송비 구매자 부담 · 하자·오배송은 전액 무료</td></tr>
<tr><th style="text-align:left;font-weight:600;color:{INK};vertical-align:top;padding:2px 0;">세탁</th><td style="padding:2px 0;">{esc(p['care'])}</td></tr>
</table>
</div>

<div style="padding:32px 0 4px;">
<div style="font-size:10px;letter-spacing:.26em;color:{SUB};">PRODUCT INFO</div>
{gosi_table(p, item, fabric_txt, p['care'])}
</div>

<div style="margin-top:30px;border-top:1px solid {LINE};padding:18px 0 4px;font-size:11px;line-height:1.9;color:{SUB};">
<b style="color:{INK};letter-spacing:.18em;">LUDEN</b> 루덴 &nbsp;|&nbsp; 대표 조규희 &nbsp;|&nbsp; 사업자등록번호 219-06-13931 &nbsp;|&nbsp; 통신판매업 신고 면제 (간이과세자)<br>
충청남도 아산시 남산로 41, 풀하우스 201호 &nbsp;|&nbsp; 010-4404-7637 &nbsp;|&nbsp; khcho0402@gmail.com
</div>

</div>'''
    return html


def main():
    tpl = os.path.join(ROOT, 'downloads', 'ExcelSaveTemplate_20260324.xlsx')
    out = os.path.join(ROOT, 'downloads', 'luden_bulk_12.xlsx')
    wb = load_workbook(tpl)
    ws = wb.worksheets[0]
    ws.delete_rows(3, 4)          # 작성가이드·예시 행 삭제 (3~6행)
    r = 3
    for p in PRODUCTS:
        cat, sub, item = CAT[p['id']]
        imgs = [IMG_BASE + i for i in p['imgs']]
        colors = [c['n'] for c in p['colors']]
        sizes = p['sizes']
        stock_each = 50
        total_stock = stock_each * len(colors) * len(sizes)
        row = {
            1: PCODE[p['id']],
            2: cat,
            3: p['name'],
            4: '신상품',
            5: p['listPrice'],
            6: 'N',
            10: '과세상품',
            12: total_stock,
            13: '조합형',
            14: '색상\n사이즈',
            15: ','.join(colors) + '\n' + ','.join(sizes),
            16: ','.join(['0'] * len(colors)),
            17: ','.join([str(stock_each)] * len(colors)),
            23: imgs[0],
            24: '\n'.join(imgs[1:]),
            25: detail_html(p),
            27: '루덴 협력사',
            30: '04',                     # 원산지 직접입력
            32: 'N',
            33: '중국',
            34: 'Y',
            36: '택배, 소포, 등기',
            37: 'HANJIN',
            38: '조건부 무료',
            39: 3000,
            40: '선결제',
            41: 70000,
            47: 4000,
            48: 8000,
            50: 'N',
            52: item,
            53: PCODE[p['id']],
            54: '해당사항 없음',
            55: '루덴 협력사',
            57: '010-4404-7637',
            58: '평일 10:00~18:00 네이버 톡톡 · 상품 Q&A 문의',
            59: '특이사항 없음',
            60: p['listPrice'] - p['price'],
            61: '원',
        }
        for col, val in row.items():
            ws.cell(row=r, column=col, value=val)
        print(p['id'], 'HTML', len(row[25]), 'chars ·', total_stock, 'stock')
        r += 1
    wb.save(out)
    print('saved', out)


if __name__ == '__main__':
    main()
