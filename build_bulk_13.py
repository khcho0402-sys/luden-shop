# -*- coding: utf-8 -*-
"""루덴 13번째 상품(블랙 와이드 데님 팬츠 5452464) 스마트스토어 일괄등록 엑셀 생성"""
import os, sys
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_bulk as B

W = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(W)

P = {
    'id': 'pants-denim-wide',
    'code': '5452464',
    'name': '루덴 블랙 워싱 와이드 데님 팬츠',
    'summary': 'A라인 루즈 와이드 · 미드라이즈 · 무신축 데님',
    'listPrice': 41000,
    'price': 31000,
    'cats': ['ALL', 'PANTS', 'NEW'],
    'badge': 'NEW',
    'imgs': ['luden_prod_13_main.jpg', 'luden_prod_13_2.jpg',
             'luden_prod_13_3.jpg', 'luden_prod_13_4.jpg'],
    'colors': [{'n': '블랙', 'c': '#17171A'}],
    'sizes': ['S', 'M', 'L', 'XL', '2XL', '3XL'],
    'sizeTable': {
        'unit': '실측 (단위 cm · 허리·엉덩이 둘레 / 밑단 단면)',
        'head': ['사이즈', '총장', '허리', '엉덩이', '밑단'],
        'rows': [
            ['S', '100', '70', '93', '25'],
            ['M', '102', '74', '97', '26'],
            ['L', '105', '80', '103', '28'],
            ['XL', '107', '84', '107', '28'],
            ['2XL', '109', '88', '111', '30'],
            ['3XL', '111', '90', '119', '32'],
        ],
        'note': '1~3cm 오차가 있을 수 있습니다. 허리·엉덩이는 <b>둘레</b>, 밑단은 <b>단면</b> 기준입니다. '
                '무신축 데님이라 평소보다 한 치수 여유 있게 선택하시는 것을 권합니다.<br>'
                '업체 권장(키 / 몸무게): S 155–160cm · 44–49kg / M 160–168 · 47.5–55 / L 165–173 · 55–65 / '
                'XL 170–178 · 65–72.5 / 2XL 175–183 · 72.5–80 / 3XL 180–185 · 80–87.5',
    },
    'fabric': [
        ['겉감', '면 60% 데님 (나머지 혼방)'],
        ['가공', '블랙 워시 · 무신축'],
        ['두께', '일반형 (봄·가을)'],
        ['허리', '미드라이즈 · 지퍼 여밈'],
        ['핏', '스트레이트 · A라인 루즈 와이드'],
        ['기장', '긴바지 · 스트레이트 밑단'],
    ],
    'points': [
        '허리에서 밑단까지 <b>일자로 떨어지는 A라인 와이드</b> 실루엣이라 다리 라인이 그대로 드러나지 않습니다.',
        '<b>미드라이즈</b>에 지퍼 여밈이라 니트를 넣어 입어도 허리선이 잡히고, 벨트 없이도 깔끔합니다.',
        '면 60% 데님에 <b>워시 가공</b>을 넣어 검정이 죽지 않고 은은하게 결이 살아 있습니다.',
        '<b>무신축</b> 원단이라 앉았다 일어나도 무릎이 늘어나지 않고, S~3XL까지 6사이즈로 나옵니다.',
    ],
    'desc': [
        '블랙 한 컬러입니다. 완전한 검정이 아니라 워싱이 살짝 들어가 데님 결이 보이는 톤입니다.',
        '와이드 실루엣이라 니트·후드·재킷 어디에 받쳐도 붙고, 남녀 공용으로 입기 좋은 사이즈입니다.',
        '밑단은 스트레이트로 떨어지고 솔기 스티치가 또렷해 신발 위에 자연스럽게 얹힙니다.',
    ],
    'care': '데님은 찬물 단독 세탁 후 뒤집어서 그늘에 건조해 주세요. 첫 세탁에서 물빠짐이 있을 수 있고, '
            '건조기 사용은 수축·색 변형의 원인이 됩니다.',
    'link': '',
}

B.CAT[P['id']] = ('50000836', '바지', '남성 팬츠')
B.PCODE[P['id']] = 'LD-WIDE-DENIM'


def main():
    tpl = os.path.join(ROOT, 'downloads', 'ExcelSaveTemplate_20260324.xlsx')
    out = os.path.join(ROOT, 'downloads', 'luden_bulk_13.xlsx')
    wb = load_workbook(tpl)
    ws = wb.worksheets[0]
    ws.delete_rows(3, 4)
    r = 3
    p = P
    cat, sub, item = B.CAT[p['id']]
    imgs = [B.IMG_BASE + i for i in p['imgs']]
    colors = [c['n'] for c in p['colors']]
    sizes = p['sizes']
    stock_each = 50
    total_stock = stock_each * len(colors) * len(sizes)
    row = {
        1: B.PCODE[p['id']],
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
        25: B.detail_html(p),
        27: '루덴 협력사',
        30: '04',
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
        53: B.PCODE[p['id']],
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
    wb.save(out)
    print('HTML', len(row[25]), 'chars ·', total_stock, 'stock · saved', out)


if __name__ == '__main__':
    main()
