"""본문 수치 검증 스크립트

결과물(대시보드·제안서)에 쓴 주요 수치를 원자료에서 build_all.py와 다른 경로(pandas)로 다시 계산해 대조한다.
사용법: python verify.py <공공데이터 폴더> <출품작_작업 폴더>
"""
import json, os, sys
import pandas as pd

D, OUT = sys.argv[1], sys.argv[2]
P = lambda n: os.path.join(D, '전국폐기물 발생 및 처리현황 _ ' + n)
NUM = ['gen', 'rec', 'mat', 'en', 'inc', 'land', 'etc']


def load(path, nkey):
    df = pd.read_excel(path, header=None, skiprows=5, dtype=str)
    df = df.iloc[:, :nkey + 3 + 7]
    df.columns = [f'k{i}' for i in range(nkey)] + ['a', 'b', 'c'] + NUM
    df[NUM] = df[NUM].astype(float)
    return df


sido = load(P('2-가-1). (시도) 생활(가정)폐기물 발생량 (1).xlsx'), 1)
gu = load(P('2-나-1). (시군구) 생활(가정)폐기물 발생량 (1).xlsx'), 2)
pop = pd.read_excel(P('1-가. (시도) 생활폐기물관리구역현황.xlsx'), header=None, skiprows=4, dtype=str).set_index(0)[2].astype(float)
gpop = pd.read_excel(P('1-나. (시군구) 생활폐기물관리구역현황.xlsx'), header=None, skiprows=4, dtype=str).set_index(1)[3].astype(float)
csv = pd.read_csv(os.path.join(D, '대구광역시_생활폐기물 처리_20241231.csv'), encoding='cp949').set_index('종류')


def row(df, keys, a, b=None, c=None):
    q = df
    for i, k in enumerate(keys): q = q[q[f'k{i}'] == k]
    q = q[q.a == a]
    q = q[q.b.isna()] if b is None else q[q.b == b]
    q = q[q.c.isna()] if c is None else q[q.c == c]
    assert len(q) == 1, (keys, a, b, c, len(q))
    return q.iloc[0]


MIX, SEP, FOOD = '종량제방식 등 혼합배출', '재활용 가능자원 분리배출', '음식물류 폐기물 분리배출'
names = [n for n in sido.k0.unique()]
T = {n: row(sido, [n], '총계') for n in names}
M = {n: row(sido, [n], MIX, '소계') for n in names}
S = {n: row(sido, [n], SEP, '소계') for n in names}
RS = {n: row(sido, [n], SEP, '재활용 잔재물') for n in names}
CW = {n: row(sido, [n], MIX, '건설폐재류', '건설폐재류(공사장생활폐기물)') for n in names}
r17 = [n for n in names if n != '전국']
land = {n: T[n].land / T[n].gen * 100 for n in names}
mixland = {n: M[n].land / M[n].gen * 100 for n in names}
resid = {n: RS[n].gen / S[n].gen * 100 for n in names}
mat = {n: T[n].mat / T[n].gen * 100 for n in names}
energy = {n: T[n].en / T[n].gen * 100 for n in names}
kg = {n: T[n].gen * 1000 / pop[n] / 366 for n in names}
cshare = {n: CW[n].land / T[n].land * 100 if T[n].land else float('nan') for n in names}
rank = lambda d, n='대구', asc=False: 1 + sum(1 for x in r17 if (d[x] < d[n] if asc else d[x] > d[n]))

dg = ('대구', '소계')
it = lambda b, c: row(gu, dg, MIX, b, c)
paper, plastic, metal, glass = it('가연성', '폐지류'), it('가연성', '폐합성수지류'), it('불연성', '폐금속류'), it('불연성', '폐유리류')
foodmix, con = it('가연성', '음식물류 폐기물'), it('건설폐재류', '건설폐재류(공사장생활폐기물)')
Tg, Mg, Sg, Fg, RSg = row(gu, dg, '총계'), row(gu, dg, MIX, '소계'), row(gu, dg, SEP, '소계'), row(gu, dg, FOOD), row(gu, dg, SEP, '재활용 잔재물')
rGen = paper.gen + plastic.gen + metal.gen + glass.gen
rLand = paper.land + plastic.land + metal.land + glass.land


def G(g):
    t, m, s = row(gu, ('대구', g), '총계'), row(gu, ('대구', g), MIX, '소계'), row(gu, ('대구', g), SEP, '소계')
    rs = row(gu, ('대구', g), SEP, '재활용 잔재물')
    return dict(kg=t.gen * 1000 / gpop[g] / 366, land=t.land / t.gen * 100, mixland=m.land / m.gen * 100, mat=t.mat / t.gen * 100,
                resid=rs.gen / s.gen * 100 if s.gen else 0)


EF = lambda doc: doc * .5 * 1.0 * .5 * 16 / 12 * .9 * 28


def scen(a, b, c, r):
    cut = a * rLand + b * con.land + c * foodmix.land
    landT = Tg.land - cut
    matNow = Sg.mat + Fg.mat + Mg.mat
    gain = a * rGen * (1 - r / 100) + b * con.gen * .8 + max(0, (RSg.gen / Sg.gen * 100 - r) / 100) * Sg.gen
    return dict(cut=cut, landT=landT, landPct=landT / Tg.gen * 100, matPct=(matNow + gain) / Tg.gen * 100,
                ghg=a * paper.land * EF(.40) + c * foodmix.land * EF(.15), left=Mg.land - cut, perDay=(Mg.land - cut) / 366)


goal, cons, aggr = scen(.3, .5, .5, resid['전국'].round(1)), scen(.2, .3, .3, 15), scen(.5, .8, .7, 8)
data = json.load(open(os.path.join(OUT, 'data.json'), encoding='utf-8'))
rb = data['cluster']['robust']
sil = [x['silhouette'] for x in data['cluster']['scan']]

# (설명, 본문에 쓴 값, 다시 계산한 값, 허용 오차)
C = [
    ('대구 매립률(%)', 26.8, land['대구'], .05), ('전국 매립률(%)', 9.9, land['전국'], .05), ('매립률 대구/전국(배)', 2.7, land['대구'] / land['전국'], .05),
    ('매립률 순위', 1, rank(land), 0),
    ('대구 종량제 매립률(%)', 49.8, mixland['대구'], .05), ('전국 종량제 매립률(%)', 19.2, mixland['전국'], .05), ('종량제 매립률 순위', 1, rank(mixland), 0),
    ('대구 잔재물률(%)', 20.9, resid['대구'], .05), ('전국 잔재물률(%)', 11.3, resid['전국'], .05), ('잔재물률 순위', 1, rank(resid), 0),
    ('대구 물질재활용률(%)', 24.5, mat['대구'], .05), ('전국 물질재활용률(%)', 46.8, mat['전국'], .05), ('물질재활용률 순위', 16, rank(mat), 0),
    ('대구 에너지회수율(%)', 36.7, energy['대구'], .05), ('전국 에너지회수율(%)', 13.0, energy['전국'], .05), ('에너지회수율 순위', 1, rank(energy), 0),
    ('대구 1인당 발생량(kg/일)', 0.86, kg['대구'], .005), ('전국 1인당 발생량(kg/일)', 0.88, kg['전국'], .005), ('1인당 발생량 적은 순위', 7, rank(kg, asc=True), 0),
    ('대구 발생량(만 t)', 75.6, Tg.gen / 1e4, .05), ('종량제 배출량(만 t)', 40.2, Mg.gen / 1e4, .05), ('종량제 매립량(만 t)', 20.0, Mg.land / 1e4, .05),
    ('대구 매립량(만 t)', 20.2, Tg.land / 1e4, .05),
    ('봉투 속 재활용 가능 품목(만 t)', 17.4, rGen / 1e4, .05), ('봉투 속 재활용 가능 비중(%)', 43, rGen / Mg.gen * 100, .5),
    ('봉투 속 재활용품 매립량(만 t)', 7.4, rLand / 1e4, .05), ('매립량 중 봉투 속 재활용품(%)', 37, rLand / Tg.land * 100, .5),
    ('공사장생활폐기물 매립량(만 t)', 5.1, con.land / 1e4, .05), ('매립량 중 공사장생활폐기물(%)', 25.3, cshare['대구'], .05),
    ('전국 공사장생활폐기물 매립 비중(%)', 7.9, cshare['전국'], .05), ('공사장 매립 비중 순위', 2, rank(cshare), 0),
    ('공사장생활폐기물 매립 비율(%)', 99.6, con.land / con.gen * 100, .05), ('공사장생활폐기물 재활용률(%)', 0.2, con.rec / con.gen * 100, .05),
    ('잔재물 발생량(만 t)', 3.2, RSg.gen / 1e4, .05), ('잔재물 중 에너지회수(%)', 75, RSg.en / RSg.gen * 100, .5),
    ('봉투 속 음식물 매립량(만 t)', 1.1, foodmix.land / 1e4, .05), ('분리배출 음식물 재활용률(%)', 100, Fg.rec / Fg.gen * 100, .5),
    ('중구 1인당 발생량', 1.08, G('중구')['kg'], .005), ('중구 물질재활용률', 11.0, G('중구')['mat'], .05),
    ('서구 잔재물률', 28.3, G('서구')['resid'], .05), ('동구 잔재물률', 26.7, G('동구')['resid'], .05), ('수성구 잔재물률', 26.3, G('수성구')['resid'], .05),
    ('북구 종량제 매립률', 62.0, G('북구')['mixland'], .05), ('달서구 종량제 매립률', 59.5, G('달서구')['mixland'], .05),
    ('달성군 잔재물률', 10.3, G('달성군')['resid'], .05), ('달성군 물질재활용률', 31.7, G('달성군')['mat'], .05), ('군위군 매립률', 36.9, G('군위군')['land'], .05),
    ('목표 시나리오 매립 감축(만 t)', 5.3, goal['cut'] / 1e4, .05), ('목표 매립 감축률(%)', 26, goal['cut'] / Tg.land * 100, .5),
    ('목표 매립률(%)', 19.7, goal['landPct'], .05), ('목표 물질재활용률(%)', 35.3, goal['matPct'], .05),
    ('목표 메탄 잠재 감축(tCO2eq)', 32744, goal['ghg'], 1), ('목표 후 남는 종량제 매립(만 t)', 14.7, goal['left'] / 1e4, .05), ('목표 후 하루 매립(t)', 402, goal['perDay'], .5),
    ('현재 종량제 하루 매립(t)', 547, Mg.land / 366, .5),
    ('보수적 매립량(만 t)', 16.9, cons['landT'] / 1e4, .05), ('보수적 매립률', 22.3, cons['landPct'], .05), ('보수적 물질재활용률', 31.3, cons['matPct'], .05), ('보수적 메탄', 21385, cons['ghg'], 1),
    ('도전적 매립량(만 t)', 11.7, aggr['landT'] / 1e4, .05), ('도전적 매립률', 15.5, aggr['landPct'], .05), ('도전적 물질재활용률', 42.1, aggr['matPct'], .05), ('도전적 메탄', 52798, aggr['ghg'], 1),
    ('종이 매립 배출계수(tCO2eq/t)', 3.36, EF(.40), .005), ('음식물 매립 배출계수', 1.26, EF(.15), .005),
    ('대구 최근접 거리', 3.58, rb['nearestNeighbor'][0]['dist'], .005), ('2위 최근접 거리', 1.81, rb['nearestNeighbor'][1]['dist'], .005),
    ('PCA 설명력(%)', 73, sum(data['cluster']['pcaVar']) * 100, .5), ('실루엣 최소', 0.24, min(sil), .005), ('실루엣 최대', 0.26, max(sil), .005),
    ('초기값 50회 중 단독(k=3)', 47, rb['seedsAlone']['3'], 0), ('지표 제외 시 단독 유지 횟수', 2, sum(rb['leaveOneFeatureOut'].values()), 0),
    ('CSV 소각 합계 차이(t)', 10, csv.loc['소각', [MIX, SEP, FOOD]].sum() - csv.loc['소각', '총발생량'], .05),
    ('CSV 총발생량 = 국가통계', csv.loc['총발생량', '총발생량'], Tg.gen, .05), ('CSV 매립 = 국가통계', csv.loc['매립', '총발생량'], Tg.land, .05),
    ('시도·구군 인구 일치', pop['대구'], gpop['소계'], 0),
]
me = {m['item']: m for m in data['methane']}; meTot = sum(m['co2e'] for m in data['methane'])
C += [
    ('2024년 매립분 메탄 잠재량(만 tCO2eq)', 12.6, sum(EF(d_) * it('가연성', n).land for n, d_ in [('폐지류', .40), ('음식물류 폐기물', .15), ('폐섬유류', .24), ('폐목재류', .43), ('폐고무류', .39)]) / 1e4, .05),
    ('메탄 잠재량 중 종이 비중(%)', 69, EF(.40) * paper.land / meTot * 100, .5),
    ('매립량 중 종이 비중(%)', 13, paper.land / Tg.land * 100, .5),
    ('목표 시나리오 메탄 감축 비중(%)', 26, goal['ghg'] / meTot * 100, .5),
    ('덴드로그램 대구 병합 높이', 4.91, data['dendro']['daeguHeight'], .005), ('그 직전 최대 병합 높이', 3.44, data['dendro']['prevMaxHeight'], .005),
    ('대구 1인당 종량제 배출(kg/년)', 168, Mg.gen * 1000 / gpop['소계'], .5),
    ('봉투 속 종이 비중(%)', 15, paper.gen / Mg.gen * 100, .5),
    ('2024년 공공 소각량 하루(t)', 233, data['publicProc']['inc'] / 366, .5), ('공공 소각량(t)', 85295.5, data['publicProc']['inc'], .05),
    ('잔재물률을 전국 평균으로 낮출 때 재활용 증가(만 t)', 1.5, (resid['대구'] - resid['전국'].round(1)) / 100 * Sg.gen / 1e4, .05),
    ('종이 30% 전환 시 메탄 잠재 감축(만 tCO2eq)', 2.6, .3 * paper.land * EF(.40) / 1e4, .05),
]
ok = True; lines = []
for name, claimed, calc, tol in C:
    good = abs(float(claimed) - float(calc)) <= tol + 1e-9
    ok &= good
    lines.append(f"{'PASS' if good else 'FAIL'}  {name:<28} 본문 {claimed:>10}  계산 {float(calc):>12.3f}")
checks = [('대구 최근접 시도 = 대전', rb['nearestNeighbor'][0]['nearest'] == '대전'),
          ('평균 연결법에서 대구가 마지막에 병합', data['dendro']['daeguMergeStep'] == data['dendro']['steps']),
          ('구·군 유형이 Ward·평균 연결에서 동일', data['guCluster']['agree']),
          ('구·군 4유형 = 중구 / 북·달서 / 달성 / 동·서·남·수성', data['guCluster']['ward'] == sorted([['중구'], ['달서구', '북구'], ['달성군'], ['남구', '동구', '서구', '수성구']])),
          ('k=3~6 모두 대구 단독', all(rb['kmeansAloneByK'][k] for k in '3456')),
          ('계층적 군집 3종 모두 단독', all(all(v.values()) for v in rb['hierarchical'].values())),
          ('목표 매립률은 전국 2위 수준', 1 + sum(1 for x in r17 if x != '대구' and land[x] > goal['landPct']) == 2)]
for name, good in checks:
    ok &= bool(good); lines.append(f"{'PASS' if good else 'FAIL'}  {name}")
lines.append(f"\n총 {len(C) + len(checks)}개 항목 중 실패 {sum(1 for l in lines if l.startswith('FAIL'))}개")
out = '\n'.join(lines); print(out)
open(os.path.join(OUT, '분석코드', '검증결과.txt'), 'w', encoding='utf-8').write(out)
sys.exit(0 if ok else 1)
