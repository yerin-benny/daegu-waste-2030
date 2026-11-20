"""대구 생활폐기물 분석 - 부문1/부문2 공통 데이터 생성

사용법: python build_all.py <공공데이터 폴더> <출력 폴더>
입력: 전국폐기물 발생 및 처리현황(2024) 엑셀 표, 대구광역시_생활폐기물 처리_20241231.csv
출력: data.json, data.js
"""
import json, os, sys
import numpy as np, pandas as pd, openpyxl
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA

D, OUT = sys.argv[1], sys.argv[2]
P = lambda n: os.path.join(D, '전국폐기물 발생 및 처리현황 _ ' + n)
L = lambda f: list(openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))
F = lambda x: float(x or 0)
DAYS = 366  # 2024년은 윤년
ZERO = dict(gen=0, rec=0, mat=0, en=0, inc=0, land=0, etc=0)


def pick(rows, key, a, b=None, c=None):
    """key=(시도,) 또는 (시도, 시군구). 그 뒤 3개 열이 폐기물 종류 계층, 다음부터 발생량·처리량."""
    n = len(key)
    for r in rows:
        if tuple(r[:n]) == tuple(key) and r[n] == a and r[n + 1] == b and r[n + 2] == c:
            o = n + 3
            return dict(gen=F(r[o]), rec=F(r[o + 1]), mat=F(r[o + 2]), en=F(r[o + 3]),
                        inc=F(r[o + 4]), land=F(r[o + 5]), etc=F(r[o + 6]))
    return dict(ZERO)


def pct(a, b):
    return round(a / b * 100, 1) if b else None


def metrics(rows, key, pop):
    t = pick(rows, key, '총계')
    m = pick(rows, key, '종량제방식 등 혼합배출', '소계')
    rc = pick(rows, key, '재활용 가능자원 분리배출', '소계')
    rs = pick(rows, key, '재활용 가능자원 분리배출', '재활용 잔재물')
    fd = pick(rows, key, '음식물류 폐기물 분리배출')
    cw = pick(rows, key, '종량제방식 등 혼합배출', '건설폐재류', '건설폐재류(공사장생활폐기물)')
    return dict(pop=int(pop), total=round(t['gen'], 1), kgDay=round(t['gen'] * 1e3 / pop / DAYS, 2),
                mat=pct(t['mat'], t['gen']), energy=pct(t['en'], t['gen']), incin=pct(t['inc'], t['gen']),
                land=pct(t['land'], t['gen']), landT=round(t['land'], 1),
                mixT=round(m['gen'], 1), mixShare=pct(m['gen'], t['gen']), mixLand=pct(m['land'], m['gen']),
                sepT=round(rc['gen'], 1), sepPerCap=round(rc['gen'] * 1e3 / pop, 1),
                foodPerCap=round(fd['gen'] * 1e3 / pop, 1),
                residue=pct(rs['gen'], rc['gen']), residueT=round(rs['gen'], 1), foodMat=pct(fd['mat'], fd['gen']),
                constrT=round(cw['gen'], 1), constrLandT=round(cw['land'], 1), constrLandShare=pct(cw['land'], t['land']))


# ---------- 시도 ----------
G = L(P('2-가-1). (시도) 생활(가정)폐기물 발생량 (1).xlsx'))[5:]
# 주의: 같은 이름의 파일이 연도별로 여러 개 있다. 2024년 파일은 대구 인구가 시군구 표(1-나) 합계와 같은 파일이다.
POP = {r[0]: F(r[2]) for r in L(P('1-가. (시도) 생활폐기물관리구역현황.xlsx'))[4:] if r[0]}
BUD = {r[0]: r for r in L(P('4-가-1). (시도) 생활폐기물관리예산.xlsx'))[5:] if r[0]}
sido = []
for s in [r[0] for r in G if r[1] == '총계' and r[2] is None]:
    d = metrics(G, (s,), POP[s]); d['name'] = s
    b = BUD.get(s)
    if b:
        d['budgetPerTon'] = round(F(b[1]) / d['total']); d['feeRatio'] = pct(F(b[9]), F(b[1]))
    sido.append(d)

# ---------- 구·군 ----------
GG = L(P('2-나-1). (시군구) 생활(가정)폐기물 발생량 (1).xlsx'))[5:]
GPOP = {r[1]: F(r[3]) for r in L(P('1-나. (시군구) 생활폐기물관리구역현황.xlsx'))[4:] if r[1]}
GB = {r[1]: r for r in L(P('4-나-1). (시군구) 생활폐기물관리예산.xlsx'))[5:] if r[1]}
GU = ['중구', '동구', '서구', '남구', '북구', '수성구', '달서구', '달성군', '군위군']
assert POP['대구'] == GPOP['소계'] == sum(GPOP[g] for g in GU), '시도·시군구 인구 불일치(연도 확인)'
gu = []
for g in GU:
    d = metrics(GG, ('대구', g), GPOP[g]); d['name'] = g
    b = GB[g]
    d['budget'] = F(b[2]); d['budgetPerTon'] = round(F(b[2]) / d['total']); d['feeRatio'] = pct(F(b[10]), F(b[2]))
    gu.append(d)

# ---------- 대구 흐름 (배출방식 x 처리방법) ----------
K = ('대구', '소계')
flows = {'종량제(혼합)': pick(GG, K, '종량제방식 등 혼합배출', '소계'),
         '재활용품 분리배출': pick(GG, K, '재활용 가능자원 분리배출', '소계'),
         '음식물 분리배출': pick(GG, K, '음식물류 폐기물 분리배출')}

# 처리 주체별: 총계 행의 공공처리 열(재활용 소계·물질·에너지·소각·매립·기타)
_tot = next(r for r in GG if r[1] == '소계' and r[2] == '총계')
publicProc = dict(rec=F(_tot[12]), mat=F(_tot[13]), en=F(_tot[14]), inc=F(_tot[15]), land=F(_tot[16]))
print('publicProc', publicProc)

# 종량제봉투 속 성상 (대구 전체만 의미 있음 - 구·군별은 동일 비율 추정치)
mixItems = []
for a, b in [('가연성', '폐지류'), ('가연성', '폐합성수지류'), ('가연성', '음식물류 폐기물'), ('가연성', '폐섬유류'),
             ('가연성', '폐목재류'), ('가연성', '폐고무류'), ('가연성', '기타'), ('불연성', '폐금속류'), ('불연성', '폐유리류'),
             ('불연성', '폐토사류'), ('불연성', '폐타일 및 도자기류'), ('불연성', '연탄재'), ('불연성', '기타'),
             ('건설폐재류', '건설폐재류(공사장생활폐기물)')]:
    v = pick(GG, K, '종량제방식 등 혼합배출', a, b)
    mixItems.append(dict(group=a, item=b, gen=v['gen'], land=v['land'], inc=v['inc'], en=v['en']))

# 분리배출 품목
sepItems = []
for r in GG:
    if r[0] == '대구' and r[1] == '소계' and r[2] == '재활용 가능자원 분리배출' and r[3] != '소계' and F(r[5]) > 0:
        name = r[3] if not r[4] else (r[3] + '(' + r[4] + ')')
        sepItems.append(dict(item=name, gen=F(r[5])))

# ---------- 메탄 발생 잠재량 (IPCC 2006 Tier1 기본값, 장기 누적, 매립가스 회수 미반영) ----------
IPCC = dict(DOC={'폐지류': .40, '음식물류 폐기물': .15, '폐섬유류': .24, '폐목재류': .43, '폐고무류': .39},
            DOCf=.5, MCF=1.0, F=.5, OX=.1, GWP=28)
ef = {k: round(v * IPCC['DOCf'] * IPCC['MCF'] * IPCC['F'] * 16 / 12 * (1 - IPCC['OX']) * IPCC['GWP'], 3)
      for k, v in IPCC['DOC'].items()}  # tCO2eq / 매립 1t
ghg_now = sum(ef[i['item']] * i['land'] for i in mixItems if i['item'] in ef)

# ---------- AI: 17개 시도 K-means 군집 ----------
# 처리 구조를 나타내는 5개 지표만 사용한다. 1인당 발생량은 처리 구조가 아니라 발생 규모 지표이고,
# 제주(관광객 영향으로 1.73kg)가 극단값이어서 군집을 왜곡하므로 제외했다.
feat = ['mat', 'energy', 'incin', 'land', 'residue']
S = pd.DataFrame([s for s in sido if s['name'] != '전국']).set_index('name')
Z = StandardScaler().fit_transform(S[feat])
IDX = list(S.index).index('대구')
def alone(labels):  # 대구가 단독 군집인가
    return int((labels == labels[IDX]).sum()) == 1
scan = []
for k in range(2, 7):
    km = KMeans(k, n_init=200, random_state=0).fit(Z)
    scan.append(dict(k=k, silhouette=round(float(silhouette_score(Z, km.labels_)), 3), daeguAlone=alone(km.labels_),
                     labels={n: int(l) for n, l in zip(S.index, km.labels_)}))
    print('k', k, scan[-1]['silhouette'], scan[-1]['daeguAlone'])
# k=3: 대구가 분리되는 가장 단순한 군집 수. 실루엣은 k=2~6에서 0.24~0.26으로 비슷하다(군집 경계가 뚜렷하지 않음).
K_SEL = int(os.environ.get('K', 3))
km = KMeans(K_SEL, n_init=200, random_state=0).fit(Z); lab = km.labels_
pc = PCA(2).fit(Z); XY = pc.transform(Z)
clusters = [dict(id=j, members=list(S.index[lab == j]), mean=S[feat][lab == j].mean().round(1).to_dict())
            for j in range(K_SEL)]
for i, n in enumerate(S.index):
    s = next(x for x in sido if x['name'] == n)
    s['cluster'] = int(lab[i]); s['pcx'] = round(float(XY[i, 0]), 3); s['pcy'] = round(float(XY[i, 1]), 3)

# ---------- 군집 결과의 견고성 점검 ----------
from scipy.cluster.hierarchy import linkage, fcluster
from scipy.spatial.distance import cdist
DM = cdist(Z, Z); np.fill_diagonal(DM, np.inf)
nn = sorted([dict(name=S.index[i], dist=round(float(DM[i].min()), 2), nearest=S.index[int(DM[i].argmin())]) for i in range(len(S))], key=lambda x: -x['dist'])
cd = np.sqrt((Z ** 2).sum(1))
robust = dict(
    kmeansAloneByK={str(x['k']): x['daeguAlone'] for x in scan},
    seedsAlone={str(k): sum(alone(KMeans(k, n_init=10, random_state=sd).fit(Z).labels_) for sd in range(50)) for k in (3, 4)},
    hierarchical={m: {str(k): alone(fcluster(linkage(Z, m), k, 'maxclust')) for k in (3, 4)} for m in ('ward', 'average', 'complete')},
    leaveOneFeatureOut={f: alone(KMeans(K_SEL, n_init=200, random_state=0).fit(StandardScaler().fit_transform(S[[g for g in feat if g != f]])).labels_) for f in feat},
    nearestNeighbor=nn,
    centroidDist=sorted([dict(name=S.index[i], dist=round(float(cd[i]), 2)) for i in range(len(S))], key=lambda x: -x['dist']),
    daeguZ=dict(zip(feat, [round(float(v), 2) for v in Z[IDX]])))
print('robust', json.dumps({k: v for k, v in robust.items() if k not in ('nearestNeighbor', 'centroidDist')}, ensure_ascii=False))
print('NN', nn[:4]); print('centroid', robust['centroidDist'][:4])

# ---------- 덴드로그램(평균 연결법): 대구가 언제 합쳐지는가 ----------
from scipy.cluster.hierarchy import dendrogram
LK = linkage(Z, 'average'); DD = dendrogram(LK, no_plot=True, labels=list(S.index))
step = [i for i, r in enumerate(LK) if IDX in (int(r[0]), int(r[1]))][0]
dendro = dict(method='average', labels=DD['ivl'], icoord=[[round(float(v), 2) for v in r] for r in DD['icoord']],
              dcoord=[[round(float(v), 3) for v in r] for r in DD['dcoord']],
              daeguMergeStep=step + 1, steps=len(LK), daeguHeight=round(float(LK[step, 2]), 2),
              prevMaxHeight=round(float(LK[:step, 2].max()), 2))
LKW = linkage(Z, 'ward'); stepW = [i for i, r in enumerate(LKW) if IDX in (int(r[0]), int(r[1]))][0]
dendro['wardMergeStep'] = stepW + 1
print('dendro', dendro['daeguMergeStep'], dendro['steps'], dendro['daeguHeight'], dendro['prevMaxHeight'], 'ward', stepW + 1)

# ---------- 구·군 유형: 계층적 군집 (군위군은 잔재물 자료가 없어 제외) ----------
GFEAT = ['kgDay', 'mat', 'mixLand', 'residue']
GD = pd.DataFrame(gu).set_index('name').drop('군위군')[GFEAT]
GZ = StandardScaler().fit_transform(GD)
def groups(method, k):
    lab = fcluster(linkage(GZ, method), k, 'maxclust')
    return sorted([sorted(GD.index[lab == j]) for j in set(lab)])
gw, ga = groups('ward', 4), groups('average', 4)
guCluster = dict(features=GFEAT, k=4, ward=gw, average=ga, agree=gw == ga)
print('guCluster', guCluster)

# ---------- 매립 유기물 성분별 메탄 잠재량 ----------
methane = [dict(item=i['item'], land=round(i['land'], 1), co2e=round(ef[i['item']] * i['land'])) for i in mixItems if i['item'] in ef]
print('methane', methane, sum(m['co2e'] for m in methane))

src = pd.read_csv(os.path.join(D, '대구광역시_생활폐기물 처리_20241231.csv'), encoding='cp949')
# 대구시 CSV와 국가통계(2-나-1) 교차 검증
csvv = src.set_index('종류')
xl = {'총발생량': 'gen', '재활용': 'rec', '소각': 'inc', '매립': 'land', '기타': 'etc'}
tot = pick(GG, K, '총계')
check = []
for k, c in xl.items():
    colsum = float(csvv.loc[k, ['종량제방식 등 혼합배출', '재활용 가능자원 분리배출', '음식물류 폐기물 분리배출']].sum())
    check.append(dict(item=k, csvTotal=float(csvv.loc[k, '총발생량']), csvSum=round(colsum, 1), stat=round(tot[c], 1)))
print('CSV 교차검증', check)
out = dict(sido=sido, gu=gu, flows=flows, mixItems=mixItems, sepItems=sepItems,
           ipcc=dict(params=IPCC, efPerTon=ef, ghgNowMixLand=round(ghg_now)),
           cluster=dict(k=K_SEL, scan=scan, features=feat, centers=clusters,
                        pcaVar=[round(float(v), 3) for v in pc.explained_variance_ratio_],
                        pcaLoadings={f: [round(float(a), 2) for a in pc.components_[:, i]] for i, f in enumerate(feat)}, robust=robust),
           daeguCsv=src.to_dict(orient='records'), csvCheck=check, dendro=dendro, guCluster=guCluster, methane=methane, publicProc=publicProc)
json.dump(out, open(os.path.join(OUT, 'data.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
open(os.path.join(OUT, 'data.js'), 'w', encoding='utf-8').write('window.DATA=' + json.dumps(out, ensure_ascii=False) + ';')

print('clusters', json.dumps(clusters, ensure_ascii=False))
print('ef', ef, 'ghgNow', round(ghg_now))
print('대구', next(s for s in sido if s['name'] == '대구'))
print('공사장생활 매립비중', sorted([(s['name'], s['constrLandShare'], s['constrT']) for s in sido], key=lambda x: -(x[1] or 0)))
print('잔재물률', sorted([(s['name'], s['residue']) for s in sido], key=lambda x: -(x[1] or 0)))
print('budget/t', sorted([(s['name'], s.get('budgetPerTon')) for s in sido], key=lambda x: -(x[1] or 0)))
print('flows', json.dumps(flows, ensure_ascii=False))
