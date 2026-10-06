# -*- coding: utf-8 -*-
# ════════════════════════════════════════════════════════════════════
#  技術分析全攻略 · 個股評分系統（Streamlit 版，POC 分價量表）
#  由 stock_analyzer_batch_20260927 HTML 版移植
#
#  與 HTML 版的差異（本次修改）：
#   1. 分價量表訊號改以 POC（成交量最大價位區）為界，不再用 VAH/VAL：
#        守穩POC買進   ：前一日收在POC區之上，今日回測POC區、長下影線、收盤守住POC價
#        突破POC追價買進：帶量長紅K，收盤由POC區下方/區內突破到POC區之上
#        反彈POC遇壓賣出：前一日收在POC區之下，今日反彈到POC區、長上影線、收盤壓回POC價之下
#        破位停損賣出   ：帶量長黑K收盤跌破POC區（2026-09-27 回測：單獨看5/10/20日報酬
#                         皆低於全體基準，確認為偏空訊號）
#      參數（回看天數、區間數、影線倍數、量能倍數）可在側邊欄調整。
#   2. 夜星限制在高檔：首日紅K收盤須位於前20日收盤區間的70%以上。
#   3. 夜星、母子懷抱(高檔)：台股與美股回測都顯示強勢股出現這兩種型態多半是洗盤後續漲，
#      重新定義為「高檔強勢整理」，照常計入「型態成形中／突破確認／剛形成」彙總旗標。
#   4. 多因子複選搜尋／飆股搜尋：
#        - 排除「冗餘組合」：多加一個條件後樣本完全沒變（例如「母子懷抱剛形成」
#          必然同時是「型態剛形成」），這種組合不再重複列出。
#        - 加上全體基準：勝率表多了「同日超額報酬%／中位數報酬%／t值(同日調整)」
#          （每筆報酬先扣掉同一天全部評估紀錄的平均報酬，排除大盤齊漲齊跌灌大t值），
#          飆股表多了「倍數(vs基準)」，並顯示全體基準數字。
#   5. 多因子複選預設依 t值(同日調整) 排序、勝率門檻 50%；指定組合分成
#        S＝選股型（扣掉同日大盤仍有超額報酬）／#＝高勝率（標⏱為擇時型）／M＝高標股。
#
#  執行方式：
#      pip install streamlit plotly pandas numpy requests openpyxl
#      streamlit run stock_analyzer_poc.py
#  每日自動追蹤命中組合（排程用，不開介面）：
#      python stock_analyzer_poc.py daily --list top100 --token 你的FinMindToken
# ════════════════════════════════════════════════════════════════════
import gc
import math
import re
import json
import time
import io
import os
import datetime as dt
from itertools import combinations

import numpy as np
import pandas as pd
try:   # pandas 2.x：開啟 Copy-on-Write，改欄位時不必整張表複製（pandas 3 預設就是）
    pd.set_option('mode.copy_on_write', True)
except Exception:  # noqa
    pass
import requests

# ────────────────────────────────────────────────────────────────────
#  JS 相容小工具
# ────────────────────────────────────────────────────────────────────
def js_round(x):
    """JS Math.round（.5 一律進位），Python round 是銀行家捨入，結果會不同"""
    return int(math.floor(x + 0.5))


def fnum(v, d=1):
    return '--' if v is None or (isinstance(v, float) and math.isnan(v)) else f'{v:.{d}f}'


def isnum(v):
    return v is not None and not (isinstance(v, float) and math.isnan(v))


# ────────────────────────────────────────────────────────────────────
#  K 棒資料 + 技術指標（全部是「因果」算法：第 i 根的值只用到第 i 根以前的資料，
#  所以回測時整條序列算一次，任何評估點直接取用，不會偷看未來）
# ────────────────────────────────────────────────────────────────────
class Bars:
    __slots__ = ('date', 'open', 'high', 'low', 'close', 'volume', 'n',
                 'ma5', 'ma10', 'ma20', 'ma60', 'macd', 'macdSig', 'macdHist', 'rsi',
                 'bbU', 'bbL', 'vm5', 'vm20', 'kdK', 'kdD', 'kdjK', 'kdjD', 'kdjJ',
                 'plusDI', 'minusDI', 'adx', 'adxr', '_zz', '_good', '_filtered', '_fmap', '_ma100')

    def __init__(self, rows):
        # rows: list of dict(date, open, high, low, close, volume)，已依日期排序
        self.date = [r['date'] for r in rows]
        self.open = [float(r['open']) for r in rows]
        self.high = [float(r['high']) for r in rows]
        self.low = [float(r['low']) for r in rows]
        self.close = [float(r['close']) for r in rows]
        self.volume = [float(r['volume']) for r in rows]
        self.n = len(rows)
        self._zz = None
        self._good = None
        self._filtered = None
        self._fmap = None
        self._ma100 = None
        self._enrich()

    # ── 指標（逐一照 HTML 版的 JS 算法，連浮點運算順序都一致）──
    def _ma(self, arr, p):
        out = [None] * self.n
        for i in range(p - 1, self.n):
            s = 0.0
            for j in range(i - p + 1, i + 1):
                s += arr[j]
            out[i] = s / p
        return out

    def _ema(self, p):
        k = 2 / (p + 1)
        res = []
        e = self.close[0] if self.n else 0
        for i in range(self.n):
            e = self.close[i] if i == 0 else self.close[i] * k + e * (1 - k)
            res.append(e)
        return res

    def _enrich(self):
        n = self.n
        c = self.close
        self.ma5, self.ma10, self.ma20, self.ma60 = (self._ma(c, 5), self._ma(c, 10),
                                                     self._ma(c, 20), self._ma(c, 60))
        # MACD(12,26,9)
        e12, e26 = self._ema(12), self._ema(26)
        dif = [e12[i] - e26[i] for i in range(n)]
        k = 2 / 10
        sig = []
        s = dif[0] if n else 0
        for i in range(n):
            s = dif[i] if i == 0 else dif[i] * k + s * (1 - k)
            sig.append(s)
        self.macd = dif
        self.macdSig = sig
        self.macdHist = [dif[i] - sig[i] for i in range(n)]
        # RSI(14)
        p = 14
        rsi = [None] * n
        if n > p:
            ag = al = 0.0
            for i in range(1, p + 1):
                d = c[i] - c[i - 1]
                if d > 0:
                    ag += d
                else:
                    al += abs(d)
            ag /= p
            al /= p
            rsi[p] = 100 - 100 / (1 + ag / (al or 0.001))
            for i in range(p + 1, n):
                d = c[i] - c[i - 1]
                ag = (ag * (p - 1) + (d if d > 0 else 0)) / p
                al = (al * (p - 1) + (abs(d) if d < 0 else 0)) / p
                rsi[i] = 100 - 100 / (1 + ag / (al or 0.001))
        self.rsi = rsi
        # 布林通道(20,2)
        self.bbU = [None] * n
        self.bbL = [None] * n
        for i in range(19, n):
            m = self.ma20[i]
            s = 0.0
            for j in range(i - 19, i + 1):
                s += (c[j] - m) * (c[j] - m)
            sd = math.sqrt(s / 20)
            self.bbU[i] = m + 2 * sd
            self.bbL[i] = m - 2 * sd
        self.vm5 = self._ma(self.volume, 5)
        self.vm20 = self._ma(self.volume, 20)
        self.kdK, self.kdD, _ = self._kdj(9, 3, 3)
        self.kdjK, self.kdjD, self.kdjJ = self._kdj(6, 3, 3)
        self._dmi(14)

    def _kdj(self, period, kN, dN):
        kA, dA, jA = [], [], []
        pk = pd_ = 50.0
        for i in range(self.n):
            if i < period - 1:
                kA.append(None); dA.append(None); jA.append(None)
                continue
            lo = min(self.low[i - period + 1:i + 1])
            hi = max(self.high[i - period + 1:i + 1])
            rsv = 50 if hi == lo else (self.close[i] - lo) / (hi - lo) * 100
            if kN == 3 and dN == 3 and period == 9:
                k = pk * 2 / 3 + rsv * 1 / 3
                d = pd_ * 2 / 3 + k * 1 / 3
            else:
                k = pk * (kN - 1) / kN + rsv * 1 / kN
                d = pd_ * (dN - 1) / dN + k * 1 / dN
            kA.append(k); dA.append(d); jA.append(3 * k - 2 * d)
            pk, pd_ = k, d
        return kA, dA, jA

    def _dmi(self, p):
        """Wilder DMI，第 b 根的值＝把資料截到第 b 根時 calcDMI() 的回傳值"""
        n = self.n
        self.plusDI = [None] * n
        self.minusDI = [None] * n
        self.adx = [None] * n
        self.adxr = [None] * n
        if n < p + 1:
            return
        pdm, mdm, tr = [], [], []
        h, l, c = self.high, self.low, self.close
        for i in range(1, n):
            up = h[i] - h[i - 1]
            dn = l[i - 1] - l[i]
            pdm.append(up if (up > dn and up > 0) else 0)
            mdm.append(dn if (dn > up and dn > 0) else 0)
            tr.append(max(h[i] - l[i], abs(h[i] - c[i - 1]), abs(l[i] - c[i - 1])))

        def wsum(arr):
            out = []
            s = 0.0
            for j in range(p):
                s += arr[j]
            out.append(s)
            for j in range(p, len(arr)):
                s = s - s / p + arr[j]
                out.append(s)
            return out

        sTR, sP, sM = wsum(tr), wsum(pdm), wsum(mdm)
        dx = []
        for i in range(len(sTR)):
            pdi = sP[i] / sTR[i] * 100 if sTR[i] else 0
            mdi = sM[i] / sTR[i] * 100 if sTR[i] else 0
            b = i + p
            self.plusDI[b] = pdi
            self.minusDI[b] = mdi
            ss = pdi + mdi
            dx.append(abs(pdi - mdi) / ss * 100 if ss else 0)
        if len(dx) >= p:
            avg = 0.0
            for v in dx[:p]:
                avg += v
            avg /= p
            self.adx[2 * p - 1] = avg
            for k in range(p, len(dx)):
                avg = (avg * (p - 1) + dx[k]) / p
                self.adx[k + p] = avg
        for b in range(3 * p - 1, n):
            if self.adx[b] is not None and self.adx[b - p] is not None:
                self.adxr[b] = (self.adx[b] + self.adx[b - p]) / 2

    # ── 壞資料過濾（型態辨識／圖表用，跟 JS 的 filter 一致）──
    def good(self, i):
        o, h, l, c = self.open[i], self.high[i], self.low[i], self.close[i]
        return (o > 0 and h > 0 and l > 0 and c > 0 and
                all(math.isfinite(x) for x in (o, h, l, c)))

    def filtered(self):
        """回傳 (filteredBars, fmap)：fmap[i] = 原始第 i 根在過濾後序列的位置（壞K棒為 -1）"""
        if self._filtered is None:
            keep = [i for i in range(self.n) if self.good(i)]
            if len(keep) == self.n:
                self._filtered = self
                self._fmap = list(range(self.n))
            else:
                fb = object.__new__(Bars)
                for attr in ('date', 'open', 'high', 'low', 'close', 'volume', 'ma5', 'ma10', 'ma20',
                             'ma60', 'macd', 'macdSig', 'macdHist', 'rsi', 'bbU', 'bbL', 'vm5', 'vm20',
                             'kdK', 'kdD', 'kdjK', 'kdjD', 'kdjJ', 'plusDI', 'minusDI', 'adx', 'adxr'):
                    src = getattr(self, attr)
                    setattr(fb, attr, [src[i] for i in keep])
                fb.n = len(keep)
                fb._zz = None
                fb._good = None
                fb._filtered = fb
                fb._fmap = list(range(fb.n))
                fmap = [-1] * self.n
                for pos, i in enumerate(keep):
                    fmap[i] = pos
                self._filtered = fb
                self._fmap = fmap
        return self._filtered, self._fmap

    # ── 轉折波：狀態機只跑一次，記下每根K棒處理完後的狀態，任何截止點都能 O(1) 重建 ──
    def _build_zz_state(self):
        n = self.n
        start = 0
        while start < n and (self.ma5[start] is None or not self._sane(start)):
            start += 1
        st = {'start': start, 'confirmed': [], 'snap': [None] * n}
        if start >= n:
            self._zz = st
            return
        state = 'above' if self.close[start] >= self.ma5[start] else 'below'
        conf = [(start, self.low[start] if state == 'above' else self.high[start],
                 'L' if state == 'above' else 'H')]
        ext_i = start
        ext_p = self.high[start] if state == 'above' else self.low[start]
        st['snap'][start] = (1, state, ext_i, ext_p)
        for i in range(start + 1, n):
            if self.ma5[i] is not None and self._sane(i):
                if state == 'above':
                    if self.high[i] > ext_p:
                        ext_p, ext_i = self.high[i], i
                    if self.close[i] < self.ma5[i]:
                        conf.append((ext_i, ext_p, 'H'))
                        state = 'below'
                        ext_i, ext_p = i, self.low[i]
                else:
                    if self.low[i] < ext_p:
                        ext_p, ext_i = self.low[i], i
                    if self.close[i] > self.ma5[i]:
                        conf.append((ext_i, ext_p, 'L'))
                        state = 'above'
                        ext_i, ext_p = i, self.high[i]
            st['snap'][i] = (len(conf), state, ext_i, ext_p)
        st['confirmed'] = conf
        self._zz = st

    def _sane(self, i):
        h, l, c, m5 = self.high[i], self.low[i], self.close[i], self.ma5[i]
        if not (h > 0) or not (l > 0) or not (c > 0):
            return False
        if not (math.isfinite(h) and math.isfinite(l) and math.isfinite(c)):
            return False
        if m5 is not None and m5 > 0:
            if l < m5 * 0.4 or h > m5 * 2.5:
                return False
        return True

    def zigzag(self, end):
        """等同 JS buildZigzag(data.slice(0, end+1))，回傳 [(idx, price, type), ...]"""
        if self._zz is None:
            self._build_zz_state()
        st = self._zz
        n = end + 1
        if st['start'] >= n - 1:
            return []
        cnt, state, ext_i, ext_p = st['snap'][end]
        pts = st['confirmed'][:cnt]
        if ext_i != end:
            pts.append((ext_i, ext_p, 'H' if state == 'above' else 'L'))
            pts.append((end, self.close[end], 'L' if state == 'above' else 'H'))
        else:
            pts.append((ext_i, self.close[end], 'H' if state == 'above' else 'L'))
        return pts


# ────────────────────────────────────────────────────────────────────
#  多方力道評分（DMI）
# ────────────────────────────────────────────────────────────────────
def score_dmi(b: Bars, e):
    pdi, mdi, adx, adxr = b.plusDI[e], b.minusDI[e], b.adx[e], b.adxr[e]
    if pdi is None or mdi is None:
        return dict(score=0, max=100, plusDI=None, minusDI=None, adx=None, adxr=None,
                    diPts=0, adxPts=0, adxrPts=0, tdir='資料不足',
                    sigs=[('DMI資料不足（可能資料天數太短）', 'neu')])
    s = pdi + mdi
    dom = pdi / s * 100 if s else 50
    diPts = dom * 0.5
    adxPts = min(adx, 40) / 40 * 30 if adx is not None else 0
    adxrPts = 20 if (adx is not None and adxr is not None and adx > adxr) else 0
    score = js_round(max(0, min(100, diPts + adxPts + adxrPts)))
    bull = pdi > mdi
    tdir = '多頭' if bull else ('空頭' if pdi < mdi else '盤整')
    sigs = [(f"+DI {pdi:.1f}{' > ' if bull else ' < '}-DI {mdi:.1f}（多方力道{'較強' if bull else '較弱'}）",
             'bull' if bull else 'bear')]
    if adx is not None:
        lbl = '趨勢明確' if adx >= 25 else ('趨勢成形中' if adx >= 20 else '盤整')
        sigs.append((f'ADX {adx:.1f}（{lbl}）', 'bull' if adx >= 25 else 'neu'))
    if adx is not None and adxr is not None:
        up = adx > adxr
        sigs.append((f"ADX{'>' if up else '<'}ADXR，趨勢{'轉強' if up else '轉弱'}", 'bull' if up else 'bear'))
    return dict(score=score, max=100, plusDI=pdi, minusDI=mdi, adx=adx, adxr=adxr,
                diPts=diPts, adxPts=adxPts, adxrPts=adxrPts, tdir=tdir, sigs=sigs)


# ────────────────────────────────────────────────────────────────────
#  回後買上漲 8 條件
# ────────────────────────────────────────────────────────────────────
def check_pullback_buy(b: Bars, e):
    C, O, H, L = b.close, b.open, b.high, b.low
    last = e
    prev = e - 1 if e >= 1 else e
    results = []
    allPass = True
    zz = b.zigzag(e)
    zH = [p for p in zz if p[2] == 'H']
    zL = [p for p in zz if p[2] == 'L']

    c1 = False
    if len(zH) >= 2 and len(zL) >= 2:
        c1 = zH[-1][1] > zH[-2][1] and zL[-1][1] > zL[-2][1]
    results.append(dict(label='①趨勢多頭（高高低低）', pass_=c1, required=True, detail=''))
    allPass &= c1

    s0 = max(0, e - 5)
    pbd = list(range(s0, e)) if e >= 1 else []  # data.slice(-6,-1)
    had = False
    for k, i in enumerate(pbd):
        ref = C[pbd[k - 1]] if k > 0 else C[i]
        if C[i] < ref or C[i] < O[i]:
            had = True
            break
    c2 = had and C[last] > C[prev]
    detail2 = ''
    if len(zH) >= 1:
        peak = zH[-1]
        pri = [p for p in zL if p[0] < peak[0]]
        if pri:
            priorLow = pri[-1][1]
            if peak[1] > priorLow:
                plow = min(L[peak[0]:e + 1])
                retr = (peak[1] - plow) / (peak[1] - priorLow)
                grade = '最強' if retr <= 0.382 else '強' if retr <= 0.5 else '弱' if retr <= 0.618 else '回檔過深'
                detail2 = f'回檔幅度{retr * 100:.1f}%（{grade}，費波0.382/0.5/0.618分級）'
    results.append(dict(label='②位置回後上漲（近期有回檔，今轉上）', pass_=c2, required=True, detail=detail2))
    allPass &= c2

    m5 = b.ma5[last]
    c3 = m5 is not None and C[last] > m5
    results.append(dict(label='③收盤站上5MA（平價不算）', pass_=c3, required=True,
                        detail=f'5MA={m5:.2f}  收盤={C[last]:.2f}' if m5 is not None else ''))
    allPass &= c3

    c4 = H[last] > H[prev]
    results.append(dict(label='④突破前一日高點（含上影線）', pass_=c4, required=True,
                        detail=f'今高={H[last]:.2f}  昨高={H[prev]:.2f}'))
    allPass &= c4

    chg = (C[last] - C[prev]) / C[prev] * 100 if C[prev] > 0 else 0
    c5 = chg >= 2.0
    results.append(dict(label='⑤漲幅2%以上', pass_=c5, required=True, detail=f'漲幅={chg:.2f}%'))
    allPass &= c5

    body = C[last] - O[last]
    maxSh = max(H[last] - C[last], O[last] - L[last])
    c6 = body > 0 and maxSh <= body
    results.append(dict(label='⑥實體紅K，影線不大於實體', pass_=c6, required=True,
                        detail=f'實體={body:.2f}  最大影線={maxSh:.2f}'))
    allPass &= c6

    vm20 = b.vm20[last]
    vr = b.volume[last] / vm20 if vm20 else 1
    c7 = vr >= 1.0
    results.append(dict(label='⑦成交量增（加分項）', pass_=c7, required=False, detail=f'量比MA20={vr:.2f}x'))

    kNow = b.kdK[last]
    kPrev = b.kdK[e - 1] if e >= 1 else None
    c8 = kNow is not None and kPrev is not None and kNow > kPrev
    results.append(dict(label='⑧指標確認（K值向上）', pass_=c8, required=True,
                        detail=(f"K={kNow:.1f}  昨K={kPrev:.1f}" if kPrev is not None else f"K={kNow:.1f}  昨K=--")
                        if kNow is not None else 'KD資料不足'))
    allPass &= c8

    reqT = sum(1 for r in results if r['required'])
    reqP = sum(1 for r in results if r['required'] and r['pass_'])
    bonus = sum(1 for r in results if not r['required'] and r['pass_'])
    return dict(results=results, allPass=bool(allPass), requiredPassed=reqP,
                requiredTotal=reqT, bonusPassed=bonus)


# ────────────────────────────────────────────────────────────────────
#  15 種進場型態
# ────────────────────────────────────────────────────────────────────
PATTERN_DEFS = [
    ('hs', '頭肩底'), ('chs', '複式頭肩底'), ('nb', 'N字底'), ('tb', '三重底'), ('rb', '圓弧底'),
    ('fb', '一字底(均線糾結)'), ('abc', '突破ABC修正下降切線'), ('channel', '突破上升軌道線'),
    ('blackk', '突破飆股大量黑K最高點'), ('kbp', 'K線橫盤的突破'),
    ('harami_bear', '母子懷抱(高檔)'), ('harami_bull', '母子懷抱(低檔)'),
    ('morning_star', '晨星'), ('evening_star', '夜星'),
]
# 夜星、母子懷抱(高檔)：原本視為看空而排除在彙總旗標外；2026-09-27 台股＋美股回測都顯示
# 強勢股出現這兩種型態多半是洗盤後續漲（美股「夜星剛形成」20日同日超額+6.6%、t=4.5），
# 改稱「高檔強勢整理」，照常計入彙總旗標。BEARISH_PATTERNS 留空保留介面，之後要排除某型態再加回來。
BEARISH_PATTERNS = set()
HIGH_CONSOLIDATION_PATTERNS = {'harami_bear', 'evening_star'}
EVENING_STAR_HIGH_POS = 0.70   # 夜星首日紅K須位於前20日收盤區間的70%以上（高檔）

_PDESC = {
    'hs': '左右肩低點相近，頭部最低，右肩不破1/2，突破頸線為買點',
    'chs': '多重肩部低點環繞單一最低頭部，最近肩部不破1/2，突破頸線為買點',
    'nb': '低點反彈後拉回不破1/2，再突破反彈高點為買點',
    'tb': '三個低點高度相近，中間兩高點形成水平頸線，突破頸線為買點',
    'rb': '價格緩跌後緩升成U型，突破起跌點高點為買點',
    'fb': '均線糾結、價格窄幅整理（區間範圍約10%內），帶量突破整理區間為買點',
    'abc': '多頭回檔呈ABC三段式下跌，A、B高點畫下降切線，C段拉回後帶量紅K突破切線為買點',
    'channel': '股價沿上升軌道緩步上漲，MA20上揚下帶量長紅收盤突破軌道上緣為買點',
    'blackk': '飆股急漲後出現大量黑K回檔，3日內帶量長紅突破其最高點為買點',
    'kbp': '三天以上(含首日)收盤未突破首日K線高低點，帶量紅K突破首日高點為買點',
    'harami_bear': '上漲高檔出現中長紅K，次日不過高不破低的黑K線為母子懷抱。回測顯示多為高檔強勢整理後續漲，不是反轉',
    'harami_bull': '下跌低檔出現中長黑K，次日不過高不破低的紅K線為母子懷抱，空頭下跌力道轉弱',
    'morning_star': '下跌出現中長黑K+變盤線+中長紅K，收盤站上首日實體中點為低檔轉折向上訊號',
    'evening_star': '高檔(前20日區間70%以上)出現中長紅K+變盤線+中長黑K。回測顯示強勢股出現此型態多為洗盤後續漲（高檔強勢整理），不是反轉',
}


def _tolerant(a, b, pct):
    base = max(abs(a), abs(b), 1e-6)
    return abs(a - b) / base <= pct


def _res(pid, name, formed=False, breakout=False, detail='尚未偵測到符合結構', line=None, line2=None, marker=None):
    return dict(id=pid, name=name, formed=formed, breakout=breakout, detail=detail,
                desc=_PDESC.get(pid, ''), line=line, line2=line2, marker=marker, justBroke=False)


def detect_patterns14(d: Bars, e):
    """14 種圖形型態（d 必須是已過濾壞K棒的序列），回傳 list[dict]"""
    O, H, L, C, V = d.open, d.high, d.low, d.close, d.volume
    last = e
    lastClose = C[last]
    vm20 = d.vm20[last]
    volConfirm = (V[last] / vm20 >= 1.3) if vm20 else False
    zz = d.zigzag(e)
    lows = [p[0] for p in zz if p[2] == 'L' and p[0] < last]
    highs = [p[0] for p in zz if p[2] == 'H' and p[0] < last]
    nlen = e + 1

    def bo_check(res):
        if res is None:
            return False, ''
        return lastClose > res, (f'頸線/壓力＝{res:.2f}　現價＝{lastClose:.2f}' + ('　(帶量突破)' if volConfirm else ''))

    def ma_up(arr, n):
        p = arr[max(0, last - n)]
        c = arr[last]
        if c is None or p is None:
            return False
        return c > p

    out = []

    # (1) 頭肩底
    r = None
    if len(lows) >= 3:
        l3 = lows[-3:]
        L1, L2, L3 = L[l3[0]], L[l3[1]], L[l3[2]]
        if _tolerant(L1, L3, 0.06) and L2 < L1 * 0.985 and L2 < L3 * 0.985:
            hb = [i for i in highs if l3[0] < i < l3[2]]
            neck = max(H[i] for i in hb) if hb else None
            if neck is not None and L3 > (L2 + neck) / 2:
                ok, det = bo_check(neck)
                r = _res('hs', '頭肩底', True, ok, det or '型態成形，等待突破頸線', line=(l3[0], neck, 0))
    out.append(r or _res('hs', '頭肩底'))

    # (2) 複式頭肩底
    r = None
    if len(lows) >= 4:
        ll = lows[-5:]
        vals = [L[i] for i in ll]
        mn = min(vals)
        hp = vals.index(mn)
        sv = [v for k, v in enumerate(vals) if k != hp]
        ok_sh = hp > 0 and hp < len(ll) - 1 and all(_tolerant(v, sv[0], 0.08) and v > mn * 1.02 for v in sv)
        if ok_sh:
            hb = [i for i in highs if ll[0] < i < ll[-1]]
            neck = max(H[i] for i in hb) if hb else None
            if neck is not None and vals[-1] > (mn + neck) / 2:
                ok, det = bo_check(neck)
                r = _res('chs', '複式頭肩底', True, ok, det or '型態成形，等待突破頸線', line=(ll[0], neck, 0))
    out.append(r or _res('chs', '複式頭肩底'))

    # (3) N字底
    r = None
    if len(lows) >= 2 and len(highs) >= 1:
        A, Cc = lows[-2], lows[-1]
        bc = [i for i in highs if A < i < Cc]
        if bc:
            B = bc[0]
            for i in bc:
                if H[i] > H[B]:
                    B = i
            half = (L[A] + H[B]) / 2
            if L[Cc] > half:
                ok, det = bo_check(H[B])
                r = _res('nb', 'N字底', True, ok, det or f'拉回未破1/2（{half:.2f}），等待突破反彈高點',
                         line=(B, H[B], 0))
    out.append(r or _res('nb', 'N字底'))

    # (4) 三重底
    r = None
    if len(lows) >= 3:
        l3 = lows[-3:]
        v = [L[i] for i in l3]
        if _tolerant(v[0], v[1], .08) and _tolerant(v[1], v[2], .08) and _tolerant(v[0], v[2], .08):
            h1 = [i for i in highs if l3[0] < i < l3[1]]
            h2 = [i for i in highs if l3[1] < i < l3[2]]
            p1 = max(H[i] for i in h1) if h1 else None
            p2 = max(H[i] for i in h2) if h2 else None
            if p1 is not None and p2 is not None and _tolerant(p1, p2, 0.05):
                res = max(p1, p2)
                ok, det = bo_check(res)
                r = _res('tb', '三重底', True, ok, det or '型態成形，等待突破壓力', line=(l3[0], res, 0))
    out.append(r or _res('tb', '三重底'))

    # (5) 圓弧底
    r = None
    ws = max(0, nlen - 40)
    wl = nlen - ws
    if wl >= 30:
        seg = wl // 3
        first = list(range(ws, ws + seg))
        mid = list(range(ws + seg, ws + wl - seg))
        tail = list(range(ws + wl - seg, ws + wl))

        def slope(idx):
            n = len(idx)
            sx = sy = sxy = sxx = 0.0
            for k, i in enumerate(idx):
                sx += k; sy += C[i]; sxy += k * C[i]; sxx += k * k
            den = (n * sxx - sx * sx) or 1
            return (n * sxy - sx * sy) / den

        def avgc(idx):
            s = 0.0
            for i in idx:
                s += C[i]
            return s / len(idx)

        midLow = min(L[i] for i in mid)
        convex = avgc(first) > midLow * 1.01 and avgc(tail) > midLow * 1.01
        shape = slope(first) < 0 and slope(tail) > 0 and convex
        ar = 0.0
        for i in range(ws, nlen):
            ar += (H[i] - L[i]) / C[i]
        ar /= wl
        if shape and ar < 0.05:
            res = max(H[i] for i in first)
            ok, det = bo_check(res)
            target = res + (res - midLow)
            r = _res('rb', '圓弧底', True, ok, (det or '弧形築底中，等待突破起跌壓力') + f'　目標價≈{target:.2f}',
                     line=(ws, res, 0))
    out.append(r or _res('rb', '圓弧底'))

    # (6) 一字底(均線糾結)
    r = None
    N = 10
    w0 = max(0, nlen - N - 1)
    win = list(range(w0, e))  # slice(-11,-1)
    if len(win) == N and all(d.ma5[i] is not None and d.ma10[i] is not None and d.ma20[i] is not None
                             and d.ma60[i] is not None for i in win):
        tangled = True
        for i in win:
            vals = [d.ma5[i], d.ma10[i], d.ma20[i], d.ma60[i]]
            if (max(vals) - min(vals)) / min(vals) > 0.05:
                tangled = False
                break
        wh = max(H[i] for i in win)
        wlw = min(L[i] for i in win)
        if tangled and (wh - wlw) / wlw <= 0.10:
            l4 = [d.ma5[last], d.ma10[last], d.ma20[last], d.ma60[last]]
            above = all(v is not None and lastClose > v for v in l4)
            bo = lastClose > wh and above and volConfirm
            t6 = wh + (wh - wlw)
            r = _res('fb', '一字底(均線糾結)', True, bo,
                     f'整理區間高點＝{wh:.2f}　現價＝{lastClose:.2f}' + ('　(帶量突破)' if volConfirm else '　(尚未帶量)')
                     + f'　目標價≈{t6:.2f}', line=(nlen - N - 1, wh, 0))
    out.append(r or _res('fb', '一字底(均線糾結)'))

    # (7) 突破ABC修正下降切線
    r = None
    rh = [i for i in highs if i >= last - 40]
    if len(rh) >= 2:
        h1, h2 = rh[-2], rh[-1]
        y1, y2 = H[h1], H[h2]
        low_between = any(h1 < li < h2 for li in lows)
        cLow = None
        for ci in range(h2 + 1, last):
            if cLow is None or L[ci] < cLow:
                cLow = L[ci]
        hasC = cLow is not None and cLow < C[h2] * 0.98
        if y2 < y1 and h2 > h1 and low_between and hasC and (h2 - h1) <= 20 and (last - h2) <= 20:
            sl = (y2 - y1) / (h2 - h1)
            line_at = y1 + sl * (last - h1)
            up20 = ma_up(d.ma20, 10)
            isRed = C[last] > O[last]
            bo = lastClose > line_at and up20 and isRed
            r = _res('abc', '突破ABC修正下降切線', True, bo,
                     f'下降切線位置≈{line_at:.2f}　C低點＝{cLow:.2f}　現價＝{lastClose:.2f}'
                     + ('　MA20上揚' if up20 else '　MA20未上揚'), line=(h1, y1, sl))
    out.append(r or _res('abc', '突破ABC修正下降切線'))

    # (8) 突破上升軌道線
    r = None
    fl = max(0, nlen - 60)
    li_ = [i for i in lows if i >= fl]
    hi_ = [i for i in highs if i >= fl]
    if len(li_) >= 2:
        l1, l2 = li_[-2], li_[-1]
        ly1, ly2 = L[l1], L[l2]
        if ly2 > ly1 and l2 > l1:
            s2 = (ly2 - ly1) / (l2 - l1)
            offs = [H[i] - (ly1 + s2 * (i - l1)) for i in hi_ if i > l1]
            offset = None
            if len(offs) >= 2:
                tol = lastClose * 0.03
                best, bc_ = None, 0
                for o in offs:
                    cnt = sum(1 for o2 in offs if abs(o - o2) <= tol)
                    if cnt > bc_ or (cnt == bc_ and (best is None or o > best)):
                        bc_, best = cnt, o
                if bc_ >= 2:
                    offset = best
            if offset is not None and offset > 0:
                upper = ly1 + s2 * (last - l1) + offset
                up20 = ma_up(d.ma20, 10)
                isRed = C[last] > O[last]
                bo = lastClose > upper and up20 and isRed and volConfirm
                r = _res('channel', '突破上升軌道線', True, bo,
                         f'軌道上緣≈{upper:.2f}　現價＝{lastClose:.2f}' + ('　帶量' if volConfirm else '　量未放大'),
                         line=(l1, ly1 + offset, s2), line2=(l1, ly1, s2))
    out.append(r or _res('channel', '突破上升軌道線'))

    # (9) 突破飆股大量黑K最高點
    r = None
    f3 = max(0, nlen - 1 - 10)
    cands = [i for i in range(f3, last) if C[i] < O[i] and d.vm20[i] and V[i] / d.vm20[i] >= 1.6]
    if cands:
        bk = cands[-1]
        if 0 <= last - bk <= 3:
            isRed = C[last] > O[last]
            up20 = ma_up(d.ma20, 10)
            bo = lastClose > H[bk] and isRed and volConfirm and up20
            r = _res('blackk', '突破飆股大量黑K最高點', True, bo,
                     f'大量黑K高點＝{H[bk]:.2f}　現價＝{lastClose:.2f}' + ('　帶量' if volConfirm else '　量未放大'),
                     line=(bk, H[bk], 0))
    out.append(r or _res('blackk', '突破飆股大量黑K最高點'))

    # (10) K線橫盤的突破
    r = None
    if nlen >= 4:
        anc = last - 3
        ok = all(not (C[j] > H[anc] or C[j] < L[anc]) for j in range(anc + 1, last))
        if ok:
            cand = anc - 1
            while cand >= 0 and (last - cand) <= 20:
                inr = all(not (C[k] > H[cand] or C[k] < L[cand]) for k in range(cand + 1, last))
                if not inr:
                    break
                anc = cand
                cand -= 1
            isRed = C[last] > O[last]
            bo = lastClose > H[anc] and isRed and volConfirm
            r = _res('kbp', 'K線橫盤的突破', True, bo,
                     f'首日K線高點＝{H[anc]:.2f}　整理天數＝{last - anc}天　現價＝{lastClose:.2f}'
                     + ('　(帶量)' if volConfirm else '　(量未放大)'), line=(anc, H[anc], 0))
    out.append(r or _res('kbp', 'K線橫盤的突破'))

    # (11) 母子懷抱(高檔)
    r = None
    if nlen >= 3:
        m, ch = last - 2, last - 1
        mb = abs(C[m] - O[m]) / C[m]
        contained = H[ch] <= H[m] and L[ch] >= L[m]
        smaller = abs(C[ch] - O[ch]) < abs(C[m] - O[m]) * 0.6
        atHigh = C[m] >= C[last - 3] if last - 3 >= 0 else True
        if C[m] > O[m] and mb >= 0.035 and contained and smaller and atHigh:
            bo = C[last] < C[ch]
            r = _res('harami_bear', '母子懷抱(高檔)', True, bo,
                     f'母K(中長紅)高點＝{H[m]:.2f}　子K收於母K範圍內　' + ('次日已確認轉折向下' if bo else '等待次日確認轉折向下'),
                     marker=(last - 1, H[m], 'up', '母子懷抱'))
    out.append(r or _res('harami_bear', '母子懷抱(高檔)'))

    # (12) 母子懷抱(低檔)
    r = None
    if nlen >= 3:
        m, ch = last - 2, last - 1
        mb = abs(C[m] - O[m]) / C[m]
        contained = H[ch] <= H[m] and L[ch] >= L[m]
        smaller = abs(C[ch] - O[ch]) < abs(C[m] - O[m]) * 0.6
        atLow = C[m] <= C[last - 3] if last - 3 >= 0 else True
        if C[m] < O[m] and mb >= 0.035 and contained and smaller and atLow:
            bo = C[last] > C[ch]
            r = _res('harami_bull', '母子懷抱(低檔)', True, bo,
                     f'母K(中長黑)低點＝{L[m]:.2f}　子K收於母K範圍內　' + ('次日已確認轉折向上' if bo else '等待次日確認轉折向上'),
                     marker=(last - 1, L[m], 'down', '母子懷抱'))
    out.append(r or _res('harami_bull', '母子懷抱(低檔)'))

    # (13) 晨星
    r = None
    if nlen >= 3:
        s1, s2, s3 = last - 2, last - 1, last
        mid1 = (O[s1] + C[s1]) / 2
        if (C[s1] < O[s1] and abs(C[s1] - O[s1]) / C[s1] >= 0.035 and abs(C[s2] - O[s2]) / C[s2] < 0.035
                and C[s3] > O[s3] and abs(C[s3] - O[s3]) / C[s3] >= 0.035 and C[s3] > mid1):
            r = _res('morning_star', '晨星', True, True,
                     f'首日黑K實體中點＝{mid1:.2f}　收盤＝{C[s3]:.2f}（已站上中點，轉折確認）',
                     marker=(last - 1, L[s2], 'down', '晨星'))
    out.append(r or _res('morning_star', '晨星'))

    # (14) 夜星（修正：限制在高檔）
    r = None
    if nlen >= 3:
        e1, e2, e3 = last - 2, last - 1, last
        mid1 = (O[e1] + C[e1]) / 2
        w = C[max(0, e1 - 19):e1 + 1]       # 含首日在內的前20日收盤
        rng = max(w) - min(w)
        pos = (C[e1] - min(w)) / rng if rng > 0 else 0
        atHigh = len(w) >= 10 and pos >= EVENING_STAR_HIGH_POS
        if (atHigh and C[e1] > O[e1] and abs(C[e1] - O[e1]) / C[e1] >= 0.035
                and abs(C[e2] - O[e2]) / C[e2] < 0.035
                and C[e3] < O[e3] and abs(C[e3] - O[e3]) / C[e3] >= 0.035 and C[e3] < mid1):
            r = _res('evening_star', '夜星', True, True,
                     f'首日紅K位於前20日區間{pos * 100:.0f}%　實體中點＝{mid1:.2f}　收盤＝{C[e3]:.2f}（已跌破中點，轉折確認）',
                     marker=(last - 1, H[e2], 'up', '夜星'))
    out.append(r or _res('evening_star', '夜星'))
    return out


def _pbup_entry(pb):
    formed = pb['allPass'] or pb['requiredPassed'] >= pb['requiredTotal'] - 1
    return dict(id='pbup', name='回後買上漲', formed=formed, breakout=pb['allPass'],
                detail=f"必要條件 {pb['requiredPassed']}/{pb['requiredTotal']} 通過" + ('　+成交量增加分' if pb['bonusPassed'] else ''),
                desc='趨勢多頭，回檔量縮價穩後，今日紅K放量突破前高為買點',
                line=None, line2=None, marker=None, justBroke=False)


class PatternCache:
    """同一檔股票逐日回測時，昨天的型態結果今天拿來判斷『剛突破』，不用重算"""

    def __init__(self, b: Bars):
        self.b = b
        self.fb, self.fmap = b.filtered()
        self.p14 = {}
        self.pbf = {}

    def p14_at(self, f):
        if f not in self.p14:
            self.p14[f] = detect_patterns14(self.fb, f)
        return self.p14[f]

    def pb_at(self, f):
        if f not in self.pbf:
            self.pbf[f] = check_pullback_buy(self.fb, f)
        return self.pbf[f]

    def detect(self, e, pb):
        """等同 JS detectPatterns(data.slice(0,e+1), pb)"""
        # 過濾後的截止位置＝e 以前（含）最後一根好K棒
        f = self.fmap[e]
        if f < 0:
            f = max([self.fmap[i] for i in range(e + 1) if self.fmap[i] >= 0] or [-1])
        if f < 0:
            return dict(results=[], anyBreakout=False, anyFormed=False, anyJustBroke=False)
        res = [dict(x) for x in self.p14_at(f)] + [_pbup_entry(pb)]
        if f >= 1:
            prev = self.p14_at(f - 1) + [_pbup_entry(self.pb_at(f - 1))]
            for r, pr in zip(res, prev):
                r['justBroke'] = bool(r['breakout'] and not pr['breakout'])
        bull = [r for r in res if r['id'] not in BEARISH_PATTERNS]
        return dict(results=res,
                    anyBreakout=any(r['breakout'] for r in bull),
                    anyFormed=any(r['formed'] for r in bull),
                    anyJustBroke=any(r['justBroke'] for r in bull))


# ────────────────────────────────────────────────────────────────────
#  布林／MACD／KDJ 綜合訊號
# ────────────────────────────────────────────────────────────────────
def core_signals(b: Bars, e):
    C = b.close
    prev = e - 1 if e >= 1 else e
    bbPos = None
    if b.bbU[e] is not None and b.bbL[e] is not None:
        w = b.bbU[e] - b.bbL[e]
        bbPos = (C[e] - b.bbL[e]) / w * 100 if w else 50
    macdState = None
    isBreakout = isPR = gcr = False
    div = False
    if b.macdHist[e] is not None and b.macdHist[prev] is not None:
        above = b.macd[e] > 0
        grow = b.macdHist[e] > b.macdHist[prev]
        if above and b.macdHist[e] > 0:
            macdState = '零軸上・紅柱增長' if grow else '零軸上・紅柱縮短'
        elif not above and b.macdHist[e] < 0:
            macdState = '零軸下・綠柱縮短' if grow else '零軸下・綠柱增長'
        else:
            macdState = '交叉轉換中'
        for i in range(max(1, e + 1 - 3), e + 1):
            if b.macd[i - 1] <= b.macdSig[i - 1] and b.macd[i] > b.macdSig[i]:
                gcr = True
                break
        widthExp = False
        if b.bbU[e] is not None and b.bbL[e] is not None:
            wn = (b.bbU[e] - b.bbL[e]) / C[e] * 100 if C[e] else 0
            ri = e - 5
            if ri >= 0 and b.bbU[ri] is not None and b.bbL[ri] is not None and C[ri]:
                widthExp = wn > (b.bbU[ri] - b.bbL[ri]) / C[ri] * 100
        isBreakout = bbPos is not None and bbPos >= 80 and widthExp and above and b.macdHist[e] > 0 and grow
        zl = [p for p in b.zigzag(e) if p[2] == 'L']
        if len(zl) >= 2:
            rl, pl = zl[-1], zl[-2]
            if rl[0] >= e - 60:
                div = rl[1] < pl[1] and b.macd[rl[0]] > b.macd[pl[0]]
        isPR = bbPos is not None and bbPos <= 20 and div and gcr
    kdjState = None
    kg = kd = False
    if b.kdjK[e] is not None and b.kdjD[e] is not None:
        kdjState = 'K>D（多方）' if b.kdjK[e] > b.kdjD[e] else ('K<D（空方）' if b.kdjK[e] < b.kdjD[e] else 'K=D')
        for i in range(max(1, e + 1 - 3), e + 1):
            a0, d0, a1, d1 = b.kdjK[i - 1], b.kdjD[i - 1], b.kdjK[i], b.kdjD[i]
            if None not in (a0, d0, a1, d1):
                if a0 <= d0 and a1 > d1:
                    kg = True
                if a0 >= d0 and a1 < d1:
                    kd = True
    return dict(bbPos=bbPos, macdState=macdState, isBreakout=isBreakout, isPullbackRebound=isPR,
                goldenCrossRecent=gcr, kdjState=kdjState, kdjGoldenCrossRecent=kg,
                kdjDeathCrossRecent=kd, divergence=div)


# ────────────────────────────────────────────────────────────────────
#  大盤（0050）、相對強弱、量能
# ────────────────────────────────────────────────────────────────────
class Benchmark:
    def __init__(self, rows):
        rows = sorted(rows, key=lambda r: r['date'])
        self.date = [r['date'] for r in rows]
        self.close = [float(r['close']) for r in rows]
        n = len(rows)
        self.ma20 = [None] * n
        self.ma60 = [None] * n
        for i in range(n):
            if i >= 19:
                s = 0.0
                for j in range(i - 19, i + 1):
                    s += self.close[j]
                self.ma20[i] = s / 20
            if i >= 59:
                s = 0.0
                for j in range(i - 59, i + 1):
                    s += self.close[j]
                self.ma60[i] = s / 60
        self.idx = {d: i for i, d in enumerate(self.date)}

    def find(self, ds):
        if ds in self.idx:
            return self.idx[ds]
        import bisect
        return bisect.bisect_right(self.date, ds) - 1

    def regime(self, ds):
        i = self.find(ds)
        if i < 0:
            return None, None
        c = self.close[i]
        return (None if self.ma20[i] is None else c > self.ma20[i],
                None if self.ma60[i] is None else c > self.ma60[i])


def relative_strength(b: Bars, e, bm, lb=20):
    if bm is None or not bm.close or e + 1 <= lb:
        return None
    pc = b.close[e - lb]
    if not pc:
        return None
    sr = (b.close[e] - pc) / pc * 100
    bi = bm.find(b.date[e])
    if bi < lb:
        return None
    bp = bm.close[bi - lb]
    if not bp:
        return None
    return sr - (bm.close[bi] - bp) / bp * 100


def vol_ratio(b: Bars, e):
    v = b.vm20[e]
    return b.volume[e] / v if v else None


def vol_range_pos(b: Bars, e):
    if e + 1 < 20:
        return None
    w = b.volume[max(0, e + 1 - 252):e + 1]
    if len(w) < 20:
        return None
    t = b.volume[e]
    return sum(1 for v in w if v < t) / len(w) * 100


def vol_trend(b: Bars, e):
    if e + 1 < 10:
        return None
    a10 = sum(b.volume[e - 9:e + 1]) / 10
    a5 = sum(b.volume[e - 4:e + 1]) / 5
    if not a10:
        return None
    return (a5 - a10) / a10 * 100


# ────────────────────────────────────────────────────────────────────
#  個股分價量表（Volume Profile）── POC 版
# ────────────────────────────────────────────────────────────────────
VP_DEFAULTS = dict(lookback=120, bins=24, shadow_mult=1.0, vol_mult=1.2, body_mult=1.2, touch_tol=0.5)

VP_LABELS = {
    'vpBuySupport': '分價量表-守穩POC買進',
    'vpBuyBreakout': '分價量表-突破POC追價買進',
    'vpSellResistance': '分價量表-反彈POC遇壓賣出',
    'vpBreakdown': '分價量表-破位停損賣出',
}
VP_SIGTYPE_FIELD = {'buy_support': 'vpBuySupport', 'buy_breakout': 'vpBuyBreakout',
                    'sell_resistance': 'vpSellResistance', 'sell_breakdown': 'vpBreakdown'}


def volume_profile(b: Bars, e, lookback=120, bins=24):
    if e + 1 < 20:
        return None
    s = max(0, e + 1 - lookback)
    lo = min(b.low[s:e + 1])
    hi = max(b.high[s:e + 1])
    if not hi > lo:
        return None
    bs = (hi - lo) / bins
    vol = [0.0] * bins
    for i in range(s, e + 1):
        mid = (b.high[i] + b.low[i]) / 2
        k = min(bins - 1, max(0, int(math.floor((mid - lo) / bs))))
        vol[k] += b.volume[i]
    poc = 0
    for i in range(1, bins):
        if vol[i] > vol[poc]:
            poc = i
    tot = sum(vol)
    li = hi_i = poc
    inc = vol[poc]
    while inc < tot * 0.7 and (li > 0 or hi_i < bins - 1):
        nl = vol[li - 1] if li > 0 else -1
        nh = vol[hi_i + 1] if hi_i < bins - 1 else -1
        if nh >= nl:
            hi_i += 1
            inc += vol[hi_i]
        else:
            li -= 1
            inc += vol[li]
    return dict(pocLow=lo + poc * bs, pocHigh=lo + (poc + 1) * bs, pocPrice=lo + (poc + 0.5) * bs,
                valLow=lo + li * bs, valHigh=lo + (hi_i + 1) * bs, binSize=bs, lo=lo, hi=hi)


def vp_signal(b: Bars, e, prof, p=None):
    """以 POC 區（成交量最大價位區間）為界的4種訊號"""
    p = p or VP_DEFAULTS
    base = dict(signal=None, sigType=None, detail=None,
                pocPrice=prof['pocPrice'] if prof else None,
                pocLow=prof['pocLow'] if prof else None, pocHigh=prof['pocHigh'] if prof else None,
                vah=prof['valHigh'] if prof else None, val=prof['valLow'] if prof else None)
    if not prof or e + 1 < 22:
        return base
    O, H, L, C, V = b.open, b.high, b.low, b.close, b.volume
    pc = C[e - 1]
    isRed, isBlack = C[e] > O[e], C[e] < O[e]
    body = abs(C[e] - O[e])
    upSh = H[e] - max(O[e], C[e])
    dnSh = min(O[e], C[e]) - L[e]
    rb = range(e - 20, e)
    avgBody = sum(abs(C[i] - O[i]) for i in rb) / 20
    avgVol = sum(V[i] for i in rb) / 20
    longRed = isRed and body > avgBody * p['body_mult']
    longBlack = isBlack and body > avgBody * p['body_mult']
    hasVol = avgVol > 0 and V[e] > avgVol * p['vol_mult']
    pL, pH, pP = prof['pocLow'], prof['pocHigh'], prof['pocPrice']
    tol = prof['binSize'] * p['touch_tol']
    longLower = dnSh > 0 and dnSh >= body * p['shadow_mult']
    longUpper = upSh > 0 and upSh >= body * p['shadow_mult']

    support = pc > pH and L[e] <= pH and L[e] >= pL - tol and C[e] >= pP and longLower
    breakout = pc <= pH and C[e] > pH and longRed and hasVol
    resist = pc < pL and H[e] >= pL and H[e] <= pH + tol and C[e] <= pP and longUpper
    breakdown = pc >= pL and C[e] < pL and longBlack and hasVol

    if support:
        base.update(signal='buy', sigType='buy_support',
                    detail='守穩POC買進：股價由上方回測POC成交密集區，長下影線且收盤守住POC價（停損設POC區之下）')
    elif breakout:
        base.update(signal='buy', sigType='buy_breakout',
                    detail='突破POC追價買進：帶量長紅K由POC區突破到上方，買方消化最大套牢量')
    elif resist:
        base.update(signal='sell', sigType='sell_resistance',
                    detail='反彈POC遇壓賣出：股價由下方反彈到POC成交密集區，長上影線且收盤壓回POC價之下')
    elif breakdown:
        base.update(signal='sell', sigType='sell_breakdown',
                    detail='破位停損賣出：帶量長黑K跌破POC區，賣方潰散，進入成交量真空區（停損/減碼）')
    return base


# ────────────────────────────────────────────────────────────────────
#  旗標列（批次即時／回測共用）
# ────────────────────────────────────────────────────────────────────
PH_FIELDS = ['ph_' + pid for pid, _ in PATTERN_DEFS] + ['ph_pbup']



# ── 新增技術面欄位（2026-09-30）：52週高點、均線多頭排列、布林收窄、向上跳空 ──
LONG_HISTORY_DAYS = 400   # 52週高點需要約250根K棒；批次分析會多抓到400天只用來算這幾個欄位


def tech_extras(b, e):
    """52週高點距離／創新高、均線多頭排列、布林寬度在近半年的百分位、近3日向上跳空、
    3日內突破季線／半年線、連續放量、MACD零軸下金叉。資料不足時為 None"""
    out = dict(high52Dist=None, newHigh52=None, maBull=None, bbwRank=None, gapUp3=None,
               crossMa60_3=None, crossMa100_3=None, vol3Ratio=None, gcBelow0_3=None,
               udVolRatio20=None, cmf20=None, obvNewHigh60=None,
               bias20=None, pullMa60=None, limitUp3=None, drop5=None, nr7=None)
    if e < 0 or e >= b.n:
        return out
    if e >= 249:
        mx = max(b.high[e - 249:e + 1])
        prev = max(b.high[e - 249:e])
        out['high52Dist'] = (b.close[e] / mx - 1) * 100 if mx > 0 else None
        out['newHigh52'] = int(b.high[e] > prev)
    m5, m20, m60 = b.ma5[e], b.ma20[e], b.ma60[e]
    if None not in (m5, m20, m60):
        out['maBull'] = int(m5 > m20 > m60 and b.close[e] > m20)

    def bw(i):
        u, l = b.bbU[i], b.bbL[i]
        if u is None or l is None or (u + l) <= 0:
            return None
        return (u - l) / ((u + l) / 2)
    cur = bw(e)
    if cur is not None and e >= 119:
        ws = [w for w in (bw(i) for i in range(e - 119, e + 1)) if w is not None]
        if len(ws) >= 100:
            out['bbwRank'] = sum(1 for w in ws if w <= cur) / len(ws) * 100
    if e >= 3:
        out['gapUp3'] = int(any(b.low[e - k] > b.high[e - k - 1] for k in range(3)))
    # 3日內收盤由下往上突破季線(MA60)／半年線(MA100)，且今天仍站在線上（2026-10-03 參考 BBW+MACD 掃描器新增）
    if b._ma100 is None:
        b._ma100 = b._ma(b.close, 100)
    for fld, ma in (('crossMa60_3', b.ma60), ('crossMa100_3', b._ma100)):
        if e >= 3 and None not in (ma[e - 3], ma[e]):
            out[fld] = int(b.close[e] >= ma[e] and any(
                b.close[j - 1] < ma[j - 1] and b.close[j] >= ma[j] for j in range(e - 2, e + 1)))
    # 連續放量：近3日均量 ÷ 再往前20日均量
    if e >= 22:
        base = sum(b.volume[e - 22:e - 2]) / 20
        out['vol3Ratio'] = (sum(b.volume[e - 2:e + 1]) / 3) / base if base > 0 else None
    # MACD 零軸下（DIF<0）3日內黃金交叉＝低檔轉折
    if e >= 3:
        out['gcBelow0_3'] = int(any(b.macd[j - 1] <= b.macdSig[j - 1] and b.macd[j] > b.macdSig[j] and b.macd[j] < 0
                                    for j in range(e - 2, e + 1)))
    # 2026-10-04 買賣力道近似（日線沒有內外盤，用收盤在當日區間的位置／漲跌日量推估）
    out.update(dict(udVolRatio20=None, cmf20=None, obvNewHigh60=None))
    if e >= 20:
        up = dn = 0.0
        mfv = vs = 0.0
        for j in range(e - 19, e + 1):
            v = b.volume[j]
            if b.close[j] > b.close[j - 1]:
                up += v
            elif b.close[j] < b.close[j - 1]:
                dn += v
            rng = b.high[j] - b.low[j]
            if rng > 0:
                mfv += ((b.close[j] - b.low[j]) - (b.high[j] - b.close[j])) / rng * v
            vs += v
        out['udVolRatio20'] = min(up / dn, 99.0) if dn > 0 else (99.0 if up > 0 else None)
        out['cmf20'] = mfv / vs if vs > 0 else None
    if e >= 61:
        # OBV 從 e-60 起算（起點＝0），今天的 OBV 高於前60天每一天＝OBV創60日新高
        obv, mx = 0.0, 0.0
        for j in range(e - 59, e):
            obv += b.volume[j] if b.close[j] > b.close[j - 1] else (-b.volume[j] if b.close[j] < b.close[j - 1] else 0.0)
            mx = max(mx, obv)
        obv += b.volume[e] if b.close[e] > b.close[e - 1] else (-b.volume[e] if b.close[e] < b.close[e - 1] else 0.0)
        out['obvNewHigh60'] = int(obv > mx)
    # 2026-10-05 新增：短線乖離、多頭回測季線、近3日漲停（台股）、近5日跌幅與NR7窄幅日（美股）
    out.update(dict(bias20=None, pullMa60=None, limitUp3=None, drop5=None, nr7=None))
    if b.ma20[e]:
        out['bias20'] = (b.close[e] / b.ma20[e] - 1) * 100
    if e >= 10 and b.ma60[e] and b.ma60[e - 10]:
        out['pullMa60'] = int(b.ma60[e] > b.ma60[e - 10] and abs(b.close[e] / b.ma60[e] - 1) <= 0.03)
    if e >= 3:
        out['limitUp3'] = int(any(b.close[j - 1] > 0 and b.close[j] / b.close[j - 1] - 1 >= 0.095 for j in range(e - 2, e + 1)))
    if e >= 5 and b.close[e - 5]:
        out['drop5'] = (b.close[e] / b.close[e - 5] - 1) * 100
    if e >= 6:
        rg = [b.high[j] - b.low[j] for j in range(e - 6, e + 1)]
        out['nr7'] = int(rg[-1] < min(rg[:-1]))
    return out


def trim_rows(rows, days):
    """只保留最近 days 天（和直接抓 days 天的結果相同）"""
    cut = (dt.date.today() - dt.timedelta(days=days)).isoformat()
    return [r for r in rows if r['date'] >= cut]


def build_flag_row(b: Bars, e, bm, pc: PatternCache, vp_params=None, extras=None):
    """回傳 (row, info)：row 是回測／條件旗標用的欄位，info 是批次表格顯示用的明細"""
    vpp = vp_params or VP_DEFAULTS
    dm = score_dmi(b, e)
    sig = core_signals(b, e)
    rs = relative_strength(b, e, bm, 20)
    vr = vol_ratio(b, e)
    vrp = vol_range_pos(b, e)
    vt = vol_trend(b, e)
    prof = volume_profile(b, e, vpp['lookback'], vpp['bins'])
    vps = vp_signal(b, e, prof, vpp)
    a20, a60 = bm.regime(b.date[e]) if bm else (None, None)
    pb = check_pullback_buy(b, e)
    pt = pc.detect(e, pb)
    row = dict(
        evalDate=b.date[e], score=dm['score'], plusDI=dm['plusDI'], minusDI=dm['minusDI'],
        adx=dm['adx'], adxr=dm['adxr'], bbPos=sig['bbPos'],
        isBreakout=int(sig['isBreakout']), isPullbackRebound=int(sig['isPullbackRebound']),
        goldenCrossRecent=int(sig['goldenCrossRecent']),
        kdjGoldenCrossRecent=int(sig['kdjGoldenCrossRecent']), kdjDeathCrossRecent=int(sig['kdjDeathCrossRecent']),
        relStrength20=rs, volRatio=vr, volRangePos=vrp, volTrend=vt,
        vpPocPrice=vps['pocPrice'],
        benchmarkAbove20=None if a20 is None else int(a20), benchmarkAbove60=None if a60 is None else int(a60),
        patternFormed=int(pt['anyFormed']), patternBreakout=int(pt['anyBreakout']),
        patternJustBroke=int(pt['anyJustBroke']), pbAllPass=int(pb['allPass']),
    )
    for f in VP_LABELS:
        row[f] = 0
    if vps['sigType']:
        row[VP_SIGTYPE_FIELD[vps['sigType']]] = 1
    for r in pt['results']:
        row['ph_' + r['id']] = int(r['justBroke'])
    ex = extras or {}
    row['divTotal'] = ex.get('divTotal')
    row['priceYoy3m'] = ex.get('priceYoy3m')
    for k in ('revYoyTurn', 'revYoy3Pos', 'revHigh24'):
        row[k] = ex.get(k)
    row['inst3m'] = ex.get('inst3m')
    row['trust5'] = ex.get('trust5')
    row['foreign5'] = ex.get('foreign5')
    row['trustStreak3'] = ex.get('trustStreak3')
    row['confluenceCount'] = confluence_count(row)
    row.update((extras or {}).get('tech') or tech_extras(b, e))
    info = dict(dm=dm, sig=sig, pb=pb, pt=pt, prof=prof, vps=vps, rs=rs, vr=vr)
    return row, info


def confluence_count(r):
    c = 0
    if r.get('patternJustBroke') == 1:
        c += 1
    if r.get('kdjGoldenCrossRecent') == 1 or r.get('kdjDeathCrossRecent') == 1:
        c += 1
    if r.get('goldenCrossRecent') == 1:
        c += 1
    bp = r.get('bbPos')
    if isnum(bp) and (bp <= 20 or bp >= 80):
        c += 1
    vr = r.get('volRatio')
    if isnum(vr) and vr >= 1.5:
        c += 1
    return c


# ────────────────────────────────────────────────────────────────────
#  條件旗標定義（多因子複選搜尋／飆股搜尋／指定組合共用）
# ────────────────────────────────────────────────────────────────────
def _col(df, c):
    return df[c] if c in df.columns else pd.Series(np.nan, index=df.index)


# 新增的技術面條件旗標（欄位, 名稱, 判斷式）
NEW_TECH_FLAGS = [
    ('newHigh52', '創52週新高', lambda s: s.eq(1)),
    ('high52Dist', '距52週高點≤5%', lambda s: s.ge(-5)),
    ('maBull', '均線多頭排列(5>20>60且站上月線)', lambda s: s.eq(1)),
    # 布林通道收窄：2026-09-30 台股三年回測報酬低於基準，已移除（bbwRank 欄位仍保留）
    ('gapUp3', '近3日向上跳空缺口', lambda s: s.eq(1)),
    # 2026-10-03 參考 BBW+MACD 掃描器新增；2026-10-04 三年三段回測：突破季線／半年線、MACD零軸下金叉無效或反向，
    # 已移除旗標（crossMa60_3／crossMa100_3／gcBelow0_3 欄位仍保留在回測紀錄裡）
    ('vol3Ratio', '連續放量(近3日均量≥1.5倍前20日均量)', lambda s: s.ge(1.5)),
    # 2026-10-04 買賣力道近似（待回測驗證）
    ('udVolRatio20', '漲時量≥跌時量1.5倍(近20日)', lambda s: s.ge(1.5)),
    ('cmf20', 'CMF資金流買方佔優(近20日≥0.1)', lambda s: s.ge(0.1)),
    ('obvNewHigh60', 'OBV能量潮創60日新高', lambda s: s.eq(1)),
    # 2026-10-05 新增（待回測驗證）
    ('pullMa60', '多頭回測季線(季線上揚、距季線±3%)', lambda s: s.eq(1)),
    ('bias20', '短線過熱(高於月線≥20%)', lambda s: s.ge(20)),
    ('limitUp3', '近3日漲停', lambda s: s.eq(1)),
]
# 不放進「多因子複選搜尋／飆股搜尋」的條件（回測貢獻極低或與個別型態重複；指定組合比對與統計仍可使用）
SEARCH_EXCLUDE_FLAGS = {'型態成形中', '型態突破確認', '型態剛形成(剛突破)', '突破上升軌道線剛形成',
                        '分價量表-守穩POC買進', '分價量表-突破POC追價買進', '分價量表-反彈POC遇壓賣出', '分價量表-破位停損賣出'}
# 2026-10-05 台股三年三段回測稽核：下列條件單獨無效（或只在單一行情有效），且幾乎不出現在三段都穩定的選股／飆股組合，
# 不再放進搜尋以減少雜訊與運算量（指定組合比對、追蹤統計、單一條件表仍保留）
SEARCH_EXCLUDE_FLAGS |= {'跌深反彈盤', '頭肩底剛形成', '複式頭肩底剛形成', '一字底(均線糾結)剛形成', '三重底剛形成', '母子懷抱(高檔)剛形成', '布林通道低檔(≤20%)', '投信近5日買超', '投信連續買超≥3日', '土洋同買(外資、投信近5日皆買超)'}
# 2026-10-05 台股三年回測：多頭回測季線、月營收YoY由負轉正 三年都明顯負報酬，不再放進搜尋（單一條件表仍顯示，可當警示）
SEARCH_EXCLUDE_FLAGS |= {'多頭回測季線(季線上揚、距季線±3%)', '月營收YoY由負轉正'}


def search_flags(df):
    return {k: v for k, v in build_condition_flags(df).items() if k not in SEARCH_EXCLUDE_FLAGS}


def _period_masks(dates):
    """資料跨度 ≥18個月切三段（配合「三年分三次回測」合併後逐段驗證），否則切前後兩半；同一天不會被切開。
    回傳 ([(段名, 布林遮罩), ...], 說明文字)"""
    ud = np.unique(dates)
    if len(ud) < 4:
        return [], ''
    span = (dt.date.fromisoformat(str(ud[-1])[:10]) - dt.date.fromisoformat(str(ud[0])[:10])).days
    k = 3 if span >= 540 and len(ud) >= 6 else 2
    cuts = [ud[len(ud) * i // k] for i in range(1, k)]
    # 三段的命名和「三年分三段」回測一致：第1段＝最近1年、第3段＝最早（欄位由舊到新排列）
    labels = ['前半', '後半'] if k == 2 else ['第3段(最早)', '第2段', '第1段(最近)']
    parts = []
    for i in range(k):
        m = np.ones(len(dates), bool)
        if i > 0:
            m &= dates >= cuts[i - 1]
        if i < k - 1:
            m &= dates < cuts[i]
        parts.append((labels[i], m))
    desc = (f'前後半段以 {cuts[0]} 為界' if k == 2 else f'依評估日分三段（以 {cuts[0]}、{cuts[1]} 為界）')
    return parts, desc


def _t_of(ex):
    n = len(ex)
    if n < 5:
        return None
    sd = ex.std(ddof=1)
    return round(float(ex.mean() / (sd / math.sqrt(n))), 2) if sd > 0 else None


def build_condition_flags(df: pd.DataFrame):
    """回傳 {旗標名稱: bool Series}；欄位全部是空值的旗標不會出現（跟 HTML 版規則一致）"""
    F = {}

    def has(c):
        return c in df.columns and df[c].notna().any()

    def eq1(c):
        return _col(df, c).eq(1)

    sc = _col(df, 'score')
    F['多方力道≥65'] = sc.ge(65)
    F['多方力道≥80'] = sc.ge(80)
    F['強勢突破盤'] = eq1('isBreakout')
    F['跌深反彈盤'] = eq1('isPullbackRebound')
    bp = _col(df, 'bbPos')
    F['布林通道低檔(≤20%)'] = bp.le(20)
    F['布林通道高檔(≥80%)'] = bp.ge(80)
    F['MACD近3日內黃金交叉'] = eq1('goldenCrossRecent')
    if has('kdjGoldenCrossRecent'):
        F['KDJ近3日內黃金交叉'] = eq1('kdjGoldenCrossRecent')
    if has('kdjDeathCrossRecent'):
        F['KDJ近3日內死亡交叉'] = eq1('kdjDeathCrossRecent')
    if has('relStrength20'):
        F['相對強弱為正(強於大盤)'] = df['relStrength20'].gt(0)
        F['相對強弱為負(弱於大盤)'] = df['relStrength20'].lt(0)
        F['近20日強於大盤≥10%'] = df['relStrength20'].ge(10)
    if has('volRatio'):
        F['爆量(≥1.5倍均量)'] = df['volRatio'].ge(1.5)
        F['爆量(≥2倍均量)'] = df['volRatio'].ge(2)
        F['地量(≤0.5倍均量)'] = df['volRatio'].le(0.5)
    if has('volRangePos'):
        F['量能區間高檔(≥90百分位)'] = df['volRangePos'].ge(90)
        F['量能區間低檔(≤10百分位)'] = df['volRangePos'].le(10)
    if has('volTrend'):
        F['量能斜率轉強(近5日均量>近10日均量20%以上)'] = df['volTrend'].ge(20)
        F['量能斜率轉弱(近5日均量<近10日均量20%以上)'] = df['volTrend'].le(-20)
    if has('vpBuySupport'):
        for fld, lbl in VP_LABELS.items():
            F[lbl] = eq1(fld)
    if has('benchmarkAbove20'):
        F['大盤站上20日均線'] = eq1('benchmarkAbove20')
        F['大盤跌破20日均線'] = _col(df, 'benchmarkAbove20').eq(0)
    if has('benchmarkAbove60'):
        F['大盤站上60日均線'] = eq1('benchmarkAbove60')
        F['大盤跌破60日均線'] = _col(df, 'benchmarkAbove60').eq(0)
    if has('patternFormed'):
        F['型態成形中'] = eq1('patternFormed')
    if has('patternBreakout'):
        F['型態突破確認'] = eq1('patternBreakout')
    if has('patternJustBroke'):
        F['型態剛形成(剛突破)'] = eq1('patternJustBroke')
    if has('pbAllPass'):
        F['回後買上漲全通過'] = eq1('pbAllPass')
    for fld, lbl, fn in NEW_TECH_FLAGS:
        if has(fld):
            F[lbl] = fn(df[fld])
    if has('divTotal'):
        F['近3月乖離度為正(營收優於股價)'] = df['divTotal'].gt(0)
        F['近3月乖離度為負(股價超前營收)'] = df['divTotal'].lt(0)
    if has('priceYoy3m'):
        F['近3月均價YoY為正'] = df['priceYoy3m'].gt(0)
        F['近3月均價YoY為負'] = df['priceYoy3m'].lt(0)
    if has('inst3m'):
        F['三大法人近3月買超'] = df['inst3m'].gt(0)
        F['三大法人近3月賣超'] = df['inst3m'].lt(0)
    # 2026-10-05 新增：月營收動能（待回測驗證）
    if has('revYoyTurn'):
        F['月營收YoY由負轉正'] = eq1('revYoyTurn')
    if has('revYoy3Pos'):
        F['月營收連3月年增'] = eq1('revYoy3Pos')
    if has('revHigh24'):
        F['月營收創24個月新高'] = eq1('revHigh24')
    if has('trust5'):
        F['投信近5日買超'] = df['trust5'].gt(0)
    if has('trustStreak3'):
        F['投信連續買超≥3日'] = eq1('trustStreak3')
    if has('foreign5'):
        F['外資近5日買超'] = df['foreign5'].gt(0)
        if has('trust5'):
            F['土洋同買(外資、投信近5日皆買超)'] = df['foreign5'].gt(0) & df['trust5'].gt(0)
    if has('adx') and has('adxr') and has('plusDI'):
        # 2026-10-03 台股三年回測：飆股率6.0%（基準3.6%）、t(依日)6.1，但勝率略低於基準
        F['DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)'] = (_col(df, 'plusDI').gt(_col(df, 'minusDI')) & _col(df, 'adx').ge(25)
                                              & _col(df, 'adx').gt(_col(df, 'adxr')))
    if any(has(f) for f in PH_FIELDS):
        for pid, name in PATTERN_DEFS:
            F[name + '剛形成'] = eq1('ph_' + pid)
    return {k: v.fillna(False).astype(bool) for k, v in F.items()}


# ────────────────────────────────────────────────────────────────────
#  回測統計
# ────────────────────────────────────────────────────────────────────
HORIZONS = [5, 10, 20]


def stat_row(label_key, label_val, g: pd.DataFrame):
    row = {label_key: label_val, '樣本數': len(g)}
    for h in HORIZONS:
        v = g[f'ret{h}d'].dropna()
        row[f'{h}日平均報酬%'] = round(v.mean(), 2) if len(v) else None
        row[f'{h}日勝率%'] = round((v > 0).mean() * 100, 1) if len(v) else None
    return row


def score_buckets(df):
    def bk(s):
        return '80-100（積極做多）' if s >= 80 else '65-79（可考慮進場）' if s >= 65 else '50-64（觀望）' if s >= 50 else '0-49（不建議）'
    order = ['0-49（不建議）', '50-64（觀望）', '65-79（可考慮進場）', '80-100（積極做多）']
    b = df['score'].map(bk)
    return pd.DataFrame([stat_row('評分區間', k, df[b == k]) for k in order if (b == k).any()])


def tag_hitrate(df, field, label):
    if field not in df.columns or df[field].isna().all():
        return pd.DataFrame()
    rows = []
    for val, nm in ((0, '否'), (1, '是')):
        g = df[df[field] == val]
        if len(g):
            rows.append(stat_row('標記', f'{label}＝{nm}', g))
    return pd.DataFrame(rows)


def baseline_row(df):
    return pd.DataFrame([stat_row('標記', '全體基準', df)])


def new_tech_table(df):
    rows = []
    flg = build_condition_flags(df)
    for fld, lbl, _ in NEW_TECH_FLAGS:
        if lbl not in flg or fld not in df.columns:
            continue
        ok = df[fld].notna().values
        yes = flg[lbl].values & ok
        for nm, m in (('是', yes), ('否', ok & ~yes)):
            if m.sum():
                rows.append(stat_row('條件', f'{lbl}＝{nm}', df[m]))
    return pd.DataFrame(rows)


def pattern_hits(df):
    rows = []
    for pid, name in PATTERN_DEFS:
        c = 'ph_' + pid
        if c not in df.columns:
            continue
        g = df[df[c] == 1]
        if len(g) >= 10:
            rows.append(stat_row('型態', name + '剛形成', g))
    rows.sort(key=lambda r: -r['樣本數'])
    return pd.DataFrame(rows)


def confluence_table(df):
    rows = []
    for n in range(6):
        g = df[df['confluenceCount'] == n]
        if len(g):
            rows.append(stat_row('訊號堆疊數量', f'{n}個訊號同時觸發', g))
    return pd.DataFrame(rows)


def grid_search(df, h):
    col = f'ret{h}d'
    v = df.dropna(subset=[col, 'plusDI', 'minusDI', 'adx'])
    if not len(v):
        return pd.DataFrame()
    pdi, mdi, adx = v['plusDI'].values, v['minusDI'].values, v['adx'].values
    adxr = v['adxr'].values
    ret = v[col].values
    out = []
    for cap in (30, 40, 50):
        for bonus in (15, 20, 25):
            for w in (0.4, 0.5, 0.6):
                amax = max(0, 100 - 100 * w - bonus)
                s = pdi + mdi
                dom = np.where(s > 0, pdi / np.where(s > 0, s, 1) * 100, 50)
                sc = dom * w + np.minimum(adx, cap) / cap * amax + np.where(~np.isnan(adxr) & (adx > adxr), bonus, 0)
                sc = np.clip(sc, 0, 100)
                for thr in (50, 65, 80):
                    m = sc >= thr
                    if m.sum() < 20:
                        continue
                    r = ret[m]
                    out.append({'方向性權重': w, 'ADX滿分上限': cap, 'ADXR加分': bonus, '買進門檻': thr,
                                '訊號數': int(m.sum()), f'{h}日平均報酬%': round(r.mean(), 2),
                                f'{h}日勝率%': round((r > 0).mean() * 100, 1)})
    out.sort(key=lambda r: -r[f'{h}日平均報酬%'])
    return pd.DataFrame(out[:15])


def _flag_matrix(df, flags):
    names = list(flags.keys())
    M = np.column_stack([flags[n].values for n in names]) if names else np.zeros((len(df), 0), bool)
    return names, M


def combo_search(df, h, max_size=3, min_samples=50, min_winrate=50.0, skip_redundant=True, sort_by='t'):
    """多因子複選搜尋：窮舉1~3個條件的AND組合；向量化，50個旗標約數秒完成"""
    col = f'ret{h}d'
    valid = df[col].notna().values
    v = df[valid].reset_index(drop=True)
    flags = {kk: pd.Series(ff.values[valid]) for kk, ff in search_flags(df).items()}
    names, M = _flag_matrix(v, flags)
    k = len(names)
    if not len(v) or not k:
        return pd.DataFrame(), 0, {}
    ret = v[col].values.astype(float)
    pos = ret > 0
    base = dict(n=len(v), avg=ret.mean(), win=pos.mean() * 100, median=float(np.median(ret)))
    # 同日超額報酬：扣掉同一評估日全部紀錄的平均報酬（去除大盤齊漲齊跌的共同成分）
    if 'evalDate' in v.columns:
        exr = ret - v.groupby('evalDate')[col].transform('mean').values.astype(float)
    else:
        exr = ret - ret.mean()
    Mi = M.astype(np.int32)
    single = Mi.sum(0)
    pair = Mi.T @ Mi
    tested = 0
    hitsz = {1: [], 2: [], 3: []}   # (combo tuple, n, sum, winrate)，依條件數分開保持窮舉順序

    def consider(combo, n, s, w):
        if n < min_samples:
            return
        wr = w / n * 100
        if wr < min_winrate:
            return
        hitsz[len(combo)].append((combo, n, s, wr))

    # size 1
    for i in range(k):
        tested += 1
        n = single[i]
        if n:
            m = M[:, i]
            consider((i,), int(n), ret[m].sum(), int(pos[m].sum()))
    if max_size >= 2:
        for i in range(k):
            idx_i = np.flatnonzero(M[:, i])
            if not len(idx_i):
                tested += k - i - 1
                continue
            sub = M[idx_i]
            rs = ret[idx_i]
            ps = pos[idx_i].astype(np.int64)
            cnt = sub.sum(0)
            sm = rs @ sub
            wn = ps @ sub
            for j in range(i + 1, k):
                tested += 1
                n = int(cnt[j])
                if skip_redundant and (n == single[i] or n == single[j]):
                    continue
                consider((i, j), n, sm[j], int(wn[j]))
            if max_size >= 3:
                for j in range(i + 1, k):
                    nij = int(cnt[j])
                    if nij < min_samples:
                        tested += k - j - 1
                        continue
                    sel = sub[:, j]
                    sub2 = sub[sel]
                    rs2 = rs[sel]
                    ps2 = ps[sel]
                    c2 = sub2.sum(0)
                    s2 = rs2 @ sub2
                    w2 = ps2 @ sub2
                    for l in range(j + 1, k):
                        tested += 1
                        n = int(c2[l])
                        if skip_redundant and (n == nij or n == pair[i, l] or n == pair[j, l]):
                            continue
                        consider((i, j, l), n, s2[l], int(w2[l]))
    rows = []
    # 前後半段驗證：依評估日切兩半，各自算同日調整 t 值；兩段都 ≥2 才比較像真訊號、不是某段行情的巧合
    parts, pdesc = _period_masks(v['evalDate'].astype(str).values) if 'evalDate' in v.columns else ([], '')
    base['parts'] = pdesc
    for combo, n, s, wr in hitsz[1] + hitsz[2] + hitsz[3]:
        m = np.all(M[:, list(combo)], axis=1)
        r = ret[m]
        avg = r.mean()
        ex = exr[m]
        exm = ex.mean()
        sd = ex.std(ddof=1) if n > 1 else float('nan')
        t = exm / (sd / math.sqrt(n)) if sd and sd > 0 else float('nan')
        rows.append({'條件組合': ' ＋ '.join(names[i] for i in combo), '條件數': len(combo), '樣本數': n,
                     f'{h}日平均報酬%': round(avg, 2), f'{h}日勝率%': round(wr, 1),
                     '同日超額報酬%': round(exm, 2), '中位數報酬%': round(float(np.median(r)), 2),
                     't值(同日調整)': round(t, 2) if not math.isnan(t) else None})
        if parts:
            ts = [_t_of(exr[m & pm]) for _, pm in parts]
            ok = [x is not None and x >= 2 for x in ts]
            for (lbl, _), t_ in zip(parts, ts):
                rows[-1][f't值({lbl})'] = t_
            rows[-1]['各段一致'] = '✅' if all(ok) else (f'⚠️{sum(ok)}/{len(ok)}' if any(ok) else '')
    out = pd.DataFrame(rows)
    if len(out):
        if sort_by == 't':   # 依同日調整t值排序：優先列出「真的比同一天其他股票強」的選股型組合
            out = out.sort_values('t值(同日調整)', ascending=False, kind='mergesort', na_position='last').reset_index(drop=True)
        else:
            out = out.sort_values(f'{h}日勝率%', ascending=False, kind='mergesort').reset_index(drop=True)
    return out, tested, base


def moonshot_search(df, h, threshold=30.0, max_size=3, min_samples=20, min_pct=10.0, skip_redundant=True):
    col = f'ret{h}d'
    valid = df[col].notna().values
    v = df[valid].reset_index(drop=True)
    flags = {kk: pd.Series(ff.values[valid]) for kk, ff in search_flags(df).items()}
    names, M = _flag_matrix(v, flags)
    k = len(names)
    if not len(v) or not k:
        return pd.DataFrame(), 0, {}
    ret = v[col].values.astype(float)
    moon = ret > threshold
    base = dict(n=len(v), moonN=int(moon.sum()), pct=moon.mean() * 100,
                avg=ret[moon].mean() if moon.any() else float('nan'))
    Mi = M.astype(np.int32)
    single = Mi.sum(0)
    pair = Mi.T @ Mi
    tested = 0
    hitsz = {1: [], 2: [], 3: []}

    def consider(combo, n, mn):
        if n < min_samples or mn == 0:
            return
        if mn / n * 100 < min_pct:
            return
        hitsz[len(combo)].append((combo, n, mn))

    for i in range(k):
        tested += 1
        if single[i]:
            consider((i,), int(single[i]), int(moon[M[:, i]].sum()))
    if max_size >= 2:
        for i in range(k):
            idx_i = np.flatnonzero(M[:, i])
            if not len(idx_i):
                tested += k - i - 1
                continue
            sub = M[idx_i]
            ms = moon[idx_i].astype(np.int64)
            cnt = sub.sum(0)
            mc = ms @ sub
            for j in range(i + 1, k):
                tested += 1
                n = int(cnt[j])
                if skip_redundant and (n == single[i] or n == single[j]):
                    continue
                consider((i, j), n, int(mc[j]))
            if max_size >= 3:
                for j in range(i + 1, k):
                    nij = int(cnt[j])
                    if nij < min_samples:
                        tested += k - j - 1
                        continue
                    sel = sub[:, j]
                    sub2 = sub[sel]
                    ms2 = ms[sel]
                    c2 = sub2.sum(0)
                    m2 = ms2 @ sub2
                    for l in range(j + 1, k):
                        tested += 1
                        n = int(c2[l])
                        if skip_redundant and (n == nij or n == pair[i, l] or n == pair[j, l]):
                            continue
                        consider((i, j, l), n, int(m2[l]))
    rows = []
    parts, pdesc = _period_masks(v['evalDate'].astype(str).values) if 'evalDate' in v.columns else ([], '')
    base['parts'] = pdesc
    for combo, n, mn in hitsz[1] + hitsz[2] + hitsz[3]:
        m = np.all(M[:, list(combo)], axis=1)
        r = ret[m & moon]
        pct = mn / n * 100
        rows.append({'條件組合': ' ＋ '.join(names[i] for i in combo), '條件數': len(combo), '樣本數': n,
                     '飆股次數': mn, '飆股比例%': round(pct, 1), '飆股平均漲幅%': round(r.mean(), 1),
                     '倍數(vs基準)': round(pct / base['pct'], 2) if base['pct'] > 0 else None})
        if parts:
            for lbl, hm in parts:
                nn = int((m & hm).sum())
                rows[-1][f'飆股比例({lbl})%'] = round((m & hm & moon).sum() / nn * 100, 1) if nn else None
    out = pd.DataFrame(rows)
    if len(out):
        out = out.sort_values('飆股比例%', ascending=False, kind='mergesort').reset_index(drop=True)
    return out, tested, base


# 擇時型因子：這些條件本身描述的是「大盤／整體位置」，含這些條件的高勝率組合，勝率多半來自
# 「出現在大盤反彈的日子」（2026-09-27 台股回測：扣掉同日大盤後超額報酬≈0），
# 適合拿來判斷大盤是否落底，不適合當選股依據。
TIMING_FACTORS = {'布林通道低檔(≤20%)', '大盤跌破20日均線', '大盤跌破60日均線', '類股跌破20日均線', '類股跌破60日均線'}


def is_timing_combo(combo):
    return any(c in TIMING_FACTORS for c in combo)


def combo_type_label(combo):
    return '擇時型' if is_timing_combo(combo) else ''


def pinned_combo_stats(df, combos, winrates, h, stats_fmt=None, type_label=None):
    """指定組合在目前回測資料的表現；多了同日超額報酬、t值(同日調整)與「判讀」（選股有效／擇時／無效）"""
    col = f'ret{h}d'
    valid = df[col].notna().values
    v = df[valid]
    flags = {kk: pd.Series(ff.values[valid]) for kk, ff in build_condition_flags(df).items()}
    ret_all = v[col].values.astype(float)
    exr = ret_all - (v.groupby('evalDate')[col].transform('mean').values.astype(float)
                     if 'evalDate' in v.columns else ret_all.mean())
    out = []
    for combo in combos:
        key = ' ＋ '.join(combo)
        star = '⭐ ' if ('KDJ近3日內黃金交叉' in combo and '晨星剛形成' in combo) else ''
        rec_txt = stats_fmt(key) if stats_fmt else format_winrates(winrates.get(key))
        row = {'條件組合': star + key, '類型': type_label or combo_type_label(combo), '樣本數': 0, '歷史記錄': rec_txt,
               f'{h}日平均報酬%': None, f'{h}日勝率%': None, '同日超額報酬%': None, 't值(同日調整)': None,
               '判讀': '', '備註': ''}
        miss = [c for c in combo if c not in flags]
        if miss:
            row['備註'] = '缺少欄位（可能沒勾選對應的可選資料）：' + '、'.join(miss)
            out.append(row)
            continue
        m = np.logical_and.reduce([flags[c].values for c in combo])
        r = v[col].values[m]
        row['樣本數'] = int(m.sum())
        if not len(r):
            row['備註'] = '目前回測資料裡沒有符合這個組合的樣本'
        else:
            row[f'{h}日平均報酬%'] = round(r.mean(), 2)
            row[f'{h}日勝率%'] = round((r > 0).mean() * 100, 1)
            ex = exr[m]
            exm = ex.mean()
            sd = ex.std(ddof=1) if len(ex) > 1 else 0
            t = exm / (sd / math.sqrt(len(ex))) if sd > 0 else None
            row['同日超額報酬%'] = round(exm, 2)
            row['t值(同日調整)'] = round(t, 2) if t is not None else None
            if len(r) < 30:
                row['判讀'] = '樣本不足'
            elif t is not None and t >= 2:
                row['判讀'] = '✅ 選股有效'
            elif row[f'{h}日勝率%'] >= 60:
                row['判讀'] = '⏱ 擇時（勝率來自大盤）'
            else:
                row['判讀'] = '❌ 無效'
            row['備註'] = '' if len(r) >= 50 else '⚠️樣本數偏少，僅供參考'
        out.append(row)
    return pd.DataFrame(out)


def format_stockpick_stats(st_):
    if not st_:
        return ''
    return '　｜　'.join(f"{h}日：{e['n']}筆 勝率{e['win']}% 同日超額{e['ex']:+.2f}% t={e['t']}"
                        for h, e in sorted(st_.items(), key=lambda kv: int(kv[0])))


def format_winrates(wr):
    if not wr:
        return ''
    return ' / '.join(f"{h}日{wr[str(h)]}%" for h in (5, 10, 20) if str(h) in wr and wr[str(h)] is not None)


def format_moonshot_stats(st_):
    if not st_:
        return ''
    parts = []
    for h in (5, 10, 20):
        e = st_.get(str(h))
        if e:
            parts.append(f"{h}日：{e['n']}筆中{e['moonshotN']}次飆股({e['pct']}%)，平均漲幅+{e['avg']}%")
    return '　｜　'.join(parts)


# ════════════════════════════════════════════════════════════════════
#  FinMind 資料抓取
# ════════════════════════════════════════════════════════════════════
FINMIND_URLS = ['https://api.finmindtrade.com/api/v4/data', 'https://api.finmind.tw/api/latest/data']
_http = requests.Session()


class FinMindQuotaError(RuntimeError):
    """FinMind 每小時請求額度用完（HTTP 402／Requests reach the upper limit）"""


def api_get(params, token, timeout=30):
    """先打主站，失敗再試備援站；兩邊的錯誤都會列出來（以前只回報最後一個，主站的真正原因會被蓋掉）。
    額度用完時直接丟 FinMindQuotaError，不再去打備援站。"""
    errs = []
    for url in FINMIND_URLS:
        host = url.split('/')[2]
        try:
            r = _http.get(url, params={**params, 'token': token}, timeout=timeout)
        except Exception as ex:  # noqa
            errs.append(f'{host} 連線失敗（{type(ex).__name__}）')
            continue
        try:
            j = r.json()
        except Exception:  # noqa
            j = {}
        msg = str(j.get('msg') or '') if isinstance(j, dict) else ''
        if r.status_code == 402 or j.get('status') == 402 or 'upper limit' in msg.lower():
            raise FinMindQuotaError(f'FinMind 請求額度已用完（{msg or "HTTP 402"}）')
        if r.status_code != 200:
            errs.append(f'{host} HTTP {r.status_code}' + (f'：{msg}' if msg else ''))
            continue
        if j.get('status') == 200:
            return j
        errs.append(f'{host}：{msg or "status=" + str(j.get("status"))}')
    raise RuntimeError('；'.join(errs) or 'fetch failed')


def _ds(d):
    return d.isoformat()


def fetch_price(sid, token, days):
    end = dt.date.today()
    start = end - dt.timedelta(days=days)
    j = api_get(dict(dataset='TaiwanStockPrice', data_id=sid, start_date=_ds(start), end_date=_ds(end)), token)
    rows = j.get('data') or []
    if not rows:
        raise RuntimeError('無資料（可能代號錯誤）')
    # FinMind 在暫停交易日可能回傳收盤價 0，會讓報酬被算成 -100%，這種日子直接略過
    out = [dict(date=str(r['date'])[:10], open=float(r['open']), high=float(r['max']), low=float(r['min']),
                close=float(r['close']), volume=float(r['Trading_Volume'])) for r in rows
           if float(r.get('close') or 0) > 0 and float(r.get('max') or 0) > 0 and float(r.get('min') or 0) > 0]
    out.sort(key=lambda r: r['date'])
    return out


def fetch_benchmark(token, days, bid='0050'):
    rows = fetch_price(bid, token, days)
    return Benchmark([dict(date=r['date'], close=r['close']) for r in rows])


def fetch_name_map(token):
    try:
        j = api_get(dict(dataset='TaiwanStockInfo'), token)
        return {d['stock_id']: d.get('stock_name') or d['stock_id'] for d in j.get('data') or []}
    except Exception:  # noqa
        return {}


def fetch_pe_range(sid, token, years=3):
    end = dt.date.today()
    start = end.replace(year=end.year - years)
    j = api_get(dict(dataset='TaiwanStockPER', data_id=sid, start_date=_ds(start), end_date=_ds(end)), token)
    rows = sorted(j.get('data') or [], key=lambda r: r['date'])
    if not rows:
        return None
    vals = []
    for r in rows:
        try:
            v = float(r['PER'])
            if math.isfinite(v) and v > 0:
                vals.append(v)
        except Exception:  # noqa
            pass
    if not vals:
        return None
    try:
        cur = float(rows[-1]['PER'])
    except Exception:  # noqa
        cur = float('nan')
    if not math.isfinite(cur) or cur <= 0:
        cur = vals[-1]
    return dict(min=min(vals), max=max(vals), current=cur)


def _add_months(y, m, k):
    t = y * 12 + (m - 1) + k
    return t // 12, t % 12 + 1


def fetch_revenue_3m(sid, token):
    end = dt.date.today()
    sy, sm = _add_months(end.year, end.month, -16)
    start = dt.date(sy, sm, min(end.day, 28))
    j = api_get(dict(dataset='TaiwanStockMonthRevenue', data_id=sid, start_date=_ds(start), end_date=_ds(end)), token)
    mp = {}
    for r in j.get('data') or []:
        mp[(int(r['revenue_year']), int(r['revenue_month']))] = float(r['revenue'])
    if not mp:
        return None
    out = []
    for (y, m) in sorted(mp)[-3:]:
        rev = mp[(y, m)]
        py_, pm = _add_months(y, m, -1)
        prv = mp.get((py_, pm))
        ly = mp.get((y - 1, m))
        out.append(dict(year=y, month=m, revenue=rev,
                        mom=(rev - prv) / prv * 100 if prv else None,
                        yoy=(rev - ly) / ly * 100 if ly else None))
    return out


def fetch_inst_monthly(sid, token, with_flow=False):
    end = dt.date.today()
    start = end - dt.timedelta(days=100)
    j = api_get(dict(dataset='TaiwanStockInstitutionalInvestorsBuySell', data_id=sid,
                     start_date=_ds(start), end_date=_ds(end)), token)
    rows = j.get('data') or []
    if not rows:
        raise RuntimeError('回傳空陣列（可能此帳號方案未開通此資料集，或該股票近100天無法人交易紀錄）')
    mm = {}
    for r in rows:
        d = str(r.get('date', ''))[:10]
        if len(d) < 7:
            continue
        mm[d[:7]] = mm.get(d[:7], 0) + (float(r.get('buy') or 0) - float(r.get('sell') or 0))
    ks = sorted(mm)[-3:]
    if not ks:
        raise RuntimeError('解析不出日期/買賣超欄位')
    out = [dict(year=int(k[:4]), month=int(k[5:7]), net=mm[k]) for k in ks]
    if with_flow:
        # 投信／外資近5日、投信連3日買超：跟回測同一套算法（inst_flow_asof）
        daily, trust, foreign = {}, {}, {}
        for r in rows:
            d = str(r.get('date', ''))[:10]
            if len(d) < 10:
                continue
            net = float(r.get('buy') or 0) - float(r.get('sell') or 0)
            daily[d] = daily.get(d, 0) + net
            if _is_trust(r):
                trust[d] = trust.get(d, 0) + net
            elif _is_foreign(r):
                foreign[d] = foreign.get(d, 0) + net
        dates = sorted(daily)
        flow = inst_flow_asof(daily, trust, foreign, dates, len(dates) - 1) if dates else {}
        return out, flow
    return out


def last_n_calendar_months(n):
    t = dt.date.today()
    return [dict(year=_add_months(t.year, t.month, -k)[0], month=_add_months(t.year, t.month, -k)[1])
            for k in range(n - 1, -1, -1)]


def fetch_monthly_avg_price_yoy(sid, token, months=None):
    months = months or last_n_calendar_months(3)
    o = months[0]
    start = dt.date(o['year'] - 1, o['month'], 1)
    j = api_get(dict(dataset='TaiwanStockPrice', data_id=sid, start_date=_ds(start), end_date=_ds(dt.date.today())), token)
    sums, cnts = {}, {}
    for r in j.get('data') or []:
        k = str(r['date'])[:7]
        try:
            c = float(r['close'])
        except Exception:  # noqa
            continue
        if not math.isfinite(c) or c <= 0:
            continue
        sums[k] = sums.get(k, 0) + c
        cnts[k] = cnts.get(k, 0) + 1
    if not sums:
        return None
    avg = {k: sums[k] / cnts[k] for k in sums}
    out = []
    for m in months:
        k = f"{m['year']}-{m['month']:02d}"
        lk = f"{m['year'] - 1}-{m['month']:02d}"
        ta, la = avg.get(k), avg.get(lk)
        out.append(dict(year=m['year'], month=m['month'], avg=ta, avgLastYear=la,
                        yoy=(ta - la) / la * 100 if (ta is not None and la) else None))
    return out


def fetch_realtime(token, ids, chunk=80):
    mp, err = {}, None
    for i in range(0, len(ids), chunk):
        params = [('data_id', s) for s in ids[i:i + chunk]] + [('token', token)]
        try:
            r = _http.get('https://api.finmindtrade.com/api/v4/taiwan_stock_tick_snapshot', params=params, timeout=30)
            if r.status_code != 200:
                raise RuntimeError(f'HTTP {r.status_code}：{r.text[:200]}')
            j = r.json()
            if j.get('status') != 200:
                raise RuntimeError(j.get('msg') or f"status={j.get('status')}")
            for row in j.get('data') or []:
                if row and row.get('stock_id'):
                    mp[row['stock_id']] = row
        except Exception as ex:  # noqa
            err = str(ex)
    return mp, (None if mp else err)


def merge_realtime(rows, snap):
    if not snap or not snap.get('date'):
        return rows, False
    today = str(snap['date'])[:10]
    if rows and today <= rows[-1]['date']:
        return rows, False
    try:
        c = float(snap['close'])
    except Exception:  # noqa
        return rows, False
    if not c or not math.isfinite(c):
        return rows, False

    def g(k):
        try:
            v = float(snap.get(k))
            return v if v else c
        except Exception:  # noqa
            return c
    vol = snap.get('total_volume', snap.get('volume'))
    try:
        vol = float(vol)
    except Exception:  # noqa
        vol = 0
    return rows + [dict(date=today, open=g('open'), high=g('high'), low=g('low'), close=c,
                        volume=vol if math.isfinite(vol) else 0)], True


def fetch_market_snapshot(token, twse, tpex, log=None):
    """全市場當日行情（需 Backer/Sponsor）：往前找最近7天內有資料的交易日"""
    info = api_get(dict(dataset='TaiwanStockInfo'), token)
    names = {d['stock_id']: d['stock_name'] for d in info.get('data') or []}
    valid = set(twse) | set(tpex)
    last_err = ''
    for back in range(7):
        ds = _ds(dt.date.today() - dt.timedelta(days=back))
        if log:
            log(f'📡 查詢 {ds} 全市場行情中...')
        try:
            j = api_get(dict(dataset='TaiwanStockPrice', start_date=ds), token, timeout=60)
            data = j.get('data') or []
            if not data:
                last_err = 'no data'
                continue
            res = []
            for r in data:
                sid = r.get('stock_id')
                if sid not in valid:
                    continue
                try:
                    c, sp = float(r['close']), float(r['spread'])
                except Exception:  # noqa
                    continue
                if c <= 0 or c - sp <= 0:
                    continue
                v = float(r.get('Trading_Volume') or 0)
                res.append(dict(id=sid, name=names.get(sid, sid), pct=sp / (c - sp) * 100, volume=v))
            if res:
                return res, ds
        except Exception as ex:  # noqa
            last_err = str(ex)
    raise RuntimeError('近7天皆無法取得全市場資料：' + last_err + '（此功能需 FinMind Backer/Sponsor 方案）')


# ── 回測用：長歷史月營收／法人，與「當時已公告」判斷 ──
def fetch_revenue_hist(sid, token, days_back):
    end = dt.date.today()
    j = api_get(dict(dataset='TaiwanStockMonthRevenue', data_id=sid,
                     start_date=_ds(end - dt.timedelta(days=days_back)), end_date=_ds(end)), token)
    mp = {}
    for r in j.get('data') or []:
        if r.get('revenue_year') is None or r.get('revenue_month') is None or r.get('revenue') is None:
            continue
        mp[(int(r['revenue_year']), int(r['revenue_month']))] = float(r['revenue'])
    return mp


def fetch_inst_hist(sid, token, days_back):
    end = dt.date.today()
    j = api_get(dict(dataset='TaiwanStockInstitutionalInvestorsBuySell', data_id=sid,
                     start_date=_ds(end - dt.timedelta(days=days_back)), end_date=_ds(end)), token)
    daily, trust, foreign = {}, {}, {}
    for r in j.get('data') or []:
        d = str(r.get('date', ''))[:10]
        if len(d) < 10:
            continue
        net = float(r.get('buy') or 0) - float(r.get('sell') or 0)
        daily[d] = daily.get(d, 0) + net
        if _is_trust(r):
            trust[d] = trust.get(d, 0) + net
        elif _is_foreign(r):
            foreign[d] = foreign.get(d, 0) + net
    return daily, trust, foreign


def _is_trust(r):
    nm = str(r.get('name') or '')
    return 'Investment_Trust' in nm or '投信' in nm


def _is_foreign(r):
    """外資＝外資及陸資（Foreign_Investor）＋外資自營商（Foreign_Dealer_Self）"""
    nm = str(r.get('name') or '')
    return 'Foreign' in nm or '外資' in nm


def inst_flow_asof(daily, trust, foreign, dates, e):
    """截至第 e 根K棒（含）：投信／外資近5個交易日買賣超合計、投信是否連續3日買超。
    這5天都沒有法人資料時全部回傳 None"""
    ds = dates[max(0, e - 4):e + 1]
    if not any(d in daily for d in ds):
        return dict(trust5=None, foreign5=None, trustStreak3=None)
    d3 = dates[max(0, e - 2):e + 1]
    return dict(trust5=sum(trust.get(d, 0) for d in ds), foreign5=sum(foreign.get(d, 0) for d in ds),
                trustStreak3=int(len(d3) == 3 and all(trust.get(d, 0) > 0 for d in d3)))


def revenue_known_by(y, m, asof):
    dy, dm = (y, m + 1) if m < 12 else (y + 1, 1)
    return dt.date.fromisoformat(asof) >= dt.date(dy, dm, 10)


def last_n_known_months(asof, n):
    d = dt.date.fromisoformat(asof)
    y, m = d.year, d.month
    out = []
    for _ in range(18):
        if revenue_known_by(y, m, asof):
            out.append((y, m))
            if len(out) >= n:
                break
        m -= 1
        if m < 1:
            m, y = 12, y - 1
    return out[::-1]


def rev_feats(rev, asof):
    """月營收動能（只用 asof 當下已公告的月份，次月10日才算已知）：
    revYoyTurn＝最新月YoY由負轉正；revYoy3Pos＝近3個月YoY皆為正；revHigh24＝最新月營收創24個月新高"""
    out = dict(revYoyTurn=None, revYoy3Pos=None, revHigh24=None)
    months = last_n_known_months(asof, 4)
    if not rev or not months:
        return out

    def yoy(y, m):
        rt, rl = rev.get((y, m)), rev.get((y - 1, m))
        return (rt - rl) / rl * 100 if (rt is not None and rl) else None
    ys = [yoy(y, m) for y, m in months]
    if len(ys) >= 2 and ys[-1] is not None and ys[-2] is not None:
        out['revYoyTurn'] = int(ys[-1] > 0 and ys[-2] <= 0)
    if len(ys) >= 3 and all(v is not None for v in ys[-3:]):
        out['revYoy3Pos'] = int(all(v > 0 for v in ys[-3:]))
    ly, lm = months[-1]
    cur = rev.get((ly, lm))
    if cur is not None:
        prev = [rev.get(_add_months(ly, lm, -k)) for k in range(1, 24)]
        prev = [p for p in prev if p is not None]
        if len(prev) >= 18:
            out['revHigh24'] = int(cur > max(prev))
    return out


def price_month_avg_map(b: Bars):
    sums, cnts = {}, {}
    for d, c in zip(b.date, b.close):
        k = (int(d[:4]), int(d[5:7]))
        sums[k] = sums.get(k, 0) + c
        cnts[k] = cnts.get(k, 0) + 1
    return {k: sums[k] / cnts[k] for k in sums}


def divergence_asof(rev, pxm, asof):
    months = last_n_known_months(asof, 3)
    if not months:
        return None, None
    div = pyoy = 0.0
    anyd = anyp = False
    for (y, m) in months:
        rt, rl = rev.get((y, m)), rev.get((y - 1, m))
        pt_, pl = pxm.get((y, m)), pxm.get((y - 1, m))
        if pl and pt_ is not None:
            pyoy += (pt_ - pl) / pl * 100
            anyp = True
        if not rl or rt is None or not pl or pt_ is None:
            continue
        div += (rt - rl) / rl * 100 - (pt_ - pl) / pl * 100
        anyd = True
    return (div if anyd else None), (pyoy if anyp else None)


def inst3m_asof(daily, asof):
    end = dt.date.fromisoformat(asof)
    start = end - dt.timedelta(days=90)
    tot, anyv = 0.0, False
    for ds, v in daily.items():
        d = dt.date.fromisoformat(ds)
        if start <= d <= end:
            tot += v
            anyv = True
    return tot if anyv else None


# ════════════════════════════════════════════════════════════════════
#  歷史回測
# ════════════════════════════════════════════════════════════════════
# ── 回測記憶體：每檔結果轉成「欄位→numpy 陣列」，最後逐欄合併（合併完一欄就釋放該欄），
# 記憶體高峰約只有最終資料的 1 倍，不會像 pd.concat＋排序＋複製那樣疊到 3 倍 ──
DEAD_BT_COLS = ('crossMa60_3', 'crossMa100_3', 'gcBelow0_3')   # 已確認無效、旗標已移除的欄位，不再存進回測紀錄


def frame_arrays(d):
    return {c: d[c].to_numpy(copy=True) for c in d.columns if c not in DEAD_BT_COLS}


def low_mem_concat(frames, sort=True):
    """frames：list of {欄位: 陣列}；依 evalDate、stockId 排序後組成 DataFrame（數值欄一律 float32）"""
    frames = [f for f in frames if f and len(next(iter(f.values()))) > 0]
    if not frames:
        return pd.DataFrame()
    cols = []
    seen = set()
    for f in frames:
        for c in f:
            if c not in seen:
                seen.add(c)
                cols.append(c)
    front = [c for c in ('evalDate', 'stockId', 'name') if c in seen]
    cols = front + [c for c in cols if c not in front]
    order = None
    if sort and 'evalDate' in seen and 'stockId' in seen:
        ed = np.concatenate([np.asarray(f['evalDate']).astype(str) for f in frames])
        sid = np.concatenate([np.asarray(f['stockId']).astype(str) for f in frames])
        order = np.lexsort((sid, ed))
        del ed, sid
    lens = [len(next(iter(f.values()))) for f in frames]
    out = {}
    for c in cols:
        parts = []
        for f, n in zip(frames, lens):
            v = f.pop(c, None)
            if v is None:
                v = np.full(n, np.nan, dtype=np.float32) if c not in front else np.array([''] * n, dtype=object)
            parts.append(v)
        arr = np.concatenate(parts)
        del parts
        if order is not None:
            arr = arr[order]
        if c not in front:
            arr = pd.to_numeric(pd.Series(arr), errors='coerce').to_numpy(dtype=np.float32)
        out[c] = arr
    frames.clear()
    return pd.DataFrame(out, copy=False)


def add_derived(df, inplace=False):
    if df.attrs.get('derived'):
        return df
    if not inplace:
        df = df.copy()

    def b01(s, cond):
        return np.where(s.isna(), np.nan, cond.astype(float))
    if 'relStrength20' in df:
        df['relStrengthPositive'] = b01(df['relStrength20'], df['relStrength20'] > 0)
    if 'volRatio' in df:
        df['volSurge15'] = b01(df['volRatio'], df['volRatio'] >= 1.5)
        df['volSurge2'] = b01(df['volRatio'], df['volRatio'] >= 2)
        df['volLow05'] = b01(df['volRatio'], df['volRatio'] <= 0.5)
    if 'volRangePos' in df:
        df['volRangeHigh90'] = b01(df['volRangePos'], df['volRangePos'] >= 90)
        df['volRangeLow10'] = b01(df['volRangePos'], df['volRangePos'] <= 10)
    return df



def store_bt_df(df):
    """回測資料存進 session 前先整理一次：衍生欄位只算一次、數值轉 float32、代號／名稱轉 category，
    全美股約20萬筆時可省下大半記憶體（Streamlit Cloud 記憶體上限約 1GB，超過會整個 App 當掉）"""
    if df is None or not len(df):
        return df
    df = add_derived(df, inplace=True)   # 剛跑完／剛載入的資料不再複製一份，省一倍記憶體高峰
    for c in [c for c in DEAD_BT_COLS if c in df.columns]:
        del df[c]
    for c in df.columns:
        if c in ('evalDate', 'stockId', 'name'):
            continue
        if (pd.api.types.is_float_dtype(df[c]) or pd.api.types.is_integer_dtype(df[c])) and df[c].dtype != np.float32:
            df[c] = df[c].astype(np.float32)
    for c in ('stockId', 'name'):
        if c in df.columns and not isinstance(df[c].dtype, pd.CategoricalDtype):
            df[c] = df[c].astype(str).astype('category')
    df['evalDate'] = df['evalDate'].astype(str)
    df.attrs['derived'] = True
    return df


def release_bt_state(ss):
    """清掉上一份回測資料與衍生的快取／壓縮檔，釋放記憶體"""
    for k in ('bt_df', '_bt_memo', '_bt_csv', '_bt_csv_sig', 'live_stats', 'live_stats_sig'):
        ss.pop(k, None)
    gc.collect()


def bt_memo_ready(ss, name, params):
    """這組參數是否已經算過（不觸發計算）"""
    df = ss.get('bt_df')
    sig = (ss.get('bt_ver'), id(df), len(df) if df is not None else 0)
    memo = ss.get('_bt_memo')
    return bool(memo) and memo.get('_sig') == sig and ((name,) + tuple(params)) in memo


def bt_memo(ss, name, params, fn):
    """回測分析結果快取（存在 session）：同一份回測資料、同樣參數只算一次，
    改任何一個選單時 Streamlit 會整頁重跑，沒有快取的話每次都要重算幾十秒。"""
    df = ss.get('bt_df')
    sig = (ss.get('bt_ver'), id(df), len(df) if df is not None else 0)
    memo = ss.get('_bt_memo')
    if memo is None or memo.get('_sig') != sig:
        memo = {'_sig': sig}
        ss['_bt_memo'] = memo
    key = (name,) + tuple(params)
    if key not in memo:
        memo[key] = fn()
    return memo[key]


def bt_csv_gz(df):
    import gzip
    buf = io.BytesIO()
    with gzip.GzipFile(fileobj=buf, mode='wb', compresslevel=5) as gz:
        with io.TextIOWrapper(gz, encoding='utf-8-sig', newline='') as tw:
            df.to_csv(tw, index=False, float_format='%.9g')   # float32 需要9位有效數字才能原值還原
    return buf.getvalue()


def load_bt_files(files):
    """一次載入一個或多個回測原始紀錄檔（例如三年分三段各一個），合併後去除重複（同日同檔）"""
    parts = []
    for f in files:
        if hasattr(f, 'seek'):
            f.seek(0)
        d = pd.read_csv(f, compression='gzip' if f.name.endswith('.gz') else None, dtype={'stockId': str})
        for c in d.columns:   # 先逐檔轉 float32，合併時記憶體高峰較低
            if c not in ('evalDate', 'stockId', 'name') and (pd.api.types.is_float_dtype(d[c]) or pd.api.types.is_integer_dtype(d[c])):
                d[c] = d[c].astype(np.float32)
        d['evalDate'] = d['evalDate'].astype(str)
        parts.append(d)
    df = pd.concat(parts, ignore_index=True) if len(parts) > 1 else parts[0]
    del parts
    if len(files) > 1:
        df = df.drop_duplicates(['evalDate', 'stockId'], keep='last').sort_values(['evalDate', 'stockId']).reset_index(drop=True)
    return store_bt_df(df)


BT_SEG_OPTIONS = ['自訂月數', '三年分三段・第1段（最近1年）', '三年分三段・第2段（1～2年前）', '三年分三段・第3段（2～3年前）']

def run_backtest(token, universe, months_back, include_div, include_inst, min_liq, vp_params,
                 delay=0.3, on_progress=None, should_stop=None, end_offset_months=0, max_quota_wait=65 * 60):
    """FinMind 每小時額度用完時（HTTP 402）不會把剩下的股票全部記成失敗，而是暫停等待、額度恢復後從同一檔繼續；
    最多等 max_quota_wait 秒。前 40 檔全部因同一個原因失敗（Token 錯誤、連線不到）會提早停止。"""
    off = int(end_offset_months or 0)   # 分段回測：評估區間往前推 off 個月（例：第2段＝1～2年前）
    buf = 60
    maxh = max(HORIZONS)
    price_days = (months_back + off) * 30 + buf + maxh + 10 + 400   # +400天：52週高點、布林收窄需要約一年歷史（營收乖離也夠用）
    rev_days = (months_back + off) * 30 + 800   # 營收創24個月新高需要多兩年的月營收
    inst_days = (months_back + off) * 30 + 100
    win_start = _ds(dt.date.today() - dt.timedelta(days=(months_back + off) * 30))
    win_end = _ds(dt.date.today() - dt.timedelta(days=off * 30)) if off else None
    fail_reasons = {}
    frames = []   # 每檔算完就轉成 float32 DataFrame（整段回測若用 list of dict 暫存，十幾萬筆會吃掉近 1GB 記憶體）
    saved = failed = 0
    total = len(universe)
    stopped = None
    state = {'waited': 0.0}

    def prog(i, sid, note=None):
        if on_progress:
            try:
                on_progress(i, total, saved, failed, sid, note)
            except TypeError:   # 舊版 callback 沒有 note 參數
                on_progress(i, total, saved, failed, sid)

    def with_quota_wait(fn, i, sid):
        """執行 fn；碰到額度用完就每 30 秒重試一次，累計最多等 max_quota_wait 秒"""
        while True:
            try:
                return fn()
            except FinMindQuotaError as qe:
                if state['waited'] >= max_quota_wait:
                    raise
                mins = int(state['waited'] // 60)
                prog(i, sid, f'⏸ {qe}，暫停等待額度恢復（已等 {mins} 分鐘，最多 {max_quota_wait // 60} 分鐘），恢復後會自動從 {sid or "開頭"} 繼續')
                for _ in range(6):
                    if should_stop and should_stop():
                        raise
                    time.sleep(5)
                state['waited'] += 30

    try:
        name_map = with_quota_wait(lambda: fetch_name_map(token), 0, None)
        bm = with_quota_wait(lambda: fetch_benchmark(token, price_days), 0, None)
    except FinMindQuotaError as qe:
        return pd.DataFrame(), dict(saved=0, failed=0, total=total, top_fail=[str(qe)],
                                    stopped=f'{qe}，等了 {max_quota_wait // 60} 分鐘仍未恢復')
    except Exception:  # noqa
        name_map = name_map if 'name_map' in locals() else {}
        bm = None

    def one(sid):
        rows = fetch_price(sid, token, price_days)
        b = Bars(rows)
        n = b.n
        ev_end = n - maxh
        if ev_end <= buf:
            return None, '股價資料天數不足'
        rev, pxm, inst, trust, foreign = {}, {}, {}, {}, {}
        if include_div:
            try:
                rev = fetch_revenue_hist(sid, token, rev_days)
                pxm = price_month_avg_map(b)
            except FinMindQuotaError:
                raise
            except Exception:  # noqa
                rev = {}
        if include_inst:
            try:
                inst, trust, foreign = fetch_inst_hist(sid, token, inst_days)
            except FinMindQuotaError:
                raise
            except Exception:  # noqa
                inst, trust, foreign = {}, {}, {}
        pc = PatternCache(b)
        recs = []
        for e in range(buf, ev_end):
            ds = b.date[e]
            if ds < win_start or (win_end and ds >= win_end):
                continue
            if min_liq > 0:
                s0 = max(0, e - 19)
                liq = sum(b.close[k] * b.volume[k] for k in range(s0, e + 1)) / (e + 1 - s0)
                if liq < min_liq:
                    continue
            ex = {}
            if include_div:
                ex['divTotal'], ex['priceYoy3m'] = divergence_asof(rev, pxm, ds)
                ex.update(rev_feats(rev, ds))
            if include_inst:
                ex['inst3m'] = inst3m_asof(inst, ds)
                ex.update(inst_flow_asof(inst, trust, foreign, b.date, e))
            row, _ = build_flag_row(b, e, bm, pc, vp_params, ex)
            ec = b.close[e]
            row['entryClose'] = ec
            for h in HORIZONS:
                row[f'ret{h}d'] = (b.close[e + h] - ec) / ec * 100 if (e + h < n and ec) else None
            recs.append(row)
        if not recs:
            return None, None   # 沒有落在評估區間／流動性門檻內，不算失敗
        d = pd.DataFrame(recs)
        del recs
        for c in d.columns:
            if c != 'evalDate':
                d[c] = pd.to_numeric(d[c], errors='coerce').astype('float32')
        d.insert(1, 'stockId', sid)
        d.insert(2, 'name', name_map.get(sid, sid))
        return d, None

    for i, sid in enumerate(universe):
        if should_stop and should_stop():
            break
        prog(i, sid)
        try:
            d, reason = with_quota_wait(lambda: one(sid), i, sid)
        except FinMindQuotaError as qe:
            stopped = f'{qe}，等了 {max_quota_wait // 60} 分鐘仍未恢復，停在第 {i + 1} 檔（{sid}）。已完成的部分仍可下載，但這一段不完整，建議額度恢復後整段重跑'
            break
        except Exception as ex:  # noqa
            d, reason = None, (str(ex)[:120] or type(ex).__name__)
        if d is not None:
            saved += len(d)
            frames.append(frame_arrays(d))
            del d
        elif reason:
            failed += 1
            fail_reasons[reason] = fail_reasons.get(reason, 0) + 1
        if i + 1 >= 40 and saved == 0 and failed >= i + 1 and len(fail_reasons) == 1:
            stopped = f'前 {i + 1} 檔全部失敗且原因相同，提早停止：{next(iter(fail_reasons))}'
            break
        if i < total - 1 and delay:
            time.sleep(delay)
    prog(total, None)
    df = low_mem_concat(frames, sort=False)
    gc.collect()
    if len(df):
        df['evalDate'] = df['evalDate'].astype(str)
    top = sorted(fail_reasons.items(), key=lambda kv: -kv[1])[:3]
    return df, dict(saved=saved, failed=failed, total=total, top_fail=[f'{k}（{v}檔）' for k, v in top],
                    stopped=stopped, quota_wait_min=int(state['waited'] // 60))


# ════════════════════════════════════════════════════════════════════
#  批次分析（單檔）
# ════════════════════════════════════════════════════════════════════
def analyze_stock(sid, name, rows, bm, vp_params, extras_data):
    b = Bars(rows)
    e = b.n - 1
    ex = {'tech': extras_data.get('tech')}
    rv, py_ = extras_data.get('revRange'), extras_data.get('priceYoYRange')
    if rv and py_:
        pxk = {(p['year'], p['month']): p for p in py_}
        s, anyv = 0.0, False
        for r in rv:
            px = pxk.get((r['year'], r['month']))
            if r['yoy'] is not None and px and px['yoy'] is not None:
                s += r['yoy'] - px['yoy']
                anyv = True
        ex['divTotal'] = s if anyv else None
    if py_:
        vals = [p['yoy'] for p in py_ if p['yoy'] is not None]
        ex['priceYoy3m'] = sum(vals) if vals else None
    ir = extras_data.get('instRange')
    if ir:
        ex['inst3m'] = sum(m['net'] for m in ir)
    ex.update(extras_data.get('instFlow') or {})
    ex.update(extras_data.get('revFeats') or {})
    pc = PatternCache(b)
    row, info = build_flag_row(b, e, bm, pc, vp_params, ex)
    return dict(stockId=sid, name=name, bars=b, row=row, info=info, total=info['dm']['score'], **extras_data)


def matched_combos(row, combos):
    df = pd.DataFrame([row])
    flags = build_condition_flags(df)
    out = []
    for idx, combo in enumerate(combos):
        if all(c in flags and bool(flags[c].iloc[0]) for c in combo):
            out.append(idx)
    return out


def score_label(t):
    return '積極做多' if t >= 80 else '可考慮進場' if t >= 65 else '觀望' if t >= 50 else '不建議進場'


# ── 組合型態：依命中類別（S/#/M/W/F）的組合分級，依據三年回測（20日報酬，含 t(依日) 檢驗，2026-10-04 三段回測重新統計） ──
COMBO_STYLE_RULES = {'S#MWF': '進攻', 'SMWF': '進攻', 'SWF': '進攻', 'SF': '進攻', 'SM': '進攻', 'WF': '進攻', 'S': '穩健', 'SW': '穩健', '#W': '穩健', 'S#WF': '穩健', '#WF': '穩健', '#F': '穩健', 'F': '彩券', 'MF': '彩券', 'M': '彩券', 'SMF': '彩券', '#MF': '彩券', '#M': '彩券'}
COMBO_STYLE_ICON = {'進攻': '🚀進攻', '穩健': '🛡️穩健', '彩券': '🎲彩券'}
COMBO_STYLE_NOTE = {'進攻': '台股三年三段回測（2023-10～2026-09，2026-10-05 重算）：20日超額+3.3~7.8%、飆股率約2.4~5.2倍、t(依日)3.1~5.9；S#MWF最強（勝率60%、超額+7.8%），波動大（跌>10%約15~27%）',
                    '穩健': '台股三年三段回測：20日勝率51~85%、跌>10%約5~14%（全體13.5%）；S單獨命中最穩（超額+2.7%、t(依日)8.3），#W勝率66%',
                    '彩券': '台股三年三段回測：20日勝率42~51%、跌>10%約17~29%，飆股率約1.6~3.3倍但期望值小（t(依日)1.2~3.6）；小部位分散＋停損'}


def combo_style(ms, mp, mm, mw, mf):
    key = ''.join(k for k, x in (('S', ms), ('#', mp), ('M', mm), ('W', mw), ('F', mf)) if x)
    return COMBO_STYLE_ICON.get(COMBO_STYLE_RULES.get(key, ''), '')


# ── 大盤狀態：先判斷大盤，再看命中的組合在「這種大盤」下回測是否有效 ──
MARKET_REGIMES = {1: ('①多頭', '加權指數站上月線（MA20）與季線（MA60）'),
                  2: ('②多頭回檔', '加權指數跌破月線、仍守住季線'),
                  3: ('③空頭反彈', '加權指數站上月線、但仍在季線下'),
                  4: ('④空頭', '加權指數跌破月線與季線')}
REGIME_ADVICE = {1: '各類組合都有效：S選股型、W高勝率、F飆股、近20日強於大盤、距52週高點、營收創新高',
                 2: 'W高勝率與#清單失效；S選股型、月營收連3月年增、外資買超仍有效，宜挑營收題材',
                 3: '#清單明顯為負（避開）；S選股型、F飆股、營收創新高、近20日強勢股有效',
                 4: '追強勢（近20日強於大盤、距52週高點）失效、投信買超為負；S選股型、月營收連3月年增、外資買超仍有效'}
REGIME_MIN_N = 50   # 該大盤狀態下回測樣本少於此數 → 不標示


def market_regime(row):
    """依 benchmarkAbove20／benchmarkAbove60 判斷大盤狀態 1～4；沒有大盤資料回傳 None"""
    a20, a60 = (row or {}).get('benchmarkAbove20'), (row or {}).get('benchmarkAbove60')
    if a20 is None or a60 is None:
        return None
    try:
        a20, a60 = int(a20) == 1, int(a60) == 1
    except Exception:  # noqa
        return None
    return 1 if (a20 and a60) else (2 if a60 else (3 if a20 else 4))


def regime_tag(combo, reg, K):
    """✅ 該大盤狀態下 t(依日)≥2 且超額為正／❌ 超額為負／⚠️ 其他；樣本不足回傳 ''"""
    if not reg:
        return ''
    row = (K.get('REGIME_STATS') or {}).get(combo_norm_key(combo))
    s = row[reg - 1] if row and len(row) >= reg else None
    if not s or s[0] < REGIME_MIN_N:
        return ''
    if s[3] is not None and s[3] >= 2 and s[2] > 0:
        return '✅'
    return '❌' if s[2] < 0 else '⚠️'


def regime_fit(groups, reg, K, show=4):
    """groups：[(代號前綴, 組合清單, 命中索引), ...] → ('✅2 S1 W3｜❌1 F12', ✅數, ❌數)"""
    if not reg:
        return '', 0, 0
    buckets = {'✅': [], '⚠️': [], '❌': []}
    for code, lst, idxs in groups:
        for i in idxs:
            t = regime_tag(lst[i], reg, K)
            if t:
                buckets[t].append(f'{code}{i + 1}')
    parts = []
    for t in ('✅', '❌', '⚠️'):
        v = buckets[t]
        if v:
            parts.append(f"{t}{len(v)} " + ' '.join(v[:show]) + (f' +{len(v) - show}' if len(v) > show else ''))
    return '｜'.join(parts), len(buckets['✅']), len(buckets['❌'])


def batch_market_regime(results):
    for r in results or []:
        reg = market_regime(r.get('row'))
        if reg:
            return reg
    return None


def regime_banner(reg):
    nm, desc = MARKET_REGIMES[reg]
    return f"🧭 今日大盤狀態：**{nm}**（{desc}）　→　{REGIME_ADVICE[reg]}。摘要表「大盤適用」欄：✅＝這種大盤下回測有效、❌＝回測為負、⚠️＝普通。"


def summary_frame(results, extras_flags, K, live=None):
    recs = []
    for r in results:
        b, info = r['bars'], r['info']
        e = b.n - 1
        last_c = b.close[e]
        prev_c = b.close[e - 1] if e >= 1 else last_c
        chg = (last_c - prev_c) / prev_c * 100 if prev_c else 0
        dm, sig, pb, pt = info['dm'], info['sig'], info['pb'], info['pt']
        bbw = (b.bbU[e] - b.bbL[e]) / last_c * 100 if (b.bbU[e] is not None and last_c) else None
        combo_tag = ''
        if sig['isBreakout']:
            combo_tag = '🚀 強勢突破盤'
        elif sig['isPullbackRebound']:
            combo_tag = '🎯 跌深反彈盤'
        kd = info['sig']['kdjState'] or ''
        vps, prof = info['vps'], info['prof']
        vp_txt = ''
        if prof:
            vp_txt = f"POC {prof['pocPrice']:.1f}（{prof['pocLow']:.1f}~{prof['pocHigh']:.1f}）"
        vp_sig = ''
        if vps['sigType']:
            vp_sig = ('🟢 ' if vps['signal'] == 'buy' else '🔴 ') + vps['detail'].split('：')[0]
        pnames = [('🔥' if x['justBroke'] else '') + x['name'] for x in pt['results']
                  if (x['breakout'] if pt['anyBreakout'] else x['formed'])]
        pt_icon = '🔥' if pt['anyJustBroke'] else ('✅' if pt['anyBreakout'] else ('🕒' if pt['anyFormed'] else '－'))
        mp = matched_combos(r['row'], K['PINNED_COMBOS'])
        mm = matched_combos(r['row'], K['MOONSHOT_COMBOS'])
        ms = matched_combos(r['row'], K['STOCK_PICK_COMBOS'])
        rec = {
            '股票': r['stockId'], '名稱': r['name'] + (' 🔴即時' if r.get('realtime') else ''),
            '多方力道': r['total'], '評等': score_label(r['total']),
            '+DI': dm['plusDI'], '-DI': dm['minusDI'], 'ADX': dm['adx'], 'ADXR': dm['adxr'],
            '布林位置%': sig['bbPos'], '布林寬度%': bbw,
            'MACD狀態': (sig['macdState'] or '') + (('　' + combo_tag) if combo_tag else ''),
            'KDJ狀態': kd, 'RS(vs0050)%': info['rs'], '量比': info['vr'],
            '分價量表(POC)': vp_txt, '分價訊號': vp_sig,
            '成交量': b.volume[e], '漲跌幅%': chg, '收盤': last_c,
        }
        if extras_flags.get('pe'):
            p = r.get('peRange')
            rec['P/E'] = p['current'] if p else None
            rec['P/E區間'] = f"{p['min']:.1f}~{p['max']:.1f}" if p else '無資料'
        if extras_flags.get('rev'):
            rv = r.get('revRange')
            rec['營收YoY%(最新)'] = rv[-1]['yoy'] if rv else None
            rec['營收YoY(近3月)'] = '　'.join(f"{m['month']}月 {fnum(m['yoy'])}%" for m in rv) if rv else '無資料'
            rec['營收MoM%(最新)'] = rv[-1]['mom'] if rv else None
        if extras_flags.get('pxyoy'):
            py_ = r.get('priceYoYRange')
            rec['均價YoY%(最新)'] = py_[-1]['yoy'] if py_ else None
        if extras_flags.get('rev') and extras_flags.get('pxyoy'):
            rec['YoY乖離度(近3月合計)pp'] = r['row'].get('divTotal')
        if extras_flags.get('inst'):
            ir = r.get('instRange')
            rec['三大法人(近3月合計)'] = sum(m['net'] for m in ir) if ir else None
            if not ir and r.get('instRangeErr'):
                rec['三大法人(近3月合計)'] = None
        rec['回後買上漲'] = '✅ 符合進場' if pb['allPass'] else f"❌ {pb['requiredPassed']}/{pb['requiredTotal']}"
        rec['型態確認'] = pt_icon + ' ' + ('、'.join(pnames) if pnames else '無')
        def _t(c, kind):
            bt = best_t(combo_stats(c, kind, K, live))
            return f'(t{bt:.1f})' if bt is not None else ''
        rec['選股型命中'] = ' '.join(f'S{i + 1}' + _t(K['STOCK_PICK_COMBOS'][i], 'S') for i in ms)
        rec['指定組合命中'] = ' '.join([f'#{i + 1}' + ('⏱' if is_timing_combo(K['PINNED_COMBOS'][i]) else '')
                                    + _t(K['PINNED_COMBOS'][i], '#') for i in mp]
                                   + [f'M{i + 1}' + _t(K['MOONSHOT_COMBOS'][i], 'M') for i in mm])
        mw = matched_combos(r['row'], K.get('BT_WIN_COMBOS', []))
        mf = matched_combos(r['row'], K.get('BT_HOT_COMBOS', []))

        def _short(idxs, code, lst, kind, key):
            st_list = [(i, combo_stats(lst[i], kind, K, live)) for i in idxs]
            if kind == 'W':
                st_list.sort(key=lambda x: -(best_t(x[1]) or 0))
                txt = [f'{code}{i + 1}' + ('⏱' if is_timing_combo(lst[i]) else '') + (f'(t{best_t(s_):.1f})' if best_t(s_) is not None else '')
                       for i, s_ in st_list[:BT_HIT_SHOW]]
            else:
                st_list.sort(key=lambda x: -max([v for v in x[1]['moon'].values() if v is not None] or [0]))
                txt = [f'{code}{i + 1}(' + str(max([v for v in s_['moon'].values() if v is not None] or [0])) + '%)'
                       for i, s_ in st_list[:BT_HIT_SHOW]]
            more = len(idxs) - BT_HIT_SHOW
            return ' '.join(txt) + (f' +{more}' if more > 0 else '')
        rec['回測高勝率命中'] = _short(mw, 'W', K.get('BT_WIN_COMBOS', []), 'W', 'win')
        rec['回測飆股命中'] = _short(mf, 'F', K.get('BT_HOT_COMBOS', []), 'F', 'moon')
        rec['命中數'] = len(ms) + len(mp) + len(mm)
        # 命中類別數：S/#/M/W/F 五類中命中幾類（同類多個組合常共用同一訊號，類別數比總命中數更能反映訊號強弱）
        rec['命中類別數'] = sum(1 for x in (ms, mp, mm, mw, mf) if x)
        rec['組合型態'] = combo_style(ms, mp, mm, mw, mf)
        _reg = market_regime(r['row'])
        rec['大盤適用'], rec['_regOk'], rec['_regBad'] = regime_fit(
            [('S', K['STOCK_PICK_COMBOS'], ms), ('#', K['PINNED_COMBOS'], mp), ('M', K['MOONSHOT_COMBOS'], mm),
             ('W', K.get('BT_WIN_COMBOS', []), mw), ('F', K.get('BT_HOT_COMBOS', []), mf)], _reg, K)
        rec['_pbPass'] = pb['allPass']
        rec['_pt'] = 'justbreak' if pt['anyJustBroke'] else ('breakout' if pt['anyBreakout'] else ('forming' if pt['anyFormed'] else 'none'))
        recs.append(rec)
    df = pd.DataFrame(recs)
    if not len(df):
        return df
    # 欄位順序：股票、名稱之後依序放 股價、漲跌幅%、選股型命中、指定組合命中、型態確認；不顯示 評等、命中數、+DI/-DI/ADX/ADXR
    df = df.rename(columns={'收盤': '股價'}).drop(columns=['評等', '命中數', '+DI', '-DI', 'ADX', 'ADXR'], errors='ignore')
    front = ['股票', '名稱', '股價', '漲跌幅%', '命中類別數', '組合型態', '大盤適用', '選股型命中', '指定組合命中', '回測高勝率命中', '回測飆股命中', '型態確認']
    return df[[c for c in front if c in df.columns] + [c for c in df.columns if c not in front]]


def _xl_col(i):
    """0 → A, 25 → Z, 26 → AA"""
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _auto_widths(df):
    out = []
    for c in df.columns:
        vals = [str(c)] + [str(v) for v in df[c].tolist()[:2000]]
        out.append(min(max(max(len(x) for x in vals) + 2, 8), 40))
    return out


def _xlsx_builtin(sheets, widths=None, wrap=False):
    """沒有 openpyxl／xlsxwriter 時的備援：用標準函式庫 zipfile 直接產生 .xlsx（Excel 可正常開啟）"""
    import zipfile
    from xml.sax.saxutils import escape

    sa = ' s="1"' if wrap else ''

    def cell(ref, v):
        if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
            return ''
        if isinstance(v, (bool, np.bool_)):
            return f'<c r="{ref}" t="b"><v>{int(v)}</v></c>'
        if isinstance(v, (int, float, np.integer, np.floating)):
            return f'<c r="{ref}"{sa}><v>{v}</v></c>'
        txt = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', escape(str(v)))
        return f'<c r="{ref}" t="inlineStr"{sa}><is><t xml:space="preserve">{txt}</t></is></c>'

    bio = io.BytesIO()
    with zipfile.ZipFile(bio, 'w', zipfile.ZIP_DEFLATED) as z:
        names = [str(n)[:31] for n, _ in sheets]
        z.writestr('[Content_Types].xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   + ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                             for i in range(len(sheets))) + '</Types>')
        z.writestr('_rels/.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        z.writestr('xl/workbook.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                   + ''.join(f'<sheet name="{escape(n)}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, n in enumerate(names))
                   + '</sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + ''.join(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
                             for i in range(len(sheets)))
                   + f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   '</Relationships>')
        z.writestr('xl/styles.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                   '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
                   '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
                   '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                   '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                   '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                   '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs>'
                   '</styleSheet>')
        for si, (_, df) in enumerate(sheets):
            ws = widths if widths else _auto_widths(df)
            cols = ''.join(f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>' for i, w in enumerate(ws))
            rows = ['<row r="1">' + ''.join(cell(f'{_xl_col(j)}1', str(c)) for j, c in enumerate(df.columns)) + '</row>']
            for ri, rec in enumerate(df.itertuples(index=False), start=2):
                rows.append(f'<row r="{ri}">' + ''.join(cell(f'{_xl_col(j)}{ri}', v) for j, v in enumerate(rec)) + '</row>')
            z.writestr(f'xl/worksheets/sheet{si + 1}.xml',
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                       + (f'<cols>{cols}</cols>' if cols else '') + '<sheetData>' + ''.join(rows) + '</sheetData></worksheet>')
    return bio.getvalue()


def sheets_to_xlsx(sheets, widths=None, wrap=False):
    """[(工作表名, DataFrame), ...] → .xlsx bytes。
    依序嘗試 openpyxl → xlsxwriter → 內建備援（標準函式庫），任何環境都不會因為少裝套件而整頁當掉。
    widths：各欄寬（None＝依內容自動）；wrap：內容自動換行、靠上對齊。"""
    bad = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')   # Excel 不接受的控制字元

    def clean(d):
        d = d if isinstance(d, pd.DataFrame) else pd.DataFrame(d)
        d = d.copy()
        for c in d.columns:
            if d[c].dtype == object:
                d[c] = d[c].map(lambda v: bad.sub('', v) if isinstance(v, str) else v)
        return d
    sheets = [(str(n)[:31], clean(d)) for n, d in sheets]
    for engine in ('openpyxl', 'xlsxwriter'):
        try:
            __import__(engine)
        except ImportError:
            continue
        try:
            return _xlsx_with_engine(sheets, engine, widths, wrap)
        except Exception:  # noqa  這個引擎失敗就換下一個，最後用內建備援
            continue
    return _xlsx_builtin(sheets, widths, wrap)


def _xlsx_with_engine(sheets, engine, widths, wrap):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine=engine) as w:
        for name, df in sheets:
            df.to_excel(w, index=False, sheet_name=name)
            ws_ = widths if widths else _auto_widths(df)
            sh = w.sheets[name]
            if engine == 'openpyxl':
                from openpyxl.styles import Alignment
                for i, wd in enumerate(ws_):
                    sh.column_dimensions[_xl_col(i)].width = wd
                if wrap:
                    for row in sh.iter_rows(min_row=2):
                        for c in row:
                            c.alignment = Alignment(wrap_text=True, vertical='top')
            else:
                fmt = w.book.add_format({'text_wrap': True, 'valign': 'top'}) if wrap else None
                for i, wd in enumerate(ws_):
                    sh.set_column(i, i, wd, fmt)
    return bio.getvalue()


def df_to_excel_bytes(df, sheet):
    return sheets_to_xlsx([(sheet, df)])


# ════════════════════════════════════════════════════════════════════
#  AI 分析（OpenAI）
# ════════════════════════════════════════════════════════════════════
AI_SYSTEM = ('你是一位精通台股技術分析的資深操盤手，熟悉DMI趨向指標（+DI／-DI／ADX／ADXR）多方力道評分方法論，'
             '以及朱家泓《技術分析全攻略》的回後買上漲、頭肩底等進場型態判斷。請根據使用者提供的個股技術數據摘要，'
             '用繁體中文給出：1) 整體技術面研判（3-4句） 2) 進場時機與風險提示 3) 綜合建議（積極做多／可考慮／觀望／不建議）。'
             '語氣專業、精簡、避免空泛用詞，並提醒這僅為技術面參考，非投資建議。')


def build_ai_prompt(r):
    b, info = r['bars'], r['info']
    e = b.n - 1
    c, p = b.close[e], b.close[e - 1] if e >= 1 else b.close[e]
    chg = c - p
    chgp = chg / p * 100 if p else 0
    dm, pb, pt = info['dm'], info['pb'], info['pt']
    L = [f"股票：{r['stockId']} {r['name']}",
         f"資料日期：{b.date[e]}　收盤：{c}　漲跌：{chg:+.2f} ({chgp:+.2f}%)",
         f"成交量：{b.volume[e]:.0f}　量比(vs MA20量)：" + (f"{b.volume[e] / b.vm20[e]:.2f}" if b.vm20[e] else 'N/A') + 'x']
    L.append(f"MACD：DIF={b.macd[e]:.2f}　Signal={b.macdSig[e]:.2f}　柱狀={b.macdHist[e]:.2f}"
             + ('（偏多）' if b.macd[e] > b.macdSig[e] else '（偏空）'))
    if b.bbU[e] is not None:
        w = b.bbU[e] - b.bbL[e]
        L.append(f"布林通道：上軌={b.bbU[e]:.2f}　下軌={b.bbL[e]:.2f}　股價位置={((c - b.bbL[e]) / w * 100 if w else 50):.0f}%")
    prof, vps = info['prof'], info['vps']
    if prof:
        L.append(f"分價量表：POC={prof['pocPrice']:.2f}（區間 {prof['pocLow']:.2f}~{prof['pocHigh']:.2f}）"
                 + (f"　訊號：{vps['detail']}" if vps['detail'] else '　無訊號'))
    L += ['', f"【DMI多方力道評分】總分 {r['total']}/100",
          f"+DI={fnum(dm['plusDI'])}　-DI={fnum(dm['minusDI'])}　ADX={fnum(dm['adx'])}　ADXR={fnum(dm['adxr'])}",
          f"方向性 {dm['diPts']:.1f}/50、趨勢強度 {dm['adxPts']:.1f}/30、趨勢動能 {dm['adxrPts']:.1f}/20"]
    bull = [s[0] for s in dm['sigs'] if s[1] == 'bull']
    bear = [s[0] for s in dm['sigs'] if s[1] == 'bear']
    if bull:
        L.append('多頭訊號：' + '、'.join(bull))
    if bear:
        L.append('空頭訊號：' + '、'.join(bear))
    L += ['', f"【回後買上漲 8條件核對】必要條件通過 {pb['requiredPassed']}/{pb['requiredTotal']}" + ('（全數通過）' if pb['allPass'] else ''),
          '', '【型態確認，15種進場型態】']
    for x in pt['results']:
        stt = '🔥剛突破（較前一交易日新增）' if x['justBroke'] else ('✅已突破' if x['breakout'] else ('🕒成形中未突破' if x['formed'] else '－未偵測到'))
        L.append(f"・{x['name']}：{stt}" + (f"（{x['detail']}）" if x['detail'] else ''))
    return '\n'.join(L)


def run_ai(prompt, key, model):
    r = requests.post('https://api.openai.com/v1/chat/completions',
                      headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + key},
                      json=dict(model=model, temperature=0.4, max_tokens=900,
                                messages=[{'role': 'system', 'content': AI_SYSTEM}, {'role': 'user', 'content': prompt}]),
                      timeout=90)
    j = r.json()
    if r.status_code != 200:
        raise RuntimeError((j.get('error') or {}).get('message') or f'HTTP {r.status_code}')
    return j['choices'][0]['message']['content']


# ════════════════════════════════════════════════════════════════════
#  圖表
# ════════════════════════════════════════════════════════════════════
LINE_STYLE = {
    'hs': ('#ffd54f', '頭肩底頸線'), 'chs': ('#ff8a65', '複式頭肩底頸線'), 'nb': ('#81c784', 'N字底壓力'),
    'tb': ('#4dd0e1', '三重底壓力'), 'rb': ('#64b5f6', '圓弧底壓力'), 'fb': ('#ba68c8', '一字底整理區間高點'),
    'abc': ('#ff2ecc', 'ABC下降切線起點'), 'channel': ('#ffa726', '上升軌道線'),
    'blackk': ('#e57373', '大量黑K高點'), 'kbp': ('rgba(255,255,255,.85)', 'K線橫盤首日高點'),
}
MARKER_STYLE = {'harami_bear': '#ff8a65', 'harami_bull': '#81c784', 'morning_star': '#4dd0e1', 'evening_star': '#ff5252'}


def draw_chart(r, vp_params):
    import plotly.graph_objects as go
    b0 = r['bars']
    fb, _ = b0.filtered()
    e = fb.n - 1
    x = fb.date
    pt = r['info']['pt']
    shapes, ann = [], []
    for p in pt['results']:
        if not p['formed']:
            continue
        if p['id'] in LINE_STYLE:
            col, lbl = LINE_STYLE[p['id']]
            for k, ln in enumerate([p.get('line'), p.get('line2')]):
                if not ln or ln[0] < 0 or ln[0] >= fb.n:
                    continue
                i1, p1, sl = ln
                shapes.append(dict(type='line', xref='x', yref='y', x0=x[i1], y0=p1, x1=x[e], y1=p1 + sl * (e - i1),
                                   line=dict(color=col, width=2 if k == 0 else 1.3, dash='solid' if p['breakout'] else 'dash')))
                if k == 0:
                    ann.append(dict(x=x[i1], y=p1, text=lbl, showarrow=True, arrowhead=2, arrowcolor=col,
                                    font=dict(color=col, size=10), ax=-10, ay=-30))
                    if p['breakout']:
                        ann.append(dict(x=x[e], y=fb.close[e], text='★ 突破' + p['name'], showarrow=True, arrowhead=2,
                                        arrowcolor=col, font=dict(color=col, size=11), ax=10, ay=-35))
        elif p['id'] in MARKER_STYLE and p.get('marker'):
            mi, mp_, dr, lb = p['marker']
            if 0 <= mi < fb.n:
                col = MARKER_STYLE[p['id']]
                ann.append(dict(x=x[mi], y=mp_, text=('★ ' if p['breakout'] else '') + lb, showarrow=True, arrowhead=2,
                                arrowcolor=col, font=dict(color=col, size=11), ax=0, ay=-32 if dr == 'up' else 32))
    prof = r['info']['prof']
    if prof:
        x0 = x[max(0, e + 1 - vp_params['lookback'])]
        shapes.append(dict(type='rect', xref='x', yref='y', x0=x0, x1=x[e], y0=prof['pocLow'], y1=prof['pocHigh'],
                           fillcolor='rgba(255,215,0,.10)', line=dict(width=0)))
        shapes.append(dict(type='line', xref='x', yref='y', x0=x0, x1=x[e], y0=prof['pocPrice'], y1=prof['pocPrice'],
                           line=dict(color='#ffd700', width=1.2, dash='dot')))
        ann.append(dict(x=x[e], y=prof['pocPrice'], text=f"POC {prof['pocPrice']:.1f}", showarrow=False,
                        xanchor='left', font=dict(color='#ffd700', size=10)))
    zz = fb.zigzag(e)
    vc = ['#ef5350' if fb.close[i] >= fb.open[i] else '#26a69a' for i in range(fb.n)]
    hc = ['#ef5350' if (fb.macdHist[i] or 0) >= 0 else '#26a69a' for i in range(fb.n)]
    fig = go.Figure([
        go.Candlestick(x=x, open=fb.open, high=fb.high, low=fb.low, close=fb.close, name='K線',
                       increasing=dict(line=dict(color='#ef5350'), fillcolor='#ef5350'),
                       decreasing=dict(line=dict(color='#26a69a'), fillcolor='#26a69a')),
        go.Scatter(x=x, y=fb.ma5, name='MA5', line=dict(color='#ffeb3b', width=1.2)),
        go.Scatter(x=x, y=fb.ma10, name='MA10', line=dict(color='#ff9800', width=1.2)),
        go.Scatter(x=x, y=fb.ma20, name='MA20', line=dict(color='#2196f3', width=1.2)),
        go.Scatter(x=x, y=fb.ma60, name='MA60', line=dict(color='#9c27b0', width=1.2)),
        go.Scatter(x=[x[p[0]] for p in zz], y=[p[1] for p in zz], name='轉折波', mode='lines+markers',
                   line=dict(color='#00e5ff', width=1.8), marker=dict(size=5, color='#00e5ff')),
        go.Scatter(x=x, y=fb.bbU, name='BB上軌', line=dict(color='rgba(100,200,255,.4)', width=1, dash='dot'), showlegend=False),
        go.Scatter(x=x, y=fb.bbL, name='BB下軌', line=dict(color='rgba(100,200,255,.4)', width=1, dash='dot'),
                   fill='tonexty', fillcolor='rgba(100,200,255,.04)', showlegend=False),
        go.Bar(x=x, y=fb.volume, marker_color=vc, name='成交量', opacity=.7, yaxis='y2'),
        go.Scatter(x=x, y=fb.vm20, name='量MA20', line=dict(color='#ff9800', width=1.5), yaxis='y2'),
        go.Bar(x=x, y=fb.macdHist, marker_color=hc, name='MACD柱', opacity=.8, yaxis='y3'),
        go.Scatter(x=x, y=fb.macd, name='MACD', line=dict(color='#2196f3', width=1.5), yaxis='y3'),
        go.Scatter(x=x, y=fb.macdSig, name='Signal', line=dict(color='#ff9800', width=1.5), yaxis='y3'),
    ])
    fig.update_layout(
        title=dict(text=f"{r['stockId']} {r['name']} 技術分析圖", font=dict(size=14)),
        template='plotly_dark', paper_bgcolor='rgba(10,14,26,1)', plot_bgcolor='rgba(15,20,35,1)',
        height=680, margin=dict(l=50, r=130, t=45, b=25), shapes=shapes, annotations=ann,
        xaxis=dict(rangeslider=dict(visible=False), gridcolor='rgba(255,255,255,.04)', type='category',
                   nticks=12, anchor='free', position=0),
        yaxis=dict(domain=[0.42, 1], gridcolor='rgba(255,255,255,.04)'),
        yaxis2=dict(domain=[0.22, 0.40], gridcolor='rgba(255,255,255,.04)'),
        yaxis3=dict(domain=[0, 0.20], gridcolor='rgba(255,255,255,.04)'),
        legend=dict(orientation='v', x=1.01, y=1, xanchor='left', bgcolor='rgba(0,0,0,0)', font=dict(size=10)),
    )
    return fig


# ════════════════════════════════════════════════════════════════════
#  批次分析主流程（Streamlit 與每日排程共用）
# ════════════════════════════════════════════════════════════════════
def batch_analyze(stocks, token, days, use_rt, ex_flags, vpp, delay=0.6, log=None, progress=None):
    log = log or (lambda m: None)
    names = fetch_name_map(token)
    try:
        bm = fetch_benchmark(token, days)
    except Exception:  # noqa
        bm = None
    rt_map = {}
    if use_rt:
        rt_map, rt_err = fetch_realtime(token, stocks)
        log(f'⚠️ 即時快照抓取失敗（已略過）：{rt_err}' if rt_err else f'🔴 已取得 {len(rt_map)} 檔即時快照')
    results = []
    for i, sid in enumerate(stocks):
        try:
            rows_full = fetch_price(sid, token, max(days, LONG_HISTORY_DAYS))
            rt = False
            if use_rt and sid in rt_map:
                rows_full, rt = merge_realtime(rows_full, rt_map[sid])
            rows = trim_rows(rows_full, days)
            bl = Bars(rows_full)
            extras = dict(realtime=rt, peRange=None, revRange=None, priceYoYRange=None,
                          instRange=None, instRangeErr=None, tech=tech_extras(bl, bl.n - 1))
            if ex_flags.get('pe'):
                try:
                    extras['peRange'] = fetch_pe_range(sid, token)
                except Exception:  # noqa
                    pass
            if ex_flags.get('rev'):
                try:
                    extras['revRange'] = fetch_revenue_3m(sid, token)
                    extras['revFeats'] = rev_feats(fetch_revenue_hist(sid, token, 800), _ds(dt.date.today()))
                except Exception:  # noqa
                    pass
            if ex_flags.get('pxyoy'):
                try:
                    ms = [dict(year=m['year'], month=m['month']) for m in extras['revRange']] if extras['revRange'] else None
                    extras['priceYoYRange'] = fetch_monthly_avg_price_yoy(sid, token, ms)
                except Exception:  # noqa
                    pass
            if ex_flags.get('inst'):
                try:
                    extras['instRange'], extras['instFlow'] = fetch_inst_monthly(sid, token, with_flow=True)
                except Exception as ex2:  # noqa
                    extras['instRangeErr'] = str(ex2)
            r = analyze_stock(sid, names.get(sid, sid), rows, bm, vpp, extras)
            results.append(r)
            log(f"✅ {sid} {r['name']}  得分:{r['total']}" + ('　🔴即時' if rt else ''))
        except Exception as ex:  # noqa
            log(f'❌ {sid} 失敗：{ex}')
        if progress:
            progress(i + 1, len(stocks))
        if i < len(stocks) - 1 and delay:
            time.sleep(delay)
    results.sort(key=lambda r: -r['total'])
    return results


# ════════════════════════════════════════════════════════════════════
#  命中組合：匯出與每日追蹤
#  追蹤紀錄存在程式同資料夾的 combo_hits_log.csv（每次批次分析自動累加，
#  同一天同一檔只記一次），之後用「更新追蹤報酬」抓最新股價，計算命中後的
#  實際表現，並依組合彙總「實盤」勝率，跟回測記錄的勝率對照。
# ════════════════════════════════════════════════════════════════════
HIT_LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'combo_hits_log.csv')
HIT_COLS = ['日期', '股票代號', '股票名', '股價', '漲跌%', '命中組合', '勝率', '飆股比例', 't值', '命中編號']


# ── 每個組合的勝率／飆股比例／t值(同日調整) ──
# 優先用「這次程式裡跑過的回測」即時計算；沒跑回測時，退回 2026-09-27 台股回測檔案的記錄值
# （COMBO_REF_STATS，只有當時進榜的組合才有），再退回各組合原本的歷史記錄。
def combo_norm_key(combo):
    return ' ＋ '.join(sorted(combo))


def compute_live_combo_stats(df, combos, moon_threshold=30.0):
    """用回測原始紀錄算每個組合 5/10/20日 的勝率、t值(同日調整)、飆股比例，回傳 {正規化key: stats}"""
    out = {}
    if df is None or not len(df):
        return out
    flags = build_condition_flags(df)
    uniq = {combo_norm_key(c): c for c in combos}
    for h in HORIZONS:
        col = f'ret{h}d'
        if col not in df.columns:
            continue
        valid = df[col].notna().values
        ret = df[col].values.astype(float)
        dm = df.groupby('evalDate')[col].transform('mean').values.astype(float) if 'evalDate' in df.columns else np.nanmean(ret)
        exr = ret - dm
        for k, c in uniq.items():
            if any(x not in flags for x in c):
                continue
            m = valid & np.logical_and.reduce([flags[x].values for x in c])
            n = int(m.sum())
            if n < 5:
                continue
            r, ex = ret[m], exr[m]
            sd = ex.std(ddof=1) if n > 1 else 0
            d = out.setdefault(k, {'win': {}, 't': {}, 'moon': {}, 'n': {}, 'src': '本次回測'})
            d['n'][str(h)] = n
            d['win'][str(h)] = round((r > 0).mean() * 100, 1)
            d['t'][str(h)] = round(ex.mean() / (sd / math.sqrt(n)), 2) if sd > 0 else None
            d['moon'][str(h)] = round((r > moon_threshold).mean() * 100, 1)
    return out


def combo_stats(combo, kind, K, live=None):
    """kind: 'S' / '#' / 'M'；回傳 {'win':{h:v}, 't':{h:v}, 'moon':{h:v}, 'src':...}"""
    k = combo_norm_key(combo)
    if live and k in live:
        return live[k]
    ref = K.get('COMBO_REF_STATS', {}).get(k, {})
    st_ = {'win': dict(ref.get('win', {})), 't': dict(ref.get('t', {})), 'moon': dict(ref.get('moon', {})),
           'src': '回測記錄值' if ref else ''}
    key = ' ＋ '.join(combo)
    if kind == 'S':
        for h, e in (K['STOCK_PICK_STATS'].get(key) or {}).items():
            st_['win'].setdefault(h, e['win'])
            st_['t'].setdefault(h, e['t'])
        st_['src'] = st_['src'] or '選股型記錄'
    elif kind == '#' and not st_['win']:
        for h, v in (K['PINNED_COMBO_WINRATES'].get(key) or {}).items():
            st_['win'][h] = v
        st_['src'] = st_['src'] or '舊版記錄'
    elif kind == 'M' and not st_['moon']:
        for h, e in (K['MOONSHOT_COMBO_STATS'].get(key) or {}).items():
            st_['moon'][h] = e['pct']
        st_['src'] = st_['src'] or '舊版記錄'
    elif kind == 'W':
        for h, e in (K.get('BT_WIN_STATS', {}).get(key) or {}).items():
            st_['win'].setdefault(h, e['win'])
            st_['t'].setdefault(h, e['t'])
        st_['src'] = st_['src'] or '三年回測'
    elif kind == 'F':
        for h, e in (K.get('BT_HOT_STATS', {}).get(key) or {}).items():
            st_['moon'].setdefault(h, e['pct'])
        st_['src'] = st_['src'] or '三年回測'
    return st_


def fmt_hz(d, pct=True):
    if not d:
        return '—'
    parts = [f"{h}日{v}{'%' if pct else ''}" for h, v in sorted(d.items(), key=lambda kv: int(kv[0])) if v is not None]
    return '/'.join(parts) if parts else '—'


def best_t(st_):
    vals = [v for v in (st_.get('t') or {}).values() if v is not None]
    return max(vals) if vals else None


def hit_records(results, K, live=None):
    """批次分析結果中有命中指定組合的股票（欄位：日期/代號/名稱/股價/漲跌%/命中組合/勝率/飆股比例/t值）
    「勝率／飆股比例／t值」每一行對應「命中組合」的同一行。"""
    out = []
    for r in results:
        mp = matched_combos(r['row'], K['PINNED_COMBOS'])
        mm = matched_combos(r['row'], K['MOONSHOT_COMBOS'])
        ms = matched_combos(r['row'], K['STOCK_PICK_COMBOS'])
        mw = matched_combos(r['row'], K.get('BT_WIN_COMBOS', []))
        mf = matched_combos(r['row'], K.get('BT_HOT_COMBOS', []))
        if not mp and not mm and not ms and not mw and not mf:
            continue
        b = r['bars']
        e = b.n - 1
        if not b.close[e] or b.close[e] <= 0:   # 0價（無成交／暫停交易）不列入命中
            continue
        pc = b.close[e - 1] if e >= 1 else b.close[e]
        lines, codes, wins, moons, ts = [], [], [], [], []

        def add(code, tag, c, kind):
            st_ = combo_stats(c, kind, K, live)
            bt = best_t(st_)
            lines.append(f"{code}{tag} " + '＋'.join(c) + (f'（t={bt}）' if bt is not None else ''))
            codes.append(code)
            wins.append(fmt_hz(st_['win']))
            moons.append(fmt_hz(st_['moon']))
            ts.append(fmt_hz(st_['t'], pct=False))
        for i in ms:
            add(f'S{i + 1}', '[選股型]', K['STOCK_PICK_COMBOS'][i], 'S')
        for i in mp:
            c = K['PINNED_COMBOS'][i]
            add(f'#{i + 1}', '[高勝率·擇時型]' if is_timing_combo(c) else '[高勝率]', c, '#')
        for i in mm:
            add(f'M{i + 1}', '[高標股]', K['MOONSHOT_COMBOS'][i], 'M')
        # 回測自動入選：每類只取最強的5組（避免一檔命中幾十組時紀錄過長）
        for i in sorted(mw, key=lambda i: -(best_t(combo_stats(K['BT_WIN_COMBOS'][i], 'W', K, live)) or 0))[:5]:
            c = K['BT_WIN_COMBOS'][i]
            add(f'W{i + 1}', '[回測高勝率·擇時型]' if is_timing_combo(c) else '[回測高勝率]', c, 'W')
        for i in sorted(mf, key=lambda i: -max(list(combo_stats(K['BT_HOT_COMBOS'][i], 'F', K, live)['moon'].values()) or [0]))[:5]:
            add(f'F{i + 1}', '[回測飆股]', K['BT_HOT_COMBOS'][i], 'F')
        out.append({'日期': b.date[e], '股票代號': r['stockId'], '股票名': r['name'], '股價': round(b.close[e], 2),
                    '漲跌%': round((b.close[e] - pc) / pc * 100, 2) if pc else None,
                    '命中組合': '\n'.join(lines), '勝率': '\n'.join(wins), '飆股比例': '\n'.join(moons),
                    't值': '\n'.join(ts), '命中編號': ' '.join(codes)})
    return out


def hits_excel_bytes(recs, sheet='命中組合股票'):
    df = pd.DataFrame(recs, columns=HIT_COLS)
    return sheets_to_xlsx([(sheet, df)], widths=[12, 10, 12, 10, 9, 80, 26, 26, 26, 22], wrap=True)


def load_hit_log():
    try:
        df = pd.read_csv(HIT_LOG_FILE, dtype={'股票代號': str, '日期': str})
        # 股價≤0 的紀錄是資料源當天回傳0價（無成交／暫停交易）造成的假命中，直接排除
        px_ = pd.to_numeric(df['股價'], errors='coerce')
        return df[~(px_ <= 0)].reset_index(drop=True)
    except Exception:  # noqa
        return pd.DataFrame(columns=HIT_COLS)


def append_hit_log(recs):
    """把命中紀錄累加到追蹤檔（同一天同一檔只保留一筆，以最新的為準），回傳新增筆數"""
    if not recs:
        return 0
    old = load_hit_log()
    new = pd.DataFrame(recs, columns=HIT_COLS)
    before = set(zip(old['日期'], old['股票代號'])) if len(old) else set()
    df = pd.concat([old, new], ignore_index=True).drop_duplicates(['日期', '股票代號'], keep='last')
    df = df.sort_values(['日期', '股票代號']).reset_index(drop=True)
    df.to_csv(HIT_LOG_FILE, index=False, encoding='utf-8-sig')
    return sum(1 for r in recs if (r['日期'], r['股票代號']) not in before)


def track_performance(log_df, token, delay=0.3, progress=None):
    """抓每檔追蹤股票從命中日到今天的股價，算命中後的表現"""
    if not len(log_df):
        return pd.DataFrame()
    first = min(log_df['日期'])
    days = (dt.date.today() - dt.date.fromisoformat(first)).days + 10
    sids = sorted(set(log_df['股票代號']))
    px = {}
    for i, sid in enumerate(sids):
        try:
            px[sid] = fetch_price(sid, token, days)
        except Exception:  # noqa
            px[sid] = []
        if progress:
            progress(i + 1, len(sids))
        if delay and i < len(sids) - 1:
            time.sleep(delay)
    out = []
    for r in log_df.to_dict('records'):
        rows = px.get(r['股票代號']) or []
        dates = [x['date'] for x in rows]
        rec = {k: r.get(k) for k in ('日期', '股票代號', '股票名', '股價', '命中編號', '命中組合')}
        import bisect
        k = bisect.bisect_right(dates, str(r['日期'])) - 1   # 命中日（或之前最近的交易日）
        if k < 0:
            rec.update({'持有天數': None, '最新價': None, '目前報酬%': None})
            out.append(rec)
            continue
        ent = rows[k]['close']
        after = rows[k + 1:]
        rec['持有天數'] = len(after)
        rec['最新價'] = rows[-1]['close']
        rec['目前報酬%'] = round((rows[-1]['close'] - ent) / ent * 100, 2)
        for h in (5, 10, 20):
            rec[f'{h}日報酬%'] = round((after[h - 1]['close'] - ent) / ent * 100, 2) if len(after) >= h else None
        rec['最大漲幅%'] = round((max(x['high'] for x in after) - ent) / ent * 100, 2) if after else None
        rec['最大回檔%'] = round((min(x['low'] for x in after) - ent) / ent * 100, 2) if after else None
        out.append(rec)
    return pd.DataFrame(out)


def combo_track_summary(perf, K):
    """依組合彙總命中後的實際表現，並附上回測記錄值對照"""
    if perf is None or not len(perf):
        return pd.DataFrame()
    rows = []
    ex = perf.assign(編號=perf['命中編號'].fillna('').str.split()).explode('編號')
    ex = ex[ex['編號'].astype(str).str.len() > 0].copy()

    def cur_combo(code):
        idx = int(code[1:]) - 1
        lst = (K['STOCK_PICK_COMBOS'] if code.startswith('S') else K['PINNED_COMBOS'] if code.startswith('#')
               else K.get('BT_WIN_COMBOS', []) if code.startswith('W') else K.get('BT_HOT_COMBOS', []) if code.startswith('F')
               else K['MOONSHOT_COMBOS'])
        return lst[idx] if 0 <= idx < len(lst) else None

    def logged_combo(code, txt):
        """從追蹤紀錄當時寫下的「命中組合」文字取出這個編號的組合（清單改版後舊紀錄仍對得上）"""
        for line in str(txt or '').split('\n'):
            m = re.match(r'^' + re.escape(code) + r'\[[^\]]*\]\s*(.+?)(（t=.*）)?$', line.strip())
            if m:
                return m.group(1).strip()
        return None
    ex['_combo'] = [logged_combo(c, t) or ('＋'.join(cur_combo(c)) if cur_combo(c) else '') for c, t in
                    zip(ex['編號'], ex['命中組合'] if '命中組合' in ex.columns else [None] * len(ex))]
    for (code, ctext), g in ex.groupby(['編號', '_combo']):
        combo = cur_combo(code)
        if combo is None or '＋'.join(combo) != ctext:   # 清單已改版：顯示當時的組合，不套用現在的回測記錄
            row_code, combo_txt, rec_txt = code + '(舊版)', ctext, '（舊版清單組合）'
        else:
            row_code, combo_txt = code, ctext
            if code.startswith('S'):
                rec_txt = format_stockpick_stats(K['STOCK_PICK_STATS'].get(' ＋ '.join(combo)))
            elif code.startswith('#'):
                rec_txt = ('[擇時型] ' if is_timing_combo(combo) else '') + format_winrates(K['PINNED_COMBO_WINRATES'].get(' ＋ '.join(combo)))
            elif code.startswith('W'):
                rec_txt = '　'.join(f"{h}日 勝率{e['win']}% t={e['t']}" for h, e in sorted(K['BT_WIN_STATS'].get(' ＋ '.join(combo), {}).items(), key=lambda kv: int(kv[0])))
            elif code.startswith('F'):
                rec_txt = format_moonshot_stats(K['BT_HOT_STATS'].get(' ＋ '.join(combo)))
            else:
                rec_txt = format_moonshot_stats(K['MOONSHOT_COMBO_STATS'].get(' ＋ '.join(combo)))
        row = {'編號': row_code, '條件組合': combo_txt, '命中次數': len(g),
               '目前平均報酬%': round(g['目前報酬%'].mean(), 2) if g['目前報酬%'].notna().any() else None}
        for h in (5, 10, 20):
            v = g[f'{h}日報酬%'].dropna() if f'{h}日報酬%' in g else pd.Series(dtype=float)
            row[f'{h}日已滿筆數'] = len(v)
            row[f'{h}日平均%'] = round(v.mean(), 2) if len(v) else None
            row[f'{h}日勝率%'] = round((v > 0).mean() * 100, 1) if len(v) else None
        row['回測記錄'] = rec_txt
        rows.append(row)
    df = pd.DataFrame(rows)
    return df.sort_values(['命中次數'], ascending=False).reset_index(drop=True)


# ── 每日排程（命令列模式）──
DAILY_USAGE = """每日追蹤（命令列）：
  python stock_analyzer_poc.py daily [--list top100|上市清單|上櫃清單|上市+上櫃|0050成分股|我的清單] \
                                     [--token XXX] [--days 180] [--no-extras] [--out 資料夾]
  token 也可用環境變數 FINMIND_TOKEN。流程：抓清單 → 批次分析 → 命中組合累加到 combo_hits_log.csv
  → 更新所有追蹤股票的報酬 → 輸出「命中組合_日期.xlsx」（今日命中／追蹤明細／組合彙總 三個工作表）。
  建議排程在台股收盤資料更新後（例如每天 18:00）執行。"""


def run_daily(argv):
    import argparse
    ap = argparse.ArgumentParser(description=DAILY_USAGE, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--list', default='top100')
    ap.add_argument('--token', default=os.environ.get('FINMIND_TOKEN', ''))
    ap.add_argument('--days', type=int, default=180)
    ap.add_argument('--no-extras', action='store_true')
    ap.add_argument('--delay', type=float, default=0.6)
    ap.add_argument('--out', default=os.path.dirname(os.path.abspath(__file__)))
    a = ap.parse_args(argv)
    if not a.token:
        print('❌ 請用 --token 或環境變數 FINMIND_TOKEN 提供 FinMind API Token')
        return 1
    K = CONSTANTS()
    if a.list == 'top100':
        res, used = fetch_market_snapshot(a.token, K['TWSE_LIST'], K['TPEX_LIST'], log=print)
        gain = [r['id'] for r in sorted(res, key=lambda r: -r['pct'])[:100]]
        vol = [r['id'] for r in sorted(res, key=lambda r: -r['volume'])[:100]]
        stocks = list(dict.fromkeys(gain + vol))
        print(f'✅ {used} 漲幅前100＋成交量前100，合併去重 {len(stocks)} 檔')
    else:
        src = {'上市清單': K['TWSE_LIST'], '上櫃清單': K['TPEX_LIST'], '上市+上櫃': K['TWSE_LIST'] + K['TPEX_LIST'],
               '0050成分股': K['TW0050_LIST']}.get(a.list)
        if src is None:
            src = load_custom_lists(K).get(a.list, [])
        stocks = list(src)
    ex = dict(pe=False, rev=not a.no_extras, pxyoy=not a.no_extras, inst=not a.no_extras)
    results = batch_analyze(stocks, a.token, a.days, False, ex, VP_DEFAULTS, a.delay, log=print)
    _reg = batch_market_regime(results)
    if _reg:
        print(regime_banner(_reg).replace('**', ''))
    recs = hit_records(results, K)
    added = append_hit_log(recs)
    print(f'📌 今日命中 {len(recs)} 檔，新增 {added} 筆追蹤紀錄 → {HIT_LOG_FILE}')
    perf = track_performance(load_hit_log(), a.token, delay=0.3)
    summ = combo_track_summary(perf, K)
    fn = os.path.join(a.out, f'命中組合_{dt.date.today()}.xlsx')
    with open(fn, 'wb') as f_:
        f_.write(sheets_to_xlsx([('今日命中', pd.DataFrame(recs, columns=HIT_COLS)), ('追蹤明細', perf), ('組合彙總', summ)]))
    print(f'📥 已輸出 {fn}')
    return 0


# ════════════════════════════════════════════════════════════════════
#  Streamlit 介面
# ════════════════════════════════════════════════════════════════════
LISTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'stock_lists.json')


def load_custom_lists(K):
    try:
        with open(LISTS_FILE, 'r', encoding='utf-8') as f:
            d = json.load(f)
    except Exception:  # noqa
        d = {}
    d.setdefault('我的清單', K['MY_LIST_DEFAULT'])
    for k in ('我的清單1', '我的清單2', '我的清單3'):
        d.setdefault(k, [])
    return d


def save_custom_lists(d):
    try:
        with open(LISTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
        return True
    except Exception:  # noqa
        return False


def parse_stocks(txt):
    import re
    return [s.strip() for s in re.split(r'[\n,，、\s]+', txt or '') if s.strip()]


CSS = """
<style>
.block-container{padding-top:1.4rem}
.tagb{display:inline-block;padding:2px 9px;margin:2px 4px 2px 0;border-radius:10px;font-size:12px}
.bull{color:#00c864;border:1px solid #00c86466}
.bear{color:#ff5555;border:1px solid #ff555566}
.neu{color:#aaa;border:1px solid #aaaaaa55}
.verdict{border-radius:10px;padding:10px 16px;font-weight:700;margin-bottom:10px}
.small{color:#888;font-size:12px}
</style>
"""


def main():
    import streamlit as st
    st.set_page_config(page_title='技術分析全攻略 · 個股評分系統', page_icon='📊', layout='wide')
    st.markdown(CSS, unsafe_allow_html=True)
    K = CONSTANTS()
    ss = st.session_state
    ss.setdefault('lists', load_custom_lists(K))
    ss.setdefault('stocks_text', '\n'.join(K['TWSE_LIST']))
    ss.setdefault('batch', None)
    ss.setdefault('bt_df', None)
    ss.setdefault('ai', {})
    ss.setdefault('msg', '')

    def list_choice_changed():
        k = ss['list_choice']
        src = {'上市清單': K['TWSE_LIST'], '上櫃清單': K['TPEX_LIST'], '0050成分股': K['TW0050_LIST']}.get(k)
        if src is None:
            src = ss['lists'].get(k, [])
        ss['stocks_text'] = '\n'.join(src)

    # ── 漲幅／成交量前100：在畫出輸入框前先把清單換掉 ──
    pending = ss.pop('pending_top100', None)
    if pending:
        tok = ss.get('token', '').strip()
        if not tok:
            ss['msg'] = '請先輸入 FinMind API Token（需 Backer/Sponsor 方案）'
        else:
            try:
                with st.spinner('📡 抓取今日全市場行情中...'):
                    res, used = fetch_market_snapshot(tok, K['TWSE_LIST'], K['TPEX_LIST'])
                key = 'pct' if pending == 'gain' else 'volume'
                top = sorted(res, key=lambda r: -r[key])[:100]
                ss['stocks_text'] = '\n'.join(r['id'] for r in top)
                t0 = top[0]
                ss['msg'] = (f"✅ 已套入 {used} {'漲幅' if pending == 'gain' else '成交量'}前{len(top)}（最高：{t0['id']} {t0['name']} "
                             + (f"+{t0['pct']:.2f}%）" if pending == 'gain' else f"量={t0['volume']:,.0f} 股）"))
                ss['auto_batch'] = True
            except Exception as ex:  # noqa
                ss['msg'] = f'❌ {ex}'

    # ────────────────────────────── 側邊欄 ──────────────────────────────
    with st.sidebar:
        st.markdown('### 📊 技術分析全攻略')
        st.caption('朱家泓方法論 · 個股評分系統（POC 分價量表版）')
        st.text_input('FinMind API Token', type='password', key='token')
        lst_names = ['上市清單', '上櫃清單', '0050成分股', '我的清單', '我的清單1', '我的清單2', '我的清單3']
        st.selectbox('載入股票清單', lst_names, key='list_choice', on_change=list_choice_changed,
                     format_func=lambda k: k + (f"({len(ss['lists'].get(k, []))})" if k.startswith('我的') else ''))
        st.text_area('批次股票代號（每行一個或逗號分隔）', key='stocks_text', height=140)
        c1, c2 = st.columns(2)
        save_to = c1.selectbox('存到', ['我的清單', '我的清單1', '我的清單2', '我的清單3'], label_visibility='collapsed', key='save_to')
        if c2.button('💾 儲存清單', use_container_width=True):
            stocks = parse_stocks(ss['stocks_text'])
            if stocks:
                ss['lists'][save_to] = stocks
                save_custom_lists(ss['lists'])
                ss['msg'] = f'✅ 已存入「{save_to}」（{len(stocks)}檔）'
        days = st.slider('分析天數', 90, 365, 180, 30, key='days')
        use_rt = st.checkbox('🔴 加入盤中即時股價（需 sponsor，非 sponsor 自動略過）', value=True, key='use_rt')
        st.caption('以下每勾一項，每檔股票多打1次API：')
        ex_flags = dict(pe=st.checkbox('📐 近3年P/E區間', key='x_pe'), rev=st.checkbox('📈 近3月營收YoY/MoM', key='x_rev'),
                        pxyoy=st.checkbox('💹 近3月均價YoY', key='x_pxyoy'),
                        inst=st.checkbox('🏦 三大法人買賣超（近3月）', key='x_inst'))
        req_delay = st.number_input('每檔間隔秒數（避免API限速）', 0.0, 5.0, 0.6, 0.1, key='req_delay')
        st.checkbox('📌 批次分析後自動把命中組合股票加入追蹤', value=True, key='auto_track')
        run_batch = st.button('🔍 批次分析', type='primary', use_container_width=True, key='run_batch')
        cg, cv = st.columns(2)
        if cg.button('🔥 漲幅前100', use_container_width=True):
            ss['pending_top100'] = 'gain'
            st.rerun()
        if cv.button('📊 成交量前100', use_container_width=True):
            ss['pending_top100'] = 'vol'
            st.rerun()
        if ss['msg']:
            st.caption(ss['msg'])

        with st.expander('📐 分價量表(POC)參數', expanded=False):
            vpp = dict(
                lookback=st.number_input('回看天數', 40, 250, VP_DEFAULTS['lookback'], 10, key='vp_lb'),
                bins=st.number_input('價格區間數', 10, 60, VP_DEFAULTS['bins'], 2, key='vp_bins'),
                shadow_mult=st.number_input('長影線：影線 ≥ 實體 × ', 0.3, 3.0, VP_DEFAULTS['shadow_mult'], 0.1, key='vp_sh'),
                vol_mult=st.number_input('帶量：成交量 > 20日均量 × ', 1.0, 3.0, VP_DEFAULTS['vol_mult'], 0.1, key='vp_vol'),
                body_mult=st.number_input('長K：實體 > 20日平均實體 × ', 1.0, 3.0, VP_DEFAULTS['body_mult'], 0.1, key='vp_body'),
                touch_tol=st.number_input('觸及容忍（區間寬度倍數）', 0.0, 2.0, VP_DEFAULTS['touch_tol'], 0.1, key='vp_tol'),
            )
            st.caption('守穩POC：前日收在POC區上方→今日回測POC區、長下影線、收盤守住POC價。'
                       '突破POC：帶量長紅收盤突破POC區。反彈遇壓：前日收在POC區下方→反彈到POC區、長上影線、收盤壓回POC價下。'
                       '破位停損：帶量長黑收盤跌破POC區。')

        st.divider()
        st.markdown('**🔬 歷史回測分析**')
        bt_univ = st.selectbox('回測股票池', ['上市清單', '上櫃清單', '上市+上櫃', '目前輸入框清單'], key='bt_univ')
        bt_seg = st.selectbox('回測期間', BT_SEG_OPTIONS, key='bt_seg',
                              help='三年一次跑完資料量太大（容易記憶體不足當掉），改成分三次：每次跑一年，各自下載原始紀錄檔，最後三個檔一起載入合併分析')
        if bt_seg == BT_SEG_OPTIONS[0]:
            bt_months = st.number_input('回測天數（月）', 1, 36, 3, 1, key='bt_months')
            bt_off = 0
        else:
            bt_months, bt_off = 12, (BT_SEG_OPTIONS.index(bt_seg) - 1) * 12
            st.caption(f'本次評估區間：{bt_off}～{bt_off + 12} 個月前。每段跑完到「🔬 歷史回測」分頁按「產生回測原始紀錄檔」下載；'
                       '三段都下載後，三個檔一起拖進下方「載入」即可合併分析，多因子／飆股搜尋會自動逐段（第1／2／3段）驗證。')
        bt_div = st.checkbox('📈 近3月YoY乖離度（股價歷史要抓超過1年）', key='bt_div')
        bt_inst = st.checkbox('🏦 三大法人買賣超（近3月）', key='bt_inst')
        bt_liq = st.number_input('最低近20日均成交金額（萬元，0＝不篩）', 0, 1000000, 0, 100, key='bt_liq')
        bt_delay = st.number_input('回測每檔間隔秒數', 0.0, 5.0, 0.3, 0.1, key='bt_delay')
        run_bt = st.button('🔬 執行歷史回測', use_container_width=True, key='run_bt')
        ups = st.file_uploader('或載入先前匯出的回測原始紀錄（.csv / .csv.gz，可一次選多個檔合併，例如三段各一個）',
                               type=['csv', 'gz'], accept_multiple_files=True)
        up_sig = tuple(sorted((u.name, u.size) for u in ups)) if ups else None
        if up_sig and ss.get('bt_loaded_name') != up_sig:
            try:
                with st.spinner(f'載入 {len(ups)} 個檔案中…'):
                    release_bt_state(ss)
                    ss['bt_df'] = load_bt_files(ups)
                    ss['bt_ver'] = ss.get('bt_ver', 0) + 1
                ss['bt_loaded_name'] = up_sig
                ss['bt_seg_label'] = ''
            except Exception as ex:  # noqa
                st.error(f'讀取失敗：{ex}')

        st.divider()
        st.text_input('OpenAI API Key（選填）', type='password', key='openai_key')
        st.selectbox('AI 模型', ['gpt-4o-mini', 'gpt-4o'], key='openai_model')
        st.caption('評分：🟢80+ 積極做多　🔵65-79 可考慮　🟡50-64 觀望　🔴<50 不適合')

    st.title('📊 技術分析全攻略 · 個股評分分析系統')
    st.caption('多方力道＝DMI（+DI／-DI／ADX／ADXR）評分；回後買上漲、15種進場型態沿用朱家泓《技術分析全攻略》；分價量表改以 POC 為界')

    token = ss.get('token', '').strip()

    # ────────────────────────────── 批次分析 ──────────────────────────────
    if run_batch or ss.pop('auto_batch', False):
        stocks = parse_stocks(ss['stocks_text'])
        if not token:
            st.error('請輸入 FinMind API Token')
        elif not stocks:
            st.error('請輸入至少一個股票代號')
        else:
            try:
                api_get(dict(dataset='TaiwanStockInfo'), token)
            except Exception as ex:  # noqa
                st.error(f'❌ 無法連線至 FinMind API：{ex}')
                st.stop()
            prog = st.progress(0.0, text='批次分析中...')
            logbox = st.empty()
            logs = []

            def _log(msg):
                logs.append(msg)
                logbox.caption('\n\n'.join(logs[-6:]))

            def _prog(i, n):
                prog.progress(i / n, text=f'批次分析中：{i} / {n}')
            results = batch_analyze(stocks, token, days, use_rt, ex_flags, vpp, req_delay, _log, _prog)
            prog.empty()
            logbox.empty()
            results.sort(key=lambda r: -r['total'])
            ss['batch'] = dict(results=results, extras=ex_flags, vpp=vpp)
            if not results:
                st.error('所有股票均無法取得資料')
            elif ss.get('auto_track', True):
                recs = hit_records(results, K, get_live_combo_stats(ss, K))
                added = append_hit_log(recs)
                ss['msg'] = f'📌 命中組合 {len(recs)} 檔，新增 {added} 筆到追蹤紀錄'

    # ────────────────────────────── 歷史回測 ──────────────────────────────
    if run_bt:
        if not token:
            st.error('請先輸入 FinMind API Token')
        else:
            univ = {'上市清單': K['TWSE_LIST'], '上櫃清單': K['TPEX_LIST'],
                    '上市+上櫃': K['TWSE_LIST'] + K['TPEX_LIST'],
                    '目前輸入框清單': parse_stocks(ss['stocks_text'])}[bt_univ]
            # 先釋放上一段回測留在 session 的資料（原始紀錄、壓縮檔、分析快取），
            # 否則第2、3段跑的時候舊資料還佔著記憶體，Streamlit Cloud（約1GB）容易直接當掉
            release_bt_state(ss)
            prog = st.progress(0.0, text='🔬 歷史回測執行中...')
            det = st.empty()

            def onp(i, total, saved, failed, sid, note=None):
                prog.progress(min(1.0, i / max(total, 1)),
                              text=f'🔬 歷史回測執行中（{bt_seg if bt_seg != BT_SEG_OPTIONS[0] else f"過去{bt_months}個月"}）：{i} / {total}')
                det.caption((f'目前：{sid}　' if sid else '完成　') + f'已存 {saved:,} 筆　失敗 {failed} 檔'
                            + (f'\n\n{note}' if note else ''))
            df, stt = run_backtest(token, univ, int(bt_months), bt_div, bt_inst, bt_liq * 10000, vpp,
                                   delay=bt_delay, on_progress=onp, end_offset_months=bt_off)
            ss['bt_seg_label'] = '' if bt_seg == BT_SEG_OPTIONS[0] else bt_seg.split('・')[1][:3]
            prog.empty()
            ss['bt_df'] = store_bt_df(df)
            ss['bt_ver'] = ss.get('bt_ver', 0) + 1
            ss['bt_loaded_name'] = None
            det.caption(f"回測完成：{stt['saved']:,} 筆評估紀錄，失敗 {stt['failed']} 檔"
                        + (f"（途中因 FinMind 額度用完暫停約 {stt['quota_wait_min']} 分鐘）" if stt.get('quota_wait_min') else ''))
            if stt.get('stopped'):
                st.error(f"⚠️ 回測提前結束：{stt['stopped']}")
            if not len(df):
                st.error('❌ 這次回測沒有產生任何評估紀錄。失敗原因：' + ('；'.join(stt.get('top_fail') or []) or '不明')
                         + '。常見原因：Token 錯誤、FinMind 連線不到。')
            elif stt['failed'] > max(20, stt['total'] * 0.2):
                st.warning(f"⚠️ 有 {stt['failed']} 檔失敗（共 {stt['total']} 檔），結果可能不完整。主要原因：" + '；'.join(stt.get('top_fail') or []))

    tab_batch, tab_track, tab_bt = st.tabs(['📋 批次分析', '📈 命中追蹤', '🔬 歷史回測'])
    with tab_batch:
        render_batch(st, ss, K)
    with tab_track:
        render_tracking(st, ss, K)
    with tab_bt:
        render_backtest(st, ss, K)


# ────────────────────────────────────────────────────────────────────
def get_live_combo_stats(ss, K):
    """有回測資料時，算一次所有指定組合的即時統計並快取（回測資料換了才重算）"""
    df = ss.get('bt_df')
    if df is None or not len(df):
        return None
    sig = (ss.get('bt_ver'), id(df), len(df))
    if ss.get('live_stats_sig') != sig:
        combos = K['STOCK_PICK_COMBOS'] + K['PINNED_COMBOS'] + K['MOONSHOT_COMBOS']
        ss['live_stats'] = compute_live_combo_stats(add_derived(df), combos)
        ss['live_stats_sig'] = sig
    return ss['live_stats']


def render_batch(st, ss, K):
    B = ss.get('batch')
    if not B or not B['results']:
        st.info('在左側輸入 FinMind API Token 與股票代號，點擊「🔍 批次分析」即可開始。')
        return
    results, exf, vpp = B['results'], B['extras'], B['vpp']
    live = get_live_combo_stats(ss, K)
    if live:
        st.caption('🧮 命中組合的 t值／勝率／飆股比例：使用本次程式內的回測即時計算（t值為同日調整，括號內為5/10/20日中最高的t）。')
    else:
        st.caption('🧮 命中組合的 t值／勝率／飆股比例：尚未跑回測，先用回測記錄值（選股型 S 為 2023-10～2026-09 三年回測；沒進榜的組合沒有 t 值）；跑過回測或載入回測紀錄後會改用即時計算。')
    df = summary_frame(results, exf, K, live)
    _reg = batch_market_regime(results)
    if _reg:
        (st.warning if _reg == 2 else st.info)(regime_banner(_reg))

    with st.expander(f"「指定組合命中」編號對照：[選股型] S1~S{len(K['STOCK_PICK_COMBOS'])}　／　[高勝率] #1~#{len(K['PINNED_COMBOS'])}　／　[高標股] M1~M{len(K['MOONSHOT_COMBOS'])}"):
        st.caption('S＝選股型：2023-10～2026-09 三年回測分三段（偏多／空頭／多頭），每一段扣掉同一天大盤後都仍有超額報酬（t≥2），最值得看。'
                   '#＝高勝率；標 ⏱ 的是擇時型（含布林低檔／大盤跌破均線），勝率主要來自大盤反彈，適合判斷大盤落底，不適合當選股依據。'
                   '⚠️ #、M 的歷史統計是舊版（VAH/VAL 分價量表、夜星未限高檔）時期記錄的。')
        t0 = pd.DataFrame([{'編號': f'S{i + 1}', '條件組合': '＋'.join(c),
                            '回測記錄(三年，2026-09-30)': format_stockpick_stats(K['STOCK_PICK_STATS'].get(' ＋ '.join(c)))}
                           for i, c in enumerate(K['STOCK_PICK_COMBOS'])])
        st.dataframe(t0, hide_index=True, use_container_width=True)
        t1 = pd.DataFrame([{'編號': f'#{i + 1}', '類型': combo_type_label(c), '條件組合': '＋'.join(c),
                            '歷史勝率': format_winrates(K['PINNED_COMBO_WINRATES'].get(' ＋ '.join(c)))}
                           for i, c in enumerate(K['PINNED_COMBOS'])])
        t2 = pd.DataFrame([{'編號': f'M{i + 1}', '條件組合': '＋'.join(c),
                            '歷史飆股統計': format_moonshot_stats(K['MOONSHOT_COMBO_STATS'].get(' ＋ '.join(c)))}
                           for i, c in enumerate(K['MOONSHOT_COMBOS'])])
        st.dataframe(t1, hide_index=True, use_container_width=True, height=240)
        st.dataframe(t2, hide_index=True, use_container_width=True, height=240)
        st.caption(f"🔵 W＝回測高勝率（{len(K.get('BT_WIN_COMBOS', []))}組，青色）：三年回測任一天期勝率>55%，且三年各自依日 t≥2、六個半年至少五個為正；標 ⏱ 的是擇時型。"
                   f"　🔴 F＝回測飆股（{len(K.get('BT_HOT_COMBOS', []))}組，粉紅色）：10/20日飆股比例>10%、飆股次數≥10，且三段飆股比例都≥2倍基準。摘要表每類只顯示最強的{BT_HIT_SHOW}個，其餘以 +N 表示。"
                   "　🔢 命中類別數＝S/#/M/W/F 五類中命中幾類；同類組合常共用同一訊號，類別數比總命中數更能反映強度（回測：≥3類 台股20日超額約+2~4%、美股+0.6~1.8%，但波動也較大）。"
                   "　🏷️ 組合型態（依命中類別組合分級）：" + '　'.join(f"{COMBO_STYLE_ICON[k]}＝{'、'.join(c for c, v in COMBO_STYLE_RULES.items() if v == k)}（{COMBO_STYLE_NOTE[k]}）" for k in ('進攻', '穩健', '彩券')) + "；其他組合不標示。")
        tw_ = pd.DataFrame([{'編號': f'W{i + 1}', '類型': combo_type_label(c), '條件組合': '＋'.join(c),
                             '三年回測': '　'.join(f"{h}日 勝率{e['win']}% t={e['t']}" for h, e in sorted(K['BT_WIN_STATS'].get(' ＋ '.join(c), {}).items(), key=lambda kv: int(kv[0])))}
                            for i, c in enumerate(K.get('BT_WIN_COMBOS', []))])
        tf_ = pd.DataFrame([{'編號': f'F{i + 1}', '條件組合': '＋'.join(c),
                             '三年回測': format_moonshot_stats(K['BT_HOT_STATS'].get(' ＋ '.join(c)))}
                            for i, c in enumerate(K.get('BT_HOT_COMBOS', []))])
        st.dataframe(tw_, hide_index=True, use_container_width=True, height=240)
        st.dataframe(tf_, hide_index=True, use_container_width=True, height=240)

    c1, c2, c3, c5, c4 = st.columns([1.2, 1.4, 1.4, 1.1, 1])
    f_rg = c5.radio('大盤適用', ['全部', '有✅', '不含❌'], horizontal=True,
                    help='依今日大盤狀態，看命中的組合在這種大盤下的回測：有✅＝至少一個組合有效；不含❌＝排除命中「這種大盤下回測為負」組合的股票')
    f_pb = c1.radio('進場條件', ['全部', '✅ 符合進場', '❌ 不符合'], horizontal=True)
    f_sc = c2.radio('評分', ['全部', '80+', '65-79', '50-64', '<50'], horizontal=True)
    f_pt = c3.radio('型態確認', ['全部', '✅ 已突破', '🔥 剛突破', '🕒 成形中', '－ 無'], horizontal=True)
    kw = c4.text_input('搜尋代號/名稱')
    m = pd.Series(True, index=df.index)
    if f_pb != '全部':
        m &= df['_pbPass'] == (f_pb == '✅ 符合進場')
    sc = df['多方力道']
    m &= {'全部': True, '80+': sc >= 80, '65-79': (sc >= 65) & (sc < 80), '50-64': (sc >= 50) & (sc < 65),
          '<50': sc < 50}[f_sc]
    ptm = {'全部': None, '✅ 已突破': ['breakout', 'justbreak'], '🔥 剛突破': ['justbreak'],
           '🕒 成形中': ['forming'], '－ 無': ['none']}[f_pt]
    if ptm:
        m &= df['_pt'].isin(ptm)
    if kw:
        m &= df['股票'].str.contains(kw) | df['名稱'].str.contains(kw)
    if f_rg == '有✅' and '_regOk' in df.columns:
        m &= df['_regOk'] > 0
    elif f_rg == '不含❌' and '_regBad' in df.columns:
        m &= df['_regBad'] == 0
    view = df[m].drop(columns=['_pbPass', '_pt', '_regOk', '_regBad'], errors='ignore')
    st.caption(f'顯示 {len(view)} / {len(df)} 檔（點欄位標題可排序）')
    colcfg = {c: st.column_config.NumberColumn(format='%.1f') for c in
              ['+DI', '-DI', 'ADX', 'ADXR', '布林位置%', '布林寬度%', 'RS(vs0050)%', 'P/E', '營收YoY%(最新)',
               '營收MoM%(最新)', '均價YoY%(最新)', 'YoY乖離度(近3月合計)pp'] if c in view.columns}
    colcfg['量比'] = st.column_config.NumberColumn(format='%.2f')
    colcfg['漲跌幅%'] = st.column_config.NumberColumn(format='%.2f')
    colcfg['股價'] = st.column_config.NumberColumn(format='%.2f')
    colcfg['成交量'] = st.column_config.NumberColumn(format='%d')
    if '三大法人(近3月合計)' in view.columns:
        colcfg['三大法人(近3月合計)'] = st.column_config.NumberColumn(format='%d')
    # 命中欄分色：選股型S＝綠、指定組合#/M＝紫、回測高勝率W＝青、回測飆股F＝粉紅
    hit_colors = {'選股型命中': '#00a050', '指定組合命中': '#7c3aed', '回測高勝率命中': '#0097a7', '回測飆股命中': '#d81b60'}
    try:
        sty = view.style
        for c, colr in hit_colors.items():
            if c in view.columns:
                sty = sty.set_properties(subset=[c], **{'color': colr, 'font-weight': '600'})
        if '組合型態' in view.columns:
            _sc = {'進攻': 'color:#d84315;font-weight:700', '穩健': 'color:#2e7d32;font-weight:700', '彩券': 'color:#8e24aa;font-weight:700'}
            def _style_css(v):
                return next((c for k, c in _sc.items() if k in str(v)), '')
            sty = sty.map(_style_css, subset=['組合型態']) if hasattr(sty, 'map') else sty.applymap(_style_css, subset=['組合型態'])
        show_obj = sty
    except Exception:  # noqa  沒有 jinja2 等套件時退回不上色
        show_obj = view
    st.dataframe(show_obj, hide_index=True, use_container_width=True, column_config=colcfg,
                 height=min(560, 38 + 35 * max(len(view), 1)))
    b1, b2, b3 = st.columns([1, 1.4, 2])
    b1.download_button('📥 匯出 Excel', df_to_excel_bytes(view, '批次分析摘要'),
                       file_name=f'批次分析摘要_{dt.date.today()}.xlsx')
    recs = hit_records(results, K, live)
    b2.download_button(f'🎯 匯出命中組合股票（{len(recs)}檔）', hits_excel_bytes(recs),
                       file_name=f'命中組合股票_{dt.date.today()}.xlsx', disabled=not recs, key='dl_hits')
    if b3.button('📌 將命中股票加入追蹤紀錄', disabled=not recs, key='add_track'):
        n = append_hit_log(recs)
        st.success(f'已加入 {n} 筆新紀錄（同一天同一檔不重複）')

    st.divider()
    opts = [f"{r['stockId']} {r['name']}（{r['total']}分）" for r in results]
    pick = st.selectbox('🏷️ 個股詳細分析', range(len(results)), format_func=lambda i: opts[i])
    render_single(st, ss, results[pick], vpp)


def _tags(sigs):
    return ''.join(f"<span class='tagb {t}'>{'▲' if t == 'bull' else '▼' if t == 'bear' else '─'} {s}</span>" for s, t in sigs)


def render_single(st, ss, r, vpp):
    b, info = r['bars'], r['info']
    e = b.n - 1
    dm, pb, pt, sig = info['dm'], info['pb'], info['pt'], info['sig']
    total = r['total']
    last_c = b.close[e]
    prev_c = b.close[e - 1] if e >= 1 else last_c
    chg = last_c - prev_c
    chgp = chg / prev_c * 100 if prev_c else 0
    vr = b.volume[e] / b.vm20[e] if b.vm20[e] else 1
    st.subheader(f"🏷️ {r['stockId']} {r['name']}")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric('最新收盤', fnum(last_c, 2), f'{chg:+.2f} ({chgp:+.2f}%)', delta_color='inverse')
    m2.metric('當日成交量', f'{b.volume[e]:,.0f}')
    m3.metric('量比 vs MA20', f"{vr:.2f}x（{'放量' if vr > 1.2 else '縮量' if vr < 0.8 else '正常'}）")
    m4.metric('資料日期', b.date[e])

    st.markdown('#### 📊 多方力道評分')
    cA, cB = st.columns([1, 2])
    lbl = '強力買進訊號' if total >= 80 else '可考慮進場' if total >= 65 else '觀望為主' if total >= 50 else '不建議進場'
    col = '#00c864' if total >= 80 else '#2196f3' if total >= 65 else '#f0a500' if total >= 50 else '#ff3c3c'
    cA.markdown(f"<div style='font-size:56px;font-weight:700;color:{col};line-height:1'>{total}</div>"
                f"<div class='small'>綜合評分 / 100</div><div style='color:{col};font-weight:700'>{lbl}</div>",
                unsafe_allow_html=True)
    for t, s, mx, d in [('🧭 方向性 (+DI vs -DI)', dm['diPts'], 50, '+DI相對-DI的優勢程度'),
                        ('💪 趨勢強度 (ADX)', dm['adxPts'], 30, 'ADX值，越高趨勢越明確'),
                        ('🚀 趨勢動能 (ADX vs ADXR)', dm['adxrPts'], 20, 'ADX>ADXR代表趨勢正在轉強')]:
        cB.progress(min(1.0, s / mx), text=f'{t}　{s:.1f}/{mx}　（{d}）')
    st.markdown('#### 🔔 技術訊號')
    st.markdown(_tags(dm['sigs']), unsafe_allow_html=True)

    st.markdown('#### 💡 操作建議')
    ma20v = b.ma20[e] or last_c
    bbUp = b.bbU[e] or last_c * 1.1
    name = f"{r['stockId']} {r['name']}"
    if total >= 80:
        act = '🟢 積極做多'
        adv = [f'**{name}** 多方力道評分 {total} 分，DMI顯示多方力道強勁，建議積極做多。']
        if dm['tdir'] == '多頭':
            adv.append('+DI大於-DI且趨勢轉強，順勢操作，逢低分批佈局。')
        sl, tg = f'停損設於 **{ma20v * 0.97:.1f}** 元（20MA下方3%）', f'目標參考 **{bbUp:.1f}** 元（布林上軌）'
    elif total >= 65:
        act = '🔵 可考慮進場'
        adv = [f'**{name}** 評分 {total} 分，DMI偏多，可考慮分批進場。', '建議等待回測均線後再進場，降低風險。']
        sl, tg = f'停損建議 **{ma20v * 0.98:.1f}** 元（20MA下方2%）', f'短線目標 **{last_c * 1.08:.1f}** 元（+8%）'
    elif total >= 50:
        act = '🟡 觀望為主'
        adv = [f'**{name}** 評分 {total} 分，DMI訊號混雜或趨勢不明確，建議觀望。', '等待ADX轉強或+DI/-DI方向明確後再行動。']
        sl, tg = '暫不建議進場', '等待更佳時機'
    else:
        act = '🔴 不適合進場'
        adv = [f'**{name}** 評分 {total} 分，DMI偏空，不建議進場。',
               '目前-DI大於+DI，空方力道較強，切忌逆勢做多，等待趨勢反轉。' if dm['tdir'] == '空頭' else 'DMI指標偏弱或資料不足，應持現金等待機會。']
        sl, tg = '持倉者建議設停損出場', '等待多頭訊號出現'
    hints = [s[0] for s in dm['sigs'] if s[1] == 'bull'][:5]
    import re as _re

    def _b(t):
        return _re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', t)
    html = (f"<div style='background:rgba(33,150,243,.08);border:1px solid rgba(33,150,243,.3);border-radius:10px;"
            f"padding:12px 18px;line-height:1.9'><div style='font-size:17px;font-weight:700'>{act}</div>"
            + '<br>'.join(_b(a) for a in adv)
            + f"<div style='margin-top:6px'>🛑 <b>停損：</b>{_b(sl)}<br>🎯 <b>目標：</b>{_b(tg)}"
            + (f"<br>✅ <b>多頭訊號：</b>{' · '.join(hints)}" if hints else '') + '</div></div>')
    st.markdown(html, unsafe_allow_html=True)

    akey = r['stockId'] + b.date[e]
    if st.button('🤖 AI 智能綜合分析', key='ai_' + akey):
        k = ss.get('openai_key', '').strip()
        if not k:
            ss['ai'][akey] = '請先在左側輸入 OpenAI API Key，才能使用 AI 智能綜合分析。'
        else:
            with st.spinner('正在請 AI 綜合研判技術面數據，請稍候…'):
                try:
                    ss['ai'][akey] = run_ai(build_ai_prompt(r), k, ss.get('openai_model', 'gpt-4o-mini'))
                except Exception as ex:  # noqa
                    ss['ai'][akey] = f'❌ AI 分析失敗：{ex}\n請確認 API Key 是否正確、額度是否足夠，或稍後再試。'
    if akey in ss['ai']:
        st.markdown(ss['ai'][akey])

    cL, cR = st.columns(2)
    with cL:
        st.markdown('#### 🎯 回後買上漲 · 進場條件核對')
        if pb['allPass']:
            st.success('✅ 符合進場條件（必要條件全部通過）' + ('  +成交量增加分' if pb['bonusPassed'] else ''))
        else:
            st.error(f"❌ 不符合進場條件（必要條件 {pb['requiredPassed']}/{pb['requiredTotal']} 通過）")
        for x in pb['results']:
            icon = '✅' if x['pass_'] else ('❌' if x['required'] else '—')
            st.markdown(f"{icon} {x['label']}" + ('' if x['required'] else '（加分）')
                        + (f"　<span class='small'>({x['detail']})</span>" if x['detail'] else ''), unsafe_allow_html=True)
    with cR:
        st.markdown('#### 🔍 型態確認（15種進場型態）')
        if pt['anyJustBroke']:
            st.warning('🔥 偵測到剛突破買點（較前一交易日新增）')
        elif pt['anyBreakout']:
            st.success('✅ 偵測到型態突破買點')
        elif pt['anyFormed']:
            st.info('🕒 型態成形中，尚未突破確認')
        else:
            st.caption('－ 目前未偵測到符合的進場型態')
        hc_hit = [x['name'] for x in pt['results'] if x['id'] in HIGH_CONSOLIDATION_PATTERNS and x['formed']]
        if hc_hit:
            st.info('📈 高檔強勢整理：' + '、'.join(hc_hit) + '（回測顯示強勢股出現多為洗盤後續漲，留意是否守住首日低點）')
        for x in pt['results']:
            icon = '🔥' if x['justBroke'] else ('✅' if x['breakout'] else ('🕒' if x['formed'] else '－'))
            st.markdown(f"{icon} **{x['name']}**" + ('（高檔強勢整理）' if x['id'] in HIGH_CONSOLIDATION_PATTERNS else '')
                        + f"<br><span class='small'>{x['desc']}" + (f"<br>{x['detail']}" if x['formed'] else '') + '</span>',
                        unsafe_allow_html=True)

    st.markdown('#### 📐 指標對照')
    g1, g2, g3, g4, g5 = st.columns(5)
    rv = b.rsi[e]
    g1.metric('RSI', fnum(rv))
    g1.caption('⚠️ 超買區（>70）' if (rv or 0) > 70 else '💚 超賣區（<30）' if (rv or 100) < 30 else '正常區間（30-70）')
    if b.kdK[e] is not None:
        g2.metric('KD', f"K={b.kdK[e]:.1f} D={b.kdD[e]:.1f}")
        g2.caption('K>D 偏多' if b.kdK[e] > b.kdD[e] else 'K<D 偏空')
    cross = ''
    if e >= 1:
        if b.macd[e - 1] <= b.macdSig[e - 1] and b.macd[e] > b.macdSig[e]:
            cross = '⚡黃金交叉'
        elif b.macd[e - 1] >= b.macdSig[e - 1] and b.macd[e] < b.macdSig[e]:
            cross = '⚡死亡交叉'
    g3.metric('MACD', f"DIF={b.macd[e]:.2f}")
    g3.caption(f"Signal={b.macdSig[e]:.2f}　柱={b.macdHist[e]:.2f}　"
               + ('DIF在Signal上方，偏多' if b.macd[e] > b.macdSig[e] else 'DIF在Signal下方，偏空') + (' ' + cross if cross else ''))
    if sig['bbPos'] is not None:
        bw = (b.bbU[e] - b.bbL[e]) / last_c * 100
        g4.metric('布林位置', f"{sig['bbPos']:.0f}%")
        g4.caption(f'上軌 {b.bbU[e]:.2f}　下軌 {b.bbL[e]:.2f}　寬度 {bw:.1f}%' + ('　🔸收窄，留意變盤' if bw < 8 else ''))
    prof, vps = info['prof'], info['vps']
    if prof:
        g5.metric('分價量表 POC', f"{prof['pocPrice']:.1f}")
        g5.caption(f"POC區 {prof['pocLow']:.1f}~{prof['pocHigh']:.1f}　"
                   + (vps['detail'].split('：')[0] if vps['detail'] else '無訊號'))
    ma_rows = [{'均線': nm, '數值': round(arr[e], 2), '股價偏離%': round((last_c - arr[e]) / arr[e] * 100, 2)}
               for nm, arr in (('MA5', b.ma5), ('MA10', b.ma10), ('MA20', b.ma20), ('MA60', b.ma60)) if arr[e] is not None]
    if vps['detail']:
        st.caption('分價量表訊號：' + vps['detail'])
    st.dataframe(pd.DataFrame(ma_rows), hide_index=True)

    st.markdown('#### 📉 技術分析圖表')
    st.plotly_chart(draw_chart(r, vpp), use_container_width=True, theme=None)
    with st.expander('📋 原始資料（最近20筆）'):
        s = max(0, e - 19)
        raw = pd.DataFrame({'日期': b.date[s:], '開盤': b.open[s:], '最高': b.high[s:], '最低': b.low[s:],
                            '收盤': b.close[s:], '成交量': b.volume[s:], 'MA5': b.ma5[s:], 'MA20': b.ma20[s:],
                            'MA60': b.ma60[s:], 'RSI': b.rsi[s:], 'KD-K': b.kdK[s:], 'KD-D': b.kdD[s:],
                            'MACD': b.macd[s:], 'Signal': b.macdSig[s:], 'BB上軌': b.bbU[s:], 'BB下軌': b.bbL[s:]})
        st.dataframe(raw.iloc[::-1].round(2), hide_index=True, use_container_width=True)



# ────────────────────────────────────────────────────────────────────
def render_tracking(st, ss, K):
    log = load_hit_log()
    st.subheader('📈 命中組合每日追蹤')
    st.caption(f'追蹤紀錄檔：{HIT_LOG_FILE}　｜　每次批次分析自動累加（可在側邊欄關閉），同一天同一檔只記一次。'
               '進場價＝命中當天收盤價；「5/10/20日報酬」要等滿天數才會出現。')
    if not len(log):
        st.info('目前沒有追蹤紀錄。跑一次批次分析（建議用「漲幅前100／成交量前100」），命中組合的股票會自動加入。')
        return
    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1])
    c1.metric('追蹤紀錄', f'{len(log)} 筆')
    c2.metric('股票數', f"{log['股票代號'].nunique()} 檔")
    c3.metric('期間', f"{log['日期'].min()} ~ {log['日期'].max()}")
    if c4.button('🔄 更新追蹤報酬', type='primary', key='upd_track'):
        tok = ss.get('token', '').strip()
        if not tok:
            st.error('請先輸入 FinMind API Token')
        else:
            prog = st.progress(0.0, text='抓取最新股價中...')
            ss['track_perf'] = track_performance(log, tok, delay=0.3,
                                                 progress=lambda i, n: prog.progress(i / n, text=f'抓取最新股價：{i}/{n}'))
            prog.empty()
    perf = ss.get('track_perf')
    days = sorted(log['日期'].unique(), reverse=True)
    pick = st.selectbox('檢視日期', ['全部'] + days, key='track_day')
    if perf is None or not len(perf):
        view = log if pick == '全部' else log[log['日期'] == pick]
        st.caption('尚未更新報酬，先顯示原始命中紀錄。按「🔄 更新追蹤報酬」抓最新股價。')
        st.dataframe(view.drop(columns=['命中組合']), hide_index=True, use_container_width=True)
    else:
        view = perf if pick == '全部' else perf[perf['日期'] == pick]
        st.markdown('##### 🧾 追蹤明細')
        st.dataframe(view, hide_index=True, use_container_width=True, height=380)
        st.markdown('##### 🧩 組合實盤表現（命中後實際報酬 vs 回測記錄）')
        summ = combo_track_summary(perf, K)
        st.caption('實盤勝率明顯低於回測記錄的組合，代表可能是過度適配，可考慮從指定組合移除；樣本少於10筆前先別下結論。')
        st.dataframe(summ, hide_index=True, use_container_width=True, height=380)
        st.download_button('📥 匯出追蹤報表', sheets_to_xlsx([('命中紀錄', log), ('追蹤明細', perf), ('組合彙總', summ)]), file_name=f'命中追蹤_{dt.date.today()}.xlsx', key='dl_track')
    with st.expander('🗂️ 管理追蹤紀錄'):
        keep_days = st.number_input('只保留最近幾天的紀錄', 5, 3650, 120, 5, key='keep_days')
        if st.button('🧹 刪除更早的紀錄', key='prune_track'):
            cut = str(dt.date.today() - dt.timedelta(days=int(keep_days)))
            log2 = log[log['日期'] >= cut]
            log2.to_csv(HIT_LOG_FILE, index=False, encoding='utf-8-sig')
            ss['track_perf'] = None
            st.success(f'已刪除 {len(log) - len(log2)} 筆')
        st.code(DAILY_USAGE, language='text')

# ────────────────────────────────────────────────────────────────────
def render_backtest(st, ss, K):
    df = ss.get('bt_df')
    if df is None or not len(df):
        st.info('點擊左側「🔬 執行歷史回測」，或載入先前匯出的回測原始紀錄。')
        return
    df = add_derived(df)
    d0, d1 = str(df['evalDate'].min()), str(df['evalDate'].max())
    st.subheader('🔬 歷史回測分析結果')
    st.caption(f"共 {len(df):,} 筆評估紀錄（{d0} ～ {d1}），{df['stockId'].nunique()} 檔股票")
    if len(df) > 450_000:
        st.warning(f'⚠️ 資料量 {len(df):,} 筆偏大，Streamlit Cloud（記憶體約1GB）可能跑不動而當掉。'
                   '可以先只載入其中兩段，或回測時提高流動性門檻／最低股價減少筆數。')
    # 原始紀錄檔按了才產生（全美股約20萬筆，每次重跑都先壓一次檔會多花十幾秒和幾百MB記憶體）
    if ss.get('_bt_csv_sig') == ss.get('bt_ver') and ss.get('_bt_csv'):
        st.download_button('💾 下載回測原始紀錄（.csv.gz，之後可直接載入，不用重抓）', ss['_bt_csv'],
                           file_name=f'回測原始紀錄{("_" + ss["bt_seg_label"]) if ss.get("bt_seg_label") else ""}_{dt.date.today()}.csv.gz', key='dl_btcsv')
    elif st.button('💾 產生回測原始紀錄檔（之後可直接載入，不用重抓）', key='mk_btcsv'):
        with st.spinner('壓縮中…'):
            ss['_bt_csv'] = bt_csv_gz(df)
            ss['_bt_csv_sig'] = ss.get('bt_ver')
        st.rerun()

    def show(title, t, note=None):
        st.markdown(f'##### {title}')
        if note:
            st.caption(note)
        if t is None or not len(t):
            st.caption('（無資料）')
        else:
            st.dataframe(t, hide_index=True, use_container_width=True)

    show('📏 全體基準（所有評估紀錄）', baseline_row(df), '以下各訊號都跟這一列比，高於基準才代表訊號有用。')
    show('📊 評分區間 vs 實際報酬', score_buckets(df))
    show('🧮 訊號堆疊數量 vs 實際報酬（Confluence Count）', confluence_table(df),
         '同一時間點同時觸發的獨立訊號數（型態剛形成、KDJ交叉、MACD黃金交叉、布林極端位置、爆量，最多5個）。')
    for title, fld, lbl in [('🚀 「強勢突破盤」', 'isBreakout', '強勢突破盤'), ('🎯 「跌深反彈盤」', 'isPullbackRebound', '跌深反彈盤'),
                            ('⚡ 「MACD近3日內黃金交叉」', 'goldenCrossRecent', 'MACD黃金交叉'),
                            ('🟢 「KDJ近3日內黃金交叉」', 'kdjGoldenCrossRecent', 'KDJ黃金交叉'),
                            ('🔴 「KDJ近3日內死亡交叉」', 'kdjDeathCrossRecent', 'KDJ死亡交叉'),
                            ('💪 「相對強弱(vs大盤0050，20日)」', 'relStrengthPositive', '相對強弱為正'),
                            ('📊 「爆量(≥1.5倍均量)」', 'volSurge15', '爆量1.5倍'), ('📊 「爆量(≥2倍均量)」', 'volSurge2', '爆量2倍'),
                            ('📉 「地量(≤0.5倍均量)」', 'volLow05', '地量'),
                            ('📈 「量能區間高檔(≥90百分位)」', 'volRangeHigh90', '量能區間高檔'),
                            ('📉 「量能區間低檔(≤10百分位)」', 'volRangeLow10', '量能區間低檔')]:
        t = tag_hitrate(df, fld, lbl)
        if len(t):
            show(title + '標記 vs 實際報酬', t)
    if 'vpBuySupport' in df:
        st.markdown('##### 📊 個股分價量表（POC）訊號 vs 實際報酬')
        st.caption('以成交量最大價位區（POC）為界：守穩POC買進、突破POC追價買進、反彈POC遇壓賣出、破位停損賣出（帶量長黑跌破POC）。')
        st.dataframe(pd.concat([tag_hitrate(df, f, l) for f, l in VP_LABELS.items()], ignore_index=True),
                     hide_index=True, use_container_width=True)
    for title, fld, lbl in [('🌐 「大盤站上20日均線」', 'benchmarkAbove20', '大盤站上20日均線'),
                            ('🌐 「大盤站上60日均線」', 'benchmarkAbove60', '大盤站上60日均線'),
                            ('🔍 「型態突破確認」', 'patternBreakout', '型態突破確認'),
                            ('🔥 「型態剛形成(剛突破)」', 'patternJustBroke', '型態剛形成'),
                            ('✅ 「回後買上漲全通過」', 'pbAllPass', '回後買上漲全通過')]:
        t = tag_hitrate(df, fld, lbl)
        if len(t):
            show(title + '標記 vs 實際報酬', t)
    show('📐 15種型態各自「剛形成」vs 實際報酬', bt_memo(ss, 'pat', (), lambda: pattern_hits(df)), '依樣本數排序，樣本數<10筆的不列出。')
    show('🆕 新增技術面條件 vs 實際報酬（52週高點／均線多頭排列／布林收窄／向上跳空）',
         bt_memo(ss, 'newtech', (), lambda: new_tech_table(df)),
         '52週高點需要約一年的歷史，回測會自動多抓資料；資料不足的評估點不列入「是／否」。')

    st.markdown('##### 🎛️ 參數網格搜尋（單一評分公式的權重調整）')
    gh = st.selectbox('優化目標天數', HORIZONS, index=1, key='gh', format_func=lambda h: f'{h}日')
    st.caption('⚠️ 在已收集的歷史資料上找「表現較好」的參數組合，樣本有限時容易過度適配，僅供方向參考。')
    st.dataframe(bt_memo(ss, 'grid', (gh,), lambda: grid_search(df, gh)), hide_index=True, use_container_width=True)

    # ── 多因子複選搜尋 ──
    st.markdown('##### 🧩 多因子複選搜尋（找出哪幾項欄位組合起來勝率最高）')
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    ch = c1.selectbox('優化目標天數', HORIZONS, index=1, key='ch', format_func=lambda h: f'{h}日')
    cwr = c2.number_input('勝率門檻(%)', 0.0, 100.0, 50.0, 5.0, key='cwr')
    cms = c3.number_input('最少樣本數', 5, 5000, 50, 5, key='cms')
    csort = c4.selectbox('排序依據', ['t值(同日調整)', '勝率'], key='csort',
                         help='t值(同日調整)：優先列出扣掉同一天大盤後仍然比較強的「選股型」組合；勝率：舊排序，容易被擇時型組合佔滿')
    cmt = c5.number_input('t值門檻', -10.0, 20.0, 1.0, 0.5, key='cmt',
                          help='排除 t值(同日調整) 低於門檻的組合（t<1：扣掉同一天大盤後幾乎沒有超額報酬，勝率多半只是跟著大盤）')
    cred = c6.checkbox('排除冗餘組合', True, key='cred', help='多加一個條件後樣本完全沒變（例如「母子懷抱剛形成」必然也是「型態剛形成」），這種組合不重複列出')
    if not bt_memo_ready(ss, 'combo', (ch, int(cms), float(cwr), cred, csort)) and not st.button('🧩 開始多因子搜尋', key='run_combo', type='primary'):
        st.info(f'回測資料共 {len(df):,} 筆；多因子搜尋需要 1～3 分鐘、也比較吃記憶體，按上方按鈕才開始計算（同一組參數算過一次就會記住，改參數要再按一次）。')
    else:
        with st.spinner('多因子複選搜尋計算中…'):
            res, tested, base = bt_memo(ss, 'combo', (ch, int(cms), float(cwr), cred, csort),
                                        lambda: combo_search(df, ch, 3, int(cms), float(cwr), cred, 't' if csort.startswith('t') else 'win'))
        n_before_t = len(res)
        if len(res) and 't值(同日調整)' in res.columns:
            res = res[res['t值(同日調整)'] >= float(cmt)].reset_index(drop=True)
        cboth = st.checkbox('只列各段 t值都 ≥ 2 的組合（✅ 各段一致）', False, key='cboth',
                            help='回測期間依評估日切段（資料超過18個月切三段，否則前後兩半），各段各自算同日調整 t 值。'
                                 '每段都 ≥2 代表不是單一段行情造成的巧合，最值得相信')
        if cboth and len(res) and '各段一致' in res.columns:
            res = res[res['各段一致'] == '✅'].reset_index(drop=True)
        if len(res):
            res.insert(1, '類型', res['條件組合'].map(lambda s: combo_type_label([x.strip() for x in s.split('＋')])))
        if base:
            st.caption(f"全體基準（{ch}日）：樣本 {base['n']:,}　平均報酬 {base['avg']:.2f}%　勝率 {base['win']:.1f}%　中位數 {base['median']:.2f}%　"
                       f"｜共測試 {tested:,} 種組合，勝率≥{cwr:.0f}% 的 {n_before_t:,} 組，再排除 t值<{cmt:g} 後剩 {len(res):,} 組"
                       + (f"｜{base['parts']}（✅＝每段 t 都 ≥2，⚠️k/n＝n段中有k段達標）。" if base.get('parts') else '。')
                       + '⚠️ 測試組合越多，純運氣突出的也越多。「同日超額報酬」＝每筆報酬扣掉同一天全部紀錄的平均，t值也用它算，已排除大盤齊漲齊跌；t>2 較可信。')
        st.dataframe(res, hide_index=True, use_container_width=True, height=420)
        if len(res):
            st.download_button('📥 匯出 Excel', df_to_excel_bytes(res, '多因子複選搜尋'),
                               file_name=f'多因子複選搜尋_勝率{cwr:.0f}%以上_t{cmt:g}以上_{dt.date.today()}.xlsx', key='dl_combo')


    # ── 飆股搜尋 ──
    st.markdown('##### 🚀 飆股搜尋（找出最容易出現大行情的組合）')
    d1, d2, d3, d4, d5 = st.columns(5)
    mh = d1.selectbox('天數', HORIZONS, index=1, key='mh', format_func=lambda h: f'{h}日')
    mthr = d2.number_input('漲幅門檻(%)', 5.0, 200.0, 30.0, 5.0, key='mthr')
    mpct = d3.number_input('飆股比例門檻(%)', 0.0, 100.0, 10.0, 1.0, key='mpct')
    mms = d4.number_input('最少樣本數', 5, 5000, 20, 5, key='mms')
    mred = d5.checkbox('排除冗餘組合', True, key='mred')
    if not bt_memo_ready(ss, 'moon', (mh, float(mthr), int(mms), float(mpct), mred)) and not st.button('🚀 開始飆股搜尋', key='run_moon', type='primary'):
        st.info(f'回測資料共 {len(df):,} 筆；飆股搜尋需要 1～3 分鐘、也比較吃記憶體，按上方按鈕才開始計算（同一組參數算過一次就會記住，改參數要再按一次）。')
    else:
        with st.spinner('飆股搜尋計算中…'):
            mres, mtested, mbase = bt_memo(ss, 'moon', (mh, float(mthr), int(mms), float(mpct), mred),
                                           lambda: moonshot_search(df, mh, float(mthr), 3, int(mms), float(mpct), mred))
        if mbase:
            st.caption(f"全體基準（{mh}日漲幅>{mthr:.0f}%）：{mbase['n']:,} 筆中 {mbase['moonN']:,} 次飆股，"
                       f"基準飆股比例 {mbase['pct']:.2f}%　｜共測試 {mtested:,} 種組合，達標 {len(mres):,} 組"
                       + (f"｜{mbase['parts']}，每段飆股比例都高才可信。" if mbase.get('parts') else '。')
                       + '「倍數」＝組合飆股比例÷基準；飆股是稀有事件，飆股次數只有個位數的不建議當真。')
        st.dataframe(mres, hide_index=True, use_container_width=True, height=420)
        if len(mres):
            st.download_button('📥 匯出 Excel', df_to_excel_bytes(mres, '飆股搜尋'),
                               file_name=f'飆股搜尋_比例{mpct:.0f}%以上_{dt.date.today()}.xlsx', key='dl_moon')



# ════════════════════════════════════════════════════════════════════
#  內建清單與指定追蹤組合（由 HTML 版原封不動搬過來）
# ════════════════════════════════════════════════════════════════════

# 上市／上櫃清單（同時當作「漲幅/成交量前100」的個股白名單，排除ETF、權證等）

TWSE_LIST = [
    "1101", "1101B", "1102", "1103", "1104", "1108", "1109", "1110", "1201", "1203", "1210", "1213", "1215", "1216", "1217", "1218", "1219", "1220", "1225", "1227",
    "1229", "1231", "1232", "1233", "1234", "1235", "1236", "1256", "1301", "1303", "1304", "1305", "1307", "1308", "1309", "1310", "1312", "1312A", "1313", "1314",
    "1315", "1316", "1319", "1321", "1323", "1324", "1325", "1326", "1337", "1338", "1339", "1340", "1341", "1342", "1402", "1409", "1410", "1413", "1414", "1416",
    "1417", "1418", "1419", "1423", "1432", "1434", "1435", "1436", "1437", "1438", "1439", "1440", "1441", "1442", "1443", "1444", "1445", "1446", "1447", "1449",
    "1451", "1452", "1453", "1454", "1455", "1456", "1457", "1459", "1460", "1463", "1464", "1465", "1466", "1467", "1468", "1470", "1471", "1472", "1473", "1474",
    "1475", "1476", "1477", "1503", "1504", "1506", "1512", "1513", "1514", "1515", "1516", "1517", "1519", "1521", "1522", "1522A", "1524", "1525", "1526", "1527",
    "1528", "1529", "1530", "1531", "1532", "1533", "1535", "1536", "1537", "1538", "1539", "1540", "1541", "1558", "1560", "1563", "1568", "1582", "1583", "1587",
    "1590", "1597", "1598", "1603", "1604", "1605", "1608", "1609", "1611", "1612", "1614", "1615", "1616", "1617", "1618", "1623", "1626", "1702", "1707", "1708",
    "1709", "1710", "1711", "1712", "1713", "1714", "1717", "1718", "1720", "1721", "1722", "1723", "1725", "1726", "1727", "1730", "1731", "1732", "1733", "1734",
    "1735", "1736", "1737", "1752", "1760", "1762", "1773", "1776", "1783", "1786", "1789", "1795", "1802", "1805", "1806", "1808", "1809", "1810", "1817", "1903",
    "1904", "1905", "1906", "1907", "1909", "2002", "2002A", "2006", "2007", "2008", "2009", "2010", "2012", "2013", "2014", "2015", "2017", "2020", "2022", "2023",
    "2024", "2025", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2038", "2049", "2059", "2062", "2069", "2072", "2101", "2102", "2103", "2104",
    "2105", "2106", "2107", "2108", "2109", "2114", "2115", "2201", "2204", "2206", "2207", "2208", "2211", "2227", "2228", "2231", "2233", "2236", "2239", "2241",
    "2243", "2247", "2248", "2250", "2254", "2258", "2301", "2302", "2303", "2305", "2308", "2312", "2313", "2314", "2316", "2317", "2321", "2323", "2324", "2327",
    "2328", "2329", "2330", "2331", "2332", "2337", "2338", "2340", "2342", "2344", "2345", "2347", "2348", "2348A", "2349", "2351", "2352", "2353", "2354", "2355",
    "2356", "2357", "2359", "2360", "2362", "2363", "2364", "2365", "2367", "2368", "2369", "2371", "2373", "2374", "2375", "2376", "2377", "2379", "2382", "2383",
    "2385", "2387", "2388", "2390", "2392", "2393", "2395", "2397", "2399", "2401", "2402", "2404", "2405", "2406", "2408", "2409", "2412", "2413", "2414", "2415",
    "2417", "2419", "2420", "2421", "2423", "2424", "2425", "2426", "2427", "2428", "2429", "2430", "2431", "2432", "2433", "2434", "2436", "2438", "2439", "2440",
    "2441", "2442", "2444", "2449", "2450", "2451", "2453", "2454", "2455", "2457", "2458", "2459", "2460", "2461", "2462", "2464", "2465", "2466", "2467", "2468",
    "2471", "2472", "2474", "2476", "2477", "2478", "2480", "2481", "2482", "2483", "2484", "2485", "2486", "2488", "2489", "2491", "2492", "2493", "2495", "2496",
    "2497", "2498", "2501", "2504", "2505", "2506", "2509", "2511", "2514", "2515", "2516", "2520", "2524", "2527", "2528", "2530", "2534", "2535", "2536", "2537",
    "2538", "2539", "2540", "2542", "2543", "2545", "2546", "2547", "2548", "2597", "2601", "2603", "2605", "2606", "2607", "2608", "2609", "2610", "2611", "2612",
    "2613", "2614", "2615", "2616", "2617", "2618", "2630", "2633", "2634", "2636", "2637", "2642", "2645", "2646", "2701", "2702", "2704", "2705", "2706", "2707",
    "2712", "2722", "2723", "2727", "2731", "2739", "2748", "2753", "2762", "2801", "2812", "2816", "2820", "2832", "2834", "2836", "2836A", "2838", "2838A", "2845",
    "2849", "2850", "2851", "2852", "2855", "2867", "2880", "2881", "2881A", "2881B", "2881C", "2882", "2882A", "2882B", "2883", "2883B", "2884", "2885", "2886", "2887",
    "2887E", "2887F", "2887G", "2887H", "2887I", "2887Z1", "2889", "2890", "2891", "2891B", "2891C", "2892", "2897", "2897B", "2901", "2903", "2904", "2905", "2906", "2908",
    "2910", "2911", "2912", "2913", "2915", "2923", "2929", "2939", "2945", "3002", "3003", "3004", "3005", "3006", "3008", "3010", "3011", "3013", "3014", "3015",
    "3016", "3017", "3018", "3019", "3021", "3022", "3023", "3024", "3025", "3026", "3027", "3028", "3029", "3030", "3031", "3032", "3033", "3034", "3035", "3036",
    "3037", "3038", "3040", "3041", "3042", "3043", "3044", "3045", "3046", "3047", "3048", "3049", "3050", "3051", "3052", "3054", "3055", "3056", "3057", "3058",
    "3059", "3060", "3062", "3090", "3092", "3094", "3130", "3135", "3138", "3149", "3150", "3164", "3167", "3168", "3189", "3209", "3229", "3231", "3257", "3266",
    "3296", "3305", "3308", "3311", "3312", "3321", "3338", "3346", "3356", "3376", "3380", "3406", "3413", "3416", "3419", "3432", "3437", "3443", "3447", "3450",
    "3481", "3494", "3501", "3504", "3515", "3518", "3528", "3530", "3532", "3533", "3535", "3543", "3545", "3550", "3557", "3563", "3576", "3583", "3588", "3591",
    "3592", "3593", "3596", "3605", "3607", "3617", "3622", "3645", "3652", "3653", "3661", "3665", "3669", "3673", "3679", "3686", "3694", "3701", "3702", "3703",
    "3704", "3705", "3706", "3708", "3711", "3712", "3714", "3715", "3716", "3717", "4104", "4106", "4108", "4119", "4133", "4137", "4142", "4148", "4155", "4164",
    "4169", "4178", "4190", "4195", "4306", "4414", "4426", "4438", "4439", "4440", "4441", "4526", "4532", "4536", "4540", "4545", "4551", "4552", "4555", "4557",
    "4560", "4562", "4564", "4566", "4569", "4571", "4572", "4576", "4581", "4582", "4583", "4585", "4588", "4590", "4720", "4722", "4736", "4737", "4739", "4746",
    "4755", "4763", "4764", "4766", "4770", "4771", "4807", "4904", "4906", "4912", "4915", "4916", "4919", "4927", "4930", "4934", "4935", "4938", "4942", "4943",
    "4949", "4952", "4956", "4958", "4960", "4961", "4967", "4968", "4976", "4977", "4989", "4994", "4999", "5007", "5203", "5215", "5222", "5225", "5234", "5243",
    "5244", "5258", "5269", "5283", "5284", "5285", "5288", "5292", "5306", "5388", "5434", "5469", "5471", "5484", "5515", "5519", "5521", "5522", "5525", "5531",
    "5533", "5534", "5538", "5546", "5607", "5608", "5706", "5871", "5871A", "5876", "5880", "5906", "5907", "6005", "6024", "6108", "6112", "6115", "6116", "6117",
    "6120", "6128", "6133", "6136", "6139", "6141", "6142", "6152", "6153", "6155", "6164", "6165", "6166", "6168", "6176", "6177", "6183", "6184", "6189", "6191",
    "6192", "6196", "6197", "6201", "6202", "6205", "6206", "6209", "6213", "6214", "6215", "6216", "6224", "6225", "6226", "6230", "6235", "6239", "6243", "6257",
    "6269", "6271", "6272", "6277", "6278", "6281", "6282", "6283", "6285", "6405", "6409", "6412", "6414", "6415", "6416", "6426", "6431", "6438", "6442", "6443",
    "6446", "6449", "6451", "6456", "6464", "6472", "6477", "6491", "6504", "6505", "6515", "6525", "6526", "6531", "6533", "6534", "6541", "6550", "6552", "6558",
    "6573", "6579", "6581", "6582", "6585", "6589", "6591", "6592", "6592A", "6592B", "6598", "6605", "6606", "6614", "6625", "6641", "6645", "6655", "6657", "6658",
    "6666", "6668", "6669", "6670", "6671", "6672", "6674", "6689", "6691", "6695", "6698", "6706", "6715", "6719", "6722", "6742", "6743", "6753", "6754", "6756",
    "6757", "6768", "6770", "6771", "6776", "6781", "6782", "6789", "6790", "6792", "6794", "6796", "6799", "6805", "6807", "6830", "6831", "6834", "6835", "6838",
    "6854", "6861", "6862", "6863", "6869", "6873", "6885", "6887", "6890", "6901", "6902", "6906", "6908", "6909", "6914", "6916", "6918", "6919", "6921", "6923",
    "6924", "6928", "6931", "6933", "6934", "6936", "6937", "6944", "6949", "6951", "6952", "6955", "6957", "6958", "6958A", "6962", "6965", "6969", "6988", "6994",
    "7610", "7631", "7705", "7711", "7721", "7722", "7730", "7732", "7736", "7740", "7749", "7750", "7760", "7765", "7768", "7769", "7780", "7786", "7788", "7791",
    "7795", "7799", "7803", "7818", "7821", "7822", "7823", "7827", "8011", "8016", "8021", "8028", "8033", "8039", "8045", "8046", "8070", "8072", "8081", "8101",
    "8103", "8104", "8105", "8110", "8112", "8112A", "8114", "8131", "8150", "8162", "8163", "8201", "8210", "8213", "8215", "8222", "8249", "8261", "8271", "8341",
    "8367", "8374", "8404", "8411", "8422", "8429", "8438", "8442", "8443", "8454", "8462", "8463", "8464", "8466", "8467", "8473", "8476", "8478", "8481", "8482",
    "8487", "8488", "8499", "8926", "8940", "8996", "9103", "910322", "9105", "910861", "9110", "911608", "911622", "911868", "912000", "9136", "9802", "9902", "9904", "9905",
    "9906", "9907", "9908", "9910", "9911", "9912", "9914", "9917", "9918", "9919", "9921", "9924", "9925", "9926", "9927", "9928", "9929", "9930", "9931", "9933",
    "9934", "9935", "9937", "9938", "9939", "9940", "9941", "9941A", "9942", "9943", "9944", "9945", "9946", "9955", "9958",
]

TPEX_LIST = [
    "1240", "1259", "1264", "1268", "1294", "1295", "1336", "1565", "1569", "1570", "1580", "1584", "1586", "1591", "1593", "1595", "1599", "1742", "1777", "1780",
    "1781", "1784", "1785", "1788", "1796", "1799", "1813", "1815", "2035", "2061", "2063", "2064", "2065", "2066", "2067", "2070", "2073", "2221", "2230", "2235",
    "2596", "2640", "2641", "2643", "2718", "2719", "2724", "2726", "2729", "2732", "2734", "2736", "2740", "2743", "2745", "2751", "2752", "2754", "2755", "2756",
    "2916", "2924", "2926", "2937", "2941", "2947", "2948", "2949", "3064", "3066", "3067", "3071", "3073", "3078", "3081", "3083", "3085", "3086", "3088", "3093",
    "3095", "3105", "3114", "3115", "3118", "3122", "3128", "3131", "3141", "3147", "3152", "3158", "3162", "3163", "3169", "3171", "3176", "3178", "3188", "3191",
    "3205", "3206", "3207", "3211", "3213", "3217", "3218", "3219", "3221", "3224", "3226", "3227", "3228", "3230", "3232", "3234", "3236", "3252", "3259", "3260",
    "3264", "3265", "3268", "3272", "3276", "3284", "3285", "3287", "3288", "3289", "3290", "3293", "3294", "3297", "3303", "3306", "3310", "3313", "3317", "3322",
    "3323", "3324", "3325", "3332", "3339", "3349", "3354", "3357", "3360", "3362", "3363", "3372", "3373", "3374", "3379", "3388", "3390", "3402", "3430", "3434",
    "3438", "3441", "3444", "3455", "3465", "3466", "3467", "3479", "3483", "3484", "3485", "3489", "3490", "3491", "3492", "3498", "3499", "3508", "3511", "3512",
    "3516", "3520", "3521", "3522", "3523", "3526", "3527", "3529", "3531", "3537", "3540", "3541", "3546", "3548", "3551", "3552", "3555", "3556", "3558", "3564",
    "3567", "3570", "3577", "3580", "3581", "3587", "3594", "3597", "3609", "3611", "3615", "3623", "3624", "3625", "3628", "3629", "3630", "3631", "3632", "3646",
    "3663", "3664", "3666", "3672", "3675", "3680", "3684", "3685", "3687", "3689", "3691", "3693", "3707", "3709", "3710", "3713", "4102", "4105", "4107", "4109",
    "4111", "4113", "4114", "4116", "4120", "4121", "4123", "4126", "4127", "4128", "4129", "4131", "4138", "4139", "4147", "4153", "4154", "4157", "4160", "4161",
    "4162", "4163", "4166", "4167", "4168", "4171", "4173", "4174", "4175", "4183", "4188", "4192", "4198", "4205", "4207", "4303", "4304", "4305", "4401", "4402",
    "4406", "4413", "4416", "4417", "4419", "4420", "4430", "4432", "4433", "4442", "4502", "4503", "4506", "4510", "4513", "4523", "4527", "4528", "4529", "4530",
    "4533", "4534", "4535", "4538", "4541", "4542", "4543", "4549", "4550", "4554", "4556", "4558", "4561", "4563", "4568", "4577", "4580", "4584", "4609", "4702",
    "4706", "4707", "4711", "4714", "4716", "4721", "4726", "4728", "4729", "4735", "4741", "4743", "4744", "4745", "4747", "4749", "4754", "4760", "4767", "4768",
    "4772", "4806", "4903", "4905", "4907", "4908", "4909", "4911", "4923", "4924", "4931", "4933", "4939", "4946", "4950", "4951", "4953", "4966", "4971", "4972",
    "4973", "4974", "4979", "4991", "4995", "5009", "5011", "5013", "5014", "5015", "5016", "5201", "5202", "5205", "5206", "5209", "5210", "5211", "5212", "5213",
    "5220", "5223", "5227", "5228", "5230", "5245", "5251", "5263", "5272", "5274", "5276", "5278", "5287", "5289", "5291", "5299", "5301", "5302", "5309", "5310",
    "5312", "5314", "5315", "5321", "5324", "5328", "5340", "5344", "5345", "5347", "5348", "5351", "5353", "5355", "5356", "5364", "5371", "5381", "5386", "5392",
    "5398", "5403", "5410", "5425", "5426", "5432", "5438", "5439", "5443", "5450", "5452", "5455", "5457", "5460", "5464", "5465", "5468", "5474", "5475", "5478",
    "5481", "5483", "5487", "5488", "5489", "5490", "5493", "5498", "5508", "5511", "5512", "5514", "5516", "5520", "5523", "5529", "5530", "5536", "5543", "5547",
    "5548", "5601", "5603", "5604", "5609", "5701", "5703", "5704", "5864", "5878", "5902", "5903", "5904", "5905", "6015", "6016", "6020", "6021", "6023", "6026",
    "6028", "6101", "6103", "6104", "6109", "6111", "6113", "6114", "6118", "6121", "6122", "6123", "6124", "6125", "6126", "6127", "6129", "6130", "6134", "6138",
    "6140", "6143", "6144", "6146", "6147", "6148", "6150", "6151", "6154", "6156", "6158", "6160", "6161", "6163", "6167", "6169", "6170", "6171", "6173", "6174",
    "6175", "6179", "6180", "6182", "6185", "6186", "6187", "6188", "6190", "6194", "6195", "6198", "6199", "6203", "6204", "6207", "6208", "6210", "6212", "6217",
    "6218", "6219", "6220", "6221", "6222", "6223", "6227", "6228", "6229", "6231", "6233", "6234", "6236", "6237", "6240", "6241", "6242", "6244", "6245", "6246",
    "6248", "6259", "6261", "6263", "6264", "6265", "6266", "6270", "6274", "6275", "6276", "6279", "6284", "6290", "6291", "6292", "6294", "6411", "6417", "6418",
    "6419", "6423", "6425", "6432", "6435", "6441", "6461", "6462", "6465", "6469", "6470", "6474", "6482", "6485", "6486", "6488", "6492", "6494", "6496", "6498",
    "6499", "6506", "6508", "6509", "6510", "6512", "6516", "6517", "6523", "6527", "6530", "6532", "6535", "6538", "6542", "6546", "6547", "6548", "6556", "6560",
    "6561", "6568", "6569", "6570", "6574", "6576", "6577", "6578", "6584", "6588", "6590", "6593", "6596", "6597", "6603", "6609", "6612", "6613", "6615", "6616",
    "6617", "6620", "6624", "6629", "6637", "6640", "6642", "6643", "6649", "6651", "6654", "6661", "6662", "6664", "6667", "6679", "6680", "6683", "6684", "6690",
    "6692", "6693", "6697", "6703", "6708", "6712", "6716", "6720", "6721", "6725", "6727", "6728", "6730", "6732", "6733", "6735", "6739", "6741", "6751", "6752",
    "6761", "6762", "6763", "6767", "6785", "6788", "6791", "6803", "6804", "6811", "6821", "6823", "6829", "6840", "6841", "6843", "6844", "6846", "6855", "6856",
    "6859", "6865", "6870", "6872", "6874", "6875", "6877", "6881", "6884", "6894", "6895", "6899", "6903", "6904", "6907", "6910", "6913", "6922", "6925", "6929",
    "6945", "6953", "6961", "6967", "6968", "6971", "6982", "6983", "6986", "6996", "6997", "7402", "7547", "7556", "7584", "7642", "7703", "7704", "7708", "7709",
    "7712", "7713", "7714", "7715", "7716", "7717", "7718", "7723", "7728", "7734", "7738", "7743", "7744", "7747", "7751", "7753", "7757", "7767", "7770", "7772",
    "7777", "7782", "7792", "7794", "7805", "7810", "7811", "7814", "7819", "7820", "7828", "7839", "7842", "8024", "8027", "8032", "8034", "8038", "8040", "8042",
    "8043", "8044", "8047", "8048", "8049", "8050", "8054", "8059", "8064", "8066", "8067", "8068", "8069", "8071", "8074", "8076", "8077", "8080", "8083", "8084",
    "8085", "8086", "8087", "8088", "8089", "8091", "8092", "8093", "8096", "8097", "8099", "8102", "8107", "8109", "8111", "8121", "8147", "8155", "8171", "8176",
    "8182", "8183", "8227", "8234", "8240", "8255", "8272", "8277", "8279", "8284", "8289", "8291", "8299", "8342", "8349", "8349A", "8354", "8358", "8383", "8390",
    "8401", "8403", "8409", "8410", "8415", "8416", "8421", "8423", "8424", "8426", "8431", "8432", "8433", "8435", "8436", "8437", "8440", "8444", "8446", "8450",
    "8455", "8472", "8477", "8489", "8905", "8906", "8908", "8916", "8917", "8921", "8923", "8924", "8927", "8928", "8929", "8930", "8931", "8932", "8933", "8935",
    "8936", "8937", "8938", "8941", "8942", "9949", "9950", "9951", "9960", "9962",
]

MY_LIST_DEFAULT = [
    "4979", "6870", "6451", "3450", "3163", "4977", "3363", "6533", "6757", "2345", "2376", "2441", "6285", "2049", "2303", "3653", "2467", "2359", "2308", "3017",
    "2368", "4576", "2327", "6442", "6531", "6683", "6257", "6223", "8150", "3105", "2634", "8299", "3532", "6213", "3033", "4991", "2449", "6781", "3189", "8046",
    "3037", "3443", "3491", "2481", "6770", "2408", "8271", "3081", "3665", "6274", "2455", "2454", "2360", "3673", "8021", "4573", "2351", "1560", "8028", "6197",
    "4958", "8358", "2330", "2383", "1303", "6669", "2137", "3231", "1815", "5475", "5340", "3771", "1519", "3661", "2059", "6805", "6510", "3211", "4931", "8210",
    "3013", "6117", "3693",
]

# 0050 成分股：2026-09-04 公開持股快照（51檔），下次調整約 2026 年 12 月

TW0050_LIST = [
    "2330", "2454", "2308", "2317", "3711", "2383", "2303", "2881", "2891", "3037", "3017", "1303", "2882", "2345", "2887", "2382", "2327", "2885", "2357", "2884",
    "3008", "2301", "3231", "2886", "2883", "2408", "2890", "2344", "2412", "2892", "2449", "2337", "4938", "2356", "2368", "5880", "1301", "1326", "2002", "1101",
    "1216", "3045", "4904", "2379", "3034", "1102", "2610", "2618", "2603", "2609", "2615",
]

# 選股型指定組合（S編號）：2023-10～2026-09 台股三年回測（分三段：偏多／空頭／多頭），三年的依日 t 值都 ≥2 的組合（S4、S7、S8 為 2026-10-05 加入的營收創新高系列）
# （2026-09-30 改版；舊版 S1～S7 多數只在多頭年有效，已移除或改列高標股 M）
STOCK_PICK_COMBOS = [
    ["距52週高點≤5%", "近3日向上跳空缺口", "近3月乖離度為正(營收優於股價)"],
    ["創52週新高", "近3日向上跳空缺口", "近3月乖離度為正(營收優於股價)"],
    ["相對強弱為正(強於大盤)", "量能區間低檔(≤10百分位)", "距52週高點≤5%"],
    ["近3日漲停", "月營收創24個月新高", "外資近5日買超"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "距52週高點≤5%"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "創52週新高"],
    ["大盤站上20日均線", "漲時量≥跌時量1.5倍(近20日)", "月營收創24個月新高"],
    ["近20日強於大盤≥10%", "大盤站上20日均線", "月營收創24個月新高"],

]
STOCK_PICK_STATS = {
    "距52週高點≤5% ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 4029, "win": 54.6, "ret": 3.33, "ex": 2.73, "t": 12.61}, "20": {"n": 4029, "win": 54.7, "ret": 5.17, "ex": 3.76, "t": 12.27}},
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2838, "win": 53.6, "ret": 3.79, "ex": 3.17, "t": 11.06}, "20": {"n": 2838, "win": 52.9, "ret": 5.53, "ex": 4.1, "t": 10.09}},
    "相對強弱為正(強於大盤) ＋ 量能區間低檔(≤10百分位) ＋ 距52週高點≤5%": {"10": {"n": 644, "win": 57.8, "ret": 2.28, "ex": 2.15, "t": 6.11}, "20": {"n": 644, "win": 56.7, "ret": 3.48, "ex": 2.29, "t": 4.93}},
    "近3日漲停 ＋ 月營收創24個月新高 ＋ 外資近5日買超": {"10": {"n": 2970, "win": 55.8, "ret": 4.88, "ex": 3.42, "t": 11.37}, "20": {"n": 2970, "win": 57.0, "ret": 8.59, "ex": 5.81, "t": 13.21}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 距52週高點≤5%": {"10": {"n": 2004, "win": 53.4, "ret": 2.69, "ex": 2.25, "t": 7.85}, "20": {"n": 2004, "win": 55.7, "ret": 5.25, "ex": 4.14, "t": 9.5}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 創52週新高": {"10": {"n": 267, "win": 59.2, "ret": 6.59, "ex": 6.22, "t": 5.5}, "20": {"n": 267, "win": 64.4, "ret": 12.11, "ex": 10.46, "t": 6.38}},
    "大盤站上20日均線 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 月營收創24個月新高": {"10": {"n": 16520, "win": 50.5, "ret": 2.37, "ex": 1.72, "t": 18.15}, "20": {"n": 16520, "win": 51.3, "ret": 4.62, "ex": 3.23, "t": 22.56}},
    "近20日強於大盤≥10% ＋ 大盤站上20日均線 ＋ 月營收創24個月新高": {"10": {"n": 8373, "win": 51.8, "ret": 3.09, "ex": 2.42, "t": 15.2}, "20": {"n": 8373, "win": 52.2, "ret": 5.35, "ex": 3.83, "t": 16.44}},

}

# 2026-09-27 台股回測（勝率榜5/10/20日＋飆股榜5/10/20日）中，指定組合的勝率／t值(同日調整)／飆股比例記錄值
# key＝條件依字元排序後以「 ＋ 」連接
COMBO_REF_STATS = {
    # ── 2026-10-05 台股三年三段回測（含月營收／外資／漲停等新條件）：S4 換成「漲停＋營收創24個月新高＋外資買超」，新增 S7、S8 營收創新高系列 ──
    "距52週高點≤5% ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 52.7, "10": 54.6, "20": 54.7}, "t": {"5": 9.65, "10": 12.61, "20": 12.27}, "moon": {"10": 5.0, "20": 9.5}},
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 52.3, "10": 53.6, "20": 52.9}, "t": {"5": 7.57, "10": 11.06, "20": 10.09}, "moon": {"10": 6.6, "20": 11.7}},
    "相對強弱為正(強於大盤) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": {"win": {"5": 58.1, "10": 57.8, "20": 56.7}, "t": {"5": 5.44, "10": 6.11, "20": 4.93}, "moon": {"10": 1.2, "20": 4.7}},
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": {"win": {"5": 53.8, "10": 55.1, "20": 55.1}, "t": {"5": 5.96, "10": 7.31, "20": 6.41}, "moon": {"10": 3.8, "20": 7.4}},
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 多方力道≥80": {"win": {"5": 49.9, "10": 52.3, "20": 49.2}, "t": {"5": 4.7, "10": 5.41, "20": 4.74}, "moon": {"10": 5.7, "20": 9.9}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 強勢突破盤": {"win": {"20": 60.2}, "t": {"20": 6.52}, "moon": {"10": 10.7, "20": 21.0}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 11.0, "20": 22.0}},
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": {"win": {"5": 55.4, "10": 55.7, "20": 57.5}, "t": {"5": 4.63, "10": 5.41, "20": 5.27}, "moon": {"10": 11.7, "20": 19.3}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3日向上跳空缺口": {"win": {"5": 57.6, "10": 71.1, "20": 72.9}, "t": {"5": 0.2, "10": 4.87, "20": 7.35}, "moon": {"10": 0.7, "20": 4.3}},
    "KDJ近3日內死亡交叉 ＋ KDJ近3日內黃金交叉 ＋ 夜星剛形成": {"moon": {"20": 10.5}},
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"moon": {"20": 15.3}},
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 11.5, "20": 23.1}},
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.0}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ N字底剛形成": {"win": {"20": 60.8}, "t": {"20": 0.66}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 大盤跌破60日均線": {"win": {"20": 61.4}, "t": {"20": 0.97}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": {"win": {"5": 63.6}, "t": {"5": 2.96}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.1}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 61.2, "10": 62.3, "20": 62.0}, "t": {"5": 1.87, "10": 1.07, "20": -0.25}},
    "KDJ近3日內死亡交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"win": {"10": 61.1, "20": 63.9}, "t": {"10": 2.25, "20": 2.19}},
    "KDJ近3日內死亡交叉 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"10": 62.2}, "t": {"10": 0.86}},
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.5}},
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"10": 60.2}, "t": {"10": 1.8}},
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"moon": {"20": 10.3}},
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.8}},
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"moon": {"10": 17.4, "20": 13.0}},
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"20": 10.5}},
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"20": 13.1}},
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.0}},
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 10.3}},
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 14.9, "20": 10.6}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 63.2, "10": 60.9, "20": 62.7}, "t": {"5": 4.59, "10": 3.52, "20": 1.07}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 62.9, "10": 66.3}, "t": {"5": 0.11, "10": -0.01}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 62.1}, "t": {"10": 0.73}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 61.8, "20": 61.7}, "t": {"5": 7.09, "20": 3.76}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 64.9, "10": 63.1, "20": 64.9}, "t": {"5": 0.32, "10": -1.13, "20": -0.69}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 60.5}, "t": {"5": 1.78}},
    "KDJ近3日內死亡交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 10.1}},
    "KDJ近3日內死亡交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.7}},
    "KDJ近3日內死亡交叉 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 12.3}},
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.8}},
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 60.6}, "t": {"10": 0.19}},
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.0}},
    "KDJ近3日內死亡交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.8}},
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"moon": {"20": 11.7}},
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 10.8, "20": 16.2}},
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 夜星剛形成": {"moon": {"20": 13.4}},
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": {"moon": {"20": 10.0}},
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 61.7}, "t": {"10": 2.28}},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"win": {"20": 62.2}, "t": {"20": 1.59}},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 多方力道≥80": {"moon": {"20": 10.5}},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.0}},
    "KDJ近3日內黃金交叉 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": {"moon": {"20": 12.2}},
    "KDJ近3日內黃金交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"moon": {"20": 13.5}},
    "KDJ近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"20": 10.8}},
    "KDJ近3日內黃金交叉 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 10.3}},
    "KDJ近3日內黃金交叉 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"10": 10.3, "20": 14.4}},
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 13.2}},
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.8}},
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.4}},
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 13.3, "20": 13.3}},
    "KDJ近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 10.2}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"win": {"5": 65.4, "10": 65.6, "20": 63.3}, "t": {"5": 1.13, "10": 0.59, "20": 1.67}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"20": 60.1}, "t": {"20": 1.39}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 跌深反彈盤": {"win": {"10": 61.3}, "t": {"10": -0.38}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"10": 60.4, "20": 60.2}, "t": {"10": 1.65, "20": -0.14}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 71.9, "10": 72.8, "20": 66.3}, "t": {"5": 1.5, "10": 0.0, "20": -0.02}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"5": 65.9, "10": 69.9, "20": 67.0}, "t": {"5": 1.41, "10": 2.92, "20": 2.13}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 60.7}, "t": {"20": 0.11}},
    "KDJ近3日內黃金交叉 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 78.5, "10": 82.5, "20": 74.9}, "t": {"5": -0.27, "10": -0.52, "20": -0.19}},
    "KDJ近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 11.8}},
    "KDJ近3日內黃金交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.8}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"win": {"20": 60.4}, "t": {"20": 3.03}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 64.1, "10": 64.3, "20": 61.6}, "t": {"5": 1.54, "10": 0.92, "20": 1.11}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.2}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.4}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"10": 60.4}, "t": {"10": 3.29}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"win": {"10": 60.2}, "t": {"10": 3.48}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.0}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 60.1}, "t": {"5": 1.78}},
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.6}},
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.6}},
    "K線橫盤的突破剛形成 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": {"moon": {"10": 11.6, "20": 23.3}},
    "K線橫盤的突破剛形成 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"win": {"20": 60.0}, "t": {"20": 1.96}},
    "K線橫盤的突破剛形成 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"moon": {"20": 12.2}},
    "K線橫盤的突破剛形成 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.3}},
    "K線橫盤的突破剛形成 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"moon": {"20": 15.0}},
    "K線橫盤的突破剛形成 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.1}},
    "K線橫盤的突破剛形成 ＋ 回後買上漲全通過 ＋ 晨星剛形成": {"moon": {"20": 10.9}},
    "K線橫盤的突破剛形成 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.8}},
    "K線橫盤的突破剛形成 ＋ 圓弧底剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.3}},
    "K線橫盤的突破剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 12.1}},
    "K線橫盤的突破剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.5}},
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"20": 11.9}},
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 12.9, "20": 12.9}},
    "K線橫盤的突破剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 14.6}},
    "K線橫盤的突破剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.4}},
    "K線橫盤的突破剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 12.6}},
    "K線橫盤的突破剛形成 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.4}},
    "K線橫盤的突破剛形成 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"moon": {"20": 17.2}},
    "K線橫盤的突破剛形成 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 15.0}},
    "K線橫盤的突破剛形成 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 13.3}},
    "K線橫盤的突破剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 13.9}},
    "K線橫盤的突破剛形成 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 15.2}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成": {"moon": {"20": 13.2}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 13.9}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 12.0}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.3}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.9}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 16.1}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 16.4}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 11.2}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 14.2}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"10": 10.9, "20": 16.3}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.7}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 12.9}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.6}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 11.5}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 11.5}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.3}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 21.9}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"20": 60.0}, "t": {"20": 1.83}, "moon": {"20": 14.0}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 12.8}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.0}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.9}},
    "K線橫盤的突破剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.9}},
    "MACD近3日內黃金交叉 ＋ N字底剛形成 ＋ 多方力道≥80": {"moon": {"20": 14.3}},
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.8}},
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"moon": {"20": 12.5}},
    "MACD近3日內黃金交叉 ＋ 三重底剛形成 ＋ 晨星剛形成": {"moon": {"20": 14.3}},
    "MACD近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": {"win": {"10": 61.9}, "t": {"10": -1.37}},
    "MACD近3日內黃金交叉 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.2}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 15.4}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 13.2}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 12.8}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.6}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 15.8}},
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.6}},
    "MACD近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.1}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"20": 60.7}, "t": {"20": -2.3}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 61.3, "10": 61.3, "20": 61.4}, "t": {"5": 1.09, "10": 0.05, "20": -0.71}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 60.2, "20": 61.9}, "t": {"10": -0.51, "20": 2.04}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 量能區間低檔(≤10百分位)": {"win": {"5": 62.9, "10": 62.9, "20": 61.4}, "t": {"5": 1.69, "10": 0.91, "20": -0.8}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 64.2, "10": 64.7, "20": 63.4}, "t": {"5": -0.17, "10": -1.18, "20": -1.75}},
    "MACD近3日內黃金交叉 ＋ 布林通道低檔(≤20%) ＋ 量能區間低檔(≤10百分位)": {"win": {"5": 60.5}, "t": {"5": 1.22}},
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 12.0}},
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.0}},
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.8}},
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.3}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.6}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 12.1}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 11.4}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.9}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.2}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 11.3}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 11.2}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.8}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"10": 61.1}, "t": {"10": 1.05}, "moon": {"20": 11.1}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.4}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.9}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 13.9}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 三大法人近3月買超": {"win": {"20": 60.6}, "t": {"20": 2.96}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 大盤站上20日均線": {"win": {"20": 60.6}, "t": {"20": 2.35}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 相對強弱為正(強於大盤)": {"win": {"20": 60.2}, "t": {"20": 2.29}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 61.1}, "t": {"10": 1.01}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"10": 60.3, "20": 61.7}, "t": {"10": 2.08, "20": 2.01}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 近3月均價YoY為正": {"win": {"10": 60.8, "20": 65.1}, "t": {"10": 1.82, "20": 2.43}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 頭肩底剛形成": {"win": {"10": 61.2, "20": 61.2}, "t": {"10": 2.53, "20": 1.6}},
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.2}},
    "N字底剛形成 ＋ 三大法人近3月賣超 ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 12.5, "20": 15.0}},
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": {"moon": {"20": 12.7}},
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 多方力道≥80": {"moon": {"20": 11.6}},
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 10.3, "20": 10.3}},
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": {"moon": {"20": 11.1}},
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%)": {"moon": {"20": 10.0}},
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"10": 10.3, "20": 13.8}},
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 61.8}, "t": {"10": -0.01}},
    "N字底剛形成 ＋ 多方力道≥65 ＋ 突破ABC修正下降切線剛形成": {"moon": {"5": 11.8, "10": 17.6, "20": 14.7}},
    "N字底剛形成 ＋ 多方力道≥65 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.4}},
    "N字底剛形成 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"10": 10.0}},
    "N字底剛形成 ＋ 多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 12.5}},
    "N字底剛形成 ＋ 多方力道≥80 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.4}},
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"10": 10.9}},
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 10.3, "20": 10.3}},
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 10.3}},
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.9}},
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"10": 11.1}},
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 10.8, "20": 12.2}},
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.8}},
    "N字底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"10": 12.9}},
    "N字底剛形成 ＋ 強勢突破盤 ＋ 突破ABC修正下降切線剛形成": {"moon": {"5": 10.3, "10": 13.8, "20": 10.3}},
    "N字底剛形成 ＋ 強勢突破盤 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.5}},
    "N字底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.0}},
    "N字底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"10": 10.5}},
    "N字底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": {"moon": {"10": 12.5, "20": 12.5}},
    "N字底剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.4}},
    "N字底剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 11.5}},
    "N字底剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 12.5}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 16.3}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 12.5}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": {"moon": {"10": 11.1}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"10": 11.6}},
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 13.4}},
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.2}},
    "一字底(均線糾結)剛形成 ＋ 三大法人近3月買超 ＋ 頭肩底剛形成": {"win": {"20": 60.0}, "t": {"20": 1.61}},
    "一字底(均線糾結)剛形成 ＋ 三重底剛形成 ＋ 頭肩底剛形成": {"win": {"5": 60.7, "10": 66.1, "20": 60.7}, "t": {"5": 1.83, "10": 1.73, "20": 0.79}},
    "一字底(均線糾結)剛形成 ＋ 回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"10": 66.1}, "t": {"10": 2.93}},
    "一字底(均線糾結)剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"5": 60.0, "10": 69.1}, "t": {"5": 1.88, "10": 0.82}},
    "一字底(均線糾結)剛形成 ＋ 大盤站上20日均線 ＋ 頭肩底剛形成": {"win": {"10": 61.1, "20": 65.3}, "t": {"10": 2.07, "20": 1.91}},
    "一字底(均線糾結)剛形成 ＋ 大盤站上60日均線 ＋ 頭肩底剛形成": {"win": {"10": 63.1, "20": 63.1}, "t": {"10": 1.99, "20": 1.54}},
    "一字底(均線糾結)剛形成 ＋ 布林通道高檔(≥80%) ＋ 頭肩底剛形成": {"win": {"20": 62.3}, "t": {"20": 1.58}},
    "一字底(均線糾結)剛形成 ＋ 強勢突破盤 ＋ 頭肩底剛形成": {"win": {"10": 61.7, "20": 63.3}, "t": {"10": 2.11, "20": 1.22}},
    "一字底(均線糾結)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 頭肩底剛形成": {"win": {"20": 61.2}, "t": {"20": 1.41}},
    "一字底(均線糾結)剛形成 ＋ 頭肩底剛形成": {"win": {"20": 61.5}, "t": {"20": 1.53}},
    "三大法人近3月買超 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": {"moon": {"20": 12.6}},
    "三大法人近3月買超 ＋ 回後買上漲全通過 ＋ 多方力道≥80": {"moon": {"20": 10.8}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"moon": {"20": 17.5}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 16.2}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.5}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 頭肩底剛形成": {"moon": {"20": 13.9}},
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"win": {"10": 62.7}, "t": {"10": 2.44}, "moon": {"10": 16.9, "20": 20.3}},
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"20": 12.8}},
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": {"moon": {"20": 13.0}},
    "三大法人近3月買超 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 12.9}},
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 夜星剛形成": {"moon": {"20": 10.9}},
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"20": 15.8}},
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.1}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 13.9}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 14.2}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.6}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.0}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"10": 61.5}, "t": {"10": 2.02}, "moon": {"10": 13.8, "20": 16.9}},
    "三大法人近3月買超 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "三大法人近3月買超 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 11.1}},
    "三大法人近3月買超 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"win": {"5": 62.5, "10": 60.6}, "t": {"5": 1.67, "10": 0.53}},
    "三大法人近3月買超 ＋ 大盤跌破20日均線 ＋ 跌深反彈盤": {"win": {"5": 66.0}, "t": {"5": 2.88}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 60.5, "10": 60.5}, "t": {"5": 4.1, "10": 4.99}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 69.3, "10": 69.1, "20": 61.8}, "t": {"5": 2.35, "10": 0.9, "20": 0.51}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"5": 60.5}, "t": {"5": 0.97}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 60.2, "20": 60.2}, "t": {"10": 7.91, "20": 7.06}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 72.9, "10": 67.8, "20": 64.4}, "t": {"5": 1.8, "10": -0.36, "20": -0.87}},
    "三大法人近3月買超 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 76.3, "10": 81.3, "20": 77.0}, "t": {"5": 0.05, "10": -0.23, "20": 0.14}},
    "三大法人近3月買超 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": {"win": {"20": 60.2}, "t": {"20": -0.94}},
    "三大法人近3月買超 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 10.8}},
    "三大法人近3月買超 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.1}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.3}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 62.1, "10": 62.1}, "t": {"5": 2.23, "10": 2.33}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 13.2}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.8}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.2}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.2}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 15.9}},
    "三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.2}},
    "三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.5}},
    "三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.6}},
    "三大法人近3月賣超 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.1}},
    "三大法人近3月賣超 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 10.7}},
    "三大法人近3月賣超 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 10.2}},
    "三大法人近3月賣超 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"win": {"10": 60.9}, "t": {"10": -1.2}},
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 64.5, "10": 66.3, "20": 63.0}, "t": {"5": -0.65, "10": -1.36, "20": -0.64}},
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 61.5}, "t": {"10": 0.54}},
    "三大法人近3月賣超 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 75.9, "10": 77.6, "20": 69.0}, "t": {"5": -0.27, "10": -1.4, "20": -0.82}},
    "三大法人近3月賣超 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 13.0}},
    "三大法人近3月賣超 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.8}},
    "三大法人近3月賣超 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.1}},
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.1}},
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 18.9}},
    "三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 18.5}},
    "三重底剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.4}},
    "三重底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 23.1}},
    "三重底剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.0}},
    "三重底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.9}},
    "三重底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.8}},
    "三重底剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.3}},
    "三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 12.0}},
    "分價量表-突破POC追價買進 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"5": 66.1}, "t": {"5": 3.55}, "moon": {"20": 20.3}},
    "回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.3}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量)": {"moon": {"20": 11.4}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥65": {"moon": {"20": 13.9}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": {"moon": {"10": 10.0, "20": 17.5}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤站上20日均線": {"moon": {"20": 12.1}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤站上60日均線": {"moon": {"20": 13.0}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": {"win": {"10": 62.5, "20": 64.3}, "t": {"10": -1.13, "20": 0.16}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%)": {"moon": {"20": 16.2}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"win": {"20": 61.4}, "t": {"20": 3.04}, "moon": {"10": 12.0, "20": 24.1}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 13.5}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 突破ABC修正下降切線剛形成": {"moon": {"5": 10.0, "10": 16.7, "20": 16.7}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 14.3}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為正": {"moon": {"20": 12.0}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.5}},
    "回後買上漲全通過 ＋ 多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.2}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 大盤站上20日均線": {"moon": {"20": 10.9}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 大盤站上60日均線": {"moon": {"20": 10.3}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.4}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.4}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 近3月均價YoY為負": {"moon": {"20": 10.8}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.7}},
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 10.3}},
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.2}},
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 60.3}, "t": {"10": 0.86}},
    "回後買上漲全通過 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 10.0}},
    "回後買上漲全通過 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.2}},
    "回後買上漲全通過 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.6}},
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"win": {"10": 60.7}, "t": {"10": 2.01}, "moon": {"10": 10.7, "20": 14.3}},
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.1}},
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.9}},
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 12.5, "20": 12.5}},
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.5}},
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.1}},
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.8}},
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.5}},
    "回後買上漲全通過 ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.5}},
    "回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.5}},
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.3}},
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.3}},
    "圓弧底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 15.1}},
    "圓弧底剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 17.0}},
    "圓弧底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.0}},
    "圓弧底剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 13.6}},
    "圓弧底剛形成 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 13.6}},
    "圓弧底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 14.6}},
    "圓弧底剛形成 ＋ 晨星剛形成": {"moon": {"20": 12.7}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.0}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 10.4}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.8}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.3}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 12.8}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.9}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.0}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 14.6}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.2}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 16.0}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 12.5}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 15.4}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 14.3}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.7}},
    "圓弧底剛形成 ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.3}},
    "圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.4}},
    "圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.3}},
    "圓弧底剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 頭肩底剛形成": {"moon": {"20": 11.9}},
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 強勢突破盤": {"moon": {"20": 12.9}},
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 夜星剛形成": {"win": {"10": 60.3}, "t": {"10": 2.36}, "moon": {"10": 14.3, "20": 14.3}},
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 強勢突破盤": {"moon": {"20": 10.2}},
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 強勢突破盤": {"moon": {"20": 12.4}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 夜星剛形成": {"moon": {"10": 14.5, "20": 17.7}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 布林通道高檔(≥80%)": {"moon": {"20": 10.1}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 強勢突破盤": {"moon": {"20": 15.5}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 18.8}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 13.3}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 夜星剛形成": {"moon": {"10": 17.4, "20": 17.4}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 布林通道高檔(≥80%)": {"moon": {"20": 12.4}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"10": 10.3, "20": 22.4}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.5}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 11.4}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成": {"moon": {"10": 13.1, "20": 16.7}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤站上20日均線": {"moon": {"10": 14.0, "20": 14.0}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤站上60日均線": {"moon": {"10": 12.5, "20": 17.2}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"10": 11.8, "20": 20.6}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"5": 10.0, "10": 15.0, "20": 15.0}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": {"moon": {"10": 16.1, "20": 19.4}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"10": 14.5, "20": 17.1}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"10": 13.9, "20": 16.7}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"10": 13.3, "20": 17.8}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月均價YoY為正": {"win": {"10": 65.5}, "t": {"10": 2.41}, "moon": {"10": 12.1, "20": 17.2}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月均價YoY為負": {"moon": {"10": 16.7, "20": 16.7}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 17.6, "20": 26.5}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 強勢突破盤": {"moon": {"20": 11.8}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 16.0}},
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 強勢突破盤": {"moon": {"20": 11.4}},
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 12.4}},
    "地量(≤0.5倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 10.4}},
    "地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 60.9, "10": 64.1}, "t": {"5": 0.96, "10": 1.13}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 13.3}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 11.1, "20": 16.7}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(高檔)剛形成": {"moon": {"10": 10.3}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 10.0, "20": 13.3}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.2}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"20": 11.1}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 13.6}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 13.0, "20": 30.4}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 13.3}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月均價YoY為正": {"moon": {"20": 11.7}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月均價YoY為負": {"moon": {"20": 10.2}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.3}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.6}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.7}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 11.4, "20": 11.4}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.0}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.4}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 12.1}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.2}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 16.8}},
    "地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.6}},
    "型態剛形成(剛突破) ＋ 多方力道≥80 ＋ 量能區間低檔(≤10百分位)": {"moon": {"20": 16.7}},
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 62.8, "10": 66.3, "20": 61.7}, "t": {"5": -1.49, "10": -1.71, "20": -0.75}},
    "型態剛形成(剛突破) ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": {"win": {"5": 70.9, "10": 71.8, "20": 71.8}, "t": {"5": 0.27, "10": -1.4, "20": -1.11}},
    "型態成形中 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"10": 60.3}, "t": {"10": 3.0}, "moon": {"10": 11.8, "20": 14.7}},
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 61.3, "10": 61.3, "20": 60.6}, "t": {"5": -1.06, "10": -2.74, "20": -1.5}},
    "型態突破確認 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 10.7}},
    "型態突破確認 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 62.7, "10": 65.6, "20": 61.7}, "t": {"5": -1.07, "10": -1.72, "20": -0.4}},
    "型態突破確認 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": {"win": {"5": 68.1, "10": 69.9, "20": 69.9}, "t": {"5": 0.16, "10": -1.36, "20": -1.0}},
    "多方力道≥65 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 10.6}},
    "多方力道≥65 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 10.1}},
    "多方力道≥65 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.1}},
    "多方力道≥65 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 11.9, "20": 11.9}},
    "多方力道≥65 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 12.7}},
    "多方力道≥65 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 11.7}},
    "多方力道≥65 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 10.3}},
    "多方力道≥65 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 64.7, "20": 61.7}, "t": {"10": 0.6, "20": 0.61}},
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 11.6}},
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.0}},
    "多方力道≥65 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.5}},
    "多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.3}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.9}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.1}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 14.9}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.4}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 13.1}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 12.1}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 15.9}},
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.5}},
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 12.9}},
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.4}},
    "多方力道≥65 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.7}},
    "多方力道≥80 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 16.3, "20": 16.3}},
    "多方力道≥80 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 14.6}},
    "多方力道≥80 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 13.8}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"moon": {"20": 11.1}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.3}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 10.3}},
    "多方力道≥80 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"moon": {"20": 15.3}},
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 12.7}},
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 61.1}, "t": {"10": 2.48}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.8}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能區間低檔(≤10百分位)": {"moon": {"20": 12.1}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.3}},
    "多方力道≥80 ＋ 晨星剛形成": {"moon": {"20": 14.0}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.0}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.9}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.4}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 13.2, "20": 15.8}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.7}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.9}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 15.5}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"10": 10.0, "20": 15.2}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 11.5}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.2}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.3}},
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 11.0}},
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.6}},
    "多方力道≥80 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.0}},
    "多方力道≥80 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"win": {"10": 63.4}, "t": {"10": 1.06}},
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.2}},
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.5}},
    "多方力道≥80 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間低檔(≤10百分位)": {"moon": {"20": 10.5}},
    "夜星剛形成 ＋ 大盤站上20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 10.6, "20": 10.6}},
    "夜星剛形成 ＋ 大盤站上60日均線 ＋ 大盤跌破20日均線": {"moon": {"20": 12.1}},
    "夜星剛形成 ＋ 大盤站上60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.0}},
    "夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 12.2}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線": {"moon": {"20": 12.3}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 布林通道高檔(≥80%)": {"moon": {"20": 10.2}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.0}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.7}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.8}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.4}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月均價YoY為正": {"moon": {"20": 15.4}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 11.7}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 13.9}},
    "夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 10.5}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 布林通道高檔(≥80%)": {"moon": {"20": 13.3}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.0}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.5}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正": {"win": {"5": 60.6}, "t": {"5": 2.21}, "moon": {"20": 12.4}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 11.0}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 15.0, "20": 15.0}},
    "夜星剛形成 ＋ 布林通道高檔(≥80%) ＋ 相對強弱為負(弱於大盤)": {"moon": {"20": 16.0}},
    "夜星剛形成 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.5}},
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月均價YoY為正": {"moon": {"20": 11.7}},
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 12.2, "20": 13.4}},
    "夜星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 14.3}},
    "夜星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 10.9, "20": 10.9}},
    "夜星剛形成 ＋ 近3月均價YoY為正 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.6}},
    "夜星剛形成 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.3}},
    "夜星剛形成 ＋ 近3月均價YoY為負 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 12.1, "20": 15.2}},
    "夜星剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 12.9}},
    "夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.8}},
    "大盤站上20日均線 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 65.3, "20": 61.8}, "t": {"10": 1.64, "20": 2.22}},
    "大盤站上20日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 11.7}},
    "大盤站上20日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.1}},
    "大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 11.1}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.1}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 10.2}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.5}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.0}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 15.7}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 15.9}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.1}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.9}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.3}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 14.3}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.7}},
    "大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.5}},
    "大盤站上60日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 12.1}},
    "大盤站上60日均線 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.5}},
    "大盤站上60日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.8}},
    "大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 10.7}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.4}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.7}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 11.2}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 13.9}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 16.2}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 11.8}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.4}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.0}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 14.0}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.1}},
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 11.3}},
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 16.1}},
    "大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 頭肩底剛形成": {"moon": {"10": 10.0, "20": 10.0}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 60.2}, "t": {"5": 4.18}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 70.8, "10": 70.8, "20": 65.2}, "t": {"5": 1.61, "10": -0.19, "20": 0.31}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 60.5, "10": 63.2, "20": 64.6}, "t": {"5": 8.94, "10": 9.01, "20": 10.12}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 66.9, "10": 65.4, "20": 61.7}, "t": {"5": 0.28, "10": -1.85, "20": -1.59}},
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 76.6, "10": 80.2, "20": 74.2}, "t": {"5": -0.22, "10": -0.92, "20": -0.29}},
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": {"win": {"5": 62.8, "10": 62.4, "20": 64.9}, "t": {"5": 1.46, "10": 3.03, "20": 1.59}},
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": {"win": {"5": 66.0, "10": 63.7, "20": 68.0}, "t": {"5": 1.75, "10": 3.09, "20": 1.07}},
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位)": {"win": {"5": 64.3, "10": 62.8, "20": 64.9}, "t": {"5": 0.12, "10": 2.34, "20": -0.84}},
    "大盤跌破20日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 10.6}},
    "大盤跌破20日均線 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 15.4}},
    "大盤跌破20日均線 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 13.2}},
    "大盤跌破20日均線 ＋ 晨星剛形成": {"win": {"5": 60.5, "10": 60.7}, "t": {"5": 0.26, "10": -0.29}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"win": {"5": 60.5, "20": 63.6}, "t": {"5": -0.68, "20": 0.59}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 60.2, "20": 64.5}, "t": {"5": 1.36, "20": 0.86}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"win": {"5": 60.6}, "t": {"5": 1.68}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 73.4, "10": 72.7, "20": 72.5}, "t": {"5": 2.12, "10": 1.5, "20": 1.91}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.2}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 64.6, "10": 64.2, "20": 60.7}, "t": {"5": 0.86, "10": -0.42, "20": -0.14}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"20": 60.2}, "t": {"20": 1.68}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"win": {"5": 65.7, "10": 64.4, "20": 64.4}, "t": {"5": 1.66, "10": 0.53, "20": 0.48}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"20": 60.0}, "t": {"20": 2.08}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 66.1, "10": 65.6, "20": 62.6}, "t": {"5": 1.15, "10": -0.41, "20": -0.88}},
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"20": 60.7}, "t": {"20": 0.32}},
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.3}},
    "大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 13.1}},
    "大盤跌破20日均線 ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 61.8}, "t": {"20": 3.65}},
    "大盤跌破20日均線 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 62.4}, "t": {"20": 2.27}},
    "大盤跌破20日均線 ＋ 複式頭肩底剛形成 ＋ 量能區間高檔(≥90百分位)": {"win": {"10": 61.2}, "t": {"10": 2.1}},
    "大盤跌破20日均線 ＋ 跌深反彈盤 ＋ 近3月均價YoY為正": {"win": {"5": 62.5}, "t": {"5": 0.9}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 79.6, "10": 83.0, "20": 77.0}, "t": {"5": 0.4, "10": -0.33, "20": 0.21}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 母子懷抱(高檔)剛形成": {"win": {"20": 60.0}, "t": {"20": 1.15}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": {"win": {"5": 65.4, "10": 65.6, "20": 68.5}, "t": {"5": -0.3, "10": 0.1, "20": 0.51}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": {"win": {"5": 67.9, "10": 66.0, "20": 71.1}, "t": {"5": -0.6, "10": -0.0, "20": -0.3}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 62.1}, "t": {"20": 3.83}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 近3月乖離度為負(股價超前營收)": {"win": {"20": 61.0}, "t": {"20": 3.43}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位)": {"win": {"5": 68.1, "10": 66.6, "20": 68.6}, "t": {"5": -2.56, "10": -1.88, "20": -2.94}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"20": 61.2}, "t": {"20": -2.76}},
    "大盤跌破60日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "大盤跌破60日均線 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.1}},
    "大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 67.2, "10": 67.8, "20": 62.3}, "t": {"5": 1.45, "10": -0.11, "20": 0.03}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"win": {"5": 66.7, "10": 66.1, "20": 70.1}, "t": {"5": 0.3, "10": 0.78, "20": 0.8}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 64.0, "10": 62.3, "20": 68.0}, "t": {"5": 0.49, "10": 0.05, "20": 0.2}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"win": {"5": 60.2}, "t": {"5": -0.1}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 80.0, "10": 80.0, "20": 76.7}, "t": {"5": 3.17, "10": 2.2, "20": 1.47}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.8}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 65.7, "10": 65.4, "20": 62.3}, "t": {"5": 0.61, "10": -0.79, "20": 0.08}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 71.2, "10": 72.3, "20": 64.6}, "t": {"5": 1.9, "10": 0.95, "20": 0.42}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"win": {"5": 68.1, "10": 69.4, "20": 60.9}, "t": {"5": 1.25, "10": 0.41, "20": -0.15}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"win": {"5": 66.9, "10": 65.8, "20": 65.1}, "t": {"5": 0.9, "10": -0.72, "20": 0.19}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 64.8, "10": 61.5, "20": 62.6}, "t": {"5": 1.85, "10": 1.42, "20": 1.49}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 67.1, "10": 68.9, "20": 64.3}, "t": {"5": 0.49, "10": -0.46, "20": -0.62}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 60.2}, "t": {"10": 1.07}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 66.4, "10": 62.2, "20": 68.1}, "t": {"5": 0.65, "10": -0.03, "20": 0.54}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 61.4, "10": 67.3, "20": 63.7}, "t": {"5": 1.57, "10": 2.74, "20": 3.07}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"10": 60.4}, "t": {"10": 1.24}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"10": 60.7, "20": 60.7}, "t": {"10": 0.73, "20": 1.08}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": {"win": {"10": 60.1}, "t": {"10": 0.96}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"10": 62.3}, "t": {"10": 0.62}},
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 64.7, "10": 62.4}, "t": {"5": 2.33, "10": 1.95}},
    "大盤跌破60日均線 ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 61.4, "20": 64.6}, "t": {"10": 1.03, "20": 2.62}},
    "大盤跌破60日均線 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 60.3, "20": 64.8}, "t": {"10": -0.51, "20": 0.98}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 跌深反彈盤": {"win": {"10": 62.4, "20": 64.1}, "t": {"10": -2.31, "20": -0.9}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收)": {"win": {"20": 61.2}, "t": {"20": 3.64}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為正": {"win": {"10": 60.0, "20": 61.3}, "t": {"10": 10.93, "20": 11.91}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位)": {"win": {"20": 60.0}, "t": {"20": 1.85}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"20": 60.3}, "t": {"20": -0.12}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"20": 60.7}, "t": {"20": 1.53}},
    "大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 61.8, "10": 61.8, "20": 60.5}, "t": {"5": -0.4, "10": -2.23, "20": -1.68}},
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 64.5, "10": 63.2, "20": 60.5}, "t": {"5": -0.29, "10": -1.88, "20": -2.79}},
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月均價YoY為正": {"win": {"5": 71.3, "10": 64.9, "20": 64.9}, "t": {"5": 0.43, "10": -2.07, "20": -0.99}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 76.1, "10": 79.6, "20": 73.3}, "t": {"5": -0.15, "10": -1.0, "20": -0.36}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 84.8, "10": 87.9, "20": 87.9}, "t": {"5": 0.95, "10": 0.63, "20": 0.14}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 77.6, "10": 81.2, "20": 76.7}, "t": {"5": 0.06, "10": -0.41, "20": -0.55}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 81.9, "10": 84.0, "20": 81.3}, "t": {"5": 0.04, "10": -0.68, "20": -0.16}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 70.1, "10": 73.8, "20": 63.6}, "t": {"5": 0.05, "10": -0.59, "20": -0.12}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"win": {"5": 69.9, "10": 74.1, "20": 64.3}, "t": {"5": -0.8, "10": -1.38, "20": -1.03}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"win": {"5": 83.9, "10": 86.6, "20": 84.8}, "t": {"5": 0.73, "10": 0.21, "20": 0.76}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 75.2, "10": 79.5, "20": 74.4}, "t": {"5": -1.29, "10": -1.84, "20": -1.41}},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為正(強於大盤)": {"win": {"5": 66.6, "10": 63.0, "20": 64.0}, "t": {"5": -6.36, "10": -5.92, "20": -6.71}},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 60.3, "20": 63.2}, "t": {"5": 0.86, "20": 1.75}},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 相對強弱為正(強於大盤)": {"win": {"5": 72.2, "10": 65.2, "20": 68.3}, "t": {"5": -5.33, "10": -3.87, "20": -6.06}},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 68.4}, "t": {"5": 0.15}},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能區間高檔(≥90百分位)": {"win": {"5": 69.8, "10": 66.6, "20": 67.9}, "t": {"5": -3.91, "10": -3.0, "20": -5.43}},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 61.2}, "t": {"5": -5.3}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 11.5}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.6}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 10.1}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 11.6}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"moon": {"20": 10.9}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 14.1}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.9}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.8}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 12.2}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.4}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 10.3, "20": 16.1}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 11.2}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.0}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.9}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 11.2}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.8}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.5}},
    "強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.9}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.4}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 11.6}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 14.3}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.4}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 13.3}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.6}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 17.6, "20": 20.6}},
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 16.0, "20": 12.0}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.1}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 61.1, "10": 60.3, "20": 62.7}, "t": {"5": 1.34, "10": 2.12, "20": 2.79}, "moon": {"20": 10.3}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 18.6}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 11.1}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.0}},
    "晨星剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 12.1}},
    "晨星剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 15.7}},
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 14.0}},
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.5}},
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.4}},
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 13.9}},
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.3}},
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.6}},
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 13.6}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 60.4}, "t": {"5": 1.58}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 63.4, "10": 60.2, "20": 61.5}, "t": {"5": -0.01, "10": -0.12, "20": 0.33}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 60.1, "10": 60.7}, "t": {"5": 2.23, "10": 2.35}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為負": {"win": {"5": 62.9, "20": 61.8}, "t": {"5": -1.08, "20": -1.52}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 64.8, "10": 62.3, "20": 63.6}, "t": {"5": 1.05, "10": -0.03, "20": 0.06}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 12.5}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 20.0}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.3}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 13.3}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.2}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 13.0}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.3}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 17.9}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 12.7}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 15.3}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 12.0}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 60.3}, "t": {"5": 2.33}, "moon": {"10": 11.5, "20": 20.5}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 15.7}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 15.2}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 14.2}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 16.0}},
    "晨星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.0}},
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 近3月均價YoY為負": {"moon": {"20": 12.8}},
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 11.0}},
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 13.0}},
    "晨星剛形成 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.8}},
    "晨星剛形成 ＋ 近3月均價YoY為負 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.3}},
    "晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.6}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.9}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.2}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 60.7}, "t": {"20": 1.72}, "moon": {"20": 13.1}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 12.1}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.7}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.2}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月均價YoY為正": {"moon": {"20": 10.2}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月均價YoY為負": {"moon": {"20": 10.4}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 11.2}},
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.3}},
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 60.0}, "t": {"10": 1.72}},
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.3}},
    "母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.6}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.7}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 11.1}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 18.9}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 13.0}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 17.4}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 12.2}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 13.8}},
    "母子懷抱(高檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.1}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 距52週高點≤5%": {"win": {"5": 52.2, "10": 53.4, "20": 55.7}, "t": {"5": 6.49, "10": 7.85, "20": 9.5}, "moon": {"10": 4.1, "20": 10.3}},
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"win": {"5": 58.8, "10": 59.2, "20": 64.4}, "t": {"5": 5.08, "10": 5.5, "20": 6.38}, "moon": {"10": 11.6, "20": 21.7}},
    "外資近5日買超 ＋ 月營收創24個月新高 ＋ 近3日漲停": {"win": {"5": 54.7, "10": 55.8, "20": 57.0}, "t": {"5": 9.96, "10": 11.37, "20": 13.21}, "moon": {"10": 7.9, "20": 15.3}},
    "大盤站上20日均線 ＋ 月營收創24個月新高 ＋ 漲時量≥跌時量1.5倍(近20日)": {"win": {"5": 48.9, "10": 50.5, "20": 51.3}, "t": {"5": 12.03, "10": 18.15, "20": 22.56}, "moon": {"10": 3.9, "20": 8.2}},
    "大盤站上20日均線 ＋ 月營收創24個月新高 ＋ 近20日強於大盤≥10%": {"win": {"5": 49.9, "10": 51.8, "20": 52.2}, "t": {"5": 10.84, "10": 15.2, "20": 16.44}, "moon": {"10": 5.7, "20": 11.4}},

}

# 高勝率指定組合（#編號）與歷史勝率記錄（舊版定義下的回測數字）

PINNED_COMBOS = [
    ["布林通道低檔(≤20%)", "爆量(≥1.5倍均量)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "近3月均價YoY為負", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "近3月乖離度為正(營收優於股價)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "KDJ近3日內黃金交叉", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "三大法人近3月買超", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "相對強弱為負(弱於大盤)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "大盤跌破20日均線", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "型態成形中", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "型態突破確認", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "型態剛形成(剛突破)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "三大法人近3月賣超", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "近3月均價YoY為正", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "相對強弱為正(強於大盤)", "爆量(≥2倍均量)"],
    ["相對強弱為負(弱於大盤)", "大盤跌破20日均線", "晨星剛形成"],
    ["大盤跌破60日均線", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["跌深反彈盤", "大盤跌破60日均線", "近3月均價YoY為正"],
    ["跌深反彈盤", "大盤跌破60日均線", "三大法人近3月買超"],
    ["布林通道低檔(≤20%)", "爆量(≥2倍均量)", "大盤跌破60日均線"],
    ["大盤跌破20日均線", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "爆量(≥2倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)"],
    ["布林通道低檔(≤20%)", "相對強弱為正(強於大盤)", "量能區間高檔(≥90百分位)"],
    ["布林通道低檔(≤20%)", "相對強弱為正(強於大盤)", "夜星剛形成"],
    ["大盤跌破60日均線", "三大法人近3月買超", "晨星剛形成"],
    ["大盤跌破60日均線", "近3月均價YoY為正", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "大盤跌破60日均線", "母子懷抱(低檔)剛形成"],
    ["大盤跌破60日均線", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["爆量(≥1.5倍均量)", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "量能區間高檔(≥90百分位)", "大盤跌破60日均線"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "爆量(≥1.5倍均量)", "大盤跌破60日均線"],
    ["布林通道低檔(≤20%)", "爆量(≥2倍均量)", "型態剛形成(剛突破)"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "近3月均價YoY為負", "夜星剛形成"],
    ["大盤跌破60日均線", "晨星剛形成"],
    ["大盤跌破60日均線", "型態成形中", "晨星剛形成"],
    ["大盤跌破60日均線", "型態突破確認", "晨星剛形成"],
    ["大盤跌破60日均線", "型態剛形成(剛突破)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "爆量(≥2倍均量)", "大盤跌破20日均線"],
    ["跌深反彈盤", "大盤跌破20日均線", "大盤跌破60日均線"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內黃金交叉", "大盤跌破20日均線", "晨星剛形成"],
    ["頭肩底剛形成", "三重底剛形成", "一字底(均線糾結)剛形成"],
    ["圓弧底剛形成", "一字底(均線糾結)剛形成", "突破飆股大量黑K最高點剛形成"],
    ["地量(≤0.5倍均量)", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "大盤跌破60日均線", "夜星剛形成"],
    ["爆量(≥1.5倍均量)", "大盤跌破60日均線", "母子懷抱(低檔)剛形成"],
    ["布林通道低檔(≤20%)", "相對強弱為正(強於大盤)", "爆量(≥1.5倍均量)"],
    ["大盤跌破60日均線", "近3月均價YoY為負", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "爆量(≥2倍均量)", "型態突破確認"],
    ["爆量(≥1.5倍均量)", "大盤跌破60日均線", "母子懷抱(高檔)剛形成"],
    ["MACD近3日內黃金交叉", "大盤跌破60日均線", "夜星剛形成"],
    ["爆量(≥1.5倍均量)", "大盤跌破20日均線", "晨星剛形成"],
    ["大盤跌破20日均線", "近3月均價YoY為負", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "大盤跌破20日均線", "晨星剛形成"],
    ["大盤跌破60日均線", "三大法人近3月賣超", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破20日均線", "夜星剛形成"],
    ["大盤站上20日均線", "大盤跌破60日均線", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "大盤跌破60日均線", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "相對強弱為負(弱於大盤)", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "爆量(≥2倍均量)", "大盤跌破60日均線"],
    ["大盤跌破60日均線", "近3月乖離度為正(營收優於股價)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破60日均線"],
    ["布林通道低檔(≤20%)", "大盤跌破20日均線", "夜星剛形成"],
    ["大盤跌破20日均線", "近3月乖離度為正(營收優於股價)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "量能區間高檔(≥90百分位)", "大盤跌破20日均線"],
    ["回後買上漲全通過", "一字底(均線糾結)剛形成", "突破飆股大量黑K最高點剛形成"],
    ["布林通道低檔(≤20%)", "爆量(≥1.5倍均量)", "大盤跌破20日均線"],
    ["相對強弱為負(弱於大盤)", "爆量(≥1.5倍均量)", "大盤跌破60日均線"],
    ["相對強弱為負(弱於大盤)", "大盤跌破20日均線", "大盤跌破60日均線"],
    ["大盤跌破60日均線", "三大法人近3月買超", "夜星剛形成"],
    ["跌深反彈盤", "KDJ近3日內死亡交叉", "大盤跌破60日均線"],
    ["跌深反彈盤", "大盤跌破60日均線", "近3月乖離度為負(股價超前營收)"],
    ["相對強弱為負(弱於大盤)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "晨星剛形成"],
    ["多方力道≥65", "相對強弱為負(弱於大盤)", "大盤跌破60日均線"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月乖離度為負(股價超前營收)", "夜星剛形成"],
    ["布林通道低檔(≤20%)", "爆量(≥1.5倍均量)", "型態剛形成(剛突破)"],
    ["近3月均價YoY為正", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["大盤跌破20日均線", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["大盤站上20日均線", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["跌深反彈盤", "大盤跌破20日均線", "三大法人近3月買超"],
    ["MACD近3日內黃金交叉", "KDJ近3日內死亡交叉", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "量能區間低檔(≤10百分位)", "大盤跌破60日均線"],
    ["地量(≤0.5倍均量)", "大盤跌破60日均線", "回後買上漲全通過"],
    ["布林通道低檔(≤20%)", "爆量(≥1.5倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)"],
    ["MACD近3日內黃金交叉", "KDJ近3日內黃金交叉", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["相對強弱為負(弱於大盤)", "近3月乖離度為正(營收優於股價)", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "大盤跌破60日均線", "型態剛形成(剛突破)"],
    ["多方力道≥80", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["大盤站上20日均線", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["布林通道低檔(≤20%)", "KDJ近3日內死亡交叉", "大盤跌破60日均線"],
    ["布林通道低檔(≤20%)", "KDJ近3日內死亡交叉", "夜星剛形成"],
    ["大盤跌破20日均線", "三大法人近3月買超", "夜星剛形成"],
    ["大盤站上60日均線", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破20日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "爆量(≥1.5倍均量)", "型態突破確認"],
    ["相對強弱為負(弱於大盤)", "三大法人近3月買超", "晨星剛形成"],
    ["跌深反彈盤", "相對強弱為負(弱於大盤)", "大盤跌破60日均線"],
    ["MACD近3日內黃金交叉", "KDJ近3日內死亡交叉", "N字底剛形成"],
    ["相對強弱為負(弱於大盤)", "近3月均價YoY為負", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "大盤跌破60日均線", "型態突破確認"],
    ["MACD近3日內黃金交叉", "近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "大盤跌破60日均線", "母子懷抱(低檔)剛形成"],
    ["相對強弱為負(弱於大盤)", "爆量(≥2倍均量)", "大盤跌破20日均線"],
    ["大盤跌破20日均線", "三大法人近3月買超", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "量能區間低檔(≤10百分位)", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "一字底(均線糾結)剛形成", "突破飆股大量黑K最高點剛形成"],
    ["頭肩底剛形成", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["布林通道低檔(≤20%)", "相對強弱為負(弱於大盤)", "大盤跌破60日均線"],
    ["MACD近3日內黃金交叉", "KDJ近3日內死亡交叉", "量能斜率轉弱(近5日均量<近10日均量20%以上)"],
    ["MACD近3日內黃金交叉", "KDJ近3日內死亡交叉", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "相對強弱為負(弱於大盤)", "大盤跌破60日均線"],
    ["相對強弱為負(弱於大盤)", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["強勢突破盤", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["跌深反彈盤", "大盤跌破60日均線"],
    ["跌深反彈盤", "布林通道低檔(≤20%)", "大盤跌破60日均線"],
    ["跌深反彈盤", "MACD近3日內黃金交叉", "大盤跌破60日均線"],
    ["布林通道低檔(≤20%)", "MACD近3日內黃金交叉", "大盤跌破60日均線"],
    ["相對強弱為正(強於大盤)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "三大法人近3月買超", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "爆量(≥1.5倍均量)", "大盤跌破20日均線"],
    ["相對強弱為負(弱於大盤)", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "近3月均價YoY為正", "晨星剛形成"],
    ["多方力道≥80", "布林通道高檔(≥80%)", "量能區間低檔(≤10百分位)"],
    ["MACD近3日內黃金交叉", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "N字底剛形成"],
    ["近3月乖離度為負(股價超前營收)", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["布林通道低檔(≤20%)", "相對強弱為正(強於大盤)", "量能斜率轉強(近5日均量>近10日均量20%以上)"],
    ["爆量(≥2倍均量)", "大盤跌破60日均線", "晨星剛形成"],
    ["大盤跌破60日均線", "近3月均價YoY為正", "夜星剛形成"],
    ["近3月均價YoY為正", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "地量(≤0.5倍均量)", "大盤跌破60日均線"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "回後買上漲全通過"],
    ["布林通道低檔(≤20%)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "大盤跌破60日均線"],
    ["KDJ近3日內死亡交叉", "相對強弱為負(弱於大盤)", "大盤跌破60日均線"],
    ["量能區間高檔(≥90百分位)", "大盤跌破20日均線", "複式頭肩底剛形成"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "近3月均價YoY為正"],
    ["量能區間高檔(≥90百分位)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤跌破20日均線", "晨星剛形成"],
    ["大盤跌破20日均線", "型態成形中", "晨星剛形成"],
    ["大盤跌破20日均線", "型態突破確認", "晨星剛形成"],
    ["大盤跌破20日均線", "型態剛形成(剛突破)", "晨星剛形成"],
    ["跌深反彈盤", "大盤跌破60日均線", "近3月乖離度為正(營收優於股價)"],
    ["布林通道高檔(≥80%)", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["KDJ近3日內黃金交叉", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["KDJ近3日內黃金交叉", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["跌深反彈盤", "大盤跌破60日均線", "型態成形中"],
    ["布林通道低檔(≤20%)", "近3月乖離度為正(營收優於股價)", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "大盤跌破20日均線", "晨星剛形成"],
    ["近3月乖離度為負(股價超前營收)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破60日均線", "母子懷抱(低檔)剛形成"],
    ["跌深反彈盤", "KDJ近3日內黃金交叉", "大盤跌破20日均線"],
    ["跌深反彈盤", "KDJ近3日內黃金交叉", "近3月乖離度為正(營收優於股價)"],
    ["MACD近3日內黃金交叉", "KDJ近3日內死亡交叉", "大盤跌破60日均線"],
    ["量能區間低檔(≤10百分位)", "大盤跌破60日均線", "母子懷抱(高檔)剛形成"],
    ["大盤站上60日均線", "圓弧底剛形成", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "夜星剛形成"],
    ["量能區間低檔(≤10百分位)", "大盤站上20日均線", "回後買上漲全通過"],
    ["MACD近3日內黃金交叉", "相對強弱為正(強於大盤)", "母子懷抱(低檔)剛形成"],
    ["爆量(≥1.5倍均量)", "大盤跌破20日均線", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "量能斜率轉強(近5日均量>近10日均量20%以上)", "晨星剛形成"],
    ["爆量(≥2倍均量)", "大盤跌破20日均線", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "量能區間低檔(≤10百分位)", "N字底剛形成"],
    ["多方力道≥80", "強勢突破盤", "地量(≤0.5倍均量)"],
    ["爆量(≥1.5倍均量)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["跌深反彈盤", "大盤跌破20日均線", "近3月均價YoY為正"],
    ["相對強弱為負(弱於大盤)", "近3月均價YoY為正", "晨星剛形成"],
    ["大盤跌破20日均線", "三大法人近3月賣超", "晨星剛形成"],
    ["大盤跌破20日均線", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "大盤跌破20日均線", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內黃金交叉", "三大法人近3月買超", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破20日均線"],
    ["相對強弱為負(弱於大盤)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "大盤跌破60日均線"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "近3月乖離度為負(股價超前營收)"],
    ["三大法人近3月買超", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["布林通道低檔(≤20%)", "MACD近3日內黃金交叉", "量能區間低檔(≤10百分位)"],
    ["大盤跌破60日均線", "三大法人近3月賣超", "母子懷抱(低檔)剛形成"],
    ["K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "大盤跌破20日均線", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "相對強弱為負(弱於大盤)", "大盤跌破60日均線"],
    ["相對強弱為負(弱於大盤)", "爆量(≥2倍均量)", "母子懷抱(低檔)剛形成"],
    ["型態成形中", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["型態突破確認", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["型態剛形成(剛突破)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤跌破60日均線", "近3月乖離度為正(營收優於股價)", "夜星剛形成"],
    ["布林通道低檔(≤20%)", "大盤跌破60日均線", "近3月乖離度為負(股價超前營收)"],
    ["布林通道高檔(≥80%)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "量能區間低檔(≤10百分位)", "大盤跌破20日均線"],
    ["相對強弱為負(弱於大盤)", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "型態成形中", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "型態突破確認", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "型態剛形成(剛突破)", "晨星剛形成"],
    ["N字底剛形成", "一字底(均線糾結)剛形成", "K線橫盤的突破剛形成"],
    ["頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["型態成形中", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["型態突破確認", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["型態剛形成(剛突破)", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["相對強弱為負(弱於大盤)", "大盤跌破20日均線", "夜星剛形成"],
    ["大盤跌破60日均線", "母子懷抱(高檔)剛形成", "夜星剛形成"],
    ["布林通道低檔(≤20%)", "三大法人近3月買超", "夜星剛形成"],
    ["大盤跌破60日均線", "近3月乖離度為正(營收優於股價)", "母子懷抱(低檔)剛形成"],
    ["大盤跌破60日均線", "近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成"],
    ["三重底剛形成", "圓弧底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["跌深反彈盤", "大盤跌破60日均線", "三大法人近3月賣超"],
    ["大盤跌破60日均線", "三大法人近3月買超", "母子懷抱(低檔)剛形成"],
    ["相對強弱為負(弱於大盤)", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["近3月均價YoY為正", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "KDJ近3日內黃金交叉", "大盤跌破60日均線"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "母子懷抱(高檔)剛形成"],
    ["相對強弱為負(弱於大盤)", "量能區間高檔(≥90百分位)", "大盤跌破60日均線"],
    ["相對強弱為負(弱於大盤)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破60日均線"],
    ["大盤站上20日均線", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["大盤站上60日均線", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤跌破20日均線", "大盤跌破60日均線", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "型態成形中", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "型態突破確認", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "型態剛形成(剛突破)", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "三重底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "N字底剛形成"],
    ["大盤跌破20日均線", "近3月均價YoY為正", "晨星剛形成"],
    ["近3月乖離度為負(股價超前營收)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "三大法人近3月買超"],
    ["布林通道低檔(≤20%)", "大盤跌破60日均線", "三大法人近3月買超"],
    ["大盤跌破60日均線", "近3月均價YoY為負", "母子懷抱(低檔)剛形成"],
    ["布林通道低檔(≤20%)", "爆量(≥2倍均量)", "三大法人近3月買超"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["三大法人近3月買超", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["KDJ近3日內死亡交叉", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破60日均線"],
    ["大盤跌破60日均線", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "N字底剛形成", "一字底(均線糾結)剛形成"],
    ["KDJ近3日內黃金交叉", "近3月乖離度為正(營收優於股價)", "晨星剛形成"],
    ["大盤跌破60日均線", "型態成形中", "母子懷抱(低檔)剛形成"],
    ["大盤跌破60日均線", "型態突破確認", "母子懷抱(低檔)剛形成"],
    ["大盤跌破60日均線", "型態剛形成(剛突破)", "母子懷抱(低檔)剛形成"],
    ["布林通道低檔(≤20%)", "大盤跌破60日均線", "母子懷抱(高檔)剛形成"],
    ["爆量(≥1.5倍均量)", "頭肩底剛形成", "一字底(均線糾結)剛形成"],
    ["量能區間低檔(≤10百分位)", "大盤跌破20日均線", "母子懷抱(高檔)剛形成"],
    ["布林通道低檔(≤20%)", "大盤跌破20日均線", "大盤跌破60日均線"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "型態成形中"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "型態剛形成(剛突破)"],
    ["大盤站上20日均線", "大盤跌破60日均線", "母子懷抱(高檔)剛形成"],
    ["多方力道≥65", "頭肩底剛形成", "三重底剛形成"],
    ["大盤跌破20日均線", "圓弧底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["多方力道≥80", "相對強弱為負(弱於大盤)", "突破飆股大量黑K最高點剛形成"],
    ["KDJ近3日內黃金交叉", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "量能區間低檔(≤10百分位)", "母子懷抱(低檔)剛形成"],
    ["跌深反彈盤", "爆量(≥1.5倍均量)", "三大法人近3月賣超"],
    ["布林通道低檔(≤20%)", "大盤跌破60日均線", "近3月均價YoY為正"],
    ["相對強弱為正(強於大盤)", "N字底剛形成", "一字底(均線糾結)剛形成"],
    # 2026-09-30 三年回測：大盤跌深時弱勢股開始跳空上漲（擇時型，三段 t 都 ≥2）
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "近3日向上跳空缺口"],
]

PINNED_COMBO_WINRATES = {
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 晨星剛形成": {"5": 83.1, "10": 86.4, "20": 86.4},
    "布林通道低檔(≤20%) ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"5": 82.8, "10": 85.9, "20": 85.9},
    "布林通道低檔(≤20%) ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成": {"5": 81.1, "10": 83.5, "20": 81.1},
    "布林通道低檔(≤20%) ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成": {"5": 77.9, "10": 82.4, "20": 74.4},
    "布林通道低檔(≤20%) ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"5": 76.5, "10": 82.4, "20": 77.3},
    "布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 78.9, "10": 82.3, "20": 77},
    "布林通道低檔(≤20%) ＋ 相對強弱為負(弱於大盤) ＋ 晨星剛形成": {"5": 77.7, "10": 81.7, "20": 77.2},
    "布林通道低檔(≤20%) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"5": 76.3, "10": 80.4, "20": 74.4},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 79.7, "10": 80, "20": 76.4},
    "布林通道低檔(≤20%) ＋ 晨星剛形成": {"5": 75.7, "10": 79.6, "20": 73.5},
    "布林通道低檔(≤20%) ＋ 型態成形中 ＋ 晨星剛形成": {"5": 75.7, "10": 79.6, "20": 73.5},
    "布林通道低檔(≤20%) ＋ 型態突破確認 ＋ 晨星剛形成": {"5": 75.7, "10": 79.6, "20": 73.5},
    "布林通道低檔(≤20%) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"5": 75.7, "10": 79.6, "20": 73.5},
    "布林通道低檔(≤20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成": {"5": 74.8, "10": 78.5, "20": 73.8},
    "布林通道低檔(≤20%) ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"5": 74.8, "10": 76.6, "20": 69.2},
    "布林通道低檔(≤20%) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"5": 69.8, "10": 75, "20": 63.5},
    "布林通道低檔(≤20%) ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"5": 70.1, "10": 74.8, "20": 63.8},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 73, "10": 73.8, "20": 66.9},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 爆量(≥2倍均量)": {"5": 73.6, "10": 65.6, "20": 68.9},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"5": 73.4, "10": 73.6, "20": 73.2},
    "大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"5": 72.3, "10": 73.6, "20": 65.1},
    "跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正": {"5": 72.6, "10": 63.1, "20": 63.1},
    "跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 三大法人近3月買超": {"5": 71.7, "10": 66, "20": 60.4},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 大盤跌破60日均線": {"5": 68.5, "10": 66.3, "20": 71.5},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 70.7, "10": 71, "20": 65.2},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"5": 70.6, "10": 60.8},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能區間高檔(≥90百分位)": {"5": 70.2, "10": 66.9, "20": 68.2},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 夜星剛形成": {"5": 70, "10": 64, "20": 60},
    "大盤跌破60日均線 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"5": 69.5, "10": 70, "20": 61.6},
    "大盤跌破60日均線 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"5": 68.6, "10": 69.8, "20": 61},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"5": 66.6, "10": 69.5, "20": 67.5},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"5": 64.9, "10": 65.6, "20": 69.5},
    "爆量(≥1.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 64.9, "10": 63.5, "20": 69.2},
    "布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位) ＋ 大盤跌破60日均線": {"5": 68.4, "10": 66.7, "20": 68.9},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 67.8, "10": 68.9, "20": 64.4},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 大盤跌破60日均線": {"5": 65.6, "10": 65.7, "20": 68.8},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 型態剛形成(剛突破)": {"5": 68.7},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 68.4, "10": 64.6, "20": 64.6},
    "布林通道低檔(≤20%) ＋ 近3月均價YoY為負 ＋ 夜星剛形成": {"5": 68.3},
    "大盤跌破60日均線 ＋ 晨星剛形成": {"5": 67.6, "10": 68.3, "20": 62.4},
    "大盤跌破60日均線 ＋ 型態成形中 ＋ 晨星剛形成": {"5": 67.6, "10": 68.3, "20": 62.4},
    "大盤跌破60日均線 ＋ 型態突破確認 ＋ 晨星剛形成": {"5": 67.6, "10": 68.3, "20": 62.4},
    "大盤跌破60日均線 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"5": 67.6, "10": 68.3, "20": 62.4},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 大盤跌破20日均線": {"5": 66.6, "10": 63.9, "20": 68.3},
    "跌深反彈盤 ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線": {"5": 67.8, "10": 65.3, "20": 61.9},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"5": 62.3, "10": 67.5, "20": 64.2},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"5": 66.6, "10": 67.3, "20": 64.8},
    "頭肩底剛形成 ＋ 三重底剛形成 ＋ 一字底(均線糾結)剛形成": {"5": 61.5, "10": 67.3},
    "圓弧底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"10": 67.3},
    "地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 63.8, "10": 67.2},
    "布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"5": 67, "10": 60.4, "20": 63.2},
    "爆量(≥1.5倍均量) ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"5": 66, "10": 61.2, "20": 67},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 爆量(≥1.5倍均量)": {"5": 66.8, "10": 62.7, "20": 64.3},
    "大盤跌破60日均線 ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"5": 66.8, "10": 66, "20": 64.8},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 型態突破確認": {"5": 66.7},
    "爆量(≥1.5倍均量) ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": {"5": 66.7, "10": 62.5},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"5": 61.1, "20": 66.7},
    "爆量(≥1.5倍均量) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"5": 61.6, "10": 60.5, "20": 66.5},
    "大盤跌破20日均線 ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"5": 66.4, "10": 65.7, "20": 64.6},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"5": 66.3, "10": 65.7, "20": 62.9},
    "大盤跌破60日均線 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"5": 65.3, "10": 66.2, "20": 63.3},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"5": 66},
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"10": 65.9, "20": 63},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 63.2, "10": 65.6},
    "KDJ近3日內黃金交叉 ＋ 相對強弱為負(弱於大盤) ＋ 晨星剛形成": {"5": 64.5, "10": 65.4, "20": 61.3},
    "相對強弱為負(弱於大盤) ＋ 爆量(≥2倍均量) ＋ 大盤跌破60日均線": {"10": 60.7, "20": 65.4},
    "大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成": {"5": 65.3, "10": 65.3, "20": 61.7},
    "MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線": {"5": 64.4, "10": 65.3, "20": 63.9},
    "布林通道低檔(≤20%) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"5": 65.2},
    "大盤跌破20日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成": {"5": 64.9, "10": 65.2, "20": 60.9},
    "布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位) ＋ 大盤跌破20日均線": {"5": 64.4, "10": 62.6, "20": 65.2},
    "回後買上漲全通過 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"10": 65.2},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 大盤跌破20日均線": {"5": 62.9, "10": 62.5, "20": 65.1},
    "相對強弱為負(弱於大盤) ＋ 爆量(≥1.5倍均量) ＋ 大盤跌破60日均線": {"10": 61.7, "20": 65.1},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線": {"5": 60.6, "10": 63.4, "20": 65},
    "大盤跌破60日均線 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"5": 64.8, "10": 60.3, "20": 60.3},
    "跌深反彈盤 ＋ KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線": {"5": 64.6, "10": 62.6, "20": 63.6},
    "跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)": {"5": 64.6, "10": 60},
    "相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成": {"5": 64.6, "10": 61.5, "20": 61.9},
    "多方力道≥65 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線": {"10": 64.5, "20": 60.8},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成": {"5": 64.4, "10": 61.5},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 型態剛形成(剛突破)": {"5": 64.3, "10": 61.7, "20": 63.6},
    "近3月均價YoY為正 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 61.4, "20": 64.2},
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": 64},
    "大盤站上20日均線 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 60.9, "20": 63.8},
    "跌深反彈盤 ＋ 大盤跌破20日均線 ＋ 三大法人近3月買超": {"5": 63.7},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成": {"5": 63.7},
    "MACD近3日內黃金交叉 ＋ 量能區間低檔(≤10百分位) ＋ 大盤跌破60日均線": {"5": 63.1, "10": 63.6, "20": 62.2},
    "地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 回後買上漲全通過": {"10": 61.5, "20": 63.5},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"5": 60.4, "20": 63.4},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"10": 63.4, "20": 61.3},
    "KDJ近3日內死亡交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 63.3, "20": 63.3},
    "相對強弱為負(弱於大盤) ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成": {"5": 63.2, "10": 60.8, "20": 61.7},
    "布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 型態剛形成(剛突破)": {"5": 61.9, "10": 63.2, "20": 62.2},
    "多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"5": 63.1},
    "大盤站上20日均線 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 63.1},
    "布林通道低檔(≤20%) ＋ KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線": {"5": 63, "10": 61, "20": 62.8},
    "布林通道低檔(≤20%) ＋ KDJ近3日內死亡交叉 ＋ 夜星剛形成": {"5": 63, "20": 61.6},
    "大盤跌破20日均線 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"5": 62.9, "20": 60.2},
    "大盤站上60日均線 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 62.9, "20": 61.3},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"20": 62.9},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 型態突破確認": {"5": 62.8, "10": 61, "20": 62.8},
    "相對強弱為負(弱於大盤) ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"5": 61.7, "10": 62.7},
    "跌深反彈盤 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線": {"10": 60.9, "20": 62.7},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ N字底剛形成": {"20": 62.7},
    "相對強弱為負(弱於大盤) ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"5": 62.6, "20": 61.3},
    "布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 型態突破確認": {"5": 61.7, "10": 62.6, "20": 62},
    "MACD近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成": {"10": 62.6},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"10": 62.6},
    "相對強弱為負(弱於大盤) ＋ 爆量(≥2倍均量) ＋ 大盤跌破20日均線": {"20": 62.6},
    "大盤跌破20日均線 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"5": 62.5, "10": 61.9},
    "KDJ近3日內死亡交叉 ＋ 量能區間低檔(≤10百分位) ＋ 母子懷抱(低檔)剛形成": {"10": 62.5},
    "KDJ近3日內死亡交叉 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"10": 62.5},
    "頭肩底剛形成 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 62.5, "20": 60.9},
    "布林通道低檔(≤20%) ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線": {"20": 62.4},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"5": 60.9, "10": 62.3, "20": 62},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成": {"20": 62.3},
    "MACD近3日內黃金交叉 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線": {"10": 60.7, "20": 62.2},
    "相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"5": 61, "10": 62.1},
    "強勢突破盤 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 60.3, "20": 62.1},
    "跌深反彈盤 ＋ 大盤跌破60日均線": {"5": 62, "10": 61.3, "20": 60.6},
    "跌深反彈盤 ＋ 布林通道低檔(≤20%) ＋ 大盤跌破60日均線": {"5": 62, "10": 61.3, "20": 60.6},
    "跌深反彈盤 ＋ MACD近3日內黃金交叉 ＋ 大盤跌破60日均線": {"5": 62, "10": 61.3, "20": 60.6},
    "布林通道低檔(≤20%) ＋ MACD近3日內黃金交叉 ＋ 大盤跌破60日均線": {"5": 62, "10": 61.9, "20": 61.9},
    "相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"5": 62},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"5": 62, "10": 60.2},
    "KDJ近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"10": 62},
    "相對強弱為負(弱於大盤) ＋ 爆量(≥1.5倍均量) ＋ 大盤跌破20日均線": {"20": 62},
    "相對強弱為負(弱於大盤) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"5": 61, "10": 60.5, "20": 61.9},
    "KDJ近3日內黃金交叉 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"5": 60.1, "10": 61.9},
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 量能區間低檔(≤10百分位)": {"10": 61.9},
    "MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ N字底剛形成": {"10": 61.9},
    "近3月乖離度為負(股價超前營收) ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 61.9, "20": 61.9},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"5": 61.8},
    "爆量(≥2倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"5": 61.8, "20": 60.9},
    "大盤跌破60日均線 ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"5": 61.8},
    "近3月均價YoY為正 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"5": 61.8},
    "MACD近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": {"10": 61.8, "20": 60.3},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過": {"20": 61.8},
    "布林通道低檔(≤20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破60日均線": {"20": 61.8},
    "KDJ近3日內死亡交叉 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線": {"5": 61.7, "20": 61.7},
    "量能區間高檔(≥90百分位) ＋ 大盤跌破20日均線 ＋ 複式頭肩底剛形成": {"10": 61.7},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正": {"10": 60.3, "20": 61.7},
    "量能區間高檔(≥90百分位) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 61.7},
    "大盤跌破20日均線 ＋ 晨星剛形成": {"5": 60.9, "10": 61.6, "20": 60.2},
    "大盤跌破20日均線 ＋ 型態成形中 ＋ 晨星剛形成": {"5": 60.9, "10": 61.6, "20": 60.2},
    "大盤跌破20日均線 ＋ 型態突破確認 ＋ 晨星剛形成": {"5": 60.9, "10": 61.6, "20": 60.2},
    "大盤跌破20日均線 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"5": 60.9, "10": 61.6, "20": 60.2},
    "跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價)": {"10": 61.6, "20": 60.3},
    "布林通道高檔(≥80%) ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 60.3, "20": 61.6},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 61.6},
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": 61.6},
    "跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 型態成形中": {"5": 61.5, "10": 60.8, "20": 60.8},
    "布林通道低檔(≤20%) ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成": {"5": 61.5},
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"5": 61.5},
    "近3月乖離度為負(股價超前營收) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"5": 61.5},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"10": 61.5},
    "跌深反彈盤 ＋ KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線": {"20": 61.5},
    "跌深反彈盤 ＋ KDJ近3日內黃金交叉 ＋ 近3月乖離度為正(營收優於股價)": {"20": 61.5},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線": {"20": 61.5},
    "量能區間低檔(≤10百分位) ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": {"20": 61.5},
    "大盤站上60日均線 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": 61.5},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"5": 61.4, "20": 61.4},
    "量能區間低檔(≤10百分位) ＋ 大盤站上20日均線 ＋ 回後買上漲全通過": {"10": 61.4},
    "MACD近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 母子懷抱(低檔)剛形成": {"20": 61.4},
    "爆量(≥1.5倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"20": 61.4},
    "多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"5": 61.3},
    "KDJ近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成": {"5": 61.3},
    "爆量(≥2倍均量) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"5": 61.3},
    "地量(≤0.5倍均量) ＋ 量能區間低檔(≤10百分位) ＋ N字底剛形成": {"10": 61.3},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 地量(≤0.5倍均量)": {"10": 60.3, "20": 61.3},
    "爆量(≥1.5倍均量) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 61.3},
    "跌深反彈盤 ＋ 大盤跌破20日均線 ＋ 近3月均價YoY為正": {"5": 61.2},
    "相對強弱為負(弱於大盤) ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"5": 60.5, "10": 61.2},
    "大盤跌破20日均線 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"10": 61.2, "20": 60.5},
    "大盤跌破20日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"10": 60.8, "20": 61.2},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"20": 61.2},
    "KDJ近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"10": 61.1},
    "MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線": {"20": 61.1},
    "相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破60日均線": {"20": 61.1},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)": {"20": 61.1},
    "三大法人近3月買超 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 61.1},
    "布林通道低檔(≤20%) ＋ MACD近3日內黃金交叉 ＋ 量能區間低檔(≤10百分位)": {"5": 61},
    "大盤跌破60日均線 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成": {"10": 61},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 61},
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": 61},
    "KDJ近3日內黃金交叉 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線": {"20": 61},
    "相對強弱為負(弱於大盤) ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成": {"20": 61},
    "型態成形中 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 61},
    "型態突破確認 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 61},
    "型態剛形成(剛突破) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 61},
    "大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成": {"5": 60.9},
    "布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)": {"10": 60, "20": 60.9},
    "布林通道高檔(≥80%) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 60.9},
    "MACD近3日內黃金交叉 ＋ 量能區間低檔(≤10百分位) ＋ 大盤跌破20日均線": {"20": 60.9},
    "相對強弱為負(弱於大盤) ＋ 晨星剛形成": {"5": 60.8, "10": 60.4},
    "相對強弱為負(弱於大盤) ＋ 型態成形中 ＋ 晨星剛形成": {"5": 60.8, "10": 60.4},
    "相對強弱為負(弱於大盤) ＋ 型態突破確認 ＋ 晨星剛形成": {"5": 60.8, "10": 60.4},
    "相對強弱為負(弱於大盤) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"5": 60.8, "10": 60.4},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ K線橫盤的突破剛形成": {"10": 60, "20": 60.8},
    "頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 60.8},
    "型態成形中 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 60.8},
    "型態突破確認 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 60.8},
    "型態剛形成(剛突破) ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 60.8},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"5": 60.7},
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成": {"5": 60.7},
    "布林通道低檔(≤20%) ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"10": 60.7},
    "大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 母子懷抱(低檔)剛形成": {"10": 60.7},
    "大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成": {"10": 60.7, "20": 60.1},
    "三重底剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"10": 60.7},
    "跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 三大法人近3月賣超": {"20": 60.7},
    "大盤跌破60日均線 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"5": 60.6},
    "相對強弱為負(弱於大盤) ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 60.6},
    "近3月均價YoY為正 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"10": 60.6},
    "布林通道低檔(≤20%) ＋ KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線": {"10": 60.4, "20": 60.6},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": {"10": 60.2, "20": 60.6},
    "相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位) ＋ 大盤跌破60日均線": {"20": 60.6},
    "相對強弱為負(弱於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線": {"20": 60.6},
    "大盤站上20日均線 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 60.6},
    "大盤站上60日均線 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 60.6},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"5": 60.5},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成": {"10": 60.5},
    "KDJ近3日內黃金交叉 ＋ 型態成形中 ＋ 晨星剛形成": {"10": 60.5},
    "KDJ近3日內黃金交叉 ＋ 型態突破確認 ＋ 晨星剛形成": {"10": 60.5},
    "KDJ近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"10": 60.5},
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"10": 60.5},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ N字底剛形成": {"10": 60.5},
    "大盤跌破20日均線 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"10": 60.5},
    "近3月乖離度為負(股價超前營收) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"10": 60.5},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 三大法人近3月買超": {"10": 60.2, "20": 60.5},
    "布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 三大法人近3月買超": {"5": 60.3, "10": 60.4},
    "大盤跌破60日均線 ＋ 近3月均價YoY為負 ＋ 母子懷抱(低檔)剛形成": {"10": 60.4},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 三大法人近3月買超": {"20": 60.4},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": 60.4},
    "三大法人近3月買超 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 60.4},
    "KDJ近3日內死亡交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線": {"5": 60.3},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"10": 60.3},
    "多方力道≥65 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"10": 60.3},
    "KDJ近3日內黃金交叉 ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成": {"10": 60.3},
    "大盤跌破60日均線 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": {"10": 60.3},
    "大盤跌破60日均線 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": {"10": 60.3},
    "大盤跌破60日均線 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": {"10": 60.3},
    "布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": {"20": 60.3},
    "爆量(≥1.5倍均量) ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 60.3},
    "量能區間低檔(≤10百分位) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"20": 60.3},
    "布林通道低檔(≤20%) ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線": {"5": 60.2},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 型態成形中": {"20": 60.2},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 型態剛形成(剛突破)": {"10": 60.1},
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": {"10": 60.1, "20": 60.1},
    "多方力道≥65 ＋ 頭肩底剛形成 ＋ 三重底剛形成": {"5": 60},
    "大盤跌破20日均線 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"5": 60},
    "多方力道≥80 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"10": 60},
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"10": 60},
    "相對強弱為正(強於大盤) ＋ 量能區間低檔(≤10百分位) ＋ 母子懷抱(低檔)剛形成": {"10": 60},
    "跌深反彈盤 ＋ 爆量(≥1.5倍均量) ＋ 三大法人近3月賣超": {"20": 60},
    "布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正": {"20": 60},
    "相對強弱為正(強於大盤) ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"20": 60},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 近3日向上跳空缺口": {"5": 57.6, "10": 71.1, "20": 72.9},
}

# 高標股指定組合（M編號）與歷史飆股統計記錄

MOONSHOT_COMBOS = [
    ["突破飆股大量黑K最高點剛形成", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "突破ABC修正下降切線剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "回後買上漲全通過"],
    ["KDJ近3日內死亡交叉", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月乖離度為負(股價超前營收)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["多方力道≥80", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["強勢突破盤", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "晨星剛形成"],
    ["多方力道≥80", "強勢突破盤", "地量(≤0.5倍均量)"],
    ["三大法人近3月買超", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["三大法人近3月買超", "圓弧底剛形成", "晨星剛形成"],
    ["近3月乖離度為負(股價超前營收)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["突破飆股大量黑K最高點剛形成", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["大盤站上20日均線", "圓弧底剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "回後買上漲全通過", "突破ABC修正下降切線剛形成"],
    ["爆量(≥1.5倍均量)", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "母子懷抱(低檔)剛形成"],
    ["三大法人近3月賣超", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["多方力道≥65", "地量(≤0.5倍均量)", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["突破ABC修正下降切線剛形成", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["近3月乖離度為負(股價超前營收)", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["近3月均價YoY為正", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "三重底剛形成", "晨星剛形成"],
    ["近3月均價YoY為負", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["大盤站上20日均線", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["多方力道≥65", "地量(≤0.5倍均量)", "夜星剛形成"],
    ["強勢突破盤", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤站上60日均線", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "回後買上漲全通過"],
    ["爆量(≥1.5倍均量)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "三大法人近3月買超", "夜星剛形成"],
    ["大盤站上60日均線", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "三大法人近3月買超", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破60日均線", "夜星剛形成"],
    ["近3月均價YoY為負", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "相對強弱為負(弱於大盤)", "夜星剛形成"],
    ["三大法人近3月賣超", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["型態成形中", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["型態突破確認", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["型態剛形成(剛突破)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["三大法人近3月賣超", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "地量(≤0.5倍均量)", "夜星剛形成"],
    ["近3月均價YoY為正", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "圓弧底剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月均價YoY為負", "圓弧底剛形成", "晨星剛形成"],
    ["三大法人近3月買超", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["多方力道≥65", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "圓弧底剛形成", "晨星剛形成"],
    ["近3月均價YoY為負", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "大盤跌破20日均線", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["爆量(≥1.5倍均量)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月乖離度為負(股價超前營收)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["三大法人近3月賣超", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["強勢突破盤", "圓弧底剛形成", "晨星剛形成"],
    ["強勢突破盤", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "回後買上漲全通過"],
    ["相對強弱為正(強於大盤)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "三大法人近3月買超", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "近3月均價YoY為負", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "三大法人近3月買超", "晨星剛形成"],
    ["大盤站上20日均線", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["型態成形中", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["型態突破確認", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["型態剛形成(剛突破)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["強勢突破盤", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["大盤跌破60日均線", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "夜星剛形成"],
    ["多方力道≥65", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "大盤跌破20日均線", "夜星剛形成"],
    ["三大法人近3月買超", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "近3月均價YoY為正", "晨星剛形成"],
    ["多方力道≥80", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["爆量(≥2倍均量)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["多方力道≥65", "強勢突破盤", "地量(≤0.5倍均量)"],
    ["KDJ近3日內黃金交叉", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["多方力道≥80", "KDJ近3日內黃金交叉", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["量能區間高檔(≥90百分位)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["近3月均價YoY為正", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "母子懷抱(高檔)剛形成", "夜星剛形成"],
    ["近3月乖離度為負(股價超前營收)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["強勢突破盤", "N字底剛形成", "晨星剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "N字底剛形成"],
    ["多方力道≥80", "大盤跌破60日均線", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["量能區間高檔(≥90百分位)", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤站上20日均線", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["大盤站上60日均線", "圓弧底剛形成", "晨星剛形成"],
    ["回後買上漲全通過", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "母子懷抱(高檔)剛形成", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "圓弧底剛形成", "晨星剛形成"],
    ["爆量(≥2倍均量)", "圓弧底剛形成", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "圓弧底剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "大盤站上20日均線", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["多方力道≥80", "大盤站上20日均線", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "大盤跌破60日均線", "夜星剛形成"],
    ["量能區間高檔(≥90百分位)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["圓弧底剛形成", "突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成"],
    ["多方力道≥80", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["回後買上漲全通過", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["多方力道≥80", "相對強弱為正(強於大盤)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "KDJ近3日內黃金交叉", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "大盤跌破60日均線", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "回後買上漲全通過", "近3月乖離度為負(股價超前營收)"],
    ["近3月乖離度為負(股價超前營收)", "圓弧底剛形成", "晨星剛形成"],
    ["近3月乖離度為負(股價超前營收)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["多方力道≥65", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "大盤跌破20日均線", "夜星剛形成"],
    ["大盤跌破20日均線", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["近3月乖離度為正(營收優於股價)", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "地量(≤0.5倍均量)", "母子懷抱(高檔)剛形成"],
    ["強勢突破盤", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["量能區間高檔(≥90百分位)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["多方力道≥80", "晨星剛形成"],
    ["多方力道≥65", "多方力道≥80", "晨星剛形成"],
    ["多方力道≥80", "型態成形中", "晨星剛形成"],
    ["多方力道≥80", "型態突破確認", "晨星剛形成"],
    ["多方力道≥80", "型態剛形成(剛突破)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "量能區間高檔(≥90百分位)", "母子懷抱(低檔)剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤站上20日均線", "晨星剛形成"],
    ["大盤站上60日均線", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["三大法人近3月買超", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["近3月均價YoY為正", "N字底剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤站上60日均線", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["多方力道≥80", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["圓弧底剛形成", "晨星剛形成"],
    ["K線橫盤的突破剛形成", "晨星剛形成"],
    ["多方力道≥80", "量能區間低檔(≤10百分位)", "型態剛形成(剛突破)"],
    ["多方力道≥80", "大盤站上60日均線", "晨星剛形成"],
    ["強勢突破盤", "大盤跌破60日均線", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "母子懷抱(高檔)剛形成", "夜星剛形成"],
    ["KDJ近3日內死亡交叉", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["爆量(≥2倍均量)", "N字底剛形成", "晨星剛形成"],
    ["量能區間高檔(≥90百分位)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤站上60日均線", "晨星剛形成"],
    ["大盤跌破60日均線", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["型態成形中", "圓弧底剛形成", "晨星剛形成"],
    ["型態成形中", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["型態突破確認", "圓弧底剛形成", "晨星剛形成"],
    ["型態突破確認", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["型態剛形成(剛突破)", "圓弧底剛形成", "晨星剛形成"],
    ["型態剛形成(剛突破)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["近3月乖離度為正(營收優於股價)", "圓弧底剛形成", "晨星剛形成"],
    ["近3月均價YoY為負", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月乖離度為正(營收優於股價)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["多方力道≥80", "KDJ近3日內死亡交叉", "晨星剛形成"],
    ["多方力道≥80", "爆量(≥2倍均量)", "晨星剛形成"],
    ["大盤站上60日均線", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["大盤跌破60日均線", "三大法人近3月買超", "夜星剛形成"],
    ["強勢突破盤", "大盤跌破20日均線", "母子懷抱(低檔)剛形成"],
    ["相對強弱為正(強於大盤)", "地量(≤0.5倍均量)", "晨星剛形成"],
    ["大盤跌破20日均線", "三大法人近3月買超", "夜星剛形成"],
    ["多方力道≥65", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "強勢突破盤", "晨星剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "型態剛形成(剛突破)"],
    ["強勢突破盤", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破20日均線", "夜星剛形成"],
    ["近3月乖離度為正(營收優於股價)", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["強勢突破盤", "相對強弱為正(強於大盤)", "地量(≤0.5倍均量)"],
    ["相對強弱為正(強於大盤)", "大盤跌破20日均線", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "回後買上漲全通過", "N字底剛形成"],
    ["近3月均價YoY為正", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "回後買上漲全通過", "晨星剛形成"],
    ["多方力道≥65", "地量(≤0.5倍均量)", "回後買上漲全通過"],
    ["多方力道≥65", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["多方力道≥65", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "量能斜率轉強(近5日均量>近10日均量20%以上)", "母子懷抱(低檔)剛形成"],
    ["相對強弱為負(弱於大盤)", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["爆量(≥2倍均量)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["回後買上漲全通過", "突破ABC修正下降切線剛形成", "母子懷抱(低檔)剛形成"],
    ["量能區間高檔(≥90百分位)", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["多方力道≥65", "三大法人近3月買超", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "地量(≤0.5倍均量)", "回後買上漲全通過"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月乖離度為負(股價超前營收)", "夜星剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3月乖離度為負(股價超前營收)"],
    ["多方力道≥80", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月均價YoY為負", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["多方力道≥65", "大盤站上20日均線", "晨星剛形成"],
    ["多方力道≥80", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["強勢突破盤", "KDJ近3日內黃金交叉", "晨星剛形成"],
    ["強勢突破盤", "三大法人近3月賣超", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "三大法人近3月賣超", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["爆量(≥1.5倍均量)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "突破ABC修正下降切線剛形成", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "回後買上漲全通過"],
    ["地量(≤0.5倍均量)", "大盤站上60日均線", "回後買上漲全通過"],
    ["地量(≤0.5倍均量)", "近3月乖離度為負(股價超前營收)", "夜星剛形成"],
    ["突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["多方力道≥80", "布林通道高檔(≥80%)", "晨星剛形成"],
    ["強勢突破盤", "三大法人近3月賣超", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤跌破20日均線", "近3月均價YoY為正", "夜星剛形成"],
    ["型態成形中", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["型態突破確認", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["型態剛形成(剛突破)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "大盤站上60日均線", "夜星剛形成"],
    ["量能區間高檔(≥90百分位)", "大盤跌破20日均線", "夜星剛形成"],
    ["多方力道≥65", "三重底剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "母子懷抱(高檔)剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "大盤站上20日均線", "晨星剛形成"],
    ["布林通道低檔(≤20%)", "MACD近3日內黃金交叉", "夜星剛形成"],
    ["KDJ近3日內死亡交叉", "三重底剛形成", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "圓弧底剛形成", "突破ABC修正下降切線剛形成"],
    ["相對強弱為負(弱於大盤)", "爆量(≥2倍均量)", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["近3月均價YoY為負", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["圓弧底剛形成", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["多方力道≥65", "近3月均價YoY為正", "晨星剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "型態突破確認"],
    ["MACD近3日內黃金交叉", "相對強弱為正(強於大盤)", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "爆量(≥2倍均量)", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月均價YoY為正", "晨星剛形成"],
    ["大盤站上20日均線", "N字底剛形成", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "三大法人近3月賣超", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["爆量(≥1.5倍均量)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["爆量(≥1.5倍均量)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "近3月均價YoY為正", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "三大法人近3月買超", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "型態成形中", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "型態突破確認", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "型態剛形成(剛突破)", "晨星剛形成"],
    ["大盤站上20日均線", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月乖離度為正(營收優於股價)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "MACD近3日內黃金交叉", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "突破ABC修正下降切線剛形成"],
    ["布林通道高檔(≥80%)", "近3月均價YoY為正", "晨星剛形成"],
    ["量能區間高檔(≥90百分位)", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "相對強弱為正(強於大盤)", "晨星剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "三大法人近3月買超"],
    ["地量(≤0.5倍均量)", "回後買上漲全通過", "近3月均價YoY為正"],
    ["地量(≤0.5倍均量)", "回後買上漲全通過", "三大法人近3月買超"],
    ["三大法人近3月買超", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["大盤站上60日均線", "N字底剛形成", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月均價YoY為負", "夜星剛形成"],
    ["多方力道≥65", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "布林通道高檔(≥80%)", "地量(≤0.5倍均量)"],
    ["多方力道≥80", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "回後買上漲全通過"],
    ["多方力道≥80", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["強勢突破盤", "MACD近3日內黃金交叉", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "大盤跌破20日均線", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "近3月均價YoY為正", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "大盤站上20日均線", "回後買上漲全通過"],
    ["大盤站上20日均線", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["大盤站上20日均線", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤站上60日均線", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "大盤站上60日均線", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "爆量(≥1.5倍均量)", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "大盤站上20日均線", "晨星剛形成"],
    ["爆量(≥2倍均量)", "三大法人近3月買超", "母子懷抱(低檔)剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破20日均線", "母子懷抱(高檔)剛形成"],
    ["多方力道≥80", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "夜星剛形成"],
    ["多方力道≥80", "大盤跌破20日均線", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["相對強弱為負(弱於大盤)", "量能區間高檔(≥90百分位)", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "型態成形中", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "型態突破確認", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "型態剛形成(剛突破)", "夜星剛形成"],
    ["大盤站上20日均線", "近3月均價YoY為正", "晨星剛形成"],
    ["大盤站上20日均線", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["回後買上漲全通過", "近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "大盤站上60日均線", "晨星剛形成"],
    ["多方力道≥80", "大盤跌破20日均線", "晨星剛形成"],
    ["強勢突破盤", "相對強弱為正(強於大盤)", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "KDJ近3日內黃金交叉", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "大盤站上20日均線", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "突破飆股大量黑K最高點剛形成", "K線橫盤的突破剛形成"],
    ["爆量(≥1.5倍均量)", "圓弧底剛形成", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "型態成形中", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "型態突破確認", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "型態剛形成(剛突破)", "夜星剛形成"],
    ["型態成形中", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["型態突破確認", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["型態剛形成(剛突破)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["回後買上漲全通過", "近3月均價YoY為正", "母子懷抱(低檔)剛形成"],
    ["回後買上漲全通過", "三大法人近3月賣超", "母子懷抱(低檔)剛形成"],
    ["三大法人近3月買超", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "地量(≤0.5倍均量)", "夜星剛形成"],
    ["量能區間高檔(≥90百分位)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["多方力道≥65", "布林通道高檔(≥80%)", "晨星剛形成"],
    ["多方力道≥80", "強勢突破盤", "量能區間低檔(≤10百分位)"],
    ["多方力道≥80", "突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成"],
    ["強勢突破盤", "近3月均價YoY為正", "晨星剛形成"],
    ["強勢突破盤", "三重底剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "相對強弱為正(強於大盤)", "晨星剛形成"],
    ["大盤跌破20日均線", "大盤跌破60日均線", "夜星剛形成"],
    ["大盤站上60日均線", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["大盤站上60日均線", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月均價YoY為正", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["三大法人近3月買超", "頭肩底剛形成", "圓弧底剛形成"],
    ["三大法人近3月賣超", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "N字底剛形成", "晨星剛形成"],
    ["大盤站上20日均線", "N字底剛形成", "母子懷抱(低檔)剛形成"],
    ["回後買上漲全通過", "N字底剛形成", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "MACD近3日內黃金交叉", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "大盤站上20日均線"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3月均價YoY為正"],
    ["強勢突破盤", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "相對強弱為正(強於大盤)", "夜星剛形成"],
    ["KDJ近3日內死亡交叉", "圓弧底剛形成", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["爆量(≥2倍均量)", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "大盤站上60日均線", "晨星剛形成"],
    ["量能區間高檔(≥90百分位)", "近3月均價YoY為正", "夜星剛形成"],
    ["大盤跌破60日均線", "近3月均價YoY為正", "夜星剛形成"],
    ["回後買上漲全通過", "圓弧底剛形成", "晨星剛形成"],
    ["回後買上漲全通過", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["回後買上漲全通過", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["近3月均價YoY為正", "三大法人近3月買超", "夜星剛形成"],
    ["多方力道≥65", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "晨星剛形成"],
    ["大盤跌破20日均線", "夜星剛形成"],
    ["多方力道≥65", "型態成形中", "晨星剛形成"],
    ["多方力道≥65", "型態突破確認", "晨星剛形成"],
    ["多方力道≥65", "型態剛形成(剛突破)", "晨星剛形成"],
    ["多方力道≥80", "MACD近3日內黃金交叉", "N字底剛形成"],
    ["強勢突破盤", "KDJ近3日內死亡交叉", "地量(≤0.5倍均量)"],
    ["強勢突破盤", "大盤跌破60日均線", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "型態成形中", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "型態突破確認", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "型態剛形成(剛突破)", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成"],
    ["大盤跌破20日均線", "型態成形中", "夜星剛形成"],
    ["大盤跌破20日均線", "型態突破確認", "夜星剛形成"],
    ["大盤跌破20日均線", "型態剛形成(剛突破)", "夜星剛形成"],
    ["近3月均價YoY為正", "N字底剛形成", "晨星剛形成"],
    ["圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "大盤跌破20日均線", "夜星剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "量能區間高檔(≥90百分位)"],
    ["布林通道高檔(≥80%)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "三大法人近3月買超", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內黃金交叉", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "三大法人近3月買超", "母子懷抱(高檔)剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "圓弧底剛形成", "K線橫盤的突破剛形成"],
    ["大盤跌破20日均線", "大盤站上60日均線", "夜星剛形成"],
    ["大盤站上60日均線", "近3月均價YoY為正", "晨星剛形成"],
    ["大盤跌破60日均線", "突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成"],
    ["型態成形中", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["型態突破確認", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["型態剛形成(剛突破)", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月乖離度為正(營收優於股價)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["多方力道≥65", "爆量(≥2倍均量)", "晨星剛形成"],
    ["多方力道≥65", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "母子懷抱(高檔)剛形成"],
    ["多方力道≥80", "量能區間低檔(≤10百分位)", "近3月乖離度為負(股價超前營收)"],
    ["多方力道≥80", "近3月均價YoY為負", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "近3月乖離度為正(營收優於股價)", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "近3月均價YoY為正", "晨星剛形成"],
    ["大盤站上20日均線", "大盤站上60日均線", "晨星剛形成"],
    ["近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "夜星剛形成"],
    ["大盤站上20日均線", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["地量(≤0.5倍均量)", "回後買上漲全通過"],
    ["多方力道≥65", "強勢突破盤", "晨星剛形成"],
    ["多方力道≥65", "大盤跌破60日均線", "夜星剛形成"],
    ["多方力道≥65", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["多方力道≥80", "回後買上漲全通過", "N字底剛形成"],
    ["布林通道高檔(≥80%)", "近3月均價YoY為正", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內黃金交叉", "地量(≤0.5倍均量)", "回後買上漲全通過"],
    ["相對強弱為正(強於大盤)", "三大法人近3月買超", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "型態成形中", "回後買上漲全通過"],
    ["地量(≤0.5倍均量)", "型態突破確認", "回後買上漲全通過"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月均價YoY為負", "晨星剛形成"],
    ["回後買上漲全通過", "圓弧底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["強勢突破盤", "晨星剛形成"],
    ["多方力道≥65", "布林通道高檔(≥80%)", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "布林通道高檔(≥80%)", "晨星剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "大盤站上60日均線"],
    ["強勢突破盤", "型態成形中", "晨星剛形成"],
    ["強勢突破盤", "型態突破確認", "晨星剛形成"],
    ["強勢突破盤", "型態剛形成(剛突破)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "KDJ近3日內死亡交叉", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "爆量(≥1.5倍均量)", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["相對強弱為正(強於大盤)", "大盤站上60日均線", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月均價YoY為正", "夜星剛形成"],
    ["大盤站上20日均線", "三大法人近3月買超", "晨星剛形成"],
    ["大盤站上20日均線", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["近3月乖離度為負(股價超前營收)", "近3月均價YoY為負", "晨星剛形成"],
    ["三大法人近3月買超", "母子懷抱(高檔)剛形成", "夜星剛形成"],
    ["三重底剛形成", "圓弧底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["多方力道≥65", "N字底剛形成", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["地量(≤0.5倍均量)", "晨星剛形成"],
    ["大盤站上20日均線", "晨星剛形成"],
    ["多方力道≥65", "MACD近3日內黃金交叉", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "量能區間高檔(≥90百分位)"],
    ["強勢突破盤", "大盤站上60日均線", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "MACD近3日內黃金交叉", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "大盤跌破60日均線", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "型態成形中", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "型態突破確認", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "型態剛形成(剛突破)", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "近3月均價YoY為負", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "近3月均價YoY為負", "夜星剛形成"],
    ["量能區間高檔(≥90百分位)", "圓弧底剛形成", "晨星剛形成"],
    ["大盤站上20日均線", "型態成形中", "晨星剛形成"],
    ["大盤站上20日均線", "型態突破確認", "晨星剛形成"],
    ["大盤站上20日均線", "型態剛形成(剛突破)", "晨星剛形成"],
    ["大盤站上20日均線", "三大法人近3月賣超", "晨星剛形成"],
    ["大盤跌破60日均線", "近3月乖離度為負(股價超前營收)", "夜星剛形成"],
    ["近3月均價YoY為負", "突破飆股大量黑K最高點剛形成", "K線橫盤的突破剛形成"],
    ["大盤跌破60日均線", "夜星剛形成"],
    ["多方力道≥80", "近3月乖離度為正(營收優於股價)", "晨星剛形成"],
    ["多方力道≥80", "三大法人近3月買超", "夜星剛形成"],
    ["強勢突破盤", "近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "N字底剛形成"],
    ["布林通道高檔(≥80%)", "三大法人近3月買超", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "三大法人近3月買超", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "地量(≤0.5倍均量)", "母子懷抱(高檔)剛形成"],
    ["爆量(≥2倍均量)", "大盤跌破20日均線", "母子懷抱(低檔)剛形成"],
    ["大盤站上60日均線", "三大法人近3月買超", "晨星剛形成"],
    ["大盤跌破60日均線", "型態成形中", "夜星剛形成"],
    ["大盤跌破60日均線", "型態突破確認", "夜星剛形成"],
    ["大盤跌破60日均線", "型態剛形成(剛突破)", "夜星剛形成"],
    ["近3月均價YoY為負", "圓弧底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)"],
    ["多方力道≥80", "強勢突破盤", "量能斜率轉弱(近5日均量<近10日均量20%以上)"],
    ["強勢突破盤", "布林通道高檔(≥80%)", "地量(≤0.5倍均量)"],
    ["強勢突破盤", "KDJ近3日內黃金交叉", "地量(≤0.5倍均量)"],
    ["布林通道高檔(≥80%)", "大盤站上60日均線", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "量能斜率轉強(近5日均量>近10日均量20%以上)", "晨星剛形成"],
    ["量能區間高檔(≥90百分位)", "大盤跌破60日均線", "夜星剛形成"],
    ["大盤站上60日均線", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["大盤站上60日均線", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["近3月均價YoY為負", "三重底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["三大法人近3月買超", "N字底剛形成", "晨星剛形成"],
    ["多方力道≥65", "近3月均價YoY為正", "夜星剛形成"],
    ["多方力道≥65", "三大法人近3月買超", "夜星剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "N字底剛形成"],
    ["多方力道≥80", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "母子懷抱(高檔)剛形成"],
    ["強勢突破盤", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "相對強弱為負(弱於大盤)", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "大盤跌破20日均線", "夜星剛形成"],
    ["布林通道高檔(≥80%)", "突破ABC修正下降切線剛形成", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "三重底剛形成", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "近3月乖離度為正(營收優於股價)", "晨星剛形成"],
    ["大盤跌破20日均線", "近3月乖離度為正(營收優於股價)", "夜星剛形成"],
    ["近3月均價YoY為正", "圓弧底剛形成", "晨星剛形成"],
    ["大盤站上60日均線", "晨星剛形成"],
    ["多方力道≥65", "KDJ近3日內黃金交叉", "晨星剛形成"],
    ["多方力道≥65", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["多方力道≥80", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "回後買上漲全通過"],
    ["強勢突破盤", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["MACD近3日內黃金交叉", "近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "三大法人近3月買超", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "大盤跌破60日均線", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "三大法人近3月買超", "晨星剛形成"],
    ["爆量(≥2倍均量)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "大盤站上60日均線", "母子懷抱(低檔)剛形成"],
    ["量能區間高檔(≥90百分位)", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "大盤站上20日均線", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["大盤站上20日均線", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["大盤站上60日均線", "型態成形中", "晨星剛形成"],
    ["大盤站上60日均線", "型態突破確認", "晨星剛形成"],
    ["大盤站上60日均線", "型態剛形成(剛突破)", "晨星剛形成"],
    ["大盤跌破60日均線", "近3月均價YoY為負", "突破飆股大量黑K最高點剛形成"],
    ["大盤跌破60日均線", "突破飆股大量黑K最高點剛形成", "K線橫盤的突破剛形成"],
    ["回後買上漲全通過", "突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成"],
    ["近3月均價YoY為正", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["多方力道≥65", "突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成"],
    ["多方力道≥80", "量能斜率轉強(近5日均量>近10日均量20%以上)", "突破ABC修正下降切線剛形成"],
    ["多方力道≥80", "回後買上漲全通過", "複式頭肩底剛形成"],
    ["多方力道≥80", "近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "KDJ近3日內死亡交叉", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "相對強弱為負(弱於大盤)", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "近3月乖離度為負(股價超前營收)", "晨星剛形成"],
    ["爆量(≥2倍均量)", "量能區間高檔(≥90百分位)", "晨星剛形成"],
    ["爆量(≥2倍均量)", "近3月乖離度為正(營收優於股價)", "母子懷抱(低檔)剛形成"],
    ["大盤跌破20日均線", "近3月乖離度為負(股價超前營收)", "夜星剛形成"],
    ["回後買上漲全通過", "突破飆股大量黑K最高點剛形成", "母子懷抱(低檔)剛形成"],
    ["近3月乖離度為正(營收優於股價)", "近3月均價YoY為正", "夜星剛形成"],
    ["三大法人近3月買超", "圓弧底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["三大法人近3月賣超", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "母子懷抱(高檔)剛形成"],
    ["爆量(≥2倍均量)", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["近3月乖離度為正(營收優於股價)", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["多方力道≥65", "KDJ近3日內死亡交叉", "晨星剛形成"],
    ["多方力道≥80", "KDJ近3日內死亡交叉", "突破飆股大量黑K最高點剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)"],
    ["布林通道高檔(≥80%)", "大盤跌破20日均線", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["爆量(≥1.5倍均量)", "N字底剛形成", "晨星剛形成"],
    ["爆量(≥2倍均量)", "三大法人近3月買超", "夜星剛形成"],
    ["回後買上漲全通過", "近3月均價YoY為正", "晨星剛形成"],
    ["三大法人近3月買超", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["大盤站上60日均線", "N字底剛形成", "母子懷抱(低檔)剛形成"],
    ["N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["多方力道≥80", "回後買上漲全通過", "三大法人近3月買超"],
    ["強勢突破盤", "回後買上漲全通過", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "KDJ近3日內死亡交叉", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "大盤站上20日均線", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內黃金交叉", "大盤站上60日均線", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "量能區間高檔(≥90百分位)", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "大盤站上60日均線", "晨星剛形成"],
    ["量能區間高檔(≥90百分位)", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤站上60日均線", "夜星剛形成"],
    ["型態成形中", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["型態突破確認", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["型態剛形成(剛突破)", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["圓弧底剛形成", "突破飆股大量黑K最高點剛形成", "K線橫盤的突破剛形成"],
    ["突破ABC修正下降切線剛形成", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "夜星剛形成"],
    ["三大法人近3月買超", "夜星剛形成"],
    ["多方力道≥80", "KDJ近3日內黃金交叉", "N字底剛形成"],
    ["多方力道≥80", "量能區間高檔(≥90百分位)", "N字底剛形成"],
    ["強勢突破盤", "三大法人近3月買超", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "相對強弱為正(強於大盤)", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "三大法人近3月買超", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "型態成形中", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "型態突破確認", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "型態剛形成(剛突破)", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成"],
    ["爆量(≥1.5倍均量)", "突破ABC修正下降切線剛形成", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "大盤跌破20日均線", "夜星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "大盤跌破60日均線", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月乖離度為正(營收優於股價)", "晨星剛形成"],
    ["大盤站上60日均線", "三大法人近3月賣超", "晨星剛形成"],
    ["型態成形中", "三大法人近3月買超", "夜星剛形成"],
    ["型態突破確認", "三大法人近3月買超", "夜星剛形成"],
    ["型態剛形成(剛突破)", "三大法人近3月買超", "夜星剛形成"],
    ["近3月均價YoY為正", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "大盤跌破20日均線", "母子懷抱(高檔)剛形成"],
    ["多方力道≥65", "圓弧底剛形成", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "量能區間高檔(≥90百分位)", "晨星剛形成"],
    ["多方力道≥80", "大盤站上20日均線", "回後買上漲全通過"],
    ["多方力道≥80", "大盤跌破20日均線", "母子懷抱(高檔)剛形成"],
    ["多方力道≥80", "三大法人近3月買超", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "突破ABC修正下降切線剛形成"],
    ["MACD近3日內黃金交叉", "量能區間高檔(≥90百分位)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "近3月乖離度為負(股價超前營收)", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "爆量(≥2倍均量)", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內黃金交叉", "近3月均價YoY為正", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "母子懷抱(高檔)剛形成"],
    ["爆量(≥1.5倍均量)", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["爆量(≥2倍均量)", "大盤站上20日均線", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "大盤站上60日均線", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤站上20日均線", "夜星剛形成"],
    ["大盤跌破60日均線", "近3月乖離度為正(營收優於股價)", "突破飆股大量黑K最高點剛形成"],
    ["近3月乖離度為負(股價超前營收)", "三大法人近3月買超", "晨星剛形成"],
    ["近3月均價YoY為負", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "KDJ近3日內黃金交叉", "夜星剛形成"],
    ["多方力道≥65", "三大法人近3月買超", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "回後買上漲全通過", "近3月均價YoY為負"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "量能斜率轉強(近5日均量>近10日均量20%以上)"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "型態成形中"],
    ["布林通道高檔(≥80%)", "爆量(≥2倍均量)", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "量能區間高檔(≥90百分位)", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "型態成形中", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "型態突破確認", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "型態剛形成(剛突破)", "母子懷抱(低檔)剛形成"],
    ["布林通道高檔(≥80%)", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["MACD近3日內黃金交叉", "量能區間高檔(≥90百分位)", "母子懷抱(高檔)剛形成"],
    ["MACD近3日內黃金交叉", "近3月均價YoY為正", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "近3月均價YoY為正", "晨星剛形成"],
    ["KDJ近3日內黃金交叉", "KDJ近3日內死亡交叉", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "量能區間高檔(≥90百分位)", "夜星剛形成"],
    ["KDJ近3日內黃金交叉", "大盤站上20日均線", "晨星剛形成"],
    ["KDJ近3日內死亡交叉", "大盤跌破20日均線", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "近3月均價YoY為正", "晨星剛形成"],
    ["爆量(≥2倍均量)", "近3月均價YoY為負", "母子懷抱(低檔)剛形成"],
    ["量能區間高檔(≥90百分位)", "大盤站上20日均線", "晨星剛形成"],
    ["大盤站上20日均線", "大盤跌破60日均線", "突破飆股大量黑K最高點剛形成"],
    ["回後買上漲全通過", "近3月乖離度為正(營收優於股價)", "突破飆股大量黑K最高點剛形成"],
    ["近3月均價YoY為正", "三大法人近3月買超", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "晨星剛形成"],
    ["多方力道≥65", "近3月乖離度為負(股價超前營收)", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "量能斜率轉強(近5日均量>近10日均量20%以上)", "晨星剛形成"],
    ["多方力道≥80", "三大法人近3月賣超", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "爆量(≥1.5倍均量)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "KDJ近3日內死亡交叉", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "爆量(≥1.5倍均量)", "頭肩底剛形成"],
    ["MACD近3日內黃金交叉", "爆量(≥1.5倍均量)", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "大盤站上60日均線", "母子懷抱(低檔)剛形成"],
    ["相對強弱為正(強於大盤)", "型態成形中", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "型態突破確認", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "型態剛形成(剛突破)", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "近3月乖離度為負(股價超前營收)", "夜星剛形成"],
    ["爆量(≥1.5倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "圓弧底剛形成"],
    ["地量(≤0.5倍均量)", "大盤站上20日均線", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "大盤跌破20日均線", "母子懷抱(高檔)剛形成"],
    ["地量(≤0.5倍均量)", "大盤跌破60日均線", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3月乖離度為正(營收優於股價)", "夜星剛形成"],
    ["大盤站上20日均線", "三重底剛形成", "母子懷抱(低檔)剛形成"],
    ["大盤跌破60日均線", "頭肩底剛形成", "複式頭肩底剛形成"],
    ["大盤跌破60日均線", "複式頭肩底剛形成", "K線橫盤的突破剛形成"],
    ["大盤跌破60日均線", "母子懷抱(高檔)剛形成", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "母子懷抱(低檔)剛形成"],
    ["N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["多方力道≥65", "布林通道高檔(≥80%)", "地量(≤0.5倍均量)"],
    ["多方力道≥65", "爆量(≥2倍均量)", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "量能斜率轉強(近5日均量>近10日均量20%以上)", "晨星剛形成"],
    ["多方力道≥65", "大盤跌破60日均線", "晨星剛形成"],
    ["強勢突破盤", "近3月均價YoY為正", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "近3月均價YoY為負", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "型態成形中", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "型態突破確認", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "型態剛形成(剛突破)", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內黃金交叉", "量能區間高檔(≥90百分位)", "母子懷抱(低檔)剛形成"],
    ["相對強弱為正(強於大盤)", "三重底剛形成", "晨星剛形成"],
    ["相對強弱為負(弱於大盤)", "大盤站上60日均線", "晨星剛形成"],
    ["爆量(≥1.5倍均量)", "大盤站上20日均線", "晨星剛形成"],
    ["爆量(≥2倍均量)", "近3月均價YoY為正", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["量能區間高檔(≥90百分位)", "近3月均價YoY為正", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "大盤跌破20日均線", "夜星剛形成"],
    ["型態成形中", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["型態突破確認", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["型態剛形成(剛突破)", "N字底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["回後買上漲全通過", "近3月均價YoY為負", "突破飆股大量黑K最高點剛形成"],
    ["近3月均價YoY為正", "母子懷抱(高檔)剛形成", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "夜星剛形成"],
    ["爆量(≥2倍均量)", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "回後買上漲全通過"],
    ["多方力道≥80", "布林通道高檔(≥80%)", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "KDJ近3日內死亡交叉", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "相對強弱為正(強於大盤)", "回後買上漲全通過"],
    ["多方力道≥80", "爆量(≥1.5倍均量)", "N字底剛形成"],
    ["多方力道≥80", "大盤站上60日均線", "回後買上漲全通過"],
    ["多方力道≥80", "回後買上漲全通過", "近3月乖離度為正(營收優於股價)"],
    ["多方力道≥80", "近3月乖離度為正(營收優於股價)", "突破飆股大量黑K最高點剛形成"],
    ["強勢突破盤", "相對強弱為正(強於大盤)", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3月均價YoY為負"],
    ["MACD近3日內黃金交叉", "爆量(≥2倍均量)", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "三大法人近3月賣超", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "型態成形中", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "型態突破確認", "夜星剛形成"],
    ["相對強弱為正(強於大盤)", "型態剛形成(剛突破)", "夜星剛形成"],
    ["爆量(≥1.5倍均量)", "爆量(≥2倍均量)", "母子懷抱(低檔)剛形成"],
    ["爆量(≥1.5倍均量)", "大盤站上60日均線", "母子懷抱(低檔)剛形成"],
    ["爆量(≥1.5倍均量)", "三大法人近3月賣超", "晨星剛形成"],
    ["爆量(≥2倍均量)", "型態成形中", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "型態突破確認", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "型態剛形成(剛突破)", "母子懷抱(低檔)剛形成"],
    ["量能區間高檔(≥90百分位)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "晨星剛形成"],
    ["量能區間高檔(≥90百分位)", "大盤站上60日均線", "晨星剛形成"],
    ["大盤跌破20日均線", "突破ABC修正下降切線剛形成", "突破飆股大量黑K最高點剛形成"],
    ["大盤跌破60日均線", "回後買上漲全通過", "複式頭肩底剛形成"],
    ["近3月乖離度為負(股價超前營收)", "三大法人近3月買超", "夜星剛形成"],
    ["三大法人近3月賣超", "母子懷抱(低檔)剛形成", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "N字底剛形成", "突破ABC修正下降切線剛形成"],
    ["量能區間低檔(≤10百分位)", "大盤跌破60日均線", "三重底剛形成"],
    ["大盤站上20日均線", "複式頭肩底剛形成", "突破ABC修正下降切線剛形成"],
    ["大盤站上60日均線", "頭肩底剛形成", "突破飆股大量黑K最高點剛形成"],
    ["多方力道≥65", "KDJ近3日內死亡交叉", "母子懷抱(低檔)剛形成"],
    ["多方力道≥80", "MACD近3日內黃金交叉", "量能區間低檔(≤10百分位)"],
    ["多方力道≥80", "大盤站上20日均線", "N字底剛形成"],
    ["多方力道≥80", "三大法人近3月買超", "N字底剛形成"],
    ["布林通道高檔(≥80%)", "近3月均價YoY為負", "晨星剛形成"],
    ["布林通道高檔(≥80%)", "三大法人近3月賣超", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "KDJ近3日內黃金交叉", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "KDJ近3日內死亡交叉", "母子懷抱(高檔)剛形成"],
    ["MACD近3日內黃金交叉", "大盤跌破60日均線", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "近3月均價YoY為負", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "大盤站上20日均線", "晨星剛形成"],
    ["爆量(≥1.5倍均量)", "大盤站上20日均線", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "回後買上漲全通過"],
    ["地量(≤0.5倍均量)", "大盤跌破60日均線", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "三大法人近3月買超", "母子懷抱(高檔)剛形成"],
    ["量能區間高檔(≥90百分位)", "頭肩底剛形成", "圓弧底剛形成"],
    ["大盤站上20日均線", "回後買上漲全通過", "突破飆股大量黑K最高點剛形成"],
    ["大盤站上20日均線", "回後買上漲全通過", "晨星剛形成"],
    ["大盤跌破20日均線", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["大盤站上60日均線", "回後買上漲全通過", "晨星剛形成"],
    ["回後買上漲全通過", "N字底剛形成", "晨星剛形成"],
    ["近3月乖離度為正(營收優於股價)", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    # 2026-09-30 三年回測：20日飆股比例三段都 ≥2 倍基準（舊版選股型 S1 移到這裡）
    ["多方力道≥80", "強勢突破盤", "地量(≤0.5倍均量)"],
    ["多方力道≥80", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "創52週新高", "近3日向上跳空缺口"],
]

MOONSHOT_COMBO_STATS = {
    "突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 22, "moonshotN": 3, "pct": 13.6, "avg": 38.9}, "20": {"n": 22, "moonshotN": 7, "pct": 31.8, "avg": 47}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 20, "moonshotN": 2, "pct": 10, "avg": 55.6}, "20": {"n": 20, "moonshotN": 6, "pct": 30, "avg": 62.1}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過": {"10": {"n": 76, "moonshotN": 10, "pct": 13.2, "avg": 41.4}, "20": {"n": 76, "moonshotN": 19, "pct": 25, "avg": 63.6}},
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 24, "moonshotN": 3, "pct": 12.5, "avg": 42.8}, "20": {"n": 24, "moonshotN": 6, "pct": 25, "avg": 45.2}},
    "近3月乖離度為負(股價超前營收) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"10": {"n": 65, "moonshotN": 8, "pct": 12.3, "avg": 42.1}, "20": {"n": 65, "moonshotN": 15, "pct": 23.1, "avg": 54.4}},
    "多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"10": {"n": 104, "moonshotN": 12, "pct": 11.5, "avg": 43.1}, "20": {"n": 104, "moonshotN": 24, "pct": 23.1, "avg": 49.4}},
    "強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"10": {"n": 31, "moonshotN": 6, "pct": 19.4, "avg": 42.5}, "20": {"n": 31, "moonshotN": 7, "pct": 22.6, "avg": 46}},
    "MACD近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 40, "moonshotN": 4, "pct": 10, "avg": 43.1}, "20": {"n": 40, "moonshotN": 9, "pct": 22.5, "avg": 38.6}},
    "突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 27, "moonshotN": 6, "pct": 22.2, "avg": 40}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": {"20": {"n": 55, "moonshotN": 12, "pct": 21.8, "avg": 60.5}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 地量(≤0.5倍均量)": {"10": {"n": 292, "moonshotN": 30, "pct": 10.3, "avg": 47.7}, "20": {"n": 292, "moonshotN": 61, "pct": 20.9, "avg": 57.7}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 29, "moonshotN": 6, "pct": 20.7, "avg": 51.8}},
    "突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 58, "moonshotN": 12, "pct": 20.7, "avg": 46.7}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 34, "moonshotN": 7, "pct": 20.6, "avg": 43.4}},
    "近3月乖離度為負(股價超前營收) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 44, "moonshotN": 9, "pct": 20.5, "avg": 52.6}},
    "突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 54, "moonshotN": 11, "pct": 20.4, "avg": 49}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"10": {"n": 65, "moonshotN": 7, "pct": 10.8, "avg": 42.6}, "20": {"n": 65, "moonshotN": 13, "pct": 20, "avg": 56.8}},
    "大盤站上20日均線 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 40, "moonshotN": 8, "pct": 20, "avg": 43.6}},
    "地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成": {"5": {"n": 26, "moonshotN": 3, "pct": 11.5, "avg": 35}, "10": {"n": 26, "moonshotN": 5, "pct": 19.2, "avg": 44.6}, "20": {"n": 26, "moonshotN": 5, "pct": 19.2, "avg": 58.2}},
    "爆量(≥1.5倍均量) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 21, "moonshotN": 4, "pct": 19, "avg": 40.5}},
    "多方力道≥65 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"5": {"n": 32, "moonshotN": 4, "pct": 12.5, "avg": 33.9}, "10": {"n": 32, "moonshotN": 6, "pct": 18.8, "avg": 42.7}, "20": {"n": 32, "moonshotN": 5, "pct": 15.6, "avg": 43.9}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 32, "moonshotN": 4, "pct": 12.5, "avg": 46}, "20": {"n": 32, "moonshotN": 6, "pct": 18.8, "avg": 68}},
    "三大法人近3月賣超 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 48, "moonshotN": 9, "pct": 18.8, "avg": 62.8}},
    "多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": {"20": {"n": 65, "moonshotN": 12, "pct": 18.5, "avg": 60.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 27, "moonshotN": 5, "pct": 18.5, "avg": 44.3}},
    "突破ABC修正下降切線剛形成 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 27, "moonshotN": 5, "pct": 18.5, "avg": 44.5}},
    "近3月乖離度為負(股價超前營收) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 38, "moonshotN": 4, "pct": 10.5, "avg": 46.8}, "20": {"n": 38, "moonshotN": 7, "pct": 18.4, "avg": 44.9}},
    "近3月均價YoY為正 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"10": {"n": 76, "moonshotN": 8, "pct": 10.5, "avg": 42.5}, "20": {"n": 76, "moonshotN": 14, "pct": 18.4, "avg": 57.9}},
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 33, "moonshotN": 4, "pct": 12.1, "avg": 41.4}, "20": {"n": 33, "moonshotN": 6, "pct": 18.2, "avg": 45.2}},
    "多方力道≥65 ＋ 三重底剛形成 ＋ 晨星剛形成": {"20": {"n": 22, "moonshotN": 4, "pct": 18.2, "avg": 54}},
    "近3月均價YoY為負 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 33, "moonshotN": 6, "pct": 18.2, "avg": 51.8}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 111, "moonshotN": 20, "pct": 18, "avg": 52.8}},
    "大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 111, "moonshotN": 20, "pct": 18, "avg": 54.1}},
    "多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"10": {"n": 67, "moonshotN": 9, "pct": 13.4, "avg": 40}, "20": {"n": 67, "moonshotN": 12, "pct": 17.9, "avg": 65}},
    "強勢突破盤 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 39, "moonshotN": 7, "pct": 17.9, "avg": 43.7}},
    "大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 117, "moonshotN": 21, "pct": 17.9, "avg": 54.1}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過": {"10": {"n": 146, "moonshotN": 16, "pct": 11, "avg": 41}, "20": {"n": 146, "moonshotN": 26, "pct": 17.8, "avg": 63.1}},
    "爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 130, "moonshotN": 23, "pct": 17.7, "avg": 55.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"10": {"n": 108, "moonshotN": 12, "pct": 11.1, "avg": 40.4}, "20": {"n": 108, "moonshotN": 19, "pct": 17.6, "avg": 53.6}},
    "大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 74, "moonshotN": 13, "pct": 17.6, "avg": 49.1}},
    "地量(≤0.5倍均量) ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"10": {"n": 97, "moonshotN": 10, "pct": 10.3, "avg": 40.4}, "20": {"n": 97, "moonshotN": 17, "pct": 17.5, "avg": 56.6}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"10": {"n": 40, "moonshotN": 4, "pct": 10, "avg": 47.5}, "20": {"n": 40, "moonshotN": 7, "pct": 17.5, "avg": 61}},
    "近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 40, "moonshotN": 7, "pct": 17.5, "avg": 46.7}},
    "布林通道高檔(≥80%) ＋ 相對強弱為負(弱於大盤) ＋ 夜星剛形成": {"20": {"n": 23, "moonshotN": 4, "pct": 17.4, "avg": 41.3}},
    "三大法人近3月賣超 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 23, "moonshotN": 4, "pct": 17.4, "avg": 52.8}},
    "突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 145, "moonshotN": 25, "pct": 17.2, "avg": 53.7}},
    "型態成形中 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 145, "moonshotN": 25, "pct": 17.2, "avg": 53.7}},
    "型態突破確認 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 145, "moonshotN": 25, "pct": 17.2, "avg": 53.7}},
    "型態剛形成(剛突破) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 145, "moonshotN": 25, "pct": 17.2, "avg": 53.7}},
    "三大法人近3月賣超 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 70, "moonshotN": 12, "pct": 17.1, "avg": 53.4}},
    "相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"10": {"n": 94, "moonshotN": 12, "pct": 12.8, "avg": 38.8}, "20": {"n": 94, "moonshotN": 16, "pct": 17, "avg": 59.8}},
    "近3月均價YoY為正 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 112, "moonshotN": 12, "pct": 10.7, "avg": 41.4}, "20": {"n": 112, "moonshotN": 19, "pct": 17, "avg": 48.5}},
    "相對強弱為正(強於大盤) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 59, "moonshotN": 10, "pct": 16.9, "avg": 45}},
    "強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 24, "moonshotN": 4, "pct": 16.7, "avg": 39.3}, "20": {"n": 24, "moonshotN": 3, "pct": 12.5, "avg": 46.4}},
    "多方力道≥65 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 48, "moonshotN": 8, "pct": 16.7, "avg": 43.5}},
    "布林通道高檔(≥80%) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 138, "moonshotN": 23, "pct": 16.7, "avg": 54.6}},
    "KDJ近3日內黃金交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 42, "moonshotN": 7, "pct": 16.7, "avg": 49.4}},
    "近3月均價YoY為負 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 36, "moonshotN": 6, "pct": 16.7, "avg": 43.6}},
    "三大法人近3月買超 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 96, "moonshotN": 16, "pct": 16.7, "avg": 48.5}},
    "布林通道高檔(≥80%) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"10": {"n": 79, "moonshotN": 8, "pct": 10.1, "avg": 42.6}, "20": {"n": 79, "moonshotN": 13, "pct": 16.5, "avg": 46.2}},
    "布林通道高檔(≥80%) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 158, "moonshotN": 26, "pct": 16.5, "avg": 48.4}},
    "多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 171, "moonshotN": 28, "pct": 16.4, "avg": 49.6}},
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 61, "moonshotN": 10, "pct": 16.4, "avg": 46.3}},
    "相對強弱為正(強於大盤) ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 55, "moonshotN": 9, "pct": 16.4, "avg": 43.1}},
    "近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 67, "moonshotN": 11, "pct": 16.4, "avg": 48.3}},
    "布林通道高檔(≥80%) ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 98, "moonshotN": 16, "pct": 16.3, "avg": 52.3}},
    "布林通道高檔(≥80%) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 80, "moonshotN": 13, "pct": 16.3, "avg": 49.8}},
    "爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 80, "moonshotN": 13, "pct": 16.3, "avg": 49.5}},
    "近3月乖離度為負(股價超前營收) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 43, "moonshotN": 7, "pct": 16.3, "avg": 47.4}},
    "三大法人近3月賣超 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 37, "moonshotN": 5, "pct": 13.5, "avg": 38.2}, "20": {"n": 37, "moonshotN": 6, "pct": 16.2, "avg": 34.4}},
    "強勢突破盤 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 37, "moonshotN": 6, "pct": 16.2, "avg": 43}},
    "強勢突破盤 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 37, "moonshotN": 6, "pct": 16.2, "avg": 40.5}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過": {"20": {"n": 179, "moonshotN": 29, "pct": 16.2, "avg": 62.1}},
    "相對強弱為正(強於大盤) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 148, "moonshotN": 24, "pct": 16.2, "avg": 49.7}},
    "突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 87, "moonshotN": 14, "pct": 16.1, "avg": 48.4}},
    "多方力道≥80 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 292, "moonshotN": 47, "pct": 16.1, "avg": 52.3}},
    "MACD近3日內黃金交叉 ＋ 近3月均價YoY為負 ＋ 夜星剛形成": {"20": {"n": 62, "moonshotN": 10, "pct": 16.1, "avg": 47}},
    "相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 137, "moonshotN": 22, "pct": 16.1, "avg": 55.1}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 87, "moonshotN": 14, "pct": 16.1, "avg": 52.9}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 211, "moonshotN": 34, "pct": 16.1, "avg": 47.2}},
    "大盤站上20日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 118, "moonshotN": 19, "pct": 16.1, "avg": 46.8}},
    "型態成形中 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 87, "moonshotN": 14, "pct": 16.1, "avg": 48.4}},
    "型態突破確認 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 87, "moonshotN": 14, "pct": 16.1, "avg": 48.4}},
    "型態剛形成(剛突破) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 87, "moonshotN": 14, "pct": 16.1, "avg": 48.4}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 62, "moonshotN": 10, "pct": 16.1, "avg": 45}},
    "強勢突破盤 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"5": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 34}, "10": {"n": 25, "moonshotN": 4, "pct": 16, "avg": 43.9}, "20": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 48.5}},
    "大盤跌破60日均線 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 25, "moonshotN": 4, "pct": 16, "avg": 42.4}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"10": {"n": 44, "moonshotN": 7, "pct": 15.9, "avg": 42.2}, "20": {"n": 44, "moonshotN": 7, "pct": 15.9, "avg": 77.3}},
    "多方力道≥65 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 63, "moonshotN": 10, "pct": 15.9, "avg": 47.2}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 82, "moonshotN": 13, "pct": 15.9, "avg": 48.8}},
    "三大法人近3月買超 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 63, "moonshotN": 10, "pct": 15.9, "avg": 46.6}},
    "多方力道≥80 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"10": {"n": 240, "moonshotN": 25, "pct": 10.4, "avg": 43.1}, "20": {"n": 240, "moonshotN": 38, "pct": 15.8, "avg": 54.3}},
    "多方力道≥80 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 203, "moonshotN": 32, "pct": 15.8, "avg": 56.5}},
    "爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 101, "moonshotN": 16, "pct": 15.8, "avg": 55.6}},
    "多方力道≥65 ＋ 強勢突破盤 ＋ 地量(≤0.5倍均量)": {"20": {"n": 452, "moonshotN": 71, "pct": 15.7, "avg": 56.6}},
    "KDJ近3日內黃金交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 70, "moonshotN": 11, "pct": 15.7, "avg": 64.1}},
    "多方力道≥80 ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成": {"10": {"n": 154, "moonshotN": 18, "pct": 11.7, "avg": 45.7}, "20": {"n": 154, "moonshotN": 24, "pct": 15.6, "avg": 54.4}},
    "布林通道高檔(≥80%) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 77, "moonshotN": 12, "pct": 15.6, "avg": 44}},
    "布林通道高檔(≥80%) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 64, "moonshotN": 10, "pct": 15.6, "avg": 45}},
    "量能區間高檔(≥90百分位) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 128, "moonshotN": 20, "pct": 15.6, "avg": 55.3}},
    "近3月均價YoY為正 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 45, "moonshotN": 7, "pct": 15.6, "avg": 50.1}},
    "爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 58, "moonshotN": 9, "pct": 15.5, "avg": 51.7}},
    "地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成": {"20": {"n": 58, "moonshotN": 9, "pct": 15.5, "avg": 55.8}},
    "近3月乖離度為負(股價超前營收) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 103, "moonshotN": 16, "pct": 15.5, "avg": 50}},
    "強勢突破盤 ＋ N字底剛形成 ＋ 晨星剛形成": {"10": {"n": 26, "moonshotN": 4, "pct": 15.4, "avg": 41.9}, "20": {"n": 26, "moonshotN": 3, "pct": 11.5, "avg": 68.4}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ N字底剛形成": {"10": {"n": 26, "moonshotN": 3, "pct": 11.5, "avg": 39.4}, "20": {"n": 26, "moonshotN": 4, "pct": 15.4, "avg": 59.3}},
    "多方力道≥80 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"20": {"n": 52, "moonshotN": 8, "pct": 15.4, "avg": 46.3}},
    "相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 78, "moonshotN": 12, "pct": 15.4, "avg": 51.1}},
    "爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 52, "moonshotN": 8, "pct": 15.4, "avg": 46.2}},
    "量能區間高檔(≥90百分位) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 26, "moonshotN": 4, "pct": 15.4, "avg": 40.5}},
    "大盤站上20日均線 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 149, "moonshotN": 23, "pct": 15.4, "avg": 49}},
    "大盤站上60日均線 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 52, "moonshotN": 8, "pct": 15.4, "avg": 43.6}},
    "回後買上漲全通過 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 39, "moonshotN": 6, "pct": 15.4, "avg": 47}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成": {"20": {"n": 59, "moonshotN": 9, "pct": 15.3, "avg": 61.4}},
    "KDJ近3日內黃金交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 33, "moonshotN": 5, "pct": 15.2, "avg": 45.3}},
    "爆量(≥2倍均量) ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 33, "moonshotN": 5, "pct": 15.2, "avg": 44.6}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 66, "moonshotN": 10, "pct": 15.2, "avg": 49.5}},
    "布林通道高檔(≥80%) ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 53, "moonshotN": 8, "pct": 15.1, "avg": 44.5}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 93, "moonshotN": 14, "pct": 15.1, "avg": 61.7}},
    "相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"10": {"n": 100, "moonshotN": 12, "pct": 12, "avg": 40.3}, "20": {"n": 100, "moonshotN": 15, "pct": 15, "avg": 59.1}},
    "多方力道≥80 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 334, "moonshotN": 50, "pct": 15, "avg": 54.8}},
    "布林通道高檔(≥80%) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 40, "moonshotN": 6, "pct": 15, "avg": 47.8}},
    "量能區間高檔(≥90百分位) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 60, "moonshotN": 9, "pct": 15, "avg": 44.5}},
    "圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 60, "moonshotN": 9, "pct": 15, "avg": 47.7}},
    "多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"10": {"n": 47, "moonshotN": 7, "pct": 14.9, "avg": 41.6}, "20": {"n": 47, "moonshotN": 6, "pct": 12.8, "avg": 72.4}},
    "回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"10": {"n": 47, "moonshotN": 5, "pct": 10.6, "avg": 35.1}, "20": {"n": 47, "moonshotN": 7, "pct": 14.9, "avg": 47.7}},
    "多方力道≥80 ＋ 相對強弱為正(強於大盤) ＋ 晨星剛形成": {"20": {"n": 397, "moonshotN": 59, "pct": 14.9, "avg": 52.9}},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內黃金交叉 ＋ 夜星剛形成": {"20": {"n": 74, "moonshotN": 11, "pct": 14.9, "avg": 50.8}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 54, "moonshotN": 8, "pct": 14.8, "avg": 48.3}},
    "地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ 近3月乖離度為負(股價超前營收)": {"20": {"n": 155, "moonshotN": 23, "pct": 14.8, "avg": 61.3}},
    "近3月乖離度為負(股價超前營收) ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 27, "moonshotN": 4, "pct": 14.8, "avg": 40.6}},
    "近3月乖離度為負(股價超前營收) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 81, "moonshotN": 12, "pct": 14.8, "avg": 49.7}},
    "多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"10": {"n": 75, "moonshotN": 8, "pct": 10.7, "avg": 40.6}, "20": {"n": 75, "moonshotN": 11, "pct": 14.7, "avg": 56.6}},
    "地量(≤0.5倍均量) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 75, "moonshotN": 11, "pct": 14.7, "avg": 61.5}},
    "大盤跌破20日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 34, "moonshotN": 5, "pct": 14.7, "avg": 51.8}},
    "近3月乖離度為正(營收優於股價) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 34, "moonshotN": 5, "pct": 14.7, "avg": 55.7}},
    "多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 157, "moonshotN": 23, "pct": 14.6, "avg": 54.4}},
    "強勢突破盤 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 103, "moonshotN": 15, "pct": 14.6, "avg": 51.8}},
    "相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 123, "moonshotN": 18, "pct": 14.6, "avg": 46.8}},
    "量能區間高檔(≥90百分位) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 137, "moonshotN": 20, "pct": 14.6, "avg": 47.9}},
    "多方力道≥80 ＋ 晨星剛形成": {"20": {"n": 408, "moonshotN": 59, "pct": 14.5, "avg": 52.9}},
    "多方力道≥65 ＋ 多方力道≥80 ＋ 晨星剛形成": {"20": {"n": 408, "moonshotN": 59, "pct": 14.5, "avg": 52.9}},
    "多方力道≥80 ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 408, "moonshotN": 59, "pct": 14.5, "avg": 52.9}},
    "多方力道≥80 ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 408, "moonshotN": 59, "pct": 14.5, "avg": 52.9}},
    "多方力道≥80 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 408, "moonshotN": 59, "pct": 14.5, "avg": 52.9}},
    "MACD近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 62, "moonshotN": 9, "pct": 14.5, "avg": 48.6}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 228, "moonshotN": 33, "pct": 14.5, "avg": 51}},
    "大盤站上60日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 125, "moonshotN": 18, "pct": 14.4, "avg": 47.2}},
    "三大法人近3月買超 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 104, "moonshotN": 15, "pct": 14.4, "avg": 49.2}},
    "近3月均價YoY為正 ＋ N字底剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 21, "moonshotN": 3, "pct": 14.3, "avg": 36.2}, "20": {"n": 21, "moonshotN": 3, "pct": 14.3, "avg": 51}},
    "大盤站上60日均線 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 63, "moonshotN": 8, "pct": 12.7, "avg": 41.9}, "20": {"n": 63, "moonshotN": 9, "pct": 14.3, "avg": 42.1}},
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"10": {"n": 35, "moonshotN": 4, "pct": 11.4, "avg": 36.2}, "20": {"n": 35, "moonshotN": 5, "pct": 14.3, "avg": 58.4}},
    "圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 63, "moonshotN": 9, "pct": 14.3, "avg": 43.1}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 189, "moonshotN": 27, "pct": 14.3, "avg": 48.1}},
    "多方力道≥80 ＋ 量能區間低檔(≤10百分位) ＋ 型態剛形成(剛突破)": {"20": {"n": 42, "moonshotN": 6, "pct": 14.3, "avg": 53.9}},
    "多方力道≥80 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 356, "moonshotN": 51, "pct": 14.3, "avg": 53.9}},
    "強勢突破盤 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 28, "moonshotN": 4, "pct": 14.3, "avg": 38.1}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": {"20": {"n": 28, "moonshotN": 4, "pct": 14.3, "avg": 42.4}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成": {"20": {"n": 56, "moonshotN": 8, "pct": 14.3, "avg": 44.7}},
    "KDJ近3日內死亡交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 63, "moonshotN": 9, "pct": 14.3, "avg": 56.3}},
    "相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 244, "moonshotN": 35, "pct": 14.3, "avg": 49.9}},
    "爆量(≥2倍均量) ＋ N字底剛形成 ＋ 晨星剛形成": {"20": {"n": 21, "moonshotN": 3, "pct": 14.3, "avg": 68.4}},
    "量能區間高檔(≥90百分位) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 84, "moonshotN": 12, "pct": 14.3, "avg": 42.7}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 21, "moonshotN": 3, "pct": 14.3, "avg": 40.6}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 273, "moonshotN": 39, "pct": 14.3, "avg": 50.6}},
    "大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 28, "moonshotN": 4, "pct": 14.3, "avg": 51.3}},
    "型態成形中 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 63, "moonshotN": 9, "pct": 14.3, "avg": 43.1}},
    "型態成形中 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 189, "moonshotN": 27, "pct": 14.3, "avg": 48.1}},
    "型態突破確認 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 63, "moonshotN": 9, "pct": 14.3, "avg": 43.1}},
    "型態突破確認 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 189, "moonshotN": 27, "pct": 14.3, "avg": 48.1}},
    "型態剛形成(剛突破) ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 63, "moonshotN": 9, "pct": 14.3, "avg": 43.1}},
    "型態剛形成(剛突破) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 189, "moonshotN": 27, "pct": 14.3, "avg": 48.1}},
    "近3月乖離度為正(營收優於股價) ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 35, "moonshotN": 5, "pct": 14.3, "avg": 45.1}},
    "近3月均價YoY為負 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 35, "moonshotN": 5, "pct": 14.3, "avg": 48.2}},
    "突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 21, "moonshotN": 3, "pct": 14.3, "avg": 46.5}},
    "近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 71, "moonshotN": 10, "pct": 14.1, "avg": 52.5}},
    "多方力道≥80 ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成": {"20": {"n": 157, "moonshotN": 22, "pct": 14, "avg": 48.6}},
    "多方力道≥80 ＋ 爆量(≥2倍均量) ＋ 晨星剛形成": {"20": {"n": 93, "moonshotN": 13, "pct": 14, "avg": 62.2}},
    "大盤站上60日均線 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 164, "moonshotN": 23, "pct": 14, "avg": 49.1}},
    "大盤跌破60日均線 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 179, "moonshotN": 25, "pct": 14, "avg": 51.5}},
    "強勢突破盤 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 36, "moonshotN": 5, "pct": 13.9, "avg": 42.8}},
    "相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": {"20": {"n": 108, "moonshotN": 15, "pct": 13.9, "avg": 56.4}},
    "大盤跌破20日均線 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 294, "moonshotN": 41, "pct": 13.9, "avg": 51}},
    "多方力道≥65 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 58, "moonshotN": 8, "pct": 13.8, "avg": 57.3}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 晨星剛形成": {"20": {"n": 174, "moonshotN": 24, "pct": 13.8, "avg": 53.4}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 型態剛形成(剛突破)": {"20": {"n": 116, "moonshotN": 16, "pct": 13.8, "avg": 61.4}},
    "強勢突破盤 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 195, "moonshotN": 27, "pct": 13.8, "avg": 49.6}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 94, "moonshotN": 13, "pct": 13.8, "avg": 56.8}},
    "近3月乖離度為正(營收優於股價) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 87, "moonshotN": 12, "pct": 13.8, "avg": 56.9}},
    "強勢突破盤 ＋ 相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量)": {"20": {"n": 527, "moonshotN": 72, "pct": 13.7, "avg": 56.6}},
    "相對強弱為正(強於大盤) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 358, "moonshotN": 49, "pct": 13.7, "avg": 53.1}},
    "地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ N字底剛形成": {"20": {"n": 51, "moonshotN": 7, "pct": 13.7, "avg": 62.4}},
    "近3月均價YoY為正 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 51, "moonshotN": 7, "pct": 13.7, "avg": 44.9}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過 ＋ 晨星剛形成": {"10": {"n": 22, "moonshotN": 3, "pct": 13.6, "avg": 35.6}, "20": {"n": 22, "moonshotN": 3, "pct": 13.6, "avg": 33.5}},
    "多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過": {"20": {"n": 191, "moonshotN": 26, "pct": 13.6, "avg": 63.1}},
    "多方力道≥65 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 346, "moonshotN": 47, "pct": 13.6, "avg": 54.9}},
    "多方力道≥65 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 118, "moonshotN": 16, "pct": 13.6, "avg": 51.5}},
    "MACD近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 59, "moonshotN": 8, "pct": 13.6, "avg": 44.9}},
    "相對強弱為負(弱於大盤) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 22, "moonshotN": 3, "pct": 13.6, "avg": 39.3}},
    "爆量(≥2倍均量) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 44, "moonshotN": 6, "pct": 13.6, "avg": 50.5}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 59, "moonshotN": 8, "pct": 13.6, "avg": 44}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 118, "moonshotN": 16, "pct": 13.6, "avg": 49.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 184, "moonshotN": 25, "pct": 13.6, "avg": 49}},
    "回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 22, "moonshotN": 3, "pct": 13.6, "avg": 42.5}},
    "量能區間高檔(≥90百分位) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 37, "moonshotN": 5, "pct": 13.5, "avg": 37.4}},
    "多方力道≥65 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 465, "moonshotN": 63, "pct": 13.5, "avg": 51.4}},
    "相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過": {"20": {"n": 222, "moonshotN": 30, "pct": 13.5, "avg": 60.9}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成": {"20": {"n": 104, "moonshotN": 14, "pct": 13.5, "avg": 51}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3月乖離度為負(股價超前營收)": {"20": {"n": 357, "moonshotN": 48, "pct": 13.4, "avg": 58}},
    "多方力道≥80 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 30, "moonshotN": 4, "pct": 13.3, "avg": 41.5}, "20": {"n": 30, "moonshotN": 4, "pct": 13.3, "avg": 56.4}},
    "近3月均價YoY為負 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 30, "moonshotN": 4, "pct": 13.3, "avg": 37}},
    "多方力道≥65 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 540, "moonshotN": 72, "pct": 13.3, "avg": 52.4}},
    "多方力道≥80 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 83, "moonshotN": 11, "pct": 13.3, "avg": 60.3}},
    "強勢突破盤 ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成": {"20": {"n": 180, "moonshotN": 24, "pct": 13.3, "avg": 50.7}},
    "強勢突破盤 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"20": {"n": 128, "moonshotN": 17, "pct": 13.3, "avg": 49.5}},
    "布林通道高檔(≥80%) ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"20": {"n": 248, "moonshotN": 33, "pct": 13.3, "avg": 48.4}},
    "相對強弱為正(強於大盤) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 60, "moonshotN": 8, "pct": 13.3, "avg": 43.4}},
    "爆量(≥1.5倍均量) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 75, "moonshotN": 10, "pct": 13.3, "avg": 45}},
    "爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 30, "moonshotN": 4, "pct": 13.3, "avg": 44}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過": {"20": {"n": 188, "moonshotN": 25, "pct": 13.3, "avg": 59.9}},
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 回後買上漲全通過": {"20": {"n": 249, "moonshotN": 33, "pct": 13.3, "avg": 59.6}},
    "地量(≤0.5倍均量) ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成": {"20": {"n": 98, "moonshotN": 13, "pct": 13.3, "avg": 52.7}},
    "突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 151, "moonshotN": 20, "pct": 13.2, "avg": 46.1}},
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"20": {"n": 295, "moonshotN": 39, "pct": 13.2, "avg": 51}},
    "強勢突破盤 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 76, "moonshotN": 10, "pct": 13.2, "avg": 46.5}},
    "強勢突破盤 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 114, "moonshotN": 15, "pct": 13.2, "avg": 53.6}},
    "布林通道高檔(≥80%) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 378, "moonshotN": 50, "pct": 13.2, "avg": 49.2}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 53, "moonshotN": 7, "pct": 13.2, "avg": 48.1}},
    "大盤跌破20日均線 ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 394, "moonshotN": 52, "pct": 13.2, "avg": 52.1}},
    "型態成形中 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 151, "moonshotN": 20, "pct": 13.2, "avg": 46.1}},
    "型態突破確認 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 151, "moonshotN": 20, "pct": 13.2, "avg": 46.1}},
    "型態剛形成(剛突破) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 151, "moonshotN": 20, "pct": 13.2, "avg": 46.1}},
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 夜星剛形成": {"20": {"n": 122, "moonshotN": 16, "pct": 13.1, "avg": 54.6}},
    "量能區間高檔(≥90百分位) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 183, "moonshotN": 24, "pct": 13.1, "avg": 50.1}},
    "多方力道≥65 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 23, "moonshotN": 3, "pct": 13, "avg": 54.8}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 77, "moonshotN": 10, "pct": 13, "avg": 54}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 92, "moonshotN": 12, "pct": 13, "avg": 74.7}},
    "強勢突破盤 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 315, "moonshotN": 41, "pct": 13, "avg": 48.7}},
    "布林通道低檔(≤20%) ＋ MACD近3日內黃金交叉 ＋ 夜星剛形成": {"20": {"n": 23, "moonshotN": 3, "pct": 13, "avg": 50.3}},
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 23, "moonshotN": 3, "pct": 13, "avg": 54.8}},
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 46, "moonshotN": 6, "pct": 13, "avg": 45.2}},
    "相對強弱為負(弱於大盤) ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 77, "moonshotN": 10, "pct": 13, "avg": 44.7}},
    "爆量(≥2倍均量) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 100, "moonshotN": 13, "pct": 13, "avg": 50.2}},
    "近3月均價YoY為負 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 46, "moonshotN": 6, "pct": 13, "avg": 42.9}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 23, "moonshotN": 3, "pct": 13, "avg": 38.5}},
    "多方力道≥65 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 426, "moonshotN": 55, "pct": 12.9, "avg": 53.3}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 型態突破確認": {"20": {"n": 357, "moonshotN": 46, "pct": 12.9, "avg": 57.7}},
    "MACD近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 70, "moonshotN": 9, "pct": 12.9, "avg": 44.2}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 夜星剛形成": {"20": {"n": 31, "moonshotN": 4, "pct": 12.9, "avg": 56.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 232, "moonshotN": 30, "pct": 12.9, "avg": 52.2}},
    "大盤站上20日均線 ＋ N字底剛形成 ＋ 晨星剛形成": {"10": {"n": 39, "moonshotN": 5, "pct": 12.8, "avg": 39.6}, "20": {"n": 39, "moonshotN": 4, "pct": 10.3, "avg": 57.4}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 352, "moonshotN": 45, "pct": 12.8, "avg": 50.5}},
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"20": {"n": 78, "moonshotN": 10, "pct": 12.8, "avg": 49.3}},
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 86, "moonshotN": 11, "pct": 12.8, "avg": 51.9}},
    "爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 78, "moonshotN": 10, "pct": 12.8, "avg": 46.2}},
    "爆量(≥1.5倍均量) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 156, "moonshotN": 20, "pct": 12.8, "avg": 45.8}},
    "地量(≤0.5倍均量) ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 117, "moonshotN": 15, "pct": 12.8, "avg": 57.4}},
    "地量(≤0.5倍均量) ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 109, "moonshotN": 14, "pct": 12.8, "avg": 54.9}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 352, "moonshotN": 45, "pct": 12.8, "avg": 50.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 352, "moonshotN": 45, "pct": 12.8, "avg": 50.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 352, "moonshotN": 45, "pct": 12.8, "avg": 50.5}},
    "大盤站上20日均線 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 39, "moonshotN": 5, "pct": 12.8, "avg": 54.3}},
    "近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 39, "moonshotN": 5, "pct": 12.8, "avg": 40.8}},
    "多方力道≥65 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 102, "moonshotN": 13, "pct": 12.7, "avg": 59.3}},
    "布林通道高檔(≥80%) ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": {"20": {"n": 126, "moonshotN": 16, "pct": 12.7, "avg": 49.1}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 55, "moonshotN": 7, "pct": 12.7, "avg": 67.6}},
    "布林通道高檔(≥80%) ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 472, "moonshotN": 60, "pct": 12.7, "avg": 49.9}},
    "量能區間高檔(≥90百分位) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 71, "moonshotN": 9, "pct": 12.7, "avg": 47.2}},
    "多方力道≥65 ＋ 相對強弱為正(強於大盤) ＋ 晨星剛形成": {"20": {"n": 635, "moonshotN": 80, "pct": 12.6, "avg": 50.8}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 三大法人近3月買超": {"20": {"n": 475, "moonshotN": 60, "pct": 12.6, "avg": 57.6}},
    "地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ 近3月均價YoY為正": {"20": {"n": 215, "moonshotN": 27, "pct": 12.6, "avg": 61}},
    "地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ 三大法人近3月買超": {"20": {"n": 223, "moonshotN": 28, "pct": 12.6, "avg": 57.6}},
    "三大法人近3月買超 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 119, "moonshotN": 15, "pct": 12.6, "avg": 43.8}},
    "相對強弱為負(弱於大盤) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 24, "moonshotN": 3, "pct": 12.5, "avg": 36.4}},
    "大盤站上60日均線 ＋ N字底剛形成 ＋ 晨星剛形成": {"10": {"n": 40, "moonshotN": 5, "pct": 12.5, "avg": 39.6}, "20": {"n": 40, "moonshotN": 4, "pct": 10, "avg": 57.4}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月均價YoY為負 ＋ 夜星剛形成": {"10": {"n": 48, "moonshotN": 5, "pct": 10.4, "avg": 44.9}, "20": {"n": 48, "moonshotN": 6, "pct": 12.5, "avg": 46.4}},
    "多方力道≥65 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 48, "moonshotN": 6, "pct": 12.5, "avg": 50.1}},
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量)": {"20": {"n": 1734, "moonshotN": 216, "pct": 12.5, "avg": 54.5}},
    "多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過": {"20": {"n": 439, "moonshotN": 55, "pct": 12.5, "avg": 56.4}},
    "多方力道≥80 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 72, "moonshotN": 9, "pct": 12.5, "avg": 57.9}},
    "強勢突破盤 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": {"20": {"n": 40, "moonshotN": 5, "pct": 12.5, "avg": 42}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 208, "moonshotN": 26, "pct": 12.5, "avg": 44.8}},
    "相對強弱為正(強於大盤) ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 513, "moonshotN": 64, "pct": 12.5, "avg": 55.8}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 回後買上漲全通過": {"20": {"n": 255, "moonshotN": 32, "pct": 12.5, "avg": 53.2}},
    "大盤站上20日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 598, "moonshotN": 75, "pct": 12.5, "avg": 51.8}},
    "大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 72, "moonshotN": 9, "pct": 12.5, "avg": 46.3}},
    "大盤站上60日均線 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 40, "moonshotN": 5, "pct": 12.5, "avg": 54.3}},
    "布林通道高檔(≥80%) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 643, "moonshotN": 80, "pct": 12.4, "avg": 48.7}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 晨星剛形成": {"20": {"n": 121, "moonshotN": 15, "pct": 12.4, "avg": 45.7}},
    "相對強弱為正(強於大盤) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 734, "moonshotN": 91, "pct": 12.4, "avg": 52.2}},
    "爆量(≥2倍均量) ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 129, "moonshotN": 16, "pct": 12.4, "avg": 50.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 137, "moonshotN": 17, "pct": 12.4, "avg": 56.9}},
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"10": {"n": 122, "moonshotN": 13, "pct": 10.7, "avg": 43.2}, "20": {"n": 122, "moonshotN": 15, "pct": 12.3, "avg": 58.7}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成": {"20": {"n": 162, "moonshotN": 20, "pct": 12.3, "avg": 55.6}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 65, "moonshotN": 8, "pct": 12.3, "avg": 64.1}},
    "多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 65, "moonshotN": 8, "pct": 12.3, "avg": 68.2}},
    "KDJ近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"20": {"n": 57, "moonshotN": 7, "pct": 12.3, "avg": 41.3}},
    "相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 114, "moonshotN": 14, "pct": 12.3, "avg": 56.6}},
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 夜星剛形成": {"20": {"n": 162, "moonshotN": 20, "pct": 12.3, "avg": 55.6}},
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 夜星剛形成": {"20": {"n": 162, "moonshotN": 20, "pct": 12.3, "avg": 55.6}},
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": {"20": {"n": 162, "moonshotN": 20, "pct": 12.3, "avg": 55.6}},
    "大盤站上20日均線 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 751, "moonshotN": 92, "pct": 12.3, "avg": 52.6}},
    "大盤站上20日均線 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 65, "moonshotN": 8, "pct": 12.3, "avg": 47.1}},
    "回後買上漲全通過 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 81, "moonshotN": 10, "pct": 12.3, "avg": 53.4}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"20": {"n": 180, "moonshotN": 22, "pct": 12.2, "avg": 53.1}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 82, "moonshotN": 10, "pct": 12.2, "avg": 45}},
    "多方力道≥65 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 599, "moonshotN": 73, "pct": 12.2, "avg": 52.2}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"20": {"n": 74, "moonshotN": 9, "pct": 12.2, "avg": 42.2}},
    "強勢突破盤 ＋ 相對強弱為正(強於大盤) ＋ 晨星剛形成": {"20": {"n": 360, "moonshotN": 44, "pct": 12.2, "avg": 48.8}},
    "布林通道高檔(≥80%) ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成": {"20": {"n": 353, "moonshotN": 43, "pct": 12.2, "avg": 51.5}},
    "布林通道高檔(≥80%) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 74, "moonshotN": 9, "pct": 12.2, "avg": 54.8}},
    "布林通道高檔(≥80%) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 607, "moonshotN": 74, "pct": 12.2, "avg": 48.7}},
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 90, "moonshotN": 11, "pct": 12.2, "avg": 51.3}},
    "KDJ近3日內死亡交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成": {"20": {"n": 205, "moonshotN": 25, "pct": 12.2, "avg": 51.6}},
    "爆量(≥1.5倍均量) ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 41, "moonshotN": 5, "pct": 12.2, "avg": 44.6}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態成形中 ＋ 夜星剛形成": {"20": {"n": 180, "moonshotN": 22, "pct": 12.2, "avg": 53.1}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態突破確認 ＋ 夜星剛形成": {"20": {"n": 180, "moonshotN": 22, "pct": 12.2, "avg": 53.1}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": {"20": {"n": 180, "moonshotN": 22, "pct": 12.2, "avg": 53.1}},
    "型態成形中 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 82, "moonshotN": 10, "pct": 12.2, "avg": 45}},
    "型態突破確認 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 82, "moonshotN": 10, "pct": 12.2, "avg": 45}},
    "型態剛形成(剛突破) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 82, "moonshotN": 10, "pct": 12.2, "avg": 45}},
    "回後買上漲全通過 ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 90, "moonshotN": 11, "pct": 12.2, "avg": 48.4}},
    "回後買上漲全通過 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 49, "moonshotN": 6, "pct": 12.2, "avg": 50.7}},
    "三大法人近3月買超 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 49, "moonshotN": 6, "pct": 12.2, "avg": 44}},
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"10": {"n": 66, "moonshotN": 8, "pct": 12.1, "avg": 35.4}},
    "量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"10": {"n": 33, "moonshotN": 4, "pct": 12.1, "avg": 35.9}},
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"20": {"n": 487, "moonshotN": 59, "pct": 12.1, "avg": 49.2}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能區間低檔(≤10百分位)": {"20": {"n": 33, "moonshotN": 4, "pct": 12.1, "avg": 53.3}},
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 190, "moonshotN": 23, "pct": 12.1, "avg": 54.6}},
    "強勢突破盤 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 224, "moonshotN": 27, "pct": 12.1, "avg": 50.4}},
    "強勢突破盤 ＋ 三重底剛形成 ＋ 晨星剛形成": {"20": {"n": 33, "moonshotN": 4, "pct": 12.1, "avg": 54.4}},
    "布林通道高檔(≥80%) ＋ 相對強弱為正(強於大盤) ＋ 晨星剛形成": {"20": {"n": 638, "moonshotN": 77, "pct": 12.1, "avg": 48.7}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 256, "moonshotN": 31, "pct": 12.1, "avg": 49.7}},
    "大盤站上60日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 727, "moonshotN": 88, "pct": 12.1, "avg": 50.7}},
    "大盤站上60日均線 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 66, "moonshotN": 8, "pct": 12.1, "avg": 47.4}},
    "近3月均價YoY為正 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 99, "moonshotN": 12, "pct": 12.1, "avg": 50}},
    "三大法人近3月買超 ＋ 頭肩底剛形成 ＋ 圓弧底剛形成": {"20": {"n": 33, "moonshotN": 4, "pct": 12.1, "avg": 53.6}},
    "三大法人近3月賣超 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 33, "moonshotN": 4, "pct": 12.1, "avg": 46.7}},
    "多方力道≥80 ＋ N字底剛形成 ＋ 晨星剛形成": {"10": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 42.2}},
    "大盤站上20日均線 ＋ N字底剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 36.2}, "20": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 51}},
    "回後買上漲全通過 ＋ N字底剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 36.2}, "20": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 51}},
    "強勢突破盤 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 42.2}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 大盤站上20日均線": {"20": {"n": 566, "moonshotN": 68, "pct": 12, "avg": 55.9}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為正": {"20": {"n": 460, "moonshotN": 55, "pct": 12, "avg": 56.8}},
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 108, "moonshotN": 13, "pct": 12, "avg": 44.6}},
    "KDJ近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 夜星剛形成": {"20": {"n": 217, "moonshotN": 26, "pct": 12, "avg": 45.9}},
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 42.8}},
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 150, "moonshotN": 18, "pct": 12, "avg": 47.4}},
    "爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 133, "moonshotN": 16, "pct": 12, "avg": 46.1}},
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 117, "moonshotN": 14, "pct": 12, "avg": 59.7}},
    "量能區間高檔(≥90百分位) ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 274, "moonshotN": 33, "pct": 12, "avg": 53.2}},
    "大盤跌破60日均線 ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 225, "moonshotN": 27, "pct": 12, "avg": 49.9}},
    "回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 25, "moonshotN": 3, "pct": 12, "avg": 54.7}},
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"20": {"n": 75, "moonshotN": 9, "pct": 12, "avg": 61.3}},
    "回後買上漲全通過 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 108, "moonshotN": 13, "pct": 12, "avg": 51.8}},
    "近3月均價YoY為正 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 532, "moonshotN": 64, "pct": 12, "avg": 50.5}},
    "多方力道≥65 ＋ 晨星剛形成": {"20": {"n": 707, "moonshotN": 84, "pct": 11.9, "avg": 51.2}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成": {"20": {"n": 746, "moonshotN": 89, "pct": 11.9, "avg": 48.5}},
    "大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 578, "moonshotN": 69, "pct": 11.9, "avg": 49.8}},
    "多方力道≥65 ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 707, "moonshotN": 84, "pct": 11.9, "avg": 51.2}},
    "多方力道≥65 ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 707, "moonshotN": 84, "pct": 11.9, "avg": 51.2}},
    "多方力道≥65 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 707, "moonshotN": 84, "pct": 11.9, "avg": 51.2}},
    "多方力道≥80 ＋ MACD近3日內黃金交叉 ＋ N字底剛形成": {"20": {"n": 84, "moonshotN": 10, "pct": 11.9, "avg": 45.3}},
    "強勢突破盤 ＋ KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量)": {"20": {"n": 109, "moonshotN": 13, "pct": 11.9, "avg": 54.3}},
    "強勢突破盤 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"20": {"n": 59, "moonshotN": 7, "pct": 11.9, "avg": 44.6}},
    "布林通道高檔(≥80%) ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 746, "moonshotN": 89, "pct": 11.9, "avg": 48.5}},
    "布林通道高檔(≥80%) ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 746, "moonshotN": 89, "pct": 11.9, "avg": 48.5}},
    "布林通道高檔(≥80%) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 746, "moonshotN": 89, "pct": 11.9, "avg": 48.5}},
    "KDJ近3日內死亡交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 67, "moonshotN": 8, "pct": 11.9, "avg": 50.8}},
    "大盤跌破20日均線 ＋ 型態成形中 ＋ 夜星剛形成": {"20": {"n": 578, "moonshotN": 69, "pct": 11.9, "avg": 49.8}},
    "大盤跌破20日均線 ＋ 型態突破確認 ＋ 夜星剛形成": {"20": {"n": 578, "moonshotN": 69, "pct": 11.9, "avg": 49.8}},
    "大盤跌破20日均線 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": {"20": {"n": 578, "moonshotN": 69, "pct": 11.9, "avg": 49.8}},
    "近3月均價YoY為正 ＋ N字底剛形成 ＋ 晨星剛形成": {"10": {"n": 34, "moonshotN": 4, "pct": 11.8, "avg": 39.2}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 51, "moonshotN": 6, "pct": 11.8, "avg": 51.8}},
    "多方力道≥65 ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 220, "moonshotN": 26, "pct": 11.8, "avg": 53}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 量能區間高檔(≥90百分位)": {"20": {"n": 186, "moonshotN": 22, "pct": 11.8, "avg": 49.3}},
    "布林通道高檔(≥80%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 372, "moonshotN": 44, "pct": 11.8, "avg": 49.2}},
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 85, "moonshotN": 10, "pct": 11.8, "avg": 51.5}},
    "KDJ近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 161, "moonshotN": 19, "pct": 11.8, "avg": 53}},
    "KDJ近3日內死亡交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 34, "moonshotN": 4, "pct": 11.8, "avg": 50}},
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 127, "moonshotN": 15, "pct": 11.8, "avg": 43}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 169, "moonshotN": 20, "pct": 11.8, "avg": 56.6}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 圓弧底剛形成 ＋ K線橫盤的突破剛形成": {"20": {"n": 34, "moonshotN": 4, "pct": 11.8, "avg": 71.8}},
    "大盤跌破20日均線 ＋ 大盤站上60日均線 ＋ 夜星剛形成": {"20": {"n": 322, "moonshotN": 38, "pct": 11.8, "avg": 49.9}},
    "大盤站上60日均線 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 905, "moonshotN": 107, "pct": 11.8, "avg": 52}},
    "大盤跌破60日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 85, "moonshotN": 10, "pct": 11.8, "avg": 42.8}},
    "型態成形中 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 51, "moonshotN": 6, "pct": 11.8, "avg": 51.8}},
    "型態突破確認 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 51, "moonshotN": 6, "pct": 11.8, "avg": 51.8}},
    "型態剛形成(剛突破) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 51, "moonshotN": 6, "pct": 11.8, "avg": 51.8}},
    "近3月乖離度為正(營收優於股價) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 76, "moonshotN": 9, "pct": 11.8, "avg": 44.7}},
    "多方力道≥65 ＋ 爆量(≥2倍均量) ＋ 晨星剛形成": {"20": {"n": 197, "moonshotN": 23, "pct": 11.7, "avg": 55.5}},
    "多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 145, "moonshotN": 17, "pct": 11.7, "avg": 57.4}},
    "多方力道≥80 ＋ 量能區間低檔(≤10百分位) ＋ 近3月乖離度為負(股價超前營收)": {"20": {"n": 315, "moonshotN": 37, "pct": 11.7, "avg": 47.8}},
    "多方力道≥80 ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"20": {"n": 163, "moonshotN": 19, "pct": 11.7, "avg": 52}},
    "布林通道高檔(≥80%) ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 315, "moonshotN": 37, "pct": 11.7, "avg": 53}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 223, "moonshotN": 26, "pct": 11.7, "avg": 47.1}},
    "地量(≤0.5倍均量) ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成": {"20": {"n": 60, "moonshotN": 7, "pct": 11.7, "avg": 60.9}},
    "地量(≤0.5倍均量) ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 128, "moonshotN": 15, "pct": 11.7, "avg": 52.8}},
    "大盤站上20日均線 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 1023, "moonshotN": 120, "pct": 11.7, "avg": 51.4}},
    "近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 282, "moonshotN": 33, "pct": 11.7, "avg": 50.3}},
    "地量(≤0.5倍均量) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"10": {"n": 43, "moonshotN": 5, "pct": 11.6, "avg": 39.1}, "20": {"n": 43, "moonshotN": 5, "pct": 11.6, "avg": 83.9}},
    "KDJ近3日內死亡交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成": {"10": {"n": 69, "moonshotN": 7, "pct": 10.1, "avg": 36.5}, "20": {"n": 69, "moonshotN": 8, "pct": 11.6, "avg": 48.2}},
    "大盤站上20日均線 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 69, "moonshotN": 7, "pct": 10.1, "avg": 38.5}, "20": {"n": 69, "moonshotN": 8, "pct": 11.6, "avg": 37.6}},
    "地量(≤0.5倍均量) ＋ 回後買上漲全通過": {"20": {"n": 301, "moonshotN": 35, "pct": 11.6, "avg": 58.2}},
    "多方力道≥65 ＋ 強勢突破盤 ＋ 晨星剛形成": {"20": {"n": 292, "moonshotN": 34, "pct": 11.6, "avg": 49.1}},
    "多方力道≥65 ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 121, "moonshotN": 14, "pct": 11.6, "avg": 54.3}},
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 207, "moonshotN": 24, "pct": 11.6, "avg": 55}},
    "多方力道≥80 ＋ 回後買上漲全通過 ＋ N字底剛形成": {"20": {"n": 638, "moonshotN": 74, "pct": 11.6, "avg": 48.5}},
    "布林通道高檔(≥80%) ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 371, "moonshotN": 43, "pct": 11.6, "avg": 51.5}},
    "KDJ近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過": {"20": {"n": 172, "moonshotN": 20, "pct": 11.6, "avg": 66.2}},
    "相對強弱為正(強於大盤) ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 490, "moonshotN": 57, "pct": 11.6, "avg": 53.2}},
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 回後買上漲全通過": {"20": {"n": 301, "moonshotN": 35, "pct": 11.6, "avg": 58.2}},
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 回後買上漲全通過": {"20": {"n": 301, "moonshotN": 35, "pct": 11.6, "avg": 58.2}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"20": {"n": 112, "moonshotN": 13, "pct": 11.6, "avg": 49.1}},
    "回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 155, "moonshotN": 18, "pct": 11.6, "avg": 57.3}},
    "強勢突破盤 ＋ 晨星剛形成": {"20": {"n": 391, "moonshotN": 45, "pct": 11.5, "avg": 48.7}},
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 391, "moonshotN": 45, "pct": 11.5, "avg": 52.5}},
    "強勢突破盤 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"20": {"n": 391, "moonshotN": 45, "pct": 11.5, "avg": 48.7}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 大盤站上60日均線": {"20": {"n": 583, "moonshotN": 67, "pct": 11.5, "avg": 56.9}},
    "強勢突破盤 ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 391, "moonshotN": 45, "pct": 11.5, "avg": 48.7}},
    "強勢突破盤 ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 391, "moonshotN": 45, "pct": 11.5, "avg": 48.7}},
    "強勢突破盤 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 391, "moonshotN": 45, "pct": 11.5, "avg": 48.7}},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 61, "moonshotN": 7, "pct": 11.5, "avg": 46}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 78, "moonshotN": 9, "pct": 11.5, "avg": 45.8}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 61, "moonshotN": 7, "pct": 11.5, "avg": 54.7}},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 113, "moonshotN": 13, "pct": 11.5, "avg": 62.7}},
    "相對強弱為正(強於大盤) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 855, "moonshotN": 98, "pct": 11.5, "avg": 51.8}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 130, "moonshotN": 15, "pct": 11.5, "avg": 56.6}},
    "大盤站上20日均線 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 695, "moonshotN": 80, "pct": 11.5, "avg": 51.2}},
    "大盤站上20日均線 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 122, "moonshotN": 14, "pct": 11.5, "avg": 59.7}},
    "近3月乖離度為負(股價超前營收) ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"20": {"n": 104, "moonshotN": 12, "pct": 11.5, "avg": 50.5}},
    "三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成": {"20": {"n": 209, "moonshotN": 24, "pct": 11.5, "avg": 51.7}},
    "三重底剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 61, "moonshotN": 7, "pct": 11.5, "avg": 55}},
    "多方力道≥65 ＋ N字底剛形成 ＋ 晨星剛形成": {"10": {"n": 35, "moonshotN": 4, "pct": 11.4, "avg": 41.9}},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 44, "moonshotN": 5, "pct": 11.4, "avg": 38.7}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成": {"20": {"n": 175, "moonshotN": 20, "pct": 11.4, "avg": 55.4}},
    "大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 1118, "moonshotN": 128, "pct": 11.4, "avg": 51}},
    "多方力道≥65 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 44, "moonshotN": 5, "pct": 11.4, "avg": 51.8}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 量能區間高檔(≥90百分位)": {"20": {"n": 35, "moonshotN": 4, "pct": 11.4, "avg": 54.6}},
    "強勢突破盤 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 332, "moonshotN": 38, "pct": 11.4, "avg": 49.5}},
    "布林通道高檔(≥80%) ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 79, "moonshotN": 9, "pct": 11.4, "avg": 52.1}},
    "爆量(≥2倍均量) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 35, "moonshotN": 4, "pct": 11.4, "avg": 36.6}},
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 175, "moonshotN": 20, "pct": 11.4, "avg": 55.4}},
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 175, "moonshotN": 20, "pct": 11.4, "avg": 55.4}},
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 175, "moonshotN": 20, "pct": 11.4, "avg": 55.4}},
    "地量(≤0.5倍均量) ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"20": {"n": 44, "moonshotN": 5, "pct": 11.4, "avg": 63.3}},
    "地量(≤0.5倍均量) ＋ 近3月均價YoY為負 ＋ 夜星剛形成": {"20": {"n": 44, "moonshotN": 5, "pct": 11.4, "avg": 50}},
    "量能區間高檔(≥90百分位) ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 44, "moonshotN": 5, "pct": 11.4, "avg": 44.6}},
    "大盤站上20日均線 ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 1118, "moonshotN": 128, "pct": 11.4, "avg": 51}},
    "大盤站上20日均線 ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 1118, "moonshotN": 128, "pct": 11.4, "avg": 51}},
    "大盤站上20日均線 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 1118, "moonshotN": 128, "pct": 11.4, "avg": 51}},
    "大盤站上20日均線 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"20": {"n": 421, "moonshotN": 48, "pct": 11.4, "avg": 50.6}},
    "大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成": {"20": {"n": 167, "moonshotN": 19, "pct": 11.4, "avg": 51.6}},
    "近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成": {"20": {"n": 370, "moonshotN": 42, "pct": 11.4, "avg": 55.4}},
    "大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 345, "moonshotN": 39, "pct": 11.3, "avg": 50}},
    "多方力道≥80 ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成": {"20": {"n": 177, "moonshotN": 20, "pct": 11.3, "avg": 48.7}},
    "多方力道≥80 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 238, "moonshotN": 27, "pct": 11.3, "avg": 59.5}},
    "強勢突破盤 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 124, "moonshotN": 14, "pct": 11.3, "avg": 47.6}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ N字底剛形成": {"20": {"n": 71, "moonshotN": 8, "pct": 11.3, "avg": 65}},
    "布林通道高檔(≥80%) ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 497, "moonshotN": 56, "pct": 11.3, "avg": 48.5}},
    "KDJ近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 213, "moonshotN": 24, "pct": 11.3, "avg": 43.4}},
    "相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 222, "moonshotN": 25, "pct": 11.3, "avg": 51.2}},
    "爆量(≥2倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 62, "moonshotN": 7, "pct": 11.3, "avg": 45.8}},
    "大盤站上60日均線 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 797, "moonshotN": 90, "pct": 11.3, "avg": 49.8}},
    "大盤跌破60日均線 ＋ 型態成形中 ＋ 夜星剛形成": {"20": {"n": 345, "moonshotN": 39, "pct": 11.3, "avg": 50}},
    "大盤跌破60日均線 ＋ 型態突破確認 ＋ 夜星剛形成": {"20": {"n": 345, "moonshotN": 39, "pct": 11.3, "avg": 50}},
    "大盤跌破60日均線 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": {"20": {"n": 345, "moonshotN": 39, "pct": 11.3, "avg": 50}},
    "近3月均價YoY為負 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 160, "moonshotN": 18, "pct": 11.3, "avg": 52.8}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量)": {"20": {"n": 676, "moonshotN": 76, "pct": 11.2, "avg": 56.3}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"20": {"n": 887, "moonshotN": 99, "pct": 11.2, "avg": 52.9}},
    "強勢突破盤 ＋ 布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量)": {"20": {"n": 676, "moonshotN": 76, "pct": 11.2, "avg": 56.3}},
    "強勢突破盤 ＋ KDJ近3日內黃金交叉 ＋ 地量(≤0.5倍均量)": {"20": {"n": 233, "moonshotN": 26, "pct": 11.2, "avg": 58}},
    "布林通道高檔(≥80%) ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 546, "moonshotN": 61, "pct": 11.2, "avg": 50.9}},
    "MACD近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 89, "moonshotN": 10, "pct": 11.2, "avg": 47.5}},
    "量能區間高檔(≥90百分位) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 116, "moonshotN": 13, "pct": 11.2, "avg": 50.6}},
    "大盤站上60日均線 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 143, "moonshotN": 16, "pct": 11.2, "avg": 59.1}},
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 402, "moonshotN": 45, "pct": 11.2, "avg": 49.4}},
    "近3月均價YoY為負 ＋ 三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 143, "moonshotN": 16, "pct": 11.2, "avg": 54.9}},
    "三大法人近3月買超 ＋ N字底剛形成 ＋ 晨星剛形成": {"10": {"n": 36, "moonshotN": 4, "pct": 11.1, "avg": 39.2}},
    "多方力道≥65 ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 395, "moonshotN": 44, "pct": 11.1, "avg": 55.6}},
    "多方力道≥65 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 407, "moonshotN": 45, "pct": 11.1, "avg": 53.3}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ N字底剛形成": {"20": {"n": 108, "moonshotN": 12, "pct": 11.1, "avg": 56.3}},
    "多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 72, "moonshotN": 8, "pct": 11.1, "avg": 57.8}},
    "強勢突破盤 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 63, "moonshotN": 7, "pct": 11.1, "avg": 46.9}},
    "布林通道高檔(≥80%) ＋ 相對強弱為負(弱於大盤) ＋ 晨星剛形成": {"20": {"n": 108, "moonshotN": 12, "pct": 11.1, "avg": 47}},
    "布林通道高檔(≥80%) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 45, "moonshotN": 5, "pct": 11.1, "avg": 48.8}},
    "布林通道高檔(≥80%) ＋ 突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 54, "moonshotN": 6, "pct": 11.1, "avg": 51.1}},
    "MACD近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 108, "moonshotN": 12, "pct": 11.1, "avg": 48.2}},
    "MACD近3日內黃金交叉 ＋ 三重底剛形成 ＋ 晨星剛形成": {"20": {"n": 27, "moonshotN": 3, "pct": 11.1, "avg": 40.9}},
    "地量(≤0.5倍均量) ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成": {"20": {"n": 72, "moonshotN": 8, "pct": 11.1, "avg": 55.1}},
    "大盤跌破20日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成": {"20": {"n": 217, "moonshotN": 24, "pct": 11.1, "avg": 48.8}},
    "近3月均價YoY為正 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"20": {"n": 27, "moonshotN": 3, "pct": 11.1, "avg": 42.2}},
    "大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 1299, "moonshotN": 143, "pct": 11, "avg": 50.3}},
    "多方力道≥65 ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成": {"20": {"n": 301, "moonshotN": 33, "pct": 11, "avg": 52.1}},
    "多方力道≥65 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 127, "moonshotN": 14, "pct": 11, "avg": 62.1}},
    "多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 127, "moonshotN": 14, "pct": 11, "avg": 57.9}},
    "強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過": {"20": {"n": 381, "moonshotN": 42, "pct": 11, "avg": 58}},
    "強勢突破盤 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 136, "moonshotN": 15, "pct": 11, "avg": 61.8}},
    "MACD近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 91, "moonshotN": 10, "pct": 11, "avg": 49}},
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 118, "moonshotN": 13, "pct": 11, "avg": 48.1}},
    "相對強弱為正(強於大盤) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 218, "moonshotN": 24, "pct": 11, "avg": 54.4}},
    "相對強弱為正(強於大盤) ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 748, "moonshotN": 82, "pct": 11, "avg": 51}},
    "爆量(≥2倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 173, "moonshotN": 19, "pct": 11, "avg": 51.5}},
    "爆量(≥2倍均量) ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 191, "moonshotN": 21, "pct": 11, "avg": 51.8}},
    "量能區間高檔(≥90百分位) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 309, "moonshotN": 34, "pct": 11, "avg": 53.6}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 400, "moonshotN": 44, "pct": 11, "avg": 48.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 100, "moonshotN": 11, "pct": 11, "avg": 49.8}},
    "大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 356, "moonshotN": 39, "pct": 11, "avg": 49.5}},
    "大盤站上60日均線 ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 1299, "moonshotN": 143, "pct": 11, "avg": 50.3}},
    "大盤站上60日均線 ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 1299, "moonshotN": 143, "pct": 11, "avg": 50.3}},
    "大盤站上60日均線 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 1299, "moonshotN": 143, "pct": 11, "avg": 50.3}},
    "大盤跌破60日均線 ＋ 近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 282, "moonshotN": 31, "pct": 11, "avg": 51.7}},
    "大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成": {"20": {"n": 155, "moonshotN": 17, "pct": 11, "avg": 53.8}},
    "回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 164, "moonshotN": 18, "pct": 11, "avg": 53.4}},
    "近3月均價YoY為正 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 100, "moonshotN": 11, "pct": 11, "avg": 65}},
    "多方力道≥65 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 338, "moonshotN": 37, "pct": 10.9, "avg": 52.1}},
    "多方力道≥80 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 442, "moonshotN": 48, "pct": 10.9, "avg": 55.9}},
    "多方力道≥80 ＋ 回後買上漲全通過 ＋ 複式頭肩底剛形成": {"20": {"n": 46, "moonshotN": 5, "pct": 10.9, "avg": 39.3}},
    "多方力道≥80 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 201, "moonshotN": 22, "pct": 10.9, "avg": 62.1}},
    "強勢突破盤 ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成": {"20": {"n": 221, "moonshotN": 24, "pct": 10.9, "avg": 46.3}},
    "MACD近3日內黃金交叉 ＋ 相對強弱為負(弱於大盤) ＋ 夜星剛形成": {"20": {"n": 137, "moonshotN": 15, "pct": 10.9, "avg": 49.6}},
    "相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成": {"20": {"n": 606, "moonshotN": 66, "pct": 10.9, "avg": 51.8}},
    "爆量(≥2倍均量) ＋ 量能區間高檔(≥90百分位) ＋ 晨星剛形成": {"20": {"n": 304, "moonshotN": 33, "pct": 10.9, "avg": 50.4}},
    "爆量(≥2倍均量) ＋ 近3月乖離度為正(營收優於股價) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 119, "moonshotN": 13, "pct": 10.9, "avg": 46.2}},
    "大盤跌破20日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成": {"20": {"n": 312, "moonshotN": 34, "pct": 10.9, "avg": 51.4}},
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 46, "moonshotN": 5, "pct": 10.9, "avg": 54.8}},
    "近3月乖離度為正(營收優於股價) ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 202, "moonshotN": 22, "pct": 10.9, "avg": 49.4}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 201, "moonshotN": 22, "pct": 10.9, "avg": 54.2}},
    "三大法人近3月賣超 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 46, "moonshotN": 5, "pct": 10.9, "avg": 37}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成": {"10": {"n": 37, "moonshotN": 4, "pct": 10.8, "avg": 38}},
    "爆量(≥2倍均量) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 37, "moonshotN": 4, "pct": 10.8, "avg": 36.1}},
    "近3月乖離度為正(營收優於股價) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 37, "moonshotN": 4, "pct": 10.8, "avg": 37}},
    "多方力道≥65 ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成": {"20": {"n": 269, "moonshotN": 29, "pct": 10.8, "avg": 48.3}},
    "多方力道≥80 ＋ KDJ近3日內死亡交叉 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 333, "moonshotN": 36, "pct": 10.8, "avg": 56.6}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"20": {"n": 295, "moonshotN": 32, "pct": 10.8, "avg": 57.6}},
    "布林通道高檔(≥80%) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"20": {"n": 139, "moonshotN": 15, "pct": 10.8, "avg": 47.4}},
    "相對強弱為負(弱於大盤) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 223, "moonshotN": 24, "pct": 10.8, "avg": 45.1}},
    "爆量(≥1.5倍均量) ＋ N字底剛形成 ＋ 晨星剛形成": {"20": {"n": 37, "moonshotN": 4, "pct": 10.8, "avg": 63.7}},
    "爆量(≥2倍均量) ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 83, "moonshotN": 9, "pct": 10.8, "avg": 35.9}},
    "回後買上漲全通過 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 213, "moonshotN": 23, "pct": 10.8, "avg": 51}},
    "三大法人近3月買超 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 111, "moonshotN": 12, "pct": 10.8, "avg": 65.7}},
    "大盤站上60日均線 ＋ N字底剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 28, "moonshotN": 3, "pct": 10.7, "avg": 36.2}, "20": {"n": 28, "moonshotN": 3, "pct": 10.7, "avg": 51}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 84, "moonshotN": 9, "pct": 10.7, "avg": 42.1}},
    "多方力道≥80 ＋ 回後買上漲全通過 ＋ 三大法人近3月買超": {"20": {"n": 3195, "moonshotN": 342, "pct": 10.7, "avg": 51.3}},
    "強勢突破盤 ＋ 回後買上漲全通過 ＋ 晨星剛形成": {"20": {"n": 177, "moonshotN": 19, "pct": 10.7, "avg": 52.5}},
    "布林通道高檔(≥80%) ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成": {"20": {"n": 364, "moonshotN": 39, "pct": 10.7, "avg": 47.4}},
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 140, "moonshotN": 15, "pct": 10.7, "avg": 45.5}},
    "KDJ近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 663, "moonshotN": 71, "pct": 10.7, "avg": 51.3}},
    "KDJ近3日內死亡交叉 ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 196, "moonshotN": 21, "pct": 10.7, "avg": 46.4}},
    "爆量(≥2倍均量) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 308, "moonshotN": 33, "pct": 10.7, "avg": 49.6}},
    "量能區間高檔(≥90百分位) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 131, "moonshotN": 14, "pct": 10.7, "avg": 58.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤站上60日均線 ＋ 夜星剛形成": {"20": {"n": 140, "moonshotN": 15, "pct": 10.7, "avg": 49.4}},
    "型態成形中 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 84, "moonshotN": 9, "pct": 10.7, "avg": 42.1}},
    "型態突破確認 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 84, "moonshotN": 9, "pct": 10.7, "avg": 42.1}},
    "型態剛形成(剛突破) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 84, "moonshotN": 9, "pct": 10.7, "avg": 42.1}},
    "圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成": {"20": {"n": 122, "moonshotN": 13, "pct": 10.7, "avg": 58.8}},
    "突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 56, "moonshotN": 6, "pct": 10.7, "avg": 47.2}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成": {"20": {"n": 218, "moonshotN": 23, "pct": 10.6, "avg": 49.7}},
    "三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 784, "moonshotN": 83, "pct": 10.6, "avg": 49.7}},
    "多方力道≥80 ＋ KDJ近3日內黃金交叉 ＋ N字底剛形成": {"20": {"n": 720, "moonshotN": 76, "pct": 10.6, "avg": 49.4}},
    "多方力道≥80 ＋ 量能區間高檔(≥90百分位) ＋ N字底剛形成": {"20": {"n": 718, "moonshotN": 76, "pct": 10.6, "avg": 47.4}},
    "強勢突破盤 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 263, "moonshotN": 28, "pct": 10.6, "avg": 48.2}},
    "布林通道高檔(≥80%) ＋ 相對強弱為正(強於大盤) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 489, "moonshotN": 52, "pct": 10.6, "avg": 50}},
    "布林通道高檔(≥80%) ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 405, "moonshotN": 43, "pct": 10.6, "avg": 50.5}},
    "MACD近3日內黃金交叉 ＋ 型態成形中 ＋ 夜星剛形成": {"20": {"n": 218, "moonshotN": 23, "pct": 10.6, "avg": 49.7}},
    "MACD近3日內黃金交叉 ＋ 型態突破確認 ＋ 夜星剛形成": {"20": {"n": 218, "moonshotN": 23, "pct": 10.6, "avg": 49.7}},
    "MACD近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": {"20": {"n": 218, "moonshotN": 23, "pct": 10.6, "avg": 49.7}},
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 161, "moonshotN": 17, "pct": 10.6, "avg": 56.2}},
    "爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 47, "moonshotN": 5, "pct": 10.6, "avg": 45.2}},
    "爆量(≥2倍均量) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 47, "moonshotN": 5, "pct": 10.6, "avg": 38}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 151, "moonshotN": 16, "pct": 10.6, "avg": 44.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成": {"20": {"n": 142, "moonshotN": 15, "pct": 10.6, "avg": 52.8}},
    "大盤站上60日均線 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"20": {"n": 500, "moonshotN": 53, "pct": 10.6, "avg": 51.2}},
    "型態成形中 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 784, "moonshotN": 83, "pct": 10.6, "avg": 49.7}},
    "型態突破確認 ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 784, "moonshotN": 83, "pct": 10.6, "avg": 49.7}},
    "型態剛形成(剛突破) ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 784, "moonshotN": 83, "pct": 10.6, "avg": 49.7}},
    "近3月均價YoY為正 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 47, "moonshotN": 5, "pct": 10.6, "avg": 41.8}},
    "多方力道≥65 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 315, "moonshotN": 33, "pct": 10.5, "avg": 54.4}},
    "多方力道≥65 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 38, "moonshotN": 4, "pct": 10.5, "avg": 38.9}},
    "多方力道≥80 ＋ 量能區間高檔(≥90百分位) ＋ 晨星剛形成": {"20": {"n": 275, "moonshotN": 29, "pct": 10.5, "avg": 54.1}},
    "多方力道≥80 ＋ 大盤站上20日均線 ＋ 回後買上漲全通過": {"20": {"n": 3550, "moonshotN": 374, "pct": 10.5, "avg": 50.5}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 143, "moonshotN": 15, "pct": 10.5, "avg": 59.1}},
    "多方力道≥80 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 295, "moonshotN": 31, "pct": 10.5, "avg": 59.6}},
    "強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 突破ABC修正下降切線剛形成": {"20": {"n": 95, "moonshotN": 10, "pct": 10.5, "avg": 57.3}},
    "MACD近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 晨星剛形成": {"20": {"n": 105, "moonshotN": 11, "pct": 10.5, "avg": 49.8}},
    "MACD近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成": {"20": {"n": 124, "moonshotN": 13, "pct": 10.5, "avg": 53.7}},
    "KDJ近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 124, "moonshotN": 13, "pct": 10.5, "avg": 45.4}},
    "KDJ近3日內黃金交叉 ＋ 近3月均價YoY為正 ＋ 夜星剛形成": {"20": {"n": 304, "moonshotN": 32, "pct": 10.5, "avg": 44.2}},
    "相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 181, "moonshotN": 19, "pct": 10.5, "avg": 58.5}},
    "爆量(≥1.5倍均量) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 153, "moonshotN": 16, "pct": 10.5, "avg": 57.9}},
    "爆量(≥2倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 276, "moonshotN": 29, "pct": 10.5, "avg": 50.5}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 440, "moonshotN": 46, "pct": 10.5, "avg": 48.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤站上20日均線 ＋ 夜星剛形成": {"20": {"n": 86, "moonshotN": 9, "pct": 10.5, "avg": 47.7}},
    "大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 315, "moonshotN": 33, "pct": 10.5, "avg": 51.6}},
    "近3月乖離度為負(股價超前營收) ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 602, "moonshotN": 63, "pct": 10.5, "avg": 51.2}},
    "近3月均價YoY為負 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 76, "moonshotN": 8, "pct": 10.5, "avg": 47.1}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 625, "moonshotN": 65, "pct": 10.4, "avg": 50.1}},
    "多方力道≥65 ＋ KDJ近3日內黃金交叉 ＋ 夜星剛形成": {"20": {"n": 164, "moonshotN": 17, "pct": 10.4, "avg": 46.9}},
    "多方力道≥65 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 560, "moonshotN": 58, "pct": 10.4, "avg": 54.4}},
    "多方力道≥80 ＋ 回後買上漲全通過 ＋ 近3月均價YoY為負": {"20": {"n": 1578, "moonshotN": 164, "pct": 10.4, "avg": 50.3}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"20": {"n": 125, "moonshotN": 13, "pct": 10.4, "avg": 58.9}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 型態成形中": {"20": {"n": 605, "moonshotN": 63, "pct": 10.4, "avg": 56.6}},
    "布林通道高檔(≥80%) ＋ 爆量(≥2倍均量) ＋ 晨星剛形成": {"20": {"n": 308, "moonshotN": 32, "pct": 10.4, "avg": 50.5}},
    "布林通道高檔(≥80%) ＋ 量能區間高檔(≥90百分位) ＋ 晨星剛形成": {"20": {"n": 491, "moonshotN": 51, "pct": 10.4, "avg": 50.8}},
    "布林通道高檔(≥80%) ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 625, "moonshotN": 65, "pct": 10.4, "avg": 50.1}},
    "布林通道高檔(≥80%) ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 625, "moonshotN": 65, "pct": 10.4, "avg": 50.1}},
    "布林通道高檔(≥80%) ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 625, "moonshotN": 65, "pct": 10.4, "avg": 50.1}},
    "布林通道高檔(≥80%) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 163, "moonshotN": 17, "pct": 10.4, "avg": 59.1}},
    "MACD近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 48, "moonshotN": 5, "pct": 10.4, "avg": 51.6}},
    "MACD近3日內黃金交叉 ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 106, "moonshotN": 11, "pct": 10.4, "avg": 50.1}},
    "MACD近3日內黃金交叉 ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 135, "moonshotN": 14, "pct": 10.4, "avg": 51.6}},
    "KDJ近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 夜星剛形成": {"20": {"n": 307, "moonshotN": 32, "pct": 10.4, "avg": 43.7}},
    "KDJ近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 夜星剛形成": {"20": {"n": 135, "moonshotN": 14, "pct": 10.4, "avg": 44.8}},
    "KDJ近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 550, "moonshotN": 57, "pct": 10.4, "avg": 52.5}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 347, "moonshotN": 36, "pct": 10.4, "avg": 49.6}},
    "相對強弱為正(強於大盤) ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 766, "moonshotN": 80, "pct": 10.4, "avg": 52.6}},
    "爆量(≥2倍均量) ＋ 近3月均價YoY為負 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 106, "moonshotN": 11, "pct": 10.4, "avg": 48.6}},
    "量能區間高檔(≥90百分位) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 536, "moonshotN": 56, "pct": 10.4, "avg": 51.5}},
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 240, "moonshotN": 25, "pct": 10.4, "avg": 49.1}},
    "回後買上漲全通過 ＋ 近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 651, "moonshotN": 68, "pct": 10.4, "avg": 53.6}},
    "近3月均價YoY為正 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"20": {"n": 814, "moonshotN": 85, "pct": 10.4, "avg": 49.6}},
    "相對強弱為正(強於大盤) ＋ 晨星剛形成": {"20": {"n": 1168, "moonshotN": 120, "pct": 10.3, "avg": 50.4}},
    "多方力道≥65 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 397, "moonshotN": 41, "pct": 10.3, "avg": 56.6}},
    "多方力道≥80 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 146, "moonshotN": 15, "pct": 10.3, "avg": 54.1}},
    "多方力道≥80 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"20": {"n": 116, "moonshotN": 12, "pct": 10.3, "avg": 55.2}},
    "布林通道高檔(≥80%) ＋ 爆量(≥1.5倍均量) ＋ 晨星剛形成": {"20": {"n": 467, "moonshotN": 48, "pct": 10.3, "avg": 49.3}},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 夜星剛形成": {"20": {"n": 97, "moonshotN": 10, "pct": 10.3, "avg": 56.5}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 頭肩底剛形成": {"20": {"n": 29, "moonshotN": 3, "pct": 10.3, "avg": 58.5}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 夜星剛形成": {"20": {"n": 68, "moonshotN": 7, "pct": 10.3, "avg": 53.4}},
    "MACD近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 136, "moonshotN": 14, "pct": 10.3, "avg": 47.9}},
    "相對強弱為正(強於大盤) ＋ 型態成形中 ＋ 晨星剛形成": {"20": {"n": 1168, "moonshotN": 120, "pct": 10.3, "avg": 50.4}},
    "相對強弱為正(強於大盤) ＋ 型態突破確認 ＋ 晨星剛形成": {"20": {"n": 1168, "moonshotN": 120, "pct": 10.3, "avg": 50.4}},
    "相對強弱為正(強於大盤) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": {"20": {"n": 1168, "moonshotN": 120, "pct": 10.3, "avg": 50.4}},
    "相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成": {"20": {"n": 418, "moonshotN": 43, "pct": 10.3, "avg": 57}},
    "爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 圓弧底剛形成": {"20": {"n": 97, "moonshotN": 10, "pct": 10.3, "avg": 57.7}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 夜星剛形成": {"20": {"n": 87, "moonshotN": 9, "pct": 10.3, "avg": 48.3}},
    "地量(≤0.5倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 203, "moonshotN": 21, "pct": 10.3, "avg": 50.1}},
    "地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"20": {"n": 58, "moonshotN": 6, "pct": 10.3, "avg": 45.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成": {"20": {"n": 68, "moonshotN": 7, "pct": 10.3, "avg": 59.1}},
    "大盤站上20日均線 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 39, "moonshotN": 4, "pct": 10.3, "avg": 49.2}},
    "大盤跌破60日均線 ＋ 頭肩底剛形成 ＋ 複式頭肩底剛形成": {"20": {"n": 29, "moonshotN": 3, "pct": 10.3, "avg": 55}},
    "大盤跌破60日均線 ＋ 複式頭肩底剛形成 ＋ K線橫盤的突破剛形成": {"20": {"n": 29, "moonshotN": 3, "pct": 10.3, "avg": 46.4}},
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成": {"20": {"n": 117, "moonshotN": 12, "pct": 10.3, "avg": 50.1}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 177, "moonshotN": 18, "pct": 10.2, "avg": 46.9}},
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 166, "moonshotN": 17, "pct": 10.2, "avg": 59.1}},
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量)": {"20": {"n": 2446, "moonshotN": 249, "pct": 10.2, "avg": 53.9}},
    "多方力道≥65 ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 108, "moonshotN": 11, "pct": 10.2, "avg": 57.4}},
    "多方力道≥65 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 264, "moonshotN": 27, "pct": 10.2, "avg": 50.4}},
    "多方力道≥65 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"20": {"n": 108, "moonshotN": 11, "pct": 10.2, "avg": 44.2}},
    "強勢突破盤 ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 137, "moonshotN": 14, "pct": 10.2, "avg": 44.1}},
    "強勢突破盤 ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"20": {"n": 166, "moonshotN": 17, "pct": 10.2, "avg": 46.5}},
    "MACD近3日內黃金交叉 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 177, "moonshotN": 18, "pct": 10.2, "avg": 46.9}},
    "MACD近3日內黃金交叉 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 177, "moonshotN": 18, "pct": 10.2, "avg": 46.9}},
    "MACD近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 177, "moonshotN": 18, "pct": 10.2, "avg": 46.9}},
    "KDJ近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 187, "moonshotN": 19, "pct": 10.2, "avg": 55.7}},
    "相對強弱為正(強於大盤) ＋ 三重底剛形成 ＋ 晨星剛形成": {"20": {"n": 49, "moonshotN": 5, "pct": 10.2, "avg": 51.4}},
    "相對強弱為負(弱於大盤) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 443, "moonshotN": 45, "pct": 10.2, "avg": 47.2}},
    "爆量(≥1.5倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 462, "moonshotN": 47, "pct": 10.2, "avg": 50.4}},
    "爆量(≥2倍均量) ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 127, "moonshotN": 13, "pct": 10.2, "avg": 51.6}},
    "爆量(≥2倍均量) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 118, "moonshotN": 12, "pct": 10.2, "avg": 57}},
    "量能區間高檔(≥90百分位) ＋ 近3月均價YoY為正 ＋ 晨星剛形成": {"20": {"n": 402, "moonshotN": 41, "pct": 10.2, "avg": 51.4}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 夜星剛形成": {"20": {"n": 206, "moonshotN": 21, "pct": 10.2, "avg": 50.5}},
    "型態成形中 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 166, "moonshotN": 17, "pct": 10.2, "avg": 59.1}},
    "型態突破確認 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 166, "moonshotN": 17, "pct": 10.2, "avg": 59.1}},
    "型態剛形成(剛突破) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 166, "moonshotN": 17, "pct": 10.2, "avg": 59.1}},
    "回後買上漲全通過 ＋ 近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 539, "moonshotN": 55, "pct": 10.2, "avg": 51.7}},
    "近3月均價YoY為正 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成": {"20": {"n": 245, "moonshotN": 25, "pct": 10.2, "avg": 49.1}},
    "相對強弱為正(強於大盤) ＋ 夜星剛形成": {"20": {"n": 818, "moonshotN": 83, "pct": 10.1, "avg": 55}},
    "爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 237, "moonshotN": 24, "pct": 10.1, "avg": 50.2}},
    "多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過": {"20": {"n": 661, "moonshotN": 67, "pct": 10.1, "avg": 54.8}},
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 228, "moonshotN": 23, "pct": 10.1, "avg": 57.1}},
    "多方力道≥80 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 159, "moonshotN": 16, "pct": 10.1, "avg": 56.1}},
    "多方力道≥80 ＋ 相對強弱為正(強於大盤) ＋ 回後買上漲全通過": {"20": {"n": 4080, "moonshotN": 413, "pct": 10.1, "avg": 51.3}},
    "多方力道≥80 ＋ 爆量(≥1.5倍均量) ＋ N字底剛形成": {"20": {"n": 605, "moonshotN": 61, "pct": 10.1, "avg": 48.5}},
    "多方力道≥80 ＋ 大盤站上60日均線 ＋ 回後買上漲全通過": {"20": {"n": 3662, "moonshotN": 371, "pct": 10.1, "avg": 51.6}},
    "多方力道≥80 ＋ 回後買上漲全通過 ＋ 近3月乖離度為正(營收優於股價)": {"20": {"n": 2024, "moonshotN": 205, "pct": 10.1, "avg": 50.1}},
    "多方力道≥80 ＋ 近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 672, "moonshotN": 68, "pct": 10.1, "avg": 52.9}},
    "強勢突破盤 ＋ 相對強弱為正(強於大盤) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 198, "moonshotN": 20, "pct": 10.1, "avg": 48.1}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為負": {"20": {"n": 208, "moonshotN": 21, "pct": 10.1, "avg": 55}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 晨星剛形成": {"20": {"n": 69, "moonshotN": 7, "pct": 10.1, "avg": 37.8}},
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 夜星剛形成": {"20": {"n": 99, "moonshotN": 10, "pct": 10.1, "avg": 51.9}},
    "相對強弱為正(強於大盤) ＋ 型態成形中 ＋ 夜星剛形成": {"20": {"n": 818, "moonshotN": 83, "pct": 10.1, "avg": 55}},
    "相對強弱為正(強於大盤) ＋ 型態突破確認 ＋ 夜星剛形成": {"20": {"n": 818, "moonshotN": 83, "pct": 10.1, "avg": 55}},
    "相對強弱為正(強於大盤) ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": {"20": {"n": 818, "moonshotN": 83, "pct": 10.1, "avg": 55}},
    "爆量(≥1.5倍均量) ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 237, "moonshotN": 24, "pct": 10.1, "avg": 50.2}},
    "爆量(≥1.5倍均量) ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 358, "moonshotN": 36, "pct": 10.1, "avg": 49.2}},
    "爆量(≥1.5倍均量) ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"20": {"n": 296, "moonshotN": 30, "pct": 10.1, "avg": 50.5}},
    "爆量(≥2倍均量) ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 237, "moonshotN": 24, "pct": 10.1, "avg": 50.2}},
    "爆量(≥2倍均量) ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 237, "moonshotN": 24, "pct": 10.1, "avg": 50.2}},
    "爆量(≥2倍均量) ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 237, "moonshotN": 24, "pct": 10.1, "avg": 50.2}},
    "量能區間高檔(≥90百分位) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成": {"20": {"n": 366, "moonshotN": 37, "pct": 10.1, "avg": 48.9}},
    "量能區間高檔(≥90百分位) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"20": {"n": 595, "moonshotN": 60, "pct": 10.1, "avg": 50.3}},
    "大盤跌破20日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 99, "moonshotN": 10, "pct": 10.1, "avg": 53.5}},
    "大盤跌破60日均線 ＋ 回後買上漲全通過 ＋ 複式頭肩底剛形成": {"20": {"n": 69, "moonshotN": 7, "pct": 10.1, "avg": 50.1}},
    "近3月乖離度為負(股價超前營收) ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"20": {"n": 405, "moonshotN": 41, "pct": 10.1, "avg": 48.9}},
    "三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成": {"20": {"n": 218, "moonshotN": 22, "pct": 10.1, "avg": 42.7}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": {"5": {"n": 30, "moonshotN": 3, "pct": 10, "avg": 34}, "10": {"n": 30, "moonshotN": 3, "pct": 10, "avg": 46.3}},
    "量能區間低檔(≤10百分位) ＋ 大盤跌破60日均線 ＋ 三重底剛形成": {"10": {"n": 20, "moonshotN": 2, "pct": 10, "avg": 35.1}, "20": {"n": 20, "moonshotN": 2, "pct": 10, "avg": 47.7}},
    "大盤站上20日均線 ＋ 複式頭肩底剛形成 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 20, "moonshotN": 2, "pct": 10, "avg": 34}, "20": {"n": 20, "moonshotN": 2, "pct": 10, "avg": 36.5}},
    "大盤站上60日均線 ＋ 頭肩底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"10": {"n": 20, "moonshotN": 2, "pct": 10, "avg": 53.1}, "20": {"n": 20, "moonshotN": 2, "pct": 10, "avg": 63.4}},
    "多方力道≥65 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 320, "moonshotN": 32, "pct": 10, "avg": 49.5}},
    "多方力道≥80 ＋ MACD近3日內黃金交叉 ＋ 量能區間低檔(≤10百分位)": {"20": {"n": 20, "moonshotN": 2, "pct": 10, "avg": 31}},
    "多方力道≥80 ＋ 大盤站上20日均線 ＋ N字底剛形成": {"20": {"n": 959, "moonshotN": 96, "pct": 10, "avg": 46.7}},
    "多方力道≥80 ＋ 三大法人近3月買超 ＋ N字底剛形成": {"20": {"n": 906, "moonshotN": 91, "pct": 10, "avg": 48.2}},
    "布林通道高檔(≥80%) ＋ 近3月均價YoY為負 ＋ 晨星剛形成": {"20": {"n": 269, "moonshotN": 27, "pct": 10, "avg": 46.3}},
    "布林通道高檔(≥80%) ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 220, "moonshotN": 22, "pct": 10, "avg": 49.3}},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成": {"20": {"n": 110, "moonshotN": 11, "pct": 10, "avg": 56.6}},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 80, "moonshotN": 8, "pct": 10, "avg": 44.3}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"20": {"n": 40, "moonshotN": 4, "pct": 10, "avg": 44}},
    "MACD近3日內黃金交叉 ＋ 近3月均價YoY為負 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 70, "moonshotN": 7, "pct": 10, "avg": 41.7}},
    "KDJ近3日內死亡交叉 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"20": {"n": 400, "moonshotN": 40, "pct": 10, "avg": 47.1}},
    "爆量(≥1.5倍均量) ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"20": {"n": 329, "moonshotN": 33, "pct": 10, "avg": 48.2}},
    "爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過": {"20": {"n": 110, "moonshotN": 11, "pct": 10, "avg": 46.4}},
    "地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 夜星剛形成": {"20": {"n": 40, "moonshotN": 4, "pct": 10, "avg": 59.3}},
    "地量(≤0.5倍均量) ＋ 三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成": {"20": {"n": 249, "moonshotN": 25, "pct": 10, "avg": 55}},
    "量能區間高檔(≥90百分位) ＋ 頭肩底剛形成 ＋ 圓弧底剛形成": {"20": {"n": 40, "moonshotN": 4, "pct": 10, "avg": 53.6}},
    "大盤站上20日均線 ＋ 回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成": {"20": {"n": 979, "moonshotN": 98, "pct": 10, "avg": 52.4}},
    "大盤站上20日均線 ＋ 回後買上漲全通過 ＋ 晨星剛形成": {"20": {"n": 270, "moonshotN": 27, "pct": 10, "avg": 50.6}},
    "大盤跌破20日均線 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"20": {"n": 40, "moonshotN": 4, "pct": 10, "avg": 43}},
    "大盤站上60日均線 ＋ 回後買上漲全通過 ＋ 晨星剛形成": {"20": {"n": 279, "moonshotN": 28, "pct": 10, "avg": 51.7}},
    "回後買上漲全通過 ＋ N字底剛形成 ＋ 晨星剛形成": {"20": {"n": 50, "moonshotN": 5, "pct": 10, "avg": 57.7}},
    "近3月乖離度為正(營收優於股價) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"20": {"n": 60, "moonshotN": 6, "pct": 10, "avg": 43.4}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 地量(≤0.5倍均量)": {"10": {"n": 301, "moonshotN": 33, "pct": 11.0, "avg": 46.4}, "20": {"n": 301, "moonshotN": 63, "pct": 20.9, "avg": 55.5}},
    "多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"10": {"n": 115, "moonshotN": 12, "pct": 10.4, "avg": 44.4}, "20": {"n": 115, "moonshotN": 24, "pct": 20.9, "avg": 49.2}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 近3日向上跳空缺口": {"10": {"n": 359, "moonshotN": 42, "pct": 11.7, "avg": 44.9}, "20": {"n": 358, "moonshotN": 69, "pct": 19.3, "avg": 57.6}},
}




# ── 回測自動入選組合（2026-10-05 三段驗證；原 2026-09-30，台股 2023-10～2026-09 三年回測）──
# W＝回測高勝率：5/10/20日任一天期勝率>60% 且 t值(同日調整)>2（樣本數≥50）；依最高 t 值排序
# F＝回測飆股：10/20日飆股(漲幅>30%)比例>10% 且飆股次數≥10；依最高飆股比例排序
# 已在 S／#／M 清單裡的組合不重複列入；含已刪除條件（布林通道收窄）的組合不列入
BT_WIN_COMBOS = [
    ["大盤跌破60日均線", "月營收連3月年增", "月營收創24個月新高"],
    ["近3日漲停", "月營收連3月年增", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "均線多頭排列(5>20>60且站上月線)"],
    ["近3日向上跳空缺口", "近3日漲停", "月營收創24個月新高"],
    ["近3日漲停", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "漲時量≥跌時量1.5倍(近20日)"],
    ["近3日漲停", "三大法人近3月買超", "月營收創24個月新高"],
    ["強勢突破盤", "近3日漲停", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "外資近5日買超"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "三大法人近3月買超"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "近3月均價YoY為正"],
    ["強勢突破盤", "創52週新高", "月營收創24個月新高"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)"],
    ["近3日向上跳空缺口", "月營收連3月年增", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "近3月乖離度為負(股價超前營收)"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3日漲停"],
    ["地量(≤0.5倍均量)", "創52週新高", "近3月均價YoY為負"],
    ["大盤跌破60日均線", "月營收創24個月新高", "外資近5日買超"],
    ["多方力道≥80", "量能區間低檔(≤10百分位)", "月營收連3月年增"],
    ["回後買上漲全通過", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["多方力道≥65", "量能區間低檔(≤10百分位)", "距52週高點≤5%"],
    ["強勢突破盤", "近20日強於大盤≥10%", "地量(≤0.5倍均量)"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "距52週高點≤5%"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "外資近5日買超"],
    ["短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)", "N字底剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3日向上跳空缺口"],

]
BT_WIN_STATS = {
    "大盤跌破60日均線 ＋ 月營收連3月年增 ＋ 月營收創24個月新高": {"5": {"n": 10070, "win": 60.3, "ex": 0.8, "t": 12.15}, "10": {"n": 10070, "win": 60.6, "ex": 1.29, "t": 13.26}, "20": {"n": 10070, "win": 60.5, "ex": 2.15, "t": 14.99}},
    "近3日漲停 ＋ 月營收連3月年增 ＋ 月營收創24個月新高": {"5": {"n": 3713, "win": 53.4, "ex": 1.85, "t": 9.9}, "10": {"n": 3713, "win": 55.4, "ex": 3.23, "t": 12.16}, "20": {"n": 3713, "win": 57.9, "ex": 5.44, "t": 14.52}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 均線多頭排列(5>20>60且站上月線)": {"5": {"n": 752, "win": 57.3, "ex": 1.29, "t": 5.92}, "10": {"n": 752, "win": 57.8, "ex": 2.06, "t": 6.49}, "20": {"n": 752, "win": 58.0, "ex": 2.47, "t": 5.68}},
    "近3日向上跳空缺口 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"5": {"n": 2745, "win": 55.0, "ex": 2.02, "t": 9.15}, "10": {"n": 2745, "win": 57.2, "ex": 3.32, "t": 11.01}, "20": {"n": 2745, "win": 58.3, "ex": 5.23, "t": 12.06}},
    "近3日漲停 ＋ 月營收創24個月新高": {"5": {"n": 4781, "win": 53.4, "ex": 1.81, "t": 10.83}, "10": {"n": 4781, "win": 54.8, "ex": 3.0, "t": 12.72}, "20": {"n": 4781, "win": 56.7, "ex": 5.06, "t": 14.81}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 漲時量≥跌時量1.5倍(近20日)": {"5": {"n": 506, "win": 55.3, "ex": 1.31, "t": 4.96}, "10": {"n": 506, "win": 54.7, "ex": 2.15, "t": 5.45}, "20": {"n": 506, "win": 53.0, "ex": 1.56, "t": 3.05}},
    "近3日漲停 ＋ 三大法人近3月買超 ＋ 月營收創24個月新高": {"5": {"n": 3267, "win": 54.0, "ex": 2.01, "t": 9.62}, "10": {"n": 3267, "win": 54.7, "ex": 3.33, "t": 11.23}, "20": {"n": 3267, "win": 56.4, "ex": 5.68, "t": 13.31}},
    "強勢突破盤 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"5": {"n": 1862, "win": 54.2, "ex": 2.8, "t": 9.76}, "10": {"n": 1862, "win": 54.4, "ex": 4.16, "t": 10.51}, "20": {"n": 1862, "win": 55.4, "ex": 6.62, "t": 11.17}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 外資近5日買超": {"5": {"n": 667, "win": 55.8, "ex": 0.95, "t": 4.43}, "10": {"n": 667, "win": 57.4, "ex": 1.83, "t": 5.98}, "20": {"n": 667, "win": 57.6, "ex": 2.0, "t": 4.84}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 三大法人近3月買超": {"5": {"n": 909, "win": 54.8, "ex": 0.89, "t": 5.03}, "10": {"n": 909, "win": 55.8, "ex": 1.52, "t": 5.93}, "20": {"n": 909, "win": 55.7, "ex": 1.86, "t": 5.16}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"5": {"n": 1475, "win": 53.9, "ex": 2.68, "t": 8.13}, "10": {"n": 1475, "win": 56.2, "ex": 4.51, "t": 9.85}, "20": {"n": 1475, "win": 56.0, "ex": 7.61, "t": 10.97}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 近3月均價YoY為正": {"5": {"n": 1041, "win": 54.7, "ex": 0.93, "t": 5.69}, "10": {"n": 1041, "win": 54.8, "ex": 1.59, "t": 6.57}, "20": {"n": 1041, "win": 56.2, "ex": 1.9, "t": 5.71}},
    "強勢突破盤 ＋ 創52週新高 ＋ 月營收創24個月新高": {"5": {"n": 2077, "win": 51.4, "ex": 1.69, "t": 6.92}, "10": {"n": 2077, "win": 53.7, "ex": 2.9, "t": 8.59}, "20": {"n": 2077, "win": 55.1, "ex": 5.7, "t": 11.12}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%)": {"5": {"n": 311, "win": 55.6, "ex": 3.67, "t": 4.78}, "10": {"n": 311, "win": 60.1, "ex": 6.44, "t": 5.7}, "20": {"n": 311, "win": 63.3, "ex": 11.14, "t": 6.81}},
    "近3日向上跳空缺口 ＋ 月營收連3月年增 ＋ 月營收創24個月新高": {"5": {"n": 5985, "win": 53.2, "ex": 1.0, "t": 8.79}, "10": {"n": 5985, "win": 55.4, "ex": 1.98, "t": 12.42}, "20": {"n": 5985, "win": 56.0, "ex": 3.34, "t": 14.76}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 近3月乖離度為負(股價超前營收)": {"5": {"n": 681, "win": 55.9, "ex": 1.13, "t": 5.26}, "10": {"n": 681, "win": 55.1, "ex": 1.71, "t": 5.59}, "20": {"n": 681, "win": 56.5, "ex": 2.26, "t": 5.03}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"5": {"n": 217, "win": 57.6, "ex": 4.68, "t": 4.79}, "10": {"n": 217, "win": 59.9, "ex": 6.92, "t": 4.72}, "20": {"n": 217, "win": 59.4, "ex": 10.72, "t": 5.05}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 近3月均價YoY為負": {"5": {"n": 129, "win": 60.5, "ex": 4.71, "t": 4.45}, "10": {"n": 129, "win": 62.0, "ex": 6.35, "t": 4.59}, "20": {"n": 129, "win": 61.2, "ex": 8.37, "t": 4.12}},
    "大盤跌破60日均線 ＋ 月營收創24個月新高 ＋ 外資近5日買超": {"5": {"n": 6770, "win": 59.4, "ex": 0.65, "t": 8.39}, "10": {"n": 6770, "win": 59.7, "ex": 1.02, "t": 8.84}, "20": {"n": 6770, "win": 59.2, "ex": 1.86, "t": 10.65}},
    "多方力道≥80 ＋ 量能區間低檔(≤10百分位) ＋ 月營收連3月年增": {"5": {"n": 352, "win": 52.3, "ex": 1.34, "t": 3.34}, "10": {"n": 352, "win": 55.7, "ex": 3.28, "t": 5.2}, "20": {"n": 352, "win": 54.5, "ex": 3.5, "t": 3.92}},
    "回後買上漲全通過 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"5": {"n": 579, "win": 51.1, "ex": 2.35, "t": 4.38}, "10": {"n": 579, "win": 54.2, "ex": 3.03, "t": 4.21}, "20": {"n": 579, "win": 56.0, "ex": 6.48, "t": 5.77}},
    "多方力道≥65 ＋ 量能區間低檔(≤10百分位) ＋ 距52週高點≤5%": {"5": {"n": 463, "win": 55.9, "ex": 1.19, "t": 4.05}, "10": {"n": 463, "win": 57.0, "ex": 2.13, "t": 4.98}, "20": {"n": 463, "win": 58.5, "ex": 2.2, "t": 3.72}},
    "強勢突破盤 ＋ 近20日強於大盤≥10% ＋ 地量(≤0.5倍均量)": {"5": {"n": 402, "win": 54.5, "ex": 2.86, "t": 4.59}, "10": {"n": 402, "win": 57.2, "ex": 4.93, "t": 5.43}, "20": {"n": 402, "win": 57.7, "ex": 8.25, "t": 6.18}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 距52週高點≤5%": {"5": {"n": 193, "win": 61.7, "ex": 2.71, "t": 4.33}, "10": {"n": 193, "win": 61.7, "ex": 5.04, "t": 5.69}, "20": {"n": 193, "win": 65.3, "ex": 5.54, "t": 4.37}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"5": {"n": 368, "win": 57.1, "ex": 2.92, "t": 4.96}, "10": {"n": 368, "win": 54.3, "ex": 3.61, "t": 4.26}, "20": {"n": 368, "win": 56.8, "ex": 6.39, "t": 5.2}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 外資近5日買超": {"5": {"n": 1682, "win": 46.8, "ex": 0.31, "t": 2.4}, "10": {"n": 1682, "win": 50.3, "ex": 0.86, "t": 4.42}, "20": {"n": 1682, "win": 55.5, "ex": 1.45, "t": 4.76}},
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價) ＋ N字底剛形成": {"5": {"n": 138, "win": 55.1, "ex": 4.87, "t": 3.8}, "10": {"n": 138, "win": 51.4, "ex": 5.75, "t": 3.2}, "20": {"n": 138, "win": 52.2, "ex": 7.55, "t": 3.06}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": {"5": {"n": 291, "win": 55.7, "ex": 3.63, "t": 4.76}, "10": {"n": 291, "win": 58.4, "ex": 6.09, "t": 5.51}, "20": {"n": 291, "win": 59.8, "ex": 9.76, "t": 6.07}},

}
BT_HOT_COMBOS = [
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "三大法人近3月賣超"],
    ["距52週高點≤5%", "近3日漲停", "月營收創24個月新高"],
    ["短線過熱(高於月線≥20%)", "月營收連3月年增", "N字底剛形成"],
    ["近20日強於大盤≥10%", "地量(≤0.5倍均量)", "近3日漲停"],
    ["近3日漲停", "突破ABC修正下降切線剛形成", "K線橫盤的突破剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3日漲停"],
    ["地量(≤0.5倍均量)", "OBV能量潮創60日新高", "近3日漲停"],
    ["地量(≤0.5倍均量)", "創52週新高", "三大法人近3月賣超"],
    ["短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)", "N字底剛形成"],
    ["近3日漲停", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["地量(≤0.5倍均量)", "CMF資金流買方佔優(近20日≥0.1)", "近3日漲停"],
    ["相對強弱為正(強於大盤)", "地量(≤0.5倍均量)", "近3日漲停"],
    ["多方力道≥65", "地量(≤0.5倍均量)", "近3日漲停"],
    ["多方力道≥80", "近3日漲停", "K線橫盤的突破剛形成"],
    ["OBV能量潮創60日新高", "近3日漲停", "N字底剛形成"],
    ["大盤站上20日均線", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["地量(≤0.5倍均量)", "近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)"],
    ["地量(≤0.5倍均量)", "創52週新高", "近3日漲停"],
    ["地量(≤0.5倍均量)", "創52週新高", "月營收創24個月新高"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["CMF資金流買方佔優(近20日≥0.1)", "近3日漲停", "N字底剛形成"],
    ["創52週新高", "近3日漲停", "K線橫盤的突破剛形成"],
    ["近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "K線橫盤的突破剛形成"],
    ["量能區間高檔(≥90百分位)", "近3日漲停", "N字底剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "近3日漲停"],
    ["距52週高點≤5%", "近3日漲停", "K線橫盤的突破剛形成"],
    ["創52週新高", "近3日漲停", "N字底剛形成"],
    ["近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "N字底剛形成"],
    ["相對強弱為負(弱於大盤)", "近3日漲停", "突破ABC修正下降切線剛形成"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["多方力道≥80", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "近3日漲停", "N字底剛形成"],
    ["大盤站上20日均線", "近3日漲停", "N字底剛形成"],
    ["漲時量≥跌時量1.5倍(近20日)", "近3日漲停", "N字底剛形成"],
    ["均線多頭排列(5>20>60且站上月線)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "近3日漲停", "N字底剛形成"],
    ["均線多頭排列(5>20>60且站上月線)", "近3日漲停", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "N字底剛形成"],
    ["近3日漲停", "月營收創24個月新高", "圓弧底剛形成"],
    ["地量(≤0.5倍均量)", "OBV能量潮創60日新高", "短線過熱(高於月線≥20%)"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3日向上跳空缺口"],
    ["多方力道≥80", "近3日漲停", "突破ABC修正下降切線剛形成"],
    ["距52週高點≤5%", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["多方力道≥80", "近3日漲停", "N字底剛形成"],
    ["多方力道≥65", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["漲時量≥跌時量1.5倍(近20日)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "N字底剛形成"],
    ["多方力道≥65", "近3日漲停", "N字底剛形成"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "近3日漲停"],
    ["近3日漲停", "外資近5日買超", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "近3日漲停", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "CMF資金流買方佔優(近20日≥0.1)", "晨星剛形成"],
    ["強勢突破盤", "近3日漲停", "N字底剛形成"],
    ["大盤跌破20日均線", "月營收創24個月新高", "夜星剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)"],
    ["地量(≤0.5倍均量)", "近3日向上跳空缺口", "短線過熱(高於月線≥20%)"],
    ["地量(≤0.5倍均量)", "均線多頭排列(5>20>60且站上月線)", "近3日漲停"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["布林通道高檔(≥80%)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["大盤跌破20日均線", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["近20日強於大盤≥10%", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["近20日強於大盤≥10%", "近3日漲停", "N字底剛形成"],
    ["地量(≤0.5倍均量)", "近3日漲停", "三大法人近3月買超"],
    ["近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "突破ABC修正下降切線剛形成"],
    ["短線過熱(高於月線≥20%)", "三大法人近3月買超", "N字底剛形成"],
    ["地量(≤0.5倍均量)", "漲時量≥跌時量1.5倍(近20日)", "近3日漲停"],
    ["近3日漲停", "月營收連3月年增", "突破ABC修正下降切線剛形成"],
    ["布林通道高檔(≥80%)", "近3日漲停", "N字底剛形成"],
    ["距52週高點≤5%", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "月營收創24個月新高"],
    ["創52週新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "母子懷抱(低檔)剛形成"],
    ["距52週高點≤5%", "近3日漲停", "N字底剛形成"],
    ["量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["回後買上漲全通過", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["地量(≤0.5倍均量)", "距52週高點≤5%", "近3日漲停"],
    ["大盤站上60日均線", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["大盤跌破60日均線", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "近3日漲停", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "晨星剛形成"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "創52週新高"],
    ["量能區間低檔(≤10百分位)", "漲時量≥跌時量1.5倍(近20日)", "短線過熱(高於月線≥20%)"],
    ["距52週高點≤5%", "月營收創24個月新高", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "近3日漲停", "N字底剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "月營收創24個月新高"],
    ["均線多頭排列(5>20>60且站上月線)", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停"],
    ["近3日漲停", "月營收連3月年增", "N字底剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "K線橫盤的突破剛形成"],
    ["爆量(≥2倍均量)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "距52週高點≤5%", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "CMF資金流買方佔優(近20日≥0.1)", "晨星剛形成"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["多方力道≥65", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "OBV能量潮創60日新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["KDJ近3日內死亡交叉", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["近3日向上跳空缺口", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "晨星剛形成"],
    ["創52週新高", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["MACD近3日內黃金交叉", "爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)"],
    ["地量(≤0.5倍均量)", "漲時量≥跌時量1.5倍(近20日)", "夜星剛形成"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "短線過熱(高於月線≥20%)"],
    ["多方力道≥65", "量能區間低檔(≤10百分位)", "創52週新高"],
    ["量能區間低檔(≤10百分位)", "大盤站上60日均線", "短線過熱(高於月線≥20%)"],
    ["量能區間低檔(≤10百分位)", "CMF資金流買方佔優(近20日≥0.1)", "短線過熱(高於月線≥20%)"],
    ["量能區間低檔(≤10百分位)", "均線多頭排列(5>20>60且站上月線)", "短線過熱(高於月線≥20%)"],
    ["量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)", "近3月均價YoY為正"],
    ["量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)", "三大法人近3月買超"],
    ["地量(≤0.5倍均量)", "量能區間高檔(≥90百分位)", "近3日漲停"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "距52週高點≤5%"],
    ["近3月均價YoY為正", "月營收創24個月新高", "夜星剛形成"],
    ["OBV能量潮創60日新高", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["MACD近3日內黃金交叉", "連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)"],
    ["月營收連3月年增", "月營收創24個月新高", "夜星剛形成"],
    ["短線過熱(高於月線≥20%)", "近3月乖離度為負(股價超前營收)", "月營收創24個月新高"],
    ["大盤跌破60日均線", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "晨星剛形成"],
    ["近20日強於大盤≥10%", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["創52週新高", "近3月均價YoY為負", "突破ABC修正下降切線剛形成"],
    ["KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "近3日漲停", "月營收連3月年增"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["多方力道≥65", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["MACD近3日內黃金交叉", "大盤跌破60日均線", "短線過熱(高於月線≥20%)"],
    ["近3日漲停", "月營收創24個月新高", "N字底剛形成"],
    ["均線多頭排列(5>20>60且站上月線)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["多方力道≥80", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "漲時量≥跌時量1.5倍(近20日)", "晨星剛形成"],
    ["爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "短線過熱(高於月線≥20%)", "近3月乖離度為負(股價超前營收)"],
    ["地量(≤0.5倍均量)", "大盤站上20日均線", "近3日漲停"],
    ["布林通道高檔(≥80%)", "月營收創24個月新高", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)"],
    ["短線過熱(高於月線≥20%)", "外資近5日買超", "N字底剛形成"],
    ["MACD近3日內黃金交叉", "OBV能量潮創60日新高", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "爆量(≥2倍均量)", "短線過熱(高於月線≥20%)"],
    ["大盤站上20日均線", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["短線過熱(高於月線≥20%)", "近3月均價YoY為負", "N字底剛形成"],
    ["創52週新高", "近3日漲停", "月營收創24個月新高"],
    ["CMF資金流買方佔優(近20日≥0.1)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "月營收連3月年增"],
    ["近20日強於大盤≥10%", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["多方力道≥65", "MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["布林通道高檔(≥80%)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["短線過熱(高於月線≥20%)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "突破ABC修正下降切線剛形成"],
    ["地量(≤0.5倍均量)", "量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)"],
    ["多方力道≥80", "短線過熱(高於月線≥20%)", "突破ABC修正下降切線剛形成"],
    ["月營收創24個月新高", "夜星剛形成"],
    ["短線過熱(高於月線≥20%)", "三大法人近3月買超", "月營收創24個月新高"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["近20日強於大盤≥10%", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["距52週高點≤5%", "近3日漲停", "近3月均價YoY為負"],
    ["近20日強於大盤≥10%", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)", "近3月乖離度為負(股價超前營收)"],
    ["布林通道高檔(≥80%)", "量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)"],
    ["大盤跌破20日均線", "月營收連3月年增", "夜星剛形成"],
    ["近20日強於大盤≥10%", "地量(≤0.5倍均量)", "夜星剛形成"],
    ["多方力道≥65", "量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "月營收創24個月新高", "晨星剛形成"],
    ["量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)"],
    ["均線多頭排列(5>20>60且站上月線)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "均線多頭排列(5>20>60且站上月線)", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能區間高檔(≥90百分位)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停"],
    ["大盤站上60日均線", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["創52週新高", "三大法人近3月買超", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["距52週高點≤5%", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["三大法人近3月買超", "月營收創24個月新高", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)"],
    ["強勢突破盤", "月營收連3月年增", "母子懷抱(低檔)剛形成"],
    ["創52週新高", "近3日漲停", "三大法人近3月賣超"],
    ["MACD近3日內黃金交叉", "KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["創52週新高", "近3日向上跳空缺口", "三大法人近3月賣超"],
    ["創52週新高", "近3日漲停", "近3月均價YoY為負"],
    ["漲時量≥跌時量1.5倍(近20日)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日向上跳空缺口", "近3日漲停"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "均線多頭排列(5>20>60且站上月線)", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "三大法人近3月買超", "突破ABC修正下降切線剛形成"],
    ["相對強弱為負(弱於大盤)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "近3日漲停"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "突破飆股大量黑K最高點剛形成"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "近3日漲停"],
    ["大盤站上60日均線", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["相對強弱為正(強於大盤)", "月營收創24個月新高", "夜星剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "創52週新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "距52週高點≤5%", "近3日漲停"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "三大法人近3月買超"],
    ["布林通道高檔(≥80%)", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "漲時量≥跌時量1.5倍(近20日)", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "近3日向上跳空缺口", "短線過熱(高於月線≥20%)"],
    ["回後買上漲全通過", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["爆量(≥2倍均量)", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["近3日向上跳空缺口", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "外資近5日買超"],
    ["大盤跌破60日均線", "外資近5日買超", "夜星剛形成"],
    ["KDJ近3日內死亡交叉", "距52週高點≤5%", "短線過熱(高於月線≥20%)"],
    ["OBV能量潮創60日新高", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近20日強於大盤≥10%", "創52週新高", "突破ABC修正下降切線剛形成"],
    ["爆量(≥2倍均量)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近3日漲停", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "外資近5日買超"],
    ["創52週新高", "近3日漲停", "突破ABC修正下降切線剛形成"],
    ["MACD近3日內黃金交叉", "大盤站上20日均線", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "相對強弱為正(強於大盤)", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "近20日強於大盤≥10%", "短線過熱(高於月線≥20%)"],
    ["創52週新高", "近3日漲停", "近3月乖離度為正(營收優於股價)"],
    ["KDJ近3日內死亡交叉", "爆量(≥2倍均量)", "短線過熱(高於月線≥20%)"],
    ["多方力道≥80", "創52週新高", "突破ABC修正下降切線剛形成"],
    ["近20日強於大盤≥10%", "近3日向上跳空缺口", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "三大法人近3月賣超"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["布林通道高檔(≥80%)", "MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["強勢突破盤", "月營收連3月年增", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "創52週新高", "月營收創24個月新高"],
    ["距52週高點≤5%", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["爆量(≥2倍均量)", "月營收連3月年增", "母子懷抱(低檔)剛形成"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "三大法人近3月賣超"],
    ["CMF資金流買方佔優(近20日≥0.1)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近3日漲停", "N字底剛形成", "K線橫盤的突破剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "三大法人近3月買超"],
    ["MACD近3日內黃金交叉", "CMF資金流買方佔優(近20日≥0.1)", "短線過熱(高於月線≥20%)"],
    ["創52週新高", "近3日向上跳空缺口", "近3月均價YoY為負"],
    ["量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["地量(≤0.5倍均量)", "創52週新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "近3月乖離度為負(股價超前營收)"],
    ["地量(≤0.5倍均量)", "近3日向上跳空缺口", "近3日漲停"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "近3月均價YoY為正"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破20日均線", "近3日漲停"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["強勢突破盤", "MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["布林通道高檔(≥80%)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近3日漲停", "近3月乖離度為正(營收優於股價)", "K線橫盤的突破剛形成"],
    ["近20日強於大盤≥10%", "地量(≤0.5倍均量)", "晨星剛形成"],
    ["距52週高點≤5%", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["近20日強於大盤≥10%", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["短線過熱(高於月線≥20%)", "近3月均價YoY為負", "月營收創24個月新高"],
    ["均線多頭排列(5>20>60且站上月線)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["相對強弱為正(強於大盤)", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["創52週新高", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)"],
    ["布林通道高檔(≥80%)", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["距52週高點≤5%", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "月營收創24個月新高", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "近3月乖離度為負(股價超前營收)"],
    ["KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["爆量(≥2倍均量)", "創52週新高", "月營收創24個月新高"],
    ["近3日漲停", "三大法人近3月買超", "突破ABC修正下降切線剛形成"],
    ["創52週新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "突破ABC修正下降切線剛形成"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["強勢突破盤", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["大盤跌破60日均線", "近3日漲停", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "近3日向上跳空缺口", "OBV能量潮創60日新高"],
    ["距52週高點≤5%", "短線過熱(高於月線≥20%)", "突破ABC修正下降切線剛形成"],
    ["近3日向上跳空缺口", "近3日漲停", "月營收創24個月新高"],
    ["近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "距52週高點≤5%", "短線過熱(高於月線≥20%)"],
    ["OBV能量潮創60日新高", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["近20日強於大盤≥10%", "三大法人近3月買超", "夜星剛形成"],
    ["OBV能量潮創60日新高", "近3日漲停", "月營收創24個月新高"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "距52週高點≤5%", "晨星剛形成"],
    ["均線多頭排列(5>20>60且站上月線)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["距52週高點≤5%", "近3日漲停", "近3月乖離度為正(營收優於股價)"],
    ["爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["KDJ近3日內黃金交叉", "量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)"],
    ["近3日漲停", "近3月乖離度為正(營收優於股價)", "月營收創24個月新高"],
    ["回後買上漲全通過", "創52週新高", "三大法人近3月賣超"],
    ["爆量(≥2倍均量)", "大盤跌破60日均線", "短線過熱(高於月線≥20%)"],
    ["KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "OBV能量潮創60日新高"],
    ["多方力道≥80", "CMF資金流買方佔優(近20日≥0.1)", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "近3月均價YoY為負", "外資近5日買超"],
    ["大盤站上20日均線", "距52週高點≤5%", "短線過熱(高於月線≥20%)"],
    ["近20日強於大盤≥10%", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["MACD近3日內黃金交叉", "量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)"],
    ["強勢突破盤", "月營收創24個月新高", "突破飆股大量黑K最高點剛形成"],
    ["KDJ近3日內黃金交叉", "近3日漲停", "突破ABC修正下降切線剛形成"],
    ["月營收創24個月新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "連續放量(近3日均量≥1.5倍前20日均量)", "OBV能量潮創60日新高"],

]
BT_HOT_STATS = {
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月賣超": {"10": {"n": 167, "moonshotN": 20, "pct": 12.0, "avg": 50.4}, "20": {"n": 167, "moonshotN": 21, "pct": 12.6, "avg": 65.7}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 1680, "moonshotN": 170, "pct": 10.1, "avg": 46.2}, "20": {"n": 1680, "moonshotN": 294, "pct": 17.5, "avg": 56.1}},
    "短線過熱(高於月線≥20%) ＋ 月營收連3月年增 ＋ N字底剛形成": {"10": {"n": 157, "moonshotN": 22, "pct": 14.0, "avg": 43.4}, "20": {"n": 157, "moonshotN": 28, "pct": 17.8, "avg": 57.9}},
    "近20日強於大盤≥10% ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 796, "moonshotN": 95, "pct": 11.9, "avg": 46.9}, "20": {"n": 796, "moonshotN": 159, "pct": 20.0, "avg": 55.2}},
    "近3日漲停 ＋ 突破ABC修正下降切線剛形成 ＋ K線橫盤的突破剛形成": {"10": {"n": 220, "moonshotN": 24, "pct": 10.9, "avg": 47.3}, "20": {"n": 220, "moonshotN": 24, "pct": 10.9, "avg": 57.8}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 217, "moonshotN": 33, "pct": 15.2, "avg": 47.5}, "20": {"n": 217, "moonshotN": 52, "pct": 24.0, "avg": 59.2}},
    "地量(≤0.5倍均量) ＋ OBV能量潮創60日新高 ＋ 近3日漲停": {"10": {"n": 172, "moonshotN": 22, "pct": 12.8, "avg": 44.9}, "20": {"n": 172, "moonshotN": 42, "pct": 24.4, "avg": 50.7}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 三大法人近3月賣超": {"10": {"n": 61, "moonshotN": 9, "pct": 14.8, "avg": 49.5}, "20": {"n": 61, "moonshotN": 17, "pct": 27.9, "avg": 49.9}},
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價) ＋ N字底剛形成": {"10": {"n": 138, "moonshotN": 16, "pct": 11.6, "avg": 51.7}, "20": {"n": 138, "moonshotN": 27, "pct": 19.6, "avg": 56.5}},
    "近3日漲停 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 251, "moonshotN": 31, "pct": 12.4, "avg": 48.5}, "20": {"n": 251, "moonshotN": 48, "pct": 19.1, "avg": 55.0}},
    "地量(≤0.5倍均量) ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 近3日漲停": {"10": {"n": 603, "moonshotN": 74, "pct": 12.3, "avg": 47.3}, "20": {"n": 603, "moonshotN": 126, "pct": 20.9, "avg": 58.0}},
    "相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 973, "moonshotN": 108, "pct": 11.1, "avg": 46.6}, "20": {"n": 973, "moonshotN": 174, "pct": 17.9, "avg": 55.4}},
    "多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 685, "moonshotN": 80, "pct": 11.7, "avg": 47.0}, "20": {"n": 685, "moonshotN": 135, "pct": 19.7, "avg": 57.7}},
    "多方力道≥80 ＋ 近3日漲停 ＋ K線橫盤的突破剛形成": {"10": {"n": 431, "moonshotN": 49, "pct": 11.4, "avg": 46.4}, "20": {"n": 431, "moonshotN": 56, "pct": 13.0, "avg": 56.4}},
    "OBV能量潮創60日新高 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 297, "moonshotN": 36, "pct": 12.1, "avg": 47.0}, "20": {"n": 297, "moonshotN": 49, "pct": 16.5, "avg": 52.9}},
    "大盤站上20日均線 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 299, "moonshotN": 36, "pct": 12.0, "avg": 45.7}, "20": {"n": 299, "moonshotN": 53, "pct": 17.7, "avg": 53.5}},
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)": {"10": {"n": 581, "moonshotN": 70, "pct": 12.0, "avg": 47.5}, "20": {"n": 581, "moonshotN": 129, "pct": 22.2, "avg": 57.6}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 近3日漲停": {"10": {"n": 285, "moonshotN": 33, "pct": 11.6, "avg": 46.3}, "20": {"n": 285, "moonshotN": 57, "pct": 20.0, "avg": 56.1}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 月營收創24個月新高": {"10": {"n": 244, "moonshotN": 27, "pct": 11.1, "avg": 47.7}, "20": {"n": 244, "moonshotN": 54, "pct": 22.1, "avg": 56.1}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 249, "moonshotN": 30, "pct": 12.0, "avg": 44.0}, "20": {"n": 249, "moonshotN": 47, "pct": 18.9, "avg": 53.7}},
    "CMF資金流買方佔優(近20日≥0.1) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 341, "moonshotN": 35, "pct": 10.3, "avg": 43.3}, "20": {"n": 341, "moonshotN": 50, "pct": 14.7, "avg": 50.5}},
    "創52週新高 ＋ 近3日漲停 ＋ K線橫盤的突破剛形成": {"10": {"n": 328, "moonshotN": 33, "pct": 10.1, "avg": 48.3}, "20": {"n": 328, "moonshotN": 39, "pct": 11.9, "avg": 56.0}},
    "近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ K線橫盤的突破剛形成": {"10": {"n": 523, "moonshotN": 56, "pct": 10.7, "avg": 47.1}, "20": {"n": 523, "moonshotN": 65, "pct": 12.4, "avg": 55.3}},
    "量能區間高檔(≥90百分位) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 396, "moonshotN": 43, "pct": 10.9, "avg": 45.3}, "20": {"n": 396, "moonshotN": 61, "pct": 15.4, "avg": 51.4}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 513, "moonshotN": 55, "pct": 10.7, "avg": 47.0}, "20": {"n": 513, "moonshotN": 100, "pct": 19.5, "avg": 56.3}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ K線橫盤的突破剛形成": {"10": {"n": 515, "moonshotN": 52, "pct": 10.1, "avg": 46.5}, "20": {"n": 515, "moonshotN": 64, "pct": 12.4, "avg": 55.4}},
    "創52週新高 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 281, "moonshotN": 35, "pct": 12.5, "avg": 46.8}, "20": {"n": 281, "moonshotN": 54, "pct": 19.2, "avg": 53.5}},
    "近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ N字底剛形成": {"10": {"n": 358, "moonshotN": 42, "pct": 11.7, "avg": 45.3}, "20": {"n": 358, "moonshotN": 54, "pct": 15.1, "avg": 51.7}},
    "相對強弱為負(弱於大盤) ＋ 近3日漲停 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 135, "moonshotN": 18, "pct": 13.3, "avg": 44.0}, "20": {"n": 135, "moonshotN": 16, "pct": 11.9, "avg": 60.2}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 490, "moonshotN": 58, "pct": 11.8, "avg": 48.2}, "20": {"n": 490, "moonshotN": 98, "pct": 20.0, "avg": 58.1}},
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 176, "moonshotN": 20, "pct": 11.4, "avg": 44.7}, "20": {"n": 176, "moonshotN": 28, "pct": 15.9, "avg": 57.8}},
    "多方力道≥80 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 289, "moonshotN": 30, "pct": 10.4, "avg": 45.1}, "20": {"n": 289, "moonshotN": 49, "pct": 17.0, "avg": 52.7}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 298, "moonshotN": 31, "pct": 10.4, "avg": 46.6}, "20": {"n": 298, "moonshotN": 39, "pct": 13.1, "avg": 53.7}},
    "大盤站上20日均線 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 463, "moonshotN": 51, "pct": 11.0, "avg": 45.9}, "20": {"n": 463, "moonshotN": 68, "pct": 14.7, "avg": 52.3}},
    "漲時量≥跌時量1.5倍(近20日) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 421, "moonshotN": 43, "pct": 10.2, "avg": 45.6}, "20": {"n": 421, "moonshotN": 59, "pct": 14.0, "avg": 51.3}},
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 314, "moonshotN": 35, "pct": 11.1, "avg": 46.2}, "20": {"n": 314, "moonshotN": 55, "pct": 17.5, "avg": 53.4}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 209, "moonshotN": 29, "pct": 13.9, "avg": 46.5}, "20": {"n": 209, "moonshotN": 34, "pct": 16.3, "avg": 57.1}},
    "短線過熱(高於月線≥20%) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 237, "moonshotN": 29, "pct": 12.2, "avg": 47.2}, "20": {"n": 237, "moonshotN": 41, "pct": 17.3, "avg": 52.3}},
    "均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 408, "moonshotN": 47, "pct": 11.5, "avg": 47.2}, "20": {"n": 408, "moonshotN": 63, "pct": 15.4, "avg": 52.9}},
    "短線過熱(高於月線≥20%) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ N字底剛形成": {"10": {"n": 304, "moonshotN": 32, "pct": 10.5, "avg": 45.8}, "20": {"n": 304, "moonshotN": 51, "pct": 16.8, "avg": 53.2}},
    "近3日漲停 ＋ 月營收創24個月新高 ＋ 圓弧底剛形成": {"10": {"n": 128, "moonshotN": 13, "pct": 10.2, "avg": 50.4}, "20": {"n": 128, "moonshotN": 17, "pct": 13.3, "avg": 49.5}},
    "地量(≤0.5倍均量) ＋ OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 335, "moonshotN": 34, "pct": 10.1, "avg": 44.2}, "20": {"n": 335, "moonshotN": 69, "pct": 20.6, "avg": 49.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 413, "moonshotN": 44, "pct": 10.7, "avg": 45.0}, "20": {"n": 413, "moonshotN": 72, "pct": 17.4, "avg": 59.1}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": {"10": {"n": 291, "moonshotN": 36, "pct": 12.4, "avg": 46.0}, "20": {"n": 291, "moonshotN": 58, "pct": 19.9, "avg": 57.2}},
    "多方力道≥80 ＋ 近3日漲停 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 353, "moonshotN": 37, "pct": 10.5, "avg": 45.3}, "20": {"n": 353, "moonshotN": 55, "pct": 15.6, "avg": 58.2}},
    "距52週高點≤5% ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 272, "moonshotN": 33, "pct": 12.1, "avg": 45.3}, "20": {"n": 272, "moonshotN": 51, "pct": 18.8, "avg": 53.9}},
    "多方力道≥80 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 315, "moonshotN": 37, "pct": 11.7, "avg": 44.5}, "20": {"n": 315, "moonshotN": 51, "pct": 16.2, "avg": 50.8}},
    "多方力道≥65 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 321, "moonshotN": 34, "pct": 10.6, "avg": 46.0}, "20": {"n": 321, "moonshotN": 55, "pct": 17.1, "avg": 53.2}},
    "漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 306, "moonshotN": 31, "pct": 10.1, "avg": 44.7}, "20": {"n": 306, "moonshotN": 52, "pct": 17.0, "avg": 52.4}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ N字底剛形成": {"10": {"n": 78, "moonshotN": 9, "pct": 11.5, "avg": 39.4}, "20": {"n": 78, "moonshotN": 14, "pct": 17.9, "avg": 50.7}},
    "多方力道≥65 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 418, "moonshotN": 49, "pct": 11.7, "avg": 45.6}, "20": {"n": 418, "moonshotN": 62, "pct": 14.8, "avg": 52.1}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 192, "moonshotN": 22, "pct": 11.5, "avg": 48.0}, "20": {"n": 192, "moonshotN": 35, "pct": 18.2, "avg": 55.1}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 145, "moonshotN": 14, "pct": 9.7, "avg": 44.1}, "20": {"n": 145, "moonshotN": 27, "pct": 18.6, "avg": 51.1}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 528, "moonshotN": 61, "pct": 11.6, "avg": 46.0}, "20": {"n": 528, "moonshotN": 116, "pct": 22.0, "avg": 56.4}},
    "近3日漲停 ＋ 外資近5日買超 ＋ N字底剛形成": {"10": {"n": 373, "moonshotN": 39, "pct": 10.5, "avg": 47.3}, "20": {"n": 373, "moonshotN": 55, "pct": 14.7, "avg": 53.9}},
    "短線過熱(高於月線≥20%) ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 1575, "moonshotN": 167, "pct": 10.6, "avg": 47.8}, "20": {"n": 1575, "moonshotN": 291, "pct": 18.5, "avg": 57.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 晨星剛形成": {"10": {"n": 116, "moonshotN": 10, "pct": 8.6, "avg": 45.0}, "20": {"n": 116, "moonshotN": 24, "pct": 20.7, "avg": 48.9}},
    "強勢突破盤 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 302, "moonshotN": 38, "pct": 12.6, "avg": 46.8}, "20": {"n": 302, "moonshotN": 38, "pct": 12.6, "avg": 55.4}},
    "大盤跌破20日均線 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 64, "moonshotN": 4, "pct": 6.2, "avg": 44.1}, "20": {"n": 64, "moonshotN": 12, "pct": 18.8, "avg": 47.0}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 311, "moonshotN": 38, "pct": 12.2, "avg": 46.8}, "20": {"n": 311, "moonshotN": 73, "pct": 23.5, "avg": 55.3}},
    "地量(≤0.5倍均量) ＋ 近3日向上跳空缺口 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 553, "moonshotN": 64, "pct": 11.6, "avg": 45.0}, "20": {"n": 553, "moonshotN": 114, "pct": 20.6, "avg": 57.3}},
    "地量(≤0.5倍均量) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停": {"10": {"n": 748, "moonshotN": 85, "pct": 11.4, "avg": 46.7}, "20": {"n": 748, "moonshotN": 145, "pct": 19.4, "avg": 56.2}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 143, "moonshotN": 16, "pct": 11.2, "avg": 47.0}, "20": {"n": 143, "moonshotN": 22, "pct": 15.4, "avg": 58.9}},
    "短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 342, "moonshotN": 38, "pct": 11.1, "avg": 45.5}, "20": {"n": 342, "moonshotN": 57, "pct": 16.7, "avg": 53.3}},
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 341, "moonshotN": 38, "pct": 11.1, "avg": 45.5}, "20": {"n": 341, "moonshotN": 57, "pct": 16.7, "avg": 53.3}},
    "大盤跌破20日均線 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 287, "moonshotN": 31, "pct": 10.8, "avg": 45.6}, "20": {"n": 287, "moonshotN": 39, "pct": 13.6, "avg": 61.0}},
    "近20日強於大盤≥10% ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 335, "moonshotN": 35, "pct": 10.4, "avg": 45.0}, "20": {"n": 335, "moonshotN": 55, "pct": 16.4, "avg": 52.5}},
    "近20日強於大盤≥10% ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 423, "moonshotN": 44, "pct": 10.4, "avg": 47.0}, "20": {"n": 423, "moonshotN": 59, "pct": 13.9, "avg": 52.4}},
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 三大法人近3月買超": {"10": {"n": 1083, "moonshotN": 112, "pct": 10.3, "avg": 46.6}, "20": {"n": 1083, "moonshotN": 189, "pct": 17.5, "avg": 54.7}},
    "近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 425, "moonshotN": 43, "pct": 10.1, "avg": 45.6}, "20": {"n": 425, "moonshotN": 64, "pct": 15.1, "avg": 58.0}},
    "短線過熱(高於月線≥20%) ＋ 三大法人近3月買超 ＋ N字底剛形成": {"10": {"n": 270, "moonshotN": 31, "pct": 11.5, "avg": 45.4}, "20": {"n": 270, "moonshotN": 49, "pct": 18.1, "avg": 54.8}},
    "地量(≤0.5倍均量) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 近3日漲停": {"10": {"n": 690, "moonshotN": 78, "pct": 11.3, "avg": 45.6}, "20": {"n": 690, "moonshotN": 132, "pct": 19.1, "avg": 55.8}},
    "近3日漲停 ＋ 月營收連3月年增 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 503, "moonshotN": 55, "pct": 10.9, "avg": 44.7}, "20": {"n": 503, "moonshotN": 71, "pct": 14.1, "avg": 53.5}},
    "布林通道高檔(≥80%) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 521, "moonshotN": 54, "pct": 10.4, "avg": 46.3}, "20": {"n": 521, "moonshotN": 72, "pct": 13.8, "avg": 52.0}},
    "距52週高點≤5% ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1669, "moonshotN": 149, "pct": 8.9, "avg": 45.2}, "20": {"n": 1669, "moonshotN": 287, "pct": 17.2, "avg": 55.1}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 月營收創24個月新高": {"10": {"n": 203, "moonshotN": 25, "pct": 12.3, "avg": 49.9}, "20": {"n": 203, "moonshotN": 49, "pct": 24.1, "avg": 59.3}},
    "創52週新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 65, "moonshotN": 6, "pct": 9.2, "avg": 49.6}, "20": {"n": 65, "moonshotN": 10, "pct": 15.4, "avg": 66.8}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 320, "moonshotN": 39, "pct": 12.2, "avg": 46.7}, "20": {"n": 320, "moonshotN": 55, "pct": 17.2, "avg": 54.0}},
    "量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 232, "moonshotN": 25, "pct": 10.8, "avg": 45.8}, "20": {"n": 232, "moonshotN": 41, "pct": 17.7, "avg": 52.7}},
    "回後買上漲全通過 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 262, "moonshotN": 28, "pct": 10.7, "avg": 44.3}, "20": {"n": 262, "moonshotN": 44, "pct": 16.8, "avg": 52.5}},
    "地量(≤0.5倍均量) ＋ 距52週高點≤5% ＋ 近3日漲停": {"10": {"n": 355, "moonshotN": 40, "pct": 11.3, "avg": 45.3}, "20": {"n": 355, "moonshotN": 73, "pct": 20.6, "avg": 54.9}},
    "大盤站上60日均線 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 304, "moonshotN": 36, "pct": 11.8, "avg": 45.2}, "20": {"n": 304, "moonshotN": 55, "pct": 18.1, "avg": 53.1}},
    "大盤跌破60日均線 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 224, "moonshotN": 15, "pct": 6.7, "avg": 37.3}, "20": {"n": 224, "moonshotN": 36, "pct": 16.1, "avg": 51.4}},
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 433, "moonshotN": 58, "pct": 13.4, "avg": 48.7}, "20": {"n": 433, "moonshotN": 100, "pct": 23.1, "avg": 56.6}},
    "地量(≤0.5倍均量) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 晨星剛形成": {"10": {"n": 66, "moonshotN": 6, "pct": 9.1, "avg": 37.7}, "20": {"n": 66, "moonshotN": 13, "pct": 19.7, "avg": 60.1}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 創52週新高": {"10": {"n": 59, "moonshotN": 3, "pct": 5.1, "avg": 30.5}, "20": {"n": 59, "moonshotN": 11, "pct": 18.6, "avg": 44.4}},
    "量能區間低檔(≤10百分位) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 60, "moonshotN": 3, "pct": 5.0, "avg": 37.5}, "20": {"n": 60, "moonshotN": 11, "pct": 18.3, "avg": 42.3}},
    "距52週高點≤5% ＋ 月營收創24個月新高 ＋ 晨星剛形成": {"10": {"n": 101, "moonshotN": 8, "pct": 7.9, "avg": 39.4}, "20": {"n": 101, "moonshotN": 18, "pct": 17.8, "avg": 45.6}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 177, "moonshotN": 18, "pct": 10.2, "avg": 46.4}, "20": {"n": 177, "moonshotN": 22, "pct": 12.4, "avg": 46.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 425, "moonshotN": 44, "pct": 10.4, "avg": 48.1}, "20": {"n": 425, "moonshotN": 77, "pct": 18.1, "avg": 59.4}},
    "均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 135, "moonshotN": 13, "pct": 9.6, "avg": 45.9}, "20": {"n": 135, "moonshotN": 19, "pct": 14.1, "avg": 61.2}},
    "強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停": {"10": {"n": 363, "moonshotN": 37, "pct": 10.2, "avg": 46.0}, "20": {"n": 363, "moonshotN": 55, "pct": 15.2, "avg": 60.4}},
    "近3日漲停 ＋ 月營收連3月年增 ＋ N字底剛形成": {"10": {"n": 244, "moonshotN": 25, "pct": 10.2, "avg": 44.8}, "20": {"n": 244, "moonshotN": 34, "pct": 13.9, "avg": 51.2}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ K線橫盤的突破剛形成": {"10": {"n": 78, "moonshotN": 9, "pct": 11.5, "avg": 44.6}, "20": {"n": 78, "moonshotN": 13, "pct": 16.7, "avg": 56.1}},
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 876, "moonshotN": 83, "pct": 9.5, "avg": 46.9}, "20": {"n": 876, "moonshotN": 154, "pct": 17.6, "avg": 57.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 距52週高點≤5% ＋ 晨星剛形成": {"10": {"n": 55, "moonshotN": 8, "pct": 14.5, "avg": 44.7}, "20": {"n": 55, "moonshotN": 12, "pct": 21.8, "avg": 46.2}},
    "地量(≤0.5倍均量) ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 晨星剛形成": {"10": {"n": 68, "moonshotN": 7, "pct": 10.3, "avg": 38.3}, "20": {"n": 68, "moonshotN": 14, "pct": 20.6, "avg": 58.9}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1531, "moonshotN": 139, "pct": 9.1, "avg": 46.4}, "20": {"n": 1531, "moonshotN": 266, "pct": 17.4, "avg": 56.1}},
    "多方力道≥65 ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 111, "moonshotN": 10, "pct": 9.0, "avg": 48.2}, "20": {"n": 111, "moonshotN": 18, "pct": 16.2, "avg": 62.5}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ OBV能量潮創60日新高": {"10": {"n": 109, "moonshotN": 8, "pct": 7.3, "avg": 44.5}, "20": {"n": 109, "moonshotN": 15, "pct": 13.8, "avg": 55.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 116, "moonshotN": 12, "pct": 10.3, "avg": 43.5}, "20": {"n": 116, "moonshotN": 17, "pct": 14.7, "avg": 55.4}},
    "KDJ近3日內死亡交叉 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 671, "moonshotN": 59, "pct": 8.8, "avg": 46.0}, "20": {"n": 671, "moonshotN": 104, "pct": 15.5, "avg": 53.9}},
    "近3日向上跳空缺口 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 274, "moonshotN": 25, "pct": 9.1, "avg": 43.4}, "20": {"n": 274, "moonshotN": 42, "pct": 15.3, "avg": 50.1}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 晨星剛形成": {"10": {"n": 162, "moonshotN": 14, "pct": 8.6, "avg": 44.1}, "20": {"n": 162, "moonshotN": 29, "pct": 17.9, "avg": 49.8}},
    "創52週新高 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 276, "moonshotN": 27, "pct": 9.8, "avg": 45.9}, "20": {"n": 276, "moonshotN": 48, "pct": 17.4, "avg": 54.2}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 591, "moonshotN": 43, "pct": 7.3, "avg": 44.4}, "20": {"n": 591, "moonshotN": 82, "pct": 13.9, "avg": 52.9}},
    "地量(≤0.5倍均量) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 夜星剛形成": {"10": {"n": 56, "moonshotN": 11, "pct": 19.6, "avg": 39.8}, "20": {"n": 56, "moonshotN": 12, "pct": 21.4, "avg": 68.3}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 55, "moonshotN": 3, "pct": 5.5, "avg": 37.5}, "20": {"n": 55, "moonshotN": 11, "pct": 20.0, "avg": 41.2}},
    "多方力道≥65 ＋ 量能區間低檔(≤10百分位) ＋ 創52週新高": {"10": {"n": 60, "moonshotN": 2, "pct": 3.3, "avg": 30.8}, "20": {"n": 60, "moonshotN": 11, "pct": 18.3, "avg": 44.4}},
    "量能區間低檔(≤10百分位) ＋ 大盤站上60日均線 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 67, "moonshotN": 3, "pct": 4.5, "avg": 37.5}, "20": {"n": 67, "moonshotN": 12, "pct": 17.9, "avg": 41.7}},
    "量能區間低檔(≤10百分位) ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 56, "moonshotN": 1, "pct": 1.8, "avg": 30.9}, "20": {"n": 56, "moonshotN": 10, "pct": 17.9, "avg": 42.9}},
    "量能區間低檔(≤10百分位) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 69, "moonshotN": 3, "pct": 4.3, "avg": 37.5}, "20": {"n": 69, "moonshotN": 12, "pct": 17.4, "avg": 41.7}},
    "量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為正": {"10": {"n": 69, "moonshotN": 3, "pct": 4.3, "avg": 37.5}, "20": {"n": 69, "moonshotN": 12, "pct": 17.4, "avg": 41.7}},
    "量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月買超": {"10": {"n": 67, "moonshotN": 3, "pct": 4.5, "avg": 37.5}, "20": {"n": 67, "moonshotN": 11, "pct": 16.4, "avg": 40.1}},
    "地量(≤0.5倍均量) ＋ 量能區間高檔(≥90百分位) ＋ 近3日漲停": {"10": {"n": 134, "moonshotN": 13, "pct": 9.7, "avg": 45.6}, "20": {"n": 134, "moonshotN": 20, "pct": 14.9, "avg": 50.3}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 距52週高點≤5%": {"10": {"n": 193, "moonshotN": 5, "pct": 2.6, "avg": 34.4}, "20": {"n": 193, "moonshotN": 27, "pct": 14.0, "avg": 42.2}},
    "近3月均價YoY為正 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 178, "moonshotN": 14, "pct": 7.9, "avg": 45.0}, "20": {"n": 178, "moonshotN": 25, "pct": 14.0, "avg": 50.6}},
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1342, "moonshotN": 107, "pct": 8.0, "avg": 47.4}, "20": {"n": 1342, "moonshotN": 222, "pct": 16.5, "avg": 56.8}},
    "MACD近3日內黃金交叉 ＋ 連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 596, "moonshotN": 51, "pct": 8.6, "avg": 45.2}, "20": {"n": 596, "moonshotN": 93, "pct": 15.6, "avg": 54.2}},
    "月營收連3月年增 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 180, "moonshotN": 16, "pct": 8.9, "avg": 44.5}, "20": {"n": 180, "moonshotN": 26, "pct": 14.4, "avg": 52.2}},
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收) ＋ 月營收創24個月新高": {"10": {"n": 1247, "moonshotN": 136, "pct": 10.9, "avg": 48.4}, "20": {"n": 1247, "moonshotN": 252, "pct": 20.2, "avg": 54.2}},
    "大盤跌破60日均線 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 晨星剛形成": {"10": {"n": 83, "moonshotN": 5, "pct": 6.0, "avg": 41.3}, "20": {"n": 83, "moonshotN": 11, "pct": 13.3, "avg": 44.2}},
    "近20日強於大盤≥10% ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"10": {"n": 186, "moonshotN": 16, "pct": 8.6, "avg": 43.3}, "20": {"n": 186, "moonshotN": 33, "pct": 17.7, "avg": 50.2}},
    "創52週新高 ＋ 近3月均價YoY為負 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 213, "moonshotN": 15, "pct": 7.0, "avg": 45.4}, "20": {"n": 213, "moonshotN": 31, "pct": 14.6, "avg": 51.8}},
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 654, "moonshotN": 50, "pct": 7.6, "avg": 41.3}, "20": {"n": 654, "moonshotN": 95, "pct": 14.5, "avg": 52.9}},
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 月營收連3月年增": {"10": {"n": 759, "moonshotN": 85, "pct": 11.2, "avg": 44.3}, "20": {"n": 759, "moonshotN": 135, "pct": 17.8, "avg": 51.1}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1175, "moonshotN": 107, "pct": 9.1, "avg": 46.9}, "20": {"n": 1175, "moonshotN": 198, "pct": 16.9, "avg": 56.7}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 239, "moonshotN": 21, "pct": 8.8, "avg": 44.3}, "20": {"n": 239, "moonshotN": 33, "pct": 13.8, "avg": 53.1}},
    "多方力道≥65 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2364, "moonshotN": 208, "pct": 8.8, "avg": 46.4}, "20": {"n": 2364, "moonshotN": 373, "pct": 15.8, "avg": 55.5}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 87, "moonshotN": 8, "pct": 9.2, "avg": 36.7}, "20": {"n": 87, "moonshotN": 11, "pct": 12.6, "avg": 44.9}},
    "近3日漲停 ＋ 月營收創24個月新高 ＋ N字底剛形成": {"10": {"n": 118, "moonshotN": 12, "pct": 10.2, "avg": 40.8}, "20": {"n": 118, "moonshotN": 18, "pct": 15.3, "avg": 46.6}},
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2351, "moonshotN": 213, "pct": 9.1, "avg": 46.6}, "20": {"n": 2351, "moonshotN": 388, "pct": 16.5, "avg": 55.7}},
    "多方力道≥80 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1965, "moonshotN": 161, "pct": 8.2, "avg": 46.8}, "20": {"n": 1965, "moonshotN": 313, "pct": 15.9, "avg": 55.4}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 晨星剛形成": {"10": {"n": 204, "moonshotN": 17, "pct": 8.3, "avg": 45.0}, "20": {"n": 204, "moonshotN": 32, "pct": 15.7, "avg": 53.5}},
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1223, "moonshotN": 114, "pct": 9.3, "avg": 46.4}, "20": {"n": 1223, "moonshotN": 213, "pct": 17.4, "avg": 55.2}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": {"10": {"n": 656, "moonshotN": 69, "pct": 10.5, "avg": 44.8}, "20": {"n": 656, "moonshotN": 114, "pct": 17.4, "avg": 54.8}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 近3日漲停": {"10": {"n": 812, "moonshotN": 71, "pct": 8.7, "avg": 44.5}, "20": {"n": 812, "moonshotN": 139, "pct": 17.1, "avg": 53.9}},
    "布林通道高檔(≥80%) ＋ 月營收創24個月新高 ＋ 晨星剛形成": {"10": {"n": 173, "moonshotN": 11, "pct": 6.4, "avg": 38.2}, "20": {"n": 173, "moonshotN": 27, "pct": 15.6, "avg": 44.5}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)": {"10": {"n": 2065, "moonshotN": 172, "pct": 8.3, "avg": 46.5}, "20": {"n": 2065, "moonshotN": 337, "pct": 16.3, "avg": 54.7}},
    "短線過熱(高於月線≥20%) ＋ 外資近5日買超 ＋ N字底剛形成": {"10": {"n": 222, "moonshotN": 27, "pct": 12.2, "avg": 46.6}, "20": {"n": 222, "moonshotN": 41, "pct": 18.5, "avg": 56.4}},
    "MACD近3日內黃金交叉 ＋ OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 628, "moonshotN": 50, "pct": 8.0, "avg": 45.2}, "20": {"n": 628, "moonshotN": 89, "pct": 14.2, "avg": 54.1}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 492, "moonshotN": 39, "pct": 7.9, "avg": 44.7}, "20": {"n": 492, "moonshotN": 67, "pct": 13.6, "avg": 54.6}},
    "大盤站上20日均線 ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 88, "moonshotN": 6, "pct": 6.8, "avg": 41.6}, "20": {"n": 88, "moonshotN": 12, "pct": 13.6, "avg": 57.2}},
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ N字底剛形成": {"10": {"n": 112, "moonshotN": 10, "pct": 8.9, "avg": 51.4}, "20": {"n": 112, "moonshotN": 22, "pct": 19.6, "avg": 49.9}},
    "創52週新高 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 1621, "moonshotN": 155, "pct": 9.6, "avg": 48.1}, "20": {"n": 1621, "moonshotN": 279, "pct": 17.2, "avg": 57.2}},
    "CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1847, "moonshotN": 160, "pct": 8.7, "avg": 48.2}, "20": {"n": 1847, "moonshotN": 307, "pct": 16.6, "avg": 55.8}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 月營收連3月年增": {"10": {"n": 376, "moonshotN": 35, "pct": 9.3, "avg": 45.9}, "20": {"n": 376, "moonshotN": 55, "pct": 14.6, "avg": 52.4}},
    "近20日強於大盤≥10% ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 126, "moonshotN": 10, "pct": 7.9, "avg": 43.7}, "20": {"n": 126, "moonshotN": 17, "pct": 13.5, "avg": 45.3}},
    "多方力道≥65 ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 599, "moonshotN": 50, "pct": 8.3, "avg": 43.1}, "20": {"n": 599, "moonshotN": 80, "pct": 13.4, "avg": 54.6}},
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2465, "moonshotN": 207, "pct": 8.4, "avg": 46.9}, "20": {"n": 2465, "moonshotN": 399, "pct": 16.2, "avg": 54.8}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 1785, "moonshotN": 131, "pct": 7.3, "avg": 43.6}, "20": {"n": 1785, "moonshotN": 274, "pct": 15.4, "avg": 50.7}},
    "短線過熱(高於月線≥20%) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 306, "moonshotN": 23, "pct": 7.5, "avg": 46.8}, "20": {"n": 306, "moonshotN": 46, "pct": 15.0, "avg": 51.4}},
    "地量(≤0.5倍均量) ＋ 量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 184, "moonshotN": 12, "pct": 6.5, "avg": 40.0}, "20": {"n": 184, "moonshotN": 28, "pct": 15.2, "avg": 45.2}},
    "多方力道≥80 ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 294, "moonshotN": 22, "pct": 7.5, "avg": 45.3}, "20": {"n": 294, "moonshotN": 41, "pct": 13.9, "avg": 52.2}},
    "月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 226, "moonshotN": 16, "pct": 7.1, "avg": 44.5}, "20": {"n": 226, "moonshotN": 28, "pct": 12.4, "avg": 51.0}},
    "短線過熱(高於月線≥20%) ＋ 三大法人近3月買超 ＋ 月營收創24個月新高": {"10": {"n": 2127, "moonshotN": 181, "pct": 8.5, "avg": 46.2}, "20": {"n": 2127, "moonshotN": 368, "pct": 17.3, "avg": 53.8}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1452, "moonshotN": 128, "pct": 8.8, "avg": 48.0}, "20": {"n": 1452, "moonshotN": 243, "pct": 16.7, "avg": 55.4}},
    "近20日強於大盤≥10% ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2499, "moonshotN": 213, "pct": 8.5, "avg": 46.2}, "20": {"n": 2499, "moonshotN": 401, "pct": 16.0, "avg": 54.8}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ 近3月均價YoY為負": {"10": {"n": 2077, "moonshotN": 173, "pct": 8.3, "avg": 45.7}, "20": {"n": 2077, "moonshotN": 292, "pct": 14.1, "avg": 53.6}},
    "近20日強於大盤≥10% ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 132, "moonshotN": 13, "pct": 9.8, "avg": 46.7}, "20": {"n": 132, "moonshotN": 16, "pct": 12.1, "avg": 60.9}},
    "量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": {"10": {"n": 58, "moonshotN": 3, "pct": 5.2, "avg": 37.5}, "20": {"n": 58, "moonshotN": 12, "pct": 20.7, "avg": 41.7}},
    "布林通道高檔(≥80%) ＋ 量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 66, "moonshotN": 3, "pct": 4.5, "avg": 37.5}, "20": {"n": 66, "moonshotN": 12, "pct": 18.2, "avg": 41.7}},
    "大盤跌破20日均線 ＋ 月營收連3月年增 ＋ 夜星剛形成": {"10": {"n": 158, "moonshotN": 13, "pct": 8.2, "avg": 40.3}, "20": {"n": 158, "moonshotN": 28, "pct": 17.7, "avg": 53.3}},
    "近20日強於大盤≥10% ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"10": {"n": 63, "moonshotN": 11, "pct": 17.5, "avg": 39.8}, "20": {"n": 63, "moonshotN": 11, "pct": 17.5, "avg": 68.2}},
    "多方力道≥65 ＋ 量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 70, "moonshotN": 2, "pct": 2.9, "avg": 39.6}, "20": {"n": 70, "moonshotN": 12, "pct": 17.1, "avg": 41.7}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 71, "moonshotN": 2, "pct": 2.8, "avg": 39.6}, "20": {"n": 71, "moonshotN": 12, "pct": 16.9, "avg": 41.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 月營收創24個月新高 ＋ 晨星剛形成": {"10": {"n": 71, "moonshotN": 4, "pct": 5.6, "avg": 40.4}, "20": {"n": 71, "moonshotN": 12, "pct": 16.9, "avg": 49.7}},
    "量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 72, "moonshotN": 3, "pct": 4.2, "avg": 37.5}, "20": {"n": 72, "moonshotN": 12, "pct": 16.7, "avg": 41.7}},
    "均線多頭排列(5>20>60且站上月線) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 152, "moonshotN": 14, "pct": 9.2, "avg": 43.2}, "20": {"n": 152, "moonshotN": 25, "pct": 16.4, "avg": 47.5}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2576, "moonshotN": 223, "pct": 8.7, "avg": 46.6}, "20": {"n": 2576, "moonshotN": 421, "pct": 16.3, "avg": 54.9}},
    "地量(≤0.5倍均量) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成": {"10": {"n": 86, "moonshotN": 8, "pct": 9.3, "avg": 37.2}, "20": {"n": 86, "moonshotN": 14, "pct": 16.3, "avg": 58.7}},
    "相對強弱為正(強於大盤) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2567, "moonshotN": 221, "pct": 8.6, "avg": 46.2}, "20": {"n": 2567, "moonshotN": 417, "pct": 16.2, "avg": 54.6}},
    "量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停": {"10": {"n": 715, "moonshotN": 52, "pct": 7.3, "avg": 44.1}, "20": {"n": 715, "moonshotN": 113, "pct": 15.8, "avg": 52.1}},
    "大盤站上60日均線 ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 96, "moonshotN": 8, "pct": 8.3, "avg": 47.2}, "20": {"n": 96, "moonshotN": 15, "pct": 15.6, "avg": 61.8}},
    "創52週新高 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 91, "moonshotN": 8, "pct": 8.8, "avg": 47.2}, "20": {"n": 91, "moonshotN": 14, "pct": 15.4, "avg": 63.8}},
    "多方力道≥65 ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 80, "moonshotN": 8, "pct": 10.0, "avg": 47.2}, "20": {"n": 80, "moonshotN": 12, "pct": 15.0, "avg": 66.5}},
    "距52週高點≤5% ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 414, "moonshotN": 31, "pct": 7.5, "avg": 46.3}, "20": {"n": 414, "moonshotN": 61, "pct": 14.7, "avg": 52.4}},
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 1004, "moonshotN": 66, "pct": 6.6, "avg": 46.5}, "20": {"n": 1004, "moonshotN": 142, "pct": 14.1, "avg": 48.2}},
    "三大法人近3月買超 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 135, "moonshotN": 11, "pct": 8.1, "avg": 45.4}, "20": {"n": 135, "moonshotN": 19, "pct": 14.1, "avg": 54.4}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)": {"10": {"n": 325, "moonshotN": 25, "pct": 7.7, "avg": 40.9}, "20": {"n": 325, "moonshotN": 44, "pct": 13.5, "avg": 53.1}},
    "強勢突破盤 ＋ 月營收連3月年增 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 112, "moonshotN": 7, "pct": 6.2, "avg": 48.2}, "20": {"n": 112, "moonshotN": 15, "pct": 13.4, "avg": 47.6}},
    "創52週新高 ＋ 近3日漲停 ＋ 三大法人近3月賣超": {"10": {"n": 1259, "moonshotN": 115, "pct": 9.1, "avg": 44.8}, "20": {"n": 1259, "moonshotN": 161, "pct": 12.8, "avg": 54.3}},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 323, "moonshotN": 22, "pct": 6.8, "avg": 39.8}, "20": {"n": 323, "moonshotN": 40, "pct": 12.4, "avg": 52.5}},
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 三大法人近3月賣超": {"10": {"n": 1065, "moonshotN": 86, "pct": 8.1, "avg": 44.7}, "20": {"n": 1065, "moonshotN": 128, "pct": 12.0, "avg": 52.8}},
    "創52週新高 ＋ 近3日漲停 ＋ 近3月均價YoY為負": {"10": {"n": 2026, "moonshotN": 162, "pct": 8.0, "avg": 46.0}, "20": {"n": 2026, "moonshotN": 279, "pct": 13.8, "avg": 53.1}},
    "漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2260, "moonshotN": 192, "pct": 8.5, "avg": 47.5}, "20": {"n": 2260, "moonshotN": 364, "pct": 16.1, "avg": 56.4}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日向上跳空缺口 ＋ 近3日漲停": {"10": {"n": 912, "moonshotN": 71, "pct": 7.8, "avg": 47.3}, "20": {"n": 912, "moonshotN": 144, "pct": 15.8, "avg": 56.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成": {"10": {"n": 206, "moonshotN": 16, "pct": 7.8, "avg": 45.8}, "20": {"n": 206, "moonshotN": 32, "pct": 15.5, "avg": 50.1}},
    "短線過熱(高於月線≥20%) ＋ 三大法人近3月買超 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 360, "moonshotN": 31, "pct": 8.6, "avg": 44.3}, "20": {"n": 360, "moonshotN": 52, "pct": 14.4, "avg": 49.2}},
    "相對強弱為負(弱於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停": {"10": {"n": 543, "moonshotN": 36, "pct": 6.6, "avg": 44.5}, "20": {"n": 543, "moonshotN": 74, "pct": 13.6, "avg": 51.7}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": {"10": {"n": 470, "moonshotN": 54, "pct": 11.5, "avg": 46.6}, "20": {"n": 470, "moonshotN": 98, "pct": 20.9, "avg": 55.8}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1475, "moonshotN": 139, "pct": 9.4, "avg": 47.3}, "20": {"n": 1475, "moonshotN": 257, "pct": 17.4, "avg": 57.7}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ 突破飆股大量黑K最高點剛形成": {"10": {"n": 184, "moonshotN": 16, "pct": 8.7, "avg": 48.6}, "20": {"n": 184, "moonshotN": 29, "pct": 15.8, "avg": 55.4}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": {"10": {"n": 795, "moonshotN": 67, "pct": 8.4, "avg": 45.3}, "20": {"n": 795, "moonshotN": 115, "pct": 14.5, "avg": 55.0}},
    "大盤站上60日均線 ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 201, "moonshotN": 16, "pct": 8.0, "avg": 47.6}, "20": {"n": 201, "moonshotN": 29, "pct": 14.4, "avg": 58.7}},
    "相對強弱為正(強於大盤) ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 154, "moonshotN": 14, "pct": 9.1, "avg": 44.8}, "20": {"n": 154, "moonshotN": 21, "pct": 13.6, "avg": 54.4}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 創52週新高": {"10": {"n": 501, "moonshotN": 42, "pct": 8.4, "avg": 43.1}, "20": {"n": 501, "moonshotN": 93, "pct": 18.6, "avg": 53.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 距52週高點≤5% ＋ 近3日漲停": {"10": {"n": 474, "moonshotN": 44, "pct": 9.3, "avg": 43.4}, "20": {"n": 474, "moonshotN": 82, "pct": 17.3, "avg": 54.0}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 三大法人近3月買超": {"10": {"n": 1317, "moonshotN": 114, "pct": 8.7, "avg": 46.7}, "20": {"n": 1317, "moonshotN": 216, "pct": 16.4, "avg": 57.2}},
    "布林通道高檔(≥80%) ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 102, "moonshotN": 8, "pct": 7.8, "avg": 47.2}, "20": {"n": 102, "moonshotN": 15, "pct": 14.7, "avg": 61.8}},
    "MACD近3日內黃金交叉 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 682, "moonshotN": 56, "pct": 8.2, "avg": 45.9}, "20": {"n": 682, "moonshotN": 97, "pct": 14.2, "avg": 56.0}},
    "MACD近3日內黃金交叉 ＋ 近3日向上跳空缺口 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 684, "moonshotN": 46, "pct": 6.7, "avg": 43.3}, "20": {"n": 684, "moonshotN": 92, "pct": 13.5, "avg": 51.9}},
    "回後買上漲全通過 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 937, "moonshotN": 69, "pct": 7.4, "avg": 44.6}, "20": {"n": 937, "moonshotN": 127, "pct": 13.6, "avg": 51.1}},
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 98, "moonshotN": 8, "pct": 8.2, "avg": 46.9}, "20": {"n": 98, "moonshotN": 13, "pct": 13.3, "avg": 51.5}},
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2340, "moonshotN": 168, "pct": 7.2, "avg": 44.3}, "20": {"n": 2340, "moonshotN": 284, "pct": 12.1, "avg": 53.3}},
    "近3日向上跳空缺口 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 140, "moonshotN": 20, "pct": 14.3, "avg": 43.6}, "20": {"n": 140, "moonshotN": 23, "pct": 16.4, "avg": 55.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 外資近5日買超": {"10": {"n": 1288, "moonshotN": 112, "pct": 8.7, "avg": 45.8}, "20": {"n": 1288, "moonshotN": 213, "pct": 16.5, "avg": 55.0}},
    "大盤跌破60日均線 ＋ 外資近5日買超 ＋ 夜星剛形成": {"10": {"n": 93, "moonshotN": 8, "pct": 8.6, "avg": 44.8}, "20": {"n": 93, "moonshotN": 15, "pct": 16.1, "avg": 57.1}},
    "KDJ近3日內死亡交叉 ＋ 距52週高點≤5% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 944, "moonshotN": 73, "pct": 7.7, "avg": 44.1}, "20": {"n": 944, "moonshotN": 134, "pct": 14.2, "avg": 51.0}},
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2284, "moonshotN": 146, "pct": 6.4, "avg": 44.2}, "20": {"n": 2284, "moonshotN": 299, "pct": 13.1, "avg": 50.8}},
    "近20日強於大盤≥10% ＋ 創52週新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 685, "moonshotN": 49, "pct": 7.2, "avg": 45.7}, "20": {"n": 685, "moonshotN": 83, "pct": 12.1, "avg": 52.9}},
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2088, "moonshotN": 130, "pct": 6.2, "avg": 43.7}, "20": {"n": 2088, "moonshotN": 252, "pct": 12.1, "avg": 51.3}},
    "近3日漲停 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"10": {"n": 71, "moonshotN": 9, "pct": 12.7, "avg": 38.6}, "20": {"n": 71, "moonshotN": 13, "pct": 18.3, "avg": 48.0}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ 外資近5日買超": {"10": {"n": 1774, "moonshotN": 154, "pct": 8.7, "avg": 46.4}, "20": {"n": 1774, "moonshotN": 299, "pct": 16.9, "avg": 55.5}},
    "創52週新高 ＋ 近3日漲停 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 420, "moonshotN": 41, "pct": 9.8, "avg": 45.9}, "20": {"n": 420, "moonshotN": 62, "pct": 14.8, "avg": 55.5}},
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 697, "moonshotN": 58, "pct": 8.3, "avg": 45.2}, "20": {"n": 697, "moonshotN": 100, "pct": 14.3, "avg": 54.6}},
    "MACD近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 823, "moonshotN": 67, "pct": 8.1, "avg": 43.9}, "20": {"n": 823, "moonshotN": 116, "pct": 14.1, "avg": 53.9}},
    "MACD近3日內黃金交叉 ＋ 近20日強於大盤≥10% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 763, "moonshotN": 63, "pct": 8.3, "avg": 44.4}, "20": {"n": 763, "moonshotN": 107, "pct": 14.0, "avg": 54.1}},
    "創52週新高 ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2850, "moonshotN": 241, "pct": 8.5, "avg": 45.5}, "20": {"n": 2850, "moonshotN": 371, "pct": 13.0, "avg": 53.8}},
    "KDJ近3日內死亡交叉 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 577, "moonshotN": 39, "pct": 6.8, "avg": 47.9}, "20": {"n": 577, "moonshotN": 69, "pct": 12.0, "avg": 54.1}},
    "多方力道≥80 ＋ 創52週新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 489, "moonshotN": 37, "pct": 7.6, "avg": 44.1}, "20": {"n": 489, "moonshotN": 57, "pct": 11.7, "avg": 54.0}},
    "近20日強於大盤≥10% ＋ 近3日向上跳空缺口 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 100, "moonshotN": 11, "pct": 11.0, "avg": 41.7}, "20": {"n": 100, "moonshotN": 17, "pct": 17.0, "avg": 58.7}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月賣超": {"10": {"n": 200, "moonshotN": 18, "pct": 9.0, "avg": 42.4}, "20": {"n": 200, "moonshotN": 37, "pct": 18.5, "avg": 53.2}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"10": {"n": 158, "moonshotN": 11, "pct": 7.0, "avg": 40.8}, "20": {"n": 158, "moonshotN": 25, "pct": 15.8, "avg": 54.2}},
    "MACD近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 593, "moonshotN": 48, "pct": 8.1, "avg": 46.4}, "20": {"n": 593, "moonshotN": 86, "pct": 14.5, "avg": 56.7}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 837, "moonshotN": 68, "pct": 8.1, "avg": 45.1}, "20": {"n": 837, "moonshotN": 118, "pct": 14.1, "avg": 54.8}},
    "布林通道高檔(≥80%) ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 836, "moonshotN": 68, "pct": 8.1, "avg": 45.1}, "20": {"n": 836, "moonshotN": 118, "pct": 14.1, "avg": 54.8}},
    "強勢突破盤 ＋ 月營收連3月年增 ＋ 晨星剛形成": {"10": {"n": 180, "moonshotN": 8, "pct": 4.4, "avg": 42.7}, "20": {"n": 180, "moonshotN": 23, "pct": 12.8, "avg": 45.9}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 創52週新高 ＋ 月營收創24個月新高": {"10": {"n": 348, "moonshotN": 31, "pct": 8.9, "avg": 43.7}, "20": {"n": 348, "moonshotN": 56, "pct": 16.1, "avg": 56.5}},
    "距52週高點≤5% ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 1958, "moonshotN": 149, "pct": 7.6, "avg": 43.5}, "20": {"n": 1958, "moonshotN": 305, "pct": 15.6, "avg": 51.0}},
    "量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1488, "moonshotN": 130, "pct": 8.7, "avg": 46.5}, "20": {"n": 1488, "moonshotN": 229, "pct": 15.4, "avg": 56.2}},
    "爆量(≥2倍均量) ＋ 月營收連3月年增 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 105, "moonshotN": 5, "pct": 4.8, "avg": 48.3}, "20": {"n": 105, "moonshotN": 15, "pct": 14.3, "avg": 49.0}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月賣超": {"10": {"n": 978, "moonshotN": 90, "pct": 9.2, "avg": 45.2}, "20": {"n": 978, "moonshotN": 137, "pct": 14.0, "avg": 54.9}},
    "CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3154, "moonshotN": 201, "pct": 6.4, "avg": 43.6}, "20": {"n": 3154, "moonshotN": 436, "pct": 13.8, "avg": 50.4}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2415, "moonshotN": 156, "pct": 6.5, "avg": 44.6}, "20": {"n": 2415, "moonshotN": 301, "pct": 12.5, "avg": 52.8}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3050, "moonshotN": 179, "pct": 5.9, "avg": 44.3}, "20": {"n": 3050, "moonshotN": 378, "pct": 12.4, "avg": 51.4}},
    "近3日漲停 ＋ N字底剛形成 ＋ K線橫盤的突破剛形成": {"10": {"n": 138, "moonshotN": 13, "pct": 9.4, "avg": 49.4}, "20": {"n": 138, "moonshotN": 17, "pct": 12.3, "avg": 52.5}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停": {"10": {"n": 622, "moonshotN": 56, "pct": 9.0, "avg": 49.1}, "20": {"n": 622, "moonshotN": 114, "pct": 18.3, "avg": 55.6}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月買超": {"10": {"n": 670, "moonshotN": 48, "pct": 7.2, "avg": 42.9}, "20": {"n": 670, "moonshotN": 97, "pct": 14.5, "avg": 52.5}},
    "MACD近3日內黃金交叉 ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 573, "moonshotN": 44, "pct": 7.7, "avg": 49.1}, "20": {"n": 573, "moonshotN": 79, "pct": 13.8, "avg": 56.6}},
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月均價YoY為負": {"10": {"n": 1787, "moonshotN": 129, "pct": 7.2, "avg": 45.3}, "20": {"n": 1787, "moonshotN": 233, "pct": 13.0, "avg": 53.0}},
    "量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3467, "moonshotN": 209, "pct": 6.0, "avg": 43.4}, "20": {"n": 3467, "moonshotN": 437, "pct": 12.6, "avg": 51.0}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)": {"10": {"n": 514, "moonshotN": 43, "pct": 8.4, "avg": 43.1}, "20": {"n": 514, "moonshotN": 95, "pct": 18.5, "avg": 53.1}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": {"10": {"n": 748, "moonshotN": 80, "pct": 10.7, "avg": 46.9}, "20": {"n": 748, "moonshotN": 135, "pct": 18.0, "avg": 54.3}},
    "地量(≤0.5倍均量) ＋ 近3日向上跳空缺口 ＋ 近3日漲停": {"10": {"n": 879, "moonshotN": 87, "pct": 9.9, "avg": 45.6}, "20": {"n": 879, "moonshotN": 149, "pct": 17.0, "avg": 55.8}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 近3月均價YoY為正": {"10": {"n": 1373, "moonshotN": 108, "pct": 7.9, "avg": 46.7}, "20": {"n": 1373, "moonshotN": 214, "pct": 15.6, "avg": 56.0}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 近3日漲停": {"10": {"n": 737, "moonshotN": 66, "pct": 9.0, "avg": 47.5}, "20": {"n": 737, "moonshotN": 112, "pct": 15.2, "avg": 53.5}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2493, "moonshotN": 194, "pct": 7.8, "avg": 43.8}, "20": {"n": 2493, "moonshotN": 341, "pct": 13.7, "avg": 53.6}},
    "強勢突破盤 ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 433, "moonshotN": 35, "pct": 8.1, "avg": 41.9}, "20": {"n": 433, "moonshotN": 58, "pct": 13.4, "avg": 50.1}},
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3835, "moonshotN": 239, "pct": 6.2, "avg": 43.6}, "20": {"n": 3835, "moonshotN": 503, "pct": 13.1, "avg": 51.0}},
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2566, "moonshotN": 162, "pct": 6.3, "avg": 43.8}, "20": {"n": 2566, "moonshotN": 316, "pct": 12.3, "avg": 51.3}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2638, "moonshotN": 157, "pct": 6.0, "avg": 42.7}, "20": {"n": 2638, "moonshotN": 317, "pct": 12.0, "avg": 51.0}},
    "近3日漲停 ＋ 近3月乖離度為正(營收優於股價) ＋ K線橫盤的突破剛形成": {"10": {"n": 596, "moonshotN": 53, "pct": 8.9, "avg": 46.9}, "20": {"n": 596, "moonshotN": 71, "pct": 11.9, "avg": 52.0}},
    "近20日強於大盤≥10% ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": {"10": {"n": 82, "moonshotN": 8, "pct": 9.8, "avg": 37.8}, "20": {"n": 82, "moonshotN": 16, "pct": 19.5, "avg": 56.3}},
    "距52週高點≤5% ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 110, "moonshotN": 11, "pct": 10.0, "avg": 41.4}, "20": {"n": 110, "moonshotN": 19, "pct": 17.3, "avg": 45.8}},
    "近20日強於大盤≥10% ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 82, "moonshotN": 8, "pct": 9.8, "avg": 47.2}, "20": {"n": 82, "moonshotN": 14, "pct": 17.1, "avg": 63.8}},
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 月營收創24個月新高": {"10": {"n": 657, "moonshotN": 51, "pct": 7.8, "avg": 44.8}, "20": {"n": 657, "moonshotN": 105, "pct": 16.0, "avg": 53.0}},
    "均線多頭排列(5>20>60且站上月線) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 65, "moonshotN": 4, "pct": 6.2, "avg": 41.2}, "20": {"n": 65, "moonshotN": 10, "pct": 15.4, "avg": 45.0}},
    "相對強弱為正(強於大盤) ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 99, "moonshotN": 8, "pct": 8.1, "avg": 47.2}, "20": {"n": 99, "moonshotN": 15, "pct": 15.2, "avg": 61.8}},
    "創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 103, "moonshotN": 8, "pct": 7.8, "avg": 47.2}, "20": {"n": 103, "moonshotN": 15, "pct": 14.6, "avg": 61.8}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 208, "moonshotN": 16, "pct": 7.7, "avg": 41.1}, "20": {"n": 208, "moonshotN": 31, "pct": 14.9, "avg": 53.1}},
    "布林通道高檔(≥80%) ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 142, "moonshotN": 12, "pct": 8.5, "avg": 46.6}, "20": {"n": 142, "moonshotN": 20, "pct": 14.1, "avg": 60.3}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 81, "moonshotN": 7, "pct": 8.6, "avg": 44.6}, "20": {"n": 81, "moonshotN": 11, "pct": 13.6, "avg": 56.0}},
    "KDJ近3日內死亡交叉 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 145, "moonshotN": 13, "pct": 9.0, "avg": 45.3}, "20": {"n": 145, "moonshotN": 16, "pct": 11.0, "avg": 51.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 近3月乖離度為負(股價超前營收)": {"10": {"n": 1120, "moonshotN": 88, "pct": 7.9, "avg": 46.8}, "20": {"n": 1120, "moonshotN": 178, "pct": 15.9, "avg": 57.2}},
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 1209, "moonshotN": 100, "pct": 8.3, "avg": 48.0}, "20": {"n": 1209, "moonshotN": 169, "pct": 14.0, "avg": 53.7}},
    "爆量(≥2倍均量) ＋ 創52週新高 ＋ 月營收創24個月新高": {"10": {"n": 1356, "moonshotN": 106, "pct": 7.8, "avg": 46.8}, "20": {"n": 1356, "moonshotN": 189, "pct": 13.9, "avg": 57.3}},
    "近3日漲停 ＋ 三大法人近3月買超 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 828, "moonshotN": 81, "pct": 9.8, "avg": 46.8}, "20": {"n": 828, "moonshotN": 113, "pct": 13.6, "avg": 54.9}},
    "創52週新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 559, "moonshotN": 42, "pct": 7.5, "avg": 46.3}, "20": {"n": 559, "moonshotN": 71, "pct": 12.7, "avg": 53.2}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2928, "moonshotN": 204, "pct": 7.0, "avg": 45.7}, "20": {"n": 2928, "moonshotN": 362, "pct": 12.4, "avg": 54.2}},
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 371, "moonshotN": 20, "pct": 5.4, "avg": 45.2}, "20": {"n": 371, "moonshotN": 43, "pct": 11.6, "avg": 52.6}},
    "大盤跌破60日均線 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 984, "moonshotN": 62, "pct": 6.3, "avg": 42.6}, "20": {"n": 984, "moonshotN": 108, "pct": 11.0, "avg": 47.8}},
    "地量(≤0.5倍均量) ＋ 近3日向上跳空缺口 ＋ OBV能量潮創60日新高": {"10": {"n": 265, "moonshotN": 24, "pct": 9.1, "avg": 43.0}, "20": {"n": 265, "moonshotN": 47, "pct": 17.7, "avg": 50.6}},
    "距52週高點≤5% ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 334, "moonshotN": 27, "pct": 8.1, "avg": 42.5}, "20": {"n": 334, "moonshotN": 48, "pct": 14.4, "avg": 50.8}},
    "近3日向上跳空缺口 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 2745, "moonshotN": 201, "pct": 7.3, "avg": 45.6}, "20": {"n": 2745, "moonshotN": 382, "pct": 13.9, "avg": 52.7}},
    "近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 79, "moonshotN": 7, "pct": 8.9, "avg": 52.8}, "20": {"n": 79, "moonshotN": 11, "pct": 13.9, "avg": 70.7}},
    "爆量(≥2倍均量) ＋ 距52週高點≤5% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 2694, "moonshotN": 203, "pct": 7.5, "avg": 45.5}, "20": {"n": 2694, "moonshotN": 349, "pct": 13.0, "avg": 55.0}},
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2755, "moonshotN": 188, "pct": 6.8, "avg": 44.3}, "20": {"n": 2755, "moonshotN": 341, "pct": 12.4, "avg": 52.1}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 3514, "moonshotN": 242, "pct": 6.9, "avg": 44.9}, "20": {"n": 3514, "moonshotN": 431, "pct": 12.3, "avg": 53.0}},
    "近20日強於大盤≥10% ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"10": {"n": 341, "moonshotN": 21, "pct": 6.2, "avg": 43.1}, "20": {"n": 341, "moonshotN": 40, "pct": 11.7, "avg": 59.0}},
    "OBV能量潮創60日新高 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 1489, "moonshotN": 130, "pct": 8.7, "avg": 48.2}, "20": {"n": 1489, "moonshotN": 240, "pct": 16.1, "avg": 58.3}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 距52週高點≤5% ＋ 晨星剛形成": {"10": {"n": 180, "moonshotN": 13, "pct": 7.2, "avg": 44.8}, "20": {"n": 180, "moonshotN": 25, "pct": 13.9, "avg": 50.6}},
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3596, "moonshotN": 235, "pct": 6.5, "avg": 42.8}, "20": {"n": 3596, "moonshotN": 483, "pct": 13.4, "avg": 51.6}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2956, "moonshotN": 261, "pct": 8.8, "avg": 45.2}, "20": {"n": 2956, "moonshotN": 384, "pct": 13.0, "avg": 54.2}},
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2956, "moonshotN": 215, "pct": 7.3, "avg": 44.4}, "20": {"n": 2956, "moonshotN": 359, "pct": 12.1, "avg": 53.4}},
    "KDJ近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 1053, "moonshotN": 80, "pct": 7.6, "avg": 45.2}, "20": {"n": 1053, "moonshotN": 126, "pct": 12.0, "avg": 54.0}},
    "近3日漲停 ＋ 近3月乖離度為正(營收優於股價) ＋ 月營收創24個月新高": {"10": {"n": 2305, "moonshotN": 147, "pct": 6.4, "avg": 46.5}, "20": {"n": 2305, "moonshotN": 271, "pct": 11.8, "avg": 53.5}},
    "回後買上漲全通過 ＋ 創52週新高 ＋ 三大法人近3月賣超": {"10": {"n": 613, "moonshotN": 42, "pct": 6.9, "avg": 42.9}, "20": {"n": 613, "moonshotN": 66, "pct": 10.8, "avg": 46.8}},
    "爆量(≥2倍均量) ＋ 大盤跌破60日均線 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 774, "moonshotN": 39, "pct": 5.0, "avg": 40.4}, "20": {"n": 774, "moonshotN": 80, "pct": 10.3, "avg": 47.9}},
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 283, "moonshotN": 33, "pct": 11.7, "avg": 45.0}, "20": {"n": 283, "moonshotN": 52, "pct": 18.4, "avg": 54.3}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ OBV能量潮創60日新高": {"10": {"n": 480, "moonshotN": 41, "pct": 8.5, "avg": 41.0}, "20": {"n": 480, "moonshotN": 82, "pct": 17.1, "avg": 48.7}},
    "多方力道≥80 ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 晨星剛形成": {"10": {"n": 251, "moonshotN": 18, "pct": 7.2, "avg": 42.3}, "20": {"n": 251, "moonshotN": 40, "pct": 15.9, "avg": 52.8}},
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 外資近5日買超": {"10": {"n": 2434, "moonshotN": 169, "pct": 6.9, "avg": 43.7}, "20": {"n": 2434, "moonshotN": 370, "pct": 15.2, "avg": 51.0}},
    "大盤站上20日均線 ＋ 距52週高點≤5% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 5254, "moonshotN": 409, "pct": 7.8, "avg": 44.5}, "20": {"n": 5254, "moonshotN": 755, "pct": 14.4, "avg": 52.7}},
    "近20日強於大盤≥10% ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3927, "moonshotN": 254, "pct": 6.5, "avg": 43.5}, "20": {"n": 3927, "moonshotN": 524, "pct": 13.3, "avg": 51.1}},
    "MACD近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 640, "moonshotN": 47, "pct": 7.3, "avg": 44.8}, "20": {"n": 640, "moonshotN": 84, "pct": 13.1, "avg": 55.0}},
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 突破飆股大量黑K最高點剛形成": {"10": {"n": 405, "moonshotN": 23, "pct": 5.7, "avg": 50.2}, "20": {"n": 405, "moonshotN": 51, "pct": 12.6, "avg": 53.5}},
    "KDJ近3日內黃金交叉 ＋ 近3日漲停 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 575, "moonshotN": 49, "pct": 8.5, "avg": 44.2}, "20": {"n": 575, "moonshotN": 70, "pct": 12.2, "avg": 54.5}},
    "月營收創24個月新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 夜星剛形成": {"10": {"n": 113, "moonshotN": 9, "pct": 8.0, "avg": 47.5}, "20": {"n": 113, "moonshotN": 13, "pct": 11.5, "avg": 57.5}},
    "地量(≤0.5倍均量) ＋ 連續放量(近3日均量≥1.5倍前20日均量) ＋ OBV能量潮創60日新高": {"10": {"n": 65, "moonshotN": 6, "pct": 9.2, "avg": 51.4}, "20": {"n": 65, "moonshotN": 12, "pct": 18.5, "avg": 60.9}},

}
BT_HIT_SHOW = 3   # 摘要表每類最多顯示幾個編號，其餘以「+N」表示

# ── 大盤狀態分組回測（2026-10-05 台股三年回測，20日超額報酬）：每個清單組合在四種大盤狀態下的 [筆數, 勝率%, 同日超額%, t(依日)]，
#    依序為 ①多頭（加權指數站上月線與季線）②多頭回檔（跌破月線、守季線）③空頭反彈（站上月線、季線下）④空頭（跌破月線與季線）；樣本<30筆為 None
REGIME_STATS = {
    "距52週高點≤5% ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": [[2988, 55.2, 4.14, 8.04], [472, 53.8, 3.07, 1.39], [364, 55.5, 2.97, 2.72], [205, 47.8, 1.21, -0.4]],
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": [[2148, 52.9, 4.11, 6.42], [303, 56.1, 5.94, 2.31], [262, 50.8, 2.56, 1.97], [125, 48.8, 2.74, 1.25]],
    "相對強弱為正(強於大盤) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[389, 60.7, 2.44, 4.17], [93, 53.8, 0.91, 1.44], [70, 54.3, 2.08, 1.4], [92, 44.6, 3.21, 2.14]],
    "外資近5日買超 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[2041, 54.3, 6.5, 7.64], [346, 50.6, 2.55, 0.59], [163, 54.6, 3.63, 1.48], [420, 76.7, 5.95, 2.01]],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 距52週高點≤5%": [[1586, 56.1, 4.04, 6.24], [182, 52.7, 3.55, 1.28], [108, 63.0, 7.98, 3.31], [128, 49.2, 3.03, 3.02]],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[229, 65.1, 10.6, 4.88], None, None, None],
    "大盤站上20日均線 ＋ 月營收創24個月新高 ＋ 漲時量≥跌時量1.5倍(近20日)": [[14517, 51.0, 3.42, 13.7], None, [2003, 53.7, 1.9, 5.97], None],
    "大盤站上20日均線 ＋ 月營收創24個月新高 ＋ 近20日強於大盤≥10%": [[7289, 51.5, 3.85, 9.92], None, [1084, 56.5, 3.69, 5.25], None],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [None, None, None, [62, 91.9, 0.44, 0.11]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [None, None, None, [110, 85.5, 0.51, -0.91]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [None, None, None, [139, 83.5, 0.06, -1.62]],
    "KDJ近3日內黃金交叉 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [207, 77.8, 0.02, -0.06]],
    "三大法人近3月買超 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [128, 81.2, 0.55, -1.75]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [235, 77.0, 0.11, 0.81]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [None, None, None, [203, 81.3, 0.03, -0.18]],
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [235, 77.0, 0.11, 0.81]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [None, None, None, [380, 78.7, 0.89, 0.59]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [235, 77.0, 0.11, 0.81]],
    "型態成形中 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [235, 77.0, 0.11, 0.81]],
    "型態突破確認 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [235, 77.0, 0.11, 0.81]],
    "型態剛形成(剛突破) ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [235, 77.0, 0.11, 0.81]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, None, None, [112, 76.8, -0.69, 0.63]],
    "三大法人近3月賣超 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [None, None, None, [107, 72.0, -0.42, 0.98]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [None, None, None, [94, 68.1, 0.27, 1.2]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [None, None, None, [125, 69.6, -0.25, 0.96]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 相對強弱為正(強於大盤)": [None, [52, 73.1, -0.64, -0.35], None, [845, 69.0, -1.64, -1.42]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [None, [97, 48.5, 1.57, 1.15], None, [380, 78.7, 0.89, 0.59]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [None, None, [39, 51.3, 1.13, -0.68], [305, 66.6, 0.44, 0.24]],
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月均價YoY為正": [None, None, None, [84, 66.7, -0.51, -1.0]],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [None, None, None, [50, 66.0, -1.02, -1.98]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [None, None, [352, 50.6, -3.55, -4.68], [2296, 73.8, 0.43, 0.56]],
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, None, [653, 65.2, 0.13, -0.94]],
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能區間高檔(≥90百分位)": [[88, 50.0, -0.75, -0.29], [211, 61.1, 1.08, 0.76], None, [1143, 70.7, -1.93, -1.82]],
    "夜星剛形成 ＋ 布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤)": [None, None, None, None],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [66, 51.5, 1.8, -1.14], [358, 64.2, 0.3, -0.89]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [None, None, [54, 46.3, -2.02, -1.56], [412, 63.1, 0.3, -0.5]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [85, 68.2, 3.4, 1.74], [267, 67.4, 1.53, 0.62]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [None, None, [30, 50.0, 0.48, -0.34], [143, 74.8, 1.0, -0.07]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [None, None, [48, 52.1, 3.92, 0.74], [197, 72.6, -0.36, -0.86]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位)": [None, None, [366, 49.2, -2.52, -3.04], [2190, 71.6, -0.36, 0.27]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, None, [47, 46.8, 1.56, -0.07], [278, 68.0, -0.43, -0.08]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": [None, None, [684, 49.9, -3.11, -6.19], [4666, 70.7, 0.5, 0.22]],
    "型態剛形成(剛突破) ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [None, None, None, [36, 77.8, -0.35, -0.74]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, [74, 68.9, 3.14, 1.62]],
    "夜星剛形成 ＋ 布林通道低檔(≤20%) ＋ 近3月均價YoY為負": [None, None, None, None],
    "大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [104, 46.2, 0.31, -1.53], [653, 65.2, 0.13, -0.94]],
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [104, 46.2, 0.31, -1.53], [653, 65.2, 0.13, -0.94]],
    "型態突破確認 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [104, 46.2, 0.31, -1.53], [653, 65.2, 0.13, -0.94]],
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [104, 46.2, 0.31, -1.53], [653, 65.2, 0.13, -0.94]],
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [None, [1010, 54.3, -0.53, 0.3], None, [2296, 73.8, 0.43, 0.56]],
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [None, None, None, [131, 62.6, -1.04, -2.0]],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤)": [None, None, [204, 64.7, 2.28, 0.55], [292, 63.0, 1.33, 0.24]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, [188, 50.0, 3.12, 1.5], None, [511, 68.1, -0.05, -2.09]],
    "一字底(均線糾結)剛形成 ＋ 三重底剛形成 ＋ 頭肩底剛形成": [[48, 64.6, 1.44, 1.75], None, None, None],
    "一字底(均線糾結)剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[45, 60.0, 0.14, -0.24], None, None, None],
    "地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, None, [60, 56.7, 2.04, 0.32]],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [None, None, None, None],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [None, None, [33, 54.5, 0.76, -1.0], [85, 74.1, 0.78, -0.01]],
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為正(強於大盤)": [[35, 48.6, -2.16, -1.36], [182, 69.2, -0.78, 0.14], None, [1710, 64.2, -1.33, -2.45]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [None, None, [45, 42.2, 2.49, 0.25], [232, 69.8, -0.07, 0.12]],
    "型態突破確認 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [None, None, None, [36, 77.8, -0.35, -0.74]],
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 爆量(≥1.5倍均量)": [None, None, None, [67, 47.8, 1.78, 0.29]],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [None, None, None, None],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [None, [107, 49.5, 2.53, 0.27], None, [197, 72.6, -0.36, -0.86]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [None, [76, 47.4, 0.86, 0.17], None, [232, 69.8, -0.07, 0.12]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, [88, 45.5, -0.91, -0.87], None, [278, 68.0, -0.43, -0.08]],
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [38, 36.8, -2.26, -2.09], [295, 66.4, -0.07, 0.01]],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [314, 62.1, 1.94, -0.23], None],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [43, 60.5, 4.98, 0.86], [193, 47.2, -1.2, -1.83]],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[243, 41.6, -0.15, 0.12], [75, 52.0, 3.19, 1.61], None, [345, 78.6, 0.7, -0.65]],
    "大盤跌破60日均線 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": [None, None, [1289, 50.7, -1.47, -2.49], [1877, 73.8, 1.39, 0.33]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [None, None, [57, 40.4, -0.19, -0.74], [320, 66.6, 0.18, -0.44]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, [1169, 51.5, -0.57, -2.42], [2014, 70.2, -0.18, -1.45]],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 布林通道低檔(≤20%)": [None, None, None, None],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [None, [107, 43.0, -0.94, 0.18], None, [320, 66.6, 0.18, -0.44]],
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位)": [None, [1508, 55.1, 0.07, -0.13], None, [2190, 71.6, -0.36, 0.27]],
    "一字底(均線糾結)剛形成 ＋ 回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成": [[92, 62.0, 2.52, 1.19], None, None, None],
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": [None, [2483, 53.9, -0.49, -0.42], None, [4666, 70.7, 0.5, 0.22]],
    "大盤跌破60日均線 ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為負(弱於大盤)": [None, None, [2697, 51.2, -1.1, -2.32], [3853, 73.0, 1.43, -0.71]],
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, None, [30974, 64.6, 0.61, -2.87]],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [None, None, [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [None, None, None, [95, 66.3, -0.74, -1.43]],
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月乖離度為負(股價超前營收)": [None, None, None, [67, 62.7, -1.34, -3.0]],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[121, 45.5, 1.34, 1.49], None, None, [171, 81.3, 0.13, 1.21]],
    "多方力道≥65 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, [2736, 56.0, -0.04, -2.18], [1383, 72.7, 0.35, 0.52]],
    "夜星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[35, 42.9, 2.94, 0.95], None, None, None],
    "型態剛形成(剛突破) ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": [[33, 36.4, -1.61, -0.63], None, None, [94, 80.9, -0.54, -0.66]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 近3月均價YoY為正": [[158, 67.1, 1.89, 2.31], None, None, None],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [None, [85, 44.7, -0.12, 1.7], None, [143, 74.8, 1.0, -0.07]],
    "一字底(均線糾結)剛形成 ＋ 大盤站上20日均線 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], None, None, None],
    "三大法人近3月買超 ＋ 大盤跌破20日均線 ＋ 跌深反彈盤": [None, [45, 40.0, -1.1, -0.86], None, [50, 66.0, -1.02, -1.98]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": [[53, 45.3, 1.18, -0.04], None, None, None],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 量能區間低檔(≤10百分位)": [None, None, [699, 54.5, -0.56, -2.97], [1106, 65.8, 0.02, -0.14]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": [None, None, [41, 73.2, 1.85, 0.6], None],
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[113, 47.8, 0.36, 0.73], [46, 65.2, -1.3, -0.92], None, [136, 75.7, 2.5, 1.61]],
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": [[64, 56.2, 1.98, 0.64], None, None, None],
    "KDJ近3日內死亡交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": [[92, 65.2, 2.67, 1.94], None, None, None],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為正(營收優於股價)": [[142, 38.7, 0.07, 0.73], None, None, [189, 83.1, 0.77, 0.33]],
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [None, None, [47, 57.4, 2.66, 1.01], [567, 61.2, -0.24, 0.38]],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[56, 37.5, 3.06, 0.87], None, None, None],
    "K線橫盤的突破剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [None, None, [1232, 53.8, -2.12, -2.69], [10213, 63.3, 0.4, -1.36]],
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 布林通道低檔(≤20%)": [None, None, None, None],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, [114, 55.3, 3.77, 1.89], None, [88, 55.7, 4.0, 1.68]],
    "一字底(均線糾結)剛形成 ＋ 大盤站上60日均線 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], None, None, None],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, [66, 50.0, 2.89, 1.07], None, [74, 68.9, 3.14, 1.62]],
    "型態突破確認 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": [[39, 41.0, -1.03, -0.45], None, None, [97, 79.4, -0.6, -0.79]],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[208, 40.9, 1.52, 0.84], [51, 49.0, 1.28, 0.58], None, [200, 79.5, 1.24, -1.54]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 跌深反彈盤": [None, None, None, [93, 66.7, -0.63, -0.79]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ N字底剛形成": [[44, 54.5, 0.32, 0.47], None, None, None],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為負": [[112, 35.7, -3.7, -1.9], None, None, [141, 84.4, 0.04, -0.15]],
    "型態突破確認 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [None, None, [51, 58.8, 2.2, 0.95], [591, 61.1, -0.1, 0.17]],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[71, 49.3, 1.45, 0.95], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [163, 66.9, 1.68, -0.74], [128, 43.8, -1.64, -1.99]],
    "大盤跌破20日均線 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": [None, [1846, 50.9, -0.57, 0.15], None, [1877, 73.8, 1.39, 0.33]],
    "三大法人近3月買超 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, [181, 49.7, 1.04, 1.3], None, [358, 64.2, 0.3, -0.89]],
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間低檔(≤10百分位)": [[35, 28.6, -3.29, -1.47], None, [32, 75.0, 1.24, 0.72], None],
    "KDJ近3日內死亡交叉 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[56, 55.4, 0.01, -0.2], None, None, None],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 頭肩底剛形成": [[54, 63.0, 1.79, 1.9], None, None, None],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 相對強弱為負(弱於大盤)": [None, None, [3688, 51.1, -2.13, -8.31], [17443, 63.6, 0.77, -1.82]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[662, 46.8, -0.55, -2.26], [183, 46.4, -0.15, 0.27], [113, 56.6, 0.22, 0.75], [673, 82.0, 0.29, 0.81]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": [[38, 50.0, 5.41, 1.31], None, None, None],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, [3742, 54.8, -0.19, -3.16], [2768, 71.4, 0.94, 0.61]],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收)": [[235, 41.3, 0.82, 1.34], [63, 50.8, 1.77, 1.22], None, [181, 76.2, 1.28, 0.8]],
    "一字底(均線糾結)剛形成 ＋ 強勢突破盤 ＋ 頭肩底剛形成": [[53, 62.3, 0.84, 1.17], None, None, None],
    "大盤跌破60日均線 ＋ 跌深反彈盤": [None, None, None, [131, 62.6, -1.04, -2.0]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 跌深反彈盤": [None, None, None, [131, 62.6, -1.04, -2.0]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [None, None, None, [131, 62.6, -1.04, -2.0]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [None, None, [90, 53.3, -2.95, -2.23], [513, 62.4, 0.2, -1.04]],
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[47, 42.6, 3.76, 1.44], None, None, None],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[45, 46.7, 6.04, 2.05], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[310, 47.7, 2.66, 2.77], [105, 54.3, 5.25, 1.83], None, [246, 67.5, -0.14, -1.71]],
    "大盤跌破20日均線 ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為負(弱於大盤)": [None, [4154, 51.4, -0.47, -0.46], None, [3853, 73.0, 1.43, -0.71]],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤)": [[130, 46.9, 3.08, 2.43], None, None, [88, 89.8, 3.05, 0.82]],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[385, 47.5, 2.88, 2.7], [135, 50.4, 3.21, 1.31], None, [318, 65.7, 0.28, -2.48]],
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 量能區間低檔(≤10百分位)": [[161, 56.5, 2.29, 1.86], None, None, None],
    "MACD近3日內黃金交叉 ＋ N字底剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[53, 49.1, -0.45, -0.25], None, None, None],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[121, 63.6, 1.17, 1.81], None, None, None],
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[97, 46.4, -1.62, -1.66], [307, 67.8, 1.37, 0.12], [31, 64.5, 0.39, 0.4], [2001, 57.9, -1.01, 0.6]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [None, None, None, [96, 64.6, -1.15, -2.12]],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正": [None, None, [44, 45.5, 0.94, 0.07], [91, 50.5, 3.06, 0.85]],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": [[65, 47.7, 6.33, 1.92], None, None, None],
    "MACD近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": [None, None, [1102, 53.0, -0.73, -4.09], [1410, 65.4, -0.21, -0.61]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[74, 62.2, 10.94, 2.99], None, None, None],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, None, [975, 51.3, -2.73, -5.29], [5759, 62.2, 0.03, 0.38]],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, [7262, 54.0, -0.48, -2.77], [10568, 66.3, 0.91, -1.49]],
    "大盤跌破20日均線 ＋ 複式頭肩底剛形成 ＋ 量能區間高檔(≥90百分位)": [None, [34, 61.8, 2.35, 0.08], None, [33, 42.4, 3.72, 1.45]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為正": [None, None, [10978, 53.2, 0.21, 1.46], [19634, 65.2, 1.11, -1.42]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[56, 60.7, 5.7, 2.27], None, None, None],
    "大盤跌破20日均線 ＋ 晨星剛形成": [None, [312, 47.4, 1.06, 1.47], None, [653, 65.2, 0.13, -0.94]],
    "型態成形中 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, [312, 47.4, 1.06, 1.47], None, [653, 65.2, 0.13, -0.94]],
    "型態突破確認 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, [312, 47.4, 1.06, 1.47], None, [653, 65.2, 0.13, -0.94]],
    "型態剛形成(剛突破) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, [312, 47.4, 1.06, 1.47], None, [653, 65.2, 0.13, -0.94]],
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月乖離度為正(營收優於股價)": [None, None, None, [61, 60.7, -1.79, -1.94]],
    "一字底(均線糾結)剛形成 ＋ 布林通道高檔(≥80%) ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], None, None, None],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": [[140, 62.9, 1.39, 1.71], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[191, 52.4, 3.61, 2.2], [56, 46.4, 0.99, 1.52], None, [116, 81.9, 2.01, 0.33]],
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [None, None, None, [119, 63.0, -0.99, -1.77]],
    "夜星剛形成 ＋ 布林通道低檔(≤20%) ＋ 近3月乖離度為正(營收優於股價)": [None, None, None, None],
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, None, None, [37, 54.1, 0.45, 0.04]],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[58, 55.2, 8.8, 2.51], None, None, None],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, [80, 67.5, 4.48, 0.72], [128, 46.1, -1.41, -3.75]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 跌深反彈盤": [None, None, None, [36, 63.9, -1.28, -0.88]],
    "KDJ近3日內黃金交叉 ＋ 跌深反彈盤 ＋ 近3月乖離度為正(營收優於股價)": [None, None, None, None],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 大盤跌破60日均線": [None, None, [655, 53.0, -0.63, -1.2], [1894, 64.1, 0.42, 0.22]],
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 量能區間低檔(≤10百分位)": [None, None, None, [47, 61.7, 1.26, 1.64]],
    "圓弧底剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], None, None, None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, None, None],
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 量能區間低檔(≤10百分位)": [[41, 48.8, 0.95, 0.5], None, None, None],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[51, 56.9, 2.81, 1.05], None, None, None],
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [None, [65, 43.1, -0.2, -0.27], None, [85, 74.1, 0.78, -0.01]],
    "多方力道≥65 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[39, 48.7, 5.95, 1.62], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[175, 43.4, 2.38, 1.7], [43, 46.5, 0.64, -0.44], None, [231, 70.1, -0.71, -2.42]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [None, [64, 45.3, 1.08, 0.31], None, [96, 64.6, -1.15, -2.12]],
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 量能區間低檔(≤10百分位)": [[33, 48.5, -1.18, -0.79], None, [31, 67.7, 1.31, 0.73], None],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 強勢突破盤": [[273, 59.3, 9.97, 4.05], None, None, None],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [[65, 58.5, 4.83, 1.91], None, None, None],
    "大盤跌破20日均線 ＋ 跌深反彈盤 ＋ 近3月均價YoY為正": [None, [52, 38.5, -2.05, -2.0], None, [84, 66.7, -0.51, -1.0]],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為正": [[294, 42.2, 1.95, 2.03], [82, 47.6, 1.25, 0.69], None, [236, 75.4, 1.38, 0.44]],
    "三大法人近3月賣超 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, [131, 44.3, 1.09, 0.78], None, [295, 66.4, -0.07, 0.01]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [None, [187, 49.7, 2.12, 2.18], None, [305, 66.6, 0.44, 0.24]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": [None, [282, 53.5, 0.6, 0.91], None, [267, 67.4, 1.53, 0.62]],
    "KDJ近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 晨星剛形成": [[340, 48.2, 2.08, 2.2], [103, 51.5, 1.88, 0.5], [30, 56.7, 2.08, 0.26], [268, 67.5, 0.42, -2.49]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, [1227, 45.5, -0.82, -2.76], None, [2014, 70.2, -0.18, -1.45]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, None, [4369, 52.7, -0.61, -3.21], [5774, 65.9, 0.75, -0.29]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收)": [None, None, [9328, 50.8, -0.87, -4.85], [16021, 66.3, 0.87, 0.52]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 三大法人近3月買超": [[153, 60.1, 2.74, 2.59], None, None, None],
    "MACD近3日內黃金交叉 ＋ 布林通道低檔(≤20%) ＋ 量能區間低檔(≤10百分位)": [[86, 37.2, -0.8, -0.24], None, None, [54, 74.1, 1.01, -0.05]],
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [160, 61.9, 0.69, -2.79], [268, 55.2, -0.2, -0.55]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], None, None, None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, [36, 47.2, 1.78, 0.7], None, None],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, [7238, 53.8, -0.79, -2.94], [9571, 65.6, 0.63, -1.2]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": [[55, 56.4, 4.93, 2.1], None, None, None],
    "K線橫盤的突破剛形成 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], None, None, None],
    "K線橫盤的突破剛形成 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], None, None, None],
    "K線橫盤的突破剛形成 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], None, None, None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價)": [None, None, [36, 50.0, 1.85, 0.75], [77, 41.6, 0.11, -0.27]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 近3月乖離度為負(股價超前營收)": [None, None, [1316, 45.5, -3.16, -8.11], [13437, 61.6, 0.62, 0.05]],
    "K線橫盤的突破剛形成 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[60, 60.0, 5.55, 2.2], None, None, None],
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 量能區間低檔(≤10百分位)": [None, [784, 51.0, -0.81, 0.18], None, [1106, 65.8, 0.02, -0.14]],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], None, [380, 78.7, 0.89, 0.59]],
    "型態成形中 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], None, [380, 78.7, 0.89, 0.59]],
    "型態突破確認 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], None, [380, 78.7, 0.89, 0.59]],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], None, [380, 78.7, 0.89, 0.59]],
    "K線橫盤的突破剛形成 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": [[113, 59.3, 2.13, 2.12], None, None, None],
    "一字底(均線糾結)剛形成 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], None, None, None],
    "一字底(均線糾結)剛形成 ＋ 型態成形中 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], None, None, None],
    "一字底(均線糾結)剛形成 ＋ 型態突破確認 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], None, None, None],
    "一字底(均線糾結)剛形成 ＋ 型態剛形成(剛突破) ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], None, None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 相對強弱為負(弱於大盤)": [None, [42, 57.1, 2.34, 1.36], None, None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": [None, None, None, [51, 47.1, 2.16, -0.24]],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 布林通道低檔(≤20%)": [None, None, None, None],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為正(營收優於股價)": [None, None, [172, 55.2, 1.76, -0.52], [278, 53.2, 0.33, -0.15]],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [None, None, [129, 69.0, 1.4, -0.17], [236, 56.8, 0.38, -0.08]],
    "三重底剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[51, 54.9, 4.0, 1.48], None, None, None],
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [None, None, None, [81, 60.5, -1.06, -1.54]],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [154, 62.3, 3.24, 0.59], [274, 52.2, -0.09, -1.76]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 相對強弱為負(弱於大盤)": [[134, 53.7, -0.01, 0.3], None, None, None],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為正": [[84, 44.0, 3.53, 1.23], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [None, None, [741, 52.1, -2.07, -5.42], [7558, 60.5, 0.18, -0.38]],
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 相對強弱為負(弱於大盤)": [None, None, [99, 51.5, 1.5, 1.78], [169, 62.1, 0.76, 0.08]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位)": [None, None, [1223, 48.7, -0.67, -1.81], [1486, 68.3, 1.35, -0.51]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, [6077, 52.0, -0.39, -1.5], [8584, 65.8, 0.21, -1.16]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 大盤站上20日均線": [[244, 58.6, 1.23, 1.82], None, None, None],
    "K線橫盤的突破剛形成 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], None, None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線": [None, None, None, [155, 48.4, 2.66, 0.74]],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "KDJ近3日內黃金交叉 ＋ 型態成形中 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "KDJ近3日內黃金交叉 ＋ 型態突破確認 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "KDJ近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[68, 47.1, 3.59, 1.22], None, None, None],
    "N字底剛形成 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, [268, 53.7, -0.82, -3.02], [31, 61.3, -1.57, -2.52]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [None, [228, 47.8, 1.21, 2.2], None, [412, 63.1, 0.3, -0.5]],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[70, 45.7, 4.62, 1.52], None, None, None],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, [10738, 52.8, 0.03, 0.39], [13073, 66.0, 0.9, -0.84]],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [None, None, [1312, 54.3, -0.85, -3.66], [12102, 59.6, 0.35, 0.73]],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": [None, None, [147, 55.1, 0.94, -0.18], [209, 56.0, 1.03, 0.68]],
    "三大法人近3月買超 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [[890, 42.7, -0.92, -2.09], [377, 58.4, -0.55, -1.29], [103, 51.5, -3.14, -2.2], [968, 76.5, 0.69, 1.4]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[48, 58.3, 5.11, 1.83], None, None, None],
    "一字底(均線糾結)剛形成 ＋ 三大法人近3月買超 ＋ 頭肩底剛形成": [[43, 62.8, 2.42, 2.21], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, [2025, 53.1, 0.23, -0.18], [5979, 61.8, -0.0, -0.48]],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [314, 62.1, 1.94, -0.23], [542, 53.7, -0.15, -1.69]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 多方力道≥65": [[129, 58.9, 2.53, 1.66], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[220, 42.3, 1.15, 0.9], [75, 42.7, 0.15, 0.52], None, [251, 70.9, 0.47, -0.7]],
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [314, 62.1, 1.94, -0.23], [542, 53.7, -0.15, -1.69]],
    "型態突破確認 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [314, 62.1, 1.94, -0.23], [542, 53.7, -0.15, -1.69]],
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, [314, 62.1, 1.94, -0.23], [542, 53.7, -0.15, -1.69]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 母子懷抱(高檔)剛形成": [None, None, None, [131, 57.3, 0.51, 1.07]],
    "一字底(均線糾結)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 頭肩底剛形成": [[54, 64.8, 1.59, 1.7], None, None, None],
    "大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 量能區間低檔(≤10百分位)": [None, None, None, [47, 61.7, 1.26, 1.64]],
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [None, None, None, [27704, 59.0, 0.25, 0.8]],
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, [22364, 53.1, -0.58, -2.38], [26340, 65.0, 0.63, -2.64]],
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [None, None, [1648, 53.5, -0.03, -0.15], [1014, 64.9, 0.29, -0.8]],
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": [None, None, [198, 57.1, 1.62, 1.59], None],
    "三重底剛形成 ＋ 多方力道≥65 ＋ 頭肩底剛形成": [[80, 45.0, -0.79, -0.5], None, None, None],
    "圓弧底剛形成 ＋ 大盤跌破20日均線 ＋ 突破飆股大量黑K最高點剛形成": [None, [41, 51.2, 3.76, 0.78], None, [30, 43.3, 3.27, 0.85]],
    "多方力道≥80 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": [[71, 56.3, 2.79, 1.01], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[70, 48.6, 4.97, 1.52], None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能區間低檔(≤10百分位)": [None, None, None, None],
    "三大法人近3月賣超 ＋ 爆量(≥1.5倍均量) ＋ 跌深反彈盤": [[47, 48.9, -1.05, -0.18], None, None, None],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 近3月均價YoY為正": [None, None, [1407, 48.6, -1.43, -5.23], [16616, 60.0, 0.62, 1.23]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 相對強弱為正(強於大盤)": [[110, 64.5, 2.74, 2.48], None, None, None],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3日向上跳空缺口": [None, None, [1909, 58.7, 0.64, 1.56], [3426, 82.2, 1.6, 0.61]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[88, 51.1, 5.76, 1.85], None, None, None],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": [[34, 47.1, 3.09, 1.07], None, None, None],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 晨星剛形成": [[46, 56.5, 9.41, 1.27], None, None, None],
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[53, 52.8, 4.54, 1.78], None, None, None],
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 晨星剛形成": [None, None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[39, 41.0, 1.08, 0.36], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[46, 52.2, 8.64, 2.45], None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "圓弧底剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], None, None, None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [None, None, None, None],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[30, 30.0, -0.39, 0.11], None, None, None],
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[33, 60.6, 11.08, 1.96], None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 晨星剛形成": [[51, 54.9, 7.42, 1.01], None, None, None],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[34, 47.1, 1.71, 0.18], None, None, None],
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "三重底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": [None, None, None, None],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": [None, None, None, None],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[91, 49.5, 6.49, 2.49], None, None, None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 夜星剛形成": [[33, 48.5, 6.95, 1.49], None, None, None],
    "K線橫盤的突破剛形成 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [[38, 60.5, 7.29, 2.51], None, None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], None, None, None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": [[134, 53.7, 4.97, 1.74], None, None, None],
    "晨星剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[102, 48.0, 6.03, 2.41], None, None, None],
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], None, None, None],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": [[33, 48.5, 5.59, 1.99], None, None, None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[37, 56.8, 6.04, 1.61], None, None, None],
    "夜星剛形成 ＋ 布林通道高檔(≥80%) ＋ 相對強弱為負(弱於大盤)": [None, None, None, None],
    "三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, None, None, None],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], None, None, None],
    "型態成形中 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], None, None, None],
    "型態突破確認 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], None, None, None],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], None, None, None],
    "K線橫盤的突破剛形成 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": [[54, 40.7, 4.02, 1.14], None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[41, 46.3, 4.11, 1.57], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[93, 41.9, 4.79, 1.68], None, None, None],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[54, 57.4, 6.06, 2.19], None, None, None],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "圓弧底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": [[33, 63.6, 7.24, 2.08], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[113, 46.0, 4.33, 1.78], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[32, 34.4, -1.51, -0.5], None, None, None],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [None, None, None, None],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[82, 41.5, 2.5, 1.03], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[72, 50.0, 2.01, 0.87], None, None, None],
    "K線橫盤的突破剛形成 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[137, 48.9, 5.61, 2.72], [31, 38.7, -1.15, -0.87], None, None],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[134, 50.7, 3.83, 1.83], None, None, None],
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": [[50, 40.0, 3.98, 1.22], None, None, None],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[36, 58.3, 7.6, 2.04], None, None, None],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[50, 46.0, 3.18, 0.61], None, None, None],
    "大盤跌破20日均線 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [None, [68, 58.8, 6.98, 2.21], None, [36, 44.4, -2.33, -0.93]],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[73, 46.6, 1.33, 0.56], None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[71, 45.1, 1.48, 0.61], None, None, None],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[37, 54.1, 5.8, 1.71], None, None, None],
    "N字底剛形成 ＋ 三大法人近3月賣超 ＋ 突破ABC修正下降切線剛形成": [[30, 46.7, 4.26, 1.25], None, None, None],
    "圓弧底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": [None, None, None, None],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[34, 47.1, 2.54, 0.42], None, None, None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%)": [[163, 52.1, 3.74, 1.63], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[121, 50.4, 6.54, 2.93], [31, 38.7, -1.67, -1.12], None, None],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], None, None, None],
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 晨星剛形成": [[251, 52.2, 5.0, 3.06], [38, 39.5, -0.78, -0.34], None, None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月均價YoY為負": [[33, 48.5, 5.47, 1.09], None, None, None],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成": [[108, 44.4, 4.49, 1.56], None, None, None],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[60, 45.0, 3.7, 1.46], None, None, None],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[149, 50.3, 2.79, 2.27], [45, 57.8, 4.69, 1.28], None, [40, 70.0, 5.54, 1.95]],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], None, None, None],
    "型態成形中 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], None, None, None],
    "型態突破確認 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], None, None, None],
    "型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[52, 55.8, 5.54, 1.81], None, None, None],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 夜星剛形成": [None, None, None, None],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[57, 50.9, 5.0, 1.69], None, None, None],
    "三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[58, 41.4, 0.36, 0.4], None, None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[207, 48.3, 4.2, 2.12], [39, 33.3, -2.4, -0.38], None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[182, 47.8, 4.93, 2.36], None, None, None],
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[77, 51.9, 6.36, 2.28], None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 強勢突破盤": [[400, 57.5, 7.41, 4.34], [46, 54.3, 1.74, 0.71], [31, 58.1, 10.86, 3.01], None],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[50, 46.0, 5.05, 1.32], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 多方力道≥80 ＋ 晨星剛形成": [[128, 49.2, 5.25, 2.17], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[71, 46.5, 2.77, 0.84], None, None, None],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": [[102, 47.1, 5.71, 2.05], None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": [[40, 37.5, -2.88, -1.23], None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[49, 51.0, 2.62, 0.72], None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[86, 39.5, 3.54, 1.3], None, None, None],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [None, None, None, None],
    "多方力道≥80 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [34, 47.1, 2.53, 0.03], None],
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成": [[68, 44.1, 1.74, 0.61], None, None, None],
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": [[42, 50.0, 5.28, 1.79], None, None, None],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], None, None, None],
    "K線橫盤的突破剛形成 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": [[37, 54.1, 5.96, 1.65], None, None, None],
    "夜星剛形成 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "KDJ近3日內黃金交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": [None, None, None, None],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [None, None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[62, 48.4, 2.69, 0.69], None, None, None],
    "圓弧底剛形成 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[35, 57.1, 5.93, 1.6], None, None, None],
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], None, None, None],
    "多方力道≥80 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], None, [34, 47.1, 2.53, 0.03], None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 布林通道高檔(≥80%)": [None, None, None, None],
    "圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[53, 54.7, 4.11, 1.44], None, None, None],
    "多方力道≥80 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[33, 48.5, 6.8, 1.59], None, None, None],
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[44, 45.5, 2.7, 0.65], None, None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[330, 47.9, 4.11, 2.36], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], None],
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 夜星剛形成": [[37, 45.9, 2.22, 1.02], None, None, None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 近3月乖離度為負(股價超前營收)": [[138, 47.8, 3.4, 1.48], None, None, None],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, None, None, None],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, None, None, None],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為正(營收優於股價)": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 母子懷抱(高檔)剛形成": [[108, 31.5, 0.14, -0.19], [37, 56.8, 10.33, 1.66], None, None],
    "K線橫盤的突破剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": [[87, 47.1, 5.2, 2.06], None, None, None],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成": [[93, 49.5, 5.92, 2.16], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[116, 51.7, 5.39, 2.38], None, None, None],
    "多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], None],
    "多方力道≥65 ＋ 多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], None],
    "型態成形中 ＋ 多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], None],
    "型態突破確認 ＋ 多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], None],
    "型態剛形成(剛突破) ＋ 多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], None],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[56, 57.1, 3.57, 1.24], None, None, None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], None, None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], None, None, None],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[86, 51.2, 5.7, 2.46], None, None, None],
    "N字底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [None, None, None, None],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], None, None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "圓弧底剛形成 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], None, None],
    "型態剛形成(剛突破) ＋ 多方力道≥80 ＋ 量能區間低檔(≤10百分位)": [[36, 36.1, 0.76, 0.28], None, None, None],
    "多方力道≥80 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], None, None],
    "大盤跌破60日均線 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [None, None, None, None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": [None, None, None, None],
    "KDJ近3日內死亡交叉 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[46, 41.3, 2.39, 0.8], None, None, None],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[167, 50.3, 3.23, 1.98], [46, 50.0, 2.31, 0.67], None, [44, 56.8, 3.09, 1.47]],
    "N字底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [None, None, None, None],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能區間高檔(≥90百分位)": [[72, 54.2, 4.72, 1.7], None, None, None],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, None, None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], None, None],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, None, None, None],
    "圓弧底剛形成 ＋ 型態成形中 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], None, None, None],
    "K線橫盤的突破剛形成 ＋ 型態成形中 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], None, None],
    "圓弧底剛形成 ＋ 型態突破確認 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], None, None, None],
    "K線橫盤的突破剛形成 ＋ 型態突破確認 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], None, None],
    "圓弧底剛形成 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], None, None, None],
    "K線橫盤的突破剛形成 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], None, None],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": [[31, 71.0, 7.56, 1.87], None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, None, None, None],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[52, 36.5, 0.81, 0.17], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 晨星剛形成": [[135, 47.4, 2.08, 1.58], None, None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[75, 40.0, 4.76, 2.06], None, None, None],
    "K線橫盤的突破剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], None, None],
    "大盤跌破20日均線 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[66, 50.0, 3.96, 0.62], None, None, [31, 58.1, 3.02, 0.71]],
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[59, 44.1, 1.42, 0.46], None, None, None],
    "多方力道≥80 ＋ 強勢突破盤 ＋ 晨星剛形成": [[149, 41.6, 2.7, 1.49], None, None, None],
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 強勢突破盤": [[104, 56.7, 4.43, 1.67], None, None, None],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[171, 46.8, 3.73, 2.3], [31, 41.9, 0.67, 0.13], None, None],
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[63, 52.4, 8.45, 1.97], None, None, None],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 相對強弱為正(強於大盤)": [[430, 57.4, 6.69, 3.94], [68, 47.1, -0.09, 0.5], [43, 58.1, 7.25, 2.31], [47, 46.8, 5.2, 2.69]],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 相對強弱為正(強於大盤)": [None, [172, 50.0, 2.27, 1.32], None, [143, 47.6, 3.19, 0.91]],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": [[37, 45.9, 2.68, 0.69], None, None, None],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為正": [[42, 45.2, 1.52, 0.2], None, None, None],
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥65": [[164, 51.8, 3.58, 1.55], None, None, None],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[296, 48.0, 4.15, 2.89], [62, 37.1, -0.36, 0.65], None, None],
    "K線橫盤的突破剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": [[103, 46.6, 3.94, 1.67], None, None, None],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[53, 54.7, 4.6, 1.44], None, None, None],
    "N字底剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[39, 56.4, 5.03, 1.47], None, None, None],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[53, 43.4, 2.84, 1.19], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[100, 49.0, 5.48, 2.46], None, None, None],
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[137, 47.4, 2.37, 2.26], [39, 48.7, 1.23, 0.28], None, [35, 71.4, 3.41, 1.3]],
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能區間高檔(≥90百分位)": [[32, 43.8, 1.52, 0.31], None, None, None],
    "三大法人近3月買超 ＋ 多方力道≥65 ＋ 晨星剛形成": [[380, 48.4, 3.35, 2.37], [66, 39.4, 0.27, 0.72], [30, 63.3, 6.62, 0.48], [51, 54.9, 0.27, 0.21]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 相對強弱為正(強於大盤)": [[187, 49.7, 2.97, 1.42], None, None, None],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月乖離度為負(股價超前營收)": [[343, 56.3, 6.29, 3.36], None, None, None],
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": [None, None, None, None],
    "多方力道≥65 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], None, [47, 48.9, 2.73, 0.06], None],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成": [[66, 42.4, 5.48, 1.54], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": [[146, 44.5, 3.65, 1.8], [34, 47.1, 1.67, 0.11], None, None],
    "三大法人近3月賣超 ＋ 強勢突破盤 ＋ 晨星剛形成": [[97, 46.4, 3.68, 1.26], None, None, None],
    "三大法人近3月賣超 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[195, 46.2, 2.88, 1.73], [45, 51.1, 4.73, 0.86], None, None],
    "N字底剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成": [[46, 41.3, 2.3, 0.49], None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[153, 44.4, 1.03, -0.02], [30, 36.7, 8.51, 0.94], None, None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤站上60日均線": [[240, 46.7, 1.84, 1.15], [38, 42.1, 5.43, 0.5], None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [None, None, None, None],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], None, None, None],
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[259, 45.2, 1.7, 1.39], [35, 37.1, -0.74, -0.46], None, None],
    "三大法人近3月賣超 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [[63, 42.9, 1.54, 0.42], None, None, None],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[89, 40.4, 1.44, 0.58], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[336, 44.9, 2.08, 1.99], [61, 50.8, 5.72, 1.83], None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月均價YoY為正": [None, [137, 55.5, 5.07, 2.59], None, [91, 50.5, 3.06, 0.85]],
    "型態成形中 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], None, None, None],
    "型態突破確認 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], None, None, None],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤站上60日均線": [[48, 45.8, 3.52, 1.47], None, None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能區間高檔(≥90百分位)": [None, [93, 46.2, 0.62, 0.26], None, [70, 42.9, 2.7, 1.33]],
    "三重底剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 母子懷抱(高檔)剛形成": [[58, 32.8, -2.41, -1.31], None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": [[77, 42.9, 3.25, 1.3], None, None, None],
    "大盤站上20日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], None, [37, 45.9, 1.33, -0.14], None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 布林通道低檔(≤20%)": [None, None, None, None],
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成": [[36, 41.7, 3.17, 0.96], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[80, 45.0, 4.77, 2.14], None, None, None],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": [[36, 50.0, 4.94, 1.52], None, None, None],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[352, 45.2, 2.45, 1.32], [72, 37.5, -1.06, 0.6], None, [36, 47.2, -3.32, -1.52]],
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 強勢突破盤": [[305, 56.4, 5.65, 3.24], [36, 61.1, 2.2, 0.75], [32, 56.2, 8.42, 2.12], None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 爆量(≥2倍均量)": [None, None, None, None],
    "晨星剛形成 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[167, 44.9, 1.91, 1.39], [43, 48.8, 1.82, 0.67], None, [47, 70.2, 4.27, 1.24]],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[40, 37.5, 0.64, 0.48], None, None, None],
    "晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], None, [74, 68.9, 3.14, 1.62]],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": [[60, 36.7, -0.48, 0.1], None, None, None],
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": [[65, 44.6, 3.62, 1.47], None, None, None],
    "晨星剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成": [[67, 43.3, 2.54, 1.01], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[126, 45.2, 3.38, 1.71], [34, 44.1, -0.53, -0.88], None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月均價YoY為正": [[35, 42.9, 2.17, 1.04], None, None, None],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": [[63, 47.6, 4.03, 0.72], None, None, [34, 58.8, 3.03, 1.22]],
    "型態成形中 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], None, [74, 68.9, 3.14, 1.62]],
    "型態突破確認 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], None, [74, 68.9, 3.14, 1.62]],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], None, [74, 68.9, 3.14, 1.62]],
    "圓弧底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[37, 51.4, 1.51, 0.36], None, None, None],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[87, 43.7, 3.36, 1.14], None, None, None],
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[115, 47.8, 1.97, 1.13], None, None, None],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 突破ABC修正下降切線剛形成": [[39, 38.5, -1.39, -0.25], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[412, 44.7, 1.91, 0.9], [70, 50.0, 4.65, 1.89], None, None],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": [[66, 43.9, 1.71, 0.56], None, None, None],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[505, 46.3, 2.86, 1.79], [104, 35.6, -1.25, 0.06], [46, 50.0, 3.01, 0.18], [61, 44.3, -1.25, -0.44]],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[428, 55.6, 5.8, 3.22], [48, 50.0, 0.08, 0.3], [32, 68.8, 13.04, 2.87], None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為正": [[185, 47.0, 2.01, 1.37], None, None, None],
    "三大法人近3月買超 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": [[185, 49.7, 2.26, 0.74], None, None, None],
    "K線橫盤的突破剛形成 ＋ 三大法人近3月買超 ＋ 晨星剛形成": [[102, 49.0, 4.39, 2.26], None, None, None],
    "N字底剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[40, 37.5, 0.64, 0.48], None, None, None],
    "夜星剛形成 ＋ 近3月均價YoY為負 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[44, 56.8, 4.42, 1.3], None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 布林通道高檔(≥80%)": [[1519, 49.8, 3.41, 3.73], [184, 43.5, 1.79, -0.58], [140, 62.1, 8.4, 2.76], [101, 42.6, 2.8, 1.91]],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[396, 46.2, 2.48, 1.37], [47, 55.3, 7.23, 1.09], [31, 58.1, 2.15, 1.32], None],
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 晨星剛形成": [[63, 46.0, 4.31, 1.38], None, None, None],
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": [[35, 48.6, 1.11, 0.45], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, [64, 51.6, 1.29, -0.19], None, [50, 48.0, 2.43, 0.04]],
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月均價YoY為正": [[283, 42.0, 0.52, 0.48], [113, 55.8, 5.55, 2.41], None, [85, 51.8, 3.6, 1.14]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤站上20日均線": [[240, 46.7, 1.84, 1.15], None, [41, 73.2, 1.85, 0.6], None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[654, 45.1, 2.19, 2.52], None, [39, 51.3, 1.13, -0.68], None],
    "大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], None, None, None],
    "圓弧底剛形成 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], None, None, None],
    "大盤站上60日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], None, None],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[103, 46.6, 1.56, 0.94], None, None, None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], None, [80, 46.2, 0.41, -1.46], None],
    "三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[101, 48.5, 3.52, 1.51], None, None, None],
    "大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, [64, 45.3, 3.23, 0.84], None, [89, 56.2, 2.78, -0.55]],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[111, 50.5, 3.56, 1.79], None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成": [[48, 45.8, 3.52, 1.47], None, None, None],
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": [None, [46, 47.8, 6.58, 1.42], None, None],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位)": [[105, 51.4, 3.91, 1.5], None, None, None],
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 夜星剛形成": [[48, 45.8, 3.52, 1.47], None, None, None],
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 夜星剛形成": [[48, 45.8, 3.52, 1.47], None, None, None],
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": [[48, 45.8, 3.52, 1.47], None, None, None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[799, 44.8, 2.19, 1.65], None, [54, 46.3, -2.02, -1.56], None],
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[73, 39.7, 2.33, 1.3], None, None, None],
    "夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], None, None, None],
    "多方力道≥65 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], None, None],
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, [55, 38.2, -1.38, -0.56], None, None],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[288, 41.3, 2.51, 1.51], [58, 41.4, -0.58, -0.83], [36, 47.2, 1.59, -0.08], [30, 46.7, 0.6, 0.64]],
    "KDJ近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[289, 47.8, 3.18, 2.58], [59, 52.5, 5.06, 1.01], None, None],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[63, 49.2, 2.55, 0.8], None, None, None],
    "大盤站上20日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], None, [62, 48.4, 1.08, -0.28], None],
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[167, 47.9, 3.9, 2.1], None, None, None],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [None, None, None, None],
    "型態成形中 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], None, None, None],
    "型態突破確認 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], None, None, None],
    "型態剛形成(剛突破) ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], None, None, None],
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[81, 44.4, 1.96, 1.34], None, None, None],
    "三大法人近3月賣超 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": [[42, 54.8, 4.23, 1.48], None, None, None],
    "K線橫盤的突破剛形成 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": [[44, 56.8, 4.48, 1.83], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": [None, None, None, None],
    "夜星剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[418, 44.0, 1.42, 1.13], [67, 37.3, -0.34, 0.11], [36, 50.0, 3.02, 0.34], [33, 42.4, -0.94, 0.09]],
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能區間低檔(≤10百分位)": [None, None, None, None],
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[161, 50.9, 4.12, 2.41], None, None, None],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[197, 42.1, 1.66, 1.23], [34, 44.1, 0.4, 0.11], None, None],
    "三重底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": [[30, 50.0, 4.33, 0.77], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[517, 45.6, 2.07, 1.73], [100, 47.0, 3.17, 1.13], [59, 47.5, 1.11, -0.34], [55, 40.0, -2.42, -0.96]],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[654, 45.1, 2.19, 2.52], [187, 49.7, 2.12, 2.18], None, None],
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 頭肩底剛形成": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 晨星剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 強勢突破盤": [[585, 53.5, 5.02, 3.75], None, [54, 51.9, 5.91, 1.84], None],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月均價YoY為正": [[437, 55.4, 5.29, 2.95], [34, 44.1, -1.97, 0.06], None, None],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[102, 45.1, 0.89, 0.54], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[102, 35.3, -2.98, -0.77], [51, 47.1, 0.65, 0.17], None, [44, 47.7, 3.5, 0.42]],
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": [None, None, None, None],
    "KDJ近3日內死亡交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[118, 46.6, 2.6, 2.12], None, None, None],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[94, 50.0, 3.71, 1.72], None, None, None],
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], None, None],
    "夜星剛形成 ＋ 近3月均價YoY為正 ＋ 量能區間高檔(≥90百分位)": [[178, 41.0, 0.37, -0.07], [51, 49.0, 4.48, 1.2], None, [40, 45.0, 3.32, 0.6]],
    "回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 晨星剛形成": [None, None, None, None],
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[63, 46.0, 3.42, 1.21], None, None, None],
    "K線橫盤的突破剛形成 ＋ 回後買上漲全通過 ＋ 晨星剛形成": [[90, 42.2, 3.19, 1.31], None, None, None],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": [[289, 45.0, 1.08, 1.06], [78, 62.8, 7.44, 2.58], None, [57, 56.1, 4.41, 1.19]],
    "多方力道≥65 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], [62, 48.4, 1.08, -0.28], [57, 38.6, -2.47, -0.9]],
    "夜星剛形成 ＋ 大盤跌破20日均線": [None, [214, 51.4, 2.28, 1.51], None, [155, 48.4, 2.66, 0.74]],
    "型態成形中 ＋ 多方力道≥65 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "型態突破確認 ＋ 多方力道≥65 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "型態剛形成(剛突破) ＋ 多方力道≥65 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "MACD近3日內黃金交叉 ＋ N字底剛形成 ＋ 多方力道≥80": [[76, 40.8, 3.99, 1.85], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[101, 49.5, 4.57, 1.92], None, None, None],
    "大盤跌破60日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": [None, None, [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "型態成形中 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], [62, 48.4, 1.08, -0.28], [57, 38.6, -2.47, -0.9]],
    "型態突破確認 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], [62, 48.4, 1.08, -0.28], [57, 38.6, -2.47, -0.9]],
    "型態剛形成(剛突破) ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], [62, 48.4, 1.08, -0.28], [57, 38.6, -2.47, -0.9]],
    "KDJ近3日內死亡交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[50, 54.0, 3.26, 1.42], None, None, None],
    "型態成形中 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, [214, 51.4, 2.28, 1.51], None, [155, 48.4, 2.66, 0.74]],
    "型態突破確認 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, [214, 51.4, 2.28, 1.51], None, [155, 48.4, 2.66, 0.74]],
    "型態剛形成(剛突破) ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, [214, 51.4, 2.28, 1.51], None, [155, 48.4, 2.66, 0.74]],
    "N字底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [None, None, None, None],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], None, None, None],
    "多方力道≥65 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, [131, 48.9, 0.75, 0.89], None, [86, 37.2, 1.89, 0.72]],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 量能區間高檔(≥90百分位)": [[151, 38.4, -0.71, 0.34], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[305, 45.6, 2.92, 2.42], [59, 45.8, 2.05, 0.87], [34, 47.1, 1.16, -0.45], None],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": [[75, 54.7, 3.01, 1.3], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[101, 40.6, -0.36, -0.22], [36, 58.3, 8.25, 1.85], None, [42, 71.4, 4.8, 1.09]],
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, None, None, None],
    "三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[96, 44.8, -0.05, -0.52], [37, 51.4, 5.09, 1.15], None, [55, 56.4, 4.96, 0.02]],
    "K線橫盤的突破剛形成 ＋ 圓弧底剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "夜星剛形成 ＋ 大盤站上60日均線 ＋ 大盤跌破20日均線": [None, [214, 51.4, 2.28, 1.51], None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[799, 44.8, 2.19, 1.65], [228, 47.8, 1.21, 2.2], None, None],
    "大盤跌破60日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, None, [61, 39.3, -0.29, 0.28], [44, 31.8, 1.33, 0.52]],
    "圓弧底剛形成 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], None, None, None],
    "圓弧底剛形成 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], None, None, None],
    "圓弧底剛形成 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[61, 52.5, 4.3, 1.41], None, None, None],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[148, 41.9, 3.72, 2.24], [35, 37.1, -1.36, -0.75], None, None],
    "多方力道≥65 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[94, 41.5, -0.56, -1.13], [32, 43.8, 6.44, 1.2], None, [34, 55.9, 6.13, 0.54]],
    "多方力道≥80 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間低檔(≤10百分位)": [[273, 51.6, 2.6, 2.8], [50, 40.0, 5.22, 1.19], None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [[133, 43.6, 2.15, 0.76], None, None, None],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[283, 49.1, 3.33, 2.05], [38, 71.1, 14.55, 3.28], None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[204, 50.0, 1.51, 1.25], None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[69, 44.9, 3.12, 0.43], None, None, [41, 53.7, 1.21, 0.29]],
    "大盤站上20日均線 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], None, None, None],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[197, 49.2, 3.56, 1.94], [48, 60.4, 5.7, 2.85], None, [64, 79.7, 1.95, 0.96]],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[32, 53.1, 6.77, 1.76], None, None, None],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], None, None, None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量)": [[240, 46.7, 1.84, 1.15], [38, 42.1, 5.43, 0.5], [41, 73.2, 1.85, 0.6], None],
    "多方力道≥65 ＋ 強勢突破盤 ＋ 晨星剛形成": [[240, 40.8, 1.46, 1.12], [45, 33.3, -3.36, -2.61], None, None],
    "多方力道≥65 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [None, None, [43, 51.2, 2.79, 0.57], [86, 37.2, 1.89, 0.72]],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[187, 47.6, 1.76, 1.04], None, None, None],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 多方力道≥80": [[536, 47.6, 3.14, 2.71], [79, 39.2, -1.83, -0.45], [53, 58.5, 3.0, 1.37], [51, 43.1, 4.35, 1.72]],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[334, 49.1, 2.4, 1.51], [46, 58.7, 10.67, 2.69], None, None],
    "KDJ近3日內黃金交叉 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": [[133, 45.1, 2.11, 0.83], None, None, None],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[296, 40.9, 0.79, 1.21], [93, 52.7, 3.72, 1.7], None, [81, 55.6, 4.6, 1.88]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 型態成形中": [[240, 46.7, 1.84, 1.15], [38, 42.1, 5.43, 0.5], [41, 73.2, 1.85, 0.6], None],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 型態突破確認": [[240, 46.7, 1.84, 1.15], [38, 42.1, 5.43, 0.5], [41, 73.2, 1.85, 0.6], None],
    "晨星剛形成 ＋ 近3月均價YoY為負 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[71, 46.5, 1.7, 0.95], None, None, None],
    "回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[121, 52.1, 4.06, 2.07], None, None, None],
    "強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[354, 44.1, 1.69, 1.36], [37, 48.6, 7.67, 1.46], None, None],
    "布林通道高檔(≥80%) ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 強勢突破盤": [[585, 53.5, 5.02, 3.75], [74, 50.0, 0.28, 0.69], None, None],
    "型態成形中 ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "型態突破確認 ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "型態剛形成(剛突破) ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [[62, 58.1, 1.7, 0.78], None, None, None],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[47, 51.1, 1.95, 0.81], None, None, None],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[89, 48.3, 6.45, 2.04], None, None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], None, None],
    "夜星剛形成 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[45, 42.2, 2.5, 1.18], None, None, None],
    "三大法人近3月買超 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[725, 45.1, 1.96, 1.38], None, [66, 51.5, 1.8, -1.14], None],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], None, None, None],
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 近3月均價YoY為負": [[73, 46.6, 5.59, 1.94], None, None, None],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": [[91, 40.7, 0.8, 0.39], [35, 48.6, 4.19, 0.99], None, None],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": [[31, 32.3, -0.9, 0.34], None, None, None],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": [[38, 39.5, -0.67, -0.5], None, None, None],
    "地量(≤0.5倍均量) ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], None, [60, 56.7, 2.04, 0.32]],
    "大盤站上20日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], None, [104, 46.2, 0.31, -1.53], None],
    "MACD近3日內黃金交叉 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[34, 41.2, -1.82, -0.09], None, None, None],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能區間高檔(≥90百分位)": [[32, 28.1, -1.81, -0.26], None, None, None],
    "大盤站上60日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], None, None],
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[68, 52.9, 2.28, 0.73], None, None, None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 爆量(≥2倍均量)": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], None, [60, 56.7, 2.04, 0.32]],
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], None, [60, 56.7, 2.04, 0.32]],
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], None, [60, 56.7, 2.04, 0.32]],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月均價YoY為負": [None, None, None, None],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [None, None, None, None],
    "型態成形中 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], None, [104, 46.2, 0.31, -1.53], None],
    "型態突破確認 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], None, [104, 46.2, 0.31, -1.53], None],
    "型態剛形成(剛突破) ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], None, [104, 46.2, 0.31, -1.53], None],
    "三大法人近3月賣超 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[438, 43.2, 1.49, 0.91], None, [38, 36.8, -2.26, -2.09], None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)": [None, None, [36, 44.4, 1.85, -0.03], [59, 54.2, 4.86, 0.6]],
    "K線橫盤的突破剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[277, 41.5, 2.56, 1.78], [43, 58.1, 4.3, 0.9], [64, 43.8, 5.09, 1.43], [38, 44.7, 6.63, 1.94]],
    "夜星剛形成 ＋ 大盤跌破60日均線": [None, None, [80, 47.5, 1.07, -0.58], [155, 48.4, 2.66, 0.74]],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[139, 41.7, 0.27, 0.05], None, None, None],
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 夜星剛形成": [[196, 42.3, 1.66, 1.67], [41, 51.2, 4.75, 1.52], None, None],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[116, 45.7, 2.79, 1.68], None, None, None],
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%)": [[65, 49.2, -0.02, 0.1], None, None, None],
    "三大法人近3月買超 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[433, 44.8, 1.4, 0.73], [58, 46.6, 2.47, 1.39], [38, 55.3, 4.33, 0.36], [37, 40.5, -2.58, -0.89]],
    "KDJ近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 夜星剛形成": [[103, 37.9, -1.7, -0.29], [30, 53.3, 1.34, -0.09], None, None],
    "地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[112, 38.4, 0.39, -0.04], [51, 52.9, 7.46, 1.54], None, [71, 53.5, 1.76, -0.54]],
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [None, [35, 40.0, -0.43, 0.26], None, [33, 75.8, 2.91, 0.43]],
    "三大法人近3月買超 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[725, 45.1, 1.96, 1.38], [181, 49.7, 1.04, 1.3], None, None],
    "型態成形中 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [None, None, [80, 47.5, 1.07, -0.58], [155, 48.4, 2.66, 0.74]],
    "型態突破確認 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [None, None, [80, 47.5, 1.07, -0.58], [155, 48.4, 2.66, 0.74]],
    "型態剛形成(剛突破) ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [None, None, [80, 47.5, 1.07, -0.58], [155, 48.4, 2.66, 0.74]],
    "圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[119, 47.9, 1.18, 0.51], None, None, None],
    "地量(≤0.5倍均量) ＋ 強勢突破盤": [[585, 53.5, 5.02, 3.75], [74, 50.0, 0.28, 0.69], [54, 51.9, 5.91, 1.84], [47, 46.8, 5.2, 2.69]],
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[770, 47.8, 3.37, 2.42], [108, 42.6, 0.26, -0.04], [73, 47.9, -0.14, 0.06], [50, 46.0, 2.36, 0.86]],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 強勢突破盤": [[585, 53.5, 5.02, 3.75], [74, 50.0, 0.28, 0.69], [54, 51.9, 5.91, 1.84], [47, 46.8, 5.2, 2.69]],
    "KDJ近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[194, 48.5, 3.96, 2.31], [32, 34.4, -1.67, 0.24], None, None],
    "大盤站上60日均線 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], None, None],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[74, 44.6, 1.75, 0.94], None, None, None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能區間高檔(≥90百分位)": [None, None, [38, 47.4, 0.39, 0.01], [70, 42.9, 2.7, 1.33]],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[377, 49.3, 2.5, 1.88], [85, 44.7, -0.12, 1.7], None, None],
    "三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[111, 48.6, 5.59, 2.57], None, None, None],
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 晨星剛形成": [None, None, None, None],
    "多方力道≥65 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": [[271, 41.3, 0.44, 0.04], [84, 58.3, 4.47, 1.8], None, [48, 43.8, 3.1, 1.28]],
    "三大法人近3月買超 ＋ 多方力道≥65 ＋ 夜星剛形成": [[301, 40.5, 0.63, 0.52], [74, 52.7, 2.28, 1.39], None, [49, 44.9, 2.93, 1.12]],
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": [[83, 51.8, 3.62, 1.35], None, None, None],
    "多方力道≥80 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[49, 42.9, -1.19, -1.11], None, None, None],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[60, 41.7, -0.67, -0.06], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[112, 43.8, 0.89, 0.17], None, None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 布林通道高檔(≥80%)": [None, None, None, None],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成": [[51, 43.1, 0.92, 0.42], None, None, None],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[92, 39.1, -2.83, -0.93], None, None, None],
    "MACD近3日內黃金交叉 ＋ 三重底剛形成 ＋ 晨星剛形成": [None, None, None, None],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[41, 43.9, 0.66, 0.62], None, None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月乖離度為正(營收優於股價)": [None, [80, 53.8, 2.36, 1.06], None, [77, 41.6, 0.11, -0.27]],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [None, None, None, None],
    "大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], [312, 47.4, 1.06, 1.47], None, None],
    "KDJ近3日內黃金交叉 ＋ 多方力道≥65 ＋ 晨星剛形成": [[234, 44.4, 2.35, 1.23], [49, 36.7, 1.11, 0.4], None, [36, 47.2, -3.08, -2.01]],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 突破飆股大量黑K最高點剛形成": [[104, 48.1, 4.7, 1.62], None, None, None],
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[106, 45.3, 1.37, 0.81], None, None, None],
    "回後買上漲全通過 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[321, 48.6, 2.99, 0.55], [47, 48.9, 4.4, 1.09], [32, 46.9, -0.99, -0.19], None],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 突破飆股大量黑K最高點剛形成": [[102, 52.0, 5.83, 1.95], None, None, None],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 夜星剛形成": [[66, 30.3, -2.84, -2.11], None, None, None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 相對強弱為正(強於大盤)": [None, None, [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[517, 46.8, 2.14, 1.72], [130, 50.0, 0.95, 0.98], [48, 52.1, 2.72, -0.67], [158, 44.9, -0.9, -0.37]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[133, 53.4, 5.52, 2.46], None, None, None],
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], None, None],
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間高檔(≥90百分位)": [[249, 47.0, 3.28, 2.21], [57, 38.6, 1.15, 0.65], None, [39, 56.4, -3.85, -1.17]],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[411, 44.3, 2.04, 2.5], None, [47, 46.8, 1.56, -0.07], None],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[71, 53.5, 3.29, 1.4], None, None, None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[377, 49.3, 2.5, 1.88], None, [30, 50.0, 0.48, -0.34], None],
    "型態成形中 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], [312, 47.4, 1.06, 1.47], None, None],
    "型態突破確認 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], [312, 47.4, 1.06, 1.47], None, None],
    "型態剛形成(剛突破) ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], [312, 47.4, 1.06, 1.47], None, None],
    "大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [None, None, [188, 47.3, 2.8, 0.79], [138, 44.2, 3.51, 1.79]],
    "K線橫盤的突破剛形成 ＋ 大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成": [None, None, [90, 43.3, 3.53, 1.6], [86, 38.4, 4.45, 2.15]],
    "回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[139, 48.9, 2.47, 0.97], None, None, None],
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": [[80, 58.8, 7.01, 2.22], None, None, None],
    "多方力道≥65 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[292, 46.9, 2.42, 2.13], [50, 52.0, 4.11, 1.44], [40, 37.5, 0.67, 0.66], None],
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[374, 46.5, 3.07, 2.62], [69, 44.9, 1.72, 0.74], [35, 42.9, 4.28, 1.0], [30, 30.0, 2.81, 0.74]],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 複式頭肩底剛形成": [[32, 50.0, 1.59, 0.25], None, None, None],
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[178, 41.6, 1.96, 1.1], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": [[178, 41.0, 0.79, 0.74], [33, 42.4, -1.66, -0.87], None, None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 相對強弱為負(弱於大盤)": [[70, 31.4, 1.39, 0.63], None, None, None],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收)": [[419, 47.3, 2.95, 2.55], [124, 49.2, 2.3, 1.72], None, [124, 52.4, -0.8, 0.2]],
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 量能區間高檔(≥90百分位)": [[216, 45.8, 3.6, 2.36], [47, 40.4, 0.6, 0.28], None, [51, 52.9, -1.52, -1.84]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月乖離度為正(營收優於股價)": [[89, 49.4, 1.87, 0.92], None, None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月乖離度為負(股價超前營收)": [None, [118, 52.5, 2.18, 1.3], None, [59, 54.2, 4.86, 0.6]],
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[45, 42.2, -0.01, -0.32], None, None, None],
    "夜星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 近3月均價YoY為正": [[93, 43.0, 0.26, -0.13], [30, 56.7, 9.43, 1.86], None, [31, 41.9, -1.54, -0.08]],
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[157, 51.0, 2.56, 1.55], None, None, None],
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[36, 33.3, -0.2, -0.18], None, None, None],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(高檔)剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[30, 33.3, 0.42, -0.2], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥65 ＋ 晨星剛形成": [[216, 41.7, 0.91, 1.55], [38, 34.2, -1.46, -0.3], None, [30, 46.7, -3.16, -1.82]],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成": [[267, 48.3, 2.91, 1.52], [46, 37.0, -1.1, -0.45], [34, 58.8, 5.21, 1.33], [34, 32.4, 5.24, 2.12]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[259, 52.9, 3.44, 1.94], [31, 45.2, 1.06, 0.81], None, None],
    "大盤跌破20日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [None, [103, 48.5, 3.46, 1.07], None, [57, 38.6, -2.47, -0.9]],
    "N字底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[30, 30.0, -0.19, 0.13], None, None, None],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 爆量(≥2倍均量)": [[56, 42.9, 0.43, 0.05], None, None, None],
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[189, 44.4, 2.08, 0.63], None, None, None],
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 突破飆股大量黑K最高點剛形成": [[90, 44.4, 5.51, 1.63], None, None, None],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [None, None, None, None],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], None, None, None],
    "三大法人近3月買超 ＋ 回後買上漲全通過 ＋ 多方力道≥80": [[2767, 47.9, 3.0, 4.52], [370, 43.0, -0.44, -0.42], [248, 57.3, 4.13, 3.5], [227, 40.1, 1.08, 0.85]],
    "回後買上漲全通過 ＋ 強勢突破盤 ＋ 晨星剛形成": [[145, 42.8, 2.46, 0.44], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[295, 43.4, 0.9, 1.37], [49, 49.0, 2.35, 1.39], [34, 64.7, 6.45, 0.86], None],
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], None, None],
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[158, 44.3, 1.88, 1.51], None, None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[274, 44.2, 2.56, 2.07], [64, 45.3, 1.08, 0.31], None, None],
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": [[98, 44.9, 4.78, 1.58], None, None, None],
    "夜星剛形成 ＋ 大盤站上60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], None, None, None],
    "N字底剛形成 ＋ 型態成形中 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], None, None, None],
    "N字底剛形成 ＋ 型態突破確認 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], None, None, None],
    "N字底剛形成 ＋ 型態剛形成(剛突破) ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], None, None, None],
    "K線橫盤的突破剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[95, 43.2, 1.25, 0.58], None, None, None],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成": [[47, 42.6, 1.56, 0.84], None, None, None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成": [[107, 34.6, -0.74, -0.79], [36, 47.2, 1.78, 0.7], None, None],
    "三大法人近3月買超 ＋ 夜星剛形成": [[434, 41.7, 0.48, 0.83], [114, 55.3, 3.77, 1.89], [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 多方力道≥80": [[623, 49.0, 3.33, 3.64], [76, 42.1, -0.09, 0.46], [61, 47.5, 0.69, 0.47], [53, 37.7, 3.36, 1.52]],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 量能區間高檔(≥90百分位)": [[601, 48.8, 3.18, 3.19], [92, 44.6, 0.01, 0.4], [57, 50.9, 1.92, 1.16], [65, 32.3, 1.76, 1.51]],
    "三大法人近3月買超 ＋ 強勢突破盤 ＋ 晨星剛形成": [[227, 39.2, 0.68, 0.57], [35, 37.1, -2.25, -0.82], None, None],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[410, 48.3, 2.32, 1.31], [65, 56.9, 6.6, 2.16], [39, 53.8, 0.77, -0.15], [36, 44.4, -2.33, -0.93]],
    "三大法人近3月買超 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[358, 47.2, 2.05, 1.49], [39, 61.5, 11.61, 2.12], None, None],
    "MACD近3日內黃金交叉 ＋ 型態成形中 ＋ 夜星剛形成": [[107, 34.6, -0.74, -0.79], [36, 47.2, 1.78, 0.7], None, None],
    "MACD近3日內黃金交叉 ＋ 型態突破確認 ＋ 夜星剛形成": [[107, 34.6, -0.74, -0.79], [36, 47.2, 1.78, 0.7], None, None],
    "MACD近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": [[107, 34.6, -0.74, -0.79], [36, 47.2, 1.78, 0.7], None, None],
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 52.6, 3.92, 2.09], [30, 60.0, 3.88, 0.98], None, None],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成": [[48, 37.5, -0.86, -0.17], None, None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 爆量(≥2倍均量)": [None, None, None, None],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, None, [47, 48.9, -0.63, -1.22], [88, 48.9, -0.27, -0.18]],
    "晨星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[89, 40.4, 0.16, 0.01], None, None, [33, 66.7, 2.73, 0.27]],
    "三大法人近3月賣超 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[438, 43.2, 1.49, 0.91], [131, 44.3, 1.09, 0.78], None, None],
    "三大法人近3月買超 ＋ 型態成形中 ＋ 夜星剛形成": [[434, 41.7, 0.48, 0.83], [114, 55.3, 3.77, 1.89], [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "三大法人近3月買超 ＋ 型態突破確認 ＋ 夜星剛形成": [[434, 41.7, 0.48, 0.83], [114, 55.3, 3.77, 1.89], [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "三大法人近3月買超 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": [[434, 41.7, 0.48, 0.83], [114, 55.3, 3.77, 1.89], [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[38, 47.4, 2.04, 0.98], None, None, None],
    "多方力道≥65 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": [None, [191, 47.6, 2.34, 0.71], None, [160, 43.1, 3.39, 1.2]],
    "圓弧底剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[32, 46.9, 1.68, 0.7], None, None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[233, 43.8, 1.9, 1.66], [42, 38.1, -1.97, -1.58], None, None],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 大盤站上20日均線": [[3625, 46.7, 2.48, 4.03], None, [369, 56.4, 3.55, 3.74], None],
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": [None, [90, 45.6, 2.71, 0.97], None, [67, 31.3, 2.26, 1.28]],
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": [[254, 43.7, 1.5, 1.71], None, None, None],
    "強勢突破盤 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[71, 40.8, -3.26, -2.38], None, None, None],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[91, 46.2, 1.74, 1.08], None, None, None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[70, 35.7, 1.72, 0.39], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[80, 45.0, 2.04, 1.04], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": [[110, 41.8, -0.62, 0.1], [43, 58.1, 4.55, 1.02], None, [30, 46.7, 0.89, -0.98]],
    "母子懷抱(高檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[97, 43.3, -0.02, -0.75], [46, 43.5, 4.63, 1.06], None, [60, 50.0, 2.91, -0.1]],
    "N字底剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[119, 50.4, 4.47, 2.09], None, None, None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[274, 44.2, 2.56, 2.07], None, None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[411, 44.3, 2.04, 2.5], [88, 45.5, -0.91, -0.87], None, None],
    "夜星剛形成 ＋ 大盤站上20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], None, None, None],
    "大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [None, None, [194, 48.5, 2.9, 0.4], [166, 36.1, 1.75, 1.14]],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[407, 46.7, 2.98, 2.74], [100, 56.0, 2.85, 2.18], None, [165, 66.7, 0.41, -0.36]],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [[62, 53.2, 3.66, 1.26], None, None, None],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], [48, 50.0, -0.24, -0.48], [36, 44.4, -2.33, -0.93]],
    "KDJ近3日內黃金交叉 ＋ 多方力道≥65 ＋ 夜星剛形成": [[109, 38.5, -1.66, -0.68], [40, 40.0, -1.25, -0.57], None, None],
    "三大法人近3月買超 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[486, 45.1, 2.12, 1.98], [56, 46.4, 6.1, 1.64], [45, 66.7, 6.54, 0.47], [50, 44.0, -1.83, -0.82]],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 近3月均價YoY為負": [[1247, 45.1, 2.69, 3.55], [171, 48.5, -0.82, -1.49], [249, 59.4, 4.88, 3.42], [139, 40.3, 3.28, 2.09]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[110, 50.0, 5.23, 2.74], None, None, None],
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 強勢突破盤": [[515, 51.8, 4.29, 3.21], [70, 50.0, 0.55, 0.75], [51, 49.0, 5.84, 1.8], [44, 45.5, 5.06, 2.61]],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[231, 45.5, 3.07, 2.12], [55, 43.6, 1.33, 0.21], None, [33, 30.3, -3.31, -1.27]],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[412, 45.1, 2.27, 1.56], [68, 38.2, 0.65, 0.14], [42, 47.6, -1.81, -0.93], [42, 38.1, -1.87, -0.53]],
    "型態成形中 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], [48, 50.0, -0.24, -0.48], [36, 44.4, -2.33, -0.93]],
    "型態突破確認 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], [48, 50.0, -0.24, -0.48], [36, 44.4, -2.33, -0.93]],
    "型態剛形成(剛突破) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], [48, 50.0, -0.24, -0.48], [36, 44.4, -2.33, -0.93]],
    "N字底剛形成 ＋ 布林通道高檔(≥80%) ＋ 突破飆股大量黑K最高點剛形成": [[130, 50.0, 4.47, 1.89], [31, 41.9, 0.08, 0.14], None, None],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(高檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[38, 39.5, 1.91, 0.52], None, None, None],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[84, 50.0, 0.25, 0.23], None, None, None],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[109, 40.4, -1.0, -0.0], None, None, None],
    "KDJ近3日內死亡交叉 ＋ KDJ近3日內黃金交叉 ＋ 夜星剛形成": [[131, 39.7, -0.76, 0.07], [52, 51.9, 0.5, -0.48], None, [42, 47.6, 1.83, -0.32]],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 量能區間高檔(≥90百分位)": [[71, 45.1, -0.13, 0.27], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], None, [49, 51.0, 2.47, 0.26], None],
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [None, [141, 47.5, 0.27, 0.97], None, [112, 48.2, 1.87, -0.19]],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月均價YoY為正": [[505, 46.3, 2.34, 1.2], [146, 47.9, 1.2, 2.23], [40, 45.0, -2.6, -1.49], [176, 46.6, -1.15, -0.67]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月均價YoY為負": [[82, 51.2, 3.32, 1.5], None, None, None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[556, 45.3, 1.82, 1.43], None, [54, 46.3, -2.06, -1.28], None],
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成": [None, None, [278, 47.1, 1.78, 0.51], None],
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[486, 48.6, 4.11, 2.95], [95, 52.6, -0.89, -0.39], [87, 51.7, 4.25, 1.39], [60, 36.7, 1.1, 0.01]],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[532, 45.5, 2.45, 1.55], [132, 50.0, 0.83, 1.73], [32, 43.8, -4.35, -2.14], [226, 62.4, 0.46, -0.8]],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], [80, 46.2, 0.41, -1.46], [273, 46.5, -0.92, -0.8]],
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[361, 44.9, 2.54, 2.01], [50, 52.0, 7.37, 1.67], None, None],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[124, 44.4, 2.24, 1.46], None, None, None],
    "三大法人近3月賣超 ＋ 多方力道≥80 ＋ 晨星剛形成": [[92, 31.5, -0.81, -0.2], None, None, None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[370, 43.5, 1.79, 1.59], [79, 46.8, 3.44, 0.81], [39, 51.3, 3.74, 0.52], [45, 37.8, -2.07, -0.86]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 夜星剛形成": [[41, 34.1, 2.86, 0.4], None, None, None],
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 頭肩底剛形成": [None, None, None, None],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 爆量(≥1.5倍均量)": [[40, 35.0, -2.63, -1.2], None, None, None],
    "MACD近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], None, None, None],
    "型態成形中 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], [80, 46.2, 0.41, -1.46], [273, 46.5, -0.92, -0.8]],
    "型態突破確認 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], [80, 46.2, 0.41, -1.46], [273, 46.5, -0.92, -0.8]],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], [80, 46.2, 0.41, -1.46], [273, 46.5, -0.92, -0.8]],
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收)": [[239, 38.1, -0.38, -0.04], [98, 51.0, 1.31, 0.92], None, [55, 54.5, 5.44, 0.77]],
    "圓弧底剛形成 ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[74, 32.4, 0.35, 0.43], None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤站上20日均線": [[48, 45.8, 3.52, 1.47], None, None, None],
    "地量(≤0.5倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": [None, [90, 51.1, 3.23, 0.82], None, [131, 57.3, 1.49, 0.05]],
    "夜星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, None, None, None],
    "三重底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[37, 62.2, 4.75, 1.52], None, None, None],
    "大盤跌破60日均線 ＋ 複式頭肩底剛形成 ＋ 頭肩底剛形成": [None, None, None, None],
    "K線橫盤的突破剛形成 ＋ 大盤跌破60日均線 ＋ 複式頭肩底剛形成": [None, None, None, None],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], None, None, None],
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 布林通道高檔(≥80%)": [[2088, 49.3, 2.61, 4.48], [302, 47.0, 2.0, -0.25], [202, 53.5, 5.06, 2.03], [162, 40.1, 1.65, 1.58]],
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[96, 44.8, 2.37, 0.87], None, None, None],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[215, 42.8, 1.82, 1.73], [45, 37.8, -2.0, -0.39], None, None],
    "多方力道≥65 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[129, 42.6, 0.56, 0.68], None, None, None],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [[126, 39.7, 1.16, -0.27], None, None, None],
    "MACD近3日內黃金交叉 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], None, None, None],
    "MACD近3日內黃金交叉 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], None, None, None],
    "MACD近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], None, None, None],
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[144, 45.8, 2.9, 1.42], None, None, None],
    "三重底剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[35, 54.3, 5.05, 1.29], None, None, None],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], None, None],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[469, 42.4, 1.5, 1.45], None, [48, 52.1, 3.92, 0.74], None],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月均價YoY為正": [[96, 50.0, 3.4, 1.25], None, None, None],
    "N字底剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[89, 50.6, 4.07, 1.67], None, None, None],
    "晨星剛形成 ＋ 近3月均價YoY為正 ＋ 量能區間高檔(≥90百分位)": [[312, 46.2, 2.21, 1.41], [73, 39.7, 0.36, 0.81], None, [56, 57.1, -2.79, -1.8]],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, [95, 49.5, 0.77, 0.67], None, [88, 48.9, -0.27, -0.18]],
    "N字底剛形成 ＋ 型態成形中 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], None, None],
    "N字底剛形成 ＋ 型態突破確認 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], None, None],
    "N字底剛形成 ＋ 型態剛形成(剛突破) ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], None, None],
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[408, 44.9, 2.26, 0.91], [58, 51.7, 0.17, 0.08], [83, 53.0, 5.26, 1.66], [56, 46.4, 2.82, 0.98]],
    "夜星剛形成 ＋ 母子懷抱(高檔)剛形成 ＋ 近3月均價YoY為正": [[114, 44.7, 1.14, 0.45], [32, 46.9, 5.71, 1.55], None, None],
    "夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[448, 40.6, 0.34, 0.59], [172, 50.0, 2.27, 1.32], [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], None, [33, 75.8, 2.91, 0.43]],
    "回後買上漲全通過 ＋ 多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[551, 45.9, 1.7, 0.92], [69, 50.7, 5.13, 1.03], [63, 63.5, 0.21, 0.21], [68, 52.9, 3.23, 1.82]],
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[201, 40.8, 0.43, 0.85], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": [[138, 35.5, -1.4, -0.94], None, None, None],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 相對強弱為正(強於大盤)": [[3439, 46.5, 2.56, 4.35], [492, 42.7, -0.87, -1.05], [366, 56.8, 3.66, 3.85], [310, 37.4, 0.86, 0.7]],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 爆量(≥1.5倍均量)": [[497, 46.3, 2.73, 2.57], [81, 35.8, -1.46, -0.35], [48, 52.1, 1.73, 0.82], [57, 38.6, 3.75, 1.9]],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 大盤站上60日均線": [[3625, 46.7, 2.48, 4.03], [492, 42.7, -0.87, -1.05], None, None],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 近3月乖離度為正(營收優於股價)": [[1612, 47.8, 3.15, 4.07], [242, 48.8, -0.89, -0.64], [266, 58.6, 4.72, 2.67], [162, 35.8, 2.39, 1.69]],
    "多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[505, 47.1, 3.36, 2.69], [108, 46.3, -0.6, -0.86], [93, 48.4, 3.66, 0.63], [66, 30.3, 3.73, 1.79]],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[174, 42.5, 1.36, 0.7], None, None, None],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月均價YoY為負": [[148, 48.0, 4.21, 3.42], [34, 47.1, 2.97, 1.31], None, None],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[60, 50.0, 2.38, 0.96], None, None, None],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 夜星剛形成": [[41, 41.5, 2.63, 0.51], None, None, None],
    "型態成形中 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[448, 40.6, 0.34, 0.59], [172, 50.0, 2.27, 1.32], [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "型態突破確認 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[448, 40.6, 0.34, 0.59], [172, 50.0, 2.27, 1.32], [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "型態剛形成(剛突破) ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[448, 40.6, 0.34, 0.59], [172, 50.0, 2.27, 1.32], [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], None, [33, 75.8, 2.91, 0.43]],
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [[337, 48.4, 2.27, 2.05], [65, 43.1, -0.2, -0.27], None, None],
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[177, 42.9, 2.13, 1.44], [54, 51.9, 3.58, 0.73], None, [85, 76.5, 0.28, 0.66]],
    "型態成形中 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], None, [33, 75.8, 2.91, 0.43]],
    "型態突破確認 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], None, [33, 75.8, 2.91, 0.43]],
    "型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], None, [33, 75.8, 2.91, 0.43]],
    "晨星剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[279, 45.9, 2.56, 1.74], [53, 41.5, -1.09, -0.98], [30, 43.3, -2.7, -1.09], [54, 64.8, 0.73, 0.57]],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[556, 45.3, 1.82, 1.43], [121, 40.5, -0.28, -0.34], None, None],
    "大盤跌破20日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [None, [80, 57.5, 4.67, 2.07], None, [44, 31.8, 1.33, 0.52]],
    "回後買上漲全通過 ＋ 大盤跌破60日均線 ＋ 複式頭肩底剛形成": [None, None, [50, 52.0, 1.58, 1.0], None],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[230, 43.5, -0.3, 0.58], [64, 59.4, 3.7, 1.81], None, [35, 62.9, 6.5, 0.97]],
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[135, 51.1, 2.32, 1.25], [38, 47.4, 0.71, 1.79], None, [66, 78.8, -0.19, 0.03]],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [None, None, None, None],
    "三重底剛形成 ＋ 大盤跌破60日均線 ＋ 量能區間低檔(≤10百分位)": [None, None, None, None],
    "大盤站上20日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 複式頭肩底剛形成": [None, None, None, None],
    "大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 頭肩底剛形成": [None, None, None, None],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[272, 40.1, 0.29, 1.03], [36, 55.6, 9.39, 1.69], None, None],
    "MACD近3日內黃金交叉 ＋ 多方力道≥80 ＋ 量能區間低檔(≤10百分位)": [None, None, None, None],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 大盤站上20日均線": [[988, 48.8, 2.56, 3.72], None, [97, 49.5, 1.54, 1.49], None],
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 多方力道≥80": [[765, 49.8, 3.04, 3.69], [117, 45.3, 1.45, 1.06], [61, 57.4, 4.52, 2.06], [79, 36.7, 2.18, 0.84]],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [[212, 45.8, 1.53, 0.33], [33, 45.5, 0.93, 0.46], [34, 47.1, 4.35, 0.55], [32, 43.8, 1.0, 0.74]],
    "三大法人近3月賣超 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[187, 50.8, 2.47, 1.36], None, None, None],
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": [[85, 44.7, 1.34, 0.59], None, None, None],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(高檔)剛形成": [[50, 38.0, 2.74, 1.36], None, None, None],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, None, [37, 54.1, 0.45, 0.04]],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": [[48, 60.4, 4.63, 1.74], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[404, 41.8, 0.13, 1.01], None, [43, 60.5, 4.98, 0.86], None],
    "大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [[337, 48.4, 2.27, 2.05], None, [33, 54.5, 0.76, -1.0], None],
    "回後買上漲全通過 ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[92, 45.7, 1.64, 0.95], None, None, None],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [None, None, None, None],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成": [[137, 38.7, 0.41, 0.02], [55, 47.3, 3.27, 0.69], None, [67, 52.2, 2.15, -0.25]],
    "圓弧底剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 頭肩底剛形成": [[33, 57.6, 4.37, 1.41], None, None, None],
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成": [[983, 46.5, 2.13, 2.2], None, [120, 50.8, 3.78, 1.72], None],
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[274, 44.5, 1.88, 0.56], None, [36, 44.4, -1.97, -1.38], None],
    "K線橫盤的突破剛形成 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [None, [34, 44.1, -0.53, -0.88], None, None],
    "回後買上漲全通過 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[274, 44.5, 1.88, 0.56], [46, 54.3, 2.98, 0.51], None, None],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 晨星剛形成": [[38, 36.8, 0.06, 0.52], None, None, None],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[47, 44.7, 3.42, 0.97], None, None, None],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": [[329, 56.8, 6.99, 3.27], None, None, None],
    "大盤跌破60日均線 ＋ 月營收創24個月新高 ＋ 月營收連3月年增": [None, None, [3867, 57.6, 2.18, 10.41], [6203, 62.4, 2.13, 6.18]],
    "月營收創24個月新高 ＋ 月營收連3月年增 ＋ 近3日漲停": [[2408, 55.0, 6.0, 8.29], [497, 51.1, 3.68, -0.34], [247, 58.3, 3.22, 1.55], [561, 76.3, 5.6, 2.3]],
    "均線多頭排列(5>20>60且站上月線) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[554, 60.5, 2.6, 4.93], [78, 55.1, 0.82, 0.98], [66, 53.0, 1.54, 1.23], [54, 42.6, 4.62, 2.02]],
    "月營收創24個月新高 ＋ 近3日向上跳空缺口 ＋ 近3日漲停": [[1828, 54.0, 5.57, 6.11], [327, 48.6, 2.08, 0.08], [160, 60.0, 6.07, 2.42], [430, 83.3, 5.9, 1.61]],
    "月營收創24個月新高 ＋ 近3日漲停": [[3172, 53.8, 5.74, 7.08], [625, 48.8, 2.35, -0.78], [283, 57.2, 4.74, 2.37], [701, 76.5, 4.58, 1.35]],
    "漲時量≥跌時量1.5倍(近20日) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[353, 53.8, 1.74, 2.96], [49, 59.2, 0.0, 0.44], [61, 54.1, 0.74, 0.42], [43, 37.2, 2.97, 1.45]],
    "三大法人近3月買超 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[2244, 55.1, 6.58, 5.95], [447, 46.8, 2.09, -0.01], [162, 54.9, 4.9, 2.39], [414, 74.2, 4.98, 2.21]],
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[1446, 57.6, 8.08, 7.78], [199, 39.7, -0.41, -2.48], [128, 57.8, 5.29, 2.26], [89, 50.6, 0.51, -0.48]],
    "外資近5日買超 ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[475, 60.4, 2.08, 4.64], [79, 55.7, 0.3, 0.65], [66, 50.0, 1.05, 0.39], [47, 42.6, 5.44, 3.24]],
    "三大法人近3月買超 ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[652, 56.9, 1.92, 4.09], [102, 56.9, 1.22, 1.24], [87, 52.9, 1.57, 1.85], [68, 45.6, 2.7, 1.74]],
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1235, 57.4, 8.26, 6.49], [117, 38.5, 0.82, -1.53], [84, 59.5, 6.79, 2.96], [39, 56.4, 9.29, 2.37]],
    "距52週高點≤5% ＋ 近3月均價YoY為正 ＋ 量能區間低檔(≤10百分位)": [[742, 57.8, 1.88, 4.61], [123, 54.5, 0.66, 0.8], [85, 54.1, 1.91, 2.08], [91, 47.3, 3.71, 2.47]],
    "創52週新高 ＋ 強勢突破盤 ＋ 月營收創24個月新高": [[1679, 56.7, 6.66, 6.44], [183, 44.8, -0.08, -0.48], [118, 55.1, 4.15, 2.55], [97, 46.4, 1.95, 1.17]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 短線過熱(高於月線≥20%)": [[263, 63.1, 11.25, 4.81], None, None, None],
    "月營收創24個月新高 ＋ 月營收連3月年增 ＋ 近3日向上跳空缺口": [[3735, 52.1, 3.41, 5.96], [805, 50.6, 1.73, 0.5], [572, 63.1, 3.68, 3.57], [873, 73.2, 4.34, 2.56]],
    "距52週高點≤5% ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間低檔(≤10百分位)": [[498, 58.8, 2.53, 4.36], [82, 59.8, 1.48, 1.15], [57, 47.4, 1.11, 0.95], [44, 36.4, 2.21, 0.99]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3日漲停": [[183, 59.0, 11.14, 3.73], None, None, None],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為負": [[90, 61.1, 6.92, 2.9], None, None, None],
    "外資近5日買超 ＋ 大盤跌破60日均線 ＋ 月營收創24個月新高": [None, None, [2838, 56.6, 1.75, 5.33], [3932, 61.1, 1.94, 5.78]],
    "多方力道≥80 ＋ 月營收連3月年增 ＋ 量能區間低檔(≤10百分位)": [[272, 56.2, 3.46, 3.62], None, None, None],
    "回後買上漲全通過 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[495, 56.4, 6.45, 4.52], [41, 41.5, 2.51, -0.88], None, None],
    "多方力道≥65 ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[354, 58.5, 2.25, 3.83], [38, 65.8, 4.22, 2.06], [41, 56.1, 0.72, 0.19], [30, 53.3, 1.13, -0.08]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近20日強於大盤≥10%": [[306, 59.5, 9.21, 4.25], [37, 48.6, -1.93, -0.33], None, [31, 45.2, 7.03, 2.54]],
    "距52週高點≤5% ＋ 近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位)": [[129, 67.4, 5.36, 3.0], None, None, None],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[244, 54.9, 4.61, 2.94], [51, 62.7, 10.08, 1.62], [55, 65.5, 13.53, 3.21], None],
    "外資近5日買超 ＋ 近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位)": [[308, 53.9, 2.04, 2.65], [123, 55.3, 6.01, 2.63], [149, 58.4, 1.0, 2.21], [1102, 55.5, 0.84, 0.64]],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[105, 50.5, 7.35, 2.44], None, None, None],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3日向上跳空缺口": [[252, 59.5, 10.02, 3.73], None, None, None],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 短線過熱(高於月線≥20%)": [[115, 49.6, 6.08, 2.34], None, None, None],
    "月營收創24個月新高 ＋ 距52週高點≤5% ＋ 近3日漲停": [[1377, 56.4, 7.26, 6.47], [164, 39.6, -1.07, -1.18], [86, 54.7, 5.99, 1.83], [53, 58.5, 4.8, 1.41]],
    "N字底剛形成 ＋ 月營收連3月年增 ＋ 短線過熱(高於月線≥20%)": [[129, 51.2, 5.65, 1.7], None, None, None],
    "地量(≤0.5倍均量) ＋ 近20日強於大盤≥10% ＋ 近3日漲停": [[549, 52.8, 5.23, 3.07], [138, 59.4, 9.71, 1.69], [31, 61.3, 8.63, 2.2], [78, 57.7, 1.74, 0.3]],
    "K線橫盤的突破剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[152, 54.6, 5.97, 3.54], [30, 63.3, 7.42, 1.67], None, None],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[143, 60.8, 9.79, 2.93], None, None, None],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 地量(≤0.5倍均量)": [[49, 69.4, 13.89, 3.4], None, None, None],
    "月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[188, 59.6, 9.62, 4.22], None, None, None],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[439, 53.5, 6.01, 2.7], [92, 59.8, 12.21, 2.34], None, [51, 58.8, -2.89, -1.19]],
    "地量(≤0.5倍均量) ＋ 相對強弱為正(強於大盤) ＋ 近3日漲停": [[616, 52.3, 5.08, 2.98], [163, 59.5, 9.59, 2.4], [35, 57.1, 6.82, 1.82], [159, 66.0, 0.58, -0.0]],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 近3日漲停": [[518, 53.7, 6.22, 3.18], [94, 54.3, 6.27, 1.44], None, [45, 60.0, 2.7, 0.88]],
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 近3日漲停": [[305, 49.5, 5.68, 4.2], [53, 43.4, -0.17, 0.05], [42, 50.0, 6.46, 1.18], [31, 35.5, -0.28, -0.28]],
    "N字底剛形成 ＋ OBV能量潮創60日新高 ＋ 近3日漲停": [[215, 47.9, 7.02, 3.23], [30, 40.0, -0.54, 0.12], [36, 50.0, 2.36, -0.21], None],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 短線過熱(高於月線≥20%)": [[268, 48.1, 4.56, 1.91], None, [31, 45.2, 1.21, 0.51], None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[442, 54.1, 7.27, 2.9], [78, 55.1, 8.25, 1.59], None, [36, 61.1, 5.0, 1.34]],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[247, 57.9, 7.86, 3.39], None, None, None],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 月營收創24個月新高": [[219, 65.8, 11.17, 4.88], None, None, None],
    "N字底剛形成 ＋ 創52週新高 ＋ 短線過熱(高於月線≥20%)": [[201, 49.3, 5.28, 2.06], None, None, None],
    "CMF資金流買方佔優(近20日≥0.1) ＋ N字底剛形成 ＋ 近3日漲停": [[255, 45.1, 4.58, 2.61], [34, 47.1, -0.5, -0.01], [34, 47.1, -1.83, -0.27], None],
    "K線橫盤的突破剛形成 ＋ 創52週新高 ＋ 近3日漲停": [[247, 51.4, 4.65, 3.47], [41, 53.7, 3.36, 0.75], None, None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ K線橫盤的突破剛形成 ＋ 近3日漲停": [[364, 49.5, 4.93, 4.46], [67, 44.8, 0.43, -0.02], [53, 54.7, 6.78, 1.42], [39, 41.0, -0.22, -0.2]],
    "N字底剛形成 ＋ 近3日漲停 ＋ 量能區間高檔(≥90百分位)": [[287, 44.6, 4.84, 2.97], [44, 38.6, -0.74, 0.29], [41, 61.0, 5.99, 1.47], None],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 近3日漲停": [[418, 53.1, 6.45, 3.25], [47, 48.9, 2.67, 0.6], None, None],
    "K線橫盤的突破剛形成 ＋ 距52週高點≤5% ＋ 近3日漲停": [[383, 50.9, 4.81, 4.19], [70, 51.4, 3.19, 1.12], [37, 51.4, 5.83, 0.75], None],
    "N字底剛形成 ＋ 創52週新高 ＋ 近3日漲停": [[219, 49.8, 7.06, 3.11], [36, 47.2, 0.43, 0.28], None, None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ N字底剛形成 ＋ 近3日漲停": [[265, 45.3, 4.06, 2.29], [40, 40.0, -1.72, -0.11], [33, 57.6, 5.7, 1.34], None],
    "相對強弱為負(弱於大盤) ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[122, 51.6, 3.96, 1.45], None, None, None],
    "地量(≤0.5倍均量) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[429, 59.0, 8.77, 4.61], [35, 31.4, -0.55, -0.22], None, None],
    "N字底剛形成 ＋ 爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[129, 48.8, 6.92, 2.35], None, None, None],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 短線過熱(高於月線≥20%)": [[225, 46.2, 3.66, 1.72], [34, 44.1, 1.7, 0.35], None, None],
    "N字底剛形成 ＋ 近3日漲停 ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[204, 45.1, 4.51, 2.59], [33, 39.4, 0.61, 0.45], [37, 56.8, 3.77, 0.64], None],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 近3日漲停": [[398, 45.5, 4.12, 3.18], None, [65, 55.4, 2.78, 1.01], None],
    "N字底剛形成 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 近3日漲停": [[306, 45.8, 4.27, 2.79], [41, 39.0, -1.69, -0.36], [50, 54.0, 3.95, 0.86], None],
    "N字底剛形成 ＋ 均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%)": [[250, 47.2, 4.53, 1.88], [33, 48.5, 2.61, 0.59], None, None],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 短線過熱(高於月線≥20%)": [[162, 53.1, 8.45, 3.08], None, None, None],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": [[185, 47.6, 5.37, 1.86], None, None, None],
    "N字底剛形成 ＋ 均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停": [[308, 46.1, 4.83, 2.69], [43, 44.2, 0.01, 0.34], [37, 48.6, 0.6, 0.07], None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ N字底剛形成 ＋ 短線過熱(高於月線≥20%)": [[238, 46.2, 3.87, 1.67], [35, 45.7, 1.7, 0.36], None, None],
    "圓弧底剛形成 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[93, 59.1, 7.66, 2.61], None, None, None],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[285, 56.5, 5.56, 2.7], None, None, None],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[353, 56.1, 6.83, 3.49], [32, 25.0, 0.07, -0.13], None, None],
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[265, 54.7, 7.53, 4.45], [44, 59.1, 6.58, 1.65], None, None],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[219, 48.9, 5.41, 1.92], [31, 51.6, 3.17, 0.48], None, None],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 近3日漲停": [[240, 45.8, 4.7, 2.47], [36, 33.3, -3.87, -0.95], None, None],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 短線過熱(高於月線≥20%)": [[254, 46.9, 4.18, 1.78], [35, 45.7, 1.7, 0.36], None, None],
    "N字底剛形成 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": [[238, 47.5, 4.06, 1.96], [32, 43.8, 1.24, 0.33], [30, 43.3, 1.05, 0.48], None],
    "N字底剛形成 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[70, 52.9, 5.0, 1.13], None, None, None],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 近3日漲停": [[313, 45.4, 4.54, 3.01], [42, 38.1, -2.02, -0.29], [36, 55.6, 5.24, 1.16], None],
    "MACD近3日內黃金交叉 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[167, 53.3, 5.43, 2.32], None, None, None],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": [[118, 67.8, 8.62, 3.5], None, None, None],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 近3日漲停": [[410, 53.9, 7.05, 3.05], [67, 53.7, 5.55, 1.21], None, None],
    "N字底剛形成 ＋ 外資近5日買超 ＋ 近3日漲停": [[270, 45.6, 4.64, 2.86], [41, 43.9, -0.38, -0.03], [42, 61.9, 6.38, 1.47], None],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": [[1325, 55.1, 7.3, 5.82], [142, 36.6, 0.33, -1.16], [77, 51.9, 4.31, 1.81], [31, 48.4, 1.91, 0.88]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[84, 58.3, 6.89, 2.47], None, None, None],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 近3日漲停": [[223, 46.6, 5.47, 2.81], [37, 45.9, 0.51, 0.68], None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 月營收創24個月新高": [None, [40, 50.0, 4.49, 1.88], None, None],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3日向上跳空缺口": [[472, 55.3, 6.62, 3.49], [49, 42.9, 6.88, 0.76], None, None],
    "地量(≤0.5倍均量) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停": [[564, 52.0, 5.12, 3.24], [110, 58.2, 9.47, 1.51], [31, 61.3, 8.63, 2.2], [43, 58.1, 1.43, 0.6]],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[108, 44.4, 6.15, 1.97], None, None, None],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%)": [[268, 48.1, 4.56, 1.91], [36, 47.2, 2.33, 0.59], [31, 45.2, 1.21, 0.51], None],
    "N字底剛形成 ＋ 布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%)": [[268, 48.1, 4.56, 1.91], [35, 48.6, 2.26, 0.54], [31, 45.2, 1.21, 0.51], None],
    "大盤跌破20日均線 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [None, [211, 34.6, -1.12, -1.71], None, [76, 51.3, 5.35, 2.12]],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10%": [[261, 47.5, 3.98, 1.62], [36, 47.2, 2.33, 0.59], [31, 45.2, 1.21, 0.51], None],
    "N字底剛形成 ＋ 近20日強於大盤≥10% ＋ 近3日漲停": [[307, 46.3, 4.29, 2.46], [45, 44.4, 0.05, -0.01], [43, 55.8, 3.7, 0.95], None],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[577, 51.1, 4.37, 2.65], [192, 64.6, 11.6, 2.82], [36, 66.7, 10.26, 2.59], [278, 78.4, 3.63, 0.84]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[323, 53.6, 6.43, 4.47], [52, 61.5, 9.65, 2.46], [30, 36.7, 6.62, 1.31], None],
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 短線過熱(高於月線≥20%)": [[216, 49.1, 5.56, 2.13], None, None, None],
    "地量(≤0.5倍均量) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 近3日漲停": [[487, 52.8, 6.14, 2.95], [98, 57.1, 6.6, 1.82], None, [77, 61.0, -1.97, -0.89]],
    "月營收連3月年增 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[370, 55.4, 6.15, 4.27], [65, 61.5, 6.92, 1.96], [43, 44.2, 2.0, -0.03], None],
    "N字底剛形成 ＋ 布林通道高檔(≥80%) ＋ 近3日漲停": [[384, 45.8, 4.35, 3.07], [52, 46.2, -0.14, 0.22], [58, 55.2, 2.92, 0.61], None],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[1423, 56.0, 6.52, 5.97], [124, 38.7, 0.46, -0.96], [74, 54.1, 7.03, 2.79], [48, 56.2, 11.39, 2.85]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 月營收創24個月新高": [[172, 66.3, 14.11, 4.52], None, None, None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": [[56, 46.4, 5.81, 1.93], None, None, None],
    "N字底剛形成 ＋ 距52週高點≤5% ＋ 近3日漲停": [[250, 50.4, 6.73, 3.37], [39, 46.2, 0.42, 0.05], None, None],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間高檔(≥90百分位)": [[177, 47.5, 5.84, 2.22], None, None, None],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 短線過熱(高於月線≥20%)": [[207, 48.8, 5.64, 2.24], None, None, None],
    "地量(≤0.5倍均量) ＋ 距52週高點≤5% ＋ 近3日漲停": [[303, 55.8, 7.32, 3.9], None, None, None],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 短線過熱(高於月線≥20%)": [[268, 48.1, 4.56, 1.91], [36, 47.2, 2.33, 0.59], None, None],
    "大盤跌破60日均線 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [None, None, [148, 50.7, 3.8, 2.18], [76, 51.3, 5.35, 2.12]],
    "地量(≤0.5倍均量) ＋ 月營收創24個月新高 ＋ 近3日漲停": [[252, 61.1, 10.57, 4.28], [81, 70.4, 12.81, 1.22], None, [90, 87.8, 7.87, -0.37]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": [[49, 55.1, 8.23, 1.12], None, None, None],
    "創52週新高 ＋ 近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位)": [[51, 72.5, 7.4, 2.83], None, None, None],
    "漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[53, 69.8, 6.31, 2.14], None, None, None],
    "晨星剛形成 ＋ 月營收創24個月新高 ＋ 距52週高點≤5%": [[89, 55.1, 4.86, 1.89], None, None, None],
    "N字底剛形成 ＋ 近3日漲停 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[118, 37.3, 2.63, 1.58], None, None, None],
    "月營收創24個月新高 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[261, 57.5, 7.94, 3.73], [75, 56.0, 8.09, 1.38], None, [77, 89.6, 10.7, 1.82]],
    "均線多頭排列(5>20>60且站上月線) ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[101, 41.6, 1.72, 0.54], None, None, None],
    "強勢突破盤 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[271, 54.6, 6.24, 2.78], [48, 52.1, 4.68, 1.31], None, None],
    "N字底剛形成 ＋ 月營收連3月年增 ＋ 近3日漲停": [[182, 48.9, 5.04, 2.3], None, None, None],
    "K線橫盤的突破剛形成 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[47, 68.1, 12.74, 3.15], None, None, None],
    "月營收創24個月新高 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": [[690, 51.2, 7.01, 4.26], [80, 36.2, -1.77, -0.84], [61, 49.2, 5.97, 2.16], [45, 60.0, 12.5, 3.38]],
    "晨星剛形成 ＋ 距52週高點≤5% ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[43, 62.8, 4.91, 1.27], None, None, None],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": [[45, 51.1, 7.92, 1.08], None, None, None],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1306, 56.0, 6.83, 6.11], [116, 38.8, 0.45, -1.06], [64, 54.7, 7.83, 3.07], [45, 55.6, 8.89, 2.5]],
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[82, 45.1, 3.03, 0.92], None, None, None],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[74, 59.5, 9.0, 3.45], None, None, None],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[92, 44.6, 0.6, -0.24], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[533, 53.8, 5.49, 3.26], [77, 35.1, -0.54, -2.32], None, [32, 50.0, 2.05, 0.65]],
    "月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日向上跳空缺口": [[214, 54.2, 5.7, 3.32], None, None, None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[120, 51.7, 4.63, 1.88], None, None, None],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": [[217, 59.4, 8.95, 4.25], None, None, None],
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[450, 43.6, 1.27, 1.98], [68, 45.6, 0.18, -0.6], [37, 51.4, 3.88, 1.24], [36, 47.2, 0.37, -0.12]],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 漲時量≥跌時量1.5倍(近20日)": [None, None, None, None],
    "短線過熱(高於月線≥20%) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[49, 75.5, 7.95, 3.03], None, None, None],
    "創52週新高 ＋ 多方力道≥65 ＋ 量能區間低檔(≤10百分位)": [[50, 68.0, 7.0, 2.72], None, None, None],
    "大盤站上60日均線 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[59, 71.2, 5.89, 2.21], None, None, None],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[46, 73.9, 7.6, 2.94], None, None, None],
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[57, 73.7, 6.28, 2.26], None, None, None],
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為正 ＋ 量能區間低檔(≤10百分位)": [[58, 72.4, 6.51, 2.51], None, None, None],
    "三大法人近3月買超 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[57, 70.2, 5.16, 1.91], None, None, None],
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 量能區間高檔(≥90百分位)": [[87, 37.9, -2.94, -0.63], None, None, None],
    "夜星剛形成 ＋ 月營收創24個月新高 ＋ 近3月均價YoY為正": [[110, 52.7, 4.85, 2.17], [34, 52.9, 5.82, 2.04], None, None],
    "OBV能量潮創60日新高 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1118, 55.0, 6.86, 5.93], [98, 34.7, -0.54, -0.73], [81, 48.1, 3.3, 1.72], [45, 55.6, 8.57, 2.33]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[455, 48.1, 4.7, 3.44], [65, 44.6, -0.87, -0.96], [39, 53.8, 3.1, 0.81], [37, 48.6, 0.7, -0.0]],
    "夜星剛形成 ＋ 月營收創24個月新高 ＋ 月營收連3月年增": [[112, 55.4, 5.65, 2.57], [33, 48.5, 3.76, 1.25], None, None],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": [[1091, 60.3, 8.73, 5.24], [91, 34.1, 0.75, -1.78], [56, 53.6, 1.77, -0.98], None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [None, None, [40, 52.5, 4.02, 0.6], [43, 41.9, -0.04, 0.62]],
    "晨星剛形成 ＋ 近20日強於大盤≥10% ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[115, 54.8, 5.63, 2.37], [31, 51.6, 4.12, 0.92], None, None],
    "創52週新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": [[140, 55.0, 4.12, 1.86], [32, 43.8, -1.78, -0.94], None, None],
    "KDJ近3日內黃金交叉 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[544, 54.8, 3.7, 2.88], [50, 32.0, -2.0, -1.68], None, [32, 65.6, 13.56, 3.22]],
    "地量(≤0.5倍均量) ＋ 月營收連3月年增 ＋ 近3日漲停": [[377, 50.9, 3.02, 2.06], [143, 69.9, 13.71, 1.64], None, [218, 83.5, 4.97, 0.37]],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[940, 52.8, 7.08, 5.06], [106, 38.7, -0.1, -1.15], [85, 52.9, 4.02, 2.3], [44, 56.8, 9.76, 2.66]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[168, 48.2, 3.18, 2.17], [32, 75.0, 13.66, 1.67], None, None],
    "多方力道≥65 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1965, 54.2, 5.89, 5.25], [205, 35.1, -1.21, -1.82], [122, 44.3, 2.46, 1.56], [72, 50.0, 5.0, 2.01]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 短線過熱(高於月線≥20%)": [None, None, [43, 51.2, 3.4, 1.22], [44, 56.8, 4.71, 0.93]],
    "N字底剛形成 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[96, 52.1, 5.66, 2.01], None, None, None],
    "均線多頭排列(5>20>60且站上月線) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1959, 53.6, 6.07, 5.37], [207, 34.8, -1.01, -1.54], [112, 43.8, 2.75, 1.91], [73, 50.7, 4.91, 1.9]],
    "多方力道≥80 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1626, 53.0, 5.78, 4.21], [183, 34.4, -1.67, -2.6], [89, 36.0, 1.08, 1.22], [67, 47.8, 2.89, 1.33]],
    "晨星剛形成 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[136, 52.9, 5.9, 2.46], [30, 50.0, 2.66, 1.02], None, None],
    "月營收創24個月新高 ＋ 爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[972, 52.3, 6.16, 4.72], [110, 39.1, 1.22, -0.33], [84, 52.4, 4.43, 1.99], [57, 59.6, 10.19, 3.25]],
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[536, 47.4, 2.98, 2.54], [78, 35.9, -0.5, 0.49], None, None],
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 近3日漲停": [[766, 49.3, 3.78, 2.04], None, [46, 63.0, 7.5, 1.95], None],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 月營收創24個月新高": [[149, 52.3, 3.7, 2.11], None, None, None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1714, 54.1, 6.04, 4.64], [186, 35.5, -1.42, -2.55], [97, 39.2, 2.12, 1.6], [68, 48.5, 3.08, 1.42]],
    "N字底剛形成 ＋ 外資近5日買超 ＋ 短線過熱(高於月線≥20%)": [[172, 47.1, 5.3, 1.87], None, None, None],
    "MACD近3日內黃金交叉 ＋ OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%)": [[497, 46.9, 2.71, 2.87], [65, 49.2, 1.57, -1.13], [36, 50.0, 2.73, 0.97], [30, 50.0, 1.14, 0.16]],
    "MACD近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": [[366, 43.4, 1.78, 1.87], [59, 49.2, 1.65, -0.4], [32, 50.0, 2.88, 1.0], [35, 48.6, 0.72, -0.12]],
    "創52週新高 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[85, 50.6, 5.08, 2.11], None, None, None],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[79, 49.4, 6.83, 2.44], None, None, None],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[1348, 55.8, 7.15, 5.51], [150, 38.7, -0.21, -0.46], [70, 55.7, 6.89, 1.81], [53, 58.5, 4.27, 1.43]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1537, 53.9, 6.5, 5.64], [155, 30.3, -0.92, -3.09], [94, 46.8, 4.62, 2.48], [61, 47.5, 2.57, 1.12]],
    "MACD近3日內黃金交叉 ＋ 月營收連3月年增 ＋ 短線過熱(高於月線≥20%)": [[314, 47.1, 2.09, 1.34], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近20日強於大盤≥10%": [[84, 48.8, 4.05, 1.92], None, None, None],
    "MACD近3日內黃金交叉 ＋ 多方力道≥65 ＋ 短線過熱(高於月線≥20%)": [[469, 51.0, 2.93, 2.26], [62, 48.4, 0.49, -0.53], [36, 50.0, 2.74, 1.12], [32, 62.5, 7.78, 1.2]],
    "布林通道高檔(≥80%) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[2061, 55.1, 6.16, 6.2], [191, 35.6, -0.25, -1.62], [141, 51.8, 3.9, 2.29], [72, 54.2, 6.87, 2.5]],
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1243, 48.3, 3.27, 3.66], [223, 57.4, 5.32, 0.8], [201, 56.2, 6.1, 2.92], [118, 50.0, 4.54, 3.4]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": [[223, 52.0, 5.12, 3.47], [47, 51.1, 2.82, 0.17], None, None],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間高檔(≥90百分位)": [[124, 38.7, -1.83, -0.38], [33, 48.5, 1.13, 0.37], None, None],
    "多方力道≥80 ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": [[212, 52.4, 4.95, 2.92], [46, 50.0, 1.17, -0.49], None, None],
    "夜星剛形成 ＋ 月營收創24個月新高": [[144, 53.5, 4.44, 2.32], [40, 50.0, 4.49, 1.88], None, None],
    "三大法人近3月買超 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1769, 56.0, 6.73, 5.67], [188, 37.2, 0.67, -0.8], [101, 57.4, 8.24, 3.88], [69, 50.7, 5.95, 2.18]],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[1148, 52.7, 6.56, 5.65], [138, 38.4, -0.37, -1.19], [102, 53.9, 4.45, 2.05], [64, 54.7, 6.84, 2.31]],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10%": [[2070, 54.8, 6.08, 5.76], [210, 34.3, -1.18, -1.7], [143, 49.0, 3.49, 2.15], [76, 51.3, 5.35, 2.12]],
    "距52週高點≤5% ＋ 近3日漲停 ＋ 近3月均價YoY為負": [[1488, 48.6, 3.81, 4.64], [253, 54.9, 4.33, 1.23], [204, 56.9, 6.65, 2.86], [132, 50.8, 2.38, 1.24]],
    "母子懷抱(低檔)剛形成 ＋ 近20日強於大盤≥10% ＋ 近3日漲停": [[85, 44.7, 1.57, 0.47], None, None, None],
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間低檔(≤10百分位)": [[49, 77.6, 8.61, 2.84], None, None, None],
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[55, 72.7, 6.8, 2.57], None, None, None],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 月營收連3月年增": [None, [93, 59.1, 6.65, 2.18], None, [65, 49.2, 2.47, 0.2]],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近20日強於大盤≥10%": [[32, 53.1, 6.48, 1.55], None, None, None],
    "多方力道≥65 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[57, 70.2, 5.51, 2.09], None, None, None],
    "短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位)": [[58, 70.7, 5.67, 2.2], None, None, None],
    "晨星剛形成 ＋ 月營收創24個月新高 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[45, 48.9, 2.2, 0.37], None, None, None],
    "短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[59, 71.2, 5.89, 2.21], None, None, None],
    "K線橫盤的突破剛形成 ＋ 均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成": [[114, 47.4, 5.67, 2.38], None, None, None],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[2141, 54.8, 6.26, 5.81], [211, 34.6, -1.12, -1.71], [148, 50.7, 3.8, 2.18], [76, 51.3, 5.35, 2.12]],
    "地量(≤0.5倍均量) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成": [[62, 51.6, 4.91, 0.52], None, None, None],
    "月營收創24個月新高 ＋ 相對強弱為正(強於大盤) ＋ 短線過熱(高於月線≥20%)": [[2132, 54.8, 6.16, 5.79], [211, 34.6, -1.12, -1.71], [148, 50.7, 3.8, 2.18], [76, 51.3, 5.35, 2.12]],
    "近3日漲停 ＋ 量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[466, 47.6, 4.47, 3.06], [134, 47.0, 2.18, 1.2], [41, 34.1, -2.34, -0.4], [74, 51.4, 0.87, 0.23]],
    "創52週新高 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [[85, 50.6, 5.08, 2.11], None, None, None],
    "三大法人近3月買超 ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": [[76, 52.6, 6.12, 2.35], None, None, None],
    "創52週新高 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[67, 47.8, 4.64, 1.78], None, None, None],
    "月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成 ＋ 距52週高點≤5%": [[313, 54.6, 6.23, 4.1], [47, 48.9, 3.78, 0.73], [34, 50.0, 8.01, 1.67], None],
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[657, 48.9, 3.08, 3.52], [108, 49.1, 1.82, -0.72], [154, 47.4, -0.16, -0.5], [85, 51.8, 4.31, 1.93]],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 月營收創24個月新高": [[85, 51.8, 5.28, 2.36], None, None, None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": [[252, 52.4, 3.45, 2.88], [37, 51.4, 1.4, 0.37], None, None],
    "強勢突破盤 ＋ 月營收連3月年增 ＋ 母子懷抱(低檔)剛形成": [[89, 46.1, 3.78, 1.84], None, None, None],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 近3日漲停": [[920, 45.3, 2.96, 2.96], [186, 49.5, 1.76, 1.29], [80, 47.5, 2.86, 1.18], [73, 49.3, -1.96, 0.23]],
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": [[245, 51.8, 1.62, 1.35], [43, 44.2, -1.91, -0.29], None, None],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 近3日向上跳空缺口": [[829, 46.4, 2.56, 3.25], [134, 50.7, 3.11, 1.54], [64, 43.8, 1.02, -0.48], [38, 50.0, 0.12, 0.82]],
    "創52週新高 ＋ 近3日漲停 ＋ 近3月均價YoY為負": [[1419, 46.2, 2.74, 3.67], [267, 55.1, 4.75, 1.37], [200, 54.0, 6.01, 1.95], [140, 48.6, 2.44, 2.22]],
    "月營收創24個月新高 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": [[1877, 54.4, 6.58, 5.55], [184, 33.7, -0.62, -1.6], [130, 50.8, 3.56, 2.09], [69, 47.8, 2.34, 1.11]],
    "近3日向上跳空缺口 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[534, 50.2, 4.38, 3.22], [122, 59.8, 10.5, 2.73], [35, 57.1, 3.45, 0.42], [221, 86.4, 6.22, 0.9]],
    "均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[154, 49.4, 2.99, 1.57], None, None, None],
    "三大法人近3月買超 ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": [[274, 48.5, 3.46, 2.59], [49, 53.1, 1.97, 0.62], None, None],
    "相對強弱為負(弱於大盤) ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[228, 48.2, 1.22, 1.0], [56, 64.3, 11.76, 2.62], None, [238, 89.5, 6.11, 0.71]],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": [[383, 52.5, 6.45, 2.99], [43, 46.5, -0.41, 0.33], None, None],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 突破飆股大量黑K最高點剛形成": [[148, 45.9, 3.74, 1.57], None, None, None],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": [[624, 48.7, 3.19, 3.15], [88, 48.9, 1.52, -0.46], [40, 47.5, 3.11, 1.0], [43, 55.8, 3.16, 0.79]],
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[149, 48.3, 3.2, 1.45], [52, 61.5, 6.82, 1.88], None, None],
    "夜星剛形成 ＋ 月營收創24個月新高 ＋ 相對強弱為正(強於大盤)": [[90, 50.0, 4.62, 2.35], [32, 50.0, 3.28, 1.25], None, None],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": [[417, 57.1, 6.27, 3.6], [42, 50.0, 10.11, 1.19], None, None],
    "距52週高點≤5% ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[371, 54.2, 5.14, 3.54], [61, 59.0, 9.41, 2.16], None, None],
    "三大法人近3月買超 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[774, 48.8, 4.1, 3.49], [253, 57.3, 8.46, 2.54], [65, 43.1, -0.33, -0.02], [225, 73.8, 5.52, 1.77]],
    "創52週新高 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[85, 50.6, 5.08, 2.11], None, None, None],
    "MACD近3日內黃金交叉 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": [[537, 49.5, 3.89, 3.22], [74, 47.3, -1.44, -1.45], [34, 47.1, 2.77, 0.8], [37, 51.4, 1.39, 0.14]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3日向上跳空缺口": [[540, 47.2, 2.18, 3.18], [76, 47.4, 1.2, -0.39], [31, 48.4, 3.68, 0.8], [37, 54.1, 3.54, 0.66]],
    "回後買上漲全通過 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[640, 46.7, 2.94, 2.45], [101, 52.5, 0.7, -0.04], [129, 62.8, 5.69, 2.95], [67, 50.7, 6.21, 1.93]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": [[71, 50.7, 5.08, 1.78], None, None, None],
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[1571, 47.0, 3.03, 4.64], [283, 53.7, 3.11, 1.72], [300, 50.0, 3.3, 1.28], [186, 47.8, 5.02, 2.86]],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近3日向上跳空缺口": [[111, 46.8, 4.05, 1.19], None, None, None],
    "外資近5日買超 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[717, 48.3, 4.23, 3.31], [225, 58.7, 8.24, 2.74], [62, 38.7, -1.02, 0.36], [284, 79.6, 5.87, 2.09]],
    "外資近5日買超 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [None, None, None, [68, 52.9, 5.03, 0.99]],
    "KDJ近3日內死亡交叉 ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[715, 51.5, 3.9, 3.51], [103, 44.7, -0.63, -1.56], [61, 49.2, 2.43, 1.61], [65, 41.5, -2.02, -0.1]],
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1573, 45.3, 2.09, 2.29], [262, 50.8, 0.94, -0.87], [291, 49.8, 3.72, 2.06], [158, 50.6, 3.68, 1.6]],
    "創52週新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近20日強於大盤≥10%": [[461, 49.7, 3.12, 3.12], [115, 53.0, 4.62, 1.14], [61, 50.8, 8.27, 2.27], [48, 41.7, -1.05, -0.52]],
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1378, 45.3, 1.99, 2.66], [246, 52.0, 2.25, 0.61], [278, 50.7, 3.5, 1.51], [186, 47.3, 3.85, 1.89]],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[50, 52.0, 8.04, 2.37], None, None, None],
    "外資近5日買超 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1460, 54.9, 6.95, 5.83], [151, 36.4, 0.25, -1.71], [101, 52.5, 4.95, 2.71], [62, 46.8, 4.86, 1.79]],
    "創52週新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[323, 49.5, 4.0, 2.93], [59, 57.6, 7.47, 1.74], None, None],
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 短線過熱(高於月線≥20%)": [[654, 48.3, 3.01, 3.12], None, [43, 51.2, 3.4, 1.22], None],
    "MACD近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 短線過熱(高於月線≥20%)": [[641, 48.2, 2.83, 3.02], [95, 49.5, 1.03, -0.66], [43, 51.2, 3.4, 1.22], [44, 56.8, 4.71, 0.93]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10%": [[583, 49.4, 2.95, 2.9], [94, 50.0, 1.3, -0.61], [42, 50.0, 2.78, 1.14], [44, 56.8, 4.71, 0.93]],
    "創52週新高 ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": [[2037, 47.9, 3.56, 6.71], [377, 54.4, 4.23, 1.97], [262, 54.6, 4.93, 1.94], [174, 46.0, 1.2, 1.17]],
    "KDJ近3日內死亡交叉 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": [[381, 47.5, 3.63, 2.82], [79, 43.0, -0.95, -0.91], [59, 55.9, 3.67, 1.34], [58, 43.1, 0.01, 0.27]],
    "創52週新高 ＋ 多方力道≥80 ＋ 突破ABC修正下降切線剛形成": [[372, 49.2, 3.73, 3.07], [57, 42.1, 1.62, -0.02], [37, 54.1, 10.7, 2.08], None],
    "母子懷抱(低檔)剛形成 ＋ 近20日強於大盤≥10% ＋ 近3日向上跳空缺口": [[60, 48.3, 3.69, -0.19], None, None, None],
    "三大法人近3月賣超 ＋ 地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[139, 49.6, 7.13, 2.25], None, [31, 45.2, -3.47, -2.09], None],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[107, 48.6, 6.04, 2.4], None, None, None],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[452, 46.9, 3.77, 2.4], [73, 47.9, -0.65, -1.13], [35, 54.3, 3.35, 0.86], [33, 48.5, -0.2, 0.07]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": [[654, 48.3, 3.01, 3.12], [96, 49.0, 0.67, -0.67], [43, 51.2, 3.4, 1.22], [44, 56.8, 4.71, 0.93]],
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%)": [[653, 48.2, 3.02, 3.14], [96, 49.0, 0.67, -0.67], [43, 51.2, 3.4, 1.22], [44, 56.8, 4.71, 0.93]],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 月營收連3月年增": [[130, 46.2, 2.07, 1.48], None, None, None],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[292, 58.9, 6.76, 4.06], None, None, None],
    "短線過熱(高於月線≥20%) ＋ 距52週高點≤5% ＋ 近3月均價YoY為負": [[1381, 48.4, 3.78, 3.84], [236, 57.2, 4.86, 0.83], [220, 57.7, 7.98, 3.52], [121, 52.1, 5.05, 3.5]],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間高檔(≥90百分位)": [[1208, 50.4, 5.19, 4.53], [138, 37.0, -1.3, -1.44], [77, 50.6, 3.96, 2.21], [65, 49.2, 6.03, 1.93]],
    "月營收連3月年增 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[69, 56.5, 7.37, 2.57], None, None, None],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 短線過熱(高於月線≥20%)": [[734, 45.9, 3.27, 2.23], [126, 49.2, 2.21, 1.26], [69, 47.8, 1.45, 0.61], [49, 44.9, -1.46, 0.1]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[2144, 47.0, 2.53, 4.15], [398, 49.7, 2.16, 0.62], [405, 54.6, 5.42, 2.37], [207, 44.9, 3.65, 2.33]],
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1651, 45.2, 2.32, 2.42], [281, 53.0, 2.22, 0.52], [301, 53.5, 4.66, 2.35], [182, 51.1, 4.14, 1.29]],
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[2031, 45.3, 1.81, 2.0], [370, 50.5, 2.52, 1.06], [395, 52.2, 3.81, 1.61], [254, 44.9, 2.84, 1.66]],
    "K線橫盤的突破剛形成 ＋ N字底剛形成 ＋ 近3日漲停": [[87, 46.0, 4.06, 1.75], None, None, None],
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[352, 48.9, 2.6, 1.64], [114, 64.9, 13.63, 2.75], None, [140, 77.9, 4.38, 0.64]],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 短線過熱(高於月線≥20%)": [[539, 48.1, 2.36, 2.86], [67, 44.8, 0.61, -0.48], [33, 48.5, 2.3, 0.79], [31, 54.8, 4.82, 0.99]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": [[453, 47.9, 3.63, 3.09], [64, 48.4, 1.54, 0.22], None, [30, 50.0, -0.61, 0.16]],
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月均價YoY為負": [[1307, 49.1, 3.34, 5.05], [213, 59.2, 7.27, 1.55], [179, 47.5, 2.63, 1.61], [88, 44.3, 1.91, 1.44]],
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 量能區間高檔(≥90百分位)": [[2346, 44.9, 1.89, 2.85], [431, 50.3, 2.3, 1.17], [417, 49.9, 2.88, 1.54], [273, 45.1, 2.5, 1.85]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 創52週新高 ＋ 地量(≤0.5倍均量)": [[426, 57.3, 6.23, 3.5], [44, 50.0, 9.3, 1.11], None, None],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": [[639, 51.5, 4.67, 2.91], [74, 41.9, -0.8, 0.47], None, None],
    "地量(≤0.5倍均量) ＋ 近3日向上跳空缺口 ＋ 近3日漲停": [[484, 53.9, 6.52, 3.59], [106, 70.8, 12.94, 2.66], None, [261, 90.8, 4.38, 1.66]],
    "近3日漲停 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[819, 47.5, 2.82, 2.41], [222, 53.2, 8.5, 2.38], [66, 50.0, 1.67, 0.5], [266, 81.2, 6.26, 0.91]],
    "大盤跌破20日均線 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [None, [338, 56.2, 7.04, 2.74], None, [399, 76.7, 4.15, 1.03]],
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[1786, 46.9, 3.14, 5.71], [298, 55.0, 5.4, 2.73], [259, 52.5, 4.26, 2.61], [150, 44.7, 3.43, 2.81]],
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 短線過熱(高於月線≥20%)": [[326, 45.1, 2.43, 2.26], [55, 47.3, -2.17, -0.39], None, None],
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[2605, 46.7, 2.45, 2.93], [458, 53.1, 3.36, 0.85], [489, 51.5, 3.72, 1.36], [283, 45.2, 2.43, 1.58]],
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1724, 46.3, 2.2, 2.46], [303, 50.5, 1.92, 0.42], [325, 49.8, 3.18, 1.54], [214, 46.3, 3.54, 2.26]],
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[1770, 45.6, 1.78, 1.93], [319, 49.8, 2.23, 0.61], [336, 51.2, 3.73, 1.66], [213, 48.8, 4.72, 2.84]],
    "K線橫盤的突破剛形成 ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": [[381, 48.0, 3.96, 3.53], [83, 51.8, 0.89, 0.06], [80, 52.5, 6.47, 1.62], [52, 44.2, -1.12, 0.23]],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近20日強於大盤≥10%": [[52, 51.9, 5.83, 0.87], None, None, None],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 距52週高點≤5%": [[81, 54.3, 7.38, 2.9], None, None, None],
    "創52週新高 ＋ 母子懷抱(低檔)剛形成 ＋ 近20日強於大盤≥10%": [[65, 49.2, 6.05, 2.08], None, None, None],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[545, 47.9, 3.52, 1.9], [40, 22.5, -2.67, -1.69], [47, 48.9, 10.18, 1.49], None],
    "K線橫盤的突破剛形成 ＋ 均線多頭排列(5>20>60且站上月線) ＋ 母子懷抱(低檔)剛形成": [[54, 55.6, 5.58, 1.96], None, None, None],
    "創52週新高 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[81, 51.9, 5.99, 2.34], None, None, None],
    "創52週新高 ＋ 母子懷抱(低檔)剛形成": [[85, 50.6, 5.08, 2.11], None, None, None],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[167, 50.3, 3.9, 2.17], None, None, None],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[108, 43.5, 2.32, 0.95], None, None, None],
    "母子懷抱(低檔)剛形成 ＋ 距52週高點≤5% ＋ 近3日漲停": [[65, 41.5, 1.61, 0.68], None, None, None],
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 月營收創24個月新高": [[94, 56.4, 4.96, 1.74], None, None, None],
    "近3日漲停 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[678, 49.1, 3.31, 3.24], [192, 54.7, 9.26, 2.48], [46, 56.5, 3.44, 1.14], [204, 80.9, 4.86, -0.03]],
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[828, 48.7, 3.14, 3.34], [139, 59.0, 7.9, 2.87], [146, 50.7, 0.97, -0.01], [96, 52.1, 5.91, 2.69]],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 爆量(≥2倍均量)": [[1043, 51.0, 6.0, 4.53], [140, 40.0, -1.88, -0.57], [82, 50.0, 6.3, 2.76], [91, 54.9, 4.21, 1.32]],
    "三大法人近3月買超 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[606, 50.5, 4.51, 4.11], [123, 63.4, 8.1, 2.27], [54, 46.3, 5.2, 1.32], [45, 40.0, -0.17, -0.16]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 創52週新高 ＋ 突破ABC修正下降切線剛形成": [[429, 51.3, 4.31, 3.98], [69, 46.4, 3.99, 1.15], [39, 51.3, 9.5, 1.94], None],
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[2059, 47.4, 3.49, 5.69], [333, 53.2, 2.42, 1.22], [345, 51.6, 3.54, 2.01], [191, 46.6, 4.16, 2.24]],
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": [[277, 55.2, 5.37, 3.56], [42, 57.1, 5.55, 1.31], [40, 50.0, 5.63, 1.83], None],
    "大盤跌破60日均線 ＋ 月營收創24個月新高 ＋ 近3日漲停": [None, None, [283, 57.2, 4.74, 2.37], [701, 76.5, 4.58, 1.35]],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": [[218, 59.2, 7.3, 3.27], None, None, None],
    "短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成 ＋ 距52週高點≤5%": [[262, 48.5, 3.13, 1.71], [44, 52.3, 3.33, 0.64], None, None],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[64, 46.9, 3.23, 0.98], None, None, None],
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[1995, 44.6, 2.67, 2.96], [347, 46.1, 1.3, 0.46], [195, 56.9, 5.07, 2.27], [157, 49.0, 3.87, 2.37]],
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[1935, 45.3, 2.17, 3.94], [300, 48.3, 1.57, 0.94], [334, 47.0, 2.4, 0.66], [186, 47.3, 3.82, 1.99]],
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[2365, 47.1, 2.93, 5.25], [437, 52.2, 3.36, 2.35], [442, 50.0, 3.13, 1.16], [270, 44.1, 2.44, 2.22]],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 近20日強於大盤≥10%": [[198, 42.4, 0.77, 1.04], [61, 55.7, 5.38, 1.39], None, [63, 55.6, 4.88, 1.92]],
    "OBV能量潮創60日新高 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[1161, 55.6, 7.73, 6.87], [146, 36.3, -1.98, -1.11], [107, 50.5, 3.56, 1.46], [75, 54.7, 3.03, 0.55]],
    "晨星剛形成 ＋ 距52週高點≤5% ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[134, 47.0, 3.95, 2.21], None, None, None],
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[2415, 45.7, 2.26, 2.65], [454, 53.5, 3.99, 1.37], [455, 51.4, 4.13, 2.56], [272, 44.5, 2.09, 2.0]],
    "距52週高點≤5% ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": [[2116, 49.4, 4.2, 6.77], [386, 53.9, 3.45, 1.34], [278, 55.0, 4.92, 2.4], [176, 46.6, 0.84, 1.14]],
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[2013, 47.6, 3.05, 5.41], [355, 52.1, 2.9, 1.99], [364, 48.6, 2.52, 1.11], [224, 46.9, 3.99, 2.6]],
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[700, 46.7, 2.15, 2.6], [139, 36.0, -0.85, -1.09], [124, 51.6, 0.92, -0.25], [90, 48.9, 6.3, 2.82]],
    "月營收創24個月新高 ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": [[1409, 50.2, 4.25, 4.92], [287, 50.5, 1.47, -0.08], [188, 55.3, 5.61, 2.48], [421, 75.3, 4.23, 1.44]],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 回後買上漲全通過": [[480, 50.4, 1.57, 1.37], [78, 46.2, -0.52, -0.11], [36, 61.1, 7.51, 1.82], None],
    "大盤跌破60日均線 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": [None, None, [459, 52.3, 2.26, 1.66], [315, 45.1, 2.13, 1.63]],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 短線過熱(高於月線≥20%)": [[229, 48.9, 5.62, 2.27], None, None, None],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": [[381, 56.2, 6.03, 3.74], [30, 50.0, -0.88, -0.39], [38, 52.6, 9.99, 2.13], [31, 35.5, -0.08, -0.23]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 多方力道≥80 ＋ 晨星剛形成": [[199, 48.7, 3.94, 2.14], None, None, None],
    "外資近5日買超 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1655, 48.6, 3.98, 4.7], [311, 58.8, 5.81, 1.63], [311, 55.6, 5.56, 1.19], [157, 47.1, 3.27, 2.28]],
    "大盤站上20日均線 ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[4832, 48.4, 3.38, 4.49], None, [422, 53.6, 3.81, 2.35], None],
    "短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10% ＋ 近3月均價YoY為負": [[2643, 46.4, 2.47, 2.97], [492, 53.0, 3.39, 1.13], [500, 51.6, 3.74, 1.49], [292, 44.5, 2.1, 1.6]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間高檔(≥90百分位)": [[492, 45.5, 1.94, 2.06], [76, 44.7, -1.1, -1.31], [37, 51.4, 3.76, 1.38], [35, 51.4, 2.67, -0.13]],
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 突破飆股大量黑K最高點剛形成": [[304, 49.0, 5.0, 3.21], [49, 55.1, 2.71, -0.39], None, None],
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[402, 52.0, 4.46, 2.19], [90, 55.6, 4.35, 0.76], [44, 45.5, 3.15, 0.84], [39, 43.6, 1.16, 0.76]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 夜星剛形成 ＋ 月營收創24個月新高": [[76, 47.4, 3.19, 1.55], None, None, None],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[44, 59.1, 13.21, 3.12], None, None, None],
}


def CONSTANTS():
    return dict(TWSE_LIST=TWSE_LIST, TPEX_LIST=TPEX_LIST, MY_LIST_DEFAULT=MY_LIST_DEFAULT, TW0050_LIST=TW0050_LIST,
                PINNED_COMBOS=PINNED_COMBOS, PINNED_COMBO_WINRATES=PINNED_COMBO_WINRATES,
                MOONSHOT_COMBOS=MOONSHOT_COMBOS, MOONSHOT_COMBO_STATS=MOONSHOT_COMBO_STATS,
                STOCK_PICK_COMBOS=STOCK_PICK_COMBOS, STOCK_PICK_STATS=STOCK_PICK_STATS, COMBO_REF_STATS=COMBO_REF_STATS,
                BT_WIN_COMBOS=BT_WIN_COMBOS, BT_WIN_STATS=BT_WIN_STATS, BT_HOT_COMBOS=BT_HOT_COMBOS, BT_HOT_STATS=BT_HOT_STATS,
                REGIME_STATS=REGIME_STATS)


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == 'daily':
        sys.exit(run_daily(sys.argv[2:]))
    main()
