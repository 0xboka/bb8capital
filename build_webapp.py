#!/usr/bin/env python3
"""Build BB8 portfolio dashboard as self-contained HTML webapp."""

import pandas as pd
import numpy as np
import json
import requests
import time
import pathlib
from collections import defaultdict
from copy import deepcopy

xl23 = pd.ExcelFile("bb8_23_dashboard_v1.xlsx")
xl24 = pd.ExcelFile("bb8_24_dashboard_v1.xlsx")
xl25 = pd.ExcelFile("bb8_25_dashboard_v2.xlsx")
xl   = xl25  # alias for _wallet / _data parsing

MONTHS_23 = ['2309','2310','2311','2312']
MONTHS_24 = ['2401','2402','2403','2404','2405','2406','2407','2408','2409','2410','2411']
MONTHS_25 = ['2412','2501','2502','2503','2504','2505','2506','2507','2508']
MONTHS    = MONTHS_23 + MONTHS_24 + MONTHS_25

MONTH_LABELS = {
    '2309':'Sep 23','2310':'Oct 23','2311':'Nov 23','2312':'Dec 23',
    '2401':'Jan 24','2402':'Feb 24','2403':'Mar 24','2404':'Apr 24','2405':'May 24',
    '2406':'Jun 24','2407':'Jul 24','2408':'Aug 24','2409':'Sep 24','2410':'Oct 24',
    '2411':'Nov 24',
    '2412':'Dec 24','2501':'Jan 25','2502':'Feb 25','2503':'Mar 25',
    '2504':'Apr 25','2505':'May 25','2506':'Jun 25','2507':'Jul 25','2508':'Aug 25',
}

# Map each month key → (ExcelFile, sheet_name)
SHEET_SOURCE = {}
for _m in MONTHS_23: SHEET_SOURCE[_m] = (xl23, _m + '_')   # 2309_ etc.
for _m in MONTHS_24: SHEET_SOURCE[_m] = (xl24, _m)
for _m in MONTHS_25: SHEET_SOURCE[_m] = (xl25, _m)

# Format A: 3 token cols (wallet|token|tok_simple), then prev cols 4-8, actual cols 9-13
FORMAT_A_MONTHS = set(MONTHS_23 + ['2401','2402','2403','2404','2405'])

# ── Synthetic month config ────────────────────────────────────────────────────
SYNTHETIC_MONTHS = ['2509','2510','2511']
SYNTHETIC_MONTH_LABELS = {
    '2509':'Sep 25 (est.)','2510':'Oct 25 (est.)',
    '2511':'Nov 25 (est.)'
}
SYNTHETIC_DATES = {
    '2509':'01-09-2025','2510':'01-10-2025',
    '2511':'01-11-2025'
}

# ── Stablecoins: always $1.00 ─────────────────────────────────────────────────
STABLECOIN_TOKENS = frozenset({
    'USDC','USDC.e','USDCe','USDT','USDT0','USD',
    'FRAX','xDAI','crvUSD','GHO','GRAI','scUSD','stkscUSD','DAI',
})

# ── Token → CoinGecko ID mapping ─────────────────────────────────────────────
TOKEN_COINGECKO_ID = {
    # ETH + derivatives
    'ETH':                  'ethereum',
    'WETH':                 'ethereum',
    'wstETH':               'wrapped-steth',
    'cbETH':                'coinbase-wrapped-staked-eth',
    'weETH':                'wrapped-eeth',
    'ezETH':                'renzo-restaked-eth',
    'scETH':                'ethereum',
    'PT-wstkscETH-26JUN2025': 'ethereum',
    # BTC
    'cbBTC':                'coinbase-wrapped-btc',
    'uBTC':                 'bitcoin',
    # HYPE
    'HYPE':                 'hyperliquid',
    'WHYPE':                'hyperliquid',
    'wstHYPE':              None,           # no CoinGecko ID → proxy HYPE
    # Sonic
    'S':                    'sonic-3',
    'stS':                  'sonic-3',
    'wS':                   'sonic-3',
    # Pendle
    'PENDLE':               'pendle',
    'mPENDLE':              'pendle',
    'PT-ENA-25APR2025':     'ethena',
    'PT-sENA':              'ethena',
    'PT-sENA-25SEP2025':    'ethena',
    # ENA
    'ENA':                  'ethena',
    'sENA':                 'ethena',
    # MAV
    'MAV':                  'maverick-protocol',
    'sdMAV':                'maverick-protocol',
    # FRAX ecosystem
    'WFRAX':                'frax',
    'FXS':                  'frax-share',
    'FXN':                  'fxn',
    # DeFi / L1 / L2
    'AAVE':                 'aave',
    'ARB':                  'arbitrum',
    'ATOM':                 'cosmos',
    'AVAX':                 'avalanche-2',
    'AXL':                  'axelar',
    'BNB':                  'binancecoin',
    'CELO':                 'celo',
    'CLEV':                 'clever',
    'CVX':                  'convex-finance',
    'EIGEN':                'eigenlayer',
    'ETHFI':                'ether-fi',
    'FLUID':                'instadapp',
    'FTM':                  'fantom',
    'GEAR':                 'gearbox-protocol',
    'GHST':                 'aavegotchi',
    'GLMR':                 'moonbeam',
    'GNO':                  'gnosis',
    'JUNO':                 'juno-network',
    'LDO':                  'lido-dao',
    'LQTY':                 'liquity',
    'METIS':                'metis-token',
    'MIMO':                 'mimo-parallel-governance-token',
    'MODE':                 'mode',
    'MORPHO':               'morpho',
    'ONDO':                 'ondo-finance',
    'ONE':                  'harmony',
    'OP':                   'optimism',
    'OSMO':                 'osmosis',
    'POL':                  'matic-network',
    'REZ':                  'renzo-protocol',
    'SCRT':                 'secret',
    'SHADOW':               'shadow-exchange',
    'xSHADOW':              'shadow-exchange',
    'SILO':                 'silo-finance',
    'STG':                  'stargate-finance',
    'TIA':                  'celestia',
    'UNI':                  'uniswap',
    'VIRTUAL':              'virtual-protocol',
    'AIXBT':                'aixbt-by-virtuals',
    'GNS':                  'gains-network',
    # No CoinGecko ID → None (roll forward)
    'stkscETH':             None,
    'PRL':                  None,
    'NFT':                  None,
}

# Proxy: when TOKEN_COINGECKO_ID is None, use another token's price
TOKEN_PRICE_PROXY = {
    'wstHYPE':  'HYPE',
    'stkscETH': 'ETH',
    'scETH':    'ETH',
}

# Non-EVM chains for KPI calculation
NON_EVM_CHAINS = frozenset({
    'cosmos','osmosis','axelar','secret network','juno',
    'celestia','stargaze','hyperliquid',
})

CACHE_FILE = pathlib.Path(__file__).parent / 'coingecko_price_cache.json'

# ─────────────────────────────────────────────────────────────────────────────

def safe_float(v):
    if v is None: return None
    if isinstance(v, str):
        if v.strip() in ('+ ∞', '+∞', '∞'): return 9999.0
        try: return float(v)
        except: return None
    if isinstance(v, float) and np.isnan(v): return None
    try: return float(v)
    except: return None

def safe_str(v):
    if v is None: return ''
    if isinstance(v, float) and np.isnan(v): return ''
    return str(v).strip()

def extract_kpis_from_df(df, start, prev, actual):
    """Keyword-based KPI extraction — works across all sheet format versions."""
    diff_val = actual - start
    diff_pct = diff_val / start if start else 0
    kpi = {
        'start': start, 'prev': prev, 'actual': actual,
        'diff_val': round(diff_val, 0), 'diff_pct': round(diff_pct * 100, 2),
        'hf_min': None, 'hf_global': None, 'debt': None,
        'apy_p': None, 'apy_c': None, 'ni_p': None, 'ni_c': None,
        'ne_p': None, 'ne_c': None, 'alt_p': None, 'alt_c': None,
    }
    hf_count = 0
    for i in range(4, 44):
        if i >= len(df): break
        r = df.iloc[i]
        label = str(r.get(1) or '').lower()
        vp = safe_float(r.get(2))
        vc = safe_float(r.get(3)) or safe_float(r.get(4))
        if 'health factor' in label:
            hf_count += 1
            val = vc if vc is not None else vp
            if hf_count == 1: kpi['hf_min']    = val
            else:             kpi['hf_global'] = val
        elif 'debt' in label:
            kpi['debt'] = vc if vc is not None else vp
        elif 'apy' in label or 'pondér' in label:
            kpi['apy_p'] = vp; kpi['apy_c'] = vc
        elif 'non investi' in label:
            kpi['ni_p'] = vp; kpi['ni_c'] = vc
        elif 'non evm' in label:
            kpi['ne_p'] = vp; kpi['ne_c'] = vc
        elif 'altcoin' in label:
            kpi['alt_p'] = vp; kpi['alt_c'] = vc
    return kpi

# ── CoinGecko price cache ─────────────────────────────────────────────────────

def load_price_cache():
    if CACHE_FILE.exists():
        with open(CACHE_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_price_cache(cache):
    with open(CACHE_FILE, 'w') as f:
        json.dump(cache, f, indent=2)

def fetch_coingecko_price(cg_id, date_str, cache, retries=3):
    """Fetch historical USD price. date_str = 'dd-mm-yyyy'. Uses/updates cache."""
    if cg_id not in cache:
        cache[cg_id] = {}
    if date_str in cache[cg_id]:
        cached = cache[cg_id][date_str]
        print(f'  [cache] {cg_id} @ {date_str} = {cached}')
        return cached

    url = f'https://api.coingecko.com/api/v3/coins/{cg_id}/history'
    params = {'date': date_str, 'localization': 'false'}

    for attempt in range(retries):
        try:
            print(f'  [CoinGecko] {cg_id} @ {date_str} (attempt {attempt+1})')
            resp = requests.get(url, params=params, timeout=15)
            if resp.status_code == 429:
                wait = 60 * (attempt + 1)
                print(f'  [rate-limit] waiting {wait}s...')
                time.sleep(wait)
                continue
            if resp.status_code == 200:
                data = resp.json()
                price = (data.get('market_data') or {}).get('current_price', {}).get('usd')
                if price is not None:
                    cache[cg_id][date_str] = float(price)
                    save_price_cache(cache)
                    time.sleep(2.5)   # ~24 req/min
                    return float(price)
                else:
                    print(f'  [warn] no market_data for {cg_id} @ {date_str}')
                    cache[cg_id][date_str] = None
                    save_price_cache(cache)
                    return None
            print(f'  [warn] HTTP {resp.status_code} for {cg_id}')
        except requests.RequestException as e:
            print(f'  [warn] request error: {e}')
            time.sleep(5 * (2 ** attempt))
    print(f'  [error] failed {cg_id} @ {date_str} after {retries} attempts')
    return None

def get_price_for_position(token_name, tok_simple, date_str, cache):
    """Return (price_usd, source). Handles stablecoins, proxies, and unknowns."""
    # Stablecoins
    if token_name in STABLECOIN_TOKENS or tok_simple in STABLECOIN_TOKENS:
        return 1.0, 'stable'
    # NFT
    if tok_simple == 'NFT' or token_name == 'NFT':
        return None, 'nft'
    # CoinGecko lookup by full token name
    cg_id = TOKEN_COINGECKO_ID.get(token_name)
    # Proxy fallback
    if cg_id is None and token_name in TOKEN_PRICE_PROXY:
        proxy = TOKEN_PRICE_PROXY[token_name]
        cg_id = TOKEN_COINGECKO_ID.get(proxy)
        if cg_id:
            print(f'  [proxy] {token_name} → {proxy} ({cg_id})')
    if cg_id is None:
        print(f'  [rollforward] no ID for {token_name!r} (simple={tok_simple})')
        return None, 'rollforward'
    price = fetch_coingecko_price(cg_id, date_str, cache)
    return price, ('coingecko' if price is not None else 'rollforward')

# ── KPI computation for synthetic months ──────────────────────────────────────

def compute_month_kpis(positions, prev_actual, start=20_000_000):
    """Compute KPI dict from positions list."""
    pos_vals = [p['val_a'] for p in positions if (p['val_a'] or 0) > 0]
    neg_vals = [p['val_a'] for p in positions if (p['val_a'] or 0) < 0]
    total_pos = sum(pos_vals) if pos_vals else 0
    total_neg = sum(neg_vals) if neg_vals else 0
    actual = total_pos + total_neg
    debt   = total_neg

    apy_num = sum((p['val_a'] or 0) * (p['apy_a'] or 0)
                  for p in positions if (p['val_a'] or 0) > 0)
    w_apy = apy_num / total_pos if total_pos else 0

    non_inv = sum((p['val_a'] or 0) for p in positions
                  if p['method'] == 'wallet' and (p['val_a'] or 0) > 0)

    non_evm_val = sum((p['val_a'] or 0) for p in positions
                      if p.get('chain','').lower() in NON_EVM_CHAINS
                      and (p['val_a'] or 0) > 0)
    non_evm_pct = non_evm_val / total_pos if total_pos else 0

    alt_val = sum((p['val_a'] or 0) for p in positions
                  if (p['val_a'] or 0) > 0
                  and token_cat.get(p['tok_simple'], 'other') not in ('infra','stable','bitcoin'))
    alt_pct = alt_val / total_pos if total_pos else 0

    diff_val = actual - start
    diff_pct = diff_val / start if start else 0
    return {
        'start': start, 'prev': prev_actual, 'actual': round(actual, 2),
        'diff_val': round(diff_val, 0), 'diff_pct': round(diff_pct * 100, 2),
        'hf_min':  None, 'hf_global': None,
        'debt':    round(debt, 0),
        'apy_p':   None, 'apy_c':  round(w_apy, 6),
        'ni_p':    None, 'ni_c':   round(non_inv, 2),
        'ne_p':    None, 'ne_c':   round(non_evm_pct, 6),
        'alt_p':   None, 'alt_c':  round(alt_pct, 6),
    }

# ── top_n: module-level (moved out of per-month loop) ─────────────────────────

def top_n(d, n=20):
    items = sorted(d.items(), key=lambda x: -x[1])
    if len(items) > n:
        other = sum(v for _, v in items[n:])
        items = items[:n]
        if other > 0: items.append(('Other', other))
    return [{'label': k, 'value': round(v, 2)} for k, v in items if v > 0]

# ── Token → category mapping from _wallet ────────────────────────────────────
wallet_raw = xl.parse('_wallet', header=None)
token_cat, token_sub = {}, {}
for i in range(len(wallet_raw)):
    r = wallet_raw.iloc[i]
    tok = r.get(1); cat = r.get(2); sub = r.get(3)
    if pd.notna(tok) and isinstance(tok, str) and tok.strip() not in ('tokens',''):
        t = tok.strip().upper()
        token_cat[t] = safe_str(cat) if pd.notna(cat) else 'other'
        token_sub[t] = safe_str(sub) if pd.notna(sub) else ''

extra_map = {
    'WFRAX':'defi','CRVUSD':'stable','USDC':'stable','USDT':'stable',
    'USDT0':'stable','XDAI':'stable','FRAX':'stable','DAI':'stable',
    'USD':'stable',
    'WSTETH':'infra','CBETH':'infra','WETH':'infra','STETH':'infra',
    'MPENDLE':'defi','STHYPE':'perp','WHYPE':'perp',
    'SDMAV':'defi','PT-SENA-25SEP2025':'defi',
    'MORPHO':'defi','GNO':'infra','ETHFI':'defi','ENA':'defi',
    'S':'infra','POL':'infra','BNB':'infra','AVAX':'infra',
    'GLMR':'infra','CELO':'infra','ONE':'infra','METIS':'infra',
}
for k, v in extra_map.items():
    if k not in token_cat: token_cat[k] = v

# ── Wallet targets ────────────────────────────────────────────────────────────
wallet_targets = []
for i in range(3, len(wallet_raw)):
    r = wallet_raw.iloc[i]
    tok = r.get(1)
    if not pd.notna(tok) or not isinstance(tok, str): continue
    tok = tok.strip()
    if tok in ('tokens',''): continue
    wallet_targets.append({
        'token': tok,
        'category': safe_str(r.get(2)),
        'sub': safe_str(r.get(3)),
        'min_upside': safe_float(r.get(5)),
        'target_min': safe_float(r.get(6)),
        'target_max': safe_float(r.get(7)),
        't0225': safe_float(r.get(11)), 'c0225': safe_float(r.get(12)),
        't0625': safe_float(r.get(15)), 'c0625': safe_float(r.get(16)),
        't0725': safe_float(r.get(19)), 'c0725': safe_float(r.get(20)),
        'comment': safe_str(r.get(21)),
    })

# ── Per-month data (real months) ─────────────────────────────────────────────
all_data = {}
for sheet_name in MONTHS:
    xl_file, sheet_raw = SHEET_SOURCE[sheet_name]
    df = xl_file.parse(sheet_raw, header=None)

    r2     = df.iloc[2]
    start  = safe_float(r2.get(1)) or 20_000_000
    prev   = safe_float(r2.get(2)) or 0
    actual = safe_float(r2.get(3)) or safe_float(r2.get(4)) or 0

    kpi = extract_kpis_from_df(df, start, prev, actual)

    # Column layout depends on format
    is_fmt_a = sheet_name in FORMAT_A_MONTHS
    has_sell = not is_fmt_a and sheet_name in set(MONTHS_25)

    if is_fmt_a:
        # 3 token cols: wallet(1), token(2), tok_simple(3)
        # prev: qty(4) px(5) val(6) pct(7)  |  actual: qty(9) px(10) val(11) pct(12) apy(13)
        # meta: comment(14) chain(15) method(16) proto(17)
        TK = 3; TH = 3
        QP, PP, VP, CP = 4, 5, 6, 7
        SC = None
        QA, PA, VA, CA, AA = 9, 10, 11, 12, 13
        CM, CH, ME, PR = 14, 15, 16, 17
    else:
        # 4 token cols: wallet(1), token(2), tok_half(3), tok_simple(4)
        # prev: qty(5) px(6) val(7) pct(8)  |  actual: qty(10) px(11) val(12) pct(13) apy(14)
        # meta: comment(15) chain(16) method(17) proto(18)
        TK = 4; TH = 3
        QP, PP, VP, CP = 5, 6, 7, 8
        SC = 9 if has_sell else None
        QA, PA, VA, CA, AA = 10, 11, 12, 13, 14
        CM, CH, ME, PR = 15, 16, 17, 18

    # Find table header row (row with 'wallet' in col 1)
    tbl_hdr = 45
    for i in range(40, 60):
        if i < len(df) and str(df.iloc[i, 1]).strip().lower() == 'wallet':
            tbl_hdr = i; break

    tbl = df.iloc[tbl_hdr+1:].copy()
    tbl.columns = range(len(tbl.columns))
    tbl = tbl.dropna(how='all')

    positions = []
    for _, rd in tbl.iterrows():
        wv = safe_str(rd.get(1))
        if not wv or wv.lower() in ('grand total','nan','total',''): continue
        tv = safe_str(rd.get(2))
        if not tv or tv.lower() == 'nan': continue
        ts = safe_str(rd.get(TK)).upper()
        if not ts or ts == 'NAN': continue
        positions.append({
            'wallet':    wv,
            'token':     tv,
            'tok_half':  safe_str(rd.get(TH)),
            'tok_simple':ts,
            'qty_p':  safe_float(rd.get(QP)),
            'px_p':   safe_float(rd.get(PP)),
            'val_p':  safe_float(rd.get(VP)),
            'pct_p':  safe_float(rd.get(CP)),
            'sell':   safe_str(rd.get(SC)) if SC else None,
            'qty_a':  safe_float(rd.get(QA)),
            'px_a':   safe_float(rd.get(PA)),
            'val_a':  safe_float(rd.get(VA)),
            'pct_a':  safe_float(rd.get(CA)),
            'apy_a':  safe_float(rd.get(AA)),
            'comment':safe_str(rd.get(CM)),
            'chain':  safe_str(rd.get(CH)),
            'method': safe_str(rd.get(ME)).lower(),
            'protocol':safe_str(rd.get(PR)),
        })

    tok_t, cat_t, meth_t, chain_t = (defaultdict(float) for _ in range(4))
    for p in positions:
        val = p['val_a'] or 0
        if val <= 0: continue
        tok = p['tok_simple'] or p['token'].upper()
        tok_t[tok] += val
        cat_t[token_cat.get(tok, 'other')] += val
        m = p['method'] or 'other'
        if m and m != 'nan': meth_t[m] += val
        c = p['chain'] or 'other'
        if c and c != 'nan': chain_t[c] += val

    all_data[sheet_name] = {
        'label': MONTH_LABELS[sheet_name],
        'synthetic': False,
        'kpi': kpi,
        'token_alloc': top_n(tok_t),
        'cat_alloc':   top_n(cat_t),
        'method_alloc':top_n(meth_t),
        'chain_alloc': top_n(chain_t),
        'positions':   positions,
    }

# ── Synthetic month generator ────────────────────────────────────────────────

def generate_synthetic_months(base_month_key='2508'):
    price_cache = load_price_cache()
    result = {}
    prev_positions = all_data[base_month_key]['positions']
    prev_actual    = all_data[base_month_key]['kpi']['actual']

    for month_idx, month_key in enumerate(SYNTHETIC_MONTHS):
        date_str  = SYNTHETIC_DATES[month_key]
        month_num = month_idx + 1   # 1=Sep, 2=Oct, 3=Nov, 4=Dec
        label     = SYNTHETIC_MONTH_LABELS[month_key]
        print(f'\n── Generating {label} ({date_str}) ──')

        new_positions = []
        for p in prev_positions:
            pos = deepcopy(p)

            # Roll prev from previous month's actual
            pos['qty_p'] = p['qty_a']
            pos['px_p']  = p['px_a']
            pos['val_p'] = p['val_a']
            pos['pct_p'] = p['pct_a']

            # Compound quantity by APY
            apy  = pos['apy_a'] or 0
            meth = pos['method'] or ''
            old_qty = pos['qty_a']
            if old_qty is not None and apy != 0:
                if meth in ('farming', 'staking'):
                    # Qty grows: monthly = (1+APY)^(1/12)
                    pos['qty_a'] = old_qty * ((1 + apy) ** (1/12))
                elif meth == 'debt':
                    # Debt grows: simple monthly rate
                    pos['qty_a'] = old_qty * (1 + apy / 12)
                # collateral/wallet/other: qty unchanged

            # PT-sENA-25SEP2025 matures Sep 25; from Oct+ convert to ENA wallet
            if pos['token'] == 'PT-sENA-25SEP2025' and month_num >= 2:
                pos['token']  = 'ENA'
                pos['tok_simple'] = 'ENA'
                pos['method'] = 'wallet'
                pos['apy_a']  = 0.0
                pos['protocol'] = ''

            # Fetch CoinGecko price
            new_px, source = get_price_for_position(
                pos['token'], pos['tok_simple'], date_str, price_cache
            )
            if new_px is not None:
                pos['px_a'] = new_px
            # else: roll forward (keep px_a from previous month)

            # Recompute val_a
            if pos['qty_a'] is not None and pos['px_a'] is not None:
                pos['val_a'] = pos['qty_a'] * pos['px_a']
            pos['pct_a'] = None  # recomputed below

            new_positions.append(pos)

        # Recompute pct_a after knowing total
        total_gross = sum(p['val_a'] for p in new_positions
                          if p['val_a'] is not None and p['val_a'] > 0)
        if total_gross > 0:
            for pos in new_positions:
                if pos['val_a'] is not None:
                    pos['pct_a'] = pos['val_a'] / total_gross

        # Build allocations
        tok_t, cat_t, meth_t, chain_t = (defaultdict(float) for _ in range(4))
        for pos in new_positions:
            val = pos['val_a'] or 0
            if val <= 0: continue
            tok = pos['tok_simple'] or pos['token'].upper()
            tok_t[tok] += val
            cat_t[token_cat.get(tok, 'other')] += val
            m = pos['method'] or 'other'
            if m and m != 'nan': meth_t[m] += val
            c = pos['chain'] or 'other'
            if c and c != 'nan': chain_t[c] += val

        kpi = compute_month_kpis(new_positions, prev_actual)
        print(f'  → total portfolio: ${kpi["actual"]:,.0f}')

        result[month_key] = {
            'label':        label,
            'synthetic':    True,
            'kpi':          kpi,
            'token_alloc':  top_n(tok_t),
            'cat_alloc':    top_n(cat_t),
            'method_alloc': top_n(meth_t),
            'chain_alloc':  top_n(chain_t),
            'positions':    new_positions,
        }

        prev_positions = new_positions
        prev_actual    = kpi['actual']

    return result

# ── Generate synthetic months ─────────────────────────────────────────────────
print('Generating synthetic months Sep–Dec 2025 (CoinGecko prices)...')
synthetic_data = generate_synthetic_months('2508')
MONTHS = MONTHS + SYNTHETIC_MONTHS
MONTH_LABELS.update(SYNTHETIC_MONTH_LABELS)
all_data.update(synthetic_data)

# Sep/Oct/Nov 25: real data — remove estimate flags
for _mk, _lbl in [('2509','Sep 25'), ('2510','Oct 25'), ('2511','Nov 25')]:
    if _mk in all_data:
        all_data[_mk]['synthetic'] = False
        all_data[_mk]['label']     = _lbl
        MONTH_LABELS[_mk]          = _lbl

# ── Portfolio value history ───────────────────────────────────────────────────
port_hist = [
    {'month': MONTH_LABELS[m], 'value': all_data[m]['kpi']['actual'],
     'prev':  all_data[m]['kpi'].get('prev') or 0,
     'year':  '20' + m[:2],
     'synthetic': all_data[m].get('synthetic', False)}
    for m in MONTHS
]

# ── Targets vs current (2508 – always real data) ──────────────────────────────
pct_sums = defaultdict(float)
for p in all_data['2508']['positions']:
    pct_sums[p['tok_simple']] += p['pct_a'] or 0

target_comp = []
for wt in wallet_targets:
    tok = wt['token'].upper()
    tgt = wt['t0725'] or wt['t0625'] or wt['t0225']
    cur = pct_sums.get(tok, 0)
    if (tgt and abs(tgt) > 0.001) or abs(cur) > 0.001:
        target_comp.append({
            'token':   wt['token'],
            'target':  round(tgt or 0, 4),
            'current': round(cur, 4),
            'cat':     wt['category'],
        })

target_tokens_set = {item['token'].upper() for item in target_comp}
for tok, cur in sorted(pct_sums.items(), key=lambda x: x[1]):
    if cur < -0.001 and tok not in target_tokens_set:
        target_comp.append({
            'token':   tok,
            'target':  0,
            'current': round(cur, 4),
            'cat':     token_cat.get(tok, 'stable'),
        })
        target_tokens_set.add(tok)

# ── JSON payload ──────────────────────────────────────────────────────────────
payload = {
    'months': MONTHS,
    'data': all_data,
    'port_hist': port_hist,
    'target_comp': target_comp,
    'wallet_targets': wallet_targets,
    'token_cat': token_cat,
}

json_str = json.dumps(payload, ensure_ascii=False, default=str)

# ── HTML template ─────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BB8 Portfolio Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
:root{
  --bg:#0D1117;--card:#161B22;--card2:#1C2128;--border:#30363D;
  --text:#E6EDF3;--muted:#7D8590;--cyan:#58A6FF;--green:#3FB950;
  --orange:#F78166;--yellow:#E3B341;--purple:#D2A8FF;--red:#FF7B72;
  --teal:#39D353;--pink:#F778BA;
}
*{box-sizing:border-box;margin:0;padding:0}
html{font-size:13px}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Inter',sans-serif;
  background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}

/* ── HEADER ── */
header{background:var(--card);border-bottom:1px solid var(--border);
  padding:8px 20px;display:flex;align-items:flex-start;gap:16px;position:sticky;top:0;z-index:100}
.logo{font-size:1.4rem;font-weight:700;letter-spacing:2px;padding-top:4px;
  background:linear-gradient(135deg,var(--cyan),var(--green));
  -webkit-background-clip:text;-webkit-text-fill-color:transparent}
.tabs{display:flex;flex-direction:column;gap:3px}
.tab-year-row{display:flex;gap:4px;align-items:center;flex-wrap:wrap}
.tab-year-label{font-size:.65rem;font-weight:700;width:34px;flex-shrink:0;
  letter-spacing:.4px;text-align:right;padding-right:4px}
.tab{padding:5px 12px;border-radius:6px;cursor:pointer;font-size:.8rem;font-weight:500;
  background:transparent;border:1px solid var(--border);color:var(--muted);transition:all .15s}
.tab:hover{border-color:var(--cyan);color:var(--cyan)}
.tab.active{background:var(--cyan);border-color:var(--cyan);color:#000;font-weight:600}
.tab.synthetic{border-style:dashed;opacity:.85}
.tab.synthetic.active{background:rgba(88,166,255,.75);border-color:var(--cyan);color:#000;opacity:1}
.synth-note{display:inline-block;font-size:.6rem;background:rgba(88,166,255,.15);
  border:1px solid rgba(88,166,255,.4);border-radius:3px;padding:0 4px;
  margin-left:4px;color:var(--cyan);vertical-align:middle}
.header-meta{margin-left:auto;font-size:.75rem;color:var(--muted)}

/* ── MAIN ── */
main{padding:16px 20px;display:flex;flex-direction:column;gap:14px}

/* ── BALANCE CARDS (in left panel) ── */
.bal-stack{display:flex;flex-direction:column;gap:6px;margin-bottom:8px}
.bal-card{background:var(--card2);border:1px solid var(--border);border-radius:8px;
  padding:8px 12px;display:flex;flex-direction:column;gap:2px}
.bal-card .label{font-size:.6rem;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.bal-card .value{font-size:1.3rem;font-weight:700;font-variant-numeric:tabular-nums}
.bal-card .sub{font-size:.7rem;color:var(--muted)}
.v-start{color:var(--muted)}.v-prev{color:var(--yellow)}.v-actual{color:var(--cyan)}
#val-mom{font-size:.9rem;font-weight:600;margin-top:3px;line-height:1.3}
#val-mom.pos{color:var(--green)}#val-mom.neg{color:var(--orange)}

/* ── DASHBOARD MAIN: KPI left | Charts right ── */
.dash-main{display:grid;grid-template-columns:260px 1fr;gap:14px;align-items:start}
@media(max-width:1000px){.dash-main{grid-template-columns:1fr}}

/* ── KPI PANEL ── */
#kpi-panel{display:flex;flex-direction:column;gap:0}
#kpi-row{background:var(--card);border:1px solid var(--border);border-radius:10px;
  padding:12px;display:grid;grid-template-columns:1fr 1fr;gap:8px;align-content:start}
.kpi{background:var(--card2);border:1px solid var(--border);border-radius:8px;
  padding:8px 10px;display:flex;flex-direction:column;gap:2px}
.kpi .kpi-label{font-size:.6rem;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;line-height:1.3}
.kpi .kpi-vals{display:flex;gap:6px;align-items:baseline;flex-wrap:wrap;margin-top:2px}
.kpi .kpi-main{font-size:.95rem;font-weight:600;font-variant-numeric:tabular-nums}
.kpi .kpi-prev{font-size:.7rem;color:var(--muted);font-variant-numeric:tabular-nums}
.pos{color:var(--green)} .neg{color:var(--red)} .neu{color:var(--cyan)}

/* ── CHARTS AREA ── */
.charts-area{display:flex;flex-direction:column;gap:12px}
.charts-top{display:grid;grid-template-columns:1fr 1fr 1.4fr;gap:12px}
.charts-bot{display:grid;grid-template-columns:2fr 1fr;gap:12px}
@media(max-width:1200px){.charts-top{grid-template-columns:1fr 1fr}}
@media(max-width:700px){.charts-top,.charts-bot{grid-template-columns:1fr}}

.chart-card{background:var(--card);border:1px solid var(--border);border-radius:10px;
  padding:14px;display:flex;flex-direction:column;gap:8px;min-height:260px}
.chart-card .card-title{font-size:.75rem;font-weight:600;color:var(--muted);
  text-transform:uppercase;letter-spacing:.5px}
.chart-wrap{position:relative;flex:1;min-height:200px}
.chart-wrap canvas{position:absolute;inset:0;width:100%!important;height:100%!important}

/* ── TRACK RECORD ── */
.tr-card{background:var(--card);border:1px solid var(--border);border-radius:10px;
  padding:14px 16px;margin-bottom:14px}
.tr-card-title{font-size:.7rem;font-weight:600;color:var(--muted);text-transform:uppercase;
  letter-spacing:.6px;margin-bottom:10px}
.tr-wrap{overflow-x:auto}
.tr-table{border-collapse:collapse;width:100%;font-size:.78rem;white-space:nowrap}
.tr-table th{color:var(--muted);font-weight:600;font-size:.65rem;letter-spacing:.4px;
  padding:6px 8px;text-align:center;border-bottom:1px solid var(--border)}
.tr-table th:first-child{text-align:left;padding-left:12px;min-width:58px}
.tr-table th:last-child{text-align:right;padding-right:12px;min-width:72px}
.tr-table td{padding:9px 8px;text-align:center;border-bottom:1px solid rgba(255,255,255,.04)}
.tr-table tr:last-child td{border-bottom:none}
.tr-year-cell{text-align:left!important;padding-left:12px!important;padding-right:8px}
.tr-total-cell{text-align:right!important;padding-right:12px!important;font-size:.82rem}
.tr-note{font-size:.62rem;color:var(--muted);margin-top:8px;line-height:1.5}

/* ── TABLE ── */
.table-card{background:var(--card);border:1px solid var(--border);border-radius:10px;overflow:hidden}
.table-toolbar{padding:12px 16px;display:flex;gap:10px;align-items:center;
  border-bottom:1px solid var(--border);flex-wrap:wrap}
.table-toolbar input,.table-toolbar select{
  background:var(--card2);border:1px solid var(--border);color:var(--text);
  border-radius:6px;padding:5px 10px;font-size:.8rem;outline:none}
.table-toolbar input:focus,.table-toolbar select:focus{border-color:var(--cyan)}
.table-toolbar .tcount{margin-left:auto;font-size:.75rem;color:var(--muted)}
.tbl-wrap{overflow-x:auto;max-height:520px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:.78rem}
thead th{background:var(--card2);color:var(--muted);font-weight:600;text-transform:uppercase;
  font-size:.65rem;letter-spacing:.3px;padding:8px 10px;text-align:left;
  position:sticky;top:0;z-index:5;white-space:nowrap;cursor:pointer;user-select:none}
thead th:hover{color:var(--cyan)}
thead th.sort-asc::after{content:' ↑';color:var(--cyan)}
thead th.sort-desc::after{content:' ↓';color:var(--cyan)}
tbody tr{border-bottom:1px solid var(--border)}
tbody tr:nth-child(even){background:#0f151c}
tbody tr:hover{background:#1a2535}
td{padding:6px 10px;white-space:nowrap;color:var(--text)}
td.token-col{font-weight:600}
td.val-col{font-variant-numeric:tabular-nums;text-align:right}
td.pct-col{text-align:right}
td.pos-val{color:var(--green)} td.neg-val{color:var(--red)}
.badge{display:inline-block;padding:1px 6px;border-radius:4px;font-size:.65rem;font-weight:500}
.b-wallet{background:#1f2a3c;color:var(--cyan)}
.b-collateral{background:#1a2b1a;color:var(--green)}
.b-debt{background:#2b1a1a;color:var(--red)}
.b-farming{background:#2b2a1a;color:var(--yellow)}
.b-staking{background:#1e1a2b;color:var(--purple)}
.b-nft{background:#2b1a2a;color:var(--pink)}
.b-other{background:#1f2128;color:var(--muted)}
.sell-hodl{color:var(--muted);font-size:.7rem}
.sell-sell{color:var(--red);font-weight:600;font-size:.7rem}

/* ── FOOTER ── */
footer{text-align:center;padding:16px;color:var(--muted);font-size:.7rem;
  border-top:1px solid var(--border);margin-top:8px}
</style>
</head>
<body>

<header>
  <div class="logo">⬡ BB8</div>
  <div class="tabs" id="tabs"></div>
  <div class="header-meta" id="header-meta"></div>
</header>

<main>
  <!-- KPI left | Charts right -->
  <div class="dash-main">
    <div id="kpi-panel">
      <!-- Balance cards stacked in left panel -->
      <div class="bal-stack">
        <div class="bal-card">
          <div class="label">Starting Balance</div>
          <div class="value v-start" id="val-start">—</div>
        </div>
        <div class="bal-card">
          <div class="label">Previous Balance</div>
          <div class="value v-prev" id="val-prev">—</div>
        </div>
        <div class="bal-card">
          <div class="label">Actual Balance</div>
          <div class="value v-actual" id="val-actual">—</div>
          <div id="val-mom">—</div>
          <div class="sub" id="val-diff" style="font-size:.6rem;opacity:.65">—</div>
        </div>
      </div>
      <!-- KPI grid -->
      <div id="kpi-row"></div>
    </div>
    <div class="charts-area">
      <div class="charts-top">
        <div class="chart-card">
          <div class="card-title" id="ct1">Portfolio allocation / token</div>
          <div class="chart-wrap"><canvas id="chart-token"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="card-title" id="ct2">Portfolio allocation / category</div>
          <div class="chart-wrap"><canvas id="chart-cat"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="card-title">Portfolio value ($) · 2023–2025</div>
          <div class="chart-wrap"><canvas id="chart-port"></canvas></div>
        </div>
      </div>
      <div class="charts-bot">
        <div class="chart-card">
          <div class="card-title">Targets wallet VS current wallet (Aug 25 · actual)</div>
          <div class="chart-wrap" style="min-height:240px"><canvas id="chart-targets"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="card-title" id="ct5">Method allocation</div>
          <div class="chart-wrap"><canvas id="chart-method"></canvas></div>
        </div>
      </div>
    </div>
  </div>

  <!-- Track Record -->
  <div class="tr-card">
    <div class="tr-card-title">Monthly Track Record</div>
    <div class="tr-wrap">
      <table class="tr-table">
        <thead><tr>
          <th>YEAR</th>__TR_HEADER__<th>TOTAL</th>
        </tr></thead>
        <tbody>__TR_ROWS__</tbody>
      </table>
    </div>
    <p class="tr-note">Snapshot taken 1st of each month → each cell = performance of the previous month &nbsp;·&nbsp; Total = (Jan snapshot year N+1) / (Jan snapshot year N) − 1</p>
  </div>

  <!-- Position table (full width) -->
  <div class="table-card">
    <div class="table-toolbar">
      <input id="search" placeholder="🔍  Search token / chain…" style="width:220px">
      <select id="filter-wallet"><option value="">All wallets</option></select>
      <select id="filter-method"><option value="">All methods</option></select>
      <select id="filter-cat"><option value="">All categories</option></select>
      <span class="tcount" id="tcount"></span>
    </div>
    <div class="tbl-wrap">
      <table id="pos-table">
        <thead>
          <tr>
            <th data-col="token">Token</th>
            <th data-col="tok_simple">Simplified</th>
            <th data-col="chain">Chain</th>
            <th data-col="method">Method</th>
            <th data-col="protocol">Protocol</th>
            <th data-col="qty_p" style="text-align:right">Qty (prev)</th>
            <th data-col="px_p"  style="text-align:right">Px (prev $)</th>
            <th data-col="val_p" style="text-align:right">Val (prev $)</th>
            <th data-col="qty_a" style="text-align:right">Qty (act)</th>
            <th data-col="px_a"  style="text-align:right">Px (act $)</th>
            <th data-col="val_a" style="text-align:right">Val (act $)</th>
            <th data-col="pct_a" style="text-align:right">% (act)</th>
            <th data-col="apy_a" style="text-align:right">APY</th>
            <th data-col="sell">Sell?</th>
            <th data-col="comment">Comment</th>
          </tr>
        </thead>
        <tbody id="tbl-body"></tbody>
      </table>
    </div>
  </div>
</main>

<footer>BB8 Portfolio Dashboard · Dec 24–Aug 25 from xlsx · Sep–Dec 25 estimated via CoinGecko historical prices</footer>

<script>
// ── DATA ──────────────────────────────────────────────────────────────────────
const D = __DATA_PLACEHOLDER__;

// Register datalabels plugin globally
Chart.register(ChartDataLabels);

// ── CHART COLORS ──────────────────────────────────────────────────────────────
const PALETTE = [
  '#58A6FF','#3FB950','#F78166','#E3B341','#D2A8FF',
  '#79C0FF','#56D364','#FFA657','#F778BA','#A5D6FF',
  '#39D353','#FF7B72','#BC8CFF','#FFAB70','#6FCCF5',
  '#2EA043','#DA3633','#1F6FEB','#9E6A03','#6E40C9',
];
const CAT_COLORS = {
  'defi':'#58A6FF','infra':'#3FB950','stable':'#E3B341','nft':'#D2A8FF',
  'bitcoin':'#F78166','rwa':'#79C0FF','perp':'#F778BA','lending':'#56D364',
  'farming':'#FFA657','dex':'#A5D6FF','oracle':'#FF7B72','ai':'#FFAB70',
  'lrt':'#39D353','lst':'#BC8CFF','bridge':'#6FCCF5','other':'#7D8590',
};
const METHOD_COLORS = {
  'collateral':'#3FB950','debt':'#F85149','wallet':'#58A6FF',
  'staking':'#D2A8FF','farming':'#E3B341','nft':'#F778BA',
  'lending':'#79C0FF','pending rewards':'#FFA657','other':'#7D8590',
};

// ── STATE ─────────────────────────────────────────────────────────────────────
let currentMonth = '2508';
let charts = {};
let sortCol = 'val_a', sortDir = -1;
let positions = [];

// ── HELPERS ───────────────────────────────────────────────────────────────────
const fmt$ = v => v == null ? '—' : (v < 0 ? '-$' : '$') + Math.abs(Math.round(v)).toLocaleString();
const fmtM = v => v == null ? '—' : (v < 0 ? '-$' : '$') + (Math.abs(v)/1e6).toFixed(2) + 'M';
const fmtPctDirect = (v, dec=2) => v == null ? '—' : (v * 100).toFixed(dec) + '%';
const sign = v => v > 0 ? '+' : '';
const colorNum = v => v == null ? 'neu' : (v > 0 ? 'pos' : v < 0 ? 'neg' : 'neu');

// Quantity: large ints comma-separated; small floats with precision
const fmtQty = v => {
  if (v == null) return '—';
  const abs = Math.abs(v);
  const neg = v < 0 ? '-' : '';
  if (abs >= 10000) return neg + Math.round(abs).toLocaleString();
  if (abs >= 1)     return v.toLocaleString('en', {maximumFractionDigits: 2});
  if (abs > 0)      return v.toPrecision(3);
  return '0';
};

// Price: $X.XX for >=1, scientific for tiny
const fmtPx = v => {
  if (v == null) return '—';
  if (v === 0) return '$0';
  if (v >= 1)  return '$' + v.toLocaleString('en', {minimumFractionDigits:2, maximumFractionDigits:2});
  if (v > 0)   return '$' + v.toPrecision(3);
  return '—';
};

function methodBadgeClass(m) {
  const map = {collateral:'b-collateral',debt:'b-debt',farming:'b-farming',
    staking:'b-staking',nft:'b-nft',wallet:'b-wallet'};
  return map[m] || 'b-other';
}

// ── TABS ──────────────────────────────────────────────────────────────────────
function buildTabs() {
  const el = document.getElementById('tabs');
  const yearDefs = [
    {pfx:'23', label:'2023', col:'#58A6FF'},
    {pfx:'24', label:'2024', col:'#D2A8FF'},
    {pfx:'25', label:'2025', col:'#3FB950'},
  ];
  yearDefs.forEach(({pfx, label, col}) => {
    const months = D.months.filter(m => m.startsWith(pfx));
    if (!months.length) return;
    const row = document.createElement('div');
    row.className = 'tab-year-row';
    const lbl = document.createElement('span');
    lbl.className = 'tab-year-label';
    lbl.textContent = label;
    lbl.style.color = col;
    row.appendChild(lbl);
    months.forEach(m => {
      const btn = document.createElement('button');
      const isSyn = D.data[m].synthetic;
      btn.className = 'tab' + (m === currentMonth ? ' active' : '') + (isSyn ? ' synthetic' : '');
      btn.textContent = D.data[m].label;
      btn.onclick = () => selectMonth(m);
      btn.id = 'tab-' + m;
      row.appendChild(btn);
    });
    el.appendChild(row);
  });
}
function selectMonth(m) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + m).classList.add('active');
  currentMonth = m;
  render();
}

// ── KPI RENDER ────────────────────────────────────────────────────────────────
function renderKPIs() {
  const md = D.data[currentMonth];
  const k  = md.kpi;
  const isSyn = md.synthetic;

  document.getElementById('val-start').textContent  = fmtM(k.start);
  document.getElementById('val-prev').textContent   = fmtM(k.prev);
  document.getElementById('val-actual').textContent = fmtM(k.actual);

  const momVal = k.actual - k.prev;
  const momPct = k.prev > 0 ? (momVal / k.prev * 100) : null;
  const momSign = momVal >= 0 ? '+' : '';
  const momEl = document.getElementById('val-mom');
  if (k.prev > 0) {
    momEl.textContent = `${momSign}${fmtM(momVal)} vs prev (${momSign}${momPct.toFixed(1)}%)`;
    momEl.className = momVal >= 0 ? 'pos' : 'neg';
  } else {
    momEl.textContent = '—'; momEl.className = '';
  }

  document.getElementById('val-diff').textContent   =
    `${sign(k.diff_val)}${fmtM(k.diff_val)}  (${sign(k.diff_pct)}${k.diff_pct.toFixed(1)}% since start)`;

  const hfMin = k.hf_min === 9999 ? '+∞'
              : k.hf_min != null  ? k.hf_min.toFixed(2)
              : isSyn ? 'N/A' : '—';
  const hfGlb = k.hf_global === 9999 ? '+∞'
              : k.hf_global != null  ? k.hf_global.toFixed(2)
              : isSyn ? 'N/A' : '—';

  const kpis = [
    { label:'différence / starting',
      main: `${sign(k.diff_val)}${fmtM(k.diff_val)}`,
      prev: `${sign(k.diff_pct)}${k.diff_pct != null ? k.diff_pct.toFixed(1)+'%' : '—'}`,
      cls: colorNum(k.diff_val) },
    { label:'min health factor',    main: hfMin, cls:'neu' },
    { label:'health factor (global)',main: hfGlb, cls:'neu' },
    { label:'debt ($)',
      main: k.debt != null ? fmt$(k.debt) : '—',
      cls: k.debt != null && k.debt < 0 ? 'neg' : 'pos' },
    { label:'weighted APY',
      main: k.apy_c != null ? fmtPctDirect(k.apy_c) : '—',
      prev: k.apy_p != null ? fmtPctDirect(k.apy_p) : '',
      cls:'pos' },
    { label:'$ non investi',
      main: k.ni_c != null ? fmtM(k.ni_c) : '—',
      prev: k.ni_p != null ? fmtM(k.ni_p) : '',
      cls:'neu' },
    { label:'% non EVM',
      main: k.ne_c != null ? fmtPctDirect(k.ne_c) : '—',
      prev: k.ne_p != null ? fmtPctDirect(k.ne_p) : '',
      cls:'neu' },
    { label:'altcoins (%)',
      main: k.alt_c != null ? fmtPctDirect(k.alt_c) : '—',
      prev: k.alt_p != null ? fmtPctDirect(k.alt_p) : '',
      cls:'neu' },
  ];

  // Extra KPIs derived from positions
  const positions = md.positions || [];
  const nPos   = positions.length;
  const nProto = new Set(positions.map(p => p.protocol).filter(p => p && p !== 'nan')).size;
  const sumM = (...methods) => positions
    .filter(p => methods.includes(p.method) && p.val_a != null && p.val_a > 0)
    .reduce((s, p) => s + p.val_a, 0);
  const collat  = sumM('collateral');
  const farming = sumM('farming');
  const staking = sumM('staking', 'staking/lp');
  kpis.push(
    { label:'# positions / protocols', main: nPos, prev: nProto + ' protocols', cls:'neu' },
    { label:'collateral ($)',  main: collat  > 0 ? fmtM(collat)  : '—', cls:'neu' },
    { label:'farming ($)',     main: farming > 0 ? fmtM(farming) : '—', cls:'pos' },
    { label:'staking ($)',     main: staking > 0 ? fmtM(staking) : '—', cls:'pos' },
  );

  const row = document.getElementById('kpi-row');
  row.innerHTML = kpis.map(k => `
    <div class="kpi">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-vals">
        <span class="kpi-main ${k.cls}">${k.main}</span>
        ${k.prev ? `<span class="kpi-prev">${k.prev}</span>` : ''}
      </div>
    </div>`).join('');

  const synthTag = isSyn ? ' · <span style="color:var(--yellow);font-size:.7rem">estimated</span>' : '';
  document.getElementById('header-meta').innerHTML =
    `Last update · ${md.label}${synthTag} · Total: ${fmtM(k.actual)}`;
}

// ── PIE / DOUGHNUT CHART ──────────────────────────────────────────────────────
function buildPie(id, items, colorMap) {
  if (charts[id]) { charts[id].destroy(); }
  const ctx = document.getElementById(id);
  if (!ctx || !items.length) return;
  const labels = items.map(d => d.label);
  const values = items.map(d => d.value);
  const colors = labels.map((l,i) => colorMap
    ? (colorMap[l.toLowerCase()] || PALETTE[i % PALETTE.length])
    : PALETTE[i % PALETTE.length]);
  const total = values.reduce((a,b) => a+b, 0);

  charts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data: values, backgroundColor: colors,
      borderColor: '#161B22', borderWidth: 2, hoverOffset: 6 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '40%',
      plugins: {
        legend: { display: false },
        datalabels: {
          color: '#fff',
          display: (ctx) => ctx.dataset.data[ctx.dataIndex] / total * 100 >= 3.5,
          font: (ctx) => {
            const pct = ctx.dataset.data[ctx.dataIndex] / total * 100;
            return { size: pct >= 10 ? 9 : 8, weight: 'bold' };
          },
          formatter: (value, ctx) => {
            const pct = value / total * 100;
            const pctStr = pct.toFixed(1) + '%';
            const label  = ctx.chart.data.labels[ctx.dataIndex];
            return pct >= 8 ? label + '\n' + pctStr : pctStr;
          },
          anchor: 'center', align: 'center', textAlign: 'center',
        },
        tooltip: { callbacks: { label: ctx => {
          const t = ctx.dataset.data.reduce((a,b)=>a+b,0);
          const pct = (ctx.parsed / t * 100).toFixed(1);
          return ` ${ctx.label}: $${Math.round(ctx.parsed).toLocaleString()} (${pct}%)`;
        }}}
      }
    }
  });
}

// ── BAR CHART – portfolio value ───────────────────────────────────────────────
function buildPortBar() {
  if (charts['chart-port']) { charts['chart-port'].destroy(); }
  const ctx = document.getElementById('chart-port');
  if (!ctx) return;
  const labels = D.port_hist.map(d => d.month);
  const values = D.port_hist.map(d => d.value);
  const colors = D.port_hist.map((d, i) => {
    const m = D.months[i];
    const syn = d.synthetic;
    if (m === currentMonth) return syn ? 'rgba(88,166,255,0.85)' : '#58A6FF';
    return syn ? 'rgba(88,166,255,0.2)' : 'rgba(88,166,255,0.35)';
  });

  // Per-year max (solid) and min (dashed) horizontal lines
  const yrColors = {'2023':'#58A6FF','2024':'#D2A8FF','2025':'#3FB950'};
  const yrDatasets = [];
  ['2023','2024','2025'].forEach(yr => {
    const yrVals = D.port_hist.filter(d => d.year === yr).map(d => d.value);
    if (!yrVals.length) return;
    const maxVal = Math.max(...yrVals);
    const minVal = Math.min(...yrVals);
    const col = yrColors[yr];
    const mask = D.port_hist.map(d => d.year === yr);
    yrDatasets.push({
      label:`Max ${yr}`, type:'line', yAxisID:'y',
      data: mask.map((ok,i) => ok ? maxVal : null),
      borderColor:col, borderWidth:1.5, borderDash:[], pointRadius:0,
      fill:false, tension:0, spanGaps:false
    });
    yrDatasets.push({
      label:`Min ${yr}`, type:'line', yAxisID:'y',
      data: mask.map((ok,i) => ok ? minVal : null),
      borderColor:col, borderWidth:1, borderDash:[4,3], pointRadius:0,
      fill:false, tension:0, spanGaps:false
    });
  });

  // MoM volatility line (secondary axis)
  const momData = D.port_hist.map(d => d.prev > 0 ? d.value - d.prev : null);
  const momDataset = {
    label:'MoM Change', type:'line', yAxisID:'y2',
    data: momData,
    borderColor:'#E3B341', borderWidth:1.5,
    pointRadius:2, pointBackgroundColor:'#E3B341',
    fill:false, tension:0.2, spanGaps:false
  };

  charts['chart-port'] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [
      { label:'Portfolio Value', data:values, backgroundColor:colors,
        borderRadius:4, borderSkipped:false, yAxisID:'y' },
      ...yrDatasets,
      momDataset,
    ]},
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: { display: false },
        tooltip: { callbacks: { label: c => {
          if (c.dataset.label === 'MoM Change') {
            const v = c.parsed.y;
            return ' ' + (v>=0?'+':'') + fmtM(v) + ' MoM';
          }
          if (c.dataset.label.startsWith('Max') || c.dataset.label.startsWith('Min')) {
            return ' ' + c.dataset.label + ': ' + fmtM(c.parsed.y);
          }
          const isSyn = D.port_hist[c.dataIndex]?.synthetic;
          return ' ' + fmtM(c.parsed.y) + (isSyn ? ' (est.)' : '');
        }}}
      },
      scales: {
        x: { grid:{color:'#21262D'}, ticks:{color:'#7D8590',font:{size:9}} },
        y: { grid:{color:'#21262D'}, ticks:{
          color:'#7D8590', font:{size:9},
          callback: v => '$'+(v/1e6).toFixed(0)+'M'
        }},
        y2: {
          position:'right',
          grid:{drawOnChartArea:false},
          ticks:{ color:'#E3B341', font:{size:8},
            callback: v => (v>=0?'+':'')+'$'+(v/1e6).toFixed(0)+'M' }
        }
      }
    }
  });
}

// ── BAR CHART – targets vs current ───────────────────────────────────────────
function buildTargetsBar() {
  if (charts['chart-targets']) { charts['chart-targets'].destroy(); }
  const ctx = document.getElementById('chart-targets');
  if (!ctx) return;
  const items = D.target_comp
    .filter(d => Math.abs(d.target) > 0.005 || Math.abs(d.current) > 0.005)
    .sort((a,b) => b.target - a.target).slice(0, 30);
  const labels  = items.map(d => d.token);
  const targets = items.map(d => +(d.target * 100).toFixed(2));
  const currents= items.map(d => +(d.current * 100).toFixed(2));
  const minVal  = Math.min(0, ...currents, ...targets);
  charts['chart-targets'] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [
      { label: '0825_target', data: targets,
        backgroundColor: 'rgba(88,166,255,0.6)',
        borderColor: '#58A6FF', borderWidth: 1, borderRadius: 3 },
      { label: '0825_current', data: currents,
        backgroundColor: currents.map(v => v < 0 ? 'rgba(248,81,73,0.7)' : 'rgba(63,185,80,0.7)'),
        borderColor: currents.map(v => v < 0 ? '#F85149' : '#3FB950'),
        borderWidth: 1, borderRadius: 3 },
    ]},
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels:{color:'#E6EDF3',font:{size:10},boxWidth:12} },
        datalabels: { display: false },
        tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.y.toFixed(2)}%` }}
      },
      scales: {
        x: { grid:{color:'#21262D'}, ticks:{color:'#7D8590',font:{size:9},maxRotation:45} },
        y: { min: minVal < 0 ? minVal * 1.2 : undefined,
          grid:{color:'#21262D'}, ticks:{color:'#7D8590',font:{size:9},callback:v=>v+'%'} }
      }
    }
  });
}

// ── TABLE ─────────────────────────────────────────────────────────────────────
function populateFilters() {
  const md = D.data[currentMonth];
  positions = md.positions;
  const wallets = [...new Set(positions.map(p => p.wallet))].sort();
  const methods = [...new Set(positions.map(p => p.method).filter(Boolean))].sort();
  const cats    = [...new Set(positions.map(p => D.token_cat[p.tok_simple] || 'other'))].sort();
  ['filter-wallet','filter-method','filter-cat'].forEach(id => {
    const sel = document.getElementById(id);
    const cur = sel.value;
    sel.innerHTML = '<option value="">' +
      (id==='filter-wallet'?'All wallets':id==='filter-method'?'All methods':'All categories') + '</option>';
    (id==='filter-wallet'?wallets:id==='filter-method'?methods:cats)
      .forEach(v => { const o=document.createElement('option');o.value=v;o.textContent=v;sel.appendChild(o); });
    sel.value = cur;
  });
}

function getFilteredPositions() {
  const q  = document.getElementById('search').value.toLowerCase();
  const fw = document.getElementById('filter-wallet').value;
  const fm = document.getElementById('filter-method').value;
  const fc = document.getElementById('filter-cat').value;
  return positions.filter(p => {
    if (fw && p.wallet !== fw) return false;
    if (fm && p.method !== fm) return false;
    if (fc && (D.token_cat[p.tok_simple] || 'other') !== fc) return false;
    if (q) {
      const hay = [p.token,p.tok_simple,p.chain,p.method,p.protocol,p.comment].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function renderTable() {
  const rows = [...getFilteredPositions()].sort((a,b) => {
    const va = a[sortCol] ?? (typeof a[sortCol]==='string' ? '' : -Infinity);
    const vb = b[sortCol] ?? (typeof b[sortCol]==='string' ? '' : -Infinity);
    if (typeof va === 'string') return va.localeCompare(vb) * sortDir;
    return (va - vb) * sortDir;
  });
  const total_val = rows.reduce((s,p) => s + (p.val_a || 0), 0);
  document.getElementById('tcount').textContent = `${rows.length} positions · total ${fmtM(total_val)}`;
  const body = document.getElementById('tbl-body');
  body.innerHTML = rows.map(p => {
    const valChg    = (p.val_a||0) - (p.val_p||0);
    const valChgCls = valChg > 0 ? 'pos-val' : valChg < 0 ? 'neg-val' : '';
    const apyColor  = (p.apy_a||0) > 0.1 ? 'color:#E3B341' : (p.apy_a||0) > 0 ? 'color:#3FB950' : '';
    const sellHtml  = p.sell && p.sell !== 'nan' && p.sell !== ''
      ? `<span class="${p.sell.toLowerCase().includes('hodl') ? 'sell-hodl' : 'sell-sell'}">${p.sell}</span>`
      : '';
    return `<tr>
      <td class="token-col">${p.token}</td>
      <td style="color:var(--yellow)">${p.tok_simple}</td>
      <td style="color:var(--muted);font-size:.72rem">${p.chain}</td>
      <td><span class="badge ${methodBadgeClass(p.method)}">${p.method||'—'}</span></td>
      <td style="color:var(--muted);font-size:.72rem">${p.protocol}</td>
      <td class="val-col" style="color:var(--muted);font-size:.72rem">${fmtQty(p.qty_p)}</td>
      <td class="val-col" style="color:var(--muted);font-size:.72rem">${fmtPx(p.px_p)}</td>
      <td class="val-col" style="color:var(--muted)">${p.val_p!=null ? fmt$(p.val_p) : '—'}</td>
      <td class="val-col" style="font-size:.72rem">${fmtQty(p.qty_a)}</td>
      <td class="val-col" style="font-size:.72rem">${fmtPx(p.px_a)}</td>
      <td class="val-col ${valChgCls}">${p.val_a!=null ? fmt$(p.val_a) : '—'}</td>
      <td class="pct-col">${p.pct_a!=null ? (p.pct_a*100).toFixed(2)+'%' : '—'}</td>
      <td class="pct-col" style="${apyColor}">${p.apy_a!=null&&p.apy_a!==0 ? (p.apy_a*100).toFixed(1)+'%' : '—'}</td>
      <td>${sellHtml}</td>
      <td style="color:var(--muted);max-width:140px;overflow:hidden;text-overflow:ellipsis" title="${p.comment}">${p.comment}</td>
    </tr>`;
  }).join('');
}

// ── SORT ──────────────────────────────────────────────────────────────────────
document.querySelectorAll('thead th[data-col]').forEach(th => {
  th.onclick = () => {
    const col = th.dataset.col;
    if (sortCol===col) sortDir*=-1; else {sortCol=col;sortDir=-1;}
    document.querySelectorAll('thead th').forEach(h=>h.classList.remove('sort-asc','sort-desc'));
    th.classList.add(sortDir===1?'sort-asc':'sort-desc');
    renderTable();
  };
});

// ── FILTER ────────────────────────────────────────────────────────────────────
['search','filter-wallet','filter-method','filter-cat'].forEach(id => {
  document.getElementById(id).addEventListener('input', renderTable);
});

// ── MAIN RENDER ───────────────────────────────────────────────────────────────
function render() {
  const md = D.data[currentMonth];
  renderKPIs();
  buildPie('chart-token',  md.token_alloc,  null);
  buildPie('chart-cat',    md.cat_alloc,    CAT_COLORS);
  buildPortBar();
  buildTargetsBar();
  buildPie('chart-method', md.method_alloc, METHOD_COLORS);
  document.getElementById('ct1').textContent = `Portfolio allocation / token · ${md.label}`;
  document.getElementById('ct2').textContent = `Portfolio allocation / category · ${md.label}`;
  document.getElementById('ct5').textContent = `Method allocation · ${md.label}`;
  populateFilters();
  renderTable();
}

// ── INIT ──────────────────────────────────────────────────────────────────────
buildTabs();
render();
document.querySelector('th[data-col="val_a"]').classList.add('sort-desc');
</script>
</body>
</html>
"""

# (injection happens after track record rows are built — see below)

# ── TRACK RECORD TABLE ────────────────────────────────────────────────────────
# Snapshot taken on 1st of month N → performance is for month N-1.
# Column shift: month_num - 2; January (1) wraps to December (11) of previous year.
# Cross-year: 2401 (Jan 24 snapshot = Dec 23 perf) → 2023 row DEC col.
#             2501 (Jan 25 snapshot = Dec 24 perf) → 2024 row DEC col.
year_months_map = {
    '2023': ['2309','2310','2311','2312','2401'],
    '2024': ['2402','2403','2404','2405','2406','2407','2408','2409','2410','2411','2412','2501'],
    '2025': ['2502','2503','2504','2505','2506','2507','2508','2509','2510','2511'],
}
COL_NAMES = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
YEAR_STYLES = {
    '2023': {'bar':'#3b82f6','bg':'#eff6ff','dot':'#3b82f6','dark':'#58A6FF'},
    '2024': {'bar':'#8b5cf6','bg':'#f5f3ff','dot':'#8b5cf6','dark':'#D2A8FF'},
    '2025': {'bar':'#22c55e','bg':'#f0fdf4','dot':'#22c55e','dark':'#3FB950'},
}

# TOTAL = (actual balance of first snapshot of next year) / (first snapshot of this year) - 1
# For 2023: end = Jan 24 actual / start = Sep 23 starting balance ($20M)
# For 2024: end = Jan 25 actual / start = Jan 24 actual
# For 2025: end = Nov 25 actual (last available) / start = Jan 25 actual (YTD)
def _act(mk): return all_data[mk]['kpi']['actual'] if mk in all_data else 0
YEAR_TOTAL = {
    '2023': (_act('2401'), all_data.get('2309',{}).get('kpi',{}).get('start') or 20_000_000),
    '2024': (_act('2501'), _act('2401')),
    '2025': (_act('2511'), _act('2501')),
}

track_rows = []
for year, months_list in year_months_map.items():
    cells = [None]*12
    is_syn = [False]*12
    for mk in months_list:
        if mk not in all_data: continue
        month_num = int(mk[2:])
        ci = 11 if month_num == 1 else month_num - 2
        k    = all_data[mk]['kpi']
        actual = k.get('actual') or 0
        prev   = k.get('prev')   or 0
        start  = k.get('start')  or 0
        base   = prev if prev > 0 else start
        syn    = all_data[mk].get('synthetic', False)
        if base > 0:
            cells[ci] = round((actual - base) / base * 100, 2)
        is_syn[ci] = syn
    end_val, start_val = YEAR_TOTAL[year]
    total = round((end_val / start_val - 1) * 100, 2) if start_val > 0 else None
    track_rows.append({'year': year, 'cells': cells, 'is_syn': is_syn, 'total': total})

def fmt_cell_light(val, syn):
    if val is None: return '<span style="color:#cbd5e1">—</span>'
    pct = f'{abs(val):.2f}%'
    if syn: return f'<span style="color:#ea580c;font-style:italic">{"+" if val>=0 else "-"}{pct}</span>'
    if val >= 0: return f'<span style="color:#16a34a">▲ {pct}</span>'
    return f'<span style="color:#64748b">-{pct}</span>'

def fmt_cell_dark(val, syn):
    if val is None: return '<span style="color:#30363D">—</span>'
    pct = f'{abs(val):.2f}%'
    if syn: return f'<span style="color:#F78166;font-style:italic">{"+" if val>=0 else "-"}{pct}</span>'
    if val >= 0: return f'<span style="color:#3FB950">▲ {pct}</span>'
    return f'<span style="color:#7D8590">-{pct}</span>'

def fmt_total_light(val, yr):
    if val is None: return '—'
    col  = YEAR_STYLES[yr]['dot']
    sign = '+' if val >= 0 else ''
    return f'<span style="color:{col};font-weight:700">{sign}{val:.2f}%</span>'

def fmt_total_dark(val, yr):
    if val is None: return '—'
    col  = YEAR_STYLES[yr]['dark']
    sign = '+' if val >= 0 else ''
    return f'<span style="color:{col};font-weight:700">{sign}{val:.2f}%</span>'

header_cols = ''.join(f'<th>{c}</th>' for c in COL_NAMES)

def build_cells(row, fmt_fn, future_dark=False):
    """Render only up to last non-None cell; collapse trailing empty ones."""
    cells  = row['cells']
    is_syn = row['is_syn']
    last_col = max((i for i in range(12) if cells[i] is not None), default=-1)
    html = ''.join(
        f'<td>{fmt_fn(cells[i], is_syn[i])}</td>'
        for i in range(last_col + 1)
    )
    remaining = 12 - (last_col + 1)
    if remaining > 0:
        sty = ('opacity:.15;color:#555;font-size:.6rem;text-align:center' if future_dark
               else 'opacity:.2;font-size:.6rem;text-align:center')
        html += f'<td colspan="{remaining}" style="{sty}">···</td>'
    return html

# ── Rows for dark dashboard embed ─────────────────────────────────────────────
tr_rows_dark = ''
for row in track_rows:
    yr  = row['year']
    col = YEAR_STYLES[yr]['dark']
    cells_html = build_cells(row, fmt_cell_dark, future_dark=True)
    tr_rows_dark += f'''
    <tr>
      <td class="tr-year-cell" style="border-left:3px solid {col}">
        <strong style="color:{col}">{yr}</strong>
      </td>
      {cells_html}
      <td class="tr-total-cell">{fmt_total_dark(row["total"], yr)}</td>
    </tr>'''

# ── Rows for standalone light file ────────────────────────────────────────────
rows_html_light = ''
for row in track_rows:
    yr = row['year']
    st = YEAR_STYLES[yr]
    cells_html = build_cells(row, fmt_cell_light)
    rows_html_light += f'''
    <tr style="background:{st["bg"]}">
      <td class="year-cell">
        <span class="year-bar" style="background:{st["bar"]}"></span>
        <strong>{yr}</strong>
      </td>
      {cells_html}
      <td class="total-cell">{fmt_total_light(row["total"], yr)}</td>
    </tr>'''

# ── Standalone track_record.html ──────────────────────────────────────────────
TR_HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>BB8 Track Record</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Inter',sans-serif;
  background:#f8fafc;color:#1e293b;padding:24px}}
h2{{font-size:1rem;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin-bottom:12px}}
.wrap{{overflow-x:auto;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
table{{border-collapse:collapse;width:100%;background:#fff;border-radius:12px;overflow:hidden;font-size:.82rem}}
th{{background:#fff;color:#94a3b8;font-weight:600;font-size:.7rem;letter-spacing:.5px;
  padding:10px 8px;text-align:center;border-bottom:1px solid #e2e8f0}}
td{{padding:11px 8px;text-align:center;border-bottom:1px solid rgba(0,0,0,.04)}}
tr:last-child td{{border-bottom:none}}
.year-cell{{text-align:left;padding-left:16px;white-space:nowrap;display:flex;align-items:center;gap:8px;min-width:80px}}
.year-bar{{display:inline-block;width:3px;height:22px;border-radius:2px;flex-shrink:0}}
.total-cell{{font-size:.85rem;font-weight:600;text-align:right;padding-right:16px}}
th:first-child{{text-align:left;padding-left:16px}}th:last-child{{text-align:right;padding-right:16px}}
</style>
</head>
<body>
<h2>BB8 · Monthly Track Record</h2>
<div class="wrap">
<table>
<thead><tr><th>DATE</th>{header_cols}<th>TOTAL</th></tr></thead>
<tbody>{rows_html_light}</tbody>
</table>
</div>
<p style="margin-top:10px;font-size:.7rem;color:#94a3b8">
  Snapshot taken 1st of each month → column = performance of previous month. Total = (Jan snapshot year N+1) / (Jan snapshot year N) − 1.
</p>
</body></html>"""

with open('track_record.html', 'w', encoding='utf-8') as f:
    f.write(TR_HTML)
print('Generated: track_record.html')

# ── Inject all into bb8_dashboard.html ────────────────────────────────────────
tr_header_html = ''.join(f'<th>{c}</th>' for c in COL_NAMES)
html_out = HTML.replace('__DATA_PLACEHOLDER__', json_str)
html_out = html_out.replace('__TR_HEADER__', tr_header_html)
html_out = html_out.replace('__TR_ROWS__',   tr_rows_dark)

with open('bb8_dashboard.html', 'w', encoding='utf-8') as f:
    f.write(html_out)

size_kb = len(html_out.encode()) / 1024
print(f'\nGenerated: bb8_dashboard.html  ({size_kb:.0f} KB)')
print(f'Months: {MONTHS}')
print(f'Total positions: {sum(len(all_data[m]["positions"]) for m in MONTHS)}')
