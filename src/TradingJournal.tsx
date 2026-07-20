import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import {
  AreaChart, Area, LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine, Legend
} from "recharts";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, ListChecks, CalendarDays, Brain, MessageCircle, Settings,
  Upload, Plus, X, TrendingUp, TrendingDown, Flame, AlertTriangle, ChevronDown,
  ChevronUp, Trash2, Edit3, Image as ImageIcon, Send, Sparkles, Target,
  Clock, Activity, Shield, Percent, BarChart3, Sun, Moon, Sunrise, Sunset,
  Filter, Download, RefreshCw, Award, Zap, Eye, EyeOff, Search,
  Layers, GitCompare, Gauge, FlaskConical, Coins, History, Tags,
  Lock, Mail, LogOut, KeyRound, CheckCircle2, ArrowLeft
} from "lucide-react";

/* ============================== DESIGN TOKENS ============================== */
const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap');`;

const GlobalStyle = () => (
  <style>{`
    ${FONT_IMPORT}
    .tj-root { font-family: 'Inter', sans-serif; color: #F2F1ED; background: #08090B; position: relative; min-height: 100vh; }
    .tj-root * { box-sizing: border-box; }
    .tj-display { font-family: 'Space Grotesk', sans-serif; }
    .tj-mono { font-family: 'JetBrains Mono', monospace; }
    .tj-bg-glow {
      position: absolute; inset: 0; pointer-events: none; z-index: 0;
      background:
        radial-gradient(600px 400px at 15% -5%, rgba(227,178,77,0.10), transparent 60%),
        radial-gradient(500px 400px at 90% 10%, rgba(52,211,153,0.06), transparent 60%),
        radial-gradient(700px 500px at 50% 110%, rgba(244,91,105,0.05), transparent 60%);
    }
    .tj-panel {
      background: linear-gradient(180deg, rgba(255,255,255,0.045), rgba(255,255,255,0.02));
      border: 1px solid rgba(255,255,255,0.09);
      border-radius: 16px;
      backdrop-filter: blur(20px);
      -webkit-backdrop-filter: blur(20px);
      box-shadow: 0 8px 32px rgba(0,0,0,0.35), inset 0 1px 0 rgba(255,255,255,0.05);
      position: relative;
    }
    .tj-panel-tight { border-radius: 12px; }
    .tj-sidebar {
      background: linear-gradient(180deg, rgba(255,255,255,0.035), rgba(255,255,255,0.01));
      border-right: 1px solid rgba(255,255,255,0.08);
      backdrop-filter: blur(24px);
    }
    .tj-navitem {
      display: flex; align-items: center; gap: 10px;
      padding: 10px 14px; border-radius: 10px; cursor: pointer;
      color: #8A8D97; font-size: 13.5px; font-weight: 500;
      transition: all .18s ease; border: 1px solid transparent;
      white-space: nowrap;
    }
    .tj-navitem:hover { color: #F2F1ED; background: rgba(255,255,255,0.045); }
    .tj-navitem.active {
      color: #E3B24D; background: rgba(227,178,77,0.10);
      border-color: rgba(227,178,77,0.25);
      box-shadow: inset 0 0 0 1px rgba(227,178,77,0.08);
    }
    .tj-gold { color: #E3B24D; }
    .tj-green { color: #34D399; }
    .tj-red { color: #F45B69; }
    .tj-dim { color: #8A8D97; }
    .tj-btn {
      display: inline-flex; align-items: center; gap: 8px;
      padding: 9px 16px; border-radius: 10px; font-size: 13px; font-weight: 600;
      cursor: pointer; transition: all .15s ease; border: 1px solid transparent;
      font-family: 'Inter', sans-serif;
    }
    .tj-btn-primary {
      background: linear-gradient(135deg, #E3B24D, #C08E2E);
      color: #0A0906; box-shadow: 0 4px 16px rgba(227,178,77,0.25);
    }
    .tj-btn-primary:hover { filter: brightness(1.08); transform: translateY(-1px); }
    .tj-btn-ghost {
      background: rgba(255,255,255,0.04); color: #F2F1ED;
      border-color: rgba(255,255,255,0.1);
    }
    .tj-btn-ghost:hover { background: rgba(255,255,255,0.08); }
    .tj-btn-danger { background: rgba(244,91,105,0.12); color: #F45B69; border-color: rgba(244,91,105,0.25); }
    .tj-btn-danger:hover { background: rgba(244,91,105,0.2); }
    .tj-input {
      background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.1);
      border-radius: 9px; padding: 9px 12px; color: #F2F1ED; font-size: 13px;
      outline: none; width: 100%; transition: border-color .15s ease;
      font-family: 'Inter', sans-serif;
    }
    .tj-input:focus { border-color: rgba(227,178,77,0.5); }
    .tj-input::placeholder { color: #5C5F68; }
    .tj-tag {
      display: inline-flex; align-items: center; gap: 4px;
      padding: 3px 9px; border-radius: 20px; font-size: 11px; font-weight: 600;
      background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.1); color: #C9CBD3;
    }
    .tj-scroll::-webkit-scrollbar { width: 6px; height: 6px; }
    .tj-scroll::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 10px; }
    .tj-scroll::-webkit-scrollbar-track { background: transparent; }
    .tj-fade-in { animation: tjFadeIn .35s ease both; }
    @keyframes tjFadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
    .tj-pulse { animation: tjPulse 2s ease-in-out infinite; }
    @keyframes tjPulse { 0%,100% { opacity: 1; } 50% { opacity: .5; } }
    .tj-row-hover:hover { background: rgba(255,255,255,0.03); }
    .tj-glow-line { filter: drop-shadow(0 0 6px rgba(227,178,77,0.5)); }
    .tj-modal-backdrop { background: rgba(4,4,5,0.7); backdrop-filter: blur(6px); }
    .tj-heat-cell { border-radius: 8px; transition: transform .15s ease; cursor: default; }
    .tj-heat-cell:hover { transform: scale(1.08); z-index: 2; }
    .tj-checklist-item { display:flex; align-items:center; gap:8px; padding:6px 0; font-size:13px; }
    .tj-focus-ring:focus-visible { outline: 2px solid #E3B24D; outline-offset: 2px; }
  `}</style>
);

/* ============================== CONSTANTS ============================== */
const CONTRACT_MULTIPLIER = 100; // XAUUSD: $ per 1.0 price move per 1.0 lot
const SETUPS = ["London Breakout", "NY Breakout", "Asia Range Fade", "Retest & Go", "Liquidity Sweep", "News Fade", "Trend Continuation"];
const EMOTIONS = ["Calm", "Confident", "Anxious", "Greedy", "Fearful", "Frustrated", "Impatient", "Focused"];
const TAG_OPTIONS = ["Rule Break", "A+ Setup", "Late Entry", "Early Exit", "Oversized", "Perfect Execution", "Chased Price"];

const SESSIONS = [
  { name: "Sydney/Asia", start: 0, end: 8, icon: Moon, color: "#8B7FD9" },
  { name: "London", start: 7, end: 16, icon: Sunrise, color: "#5FB4E8" },
  { name: "New York", start: 12, end: 21, icon: Sun, color: "#E3B24D" },
  { name: "Off-Session", start: 21, end: 24, icon: Sunset, color: "#5C5F68" },
];

function getSession(hourUTC) {
  if (hourUTC >= 12 && hourUTC < 16) return "London/NY Overlap";
  if (hourUTC >= 7 && hourUTC < 12) return "London";
  if (hourUTC >= 16 && hourUTC < 21) return "New York";
  if (hourUTC >= 0 && hourUTC < 7) return "Asia";
  return "Off-Session";
}

const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

/* ============================== DEMO SEED DATA ============================== */
function genSeedTrades() {
  const trades = [];
  let id = 1;
  const start = new Date("2026-03-02T00:00:00Z");
  let px = 2085;
  const setupsW = SETUPS;
  for (let d = 0; d < 95; d++) {
    const day = new Date(start.getTime() + d * 86400000);
    const wd = day.getUTCDay();
    if (wd === 0 || wd === 6) continue; // skip weekends
    const numTrades = Math.random() < 0.55 ? (Math.random() < 0.7 ? 1 : 2) : (Math.random() < 0.5 ? 0 : 3);
    for (let t = 0; t < numTrades; t++) {
      const hour = [1, 3, 7, 8, 9, 13, 14, 15, 16, 19][Math.floor(Math.random() * 10)];
      const entryTime = new Date(day);
      entryTime.setUTCHours(hour, Math.floor(Math.random() * 60), 0, 0);
      const side = Math.random() < 0.55 ? "buy" : "sell";
      const lots = [0.1, 0.2, 0.25, 0.5, 0.5, 1][Math.floor(Math.random() * 6)];
      const entryPrice = +(px + (Math.random() - 0.5) * 6).toFixed(2);
      const riskPts = +(3 + Math.random() * 5).toFixed(2);
      const sl = side === "buy" ? +(entryPrice - riskPts).toFixed(2) : +(entryPrice + riskPts).toFixed(2);
      const winProb = 0.5;
      const isWin = Math.random() < winProb;
      let exitPrice, tp;
      const rr = 1.5 + Math.random() * 2;
      tp = side === "buy" ? +(entryPrice + riskPts * rr).toFixed(2) : +(entryPrice - riskPts * rr).toFixed(2);
      if (isWin) {
        const frac = 0.5 + Math.random() * 0.6;
        exitPrice = side === "buy" ? +(entryPrice + riskPts * rr * Math.min(frac, 1.05)).toFixed(2) : +(entryPrice - riskPts * rr * Math.min(frac, 1.05)).toFixed(2);
      } else {
        const frac = 0.6 + Math.random() * 0.8;
        exitPrice = side === "buy" ? +(entryPrice - riskPts * Math.min(frac, 1.4)).toFixed(2) : +(entryPrice + riskPts * Math.min(frac, 1.4)).toFixed(2);
      }
      const durMin = Math.floor(5 + Math.random() * 180);
      const exitTime = new Date(entryTime.getTime() + durMin * 60000);
      const commission = +(lots * 7).toFixed(2);
      const swap = +((Math.random() - 0.5) * lots * 4).toFixed(2);
      const pnl = +(((side === "buy" ? exitPrice - entryPrice : entryPrice - exitPrice) * lots * CONTRACT_MULTIPLIER) - commission - Math.abs(swap)).toFixed(2);
      trades.push({
        id: id++, symbol: "XAUUSD", side, lots,
        entryPrice, exitPrice, sl, tp,
        entryTime: entryTime.toISOString(), exitTime: exitTime.toISOString(),
        commission, swap, pnl,
        setup: setupsW[Math.floor(Math.random() * setupsW.length)],
        emotion: EMOTIONS[Math.floor(Math.random() * EMOTIONS.length)],
        tags: Math.random() < 0.2 ? [TAG_OPTIONS[Math.floor(Math.random() * TAG_OPTIONS.length)]] : [],
        notes: "", screenshots: [],
      });
      px = exitPrice;
    }
  }
  return trades;
}

/* ============================== METRICS ENGINE ============================== */
function computePnl(t) {
  if (typeof t.pnl === "number" && !isNaN(t.pnl)) return t.pnl;
  const raw = (t.side === "buy" ? t.exitPrice - t.entryPrice : t.entryPrice - t.exitPrice) * (t.lots || 0) * CONTRACT_MULTIPLIER;
  return +(raw - (t.commission || 0) - Math.abs(t.swap || 0)).toFixed(2);
}
function riskDollars(t) {
  if (!t.sl) return null;
  return Math.abs(t.entryPrice - t.sl) * (t.lots || 0) * CONTRACT_MULTIPLIER;
}
function rMultiple(t) {
  const r = riskDollars(t);
  if (!r || r === 0) return null;
  return computePnl(t) / r;
}
function durationMin(t) {
  if (!t.entryTime || !t.exitTime) return null;
  return Math.max(0, (new Date(t.exitTime) - new Date(t.entryTime)) / 60000);
}
function stdev(arr) {
  if (arr.length < 2) return 0;
  const m = arr.reduce((a, b) => a + b, 0) / arr.length;
  const v = arr.reduce((a, b) => a + (b - m) ** 2, 0) / (arr.length - 1);
  return Math.sqrt(v);
}

function computeMetrics(trades) {
  const sorted = [...trades].sort((a, b) => new Date(a.entryTime) - new Date(b.entryTime));
  const n = sorted.length;
  const pnls = sorted.map(computePnl);
  const wins = pnls.filter(p => p > 0);
  const losses = pnls.filter(p => p < 0);
  const grossProfit = wins.reduce((a, b) => a + b, 0);
  const grossLoss = losses.reduce((a, b) => a + b, 0);
  const netProfit = grossProfit + grossLoss;
  const winRate = n ? wins.length / n : 0;
  const lossRate = n ? losses.length / n : 0;
  const avgWin = wins.length ? grossProfit / wins.length : 0;
  const avgLoss = losses.length ? grossLoss / losses.length : 0;
  const largestWin = wins.length ? Math.max(...wins) : 0;
  const largestLoss = losses.length ? Math.min(...losses) : 0;
  const profitFactor = grossLoss !== 0 ? Math.abs(grossProfit / grossLoss) : (grossProfit > 0 ? Infinity : 0);
  const expectancy = n ? netProfit / n : 0;
  const avgWinLossRatio = avgLoss !== 0 ? Math.abs(avgWin / avgLoss) : 0;

  // equity curve
  let bal = 0; const equityCurve = [{ x: 0, y: 0, date: "Start" }];
  let peak = 0, maxDD = 0, maxDDPct = 0;
  const ddSeries = [];
  sorted.forEach((t, i) => {
    bal += pnls[i];
    peak = Math.max(peak, bal);
    const dd = peak - bal;
    maxDD = Math.max(maxDD, dd);
    ddSeries.push({ x: i + 1, dd: -dd });
    equityCurve.push({ x: i + 1, y: +bal.toFixed(2), date: t.entryTime ? t.entryTime.slice(0, 10) : i });
  });
  const recoveryFactor = maxDD > 0 ? netProfit / maxDD : (netProfit > 0 ? Infinity : 0);

  // streaks
  let curWin = 0, curLoss = 0, maxWinStreak = 0, maxLossStreak = 0;
  pnls.forEach(p => {
    if (p > 0) { curWin++; curLoss = 0; maxWinStreak = Math.max(maxWinStreak, curWin); }
    else if (p < 0) { curLoss++; curWin = 0; maxLossStreak = Math.max(maxLossStreak, curLoss); }
  });
  const currentStreak = (() => {
    let c = 0, sign = 0;
    for (let i = pnls.length - 1; i >= 0; i--) {
      const s = pnls[i] > 0 ? 1 : pnls[i] < 0 ? -1 : 0;
      if (i === pnls.length - 1) { sign = s; c = s === 0 ? 0 : 1; }
      else if (s === sign && s !== 0) c++;
      else break;
    }
    return { count: c, sign };
  })();

  // R multiples / SQN
  const rVals = sorted.map(rMultiple).filter(v => v !== null);
  const rMean = rVals.length ? rVals.reduce((a, b) => a + b, 0) / rVals.length : 0;
  const rStd = stdev(rVals);
  const sqn = rVals.length && rStd > 0 ? Math.sqrt(rVals.length) * (rMean / rStd) : 0;

  // sharpe / sortino (per-trade return basis)
  const pnlMean = pnls.length ? pnls.reduce((a, b) => a + b, 0) / pnls.length : 0;
  const pnlStd = stdev(pnls);
  const downside = pnls.filter(p => p < 0);
  const downsideStd = downside.length ? Math.sqrt(downside.reduce((a, b) => a + b * b, 0) / downside.length) : 0;
  const sharpe = pnlStd > 0 ? (pnlMean / pnlStd) * Math.sqrt(n) : 0;
  const sortino = downsideStd > 0 ? (pnlMean / downsideStd) * Math.sqrt(n) : 0;
  const calmar = maxDD > 0 ? netProfit / maxDD : 0;

  // kelly & risk of ruin (approximate)
  const kellyPct = avgWinLossRatio > 0 ? (winRate - (1 - winRate) / avgWinLossRatio) * 100 : 0;
  const edge = 2 * winRate - 1;
  const riskOfRuin = edge > 0 ? Math.max(0, Math.min(100, ((1 - edge) / (1 + edge)) ** 10 * 100)) : 100;

  // avg duration win vs loss
  const winDurs = sorted.filter(t => computePnl(t) > 0).map(durationMin).filter(v => v !== null);
  const lossDurs = sorted.filter(t => computePnl(t) < 0).map(durationMin).filter(v => v !== null);
  const avgWinDur = winDurs.length ? winDurs.reduce((a, b) => a + b, 0) / winDurs.length : 0;
  const avgLossDur = lossDurs.length ? lossDurs.reduce((a, b) => a + b, 0) / lossDurs.length : 0;

  // buy vs sell
  const buyTrades = sorted.filter(t => t.side === "buy");
  const sellTrades = sorted.filter(t => t.side === "sell");
  const buyWinRate = buyTrades.length ? buyTrades.filter(t => computePnl(t) > 0).length / buyTrades.length : 0;
  const sellWinRate = sellTrades.length ? sellTrades.filter(t => computePnl(t) > 0).length / sellTrades.length : 0;
  const buyPnl = buyTrades.reduce((a, t) => a + computePnl(t), 0);
  const sellPnl = sellTrades.reduce((a, t) => a + computePnl(t), 0);

  return {
    n, netProfit, grossProfit, grossLoss, winRate, lossRate, avgWin, avgLoss,
    largestWin, largestLoss, profitFactor, expectancy, avgWinLossRatio,
    equityCurve, maxDD, maxDDPct: peak > 0 ? (maxDD / peak) * 100 : 0, recoveryFactor,
    maxWinStreak, maxLossStreak, currentStreak, rVals, rMean, rStd, sqn,
    sharpe, sortino, calmar, kellyPct, riskOfRuin, avgWinDur, avgLossDur,
    buyTrades, sellTrades, buyWinRate, sellWinRate, buyPnl, sellPnl, ddSeries,
    sorted,
  };
}

/* ============================== PSYCHOLOGY DETECTION ============================== */
const DEFAULT_PSYCH_CONFIG = {
  revengeGapMin: 20, revengeLotMultiplier: 1.15,
  overtradingCount: 4,
  cuttingWinnersFraction: 0.35,
  holdingLosersMultiplier: 2.5,
  sessionStartHour: 7, sessionEndHour: 21,
};

const RULE_FIELDS = [
  { key: "pnl", label: "P&L ($)" },
  { key: "lots", label: "Lot Size" },
  { key: "durationMin", label: "Duration (min)" },
  { key: "rMultiple", label: "R-Multiple" },
  { key: "entryHourUTC", label: "Entry Hour (UTC)" },
  { key: "setup", label: "Setup (text)" },
  { key: "emotion", label: "Emotion (text)" },
  { key: "tags", label: "Tags (contains)" },
  { key: "side", label: "Side (buy/sell)" },
];
const RULE_OPERATORS = [
  { key: "gt", label: ">" }, { key: "gte", label: "≥" },
  { key: "lt", label: "<" }, { key: "lte", label: "≤" },
  { key: "eq", label: "=" }, { key: "contains", label: "contains" },
];

function evalCustomRule(t, rule) {
  let val;
  switch (rule.field) {
    case "pnl": val = computePnl(t); break;
    case "lots": val = t.lots; break;
    case "durationMin": val = durationMin(t); break;
    case "rMultiple": val = rMultiple(t); break;
    case "entryHourUTC": val = t.entryTime ? new Date(t.entryTime).getUTCHours() : null; break;
    case "setup": val = t.setup || ""; break;
    case "emotion": val = t.emotion || ""; break;
    case "tags": val = (t.tags || []).join(","); break;
    case "side": val = t.side || ""; break;
    default: val = null;
  }
  if (val === null || val === undefined) return false;
  const target = rule.operator === "contains" ? rule.value.toString().toLowerCase() : parseFloat(rule.value);
  switch (rule.operator) {
    case "gt": return parseFloat(val) > target;
    case "gte": return parseFloat(val) >= target;
    case "lt": return parseFloat(val) < target;
    case "lte": return parseFloat(val) <= target;
    case "eq": return val.toString().toLowerCase() === rule.value.toString().toLowerCase();
    case "contains": return val.toString().toLowerCase().includes(target);
    default: return false;
  }
}

function detectPsychology(trades, config = DEFAULT_PSYCH_CONFIG, customRules = []) {
  const cfg = { ...DEFAULT_PSYCH_CONFIG, ...config };
  const sorted = [...trades].filter(t => t.entryTime).sort((a, b) => new Date(a.entryTime) - new Date(b.entryTime));
  const flags = [];
  const metrics = computeMetrics(trades);
  const dateOf = (t) => t.entryTime.slice(0, 10);

  // revenge trading
  for (let i = 1; i < sorted.length; i++) {
    const prev = sorted[i - 1], cur = sorted[i];
    if (!prev.exitTime) continue;
    const gapMin = (new Date(cur.entryTime) - new Date(prev.exitTime)) / 60000;
    if (computePnl(prev) < 0 && gapMin >= 0 && gapMin < cfg.revengeGapMin && cur.lots > prev.lots * cfg.revengeLotMultiplier) {
      flags.push({ type: "Revenge Trading", tradeId: cur.id, severity: "high", date: dateOf(cur),
        detail: `Re-entered ${gapMin.toFixed(0)}min after a loss with lot size increased from ${prev.lots} to ${cur.lots}.` });
    }
  }

  // overtrading days
  const byDay = {};
  sorted.forEach(t => { const d = dateOf(t); (byDay[d] = byDay[d] || []).push(t); });
  Object.entries(byDay).forEach(([d, ts]) => {
    if (ts.length >= cfg.overtradingCount) flags.push({ type: "Overtrading", tradeId: null, day: d, date: d, severity: "medium",
      detail: `${ts.length} trades placed on ${d}, above your ${cfg.overtradingCount}-trade daily threshold.` });
  });

  // cutting winners early
  sorted.forEach(t => {
    if (!t.tp || !t.sl) return;
    const pnl = computePnl(t);
    if (pnl <= 0) return;
    const plannedDist = Math.abs(t.tp - t.entryPrice);
    const gotDist = Math.abs(t.exitPrice - t.entryPrice);
    if (plannedDist > 0 && gotDist / plannedDist < cfg.cuttingWinnersFraction) {
      flags.push({ type: "Cutting Winners Short", tradeId: t.id, severity: "medium", date: dateOf(t),
        detail: `Closed at only ${((gotDist / plannedDist) * 100).toFixed(0)}% of planned distance to TP.` });
    }
  });

  // holding losers
  const avgWinDur = metrics.avgWinDur || 30;
  sorted.forEach(t => {
    const pnl = computePnl(t);
    const dur = durationMin(t);
    if (pnl < 0 && dur !== null && avgWinDur > 0 && dur > avgWinDur * cfg.holdingLosersMultiplier) {
      flags.push({ type: "Holding Losers Too Long", tradeId: t.id, severity: "high", date: dateOf(t),
        detail: `Loss held for ${dur.toFixed(0)}min, vs an average winner duration of ${avgWinDur.toFixed(0)}min.` });
    }
  });

  // no plan / rule break tags
  sorted.forEach(t => {
    if (!t.setup) flags.push({ type: "No Defined Setup", tradeId: t.id, severity: "low", date: dateOf(t),
      detail: "Trade taken without a tagged setup — a common sign of impulsive or FOMO entries." });
    if ((t.tags || []).includes("Rule Break")) flags.push({ type: "Rule Break Logged", tradeId: t.id, severity: "high", date: dateOf(t),
      detail: "Trade self-tagged as a rule break." });
    if ((t.tags || []).includes("Oversized")) flags.push({ type: "Position Oversized", tradeId: t.id, severity: "medium", date: dateOf(t),
      detail: "Trade self-tagged as oversized relative to plan." });
  });

  // outside session
  sorted.forEach(t => {
    const hr = new Date(t.entryTime).getUTCHours();
    if (hr < cfg.sessionStartHour || hr >= cfg.sessionEndHour) {
      flags.push({ type: "Trading Outside Core Session", tradeId: t.id, severity: "low", date: dateOf(t),
        detail: `Entered at ${String(hr).padStart(2, "0")}:00 UTC, outside your ${cfg.sessionStartHour}:00–${cfg.sessionEndHour}:00 window.` });
    }
  });

  // user-defined custom rules
  (customRules || []).forEach(rule => {
    sorted.forEach(t => {
      try {
        if (evalCustomRule(t, rule)) {
          flags.push({ type: rule.label || "Custom Rule", tradeId: t.id, severity: rule.severity || "medium", date: dateOf(t),
            detail: `Matched your rule: ${RULE_FIELDS.find(f => f.key === rule.field)?.label || rule.field} ${RULE_OPERATORS.find(o => o.key === rule.operator)?.label || rule.operator} ${rule.value}.` });
        }
      } catch (e) { /* skip bad rule */ }
    });
  });

  const counts = {};
  flags.forEach(f => { counts[f.type] = (counts[f.type] || 0) + 1; });
  return { flags, counts };
}

/* ============================== DEEP ANALYSIS ENGINE ============================== */

// 1 & 2. Generic "group trades by a key function, compute stats per group" — powers setup, tag, and custom-column breakdowns
function groupStats(trades, keyFn) {
  const groups = {};
  trades.forEach(t => {
    const keys = keyFn(t);
    (Array.isArray(keys) ? keys : [keys]).forEach(k => {
      if (k === null || k === undefined || k === "") return;
      if (!groups[k]) groups[k] = [];
      groups[k].push(t);
    });
  });
  return Object.entries(groups).map(([key, list]) => {
    const pnls = list.map(computePnl);
    const wins = pnls.filter(p => p > 0);
    const losses = pnls.filter(p => p < 0);
    const net = pnls.reduce((a, b) => a + b, 0);
    const grossLoss = losses.reduce((a, b) => a + b, 0);
    const grossProfit = wins.reduce((a, b) => a + b, 0);
    return {
      key, count: list.length,
      winRate: list.length ? wins.length / list.length : 0,
      expectancy: list.length ? net / list.length : 0,
      profitFactor: grossLoss !== 0 ? Math.abs(grossProfit / grossLoss) : (grossProfit > 0 ? Infinity : 0),
      netPnl: net,
    };
  }).sort((a, b) => b.expectancy - a.expectancy);
}

function computeSetupTagStats(trades) {
  const bySetup = groupStats(trades, t => t.setup || null);
  const byTag = groupStats(trades, t => (t.tags && t.tags.length ? t.tags : null));
  return { bySetup, byTag };
}

function computeCustomColumnStats(trades, customColumns) {
  return customColumns.map(col => ({
    column: col,
    stats: groupStats(trades, t => (t.customFields && t.customFields[col.id]) || null).slice(0, 10),
  })).filter(c => c.stats.length > 0);
}

// 3. MFE/MAE efficiency — uses optional mfe/mae $ fields logged per trade
function computeEfficiency(trades) {
  const withData = trades.filter(t => typeof t.mfe === "number" && t.mfe > 0);
  if (!withData.length) return { available: false, avgEfficiency: 0, list: [] };
  const list = withData.map(t => {
    const pnl = computePnl(t);
    const efficiency = t.mfe > 0 ? clamp(pnl / t.mfe, -2, 1) : 0;
    return { id: t.id, symbol: t.symbol, pnl, mfe: t.mfe, mae: t.mae || 0, efficiency };
  });
  const avgEfficiency = list.reduce((a, b) => a + b.efficiency, 0) / list.length;
  return { available: true, avgEfficiency, list: list.sort((a, b) => a.efficiency - b.efficiency) };
}

// 4. Rolling / edge-decay stats over trade sequence
function computeRollingStats(trades, windowSize = 15) {
  const sorted = [...trades].filter(t => t.entryTime).sort((a, b) => new Date(a.entryTime) - new Date(b.entryTime));
  const pnls = sorted.map(computePnl);
  const out = [];
  for (let i = 0; i < pnls.length; i++) {
    const start = Math.max(0, i - windowSize + 1);
    const window = pnls.slice(start, i + 1);
    const wins = window.filter(p => p > 0).length;
    out.push({
      x: i + 1,
      winRate: +(wins / window.length * 100).toFixed(1),
      expectancy: +(window.reduce((a, b) => a + b, 0) / window.length).toFixed(2),
    });
  }
  return out;
}

// 5. Day-after-loss / day-after-win behavior
function computeDayAfterStats(trades) {
  const byDay = {};
  trades.forEach(t => { if (!t.entryTime) return; const d = t.entryTime.slice(0, 10); (byDay[d] = byDay[d] || []).push(t); });
  const days = Object.keys(byDay).sort();
  const buckets = { afterLoss: [], afterWin: [], afterFlat: [] };
  for (let i = 1; i < days.length; i++) {
    const prevDayPnl = byDay[days[i - 1]].reduce((a, t) => a + computePnl(t), 0);
    const bucket = prevDayPnl < 0 ? "afterLoss" : prevDayPnl > 0 ? "afterWin" : "afterFlat";
    buckets[bucket].push(...byDay[days[i]]);
  }
  const summarize = (list) => {
    if (!list.length) return { count: 0, winRate: 0, expectancy: 0, netPnl: 0 };
    const pnls = list.map(computePnl);
    const wins = pnls.filter(p => p > 0).length;
    const net = pnls.reduce((a, b) => a + b, 0);
    return { count: list.length, winRate: wins / list.length, expectancy: net / list.length, netPnl: net };
  };
  return {
    afterLoss: summarize(buckets.afterLoss),
    afterWin: summarize(buckets.afterWin),
    afterFlat: summarize(buckets.afterFlat),
  };
}

// 6. Discipline / consistency score — composite of flag density, tag hygiene, sizing consistency
function computeDisciplineScore(trades, flags) {
  if (!trades.length) return { score: 100, trend: [] };
  const flagPenalty = clamp(100 - (flags.length / trades.length) * 40, 0, 100);
  const lots = trades.map(t => t.lots).filter(l => typeof l === "number" && l > 0);
  const lotMean = lots.length ? lots.reduce((a, b) => a + b, 0) / lots.length : 0;
  const lotCV = lotMean > 0 ? stdev(lots) / lotMean : 0; // coefficient of variation
  const sizingScore = clamp(100 - lotCV * 100, 0, 100);
  const ruleBreaks = trades.filter(t => (t.tags || []).includes("Rule Break")).length;
  const ruleScore = clamp(100 - (ruleBreaks / trades.length) * 150, 0, 100);
  const score = Math.round(flagPenalty * 0.45 + sizingScore * 0.30 + ruleScore * 0.25);

  // weekly trend
  const byWeek = {};
  trades.forEach(t => { if (!t.entryTime) return; const wk = weekKey(t.entryTime.slice(0, 10)); (byWeek[wk] = byWeek[wk] || []).push(t); });
  const trend = Object.entries(byWeek).sort((a, b) => a[0].localeCompare(b[0])).map(([wk, list]) => {
    const wkFlags = flags.filter(f => list.some(t => t.id === f.tradeId));
    const fp = clamp(100 - (wkFlags.length / list.length) * 40, 0, 100);
    const wkLots = list.map(t => t.lots).filter(l => typeof l === "number" && l > 0);
    const wkMean = wkLots.length ? wkLots.reduce((a, b) => a + b, 0) / wkLots.length : 0;
    const wkCV = wkMean > 0 ? stdev(wkLots) / wkMean : 0;
    const sz = clamp(100 - wkCV * 100, 0, 100);
    return { week: wk, score: Math.round(fp * 0.6 + sz * 0.4) };
  });
  return { score, flagPenalty, sizingScore, ruleScore, trend };
}

// 7. Cost drag — commission + swap as a share of gross profit
function computeCostDrag(trades) {
  const totalCommission = trades.reduce((a, t) => a + Math.abs(t.commission || 0), 0);
  const totalSwap = trades.reduce((a, t) => a + Math.abs(t.swap || 0), 0);
  const totalCost = totalCommission + totalSwap;
  const grossProfit = trades.map(computePnl).filter(p => p > 0).reduce((a, b) => a + b, 0);
  const pctOfGrossProfit = grossProfit > 0 ? (totalCost / grossProfit) * 100 : 0;
  return { totalCommission, totalSwap, totalCost, pctOfGrossProfit };
}

// 8. What-if simulator — recompute equity with certain flagged trades excluded
function computeWhatIf(trades, flags, excludedTypes) {
  const excludedIds = new Set(flags.filter(f => excludedTypes.includes(f.type) && f.tradeId).map(f => f.tradeId));
  const cleanTrades = trades.filter(t => !excludedIds.has(t.id));
  const actual = computeMetrics(trades);
  const hypothetical = computeMetrics(cleanTrades);
  return { actual, hypothetical, excludedCount: trades.length - cleanTrades.length, delta: hypothetical.netProfit - actual.netProfit };
}

// 9. Period comparison — current vs previous window of the same length
function computePeriodComparison(trades, periodType) {
  const now = new Date();
  const msDay = 86400000;
  const spanMs = periodType === "week" ? 7 * msDay : 30 * msDay;
  const currentStart = new Date(now.getTime() - spanMs);
  const prevStart = new Date(now.getTime() - spanMs * 2);
  const inRange = (t, from, to) => t.entryTime && new Date(t.entryTime) >= from && new Date(t.entryTime) < to;
  const current = trades.filter(t => inRange(t, currentStart, now));
  const previous = trades.filter(t => inRange(t, prevStart, currentStart));
  return { current: computeMetrics(current), previous: computeMetrics(previous) };
}
const fmt$ = (v) => (v < 0 ? "-$" : "$") + Math.abs(v).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
const fmtPct = (v) => `${(v * 100).toFixed(1)}%`;
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

function saveTrades(trades, journalId) {
  try { window.storage?.set(`trades:${journalId || "default"}`, JSON.stringify(trades), false); } catch (e) { /* noop */ }
}

/* ============================== UI ATOMS ============================== */
function StatCard({ label, value, sub, icon: Icon, tone = "neutral", mono = true }) {
  const toneColor = tone === "good" ? "#34D399" : tone === "bad" ? "#F45B69" : tone === "gold" ? "#E3B24D" : "#F2F1ED";
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: "16px 18px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div style={{ fontSize: 11.5, color: "#8A8D97", fontWeight: 600, letterSpacing: "0.03em", textTransform: "uppercase" }}>{label}</div>
        {Icon && <Icon size={15} color="#5C5F68" />}
      </div>
      <div className={mono ? "tj-mono" : "tj-display"} style={{ fontSize: 24, fontWeight: 600, marginTop: 8, color: toneColor }}>{value}</div>
      {sub && <div style={{ fontSize: 11.5, color: "#6B6E78", marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

function SectionTitle({ children, icon: Icon, right }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        {Icon && <Icon size={16} className="tj-gold" />}
        <h3 className="tj-display" style={{ fontSize: 15, fontWeight: 600, letterSpacing: "0.01em" }}>{children}</h3>
      </div>
      {right}
    </div>
  );
}

function Pill({ children, tone = "neutral" }) {
  const styles = {
    good: { background: "rgba(52,211,153,0.12)", color: "#34D399", border: "1px solid rgba(52,211,153,0.3)" },
    bad: { background: "rgba(244,91,105,0.12)", color: "#F45B69", border: "1px solid rgba(244,91,105,0.3)" },
    warn: { background: "rgba(227,178,77,0.12)", color: "#E3B24D", border: "1px solid rgba(227,178,77,0.3)" },
    neutral: { background: "rgba(255,255,255,0.06)", color: "#C9CBD3", border: "1px solid rgba(255,255,255,0.1)" },
  };
  return <span style={{ ...styles[tone], padding: "2px 9px", borderRadius: 20, fontSize: 11, fontWeight: 700 }}>{children}</span>;
}

/* ============================== CHARTS ============================== */
function CustomTooltip({ active, payload, label, prefix = "$" }) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div className="tj-panel tj-panel-tight" style={{ padding: "8px 12px", fontSize: 12 }}>
      <div className="tj-dim" style={{ marginBottom: 2 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} className="tj-mono" style={{ color: p.color || "#E3B24D", fontWeight: 600 }}>
          {prefix}{typeof p.value === "number" ? p.value.toFixed(2) : p.value}
        </div>
      ))}
    </div>
  );
}

function EquityHero({ metrics }) {
  const data = metrics.equityCurve;
  const isUp = metrics.netProfit >= 0;
  const color = isUp ? "#E3B24D" : "#F45B69";
  return (
    <div className="tj-panel tj-fade-in" style={{ padding: "24px 24px 8px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 8 }}>
        <div>
          <div className="tj-dim" style={{ fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>Account Equity Curve</div>
          <div className="tj-display tj-mono" style={{ fontSize: 36, fontWeight: 700, marginTop: 4, color }}>
            {fmt$(metrics.netProfit)}
          </div>
          <div style={{ fontSize: 12.5, color: "#8A8D97", marginTop: 2 }}>
            {metrics.n} trades · {fmtPct(metrics.winRate)} win rate · PF {metrics.profitFactor === Infinity ? "∞" : metrics.profitFactor.toFixed(2)}
          </div>
        </div>
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
          <MiniStat label="Max DD" value={fmt$(-metrics.maxDD)} tone="bad" />
          <MiniStat label="Expectancy" value={fmt$(metrics.expectancy)} tone={metrics.expectancy >= 0 ? "good" : "bad"} />
          <MiniStat label="Profit Factor" value={metrics.profitFactor === Infinity ? "∞" : metrics.profitFactor.toFixed(2)} tone="gold" />
          <MiniStat label="SQN" value={metrics.sqn.toFixed(2)} tone="gold" />
        </div>
      </div>
      <div style={{ height: 260 }}>
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart data={data} margin={{ top: 10, right: 10, left: -10, bottom: 0 }}>
            <defs>
              <linearGradient id="equityFill" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={color} stopOpacity={0.35} />
                <stop offset="100%" stopColor={color} stopOpacity={0} />
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" vertical={false} />
            <XAxis dataKey="x" tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} width={55}
              tickFormatter={(v) => `$${v}`} />
            <Tooltip content={<CustomTooltip />} />
            <ReferenceLine y={0} stroke="rgba(255,255,255,0.15)" />
            <Area type="monotone" dataKey="y" stroke={color} strokeWidth={2.5} fill="url(#equityFill)" className="tj-glow-line" />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function MiniStat({ label, value, tone }) {
  const color = tone === "good" ? "#34D399" : tone === "bad" ? "#F45B69" : tone === "gold" ? "#E3B24D" : "#F2F1ED";
  return (
    <div>
      <div style={{ fontSize: 10.5, color: "#6B6E78", fontWeight: 600, textTransform: "uppercase" }}>{label}</div>
      <div className="tj-mono" style={{ fontSize: 17, fontWeight: 700, color }}>{value}</div>
    </div>
  );
}

function DrawdownChart({ metrics }) {
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: 18 }}>
      <SectionTitle icon={Activity}>Drawdown Curve</SectionTitle>
      <div style={{ height: 140 }}>
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart data={metrics.ddSeries}>
            <defs>
              <linearGradient id="ddFill" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#F45B69" stopOpacity={0.05} />
                <stop offset="100%" stopColor="#F45B69" stopOpacity={0.35} />
              </linearGradient>
            </defs>
            <XAxis dataKey="x" hide />
            <YAxis tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} width={50} />
            <Tooltip content={<CustomTooltip />} />
            <Area type="monotone" dataKey="dd" stroke="#F45B69" strokeWidth={1.5} fill="url(#ddFill)" />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function WeekdayChart({ trades }) {
  const data = useMemo(() => {
    const byDay = WEEKDAYS.map((d, i) => ({ day: d, pnl: 0, count: 0 }));
    trades.forEach(t => {
      if (!t.entryTime) return;
      const wd = new Date(t.entryTime).getUTCDay();
      byDay[wd].pnl += computePnl(t);
      byDay[wd].count++;
    });
    return byDay.filter((_, i) => i >= 1 && i <= 5);
  }, [trades]);
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: 18 }}>
      <SectionTitle icon={CalendarDays}>Performance by Weekday</SectionTitle>
      <div style={{ height: 180 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={data}>
            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" vertical={false} />
            <XAxis dataKey="day" tick={{ fill: "#8A8D97", fontSize: 11 }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} width={45} />
            <Tooltip content={<CustomTooltip />} />
            <Bar dataKey="pnl" radius={[6, 6, 0, 0]}>
              {data.map((d, i) => <Cell key={i} fill={d.pnl >= 0 ? "#34D399" : "#F45B69"} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function SessionChart({ trades }) {
  const data = useMemo(() => {
    const map = {};
    trades.forEach(t => {
      if (!t.entryTime) return;
      const s = getSession(new Date(t.entryTime).getUTCHours());
      if (!map[s]) map[s] = { session: s, pnl: 0, count: 0, wins: 0 };
      map[s].pnl += computePnl(t);
      map[s].count++;
      if (computePnl(t) > 0) map[s].wins++;
    });
    return Object.values(map).map(d => ({ ...d, winRate: d.count ? (d.wins / d.count) * 100 : 0 }));
  }, [trades]);
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: 18 }}>
      <SectionTitle icon={Clock}>Session Analysis</SectionTitle>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {data.sort((a, b) => b.pnl - a.pnl).map((d, i) => (
          <div key={i}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 4 }}>
              <span style={{ fontWeight: 600 }}>{d.session}</span>
              <span className="tj-mono" style={{ color: d.pnl >= 0 ? "#34D399" : "#F45B69", fontWeight: 700 }}>{fmt$(d.pnl)}</span>
            </div>
            <div style={{ height: 6, background: "rgba(255,255,255,0.06)", borderRadius: 4, overflow: "hidden" }}>
              <div style={{ width: `${clamp(d.winRate, 0, 100)}%`, height: "100%", background: "linear-gradient(90deg,#E3B24D,#C08E2E)" }} />
            </div>
            <div style={{ fontSize: 10.5, color: "#6B6E78", marginTop: 2 }}>{d.count} trades · {d.winRate.toFixed(0)}% win rate</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function BuySellDonut({ metrics }) {
  const data = [
    { name: "Buy", value: metrics.buyTrades.length, pnl: metrics.buyPnl },
    { name: "Sell", value: metrics.sellTrades.length, pnl: metrics.sellPnl },
  ];
  const colors = ["#34D399", "#F45B69"];
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: 18 }}>
      <SectionTitle icon={BarChart3}>Buy vs Sell</SectionTitle>
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <div style={{ width: 120, height: 120 }}>
          <ResponsiveContainer>
            <PieChart>
              <Pie data={data} dataKey="value" innerRadius={35} outerRadius={55} paddingAngle={4}>
                {data.map((d, i) => <Cell key={i} fill={colors[i]} stroke="none" />)}
              </Pie>
            </PieChart>
          </ResponsiveContainer>
        </div>
        <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 8 }}>
          {data.map((d, i) => (
            <div key={i} style={{ fontSize: 12.5 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, fontWeight: 700, color: colors[i] }}>
                <span style={{ width: 8, height: 8, borderRadius: 4, background: colors[i] }} />
                {d.name} <span className="tj-dim" style={{ fontWeight: 500 }}>({d.value})</span>
              </div>
              <div className="tj-mono tj-dim" style={{ fontSize: 11, marginLeft: 14 }}>{fmt$(d.pnl)}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function RDistribution({ metrics }) {
  const buckets = [-3, -2, -1, 0, 1, 2, 3, 4, 5];
  const data = buckets.slice(0, -1).map((b, i) => {
    const upper = buckets[i + 1];
    const count = metrics.rVals.filter(r => r >= b && r < upper).length;
    return { label: `${b}R`, count };
  });
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: 18 }}>
      <SectionTitle icon={Target}>R-Multiple Distribution</SectionTitle>
      <div style={{ height: 160 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={data}>
            <XAxis dataKey="label" tick={{ fill: "#8A8D97", fontSize: 10 }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
            <Tooltip content={<CustomTooltip prefix="" />} />
            <Bar dataKey="count" radius={[4, 4, 0, 0]}>
              {data.map((d, i) => <Cell key={i} fill={buckets[i] < 0 ? "#F45B69" : "#34D399"} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

/* ============================== CALENDAR ============================== */
function TradeCalendar({ trades, onDayClick }) {
  const [monthOffset, setMonthOffset] = useState(0);
  const byDay = useMemo(() => {
    const m = {};
    trades.forEach(t => {
      if (!t.entryTime) return;
      const d = t.entryTime.slice(0, 10);
      if (!m[d]) m[d] = { pnl: 0, count: 0, wins: 0, losses: 0 };
      const p = computePnl(t);
      m[d].pnl += p; m[d].count++;
      if (p > 0) m[d].wins++; else if (p < 0) m[d].losses++;
    });
    return m;
  }, [trades]);

  const base = new Date();
  base.setUTCMonth(base.getUTCMonth() + monthOffset);
  const year = base.getUTCFullYear(), month = base.getUTCMonth();
  const firstDay = new Date(Date.UTC(year, month, 1));
  const startWeekday = firstDay.getUTCDay();
  const daysInMonth = new Date(Date.UTC(year, month + 1, 0)).getUTCDate();
  const cells = [];
  for (let i = 0; i < startWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  let monthPnl = 0;
  Object.entries(byDay).forEach(([k, v]) => { if (k.startsWith(`${year}-${String(month + 1).padStart(2, "0")}`)) monthPnl += v.pnl; });

  const maxAbs = Math.max(1, ...Object.values(byDay).map(v => Math.abs(v.pnl)));

  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div>
          <h3 className="tj-display" style={{ fontSize: 17, fontWeight: 700 }}>
            {base.toLocaleString("en-US", { month: "long", timeZone: "UTC" })} {year}
          </h3>
          <div className="tj-mono" style={{ fontSize: 13, color: monthPnl >= 0 ? "#34D399" : "#F45B69", fontWeight: 700 }}>{fmt$(monthPnl)}</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="tj-btn tj-btn-ghost" onClick={() => setMonthOffset(m => m - 1)}><ChevronDown size={14} style={{ transform: "rotate(90deg)" }} /></button>
          <button className="tj-btn tj-btn-ghost" onClick={() => setMonthOffset(m => m + 1)}><ChevronDown size={14} style={{ transform: "rotate(-90deg)" }} /></button>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 6, marginBottom: 6 }}>
        {["S", "M", "T", "W", "T", "F", "S"].map((d, i) => <div key={i} className="tj-dim" style={{ textAlign: "center", fontSize: 11, fontWeight: 700 }}>{d}</div>)}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 6 }}>
        {cells.map((d, i) => {
          if (!d) return <div key={i} />;
          const key = `${year}-${String(month + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
          const info = byDay[key];
          const intensity = info ? clamp(Math.abs(info.pnl) / maxAbs, 0.15, 1) : 0;
          const bg = !info ? "rgba(255,255,255,0.02)" : info.pnl >= 0
            ? `rgba(52,211,153,${0.12 + intensity * 0.45})`
            : `rgba(244,91,105,${0.12 + intensity * 0.45})`;
          return (
            <div key={i} className="tj-heat-cell" onClick={() => info && onDayClick && onDayClick(key)}
              style={{ background: bg, border: "1px solid rgba(255,255,255,0.06)", padding: "8px 6px", minHeight: 58, cursor: info ? "pointer" : "default" }}>
              <div style={{ fontSize: 10.5, color: "#8A8D97", fontWeight: 600 }}>{d}</div>
              {info && (
                <>
                  <div className="tj-mono" style={{ fontSize: 11, fontWeight: 700, color: info.pnl >= 0 ? "#34D399" : "#F45B69", marginTop: 4 }}>
                    {info.pnl >= 0 ? "+" : ""}{info.pnl.toFixed(0)}
                  </div>
                  <div style={{ fontSize: 9.5, color: "#6B6E78" }}>{info.count} trades</div>
                </>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ============================== DAY DETAIL MODAL ============================== */
function DayDetailModal({ date, trades, onClose, onEditTrade, onDeleteTrade }) {
  const dayTrades = useMemo(() =>
    trades.filter(t => t.entryTime && t.entryTime.slice(0, 10) === date)
      .sort((a, b) => new Date(a.entryTime) - new Date(b.entryTime)),
    [trades, date]
  );
  const pnl = dayTrades.reduce((a, t) => a + computePnl(t), 0);
  const wins = dayTrades.filter(t => computePnl(t) > 0).length;
  const losses = dayTrades.filter(t => computePnl(t) < 0).length;
  const winRate = dayTrades.length ? wins / dayTrades.length : 0;
  const bestTrade = dayTrades.length ? dayTrades.reduce((a, b) => computePnl(a) > computePnl(b) ? a : b) : null;
  const worstTrade = dayTrades.length ? dayTrades.reduce((a, b) => computePnl(a) < computePnl(b) ? a : b) : null;
  const sessions = {};
  dayTrades.forEach(t => { const s = getSession(new Date(t.entryTime).getUTCHours()); sessions[s] = (sessions[s] || 0) + 1; });
  const emotions = {};
  dayTrades.forEach(t => { if (t.emotion) emotions[t.emotion] = (emotions[t.emotion] || 0) + 1; });

  let running = 0;
  const dayCurve = dayTrades.map((t, i) => { running += computePnl(t); return { x: i + 1, y: +running.toFixed(2) }; });

  const dateLabel = new Date(date + "T00:00:00Z").toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric", timeZone: "UTC" });

  return (
    <div className="tj-modal-backdrop" style={{ position: "fixed", inset: 0, zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onClose}>
      <div className="tj-panel tj-fade-in tj-scroll" style={{ width: 780, maxHeight: "88vh", overflowY: "auto", padding: 24 }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18 }}>
          <div>
            <div className="tj-dim" style={{ fontSize: 11.5, fontWeight: 600, textTransform: "uppercase" }}>Day Breakdown</div>
            <h3 className="tj-display" style={{ fontSize: 19, fontWeight: 700 }}>{dateLabel}</h3>
          </div>
          <button className="tj-btn tj-btn-ghost" style={{ padding: 6 }} onClick={onClose}><X size={16} /></button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 10, marginBottom: 18 }}>
          <MiniStat label="Net P&L" value={fmt$(pnl)} tone={pnl >= 0 ? "good" : "bad"} />
          <MiniStat label="Trades" value={dayTrades.length} />
          <MiniStat label="Win Rate" value={fmtPct(winRate)} tone={winRate >= 0.5 ? "good" : "bad"} />
          <MiniStat label="Wins / Losses" value={`${wins}W / ${losses}L`} />
          <MiniStat label="Best Trade" value={bestTrade ? fmt$(computePnl(bestTrade)) : "—"} tone="good" />
        </div>

        {dayCurve.length > 1 && (
          <div style={{ height: 100, marginBottom: 18 }}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={dayCurve}>
                <defs>
                  <linearGradient id="dayFill" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={pnl >= 0 ? "#34D399" : "#F45B69"} stopOpacity={0.35} />
                    <stop offset="100%" stopColor={pnl >= 0 ? "#34D399" : "#F45B69"} stopOpacity={0} />
                  </linearGradient>
                </defs>
                <XAxis dataKey="x" hide />
                <YAxis hide />
                <Tooltip content={<CustomTooltip />} />
                <ReferenceLine y={0} stroke="rgba(255,255,255,0.15)" />
                <Area type="monotone" dataKey="y" stroke={pnl >= 0 ? "#34D399" : "#F45B69"} strokeWidth={2} fill="url(#dayFill)" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        )}

        <div style={{ display: "flex", gap: 20, marginBottom: 18, flexWrap: "wrap" }}>
          <div>
            <div className="tj-dim" style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", marginBottom: 6 }}>Sessions Traded</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {Object.entries(sessions).map(([s, c], i) => <span key={i} className="tj-tag">{s} × {c}</span>)}
            </div>
          </div>
          {Object.keys(emotions).length > 0 && (
            <div>
              <div className="tj-dim" style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", marginBottom: 6 }}>Emotional State</div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {Object.entries(emotions).map(([e, c], i) => <span key={i} className="tj-tag">{e} × {c}</span>)}
              </div>
            </div>
          )}
        </div>

        <SectionTitle icon={ListChecks}>All Trades ({dayTrades.length})</SectionTitle>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {dayTrades.map(t => {
            const p = computePnl(t);
            const r = rMultiple(t);
            return (
              <div key={t.id} className="tj-row-hover" style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 12px", borderRadius: 10, border: "1px solid rgba(255,255,255,0.06)", background: "rgba(255,255,255,0.02)" }}>
                <Pill tone={t.side === "buy" ? "good" : "bad"}>{t.side === "buy" ? "BUY" : "SELL"}</Pill>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 12.5, fontWeight: 700 }}>{t.symbol} <span className="tj-dim" style={{ fontWeight: 500 }}>· {t.lots} lots</span></div>
                  <div className="tj-mono tj-dim" style={{ fontSize: 11 }}>{t.entryPrice} → {t.exitPrice} · {new Date(t.entryTime).toISOString().slice(11, 16)} UTC · {t.setup || "No setup tagged"}</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div className="tj-mono" style={{ fontSize: 13, fontWeight: 700, color: p >= 0 ? "#34D399" : "#F45B69" }}>{fmt$(p)}</div>
                  {r != null && <div className="tj-mono tj-dim" style={{ fontSize: 10.5 }}>{r.toFixed(2)}R</div>}
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  <button className="tj-btn tj-btn-ghost" style={{ padding: 6 }} onClick={() => onEditTrade(t)}><Edit3 size={13} /></button>
                  <button className="tj-btn tj-btn-danger" style={{ padding: 6 }} onClick={() => onDeleteTrade(t.id)}><Trash2 size={13} /></button>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ============================== DAILY PATTERN ANALYSIS ============================== */
function DailyPatterns({ trades }) {
  const patterns = useMemo(() => {
    const byDay = {};
    trades.forEach(t => {
      if (!t.entryTime) return;
      const d = t.entryTime.slice(0, 10);
      (byDay[d] = byDay[d] || []).push(t);
    });
    const counts = {};
    Object.values(byDay).forEach(list => {
      const w = list.filter(t => computePnl(t) > 0).length;
      const l = list.filter(t => computePnl(t) < 0).length;
      const label = `${w}W / ${l}L${list.length !== w + l ? " / BE" : ""}`;
      counts[label] = (counts[label] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 8);
  }, [trades]);
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: 18 }}>
      <SectionTitle icon={Flame}>Daily Trade Patterns</SectionTitle>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {patterns.length === 0 && <div className="tj-dim" style={{ fontSize: 12.5 }}>Not enough data yet.</div>}
        {patterns.map(([label, count], i) => (
          <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12.5 }}>
            <span className="tj-tag">{label}</span>
            <span className="tj-mono tj-dim">{count} day{count !== 1 ? "s" : ""}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================== PSYCHOLOGY PANEL ============================== */
function weekKey(dateStr) {
  const d = new Date(dateStr + "T00:00:00Z");
  const day = d.getUTCDay();
  const diffToMonday = (day === 0 ? -6 : 1) - day;
  const monday = new Date(d);
  monday.setUTCDate(d.getUTCDate() + diffToMonday);
  return monday.toISOString().slice(0, 10);
}
function monthKey(dateStr) { return dateStr.slice(0, 7); }

function RuleSettingsPanel({ config, setConfig, customRules, setCustomRules }) {
  const [newRule, setNewRule] = useState({ label: "", field: "pnl", operator: "gt", value: "", severity: "medium" });
  const setCfg = (k, v) => setConfig(c => ({ ...c, [k]: v }));
  const addRule = () => {
    if (!newRule.label.trim() || newRule.value === "") return;
    setCustomRules(rs => [...rs, { ...newRule, id: Date.now().toString(36) }]);
    setNewRule({ label: "", field: "pnl", operator: "gt", value: "", severity: "medium" });
  };
  const removeRule = (id) => setCustomRules(rs => rs.filter(r => r.id !== id));

  const NumField = ({ label, k, step = 1 }) => (
    <div>
      <label style={{ fontSize: 11, color: "#8A8D97", fontWeight: 600 }}>{label}</label>
      <input type="number" step={step} className="tj-input" style={{ marginTop: 4 }} value={config[k]} onChange={e => setCfg(k, parseFloat(e.target.value) || 0)} />
    </div>
  );

  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={Settings}>Detection Rules — Tune to Your Strategy</SectionTitle>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 12, marginBottom: 20 }}>
        <NumField label="Revenge gap (min)" k="revengeGapMin" />
        <NumField label="Revenge lot multiplier" k="revengeLotMultiplier" step={0.05} />
        <NumField label="Overtrading trades/day" k="overtradingCount" />
        <NumField label="Cut-winner fraction (0-1)" k="cuttingWinnersFraction" step={0.05} />
        <NumField label="Hold-loser multiplier" k="holdingLosersMultiplier" step={0.1} />
        <NumField label="Session start (UTC hr)" k="sessionStartHour" />
        <NumField label="Session end (UTC hr)" k="sessionEndHour" />
      </div>

      <div style={{ borderTop: "1px solid rgba(255,255,255,0.08)", paddingTop: 16 }}>
        <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 10 }}>Your Custom Strategy Rules</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 14 }}>
          {customRules.length === 0 && <div className="tj-dim" style={{ fontSize: 12.5 }}>No custom rules yet — add one below (e.g. "Lot Size &gt; 0.5 on London session").</div>}
          {customRules.map(r => (
            <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12.5, background: "rgba(255,255,255,0.03)", borderRadius: 8, padding: "8px 12px" }}>
              <span><strong>{r.label}</strong> — {RULE_FIELDS.find(f => f.key === r.field)?.label} {RULE_OPERATORS.find(o => o.key === r.operator)?.label} {r.value}</span>
              <button onClick={() => removeRule(r.id)} style={{ background: "none", border: "none", color: "#F45B69", cursor: "pointer" }}><X size={14} /></button>
            </div>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1.2fr 1fr 0.7fr 0.8fr 0.9fr auto", gap: 8, alignItems: "end" }}>
          <div>
            <label style={{ fontSize: 10.5, color: "#8A8D97" }}>Rule Name</label>
            <input className="tj-input" placeholder="e.g. Oversized London Entry" value={newRule.label} onChange={e => setNewRule(r => ({ ...r, label: e.target.value }))} />
          </div>
          <div>
            <label style={{ fontSize: 10.5, color: "#8A8D97" }}>Field</label>
            <select className="tj-input" value={newRule.field} onChange={e => setNewRule(r => ({ ...r, field: e.target.value }))}>
              {RULE_FIELDS.map(f => <option key={f.key} value={f.key}>{f.label}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize: 10.5, color: "#8A8D97" }}>Operator</label>
            <select className="tj-input" value={newRule.operator} onChange={e => setNewRule(r => ({ ...r, operator: e.target.value }))}>
              {RULE_OPERATORS.map(o => <option key={o.key} value={o.key}>{o.label}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize: 10.5, color: "#8A8D97" }}>Value</label>
            <input className="tj-input" value={newRule.value} onChange={e => setNewRule(r => ({ ...r, value: e.target.value }))} />
          </div>
          <div>
            <label style={{ fontSize: 10.5, color: "#8A8D97" }}>Severity</label>
            <select className="tj-input" value={newRule.severity} onChange={e => setNewRule(r => ({ ...r, severity: e.target.value }))}>
              <option value="low">Low</option><option value="medium">Medium</option><option value="high">High</option>
            </select>
          </div>
          <button className="tj-btn tj-btn-primary" onClick={addRule}><Plus size={14} /></button>
        </div>
      </div>
    </div>
  );
}

function PsychologyPeriodBreakdown({ flags }) {
  const [period, setPeriod] = useState("daily");
  const grouped = useMemo(() => {
    const m = {};
    flags.forEach(f => {
      if (!f.date) return;
      const key = period === "daily" ? f.date : period === "weekly" ? weekKey(f.date) : monthKey(f.date);
      m[key] = (m[key] || 0) + 1;
    });
    return Object.entries(m).sort((a, b) => a[0].localeCompare(b[0])).map(([k, v]) => ({ label: k, count: v }));
  }, [flags, period]);

  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <SectionTitle icon={CalendarDays}>Psychology Over Time</SectionTitle>
        <div style={{ display: "flex", gap: 6 }}>
          {["daily", "weekly", "monthly"].map(p => (
            <button key={p} className="tj-btn" style={{
              padding: "7px 12px",
              background: period === p ? "rgba(227,178,77,0.14)" : "rgba(255,255,255,0.04)",
              color: period === p ? "#E3B24D" : "#F2F1ED",
              borderColor: period === p ? "rgba(227,178,77,0.3)" : "rgba(255,255,255,0.1)",
            }} onClick={() => setPeriod(p)}>{p[0].toUpperCase() + p.slice(1)}</button>
          ))}
        </div>
      </div>
      {grouped.length === 0 ? (
        <div className="tj-dim" style={{ fontSize: 13 }}>No flagged events yet for this view.</div>
      ) : (
        <div style={{ height: 220 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={grouped}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" vertical={false} />
              <XAxis dataKey="label" tick={{ fill: "#8A8D97", fontSize: 10 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} width={30} allowDecimals={false} />
              <Tooltip content={<CustomTooltip prefix="" />} />
              <Bar dataKey="count" radius={[5, 5, 0, 0]} fill="#E3B24D" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

function PsychologyPanel({ trades, journalId }) {
  const [config, setConfig] = useState(DEFAULT_PSYCH_CONFIG);
  const [customRules, setCustomRules] = useState([]);
  const [showSettings, setShowSettings] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const jid = journalId || "default";

  useEffect(() => {
    (async () => {
      try {
        const cfgRes = await window.storage?.get(`psychConfig:${jid}`, false);
        if (cfgRes?.value) setConfig(JSON.parse(cfgRes.value));
      } catch (e) {}
      try {
        const rulesRes = await window.storage?.get(`psychRules:${jid}`, false);
        if (rulesRes?.value) setCustomRules(JSON.parse(rulesRes.value));
      } catch (e) {}
      setLoaded(true);
    })();
  }, [jid]);
  useEffect(() => { if (loaded) { try { window.storage?.set(`psychConfig:${jid}`, JSON.stringify(config), false); } catch (e) {} } }, [config, loaded, jid]);
  useEffect(() => { if (loaded) { try { window.storage?.set(`psychRules:${jid}`, JSON.stringify(customRules), false); } catch (e) {} } }, [customRules, loaded, jid]);

  const { flags, counts } = useMemo(() => detectPsychology(trades, config, customRules), [trades, config, customRules]);
  const severityColor = { high: "bad", medium: "warn", low: "neutral" };

  return (
    <div className="tj-fade-in" style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div className="tj-panel" style={{ padding: 20 }}>
        <SectionTitle icon={Brain} right={
          <button className="tj-btn tj-btn-ghost" onClick={() => setShowSettings(s => !s)}>
            <Settings size={14} /> {showSettings ? "Hide Rules" : "Set My Rules"}
          </button>
        }>Trading Psychology Radar</SectionTitle>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
          {Object.entries(counts).length === 0 && (
            <div className="tj-dim" style={{ fontSize: 13 }}>No behavioral warning patterns detected — clean execution across your logged trades.</div>
          )}
          {Object.entries(counts).map(([type, count], i) => (
            <div key={i} className="tj-panel-tight" style={{ padding: 14, background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <AlertTriangle size={14} className="tj-gold" />
                <span className="tj-mono" style={{ fontSize: 20, fontWeight: 700 }}>{count}</span>
              </div>
              <div style={{ fontSize: 12, fontWeight: 600, marginTop: 6 }}>{type}</div>
            </div>
          ))}
        </div>
      </div>

      {showSettings && <RuleSettingsPanel config={config} setConfig={setConfig} customRules={customRules} setCustomRules={setCustomRules} />}

      <PsychologyPeriodBreakdown flags={flags} />

      <div className="tj-panel" style={{ padding: 20 }}>
        <SectionTitle icon={Eye}>Flagged Trade Events</SectionTitle>
        <div className="tj-scroll" style={{ maxHeight: 420, overflowY: "auto", display: "flex", flexDirection: "column", gap: 10 }}>
          {flags.length === 0 && <div className="tj-dim" style={{ fontSize: 13 }}>Nothing to flag. Keep it up.</div>}
          {flags.map((f, i) => (
            <div key={i} style={{ display: "flex", gap: 12, padding: "10px 12px", borderRadius: 10, background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)" }}>
              <Pill tone={severityColor[f.severity]}>{f.severity.toUpperCase()}</Pill>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 12.5, fontWeight: 700 }}>{f.type} {f.tradeId ? <span className="tj-dim" style={{ fontWeight: 500 }}>· Trade #{f.tradeId}</span> : f.day ? <span className="tj-dim" style={{ fontWeight: 500 }}>· {f.day}</span> : null}</div>
                <div style={{ fontSize: 12, color: "#8A8D97", marginTop: 2 }}>{f.detail}</div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ============================== TRADE MODAL ============================== */
function emptyTrade() {
  return {
    id: null, symbol: "XAUUSD", side: "buy", lots: 0.1,
    entryPrice: "", exitPrice: "", sl: "", tp: "",
    entryTime: new Date().toISOString().slice(0, 16),
    exitTime: new Date().toISOString().slice(0, 16),
    commission: 0, swap: 0, setup: SETUPS[0], emotion: EMOTIONS[0],
    tags: [], notes: "", screenshots: [], checklist: [], customFields: {},
    mfe: "", mae: "",
  };
}

const DEFAULT_CUSTOM_COLUMNS = [
  { id: "cc_4h", label: "4H Candle" },
  { id: "cc_1h", label: "1H Candle" },
  { id: "cc_daily", label: "Daily Candle" },
  { id: "cc_minrrr", label: "Min RRR" },
  { id: "cc_holdtiming", label: "Trade Hold Timing" },
  { id: "cc_breakout1", label: "1st Breakout" },
  { id: "cc_breakout2", label: "2nd Breakout" },
];

function TradeModal({ trade, onSave, onClose, customColumns }) {
  const [form, setForm] = useState(() => {
    const base = trade || emptyTrade();
    return { ...base, customFields: base.customFields || {} };
  });
  const fileRef = useRef();
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const setCustom = (colId, v) => setForm(f => ({ ...f, customFields: { ...f.customFields, [colId]: v } }));

  const toggleTag = (tag) => {
    setForm(f => ({ ...f, tags: f.tags.includes(tag) ? f.tags.filter(t => t !== tag) : [...f.tags, tag] }));
  };

  const handleFile = (e) => {
    const files = Array.from(e.target.files || []);
    files.forEach(file => {
      const reader = new FileReader();
      reader.onload = () => setForm(f => ({ ...f, screenshots: [...f.screenshots, reader.result] }));
      reader.readAsDataURL(file);
    });
  };

  const submit = () => {
    const t = {
      ...form,
      id: form.id || Date.now(),
      lots: parseFloat(form.lots) || 0,
      entryPrice: parseFloat(form.entryPrice) || 0,
      exitPrice: parseFloat(form.exitPrice) || 0,
      sl: form.sl === "" ? null : parseFloat(form.sl),
      tp: form.tp === "" ? null : parseFloat(form.tp),
      commission: parseFloat(form.commission) || 0,
      swap: parseFloat(form.swap) || 0,
      mfe: form.mfe === "" ? undefined : parseFloat(form.mfe),
      mae: form.mae === "" ? undefined : parseFloat(form.mae),
      entryTime: new Date(form.entryTime).toISOString(),
      exitTime: new Date(form.exitTime).toISOString(),
      customFields: form.customFields || {},
    };
    onSave(t);
  };

  const Field = ({ label, children }) => (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      <label style={{ fontSize: 11, fontWeight: 600, color: "#8A8D97", textTransform: "uppercase", letterSpacing: "0.03em" }}>{label}</label>
      {children}
    </div>
  );

  return (
    <div className="tj-modal-backdrop" style={{ position: "fixed", inset: 0, zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onClose}>
      <div className="tj-panel tj-fade-in tj-scroll" style={{ width: 640, maxHeight: "90vh", overflowY: "auto", padding: 24 }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h3 className="tj-display" style={{ fontSize: 18, fontWeight: 700 }}>{form.id ? "Edit Trade" : "Log New Trade"}</h3>
          <button className="tj-btn tj-btn-ghost" style={{ padding: 6 }} onClick={onClose}><X size={16} /></button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
          <Field label="Symbol"><input className="tj-input" value={form.symbol} onChange={e => set("symbol", e.target.value)} /></Field>
          <Field label="Side">
            <select className="tj-input" value={form.side} onChange={e => set("side", e.target.value)}>
              <option value="buy">Buy</option><option value="sell">Sell</option>
            </select>
          </Field>
          <Field label="Lot Size"><input type="number" step="0.01" className="tj-input" value={form.lots} onChange={e => set("lots", e.target.value)} /></Field>

          <Field label="Entry Price"><input type="number" step="0.01" className="tj-input" value={form.entryPrice} onChange={e => set("entryPrice", e.target.value)} /></Field>
          <Field label="Exit Price"><input type="number" step="0.01" className="tj-input" value={form.exitPrice} onChange={e => set("exitPrice", e.target.value)} /></Field>
          <Field label="Stop Loss"><input type="number" step="0.01" className="tj-input" value={form.sl} onChange={e => set("sl", e.target.value)} /></Field>

          <Field label="Take Profit"><input type="number" step="0.01" className="tj-input" value={form.tp} onChange={e => set("tp", e.target.value)} /></Field>
          <Field label="Commission"><input type="number" step="0.01" className="tj-input" value={form.commission} onChange={e => set("commission", e.target.value)} /></Field>
          <Field label="Swap"><input type="number" step="0.01" className="tj-input" value={form.swap} onChange={e => set("swap", e.target.value)} /></Field>

          <Field label="Entry Time"><input type="datetime-local" className="tj-input" value={form.entryTime?.slice(0, 16)} onChange={e => set("entryTime", e.target.value)} /></Field>
          <Field label="Exit Time"><input type="datetime-local" className="tj-input" value={form.exitTime?.slice(0, 16)} onChange={e => set("exitTime", e.target.value)} /></Field>
          <Field label="Setup">
            <select className="tj-input" value={form.setup} onChange={e => set("setup", e.target.value)}>
              {SETUPS.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </Field>

          <Field label="Emotion">
            <select className="tj-input" value={form.emotion} onChange={e => set("emotion", e.target.value)}>
              {EMOTIONS.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </Field>
        </div>

        <div style={{ marginTop: 14 }}>
          <Field label="Tags">
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 2 }}>
              {TAG_OPTIONS.map(tag => (
                <button key={tag} onClick={() => toggleTag(tag)} className="tj-tag" style={{
                  cursor: "pointer", background: form.tags.includes(tag) ? "rgba(227,178,77,0.18)" : undefined,
                  borderColor: form.tags.includes(tag) ? "rgba(227,178,77,0.4)" : undefined,
                  color: form.tags.includes(tag) ? "#E3B24D" : undefined,
                }}>{tag}</button>
              ))}
            </div>
          </Field>
        </div>

        <div style={{ marginTop: 14, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Max Favorable Excursion ($)">
            <input type="number" step="0.01" className="tj-input" placeholder="Best unrealized profit reached" value={form.mfe} onChange={e => set("mfe", e.target.value)} />
          </Field>
          <Field label="Max Adverse Excursion ($)">
            <input type="number" step="0.01" className="tj-input" placeholder="Worst unrealized drawdown reached" value={form.mae} onChange={e => set("mae", e.target.value)} />
          </Field>
        </div>

        <div style={{ marginTop: 14 }}>
          <Field label="Notes">
            <textarea className="tj-input" rows={3} value={form.notes} onChange={e => set("notes", e.target.value)} placeholder="What was your read? What did you see at entry/exit?" />
          </Field>
        </div>

        <div style={{ marginTop: 14 }}>
          <Field label="Screenshots (before / after)">
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 4 }}>
              {form.screenshots.map((s, i) => (
                <div key={i} style={{ position: "relative" }}>
                  <img src={s} alt="" style={{ width: 72, height: 72, objectFit: "cover", borderRadius: 8, border: "1px solid rgba(255,255,255,0.1)" }} />
                  <button onClick={() => setForm(f => ({ ...f, screenshots: f.screenshots.filter((_, j) => j !== i) }))}
                    style={{ position: "absolute", top: -6, right: -6, background: "#F45B69", borderRadius: "50%", width: 18, height: 18, border: "none", color: "#fff", cursor: "pointer", fontSize: 11 }}>×</button>
                </div>
              ))}
              <button onClick={() => fileRef.current?.click()} className="tj-btn tj-btn-ghost" style={{ width: 72, height: 72, justifyContent: "center" }}>
                <ImageIcon size={18} />
              </button>
              <input ref={fileRef} type="file" accept="image/*" multiple hidden onChange={handleFile} />
            </div>
          </Field>
        </div>

        {customColumns && customColumns.length > 0 && (
          <div style={{ marginTop: 14 }}>
            <label style={{ fontSize: 11, fontWeight: 600, color: "#8A8D97", textTransform: "uppercase", letterSpacing: "0.03em" }}>Custom Details</label>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginTop: 6 }}>
              {customColumns.map(col => (
                <Field key={col.id} label={col.label}>
                  <input className="tj-input" value={form.customFields?.[col.id] || ""} onChange={e => setCustom(col.id, e.target.value)} placeholder={col.label} />
                </Field>
              ))}
            </div>
          </div>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 22 }}>
          <button className="tj-btn tj-btn-ghost" onClick={onClose}>Cancel</button>
          <button className="tj-btn tj-btn-primary" onClick={submit}>Save Trade</button>
        </div>
      </div>
    </div>
  );
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function extractTradesFromImage(base64Data, mediaType) {
  const system = `You are a precise data-extraction engine for trading platform screenshots (MT4, MT5, broker apps, trade history tables). Extract every trade row visible in the image into a strict JSON array and output ONLY the JSON array — no markdown fences, no commentary, no explanation.

Each object must use exactly these fields:
- symbol (string)
- side ("buy" or "sell")
- lots (number)
- entryPrice (number)
- exitPrice (number, null if not visible/still open)
- sl (number or null)
- tp (number or null)
- openTime (ISO 8601 string — infer the year as 2026 if not shown, keep month/day/time as shown)
- closeTime (ISO 8601 string, same rule, null if not visible)
- swap (number, 0 if not shown)
- pnl (number — the P/L or profit column, null if not visible)
- reason (string or null — e.g. "Stop Loss", "Take Profit", "Manual")

If a value is not visible in the image, use null. Never invent numbers. Output valid JSON only.`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 2000,
      system,
      messages: [{
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } },
          { type: "text", text: "Extract every trade row from this screenshot as a JSON array, following the schema exactly." },
        ],
      }],
    }),
  });
  const data = await response.json();
  const text = (data.content || []).map(c => c.text || "").join("\n");
  const clean = text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(clean);
  return Array.isArray(parsed) ? parsed : [];
}

function makeEditableRow(t = {}, screenshotDataUrl) {
  return {
    _rowId: Math.random().toString(36).slice(2),
    symbol: t.symbol || "XAUUSD",
    side: (t.side || "buy").toLowerCase().includes("sell") ? "sell" : "buy",
    lots: t.lots ?? t.volume ?? 0.01,
    entryPrice: t.entryPrice ?? "",
    exitPrice: t.exitPrice ?? "",
    sl: t.sl ?? "",
    tp: t.tp ?? "",
    openTime: t.openTime ? new Date(t.openTime).toISOString().slice(0, 16) : new Date().toISOString().slice(0, 16),
    closeTime: t.closeTime ? new Date(t.closeTime).toISOString().slice(0, 16) : new Date().toISOString().slice(0, 16),
    swap: t.swap ?? 0,
    pnl: t.pnl ?? "",
    reason: t.reason || "",
    screenshot: screenshotDataUrl || null,
  };
}

/* ============================== CSV / MT4 IMPORT ============================== */
function detectField(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const found = keys.find(k => k.toLowerCase().replace(/[\s_]/g, "").includes(c));
    if (found) return found;
  }
  return null;
}

function parseImportedRows(rows) {
  if (!rows.length) return [];
  const sample = rows[0];
  const symbolKey = detectField(sample, ["symbol", "item", "instrument"]);
  const sideKey = detectField(sample, ["type", "side", "direction"]);
  const lotsKey = detectField(sample, ["volume", "lots", "size", "qty"]);
  const openPriceKey = detectField(sample, ["openprice", "entryprice", "priceopen"]);
  const closePriceKey = detectField(sample, ["closeprice", "exitprice", "priceclose"]);
  const slKey = detectField(sample, ["s/l", "sl", "stoploss"]);
  const tpKey = detectField(sample, ["t/p", "tp", "takeprofit"]);
  const openTimeKey = detectField(sample, ["opentime", "entrytime", "timeopen", "date/time"]);
  const closeTimeKey = detectField(sample, ["closetime", "exittime", "timeclose"]);
  const swapKey = detectField(sample, ["swap"]);
  const profitKey = detectField(sample, ["profit", "p/l", "pnl", "netprofit"]);
  const reasonKey = detectField(sample, ["reason", "comment"]);

  return rows.filter(r => r[symbolKey || ""] || r[openPriceKey || ""]).map(r => ({
    symbol: (r[symbolKey] || "XAUUSD").toString().toUpperCase(),
    side: r[sideKey], lots: parseFloat(r[lotsKey]) || 0.01,
    entryPrice: parseFloat(r[openPriceKey]) || "",
    exitPrice: parseFloat(r[closePriceKey]) || "",
    sl: r[slKey] ? parseFloat(r[slKey]) : "",
    tp: r[tpKey] ? parseFloat(r[tpKey]) : "",
    openTime: r[openTimeKey] || new Date().toISOString(),
    closeTime: r[closeTimeKey] || new Date().toISOString(),
    swap: parseFloat(r[swapKey]) || 0,
    pnl: r[profitKey] !== undefined ? parseFloat(r[profitKey]) : "",
    reason: r[reasonKey] || "",
  }));
}

/* -- Editable preview grid: used for CSV/HTML rows AND AI-extracted screenshot rows -- */
function ImportPreviewTable({ rows, setRows }) {
  const update = (rowId, field, value) => setRows(rs => rs.map(r => (r._rowId === rowId ? { ...r, [field]: value } : r)));
  const removeRow = (rowId) => setRows(rs => rs.filter(r => r._rowId !== rowId));
  const addRow = () => setRows(rs => [...rs, makeEditableRow()]);
  const cellStyle = { padding: "6px 6px" };
  const inp = { padding: "6px 8px", fontSize: 11.5, width: "100%" };

  return (
    <div>
      <div className="tj-scroll" style={{ maxHeight: 320, overflow: "auto", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
          <thead style={{ position: "sticky", top: 0, background: "#111214", zIndex: 1 }}>
            <tr>
              {["Symbol", "Side", "Lots", "Entry", "Exit", "SL", "TP", "Open Time", "Close Time", "Swap", "P/L", "Reason", ""].map((h, i) => (
                <th key={i} style={{ padding: "8px 6px", fontSize: 10, color: "#8A8D97", fontWeight: 700, textTransform: "uppercase", textAlign: "left", whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(r => (
              <tr key={r._rowId} style={{ borderTop: "1px solid rgba(255,255,255,0.05)" }}>
                <td style={cellStyle}><input className="tj-input" style={{ ...inp, width: 70 }} value={r.symbol} onChange={e => update(r._rowId, "symbol", e.target.value)} /></td>
                <td style={cellStyle}>
                  <select className="tj-input" style={{ ...inp, width: 62 }} value={r.side} onChange={e => update(r._rowId, "side", e.target.value)}>
                    <option value="buy">Buy</option><option value="sell">Sell</option>
                  </select>
                </td>
                <td style={cellStyle}><input type="number" step="0.01" className="tj-input" style={{ ...inp, width: 55 }} value={r.lots} onChange={e => update(r._rowId, "lots", e.target.value)} /></td>
                <td style={cellStyle}><input type="number" step="0.001" className="tj-input" style={{ ...inp, width: 80 }} value={r.entryPrice} onChange={e => update(r._rowId, "entryPrice", e.target.value)} /></td>
                <td style={cellStyle}><input type="number" step="0.001" className="tj-input" style={{ ...inp, width: 80 }} value={r.exitPrice} onChange={e => update(r._rowId, "exitPrice", e.target.value)} /></td>
                <td style={cellStyle}><input type="number" step="0.001" className="tj-input" style={{ ...inp, width: 75 }} value={r.sl} onChange={e => update(r._rowId, "sl", e.target.value)} /></td>
                <td style={cellStyle}><input type="number" step="0.001" className="tj-input" style={{ ...inp, width: 75 }} value={r.tp} onChange={e => update(r._rowId, "tp", e.target.value)} /></td>
                <td style={cellStyle}><input type="datetime-local" className="tj-input" style={{ ...inp, width: 150 }} value={r.openTime?.slice(0, 16)} onChange={e => update(r._rowId, "openTime", e.target.value)} /></td>
                <td style={cellStyle}><input type="datetime-local" className="tj-input" style={{ ...inp, width: 150 }} value={r.closeTime?.slice(0, 16)} onChange={e => update(r._rowId, "closeTime", e.target.value)} /></td>
                <td style={cellStyle}><input type="number" step="0.01" className="tj-input" style={{ ...inp, width: 60 }} value={r.swap} onChange={e => update(r._rowId, "swap", e.target.value)} /></td>
                <td style={cellStyle}><input type="number" step="0.01" className="tj-input" style={{ ...inp, width: 70 }} value={r.pnl} onChange={e => update(r._rowId, "pnl", e.target.value)} /></td>
                <td style={cellStyle}><input className="tj-input" style={{ ...inp, width: 90 }} value={r.reason} onChange={e => update(r._rowId, "reason", e.target.value)} /></td>
                <td style={cellStyle}><button onClick={() => removeRow(r._rowId)} style={{ background: "none", border: "none", color: "#F45B69", cursor: "pointer" }}><Trash2 size={13} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <button className="tj-btn tj-btn-ghost" style={{ marginTop: 10 }} onClick={addRow}><Plus size={13} /> Add Row Manually</button>
    </div>
  );
}

function ImportModal({ onImport, onClose }) {
  const [status, setStatus] = useState("");
  const [rows, setRows] = useState([]);
  const [loadingAI, setLoadingAI] = useState(false);
  const [mode, setMode] = useState("file"); // 'file' | 'screenshot'
  const fileRef = useRef();
  const imgRef = useRef();

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setStatus("Parsing...");
    const isHtml = file.name.toLowerCase().endsWith(".html") || file.name.toLowerCase().endsWith(".htm");
    if (isHtml) {
      const reader = new FileReader();
      reader.onload = () => {
        const html = reader.result;
        const doc = new DOMParser().parseFromString(html, "text/html");
        const tables = Array.from(doc.querySelectorAll("table"));
        let parsedRows = [];
        tables.forEach(table => {
          const trs = Array.from(table.querySelectorAll("tr"));
          if (trs.length < 2) return;
          const headerCells = Array.from(trs[0].querySelectorAll("th,td")).map(td => td.textContent.trim());
          if (!headerCells.some(h => /open|close|symbol|type|price/i.test(h))) return;
          for (let i = 1; i < trs.length; i++) {
            const cells = Array.from(trs[i].querySelectorAll("td")).map(td => td.textContent.trim());
            if (cells.length < headerCells.length - 2) continue;
            const obj = {};
            headerCells.forEach((h, idx) => obj[h] = cells[idx]);
            parsedRows.push(obj);
          }
        });
        const parsed = parseImportedRows(parsedRows).map(r => makeEditableRow(r));
        setRows(parsed);
        setStatus(parsed.length ? `Found ${parsed.length} trades — review and edit below, then import.` : "No recognizable trade rows found.");
      };
      reader.readAsText(file);
    } else {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: (res) => {
          const parsed = parseImportedRows(res.data).map(r => makeEditableRow(r));
          setRows(parsed);
          setStatus(parsed.length ? `Found ${parsed.length} trades — review and edit below, then import.` : "No recognizable trade rows found.");
        },
        error: () => setStatus("Could not parse this file."),
      });
    }
  };

  const handleScreenshot = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoadingAI(true);
    setStatus("Reading the screenshot and extracting trade data…");
    try {
      const base64 = await fileToBase64(file);
      const dataUrl = `data:${file.type};base64,${base64}`;
      const extracted = await extractTradesFromImage(base64, file.type || "image/png");
      const parsed = extracted.map(r => makeEditableRow(r, dataUrl));
      setRows(parsed);
      setStatus(parsed.length
        ? `Extracted ${parsed.length} trades from the screenshot — check every field below before importing.`
        : "Couldn't confidently extract any rows. Try a clearer screenshot, or add rows manually below.");
    } catch (err) {
      setStatus("Extraction failed — the image may be unclear. You can still add trades manually below.");
    } finally {
      setLoadingAI(false);
    }
  };

  const confirmImport = () => {
    const trades = rows.filter(r => r.entryPrice !== "" && r.entryPrice !== null).map(r => {
      const commission = 0;
      const t = {
        id: Date.now() + Math.random(),
        symbol: r.symbol, side: r.side, lots: parseFloat(r.lots) || 0.01,
        entryPrice: parseFloat(r.entryPrice) || 0,
        exitPrice: r.exitPrice === "" ? null : parseFloat(r.exitPrice),
        sl: r.sl === "" ? null : parseFloat(r.sl),
        tp: r.tp === "" ? null : parseFloat(r.tp),
        entryTime: new Date(r.openTime).toISOString(),
        exitTime: new Date(r.closeTime).toISOString(),
        commission, swap: parseFloat(r.swap) || 0,
        pnl: r.pnl === "" ? undefined : parseFloat(r.pnl),
        setup: "", emotion: "",
        tags: r.reason ? [r.reason] : [],
        notes: r.reason ? `Closed via: ${r.reason}` : "",
        screenshots: r.screenshot ? [r.screenshot] : [],
        customFields: {},
      };
      return t;
    });
    onImport(trades);
    onClose();
  };

  return (
    <div className="tj-modal-backdrop" style={{ position: "fixed", inset: 0, zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onClose}>
      <div className="tj-panel tj-fade-in tj-scroll" style={{ width: 960, maxHeight: "90vh", overflowY: "auto", padding: 24 }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h3 className="tj-display" style={{ fontSize: 18, fontWeight: 700 }}>Import Trades</h3>
          <button className="tj-btn tj-btn-ghost" style={{ padding: 6 }} onClick={onClose}><X size={16} /></button>
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          <button className="tj-btn" style={{ background: mode === "file" ? "rgba(227,178,77,0.14)" : "rgba(255,255,255,0.04)", color: mode === "file" ? "#E3B24D" : "#F2F1ED", borderColor: mode === "file" ? "rgba(227,178,77,0.3)" : "rgba(255,255,255,0.1)" }}
            onClick={() => { setMode("file"); setRows([]); setStatus(""); }}>
            <Upload size={13} /> File (HTML / CSV)
          </button>
          <button className="tj-btn" style={{ background: mode === "screenshot" ? "rgba(227,178,77,0.14)" : "rgba(255,255,255,0.04)", color: mode === "screenshot" ? "#E3B24D" : "#F2F1ED", borderColor: mode === "screenshot" ? "rgba(227,178,77,0.3)" : "rgba(255,255,255,0.1)" }}
            onClick={() => { setMode("screenshot"); setRows([]); setStatus(""); }}>
            <ImageIcon size={13} /> Screenshot (AI extract)
          </button>
        </div>

        {mode === "file" ? (
          <>
            <p className="tj-dim" style={{ fontSize: 12.5, marginBottom: 14 }}>
              MT4/MT5 exported HTML statements, CSV, and Excel-exported CSV. Columns are auto-detected — symbol, side, volume, open/close price, S/L, T/P, timestamps, swap, and profit.
            </p>
            <div onClick={() => fileRef.current?.click()} style={{
              border: "1.5px dashed rgba(227,178,77,0.35)", borderRadius: 14, padding: "26px 20px",
              textAlign: "center", cursor: "pointer", background: "rgba(227,178,77,0.04)",
            }}>
              <Upload size={24} className="tj-gold" style={{ margin: "0 auto 8px" }} />
              <div style={{ fontSize: 13, fontWeight: 600 }}>Click to choose a file</div>
              <div className="tj-dim" style={{ fontSize: 11.5, marginTop: 4 }}>.html · .csv · .xlsx (exported as CSV)</div>
            </div>
            <input ref={fileRef} type="file" accept=".html,.htm,.csv,.txt" hidden onChange={handleFile} />
          </>
        ) : (
          <>
            <p className="tj-dim" style={{ fontSize: 12.5, marginBottom: 14 }}>
              Upload a screenshot of your trade history (MT4/MT5 terminal, mobile app, or broker portal). The AI reads every visible row and drops it into an editable table below — nothing gets imported until you confirm it.
            </p>
            <div onClick={() => !loadingAI && imgRef.current?.click()} style={{
              border: "1.5px dashed rgba(227,178,77,0.35)", borderRadius: 14, padding: "26px 20px",
              textAlign: "center", cursor: loadingAI ? "wait" : "pointer", background: "rgba(227,178,77,0.04)",
            }}>
              {loadingAI ? <RefreshCw size={24} className="tj-gold tj-pulse" style={{ margin: "0 auto 8px" }} /> : <ImageIcon size={24} className="tj-gold" style={{ margin: "0 auto 8px" }} />}
              <div style={{ fontSize: 13, fontWeight: 600 }}>{loadingAI ? "Extracting…" : "Click to upload a screenshot"}</div>
              <div className="tj-dim" style={{ fontSize: 11.5, marginTop: 4 }}>.png · .jpg · .webp</div>
            </div>
            <input ref={imgRef} type="file" accept="image/*" hidden onChange={handleScreenshot} />
          </>
        )}

        {status && <div className="tj-mono" style={{ fontSize: 12, marginTop: 14, color: "#E3B24D" }}>{status}</div>}

        {rows.length > 0 && (
          <div style={{ marginTop: 18 }}>
            <SectionTitle icon={Edit3}>Review &amp; Edit Before Import</SectionTitle>
            <ImportPreviewTable rows={rows} setRows={setRows} />
          </div>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
          <button className="tj-btn tj-btn-ghost" onClick={onClose}>Cancel</button>
          <button className="tj-btn tj-btn-primary" disabled={!rows.length} style={{ opacity: rows.length ? 1 : 0.5 }} onClick={confirmImport}>
            Import {rows.length || ""} Trades
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================== TRADE TABLE ============================== */
function ColumnManager({ customColumns, onAddColumn, onRemoveColumn }) {
  const [open, setOpen] = useState(false);
  const [newLabel, setNewLabel] = useState("");
  const add = () => {
    if (!newLabel.trim()) return;
    onAddColumn(newLabel.trim());
    setNewLabel("");
  };
  return (
    <div style={{ position: "relative" }}>
      <button className="tj-btn tj-btn-ghost" onClick={() => setOpen(o => !o)}>
        <ListChecks size={14} /> Columns {open ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
      </button>
      {open && (
        <div className="tj-panel tj-fade-in" style={{
          position: "absolute", right: 0, top: "calc(100% + 8px)", width: 300, padding: 16, zIndex: 20,
        }}>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: "#8A8D97", textTransform: "uppercase", marginBottom: 10 }}>Custom Columns</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 220, overflowY: "auto" }} className="tj-scroll">
            {customColumns.length === 0 && <div className="tj-dim" style={{ fontSize: 12.5 }}>No custom columns yet.</div>}
            {customColumns.map(col => (
              <div key={col.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12.5, background: "rgba(255,255,255,0.03)", borderRadius: 8, padding: "6px 10px" }}>
                <span>{col.label}</span>
                <button onClick={() => onRemoveColumn(col.id)} style={{ background: "none", border: "none", color: "#F45B69", cursor: "pointer" }}><X size={14} /></button>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 6, marginTop: 12 }}>
            <input className="tj-input" placeholder="New column name..." value={newLabel}
              onChange={e => setNewLabel(e.target.value)} onKeyDown={e => e.key === "Enter" && add()} />
            <button className="tj-btn tj-btn-primary" style={{ padding: "9px 12px" }} onClick={add}><Plus size={14} /></button>
          </div>
          <div className="tj-dim" style={{ fontSize: 11, marginTop: 8 }}>Add or remove as many columns as you need — they show up here and in the trade form instantly.</div>
        </div>
      )}
    </div>
  );
}

function TradeTable({ trades, onEdit, onDelete, customColumns, onAddColumn, onRemoveColumn }) {
  const [search, setSearch] = useState("");
  const [sortKey, setSortKey] = useState("entryTime");
  const [sortDir, setSortDir] = useState(-1);

  const filtered = useMemo(() => {
    let list = trades.filter(t =>
      !search || t.symbol.toLowerCase().includes(search.toLowerCase()) || (t.setup || "").toLowerCase().includes(search.toLowerCase())
    );
    list = [...list].sort((a, b) => {
      let av = sortKey === "pnl" ? computePnl(a) : a[sortKey];
      let bv = sortKey === "pnl" ? computePnl(b) : b[sortKey];
      if (av < bv) return -1 * sortDir;
      if (av > bv) return 1 * sortDir;
      return 0;
    });
    return list;
  }, [trades, search, sortKey, sortDir]);

  const Th = ({ label, k }) => (
    <th onClick={() => { setSortKey(k); setSortDir(d => (sortKey === k ? -d : -1)); }}
      style={{ padding: "10px 12px", textAlign: "left", fontSize: 11, color: "#8A8D97", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.03em", cursor: "pointer", whiteSpace: "nowrap" }}>
      {label} {sortKey === k && (sortDir === 1 ? "↑" : "↓")}
    </th>
  );

  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
        <SectionTitle icon={ListChecks}>Trade Log ({filtered.length})</SectionTitle>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <div style={{ position: "relative" }}>
            <Search size={14} style={{ position: "absolute", left: 10, top: 10, color: "#5C5F68" }} />
            <input className="tj-input" placeholder="Search symbol or setup..." value={search} onChange={e => setSearch(e.target.value)} style={{ paddingLeft: 30, width: 220 }} />
          </div>
          <ColumnManager customColumns={customColumns} onAddColumn={onAddColumn} onRemoveColumn={onRemoveColumn} />
        </div>
      </div>
      <div className="tj-scroll" style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
              <Th label="Date" k="entryTime" /><Th label="Symbol" k="symbol" /><Th label="Side" k="side" />
              <Th label="Lots" k="lots" /><Th label="Entry" k="entryPrice" /><Th label="Exit" k="exitPrice" />
              <Th label="R" k="pnl" /><Th label="P&L" k="pnl" /><Th label="Setup" k="setup" /><Th label="Session" k="entryTime" />
              {customColumns.map(col => (
                <th key={col.id} style={{ padding: "10px 12px", textAlign: "left", fontSize: 11, color: "#8A8D97", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.03em", whiteSpace: "nowrap" }}>{col.label}</th>
              ))}
              <th style={{ padding: "10px 12px" }}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map(t => {
              const pnl = computePnl(t);
              const r = rMultiple(t);
              return (
                <tr key={t.id} className="tj-row-hover" style={{ borderBottom: "1px solid rgba(255,255,255,0.04)" }}>
                  <td className="tj-mono" style={{ padding: "10px 12px", fontSize: 12, color: "#8A8D97", whiteSpace: "nowrap" }}>{t.entryTime?.slice(0, 16).replace("T", " ")}</td>
                  <td style={{ padding: "10px 12px", fontSize: 12.5, fontWeight: 700 }}>{t.symbol}</td>
                  <td style={{ padding: "10px 12px" }}>
                    <Pill tone={t.side === "buy" ? "good" : "bad"}>{t.side === "buy" ? "BUY" : "SELL"}</Pill>
                  </td>
                  <td className="tj-mono" style={{ padding: "10px 12px", fontSize: 12.5 }}>{t.lots}</td>
                  <td className="tj-mono" style={{ padding: "10px 12px", fontSize: 12.5 }}>{t.entryPrice}</td>
                  <td className="tj-mono" style={{ padding: "10px 12px", fontSize: 12.5 }}>{t.exitPrice}</td>
                  <td className="tj-mono" style={{ padding: "10px 12px", fontSize: 12.5, color: r == null ? "#6B6E78" : r >= 0 ? "#34D399" : "#F45B69" }}>{r == null ? "—" : `${r.toFixed(2)}R`}</td>
                  <td className="tj-mono" style={{ padding: "10px 12px", fontSize: 12.5, fontWeight: 700, color: pnl >= 0 ? "#34D399" : "#F45B69" }}>{fmt$(pnl)}</td>
                  <td style={{ padding: "10px 12px", fontSize: 11.5, color: "#8A8D97" }}>{t.setup || "—"}</td>
                  <td style={{ padding: "10px 12px", fontSize: 11.5, color: "#8A8D97" }}>{t.entryTime ? getSession(new Date(t.entryTime).getUTCHours()) : "—"}</td>
                  {customColumns.map(col => (
                    <td key={col.id} style={{ padding: "10px 12px", fontSize: 12, color: "#C9CBD3", whiteSpace: "nowrap" }}>{t.customFields?.[col.id] || "—"}</td>
                  ))}
                  <td style={{ padding: "10px 12px" }}>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button className="tj-btn tj-btn-ghost" style={{ padding: 6 }} onClick={() => onEdit(t)}><Edit3 size={13} /></button>
                      <button className="tj-btn tj-btn-danger" style={{ padding: 6 }} onClick={() => onDelete(t.id)}><Trash2 size={13} /></button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={11 + customColumns.length} style={{ padding: 30, textAlign: "center", color: "#6B6E78", fontSize: 13 }}>No trades match. Log a trade or import your broker statement to get started.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== ADVANCED METRICS GRID ============================== */
function AdvancedMetrics({ m }) {
  const items = [
    { label: "Win Rate", value: fmtPct(m.winRate), tone: m.winRate >= 0.5 ? "good" : "bad" },
    { label: "Loss Rate", value: fmtPct(m.lossRate), tone: "neutral" },
    { label: "Profit Factor", value: m.profitFactor === Infinity ? "∞" : m.profitFactor.toFixed(2), tone: m.profitFactor >= 1.5 ? "good" : m.profitFactor >= 1 ? "warn" : "bad" },
    { label: "Expectancy", value: fmt$(m.expectancy), tone: m.expectancy >= 0 ? "good" : "bad" },
    { label: "Avg Winner", value: fmt$(m.avgWin), tone: "good" },
    { label: "Avg Loser", value: fmt$(m.avgLoss), tone: "bad" },
    { label: "Win/Loss Ratio", value: m.avgWinLossRatio.toFixed(2), tone: m.avgWinLossRatio >= 1.5 ? "good" : "warn" },
    { label: "Largest Win", value: fmt$(m.largestWin), tone: "good" },
    { label: "Largest Loss", value: fmt$(m.largestLoss), tone: "bad" },
    { label: "Max Drawdown", value: `${fmt$(-m.maxDD)} (${m.maxDDPct.toFixed(1)}%)`, tone: "bad" },
    { label: "Recovery Factor", value: m.recoveryFactor === Infinity ? "∞" : m.recoveryFactor.toFixed(2), tone: "good" },
    { label: "Sharpe Ratio", value: m.sharpe.toFixed(2), tone: m.sharpe >= 1 ? "good" : "warn" },
    { label: "Sortino Ratio", value: m.sortino.toFixed(2), tone: m.sortino >= 1 ? "good" : "warn" },
    { label: "Calmar Ratio", value: m.calmar.toFixed(2), tone: m.calmar >= 1 ? "good" : "warn" },
    { label: "SQN", value: m.sqn.toFixed(2), tone: m.sqn >= 2 ? "good" : "warn" },
    { label: "Kelly %", value: `${m.kellyPct.toFixed(1)}%`, tone: m.kellyPct > 0 ? "good" : "bad" },
    { label: "Risk of Ruin (est.)", value: `${m.riskOfRuin.toFixed(1)}%`, tone: m.riskOfRuin < 5 ? "good" : m.riskOfRuin < 20 ? "warn" : "bad" },
    { label: "Max Win Streak", value: m.maxWinStreak, tone: "good" },
    { label: "Max Loss Streak", value: m.maxLossStreak, tone: "bad" },
    { label: "Avg Win Duration", value: `${m.avgWinDur.toFixed(0)}m`, tone: "neutral" },
    { label: "Avg Loss Duration", value: `${m.avgLossDur.toFixed(0)}m`, tone: "neutral" },
  ];
  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={Shield}>Institutional-Grade Metrics</SectionTitle>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(150px,1fr))", gap: 12 }}>
        {items.map((it, i) => {
          const color = it.tone === "good" ? "#34D399" : it.tone === "bad" ? "#F45B69" : it.tone === "warn" ? "#E3B24D" : "#F2F1ED";
          return (
            <div key={i} style={{ background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 10, padding: "10px 12px" }}>
              <div style={{ fontSize: 10.5, color: "#6B6E78", fontWeight: 600, textTransform: "uppercase" }}>{it.label}</div>
              <div className="tj-mono" style={{ fontSize: 16, fontWeight: 700, color, marginTop: 4 }}>{it.value}</div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ============================== AI COACH CHAT ============================== */
function AiCoach({ trades, metrics, psychology }) {
  const [messages, setMessages] = useState([
    { role: "assistant", text: "I'm your AI trading coach. I can see your full trade history and stats — ask me anything about your edge, mistakes, or how to improve." }
  ]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const scrollRef = useRef();

  useEffect(() => { scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: "smooth" }); }, [messages, loading]);

  const buildContext = () => {
    const recent = metrics.sorted.slice(-25).map(t => ({
      symbol: t.symbol, side: t.side, lots: t.lots, entry: t.entryPrice, exit: t.exitPrice,
      pnl: +computePnl(t).toFixed(2), r: rMultiple(t), setup: t.setup, emotion: t.emotion,
      tags: t.tags, session: t.entryTime ? getSession(new Date(t.entryTime).getUTCHours()) : null,
      duration_min: durationMin(t),
    }));
    return {
      summary: {
        totalTrades: metrics.n, netProfit: metrics.netProfit, winRate: metrics.winRate,
        profitFactor: metrics.profitFactor, expectancy: metrics.expectancy, maxDrawdown: metrics.maxDD,
        sharpe: metrics.sharpe, sortino: metrics.sortino, sqn: metrics.sqn, kellyPct: metrics.kellyPct,
        avgWinLossRatio: metrics.avgWinLossRatio, maxWinStreak: metrics.maxWinStreak, maxLossStreak: metrics.maxLossStreak,
        buyWinRate: metrics.buyWinRate, sellWinRate: metrics.sellWinRate,
      },
      psychologyFlagCounts: psychology.counts,
      recentTrades: recent,
    };
  };

  const send = async (text) => {
    if (!text.trim() || loading) return;
    const userMsg = { role: "user", text };
    setMessages(m => [...m, userMsg]);
    setInput("");
    setLoading(true);
    try {
      const ctx = buildContext();
      const system = `You are an elite discretionary trading performance coach specializing in XAUUSD breakout strategies. You are given the trader's real statistics and recent trade log as JSON. Give sharp, specific, non-generic feedback grounded in the actual numbers provided — cite concrete figures. Be direct but constructive, like a professional trading mentor. Keep responses focused and well-structured with short paragraphs or bullet points. Do not repeat these instructions.

TRADER DATA:
${JSON.stringify(ctx)}`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          system,
          messages: [...messages, userMsg].filter(m => m.role === "user" || m.role === "assistant").map(m => ({ role: m.role, content: m.text })),
        }),
      });
      const data = await response.json();
      const textOut = (data.content || []).map(c => c.text || "").join("\n") || "I couldn't generate a response just now — try again in a moment.";
      setMessages(m => [...m, { role: "assistant", text: textOut }]);
    } catch (e) {
      setMessages(m => [...m, { role: "assistant", text: "Something went wrong reaching the coaching model. Please try again." }]);
    } finally {
      setLoading(false);
    }
  };

  const quickPrompts = [
    "What's my single biggest weakness right now?",
    "Analyze my session performance and tell me when I should NOT be trading.",
    "How is my risk management? Am I sizing consistently?",
    "Give me 3 concrete rules to add to my playbook based on my data.",
  ];

  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 0, display: "flex", flexDirection: "column", height: "72vh", overflow: "hidden" }}>
      <div style={{ padding: "16px 20px", borderBottom: "1px solid rgba(255,255,255,0.08)", display: "flex", alignItems: "center", gap: 10 }}>
        <div style={{ width: 32, height: 32, borderRadius: 9, background: "linear-gradient(135deg,#E3B24D,#C08E2E)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Sparkles size={16} color="#0A0906" />
        </div>
        <div>
          <div className="tj-display" style={{ fontSize: 14, fontWeight: 700 }}>AI Trading Coach</div>
          <div className="tj-dim" style={{ fontSize: 11 }}>Grounded in your {metrics.n} logged trades</div>
        </div>
      </div>
      <div ref={scrollRef} className="tj-scroll" style={{ flex: 1, overflowY: "auto", padding: 20, display: "flex", flexDirection: "column", gap: 14 }}>
        {messages.map((m, i) => (
          <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start" }}>
            <div style={{
              maxWidth: "78%", padding: "10px 14px", borderRadius: 14, fontSize: 13, lineHeight: 1.55, whiteSpace: "pre-wrap",
              background: m.role === "user" ? "linear-gradient(135deg,#E3B24D,#C08E2E)" : "rgba(255,255,255,0.045)",
              color: m.role === "user" ? "#0A0906" : "#F2F1ED",
              border: m.role === "user" ? "none" : "1px solid rgba(255,255,255,0.08)",
              fontWeight: m.role === "user" ? 600 : 400,
            }}>{m.text}</div>
          </div>
        ))}
        {loading && (
          <div style={{ display: "flex", justifyContent: "flex-start" }}>
            <div className="tj-pulse tj-dim" style={{ padding: "10px 14px", fontSize: 12.5 }}>Analyzing your trade data…</div>
          </div>
        )}
      </div>
      <div style={{ padding: "12px 16px", borderTop: "1px solid rgba(255,255,255,0.08)" }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
          {quickPrompts.map((q, i) => (
            <button key={i} className="tj-tag" style={{ cursor: "pointer" }} onClick={() => send(q)}>{q}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <input className="tj-input" placeholder="Ask about your trading data..." value={input}
            onChange={e => setInput(e.target.value)} onKeyDown={e => e.key === "Enter" && send(input)} />
          <button className="tj-btn tj-btn-primary" onClick={() => send(input)} disabled={loading}><Send size={14} /></button>
        </div>
      </div>
    </div>
  );
}

/* ============================== GOALS ============================== */
const GOAL_METRICS = [
  { key: "netProfit", label: "Net Profit ($)", getCurrent: (m) => Math.max(0, m.netProfit), invertDefault: false, suggestedTarget: 2000 },
  { key: "winRate", label: "Win Rate (%)", getCurrent: (m) => m.winRate * 100, invertDefault: false, suggestedTarget: 55 },
  { key: "profitFactor", label: "Profit Factor", getCurrent: (m) => (m.profitFactor === Infinity ? 99 : m.profitFactor), invertDefault: false, suggestedTarget: 1.8 },
  { key: "maxDDPct", label: "Max Drawdown (%)", getCurrent: (m) => m.maxDDPct, invertDefault: true, suggestedTarget: 10 },
  { key: "expectancy", label: "Expectancy ($/trade)", getCurrent: (m) => m.expectancy, invertDefault: false, suggestedTarget: 20 },
  { key: "tradeCount", label: "Trade Count", getCurrent: (m) => m.n, invertDefault: false, suggestedTarget: 50 },
  { key: "sqn", label: "SQN", getCurrent: (m) => m.sqn, invertDefault: false, suggestedTarget: 2.5 },
  { key: "maxLossStreak", label: "Max Loss Streak", getCurrent: (m) => m.maxLossStreak, invertDefault: true, suggestedTarget: 3 },
];

const DEFAULT_GOALS = [
  { id: "g1", label: "Monthly Profit Target", metricKey: "netProfit", target: 2000, invert: false },
  { id: "g2", label: "Win Rate ≥ 55%", metricKey: "winRate", target: 55, invert: false },
  { id: "g3", label: "Max Drawdown ≤ 10%", metricKey: "maxDDPct", target: 10, invert: true },
  { id: "g4", label: "Profit Factor ≥ 1.8", metricKey: "profitFactor", target: 1.8, invert: false },
];

function GoalEditor({ initial, onSave, onCancel }) {
  const [form, setForm] = useState(initial || { label: "", metricKey: "netProfit", target: "", invert: false });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const metricDef = GOAL_METRICS.find(m => m.key === form.metricKey);
  return (
    <div className="tj-panel-tight" style={{ padding: 16, background: "rgba(227,178,77,0.05)", border: "1px solid rgba(227,178,77,0.2)", marginBottom: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr 0.8fr auto", gap: 10, alignItems: "end" }}>
        <div>
          <label style={{ fontSize: 10.5, color: "#8A8D97" }}>Goal Name</label>
          <input className="tj-input" placeholder="e.g. Beat last month's P&L" value={form.label} onChange={e => set("label", e.target.value)} />
        </div>
        <div>
          <label style={{ fontSize: 10.5, color: "#8A8D97" }}>Metric</label>
          <select className="tj-input" value={form.metricKey} onChange={e => {
            const md = GOAL_METRICS.find(m => m.key === e.target.value);
            setForm(f => ({ ...f, metricKey: e.target.value, invert: md.invertDefault, target: md.suggestedTarget }));
          }}>
            {GOAL_METRICS.map(m => <option key={m.key} value={m.key}>{m.label}</option>)}
          </select>
        </div>
        <div>
          <label style={{ fontSize: 10.5, color: "#8A8D97" }}>Target</label>
          <input type="number" step="0.1" className="tj-input" value={form.target} onChange={e => set("target", e.target.value)} />
        </div>
        <button className="tj-btn tj-btn-primary" onClick={() => form.label.trim() && form.target !== "" && onSave({ ...form, target: parseFloat(form.target) })}>
          <Plus size={14} /> Save
        </button>
      </div>
      <label style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 10, fontSize: 12, color: "#8A8D97", cursor: "pointer" }}>
        <input type="checkbox" checked={form.invert} onChange={e => set("invert", e.target.checked)} />
        Lower is better (e.g. drawdown, loss streaks)
      </label>
      {onCancel && <button className="tj-btn tj-btn-ghost" style={{ marginTop: 10 }} onClick={onCancel}>Cancel</button>}
    </div>
  );
}

function GoalsPanel({ metrics, journalId }) {
  const [goals, setGoals] = useState(DEFAULT_GOALS);
  const [loaded, setLoaded] = useState(false);
  const [adding, setAdding] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const jid = journalId || "default";

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage?.get(`goals:${jid}`, false);
        if (res?.value) setGoals(JSON.parse(res.value));
      } catch (e) { /* keep defaults */ }
      setLoaded(true);
    })();
  }, [jid]);
  useEffect(() => { if (loaded) { try { window.storage?.set(`goals:${jid}`, JSON.stringify(goals), false); } catch (e) {} } }, [goals, loaded, jid]);

  const addGoal = (g) => { setGoals(gs => [...gs, { ...g, id: Date.now().toString(36) }]); setAdding(false); };
  const updateGoal = (id, g) => { setGoals(gs => gs.map(x => (x.id === id ? { ...g, id } : x))); setEditingId(null); };
  const removeGoal = (id) => setGoals(gs => gs.filter(g => g.id !== id));

  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={Award} right={
        <button className="tj-btn tj-btn-primary" onClick={() => setAdding(a => !a)}><Plus size={14} /> Add Goal</button>
      }>Goal Tracking</SectionTitle>

      {adding && <GoalEditor onSave={addGoal} onCancel={() => setAdding(false)} />}

      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        {goals.length === 0 && !adding && <div className="tj-dim" style={{ fontSize: 13 }}>No goals yet — click "Add Goal" to set your first target.</div>}
        {goals.map(g => {
          if (editingId === g.id) {
            return <GoalEditor key={g.id} initial={g} onSave={(ng) => updateGoal(g.id, ng)} onCancel={() => setEditingId(null)} />;
          }
          const metricDef = GOAL_METRICS.find(m => m.key === g.metricKey) || GOAL_METRICS[0];
          const current = metricDef.getCurrent(metrics);
          const pct = g.invert ? clamp(100 - (current / g.target) * 100, 0, 100) : clamp((current / g.target) * 100, 0, 100);
          const ok = g.invert ? current <= g.target : current >= g.target;
          return (
            <div key={g.id}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12.5, marginBottom: 5 }}>
                <span style={{ fontWeight: 600 }}>{g.label} <span className="tj-dim" style={{ fontWeight: 500 }}>({metricDef.label})</span></span>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span className="tj-mono" style={{ color: ok ? "#34D399" : "#E3B24D", fontWeight: 700 }}>{current.toFixed(1)} / {g.target}</span>
                  <button onClick={() => setEditingId(g.id)} style={{ background: "none", border: "none", color: "#8A8D97", cursor: "pointer" }}><Edit3 size={13} /></button>
                  <button onClick={() => removeGoal(g.id)} style={{ background: "none", border: "none", color: "#F45B69", cursor: "pointer" }}><Trash2 size={13} /></button>
                </div>
              </div>
              <div style={{ height: 8, background: "rgba(255,255,255,0.06)", borderRadius: 5, overflow: "hidden" }}>
                <div style={{ width: `${pct}%`, height: "100%", background: ok ? "linear-gradient(90deg,#34D399,#22B586)" : "linear-gradient(90deg,#E3B24D,#C08E2E)", transition: "width .4s ease" }} />
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ============================== DEEP ANALYSIS TAB ============================== */
function LeaderboardTable({ title, rows, icon: Icon }) {
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: 18 }}>
      <SectionTitle icon={Icon}>{title}</SectionTitle>
      {rows.length === 0 ? (
        <div className="tj-dim" style={{ fontSize: 12.5 }}>Not enough tagged data yet.</div>
      ) : (
        <div className="tj-scroll" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
                {["Name", "Trades", "Win Rate", "Expectancy", "Profit Factor", "Net P&L"].map((h, i) => (
                  <th key={i} style={{ padding: "7px 10px", fontSize: 10.5, color: "#8A8D97", fontWeight: 700, textTransform: "uppercase", textAlign: i === 0 ? "left" : "right" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i} className="tj-row-hover" style={{ borderBottom: "1px solid rgba(255,255,255,0.04)" }}>
                  <td style={{ padding: "8px 10px", fontSize: 12.5, fontWeight: 600 }}>{r.key}</td>
                  <td className="tj-mono" style={{ padding: "8px 10px", fontSize: 12, textAlign: "right" }}>{r.count}</td>
                  <td className="tj-mono" style={{ padding: "8px 10px", fontSize: 12, textAlign: "right", color: r.winRate >= 0.5 ? "#34D399" : "#F45B69" }}>{fmtPct(r.winRate)}</td>
                  <td className="tj-mono" style={{ padding: "8px 10px", fontSize: 12, textAlign: "right", color: r.expectancy >= 0 ? "#34D399" : "#F45B69" }}>{fmt$(r.expectancy)}</td>
                  <td className="tj-mono" style={{ padding: "8px 10px", fontSize: 12, textAlign: "right" }}>{r.profitFactor === Infinity ? "∞" : r.profitFactor.toFixed(2)}</td>
                  <td className="tj-mono" style={{ padding: "8px 10px", fontSize: 12, textAlign: "right", fontWeight: 700, color: r.netPnl >= 0 ? "#34D399" : "#F45B69" }}>{fmt$(r.netPnl)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function CustomColumnCorrelation({ trades, customColumns }) {
  const data = useMemo(() => computeCustomColumnStats(trades, customColumns), [trades, customColumns]);
  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={Layers}>Custom Column Correlation</SectionTitle>
      {data.length === 0 ? (
        <div className="tj-dim" style={{ fontSize: 12.5 }}>Fill in your custom columns (4H Candle, Breakout stage, etc.) on a few trades to see which values actually correlate with winners.</div>
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 16 }}>
          {data.map(({ column, stats }, i) => (
            <div key={i} style={{ background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 8 }}>{column.label}</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {stats.map((s, j) => (
                  <div key={j} style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5 }}>
                    <span className="tj-tag">{s.key}</span>
                    <span className="tj-mono" style={{ color: s.expectancy >= 0 ? "#34D399" : "#F45B69" }}>{fmtPct(s.winRate)} · {fmt$(s.expectancy)}</span>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function EfficiencyPanel({ trades }) {
  const eff = useMemo(() => computeEfficiency(trades), [trades]);
  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={Gauge}>Trade Efficiency (MFE / MAE)</SectionTitle>
      {!eff.available ? (
        <div className="tj-dim" style={{ fontSize: 12.5 }}>Log "Max Favorable Excursion" on your trades (in the trade form) to unlock this — it shows how much of the available move you actually captured.</div>
      ) : (
        <>
          <div style={{ display: "flex", gap: 24, marginBottom: 16 }}>
            <MiniStat label="Avg Efficiency" value={`${(eff.avgEfficiency * 100).toFixed(0)}%`} tone={eff.avgEfficiency >= 0.5 ? "good" : "warn"} />
            <MiniStat label="Trades Logged" value={eff.list.length} />
          </div>
          <div className="tj-dim" style={{ fontSize: 11.5, marginBottom: 8 }}>Lowest-efficiency trades — biggest gap between what the market offered and what you captured:</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {eff.list.slice(0, 5).map((t, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, padding: "6px 10px", background: "rgba(255,255,255,0.02)", borderRadius: 8 }}>
                <span>{t.symbol} · Trade #{t.id}</span>
                <span className="tj-mono">Captured {fmt$(t.pnl)} of {fmt$(t.mfe)} MFE ({(t.efficiency * 100).toFixed(0)}%)</span>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

function EdgeDecayChart({ trades }) {
  const [window, setWindowSize] = useState(15);
  const data = useMemo(() => computeRollingStats(trades, window), [trades, window]);
  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <SectionTitle icon={History}>Edge Decay — Rolling {window}-Trade Window</SectionTitle>
        <div style={{ display: "flex", gap: 6 }}>
          {[10, 15, 25].map(w => (
            <button key={w} className="tj-btn" style={{ padding: "6px 10px", background: window === w ? "rgba(227,178,77,0.14)" : "rgba(255,255,255,0.04)", color: window === w ? "#E3B24D" : "#F2F1ED", borderColor: window === w ? "rgba(227,178,77,0.3)" : "rgba(255,255,255,0.1)" }} onClick={() => setWindowSize(w)}>{w}</button>
          ))}
        </div>
      </div>
      {data.length < 5 ? <div className="tj-dim" style={{ fontSize: 12.5 }}>Need a few more trades to plot a rolling trend.</div> : (
        <div style={{ height: 200 }}>
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={data}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" vertical={false} />
              <XAxis dataKey="x" tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} />
              <YAxis yAxisId="left" tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} width={35} />
              <YAxis yAxisId="right" orientation="right" tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} width={40} />
              <Tooltip content={<CustomTooltip prefix="" />} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Line yAxisId="left" type="monotone" dataKey="winRate" name="Win Rate %" stroke="#5FB4E8" strokeWidth={2} dot={false} />
              <Line yAxisId="right" type="monotone" dataKey="expectancy" name="Expectancy $" stroke="#E3B24D" strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

function DayAfterPanel({ trades }) {
  const stats = useMemo(() => computeDayAfterStats(trades), [trades]);
  const cards = [
    { key: "afterLoss", label: "Day After a Loss", tone: "bad" },
    { key: "afterWin", label: "Day After a Win", tone: "good" },
    { key: "afterFlat", label: "Day After Flat", tone: "neutral" },
  ];
  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={GitCompare}>Day-After Behavior</SectionTitle>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
        {cards.map(c => {
          const s = stats[c.key];
          return (
            <div key={c.key} style={{ background: "rgba(255,255,255,0.02)", border: "1px solid rgba(255,255,255,0.06)", borderRadius: 10, padding: 14 }}>
              <div className="tj-dim" style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", marginBottom: 8 }}>{c.label}</div>
              {s.count === 0 ? <div className="tj-dim" style={{ fontSize: 12 }}>No data</div> : (
                <>
                  <div className="tj-mono" style={{ fontSize: 18, fontWeight: 700, color: s.expectancy >= 0 ? "#34D399" : "#F45B69" }}>{fmt$(s.expectancy)}<span style={{ fontSize: 11, color: "#6B6E78", fontWeight: 500 }}> / trade</span></div>
                  <div style={{ fontSize: 11.5, color: "#8A8D97", marginTop: 4 }}>{fmtPct(s.winRate)} win rate · {s.count} trades</div>
                </>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function DisciplineScorePanel({ trades, flags }) {
  const result = useMemo(() => computeDisciplineScore(trades, flags), [trades, flags]);
  const color = result.score >= 75 ? "#34D399" : result.score >= 50 ? "#E3B24D" : "#F45B69";
  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={Shield}>Discipline &amp; Consistency Score</SectionTitle>
      <div style={{ display: "flex", alignItems: "center", gap: 24, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ position: "relative", width: 100, height: 100 }}>
          <svg viewBox="0 0 100 100" style={{ transform: "rotate(-90deg)" }}>
            <circle cx="50" cy="50" r="42" fill="none" stroke="rgba(255,255,255,0.08)" strokeWidth="9" />
            <circle cx="50" cy="50" r="42" fill="none" stroke={color} strokeWidth="9" strokeLinecap="round"
              strokeDasharray={`${(result.score / 100) * 264} 264`} />
          </svg>
          <div className="tj-mono" style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, fontWeight: 700, color }}>{result.score}</div>
        </div>
        <div style={{ display: "flex", gap: 20 }}>
          <MiniStat label="Flag Discipline" value={Math.round(result.flagPenalty || 0)} tone="gold" />
          <MiniStat label="Sizing Consistency" value={Math.round(result.sizingScore || 0)} tone="gold" />
          <MiniStat label="Rule Adherence" value={Math.round(result.ruleScore || 0)} tone="gold" />
        </div>
      </div>
      {result.trend.length > 1 && (
        <div style={{ height: 140 }}>
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={result.trend}>
              <XAxis dataKey="week" tick={{ fill: "#5C5F68", fontSize: 9 }} axisLine={false} tickLine={false} />
              <YAxis domain={[0, 100]} tick={{ fill: "#5C5F68", fontSize: 10 }} axisLine={false} tickLine={false} width={30} />
              <Tooltip content={<CustomTooltip prefix="" />} />
              <Line type="monotone" dataKey="score" stroke="#E3B24D" strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

function CostDragPanel({ trades }) {
  const c = useMemo(() => computeCostDrag(trades), [trades]);
  return (
    <div className="tj-panel tj-panel-tight tj-fade-in" style={{ padding: 18 }}>
      <SectionTitle icon={Coins}>Cost Drag</SectionTitle>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        <MiniStat label="Total Commission" value={fmt$(-c.totalCommission)} tone="bad" />
        <MiniStat label="Total Swap" value={fmt$(-c.totalSwap)} tone="bad" />
        <MiniStat label="% of Gross Profit Consumed" value={`${c.pctOfGrossProfit.toFixed(1)}%`} tone={c.pctOfGrossProfit > 15 ? "bad" : "gold"} />
      </div>
    </div>
  );
}

function WhatIfSimulator({ trades, flags }) {
  const flagTypes = useMemo(() => Array.from(new Set(flags.map(f => f.type))), [flags]);
  const [excluded, setExcluded] = useState([]);
  const result = useMemo(() => computeWhatIf(trades, flags, excluded), [trades, flags, excluded]);
  const toggle = (type) => setExcluded(ex => ex.includes(type) ? ex.filter(t => t !== type) : [...ex, type]);

  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={FlaskConical}>What-If Simulator</SectionTitle>
      <div className="tj-dim" style={{ fontSize: 12, marginBottom: 12 }}>Select behavior patterns to strip out and see what your equity curve would look like without them.</div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 16 }}>
        {flagTypes.length === 0 && <div className="tj-dim" style={{ fontSize: 12.5 }}>No flagged patterns to simulate yet.</div>}
        {flagTypes.map(type => (
          <button key={type} onClick={() => toggle(type)} className="tj-tag" style={{
            cursor: "pointer", background: excluded.includes(type) ? "rgba(244,91,105,0.16)" : undefined,
            borderColor: excluded.includes(type) ? "rgba(244,91,105,0.4)" : undefined,
            color: excluded.includes(type) ? "#F45B69" : undefined,
          }}>{excluded.includes(type) ? "✕ " : ""}{type}</button>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 12 }}>
        <MiniStat label="Actual Net P&L" value={fmt$(result.actual.netProfit)} />
        <MiniStat label="Without Selected Patterns" value={fmt$(result.hypothetical.netProfit)} tone={result.delta >= 0 ? "good" : "bad"} />
        <MiniStat label="Cost of That Behavior" value={fmt$(result.delta)} tone={result.delta >= 0 ? "good" : "bad"} />
        <MiniStat label="Trades Removed" value={result.excludedCount} />
      </div>
    </div>
  );
}

function PeriodComparisonPanel({ trades }) {
  const [type, setType] = useState("week");
  const { current, previous } = useMemo(() => computePeriodComparison(trades, type), [trades, type]);
  const rows = [
    { label: "Net P&L", cur: fmt$(current.netProfit), prev: fmt$(previous.netProfit), delta: current.netProfit - previous.netProfit },
    { label: "Win Rate", cur: fmtPct(current.winRate), prev: fmtPct(previous.winRate), delta: (current.winRate - previous.winRate) * 100 },
    { label: "Trades", cur: current.n, prev: previous.n, delta: current.n - previous.n },
    { label: "Profit Factor", cur: current.profitFactor === Infinity ? "∞" : current.profitFactor.toFixed(2), prev: previous.profitFactor === Infinity ? "∞" : previous.profitFactor.toFixed(2), delta: (current.profitFactor === Infinity ? 0 : current.profitFactor) - (previous.profitFactor === Infinity ? 0 : previous.profitFactor) },
    { label: "Expectancy", cur: fmt$(current.expectancy), prev: fmt$(previous.expectancy), delta: current.expectancy - previous.expectancy },
  ];
  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <SectionTitle icon={GitCompare}>Period Comparison</SectionTitle>
        <div style={{ display: "flex", gap: 6 }}>
          {[{ k: "week", l: "This Week vs Last" }, { k: "month", l: "This Month vs Last" }].map(o => (
            <button key={o.k} className="tj-btn" style={{ padding: "6px 12px", background: type === o.k ? "rgba(227,178,77,0.14)" : "rgba(255,255,255,0.04)", color: type === o.k ? "#E3B24D" : "#F2F1ED", borderColor: type === o.k ? "rgba(227,178,77,0.3)" : "rgba(255,255,255,0.1)" }} onClick={() => setType(o.k)}>{o.l}</button>
          ))}
        </div>
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
            {["Metric", "Current", "Previous", "Change"].map((h, i) => <th key={i} style={{ padding: "7px 10px", fontSize: 10.5, color: "#8A8D97", fontWeight: 700, textTransform: "uppercase", textAlign: i === 0 ? "left" : "right" }}>{h}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i} style={{ borderBottom: "1px solid rgba(255,255,255,0.04)" }}>
              <td style={{ padding: "9px 10px", fontSize: 12.5, fontWeight: 600 }}>{r.label}</td>
              <td className="tj-mono" style={{ padding: "9px 10px", fontSize: 12.5, textAlign: "right" }}>{r.cur}</td>
              <td className="tj-mono tj-dim" style={{ padding: "9px 10px", fontSize: 12.5, textAlign: "right" }}>{r.prev}</td>
              <td className="tj-mono" style={{ padding: "9px 10px", fontSize: 12.5, textAlign: "right", fontWeight: 700, color: r.delta >= 0 ? "#34D399" : "#F45B69" }}>{r.delta >= 0 ? "+" : ""}{typeof r.delta === "number" ? r.delta.toFixed(1) : r.delta}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function InsightsPanel({ trades, metrics, psychology, customColumns }) {
  const { bySetup, byTag } = useMemo(() => computeSetupTagStats(trades), [trades]);
  return (
    <div className="tj-fade-in" style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 }}>
        <LeaderboardTable title="Performance by Setup" rows={bySetup} icon={Target} />
        <LeaderboardTable title="Performance by Tag" rows={byTag} icon={Tags} />
      </div>
      <CustomColumnCorrelation trades={trades} customColumns={customColumns} />
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 }}>
        <EfficiencyPanel trades={trades} />
        <DayAfterPanel trades={trades} />
      </div>
      <EdgeDecayChart trades={trades} />
      <DisciplineScorePanel trades={trades} flags={psychology.flags} />
      <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 18 }}>
        <CostDragPanel trades={trades} />
        <PeriodComparisonPanel trades={trades} />
      </div>
      <WhatIfSimulator trades={trades} flags={psychology.flags} />
    </div>
  );
}

/* ============================== AUTHENTICATION ============================== */
async function sha256Hex(text) {
  try {
    const enc = new TextEncoder().encode(text);
    const buf = await crypto.subtle.digest("SHA-256", enc);
    return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
  } catch (e) {
    // fallback simple hash if Web Crypto isn't available
    let h = 0;
    for (let i = 0; i < text.length; i++) { h = (h << 5) - h + text.charCodeAt(i); h |= 0; }
    return "fallback_" + h.toString(16);
  }
}
const isValidEmail = (e) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);

async function getAuthUsers() {
  try {
    const res = await window.storage?.get("authUsers", true);
    return res?.value ? JSON.parse(res.value) : [];
  } catch (e) { return []; }
}
async function setAuthUsers(users) {
  try { await window.storage?.set("authUsers", JSON.stringify(users), true); } catch (e) {}
}

function AuthShell({ children, title, subtitle }) {
  return (
    <div className="tj-root" style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <GlobalStyle />
      <div className="tj-bg-glow" />
      <div className="tj-panel tj-fade-in" style={{ width: 400, padding: 32, position: "relative", zIndex: 1 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 22 }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: "linear-gradient(135deg,#E3B24D,#8A6C2E)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <TrendingUp size={18} color="#0A0906" />
          </div>
          <div className="tj-display" style={{ fontSize: 18, fontWeight: 700 }}>Aurlis<span className="tj-gold">.</span></div>
        </div>
        <h2 className="tj-display" style={{ fontSize: 19, fontWeight: 700, marginBottom: 4 }}>{title}</h2>
        {subtitle && <p className="tj-dim" style={{ fontSize: 12.5, marginBottom: 22 }}>{subtitle}</p>}
        {children}
      </div>
    </div>
  );
}

function AuthField({ icon: Icon, ...props }) {
  return (
    <div style={{ position: "relative", marginBottom: 12 }}>
      <Icon size={15} style={{ position: "absolute", left: 12, top: 12, color: "#5C5F68" }} />
      <input className="tj-input" style={{ paddingLeft: 36 }} {...props} />
    </div>
  );
}

function AuthGate({ onAuthenticated }) {
  const [mode, setMode] = useState("login"); // login | signup | forgot
  const [checkingSession, setCheckingSession] = useState(true);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");

  // login
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [showLoginPw, setShowLoginPw] = useState(false);

  // signup
  const [suEmail, setSuEmail] = useState("");
  const [suPassword, setSuPassword] = useState("");
  const [suConfirm, setSuConfirm] = useState("");
  const [suQuestion, setSuQuestion] = useState("What was the name of your first trading strategy?");
  const [suAnswer, setSuAnswer] = useState("");
  const [showSuPw, setShowSuPw] = useState(false);

  // forgot password
  const [fpStep, setFpStep] = useState(1);
  const [fpEmail, setFpEmail] = useState("");
  const [fpUser, setFpUser] = useState(null);
  const [fpAnswer, setFpAnswer] = useState("");
  const [fpNewPassword, setFpNewPassword] = useState("");
  const [fpConfirmPassword, setFpConfirmPassword] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage?.get("authSession", false);
        if (res?.value) {
          const users = await getAuthUsers();
          if (users.some(u => u.email === res.value)) { onAuthenticated(res.value); return; }
        }
      } catch (e) { /* no session */ }
      setCheckingSession(false);
    })();
  }, []);

  const resetMessages = () => { setError(""); setInfo(""); };

  const handleLogin = async () => {
    resetMessages();
    if (!isValidEmail(loginEmail)) return setError("Enter a valid email address.");
    if (!loginPassword) return setError("Enter your password.");
    setLoading(true);
    try {
      const users = await getAuthUsers();
      const user = users.find(u => u.email.toLowerCase() === loginEmail.trim().toLowerCase());
      if (!user) { setError("No account found with that email."); setLoading(false); return; }
      const hash = await sha256Hex(loginPassword);
      if (hash !== user.passwordHash) { setError("Incorrect password."); setLoading(false); return; }
      try { await window.storage?.set("authSession", user.email, false); } catch (e) {}
      onAuthenticated(user.email);
    } catch (e) {
      setError("Something went wrong logging in. Please try again.");
    } finally { setLoading(false); }
  };

  const handleSignup = async () => {
    resetMessages();
    if (!isValidEmail(suEmail)) return setError("Enter a valid email address.");
    if (suPassword.length < 6) return setError("Password must be at least 6 characters.");
    if (suPassword !== suConfirm) return setError("Passwords don't match.");
    if (!suAnswer.trim()) return setError("Answer your security question — it's used to recover your password later.");
    setLoading(true);
    try {
      const users = await getAuthUsers();
      if (users.some(u => u.email.toLowerCase() === suEmail.trim().toLowerCase())) {
        setError("An account with that email already exists."); setLoading(false); return;
      }
      const passwordHash = await sha256Hex(suPassword);
      const securityAnswerHash = await sha256Hex(suAnswer.trim().toLowerCase());
      const newUser = { email: suEmail.trim(), passwordHash, securityQuestion: suQuestion, securityAnswerHash, createdAt: new Date().toISOString() };
      await setAuthUsers([...users, newUser]);
      try { await window.storage?.set("authSession", newUser.email, false); } catch (e) {}
      onAuthenticated(newUser.email);
    } catch (e) {
      setError("Something went wrong creating your account. Please try again.");
    } finally { setLoading(false); }
  };

  const fpFindAccount = async () => {
    resetMessages();
    if (!isValidEmail(fpEmail)) return setError("Enter a valid email address.");
    setLoading(true);
    try {
      const users = await getAuthUsers();
      const user = users.find(u => u.email.toLowerCase() === fpEmail.trim().toLowerCase());
      if (!user) { setError("No account found with that email."); setLoading(false); return; }
      setFpUser(user);
      setFpStep(2);
    } catch (e) { setError("Something went wrong. Please try again."); }
    finally { setLoading(false); }
  };

  const fpVerifyAnswer = async () => {
    resetMessages();
    if (!fpAnswer.trim()) return setError("Enter your answer.");
    setLoading(true);
    try {
      const hash = await sha256Hex(fpAnswer.trim().toLowerCase());
      if (hash !== fpUser.securityAnswerHash) { setError("That answer doesn't match our records."); setLoading(false); return; }
      setFpStep(3);
    } catch (e) { setError("Something went wrong. Please try again."); }
    finally { setLoading(false); }
  };

  const fpResetPassword = async () => {
    resetMessages();
    if (fpNewPassword.length < 6) return setError("Password must be at least 6 characters.");
    if (fpNewPassword !== fpConfirmPassword) return setError("Passwords don't match.");
    setLoading(true);
    try {
      const users = await getAuthUsers();
      const passwordHash = await sha256Hex(fpNewPassword);
      const updated = users.map(u => (u.email === fpUser.email ? { ...u, passwordHash } : u));
      await setAuthUsers(updated);
      setMode("login");
      setLoginEmail(fpUser.email);
      setInfo("Password reset — sign in with your new password.");
      setFpStep(1); setFpEmail(""); setFpUser(null); setFpAnswer(""); setFpNewPassword(""); setFpConfirmPassword("");
    } catch (e) { setError("Something went wrong resetting your password."); }
    finally { setLoading(false); }
  };

  if (checkingSession) {
    return (
      <div className="tj-root" style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <GlobalStyle />
        <RefreshCw size={22} className="tj-gold tj-pulse" />
      </div>
    );
  }

  const ErrorInfo = () => (
    <>
      {error && <div style={{ background: "rgba(244,91,105,0.1)", border: "1px solid rgba(244,91,105,0.3)", color: "#F45B69", fontSize: 12, padding: "8px 12px", borderRadius: 8, marginBottom: 12 }}>{error}</div>}
      {info && <div style={{ background: "rgba(52,211,153,0.1)", border: "1px solid rgba(52,211,153,0.3)", color: "#34D399", fontSize: 12, padding: "8px 12px", borderRadius: 8, marginBottom: 12, display: "flex", alignItems: "center", gap: 6 }}><CheckCircle2 size={13} />{info}</div>}
    </>
  );

  if (mode === "login") {
    return (
      <AuthShell title="Welcome back" subtitle="Sign in to your trading journal.">
        <ErrorInfo />
        <AuthField icon={Mail} type="email" placeholder="Email address" value={loginEmail} onChange={e => setLoginEmail(e.target.value)} onKeyDown={e => e.key === "Enter" && handleLogin()} />
        <div style={{ position: "relative", marginBottom: 6 }}>
          <Lock size={15} style={{ position: "absolute", left: 12, top: 12, color: "#5C5F68" }} />
          <input className="tj-input" style={{ paddingLeft: 36, paddingRight: 36 }} type={showLoginPw ? "text" : "password"} placeholder="Password" value={loginPassword} onChange={e => setLoginPassword(e.target.value)} onKeyDown={e => e.key === "Enter" && handleLogin()} />
          <button onClick={() => setShowLoginPw(s => !s)} style={{ position: "absolute", right: 10, top: 9, background: "none", border: "none", color: "#5C5F68", cursor: "pointer" }}>
            {showLoginPw ? <EyeOff size={15} /> : <Eye size={15} />}
          </button>
        </div>
        <div style={{ textAlign: "right", marginBottom: 16 }}>
          <span className="tj-gold" style={{ fontSize: 12, cursor: "pointer" }} onClick={() => { resetMessages(); setMode("forgot"); }}>Forgot password?</span>
        </div>
        <button className="tj-btn tj-btn-primary" style={{ width: "100%", justifyContent: "center", padding: 11 }} onClick={handleLogin} disabled={loading}>
          {loading ? "Signing in…" : "Sign In"}
        </button>
        <div className="tj-dim" style={{ fontSize: 12.5, textAlign: "center", marginTop: 18 }}>
          Don't have an account? <span className="tj-gold" style={{ cursor: "pointer", fontWeight: 700 }} onClick={() => { resetMessages(); setMode("signup"); }}>Create one</span>
        </div>
      </AuthShell>
    );
  }

  if (mode === "signup") {
    return (
      <AuthShell title="Create your account" subtitle="Your own private trading journal, secured by email and password.">
        <ErrorInfo />
        <AuthField icon={Mail} type="email" placeholder="Email address" value={suEmail} onChange={e => setSuEmail(e.target.value)} />
        <div style={{ position: "relative", marginBottom: 12 }}>
          <Lock size={15} style={{ position: "absolute", left: 12, top: 12, color: "#5C5F68" }} />
          <input className="tj-input" style={{ paddingLeft: 36, paddingRight: 36 }} type={showSuPw ? "text" : "password"} placeholder="Password (min 6 characters)" value={suPassword} onChange={e => setSuPassword(e.target.value)} />
          <button onClick={() => setShowSuPw(s => !s)} style={{ position: "absolute", right: 10, top: 9, background: "none", border: "none", color: "#5C5F68", cursor: "pointer" }}>
            {showSuPw ? <EyeOff size={15} /> : <Eye size={15} />}
          </button>
        </div>
        <AuthField icon={Lock} type="password" placeholder="Confirm password" value={suConfirm} onChange={e => setSuConfirm(e.target.value)} />
        <div style={{ marginBottom: 12 }}>
          <label style={{ fontSize: 11, color: "#8A8D97", fontWeight: 600 }}>Security Question (used for password recovery)</label>
          <select className="tj-input" style={{ marginTop: 4, marginBottom: 8 }} value={suQuestion} onChange={e => setSuQuestion(e.target.value)}>
            <option>What was the name of your first trading strategy?</option>
            <option>What city did you place your first trade from?</option>
            <option>What's your trading mentor's name?</option>
            <option>What was your first broker's name?</option>
          </select>
          <AuthField icon={KeyRound} type="text" placeholder="Your answer" value={suAnswer} onChange={e => setSuAnswer(e.target.value)} />
        </div>
        <button className="tj-btn tj-btn-primary" style={{ width: "100%", justifyContent: "center", padding: 11 }} onClick={handleSignup} disabled={loading}>
          {loading ? "Creating account…" : "Create Account"}
        </button>
        <div className="tj-dim" style={{ fontSize: 12.5, textAlign: "center", marginTop: 18 }}>
          Already have an account? <span className="tj-gold" style={{ cursor: "pointer", fontWeight: 700 }} onClick={() => { resetMessages(); setMode("login"); }}>Sign in</span>
        </div>
      </AuthShell>
    );
  }

  // forgot password
  return (
    <AuthShell title="Reset your password" subtitle={fpStep === 1 ? "Enter your account email to begin." : fpStep === 2 ? "Answer your security question." : "Choose a new password."}>
      <ErrorInfo />
      {fpStep === 1 && (
        <>
          <AuthField icon={Mail} type="email" placeholder="Email address" value={fpEmail} onChange={e => setFpEmail(e.target.value)} onKeyDown={e => e.key === "Enter" && fpFindAccount()} />
          <button className="tj-btn tj-btn-primary" style={{ width: "100%", justifyContent: "center", padding: 11 }} onClick={fpFindAccount} disabled={loading}>{loading ? "Checking…" : "Continue"}</button>
        </>
      )}
      {fpStep === 2 && (
        <>
          <div className="tj-dim" style={{ fontSize: 12.5, marginBottom: 10 }}>{fpUser?.securityQuestion}</div>
          <AuthField icon={KeyRound} type="text" placeholder="Your answer" value={fpAnswer} onChange={e => setFpAnswer(e.target.value)} onKeyDown={e => e.key === "Enter" && fpVerifyAnswer()} />
          <button className="tj-btn tj-btn-primary" style={{ width: "100%", justifyContent: "center", padding: 11 }} onClick={fpVerifyAnswer} disabled={loading}>{loading ? "Verifying…" : "Verify"}</button>
        </>
      )}
      {fpStep === 3 && (
        <>
          <AuthField icon={Lock} type="password" placeholder="New password (min 6 characters)" value={fpNewPassword} onChange={e => setFpNewPassword(e.target.value)} />
          <AuthField icon={Lock} type="password" placeholder="Confirm new password" value={fpConfirmPassword} onChange={e => setFpConfirmPassword(e.target.value)} onKeyDown={e => e.key === "Enter" && fpResetPassword()} />
          <button className="tj-btn tj-btn-primary" style={{ width: "100%", justifyContent: "center", padding: 11 }} onClick={fpResetPassword} disabled={loading}>{loading ? "Saving…" : "Reset Password"}</button>
        </>
      )}
      <div className="tj-dim" style={{ fontSize: 12.5, textAlign: "center", marginTop: 18, display: "flex", alignItems: "center", justifyContent: "center", gap: 5, cursor: "pointer" }}
        onClick={() => { resetMessages(); setMode("login"); setFpStep(1); }}>
        <ArrowLeft size={13} /> Back to sign in
      </div>
    </AuthShell>
  );
}

/* ============================== DATA EXPORT ============================== */
function filterTradesByDateRange(trades, from, to) {
  return trades.filter(t => {
    if (!t.entryTime) return false;
    const d = t.entryTime.slice(0, 10);
    return (!from || d >= from) && (!to || d <= to);
  });
}

function buildExportRows(trades, customColumns) {
  return [...trades].sort((a, b) => new Date(a.entryTime) - new Date(b.entryTime)).map(t => {
    const pnl = computePnl(t);
    const r = rMultiple(t);
    const base = {
      Date: t.entryTime ? t.entryTime.slice(0, 10) : "",
      "Entry Time": t.entryTime ? t.entryTime.slice(11, 16) : "",
      "Exit Time": t.exitTime ? t.exitTime.slice(11, 16) : "",
      Symbol: t.symbol, Side: t.side, Lots: t.lots,
      Entry: t.entryPrice, Exit: t.exitPrice, SL: t.sl ?? "", TP: t.tp ?? "",
      Commission: t.commission || 0, Swap: t.swap || 0,
      "P&L": +pnl.toFixed(2), "R-Multiple": r == null ? "" : +r.toFixed(2),
      Setup: t.setup || "", Emotion: t.emotion || "", Tags: (t.tags || []).join("; "),
      Session: t.entryTime ? getSession(new Date(t.entryTime).getUTCHours()) : "",
      "Duration (min)": durationMin(t) ?? "", Notes: t.notes || "",
    };
    (customColumns || []).forEach(col => { base[col.label] = (t.customFields && t.customFields[col.id]) || ""; });
    return base;
  });
}

function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportCSV(rows, filename) {
  const csv = Papa.unparse(rows);
  downloadBlob(csv, filename, "text/csv;charset=utf-8;");
}
function exportXLSX(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Trades");
  XLSX.writeFile(wb, filename);
}
function exportJSON(trades, filename) {
  downloadBlob(JSON.stringify(trades, null, 2), filename, "application/json");
}
function exportPDFReport(trades, from, to, journalName, profileName) {
  const m = computeMetrics(trades);
  const rows = buildExportRows(trades, []);
  const win = window.open("", "_blank");
  if (!win) return;
  const rowsHtml = rows.map(r => `<tr>
    <td>${r.Date}</td><td>${r.Symbol}</td><td style="text-transform:uppercase">${r.Side}</td>
    <td>${r.Lots}</td><td>${r.Entry}</td><td>${r.Exit}</td>
    <td>${r["R-Multiple"]}</td><td style="color:${r["P&L"] >= 0 ? "#0a7a4a" : "#b3261e"};font-weight:700">${r["P&L"]}</td>
    <td>${r.Setup}</td><td>${r.Session}</td>
  </tr>`).join("");
  win.document.write(`
    <html><head><title>Trading Report</title>
    <style>
      body { font-family: Arial, sans-serif; padding: 32px; color: #1a1a1a; }
      h1 { margin-bottom: 2px; } .sub { color: #666; margin-bottom: 20px; font-size: 13px; }
      .stats { display: flex; flex-wrap: wrap; gap: 14px; margin-bottom: 24px; }
      .stat { border: 1px solid #ddd; border-radius: 8px; padding: 10px 14px; min-width: 120px; }
      .stat .l { font-size: 10px; color: #777; text-transform: uppercase; font-weight: 700; }
      .stat .v { font-size: 18px; font-weight: 700; margin-top: 2px; }
      table { width: 100%; border-collapse: collapse; font-size: 11px; }
      th, td { border: 1px solid #ddd; padding: 5px 7px; text-align: left; }
      th { background: #f2f2f2; }
      @media print { button { display: none; } }
    </style></head><body>
    <h1>${profileName || "Trading"} — Performance Report</h1>
    <div class="sub">${journalName || "Journal"} · ${from || "All time"} to ${to || "present"} · Generated ${new Date().toLocaleDateString()}</div>
    <div class="stats">
      <div class="stat"><div class="l">Net P&L</div><div class="v">${fmt$(m.netProfit)}</div></div>
      <div class="stat"><div class="l">Trades</div><div class="v">${m.n}</div></div>
      <div class="stat"><div class="l">Win Rate</div><div class="v">${fmtPct(m.winRate)}</div></div>
      <div class="stat"><div class="l">Profit Factor</div><div class="v">${m.profitFactor === Infinity ? "∞" : m.profitFactor.toFixed(2)}</div></div>
      <div class="stat"><div class="l">Expectancy</div><div class="v">${fmt$(m.expectancy)}</div></div>
      <div class="stat"><div class="l">Max Drawdown</div><div class="v">${fmt$(-m.maxDD)}</div></div>
      <div class="stat"><div class="l">SQN</div><div class="v">${m.sqn.toFixed(2)}</div></div>
    </div>
    <table><thead><tr><th>Date</th><th>Symbol</th><th>Side</th><th>Lots</th><th>Entry</th><th>Exit</th><th>R</th><th>P&L</th><th>Setup</th><th>Session</th></tr></thead>
    <tbody>${rowsHtml}</tbody></table>
    <button onclick="window.print()" style="margin-top:20px;padding:10px 18px;background:#111;color:#fff;border:none;border-radius:8px;cursor:pointer;">Print / Save as PDF</button>
    </body></html>`);
  win.document.close();
}

function ExportPanel({ trades, customColumns, journalName, profileName }) {
  const dates = trades.map(t => t.entryTime).filter(Boolean).sort();
  const [from, setFrom] = useState(dates.length ? dates[0].slice(0, 10) : "");
  const [to, setTo] = useState(dates.length ? dates[dates.length - 1].slice(0, 10) : "");
  const [format, setFormat] = useState("csv");
  const filtered = useMemo(() => filterTradesByDateRange(trades, from, to), [trades, from, to]);

  const runExport = () => {
    if (!filtered.length) return;
    const stamp = `${from || "all"}_to_${to || "now"}`;
    if (format === "csv") exportCSV(buildExportRows(filtered, customColumns), `trades_${stamp}.csv`);
    else if (format === "xlsx") exportXLSX(buildExportRows(filtered, customColumns), `trades_${stamp}.xlsx`);
    else if (format === "json") exportJSON(filtered, `trades_${stamp}.json`);
    else if (format === "pdf") exportPDFReport(filtered, from, to, journalName, profileName);
  };

  const FORMATS = [
    { key: "csv", label: "CSV" },
    { key: "xlsx", label: "Excel (.xlsx)" },
    { key: "json", label: "JSON" },
    { key: "pdf", label: "PDF Report" },
  ];

  return (
    <div className="tj-panel tj-fade-in" style={{ padding: 20 }}>
      <SectionTitle icon={Download}>Export Journal Data</SectionTitle>
      <p className="tj-dim" style={{ fontSize: 12.5, marginBottom: 16 }}>
        Pick a date range and format — everything you've logged in this journal (including your custom columns) exports in one click.
      </p>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
        <div>
          <label style={{ fontSize: 11, fontWeight: 600, color: "#8A8D97", textTransform: "uppercase" }}>From</label>
          <input type="date" className="tj-input" style={{ marginTop: 4 }} value={from} onChange={e => setFrom(e.target.value)} />
        </div>
        <div>
          <label style={{ fontSize: 11, fontWeight: 600, color: "#8A8D97", textTransform: "uppercase" }}>To</label>
          <input type="date" className="tj-input" style={{ marginTop: 4 }} value={to} onChange={e => setTo(e.target.value)} />
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
        {FORMATS.map(f => (
          <button key={f.key} className="tj-btn" style={{
            padding: "8px 14px",
            background: format === f.key ? "rgba(227,178,77,0.14)" : "rgba(255,255,255,0.04)",
            color: format === f.key ? "#E3B24D" : "#F2F1ED",
            borderColor: format === f.key ? "rgba(227,178,77,0.3)" : "rgba(255,255,255,0.1)",
          }} onClick={() => setFormat(f.key)}>{f.label}</button>
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
        <div className="tj-dim" style={{ fontSize: 12 }}>{filtered.length} trade{filtered.length !== 1 ? "s" : ""} in this range</div>
        <button className="tj-btn tj-btn-primary" disabled={!filtered.length} style={{ opacity: filtered.length ? 1 : 0.5 }} onClick={runExport}>
          <Download size={14} /> Export {FORMATS.find(f => f.key === format)?.label}
        </button>
      </div>
    </div>
  );
}

/* ============================== MAIN APP ============================== */
const NAV = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "trades", label: "Trade Log", icon: ListChecks },
  { key: "calendar", label: "Calendar", icon: CalendarDays },
  { key: "psychology", label: "Psychology", icon: Brain },
  { key: "coach", label: "AI Coach", icon: MessageCircle },
  { key: "goals", label: "Goals", icon: Award },
];

function TradingJournalApp({ userEmail, onLogout }) {
  const [trades, setTrades] = useState([]);
  const [tab, setTab] = useState("dashboard");
  const [modalTrade, setModalTrade] = useState(undefined);
  const [showImport, setShowImport] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const [profileName, setProfileName] = useState("Aurlis");
  const [editingName, setEditingName] = useState(false);
  const [nameDraft, setNameDraft] = useState("Aurlis");
  const DEFAULT_TAGLINE = "Institutional-grade analytics for discretionary XAUUSD breakout trading";
  const [tagline, setTagline] = useState(DEFAULT_TAGLINE);
  const [editingTagline, setEditingTagline] = useState(false);
  const [taglineDraft, setTaglineDraft] = useState(DEFAULT_TAGLINE);
  const [customColumns, setCustomColumns] = useState(DEFAULT_CUSTOM_COLUMNS);
  const [selectedDay, setSelectedDay] = useState(null);
  const [primaryInstrument, setPrimaryInstrument] = useState("XAUUSD");
  const [editingInstrument, setEditingInstrument] = useState(false);
  const [instrumentDraft, setInstrumentDraft] = useState("XAUUSD");

  const [journals, setJournals] = useState([{ id: "default", name: "Main Journal" }]);
  const [activeJournalId, setActiveJournalId] = useState("default");
  const [journalMenuOpen, setJournalMenuOpen] = useState(false);
  const [newJournalName, setNewJournalName] = useState("");
  const [journalDataLoaded, setJournalDataLoaded] = useState(false);

  const loadJournalData = async (journalId, seedIfEmpty) => {
    setJournalDataLoaded(false);
    try {
      const res = await window.storage?.get(`trades:${journalId}`, false);
      if (res?.value) setTrades(JSON.parse(res.value));
      else setTrades(seedIfEmpty ? genSeedTrades() : []);
    } catch (e) {
      setTrades(seedIfEmpty ? genSeedTrades() : []);
    }
    try {
      const colRes = await window.storage?.get(`customColumns:${journalId}`, false);
      if (colRes?.value) setCustomColumns(JSON.parse(colRes.value));
      else setCustomColumns(DEFAULT_CUSTOM_COLUMNS);
    } catch (e) {
      setCustomColumns(DEFAULT_CUSTOM_COLUMNS);
    }
    try {
      const instRes = await window.storage?.get(`primaryInstrument:${journalId}`, false);
      const val = instRes?.value || "XAUUSD";
      setPrimaryInstrument(val);
      setInstrumentDraft(val);
    } catch (e) {
      setPrimaryInstrument("XAUUSD");
      setInstrumentDraft("XAUUSD");
    }
    setJournalDataLoaded(true);
  };

  useEffect(() => {
    (async () => {
      let journalList = [{ id: "default", name: "Main Journal" }];
      let activeId = "default";
      try {
        const jRes = await window.storage?.get("journals", false);
        if (jRes?.value) { journalList = JSON.parse(jRes.value); }
        else { try { await window.storage?.set("journals", JSON.stringify(journalList), false); } catch (e) {} }
      } catch (e) { /* keep default single journal */ }
      try {
        const aRes = await window.storage?.get("activeJournalId", false);
        if (aRes?.value) activeId = aRes.value;
      } catch (e) { /* keep default */ }
      setJournals(journalList);
      setActiveJournalId(activeId);

      try {
        const nameRes = await window.storage?.get("profileName", false);
        if (nameRes?.value) { setProfileName(nameRes.value); setNameDraft(nameRes.value); }
      } catch (e) { /* keep default */ }
      try {
        const taglineRes = await window.storage?.get("tagline", false);
        if (taglineRes?.value) { setTagline(taglineRes.value); setTaglineDraft(taglineRes.value); }
      } catch (e) { /* keep default */ }

      await loadJournalData(activeId, true);
      setLoaded(true);
    })();
  }, []);

  useEffect(() => { if (loaded && journalDataLoaded) saveTrades(trades, activeJournalId); }, [trades, loaded, journalDataLoaded, activeJournalId]);
  useEffect(() => {
    if (loaded && journalDataLoaded) { try { window.storage?.set(`customColumns:${activeJournalId}`, JSON.stringify(customColumns), false); } catch (e) {} }
  }, [customColumns, loaded, journalDataLoaded, activeJournalId]);

  const switchJournal = async (id) => {
    if (id === activeJournalId) { setJournalMenuOpen(false); return; }
    setActiveJournalId(id);
    setJournalMenuOpen(false);
    try { window.storage?.set("activeJournalId", id, false); } catch (e) {}
    await loadJournalData(id, false);
  };

  const createJournal = async () => {
    const name = newJournalName.trim();
    if (!name) return;
    const id = "j_" + Date.now().toString(36);
    const updated = [...journals, { id, name }];
    setJournals(updated);
    try { window.storage?.set("journals", JSON.stringify(updated), false); } catch (e) {}
    setNewJournalName("");
    await switchJournal(id);
  };

  const renameJournal = (id, name) => {
    setJournals(js => {
      const updated = js.map(j => (j.id === id ? { ...j, name } : j));
      try { window.storage?.set("journals", JSON.stringify(updated), false); } catch (e) {}
      return updated;
    });
  };

  const deleteJournal = (id) => {
    if (journals.length <= 1) return;
    const updated = journals.filter(j => j.id !== id);
    setJournals(updated);
    try { window.storage?.set("journals", JSON.stringify(updated), false); } catch (e) {}
    if (id === activeJournalId) switchJournal(updated[0].id);
  };

  const saveProfileName = () => {
    const val = nameDraft.trim() || "Aurlis";
    setProfileName(val);
    setEditingName(false);
    try { window.storage?.set("profileName", val, false); } catch (e) { /* noop */ }
  };

  const saveTagline = () => {
    const val = taglineDraft.trim() || DEFAULT_TAGLINE;
    setTagline(val);
    setEditingTagline(false);
    try { window.storage?.set("tagline", val, false); } catch (e) { /* noop */ }
  };

  const saveInstrument = () => {
    const val = (instrumentDraft.trim() || "XAUUSD").toUpperCase();
    setPrimaryInstrument(val);
    setEditingInstrument(false);
    try { window.storage?.set(`primaryInstrument:${activeJournalId}`, val, false); } catch (e) { /* noop */ }
  };

  const addCustomColumn = (label) => {
    const id = "cc_" + label.toLowerCase().replace(/[^a-z0-9]+/g, "_") + "_" + Date.now();
    setCustomColumns(cols => [...cols, { id, label }]);
  };
  const removeCustomColumn = (id) => setCustomColumns(cols => cols.filter(c => c.id !== id));

  const metrics = useMemo(() => computeMetrics(trades), [trades]);
  const psychology = useMemo(() => detectPsychology(trades), [trades]);

  const upsertTrade = (t) => {
    setTrades(prev => {
      const exists = prev.some(p => p.id === t.id);
      return exists ? prev.map(p => (p.id === t.id ? t : p)) : [...prev, t];
    });
    setModalTrade(undefined);
  };
  const deleteTrade = (id) => setTrades(prev => prev.filter(p => p.id !== id));
  const importTrades = (rows) => setTrades(prev => [...prev, ...rows]);
  const resetDemo = () => setTrades(genSeedTrades());


  return (
    <div className="tj-root">
      <GlobalStyle />
      <div className="tj-bg-glow" />
      <div style={{ position: "relative", zIndex: 1, display: "flex", minHeight: "100vh" }}>
        {/* Sidebar */}
        <div className="tj-sidebar" style={{ width: 220, flexShrink: 0, padding: "22px 14px", display: "flex", flexDirection: "column", gap: 4 }}>
          <div style={{ position: "relative", padding: "0 8px 10px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
              <div style={{ width: 30, height: 30, borderRadius: 8, background: "linear-gradient(135deg,#E3B24D,#8A6C2E)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <TrendingUp size={16} color="#0A0906" />
              </div>
              {editingName ? (
                <div style={{ display: "flex", alignItems: "center", gap: 4, minWidth: 0 }}>
                  <input autoFocus className="tj-input" value={nameDraft} onChange={e => setNameDraft(e.target.value)}
                    onKeyDown={e => e.key === "Enter" && saveProfileName()}
                    style={{ padding: "5px 8px", fontSize: 13.5, width: 100 }} />
                  <button className="tj-btn tj-btn-primary" style={{ padding: 6 }} onClick={saveProfileName}><Send size={12} /></button>
                </div>
              ) : (
                <div className="tj-display" style={{ fontSize: 15.5, fontWeight: 700, display: "flex", alignItems: "center", gap: 5, minWidth: 0, flex: 1 }}>
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{profileName}<span className="tj-gold">.</span></span>
                  <Edit3 size={12} className="tj-dim" style={{ cursor: "pointer", flexShrink: 0 }} onClick={() => { setNameDraft(profileName); setEditingName(true); }} />
                  <div style={{ flex: 1 }} />
                  <button onClick={() => setJournalMenuOpen(o => !o)} style={{ background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 7, padding: 4, cursor: "pointer", flexShrink: 0, display: "flex" }}>
                    {journalMenuOpen ? <ChevronUp size={13} color="#E3B24D" /> : <ChevronDown size={13} color="#8A8D97" />}
                  </button>
                </div>
              )}
            </div>
            {!editingName && (
              <div className="tj-dim" style={{ fontSize: 11, marginTop: 4, cursor: "pointer" }} onClick={() => setJournalMenuOpen(o => !o)}>
                📓 {journals.find(j => j.id === activeJournalId)?.name || "Main Journal"}
              </div>
            )}

            {journalMenuOpen && (
              <div className="tj-fade-in tj-scroll" style={{
                position: "absolute", top: "100%", left: 0, marginTop: 6, padding: 12, zIndex: 200,
                maxHeight: 320, overflowY: "auto", width: 240,
                background: "#111318", border: "1px solid rgba(255,255,255,0.14)", borderRadius: 14,
                boxShadow: "0 24px 60px rgba(0,0,0,0.65), 0 0 0 1px rgba(0,0,0,0.4)",
              }}>
                <div style={{ fontSize: 10.5, fontWeight: 700, color: "#8A8D97", textTransform: "uppercase", marginBottom: 8 }}>Your Journals</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
                  {journals.map(j => (
                    <div key={j.id} onClick={() => switchJournal(j.id)} style={{
                      display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6,
                      padding: "8px 10px", borderRadius: 8, cursor: "pointer", fontSize: 12.5,
                      background: j.id === activeJournalId ? "rgba(227,178,77,0.12)" : "rgba(255,255,255,0.03)",
                      border: j.id === activeJournalId ? "1px solid rgba(227,178,77,0.3)" : "1px solid transparent",
                      color: j.id === activeJournalId ? "#E3B24D" : "#F2F1ED",
                    }}>
                      <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: j.id === activeJournalId ? 700 : 500 }}>{j.name}</span>
                      {journals.length > 1 && (
                        <button onClick={(e) => { e.stopPropagation(); deleteJournal(j.id); }} style={{ background: "none", border: "none", color: "#6B6E78", cursor: "pointer", flexShrink: 0 }}>
                          <Trash2 size={12} />
                        </button>
                      )}
                    </div>
                  ))}
                </div>
                <div style={{ borderTop: "1px solid rgba(255,255,255,0.08)", paddingTop: 10 }}>
                  <div style={{ fontSize: 10.5, fontWeight: 700, color: "#8A8D97", textTransform: "uppercase", marginBottom: 6 }}>New Journal</div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <input className="tj-input" placeholder="e.g. Scalping Strategy" value={newJournalName}
                      onChange={e => setNewJournalName(e.target.value)} onKeyDown={e => e.key === "Enter" && createJournal()}
                      style={{ fontSize: 12.5 }} />
                    <button className="tj-btn tj-btn-primary" style={{ padding: "8px 10px" }} onClick={createJournal}><Plus size={13} /></button>
                  </div>
                  <div className="tj-dim" style={{ fontSize: 10.5, marginTop: 6 }}>Each journal keeps its own trades, columns, goals, and psychology rules — nothing mixes.</div>
                </div>
              </div>
            )}
          </div>
          {NAV.map(item => (
            <div key={item.key} className={`tj-navitem ${tab === item.key ? "active" : ""}`} onClick={() => setTab(item.key)}>
              <item.icon size={16} /> {item.label}
            </div>
          ))}
          <div style={{ flex: 1 }} />
          <div className="tj-navitem" onClick={resetDemo}><RefreshCw size={16} /> Reset Demo Data</div>
          <div className="tj-navitem" onClick={onLogout}><LogOut size={16} /> Log Out</div>
          {userEmail && (
            <div style={{ padding: "8px 8px 0", fontSize: 10.5, color: "#5C5F68", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              Signed in as {userEmail}
            </div>
          )}
          <div style={{ padding: "12px 8px 0", borderTop: "1px solid rgba(255,255,255,0.06)", marginTop: 8 }}>
            <div style={{ fontSize: 11, color: "#5C5F68" }}>Primary Instrument</div>
            {editingInstrument ? (
              <div style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 4 }}>
                <input autoFocus className="tj-input" value={instrumentDraft} onChange={e => setInstrumentDraft(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && saveInstrument()}
                  style={{ padding: "5px 8px", fontSize: 12.5, width: 90, textTransform: "uppercase" }} />
                <button className="tj-btn tj-btn-primary" style={{ padding: 6 }} onClick={saveInstrument}><Send size={11} /></button>
              </div>
            ) : (
              <div style={{ fontSize: 13, fontWeight: 700, color: "#E3B24D", display: "flex", alignItems: "center", gap: 6 }}>
                {primaryInstrument}
                <Edit3 size={11} className="tj-dim" style={{ cursor: "pointer" }} onClick={() => { setInstrumentDraft(primaryInstrument); setEditingInstrument(true); }} />
              </div>
            )}
          </div>
        </div>

        {/* Main */}
        <div style={{ flex: 1, minWidth: 0, padding: "22px 28px 60px" }}>
          {/* Topbar */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22, flexWrap: "wrap", gap: 12 }}>
            <div>
              <h1 className="tj-display" style={{ fontSize: 22, fontWeight: 700 }}>
                {NAV.find(n => n.key === tab)?.label}
              </h1>
              <div className="tj-dim" style={{ fontSize: 12.5, marginTop: 2, display: "flex", alignItems: "center", gap: 6 }}>
                {editingTagline ? (
                  <>
                    <input autoFocus className="tj-input" value={taglineDraft} onChange={e => setTaglineDraft(e.target.value)}
                      onKeyDown={e => e.key === "Enter" && saveTagline()}
                      style={{ padding: "5px 8px", fontSize: 12.5, width: 380 }} />
                    <button className="tj-btn tj-btn-primary" style={{ padding: 6 }} onClick={saveTagline}><Send size={12} /></button>
                  </>
                ) : (
                  <>
                    <span>{tagline}</span>
                    <Edit3 size={11} className="tj-dim" style={{ cursor: "pointer", flexShrink: 0 }} onClick={() => { setTaglineDraft(tagline); setEditingTagline(true); }} />
                  </>
                )}
              </div>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button className="tj-btn tj-btn-ghost" onClick={() => setShowImport(true)}><Upload size={14} /> Import</button>
              <button className="tj-btn tj-btn-primary" onClick={() => setModalTrade(emptyTrade())}><Plus size={14} /> Log Trade</button>
            </div>
          </div>

          {tab === "dashboard" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 12 }}>
                <StatCard label="Net P&L" value={fmt$(metrics.netProfit)} icon={TrendingUp} tone={metrics.netProfit >= 0 ? "good" : "bad"} />
                <StatCard label="Win Rate" value={fmtPct(metrics.winRate)} icon={Percent} tone={metrics.winRate >= 0.5 ? "good" : "bad"} />
                <StatCard label="Profit Factor" value={metrics.profitFactor === Infinity ? "∞" : metrics.profitFactor.toFixed(2)} icon={Target} tone="gold" />
                <StatCard label="Total Trades" value={metrics.n} icon={ListChecks} />
                <StatCard label="Current Streak" value={`${metrics.currentStreak.count}${metrics.currentStreak.sign > 0 ? "W" : metrics.currentStreak.sign < 0 ? "L" : ""}`} icon={Flame} tone={metrics.currentStreak.sign > 0 ? "good" : "bad"} />
              </div>
              <EquityHero metrics={metrics} />
              <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 18 }}>
                <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
                  <DrawdownChart metrics={metrics} />
                  <WeekdayChart trades={trades} />
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
                  <SessionChart trades={trades} />
                  <BuySellDonut metrics={metrics} />
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 }}>
                <RDistribution metrics={metrics} />
                <DailyPatterns trades={trades} />
              </div>
              <AdvancedMetrics m={metrics} />
            </div>
          )}

          {tab === "trades" && (
            <TradeTable trades={trades} onEdit={setModalTrade} onDelete={deleteTrade}
              customColumns={customColumns} onAddColumn={addCustomColumn} onRemoveColumn={removeCustomColumn} />
          )}

          {tab === "calendar" && (
            <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 18 }}>
              <TradeCalendar trades={trades} onDayClick={setSelectedDay} />
              <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
                <SessionChart trades={trades} />
                <DailyPatterns trades={trades} />
              </div>
            </div>
          )}

          {tab === "psychology" && <PsychologyPanel trades={trades} journalId={activeJournalId} key={`psych-${activeJournalId}`} />}

          {tab === "coach" && <AiCoach trades={trades} metrics={metrics} psychology={psychology} />}

          {tab === "goals" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
              <GoalsPanel metrics={metrics} journalId={activeJournalId} key={`goals-${activeJournalId}`} />
              <ExportPanel trades={trades} customColumns={customColumns} journalName={journals.find(j => j.id === activeJournalId)?.name} profileName={profileName} />
            </div>
          )}
        </div>
      </div>

      {modalTrade !== undefined && (
        <TradeModal trade={modalTrade.id ? modalTrade : null} onSave={upsertTrade} onClose={() => setModalTrade(undefined)} customColumns={customColumns} />
      )}
      {showImport && <ImportModal onImport={importTrades} onClose={() => setShowImport(false)} />}
      {selectedDay && (
        <DayDetailModal
          date={selectedDay}
          trades={trades}
          onClose={() => setSelectedDay(null)}
          onEditTrade={(t) => { setSelectedDay(null); setModalTrade(t); }}
          onDeleteTrade={deleteTrade}
        />
      )}
    </div>
  );
}

/* ============================== ROOT (AUTH-GATED) ============================== */
export default function TradingJournalRoot() {
  const [userEmail, setUserEmail] = useState(null);
  const [checked, setChecked] = useState(false);

  useEffect(() => { setChecked(true); }, []);

  const handleLogout = async () => {
    try { await window.storage?.set("authSession", "", false); } catch (e) {}
    setUserEmail(null);
  };

  if (!checked) return null;
  if (!userEmail) return <AuthGate onAuthenticated={setUserEmail} />;
  return <TradingJournalApp userEmail={userEmail} onLogout={handleLogout} />;
}
