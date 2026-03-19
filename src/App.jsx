import { useState, useEffect, useCallback } from 'react'
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Legend, ReferenceLine,
} from 'recharts'

// ── API設定 ──
const API_URL = 'https://script.google.com/macros/s/AKfycbx3fK6vqL6sO0O23cvNonlLflirOKmW4ZCTaKpFMUasoaxJt2P0LbL6XaAyNKi1/exec'

// ── カラーテーマ ──
const C = {
  accent: '#0ea5c7',
  accentLight: '#43d7ff',
  accentBg: 'rgba(67,215,255,0.06)',
  accentBorder: 'rgba(67,215,255,0.15)',
  bg: '#ffffff',
  pageBg: '#f4f7f9',
  card: '#f4f7f9',
  cardBorder: '#dce3ea',
  text: '#1a2229',
  textSub: '#4a5c66',
  textMuted: '#94a4ac',
  red: '#d64545',
  blue: '#3b82f6',
}

// ── ユーティリティ ──

/** 日付文字列やタイムスタンプを Date に変換 */
const parseDate = (raw) => {
  const d = new Date(raw)
  if (!isNaN(d.getTime())) return d
  return null
}

/** X軸用: MM/DD 形式 */
const formatDate = (raw) => {
  const d = parseDate(raw)
  if (d) return `${d.getMonth() + 1}/${d.getDate()}`
  const m = String(raw).match(/(\d{1,2})[\/\-](\d{1,2})$/)
  return m ? `${parseInt(m[1])}/${parseInt(m[2])}` : String(raw)
}

/** ツールチップ用: YYYY/MM/DD 形式 */
const formatDateLong = (raw) => {
  const d = parseDate(raw)
  if (d) return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`
  const m = String(raw).match(/(\d{2,4})[\/\-](\d{1,2})[\/\-](\d{1,2})/)
  if (m) return `${m[1].padStart(4, '20')}/${m[2].padStart(2, '0')}/${m[3].padStart(2, '0')}`
  return String(raw)
}

/** データ配列から年度(4月始まり)・四半期境界を抽出する */
const getFiscalBoundaries = (chartData) => {
  const results = []
  for (let i = 1; i < chartData.length; i++) {
    const prev = parseDate(chartData[i - 1].rawDate)
    const curr = parseDate(chartData[i].rawDate)
    if (!prev || !curr) continue
    if (prev.getMonth() === curr.getMonth()) continue
    const cm = curr.getMonth() + 1
    if (cm === 4 || cm === 7 || cm === 10 || cm === 1) {
      results.push({
        idx: i,
        type: cm === 4 ? 'year' : 'quarter',
        dateLabel: `${cm}/1`,
      })
    }
  }
  return results
}

/** 年度ラベルの表示位置（各年度区間の中央インデックス）を算出する */
const getFiscalYearLabels = (chartData, boundaries) => {
  const yearStarts = boundaries.filter(b => b.type === 'year').map(b => b.idx)
  const points = [0, ...yearStarts, chartData.length]
  const labels = []
  for (let i = 0; i < points.length - 1; i++) {
    const midIdx = Math.floor((points[i] + points[i + 1]) / 2)
    const d = parseDate(chartData[midIdx]?.rawDate)
    if (d) {
      const fy = d.getMonth() + 1 >= 4 ? d.getFullYear() : d.getFullYear() - 1
      labels.push({ idx: midIdx, label: `${fy}年度` })
    }
  }
  return labels
}

/** 境界線の下部に日付ラベルを表示するカスタムコンポーネント */
const BoundaryDateLabel = ({ viewBox, value }) => {
  if (!viewBox) return null
  return (
    <text x={viewBox.x} y={viewBox.y + viewBox.height + 14} textAnchor="middle"
      fontSize={9} fontWeight={600} fill={C.textMuted}>
      {value}
    </text>
  )
}

// ── 共通コンポーネント ──

const KpiCard = ({ label, value, sub, accent }) => (
  <div style={{
    flex: 1, minWidth: 140, padding: '16px 14px', textAlign: 'center',
    background: accent ? `linear-gradient(135deg, ${C.accent}, ${C.accentLight})` : C.card,
    border: `1px solid ${accent ? 'transparent' : C.cardBorder}`,
    borderRadius: 12,
  }}>
    <div style={{ fontSize: 11, color: accent ? 'rgba(255,255,255,0.8)' : C.textMuted, fontWeight: 500, marginBottom: 3 }}>{label}</div>
    <div style={{ fontSize: 24, fontWeight: 800, color: accent ? '#fff' : C.accent, letterSpacing: '-0.5px', lineHeight: 1.2 }}>{value}</div>
    {sub && <div style={{ fontSize: 10, color: accent ? 'rgba(255,255,255,0.7)' : C.textMuted, marginTop: 3 }}>{sub}</div>}
  </div>
)

const SectionTitle = ({ children }) => (
  <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 10, marginTop: 20, display: 'flex', alignItems: 'center', gap: 8 }}>
    <div style={{ width: 3, height: 18, background: C.accent, borderRadius: 2 }} />
    {children}
  </div>
)

const ChartTooltip = ({ active, payload, label }) => {
  if (!active || !payload || payload.length === 0) return null
  const displayDate = payload[0]?.payload?.tooltipDate || label
  return (
    <div style={{ background: '#fff', border: `1px solid ${C.cardBorder}`, borderRadius: 8, padding: '8px 12px', fontSize: 11, boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
      <div style={{ fontWeight: 700, marginBottom: 4 }}>{displayDate}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color }}>{p.name}: {Number(p.value).toLocaleString()}</div>
      ))}
    </div>
  )
}

/** ソート可能なカラムか判定（数値 or 日付文字列） */
const isSortableColumn = (rows, colIndex) => {
  for (const row of rows) {
    const cell = row[colIndex]
    if (cell == null || cell === '' || cell === 0) continue
    if (typeof cell === 'number') return true
    if (typeof cell === 'string' && /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(cell)) return true
    return false
  }
  return false
}

/** セル値の比較（数値 or 日付文字列） */
const compareCells = (a, b) => {
  if (typeof a === 'number' && typeof b === 'number') return a - b
  return String(a || '').localeCompare(String(b || ''), 'ja')
}

/** ソートインジケーター */
const SortIndicator = ({ direction }) => (
  <span style={{ marginLeft: 3, fontSize: 10, color: C.accent }}>
    {direction === 'asc' ? '▲' : '▼'}
  </span>
)

const useSortableTable = (rows, defaultSortCol = 0, defaultDir = 'desc') => {
  const [sortCol, setSortCol] = useState(defaultSortCol)
  const [sortDir, setSortDir] = useState(defaultDir)

  const handleSort = (colIndex) => {
    if (sortCol === colIndex) {
      setSortDir(prev => prev === 'asc' ? 'desc' : 'asc')
    } else {
      setSortCol(colIndex)
      setSortDir('desc')
    }
  }

  const sortedRows = [...rows].sort((a, b) => {
    const result = compareCells(a[sortCol], b[sortCol])
    return sortDir === 'asc' ? result : -result
  })

  return { sortedRows, sortCol, sortDir, handleSort }
}

const DataTable = ({ headers, rows, maxRows = 10 }) => {
  const { sortedRows, sortCol, sortDir, handleSort } = useSortableTable(rows)
  const sortable = headers.map((_, i) => isSortableColumn(rows, i))

  return (
    <div style={{ overflowX: 'auto' }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
        <thead>
          <tr>
            {headers.map((h, i) => (
              <th key={i} onClick={() => sortable[i] && handleSort(i)} style={{
                textAlign: i === 0 || i === 1 ? 'left' : 'right',
                padding: '8px 10px', color: sortCol === i ? C.accent : C.textMuted, fontWeight: 600,
                borderBottom: `2px solid ${C.cardBorder}`, whiteSpace: 'nowrap',
                cursor: sortable[i] ? 'pointer' : 'default',
                userSelect: 'none',
              }}>
                {h}{sortCol === i && <SortIndicator direction={sortDir} />}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sortedRows.slice(0, maxRows).map((row, i) => (
            <tr key={i} style={{ borderBottom: `1px solid ${C.cardBorder}` }}>
              {row.map((cell, j) => (
                <td key={j} style={{
                  textAlign: j === 0 || j === 1 ? 'left' : 'right',
                  padding: '8px 10px', color: j <= 1 ? C.text : C.textSub,
                  fontWeight: j <= 1 ? 600 : 400,
                  maxWidth: j === 1 ? 200 : 'none',
                  overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                }}>{typeof cell === 'number' ? cell.toLocaleString() : cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

const Loading = () => (
  <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: 300, color: C.textMuted }}>
    <div>データを読み込み中...</div>
  </div>
)

const ErrorMsg = ({ message, onRetry }) => (
  <div style={{ textAlign: 'center', padding: 40, color: C.textSub }}>
    <div style={{ fontSize: 14, marginBottom: 12 }}>{message}</div>
    {onRetry && (
      <button onClick={onRetry} style={{
        padding: '8px 20px', border: `1px solid ${C.accent}`, borderRadius: 8,
        background: 'transparent', color: C.accent, fontSize: 13, fontWeight: 600, cursor: 'pointer',
      }}>再読み込み</button>
    )}
  </div>
)


// ── アカウント概要タブ ──

const PERIOD_OPTIONS = [
  { key: 7, label: '7日' },
  { key: 14, label: '14日' },
  { key: 30, label: '30日' },
  { key: 0, label: '全期間' },
]

function AccountTab({ data }) {
  const [period, setPeriod] = useState(0)

  if (!data || data.length === 0) return <ErrorMsg message="アカウントデータがまだありません" />

  const filtered = period > 0 ? data.slice(-period) : data

  const latest = filtered[filtered.length - 1]
  const first = filtered[0]
  const startFollowers = (first['フォロワー数'] || 0) - (first['フォロワー増減'] || 0)
  const endFollowers = latest['フォロワー数'] || 0
  const netGain = endFollowers - startFollowers
  const dailyAvg = filtered.length > 0 ? (netGain / filtered.length).toFixed(1) : 0

  // フォロワー推移チャート用データ
  const followerChart = filtered.map((d, i) => ({
    idx: i,
    date: formatDate(d['日付']),
    rawDate: d['日付'],
    tooltipDate: formatDateLong(d['日付']),
    followers: d['フォロワー数'],
    delta: d['フォロワー増減'],
  }))

  // 日次指標チャート
  const dailyChart = filtered.map((d, i) => ({
    idx: i,
    date: formatDate(d['日付']),
    rawDate: d['日付'],
    tooltipDate: formatDateLong(d['日付']),
    views: d['閲覧数'] || 0,
    interactions: d['インタラクション数'] || 0,
  }))

  const followerBoundaries = getFiscalBoundaries(followerChart)
  const followerFYLabels = getFiscalYearLabels(followerChart, followerBoundaries)
  const dailyBoundaries = getFiscalBoundaries(dailyChart)
  const dailyFYLabels = getFiscalYearLabels(dailyChart, dailyBoundaries)

  return (
    <div>
      {/* 期間セレクター */}
      <div style={{ display: 'flex', gap: 4, marginBottom: 14, background: C.bg, padding: 3, borderRadius: 8, border: `1px solid ${C.cardBorder}`, width: 'fit-content' }}>
        {PERIOD_OPTIONS.map(o => (
          <button key={o.key} onClick={() => setPeriod(o.key)} style={{
            padding: '5px 14px', border: 'none', borderRadius: 6, cursor: 'pointer',
            fontSize: 12, fontWeight: 600, transition: 'all 0.2s',
            background: period === o.key ? C.accent : 'transparent',
            color: period === o.key ? '#fff' : C.textMuted,
          }}>{o.label}</button>
        ))}
      </div>

      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
        <KpiCard label="現在のフォロワー" value={endFollowers.toLocaleString()} sub={`開始時 ${startFollowers.toLocaleString()}`} accent />
        <KpiCard label="フォロワー純増" value={`+${netGain.toLocaleString()}`} sub={`${filtered.length}日間`} />
        <KpiCard label="日平均純増" value={`${dailyAvg}人`} />
        <KpiCard label="昨日の閲覧数" value={(latest['閲覧数'] || 0).toLocaleString()} />
        <KpiCard label="昨日のインタラクション" value={(latest['インタラクション数'] || 0).toLocaleString()} />
      </div>

      <SectionTitle>フォロワー推移</SectionTitle>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.cardBorder}`, padding: '16px 12px 8px' }}>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={followerChart} margin={{ top: 20 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.cardBorder} vertical={followerBoundaries.length === 0} />
            <XAxis dataKey="idx" type="number" domain={[0, followerChart.length - 1]}
              tickFormatter={i => followerChart[Math.round(i)]?.date || ''}
              tick={{ fontSize: 10, fill: C.textMuted }} tickLine={false} />
            <YAxis tick={{ fontSize: 10, fill: C.textMuted }} tickLine={false} axisLine={false}
              domain={['dataMin - 100', 'dataMax + 100']}
              tickFormatter={v => `${(v / 1000).toFixed(1)}K`} />
            <Tooltip content={<ChartTooltip />} />
            {followerFYLabels.map(fl => (
              <ReferenceLine key={fl.label} x={fl.idx} stroke="none"
                label={{ value: fl.label, position: 'top', fontSize: 10, fontWeight: 700, fill: C.textSub }} />
            ))}
            {followerBoundaries.map((b, i) => (
              <ReferenceLine key={`fb-${i}`} x={b.idx}
                stroke={b.type === 'year' ? C.textMuted : C.cardBorder}
                strokeWidth={b.type === 'year' ? 1.5 : 1}
                strokeDasharray={b.type === 'year' ? '' : '3 3'}
                label={<BoundaryDateLabel value={b.dateLabel} />} />
            ))}
            <Line type="monotone" dataKey="followers" name="フォロワー数" stroke={C.accent} strokeWidth={2.5} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <SectionTitle>フォロワー増減（日次）</SectionTitle>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.cardBorder}`, padding: '16px 12px 8px' }}>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={followerChart} margin={{ top: 20 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.cardBorder} vertical={false} />
            <XAxis dataKey="idx" type="number" domain={[0, followerChart.length - 1]}
              tickFormatter={i => followerChart[Math.round(i)]?.date || ''}
              tick={{ fontSize: 10, fill: C.textMuted }} tickLine={false} />
            <YAxis tick={{ fontSize: 10, fill: C.textMuted }} tickLine={false} axisLine={false} />
            <Tooltip content={<ChartTooltip />} />
            {followerFYLabels.map(fl => (
              <ReferenceLine key={fl.label} x={fl.idx} stroke="none"
                label={{ value: fl.label, position: 'top', fontSize: 10, fontWeight: 700, fill: C.textSub }} />
            ))}
            {followerBoundaries.map((b, i) => (
              <ReferenceLine key={`fb2-${i}`} x={b.idx}
                stroke={b.type === 'year' ? C.textMuted : C.cardBorder}
                strokeWidth={b.type === 'year' ? 1.5 : 1}
                strokeDasharray={b.type === 'year' ? '' : '3 3'}
                label={<BoundaryDateLabel value={b.dateLabel} />} />
            ))}
            <Bar dataKey="delta" name="増減" radius={[3, 3, 0, 0]}>
              {followerChart.map((d, i) => (
                <rect key={i} fill={d.delta >= 0 ? C.accentLight : C.red} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <SectionTitle>閲覧数・インタラクション（日次）</SectionTitle>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.cardBorder}`, padding: '16px 12px 8px' }}>
        <ResponsiveContainer width="100%" height={200}>
          <LineChart data={dailyChart} margin={{ top: 20 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.cardBorder} vertical={dailyBoundaries.length === 0} />
            <XAxis dataKey="idx" type="number" domain={[0, dailyChart.length - 1]}
              tickFormatter={i => dailyChart[Math.round(i)]?.date || ''}
              tick={{ fontSize: 10, fill: C.textMuted }} tickLine={false} />
            <YAxis tick={{ fontSize: 10, fill: C.textMuted }} tickLine={false} axisLine={false} />
            <Tooltip content={<ChartTooltip />} />
            <Legend wrapperStyle={{ fontSize: 11 }} />
            {dailyFYLabels.map(fl => (
              <ReferenceLine key={fl.label} x={fl.idx} stroke="none"
                label={{ value: fl.label, position: 'top', fontSize: 10, fontWeight: 700, fill: C.textSub }} />
            ))}
            {dailyBoundaries.map((b, i) => (
              <ReferenceLine key={`db-${i}`} x={b.idx}
                stroke={b.type === 'year' ? C.textMuted : C.cardBorder}
                strokeWidth={b.type === 'year' ? 1.5 : 1}
                strokeDasharray={b.type === 'year' ? '' : '3 3'}
                label={<BoundaryDateLabel value={b.dateLabel} />} />
            ))}
            <Line type="monotone" dataKey="views" name="閲覧数" stroke={C.blue} strokeWidth={2} dot={false} />
            <Line type="monotone" dataKey="interactions" name="インタラクション" stroke={C.accent} strokeWidth={2} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      {/* フォロワー属性（常に最新データを使用） */}
      {data[data.length - 1]['フォロワー都市TOP5'] && (
        <>
          <SectionTitle>フォロワー属性</SectionTitle>
          <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.cardBorder}`, padding: 16, fontSize: 12, color: C.textSub, lineHeight: 1.8 }}>
            <div><strong style={{ color: C.text }}>都市TOP5:</strong> {data[data.length - 1]['フォロワー都市TOP5']}</div>
            <div><strong style={{ color: C.text }}>国TOP5:</strong> {data[data.length - 1]['フォロワー国TOP5']}</div>
            <div><strong style={{ color: C.text }}>性別×年齢:</strong> {data[data.length - 1]['フォロワー性別年齢']}</div>
          </div>
        </>
      )}
    </div>
  )
}


// ── タイトル編集コンポーネント ──

const EditableTitle = ({ content, title, onSave }) => {
  const [editing, setEditing] = useState(false)
  const [value, setValue] = useState(title || '')

  const handleSave = () => {
    onSave(content, value)
    setEditing(false)
  }

  const handleKeyDown = (e) => {
    if (e.key === 'Enter') handleSave()
    if (e.key === 'Escape') { setValue(title || ''); setEditing(false) }
  }

  if (editing) {
    return (
      <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
        <input
          value={value} onChange={e => setValue(e.target.value)}
          onKeyDown={handleKeyDown} autoFocus
          style={{
            flex: 1, padding: '3px 6px', border: `1px solid ${C.accent}`, borderRadius: 4,
            fontSize: 12, fontFamily: 'inherit', outline: 'none', minWidth: 120,
          }}
          placeholder="タイトルを入力"
        />
        <button onClick={handleSave} style={{
          padding: '2px 8px', border: 'none', borderRadius: 4, background: C.accent,
          color: '#fff', fontSize: 10, cursor: 'pointer', fontWeight: 600, whiteSpace: 'nowrap',
        }}>保存</button>
        <button onClick={() => { setValue(title || ''); setEditing(false) }} style={{
          padding: '2px 8px', border: `1px solid ${C.cardBorder}`, borderRadius: 4,
          background: 'transparent', color: C.textMuted, fontSize: 10, cursor: 'pointer', whiteSpace: 'nowrap',
        }}>✕</button>
      </div>
    )
  }

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer' }} onClick={() => setEditing(true)}>
      <span style={{
        maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
        color: title ? C.text : C.textMuted,
      }}>
        {title || content || '—'}
      </span>
      <span style={{ fontSize: 10, color: C.textMuted, flexShrink: 0 }}>✏️</span>
    </div>
  )
}

// ── タイトル付きデータテーブル ──

const TitledDataTable = ({ headers, rows, titleMap, onSaveTitle, titleColIndex = 1, maxRows = 25 }) => {
  const { sortedRows, sortCol, sortDir, handleSort } = useSortableTable(rows)
  const sortable = headers.map((_, i) => i !== titleColIndex && isSortableColumn(rows, i))

  return (
    <div style={{ overflowX: 'auto' }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
        <thead>
          <tr>
            {headers.map((h, i) => (
              <th key={i} onClick={() => sortable[i] && handleSort(i)} style={{
                textAlign: i === 0 || i === titleColIndex ? 'left' : 'right',
                padding: '8px 10px', color: sortCol === i ? C.accent : C.textMuted, fontWeight: 600,
                borderBottom: `2px solid ${C.cardBorder}`, whiteSpace: 'nowrap',
                cursor: sortable[i] ? 'pointer' : 'default',
                userSelect: 'none',
              }}>
                {h}{sortCol === i && <SortIndicator direction={sortDir} />}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sortedRows.slice(0, maxRows).map((row, i) => (
            <tr key={i} style={{ borderBottom: `1px solid ${C.cardBorder}` }}>
              {row.map((cell, j) => (
                <td key={j} style={{
                  textAlign: j === 0 || j === titleColIndex ? 'left' : 'right',
                  padding: '8px 10px', color: j <= titleColIndex ? C.text : C.textSub,
                  fontWeight: j <= titleColIndex ? 600 : 400,
                  maxWidth: j === titleColIndex ? 220 : 'none',
                  overflow: j === titleColIndex ? 'visible' : 'hidden',
                  textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                }}>
                  {j === titleColIndex ? (
                    <EditableTitle
                      content={cell.content}
                      title={titleMap[cell.content]}
                      onSave={onSaveTitle}
                    />
                  ) : (
                    typeof cell === 'number' ? cell.toLocaleString() : cell
                  )}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}


// ── 通常投稿タブ ──

function FeedTab({ data, titleMap, onSaveTitle }) {
  if (!data || data.length === 0) return <ErrorMsg message="通常投稿のデータがまだありません" />

  const totalReach = data.reduce((s, d) => s + (d['リーチ'] || 0), 0)
  const totalViews = data.reduce((s, d) => s + (d['閲覧数'] || 0), 0)
  const totalLikes = data.reduce((s, d) => s + (d['いいね'] || 0), 0)
  const totalSaves = data.reduce((s, d) => s + (d['保存数'] || 0), 0)
  const totalFollows = data.reduce((s, d) => s + (d['フォロー数'] || 0), 0)

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
        <KpiCard label="投稿数" value={data.length} accent />
        <KpiCard label="閲覧数合計" value={totalViews.toLocaleString()} />
        <KpiCard label="リーチ合計" value={totalReach.toLocaleString()} />
        <KpiCard label="いいね合計" value={totalLikes.toLocaleString()} />
        <KpiCard label="保存数合計" value={totalSaves.toLocaleString()} />
        <KpiCard label="フォロー獲得" value={totalFollows.toLocaleString()} />
      </div>

      <SectionTitle>投稿一覧</SectionTitle>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.cardBorder}`, padding: 16 }}>
        <TitledDataTable
          headers={['投稿日', 'タイトル', 'タイプ', '閲覧数', 'リーチ', 'いいね', '保存', 'シェア', 'フォロー']}
          rows={data.map(d => [
            formatDateLong(d['投稿日']), { content: d['内容'] }, d['メディアタイプ'],
            d['閲覧数'] || 0, d['リーチ'] || 0, d['いいね'] || 0,
            d['保存数'] || 0, d['シェア'] || 0, d['フォロー数'] || 0,
          ])}
          titleMap={titleMap}
          onSaveTitle={onSaveTitle}
          maxRows={25}
        />
      </div>
    </div>
  )
}


// ── リールタブ ──

function ReelsTab({ data, titleMap, onSaveTitle }) {
  if (!data || data.length === 0) return <ErrorMsg message="リールのデータがまだありません" />

  const totalViews = data.reduce((s, d) => s + (d['閲覧数'] || 0), 0)
  const totalReach = data.reduce((s, d) => s + (d['リーチ'] || 0), 0)
  const totalSaves = data.reduce((s, d) => s + (d['保存数'] || 0), 0)
  const totalShares = data.reduce((s, d) => s + (d['シェア'] || 0), 0)

  const avgWatchTimes = data.filter(d => d['平均視聴時間(秒)'] > 0)
  const overallAvgWatch = avgWatchTimes.length > 0
    ? (avgWatchTimes.reduce((s, d) => s + d['平均視聴時間(秒)'], 0) / avgWatchTimes.length / 1000).toFixed(2)
    : '—'

  const sorted = [...data].sort((a, b) => (b['閲覧数'] || 0) - (a['閲覧数'] || 0))

  // リーチ棒グラフ
  const chartData = sorted.slice(0, 10).map(d => ({
    name: (titleMap[d['内容']] || (d['内容'] || '').substring(0, 15)),
    views: d['閲覧数'] || 0,
    reach: d['リーチ'] || 0,
  }))

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
        <KpiCard label="リール数" value={data.length} accent />
        <KpiCard label="閲覧数合計" value={totalViews.toLocaleString()} />
        <KpiCard label="リーチ合計" value={totalReach.toLocaleString()} />
        <KpiCard label="保存数合計" value={totalSaves.toLocaleString()} />
        <KpiCard label="シェア合計" value={totalShares.toLocaleString()} />
        <KpiCard label="平均視聴時間" value={`${overallAvgWatch}秒`} />
      </div>

      <SectionTitle>閲覧数・リーチ TOP10</SectionTitle>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.cardBorder}`, padding: '16px 12px 8px' }}>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={chartData} layout="vertical" margin={{ left: 80 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.cardBorder} horizontal={false} />
            <XAxis type="number" tick={{ fontSize: 10, fill: C.textMuted }} tickFormatter={v => `${(v / 1000).toFixed(0)}K`} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 10, fill: C.textSub }} width={80} />
            <Tooltip content={<ChartTooltip />} />
            <Legend wrapperStyle={{ fontSize: 11 }} />
            <Bar dataKey="views" name="閲覧数" fill={C.accentLight} radius={[0, 4, 4, 0]} barSize={14} />
            <Bar dataKey="reach" name="リーチ" fill={C.blue} radius={[0, 4, 4, 0]} barSize={14} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <SectionTitle>リール一覧</SectionTitle>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.cardBorder}`, padding: 16 }}>
        <TitledDataTable
          headers={['投稿日', 'タイトル', '閲覧数', 'リーチ', 'いいね', '保存', 'シェア', '平均視聴(秒)']}
          rows={data.map(d => [
            formatDateLong(d['投稿日']), { content: d['内容'] },
            d['閲覧数'] || 0, d['リーチ'] || 0, d['いいね'] || 0,
            d['保存数'] || 0, d['シェア'] || 0, +((d['平均視聴時間(秒)'] || 0) / 1000).toFixed(2),
          ])}
          titleMap={titleMap}
          onSaveTitle={onSaveTitle}
          maxRows={25}
        />
      </div>
    </div>
  )
}


// ── ストーリーズタブ ──

function StoriesTab({ data }) {
  if (!data || data.length === 0) return <ErrorMsg message="ストーリーズのデータがまだありません（24時間以内に公開されたものが対象）" />

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
        <KpiCard label="ストーリー数" value={data.length} accent />
        <KpiCard label="閲覧数合計" value={data.reduce((s, d) => s + (d['閲覧数'] || 0), 0).toLocaleString()} />
        <KpiCard label="リーチ合計" value={data.reduce((s, d) => s + (d['リーチ'] || 0), 0).toLocaleString()} />
      </div>

      <SectionTitle>ストーリーズ一覧</SectionTitle>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.cardBorder}`, padding: 16 }}>
        <DataTable
          headers={['投稿日', '内容', '閲覧数', 'リーチ', 'シェア', 'フォロー', 'ナビゲーション']}
          rows={data.map(d => [
            formatDateLong(d['投稿日']), d['内容'],
            d['閲覧数'] || 0, d['リーチ'] || 0, d['シェア'] || 0,
            d['フォロー数'] || 0, d['ナビゲーション'] || 0,
          ])}
        />
      </div>
    </div>
  )
}


// ── メインApp ──

export default function App() {
  const [tab, setTab] = useState('account')
  const [data, setData] = useState({})
  const [titleMap, setTitleMap] = useState({})
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)
  const [lastUpdated, setLastUpdated] = useState(null)

  const fetchData = useCallback(async () => {
    setLoading(true)
    setError(null)
    try {
      const res = await fetch(`${API_URL}?type=all`)
      const json = await res.json()
      setData({
        account: json.account?.data || [],
        feed: json.feed?.data || [],
        reels: json.reels?.data || [],
        stories: json.stories?.data || [],
      })
      setTitleMap(json.titles?.data || {})
      setLastUpdated(new Date())
    } catch (e) {
      setError(`データの取得に失敗しました: ${e.message}`)
    } finally {
      setLoading(false)
    }
  }, [])

  const handleSaveTitle = useCallback(async (contentKey, title) => {
    setTitleMap(prev => ({ ...prev, [contentKey]: title }))
    try {
      const params = new URLSearchParams({ action: 'setTitle', key: contentKey, title })
      await fetch(`${API_URL}?${params}`)
    } catch { /* 保存失敗時もローカル状態は維持 */ }
  }, [])

  useEffect(() => { fetchData() }, [fetchData])

  const tabs = [
    { key: 'account', label: 'アカウント概要', icon: '📊' },
    { key: 'feed', label: '通常投稿', icon: '🏔' },
    { key: 'reels', label: 'リール', icon: '🎬' },
    { key: 'stories', label: 'ストーリーズ', icon: '📱' },
  ]

  return (
    <div style={{ maxWidth: 1100, margin: '0 auto', padding: '0 16px 40px' }}>
      {/* ヘッダー */}
      <header style={{ padding: '24px 0 16px', borderBottom: `3px solid ${C.accent}`, marginBottom: 20 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end' }}>
          <div>
            <div style={{ fontSize: 11, color: C.accentLight, letterSpacing: 2, fontWeight: 600, marginBottom: 2 }}>
              WEBTV INSIGHTS DASHBOARD
            </div>
            <h1 style={{ fontSize: 22, fontWeight: 800, color: C.text }}>
              webtv インサイトダッシュボード
            </h1>
          </div>
          <div style={{ textAlign: 'right', fontSize: 11, color: C.textMuted }}>
            {lastUpdated && <div>最終取得: {lastUpdated.toLocaleString('ja-JP')}</div>}
            <button onClick={fetchData} style={{
              marginTop: 4, padding: '4px 12px', border: `1px solid ${C.cardBorder}`,
              borderRadius: 6, background: C.bg, fontSize: 11, cursor: 'pointer', color: C.textSub,
            }}>↻ 更新</button>
          </div>
        </div>
      </header>

      {/* タブ */}
      <nav style={{ display: 'flex', gap: 4, marginBottom: 20, background: C.bg, padding: 4, borderRadius: 10, border: `1px solid ${C.cardBorder}` }}>
        {tabs.map(t => (
          <button key={t.key} onClick={() => setTab(t.key)} style={{
            flex: 1, padding: '10px 16px', border: 'none', borderRadius: 7, cursor: 'pointer',
            fontSize: 13, fontWeight: 600, transition: 'all 0.2s',
            background: tab === t.key ? C.accent : 'transparent',
            color: tab === t.key ? '#fff' : C.textMuted,
          }}>
            {t.icon} {t.label}
            {data[t.key] && <span style={{
              marginLeft: 6, fontSize: 10, opacity: 0.8,
              background: tab === t.key ? 'rgba(255,255,255,0.2)' : C.accentBg,
              padding: '1px 6px', borderRadius: 10,
              color: tab === t.key ? '#fff' : C.accent,
            }}>{data[t.key].length}</span>}
          </button>
        ))}
      </nav>

      {/* コンテンツ */}
      {loading ? <Loading /> : error ? <ErrorMsg message={error} onRetry={fetchData} /> : (
        <>
          {tab === 'account' && <AccountTab data={data.account} />}
          {tab === 'feed' && <FeedTab data={data.feed} titleMap={titleMap} onSaveTitle={handleSaveTitle} />}
          {tab === 'reels' && <ReelsTab data={data.reels} titleMap={titleMap} onSaveTitle={handleSaveTitle} />}
          {tab === 'stories' && <StoriesTab data={data.stories} />}
        </>
      )}

      {/* フッター */}
      <footer style={{ marginTop: 40, paddingTop: 16, borderTop: `1px solid ${C.cardBorder}`, fontSize: 10, color: C.textMuted, textAlign: 'center' }}>
        webtv インサイトダッシュボード ｜ データはInstagram Graph API経由で自動取得
      </footer>
    </div>
  )
}
