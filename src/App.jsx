
import React, { useEffect, useMemo, useRef, useState } from 'react'
import {
  Home,
  FolderKanban,
  Settings,
  PlusCircle,
  Search,
  Download,
  FileText,
  Eye,
  Trash2,
  Pencil,
  Save,
  ArrowLeft,
  ShieldCheck,
  Upload,
  RotateCcw,
  AlertTriangle,
  ReceiptText,
  BarChart3
} from 'lucide-react'
import * as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import { isSupabaseConfigured, supabase } from './supabaseClient'

const STORAGE_KEY = 'eventExpenseTracker.activities.v1'
const SETTINGS_KEY = 'eventExpenseTracker.settings.v1'
const CLOUD_TABLE = 'user_app_data'

const defaultSettings = {
  systemName: '活動支出追蹤系統',
  passwordHash: '',
  activityCategories: ['社區活動', '小組活動', '學校活動', '宣傳活動', '培訓活動'],
  expenseCategories: ['物資', '交通', '膳食', '場地', '印刷', '其他'],
  sortMode: 'created_desc'
}

const emptyActivity = {
  activityName: '',
  activityCode: '',
  startDate: '',
  endDate: '',
  budget: '',
  category: '社區活動',
  accountingCode: '',
  expenseCategory: '物資',
  personInCharge: '',
  advanceApplied: '沒有',
  advanceReceipts: []
}

const emptyExpense = {
  receiptNo: '',
  date: '',
  category: '物資',
  description: '',
  amount: '',
  payer: ''
}

function uid() {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function currency(value) {
  const num = Number(value || 0)
  return new Intl.NumberFormat('zh-HK', { style: 'currency', currency: 'HKD' }).format(num)
}

function today() {
  return new Date().toISOString().slice(0, 10)
}

async function sha256(text) {
  const enc = new TextEncoder()
  const buffer = await crypto.subtle.digest('SHA-256', enc.encode(text))
  return Array.from(new Uint8Array(buffer)).map(b => b.toString(16).padStart(2, '0')).join('')
}

function downloadText(filename, text, type = 'application/json') {
  const blob = new Blob([text], { type })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

function safeFileName(text) {
  return String(text || 'receipt')
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_') || 'receipt'
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result)
    reader.onerror = reject
    reader.readAsDataURL(file)
  })
}

function downloadDataUrl(filename, dataUrl) {
  const a = document.createElement('a')
  a.href = dataUrl
  a.download = filename
  a.click()
}

function openPdfPreview(dataUrl) {
  const win = window.open()
  if (!win) {
    alert('瀏覽器阻擋了預覽視窗，請允許彈出視窗後再試。')
    return
  }
  win.document.write(`
    <!doctype html>
    <html>
      <head>
        <title>PDF 預覽</title>
        <style>
          html, body { margin: 0; height: 100%; background: #0f172a; }
          iframe { width: 100%; height: 100%; border: 0; background: white; }
        </style>
      </head>
      <body>
        <iframe src="${dataUrl}"></iframe>
      </body>
    </html>
  `)
  win.document.close()
}

function openImagePreview(dataUrl, title = '圖片預覽') {
  const win = window.open()
  if (!win) {
    alert('瀏覽器阻擋了預覽視窗，請允許彈出視窗後再試。')
    return
  }
  win.document.write(`
    <!doctype html>
    <html>
      <head>
        <title>${title}</title>
        <style>
          html, body { margin: 0; min-height: 100%; background: #0f172a; display: grid; place-items: center; }
          img { max-width: 100%; max-height: 100vh; object-fit: contain; background: white; }
        </style>
      </head>
      <body>
        <img src="${dataUrl}" alt="${title}" />
      </body>
    </html>
  `)
  win.document.close()
}

function parseQuotedLine(line) {
  const matches = line.match(/"[^"]*"|\S+/g) || []
  return matches.map(item => item.replace(/^"|"$/g, '').trim())
}

function parseExpenseLine(line) {
  const parts = parseQuotedLine(line)
  if (parts.length < 6) return { error: '資料不足，請使用：收據編號 日期 類別 支出描述 金額 支付者' }

  const receiptNo = parts[0]
  const date = parts[1]
  const category = parts[2]
  const amountIndex = parts.findIndex((p, i) => i >= 3 && /^-?\d+(\.\d+)?$/.test(p))
  if (amountIndex === -1) return { error: '找不到有效金額' }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { error: '日期格式必須為 YYYY-MM-DD' }

  const description = parts.slice(3, amountIndex).join(' ')
  const amount = Number(parts[amountIndex])
  const payer = parts.slice(amountIndex + 1).join(' ')
  if (!description || !payer) return { error: '支出描述或支付者不可留空' }

  return {
    expense: {
      id: uid(),
      receiptNo,
      date,
      category,
      description,
      amount,
      payer,
      pdfData: '',
      pdfFileName: '',
      createdAt: new Date().toISOString()
    }
  }
}

function getUsed(activity) {
  return (activity.expenses || []).reduce((sum, item) => sum + Number(item.amount || 0), 0)
}

function getRemaining(activity) {
  return Number(activity.budget || 0) - getUsed(activity)
}

function hasAdvanceReceiptData(record) {
  return !!(record?.expectedDate || record?.amount || record?.imageData)
}

function getAdvanceReceipts(activity) {
  if (Array.isArray(activity.advanceReceipts)) {
    return activity.advanceReceipts.map((item, index) => ({
      id: item.id || `${activity.id || 'activity'}-advance-${index}`,
      ...item
    }))
  }
  if (hasAdvanceReceiptData(activity.advanceReceipt)) {
    return [{ id: `${activity.id || 'activity'}-advance-legacy`, ...activity.advanceReceipt }]
  }
  return []
}

function getAdvanceTotal(activity) {
  return getAdvanceReceipts(activity).reduce((sum, item) => sum + Number(item.amount || 0), 0)
}

function StatCard({ label, value, tone = 'blue', sub }) {
  const toneMap = {
    blue: 'bg-blue-50 text-blue-700 border-blue-100',
    green: 'bg-emerald-50 text-emerald-700 border-emerald-100',
    red: 'bg-rose-50 text-rose-700 border-rose-100',
    slate: 'bg-slate-50 text-slate-700 border-slate-100'
  }
  return (
    <div className={`rounded-3xl border p-4 ${toneMap[tone]}`}>
      <p className="text-sm opacity-80">{label}</p>
      <p className="mt-1 text-2xl font-bold">{value}</p>
      {sub && <p className="mt-1 text-xs opacity-75">{sub}</p>}
    </div>
  )
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="mb-1 block text-sm font-semibold text-slate-700">{label}</span>
      {children}
    </label>
  )
}

function inputClass() {
  return 'w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 outline-none transition focus:border-blue-500 focus:ring-4 focus:ring-blue-100'
}

function App() {
  const [page, setPage] = useState('home')
  const [activities, setActivities] = useState([])
  const [settings, setSettings] = useState(defaultSettings)
  const [activeId, setActiveId] = useState(null)
  const [toast, setToast] = useState('')
  const [form, setForm] = useState(emptyActivity)
  const [session, setSession] = useState(null)
  const [authLoading, setAuthLoading] = useState(isSupabaseConfigured)
  const [cloudReady, setCloudReady] = useState(false)
  const [cloudStatus, setCloudStatus] = useState(isSupabaseConfigured ? '尚未登入雲端同步' : '未設定 Supabase')
  const applyingRemoteRef = useRef(false)

  useEffect(() => {
    try {
      const savedActivities = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]')
      const savedSettings = JSON.parse(localStorage.getItem(SETTINGS_KEY) || 'null')
      setActivities(Array.isArray(savedActivities) ? savedActivities : [])
      setSettings({ ...defaultSettings, ...(savedSettings || {}) })
    } catch {
      setActivities([])
      setSettings(defaultSettings)
    }
  }, [])

  useEffect(() => {
    if (!supabase) return

    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session)
      setAuthLoading(false)
    })

    const { data: listener } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession)
      setCloudReady(false)
      setCloudStatus(nextSession ? '正在同步雲端資料...' : '尚未登入雲端同步')
    })

    return () => listener.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(activities))
  }, [activities])

  useEffect(() => {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings))
  }, [settings])

  useEffect(() => {
    if (!supabase || !session?.user) return

    let active = true
    const userId = session.user.id

    async function loadCloudSnapshot() {
      setCloudStatus('正在讀取雲端資料...')
      const localActivities = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]')
      const localSettings = JSON.parse(localStorage.getItem(SETTINGS_KEY) || 'null')

      const { data, error } = await supabase
        .from(CLOUD_TABLE)
        .select('activities, settings, updated_at')
        .eq('user_id', userId)
        .maybeSingle()

      if (!active) return

      if (error) {
        console.error(error)
        setCloudStatus('雲端同步讀取失敗，暫用本機資料')
        return
      }

      if (!data) {
        const snapshot = {
          user_id: userId,
          activities: Array.isArray(localActivities) ? localActivities : [],
          settings: { ...defaultSettings, ...(localSettings || {}) },
          updated_at: new Date().toISOString()
        }

        const { error: upsertError } = await supabase
          .from(CLOUD_TABLE)
          .upsert(snapshot, { onConflict: 'user_id' })

        if (upsertError) {
          console.error(upsertError)
          setCloudStatus('雲端同步建立失敗，暫用本機資料')
          return
        }

        if (!active) return
        setCloudReady(true)
        setCloudStatus('雲端同步已啟用')
        return
      }

      applyingRemoteRef.current = true
      setActivities(Array.isArray(data.activities) ? data.activities : [])
      setSettings({ ...defaultSettings, ...(data.settings || {}) })
      window.setTimeout(() => {
        applyingRemoteRef.current = false
      }, 0)
      setCloudReady(true)
      setCloudStatus('雲端同步已啟用')
    }

    loadCloudSnapshot()

    const channel = supabase
      .channel(`user-app-data-${userId}`)
      .on(
        'postgres_changes',
        { event: '*', schema: 'public', table: CLOUD_TABLE, filter: `user_id=eq.${userId}` },
        (payload) => {
          if (!payload.new) return
          applyingRemoteRef.current = true
          setActivities(Array.isArray(payload.new.activities) ? payload.new.activities : [])
          setSettings({ ...defaultSettings, ...(payload.new.settings || {}) })
          window.setTimeout(() => {
            applyingRemoteRef.current = false
          }, 0)
          setCloudStatus('已收到其他裝置更新')
        }
      )
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') setCloudStatus('雲端同步已啟用')
      })

    return () => {
      active = false
      supabase.removeChannel(channel)
    }
  }, [session])

  useEffect(() => {
    if (!supabase || !session?.user || !cloudReady || applyingRemoteRef.current) return

    const timer = window.setTimeout(async () => {
      const { error } = await supabase
        .from(CLOUD_TABLE)
        .upsert({
          user_id: session.user.id,
          activities,
          settings,
          updated_at: new Date().toISOString()
        }, { onConflict: 'user_id' })

      if (error) {
        console.error(error)
        setCloudStatus('雲端同步儲存失敗，已保留本機快取')
      } else {
        setCloudStatus('雲端同步已更新')
      }
    }, 500)

    return () => window.clearTimeout(timer)
  }, [activities, settings, session, cloudReady])

  function flash(message) {
    setToast(message)
    setTimeout(() => setToast(''), 2400)
  }

  function createActivity(e) {
    e.preventDefault()
    const required = ['activityName', 'activityCode', 'startDate', 'endDate', 'budget', 'category', 'expenseCategory', 'personInCharge']
    if (required.some(k => !String(form[k]).trim())) {
      flash('請填妥所有必填欄位')
      return
    }
    if (activities.some(a => a.activityCode.trim().toLowerCase() === form.activityCode.trim().toLowerCase())) {
      flash('活動編號已存在，請使用另一個編號')
      return
    }
    const item = {
      id: uid(),
      ...form,
      budget: Number(form.budget),
      advanceReceipts: [],
      expenses: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }
    setActivities(prev => [item, ...prev])
    setForm({
      ...emptyActivity,
      category: settings.activityCategories[0] || '社區活動',
      expenseCategory: settings.expenseCategories[0] || '物資'
    })
    setPage('manage')
    flash('活動已建立')
  }

  function updateActivity(updated) {
    setActivities(prev => prev.map(a => a.id === updated.id ? { ...updated, updatedAt: new Date().toISOString() } : a))
  }

  async function requirePassword(action) {
    // 首次使用需先到設定頁設置密碼，才可啟動刪除及清空等高風險功能。
    if (!settings.passwordHash) {
      flash('首次使用需設置密碼才能啟動刪除功能')
      setPage('settings')
      return false
    }

    const input = window.prompt('此操作需要密碼，請輸入密碼：')
    if (!input) return false

    const hash = await sha256(input)
    if (hash !== settings.passwordHash) {
      flash('密碼錯誤')
      return false
    }

    return action()
  }

  function exportExcel(activity) {
    const used = getUsed(activity)
    const remaining = getRemaining(activity)
    const advanceReceipts = getAdvanceReceipts(activity)
    const summary = [
      ['活動名稱', activity.activityName],
      ['活動編號', activity.activityCode],
      ['會計編號', activity.accountingCode || ''],
      ['活動日期', `${activity.startDate} 至 ${activity.endDate}`],
      ['活動類別', activity.category],
      ['支出類別', activity.expenseCategory || ''],
      ['活動負責人', activity.personInCharge],
      ['預算金額', activity.budget],
      ['已使用金額', used],
      ['剩餘金額', remaining],
      ['預支狀態', activity.advanceApplied],
      ['預支紀錄數量', advanceReceipts.length],
      ['預支領取總額', getAdvanceTotal(activity)],
      ['匯出日期', today()]
    ]
    const advances = advanceReceipts.map((x, index) => ({
      序號: index + 1,
      預計領取日期: x.expectedDate || '',
      金額: Number(x.amount || 0),
      圖片: x.imageData ? (x.imageName || '已上傳') : '未上傳',
      建立時間: x.createdAt || '',
      更新時間: x.updatedAt || ''
    }))
    const expenses = (activity.expenses || []).map(x => ({
      收據編號: x.receiptNo,
      日期: x.date,
      類別: x.category,
      支出描述: x.description,
      金額: Number(x.amount || 0),
      支付者: x.payer,
      PDF文件: x.pdfData ? `${x.receiptNo}.pdf` : '未上傳'
    }))
    const wb = XLSX.utils.book_new()
    const ws1 = XLSX.utils.aoa_to_sheet(summary)
    const ws2 = XLSX.utils.json_to_sheet(expenses)
    const ws3 = XLSX.utils.json_to_sheet(advances)
    XLSX.utils.book_append_sheet(wb, ws1, '活動基本資料')
    XLSX.utils.book_append_sheet(wb, ws2, '支出紀錄')
    XLSX.utils.book_append_sheet(wb, ws3, '預支紀錄')
    XLSX.writeFile(wb, `${activity.activityCode}_${activity.activityName}_支出紀錄.xlsx`)
  }

  function exportPDF(activity) {
    const doc = new jsPDF({ orientation: 'landscape' })
    const used = getUsed(activity)
    const remaining = getRemaining(activity)
    const advanceReceipts = getAdvanceReceipts(activity)

    doc.setFont('helvetica', 'bold')
    doc.setFontSize(18)
    doc.text(settings.systemName || '活動支出追蹤系統', 14, 18)
    doc.setFontSize(11)
    doc.setFont('helvetica', 'normal')
    doc.text(`Event Expense Report / Export Date: ${today()}`, 14, 27)

    autoTable(doc, {
      startY: 34,
      theme: 'grid',
      head: [['Item', 'Content', 'Item', 'Content']],
      body: [
        ['Activity Name', activity.activityName, 'Activity Code', activity.activityCode],
        ['Accounting Code', activity.accountingCode || '', 'Expense Category', activity.expenseCategory || ''],
        ['Date', `${activity.startDate} to ${activity.endDate}`, 'Category', activity.category],
        ['Person In Charge', activity.personInCharge, 'Advance Applied', activity.advanceApplied],
        ['Advance Records', String(advanceReceipts.length), 'Advance Total', currency(getAdvanceTotal(activity))],
        ['Budget', currency(activity.budget), 'Used', currency(used)],
        ['Remaining', currency(remaining), 'Status', remaining < 0 ? 'Over Budget' : 'Within Budget']
      ],
      styles: { font: 'helvetica', fontSize: 9 },
      headStyles: { fillColor: [37, 99, 235] }
    })

    if (advanceReceipts.length) {
      autoTable(doc, {
        startY: doc.lastAutoTable.finalY + 8,
        theme: 'striped',
        head: [['No.', 'Expected Date', 'Amount', 'Image']],
        body: advanceReceipts.map((x, index) => [
          index + 1, x.expectedDate || '', currency(x.amount || 0), x.imageData ? (x.imageName || 'Uploaded') : 'Not uploaded'
        ]),
        styles: { font: 'helvetica', fontSize: 8 },
        headStyles: { fillColor: [37, 99, 235] }
      })
    }

    autoTable(doc, {
      startY: doc.lastAutoTable.finalY + 8,
      theme: 'striped',
      head: [['Receipt No.', 'Date', 'Category', 'Description', 'Amount', 'Payer', 'PDF']],
      body: (activity.expenses || []).map(x => [
        x.receiptNo, x.date, x.category, x.description, currency(x.amount), x.payer, x.pdfData ? `${x.receiptNo}.pdf` : 'Not uploaded'
      ]),
      styles: { font: 'helvetica', fontSize: 8 },
      headStyles: { fillColor: [15, 23, 42] }
    })

    doc.save(`${activity.activityCode}_${activity.activityName}_支出報告.pdf`)
  }

  const activeActivity = activities.find(a => a.id === activeId)

  return (
    <div className="min-h-screen">
      <div className="mx-auto flex max-w-7xl">
        <aside className="sticky top-0 hidden h-screen w-72 shrink-0 border-r border-slate-200 bg-white/80 p-6 backdrop-blur md:block">
          <Brand systemName={settings.systemName} />
          <DesktopNav page={page} setPage={setPage} />
        </aside>

        <main className="safe-bottom w-full px-4 py-5 md:px-8 md:py-8">
          <MobileHeader systemName={settings.systemName} />
          {toast && <div className="fixed left-1/2 top-4 z-50 -translate-x-1/2 rounded-2xl bg-slate-950 px-4 py-3 text-sm font-semibold text-white shadow-soft">{toast}</div>}
          <CloudSyncBar
            session={session}
            authLoading={authLoading}
            cloudStatus={cloudStatus}
            flash={flash}
          />

          {page === 'home' && (
            <HomePage
              settings={settings}
              form={form}
              setForm={setForm}
              createActivity={createActivity}
              setPage={setPage}
            />
          )}

          {page === 'manage' && !activeActivity && (
            <ManagePage
              activities={activities}
              settings={settings}
              setSettings={setSettings}
              setActiveId={setActiveId}
              exportExcel={exportExcel}
              exportPDF={exportPDF}
              requirePassword={requirePassword}
              setActivities={setActivities}
              flash={flash}
            />
          )}

          {page === 'manage' && activeActivity && (
            <ActivityDetail
              activity={activeActivity}
              settings={settings}
              updateActivity={updateActivity}
              setActiveId={setActiveId}
              exportExcel={exportExcel}
              exportPDF={exportPDF}
              requirePassword={requirePassword}
              flash={flash}
            />
          )}

          {page === 'settings' && (
            <SettingsPage
              settings={settings}
              setSettings={setSettings}
              activities={activities}
              setActivities={setActivities}
              requirePassword={requirePassword}
              flash={flash}
            />
          )}
        </main>
      </div>

      <BottomNav page={page} setPage={(p) => { setActiveId(null); setPage(p) }} />
    </div>
  )
}

function CloudSyncBar({ session, authLoading, cloudStatus, flash }) {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [busy, setBusy] = useState(false)

  async function signIn(e) {
    e.preventDefault()
    if (!supabase || !email || !password) return
    setBusy(true)
    const { error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) flash(error.message)
    else {
      setEmail('')
      setPassword('')
      flash('已登入雲端同步')
    }
    setBusy(false)
  }

  async function signUp() {
    if (!supabase || !email || !password) return
    setBusy(true)
    const { error } = await supabase.auth.signUp({ email, password })
    if (error) flash(error.message)
    else flash('帳號已建立，請按 Supabase Auth 設定完成確認')
    setBusy(false)
  }

  async function signOut() {
    if (!supabase) return
    await supabase.auth.signOut()
    flash('已登出雲端同步')
  }

  if (!isSupabaseConfigured) {
    return (
      <div className="mb-5 rounded-3xl border border-amber-200 bg-amber-50 p-4 text-sm font-semibold text-amber-800 shadow-soft">
        Supabase 尚未設定，請加入 VITE_SUPABASE_URL 與 VITE_SUPABASE_ANON_KEY。
      </div>
    )
  }

  if (authLoading) {
    return <div className="mb-5 rounded-3xl bg-white p-4 text-sm font-semibold text-slate-600 shadow-soft">正在檢查雲端登入...</div>
  }

  if (session?.user) {
    return (
      <div className="mb-5 flex flex-col gap-3 rounded-3xl bg-white p-4 shadow-soft md:flex-row md:items-center md:justify-between">
        <div>
          <p className="text-sm font-black text-slate-800">{session.user.email}</p>
          <p className="text-xs font-semibold text-slate-500">{cloudStatus}</p>
        </div>
        <button onClick={signOut} className="btn-muted">登出</button>
      </div>
    )
  }

  return (
    <form onSubmit={signIn} className="mb-5 grid gap-2 rounded-3xl bg-white p-4 shadow-soft md:grid-cols-[1fr_1fr_auto_auto]">
      <input type="email" className={inputClass()} placeholder="Email" value={email} onChange={e => setEmail(e.target.value)} />
      <input type="password" className={inputClass()} placeholder="Password" value={password} onChange={e => setPassword(e.target.value)} />
      <button disabled={busy} className="btn-primary">登入同步</button>
      <button type="button" disabled={busy} onClick={signUp} className="btn-muted">建立帳號</button>
    </form>
  )
}

function Brand({ systemName }) {
  return (
    <div className="mb-8">
      <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-3xl bg-blue-600 text-white shadow-soft">
        <ReceiptText />
      </div>
      <h1 className="text-xl font-black text-slate-900">{systemName}</h1>
      <p className="mt-1 text-sm text-slate-500">清楚管理每個活動的預算、收據及支出紀錄。</p>
    </div>
  )
}

function DesktopNav({ page, setPage }) {
  const items = [
    ['home', Home, '首頁'],
    ['manage', FolderKanban, '活動管理'],
    ['settings', Settings, '設定']
  ]
  return <div className="space-y-2">{items.map(([key, Icon, label]) => (
    <button key={key} onClick={() => setPage(key)} className={`flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left font-bold ${page === key ? 'bg-blue-600 text-white shadow-soft' : 'text-slate-600 hover:bg-slate-100'}`}>
      <Icon size={20} /> {label}
    </button>
  ))}</div>
}

function MobileHeader({ systemName }) {
  return (
    <header className="mb-5 rounded-3xl bg-white/85 p-5 shadow-soft backdrop-blur md:hidden">
      <p className="text-xs font-bold uppercase tracking-widest text-blue-600">Event Expense Tracker</p>
      <h1 className="mt-1 text-2xl font-black text-slate-950">{systemName}</h1>
      <p className="mt-1 text-sm text-slate-500">預算、收據、支出一頁管理。</p>
    </header>
  )
}

function BottomNav({ page, setPage }) {
  const items = [
    ['home', Home, '首頁'],
    ['manage', FolderKanban, '活動'],
    ['settings', Settings, '設定']
  ]
  return (
    <nav className="fixed inset-x-0 bottom-0 z-40 border-t border-slate-200 bg-white/95 px-4 pb-[env(safe-area-inset-bottom)] pt-2 shadow-soft backdrop-blur md:hidden">
      <div className="mx-auto grid max-w-md grid-cols-3 gap-2">
        {items.map(([key, Icon, label]) => (
          <button key={key} onClick={() => setPage(key)} className={`rounded-2xl px-3 py-2 text-xs font-bold ${page === key ? 'bg-blue-600 text-white' : 'text-slate-500'}`}>
            <Icon className="mx-auto mb-1" size={20} />
            {label}
          </button>
        ))}
      </div>
    </nav>
  )
}

function HomePage({ settings, form, setForm, createActivity, setPage }) {
  return (
    <div className="space-y-6">
      <section className="grid gap-4 md:grid-cols-3">
        <ActionCard icon={<PlusCircle />} title="創建活動" text="建立活動檔案、預算及負責人資料。" />
        <button onClick={() => setPage('manage')}><ActionCard icon={<FolderKanban />} title="管理活動" text="查看支出、匯出 Excel / PDF。" /></button>
        <button onClick={() => setPage('settings')}><ActionCard icon={<ShieldCheck />} title="設定" text="管理密碼、類別及備份還原。" /></button>
      </section>

      <section className="rounded-[2rem] bg-white p-5 shadow-soft md:p-8">
        <div className="mb-5">
          <p className="text-sm font-bold text-blue-600">Create Activity</p>
          <h2 className="text-2xl font-black text-slate-950">創建活動</h2>
        </div>

        <form onSubmit={createActivity} className="grid gap-4 md:grid-cols-2">
          <Field label="活動名稱"><input className={inputClass()} value={form.activityName} onChange={e => setForm({ ...form, activityName: e.target.value })} /></Field>
          <Field label="活動編號"><input className={inputClass()} value={form.activityCode} onChange={e => setForm({ ...form, activityCode: e.target.value })} placeholder="例如：ACT-2026-001" /></Field>
          <Field label="開始日期"><input type="date" className={inputClass()} value={form.startDate} onChange={e => setForm({ ...form, startDate: e.target.value })} /></Field>
          <Field label="結束日期"><input type="date" className={inputClass()} value={form.endDate} onChange={e => setForm({ ...form, endDate: e.target.value })} /></Field>
          <Field label="預算金額"><input type="number" min="0" step="0.01" className={inputClass()} value={form.budget} onChange={e => setForm({ ...form, budget: e.target.value })} /></Field>
          <Field label="活動類別">
            <select className={inputClass()} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>
              {settings.activityCategories.map(c => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="會計編號"><input className={inputClass()} value={form.accountingCode} onChange={e => setForm({ ...form, accountingCode: e.target.value })} placeholder="例如：AC-2026-001" /></Field>
          <Field label="支出類別">
            <select className={inputClass()} value={form.expenseCategory} onChange={e => setForm({ ...form, expenseCategory: e.target.value })}>
              {settings.expenseCategories.map(c => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="活動負責人"><input className={inputClass()} value={form.personInCharge} onChange={e => setForm({ ...form, personInCharge: e.target.value })} /></Field>
          <Field label="有否申請預支">
            <select className={inputClass()} value={form.advanceApplied} onChange={e => setForm({ ...form, advanceApplied: e.target.value })}>
              <option>有</option><option>沒有</option>
            </select>
          </Field>
          <button className="md:col-span-2 rounded-2xl bg-blue-600 px-5 py-4 font-black text-white shadow-soft hover:bg-blue-700">確認創建</button>
        </form>
      </section>
    </div>
  )
}

function ActionCard({ icon, title, text }) {
  return (
    <div className="h-full rounded-[2rem] border border-white bg-white p-5 text-left shadow-soft transition hover:-translate-y-1">
      <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-50 text-blue-600">{icon}</div>
      <h3 className="text-lg font-black">{title}</h3>
      <p className="mt-1 text-sm text-slate-500">{text}</p>
    </div>
  )
}


function EditActivityModal({ activity, settings, onClose, onSave, activities }) {
  const [editForm, setEditForm] = useState({
    activityName: activity.activityName || '',
    activityCode: activity.activityCode || '',
    startDate: activity.startDate || '',
    endDate: activity.endDate || '',
    budget: activity.budget || '',
    category: activity.category || (settings.activityCategories[0] || '社區活動'),
    accountingCode: activity.accountingCode || '',
    expenseCategory: activity.expenseCategory || (settings.expenseCategories[0] || '物資'),
    personInCharge: activity.personInCharge || '',
    advanceApplied: activity.advanceApplied || '沒有'
  })

  function submit(e) {
    e.preventDefault()
    const required = ['activityName', 'activityCode', 'startDate', 'endDate', 'budget', 'category', 'expenseCategory', 'personInCharge']
    if (required.some(k => !String(editForm[k]).trim())) {
      alert('請填妥所有必填欄位')
      return
    }
    const duplicated = activities.some(a =>
      a.id !== activity.id &&
      String(a.activityCode).trim().toLowerCase() === String(editForm.activityCode).trim().toLowerCase()
    )
    if (duplicated) {
      alert('活動編號已存在，請使用另一個編號')
      return
    }
    onSave({
      ...activity,
      ...editForm,
      budget: Number(editForm.budget),
      updatedAt: new Date().toISOString()
    })
  }

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-950/50 p-0 backdrop-blur-sm md:items-center md:p-6">
      <div className="max-h-[92vh] w-full max-w-3xl overflow-y-auto rounded-t-[2rem] bg-white p-5 shadow-soft md:rounded-[2rem] md:p-8">
        <div className="mb-5 flex items-start justify-between gap-4">
          <div>
            <p className="text-sm font-bold text-blue-600">Edit Activity</p>
            <h3 className="text-2xl font-black text-slate-950">更改活動內容</h3>
            <p className="mt-1 text-sm text-slate-500">只會更新活動基本資料，原有支出紀錄及已上傳 PDF 會保留。</p>
          </div>
          <button type="button" onClick={onClose} className="rounded-2xl bg-slate-100 px-4 py-2 font-black text-slate-600">關閉</button>
        </div>

        <form onSubmit={submit} className="grid gap-4 md:grid-cols-2">
          <Field label="活動名稱">
            <input className={inputClass()} value={editForm.activityName} onChange={e => setEditForm({ ...editForm, activityName: e.target.value })} />
          </Field>
          <Field label="活動編號">
            <input className={inputClass()} value={editForm.activityCode} onChange={e => setEditForm({ ...editForm, activityCode: e.target.value })} />
          </Field>
          <Field label="開始日期">
            <input type="date" className={inputClass()} value={editForm.startDate} onChange={e => setEditForm({ ...editForm, startDate: e.target.value })} />
          </Field>
          <Field label="結束日期">
            <input type="date" className={inputClass()} value={editForm.endDate} onChange={e => setEditForm({ ...editForm, endDate: e.target.value })} />
          </Field>
          <Field label="預算金額">
            <input type="number" min="0" step="0.01" className={inputClass()} value={editForm.budget} onChange={e => setEditForm({ ...editForm, budget: e.target.value })} />
          </Field>
          <Field label="活動類別">
            <select className={inputClass()} value={editForm.category} onChange={e => setEditForm({ ...editForm, category: e.target.value })}>
              {settings.activityCategories.map(c => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="會計編號">
            <input className={inputClass()} value={editForm.accountingCode} onChange={e => setEditForm({ ...editForm, accountingCode: e.target.value })} />
          </Field>
          <Field label="支出類別">
            <select className={inputClass()} value={editForm.expenseCategory} onChange={e => setEditForm({ ...editForm, expenseCategory: e.target.value })}>
              {settings.expenseCategories.map(c => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="活動負責人">
            <input className={inputClass()} value={editForm.personInCharge} onChange={e => setEditForm({ ...editForm, personInCharge: e.target.value })} />
          </Field>
          <Field label="有否申請預支">
            <select className={inputClass()} value={editForm.advanceApplied} onChange={e => setEditForm({ ...editForm, advanceApplied: e.target.value })}>
              <option>有</option>
              <option>沒有</option>
            </select>
          </Field>

          <div className="flex flex-col gap-2 md:col-span-2 md:flex-row">
            <button className="btn-primary flex-1"><Save size={16} /> 儲存更改</button>
            <button type="button" onClick={onClose} className="btn-muted flex-1">取消</button>
          </div>
        </form>
      </div>
    </div>
  )
}

function ManagePage({ activities, settings, setSettings, setActiveId, exportExcel, exportPDF, requirePassword, setActivities, flash }) {
  const [query, setQuery] = useState('')
  const [filter, setFilter] = useState('全部')
  const [editingActivity, setEditingActivity] = useState(null)

  const filtered = useMemo(() => {
    let list = activities.filter(a => {
      const hit = `${a.activityName} ${a.activityCode} ${a.accountingCode || ''} ${a.personInCharge}`.toLowerCase().includes(query.toLowerCase())
      const cat = filter === '全部' || a.category === filter
      return hit && cat
    })
    const sort = settings.sortMode
    list = [...list].sort((a, b) => {
      if (sort === 'created_asc') return new Date(a.createdAt) - new Date(b.createdAt)
      if (sort === 'start_asc') return new Date(a.startDate) - new Date(b.startDate)
      if (sort === 'start_desc') return new Date(b.startDate) - new Date(a.startDate)
      if (sort === 'name_az') return a.activityName.localeCompare(b.activityName, 'zh-Hant')
      if (sort === 'remain_desc') return getRemaining(b) - getRemaining(a)
      if (sort === 'used_desc') return getUsed(b) - getUsed(a)
      return new Date(b.createdAt) - new Date(a.createdAt)
    })
    return list
  }, [activities, query, filter, settings.sortMode])

  async function deleteActivity(id) {
    await requirePassword(() => {
      if (!window.confirm('確定刪除此活動？此操作不能復原。')) return false
      setActivities(prev => prev.filter(a => a.id !== id))
      flash('活動已刪除')
      return true
    })
  }

  function saveActivityEdit(updatedActivity) {
    setActivities(prev => prev.map(a => a.id === updatedActivity.id ? updatedActivity : a))
    setEditingActivity(null)
    flash('活動內容已更新')
  }

  return (
    <div className="space-y-5">
      {editingActivity && (
        <EditActivityModal
          activity={editingActivity}
          settings={settings}
          activities={activities}
          onClose={() => setEditingActivity(null)}
          onSave={saveActivityEdit}
        />
      )}
      <HeaderBlock title="活動管理" subtitle="查看、排序、搜尋及匯出活動支出紀錄。" />
      {!settings.passwordHash && (
        <div className="flex items-start gap-3 rounded-3xl border border-amber-200 bg-amber-50 p-4 text-amber-800 shadow-sm">
          <ShieldCheck className="mt-0.5 shrink-0" size={20} />
          <div>
            <p className="font-black">首次使用需設置密碼才能啟動刪除功能</p>
            <p className="mt-1 text-sm">請先到「設定」頁設定密碼；完成後，刪除活動、清空資料及還原資料等高風險操作才可使用。</p>
          </div>
        </div>
      )}
      <div className="grid gap-3 rounded-[2rem] bg-white p-4 shadow-soft md:grid-cols-4">
        <div className="relative md:col-span-2">
          <Search className="absolute left-4 top-3.5 text-slate-400" size={18} />
          <input className={`${inputClass()} pl-11`} placeholder="搜尋活動名稱、活動編號、會計編號、負責人" value={query} onChange={e => setQuery(e.target.value)} />
        </div>
        <select className={inputClass()} value={filter} onChange={e => setFilter(e.target.value)}>
          <option>全部</option>
          {settings.activityCategories.map(c => <option key={c}>{c}</option>)}
        </select>
        <select className={inputClass()} value={settings.sortMode} onChange={e => setSettings({ ...settings, sortMode: e.target.value })}>
          <option value="created_desc">建立日期由新至舊</option>
          <option value="created_asc">建立日期由舊至新</option>
          <option value="start_asc">活動開始日期由近至遠</option>
          <option value="start_desc">活動開始日期由遠至近</option>
          <option value="name_az">活動名稱 A-Z</option>
          <option value="remain_desc">剩餘預算由高至低</option>
          <option value="used_desc">已使用金額由高至低</option>
        </select>
      </div>

      <div className="grid gap-4 lg:grid-cols-2">
        {filtered.map(activity => {
          const used = getUsed(activity)
          const remaining = getRemaining(activity)
          const percent = activity.budget > 0 ? Math.round((used / activity.budget) * 100) : 0
          const advanceReceipts = getAdvanceReceipts(activity)
          const nextAdvance = [...advanceReceipts].filter(x => x.expectedDate).sort((a, b) => new Date(a.expectedDate) - new Date(b.expectedDate))[0]
          return (
            <article key={activity.id} className="rounded-[2rem] bg-white p-5 shadow-soft">
              <div className="flex items-start justify-between gap-3">
                <div>
                  <p className="text-xs font-bold uppercase tracking-widest text-blue-600">{activity.activityCode}</p>
                  <h3 className="text-xl font-black text-slate-950">{activity.activityName}</h3>
                  <p className="mt-1 text-sm text-slate-500">{activity.startDate} 至 {activity.endDate} · {activity.category} · {activity.personInCharge}</p>
                  <p className="mt-1 text-sm text-slate-500">會計編號：{activity.accountingCode || '未填寫'} · 支出類別：{activity.expenseCategory || '未設定'}</p>
                </div>
                {remaining < 0 && <span className="rounded-full bg-rose-100 px-3 py-1 text-xs font-black text-rose-700">超支</span>}
              </div>

              <div className="mt-4 grid grid-cols-3 gap-2">
                <MiniMoney label="預算" value={currency(activity.budget)} />
                <MiniMoney label="已使用" value={currency(used)} />
                <MiniMoney label="剩餘" value={currency(remaining)} danger={remaining < 0} />
              </div>
              <div className="mt-4 h-3 overflow-hidden rounded-full bg-slate-100">
                <div className={`h-full ${percent > 100 ? 'bg-rose-500' : 'bg-blue-600'}`} style={{ width: `${Math.min(percent, 100)}%` }} />
              </div>
              <p className="mt-2 text-xs text-slate-500">預算使用率：{percent}% · 預支：{activity.advanceApplied}</p>
              <div className="mt-3 rounded-2xl bg-slate-50 p-3 text-sm text-slate-600">
                <p className="font-black text-slate-800">預支領取</p>
                <p className="mt-1">紀錄：{advanceReceipts.length} 筆 · 總額：{currency(getAdvanceTotal(activity))}</p>
                <p className="mt-1 text-xs font-semibold text-slate-500">最近預計日期：{nextAdvance?.expectedDate || '未填寫'}</p>
              </div>

              <div className="mt-4 flex flex-wrap gap-2">
                <button onClick={() => setActiveId(activity.id)} className="btn-primary"><Eye size={16} /> 查看</button>
                <button onClick={(e) => { e.preventDefault(); e.stopPropagation(); setEditingActivity(activity) }} className="btn-muted"><Pencil size={16} /> 更改活動內容</button>
                <button onClick={() => exportExcel(activity)} className="btn-muted"><Download size={16} /> Excel</button>
                <button onClick={() => exportPDF(activity)} className="btn-muted"><FileText size={16} /> PDF</button>
                <button title={!settings.passwordHash ? '首次使用需設置密碼才能啟動刪除功能' : '刪除活動'} onClick={(e) => { e.preventDefault(); e.stopPropagation(); deleteActivity(activity.id) }} className="btn-danger"><Trash2 size={16} /> 刪除</button>
              </div>
            </article>
          )
        })}
      </div>

      {filtered.length === 0 && <Empty text="暫時未有活動，請先到首頁建立活動。" />}
    </div>
  )
}

function ActivityDetail({ activity, settings, updateActivity, setActiveId, exportExcel, exportPDF, requirePassword, flash }) {
  const [bulk, setBulk] = useState('')
  const [preview, setPreview] = useState([])
  const [errors, setErrors] = useState([])
  const [single, setSingle] = useState({ ...emptyExpense, category: settings.expenseCategories[0] || '物資' })
  const [expenseQuery, setExpenseQuery] = useState('')
  const [catFilter, setCatFilter] = useState('全部')
  const [editingId, setEditingId] = useState(null)
  const [draft, setDraft] = useState(null)

  const used = getUsed(activity)
  const remaining = getRemaining(activity)
  const percent = activity.budget > 0 ? Math.round((used / activity.budget) * 100) : 0
  const advanceReceipts = getAdvanceReceipts(activity)

  const filteredExpenses = useMemo(() => {
    return [...(activity.expenses || [])]
      .filter(x => catFilter === '全部' || x.category === catFilter)
      .filter(x => `${x.receiptNo} ${x.description} ${x.payer}`.toLowerCase().includes(expenseQuery.toLowerCase()))
      .sort((a, b) => new Date(a.date) - new Date(b.date))
  }, [activity.expenses, expenseQuery, catFilter])

  const categoryStats = useMemo(() => {
    const map = {}
    for (const item of activity.expenses || []) map[item.category] = (map[item.category] || 0) + Number(item.amount || 0)
    return Object.entries(map).sort((a, b) => b[1] - a[1])
  }, [activity.expenses])

  function buildPreview() {
    const lines = bulk.split('\n').map(x => x.trim()).filter(Boolean)
    const good = []
    const bad = []
    lines.forEach((line, index) => {
      const parsed = parseExpenseLine(line)
      if (parsed.error) bad.push(`第 ${index + 1} 行：${parsed.error}`)
      else good.push(parsed.expense)
    })
    setPreview(good)
    setErrors(bad)
    if (!lines.length) flash('請先貼上收據資料')
  }

  function confirmBulk() {
    if (!preview.length) {
      flash('沒有可儲存的有效資料')
      return
    }
    updateActivity({ ...activity, expenses: [...(activity.expenses || []), ...preview] })
    setBulk('')
    setPreview([])
    setErrors([])
    flash('批量支出已儲存')
  }

  function addSingle(e) {
    e.preventDefault()
    if (!single.receiptNo || !single.date || !single.category || !single.description || !single.amount || !single.payer) {
      flash('請填妥單筆支出資料')
      return
    }
    updateActivity({ ...activity, expenses: [...(activity.expenses || []), { ...single, id: uid(), amount: Number(single.amount), createdAt: new Date().toISOString() }] })
    setSingle({ ...emptyExpense, category: settings.expenseCategories[0] || '物資' })
    flash('支出已新增')
  }

  async function deleteExpense(id) {
    await requirePassword(() => {
      if (!window.confirm('確定刪除此支出紀錄？')) return false
      updateActivity({ ...activity, expenses: activity.expenses.filter(x => x.id !== id) })
      flash('支出紀錄已刪除')
      return true
    })
  }

  function startEdit(expense) {
    setEditingId(expense.id)
    setDraft({ ...expense })
  }

  function saveEdit() {
    updateActivity({ ...activity, expenses: activity.expenses.map(x => x.id === editingId ? { ...draft, amount: Number(draft.amount) } : x) })
    setEditingId(null)
    setDraft(null)
    flash('支出已更新')
  }

  async function uploadReceiptPdf(expenseId, file) {
    if (!file) return
    if (file.type !== 'application/pdf') {
      flash('請上傳 PDF 檔案')
      return
    }
    const dataUrl = await fileToDataUrl(file)
    updateActivity({
      ...activity,
      expenses: activity.expenses.map(x => {
        if (x.id !== expenseId) return x
        const generatedName = `${safeFileName(x.receiptNo)}.pdf`
        return {
          ...x,
          pdfData: dataUrl,
          pdfFileName: generatedName,
          pdfOriginalName: file.name,
          pdfUploadedAt: new Date().toISOString()
        }
      })
    })
    flash('PDF 已上傳，請按 PDF 欄目檢視或下載')
  }

  function previewReceiptPdf(expense) {
    if (!expense.pdfData) {
      flash('此收據尚未上傳 PDF')
      return
    }
    openPdfPreview(expense.pdfData)
  }

  function downloadReceiptPdf(expense) {
    if (!expense.pdfData) {
      flash('此收據尚未上傳 PDF')
      return
    }
    downloadDataUrl(`${safeFileName(expense.receiptNo)}.pdf`, expense.pdfData)
  }

  function syncAdvanceReceipts(nextReceipts) {
    updateActivity({
      ...activity,
      advanceReceipts: nextReceipts
    })
  }

  function addAdvanceReceipt() {
    syncAdvanceReceipts([
      ...advanceReceipts,
      {
        id: uid(),
        expectedDate: '',
        amount: '',
        imageData: '',
        imageName: '',
        imageUploadedAt: '',
        createdAt: new Date().toISOString()
      }
    ])
    flash('已新增一筆預支紀錄')
  }

  function updateAdvanceReceipt(receiptId, patch) {
    syncAdvanceReceipts(advanceReceipts.map(item => (
      item.id === receiptId
        ? { ...item, ...patch, updatedAt: new Date().toISOString() }
        : item
    )))
  }

  async function deleteAdvanceReceipt(receiptId) {
    await requirePassword(() => {
      if (!window.confirm('確定刪除此預支紀錄？')) return false
      syncAdvanceReceipts(advanceReceipts.filter(item => item.id !== receiptId))
      flash('預支紀錄已刪除')
      return true
    })
  }

  async function uploadAdvanceReceiptImage(receiptId, file) {
    if (!file) return
    if (!file.type.startsWith('image/')) {
      flash('請上傳圖片檔案')
      return
    }
    if (file.size > 3 * 1024 * 1024) {
      flash('圖片不可大於 3MB，請先壓縮後再上傳')
      return
    }
    const dataUrl = await fileToDataUrl(file)
    updateAdvanceReceipt(receiptId, {
      imageData: dataUrl,
      imageName: file.name,
      imageUploadedAt: new Date().toISOString()
    })
    flash('預支領取圖片已上傳')
  }

  function previewAdvanceReceiptImage(record) {
    if (!record.imageData) {
      flash('尚未上傳預支領取圖片')
      return
    }
    openImagePreview(record.imageData, '預支領取圖片')
  }

  function downloadAdvanceReceiptImage(record) {
    if (!record.imageData) {
      flash('尚未上傳預支領取圖片')
      return
    }
    downloadDataUrl(record.imageName || `${safeFileName(activity.activityCode)}_advance.jpg`, record.imageData)
  }

  return (
    <div className="space-y-5">
      <button onClick={() => setActiveId(null)} className="inline-flex items-center gap-2 rounded-2xl bg-white px-4 py-3 font-bold text-slate-600 shadow-soft"><ArrowLeft size={18} /> 返回活動列表</button>
      <HeaderBlock title={activity.activityName} subtitle={`${activity.activityCode} · ${activity.startDate} 至 ${activity.endDate} · ${activity.personInCharge}`} />

      <section className="grid gap-3 md:grid-cols-3">
        <MiniMoney label="會計編號" value={activity.accountingCode || '未填寫'} />
        <MiniMoney label="活動類別" value={activity.category || '未設定'} />
        <MiniMoney label="支出類別" value={activity.expenseCategory || '未設定'} />
      </section>

      <section className="grid gap-3 md:grid-cols-4">
        <StatCard label="預算金額" value={currency(activity.budget)} />
        <StatCard label="已使用金額" value={currency(used)} tone="slate" />
        <StatCard label="剩餘金額" value={currency(remaining)} tone={remaining < 0 ? 'red' : 'green'} sub={remaining < 0 ? '已超出預算' : '仍在預算內'} />
        <StatCard label="預算使用率" value={`${percent}%`} tone={percent > 100 ? 'red' : 'blue'} />
      </section>

      {remaining < 0 && (
        <div className="flex items-center gap-3 rounded-3xl border border-rose-200 bg-rose-50 p-4 font-bold text-rose-700">
          <AlertTriangle /> 此活動已超出預算，請檢查支出或調整預算。
        </div>
      )}

      <div className="flex flex-wrap gap-2 rounded-[2rem] bg-white p-4 shadow-soft">
        <button onClick={() => exportExcel(activity)} className="btn-primary"><Download size={16} /> 生成 Excel</button>
        <button onClick={() => exportPDF(activity)} className="btn-muted"><FileText size={16} /> 生成 PDF</button>
      </div>

      <section className="rounded-[2rem] bg-white p-5 shadow-soft">
        <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
          <div>
            <h3 className="text-xl font-black">預支領取</h3>
            <p className="mt-1 text-sm text-slate-500">可新增多筆預支紀錄；每筆紀錄都可保存預計領取日期、金額及圖片。</p>
          </div>
          <button onClick={addAdvanceReceipt} className="btn-primary"><PlusCircle size={16} /> 新增預支</button>
        </div>
        <div className="mb-4 grid gap-2 md:grid-cols-3">
          <MiniMoney label="預支紀錄" value={`${advanceReceipts.length} 筆`} />
          <MiniMoney label="預支總額" value={currency(getAdvanceTotal(activity))} />
          <MiniMoney label="已上傳圖片" value={`${advanceReceipts.filter(x => x.imageData).length} 張`} />
        </div>

        <div className="space-y-3">
          {advanceReceipts.map((record, index) => (
            <div key={record.id} className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
              <div className="mb-3 flex items-center justify-between gap-3">
                <p className="font-black text-slate-800">預支紀錄 {index + 1}</p>
                <button onClick={() => deleteAdvanceReceipt(record.id)} className="btn-danger"><Trash2 size={16} /> 刪除</button>
              </div>
              <div className="grid gap-3 md:grid-cols-3">
                <Field label="預計領取日期">
                  <input type="date" className={inputClass()} value={record.expectedDate || ''} onChange={e => updateAdvanceReceipt(record.id, { expectedDate: e.target.value })} />
                </Field>
                <Field label="預支領取金額">
                  <input type="number" min="0" step="0.01" className={inputClass()} value={record.amount || ''} onChange={e => updateAdvanceReceipt(record.id, { amount: e.target.value })} />
                </Field>
                <Field label="圖片">
                  <label className="flex min-h-[50px] cursor-pointer items-center justify-center gap-2 rounded-2xl border border-dashed border-slate-300 bg-white px-4 py-3 font-black text-slate-600 hover:bg-slate-100">
                    <Upload size={16} /> 上傳圖片
                    <input type="file" accept="image/*" className="hidden" onChange={e => uploadAdvanceReceiptImage(record.id, e.target.files?.[0])} />
                  </label>
                </Field>
              </div>
              {record.imageData && (
                <div className="mt-3 flex flex-col gap-3 rounded-2xl bg-white p-3 md:flex-row md:items-center md:justify-between">
                  <div>
                    <p className="font-bold text-slate-800">{record.imageName || '已上傳圖片'}</p>
                    <p className="text-xs text-slate-500">{record.imageUploadedAt ? `上傳時間：${new Date(record.imageUploadedAt).toLocaleString('zh-HK')}` : '已儲存圖片'}</p>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    <button onClick={() => previewAdvanceReceiptImage(record)} className="btn-primary"><Eye size={16} /> 預覽</button>
                    <button onClick={() => downloadAdvanceReceiptImage(record)} className="btn-muted"><Download size={16} /> 下載</button>
                  </div>
                </div>
              )}
            </div>
          ))}
          {!advanceReceipts.length && <Empty text="未有預支紀錄。按「新增預支」建立第一筆紀錄。" />}
        </div>
      </section>

      <section className="rounded-[2rem] bg-white p-5 shadow-soft">
        <h3 className="mb-4 flex items-center gap-2 text-xl font-black"><PlusCircle className="text-blue-600" /> 新增單筆支出</h3>
        <form onSubmit={addSingle} className="grid gap-3 md:grid-cols-6">
          <input className={inputClass()} placeholder="收據編號" value={single.receiptNo} onChange={e => setSingle({ ...single, receiptNo: e.target.value })} />
          <input type="date" className={inputClass()} value={single.date} onChange={e => setSingle({ ...single, date: e.target.value })} />
          <select className={inputClass()} value={single.category} onChange={e => setSingle({ ...single, category: e.target.value })}>{settings.expenseCategories.map(c => <option key={c}>{c}</option>)}</select>
          <input className={inputClass()} placeholder="支出描述" value={single.description} onChange={e => setSingle({ ...single, description: e.target.value })} />
          <input type="number" step="0.01" className={inputClass()} placeholder="金額" value={single.amount} onChange={e => setSingle({ ...single, amount: e.target.value })} />
          <input className={inputClass()} placeholder="支付者" value={single.payer} onChange={e => setSingle({ ...single, payer: e.target.value })} />
          <button className="rounded-2xl bg-blue-600 px-4 py-3 font-black text-white md:col-span-6">新增支出</button>
        </form>
      </section>

      <section className="rounded-[2rem] bg-white p-5 shadow-soft">
        <h3 className="mb-2 text-xl font-black">批量輸入支出紀錄</h3>
        <p className="mb-3 text-sm text-slate-500">格式：收據編號 日期 類別 支出描述 金額 支付者。支援雙引號，例如 "大型活動佈置材料"。</p>
        <textarea className={`${inputClass()} min-h-40 font-mono text-sm`} value={bulk} onChange={e => setBulk(e.target.value)} placeholder={'R001 2026-05-01 物資 文具及活動材料 238.5 陳大文\nR004 2026-05-04 物資 "大型活動佈置材料" 1200 "陳大文"'} />
        <div className="mt-3 flex flex-wrap gap-2">
          <button onClick={buildPreview} className="btn-primary">產生預覽</button>
          <button onClick={confirmBulk} className="btn-muted">確認儲存</button>
        </div>
        {!!errors.length && <div className="mt-4 rounded-2xl bg-rose-50 p-4 text-sm font-semibold text-rose-700">{errors.map(e => <p key={e}>{e}</p>)}</div>}
        {!!preview.length && (
          <div className="table-wrap mt-4">
            <table className="w-full min-w-[720px] text-left text-sm">
              <thead><tr className="border-b bg-slate-50">{['收據編號','日期','類別','支出描述','金額','支付者'].map(h => <th className="p-3" key={h}>{h}</th>)}</tr></thead>
              <tbody>{preview.map(x => <tr className="border-b" key={x.id}><td className="p-3">{x.receiptNo}</td><td className="p-3">{x.date}</td><td className="p-3">{x.category}</td><td className="p-3">{x.description}</td><td className="p-3">{currency(x.amount)}</td><td className="p-3">{x.payer}</td></tr>)}</tbody>
            </table>
          </div>
        )}
      </section>

      <section className="rounded-[2rem] bg-white p-5 shadow-soft">
        <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <div>
            <h3 className="flex items-center gap-2 text-xl font-black"><BarChart3 className="text-blue-600" /> 支出紀錄</h3>
            <p className="mt-1 text-sm text-slate-500">可於每筆支出的 PDF 欄上傳收據 PDF；上傳後可預覽及下載，檔名會根據收據編號產生。</p>
          </div>
          <div className="flex flex-col gap-2 md:flex-row">
            <input className={inputClass()} placeholder="搜尋收據、描述、支付者" value={expenseQuery} onChange={e => setExpenseQuery(e.target.value)} />
            <select className={inputClass()} value={catFilter} onChange={e => setCatFilter(e.target.value)}><option>全部</option>{settings.expenseCategories.map(c => <option key={c}>{c}</option>)}</select>
          </div>
        </div>

        {!!categoryStats.length && (
          <div className="mb-4 grid gap-2 md:grid-cols-4">
            {categoryStats.map(([cat, val]) => <MiniMoney key={cat} label={cat} value={currency(val)} />)}
          </div>
        )}

        <div className="table-wrap">
          <table className="w-full min-w-[880px] text-left text-sm">
            <thead><tr className="border-b bg-slate-50">{['收據編號','日期','類別','支出描述','金額','支付者','PDF','操作'].map(h => <th className="p-3" key={h}>{h}</th>)}</tr></thead>
            <tbody>
              {filteredExpenses.map(x => (
                <tr key={x.id} className="border-b align-top">
                  {editingId === x.id ? (
                    <>
                      <td className="p-2"><input className={inputClass()} value={draft.receiptNo} onChange={e => setDraft({ ...draft, receiptNo: e.target.value })} /></td>
                      <td className="p-2"><input type="date" className={inputClass()} value={draft.date} onChange={e => setDraft({ ...draft, date: e.target.value })} /></td>
                      <td className="p-2"><select className={inputClass()} value={draft.category} onChange={e => setDraft({ ...draft, category: e.target.value })}>{settings.expenseCategories.map(c => <option key={c}>{c}</option>)}</select></td>
                      <td className="p-2"><input className={inputClass()} value={draft.description} onChange={e => setDraft({ ...draft, description: e.target.value })} /></td>
                      <td className="p-2"><input type="number" className={inputClass()} value={draft.amount} onChange={e => setDraft({ ...draft, amount: e.target.value })} /></td>
                      <td className="p-2"><input className={inputClass()} value={draft.payer} onChange={e => setDraft({ ...draft, payer: e.target.value })} /></td>
                      <td className="p-2 text-sm text-slate-500">PDF 可在完成編輯後上傳或檢視</td>
                      <td className="p-2"><button onClick={saveEdit} className="btn-primary"><Save size={16}/>儲存</button></td>
                    </>
                  ) : (
                    <>
                      <td className="p-3 font-bold">{x.receiptNo}</td>
                      <td className="p-3">{x.date}</td>
                      <td className="p-3">{x.category}</td>
                      <td className="p-3">{x.description}</td>
                      <td className="p-3 font-bold">{currency(x.amount)}</td>
                      <td className="p-3">{x.payer}</td>
                      <td className="p-3">
                        <div className="flex flex-wrap gap-2">
                          <label className="btn-muted cursor-pointer">
                            <Upload size={16} /> 上傳
                            <input
                              type="file"
                              accept="application/pdf"
                              className="hidden"
                              onChange={(e) => uploadReceiptPdf(x.id, e.target.files?.[0])}
                            />
                          </label>
                          {x.pdfData ? (
                            <>
                              <button onClick={() => previewReceiptPdf(x)} className="btn-primary"><Eye size={16} /> PDF</button>
                              <button onClick={() => downloadReceiptPdf(x)} className="btn-muted"><Download size={16} /> 下載</button>
                            </>
                          ) : (
                            <span className="rounded-2xl bg-slate-100 px-3 py-2 text-xs font-bold text-slate-500">未上傳</span>
                          )}
                        </div>
                        {x.pdfData && <p className="mt-2 text-xs text-slate-500">檔名：{safeFileName(x.receiptNo)}.pdf</p>}
                      </td>
                      <td className="p-3">
                        <div className="flex gap-2">
                          <button onClick={() => startEdit(x)} className="btn-muted"><Pencil size={16} /> 編輯</button>
                          <button onClick={() => deleteExpense(x.id)} className="btn-danger"><Trash2 size={16} /> 刪除</button>
                        </div>
                      </td>
                    </>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {!filteredExpenses.length && <Empty text="未有支出紀錄。" />}
      </section>
    </div>
  )
}

function SettingsPage({ settings, setSettings, activities, setActivities, requirePassword, flash }) {
  const [password, setPassword] = useState('')
  const [activityCats, setActivityCats] = useState(settings.activityCategories.join('\n'))
  const [expenseCats, setExpenseCats] = useState(settings.expenseCategories.join('\n'))

  async function savePassword() {
    if (password.length < 4) {
      flash('密碼至少需要 4 個字元')
      return
    }
    const passwordHash = await sha256(password)
    setSettings({ ...settings, passwordHash })
    setPassword('')
    flash('密碼已儲存')
  }

  function saveSettings() {
    setSettings({
      ...settings,
      activityCategories: activityCats.split('\n').map(x => x.trim()).filter(Boolean),
      expenseCategories: expenseCats.split('\n').map(x => x.trim()).filter(Boolean)
    })
    flash('設定已儲存')
  }

  function backup() {
    downloadText(`活動支出追蹤系統_備份_${today()}.json`, JSON.stringify({ activities, settings }, null, 2))
  }

  async function restore(file) {
    const text = await file.text()
    try {
      const data = JSON.parse(text)
      await requirePassword(() => {
        if (!window.confirm('還原資料會覆蓋現有資料，確定繼續？')) return false
        setActivities(Array.isArray(data.activities) ? data.activities : [])
        setSettings({ ...defaultSettings, ...(data.settings || {}) })
        flash('資料已還原')
        return true
      })
    } catch {
      flash('JSON 檔案格式不正確')
    }
  }

  async function clearAll() {
    await requirePassword(() => {
      if (!window.confirm('確定清空所有活動資料？此操作不能復原。')) return false
      setActivities([])
      flash('所有活動資料已清空')
      return true
    })
  }

  function loadDemo() {
    setActivities([
      {
        id: uid(),
        activityName: '樂齡同行情緒支援小組',
        activityCode: 'JA-2026-001',
        startDate: '2026-06-01',
        endDate: '2026-06-30',
        budget: 5000,
        category: '小組活動',
        accountingCode: 'AC-2026-001',
        expenseCategory: '物資',
        personInCharge: '陳大文',
        advanceApplied: '有',
        advanceReceipts: [
          {
            id: uid(),
            expectedDate: '2026-05-25',
            amount: 2000,
            imageData: '',
            imageName: '',
            imageUploadedAt: '',
            createdAt: new Date().toISOString()
          },
          {
            id: uid(),
            expectedDate: '2026-06-10',
            amount: 1500,
            imageData: '',
            imageName: '',
            imageUploadedAt: '',
            createdAt: new Date().toISOString()
          }
        ],
        expenses: [
          { id: uid(), receiptNo: 'R001', date: '2026-06-01', category: '物資', description: '文具及活動材料', amount: 238.5, payer: '陳大文' },
          { id: uid(), receiptNo: 'R002', date: '2026-06-03', category: '膳食', description: '參加者茶點', amount: 520, payer: '李小明' }
        ],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      }
    ])
    flash('已加入測試資料')
  }

  return (
    <div className="space-y-5">
      <HeaderBlock title="設定" subtitle="密碼、類別、系統名稱、資料備份及還原。" />

      <section className="rounded-[2rem] bg-white p-5 shadow-soft">
        <h3 className="mb-4 text-xl font-black">系統名稱設定</h3>
        <input className={inputClass()} value={settings.systemName} onChange={e => setSettings({ ...settings, systemName: e.target.value })} />
      </section>

      <section className="rounded-[2rem] bg-white p-5 shadow-soft">
        <h3 className="mb-2 text-xl font-black">密碼管理</h3>
        <p className="mb-4 text-sm text-slate-500">{settings.passwordHash ? '已設定密碼。可在此輸入新密碼更改。' : '首次使用需設置密碼才能啟動刪除功能。設定後才可刪除活動、清空資料或還原資料。'}</p>
        <div className="flex flex-col gap-2 md:flex-row">
          <input type="password" className={inputClass()} placeholder="輸入新密碼" value={password} onChange={e => setPassword(e.target.value)} />
          <button onClick={savePassword} className="btn-primary whitespace-nowrap"><ShieldCheck size={16} /> 儲存密碼</button>
        </div>
      </section>

      <section className="grid gap-4 md:grid-cols-2">
        <div className="rounded-[2rem] bg-white p-5 shadow-soft">
          <h3 className="mb-3 text-xl font-black">預設活動類別</h3>
          <textarea className={`${inputClass()} min-h-44`} value={activityCats} onChange={e => setActivityCats(e.target.value)} />
        </div>
        <div className="rounded-[2rem] bg-white p-5 shadow-soft">
          <h3 className="mb-3 text-xl font-black">預設支出類別</h3>
          <textarea className={`${inputClass()} min-h-44`} value={expenseCats} onChange={e => setExpenseCats(e.target.value)} />
        </div>
      </section>

      <button onClick={saveSettings} className="btn-primary"><Save size={16} /> 儲存設定</button>

      <section className="rounded-[2rem] bg-white p-5 shadow-soft">
        <h3 className="mb-4 text-xl font-black">資料備份及還原</h3>
        <div className="flex flex-wrap gap-2">
          <button onClick={backup} className="btn-primary"><Download size={16} /> 一鍵備份 JSON</button>
          <label className="btn-muted cursor-pointer"><Upload size={16} /> 還原 JSON<input type="file" accept="application/json" className="hidden" onChange={e => e.target.files?.[0] && restore(e.target.files[0])} /></label>
          <button onClick={loadDemo} className="btn-muted"><RotateCcw size={16} /> 加入測試資料</button>
          <button onClick={clearAll} className="btn-danger"><Trash2 size={16} /> 清空所有資料</button>
        </div>
      </section>
    </div>
  )
}

function HeaderBlock({ title, subtitle }) {
  return (
    <div className="rounded-[2rem] bg-slate-950 p-6 text-white shadow-soft">
      <p className="text-xs font-bold uppercase tracking-widest text-blue-300">活動支出追蹤系統</p>
      <h2 className="mt-1 text-3xl font-black">{title}</h2>
      <p className="mt-2 text-sm text-slate-300">{subtitle}</p>
    </div>
  )
}

function MiniMoney({ label, value, danger = false }) {
  return (
    <div className={`rounded-2xl p-3 ${danger ? 'bg-rose-50 text-rose-700' : 'bg-slate-50 text-slate-700'}`}>
      <p className="text-xs font-bold opacity-70">{label}</p>
      <p className="mt-1 font-black">{value}</p>
    </div>
  )
}

function Empty({ text }) {
  return <div className="rounded-3xl border border-dashed border-slate-300 p-8 text-center font-semibold text-slate-500">{text}</div>
}

export default App
