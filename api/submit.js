const { google } = require('googleapis')
const nodemailer = require('nodemailer')

const HEADERS = [
  'Submitted At', 'Full Name', 'Email', 'Gender', 'Age', 'University', 'Faculty', 'Year', 'District', 'Living Situation',
  'Income: Mahapola (LKR)', 'Income: Bursary (LKR)', 'Income: Family (LKR)',
  'Income: Part-time (LKR)', 'Income: Other (LKR)', 'Total Income (LKR)',
  'Primary Income Source',
  'Expense: Rent (LKR)', 'Expense: Food (LKR)', 'Expense: Transport (LKR)',
  'Expense: Academic (LKR)', 'Expense: Mobile (LKR)', 'Expense: Personal (LKR)',
  'Expense: Entertainment (LKR)', 'Expense: Other (LKR)', 'Total Expenses (LKR)',
  'Net (Income - Expenses) (LKR)',
  'Runs Out of Money', 'Saves Monthly', 'Biggest Challenges', 'Comments',
]

async function getSheet() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n').replace(/\\\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  })
  const sheets = google.sheets({ version: 'v4', auth })
  const spreadsheetId = process.env.GOOGLE_SHEET_ID
  const tab = process.env.GOOGLE_SHEET_TAB || 'Responses'

  // Ensure header row exists
  const existing = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${tab}!A1:A1`,
  }).catch(() => null)

  if (!existing?.data?.values?.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${tab}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] },
    })
  }

  return { sheets, spreadsheetId, tab }
}

async function appendRow(data) {
  const { sheets, spreadsheetId, tab } = await getSheet()

  const inc = data.income || {}
  const exp = data.expenses || {}
  const totalInc = (inc.mahapola||0) + (inc.bursary||0) + (inc.family||0) + (inc.parttime||0) + (inc.other||0)
  const totalExp = (exp.rent||0) + (exp.food||0) + (exp.transport||0) + (exp.academic||0) + (exp.mobile||0) + (exp.personal||0) + (exp.entertainment||0) + (exp.other||0)

  const row = [
    new Date().toLocaleString('en-GB', { timeZone: 'Asia/Colombo' }),
    data.full_name || '', data.email || '',
    data.gender, data.age, data.university, data.faculty, data.year, data.district, data.living,
    inc.mahapola||0, inc.bursary||0, inc.family||0, inc.parttime||0, inc.other||0, totalInc,
    data.primary_income,
    exp.rent||0, exp.food||0, exp.transport||0, exp.academic||0, exp.mobile||0, exp.personal||0, exp.entertainment||0, exp.other||0, totalExp,
    totalInc - totalExp,
    data.run_out_frequency, data.saves,
    (data.challenges || []).join('; '),
    data.comments || '',
  ]

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${tab}!A1`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] },
  })

  return { totalInc, totalExp }
}

async function sendNotification(data, totalInc, totalExp) {
  if (!process.env.NOTIFY_FROM_EMAIL || !process.env.NOTIFY_FROM_PASSWORD) return

  const transport = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: process.env.NOTIFY_FROM_EMAIL,
      pass: process.env.NOTIFY_FROM_PASSWORD,
    },
  })

  const inc = data.income || {}
  const html = `
<div style="font-family:sans-serif;max-width:520px;margin:0 auto;">
  <div style="background:#1A6B49;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">
    <h2 style="margin:0;font-size:18px;">📋 New Survey Response</h2>
    <p style="margin:6px 0 0;font-size:13px;opacity:.85;">iLstable Student Income Survey</p>
  </div>
  <div style="border:1px solid #e2e8f0;border-top:none;padding:20px 24px;border-radius:0 0 8px 8px;background:#fff;">
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      <tr><td style="padding:6px 0;color:#718096;width:160px;">Name</td><td style="color:#1a202c;font-weight:600;">${data.full_name || '—'}</td></tr>
      <tr><td style="padding:6px 0;color:#718096;">Email</td><td style="color:#1a202c;">${data.email || '—'}</td></tr>
      <tr><td style="padding:6px 0;color:#718096;">University</td><td style="color:#1a202c;font-weight:600;">${data.university}</td></tr>
      <tr><td style="padding:6px 0;color:#718096;">Faculty / Year</td><td style="color:#1a202c;">${data.faculty} · ${data.year}</td></tr>
      <tr><td style="padding:6px 0;color:#718096;">District</td><td style="color:#1a202c;">${data.district}</td></tr>
      <tr><td style="padding:6px 0;color:#718096;">Living</td><td style="color:#1a202c;">${data.living}</td></tr>
      <tr><td style="padding:6px 0;color:#718096;">Primary Income</td><td style="color:#1a202c;">${data.primary_income}</td></tr>
      <tr style="background:#f7fafc;"><td style="padding:8px 6px;font-weight:700;color:#1A6B49;">Total Monthly Income</td><td style="padding:8px 6px;font-weight:700;color:#1A6B49;">LKR ${totalInc.toLocaleString()}</td></tr>
      <tr style="background:#fff7ed;"><td style="padding:8px 6px;font-weight:700;color:#d97706;">Total Monthly Expenses</td><td style="padding:8px 6px;font-weight:700;color:#d97706;">LKR ${totalExp.toLocaleString()}</td></tr>
      <tr style="background:${(totalInc-totalExp)>=0?'#f0fdf4':'#fff5f5'};"><td style="padding:8px 6px;font-weight:700;color:${(totalInc-totalExp)>=0?'#16a34a':'#e53e3e'};">Net (Surplus / Deficit)</td><td style="padding:8px 6px;font-weight:700;color:${(totalInc-totalExp)>=0?'#16a34a':'#e53e3e'};">LKR ${(totalInc-totalExp).toLocaleString()}</td></tr>
      <tr><td style="padding:6px 0;color:#718096;">Runs Out of Money</td><td style="color:#1a202c;">${data.run_out_frequency}</td></tr>
      <tr><td style="padding:6px 0;color:#718096;">Saves Monthly</td><td style="color:#1a202c;">${data.saves}</td></tr>
    </table>
    ${data.comments ? `<div style="margin-top:16px;padding:12px;background:#f7fafc;border-radius:6px;font-size:13px;color:#4a5568;">"${data.comments}"</div>` : ''}
    <p style="margin-top:20px;font-size:12px;color:#a0aec0;">View all responses in your <a href="https://docs.google.com/spreadsheets/d/${process.env.GOOGLE_SHEET_ID}" style="color:#1A6B49;">Google Sheet →</a></p>
  </div>
</div>`

  await transport.sendMail({
    from: `"iLstable Survey" <${process.env.NOTIFY_FROM_EMAIL}>`,
    to: process.env.NOTIFY_TO_EMAIL,
    subject: `📋 New Survey Response — ${data.university} (${data.year})`,
    html,
  })
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')

  if (req.method === 'OPTIONS') return res.status(200).end()
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' })

  try {
    const data = typeof req.body === 'string' ? JSON.parse(req.body) : req.body

    const { totalInc, totalExp } = await appendRow(data)
    await sendNotification(data, totalInc, totalExp)

    return res.status(200).json({ ok: true, message: 'Response recorded. Thank you!' })
  } catch (err) {
    console.error('Survey submit error:', err?.message || err)
    return res.status(500).json({ ok: false, error: err?.message || 'Failed to save response. Please try again.' })
  }
}
