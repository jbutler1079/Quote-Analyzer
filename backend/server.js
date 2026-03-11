'use strict';

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');
const PptxGenJS = require('pptxgenjs');
const OpenAI = require('openai');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const COVERAGE_TIERS = ['ee', 'es', 'ec', 'ef'];
const PREMIUM_FIELD_BY_TIER = { ee: 'premiumEE', es: 'premiumES', ec: 'premiumEC', ef: 'premiumEF' };
const PAY_PERIODS_PER_MONTH = {
  weekly: 52 / 12,
  biweekly: 26 / 12,
  semimonthly: 24 / 12,
  monthly: 1,
};

// ── In-memory store ──────────────────────────────────────────────────────────
const caseStore = new Map(); // caseId → { files, plans, census, recommendations }
let lastExtractDebug = null; // Stores last extraction debug info

function defaultContributionConfig() {
  return {
    payrollFrequency: 'biweekly',
    tiers: {
      ee: { type: 'percent', value: 0 },
      es: { type: 'percent', value: 0 },
      ec: { type: 'percent', value: 0 },
      ef: { type: 'percent', value: 0 },
    },
  };
}

function normalizeContributionConfig(input) {
  const defaults = defaultContributionConfig();
  const freq = input && typeof input.payrollFrequency === 'string' ? input.payrollFrequency.toLowerCase() : defaults.payrollFrequency;
  const payrollFrequency = Object.prototype.hasOwnProperty.call(PAY_PERIODS_PER_MONTH, freq) ? freq : defaults.payrollFrequency;

  const tiers = {};
  for (const tier of COVERAGE_TIERS) {
    const rawTier = input && input.tiers ? input.tiers[tier] : null;
    const type = rawTier && rawTier.type === 'dollar' ? 'dollar' : 'percent';
    const rawValue = rawTier ? Number(rawTier.value) : 0;
    const value = Number.isFinite(rawValue) && rawValue > 0 ? rawValue : 0;
    tiers[tier] = { type, value };
  }

  return { payrollFrequency, tiers };
}

function calculatePlanCostShares(plan, census, contribution) {
  const normalized = normalizeContributionConfig(contribution);
  const payPeriodsPerMonth = PAY_PERIODS_PER_MONTH[normalized.payrollFrequency] || PAY_PERIODS_PER_MONTH.biweekly;

  // Get EE base premium and EE rule for dependent surplus methodology
  const eePremium = Number(plan[PREMIUM_FIELD_BY_TIER.ee]) || 0;
  const eeRule = normalized.tiers.ee || { type: 'percent', value: 0 };

  // Calculate how much employer covers at EE level (the base)
  let eeBaseEmployer = 0;
  if (eePremium > 0 && eeRule.value > 0) {
    if (eeRule.type === 'dollar') {
      eeBaseEmployer = Math.min(eePremium, eeRule.value);
    } else {
      eeBaseEmployer = eePremium * (Math.max(0, Math.min(100, eeRule.value)) / 100);
    }
  }

  const byTier = {};
  let employerMonthlyTotal = 0;
  let employeeMonthlyTotal = 0;

  for (const tier of COVERAGE_TIERS) {
    const premiumField = PREMIUM_FIELD_BY_TIER[tier];
    const premiumPerMemberMonthly = Number(plan[premiumField]) || 0;
    const enrolled = Number(census[tier]) || 0;
    const rule = normalized.tiers[tier] || { type: 'percent', value: 0 };

    let employerPerMemberMonthly = 0;
    if (premiumPerMemberMonthly > 0) {
      if (tier === 'ee') {
        // EE tier: straightforward — apply EE rule to EE premium
        employerPerMemberMonthly = eeBaseEmployer;
      } else {
        // Dependent tiers: employer pays EE base + tier% of dependent surplus
        // Dependent surplus = tier premium - EE premium (the additional cost for dependents)
        const dependentSurplus = Math.max(0, premiumPerMemberMonthly - eePremium);
        let surplusContribution = 0;
        if (rule.value > 0 && dependentSurplus > 0) {
          if (rule.type === 'dollar') {
            surplusContribution = Math.min(dependentSurplus, rule.value);
          } else {
            surplusContribution = dependentSurplus * (Math.max(0, Math.min(100, rule.value)) / 100);
          }
        }
        employerPerMemberMonthly = Math.min(premiumPerMemberMonthly, eeBaseEmployer + surplusContribution);
      }
    }

    const employeePerMemberMonthly = Math.max(0, premiumPerMemberMonthly - employerPerMemberMonthly);
    const employerMonthly = employerPerMemberMonthly * enrolled;
    const employeeMonthly = employeePerMemberMonthly * enrolled;

    employerMonthlyTotal += employerMonthly;
    employeeMonthlyTotal += employeeMonthly;

    byTier[tier] = {
      enrolled,
      premiumPerMemberMonthly,
      employerPerMemberMonthly,
      employeePerMemberMonthly,
      employerMonthly,
      employeeMonthly,
      employerPerPay: employerMonthly / payPeriodsPerMonth,
      employeePerPay: employeeMonthly / payPeriodsPerMonth,
    };
  }

  return {
    payrollFrequency: normalized.payrollFrequency,
    payPeriodsPerMonth,
    employerMonthlyTotal,
    employeeMonthlyTotal,
    employerPerPayTotal: employerMonthlyTotal / payPeriodsPerMonth,
    employeePerPayTotal: employeeMonthlyTotal / payPeriodsPerMonth,
    byTier,
  };
}

// ── Multer (memory storage) ──────────────────────────────────────────────────
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const allowed = [
      'application/pdf',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'text/csv',
      'application/octet-stream',
    ];
    const ext = file.originalname.split('.').pop().toLowerCase();
    if (allowed.includes(file.mimetype) || ['pdf', 'xlsx', 'xls', 'csv'].includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error(`Unsupported file type: ${file.mimetype}`));
    }
  },
});

// ── Auth middleware (disabled – internal tool) ───────────────────────────────

// ── Health ────────────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ status: 'ok', version: 'llm-strategy13-v1', strategies: 13, llmEnabled: !!process.env.OPENAI_API_KEY }));

// ── Debug: last extraction info ───────────────────────────────────────────────
app.get('/debug/lastextract', (_req, res) => {
  if (!lastExtractDebug) return res.status(404).json({ error: 'No extraction yet' });
  res.json(lastExtractDebug);
});

// ── Debug: extract plans from raw text (for testing without PDF) ──────────────
app.post('/debug/extract', async (req, res) => {
  try {
    const { text, sourceFile } = req.body;
    if (!text) return res.status(400).json({ error: 'text required' });
    const plans = await extractPlanFromText(text, sourceFile || 'debug-input.txt');
    res.json({ plans, count: plans.length });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── Debug: see raw parsed text from uploaded PDF ──────────────────────────────
app.post('/debug/parsetext', async (req, res) => {
  try {
    const { caseId } = req.body;
    if (!caseId) return res.status(400).json({ error: 'caseId required' });
    const caseData = caseStore.get(caseId);
    if (!caseData) return res.status(404).json({ error: 'Case not found' });
    const results = [];
    for (const file of caseData.files) {
      const ext = file.originalname.split('.').pop().toLowerCase();
      if (ext === 'pdf') {
        const data = await pdfParse(file.buffer);
        results.push({ file: file.originalname, textLength: data.text.length, text: data.text });
      }
    }
    res.json({ caseId, results });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── Debug: upload a PDF and get raw text + per-page text ──────────────────────
app.post('/debug/pdftext', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'file required' });
    const data = await pdfParse(req.file.buffer);
    const pages = data.text.split(/\f/);
    res.json({
      file: req.file.originalname,
      totalLength: data.text.length,
      numPages: pages.length,
      pages: pages.map((p, i) => ({ page: i + 1, length: p.length, text: p })),
      fullText: data.text.substring(0, 20000),
    });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── Debug: dump raw text from latest case ─────────────────────────────────────
app.get('/debug/latestcase', async (req, res) => {
  try {
    const entries = Array.from(caseStore.entries());
    if (entries.length === 0) return res.status(404).json({ error: 'No cases in store' });
    const [caseId, caseData] = entries[entries.length - 1];
    const full = req.query.full === '1';
    const results = [];
    for (const file of caseData.files) {
      const ext = file.originalname.split('.').pop().toLowerCase();
      if (ext === 'pdf') {
        const data = await pdfParse(file.buffer);
        const pages = data.text.split(/\f/);
        results.push({
          file: file.originalname,
          totalLength: data.text.length,
          numPages: pages.length,
          pages: pages.map((p, i) => ({
            page: i + 1,
            length: p.length,
            ...(full ? { text: p } : { preview: p.substring(0, 1500) }),
          })),
        });
      }
    }
    res.json({ caseId, numFiles: caseData.files.length, results });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ── Frontend static hosting ───────────────────────────────────────────────────
const frontendDir = path.join(__dirname, '..', 'frontend');
app.use(express.static(frontendDir));

app.get('/', (_req, res) => {
  res.sendFile(path.join(frontendDir, 'index.html'));
});

// ── Upload ────────────────────────────────────────────────────────────────────
app.post('/upload', upload.array('files[]', 20), (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }
    const caseId = uuidv4();
    const files = req.files.map(f => ({
      originalname: f.originalname,
      mimetype: f.mimetype,
      size: f.size,
      buffer: f.buffer,
    }));
    caseStore.set(caseId, {
      files,
      plans: [],
      census: {},
      contribution: defaultContributionConfig(),
      recommendations: null,
    });
    res.json({
      caseId,
      files: files.map(f => ({ name: f.originalname, size: f.size, type: f.mimetype })),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Parse helpers ─────────────────────────────────────────────────────────────

function parseMoney(str) {
  if (!str) return null;
  const cleaned = String(str).replace(/[$,\s]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

function firstMatch(text, patterns) {
  for (const re of patterns) {
    const m = text.match(re);
    if (m) return m[1] ? m[1].trim() : m[0].trim();
  }
  return null;
}

function allMatches(text, patterns) {
  const results = [];
  for (const re of patterns) {
    const flags = re.flags.includes('g') ? re.flags : re.flags + 'g';
    const gre = new RegExp(re.source, flags);
    let m;
    while ((m = gre.exec(text)) !== null) {
      results.push(m[1] ? m[1].trim() : m[0].trim());
    }
  }
  return results;
}

// ── Known carrier names for detection ─────────────────────────────────────────
const KNOWN_CARRIERS = [
  'Anthem', 'Aetna', 'Cigna', 'United\\s*Health(?:care)?', 'UHC', 'Kaiser',
  'Blue\\s*Cross\\s*Blue\\s*Shield', 'BlueCross\\s*BlueShield', 'BCBS',
  'Blue\\s*Cross', 'Blue\\s*Shield', 'Humana', 'Molina', 'Oscar', 'Centene',
  'Wellmark', 'Harvard\\s*Pilgrim', 'Tufts', 'HCSC', 'Premera', 'Regence',
  'Providence', 'Health\\s*Net', 'HealthNet', 'Coventry', 'WellCare',
  'Magellan', 'Ambetter', 'CareFirst', 'Highmark', 'Florida\\s*Blue',
  'Excellus', 'Independence', 'Medica', 'Priority\\s*Health', 'SelectHealth',
  'Allina', 'HealthPartners', 'Dean\\s*Health', 'Geisinger', 'MVP',
  'ConnectiCare', 'EmblemHealth', 'Oxford', 'AvMed',
  'Baylor\\s*Scott\\s*(?:&|and)\\s*White', 'BSW(?:HP)?'
];
const CARRIER_PATTERN = new RegExp('\\b(' + KNOWN_CARRIERS.join('|') + ')\\b', 'i');

// ── Strategy 13: LLM Universal Extractor ──────────────────────────────────────
// Uses OpenAI (GPT-4o-mini by default) to extract plan data from ANY carrier PDF
// format. Only runs when OPENAI_API_KEY environment variable is set.
async function extractWithLLM(text, sourceFile) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) return null; // No key = skip LLM strategy

  const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';
  const client = new OpenAI({ apiKey, timeout: 55000 });

  // Truncate to fit context window (reserve ~4K for prompt + response structure)
  // gpt-4o-mini: 128K context; gpt-4o: 128K context
  const MAX_TEXT_CHARS = 100000;
  const inputText = text.length > MAX_TEXT_CHARS
    ? text.substring(0, MAX_TEXT_CHARS) + '\n\n[... truncated ...]'
    : text;

  const systemPrompt = `You are an insurance benefits extraction engine. Extract ALL health insurance plan options from the provided document text.

For EACH plan, return a JSON object with these exact fields (use null if not found):
- planName: string — Plan display name (e.g., "Gold PPO 1500", "Choice Plus HSA")
- planCode: string — Internal plan code if present (e.g., "EI2J", "AAG1")
- carrier: string — Insurance carrier (e.g., "UnitedHealthcare", "Aetna", "Blue Cross Blue Shield")
- networkType: string — HMO, PPO, EPO, POS, HDHP, or null
- deductibleIndividual: number — In-network individual deductible in dollars
- deductibleFamily: number — In-network family deductible in dollars
- oopMaxIndividual: number — In-network individual out-of-pocket maximum in dollars
- oopMaxFamily: number — In-network family out-of-pocket maximum in dollars
- coinsurance: string — Coinsurance percentage (e.g., "20%")
- copayPCP: number — Primary care copay in dollars
- copaySpecialist: number — Specialist copay in dollars
- copayUrgentCare: number — Urgent care copay in dollars (null if not listed)
- copayER: number — Emergency room copay in dollars
- rxTier1: string — Generic Rx cost (e.g., "$10" or "$10/$20/$30")
- rxTier2: string — Preferred brand Rx cost
- rxTier3: string — Non-preferred Rx cost
- premiumEE: number — Employee-only monthly premium in dollars
- premiumES: number — Employee + Spouse monthly premium in dollars
- premiumEC: number — Employee + Child(ren) monthly premium in dollars
- premiumEF: number — Employee + Family monthly premium in dollars
- effectiveDate: string — Plan effective date if found (e.g., "05/01/2026")
- hsaEligible: boolean — Whether plan is HSA-eligible
- product: string — Product line/family if stated (e.g., "Insurance Choice", "Navigate")

CRITICAL RULES:
- Extract EVERY plan option in the document, including alternates and variations
- Premiums must be monthly amounts. If the document shows per-pay-period amounts, note that but still report as-is.
- Dollar amounts should be numbers without $ sign (e.g., 1500 not "$1,500")
- Return ONLY the JSON array, no explanatory text
- If the document contains plan comparison tables, extract each column as a plan
- If the document contains "Medical Plan Alternates" pages, extract every alternate plan listed`;

  const userPrompt = `Extract all insurance plans from this document (source file: "${sourceFile}"):\n\n${inputText}`;

  console.log(`[LLM] Sending ${inputText.length} chars to ${model} for extraction...`);
  const startTime = Date.now();

  try {
    const completion = await client.chat.completions.create({
      model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt },
      ],
      response_format: { type: 'json_object' },
      temperature: 0,
      max_tokens: 16000,
    });

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    const content = completion.choices[0]?.message?.content;
    if (!content) {
      console.log(`[LLM] Empty response after ${elapsed}s`);
      return [];
    }

    console.log(`[LLM] Received response in ${elapsed}s (${content.length} chars)`);

    const parsed = JSON.parse(content);
    // The response might be { plans: [...] } or just [...]
    const rawPlans = Array.isArray(parsed) ? parsed : (parsed.plans || parsed.data || []);

    if (!Array.isArray(rawPlans) || rawPlans.length === 0) {
      console.log(`[LLM] No plans in response`);
      return [];
    }

    // Normalize LLM output into our standard plan schema
    const plans = rawPlans.map(p => ({
      id: uuidv4(),
      carrier: p.carrier || null,
      planName: p.planName || p.plan_name || 'Unknown Plan',
      planCode: p.planCode || p.plan_code || null,
      networkType: p.networkType || p.network_type || null,
      metalLevel: p.metalLevel || p.metal_level || null,
      deductibleIndividual: toNum(p.deductibleIndividual ?? p.deductible_individual),
      deductibleFamily: toNum(p.deductibleFamily ?? p.deductible_family),
      oopMaxIndividual: toNum(p.oopMaxIndividual ?? p.oop_max_individual),
      oopMaxFamily: toNum(p.oopMaxFamily ?? p.oop_max_family),
      coinsurance: p.coinsurance || null,
      copayPCP: toNum(p.copayPCP ?? p.copay_pcp),
      copaySpecialist: toNum(p.copaySpecialist ?? p.copay_specialist),
      copayUrgentCare: toNum(p.copayUrgentCare ?? p.copay_urgent_care),
      copayER: toNum(p.copayER ?? p.copay_er),
      rxDeductible: toNum(p.rxDeductible ?? p.rx_deductible),
      rxTier1: p.rxTier1 || p.rx_tier1 || null,
      rxTier2: p.rxTier2 || p.rx_tier2 || null,
      rxTier3: p.rxTier3 || p.rx_tier3 || null,
      premiumEE: toNum(p.premiumEE ?? p.premium_ee),
      premiumES: toNum(p.premiumES ?? p.premium_es),
      premiumEC: toNum(p.premiumEC ?? p.premium_ec),
      premiumEF: toNum(p.premiumEF ?? p.premium_ef),
      effectiveDate: p.effectiveDate || p.effective_date || null,
      product: p.product || null,
      hsaEligible: p.hsaEligible ?? p.hsa_eligible ?? false,
      ratingArea: null,
      underwritingNotes: null,
      extractionConfidence: 0,
      sourceFile,
    }));

    // Calculate confidence per plan
    const BENEFIT_FIELDS = [
      'deductibleIndividual', 'deductibleFamily', 'oopMaxIndividual', 'oopMaxFamily',
      'copayPCP', 'copaySpecialist', 'copayER',
      'premiumEE', 'premiumES', 'premiumEC', 'premiumEF',
    ];
    for (const plan of plans) {
      const found = BENEFIT_FIELDS.filter(f => plan[f] != null).length;
      plan.extractionConfidence = Math.min(1, 0.3 + (found * 0.07));
    }

    const usage = completion.usage;
    console.log(`[LLM] Extracted ${plans.length} plans via ${model} in ${elapsed}s` +
      (usage ? ` (${usage.prompt_tokens}+${usage.completion_tokens} tokens)` : ''));

    return plans;
  } catch (err) {
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    console.error(`[LLM] Error after ${elapsed}s: ${err.message}`);
    return [];
  }
}

// Helper: coerce to number, return null if invalid
function toNum(v) {
  if (v == null) return null;
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const n = parseFloat(v.replace(/[$,]/g, ''));
    return isNaN(n) ? null : n;
  }
  return null;
}

async function extractPlanFromText(text, sourceFile) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const fullText = text;

  console.log(`[EXTRACT] Starting extraction for "${sourceFile}" — ${lines.length} non-empty lines`);

  // ── Quality scoring: pick the strategy that produces best data, not just most plans ──
  // Uses AVERAGE benefit fields per plan so 3 complete plans beat 6 half-empty stubs.
  function scoreStrategyResult(plans) {
    if (!plans || plans.length === 0) return 0;
    const BENEFIT_FIELDS = [
      'deductibleIndividual', 'deductibleFamily',
      'oopMaxIndividual', 'oopMaxFamily',
      'copayPCP', 'copaySpecialist', 'copayER',
      'premiumEE', 'premiumES', 'premiumEC', 'premiumEF',
    ];
    let totalBenefitFields = 0;
    let plansWithBenefits = 0;
    for (const p of plans) {
      const benefitCount = BENEFIT_FIELDS.filter(f => p[f] !== null && p[f] !== undefined).length;
      totalBenefitFields += benefitCount;
      if (benefitCount >= 2) plansWithBenefits++;
    }
    // Average benefit fields per plan × number of plans that have real data
    const avgBenefits = totalBenefitFields / plans.length;
    // Reward: more plans WITH actual benefit data, penalize empty stubs
    return avgBenefits * plansWithBenefits;
  }

  // Collect results from ALL strategies
  const candidates = [];

  // ── Strategy 1: Find all distinct plan names in text, split into blocks ──
  try {
    const planNamePlans = extractByPlanNames(text, lines, sourceFile);
    if (planNamePlans.length > 0) {
      const score = scoreStrategyResult(planNamePlans);
      console.log(`[EXTRACT] Strategy 1 (plan names): ${planNamePlans.length} plans, score=${score}`);
      candidates.push({ name: 'plan names', plans: planNamePlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 1 error: ${e.message}`); }

  // ── Strategy 2: Benefit comparison grid ─────────────────────────────────
  try {
    const gridPlans = extractFromBenefitGrid(text, sourceFile);
    if (gridPlans.length > 0) {
      const score = scoreStrategyResult(gridPlans);
      console.log(`[EXTRACT] Strategy 2 (benefit grid): ${gridPlans.length} plans, score=${score}`);
      candidates.push({ name: 'benefit grid', plans: gridPlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 2 error: ${e.message}`); }

  // ── Strategy 3: Table-based extraction (rate sheets) ────────────────────
  try {
    const tablePlans = extractFromTable(lines, fullText, sourceFile);
    if (tablePlans.length > 0) {
      const score = scoreStrategyResult(tablePlans);
      console.log(`[EXTRACT] Strategy 3 (rate table): ${tablePlans.length} plans, score=${score}`);
      candidates.push({ name: 'rate table', plans: tablePlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 3 error: ${e.message}`); }

  // ── Strategy 4: Page-based splitting ────────────────────────────────────
  try {
    const pages = text.split(/\f/).filter(p => p.trim().length > 50);
    if (pages.length > 1) {
      const pagePlans = [];
      for (const page of pages) {
        const plan = extractFieldsFromBlock(page, sourceFile);
        if (plan) pagePlans.push(plan);
      }
      if (pagePlans.length > 0) {
        const score = scoreStrategyResult(pagePlans);
        console.log(`[EXTRACT] Strategy 4 (pages): ${pagePlans.length} plans from ${pages.length} pages, score=${score}`);
        candidates.push({ name: 'pages', plans: pagePlans, score });
      }
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 4 error: ${e.message}`); }

  // ── Strategy 5: Plan-block-based extraction (SBC / summary layouts) ─────
  try {
    const planBlocks = splitIntoPlanBlocks(lines, fullText);
    if (planBlocks.length > 1) {
      const blockPlans = [];
      for (const block of planBlocks) {
        const plan = extractFieldsFromBlock(block, sourceFile);
        if (plan) blockPlans.push(plan);
      }
      if (blockPlans.length > 0) {
        const score = scoreStrategyResult(blockPlans);
        console.log(`[EXTRACT] Strategy 5 (blocks): ${blockPlans.length} plans from ${planBlocks.length} blocks, score=${score}`);
        candidates.push({ name: 'blocks', plans: blockPlans, score });
      }
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 5 error: ${e.message}`); }

  // ── Strategy 6: Repeated benefit keywords ─────────────────────────────────
  try {
    const repPlans = extractByRepeatedKeywords(text, lines, sourceFile);
    if (repPlans.length > 0) {
      const score = scoreStrategyResult(repPlans);
      console.log(`[EXTRACT] Strategy 6 (repeated keywords): ${repPlans.length} plans, score=${score}`);
      candidates.push({ name: 'repeated keywords', plans: repPlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 6 error: ${e.message}`); }

  // ── Strategy 7: Plan code benefit rows (BSW-style) ────────────────────────
  try {
    const codeRowPlans = extractFromPlanCodeBenefitRows(text, sourceFile);
    if (codeRowPlans.length > 0) {
      const score = scoreStrategyResult(codeRowPlans);
      console.log(`[EXTRACT] Strategy 7 (plan code rows): ${codeRowPlans.length} plans, score=${score}`);
      candidates.push({ name: 'plan code benefit rows', plans: codeRowPlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 7 error: ${e.message}`); }

  // ── Strategy 8: Aetna Medical Cost Grid ─────────────────────────────────
  try {
    const aetnaPlans = extractFromAetnaCostGrid(text, sourceFile);
    if (aetnaPlans.length > 0) {
      const score = scoreStrategyResult(aetnaPlans);
      console.log(`[EXTRACT] Strategy 8 (Aetna cost grid): ${aetnaPlans.length} plans, score=${score}`);
      candidates.push({ name: 'Aetna cost grid', plans: aetnaPlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 8 error: ${e.message}`); }

  // ── Strategy 9: BCBS Proposal Grid ──────────────────────────────────────
  try {
    const bcbsPlans = extractFromBCBSGrid(text, sourceFile);
    if (bcbsPlans.length > 0) {
      const score = scoreStrategyResult(bcbsPlans);
      console.log(`[EXTRACT] Strategy 9 (BCBS grid): ${bcbsPlans.length} plans, score=${score}`);
      candidates.push({ name: 'BCBS proposal grid', plans: bcbsPlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 9 error: ${e.message}`); }

  // ── Strategy 10: BSW Vertical Format ─────────────────────────────────────
  try {
    const bswPlans = extractFromBSWVertical(text, sourceFile);
    if (bswPlans.length > 0) {
      const score = scoreStrategyResult(bswPlans);
      console.log(`[EXTRACT] Strategy 10 (BSW vertical): ${bswPlans.length} plans, score=${score}`);
      candidates.push({ name: 'BSW vertical format', plans: bswPlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 10 error: ${e.message}`); }

  // ── Strategy 11: BSW Rate Sheet (concatenated tabular rows) ──────────────
  try {
    const bswRatePlans = extractFromBSWRateSheet(text, sourceFile);
    if (bswRatePlans.length > 0) {
      const score = scoreStrategyResult(bswRatePlans);
      console.log(`[EXTRACT] Strategy 11 (BSW rate sheet): ${bswRatePlans.length} plans, score=${score}`);
      candidates.push({ name: 'BSW rate sheet', plans: bswRatePlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 11 error: ${e.message}`); }

  // ── Strategy 12: UHC / UnitedHealthcare Quote Grid ────────────────────────
  try {
    const uhcPlans = extractFromUHCQuote(text, sourceFile);
    if (uhcPlans.length > 0) {
      // Small tiebreaker bonus for carrier-specific strategies (same as Aetna/BCBS/BSW)
      const score = scoreStrategyResult(uhcPlans) + 0.1;
      console.log(`[EXTRACT] Strategy 12 (UHC quote): ${uhcPlans.length} plans, score=${score}`);
      candidates.push({ name: 'UHC quote grid', plans: uhcPlans, score });
    }
  } catch (e) { console.log(`[EXTRACT] Strategy 12 error: ${e.message}`); }

  // ── Strategy 13: LLM Universal Extractor ──────────────────────────────────
  // Smart fallback: only calls the LLM when regex strategies produced poor results.
  // This avoids adding 10-30s of latency when regex already found good data.
  let llmStatus = 'no API key';
  try {
    if (process.env.OPENAI_API_KEY) {
      const bestRegexScore = candidates.length > 0
        ? Math.max(...candidates.map(c => c.score))
        : 0;
      const bestRegexPlans = candidates.length > 0
        ? Math.max(...candidates.map(c => c.plans.length))
        : 0;
      // Only call LLM if: no plans found, or best regex score is weak (< 20 = few plans with few fields)
      if (bestRegexPlans === 0 || bestRegexScore < 20) {
        console.log(`[LLM] Regex best: ${bestRegexPlans} plans, score=${bestRegexScore.toFixed(1)} — invoking LLM fallback`);
        llmStatus = 'invoked';
        const llmPlans = await extractWithLLM(text, sourceFile);
        if (llmPlans && llmPlans.length > 0) {
          const score = scoreStrategyResult(llmPlans) + 0.1;
          console.log(`[EXTRACT] Strategy 13 (LLM universal): ${llmPlans.length} plans, score=${score}`);
          candidates.push({ name: 'LLM universal', plans: llmPlans, score });
          llmStatus = `success: ${llmPlans.length} plans`;
        } else {
          llmStatus = 'invoked but returned 0 plans';
        }
      } else {
        llmStatus = `skipped (regex score=${bestRegexScore.toFixed(1)})`;
        console.log(`[LLM] Regex found ${bestRegexPlans} plans with score=${bestRegexScore.toFixed(1)} — skipping LLM`);
      }
    }
  } catch (e) {
    llmStatus = `error: ${e.message}`;
    console.log(`[EXTRACT] Strategy 13 error: ${e.message}`);
  }

  // ── Pick the best strategy by quality score ─────────────────────────────
  let plans = [];
  if (candidates.length > 0) {
    candidates.sort((a, b) => b.score - a.score);
    const best = candidates[0];
    console.log(`[EXTRACT] Winner: "${best.name}" with ${best.plans.length} plans, score=${best.score}`);
    console.log(`[EXTRACT] All candidates:`, candidates.map(c => `${c.name}: ${c.plans.length} plans, score=${c.score}`));
    plans = best.plans;

    // Store debug info for last extraction
    lastExtractDebug = {
      timestamp: new Date().toISOString(),
      sourceFile,
      winner: best.name,
      winnerPlans: best.plans.length,
      winnerScore: best.score,
      allCandidates: candidates.map(c => ({ name: c.name, plans: c.plans.length, score: c.score })),
      planNames: best.plans.map(p => p.planName),
      llmStatus,
    };
  }

  // ── Strategy 7: Fallback — entire text as one block ─────────────────────
  if (plans.length === 0) {
    const plan = extractFieldsFromBlock(lines.join('\n'), sourceFile);
    if (plan) {
      console.log(`[EXTRACT] Strategy 7 (fallback): extracted 1 plan from entire text`);
      plans.push(plan);
    }
  }

  // ── Post-extraction: clean plan names ───────────────────────────────────
  for (const plan of plans) {
    plan.planName = cleanPlanName(plan.planName);
  }

  // ── Post-extraction: enrich plans with missing benefit fields ───────────
  enrichPlansFromText(plans, fullText);

  // ── Post-extraction: apply carrier from full text if not found per-plan ─
  const globalCarrier = detectCarrier(fullText, sourceFile);
  for (const plan of plans) {
    if (!plan.carrier && globalCarrier) plan.carrier = globalCarrier;
  }

  // ── Post-extraction: apply effective date from full text if not per-plan ─
  const hasDateGap = plans.some(p => !p.effectiveDate);
  if (hasDateGap) {
    const effDateGlobal = firstMatch(fullText, [
      /effective\s*(?:date)?\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i,
      /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})\s*(?:effective|start)/i,
    ]);
    if (effDateGlobal) {
      for (const plan of plans) {
        if (!plan.effectiveDate) plan.effectiveDate = effDateGlobal;
      }
    }
  }

  // De-duplicate plans that have the same planName + carrier
  plans = deduplicatePlans(plans);

  console.log(`[EXTRACT] Final result: ${plans.length} plans`);
  return plans;
}

// ── Clean plan name: strip headers, embedded newlines, common prefixes ────────
function cleanPlanName(name) {
  if (!name) return name;
  // Remove "Illustrative Quote" header prefix (appears in BCBS TX PDFs)
  let cleaned = name.replace(/^(?:Illustrative\s*)?Quote\s*/i, '');
  // Collapse newlines and multiple spaces
  cleaned = cleaned.replace(/[\r\n]+/g, ' ').replace(/\s{2,}/g, ' ').trim();

  // Remove carrier boilerplate prefix ("BlueCross BlueShield of Texas", "Anthem", etc.)
  // Only if there's a meaningful plan descriptor after it
  const carrierPrefixRe = new RegExp('^(?:' + KNOWN_CARRIERS.join('|') + ')\\s+(?:of\\s+\\w+\\s+)?', 'i');
  const withoutCarrier = cleaned.replace(carrierPrefixRe, '').trim();
  if (withoutCarrier.length > 5 && /[A-Z]/i.test(withoutCarrier)) {
    cleaned = withoutCarrier;
  }

  // Note: Do NOT strip BCBS plan-ID suffixes (e.g., P9M1CHC, G654CHC, B661ADT)
  // since they are the only way to differentiate plans within the same network+metal tier.

  return cleaned || name;
}

// ── Strategy 1: Extract by finding all plan names, then building context ──────
// Scans text for all distinct plan-name-like strings, then for each plan name,
// looks within a context window for benefit values.
function extractByPlanNames(text, lines, sourceFile) {
  // Find all plan name occurrences in the text
  const planNamePatterns = [
    /([A-Z][A-Za-z0-9\s\/\-&']{2,50}?(?:HMO|PPO|EPO|HDHP|HSA)\s*(?:Gold|Silver|Bronze|Platinum)?[A-Za-z0-9\s\/\-&']{0,20})/g,
    /([A-Z][A-Za-z0-9\s\/\-&']{2,50}?(?:Gold|Silver|Bronze|Platinum)\s*(?:HMO|PPO|EPO|HDHP|HSA)?[A-Za-z0-9\s\/\-&']{0,20})/g,
    // Also catch carrier-prefixed plan names with plan numbers/codes
    new RegExp('(' + KNOWN_CARRIERS.join('|') + ')\\s+[A-Za-z0-9\\s\\/\\-&\']{3,50}', 'gi'),
  ];

  // Collect all candidate plan names with their line positions
  const candidates = [];
  for (let i = 0; i < lines.length; i++) {
    for (const pattern of planNamePatterns) {
      pattern.lastIndex = 0;
      let m;
      while ((m = pattern.exec(lines[i])) !== null) {
        const name = m[1].trim();
        // Filter out noise: skip if too short, too long, or looks like a label
        if (name.length < 6 || name.length > 80) continue;
        if (/^(deductible|copay|premium|coinsurance|out.of.pocket|oop|maximum|benefit|coverage|plan\s*type)/i.test(name)) continue;
        candidates.push({ name, lineIndex: i, text: lines[i] });
      }
    }
  }

  if (candidates.length === 0) return [];

  // Deduplicate plan names — group similar names
  const uniqueNames = [];
  const seen = new Set();
  for (const c of candidates) {
    const normalized = c.name.toLowerCase().replace(/\s+/g, ' ').trim();
    if (seen.has(normalized)) continue;
    // Also check for substring matches
    let isDupe = false;
    for (const s of seen) {
      if (s.includes(normalized) || normalized.includes(s)) { isDupe = true; break; }
    }
    if (isDupe) continue;
    seen.add(normalized);
    uniqueNames.push(c);
  }

  console.log(`[EXTRACT] Plan name candidates: ${uniqueNames.map(u => u.name).join(' | ')}`);

  if (uniqueNames.length <= 1) return [];

  // For each unique plan name, build a context window and extract fields
  const plans = [];
  for (let u = 0; u < uniqueNames.length; u++) {
    const { name, lineIndex } = uniqueNames[u];
    // Context: from this plan name to the next plan name (or end of text)
    const nextLineIndex = u + 1 < uniqueNames.length ? uniqueNames[u + 1].lineIndex : lines.length;
    // Also look a few lines before the plan name for context
    const startLine = Math.max(0, lineIndex - 2);
    const blockLines = lines.slice(startLine, nextLineIndex);
    const blockText = blockLines.join('\n');

    const plan = extractFieldsFromBlock(blockText, sourceFile);
    if (plan) {
      // Override plan name with the one we found
      if (!plan.planName || plan.planName.length < name.length) {
        plan.planName = name;
      }
      plans.push(plan);
    }
  }

  return plans;
}

// ── Strategy 7: Plan code benefit rows ────────────────────────────────────────
// Handles PDFs where each line is a plan-code followed by concatenated benefit
// data.  Common in BSW (Baylor Scott & White) and similar carriers that produce
// benefit-comparison tables with columns: Code | Deductible | PCP/Specialist |
// Coinsurance | OOP Max.  PDF extraction concatenates the cells, giving lines
// like: PHG26P44$1,000$5 / $2020% copayment after deductible$3,600

/**
 * Split concatenated specialist-copay + coinsurance digits.
 * E.g. "2020" → { spec: 20, coins: 20 }, "250" → { spec: 25, coins: 0 }
 */
function splitSpecCoins(combined) {
  const candidates = [];
  for (let i = 1; i < combined.length; i++) {
    const spec = parseInt(combined.slice(0, i), 10);
    const coins = parseInt(combined.slice(i), 10);
    if (isNaN(spec) || isNaN(coins)) continue;
    if (coins >= 0 && coins <= 50 && coins % 5 === 0 && spec >= 0 && spec <= 200) {
      candidates.push({ spec, coins });
    }
  }
  if (candidates.length === 0) return null;
  if (candidates.length === 1) return candidates[0];
  // Prefer candidates with specialist in typical range ($5-$150), pick largest specialist
  const typical = candidates.filter(c => c.spec >= 5 && c.spec <= 150);
  if (typical.length > 0) {
    typical.sort((a, b) => b.spec - a.spec);
    return typical[0];
  }
  candidates.sort((a, b) => b.spec - a.spec);
  return candidates[0];
}

function extractFromPlanCodeBenefitRows(text, sourceFile) {
  const lines = text.split('\n');
  const plans = [];

  // BSW-style codes: [Metal][Net]G[YY][SubNet][Num]
  const PLAN_CODE_RE = /^([PGBS][HP]G\d{2}[A-Z](?:\d{2,3}|IV))/;

  const METAL_MAP = { P: 'Platinum', G: 'Gold', S: 'Silver', B: 'Bronze' };
  const NET_MAP   = { H: 'HMO', P: 'PPO' };
  const SUBNET_MAP = { P: 'BSW Premier', A: 'BSW Access', D: 'BSW Plus' };

  // Detect carrier
  let carrier = null;
  if (/baylor\s*scott\s*(?:&|and)\s*white/i.test(text)) carrier = 'Baylor Scott & White';

  // Regex for premium line: $EE$ES$EC$EF$Total (5 dollar amounts with decimals)
  const PREMIUM_LINE_RE = /^\$([\d,]+\.\d{2})\$([\d,]+\.\d{2})\$([\d,]+\.\d{2})\$([\d,]+\.\d{2})\$([\d,]+\.\d{2})$/;
  // Regex for Rx tier line: $3/$50/$125/$250 etc
  const RX_LINE_RE = /^\$(\d+)\/\$(\d+)\/\$(\d+)\/\$(\d+)$/;

  /**
   * Look ahead from the plan code line to find Rx tiers and premiums.
   * In the BSW format, the sequence after a plan code line is:
   *   [optional Rx line like "$3/$50/$125/$250"]
   *   [optional "0% copayment after" / "deductible" for HSA]
   *   [premium line like "$825.02$1,650.04$1,650.04$2,475.06$7,425.14"]
   */
  function lookAheadForPremiums(startIdx) {
    const result = { premiumEE: null, premiumES: null, premiumEC: null, premiumEF: null,
                     rxTier1: null, rxTier2: null, rxTier3: null };
    // Search the next 5 lines for premium and Rx data
    for (let j = startIdx + 1; j < Math.min(startIdx + 6, lines.length); j++) {
      const ahead = lines[j].trim();
      if (!ahead) continue;

      // Stop if we hit another plan code line or a section header
      if (PLAN_CODE_RE.test(ahead)) break;

      // Check for Rx tier line
      const rxMatch = RX_LINE_RE.exec(ahead);
      if (rxMatch) {
        result.rxTier1 = parseInt(rxMatch[1], 10);
        result.rxTier2 = parseInt(rxMatch[2], 10);
        result.rxTier3 = parseInt(rxMatch[3], 10);
        continue;
      }

      // Check for premium line (5 dollar amounts with cents)
      const premMatch = PREMIUM_LINE_RE.exec(ahead);
      if (premMatch) {
        result.premiumEE = parseMoney(premMatch[1]);
        result.premiumES = parseMoney(premMatch[2]);
        result.premiumEC = parseMoney(premMatch[3]);
        result.premiumEF = parseMoney(premMatch[4]);
        break; // Found premiums, stop looking
      }

      // Skip "0% copayment after" / "deductible" / "Total Monthly" / "Premium" continuations
      if (/^(0%|deductible|total\s*monthly|premium)/i.test(ahead)) continue;

      // If it's a recognizable section header (e.g. "HMO Gold Plans"), stop
      if (/^(HMO|PPO|BSW|Composite|Group|Plan)\b/i.test(ahead)) break;
    }
    return result;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const codeMatch = PLAN_CODE_RE.exec(line);
    if (!codeMatch) continue;

    const code = codeMatch[1];
    const rest = line.substring(code.length);
    if (!rest.startsWith('$')) continue;

    const metalCode = code[0], netCode = code[1];
    const subMatch = code.match(/G\d{2}([A-Z])/);
    const subNetCode = subMatch ? subMatch[1] : null;
    const metal = METAL_MAP[metalCode] || null;
    const network = NET_MAP[netCode] || null;
    const subNetwork = subNetCode ? (SUBNET_MAP[subNetCode] || null) : null;

    // Look ahead for premiums and Rx
    const extras = lookAheadForPremiums(i);

    // --- HSA format: $DED0% AD / 0% AD0% copayment after deductible$OOP ---
    const hsaRe = /^\$(\d{1,3}(?:,\d{3})*)0%\s*AD\s*\/\s*0%\s*AD\s*0%\s*co(?:payment|insurance)\s+after\s+deductible\$([\d,]+)/;
    const hsaMatch = hsaRe.exec(rest);
    if (hsaMatch) {
      const ded = parseMoney(hsaMatch[1]);
      const oop = parseMoney(hsaMatch[2]);
      const hasPremiums = extras.premiumEE != null;
      const benefitCount = [ded, oop].filter(v => v != null).length + (hasPremiums ? 4 : 0);

      plans.push({
        id: uuidv4(), carrier,
        planName: `${metal} ${network} HSA $${ded.toLocaleString()}${subNetwork ? ' (' + subNetwork + ')' : ''}`,
        planCode: code, networkType: network, metalLevel: metal,
        deductibleIndividual: ded, deductibleFamily: null,
        oopMaxIndividual: oop, oopMaxFamily: null,
        coinsurance: '0% after deductible',
        copayPCP: null, copaySpecialist: null,
        copayUrgentCare: null, copayER: null,
        rxDeductible: null,
        rxTier1: extras.rxTier1, rxTier2: extras.rxTier2, rxTier3: extras.rxTier3,
        premiumEE: extras.premiumEE, premiumES: extras.premiumES,
        premiumEC: extras.premiumEC, premiumEF: extras.premiumEF,
        effectiveDate: null, ratingArea: null, underwritingNotes: null,
        extractionConfidence: Math.min(1, 0.4 + (benefitCount * 0.1)),
        sourceFile,
      });
      continue;
    }

    // --- Standard format: $DED$PCP [AD] / $SPEC_COINS% copay... $OOP ---
    const stdRe = /^\$([\d,]+)\$(\d+)\s*(?:AD\s*)?\/\s*\$(.+?)%\s*co(?:payment|insurance)(?:\s+after\s+deductible)?\$([\d,]+)/;
    const stdMatch = stdRe.exec(rest);
    if (!stdMatch) continue;

    const ded = parseMoney(stdMatch[1]);
    const pcp = parseInt(stdMatch[2], 10);
    const specCoinsRaw = stdMatch[3].trim();
    const oop = parseMoney(stdMatch[4]);

    let spec = null, coins = null;
    const adSplit = specCoinsRaw.match(/^(\d+)\s*AD\s*(\d+)$/);
    if (adSplit) {
      spec = parseInt(adSplit[1], 10);
      coins = parseInt(adSplit[2], 10);
    } else {
      const pureDigits = specCoinsRaw.replace(/\s+/g, '');
      if (/^\d+$/.test(pureDigits)) {
        const result = splitSpecCoins(pureDigits);
        if (result) { spec = result.spec; coins = result.coins; }
      }
    }

    let planName;
    if (ded === 0 && coins != null) {
      planName = `${metal} ${network} Copay $0/${oop != null ? '$' + oop.toLocaleString() + ' OOP' : ''}`;
    } else {
      const planPays = coins != null ? (100 - coins) : '?';
      planName = `${metal} ${network} ${planPays} $${ded.toLocaleString()} Ded`;
    }
    if (subNetwork) planName += ` (${subNetwork})`;

    const hasPremiums = extras.premiumEE != null;
    const benefitFields = [ded, oop, pcp, spec].filter(v => v != null).length + (hasPremiums ? 4 : 0);

    plans.push({
      id: uuidv4(), carrier,
      planName, planCode: code, networkType: network, metalLevel: metal,
      deductibleIndividual: ded, deductibleFamily: null,
      oopMaxIndividual: oop, oopMaxFamily: null,
      coinsurance: coins != null ? `${coins}%` : null,
      copayPCP: pcp, copaySpecialist: spec,
      copayUrgentCare: null, copayER: null,
      rxDeductible: null,
      rxTier1: extras.rxTier1, rxTier2: extras.rxTier2, rxTier3: extras.rxTier3,
      premiumEE: extras.premiumEE, premiumES: extras.premiumES,
      premiumEC: extras.premiumEC, premiumEF: extras.premiumEF,
      effectiveDate: null, ratingArea: null, underwritingNotes: null,
      extractionConfidence: Math.min(1, 0.4 + (benefitFields * 0.075)),
      sourceFile,
    });
  }

  if (plans.length < 1) return [];
  console.log(`[EXTRACT] Plan code rows: found ${plans.length} plans from code-prefixed lines`);
  return plans;
}

// ── Strategy 10: BSW Vertical/Sequential Quote Format ─────────────────────────
// Handles BSW (Baylor Scott & White) PDFs where benefits appear on separate lines
// rather than concatenated after a plan code.  The format is:
//   [Section: HMO Plans / PPO Plans]
//   [Metal: Gold / Silver / Bronze]
//   Plan ID
//   GHG26A01
//   Individual Deductible
//   $1,500
//   Family Deductible
//   $3,000
//   ...premiums...
function extractFromBSWVertical(text, sourceFile) {
  // Gate: must look like a BSW document
  if (!/baylor\s*scott\s*(?:&|and)\s*white|bswhp/i.test(text) &&
      !/BSW\s*(?:Premier|Access|Plus)/i.test(text)) return [];

  const lines = text.split('\n');
  const plans = [];
  const carrier = 'Baylor Scott & White';

  // BSW plan code patterns (flexible)
  const BSW_CODE_RE = /^([PGBS][HP]G\d{2}[A-Z](?:\d{1,3}|IV))$/;
  // Also match standalone plan codes that may have different patterns
  const ALT_CODE_RE = /^([A-Z]{2,4}[\-]?\d{3,6}[A-Z]*)$/;

  const METAL_MAP = { P: 'Platinum', G: 'Gold', S: 'Silver', B: 'Bronze' };
  const NET_MAP   = { H: 'HMO', P: 'PPO' };
  const SUBNET_MAP = { P: 'BSW Premier', A: 'BSW Access', D: 'BSW Plus' };

  // Effective date
  const effMatch = text.match(/Effective\s*Date[:\s]+([\d\/\-]+)/i);
  const effectiveDate = effMatch ? effMatch[1] : null;

  // Track current context from section headers
  let currentNetwork = null;
  let currentMetal = null;
  let isHSA = false;

  // Find all plan code positions first
  const planPositions = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Track section headers for network
    if (/^HMO\s*Plans?$/i.test(line) || /^HMO$/i.test(line)) { currentNetwork = 'HMO'; continue; }
    if (/^PPO\s*Plans?$/i.test(line) || /^PPO$/i.test(line)) { currentNetwork = 'PPO'; continue; }
    if (/^HSA\s*Plans?$/i.test(line)) { isHSA = true; continue; }
    if (/^(?:HMO|PPO)\s*Plans?$/i.test(line)) { isHSA = false; continue; }

    // Track metal level headers
    if (/^Platinum$/i.test(line)) { currentMetal = 'Platinum'; continue; }
    if (/^Gold$/i.test(line))     { currentMetal = 'Gold'; continue; }
    if (/^Silver$/i.test(line))   { currentMetal = 'Silver'; continue; }
    if (/^Bronze$/i.test(line))   { currentMetal = 'Bronze'; continue; }
    if (/^Expanded\s*Bronze$/i.test(line)) { currentMetal = 'Bronze'; continue; }

    // Check if previous line was "Plan ID" label
    const prevLine = i > 0 ? lines[i - 1].trim() : '';
    const isPlanCodeContext = /^Plan\s*(?:ID|Code|Number)$/i.test(prevLine);

    if (BSW_CODE_RE.test(line) || (isPlanCodeContext && ALT_CODE_RE.test(line))) {
      planPositions.push({
        lineIndex: i,
        code: line,
        network: currentNetwork,
        metal: currentMetal,
        isHSA,
      });
    }
  }

  if (planPositions.length === 0) return [];

  // For each plan code, scan the surrounding lines for benefit data
  for (let p = 0; p < planPositions.length; p++) {
    const pos = planPositions[p];
    const code = pos.code;
    const startLine = pos.lineIndex;
    // End at next plan code or end of text
    const endLine = p + 1 < planPositions.length
      ? planPositions[p + 1].lineIndex
      : Math.min(startLine + 60, lines.length);

    // Extract benefit fields from the block between this plan code and the next
    const block = lines.slice(startLine, endLine);
    const blockText = block.join('\n');

    // Parse metal/network from code if not from headers
    let metal = pos.metal;
    let network = pos.network;
    let subNetwork = null;

    const codeM = code.match(/^([PGBS])([HP])G\d{2}([A-Z])/);
    if (codeM) {
      if (!metal) metal = METAL_MAP[codeM[1]] || metal;
      if (!network) network = NET_MAP[codeM[2]] || network;
      subNetwork = SUBNET_MAP[codeM[3]] || null;
    }

    // Scan block for labeled values using label/value pairs (may be on same or adjacent lines)
    let ded = null, dedFam = null, oop = null, oopFam = null;
    let pcp = null, spec = null, urgent = null, er = null;
    let coins = null, rxT1 = null, rxT2 = null, rxT3 = null;
    let premEE = null, premES = null, premEC = null, premEF = null;

    for (let j = 0; j < block.length; j++) {
      const line = block[j].trim();
      const nextLine = j + 1 < block.length ? block[j + 1].trim() : '';

      // Try to extract value from same line or next line
      const valueOnLine = line.match(/\$([\d,]+(?:\.\d{2})?)/);
      const valueNextLine = nextLine.match(/^\$([\d,]+(?:\.\d{2})?)$/);

      // Combined label + value on same line (e.g., "Individual Deductible $1,500")
      const sameLineVal = line.match(/^(.+?)\s+\$([\d,]+(?:\.\d{2})?)\s*$/);

      // Individual Deductible
      if (/^Individual\s*Deductible$/i.test(line) && valueNextLine) {
        ded = parseMoney(valueNextLine[1]); j++;
      } else if (/Individual\s*Deductible/i.test(line) && sameLineVal) {
        ded = parseMoney(sameLineVal[2]);
      } else if (/^(?:In-?Network\s*)?Deductible$/i.test(line) && valueNextLine) {
        ded = parseMoney(valueNextLine[1]); j++;
      }

      // Family Deductible
      if (/^Family\s*Deductible$/i.test(line) && valueNextLine) {
        dedFam = parseMoney(valueNextLine[1]); j++;
      } else if (/Family\s*Deductible/i.test(line) && sameLineVal) {
        dedFam = parseMoney(sameLineVal[2]);
      }

      // OOP Max Individual
      if (/^(?:Individual\s*)?(?:OOP|Out[\s-]*of[\s-]*Pocket)\s*(?:Max(?:imum)?|Limit)?(?:\s*Individual)?$/i.test(line) && valueNextLine) {
        oop = parseMoney(valueNextLine[1]); j++;
      } else if (/(?:OOP|Out[\s-]*of[\s-]*Pocket)\s*(?:Max|Limit)?\s*Individual/i.test(line) && sameLineVal) {
        oop = parseMoney(sameLineVal[2]);
      } else if (/Individual\s*(?:OOP|Out[\s-]*of[\s-]*Pocket)/i.test(line) && sameLineVal) {
        oop = parseMoney(sameLineVal[2]);
      }

      // OOP Max Family
      if (/^(?:Family\s*)?(?:OOP|Out[\s-]*of[\s-]*Pocket)\s*(?:Max(?:imum)?|Limit)?(?:\s*Family)?$/i.test(line) && valueNextLine) {
        oopFam = parseMoney(valueNextLine[1]); j++;
      } else if (/(?:OOP|Out[\s-]*of[\s-]*Pocket)\s*(?:Max|Limit)?\s*Family/i.test(line) && sameLineVal) {
        oopFam = parseMoney(sameLineVal[2]);
      } else if (/Family\s*(?:OOP|Out[\s-]*of[\s-]*Pocket)/i.test(line) && sameLineVal) {
        oopFam = parseMoney(sameLineVal[2]);
      }

      // Coinsurance
      const coinsMatch = line.match(/(\d+)\s*%\s*(?:Coinsurance|co[\s-]*insurance)/i) ||
                          line.match(/Coinsurance\s*(\d+)\s*%/i);
      if (coinsMatch) {
        coins = coinsMatch[1];
      } else if (/^Coinsurance$/i.test(line)) {
        const nextVal = nextLine.match(/^(\d+)\s*%/);
        if (nextVal) { coins = nextVal[1]; j++; }
      }

      // PCP / Primary Care
      if (/^(?:Primary\s*Care|PCP)(?:\s*(?:Office\s*)?(?:Visit|Copay))?$/i.test(line) && valueNextLine) {
        pcp = parseMoney(valueNextLine[1]); j++;
      } else if (/(?:Primary\s*Care|PCP)/i.test(line) && sameLineVal) {
        pcp = parseMoney(sameLineVal[2]);
      }

      // Specialist
      if (/^Specialist(?:\s*(?:Office\s*)?(?:Visit|Copay))?$/i.test(line) && valueNextLine) {
        spec = parseMoney(valueNextLine[1]); j++;
      } else if (/Specialist/i.test(line) && sameLineVal && !/OOP|Deductible|Premium/i.test(line)) {
        spec = parseMoney(sameLineVal[2]);
      }

      // Urgent Care
      if (/^Urgent\s*Care$/i.test(line) && valueNextLine) {
        urgent = parseMoney(valueNextLine[1]); j++;
      } else if (/Urgent\s*Care/i.test(line) && sameLineVal) {
        urgent = parseMoney(sameLineVal[2]);
      }

      // ER / Emergency Room
      if (/^(?:ER|Emergency\s*Room)$/i.test(line) && valueNextLine) {
        er = parseMoney(valueNextLine[1]); j++;
      } else if (/(?:ER|Emergency\s*Room)/i.test(line) && sameLineVal) {
        er = parseMoney(sameLineVal[2]);
      }

      // Rx tiers — "Rx: $10/$40/$80" or "$10/$40/$80" or separate lines
      const rxMatch = line.match(/\$(\d+)\s*\/\s*\$(\d+)\s*\/\s*\$(\d+)/);
      if (rxMatch) {
        rxT1 = parseInt(rxMatch[1], 10);
        rxT2 = parseInt(rxMatch[2], 10);
        rxT3 = parseInt(rxMatch[3], 10);
      }

      // Rx tiers with 4 values: $10/$40/$80/$150
      const rx4Match = line.match(/\$(\d+)\s*\/\s*\$(\d+)\s*\/\s*\$(\d+)\s*\/\s*\$(\d+)/);
      if (rx4Match) {
        rxT1 = parseInt(rx4Match[1], 10);
        rxT2 = parseInt(rx4Match[2], 10);
        rxT3 = parseInt(rx4Match[3], 10);
      }

      // Premiums — label on one line, value on next
      if (/^Employee\s*Only$/i.test(line) && valueNextLine) {
        premEE = parseMoney(valueNextLine[1]); j++;
      } else if (/Employee\s*Only/i.test(line) && sameLineVal) {
        premEE = parseMoney(sameLineVal[2]);
      } else if (/^(?:EE|Single)$/i.test(line) && valueNextLine) {
        premEE = parseMoney(valueNextLine[1]); j++;
      }

      if (/^Employee\s*\+\s*Spouse$/i.test(line) && valueNextLine) {
        premES = parseMoney(valueNextLine[1]); j++;
      } else if (/Employee\s*\+\s*Spouse/i.test(line) && sameLineVal) {
        premES = parseMoney(sameLineVal[2]);
      } else if (/^(?:ES|Emp\s*\+\s*Sp)$/i.test(line) && valueNextLine) {
        premES = parseMoney(valueNextLine[1]); j++;
      }

      if (/^Employee\s*\+\s*Child(?:ren|\(ren\))?$/i.test(line) && valueNextLine) {
        premEC = parseMoney(valueNextLine[1]); j++;
      } else if (/Employee\s*\+\s*Child/i.test(line) && sameLineVal) {
        premEC = parseMoney(sameLineVal[2]);
      } else if (/^(?:EC|Emp\s*\+\s*Ch)$/i.test(line) && valueNextLine) {
        premEC = parseMoney(valueNextLine[1]); j++;
      }

      if (/^Family$/i.test(line) && valueNextLine && premEE != null) {
        // Only match "Family" as premium if we already found EE premium (avoids matching "Family Deductible")
        premEF = parseMoney(valueNextLine[1]); j++;
      } else if (/^Family\b/i.test(line) && !/Deductible|OOP|Out.of.Pocket/i.test(line) && sameLineVal && premEE != null) {
        premEF = parseMoney(sameLineVal[2]);
      } else if (/^Employee\s*\+\s*Family$/i.test(line) && valueNextLine) {
        premEF = parseMoney(valueNextLine[1]); j++;
      } else if (/Employee\s*\+\s*Family/i.test(line) && sameLineVal) {
        premEF = parseMoney(sameLineVal[2]);
      } else if (/^(?:EF|Emp\s*\+\s*Fam)$/i.test(line) && valueNextLine) {
        premEF = parseMoney(valueNextLine[1]); j++;
      }

      // Handle "Deductible / OOP" combo line: "$1,500 / $6,000" type
      const comboMatch = line.match(/^\$(\d[\d,]*)\s*\/\s*\$(\d[\d,]*)$/);
      if (comboMatch && !ded && !oop) {
        ded = parseMoney(comboMatch[1]);
        oop = parseMoney(comboMatch[2]);
      }
    }

    // Build plan name
    let planName;
    const metalStr = metal || 'Plan';
    const netStr = network || '';
    if (pos.isHSA) {
      planName = `${metalStr} ${netStr} HSA`.trim();
    } else if (ded != null && ded === 0) {
      planName = `${metalStr} ${netStr} Copay`.trim();
    } else if (ded != null && coins != null) {
      const planPays = 100 - parseInt(coins, 10);
      planName = `${metalStr} ${netStr} ${planPays} $${ded.toLocaleString()} Ded`.trim();
    } else if (ded != null) {
      planName = `${metalStr} ${netStr} $${ded.toLocaleString()} Ded`.trim();
    } else {
      planName = `${metalStr} ${netStr}`.trim();
    }
    if (subNetwork) planName += ` (${subNetwork})`;

    const BENEFIT_FIELDS = [ded, dedFam, oop, oopFam, pcp, spec, urgent, er, premEE, premES, premEC, premEF];
    const benefitCount = BENEFIT_FIELDS.filter(v => v != null).length;

    // Only add plan if we found at least some real data
    if (benefitCount < 2) continue;

    plans.push({
      id: uuidv4(), carrier,
      planName, planCode: code, networkType: network, metalLevel: metal,
      deductibleIndividual: ded, deductibleFamily: dedFam,
      oopMaxIndividual: oop, oopMaxFamily: oopFam,
      coinsurance: coins ? `${coins}%` : null,
      copayPCP: pcp, copaySpecialist: spec,
      copayUrgentCare: urgent, copayER: er,
      rxDeductible: null,
      rxTier1: rxT1, rxTier2: rxT2, rxTier3: rxT3,
      premiumEE: premEE, premiumES: premES,
      premiumEC: premEC, premiumEF: premEF,
      effectiveDate, ratingArea: null, underwritingNotes: null,
      extractionConfidence: Math.min(1, 0.3 + (benefitCount * 0.06)),
      sourceFile,
    });
  }

  if (plans.length === 0) return [];
  console.log(`[EXTRACT] BSW vertical: found ${plans.length} plans`);
  return plans;
}

// ── Strategy 11: BSW Rate Sheet (concatenated tabular rows) ───────────────────
// Handles BSW renewal/proposal rate sheets where each plan is a single row with
// concatenated columns:
//   [Code][BSW Network][ProductType][Coins%][$Ded/$OOP][$PCP/$Spec][RxCode][RxTiers]
// followed by a premium line:
//   $EE$ES$EC$EF
// Example:
//   LC6QB2L2BSW Premier HMOCC8020%$5000/$7000$25/$50LGRXHE26$5/$15/$60/$130
//   $634.50$1,395.90$1,142.10$2,030.39
function extractFromBSWRateSheet(text, sourceFile) {
  // Gate: must contain a BSW network name in a concatenated data context
  if (!/BSW\s+(?:Premier|Access|Plus)\s+(?:HMO|PPO|EPO)/i.test(text)) return [];

  const lines = text.split('\n');
  const plans = [];
  const carrier = 'Baylor Scott & White';

  // Effective date
  let effectiveDate = null;
  const effMatch1 = text.match(/(?:Renewal\s+)?Effective(?:\s+Date)?(?:\s+of)?\s*[:\s]\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
  const effMatch2 = text.match(/(\w+\s+\d{1,2},?\s*\d{4})\s+to\s+/i);
  if (effMatch1) effectiveDate = effMatch1[1];
  else if (effMatch2) effectiveDate = effMatch2[1];

  // Plan data line: [planCode][BSW network]...
  const BSW_PLAN_RE = /^([A-Z0-9]{4,16})(BSW\s+(?:Premier|Access|Plus)\s+(?:HMO|PPO|EPO))\s*/i;

  // Product type + coinsurance: CC80→20%, CC100 HDHP→0%, etc.
  // Use alternation so CC100 is tried before \d{2} to prevent CC10 + 0... mis-parse
  const PRODUCT_RE = /^(CC(?:100|\d{2}))(\s*HDHP)?\s*(\d{1,2})%\s*/;

  // Ded/OOP: $5000/$7000
  const DED_OOP_RE = /^\$([\d,]+)\/\$([\d,]+)\s*/;

  // Copay: $25/$50
  const COPAY_RE = /^\$(\d+)\/\$(\d+)\s*/;

  // "Ded + 0%" or "Ded +0%" style (HDHP copay placeholder)
  const DED_COPAY_RE = /^Ded\s*\+\s*\d+%\s*/i;

  // Premium line: 4+ consecutive $amounts with cents
  const PREM_4_RE = /^\$([\d,]+\.\d{2})\$([\d,]+\.\d{2})\$([\d,]+\.\d{2})\$([\d,]+\.\d{2})/;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const planMatch = BSW_PLAN_RE.exec(line);
    if (!planMatch) continue;

    const planCode = planMatch[1];
    const networkRaw = planMatch[2].trim();
    let rest = line.slice(planMatch[0].length);

    // Parse product type and coinsurance
    const prodMatch = PRODUCT_RE.exec(rest);
    if (!prodMatch) continue;

    const productLabel = prodMatch[1] + (prodMatch[2] || '').trim();
    const coinsurance = parseInt(prodMatch[3], 10);
    const isHDHP = /HDHP/i.test(productLabel);
    rest = rest.slice(prodMatch[0].length);

    // Parse deductible / OOP max
    const dedMatch = DED_OOP_RE.exec(rest);
    if (!dedMatch) continue;

    const deductible = parseMoney(dedMatch[1]);
    const oopMax = parseMoney(dedMatch[2]);
    rest = rest.slice(dedMatch[0].length);

    // Parse PCP/Spec copays or "Ded + X%"
    let pcp = null, spec = null;
    const copayMatch = COPAY_RE.exec(rest);
    const dedCopayMatch = DED_COPAY_RE.exec(rest);

    if (copayMatch) {
      pcp = parseInt(copayMatch[1], 10);
      spec = parseInt(copayMatch[2], 10);
      rest = rest.slice(copayMatch[0].length);
    } else if (dedCopayMatch) {
      rest = rest.slice(dedCopayMatch[0].length);
    }

    // Parse Rx tiers — skip Rx plan code (alphanumeric), find $X/$Y/$Z or $X/$Y/$Z/$W
    let rxT1 = null, rxT2 = null, rxT3 = null;
    const rx4Match = rest.match(/\$(\d+)\/\$(\d+)\/\$(\d+)\/\$(\d+)/);
    const rx3Match = rest.match(/\$(\d+)\/\$(\d+)\/\$(\d+)/);

    if (rx4Match) {
      rxT1 = parseInt(rx4Match[1], 10);
      rxT2 = parseInt(rx4Match[2], 10);
      rxT3 = parseInt(rx4Match[3], 10);
    } else if (rx3Match) {
      rxT1 = parseInt(rx3Match[1], 10);
      rxT2 = parseInt(rx3Match[2], 10);
      rxT3 = parseInt(rx3Match[3], 10);
    }

    // Look ahead for premium line (next non-empty line)
    let premEE = null, premES = null, premEC = null, premEF = null;
    for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
      const nextLine = lines[j].trim();
      if (!nextLine) continue;
      if (BSW_PLAN_RE.test(nextLine)) break;

      const premMatch = PREM_4_RE.exec(nextLine);
      if (premMatch) {
        premEE = parseMoney(premMatch[1]);
        premES = parseMoney(premMatch[2]);
        premEC = parseMoney(premMatch[3]);
        premEF = parseMoney(premMatch[4]);
        break;
      }
    }

    // Network type
    let networkType = null;
    if (/HMO/i.test(networkRaw)) networkType = 'HMO';
    else if (/PPO/i.test(networkRaw)) networkType = 'PPO';
    else if (/EPO/i.test(networkRaw)) networkType = 'EPO';

    // Sub-network
    let subNetwork = null;
    if (/Premier/i.test(networkRaw)) subNetwork = 'BSW Premier';
    else if (/Access/i.test(networkRaw)) subNetwork = 'BSW Access';
    else if (/Plus/i.test(networkRaw)) subNetwork = 'BSW Plus';

    // Metal level estimate from plan design
    let metalLevel = null;
    if (isHDHP) {
      metalLevel = deductible >= 6000 ? 'Bronze' : 'Silver';
    } else {
      if (coinsurance <= 10) metalLevel = 'Platinum';
      else if (coinsurance <= 20) metalLevel = 'Gold';
      else if (coinsurance <= 30) metalLevel = 'Silver';
      else metalLevel = 'Bronze';
    }

    // Build descriptive plan name
    let planName;
    if (isHDHP) {
      planName = `${subNetwork || networkRaw} ${networkType} HDHP $${deductible.toLocaleString()} Ded`;
    } else {
      const planPays = 100 - coinsurance;
      planName = `${subNetwork || networkRaw} ${networkType} ${planPays}/${coinsurance} $${deductible.toLocaleString()} Ded`;
    }

    const benefitFields = [deductible, oopMax, pcp, spec, premEE, premES, premEC, premEF, rxT1].filter(v => v != null).length;

    plans.push({
      id: uuidv4(), carrier,
      planName, planCode, networkType, metalLevel,
      deductibleIndividual: deductible, deductibleFamily: null,
      oopMaxIndividual: oopMax, oopMaxFamily: null,
      coinsurance: `${coinsurance}%`,
      copayPCP: pcp, copaySpecialist: spec,
      copayUrgentCare: null, copayER: null,
      rxDeductible: null,
      rxTier1: rxT1, rxTier2: rxT2, rxTier3: rxT3,
      premiumEE: premEE, premiumES: premES,
      premiumEC: premEC, premiumEF: premEF,
      effectiveDate, ratingArea: null, underwritingNotes: null,
      extractionConfidence: Math.min(1, 0.3 + (benefitFields * 0.08)),
      sourceFile,
    });
  }

  if (plans.length === 0) return [];
  console.log(`[EXTRACT] BSW rate sheet: found ${plans.length} plans`);
  return plans;
}

// ── Strategy 12: UHC / UnitedHealthcare Quote Grid ────────────────────────────
// Handles UHC "Group Specialty" (GS) and standard multi-plan quote PDFs.
// UHC quotes typically have a columnar benefit comparison grid where plan names
// appear as column headers and benefit labels as rows.  pdf-parse often renders
// these as one of:
// UHC "Medical Proposed Rates" format:
//   - Multi-page document, each page has 2-3 "Options"
//   - Pages separated by "UnitedHealthcare\nMedical Proposed Rates for <CLIENT>"
//   - Each page: Option headers, Plan Name codes, Product (INS-Choice etc),
//     benefits (concatenated values), rates (concatenated $amounts)
//   - Rate pages identified by "Option \d+" pattern
function extractFromUHCQuote(text, sourceFile) {
  // Gate: must look like a UHC/UnitedHealthcare document
  if (!/United\s*Health(?:care)?|UHC|\bINS-Choice|\bINS-Surest/i.test(text) &&
      !/UHC/i.test(sourceFile || '')) return [];

  const carrier = 'UnitedHealthcare';

  // ── Effective date ──
  let effectiveDate = null;
  const effM = text.match(/Effective\s*(?:Date)?\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i);
  if (effM) effectiveDate = effM[1];

  // ── Split into page blocks ──
  // Split on "UnitedHealthcare\nMedical Proposed Rates" which starts each page
  const pageBlocks = text.split(/(?=UnitedHealthcare\nMedical Proposed Rates)/);
  const ratePages = pageBlocks.filter(b => /Option\s+\d+/i.test(b.substring(0, 600)));

  if (ratePages.length === 0) return [];
  console.log(`[UHC] Found ${ratePages.length} rate page(s)`);

  const allPlans = [];

  for (const page of ratePages) {
    const lines = page.split('\n').map(l => l.trim());

    // ── Find Option numbers on this page ──
    // e.g. "Option 1Option 2Option 3" or "Option 16Option 17"
    let optionNumbers = [];
    let optionLineIdx = -1;
    for (let i = 0; i < Math.min(15, lines.length); i++) {
      const optMatches = [...lines[i].matchAll(/Option\s+(\d+)/gi)];
      if (optMatches.length >= 2) {
        optionNumbers = optMatches.map(m => parseInt(m[1], 10));
        optionLineIdx = i;
        break;
      }
    }
    if (optionNumbers.length === 0) continue;
    const numPlans = optionNumbers.length;

    // ── Extract plan name codes ──
    // After "Plan Name" line, expect N plan code blocks (may wrap across lines)
    // e.g. "EITP (EPO HSA Emb) Rx plan: \nMM"
    let planNameStartIdx = -1;
    for (let i = optionLineIdx + 1; i < Math.min(optionLineIdx + 5, lines.length); i++) {
      if (/^Plan\s+Name$/i.test(lines[i])) { planNameStartIdx = i; break; }
    }

    const planNames = [];
    if (planNameStartIdx >= 0) {
      // Collect lines between "Plan Name" and "Product" — these contain the plan codes
      let productIdx = -1;
      for (let i = planNameStartIdx + 1; i < Math.min(planNameStartIdx + 20, lines.length); i++) {
        if (/^Product$/i.test(lines[i])) { productIdx = i; break; }
      }
      if (productIdx > planNameStartIdx) {
        // Join all non-empty lines between Plan Name and Product
        const nameLines = lines.slice(planNameStartIdx + 1, productIdx).filter(l => l.length > 0);
        // Plan codes look like "EITP (EPO HSA Emb) Rx plan:" or "EIZ4 (EPO PROformance) Rx\nplan: Z9"
        // They start with a 4-char alphanumeric code
        const rawText = nameLines.join('\n');
        // Split on plan code pattern: alphanumeric/underscore code followed by " ("
        const codeRe = /(?:^|\n)\s*([A-Z][A-Z0-9_]{1,20})\s*\(/gi;
        const codeParts = [];
        let match;
        while ((match = codeRe.exec(rawText)) !== null) {
          codeParts.push({ code: match[1], pos: match.index });
        }
        if (codeParts.length >= numPlans) {
          for (let p = 0; p < numPlans; p++) {
            const start = codeParts[p].pos;
            const end = p + 1 < codeParts.length ? codeParts[p + 1].pos : rawText.length;
            let nameChunk = rawText.substring(start, end).replace(/\n/g, ' ').trim();
            // Clean trailing whitespace and partial Rx info
            nameChunk = nameChunk.replace(/\s+/g, ' ').replace(/\s*$/, '');
            planNames.push(nameChunk);
          }
        } else {
          // Fallback: just join lines and split evenly
          const joinedNames = nameLines.join(' ').replace(/\s+/g, ' ');
          planNames.push(joinedNames);
        }
      }
    }

    // ── Extract Product line (INS-Choice, HMO-NexusACO R, etc) ──
    const products = [];
    for (let i = optionLineIdx; i < Math.min(optionLineIdx + 25, lines.length); i++) {
      if (/^Product$/i.test(lines[i])) {
        // Next non-empty line has concatenated products
        for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
          const pl = lines[j];
          if (!pl) continue;
          // Split concatenated products: "INS-ChoiceINS-ChoiceINS-Choice"
          // or "INS-Choice +INS-Choice +INS-Choice +"
          const productRe = /((?:INS-[A-Za-z+ ]+|HMO-[A-Za-z+ ]+|PPO-[A-Za-z+ ]+|EPO-[A-Za-z+ ]+|Surest[A-Za-z+ ]*|[A-Z]{3,}-[A-Za-z+ ]+?)(?=(?:INS-|HMO-|PPO-|EPO-|Surest|[A-Z]{3,}-)|$))/gi;
          const pMatches = [...pl.matchAll(productRe)];
          if (pMatches.length >= numPlans) {
            for (const pm of pMatches) products.push(pm[1].trim());
          } else if (pl.length > 3) {
            // Try to split by known product prefixes
            const splits = pl.split(/(INS-|HMO-|PPO-|EPO-)/i).filter(Boolean);
            let current = '';
            for (const s of splits) {
              if (/^(INS-|HMO-|PPO-|EPO-)$/i.test(s)) {
                if (current) products.push(current.trim());
                current = s;
              } else {
                current += s;
              }
            }
            if (current) products.push(current.trim());
          }
          if (products.length > 0) break;
        }
        break;
      }
    }

    // ── Derive network type from plan name/product ──
    function deriveNetwork(planName, product) {
      const combined = (planName + ' ' + product).toLowerCase();
      if (/\bepo\b/.test(combined)) return 'EPO';
      if (/\bhmo\b/.test(combined)) return 'HMO';
      if (/\bppo\b/.test(combined)) return 'PPO';
      if (/\bpos\b/.test(combined)) return 'POS';
      if (/\bhsa\b/.test(combined)) return 'HSA';
      if (/surest/i.test(combined)) return 'Surest';
      return null;
    }

    // ── Extract HRA/HSA line ──
    const hsaFlags = [];
    for (let i = optionLineIdx; i < Math.min(optionLineIdx + 30, lines.length); i++) {
      if (/^HRA\/HSA$/i.test(lines[i])) {
        for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
          const hl = lines[j];
          if (!hl) continue;
          // Concatenated: "HSAHSANo" or "NoHSANo" — no word boundaries
          const hsaRe = /(HSA|HRA|No)/gi;
          const hMatches = [...hl.matchAll(hsaRe)];
          if (hMatches.length >= numPlans) {
            for (const hm of hMatches) hsaFlags.push(hm[1].toUpperCase());
          }
          if (hsaFlags.length > 0) break;
        }
        break;
      }
    }

    // ── Extract Deductible line ──
    // Line: "$6,750 / $13,500 / Emb$5,000 / $10,000 / Emb$5,000 / $10,000 / Emb"
    // Each option has "ind / fam / Emb" pattern separated by concatenation
    function extractDeductibleOOP(label) {
      const results = [];
      for (let i = optionLineIdx; i < lines.length; i++) {
        if (lines[i] === label) {
          // Scan next few lines for dollar amounts
          for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
            const dl = lines[j];
            if (!dl || dl === ' ') continue;
            // Split by "Emb" or "N/A" boundaries to separate option blocks
            // Pattern: "$X / $Y / Emb" repeating, or "$0 / $0 / N/A" 
            const blockRe = /(\$[\d,]+(?:\.\d+)?\s*\/\s*\$[\d,]+(?:\.\d+)?\s*\/\s*(?:Emb|N\/A|Agg))/gi;
            const blocks = [...dl.matchAll(blockRe)];
            if (blocks.length >= numPlans) {
              for (const b of blocks) {
                const amts = [...b[1].matchAll(/\$([\d,]+(?:\.\d+)?)/g)];
                const ind = amts[0] ? parseMoney(amts[0][1]) : null;
                const fam = amts[1] ? parseMoney(amts[1][1]) : null;
                results.push({ individual: ind, family: fam });
              }
              return results;
            }
            // Fallback: just extract all dollar amounts and pair them
            const allAmts = [...dl.matchAll(/\$([\d,]+(?:\.\d+)?)/g)];
            if (allAmts.length >= numPlans * 2) {
              for (let p = 0; p < numPlans; p++) {
                const ind = parseMoney(allAmts[p * 2][1]);
                const fam = parseMoney(allAmts[p * 2 + 1][1]);
                results.push({ individual: ind, family: fam });
              }
              return results;
            }
            // Single value per plan
            if (allAmts.length >= numPlans) {
              for (let p = 0; p < numPlans; p++) {
                results.push({ individual: parseMoney(allAmts[p][1]), family: null });
              }
              return results;
            }
          }
          break;
        }
      }
      return results;
    }

    // Search for "Deductible" or "Out-of-Pocket" under In-Network section only
    // The structure: "In Network" → benefit rows → "Out of Network"
    // Values are concatenated per-option blocks ending in Emb|N/A|Agg:
    //   "$3,000 / $6,000 / Emb$3,500 / $7,000 / EmbN/A / Emb"
    function findInNetworkDeductibleOOP(label) {
      let inNetStart = -1;
      let outNetStart = lines.length;
      for (let i = optionLineIdx; i < lines.length; i++) {
        if (/^In Network/i.test(lines[i]) && inNetStart === -1) inNetStart = i;
        if (/^Out of Network/i.test(lines[i]) && inNetStart >= 0) { outNetStart = i; break; }
      }
      if (inNetStart < 0) inNetStart = optionLineIdx;

      const results = [];
      for (let i = inNetStart; i < outNetStart; i++) {
        if (lines[i] === label) {
          for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
            const dl = lines[j];
            if (!dl || dl === ' ') continue;
            // Split into per-option blocks by splitting on Emb|N/A|Agg terminators
            // Each block: "$X / $Y / Emb" or "N/A / Emb" or "$X / $Y / N/A"
            const blockParts = dl.split(/(?<=Emb|N\/A|Agg)(?=[^\s\/]|$)/i).filter(b => b.trim());
            if (blockParts.length >= numPlans) {
              for (let p = 0; p < numPlans; p++) {
                const block = blockParts[p];
                const amts = [...block.matchAll(/\$([\d,]+(?:\.\d+)?)/g)];
                if (amts.length >= 2) {
                  results.push({ individual: parseMoney(amts[0][1]), family: parseMoney(amts[1][1]) });
                } else if (amts.length === 1) {
                  results.push({ individual: parseMoney(amts[0][1]), family: null });
                } else {
                  results.push({ individual: null, family: null }); // N/A block
                }
              }
              return results;
            }
            // Fallback: all dollar amounts paired (2 per plan)
            const allAmts = [...dl.matchAll(/\$([\d,]+(?:\.\d+)?)/g)];
            if (allAmts.length >= numPlans * 2) {
              for (let p = 0; p < numPlans; p++) {
                results.push({ individual: parseMoney(allAmts[p * 2][1]), family: parseMoney(allAmts[p * 2 + 1][1]) });
              }
              return results;
            }
            if (allAmts.length >= numPlans) {
              for (let p = 0; p < numPlans; p++) {
                results.push({ individual: parseMoney(allAmts[p][1]), family: null });
              }
              return results;
            }
          }
          break;
        }
      }
      return results;
    }

    const deductibles = findInNetworkDeductibleOOP('Deductible');
    const oopValues = findInNetworkDeductibleOOP('Out-of-Pocket');

    // ── Extract copay from "Office Copay (PCP/SPC)" line ──
    // Format: entire copay section spans from "Office Copay (PCP/SPC)" to "Hospital Copays"
    // The label may be merged with the first plan's copay on the same line:
    //   "Office Copay (PCP/SPC)PCP $15, SCP $50/$100"
    //   "PCP Ded+100%, SCP Ded+100%"
    //   "PCP $10, SCP $30"
    const copays = { pcp: [], specialist: [] };
    for (let i = optionLineIdx; i < lines.length; i++) {
      if (/Office\s*Copay.*PCP/i.test(lines[i])) {
        // Include the copay label line itself (text after the label) and subsequent lines
        const afterLabel = lines[i].replace(/^.*Office\s*Copay\s*\([^)]*\)/i, '').trim();
        const copayLines = afterLabel ? [afterLabel] : [];
        for (let j = i + 1; j < Math.min(i + 10, lines.length); j++) {
          if (/^(?:Hospital|UC\/ER|Major|X-ray|Deductible|Coinsurance|Out-of-Pocket|Pharmacy|Enrollment|Rates)/i.test(lines[j])) break;
          if (lines[j] && lines[j] !== ' ') copayLines.push(lines[j]);
        }
        // Join and parse all PCP/SPC copay amounts
        const copayText = copayLines.join(' ');
        // Match "PCP $XX" and "PCP Ded+" patterns per plan
        const pcpDollarMatches = [...copayText.matchAll(/PCP\s+\$(\d+)/gi)];
        const pcpDedMatches = [...copayText.matchAll(/PCP\s+Ded\+/gi)];
        const spcDollarMatches = [...copayText.matchAll(/S(?:CP|PC)\s+\$(\d+)/gi)];
        const spcDedMatches = [...copayText.matchAll(/S(?:CP|PC)\s+Ded\+/gi)];
        // Build per-plan copays in order of appearance
        const pcpAll = [];
        const spcAll = [];
        // Merge dollar and ded matches by position
        const pcpEntries = [
          ...pcpDollarMatches.map(m => ({ pos: m.index, val: parseInt(m[1], 10) })),
          ...pcpDedMatches.map(m => ({ pos: m.index, val: null }))
        ].sort((a, b) => a.pos - b.pos);
        const spcEntries = [
          ...spcDollarMatches.map(m => ({ pos: m.index, val: parseInt(m[1], 10) })),
          ...spcDedMatches.map(m => ({ pos: m.index, val: null }))
        ].sort((a, b) => a.pos - b.pos);
        for (const e of pcpEntries) copays.pcp.push(e.val);
        for (const e of spcEntries) copays.specialist.push(e.val);
        break;
      }
    }

    // ── Extract ER copay from UC/ER line ──
    const erCopays = [];
    for (let i = optionLineIdx; i < lines.length; i++) {
      if (/^UC\/ER$/i.test(lines[i])) {
        const ucerLines = [];
        for (let j = i + 1; j < Math.min(i + 6, lines.length); j++) {
          if (/^(?:Major|X-ray|Deductible|Coinsurance|Out-of-Pocket|Pharmacy|Enrollment|Rates|Office)/i.test(lines[j])) break;
          if (lines[j] && lines[j] !== ' ') ucerLines.push(lines[j]);
        }
        const ucerText = ucerLines.join(' ');
        const erMatches = [...ucerText.matchAll(/ER\s+\$(\d+)/gi)];
        for (const em of erMatches) erCopays.push(parseInt(em[1], 10));
        break;
      }
    }

    // ── Extract Coinsurance ──
    // Format: "100%100%80%" or "75%100%100%" or "100%N/A"
    // Must parse in order, interleaving percentages and N/A
    const coinsurances = [];
    for (let i = optionLineIdx; i < lines.length; i++) {
      if (/^Coinsurance$/i.test(lines[i])) {
        for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
          const cl = lines[j];
          if (!cl || cl === ' ') continue;
          // Split into tokens by matching percentages and N/A in order
          const tokenRe = /(\d+)%|N\/A/gi;
          let tm;
          while ((tm = tokenRe.exec(cl)) !== null) {
            if (tm[1]) coinsurances.push(parseInt(tm[1], 10));
            else coinsurances.push(null); // N/A
          }
          if (coinsurances.length >= numPlans) break;
        }
        break; // Only use first "Coinsurance" (In Network section)
      }
    }

    // ── Extract Rates ──
    // Find "EE Only" line, followed by concatenated rates "$417.17$453.18$480.37"
    const rates = { ee: [], es: [], ec: [], ef: [] };
    const rateLabels = [
      { key: 'ee', re: /^EE\s*Only$/i },
      { key: 'es', re: /^EE\+Spouse\b/i },
      { key: 'ec', re: /^EE\+Ch\(?ren\)?\b/i },
      { key: 'ef', re: /^Family\b/i },
    ];

    // Find the "Rates" section marker
    let ratesStartIdx = -1;
    for (let i = optionLineIdx; i < lines.length; i++) {
      if (/^Rates$/i.test(lines[i])) { ratesStartIdx = i; break; }
    }
    if (ratesStartIdx < 0) ratesStartIdx = Math.floor(lines.length * 0.6); // fallback

    for (const rl of rateLabels) {
      for (let i = ratesStartIdx; i < lines.length; i++) {
        if (rl.re.test(lines[i])) {
          // Look at this line and next few for dollar amounts
          for (let j = i; j < Math.min(i + 3, lines.length); j++) {
            const rateLine = lines[j];
            const amts = [...rateLine.matchAll(/\$([\d,]+(?:\.\d{1,2})?)/g)];
            if (amts.length >= numPlans) {
              for (const a of amts) rates[rl.key].push(parseMoney(a[1]));
              break;
            }
          }
          break;
        }
        // Also handle inline: "EE+Spouse $917.77$996.99$1,056.81"
        if (rl.key !== 'ee' && rl.re.test(lines[i].split(/\$/)[0])) {
          const amts = [...lines[i].matchAll(/\$([\d,]+(?:\.\d{1,2})?)/g)];
          if (amts.length >= numPlans) {
            for (const a of amts) rates[rl.key].push(parseMoney(a[1]));
            break;
          }
        }
      }
    }

    // ── Build plan objects ──
    for (let p = 0; p < numPlans; p++) {
      const optionNum = optionNumbers[p];
      const planName = planNames[p] || `Option ${optionNum}`;
      const product = products[p] || '';
      const networkType = deriveNetwork(planName, product);
      const isHSA = (hsaFlags[p] === 'HSA' || hsaFlags[p] === 'HRA');

      const plan = {
        id: uuidv4(),
        carrier,
        planName: `Option ${optionNum}: ${planName}`,
        planCode: planName.match(/^([A-Z][A-Z0-9_]+)/)?.[1] || null,
        networkType: isHSA ? (networkType || 'HSA') : networkType,
        metalLevel: null,
        deductibleIndividual: deductibles[p]?.individual ?? null,
        deductibleFamily: deductibles[p]?.family ?? null,
        oopMaxIndividual: oopValues[p]?.individual ?? null,
        oopMaxFamily: oopValues[p]?.family ?? null,
        coinsurance: coinsurances[p] != null ? `${coinsurances[p]}%` : null,
        copayPCP: copays.pcp[p] ?? null,
        copaySpecialist: copays.specialist[p] ?? null,
        copayUrgentCare: null,
        copayER: erCopays[p] ?? null,
        rxDeductible: null, rxTier1: null, rxTier2: null, rxTier3: null,
        premiumEE: rates.ee[p] ?? null,
        premiumES: rates.es[p] ?? null,
        premiumEC: rates.ec[p] ?? null,
        premiumEF: rates.ef[p] ?? null,
        effectiveDate,
        product: product || null,
        hsaEligible: isHSA,
        ratingArea: null,
        underwritingNotes: null,
        extractionConfidence: 0,
        sourceFile,
      };

      // Calculate confidence
      const fields = [plan.deductibleIndividual, plan.deductibleFamily,
                      plan.oopMaxIndividual, plan.oopMaxFamily,
                      plan.copayPCP, plan.copaySpecialist, plan.copayER,
                      plan.premiumEE, plan.premiumES, plan.premiumEC, plan.premiumEF];
      const found = fields.filter(v => v != null).length;
      plan.extractionConfidence = Math.min(1, 0.3 + (found * 0.07));

      if (found >= 2) {
        allPlans.push(plan);
      }
    }
  }

  if (allPlans.length > 0) {
    console.log(`[UHC] Extracted ${allPlans.length} plans from ${ratePages.length} rate pages`);
  }

  // ── Parse "Medical Plan Alternates" (MPE) pages ──
  // These are dense tabular pages listing alternate plans for each product line.
  // Format per entry (spans ~8-12 lines):
  //   <number><planCode>        e.g. "1EIZ4" or "24EIXU"
  //   <network type>            e.g. "EPO \nPROformance" or "POS \nPremier"
  //   <copays+benefits>         concatenated copays, deductibles, OOP, rates
  //   Final rates line ends with: $EE$ES$EC$EF<variance>%
  // Continuation pages (no "Medical Plan Alternates" header) have same format.
  const mpePages = pageBlocks.filter(b =>
    /MPE-\d+/i.test(b.substring(0, 200)) &&
    !/Option\s+\d+/i.test(b.substring(0, 600))
  );

  if (mpePages.length > 0) {
    console.log(`[UHC] Found ${mpePages.length} MPE alternate page(s)`);
    let mpeCount = 0;
    let lastProductFamily = null; // For continuation pages that lack header
    // Track plan codes from rate pages to avoid duplicating base plans
    const ratePageCodes = new Set(allPlans.map(p => p.planCode));

    for (const page of mpePages) {
      const lines = page.split('\n');

      // Extract product family from header: "Medical Plan Alternates for Insurance Choice, ..."
      let productFamily = null;
      const famMatch = page.match(/Medical Plan Alternates for ([^,*]+)/i);
      if (famMatch) {
        productFamily = famMatch[1].trim();
        lastProductFamily = productFamily;
      } else {
        productFamily = lastProductFamily; // Continuation page inherits
      }

      // Find plan entries: lines starting with <number><alphaCode> (e.g. "1EIZ4", "24EIXU")
      // or just <number>\n<code> for Surest-style entries
      const entryStarts = [];
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        // Match: digit(s) immediately followed by plan code (letter + alphanumerics)
        const em = line.match(/^(\d+)([A-Z][A-Z0-9_]+)$/i);
        if (em) {
          entryStarts.push({ lineIdx: i, seqNum: parseInt(em[1], 10), planCode: em[2] });
          continue;
        }
        // Surest-style: just a number on one line, then Surest code wrapped across next lines
        if (/^\d+$/.test(line) && i + 1 < lines.length) {
          const nextLine = lines[i + 1].trim();
          // Surest codes wrap: "Surest_AP1500_(" on one line, "2025)" on next
          if (/^Surest_/i.test(nextLine)) {
            // Collect the full code by joining lines until we hit the description line
            let fullCode = '';
            for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
              const part = lines[j].trim();
              if (/^Surest\s+\d{4}/i.test(part) && fullCode) break; // Description line starts
              fullCode += part;
            }
            // Clean: "Surest_AP1500_(2025)" → "Surest_AP1500_(2025)"
            fullCode = fullCode.replace(/\s+/g, '');
            entryStarts.push({ lineIdx: i, seqNum: parseInt(line, 10), planCode: fullCode });
          }
        }
      }

      for (let e = 0; e < entryStarts.length; e++) {
        const entry = entryStarts[e];
        const startLine = entry.lineIdx;
        const endLine = e + 1 < entryStarts.length
          ? entryStarts[e + 1].lineIdx
          : Math.min(startLine + 20, lines.length);

        // Gather all text for this entry
        const entryLines = lines.slice(startLine, endLine).map(l => l.trim()).filter(Boolean);
        const entryText = entryLines.join(' ');

        // Extract network type from lines after the plan code line
        let networkType = null;
        let planDescription = '';
        for (let j = 1; j < Math.min(4, entryLines.length); j++) {
          const el = entryLines[j];
          if (/^\$/.test(el) || /D&C|POD|Ded\+/i.test(el)) break;
          planDescription += ' ' + el;
        }
        planDescription = planDescription.trim()
          .replace(/PROforman\s+ce/gi, 'PROformance');
        if (/\bEPO\b/i.test(planDescription)) networkType = 'EPO';
        else if (/\bHMO\b/i.test(planDescription)) networkType = 'HMO';
        else if (/\bPOS\b/i.test(planDescription)) networkType = 'POS';
        else if (/\bPPO\b/i.test(planDescription)) networkType = 'PPO';
        else if (/\bSurest\b/i.test(planDescription) || /Surest/i.test(entry.planCode)) networkType = 'Surest';

        // Extract the 4 premium rates — they appear as 4 consecutive $ amounts
        // near the end of the entry, just before the variance percentage
        // Pattern: $EE$ES$EC$EF followed by optional -X.X% or 0%
        const allAmounts = [...entryText.matchAll(/\$([\d,]+\.\d{2})/g)];

        let premiumEE = null, premiumES = null, premiumEC = null, premiumEF = null;
        // The last 4 dollar amounts with cents are the premiums (EE, ES, EC, EF)
        if (allAmounts.length >= 4) {
          const last4 = allAmounts.slice(-4);
          premiumEE = parseMoney(last4[0][1]);
          premiumES = parseMoney(last4[1][1]);
          premiumEC = parseMoney(last4[2][1]);
          premiumEF = parseMoney(last4[3][1]);
        }

        // Extract deductible — look for "$X,XXX / $Y,YYY" pattern (ind / fam)
        // Require spaces around / to avoid matching copay splits like $70/$100
        // Use strict comma-formatted number pattern to avoid capturing trailing digits (e.g. 75%)
        let dedIndividual = null, dedFamily = null;
        const dedMatches = [...entryText.matchAll(/\$(\d{1,3}(?:,\d{3})*)\s+\/\s+\$(\d{1,3}(?:,\d{3})*)/g)];
        if (dedMatches.length > 0) {
          // First ind/fam pair is in-network deductible
          dedIndividual = parseMoney(dedMatches[0][1]);
          dedFamily = parseMoney(dedMatches[0][2]);
        }

        // Extract OOP — second ind/fam pair (or after "N/A" gap for out-of-network)
        let oopIndividual = null, oopFamily = null;
        // In-network OOP is typically the 2nd pair, but for HMO plans there may be
        // fewer pairs. Use a heuristic: last pair before the premiums.
        if (dedMatches.length >= 2) {
          // Find the pair that's the in-network OOP (2nd pair usually)
          oopIndividual = parseMoney(dedMatches[1][1]);
          oopFamily = parseMoney(dedMatches[1][2]);
        }

        // Extract coinsurance — percentage between deductible and OOP data
        let coinsurance = null;
        const coinsMatch = entryText.match(/(\d+)%\s*\$/);
        if (coinsMatch && parseInt(coinsMatch[1], 10) <= 100) {
          coinsurance = `${coinsMatch[1]}%`;
        }

        // Extract copays — PCP is first dollar amount, SPC is second
        let copayPCP = null, copaySpecialist = null, copayER = null;
        // The copay amounts appear early in the entry after network description
        // Look for small dollar amounts (< $500) before deductible
        const copayLine = entryLines.slice(1, 5).join(' ');
        const smallAmts = [...copayLine.matchAll(/\$(\d+)/g)];
        if (smallAmts.length >= 2) {
          const v1 = parseInt(smallAmts[0][1], 10);
          const v2 = parseInt(smallAmts[1][1], 10);
          if (v1 < 200) copayPCP = v1;
          if (v2 < 200) copaySpecialist = v2;
        }
        // ER copay — look for a value after "ER" or the 6th small amount
        const erMatch = entryText.match(/ER\s*\$(\d+)/i);
        if (erMatch) copayER = parseInt(erMatch[1], 10);
        else if (smallAmts.length >= 6) {
          const erVal = parseInt(smallAmts[5][1], 10);
          if (erVal < 2000) copayER = erVal;
        }

        // Determine HSA from plan code or description
        const isHSA = /HSA/i.test(planDescription) || /HSA/i.test(entry.planCode);

        // Derive product from family header or fallback
        let product = productFamily || null;

        if (premiumEE != null) {
          // Skip plans that already exist from rate pages (base plans appear in both)
          if (ratePageCodes.has(entry.planCode)) continue;

          const plan = {
            id: uuidv4(),
            carrier,
            planName: `${entry.planCode} (${planDescription || networkType || 'Unknown'})`,
            planCode: entry.planCode,
            networkType: isHSA ? (networkType || 'HSA') : networkType,
            metalLevel: null,
            deductibleIndividual: dedIndividual,
            deductibleFamily: dedFamily,
            oopMaxIndividual: oopIndividual,
            oopMaxFamily: oopFamily,
            coinsurance,
            copayPCP, copaySpecialist,
            copayUrgentCare: null,
            copayER,
            rxDeductible: null, rxTier1: null, rxTier2: null, rxTier3: null,
            premiumEE, premiumES, premiumEC, premiumEF,
            effectiveDate,
            product,
            hsaEligible: isHSA,
            ratingArea: null,
            underwritingNotes: null,
            extractionConfidence: 0,
            sourceFile,
          };

          const fields = [plan.deductibleIndividual, plan.deductibleFamily,
                          plan.oopMaxIndividual, plan.oopMaxFamily,
                          plan.copayPCP, plan.copaySpecialist, plan.copayER,
                          plan.premiumEE, plan.premiumES, plan.premiumEC, plan.premiumEF];
          const found = fields.filter(v => v != null).length;
          plan.extractionConfidence = Math.min(1, 0.3 + (found * 0.07));

          allPlans.push(plan);
          mpeCount++;
        }
      }
    }
    if (mpeCount > 0) {
      console.log(`[UHC] Extracted ${mpeCount} alternate plans from ${mpePages.length} MPE pages`);
    }
  }

  if (allPlans.length > 0) {
    console.log(`[UHC] Total: ${allPlans.length} plans (rate pages + alternates)`);
  }
  return allPlans;
}

// ── UHC helper: extract dollar amounts from a line ────────────────────────────
// Handles both spaced ("$3,000  $1,500  $2,000") and concatenated ("$3,000$1,500$2,000")
function extractDollarAmounts(line) {
  const cleaned = line
    .replace(/\b(?:after|AD)\s*(?:deductible|ded)?\b/gi, '')
    .replace(/\bcopay(?:ment)?\b/gi, '')
    .replace(/\bcoinsurance\b/gi, '')
    .replace(/\bper\s*(?:visit|occurrence)\b/gi, '');

  const amounts = [];
  const re = /\$\s*([\d,]+(?:\.\d{1,2})?)/g;
  let m;
  while ((m = re.exec(cleaned)) !== null) {
    const val = parseMoney(m[1]);
    if (val != null) amounts.push(val);
  }
  return amounts;
}

// ── UHC helper: assign benefit amounts to plan array ──────────────────────────
function assignBenefitToPlans(planData, numPlans, benefitDef, amounts) {
  if (benefitDef.special === 'coinsurance') {
    for (let p = 0; p < numPlans && p < amounts.length; p++) {
      if (planData[p].coinsurance == null && amounts[p] != null && amounts[p] <= 100) {
        planData[p].coinsurance = `${amounts[p]}%`;
      }
    }
    return;
  }
  const field = benefitDef.field;
  if (!field) return;
  for (let p = 0; p < numPlans && p < amounts.length; p++) {
    if (planData[p][field] == null && amounts[p] != null) {
      planData[p][field] = amounts[p];
    }
  }
}

// ── Strategy 8: Aetna Medical Cost Grid ───────────────────────────────────────
// Handles Aetna "Medical Cost Grid - Single Options" PDFs where each plan block
// looks like:
//   AFA OAAS 9100 100%Value CY V25          ← plan name (may wrap to next line)
//   ID: 30021412                             ← plan ID
//   $9100,100/0,0/0                          ← $ded,planCoins/memberCoins,pcp/spec
//   0%Med Ded Applies                        ← Rx info OR 3/10/50/80/20%up to...
//   OAAS$519.93                              ← network + EE premium
//   (3)                                      ← EE enrolled count
//   $1,307.56                                ← ES premium
//   (0)                                      ← ES enrolled
//   $1,042.59                                ← EC premium
//   (0)                                      ← EC enrolled
//   $1,797.05                                ← EF premium
//   (2)                                      ← EF enrolled
//   $5,153.89                                ← total premium
//   (5)                                      ← total enrolled
//   $1,468.20$3,151.20$486.23$48.26          ← agg, stoploss, admin, TRO
function extractFromAetnaCostGrid(text, sourceFile) {
  // Quick gate: must look like an Aetna cost grid
  if (!/Medical\s*Cost\s*Grid|AFA\s+(OAAS|CPOS)/i.test(text)) return [];

  const plans = [];
  const lines = text.split('\n');

  // Detect carrier from text/filename
  const carrier = detectCarrier(text, sourceFile) || 'Aetna';

  // Network map for Aetna plan codes
  const AETNA_NETWORKS = {
    'OAAS': 'OA',    // Open Access Aetna Select
    'CPOS II': 'CPOS',
    'CPOS': 'CPOS',
  };

  // Scan lines for plan blocks.  An Aetna plan block starts with a line
  // beginning with "AFA " and is followed by an ID line, benefit line, etc.
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Match plan name line: starts with "AFA "
    if (!/^AFA\s+/i.test(line)) continue;

    // Collect the full plan name (may wrap to the next line)
    let planName = line;
    let nextIdx = i + 1;

    // Check if next line is a continuation (starts with V25, V24, etc. or other
    // short continuation that doesn't start with ID: or $)
    if (nextIdx < lines.length) {
      const nextLine = lines[nextIdx].trim();
      if (/^V\d{2}\b/.test(nextLine) && nextLine.length < 15) {
        planName += ' ' + nextLine;
        nextIdx++;
      }
    }

    // Next must be "ID: NNNNNN"
    if (nextIdx >= lines.length) continue;
    const idLine = lines[nextIdx].trim();
    const idMatch = idLine.match(/^ID:\s*(\d+)/);
    if (!idMatch) continue;
    const planCode = idMatch[1];
    nextIdx++;

    // Next: benefit data line "$DED,COINS_PLAN/COINS_MEMBER,PCP/SPEC"
    if (nextIdx >= lines.length) continue;
    const benefitLine = lines[nextIdx].trim();
    const benefitMatch = benefitLine.match(/^\$([\d,]+),(\d+)\/(\d+),(\d+)\/(\d+)/);
    if (!benefitMatch) continue;
    nextIdx++;

    const deductibleIndividual = parseMoney(benefitMatch[1]);
    const coinsurancePlan = parseInt(benefitMatch[2], 10);
    const coinsuranceMember = parseInt(benefitMatch[3], 10);
    const copayPCP = parseInt(benefitMatch[4], 10) || null;
    const copaySpecialist = parseInt(benefitMatch[5], 10) || null;
    const coinsurance = coinsuranceMember > 0 ? `${coinsuranceMember}%` : null;

    // If copayPCP is 0 and coinsuranceMember is 0, it's likely "0% after ded"
    // (no copay — deductible applies to everything)
    const effectivePCP = copayPCP === 0 ? null : copayPCP;
    const effectiveSpec = copaySpecialist === 0 ? null : copaySpecialist;

    // Next line(s): Rx info — skip past them to find network+premium
    // Rx lines can be: "0%Med Ded Applies", "0%Med Ded Applies Tiers\n2-5",
    //                   "3/10/50/80/20%up to\n250/40%up to 500"
    let rxTier1 = null, rxTier2 = null, rxTier3 = null;

    // Collect Rx lines until we hit a network+premium line
    let rxText = '';
    while (nextIdx < lines.length) {
      const rxLine = lines[nextIdx].trim();
      // Network+premium line: NETWORK_NAME$AMOUNT e.g. "OAAS$519.93" or "CPOS II$542.85"
      if (/^(?:OAAS|CPOS\s*II?)\$[\d,]+\.\d{2}$/i.test(rxLine)) break;
      rxText += ' ' + rxLine;
      nextIdx++;
    }

    // Parse Rx tiers from collected Rx text (e.g. "3/10/50/80/20%up to 250/40%up to 500")
    const rxMatch = rxText.match(/(\d+)\/(\d+)\/(\d+)\/(\d+)/);
    if (rxMatch) {
      rxTier1 = parseInt(rxMatch[1], 10);
      rxTier2 = parseInt(rxMatch[2], 10);
      rxTier3 = parseInt(rxMatch[3], 10);
    }

    // Now parse network + EE premium line
    if (nextIdx >= lines.length) continue;
    const netPremLine = lines[nextIdx].trim();
    const netPremMatch = netPremLine.match(/^(OAAS|CPOS\s*II?)\$([\d,]+\.\d{2})$/i);
    if (!netPremMatch) continue;
    nextIdx++;

    const networkRaw = netPremMatch[1].trim().toUpperCase();
    const premiumEE = parseMoney(netPremMatch[2]);

    // Determine network type from plan name
    let networkType = null;
    if (/HSA/i.test(planName)) networkType = 'HSA';
    else if (/CPOS/i.test(networkRaw)) networkType = 'PPO';
    else if (/OAAS/i.test(networkRaw)) networkType = 'HMO';

    // Next lines: (EE_count), $ES_prem, (ES_count), $EC_prem, (EC_count), $EF_prem, (EF_count)
    // Parse premium lines — each premium is $AMOUNT followed by (count) on next line
    function readPremium() {
      if (nextIdx >= lines.length) return null;
      // Skip enrollment count lines like "(3)"
      let l = lines[nextIdx].trim();
      if (/^\(\d+\)$/.test(l)) { nextIdx++; l = nextIdx < lines.length ? lines[nextIdx].trim() : ''; }
      const pm = l.match(/^\$([\d,]+\.\d{2})$/);
      if (pm) { nextIdx++; return parseMoney(pm[1]); }
      return null;
    }

    // Skip EE enrolled count
    if (nextIdx < lines.length && /^\(\d+\)$/.test(lines[nextIdx].trim())) nextIdx++;

    const premiumES = readPremium();
    const premiumEC = readPremium();
    const premiumEF = readPremium();

    // Count filled benefit fields for confidence
    const fields = [deductibleIndividual, coinsurance, effectivePCP, effectiveSpec,
                    premiumEE, premiumES, premiumEC, premiumEF, rxTier1].filter(v => v != null);
    const confidence = Math.min(1, 0.5 + (fields.length * 0.06));

    plans.push({
      id: uuidv4(),
      carrier,
      planName: planName.trim(),
      planCode,
      networkType,
      metalLevel: null,  // Aetna cost grids don't use metal levels
      deductibleIndividual,
      deductibleFamily: null,
      oopMaxIndividual: null,  // Not in this grid format
      oopMaxFamily: null,
      coinsurance,
      copayPCP: effectivePCP,
      copaySpecialist: effectiveSpec,
      copayUrgentCare: null,
      copayER: null,
      rxDeductible: null,
      rxTier1,
      rxTier2,
      rxTier3,
      premiumEE,
      premiumES,
      premiumEC,
      premiumEF,
      effectiveDate: null,
      ratingArea: null,
      underwritingNotes: null,
      extractionConfidence: confidence,
      sourceFile,
    });
  }

  // Try to extract effective date from header
  const effDateMatch = text.match(/Eff\s*Date:\s*(\d{2}\/\d{2}\/\d{2,4})/i);
  if (effDateMatch && plans.length > 0) {
    for (const plan of plans) {
      plan.effectiveDate = effDateMatch[1];
    }
  }

  console.log(`[AETNA GRID] Extracted ${plans.length} plans from Aetna cost grid`);
  return plans;
}

// ── Strategy 9: BCBS Proposal Grid ────────────────────────────────────────────
// Handles Blue Cross Blue Shield "Illustrative Composite Billed Rates" PDFs.
// Grid section layout per plan block:
//   PLAN_ID (e.g. P9M1CHC, G654CHC)   ← plan code
//   [*N footnotes]                     ← optional
//   $DED_IN//                          ← deductible in-network
//   $DED_OUT or "Not Covered"          ← deductible out-of-network
//   $OOP_IN//                          ← OOP max in-network
//   $OOP_OUT or "Unlimited"/"Not Covered"
//   COINS_IN%//                        ← coinsurance in-network
//   COINS_OUT%                         ← coinsurance out-of-network
//   $PCP/$VIRT$SPEC                    ← PCP/Virtual, Specialist copay  (or DC/DC DC)
//   $ER//                              ← ER copay
//   COINS%                             ← ER coinsurance
//   $UC                                ← urgent care copay
//   InPat ded lines (skip)
//   OutPat ded lines (skip)
//   Rx tiers: $T1/$T2/$T3/$T4/$T5/$T6  ← or "100%" or "X%/X%/X%..."
//   $EO$ES$EC$EF$TOTAL                 ← premiums concatenated on one line
function extractFromBCBSGrid(text, sourceFile) {
  // Gate: must look like a BCBS proposal grid
  if (!/Illustrative\s*Composite|Blue\s*Choice\s*PPO|Blue\s*Advantage\s*HMO/i.test(text)) return [];

  const plans = [];
  const lines = text.split('\n');
  const carrier = 'BCBS';

  // Track current context as we scan
  let currentNetwork = null;  // 'PPO' or 'HMO'
  let currentMetal = null;    // 'Platinum', 'Gold', 'Silver', 'Bronze'
  let isHSA = false;

  // Extract effective date from header
  const effMatch = text.match(/Effective\s*Date:\s*(\d{2}\/\d{2}\/\d{4})/i);
  const effectiveDate = effMatch ? effMatch[1] : null;

  // BCBS plan IDs: letter + alphanumeric + 3-letter suffix (e.g. P9M1CHC, G654CHC, S663CHC, B662CHC, P9M1ADT)
  const PLAN_ID_RE = /^([A-Z]\w{2,8}(?:CHC|ADT|ADV|HMO|PPO))$/;
  // Footnote line
  const FOOTNOTE_RE = /^(\*\d+)+$/;
  // Deductible/OOP line: "$X//" pattern
  const DOLLAR_SLASH_RE = /^\$(\d[\d,]*)\s*\/\/$/;
  // OOP/Ded out-of-network value
  const OON_VALUE_RE = /^(?:\$(\d[\d,]*)|Unlimited|Not\s*Covered)$/;
  // Coinsurance line: "X%//"
  const COINS_IN_RE = /^(\d+)%\s*\/\/$/;
  // Coinsurance out: "X%" or "Not Covered"
  const COINS_OUT_RE = /^(?:(\d+)%|Not\s*Covered)$/;
  // Premium line: multiple $X,XXX.XX concatenated
  const PREMIUM_LINE_RE = /^(\$[\d,]+\.\d{2}){3,5}$/;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Track network context
    if (/Blue\s*Choice\s*PPO/i.test(line)) { currentNetwork = 'PPO'; continue; }
    if (/Blue\s*Advantage\s*HMO/i.test(line)) { currentNetwork = 'HMO'; continue; }

    // Track metal level context
    if (/^Platinum$/i.test(line)) { currentMetal = 'Platinum'; continue; }
    if (/^Gold$/i.test(line)) { currentMetal = 'Gold'; continue; }
    if (/^Silver$/i.test(line)) { currentMetal = 'Silver'; continue; }
    if (/^Bronze$/i.test(line)) { currentMetal = 'Bronze'; continue; }
    if (/^Expanded\s*Bronze$/i.test(line)) { currentMetal = 'Bronze'; continue; }
    if (/^HSA\s*Plans$/i.test(line)) { isHSA = true; continue; }
    if (/^(?:PPO|HMO)\s*Plans$/i.test(line)) { isHSA = false; continue; }

    // Skip header/footer lines
    if (/^Plan\s*ID$/i.test(line)) continue;
    if (/^(?:Individual|Coinsurance|Primary|ER|Urgent|In-Patient|Out-Patient|Non-|EOESECEF|Total|Monthly|Medical|Cost|Ded|Network|Preferred|Pharmacy)/i.test(line)) continue;
    if (/^Go\s*to\s*Proposal/i.test(line)) continue;
    if (/^Blue\s*Cross\s*and/i.test(line)) continue;
    if (/^Quote\s*ID:/i.test(line)) continue;
    if (/^No\.\s*of\s*Employees/i.test(line)) continue;
    if (/^Printed:/i.test(line)) continue;
    if (/^\*\d/.test(line)) continue;

    // Match plan ID
    const planIdMatch = line.match(PLAN_ID_RE);
    if (!planIdMatch) continue;

    const planCode = planIdMatch[1];
    let nextIdx = i + 1;

    // Skip optional footnote markers
    if (nextIdx < lines.length && FOOTNOTE_RE.test(lines[nextIdx].trim())) nextIdx++;

    // Helper to read next non-empty line
    function nextLine() {
      while (nextIdx < lines.length) {
        const l = lines[nextIdx].trim();
        nextIdx++;
        if (l.length > 0) return l;
      }
      return '';
    }

    // Parse deductible in-network: "$X//"
    let l = nextLine();
    let dedInMatch = l.match(/^\$(\d[\d,]*)\s*\/\//);
    if (!dedInMatch) continue;
    const deductibleIndividual = parseMoney(dedInMatch[1]);

    // Deductible out-of-network
    l = nextLine();
    // Could be "$5000", "Not Covered", or another value

    // OOP max in-network: "$X//"
    l = nextLine();
    let oopInMatch = l.match(/^\$(\d[\d,]*)\s*\/\//);
    const oopMaxIndividual = oopInMatch ? parseMoney(oopInMatch[1]) : null;

    // OOP max out-of-network
    l = nextLine();

    // Coinsurance in-network: "X%//"
    l = nextLine();
    let coinsInMatch = l.match(/^(\d+)%\s*\/\//);
    const coinsurance = coinsInMatch ? `${coinsInMatch[1]}%` : null;

    // Coinsurance out-of-network
    l = nextLine();

    // PCP/Virtual + Specialist copay: "$PCP/$VIRT$SPEC" or "DC/DC DC" or "$PCP/$VIRT DC"
    l = nextLine();
    let copayPCP = null;
    let copaySpecialist = null;
    // Patterns: "$20/$20$40", "DC/DCDC", "$0/$0$40"
    const pcpMatch = l.match(/^\$(\d+)\/\$?(\d+)\$?(\d+)/);
    if (pcpMatch) {
      copayPCP = parseInt(pcpMatch[1], 10);
      copaySpecialist = parseInt(pcpMatch[3], 10);
    } else {
      const dcMatch = l.match(/^DC\/DC\$?(\d+)/);
      if (dcMatch) {
        copaySpecialist = parseInt(dcMatch[1], 10);
      }
      // If fully "DC/DCDC" — all null, which is fine
    }

    // ER copay: "$X//" or "DC//"
    l = nextLine();
    let copayER = null;
    const erMatch = l.match(/^\$(\d+)\s*\/\//);
    if (erMatch) copayER = parseInt(erMatch[1], 10);

    // ER coinsurance: skip
    l = nextLine();

    // Urgent care: "$X" or "DC"
    l = nextLine();
    let copayUrgentCare = null;
    const ucMatch = l.match(/^\$(\d+)$/);
    if (ucMatch) copayUrgentCare = parseInt(ucMatch[1], 10);

    // In-patient deductible in/out: 2 lines
    l = nextLine(); // e.g. "DC//" or "$150//"
    l = nextLine(); // e.g. "DC" or "Not Covered"

    // Out-patient deductible in/out: 2 lines
    l = nextLine(); // e.g. "DC//" or "$100//"
    l = nextLine(); // e.g. "DC" or "$200"

    // Rx tiers: "$T1/$T2/$T3/" (may wrap) or "100%" or "X%/X%/X%..."
    // NOTE: Sometimes "100%" or "X%..." is concatenated with premiums on same line,
    // e.g. "100%$982.85$1,965.70$1,965.70$2,948.55$16,708.45"
    l = nextLine();
    let rxTier1 = null, rxTier2 = null, rxTier3 = null;
    let premiumLineOverride = null;  // If premiums are on the Rx line
    // Match dollar-based Rx: "$10/$20/$70/" or "$10/$20/$55/"
    let rxMatch = l.match(/^\$(\d+)\/\$(\d+)\/\$(\d+)/);
    if (rxMatch) {
      rxTier1 = parseInt(rxMatch[1], 10);
      rxTier2 = parseInt(rxMatch[2], 10);
      rxTier3 = parseInt(rxMatch[3], 10);
      // Next line has more Rx tiers (e.g. "$120/$150/$250") — skip
      l = nextLine();
    } else if (/^\d+%.*\$[\d,]+\.\d{2}/.test(l)) {
      // Percentage Rx concatenated with premiums: "100%$982.85..." or "80%/80%/...$1,234.56..."
      premiumLineOverride = l;
    } else if (/^\d+%/.test(l)) {
      // Percentage-based Rx like "80%/80%/70%/60%/60%/50%"
      // Can't map to dollar copays — leave as null
      // But check for continuation
      if (l.endsWith('/')) l = nextLine();
    } else if (/^100%$/.test(l)) {
      // 100% coverage — no Rx copays — premiums on next line
    }

    // Premium line: "$EO$ES$EC$EF$TOTAL" all concatenated
    l = premiumLineOverride || nextLine();
    // Match: $1,638.23$3,276.46$3,276.46$4,914.69$27,849.91
    const premiums = [];
    const premRe = /\$(\d[\d,]*\.\d{2})/g;
    let pm;
    while ((pm = premRe.exec(l)) !== null) {
      premiums.push(parseMoney(pm[1]));
    }

    const premiumEE = premiums[0] || null;
    const premiumES = premiums[1] || null;
    const premiumEC = premiums[2] || null;
    const premiumEF = premiums[3] || null;

    // Determine network type
    let networkType = isHSA ? 'HSA' : currentNetwork;

    // Build plan name from context + plan code
    const metalStr = currentMetal || '';
    const networkStr = currentNetwork === 'HMO' ? 'Blue Advantage HMO' : 'Blue Choice PPO';
    const planName = `${networkStr} ${metalStr} ${planCode}`.replace(/\s+/g, ' ').trim();

    // Score confidence
    const fields = [deductibleIndividual, oopMaxIndividual, coinsurance,
                    copayPCP, copaySpecialist, copayER, copayUrgentCare,
                    premiumEE, premiumES, premiumEC, premiumEF,
                    rxTier1].filter(v => v != null);
    const confidence = Math.min(1, 0.5 + (fields.length * 0.05));

    plans.push({
      id: uuidv4(),
      carrier,
      planName,
      planCode,
      networkType,
      metalLevel: currentMetal,
      deductibleIndividual,
      deductibleFamily: null,
      oopMaxIndividual,
      oopMaxFamily: null,
      coinsurance,
      copayPCP: copayPCP === 0 ? null : copayPCP,
      copaySpecialist: copaySpecialist === 0 ? null : copaySpecialist,
      copayUrgentCare,
      copayER,
      rxDeductible: null,
      rxTier1,
      rxTier2,
      rxTier3,
      premiumEE,
      premiumES,
      premiumEC,
      premiumEF,
      effectiveDate,
      ratingArea: null,
      underwritingNotes: null,
      extractionConfidence: confidence,
      sourceFile,
    });
  }

  console.log(`[BCBS GRID] Extracted ${plans.length} plans from BCBS proposal grid`);
  return plans;
}

// ── Strategy 6: Repeated-keyword extraction ───────────────────────────────────
// If the same benefit keyword (e.g., "deductible") appears multiple times in the
// text, each time followed by a dollar amount, that likely means the PDF text was
// extracted column-by-column (one plan at a time).  Each Nth occurrence of each
// keyword belongs to plan N.
function extractByRepeatedKeywords(text, lines, sourceFile) {
  const DOLLAR_RE = /\$\s*([\d,]+(?:\.\d{1,2})?)/;

  // Benefit keywords to look for — each may appear once per plan
  const benefitDefs = [
    { field: 'deductibleIndividual', re: /deductible/i, exclude: /family|rx|pharm|drug/i },
    { field: 'oopMaxIndividual', re: /out[\s-]*of[\s-]*pocket|oop|moop/i, exclude: /family/i },
    { field: 'copayPCP', re: /pcp|primary\s*care|office\s*visit/i },
    { field: 'copaySpecialist', re: /specialist/i },
    { field: 'copayER', re: /emergency/i },
    { field: 'premiumEE', re: /employee\s*only|emp\s*only|\bee\b(?!\+)|single/i, exclude: /spouse|child|family/i },
    { field: 'premiumES', re: /emp(?:loyee)?\s*[\+\/&]\s*(?:spouse|sp)|\bes\b/i },
    { field: 'premiumEC', re: /emp(?:loyee)?\s*[\+\/&]\s*(?:child|ch)|\bec\b/i },
    { field: 'premiumEF', re: /\bfamily\b|\bef\b/i, exclude: /deductible|oop|out.of.pocket|family.*deductible/i },
  ];

  // For each benefit keyword, find all lines that have it + a dollar amount
  const fieldOccurrences = {};
  let maxOccurrences = 0;
  for (const { field, re, exclude } of benefitDefs) {
    const values = [];
    for (const line of lines) {
      if (!re.test(line)) continue;
      if (exclude && exclude.test(line)) continue;
      const dm = line.match(DOLLAR_RE);
      if (dm) {
        values.push(parseMoney(dm[1]));
      }
    }
    fieldOccurrences[field] = values;
    if (values.length > maxOccurrences) maxOccurrences = values.length;
  }

  // We need at least 2 occurrences of some benefit keyword, and ideally
  // multiple keywords all repeating the same number of times
  if (maxOccurrences < 2) return [];

  // Determine numPlans: the most common repetition count >= 2
  const repCounts = {};
  for (const values of Object.values(fieldOccurrences)) {
    if (values.length >= 2) {
      repCounts[values.length] = (repCounts[values.length] || 0) + 1;
    }
  }
  let numPlans = maxOccurrences;
  let bestVotes = 0;
  for (const [cnt, votes] of Object.entries(repCounts)) {
    if (votes > bestVotes || (votes === bestVotes && Number(cnt) > numPlans)) {
      numPlans = Number(cnt);
      bestVotes = votes;
    }
  }

  console.log(`[EXTRACT] Repeated-keyword detection: numPlans=${numPlans}, maxOccurrences=${maxOccurrences}`);
  console.log(`[EXTRACT] Field occurrences:`, Object.fromEntries(Object.entries(fieldOccurrences).map(([k, v]) => [k, v.length])));

  // Build plans
  const plans = [];
  for (let p = 0; p < numPlans; p++) {
    const plan = {
      id: uuidv4(),
      carrier: null, planName: `Plan ${p + 1}`, planCode: null,
      networkType: null, metalLevel: null,
      deductibleIndividual: null, deductibleFamily: null,
      oopMaxIndividual: null, oopMaxFamily: null,
      coinsurance: null, copayPCP: null, copaySpecialist: null,
      copayUrgentCare: null, copayER: null,
      rxDeductible: null, rxTier1: null, rxTier2: null, rxTier3: null,
      premiumEE: null, premiumES: null, premiumEC: null, premiumEF: null,
      effectiveDate: null, ratingArea: null, underwritingNotes: null,
      extractionConfidence: 0, sourceFile,
    };

    for (const [field, values] of Object.entries(fieldOccurrences)) {
      if (values.length >= numPlans && values[p] != null) {
        plan[field] = values[p];
      }
    }
    plans.push(plan);
  }

  // Try to find plan names — look for lines matching plan-name patterns
  // that appear before or near each plan's data
  const allNames = [];
  for (const line of lines) {
    const names = findPlanNamesInLine(line);
    for (const n of names) {
      if (!allNames.find(a => a.name === n.name)) allNames.push(n);
    }
  }
  for (let p = 0; p < numPlans && p < allNames.length; p++) {
    plans[p].planName = allNames[p].name;
    if (allNames[p].network) plans[p].networkType = allNames[p].network;
    if (allNames[p].metal) plans[p].metalLevel = allNames[p].metal;
  }

  // Filter: at least 2 non-null key fields
  return plans.filter(p => {
    const kf = [p.carrier, p.planName, p.networkType, p.deductibleIndividual,
                p.oopMaxIndividual, p.copayPCP, p.premiumEE, p.premiumES, p.premiumEC, p.premiumEF];
    const found = kf.filter(v => v !== null && v !== undefined).length;
    p.extractionConfidence = Math.min(1, found / 6);
    return found >= 2;
  });
}

// ── Benefit comparison grid extractor ─────────────────────────────────────────
// Detects common carrier quote format:
//   Row labels (Deductible, OOP Max, Copay, etc.) in first "column", and
//   dollar amounts aligned in subsequent columns — one column per plan.
// Also handles the "$X / $Y" individual/family pair format.
function extractFromBenefitGrid(text, sourceFile) {
  const lines = text.split('\n');

  // Identify "benefit label" lines that contain both a recognizable label
  // and at least one dollar amount.
  const BENEFIT_LABEL_RE = /deductible|out[\s-]*of[\s-]*pocket|oop|copay|co-?pay|coinsurance|premium|office\s*visit|pcp|primary\s*care|specialist|emergency|urgent\s*care|generic|preferred\s*brand|non[\s-]*preferred|tier\s*[123]|rx|pharmacy|moop|max(?:imum)?\s*out/i;
  const DOLLAR_RE = /\$\s*[\d,]+(?:\.\d{1,2})?/g;
  const TIER_LABEL_RE = /\b(?:employee\s*only|emp(?:loyee)?\s*[\+\/&]\s*(?:spouse|sp|child(?:ren)?|ch|family|fam)|single|ee|es|ec|ef|family)\b/i;

  // Helper: detect "$X / $Y" pairs vs standalone dollar amounts
  const PAIR_RE = /\$\s*([\d,]+(?:\.\d{1,2})?)\s*[\/|]\s*\$\s*([\d,]+(?:\.\d{1,2})?)/g;
  function parseDollarGroups(line) {
    // First detect "$X / $Y" pairs
    const pairs = [];
    let pm;
    const pairRe = new RegExp(PAIR_RE.source, 'g');
    while ((pm = pairRe.exec(line)) !== null) {
      pairs.push({ primary: parseMoney(pm[1]), secondary: parseMoney(pm[2]), fullMatch: pm[0] });
    }
    if (pairs.length > 0) {
      return { type: 'pairs', count: pairs.length, primary: pairs.map(p => p.primary), secondary: pairs.map(p => p.secondary), all: pairs };
    }
    // No pairs — extract standalone amounts
    const amounts = (line.match(DOLLAR_RE) || []).map(s => parseMoney(s)).filter(v => v != null);
    return { type: 'singles', count: amounts.length, primary: amounts, secondary: [], all: amounts };
  }

  // Pass 1: find lines with a benefit label and dollar amounts
  // Also check the next 1-3 lines if the label line has no dollar amounts
  const labeledRows = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!BENEFIT_LABEL_RE.test(line) && !TIER_LABEL_RE.test(line)) continue;
    let groups = parseDollarGroups(line);
    // If the label line has no dollar amounts, check the next 1-3 lines
    if (groups.count === 0) {
      for (let j = 1; j <= 3 && i + j < lines.length; j++) {
        const nextLine = lines[i + j];
        if (BENEFIT_LABEL_RE.test(nextLine) || TIER_LABEL_RE.test(nextLine)) break;
        const nextGroups = parseDollarGroups(nextLine);
        if (nextGroups.count > 0) {
          groups = nextGroups;
          break;
        }
      }
    }
    if (groups.count === 0) continue;
    labeledRows.push({ index: i, line, groups, label: line });
  }

  if (labeledRows.length < 1) {
    console.log(`[GRID] Only ${labeledRows.length} labeled rows found — skipping grid strategy`);
    return [];
  }

  console.log(`[GRID] Found ${labeledRows.length} labeled rows:`, labeledRows.map(r => ({ label: r.line.substring(0, 60), count: r.groups.count, type: r.groups.type })));

  // Find the number of plans from column counts
  // Use "groups.count" (which handles $X/$Y pairs as 1 group)
  const countFreq = {};
  for (const r of labeledRows) {
    const c = r.groups.count;
    if (c >= 1) countFreq[c] = (countFreq[c] || 0) + 1;
  }
  console.log(`[GRID] Column count frequency:`, countFreq);

  // Pick numPlans: prefer the highest count with freq >= 2, else highest count with freq >= 1
  let numPlans = 1;
  const countsWithFreq2 = Object.entries(countFreq)
    .filter(([cnt, freq]) => Number(cnt) >= 2 && freq >= 1)
    .map(([cnt]) => Number(cnt));
  if (countsWithFreq2.length > 0) {
    // Among counts >= 2, prefer the one with the highest frequency; break ties by higher count
    let bestCount = 0, bestFreq = 0;
    for (const [cnt, freq] of Object.entries(countFreq)) {
      const n = Number(cnt);
      if (n < 2) continue;
      if (freq > bestFreq || (freq === bestFreq && n > bestCount)) {
        bestCount = n;
        bestFreq = freq;
      }
    }
    numPlans = bestCount;
  } else if (countFreq[1] && countFreq[1] >= 2) {
    // Multiple rows with 1 amount each — might be sequential per-plan layout
    numPlans = 1;
  }

  if (numPlans < 1) return [];

  // Build plan stubs
  const planData = [];
  for (let p = 0; p < numPlans; p++) {
    planData.push({
      id: uuidv4(),
      carrier: null, planName: null, planCode: null,
      networkType: null, metalLevel: null,
      deductibleIndividual: null, deductibleFamily: null,
      oopMaxIndividual: null, oopMaxFamily: null,
      coinsurance: null, copayPCP: null, copaySpecialist: null,
      copayUrgentCare: null, copayER: null,
      rxDeductible: null, rxTier1: null, rxTier2: null, rxTier3: null,
      premiumEE: null, premiumES: null, premiumEC: null, premiumEF: null,
      effectiveDate: null, ratingArea: null, underwritingNotes: null,
      extractionConfidence: 0, sourceFile,
    });
  }

  // Try to find plan name header row (line before first labeled row with plan-like names)
  const firstLabelIdx = labeledRows[0].index;
  for (let j = Math.max(0, firstLabelIdx - 15); j < firstLabelIdx; j++) {
    const hLine = lines[j];
    const names = findPlanNamesInLine(hLine);
    if (names.length >= numPlans) {
      for (let p = 0; p < numPlans && p < names.length; p++) {
        planData[p].planName = names[p].name;
        if (names[p].network) planData[p].networkType = names[p].network;
        if (names[p].metal) planData[p].metalLevel = names[p].metal;
      }
      break;
    }
  }

  // If no plan names found from header, try to find them via other patterns
  if (!planData[0].planName) {
    // Look for plan names anywhere in the first ~30 lines
    for (let j = 0; j < Math.min(30, lines.length); j++) {
      const names = findPlanNamesInLine(lines[j]);
      if (names.length >= numPlans) {
        for (let p = 0; p < numPlans && p < names.length; p++) {
          planData[p].planName = names[p].name;
          if (names[p].network) planData[p].networkType = names[p].network;
          if (names[p].metal) planData[p].metalLevel = names[p].metal;
        }
        break;
      }
    }
    // Fallback names
    for (let p = 0; p < numPlans; p++) {
      if (!planData[p].planName) planData[p].planName = `Plan ${p + 1}`;
    }
  }

  // Map labeled rows to plan fields
  for (const row of labeledRows) {
    const lbl = row.label.toLowerCase();
    const amounts = row.groups.primary;
    const secondaryAmounts = row.groups.secondary || [];
    const isPair = row.groups.type === 'pairs';

    // Determine what benefit this row describes
    const assignAllPlans = (field) => {
      for (let p = 0; p < numPlans; p++) {
        if (p < amounts.length && amounts[p] != null && planData[p][field] == null) {
          planData[p][field] = amounts[p];
        }
      }
    };

    const assignSecondary = (field) => {
      for (let p = 0; p < numPlans; p++) {
        if (p < secondaryAmounts.length && secondaryAmounts[p] != null && planData[p][field] == null) {
          planData[p][field] = secondaryAmounts[p];
        }
      }
    };

    // Deductible (individual — primary; family — secondary if "$X / $Y" pair)
    if (/deductible/i.test(lbl) && !/rx|pharm|drug/i.test(lbl)) {
      if (/family/i.test(lbl)) {
        assignAllPlans('deductibleFamily');
      } else {
        assignAllPlans('deductibleIndividual');
        if (isPair) {
          assignSecondary('deductibleFamily');
        } else if (numPlans === 1 && amounts.length === 2 && !(/family/i.test(lbl))) {
          planData[0].deductibleIndividual = amounts[0];
          planData[0].deductibleFamily = amounts[1];
        }
      }
    }
    // OOP Max
    else if (/out[\s-]*of[\s-]*pocket|oop|moop/i.test(lbl) || /max(?:imum)?\s*out/i.test(lbl)) {
      if (/family/i.test(lbl)) {
        assignAllPlans('oopMaxFamily');
      } else {
        assignAllPlans('oopMaxIndividual');
        if (isPair) {
          assignSecondary('oopMaxFamily');
        } else if (numPlans === 1 && amounts.length === 2 && !(/family/i.test(lbl))) {
          planData[0].oopMaxIndividual = amounts[0];
          planData[0].oopMaxFamily = amounts[1];
        }
      }
    }
    // PCP / Office visit copay
    else if (/pcp|primary\s*care|office\s*visit|doctor\s*visit|physician/i.test(lbl)) {
      assignAllPlans('copayPCP');
    }
    // Specialist copay
    else if (/specialist/i.test(lbl)) {
      assignAllPlans('copaySpecialist');
    }
    // ER copay
    else if (/emergency/i.test(lbl)) {
      assignAllPlans('copayER');
    }
    // Urgent care copay
    else if (/urgent/i.test(lbl)) {
      assignAllPlans('copayUrgentCare');
    }
    // Rx tiers
    else if (/generic|tier\s*1/i.test(lbl)) {
      assignAllPlans('rxTier1');
    }
    else if (/preferred\s*brand|tier\s*2/i.test(lbl)) {
      assignAllPlans('rxTier2');
    }
    else if (/non[\s-]*preferred|tier\s*3|specialty/i.test(lbl)) {
      assignAllPlans('rxTier3');
    }
    else if (/rx\s*deductible|pharmacy\s*deductible|drug\s*deductible/i.test(lbl)) {
      assignAllPlans('rxDeductible');
    }
    // Premiums — by tier label in the row
    else if (/premium|rate|monthly/i.test(lbl) || TIER_LABEL_RE.test(lbl)) {
      if (/employee\s*only|emp\s*only|\bee\b|single/i.test(lbl)) {
        assignAllPlans('premiumEE');
      } else if (/emp(?:loyee)?\s*[\+\/&]\s*(?:spouse|sp)|\bes\b/i.test(lbl)) {
        assignAllPlans('premiumES');
      } else if (/emp(?:loyee)?\s*[\+\/&]\s*(?:child|ch)|\bec\b/i.test(lbl)) {
        assignAllPlans('premiumEC');
      } else if (/family|emp(?:loyee)?\s*[\+\/&]\s*(?:fam)|\bef\b/i.test(lbl)) {
        assignAllPlans('premiumEF');
      }
    }
  }

  // Calculate confidence and filter
  const result = planData.filter(p => {
    const kf = [p.carrier, p.planName, p.networkType, p.deductibleIndividual,
                p.oopMaxIndividual, p.copayPCP, p.premiumEE, p.premiumES, p.premiumEC, p.premiumEF];
    const found = kf.filter(v => v !== null && v !== undefined).length;
    p.extractionConfidence = Math.min(1, found / 6);
    return found >= 2;
  });

  return result;
}

// ── Post-extraction enrichment ────────────────────────────────────────────────
// Scans the full text for labeled benefit values and fills any null fields on
// all extracted plans.  If multiple plans exist but a benefit appears only once,
// we assume it applies to all plans (common for same-carrier quotes).
function enrichPlansFromText(plans, fullText) {
  if (plans.length === 0) return;

  // Build a set of {field, patterns} to try
  const fieldPatterns = [
    { field: 'deductibleIndividual', patterns: [
      /(?:individual|in[\s-]?network)\s*deductible\s*[:\-]?\s*\$?([\d,]+)/i,
      /deductible\s*(?:\([^)]*\))?\s*[:\-·.…]*\s*\$?([\d,]+)\s*(?:\/\s*\$?[\d,]+)?(?:\s*(?:individual|person|member|single))?/i,
      /(?:in[\s-]?network|medical)\s*deductible\s*[:\-·.…]*\s*\$?([\d,]+)/i,
      /deductible[^$\n]{0,40}\$\s*([\d,]+)/i,
    ]},
    { field: 'deductibleFamily', patterns: [
      /family\s*deductible\s*[:\-]?\s*\$?([\d,]+)/i,
      /deductible[^$\n]*\$\s*[\d,]+(?:\.\d+)?\s*[\/|]\s*\$\s*([\d,]+)/i,
      /deductible\s*\([^)]*\)\s*[:\-·.…]*\s*\$\s*[\d,]+(?:\.\d+)?\s*[\/|]\s*\$\s*([\d,]+)/i,
    ]},
    { field: 'oopMaxIndividual', patterns: [
      /(?:individual|in[\s-]?network)\s*(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*[:\-]?\s*\$?([\d,]+)/i,
      /(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*(?:\([^)]*\))?\s*[:\-·.…]*\s*\$?([\d,]+)/i,
      /(?:max(?:imum)?\s*)?out[\s-]*of[\s-]*pocket\s*[:\-·.…]*\s*\$?([\d,]+)/i,
      /moop\s*[:\-·.…]*\s*\$?([\d,]+)/i,
      /oop[^$\n]{0,40}\$\s*([\d,]+)/i,
    ]},
    { field: 'oopMaxFamily', patterns: [
      /family\s*(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*[:\-]?\s*\$?([\d,]+)/i,
      /(?:out[\s-]*of[\s-]*pocket|oop)[^$\n]*\$\s*[\d,]+(?:\.\d+)?\s*[\/|]\s*\$\s*([\d,]+)/i,
      /(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*\([^)]*\)\s*[:\-·.…]*\s*\$\s*[\d,]+(?:\.\d+)?\s*[\/|]\s*\$\s*([\d,]+)/i,
    ]},
    { field: 'copayPCP', patterns: [
      /(?:pcp|primary\s*care|office\s*visit|doctor|physician)\s*(?:copay?|co[\s-]?pay(?:ment)?|visit)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
      /(?:pcp|office\s*visit|primary\s*care)\s*[:\-·.…]*\s*\$?([\d]+)/i,
    ]},
    { field: 'copaySpecialist', patterns: [
      /specialist\s*(?:copay?|co[\s-]?pay(?:ment)?|visit)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    ]},
    { field: 'copayER', patterns: [
      /(?:emergency\s*(?:room|dept|department|services?)|er\b)\s*(?:copay?|co[\s-]?pay(?:ment)?)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    ]},
    { field: 'copayUrgentCare', patterns: [
      /urgent\s*care\s*(?:copay?|co[\s-]?pay(?:ment)?)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    ]},
    { field: 'coinsurance', patterns: [
      /coinsurance\s*[:\-·.…]*\s*(\d+)\s*%/i,
      /(\d+)\s*%\s*(?:co[\s-]?insurance|after\s+deductible)/i,
    ]},
  ];

  for (const { field, patterns } of fieldPatterns) {
    // Check if ALL plans are missing this field
    const allMissing = plans.every(p => p[field] == null);
    if (!allMissing) continue;

    const raw = firstMatch(fullText, patterns);
    if (raw == null) continue;
    const val = field === 'coinsurance' ? raw : parseMoney(raw);
    if (val == null) continue;

    // Apply to all plans that are missing this field
    for (const plan of plans) {
      if (plan[field] == null) plan[field] = val;
    }
  }

  // Recalculate confidence for each plan
  for (const plan of plans) {
    const kf = [plan.carrier, plan.planName, plan.networkType, plan.metalLevel,
                plan.deductibleIndividual, plan.oopMaxIndividual, plan.copayPCP,
                plan.premiumEE, plan.premiumES, plan.premiumEC, plan.premiumEF];
    const found = kf.filter(v => v !== null && v !== undefined).length;
    plan.extractionConfidence = Math.min(1, found / 6);
  }
}

// ── De-duplicate plans ────────────────────────────────────────────────────────
function deduplicatePlans(plans) {
  // Normalize a plan name for dedup comparison
  function normalizeName(name) {
    return (name || '')
      .toLowerCase()
      .replace(/^(?:illustrative\s*)?quote\s*/i, '')
      .replace(/[\r\n]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function recalcConfidence(plan) {
    const kf = [plan.carrier, plan.planName, plan.networkType, plan.metalLevel,
                plan.deductibleIndividual, plan.oopMaxIndividual, plan.copayPCP,
                plan.premiumEE, plan.premiumES, plan.premiumEC, plan.premiumEF];
    plan.extractionConfidence = Math.min(1, kf.filter(v => v != null).length / 6);
  }

  function mergePlanInto(existing, donor) {
    for (const k of Object.keys(donor)) {
      if (k === 'id' || k === 'extractionConfidence') continue;
      if (existing[k] == null && donor[k] != null) existing[k] = donor[k];
    }
    recalcConfidence(existing);
  }

  const result = [];
  for (const plan of plans) {
    const normA = normalizeName(plan.planName);
    // Try to find an existing plan to merge with
    let merged = false;
    for (const existing of result) {
      const normB = normalizeName(existing.planName);
      // Match if: same name, OR one is a substring of the other (handles truncated names)
      const nameMatch = normA === normB ||
        (normA.length > 10 && normB.length > 10 && (normA.includes(normB) || normB.includes(normA)));
      // Carrier must match if both are set
      const carrierMatch = !plan.carrier || !existing.carrier ||
        String(plan.carrier).toLowerCase() === String(existing.carrier).toLowerCase();
      // Don't merge if both plans have distinct premium values (they are different plans)
      const bothHavePremiums = plan.premiumEE != null && existing.premiumEE != null;
      const premiumsDiffer = bothHavePremiums && Math.abs(plan.premiumEE - existing.premiumEE) > 0.01;
      if (nameMatch && carrierMatch && !premiumsDiffer) {
        mergePlanInto(existing, plan);
        merged = true;
        break;
      }
    }
    if (!merged) {
      result.push(plan);
    }
  }
  return result;
}

// ── Detect carrier from text or filename ──────────────────────────────────────
function detectCarrier(text, sourceFile) {
  // Check filename first
  const filePart = (sourceFile || '').replace(/[_\-\.]/g, ' ');
  const fm = filePart.match(CARRIER_PATTERN);
  if (fm) return normalizeCarrier(fm[1]);
  // Check text
  const tm = text.match(CARRIER_PATTERN);
  if (tm) return normalizeCarrier(tm[1]);
  return null;
}

function normalizeCarrier(raw) {
  if (!raw) return null;
  const up = raw.replace(/\s+/g, ' ').trim();
  if (/bcbs|blue\s*cross/i.test(up)) return 'BCBS';
  if (/united\s*health|uhc/i.test(up)) return 'UnitedHealthcare';
  if (/baylor|bsw/i.test(up)) return 'Baylor Scott & White';
  return up.charAt(0).toUpperCase() + up.slice(1);
}

// ── Table-based extraction (rate sheets) ──────────────────────────────────────
// Real carrier PDFs often render as rows of data like:
//   "Blue Choice PPO Gold   $40   $2,000  $6,000  $500.00  $1,200.00  $900.00  $1,800.00"
// or columnar rate tables where premiums appear on rows labeled EE, ES, EC, EF
function extractFromTable(lines, fullText, sourceFile) {
  const plans = [];

  // Strategy 1: Look for premium rate table rows
  // e.g. "Employee Only   $500.00  $520.00  $480.00"
  //       header row lists plan names across columns
  const rateTablePlans = extractFromRateTable(lines, sourceFile);
  if (rateTablePlans.length > 0) return rateTablePlans;

  // Strategy 2: Look for rows that contain plan-like data with $ amounts
  // e.g. "Blue Choice PPO 500   $40 copay   $2,000/$4,000   $6,000/$12,000   $500.00"
  const rowPlans = extractFromDataRows(lines, fullText, sourceFile);
  if (rowPlans.length > 0) return rowPlans;

  return plans;
}

// Look for a rate table structure:
// header row with plan names, followed by rows for EE/ES/EC/EF rates
// Also looks for benefit rows (deductible, OOP, copay) with dollar amounts per plan column
function extractFromRateTable(lines, sourceFile) {
  const plans = [];
  const tierLabels = {
    ee: /\b(?:employee\s*only|ee\b|single\b|emp\s*only)/i,
    es: /\b(?:emp(?:loyee)?\s*[\+\/&]\s*(?:spouse|sp)|ee\s*[\+\/&]\s*sp|es\b|emp\s*(?:\+\s*)?spouse)/i,
    ec: /\b(?:emp(?:loyee)?\s*[\+\/&]\s*(?:child(?:ren)?|ch)|ee\s*[\+\/&]\s*ch|ec\b|emp\s*(?:\+\s*)?child)/i,
    ef: /\b(?:family|emp(?:loyee)?\s*[\+\/&]\s*(?:family|fam)|ee\s*[\+\/&]\s*fam|ef\b)/i,
  };

  // Benefit label patterns for benefit rows in rate tables
  const benefitLabels = {
    deductibleIndividual: /\bdeductible\b(?!.*family)/i,
    deductibleFamily: /\bdeductible\b.*family/i,
    oopMaxIndividual: /\b(?:out[\s-]*of[\s-]*pocket|oop|moop)\b(?!.*family)/i,
    oopMaxFamily: /\b(?:out[\s-]*of[\s-]*pocket|oop|moop)\b.*family/i,
    copayPCP: /\b(?:pcp|primary\s*care|office\s*visit|doctor\s*visit|physician)\b/i,
    copaySpecialist: /\bspecialist\b/i,
    copayER: /\b(?:emergency|er\b)/i,
    copayUrgentCare: /\burgent\s*care\b/i,
    rxTier1: /\b(?:generic|tier\s*1)\b/i,
    rxTier2: /\b(?:preferred\s*brand|tier\s*2)\b/i,
    rxTier3: /\b(?:non[\s-]*preferred|tier\s*3|specialty)\b/i,
    coinsurance: /\bcoinsurance\b/i,
  };

  // Find rows with dollar amounts
  const dollarRowRe = /\$[\d,]+(?:\.\d{2})?/g;

  // Also track benefit-labeled rows for later assignment
  const benefitRows = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const amounts = (line.match(dollarRowRe) || []).map(s => parseMoney(s));

    // Check if this is a premium tier label row with dollar amounts
    let matchedTier = null;
    for (const [tier, re] of Object.entries(tierLabels)) {
      if (re.test(line)) { matchedTier = tier; break; }
    }

    // Check if this is a benefit-labeled row
    let matchedBenefit = null;
    if (!matchedTier) {
      for (const [field, re] of Object.entries(benefitLabels)) {
        if (re.test(line) && amounts.length > 0) { matchedBenefit = field; break; }
      }
    }

    if (!matchedTier && !matchedBenefit) continue;
    if (amounts.length === 0) continue;

    // Look backwards for a header row with plan names to create plan stubs
    if (!plans.length) {
      for (let j = Math.max(0, i - 15); j < i; j++) {
        const hLine = lines[j];
        const planNames = findPlanNamesInLine(hLine);
        if (planNames.length >= amounts.length) {
          for (let k = 0; k < amounts.length && k < planNames.length; k++) {
            const pn = planNames[k];
            plans.push({
              id: uuidv4(),
              carrier: null, planName: pn.name, planCode: null,
              networkType: pn.network || null, metalLevel: pn.metal || null,
              deductibleIndividual: null, deductibleFamily: null,
              oopMaxIndividual: null, oopMaxFamily: null,
              coinsurance: null, copayPCP: null, copaySpecialist: null,
              copayUrgentCare: null, copayER: null,
              rxDeductible: null, rxTier1: null, rxTier2: null, rxTier3: null,
              premiumEE: null, premiumES: null, premiumEC: null, premiumEF: null,
              effectiveDate: null, ratingArea: null, underwritingNotes: null,
              extractionConfidence: 0, sourceFile,
            });
          }
          break;
        }
      }
      // If still no plan stubs and we have amounts, create generic plans
      if (!plans.length && amounts.length > 0) {
        for (let k = 0; k < amounts.length; k++) {
          plans.push({
            id: uuidv4(),
            carrier: null, planName: `Plan ${k + 1}`, planCode: null,
            networkType: null, metalLevel: null,
            deductibleIndividual: null, deductibleFamily: null,
            oopMaxIndividual: null, oopMaxFamily: null,
            coinsurance: null, copayPCP: null, copaySpecialist: null,
            copayUrgentCare: null, copayER: null,
            rxDeductible: null, rxTier1: null, rxTier2: null, rxTier3: null,
            premiumEE: null, premiumES: null, premiumEC: null, premiumEF: null,
            effectiveDate: null, ratingArea: null, underwritingNotes: null,
            extractionConfidence: 0, sourceFile,
          });
        }
      }
    }

    // Assign premium amounts to the matching tier
    if (matchedTier) {
      const premKey = { ee: 'premiumEE', es: 'premiumES', ec: 'premiumEC', ef: 'premiumEF' }[matchedTier];
      for (let k = 0; k < amounts.length && k < plans.length; k++) {
        if (amounts[k] != null) plans[k][premKey] = amounts[k];
      }
    }

    // Assign benefit amounts to the matching benefit field
    if (matchedBenefit) {
      for (let k = 0; k < amounts.length && k < plans.length; k++) {
        if (amounts[k] != null && plans[k][matchedBenefit] == null) {
          plans[k][matchedBenefit] = amounts[k];
        }
      }
    }
  }

  // Fill in confidence and filter
  const result = plans.filter(p => {
    const kf = [p.carrier, p.planName, p.networkType, p.deductibleIndividual, p.oopMaxIndividual, p.copayPCP, p.premiumEE, p.premiumES, p.premiumEC, p.premiumEF];
    const found = kf.filter(v => v !== null && v !== undefined).length;
    p.extractionConfidence = Math.min(1, found / 6);
    return found >= 2;
  });
  return result;
}

function findPlanNamesInLine(line) {
  const plans = [];
  const PLAN_KW = /\b(HMO|PPO|EPO|HDHP|HSA|Gold|Silver|Bronze|Platinum|Plus|Select|Choice|Value|Basic|Premier|Standard|Preferred|Essential|Core|Navigate|Charter|Compass)\b/i;

  // ── Approach 1: Split by tabs or 2+ spaces (common in pdf-parse table output)
  const segments = line.split(/\t|\s{2,}/).map(s => s.trim()).filter(s => s.length > 3 && s.length < 80);
  if (segments.length >= 2) {
    const planSegs = segments.filter(s => PLAN_KW.test(s) && !/^\$|^\d+$/.test(s));
    if (planSegs.length >= 2) {
      for (const seg of planSegs) {
        const network = (seg.match(/\b(HMO|PPO|EPO|HDHP|HSA)\b/i) || [])[1];
        const metal = (seg.match(/\b(Gold|Silver|Bronze|Platinum)\b/i) || [])[1];
        plans.push({ name: seg, network: network ? network.toUpperCase() : null, metal: metal ? metal.charAt(0).toUpperCase() + metal.slice(1).toLowerCase() : null });
      }
      return plans;
    }
  }

  // ── Approach 2: Find all network/metal keyword positions and split around them
  const kwPositions = [];
  const kwRe = /\b(HMO|PPO|EPO|HDHP|HSA|Gold|Silver|Bronze|Platinum)\b/gi;
  let km;
  while ((km = kwRe.exec(line)) !== null) {
    kwPositions.push({ keyword: km[1], index: km.index, end: km.index + km[0].length });
  }
  if (kwPositions.length >= 2) {
    for (let i = 0; i < kwPositions.length; i++) {
      const start = i === 0 ? 0 : kwPositions[i - 1].end;
      let end = kwPositions[i].end;
      // Include trailing digits/spaces (plan codes like "500", "1500")
      const afterKw = line.substring(end);
      const trailer = afterKw.match(/^[\s]*[\d]+/);
      if (trailer) end += trailer[0].length;
      const name = line.substring(start, end).trim();
      if (name.length >= 4 && name.length < 80 && !/^\$/.test(name)) {
        const network = (name.match(/\b(HMO|PPO|EPO|HDHP|HSA)\b/i) || [])[1];
        const metal = (name.match(/\b(Gold|Silver|Bronze|Platinum)\b/i) || [])[1];
        plans.push({ name, network: network ? network.toUpperCase() : null, metal: metal ? metal.charAt(0).toUpperCase() + metal.slice(1).toLowerCase() : null });
      }
    }
    if (plans.length >= 2) return plans;
    plans.length = 0;
  }

  // ── Approach 3: Single (non-greedy) match — fallback for a header with 1 plan name
  const singleRe = /([A-Za-z][A-Za-z0-9\s\/\-&']{2,60}?(?:HMO|PPO|EPO|HDHP|HSA|Gold|Silver|Bronze|Platinum|Plus|Select|Choice|Value|Basic|Premier|Standard|Preferred|Essential|Core)(?:\s*\d{0,5})?)/gi;
  let sm;
  while ((sm = singleRe.exec(line)) !== null) {
    const name = sm[1].trim();
    if (name.length > 3 && name.length < 80) {
      const network = (name.match(/\b(HMO|PPO|EPO|HDHP|HSA)\b/i) || [])[1];
      const metal = (name.match(/\b(Gold|Silver|Bronze|Platinum)\b/i) || [])[1];
      plans.push({ name, network: network ? network.toUpperCase() : null, metal: metal ? metal.charAt(0).toUpperCase() + metal.slice(1).toLowerCase() : null });
    }
  }
  return plans;
}

// Look for data rows where a single line contains plan info + dollar amounts
function extractFromDataRows(lines, fullText, sourceFile) {
  const plans = [];
  // Pattern: line with a plan-name-like string followed by multiple $ amounts
  for (const line of lines) {
    const dollarAmounts = [];
    const dollarRe = /\$\s*([\d,]+(?:\.\d{1,2})?)/g;
    let dm;
    while ((dm = dollarRe.exec(line)) !== null) {
      dollarAmounts.push(parseMoney(dm[1]));
    }
    if (dollarAmounts.length < 2) continue; // Need at least 2 numeric values

    // Check if line contains a plan name
    const planNameMatch = line.match(/([A-Za-z][A-Za-z0-9\s\/\-&']*(?:HMO|PPO|EPO|HDHP|HSA|Gold|Silver|Bronze|Platinum|Plus|Choice|Select|Value|Preferred|Essential|Core|Standard|Basic|Premier)[A-Za-z0-9\s\/\-&']*)/i);
    if (!planNameMatch) continue;

    const planName = planNameMatch[1].trim();
    if (planName.length < 4) continue;

    const network = (planName.match(/\b(HMO|PPO|EPO|HDHP|HSA)\b/i) || [])[1];
    const metal = (planName.match(/\b(Gold|Silver|Bronze|Platinum)\b/i) || [])[1];

    // Heuristic assignment of dollar amounts:
    // Sort amounts to help guess which is premium vs deductible vs copay
    // Premiums are typically $200-$3000/month, deductibles $250-$10000, copays $10-$100, OOP $2000-$16000
    const sorted = [...dollarAmounts].filter(v => v != null);
    let premiumEE = null, deductibleIndividual = null, oopMaxIndividual = null, copayPCP = null;
    let premiumES = null, premiumEC = null, premiumEF = null;

    // Try to identify based on typical ranges and position
    const copays = sorted.filter(v => v >= 5 && v <= 100);
    const deductibles = sorted.filter(v => v >= 100 && v <= 10000);
    const oops = sorted.filter(v => v >= 2000 && v <= 20000);
    const premiums = sorted.filter(v => v >= 100 && v <= 5000);

    // If we have 4+ amounts that look like premiums (all in $100-$3000 range), treat as EE/ES/EC/EF
    const premLike = sorted.filter(v => v >= 50 && v <= 4000);
    if (premLike.length >= 3) {
      // Likely a premium row: EE, ES, EC, EF (ascending)
      const premSorted = [...premLike].sort((a, b) => a - b);
      premiumEE = premSorted[0] || null;
      premiumES = premSorted.length >= 2 ? premSorted[1] : null;
      premiumEC = premSorted.length >= 3 ? premSorted[2] : null;
      premiumEF = premSorted.length >= 4 ? premSorted[3] : null;
    } else {
      // Mixed types — try copay (smallest), premium, deductible, OOP (largest)
      if (copays.length > 0) copayPCP = Math.min(...copays);
      const remaining = sorted.filter(v => v !== copayPCP);
      if (remaining.length >= 1) premiumEE = remaining.find(v => v >= 100 && v <= 3000) || null;
      if (remaining.length >= 2) deductibleIndividual = remaining.find(v => v >= 500 && v <= 10000 && v !== premiumEE) || null;
      if (remaining.length >= 3) oopMaxIndividual = remaining.find(v => v >= 2000 && v <= 20000 && v !== premiumEE && v !== deductibleIndividual) || null;
    }

    const keyFields = [planName, network, metal, deductibleIndividual, oopMaxIndividual, copayPCP, premiumEE];
    const found = keyFields.filter(v => v !== null && v !== undefined).length;
    if (found < 2) continue;

    plans.push({
      id: uuidv4(),
      carrier: null, planName, planCode: null,
      networkType: network ? network.toUpperCase() : null,
      metalLevel: metal ? metal.charAt(0).toUpperCase() + metal.slice(1).toLowerCase() : null,
      deductibleIndividual, deductibleFamily: null,
      oopMaxIndividual, oopMaxFamily: null,
      coinsurance: null, copayPCP,
      copaySpecialist: null, copayUrgentCare: null, copayER: null,
      rxDeductible: null, rxTier1: null, rxTier2: null, rxTier3: null,
      premiumEE, premiumES, premiumEC, premiumEF,
      effectiveDate: null, ratingArea: null, underwritingNotes: null,
      extractionConfidence: Math.min(1, found / 6),
      sourceFile,
    });
  }
  return plans;
}

function splitIntoPlanBlocks(lines, fullText) {
  // Look for lines that look like plan headers — broadened patterns
  const headerPatterns = [
    // Lines that are just a plan name like "Blue Choice PPO Gold 500"
    /^[A-Z][A-Za-z0-9\s\/\-&']{3,60}\b(?:HMO|PPO|EPO|HDHP|HSA|Platinum|Gold|Silver|Bronze|Plus|Choice|Select|Value|Preferred|Essential|Core|Standard|Basic|Premier)\b[A-Za-z0-9\s\/\-&']{0,30}$/i,
    // "Plan Name: ..." or "Plan 1:" etc.
    /^(?:plan\s*(?:name|type|option|design)?\s*[:\-#]?\s*\d*\.?\s*)([A-Za-z].{3,60})$/i,
    // "Option A:", "Option 1:" headers
    /^(?:option|plan|benefit)\s*[A-Z1-9][:\-\s]/i,
    // SBC-style "Summary of Benefits and Coverage" header
    /^summary\s+of\s+benefits/i,
    // "Schedule of Benefits" header
    /^schedule\s+of\s+benefits/i,
    // "Benefit Highlights" header
    /^benefit\s+(?:highlights|summary|details)/i,
    // Lines that start with a carrier name followed by plan identifier
    /^(?:Anthem|Aetna|Cigna|United|Kaiser|BCBS|Blue\s*Cross|Humana|Oscar)[A-Za-z\s\/\-&']{4,60}(?:HMO|PPO|EPO|HDHP|HSA|Gold|Silver|Bronze|Platinum)/i,
  ];
  const planStartIndices = [];

  lines.forEach((line, i) => {
    if (line.length > 120) return; // skip very long lines
    for (const re of headerPatterns) {
      if (re.test(line)) {
        // Avoid marking consecutive header lines — require at least 5 lines gap
        if (planStartIndices.length === 0 || i - planStartIndices[planStartIndices.length - 1] >= 5) {
          planStartIndices.push(i);
        }
        break;
      }
    }
  });

  if (planStartIndices.length <= 1) {
    return [fullText]; // treat whole file as one block
  }

  const blocks = [];
  for (let b = 0; b < planStartIndices.length; b++) {
    const start = planStartIndices[b];
    const end = b + 1 < planStartIndices.length ? planStartIndices[b + 1] : lines.length;
    blocks.push(lines.slice(start, end).join('\n'));
  }
  return blocks;
}

function extractFieldsFromBlock(text, sourceFile) {
  const get = (patterns) => firstMatch(text, patterns);

  // Carrier
  const carrier = detectCarrier(text, sourceFile);

  // Plan name — broadened
  const planName = get([
    /(?:plan\s*name|plan\s*title|plan\s*option|plan\s*design)\s*[:\-]\s*([^\n]{3,70})/i,
    /([A-Za-z][A-Za-z0-9\s\/\-&']*(?:HMO|PPO|EPO|HDHP|HSA)[A-Za-z0-9\s\/\-&']{0,40})/i,
    /([A-Za-z][A-Za-z0-9\s\/\-&']*(?:Gold|Silver|Bronze|Platinum)[A-Za-z0-9\s\/\-&']{0,40})/i,
    /([A-Za-z][A-Za-z0-9\s\/\-&']*(?:Plus|Choice|Select|Value|Preferred|Essential|Core|Standard|Basic|Premier)[A-Za-z0-9\s\/\-&']{0,40})/i,
  ]);

  // Network type
  const networkRaw = get([/\b(HDHP|EPO|HMO|PPO|HSA|POS|OAP)\b/i]);
  const networkType = networkRaw ? networkRaw.toUpperCase() : null;

  // Metal level
  const metalRaw = get([/\b(Platinum|Gold|Silver|Bronze|Catastrophic)\b/i]);
  const metalLevel = metalRaw ? metalRaw.charAt(0).toUpperCase() + metalRaw.slice(1).toLowerCase() : null;

  // ── Deductibles — expanded patterns ───────────────────────────────────────
  // Handles: "Deductible: $1,500", "Deductible.....$1,500", "Individual Deductible $1,500",
  //          "Deductible $1,500 / $3,000", "In-Network Deductible: $1,500",
  //          "Medical Deductible $1,500", "$1,500 Deductible", etc.
  const deductibleIndividual = parseMoney(get([
    /individual\s+deductible\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /deductible\s*[:\-]?\s*individual\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /deductible\s*(?:\(individual\))?\s*[:\-·.…]*\s*\$?([\d,]+)\s*(?:individual|per\s*person|person|member|single)/i,
    /deductible\s*(?:\(individual\))?\s*[:\-·.…]*\s*\$?([\d,]+)\s*[\/|]\s*\$?[\d,]+/i,
    /(?:in[\s-]?network\s+)?(?:medical\s+)?deductible\s*(?:\((?:individual|in[\s-]?network)\))?\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /\$\s*([\d,]+)\s*(?:individual\s+)?deductible/i,
    /deductible[^$\n]{0,50}\$\s*([\d,]+)/i,
    /ded(?:uctible)?\s*[:\-·.…]*\s*\$?([\d,]+)/i,
  ]));

  const deductibleFamily = parseMoney(get([
    /family\s+deductible\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /deductible\s*[:\-]?\s*family\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    // "$X / $Y" pair — capture Y (family). Must have $ before both X and Y.
    /deductible[^$\n]*\$\s*[\d,]+(?:\.\d+)?\s*[\/|]\s*\$\s*([\d,]+)/i,
    // Handle "(Individual/Family): $X / $Y"
    /deductible\s*\([^)]*\)\s*[:\-·.…]*\s*\$\s*[\d,]+(?:\.\d+)?\s*[\/|]\s*\$\s*([\d,]+)/i,
  ]));

  // ── Out-of-Pocket Max — expanded patterns ────────────────────────────────
  // Handles: "OOP Max: $6,000", "Out-of-Pocket Maximum.....$6,000",
  //          "Maximum Out of Pocket $6,000 / $12,000", "MOOP: $6,000",
  //          "Individual OOP Max $6,000", "$6,000 Out of Pocket Maximum"
  const oopMaxIndividual = parseMoney(get([
    /individual\s+(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*[:\-]?\s*(?:individual|in[\s-]?network)\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*(?:\(individual\))?\s*[:\-·.…]*\s*\$?([\d,]+)\s*(?:individual|per\s*person|person|member|single)/i,
    /(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*(?:\(individual\))?\s*[:\-·.…]*\s*\$?([\d,]+)\s*[\/|]\s*\$?[\d,]+/i,
    /(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*(?:\([^)]*\))?\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /(?:max(?:imum)?\s*)?out[\s-]*of[\s-]*pocket\s*(?:max(?:imum)?)?\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /\$\s*([\d,]+)\s*(?:individual\s+)?(?:out[\s-]*of[\s-]*pocket|oop)/i,
    /moop\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    /oop[^$\n]{0,50}\$\s*([\d,]+)/i,
  ]));

  const oopMaxFamily = parseMoney(get([
    /family\s+(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*[:\-·.…]*\s*\$?([\d,]+)/i,
    // "$X / $Y" pair — capture Y (family). Must have $ before both X and Y.
    /(?:out[\s-]*of[\s-]*pocket|oop)[^$\n]*\$\s*[\d,]+(?:\.\d+)?\s*[\/|]\s*\$\s*([\d,]+)/i,
    // Handle "(Individual/Family): $X / $Y"
    /(?:out[\s-]*of[\s-]*pocket|oop)\s*(?:max(?:imum)?)?\s*\([^)]*\)\s*[:\-·.…]*\s*\$\s*[\d,]+(?:\.\d+)?\s*[\/|]\s*\$\s*([\d,]+)/i,
  ]));

  // Coinsurance
  const coinsurance = get([
    /coinsurance\s*[:\-·.…]*\s*(\d+%(?:\s*[\/\-]\s*\d+%)?)/i,
    /(\d+%)\s*(?:co[\s-]?insurance|after\s+deductible)/i,
    /(?:you\s+pay|member\s+pays?|plan\s+pays?)\s*[:\-]?\s*(\d+%)/i,
  ]);

  // ── Copays — expanded patterns ──────────────────────────────────────────
  // Handles: "PCP Copay: $25", "Office Visit.....$25", "Primary Care $25 copay",
  //          "$25 PCP", "Doctor Visit Copay $25"
  const copayPCP = parseMoney(get([
    /(?:pcp|primary\s*care(?:\s*physician)?|office\s*visit|doctor\s*visit|physician\s*visit)\s*(?:copay?|co[\s-]?pay(?:ment)?)\s*[:\-·.…]*\s*\$?([\d]+)/i,
    /(?:copay?|co[\s-]?pay)\s*[:\-]?\s*\$?([\d]+)\s*(?:pcp|primary\s*care|office\s*visit)/i,
    /(?:pcp|office\s*visit|primary\s*care)\s*[:\-·.…]*\s*\$?([\d]+)/i,
    /\$\s*([\d]+)\s*(?:pcp|office\s*visit|primary\s*care)/i,
    /(?:pcp|primary\s*care|office\s*visit)[^$\n]{0,30}\$\s*([\d]+)/i,
  ]));

  const copaySpecialist = parseMoney(get([
    /specialist\s*(?:copay?|co[\s-]?pay(?:ment)?|visit)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    /(?:copay?|co[\s-]?pay)\s*[:\-]?\s*\$?([\d]+)\s*(?:specialist)/i,
    /specialist[^$\n]{0,30}\$\s*([\d]+)/i,
  ]));

  const copayUrgentCare = parseMoney(get([
    /urgent\s*care\s*(?:copay?|co[\s-]?pay(?:ment)?)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    /urgent\s*care[^$\n]{0,30}\$\s*([\d]+)/i,
  ]));

  const copayER = parseMoney(get([
    /(?:emergency\s*(?:room|dept|department|services?)|er)\s*(?:copay?|co[\s-]?pay(?:ment)?)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    /(?:copay?|co[\s-]?pay)\s*[:\-]?\s*\$?([\d]+)\s*(?:emergency|er\b)/i,
    /emergency[^$\n]{0,30}\$\s*([\d]+)/i,
  ]));

  // ── Rx / Pharmacy ──────────────────────────────────────────────────────
  const rxDeductible = parseMoney(get([
    /(?:rx|pharmacy|drug|prescription)\s*deductible\s*[:\-·.…]*\s*\$?([\d,]+)/i,
  ]));
  const rxTier1 = parseMoney(get([
    /(?:generic|tier\s*(?:1|i)\b|tier1)\s*(?:rx|drug|copay?|co[\s-]?pay)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    /(?:generic|tier\s*1)[^$\n]{0,30}\$\s*([\d]+)/i,
  ]));
  const rxTier2 = parseMoney(get([
    /(?:preferred\s*brand|tier\s*(?:2|ii)\b|tier2)\s*(?:rx|drug|copay?|co[\s-]?pay)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    /(?:preferred\s*brand|tier\s*2)[^$\n]{0,30}\$\s*([\d]+)/i,
  ]));
  const rxTier3 = parseMoney(get([
    /(?:non[\s-]?preferred\s*brand|tier\s*(?:3|iii)\b|tier3|specialty)\s*(?:rx|drug|copay?|co[\s-]?pay)?\s*[:\-·.…]*\s*\$?([\d]+)/i,
    /(?:non[\s-]?preferred|tier\s*3|specialty)[^$\n]{0,30}\$\s*([\d]+)/i,
  ]));

  // ── Premiums — expanded patterns ───────────────────────────────────────
  // Handles: "Employee Only: $500.00", "EE.....$500.00", "EE Premium $500.00",
  //          "Single Rate: $500.00", "Employee Only $500.00"
  const premiumEE = parseMoney(get([
    /(?:employee\s*only|ee\b|emp\s*only|single)\s*(?:monthly\s*)?(?:premium|rate|cost)?\s*[:\-·.…]*\s*\$?([\d,]+\.?\d*)/i,
    /(?:premium|rate|cost)\s*[:\-]?\s*(?:employee\s*only|ee)\s*[:\-·.…]*\s*\$?([\d,]+\.?\d*)/i,
    /(?:employee\s*only|ee|single)\s+\$?([\d,]+\.\d{2})/i,
    /(?:employee\s*only|ee|single)[^$\n]{0,30}\$\s*([\d,]+\.\d{2})/i,
  ]));

  const premiumES = parseMoney(get([
    /(?:emp(?:loyee)?\s*[\+\/&]\s*(?:spouse|sp)|emp\s*spouse|ee\s*[\+\/&]\s*sp|es\b|employee\s*[\+\/&\s]\s*spouse|emp(?:loyee)?\/spouse)\s*(?:monthly\s*)?(?:premium|rate|cost)?\s*[:\-·.…]*\s*\$?([\d,]+\.?\d*)/i,
    /(?:employee\s*[\+\/&]\s*spouse|emp\+sp|ee\+sp|es)\s+\$?([\d,]+\.\d{2})/i,
    /(?:emp(?:loyee)?\s*[\+\/&]\s*spouse|es)[^$\n]{0,30}\$\s*([\d,]+\.\d{2})/i,
  ]));

  const premiumEC = parseMoney(get([
    /(?:emp(?:loyee)?\s*[\+\/&]\s*(?:child(?:ren)?|ch|dep)|ee\s*[\+\/&]\s*ch|ec\b|employee\s*[\+\/&\s]\s*child(?:ren)?|emp(?:loyee)?\/child(?:ren)?)\s*(?:monthly\s*)?(?:premium|rate|cost)?\s*[:\-·.…]*\s*\$?([\d,]+\.?\d*)/i,
    /(?:employee\s*[\+\/&]\s*child(?:ren)?|emp\+ch|ee\+ch|ec)\s+\$?([\d,]+\.\d{2})/i,
    /(?:emp(?:loyee)?\s*[\+\/&]\s*child|ec)[^$\n]{0,30}\$\s*([\d,]+\.\d{2})/i,
  ]));

  const premiumEF = parseMoney(get([
    /(?:family|emp(?:loyee)?\s*[\+\/&]\s*(?:family|fam)|ee\s*[\+\/&]\s*fam|ef\b|employee\s*[\+\/&\s]\s*family|emp(?:loyee)?\/family)\s*(?:monthly\s*)?(?:premium|rate|cost)?\s*[:\-·.…]*\s*\$?([\d,]+\.?\d*)/i,
    /(?:employee\s*[\+\/&]\s*family|family|emp\+fam|ee\+fam|ef)\s+\$?([\d,]+\.\d{2})/i,
    /(?:family|ef)[^$\n]{0,30}\$\s*([\d,]+\.\d{2})/i,
  ]));

  // Effective date
  const effectiveDate = get([
    /effective\s*(?:date)?\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i,
    /effective\s*[:\-]?\s*(\w+ \d{1,2},?\s*\d{4})/i,
    /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})\s*(?:effective|start|begin)/i,
  ]);

  // Plan code
  const planCode = get([/plan\s*(?:code|id|#|number)\s*[:\-]?\s*([A-Z0-9\-]{3,20})/i]);

  // Rating area
  const ratingArea = get([/rating\s*(?:area|region|zone)\s*[:\-]?\s*([A-Za-z0-9 \-]+)/i]);

  // Underwriting notes
  const underwritingNotes = get([/underwriting\s*(?:notes?|class|tier)\s*[:\-]?\s*([^\n]{1,100})/i]);

  // Confidence: how many key fields were found
  const keyFields = [
    carrier, planName, networkType, metalLevel, deductibleIndividual,
    oopMaxIndividual, copayPCP, premiumEE, premiumES, premiumEC, premiumEF,
  ];
  const found = keyFields.filter(v => v !== null && v !== undefined).length;
  const extractionConfidence = Math.min(1, found / 6);

  // Only create a plan if we found at least 2 fields
  if (found < 2) return null;

  return {
    id: uuidv4(),
    carrier: carrier || null,
    planName: planName || null,
    planCode: planCode || null,
    networkType: networkType || null,
    metalLevel: metalLevel || null,
    deductibleIndividual,
    deductibleFamily: deductibleFamily || null,
    oopMaxIndividual,
    oopMaxFamily: oopMaxFamily || null,
    coinsurance: coinsurance || null,
    copayPCP,
    copaySpecialist: copaySpecialist || null,
    copayUrgentCare: copayUrgentCare || null,
    copayER: copayER || null,
    rxDeductible: rxDeductible || null,
    rxTier1: rxTier1 || null,
    rxTier2: rxTier2 || null,
    rxTier3: rxTier3 || null,
    premiumEE,
    premiumES: premiumES || null,
    premiumEC: premiumEC || null,
    premiumEF: premiumEF || null,
    effectiveDate: effectiveDate || null,
    ratingArea: ratingArea || null,
    underwritingNotes: underwritingNotes || null,
    extractionConfidence,
    sourceFile,
  };
}

async function extractPlansFromExcel(buffer, sourceFile) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const plans = [];

  workbook.eachSheet(sheet => {
    // Collect all rows as arrays of values
    const aoa = [];
    sheet.eachRow({ includeEmpty: false }, row => {
      const rowVals = [];
      row.eachCell({ includeEmpty: true }, cell => {
        const v = cell.value;
        if (v === null || v === undefined) {
          rowVals.push(null);
        } else if (typeof v === 'object' && v.result !== undefined) {
          // formula cell
          rowVals.push(v.result != null ? String(v.result) : null);
        } else if (typeof v === 'object' && v.text !== undefined) {
          // rich text
          rowVals.push(v.text);
        } else {
          rowVals.push(String(v));
        }
      });
      aoa.push(rowVals);
    });

    if (aoa.length === 0) return;

    // Check first row for column headers matching benefit fields
    const firstRow = aoa[0] || [];
    const hasHeaders = firstRow.some(
      h => h && /plan|carrier|network|deductible|premium|copay|oop/i.test(String(h))
    );

    if (hasHeaders && aoa.length > 1) {
      // Row-per-plan: first row = headers
      const headers = firstRow.map(h => (h ? String(h) : ''));
      for (let i = 1; i < aoa.length; i++) {
        const rowObj = {};
        headers.forEach((h, ci) => {
          rowObj[h] = aoa[i][ci] !== undefined ? aoa[i][ci] : null;
        });
        const plan = buildPlanFromRow(rowObj, sourceFile);
        if (plan) plans.push(plan);
      }
    } else {
      // Key-value layout: column A = labels, subsequent cols = plan values
      const plansByCol = buildPlansFromKeyValueSheet(aoa, sourceFile);
      plans.push(...plansByCol);
    }
  });

  return plans;
}

function buildPlanFromRow(row, sourceFile) {
  const find = (...keys) => {
    for (const k of keys) {
      for (const rk of Object.keys(row)) {
        if (new RegExp(k, 'i').test(rk) && row[rk] !== null && row[rk] !== '') {
          return String(row[rk]);
        }
      }
    }
    return null;
  };

  const carrier = find('carrier', 'insurer', 'insurance company');
  const planName = find('plan name', 'plan title', 'plan');
  const networkType = find('network', 'type', 'hmo|ppo|epo|hdhp');
  const metalLevel = find('metal', 'tier level');
  const deductibleIndividual = parseMoney(find('ind.*deductible', 'deductible.*ind', 'deductible'));
  const deductibleFamily = parseMoney(find('fam.*deductible', 'deductible.*fam'));
  const oopMaxIndividual = parseMoney(find('ind.*oop', 'oop.*ind', 'out.*of.*pocket.*ind', 'oop max', 'oop'));
  const oopMaxFamily = parseMoney(find('fam.*oop', 'oop.*fam', 'out.*of.*pocket.*fam'));
  const copayPCP = parseMoney(find('pcp', 'primary care', 'office visit'));
  const copaySpecialist = parseMoney(find('specialist'));
  const copayER = parseMoney(find('emergency', 'er copay', '\\ber\\b'));
  const copayUrgentCare = parseMoney(find('urgent'));
  const premiumEE = parseMoney(find('ee.*premium', 'premium.*ee', 'employee only', '\\bee\\b'));
  const premiumES = parseMoney(find('es.*premium', 'premium.*es', 'emp.*spouse', '\\bes\\b'));
  const premiumEC = parseMoney(find('ec.*premium', 'premium.*ec', 'emp.*child', '\\bec\\b'));
  const premiumEF = parseMoney(find('ef.*premium', 'premium.*ef', 'family', '\\bef\\b'));
  const coinsurance = find('coinsurance');
  const rxTier1 = parseMoney(find('generic', 'tier.?1', 'rx.*1'));
  const rxTier2 = parseMoney(find('preferred brand', 'tier.?2', 'rx.*2'));
  const rxTier3 = parseMoney(find('non.*preferred', 'tier.?3', 'rx.*3'));
  const effectiveDate = find('effective date', 'effective');
  const planCode = find('plan code', 'plan id', 'plan #');
  const ratingArea = find('rating area', 'region', 'zone');

  const keyFields = [carrier, planName, networkType, deductibleIndividual, oopMaxIndividual, copayPCP, premiumEE];
  const found = keyFields.filter(v => v !== null).length;
  if (found < 2) return null;

  const extractionConfidence = Math.min(1, found / 6);
  let netType = networkType;
  if (netType) {
    const m = netType.match(/\b(HMO|PPO|EPO|HDHP|HSA)\b/i);
    netType = m ? m[1].toUpperCase() : netType;
  }
  let metal = metalLevel;
  if (metal) {
    const m = metal.match(/\b(Platinum|Gold|Silver|Bronze)\b/i);
    metal = m ? m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase() : metal;
  }

  return {
    id: uuidv4(),
    carrier: carrier || null,
    planName: planName || null,
    planCode: planCode || null,
    networkType: netType || null,
    metalLevel: metal || null,
    deductibleIndividual,
    deductibleFamily,
    oopMaxIndividual,
    oopMaxFamily,
    coinsurance: coinsurance || null,
    copayPCP,
    copaySpecialist,
    copayUrgentCare,
    copayER,
    rxDeductible: null,
    rxTier1,
    rxTier2,
    rxTier3,
    premiumEE,
    premiumES,
    premiumEC,
    premiumEF,
    effectiveDate: effectiveDate || null,
    ratingArea: ratingArea || null,
    underwritingNotes: null,
    extractionConfidence,
    sourceFile,
  };
}

function buildPlansFromKeyValueSheet(aoa, sourceFile) {
  if (!aoa || aoa.length === 0) return [];
  // First row may be headers; first column is field labels
  // Each subsequent column is a plan
  const numCols = aoa[0] ? aoa[0].length : 0;
  if (numCols < 2) return [];

  const plans = [];
  for (let col = 1; col < numCols; col++) {
    const rowObj = {};
    for (const row of aoa) {
      const key = row[0] ? String(row[0]) : null;
      const val = row[col] !== null && row[col] !== undefined ? String(row[col]) : null;
      if (key) rowObj[key] = val;
    }
    const plan = buildPlanFromRow(rowObj, sourceFile);
    if (plan) plans.push(plan);
  }
  return plans;
}

// ── Parse endpoint ────────────────────────────────────────────────────────────
app.post('/parse', async (req, res) => {
  try {
    const { caseId } = req.body;
    if (!caseId) return res.status(400).json({ error: 'caseId is required' });
    const caseData = caseStore.get(caseId);
    if (!caseData) return res.status(404).json({ error: 'Case not found' });

    const allPlans = [];
    const warnings = [];

    for (const file of caseData.files) {
      const ext = file.originalname.split('.').pop().toLowerCase();
      try {
        if (ext === 'pdf') {
          const data = await pdfParse(file.buffer);
          console.log(`[PARSE] PDF "${file.originalname}" — extracted ${data.text.length} chars, ${data.text.split('\n').length} lines`);
          console.log(`[PARSE] First 2000 chars:\n${data.text.substring(0, 2000)}`);
          const plans = await extractPlanFromText(data.text, file.originalname);
          console.log(`[PARSE] Extracted ${plans.length} plans from "${file.originalname}":`, plans.map(p => ({ name: p.planName, ded: p.deductibleIndividual, oop: p.oopMaxIndividual, pcp: p.copayPCP, eeP: p.premiumEE })));
          if (plans.length === 0) {
            warnings.push(`No plans extracted from ${file.originalname} — check file formatting`);
          }
          allPlans.push(...plans);
        } else if (['xlsx', 'xls', 'csv'].includes(ext)) {
          const plans = await extractPlansFromExcel(file.buffer, file.originalname);
          if (plans.length === 0) {
            warnings.push(`No plans extracted from ${file.originalname} — check column headers`);
          }
          allPlans.push(...plans);
        } else {
          warnings.push(`Skipped unsupported file: ${file.originalname}`);
        }
      } catch (parseErr) {
        warnings.push(`Error parsing ${file.originalname}: ${parseErr.message}`);
      }
    }

    caseData.plans = allPlans;
    caseStore.set(caseId, caseData);

    // Include a raw text preview for debugging extraction issues
    const rawTextPreviews = [];
    for (const file of caseData.files) {
      const ext = file.originalname.split('.').pop().toLowerCase();
      if (ext === 'pdf') {
        try {
          const data = await pdfParse(file.buffer);
          rawTextPreviews.push({ file: file.originalname, preview: data.text.substring(0, 5000), totalLength: data.text.length });
        } catch (e) { /* skip */ }
      }
    }

    res.json({ caseId, plans: allPlans, census: caseData.census || {}, warnings, rawTextPreviews });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Recommend endpoint ────────────────────────────────────────────────────────
function scorePlans(plans, census, contribution) {
  const { ee = 0, es = 0, ec = 0, ef = 0 } = census;
  const totalEnrolled = ee + es + ec + ef;

  const scored = plans.map(plan => {
    // ── Premium Cost (weighted by enrollment) ──
    let monthlyTotal = 0;
    if (plan.premiumEE) monthlyTotal += plan.premiumEE * ee;
    if (plan.premiumES) monthlyTotal += plan.premiumES * es;
    if (plan.premiumEC) monthlyTotal += plan.premiumEC * ec;
    if (plan.premiumEF) monthlyTotal += plan.premiumEF * ef;

    // ── Risk Protection ──
    // When benefits are null (unknown), use neutral score (0.5), NOT perfect (1.0)
    const dedKnown = plan.deductibleIndividual != null;
    const oopKnown = plan.oopMaxIndividual != null;
    let riskScore = 0.5; // neutral default for unknown
    if (dedKnown || oopKnown) {
      const ded = plan.deductibleIndividual || 0;
      const oop = plan.oopMaxIndividual || 0;
      // $0 combined = 1.0, $15,000+ combined = 0
      riskScore = Math.max(0, 1 - (ded + oop) / 15000);
    }

    // ── Copay Usability ──
    let copayUsability = 0.5; // neutral default for unknown
    if (plan.copayPCP != null) {
      const copayScore = Math.max(0, 1 - plan.copayPCP / 100);
      // Bonus for copay-first plans (no deductible before copay kicks in)
      const copayFirst = plan.deductibleIndividual === 0 || (plan.copayPCP > 0 && plan.deductibleIndividual == null);
      copayUsability = copayFirst ? Math.min(1, copayScore + 0.15) : copayScore;
    }

    // ── Network Preference ──
    // PPO plans get a significant boost: broader networks, out-of-network coverage,
    // no referral requirements — clients strongly prefer these
    const networkScores = { PPO: 1.0, EPO: 0.7, POS: 0.7, HMO: 0.5, HDHP: 0.4, HSA: 0.4 };
    const networkScore = networkScores[(plan.networkType || '').toUpperCase()] || 0.5;

    // ── Metal Level Value ──
    // Higher metal levels = richer benefits, better day-to-day coverage
    const metalScores = { Platinum: 1.0, Gold: 0.85, Silver: 0.6, Bronze: 0.3, Catastrophic: 0.1 };
    const metalScore = metalScores[plan.metalLevel] || 0.5;

    const contributionSummary = calculatePlanCostShares(plan, census, contribution);

    return {
      ...plan,
      _monthlyTotal: monthlyTotal,
      _riskScore: riskScore,
      _copayUsability: copayUsability,
      _networkScore: networkScore,
      _metalScore: metalScore,
      _contributionSummary: contributionSummary,
    };
  });

  // Normalize premium efficiency across plans (relative ranking)
  const premiums = scored.map(p => p._monthlyTotal).filter(v => v > 0);
  const maxPremium = premiums.length > 0 ? Math.max(...premiums) : 1;
  const minPremium = premiums.length > 0 ? Math.min(...premiums) : 0;
  const premRange = maxPremium - minPremium || 1;

  // Check if benefit data is available across plans
  const hasBenefitData = scored.some(p => p.deductibleIndividual != null || p.oopMaxIndividual != null);

  const result = scored.map(plan => {
    const premEfficiency = plan._monthlyTotal > 0
      ? Math.max(0, 1 - (plan._monthlyTotal - minPremium) / premRange)
      : 0.5; // unknown premium gets middle score

    // Adaptive weighting: when benefits are unknown, rely more on premium + network + metal
    // When benefits are known, use a more balanced scoring
    let totalScore;
    if (hasBenefitData) {
      // Full scoring with benefit data available
      totalScore =
        premEfficiency     * 0.30 +   // Premium cost (reduced from 40%)
        plan._riskScore    * 0.20 +   // Deductible + OOP protection
        plan._copayUsability * 0.15 + // Copay accessibility
        plan._networkScore * 0.20 +   // Network breadth (PPO boost)
        plan._metalScore   * 0.15;    // Metal level richness
    } else {
      // Rate-only scoring: no benefit data, weight network/metal more heavily
      totalScore =
        premEfficiency     * 0.30 +   // Premium cost
        plan._networkScore * 0.30 +   // Network breadth (PPO gets big boost)
        plan._metalScore   * 0.30 +   // Metal level as proxy for benefit richness
        plan._copayUsability * 0.05 + // Copay (likely unknown)
        plan._riskScore    * 0.05;    // Risk (likely unknown)
    }

    const reasons = [];
    if (premEfficiency >= 0.7) reasons.push('competitive premium costs');
    if (plan._riskScore >= 0.7 && (plan.deductibleIndividual != null)) reasons.push('strong risk protection (low deductible + OOP max)');
    if (plan._copayUsability >= 0.7 && plan.copayPCP != null) reasons.push('excellent copay accessibility');
    if (plan._networkScore >= 0.9) reasons.push('broad PPO network — no referrals required, out-of-network coverage');
    if (plan._metalScore >= 0.8) reasons.push('rich benefit design (' + (plan.metalLevel || 'high tier') + ')');
    if (plan._networkScore >= 0.7 && plan._networkScore < 0.9) reasons.push('good network flexibility');
    if (premEfficiency >= 0.4 && premEfficiency < 0.7 && plan._networkScore >= 0.9) reasons.push('premium value for PPO-level access');
    if (reasons.length === 0) reasons.push('balanced overall value');

    const whyRecommended = `This plan scores well due to ${reasons.slice(0, 3).join(', ')}.`;

    return {
      ...plan,
      premiumEfficiencyScore: Math.round(premEfficiency * 100) / 100,
      riskProtectionScore: Math.round(plan._riskScore * 100) / 100,
      copayUsabilityScore: Math.round(plan._copayUsability * 100) / 100,
      networkScore: Math.round(plan._networkScore * 100) / 100,
      metalScore: Math.round(plan._metalScore * 100) / 100,
      totalScore: Math.round(totalScore * 1000) / 1000,
      monthlyTotalCost: Math.round(plan._monthlyTotal * 100) / 100,
      employerMonthlyCost: Math.round(plan._contributionSummary.employerMonthlyTotal * 100) / 100,
      employeeMonthlyCost: Math.round(plan._contributionSummary.employeeMonthlyTotal * 100) / 100,
      employerPerPayCost: Math.round(plan._contributionSummary.employerPerPayTotal * 100) / 100,
      employeePerPayCost: Math.round(plan._contributionSummary.employeePerPayTotal * 100) / 100,
      payrollFrequency: plan._contributionSummary.payrollFrequency,
      contributionBreakdown: plan._contributionSummary.byTier,
      whyRecommended,
    };
  });

  return result.sort((a, b) => b.totalScore - a.totalScore);
}

app.post('/recommend', (req, res) => {
  try {
    const { caseId, census, contribution } = req.body;
    if (!caseId) return res.status(400).json({ error: 'caseId is required' });
    const caseData = caseStore.get(caseId);
    if (!caseData) return res.status(404).json({ error: 'Case not found' });

    if (census) caseData.census = census;
    if (contribution) caseData.contribution = normalizeContributionConfig(contribution);
    if (!caseData.contribution) caseData.contribution = defaultContributionConfig();
    const plans = caseData.plans || [];
    if (plans.length === 0) return res.status(400).json({ error: 'No plans to score — run /parse first' });

    const allScored = scorePlans(plans, caseData.census || {}, caseData.contribution);

    // Pick top 3, but always ensure the lowest-cost plan is included
    let top = allScored.slice(0, 3);

    // Find lowest-cost plan (by EE premium as baseline when no census, else weighted total)
    const plansWithPremium = allScored.filter(p => p.monthlyTotalCost > 0 || (p.premiumEE != null && p.premiumEE > 0));
    let lowestCost = null;
    if (plansWithPremium.length > 0) {
      lowestCost = plansWithPremium.reduce((best, p) => {
        const costA = best.monthlyTotalCost > 0 ? best.monthlyTotalCost : (best.premiumEE || Infinity);
        const costB = p.monthlyTotalCost > 0 ? p.monthlyTotalCost : (p.premiumEE || Infinity);
        return costB < costA ? p : best;
      });
    }

    // Tag which plans are already in top 3
    const topIds = new Set(top.map(p => p.planName + '|' + p.carrier));
    const lowestCostId = lowestCost ? (lowestCost.planName + '|' + lowestCost.carrier) : null;

    if (lowestCost && !topIds.has(lowestCostId)) {
      // Replace #3 with lowest cost; it becomes the 3rd tile
      top = top.slice(0, 2);
      top.push(lowestCost);
    }

    // Assign ranks and labels
    const recommendations = top.map((p, i) => {
      const isLowest = lowestCost && (p.planName + '|' + p.carrier) === lowestCostId;
      return {
        rank: i + 1,
        ...p,
        recommendationLabel: isLowest ? 'Lowest Cost Option' : (i === 0 ? 'Top Pick' : i === 1 ? 'Runner-Up' : 'Strong Alternative'),
      };
    });

    caseData.recommendations = {
      recommendations,
      allPlans: allScored,
      contribution: caseData.contribution,
    };
    caseStore.set(caseId, caseData);

    res.json({ caseId, recommendations, allPlans: allScored, contribution: caseData.contribution });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Export PPTX ───────────────────────────────────────────────────────────────
app.post('/export/pptx', async (req, res) => {
  try {
    const { caseId, clientName = 'Client', effectiveDate = '', contribution, selectedPlanIndices } = req.body;
    if (!caseId) return res.status(400).json({ error: 'caseId is required' });
    const caseData = caseStore.get(caseId);
    if (!caseData) return res.status(404).json({ error: 'Case not found' });

    const plans = caseData.plans || [];
    const recData = caseData.recommendations || {};
    const contributionConfig = normalizeContributionConfig(contribution || recData.contribution || caseData.contribution || defaultContributionConfig());
    caseData.contribution = contributionConfig;

    // If account manager selected specific plans, use those; otherwise fall back to recommendations
    const useManualSelection = Array.isArray(selectedPlanIndices) && selectedPlanIndices.length > 0;
    let exportPlans;
    if (useManualSelection) {
      exportPlans = selectedPlanIndices
        .filter(i => i >= 0 && i < plans.length)
        .slice(0, 5)
        .map((idx, rank) => {
          const plan = plans[idx];
          const shares = calculatePlanCostShares(plan, caseData.census || {}, contributionConfig);
          return {
            ...plan,
            rank: rank + 1,
            recommendationLabel: `Selected Plan ${rank + 1}`,
            employerMonthlyCost: Math.round(shares.employerMonthlyTotal * 100) / 100,
            employeeMonthlyCost: Math.round(shares.employeeMonthlyTotal * 100) / 100,
            employerPerPayCost: Math.round(shares.employerPerPayTotal * 100) / 100,
            employeePerPayCost: Math.round(shares.employeePerPayTotal * 100) / 100,
            monthlyTotalCost: Math.round((shares.employerMonthlyTotal + shares.employeeMonthlyTotal) * 100) / 100,
            payrollFrequency: shares.payrollFrequency,
            contributionBreakdown: shares.byTier,
          };
        });
      console.log(`[PPTX] Using ${exportPlans.length} manually selected plans`);
    } else {
      exportPlans = (recData.recommendations || plans.slice(0, 3).map((p, i) => ({ rank: i + 1, ...p }))).map(plan => {
        if (typeof plan.employerPerPayCost === 'number' && typeof plan.employeePerPayCost === 'number') return plan;
        const shares = calculatePlanCostShares(plan, caseData.census || {}, contributionConfig);
        return {
          ...plan,
          employerMonthlyCost: Math.round(shares.employerMonthlyTotal * 100) / 100,
          employeeMonthlyCost: Math.round(shares.employeeMonthlyTotal * 100) / 100,
          employerPerPayCost: Math.round(shares.employerPerPayTotal * 100) / 100,
          employeePerPayCost: Math.round(shares.employeePerPayTotal * 100) / 100,
          monthlyTotalCost: plan.monthlyTotalCost || Math.round((shares.employerMonthlyTotal + shares.employeeMonthlyTotal) * 100) / 100,
          payrollFrequency: shares.payrollFrequency,
          contributionBreakdown: shares.byTier,
        };
      });
      console.log(`[PPTX] Using ${exportPlans.length} recommended plans`);
    }
    const recommendations = exportPlans;
    const census = caseData.census || {};

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.theme = { headFontFace: 'Calibri', bodyFontFace: 'Calibri' };

    const PRIMARY = '1e3a5f';
    const ACCENT = '2e86de';
    const WHITE = 'FFFFFF';
    const LIGHT = 'f4f6f9';
    const TEXT_DARK = '2c3e50';
    const GOLD = 'd4af37';
    const SILVER = '9e9e9e';
    const BRONZE = 'cd7f32';
    const GREEN = '27ae60';
    const MUTED = '7f8c8d';
    const ORANGE = 'e67e22';

    const fmtDol = v => v != null ? `$${Number(v).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : '—';
    const fmtInt = v => v != null ? `$${Number(v).toLocaleString()}` : '—';

    const addSlideHeader = (slide, title, subtitle) => {
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: '100%', h: 1.0,
        fill: { color: PRIMARY },
      });
      slide.addText(title, {
        x: 0.5, y: 0.08, w: '85%', h: subtitle ? 0.5 : 0.8,
        fontSize: 22, bold: true, color: WHITE, fontFace: 'Calibri',
      });
      if (subtitle) {
        slide.addText(subtitle, {
          x: 0.5, y: 0.52, w: '85%', h: 0.4,
          fontSize: 13, color: 'a8d4f5', fontFace: 'Calibri',
        });
      }
      // Accent line under header
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 1.0, w: '100%', h: 0.04,
        fill: { color: ACCENT },
      });
    };

    const addFooter = (slide, text) => {
      slide.addText(text || `${clientName} — Benefits Analysis`, {
        x: 0.3, y: 7.15, w: 9.4, h: 0.3,
        fontSize: 8, color: MUTED, italic: true, align: 'right',
      });
    };

    // ── Slide 1: Title ──
    const s1 = pptx.addSlide();
    s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: PRIMARY } });
    // Accent stripe
    s1.addShape(pptx.ShapeType.rect, { x: 0, y: 3.4, w: '100%', h: 0.06, fill: { color: ACCENT } });
    s1.addText('Benefits Plan Analysis', {
      x: 0.5, y: 1.2, w: '90%', h: 1.2,
      fontSize: 42, bold: true, color: WHITE, align: 'center', fontFace: 'Calibri',
    });
    s1.addText(clientName, {
      x: 0.5, y: 2.5, w: '90%', h: 0.7,
      fontSize: 28, color: 'a8d4f5', align: 'center', fontFace: 'Calibri',
    });
    if (effectiveDate) {
      s1.addText(`Effective: ${effectiveDate}`, {
        x: 0.5, y: 3.7, w: '90%', h: 0.5,
        fontSize: 18, color: 'c0d8f0', align: 'center',
      });
    }
    s1.addText('Confidential — Prepared for Decision Makers', {
      x: 0.5, y: 5.0, w: '90%', h: 0.4,
      fontSize: 13, color: '7fa8cc', align: 'center', italic: true,
    });
    s1.addShape(pptx.ShapeType.rect, { x: 3.5, y: 5.8, w: 3.0, h: 0.04, fill: { color: ACCENT } });
    s1.addText('Prepared by Your Benefits Brokerage', {
      x: 0.5, y: 6.0, w: '90%', h: 0.4,
      fontSize: 12, color: '7fa8cc', align: 'center',
    });

    // ── Slide 2: Case Summary ──
    const s2 = pptx.addSlide();
    addSlideHeader(s2, 'Case Summary', `${clientName} ${effectiveDate ? '— Effective ' + effectiveDate : ''}`);
    const carriers = [...new Set(plans.map(p => p.carrier).filter(Boolean))];
    const freqLabel = { weekly: 'Weekly', biweekly: 'Bi-Weekly', semimonthly: 'Semi-Monthly', monthly: 'Monthly' };
    const summaryRows = [
      ['Client', clientName],
      ['Effective Date', effectiveDate || 'Not specified'],
      ['Plans Analyzed', String(plans.length)],
      ['Carriers', carriers.join(', ') || 'Various'],
      ['Census — EE (Employee Only)', String(census.ee || 0)],
      ['Census — ES (Emp + Spouse)', String(census.es || 0)],
      ['Census — EC (Emp + Child)', String(census.ec || 0)],
      ['Census — EF (Family)', String(census.ef || 0)],
      ['Total Enrolled', String((census.ee || 0) + (census.es || 0) + (census.ec || 0) + (census.ef || 0))],
      ['Payroll Frequency', freqLabel[contributionConfig.payrollFrequency] || 'Bi-Weekly'],
    ];
    // Summary table with alternating rows
    summaryRows.forEach(([label, value], i) => {
      const yPos = 1.25 + i * 0.52;
      const bg = i % 2 === 0 ? LIGHT : WHITE;
      s2.addShape(pptx.ShapeType.rect, { x: 0.5, y: yPos, w: 9.0, h: 0.48, fill: { color: bg }, line: { color: 'dddddd', width: 0.5 }, rectRadius: 0.03 });
      s2.addText(label, { x: 0.65, y: yPos + 0.07, w: 4.2, h: 0.35, fontSize: 12, bold: true, color: TEXT_DARK });
      s2.addText(value, { x: 5.0, y: yPos + 0.07, w: 4.3, h: 0.35, fontSize: 12, color: ACCENT, bold: true });
    });
    addFooter(s2);

    // ── Individual Plan Slides ──
    const rankLabels = useManualSelection
      ? recommendations.map((_, i) => `Selected Plan ${i + 1}`)
      : ['Top Pick', 'Runner-Up', 'Strong Alternative', 'Alternative #4', 'Alternative #5'];
    const rankColors = [GOLD, SILVER, BRONZE, ACCENT, MUTED];
    const tierLabels = { ee: 'EE Only', es: 'EE + Spouse', ec: 'EE + Child(ren)', ef: 'Family' };
    const payPeriodsMap = PAY_PERIODS_PER_MONTH[contributionConfig.payrollFrequency] || PAY_PERIODS_PER_MONTH.biweekly;

    recommendations.forEach((plan, idx) => {
      const s = pptx.addSlide();
      const label = plan.recommendationLabel || rankLabels[idx] || `Recommendation #${idx + 1}`;
      const isLowest = label === 'Lowest Cost Option';
      const badgeColor = isLowest ? GREEN : (rankColors[idx] || ACCENT);

      addSlideHeader(s, label, `${plan.carrier || 'Unknown Carrier'} — ${plan.planName || 'Unknown Plan'}`);

      // Score bar
      const scoreVal = Math.round((plan.totalScore || 0) * 100);
      s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.2, w: 9.0, h: 0.32, fill: { color: 'e0e0e0' }, rectRadius: 0.08 });
      s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.2, w: Math.max(0.3, 9.0 * (plan.totalScore || 0)), h: 0.32, fill: { color: badgeColor }, rectRadius: 0.08 });
      s.addText(`Score: ${scoreVal}/100`, { x: 0.5, y: 1.2, w: 9.0, h: 0.32, fontSize: 11, bold: true, color: WHITE, align: 'center' });

      // ── Left column: Plan Details ──
      const leftX = 0.5;
      const colW = 4.3;
      s.addText('PLAN DETAILS', { x: leftX, y: 1.72, w: colW, h: 0.3, fontSize: 9, bold: true, color: ACCENT, letterSpacing: 1.5 });

      const details = [
        ['Network', plan.networkType || '—'],
        ['Metal Level', plan.metalLevel || '—'],
        ['Deductible (Ind)', fmtInt(plan.deductibleIndividual)],
        ['OOP Max (Ind)', fmtInt(plan.oopMaxIndividual)],
        ['PCP Copay', plan.copayPCP != null ? `$${plan.copayPCP}` : '—'],
        ['Specialist Copay', plan.copaySpecialist != null ? `$${plan.copaySpecialist}` : '—'],
        ['Coinsurance', plan.coinsurance != null ? `${plan.coinsurance}%` : '—'],
      ];
      details.forEach(([lbl, val], i) => {
        const yPos = 2.02 + i * 0.4;
        const bg = i % 2 === 0 ? LIGHT : WHITE;
        s.addShape(pptx.ShapeType.rect, { x: leftX, y: yPos, w: colW, h: 0.38, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
        s.addText(lbl, { x: leftX + 0.1, y: yPos + 0.05, w: 2.0, h: 0.28, fontSize: 10, bold: true, color: TEXT_DARK });
        s.addText(val, { x: leftX + 2.1, y: yPos + 0.05, w: 2.0, h: 0.28, fontSize: 10, color: ACCENT, bold: true, align: 'right' });
      });

      // ── Right column: Composite Monthly Rates ──
      const rightX = 5.2;
      const rightW = 4.3;
      s.addText('COMPOSITE MONTHLY RATES', { x: rightX, y: 1.72, w: rightW, h: 0.3, fontSize: 9, bold: true, color: ACCENT, letterSpacing: 1.5 });

      const rates = [
        ['EE Only', fmtDol(plan.premiumEE)],
        ['EE + Spouse', fmtDol(plan.premiumES)],
        ['EE + Child(ren)', fmtDol(plan.premiumEC)],
        ['Family', fmtDol(plan.premiumEF)],
      ];
      rates.forEach(([lbl, val], i) => {
        const yPos = 2.02 + i * 0.4;
        const bg = i % 2 === 0 ? LIGHT : WHITE;
        s.addShape(pptx.ShapeType.rect, { x: rightX, y: yPos, w: rightW, h: 0.38, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
        s.addText(lbl, { x: rightX + 0.1, y: yPos + 0.05, w: 2.0, h: 0.28, fontSize: 10, bold: true, color: TEXT_DARK });
        s.addText(val, { x: rightX + 2.1, y: yPos + 0.05, w: 2.0, h: 0.28, fontSize: 10, color: ACCENT, bold: true, align: 'right' });
      });

      // ── Per Employee Per Pay Period Section ──
      const shares = calculatePlanCostShares(plan, census, contributionConfig);
      const payFreqLabel = freqLabel[shares.payrollFrequency] || 'Bi-Weekly';
      const ppY = 3.85;

      s.addText(`PER EMPLOYEE PER PAY PERIOD (${payFreqLabel.toUpperCase()})`, { x: 0.5, y: ppY, w: 9.0, h: 0.3, fontSize: 9, bold: true, color: ACCENT, letterSpacing: 1.5 });

      // Header row
      const ppHeaderY = ppY + 0.3;
      s.addShape(pptx.ShapeType.rect, { x: 0.5, y: ppHeaderY, w: 9.0, h: 0.35, fill: { color: PRIMARY } });
      s.addText('Coverage Tier', { x: 0.6, y: ppHeaderY + 0.05, w: 2.5, h: 0.25, fontSize: 9, bold: true, color: WHITE });
      s.addText('Monthly Rate', { x: 3.2, y: ppHeaderY + 0.05, w: 1.8, h: 0.25, fontSize: 9, bold: true, color: WHITE, align: 'center' });
      s.addText('Employer Pays', { x: 5.1, y: ppHeaderY + 0.05, w: 2.0, h: 0.25, fontSize: 9, bold: true, color: WHITE, align: 'center' });
      s.addText('Employee Pays', { x: 7.2, y: ppHeaderY + 0.05, w: 2.0, h: 0.25, fontSize: 9, bold: true, color: WHITE, align: 'center' });

      // Tier rows
      const tierKeys = ['ee', 'es', 'ec', 'ef'];
      tierKeys.forEach((tier, i) => {
        const td = shares.byTier[tier];
        const yPos = ppHeaderY + 0.37 + i * 0.38;
        const bg = i % 2 === 0 ? LIGHT : WHITE;
        s.addShape(pptx.ShapeType.rect, { x: 0.5, y: yPos, w: 9.0, h: 0.36, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
        s.addText(tierLabels[tier], { x: 0.6, y: yPos + 0.05, w: 2.5, h: 0.26, fontSize: 10, bold: true, color: TEXT_DARK });
        s.addText(fmtDol(td.premiumPerMemberMonthly), { x: 3.2, y: yPos + 0.05, w: 1.8, h: 0.26, fontSize: 10, color: TEXT_DARK, align: 'center' });
        s.addText(fmtDol(td.employerPerMemberMonthly / payPeriodsMap), { x: 5.1, y: yPos + 0.05, w: 2.0, h: 0.26, fontSize: 10, bold: true, color: PRIMARY, align: 'center' });
        s.addText(fmtDol(td.employeePerMemberMonthly / payPeriodsMap), { x: 7.2, y: yPos + 0.05, w: 2.0, h: 0.26, fontSize: 10, bold: true, color: ORANGE, align: 'center' });
      });

      // ── Aggregate totals bar ──
      const aggY = ppHeaderY + 0.37 + 4 * 0.38 + 0.15;
      s.addShape(pptx.ShapeType.rect, { x: 0.5, y: aggY, w: 4.3, h: 0.55, fill: { color: 'e8f4fd' }, line: { color: ACCENT, width: 1 }, rectRadius: 0.05 });
      s.addText('Employer Aggregate / Pay Period', { x: 0.6, y: aggY + 0.03, w: 4.0, h: 0.2, fontSize: 8, bold: true, color: MUTED });
      s.addText(fmtDol(shares.employerPerPayTotal), { x: 0.6, y: aggY + 0.22, w: 4.0, h: 0.28, fontSize: 14, bold: true, color: PRIMARY });

      s.addShape(pptx.ShapeType.rect, { x: 5.2, y: aggY, w: 4.3, h: 0.55, fill: { color: 'fef5e7' }, line: { color: ORANGE, width: 1 }, rectRadius: 0.05 });
      s.addText('Employee Aggregate / Pay Period', { x: 5.3, y: aggY + 0.03, w: 4.0, h: 0.2, fontSize: 8, bold: true, color: MUTED });
      s.addText(fmtDol(shares.employeePerPayTotal), { x: 5.3, y: aggY + 0.22, w: 4.0, h: 0.28, fontSize: 14, bold: true, color: ORANGE });

      // Why recommended
      if (plan.whyRecommended) {
        const whyY = aggY + 0.7;
        s.addShape(pptx.ShapeType.rect, { x: 0.5, y: whyY, w: 9.0, h: 0.55, fill: { color: 'e8f4fd' }, line: { color: ACCENT, width: 0.5 }, rectRadius: 0.05 });
        s.addText(plan.whyRecommended, { x: 0.7, y: whyY + 0.05, w: 8.6, h: 0.45, fontSize: 10, italic: true, color: PRIMARY });
      }

      addFooter(s);
    });

    // ── Comparison Slide: Side-by-Side ──
    const s6 = pptx.addSlide();
    const compPlans = recommendations;
    const numPlans = compPlans.length;
    const compLabels = compPlans.map((p, i) => p.recommendationLabel || rankLabels[i] || `#${i + 1}`);
    addSlideHeader(s6, 'Plan Comparison', `Side-by-side view of ${numPlans} ${useManualSelection ? 'selected' : 'recommended'} plans`);

    const rowLabels = [
      'Carrier', 'Network', 'Metal Level', 'Deductible (Ind)', 'OOP Max (Ind)',
      'PCP Copay', 'Specialist Copay', 'Coinsurance',
      'EE Premium/mo', 'EE+Spouse/mo', 'EE+Child/mo', 'Family/mo',
    ];
    const getVal = (plan, label) => {
      switch (label) {
        case 'Carrier': return plan.carrier || '—';
        case 'Network': return plan.networkType || '—';
        case 'Metal Level': return plan.metalLevel || '—';
        case 'Deductible (Ind)': return fmtInt(plan.deductibleIndividual);
        case 'OOP Max (Ind)': return fmtInt(plan.oopMaxIndividual);
        case 'PCP Copay': return plan.copayPCP != null ? `$${plan.copayPCP}` : '—';
        case 'Specialist Copay': return plan.copaySpecialist != null ? `$${plan.copaySpecialist}` : '—';
        case 'Coinsurance': return plan.coinsurance != null ? `${plan.coinsurance}%` : '—';
        case 'EE Premium/mo': return fmtDol(plan.premiumEE);
        case 'EE+Spouse/mo': return fmtDol(plan.premiumES);
        case 'EE+Child/mo': return fmtDol(plan.premiumEC);
        case 'Family/mo': return fmtDol(plan.premiumEF);
        default: return '—';
      }
    };

    // Dynamic column widths based on number of plans (total slide width ~10")
    const startX = 0.3;
    const totalTableW = 9.4;
    const labelColW = numPlans <= 3 ? 2.2 : (numPlans <= 4 ? 1.9 : 1.7);
    const colW = (totalTableW - labelColW) / numPlans;
    const compFontSize = numPlans <= 3 ? 9 : (numPlans <= 4 ? 8 : 7);
    const compNameMaxChars = numPlans <= 3 ? 22 : (numPlans <= 4 ? 18 : 15);

    // Header row
    s6.addShape(pptx.ShapeType.rect, { x: startX, y: 1.18, w: labelColW, h: 0.45, fill: { color: PRIMARY } });
    s6.addText('Benefit', { x: startX, y: 1.21, w: labelColW, h: 0.38, fontSize: compFontSize + 1, bold: true, color: WHITE, align: 'center' });
    compPlans.forEach((plan, ci) => {
      const cx = startX + labelColW + ci * colW;
      const badgeColor = (plan.recommendationLabel === 'Lowest Cost Option') ? GREEN : (rankColors[ci] || ACCENT);
      s6.addShape(pptx.ShapeType.rect, { x: cx, y: 1.18, w: colW, h: 0.45, fill: { color: badgeColor } });
      s6.addText(`${compLabels[ci]}`, { x: cx, y: 1.18, w: colW, h: 0.22, fontSize: compFontSize - 1, bold: true, color: WHITE, align: 'center' });
      s6.addText(`${(plan.planName || 'Plan').substring(0, compNameMaxChars)}`, { x: cx, y: 1.38, w: colW, h: 0.22, fontSize: compFontSize, color: WHITE, align: 'center' });
    });

    rowLabels.forEach((label, ri) => {
      const yPos = 1.66 + ri * 0.43;
      const bg = ri % 2 === 0 ? LIGHT : WHITE;
      const isPremiumRow = ri >= 8;
      s6.addShape(pptx.ShapeType.rect, { x: startX, y: yPos, w: labelColW, h: 0.41, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
      s6.addText(label, { x: startX + 0.08, y: yPos + 0.07, w: labelColW - 0.1, h: 0.28, fontSize: compFontSize, bold: true, color: TEXT_DARK });
      compPlans.forEach((plan, ci) => {
        const cx = startX + labelColW + ci * colW;
        s6.addShape(pptx.ShapeType.rect, { x: cx, y: yPos, w: colW, h: 0.41, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
        s6.addText(getVal(plan, label), { x: cx + 0.05, y: yPos + 0.07, w: colW - 0.1, h: 0.28, fontSize: compFontSize, color: isPremiumRow ? ACCENT : TEXT_DARK, bold: isPremiumRow, align: 'center' });
      });
    });
    addFooter(s6);

    // ── Per-Employee Cost Breakdown Slides ──
    // Split into pages of 3 plans each to keep them readable
    const payFreqLbl = freqLabel[contributionConfig.payrollFrequency] || 'Bi-Weekly';
    const contribDescParts = COVERAGE_TIERS.map(t => {
      const rule = contributionConfig.tiers[t] || { type: 'percent', value: 0 };
      const lbl = { ee: 'EE', es: 'ES', ec: 'EC', ef: 'EF' }[t];
      return rule.type === 'dollar' ? `${lbl}: $${rule.value}` : `${lbl}: ${rule.value}%`;
    });
    const contribSummaryText = `Employer Contribution: ${contribDescParts.join('  |  ')}  •  Payroll: ${payFreqLbl}`;

    const plansPerCostPage = 3;
    for (let pageStart = 0; pageStart < compPlans.length; pageStart += plansPerCostPage) {
      const pagePlans = compPlans.slice(pageStart, pageStart + plansPerCostPage);
      const pageNum = Math.floor(pageStart / plansPerCostPage) + 1;
      const totalPages = Math.ceil(compPlans.length / plansPerCostPage);
      const pageLabel = totalPages > 1 ? ` (${pageNum}/${totalPages})` : '';
      const s7 = pptx.addSlide();
      addSlideHeader(s7, `Per-Employee Cost Breakdown${pageLabel}`, `What each employee pays per ${payFreqLbl.toLowerCase()} pay period`);
      s7.addText(contribSummaryText, { x: 0.5, y: 1.15, w: 9.0, h: 0.3, fontSize: 9, color: MUTED, italic: true });

    const ppColW = 2.1;
    const ppLabelW = 2.8;
    const ppStartX = 0.5;
    const ppTiers = ['ee', 'es', 'ec', 'ef'];

    // One table per plan on this page
    pagePlans.forEach((plan, pi) => {
      const blockY = 1.6 + pi * 2.0;
      const shares = calculatePlanCostShares(plan, census, contributionConfig);
      const planLabel = plan.recommendationLabel || rankLabels[pageStart + pi] || `Plan ${pageStart + pi + 1}`;
      const badgeColor = (plan.recommendationLabel === 'Lowest Cost Option') ? GREEN : (rankColors[pageStart + pi] || ACCENT);

      // Plan label bar
      s7.addShape(pptx.ShapeType.rect, { x: ppStartX, y: blockY, w: 9.0, h: 0.35, fill: { color: badgeColor } });
      s7.addText(`${planLabel}: ${(plan.planName || 'Plan').substring(0, 40)}`, { x: ppStartX + 0.1, y: blockY + 0.04, w: 8.8, h: 0.28, fontSize: 10, bold: true, color: WHITE });

      // Table header
      const hdrY = blockY + 0.37;
      s7.addShape(pptx.ShapeType.rect, { x: ppStartX, y: hdrY, w: ppLabelW, h: 0.3, fill: { color: PRIMARY } });
      s7.addText('Tier', { x: ppStartX + 0.08, y: hdrY + 0.04, w: ppLabelW - 0.1, h: 0.22, fontSize: 9, bold: true, color: WHITE });
      ['Monthly Rate', 'Employer / Pay', 'Employee / Pay'].forEach((hdr, hi) => {
        const hx = ppStartX + ppLabelW + hi * ppColW;
        s7.addShape(pptx.ShapeType.rect, { x: hx, y: hdrY, w: ppColW, h: 0.3, fill: { color: PRIMARY } });
        s7.addText(hdr, { x: hx, y: hdrY + 0.04, w: ppColW, h: 0.22, fontSize: 9, bold: true, color: WHITE, align: 'center' });
      });

      // Data rows
      ppTiers.forEach((tier, ti) => {
        const td = shares.byTier[tier];
        const rY = hdrY + 0.32 + ti * 0.3;
        const bg = ti % 2 === 0 ? LIGHT : WHITE;
        s7.addShape(pptx.ShapeType.rect, { x: ppStartX, y: rY, w: ppLabelW, h: 0.28, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
        s7.addText(tierLabels[tier], { x: ppStartX + 0.08, y: rY + 0.03, w: ppLabelW - 0.1, h: 0.22, fontSize: 9, bold: true, color: TEXT_DARK });
        const vals = [
          fmtDol(td.premiumPerMemberMonthly),
          fmtDol(td.employerPerMemberMonthly / payPeriodsMap),
          fmtDol(td.employeePerMemberMonthly / payPeriodsMap),
        ];
        vals.forEach((v, vi) => {
          const vx = ppStartX + ppLabelW + vi * ppColW;
          s7.addShape(pptx.ShapeType.rect, { x: vx, y: rY, w: ppColW, h: 0.28, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
          s7.addText(v, { x: vx, y: rY + 0.03, w: ppColW, h: 0.22, fontSize: 9, bold: true, color: vi === 1 ? PRIMARY : vi === 2 ? ORANGE : TEXT_DARK, align: 'center' });
        });
      });
    });
    addFooter(s7);
    } // end cost breakdown page loop

    // ── Aggregate Monthly Costs Slide ──
    const s8 = pptx.addSlide();
    addSlideHeader(s8, 'Aggregate Monthly Costs', `Total employer and employee costs across all enrolled members`);

    // Dynamic column widths for aggregate slide
    const aggTotalW = 9.4;
    const aggLabelColW = numPlans <= 3 ? 2.2 : (numPlans <= 4 ? 1.8 : 1.5);
    const aggEnrolledW = numPlans <= 3 ? 1.0 : 0.8;
    const premColW = (aggTotalW - aggLabelColW - aggEnrolledW) / numPlans;
    const premStartX = 0.3;
    const aggFontSize = numPlans <= 3 ? 9 : (numPlans <= 4 ? 8 : 7);
    const premHeaders = ['Tier', 'Enrolled', ...compPlans.map((p, i) => (p.recommendationLabel || rankLabels[i] || `Plan ${i+1}`).substring(0, numPlans <= 3 ? 16 : 12))];

    // Header — positioned with dynamic widths
    const aggHeaderWidths = [aggLabelColW, aggEnrolledW, ...compPlans.map(() => premColW)];
    let aggHdrX = premStartX;
    premHeaders.forEach((hdr, hi) => {
      const w = aggHeaderWidths[hi];
      s8.addShape(pptx.ShapeType.rect, { x: aggHdrX, y: 1.2, w, h: 0.45, fill: { color: PRIMARY } });
      s8.addText(hdr, { x: aggHdrX + 0.05, y: 1.24, w: w - 0.1, h: 0.36, fontSize: aggFontSize, bold: true, color: WHITE, align: hi === 0 ? 'left' : 'center' });
      aggHdrX += w;
    });

    const tierCensus = { ee: census.ee || 0, es: census.es || 0, ec: census.ec || 0, ef: census.ef || 0 };
    const premierTiers = [
      { key: 'ee', label: 'EE Only', field: 'premiumEE' },
      { key: 'es', label: 'EE + Spouse', field: 'premiumES' },
      { key: 'ec', label: 'EE + Child(ren)', field: 'premiumEC' },
      { key: 'ef', label: 'Family', field: 'premiumEF' },
    ];

    premierTiers.forEach(({ key, label, field }, ri) => {
      const yPos = 1.7 + ri * 0.42;
      const bg = ri % 2 === 0 ? LIGHT : WHITE;
      // Tier label
      s8.addShape(pptx.ShapeType.rect, { x: premStartX, y: yPos, w: aggLabelColW, h: 0.4, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
      s8.addText(label, { x: premStartX + 0.08, y: yPos + 0.07, w: aggLabelColW - 0.1, h: 0.26, fontSize: aggFontSize + 1, bold: true, color: TEXT_DARK });
      // Enrolled
      const enrolledX = premStartX + aggLabelColW;
      s8.addShape(pptx.ShapeType.rect, { x: enrolledX, y: yPos, w: aggEnrolledW, h: 0.4, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
      s8.addText(String(tierCensus[key]), { x: enrolledX, y: yPos + 0.07, w: aggEnrolledW, h: 0.26, fontSize: aggFontSize + 1, color: TEXT_DARK, align: 'center' });
      // Plan premiums
      compPlans.forEach((plan, ci) => {
        const cx = premStartX + aggLabelColW + aggEnrolledW + ci * premColW;
        s8.addShape(pptx.ShapeType.rect, { x: cx, y: yPos, w: premColW, h: 0.4, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
        s8.addText(fmtDol(plan[field]), { x: cx, y: yPos + 0.07, w: premColW, h: 0.26, fontSize: aggFontSize, bold: true, color: ACCENT, align: 'center' });
      });
    });

    // Totals section
    const totY = 1.7 + 4 * 0.42 + 0.15;
    const totLabelW = aggLabelColW + aggEnrolledW;
    const totalRowLabels = ['Est. Monthly Total', `Employer / ${payFreqLbl}`, `Employee / ${payFreqLbl}`];
    const totalRowColors = [TEXT_DARK, PRIMARY, ORANGE];
    totalRowLabels.forEach((lbl, ri) => {
      const yPos = totY + ri * 0.42;
      s8.addShape(pptx.ShapeType.rect, { x: premStartX, y: yPos, w: totLabelW, h: 0.4, fill: { color: ri === 0 ? PRIMARY : 'e8f4fd' }, line: { color: 'e8e8e8', width: 0.3 } });
      s8.addText(lbl, { x: premStartX + 0.08, y: yPos + 0.07, w: totLabelW - 0.1, h: 0.26, fontSize: aggFontSize + 1, bold: true, color: ri === 0 ? WHITE : totalRowColors[ri] });
      compPlans.forEach((plan, ci) => {
        const cx = premStartX + totLabelW + ci * premColW;
        let val;
        if (ri === 0) val = fmtDol(plan.monthlyTotalCost);
        else if (ri === 1) val = fmtDol(plan.employerPerPayCost);
        else val = fmtDol(plan.employeePerPayCost);
        const bg = ri === 0 ? PRIMARY : 'e8f4fd';
        s8.addShape(pptx.ShapeType.rect, { x: cx, y: yPos, w: premColW, h: 0.4, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.3 } });
        s8.addText(val, { x: cx, y: yPos + 0.07, w: premColW, h: 0.26, fontSize: aggFontSize, bold: true, color: ri === 0 ? WHITE : totalRowColors[ri], align: 'center' });
      });
    });
    addFooter(s8);

    // ── Slide 9: Appendix — All Plans ──
    const s9 = pptx.addSlide();
    addSlideHeader(s9, 'Appendix', 'All plans analyzed, ranked by overall score');
    const allPlansList = (recData.allPlans || plans).slice(0, 13);

    // Mini table header
    const appColWidths = [0.4, 3.0, 1.5, 1.0, 1.2, 1.2, 1.6];
    const appHeaders = ['#', 'Plan Name', 'Network', 'Metal', 'Deductible', 'OOP Max', 'EE Premium'];
    const appStartY = 1.2;
    let appX = 0.3;
    appHeaders.forEach((hdr, hi) => {
      s9.addShape(pptx.ShapeType.rect, { x: appX, y: appStartY, w: appColWidths[hi], h: 0.35, fill: { color: PRIMARY } });
      s9.addText(hdr, { x: appX + 0.04, y: appStartY + 0.05, w: appColWidths[hi] - 0.08, h: 0.25, fontSize: 8, bold: true, color: WHITE, align: hi === 0 ? 'center' : 'left' });
      appX += appColWidths[hi];
    });

    allPlansList.forEach((plan, i) => {
      const yPos = appStartY + 0.37 + i * 0.38;
      if (yPos > 7.0) return;
      const bg = i % 2 === 0 ? LIGHT : WHITE;
      const vals = [
        String(i + 1),
        `${plan.carrier || ''} — ${(plan.planName || '').substring(0, 28)}`,
        plan.networkType || '—',
        plan.metalLevel || '—',
        fmtInt(plan.deductibleIndividual),
        fmtInt(plan.oopMaxIndividual),
        fmtDol(plan.premiumEE),
      ];
      let vx = 0.3;
      vals.forEach((val, vi) => {
        s9.addShape(pptx.ShapeType.rect, { x: vx, y: yPos, w: appColWidths[vi], h: 0.36, fill: { color: bg }, line: { color: 'e8e8e8', width: 0.2 } });
        s9.addText(val, { x: vx + 0.04, y: yPos + 0.05, w: appColWidths[vi] - 0.08, h: 0.26, fontSize: 8, color: TEXT_DARK, align: vi === 0 ? 'center' : 'left' });
        vx += appColWidths[vi];
      });
    });
    addFooter(s9);

    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    const safeName = (clientName || 'Client').replace(/[^a-z0-9]/gi, '_');
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-Disposition': `attachment; filename="BenefitsAnalysis_${safeName}.pptx"`,
      'Content-Length': buffer.length,
    });
    res.send(buffer);
  } catch (err) {
    console.error('PPTX export error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ── Export XLSX ───────────────────────────────────────────────────────────────
app.post('/export/xlsx', async (req, res) => {
  try {
    const { caseId, clientName = 'Client', effectiveDate = '', contribution } = req.body;
    if (!caseId) return res.status(400).json({ error: 'caseId is required' });
    const caseData = caseStore.get(caseId);
    if (!caseData) return res.status(404).json({ error: 'Case not found' });

    const plans = caseData.plans || [];
    const recData = caseData.recommendations || {};
    const contributionConfig = normalizeContributionConfig(contribution || recData.contribution || caseData.contribution || defaultContributionConfig());
    caseData.contribution = contributionConfig;
    const recommendations = (recData.recommendations || plans.slice(0, 3).map((p, i) => ({ rank: i + 1, ...p }))).map(plan => {
      if (typeof plan.employerPerPayCost === 'number' && typeof plan.employeePerPayCost === 'number') return plan;
      const shares = calculatePlanCostShares(plan, caseData.census || {}, contributionConfig);
      return {
        ...plan,
        employerMonthlyCost: Math.round(shares.employerMonthlyTotal * 100) / 100,
        employeeMonthlyCost: Math.round(shares.employeeMonthlyTotal * 100) / 100,
        employerPerPayCost: Math.round(shares.employerPerPayTotal * 100) / 100,
        employeePerPayCost: Math.round(shares.employeePerPayTotal * 100) / 100,
        payrollFrequency: shares.payrollFrequency,
        contributionBreakdown: shares.byTier,
      };
    });
    const census = caseData.census || {};

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Quote Analyzer';
    workbook.lastModifiedBy = 'Quote Analyzer';
    workbook.created = new Date();

    const PRIMARY_HEX = '1e3a5f';
    const ACCENT_HEX = '2e86de';
    const ALT_ROW = 'f4f6f9';

    const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + PRIMARY_HEX } };
    const altFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + ALT_ROW } };
    const whiteFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFdddddd' } },
      bottom: { style: 'thin', color: { argb: 'FFdddddd' } },
      left: { style: 'thin', color: { argb: 'FFdddddd' } },
      right: { style: 'thin', color: { argb: 'FFdddddd' } },
    };

    // ── Sheet 1: Data ──
    const dataSheet = workbook.addWorksheet('Data');
    const dataColumns = [
      { header: 'ID', key: 'id', width: 36 },
      { header: 'Carrier', key: 'carrier', width: 20 },
      { header: 'Plan Name', key: 'planName', width: 28 },
      { header: 'Plan Code', key: 'planCode', width: 15 },
      { header: 'Network Type', key: 'networkType', width: 14 },
      { header: 'Metal Level', key: 'metalLevel', width: 14 },
      { header: 'Deductible (Ind)', key: 'deductibleIndividual', width: 18 },
      { header: 'Deductible (Fam)', key: 'deductibleFamily', width: 18 },
      { header: 'OOP Max (Ind)', key: 'oopMaxIndividual', width: 16 },
      { header: 'OOP Max (Fam)', key: 'oopMaxFamily', width: 16 },
      { header: 'Coinsurance', key: 'coinsurance', width: 14 },
      { header: 'Copay PCP', key: 'copayPCP', width: 12 },
      { header: 'Copay Specialist', key: 'copaySpecialist', width: 18 },
      { header: 'Copay Urgent Care', key: 'copayUrgentCare', width: 18 },
      { header: 'Copay ER', key: 'copayER', width: 12 },
      { header: 'Rx Deductible', key: 'rxDeductible', width: 15 },
      { header: 'Rx Tier 1', key: 'rxTier1', width: 12 },
      { header: 'Rx Tier 2', key: 'rxTier2', width: 12 },
      { header: 'Rx Tier 3', key: 'rxTier3', width: 12 },
      { header: 'Premium EE', key: 'premiumEE', width: 14 },
      { header: 'Premium ES', key: 'premiumES', width: 14 },
      { header: 'Premium EC', key: 'premiumEC', width: 14 },
      { header: 'Premium EF', key: 'premiumEF', width: 14 },
      { header: 'Effective Date', key: 'effectiveDate', width: 16 },
      { header: 'Rating Area', key: 'ratingArea', width: 14 },
      { header: 'Confidence', key: 'extractionConfidence', width: 12 },
      { header: 'Source File', key: 'sourceFile', width: 24 },
    ];
    dataSheet.columns = dataColumns;

    // Header row styling
    const dataHeaderRow = dataSheet.getRow(1);
    dataHeaderRow.eachCell(cell => {
      cell.font = headerFont;
      cell.fill = headerFill;
      cell.border = thinBorder;
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
    dataHeaderRow.height = 28;

    plans.forEach((plan, i) => {
      const row = dataSheet.addRow({
        id: plan.id,
        carrier: plan.carrier,
        planName: plan.planName,
        planCode: plan.planCode,
        networkType: plan.networkType,
        metalLevel: plan.metalLevel,
        deductibleIndividual: plan.deductibleIndividual,
        deductibleFamily: plan.deductibleFamily,
        oopMaxIndividual: plan.oopMaxIndividual,
        oopMaxFamily: plan.oopMaxFamily,
        coinsurance: plan.coinsurance,
        copayPCP: plan.copayPCP,
        copaySpecialist: plan.copaySpecialist,
        copayUrgentCare: plan.copayUrgentCare,
        copayER: plan.copayER,
        rxDeductible: plan.rxDeductible,
        rxTier1: plan.rxTier1,
        rxTier2: plan.rxTier2,
        rxTier3: plan.rxTier3,
        premiumEE: plan.premiumEE,
        premiumES: plan.premiumES,
        premiumEC: plan.premiumEC,
        premiumEF: plan.premiumEF,
        effectiveDate: plan.effectiveDate,
        ratingArea: plan.ratingArea,
        extractionConfidence: plan.extractionConfidence != null ? Math.round(plan.extractionConfidence * 100) + '%' : null,
        sourceFile: plan.sourceFile,
      });
      row.eachCell(cell => {
        cell.fill = i % 2 === 0 ? whiteFill : altFill;
        cell.border = thinBorder;
        cell.alignment = { vertical: 'middle' };
      });
      row.height = 20;
    });

    dataSheet.autoFilter = { from: 'A1', to: dataColumns[dataColumns.length - 1].letter + '1' };

    // ── Sheet 2: Summary ──
    const summSheet = workbook.addWorksheet('Summary');

    // Title block
    summSheet.mergeCells('A1:G1');
    const titleCell = summSheet.getCell('A1');
    titleCell.value = `Benefits Analysis Summary — ${clientName}`;
    titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' }, name: 'Calibri' };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + PRIMARY_HEX } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    summSheet.getRow(1).height = 32;

    summSheet.mergeCells('A2:G2');
    const subCell = summSheet.getCell('A2');
    subCell.value = effectiveDate ? `Effective Date: ${effectiveDate}` : 'Draft Analysis';
    subCell.font = { size: 12, color: { argb: 'FF' + ACCENT_HEX }, italic: true };
    subCell.alignment = { horizontal: 'center' };
    summSheet.getRow(2).height = 22;

    // Census block
    summSheet.getRow(4).height = 22;
    ['A4', 'B4', 'C4', 'D4'].forEach((addr, i) => {
      const cell = summSheet.getCell(addr);
      cell.value = ['Census Tier', 'Count', 'Description', ''].filter((_, idx) => idx === i)[0] || '';
    });
    const censusHdrRow = summSheet.getRow(4);
    censusHdrRow.eachCell(cell => {
      cell.font = headerFont;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + ACCENT_HEX } };
      cell.border = thinBorder;
    });

    const censusData = [
      ['EE', census.ee || 0, 'Employee Only'],
      ['ES', census.es || 0, 'Employee + Spouse'],
      ['EC', census.ec || 0, 'Employee + Child(ren)'],
      ['EF', census.ef || 0, 'Family'],
      ['Total', (census.ee || 0) + (census.es || 0) + (census.ec || 0) + (census.ef || 0), ''],
    ];
    censusData.forEach(([tier, count, desc], i) => {
      const rowNum = 5 + i;
      summSheet.getCell(`A${rowNum}`).value = tier;
      summSheet.getCell(`B${rowNum}`).value = count;
      summSheet.getCell(`C${rowNum}`).value = desc;
      const r = summSheet.getRow(rowNum);
      r.eachCell(cell => {
        cell.fill = i % 2 === 0 ? whiteFill : altFill;
        cell.border = thinBorder;
      });
      r.height = 20;
      if (tier === 'Total') {
        r.eachCell(cell => { cell.font = { bold: true }; });
      }
    });

    // Top plans table
    const planHdrRow = 12;
    const planHdrCols = ['Rank', 'Carrier', 'Plan Name', 'Network', 'Metal', 'Score', 'Monthly Total Cost',
      'Employer Monthly', 'Employee Monthly', 'Employer Per Pay', 'Employee Per Pay',
      'Deductible (Ind)', 'OOP Max (Ind)', 'PCP Copay', 'EE Premium', 'EF Premium', 'EE×Count', 'ES×Count', 'EC×Count', 'EF×Count'];
    planHdrCols.forEach((h, i) => {
      const cell = summSheet.getCell(planHdrRow, i + 1);
      cell.value = h;
      cell.font = headerFont;
      cell.fill = headerFill;
      cell.border = thinBorder;
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
    summSheet.getRow(planHdrRow).height = 28;

    recommendations.forEach((plan, i) => {
      const rowNum = planHdrRow + 1 + i;
      const monthly = plan.monthlyTotalCost || 0;
      const annual = monthly * 12;
      const rowData = [
        plan.rank,
        plan.carrier,
        plan.planName,
        plan.networkType,
        plan.metalLevel,
        plan.totalScore != null ? Math.round(plan.totalScore * 100) + '/100' : '—',
        monthly > 0 ? `$${monthly.toFixed(2)}` : '—',
        plan.employerMonthlyCost != null ? `$${Number(plan.employerMonthlyCost).toFixed(2)}` : '—',
        plan.employeeMonthlyCost != null ? `$${Number(plan.employeeMonthlyCost).toFixed(2)}` : '—',
        plan.employerPerPayCost != null ? `$${Number(plan.employerPerPayCost).toFixed(2)}` : '—',
        plan.employeePerPayCost != null ? `$${Number(plan.employeePerPayCost).toFixed(2)}` : '—',
        plan.deductibleIndividual != null ? `$${plan.deductibleIndividual.toLocaleString()}` : '—',
        plan.oopMaxIndividual != null ? `$${plan.oopMaxIndividual.toLocaleString()}` : '—',
        plan.copayPCP != null ? `$${plan.copayPCP}` : '—',
        plan.premiumEE != null ? `$${plan.premiumEE.toFixed(2)}` : '—',
        plan.premiumEF != null ? `$${plan.premiumEF.toFixed(2)}` : '—',
        plan.premiumEE != null ? `$${(plan.premiumEE * (census.ee || 0)).toFixed(2)}` : '—',
        plan.premiumES != null ? `$${(plan.premiumES * (census.es || 0)).toFixed(2)}` : '—',
        plan.premiumEC != null ? `$${(plan.premiumEC * (census.ec || 0)).toFixed(2)}` : '—',
        plan.premiumEF != null ? `$${(plan.premiumEF * (census.ef || 0)).toFixed(2)}` : '—',
      ];
      rowData.forEach((val, ci) => {
        const cell = summSheet.getCell(rowNum, ci + 1);
        cell.value = val != null ? val : '—';
        cell.fill = i % 2 === 0 ? whiteFill : altFill;
        cell.border = thinBorder;
        cell.alignment = { vertical: 'middle' };
      });
      summSheet.getRow(rowNum).height = 22;
    });

    const contributionStart = planHdrRow + 6;
    const contributionHeader = summSheet.getCell(`A${contributionStart}`);
    contributionHeader.value = 'Employer Contribution Setup';
    contributionHeader.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
    contributionHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + ACCENT_HEX } };
    contributionHeader.border = thinBorder;

    const contributionRows = [
      ['Payroll Frequency', contributionConfig.payrollFrequency, ''],
      ['EE Rule', `${contributionConfig.tiers.ee.type} ${contributionConfig.tiers.ee.value}`, ''],
      ['ES Rule', `${contributionConfig.tiers.es.type} ${contributionConfig.tiers.es.value}`, ''],
      ['EC Rule', `${contributionConfig.tiers.ec.type} ${contributionConfig.tiers.ec.value}`, ''],
      ['EF Rule', `${contributionConfig.tiers.ef.type} ${contributionConfig.tiers.ef.value}`, ''],
    ];
    contributionRows.forEach((vals, idx) => {
      const rowNum = contributionStart + 1 + idx;
      vals.forEach((val, ci) => {
        const cell = summSheet.getCell(rowNum, ci + 1);
        cell.value = val;
        cell.fill = idx % 2 === 0 ? whiteFill : altFill;
        cell.border = thinBorder;
      });
    });

    // Column widths for summary
    summSheet.columns = planHdrCols.map((h, i) => ({ width: [8, 20, 28, 12, 10, 10, 18, 16, 16, 16, 16, 16, 16, 12, 12, 12, 14, 14, 14, 14][i] || 14 }));

    const xlsxBuffer = await workbook.xlsx.writeBuffer();
    const safeName = (clientName || 'Client').replace(/[^a-z0-9]/gi, '_');
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="BenefitsAnalysis_${safeName}.xlsx"`,
      'Content-Length': xlsxBuffer.length,
    });
    res.send(Buffer.from(xlsxBuffer));
  } catch (err) {
    console.error('XLSX export error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ── Error handler ─────────────────────────────────────────────────────────────
app.use((err, _req, res, _next) => {
  console.error(err.stack);
  res.status(err.status || 500).json({ error: err.message || 'Internal server error' });
});

// ── Start server ──────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Quote Analyzer API listening on port ${PORT}`);
});

module.exports = app;
