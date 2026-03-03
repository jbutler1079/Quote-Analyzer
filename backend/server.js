'use strict';

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');
const PptxGenJS = require('pptxgenjs');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// ── In-memory store ──────────────────────────────────────────────────────────
const caseStore = new Map(); // caseId → { files, plans, census, recommendations }

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

// ── Auth middleware ───────────────────────────────────────────────────────────
const API_TOKEN = process.env.API_TOKEN || 'internal-token-2024';

function authMiddleware(req, res, next) {
  const token = req.headers['x-api-token'];
  if (!token || token !== API_TOKEN) {
    return res.status(401).json({ error: 'Unauthorized: invalid or missing X-API-Token' });
  }
  next();
}

// ── Health ────────────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ status: 'ok' }));

// ── Upload ────────────────────────────────────────────────────────────────────
app.post('/upload', authMiddleware, upload.array('files[]', 20), (req, res) => {
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
    caseStore.set(caseId, { files, plans: [], census: {}, recommendations: null });
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

function extractPlanFromText(text, sourceFile) {
  const plans = [];
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const fullText = text;

  // Detect network type occurrences to split into plan blocks if multiple
  const networkPattern = /\b(HMO|PPO|EPO|HDHP|HSA)\b/gi;
  const planNamePattern =
    /(?:plan\s*name\s*[:\-]?\s*)([^\n]{3,60})|([A-Z][A-Za-z0-9 \-\/]*(HMO|PPO|EPO|HDHP|HDHP)[A-Za-z0-9 \-\/]*)/gi;

  // ── Attempt to find multiple plan blocks separated by plan names ──
  const planBlocks = splitIntoPlanBlocks(lines, fullText);

  for (const block of planBlocks) {
    const plan = extractFieldsFromBlock(block, sourceFile);
    if (plan) plans.push(plan);
  }

  // fallback: if nothing parsed, create one plan from full text
  if (plans.length === 0) {
    const plan = extractFieldsFromBlock(lines.join('\n'), sourceFile);
    if (plan) plans.push(plan);
  }

  return plans;
}

function splitIntoPlanBlocks(lines, fullText) {
  // Look for lines that look like plan headers
  const headerRe =
    /^(?:plan\s*(name|type)?\s*[:\-]?\s*\d*\.?\s*)?([A-Z][A-Za-z0-9\s\/\-]*(HMO|PPO|EPO|HDHP|HSA|Platinum|Gold|Silver|Bronze)[A-Za-z0-9\s\/\-]*)\s*$/i;
  const planStartIndices = [];

  lines.forEach((line, i) => {
    if (headerRe.test(line) && line.length < 80) {
      planStartIndices.push(i);
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
  const carrier = get([
    /(?:carrier|insurance\s*company|insurer)\s*[:\-]\s*([A-Za-z][A-Za-z\s&,.]+)/i,
    /^(Anthem|Aetna|Cigna|United\s*Health|UHC|Kaiser|BlueCross|Blue\s*Cross|BCBS|Humana|Molina|Oscar|Centene|Wellmark|Harvard\s*Pilgrim|Tufts|HCSC|Premera|Regence|Providence|HealthNet|Health\s*Net|Coventry|WellCare|Magellan|Ambetter)/im,
  ]);

  // Plan name
  const planName = get([
    /(?:plan\s*name|plan\s*title)\s*[:\-]\s*([^\n]{3,60})/i,
    /([A-Za-z][A-Za-z0-9\s\/\-]*(HMO|PPO|EPO|HDHP)[A-Za-z0-9\s\/\-]{0,30})/i,
  ]);

  // Network type
  const networkRaw = get([/\b(HDHP|EPO|HMO|PPO|HSA)\b/i]);
  const networkType = networkRaw ? networkRaw.toUpperCase() : null;

  // Metal level
  const metalRaw = get([/\b(Platinum|Gold|Silver|Bronze)\b/i]);
  const metalLevel = metalRaw ? metalRaw.charAt(0).toUpperCase() + metalRaw.slice(1).toLowerCase() : null;

  // Deductible Individual
  const dedIndRaw = get([
    /individual\s+deductible\s*[:\-]?\s*\$?([\d,]+)/i,
    /deductible\s*[:\-]?\s*individual\s*[:\-]?\s*\$?([\d,]+)/i,
    /deductible\s*[:\-]?\s*\$?([\d,]+)/i,
  ]);
  const deductibleIndividual = parseMoney(dedIndRaw);

  // Deductible Family
  const dedFamRaw = get([
    /family\s+deductible\s*[:\-]?\s*\$?([\d,]+)/i,
    /deductible\s*[:\-]?\s*family\s*[:\-]?\s*\$?([\d,]+)/i,
  ]);
  const deductibleFamily = parseMoney(dedFamRaw);

  // OOP Max Individual
  const oopIndRaw = get([
    /individual\s+out[\s-]of[\s-]pocket\s*(?:max(?:imum)?)?\s*[:\-]?\s*\$?([\d,]+)/i,
    /out[\s-]of[\s-]pocket\s*(?:max(?:imum)?)?\s*[:\-]?\s*individual\s*[:\-]?\s*\$?([\d,]+)/i,
    /oop\s*(?:max(?:imum)?)?\s*[:\-]?\s*\$?([\d,]+)/i,
    /out[\s-]of[\s-]pocket\s*(?:max(?:imum)?)?\s*[:\-]?\s*\$?([\d,]+)/i,
  ]);
  const oopMaxIndividual = parseMoney(oopIndRaw);

  // OOP Max Family
  const oopFamRaw = get([
    /family\s+out[\s-]of[\s-]pocket\s*(?:max(?:imum)?)?\s*[:\-]?\s*\$?([\d,]+)/i,
    /out[\s-]of[\s-]pocket\s*(?:max(?:imum)?)?\s*family\s*[:\-]?\s*\$?([\d,]+)/i,
  ]);
  const oopMaxFamily = parseMoney(oopFamRaw);

  // Coinsurance
  const coinsurance = get([/coinsurance\s*[:\-]?\s*(\d+%(?:\s*\/\s*\d+%)?)/i]);

  // Copays
  const copayPCPRaw = get([
    /pcp\s*copay?\s*[:\-]?\s*\$?([\d]+)/i,
    /primary\s*care\s*(?:physician|visit)?\s*(?:copay?)?\s*[:\-]?\s*\$?([\d]+)/i,
    /office\s*visit\s*(?:copay?)?\s*[:\-]?\s*\$?([\d]+)/i,
  ]);
  const copayPCP = parseMoney(copayPCPRaw);

  const copaySpecRaw = get([
    /specialist\s*(?:copay?)?\s*[:\-]?\s*\$?([\d]+)/i,
    /specialist\s*visit\s*(?:copay?)?\s*[:\-]?\s*\$?([\d]+)/i,
  ]);
  const copaySpecialist = parseMoney(copaySpecRaw);

  const copayUCRaw = get([
    /urgent\s*care\s*(?:copay?)?\s*[:\-]?\s*\$?([\d]+)/i,
    /urgent\s*(?:care)?\s*[:\-]?\s*\$?([\d]+)/i,
  ]);
  const copayUrgentCare = parseMoney(copayUCRaw);

  const copayERRaw = get([
    /emergency\s*(?:room|dept|department)?\s*(?:copay?)?\s*[:\-]?\s*\$?([\d]+)/i,
    /er\s*(?:copay?)?\s*[:\-]?\s*\$?([\d]+)/i,
  ]);
  const copayER = parseMoney(copayERRaw);

  // Rx
  const rxDedRaw = get([/rx\s*deductible\s*[:\-]?\s*\$?([\d,]+)/i]);
  const rxDeductible = parseMoney(rxDedRaw);

  const rxT1Raw = get([/(?:generic|tier\s*1|tier1)\s*(?:rx|drug|copay?)?\s*[:\-]?\s*\$?([\d]+)/i]);
  const rxTier1 = parseMoney(rxT1Raw);

  const rxT2Raw = get([/(?:preferred\s*brand|tier\s*2|tier2)\s*(?:rx|drug|copay?)?\s*[:\-]?\s*\$?([\d]+)/i]);
  const rxTier2 = parseMoney(rxT2Raw);

  const rxT3Raw = get([/(?:non[\s-]preferred\s*brand|tier\s*3|tier3)\s*(?:rx|drug|copay?)?\s*[:\-]?\s*\$?([\d]+)/i]);
  const rxTier3 = parseMoney(rxT3Raw);

  // Premiums
  const premEERaw = get([
    /(?:ee|employee\s*only|employee)\s*(?:monthly)?\s*(?:premium|rate)?\s*[:\-]?\s*\$?([\d,]+\.?\d*)/i,
    /premium\s*(?:ee|employee\s*only)\s*[:\-]?\s*\$?([\d,]+\.?\d*)/i,
  ]);
  const premiumEE = parseMoney(premEERaw);

  const premESRaw = get([
    /(?:es|emp\+sp|employee\s*\+?\s*spouse|employee\/spouse)\s*(?:monthly)?\s*(?:premium|rate)?\s*[:\-]?\s*\$?([\d,]+\.?\d*)/i,
  ]);
  const premiumES = parseMoney(premESRaw);

  const premECRaw = get([
    /(?:ec|emp\+ch|employee\s*\+?\s*child(?:ren)?|employee\/child(?:ren)?)\s*(?:monthly)?\s*(?:premium|rate)?\s*[:\-]?\s*\$?([\d,]+\.?\d*)/i,
  ]);
  const premiumEC = parseMoney(premECRaw);

  const premEFRaw = get([
    /(?:ef|family|employee\s*\+?\s*family|employee\/family)\s*(?:monthly)?\s*(?:premium|rate)?\s*[:\-]?\s*\$?([\d,]+\.?\d*)/i,
  ]);
  const premiumEF = parseMoney(premEFRaw);

  // Effective date
  const effectiveDate = get([
    /effective\s*(?:date)?\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i,
    /effective\s*[:\-]?\s*(\w+ \d{1,2},?\s*\d{4})/i,
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
    oopMaxIndividual, copayPCP, premiumEE,
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
    deductibleFamily,
    oopMaxIndividual,
    oopMaxFamily,
    coinsurance: coinsurance || null,
    copayPCP,
    copaySpecialist,
    copayUrgentCare,
    copayER,
    rxDeductible,
    rxTier1,
    rxTier2,
    rxTier3,
    premiumEE,
    premiumES,
    premiumEC,
    premiumEF,
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
app.post('/parse', authMiddleware, async (req, res) => {
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
          const plans = extractPlanFromText(data.text, file.originalname);
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

    res.json({ caseId, plans: allPlans, census: caseData.census || {}, warnings });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Recommend endpoint ────────────────────────────────────────────────────────
function scorePlans(plans, census) {
  const { ee = 0, es = 0, ec = 0, ef = 0 } = census;
  const totalHeads = ee + es + ec + ef;

  const scored = plans.map(plan => {
    // ── Premium Efficiency (40%) ──
    let monthlyTotal = 0;
    if (plan.premiumEE) monthlyTotal += plan.premiumEE * ee;
    if (plan.premiumES) monthlyTotal += plan.premiumES * es;
    if (plan.premiumEC) monthlyTotal += plan.premiumEC * ec;
    if (plan.premiumEF) monthlyTotal += plan.premiumEF * ef;

    // ── Risk Protection (30%) ──
    const ded = plan.deductibleIndividual || 0;
    const oop = plan.oopMaxIndividual || 0;
    // $0 deductible+OOP = 100, $15000 = 0
    const riskRaw = Math.max(0, 1 - (ded + oop) / 15000);

    // ── Copay Usability (20%) ──
    const pcp = plan.copayPCP || 50; // default 50 if unknown
    const copayScore = Math.max(0, 1 - pcp / 100);
    // bonus: if copayPCP exists and deductible is 0 or copay-first model
    const copayFirst = plan.copayPCP !== null && (plan.deductibleIndividual === 0 || plan.deductibleIndividual === null);
    const copayUsability = copayFirst ? Math.min(1, copayScore + 0.2) : copayScore;

    // ── Network (10%) ──
    const networkScores = { PPO: 1.0, EPO: 0.8, HMO: 0.7, HDHP: 0.5, HSA: 0.5 };
    const networkScore = networkScores[(plan.networkType || '').toUpperCase()] || 0.65;

    return {
      ...plan,
      _monthlyTotal: monthlyTotal,
      _riskRaw: riskRaw,
      _copayUsability: copayUsability,
      _networkScore: networkScore,
    };
  });

  // Normalize premium efficiency across plans
  const premiums = scored.map(p => p._monthlyTotal).filter(v => v > 0);
  const maxPremium = premiums.length > 0 ? Math.max(...premiums) : 1;
  const minPremium = premiums.length > 0 ? Math.min(...premiums) : 0;
  const premRange = maxPremium - minPremium || 1;

  const result = scored.map(plan => {
    const premEfficiency = plan._monthlyTotal > 0
      ? Math.max(0, 1 - (plan._monthlyTotal - minPremium) / premRange)
      : 0.5; // unknown premium gets middle score

    const totalScore =
      premEfficiency * 0.40 +
      plan._riskRaw * 0.30 +
      plan._copayUsability * 0.20 +
      plan._networkScore * 0.10;

    const reasons = [];
    if (premEfficiency >= 0.7) reasons.push('competitive premium costs');
    if (plan._riskRaw >= 0.7) reasons.push('strong risk protection (low deductible + OOP max)');
    if (plan._copayUsability >= 0.7) reasons.push('excellent copay accessibility');
    if (plan._networkScore >= 0.9) reasons.push('broad PPO network flexibility');
    if (plan.metalLevel === 'Gold' || plan.metalLevel === 'Platinum') reasons.push('rich benefit structure');
    if (reasons.length === 0) reasons.push('balanced overall value');

    const whyRecommended = `This plan scores well due to ${reasons.slice(0, 3).join(', ')}.`;

    return {
      ...plan,
      premiumEfficiencyScore: Math.round(premEfficiency * 100) / 100,
      riskProtectionScore: Math.round(plan._riskRaw * 100) / 100,
      copayUsabilityScore: Math.round(plan._copayUsability * 100) / 100,
      networkScore: Math.round(plan._networkScore * 100) / 100,
      totalScore: Math.round(totalScore * 1000) / 1000,
      monthlyTotalCost: Math.round(plan._monthlyTotal * 100) / 100,
      whyRecommended,
    };
  });

  return result.sort((a, b) => b.totalScore - a.totalScore);
}

app.post('/recommend', authMiddleware, (req, res) => {
  try {
    const { caseId, census } = req.body;
    if (!caseId) return res.status(400).json({ error: 'caseId is required' });
    const caseData = caseStore.get(caseId);
    if (!caseData) return res.status(404).json({ error: 'Case not found' });

    if (census) caseData.census = census;
    const plans = caseData.plans || [];
    if (plans.length === 0) return res.status(400).json({ error: 'No plans to score — run /parse first' });

    const allScored = scorePlans(plans, caseData.census || {});
    const recommendations = allScored.slice(0, 3).map((p, i) => ({ rank: i + 1, ...p }));

    caseData.recommendations = { recommendations, allPlans: allScored };
    caseStore.set(caseId, caseData);

    res.json({ caseId, recommendations, allPlans: allScored });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Export PPTX ───────────────────────────────────────────────────────────────
app.post('/export/pptx', authMiddleware, async (req, res) => {
  try {
    const { caseId, clientName = 'Client', effectiveDate = '' } = req.body;
    if (!caseId) return res.status(400).json({ error: 'caseId is required' });
    const caseData = caseStore.get(caseId);
    if (!caseData) return res.status(404).json({ error: 'Case not found' });

    const plans = caseData.plans || [];
    const recData = caseData.recommendations || {};
    const recommendations = recData.recommendations || plans.slice(0, 3).map((p, i) => ({ rank: i + 1, ...p }));
    const census = caseData.census || {};

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.theme = { headFontFace: 'Calibri', bodyFontFace: 'Calibri' };

    const PRIMARY = '1e3a5f';
    const ACCENT = '2e86de';
    const WHITE = 'FFFFFF';
    const LIGHT = 'f4f6f9';
    const TEXT_DARK = '2c3e50';

    const addSlideHeader = (slide, title) => {
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: '100%', h: 1.0,
        fill: { color: PRIMARY },
      });
      slide.addText(title, {
        x: 0.4, y: 0.1, w: '90%', h: 0.8,
        fontSize: 22, bold: true, color: WHITE, fontFace: 'Calibri',
      });
    };

    // ── Slide 1: Title ──
    const s1 = pptx.addSlide();
    s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: PRIMARY } });
    s1.addText('Benefits Plan Analysis', {
      x: 0.5, y: 1.5, w: '90%', h: 1.2,
      fontSize: 40, bold: true, color: WHITE, align: 'center',
    });
    s1.addText(clientName, {
      x: 0.5, y: 3.0, w: '90%', h: 0.7,
      fontSize: 26, color: 'a8d4f5', align: 'center',
    });
    if (effectiveDate) {
      s1.addText(`Effective: ${effectiveDate}`, {
        x: 0.5, y: 3.8, w: '90%', h: 0.5,
        fontSize: 18, color: 'c0d8f0', align: 'center',
      });
    }
    s1.addText('Prepared by: Your Benefits Brokerage', {
      x: 0.5, y: 5.8, w: '90%', h: 0.4,
      fontSize: 13, color: '7fa8cc', align: 'center', italic: true,
    });

    // ── Slide 2: Case Summary ──
    const s2 = pptx.addSlide();
    addSlideHeader(s2, 'Case Summary');
    const carriers = [...new Set(plans.map(p => p.carrier).filter(Boolean))];
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
    ];
    summaryRows.forEach(([label, value], i) => {
      const yPos = 1.2 + i * 0.55;
      const bg = i % 2 === 0 ? LIGHT : WHITE;
      s2.addShape(pptx.ShapeType.rect, { x: 0.3, y: yPos, w: 9.4, h: 0.5, fill: { color: bg }, line: { color: 'dddddd', width: 0.5 } });
      s2.addText(label, { x: 0.4, y: yPos + 0.08, w: 4.5, h: 0.35, fontSize: 13, bold: true, color: TEXT_DARK });
      s2.addText(value, { x: 5.0, y: yPos + 0.08, w: 4.5, h: 0.35, fontSize: 13, color: TEXT_DARK });
    });

    // ── Slides 3-5: Top Recommendations ──
    recommendations.slice(0, 3).forEach((plan, idx) => {
      const s = pptx.addSlide();
      const rankColors = ['d4af37', '9e9e9e', 'cd7f32'];
      const rankColor = rankColors[idx] || ACCENT;
      addSlideHeader(s, `Recommendation #${idx + 1}`);
      s.addShape(pptx.ShapeType.ellipse, { x: 0.3, y: 1.1, w: 0.8, h: 0.8, fill: { color: rankColor } });
      s.addText(String(idx + 1), { x: 0.3, y: 1.2, w: 0.8, h: 0.6, fontSize: 20, bold: true, color: WHITE, align: 'center' });

      s.addText(plan.planName || 'Unknown Plan', { x: 1.3, y: 1.1, w: 8.0, h: 0.55, fontSize: 22, bold: true, color: PRIMARY });
      s.addText(plan.carrier || '', { x: 1.3, y: 1.65, w: 8.0, h: 0.4, fontSize: 15, color: '555555' });

      const scoreBar = Math.round((plan.totalScore || 0) * 100);
      s.addShape(pptx.ShapeType.rect, { x: 0.3, y: 2.2, w: 9.4, h: 0.35, fill: { color: 'e0e0e0' } });
      s.addShape(pptx.ShapeType.rect, { x: 0.3, y: 2.2, w: 9.4 * (plan.totalScore || 0), h: 0.35, fill: { color: ACCENT } });
      s.addText(`Score: ${scoreBar}/100`, { x: 0.3, y: 2.2, w: 9.4, h: 0.35, fontSize: 12, bold: true, color: WHITE, align: 'center' });

      const detailRows = [
        ['Network Type', plan.networkType || '—'],
        ['Metal Level', plan.metalLevel || '—'],
        ['Deductible (Ind)', plan.deductibleIndividual != null ? `$${plan.deductibleIndividual.toLocaleString()}` : '—'],
        ['OOP Max (Ind)', plan.oopMaxIndividual != null ? `$${plan.oopMaxIndividual.toLocaleString()}` : '—'],
        ['PCP Copay', plan.copayPCP != null ? `$${plan.copayPCP}` : '—'],
        ['Specialist Copay', plan.copaySpecialist != null ? `$${plan.copaySpecialist}` : '—'],
        ['EE Monthly Premium', plan.premiumEE != null ? `$${plan.premiumEE.toFixed(2)}` : '—'],
        ['Family Monthly Premium', plan.premiumEF != null ? `$${plan.premiumEF.toFixed(2)}` : '—'],
      ];
      detailRows.forEach(([label, value], i) => {
        const col = i < 4 ? 0 : 1;
        const row = i % 4;
        const xPos = col === 0 ? 0.3 : 5.1;
        const yPos = 2.8 + row * 0.65;
        s.addShape(pptx.ShapeType.rect, { x: xPos, y: yPos, w: 4.5, h: 0.55, fill: { color: i % 2 === 0 ? LIGHT : WHITE }, line: { color: 'dddddd', width: 0.5 } });
        s.addText(label, { x: xPos + 0.1, y: yPos + 0.08, w: 2.2, h: 0.4, fontSize: 11, bold: true, color: TEXT_DARK });
        s.addText(value, { x: xPos + 2.3, y: yPos + 0.08, w: 2.1, h: 0.4, fontSize: 11, color: ACCENT });
      });

      if (plan.whyRecommended) {
        s.addShape(pptx.ShapeType.rect, { x: 0.3, y: 5.5, w: 9.4, h: 0.9, fill: { color: 'e8f4fd' }, line: { color: ACCENT, width: 1 } });
        s.addText(`💡 ${plan.whyRecommended}`, { x: 0.5, y: 5.55, w: 9.0, h: 0.8, fontSize: 12, italic: true, color: PRIMARY });
      }
    });

    // ── Slide 6: Comparison Table ──
    const s6 = pptx.addSlide();
    addSlideHeader(s6, 'Plan Comparison Table');
    const compPlans = recommendations.slice(0, 3);
    const rowLabels = ['Carrier', 'Network', 'Metal Level', 'Deductible (Ind)', 'OOP Max (Ind)', 'PCP Copay', 'Specialist Copay', 'ER Copay', 'EE Premium', 'Family Premium'];
    const getVal = (plan, label) => {
      const m = v => v != null ? `$${Number(v).toLocaleString()}` : '—';
      switch (label) {
        case 'Carrier': return plan.carrier || '—';
        case 'Network': return plan.networkType || '—';
        case 'Metal Level': return plan.metalLevel || '—';
        case 'Deductible (Ind)': return m(plan.deductibleIndividual);
        case 'OOP Max (Ind)': return m(plan.oopMaxIndividual);
        case 'PCP Copay': return plan.copayPCP != null ? `$${plan.copayPCP}` : '—';
        case 'Specialist Copay': return plan.copaySpecialist != null ? `$${plan.copaySpecialist}` : '—';
        case 'ER Copay': return plan.copayER != null ? `$${plan.copayER}` : '—';
        case 'EE Premium': return plan.premiumEE != null ? `$${plan.premiumEE.toFixed(2)}` : '—';
        case 'Family Premium': return plan.premiumEF != null ? `$${plan.premiumEF.toFixed(2)}` : '—';
        default: return '—';
      }
    };

    // Header row
    const colW = 2.5;
    const startX = 0.3;
    s6.addShape(pptx.ShapeType.rect, { x: startX, y: 1.1, w: 2.2, h: 0.5, fill: { color: PRIMARY } });
    s6.addText('Benefit', { x: startX, y: 1.15, w: 2.2, h: 0.4, fontSize: 12, bold: true, color: WHITE, align: 'center' });
    compPlans.forEach((plan, ci) => {
      const cx = startX + 2.2 + ci * colW;
      s6.addShape(pptx.ShapeType.rect, { x: cx, y: 1.1, w: colW, h: 0.5, fill: { color: ACCENT } });
      s6.addText(`#${ci + 1}: ${(plan.planName || 'Plan').substring(0, 18)}`, { x: cx, y: 1.15, w: colW, h: 0.4, fontSize: 10, bold: true, color: WHITE, align: 'center' });
    });

    rowLabels.forEach((label, ri) => {
      const yPos = 1.65 + ri * 0.52;
      const bg = ri % 2 === 0 ? LIGHT : WHITE;
      s6.addShape(pptx.ShapeType.rect, { x: startX, y: yPos, w: 2.2, h: 0.48, fill: { color: bg }, line: { color: 'dddddd', width: 0.5 } });
      s6.addText(label, { x: startX + 0.05, y: yPos + 0.08, w: 2.1, h: 0.35, fontSize: 10, bold: true, color: TEXT_DARK });
      compPlans.forEach((plan, ci) => {
        const cx = startX + 2.2 + ci * colW;
        s6.addShape(pptx.ShapeType.rect, { x: cx, y: yPos, w: colW, h: 0.48, fill: { color: bg }, line: { color: 'dddddd', width: 0.5 } });
        s6.addText(getVal(plan, label), { x: cx + 0.05, y: yPos + 0.08, w: colW - 0.1, h: 0.35, fontSize: 10, color: TEXT_DARK, align: 'center' });
      });
    });

    // ── Slide 7: Monthly Premium Summary ──
    const s7 = pptx.addSlide();
    addSlideHeader(s7, 'Monthly Premium Summary');
    const premRows = [
      ['Tier', 'Count', ...compPlans.map(p => (p.planName || 'Plan').substring(0, 16))],
      ['EE (Employee Only)', String(census.ee || 0), ...compPlans.map(p => p.premiumEE != null ? `$${p.premiumEE.toFixed(2)}` : '—')],
      ['ES (Emp + Spouse)', String(census.es || 0), ...compPlans.map(p => p.premiumES != null ? `$${p.premiumES.toFixed(2)}` : '—')],
      ['EC (Emp + Child)', String(census.ec || 0), ...compPlans.map(p => p.premiumEC != null ? `$${p.premiumEC.toFixed(2)}` : '—')],
      ['EF (Family)', String(census.ef || 0), ...compPlans.map(p => p.premiumEF != null ? `$${p.premiumEF.toFixed(2)}` : '—')],
      ['Est. Monthly Total', '', ...compPlans.map(p => p.monthlyTotalCost != null ? `$${p.monthlyTotalCost.toFixed(2)}` : '—')],
    ];
    premRows.forEach((row, ri) => {
      const isHeader = ri === 0;
      const yPos = 1.2 + ri * 0.7;
      const bg = isHeader ? PRIMARY : (ri % 2 === 0 ? LIGHT : WHITE);
      const fc = isHeader ? WHITE : TEXT_DARK;
      row.forEach((cell, ci) => {
        const xPos = 0.3 + ci * 2.3;
        s7.addShape(pptx.ShapeType.rect, { x: xPos, y: yPos, w: 2.2, h: 0.6, fill: { color: bg }, line: { color: 'dddddd', width: 0.5 } });
        s7.addText(cell, { x: xPos + 0.05, y: yPos + 0.1, w: 2.1, h: 0.45, fontSize: isHeader ? 11 : 12, bold: isHeader, color: fc, align: ci === 0 ? 'left' : 'center' });
      });
    });

    // ── Slide 8: Appendix ──
    const s8 = pptx.addSlide();
    addSlideHeader(s8, 'Appendix — All Plans Analyzed');
    const allPlansList = (recData.allPlans || plans).slice(0, 15);
    allPlansList.forEach((plan, i) => {
      const yPos = 1.1 + i * 0.45;
      if (yPos > 7.0) return;
      const bg = i % 2 === 0 ? LIGHT : WHITE;
      s8.addShape(pptx.ShapeType.rect, { x: 0.3, y: yPos, w: 9.4, h: 0.42, fill: { color: bg }, line: { color: 'dddddd', width: 0.3 } });
      const label = `${i + 1}. ${plan.carrier || '—'}  |  ${plan.planName || '—'}  |  ${plan.networkType || '—'}  |  Score: ${plan.totalScore != null ? Math.round(plan.totalScore * 100) : '—'}`;
      s8.addText(label, { x: 0.4, y: yPos + 0.06, w: 9.2, h: 0.32, fontSize: 11, color: TEXT_DARK });
    });

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
app.post('/export/xlsx', authMiddleware, async (req, res) => {
  try {
    const { caseId, clientName = 'Client', effectiveDate = '' } = req.body;
    if (!caseId) return res.status(400).json({ error: 'caseId is required' });
    const caseData = caseStore.get(caseId);
    if (!caseData) return res.status(404).json({ error: 'Case not found' });

    const plans = caseData.plans || [];
    const recData = caseData.recommendations || {};
    const recommendations = recData.recommendations || plans.slice(0, 3).map((p, i) => ({ rank: i + 1, ...p }));
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

    // Column widths for summary
    summSheet.columns = planHdrCols.map((h, i) => ({ width: [8, 20, 28, 12, 10, 10, 18, 16, 16, 12, 12, 12, 14, 14, 14, 14][i] || 14 }));

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
