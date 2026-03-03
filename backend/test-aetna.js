/**
 * Test script for Strategy 8: Aetna Medical Cost Grid parser
 * Feeds actual Aetna PDF text into the extractFromAetnaCostGrid function
 * and verifies correct extraction of plan details and premiums.
 */

const { v4: uuidv4 } = require('uuid');

// ── Copy utility functions from server.js ──────────────────────────────────
function parseMoney(str) {
  if (!str) return null;
  const cleaned = String(str).replace(/[$,\s]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

function detectCarrier(text, sourceFile) {
  if (/aetna/i.test(text) || /aetna/i.test(sourceFile || '')) return 'Aetna';
  return null;
}

// ── Copy the Strategy 8 function from server.js ────────────────────────────
function extractFromAetnaCostGrid(text, sourceFile) {
  if (!/Medical\s*Cost\s*Grid|AFA\s+(OAAS|CPOS)/i.test(text)) return [];

  const plans = [];
  const lines = text.split('\n');
  const carrier = detectCarrier(text, sourceFile) || 'Aetna';

  const AETNA_NETWORKS = {
    'OAAS': 'OA',
    'CPOS II': 'CPOS',
    'CPOS': 'CPOS',
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!/^AFA\s+/i.test(line)) continue;

    let planName = line;
    let nextIdx = i + 1;

    if (nextIdx < lines.length) {
      const nextLine = lines[nextIdx].trim();
      if (/^V\d{2}\b/.test(nextLine) && nextLine.length < 15) {
        planName += ' ' + nextLine;
        nextIdx++;
      }
    }

    if (nextIdx >= lines.length) continue;
    const idLine = lines[nextIdx].trim();
    const idMatch = idLine.match(/^ID:\s*(\d+)/);
    if (!idMatch) continue;
    const planCode = idMatch[1];
    nextIdx++;

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

    const effectivePCP = copayPCP === 0 ? null : copayPCP;
    const effectiveSpec = copaySpecialist === 0 ? null : copaySpecialist;

    let rxTier1 = null, rxTier2 = null, rxTier3 = null;
    let rxText = '';
    while (nextIdx < lines.length) {
      const rxLine = lines[nextIdx].trim();
      if (/^(?:OAAS|CPOS\s*II?)\$[\d,]+\.\d{2}$/i.test(rxLine)) break;
      rxText += ' ' + rxLine;
      nextIdx++;
    }

    const rxMatch = rxText.match(/(\d+)\/(\d+)\/(\d+)\/(\d+)/);
    if (rxMatch) {
      rxTier1 = parseInt(rxMatch[1], 10);
      rxTier2 = parseInt(rxMatch[2], 10);
      rxTier3 = parseInt(rxMatch[3], 10);
    }

    if (nextIdx >= lines.length) continue;
    const netPremLine = lines[nextIdx].trim();
    const netPremMatch = netPremLine.match(/^(OAAS|CPOS\s*II?)\$([\d,]+\.\d{2})$/i);
    if (!netPremMatch) continue;
    nextIdx++;

    const networkRaw = netPremMatch[1].trim().toUpperCase();
    const premiumEE = parseMoney(netPremMatch[2]);

    let networkType = null;
    if (/HSA/i.test(planName)) networkType = 'HSA';
    else if (/CPOS/i.test(networkRaw)) networkType = 'PPO';
    else if (/OAAS/i.test(networkRaw)) networkType = 'HMO';

    function readPremium() {
      if (nextIdx >= lines.length) return null;
      let l = lines[nextIdx].trim();
      if (/^\(\d+\)$/.test(l)) { nextIdx++; l = nextIdx < lines.length ? lines[nextIdx].trim() : ''; }
      const pm = l.match(/^\$([\d,]+\.\d{2})$/);
      if (pm) { nextIdx++; return parseMoney(pm[1]); }
      return null;
    }

    if (nextIdx < lines.length && /^\(\d+\)$/.test(lines[nextIdx].trim())) nextIdx++;

    const premiumES = readPremium();
    const premiumEC = readPremium();
    const premiumEF = readPremium();

    const fields = [deductibleIndividual, coinsurance, effectivePCP, effectiveSpec,
                    premiumEE, premiumES, premiumEC, premiumEF, rxTier1].filter(v => v != null);
    const confidence = Math.min(1, 0.5 + (fields.length * 0.06));

    plans.push({
      carrier,
      planName: planName.trim(),
      planCode,
      networkType,
      deductibleIndividual,
      coinsurance,
      copayPCP: effectivePCP,
      copaySpecialist: effectiveSpec,
      rxTier1,
      rxTier2,
      rxTier3,
      premiumEE,
      premiumES,
      premiumEC,
      premiumEF,
      extractionConfidence: confidence,
    });
  }

  const effDateMatch = text.match(/Eff\s*Date:\s*(\d{2}\/\d{2}\/\d{2,4})/i);
  if (effDateMatch && plans.length > 0) {
    for (const plan of plans) {
      plan.effectiveDate = effDateMatch[1];
    }
  }

  return plans;
}

// ── Test with actual Aetna PDF text ────────────────────────────────────────
const sampleText = `Group Name: AMARILLOBROKERAGE COMPANYQuote ID: 17130578Eff Date: 01/01/26to 01/01/27
NewBusiness AFAMedical Cost Grid- Single Options
Plan Name
Plan ID
Ded/Co-ins,
PCP/SPECRX
NetworkEEEE +SPEE +CHFAMTotal
AFA OAAS 9100 100%Value CY V25
ID: 30021412
$9100,100/0,0/0
0%Med Ded Applies
OAAS$519.93
(3)
$1,307.56
(0)
$1,042.59
(0)
$1,797.05
(2)
$5,153.89
(5)
$1,468.20$3,151.20$486.23$48.26
AFA OAAS 6500 HSA 70%E CY V25
ID: 30021432
$6500,70/30,40/80
3/10/50/100/20%up to
250/40%up to 500
OAAS$552.27
(3)
$1,393.82
(0)
$1,110.72
(0)
$1,916.85
(2)
$5,490.51
(5)
$1,693.09$3,284.88$453.62$58.92
AFA CPOSII 7500 HSA 100/50 E CY
V25
ID: 30021239
$7500,100/0,0/0
0%Med Ded Applies
CPOS II$578.90
(3)
$1,464.87
(0)
$1,166.82
(0)
$2,015.49
(2)
$5,767.68
(5)
$1,771.68$3,496.21$434.63$65.16
AFA OAAS 6000 70%CY V25
ID: 30021396
$6000,70/30,40/80
3/10/50/80/20%up to
250/40%up to 500
OAAS$620.71
(3)
$1,576.36
(0)
$1,254.87
(0)
$2,170.29
(2)
$6,202.71
(5)
$2,447.17$3,335.29$364.34$55.91`;

const plans = extractFromAetnaCostGrid(sampleText, 'Aetna-Quote.pdf');

console.log(`\n═══ RESULTS: ${plans.length} plans extracted ═══\n`);

let passed = 0;
let failed = 0;

function assert(label, actual, expected) {
  if (actual === expected) {
    passed++;
    return true;
  } else {
    failed++;
    console.log(`  ✗ ${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    return false;
  }
}

// Plan 1: AFA OAAS 9100 100%Value CY V25
if (plans.length >= 1) {
  const p = plans[0];
  console.log(`Plan 1: ${p.planName}`);
  assert('carrier', p.carrier, 'Aetna');
  assert('planCode', p.planCode, '30021412');
  assert('networkType', p.networkType, 'HMO');
  assert('deductible', p.deductibleIndividual, 9100);
  assert('coinsurance', p.coinsurance, null);  // 0% member = null
  assert('copayPCP', p.copayPCP, null);  // 0 = null
  assert('copaySpecialist', p.copaySpecialist, null);  // 0 = null
  assert('premiumEE', p.premiumEE, 519.93);
  assert('premiumES', p.premiumES, 1307.56);
  assert('premiumEC', p.premiumEC, 1042.59);
  assert('premiumEF', p.premiumEF, 1797.05);
  assert('effectiveDate', p.effectiveDate, '01/01/26');
  console.log();
}

// Plan 2: AFA OAAS 6500 HSA 70%E CY V25
if (plans.length >= 2) {
  const p = plans[1];
  console.log(`Plan 2: ${p.planName}`);
  assert('carrier', p.carrier, 'Aetna');
  assert('planCode', p.planCode, '30021432');
  assert('networkType', p.networkType, 'HSA');
  assert('deductible', p.deductibleIndividual, 6500);
  assert('coinsurance', p.coinsurance, '30%');
  assert('copayPCP', p.copayPCP, 40);
  assert('copaySpecialist', p.copaySpecialist, 80);
  assert('rxTier1', p.rxTier1, 3);
  assert('rxTier2', p.rxTier2, 10);
  assert('rxTier3', p.rxTier3, 50);
  assert('premiumEE', p.premiumEE, 552.27);
  assert('premiumES', p.premiumES, 1393.82);
  assert('premiumEC', p.premiumEC, 1110.72);
  assert('premiumEF', p.premiumEF, 1916.85);
  console.log();
}

// Plan 3: AFA CPOSII 7500 HSA 100/50 E CY V25 (multi-line name)
if (plans.length >= 3) {
  const p = plans[2];
  console.log(`Plan 3: ${p.planName}`);
  assert('planName', p.planName, 'AFA CPOSII 7500 HSA 100/50 E CY V25');
  assert('planCode', p.planCode, '30021239');
  assert('networkType', p.networkType, 'HSA');
  assert('deductible', p.deductibleIndividual, 7500);
  assert('coinsurance', p.coinsurance, null);  // 100/0 = 0% member
  assert('premiumEE', p.premiumEE, 578.90);
  assert('premiumES', p.premiumES, 1464.87);
  assert('premiumEC', p.premiumEC, 1166.82);
  assert('premiumEF', p.premiumEF, 2015.49);
  console.log();
}

// Plan 4: AFA OAAS 6000 70%CY V25
if (plans.length >= 4) {
  const p = plans[3];
  console.log(`Plan 4: ${p.planName}`);
  assert('carrier', p.carrier, 'Aetna');
  assert('planCode', p.planCode, '30021396');
  assert('networkType', p.networkType, 'HMO');
  assert('deductible', p.deductibleIndividual, 6000);
  assert('coinsurance', p.coinsurance, '30%');
  assert('copayPCP', p.copayPCP, 40);
  assert('copaySpecialist', p.copaySpecialist, 80);
  assert('premiumEE', p.premiumEE, 620.71);
  assert('premiumES', p.premiumES, 1576.36);
  assert('premiumEC', p.premiumEC, 1254.87);
  assert('premiumEF', p.premiumEF, 2170.29);
  console.log();
}

console.log(`═══ ${passed} passed, ${failed} failed ═══`);
process.exit(failed > 0 ? 1 : 0);
