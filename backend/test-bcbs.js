/**
 * Test script for Strategy 9: BCBS Proposal Grid parser
 * Uses actual raw text from the uploaded BCBS PDF.
 */

const { v4: uuidv4 } = require('uuid');
const fs = require('fs');

// ── Utility functions from server.js ──
function parseMoney(str) {
  if (!str) return null;
  const cleaned = String(str).replace(/[$,\s]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

// ── Copy the BCBS parser from server.js ──
function extractFromBCBSGrid(text, sourceFile) {
  if (!/Illustrative\s*Composite|Blue\s*Choice\s*PPO|Blue\s*Advantage\s*HMO/i.test(text)) return [];

  const plans = [];
  const lines = text.split('\n');
  const carrier = 'BCBS';

  let currentNetwork = null;
  let currentMetal = null;
  let isHSA = false;

  const effMatch = text.match(/Effective\s*Date:\s*(\d{2}\/\d{2}\/\d{4})/i);
  const effectiveDate = effMatch ? effMatch[1] : null;

  const PLAN_ID_RE = /^([A-Z]\w{2,8}(?:CHC|ADT|ADV|HMO|PPO))$/;
  const FOOTNOTE_RE = /^(\*\d+)+$/;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (/Blue\s*Choice\s*PPO/i.test(line)) { currentNetwork = 'PPO'; continue; }
    if (/Blue\s*Advantage\s*HMO/i.test(line)) { currentNetwork = 'HMO'; continue; }

    if (/^Platinum$/i.test(line)) { currentMetal = 'Platinum'; continue; }
    if (/^Gold$/i.test(line)) { currentMetal = 'Gold'; continue; }
    if (/^Silver$/i.test(line)) { currentMetal = 'Silver'; continue; }
    if (/^Bronze$/i.test(line)) { currentMetal = 'Bronze'; continue; }
    if (/^Expanded\s*Bronze$/i.test(line)) { currentMetal = 'Bronze'; continue; }
    if (/^HSA\s*Plans$/i.test(line)) { isHSA = true; continue; }
    if (/^(?:PPO|HMO)\s*Plans$/i.test(line)) { isHSA = false; continue; }

    if (/^Plan\s*ID$/i.test(line)) continue;
    if (/^(?:Individual|Coinsurance|Primary|ER|Urgent|In-Patient|Out-Patient|Non-|EOESECEF|Total|Monthly|Medical|Cost|Ded|Network|Preferred|Pharmacy)/i.test(line)) continue;
    if (/^Go\s*to\s*Proposal/i.test(line)) continue;
    if (/^Blue\s*Cross\s*and/i.test(line)) continue;
    if (/^Quote\s*ID:/i.test(line)) continue;
    if (/^No\.\s*of\s*Employees/i.test(line)) continue;
    if (/^Printed:/i.test(line)) continue;
    if (/^\*\d/.test(line)) continue;

    const planIdMatch = line.match(PLAN_ID_RE);
    if (!planIdMatch) continue;

    const planCode = planIdMatch[1];
    let nextIdx = i + 1;

    if (nextIdx < lines.length && FOOTNOTE_RE.test(lines[nextIdx].trim())) nextIdx++;

    function nextLine() {
      while (nextIdx < lines.length) {
        const l = lines[nextIdx].trim();
        nextIdx++;
        if (l.length > 0) return l;
      }
      return '';
    }

    let l = nextLine();
    let dedInMatch = l.match(/^\$(\d[\d,]*)\s*\/\//);
    if (!dedInMatch) continue;
    const deductibleIndividual = parseMoney(dedInMatch[1]);

    l = nextLine();

    l = nextLine();
    let oopInMatch = l.match(/^\$(\d[\d,]*)\s*\/\//);
    const oopMaxIndividual = oopInMatch ? parseMoney(oopInMatch[1]) : null;

    l = nextLine();

    l = nextLine();
    let coinsInMatch = l.match(/^(\d+)%\s*\/\//);
    const coinsurance = coinsInMatch ? `${coinsInMatch[1]}%` : null;

    l = nextLine();

    l = nextLine();
    let copayPCP = null;
    let copaySpecialist = null;
    const pcpMatch = l.match(/^\$(\d+)\/\$?(\d+)\$?(\d+)/);
    if (pcpMatch) {
      copayPCP = parseInt(pcpMatch[1], 10);
      copaySpecialist = parseInt(pcpMatch[3], 10);
    } else {
      const dcMatch = l.match(/^DC\/DC\$?(\d+)/);
      if (dcMatch) {
        copaySpecialist = parseInt(dcMatch[1], 10);
      }
    }

    l = nextLine();
    let copayER = null;
    const erMatch = l.match(/^\$(\d+)\s*\/\//);
    if (erMatch) copayER = parseInt(erMatch[1], 10);

    l = nextLine();

    l = nextLine();
    let copayUrgentCare = null;
    const ucMatch = l.match(/^\$(\d+)$/);
    if (ucMatch) copayUrgentCare = parseInt(ucMatch[1], 10);

    l = nextLine();
    l = nextLine();

    l = nextLine();
    l = nextLine();

    l = nextLine();
    let rxTier1 = null, rxTier2 = null, rxTier3 = null;
    let premiumLineOverride = null;
    let rxMatch = l.match(/^\$(\d+)\/\$(\d+)\/\$(\d+)/);
    if (rxMatch) {
      rxTier1 = parseInt(rxMatch[1], 10);
      rxTier2 = parseInt(rxMatch[2], 10);
      rxTier3 = parseInt(rxMatch[3], 10);
      l = nextLine();
    } else if (/^\d+%.*\$[\d,]+\.\d{2}/.test(l)) {
      premiumLineOverride = l;
    } else if (/^\d+%/.test(l)) {
      if (l.endsWith('/')) l = nextLine();
    } else if (/^100%$/.test(l)) {
      // no-op
    }

    l = premiumLineOverride || nextLine();
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

    let networkType = isHSA ? 'HSA' : currentNetwork;
    const metalStr = currentMetal || '';
    const networkStr = currentNetwork === 'HMO' ? 'Blue Advantage HMO' : 'Blue Choice PPO';
    const planName = `${networkStr} ${metalStr} ${planCode}`.replace(/\s+/g, ' ').trim();

    const fields = [deductibleIndividual, oopMaxIndividual, coinsurance,
                    copayPCP, copaySpecialist, copayER, copayUrgentCare,
                    premiumEE, premiumES, premiumEC, premiumEF,
                    rxTier1].filter(v => v != null);
    const confidence = Math.min(1, 0.5 + (fields.length * 0.05));

    plans.push({
      id: 'test-' + plans.length,
      carrier,
      planName,
      planCode,
      networkType,
      metalLevel: currentMetal,
      deductibleIndividual,
      oopMaxIndividual,
      coinsurance,
      copayPCP: copayPCP === 0 ? null : copayPCP,
      copaySpecialist: copaySpecialist === 0 ? null : copaySpecialist,
      copayUrgentCare,
      copayER,
      rxTier1,
      rxTier2,
      rxTier3,
      premiumEE,
      premiumES,
      premiumEC,
      premiumEF,
      effectiveDate,
      extractionConfidence: confidence,
    });
  }

  return plans;
}

// ── Load actual BCBS raw text ──
const rawText = fs.readFileSync('/tmp/bcbs-raw.txt', 'utf8');
const plans = extractFromBCBSGrid(rawText, '2026 QCENT BC Quote.pdf');

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

// Verify first PPO Platinum plan: P9M1CHC
const p1 = plans.find(p => p.planCode === 'P9M1CHC');
if (p1) {
  console.log(`Plan: ${p1.planName} (${p1.planCode})`);
  assert('carrier', p1.carrier, 'BCBS');
  assert('networkType', p1.networkType, 'PPO');
  assert('metalLevel', p1.metalLevel, 'Platinum');
  assert('deductible', p1.deductibleIndividual, 0);
  assert('oopMax', p1.oopMaxIndividual, 6300);
  assert('coinsurance', p1.coinsurance, '90%');
  assert('copayPCP', p1.copayPCP, 20);
  assert('copaySpecialist', p1.copaySpecialist, 40);
  assert('copayER', p1.copayER, 500);
  assert('copayUrgentCare', p1.copayUrgentCare, 75);
  assert('rxTier1', p1.rxTier1, 10);
  assert('rxTier2', p1.rxTier2, 20);
  assert('rxTier3', p1.rxTier3, 70);
  assert('premiumEE', p1.premiumEE, 1638.23);
  assert('premiumES', p1.premiumES, 3276.46);
  assert('premiumEC', p1.premiumEC, 3276.46);
  assert('premiumEF', p1.premiumEF, 4914.69);
  assert('effectiveDate', p1.effectiveDate, '01/01/2026');
  console.log();
} else {
  console.log('✗ Plan P9M1CHC not found!');
  failed++;
}

// Verify a Gold plan: G654CHC
const p2 = plans.find(p => p.planCode === 'G654CHC');
if (p2) {
  console.log(`Plan: ${p2.planName} (${p2.planCode})`);
  assert('metalLevel', p2.metalLevel, 'Gold');
  assert('deductible', p2.deductibleIndividual, 1300);
  assert('premiumEE', p2.premiumEE, 1401.39);
  console.log();
} else {
  console.log('✗ Plan G654CHC not found!');
  failed++;
}

// Verify an HMO plan: P9M1ADT (first HMO Platinum)
const p3 = plans.find(p => p.planCode === 'P9M1ADT');
if (p3) {
  console.log(`Plan: ${p3.planName} (${p3.planCode})`);
  assert('networkType', p3.networkType, 'HMO');
  assert('metalLevel', p3.metalLevel, 'Platinum');
  assert('deductible', p3.deductibleIndividual, 0);
  assert('premiumEE', p3.premiumEE, 1025.28);
  assert('premiumEF', p3.premiumEF, 3075.84);
  console.log();
} else {
  console.log('✗ Plan P9M1ADT not found!');
  failed++;
}

// Print summary of all plans
console.log('── All plans extracted ──');
for (const p of plans) {
  const fields = [p.deductibleIndividual, p.oopMaxIndividual, p.coinsurance,
                  p.copayPCP, p.premiumEE].filter(v => v != null).length;
  console.log(`  ${p.planCode.padEnd(10)} ${(p.networkType||'-').padEnd(5)} ${(p.metalLevel||'-').padEnd(10)} Ded=$${p.deductibleIndividual||'-'}  OOP=$${p.oopMaxIndividual||'-'}  EE=$${p.premiumEE||'-'}  Conf=${(p.extractionConfidence*100).toFixed(0)}%`);
}

console.log(`\n═══ ${passed} passed, ${failed} failed ═══`);
process.exit(failed > 0 ? 1 : 0);
