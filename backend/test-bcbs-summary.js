/**
 * Test script for Strategy 13: BCBS Benefit Summary / Comparison Format parser
 * Sends BCBS-format text to /debug/extract and verifies multi-plan extraction.
 * Usage: node test-bcbs-summary.js (with server running on port 3001)
 */

const http = require('http');
const API = 'http://localhost:3001';

function httpPost(path, body) {
  return new Promise((resolve, reject) => {
    const url = new URL(path, API);
    const data = JSON.stringify(body);
    const opts = { hostname: url.hostname, port: url.port, path: url.pathname, method: 'POST', headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) } };
    const req = http.request(opts, (res) => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString();
        try { resolve(JSON.parse(raw)); } catch { resolve(raw); }
      });
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

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

async function main() {
  console.log('═══ Strategy 13: BCBS Benefit Summary Tests ═══\n');

  // ── Test 1: BCBS TX Benefit Summary with Blue Choice PPO and Blue Advantage HMO ──
  console.log('Test 1: BCBS TX Benefit Summary (PPO + HMO plans)');
  const test1 = await httpPost('/debug/extract', {
    text: `Blue Cross Blue Shield of Texas
Effective Date: 01/01/2026

Blue Choice PPO Gold 500  Blue Choice PPO Silver 1500  Blue Advantage HMO Bronze 3000

Individual Deductible     $500      $1,500     $3,000
Family Deductible         $1,000    $3,000     $6,000
Individual OOP Maximum    $6,000    $7,350     $6,550
Family OOP Maximum        $12,000   $14,700    $13,100
Coinsurance               20%       30%        40%
PCP Office Visit          $25       $40        $60
Specialist                $50       $75        $90
Emergency Room            $250      $350       $500
Urgent Care               $75       $100       $150
Generic Rx                $10       $15        $20
Preferred Brand Rx        $40       $55        $70
Non-Preferred Rx          $80       $100       $120

Monthly Premiums:
Employee Only             $500.00   $450.00    $380.00
Employee + Spouse         $1,000.00 $900.00    $760.00
Employee + Child          $850.00   $765.00    $646.00
Family                    $1,500.00 $1,350.00  $1,140.00`,
    sourceFile: 'BCBS-TX-Quote-2026.pdf'
  });

  const plans1 = test1.plans || [];
  assert('plan count', plans1.length, 3);

  if (plans1.length >= 3) {
    // Plan 1: Blue Choice PPO Gold 500
    const p1 = plans1[0];
    console.log(`  Plan 1: ${p1.planName}`);
    assert('carrier', p1.carrier, 'BCBS');
    assert('network', p1.networkType, 'PPO');
    assert('metal', p1.metalLevel, 'Gold');
    assert('deductibleIndividual', p1.deductibleIndividual, 500);
    assert('deductibleFamily', p1.deductibleFamily, 1000);
    assert('oopMaxIndividual', p1.oopMaxIndividual, 6000);
    assert('oopMaxFamily', p1.oopMaxFamily, 12000);
    assert('coinsurance', p1.coinsurance, '20%');
    assert('copayPCP', p1.copayPCP, 25);
    assert('copaySpecialist', p1.copaySpecialist, 50);
    assert('copayER', p1.copayER, 250);
    assert('copayUrgentCare', p1.copayUrgentCare, 75);
    assert('rxTier1', p1.rxTier1, 10);
    assert('rxTier2', p1.rxTier2, 40);
    assert('rxTier3', p1.rxTier3, 80);
    assert('premiumEE', p1.premiumEE, 500);
    assert('premiumES', p1.premiumES, 1000);
    assert('premiumEC', p1.premiumEC, 850);
    assert('premiumEF', p1.premiumEF, 1500);
    assert('effectiveDate', p1.effectiveDate, '01/01/2026');

    // Plan 2: Blue Choice PPO Silver 1500
    const p2 = plans1[1];
    console.log(`  Plan 2: ${p2.planName}`);
    assert('carrier', p2.carrier, 'BCBS');
    assert('network', p2.networkType, 'PPO');
    assert('metal', p2.metalLevel, 'Silver');
    assert('deductibleIndividual', p2.deductibleIndividual, 1500);
    assert('premiumEE', p2.premiumEE, 450);

    // Plan 3: Blue Advantage HMO Bronze 3000
    const p3 = plans1[2];
    console.log(`  Plan 3: ${p3.planName}`);
    assert('carrier', p3.carrier, 'BCBS');
    assert('network', p3.networkType, 'HMO');
    assert('metal', p3.metalLevel, 'Bronze');
    assert('deductibleIndividual', p3.deductibleIndividual, 3000);
    assert('premiumEE', p3.premiumEE, 380);
  }

  // ── Test 2: BCBS with Individual/Family pair format ──
  console.log('\nTest 2: BCBS with Deductible/OOP pair format');
  const test2 = await httpPost('/debug/extract', {
    text: `Blue Cross Blue Shield
Effective Date: 03/01/2026

Blue Choice PPO Gold  Blue Advantage HMO Silver

Deductible           $500 / $1,000     $2,000 / $4,000
Out-of-Pocket Max    $5,000 / $10,000  $6,350 / $12,700
PCP Copay            $25               $40
Specialist Copay     $50               $75
Emergency Room       $300              $400
Urgent Care          $75               $100

Employee Only        $600.00           $500.00
Emp + Spouse         $1,200.00         $1,000.00
Emp + Child          $960.00           $800.00
Family               $1,800.00         $1,500.00`,
    sourceFile: 'BCBS-Pair-Quote.pdf'
  });

  const plans2 = test2.plans || [];
  assert('plan count', plans2.length, 2);

  if (plans2.length >= 2) {
    const p1 = plans2[0];
    console.log(`  Plan 1: ${p1.planName}`);
    assert('carrier', p1.carrier, 'BCBS');
    assert('deductibleIndividual', p1.deductibleIndividual, 500);
    assert('deductibleFamily', p1.deductibleFamily, 1000);
    assert('oopMaxIndividual', p1.oopMaxIndividual, 5000);
    assert('oopMaxFamily', p1.oopMaxFamily, 10000);
    assert('copayPCP', p1.copayPCP, 25);
    assert('premiumEE', p1.premiumEE, 600);
    assert('premiumEF', p1.premiumEF, 1800);

    const p2 = plans2[1];
    console.log(`  Plan 2: ${p2.planName}`);
    assert('deductibleIndividual', p2.deductibleIndividual, 2000);
    assert('deductibleFamily', p2.deductibleFamily, 4000);
    assert('premiumEE', p2.premiumEE, 500);
  }

  console.log(`\n═══ ${passed} passed, ${failed} failed ═══`);
  process.exit(failed > 0 ? 1 : 0);
}

main().catch(console.error);
