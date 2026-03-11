/**
 * Test script for Strategy 14: UHC / UnitedHealthcare Proposal Format parser
 * Sends UHC-format text to /debug/extract and verifies multi-plan extraction.
 * Usage: node test-uhc.js (with server running on port 3001)
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
  console.log('═══ Strategy 14: UHC Proposal Format Tests ═══\n');

  // ── Test 1: UHC Choice Plus and Navigate plans ──
  console.log('Test 1: UHC Choice Plus PPO + Navigate HMO (3 plans)');
  const test1 = await httpPost('/debug/extract', {
    text: `UnitedHealthcare Group Proposal
Client: ABC Corporation
Effective Date: 07/01/2026

Choice Plus PPO Gold  Choice Plus PPO Silver  Navigate HMO Bronze

Benefit Summary:
Deductible              $750       $1,500      $3,000
OOP Maximum             $6,000     $7,350      $6,550
Coinsurance             20%        30%         40%
PCP Office Visit        $30        $50         $40
Specialist Copay        $60        $80         $70
ER Copay                $300       $400        $350
Urgent Care             $75        $100        $150
Generic Rx (Tier 1)     $10        $15         $20
Preferred Brand (Tier 2) $40       $60         $80
Non-Preferred (Tier 3)  $80        $100        $150

Rate Summary (Monthly):
Employee Only           $545.00    $489.00     $412.00
Employee + Spouse       $1,090.00  $978.00     $824.00
Employee + Child        $926.50    $831.30     $700.40
Family                  $1,635.00  $1,467.00   $1,236.00`,
    sourceFile: 'UHC-Quote-2026.pdf'
  });

  const plans1 = test1.plans || [];
  assert('plan count', plans1.length, 3);

  if (plans1.length >= 3) {
    // Plan 1: Choice Plus PPO Gold
    const p1 = plans1[0];
    console.log(`  Plan 1: ${p1.planName}`);
    assert('carrier', p1.carrier, 'UnitedHealthcare');
    assert('network', p1.networkType, 'PPO');
    assert('metal', p1.metalLevel, 'Gold');
    assert('deductibleIndividual', p1.deductibleIndividual, 750);
    assert('oopMaxIndividual', p1.oopMaxIndividual, 6000);
    assert('coinsurance', p1.coinsurance, '20%');
    assert('copayPCP', p1.copayPCP, 30);
    assert('copaySpecialist', p1.copaySpecialist, 60);
    assert('copayER', p1.copayER, 300);
    assert('copayUrgentCare', p1.copayUrgentCare, 75);
    assert('rxTier1', p1.rxTier1, 10);
    assert('rxTier2', p1.rxTier2, 40);
    assert('rxTier3', p1.rxTier3, 80);
    assert('premiumEE', p1.premiumEE, 545);
    assert('premiumES', p1.premiumES, 1090);
    assert('premiumEC', p1.premiumEC, 926.5);
    assert('premiumEF', p1.premiumEF, 1635);
    assert('effectiveDate', p1.effectiveDate, '07/01/2026');

    // Plan 2: Choice Plus PPO Silver
    const p2 = plans1[1];
    console.log(`  Plan 2: ${p2.planName}`);
    assert('carrier', p2.carrier, 'UnitedHealthcare');
    assert('network', p2.networkType, 'PPO');
    assert('metal', p2.metalLevel, 'Silver');
    assert('deductibleIndividual', p2.deductibleIndividual, 1500);
    assert('oopMaxIndividual', p2.oopMaxIndividual, 7350);
    assert('coinsurance', p2.coinsurance, '30%');
    assert('premiumEE', p2.premiumEE, 489);
    assert('premiumEF', p2.premiumEF, 1467);

    // Plan 3: Navigate HMO Bronze
    const p3 = plans1[2];
    console.log(`  Plan 3: ${p3.planName}`);
    assert('carrier', p3.carrier, 'UnitedHealthcare');
    assert('network', p3.networkType, 'HMO');
    assert('metal', p3.metalLevel, 'Bronze');
    assert('deductibleIndividual', p3.deductibleIndividual, 3000);
    assert('oopMaxIndividual', p3.oopMaxIndividual, 6550);
    assert('coinsurance', p3.coinsurance, '40%');
    assert('premiumEE', p3.premiumEE, 412);
    assert('premiumEF', p3.premiumEF, 1236);
  }

  // ── Test 2: UHC with Individual/Family pair deductibles ──
  console.log('\nTest 2: UHC with Deductible/OOP pairs');
  const test2 = await httpPost('/debug/extract', {
    text: `UnitedHealthcare Small Group Quote
Effective Date: 01/01/2026

Choice Plus PPO Gold  Navigate HMO Silver

Deductible           $500 / $1,000     $2,000 / $4,000
Out-of-Pocket Max    $5,000 / $10,000  $7,000 / $14,000
PCP Copay            $25               $40
Specialist           $50               $75
Emergency Room       $250              $350
Urgent Care          $50               $75

Employee Only        $520.00           $425.00
Emp + Spouse         $1,040.00         $850.00
Emp + Child          $884.00           $722.50
Family               $1,560.00         $1,275.00`,
    sourceFile: 'UHC-Pair-Quote.pdf'
  });

  const plans2 = test2.plans || [];
  assert('plan count', plans2.length, 2);

  if (plans2.length >= 2) {
    const p1 = plans2[0];
    console.log(`  Plan 1: ${p1.planName}`);
    assert('carrier', p1.carrier, 'UnitedHealthcare');
    assert('deductibleIndividual', p1.deductibleIndividual, 500);
    assert('deductibleFamily', p1.deductibleFamily, 1000);
    assert('oopMaxIndividual', p1.oopMaxIndividual, 5000);
    assert('oopMaxFamily', p1.oopMaxFamily, 10000);
    assert('copayPCP', p1.copayPCP, 25);
    assert('copaySpecialist', p1.copaySpecialist, 50);
    assert('copayER', p1.copayER, 250);
    assert('premiumEE', p1.premiumEE, 520);
    assert('premiumEF', p1.premiumEF, 1560);

    const p2 = plans2[1];
    console.log(`  Plan 2: ${p2.planName}`);
    assert('carrier', p2.carrier, 'UnitedHealthcare');
    assert('deductibleIndividual', p2.deductibleIndividual, 2000);
    assert('deductibleFamily', p2.deductibleFamily, 4000);
    assert('oopMaxIndividual', p2.oopMaxIndividual, 7000);
    assert('oopMaxFamily', p2.oopMaxFamily, 14000);
    assert('premiumEE', p2.premiumEE, 425);
    assert('premiumEF', p2.premiumEF, 1275);
  }

  // ── Test 3: UHC with Options PPO and Select Plus ──
  console.log('\nTest 3: UHC Options PPO and Select Plus plans');
  const test3 = await httpPost('/debug/extract', {
    text: `UnitedHealthcare Renewal Proposal
Coverage Period: 04/01/2026

Options PPO Gold 500  Select Plus HMO Silver 1500

Deductible              $500       $1,500
OOP Maximum             $5,000     $7,350
PCP Copay               $25        $40
Specialist              $50        $75
Emergency Room          $250       $350

Employee Only           $485.00    $395.00
Employee + Spouse       $970.00    $790.00
Employee + Child        $824.50    $671.50
Family                  $1,455.00  $1,185.00`,
    sourceFile: 'UHC-Options-Quote.pdf'
  });

  const plans3 = test3.plans || [];
  assert('plan count', plans3.length, 2);

  if (plans3.length >= 2) {
    const p1 = plans3[0];
    console.log(`  Plan 1: ${p1.planName}`);
    assert('carrier', p1.carrier, 'UnitedHealthcare');
    assert('deductibleIndividual', p1.deductibleIndividual, 500);
    assert('premiumEE', p1.premiumEE, 485);

    const p2 = plans3[1];
    console.log(`  Plan 2: ${p2.planName}`);
    assert('carrier', p2.carrier, 'UnitedHealthcare');
    assert('deductibleIndividual', p2.deductibleIndividual, 1500);
    assert('premiumEE', p2.premiumEE, 395);
  }

  console.log(`\n═══ ${passed} passed, ${failed} failed ═══`);
  process.exit(failed > 0 ? 1 : 0);
}

main().catch(console.error);
