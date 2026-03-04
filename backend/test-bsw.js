#!/usr/bin/env node
'use strict';
/**
 * Test suite for BSW (Baylor Scott & White) PDF parsing improvements.
 *
 * Covers:
 *  1. Original plan-code row format (regression check – must still work)
 *  2. Benefit lines that include family deductible / OOP pairs ($IND / $FAM)
 *  3. Premium lines with only 4 values (no monthly total)
 *  4. Benefit lines with ER and Urgent Care copays embedded before OOP
 *  5. BSW Tabular Comparison Grid (Strategy 10) – plan codes as column headers
 *  6. BSW Tabular Grid with descriptive plan names (BSW Premier Gold 1000 …)
 *
 * Usage: node test-bsw.js  (server must be running on port 3001)
 */

const http = require('http');

const API = 'http://localhost:3001';

function httpPost(path, body) {
  return new Promise((resolve, reject) => {
    const url = new URL(path, API);
    const data = JSON.stringify(body);
    const opts = {
      hostname: url.hostname, port: url.port, path: url.pathname,
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) },
    };
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

function assertNotNull(label, actual) {
  if (actual !== null && actual !== undefined) {
    passed++;
    return true;
  } else {
    failed++;
    console.log(`  ✗ ${label}: expected non-null, got ${JSON.stringify(actual)}`);
    return false;
  }
}

async function testScenario(name, textContent, validate) {
  console.log(`\n${'='.repeat(70)}`);
  console.log(`TEST: ${name}`);
  console.log('='.repeat(70));

  const result = await httpPost('/debug/extract', { text: textContent, sourceFile: 'bsw-test.pdf' });
  const plans = result.plans || [];

  console.log(`Plans found: ${plans.length}`);
  for (const p of plans) {
    console.log(`  - ${p.planName || '(no name)'} [${p.planCode || '-'}] | Network: ${p.networkType || '-'} | Metal: ${p.metalLevel || '-'}`);
    console.log(`    Ded: $${p.deductibleIndividual ?? '-'} / $${p.deductibleFamily ?? '-'} | OOP: $${p.oopMaxIndividual ?? '-'} / $${p.oopMaxFamily ?? '-'}`);
    console.log(`    PCP: $${p.copayPCP ?? '-'} | Spec: $${p.copaySpecialist ?? '-'} | ER: $${p.copayER ?? '-'} | UC: $${p.copayUrgentCare ?? '-'} | Coins: ${p.coinsurance ?? '-'}`);
    console.log(`    Premiums: EE=$${p.premiumEE ?? '-'} ES=$${p.premiumES ?? '-'} EC=$${p.premiumEC ?? '-'} EF=$${p.premiumEF ?? '-'}`);
    console.log(`    Rx: $${p.rxTier1 ?? '-'}/$${p.rxTier2 ?? '-'}/$${p.rxTier3 ?? '-'} | Conf: ${((p.extractionConfidence || 0) * 100).toFixed(0)}%`);
  }

  if (validate) validate(plans);
}

async function main() {
  // ── Test 1: Original BSW plan-code row format (regression) ───────────────
  await testScenario('BSW Original Code-Row Format (regression)', `
Baylor Scott & White Health Plan
Group Medical Quote
Effective Date: 01/01/2026

HMO Gold Plans
GHG26P44$1,000$30 / $6020% copayment after deductible$5,000
$3/$50/$125/$250
$500.00$1,000.00$850.00$1,500.00$3,850.00

HMO Silver Plans
SHG26A52$2,500$45 / $7530% copayment after deductible$7,000
$3/$50/$125/$250
$400.00$800.00$680.00$1,200.00$3,080.00

HMO Bronze Plans
BHG26D60$5,000$60 / $9040% copayment after deductible$8,150
$3/$50/$125/$250
$320.00$640.00$544.00$960.00$2,464.00
`, (plans) => {
    assert('plan count >= 3', plans.length >= 3, true);
    const gold = plans.find(p => /gold/i.test(p.metalLevel || '') || /GHG26P44/.test(p.planCode || ''));
    if (gold) {
      assert('gold deductible', gold.deductibleIndividual, 1000);
      assert('gold OOP', gold.oopMaxIndividual, 5000);
      assert('gold PCP', gold.copayPCP, 30);
      assert('gold premiumEE', gold.premiumEE, 500);
      assert('gold premiumEF', gold.premiumEF, 1500);
      assert('gold rxTier1', gold.rxTier1, 3);
      assert('gold carrier', gold.carrier, 'Baylor Scott & White');
    } else {
      failed++; console.log('  ✗ Gold plan not found');
    }
  });

  // ── Test 2: Benefit lines with family deductible/OOP pairs ───────────────
  await testScenario('BSW Code-Row with IND/FAM Deductible and OOP Pairs', `
Baylor Scott & White Health Plan
Group Quote

GHG26P44$1,000 / $2,000$30 / $6020% copayment after deductible$5,000 / $10,000
$3/$50/$125/$250
$500.00$1,000.00$850.00$1,500.00

SHG26A52$2,500 / $5,000$45 / $7530% copayment after deductible$7,000 / $14,000
$3/$50/$125/$250
$400.00$800.00$680.00$1,200.00

BHG26D60$5,000 / $10,000$60 / $9040% copayment after deductible$8,150 / $16,300
$3/$50/$125/$250
$320.00$640.00$544.00$960.00
`, (plans) => {
    assert('plan count >= 3', plans.length >= 3, true);
    const gold = plans.find(p => /GHG26P44/.test(p.planCode || ''));
    if (gold) {
      assert('gold ind deductible', gold.deductibleIndividual, 1000);
      assert('gold fam deductible', gold.deductibleFamily, 2000);
      assert('gold ind OOP', gold.oopMaxIndividual, 5000);
      assert('gold fam OOP', gold.oopMaxFamily, 10000);
      assert('gold premiumEE', gold.premiumEE, 500);
      assert('gold premiumEF', gold.premiumEF, 1500);
    } else {
      failed++; console.log('  ✗ Gold plan GHG26P44 not found');
    }
    const silver = plans.find(p => /SHG26A52/.test(p.planCode || ''));
    if (silver) {
      assert('silver ind deductible', silver.deductibleIndividual, 2500);
      assert('silver fam deductible', silver.deductibleFamily, 5000);
    } else {
      failed++; console.log('  ✗ Silver plan SHG26A52 not found');
    }
  });

  // ── Test 3: Premium lines with only 4 values (no monthly total) ──────────
  await testScenario('BSW Code-Row with 4-Value Premium Lines (no total)', `
Baylor Scott & White Health Plan
Group Quote – Illustrative Rates

GHG26P44$1,000$30 / $6020% copayment after deductible$5,000
$3/$50/$125/$250
$500.00$1,000.00$850.00$1,500.00

SHG26A52$2,500$45 / $7530% copayment after deductible$7,000
$3/$50/$125/$250
$400.00$800.00$680.00$1,200.00

BHG26D60$5,000$60 / $9040% copayment after deductible$8,150
$3/$50/$125/$250
$320.00$640.00$544.00$960.00
`, (plans) => {
    assert('plan count >= 3', plans.length >= 3, true);
    const gold = plans.find(p => /GHG26P44/.test(p.planCode || ''));
    if (gold) {
      assert('gold premiumEE (4-val line)', gold.premiumEE, 500);
      assert('gold premiumES (4-val line)', gold.premiumES, 1000);
      assert('gold premiumEC (4-val line)', gold.premiumEC, 850);
      assert('gold premiumEF (4-val line)', gold.premiumEF, 1500);
    } else {
      failed++; console.log('  ✗ Gold plan not found in 4-value premium test');
    }
  });

  // ── Test 4: BSW Tabular Comparison Grid (Strategy 10) ────────────────────
  await testScenario('BSW Tabular Comparison Grid (plan codes as column headers)', `
Baylor Scott & White Health Plan
Small Group Quote
Effective Date: 01/01/2026

Plan Code  GHG26P44  SHG26A52  BHG26D60
Network    BSW Premier  BSW Access  BSW Plus
Metal      Gold      Silver    Bronze
Plan Type  HMO       HMO       HMO

Benefits
Individual Deductible  $1,000  $2,500  $5,000
Family Deductible      $2,000  $5,000  $10,000
Individual OOP Max     $5,000  $7,500  $8,150
Family OOP Max         $10,000 $15,000 $16,300
Primary Care Copay     $30     $45     $60
Specialist Copay       $60     $75     $90
Emergency Room         $300    $400    $500
Urgent Care            $75     $100    $150
Coinsurance            20%     30%     40%

Monthly Rates
Employee Only          $500.00 $400.00 $320.00
Employee + Spouse      $1,000.00 $800.00 $640.00
Employee + Child(ren)  $850.00 $680.00 $544.00
Family                 $1,500.00 $1,200.00 $960.00
`, (plans) => {
    assert('plan count >= 3', plans.length >= 3, true);
    const gold = plans.find(p => /GHG26P44/.test(p.planCode || '') || /gold/i.test(p.metalLevel || ''));
    if (gold) {
      assert('tabular gold carrier', gold.carrier, 'Baylor Scott & White');
      assert('tabular gold metal', gold.metalLevel, 'Gold');
      assert('tabular gold network', gold.networkType, 'HMO');
      assert('tabular gold ind deductible', gold.deductibleIndividual, 1000);
      assert('tabular gold fam deductible', gold.deductibleFamily, 2000);
      assert('tabular gold ind OOP', gold.oopMaxIndividual, 5000);
      assert('tabular gold fam OOP', gold.oopMaxFamily, 10000);
      assert('tabular gold PCP', gold.copayPCP, 30);
      assert('tabular gold specialist', gold.copaySpecialist, 60);
      assert('tabular gold ER', gold.copayER, 300);
      assert('tabular gold UC', gold.copayUrgentCare, 75);
      assert('tabular gold premiumEE', gold.premiumEE, 500);
      assert('tabular gold premiumEF', gold.premiumEF, 1500);
    } else {
      failed++; console.log('  ✗ Gold plan not found in tabular grid test');
    }
    const bronze = plans.find(p => /BHG26D60/.test(p.planCode || '') || /bronze/i.test(p.metalLevel || ''));
    if (bronze) {
      assert('tabular bronze deductible', bronze.deductibleIndividual, 5000);
      assert('tabular bronze premiumEE', bronze.premiumEE, 320);
    } else {
      failed++; console.log('  ✗ Bronze plan not found in tabular grid test');
    }
  });

  // ── Test 5: BSW Tabular Grid with descriptive plan names ─────────────────
  await testScenario('BSW Tabular Grid with Descriptive Plan Names', `
Baylor Scott & White Health Plan
Group Proposal

BSW Premier Gold 1000  BSW Access Silver 2500  BSW Plus Bronze 5000

Individual Deductible  $1,000  $2,500  $5,000
Family Deductible      $2,000  $5,000  $10,000
OOP Max Individual     $5,000  $7,500  $8,150
OOP Max Family         $10,000 $15,000 $16,300
Primary Care           $30     $45     $60
Specialist             $60     $75     $90
Emergency Room         $300    $400    $500
Urgent Care            $75     $100    $150

Employee Only          $500.00 $400.00 $320.00
Employee + Spouse      $1,000.00 $800.00 $640.00
Employee + Children    $850.00 $680.00 $544.00
Family                 $1,500.00 $1,200.00 $960.00
`, (plans) => {
    assert('descriptive plan count >= 3', plans.length >= 3, true);
    const gold = plans.find(p => /gold/i.test(p.metalLevel || '') || /gold/i.test(p.planName || ''));
    if (gold) {
      assert('descriptive gold carrier', gold.carrier, 'Baylor Scott & White');
      assert('descriptive gold deductible', gold.deductibleIndividual, 1000);
      assert('descriptive gold OOP', gold.oopMaxIndividual, 5000);
      assert('descriptive gold premiumEE', gold.premiumEE, 500);
    } else {
      failed++; console.log('  ✗ Gold plan not found in descriptive name tabular test');
    }
  });

  // ── Summary ────────────────────────────────────────────────────────────────
  console.log('\n' + '='.repeat(70));
  console.log(`ALL TESTS COMPLETE: ${passed} passed, ${failed} failed`);
  console.log('='.repeat(70));
  process.exit(failed > 0 ? 1 : 0);
}

main().catch(err => { console.error('Fatal:', err.message); process.exit(1); });
