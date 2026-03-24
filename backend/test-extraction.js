#!/usr/bin/env node
'use strict';
/**
 * Test script: sends sample text to /debug/extract and verifies multi-plan extraction.
 * Usage: node test-extraction.js (with server running on port 3001)
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

async function testScenario(name, textContent, expectedMinPlans) {
  console.log(`\n${'='.repeat(70)}`);
  console.log(`TEST: ${name}`);
  console.log('='.repeat(70));

  const result = await httpPost('/debug/extract', { text: textContent, sourceFile: 'test-quote.pdf' });
  const plans = result.plans || [];

  console.log(`Plans found: ${plans.length}`);
  for (const p of plans) {
    console.log(`  - ${p.planName || '(no name)'} | Network: ${p.networkType || '-'} | Metal: ${p.metalLevel || '-'}`);
    console.log(`    Ded: $${p.deductibleIndividual ?? '-'} / $${p.deductibleFamily ?? '-'} | OOP: $${p.oopMaxIndividual ?? '-'} / $${p.oopMaxFamily ?? '-'}`);
    console.log(`    PCP: $${p.copayPCP ?? '-'} | Specialist: $${p.copaySpecialist ?? '-'} | ER: $${p.copayER ?? '-'} | Urgent: $${p.copayUrgentCare ?? '-'}`);
    console.log(`    Premiums: EE=$${p.premiumEE ?? '-'} ES=$${p.premiumES ?? '-'} EC=$${p.premiumEC ?? '-'} EF=$${p.premiumEF ?? '-'}`);
    console.log(`    Confidence: ${((p.extractionConfidence || 0) * 100).toFixed(0)}%`);
  }
  const pass = plans.length >= expectedMinPlans;
  console.log(`\n  RESULT: ${pass ? 'PASS' : 'FAIL'} (expected >= ${expectedMinPlans} plans, got ${plans.length})`);
  return plans;
}

async function main() {
  // ── Test 1: Comparison grid with multiple spaces between plan columns ──
  const test1 = await testScenario('Comparison Grid (multi-space separated)', `
Anthem Blue Cross Group Medical Proposal
Effective Date: 01/01/2025

Plan Options

Anthem PPO Gold 500  Anthem PPO Silver 1500  Anthem HDHP Bronze 3000

Individual Deductible     $500      $1,500     $3,000
Family Deductible         $1,000    $3,000     $6,000
Individual OOP Maximum    $6,000    $7,350     $6,550
Family OOP Maximum        $12,000   $14,700    $13,100
PCP Office Visit          $25       $40        $60
Specialist                $50       $75        $90
Emergency Room            $250      $350       $500
Urgent Care               $75       $100       $150

Monthly Premiums:
Employee Only             $500.00   $450.00    $380.00
Employee + Spouse         $1,000.00 $900.00    $760.00
Employee + Child          $850.00   $765.00    $646.00
Family                    $1,500.00 $1,350.00  $1,140.00
`, 3);

  // ── Test 2: Comparison grid with "$X / $Y" individual/family pairs ──
  const test2 = await testScenario('Comparison Grid (ind/fam pairs)', `
Aetna Medical Plan Options

Aetna Choice PPO 250   Aetna Choice PPO 500   Aetna Select HMO 1000

Deductible           $250 / $500     $500 / $1,000   $1,000 / $2,000
Out-of-Pocket Max    $5,000 / $10,000  $6,350 / $12,700  $8,150 / $16,300
PCP Copay            $20             $30              $40
Specialist Copay     $40             $60              $80
Emergency Room       $150            $250             $350
Urgent Care          $50             $75              $100

Premium Rates (Monthly):
Employee Only        $475.00         $425.00          $375.00
Emp + Spouse         $950.00         $850.00          $750.00
Emp + Child          $807.50         $722.50          $637.50
Family               $1,425.00       $1,275.00        $1,125.00
`, 3);

  // ── Test 3: Sequential plan layout (column-by-column extraction) ──
  const test3 = await testScenario('Sequential Plan Layout', `
Cigna Medical Benefits

Plan: Cigna PPO Gold
Deductible: $500
Out-of-Pocket Maximum: $6,000
PCP Office Visit: $25
Specialist: $50
Emergency Room: $250
Employee Only: $520.00
Employee + Spouse: $1,040.00
Employee + Child(ren): $884.00
Family: $1,560.00

Plan: Cigna PPO Silver
Deductible: $1,500
Out-of-Pocket Maximum: $7,350
PCP Office Visit: $40
Specialist: $75
Emergency Room: $350
Employee Only: $450.00
Employee + Spouse: $900.00
Employee + Child(ren): $765.00
Family: $1,350.00

Plan: Cigna HDHP Bronze
Deductible: $3,000
Out-of-Pocket Maximum: $6,550
PCP Office Visit: $0 (after deductible)
Specialist: $0 (after deductible)
Emergency Room: $0 (after deductible)
Employee Only: $380.00
Employee + Spouse: $760.00
Employee + Child(ren): $646.00
Family: $1,140.00
`, 3);

  // ── Test 4: Rate table with separate benefit and premium sections ──
  const test4 = await testScenario('Rate Table + Benefits', `
UnitedHealthcare Group Quote
Client: Sample Corp

Choice Plus PPO Gold  Choice Plus PPO Silver  Navigate HMO Bronze

Benefit Summary:
Deductible              $750       $1,500      $2,500
OOP Maximum             $6,000     $7,350      $6,550
Office Visit Copay      $30        $50         $40
Specialist Copay        $60        $80         $70
ER Copay                $300       $400        $350

Rate Summary:
EE                      $545.00    $489.00     $412.00
ES                      $1,090.00  $978.00     $824.00
EC                      $926.50    $831.30     $700.40
EF                      $1,635.00  $1,467.00   $1,236.00
`, 3);
  console.log('\n' + '='.repeat(70));
  console.log('ALL TESTS COMPLETE');
  console.log('='.repeat(70));
}

main().catch(console.error);
