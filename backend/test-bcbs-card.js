/**
 * Test script for Strategy 14: BCBS Proposal Card parser
 * Uses actual raw text from the Aegis BCBS PDF.
 */
const fs = require('fs');
const path = require('path');

// We need parseMoney and the new function from server.js
// For quick testing, inline parseMoney and load the raw text
function parseMoney(str) {
  if (!str) return null;
  const cleaned = String(str).replace(/[$,\s]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

// Read the raw text saved from the Aegis PDF
const rawText = fs.readFileSync(path.join(__dirname, 'test-bcbs-aegis-raw.txt'), 'utf8');

console.log('=== Raw text length:', rawText.length);
console.log('=== First 200 chars:', rawText.substring(0, 200));
console.log();

// Test the card pattern matching
const cardPattern = /(?:^|\n)\s*(Platinum|Gold|Silver|Bronze|Expanded\s*Bronze)\s*\n\s*(\d+)\s*\n\s*(BlueCross\s*BlueShield\s*of\s*Texas\s+\S+.*?)(?=\n\s*(?:Platinum|Gold|Silver|Bronze|Expanded\s*Bronze)\s*\n\s*\d+\s*\n\s*BlueCross|Medical\s*Coverage\s*\n|Medical\s*Employer\s*Contribution|\n\s*-\s*\d+\s*Available\s*Plans|$)/gis;

let match;
let planCount = 0;
while ((match = cardPattern.exec(rawText)) !== null) {
  planCount++;
  const metalLevel = match[1].trim();
  const planNum = match[2];
  const cardText = match[3];
  
  console.log(`\n=== Plan ${planCount} ===`);
  console.log(`Metal: ${metalLevel}, Number: ${planNum}`);
  
  // Plan code
  const codeMatch = cardText.match(/BlueCross\s*BlueShield\s*of\s*Texas\s+(\S+)\s+(.*?)(?:\n|$)/i);
  console.log(`Plan code: ${codeMatch ? codeMatch[1] : 'NOT FOUND'}`);
  
  // Network
  const networkMatch = cardText.match(/(PPO|HMO)\s*Network/i);
  console.log(`Network: ${networkMatch ? networkMatch[1] : 'NOT FOUND'}`);
  
  // Deductible
  const dedInMatch = cardText.match(/\(In\)\s*Ind\s*\/\s*Fam\s*\$?([\d,]+(?:\.\d+)?)\s*\/\s*\$?([\d,]+(?:\.\d+)?)/i);
  console.log(`Deductible In: ${dedInMatch ? `$${dedInMatch[1]} / $${dedInMatch[2]}` : 'NOT FOUND'}`);
  
  // OOP
  const oopSection = cardText.match(/Out-of-Pocket\s*Max[\s\S]*?\(In\)\s*Ind\s*\/\s*Fam\s*\$?([\d,]+(?:\.\d+)?)\s*\/\s*\$?([\d,]+(?:\.\d+)?)/i);
  console.log(`OOP In: ${oopSection ? `$${oopSection[1]} / $${oopSection[2]}` : 'NOT FOUND'}`);
  
  // Coinsurance
  const coinsMatch = cardText.match(/In-Network\s*(\d+)%/i);
  console.log(`Coinsurance: ${coinsMatch ? `${coinsMatch[1]}%` : 'NOT FOUND'}`);
  
  // Copays
  const pcpMatch = cardText.match(/Doctor\s*Visit\s*\n?\s*\$(\d+)\s*copay/i);
  const specMatch = cardText.match(/Specialist\s*Visit\s*\n?\s*\$(\d+)\s*copay/i);
  const erMatch = cardText.match(/Emergency\s*Room\s*\n?\s*\$(\d+)\s*copay/i);
  const ucMatch = cardText.match(/Urgent\s*Care\s*\n?\s*\$(\d+)\s*copay/i);
  console.log(`PCP: $${pcpMatch?.[1] || 'N/A'}, Spec: $${specMatch?.[1] || 'N/A'}, ER: $${erMatch?.[1] || 'N/A'}, UC: $${ucMatch?.[1] || 'N/A'}`);
  
  // Rx
  const rxMatch = cardText.match(/Prescription\s*Drugs\s*\n?\s*\$(\d+)\/\$(\d+)\/\$(\d+)/i);
  console.log(`Rx: ${rxMatch ? `$${rxMatch[1]}/$${rxMatch[2]}/$${rxMatch[3]}` : 'NOT FOUND'}`);
  
  // Premiums
  const eeMatch = cardText.match(/Employee\s*Only\s*\(\d+\)\s*\$?([\d,]+\.\d{2})/i);
  const esMatch = cardText.match(/Employee\s*&\s*Spouse\s*\(\d+\)\s*\$?([\d,]+\.\d{2})/i);
  const ecMatch = cardText.match(/Employee\s*&\s*Child(?:ren)?\s*\(\d+\)\s*\$?([\d,]+\.\d{2})/i);
  const efMatch = cardText.match(/Employee\s*&\s*Family\s*\(\d+\)\s*\$?([\d,]+\.\d{2})/i);
  console.log(`Premiums: EE=$${eeMatch?.[1] || 'N/A'}, ES=$${esMatch?.[1] || 'N/A'}, EC=$${ecMatch?.[1] || 'N/A'}, EF=$${efMatch?.[1] || 'N/A'}`);
}

console.log(`\n=== TOTAL PLANS FOUND: ${planCount} ===`);
if (planCount === 3) {
  console.log('✓ SUCCESS: Found expected 3 plans');
} else {
  console.log(`✗ EXPECTED 3 plans, got ${planCount}`);
}
