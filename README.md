# Quote-Analyzer
tool to analyze and organize quotes for account managers to use internally

Internal Quote Analyzer Dashboard
Butler Benefits – Internal Tool
Purpose
The Internal Quote Analyzer Dashboard is an internal-only web application designed to ingest carrier quote PDFs and Excel files, extract structured plan data, normalize plan comparisons, score and rank plan value, and generate professional PowerPoint and Excel presentation outputs.
Core Capabilities (MVP)
1. File Upload & Parsing
• Accepts: PDF, XLSX, XLS, CSV
• Multi-file upload per case
• Extracts carrier, plan name/code, network type, deductibles, OOP maximums, coinsurance, copays, Rx tiers, premium rates, effective date, and census counts (if present).
• Provides extraction confidence scoring and editable corrections in UI.
2. Census Input (Manual Override)
If census cannot be reliably extracted, user manually enters:
• Employee Only (EE)
• EE/Spouse (ES)
• EE/Children (EC)
• EE/Family (EF)
The system calculates total monthly premium, annual premium, and tier distribution summaries.
3. Value Scoring Engine
Plans are scored using a 100-point composite model based on:
• Premium Efficiency (30%)
• Risk Protection (25%)
• Cost Sharing Usability (20%)
• Network Configuration (15%)
• Actuarial Strength Estimate (10%)
Top 3 plans are automatically identified while all plans remain visible.
Outputs
1. PowerPoint Export (.pptx)
Includes title slide, case overview, top 3 recommendations, comparison tables, premium charts, network notes, and appendix of all plans.
2. Excel Export (.xlsx)
Includes normalized data table, census tab, summary tab, and premium charts.
Architecture Overview
Frontend: HTML/CSS/JS embedded in WordPress.
Backend API handles file upload, parsing, scoring, and export generation.
MVP storage: in-memory session; future: SQL-based persistence.
Security Notes
• Internal-only tool
• API token validation required
• No persistent PHI storage
• Files processed in memory and discarded after export
Roadmap (Post-MVP)
• Case saving & history
• CRM integration
• Enhanced carrier-specific parsers
• Advanced actuarial modeling
• Contribution modeling scenarios
• AI-generated executive summaries
Philosophy
This tool standardizes underwriting logic, compresses analysis time, elevates advisory consistency, and creates institutional knowledge to scale advisory leverage.

