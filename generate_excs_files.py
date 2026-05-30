#!/usr/bin/env python3
"""Generates ExCS IGNITION GTM package files for NAKACHI Consulting."""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import docx.oxml.ns as qn
from docx.oxml import OxmlElement

OUTPUT_DIR = "/home/user/nexus-project1"

# ─────────────────────────────────────────────
# 1.  excs_ignition_control.html  (~65 KB)
# ─────────────────────────────────────────────
CONTROL_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ExCS IGNITION — Master GTM Control Document | NAKACHI Consulting</title>
<style>
  :root {
    --primary: #0a2342;
    --accent:  #e63946;
    --gold:    #f4a261;
    --light:   #f8f9fa;
    --mid:     #dee2e6;
    --text:    #212529;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; color: var(--text);
         background: #fff; line-height: 1.7; font-size: 15px; }
  header { background: var(--primary); color: #fff; padding: 48px 64px; }
  header h1 { font-size: 2.6rem; letter-spacing: 1px; margin-bottom: 8px; }
  header .sub { font-size: 1.1rem; color: var(--gold); text-transform: uppercase;
                letter-spacing: 3px; }
  header .meta { margin-top: 16px; font-size: 0.9rem; opacity: .75; }
  .banner { background: var(--accent); color: #fff; text-align: center;
            padding: 12px; font-size: 0.85rem; letter-spacing: 1px;
            text-transform: uppercase; }
  nav { background: #1b3a5e; padding: 0 64px; display: flex; gap: 0; flex-wrap: wrap; }
  nav a { color: #cdd6e0; text-decoration: none; padding: 14px 20px; font-size: 0.9rem;
          display: block; transition: background .2s; }
  nav a:hover { background: var(--accent); color: #fff; }
  main { max-width: 1100px; margin: 0 auto; padding: 40px 40px 80px; }
  h2.section-title { font-size: 1.7rem; color: var(--primary); border-left: 5px solid var(--accent);
                     padding-left: 16px; margin: 48px 0 20px; }
  h3 { font-size: 1.2rem; color: #1b3a5e; margin: 28px 0 10px; }
  h4 { font-size: 1rem; color: var(--accent); margin: 18px 0 8px; text-transform: uppercase;
       letter-spacing: 1px; }
  p  { margin-bottom: 14px; }
  ul, ol { margin: 0 0 14px 24px; }
  li { margin-bottom: 6px; }
  table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 0.9rem; }
  th { background: var(--primary); color: #fff; padding: 12px 14px; text-align: left; }
  td { padding: 10px 14px; border-bottom: 1px solid var(--mid); }
  tr:nth-child(even) td { background: var(--light); }
  .card-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
               gap: 20px; margin: 24px 0; }
  .card { border: 1px solid var(--mid); border-radius: 8px; padding: 24px;
          background: var(--light); }
  .card .card-num { font-size: 2rem; font-weight: 700; color: var(--accent); }
  .card .card-label { font-size: 0.85rem; color: #6c757d; text-transform: uppercase;
                      letter-spacing: 1px; margin-top: 4px; }
  .card .card-body { margin-top: 12px; font-size: 0.95rem; }
  .highlight-box { background: #fff3cd; border-left: 4px solid var(--gold);
                   padding: 16px 20px; margin: 20px 0; border-radius: 0 6px 6px 0; }
  .risk-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 20px 0; }
  .risk-item { border: 1px solid var(--mid); border-radius: 6px; padding: 16px; }
  .risk-item.high { border-left: 4px solid #dc3545; }
  .risk-item.med  { border-left: 4px solid #fd7e14; }
  .risk-item.low  { border-left: 4px solid #28a745; }
  .timeline { border-left: 3px solid var(--primary); padding-left: 24px; margin: 24px 0; }
  .tl-event { position: relative; margin-bottom: 28px; }
  .tl-event::before { content: ''; position: absolute; left: -31px; top: 6px;
                      width: 14px; height: 14px; border-radius: 50%;
                      background: var(--accent); border: 2px solid var(--primary); }
  .tl-event .tl-date { font-size: 0.8rem; color: #6c757d; text-transform: uppercase;
                        letter-spacing: 1px; }
  .tl-event h4 { margin-top: 4px; }
  .phase-table td:first-child { font-weight: 600; color: var(--primary); }
  .kpi-row { display: flex; gap: 16px; flex-wrap: wrap; margin: 20px 0; }
  .kpi-box { flex: 1; min-width: 160px; background: var(--primary); color: #fff;
             border-radius: 8px; padding: 20px; text-align: center; }
  .kpi-box .kpi-val { font-size: 1.8rem; font-weight: 700; color: var(--gold); }
  .kpi-box .kpi-lbl { font-size: 0.8rem; opacity: .8; margin-top: 4px;
                       text-transform: uppercase; letter-spacing: 1px; }
  .appendix { background: var(--light); border: 1px solid var(--mid);
              border-radius: 8px; padding: 24px; margin: 20px 0; }
  footer { background: var(--primary); color: #adb5bd; text-align: center;
           padding: 28px; font-size: 0.85rem; margin-top: 60px; }
  footer strong { color: var(--gold); }
  .confidential { background: #dc3545; color: #fff; text-align: center;
                  padding: 8px; font-size: 0.8rem; letter-spacing: 2px;
                  text-transform: uppercase; font-weight: 700; }
  @media print {
    nav, .banner { display: none; }
    header { padding: 24px 32px; }
    main { padding: 20px; }
  }
</style>
</head>
<body>

<div class="confidential">CONFIDENTIAL — FOR INTERNAL USE ONLY — NOT FOR DISTRIBUTION</div>

<header>
  <div class="sub">NAKACHI Consulting · ExCS Division</div>
  <h1>ExCS IGNITION</h1>
  <p style="font-size:1.2rem; color:#cdd6e0; margin-top:8px;">
    Master Go-To-Market Control Document
  </p>
  <div class="meta">
    Version 1.0 &nbsp;|&nbsp; Prepared: May 2026 &nbsp;|&nbsp;
    Classification: Confidential &nbsp;|&nbsp; Owner: GTM Strategy Lead
  </div>
</header>

<div class="banner">
  ExCS IGNITION GTM Package — South-South Nigeria Telecom Revenue Initiative
</div>

<nav>
  <a href="#exec">Executive Summary</a>
  <a href="#market">Market Context</a>
  <a href="#framework">ExCS Framework</a>
  <a href="#gtm">GTM Strategy</a>
  <a href="#segments">Target Segments</a>
  <a href="#financial">Financials</a>
  <a href="#timeline">Timeline</a>
  <a href="#kpis">KPIs</a>
  <a href="#risk">Risk</a>
  <a href="#appendix">Appendix</a>
</nav>

<main>

<!-- ── EXECUTIVE SUMMARY ── -->
<h2 class="section-title" id="exec">1. Executive Summary</h2>

<p>
  <strong>ExCS IGNITION</strong> is NAKACHI Consulting's flagship go-to-market initiative targeting
  South-South Nigerian state governments seeking to leverage the NCC (Nigerian Communications Commission)
  Policy Review that now permits State Special Purpose Vehicles (SPVs) to hold equity in
  licensed telecommunications operators.
</p>
<p>
  This document serves as the master control reference for all ExCS IGNITION campaign activities,
  stakeholder engagement protocols, financial modelling assumptions, and programme governance
  structures. It is to be read in conjunction with the Public-Facing Summary (excs_ignition_public.html)
  and the Direct Mail Wave 1 sequence.
</p>

<div class="card-grid">
  <div class="card">
    <div class="card-num">6</div>
    <div class="card-label">Target States</div>
    <div class="card-body">Rivers, Bayelsa, Delta, Akwa Ibom, Cross River, Edo — the full South-South
    geo-political zone of Nigeria.</div>
  </div>
  <div class="card">
    <div class="card-num">NGN 2.4T</div>
    <div class="card-label">Aggregate IGR Potential</div>
    <div class="card-body">Estimated combined annual Internal Generated Revenue uplift across all six
    states at median penetration assumptions.</div>
  </div>
  <div class="card">
    <div class="card-num">40%</div>
    <div class="card-label">Target SPV Equity</div>
    <div class="card-body">Recommended state equity stake in the licensed SPV operator, balancing
    control with private capital attraction.</div>
  </div>
  <div class="card">
    <div class="card-num">18 Mo.</div>
    <div class="card-label">Campaign Horizon</div>
    <div class="card-body">Full GTM wave from initial outreach through signed engagement mandates,
    project mobilisation, and first revenue.</div>
  </div>
</div>

<div class="highlight-box">
  <strong>Strategic Imperative:</strong> The NCC Policy Review creates a narrow first-mover window.
  States that establish their SPVs within 24 months will secure the most attractive spectrum
  allocations and avoid market saturation. NAKACHI's ExCS IGNITION is designed to accelerate
  state readiness and compress the path to revenue.
</div>


<!-- ── MARKET CONTEXT ── -->
<h2 class="section-title" id="market">2. Market Context &amp; Opportunity</h2>

<h3>2.1 The NCC Policy Shift</h3>
<p>
  In Q4 2025, the Nigerian Communications Commission finalised amendments to the
  <em>National Telecommunications Policy Framework</em> that explicitly permit sub-national
  government entities to hold equity positions — up to 49% — in licensed Mobile Network
  Operators (MNOs) and Internet Service Providers (ISPs) operating within their geographic
  jurisdiction. This represents the most significant structural reform in Nigeria's telecoms
  sector since the 2001 liberalisation.
</p>
<p>
  The practical effect is that State Governments, through properly constituted SPVs, can now
  participate directly in the economics of telecoms services delivered to their citizens —
  converting what was previously an invisible revenue stream into a measurable contribution
  to Internal Generated Revenue.
</p>

<h3>2.2 South-South Zone Profile</h3>
<table>
  <thead>
    <tr>
      <th>State</th><th>Population</th><th>Est. Current IGR</th>
      <th>Telecom Penetration (Existing)</th><th>Upside Category</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>Rivers</td><td>7.2 M</td><td>NGN 8.0 B/yr</td><td>62%</td><td>High</td></tr>
    <tr><td>Delta</td><td>5.6 M</td><td>NGN 6.5 B/yr</td><td>58%</td><td>High</td></tr>
    <tr><td>Akwa Ibom</td><td>5.4 M</td><td>NGN 6.0 B/yr</td><td>54%</td><td>High</td></tr>
    <tr><td>Edo</td><td>4.7 M</td><td>NGN 5.5 B/yr</td><td>51%</td><td>Medium</td></tr>
    <tr><td>Cross River</td><td>3.8 M</td><td>NGN 4.0 B/yr</td><td>46%</td><td>Medium</td></tr>
    <tr><td>Bayelsa</td><td>2.4 M</td><td>NGN 3.0 B/yr</td><td>41%</td><td>Medium</td></tr>
  </tbody>
</table>

<h3>2.3 Revenue Mechanics</h3>
<p>
  Under the SPV model, the state acquires equity in an MNO/ISP licensed for its territory.
  Revenue flows to the state via:
</p>
<ol>
  <li><strong>Dividend distributions</strong> — proportional to equity share and operator profit.</li>
  <li><strong>Management fees</strong> — for state-provided assets (spectrum, rights-of-way, infrastructure).</li>
  <li><strong>Licence revenue</strong> — direct fees from the NCC assigned to the sub-national entity.</li>
  <li><strong>Taxation</strong> — indirect IGR from employee income tax, corporate taxes, and vendor ecosystem activity.</li>
</ol>

<h3>2.4 Competitive Landscape</h3>
<p>
  At time of publication, no South-South state has operationalised an SPV under the new NCC framework.
  Consultancy interest is nascent, with two Lagos-based firms known to be preparing proposals for
  Rivers and Delta. The 6–9 month first-mover window is real and finite.
</p>


<!-- ── ExCS FRAMEWORK ── -->
<h2 class="section-title" id="framework">3. The ExCS IGNITION Framework</h2>

<h3>3.1 What is ExCS?</h3>
<p>
  <strong>ExCS</strong> (Executive Consulting Strategy) is NAKACHI's proprietary methodology for
  engaging government counterparts at the executive level — Governors, Deputy Governors,
  Commissioners for Finance, and Permanent Secretaries. ExCS combines structured stakeholder
  intelligence, bespoke financial modelling, and a sequenced engagement protocol to move
  decision-makers from awareness to mandate commitment within a defined timeframe.
</p>
<p>
  The IGNITION variant is purpose-built for <em>new market entry</em> scenarios — contexts where
  the opportunity is real but the institutional knowledge of the client is low, requiring
  NAKACHI to build both the business case and the organisational appetite simultaneously.
</p>

<h3>3.2 IGNITION's Five Pillars</h3>
<div class="card-grid">
  <div class="card">
    <div class="card-num">01</div>
    <div class="card-label">Intelligence</div>
    <div class="card-body">Deep stakeholder mapping: decision-makers, gatekeepers, champions, and
    blockers in each target state government.</div>
  </div>
  <div class="card">
    <div class="card-num">02</div>
    <div class="card-label">Narrative</div>
    <div class="card-body">Bespoke financial models and impact narratives tailored to each state's
    existing IGR base, demographics, and development priorities.</div>
  </div>
  <div class="card">
    <div class="card-num">03</div>
    <div class="card-label">Insertion</div>
    <div class="card-body">Multi-channel outreach — direct mail, introductory briefings, trusted
    intermediary activation — to secure initial access.</div>
  </div>
  <div class="card">
    <div class="card-num">04</div>
    <div class="card-label">Transformation</div>
    <div class="card-body">Workshop facilitation and roadmap development that converts a curious
    counterpart into a committed sponsor.</div>
  </div>
  <div class="card">
    <div class="card-num">05</div>
    <div class="card-label">Ignition</div>
    <div class="card-body">Signed engagement mandate, project mobilisation, and first milestone
    delivery — the point of no return toward live revenue.</div>
  </div>
</div>

<h3>3.3 Engagement Sequencing Model</h3>
<table class="phase-table">
  <thead>
    <tr><th>Phase</th><th>Label</th><th>Duration</th><th>Primary Deliverable</th><th>Exit Criterion</th></tr>
  </thead>
  <tbody>
    <tr><td>0</td><td>Preparation</td><td>Weeks 1–3</td><td>Stakeholder maps, state briefs, outreach collateral</td><td>All materials approved</td></tr>
    <tr><td>1</td><td>Outreach</td><td>Weeks 4–8</td><td>DM Wave 1 dispatched; 2+ meetings secured per state</td><td>Meeting rate &gt;40%</td></tr>
    <tr><td>2</td><td>Discovery</td><td>Weeks 9–14</td><td>State-specific feasibility briefs; IGR impact models</td><td>Brief accepted by sponsor</td></tr>
    <tr><td>3</td><td>Proposal</td><td>Weeks 15–20</td><td>Full ExCS IGNITION proposal submitted</td><td>No rejection at 2 weeks</td></tr>
    <tr><td>4</td><td>Negotiation</td><td>Weeks 21–26</td><td>Engagement mandate negotiated and signed</td><td>Executed mandate</td></tr>
    <tr><td>5</td><td>Mobilisation</td><td>Weeks 27–36</td><td>SPV structure, NCC application, operator partner</td><td>NCC application filed</td></tr>
    <tr><td>6</td><td>First Revenue</td><td>Months 12–18</td><td>Commercial launch; first IGR receipts</td><td>Revenue &gt; NGN 50M/month</td></tr>
  </tbody>
</table>


<!-- ── GTM STRATEGY ── -->
<h2 class="section-title" id="gtm">4. Go-To-Market Strategy</h2>

<h3>4.1 Channel Strategy</h3>
<p>
  The ExCS IGNITION GTM strategy deploys four primary channels in coordinated waves,
  with the Direct Mail sequence (Wave 1) serving as the primary cold outreach vehicle.
</p>

<h4>Channel 1 — Direct Mail</h4>
<p>
  A four-letter direct mail sequence targeting Governors and Commissioners of Finance in all six
  states. Each letter is individually personalised with state-specific data. The sequence is
  designed on a "reluctance-reduction" model: each successive letter addresses the most likely
  objection raised by a reader who has not yet responded to the previous letter.
</p>
<ul>
  <li><strong>Letter 1 (Attention):</strong> The NCC policy change and what it means specifically for the recipient's state.</li>
  <li><strong>Letter 2 (Interest):</strong> The financial upside — state-specific IGR projections with credible assumptions.</li>
  <li><strong>Letter 3 (Desire):</strong> Case studies and peer comparisons; urgency framing around the first-mover window.</li>
  <li><strong>Letter 4 (Action):</strong> A direct, time-bounded invitation to a private briefing with NAKACHI's senior partner.</li>
</ul>

<h4>Channel 2 — Trusted Intermediaries</h4>
<p>
  Activation of NAKACHI's network of former commissioners, senior civil servants, and development
  bank officials with existing relationships in each target state. Each intermediary is briefed
  separately and provided with the public summary document (excs_ignition_public.html) as
  a conversation starter.
</p>

<h4>Channel 3 — Events &amp; Conferences</h4>
<p>
  Targeted presence at the Annual Governors' Forum, the South-South Economic Summit, and the
  NCC Industry Dialogue Series. NAKACHI will secure speaking slots where possible, and position
  ExCS IGNITION as the de facto framework through which states should approach the NCC opportunity.
</p>

<h4>Channel 4 — Digital Presence</h4>
<p>
  A restricted-access microsite mirroring the public HTML document, distributed via personalised
  URLs (PURLs) in the DM sequence to allow tracking of engagement. LinkedIn outreach to deputy
  governors and senior commissioner-level officials.
</p>

<h3>4.2 Messaging Architecture</h3>
<table>
  <thead>
    <tr><th>Audience</th><th>Primary Message</th><th>Proof Point</th><th>CTA</th></tr>
  </thead>
  <tbody>
    <tr>
      <td>Governor</td>
      <td>Transform telecom revenues into a permanent IGR stream — without borrowing</td>
      <td>State-specific NGN projection</td>
      <td>Request private briefing</td>
    </tr>
    <tr>
      <td>Commissioner for Finance</td>
      <td>The SPV model is low-risk, off-balance-sheet, and NCC-compliant</td>
      <td>Legal opinion summary + NCC guidance note</td>
      <td>Download full feasibility brief</td>
    </tr>
    <tr>
      <td>Permanent Secretary</td>
      <td>NAKACHI provides end-to-end implementation — your office owns the outcome, not the workload</td>
      <td>Phase plan with PS role defined at each stage</td>
      <td>Schedule working session</td>
    </tr>
    <tr>
      <td>Technical Advisors</td>
      <td>Proven methodology, NCC-cleared, with operator partnership options already identified</td>
      <td>NAKACHI credentials + operator MOU template</td>
      <td>Access technical annexe</td>
    </tr>
  </tbody>
</table>

<h3>4.3 Objection-Handling Framework</h3>
<table>
  <thead>
    <tr><th>Objection</th><th>Category</th><th>Response Approach</th></tr>
  </thead>
  <tbody>
    <tr><td>"We don't have capacity to manage a telco"</td><td>Capability</td>
        <td>SPV uses professional management; state role is governance + dividend receipt only</td></tr>
    <tr><td>"The private sector won't accept state equity"</td><td>Market</td>
        <td>State equity reduces sovereign risk for lenders; MNC operators already expressing interest</td></tr>
    <tr><td>"This will take too long before we see money"</td><td>Timeline</td>
        <td>Management fees start from month 6; first dividends by month 14 at conservative assumptions</td></tr>
    <tr><td>"We've heard this before — consultants talk, nothing happens"</td><td>Trust</td>
        <td>Success-fee element in NAKACHI engagement; no results, no full fee</td></tr>
    <tr><td>"What happens if the NCC changes policy again?"</td><td>Risk</td>
        <td>SPV structure is robust to policy reversal; regulatory risk analysis provided</td></tr>
  </tbody>
</table>


<!-- ── TARGET SEGMENTS ── -->
<h2 class="section-title" id="segments">5. Target Segments &amp; Stakeholder Intelligence</h2>

<h3>5.1 Priority Tier 1 States</h3>
<p>Rivers, Delta, and Akwa Ibom are designated <strong>Tier 1</strong> based on population size,
current IGR trajectory, and perceived institutional appetite for innovation.</p>

<h3>5.2 Priority Tier 2 States</h3>
<p>Edo, Cross River, and Bayelsa are <strong>Tier 2</strong> — viable opportunities but with lower
near-term probability scores, warranting lighter outreach investment in Wave 1.</p>

<h3>5.3 Stakeholder Roles</h3>
<table>
  <thead>
    <tr><th>Role</th><th>Title Examples</th><th>Engagement Mode</th><th>Decision Weight</th></tr>
  </thead>
  <tbody>
    <tr><td>Economic Sponsor</td><td>Governor, Deputy Governor</td><td>DM + private briefing</td><td>Final</td></tr>
    <tr><td>Financial Gatekeeper</td><td>Commissioner for Finance, AG</td><td>DM + technical brief</td><td>High</td></tr>
    <tr><td>Implementation Owner</td><td>Perm Sec (Finance/ICT)</td><td>Working session</td><td>Medium</td></tr>
    <tr><td>Technical Validator</td><td>State ICT Director</td><td>Technical annexe + call</td><td>Medium</td></tr>
    <tr><td>Trusted Gatekeeper</td><td>Chief of Staff, SPA (Econ)</td><td>Intermediary activation</td><td>Access</td></tr>
  </tbody>
</table>


<!-- ── FINANCIAL PROJECTIONS ── -->
<h2 class="section-title" id="financial">6. Financial Projections</h2>

<h3>6.1 Revenue Model Assumptions</h3>
<ul>
  <li>Average Revenue Per User (ARPU): NGN 2,500/month</li>
  <li>Target Market Penetration: 15% of state population in Year 1, growing to 22% by Year 3</li>
  <li>State SPV Equity: 40% of operator entity</li>
  <li>Operator EBITDA Margin: 28% (conservative; sector median is 34%)</li>
  <li>Dividend Payout Ratio: 60% of net profit</li>
</ul>

<h3>6.2 State-by-State Projections (Year 1, Base Case)</h3>
<table>
  <thead>
    <tr><th>State</th><th>Subscribers (Y1)</th><th>Operator Revenue</th>
        <th>State EBITDA Share</th><th>Annual Dividend to State</th></tr>
  </thead>
  <tbody>
    <tr><td>Rivers</td><td>1,080,000</td><td>NGN 32.4 B</td><td>NGN 3.6 B</td><td>NGN 2.2 B</td></tr>
    <tr><td>Delta</td><td>840,000</td><td>NGN 25.2 B</td><td>NGN 2.8 B</td><td>NGN 1.7 B</td></tr>
    <tr><td>Akwa Ibom</td><td>810,000</td><td>NGN 24.3 B</td><td>NGN 2.7 B</td><td>NGN 1.6 B</td></tr>
    <tr><td>Edo</td><td>705,000</td><td>NGN 21.2 B</td><td>NGN 2.4 B</td><td>NGN 1.4 B</td></tr>
    <tr><td>Cross River</td><td>570,000</td><td>NGN 17.1 B</td><td>NGN 1.9 B</td><td>NGN 1.2 B</td></tr>
    <tr><td>Bayelsa</td><td>360,000</td><td>NGN 10.8 B</td><td>NGN 1.2 B</td><td>NGN 0.7 B</td></tr>
  </tbody>
</table>

<h3>6.3 NAKACHI Fee Structure</h3>
<p>
  NAKACHI's ExCS IGNITION engagement is structured as a blended retainer + success fee model
  to align interests and reduce the state's financial risk:
</p>
<ul>
  <li><strong>Mobilisation Retainer:</strong> NGN 150M (paid in three tranches over 12 months)</li>
  <li><strong>NCC Application Success Fee:</strong> NGN 50M upon filed application</li>
  <li><strong>Commercial Launch Bonus:</strong> NGN 100M upon first revenue receipt exceeding NGN 100M/month</li>
  <li><strong>Year 2 Advisory Retainer:</strong> NGN 80M/year (optional renewal)</li>
</ul>
<p>Total engagement value: approximately NGN 380M per state over the initial programme period.</p>


<!-- ── TIMELINE ── -->
<h2 class="section-title" id="timeline">7. Programme Timeline</h2>

<div class="timeline">
  <div class="tl-event">
    <div class="tl-date">May–June 2026 · Phase 0</div>
    <h4>Preparation &amp; Asset Finalisation</h4>
    <p>Stakeholder intelligence compiled. State briefs authored. Direct Mail Wave 1 printed and
    dispatched. Public summary document uploaded to NAKACHI KB.</p>
  </div>
  <div class="tl-event">
    <div class="tl-date">June–August 2026 · Phase 1</div>
    <h4>Outreach Wave 1</h4>
    <p>Four-letter DM sequence dispatched to all Tier 1 and Tier 2 targets. Intermediary
    network activated. Conference presence at South-South Economic Summit secured.</p>
  </div>
  <div class="tl-event">
    <div class="tl-date">August–October 2026 · Phase 2</div>
    <h4>Discovery Engagements</h4>
    <p>Private briefings held with interested states. State-specific feasibility briefs authored
    and delivered. Second DM wave deployed to non-responders from Wave 1.</p>
  </div>
  <div class="tl-event">
    <div class="tl-date">October–December 2026 · Phase 3–4</div>
    <h4>Proposal &amp; Negotiation</h4>
    <p>Full ExCS IGNITION proposals submitted to at least three states. Engagement mandates
    under negotiation. Target: 2 signed mandates by December 2026.</p>
  </div>
  <div class="tl-event">
    <div class="tl-date">January–September 2027 · Phase 5</div>
    <h4>Mobilisation &amp; NCC Application</h4>
    <p>SPV incorporation. Operator partner selection and MOU execution. NCC application
    filed. Spectrum allocation process commenced.</p>
  </div>
  <div class="tl-event">
    <div class="tl-date">Q3–Q4 2027 · Phase 6</div>
    <h4>Commercial Launch &amp; First Revenue</h4>
    <p>Pilot commercial launch in lead state. First revenue receipts to state SPV.
    NAKACHI commercial launch bonus triggered.</p>
  </div>
</div>


<!-- ── KPIs ── -->
<h2 class="section-title" id="kpis">8. Key Performance Indicators</h2>

<div class="kpi-row">
  <div class="kpi-box">
    <div class="kpi-val">≥40%</div>
    <div class="kpi-lbl">DM Wave 1 Meeting Rate</div>
  </div>
  <div class="kpi-box">
    <div class="kpi-val">3</div>
    <div class="kpi-lbl">Discovery Briefings by Oct '26</div>
  </div>
  <div class="kpi-box">
    <div class="kpi-val">2</div>
    <div class="kpi-lbl">Signed Mandates by Dec '26</div>
  </div>
  <div class="kpi-box">
    <div class="kpi-val">1</div>
    <div class="kpi-lbl">NCC Application Filed by Q2 '27</div>
  </div>
  <div class="kpi-box">
    <div class="kpi-val">NGN 100M</div>
    <div class="kpi-lbl">Monthly Revenue Threshold by Q4 '27</div>
  </div>
</div>

<h3>8.1 Leading Indicators (Monthly Tracking)</h3>
<table>
  <thead>
    <tr><th>Metric</th><th>Target</th><th>Frequency</th><th>Owner</th></tr>
  </thead>
  <tbody>
    <tr><td>DM letters dispatched</td><td>100% of list by Week 8</td><td>Weekly</td><td>Outreach Manager</td></tr>
    <tr><td>PURL open rate</td><td>≥35%</td><td>Weekly</td><td>Digital Lead</td></tr>
    <tr><td>Meeting requests received</td><td>≥12 by Week 12</td><td>Weekly</td><td>BD Lead</td></tr>
    <tr><td>Briefings completed</td><td>≥6 by Month 5</td><td>Monthly</td><td>Partner</td></tr>
    <tr><td>Proposals submitted</td><td>≥3 by Month 7</td><td>Monthly</td><td>Partner</td></tr>
    <tr><td>Mandates executed</td><td>≥2 by Month 9</td><td>Monthly</td><td>Managing Partner</td></tr>
  </tbody>
</table>


<!-- ── RISK ── -->
<h2 class="section-title" id="risk">9. Risk Register</h2>

<div class="risk-grid">
  <div class="risk-item high">
    <h4>NCC Policy Reversal</h4>
    <p><strong>Likelihood:</strong> Low &nbsp;|&nbsp; <strong>Impact:</strong> Critical</p>
    <p>Monitor NCC quarterly bulletins. Engage NCC through industry body. Structure SPV to be
    convertible if policy rolls back.</p>
  </div>
  <div class="risk-item high">
    <h4>Competitor Captures Lead State</h4>
    <p><strong>Likelihood:</strong> Medium &nbsp;|&nbsp; <strong>Impact:</strong> High</p>
    <p>Accelerate Rivers State engagement to Q1. Secure intermediary exclusivity in Tier 1 states
    where possible.</p>
  </div>
  <div class="risk-item med">
    <h4>Political Transition (Elections)</h4>
    <p><strong>Likelihood:</strong> Medium &nbsp;|&nbsp; <strong>Impact:</strong> Medium</p>
    <p>Engage both incumbent and likely successor administrations. Build civil service champions
    who survive transition.</p>
  </div>
  <div class="risk-item med">
    <h4>Operator Partner Withdrawal</h4>
    <p><strong>Likelihood:</strong> Low &nbsp;|&nbsp; <strong>Impact:</strong> High</p>
    <p>Maintain relationships with three operator candidates simultaneously. Operator MOU includes
    exclusivity break fee.</p>
  </div>
  <div class="risk-item low">
    <h4>State Budget Constraints</h4>
    <p><strong>Likelihood:</strong> Medium &nbsp;|&nbsp; <strong>Impact:</strong> Medium</p>
    <p>Structure mobilisation retainer as success-linked. Provide bridge financing options through
    NAKACHI's DFI relationships.</p>
  </div>
  <div class="risk-item low">
    <h4>Technical Execution Delays</h4>
    <p><strong>Likelihood:</strong> Medium &nbsp;|&nbsp; <strong>Impact:</strong> Low</p>
    <p>Use proven operator partner with existing South-South infrastructure. NAKACHI PM embedded
    in project team.</p>
  </div>
</div>


<!-- ── APPENDIX ── -->
<h2 class="section-title" id="appendix">10. Appendix</h2>

<div class="appendix">
  <h3>Appendix A — Document Register</h3>
  <table>
    <thead>
      <tr><th>Document</th><th>File Name</th><th>Status</th><th>Location</th></tr>
    </thead>
    <tbody>
      <tr><td>Master GTM Control</td><td>excs_ignition_control.html</td><td>Published</td><td>NAKACHI KB</td></tr>
      <tr><td>Public Summary</td><td>excs_ignition_public.html</td><td>Published</td><td>NAKACHI KB</td></tr>
      <tr><td>Direct Mail Sequence</td><td>NAKACHI_Direct_Mail_Wave1.docx</td><td>Published</td><td>NAKACHI KB</td></tr>
      <tr><td>NCC Policy Guidance Note</td><td>ncc_policy_review_guidance.pdf</td><td>Pending</td><td>TBC</td></tr>
      <tr><td>SPV Legal Opinion</td><td>excs_spv_legal_opinion.pdf</td><td>Pending</td><td>TBC</td></tr>
      <tr><td>Operator MOU Template</td><td>operator_mou_template.docx</td><td>Pending</td><td>TBC</td></tr>
    </tbody>
  </table>
</div>

<div class="appendix">
  <h3>Appendix B — Glossary</h3>
  <table>
    <thead><tr><th>Term</th><th>Definition</th></tr></thead>
    <tbody>
      <tr><td>ExCS</td><td>Executive Consulting Strategy — NAKACHI's government-facing methodology</td></tr>
      <tr><td>IGNITION</td><td>The new-market-entry variant of the ExCS framework</td></tr>
      <tr><td>SPV</td><td>Special Purpose Vehicle — a ring-fenced corporate entity for the telecom equity</td></tr>
      <tr><td>IGR</td><td>Internally Generated Revenue — non-federal revenue raised by state governments</td></tr>
      <tr><td>NCC</td><td>Nigerian Communications Commission — the sector regulator</td></tr>
      <tr><td>MNO</td><td>Mobile Network Operator</td></tr>
      <tr><td>ISP</td><td>Internet Service Provider</td></tr>
      <tr><td>ARPU</td><td>Average Revenue Per User</td></tr>
      <tr><td>EBITDA</td><td>Earnings Before Interest, Tax, Depreciation &amp; Amortisation</td></tr>
      <tr><td>PURL</td><td>Personalised URL — trackable web link used in the DM sequence</td></tr>
      <tr><td>DFI</td><td>Development Finance Institution</td></tr>
    </tbody>
  </table>
</div>

<div class="appendix">
  <h3>Appendix C — NAKACHI Consulting Credentials</h3>
  <p>
    NAKACHI Consulting is a strategy and implementation firm specialising in public-sector revenue
    optimisation, infrastructure finance, and technology-enabled governance reform across
    Sub-Saharan Africa. Founded with a focus on the Nigerian market, NAKACHI combines deep
    government relationships with sector-specific technical expertise to deliver outcomes
    that outlast the engagement.
  </p>
  <p>
    The ExCS IGNITION team brings together former NCC commissioners, Big-4 telecoms partners,
    DFI transaction advisors, and state government veterans — a combination specifically
    assembled to navigate the political, technical, and financial dimensions of this initiative.
  </p>
  <p>
    <strong>Contact:</strong> ExCS IGNITION Programme Office &nbsp;|&nbsp;
    NAKACHI Consulting &nbsp;|&nbsp; Lagos, Nigeria<br>
    For secure document delivery, use NAKACHI AI Knowledge Base (samikoku@gmail.com).
  </p>
</div>

</main>

<footer>
  <strong>NAKACHI Consulting</strong> &nbsp;·&nbsp; ExCS IGNITION GTM Package &nbsp;·&nbsp;
  Version 1.0 &nbsp;·&nbsp; May 2026<br>
  <span style="font-size:0.78rem; opacity:.7;">
    This document is confidential and intended solely for internal NAKACHI use and
    authorised state government counterparts. Unauthorised distribution is prohibited.
  </span>
</footer>

</body>
</html>
"""

# ─────────────────────────────────────────────
# 2.  excs_ignition_public.html  (~12.5 KB)
# ─────────────────────────────────────────────
PUBLIC_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ExCS IGNITION | NAKACHI Consulting</title>
<style>
  :root { --primary:#0a2342; --accent:#e63946; --gold:#f4a261; }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Segoe UI',Arial,sans-serif; color:#212529; background:#fff;
         line-height:1.75; font-size:16px; }
  header { background:var(--primary); color:#fff; padding:56px 48px; text-align:center; }
  header .eyebrow { color:var(--gold); text-transform:uppercase; letter-spacing:4px;
                    font-size:0.85rem; margin-bottom:12px; }
  header h1 { font-size:3rem; letter-spacing:2px; margin-bottom:16px; }
  header p  { font-size:1.15rem; max-width:640px; margin:0 auto; opacity:.85; }
  .strip { background:var(--accent); color:#fff; text-align:center; padding:14px;
           font-size:0.9rem; font-weight:600; letter-spacing:1px; }
  section { max-width:900px; margin:0 auto; padding:56px 32px; }
  h2 { font-size:1.8rem; color:var(--primary); margin-bottom:20px;
       padding-bottom:12px; border-bottom:3px solid var(--accent); }
  p  { margin-bottom:16px; }
  .value-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
                gap:20px; margin:32px 0; }
  .value-card { background:#f8f9fa; border-radius:10px; padding:28px 24px;
                border-top:4px solid var(--accent); }
  .value-card h3 { font-size:1.1rem; color:var(--primary); margin-bottom:8px; }
  .value-card p  { font-size:0.9rem; margin:0; }
  .stat-row { display:flex; gap:16px; flex-wrap:wrap; justify-content:center;
              margin:36px 0; }
  .stat { background:var(--primary); color:#fff; border-radius:10px; padding:28px;
          text-align:center; flex:1; min-width:160px; }
  .stat .num { font-size:2.2rem; font-weight:700; color:var(--gold); }
  .stat .lbl { font-size:0.8rem; opacity:.8; margin-top:6px; text-transform:uppercase;
               letter-spacing:1px; }
  .steps { counter-reset:step; margin:24px 0; }
  .step  { display:flex; gap:20px; align-items:flex-start; margin-bottom:24px; }
  .step-num { background:var(--accent); color:#fff; border-radius:50%;
              width:40px; height:40px; display:flex; align-items:center;
              justify-content:center; font-weight:700; flex-shrink:0; font-size:1.1rem; }
  .step-body h3 { font-size:1rem; color:var(--primary); margin-bottom:4px; }
  .cta { background:var(--accent); color:#fff; text-align:center; padding:52px 32px; }
  .cta h2 { color:#fff; border-color:rgba(255,255,255,.4); }
  .cta p  { opacity:.9; max-width:560px; margin:0 auto 28px; }
  .cta-btn { display:inline-block; background:#fff; color:var(--accent);
             font-weight:700; padding:16px 40px; border-radius:6px;
             text-decoration:none; font-size:1rem; letter-spacing:1px; }
  footer { background:#1b3a5e; color:#adb5bd; text-align:center;
           padding:28px; font-size:0.85rem; }
  footer strong { color:var(--gold); }
</style>
</head>
<body>

<header>
  <div class="eyebrow">NAKACHI Consulting · ExCS Division</div>
  <h1>ExCS IGNITION</h1>
  <p>Helping South-South Nigerian state governments unlock a new, permanent source of
     Internally Generated Revenue through the NCC telecom equity framework.</p>
</header>

<div class="strip">
  First-Mover Advantage Is Available — But Only for the Next 18 Months
</div>

<section>
  <h2>The Opportunity</h2>
  <p>
    The Nigerian Communications Commission's 2025 Policy Review now permits state governments
    to hold equity in licensed telecom operators serving their populations. For the first time,
    the revenue your citizens pay for mobile and internet services can flow — in part — back to
    your state treasury.
  </p>
  <p>
    NAKACHI Consulting's <strong>ExCS IGNITION</strong> programme is the structured pathway
    from policy awareness to first revenue receipt. We have done the legal analysis, built
    the financial models, and identified the operator partners. Your state can move faster
    because we have already done the preparatory work.
  </p>

  <div class="stat-row">
    <div class="stat"><div class="num">6</div><div class="lbl">South-South States</div></div>
    <div class="stat"><div class="num">NGN 2.4T</div><div class="lbl">Combined IGR Potential</div></div>
    <div class="stat"><div class="num">18 Mo.</div><div class="lbl">Path to First Revenue</div></div>
    <div class="stat"><div class="num">40%</div><div class="lbl">Recommended SPV Equity</div></div>
  </div>
</section>

<section style="background:#f8f9fa; max-width:100%; padding:56px 0;">
  <div style="max-width:900px; margin:0 auto; padding:0 32px;">
    <h2>What ExCS IGNITION Delivers</h2>
    <div class="value-grid">
      <div class="value-card">
        <h3>Financial Clarity</h3>
        <p>A state-specific IGR impact model showing exactly how much your SPV equity stake
        is worth under three scenarios — conservative, base, and optimistic.</p>
      </div>
      <div class="value-card">
        <h3>Legal Certainty</h3>
        <p>NCC-reviewed SPV structure, legal opinions from leading telecoms counsel, and
        a complete compliance roadmap.</p>
      </div>
      <div class="value-card">
        <h3>Operator Partnership</h3>
        <p>NAKACHI has pre-qualified three operator candidates with proven South-South
        infrastructure. You inherit a relationship, not a cold search.</p>
      </div>
      <div class="value-card">
        <h3>Implementation Support</h3>
        <p>End-to-end programme management from SPV incorporation through NCC application
        to commercial launch — your office governs, we execute.</p>
      </div>
    </div>
  </div>
</section>

<section>
  <h2>How It Works</h2>
  <div class="steps">
    <div class="step">
      <div class="step-num">1</div>
      <div class="step-body">
        <h3>Discovery Briefing (Week 1–2)</h3>
        <p>A private 90-minute session with NAKACHI's senior partner. We present your
        state's specific financial model and answer questions without obligation.</p>
      </div>
    </div>
    <div class="step">
      <div class="step-num">2</div>
      <div class="step-body">
        <h3>Feasibility &amp; Proposal (Months 1–3)</h3>
        <p>Deep-dive feasibility study. SPV legal structure designed. Operator partner
        shortlisted. Full ExCS IGNITION proposal submitted.</p>
      </div>
    </div>
    <div class="step">
      <div class="step-num">3</div>
      <div class="step-body">
        <h3>Mobilisation (Months 4–9)</h3>
        <p>SPV incorporated. NCC application filed. Operator MOU executed. Management
        team recruited. Infrastructure plan agreed.</p>
      </div>
    </div>
    <div class="step">
      <div class="step-num">4</div>
      <div class="step-body">
        <h3>Commercial Launch (Months 12–18)</h3>
        <p>Services launched. First subscribers acquired. First revenue receipts flow to
        state SPV. NAKACHI transitions to advisory role.</p>
      </div>
    </div>
  </div>
</section>

<div class="cta">
  <h2>Request a Private Briefing</h2>
  <p>
    This programme is available to all six South-South states, but NAKACHI works with
    a maximum of three states simultaneously to protect quality. Availability for 2026
    onboarding is limited.
  </p>
  <a class="cta-btn" href="mailto:ignition@nakachiconsulting.com?subject=ExCS IGNITION Briefing Request">
    Request Your Briefing
  </a>
</div>

<footer>
  <strong>NAKACHI Consulting</strong> &nbsp;·&nbsp; ExCS IGNITION &nbsp;·&nbsp; May 2026<br>
  <span style="font-size:0.8rem; opacity:.7;">
    © 2026 NAKACHI Consulting. All rights reserved. For enquiries: ignition@nakachiconsulting.com
  </span>
</footer>

</body>
</html>
"""

# ─────────────────────────────────────────────
# 3.  NAKACHI_Direct_Mail_Wave1.docx
# ─────────────────────────────────────────────
def create_docx(filepath):
    doc = Document()

    # ── Document-level styles ──────────────────
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    def heading(text, level=1):
        h = doc.add_heading(text, level=level)
        h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = h.runs[0]
        run.font.color.rgb = RGBColor(0x0a, 0x23, 0x42)
        return h

    def body(text):
        p = doc.add_paragraph(text)
        p.style = doc.styles['Normal']
        return p

    def signature():
        doc.add_paragraph()
        body("Warm regards,")
        doc.add_paragraph()
        p = doc.add_paragraph()
        r = p.add_run("Dr Nakachi A.\nManaging Partner, ExCS Division\nNAKACHI Consulting\nLagos, Nigeria\nignition@nakachiconsulting.com")
        r.font.bold = True
        doc.add_paragraph()

    def page_break():
        doc.add_page_break()

    # ──────────────────────────────────────────
    # COVER / HEADER
    # ──────────────────────────────────────────
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("NAKACHI CONSULTING\nExCS IGNITION — Direct Mail Wave 1")
    tr.font.size = Pt(16)
    tr.font.bold = True
    tr.font.color.rgb = RGBColor(0x0a, 0x23, 0x42)
    doc.add_paragraph()
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("Four-Letter Direct Mail Sequence · May 2026\nConfidential — Campaign Use Only")
    sr.font.size = Pt(10)
    sr.font.color.rgb = RGBColor(0x6c, 0x75, 0x7d)
    doc.add_paragraph()
    doc.add_paragraph("─" * 72)

    page_break()

    # ──────────────────────────────────────────
    # LETTER 1 — ATTENTION
    # ──────────────────────────────────────────
    heading("LETTER 1 OF 4 — ATTENTION", level=1)
    doc.add_paragraph("Subject: A New Source of Revenue for [STATE] State — NCC Policy Change")
    doc.add_paragraph("Date: [DATE]")
    doc.add_paragraph()
    body("His Excellency [GOVERNOR NAME]\nGovernor, [STATE] State\nGovernment House, [CAPITAL CITY]")
    doc.add_paragraph()
    body("Your Excellency,")
    doc.add_paragraph()
    body(
        "I write to draw your attention to a regulatory development that creates a direct "
        "and significant new revenue opportunity for [STATE] State — one that, to my "
        "knowledge, no South-South state has yet moved to capture."
    )
    body(
        "In Q4 2025, the Nigerian Communications Commission finalised amendments to the "
        "National Telecommunications Policy Framework. Under the revised framework, state "
        "governments are now explicitly permitted to hold equity — up to 49 percent — in "
        "licensed mobile and internet service operators providing services within their territory."
    )
    body(
        "For [STATE] State, with a population of [POPULATION] and existing mobile penetration "
        "of approximately [PENETRATION]%, this means that a properly structured Special Purpose "
        "Vehicle (SPV) — holding as little as 40% equity in a licensed operator — could generate "
        "an estimated NGN [ANNUAL_PROJECTION] per annum in additional Internally Generated Revenue, "
        "beginning within 18 months of a decision to proceed."
    )
    body(
        "NAKACHI Consulting has been preparing for this opportunity since the policy consultation "
        "stage. We have the legal structure, the financial models, and the operator relationships "
        "already in place. What is required now is a state executive prepared to move."
    )
    body(
        "I would welcome the opportunity to present this analysis to you or your Commissioner "
        "for Finance in a private 90-minute briefing at a time of your choosing. There is no "
        "obligation, and the analysis will be delivered at no cost."
    )
    body(
        "The window for first-mover advantage in the South-South zone is real and, by our "
        "assessment, will close within 12 to 18 months as competitors and other states begin "
        "to activate. I believe [STATE] State is positioned to lead."
    )
    signature()
    body("[P.S. A one-page summary of the financial projections specific to [STATE] State is enclosed.]")
    page_break()

    # ──────────────────────────────────────────
    # LETTER 2 — INTEREST
    # ──────────────────────────────────────────
    heading("LETTER 2 OF 4 — INTEREST", level=1)
    doc.add_paragraph("Subject: The Numbers for [STATE] State — What the Telecom Equity Model Actually Yields")
    doc.add_paragraph("Date: [DATE + 10 DAYS]")
    doc.add_paragraph()
    body("Your Excellency,")
    doc.add_paragraph()
    body(
        "I wrote to you ten days ago about the NCC Policy Review and the revenue opportunity "
        "it creates for [STATE] State. I recognise that correspondence of this nature arrives "
        "alongside many competing demands on your attention. I therefore want to make the "
        "financial case as concrete and specific as possible."
    )
    body("Here is what the model shows for [STATE] State under conservative assumptions:")
    doc.add_paragraph()
    table = doc.add_table(rows=6, cols=2)
    table.style = 'Table Grid'
    rows_data = [
        ("Population of [STATE] State", "[POPULATION]"),
        ("Target subscribers at 15% penetration (Year 1)", "[SUBSCRIBERS]"),
        ("Average Revenue Per User per month", "NGN 2,500"),
        ("Operator monthly revenue", "NGN [MONTHLY_REV]"),
        ("State SPV share at 40% equity (monthly)", "NGN [MONTHLY_STATE]"),
        ("Annual IGR contribution to State", "NGN [ANNUAL_STATE]"),
    ]
    for i, (label, value) in enumerate(rows_data):
        table.rows[i].cells[0].text = label
        table.rows[i].cells[1].text = value
    doc.add_paragraph()
    body(
        "These figures assume only 15% penetration — conservative relative to the current "
        "national average of 49%. At 20% penetration, the annual contribution rises to "
        "approximately NGN [UPSIDE_20PCT]. At 25%, it approaches NGN [UPSIDE_25PCT]."
    )
    body(
        "The state's exposure is limited. The SPV structure keeps telecom equity off the "
        "state's balance sheet. The NCC application cost is under NGN 50 million. The "
        "mobilisation investment — across legal, regulatory, and operator negotiation — "
        "is recovered within the first four months of commercial operation."
    )
    body(
        "I am aware that [STATE] State has ambitious development plans that require sustainable, "
        "non-oil revenue. This model does not require borrowing, does not depend on federal "
        "allocations, and creates a revenue stream that compounds as subscriber penetration grows "
        "year on year."
    )
    body(
        "I would still welcome the briefing I proposed in my earlier letter. If it is more "
        "convenient, my colleague and ExCS IGNITION technical lead would be glad to meet "
        "with your Commissioner for Finance or Permanent Secretary (Finance) as a first step."
    )
    body("I remain at your service.")
    signature()
    page_break()

    # ──────────────────────────────────────────
    # LETTER 3 — DESIRE
    # ──────────────────────────────────────────
    heading("LETTER 3 OF 4 — DESIRE", level=1)
    doc.add_paragraph("Subject: Why [STATE] State Must Move Now — and What Happens If It Doesn't")
    doc.add_paragraph("Date: [DATE + 24 DAYS]")
    doc.add_paragraph()
    body("Your Excellency,")
    doc.add_paragraph()
    body(
        "This is my third letter on the subject of the NCC telecom equity opportunity. "
        "I write again not to repeat the financial case — you have those numbers — but to "
        "address the question of urgency, because it is the question I most commonly hear "
        "from state officials who are interested but have not yet moved."
    )
    body("The question is: 'Why now? Why not wait and see how other states do it first?'")
    body(
        "The answer is spectrum. The NCC allocates spectrum on a first-come, first-served "
        "basis at the sub-national level. The states that file SPV applications earliest will "
        "receive the most favourable spectrum bands — the bands that carry the highest subscriber "
        "capacity, the best rural coverage economics, and the strongest competitive moat against "
        "later entrants."
    )
    body(
        "A state that waits 18 months to observe others will find that the premium spectrum "
        "has been assigned, operator partners have committed exclusivities elsewhere, and the "
        "NCC is managing a queue. The revenue projections I have shared assume access to "
        "Tier 1 spectrum. A late entrant's projections would be materially lower."
    )
    body(
        "To be transparent: NAKACHI is in discussion with two other South-South states. "
        "We are in a position to begin mobilisation with one or at most two states in the "
        "2026 cohort. We would prefer that [STATE] State be among them."
    )
    body(
        "I have enclosed a one-page comparison showing the projected IGR trajectory for "
        "a state that enters in 2026 versus a state that enters in 2028. The gap in cumulative "
        "revenue over five years, at base case assumptions, exceeds NGN [LATE_ENTRY_GAP]."
    )
    body(
        "Your Excellency, the opportunity is real. The risk is modest. The window is finite. "
        "I respectfully request 90 minutes of your time — or that of your senior financial "
        "officers — to present the full ExCS IGNITION programme and answer every question "
        "your team has."
    )
    signature()
    body("[Enclosure: First-Mover vs. Late-Entrant IGR Comparison — [STATE] State]")
    page_break()

    # ──────────────────────────────────────────
    # LETTER 4 — ACTION
    # ──────────────────────────────────────────
    heading("LETTER 4 OF 4 — ACTION", level=1)
    doc.add_paragraph("Subject: A Direct Invitation — Private Briefing, [PROPOSED DATE]")
    doc.add_paragraph("Date: [DATE + 38 DAYS]")
    doc.add_paragraph()
    body("Your Excellency,")
    doc.add_paragraph()
    body(
        "I have written to you three times regarding the ExCS IGNITION programme and the "
        "NCC telecom equity opportunity for [STATE] State. I appreciate that you and your "
        "office manage extraordinary volumes of correspondence, and I understand that my "
        "earlier letters may not have reached you directly."
    )
    body(
        "I am therefore writing to extend a specific, time-bounded invitation:"
    )
    body(
        "NAKACHI Consulting respectfully requests a 90-minute private briefing with you, "
        "or with a senior representative of your choice, on [PROPOSED DATE] at Government "
        "House [CAPITAL CITY] or at a location of your preference in Lagos. I will personally "
        "lead the session."
    )
    body("The briefing will cover:")
    items = [
        "The NCC framework — what it allows and what it requires of the state",
        "[STATE]-specific financial modelling — three scenarios, fully documented assumptions",
        "SPV legal structure and compliance pathway",
        "Operator partner options and indicative commercial terms",
        "The ExCS IGNITION programme timeline and NAKACHI's role",
        "Answers to every question your team has",
    ]
    for item in items:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(item)
    doc.add_paragraph()
    body(
        "There is no fee for the briefing. There is no obligation to engage NAKACHI beyond it. "
        "My only ask is 90 minutes and a genuine conversation."
    )
    body(
        "To confirm attendance or to propose an alternative date, please respond to this "
        "letter directly, or contact my office at ignition@nakachiconsulting.com or "
        "+234 (0) 800 NAKACHI."
    )
    body(
        "Your Excellency, [STATE] State has the population, the resources, and the "
        "institutional capacity to lead the South-South zone into the new era of "
        "state-backed telecommunications. NAKACHI is ready to help you do exactly that."
    )
    body("I look forward to hearing from you.")
    signature()
    doc.add_paragraph()
    doc.add_paragraph("─" * 72)
    doc.add_paragraph()
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note.add_run(
        "NAKACHI Consulting · ExCS IGNITION · Direct Mail Wave 1\n"
        "Production File — Confidential — May 2026\n"
        "Personalise all bracketed fields before dispatch."
    ).font.size = Pt(9)

    doc.save(filepath)
    print(f"  ✓ {filepath}")


# ──────────────────────────────────────────────────
# Write files
# ──────────────────────────────────────────────────
control_path = os.path.join(OUTPUT_DIR, "excs_ignition_control.html")
public_path  = os.path.join(OUTPUT_DIR, "excs_ignition_public.html")
docx_path    = os.path.join(OUTPUT_DIR, "NAKACHI_Direct_Mail_Wave1.docx")

with open(control_path, "w", encoding="utf-8") as f:
    f.write(CONTROL_HTML)
print(f"  ✓ {control_path}")

with open(public_path, "w", encoding="utf-8") as f:
    f.write(PUBLIC_HTML)
print(f"  ✓ {public_path}")

create_docx(docx_path)

print()
print("File sizes:")
for p in [control_path, public_path, docx_path]:
    print(f"  {os.path.basename(p)}: {os.path.getsize(p)/1024:.1f} KB")
