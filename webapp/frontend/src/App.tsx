import { useState, useRef, useCallback, useEffect } from "react";
import { motion, useInView } from "framer-motion";
import { Radar, RadarChart, PolarGrid, PolarAngleAxis, ResponsiveContainer, Tooltip } from "recharts";
import gsap from "gsap";
import "./index.css";

// ── Data ──
const rgba = (hex: string, a: number) => {
  const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${a})`;
};

const PROJECTS = [
  { id:"ps", name:"P&S / P&T", headline:"Process-Driven Execution", color:"#002060", maturity:0.9,
    subtitle:"4-phase lifecycle in production. 264 slides of playbooks. Flight Level governance with RACI, KPIs, and Big Room Planning." },
  { id:"coe", name:"Digital CoE", headline:"Discovery-First Governance", color:"#0070C0", maturity:0.7,
    subtitle:"Separates Discovery from Delivery with feedback loops. 41-page JIRA Playbook. Portfolio Learn Cycle quarterly. HEVA cycle." },
  { id:"infra", name:"Infrastructure", headline:"Team-Centric Deployment", color:"#009688", maturity:0.5,
    subtitle:"Team Topologies + Pilot → Expansion → Optimization. Per-unit-served investment model. Built for multi-team rollout." },
  { id:"dto", name:"Renewables DTO", headline:"Lean Kanban Operations", color:"#D46B08", maturity:0.55,
    subtitle:"4 pillars: Intake → Prioritize → Deliver → Govern. Jira-only intake. Pull-based. WIP limits. Flow Metrics dashboard." },
];

const FLOWS = [
  { pid:"ps", nodes:["Area Requests","Shark Tank","Portfolio Committee","Discovery Phase","Solution Proposal","Project Execution","Ongoing / Close"],
    gates:[1,2], details:{
      philosophy:"Process-driven execution with 3-tier Flight Level governance (Portfolio → Program → Project). Most mature — running in production.",
      tools:"Entry: Email, ServiceNow, Teams → Execution: Jira, Confluence, Figma+Maze → Output: Jira Dashboards, PPMNow, Scorecards",
      prioritization:"PBI(x) = (BV+UX)/E with 1–10 weighted scoring. Business Value Criteria: Delight Users, Asset Health, Productivity, Regulatory, Safety. Effort: S/M/L.",
      roles:"PO: prioritizes backlog | SM: facilitates flow, WIP limits | PM: health dashboards | Portfolio Mgr: investment decisions | Senior Leadership: Shark Tank GO/No-Go"
    }},
  { pid:"coe", nodes:["Executive Idea","Lean Biz Case","Shark Tank","Think Tank","Discovery Board","Feedback Loop","Delivery Board","Release"],
    gates:[1,2,3], loop:5, details:{
      philosophy:"Discovery-first governance. No work enters Delivery without Discovery validation. Two separate Kanban boards connected by feedback loops.",
      tools:"Entry: Strategic Funnel → Execution: Jira (Discovery+Delivery), Ziflow, Maze, Figma, Monday → Output: Portfolio Learn Cycle, Sprint Report Canvas, Health Radars",
      prioritization:"PBI(x) = (BV+UX)/E at backlog level + HEVA cycle (Hypothesis→Experiment→Validation→Apprenticeship). COORG: Classify, Order, Organize. Planning Poker.",
      roles:"PO: moves items through discovery | Agile Master: hygiene checklist, sprint health | VP Digital: strategic alignment | Enterprise Architect: technical feasibility"
    }},
  { pid:"infra", nodes:["Request","Assessment","Pilot (1 team)","Learn & Measure","Expansion (7+ teams)","Optimization"],
    gates:[1], details:{
      philosophy:"Organizational change, not lifecycle definition. Team Topologies (4 types, 3 interaction modes) as structural foundation. Phased rollout with coaching.",
      tools:"Entry: Business requests → Execution: Jira, Businessmap → Output: Flow Metrics, Portfolio Roadmap, Investment Model",
      prioritization:"No formal formula. Portfolio Roadmap drives sequencing. Investment model per unit served. Blueprint Design for scoping.",
      roles:"Enterprise Agile Coach: leads rollout | Sensei Coach: daily team coaching | Sempai Coach: logistics | Portfolio Mgr: roadmap + investment"
    }},
  { pid:"dto", nodes:["Jira Intake","Prioritize","Pull & Deliver","Governance"],
    gates:[3], details:{
      philosophy:"Lean Kanban with zero tolerance for informal channels. ALL work via Jira. Pull-based — never pushed. 4 pillars: Intake → Prioritize → Deliver → Govern.",
      tools:"Entry: Jira Intake Ticket ONLY (no email/chat/meetings) → Execution: Jira 2-level (Epic Kanban + Project Kanban) → Output: Flow Metrics Dashboard, Blocker Tracker",
      prioritization:"Continuous (not batch). Scoring: Business Value × Urgency × Capacity. Allocation: Build 60%, BAU 25%, Upgrades 15%. Classes of Service: Standard, Expedite, Fixed Date, Intangible.",
      roles:"PO: triages intake, orders backlog | Dev Team: pulls work (never assigned) | QA Lead: pulls from Dev | Kanban Lead: flow metrics, cadences, blockers"
    }},
];

const DELIVERABLES = [
  { name:"E2E Model", ps:1, coe:.75, infra:.25, dto:.75, action:"COMBINE", reason:"P&S lifecycle (production-proven) + CoE feedback loops + Infra rollout + DTO pull-based execution" },
  { name:"Prioritization", ps:.75, coe:.5, infra:0, dto:.75, action:"ADAPT", reason:"PBI formula already used by P&S+CoE (familiarity). DTO adds continuous model + capacity allocation." },
  { name:"Intake", ps:1, coe:.25, infra:0, dto:1, action:"ADOPT", reason:"DTO's Jira-only intake solves Digital Products' #1 pain: fragmented channels. P&S adds governance layer." },
  { name:"RACI", ps:1, coe:.25, infra:.25, dto:.5, action:"ADOPT", reason:"PM Playbook has the only formalized RACI. DTO adds per-stage ownership. Avangrid defines team positioning." },
  { name:"Jira & Tools", ps:.75, coe:1, infra:.5, dto:1, action:"ADOPT", reason:"CoE JIRA Playbook (41p) most comprehensive. DTO 2-level board gives portfolio + execution visibility." },
  { name:"Portfolio", ps:.75, coe:.75, infra:.5, dto:.5, action:"ADAPT", reason:"nWOW Portfolio Kanban (clearest). CoE Learn Cycle (quarterly alignment). Avangrid investment model." },
  { name:"Baseline", ps:.25, coe:.25, infra:.25, dto:.75, action:"ADOPT", reason:"DTO deliverable IS a baseline (exec summary + findings per pillar) — exactly what proposal requires." },
  { name:"Pilot", ps:.25, coe:.25, infra:.5, dto:.75, action:"ADOPT", reason:"DTO most execution-ready. Avangrid's Pilot→Expansion→Optimization designed for multi-OPCO rollout." },
  { name:"Playbook", ps:1, coe:.75, infra:0, dto:.5, action:"ADAPT", reason:"P&S 264 slides of modular structure + CoE agile practices + DTO Kanban guide = 375+ slides to adapt." },
  { name:"KPIs", ps:1, coe:.75, infra:0, dto:.5, action:"ADAPT", reason:"PM Playbook has only full KPI hierarchy with Jira automation. CoE adds strategic metrics. DTO adds Flow Metrics." },
];

// ── Components ──

function Hero() {
  return (
    <section style={{ minHeight:"100vh", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", position:"relative" }} className="bg-gradient-animated">
      <motion.div initial={{opacity:0,y:30}} animate={{opacity:1,y:0}} transition={{duration:1}} style={{textAlign:"center", zIndex:1}}>
        <div style={{width:120, height:2, background:"#19A3FC", margin:"0 auto 32px"}} />
        <p style={{color:"#19A3FC", fontSize:13, letterSpacing:"0.3em", textTransform:"uppercase", marginBottom:16, fontWeight:500}}>Digital Products</p>
        <h1 style={{fontSize:64, fontWeight:800, color:"#fff", marginBottom:16, letterSpacing:"-0.02em"}}>Prior Art Review</h1>
        <p style={{fontSize:18, color:"#9CA3AF", maxWidth:600, margin:"0 auto 32px"}}>How 4 existing projects inform the Wave 1 operating model</p>
        <div style={{display:"flex", gap:48, justifyContent:"center"}}>
          {[{v:"4",l:"Projects"},{v:"10",l:"Deliverables"},{v:"375+",l:"Slides"},{v:"14",l:"Weeks"}].map(s=>(
            <div key={s.l} style={{textAlign:"center"}}>
              <div style={{fontSize:28, fontWeight:700, color:"#fff"}}>{s.v}</div>
              <div style={{fontSize:11, color:"#6B7280", textTransform:"uppercase", letterSpacing:"0.05em"}}>{s.l}</div>
            </div>
          ))}
        </div>
      </motion.div>
      <motion.div initial={{opacity:0}} animate={{opacity:1}} transition={{delay:2}} style={{position:"absolute", bottom:32, color:"#6B7280", fontSize:12, textTransform:"uppercase", letterSpacing:"0.1em"}}>
        Scroll to explore ↓
      </motion.div>
    </section>
  );
}

function ProjectCards() {
  const ref = useRef(null);
  const inView = useInView(ref, {once:true, margin:"-100px"});
  return (
    <section ref={ref} style={{minHeight:"100vh", display:"flex", flexDirection:"column", justifyContent:"center", padding:"80px 48px"}}>
      <motion.div initial={{opacity:0,y:20}} animate={inView?{opacity:1,y:0}:{}} style={{textAlign:"center", marginBottom:48}}>
        <h2 style={{fontSize:32, fontWeight:700, color:"#fff", marginBottom:12}}>The Four Projects</h2>
        <p style={{color:"#9CA3AF", maxWidth:600, margin:"0 auto"}}>Each area solved different parts of the same problem — understanding what was built reveals what's reusable for Digital Products.</p>
      </motion.div>
      <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:24, maxWidth:1100, margin:"0 auto", width:"100%"}}>
        {PROJECTS.map((p,i)=>(
          <motion.div key={p.id} initial={{opacity:0,y:40}} animate={inView?{opacity:1,y:0}:{}} transition={{delay:0.2+i*0.15}} className="glass" style={{padding:24, borderColor:rgba(p.color,0.12)}}>
            <div style={{height:4, width:64, borderRadius:4, background:p.color, marginBottom:16}} />
            <div style={{display:"flex", alignItems:"center", gap:12, marginBottom:8}}>
              <h3 style={{fontSize:20, fontWeight:700, color:"#fff"}}>{p.name}</h3>
              <span style={{fontSize:11, padding:"2px 10px", borderRadius:20, background:rgba(p.color,0.12), color:p.color, fontWeight:500}}>{p.headline}</span>
            </div>
            <p style={{fontSize:14, color:"#9CA3AF", lineHeight:1.6, marginBottom:16}}>{p.subtitle}</p>
            <div>
              <div style={{display:"flex", justifyContent:"space-between", fontSize:12, color:"#6B7280", marginBottom:4}}>
                <span>Maturity</span><span>{Math.round(p.maturity*100)}%</span>
              </div>
              <div style={{height:6, background:"rgba(255,255,255,0.05)", borderRadius:4, overflow:"hidden"}}>
                <motion.div initial={{width:0}} animate={inView?{width:`${p.maturity*100}%`}:{}} transition={{duration:1.2, delay:0.3+i*0.15}} style={{height:"100%", borderRadius:4, background:p.color}} />
              </div>
            </div>
          </motion.div>
        ))}
      </div>
    </section>
  );
}

function FlowSection() {
  const ref = useRef(null);
  const inView = useInView(ref, {once:true, margin:"-50px"});
  const [selected, setSelected] = useState<string|null>(null);
  const tokenRefs = useRef<Record<string, SVGCircleElement|null>>({});
  const glowRefs = useRef<Record<string, SVGCircleElement|null>>({});

  const runTrack = useCallback((pid: string) => {
    const flow = FLOWS.find(f=>f.pid===pid)!;
    const proj = PROJECTS.find(p=>p.id===pid)!;
    const token = tokenRefs.current[pid];
    const glow = glowRefs.current[pid];
    if (!token || !glow) return;
    setSelected(pid);

    const n = flow.nodes.length;
    const w = 900, pad = 60;
    const spacing = (w - pad*2) / (n-1);
    const getX = (i:number) => pad + i * spacing;
    const y = 35;

    const tl = gsap.timeline();
    tl.set([token, glow], {opacity:1, attr:{cx:getX(0), cy:y}});
    for (let i=0; i<n-1; i++) {
      const toX = getX(i+1);
      const isGate = flow.gates.includes(i+1);
      if (flow.loop === i+1) {
        const midX = (getX(i) + toX) / 2;
        tl.to([token,glow], {duration:0.3, attr:{cx:midX, cy:y-20}, ease:"power2.out"});
        tl.to([token,glow], {duration:0.3, attr:{cx:getX(i), cy:y}, ease:"power2.in"});
        tl.to([token,glow], {duration:0.3, attr:{cx:toX, cy:y}, ease:"power2.inOut"});
      } else {
        tl.to([token,glow], {duration:0.5, attr:{cx:toX}, ease:"power2.inOut"});
      }
      if (isGate) tl.to({}, {duration:0.25});
    }
    tl.to([token,glow], {duration:0.3, opacity:0}, "+=0.4");
  }, []);

  const runAll = () => FLOWS.forEach(f => runTrack(f.pid));

  const selFlow = selected ? FLOWS.find(f=>f.pid===selected) : null;
  const selProj = selected ? PROJECTS.find(p=>p.id===selected) : null;

  return (
    <section ref={ref} style={{padding:"80px 48px"}}>
      <motion.div initial={{opacity:0,y:20}} animate={inView?{opacity:1,y:0}:{}} style={{textAlign:"center", marginBottom:40}}>
        <h2 style={{fontSize:32, fontWeight:700, color:"#fff", marginBottom:12}}>How a Demand Flows</h2>
        <p style={{color:"#9CA3AF", maxWidth:600, margin:"0 auto"}}>Watch each project's pipeline. Click a project for details on tools, roles, and prioritization.</p>
      </motion.div>

      {/* Controls */}
      <div style={{display:"flex", gap:8, justifyContent:"center", marginBottom:40, flexWrap:"wrap"}}>
        <button onClick={runAll} style={{display:"flex", alignItems:"center", gap:8, padding:"10px 20px", borderRadius:12, background:"rgba(25,163,252,0.12)", color:"#19A3FC", border:"1px solid rgba(25,163,252,0.25)", cursor:"pointer", fontSize:14, fontWeight:600}}>▶ Run All</button>
        {PROJECTS.map(p=>(
          <button key={p.id} onClick={()=>runTrack(p.id)} style={{padding:"10px 16px", borderRadius:12, background:rgba(p.color,0.1), color:p.color, border:`1px solid ${rgba(p.color,0.25)}`, cursor:"pointer", fontSize:13, fontWeight:500, outline: selected===p.id ? `2px solid ${p.color}` : "none", outlineOffset:2}}>{p.name}</button>
        ))}
      </div>

      <div style={{display:"flex", gap:24, maxWidth:1400, margin:"0 auto"}}>
        {/* Tracks */}
        <div style={{flex:1, display:"flex", flexDirection:"column", gap:12}}>
          {FLOWS.map((flow) => {
            const proj = PROJECTS.find(p=>p.id===flow.pid)!;
            const n = flow.nodes.length;
            const w = 900, pad = 60, h = 75;
            const spacing = (w - pad*2) / (n-1);
            const getX = (i:number) => pad + i * spacing;
            const y = 35;
            return (
              <div key={flow.pid} style={{display:"flex", alignItems:"center", gap:16}}>
                <div style={{width:120, textAlign:"right", flexShrink:0}}>
                  <span style={{fontSize:12, fontWeight:700, color:proj.color}}>{proj.name}</span>
                </div>
                <div style={{flex:1, borderRadius:12, padding:"8px 4px", background:`linear-gradient(135deg, ${rgba(proj.color,0.04)}, ${rgba(proj.color,0.02)})`, border:`1px solid ${rgba(proj.color,0.1)}`}}>
                  <svg viewBox={`0 0 ${w} ${h}`} style={{width:"100%", height:h, overflow:"visible"}}>
                    {/* Edges */}
                    {flow.nodes.slice(0,-1).map((_,i) => (
                      <line key={i} x1={getX(i)} y1={y} x2={getX(i+1)} y2={y} stroke={rgba(proj.color,0.3)} strokeWidth={2} strokeDasharray={flow.loop===i+1?"5 5":"none"} />
                    ))}
                    {/* Nodes */}
                    {flow.nodes.map((label,i) => {
                      const x = getX(i);
                      const isGate = flow.gates.includes(i);
                      return (
                        <g key={i}>
                          {isGate ? (
                            <rect x={x-8} y={y-8} width={16} height={16} rx={3} fill={rgba(proj.color,0.15)} stroke={proj.color} strokeWidth={1.5} transform={`rotate(45 ${x} ${y})`} />
                          ) : (
                            <circle cx={x} cy={y} r={8} fill={rgba(proj.color,0.1)} stroke={proj.color} strokeWidth={1.5} />
                          )}
                          <text x={x} y={y+24} textAnchor="middle" fill="#94A3B8" fontSize={10} fontFamily="Inter, system-ui">{label}</text>
                        </g>
                      );
                    })}
                    {/* Loop arc */}
                    {flow.loop !== undefined && (
                      <path d={`M ${getX(flow.loop)} ${y} Q ${(getX(flow.loop-1)+getX(flow.loop))/2} ${y-28} ${getX(flow.loop-1)} ${y}`} fill="none" stroke={rgba(proj.color,0.2)} strokeWidth={1.5} strokeDasharray="4 3" />
                    )}
                    {/* Token */}
                    <circle ref={el=>{glowRefs.current[flow.pid]=el}} cx={getX(0)} cy={y} r={14} fill={proj.color} opacity={0} style={{filter:"blur(10px)"}} />
                    <circle ref={el=>{tokenRefs.current[flow.pid]=el}} cx={getX(0)} cy={y} r={5} fill="#fff" stroke={proj.color} strokeWidth={2.5} opacity={0} className="token-trail" />
                  </svg>
                </div>
              </div>
            );
          })}
        </div>

        {/* Detail panel */}
        <div style={{width:340, flexShrink:0}}>
          {selFlow && selProj ? (
            <motion.div key={selFlow.pid} initial={{opacity:0,x:20}} animate={{opacity:1,x:0}} style={{borderRadius:12, padding:20, background:rgba(selProj.color,0.05), border:`1px solid ${rgba(selProj.color,0.15)}`}}>
              <h3 style={{color:"#fff", fontSize:14, fontWeight:700, marginBottom:4}}>{selProj.name}</h3>
              <div style={{height:2, width:40, borderRadius:2, background:selProj.color, marginBottom:12}} />
              {(["philosophy","tools","prioritization","roles"] as const).map(key=>(
                <div key={key} style={{marginBottom:12}}>
                  <h4 style={{fontSize:10, color:"#6B7280", textTransform:"uppercase", letterSpacing:"0.05em", fontWeight:600, marginBottom:4}}>{key}</h4>
                  <p style={{fontSize:11, color:"#9CA3AF", lineHeight:1.6}}>{selFlow.details[key]}</p>
                </div>
              ))}
            </motion.div>
          ) : (
            <div style={{borderRadius:12, padding:20, border:"1px solid rgba(255,255,255,0.05)", background:"rgba(255,255,255,0.02)", textAlign:"center"}}>
              <p style={{color:"#4B5563", fontSize:12}}>Click a project to see philosophy, tools, prioritization, and roles.</p>
            </div>
          )}
        </div>
      </div>
    </section>
  );
}

function RadarSection() {
  const ref = useRef(null);
  const inView = useInView(ref, {once:true, margin:"-100px"});
  const [sel, setSel] = useState(0);
  const d = DELIVERABLES[sel];
  const actionColor: Record<string,string> = { ADOPT:"#00B050", ADAPT:"#0070C0", "BUILD NEW":"#D46B08", COMBINE:"#7C3AED" };

  const data = DELIVERABLES.map(d=>({subject:d.name, ps:d.ps*100, coe:d.coe*100, infra:d.infra*100, dto:d.dto*100}));

  return (
    <section ref={ref} style={{padding:"80px 48px"}}>
      <motion.div initial={{opacity:0,y:20}} animate={inView?{opacity:1,y:0}:{}} style={{textAlign:"center", marginBottom:40}}>
        <h2 style={{fontSize:32, fontWeight:700, color:"#fff", marginBottom:12}}>What Each Project Brings</h2>
        <p style={{color:"#9CA3AF", maxWidth:600, margin:"0 auto"}}>Radar shows coverage across 10 deliverables. Click a deliverable for reuse recommendation.</p>
      </motion.div>
      <div style={{display:"flex", gap:32, maxWidth:1200, margin:"0 auto"}}>
        <div className="glass" style={{flex:1, padding:24}}>
          <ResponsiveContainer width="100%" height={450}>
            <RadarChart data={data}>
              <PolarGrid stroke="rgba(255,255,255,0.06)" />
              <PolarAngleAxis dataKey="subject" tick={{fill:"#94A3B8", fontSize:11}} />
              {PROJECTS.map(p => (
                <Radar key={p.id} name={p.name} dataKey={p.id} stroke={p.color} fill={p.color} fillOpacity={0.08} strokeWidth={2} />
              ))}
              <Tooltip contentStyle={{background:"#1a1f2e", border:"1px solid rgba(255,255,255,0.1)", borderRadius:8, fontSize:12}} itemStyle={{color:"#e2e8f0"}} labelStyle={{color:"#fff", fontWeight:600}} formatter={(v:number)=>`${v}%`} />
            </RadarChart>
          </ResponsiveContainer>
          <div style={{display:"flex", gap:24, justifyContent:"center", marginTop:16}}>
            {PROJECTS.map(p=>(
              <div key={p.id} style={{display:"flex", alignItems:"center", gap:8}}>
                <div style={{width:12, height:12, borderRadius:6, background:p.color}} />
                <span style={{fontSize:11, color:"#9CA3AF"}}>{p.name}</span>
              </div>
            ))}
          </div>
        </div>
        <div style={{width:340, flexShrink:0}}>
          <div className="glass" style={{padding:20, marginBottom:16, borderColor:rgba(actionColor[d.action],0.2)}}>
            <h3 style={{color:"#fff", fontWeight:700, fontSize:15, marginBottom:8}}>{d.name}</h3>
            {PROJECTS.map(p=>{
              const val = d[p.id as keyof typeof d] as number;
              return (
                <div key={p.id} style={{marginBottom:6}}>
                  <div style={{display:"flex", justifyContent:"space-between", fontSize:11, marginBottom:2}}>
                    <span style={{color:"#9CA3AF"}}>{p.name}</span>
                    <span style={{color:p.color}}>{Math.round(val*100)}%</span>
                  </div>
                  <div style={{height:5, background:"rgba(255,255,255,0.05)", borderRadius:4, overflow:"hidden"}}>
                    <div style={{height:"100%", width:`${val*100}%`, borderRadius:4, background:p.color, transition:"width 0.5s"}} />
                  </div>
                </div>
              );
            })}
            <div style={{borderTop:"1px solid rgba(255,255,255,0.05)", paddingTop:12, marginTop:12}}>
              <span style={{fontSize:11, padding:"2px 10px", borderRadius:20, background:rgba(actionColor[d.action],0.12), color:actionColor[d.action], fontWeight:600}}>{d.action}</span>
              <span style={{fontSize:11, color:"#6B7280", marginLeft:8}}>for Digital Products</span>
              <p style={{fontSize:11, color:"#9CA3AF", marginTop:8, lineHeight:1.6}}>{d.reason}</p>
            </div>
          </div>
          <div style={{display:"flex", flexDirection:"column", gap:4}}>
            {DELIVERABLES.map((dd,i)=>(
              <button key={dd.name} onClick={()=>setSel(i)} style={{width:"100%", textAlign:"left", padding:"8px 12px", borderRadius:8, border:"none", cursor:"pointer", fontSize:12, fontWeight:sel===i?600:400, color:sel===i?"#fff":"#6B7280", background:sel===i?"rgba(255,255,255,0.08)":"transparent", transition:"all 0.15s"}}>
                {dd.name}
                <span style={{marginLeft:8, fontSize:10, padding:"1px 6px", borderRadius:10, background:rgba(actionColor[dd.action],0.1), color:actionColor[dd.action]}}>{dd.action}</span>
              </button>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}

// ── Operating Model Blueprints ──

const MODEL_BLUEPRINTS = [
  {
    id: "ps", name: "P&S / P&T", color: "#002060",
    title: "Flight Level Kanban Governance",
    layers: [
      { level: "Flight Level 3 — Portfolio", color: "#002060", desc: "Strategic direction. Portfolio Manager maintains and prioritizes the Portfolio Backlog. Evaluates and sequences EPICs according to business value and OKR alignment.",
        stages: ["Portfolio Backlog", "Discovery", "Ideation", "Solutioning", "Done"],
        swimlanes: ["High Risks & Blockers", "High Value", "Regulatory", "OKR's & Strategic Goals"],
        forums: "Big Room Planning (quarterly) · Portfolio Committee (monthly)" },
      { level: "Flight Level 2 — Program", color: "#0050A0", desc: "Tactical coordination. Programs break EPICs into features and plan releases. Epics and Solutions in flight.",
        stages: ["Program Backlog", "Work In Progress", "Discovery", "Shark Tank", "Done"],
        swimlanes: ["Program A", "Program B", "Program C"],
        forums: "Program Sync (bi-weekly) · Replenishment" },
      { level: "Flight Level 1 — Project", color: "#0070C0", desc: "Operational execution. Project teams break features into User Stories. Day-to-day delivery work.",
        stages: ["Backlog", "In Progress", "QA", "Deploy", "Done"],
        swimlanes: [],
        forums: "Daily Standup · Sprint Review · Retrospective" },
    ],
    lifecycle: { title: "E2E Management Process", phases: [
      { name: "Phase 0: Discovery", steps: ["0.1 Research", "0.2 Understanding", "0.3 Problem Statement", "0.4 Inspiration Gate"], gate: "GO/No-Go → Shark Tank" },
      { name: "Phase 1: Materialization", steps: ["1.2 Ideation", "1.3 Validation", "1.4 Solution Proposal"], gate: "Portfolio Committee Approval" },
      { name: "Phase 2: Creation", steps: ["2.1 Project Book", "2.2 Budget Release", "2.3 Resourcing"], gate: "Budget Release Gate" },
      { name: "Phase 3: Ongoing", steps: ["3.1 Launch", "3.2 PoC/MVP", "3.3 Build", "3.4 Testing", "3.5 Deploy"], gate: "Health Dashboard Review" },
    ]},
  },
  {
    id: "coe", name: "Digital CoE", color: "#0070C0",
    title: "Discovery-Delivery Dual Board with Feedback Loops",
    layers: [
      { level: "Portfolio Governance — 3 Tier Model", color: "#0070C0", desc: "Portfolio (Initiative Level) → Strategic. Program (Feature/Epic Level) → Tactical. Delivery (User Story Level) → Executional.",
        stages: ["Strategic Initiatives Funnel", "Lean Business Case", "PRD", "Quarterly Program Plan"],
        swimlanes: [],
        forums: "Portfolio Learn Cycle (quarterly: 30 min past, 30 min future)" },
      { level: "Discovery Board (Upstream)", color: "#0090E0", desc: "Structured discovery workflow. Each PO moves discoveries through phases by changing status. Measures Cycle Time per phase and Discovery Success Rate.",
        stages: ["Opportunity Identification", "Ideation Phase", "Technical Feasibility", "Design Phase", "Business Signoff", "Done"],
        swimlanes: [],
        forums: "Discovery Review · Feedback from Delivery informs next Discovery" },
      { level: "Delivery Board (Downstream)", color: "#19A3FC", desc: "Sprint-based delivery. 2-week time-boxed cycles. Sprint Report Canvas tracks baselines across cycles.",
        stages: ["Sprint Backlog", "In Progress", "QA", "Done / Deployed"],
        swimlanes: [],
        forums: "Sprint Planning · Daily · Sprint Review · Retrospective" },
    ],
    lifecycle: { title: "E2E Agile Framework", phases: [
      { name: "Discovery Upstream", steps: ["User Research", "Ideation", "Prototyping", "Validation"], gate: "PRD + Lean Business Case" },
      { name: "Feedback Loop", steps: ["Delivery insights → Discovery", "Real-world data → Refinement"], gate: "Continuous (not gated)" },
      { name: "Delivery Downstream", steps: ["Backlog Refinement", "Development", "Testing", "Deployment"], gate: "Sprint Review + DoD" },
      { name: "Portfolio Review", steps: ["Learn Cycle", "Strategic Realignment", "Next Quarter Planning"], gate: "Quarterly Portfolio Learn Cycle" },
    ]},
  },
  {
    id: "infra", name: "Infrastructure", color: "#009688",
    title: "Team Topologies + Phased Deployment",
    layers: [
      { level: "Team Topologies Framework", color: "#009688", desc: "4 team types with 3 interaction modes. Selection depends on volume/frequency of demand, reuse of deliverable, and current capabilities.",
        stages: ["Stream-Aligned Team", "Enabling Team", "Platform Team", "Complicated Subsystem Team"],
        swimlanes: ["Collaboration", "X-as-a-Service", "Facilitating"],
        forums: "Interaction mode defines how teams relate structurally" },
      { level: "Deployment Methodology", color: "#00B8A0", desc: "Rollout-iterative strategy. Learn from each phase before expanding.",
        stages: ["Assessment (ODAS)", "Preparation & Onboarding", "Kanban Model Rollout", "Portfolio Roadmap Design", "Stabilization"],
        swimlanes: [],
        forums: "Sapiens Learning Cycles · Coaching Sessions" },
    ],
    lifecycle: { title: "Pilot → Expansion → Optimization", phases: [
      { name: "Pilot (1 team)", steps: ["ODAS Assessment", "Team Blueprint", "Kanban System Design", "Close Coaching"], gate: "Flow Metrics Baseline" },
      { name: "Learn & Measure", steps: ["Lead Time", "Cycle Time", "Throughput", "Team Health"], gate: "Results Report" },
      { name: "Expansion (7+ teams)", steps: ["Sapiens Learning Cycle", "Team Coaching", "Knowledge Transfer"], gate: "Expansion Readiness" },
      { name: "Optimization", steps: ["Portfolio Roadmap", "Investment Model", "Continuous Improvement"], gate: "Maturity Assessment" },
    ]},
  },
  {
    id: "dto", name: "Renewables DTO", color: "#D46B08",
    title: "Lean Kanban 4-Pillar Model",
    layers: [
      { level: "Level 1 — Epic Kanban (Portfolio)", color: "#D46B08", desc: "Connects strategy to execution. Swimlanes by business area. Ideas evaluated by DTO Project Committee.",
        stages: ["Backlog", "PMO", "Project Prep", "Ready", "Design & Build", "Done"],
        swimlanes: ["Growth (Product)", "Power Trading", "O&M", "Canceled"],
        forums: "DTO Project Committee · Replenishment Meeting" },
      { level: "Level 2 — Project Kanban (Execution)", color: "#E88B30", desc: "Pull-based execution. WIP limits enforced. QA pulls from Dev immediately. Definition of Done (DoD) required.",
        stages: ["Backlog", "In Progress", "QA / Peer Review", "UAT", "Done"],
        swimlanes: ["Capacity: Build 60%", "BAU 25%", "Upgrades 15%"],
        forums: "Daily Flow Review · Retrospective / Kaizen" },
    ],
    lifecycle: { title: "4 Pillars: Intake → Prioritize → Deliver → Govern", phases: [
      { name: "Demand Intake", steps: ["Jira Ticket (standardized)", "Classify: Class of Service", "Triage: accept/clarify/reject", "DoR Compliance"], gate: "Definition of Ready met" },
      { name: "Prioritization", steps: ["Continuous ordering (not batch)", "BV × Urgency × Capacity scoring", "Pre-Estimation (lightweight)"], gate: "Prioritized Backlog" },
      { name: "Delivery", steps: ["Pull from backlog (no push)", "WIP limits enforced", "QA pulls from Dev", "DoD validation"], gate: "Definition of Done met" },
      { name: "Governance", steps: ["Flow Metrics: LT, CT, Throughput, Aging", "Daily Flow Review", "Kaizen / Continuous Improvement"], gate: "Metrics reviewed regularly" },
    ]},
  },
];

function ModelBlueprints() {
  const ref = useRef(null);
  const inView = useInView(ref, {once:true, margin:"-50px"});
  const [activeModel, setActiveModel] = useState(0);
  const m = MODEL_BLUEPRINTS[activeModel];
  const proj = PROJECTS.find(p=>p.id===m.id)!;

  return (
    <section ref={ref} style={{padding:"80px 48px"}}>
      <motion.div initial={{opacity:0,y:20}} animate={inView?{opacity:1,y:0}:{}} style={{textAlign:"center", marginBottom:40}}>
        <h2 style={{fontSize:32, fontWeight:700, color:"#fff", marginBottom:12}}>Operating Model Blueprints</h2>
        <p style={{color:"#9CA3AF", maxWidth:700, margin:"0 auto"}}>The actual governance structure, board design, and lifecycle phases as defined in each project's materials.</p>
      </motion.div>

      {/* Tabs */}
      <div style={{display:"flex", gap:8, justifyContent:"center", marginBottom:32}}>
        {MODEL_BLUEPRINTS.map((bp,i)=>(
          <button key={bp.id} onClick={()=>setActiveModel(i)}
            style={{padding:"10px 20px", borderRadius:12, cursor:"pointer", fontSize:13, fontWeight: activeModel===i ? 700 : 500,
              background: activeModel===i ? rgba(bp.color,0.15) : "rgba(255,255,255,0.03)",
              color: activeModel===i ? bp.color : "#6B7280",
              border: `1px solid ${activeModel===i ? rgba(bp.color,0.3) : "rgba(255,255,255,0.06)"}`,
              transition:"all 0.2s"
            }}>
            {bp.name}
          </button>
        ))}
      </div>

      <motion.div key={m.id} initial={{opacity:0, y:10}} animate={{opacity:1, y:0}} style={{maxWidth:1200, margin:"0 auto"}}>
        {/* Title */}
        <div style={{textAlign:"center", marginBottom:32}}>
          <h3 style={{fontSize:22, fontWeight:700, color:"#fff", marginBottom:4}}>{m.title}</h3>
          <div style={{height:3, width:60, borderRadius:2, background:proj.color, margin:"0 auto"}} />
        </div>

        {/* Board Layers */}
        <div style={{display:"flex", flexDirection:"column", gap:16, marginBottom:40}}>
          {m.layers.map((layer, li) => (
            <motion.div key={li} initial={{opacity:0, x:-20}} animate={{opacity:1, x:0}} transition={{delay:li*0.15}}
              style={{borderRadius:12, overflow:"hidden", border:`1px solid ${rgba(layer.color, 0.2)}`}}>
              {/* Layer header */}
              <div style={{padding:"12px 20px", background:rgba(layer.color, 0.1), display:"flex", justifyContent:"space-between", alignItems:"center"}}>
                <div>
                  <h4 style={{fontSize:13, fontWeight:700, color:"#fff"}}>{layer.level}</h4>
                  <p style={{fontSize:11, color:"#9CA3AF", marginTop:2}}>{layer.desc}</p>
                </div>
                {layer.forums && (
                  <span style={{fontSize:10, color:layer.color, background:rgba(layer.color,0.1), padding:"4px 12px", borderRadius:20, whiteSpace:"nowrap"}}>{layer.forums}</span>
                )}
              </div>
              {/* Workflow stages */}
              <div style={{padding:"16px 20px", background:"rgba(255,255,255,0.02)"}}>
                <div style={{display:"flex", alignItems:"center", gap:4}}>
                  {layer.stages.map((stage, si) => (
                    <div key={si} style={{display:"flex", alignItems:"center", gap:4, flex:1}}>
                      <div style={{flex:1, padding:"8px 12px", borderRadius:8, textAlign:"center", fontSize:11, fontWeight:500, color:"#fff",
                        background:rgba(layer.color, 0.08), border:`1px solid ${rgba(layer.color, 0.15)}`}}>
                        {stage}
                      </div>
                      {si < layer.stages.length - 1 && <span style={{color:rgba(layer.color,0.4), fontSize:16}}>→</span>}
                    </div>
                  ))}
                </div>
                {/* Swimlanes */}
                {layer.swimlanes.length > 0 && (
                  <div style={{marginTop:12, display:"flex", gap:8, flexWrap:"wrap"}}>
                    <span style={{fontSize:10, color:"#6B7280", fontWeight:600}}>Swimlanes:</span>
                    {layer.swimlanes.map((sw,i)=>(
                      <span key={i} style={{fontSize:10, padding:"2px 8px", borderRadius:10, background:"rgba(255,255,255,0.04)", color:"#9CA3AF", border:"1px solid rgba(255,255,255,0.06)"}}>{sw}</span>
                    ))}
                  </div>
                )}
              </div>
            </motion.div>
          ))}
        </div>

        {/* Lifecycle Phases */}
        <div>
          <h4 style={{fontSize:14, fontWeight:700, color:"#fff", marginBottom:16, textAlign:"center"}}>{m.lifecycle.title}</h4>
          <div style={{display:"grid", gridTemplateColumns:`repeat(${m.lifecycle.phases.length}, 1fr)`, gap:12}}>
            {m.lifecycle.phases.map((phase, pi) => (
              <motion.div key={pi} initial={{opacity:0, y:20}} animate={{opacity:1, y:0}} transition={{delay:0.3+pi*0.1}}
                style={{borderRadius:12, overflow:"hidden", border:`1px solid ${rgba(proj.color, 0.15)}`, background:"rgba(255,255,255,0.02)"}}>
                <div style={{padding:"10px 14px", background:rgba(proj.color, 0.08)}}>
                  <h5 style={{fontSize:12, fontWeight:700, color:proj.color}}>{phase.name}</h5>
                </div>
                <div style={{padding:"12px 14px"}}>
                  {phase.steps.map((step,si) => (
                    <div key={si} style={{fontSize:11, color:"#9CA3AF", padding:"3px 0", display:"flex", alignItems:"center", gap:6}}>
                      <span style={{width:4, height:4, borderRadius:2, background:proj.color, flexShrink:0}} />
                      {step}
                    </div>
                  ))}
                  <div style={{marginTop:8, paddingTop:8, borderTop:"1px solid rgba(255,255,255,0.05)"}}>
                    <span style={{fontSize:10, color:"#EAB308", fontWeight:600}}>Gate: </span>
                    <span style={{fontSize:10, color:"#D1D5DB"}}>{phase.gate}</span>
                  </div>
                </div>
              </motion.div>
            ))}
          </div>
        </div>
      </motion.div>
    </section>
  );
}

function Closing() {
  const ref = useRef(null);
  const inView = useInView(ref, {once:true});
  return (
    <section ref={ref} style={{minHeight:"100vh", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center"}} className="bg-gradient-animated">
      <motion.div initial={{opacity:0}} animate={inView?{opacity:1}:{}} transition={{duration:1}} style={{textAlign:"center"}}>
        <div style={{width:160, height:2, background:"#19A3FC", margin:"0 auto 40px"}} />
        <h2 style={{fontSize:48, fontWeight:800, color:"#fff", maxWidth:800, lineHeight:1.2, marginBottom:24}}>
          4 projects. 375+ slides of proven content.<br/><span style={{color:"#19A3FC"}}>14 weeks</span> to build the Digital Products operating model.
        </h2>
        <p style={{color:"#6B7280", fontSize:18, marginTop:32}}>NTT DATA × Avangrid — Wave 1: Review & Baseline</p>
      </motion.div>
    </section>
  );
}

// ── App ──
export default function App() {
  return (
    <div>
      <Hero />
      <ProjectCards />
      <FlowSection />
      <ModelBlueprints />
      <RadarSection />
      <Closing />
    </div>
  );
}
