import { useState, useEffect, useRef } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import { Radar, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, ResponsiveContainer, Legend } from 'recharts'
import './App.css'

// ═══ Avangrid Colors ═══
const A = { blue:'#00529B', green:'#78BE20', orange:'#F7941D', red:'#e85555', purple:'#9B59B6',
  text:'#e0e8f0', sub:'#8899b0', dim:'#3d5070', bg:'#02111e' }

type View = 'home'|'explore'|'context'|'assess'|'structures'

// ═══════════════════════════════════════
// HOME
// ═══════════════════════════════════════
// Floating particles component
function Particles() {
  const particles = Array.from({length:20},(_,i)=>({
    id:i, x:Math.random()*100, y:Math.random()*100, size:3+Math.random()*5,
    dur:15+Math.random()*20, delay:Math.random()*10,
    color: i%3===0?A.green:i%3===1?A.blue:A.orange
  }))
  return (
    <div style={{position:'absolute',inset:0,overflow:'hidden',pointerEvents:'none',zIndex:0}}>
      {particles.map(p=>(
        <motion.div key={p.id}
          initial={{x:`${p.x}vw`,y:`${p.y}vh`,opacity:0}}
          animate={{x:[`${p.x}vw`,`${(p.x+15)%100}vw`,`${p.x}vw`],y:[`${p.y}vh`,`${(p.y-20+100)%100}vh`,`${p.y}vh`],opacity:[0,.4,0]}}
          transition={{duration:p.dur,delay:p.delay,repeat:Infinity,ease:'easeInOut'}}
          style={{position:'absolute',width:p.size,height:p.size,borderRadius:'50%',background:p.color}} />
      ))}
      {/* Animated gradient orbs */}
      <motion.div animate={{scale:[1,1.2,1],opacity:[.12,.2,.12]}} transition={{duration:8,repeat:Infinity,ease:'easeInOut'}}
        style={{position:'absolute',top:'-10%',right:'-5%',width:500,height:500,borderRadius:'50%',background:`radial-gradient(circle,${A.blue}40,transparent 70%)`,filter:'blur(60px)'}} />
      <motion.div animate={{scale:[1,1.15,1],opacity:[.08,.16,.08]}} transition={{duration:10,delay:2,repeat:Infinity,ease:'easeInOut'}}
        style={{position:'absolute',bottom:'0%',left:'-5%',width:400,height:400,borderRadius:'50%',background:`radial-gradient(circle,${A.green}40,transparent 70%)`,filter:'blur(60px)'}} />
      <motion.div animate={{scale:[1,1.1,1],opacity:[.06,.12,.06]}} transition={{duration:12,delay:4,repeat:Infinity,ease:'easeInOut'}}
        style={{position:'absolute',top:'40%',right:'15%',width:300,height:300,borderRadius:'50%',background:`radial-gradient(circle,${A.orange}30,transparent 70%)`,filter:'blur(60px)'}} />
    </div>
  )
}

function Home({go}:{go:(v:View)=>void}) {
  return (
    <div className="scene" style={{display:'flex',alignItems:'center',justifyContent:'center',padding:'40px 32px',position:'relative',overflow:'hidden'}}>
      <Particles />
      <div className="aurora" />
      <div className="aurora" style={{animationDelay:'4s',top:'-30%',opacity:.7}} />
      <div style={{maxWidth:900,width:'100%',position:'relative',zIndex:1}}>
        <motion.div initial={{opacity:0,y:24}} animate={{opacity:1,y:0}} transition={{duration:.8}}>
          <p style={{fontSize:12,color:A.sub,letterSpacing:4,textTransform:'uppercase',marginBottom:16}}>Inspiring Session  •  Avangrid × NTT DATA</p>
          <h1 style={{fontSize:64,fontWeight:900,lineHeight:1,letterSpacing:-3}}>
            From <motion.span className="text-ar" animate={{backgroundPosition:['0%','100%','0%']}} transition={{duration:6,repeat:Infinity}} style={{backgroundSize:'200% 200%'}}>Project</motion.span><br/>
            to <motion.span className="text-ag" animate={{backgroundPosition:['0%','100%','0%']}} transition={{duration:6,delay:1,repeat:Infinity}} style={{backgroundSize:'200% 200%'}}>Product</motion.span>
          </h1>
          <p style={{fontSize:20,color:A.sub,marginTop:20,maxWidth:540,lineHeight:1.7}}>What changes when you shift from project-based delivery to product-centric operations — and why it matters for Digital Products.</p>
        </motion.div>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16,marginTop:56}}>
          {([
            {v:'explore' as View,n:'01',t:'Explore the Shift',s:'5 key dimensions with interactive comparisons',c:A.green,i:'⚡'},
            {v:'structures' as View,n:'02',t:'How Others Do It',s:'Gartner: Aon, Ford, Globo + Avangrid proposal',c:A.blue,i:'🏗️'},
            {v:'context' as View,n:'03',t:'Your Reality Today',s:'12 questions about how Digital Products works now',c:A.orange,i:'🏢'},
            {v:'assess' as View,n:'04',t:'Maturity Radar',s:'Rate where you are and where you want to be',c:A.purple,i:'📊'},
          ]).map((x,i)=>(
            <motion.div key={x.v} initial={{opacity:0,y:24}} animate={{opacity:1,y:0}} transition={{delay:.4+i*.12}}
              className="card card-glow" onClick={()=>go(x.v)}
              whileHover={{scale:1.02,borderColor:`${x.c}30`,boxShadow:`0 8px 32px ${x.c}10`}}
              style={{cursor:'pointer',display:'flex',gap:20,alignItems:'center',transition:'box-shadow .3s'}}>
              <motion.div whileHover={{rotate:5,scale:1.1}} style={{width:56,height:56,borderRadius:14,background:`${x.c}10`,border:`1px solid ${x.c}20`,display:'flex',alignItems:'center',justifyContent:'center',fontSize:26,flexShrink:0}}>{x.i}</motion.div>
              <div style={{flex:1}}><p className="mono" style={{fontSize:11,color:x.c,marginBottom:4}}>{x.n}</p><p style={{fontSize:18,fontWeight:700,color:A.text,margin:0}}>{x.t}</p><p style={{fontSize:14,color:A.sub,margin:'4px 0 0'}}>{x.s}</p></div>
              <motion.span whileHover={{x:4}} style={{fontSize:20,color:A.dim}}>→</motion.span>
            </motion.div>
          ))}
        </div>
        <motion.p initial={{opacity:0}} animate={{opacity:1}} transition={{delay:1.2}} style={{textAlign:'center',marginTop:48,fontSize:13,color:A.dim}}>Prepared for Heather Lajoie, Director of IT Product Management</motion.p>
      </div>
    </div>
  )
}

function Back({go}:{go:(v:View)=>void}) {
  return <button onClick={()=>go('home')} style={{position:'fixed',top:20,left:20,zIndex:100,padding:'8px 18px',borderRadius:10,background:'rgba(4,22,44,.9)',border:'1px solid rgba(255,255,255,.08)',color:A.sub,fontSize:14}}>← Back</button>
}

// ═══════════════════════════════════════
// EXPLORE — 5 Dimensions
// ═══════════════════════════════════════
const DIMS = [
  { id:'success', title:'How We Define Success', icon:'🎯', color:A.green,
    project:{title:'Deliver scope on time',points:['Success = milestones completed on schedule','Measured by: on time, on budget, on scope','Key question: "Are we on track?"','Risk: building the wrong thing perfectly on time'],
      phases:[{n:'Requirements',d:4},{n:'Design',d:3},{n:'Build',d:6},{n:'Testing',d:2},{n:'Deploy',d:1}]},
    product:{title:'Deliver outcomes that matter',points:['Success = measurable business impact','Measured by: adoption, retention, NPS, revenue','Key question: "Are users better off?"','Advantage: course-correct every 2 weeks based on real data'],
      phases:[{n:'Discover',d:1},{n:'Build MVP',d:2},{n:'User Test',d:1},{n:'Build v1',d:3},{n:'Measure',d:1},{n:'Iterate',d:2}]},
    metrics:[{l:'Time to first value',p:'16 weeks',d:'4 weeks'},{l:'User feedback rounds',p:'0',d:'3'},{l:'Risk of building wrong thing',p:'HIGH',d:'LOW'}],
    insight:'The fundamental shift: from "did we deliver what was asked?" to "did we create real value?"'
  },
  { id:'teams', title:'Team Structure', icon:'👥', color:A.blue,
    project:{title:'Temporary teams — assemble and disband',points:['Teams assembled per project, disbanded after','People rotate constantly — no continuity','Knowledge and context lost at every transition','Team = cost center, minimized when possible'],
      phases:[{n:'Assemble',d:2},{n:'Project A',d:5},{n:'Disband',d:1},{n:'Reassemble',d:2},{n:'Project B',d:5},{n:'Disband',d:1}]},
    product:{title:'Persistent teams that grow expertise',points:['Cross-functional team owns a product area permanently','Deep domain expertise grows sprint over sprint','Team improves its own processes continuously','Team = value center, invested in long-term'],
      phases:[{n:'Sprint 1',d:2},{n:'Sprint 2',d:2},{n:'Sprint 3',d:2},{n:'Sprint 4',d:2},{n:'Sprint 5',d:2},{n:'Sprint 6',d:2}]},
    metrics:[{l:'Knowledge retention',p:'Lost every 6mo',d:'Grows forever'},{l:'Ramp-up time',p:'3-4 weeks',d:'0 — permanent'},{l:'Team ownership',p:'Ends at close',d:'Permanent'}],
    insight:'Gartner: Aon (100%), Ford (100%), Globo (70%) all run persistent product teams of 8-15.'
  },
  { id:'discovery', title:'Discovery & Validation', icon:'🔍', color:A.purple,
    project:{title:'Validate at the very end',points:['Requirements gathered upfront, then frozen','BA writes spec → IT builds exactly that','Users see product only at UAT — too late','Assumptions untested until launch day'],
      phases:[{n:'Write BRD',d:3},{n:'Approve',d:1},{n:'Build',d:8},{n:'UAT',d:2},{n:'Launch 🤞',d:1}]},
    product:{title:'Validate continuously',points:['Discovery runs in parallel with delivery — always','Product trio tests assumptions weekly','Prototypes reduce risk BEFORE building','Confidence grows incrementally'],
      phases:[{n:'Interview',d:1},{n:'Prototype',d:1},{n:'Test',d:1},{n:'Build',d:2},{n:'Ship',d:1},{n:'Learn',d:1},{n:'Iterate',d:2}]},
    metrics:[{l:'When users first see it',p:'Month 14',d:'Week 2'},{l:'Assumptions tested before build',p:'0',d:'Weekly'},{l:'Cost of being wrong',p:'Entire budget',d:'1 sprint'}],
    insight:'Teresa Torres: "The biggest risk isn\'t building wrong — it\'s spending MONTHS building wrong."'
  },
  { id:'delivery', title:'Release Cadence', icon:'🚀', color:A.orange,
    project:{title:'Big bang at the end',points:['One massive release after months of building','Huge integration effort before deployment','If it fails, everything fails — maximum risk','Users wait months for any improvement'],
      phases:[{n:'Build',d:4},{n:'Build',d:4},{n:'Build',d:4},{n:'Integrate',d:2},{n:'Deploy 💥',d:1}]},
    product:{title:'Continuous small releases',points:['Small releases every 1-2 weeks','Each release is small, safe, and reversible','If something breaks, rollback is trivial','Users see improvements every sprint'],
      phases:[{n:'v1',d:1},{n:'v2',d:1},{n:'v3',d:1},{n:'v4',d:1},{n:'v5',d:1},{n:'v6',d:1},{n:'v7',d:1}]},
    metrics:[{l:'Releases per quarter',p:'1',d:'6-12'},{l:'Rollback risk',p:'Catastrophic',d:'Trivial'},{l:'User feedback loop',p:'6+ months',d:'2 weeks'}],
    insight:'Amazon deploys every 11.6 seconds. Ford: 330+ product lines deploying independently.'
  },
  { id:'priority', title:'How We Prioritize', icon:'⚖️', color:A.green,
    project:{title:'Whoever shouts loudest wins',points:['HiPPO decides (Highest Paid Person\'s Opinion)','Everything is "urgent" and "critical"','First-come-first-served queue','No structured way to say "no"'],
      phases:[{n:'Request 1',d:3},{n:'Request 2',d:3},{n:'Request 3',d:3},{n:'Request 4',d:3},{n:'Request 5',d:3}]},
    product:{title:'Evidence-based decisions',points:['Scoring frameworks: RICE, WSJF, ICE','Value vs Effort matrix drives sequencing','Data and user research inform priorities','Saying "no" to low-value work is celebrated'],
      phases:[{n:'Quick Win',d:2},{n:'Validate',d:3},{n:'Strategic',d:4},{n:'Fill-in',d:1}]},
    metrics:[{l:'Basis for decisions',p:'Opinion',d:'Data + evidence'},{l:'Low-value work',p:'Gets done anyway',d:'Killed early'},{l:'Capacity wasted',p:'~40%',d:'<10%'}],
    insight:'The PM\'s most important skill is deciding what NOT to build.'
  },
]

// ═══════════════════════════════════════
// SIMULATION RENDERER — unique visual per theme
// ═══════════════════════════════════════


function SimRenderer({themeId, progress:p, mode='project'}:{themeId:string, progress:number, mode?:'project'|'product'}) {
  const isProj = mode === 'project'
  const col = isProj ? A.red : A.green

  // When progress is 0, show static content but no animation
  // (p is treated as 0 in all animations — content visible but frozen)

  // ── MINDSET: The fundamental shift ──
  if(themeId==='mindset') {
    if(isProj) {
      // Project mindset: linear, one-shot, handoff
      const step = Math.floor((p/100)*5)
      return (<div>
        <div style={{display:'flex',alignItems:'baseline',gap:10,marginBottom:20}}>
          <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Project Mindset</span>
        </div>
        <div style={{position:'relative',paddingLeft:32,borderLeft:`3px solid ${col}30`,marginLeft:8}}>
          {[
            {label:'Define the scope',sub:'Lock requirements early. Get sign-off. No changes after this.',icon:'📋'},
            {label:'Assemble a team',sub:'Pull people from different areas. They don\'t know each other yet.',icon:'👥'},
            {label:'Execute the plan',sub:'Build exactly what was specified. Track against the Gantt chart.',icon:'⚙️'},
            {label:'Hand off and disband',sub:'Ship it. Move people to the next project. Knowledge walks out the door.',icon:'📦'},
            {label:'Hope it works',sub:'Was it the right thing? We\'ll find out... eventually.',icon:'🤞'},
          ].map((s,i)=>{
            const active = step >= i
            return <div key={i} style={{marginBottom:28,opacity:active?1:.15,transition:'all .5s'}}>
              <div style={{position:'absolute',left:-9,width:15,height:15,borderRadius:8,background:active?col:A.dim,marginTop:4,boxShadow:step===i?`0 0 12px ${col}`:'none'}} />
              <p style={{fontSize:22,fontWeight:700,color:active?A.text:A.dim,margin:0}}>{s.icon} {s.label}</p>
              <p style={{fontSize:16,color:active?A.sub:A.dim,margin:'6px 0 0'}}>{s.sub}</p>
            </div>
          })}
        </div>
        {p>90 && <div style={{padding:'16px 24px',borderRadius:12,background:`${col}08`,border:`1px solid ${col}20`}}>
          <p style={{fontSize:17,fontWeight:600,color:col,margin:0}}>The project is "done" — but was it the right thing to build? Nobody knows until users complain.</p>
        </div>}
      </div>)
    } else {
      // Product mindset: continuous, iterative, owned
      const cycle = Math.floor((p/100)*4) // 4 cycles
      return (<div>
        <div style={{display:'flex',alignItems:'baseline',gap:10,marginBottom:20}}>
          <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Product Mindset</span>
        </div>
        {/* Cycles */}
        {[0,1,2,3].map(c => {
          const active = cycle >= c
          const current = cycle === c
          return <div key={c} style={{marginBottom:20,padding:20,borderRadius:14,
            background:active?`${col}08`:'rgba(255,255,255,.01)',
            border:`1px solid ${active?col+'20':'rgba(255,255,255,.03)'}`,
            opacity:active?1:.12,transition:'all .5s'}}>
            <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:10}}>
              <div style={{width:32,height:32,borderRadius:10,background:active?`${col}20`:'rgba(255,255,255,.03)',
                display:'flex',alignItems:'center',justifyContent:'center',fontSize:14,fontWeight:800,color:active?col:A.dim}}>
                {active?'✓':(c+1)}
              </div>
              <span style={{fontSize:18,fontWeight:700,color:active?A.text:A.dim}}>Cycle {c+1}</span>
              {current && <span style={{fontSize:13,color:col,fontWeight:600}}>Running...</span>}
            </div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr 1fr 1fr',gap:8}}>
              {['Discover','Build','Ship','Measure','Learn'].map((step,si) => (
                <div key={si} style={{padding:'8px 12px',borderRadius:8,textAlign:'center',fontSize:14,fontWeight:500,
                  background:active&&(current?p/100*4-c>.2*si:true)?`${col}15`:'rgba(255,255,255,.02)',
                  color:active?col:A.dim,transition:'all .4s'}}>
                  {step}
                </div>
              ))}
            </div>
            {active && c > 0 && <p style={{fontSize:13,color:col,marginTop:8}}>
              {c===1?'Users gave feedback → we adapted':c===2?'Adoption growing → data-driven decisions':c===3?'Product-market fit validated → ready to scale':''}
            </p>}
          </div>
        })}
        {p>90 && <div style={{padding:'16px 24px',borderRadius:12,background:`${col}08`,border:`1px solid ${col}20`}}>
          <p style={{fontSize:17,fontWeight:600,color:col,margin:0}}>The product is never "done" — it continuously improves based on real user feedback and business data.</p>
        </div>}
      </div>)
    }
  }

  // ── SUCCESS: Timeline with what gets measured at each step ──
  if(themeId==='success') {
    const steps = isProj
      ? [{label:'Requirements Gathering',dur:'4 weeks',measure:'Measuring: scope document completion %',measureColor:A.dim},
         {label:'Design Phase',dur:'3 weeks',measure:'Measuring: design sign-offs collected',measureColor:A.dim},
         {label:'Development',dur:'6 weeks',measure:'Measuring: burndown chart, code commits',measureColor:A.dim},
         {label:'QA & Testing',dur:'2 weeks',measure:'Measuring: bugs found and fixed',measureColor:A.dim},
         {label:'Deploy to Production',dur:'1 week',measure:'💥 First time users see ANYTHING — zero feedback until now',measureColor:A.red}]
      : [{label:'Discovery & User Research',dur:'1 week',measure:'Measuring: user pain points identified',measureColor:A.green},
         {label:'Build MVP Prototype',dur:'2 weeks',measure:'Measuring: prototype usability score',measureColor:A.green},
         {label:'User Testing & Feedback',dur:'1 week',measure:'✓ Adoption: 23% | Feedback: Round 1',measureColor:A.green},
         {label:'Build Version 1',dur:'3 weeks',measure:'✓ Adoption: 52% | NPS: +18 | Feedback: Round 2',measureColor:A.green},
         {label:'Measure & Analyze',dur:'1 week',measure:'✓ Adoption: 68% | NPS: +35 | Time-to-value: 5 days',measureColor:A.green},
         {label:'Iterate Based on Data',dur:'2 weeks',measure:'✓ Adoption: 78% | NPS: +48 | Time-to-value: 3 days',measureColor:A.green}]
    const totalW = isProj ? 16 : 10
    const week = (p/100)*totalW
    const stepCum:number[] = []; let c2=0; steps.forEach(s=>{stepCum.push(c2);c2+=parseInt(s.dur)})

    return (<div>
      <div style={{display:'flex',alignItems:'baseline',gap:10,marginBottom:20}}>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>{isProj?'Project Mode':'Product Mode'}</span>
        <span className="mono" style={{fontSize:40,fontWeight:800,color:col}}>Week {Math.floor(week)}</span>
        <span style={{fontSize:16,color:A.dim}}>/ {totalW}</span>
      </div>
      <div style={{position:'relative',paddingLeft:32,borderLeft:`3px solid ${col}30`,marginLeft:8}}>
        {steps.map((step,i)=>{
          const active = week >= stepCum[i]
          const current = active && week < stepCum[i] + parseInt(step.dur)
          return (
            <div key={i} style={{marginBottom:24,opacity:active?1:.2,transition:'all .5s'}}>
              <div style={{position:'absolute',left:-9,width:15,height:15,borderRadius:8,background:active?col:A.dim,border:`3px solid ${active?col:A.dim}`,marginTop:4,transition:'all .3s',boxShadow:current?`0 0 12px ${col}`:'none'}} />
              <p style={{fontSize:20,fontWeight:700,color:active?A.text:A.dim,margin:0}}>{step.label}</p>
              <p style={{fontSize:15,color:A.dim,margin:'4px 0 0'}}>{step.dur}</p>
              {active && <p style={{fontSize:14,color:step.measureColor,fontWeight:500,margin:'6px 0 0'}}>{step.measure}</p>}
            </div>
          )
        })}
      </div>
      {p>90 && <div style={{padding:'16px 24px',borderRadius:12,background:`${col}08`,border:`1px solid ${col}20`,marginTop:8}}>
        <p style={{fontSize:16,fontWeight:600,color:col,margin:0}}>
          {isProj ? 'All metrics green ✓ — but nobody used it until Week 16. Zero user feedback. Maximum risk.'
                  : 'Real outcomes from Week 3. Adoption 78%. NPS +48. 3 rounds of learning. We KNOW it works.'}
        </p>
      </div>}
    </div>)
  }

  // ── TEAMS ──
  if(themeId==='teams') {
    if(isProj) {
      // 9 phases: for each of 3 teams → assembling, working, disbanded
      // p goes 0-100, split into 3 teams of ~33% each
      // Each team: 0-8% assembling, 8-25% working, 25-33% disbanded
      const teams = ['Alpha','Beta','Gamma']
      const teamPhase = (teamIdx:number):{status:'waiting'|'assembling'|'working'|'disbanded',pct:number} => {
        const teamStart = teamIdx * 33
        const teamEnd = teamStart + 33
        const tp = p - teamStart // progress within this team's window
        if(p < teamStart) return {status:'waiting',pct:0}
        if(tp < 8) return {status:'assembling',pct:tp/8*100}
        if(tp < 25) return {status:'working',pct:(tp-8)/17*100}
        return {status:'disbanded',pct:100}
      }

      return (<div>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Project Mode — Temporary Teams</span>
        <div style={{marginTop:16}}>
          {teams.map((name,i)=>{
            const {status,pct} = teamPhase(i)
            const statusLabel = {
              waiting: '⏳ Not started yet',
              assembling: '🔄 Assembling team...',
              working: '⚡ Working on project...',
              disbanded: '❌ Disbanded — knowledge lost'
            }[status]
            const statusColor = {waiting:A.dim, assembling:A.orange, working:A.blue, disbanded:col}[status]

            return (
              <div key={i} style={{padding:20,borderRadius:14,marginBottom:12,
                background: status==='working'?`${A.blue}08` : status==='assembling'?`${A.orange}06` : 'rgba(255,255,255,.015)',
                border:`1px solid ${status==='working'?A.blue+'25':status==='assembling'?A.orange+'20':'rgba(255,255,255,.04)'}`,
                opacity: status==='disbanded'?.3 : status==='waiting'?.15 : 1,
                transition:'all .5s'}}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
                  <span style={{fontSize:18,fontWeight:700,color:status==='working'||status==='assembling'?A.text:A.dim}}>Team {name}</span>
                  <span style={{fontSize:14,color:statusColor,fontWeight:600}}>{statusLabel}</span>
                </div>
                {/* People assembling animation */}
                <div style={{display:'flex',gap:10}}>
                  {['Project Manager','Developer','Developer','QA Tester'].map((role,j)=>{
                    const roleVisible = status==='assembling' ? pct > j*25 : status!=='waiting'
                    return (
                      <div key={j} style={{padding:'6px 14px',borderRadius:8,fontSize:14,
                        background: status==='working'?'rgba(255,255,255,.08)' : roleVisible?'rgba(255,255,255,.04)':'rgba(255,255,255,.01)',
                        color: status==='working'?A.sub : roleVisible?A.dim:'rgba(255,255,255,.05)',
                        transition:'all .4s',
                        opacity: roleVisible?1:0,
                        transform: roleVisible?'translateY(0)':'translateY(8px)'}}>
                        {role}
                      </div>
                    )
                  })}
                </div>
                {/* Working progress bar */}
                {status==='working' && <div style={{marginTop:10}}>
                  <div style={{height:4,borderRadius:2,background:'rgba(255,255,255,.04)',overflow:'hidden'}}>
                    <div style={{height:'100%',borderRadius:2,background:A.blue,width:`${pct}%`,transition:'width .1s'}} />
                  </div>
                  <p style={{fontSize:12,color:A.blue,marginTop:4}}>Building... {Math.floor(pct)}% complete</p>
                </div>}
                {/* Disbanded message */}
                {status==='disbanded' && <p style={{fontSize:13,color:col,marginTop:10}}>Team members reassigned. All context and domain knowledge lost — next team starts from zero.</p>}
              </div>
            )
          })}
        </div>
        {p>95 && <div style={{padding:'16px 24px',borderRadius:12,background:`${col}08`,border:`1px solid ${col}20`}}>
          <p style={{fontSize:16,fontWeight:600,color:col,margin:0}}>3 teams formed and disbanded. Knowledge lost 3 times. Every new project starts from zero context.</p>
        </div>}
      </div>)
    } else {
      const started = p > 0
      return (<div>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Product Mode — Persistent Team</span>
        <div style={{padding:24,borderRadius:14,marginTop:16,background:started?`${col}06`:'rgba(255,255,255,.02)',border:`1px solid ${started?col+'15':'rgba(255,255,255,.04)'}`,transition:'all .5s'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
            <span style={{fontSize:20,fontWeight:800,color:started?A.text:A.dim}}>Digital Products Squad</span>
            {started && <span className="mono" style={{fontSize:14,color:col}}>Sprint {Math.floor(p/12)+1}</span>}
          </div>
          <div style={{display:'flex',gap:10,flexWrap:'wrap',marginBottom:20}}>
            {['Product Owner','UX Designer','Dev Lead','Developer','Developer','QA Engineer'].map((role,j)=>(
              <div key={j}
                style={{padding:'8px 18px',borderRadius:10,fontSize:15,fontWeight:600,
                  background: started && p>j*12 ? `${col}10` : 'rgba(255,255,255,.03)',
                  border:`1px solid ${started && p>j*12 ? col+'18' : 'rgba(255,255,255,.04)'}`,
                  color: started && p>j*12 ? col : A.dim,
                  opacity: started ? (p>j*12?1:.2) : .3,
                  transition:'all .5s'}}>
                {role}
              </div>
            ))}
          </div>
          {started && <div style={{marginBottom:12}}>
            <div style={{display:'flex',justifyContent:'space-between',marginBottom:6}}>
              <span style={{fontSize:15,color:A.sub}}>Domain Expertise</span>
              <span style={{fontSize:15,color:col,fontWeight:700}}>
                {p<20?'🌱 Forming':p<40?'📈 Growing':p<60?'🤝 Norming':p<80?'🚀 Performing':'🏆 High-Performing'}
              </span>
            </div>
            <div style={{height:10,borderRadius:5,background:'rgba(255,255,255,.04)',overflow:'hidden'}}>
              <div style={{height:'100%',borderRadius:5,background:`linear-gradient(90deg,${col}60,${col})`,width:`${p}%`,transition:'width .1s'}} />
            </div>
          </div>}
          <p style={{fontSize:14,color:started?A.dim:'rgba(255,255,255,.15)',transition:'color .5s'}}>{started?'Same team → same context → compounding knowledge every sprint':'Click ▶ Run to see how a persistent team builds expertise'}</p>
        </div>
        {p>90 && <div style={{padding:'16px 24px',borderRadius:12,background:`${col}08`,border:`1px solid ${col}20`,marginTop:12}}>
          <p style={{fontSize:16,fontWeight:600,color:col,margin:0}}>One team. Zero context switches. Expertise compounds. They know the product better than anyone.</p>
        </div>}
      </div>)
    }
  }

  // ── DISCOVERY ──
  if(themeId==='discovery') {
    const step = Math.floor((p/100)*(isProj?5:7))
    if(isProj) {
      return (<div>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Project Mode — Linear, One-Shot Path</span>
        <div style={{position:'relative',paddingLeft:32,borderLeft:`3px solid ${col}30`,marginLeft:8,marginTop:16}}>
          {[{label:'Write Business Requirements Document',sub:'3 weeks of documentation'},{label:'Approval Gate — Sign Off',sub:'Scope locked. No changes allowed.'},{label:'Build Everything',sub:'8 weeks — heads down, no user contact'},{label:'UAT — Users See It For The First Time',sub:'⚠️ "This isn\'t what we asked for..."'},{label:'Launch & Hope For The Best',sub:'🎰 All assumptions untested until NOW'}].map((s,i)=>{
            const active = step >= i
            return <div key={i} style={{marginBottom:24,opacity:active?1:.15,transition:'all .5s'}}>
              <div style={{position:'absolute',left:-9,width:15,height:15,borderRadius:8,background:active?col:A.dim,marginTop:4,boxShadow:step===i?`0 0 12px ${col}`:'none'}} />
              <p style={{fontSize:18,fontWeight:700,color:active?A.text:A.dim,margin:0}}>{s.label}</p>
              <p style={{fontSize:14,color:active?(i>=3?col:A.sub):A.dim,margin:'4px 0 0'}}>{s.sub}</p>
            </div>
          })}
        </div>
      </div>)
    } else {
      // Product Discovery: show 3 iterative CYCLES that ping-pong between Discovery and Delivery
      // Each cycle: Discovery phase → Delivery phase → Learning
      const cycles = [
        {disco:'Interview 5 users about their daily workflow',deliv:'Build clickable prototype',learn:'3 of 5 users confused by step 2 → redesign'},
        {disco:'Test redesigned flow with same users',deliv:'Build & ship MVP to pilot group',learn:'Pilot group adoption: 73% in week 1 → validate'},
        {disco:'Analyze usage data + 3 follow-up interviews',deliv:'Ship v1.1 with top 3 fixes',learn:'NPS jumped from +12 to +48 → ready to scale'},
      ]
      // Which cycle/phase are we in? p goes 0-100, 3 cycles = ~33% each, each has 3 phases
      const cycleIdx = Math.min(Math.floor(p / 33), 2)
      const withinCycle = (p - cycleIdx * 33) / 33 // 0-1 within current cycle
      const phase = withinCycle < .33 ? 'discovery' : withinCycle < .66 ? 'delivery' : 'learn'

      return (<div>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Product Mode — Iterative Discovery + Delivery Cycles</span>
        <div style={{marginTop:16}}>
          {cycles.map((cycle,i)=>{
            const isActive = i === cycleIdx && p > 0
            const isDone = i < cycleIdx
            const isFuture = i > cycleIdx || p === 0
            const activePhase = isActive ? phase : isDone ? 'done' : 'waiting'

            return (
              <div key={i} style={{marginBottom:16,padding:20,borderRadius:14,
                background: isDone?`${col}04` : isActive?`${col}08` : 'rgba(255,255,255,.01)',
                border:`1px solid ${isActive?col+'25':isDone?col+'10':'rgba(255,255,255,.03)'}`,
                opacity: isFuture&&p>0?.2:isFuture?.15:1,
                transition:'all .5s'}}>
                <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:12}}>
                  <div style={{width:32,height:32,borderRadius:10,
                    background:isDone?`${col}20`:isActive?`${col}15`:'rgba(255,255,255,.03)',
                    display:'flex',alignItems:'center',justifyContent:'center',
                    fontSize:14,fontWeight:800,color:isDone||isActive?col:A.dim}}>
                    {isDone?'✓':(i+1)}
                  </div>
                  <span style={{fontSize:18,fontWeight:700,color:isDone||isActive?A.text:A.dim}}>Cycle {i+1}</span>
                  {isDone && <span style={{fontSize:13,color:col,fontWeight:600}}>Complete</span>}
                  {isActive && <span style={{fontSize:13,color:A.orange,fontWeight:600}}>
                    {activePhase==='discovery'?'🔍 Discovering...':activePhase==='delivery'?'🚀 Delivering...':'💡 Learning...'}
                  </span>}
                </div>
                {/* Three phases within the cycle */}
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:10}}>
                  <div style={{padding:12,borderRadius:10,
                    background:activePhase==='discovery'?`${A.purple}12`:isDone?`${A.purple}06`:'rgba(255,255,255,.02)',
                    border:`1px solid ${activePhase==='discovery'?A.purple+'30':'rgba(255,255,255,.04)'}`,
                    transition:'all .4s'}}>
                    <p style={{fontSize:12,fontWeight:700,color:A.purple,margin:'0 0 6px'}}>🔍 Discovery</p>
                    <p style={{fontSize:13,color:isDone||activePhase!=='waiting'?A.sub:A.dim,margin:0,lineHeight:1.4}}>{cycle.disco}</p>
                  </div>
                  <div style={{padding:12,borderRadius:10,
                    background:activePhase==='delivery'?`${col}12`:isDone?`${col}06`:'rgba(255,255,255,.02)',
                    border:`1px solid ${activePhase==='delivery'?col+'30':'rgba(255,255,255,.04)'}`,
                    transition:'all .4s'}}>
                    <p style={{fontSize:12,fontWeight:700,color:col,margin:'0 0 6px'}}>🚀 Delivery</p>
                    <p style={{fontSize:13,color:isDone||(activePhase==='delivery'||activePhase==='learn')?A.sub:A.dim,margin:0,lineHeight:1.4}}>{cycle.deliv}</p>
                  </div>
                  <div style={{padding:12,borderRadius:10,
                    background:activePhase==='learn'?`${A.orange}12`:isDone?`${A.orange}06`:'rgba(255,255,255,.02)',
                    border:`1px solid ${activePhase==='learn'?A.orange+'30':'rgba(255,255,255,.04)'}`,
                    transition:'all .4s'}}>
                    <p style={{fontSize:12,fontWeight:700,color:A.orange,margin:'0 0 6px'}}>💡 Learning</p>
                    <p style={{fontSize:13,color:isDone||activePhase==='learn'?A.sub:A.dim,margin:0,lineHeight:1.4}}>{cycle.learn}</p>
                  </div>
                </div>
              </div>
            )
          })}
        </div>
        {p>95 && <div style={{padding:'16px 24px',borderRadius:12,background:`${col}08`,border:`1px solid ${col}20`}}>
          <p style={{fontSize:16,fontWeight:600,color:col,margin:0}}>3 cycles of Discovery → Delivery → Learning. Each cycle makes the product better. NPS: +12 → +48.</p>
        </div>}
      </div>)
    }
  }

  // ── DELIVERY ──
  if(themeId==='delivery') {
    const bars = 24
    if(isProj) {
      return (<div>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Project Mode — Months of Silence, Then Big Bang</span>
        <div style={{height:140,borderRadius:14,background:'rgba(255,255,255,.02)',border:'1px solid rgba(255,255,255,.04)',marginTop:16,display:'flex',alignItems:'flex-end',padding:'0 16px 16px',gap:5,overflow:'hidden'}}>
          {Array.from({length:bars},(_,i)=>{
            const isLast = i === bars-1
            const active = p > (i/bars)*100
            const h = isLast && p > 95 ? 110 : active ? 6 : 4
            return <motion.div key={i} animate={{height:h}} transition={{duration:.3}}
              style={{flex:1,borderRadius:'6px 6px 0 0',background:isLast&&p>95?col:`${col}${active?'25':'08'}`}} />
          })}
        </div>
        <div style={{display:'flex',justifyContent:'space-between',marginTop:10}}>
          <span style={{fontSize:14,color:A.dim}}>Month 1</span>
          <span style={{fontSize:14,color:p>95?col:p>0?A.sub:A.dim,fontWeight:p>95?700:500}}>{p>95?'💥 ONE massive, risky release':p>0?'⏳ No value delivered yet — users waiting...':'Click ▶ Run to start'}</span>
          <span style={{fontSize:14,color:A.dim}}>Month 16</span>
        </div>
      </div>)
    } else {
      return (<div>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Product Mode — Continuous Flow of Small Releases</span>
        <div style={{height:140,borderRadius:14,background:'rgba(255,255,255,.02)',border:'1px solid rgba(255,255,255,.04)',marginTop:16,display:'flex',alignItems:'flex-end',padding:'0 16px 16px',gap:5,overflow:'hidden'}}>
          {Array.from({length:bars},(_,i)=>{
            const active = p > (i/bars)*100
            const h = active ? 24 + Math.sin(i*0.8)*30 + 30 : 4
            return <motion.div key={i} animate={{height:h}} transition={{duration:.4,delay:i*.02}}
              style={{flex:1,borderRadius:'6px 6px 0 0',background:active?`linear-gradient(0deg,${col}50,${col})`:`${col}08`}} />
          })}
        </div>
        <div style={{display:'flex',justifyContent:'space-between',marginTop:10}}>
          <span style={{fontSize:14,color:A.dim}}>Week 1</span>
          <span style={{fontSize:14,color:col,fontWeight:700}}>Each bar = a small release delivering value</span>
          <span style={{fontSize:14,color:A.dim}}>Week 10</span>
        </div>
      </div>)
    }
  }

  // ── PRIORITIZATION ──
  if(themeId==='priority') {
    const items = [
      {name:'SAP Break-Fix',value:'High',effort:'Low',fifo:1,prodOrder:1,action:'✅ Do NOW',actionColor:A.green,score:92},
      {name:'BRS Onboarding Redesign',value:'High',effort:'High',fifo:2,prodOrder:2,action:'🔍 Validate first',actionColor:A.blue,score:78},
      {name:'Legacy Report Migration',value:'Low',effort:'High',fifo:3,prodOrder:5,action:'🚫 Kill it — waste',actionColor:A.red,score:8},
      {name:'ServiceNow Field Change',value:'Low',effort:'Low',fifo:4,prodOrder:4,action:'📋 Fill-in if capacity',actionColor:A.dim,score:22},
      {name:'Cross-OPCO Integration API',value:'High',effort:'High',fifo:5,prodOrder:3,action:'🔍 Validate first',actionColor:A.blue,score:85},
    ]

    if(isProj) {
      // Show items appearing in FIFO order with a progress indicator showing we're just doing them in order
      const activeIdx = Math.min(Math.floor((p/100)*5), 4)
      return (<div>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Project Mode — First Come, First Served</span>
        <p style={{fontSize:14,color:A.dim,margin:'8px 0 16px'}}>Requests processed in the order they arrived — no scoring, no trade-offs</p>
        <div>
          {items.sort((a,b)=>a.fifo-b.fifo).map((item,i)=>{
            const active = p > 0 && i === activeIdx
            const done = p > 0 && i < activeIdx
            const waiting = p === 0 || i > activeIdx
            return (
              <div key={i} style={{display:'flex',alignItems:'center',gap:16,padding:'16px 20px',marginBottom:8,borderRadius:12,
                background: active?`${col}10` : done?`${col}04` : 'rgba(255,255,255,.02)',
                border: active?`2px solid ${col}30` : `1px solid rgba(255,255,255,.05)`,
                transition:'all .4s'}}>
                <div style={{width:40,height:40,borderRadius:10,background:active?`${col}20`:done?`${col}10`:'rgba(255,255,255,.04)',
                  display:'flex',alignItems:'center',justifyContent:'center',transition:'all .4s'}}>
                  <span className="mono" style={{fontSize:18,fontWeight:800,color:done||active?col:A.dim}}>#{item.fifo}</span>
                </div>
                <div style={{flex:1}}>
                  <p style={{fontSize:16,fontWeight:600,color:done||active?A.text:A.sub,margin:0}}>{item.name}</p>
                  <div style={{display:'flex',gap:12,marginTop:4}}>
                    <span style={{fontSize:12,color:item.value==='High'?A.green:A.dim}}>Value: {item.value}</span>
                    <span style={{fontSize:12,color:item.effort==='High'?A.orange:A.dim}}>Effort: {item.effort}</span>
                  </div>
                </div>
                {active && <span style={{fontSize:13,color:col,fontWeight:600}}>⚡ Working...</span>}
                {done && <span style={{fontSize:13,color:A.dim}}>✓ Done</span>}
                {waiting && p>0 && <span style={{fontSize:13,color:A.dim}}>Queued</span>}
              </div>
            )
          })}
        </div>
        {p>90 && <div style={{padding:'16px 24px',borderRadius:12,marginTop:12,background:`${col}08`,border:`1px solid ${col}20`}}>
          <p style={{fontSize:16,fontWeight:600,color:col,margin:0}}>⚠️ "Legacy Migration" (LOW value, HIGH effort) was done at #3 — before the HIGHEST value "Cross-OPCO API" at #5. ~40% of capacity wasted on low-value work.</p>
        </div>}
      </div>)
    } else {
      // Product: items sorted by score, with clear value-based ranking
      const sorted = [...items].sort((a,b)=>b.score-a.score)
      const activeIdx = Math.min(Math.floor((p/100)*5), 4)
      return (<div>
        <span style={{fontSize:13,fontWeight:700,color:col,letterSpacing:1,textTransform:'uppercase'}}>Product Mode — Scored by Value vs Effort</span>
        <p style={{fontSize:14,color:A.dim,margin:'8px 0 16px'}}>Each request scored objectively — highest value/effort ratio goes first</p>
        <div>
          {sorted.map((item,i)=>{
            const active = p > 0 && i === activeIdx
            const done = p > 0 && i < activeIdx
            return (
              <div key={i} style={{display:'flex',alignItems:'center',gap:16,padding:'16px 20px',marginBottom:8,borderRadius:12,
                background: active?`${item.actionColor}10` : done?`${item.actionColor}04` : 'rgba(255,255,255,.02)',
                border: active?`2px solid ${item.actionColor}30` : `1px solid rgba(255,255,255,.05)`,
                transition:'all .4s'}}>
                <div className="mono" style={{width:48,height:40,borderRadius:10,
                  background:`${item.actionColor}15`,display:'flex',alignItems:'center',justifyContent:'center',
                  fontSize:18,fontWeight:800,color:item.actionColor}}>
                  {item.score}
                </div>
                <div style={{flex:1}}>
                  <p style={{fontSize:16,fontWeight:600,color:done||active?A.text:A.sub,margin:0}}>{item.name}</p>
                  <div style={{display:'flex',gap:12,marginTop:4}}>
                    <span style={{fontSize:12,color:item.value==='High'?A.green:A.dim}}>Value: {item.value}</span>
                    <span style={{fontSize:12,color:item.effort==='High'?A.orange:A.dim}}>Effort: {item.effort}</span>
                  </div>
                </div>
                <span style={{padding:'6px 16px',borderRadius:8,fontSize:13,fontWeight:700,
                  background:`${item.actionColor}12`,color:item.actionColor,
                  opacity:done||active?1:.5,transition:'opacity .4s'}}>
                  {item.action}
                </span>
              </div>
            )
          })}
        </div>
        {p>90 && <div style={{padding:'16px 24px',borderRadius:12,marginTop:12,background:`${col}08`,border:`1px solid ${col}20`}}>
          <p style={{fontSize:16,fontWeight:600,color:col,margin:0}}>✅ "Legacy Migration" killed (score: 8). Capacity redirected to "Cross-OPCO API" (score: 85). Less than 10% wasted.</p>
        </div>}
      </div>)
    }
  }

  return <p style={{color:A.dim,textAlign:'center',padding:32,fontSize:16}}>Select a dimension and click Run</p>
}

// ═══════════════════════════════════════
// CONCEPT VISUALS — SVG mini-diagrams per dimension
// ═══════════════════════════════════════
function ConceptVisualSide({themeId,side}:{themeId:string,side:'project'|'product'}) {
  const isProj = side === 'project'
  const col = isProj ? A.red : A.green
  const vs = {marginTop:16,borderRadius:10,padding:'12px 16px',background:'rgba(255,255,255,.02)',border:'1px solid rgba(255,255,255,.04)'}

  if(themeId==='mindset') return (
    <div style={vs}>
      {isProj ? (
        <svg width="100%" height="50" viewBox="0 0 400 50">
          <line x1={20} y1={25} x2={350} y2={25} stroke={col} strokeWidth={2}/>
          {['Plan','Build','Ship','Done?'].map((s,i)=>(
            <g key={i}><rect x={20+i*88} y={8} width={78} height={34} rx={6} fill={`${col}15`} stroke={`${col}30`} strokeWidth={1}/>
            <text x={59+i*88} y={30} textAnchor="middle" fill={col} fontSize={11} fontWeight={500}>{s}</text></g>
          ))}
        </svg>
      ) : (
        <svg width="100%" height="50" viewBox="0 0 400 50">
          <path d="M 20,25 Q 100,5 180,25 Q 260,45 340,25" fill="none" stroke={col} strokeWidth={2}/>
          {['Discover','Build','Ship','Learn','Repeat'].map((s,i)=>(
            <g key={i}><circle cx={20+i*80} cy={25} r={16} fill={`${col}15`} stroke={`${col}30`} strokeWidth={1}/>
            <text x={20+i*80} y={29} textAnchor="middle" fill={col} fontSize={8} fontWeight={600}>{s}</text></g>
          ))}
        </svg>
      )}
    </div>
  )

  if(themeId==='success') return (
    <div style={vs}>
      {isProj ? (
        <svg width="100%" height="60" viewBox="0 0 400 60">
          <rect x={10} y={15} width={350} height={30} rx={6} fill={`${col}12`} stroke={`${col}30`} strokeWidth={1}/>
          <text x={180} y={35} textAnchor="middle" fill={col} fontSize={12} fontWeight={500} fontFamily="Inter">16 weeks → One delivery at the end</text>
          <circle cx={360} cy={30} r={12} fill={col}/>
          <text x={360} y={35} textAnchor="middle" fill="#fff" fontSize={10}>💥</text>
        </svg>
      ) : (
        <svg width="100%" height="60" viewBox="0 0 400 60">
          {[0,1,2,3,4].map(i=>(
            <g key={i}>
              <rect x={10+i*78} y={10} width={68} height={35} rx={6} fill={`${col}${10+i*4}`} stroke={`${col}30`} strokeWidth={1}/>
              {i>=1 && <text x={44+i*78} y={32} textAnchor="middle" fill={col} fontSize={11} fontWeight={600}>✓ value</text>}
              {i===0 && <text x={44} y={32} textAnchor="middle" fill={A.dim} fontSize={10}>discover</text>}
            </g>
          ))}
        </svg>
      )}
    </div>
  )

  if(themeId==='teams') return (
    <div style={vs}>
      {isProj ? (
        <svg width="100%" height="55" viewBox="0 0 400 55">
          {['Team A','Team B','Team C'].map((t,i)=>(
            <g key={i}>
              <rect x={10+i*135} y={8} width={120} height={28} rx={6} fill={`${col}12`} stroke={i<2?`${A.dim}30`:`${col}30`} strokeWidth={1} strokeDasharray={i<2?"4 3":""}/>
              <text x={70+i*135} y={27} textAnchor="middle" fill={i<2?A.dim:col} fontSize={11}>{t}</text>
              {i<2 && <text x={70+i*135} y={48} textAnchor="middle" fill={A.dim} fontSize={9}>disbanded</text>}
            </g>
          ))}
        </svg>
      ) : (
        <svg width="100%" height="55" viewBox="0 0 400 55">
          <rect x={10} y={5} width={380} height={30} rx={8} fill={`${col}12`} stroke={`${col}25`} strokeWidth={1}/>
          {['PO','UX','Dev','Dev','QA','Eng'].map((r,i)=>(
            <text key={i} x={40+i*62} y={25} textAnchor="middle" fill={col} fontSize={11} fontWeight={500}>{r}</text>
          ))}
          <rect x={10} y={42} width={300} height={5} rx={3} fill={`${col}25`}/>
          <text x={340} y={48} fill={A.dim} fontSize={8}>expertise →</text>
        </svg>
      )}
    </div>
  )

  if(themeId==='discovery') return (
    <div style={vs}>
      {isProj ? (
        <svg width="100%" height="55" viewBox="0 0 400 55">
          <line x1={20} y1={25} x2={350} y2={25} stroke={col} strokeWidth={2}/>
          {['BRD','Approve','Build','Build','UAT','🤞'].map((s,i)=>(
            <g key={i}><circle cx={20+i*66} cy={25} r={8} fill={`${col}25`} stroke={col} strokeWidth={1}/>
            <text x={20+i*66} y={48} textAnchor="middle" fill={A.dim} fontSize={9}>{s}</text></g>
          ))}
        </svg>
      ) : (
        <svg width="100%" height="55" viewBox="0 0 400 55">
          <polyline points="20,35 80,12 140,35 200,12 260,35 320,12 380,35" fill="none" stroke={col} strokeWidth={2}/>
          {[0,1,2,3,4,5,6].map(i=>(
            <circle key={i} cx={20+i*60} cy={i%2===0?35:12} r={5} fill={col}/>
          ))}
          <text x={200} y={52} textAnchor="middle" fill={A.dim} fontSize={9}>feedback at every turn</text>
        </svg>
      )}
    </div>
  )

  if(themeId==='delivery') return (
    <div style={vs}>
      {isProj ? (
        <svg width="100%" height="55" viewBox="0 0 400 55">
          <line x1={20} y1={35} x2={330} y2={35} stroke={`${col}30`} strokeWidth={2}/>
          <rect x={340} y={8} width={40} height={38} rx={5} fill={col}/>
          <text x={360} y={32} textAnchor="middle" fill="#fff" fontSize={12}>💥</text>
          <text x={180} y={52} textAnchor="middle" fill={A.dim} fontSize={9}>months of silence → big bang</text>
        </svg>
      ) : (
        <svg width="100%" height="55" viewBox="0 0 400 55">
          {Array.from({length:14},(_,i)=>(
            <rect key={i} x={15+i*27} y={45-(12+Math.sin(i*.7)*14+10)} width={22} height={12+Math.sin(i*.7)*14+10} rx={3} fill={`${col}${50+i*4}`}/>
          ))}
        </svg>
      )}
    </div>
  )

  if(themeId==='priority') return (
    <div style={vs}>
      {isProj ? (
        <svg width="100%" height="50" viewBox="0 0 400 50">
          {['#1','#2','#3','#4','#5'].map((n,i)=>(
            <g key={i}>
              <rect x={10+i*78} y={8} width={68} height={30} rx={5} fill={`${col}10`} stroke={`${col}20`} strokeWidth={1}/>
              <text x={44+i*78} y={28} textAnchor="middle" fill={col} fontSize={12} fontWeight={600}>{n}</text>
            </g>
          ))}
          <text x={200} y={48} textAnchor="middle" fill={A.dim} fontSize={9}>FIFO — first come, first served</text>
        </svg>
      ) : (
        <svg width="100%" height="60" viewBox="0 0 400 60">
          <rect x={10} y={2} width={180} height={25} rx={4} fill={`${A.red}10`}/>
          <text x={100} y={19} textAnchor="middle" fill={A.red} fontSize={9}>🚫 Kill (low value)</text>
          <rect x={200} y={2} width={190} height={25} rx={4} fill={`${A.blue}10`}/>
          <text x={295} y={19} textAnchor="middle" fill={A.blue} fontSize={9}>🔍 Validate first</text>
          <rect x={10} y={32} width={180} height={25} rx={4} fill={`${A.dim}10`}/>
          <text x={100} y={49} textAnchor="middle" fill={A.dim} fontSize={9}>📋 Fill-in</text>
          <rect x={200} y={32} width={190} height={25} rx={4} fill={`${A.green}10`}/>
          <text x={295} y={49} textAnchor="middle" fill={A.green} fontSize={9}>✅ Do NOW</text>
        </svg>
      )}
    </div>
  )

  return null
}


function Explore({go}:{go:(v:View)=>void}) {
  const [selId,setSelId]=useState(DIMS[0].id)
  const [mode,setMode]=useState<'none'|'project'|'product'>('none')
  const [progress,setProgress]=useState(0)
  const [projProgress,setProjProgress]=useState(0) // keeps project result visible
  const [done,setDone]=useState(false)
  const [productDone,setProductDone]=useState(false)
  const ref=useRef<number>(0)
  const dim=DIMS.find(d=>d.id===selId)!

  // Per-theme duration: teams and discovery get 12s, others 6s
  const getDuration = () => {
    if(selId==='teams'||selId==='discovery') return 0.26 // ~19 seconds
    return 0.52 // ~10 seconds
  }

  const run=(m:'project'|'product')=>{
    clearInterval(ref.current);setMode(m);setProgress(0);setDone(false)
    if(m==='project') setProjProgress(0)
    const speed = getDuration()
    let p=0;ref.current=window.setInterval(()=>{
      p+=speed
      if(p>=100){
        clearInterval(ref.current);setDone(true);setProgress(100)
        if(m==='project') setProjProgress(100)
        if(m==='product') setProductDone(true)
      }
      setProgress(Math.min(p,100))
    },50)
  }
  const reset=()=>{clearInterval(ref.current);setMode('none');setProgress(0);setProjProgress(0);setDone(false);setProductDone(false)}
  useEffect(()=>()=>clearInterval(ref.current),[])
  useEffect(()=>reset(),[selId])

  const phases = mode==='project'?dim.project.phases:mode==='product'?dim.product.phases:[]
  const totalDur = phases.reduce((a,p)=>a+p.d,0)
  const elapsed = (progress/100)*totalDur

  return (
    <div className="scene" style={{padding:'32px'}}>
      <Back go={go} />
      <div style={{maxWidth:1100,margin:'0 auto',paddingTop:24}}>
        {/* Tabs */}
        <div style={{display:'flex',gap:8,marginBottom:32,overflowX:'auto'}}>
          {DIMS.map(d=>(
            <button key={d.id} onClick={()=>setSelId(d.id)} style={{padding:'10px 20px',borderRadius:10,fontSize:15,fontWeight:selId===d.id?700:500,whiteSpace:'nowrap',
              background:selId===d.id?`${d.color}15`:'transparent',color:selId===d.id?A.text:A.sub,border:selId===d.id?`1px solid ${d.color}30`:'1px solid transparent'}}>
              {d.icon} {d.title}
            </button>
          ))}
        </div>

        <AnimatePresence mode="wait">
          <motion.div key={dim.id} initial={{opacity:0,y:16}} animate={{opacity:1,y:0}} exit={{opacity:0}} transition={{duration:.3}}>
            <h2 style={{fontSize:36,fontWeight:800,color:A.text,marginBottom:24}}>{dim.icon} {dim.title}</h2>

            {/* Comparison */}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:20,marginBottom:28}}>
              <div className="card" style={{borderTop:`3px solid ${A.red}`,opacity:mode==='product'?.35:1,transition:'opacity .4s'}}>
                <p style={{fontSize:13,fontWeight:700,color:A.red,letterSpacing:1,textTransform:'uppercase',marginBottom:12}}>Project Mode</p>
                <p style={{fontSize:17,fontWeight:600,color:A.text,marginBottom:16}}>{dim.project.title}</p>
                {dim.project.points.map((pt,i)=><div key={i} style={{display:'flex',gap:10,marginBottom:8}}><span style={{color:A.red,marginTop:2}}>•</span><span style={{fontSize:15,color:A.sub,lineHeight:1.5}}>{pt}</span></div>)}
                <ConceptVisualSide themeId={dim.id} side="project" />
              </div>
              <div className="card" style={{borderTop:`3px solid ${A.green}`,opacity:mode==='project'?.35:1,transition:'opacity .4s'}}>
                <p style={{fontSize:13,fontWeight:700,color:A.green,letterSpacing:1,textTransform:'uppercase',marginBottom:12}}>Product Mode</p>
                <p style={{fontSize:17,fontWeight:600,color:A.text,marginBottom:16}}>{dim.product.title}</p>
                {dim.product.points.map((pt,i)=><div key={i} style={{display:'flex',gap:10,marginBottom:8}}><span style={{color:A.green,marginTop:2}}>•</span><span style={{fontSize:15,color:A.sub,lineHeight:1.5}}>{pt}</span></div>)}
                <ConceptVisualSide themeId={dim.id} side="product" />
              </div>
            </div>

            {/* Simulation */}
            {/* BOTH MODES SIDE BY SIDE — always visible, play independently */}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:20}}>
              {/* PROJECT SIDE */}
              <div className="card" style={{padding:24,borderTop:`3px solid ${A.red}`,opacity:mode==='product'&&!done?.4:1,transition:'opacity .5s'}}>
                <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:16}}>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <div style={{width:7,height:7,borderRadius:4,background:mode==='project'?A.red:A.dim,boxShadow:mode==='project'?`0 0 10px ${A.red}`:'none',transition:'all .3s'}}/>
                    <span style={{fontSize:14,fontWeight:700,color:A.red,letterSpacing:1,textTransform:'uppercase'}}>Project Mode</span>
                  </div>
                  {projProgress===0&&mode!=='project'&&<button onClick={()=>run('project')} style={{padding:'8px 18px',borderRadius:8,background:A.red,color:'#fff',fontSize:13,fontWeight:700}}>▶ Run</button>}
                  {mode==='project'&&!done&&<button onClick={reset} style={{padding:'8px 18px',borderRadius:8,background:'transparent',border:`1px solid ${A.red}30`,color:A.red,fontSize:13}}>⏹ Stop</button>}
                  {projProgress>=100&&<span style={{fontSize:12,color:A.red,fontWeight:600}}>✓ Complete</span>}
                </div>
                <SimRenderer themeId={dim.id} progress={mode==='project'?progress:projProgress} mode="project" />
              </div>

              {/* PRODUCT SIDE */}
              <div className="card" style={{padding:24,borderTop:`3px solid ${A.green}`,
                opacity: mode==='project'&&!done ? .4 : 1,
                boxShadow: projProgress>=100&&!productDone ? `0 0 20px ${A.green}15` : 'none',
                transition:'all .5s'}}>
                <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:16}}>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <div style={{width:7,height:7,borderRadius:4,
                      background: mode==='product'?A.green : projProgress>=100&&!productDone?A.green:A.dim,
                      boxShadow: mode==='product'||( projProgress>=100&&!productDone)?`0 0 10px ${A.green}`:'none',
                      transition:'all .3s'}}/>
                    <span style={{fontSize:14,fontWeight:700,color:A.green,letterSpacing:1,textTransform:'uppercase'}}>
                      Product Mode {projProgress>=100&&!productDone&&mode!=='product'?' — Your turn!':''}
                    </span>
                  </div>
                  {!productDone&&mode!=='product'&&<button onClick={()=>run('product')}
                    style={{padding:'8px 18px',borderRadius:8,fontSize:13,fontWeight:700,color:'#fff',
                      background: projProgress>=100 ? A.green : `${A.green}60`,
                      boxShadow: projProgress>=100 ? `0 0 12px ${A.green}40` : 'none',
                      transition:'all .3s'}}>
                    ▶ Run
                  </button>}
                  {mode==='product'&&!done&&<button onClick={reset} style={{padding:'8px 18px',borderRadius:8,background:'transparent',border:`1px solid ${A.green}30`,color:A.green,fontSize:13}}>⏹ Stop</button>}
                  {productDone&&<span style={{fontSize:12,color:A.green,fontWeight:600}}>✓ Complete</span>}
                </div>
                <SimRenderer themeId={dim.id} progress={mode==='product'?progress:0} mode="product" />
              </div>
            </div>

            {/* Metrics + Insight — only after PRODUCT mode completes */}
            {productDone&&<motion.div initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} transition={{delay:.3}}>
              <div style={{display:'grid',gridTemplateColumns:`repeat(${dim.metrics.length},1fr)`,gap:16,marginTop:20}}>
                {dim.metrics.map((m,i)=><div key={i} className="card" style={{textAlign:'center',padding:'20px 14px'}}>
                  <p style={{fontSize:12,color:A.dim,fontWeight:600,letterSpacing:.5,textTransform:'uppercase',marginBottom:10}}>{m.l}</p>
                  <div style={{display:'flex',justifyContent:'center',gap:20}}>
                    <div><p className="mono" style={{fontSize:24,fontWeight:800,color:A.red,margin:0}}>{m.p}</p><p style={{fontSize:10,color:A.dim,margin:'4px 0 0'}}>Project</p></div>
                    <div style={{width:1,background:'rgba(255,255,255,.06)'}} />
                    <div><p className="mono" style={{fontSize:24,fontWeight:800,color:A.green,margin:0}}>{m.d}</p><p style={{fontSize:10,color:A.dim,margin:'4px 0 0'}}>Product</p></div>
                  </div>
                </div>)}
              </div>
              <div style={{marginTop:16,padding:'16px 24px',borderRadius:12,background:`${dim.color}08`,borderLeft:`4px solid ${dim.color}`}}>
                <p style={{fontSize:16,color:A.text,lineHeight:1.6,margin:0}}>💡 <strong style={{color:dim.color}}>Insight:</strong> {dim.insight}</p>
              </div>
              <div style={{textAlign:'center',marginTop:16}}>
                <button onClick={reset} style={{padding:'10px 24px',borderRadius:10,background:'transparent',border:`1px solid ${A.dim}`,color:A.sub,fontSize:14}}>↻ Reset Both</button>
              </div>
            </motion.div>}
          </motion.div>
        </AnimatePresence>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════
// CONTEXT QUESTIONS
// ═══════════════════════════════════════
const QS=[
  {t:'Success',q:'How does Digital Products define "success" for an initiative today?',o:['On time and on budget','User adoption rates','Measurable business outcome','Not clearly defined']},
  {t:'Teams',q:'Are your teams persistent or assembled per project?',o:['Fully persistent','Mostly persistent','Mostly project-based','Fully project-based']},
  {t:'Planning',q:'How far ahead do you plan work in detail?',o:['2-4 weeks','1-3 months','6-12 months','12+ months']},
  {t:'Discovery',q:'How often does someone on your team talk to an actual end user?',o:['Weekly','Monthly','Quarterly','Almost never']},
  {t:'Delivery',q:'How frequently does your team ship value to production?',o:['Daily or weekly','Every 2-4 weeks','Quarterly','Big-bang releases']},
  {t:'Ownership',q:'Who owns a product capability after it goes live?',o:['Same team that built it','Separate operations team','Unclear','Nobody specifically']},
  {t:'Priority',q:'How are priorities decided across your portfolio?',o:['Evidence-based scoring','Committee / governance body','Whoever asks loudest','First come, first served']},
  {t:'KPIs',q:'What KPIs does your team track most actively?',o:['Business outcomes (adoption, NPS)','Delivery metrics (velocity)','Traditional (time/budget)','Not consistently tracked']},
  {t:'Roadmap',q:'What does your roadmap look like today?',o:['Now / Next / Later themes','Features with fixed dates','Gantt chart','We don\'t really have one']},
  {t:'Funding',q:'How is work funded?',o:['Persistent team funding','Per-project business case','Annual budget allocation','Mix of approaches']},
  {t:'Governance',q:'What do governance reviews look like?',o:['Demos and outcome metrics','Stage-gate milestone reviews','Steering committee decks','Varies by project']},
  {t:'Customer',q:'How close is your team to actual end users?',o:['Direct weekly contact','Through a BA/PM proxy','Only during testing','Very distant']},
]

function Context({go}:{go:(v:View)=>void}) {
  const [ans,setAns]=useState<Record<number,number>>({})
  const [cur,setCur]=useState(0)
  const [showResult,setShowResult]=useState(false)
  const q=QS[cur]
  const allAnswered = Object.keys(ans).length === QS.length

  // Maturity scoring: option index 0=1pt, 1=2pt, 2=3pt, 3=4pt (higher = more product-oriented)
  const maturityScore = Object.values(ans).reduce((a,v)=>a+(v+1),0) / (Object.keys(ans).length || 1)

  const exportData = () => {
    const data = {
      session: 'Avangrid Digital Products — Reality Check',
      date: new Date().toISOString(),
      participant: 'Heather Lajoie',
      maturityScore: maturityScore.toFixed(1),
      answers: QS.map((q,i) => ({
        theme: q.t,
        question: q.q,
        answer: q.o[ans[i]] || 'Not answered',
        answerIndex: ans[i] ?? -1,
        score: (ans[i] ?? -1) + 1,
      }))
    }
    // Save to localStorage
    localStorage.setItem('avangrid_reality_check', JSON.stringify(data))
    // Download as JSON file
    const blob = new Blob([JSON.stringify(data,null,2)], {type:'application/json'})
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url; a.download = 'avangrid_reality_check.json'; a.click()
    URL.revokeObjectURL(url)
  }

  const finish = () => {
    // Save to localStorage immediately
    const data = {
      session: 'Avangrid Digital Products — Reality Check',
      date: new Date().toISOString(),
      maturityScore: maturityScore.toFixed(1),
      answers: QS.map((q,i) => ({theme:q.t, question:q.q, answer:q.o[ans[i]]||'', score:(ans[i]??-1)+1}))
    }
    localStorage.setItem('avangrid_reality_check', JSON.stringify(data))
    setShowResult(true)
  }

  if(showResult) {
    return (
      <div className="scene" style={{padding:'32px'}}><Back go={go}/>
        <div style={{maxWidth:800,margin:'0 auto',paddingTop:32}}>
          <p className="mono" style={{fontSize:11,color:A.orange,letterSpacing:2}}>RESULTS</p>
          <h2 style={{fontSize:32,fontWeight:800,color:A.text,marginBottom:24}}>Your Product Maturity</h2>

          {/* Score card */}
          <div className="card" style={{textAlign:'center',padding:'32px 24px',marginBottom:24,borderTop:`3px solid ${maturityScore<2?A.red:maturityScore<3?A.orange:A.green}`}}>
            <p style={{fontSize:16,color:A.sub,margin:'0 0 8px'}}>Average Maturity Score</p>
            <p className="mono" style={{fontSize:64,fontWeight:900,color:maturityScore<2?A.red:maturityScore<3?A.orange:A.green,margin:0,lineHeight:1}}>
              {maturityScore.toFixed(1)}
            </p>
            <p style={{fontSize:14,color:A.dim,margin:'8px 0 0'}}>out of 4.0</p>
            <p style={{fontSize:18,color:A.text,marginTop:16,fontWeight:600}}>
              {maturityScore<1.5?'Early stage — significant opportunity to build product habits'
                :maturityScore<2.5?'Foundation in place — ready to accelerate the shift'
                :maturityScore<3.5?'Strong base — focus on optimization and scaling'
                :'Product-led — maintain momentum and innovate'}
            </p>
          </div>

          {/* Per-question breakdown */}
          <div className="card" style={{padding:24,marginBottom:24}}>
            <p style={{fontSize:16,fontWeight:700,color:A.text,marginBottom:16}}>Breakdown by Dimension</p>
            {QS.map((q,i) => {
              const score = (ans[i] ?? 0) + 1
              const color = score<=1?A.red:score<=2?A.orange:score<=3?A.blue:A.green
              return (
                <div key={i} style={{display:'flex',alignItems:'center',gap:12,marginBottom:10}}>
                  <span style={{width:100,fontSize:13,color:A.sub,flexShrink:0}}>{q.t}</span>
                  <div style={{flex:1,height:8,borderRadius:4,background:'rgba(255,255,255,.04)',overflow:'hidden'}}>
                    <div style={{height:'100%',borderRadius:4,background:color,width:`${(score/4)*100}%`}} />
                  </div>
                  <span className="mono" style={{width:24,fontSize:13,fontWeight:700,color,textAlign:'right'}}>{score}</span>
                  <span style={{width:160,fontSize:12,color:A.dim}}>{q.o[ans[i]]}</span>
                </div>
              )
            })}
          </div>

          {/* Export buttons */}
          <div style={{display:'flex',gap:12,justifyContent:'center'}}>
            <button onClick={exportData} style={{padding:'12px 28px',borderRadius:10,background:A.blue,color:'#fff',fontSize:15,fontWeight:700}}>
              📥 Export as JSON (for KB)
            </button>
            <button onClick={()=>{setShowResult(false);setAns({});setCur(0)}} style={{padding:'12px 28px',borderRadius:10,background:'transparent',border:`1px solid ${A.dim}`,color:A.sub,fontSize:15}}>
              ↻ Retake
            </button>
            <button onClick={()=>go('home')} style={{padding:'12px 28px',borderRadius:10,background:'transparent',border:`1px solid ${A.dim}`,color:A.sub,fontSize:15}}>
              ← Back to Home
            </button>
          </div>

          <p style={{textAlign:'center',marginTop:20,fontSize:12,color:A.dim}}>
            Results saved to browser storage and can be exported as JSON for the Knowledge Base
          </p>
        </div>
      </div>
    )
  }

  return (
    <div className="scene" style={{padding:'32px'}}><Back go={go}/>
      <div style={{maxWidth:700,margin:'0 auto',paddingTop:32}}>
        <p className="mono" style={{fontSize:11,color:A.dim,letterSpacing:2}}>AVANGRID REALITY CHECK</p>
        <h2 style={{fontSize:32,fontWeight:800,color:A.text,marginBottom:8}}>Your Reality Today</h2>
        <p style={{fontSize:16,color:A.sub,marginBottom:28}}>Each question connects to the dimensions we explored.</p>
        <div style={{display:'flex',gap:4,marginBottom:28}}>{QS.map((_,i)=><div key={i} onClick={()=>setCur(i)} style={{flex:1,height:4,borderRadius:2,cursor:'pointer',background:ans[i]!==undefined?A.orange:i===cur?`${A.orange}50`:'rgba(255,255,255,.06)'}}/>)}</div>
        <AnimatePresence mode="wait"><motion.div key={cur} initial={{opacity:0,x:30}} animate={{opacity:1,x:0}} exit={{opacity:0,x:-30}} transition={{duration:.2}}>
          <p className="mono" style={{fontSize:12,color:A.orange,marginBottom:8}}>{cur+1} / {QS.length} — {q.t}</p>
          <h3 style={{fontSize:22,fontWeight:700,color:A.text,marginBottom:20,lineHeight:1.4}}>{q.q}</h3>
          <div style={{display:'flex',flexDirection:'column',gap:10}}>
            {q.o.map((opt,i)=><button key={i} className="card" onClick={()=>{setAns(p=>({...p,[cur]:i}));if(cur<QS.length-1)setTimeout(()=>setCur(cur+1),250)}}
              style={{textAlign:'left',padding:'16px 20px',cursor:'pointer',borderColor:ans[cur]===i?`${A.orange}40`:undefined,background:ans[cur]===i?`${A.orange}08`:undefined}}>
              <span style={{fontSize:16,color:ans[cur]===i?A.text:A.sub,fontWeight:ans[cur]===i?600:400}}>{opt}</span></button>)}
          </div>
        </motion.div></AnimatePresence>
        <div style={{display:'flex',justifyContent:'space-between',marginTop:24}}>
          <button onClick={()=>setCur(Math.max(0,cur-1))} style={{padding:'10px 20px',borderRadius:10,background:'transparent',border:'1px solid rgba(255,255,255,.08)',color:A.sub,fontSize:14}}>← Previous</button>
          {cur<QS.length-1?
            <button onClick={()=>setCur(cur+1)} style={{padding:'10px 20px',borderRadius:10,background:A.orange,color:'#fff',fontSize:14,fontWeight:600}}>Next →</button>
          : <button onClick={finish} style={{padding:'10px 20px',borderRadius:10,background:A.green,color:'#fff',fontSize:14,fontWeight:600}}>✓ See Results</button>}
        </div>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════
// ASSESSMENT RADAR
// ═══════════════════════════════════════
const AX:{name:string,q:string,opts:{label:string,desc:string}[]}[] = [
  {name:'Discovery Habits',q:'How often does your team explore user problems before building solutions?',
    opts:[{label:'Never',desc:'We build what is requested — no user research'},{label:'Occasionally',desc:'Some interviews happen but not structured'},{label:'Regularly',desc:'User research is part of every initiative'},{label:'Continuously',desc:'Weekly discovery rituals with real users'}]},
  {name:'Outcome Focus',q:'How does your team define and measure success?',
    opts:[{label:'Delivery-focused',desc:'Success = on time, on budget, on scope'},{label:'Output-focused',desc:'Success = features shipped and tasks done'},{label:'Outcome-aware',desc:'We track adoption but delivery metrics dominate'},{label:'Outcome-driven',desc:'OKRs with measurable business impact'}]},
  {name:'Team Persistence',q:'How stable are your teams across initiatives?',
    opts:[{label:'Fully project-based',desc:'New team assembled for each project'},{label:'Mostly temporary',desc:'Core people stay but most rotate'},{label:'Mostly persistent',desc:'Teams stay together across sprints'},{label:'Fully persistent',desc:'Cross-functional teams own product areas long-term'}]},
  {name:'User Empathy',q:'How well does your team understand end-user pain points?',
    opts:[{label:'Very distant',desc:'We rarely interact with actual users'},{label:'Through proxies',desc:'BA or PM relays user needs to the team'},{label:'Periodic contact',desc:'User interviews happen monthly or quarterly'},{label:'Deeply embedded',desc:'Team members observe and talk to users weekly'}]},
  {name:'Data-Driven Decisions',q:'How much do data and experiments drive your prioritization?',
    opts:[{label:'Gut feeling',desc:'Decisions based on opinions and politics'},{label:'Some data',desc:'We look at data but opinions still dominate'},{label:'Data-informed',desc:'Metrics inform decisions alongside judgment'},{label:'Data-driven',desc:'A/B tests and experiments drive every major decision'}]},
  {name:'Continuous Delivery',q:'How frequently does your team ship value to production?',
    opts:[{label:'Quarterly or less',desc:'Big-bang releases a few times a year'},{label:'Monthly',desc:'Planned monthly release cycles'},{label:'Bi-weekly',desc:'Releases aligned with sprint cadence'},{label:'On demand',desc:'Teams deploy independently, multiple times per week'}]},
  {name:'Portfolio Governance',q:'How is work across the portfolio prioritized and reviewed?',
    opts:[{label:'No formal process',desc:'Each team decides independently'},{label:'Heavy stage-gates',desc:'Steering committees with status decks'},{label:'Balanced reviews',desc:'Regular portfolio reviews with clear criteria'},{label:'Lean governance',desc:'Quarterly learn cycles, demo-based reviews'}]},
  {name:'Product Culture',q:'How does leadership view experimentation and learning from failure?',
    opts:[{label:'Failure is punished',desc:'Mistakes lead to blame and consequences'},{label:'Failure is tolerated',desc:'Mistakes are OK but not encouraged'},{label:'Learning is valued',desc:'Retrospectives happen and insights are shared'},{label:'Experimentation celebrated',desc:'Teams are expected to test, fail, and learn fast'}]},
  {name:'Prioritization',q:'How are competing demands ranked and sequenced?',
    opts:[{label:'Loudest voice wins',desc:'Whoever pushes hardest gets prioritized'},{label:'Committee decides',desc:'Governance body reviews and approves'},{label:'Criteria-based',desc:'Scoring framework exists but not always followed'},{label:'Evidence-based',desc:'RICE/WSJF scoring with data drives every decision'}]},
  {name:'Roadmap Maturity',q:'What does your roadmap look like?',
    opts:[{label:'No roadmap',desc:'Work is reactive — no forward planning'},{label:'Gantt chart',desc:'Features mapped to fixed dates and quarters'},{label:'Theme-based',desc:'Organized by outcomes but still date-driven'},{label:'Now / Next / Later',desc:'Outcome-based with decreasing certainty over time'}]},
  {name:'Funding Model',q:'How is work funded in your organization?',
    opts:[{label:'Per-project',desc:'Each initiative needs its own business case'},{label:'Annual budget',desc:'Budget allocated yearly, use-it-or-lose-it'},{label:'Mixed approach',desc:'Some persistent funding, some project-based'},{label:'Team-based',desc:'Persistent teams funded as capacity, rebalanced quarterly'}]},
  {name:'Customer Proximity',q:'How close is your team to the actual end users?',
    opts:[{label:'Very distant',desc:'No direct contact with users'},{label:'Through proxy',desc:'BA or PM is the voice of the customer'},{label:'Periodic access',desc:'User interviews or surveys happen occasionally'},{label:'Embedded',desc:'Team has direct, weekly contact with real users'}]},
]
function Assess({go}:{go:(v:View)=>void}) {
  const [know,setKnow]=useState<Record<number,number>>({});const [want,setWant]=useState<Record<number,number>>({})
  const [phase,setPhase]=useState<'know'|'want'|'done'>('know');const [cur,setCur]=useState(0)
  const answer=(v:number)=>{if(phase==='know')setKnow(p=>({...p,[cur]:v}));else setWant(p=>({...p,[cur]:v}));if(cur<AX.length-1)setCur(cur+1);else if(phase==='know'){setPhase('want');setCur(0)}else setPhase('done')}
  const data=AX.map((d,i)=>({subject:d.name,today:know[i]||0,aspiration:want[i]||0}))
  const avg=Object.values(know).length?Object.values(know).reduce((a,b)=>a+b,0)/Object.values(know).length:0

  const exportAssess = () => {
    const d = {session:'Product Maturity Assessment',date:new Date().toISOString(),participant:'Heather Lajoie',
      averageToday:avg.toFixed(1), averageAspiration:avgWant.toFixed(1),
      dimensions:AX.map((ax,i)=>({name:ax.name,today:know[i]||0,todayLabel:ax.opts[(know[i]||1)-1]?.label,aspiration:want[i]||0,aspirationLabel:ax.opts[(want[i]||1)-1]?.label,gap:(want[i]||0)-(know[i]||0)}))}
    localStorage.setItem('avangrid_maturity_assessment',JSON.stringify(d))
    const blob = new Blob([JSON.stringify(d,null,2)],{type:'application/json'})
    const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='avangrid_maturity_assessment.json'; a.click(); URL.revokeObjectURL(url)
  }

  const avgWant = Object.values(want).length ? Object.values(want).reduce((a,b)=>a+b,0)/Object.values(want).length : 0

  if(phase==='done') return (
    <div className="scene" style={{padding:'32px'}}><Back go={go}/>
      <div style={{maxWidth:1100,margin:'0 auto',paddingTop:32}}>
        <h2 style={{fontSize:32,fontWeight:800,color:A.text,marginBottom:24}}>Product Maturity Snapshot</h2>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:28}}>
          {/* Radar chart */}
          <div>
            <ResponsiveContainer width="100%" height={400}><RadarChart data={data}><PolarGrid stroke="rgba(255,255,255,.06)"/>
              <PolarAngleAxis dataKey="subject" tick={{fill:A.sub,fontSize:10}}/>
              <PolarRadiusAxis angle={90} domain={[0,4]} tick={false}/>
              <Radar name="Where you want to be (12mo)" dataKey="aspiration" stroke={A.green} fill={A.green} fillOpacity={.1} strokeWidth={2} strokeDasharray="6 4"/>
              <Radar name="Where you are today" dataKey="today" stroke={A.red} fill={A.red} fillOpacity={.15} strokeWidth={2}/>
              <Legend wrapperStyle={{fontSize:11}}/>
            </RadarChart></ResponsiveContainer>
            <div style={{display:'flex',gap:20,justifyContent:'center',marginTop:8}}>
              <span style={{display:'flex',alignItems:'center',gap:6,fontSize:13,color:A.sub}}><span style={{width:14,height:3,background:A.red,borderRadius:2,display:'inline-block'}}/> Today</span>
              <span style={{display:'flex',alignItems:'center',gap:6,fontSize:13,color:A.sub}}><span style={{width:14,height:3,background:A.green,borderRadius:2,display:'inline-block',borderTop:'1px dashed '+A.green}}/> 12-month target</span>
            </div>
          </div>

          {/* Score + Gap */}
          <div>
            {/* Score cards */}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12,marginBottom:20}}>
              <div className="card" style={{textAlign:'center',borderTop:`3px solid ${A.red}`}}>
                <p style={{fontSize:12,color:A.dim,margin:'0 0 4px'}}>Today</p>
                <p className="mono" style={{fontSize:40,fontWeight:900,color:avg<2?A.red:avg<3?A.orange:A.green,margin:0}}>{avg.toFixed(1)}</p>
                <p style={{fontSize:11,color:A.dim,margin:'4px 0 0'}}>out of 4.0</p>
              </div>
              <div className="card" style={{textAlign:'center',borderTop:`3px solid ${A.green}`}}>
                <p style={{fontSize:12,color:A.dim,margin:'0 0 4px'}}>12-Month Target</p>
                <p className="mono" style={{fontSize:40,fontWeight:900,color:A.green,margin:0}}>{avgWant.toFixed(1)}</p>
                <p style={{fontSize:11,color:A.dim,margin:'4px 0 0'}}>out of 4.0</p>
              </div>
            </div>

            {/* Per-dimension gap */}
            <p style={{fontSize:16,fontWeight:600,color:A.text,marginBottom:12}}>Gap by Dimension</p>
            {AX.map((d,i)=>{
              const today=know[i]||0; const target=want[i]||0; const gap=target-today
              return <div key={i} style={{marginBottom:10}}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:3}}>
                  <span style={{fontSize:13,color:A.sub}}>{d.name}</span>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <span className="mono" style={{fontSize:12,color:A.red}}>{today}</span>
                    <span style={{fontSize:10,color:A.dim}}>→</span>
                    <span className="mono" style={{fontSize:12,color:A.green}}>{target}</span>
                    <span className="mono" style={{fontSize:11,fontWeight:700,color:gap>1.5?A.red:gap>.5?A.orange:A.green,width:28,textAlign:'right'}}>+{gap}</span>
                  </div>
                </div>
                <div style={{height:8,borderRadius:4,background:'rgba(255,255,255,.04)',position:'relative',overflow:'hidden'}}>
                  <div style={{position:'absolute',left:0,top:0,bottom:0,width:`${(today/4)*100}%`,borderRadius:4,background:A.red+60}} />
                  <div style={{position:'absolute',left:`${(today/4)*100}%`,top:0,bottom:0,width:`${(gap/4)*100}%`,borderRadius:'0 4px 4px 0',background:`${A.green}30`,borderRight:`2px solid ${A.green}`}} />
                </div>
              </div>
            })}

            {/* Actions */}
            <div style={{display:'flex',gap:10,marginTop:20}}>
              <button onClick={exportAssess} style={{padding:'10px 20px',borderRadius:10,background:A.blue,color:'#fff',fontSize:14,fontWeight:700}}>📥 Export JSON</button>
              <button onClick={()=>{setKnow({});setWant({});setPhase('know');setCur(0)}} style={{padding:'10px 20px',borderRadius:10,background:'transparent',border:`1px solid ${A.dim}`,color:A.sub,fontSize:14}}>↻ Retake</button>
            </div>
            <p style={{fontSize:11,color:A.dim,marginTop:8}}>Results saved to browser + exportable as JSON for Knowledge Base</p>
          </div>
        </div>
      </div>
    </div>)

  return (
    <div className="scene" style={{padding:'32px'}}><Back go={go}/>
      <div style={{maxWidth:700,margin:'0 auto',paddingTop:32}}>
        <p className="mono" style={{fontSize:12,color:phase==='know'?A.blue:A.green,letterSpacing:2}}>{phase==='know'?'PART 1: WHERE ARE YOU TODAY?':'PART 2: WHERE DO YOU WANT TO BE IN 12 MONTHS?'}</p>
        <h2 style={{fontSize:28,fontWeight:800,color:A.text,marginTop:4,marginBottom:8}}>{AX[cur].name}</h2>
        <p style={{fontSize:16,color:A.sub,marginBottom:8}}>{AX[cur].q}</p>
        {phase==='want' && <p style={{fontSize:14,color:A.green,marginBottom:16,fontWeight:500}}>Now select the level you want your team to reach within the next 12 months:</p>}
        <div style={{display:'flex',gap:4,marginBottom:24}}>{AX.map((_,i)=><div key={i} style={{flex:1,height:4,borderRadius:2,background:(phase==='know'?know[i]:want[i])!==undefined?(phase==='know'?A.blue:A.green):i===cur?'rgba(255,255,255,.12)':'rgba(255,255,255,.04)'}}/>)}</div>
        <div style={{display:'flex',flexDirection:'column',gap:10}}>
          {AX[cur].opts.map((opt,oi)=>
            <button key={oi} className="card" onClick={()=>answer(oi+1)} style={{display:'flex',alignItems:'center',gap:16,padding:'16px 20px',textAlign:'left',cursor:'pointer'}}>
              <div style={{width:40,height:40,borderRadius:10,background:`${phase==='know'?A.blue:A.green}12`,display:'flex',alignItems:'center',justifyContent:'center',fontWeight:800,fontSize:18,color:phase==='know'?A.blue:A.green,flexShrink:0}}>{oi+1}</div>
              <div><p style={{fontSize:16,fontWeight:600,color:A.text,margin:0}}>{opt.label}</p><p style={{fontSize:13,color:A.sub,margin:'2px 0 0'}}>{opt.desc}</p></div>
            </button>)}
        </div>
      </div>
    </div>)
}

// ═══════════════════════════════════════
// STRUCTURES (Gartner) — with SVG org charts
// ═══════════════════════════════════════
type OrgNode = {id:string,label:string,sub?:string,x:number,y:number,w:number,h?:number,color:string,level:number}
type OrgEdge = [string,string]

interface OrgData {
  name:string; ind:string; rev:string; pct:string; c:string
  team:string; uniq:string; employees?:string
  nodes:OrgNode[]; edges:OrgEdge[]
}

const ORGS:OrgData[] = [
  {name:'Aon',ind:'Professional Services',rev:'$13.4B',pct:'100%',employees:'50,000+',c:A.green,
    team:'Product Owner + Software Engineers + UX Designer (dedicated per team). Platform Owner for foundational teams. Dedicated Data team: Analyst + Engineer + Steward.',
    uniq:'Platform teams see themselves as a "link in value creation chain" — they understand how their work enables product teams to serve end customers. 4-layer platform stack is the most granular of all examples.',
    nodes:[
      {id:'cio',label:'CIO',x:400,y:15,w:120,color:A.green,level:0},
      {id:'sl1',label:'Reinsurance',sub:'Solution Line',x:120,y:90,w:150,color:A.green,level:1},
      {id:'sl2',label:'Commercial Risk',sub:'Solution Line',x:340,y:90,w:160,color:A.green,level:1},
      {id:'sl3',label:'Health Solutions',sub:'Solution Line',x:570,y:90,w:160,color:A.green,level:1},
      {id:'pl1',label:'Risk Advisory',sub:'Product Line',x:60,y:170,w:130,color:A.blue,level:2},
      {id:'pl2',label:'Risk Transfer',sub:'Product Line',x:210,y:170,w:130,color:A.blue,level:2},
      {id:'pl3',label:'Client Mgmt',sub:'Product Line',x:370,y:170,w:130,color:A.blue,level:2},
      {id:'pl4',label:'Claims',sub:'Product Line',x:520,y:170,w:110,color:A.blue,level:2},
      {id:'t1',label:'Team',sub:'8-12',x:60,y:240,w:70,color:A.sub,level:3},
      {id:'t2',label:'Team',sub:'8-12',x:145,y:240,w:70,color:A.sub,level:3},
      {id:'t3',label:'Team',sub:'8-12',x:230,y:240,w:70,color:A.sub,level:3},
      {id:'t4',label:'Team',sub:'8-12',x:390,y:240,w:70,color:A.sub,level:3},
      {id:'t5',label:'Team',sub:'8-12',x:475,y:240,w:70,color:A.sub,level:3},
      {id:'t6',label:'Team',sub:'8-12',x:540,y:240,w:70,color:A.sub,level:3},
      {id:'pf',label:'IT Enablement Platforms (APIs, Data & Analytics, Cloud Infra, End User)',x:120,y:310,w:560,color:A.dim,level:4},
    ],
    edges:[['cio','sl1'],['cio','sl2'],['cio','sl3'],['sl1','pl1'],['sl1','pl2'],['sl2','pl3'],['sl2','pl4'],['pl1','t1'],['pl1','t2'],['pl2','t3'],['pl3','t4'],['pl3','t5'],['pl4','t6']]
  },
  {name:'Ford',ind:'Automotive',rev:'$136B',pct:'100%',employees:'183,000+',c:A.orange,
    team:'Quad leadership per product line: Product/Service Owner + Product/Service Manager + Product Designer + Architect. Each manages talent for their domain. Teams: 8-12 members.',
    uniq:'Largest scale: 8 PDOs, 140+ product groups, 330+ product lines. 4 co-equal leaders per line — not one PM but four disciplines with shared accountability. Some teams up to 100% contract staff.',
    nodes:[
      {id:'cto',label:'CTO / Exec Director',sub:'per PDO',x:350,y:15,w:200,color:A.orange,level:0},
      {id:'pdo',label:'Product-Driven Org',sub:'8 PDOs total',x:350,y:85,w:200,color:A.orange,level:1},
      {id:'pg1',label:'Collections',sub:'Product Group',x:100,y:155,w:140,color:A.blue,level:2},
      {id:'pg2',label:'Order Fulfillment',sub:'Product Group',x:300,y:155,w:160,color:A.blue,level:2},
      {id:'pg3',label:'Supply Chain',sub:'Product Group',x:520,y:155,w:140,color:A.blue,level:2},
      {id:'pl1',label:'Scheduling',sub:'Product Line',x:230,y:230,w:120,color:A.sub,level:3},
      {id:'pl2',label:'Order Mgmt',sub:'Product Line',x:380,y:230,w:120,color:A.sub,level:3},
      {id:'t1',label:'4-20 Teams',sub:'8-12 each',x:230,y:295,w:120,color:A.dim,level:3},
      {id:'t2',label:'4-20 Teams',sub:'8-12 each',x:380,y:295,w:120,color:A.dim,level:3},
      {id:'pf',label:'API-Enabled Reusable Products (Infrastructure)',x:160,y:360,w:480,color:A.dim,level:4},
    ],
    edges:[['cto','pdo'],['pdo','pg1'],['pdo','pg2'],['pdo','pg3'],['pg2','pl1'],['pg2','pl2'],['pl1','t1'],['pl2','t2']]
  },
  {name:'Globo',ind:'Media',rev:'$3B',pct:'70%',employees:'14,000+',c:A.purple,
    team:'Team Leader + UX Designer + Product Owner + Agile Coach + Software Engineers + Data Scientists. Squads: 8-15 members. Dual reporting: tribe leader (work) + technology manager (skills).',
    uniq:'Multidisciplinary fusion teams: IT + business experts in every squad. Cross-cutting CoEs for Lean Agile and Design. Squads organized by user journeys. 3-layer platform stack.',
    nodes:[
      {id:'ed',label:'Exec Director',sub:'Digital Hub',x:350,y:15,w:180,color:A.purple,level:0},
      {id:'c1',label:'Corporate',sub:'Value Stream',x:100,y:85,w:130,color:A.purple,level:1},
      {id:'c2',label:'Advertising',sub:'Value Stream',x:300,y:85,w:130,color:A.purple,level:1},
      {id:'c3',label:'Direct-to-Consumer',sub:'Value Stream',x:500,y:85,w:180,color:A.purple,level:1},
      {id:'p1',label:'Globoplay',sub:'Product Line',x:440,y:160,w:120,color:A.blue,level:2},
      {id:'p2',label:'Other Products',sub:'Product Line',x:580,y:160,w:130,color:A.blue,level:2},
      {id:'s1',label:'Squad',sub:'8-15',x:420,y:230,w:80,color:A.sub,level:3},
      {id:'s2',label:'Squad',sub:'8-15',x:520,y:230,w:80,color:A.sub,level:3},
      {id:'s3',label:'Squad',sub:'8-15',x:620,y:230,w:80,color:A.sub,level:3},
      {id:'cc',label:'Cross-cutting: EA, DevOps, DevFinOps, Security, Lean Agile CoE',x:100,y:295,w:500,color:A.dim,level:4},
      {id:'pf',label:'Platforms: CX & Business | Digital Dev | Foundational (GloboID, AI, Cloud)',x:100,y:345,w:600,color:A.dim,level:4},
    ],
    edges:[['ed','c1'],['ed','c2'],['ed','c3'],['c3','p1'],['c3','p2'],['p1','s1'],['p1','s2'],['p2','s3']]
  },
  {name:'HealthPartners',ind:'Health Insurance',rev:'N/A',pct:'60%',employees:'26,000+',c:A.blue,
    team:'Product Owner (sourced from business) + Scrum Master + Engineers + Business Analyst (dedicated). Architects, UI/UX, Security, Digital Content (shared across teams). POs can support multiple related teams.',
    uniq:'CTO (not CIO) at top — technical leadership emphasis. Product Owners sourced from business side — they own the vision. Separate Resourcing Managers handle talent development. Simplest foundational layer.',
    nodes:[
      {id:'cto',label:'CTO',x:370,y:15,w:100,color:A.blue,level:0},
      {id:'c1',label:'Care Delivery',sub:'Category',x:80,y:85,w:140,color:A.blue,level:1},
      {id:'c2',label:'Digital',sub:'Category',x:260,y:85,w:110,color:A.blue,level:1},
      {id:'c3',label:'Health Plan',sub:'Category',x:410,y:85,w:130,color:A.blue,level:1},
      {id:'c4',label:'Enterprise',sub:'Category',x:580,y:85,w:120,color:A.blue,level:1},
      {id:'p1',label:'B2B',sub:'Product Line',x:220,y:160,w:90,color:A.sub,level:2},
      {id:'p2',label:'Consumer',sub:'Product Line',x:330,y:160,w:100,color:A.sub,level:2},
      {id:'t1',label:'Team',sub:'8-12',x:220,y:225,w:80,color:A.dim,level:3},
      {id:'t2',label:'Team',sub:'8-12',x:320,y:225,w:80,color:A.dim,level:3},
      {id:'t3',label:'Team',sub:'8-12',x:420,y:225,w:80,color:A.dim,level:3},
      {id:'pf',label:'Technology Platforms (Cloud, Database, Content Management)',x:130,y:295,w:520,color:A.dim,level:4},
    ],
    edges:[['cto','c1'],['cto','c2'],['cto','c3'],['cto','c4'],['c2','p1'],['c2','p2'],['p1','t1'],['p2','t2'],['c3','t3']]
  },
  {name:'Avangrid',ind:'Energy Utility (Iberdrola)',rev:'$7.4B',pct:'Proposed — Wave 1',employees:'8,000+',c:A.green,
    team:'Per team (8-10 members): Product Owner + Dev Lead + 2-3 Developers + QA Engineer + UX Designer (shared across 2 teams). Champions per area act as business liaison. BRS as first pilot consumer OPCO.',
    uniq:'Built on proven assets from 4 existing Avangrid projects (P&S, Digital CoE, Infrastructure, Renewables DTO). Wave 1 validates the model with BRS pilot in 14 weeks. Platform team designed to serve ALL OPCOs (BRS, CX, Networks). Scalable from day one — add product lines without adding coordination overhead.',
    nodes:[
      {id:'vp',label:'VP Digital Products',sub:'Vitor Cruz',x:310,y:15,w:220,color:A.green,level:0},
      {id:'sap',label:'SAP Products',sub:'Product Category',x:50,y:95,w:150,color:A.blue,level:1},
      {id:'dig',label:'Digital Applications',sub:'Product Category',x:250,y:95,w:180,color:A.purple,level:1},
      {id:'data',label:'Data & Analytics',sub:'Product Category',x:490,y:95,w:170,color:A.orange,level:1},
      {id:'sf',label:'SAP Finance',sub:'Product Line',x:10,y:180,w:120,color:A.blue,level:2},
      {id:'ss',label:'SAP SCM',sub:'Product Line',x:150,y:180,w:110,color:A.blue,level:2},
      {id:'cp',label:'Customer Portal',sub:'Product Line',x:280,y:180,w:130,color:A.purple,level:2},
      {id:'ma',label:'Mobile App',sub:'Product Line',x:430,y:180,w:110,color:A.purple,level:2},
      {id:'bi',label:'BI & Reporting',sub:'Product Line',x:560,y:180,w:120,color:A.orange,level:2},
      {id:'dp',label:'Data Platform',sub:'Product Line',x:700,y:180,w:120,color:A.orange,level:2},
      {id:'ts',label:'Team (8-10)',x:10,y:255,w:90,color:A.dim,level:3},
      {id:'ts2',label:'Team (8-10)',x:120,y:255,w:90,color:A.dim,level:3},
      {id:'ts3',label:'Team (8-10)',x:290,y:255,w:90,color:A.dim,level:3},
      {id:'ts4',label:'Team (8-10)',x:400,y:255,w:90,color:A.dim,level:3},
      {id:'ts5',label:'Team (8-10)',x:570,y:255,w:90,color:A.dim,level:3},
      {id:'ts6',label:'Team (8-10)',x:680,y:255,w:90,color:A.dim,level:3},
      {id:'pf',label:'Foundational Platform: API Gateway, Infrastructure, Security, DevOps, Shared Services',x:80,y:330,w:660,color:A.dim,level:4},
      {id:'brs',label:'BRS (Pilot OPCO)',sub:'First consumer',x:300,y:400,w:160,color:A.green,level:5},
      {id:'cx',label:'CX & Digital CoE',sub:'Wave 2',x:500,y:400,w:150,color:A.dim,level:5},
      {id:'net',label:'Networks',sub:'Wave 2',x:130,y:400,w:110,color:A.dim,level:5},
    ],
    edges:[['vp','sap'],['vp','dig'],['vp','data'],['sap','sf'],['sap','ss'],['dig','cp'],['dig','ma'],['data','bi'],['data','dp'],['sf','ts'],['ss','ts2'],['cp','ts3'],['ma','ts4'],['bi','ts5'],['dp','ts6']]
  },
]

function OrgChart({org}:{org:OrgData}) {
  const maxX = Math.max(...org.nodes.map(n=>n.x+n.w)) + 40
  const maxY = Math.max(...org.nodes.map(n=>n.y+(n.h||35))) + 50
  return (
    <svg viewBox={`0 0 ${maxX} ${maxY}`} style={{width:'100%',display:'block'}}>
      {/* Edges */}
      {org.edges.map(([from,to],i)=>{
        const f=org.nodes.find(n=>n.id===from); const t=org.nodes.find(n=>n.id===to)
        if(!f||!t) return null
        return <line key={i} x1={f.x+f.w/2} y1={f.y+(f.h||35)} x2={t.x+t.w/2} y2={t.y}
          stroke={`${org.c}30`} strokeWidth={1.5} strokeDasharray="4 3" />
      })}
      {/* Platform connections (dashed to bottom) */}
      {org.nodes.filter(n=>n.level===4).length > 0 && org.nodes.filter(n=>n.level===3).map((n,i)=>{
        const pf = org.nodes.find(n=>n.level===4)
        if(!pf) return null
        return <line key={`pf${i}`} x1={n.x+n.w/2} y1={n.y+(n.h||35)} x2={pf.x+pf.w/2} y2={pf.y}
          stroke={`${org.c}15`} strokeWidth={1} strokeDasharray="2 4" />
      })}
      {/* Nodes */}
      {org.nodes.map(node=>(
        <g key={node.id}>
          <rect x={node.x} y={node.y} width={node.w} height={node.h||35} rx={8}
            fill={node.level>=4?`${node.color}08`:`${node.color}12`}
            stroke={`${node.color}25`} strokeWidth={1} />
          <text x={node.x+node.w/2} y={node.y+(node.sub?14:20)} textAnchor="middle"
            fill={node.level<2?'#fff':node.color} fontSize={node.level<1?13:node.level<2?12:11}
            fontWeight={node.level<2?700:500} fontFamily="Inter,system-ui">
            {node.label}
          </text>
          {node.sub && <text x={node.x+node.w/2} y={node.y+28} textAnchor="middle"
            fill={A.dim} fontSize={9} fontFamily="Inter,system-ui">
            {node.sub}
          </text>}
        </g>
      ))}
      {/* Gartner attribution */}
      <text x={10} y={maxY-10} fill={A.dim} fontSize={8} fontFamily="Inter">
        Based on Gartner Product-Centric Delivery Organizational Structures (2024)
      </text>
    </svg>
  )
}

function Structures({go}:{go:(v:View)=>void}) {
  const [sel,setSel]=useState(0);const org=ORGS[sel]
  return (
    <div className="scene" style={{padding:'32px'}}><Back go={go}/>
      <div style={{maxWidth:1100,margin:'0 auto',paddingTop:32}}>
        <p className="mono" style={{fontSize:11,color:A.dim,letterSpacing:2}}>GARTNER RESEARCH 2024</p>
        <h2 style={{fontSize:32,fontWeight:800,color:A.text,marginBottom:20}}>Product-Centric Org Structures</h2>
        <div style={{display:'flex',gap:6,marginBottom:24,overflowX:'auto'}}>{ORGS.map((o,i)=><button key={i} onClick={()=>setSel(i)} style={{padding:'10px 18px',borderRadius:10,fontSize:14,fontWeight:sel===i?700:500,whiteSpace:'nowrap',background:sel===i?`${o.c}12`:'transparent',color:sel===i?A.text:A.sub,borderBottom:sel===i?`2px solid ${o.c}`:'2px solid transparent'}}>{o.name}</button>)}</div>
        <AnimatePresence mode="wait"><motion.div key={sel} initial={{opacity:0,y:10}} animate={{opacity:1,y:0}} transition={{duration:.25}}>
          {/* Header */}
          <div className="card" style={{marginBottom:16,borderTop:`3px solid ${org.c}`}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
              <div><h3 style={{fontSize:24,fontWeight:800,color:org.c,margin:0}}>{org.name}</h3><p style={{fontSize:15,color:A.sub,margin:'4px 0 0'}}>{org.ind}</p></div>
              <div style={{display:'flex',gap:8}}>
                {org.rev!=='—'&&<span className="mono" style={{fontSize:13,color:A.dim,padding:'4px 12px',borderRadius:8,background:'rgba(255,255,255,.04)'}}>{org.rev}</span>}
                {org.employees&&<span className="mono" style={{fontSize:13,color:A.dim,padding:'4px 12px',borderRadius:8,background:'rgba(255,255,255,.04)'}}>{org.employees}</span>}
                <span className="mono" style={{fontSize:13,color:org.c,padding:'4px 12px',borderRadius:8,background:`${org.c}10`}}>{org.pct} product</span>
              </div>
            </div>
          </div>
          {/* Visual Org Chart */}
          <div className="card" style={{marginBottom:16,padding:24}}>
            <p style={{fontSize:14,fontWeight:700,color:org.c,letterSpacing:1,textTransform:'uppercase',marginBottom:14}}>Organizational Structure</p>
            <OrgChart org={org} />
          </div>
          {/* Team + Unique */}
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16}}>
            <div className="card"><p style={{fontSize:14,fontWeight:700,color:org.c,letterSpacing:1,textTransform:'uppercase',marginBottom:10}}>Team Composition</p><p style={{fontSize:15,color:A.sub,lineHeight:1.7}}>{org.team}</p></div>
            <div className="card"><p style={{fontSize:14,fontWeight:700,color:org.c,letterSpacing:1,textTransform:'uppercase',marginBottom:10}}>What Makes Them Unique</p><p style={{fontSize:15,color:A.sub,lineHeight:1.7}}>{org.uniq}</p></div>
          </div>
        </motion.div></AnimatePresence>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════
// APP
// ═══════════════════════════════════════
export default function App() {
  const [v,setV]=useState<View>('home')
  return <AnimatePresence mode="wait">
    {v==='home'&&<Home key="h" go={setV}/>}
    {v==='explore'&&<Explore key="e" go={setV}/>}
    {v==='context'&&<Context key="c" go={setV}/>}
    {v==='assess'&&<Assess key="a" go={setV}/>}
    {v==='structures'&&<Structures key="s" go={setV}/>}
  </AnimatePresence>
}
