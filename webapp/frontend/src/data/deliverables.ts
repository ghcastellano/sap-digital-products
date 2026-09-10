export interface Deliverable {
  name: string;
  shortName: string;
  coverage: { ps: number; coe: number; infra: number; dto: number };
  reuse: {
    action: "ADOPT" | "ADAPT" | "BUILD NEW" | "COMBINE";
    from: string;
    reason: string;
  };
}

export const deliverables: Deliverable[] = [
  {
    name: "E2E Operating Model",
    shortName: "E2E Model",
    coverage: { ps: 1.0, coe: 0.75, infra: 0.25, dto: 0.75 },
    reuse: {
      action: "COMBINE",
      from: "P&S lifecycle + CoE feedback loops + Infra rollout + DTO Kanban",
      reason: "P&S is the only lifecycle in production. CoE adds discovery rigor. Infra provides the rollout methodology. DTO adds pull-based execution. Digital Products needs all four combined.",
    },
  },
  {
    name: "Value & Prioritization",
    shortName: "Prioritization",
    coverage: { ps: 0.75, coe: 0.5, infra: 0, dto: 0.75 },
    reuse: {
      action: "ADAPT",
      from: "P&S calculator + CoE HEVA + DTO continuous model",
      reason: "The PBI formula (BV+UX)/E is already used by P&S and CoE — organizational familiarity means no learning curve. DTO's continuous prioritization eliminates batch bottlenecks. Capacity allocation (Build/BAU/Upgrades) is essential for balancing maintenance with new capabilities.",
    },
  },
  {
    name: "Demand Typology & Intake",
    shortName: "Intake",
    coverage: { ps: 1.0, coe: 0.25, infra: 0, dto: 1.0 },
    reuse: {
      action: "ADOPT",
      from: "DTO Jira-only intake + P&S demand funnel + CoE LBC gate",
      reason: "DTO's Jira-only intake directly solves Digital Products' #1 pain point: fragmented channels (email, Teams, ServiceNow). P&S adds the strategic governance layer. CoE's Lean Business Case forces value justification before work enters the system.",
    },
  },
  {
    name: "Roles & RACI",
    shortName: "RACI",
    coverage: { ps: 1.0, coe: 0.25, infra: 0.25, dto: 0.5 },
    reuse: {
      action: "ADOPT",
      from: "PM Playbook RACI + DTO per-stage ownership + Avangrid Team Topologies",
      reason: "PM Playbook has the only formalized RACI — the proposal explicitly requires 'Roles & RACI Matrix per phase'. DTO adds per-workflow-stage ownership for Kanban execution. Avangrid's Team Topologies defines Digital Products as a Platform Team providing X-as-a-Service.",
    },
  },
  {
    name: "Tool Ecosystem & Jira",
    shortName: "Jira & Tools",
    coverage: { ps: 0.75, coe: 1.0, infra: 0.5, dto: 1.0 },
    reuse: {
      action: "ADOPT",
      from: "CoE JIRA Playbook (41p) + DTO 2-level board design",
      reason: "CoE's 41-page JIRA Playbook is the most comprehensive Jira documentation across Avangrid. DTO's 2-level board (Epic Kanban + Project Kanban) gives Digital Products both portfolio visibility and execution tracking. PM Playbook's issue hierarchy (Initiative → Epic → Story → Task) is the Avangrid standard.",
    },
  },
  {
    name: "Portfolio & PPMNow",
    shortName: "Portfolio",
    coverage: { ps: 0.75, coe: 0.75, infra: 0.5, dto: 0.5 },
    reuse: {
      action: "ADAPT",
      from: "nWOW Portfolio Kanban + CoE Learn Cycle + Avangrid investment model",
      reason: "nWOW's Portfolio Kanban provides the clearest visualization. CoE's quarterly Learn Cycle ensures strategic alignment — critical for a new area. Avangrid's per-unit-served model solves the OPCO cost allocation challenge that PPMNow needs to address.",
    },
  },
  {
    name: "Baseline & Demand Map",
    shortName: "Baseline",
    coverage: { ps: 0.25, coe: 0.25, infra: 0.25, dto: 0.75 },
    reuse: {
      action: "ADOPT",
      from: "DTO baseline format + PD Playbook discovery methodology",
      reason: "The DTO deliverable IS a baseline — its executive summary + findings-per-pillar structure is exactly what the proposal's 'Baseline Report' requires. PD Playbook provides the interview and assessment methodology to execute the baseline for Digital Products.",
    },
  },
  {
    name: "Pilot Execution",
    shortName: "Pilot",
    coverage: { ps: 0.25, coe: 0.25, infra: 0.5, dto: 0.75 },
    reuse: {
      action: "ADOPT",
      from: "DTO pilot design + Avangrid Pilot→Expansion→Optimization",
      reason: "DTO's pilot is the most operationally ready with defined cadences, flow metrics, and pull-based execution — exactly what the BRS pilot needs. Avangrid's methodology was designed for scaling from 1 team to 7+ teams, which is how Digital Products will expand from BRS to all OPCOs.",
    },
  },
  {
    name: "Ways of Working Playbook",
    shortName: "Playbook",
    coverage: { ps: 1.0, coe: 0.75, infra: 0, dto: 0.5 },
    reuse: {
      action: "ADAPT",
      from: "P&S modular structure + CoE agile practices + DTO Kanban guide",
      reason: "P&S has 264 slides organized by lifecycle phase — proven modular structure. CoE teaches 'how to be agile' with 111 slides of fundamentals. DTO adds step-by-step Kanban process instructions. Total: 375+ slides of reference content to adapt, not build from scratch.",
    },
  },
  {
    name: "KPI & Reporting",
    shortName: "KPIs",
    coverage: { ps: 1.0, coe: 0.75, infra: 0, dto: 0.5 },
    reuse: {
      action: "ADAPT",
      from: "PM Playbook multi-level KPIs + CoE strategic metrics + DTO flow metrics",
      reason: "PM Playbook has the only fully defined KPI hierarchy with Jira-automated formulas. CoE adds strategic metrics (Initiative Lead Time, ROI per Initiative) for portfolio visibility. DTO's Flow Metrics (Lead Time, Cycle Time, Throughput, Aging) are the operational heartbeat of Kanban execution.",
    },
  },
];
