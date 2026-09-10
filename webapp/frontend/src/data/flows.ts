export interface FlowNode {
  id: string;
  label: string;
  type: "stage" | "gate" | "loop";
  owner: string;
  description: string;
  artifacts: string[];
  tools?: string[];
  criteria?: string;
  roles?: { role: string; action: string }[];
}

export interface FlowDefinition {
  projectId: string;
  philosophy: string;
  toolLandscape: { entry: string[]; execution: string[]; output: string[] };
  prioritizationMethod: string;
  nodes: FlowNode[];
}

export const flowDefinitions: FlowDefinition[] = [
  {
    projectId: "ps",
    philosophy: "Process-driven execution with Flight Level governance. Work flows through 4 sequential phases with defined gates, artifacts, and RACI per phase. The most operationally mature model.",
    toolLandscape: {
      entry: ["Email", "ServiceNow", "Teams Meetings"],
      execution: ["Jira (Kanban/Scrum boards)", "Confluence (RACI, docs)", "Figma + Maze (prototyping)"],
      output: ["Jira Dashboards (KPIs)", "PPMNow (Portfolio view)", "Health Scorecards"],
    },
    prioritizationMethod: "PBI(x) = (BV + UX) / E — Weighted scoring calculator with 1–10 scale. Business Value Criteria tied to Avangrid axes: Delight Users, Asset Health, Field Productivity, O&M Reduction, Regulatory, Safety, Environmental. Effort Matrix: S (1–5 features), M (5–10), L (10+). Demand classified as Strategic, Operational, or Regulatory.",
    nodes: [
      {
        id: "ps-1", label: "Area Requests", type: "stage",
        owner: "Requestor / Business Area",
        description: "Demand enters from 3 sources: (1) Strategic initiatives from executives, (2) Product feedback from user research and field observations, (3) Regulatory/compliance mandates. Requests are submitted through multiple channels.",
        artifacts: ["Demand Request Form"],
        tools: ["Email", "ServiceNow", "Teams"],
        roles: [
          { role: "Requestor", action: "Submits demand through available channel" },
          { role: "Portfolio Manager", action: "Receives and logs incoming demand" },
        ],
      },
      {
        id: "ps-2", label: "Shark Tank", type: "gate",
        owner: "Senior Leadership",
        description: "Quarterly governance gate (1–2 hours). Evaluates viability, potential value, and risk profile of initiatives. Grants or denies GO/No-Go. Structured 'pitch' format where PMs present their initiative to leadership panel.",
        artifacts: ["GO/No-Go Decision", "Triage Assessment"],
        tools: ["PowerPoint (pitch)", "Confluence (scoring)"],
        criteria: "Strategic alignment, feasibility, resource availability, risk profile",
        roles: [
          { role: "Product Manager", action: "Presents initiative pitch" },
          { role: "Senior Leadership", action: "Evaluates and votes GO/No-Go" },
          { role: "Portfolio Manager", action: "Facilitates, records decisions" },
        ],
      },
      {
        id: "ps-3", label: "Portfolio Committee", type: "gate",
        owner: "Portfolio Manager",
        description: "Monthly/quarterly investment decision forum. Ensures the organization invests in the right work by selecting and prioritizing initiatives that maximize strategic value. Evaluates portfolio balance using the PBI formula: PBI(x) = (BV+UX)/E.",
        artifacts: ["Portfolio Kanban Entry", "Priority Score"],
        tools: ["Portfolio Kanban (Jira/PPMNow)", "Scoring Calculator"],
        criteria: "PBI formula score, OKR alignment, capacity availability",
        roles: [
          { role: "Portfolio Manager", action: "Maintains portfolio kanban, facilitates prioritization" },
          { role: "Program Managers", action: "Present initiative status and dependencies" },
          { role: "Senior Leadership", action: "Final investment approval" },
        ],
      },
      {
        id: "ps-4", label: "Discovery Phase", type: "stage",
        owner: "Product Manager",
        description: "Design Thinking process: Empathize → Define → Ideate → Prototype → Test. Contextual research, field observations, empathy maps, problem statements. User testing with Figma prototypes via Maze platform. Validation Report with feasibility matrix.",
        artifacts: ["Problem Statement Card", "Research Notes", "Empathy Maps", "Validation Report"],
        tools: ["Figma (prototyping)", "Maze (user testing)", "Confluence (documentation)"],
        roles: [
          { role: "Product Manager", action: "Leads discovery, defines problem statement" },
          { role: "UX Designer", action: "Creates prototypes, runs user tests" },
          { role: "Business Stakeholder", action: "Validates problem relevance" },
        ],
      },
      {
        id: "ps-5", label: "Solution Proposal", type: "stage",
        owner: "Product Manager + PM",
        description: "Solutioning phase: Project Book creation with scope, timeline, budget. Budget Release & Resourcing approval. Cross-functional team assembly. Project plan with milestones and deliverables defined.",
        artifacts: ["Project Book", "Budget Release", "Resource Plan", "Project Charter"],
        tools: ["Confluence (Project Book)", "Jira (project setup)", "PPMNow (budget tracking)"],
        roles: [
          { role: "Product Manager", action: "Defines solution scope and requirements" },
          { role: "Project Manager", action: "Creates Project Book, plans timeline" },
          { role: "Sponsor", action: "Approves budget release" },
        ],
      },
      {
        id: "ps-6", label: "Project Execution", type: "stage",
        owner: "Dev Team (PO / SM / Devs)",
        description: "Kanban-managed execution at Flight Level 1. Workflow: Backlog → In Progress → QA → Done. PoC/MVP → Build → Testing → Deploy. Scrum Master ensures Kanban hygiene, WIP limits, and flow metrics. PM tracks health via Jira dashboards.",
        artifacts: ["Working Software", "Test Reports", "Sprint/Flow Reports"],
        tools: ["Jira (Kanban boards)", "Confluence (technical docs)", "CI/CD pipeline"],
        roles: [
          { role: "Product Owner", action: "Prioritizes backlog, accepts deliverables" },
          { role: "Scrum Master", action: "Facilitates flow, removes blockers, ensures WIP limits" },
          { role: "Dev Team", action: "Pulls work, builds, tests, deploys" },
          { role: "Project Manager", action: "Monitors health dashboards, manages dependencies" },
        ],
      },
      {
        id: "ps-7", label: "Ongoing / Close", type: "stage",
        owner: "Project Manager",
        description: "Continuous improvement and reporting. Multi-level KPIs tracked: Throughput, Lead Time, Velocity (Story Points), Scope Creep, WIP. Health Dashboards and Scorecards with 2026 targets. Governance at Flight Level 2 (Program) and 3 (Portfolio).",
        artifacts: ["Health Dashboard", "Scorecard", "Lessons Learned", "Contracts"],
        tools: ["Jira Dashboards", "PPMNow (portfolio reporting)", "Confluence (retrospectives)"],
        roles: [
          { role: "Project Manager", action: "Closes project, documents lessons learned" },
          { role: "Portfolio Manager", action: "Updates portfolio status, reports to leadership" },
        ],
      },
    ],
  },
  {
    projectId: "coe",
    philosophy: "Discovery-first governance with separate Discovery and Delivery flows connected by feedback loops. No work enters Delivery without passing through Discovery validation first. More framework than playbook.",
    toolLandscape: {
      entry: ["Strategic Initiatives Funnel", "Executive demand"],
      execution: ["Jira (Discovery + Delivery boards)", "Ziflow", "Maze", "Figma", "Monday"],
      output: ["Portfolio Learn Cycle reports", "Sprint Report Canvas", "Agility Health Radars"],
    },
    prioritizationMethod: "PBI(x) = (BV + UX) / E at backlog level (same formula as P&S but applied to features/stories, not portfolio). HEVA Cycle: Hypothesis → Experiment → Validation → Apprenticeship for continuous discovery. COORG method: Classify, Order, Organize for backlog sequencing. Planning Poker for complexity estimation. Importance × Urgency Matrix for strategic initiatives.",
    nodes: [
      {
        id: "coe-1", label: "Executive Idea", type: "stage",
        owner: "VP Digital / Business Owner",
        description: "Strategic initiatives from executive levels. First source of demand, focused on sensitive business changes. Ideas enter the Agile Portfolio Funnel for evaluation.",
        artifacts: ["Initiative Proposal"],
        tools: ["Portfolio Funnel (visual board)"],
        roles: [
          { role: "VP Digital & Customer Care", action: "Proposes strategic initiatives" },
          { role: "Enterprise Architect", action: "Assesses technical feasibility" },
        ],
      },
      {
        id: "coe-2", label: "Lean Biz Case", type: "gate",
        owner: "Portfolio Management Team",
        description: "Lean Business Case (LBC) is a streamlined artifact that evaluates value, feasibility, and alignment. Emphasizes brevity — only essential information for decision-making. Forces value justification before any work enters the system.",
        artifacts: ["Lean Business Case (LBC)", "Value Charter"],
        tools: ["Confluence (LBC template)"],
        criteria: "Business value justification, strategic alignment, feasibility assessment",
        roles: [
          { role: "Product Manager", action: "Prepares Lean Business Case" },
          { role: "Business Owner", action: "Validates business value claims" },
        ],
      },
      {
        id: "coe-3", label: "Shark Tank", type: "gate",
        owner: "Enterprise Architect + PMT",
        description: "Initial triage in the Agile Portfolio Funnel. Shark Tank → Think Tank → Biz Tank sequence. Portfolio Management Team: Enterprise Architect, VP Digital, Enterprise Backlog.",
        artifacts: ["Triage Decision", "PRD (Light Requirements)"],
        tools: ["Portfolio Board"],
        criteria: "Viability, strategic fit, technical feasibility",
        roles: [
          { role: "Enterprise Architect", action: "Evaluates technical viability" },
          { role: "VP Digital", action: "Evaluates strategic alignment" },
          { role: "Product Owner", action: "Presents initiative details" },
        ],
      },
      {
        id: "coe-4", label: "Think Tank", type: "gate",
        owner: "Portfolio Management Team",
        description: "Technical feasibility and design phase assessment. Product Blueprints (Requirements) are refined. Quarterly Program Plan alignment check.",
        artifacts: ["Product Blueprint", "Program Plan alignment"],
        tools: ["Monday (planning)", "Confluence"],
        roles: [
          { role: "Technical Lead", action: "Validates technical approach" },
          { role: "Product Owner", action: "Refines requirements" },
        ],
      },
      {
        id: "coe-5", label: "Discovery Board", type: "stage",
        owner: "Product Owner",
        description: "Structured discovery workflow implemented as a Jira Kanban board. Phases: Opportunity Identification → Ideation → Technical Feasibility → Design → Business Signoff. Measures Cycle Time per phase and Discovery Success Rate. Tools per phase: Ziflow, Maze, Figma, Monday.",
        artifacts: ["Discovery Backlog", "Validated Hypothesis", "User Research"],
        tools: ["Jira (Discovery Board)", "Ziflow", "Maze", "Figma", "Monday"],
        roles: [
          { role: "Product Owner", action: "Moves items through discovery phases" },
          { role: "UX/Research", action: "Conducts user research and testing" },
          { role: "Tech Lead", action: "Validates technical feasibility" },
        ],
      },
      {
        id: "coe-6", label: "Feedback Loop", type: "loop",
        owner: "Product Owner",
        description: "Delivery insights feed back to Discovery. Once features are delivered, real-world feedback and performance data inform future discoveries. This creates continuous interaction in two loops: Discovery feeds Delivery (validated needs become backlog items), Delivery feeds Discovery (real-world feedback refines future discoveries).",
        artifacts: ["Feedback Report", "Performance Data"],
        tools: ["Jira (linked boards)", "Analytics"],
        roles: [
          { role: "Product Owner", action: "Routes feedback to appropriate discovery items" },
          { role: "Agile Master", action: "Facilitates retrospective analysis" },
        ],
      },
      {
        id: "coe-7", label: "Delivery Board", type: "stage",
        owner: "Scrum Master / Agile Master",
        description: "Sprint-based delivery with 2-week time-boxed cycles. Jira Delivery Board: Backlog → In Progress → QA → Done. Sprint Report Canvas tracks team baselines with comparison across cycles. Agile Master ensures hygiene checklist compliance.",
        artifacts: ["Product Increment", "Sprint Report Canvas", "Burndown Chart"],
        tools: ["Jira (Delivery Board)", "Sprint Report Canvas template"],
        roles: [
          { role: "Scrum Master / Agile Master", action: "Facilitates sprints, ensures hygiene" },
          { role: "Dev Team", action: "Delivers product increments" },
          { role: "Product Owner", action: "Accepts/rejects deliverables" },
        ],
      },
      {
        id: "coe-8", label: "Release", type: "stage",
        owner: "Dev Team + PO",
        description: "Deployment to production. Portfolio Learn Cycle quarterly (30 min past, 30 min future). Portfolio Metrics: Initiative Lead Time, ROI per Initiative, Lean Portfolio Radar, % per Strategic Theme. Agility Health Radars for team assessment.",
        artifacts: ["Release Notes", "Portfolio Review", "Health Radar"],
        tools: ["CI/CD", "Jira (release tracking)", "Health Radar template"],
        roles: [
          { role: "Dev Team", action: "Deploys to production" },
          { role: "Product Owner", action: "Validates release against acceptance criteria" },
          { role: "Portfolio Manager", action: "Updates portfolio metrics" },
        ],
      },
    ],
  },
  {
    projectId: "infra",
    philosophy: "Team-centric deployment focused on organizational change, not lifecycle definition. Uses Team Topologies to structure relationships between teams and a phased rollout (Pilot → Expansion → Optimization) with coaching cycles.",
    toolLandscape: {
      entry: ["Business area requests", "Portfolio backlog"],
      execution: ["Jira (Kanban boards)", "Businessmap/LeanKit"],
      output: ["Flow Metrics Reports", "Portfolio Roadmap", "Investment Model"],
    },
    prioritizationMethod: "No formal prioritization formula defined. Portfolio Roadmap drives sequencing. Investment model per unit served determines budget allocation. Blueprint Design method for project scoping.",
    nodes: [
      {
        id: "infra-1", label: "Request", type: "stage",
        owner: "Business Area / TPO",
        description: "Work request enters from TPO (Technology, Projects & Operations) areas. Assessed for alignment with the agile transformation rollout plan. Project Inception: new projects emerge or are approved to the portfolio backlog.",
        artifacts: ["Work Request", "Portfolio Backlog Entry"],
        tools: ["Jira", "Email"],
        roles: [
          { role: "Business Area Lead", action: "Submits transformation request" },
          { role: "Portfolio Manager", action: "Logs in portfolio backlog" },
        ],
      },
      {
        id: "infra-2", label: "Assessment", type: "gate",
        owner: "Enterprise Agile Coach",
        description: "ODAS Model assessment: organizational readiness evaluation. Determine Team Topologies: which of the 4 team types (Stream-Aligned, Enabling, Platform, Complicated Subsystem) and 3 interaction modes (Collaboration, X-as-a-Service, Facilitating) apply.",
        artifacts: ["Assessment Report", "Team Topology Map", "ODAS Plan"],
        tools: ["Assessment templates", "Wardley Maps"],
        criteria: "Org readiness, team maturity, volume/frequency of demand, capability gaps",
        roles: [
          { role: "Enterprise Agile Coach", action: "Conducts ODAS assessment" },
          { role: "Senior Talent Consultant", action: "Evaluates team composition" },
        ],
      },
      {
        id: "infra-3", label: "Pilot (1 team)", type: "stage",
        owner: "Sensei Agile Coach",
        description: "First team implements Kanban method. Close support with team coaching and mentoring. Blueprint Design: step-by-step project scoping for PMs. Sapiens Learning Cycle: Courses, Tutoring, Knowledge practice. Design the team's Kanban system.",
        artifacts: ["Kanban Board", "Team Blueprint", "Training Materials"],
        tools: ["Jira (Kanban board)", "Confluence (Blueprint)"],
        roles: [
          { role: "Sensei Agile Coach", action: "Leads pilot, coaches team daily" },
          { role: "Sempai Coach", action: "Supports coaching, handles logistics" },
          { role: "Team Members", action: "Learn and adopt Kanban practices" },
        ],
      },
      {
        id: "infra-4", label: "Learn & Measure", type: "stage",
        owner: "Kanban Lead / Coach",
        description: "Measure Flow Metrics: Lead Time, Cycle Time, Throughput — the foundational definition of a successful Kanban team. Team Success & Workflow Improvement Metrics tracked. Results collected and reported using Flow Metrics approach.",
        artifacts: ["Flow Metrics Report", "Improvement Backlog"],
        tools: ["Jira Dashboards", "Flow Analytics"],
        roles: [
          { role: "Kanban Lead", action: "Monitors flow metrics, identifies bottlenecks" },
          { role: "Agile Coach", action: "Facilitates retrospectives, captures learnings" },
        ],
      },
      {
        id: "infra-5", label: "Expansion (7+ teams)", type: "stage",
        owner: "Enterprise Agile Coach",
        description: "Sapiens Learning Cycles for Kanban implementation across 7+ additional teams. Reference learnings from pilot. Clear steps of team coaching and support. Design additional team Kanban systems. Coordination via portfolio-level management system.",
        artifacts: ["Expansion Plan", "Training Curriculum", "7+ Team Kanban Systems"],
        tools: ["Jira (multiple boards)", "Training platform"],
        roles: [
          { role: "Enterprise Agile Coach", action: "Coordinates expansion across teams" },
          { role: "Sensei/Sempai Coaches", action: "Coach individual teams" },
          { role: "Team Leads", action: "Champion adoption within their teams" },
        ],
      },
      {
        id: "infra-6", label: "Optimization", type: "stage",
        owner: "Portfolio Manager",
        description: "Stabilization phase. Portfolio Roadmap across value streams. Per-unit-served investment model for OPCO cost allocation. Deployment governance with phased rollout planning and milestone tracking. Continuous improvement embedded.",
        artifacts: ["Portfolio Roadmap", "Investment Model", "Maturity Assessment"],
        tools: ["PPMNow / Portfolio tools", "Investment tracking"],
        roles: [
          { role: "Portfolio Manager", action: "Manages portfolio roadmap and investment" },
          { role: "Enterprise Agile Coach", action: "Ensures sustained adoption" },
        ],
      },
    ],
  },
  {
    projectId: "dto",
    philosophy: "Lean Kanban operations with 4 pillars: Demand Intake → Prioritization → Delivery → Governance. Pull-based execution — work is never pushed, only pulled by available capacity. Jira is the single source of truth with zero tolerance for informal channels.",
    toolLandscape: {
      entry: ["Jira Intake Ticket (ONLY — no email, no chat, no meetings)"],
      execution: ["Jira (2-level: Epic Kanban + Project Kanban)"],
      output: ["Flow Metrics Dashboard (Lead Time, Cycle Time, Throughput, Aging)", "Blocker Tracker"],
    },
    prioritizationMethod: "Continuous backlog prioritization (not batch/sprint-based). Simple scoring: Business Value × Urgency × Capacity. Capacity Allocation model: Build 60%, BAU 25%, Upgrades 15%. Classes of Service: Standard, Expedite, Fixed Date, Intangible. Work types: BAU, Build, Upgrades. Backlog is dynamically ordered — priorities re-evaluated continuously.",
    nodes: [
      {
        id: "dto-1", label: "Jira Intake", type: "stage",
        owner: "Product Owner / Tech Lead",
        description: "ALL requests must be submitted via Jira — no emails, chats, or meetings accepted. Standardized intake form with required fields. Work classified by Class of Service (Standard, Expedite, Fixed Date, Intangible) and Work Type (BAU, Build, Upgrades). Triage: accept, request clarification, or reject. Definition of Ready (DoR) must be met before entering backlog.",
        artifacts: ["Jira Intake Ticket", "Intake Form", "DoR Checklist"],
        tools: ["Jira Service Desk / Kanban Intake Board"],
        criteria: "DoR: Title, description, BV, acceptance criteria, Class of Service, work type all filled",
        roles: [
          { role: "Requestor", action: "Submits Jira ticket with standardized form" },
          { role: "Product Owner / Tech Lead", action: "Triages: accept, clarify, or reject" },
          { role: "Project Manager", action: "Governs intake process compliance" },
        ],
      },
      {
        id: "dto-2", label: "Prioritize", type: "stage",
        owner: "Product Owner",
        description: "Continuous backlog prioritization — not sprint-based or batch-based. Scoring model: Business Value × Urgency × Capacity. Capacity Allocation enforced: Build 60%, BAU 25%, Upgrades 15%. Weekly backlog review session with PO, key stakeholders, and PM. Pre-Estimation: lightweight effort sizing (Hours-Time) during refinement. Backlog serves as single ordered queue for delivery.",
        artifacts: ["Prioritized Backlog", "Capacity Plan", "Priority Framework"],
        tools: ["Jira (Backlog view)", "Scoring template"],
        criteria: "Priorities are objective, not opinion-based. Work types balanced per allocation model.",
        roles: [
          { role: "Product Owner", action: "Orders backlog based on scoring criteria" },
          { role: "Dev Team", action: "Pre-estimates effort during refinement" },
          { role: "Key Stakeholders", action: "Validate business priority alignment" },
          { role: "Project Manager", action: "Facilitates weekly backlog review" },
        ],
      },
      {
        id: "dto-3", label: "Pull & Deliver", type: "stage",
        owner: "Dev Team",
        description: "Pull-based execution — teams pull highest priority item when capacity is available. NO push assignments. WIP limits enforced at each stage. 2-level Jira board: Level 1 Epic Kanban (portfolio: Backlog → PMO → Prep → Ready → Build → Done) with swimlanes by business area (Growth, Power Trading, O&M). Level 2 Project Kanban (execution: Backlog → In Progress → QA/Peer Review → UAT → Done). QA pulls from Dev immediately — continuous testing, not end-of-sprint.",
        artifacts: ["Kanban Board (2 levels)", "Working Software", "DoD Checklist"],
        tools: ["Jira (Epic Kanban + Project Kanban)"],
        roles: [
          { role: "Dev Team", action: "Pulls work, builds, commits. Never exceeds WIP limit" },
          { role: "QA Lead", action: "Pulls from Dev, executes tests. Returns defects immediately" },
          { role: "Product Owner", action: "Validates DoD, accepts deliverables" },
        ],
      },
      {
        id: "dto-4", label: "Governance", type: "gate",
        owner: "Kanban Lead",
        description: "Flow monitoring via Jira Dashboard: Lead Time, Cycle Time, Throughput, Aging. Cadences: Daily Flow Review (blockers, WIP breaches, aging work), Replenishment Meeting (review backlog, confirm priorities), Retrospective/Kaizen (identify inefficiencies, implement improvements). Blocker Tracker for impediment management. Continuous Improvement (Kaizen) embedded — every retrospective produces measurable improvement items.",
        artifacts: ["Flow Metrics Dashboard", "Blocker Tracker", "Improvement Log", "Kaizen Items"],
        tools: ["Jira Dashboard", "Flow Analytics"],
        roles: [
          { role: "Kanban Lead", action: "Monitors flow metrics, facilitates all cadences" },
          { role: "Project Manager", action: "Manages blockers, escalations, expedite requests" },
          { role: "Entire Team", action: "Participates in Kaizen, proposes improvements" },
        ],
      },
    ],
  },
];
