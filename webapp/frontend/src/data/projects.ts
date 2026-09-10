export interface Project {
  id: string;
  name: string;
  headline: string;
  subtitle: string;
  color: string;
  colorLight: string;
  maturity: number;
}

export const projects: Project[] = [
  {
    id: "ps",
    name: "P&S / P&T",
    headline: "Process-Driven Execution",
    subtitle: "4-phase lifecycle running in production. Most mature: governance, RACI, KPIs all defined. 264 slides of proven playbooks actively used by teams.",
    color: "#002060",
    colorLight: "#E8EFF7",
    maturity: 0.9,
  },
  {
    id: "coe",
    name: "Digital CoE",
    headline: "Discovery-First Governance",
    subtitle: "Separates Discovery from Delivery with feedback loops. Strongest in portfolio metrics, strategic alignment, and Jira configuration (41-page playbook).",
    color: "#0070C0",
    colorLight: "#E3F2FD",
    maturity: 0.7,
  },
  {
    id: "infra",
    name: "Infrastructure",
    headline: "Team-Centric Deployment",
    subtitle: "Team Topologies + Pilot → Expansion → Optimization. Built for multi-team rollout across OPCOs with coaching cycles and flow metrics.",
    color: "#009688",
    colorLight: "#E0F2F1",
    maturity: 0.5,
  },
  {
    id: "dto",
    name: "Renewables DTO",
    headline: "Lean Kanban Operations",
    subtitle: "4 pillars: Intake → Prioritize → Deliver → Govern. Pull-based, Jira-native, most operationally detailed intake and execution process.",
    color: "#D46B08",
    colorLight: "#FFF8E1",
    maturity: 0.55,
  },
];
