import { useRef } from "react";
import { motion, useInView } from "framer-motion";

const steps = [
  {
    number: "01",
    title: "Review & Baseline",
    weeks: "Weeks 1-2",
    color: "#19A3FC",
    deliverables: [
      { name: "Baseline Report", action: "ADOPT" as const },
      { name: "Demand Map", action: "ADOPT" as const },
      { name: "Portfolio Landscape", action: "BUILD NEW" as const },
    ],
    description: "Where are we starting from? Current state diagnosis using DTO baseline format + PD Playbook discovery methodology.",
  },
  {
    number: "02",
    title: "Define Lifecycle",
    weeks: "Weeks 2-5",
    color: "#0070C0",
    deliverables: [
      { name: "E2E Lifecycle Map", action: "COMBINE" as const },
      { name: "Demand Typology", action: "ADOPT" as const },
      { name: "Prioritization Framework", action: "ADAPT" as const },
      { name: "Roles & RACI", action: "ADOPT" as const },
      { name: "Value Criteria v1.0", action: "ADAPT" as const },
    ],
    description: "How should we work? Combining P&S lifecycle, CoE feedback loops, DTO intake, and Avangrid Team Topologies.",
  },
  {
    number: "03",
    title: "Design Tool Ecosystem",
    weeks: "Weeks 4-7",
    color: "#002060",
    deliverables: [
      { name: "Jira Configuration v1.0", action: "ADOPT" as const },
      { name: "Intake Form Design", action: "ADOPT" as const },
      { name: "PPMNow Integration", action: "ADAPT" as const },
    ],
    description: "What supports the process? CoE JIRA Playbook (41p) as config standard + DTO 2-level board design.",
  },
  {
    number: "04",
    title: "Pilot with BRS",
    weeks: "Weeks 7-13",
    color: "#00B050",
    deliverables: [
      { name: "BRS Pilot Execution", action: "ADOPT" as const },
      { name: "Feedback Log", action: "BUILD NEW" as const },
      { name: "Pilot Report", action: "ADOPT" as const },
    ],
    description: "Does it work in reality? DTO pilot design + Avangrid's Pilot \u2192 Expansion \u2192 Optimization methodology.",
  },
  {
    number: "05",
    title: "Learn & Adjust",
    weeks: "Weeks 13-14",
    color: "#D46B08",
    deliverables: [
      { name: "WoW Playbook", action: "ADAPT" as const },
      { name: "KPI Framework", action: "ADAPT" as const },
      { name: "Jira Config v2.0", action: "ADAPT" as const },
    ],
    description: "What needs refinement? 375+ slides of reference content adapted into the Digital Products playbook.",
  },
];

const actionColors: Record<string, { bg: string; text: string }> = {
  ADOPT: { bg: "#00B05018", text: "#00B050" },
  ADAPT: { bg: "#0070C018", text: "#0070C0" },
  "BUILD NEW": { bg: "#D46B0818", text: "#D46B08" },
  COMBINE: { bg: "#7C3AED18", text: "#7C3AED" },
};

export function TimelineWave1() {
  const ref = useRef(null);
  const isInView = useInView(ref, { once: true, margin: "-100px" });

  return (
    <section ref={ref} className="min-h-screen flex flex-col px-8 py-20">
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={isInView ? { opacity: 1, y: 0 } : {}}
        className="text-center mb-12"
      >
        <h2 className="text-3xl font-bold text-white mb-3">The 14-Week Plan</h2>
        <p className="text-gray-400 max-w-2xl mx-auto">
          Wave 1 delivers a complete operating model in 5 steps, reusing proven assets from all 4 projects.
        </p>
      </motion.div>

      {/* Timeline */}
      <div className="max-w-6xl mx-auto w-full relative">
        {/* Vertical line */}
        <div className="absolute left-8 top-0 bottom-0 w-px bg-white/10" />

        {steps.map((step, i) => (
          <motion.div
            key={step.number}
            initial={{ opacity: 0, x: -30 }}
            animate={isInView ? { opacity: 1, x: 0 } : {}}
            transition={{ delay: 0.2 + i * 0.15 }}
            className="relative pl-20 pb-10"
          >
            {/* Circle on timeline */}
            <div
              className="absolute left-5 w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold border-2"
              style={{ borderColor: step.color, backgroundColor: `${step.color}20`, color: step.color }}
            >
              {step.number}
            </div>

            {/* Card */}
            <div className="glass p-5" style={{ borderColor: `${step.color}20` }}>
              <div className="flex items-center gap-3 mb-2">
                <h3 className="text-lg font-bold text-white">{step.title}</h3>
                <span
                  className="text-xs px-2 py-0.5 rounded-full font-medium"
                  style={{ backgroundColor: `${step.color}20`, color: step.color }}
                >
                  {step.weeks}
                </span>
              </div>
              <p className="text-sm text-gray-400 mb-4">{step.description}</p>

              {/* Deliverables */}
              <div className="flex gap-2 flex-wrap">
                {step.deliverables.map((d) => {
                  const ac = actionColors[d.action];
                  return (
                    <div
                      key={d.name}
                      className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs"
                      style={{ backgroundColor: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.06)" }}
                    >
                      <span className="text-gray-300 font-medium">{d.name}</span>
                      <span
                        className="text-[10px] px-1.5 py-0.5 rounded-full font-semibold"
                        style={{ backgroundColor: ac.bg, color: ac.text }}
                      >
                        {d.action}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </motion.div>
        ))}
      </div>
    </section>
  );
}
