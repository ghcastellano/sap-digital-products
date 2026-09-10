import { useRef, useState } from "react";
import { motion, useInView } from "framer-motion";
import {
  Radar,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  ResponsiveContainer,
  Tooltip,
} from "recharts";
import { deliverables } from "../../data/deliverables";
import { projects } from "../../data/projects";

export function RadarCoverage() {
  const ref = useRef(null);
  const isInView = useInView(ref, { once: true, margin: "-100px" });
  const [selectedDeliverable, setSelectedDeliverable] = useState<number | null>(null);

  const radarData = deliverables.map((d) => ({
    subject: d.shortName,
    ps: d.coverage.ps * 100,
    coe: d.coverage.coe * 100,
    infra: d.coverage.infra * 100,
    dto: d.coverage.dto * 100,
    fullMark: 100,
  }));

  const projectKeys = [
    { key: "ps", project: projects[0] },
    { key: "coe", project: projects[1] },
    { key: "infra", project: projects[2] },
    { key: "dto", project: projects[3] },
  ];

  const selected = selectedDeliverable !== null ? deliverables[selectedDeliverable] : null;

  return (
    <section ref={ref} className="min-h-screen flex flex-col px-8 py-20">
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={isInView ? { opacity: 1, y: 0 } : {}}
        className="text-center mb-10"
      >
        <h2 className="text-3xl font-bold text-white mb-3">What Each Project Brings</h2>
        <p className="text-gray-400 max-w-2xl mx-auto">
          Radar view shows coverage across all 10 deliverables. Click any axis to see details and reuse recommendations.
        </p>
      </motion.div>

      <div className="flex gap-8 max-w-7xl mx-auto w-full flex-1 items-start">
        {/* Radar Chart */}
        <motion.div
          initial={{ opacity: 0, scale: 0.9 }}
          animate={isInView ? { opacity: 1, scale: 1 } : {}}
          transition={{ delay: 0.3, duration: 0.6 }}
          className="flex-1 glass p-6"
        >
          <ResponsiveContainer width="100%" height={500}>
            <RadarChart data={radarData} onClick={(e: any) => {
              if (e?.activeTooltipIndex !== undefined) {
                setSelectedDeliverable(e.activeTooltipIndex);
              }
            }}>
              <PolarGrid stroke="rgba(255,255,255,0.06)" />
              <PolarAngleAxis
                dataKey="subject"
                tick={{ fill: "#94A3B8", fontSize: 11 }}
              />
              <PolarRadiusAxis
                angle={90}
                domain={[0, 100]}
                tick={{ fill: "#4B5563", fontSize: 9 }}
                axisLine={false}
              />
              {projectKeys.map(({ key, project }) => (
                <Radar
                  key={key}
                  name={project.name}
                  dataKey={key}
                  stroke={project.color}
                  fill={project.color}
                  fillOpacity={0.08}
                  strokeWidth={2}
                />
              ))}
              <Tooltip
                contentStyle={{
                  backgroundColor: "#1a1f2e",
                  border: "1px solid rgba(255,255,255,0.1)",
                  borderRadius: 8,
                  fontSize: 12,
                }}
                itemStyle={{ color: "#e2e8f0" }}
                labelStyle={{ color: "#fff", fontWeight: 600 }}
                formatter={(value: number) => `${value}%`}
              />
            </RadarChart>
          </ResponsiveContainer>

          {/* Legend */}
          <div className="flex gap-6 justify-center mt-4">
            {projectKeys.map(({ project }) => (
              <div key={project.id} className="flex items-center gap-2">
                <div className="w-3 h-3 rounded-full" style={{ backgroundColor: project.color }} />
                <span className="text-xs text-gray-400">{project.name}</span>
              </div>
            ))}
          </div>
        </motion.div>

        {/* Detail panel */}
        <motion.div
          initial={{ opacity: 0, x: 20 }}
          animate={isInView ? { opacity: 1, x: 0 } : {}}
          transition={{ delay: 0.5 }}
          className="w-96 shrink-0"
        >
          {selected ? (
            <div className="glass p-6" style={{ borderColor: "rgba(25,163,252,0.15)" }}>
              <h3 className="text-lg font-bold text-white mb-2">{selected.name}</h3>

              {/* Coverage bars */}
              <div className="space-y-2 mb-4">
                {projectKeys.map(({ key, project }) => {
                  const val = selected.coverage[key as keyof typeof selected.coverage];
                  return (
                    <div key={key}>
                      <div className="flex justify-between text-xs mb-0.5">
                        <span className="text-gray-400">{project.name}</span>
                        <span style={{ color: project.color }}>{Math.round(val * 100)}%</span>
                      </div>
                      <div className="h-1.5 bg-white/5 rounded-full overflow-hidden">
                        <div
                          className="h-full rounded-full transition-all duration-500"
                          style={{ width: `${val * 100}%`, backgroundColor: project.color }}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Reuse recommendation */}
              <div className="border-t border-white/5 pt-4">
                <div className="flex items-center gap-2 mb-2">
                  <span
                    className="text-xs px-2 py-0.5 rounded-full font-semibold"
                    style={{
                      backgroundColor:
                        selected.reuse.action === "ADOPT" ? "#00B05020" :
                        selected.reuse.action === "ADAPT" ? "#0070C020" :
                        selected.reuse.action === "BUILD NEW" ? "#D46B0820" : "#7C3AED20",
                      color:
                        selected.reuse.action === "ADOPT" ? "#00B050" :
                        selected.reuse.action === "ADAPT" ? "#0070C0" :
                        selected.reuse.action === "BUILD NEW" ? "#D46B08" : "#7C3AED",
                    }}
                  >
                    {selected.reuse.action}
                  </span>
                  <span className="text-xs text-gray-500">for Digital Products</span>
                </div>
                <p className="text-xs text-gray-300 mb-2"><strong>From:</strong> {selected.reuse.from}</p>
                <p className="text-xs text-gray-400 leading-relaxed">{selected.reuse.reason}</p>
              </div>
            </div>
          ) : (
            <div className="glass p-6 text-center">
              <p className="text-gray-500 text-sm">Click any axis on the radar chart to see deliverable details and reuse recommendations.</p>
            </div>
          )}

          {/* Deliverable list */}
          <div className="mt-4 space-y-1">
            {deliverables.map((d, i) => (
              <button
                key={d.name}
                onClick={() => setSelectedDeliverable(i)}
                className={`w-full text-left px-3 py-2 rounded-lg text-xs transition-all ${
                  selectedDeliverable === i
                    ? "bg-white/10 text-white"
                    : "text-gray-500 hover:bg-white/5 hover:text-gray-300"
                }`}
              >
                <span className="font-medium">{d.name}</span>
                <span
                  className="ml-2 text-[10px] px-1.5 py-0.5 rounded-full"
                  style={{
                    backgroundColor:
                      d.reuse.action === "ADOPT" ? "#00B05015" :
                      d.reuse.action === "ADAPT" ? "#0070C015" :
                      d.reuse.action === "BUILD NEW" ? "#D46B0815" : "#7C3AED15",
                    color:
                      d.reuse.action === "ADOPT" ? "#00B050" :
                      d.reuse.action === "ADAPT" ? "#0070C0" :
                      d.reuse.action === "BUILD NEW" ? "#D46B08" : "#7C3AED",
                  }}
                >
                  {d.reuse.action}
                </span>
              </button>
            ))}
          </div>
        </motion.div>
      </div>
    </section>
  );
}
