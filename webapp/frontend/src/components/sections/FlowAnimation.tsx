import { useState, useRef, useCallback } from "react";
import { motion, useInView } from "framer-motion";
import { Play, RotateCcw } from "lucide-react";
import { projects } from "../../data/projects";
import { flowDefinitions, type FlowNode } from "../../data/flows";
import { FlowTrack } from "../flow/FlowTrack";
import { rgba } from "../../utils/colors";

export function FlowAnimation() {
  const ref = useRef(null);
  const isInView = useInView(ref, { once: true, margin: "-50px" });
  const [hoveredNode, setHoveredNode] = useState<FlowNode | null>(null);
  const [selectedProject, setSelectedProject] = useState<string | null>(null);
  const trackRefs = useRef<Record<string, { reset: () => void; play: () => void }>>({});

  const registerTrack = useCallback((id: string, r: { reset: () => void; play: () => void }) => {
    trackRefs.current[id] = r;
  }, []);

  const runAll = () => Object.values(trackRefs.current).forEach((t) => t.play());
  const runSingle = (id: string) => trackRefs.current[id]?.play();
  const resetAll = () => Object.values(trackRefs.current).forEach((t) => t.reset());

  const activeFlow = selectedProject ? flowDefinitions.find((f) => f.projectId === selectedProject) : null;
  const activeProject = selectedProject ? projects.find((p) => p.id === selectedProject) : null;

  return (
    <section ref={ref} className="py-20 px-4 md:px-12 relative overflow-hidden">
      {/* Header */}
      <motion.div initial={{ opacity: 0, y: 20 }} animate={isInView ? { opacity: 1, y: 0 } : {}} className="text-center mb-10 relative z-10">
        <h2 className="text-3xl md:text-4xl font-bold text-white mb-4">How a Demand Flows</h2>
        <p className="text-gray-400 max-w-2xl mx-auto text-sm leading-relaxed">
          Watch how each project processes a demand request from intake to delivery. Click a project for details on tools, roles, and prioritization.
        </p>
      </motion.div>

      {/* Controls */}
      <motion.div initial={{ opacity: 0 }} animate={isInView ? { opacity: 1 } : {}} transition={{ delay: 0.3 }} className="flex gap-2 justify-center mb-10 flex-wrap relative z-10">
        <button onClick={runAll} className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-sm font-semibold transition-all" style={{ backgroundColor: "rgba(25,163,252,0.12)", color: "#19A3FC", border: "1px solid rgba(25,163,252,0.25)" }}>
          <Play size={15} /> Run All
        </button>
        {projects.map((p) => (
          <button key={p.id} onClick={() => { runSingle(p.id); setSelectedProject(p.id); }}
            className="px-4 py-2.5 rounded-xl text-sm font-medium transition-all"
            style={{ backgroundColor: rgba(p.color, 0.1), border: `1px solid ${rgba(p.color, 0.25)}`, color: p.color, outline: selectedProject === p.id ? `2px solid ${p.color}` : "none", outlineOffset: 2 }}>
            {p.name}
          </button>
        ))}
        <button onClick={() => { resetAll(); setSelectedProject(null); }} className="flex items-center gap-2 px-4 py-2.5 rounded-xl text-sm transition-all" style={{ backgroundColor: "rgba(255,255,255,0.04)", color: "#9CA3AF", border: "1px solid rgba(255,255,255,0.08)" }}>
          <RotateCcw size={14} /> Reset
        </button>
      </motion.div>

      {/* Main: tracks + detail */}
      <div className="max-w-[1400px] mx-auto relative z-10 flex gap-6">
        <div className="flex-1 flex flex-col gap-3 min-w-0">
          {flowDefinitions.map((flow, i) => {
            const project = projects.find((p) => p.id === flow.projectId)!;
            return (
              <motion.div key={flow.projectId} initial={{ opacity: 0, x: -20 }} animate={isInView ? { opacity: 1, x: 0 } : {}} transition={{ delay: 0.2 + i * 0.1 }}>
                <FlowTrack flow={flow} project={project} onRegister={registerTrack} onNodeHover={setHoveredNode} />
              </motion.div>
            );
          })}
        </div>

        {/* Right panel */}
        <div className="w-[340px] shrink-0 hidden lg:block">
          {activeFlow && activeProject ? (
            <motion.div key={activeFlow.projectId} initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }}
              className="rounded-xl p-5 sticky top-8" style={{ background: rgba(activeProject.color, 0.05), border: `1px solid ${rgba(activeProject.color, 0.15)}` }}>
              <h3 className="font-bold text-white text-sm mb-1">{activeProject.name}</h3>
              <div className="h-0.5 w-10 rounded mb-3" style={{ backgroundColor: activeProject.color }} />
              <p className="text-gray-400 text-xs leading-relaxed mb-4">{activeFlow.philosophy}</p>

              <div className="mb-4">
                <h4 className="text-gray-500 uppercase tracking-wider font-semibold mb-2" style={{ fontSize: 10 }}>Tool Landscape</h4>
                <div className="space-y-1.5">
                  {(["entry", "execution", "output"] as const).map((key) => (
                    <div key={key}>
                      <span className="font-medium capitalize" style={{ fontSize: 10, color: key === "entry" ? "#EF4444" : key === "execution" ? "#19A3FC" : "#00B050" }}>{key}: </span>
                      <span className="text-gray-400" style={{ fontSize: 10 }}>{activeFlow.toolLandscape[key].join(", ")}</span>
                    </div>
                  ))}
                </div>
              </div>

              <div>
                <h4 className="text-gray-500 uppercase tracking-wider font-semibold mb-2" style={{ fontSize: 10 }}>Prioritization Method</h4>
                <p className="text-gray-400 leading-relaxed" style={{ fontSize: 10 }}>{activeFlow.prioritizationMethod}</p>
              </div>
            </motion.div>
          ) : (
            <div className="rounded-xl p-5 text-center" style={{ border: "1px solid rgba(255,255,255,0.05)", background: "rgba(255,255,255,0.02)" }}>
              <p className="text-gray-600 text-xs">Click a project to see tools, roles, and prioritization details.</p>
            </div>
          )}
        </div>
      </div>

      {/* Hover tooltip */}
      {hoveredNode && (
        <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} className="fixed bottom-6 left-1/2 -translate-x-1/2 z-50 w-full max-w-xl px-4">
          <div className="rounded-xl p-4 shadow-2xl" style={{ background: "#141822", border: "1px solid rgba(255,255,255,0.1)" }}>
            <div className="flex items-center justify-between mb-2">
              <h4 className="text-white font-bold text-sm">{hoveredNode.label}</h4>
              <span className="px-2 py-0.5 rounded-full" style={{ fontSize: 10, backgroundColor: "rgba(25,163,252,0.1)", color: "#19A3FC" }}>{hoveredNode.owner}</span>
            </div>
            <p className="text-gray-400 text-xs leading-relaxed mb-3">{hoveredNode.description}</p>
            {hoveredNode.criteria && (
              <p className="text-xs mb-3"><span className="text-yellow-500 font-semibold">Gate Criteria: </span><span className="text-gray-300">{hoveredNode.criteria}</span></p>
            )}
            <div className="grid grid-cols-3 gap-3">
              {hoveredNode.roles && hoveredNode.roles.length > 0 && (
                <div>
                  <p className="text-gray-500 uppercase tracking-wider font-semibold mb-1" style={{ fontSize: 10 }}>Roles</p>
                  {hoveredNode.roles.map((r) => (
                    <p key={r.role} className="text-gray-400 mb-0.5" style={{ fontSize: 10 }}><span className="text-white font-medium">{r.role}: </span>{r.action}</p>
                  ))}
                </div>
              )}
              {hoveredNode.tools && hoveredNode.tools.length > 0 && (
                <div>
                  <p className="text-gray-500 uppercase tracking-wider font-semibold mb-1" style={{ fontSize: 10 }}>Tools</p>
                  <div className="flex gap-1 flex-wrap">
                    {hoveredNode.tools.map((t) => (
                      <span key={t} className="rounded" style={{ fontSize: 9, padding: "2px 6px", backgroundColor: "rgba(255,255,255,0.05)", color: "#D1D5DB" }}>{t}</span>
                    ))}
                  </div>
                </div>
              )}
              {hoveredNode.artifacts && hoveredNode.artifacts.length > 0 && (
                <div>
                  <p className="text-gray-500 uppercase tracking-wider font-semibold mb-1" style={{ fontSize: 10 }}>Artifacts</p>
                  <div className="flex gap-1 flex-wrap">
                    {hoveredNode.artifacts.map((a) => (
                      <span key={a} className="rounded" style={{ fontSize: 9, padding: "2px 6px", backgroundColor: "rgba(0,176,80,0.08)", color: "rgba(0,176,80,0.8)" }}>{a}</span>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </motion.div>
      )}

      {/* Complexity bar */}
      <motion.div initial={{ opacity: 0 }} animate={isInView ? { opacity: 1 } : {}} transition={{ delay: 0.8 }} className="mt-10 max-w-[1000px] mx-auto w-full relative z-10">
        <p className="text-gray-500 mb-3 text-center uppercase tracking-wider" style={{ fontSize: 10 }}>Relative Process Complexity</p>
        <div className="flex gap-6 items-end justify-center">
          {flowDefinitions.map((flow) => {
            const project = projects.find((p) => p.id === flow.projectId)!;
            const maxN = Math.max(...flowDefinitions.map((f) => f.nodes.length));
            const w = (flow.nodes.length / maxN) * 150;
            return (
              <div key={flow.projectId} className="flex flex-col items-center gap-1.5">
                <motion.div initial={{ width: 0 }} animate={isInView ? { width: w } : {}} transition={{ delay: 1, duration: 0.8 }} className="h-3 rounded-full" style={{ backgroundColor: project.color, minWidth: 24 }} />
                <span className="text-gray-400 font-medium" style={{ fontSize: 10 }}>{project.name}</span>
                <span className="text-gray-600" style={{ fontSize: 10 }}>{flow.nodes.length} stages</span>
              </div>
            );
          })}
        </div>
      </motion.div>
    </section>
  );
}
