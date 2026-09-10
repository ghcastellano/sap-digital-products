import { motion, useInView } from "framer-motion";
import { useRef } from "react";
import { projects } from "../../data/projects";
import { rgba } from "../../utils/colors";

function MaturityBar({ value, color }: { value: number; color: string }) {
  const ref = useRef(null);
  const isInView = useInView(ref, { once: true });
  return (
    <div ref={ref} className="mt-4">
      <div className="flex justify-between text-xs text-gray-500 mb-1">
        <span>Maturity</span>
        <span>{Math.round(value * 100)}%</span>
      </div>
      <div className="h-1.5 rounded-full overflow-hidden" style={{ background: "rgba(255,255,255,0.05)" }}>
        <motion.div
          initial={{ width: 0 }}
          animate={isInView ? { width: `${value * 100}%` } : {}}
          transition={{ duration: 1.2, ease: "easeOut", delay: 0.3 }}
          className="h-full rounded-full"
          style={{ backgroundColor: color }}
        />
      </div>
    </div>
  );
}

export function ProjectCards() {
  const ref = useRef(null);
  const isInView = useInView(ref, { once: true, margin: "-100px" });
  return (
    <section className="min-h-screen flex flex-col justify-center px-8 py-20">
      <motion.div ref={ref} initial={{ opacity: 0, y: 20 }} animate={isInView ? { opacity: 1, y: 0 } : {}} transition={{ duration: 0.6 }} className="text-center mb-12">
        <h2 className="text-3xl font-bold text-white mb-3">The Four Projects</h2>
        <p className="text-gray-400 max-w-2xl mx-auto">
          Each area solved different parts of the same problem — understanding what was actually built reveals what's directly reusable for Digital Products.
        </p>
      </motion.div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 max-w-6xl mx-auto w-full">
        {projects.map((project, i) => (
          <motion.div
            key={project.id}
            initial={{ opacity: 0, y: 40 }}
            animate={isInView ? { opacity: 1, y: 0 } : {}}
            transition={{ delay: 0.2 + i * 0.15, duration: 0.6 }}
            whileHover={{ scale: 1.02, y: -4 }}
            className="glass p-6 cursor-default group"
            style={{ borderColor: rgba(project.color, 0.12) }}
          >
            <div className="h-1 w-16 rounded-full mb-4 transition-all duration-300 group-hover:w-24" style={{ backgroundColor: project.color }} />
            <div className="flex items-center gap-3 mb-2">
              <h3 className="text-xl font-bold text-white">{project.name}</h3>
              <span className="text-xs px-2 py-0.5 rounded-full font-medium" style={{ backgroundColor: rgba(project.color, 0.12), color: project.color }}>
                {project.headline}
              </span>
            </div>
            <p className="text-sm text-gray-400 leading-relaxed">{project.subtitle}</p>
            <MaturityBar value={project.maturity} color={project.color} />
          </motion.div>
        ))}
      </div>
    </section>
  );
}
