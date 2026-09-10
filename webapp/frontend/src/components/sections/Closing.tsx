import { motion, useInView } from "framer-motion";
import { useRef } from "react";

export function Closing() {
  const ref = useRef(null);
  const isInView = useInView(ref, { once: true });

  return (
    <section ref={ref} className="min-h-screen flex flex-col items-center justify-center bg-gradient-animated relative">
      <div className="absolute inset-0 overflow-hidden pointer-events-none">
        <div className="absolute top-1/3 left-1/3 w-96 h-96 bg-blue-primary/5 rounded-full blur-[120px]" />
      </div>

      <motion.div
        initial={{ opacity: 0 }}
        animate={isInView ? { opacity: 1 } : {}}
        transition={{ duration: 1 }}
        className="relative z-10 text-center"
      >
        <motion.div
          initial={{ width: 0 }}
          animate={isInView ? { width: 160 } : {}}
          transition={{ duration: 0.8 }}
          className="h-0.5 bg-blue-accent mx-auto mb-10"
        />

        <h2 className="text-4xl md:text-5xl font-bold text-white mb-6 leading-tight max-w-4xl">
          4 projects. 375+ slides of proven content.
          <br />
          <span className="text-blue-accent">14 weeks</span> to build the Digital Products operating model.
        </h2>

        <p className="text-gray-500 text-lg mt-8">
          NTT DATA × Avangrid — Wave 1: Review & Baseline
        </p>
      </motion.div>
    </section>
  );
}
