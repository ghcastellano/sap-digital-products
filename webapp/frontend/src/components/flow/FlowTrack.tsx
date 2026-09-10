import { useRef, useEffect, useCallback, useState } from "react";
import gsap from "gsap";
import { MotionPathPlugin } from "gsap/MotionPathPlugin";
import type { FlowDefinition, FlowNode } from "../../data/flows";
import type { Project } from "../../data/projects";
import { rgba } from "../../utils/colors";

gsap.registerPlugin(MotionPathPlugin);

interface Props {
  flow: FlowDefinition;
  project: Project;
  onRegister: (id: string, ref: { reset: () => void; play: () => void }) => void;
  onNodeHover: (node: FlowNode | null) => void;
}

export function FlowTrack({ flow, project, onRegister, onNodeHover }: Props) {
  const tokenRef = useRef<SVGCircleElement>(null);
  const glowRef = useRef<SVGCircleElement>(null);
  const tlRef = useRef<gsap.core.Timeline | null>(null);
  const [activeNodeIndex, setActiveNodeIndex] = useState(-1);

  const nodeCount = flow.nodes.length;
  const svgWidth = 1000;
  const svgHeight = 80;
  const paddingX = 70;
  const nodeSpacing = (svgWidth - paddingX * 2) / Math.max(nodeCount - 1, 1);

  const getNodeX = (i: number) => paddingX + i * nodeSpacing;
  const nodeY = 32;

  const play = useCallback(() => {
    if (tlRef.current) tlRef.current.kill();
    const tl = gsap.timeline({ onComplete: () => setActiveNodeIndex(-1) });
    tlRef.current = tl;

    tl.set([tokenRef.current, glowRef.current], { opacity: 1 });
    tl.set([tokenRef.current, glowRef.current], { attr: { cx: getNodeX(0), cy: nodeY } });

    for (let i = 0; i < nodeCount - 1; i++) {
      const fromX = getNodeX(i);
      const toX = getNodeX(i + 1);
      const isGate = flow.nodes[i + 1].type === "gate";
      const isLoop = flow.nodes[i + 1].type === "loop";

      tl.call(() => setActiveNodeIndex(i), [], "<");

      if (isLoop) {
        const midX = (fromX + toX) / 2;
        tl.to([tokenRef.current, glowRef.current], { duration: 0.35, attr: { cx: midX, cy: nodeY - 22 }, ease: "power2.out" });
        tl.to([tokenRef.current, glowRef.current], { duration: 0.35, attr: { cx: fromX, cy: nodeY }, ease: "power2.in" });
        tl.to([tokenRef.current, glowRef.current], { duration: 0.35, attr: { cx: toX, cy: nodeY }, ease: "power2.inOut" });
      } else {
        tl.to([tokenRef.current, glowRef.current], { duration: 0.5, attr: { cx: toX }, ease: "power2.inOut" });
      }

      if (isGate) {
        tl.call(() => setActiveNodeIndex(i + 1));
        tl.to({}, { duration: 0.25 });
      }
    }

    tl.call(() => setActiveNodeIndex(nodeCount - 1));
    tl.to([tokenRef.current, glowRef.current], { duration: 0.3, opacity: 0 }, "+=0.4");
    tl.play();
  }, [flow, nodeCount]);

  const reset = useCallback(() => {
    if (tlRef.current) tlRef.current.kill();
    setActiveNodeIndex(-1);
    gsap.set([tokenRef.current, glowRef.current], { opacity: 0 });
  }, []);

  useEffect(() => {
    onRegister(flow.projectId, { play, reset });
    gsap.set([tokenRef.current, glowRef.current], { opacity: 0 });
  }, [flow.projectId, onRegister, play, reset]);

  return (
    <div className="flex items-center gap-4">
      {/* Project label */}
      <div className="w-32 shrink-0 text-right pr-1">
        <span className="text-xs font-bold tracking-wide" style={{ color: project.color }}>
          {project.name}
        </span>
      </div>

      {/* SVG Track */}
      <div
        className="flex-1 rounded-xl overflow-visible"
        style={{
          background: `linear-gradient(135deg, ${rgba(project.color, 0.04)}, ${rgba(project.color, 0.02)})`,
          border: `1px solid ${rgba(project.color, 0.1)}`,
          padding: "8px 4px",
        }}
      >
        <svg
          viewBox={`0 0 ${svgWidth} ${svgHeight}`}
          className="w-full"
          style={{ height: 80, overflow: "visible" }}
          preserveAspectRatio="xMidYMid meet"
        >
          {/* Edges */}
          {flow.nodes.slice(0, -1).map((_, i) => {
            const fromX = getNodeX(i);
            const toX = getNodeX(i + 1);
            const nextNode = flow.nodes[i + 1];
            return (
              <g key={`edge-${i}`}>
                <line
                  x1={fromX} y1={nodeY} x2={toX} y2={nodeY}
                  stroke={rgba(project.color, 0.3)}
                  strokeWidth={2}
                  strokeDasharray={nextNode.type === "loop" ? "5 5" : "none"}
                />
                {/* Arrow head */}
                <polygon
                  points={`${toX - 6},${nodeY - 3} ${toX},${nodeY} ${toX - 6},${nodeY + 3}`}
                  fill={rgba(project.color, 0.3)}
                />
                {/* Loop arc */}
                {nextNode.type === "loop" && (
                  <path
                    d={`M ${toX} ${nodeY} Q ${(fromX + toX) / 2} ${nodeY - 30} ${fromX} ${nodeY}`}
                    fill="none"
                    stroke={rgba(project.color, 0.18)}
                    strokeWidth={1.5}
                    strokeDasharray="4 3"
                    markerEnd=""
                  />
                )}
              </g>
            );
          })}

          {/* Nodes */}
          {flow.nodes.map((node, i) => {
            const x = getNodeX(i);
            const isActive = i === activeNodeIndex;
            const isGate = node.type === "gate";
            const isLoop = node.type === "loop";

            return (
              <g
                key={node.id}
                onMouseEnter={() => onNodeHover(node)}
                onMouseLeave={() => onNodeHover(null)}
                style={{ cursor: "pointer" }}
              >
                {/* Active pulse */}
                {isActive && (
                  <circle cx={x} cy={nodeY} r={18} fill="none" stroke={project.color} strokeWidth={1.5} opacity={0.2}>
                    <animate attributeName="r" from="14" to="22" dur="0.8s" repeatCount="indefinite" />
                    <animate attributeName="opacity" from="0.3" to="0" dur="0.8s" repeatCount="indefinite" />
                  </circle>
                )}

                {/* Node shape */}
                {isGate ? (
                  <rect
                    x={x - 9} y={nodeY - 9} width={18} height={18} rx={3}
                    fill={isActive ? project.color : rgba(project.color, 0.15)}
                    stroke={project.color}
                    strokeWidth={1.5}
                    transform={`rotate(45 ${x} ${nodeY})`}
                  />
                ) : (
                  <circle
                    cx={x} cy={nodeY} r={isActive ? 10 : 8}
                    fill={isActive ? project.color : rgba(project.color, 0.08)}
                    stroke={project.color}
                    strokeWidth={isLoop ? 2 : 1.5}
                    strokeDasharray={isLoop ? "3 2" : "none"}
                    style={{ transition: "r 0.2s, fill 0.2s" }}
                  />
                )}

                {/* Inner dot for stages */}
                {!isGate && !isLoop && (
                  <circle cx={x} cy={nodeY} r={3} fill={isActive ? "#fff" : project.color} opacity={isActive ? 1 : 0.5} />
                )}

                {/* Label */}
                <text
                  x={x} y={nodeY + 26}
                  textAnchor="middle"
                  fill={isActive ? "#fff" : "#94A3B8"}
                  fontSize={10}
                  fontFamily="Inter, system-ui"
                  fontWeight={isActive ? 600 : 400}
                  style={{ transition: "fill 0.2s" }}
                >
                  {node.label}
                </text>
              </g>
            );
          })}

          {/* Glow token */}
          <circle ref={glowRef} cx={getNodeX(0)} cy={nodeY} r={14} fill={project.color} opacity={0} style={{ filter: "blur(10px)" }} />
          {/* Main token */}
          <circle ref={tokenRef} cx={getNodeX(0)} cy={nodeY} r={5} fill="#fff" stroke={project.color} strokeWidth={2.5} opacity={0} className="token-trail" />
        </svg>
      </div>
    </div>
  );
}
