import * as React from "react";
import { BorderBeam } from "border-beam";
import { useReducedMotion } from "motion/react";
import { Check, CornerDownLeft, Sparkles, X } from "lucide-react";
import type { XlsxCellEditorOverlayRenderProps } from "@extend-ai/react-xlsx";
import "./formula-editor-demo.css";

const DEMO_FORMULA = "=SUM(10,20,30)";

function generateDemoFormula(signal: AbortSignal): Promise<string> {
  return new Promise((resolve, reject) => {
    if (signal.aborted) { reject(new DOMException("Cancelled", "AbortError")); return; }
    const abort = () => { clearTimeout(timer); reject(new DOMException("Cancelled", "AbortError")); };
    const timer = setTimeout(() => {
      signal.removeEventListener("abort", abort);
      resolve(DEMO_FORMULA);
    }, 650);
    signal.addEventListener("abort", abort, { once: true });
  });
}

export function FormulaEditorDemo({ session, isDark }: { session: XlsxCellEditorOverlayRenderProps; isDark: boolean }) {
  const [phase, setPhase] = React.useState<"suggest" | "typing" | "compose" | "generating">("suggest");
  const [mode, setMode] = React.useState<"review" | "apply">("review");
  const [message, setMessage] = React.useState("");
  const reducedMotion = useReducedMotion();
  React.useEffect(() => { setPhase("suggest"); setMessage(""); }, [session.signal]);
  const active = phase === "compose" || phase === "generating" || session.isPreview;
  const suggesting = phase === "suggest" && !!session.value.trim() && session.value !== session.context.originalValue;
  const showActions = suggesting || active;
  const above = session.rect.top + session.rect.height + (phase === "compose" ? 140 : 100) > window.innerHeight;
  const left = Math.min(0, Math.max(8 - session.rect.left, window.innerWidth - session.rect.left - 292));

  async function confirm() {
    if (phase === "generating") return;
    const prompt = session.value;
    setPhase("generating");
    setMessage("");
    try {
      const formula = await generateDemoFormula(session.signal);
      if (session.signal.aborted) return;
      if (session.getValue() !== prompt) {
        setPhase("compose");
        setMessage("Draft changed. Confirm to generate again.");
        return;
      }
      if (mode === "apply") session.commit(formula);
      else { session.preview(formula); setPhase("compose"); session.focus(); }
    } catch (error) {
      if (session.signal.aborted) return;
      setPhase("compose");
      setMessage(error instanceof Error ? error.message : "Could not generate the formula.");
    }
  }

  return (
    <>
      {active && (
        <BorderBeam
          className="playground-editor-beam"
          size="sm"
          strength={0.7}
          borderRadius={2}
          theme={isDark ? "dark" : "light"}
          active={!reducedMotion}
          style={{ position: "absolute", inset: 0, width: "100%", height: "100%", pointerEvents: "none", outline: reducedMotion ? "2px solid #7280df" : undefined }}
        >
          <div aria-hidden="true" style={{ width: "100%", height: "100%", borderRadius: 2 }} />
        </BorderBeam>
      )}
      {showActions && (
        <div
          className={`playground-editor-actions${isDark ? " dark" : ""}`}
          role="group"
          aria-label="Cell editor actions"
          style={{ position: "absolute", left, ...(above ? { bottom: "calc(100% + 8px)" } : { top: "calc(100% + 8px)" }), pointerEvents: "auto" }}
        >
          {session.isPreview ? (
            <>
              <div className="playground-editor-note">Review the formula in the cell.</div>
              <div className="playground-editor-buttons">
                <button type="button" className="playground-editor-primary" onClick={() => session.commit()}><Check size={14} />Keep</button>
                <button type="button" onClick={() => { session.discardPreview(); setPhase("typing"); }}><X size={14} />Discard</button>
              </div>
            </>
          ) : suggesting ? (
            <>
              <button type="button" className="playground-editor-menu-item" onClick={() => { setPhase("compose"); session.focus(); }}><Sparkles size={14} />Write formula</button>
              <button type="button" className="playground-editor-menu-item" onClick={() => { setPhase("typing"); session.focus(); }}><CornerDownLeft size={14} />Keep typing</button>
            </>
          ) : (
            <>
              <div className="playground-editor-heading"><Sparkles size={14} />Write formula<span>Demo</span></div>
              <div className="playground-editor-note">Always generates <code>{DEMO_FORMULA}</code></div>
              <div className="playground-editor-buttons">
                <select aria-label="After generation" value={mode} disabled={phase === "generating"} onChange={event => setMode(event.target.value as "review" | "apply")}>
                  <option value="review">Review first</option>
                  <option value="apply">Apply directly</option>
                </select>
                <button type="button" className="playground-editor-primary" disabled={phase === "generating"} onClick={() => void confirm()}>{phase === "generating" ? "Generating…" : "Confirm"}</button>
                <button type="button" aria-label="Cancel formula assistance" disabled={phase === "generating"} onClick={() => { setPhase("typing"); session.focus(); }}><X size={14} /></button>
              </div>
              {message && <div className="playground-editor-note" role="status">{message}</div>}
              {phase === "generating" && <span className="sr-only" role="status">Generating demo formula</span>}
            </>
          )}
        </div>
      )}
    </>
  );
}
