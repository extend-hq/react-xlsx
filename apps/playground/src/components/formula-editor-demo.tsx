import * as React from "react";
import { Check, CornerDownLeft, Sparkles, X } from "lucide-react";
import type { XlsxCellEditorOverlayRenderProps } from "@extend-ai/react-xlsx";
import { Button } from "@/components/ui/button";
import { FormulaDemoPopover, formulaActionClassName, getDemoFormula } from "./formula-demo-ui";

export function FormulaEditorDemo({ session, isDark }: { session: XlsxCellEditorOverlayRenderProps; isDark: boolean }) {
  const [dismissed, setDismissed] = React.useState(false);
  const [closingPreview, setClosingPreview] = React.useState(false);
  React.useEffect(() => { setDismissed(false); setClosingPreview(false); }, [session.signal]);
  React.useEffect(() => {
    session.editorElement.classList.toggle("playground-formula-preview", session.isPreview);
    return () => session.editorElement.classList.remove("playground-formula-preview");
  }, [session.editorElement, session.isPreview]);
  const suggesting = !!session.value.trim() && session.value !== session.context.originalValue;
  const showActions = !dismissed && (suggesting || session.isPreview);
  const reviewing = session.isPreview || closingPreview;

  return (
    <FormulaDemoPopover open={showActions} isDark={isDark} onEscape={session.cancel} label="Cell editor actions">
      <div className="flex flex-col">
        {reviewing ? (
          <>
            <Button variant="ghost" className={formulaActionClassName} onClick={() => session.commit()}>
              <Check className="size-4" />Keep
            </Button>
            <Button variant="ghost" className={formulaActionClassName} onClick={() => { setClosingPreview(true); setDismissed(true); session.discardPreview(); }}>
              <X className="size-4" />Discard
            </Button>
          </>
        ) : (
          <>
            <Button variant="ghost" className={formulaActionClassName} onClick={() => { session.preview(getDemoFormula(session.context.cell, session.context.sheetName)); session.focus(); }}>
              <Sparkles className="size-4" /><span>Write formula <span className="text-muted-foreground">(Demo formula)</span></span>
            </Button>
            <Button variant="ghost" className={formulaActionClassName} onClick={() => { setDismissed(true); session.focus(); }}>
              <CornerDownLeft className="size-4" />Keep typing
            </Button>
          </>
        )}
      </div>
    </FormulaDemoPopover>
  );
}
