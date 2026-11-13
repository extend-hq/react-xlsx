import * as React from "react";
import { Sparkles } from "lucide-react";
import type { XlsxContextMenuAction, XlsxContextMenuRenderProps, XlsxViewerController } from "@extend-ai/react-xlsx";
import { ContextMenu, ContextMenuCheckboxItem, ContextMenuItem, ContextMenuPopup, ContextMenuSeparator, ContextMenuShortcut, ContextMenuSub, ContextMenuSubTrigger } from "@/components/ui/context-menu";
import { FormulaSelectionDemo } from "./formula-selection-demo";
import type { DemoStylePreview } from "./formula-demo-actions";

export function FormulaContextMenu({ menu, controller, onPreview }: {
  menu: XlsxContextMenuRenderProps;
  controller: XlsxViewerController;
  onPreview: (preview: DemoStylePreview | null) => void;
}) {
  const [host, setHost] = React.useState<HTMLDivElement | null>(null);
  const [assistantOpen, setAssistantOpen] = React.useState(false);
  const renderActions = (actions: XlsxContextMenuAction[]): React.ReactNode => actions.map(action => (
    <React.Fragment key={action.id}>
      {action.separatorBefore && <ContextMenuSeparator />}
      {action.children ? <ContextMenuSub>
        <ContextMenuSubTrigger disabled={action.disabled || menu.busy}>{action.label}</ContextMenuSubTrigger>
        <ContextMenuPopup portalProps={{ container: host }} positionMethod="fixed" side="inline-end" align="start" sideOffset={2} className="min-w-56">
          {renderActions(action.children)}
        </ContextMenuPopup>
      </ContextMenuSub> : typeof action.checked === "boolean" ? <ContextMenuCheckboxItem checked={action.checked} indicatorPlacement="end" disabled={action.disabled || menu.busy} closeOnClick={false} onCheckedChange={() => void menu.runAction(action.id)}>
        {action.label}
      </ContextMenuCheckboxItem> : <ContextMenuItem disabled={action.disabled || menu.busy} variant={action.destructive ? "destructive" : "default"} closeOnClick={false} onClick={() => void menu.runAction(action.id)}>
        {action.label}
        {action.shortcut && <ContextMenuShortcut>{action.shortcut}</ContextMenuShortcut>}
      </ContextMenuItem>}
    </React.Fragment>
  ));
  return (
    <div ref={setHost} className={menu.isDark ? "dark" : undefined} style={{ width: 1, height: 1 }} onKeyDown={event => event.stopPropagation()}>
      <ContextMenu open={host !== null} onOpenChange={open => { if (!open) menu.closeMenu(); }}>
        <ContextMenuPopup anchor={host} positionMethod="fixed" portalProps={{ container: host }} side="bottom" sideOffset={0} align="start" finalFocus={false} className="w-56">
          {!menu.context.readOnly && !menu.context.chart && !menu.context.image && !controller.getPivotTableAtCell(menu.context.cell, menu.context.sheetName) && (
            <>
              <ContextMenuSub open={assistantOpen} onOpenChange={setAssistantOpen}>
                <ContextMenuSubTrigger className="[&_svg]:size-4"><Sparkles className="size-4" />Generate formula</ContextMenuSubTrigger>
                <ContextMenuPopup
                  portalProps={{ container: host }}
                  positionMethod="fixed"
                  side="inline-end"
                  align="start"
                  sideOffset={4}
                  className="w-72 transition-[scale,opacity] duration-150 data-starting-style:scale-98 data-starting-style:opacity-0 data-ending-style:scale-98 data-ending-style:opacity-0 motion-reduce:transition-none"
                >
                  <FormulaSelectionDemo
                    inline
                    isDark={menu.isDark}
                    controller={controller}
                    onPreview={onPreview}
                    selection={{ context: menu.context, signal: menu.signal, close: menu.closeMenu, commit: menu.commit,
                      applyFormula: formula => menu.commit(context => controller.setCellFormula(context.cell, formula)) }}
                  />
                </ContextMenuPopup>
              </ContextMenuSub>
              <ContextMenuSeparator />
            </>
          )}
          {renderActions(menu.actions)}
        </ContextMenuPopup>
      </ContextMenu>
    </div>
  );
}
