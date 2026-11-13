import * as React from "react";
import type { XlsxContextMenuAction, XlsxContextMenuActionId, XlsxViewerProps } from "./types";

export function menuItems(root: Element) {
  return Array.from(root.querySelectorAll<HTMLElement>('[role="menuitem"],[role="menuitemcheckbox"]'))
    .filter(item => item.closest('[role="menu"]') === root && !item.matches(':disabled,[aria-disabled="true"]'));
}

export function handleMenuNavigation(event: React.KeyboardEvent) {
  const root = (event.target as HTMLElement).closest('[role="menu"]');
  if (!root) return;
  const items = menuItems(root);
  const current = items.indexOf(document.activeElement as HTMLElement);
  let next: number | undefined;
  if (event.key === "ArrowDown") next = (current + 1) % items.length;
  if (event.key === "ArrowUp") next = (current - 1 + items.length) % items.length;
  if (event.key === "Home") next = 0;
  if (event.key === "End") next = items.length - 1;
  if (event.key.length === 1 && !event.altKey && !event.metaKey && !event.ctrlKey && event.key !== " ") {
    next = items.findIndex((_, i) => items[(current + 1 + i) % items.length].textContent?.trim().replace(/^✓\s*/, "").toLowerCase().startsWith(event.key.toLowerCase()));
    next = next < 0 ? undefined : (current + 1 + next) % items.length;
  }
  if (next !== undefined) {
    event.preventDefault();
    event.stopPropagation();
    items[next]?.focus();
  }
}

export function ContextMenuItems({ actions, busy, runAction, contextMenuClassName, contextMenuStyle }: {
  actions: XlsxContextMenuAction[];
  busy: boolean;
  runAction: (id: XlsxContextMenuActionId) => Promise<boolean>;
} & Pick<XlsxViewerProps, "contextMenuClassName" | "contextMenuStyle">) {
  const [openId, setOpenId] = React.useState<XlsxContextMenuActionId | null>(null);
  return <>{actions.map(action => <React.Fragment key={action.id}>
    {action.separatorBefore && <div className="react-xlsx-context-menu-separator" role="separator" />}
    <ContextMenuItem action={action} busy={busy} runAction={runAction} open={openId === action.id}
      setOpen={open => setOpenId(current => open ? action.id : current === action.id ? null : current)} contextMenuClassName={contextMenuClassName} contextMenuStyle={contextMenuStyle} />
  </React.Fragment>)}</>;
}

function ContextMenuItem({ action, busy, runAction, open, setOpen, contextMenuClassName, contextMenuStyle }: {
  action: XlsxContextMenuAction;
  busy: boolean;
  runAction: (id: XlsxContextMenuActionId) => Promise<boolean>;
  open: boolean;
  setOpen: (open: boolean) => void;
} & Pick<XlsxViewerProps, "contextMenuClassName" | "contextMenuStyle">) {
  const trigger = React.useRef<HTMLButtonElement>(null);
  const popup = React.useRef<HTMLDivElement>(null);
  const focusOnOpen = React.useRef(false);
  const closeTimer = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const [position, setPosition] = React.useState({ left: 0, top: 0 });
  const submenuId = React.useId();
  React.useEffect(() => () => clearTimeout(closeTimer.current), []);
  React.useLayoutEffect(() => {
    if (!open || !trigger.current || !popup.current) return;
    const anchor = trigger.current.getBoundingClientRect();
    const rect = popup.current.getBoundingClientRect();
    setPosition({ left: Math.max(8, anchor.right + rect.width > innerWidth - 8 ? anchor.left - rect.width : anchor.right), top: Math.max(8, Math.min(anchor.top - 4, innerHeight - rect.height - 8)) });
    if (focusOnOpen.current) { menuItems(popup.current)[0]?.focus(); focusOnOpen.current = false; }
  }, [open]);
  const disabled = busy || action.disabled;
  return <div onPointerEnter={() => { clearTimeout(closeTimer.current); setOpen(!disabled && Boolean(action.children)); }}
    onPointerLeave={() => { closeTimer.current = setTimeout(() => setOpen(false), 180); }}>
    <button ref={trigger} className="react-xlsx-context-menu-item" data-action={action.id} data-destructive={action.destructive || undefined}
      disabled={disabled} type="button" tabIndex={-1} role={action.checked === undefined ? "menuitem" : "menuitemcheckbox"}
      aria-checked={action.checked} aria-haspopup={action.children ? "menu" : undefined} aria-expanded={action.children ? open : undefined}
      aria-controls={open ? submenuId : undefined}
      onFocus={() => { if (!action.children) setOpen(false); }}
      onClick={() => { if (action.children) { focusOnOpen.current = true; setOpen(!open); } else void runAction(action.id); }}
      onKeyDown={event => {
        if (event.key === "ArrowRight" && action.children && !disabled) {
          event.preventDefault(); event.stopPropagation(); focusOnOpen.current = true; setOpen(true);
          if (open && popup.current) menuItems(popup.current)[0]?.focus();
        }
      }}>
      <span>{action.checked !== undefined && <span aria-hidden="true" style={{ display: "inline-block", width: 18 }}>{action.checked ? "✓" : ""}</span>}{action.label}</span>
      {action.children ? <span aria-hidden="true">›</span> : action.shortcut && <span className="react-xlsx-context-menu-shortcut">{action.shortcut}</span>}
    </button>
    {open && action.children && <div ref={popup} id={submenuId} role="menu" aria-label={action.label}
      className={["react-xlsx-context-menu", "react-xlsx-context-submenu", contextMenuClassName].filter(Boolean).join(" ")}
      style={{ ...contextMenuStyle, position: "fixed", ...position, maxHeight: "calc(100vh - 16px)", maxWidth: "calc(100vw - 16px)", overflow: "auto", zIndex: 10001, boxSizing: "border-box" }}
      onKeyDown={event => {
        if (event.key === "ArrowLeft" || event.key === "Escape") {
          event.preventDefault(); event.stopPropagation(); setOpen(false); trigger.current?.focus();
        } else handleMenuNavigation(event);
      }}>
      <ContextMenuItems actions={action.children} busy={busy} runAction={runAction} contextMenuClassName={contextMenuClassName} contextMenuStyle={contextMenuStyle} />
    </div>}
  </div>;
}
