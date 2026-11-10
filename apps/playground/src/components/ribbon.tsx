import * as React from "react";
import { ChevronDown } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverPopup, PopoverTrigger } from "@/components/ui/popover";
import { Tooltip, TooltipPopup, TooltipTrigger } from "@/components/ui/tooltip";
import { cn } from "@/lib/utils";

/** A labelled cluster of ribbon commands, separated from its neighbours by a hairline. */
export function RibbonGroup({ children, className }: { children: React.ReactNode; className?: string }) {
  return (
    <div className={cn("flex shrink-0 items-center gap-0.5 px-1.5 first:ps-0 [&+&]:border-s", className)} role="group">
      {children}
    </div>
  );
}

type RibbonButtonProps = Omit<React.ComponentProps<typeof Button>, "size" | "variant"> & {
  /** Accessible name and tooltip text. */
  label: string;
  /** Optional keyboard hint shown in the tooltip. */
  shortcut?: string;
  /** Render the label next to the icon. */
  showLabel?: boolean;
  active?: boolean;
};

/** Flat, Excel-style command button. Icon-only by default, with a tooltip. */
export function RibbonButton({ active, children, className, label, shortcut, showLabel, ...props }: RibbonButtonProps) {
  return (
    <Tooltip>
      <TooltipTrigger
        render={
          <Button
            aria-label={label}
            aria-pressed={active}
            className={cn(
              "h-7 min-w-7 rounded-md px-1.5 font-normal text-[13px] sm:h-7 [&_svg:not([class*='size-'])]:size-4",
              showLabel && "px-2",
              active && "bg-brand/10 text-brand hover:bg-brand/15",
              className,
            )}
            size="sm"
            variant="ghost"
            {...props}
          />
        }
      >
        {children}
        {showLabel ? <span>{label}</span> : null}
      </TooltipTrigger>
      <TooltipPopup side="bottom">
        {label}
        {shortcut ? <span className="ms-2 text-muted-foreground">{shortcut}</span> : null}
      </TooltipPopup>
    </Tooltip>
  );
}

/** Ghost dropdown trigger used for ribbon menus (Merge, Borders, etc). */
export function RibbonMenuButton({
  className,
  icon,
  label,
  showLabel = true,
  ...props
}: React.ComponentProps<typeof Button> & { label: string; icon?: React.ReactNode; showLabel?: boolean }) {
  return (
    <Button
      aria-label={label}
      className={cn(
        "h-7 gap-1 rounded-md px-1.5 font-normal text-[13px] sm:h-7 [&_svg:not([class*='size-'])]:size-4",
        className,
      )}
      size="sm"
      variant="ghost"
      {...props}
    >
      {icon}
      {showLabel ? <span>{label}</span> : null}
      <ChevronDown className="size-3! opacity-60" />
    </Button>
  );
}

// Excel's default "Office" theme palette (top row) with lighter/darker tints.
const THEME_COLORS = [
  ["#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
  ["#F2F2F2", "#7F7F7F", "#D0CECE", "#D6DCE4", "#D9E2F3", "#FBE5D5", "#EDEDED", "#FFF2CC", "#DEEBF6", "#E2EFD9"],
  ["#D8D8D8", "#595959", "#AEABAB", "#ADB9CA", "#B4C6E7", "#F7CBAC", "#DBDBDB", "#FEE599", "#BDD7EE", "#C5E0B3"],
  ["#BFBFBF", "#3F3F3F", "#757070", "#8496B0", "#8EAADB", "#F4B183", "#C9C9C9", "#FFD965", "#9DC3E6", "#A8D08D"],
  ["#A5A5A5", "#262626", "#3A3838", "#323F4F", "#2F5496", "#C55A11", "#7B7B7B", "#BF9000", "#2E75B5", "#538135"],
];
const STANDARD_COLORS = ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"];

/**
 * Excel-style split color button: the left half applies the current color,
 * the chevron opens a theme palette.
 */
export function RibbonColorButton({
  color,
  disabled,
  icon,
  label,
  onApply,
  onColorChange,
}: {
  color: string;
  disabled?: boolean;
  icon: React.ReactNode;
  label: string;
  onApply: (color: string) => void;
  onColorChange: (color: string) => void;
}) {
  const [open, setOpen] = React.useState(false);
  const pick = (next: string) => {
    onColorChange(next);
    onApply(next);
    setOpen(false);
  };

  return (
    <div className="flex items-center">
      <RibbonButton className="rounded-e-none pe-1" disabled={disabled} label={label} onClick={() => onApply(color)}>
        <span className="relative flex flex-col items-center">
          {icon}
          <span className="-mt-px h-[3px] w-4 rounded-[1px]" style={{ backgroundColor: color }} />
        </span>
      </RibbonButton>
      <Popover onOpenChange={setOpen} open={open}>
        <PopoverTrigger
          render={
            <Button
              aria-label={`${label} options`}
              className="h-7 w-4 min-w-4 rounded-s-none rounded-e-md px-0 sm:h-7"
              disabled={disabled}
              size="sm"
              variant="ghost"
            />
          }
        >
          <ChevronDown className="size-3! opacity-60" />
        </PopoverTrigger>
        <PopoverPopup align="start" className="w-auto [&>div]:p-2.5">
          <div className="flex flex-col gap-2">
            <div className="text-[11px] font-medium text-muted-foreground">Theme colors</div>
            <div className="grid grid-cols-10 gap-1">
              {THEME_COLORS.flatMap((row) => row).map((swatch, index) => (
                <Swatch color={swatch} key={`${index}-${swatch}`} onPick={pick} selected={swatch === color.toUpperCase()} />
              ))}
            </div>
            <div className="text-[11px] font-medium text-muted-foreground">Standard colors</div>
            <div className="grid grid-cols-10 gap-1">
              {STANDARD_COLORS.map((swatch) => (
                <Swatch color={swatch} key={swatch} onPick={pick} selected={swatch === color.toUpperCase()} />
              ))}
            </div>
            <label className="mt-1 flex cursor-pointer items-center gap-2 rounded-md px-1 py-1 text-xs hover:bg-accent">
              <input
                className="size-5 cursor-pointer rounded border-0 bg-transparent p-0"
                onChange={(event) => pick(event.target.value.toUpperCase())}
                type="color"
                value={color}
              />
              More colors…
            </label>
          </div>
        </PopoverPopup>
      </Popover>
    </div>
  );
}

function Swatch({ color, onPick, selected }: { color: string; onPick: (color: string) => void; selected: boolean }) {
  return (
    <button
      aria-label={color}
      className={cn(
        "size-4 rounded-[2px] border border-black/10 outline-none transition-transform hover:scale-125 hover:ring-1 hover:ring-foreground/40 focus-visible:ring-2 focus-visible:ring-ring",
        selected && "ring-2 ring-brand ring-offset-1",
      )}
      onClick={() => onPick(color)}
      style={{ backgroundColor: color }}
      type="button"
    />
  );
}

export function RibbonDivider() {
  return <div aria-hidden className="mx-1 h-5 w-px shrink-0 bg-border" />;
}
