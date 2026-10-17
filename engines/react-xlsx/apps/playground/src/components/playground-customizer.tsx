import * as React from "react";
import { HugeiconsIcon, type IconSvgElement } from "@hugeicons/react";
import {
  ArrowDown01Icon,
  ArrowLeft01Icon,
  ArrowRight01Icon,
  Delete02Icon,
  Download01Icon,
  File01Icon,
  Link01Icon,
  MinusSignIcon,
  Moon01Icon,
  PaintBrush02Icon,
  PlusSignIcon,
  RedoIcon,
  RefreshIcon,
  Settings01Icon,
  SparklesIcon,
  Sun01Icon,
  UndoIcon,
  Upload01Icon,
} from "@hugeicons/core-free-icons";
import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Download,
  FileSpreadsheet,
  Link2,
  Minus,
  Moon,
  Paintbrush,
  Plus,
  Redo2,
  RefreshCcw,
  Settings2,
  Sparkles,
  Sun,
  Trash2,
  Undo2,
  Upload,
  type LucideProps,
} from "lucide-react";
import { useTheme } from "next-themes";
import { Badge } from "./ui/badge";
import { Button } from "./ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "./ui/card";
import { Sheet, SheetClose, SheetContent, SheetDescription, SheetHeader, SheetTitle, SheetTrigger } from "./ui/sheet";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "./ui/tabs";
import { cn } from "../lib/utils";

type Appearance = "dark" | "light" | "system";
type FontPreset = "inter" | "mono" | "system";
type IconFamily = "hugeicons" | "lucide";
type RadiusPreset = "large" | "medium" | "none" | "small";
type ThemePresetName = "ember" | "graphite" | "ocean" | "rose" | "spruce";

type ThemeScale = Record<
  | "accent"
  | "accent-foreground"
  | "background"
  | "border"
  | "card"
  | "card-foreground"
  | "chart-1"
  | "chart-2"
  | "chart-3"
  | "chart-4"
  | "chart-5"
  | "destructive"
  | "foreground"
  | "input"
  | "muted"
  | "muted-foreground"
  | "popover"
  | "popover-foreground"
  | "primary"
  | "primary-foreground"
  | "ring"
  | "secondary"
  | "secondary-foreground"
  | "sidebar"
  | "sidebar-accent"
  | "sidebar-accent-foreground"
  | "sidebar-border"
  | "sidebar-foreground"
  | "sidebar-primary"
  | "sidebar-primary-foreground"
  | "sidebar-ring",
  string
>;

type PlaygroundCustomizerSettings = {
  appearance: Appearance;
  font: FontPreset;
  iconFamily: IconFamily;
  radius: RadiusPreset;
  theme: ThemePresetName;
};

type PlaygroundCustomizerContextValue = {
  resolvedAppearance: "dark" | "light";
  settings: PlaygroundCustomizerSettings;
  updateSettings: (nextSettings: Partial<PlaygroundCustomizerSettings>) => void;
};

type PlaygroundIconName =
  | "chevron-down"
  | "chevron-left"
  | "chevron-right"
  | "close"
  | "download"
  | "link"
  | "minus"
  | "moon"
  | "open"
  | "palette"
  | "plus"
  | "redo"
  | "refresh"
  | "settings"
  | "sparkles"
  | "spreadsheet"
  | "sun"
  | "trash"
  | "undo";

const STORAGE_KEY = "react-xlsx-playground-customizer";

const FONT_STACKS: Record<FontPreset, string> = {
  inter: "'Inter Variable', sans-serif",
  mono: "'SFMono-Regular', 'Monaco', 'Cascadia Mono', 'Roboto Mono', monospace",
  system:
    "ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
};

const RADIUS_VALUES: Record<RadiusPreset, string> = {
  none: "0rem",
  small: "0.45rem",
  medium: "0.625rem",
  large: "0.95rem",
};

const DEFAULT_SETTINGS: PlaygroundCustomizerSettings = {
  appearance: "system",
  font: "inter",
  iconFamily: "lucide",
  radius: "medium",
  theme: "graphite",
};

const THEME_PRESETS: Record<
  ThemePresetName,
  {
    description: string;
    label: string;
    light: ThemeScale;
    dark: ThemeScale;
    swatch: string;
  }
> = {
  graphite: {
    label: "Graphite",
    description: "Neutral grayscale with crisp contrast.",
    swatch: "linear-gradient(135deg, oklch(0.98 0 0) 0%, oklch(0.16 0 0) 100%)",
    light: {
      background: "oklch(0.985 0 0)",
      foreground: "oklch(0.17 0 0)",
      card: "oklch(1 0 0)",
      "card-foreground": "oklch(0.17 0 0)",
      popover: "oklch(1 0 0)",
      "popover-foreground": "oklch(0.17 0 0)",
      primary: "oklch(0.22 0 0)",
      "primary-foreground": "oklch(0.985 0 0)",
      secondary: "oklch(0.965 0 0)",
      "secondary-foreground": "oklch(0.22 0 0)",
      muted: "oklch(0.965 0 0)",
      "muted-foreground": "oklch(0.53 0 0)",
      accent: "oklch(0.96 0 0)",
      "accent-foreground": "oklch(0.22 0 0)",
      destructive: "oklch(0.62 0.22 25)",
      border: "oklch(0.91 0 0)",
      input: "oklch(0.92 0 0)",
      ring: "oklch(0.68 0 0)",
      "chart-1": "oklch(0.75 0.12 250)",
      "chart-2": "oklch(0.67 0.16 220)",
      "chart-3": "oklch(0.59 0.18 200)",
      "chart-4": "oklch(0.52 0.14 260)",
      "chart-5": "oklch(0.46 0.11 280)",
      sidebar: "oklch(0.99 0 0)",
      "sidebar-foreground": "oklch(0.17 0 0)",
      "sidebar-primary": "oklch(0.22 0 0)",
      "sidebar-primary-foreground": "oklch(0.985 0 0)",
      "sidebar-accent": "oklch(0.96 0 0)",
      "sidebar-accent-foreground": "oklch(0.22 0 0)",
      "sidebar-border": "oklch(0.91 0 0)",
      "sidebar-ring": "oklch(0.68 0 0)",
    },
    dark: {
      background: "oklch(0.15 0 0)",
      foreground: "oklch(0.985 0 0)",
      card: "oklch(0.2 0 0)",
      "card-foreground": "oklch(0.985 0 0)",
      popover: "oklch(0.2 0 0)",
      "popover-foreground": "oklch(0.985 0 0)",
      primary: "oklch(0.9 0 0)",
      "primary-foreground": "oklch(0.18 0 0)",
      secondary: "oklch(0.28 0 0)",
      "secondary-foreground": "oklch(0.985 0 0)",
      muted: "oklch(0.28 0 0)",
      "muted-foreground": "oklch(0.72 0 0)",
      accent: "oklch(0.32 0 0)",
      "accent-foreground": "oklch(0.985 0 0)",
      destructive: "oklch(0.68 0.18 23)",
      border: "oklch(1 0 0 / 12%)",
      input: "oklch(1 0 0 / 14%)",
      ring: "oklch(0.58 0 0)",
      "chart-1": "oklch(0.76 0.12 250)",
      "chart-2": "oklch(0.69 0.16 220)",
      "chart-3": "oklch(0.61 0.18 200)",
      "chart-4": "oklch(0.55 0.14 260)",
      "chart-5": "oklch(0.49 0.11 280)",
      sidebar: "oklch(0.2 0 0)",
      "sidebar-foreground": "oklch(0.985 0 0)",
      "sidebar-primary": "oklch(0.72 0.1 240)",
      "sidebar-primary-foreground": "oklch(0.15 0 0)",
      "sidebar-accent": "oklch(0.28 0 0)",
      "sidebar-accent-foreground": "oklch(0.985 0 0)",
      "sidebar-border": "oklch(1 0 0 / 12%)",
      "sidebar-ring": "oklch(0.58 0 0)",
    },
  },
  ocean: {
    label: "Ocean",
    description: "Deep blue surfaces with marine highlights.",
    swatch: "linear-gradient(135deg, oklch(0.95 0.03 230) 0%, oklch(0.44 0.16 240) 100%)",
    light: {
      background: "oklch(0.97 0.01 230)",
      foreground: "oklch(0.23 0.03 247)",
      card: "oklch(0.995 0.01 230)",
      "card-foreground": "oklch(0.23 0.03 247)",
      popover: "oklch(0.995 0.01 230)",
      "popover-foreground": "oklch(0.23 0.03 247)",
      primary: "oklch(0.52 0.16 245)",
      "primary-foreground": "oklch(0.985 0 0)",
      secondary: "oklch(0.93 0.02 230)",
      "secondary-foreground": "oklch(0.24 0.04 245)",
      muted: "oklch(0.94 0.02 230)",
      "muted-foreground": "oklch(0.51 0.03 240)",
      accent: "oklch(0.92 0.03 215)",
      "accent-foreground": "oklch(0.24 0.04 245)",
      destructive: "oklch(0.65 0.22 25)",
      border: "oklch(0.88 0.02 230)",
      input: "oklch(0.9 0.02 230)",
      ring: "oklch(0.62 0.13 245)",
      "chart-1": "oklch(0.74 0.16 235)",
      "chart-2": "oklch(0.69 0.17 210)",
      "chart-3": "oklch(0.62 0.13 200)",
      "chart-4": "oklch(0.55 0.1 255)",
      "chart-5": "oklch(0.48 0.09 275)",
      sidebar: "oklch(0.98 0.01 230)",
      "sidebar-foreground": "oklch(0.23 0.03 247)",
      "sidebar-primary": "oklch(0.52 0.16 245)",
      "sidebar-primary-foreground": "oklch(0.985 0 0)",
      "sidebar-accent": "oklch(0.92 0.03 215)",
      "sidebar-accent-foreground": "oklch(0.24 0.04 245)",
      "sidebar-border": "oklch(0.88 0.02 230)",
      "sidebar-ring": "oklch(0.62 0.13 245)",
    },
    dark: {
      background: "oklch(0.18 0.02 245)",
      foreground: "oklch(0.97 0.01 220)",
      card: "oklch(0.22 0.02 245)",
      "card-foreground": "oklch(0.97 0.01 220)",
      popover: "oklch(0.22 0.02 245)",
      "popover-foreground": "oklch(0.97 0.01 220)",
      primary: "oklch(0.72 0.15 230)",
      "primary-foreground": "oklch(0.18 0.02 245)",
      secondary: "oklch(0.28 0.03 240)",
      "secondary-foreground": "oklch(0.97 0.01 220)",
      muted: "oklch(0.28 0.03 240)",
      "muted-foreground": "oklch(0.75 0.02 220)",
      accent: "oklch(0.33 0.05 220)",
      "accent-foreground": "oklch(0.97 0.01 220)",
      destructive: "oklch(0.7 0.17 25)",
      border: "oklch(1 0 0 / 12%)",
      input: "oklch(1 0 0 / 13%)",
      ring: "oklch(0.68 0.12 230)",
      "chart-1": "oklch(0.74 0.16 235)",
      "chart-2": "oklch(0.69 0.17 210)",
      "chart-3": "oklch(0.62 0.13 200)",
      "chart-4": "oklch(0.55 0.1 255)",
      "chart-5": "oklch(0.48 0.09 275)",
      sidebar: "oklch(0.22 0.02 245)",
      "sidebar-foreground": "oklch(0.97 0.01 220)",
      "sidebar-primary": "oklch(0.72 0.15 230)",
      "sidebar-primary-foreground": "oklch(0.18 0.02 245)",
      "sidebar-accent": "oklch(0.33 0.05 220)",
      "sidebar-accent-foreground": "oklch(0.97 0.01 220)",
      "sidebar-border": "oklch(1 0 0 / 12%)",
      "sidebar-ring": "oklch(0.68 0.12 230)",
    },
  },
  spruce: {
    label: "Spruce",
    description: "Muted green UI with spreadsheet-adjacent warmth.",
    swatch: "linear-gradient(135deg, oklch(0.96 0.03 160) 0%, oklch(0.42 0.1 160) 100%)",
    light: {
      background: "oklch(0.975 0.01 160)",
      foreground: "oklch(0.25 0.03 160)",
      card: "oklch(0.995 0.01 160)",
      "card-foreground": "oklch(0.25 0.03 160)",
      popover: "oklch(0.995 0.01 160)",
      "popover-foreground": "oklch(0.25 0.03 160)",
      primary: "oklch(0.54 0.11 161)",
      "primary-foreground": "oklch(0.985 0 0)",
      secondary: "oklch(0.94 0.02 160)",
      "secondary-foreground": "oklch(0.24 0.03 160)",
      muted: "oklch(0.94 0.02 160)",
      "muted-foreground": "oklch(0.5 0.03 160)",
      accent: "oklch(0.92 0.03 170)",
      "accent-foreground": "oklch(0.24 0.03 160)",
      destructive: "oklch(0.65 0.21 25)",
      border: "oklch(0.89 0.02 160)",
      input: "oklch(0.91 0.02 160)",
      ring: "oklch(0.63 0.09 161)",
      "chart-1": "oklch(0.72 0.13 161)",
      "chart-2": "oklch(0.67 0.12 145)",
      "chart-3": "oklch(0.61 0.1 175)",
      "chart-4": "oklch(0.56 0.09 130)",
      "chart-5": "oklch(0.48 0.08 190)",
      sidebar: "oklch(0.985 0.01 160)",
      "sidebar-foreground": "oklch(0.25 0.03 160)",
      "sidebar-primary": "oklch(0.54 0.11 161)",
      "sidebar-primary-foreground": "oklch(0.985 0 0)",
      "sidebar-accent": "oklch(0.92 0.03 170)",
      "sidebar-accent-foreground": "oklch(0.24 0.03 160)",
      "sidebar-border": "oklch(0.89 0.02 160)",
      "sidebar-ring": "oklch(0.63 0.09 161)",
    },
    dark: {
      background: "oklch(0.18 0.02 160)",
      foreground: "oklch(0.97 0.01 160)",
      card: "oklch(0.22 0.02 160)",
      "card-foreground": "oklch(0.97 0.01 160)",
      popover: "oklch(0.22 0.02 160)",
      "popover-foreground": "oklch(0.97 0.01 160)",
      primary: "oklch(0.7 0.12 161)",
      "primary-foreground": "oklch(0.18 0.02 160)",
      secondary: "oklch(0.28 0.03 160)",
      "secondary-foreground": "oklch(0.97 0.01 160)",
      muted: "oklch(0.28 0.03 160)",
      "muted-foreground": "oklch(0.74 0.02 160)",
      accent: "oklch(0.33 0.04 170)",
      "accent-foreground": "oklch(0.97 0.01 160)",
      destructive: "oklch(0.7 0.17 25)",
      border: "oklch(1 0 0 / 12%)",
      input: "oklch(1 0 0 / 14%)",
      ring: "oklch(0.67 0.1 161)",
      "chart-1": "oklch(0.72 0.13 161)",
      "chart-2": "oklch(0.67 0.12 145)",
      "chart-3": "oklch(0.61 0.1 175)",
      "chart-4": "oklch(0.56 0.09 130)",
      "chart-5": "oklch(0.48 0.08 190)",
      sidebar: "oklch(0.22 0.02 160)",
      "sidebar-foreground": "oklch(0.97 0.01 160)",
      "sidebar-primary": "oklch(0.7 0.12 161)",
      "sidebar-primary-foreground": "oklch(0.18 0.02 160)",
      "sidebar-accent": "oklch(0.33 0.04 170)",
      "sidebar-accent-foreground": "oklch(0.97 0.01 160)",
      "sidebar-border": "oklch(1 0 0 / 12%)",
      "sidebar-ring": "oklch(0.67 0.1 161)",
    },
  },
  rose: {
    label: "Rose",
    description: "Warm editorial tones with strong highlights.",
    swatch: "linear-gradient(135deg, oklch(0.96 0.03 15) 0%, oklch(0.55 0.18 15) 100%)",
    light: {
      background: "oklch(0.975 0.01 15)",
      foreground: "oklch(0.26 0.03 15)",
      card: "oklch(0.995 0.01 15)",
      "card-foreground": "oklch(0.26 0.03 15)",
      popover: "oklch(0.995 0.01 15)",
      "popover-foreground": "oklch(0.26 0.03 15)",
      primary: "oklch(0.58 0.18 18)",
      "primary-foreground": "oklch(0.985 0 0)",
      secondary: "oklch(0.94 0.02 15)",
      "secondary-foreground": "oklch(0.26 0.03 15)",
      muted: "oklch(0.94 0.02 15)",
      "muted-foreground": "oklch(0.51 0.04 15)",
      accent: "oklch(0.92 0.03 20)",
      "accent-foreground": "oklch(0.26 0.03 15)",
      destructive: "oklch(0.64 0.23 28)",
      border: "oklch(0.89 0.02 15)",
      input: "oklch(0.91 0.02 15)",
      ring: "oklch(0.63 0.15 18)",
      "chart-1": "oklch(0.74 0.16 18)",
      "chart-2": "oklch(0.69 0.15 35)",
      "chart-3": "oklch(0.63 0.12 340)",
      "chart-4": "oklch(0.57 0.11 300)",
      "chart-5": "oklch(0.49 0.1 265)",
      sidebar: "oklch(0.985 0.01 15)",
      "sidebar-foreground": "oklch(0.26 0.03 15)",
      "sidebar-primary": "oklch(0.58 0.18 18)",
      "sidebar-primary-foreground": "oklch(0.985 0 0)",
      "sidebar-accent": "oklch(0.92 0.03 20)",
      "sidebar-accent-foreground": "oklch(0.26 0.03 15)",
      "sidebar-border": "oklch(0.89 0.02 15)",
      "sidebar-ring": "oklch(0.63 0.15 18)",
    },
    dark: {
      background: "oklch(0.18 0.02 15)",
      foreground: "oklch(0.97 0.01 20)",
      card: "oklch(0.22 0.02 15)",
      "card-foreground": "oklch(0.97 0.01 20)",
      popover: "oklch(0.22 0.02 15)",
      "popover-foreground": "oklch(0.97 0.01 20)",
      primary: "oklch(0.73 0.16 18)",
      "primary-foreground": "oklch(0.18 0.02 15)",
      secondary: "oklch(0.29 0.03 15)",
      "secondary-foreground": "oklch(0.97 0.01 20)",
      muted: "oklch(0.29 0.03 15)",
      "muted-foreground": "oklch(0.75 0.02 15)",
      accent: "oklch(0.34 0.05 20)",
      "accent-foreground": "oklch(0.97 0.01 20)",
      destructive: "oklch(0.7 0.18 28)",
      border: "oklch(1 0 0 / 12%)",
      input: "oklch(1 0 0 / 13%)",
      ring: "oklch(0.69 0.15 18)",
      "chart-1": "oklch(0.74 0.16 18)",
      "chart-2": "oklch(0.69 0.15 35)",
      "chart-3": "oklch(0.63 0.12 340)",
      "chart-4": "oklch(0.57 0.11 300)",
      "chart-5": "oklch(0.49 0.1 265)",
      sidebar: "oklch(0.22 0.02 15)",
      "sidebar-foreground": "oklch(0.97 0.01 20)",
      "sidebar-primary": "oklch(0.73 0.16 18)",
      "sidebar-primary-foreground": "oklch(0.18 0.02 15)",
      "sidebar-accent": "oklch(0.34 0.05 20)",
      "sidebar-accent-foreground": "oklch(0.97 0.01 20)",
      "sidebar-border": "oklch(1 0 0 / 12%)",
      "sidebar-ring": "oklch(0.69 0.15 18)",
    },
  },
  ember: {
    label: "Ember",
    description: "Amber-heavy palette with dark bronze contrast.",
    swatch: "linear-gradient(135deg, oklch(0.97 0.02 85) 0%, oklch(0.56 0.13 75) 100%)",
    light: {
      background: "oklch(0.98 0.01 85)",
      foreground: "oklch(0.28 0.03 75)",
      card: "oklch(1 0.01 85)",
      "card-foreground": "oklch(0.28 0.03 75)",
      popover: "oklch(1 0.01 85)",
      "popover-foreground": "oklch(0.28 0.03 75)",
      primary: "oklch(0.61 0.13 75)",
      "primary-foreground": "oklch(0.98 0.01 85)",
      secondary: "oklch(0.95 0.02 85)",
      "secondary-foreground": "oklch(0.28 0.03 75)",
      muted: "oklch(0.95 0.02 85)",
      "muted-foreground": "oklch(0.54 0.03 75)",
      accent: "oklch(0.93 0.03 80)",
      "accent-foreground": "oklch(0.28 0.03 75)",
      destructive: "oklch(0.65 0.21 28)",
      border: "oklch(0.9 0.02 80)",
      input: "oklch(0.92 0.02 80)",
      ring: "oklch(0.67 0.1 75)",
      "chart-1": "oklch(0.73 0.14 80)",
      "chart-2": "oklch(0.68 0.12 60)",
      "chart-3": "oklch(0.62 0.1 45)",
      "chart-4": "oklch(0.56 0.09 30)",
      "chart-5": "oklch(0.5 0.08 20)",
      sidebar: "oklch(1 0.01 85)",
      "sidebar-foreground": "oklch(0.28 0.03 75)",
      "sidebar-primary": "oklch(0.61 0.13 75)",
      "sidebar-primary-foreground": "oklch(0.98 0.01 85)",
      "sidebar-accent": "oklch(0.93 0.03 80)",
      "sidebar-accent-foreground": "oklch(0.28 0.03 75)",
      "sidebar-border": "oklch(0.9 0.02 80)",
      "sidebar-ring": "oklch(0.67 0.1 75)",
    },
    dark: {
      background: "oklch(0.19 0.02 75)",
      foreground: "oklch(0.97 0.01 85)",
      card: "oklch(0.23 0.02 75)",
      "card-foreground": "oklch(0.97 0.01 85)",
      popover: "oklch(0.23 0.02 75)",
      "popover-foreground": "oklch(0.97 0.01 85)",
      primary: "oklch(0.74 0.13 80)",
      "primary-foreground": "oklch(0.19 0.02 75)",
      secondary: "oklch(0.29 0.03 75)",
      "secondary-foreground": "oklch(0.97 0.01 85)",
      muted: "oklch(0.29 0.03 75)",
      "muted-foreground": "oklch(0.76 0.02 80)",
      accent: "oklch(0.34 0.04 80)",
      "accent-foreground": "oklch(0.97 0.01 85)",
      destructive: "oklch(0.7 0.18 28)",
      border: "oklch(1 0 0 / 12%)",
      input: "oklch(1 0 0 / 13%)",
      ring: "oklch(0.7 0.1 80)",
      "chart-1": "oklch(0.73 0.14 80)",
      "chart-2": "oklch(0.68 0.12 60)",
      "chart-3": "oklch(0.62 0.1 45)",
      "chart-4": "oklch(0.56 0.09 30)",
      "chart-5": "oklch(0.5 0.08 20)",
      sidebar: "oklch(0.23 0.02 75)",
      "sidebar-foreground": "oklch(0.97 0.01 85)",
      "sidebar-primary": "oklch(0.74 0.13 80)",
      "sidebar-primary-foreground": "oklch(0.19 0.02 75)",
      "sidebar-accent": "oklch(0.34 0.04 80)",
      "sidebar-accent-foreground": "oklch(0.97 0.01 85)",
      "sidebar-border": "oklch(1 0 0 / 12%)",
      "sidebar-ring": "oklch(0.7 0.1 80)",
    },
  },
};

const PlaygroundCustomizerContext = React.createContext<PlaygroundCustomizerContextValue | null>(null);

function readStoredSettings(): PlaygroundCustomizerSettings {
  if (typeof window === "undefined") {
    return DEFAULT_SETTINGS;
  }

  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return DEFAULT_SETTINGS;
    }

    const parsed = JSON.parse(raw) as Partial<PlaygroundCustomizerSettings>;
    return {
      appearance:
        parsed.appearance === "dark" || parsed.appearance === "light" || parsed.appearance === "system"
          ? parsed.appearance
          : DEFAULT_SETTINGS.appearance,
      font: parsed.font === "inter" || parsed.font === "mono" || parsed.font === "system" ? parsed.font : DEFAULT_SETTINGS.font,
      iconFamily: parsed.iconFamily === "hugeicons" || parsed.iconFamily === "lucide" ? parsed.iconFamily : DEFAULT_SETTINGS.iconFamily,
      radius:
        parsed.radius === "none" || parsed.radius === "small" || parsed.radius === "medium" || parsed.radius === "large"
          ? parsed.radius
          : DEFAULT_SETTINGS.radius,
      theme:
        parsed.theme === "ember" ||
        parsed.theme === "graphite" ||
        parsed.theme === "ocean" ||
        parsed.theme === "rose" ||
        parsed.theme === "spruce"
          ? parsed.theme
          : DEFAULT_SETTINGS.theme,
    };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

function applyThemeScale(scale: ThemeScale) {
  const root = document.documentElement;

  for (const [token, value] of Object.entries(scale)) {
    root.style.setProperty(`--${token}`, value);
  }
}

function applyCustomizerSettings(settings: PlaygroundCustomizerSettings, resolvedAppearance: "dark" | "light") {
  const root = document.documentElement;
  const activeScale = THEME_PRESETS[settings.theme][resolvedAppearance];

  applyThemeScale(activeScale);
  root.style.setProperty("--radius", RADIUS_VALUES[settings.radius]);
  root.style.setProperty("--font-sans", FONT_STACKS[settings.font]);
  root.dataset.playgroundIcon = settings.iconFamily;
}

export function PlaygroundCustomizerProvider({ children }: { children: React.ReactNode }) {
  const { resolvedTheme, setTheme, theme } = useTheme();
  const [settings, setSettings] = React.useState<PlaygroundCustomizerSettings>(() => readStoredSettings());
  const resolvedAppearance = (resolvedTheme ?? theme ?? "light") === "dark" ? "dark" : "light";

  React.useEffect(() => {
    setTheme(settings.appearance);
  }, [setTheme, settings.appearance]);

  React.useEffect(() => {
    applyCustomizerSettings(settings, resolvedAppearance);
  }, [resolvedAppearance, settings]);

  React.useEffect(() => {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  }, [settings]);

  const value = React.useMemo<PlaygroundCustomizerContextValue>(
    () => ({
      resolvedAppearance,
      settings,
      updateSettings: (nextSettings) => {
        setSettings((current) => ({ ...current, ...nextSettings }));
      },
    }),
    [resolvedAppearance, settings]
  );

  return (
    <PlaygroundCustomizerContext.Provider value={value}>
      {children}
    </PlaygroundCustomizerContext.Provider>
  );
}

export function usePlaygroundCustomizer() {
  const context = React.useContext(PlaygroundCustomizerContext);

  if (!context) {
    throw new Error("usePlaygroundCustomizer must be used within PlaygroundCustomizerProvider");
  }

  return context;
}

function IconTile({
  active,
  children,
  className,
  onClick,
}: {
  active: boolean;
  children: React.ReactNode;
  className?: string;
  onClick: () => void;
}) {
  return (
    <button
      className={cn(
        "hover:border-primary/50 hover:bg-accent/70 flex min-h-16 flex-col items-start justify-between rounded-xl border p-3 text-left transition-all",
        active ? "border-primary bg-accent/80 shadow-sm" : "border-border bg-card",
        className
      )}
      onClick={onClick}
      type="button"
    >
      {children}
    </button>
  );
}

function ThemeSwatch({
  active,
  description,
  label,
  swatch,
  onClick,
}: {
  active: boolean;
  description: string;
  label: string;
  swatch: string;
  onClick: () => void;
}) {
  return (
    <button
      className={cn(
        "group text-left",
        active ? "opacity-100" : "opacity-90 hover:opacity-100"
      )}
      onClick={onClick}
      type="button"
    >
      <div
        className={cn(
          "h-18 rounded-xl border p-1 transition-all",
          active ? "border-primary shadow-sm" : "border-border"
        )}
      >
        <div
          className="flex h-full items-end rounded-[calc(var(--radius)*0.9)] p-2"
          style={{ background: swatch }}
        >
          <div className="rounded-full bg-white/70 px-2 py-0.5 text-[10px] font-medium text-black/80 backdrop-blur-sm">
            {label}
          </div>
        </div>
      </div>
      <div className="mt-2 space-y-0.5">
        <div className="text-foreground text-xs font-medium">{label}</div>
        <div className="text-muted-foreground text-[11px] leading-4">{description}</div>
      </div>
    </button>
  );
}

type LucideComponent = React.ComponentType<LucideProps>;

const LUCIDE_ICON_MAP: Record<Exclude<PlaygroundIconName, "close">, LucideComponent> = {
  "chevron-down": ChevronDown,
  "chevron-left": ChevronLeft,
  "chevron-right": ChevronRight,
  download: Download,
  link: Link2,
  minus: Minus,
  moon: Moon,
  open: Upload,
  palette: Paintbrush,
  plus: Plus,
  redo: Redo2,
  refresh: RefreshCcw,
  settings: Settings2,
  sparkles: Sparkles,
  spreadsheet: FileSpreadsheet,
  sun: Sun,
  trash: Trash2,
  undo: Undo2,
};

const HUGEICON_MAP: Record<PlaygroundIconName, IconSvgElement> = {
  "chevron-down": ArrowDown01Icon,
  "chevron-left": ArrowLeft01Icon,
  "chevron-right": ArrowRight01Icon,
  close: Delete02Icon,
  download: Download01Icon,
  link: Link01Icon,
  minus: MinusSignIcon,
  moon: Moon01Icon,
  open: Upload01Icon,
  palette: PaintBrush02Icon,
  plus: PlusSignIcon,
  redo: RedoIcon,
  refresh: RefreshIcon,
  settings: Settings01Icon,
  sparkles: SparklesIcon,
  spreadsheet: File01Icon,
  sun: Sun01Icon,
  trash: Delete02Icon,
  undo: UndoIcon,
};

export function PlaygroundIcon({
  className,
  name,
  strokeWidth,
}: {
  className?: string;
  name: PlaygroundIconName;
  strokeWidth?: number;
}) {
  const { settings } = usePlaygroundCustomizer();

  if (settings.iconFamily === "hugeicons") {
    return (
      <HugeiconsIcon
        className={className}
        icon={HUGEICON_MAP[name]}
        strokeWidth={strokeWidth ?? 1.9}
      />
    );
  }

  const LucideIcon = LUCIDE_ICON_MAP[name === "close" ? "trash" : name];
  return <LucideIcon className={className} strokeWidth={strokeWidth ?? 1.9} />;
}

function PreviewToolbar() {
  return (
    <div className="bg-background/90 flex items-center gap-2 rounded-xl border p-2 shadow-sm">
      <Button size="sm">
        <PlaygroundIcon name="open" />
        Open
      </Button>
      <Button size="sm" variant="outline">
        <PlaygroundIcon name="download" />
        Export
      </Button>
      <Button size="icon-sm" variant="ghost">
        <PlaygroundIcon name="undo" />
      </Button>
      <Button size="icon-sm" variant="ghost">
        <PlaygroundIcon name="redo" />
      </Button>
    </div>
  );
}

export function PlaygroundCustomizerPanel() {
  const { resolvedAppearance, settings, updateSettings } = usePlaygroundCustomizer();

  return (
    <Sheet>
      <SheetTrigger
        render={<Button size="sm" variant="outline" />}
      >
        <PlaygroundIcon name="settings" />
        Customize
      </SheetTrigger>
      <SheetContent className="w-full max-w-[440px] overflow-y-auto border-l" showCloseButton={false} side="right">
        <SheetHeader className="border-b pb-4">
          <div className="flex items-start justify-between gap-4">
            <div className="space-y-1">
              <SheetTitle>Playground Customizer</SheetTitle>
              <SheetDescription>
                Adjust the shell theme, icon language, radius, and typography. The workbook viewer stays live underneath.
              </SheetDescription>
            </div>
            <SheetClose
              render={<Button size="icon-sm" variant="ghost" />}
            >
              <PlaygroundIcon name="close" />
              <span className="sr-only">Close customizer</span>
            </SheetClose>
          </div>
        </SheetHeader>

        <div className="space-y-5 p-5">
          <Card className="bg-muted/35">
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <PlaygroundIcon name="sparkles" className="size-4" />
                Live Preview
              </CardTitle>
              <CardDescription>
                The controls below affect the toolbar, dialogs, cards, tabs, and general playground chrome.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-3">
              <PreviewToolbar />
              <div className="grid grid-cols-[1fr_auto] gap-2 rounded-xl border bg-card p-3 shadow-sm">
                <div className="space-y-1">
                  <div className="text-sm font-medium">Financial Model.xlsx</div>
                  <div className="text-muted-foreground text-[11px]">
                    {THEME_PRESETS[settings.theme].label} / {settings.iconFamily} / {settings.font}
                  </div>
                </div>
                <Badge variant="secondary">
                  {resolvedAppearance}
                </Badge>
              </div>
            </CardContent>
          </Card>

          <Tabs className="gap-4" defaultValue="theme">
            <TabsList className="w-full justify-start" variant="line">
              <TabsTrigger value="theme">Theme</TabsTrigger>
              <TabsTrigger value="style">Style</TabsTrigger>
            </TabsList>

            <TabsContent value="theme">
              <div className="space-y-5">
                <section className="space-y-2">
                  <div className="flex items-center justify-between">
                    <div className="text-xs font-medium">Appearance</div>
                    <Badge variant="outline">{settings.appearance}</Badge>
                  </div>
                  <div className="grid grid-cols-3 gap-2">
                    {(["light", "dark", "system"] as const).map((appearance) => (
                      <IconTile
                        active={settings.appearance === appearance}
                        key={appearance}
                        onClick={() => updateSettings({ appearance })}
                      >
                        <div className="text-foreground text-xs font-medium capitalize">{appearance}</div>
                        <div className="text-muted-foreground text-[11px]">
                          {appearance === "system" ? "Follow OS preference" : `Lock the shell to ${appearance}`}
                        </div>
                      </IconTile>
                    ))}
                  </div>
                </section>

                <section className="space-y-2">
                  <div className="text-xs font-medium">Color System</div>
                  <div className="grid gap-3">
                    {(Object.entries(THEME_PRESETS) as [ThemePresetName, (typeof THEME_PRESETS)[ThemePresetName]][]).map(([themeName, themePreset]) => (
                      <ThemeSwatch
                        active={settings.theme === themeName}
                        description={themePreset.description}
                        key={themeName}
                        label={themePreset.label}
                        onClick={() => updateSettings({ theme: themeName })}
                        swatch={themePreset.swatch}
                      />
                    ))}
                  </div>
                </section>
              </div>
            </TabsContent>

            <TabsContent value="style">
              <div className="space-y-5">
                <section className="space-y-2">
                  <div className="flex items-center gap-2 text-xs font-medium">
                    <PlaygroundIcon name="palette" className="size-4" />
                    Radius
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    {([
                      ["none", "Sharp edges"],
                      ["small", "Tight corners"],
                      ["medium", "Default balance"],
                      ["large", "Soft panels"],
                    ] as const).map(([radius, description]) => (
                      <IconTile
                        active={settings.radius === radius}
                        key={radius}
                        onClick={() => updateSettings({ radius })}
                      >
                        <div className="text-foreground text-xs font-medium capitalize">{radius}</div>
                        <div className="text-muted-foreground text-[11px]">{description}</div>
                      </IconTile>
                    ))}
                  </div>
                </section>

                <section className="space-y-2">
                  <div className="flex items-center gap-2 text-xs font-medium">
                    <PlaygroundIcon name="settings" className="size-4" />
                    Icon Family
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    {([
                      ["lucide", "Clean and technical"],
                      ["hugeicons", "More expressive silhouettes"],
                    ] as const).map(([iconFamily, description]) => (
                      <IconTile
                        active={settings.iconFamily === iconFamily}
                        key={iconFamily}
                        onClick={() => updateSettings({ iconFamily })}
                      >
                        <div className="flex items-center gap-2">
                          <PlaygroundIcon name="settings" className="size-4" />
                          <div className="text-foreground text-xs font-medium capitalize">{iconFamily}</div>
                        </div>
                        <div className="text-muted-foreground text-[11px]">{description}</div>
                      </IconTile>
                    ))}
                  </div>
                </section>

                <section className="space-y-2">
                  <div className="text-xs font-medium">Typography</div>
                  <div className="grid grid-cols-3 gap-2">
                    {([
                      ["inter", "Balanced"],
                      ["system", "Native"],
                      ["mono", "Tooling"],
                    ] as const).map(([font, description]) => (
                      <IconTile
                        active={settings.font === font}
                        className={font === "mono" ? "font-mono" : ""}
                        key={font}
                        onClick={() => updateSettings({ font })}
                      >
                        <div className="text-foreground text-xs font-medium capitalize">{font}</div>
                        <div className="text-muted-foreground text-[11px]">{description}</div>
                      </IconTile>
                    ))}
                  </div>
                </section>
              </div>
            </TabsContent>
          </Tabs>

          <div className="flex items-center justify-between gap-3 rounded-xl border bg-muted/35 p-3">
            <div className="space-y-1">
              <div className="text-xs font-medium">Reset customizer</div>
              <div className="text-muted-foreground text-[11px]">
                Restore the default graphite theme, medium radius, Lucide icons, and system appearance.
              </div>
            </div>
            <Button
              onClick={() => updateSettings(DEFAULT_SETTINGS)}
              size="sm"
              variant="outline"
            >
              Reset
            </Button>
          </div>
        </div>
      </SheetContent>
    </Sheet>
  );
}
