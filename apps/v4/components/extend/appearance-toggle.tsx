import { Button } from "@/components/ui/button"
import { IconPlaceholder } from "@/components/icon-placeholder"

// The light/dark switch in a document toolbar; the page owns the theme.
export function AppearanceToggle({
  isDark,
  onIsDarkChange,
}: {
  isDark: boolean
  onIsDarkChange: (isDark: boolean) => void
}) {
  const label = isDark ? "Use light appearance" : "Use dark appearance"
  return (
    <Button
      type="button"
      variant="ghost"
      size="icon-sm"
      onClick={() => onIsDarkChange(!isDark)}
      aria-label={label}
      title={label}
    >
      {isDark ? (
        <IconPlaceholder
          lucide="Sun"
          tabler="IconSun"
          hugeicons="Sun03Icon"
          phosphor="SunIcon"
          remixicon="RiSunLine"
          className="size-4"
        />
      ) : (
        <IconPlaceholder
          lucide="Moon"
          tabler="IconMoon"
          hugeicons="Moon02Icon"
          phosphor="MoonIcon"
          remixicon="RiMoonLine"
          className="size-4"
        />
      )}
    </Button>
  )
}
