export function coerceUserEnteredValue(value: string): unknown {
  if (value.startsWith("'")) return value.slice(1);
  const trimmed = value.trim();
  if (!trimmed) return value;

  if (/^(true|false)$/i.test(trimmed)) {
    return trimmed.toLowerCase() === "true";
  }

  const number = trimmed.endsWith("%") ? trimmed.slice(0, -1).trim() : trimmed;
  if (/^[+-]?(?:(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(number)) {
    const parsed = Number(number.replaceAll(",", "")) / (trimmed.endsWith("%") ? 100 : 1);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }

  return value;
}
