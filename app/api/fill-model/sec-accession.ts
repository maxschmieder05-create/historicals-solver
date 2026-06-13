export function normalizeAccession(input: string | number | null | undefined) {
  const value = String(input ?? "").trim();
  if (!value) return "";

  const dashed = value.match(/(\d{1,10})\s*-\s*(\d{2})\s*-\s*(\d{1,6})/);
  if (dashed) return `${dashed[1].padStart(10, "0")}${dashed[2]}${dashed[3].padStart(6, "0")}`;

  const digitRuns = value.replace(/-/g, "").match(/\d+/g) ?? [];
  const accessionLike = digitRuns
    .filter((run) => run.length >= 9)
    .sort((a, b) => b.length - a.length)[0];
  const compact = accessionLike ?? value.replace(/-/g, "").replace(/\s+/g, "");
  if (!/^\d+$/.test(compact)) return compact;
  if (compact.length >= 18) return compact;
  if (compact.length > 8) return `${compact.slice(0, -8).padStart(10, "0")}${compact.slice(-8)}`;
  return compact;
}

export function normalizeAccessionList(input: string | number | null | undefined) {
  return unique(
    String(input ?? "")
      .split(/[;,\n]/)
      .map((item) => normalizeAccession(item))
      .filter(Boolean)
  );
}

export function normalizeCik(input: string | number | null | undefined) {
  const digits = String(input ?? "").trim().replace(/\D/g, "");
  return digits ? digits.padStart(10, "0") : "";
}

export function cikFromAccession(input: string | number | null | undefined) {
  const accession = normalizeAccession(input);
  if (!/^\d{9,}$/.test(accession)) return "";
  return normalizeCik(accession.slice(0, -8));
}

function unique<T>(items: T[]) {
  return Array.from(new Set(items));
}
