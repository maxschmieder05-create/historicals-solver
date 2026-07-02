export type CurrentNonCurrentSignal = "current" | "non-current" | "current-and-non-current" | "none";

export function financialWordsText(input: string) {
  return input.replace(/([a-z])([A-Z])/g, "$1 $2").toLowerCase();
}

export function compactFinancialText(input: string) {
  return financialWordsText(input).replace(/[^a-z0-9]/g, "");
}

export function currentNonCurrentSignalFromText(input: string): CurrentNonCurrentSignal {
  const text = financialWordsText(input);
  const compact = compactFinancialText(text);
  if (textLooksLikeCurrentDebtPortion(text)) return "current";
  if (textMentionsCurrentAndNonCurrent(text)) return "current-and-non-current";
  if (/\bnon[-\s]?current\b|\blong[-\s]?term\b/.test(text) || /noncurrent|longterm/.test(compact)) return "non-current";
  if (/\bcurrent\b/.test(text) || /current/.test(compact)) return "current";
  return "none";
}

export function textMentionsCurrentAndNonCurrent(input: string) {
  const text = financialWordsText(input);
  return /\bcurrent\s+and\s+non[-\s]?current\b/.test(text) || /currentandnoncurrent/.test(compactFinancialText(text));
}

export function textLooksLikeCurrentDebtPortion(input: string) {
  const text = financialWordsText(input);
  const compact = compactFinancialText(text);
  return (
    /\bcurrent maturit|\bcurrent portion\b.*\blong[-\s]?term debt|\blong[-\s]?term debt\b.*\bcurrent\b|\bdebt due within one year\b/.test(text) ||
    /longtermdebtcurrent|currentmaturitiesoflongtermdebt|currentportionoflongtermdebt/.test(compact)
  );
}
