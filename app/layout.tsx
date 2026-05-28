import type { Metadata } from "next";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import "./styles.css";

export const metadata: Metadata = {
  title: "Historicals Solver",
  description: "Fill Excel model historical financials from SEC EDGAR filings."
};

const inlineStyles = readFileSync(join(process.cwd(), "app/styles.css"), "utf8");

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>
        <style dangerouslySetInnerHTML={{ __html: inlineStyles }} />
        {children}
      </body>
    </html>
  );
}
