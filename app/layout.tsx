import type { Metadata } from "next";
import "./styles.css";

export const metadata: Metadata = {
  title: "Historicals Solver",
  description: "Fill Excel model historical financials from SEC EDGAR filings."
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
