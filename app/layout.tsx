// app/layout.tsx
// The root layout Next.js wraps every page in. It just loads the global
// stylesheet and sets the page title shown in the browser tab.

import "./globals.css";
import type { ReactNode } from "react";

export const metadata = {
  title: "BDA File Cleaner",
  description: "Upload a raw CSV/Excel export and download the cleaned version.",
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
