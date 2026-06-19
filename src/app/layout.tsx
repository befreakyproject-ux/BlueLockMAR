import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Project: bluelockMA",
  description: "TMSW Gamified Data Dashboard",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body style={{ margin: 0, padding: 0, backgroundColor: "#050811" }}>
        {children}
      </body>
    </html>
  );
}
