import './globals.css';
import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'PPT AI Generator',
  description: 'Bikin presentasi PowerPoint otomatis pakai AI',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="id">
      <body className="antialiased">{children}</body>
    </html>
  );
}
