import type { Metadata } from "next";
import { Inter } from "next/font/google";
import "./globals.css";

const inter = Inter({ subsets: ["latin"] });

export const metadata: Metadata = {
    title: "AI Financial Modeler | Institutional-Grade Excel Models",
    description: "Generate professional financial models powered by AI. DCF, 3-Statement, and industry-specific templates for Indian stocks.",
    keywords: ["financial modeling", "DCF", "Excel", "AI", "stock analysis", "India"],
};

export default function RootLayout({
    children,
}: Readonly<{
    children: React.ReactNode;
}>) {
    return (
        <html lang="en" className="dark">
            <body className={inter.className}>
                <div className="mesh-bg" />
                {children}
            </body>
        </html>
    );
}
