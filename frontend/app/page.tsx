"use client";

import { useState, useEffect } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
    FileSpreadsheet,
    Building2,
    TrendingUp,
    Zap,
    Download,
    Search,
    ChevronRight,
    AlertCircle,
    CheckCircle,
    Loader2,
    BarChart3,
    DollarSign,
    LineChart,
    Sparkles,
    Database,
    Upload,
    Filter,
    ChevronDown,
    X,
    PieChart,
    Table,
    Sun,
    Moon,
    History,
    FileText,
    Presentation,
} from "lucide-react";

// Types
interface Stock {
    symbol: string;
    name: string;
    sector: string;
    sector_code?: string;
}

interface CompanyInfo {
    symbol: string;
    name: string;
    sector: string;
    industry: string;
    market_cap: number | null;
    current_price: number | null;
}

interface JobStatus {
    job_id: string;
    status: string;
    progress: number;
    message: string;
    company_name?: string;
    industry?: string;
    download_url?: string;
    filename?: string;
    validation?: {
        is_valid: boolean;
        errors: any[];
    };
}

interface JobHistoryItem {
    id: string;
    company_name: string;
    industry?: string;
    status: string;
    created_at: string;
    file_path?: string;
}

interface ExportFormat {
    id: string;
    name: string;
    extension: string | null;
    available: boolean;
    description: string;
}


// API Functions
// API Configuration
// Production: Uses relative path (hits Vercel backend via rewrites)
// Development: Hits local Python server
const API_BASE = process.env.NODE_ENV === 'production'
    ? ""
    : "http://localhost:8000";

async function fetchAllStocks(sector?: string): Promise<{ stocks: Stock[]; count: number }> {
    const url = sector ? `${API_BASE}/api/stocks?sector=${sector}` : `${API_BASE}/api/stocks`;
    const response = await fetch(url);
    if (!response.ok) throw new Error("Failed to fetch stocks");
    return response.json();
}

async function searchStocks(query: string): Promise<{ results: Stock[] }> {
    const response = await fetch(`${API_BASE}/api/stocks/search/${query}`);
    if (!response.ok) throw new Error("Search failed");
    return response.json();
}

async function fetchSectors(): Promise<{ sectors: string[] }> {
    const response = await fetch(`${API_BASE}/api/sectors`);
    if (!response.ok) throw new Error("Failed to fetch sectors");
    return response.json();
}

async function generateModel(symbol: string, forecastYears: number): Promise<{ job_id: string }> {
    const response = await fetch(`${API_BASE}/api/model/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            symbol,
            exchange: "NSE",
            forecast_years: forecastYears,
            model_types: ["three_statement", "dcf"],
        }),
    });
    if (!response.ok) throw new Error("Failed to start model generation");
    return response.json();
}

async function generateRawModel(data: RawModelData): Promise<{ job_id: string }> {
    const response = await fetch(`${API_BASE}/api/model/generate-raw`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
    });
    if (!response.ok) throw new Error("Failed to start raw model generation");
    return response.json();
}

interface LBOModelData {
    symbol: string;
    exchange: string;
    holding_period: number;
    entry_multiple: number;
    exit_multiple: number;
    senior_debt_multiple: number;
    senior_interest_rate: number;
    mezz_debt_multiple: number;
    mezz_interest_rate: number;
    sub_debt_multiple: number;
    sub_interest_rate: number;
    revenue_growth: number;
    ebitda_margin: number;
}

async function generateLBOModel(data: LBOModelData): Promise<{ job_id: string }> {
    const response = await fetch(`${API_BASE}/api/model/generate-lbo`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
    });
    if (!response.ok) throw new Error("Failed to start LBO model generation");
    return response.json();
}

interface MAModelData {
    acquirer_symbol: string;
    target_symbol: string;
    exchange: string;
    offer_premium: number;
    percent_stock: number;
    percent_cash: number;
    synergies_revenue: number;
    synergies_cost: number;
    acquirer_growth_rate: number;
    target_growth_rate: number;
}

async function generateMAModel(data: MAModelData): Promise<{ job_id: string }> {
    const response = await fetch(`${API_BASE}/api/model/generate-ma`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
    });
    if (!response.ok) throw new Error("Failed to start M&A model generation");
    return response.json();
}

interface LBOTemplate {
    id: string;
    name: string;
    description: string;
    icon: string;
    assumptions: {
        holding_period: number;
        entry_multiple: number;
        exit_multiple: number;
        senior_debt_multiple: number;
        senior_interest_rate: number;
        mezz_debt_multiple: number;
        mezz_interest_rate: number;
        sub_debt_multiple: number;
        sub_interest_rate: number;
        revenue_growth: number;
        ebitda_margin: number;
    };
    key_metrics: string[];
}

async function fetchLBOTemplates(): Promise<{ templates: LBOTemplate[] }> {
    const response = await fetch(`${API_BASE}/api/templates/lbo`);
    if (!response.ok) return { templates: [] };
    return response.json();
}

async function downloadExportFile(jobId: string, format: "pdf" | "pptx"): Promise<void> {
    const response = await fetch(`${API_BASE}/api/export/${jobId}/${format}`);
    if (!response.ok) throw new Error(`Export failed: ${format}`);

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `model.${format}`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
}

async function uploadExcelFile(
    file: File,
    companyName: string,
    industry: string,
    forecastYears: number
): Promise<{ job_id: string }> {
    const formData = new FormData();
    formData.append("file", file);
    formData.append("company_name", companyName);
    formData.append("industry", industry);
    formData.append("forecast_years", forecastYears.toString());

    const response = await fetch(`${API_BASE}/api/model/upload-excel`, {
        method: "POST",
        body: formData,
    });
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.detail || `Upload failed: ${response.status}`);
    }
    return response.json();
}

async function getJobStatus(jobId: string): Promise<JobStatus> {
    const response = await fetch(`${API_BASE}/api/job/${jobId}`);
    if (!response.ok) throw new Error("Failed to get job status");
    return response.json();
}

async function getJobHistory(limit: number = 20): Promise<{ jobs: JobHistoryItem[] }> {
    const response = await fetch(`${API_BASE}/api/jobs/history?limit=${limit}`);
    if (!response.ok) return { jobs: [] };
    return response.json();
}

async function getExportFormats(): Promise<{ formats: ExportFormat[] }> {
    const response = await fetch(`${API_BASE}/api/export/formats`);
    if (!response.ok) return { formats: [] };
    return response.json();
}

async function downloadExport(jobId: string, format: string): Promise<void> {
    const response = await fetch(`${API_BASE}/api/export/${jobId}/${format}`);
    if (!response.ok) throw new Error(`Export failed: ${format}`);

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `model.${format}`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
}

interface RawModelData {
    company_name: string;
    industry: string;
    forecast_years: number;
    historical_data: {
        income_statement?: { revenue?: number; ebitda?: number; net_income?: number };
        balance_sheet?: { total_assets?: number; total_liabilities?: number };
    };
    assumptions: {
        revenue_growth?: number;
        ebitda_margin?: number;
        tax_rate?: number;
    };
}

// Feature cards data
const FEATURES = [
    {
        icon: Zap,
        title: "AI-Powered",
        description: "Gemini AI classifies industry & builds model logic",
        color: "from-yellow-500 to-orange-500",
    },
    {
        icon: FileSpreadsheet,
        title: "Real Excel Formulas",
        description: "Linked cells, charts, sensitivity analysis",
        color: "from-green-500 to-emerald-500",
    },
    {
        icon: Building2,
        title: "150+ Stocks",
        description: "Power, Banking, FMCG, IT, Pharma & more",
        color: "from-blue-500 to-cyan-500",
    },
    {
        icon: TrendingUp,
        title: "Advanced DCF",
        description: "WACC, scenarios, dashboard with charts",
        color: "from-purple-500 to-pink-500",
    },
];

const SECTOR_COLORS: { [key: string]: string } = {
    power: "bg-yellow-500/20 text-yellow-400 border-yellow-500/30",
    banking: "bg-blue-500/20 text-blue-400 border-blue-500/30",
    it: "bg-purple-500/20 text-purple-400 border-purple-500/30",
    pharma: "bg-green-500/20 text-green-400 border-green-500/30",
    fmcg: "bg-orange-500/20 text-orange-400 border-orange-500/30",
    auto: "bg-red-500/20 text-red-400 border-red-500/30",
    metals: "bg-slate-500/20 text-slate-400 border-slate-500/30",
    oil_gas: "bg-amber-500/20 text-amber-400 border-amber-500/30",
    cement: "bg-stone-500/20 text-stone-400 border-stone-500/30",
    infra: "bg-cyan-500/20 text-cyan-400 border-cyan-500/30",
    nbfc: "bg-indigo-500/20 text-indigo-400 border-indigo-500/30",
    telecom: "bg-pink-500/20 text-pink-400 border-pink-500/30",
    chemicals: "bg-lime-500/20 text-lime-400 border-lime-500/30",
    consumer: "bg-rose-500/20 text-rose-400 border-rose-500/30",
};

type TabType = "stocks" | "excel" | "lbo" | "ma" | "compare";

export default function Home() {
    const [activeTab, setActiveTab] = useState<TabType>("stocks");
    const [stocks, setStocks] = useState<Stock[]>([]);
    const [sectors, setSectors] = useState<string[]>([]);
    const [selectedSector, setSelectedSector] = useState<string>("");
    const [searchQuery, setSearchQuery] = useState("");
    const [searchResults, setSearchResults] = useState<Stock[]>([]);

    const [selectedStock, setSelectedStock] = useState<Stock | null>(null);
    const [forecastYears, setForecastYears] = useState(5);
    const [jobStatus, setJobStatus] = useState<JobStatus | null>(null);
    const [error, setError] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState(false);

    // Excel file upload state
    const [excelFile, setExcelFile] = useState<File | null>(null);
    const [excelCompanyName, setExcelCompanyName] = useState("");
    const [excelIndustry, setExcelIndustry] = useState("general");
    const [isDragging, setIsDragging] = useState(false);

    // LBO model state
    const [lboAssumptions, setLboAssumptions] = useState({
        holding_period: 5,
        entry_multiple: 8.0,
        exit_multiple: 8.0,
        senior_debt_multiple: 3.0,
        senior_interest_rate: 0.08,
        mezz_debt_multiple: 1.5,
        mezz_interest_rate: 0.12,
        sub_debt_multiple: 0.5,
        sub_interest_rate: 0.14,
        revenue_growth: 0.08,
        ebitda_margin: 0.25,
    });

    // M&A model state
    const [acquirerStock, setAcquirerStock] = useState<Stock | null>(null);
    const [targetStock, setTargetStock] = useState<Stock | null>(null);
    const [maSearchQuery, setMaSearchQuery] = useState("");
    const [maSearchResults, setMaSearchResults] = useState<Stock[]>([]);
    const [maSearchType, setMaSearchType] = useState<"acquirer" | "target">("acquirer");
    const [maAssumptions, setMaAssumptions] = useState({
        offer_premium: 0.25,
        percent_stock: 0.50,
        percent_cash: 0.50,
        synergies_revenue: 0,
        synergies_cost: 0,
        acquirer_growth_rate: 0.05,
        target_growth_rate: 0.05,
    });

    // Validation warnings for LBO/M&A inputs
    const validationWarnings = {
        lbo: {
            ebitda_margin: lboAssumptions.ebitda_margin > 0.40
                ? "⚠️ EBITDA margin above 40% is unusually high for most industries"
                : lboAssumptions.ebitda_margin < 0.10
                    ? "⚠️ EBITDA margin below 10% may indicate challenging economics"
                    : null,
            entry_multiple: lboAssumptions.entry_multiple > 15
                ? "⚠️ Entry multiple above 15x is very high for LBO transactions"
                : lboAssumptions.entry_multiple < 5
                    ? "💡 Entry multiple below 5x indicates value opportunity"
                    : null,
            exit_multiple: lboAssumptions.exit_multiple > lboAssumptions.entry_multiple * 1.5
                ? "⚠️ Exit multiple significantly higher than entry - aggressive assumption"
                : null,
            total_debt: (lboAssumptions.senior_debt_multiple + lboAssumptions.mezz_debt_multiple + lboAssumptions.sub_debt_multiple) > 6
                ? "⚠️ Total debt > 6x EBITDA is very aggressive leverage"
                : null,
            revenue_growth: lboAssumptions.revenue_growth > 0.20
                ? "⚠️ Revenue growth above 20% annually is aggressive"
                : lboAssumptions.revenue_growth < 0
                    ? "⚠️ Negative revenue growth - ensure this is intentional"
                    : null,
        },
        ma: {
            offer_premium: maAssumptions.offer_premium > 0.50
                ? "⚠️ Offer premium above 50% is very high"
                : maAssumptions.offer_premium < 0.10
                    ? "💡 Low premium may face target resistance"
                    : null,
            consideration_mix: maAssumptions.percent_stock + maAssumptions.percent_cash !== 1.0
                ? "⚠️ Stock + Cash should equal 100%"
                : null,
            synergies: (maAssumptions.synergies_revenue + maAssumptions.synergies_cost) > 0 &&
                (maAssumptions.synergies_revenue + maAssumptions.synergies_cost) < 100
                ? "💡 Consider if synergy values are in correct units (₹ Cr)"
                : null,
        },
    };

    // Count active warnings
    const lboWarningCount = Object.values(validationWarnings.lbo).filter(w => w).length;
    const maWarningCount = Object.values(validationWarnings.ma).filter(w => w).length;

    // Theme and enhanced features state
    const [theme, setTheme] = useState<"dark" | "light">("dark");
    const [showHistory, setShowHistory] = useState(false);
    const [jobHistory, setJobHistory] = useState<JobHistoryItem[]>([]);
    const [exportFormats, setExportFormats] = useState<ExportFormat[]>([]);
    const [lboTemplates, setLboTemplates] = useState<LBOTemplate[]>([]);

    // Company Comparison state
    const [compareStocks, setCompareStocks] = useState<Stock[]>([]);
    const [compareSearch, setCompareSearch] = useState("");
    const [compareSearchResults, setCompareSearchResults] = useState<Stock[]>([]);
    const [comparisonData, setComparisonData] = useState<any>(null);
    const [isComparing, setIsComparing] = useState(false);


    // Theme toggle
    useEffect(() => {
        document.documentElement.setAttribute("data-theme", theme);
    }, [theme]);

    // Load stocks, sectors, history, export formats, and templates
    useEffect(() => {
        fetchSectors().then((data) => setSectors(data.sectors)).catch(console.error);
        fetchAllStocks().then((data) => setStocks(data.stocks)).catch(console.error);
        getJobHistory().then((data) => setJobHistory(data.jobs)).catch(console.error);
        getExportFormats().then((data) => setExportFormats(data.formats)).catch(console.error);
        fetchLBOTemplates().then((data) => setLboTemplates(data.templates)).catch(console.error);
    }, []);


    // Filter by sector
    useEffect(() => {
        if (selectedSector) {
            fetchAllStocks(selectedSector).then((data) => setStocks(data.stocks)).catch(console.error);
        } else {
            fetchAllStocks().then((data) => setStocks(data.stocks)).catch(console.error);
        }
    }, [selectedSector]);

    // Search stocks
    useEffect(() => {
        if (searchQuery.length >= 2) {
            searchStocks(searchQuery).then((data) => setSearchResults(data.results)).catch(console.error);
        } else {
            setSearchResults([]);
        }
    }, [searchQuery]);

    // M&A Search stocks
    useEffect(() => {
        if (maSearchQuery.length >= 2) {
            searchStocks(maSearchQuery).then((data) => setMaSearchResults(data.results)).catch(console.error);
        } else {
            setMaSearchResults([]);
        }
    }, [maSearchQuery]);

    // Compare Search stocks
    useEffect(() => {
        if (compareSearch.length >= 2) {
            searchStocks(compareSearch).then((data) => setCompareSearchResults(data.results)).catch(console.error);
        } else {
            setCompareSearchResults([]);
        }
    }, [compareSearch]);


    // Generate model
    const handleGenerate = async () => {
        if (activeTab === "stocks" && !selectedStock) return;
        if (activeTab === "excel" && !excelFile) return;
        if (activeTab === "lbo" && !selectedStock) return;
        if (activeTab === "ma" && (!acquirerStock || !targetStock)) return;

        setIsLoading(true);
        setError(null);
        setJobStatus(null);

        try {
            let job_id: string;

            if (activeTab === "stocks" && selectedStock) {
                const result = await generateModel(selectedStock.symbol, forecastYears);
                job_id = result.job_id;
            } else if (activeTab === "excel" && excelFile) {
                const result = await uploadExcelFile(
                    excelFile,
                    excelCompanyName || excelFile.name.replace(/\.xlsx?$/, ''),
                    excelIndustry,
                    forecastYears
                );
                job_id = result.job_id;
            } else if (activeTab === "lbo" && selectedStock) {
                const result = await generateLBOModel({
                    symbol: selectedStock.symbol,
                    exchange: "NSE",
                    ...lboAssumptions,
                });
                job_id = result.job_id;
            } else if (activeTab === "ma" && acquirerStock && targetStock) {
                const result = await generateMAModel({
                    acquirer_symbol: acquirerStock.symbol,
                    target_symbol: targetStock.symbol,
                    exchange: "NSE",
                    ...maAssumptions,
                });
                job_id = result.job_id;
            } else {
                return;
            }

            setJobStatus({ job_id, status: "pending", progress: 0, message: "Starting..." });
        } catch (err) {
            const errorMessage = err instanceof Error ? err.message : "Failed to start model generation. Please try again.";
            setError(errorMessage);
            setIsLoading(false);
        }
    };

    // Compare companies
    const handleCompare = async () => {
        if (compareStocks.length < 2) {
            setError("Select at least 2 companies to compare");
            return;
        }

        setIsComparing(true);
        setError(null);

        try {
            const response = await fetch(`${API_BASE}/api/compare`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    symbols: compareStocks.map(s => s.symbol),
                    exchange: "NSE"
                }),
            });

            if (!response.ok) throw new Error("Comparison failed");

            const data = await response.json();
            setComparisonData(data);
        } catch (err) {
            setError(err instanceof Error ? err.message : "Comparison failed");
        } finally {
            setIsComparing(false);
        }
    };


    // Poll job status
    useEffect(() => {
        if (!jobStatus || jobStatus.status === "completed" || jobStatus.status === "failed") {
            setIsLoading(false);
            return;
        }

        const interval = setInterval(async () => {
            try {
                const status = await getJobStatus(jobStatus.job_id);
                setJobStatus(status);

                if (status.status === "completed" || status.status === "failed") {
                    setIsLoading(false);
                    clearInterval(interval);
                }
            } catch (err) {
                console.error("Failed to get job status:", err);
            }
        }, 1000);

        return () => clearInterval(interval);
    }, [jobStatus?.job_id, jobStatus?.status]);

    const displayedStocks = searchQuery.length >= 2 ? searchResults : stocks;

    return (
        <main className="min-h-screen py-8 px-4 sm:px-6 lg:px-8">
            <div className="max-w-7xl mx-auto">
                {/* Header */}
                <motion.div
                    initial={{ opacity: 0, y: -20 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="text-center mb-8 relative"
                >
                    {/* Theme and History Controls */}
                    <div className="absolute right-0 top-0 flex items-center gap-2">
                        <button
                            onClick={() => setShowHistory(!showHistory)}
                            className="p-2.5 rounded-xl bg-dark-800/50 border border-dark-600 hover:bg-dark-700 hover:border-dark-500 transition-all"
                            title="View History"
                        >
                            <History className="w-5 h-5 text-dark-300" />
                        </button>
                        <button
                            onClick={() => setTheme(theme === "dark" ? "light" : "dark")}
                            className="p-2.5 rounded-xl bg-dark-800/50 border border-dark-600 hover:bg-dark-700 hover:border-dark-500 transition-all"
                            title={theme === "dark" ? "Switch to Light Mode" : "Switch to Dark Mode"}
                        >
                            {theme === "dark" ? (
                                <Sun className="w-5 h-5 text-yellow-400" />
                            ) : (
                                <Moon className="w-5 h-5 text-dark-300" />
                            )}
                        </button>
                    </div>

                    <div className="flex items-center justify-center gap-3 mb-4">
                        <div className="p-3 rounded-2xl bg-gradient-to-br from-primary-500 to-purple-600 shadow-glow-md">
                            <BarChart3 className="w-8 h-8 text-white" />
                        </div>
                        <h1 className="text-4xl sm:text-5xl font-bold bg-gradient-to-r from-white via-primary-200 to-purple-200 bg-clip-text text-transparent">
                            AI Financial Modeler
                        </h1>
                    </div>
                    <p className="text-dark-300 text-lg max-w-2xl mx-auto">
                        Generate institutional-grade Excel models with{" "}
                        <span className="text-primary-400">DCF, charts, sensitivity analysis</span>
                        {" "}for 150+ Indian stocks
                    </p>
                </motion.div>

                {/* History Sidebar */}
                <AnimatePresence>
                    {showHistory && (
                        <motion.div
                            initial={{ opacity: 0, x: 300 }}
                            animate={{ opacity: 1, x: 0 }}
                            exit={{ opacity: 0, x: 300 }}
                            className="fixed right-0 top-0 h-full w-80 bg-dark-900/95 backdrop-blur-lg border-l border-dark-700 z-50 p-6 overflow-y-auto"
                        >
                            <div className="flex items-center justify-between mb-6">
                                <h3 className="text-lg font-semibold text-white">Model History</h3>
                                <button
                                    onClick={() => setShowHistory(false)}
                                    className="p-1.5 rounded-lg hover:bg-dark-800 transition-colors"
                                >
                                    <X className="w-5 h-5 text-dark-400" />
                                </button>
                            </div>

                            {jobHistory.length === 0 ? (
                                <p className="text-dark-400 text-sm">No models generated yet</p>
                            ) : (
                                <div className="space-y-3">
                                    {jobHistory.map((job) => (
                                        <div
                                            key={job.id}
                                            className="p-3 rounded-xl bg-dark-800/50 border border-dark-700 hover:border-dark-600 transition-colors"
                                        >
                                            <p className="font-medium text-white truncate">{job.company_name}</p>
                                            <p className="text-xs text-dark-400 mt-1">
                                                {job.industry || "General"} • {new Date(job.created_at).toLocaleDateString()}
                                            </p>
                                            <div className="flex items-center gap-2 mt-2">
                                                <span className={`px-2 py-0.5 rounded text-xs ${job.status === "completed"
                                                    ? "bg-green-500/20 text-green-400"
                                                    : "bg-yellow-500/20 text-yellow-400"
                                                    }`}>
                                                    {job.status}
                                                </span>
                                                {job.status === "completed" && job.file_path && (
                                                    <a
                                                        href={`/api/download/${job.id}`}
                                                        className="text-xs text-primary-400 hover:underline"
                                                    >
                                                        Download
                                                    </a>
                                                )}
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            )}
                        </motion.div>
                    )}
                </AnimatePresence>


                {/* Features Grid */}
                <motion.div
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.1 }}
                    className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8"
                >
                    {FEATURES.map((feature, index) => (
                        <motion.div
                            key={feature.title}
                            initial={{ opacity: 0, y: 20 }}
                            animate={{ opacity: 1, y: 0 }}
                            transition={{ delay: 0.1 + index * 0.05 }}
                            className="card p-4 group hover:border-primary-500/30 transition-all duration-300"
                        >
                            <div className={`w-10 h-10 rounded-xl bg-gradient-to-br ${feature.color} p-2 mb-3 group-hover:scale-110 transition-transform`}>
                                <feature.icon className="w-full h-full text-white" />
                            </div>
                            <h3 className="font-semibold text-white mb-1">{feature.title}</h3>
                            <p className="text-sm text-dark-400">{feature.description}</p>
                        </motion.div>
                    ))}
                </motion.div>

                {/* Main Content */}
                <div className="grid lg:grid-cols-3 gap-6">
                    {/* Left Panel - Stock Selection */}
                    <motion.div
                        initial={{ opacity: 0, x: -20 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: 0.2 }}
                        className="lg:col-span-2 card p-6"
                    >
                        {/* Tab Navigation */}
                        <div className="flex gap-2 mb-6 flex-wrap">
                            <button
                                onClick={() => setActiveTab("stocks")}
                                className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium transition-all ${activeTab === "stocks"
                                    ? "bg-primary-500 text-white"
                                    : "bg-dark-700/50 text-dark-300 hover:bg-dark-600"
                                    }`}
                            >
                                <Database className="w-4 h-4" />
                                Select Stock
                            </button>
                            <button
                                onClick={() => setActiveTab("excel")}
                                className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium transition-all ${activeTab === "excel"
                                    ? "bg-primary-500 text-white"
                                    : "bg-dark-700/50 text-dark-300 hover:bg-dark-600"
                                    }`}
                            >
                                <Upload className="w-4 h-4" />
                                Excel Input
                            </button>
                            <button
                                onClick={() => setActiveTab("lbo")}
                                className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium transition-all ${activeTab === "lbo"
                                    ? "bg-gradient-to-r from-purple-500 to-pink-500 text-white"
                                    : "bg-dark-700/50 text-dark-300 hover:bg-dark-600"
                                    }`}
                            >
                                <TrendingUp className="w-4 h-4" />
                                LBO Model
                            </button>
                            <button
                                onClick={() => setActiveTab("ma")}
                                className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium transition-all ${activeTab === "ma"
                                    ? "bg-gradient-to-r from-blue-500 to-green-500 text-white"
                                    : "bg-dark-700/50 text-dark-300 hover:bg-dark-600"
                                    }`}
                            >
                                <Building2 className="w-4 h-4" />
                                M&A Model
                            </button>
                            <button
                                onClick={() => setActiveTab("compare")}
                                className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium transition-all ${activeTab === "compare"
                                    ? "bg-gradient-to-r from-amber-500 to-orange-500 text-white"
                                    : "bg-dark-700/50 text-dark-300 hover:bg-dark-600"
                                    }`}
                            >
                                <BarChart3 className="w-4 h-4" />
                                Compare
                            </button>
                        </div>

                        {activeTab === "stocks" ? (
                            <>
                                {/* Search & Filter */}
                                <div className="flex flex-wrap gap-3 mb-4">
                                    <div className="relative flex-1 min-w-[200px]">
                                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-dark-400" />
                                        <input
                                            type="text"
                                            value={searchQuery}
                                            onChange={(e) => setSearchQuery(e.target.value.toUpperCase())}
                                            placeholder="Search stocks..."
                                            className="input-field w-full pl-12"
                                        />
                                    </div>
                                    <div className="relative">
                                        <select
                                            value={selectedSector}
                                            onChange={(e) => setSelectedSector(e.target.value)}
                                            className="input-field pr-10 appearance-none"
                                        >
                                            <option value="">All Sectors</option>
                                            {sectors.map((sector) => (
                                                <option key={sector} value={sector}>
                                                    {sector.replace("_", " ").toUpperCase()}
                                                </option>
                                            ))}
                                        </select>
                                        <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-5 h-5 text-dark-400 pointer-events-none" />
                                    </div>
                                </div>

                                {/* Stocks Grid */}
                                <div className="max-h-[400px] overflow-y-auto custom-scrollbar">
                                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                        {displayedStocks.map((stock) => (
                                            <button
                                                key={stock.symbol}
                                                onClick={() => {
                                                    setSelectedStock(stock);
                                                    setJobStatus(null);
                                                    setError(null);
                                                }}
                                                className={`p-3 rounded-xl border text-left transition-all ${selectedStock?.symbol === stock.symbol
                                                    ? "border-primary-500 bg-primary-500/10"
                                                    : "border-dark-600 hover:border-dark-500 bg-dark-800/50"
                                                    }`}
                                            >
                                                <div className="flex items-center justify-between">
                                                    <div>
                                                        <p className="font-semibold text-white">{stock.symbol}</p>
                                                        <p className="text-sm text-dark-400 truncate max-w-[180px]">{stock.name}</p>
                                                    </div>
                                                    <span className={`text-xs px-2 py-1 rounded-lg border ${SECTOR_COLORS[stock.sector_code || ""] || "bg-dark-600/50 text-dark-300 border-dark-500"
                                                        }`}>
                                                        {stock.sector_code?.replace("_", " ")}
                                                    </span>
                                                </div>
                                            </button>
                                        ))}
                                    </div>
                                </div>
                            </>
                        ) : activeTab === "excel" ? (
                            /* Excel File Upload */
                            <div className="space-y-4">
                                {/* File Drop Zone */}
                                <div
                                    className={`relative p-8 rounded-xl border-2 border-dashed transition-all cursor-pointer ${isDragging
                                        ? "border-primary-500 bg-primary-500/10"
                                        : excelFile
                                            ? "border-green-500 bg-green-500/10"
                                            : "border-dark-600 hover:border-dark-500 bg-dark-800/50"
                                        }`}
                                    onDragOver={(e) => {
                                        e.preventDefault();
                                        setIsDragging(true);
                                    }}
                                    onDragLeave={(e) => {
                                        e.preventDefault();
                                        setIsDragging(false);
                                    }}
                                    onDrop={(e) => {
                                        e.preventDefault();
                                        setIsDragging(false);
                                        const files = e.dataTransfer.files;
                                        if (files.length > 0 && files[0].name.match(/\.xlsx?$/i)) {
                                            setExcelFile(files[0]);
                                        }
                                    }}
                                    onClick={() => document.getElementById('excel-file-input')?.click()}
                                >
                                    <input
                                        id="excel-file-input"
                                        type="file"
                                        accept=".xlsx,.xls"
                                        className="hidden"
                                        onChange={(e) => {
                                            const files = e.target.files;
                                            if (files && files.length > 0) {
                                                setExcelFile(files[0]);
                                            }
                                        }}
                                    />
                                    <div className="text-center">
                                        {excelFile ? (
                                            <>
                                                <div className="w-16 h-16 mx-auto mb-3 rounded-2xl bg-green-500/20 flex items-center justify-center">
                                                    <FileSpreadsheet className="w-8 h-8 text-green-400" />
                                                </div>
                                                <p className="font-semibold text-white mb-1">{excelFile.name}</p>
                                                <p className="text-sm text-dark-400">
                                                    {(excelFile.size / 1024).toFixed(1)} KB
                                                </p>
                                                <button
                                                    onClick={(e) => {
                                                        e.stopPropagation();
                                                        setExcelFile(null);
                                                    }}
                                                    className="mt-3 text-sm text-red-400 hover:text-red-300 flex items-center gap-1 mx-auto"
                                                >
                                                    <X className="w-4 h-4" />
                                                    Remove file
                                                </button>
                                            </>
                                        ) : (
                                            <>
                                                <div className="w-16 h-16 mx-auto mb-3 rounded-2xl bg-dark-700/50 flex items-center justify-center">
                                                    <Upload className="w-8 h-8 text-dark-400" />
                                                </div>
                                                <p className="font-semibold text-white mb-1">Drop Excel file here</p>
                                                <p className="text-sm text-dark-400">or click to browse</p>
                                                <p className="text-xs text-dark-500 mt-2">Supports .xlsx and .xls files</p>
                                            </>
                                        )}
                                    </div>
                                </div>

                                {/* Company Name & Industry */}
                                <div className="grid md:grid-cols-2 gap-4">
                                    <div>
                                        <label className="block text-sm font-medium text-dark-300 mb-2">
                                            Company Name (optional)
                                        </label>
                                        <input
                                            type="text"
                                            value={excelCompanyName}
                                            onChange={(e) => setExcelCompanyName(e.target.value)}
                                            placeholder="Extracted from file if empty"
                                            className="input-field w-full"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-dark-300 mb-2">
                                            Industry
                                        </label>
                                        <select
                                            value={excelIndustry}
                                            onChange={(e) => setExcelIndustry(e.target.value)}
                                            className="input-field w-full"
                                        >
                                            <option value="general">General Corporate</option>
                                            <option value="power">Power & Utilities</option>
                                            <option value="banking">Banking</option>
                                            <option value="it">IT Services</option>
                                            <option value="pharma">Pharmaceuticals</option>
                                            <option value="fmcg">FMCG</option>
                                            <option value="auto">Automobiles</option>
                                            <option value="metals">Metals & Mining</option>
                                            <option value="oil_gas">Oil & Gas</option>
                                            <option value="cement">Cement</option>
                                            <option value="infra">Infrastructure</option>
                                        </select>
                                    </div>
                                </div>

                                {/* Expected Format Info */}
                                <div className="p-4 rounded-xl bg-dark-800/50 border border-dark-600">
                                    <h4 className="font-medium text-white mb-2 flex items-center gap-2">
                                        <FileSpreadsheet className="w-4 h-4 text-primary-400" />
                                        Supported Excel Formats
                                    </h4>
                                    <ul className="text-sm text-dark-400 space-y-1">
                                        <li>• Rows labeled: Revenue, EBITDA, Net Income, etc.</li>
                                        <li>• Multiple sheets: Income Statement, Balance Sheet</li>
                                        <li>• Screener.in or similar export formats</li>
                                        <li>• Company name auto-detected from header</li>
                                    </ul>
                                </div>
                            </div>
                        ) : activeTab === "lbo" ? (
                            /* LBO Model Configuration */
                            <div className="space-y-4">
                                {/* Stock Selection for LBO */}
                                <div>
                                    <label className="block text-sm font-medium text-dark-300 mb-2">
                                        Select Target Company
                                    </label>
                                    <div className="relative">
                                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-dark-400" />
                                        <input
                                            type="text"
                                            value={searchQuery}
                                            onChange={(e) => setSearchQuery(e.target.value.toUpperCase())}
                                            placeholder="Search stocks..."
                                            className="input-field w-full pl-12"
                                        />
                                    </div>
                                    {searchResults.length > 0 && (
                                        <div className="mt-2 max-h-32 overflow-y-auto border border-dark-600 rounded-lg">
                                            {searchResults.slice(0, 5).map((stock) => (
                                                <button
                                                    key={stock.symbol}
                                                    onClick={() => {
                                                        setSelectedStock(stock);
                                                        setSearchQuery("");
                                                        setSearchResults([]);
                                                    }}
                                                    className="w-full p-2 text-left hover:bg-dark-700 text-sm"
                                                >
                                                    <span className="font-semibold text-white">{stock.symbol}</span>
                                                    <span className="text-dark-400 ml-2">{stock.name}</span>
                                                </button>
                                            ))}
                                        </div>
                                    )}
                                    {selectedStock && (
                                        <div className="mt-2 p-2 rounded-lg bg-primary-500/10 border border-primary-500/30 flex items-center justify-between">
                                            <span className="text-white font-medium">{selectedStock.symbol} - {selectedStock.name}</span>
                                            <button onClick={() => setSelectedStock(null)} className="text-dark-400 hover:text-white">
                                                <X className="w-4 h-4" />
                                            </button>
                                        </div>
                                    )}
                                </div>

                                {/* Transaction Assumptions */}
                                <div className="p-4 rounded-xl bg-gradient-to-r from-purple-500/10 to-pink-500/10 border border-purple-500/20">
                                    <h4 className="font-medium text-white mb-3 flex items-center gap-2">
                                        <DollarSign className="w-4 h-4 text-purple-400" />
                                        Transaction Assumptions
                                    </h4>
                                    <div className="grid md:grid-cols-3 gap-3">
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Hold Period (Yrs)</label>
                                            <input
                                                type="number"
                                                value={lboAssumptions.holding_period}
                                                onChange={(e) => setLboAssumptions({ ...lboAssumptions, holding_period: parseInt(e.target.value) || 5 })}
                                                className="input-field w-full text-sm"
                                                min={3} max={10}
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Entry Multiple (x)</label>
                                            <input
                                                type="number"
                                                step="0.5"
                                                value={lboAssumptions.entry_multiple}
                                                onChange={(e) => setLboAssumptions({ ...lboAssumptions, entry_multiple: parseFloat(e.target.value) || 8 })}
                                                className="input-field w-full text-sm"
                                                min={3} max={20}
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Exit Multiple (x)</label>
                                            <input
                                                type="number"
                                                step="0.5"
                                                value={lboAssumptions.exit_multiple}
                                                onChange={(e) => setLboAssumptions({ ...lboAssumptions, exit_multiple: parseFloat(e.target.value) || 8 })}
                                                className="input-field w-full text-sm"
                                                min={3} max={20}
                                            />
                                        </div>
                                    </div>
                                </div>

                                {/* Debt Structure */}
                                <div className="p-4 rounded-xl bg-dark-800/50 border border-dark-600">
                                    <h4 className="font-medium text-white mb-3 flex items-center gap-2">
                                        <BarChart3 className="w-4 h-4 text-blue-400" />
                                        Debt Structure
                                    </h4>
                                    <div className="space-y-3">
                                        {/* Senior Debt */}
                                        <div className="grid grid-cols-2 gap-2">
                                            <div>
                                                <label className="block text-xs text-blue-400 mb-1">Senior Debt (x EBITDA)</label>
                                                <input
                                                    type="number"
                                                    step="0.5"
                                                    value={lboAssumptions.senior_debt_multiple}
                                                    onChange={(e) => setLboAssumptions({ ...lboAssumptions, senior_debt_multiple: parseFloat(e.target.value) || 3 })}
                                                    className="input-field w-full text-sm"
                                                    min={0} max={6}
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs text-blue-400 mb-1">Senior Rate (%)</label>
                                                <input
                                                    type="number"
                                                    step="0.5"
                                                    value={(lboAssumptions.senior_interest_rate * 100).toFixed(1)}
                                                    onChange={(e) => setLboAssumptions({ ...lboAssumptions, senior_interest_rate: (parseFloat(e.target.value) || 8) / 100 })}
                                                    className="input-field w-full text-sm"
                                                    min={3} max={20}
                                                />
                                            </div>
                                        </div>
                                        {/* Mezz Debt */}
                                        <div className="grid grid-cols-2 gap-2">
                                            <div>
                                                <label className="block text-xs text-purple-400 mb-1">Mezz Debt (x EBITDA)</label>
                                                <input
                                                    type="number"
                                                    step="0.5"
                                                    value={lboAssumptions.mezz_debt_multiple}
                                                    onChange={(e) => setLboAssumptions({ ...lboAssumptions, mezz_debt_multiple: parseFloat(e.target.value) || 1.5 })}
                                                    className="input-field w-full text-sm"
                                                    min={0} max={3}
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs text-purple-400 mb-1">Mezz Rate (%)</label>
                                                <input
                                                    type="number"
                                                    step="0.5"
                                                    value={(lboAssumptions.mezz_interest_rate * 100).toFixed(1)}
                                                    onChange={(e) => setLboAssumptions({ ...lboAssumptions, mezz_interest_rate: (parseFloat(e.target.value) || 12) / 100 })}
                                                    className="input-field w-full text-sm"
                                                    min={5} max={25}
                                                />
                                            </div>
                                        </div>
                                        {/* Sub Debt */}
                                        <div className="grid grid-cols-2 gap-2">
                                            <div>
                                                <label className="block text-xs text-red-400 mb-1">Sub Debt (x EBITDA)</label>
                                                <input
                                                    type="number"
                                                    step="0.5"
                                                    value={lboAssumptions.sub_debt_multiple}
                                                    onChange={(e) => setLboAssumptions({ ...lboAssumptions, sub_debt_multiple: parseFloat(e.target.value) || 0.5 })}
                                                    className="input-field w-full text-sm"
                                                    min={0} max={2}
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs text-red-400 mb-1">Sub Rate (%)</label>
                                                <input
                                                    type="number"
                                                    step="0.5"
                                                    value={(lboAssumptions.sub_interest_rate * 100).toFixed(1)}
                                                    onChange={(e) => setLboAssumptions({ ...lboAssumptions, sub_interest_rate: (parseFloat(e.target.value) || 14) / 100 })}
                                                    className="input-field w-full text-sm"
                                                    min={5} max={25}
                                                />
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                {/* Operating Assumptions */}
                                <div className="p-4 rounded-xl bg-dark-800/50 border border-dark-600">
                                    <h4 className="font-medium text-white mb-3 flex items-center gap-2">
                                        <LineChart className="w-4 h-4 text-green-400" />
                                        Operating Assumptions
                                    </h4>
                                    <div className="grid md:grid-cols-2 gap-3">
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Revenue Growth (%)</label>
                                            <input
                                                type="number"
                                                step="1"
                                                value={(lboAssumptions.revenue_growth * 100).toFixed(0)}
                                                onChange={(e) => setLboAssumptions({ ...lboAssumptions, revenue_growth: (parseFloat(e.target.value) || 8) / 100 })}
                                                className="input-field w-full text-sm"
                                                min={-20} max={50}
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">EBITDA Margin (%)</label>
                                            <input
                                                type="number"
                                                step="1"
                                                value={(lboAssumptions.ebitda_margin * 100).toFixed(0)}
                                                onChange={(e) => setLboAssumptions({ ...lboAssumptions, ebitda_margin: (parseFloat(e.target.value) || 25) / 100 })}
                                                className="input-field w-full text-sm"
                                                min={5} max={60}
                                            />
                                        </div>
                                    </div>
                                </div>

                                {/* LBO Model Info */}
                                <div className="p-4 rounded-xl bg-gradient-to-r from-purple-500/10 to-pink-500/10 border border-purple-500/30">
                                    <h4 className="font-medium text-white mb-2">LBO Model Includes:</h4>
                                    <ul className="text-sm text-dark-300 space-y-1">
                                        <li>• Sources & Uses of Funds</li>
                                        <li>• Debt Schedules (Senior, Mezz, Sub)</li>
                                        <li>• Operating Model & Cash Flow</li>
                                        <li>• Returns Analysis (IRR, MoIC)</li>
                                        <li>• Sensitivity Tables</li>
                                    </ul>
                                </div>

                                {/* LBO Validation Warnings */}
                                {lboWarningCount > 0 && (
                                    <div className="p-4 rounded-xl bg-amber-500/10 border border-amber-500/30">
                                        <h4 className="font-medium text-amber-400 mb-2 flex items-center gap-2">
                                            <AlertCircle className="w-4 h-4" />
                                            Assumption Warnings ({lboWarningCount})
                                        </h4>
                                        <ul className="text-sm text-amber-200 space-y-1">
                                            {validationWarnings.lbo.ebitda_margin && (
                                                <li className="flex items-start gap-2">
                                                    <span className="text-xs">{validationWarnings.lbo.ebitda_margin}</span>
                                                </li>
                                            )}
                                            {validationWarnings.lbo.entry_multiple && (
                                                <li className="flex items-start gap-2">
                                                    <span className="text-xs">{validationWarnings.lbo.entry_multiple}</span>
                                                </li>
                                            )}
                                            {validationWarnings.lbo.exit_multiple && (
                                                <li className="flex items-start gap-2">
                                                    <span className="text-xs">{validationWarnings.lbo.exit_multiple}</span>
                                                </li>
                                            )}
                                            {validationWarnings.lbo.total_debt && (
                                                <li className="flex items-start gap-2">
                                                    <span className="text-xs">{validationWarnings.lbo.total_debt}</span>
                                                </li>
                                            )}
                                            {validationWarnings.lbo.revenue_growth && (
                                                <li className="flex items-start gap-2">
                                                    <span className="text-xs">{validationWarnings.lbo.revenue_growth}</span>
                                                </li>
                                            )}
                                        </ul>
                                    </div>
                                )}
                            </div>
                        ) : activeTab === "ma" ? (
                            /* M&A Model Configuration */
                            <div className="space-y-4">
                                {/* Acquirer Selection */}
                                <div className="p-4 rounded-xl bg-blue-500/10 border border-blue-500/20">
                                    <h4 className="font-medium text-blue-400 mb-3 flex items-center gap-2">
                                        <Building2 className="w-4 h-4" />
                                        Acquirer Company
                                    </h4>
                                    <div className="relative">
                                        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-dark-400" />
                                        <input
                                            type="text"
                                            value={maSearchType === "acquirer" ? maSearchQuery : ""}
                                            onChange={(e) => {
                                                setMaSearchType("acquirer");
                                                setMaSearchQuery(e.target.value.toUpperCase());
                                            }}
                                            onFocus={() => setMaSearchType("acquirer")}
                                            placeholder="Search acquirer..."
                                            className="input-field w-full pl-10 text-sm"
                                        />
                                    </div>
                                    {maSearchType === "acquirer" && maSearchResults.length > 0 && (
                                        <div className="mt-2 max-h-24 overflow-y-auto border border-dark-600 rounded-lg">
                                            {maSearchResults.slice(0, 4).map((stock) => (
                                                <button
                                                    key={stock.symbol}
                                                    onClick={() => {
                                                        setAcquirerStock(stock);
                                                        setMaSearchQuery("");
                                                        setMaSearchResults([]);
                                                    }}
                                                    className="w-full p-2 text-left hover:bg-dark-700 text-xs"
                                                >
                                                    <span className="font-semibold text-white">{stock.symbol}</span>
                                                    <span className="text-dark-400 ml-2">{stock.name}</span>
                                                </button>
                                            ))}
                                        </div>
                                    )}
                                    {acquirerStock && (
                                        <div className="mt-2 p-2 rounded-lg bg-blue-500/20 border border-blue-500/30 flex items-center justify-between">
                                            <span className="text-white text-sm font-medium">{acquirerStock.symbol}</span>
                                            <button onClick={() => setAcquirerStock(null)} className="text-dark-400 hover:text-white">
                                                <X className="w-4 h-4" />
                                            </button>
                                        </div>
                                    )}
                                </div>

                                {/* Target Selection */}
                                <div className="p-4 rounded-xl bg-green-500/10 border border-green-500/20">
                                    <h4 className="font-medium text-green-400 mb-3 flex items-center gap-2">
                                        <TrendingUp className="w-4 h-4" />
                                        Target Company
                                    </h4>
                                    <div className="relative">
                                        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-dark-400" />
                                        <input
                                            type="text"
                                            value={maSearchType === "target" ? maSearchQuery : ""}
                                            onChange={(e) => {
                                                setMaSearchType("target");
                                                setMaSearchQuery(e.target.value.toUpperCase());
                                            }}
                                            onFocus={() => setMaSearchType("target")}
                                            placeholder="Search target..."
                                            className="input-field w-full pl-10 text-sm"
                                        />
                                    </div>
                                    {maSearchType === "target" && maSearchResults.length > 0 && (
                                        <div className="mt-2 max-h-24 overflow-y-auto border border-dark-600 rounded-lg">
                                            {maSearchResults.slice(0, 4).map((stock) => (
                                                <button
                                                    key={stock.symbol}
                                                    onClick={() => {
                                                        setTargetStock(stock);
                                                        setMaSearchQuery("");
                                                        setMaSearchResults([]);
                                                    }}
                                                    className="w-full p-2 text-left hover:bg-dark-700 text-xs"
                                                >
                                                    <span className="font-semibold text-white">{stock.symbol}</span>
                                                    <span className="text-dark-400 ml-2">{stock.name}</span>
                                                </button>
                                            ))}
                                        </div>
                                    )}
                                    {targetStock && (
                                        <div className="mt-2 p-2 rounded-lg bg-green-500/20 border border-green-500/30 flex items-center justify-between">
                                            <span className="text-white text-sm font-medium">{targetStock.symbol}</span>
                                            <button onClick={() => setTargetStock(null)} className="text-dark-400 hover:text-white">
                                                <X className="w-4 h-4" />
                                            </button>
                                        </div>
                                    )}
                                </div>

                                {/* Transaction Terms */}
                                <div className="p-4 rounded-xl bg-dark-800/50 border border-dark-600">
                                    <h4 className="font-medium text-white mb-3 flex items-center gap-2">
                                        <DollarSign className="w-4 h-4 text-purple-400" />
                                        Transaction Terms
                                    </h4>
                                    <div className="grid grid-cols-3 gap-3">
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Offer Premium (%)</label>
                                            <input
                                                type="number"
                                                step="5"
                                                value={(maAssumptions.offer_premium * 100).toFixed(0)}
                                                onChange={(e) => setMaAssumptions({ ...maAssumptions, offer_premium: (parseFloat(e.target.value) || 25) / 100 })}
                                                className="input-field w-full text-sm"
                                                min={0} max={100}
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Stock (%)</label>
                                            <input
                                                type="number"
                                                step="10"
                                                value={(maAssumptions.percent_stock * 100).toFixed(0)}
                                                onChange={(e) => {
                                                    const stock = (parseFloat(e.target.value) || 50) / 100;
                                                    setMaAssumptions({ ...maAssumptions, percent_stock: stock, percent_cash: 1 - stock });
                                                }}
                                                className="input-field w-full text-sm"
                                                min={0} max={100}
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Cash (%)</label>
                                            <input
                                                type="number"
                                                value={(maAssumptions.percent_cash * 100).toFixed(0)}
                                                className="input-field w-full text-sm bg-dark-700"
                                                disabled
                                            />
                                        </div>
                                    </div>
                                </div>

                                {/* Synergies */}
                                <div className="p-4 rounded-xl bg-dark-800/50 border border-dark-600">
                                    <h4 className="font-medium text-white mb-3 flex items-center gap-2">
                                        <Sparkles className="w-4 h-4 text-yellow-400" />
                                        Synergies (₹ Crores)
                                    </h4>
                                    <div className="grid grid-cols-2 gap-3">
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Revenue Synergies</label>
                                            <input
                                                type="number"
                                                value={maAssumptions.synergies_revenue}
                                                onChange={(e) => setMaAssumptions({ ...maAssumptions, synergies_revenue: parseFloat(e.target.value) || 0 })}
                                                className="input-field w-full text-sm"
                                                min={0}
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Cost Synergies</label>
                                            <input
                                                type="number"
                                                value={maAssumptions.synergies_cost}
                                                onChange={(e) => setMaAssumptions({ ...maAssumptions, synergies_cost: parseFloat(e.target.value) || 0 })}
                                                className="input-field w-full text-sm"
                                                min={0}
                                            />
                                        </div>
                                    </div>
                                </div>

                                {/* M&A Model Info */}
                                <div className="p-4 rounded-xl bg-gradient-to-r from-blue-500/10 to-green-500/10 border border-blue-500/30">
                                    <h4 className="font-medium text-white mb-2">M&A Model Includes:</h4>
                                    <ul className="text-sm text-dark-300 space-y-1">
                                        <li>• Accretion / Dilution Analysis</li>
                                        <li>• Pro Forma Combined Financials</li>
                                        <li>• Synergy Phase-in Schedule</li>
                                        <li>• Sources & Uses</li>
                                        <li>• Sensitivity Tables</li>
                                    </ul>
                                </div>
                            </div>
                        ) : activeTab === "compare" ? (
                            /* Company Comparison */
                            <div className="space-y-4">
                                <div className="p-4 rounded-xl bg-gradient-to-r from-amber-500/10 to-orange-500/10 border border-amber-500/20">
                                    <h3 className="font-medium text-white mb-3">Select Companies to Compare</h3>

                                    {/* Search */}
                                    <div className="relative mb-3">
                                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-dark-400" />
                                        <input
                                            type="text"
                                            value={compareSearch}
                                            onChange={(e) => setCompareSearch(e.target.value.toUpperCase())}
                                            placeholder="Search stocks to add..."
                                            className="input-field w-full pl-12"
                                        />
                                    </div>

                                    {/* Search Results */}
                                    {compareSearchResults.length > 0 && (
                                        <div className="max-h-40 overflow-y-auto custom-scrollbar mb-3">
                                            <div className="space-y-1">
                                                {compareSearchResults.map((stock) => (
                                                    <button
                                                        key={stock.symbol}
                                                        onClick={() => {
                                                            if (!compareStocks.find(s => s.symbol === stock.symbol) && compareStocks.length < 5) {
                                                                setCompareStocks([...compareStocks, stock]);
                                                                setCompareSearch("");
                                                                setCompareSearchResults([]);
                                                            }
                                                        }}
                                                        className="w-full p-2 rounded-lg border border-dark-600 hover:border-primary-500 bg-dark-800/50 text-left transition-all"
                                                        disabled={compareStocks.find(s => s.symbol === stock.symbol) !== undefined || compareStocks.length >= 5}
                                                    >
                                                        <p className="font-medium text-white text-sm">{stock.symbol}</p>
                                                        <p className="text-xs text-dark-400">{stock.name}</p>
                                                    </button>
                                                ))}
                                            </div>
                                        </div>
                                    )}

                                    {/* Selected Stocks */}
                                    {compareStocks.length > 0 && (
                                        <div>
                                            <h4 className="text-sm font-medium text-dark-300 mb-2">Selected ({compareStocks.length}/5)</h4>
                                            <div className="flex flex-wrap gap-2">
                                                {compareStocks.map((stock) => (
                                                    <div
                                                        key={stock.symbol}
                                                        className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-primary-500/20 border border-primary-500/30"
                                                    >
                                                        <span className="text-sm text-white font-medium">{stock.symbol}</span>
                                                        <button
                                                            onClick={() => setCompareStocks(compareStocks.filter(s => s.symbol !== stock.symbol))}
                                                            className="text-dark-400 hover:text-white"
                                                        >
                                                            <X className="w-3 h-3" />
                                                        </button>
                                                    </div>
                                                ))}
                                            </div>
                                        </div>
                                    )}

                                    {/* Compare Button */}
                                    <button
                                        onClick={handleCompare}
                                        disabled={compareStocks.length < 2 || isComparing}
                                        className="btn-primary w-full mt-3 disabled:opacity-50 disabled:cursor-not-allowed inline-flex items-center justify-center gap-2"
                                    >
                                        {isComparing ? (
                                            <>
                                                <Loader2 className="w-5 h-5 animate-spin" />
                                                Comparing...
                                            </>
                                        ) : (
                                            <>
                                                <BarChart3 className="w-5 h-5" />
                                                Compare Companies
                                            </>
                                        )}
                                    </button>
                                </div>

                                {/* Comparison Results */}
                                {comparisonData && (
                                    <div className="p-4 rounded-xl bg-dark-800/50 border border-dark-600">
                                        <h3 className="font-medium text-white mb-3">Comparison Results</h3>
                                        <div className="overflow-x-auto">
                                            <table className="w-full text-sm">
                                                <thead>
                                                    <tr className="border-b border-dark-600">
                                                        <th className="text-left py-2 px-3 text-dark-300 font-medium">Metric</th>
                                                        {comparisonData.companies.map((company: any) => (
                                                            <th key={company.symbol} className="text-right py-2 px-3 text-white font-medium">
                                                                {company.symbol}
                                                            </th>
                                                        ))}
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr className="border-b border-dark-700">
                                                        <td className="py-2 px-3 text-dark-400">Market Cap (Cr)</td>
                                                        {comparisonData.companies.map((company: any) => (
                                                            <td key={company.symbol} className="text-right py-2 px-3 text-white">
                                                                {company.market_cap ? company.market_cap.toFixed(0) : "N/A"}
                                                            </td>
                                                        ))}
                                                    </tr>
                                                    <tr className="border-b border-dark-700">
                                                        <td className="py-2 px-3 text-dark-400">P/E Ratio</td>
                                                        {comparisonData.companies.map((company: any) => (
                                                            <td key={company.symbol} className="text-right py-2 px-3 text-white">
                                                                {company.pe_ratio ? company.pe_ratio.toFixed(1) : "N/A"}
                                                            </td>
                                                        ))}
                                                    </tr>
                                                    <tr className="border-b border-dark-700">
                                                        <td className="py-2 px-3 text-dark-400">EBITDA Margin (%)</td>
                                                        {comparisonData.companies.map((company: any) => (
                                                            <td key={company.symbol} className="text-right py-2 px-3 text-white">
                                                                {company.ebitda_margin ? (company.ebitda_margin * 100).toFixed(1) + "%" : "N/A"}
                                                            </td>
                                                        ))}
                                                    </tr>
                                                    <tr className="border-b border-dark-700">
                                                        <td className="py-2 px-3 text-dark-400">ROE (%)</td>
                                                        {comparisonData.companies.map((company: any) => (
                                                            <td key={company.symbol} className="text-right py-2 px-3 text-white">
                                                                {company.roe ? (company.roe * 100).toFixed(1) + "%" : "N/A"}
                                                            </td>
                                                        ))}
                                                    </tr>
                                                    <tr>
                                                        <td className="py-2 px-3 text-dark-400">Debt/Equity</td>
                                                        {comparisonData.companies.map((company: any) => (
                                                            <td key={company.symbol} className="text-right py-2 px-3 text-white">
                                                                {company.debt_to_equity ? company.debt_to_equity.toFixed(2) : "N/A"}
                                                            </td>
                                                        ))}
                                                    </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                )}
                            </div>
                        ) : null}

                        {/* Error Display */}
                        <AnimatePresence>
                            {error && (
                                <motion.div
                                    initial={{ opacity: 0, height: 0 }}
                                    animate={{ opacity: 1, height: "auto" }}
                                    exit={{ opacity: 0, height: 0 }}
                                    className="mt-4"
                                >
                                    <div className="flex items-center gap-3 p-4 rounded-xl bg-red-500/10 border border-red-500/30 text-red-400">
                                        <AlertCircle className="w-5 h-5 flex-shrink-0" />
                                        <p>{error}</p>
                                    </div>
                                </motion.div>
                            )}
                        </AnimatePresence>
                    </motion.div>

                    {/* Right Panel - Configuration & Progress */}
                    <motion.div
                        initial={{ opacity: 0, x: 20 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: 0.3 }}
                        className="card p-6"
                    >
                        {/* Selected Stock / Company Info */}
                        {(selectedStock || (activeTab === "excel" && excelFile) || (activeTab === "ma" && (acquirerStock || targetStock))) && (
                            <div className="mb-6">
                                <div className="p-4 rounded-xl bg-gradient-to-r from-primary-500/10 to-purple-500/10 border border-primary-500/20">
                                    <div className="flex items-center gap-2 mb-1">
                                        <Building2 className="w-5 h-5 text-primary-400" />
                                        <h3 className="font-semibold text-white">
                                            {activeTab === "stocks" || activeTab === "lbo"
                                                ? selectedStock?.name
                                                : activeTab === "ma"
                                                    ? `${acquirerStock?.symbol || "?"} + ${targetStock?.symbol || "?"}`
                                                    : (excelCompanyName || excelFile?.name || "Uploaded File")
                                            }
                                        </h3>
                                    </div>
                                    <p className="text-sm text-dark-400">
                                        {activeTab === "stocks"
                                            ? `${selectedStock?.symbol} • ${selectedStock?.sector}`
                                            : activeTab === "lbo"
                                                ? `LBO Model • ${selectedStock?.symbol}`
                                                : activeTab === "ma"
                                                    ? `M&A Model • ${acquirerStock?.name || "Acquirer"} acquiring ${targetStock?.name || "Target"}`
                                                    : `Excel Upload • ${excelIndustry.replace("_", " ")}`
                                        }
                                    </p>
                                </div>
                            </div>
                        )}

                        {/* Configuration */}
                        {!jobStatus && (
                            <div className="space-y-4">
                                <div>
                                    <label className="block text-sm font-medium text-dark-300 mb-2">
                                        Forecast Period
                                    </label>
                                    <select
                                        value={forecastYears}
                                        onChange={(e) => {
                                            setForecastYears(parseInt(e.target.value));
                                        }}
                                        className="input-field w-full"
                                    >
                                        {[3, 5, 7, 10].map((years) => (
                                            <option key={years} value={years}>
                                                {years} Years
                                            </option>
                                        ))}
                                    </select>
                                </div>

                                <button
                                    onClick={handleGenerate}
                                    disabled={
                                        isLoading ||
                                        (activeTab === "stocks" && !selectedStock) ||
                                        (activeTab === "excel" && !excelFile) ||
                                        (activeTab === "lbo" && !selectedStock) ||
                                        (activeTab === "ma" && (!acquirerStock || !targetStock))
                                    }
                                    className="btn-primary w-full flex items-center justify-center gap-2"
                                >
                                    {isLoading ? (
                                        <Loader2 className="w-5 h-5 animate-spin" />
                                    ) : (
                                        <Sparkles className="w-5 h-5" />
                                    )}
                                    Generate Model
                                    <ChevronRight className="w-5 h-5" />
                                </button>

                                {/* Model Features */}
                                <div className="pt-4 border-t border-dark-600">
                                    <h4 className="text-sm font-medium text-dark-300 mb-3">Model Includes:</h4>
                                    <ul className="space-y-2">
                                        {(activeTab === "lbo" ? [
                                            "Sources & Uses of Funds",
                                            "Debt Schedules (Senior, Mezz, Sub)",
                                            "Operating Model & Cash Flow",
                                            "Returns Analysis (IRR, MoIC)",
                                            "Sensitivity Tables",
                                        ] : activeTab === "ma" ? [
                                            "Accretion / Dilution Analysis",
                                            "Pro Forma Combined Financials",
                                            "Synergy Phase-in Schedule",
                                            "Sources & Uses",
                                            "Sensitivity Tables",
                                        ] : [
                                            "Income Statement (5Y Historical + Forecast)",
                                            "Balance Sheet with Balance Check",
                                            "Cash Flow Statement",
                                            "DCF Valuation with WACC",
                                            "Sensitivity Analysis",
                                            "Scenario Analysis (Bear/Base/Bull)",
                                            "Dashboard with Charts",
                                        ]).map((item) => (
                                            <li key={item} className="flex items-center gap-2 text-sm text-dark-400">
                                                <CheckCircle className="w-4 h-4 text-green-400" />
                                                {item}
                                            </li>
                                        ))}
                                    </ul>
                                </div>
                            </div>
                        )}

                        {/* Progress Display */}
                        <AnimatePresence>
                            {jobStatus && (
                                <motion.div
                                    initial={{ opacity: 0, y: 10 }}
                                    animate={{ opacity: 1, y: 0 }}
                                >
                                    <div className="space-y-4">
                                        <div className="flex items-center justify-between">
                                            <div className="flex items-center gap-3">
                                                {jobStatus.status === "completed" ? (
                                                    <CheckCircle className="w-6 h-6 text-green-400" />
                                                ) : jobStatus.status === "failed" ? (
                                                    <AlertCircle className="w-6 h-6 text-red-400" />
                                                ) : (
                                                    <Loader2 className="w-6 h-6 text-primary-400 animate-spin" />
                                                )}
                                                <div>
                                                    <p className="font-medium text-white">
                                                        {jobStatus.company_name || selectedStock?.name || excelCompanyName || "Custom Model"}
                                                    </p>
                                                    <p className="text-sm text-dark-400">{jobStatus.industry}</p>
                                                </div>
                                            </div>
                                            <span className={`px-3 py-1 rounded-full text-sm font-medium ${jobStatus.status === "completed"
                                                ? "bg-green-500/20 text-green-400"
                                                : jobStatus.status === "failed"
                                                    ? "bg-red-500/20 text-red-400"
                                                    : "bg-primary-500/20 text-primary-400"
                                                }`}>
                                                {jobStatus.status === "completed" ? "Ready" :
                                                    jobStatus.status === "failed" ? "Failed" :
                                                        `${jobStatus.progress}%`}
                                            </span>
                                        </div>

                                        {/* Progress bar */}
                                        <div className="h-2 bg-dark-600 rounded-full overflow-hidden">
                                            <motion.div
                                                initial={{ width: 0 }}
                                                animate={{ width: `${jobStatus.progress}%` }}
                                                className="h-full progress-bar rounded-full"
                                            />
                                        </div>

                                        <p className="text-sm text-dark-400">{jobStatus.message}</p>

                                        {/* Download button */}
                                        {jobStatus.status === "completed" && jobStatus.download_url && (
                                            <motion.div
                                                initial={{ opacity: 0, y: 10 }}
                                                animate={{ opacity: 1, y: 0 }}
                                            >
                                                <a
                                                    href={jobStatus.download_url}
                                                    download
                                                    target="_blank"
                                                    rel="noopener noreferrer"
                                                    className="btn-primary w-full inline-flex items-center justify-center gap-2"
                                                    onClick={(e) => {
                                                        e.preventDefault();
                                                        // Programmatic download to ensure file is downloaded
                                                        fetch(jobStatus.download_url!)
                                                            .then(res => res.blob())
                                                            .then(blob => {
                                                                const url = window.URL.createObjectURL(blob);
                                                                const a = document.createElement('a');
                                                                a.href = url;
                                                                a.download = jobStatus.filename || 'financial_model.xlsx';
                                                                document.body.appendChild(a);
                                                                a.click();
                                                                a.remove();
                                                                window.URL.revokeObjectURL(url);
                                                            });
                                                    }}
                                                >
                                                    <Download className="w-5 h-5" />
                                                    Download Excel Model
                                                </a>

                                                {/* PDF Export Button */}
                                                <button
                                                    onClick={async () => {
                                                        try {
                                                            await downloadExportFile(jobStatus.job_id, "pdf");
                                                        } catch (err) {
                                                            alert("PDF export failed. Please try again.");
                                                        }
                                                    }}
                                                    className="btn-secondary w-full mt-2 inline-flex items-center justify-center gap-2"
                                                >
                                                    <FileText className="w-5 h-5" />
                                                    Export to PDF
                                                </button>

                                                {/* PowerPoint Export Button */}
                                                <button
                                                    onClick={async () => {
                                                        try {
                                                            await downloadExportFile(jobStatus.job_id, "pptx");
                                                        } catch (err) {
                                                            alert("PowerPoint export failed. Please try again.");
                                                        }
                                                    }}
                                                    className="btn-secondary w-full mt-2 inline-flex items-center justify-center gap-2"
                                                >
                                                    <Presentation className="w-5 h-5" />
                                                    Export to PowerPoint
                                                </button>

                                                {jobStatus.validation && (
                                                    <div className="mt-4 p-4 rounded-lg bg-dark-800/50">
                                                        <div className="flex items-center gap-2 mb-2">
                                                            {jobStatus.validation.is_valid ? (
                                                                <CheckCircle className="w-4 h-4 text-green-400" />
                                                            ) : (
                                                                <AlertCircle className="w-4 h-4 text-yellow-400" />
                                                            )}
                                                            <span className="text-sm font-medium text-white">
                                                                Model Validation
                                                            </span>
                                                        </div>
                                                        {jobStatus.validation.errors.length > 0 && (
                                                            <ul className="text-sm text-dark-400 space-y-1">
                                                                {jobStatus.validation.errors.slice(0, 3).map((err, i) => (
                                                                    <li key={i}>• {err.message}</li>
                                                                ))}
                                                            </ul>
                                                        )}
                                                    </div>
                                                )}
                                            </motion.div>
                                        )}

                                        {/* New model button */}
                                        {(jobStatus.status === "completed" || jobStatus.status === "failed") && (
                                            <button
                                                onClick={() => {
                                                    setJobStatus(null);
                                                    setSelectedStock(null);
                                                    setExcelFile(null);
                                                    setExcelCompanyName("");
                                                }}
                                                className="btn-secondary w-full mt-2"
                                            >
                                                Generate Another Model
                                            </button>
                                        )}
                                    </div>
                                </motion.div>
                            )}
                        </AnimatePresence>
                    </motion.div>
                </div>

                {/* Footer */}
                <motion.footer
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    transition={{ delay: 0.5 }}
                    className="text-center text-dark-500 text-sm mt-8"
                >
                    <p className="flex items-center justify-center gap-2">
                        <LineChart className="w-4 h-4" />
                        Built with AI • DCF Valuation • Charts & Sensitivity Analysis • 150+ Stocks
                    </p>
                </motion.footer>
            </div>
        </main >
    );
}

